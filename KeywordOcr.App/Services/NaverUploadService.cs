using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace KeywordOcr.App.Services;

public sealed record NaverUploadOptions
{
    public int RowStart { get; set; }
    public int RowEnd { get; set; }
    public bool DryRun { get; set; } = true;
}

public sealed record NaverUploadResultItem(
    int Row,
    string Name,
    string Status,
    string ProductId,
    string Error);

public sealed record NaverUploadResult(
    IReadOnlyList<NaverUploadResultItem> Items,
    int SuccessCount,
    int FailCount,
    int TotalCount);

/// <summary>
/// 네이버 스마트스토어 상품 업로드 서비스 (순수 C# — Python 의존성 없음)
/// </summary>
public sealed class NaverUploadService
{
    public async Task<NaverUploadResult> UploadAsync(
        string sourcePath,
        NaverUploadOptions options,
        IProgress<string>? progress = null,
        CancellationToken ct = default)
    {
        void Log(string msg) => progress?.Report(msg);

        using var api = NaverCommerceApiClient.FromKeyFile();

        // 엑셀 읽기
        Log("가공파일 읽는 중...");
        var allRows = ReadSourceFile(sourcePath);
        Log($"{allRows.Count}개 상품 로드 완료");

        // 행 필터
        List<Dictionary<string, object?>> targetRows;
        if (options.RowStart > 0)
        {
            var end = options.RowEnd > 0 ? options.RowEnd : options.RowStart;
            targetRows = allRows.Where(r =>
            {
                var rowNum = (int)r["_row_num"]! - 1;
                return rowNum >= options.RowStart && rowNum <= end;
            }).ToList();
        }
        else
        {
            targetRows = allRows;
        }

        Log($"처리 대상: {targetRows.Count}개");
        var results = new List<NaverUploadResultItem>();

        for (int idx = 0; idx < targetRows.Count; idx++)
        {
            ct.ThrowIfCancellationRequested();
            var row = targetRows[idx];
            var rowNum = (int)row["_row_num"]!;
            var productName = GetStr(row, "상품명")
                .OrIfEmpty(GetStr(row, "최종키워드2차"))
                .OrIfEmpty(GetStr(row, "1차키워드"));
            var shortName = productName.Length > 30 ? productName[..30] : productName;

            Log($"[{idx + 1}/{targetRows.Count}] {shortName}...");

            // 카테고리 결정
            string categoryId;
            string catName;

            var presetCat = GetStr(row, "네이버카테고리코드").OrIfEmpty(GetStr(row, "네이버카테고리"));
            if (!string.IsNullOrEmpty(presetCat))
            {
                categoryId = ((long)double.Parse(presetCat)).ToString();
                catName = $"엑셀지정({categoryId})";
                Log($"  → 엑셀 카테고리 사용: {categoryId}");
            }
            else
            {
                try
                {
                    using var catDoc = await api.PredictCategoryAsync(productName, ct);
                    var root = catDoc.RootElement;

                    if (root.TryGetProperty("_error", out _))
                    {
                        var msg = root.TryGetProperty("_msg", out var mp) ? mp.ToString() : "카테고리 추천 실패";
                        results.Add(new NaverUploadResultItem(rowNum, shortName, "CATEGORY_FAIL", "", msg.Length > 200 ? msg[..200] : msg));
                        continue;
                    }

                    if (root.TryGetProperty("contents", out var contents) && contents.GetArrayLength() > 0)
                    {
                        var top = contents[0];
                        categoryId = top.GetProperty("categoryId").ToString();
                        var wholeName = top.TryGetProperty("wholeCategoryName", out var wn) ? wn.GetString() ?? "" : "";
                        catName = wholeName;
                    }
                    else
                    {
                        results.Add(new NaverUploadResultItem(rowNum, shortName, "CATEGORY_FAIL", "", "유사 상품 없음"));
                        continue;
                    }
                }
                catch (Exception ex)
                {
                    results.Add(new NaverUploadResultItem(rowNum, shortName, "CATEGORY_FAIL", "", ex.Message.Length > 200 ? ex.Message[..200] : ex.Message));
                    continue;
                }
            }

            if (options.DryRun)
            {
                Log($"  → {catName}");
                results.Add(new NaverUploadResultItem(rowNum, shortName, "DRY_RUN_OK", $"{catName} ({categoryId})", ""));
                continue;
            }

            // 실제 등록
            try
            {
                // 이미지: listing_images 가공이미지 우선, 없으면 엑셀 URL fallback
                var exportRoot = GetStr(row, "_export_root");
                var sku = GetStr(row, "자체 상품코드");
                var listingImages = !string.IsNullOrEmpty(exportRoot) && !string.IsNullOrEmpty(sku)
                    ? FindListingImages(exportRoot, sku)
                    : new List<string>();

                var imageUrls = listingImages.Count > 0 ? listingImages : CollectImageUrls(row);
                JsonObject? imagesNode = null;

                if (imageUrls.Count > 0)
                {
                    var uploadedUrls = new List<string>();
                    foreach (var imgUrl in imageUrls.Take(9))
                    {
                        try
                        {
                            var uploaded = await api.UploadImageAsync(imgUrl, ct);
                            uploadedUrls.Add(uploaded);
                        }
                        catch { /* skip failed images */ }
                    }

                    if (uploadedUrls.Count > 0)
                    {
                        imagesNode = new JsonObject
                        {
                            ["representativeImage"] = new JsonObject { ["url"] = uploadedUrls[0] },
                        };
                        if (uploadedUrls.Count > 1)
                        {
                            var optImages = new JsonArray();
                            foreach (var u in uploadedUrls.Skip(1))
                                optImages.Add(new JsonObject { ["url"] = u });
                            imagesNode["optionalImages"] = optImages;
                        }
                    }
                }

                var productJson = BuildNaverProduct(row, categoryId, imagesNode);
                var productElement = JsonSerializer.Deserialize<JsonElement>(productJson.ToJsonString());
                using var resp = await api.CreateProductAsync(productElement, ct);
                var respRoot = resp.RootElement;

                // 등록불가 태그 처리 후 재시도
                if (respRoot.TryGetProperty("_error", out _))
                {
                    var errMsg = respRoot.TryGetProperty("_msg", out var mp) ? mp.ToString() : "";
                    var restrictedTags = ExtractRestrictedTags(errMsg);
                    if (restrictedTags.Count > 0)
                    {
                        // 태그 제거 후 재시도
                        RemoveRestrictedTags(productJson, restrictedTags);
                        var retryElement = JsonSerializer.Deserialize<JsonElement>(productJson.ToJsonString());
                        using var resp2 = await api.CreateProductAsync(retryElement, ct);
                        var resp2Root = resp2.RootElement;
                        if (resp2Root.TryGetProperty("_error", out _))
                        {
                            var msg2 = resp2Root.TryGetProperty("_msg", out var mp2) ? mp2.ToString() : "등록 실패";
                            if (msg2.Length > 200) msg2 = msg2[..200];
                            results.Add(new NaverUploadResultItem(rowNum, shortName, "FAIL", "", msg2));
                        }
                        else
                        {
                            var pid = ExtractProductId(resp2Root);
                            results.Add(new NaverUploadResultItem(rowNum, shortName, "OK", pid, ""));
                        }
                        continue;
                    }

                    if (errMsg.Length > 200) errMsg = errMsg[..200];
                    results.Add(new NaverUploadResultItem(rowNum, shortName, "FAIL", "", errMsg));
                }
                else
                {
                    var pid = ExtractProductId(respRoot);
                    results.Add(new NaverUploadResultItem(rowNum, shortName, "OK", pid, ""));
                }
            }
            catch (Exception ex)
            {
                var msg = ex.Message.Length > 200 ? ex.Message[..200] : ex.Message;
                results.Add(new NaverUploadResultItem(rowNum, shortName, "FAIL", "", msg));
            }

            // 속도 제한
            if ((idx + 1) % 5 == 0)
                await Task.Delay(1000, ct);
        }

        var successCount = results.Count(r => r.Status is "OK" or "DRY_RUN_OK");
        var failCount = results.Count - successCount;
        return new NaverUploadResult(results, successCount, failCount, results.Count);
    }

    // ── 상품 JSON 빌드 ─────────────────────────────

    private static JsonObject BuildNaverProduct(
        Dictionary<string, object?> row, string categoryId, JsonObject? images)
    {
        var productName = GetStr(row, "상품명")
            .OrIfEmpty(GetStr(row, "최종키워드2차"))
            .OrIfEmpty(GetStr(row, "1차키워드"));
        if (productName.Length > 100) productName = productName[..100];

        var salePrice = Math.Max(GetInt(row, "판매가"), 100);
        var detailHtml = GetStr(row, "상품 상세설명").OrIfEmpty(GetStr(row, "상세설명"));
        var sellerCode = GetStr(row, "자체 상품코드");

        // 검색 태그
        var rawTags = GetStr(row, "네이버태그").OrIfEmpty(GetStr(row, "검색키워드"));
        string[] tagParts;
        if (rawTags.Contains('|') || rawTags.Contains(',') || rawTags.Contains('\n'))
            tagParts = Regex.Split(rawTags, @"[|,\n]+");
        else
            tagParts = rawTags.Split(' ', StringSplitOptions.RemoveEmptyEntries);

        var tagList = new JsonArray();
        var seenTags = new HashSet<string>();
        foreach (var raw in tagParts)
        {
            var tag = SanitizeTag(raw);
            if (!string.IsNullOrEmpty(tag) && seenTags.Add(tag))
            {
                tagList.Add(new JsonObject { ["text"] = tag });
                if (tagList.Count >= 10) break;
            }
        }

        // 옵션
        var options = ParseOptions(GetStr(row, "옵션입력"), GetStr(row, "옵션추가금"));

        var originProduct = new JsonObject
        {
            ["statusType"] = "SALE",
            ["saleType"] = "NEW",
            ["leafCategoryId"] = categoryId,
            ["name"] = productName,
            ["detailContent"] = detailHtml,
            ["salePrice"] = salePrice,
            ["stockQuantity"] = 999,
            ["deliveryInfo"] = new JsonObject
            {
                ["deliveryType"] = "DELIVERY",
                ["deliveryAttributeType"] = "NORMAL",
                ["deliveryCompany"] = "CJGLS",
                ["deliveryFee"] = new JsonObject
                {
                    ["deliveryFeeType"] = "FREE",
                    ["baseFee"] = 0,
                },
                ["claimDeliveryInfo"] = new JsonObject
                {
                    ["returnDeliveryFee"] = 3000,
                    ["exchangeDeliveryFee"] = 3000,
                },
            },
            ["sellerCodeInfo"] = new JsonObject
            {
                ["sellerManagementCode"] = sellerCode,
            },
            ["detailAttribute"] = new JsonObject
            {
                ["naverShoppingSearchInfo"] = new JsonObject
                {
                    ["manufacturerName"] = "상세페이지 참조",
                    ["brandName"] = "",
                },
                ["afterServiceInfo"] = new JsonObject
                {
                    ["afterServiceTelephoneNumber"] = "010-2324-8352",
                    ["afterServiceGuideContent"] = "전화 문의",
                },
                ["originAreaInfo"] = new JsonObject
                {
                    ["originAreaCode"] = "0200037",
                    ["importer"] = "상세페이지 참조",
                    ["content"] = "상세설명 참조",
                    ["plural"] = false,
                },
                ["productInfoProvidedNotice"] = new JsonObject
                {
                    ["productInfoProvidedNoticeType"] = "ETC",
                    ["etc"] = new JsonObject
                    {
                        ["returnCostReason"] = "상세페이지 참조",
                        ["noRefundReason"] = "상세페이지 참조",
                        ["qualityAssuranceStandard"] = "상세페이지 참조",
                        ["compensationProcedure"] = "상세페이지 참조",
                        ["troubleShootingContents"] = "상세페이지 참조",
                        ["itemName"] = "상세페이지 참조",
                        ["modelName"] = "상세페이지 참조",
                        ["manufacturer"] = "상세페이지 참조",
                        ["afterServiceDirector"] = "상세페이지 참조",
                    },
                },
                ["minorPurchasable"] = true,
                ["seoInfo"] = new JsonObject
                {
                    ["sellerTags"] = tagList,
                },
            },
        };

        if (images is not null)
            originProduct["images"] = images;

        // 옵션 설정
        if (options.Count > 0)
        {
            var optionCombinations = new JsonArray();
            foreach (var opt in options)
            {
                optionCombinations.Add(new JsonObject
                {
                    ["optionName1"] = opt.Name,
                    ["stockQuantity"] = 999,
                    ["price"] = salePrice + opt.Price,
                    ["usable"] = true,
                });
            }
            originProduct["optionInfo"] = new JsonObject
            {
                ["optionCombinationSortType"] = "CREATE",
                ["optionCombinationGroupNames"] = new JsonObject
                {
                    ["optionGroupName1"] = "옵션",
                },
                ["optionCombinations"] = optionCombinations,
            };
        }

        return new JsonObject
        {
            ["originProduct"] = originProduct,
            ["smartstoreChannelProduct"] = new JsonObject
            {
                ["channelProductDisplayStatusType"] = "ON",
                ["storeKeepExclusiveProduct"] = false,
                ["naverShoppingRegistration"] = true,
            },
        };
    }

    // ── 헬퍼 ───────────────────────────────────────

    private record OptionItem(string Name, int Price);

    private static List<OptionItem> ParseOptions(string? optionStr, string? extraPriceStr)
    {
        if (string.IsNullOrEmpty(optionStr)) return new();

        var matches = Regex.Matches(optionStr, @"([A-Z])\s+([^,}|]+)");
        var prices = new List<int>();
        if (!string.IsNullOrEmpty(extraPriceStr))
        {
            foreach (var p in Regex.Split(extraPriceStr, @"[,|]"))
            {
                var trimmed = p.Trim();
                if (!string.IsNullOrEmpty(trimmed) && double.TryParse(trimmed, out var v))
                    prices.Add((int)v);
            }
        }

        var options = new List<OptionItem>();
        for (int i = 0; i < matches.Count; i++)
        {
            var name = matches[i].Groups[2].Value.Trim();
            var price = i < prices.Count ? prices[i] : 0;
            options.Add(new OptionItem(name, price));
        }
        return options;
    }

    private static List<string> CollectImageUrls(Dictionary<string, object?> row)
    {
        var urls = new List<string>();
        var seen = new HashSet<string>();

        // "이미지등록(목록)"만 사용 — 대표이미지 + 추가이미지
        // "이미지등록(상세)"는 상세페이지 HTML에 이미 포함되어 있으므로 여기선 제외
        var val = GetStr(row, "이미지등록(목록)");
        if (!string.IsNullOrEmpty(val))
        {
            foreach (var u in Regex.Split(val, @"[|\n]"))
            {
                var trimmed = u.Trim();
                if (!string.IsNullOrEmpty(trimmed) && seen.Add(trimmed))
                    urls.Add(trimmed);
            }
        }

        // 목록 이미지 없으면 상세이미지 첫 1장을 대표이미지 fallback
        if (urls.Count == 0)
        {
            var detailVal = GetStr(row, "이미지등록(상세)");
            if (!string.IsNullOrEmpty(detailVal))
            {
                foreach (var u in Regex.Split(detailVal, @"[|\n]"))
                {
                    var trimmed = u.Trim();
                    if (!string.IsNullOrEmpty(trimmed) && seen.Add(trimmed))
                    {
                        urls.Add(trimmed);
                        break;
                    }
                }
            }
        }

        return urls.Take(9).ToList();
    }

    private static string SanitizeTag(string tag)
    {
        var cleaned = Regex.Replace(tag, @"[^0-9A-Za-z가-힣\s]", "");
        cleaned = Regex.Replace(cleaned, @"\s+", " ").Trim();
        return cleaned;
    }

    private static List<string> ExtractRestrictedTags(string errorText)
    {
        var restricted = new List<string>();
        try
        {
            using var doc = JsonDocument.Parse(errorText);
            if (doc.RootElement.TryGetProperty("invalidInputs", out var inputs))
            {
                foreach (var item in inputs.EnumerateArray())
                {
                    var msg = item.TryGetProperty("message", out var mp) ? mp.GetString() ?? "" : "";
                    var matches = Regex.Matches(msg, @"등록불가인 단어\(([^)]+)\)");
                    foreach (Match m in matches)
                        restricted.Add(m.Groups[1].Value.Trim());
                }
            }
        }
        catch { }
        return restricted;
    }

    private static void RemoveRestrictedTags(JsonObject productJson, List<string> restricted)
    {
        var origin = productJson["originProduct"]?.AsObject();
        var detail = origin?["detailAttribute"]?.AsObject();
        var seo = detail?["seoInfo"]?.AsObject();
        var tags = seo?["sellerTags"]?.AsArray();
        if (tags is null) return;

        var toRemove = new List<JsonNode>();
        foreach (var tag in tags)
        {
            var text = tag?["text"]?.GetValue<string>() ?? "";
            if (restricted.Contains(text)) toRemove.Add(tag!);
        }
        foreach (var node in toRemove) tags.Remove(node);
    }

    private static string ExtractProductId(JsonElement root)
    {
        if (root.TryGetProperty("smartstoreChannelProductNo", out var spn))
            return spn.ToString();
        if (root.TryGetProperty("originProductNo", out var opn))
            return opn.ToString();
        return "";
    }

    private static List<Dictionary<string, object?>> ReadSourceFile(string filePath)
    {
        // A마켓 시트 또는 기본 시트 사용
        using var wb = new XLWorkbook(filePath);
        IXLWorksheet ws;
        if (wb.TryGetWorksheet("A마켓", out var aSheet))
            ws = aSheet;
        else
            ws = wb.Worksheets.First();

        var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
        var lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 1;

        var headers = new Dictionary<int, string>();
        for (int c = 1; c <= lastCol; c++)
        {
            var val = ws.Cell(1, c).GetString().Trim();
            if (!string.IsNullOrEmpty(val)) headers[c] = val;
        }

        var rows = new List<Dictionary<string, object?>>();
        for (int r = 2; r <= lastRow; r++)
        {
            var row = new Dictionary<string, object?>();
            foreach (var (col, name) in headers)
            {
                var cell = ws.Cell(r, col);
                row[name] = cell.IsEmpty() ? null : cell.Value.IsNumber ? cell.Value.GetNumber() : cell.GetString();
            }
            row["_row_num"] = r;
            row["_source_file_path"] = filePath;
            row["_export_root"] = ResolveExportRoot(filePath);
            rows.Add(row);
        }
        return rows;
    }

    private static string ResolveExportRoot(string sourceFilePath)
    {
        var path = System.IO.Path.GetFullPath(sourceFilePath);
        var parent = System.IO.Path.GetDirectoryName(path) ?? "";
        var parentName = System.IO.Path.GetFileName(parent).ToLower();
        var grandParent = System.IO.Path.GetDirectoryName(parent) ?? "";
        var grandName = System.IO.Path.GetFileName(grandParent).ToLower();

        if (parentName == "llm_result" && grandName == "llm_chunks")
            return System.IO.Path.GetDirectoryName(grandParent) ?? grandParent;
        if (parentName == "llm_result")
            return grandParent;
        return parent;
    }

    /// <summary>listing_images 폴더에서 GS코드 가공이미지 파일 찾기 (이미지 선택 반영)</summary>
    private static List<string> FindListingImages(string exportRoot, string gsCode)
    {
        var gsBase = Regex.Replace(gsCode.Trim(), @"[A-Z]$", "", RegexOptions.IgnoreCase);

        // image_selections.json 로드
        ImageSelection? selection = null;
        var selectionsPath = System.IO.Path.Combine(exportRoot, "image_selections.json");
        if (System.IO.File.Exists(selectionsPath))
        {
            try
            {
                var json = System.IO.File.ReadAllText(selectionsPath);
                using var doc = JsonDocument.Parse(json);
                var gs9 = gsBase.Length >= 9 ? gsBase[..9] : gsBase;
                if (doc.RootElement.TryGetProperty(gs9, out var sel))
                {
                    int? mainIdx = sel.TryGetProperty("main", out var m) && m.ValueKind == JsonValueKind.Number ? m.GetInt32() : null;
                    int? mainIdxB = sel.TryGetProperty("mainB", out var mb) && mb.ValueKind == JsonValueKind.Number ? mb.GetInt32() : null;
                    var addIndices = new List<int>();
                    if (sel.TryGetProperty("additional", out var addArr) && addArr.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var a in addArr.EnumerateArray())
                            if (a.ValueKind == JsonValueKind.Number) addIndices.Add(a.GetInt32());
                    }
                    selection = new ImageSelection(mainIdx, addIndices, mainIdxB);
                }
            }
            catch { }
        }

        var listingRoot = System.IO.Path.Combine(exportRoot, "listing_images");
        if (!System.IO.Directory.Exists(listingRoot))
            return new List<string>();

        var searchDirs = new List<string> { listingRoot };
        try
        {
            foreach (var sub in System.IO.Directory.GetDirectories(listingRoot))
                searchDirs.Add(sub);
        }
        catch { }

        foreach (var dir in searchDirs)
        {
            var gsFolder = System.IO.Path.Combine(dir, gsBase);
            if (!System.IO.Directory.Exists(gsFolder))
                gsFolder = System.IO.Path.Combine(dir, gsCode);
            if (!System.IO.Directory.Exists(gsFolder)) continue;

            var allFiles = System.IO.Directory.GetFiles(gsFolder)
                .Where(f => Regex.IsMatch(f, @"\.(jpg|jpeg|png|bmp|webp)$", RegexOptions.IgnoreCase))
                .OrderBy(f => f)
                .ToList();

            if (allFiles.Count == 0) continue;

            if (selection?.MainIndex is not null)
            {
                var (mainPath, addPaths) = Cafe24UploadSupport.PickImagesBySelection(gsFolder, selection);
                if (mainPath is not null)
                {
                    var result = new List<string> { mainPath };
                    result.AddRange(addPaths);
                    return result;
                }
            }

            return allFiles;
        }

        return new List<string>();
    }

    private static string GetStr(Dictionary<string, object?> row, string key)
        => row.TryGetValue(key, out var v) && v is not null ? v.ToString()?.Trim() ?? "" : "";

    private static int GetInt(Dictionary<string, object?> row, string key)
    {
        if (!row.TryGetValue(key, out var v) || v is null) return 0;
        if (v is double d) return (int)d;
        if (int.TryParse(v.ToString(), out var i)) return i;
        if (double.TryParse(v.ToString(), out var d2)) return (int)d2;
        return 0;
    }
}
