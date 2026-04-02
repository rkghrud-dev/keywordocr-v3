using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace KeywordOcr.App.Services;

public sealed record CoupangUploadOptions
{
    public int RowStart { get; set; }
    public int RowEnd { get; set; }
    public bool DryRun { get; set; } = true;
}

public sealed record CoupangUploadResultItem(
    int Row,
    string Name,
    string Status,
    string Category,
    string SellerProductId,
    string Error);

public sealed record CoupangUploadResult(
    IReadOnlyList<CoupangUploadResultItem> Items,
    int SuccessCount,
    int FailCount,
    int TotalCount);

/// <summary>
/// 쿠팡 상품 업로드 서비스 (순수 C# — Python 의존성 없음)
/// </summary>
public sealed class CoupangUploadService
{
    private const int CategoryBatchSize = 5;
    private const int RegisterBatchSize = 5;

    public async Task<CoupangUploadResult> UploadAsync(
        string sourcePath,
        CoupangUploadOptions options,
        IProgress<string>? progress = null,
        CancellationToken ct = default)
    {
        void Log(string msg) => progress?.Report(msg);

        // 키 로드 + API 클라이언트
        using var api = CoupangApiClient.FromKeyFile();

        // 엑셀 읽기
        Log("가공파일 읽는 중...");
        var allRows = CoupangProductBuilder.ReadSourceFile(sourcePath);
        Log($"{allRows.Count}개 상품 로드 완료");

        // 행 필터
        List<Dictionary<string, object?>> targetRows;
        if (options.RowStart > 0)
        {
            var end = options.RowEnd > 0 ? options.RowEnd : options.RowStart;
            targetRows = allRows.Where(r =>
            {
                var rowNum = (int)r["_row_num"]! - 1; // 0-based
                return rowNum >= options.RowStart && rowNum <= end;
            }).ToList();
        }
        else
        {
            targetRows = allRows;
        }

        Log($"처리 대상: {targetRows.Count}개");
        var results = new List<CoupangUploadResultItem>();
        var categoryCache = new Dictionary<long, JsonElement>();

        // ── 1단계: 카테고리 추천 ──────────────────

        // 카테고리 결과 저장 (index → categoryCode, categoryName)
        var catResults = new (long Code, string Name, bool Ok)[targetRows.Count];

        // 엑셀에 카테고리 지정된 행 처리
        var needPredict = new List<int>(); // index into targetRows
        for (int i = 0; i < targetRows.Count; i++)
        {
            var row = targetRows[i];
            var presetCat = GetStr(row, "쿠팡카테고리코드").OrIfEmpty(GetStr(row, "쿠팡카테고리"));
            if (!string.IsNullOrEmpty(presetCat) && long.TryParse(presetCat, out var code))
            {
                catResults[i] = (code, $"엑셀지정({code})", true);
                Log($"  행{row["_row_num"]}: 엑셀 카테고리 사용 ({code})");
            }
            else
            {
                needPredict.Add(i);
            }
        }

        // API로 카테고리 추천 (배치)
        if (needPredict.Count > 0)
        {
            Log($"카테고리 추천 중... ({needPredict.Count}건 API 호출)");

            for (int batchStart = 0; batchStart < needPredict.Count; batchStart += CategoryBatchSize)
            {
                ct.ThrowIfCancellationRequested();
                var batchEnd = Math.Min(batchStart + CategoryBatchSize, needPredict.Count);
                var batchT0 = DateTime.UtcNow;

                var tasks = new List<Task>();
                for (int b = batchStart; b < batchEnd; b++)
                {
                    var idx = needPredict[b];
                    var row = targetRows[idx];
                    var productName = GetStr(row, "상품명");
                    var capturedIdx = idx;

                    tasks.Add(Task.Run(async () =>
                    {
                        try
                        {
                            using var doc = await api.PredictCategoryAsync(productName, ct);
                            var root = doc.RootElement;

                            // 에러 응답 체크
                            if (root.TryGetProperty("code", out var codeProp))
                            {
                                var codeStr = codeProp.ValueKind == System.Text.Json.JsonValueKind.String
                                    ? codeProp.GetString() : codeProp.ToString();
                                if (codeStr == "ERROR")
                                {
                                    var errMsg = root.TryGetProperty("message", out var mp) ? mp.ToString() : "API ERROR";
                                    catResults[capturedIdx] = (0, errMsg, false);
                                    return;
                                }
                            }

                            if (!root.TryGetProperty("data", out var data))
                            {
                                catResults[capturedIdx] = (0, "응답에 data 없음", false);
                                return;
                            }

                            var resultType = data.GetProperty("autoCategorizationPredictionResultType").GetString();
                            if (resultType == "SUCCESS")
                            {
                                var catIdProp = data.GetProperty("predictedCategoryId");
                                long code = catIdProp.ValueKind == System.Text.Json.JsonValueKind.Number
                                    ? catIdProp.GetInt64()
                                    : long.Parse(catIdProp.GetString() ?? "0");
                                var name = data.GetProperty("predictedCategoryName").GetString() ?? "";
                                catResults[capturedIdx] = (code, name, true);
                            }
                            else
                            {
                                catResults[capturedIdx] = (0, resultType ?? "UNKNOWN", false);
                            }
                        }
                        catch (Exception ex)
                        {
                            catResults[capturedIdx] = (0, ex.Message, false);
                        }
                    }, ct));
                }

                await Task.WhenAll(tasks);
                Log($"[{batchEnd}/{needPredict.Count}] 카테고리 추천 중...");

                // 배치 간 최소 1.5초 대기 (429 방지)
                var elapsed = (DateTime.UtcNow - batchT0).TotalSeconds;
                if (elapsed < 1.5)
                    await Task.Delay(TimeSpan.FromSeconds(1.5 - elapsed), ct);
            }
        }

        Log("카테고리 추천 완료");

        // 카테고리 실패 처리 + 메타 로드
        for (int i = 0; i < targetRows.Count; i++)
        {
            var row = targetRows[i];
            var rowNum = (int)row["_row_num"]!;
            var shortName = GetStr(row, "상품명");
            if (shortName.Length > 50) shortName = shortName[..50];

            if (!catResults[i].Ok)
            {
                results.Add(new CoupangUploadResultItem(rowNum, shortName, "CATEGORY_FAIL", "", "", catResults[i].Name));
                continue;
            }

            var catCode = catResults[i].Code;
            if (!categoryCache.ContainsKey(catCode))
            {
                try
                {
                    using var metaDoc = await api.GetCategoryMetaAsync(catCode, ct);
                    categoryCache[catCode] = metaDoc.RootElement.Clone();
                }
                catch
                {
                    categoryCache[catCode] = JsonDocument.Parse("""{"data":{"attributes":[],"noticeCategories":[]}}""").RootElement.Clone();
                }
            }

            row["_category_code"] = catCode;
            row["_category_name"] = catResults[i].Name;
            row["_category_meta"] = categoryCache[catCode];
        }

        // ── 1.5단계: Cafe24 B마켓에서 가공이미지 URL 가져오기 ──

        Log("Cafe24 기본마켓 이미지 URL 조회 중...");
        try
        {
            var cafe24Client = new Cafe24ApiClient();
            var configStore = new Cafe24ConfigStore("", "");
            var tokenState = configStore.LoadTokenState();

            var cafe24Products = await cafe24Client.GetProductsAsync(tokenState.Config, true, ct);
            Log($"Cafe24 기본마켓 상품 {cafe24Products.Count}개 로드");

            // GS코드 → productNo 매핑 (custom_product_code에 GS코드 저장됨)
            var gsToCafe24 = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var p in cafe24Products)
            {
                if (!string.IsNullOrEmpty(p.CustomProductCode))
                {
                    // GS코드 정규화: GS3500169A → GS3500169
                    var gsKey = Regex.Replace(p.CustomProductCode.Trim(), @"[A-Z]$", "", RegexOptions.IgnoreCase);
                    gsToCafe24.TryAdd(gsKey, p.ProductNo);
                }
            }

            // 각 행에 Cafe24 이미지 URL 주입
            foreach (var row in targetRows)
            {
                var sku = GetStr(row, "자체 상품코드");
                if (string.IsNullOrEmpty(sku)) continue;
                var gsKey = Regex.Replace(sku.Trim(), @"[A-Z]$", "", RegexOptions.IgnoreCase);

                if (!gsToCafe24.TryGetValue(gsKey, out var productNo)) continue;

                try
                {
                    var (detailImg, additionalImgs) = await cafe24Client.GetProductImageUrlsAsync(
                        tokenState.Config, productNo, ct);

                    if (!string.IsNullOrEmpty(detailImg))
                    {
                        var allImgUrls = new List<string> { detailImg };
                        allImgUrls.AddRange(additionalImgs);
                        row["_cafe24_image_urls"] = allImgUrls;
                        Log($"  {sku}: Cafe24 이미지 {allImgUrls.Count}장 확보");
                    }
                }
                catch { /* 개별 실패는 무시, esmplus URL fallback */ }
            }
        }
        catch (Exception ex)
        {
            Log($"Cafe24 이미지 조회 실패 (엑셀 URL fallback): {ex.Message}");
        }

        // ── 2단계: JSON 생성 ─────────────────────

        Log("상품 JSON 생성 중...");
        var products = new List<(int Row, string Name, string Category, JsonObject Json)>();

        foreach (var row in targetRows)
        {
            if (!row.ContainsKey("_category_code")) continue;
            var catCode = (long)row["_category_code"]!;
            var catName = (string)row["_category_name"]!;
            var catMeta = (JsonElement)row["_category_meta"]!;

            var productJson = CoupangProductBuilder.BuildProduct(row, catCode, catMeta, api.VendorId);
            var shortName = GetStr(row, "상품명");
            if (shortName.Length > 50) shortName = shortName[..50];
            products.Add(((int)row["_row_num"]!, shortName, $"[{catCode}] {catName}", productJson));
        }

        Log($"JSON 생성 완료: {products.Count}개");

        // ── 3단계: 등록 또는 DRY RUN ──────────────

        if (!options.DryRun)
        {
            Log($"쿠팡 등록 시작 ({products.Count}개)...");
            var regResults = new CoupangUploadResultItem[products.Count];

            for (int batchStart = 0; batchStart < products.Count; batchStart += RegisterBatchSize)
            {
                ct.ThrowIfCancellationRequested();
                var batchEnd = Math.Min(batchStart + RegisterBatchSize, products.Count);
                var batchT0 = DateTime.UtcNow;

                var tasks = new List<Task>();
                for (int b = batchStart; b < batchEnd; b++)
                {
                    var p = products[b];
                    var capturedIdx = b;

                    tasks.Add(Task.Run(async () =>
                    {
                        try
                        {
                            using var respDoc = await api.CreateProductAsync(
                                JsonSerializer.Deserialize<JsonElement>(p.Json.ToJsonString()), ct);
                            var root = respDoc.RootElement;
                            var code = root.TryGetProperty("code", out var codeProp)
                                ? (codeProp.ValueKind == JsonValueKind.String ? codeProp.GetString() : codeProp.ToString())
                                : null;

                            if (code == "SUCCESS")
                            {
                                var spid = root.TryGetProperty("data", out var dataProp) ? dataProp.ToString() : "";
                                regResults[capturedIdx] = new(p.Row, p.Name, "SUCCESS", p.Category, spid, "");
                            }
                            else
                            {
                                var msg = root.TryGetProperty("message", out var msgProp) ? msgProp.ToString() : "";
                                if (msg.Length > 200) msg = msg[..200];
                                regResults[capturedIdx] = new(p.Row, p.Name, $"FAIL_{code}", p.Category, "", msg);
                            }
                        }
                        catch (Exception ex)
                        {
                            regResults[capturedIdx] = new(p.Row, p.Name, "REGISTER_FAIL", p.Category, "", ex.Message.Length > 200 ? ex.Message[..200] : ex.Message);
                        }
                    }, ct));
                }

                await Task.WhenAll(tasks);
                Log($"[{batchEnd}/{products.Count}] 등록 중...");

                var elapsed = (DateTime.UtcNow - batchT0).TotalSeconds;
                if (elapsed < 1.5)
                    await Task.Delay(TimeSpan.FromSeconds(1.5 - elapsed), ct);
            }

            results.AddRange(regResults);
        }
        else
        {
            Log("DRY RUN 완료 - 등록하지 않음");
            foreach (var p in products)
            {
                // 옵션별 가격 로그 출력 (확인용)
                if (p.Json.TryGetPropertyValue("items", out var itemsNode) && itemsNode is JsonArray itemsArr)
                {
                    var baseSale = itemsArr.Count > 0 ? (int)(itemsArr[0]?["salePrice"]?.GetValue<int>() ?? 0) : 0;
                    foreach (var item in itemsArr)
                    {
                        var name = item?["itemName"]?.GetValue<string>() ?? "";
                        var sale = item?["salePrice"]?.GetValue<int>() ?? 0;
                        var diff = sale - baseSale;
                        var diffStr = diff >= 0 ? $"+{diff:#,0}원" : $"{diff:#,0}원";
                        Log($"  옵션: {name} = {sale:#,0}원 ({diffStr})");
                    }
                }

                // 이미지 수 로그
                if (p.Json.TryGetPropertyValue("items", out var items2) && items2 is JsonArray arr2 && arr2.Count > 0)
                {
                    var firstItem = arr2[0];
                    var imgCount = 0;
                    if (firstItem?["images"] is JsonArray imgs) imgCount = imgs.Count;
                    var contentCount = 0;
                    if (firstItem?["contents"] is JsonArray contents)
                    {
                        foreach (var c in contents)
                        {
                            if (c?["contentDetails"] is JsonArray details) contentCount += details.Count;
                        }
                    }
                    Log($"  이미지: 대표+추가 {imgCount}장, 상세 {contentCount}장");
                }

                results.Add(new CoupangUploadResultItem(p.Row, p.Name, "DRY_RUN", p.Category, "", ""));
            }
        }

        var successCount = results.Count(r => r.Status is "SUCCESS" or "DRY_RUN");
        var failCount = results.Count - successCount;
        return new CoupangUploadResult(results, successCount, failCount, results.Count);
    }

    private static string GetStr(Dictionary<string, object?> row, string key)
        => row.TryGetValue(key, out var v) && v is not null ? v.ToString()?.Trim() ?? "" : "";
}
