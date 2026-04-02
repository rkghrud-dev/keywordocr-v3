using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace KeywordOcr.App.Services;

/// <summary>
/// 엑셀 데이터를 쿠팡 상품 등록 JSON으로 변환
/// </summary>
public static class CoupangProductBuilder
{
    public const int OutboundCode = 23273329;
    public const int ReturnCenterCode = 1002256451;

    // ── 엑셀 읽기 ──────────────────────────────────

    public static List<Dictionary<string, object?>> ReadSourceFile(string filePath)
    {
        using var wb = new XLWorkbook(filePath);
        IXLWorksheet ws;
        if (wb.TryGetWorksheet("B마켓", out var bSheet))
            ws = bSheet;
        else if (wb.TryGetWorksheet("분리추출후", out var splitSheet))
            ws = splitSheet;
        else
            ws = wb.Worksheets.First();

        var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
        var lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 1;

        // 헤더
        var headers = new Dictionary<int, string>();
        for (int c = 1; c <= lastCol; c++)
        {
            var val = ws.Cell(1, c).GetString().Trim();
            if (!string.IsNullOrEmpty(val))
                headers[c] = val;
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

    // ── 상품 JSON 빌드 ─────────────────────────────

    public static JsonObject BuildProduct(
        Dictionary<string, object?> row,
        long categoryCode,
        JsonElement categoryMeta,
        string vendorId)
    {
        var productName = GetStr(row, "상품명")
            .OrIfEmpty(GetStr(row, "최종키워드2차"))
            .OrIfEmpty(GetStr(row, "1차키워드"));
        var displayName = productName.Length > 100 ? productName[..100] : productName;

        // ── 브랜드 표준화: 공백/특수문자 제거 (쿠팡 권장) ──
        var rawBrand = GetStr(row, "브랜드").OrIfEmpty("자체브랜드");
        var brand = Regex.Replace(rawBrand, @"[^0-9A-Za-z가-힣]", "");
        if (string.IsNullOrEmpty(brand)) brand = "자체브랜드";

        // ── generalProductName: 옵션 정보 제외한 순수 제품명 ──
        var generalName = Regex.Replace(displayName, @"\d+(mm|cm|m|g|kg|ml|L|개|매|장|ea)\b", "",
            RegexOptions.IgnoreCase).Trim();
        generalName = Regex.Replace(generalName, @"\s{2,}", " ").Trim();
        if (generalName.Length > 100) generalName = generalName[..100];

        // ── sellerProductName: 내부 관리용 (발주서용) ──
        var extSku = GetStr(row, "자체 상품코드");
        var sellerProductName = string.IsNullOrEmpty(extSku)
            ? displayName
            : $"{extSku}_{displayName}";
        if (sellerProductName.Length > 100) sellerProductName = sellerProductName[..100];

        var salePrice = Math.Max(GetInt(row, "판매가"), 1000);
        var originalPrice = GetInt(row, "소비자가");
        if (originalPrice < salePrice) originalPrice = salePrice;

        // 이미지
        var detailHtml = GetStr(row, "상품 상세설명").OrIfEmpty(GetStr(row, "상세설명"));
        var detailImageUrls = BuildDetailImageUrls(row);

        // Cafe24 기본마켓 가공이미지 URL 우선 + 부족하면 상세HTML 이미지로 보충
        List<string> listingImageUrls;
        if (row.TryGetValue("_cafe24_image_urls", out var cafe24Imgs) && cafe24Imgs is List<string> cafe24List && cafe24List.Count > 0)
        {
            listingImageUrls = new List<string>(cafe24List);
        }
        else
        {
            listingImageUrls = BuildImageUrls(row);
        }

        // 대표이미지만 있고 추가이미지가 없으면 상세HTML 이미지로 보충
        if (listingImageUrls.Count <= 1 && detailImageUrls.Count > 0)
        {
            var seen = new HashSet<string>(listingImageUrls, StringComparer.OrdinalIgnoreCase);
            foreach (var imgUrl in detailImageUrls)
            {
                if (seen.Add(imgUrl))
                    listingImageUrls.Add(imgUrl);
                if (listingImageUrls.Count >= 10) break;
            }
        }

        var images = new JsonArray();
        for (int i = 0; i < Math.Min(listingImageUrls.Count, 10); i++)
        {
            images.Add(new JsonObject
            {
                ["imageOrder"] = i,
                ["imageType"] = i == 0 ? "REPRESENTATION" : "DETAIL",
                ["vendorPath"] = listingImageUrls[i],
            });
        }

        // 옵션
        var options = ParseOptions(GetStr(row, "옵션입력"), GetStr(row, "옵션추가금"));

        // 고시정보 / 속성
        var noticeContent = BuildNoticeContent(categoryMeta);
        var baseAttributes = BuildAttributes(categoryMeta);
        var optionAttrName = FindExposedAttributeName(categoryMeta);

        // 검색태그 (중복 제거)
        var searchTags = GetStr(row, "쿠팡검색태그").OrIfEmpty(GetStr(row, "검색어설정"));
        var tagList = new JsonArray();
        var tagSeen = new HashSet<string>();
        foreach (var t in searchTags.Replace(",", " ").Split(' ', StringSplitOptions.RemoveEmptyEntries).Take(10))
        {
            var trimmed = t.Trim();
            if (tagSeen.Add(trimmed)) tagList.Add(trimmed);
        }

        // items
        var items = new JsonArray();
        if (options.Count > 0)
        {
            for (int i = 0; i < options.Count; i++)
            {
                var opt = options[i];
                // 옵션별 고유 속성 추가
                var itemAttrs = baseAttributes.DeepClone().AsArray();
                itemAttrs.Add(new JsonObject
                {
                    ["attributeTypeName"] = optionAttrName,
                    ["attributeValueName"] = opt.Name,
                });
                // SKU: 불변키+옵션 형식 (문서 권장)
                var optSku = string.IsNullOrEmpty(extSku)
                    ? $"OPT{i + 1}"
                    : $"{extSku}_{Regex.Replace(opt.Name, @"[^0-9A-Za-z가-힣]", "")}";
                items.Add(MakeItem(
                    opt.Name, salePrice + opt.Price, originalPrice + opt.Price,
                    optSku,
                    noticeContent, itemAttrs, images, tagList, detailImageUrls));
            }
        }
        else
        {
            items.Add(MakeItem(
                displayName, salePrice, originalPrice, extSku,
                noticeContent, baseAttributes, images, tagList, detailImageUrls));
        }

        return new JsonObject
        {
            ["displayCategoryCode"] = categoryCode,
            ["sellerProductName"] = sellerProductName,
            ["vendorId"] = vendorId,
            ["saleStartedAt"] = "2020-01-01T00:00:00",
            ["saleEndedAt"] = "2099-12-31T00:00:00",
            ["displayProductName"] = displayName,
            ["brand"] = brand,
            ["generalProductName"] = generalName,
            ["productGroup"] = "",
            ["deliveryMethod"] = "SEQUENCIAL",
            ["deliveryCompanyCode"] = "CJGLS",
            ["deliveryChargeType"] = "FREE",
            ["deliveryCharge"] = 0,
            ["freeShipOverAmount"] = 0,
            ["deliveryChargeOnReturn"] = 3000,
            ["returnCharge"] = 3000,
            ["outboundShippingPlaceCode"] = OutboundCode,
            ["returnCenterCode"] = ReturnCenterCode,
            ["returnChargeName"] = "명일우진반품",
            ["companyContactNumber"] = "010-2324-8352",
            ["returnZipCode"] = "05287",
            ["returnAddress"] = "서울특별시 강동구 상일로 74",
            ["returnAddressDetail"] = "고덕리엔파크3단지아파트 고덕리엔파크 321동 CJ대한통운 명일우진대리점",
            ["remoteAreaDeliverable"] = "Y",
            ["unionDeliveryType"] = "UNION_DELIVERY",
            ["vendorUserId"] = "rkghrud",
            ["afterServiceInformation"] = "010-2324-8352",
            ["afterServiceContactNumber"] = "010-2324-8352",
            ["requested"] = true,
            ["items"] = items,
            ["requiredDocuments"] = new JsonArray(),
            ["extraInfoMessage"] = "",
            ["manufacture"] = "",
        };
    }

    // ── 내부 헬퍼 ──────────────────────────────────

    private static JsonObject MakeItem(
        string itemName, int salePrice, int originalPrice, string sku,
        JsonArray noticeContent, JsonNode attributes,
        JsonArray images, JsonArray searchTags, List<string> detailImageUrls)
    {
        // 상세이미지 → contents (IMAGE_NO_SPACE)
        var contents = new JsonArray();
        if (detailImageUrls.Count > 0)
        {
            var contentDetails = new JsonArray();
            foreach (var imgUrl in detailImageUrls)
            {
                contentDetails.Add(new JsonObject
                {
                    ["content"] = imgUrl,
                    ["detailType"] = "IMAGE",
                });
            }
            contents.Add(new JsonObject
            {
                ["contentsType"] = "IMAGE_NO_SPACE",
                ["contentDetails"] = contentDetails,
            });
        }

        return new JsonObject
        {
            ["itemName"] = itemName,
            ["originalPrice"] = originalPrice,
            ["salePrice"] = salePrice,
            ["maximumBuyCount"] = 9999,
            ["maximumBuyForPerson"] = 9999,
            ["outboundShippingTimeDay"] = 2,
            ["maximumBuyForPersonPeriod"] = 1,
            ["unitCount"] = 1,
            ["adultOnly"] = "EVERYONE",
            ["taxType"] = "TAX",
            ["parallelImported"] = "NOT_PARALLEL_IMPORTED",
            ["overseasPurchased"] = "NOT_OVERSEAS_PURCHASED",
            ["pccNeeded"] = false,
            ["externalVendorSku"] = sku,
            ["barcode"] = "",
            ["emptyBarcode"] = true,
            ["emptyBarcodeReason"] = "",
            ["notices"] = noticeContent.DeepClone(),
            ["attributes"] = attributes.DeepClone(),
            ["contents"] = contents,
            ["images"] = images.DeepClone(),
            ["searchTags"] = searchTags.DeepClone(),
        };
    }

    /// <summary>상세페이지용 이미지 추출 (상품 상세설명 HTML의 img 태그에서 추출)</summary>
    private static List<string> BuildDetailImageUrls(Dictionary<string, object?> row)
    {
        var urls = new List<string>();
        var seen = new HashSet<string>();

        // "상품 상세설명" HTML에서 <img src="..."> 추출 — 이것이 진짜 상세이미지
        var detailHtml = GetStr(row, "상품 상세설명").OrIfEmpty(GetStr(row, "상세설명"));
        if (!string.IsNullOrEmpty(detailHtml))
        {
            var imgMatches = Regex.Matches(detailHtml, @"<img[^>]+src=[""']([^""']+)", RegexOptions.IgnoreCase);
            foreach (Match m in imgMatches)
            {
                var imgUrl = m.Groups[1].Value.Trim();
                if (!string.IsNullOrEmpty(imgUrl) && seen.Add(imgUrl))
                    urls.Add(imgUrl);
            }
        }

        return urls;
    }

    /// <summary>목록용 이미지만 추출 (대표이미지 + 추가이미지)</summary>
    private static List<string> BuildImageUrls(Dictionary<string, object?> row)
    {
        var urls = new List<string>();
        var seen = new HashSet<string>();

        // "이미지등록(목록)"만 사용 — 대표이미지 + 추가이미지
        // "이미지등록(상세)"는 상세페이지 HTML용이므로 여기에 넣지 않음
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

        // 목록 이미지가 없으면 상세이미지 첫 1장을 대표이미지로 fallback
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
                        break; // 대표이미지 1장만
                    }
                }
            }
        }

        return urls.Take(10).ToList();
    }

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

    private static JsonArray BuildNoticeContent(JsonElement categoryMeta)
    {
        var arr = new JsonArray();
        if (!categoryMeta.TryGetProperty("data", out var data)) return arr;
        if (!data.TryGetProperty("noticeCategories", out var notices)) return arr;
        if (notices.GetArrayLength() == 0) return arr;

        var notice = notices[0];
        var noticeName = notice.GetProperty("noticeCategoryName").GetString() ?? "";
        if (notice.TryGetProperty("noticeCategoryDetailNames", out var details))
        {
            foreach (var d in details.EnumerateArray())
            {
                arr.Add(new JsonObject
                {
                    ["noticeCategoryName"] = noticeName,
                    ["noticeCategoryDetailName"] = d.GetProperty("noticeCategoryDetailName").GetString() ?? "",
                    ["content"] = "상세페이지 참조",
                });
            }
        }
        return arr;
    }

    private static JsonArray BuildAttributes(JsonElement categoryMeta)
    {
        var arr = new JsonArray();
        if (!categoryMeta.TryGetProperty("data", out var data)) return arr;
        if (!data.TryGetProperty("attributes", out var attrs)) return arr;

        foreach (var a in attrs.EnumerateArray())
        {
            var required = a.GetProperty("required").GetString();
            if (required != "MANDATORY") continue;

            if (a.TryGetProperty("exposed", out var exposed) && exposed.GetString() == "EXPOSED")
                continue;

            var attrName = a.GetProperty("attributeTypeName").GetString() ?? "";
            string val;

            var inputType = a.GetProperty("inputType").GetString();
            if (inputType == "SELECT" && a.TryGetProperty("inputValues", out var inputValues)
                && inputValues.GetArrayLength() > 0)
            {
                var first = inputValues[0];
                val = first.ValueKind == JsonValueKind.Object
                    ? first.GetProperty("inputValueName").GetString() ?? first.ToString()
                    : first.ToString();
            }
            else if (attrName is "수량" or "총 수량")
                val = "1";
            else if (attrName == "색상")
                val = "기타";
            else if (a.GetProperty("dataType").GetString() == "NUMBER")
                val = "1";
            else
                val = "상세페이지 참조";

            var entry = new JsonObject
            {
                ["attributeTypeName"] = attrName,
                ["attributeValueName"] = val,
            };

            if (a.TryGetProperty("basicUnit", out var unit))
            {
                var unitStr = unit.GetString();
                if (!string.IsNullOrEmpty(unitStr) && unitStr != "없음")
                    entry["unitCodeName"] = unitStr;
            }

            arr.Add(entry);
        }
        return arr;
    }

    /// <summary>카테고리 메타에서 EXPOSED 속성명 찾기 (옵션 구분용)</summary>
    private static string FindExposedAttributeName(JsonElement categoryMeta)
    {
        if (!categoryMeta.TryGetProperty("data", out var data)) return "옵션";
        if (!data.TryGetProperty("attributes", out var attrs)) return "옵션";

        foreach (var a in attrs.EnumerateArray())
        {
            if (a.TryGetProperty("exposed", out var exposed) && exposed.GetString() == "EXPOSED")
            {
                var name = a.GetProperty("attributeTypeName").GetString();
                if (!string.IsNullOrEmpty(name))
                    return name;
            }
        }
        return "옵션";
    }

    private static string ResolveExportRoot(string sourceFilePath)
    {
        var path = Path.GetFullPath(sourceFilePath);
        var parent = Path.GetDirectoryName(path) ?? "";
        var parentName = Path.GetFileName(parent).ToLower();
        var grandParent = Path.GetDirectoryName(parent) ?? "";
        var grandName = Path.GetFileName(grandParent).ToLower();

        if (parentName == "llm_result" && grandName == "llm_chunks")
            return Path.GetDirectoryName(grandParent) ?? grandParent;
        if (parentName == "llm_result")
            return grandParent;
        return parent;
    }

    // ── 유틸 ───────────────────────────────────────

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

internal static class StringExt
{
    public static string OrIfEmpty(this string s, string fallback)
        => string.IsNullOrWhiteSpace(s) ? fallback : s;
}
