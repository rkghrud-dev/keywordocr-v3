using System.Reflection;
using System.Text.Json;
using System.Text.Json.Nodes;
using KeywordOcr.App.Services;

TestPrefersSizeAttributeOverQuantity();
TestQuantityOptionUsesNumericValueAndUnit();
TestNaverFallbackImagesIncludeAdditionalColumn();
Console.WriteLine("PASS");

static void TestPrefersSizeAttributeOverQuantity()
{
    var row = BuildRow("A 50mm, B 110mm");
    using var metaDoc = JsonDocument.Parse("""
    {
      "data": {
        "noticeCategories": [],
        "attributes": [
          {
            "attributeTypeName": "수량",
            "dataType": "NUMBER",
            "inputType": "INPUT",
            "inputValues": [],
            "basicUnit": "개",
            "usableUnits": ["개", "박스", "세트"],
            "required": "MANDATORY",
            "groupNumber": "NONE",
            "exposed": "EXPOSED"
          },
          {
            "attributeTypeName": "사이즈",
            "dataType": "STRING",
            "inputType": "INPUT",
            "inputValues": [],
            "basicUnit": "없음",
            "usableUnits": [],
            "required": "MANDATORY",
            "groupNumber": "NONE",
            "exposed": "EXPOSED"
          }
        ]
      }
    }
    """);

    var product = CoupangProductBuilder.BuildProduct(row, 64367, metaDoc.RootElement, "VENDOR");
    var sizeAttr = GetAttribute(product, 0, "사이즈");
    var quantityAttr = GetAttribute(product, 0, "수량");

    AssertEqual("50mm", sizeAttr["attributeValueName"]?.GetValue<string>(), "size attribute value");
    AssertNull(sizeAttr["unitCodeName"], "size attribute unit");
    AssertEqual("1개", quantityAttr["attributeValueName"]?.GetValue<string>(), "fixed quantity value");
    AssertNull(quantityAttr["unitCodeName"], "fixed quantity unit");
}

static void TestQuantityOptionUsesNumericValueAndUnit()
{
    var row = BuildRow("A 2개, B 3개");
    using var metaDoc = JsonDocument.Parse("""
    {
      "data": {
        "noticeCategories": [],
        "attributes": [
          {
            "attributeTypeName": "수량",
            "dataType": "NUMBER",
            "inputType": "INPUT",
            "inputValues": [],
            "basicUnit": "개",
            "usableUnits": ["개", "박스", "세트"],
            "required": "MANDATORY",
            "groupNumber": "NONE",
            "exposed": "EXPOSED"
          }
        ]
      }
    }
    """);

    var product = CoupangProductBuilder.BuildProduct(row, 64367, metaDoc.RootElement, "VENDOR");
    var quantityAttr = GetAttribute(product, 0, "수량");

    AssertEqual("2개", quantityAttr["attributeValueName"]?.GetValue<string>(), "quantity attribute value");
    AssertNull(quantityAttr["unitCodeName"], "quantity attribute unit");
}

static void TestNaverFallbackImagesIncludeAdditionalColumn()
{
    var row = new Dictionary<string, object?>
    {
        ["이미지등록(목록)"] = "https://example.com/main.jpg",
        ["이미지등록(추가)"] = "https://example.com/add-1.jpg|https://example.com/add-2.jpg",
        ["이미지등록(상세)"] = "https://example.com/detail.jpg",
    };

    var method = typeof(NaverUploadService).GetMethod(
        "CollectImageUrls",
        BindingFlags.NonPublic | BindingFlags.Static)
        ?? throw new Exception("CollectImageUrls not found");

    var result = (List<string>?)method.Invoke(null, new object?[] { row })
        ?? throw new Exception("CollectImageUrls returned null");

    AssertEqualInt(3, result.Count, "naver fallback image count");
    AssertEqual("https://example.com/main.jpg", result[0], "naver fallback main image");
    AssertEqual("https://example.com/add-1.jpg", result[1], "naver fallback additional image 1");
    AssertEqual("https://example.com/add-2.jpg", result[2], "naver fallback additional image 2");
}

static Dictionary<string, object?> BuildRow(string optionInput)
{
    return new Dictionary<string, object?>
    {
        ["상품명"] = "테스트 상품",
        ["최종키워드2차"] = "테스트 상품",
        ["1차키워드"] = "테스트",
        ["판매가"] = 1000d,
        ["소비자가"] = 1000d,
        ["옵션입력"] = optionInput,
        ["옵션추가금"] = "0,0",
        ["자체 상품코드"] = "GS0000001A",
        ["이미지등록(목록)"] = "https://example.com/a.jpg",
        ["상품 상세설명"] = "<img src='https://example.com/detail.jpg'>"
    };
}

static JsonObject GetAttribute(JsonObject product, int itemIndex, string attrName)
{
    var items = product["items"]?.AsArray() ?? throw new Exception("items missing");
    var attrs = items[itemIndex]?["attributes"]?.AsArray() ?? throw new Exception("attributes missing");
    foreach (var attr in attrs)
    {
        var obj = attr?.AsObject();
        if (obj is null) continue;
        var name = obj["attributeTypeName"]?.GetValue<string>();
        if (string.Equals(name, attrName, StringComparison.Ordinal))
            return obj;
    }

    throw new Exception($"attribute '{attrName}' missing");
}

static void AssertEqual(string expected, string? actual, string label)
{
    if (!string.Equals(expected, actual, StringComparison.Ordinal))
        throw new Exception($"{label}: expected '{expected}', got '{actual ?? "<null>"}'");
}

static void AssertEqualInt(int expected, int actual, string label)
{
    if (expected != actual)
        throw new Exception($"{label}: expected '{expected}', got '{actual}'");
}

static void AssertNull(JsonNode? actual, string label)
{
    if (actual is not null)
        throw new Exception($"{label}: expected null, got '{actual}'");
}

