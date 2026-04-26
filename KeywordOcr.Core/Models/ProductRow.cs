namespace KeywordOcr.Core.Models;

/// <summary>
/// 입력 Excel 한 행 (상품 한 개)
/// </summary>
public class ProductRow
{
    public int RowIndex { get; set; }

    /// <summary>GS코드 (예: GS1234567A)</summary>
    public string GsCode { get; set; } = "";

    /// <summary>상품명</summary>
    public string ProductName { get; set; } = "";

    /// <summary>상품코드 (마켓 고유)</summary>
    public string ProductCode { get; set; } = "";

    /// <summary>옵션명</summary>
    public string OptionName { get; set; } = "";

    /// <summary>공급가</summary>
    public decimal SupplyPrice { get; set; }

    /// <summary>판매가</summary>
    public decimal SalePrice { get; set; }

    /// <summary>소비자가</summary>
    public decimal ConsumerPrice { get; set; }

    /// <summary>상세설명 HTML (이미지태그 포함)</summary>
    public string DetailHtml { get; set; } = "";

    /// <summary>기존 검색어설정 값 (있으면 참고용)</summary>
    public string ExistingSearchKeywords { get; set; } = "";

    /// <summary>이미지 폴더 경로 (로컬)</summary>
    public string ImageFolder { get; set; } = "";

    /// <summary>원본 행의 모든 컬럼 (출력 Excel에 그대로 전달)</summary>
    public Dictionary<string, object?> RawColumns { get; set; } = [];
}
