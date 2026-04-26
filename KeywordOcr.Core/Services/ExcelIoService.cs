using ClosedXML.Excel;
using KeywordOcr.Core.Models;
using System.Text.RegularExpressions;

namespace KeywordOcr.Core.Services;

/// <summary>
/// Excel 입출력 서비스 (Python io_excel.py + pipeline.py 컬럼 처리 포팅)
/// </summary>
public class ExcelIoService
{
    // ── 컬럼 이름 상수 ────────────────────────────────────────────────────────

    public const string ColProductCode  = "상품코드";
    public const string ColProductName  = "상품명";
    public const string ColOptionName   = "옵션명";
    public const string ColSupplyPrice  = "공급가";
    public const string ColSalePrice    = "판매가";
    public const string ColConsumerPrice = "소비자가";
    public const string ColDetail       = "상품 상세설명";
    public const string ColSearchKw     = "검색어설정";
    public const string ColSearchKwB    = "검색어설정_B";
    public const string ColNaverTag     = "검색키워드";
    public const string ColNaverTagB    = "검색키워드_B";
    public const string ColCoupangTag   = "쿠팡태그";
    public const string ColCoupangTagB  = "쿠팡태그_B";

    private static readonly Regex _gsRe = new(@"GS\d{7}[A-Za-z]?", RegexOptions.Compiled);

    // ── 입력 파일 로드 ────────────────────────────────────────────────────────

    /// <summary>CSV 또는 Excel(.xlsx/.xls) 파일을 읽어 ProductRow 목록으로 변환</summary>
    public List<ProductRow> LoadInputFile(string filePath, string sheetName = "")
    {
        var ext = Path.GetExtension(filePath).TrimStart('.').ToLowerInvariant();
        return ext is "csv"
            ? LoadCsv(filePath)
            : LoadExcel(filePath, sheetName);
    }

    // ── Excel 읽기 ────────────────────────────────────────────────────────────

    private static List<ProductRow> LoadExcel(string path, string sheetName)
    {
        using var wb = new XLWorkbook(path);
        var ws = string.IsNullOrEmpty(sheetName)
            ? wb.Worksheets.First()
            : wb.Worksheets.FirstOrDefault(s => s.Name == sheetName)
              ?? wb.Worksheets.First();
        return ParseSheet(ws);
    }

    private static List<ProductRow> ParseSheet(IXLWorksheet ws)
    {
        var rows = ws.RowsUsed().ToList();
        if (rows.Count < 2) return [];

        // 헤더 행
        var headerRow = rows[0];
        var headers = headerRow.CellsUsed()
            .ToDictionary(c => c.WorksheetColumn().ColumnNumber(),
                          c => c.GetString().Trim(),
                          EqualityComparer<int>.Default);

        // 컬럼 인덱스 해결
        int ColIdx(params string[] candidates)
        {
            foreach (var cand in candidates)
                foreach (var (col, name) in headers)
                    if (name.Contains(cand)) return col;
            return -1;
        }

        int idxCode    = ColIdx("상품코드", "코드");
        int idxName    = ColIdx("상품명");
        int idxOption  = ColIdx("옵션명");
        int idxSupply  = ColIdx("공급가");
        int idxSale    = ColIdx("판매가");
        int idxConsumer= ColIdx("소비자가");
        int idxDetail  = ColIdx("상품 상세설명", "상세설명");
        int idxKw      = ColIdx("검색어설정");

        var result = new List<ProductRow>();
        for (int i = 1; i < rows.Count; i++)
        {
            var row = rows[i];
            string GetCell(int colIdx)
                => colIdx > 0 ? row.Cell(colIdx).GetString().Trim() : "";

            var raw = new Dictionary<string, object?>(StringComparer.Ordinal);
            foreach (var (hCol, hName) in headers)
                raw[hName] = row.Cell(hCol).Value.ToString();

            var name = GetCell(idxName);
            var code = GetCell(idxCode);

            // GS코드 추출: 코드 컬럼 또는 상품명에서
            var gsCode = ExtractGsCode(code) ?? ExtractGsCode(name) ?? code;

            // A접미사만 처리 (B/C/D 제외)
            if (!IsGsCodeA(gsCode)) continue;

            result.Add(new ProductRow
            {
                RowIndex             = row.RowNumber(),
                GsCode               = gsCode,
                ProductName          = name,
                ProductCode          = code,
                OptionName           = GetCell(idxOption),
                SupplyPrice          = ParseDecimal(GetCell(idxSupply)),
                SalePrice            = ParseDecimal(GetCell(idxSale)),
                ConsumerPrice        = ParseDecimal(GetCell(idxConsumer)),
                DetailHtml           = GetCell(idxDetail),
                ExistingSearchKeywords = GetCell(idxKw),
                RawColumns           = raw,
            });
        }
        return result;
    }

    // ── CSV 읽기 ──────────────────────────────────────────────────────────────

    private static List<ProductRow> LoadCsv(string path)
    {
        // 임시 Excel 변환 방식 대신 ClosedXML 읽기 불가이므로 직접 파싱
        var lines = File.ReadAllLines(path, System.Text.Encoding.UTF8);
        if (lines.Length < 2) return [];

        var headers = SplitCsv(lines[0]);
        int Col(params string[] cands)
        {
            for (int ci = 0; ci < headers.Count; ci++)
                foreach (var c in cands)
                    if (headers[ci].Contains(c)) return ci;
            return -1;
        }

        int idxCode    = Col("상품코드","코드");
        int idxName    = Col("상품명");
        int idxOption  = Col("옵션명");
        int idxSupply  = Col("공급가");
        int idxSale    = Col("판매가");
        int idxConsumer= Col("소비자가");
        int idxDetail  = Col("상품 상세설명","상세설명");
        int idxKw      = Col("검색어설정");

        string GetF(List<string> fields, int idx)
            => idx >= 0 && idx < fields.Count ? fields[idx].Trim() : "";

        var result = new List<ProductRow>();
        for (int i = 1; i < lines.Length; i++)
        {
            var fields = SplitCsv(lines[i]);
            var name   = GetF(fields, idxName);
            var code   = GetF(fields, idxCode);
            var gsCode = ExtractGsCode(code) ?? ExtractGsCode(name) ?? code;
            if (!IsGsCodeA(gsCode)) continue;

            var raw = new Dictionary<string, object?>(StringComparer.Ordinal);
            for (int ci = 0; ci < headers.Count && ci < fields.Count; ci++)
                raw[headers[ci]] = fields[ci];

            result.Add(new ProductRow
            {
                RowIndex             = i + 1,
                GsCode               = gsCode,
                ProductName          = name,
                ProductCode          = code,
                OptionName           = GetF(fields, idxOption),
                SupplyPrice          = ParseDecimal(GetF(fields, idxSupply)),
                SalePrice            = ParseDecimal(GetF(fields, idxSale)),
                ConsumerPrice        = ParseDecimal(GetF(fields, idxConsumer)),
                DetailHtml           = GetF(fields, idxDetail),
                ExistingSearchKeywords = GetF(fields, idxKw),
                RawColumns           = raw,
            });
        }
        return result;
    }

    private static List<string> SplitCsv(string line)
    {
        var result = new List<string>();
        bool inQuote = false;
        var sb = new System.Text.StringBuilder();
        foreach (var ch in line)
        {
            if (ch == '"') { inQuote = !inQuote; continue; }
            if (ch == ',' && !inQuote) { result.Add(sb.ToString()); sb.Clear(); continue; }
            sb.Append(ch);
        }
        result.Add(sb.ToString());
        return result;
    }

    // ── 결과 Excel 저장 ───────────────────────────────────────────────────────

    /// <summary>파이프라인 결과를 업로드용 Excel로 저장</summary>
    public void WriteOutputFile(
        string outputPath,
        List<PipelineResult> results,
        IReadOnlyList<string>? originalHeaders = null)
    {
        using var wb = new XLWorkbook();
        WriteSheet(wb, "분리추출후", results, market: "A", originalHeaders);
        WriteSheet(wb, "B마켓",     results, market: "B", originalHeaders);
        wb.SaveAs(outputPath);
    }

    private static void WriteSheet(
        XLWorkbook wb,
        string sheetName,
        List<PipelineResult> results,
        string market,
        IReadOnlyList<string>? originalHeaders)
    {
        var ws = wb.Worksheets.Add(sheetName);

        // ── 헤더 ────────────────────────────────────────────────────────────
        var extraHeaders = new[] { ColSearchKw, ColNaverTag, ColCoupangTag };

        // 원본 헤더 + 추가 컬럼
        var allHeaders = new List<string>();
        if (originalHeaders?.Count > 0) allHeaders.AddRange(originalHeaders);
        foreach (var h in extraHeaders)
            if (!allHeaders.Contains(h)) allHeaders.Add(h);

        for (int ci = 0; ci < allHeaders.Count; ci++)
            ws.Cell(1, ci + 1).Value = allHeaders[ci];

        // 헤더 스타일
        var headerRange = ws.Range(1, 1, 1, allHeaders.Count);
        headerRange.Style.Font.Bold = true;
        headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;

        // ── 데이터 행 ────────────────────────────────────────────────────────
        for (int ri = 0; ri < results.Count; ri++)
        {
            var r   = results[ri];
            int row = ri + 2;

            var kw      = market == "B" ? r.SearchKeywordsB : r.SearchKeywordsA;
            var naver   = market == "B" ? r.NaverTagsB      : r.NaverTagsA;
            var coupang = market == "B" ? r.CoupangTagsB    : r.CoupangTagsA;

            // 원본 컬럼 복원
            for (int ci = 0; ci < allHeaders.Count; ci++)
            {
                var h = allHeaders[ci];
                if (h == ColSearchKw)
                    ws.Cell(row, ci + 1).Value = kw;
                else if (h == ColNaverTag)
                    ws.Cell(row, ci + 1).Value = string.Join("|", naver);
                else if (h == ColCoupangTag)
                    ws.Cell(row, ci + 1).Value = string.Join(",", coupang);
                else if (r.RawColumns.TryGetValue(h, out var v))
                    ws.Cell(row, ci + 1).Value = v?.ToString() ?? "";
            }
        }

        // 컬럼 너비 자동
        ws.Columns().AdjustToContents();
    }

    // ── 유틸 ─────────────────────────────────────────────────────────────────

    private static string? ExtractGsCode(string? text)
    {
        if (string.IsNullOrEmpty(text)) return null;
        var m = _gsRe.Match(text);
        return m.Success ? m.Value : null;
    }

    /// <summary>GS코드가 A 접미사이거나 접미사 없으면 true</summary>
    private static bool IsGsCodeA(string gsCode)
    {
        if (string.IsNullOrEmpty(gsCode)) return false;
        var m = _gsRe.Match(gsCode);
        if (!m.Success) return false;
        var code = m.Value;
        if (code.Length == 9) return true;          // 접미사 없음
        var suffix = code[9..].ToUpperInvariant();
        return suffix == "A";
    }

    private static decimal ParseDecimal(string s)
    {
        if (string.IsNullOrEmpty(s)) return 0m;
        var cleaned = Regex.Replace(s, @"[^\d.]", "");
        return decimal.TryParse(cleaned, out var d) ? d : 0m;
    }
}
