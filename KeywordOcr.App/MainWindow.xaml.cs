using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using ClosedXML.Excel;
using KeywordOcr.App.Services;
using Microsoft.Win32;

namespace KeywordOcr.App;

public partial class MainWindow : Window
{
    private readonly string _legacyRoot;
    private readonly string _pythonRoot;   // v3/backend — Python import 경로
    private readonly string _v3Root;
    private string? _sourcePath;
    private string? _lastOutputRoot;
    private string? _lastOutputFile;
    private string? _lastUploadLogPath;
    private CancellationTokenSource? _cts;
    private readonly ObservableCollection<ProductItem> _products = new();
    private readonly ObservableCollection<PriceRow> _priceRows = new();
    private readonly ObservableCollection<string> _imageGsCodes = new();
    private readonly ObservableCollection<ImageThumbnailItem> _imageThumbnails = new();
    private readonly Dictionary<string, ImageSelection> _imageSelections = new(StringComparer.OrdinalIgnoreCase);
    private bool _selectingBMarket; // true면 B마켓 대표 선택 모드
    private string _bMarketTokenPath = ""; // 준비몰 토큰 JSON 경로 (비어 있으면 기본 경로 사용)
    private string? _imageListingRoot;
    private JobHistoryService? _jobHistory;
    private string _settingsPath = "";
    private bool _syncingKeywordVersion;
    private bool _syncingCafe24MarketTargetChecks;

    // 상품 선택 목록
    private readonly ObservableCollection<UploadProductItem> _cafe24Items = new();
    private readonly ObservableCollection<UploadProductItem> _coupangItems = new();
    private readonly ObservableCollection<UploadProductItem> _basicCafe24Items = new();
    private int _cafe24LastClickIndex = -1;
    private int _coupangLastClickIndex = -1;
    private int _basicCafe24LastClickIndex = -1;
    private Services.UploadHistoryStore _uploadHistory = new();

    public MainWindow()
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        InitializeComponent();

        (_v3Root, _legacyRoot, _pythonRoot) = ResolveApplicationRoots();

        _settingsPath = Path.Combine(_legacyRoot, "app_settings.json");

        ProductList.ItemsSource = _products;
        PriceGrid.ItemsSource = _priceRows;
        ImageGsListBox.ItemsSource = _imageGsCodes;
        ThumbnailPanel.ItemsSource = _imageThumbnails;
        Cafe24ProductList.ItemsSource = _cafe24Items;
        CoupangProductList.ItemsSource = _coupangItems;
        BasicCafe24ProductGrid.ItemsSource = _basicCafe24Items;

        // 설정 탭 초기값
        SettingsLegacyRoot.Text = _legacyRoot;
        SettingsV3Root.Text = _v3Root;
        Cafe24DateTag.Text = DateTime.Now.ToString("yyyyMMdd");
        LoadTokenInfo();
        LoadAppSettings();
        ApplyDefaultWorkflowSelections();
        if (string.IsNullOrEmpty(SettingsBTokenPath.Text))
            LoadTokenInfoB(); // 설정 파일 없을 때 기본 경로로 시도
        SyncCafe24MarketTargetCheckBoxes(true, true);

        _jobHistory = new JobHistoryService(_legacyRoot);
        RefreshHistoryGrid();

        Log("KeywordOCR v3 시작");
        Log($"Python 루트: {_pythonRoot}");
    }

    private void ApplyDefaultWorkflowSelections()
    {
        TestChunkSizeCombo.SelectedIndex = 4; // 분할안함
        SetKeywordVersionSelection("3.0");
    }

    private static (string V3Root, string LegacyRoot, string PythonRoot) ResolveApplicationRoots()
    {
        var baseDir = Path.GetFullPath(AppContext.BaseDirectory);
        var candidates = new List<string>();
        var current = baseDir;
        for (var i = 0; i < 12; i++)
        {
            candidates.Add(current);
            var parent = Directory.GetParent(current);
            if (parent is null)
                break;
            current = parent.FullName;
        }

        var v3Root = candidates.FirstOrDefault(IsV3Root)
            ?? Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", "..", ".."));

        var legacyRoot = File.Exists(Path.Combine(v3Root, "app_settings.json"))
            || Directory.Exists(Path.Combine(v3Root, "backend"))
                ? v3Root
                : Path.GetFullPath(Path.Combine(v3Root, ".."));

        var pythonRoot = ResolvePythonRoot(v3Root, legacyRoot, baseDir);
        return (v3Root, legacyRoot, pythonRoot);
    }

    private static bool IsV3Root(string root)
    {
        return Directory.Exists(Path.Combine(root, "backend", "app"))
            && (Directory.Exists(Path.Combine(root, "KeywordOcr.App"))
                || Directory.Exists(Path.Combine(root, "Bridge")));
    }

    private static string ResolvePythonRoot(string v3Root, string legacyRoot, string baseDir)
    {
        var candidates = new[]
        {
            Path.Combine(v3Root, "backend"),
            Path.Combine(baseDir, "backend"),
            legacyRoot
        };

        return candidates.FirstOrDefault(path => Directory.Exists(Path.Combine(path, "app")))
            ?? Path.Combine(v3Root, "backend");
    }

    #region ═══ 드래그 앤 드롭 ═══

    private void SyncCafe24MarketTargetCheckBoxes(bool homeSelected, bool readySelected)
    {
        _syncingCafe24MarketTargetChecks = true;
        try
        {
            if (Cafe24HomeCheckBox is not null) Cafe24HomeCheckBox.IsChecked = homeSelected;
            if (TestCafe24HomeCheckBox is not null) TestCafe24HomeCheckBox.IsChecked = homeSelected;
            if (Cafe24ReadyCheckBox is not null) Cafe24ReadyCheckBox.IsChecked = readySelected;
            if (TestCafe24ReadyCheckBox is not null) TestCafe24ReadyCheckBox.IsChecked = readySelected;
        }
        finally
        {
            _syncingCafe24MarketTargetChecks = false;
        }
    }

    private void Cafe24MarketTargetCheck_Changed(object sender, RoutedEventArgs e)
    {
        if (_syncingCafe24MarketTargetChecks || sender is not CheckBox checkBox)
        {
            return;
        }

        var homeSelected = IsCafe24HomeSelected();
        var readySelected = IsCafe24ReadySelected();

        if (checkBox == Cafe24HomeCheckBox || checkBox == TestCafe24HomeCheckBox)
        {
            homeSelected = checkBox.IsChecked == true;
        }
        else if (checkBox == Cafe24ReadyCheckBox || checkBox == TestCafe24ReadyCheckBox)
        {
            readySelected = checkBox.IsChecked == true;
        }

        SyncCafe24MarketTargetCheckBoxes(homeSelected, readySelected);
    }

    private bool IsCafe24HomeSelected() => Cafe24HomeCheckBox?.IsChecked == true || TestCafe24HomeCheckBox?.IsChecked == true;

    private bool IsCafe24ReadySelected() => Cafe24ReadyCheckBox?.IsChecked == true || TestCafe24ReadyCheckBox?.IsChecked == true;

    private bool TryGetSelectedCafe24Markets(out bool homeSelected, out bool readySelected, out string marketLabel)
    {
        homeSelected = IsCafe24HomeSelected();
        readySelected = IsCafe24ReadySelected();
        marketLabel = GetSelectedCafe24MarketLabel(homeSelected, readySelected);
        if (homeSelected || readySelected)
        {
            return true;
        }

        MessageBox.Show("Cafe24 대상 몰을 하나 이상 선택하세요.", "알림", MessageBoxButton.OK, MessageBoxImage.Warning);
        return false;
    }

    private static string GetSelectedCafe24MarketLabel(bool homeSelected, bool readySelected)
    {
        var markets = new List<string>();
        if (homeSelected) markets.Add("홈런마켓");
        if (readySelected) markets.Add("준비몰");
        return markets.Count == 0 ? "선택 없음" : string.Join(" + ", markets);
    }

    private void ShowTab_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not MenuItem { Tag: string tabName })
        {
            return;
        }

        if (FindName(tabName) is TabItem tab)
        {
            tab.Visibility = Visibility.Visible;
            tab.IsSelected = true;
        }
    }


    private void DropZone_DragEnter(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            e.Effects = DragDropEffects.Copy;
            DropZone.BorderBrush = new SolidColorBrush(Color.FromRgb(0x1a, 0x1a, 0x2e));
            DropZone.Background = new SolidColorBrush(Color.FromRgb(0xEE, 0xEE, 0xFF));
        }
        else
        {
            e.Effects = DragDropEffects.None;
        }
        e.Handled = true;
    }

    private void DropZone_DragLeave(object sender, DragEventArgs e)
    {
        DropZone.BorderBrush = new SolidColorBrush(Color.FromRgb(0xCC, 0xCC, 0xCC));
        DropZone.Background = Brushes.White;
    }

    private void DropZone_Drop(object sender, DragEventArgs e)
    {
        DropZone.BorderBrush = new SolidColorBrush(Color.FromRgb(0xCC, 0xCC, 0xCC));
        DropZone.Background = Brushes.White;

        if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;

        var files = (string[])e.Data.GetData(DataFormats.FileDrop)!;
        var file = files.FirstOrDefault(f =>
        {
            var ext = Path.GetExtension(f).ToLowerInvariant();
            return ext is ".csv" or ".xlsx" or ".xls";
        });

        if (file == null)
        {
            MessageBox.Show("CSV 또는 Excel 파일만 지원합니다.", "알림", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        LoadFile(file);
    }

    #endregion

    #region ═══ 파일 선택 / 로딩 ═══

    private void SelectFile_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog
        {
            Filter = "CSV/Excel|*.csv;*.xlsx;*.xls|모든 파일|*.*",
            Title = "원본 파일 선택",
        };
        if (dlg.ShowDialog() == true)
            LoadFile(dlg.FileName);
    }

    private void LoadFile(string filePath)
    {
        _sourcePath = filePath;
        DropZoneFile.Text = filePath;
        DropZoneText.Text = "선택된 파일:";
        TestDropZoneFile.Text = filePath;
        Log($"파일 선택: {Path.GetFileName(filePath)}");
        LoadProductList(filePath);
    }

    private void LoadProductList(string filePath)
    {
        _products.Clear();
        try
        {
            var ext = Path.GetExtension(filePath).ToLowerInvariant();
            var items = ext is ".xlsx" or ".xls"
                ? ReadProductsFromExcel(filePath)
                : ReadProductsFromCsv(filePath);

            if (items.Count == 0)
            {
                Log("상품코드를 찾지 못했습니다. 전체 파일이 처리됩니다.");
                ProductListPanel.Visibility = Visibility.Collapsed;
                TestProductListPanel.Visibility = Visibility.Collapsed;
                SetPipelineEnabled(true);
                return;
            }

            foreach (var (code, name) in items)
                _products.Add(new ProductItem { Code = code, Name = name, IsSelected = true });

            ProductListPanel.Visibility = Visibility.Visible;
            TestProductListPanel.Visibility = Visibility.Visible;
            TestProductList.ItemsSource = _products;
            UpdateProductCount();
            ApplyHistoryToProducts();
            SetPipelineEnabled(true);
            Log($"상품 {items.Count}개 로드됨");
        }
        catch (Exception ex)
        {
            Log($"파일 읽기 오류: {ex.Message}");
            ProductListPanel.Visibility = Visibility.Collapsed;
            TestProductListPanel.Visibility = Visibility.Collapsed;
            SetPipelineEnabled(true);
        }
    }

    private static readonly string[] CodeColumns =
        { "상품코드", "자체상품코드", "자체 상품코드", "상품코드B", "코드", "코드B", "GS코드", "product_code", "gs_code" };

    private static readonly string[] NameColumns =
        { "상품명", "제품명", "product_name", "name" };

    private List<(string code, string name)> ReadProductsFromExcel(string filePath)
    {
        var results = new List<(string code, string name)>();
        using var wb = new XLWorkbook(filePath);
        var ws = wb.Worksheets.First();
        var headerRow = ws.FirstRowUsed();
        if (headerRow == null) return results;

        int nameCol = -1;
        var codeCols = new List<int>();
        var lastCol = headerRow.LastCellUsed()?.Address.ColumnNumber ?? 0;

        for (int c = 1; c <= lastCol; c++)
        {
            var header = headerRow.Cell(c).GetString().Trim();
            if (CodeColumns.Any(h => h.Equals(header, StringComparison.OrdinalIgnoreCase)))
                codeCols.Add(c);
            if (nameCol < 0 && NameColumns.Any(h => h.Equals(header, StringComparison.OrdinalIgnoreCase)))
                nameCol = c;
        }

        if (codeCols.Count == 0) return results;

        var lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
        var seen = new HashSet<string>();

        for (int r = headerRow.RowNumber() + 1; r <= lastRow; r++)
        {
            // 코드 컬럼들 중 비어있지 않은 첫 번째 값 사용
            var code = "";
            foreach (var cc in codeCols)
            {
                code = ws.Cell(r, cc).GetString().Trim();
                if (!string.IsNullOrEmpty(code)) break;
            }
            if (string.IsNullOrEmpty(code)) continue;

            var name = nameCol > 0 ? ws.Cell(r, nameCol).GetString().Trim() : "";

            // 코드 컬럼과 상품명 컬럼 둘 다에서 GS코드 찾기
            var gsMatch = Regex.Match(code, @"(GS\d{7})([A-Za-z])?", RegexOptions.IgnoreCase);
            if (!gsMatch.Success)
                gsMatch = Regex.Match(name, @"(GS\d{7})([A-Za-z])?", RegexOptions.IgnoreCase);

            if (gsMatch.Success && gsMatch.Groups[2].Success)
            {
                var suffix = gsMatch.Groups[2].Value.ToUpper();
                if (suffix != "A") continue;
            }
            var displayCode = gsMatch.Success ? gsMatch.Groups[1].Value : code;
            if (!seen.Add(displayCode)) continue;

            results.Add((displayCode, name));
        }
        return results;
    }

    private List<(string code, string name)> ReadProductsFromCsv(string filePath)
    {
        var results = new List<(string code, string name)>();
        string[] lines;
        try { lines = File.ReadAllLines(filePath, Encoding.UTF8); }
        catch { lines = File.ReadAllLines(filePath, Encoding.GetEncoding(949)); }

        if (lines.Length < 2) return results;

        var headers = ParseCsvLine(lines[0]);
        int codeIdx = -1, nameIdx = -1;
        for (int i = 0; i < headers.Length; i++)
        {
            var h = headers[i].Trim();
            if (codeIdx < 0 && CodeColumns.Any(c => c.Equals(h, StringComparison.OrdinalIgnoreCase)))
                codeIdx = i;
            if (nameIdx < 0 && NameColumns.Any(c => c.Equals(h, StringComparison.OrdinalIgnoreCase)))
                nameIdx = i;
        }

        if (codeIdx < 0) return results;

        var seen = new HashSet<string>();
        for (int r = 1; r < lines.Length; r++)
        {
            var cols = ParseCsvLine(lines[r]);
            if (codeIdx >= cols.Length) continue;
            var code = cols[codeIdx].Trim();
            if (string.IsNullOrEmpty(code)) continue;

            var name = (nameIdx >= 0 && nameIdx < cols.Length) ? cols[nameIdx].Trim() : "";

            // 코드 컬럼과 상품명 컬럼 둘 다에서 GS코드 찾기
            var gsMatch = Regex.Match(code, @"(GS\d{7})([A-Za-z])?", RegexOptions.IgnoreCase);
            if (!gsMatch.Success)
                gsMatch = Regex.Match(name, @"(GS\d{7})([A-Za-z])?", RegexOptions.IgnoreCase);

            if (gsMatch.Success && gsMatch.Groups[2].Success)
            {
                var suffix = gsMatch.Groups[2].Value.ToUpper();
                if (suffix != "A") continue;
            }
            var displayCode = gsMatch.Success ? gsMatch.Groups[1].Value : code;
            if (!seen.Add(displayCode)) continue;

            results.Add((displayCode, name));
        }
        return results;
    }

    private static string[] ParseCsvLine(string line)
    {
        var result = new List<string>();
        bool inQuote = false;
        var sb = new StringBuilder();
        foreach (char c in line)
        {
            if (c == '"') inQuote = !inQuote;
            else if (c == ',' && !inQuote) { result.Add(sb.ToString()); sb.Clear(); }
            else sb.Append(c);
        }
        result.Add(sb.ToString());
        return result.ToArray();
    }

    #endregion

    #region ═══ 상품 선택 ═══

    private void SelectAll_Click(object sender, RoutedEventArgs e)
    {
        foreach (var p in _products) p.IsSelected = true;
        ProductList.Items.Refresh();
        UpdateProductCount();
    }

    private void DeselectAll_Click(object sender, RoutedEventArgs e)
    {
        foreach (var p in _products) p.IsSelected = false;
        ProductList.Items.Refresh();
        UpdateProductCount();
    }

    private void ProductCheck_Changed(object sender, RoutedEventArgs e) => UpdateProductCount();

    /// <summary>
    /// _jobHistory에서 각 상품의 마지막 처리 날짜를 _products에 반영
    /// </summary>
    private void ApplyHistoryToProducts()
    {
        if (_jobHistory == null || _products.Count == 0) return;

        // GS코드 → 가장 최근 처리 시각
        var lastProcessed = new Dictionary<string, DateTime>(StringComparer.OrdinalIgnoreCase);
        foreach (var record in _jobHistory.Records)
        {
            foreach (var code in record.SelectedCodes)
            {
                if (!lastProcessed.TryGetValue(code, out var existing) || record.Timestamp > existing)
                    lastProcessed[code] = record.Timestamp;
            }
        }

        foreach (var product in _products)
        {
            product.LastProcessedAt = lastProcessed.TryGetValue(product.Code, out var date) ? date : null;
        }

        // 이력이 있는 항목을 위로, 없는 항목을 아래로 재정렬
        var sorted = _products
            .OrderByDescending(p => p.LastProcessedAt.HasValue)
            .ThenByDescending(p => p.LastProcessedAt)
            .ToList();

        _products.Clear();
        foreach (var p in sorted)
            _products.Add(p);

        ProductList.Items.Refresh();
        UpdateProductCount();
    }

    private void UpdateProductCount()
    {
        var selected = _products.Count(p => p.IsSelected);
        var text = $"({selected}/{_products.Count} 선택)";
        ProductCountText.Text = text;
        TestProductCountText.Text = text;
    }

    #endregion

    #region ═══ 필터링 ═══

    private string? CreateFilteredFile()
    {
        if (_products.Count == 0 || _products.All(p => p.IsSelected))
            return _sourcePath;

        var selectedCodes = new HashSet<string>(
            _products.Where(p => p.IsSelected).Select(p => p.Code),
            StringComparer.OrdinalIgnoreCase);

        if (selectedCodes.Count == 0) return null;

        var ext = Path.GetExtension(_sourcePath!).ToLowerInvariant();
        var dir = Path.GetDirectoryName(_sourcePath!)!;
        var baseName = Path.GetFileNameWithoutExtension(_sourcePath!);
        var filteredPath = Path.Combine(dir, $"{baseName}_filtered{ext}");

        try
        {
            if (ext is ".xlsx" or ".xls")
                CreateFilteredExcel(_sourcePath!, filteredPath, selectedCodes);
            else
                CreateFilteredCsv(_sourcePath!, filteredPath, selectedCodes);

            Log($"선택된 {selectedCodes.Count}개 상품으로 필터링 완료");
            return filteredPath;
        }
        catch (Exception ex)
        {
            Log($"필터링 오류: {ex.Message}, 원본 파일 사용");
            return _sourcePath;
        }
    }

    private void CreateFilteredExcel(string source, string dest, HashSet<string> codes)
    {
        using var wb = new XLWorkbook(source);
        var ws = wb.Worksheets.First();
        var headerRow = ws.FirstRowUsed()!;
        var lastCol = headerRow.LastCellUsed()?.Address.ColumnNumber ?? 0;
        var lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;

        int codeCol = -1, nameCol = -1;
        for (int c = 1; c <= lastCol; c++)
        {
            var header = headerRow.Cell(c).GetString().Trim();
            if (codeCol < 0 && CodeColumns.Any(h => h.Equals(header, StringComparison.OrdinalIgnoreCase)))
                codeCol = c;
            if (nameCol < 0 && NameColumns.Any(h => h.Equals(header, StringComparison.OrdinalIgnoreCase)))
                nameCol = c;
        }
        if (codeCol < 0) return;

        var rowsToDelete = new List<int>();
        for (int r = headerRow.RowNumber() + 1; r <= lastRow; r++)
        {
            var code = ws.Cell(r, codeCol).GetString().Trim();
            var name = nameCol > 0 ? ws.Cell(r, nameCol).GetString().Trim() : "";

            var gsMatch = Regex.Match(code, @"(GS\d{7})", RegexOptions.IgnoreCase);
            if (!gsMatch.Success)
                gsMatch = Regex.Match(name, @"(GS\d{7})", RegexOptions.IgnoreCase);

            var checkCode = gsMatch.Success ? gsMatch.Value : code;
            if (!codes.Contains(checkCode))
                rowsToDelete.Add(r);
        }

        for (int i = rowsToDelete.Count - 1; i >= 0; i--)
            ws.Row(rowsToDelete[i]).Delete();

        wb.SaveAs(dest);
    }

    private void CreateFilteredCsv(string source, string dest, HashSet<string> codes)
    {
        string[] lines;
        Encoding enc;
        try { lines = File.ReadAllLines(source, Encoding.UTF8); enc = Encoding.UTF8; }
        catch { enc = Encoding.GetEncoding(949); lines = File.ReadAllLines(source, enc); }

        if (lines.Length < 2) return;

        var headers = ParseCsvLine(lines[0]);
        int codeIdx = -1, nameIdx = -1;
        for (int i = 0; i < headers.Length; i++)
        {
            var h = headers[i].Trim();
            if (codeIdx < 0 && CodeColumns.Any(c => c.Equals(h, StringComparison.OrdinalIgnoreCase)))
                codeIdx = i;
            if (nameIdx < 0 && NameColumns.Any(c => c.Equals(h, StringComparison.OrdinalIgnoreCase)))
                nameIdx = i;
        }
        if (codeIdx < 0) return;

        var output = new List<string> { lines[0] };
        for (int r = 1; r < lines.Length; r++)
        {
            var cols = ParseCsvLine(lines[r]);
            if (codeIdx >= cols.Length) continue;
            var code = cols[codeIdx].Trim();
            var name = (nameIdx >= 0 && nameIdx < cols.Length) ? cols[nameIdx].Trim() : "";

            var gsMatch = Regex.Match(code, @"(GS\d{7})", RegexOptions.IgnoreCase);
            if (!gsMatch.Success)
                gsMatch = Regex.Match(name, @"(GS\d{7})", RegexOptions.IgnoreCase);

            var checkCode = gsMatch.Success ? gsMatch.Value : code;
            if (codes.Contains(checkCode))
                output.Add(lines[r]);
        }
        File.WriteAllLines(dest, output, enc);
    }

    #endregion

    #region ═══ STEP 1: 파이프라인 실행 ═══

    private ListingImageSettings BuildListingSettings()
    {
        var s = new ListingImageSettings(
            MakeListing: MakeListingCheck.IsChecked == true,
            ListingSize: ParseInt(SettingsListingSize, 1200),
            LogoPath: SettingsLogoPath.Text.Trim(),
            LogoRatio: ParseInt(SettingsLogoRatio, 14),
            LogoOpacity: ParseInt(SettingsLogoOpacity, 65),
            LogoPosition: (SettingsLogoPos.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "tr",
            UseAutoContrast: SettingsAutoContrast.IsChecked == true,
            UseSharpen: SettingsSharpen.IsChecked == true,
            UseSmallRotate: SettingsSmallRotate.IsChecked == true,
            RotateZoom: ParseDouble(SettingsRotateZoom, 1.04),
            JpegQualityMin: ParseInt(SettingsJpegMin, 88),
            JpegQualityMax: ParseInt(SettingsJpegMax, 92),
            FlipLeftRight: SettingsFlipLR.IsChecked == true,
            LogoPathB: SettingsLogoPathB.Text.Trim(),
            ImgTag: SettingsImgTag.Text.Trim(),
            ImgTagB: SettingsImgTagB.Text.Trim(),
            ANameMin: ParseInt(SettingsANameMin, 80),
            ANameMax: ParseInt(SettingsANameMax, 100),
            BNameMin: ParseInt(SettingsBNameMin, 63),
            BNameMax: ParseInt(SettingsBNameMax, 98),
            ATagCount: ParseInt(SettingsATagCount, 20),
            BTagCount: ParseInt(SettingsBTagCount, 14),
            KeywordVersion: GetSelectedKeywordVersion(),
            BMarketTokenPath: _bMarketTokenPath
        );
        SaveAppSettings(s);
        return s;
    }

    private void SaveAppSettings(ListingImageSettings s)
    {
        try
        {
            var json = JsonSerializer.Serialize(s, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(_settingsPath, json);
        }
        catch { }
    }

    private void LoadAppSettings()
    {
        try
        {
            if (!File.Exists(_settingsPath)) return;
            var json = File.ReadAllText(_settingsPath);
            var s = JsonSerializer.Deserialize<ListingImageSettings>(json);
            if (s is null) return;

            SettingsLogoPath.Text = s.LogoPath;
            SettingsLogoPathB.Text = s.LogoPathB;
            SettingsImgTag.Text = s.ImgTag;
            SettingsImgTagB.Text = s.ImgTagB;
            SettingsLogoRatio.Text = s.LogoRatio.ToString();
            SettingsLogoOpacity.Text = s.LogoOpacity.ToString();
            SettingsListingSize.Text = s.ListingSize.ToString();
            SettingsJpegMin.Text = s.JpegQualityMin.ToString();
            SettingsJpegMax.Text = s.JpegQualityMax.ToString();
            SettingsRotateZoom.Text = s.RotateZoom.ToString(CultureInfo.InvariantCulture);
            SettingsAutoContrast.IsChecked = s.UseAutoContrast;
            SettingsSharpen.IsChecked = s.UseSharpen;
            SettingsSmallRotate.IsChecked = s.UseSmallRotate;
            SettingsFlipLR.IsChecked = s.FlipLeftRight;
            MakeListingCheck.IsChecked = s.MakeListing;
            SettingsANameMin.Text = s.ANameMin.ToString();
            SettingsANameMax.Text = s.ANameMax.ToString();
            SettingsBNameMin.Text = s.BNameMin.ToString();
            SettingsBNameMax.Text = s.BNameMax.ToString();
            SettingsATagCount.Text = s.ATagCount.ToString();
            SettingsBTagCount.Text = s.BTagCount.ToString();
            SetKeywordVersionSelection(string.IsNullOrWhiteSpace(s.KeywordVersion) ? "2.0" : s.KeywordVersion);

            // 로고 위치 콤보박스
            SetComboSelection(SettingsLogoPos, s.LogoPosition);

            // B마켓 토큰 경로
            if (!string.IsNullOrWhiteSpace(s.BMarketTokenPath))
            {
                _bMarketTokenPath = s.BMarketTokenPath;
                SettingsBTokenPath.Text = s.BMarketTokenPath;
                LoadTokenInfoB();
            }
            else
            {
                LoadTokenInfoB(); // 기본 경로로 시도
            }
        }
        catch { }
    }

    private string GetSelectedKeywordVersion()
    {
        var selected = GetComboSelectedText(TestKeywordVersionCombo)
            ?? GetComboSelectedText(KeywordVersionCombo);

        return selected switch
        {
            "1.0" => "1.0",
            "3.0" => "3.0",
            _ => "3.0",
        };
    }

    private void SetKeywordVersionSelection(string? version)
    {
        var trimmed = version?.Trim();
        var normalized = string.Equals(trimmed, "1.0", StringComparison.OrdinalIgnoreCase) ? "1.0"
            : string.Equals(trimmed, "3.0", StringComparison.OrdinalIgnoreCase) ? "3.0"
            : "3.0";
        _syncingKeywordVersion = true;
        try
        {
            SetComboSelection(KeywordVersionCombo, normalized);
            SetComboSelection(TestKeywordVersionCombo, normalized);
        }
        finally
        {
            _syncingKeywordVersion = false;
        }
    }

    private void KeywordVersionCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_syncingKeywordVersion) return;

        var normalized = GetComboSelectedText(sender as ComboBox) switch
        {
            "1.0" => "1.0",
            "3.0" => "3.0",
            _ => "3.0",
        };

        SetKeywordVersionSelection(normalized);

        if (!string.IsNullOrWhiteSpace(_testOutputRoot) && Directory.Exists(_testOutputRoot))
            RefreshTestCodexCommands(_testOutputRoot);
    }

    private static string? GetComboSelectedText(ComboBox? comboBox)
        => (comboBox?.SelectedItem as ComboBoxItem)?.Content?.ToString()?.Trim();

    private static void SetComboSelection(ComboBox comboBox, string? value)
    {
        if (comboBox is null || string.IsNullOrWhiteSpace(value)) return;
        foreach (ComboBoxItem item in comboBox.Items)
        {
            if (string.Equals(item.Content?.ToString(), value, StringComparison.OrdinalIgnoreCase))
            {
                comboBox.SelectedItem = item;
                return;
            }
        }
    }

    private async void RunPipeline_Click(object sender, RoutedEventArgs e)
    {
        if (!ValidateSource()) return;
        if (_products.Count > 0 && !_products.Any(p => p.IsSelected))
        {
            MessageBox.Show("처리할 상품을 선택하세요.", "알림", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        SetRunning(true);
        _cts = new CancellationTokenSource();

        try
        {
            var inputFile = CreateFilteredFile();
            if (inputFile == null) { SetRunning(false); return; }

            var settings = BuildListingSettings();
            var bridge = new PythonPipelineBridgeService(_v3Root, _pythonRoot);
            var progress = new Progress<string>(msg => Log(msg));
            var selectedModel = (ModelCombo.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "";
            var selectedKeywordVersion = GetSelectedKeywordVersion();

            if (settings.MakeListing)
            {
                // ── 2-Phase 실행: 이미지 먼저 → 피커 → 분석 병렬 ──
                Log($"Phase 1: 이미지 다운로드 + 가공 시작... (모델: {selectedModel}, 키워드 버전: {selectedKeywordVersion})");
                StatusText.Text = "Phase 1: 이미지 처리 중...";
                ProgressBar.IsIndeterminate = true;

                var phase1 = await bridge.RunPipelineAsync(
                    inputFile, settings, progress, _cts.Token, phase: "images", model: selectedModel, keywordVersion: selectedKeywordVersion);

                _lastOutputRoot = phase1.OutputRoot;
                Log($"Phase 1 완료 — 이미지 폴더: {phase1.OutputRoot}");

                // Phase 2: 분석 백그라운드 시작
                Log("Phase 2: OCR + Vision + 키워드 생성 (백그라운드)...");
                StatusText.Text = "이미지 선택 중... (백그라운드에서 키워드 생성 중)";
                var phase2Progress = new Progress<string>(msg => Log($"[Phase2] {msg}"));
                var phase2Task = bridge.RunPipelineAsync(
                    inputFile, settings, phase2Progress, _cts.Token,
                    phase: "analysis", exportRoot: phase1.OutputRoot, model: selectedModel, keywordVersion: selectedKeywordVersion);

                // 이미지 선택 탭으로 전환 + 이미지 로드
                LoadListingImagesFromRoot(phase1.OutputRoot);

                // Phase 2 완료 대기
                var phase2Result = await phase2Task;
                OnPipelineComplete(phase2Result);
            }
            else
            {
                // ── 기존 단일 실행 (이미지 없이 키워드만) ──
                Log($"전체 파이프라인 실행 시작... (모델: {selectedModel}, 키워드 버전: {selectedKeywordVersion})");
                StatusText.Text = "실행 중...";
                ProgressBar.IsIndeterminate = true;

                var result = await bridge.RunPipelineAsync(inputFile, settings, progress, _cts.Token, model: selectedModel, keywordVersion: selectedKeywordVersion);
                OnPipelineComplete(result);
            }
        }
        catch (OperationCanceledException) { Log("작업 취소됨"); StatusText.Text = "취소됨"; }
        catch (Exception ex) { HandlePipelineError(ex); }
        finally { SetRunning(false); ProgressBar.IsIndeterminate = false; }
    }

    private async void RunKeywordOnly_Click(object sender, RoutedEventArgs e)
    {
        if (!ValidateSource()) return;

        SetRunning(true);
        _cts = new CancellationTokenSource();

        try
        {
            var inputFile = CreateFilteredFile() ?? _sourcePath!;
            var settings = BuildListingSettings() with { MakeListing = false };
            var bridge = new PythonPipelineBridgeService(_v3Root, _pythonRoot);
            var progress = new Progress<string>(msg => Log(msg));
            var selectedModel = (ModelCombo.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "";
            var selectedKeywordVersion = GetSelectedKeywordVersion();

            Log($"키워드만 생성 시작... (모델: {selectedModel}, 키워드 버전: {selectedKeywordVersion})");
            StatusText.Text = "키워드 생성 중...";
            ProgressBar.IsIndeterminate = true;

            var result = await bridge.RunPipelineAsync(inputFile, settings, progress, _cts.Token, model: selectedModel, keywordVersion: selectedKeywordVersion);
            OnPipelineComplete(result);
        }
        catch (OperationCanceledException) { Log("작업 취소됨"); StatusText.Text = "취소됨"; }
        catch (Exception ex) { HandlePipelineError(ex); }
        finally { SetRunning(false); ProgressBar.IsIndeterminate = false; }
    }

    private async void RunListingOnly_Click(object sender, RoutedEventArgs e)
    {
        if (!ValidateSource()) return;

        SetRunning(true);
        _cts = new CancellationTokenSource();

        try
        {
            var inputFile = CreateFilteredFile() ?? _sourcePath!;
            var settings = BuildListingSettings();
            var bridge = new PythonPipelineBridgeService(_v3Root, _pythonRoot);
            var progress = new Progress<string>(msg => Log(msg));
            var selectedModel = (ModelCombo.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "";
            var selectedKeywordVersion = GetSelectedKeywordVersion();

            Log($"대표이미지만 생성 시작... (모델: {selectedModel}, 키워드 버전: {selectedKeywordVersion})");
            StatusText.Text = "대표이미지 생성 중...";
            ProgressBar.IsIndeterminate = true;

            var result = await bridge.RunPipelineAsync(inputFile, settings, progress, _cts.Token, model: selectedModel, keywordVersion: selectedKeywordVersion);
            OnPipelineComplete(result);
        }
        catch (OperationCanceledException) { Log("작업 취소됨"); StatusText.Text = "취소됨"; }
        catch (Exception ex) { HandlePipelineError(ex); }
        finally { SetRunning(false); ProgressBar.IsIndeterminate = false; }
    }

    private void OnPipelineComplete(PythonPipelineBridgeResult result)
    {
        _lastOutputRoot = result.OutputRoot;
        _lastOutputFile = result.OutputFile;

        OpenUploadExcelButton.IsEnabled = true;
        OpenOutputFolderButton.IsEnabled = true;
        Cafe24UploadButton.IsEnabled = true;
        Cafe24CreateButton.IsEnabled = true;

        Log($"완료: {result.OutputFile}");
        StatusText.Text = "완료 — Cafe24 업로드 가능";
        OutputFileText.Text = result.OutputFile;

        var uploadFile = FindLatestFile(_lastOutputRoot, "업로드용_*.xlsx");
        if (uploadFile != null)
        {
            Clipboard.SetText(uploadFile);
            Log($"업로드용 엑셀 클립보드 복사: {Path.GetFileName(uploadFile)}");
        }

        // 결과 폴더 자동 열기
        if (!string.IsNullOrEmpty(_lastOutputRoot) && Directory.Exists(_lastOutputRoot))
            Process.Start(new ProcessStartInfo("explorer.exe", _lastOutputRoot));

        // 완료 알림 + 앱 포커스
        Activate();
        Topmost = true;
        Topmost = false;
        System.Media.SystemSounds.Asterisk.Play();
        MessageBox.Show(
            $"파이프라인 완료!\n\n" +
            $"파일: {Path.GetFileName(result.OutputFile)}\n" +
            $"폴더: {_lastOutputRoot}",
            "작업 완료", MessageBoxButton.OK, MessageBoxImage.Information);

        // 실행 이력 저장
        var selectedModel = (ModelCombo.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "";
        var job = new JobRecord
        {
            SourceFile = _sourcePath ?? "",
            OutputRoot = result.OutputRoot,
            OutputFile = result.OutputFile,
            ProductCount = _products.Count(p => p.IsSelected),
            SelectedCodes = _products.Where(p => p.IsSelected).Select(p => p.Code).ToList(),
            Model = selectedModel,
            MakeListing = MakeListingCheck.IsChecked == true,
            Status = "완료",
        };
        _jobHistory?.Add(job);
        RefreshHistoryGrid();
        ApplyHistoryToProducts();
    }

    private void HandlePipelineError(Exception ex)
    {
        Log($"오류: {ex.Message}");
        StatusText.Text = "오류 발생";
        MessageBox.Show(ex.Message, "파이프라인 오류", MessageBoxButton.OK, MessageBoxImage.Error);
    }

    #endregion

    #region ═══ 테스트실행 (OCR Only + LLM 수동) ═══

    private string? _testOutputRoot;
    private string? _testLlmResultFile;
    private List<string> _testLlmResultFiles = new();
    private string? _testSkipOcrFolder;

    private int GetTestChunkSize()
    {
        var selected = (TestChunkSizeCombo.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "10개";
        if (selected == "분할안함") return 0;
        return int.TryParse(selected.Replace("개", ""), out var n) ? n : 10;
    }

    private void TestSkipOcrCheck_Changed(object sender, RoutedEventArgs e)
    {
        var isChecked = TestSkipOcrCheck.IsChecked == true;
        TestSkipOcrFolderPanel.Visibility = isChecked ? Visibility.Visible : Visibility.Collapsed;
        TestRunOcrOnlyButton.Content = isChecked ? "청크 재생성 (OCR 제외)" : "1차 가공 실행";
        TestRunOcrOnlyButton.Background = isChecked
            ? new System.Windows.Media.SolidColorBrush(
                (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#e67e22"))
            : new System.Windows.Media.SolidColorBrush(
                (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#6c5ce7"));
    }

    private void TestSkipOcrSelectFolder_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "업로드용 엑셀|업로드용_*.xlsx|모든 파일|*.*",
            Title = "기존 업로드용 엑셀 선택 (OCR결과 포함된 파일)",
            InitialDirectory = @"C:\code\exports",
        };

        if (dlg.ShowDialog() == true)
        {
            _testSkipOcrFolder = Path.GetDirectoryName(dlg.FileName)!;
            TestSkipOcrFolderText.Text = _testSkipOcrFolder;
            Log($"OCR 재사용 폴더: {_testSkipOcrFolder}");
            Log($"업로드용 엑셀: {Path.GetFileName(dlg.FileName)}");
        }
    }

    private async void TestRunOcrOnly_Click(object sender, RoutedEventArgs e)
    {
        // OCR 제외 모드
        if (TestSkipOcrCheck.IsChecked == true)
        {
            await TestRunSkipOcr_Execute();
            return;
        }

        if (!ValidateSource()) return;
        if (_products.Count > 0 && !_products.Any(p => p.IsSelected))
        {
            MessageBox.Show("처리할 상품을 선택하세요.", "알림", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        SetRunning(true);
        _cts = new CancellationTokenSource();
        var chunkSize = GetTestChunkSize();

        try
        {
            var inputFile = CreateFilteredFile();
            if (inputFile == null) { SetRunning(false); return; }

            var settings = BuildListingSettings();
            var bridge = new PythonPipelineBridgeService(_v3Root, _pythonRoot);
            var progress = new Progress<string>(msg => Log(msg));
            var selectedModel = (ModelCombo.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "";
            var selectedKeywordVersion = GetSelectedKeywordVersion();

            Log($"테스트실행: 1차 가공 시작 (OCR + 이미지, LLM 스킵, 분할: {(chunkSize > 0 ? $"{chunkSize}개씩" : "안함")}, 키워드 버전: {selectedKeywordVersion})...");
            StatusText.Text = "1차 가공 중 (LLM 스킵)...";
            ProgressBar.IsIndeterminate = true;

            if (settings.MakeListing)
            {
                // Phase 1: 이미지 처리
                Log("Phase 1: 이미지 다운로드 + 가공...");
                var phase1 = await bridge.RunPipelineAsync(
                    inputFile, settings, progress, _cts.Token, phase: "images", model: selectedModel, keywordVersion: selectedKeywordVersion);

                _testOutputRoot = phase1.OutputRoot;
                Log($"Phase 1 완료 — 이미지 폴더: {phase1.OutputRoot}");

                // 이미지 선택 탭으로 전환
                LoadListingImagesFromRoot(phase1.OutputRoot);

                // Phase 2: OCR only (LLM 스킵)
                Log("Phase 2: OCR only 실행 (키워드 생성 스킵)...");
                StatusText.Text = "OCR 처리 중 (LLM 스킵)...";
                var phase2Progress = new Progress<string>(msg => Log($"[OCR] {msg}"));
                var phase2Result = await bridge.RunPipelineAsync(
                    inputFile, settings, phase2Progress, _cts.Token,
                    phase: "ocr_only", exportRoot: phase1.OutputRoot, model: selectedModel,
                    chunkSize: chunkSize, keywordVersion: selectedKeywordVersion);

                OnTestOcrComplete(phase2Result);
            }
            else
            {
                var result = await bridge.RunPipelineAsync(
                    inputFile, settings, progress, _cts.Token, phase: "ocr_only", model: selectedModel,
                    chunkSize: chunkSize, keywordVersion: selectedKeywordVersion);
                OnTestOcrComplete(result);
            }
        }
        catch (OperationCanceledException) { Log("작업 취소됨"); StatusText.Text = "취소됨"; }
        catch (Exception ex) { HandlePipelineError(ex); }
        finally { SetRunning(false); ProgressBar.IsIndeterminate = false; }
    }

    private async Task TestRunSkipOcr_Execute()
    {
        if (string.IsNullOrEmpty(_testSkipOcrFolder) || !Directory.Exists(_testSkipOcrFolder))
        {
            MessageBox.Show("기존 폴더를 먼저 선택하세요.", "알림", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // 업로드용 엑셀 찾기
        var uploadFiles = Directory.GetFiles(_testSkipOcrFolder, "업로드용_*.xlsx");
        if (uploadFiles.Length == 0)
        {
            MessageBox.Show("선택한 폴더에 업로드용 엑셀이 없습니다.", "알림", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        Array.Sort(uploadFiles);
        var uploadPath = uploadFiles[^1]; // 최신 파일
        var chunkSize = GetTestChunkSize();

        SetRunning(true);
        ProgressBar.IsIndeterminate = true;
        StatusText.Text = "청크 재생성 중 (OCR 제외)...";
        Log($"OCR 제외 모드: 기존 업로드용 엑셀로 skill.md + 청크만 재생성");
        Log($"엑셀: {Path.GetFileName(uploadPath)}");

        try
        {
            var exportRoot = _testSkipOcrFolder;

            // Python bridge로 phase=ocr_only 실행 (이미 OCR 결과가 엑셀에 포함되어 있으므로 OCR은 스킵되고 skill.md + 청크만 생성됨)
            var bridgeService = new PythonPipelineBridgeService(_v3Root, _pythonRoot);
            var progress = new Progress<string>(msg => Log(msg));
            var selectedModel = (ModelCombo.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "";
            var selectedKeywordVersion = GetSelectedKeywordVersion();

            var result = await bridgeService.RunPipelineAsync(
                uploadPath, BuildListingSettings(), progress, CancellationToken.None,
                phase: "ocr_only", exportRoot: exportRoot, model: selectedModel,
                chunkSize: chunkSize, keywordVersion: selectedKeywordVersion);

            _testOutputRoot = exportRoot;
            OnTestOcrComplete(result);
        }
        catch (Exception ex)
        {
            Log($"청크 재생성 오류: {ex.Message}");
            MessageBox.Show(ex.Message, "오류", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            SetRunning(false);
            ProgressBar.IsIndeterminate = false;
        }
    }

    private List<string> _codexCommands = new();
    private List<string> _codexCommandsExt = new();

    private static string GetKeywordVersionSuffix(string version)
        => version switch
        {
            "1.0" => "v1_0",
            "3.0" => "v3_0",
            _ => "v2_0",
        };

    private static string GetKeywordVersionLabel(string version)
        => version switch
        {
            "1.0" => "v1.0 확장형",
            "3.0" => "v3.0 타겟형",
            _ => "v2.0 근거우선",
        };

    private static string GetChunksRoot(string outputRoot)
        => Path.Combine(outputRoot, "llm_chunks");

    private static string GetActiveChunksMarkerPath(string outputRoot, string version)
        => Path.Combine(GetChunksRoot(outputRoot), $"_active_{GetKeywordVersionSuffix(version)}.txt");

    private string GetActiveChunksDir(string outputRoot, string? version = null)
    {
        var selectedVersion = version ?? GetSelectedKeywordVersion();
        var versionSuffix = GetKeywordVersionSuffix(selectedVersion);
        var chunksRoot = GetChunksRoot(outputRoot);
        if (!Directory.Exists(chunksRoot))
            return chunksRoot;

        var markerPath = GetActiveChunksMarkerPath(outputRoot, selectedVersion);
        try
        {
            if (File.Exists(markerPath))
            {
                var markedDir = File.ReadAllText(markerPath).Trim();
                if (!string.IsNullOrWhiteSpace(markedDir) && Directory.Exists(markedDir))
                    return markedDir;
            }
        }
        catch { }

        if (Directory.GetFiles(chunksRoot, "chunk_*.xlsx").Length > 0)
            return chunksRoot;

        var versionedSessions = Directory.GetDirectories(chunksRoot, $"session_*_{versionSuffix}_*")
            .Where(dir => Directory.GetFiles(dir, "chunk_*.xlsx").Length > 0)
            .OrderByDescending(Directory.GetCreationTimeUtc)
            .ToArray();
        if (versionedSessions.Length > 0)
            return versionedSessions[0];

        var anySessions = Directory.GetDirectories(chunksRoot, "session_*")
            .Where(dir => Directory.GetFiles(dir, "chunk_*.xlsx").Length > 0)
            .OrderByDescending(Directory.GetCreationTimeUtc)
            .ToArray();
        if (anySessions.Length > 0)
            return anySessions[0];

        return chunksRoot;
    }

    private IEnumerable<string> GetPreferredLlmDirs(string outputRoot, string version)
    {
        var versionSuffix = GetKeywordVersionSuffix(version);
        var chunksRoot = GetChunksRoot(outputRoot);
        var activeChunksDir = GetActiveChunksDir(outputRoot, version);

        return new[]
        {
            Path.Combine(activeChunksDir, $"llm_result_{versionSuffix}"),
            Path.Combine(outputRoot, $"llm_result_{versionSuffix}"),
            Path.Combine(activeChunksDir, $"llm_result_ext_{versionSuffix}"),
            Path.Combine(outputRoot, $"llm_result_ext_{versionSuffix}"),
            Path.Combine(chunksRoot, $"llm_result_{versionSuffix}"),
            Path.Combine(chunksRoot, $"llm_result_ext_{versionSuffix}"),
            Path.Combine(activeChunksDir, "llm_result"),
            Path.Combine(outputRoot, "llm_result"),
            Path.Combine(activeChunksDir, "llm_result_ext"),
            Path.Combine(outputRoot, "llm_result_ext"),
            Path.Combine(chunksRoot, "llm_result"),
            Path.Combine(chunksRoot, "llm_result_ext"),
        }.Distinct().Where(Directory.Exists);
    }

    private static string GetKeywordVersionCommandGuide(string version, bool extended)
    {
        if (version == "1.0")
        {
            return extended
                ? "키워드 버전 1.0 확장형으로 처리해. 핵심상품명은 맨 앞에 두고, 온토픽 범위에서만 실무 유사어와 사용처를 조금 더 활용하되 다른 상품군, 오타 확장, 과장 문구는 금지해."
                : "키워드 버전 1.0 확장형으로 처리해. 핵심상품명은 맨 앞에 두고, 온토픽 범위에서 검색 커버리지를 조금 더 확보하되 다른 상품군, 오타 확장, 과장 문구는 금지해.";
        }

        if (version == "3.0")
        {
            return extended
                ? "키워드 버전 3.0 타겟형으로 처리해. 근거는 원본 상품명과 OCR결과 시트만 사용하고, 외부 검색/연관어/자동완성은 절대 금지해. 단위 없는 순수 숫자 토큰(801, 2024 같은 제품코드·연도·바코드)은 제외하되 단위 붙은 규격(35mm, 2M, 500ml)은 유지해. A마켓은 기능·규격·재질 중심으로 확장하고, B마켓은 A의 부분집합이 아니라 핵심상품명·규격만 공유하며 용도·사용처·대상 축으로 독립 패키징해. A 제목 뒷부분과 B 제목 뒷부분 토큰이 50% 이상 겹치지 않게 해."
                : "키워드 버전 3.0 타겟형으로 처리해. 근거는 상품명과 OCR뿐. 외부 검색/연관어 금지. OCR 숫자 필터(단위 붙은 규격 유지, 순수 숫자 제외) 적용. A=기능/규격 확장, B=용도/사용처 독립 패키징. A와 B 제목 뒷부분은 50% 이상 겹치면 안 됨.";
        }

        return extended
            ? "키워드 버전 2.0 근거 우선 규칙으로 처리해. evidence-first로 상품명과 OCR 근거 안에서만 확장하고, 짧아도 무관 확장과 오타 확장은 금지해."
            : "키워드 버전 2.0 근거 우선 규칙으로 처리해. evidence-first로 상품명과 OCR 근거 안에서만 조립하고, 짧아도 무관 확장과 오타 확장은 금지해.";
    }

    private static string BuildTestCodexCommand(string workingDir, string instruction)
        => $"cd \"{workingDir}\"; codex --full-auto \"{instruction}\"";

    private static string GetCategoryMatchingCommandGuide(string versionSuffix, bool extended, string? inputFileName = null)
    {
        var resultDir = extended ? $"llm_result_ext_{versionSuffix}" : $"llm_result_{versionSuffix}";
        var outputStem = extended ? $"category_match_ext_{versionSuffix}" : $"category_match_{versionSuffix}";
        var outputFile = string.IsNullOrWhiteSpace(inputFileName)
            ? $"{resultDir}/{outputStem}.xlsx"
            : $"{resultDir}/{Path.GetFileNameWithoutExtension(inputFileName)}_{outputStem}.xlsx";

        return $"키워드 저장 후 상품코드/상품명 기준 마켓별 카테고리 매칭 파일도 `{outputFile}`로 추가 생성해. " +
               "카테고리 기준표는 `category_reference/` 또는 같은 폴더의 `*_categories.csv`, `lotteon_standard_categories.csv`, `lotteon_display_categories.csv`, `esm_auction_gmarket_category_matching.csv`를 우선 사용해. " +
               "포함 열: 상품코드, 상품명, 네이버카테고리코드/경로, 쿠팡카테고리코드/경로, 11번가카테고리코드/경로, " +
               "롯데ON표준카테고리코드/경로, 롯데ON전시카테고리코드/경로, 옥션카테고리코드/경로, G마켓카테고리경로, ESM카테고리경로, 확신도, 검수필요, 매칭근거, " +
               "마켓플러스검증상태, 마켓플러스차단마켓, 마켓플러스검증메모, G마켓옵션위험, 롯데ON옵션위험, 옵션검수필요. " +
               "마켓플러스검증상태는 PASS/WARN/BLOCK으로 기록하고, 상품명 100자 초과, 옵션명/옵션값 25자 또는 50바이트 초과 위험, G마켓/옥션 권장 옵션명 불일치 위험, 롯데ON 표준/전시카테고리 누락 또는 옵션명 매칭 필요 여부를 표시해. " +
               "ESM 매칭표는 사이트=G마켓 행의 G/A 카테고리명을 G마켓카테고리경로로 기록해. " +
               "외부 검색 없이 상품명/OCR/생성 키워드와 제공된 카테고리표만 근거로 해.";
    }

    private void RefreshTestCodexCommands(string outputRoot)
    {
        var selectedKeywordVersion = GetSelectedKeywordVersion();
        var versionSuffix = GetKeywordVersionSuffix(selectedKeywordVersion);
        var versionLabel = GetKeywordVersionLabel(selectedKeywordVersion);

        var skillMd = Path.Combine(outputRoot, "keyword_skill.md");
        var skillMdExt = Path.Combine(outputRoot, "keyword_skill_extended.md");
        if (File.Exists(skillMd) && File.Exists(skillMdExt))
            TestSkillMdPathText.Text = $"keyword_skill.md + extended 생성됨 · 현재 명령 버전 {selectedKeywordVersion}";
        else if (File.Exists(skillMd))
            TestSkillMdPathText.Text = $"keyword_skill.md 생성됨 · 현재 명령 버전 {selectedKeywordVersion}";
        else
            TestSkillMdPathText.Text = $"현재 명령 버전 {selectedKeywordVersion}";

        _codexCommands.Clear();
        _codexCommandsExt.Clear();

        var chunksDir = GetActiveChunksDir(outputRoot, selectedKeywordVersion);
        if (Directory.Exists(chunksDir))
        {
            var chunkFiles = Directory.GetFiles(chunksDir, "chunk_*.xlsx");
            if (chunkFiles.Length > 0)
            {
                Array.Sort(chunkFiles);
                Directory.CreateDirectory(Path.Combine(chunksDir, $"llm_result_{versionSuffix}"));
                Directory.CreateDirectory(Path.Combine(chunksDir, $"llm_result_ext_{versionSuffix}"));

                TestCodexCmdTitle.Text = $"Codex 병렬 실행 ({chunkFiles.Length}개 세션 × 2세트, {versionLabel})";
                foreach (var cf in chunkFiles)
                {
                    var fileName = Path.GetFileName(cf);
                    var outputFileName = $"{Path.GetFileNameWithoutExtension(fileName)}_llm_{versionSuffix}.xlsx";
                    var cmd = BuildTestCodexCommand(
                        chunksDir,
                        $"keyword_skill.md 지시서에 따라 {fileName} 파일의 키워드를 채워서 llm_result_{versionSuffix}/{outputFileName} 로 저장해. {GetKeywordVersionCommandGuide(selectedKeywordVersion, extended: false)} {GetCategoryMatchingCommandGuide(versionSuffix, extended: false, inputFileName: fileName)}");
                    _codexCommands.Add(cmd);

                    var cmdExt = BuildTestCodexCommand(
                        chunksDir,
                        $"keyword_skill_extended.md 지시서에 따라 {fileName} 파일의 키워드를 채워서 llm_result_ext_{versionSuffix}/{outputFileName} 로 저장해. {GetKeywordVersionCommandGuide(selectedKeywordVersion, extended: true)} {GetCategoryMatchingCommandGuide(versionSuffix, extended: true, inputFileName: fileName)}");
                    _codexCommandsExt.Add(cmdExt);
                }

                Log($"분할 엑셀 {chunkFiles.Length}개 → 기본/확장 2세트 명령어 생성 ({versionLabel})");
                BuildCodexCommandCards();
                TestCodexCmdPanel.Visibility = Visibility.Visible;
                StatusText.Text = $"1차 가공 완료 — LLM 키워드 처리 대기 ({versionLabel})";
                return;
            }
        }

        Directory.CreateDirectory(Path.Combine(outputRoot, $"llm_result_{versionSuffix}"));
        Directory.CreateDirectory(Path.Combine(outputRoot, $"llm_result_ext_{versionSuffix}"));

        var uploadFile = !string.IsNullOrWhiteSpace(_lastOutputFile) &&
                         string.Equals(Path.GetDirectoryName(_lastOutputFile), outputRoot, StringComparison.OrdinalIgnoreCase)
            ? _lastOutputFile
            : FindLatestFile(outputRoot, "업로드용_*.xlsx");
        var uploadName = !string.IsNullOrWhiteSpace(uploadFile) ? Path.GetFileName(uploadFile) : "업로드용 엑셀";

        _codexCommands.Add(BuildTestCodexCommand(
            outputRoot,
            $"keyword_skill.md 지시서에 따라 {uploadName} 파일의 키워드를 채워서 llm_result_{versionSuffix}/ 아래에 저장해. 파일명은 입력 파일명 기준으로 `_llm_{versionSuffix}.xlsx` 형식으로 저장해. {GetKeywordVersionCommandGuide(selectedKeywordVersion, extended: false)} {GetCategoryMatchingCommandGuide(versionSuffix, extended: false, inputFileName: uploadName)}"));
        _codexCommandsExt.Add(BuildTestCodexCommand(
            outputRoot,
            $"keyword_skill_extended.md 지시서에 따라 {uploadName} 파일의 키워드를 채워서 llm_result_ext_{versionSuffix}/ 아래에 저장해. 파일명은 입력 파일명 기준으로 `_llm_{versionSuffix}.xlsx` 형식으로 저장해. {GetKeywordVersionCommandGuide(selectedKeywordVersion, extended: true)} {GetCategoryMatchingCommandGuide(versionSuffix, extended: true, inputFileName: uploadName)}"));

        TestCodexCmdTitle.Text = $"Codex 실행 (기본 + 확장, {versionLabel})";
        Log($"keyword_skill.md + extended → Codex에서 실행 ({versionLabel})");
        BuildCodexCommandCards();
        TestCodexCmdPanel.Visibility = Visibility.Visible;
        StatusText.Text = $"1차 가공 완료 — LLM 키워드 처리 대기 ({versionLabel})";
    }

    private void OnTestOcrComplete(PythonPipelineBridgeResult result)
    {
        _testOutputRoot = result.OutputRoot;
        _lastOutputFile = result.OutputFile;
        TestOutputPathText.Text = $"결과 폴더: {result.OutputRoot}";
        TestOpenOutputButton.IsEnabled = true;

        // OCR 제외 모드용 폴더 자동 설정 (재실행 시 폴더 재선택 불필요)
        _testSkipOcrFolder = result.OutputRoot;
        TestSkipOcrFolderText.Text = result.OutputRoot;

        // 이미지 선택 탭 자동 로드 (listing_images 폴더가 있으면)
        var listingDir = Path.Combine(result.OutputRoot, "listing_images");
        if (Directory.Exists(listingDir) && Directory.GetDirectories(listingDir).Length > 0)
            LoadListingImagesFromRoot(result.OutputRoot);

        RefreshTestCodexCommands(result.OutputRoot);

        var selectedKeywordVersion = GetSelectedKeywordVersion();
        var versionSuffix = GetKeywordVersionSuffix(selectedKeywordVersion);
        var chunksDir = GetActiveChunksDir(result.OutputRoot, selectedKeywordVersion);
        var hasChunks = Directory.Exists(chunksDir) && Directory.GetFiles(chunksDir, "chunk_*.xlsx").Length > 0;
        var llmDir = hasChunks
            ? Path.Combine(chunksDir, $"llm_result_{versionSuffix}")
            : Path.Combine(result.OutputRoot, $"llm_result_{versionSuffix}");
        var llmDirExt = hasChunks
            ? Path.Combine(chunksDir, $"llm_result_ext_{versionSuffix}")
            : Path.Combine(result.OutputRoot, $"llm_result_ext_{versionSuffix}");

        Log($"1차 가공 완료!");
        Log($"기본 결과 → {llmDir}");
        Log($"확장 결과 → {llmDirExt}");

        Activate();
        System.Media.SystemSounds.Asterisk.Play();

        if (!string.IsNullOrEmpty(result.OutputRoot) && Directory.Exists(result.OutputRoot))
            Process.Start(new ProcessStartInfo("explorer.exe", result.OutputRoot));
    }

    private void BuildCodexCommandCards()
    {
        TestCodexCmdList.Items.Clear();

        // ── 기본 키워드셋 섹션 ──
        _AddSectionHeader("기본 키워드셋", "#a29bfe");
        _AddCommandCards(_codexCommands, "기본", "#2d2d44", "#6c5ce7");

        // ── 확장 키워드셋 섹션 (PDF 보고서 규칙) ──
        if (_codexCommandsExt.Count > 0)
        {
            _AddSectionHeader("확장 키워드셋 (SEO 최적화)", "#00b894");
            _AddCommandCards(_codexCommandsExt, "확장", "#2d3d44", "#00b894");
        }
    }

    private void _AddSectionHeader(string title, string colorHex)
    {
        var header = new TextBlock
        {
            Text = $"▸ {title}",
            FontSize = 12,
            FontWeight = FontWeights.Bold,
            Foreground = new System.Windows.Media.SolidColorBrush(
                (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(colorHex)),
            Margin = new Thickness(0, 8, 0, 4)
        };
        TestCodexCmdList.Items.Add(header);
    }

    private void _AddCommandCards(List<string> commands, string label, string bgHex, string accentHex)
    {
        for (int i = 0; i < commands.Count; i++)
        {
            var idx = i;
            var cmd = commands[i];

            var border = new Border
            {
                Background = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(bgHex)),
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(8, 6, 8, 6),
                Margin = new Thickness(0, 0, 0, 6)
            };

            var stack = new StackPanel();

            var header = new TextBlock
            {
                Text = commands.Count > 1 ? $"{label} 세션 {i + 1}" : $"{label} 실행 명령어",
                FontSize = 10,
                FontWeight = FontWeights.Bold,
                Foreground = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(accentHex)),
                Margin = new Thickness(0, 0, 0, 4)
            };

            var cmdText = new TextBox
            {
                Text = cmd,
                FontSize = 10,
                FontFamily = new System.Windows.Media.FontFamily("Consolas"),
                Background = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#1e1e2e")),
                Foreground = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#f8f8f2")),
                BorderThickness = new Thickness(0),
                IsReadOnly = true,
                TextWrapping = TextWrapping.Wrap,
                Padding = new Thickness(6, 4, 6, 4)
            };

            var copyBtn = new Button
            {
                Content = "복사",
                Height = 24,
                FontSize = 10,
                Padding = new Thickness(12, 0, 12, 0),
                Margin = new Thickness(0, 4, 0, 0),
                HorizontalAlignment = HorizontalAlignment.Left,
                Background = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(accentHex)),
                Foreground = System.Windows.Media.Brushes.White
            };
            var capturedLabel = label;
            copyBtn.Click += (s, e) =>
            {
                Clipboard.SetText(cmd);
                Log($"{capturedLabel} 세션 {idx + 1} 명령어 복사됨");
                StatusText.Text = $"{capturedLabel} 세션 {idx + 1} 명령어 복사 완료";
            };

            stack.Children.Add(header);
            stack.Children.Add(cmdText);
            stack.Children.Add(copyBtn);
            border.Child = stack;
            TestCodexCmdList.Items.Add(border);
        }
    }

    private void TestOpenOutput_Click(object sender, RoutedEventArgs e)
    {
        if (!string.IsNullOrEmpty(_testOutputRoot) && Directory.Exists(_testOutputRoot))
            Process.Start(new ProcessStartInfo("explorer.exe", _testOutputRoot));
    }

    private void TestOpenChunksFolder_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_testOutputRoot)) return;
        var chunksDir = GetActiveChunksDir(_testOutputRoot, GetSelectedKeywordVersion());
        if (Directory.Exists(chunksDir))
            Process.Start(new ProcessStartInfo("explorer.exe", chunksDir));
    }

    private void TestCopyAllCodexCmd_Click(object sender, RoutedEventArgs e)
    {
        var sections = new List<string>();
        if (_codexCommands.Count > 0)
        {
            sections.Add("# ── 기본 키워드셋 ──");
            sections.AddRange(_codexCommands.Select((c, i) => $"# 기본 세션 {i + 1}\n{c}"));
        }
        if (_codexCommandsExt.Count > 0)
        {
            sections.Add("\n# ── 확장 키워드셋 (SEO 최적화) ──");
            sections.AddRange(_codexCommandsExt.Select((c, i) => $"# 확장 세션 {i + 1}\n{c}"));
        }
        if (sections.Count > 0)
        {
            Clipboard.SetText(string.Join("\n\n", sections));
            var total = _codexCommands.Count + _codexCommandsExt.Count;
            Log($"전체 Codex 명령어 {total}개 복사됨 (기본 {_codexCommands.Count} + 확장 {_codexCommandsExt.Count})");
            StatusText.Text = $"전체 {total}개 명령어 복사 완료";
        }
    }

    private void TestLoadLlmResult_Click(object sender, RoutedEventArgs e)
    {
        // 현재 선택된 버전 결과 폴더를 우선 탐색
        var startDir = "";
        if (_testOutputRoot != null)
        {
            startDir = GetPreferredLlmDirs(_testOutputRoot, GetSelectedKeywordVersion()).FirstOrDefault() ?? "";
        }

        var dlg = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "Excel|*.xlsx|모든 파일|*.*",
            Title = "LLM 결과 엑셀 선택 (여러 파일 선택 가능)",
            InitialDirectory = Directory.Exists(startDir) ? startDir : "",
            Multiselect = true,
        };

        if (dlg.ShowDialog() == true && dlg.FileNames.Length > 0)
        {
            _testLlmResultFiles = dlg.FileNames.OrderBy(f => f).ToList();
            _testLlmResultFile = _testLlmResultFiles[0]; // 호환용

            // LLM 결과 파일에서 _testOutputRoot 자동 추론
            // 경로 예: .../exports/20260331_xxx/llm_chunks/llm_result/chunk_01_llm.xlsx
            var firstDir = Path.GetDirectoryName(_testLlmResultFiles[0])!;
            if (firstDir.Contains("llm_chunks"))
            {
                // llm_chunks 상위 = export root
                var chunksDir = firstDir;
                while (!string.IsNullOrEmpty(chunksDir) && Path.GetFileName(chunksDir) != "llm_chunks")
                    chunksDir = Path.GetDirectoryName(chunksDir);
                if (!string.IsNullOrEmpty(chunksDir))
                    _testOutputRoot = Path.GetDirectoryName(chunksDir);
            }
            else if (firstDir.Contains("llm_result"))
            {
                _testOutputRoot = Path.GetDirectoryName(firstDir);
            }
            if (string.IsNullOrEmpty(_testOutputRoot))
                _testOutputRoot = firstDir;

            if (_testLlmResultFiles.Count == 1)
                TestLlmResultFileText.Text = $"LLM 결과: {Path.GetFileName(_testLlmResultFiles[0])}";
            else
                TestLlmResultFileText.Text = $"LLM 결과: {_testLlmResultFiles.Count}개 파일 선택됨";

            TestCafe24UploadButton.IsEnabled = true;
            TestCafe24CreateButton.IsEnabled = true;
            TestCoupangUploadButton.IsEnabled = true;
            TestNaverUploadButton.IsEnabled = true;

            foreach (var f in _testLlmResultFiles)
            {
                Log($"LLM 결과 파일: {Path.GetFileName(f)}");
                if (!HasBMarketSheet(f))
                    Log($"  ⚠ B마켓 시트 없음: {Path.GetFileName(f)} — 준비몰 신규등록은 이 파일에서 스킵됩니다.");
            }

            LoadBasicCafe24ProductList(_testLlmResultFiles[0]);
            Log($"신규등록 목록 자동 로드: {_basicCafe24Items.Count}개");
        }
    }

    /// <summary>엑셀 파일에 'B마켓' 시트가 있는지 확인</summary>
    private static bool HasBMarketSheet(string excelPath)
    {
        try
        {
            using var wb = WorkbookFileLoader.OpenReadOnly(excelPath);
            return wb.Worksheets.Any(ws =>
                string.Equals(ws.Name.Trim(), "B마켓", StringComparison.OrdinalIgnoreCase));
        }
        catch { return false; }
    }

    private void TestCafe24Create_Click(object sender, RoutedEventArgs e)
    {
        var files = _testLlmResultFiles.Where(File.Exists).ToList();
        if (files.Count == 0)
        {
            MessageBox.Show("LLM 결과 파일을 먼저 선택하세요.", "알림", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        LoadBasicCafe24ProductList(files[0]);
        Log($"상품 목록 {_basicCafe24Items.Count}개 로드 — 항목 선택 후 '신규등록 실행' 버튼을 클릭하세요.");
    }

    private void TestCafe24Upload_Click(object sender, RoutedEventArgs e)
    {
        var files = _testLlmResultFiles.Where(File.Exists).ToList();
        if (files.Count == 0)
        {
            MessageBox.Show("LLM 결과 파일을 먼저 선택하세요.", "알림", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // 첫 번째 파일로 기존 업로드 로직 호출 (단일 파일용 호환)
        _lastOutputRoot = _testOutputRoot ?? Path.GetDirectoryName(files[0])!;
        _lastOutputFile = files[0];

        // 기존 Cafe24 업로드 로직 재사용
        Cafe24Upload_Click(sender, e);
    }

    private async void TestCoupangUpload_Click(object sender, RoutedEventArgs e)
    {
        // _testOutputRoot에서 업로드용 엑셀 자동 탐색
        var sourcePath = FindUploadExcel();
        if (sourcePath == null)
        {
            MessageBox.Show("업로드용 엑셀 파일을 찾을 수 없습니다.\nCafe24 업로드를 먼저 실행하세요.", "알림",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var dryRun = MarketDryRun.IsChecked == true;
        if (!dryRun)
        {
            var confirm = MessageBox.Show(
                $"쿠팡에 상품을 실제 등록합니다.\n\n파일: {Path.GetFileName(sourcePath)}\n\n계속하시겠습니까?",
                "쿠팡 실제 등록 확인", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (confirm != MessageBoxResult.Yes) return;
        }

        TestCoupangUploadButton.IsEnabled = false;
        _cts = new CancellationTokenSource();

        try
        {
            StatusText.Text = dryRun ? "쿠팡 DRY RUN 중..." : "쿠팡 등록 중...";
            ProgressBar.IsIndeterminate = true;
            Log($"[쿠팡] 업로드 시작: {Path.GetFileName(sourcePath)} (DRY RUN: {dryRun})");

            var options = new CoupangUploadOptions
            {
                RowStart = ParseInt(MarketRowStart, 0),
                RowEnd = ParseInt(MarketRowEnd, 0),
                DryRun = dryRun,
            };

            var service = new CoupangUploadService();
            var progress = new Progress<string>(msg => Log(msg));
            var result = await service.UploadAsync(sourcePath, options, progress, _cts.Token);

            var gridItems = result.Items.Select(item => new MarketResultRow
            {
                Market = "쿠팡",
                Row = item.Row,
                Name = item.Name,
                Status = item.Status,
                Info = !string.IsNullOrEmpty(item.SellerProductId) ? item.SellerProductId : item.Category,
                Error = item.Error,
            }).ToList();
            MarketUploadResultGrid.ItemsSource = gridItems;
            MarketUploadSummary.Text = $"[쿠팡] 성공 {result.SuccessCount} / 실패 {result.FailCount} / 전체 {result.TotalCount}";
            StatusText.Text = $"쿠팡 {(dryRun ? "DRY RUN" : "등록")} 완료";
            Log($"[쿠팡] 완료: 성공 {result.SuccessCount} / 실패 {result.FailCount} / 전체 {result.TotalCount}");
            foreach (var item in result.Items.Where(i => !string.IsNullOrEmpty(i.Error)))
                Log($"  [쿠팡 오류] 행{item.Row} {item.Name} → {item.Error}");
        }
        catch (OperationCanceledException) { Log("[쿠팡] 취소됨"); StatusText.Text = "취소됨"; }
        catch (Exception ex)
        {
            Log($"[쿠팡] 오류: {ex.Message}");
            StatusText.Text = "쿠팡 업로드 오류";
            MessageBox.Show(ex.Message, "쿠팡 업로드 오류", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            TestCoupangUploadButton.IsEnabled = true;
            ProgressBar.IsIndeterminate = false;
        }
    }

    private async void TestNaverUpload_Click(object sender, RoutedEventArgs e)
    {
        var sourcePath = FindUploadExcel();
        if (sourcePath == null)
        {
            MessageBox.Show("업로드용 엑셀 파일을 찾을 수 없습니다.\nCafe24 업로드를 먼저 실행하세요.", "알림",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var dryRun = MarketDryRun.IsChecked == true;
        if (!dryRun)
        {
            var confirm = MessageBox.Show(
                $"네이버 스마트스토어에 상품을 실제 등록합니다.\n\n파일: {Path.GetFileName(sourcePath)}\n\n계속하시겠습니까?",
                "네이버 실제 등록 확인", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (confirm != MessageBoxResult.Yes) return;
        }

        TestNaverUploadButton.IsEnabled = false;
        _cts = new CancellationTokenSource();

        try
        {
            StatusText.Text = dryRun ? "네이버 DRY RUN 중..." : "네이버 등록 중...";
            ProgressBar.IsIndeterminate = true;
            Log($"[네이버] 업로드 시작: {Path.GetFileName(sourcePath)} (DRY RUN: {dryRun})");

            var options = new NaverUploadOptions
            {
                RowStart = ParseInt(MarketRowStart, 0),
                RowEnd = ParseInt(MarketRowEnd, 0),
                DryRun = dryRun,
            };

            var service = new NaverUploadService();
            var progress = new Progress<string>(msg => Log(msg));
            var result = await service.UploadAsync(sourcePath, options, progress, _cts.Token);

            var gridItems = result.Items.Select(item => new MarketResultRow
            {
                Market = "네이버",
                Row = item.Row,
                Name = item.Name,
                Status = item.Status,
                Info = item.ProductId,
                Error = item.Error,
            }).ToList();
            MarketUploadResultGrid.ItemsSource = gridItems;
            MarketUploadSummary.Text = $"[네이버] 성공 {result.SuccessCount} / 실패 {result.FailCount} / 전체 {result.TotalCount}";
            StatusText.Text = $"네이버 {(dryRun ? "DRY RUN" : "등록")} 완료";
            Log($"[네이버] 완료: 성공 {result.SuccessCount} / 실패 {result.FailCount} / 전체 {result.TotalCount}");
            foreach (var item in result.Items.Where(i => !string.IsNullOrEmpty(i.Error)))
                Log($"  [네이버 오류] 행{item.Row} {item.Name} → {item.Error}");
        }
        catch (OperationCanceledException) { Log("[네이버] 취소됨"); StatusText.Text = "취소됨"; }
        catch (Exception ex)
        {
            Log($"[네이버] 오류: {ex.Message}");
            StatusText.Text = "네이버 업로드 오류";
            MessageBox.Show(ex.Message, "네이버 업로드 오류", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            TestNaverUploadButton.IsEnabled = true;
            ProgressBar.IsIndeterminate = false;
        }
    }

    private string? FindUploadExcel()
    {
        // 1) LLM 결과 파일 우선 사용 (키워드/검색어/태그가 적용된 파일)
        if (_testLlmResultFiles.Count > 0)
        {
            var first = _testLlmResultFiles.FirstOrDefault(File.Exists);
            if (first != null)
            {
                Log($"LLM 결과 파일 사용: {Path.GetFileName(first)}");
                return first;
            }
        }
        // 2) _testOutputRoot에서 업로드용 엑셀 탐색
        if (!string.IsNullOrEmpty(_testOutputRoot) && Directory.Exists(_testOutputRoot))
        {
            var found = FindLatestFile(_testOutputRoot, "업로드용_*.xlsx");
            if (found != null) return found;
        }
        // 3) _lastOutputRoot에서 업로드용 엑셀 탐색
        if (!string.IsNullOrEmpty(_lastOutputRoot) && Directory.Exists(_lastOutputRoot))
        {
            var found = FindLatestFile(_lastOutputRoot, "업로드용_*.xlsx");
            if (found != null) return found;
        }
        return null;
    }

    private async Task TryUploadLatestMarketPlusCategoryMapAsync(
        string uploadFile,
        IEnumerable<string> llmResultFiles,
        CancellationToken cancellationToken)
    {
        try
        {
            var candidateRoots = new List<string?>
            {
                Path.GetDirectoryName(uploadFile),
                _testOutputRoot,
                _lastOutputRoot
            };

            var priorityFiles = new List<string?> { uploadFile };
            foreach (var file in llmResultFiles)
            {
                priorityFiles.Add(file);

                var fileDir = Path.GetDirectoryName(file);
                candidateRoots.Add(fileDir);
                if (!string.IsNullOrWhiteSpace(fileDir))
                    candidateRoots.Add(Path.GetDirectoryName(fileDir));
            }

            var uploader = new MarketPlusCategoryMapAutoUploader(
                _v3Root,
                new Progress<string>(msg => Log(msg)));

            await uploader.UploadLatestAsync(candidateRoots, priorityFiles, cancellationToken);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            Log($"[카테고리맵] 자동 업로드 실패: {ex.Message}");
        }
    }

    private sealed class MarketResultRow
    {
        public string Market { get; set; } = "";
        public int Row { get; set; }
        public string Name { get; set; } = "";
        public string Status { get; set; } = "";
        public string Info { get; set; } = "";
        public string Error { get; set; } = "";
    }

    #endregion

    #region ═══ STEP 2: Cafe24 업로드 ═══

    private Cafe24UploadOptions BuildUploadOptions()
    {
        return new Cafe24UploadOptions
        {
            TokenFilePath = string.IsNullOrWhiteSpace(SettingsTokenPath.Text) ? null : SettingsTokenPath.Text.Trim(),
            DateTag = Cafe24DateTag.Text.Trim(),
            ExportDir = _lastOutputRoot ?? "",
            MainIndex = ParseInt(Cafe24MainIdx, 2),
            AddStart = ParseInt(Cafe24AddStart, 3),
            AddMax = ParseInt(Cafe24AddMax, 10),
            RetryCount = ParseInt(Cafe24RetryCount, 1),
            RetryDelaySeconds = ParseDouble(Cafe24RetryDelay, 1.0),
            MatchMode = (Cafe24MatchMode.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "PREFIX",
            MatchPrefix = ParseInt(Cafe24MatchPrefix, 20),
        };
    }

    private async void Cafe24Upload_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_lastOutputRoot) || !Directory.Exists(_lastOutputRoot))
        {
            MessageBox.Show("먼저 STEP 1을 실행하세요.", "알림", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (!TryGetSelectedCafe24Markets(out var runHomeMarket, out var runReadyMarket, out var marketLabel))
        {
            return;
        }

        var uploadFile = !string.IsNullOrEmpty(_lastOutputFile) && File.Exists(_lastOutputFile)
            ? _lastOutputFile
            : FindLatestFile(_lastOutputRoot, "업로드용_*.xlsx");
        if (uploadFile == null)
        {
            MessageBox.Show("업로드용 엑셀 파일을 찾을 수 없습니다.", "알림", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var confirm = MessageBox.Show(
            $"Cafe24에 이미지 업로드 + 옵션가격을 반영합니다.\n\n" +
            $"대상 몰: {marketLabel}\n" +
            $"업로드 파일: {Path.GetFileName(uploadFile)}\n" +
            $"결과 폴더: {_lastOutputRoot}\n\n계속하시겠습니까?",
            "Cafe24 업로드 확인", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (confirm != MessageBoxResult.Yes) return;

        SavePriceReviewJson();

        Cafe24UploadButton.IsEnabled = false;
        _cts = new CancellationTokenSource();

        try
        {
            StatusText.Text = "Cafe24 업로드 중...";
            ProgressBar.IsIndeterminate = true;

            var options = BuildUploadOptions();
            var priceDataPath = Path.Combine(_lastOutputRoot, "cafe24_price_upload_data.json");
            if (File.Exists(priceDataPath))
                options.PriceDataPath = priceDataPath;

            var uploadService = new Cafe24UploadService(_v3Root, _legacyRoot);
            var progress = new Progress<string>(msg => Log(msg));
            var totalCount = 0;
            var totalSuccess = 0;
            var totalError = 0;
            var totalSkipped = 0;
            string? lastLogPath = null;

            void Accumulate(Cafe24UploadResult result)
            {
                totalCount += result.TotalCount;
                totalSuccess += result.SuccessCount;
                totalError += result.ErrorCount;
                totalSkipped += result.SkippedCount;
                if (!string.IsNullOrWhiteSpace(result.LogPath))
                {
                    lastLogPath = result.LogPath;
                }
            }

            async Task RunReadyMarketUploadAsync()
            {
                StatusText.Text = "준비몰 Cafe24 업로드 중...";
                var resultB = await uploadService.UploadBMarketAsync(uploadFile, _lastOutputRoot, options, progress, _cts.Token, _bMarketTokenPath);
                Accumulate(resultB);
                if (resultB.TotalCount > 0)
                {
                    Log($"Cafe24 준비몰 업로드 완료: 성공 {resultB.SuccessCount} / 오류 {resultB.ErrorCount} / 스킵 {resultB.SkippedCount}");
                }
                else
                {
                    Log("Cafe24 준비몰 업로드 스킵: B마켓 시트 또는 대상 상품이 없습니다.");
                }
            }

            if (runHomeMarket)
            {
                StatusText.Text = "홈런마켓 Cafe24 업로드 중...";
                var result = await uploadService.UploadAsync(uploadFile, _lastOutputRoot, options, progress, _cts.Token);
                Accumulate(result);
                Log($"Cafe24 홈런마켓 업로드 완료: 성공 {result.SuccessCount} / 오류 {result.ErrorCount} / 스킵 {result.SkippedCount}");
            }

            if (runReadyMarket)
            {
                if (runHomeMarket)
                {
                    try
                    {
                        await RunReadyMarketUploadAsync();
                    }
                    catch (Cafe24ReauthenticationRequiredException exB)
                    {
                        Log($"준비몰 Cafe24 토큰 오류: {exB.Message}");
                        MessageBox.Show("준비몰 토큰이 만료됐습니다. 설정 탭에서 토큰 파일을 교체해 주세요.", "준비몰 토큰 오류", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                    catch (Exception exB)
                    {
                        Log($"준비몰 업로드 오류 (홈런마켓은 성공): {exB.Message}");
                    }
                }
                else
                {
                    await RunReadyMarketUploadAsync();
                }
            }

            _lastUploadLogPath = lastLogPath;
            StatusText.Text = $"업로드 완료 ({marketLabel}, 성공: {totalSuccess})";
            UploadSummaryText.Text = $"{marketLabel} | 총 {totalCount} | 성공 {totalSuccess} | 오류 {totalError} | 스킵 {totalSkipped}";
            OpenUploadLogButton.IsEnabled = !string.IsNullOrEmpty(lastLogPath);
            if (!string.IsNullOrEmpty(lastLogPath))
            {
                LoadUploadLog(lastLogPath);
            }
        }
        catch (OperationCanceledException) { Log("업로드 취소됨"); StatusText.Text = "취소됨"; }
        catch (Cafe24ReauthenticationRequiredException ex)
        {
            Log($"Cafe24 토큰 오류: {ex.Message}");
            StatusText.Text = "토큰 오류";
            MessageBox.Show("Cafe24 토큰이 만료됐습니다. 설정 탭에서 토큰 파일을 교체해 주세요.", "Cafe24 토큰 오류", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
        catch (Exception ex)
        {
            Log($"Cafe24 오류: {ex.Message}");
            StatusText.Text = "업로드 오류";
            MessageBox.Show(ex.Message, "Cafe24 업로드 오류", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            Cafe24UploadButton.IsEnabled = true;
            ProgressBar.IsIndeterminate = false;
        }
    }

    private async void Cafe24Create_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_lastOutputRoot) || !Directory.Exists(_lastOutputRoot))
        {
            MessageBox.Show("먼저 STEP 1을 실행하세요.", "알림", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (!TryGetSelectedCafe24Markets(out var runHomeMarket, out var runReadyMarket, out var marketLabel))
        {
            return;
        }

        var uploadFile = !string.IsNullOrEmpty(_lastOutputFile) && File.Exists(_lastOutputFile)
            ? _lastOutputFile
            : FindLatestFile(_lastOutputRoot, "업로드용_*.xlsx");
        if (uploadFile == null)
        {
            MessageBox.Show("업로드용 엑셀을 찾을 수 없습니다.", "알림", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // 준비몰 선택 시 B마켓 시트 존재 여부 미리 확인
        var bMarketNote = "";
        if (runReadyMarket && !HasBMarketSheet(uploadFile))
            bMarketNote = "\n\n⚠ 업로드 파일에 B마켓 시트가 없습니다.\n   준비몰 신규등록은 스킵됩니다.";

        var confirm = MessageBox.Show(
            $"Cafe24에 신규 상품을 등록합니다.\n\n" +
            $"대상 몰: {marketLabel}\n" +
            $"업로드 파일: {Path.GetFileName(uploadFile)}{bMarketNote}\n\n계속하시겠습니까?",
            "신규상품 등록 확인", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (confirm != MessageBoxResult.Yes) return;

        // 네이버(홈런마켓) 중복 확인
        if (runHomeMarket)
        {
            StatusText.Text = "네이버 중복 확인 중...";
            var duplicateInfo = await CheckNaverDuplicatesAsync(uploadFile);
            if (duplicateInfo.Count > 0)
            {
                var dupLines = duplicateInfo
                    .Select(d => $"  • {d.GsCode}  {d.ProductName}")
                    .ToList();
                var msg = $"다음 {duplicateInfo.Count}개 상품이 네이버(홈런마켓)에 이미 등록되어 있습니다:\n\n" +
                          string.Join("\n", dupLines) +
                          "\n\n이미 등록된 상품도 포함하여 계속 진행하시겠습니까?";
                var dupResult = MessageBox.Show(msg, "네이버 중복 확인", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (dupResult != MessageBoxResult.Yes) return;
            }
        }

        Cafe24CreateButton.IsEnabled = false;
        _cts = new CancellationTokenSource();

        try
        {
            StatusText.Text = "신규상품 등록 중...";
            ProgressBar.IsIndeterminate = true;

            var options = BuildUploadOptions();
            var createService = new Cafe24CreateProductService(_v3Root, _legacyRoot);
            var progress = new Progress<string>(msg => Log(msg));
            var totalCreated = 0;
            var totalError = 0;
            var totalSkipped = 0;

            // 체크된 GS 코드만 등록 (목록이 있을 때)
            IReadOnlySet<string>? selectedGs = _cafe24Items.Count > 0
                ? new HashSet<string>(_cafe24Items.Where(i => i.IsChecked).Select(i => i.GsCode), StringComparer.OrdinalIgnoreCase)
                : null;

            async Task RunReadyMarketCreateAsync()
            {
                StatusText.Text = "준비몰 신규상품 등록 중...";
                Log("── [준비몰] 신규등록 시작 ──");
                var resultB = await createService.CreateBMarketAsync(uploadFile, _lastOutputRoot, progress, _cts.Token, _bMarketTokenPath, selectedGs);
                totalCreated += resultB.CreatedCount;
                totalError += resultB.ErrorCount;
                totalSkipped += resultB.SkippedCount;
                if (resultB.TotalCount > 0)
                {
                    Log($"[준비몰] 신규등록 완료: 생성 {resultB.CreatedCount} / 오류 {resultB.ErrorCount} / 스킵 {resultB.SkippedCount}");
                }
                else
                {
                    Log("[준비몰] 신규등록 스킵: B마켓 시트에 등록 대상이 없습니다.");
                }
            }

            if (runHomeMarket)
            {
                StatusText.Text = "홈런마켓 신규상품 등록 중...";
                var aTokenPath = string.IsNullOrWhiteSpace(SettingsTokenPath.Text) ? null : SettingsTokenPath.Text.Trim();
                var result = await createService.CreateAsync(uploadFile, _lastOutputRoot, progress, _cts.Token, tokenPath: aTokenPath, allowedGsCodes: selectedGs);
                totalCreated += result.CreatedCount;
                totalError += result.ErrorCount;
                totalSkipped += result.SkippedCount;
                // 업로드 이력 기록
                foreach (var item in _cafe24Items.Where(i => i.IsChecked))
                {
                    _uploadHistory.Mark(item.GsCode, "homemarket");
                    item.HomeMarketStatus = UploadProductItem.FormatDate(DateTime.Now);
                }
                Log($"[홈런마켓] 신규등록 완료: 생성 {result.CreatedCount} / 오류 {result.ErrorCount} / 스킵 {result.SkippedCount}");
            }

            if (runReadyMarket)
            {
                if (runHomeMarket)
                {
                    try
                    {
                        await RunReadyMarketCreateAsync();
                    }
                    catch (Cafe24ReauthenticationRequiredException exB)
                    {
                        Log($"[준비몰] 토큰 오류: {exB.Message}");
                        MessageBox.Show("준비몰 토큰이 만료됐습니다. 설정 탭에서 토큰 파일을 교체해 주세요.", "준비몰 토큰 오류", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                    catch (Exception exB)
                    {
                        Log($"[준비몰] 신규등록 오류: {exB.Message}");
                    }
                }
                else
                {
                    await RunReadyMarketCreateAsync();
                }
            }

            StatusText.Text = $"등록 완료 ({marketLabel}, 생성: {totalCreated})";
        }
        catch (OperationCanceledException) { Log("등록 취소됨"); StatusText.Text = "취소됨"; }
        catch (Cafe24ReauthenticationRequiredException ex)
        {
            Log($"Cafe24 토큰 오류: {ex.Message}");
            StatusText.Text = "토큰 오류";
            MessageBox.Show("Cafe24 토큰이 만료됐습니다. 설정 탭에서 토큰 파일을 교체해 주세요.", "Cafe24 토큰 오류", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
        catch (Exception ex)
        {
            Log($"등록 오류: {ex.Message}");
            MessageBox.Show(ex.Message, "신규상품 등록 오류", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            Cafe24CreateButton.IsEnabled = true;
            ProgressBar.IsIndeterminate = false;
        }
    }

    private void CoupangSource_DragOver(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
            e.Effects = DragDropEffects.Copy;
        else
            e.Effects = DragDropEffects.None;
        e.Handled = true;
    }

    private void CoupangSource_Drop(object sender, DragEventArgs e)
    {
        if (e.Data.GetData(DataFormats.FileDrop) is string[] files && files.Length > 0)
        {
            CoupangSourcePath.Text = files[0];
        }
    }

    private void CoupangBrowseSource_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "Excel Files|*.xlsx;*.xls",
            Title = "가공파일 선택",
            InitialDirectory = @"C:\code\exports",
        };
        if (dlg.ShowDialog() == true)
        {
            CoupangSourcePath.Text = dlg.FileName;
        }
    }

    private async void CoupangUpload_Click(object sender, RoutedEventArgs e)
    {
        var sourcePath = CoupangSourcePath.Text.Trim();
        if (string.IsNullOrEmpty(sourcePath) || !File.Exists(sourcePath))
        {
            // _lastOutputRoot에서 자동 탐색
            if (!string.IsNullOrEmpty(_lastOutputRoot) && Directory.Exists(_lastOutputRoot))
            {
                var found = FindLatestFile(_lastOutputRoot, "업로드용_*.xlsx");
                if (found != null)
                {
                    sourcePath = found;
                    CoupangSourcePath.Text = sourcePath;
                }
                else
                {
                    MessageBox.Show("가공파일(업로드용 엑셀)을 선택해주세요.", "알림", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
            }
            else
            {
                MessageBox.Show("가공파일(업로드용 엑셀)을 선택해주세요.", "알림", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
        }

        var dryRun = CoupangDryRun.IsChecked == true;
        if (!dryRun)
        {
            var confirm = MessageBox.Show(
                $"쿠팡에 상품을 실제 등록합니다.\n\n파일: {Path.GetFileName(sourcePath)}\n\n계속하시겠습니까?",
                "쿠팡 실제 등록 확인", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (confirm != MessageBoxResult.Yes) return;
        }

        CoupangUploadButton.IsEnabled = false;
        _cts = new CancellationTokenSource();

        try
        {
            StatusText.Text = dryRun ? "쿠팡 DRY RUN 중..." : "쿠팡 등록 중...";
            ProgressBar.IsIndeterminate = true;
            Log($"쿠팡 업로드 시작: {Path.GetFileName(sourcePath)} (DRY RUN: {dryRun})");

            var options = new CoupangUploadOptions { DryRun = dryRun };

            // 체크된 행만 처리 (목록이 있을 때)
            IReadOnlySet<int>? selectedRows = _coupangItems.Count > 0
                ? new HashSet<int>(_coupangItems.Where(i => i.IsChecked).Select(i => i.RowNum))
                : null;

            var service = new CoupangUploadService();
            var progress = new Progress<string>(msg => Log(msg));

            var result = await service.UploadAsync(sourcePath, options, progress, _cts.Token, selectedRows);

            // 업로드 이력 기록
            if (!dryRun)
            {
                foreach (var item in _coupangItems.Where(i => i.IsChecked && !string.IsNullOrEmpty(i.GsCode)))
                {
                    _uploadHistory.Mark(item.GsCode, "coupang");
                    item.CoupangStatus = UploadProductItem.FormatDate(DateTime.Now);
                }
            }

            // 결과 그리드
            var gridItems = result.Items.Select(item => new CoupangResultRow
            {
                Row = item.Row,
                Name = item.Name,
                Status = item.Status,
                Info = !string.IsNullOrEmpty(item.SellerProductId) ? item.SellerProductId : item.Category,
                Error = item.Error,
            }).ToList();
            CoupangResultGrid.ItemsSource = gridItems;
            CoupangSummaryText.Text = $"성공 {result.SuccessCount} / 실패 {result.FailCount} / 전체 {result.TotalCount}";
            StatusText.Text = $"쿠팡 {(dryRun ? "DRY RUN" : "등록")} 완료";
            Log($"쿠팡 완료: 성공 {result.SuccessCount} / 실패 {result.FailCount} / 전체 {result.TotalCount}");
        }
        catch (OperationCanceledException) { Log("쿠팡 업로드 취소됨"); StatusText.Text = "취소됨"; }
        catch (Exception ex)
        {
            Log($"쿠팡 오류: {ex.Message}");
            StatusText.Text = "쿠팡 업로드 오류";
            MessageBox.Show(ex.Message, "쿠팡 업로드 오류", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            CoupangUploadButton.IsEnabled = true;
            ProgressBar.IsIndeterminate = false;
        }
    }

    private sealed class CoupangResultRow
    {
        public int Row { get; set; }
        public string Name { get; set; } = "";
        public string Status { get; set; } = "";
        public string Info { get; set; } = "";
        public string Error { get; set; } = "";
    }

    private void LoadUploadLog(string? logPath)
    {
        if (string.IsNullOrEmpty(logPath) || !File.Exists(logPath)) return;

        try
        {
            using var wb = new XLWorkbook(logPath);
            var ws = wb.Worksheets.First();
            var rows = new List<UploadResultRow>();
            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;

            for (int r = 2; r <= lastRow; r++)
            {
                rows.Add(new UploadResultRow
                {
                    Gs = ws.Cell(r, 1).GetString(),
                    ProductNo = ws.Cell(r, 2).GetString(),
                    Status = ws.Cell(r, 3).GetString(),
                    MainImage = ws.Cell(r, 4).GetString(),
                    AddCount = ws.Cell(r, 5).GetString(),
                    PriceStatus = ws.Cell(r, 6).GetString(),
                });
            }
            UploadResultGrid.ItemsSource = rows;
        }
        catch (Exception ex)
        {
            Log($"로그 읽기 오류: {ex.Message}");
        }
    }

    private void OpenUploadLog_Click(object sender, RoutedEventArgs e)
    {
        if (!string.IsNullOrEmpty(_lastUploadLogPath) && File.Exists(_lastUploadLogPath))
            Process.Start(new ProcessStartInfo(_lastUploadLogPath) { UseShellExecute = true });
    }

    #endregion

    #region ═══ 옵션 가격 관리 ═══

    private void LoadPriceData_Click(object sender, RoutedEventArgs e)
    {
        string? gptFile = null;

        // 결과 폴더에서 자동 탐색
        if (!string.IsNullOrEmpty(_lastOutputRoot))
            gptFile = FindLatestFile(_lastOutputRoot, "상품전처리GPT_*.xlsx");

        // 없으면 파일 선택
        if (gptFile == null)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "Excel|*.xlsx|모든 파일|*.*",
                Title = "상품전처리GPT 파일 선택",
            };
            if (dlg.ShowDialog() != true) return;
            gptFile = dlg.FileName;
        }

        LoadPriceFromExcel(gptFile);
    }

    private void LoadPriceFromExcel(string filePath)
    {
        _priceRows.Clear();
        try
        {
            using var wb = new XLWorkbook(filePath);

            // 분리추출전 시트에서 옵션+공급가 로드
            var sheetName = wb.Worksheets.Any(w => w.Name == "분리추출전") ? "분리추출전" : wb.Worksheets.First().Name;
            var ws = wb.Worksheet(sheetName);
            var headerRow = ws.FirstRowUsed();
            if (headerRow == null) return;

            var lastCol = headerRow.LastCellUsed()?.Address.ColumnNumber ?? 0;
            var cols = new Dictionary<string, int>();

            for (int c = 1; c <= lastCol; c++)
            {
                var h = headerRow.Cell(c).GetString().Trim();
                if (!string.IsNullOrEmpty(h))
                    cols[h] = c;
            }

            int codeCol = FindCol(cols, CodeColumns);
            int nameCol = FindCol(cols, NameColumns);
            int supplyCol = FindCol(cols, new[] { "공급가", "supply_price" });
            int sellingCol = FindCol(cols, new[] { "판매가", "selling_price" });
            int consumerCol = FindCol(cols, new[] { "소비자가", "consumer_price" });

            if (codeCol < 0) { Log("상품코드 컬럼을 찾을 수 없습니다."); return; }

            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
            for (int r = headerRow.RowNumber() + 1; r <= lastRow; r++)
            {
                var code = ws.Cell(r, codeCol).GetString().Trim();
                if (string.IsNullOrEmpty(code)) continue;

                var gsMatch = Regex.Match(code, @"(GS\d{7})", RegexOptions.IgnoreCase);
                var gsCode = gsMatch.Success ? gsMatch.Value : code;

                var supply = supplyCol > 0 ? GetDecimal(ws.Cell(r, supplyCol)) : 0;
                var selling = sellingCol > 0 ? GetDecimal(ws.Cell(r, sellingCol)) : 0;
                var consumer = consumerCol > 0 ? GetDecimal(ws.Cell(r, consumerCol)) : 0;
                var optionName = nameCol > 0 ? ws.Cell(r, nameCol).GetString().Trim() : "";

                _priceRows.Add(new PriceRow
                {
                    IsChecked = true,
                    GsCode = gsCode,
                    OptionName = optionName,
                    SupplyPrice = supply,
                    SellingPrice = selling,
                    AdditionalAmount = 0,
                    ConsumerPrice = consumer,
                });
            }

            PriceFileText.Text = Path.GetFileName(filePath);
            PriceSummaryText.Text = $"{_priceRows.Count}개 행 로드됨";
            Log($"가격 데이터 로드: {_priceRows.Count}개 ({Path.GetFileName(filePath)})");
        }
        catch (Exception ex)
        {
            Log($"가격 파일 오류: {ex.Message}");
        }
    }

    private void RecalcPrices_Click(object sender, RoutedEventArgs e)
    {
        if (_priceRows.Count == 0) return;

        // GS코드별 그룹 → 최고 공급가 기준 추가금액 계산
        var groups = _priceRows.GroupBy(r => r.GsCode).ToList();
        foreach (var group in groups)
        {
            var items = group.ToList();
            if (items.Count <= 1) continue;

            var maxSupply = items.Max(r => r.SupplyPrice);
            foreach (var row in items)
            {
                row.AdditionalAmount = row.SupplyPrice - items.Min(r => r.SupplyPrice);
            }
        }

        PriceGrid.Items.Refresh();
        Log("추가금액 재계산 완료");
    }

    private void SavePriceData_Click(object sender, RoutedEventArgs e) => SavePriceReviewJson();

    private void SavePriceReviewJson()
    {
        if (_priceRows.Count == 0 || string.IsNullOrEmpty(_lastOutputRoot)) return;

        try
        {
            var checkedGs = _priceRows.Where(r => r.IsChecked)
                .Select(r => r.GsCode).Distinct().ToList();

            var editedAmounts = new Dictionary<string, List<decimal>>();
            foreach (var group in _priceRows.GroupBy(r => r.GsCode))
            {
                var amounts = group.Select(r => r.AdditionalAmount).ToList();
                if (amounts.Any(a => a != 0))
                    editedAmounts[group.Key] = amounts;
            }

            var data = new
            {
                checked_gs = checkedGs,
                edited_amounts = editedAmounts,
                image_selections = new Dictionary<string, object>(),
            };

            var json = JsonSerializer.Serialize(data, new JsonSerializerOptions { WriteIndented = true });
            var path = Path.Combine(_lastOutputRoot, "cafe24_price_upload_data.json");
            File.WriteAllText(path, json, Encoding.UTF8);
            Log($"가격 데이터 저장: {Path.GetFileName(path)}");
        }
        catch (Exception ex)
        {
            Log($"저장 오류: {ex.Message}");
        }
    }

    private void PriceSelectAll_Click(object sender, RoutedEventArgs e)
    {
        foreach (var r in _priceRows) r.IsChecked = true;
        PriceGrid.Items.Refresh();
    }

    private void PriceDeselectAll_Click(object sender, RoutedEventArgs e)
    {
        foreach (var r in _priceRows) r.IsChecked = false;
        PriceGrid.Items.Refresh();
    }

    #endregion

    #region ═══ 실행 이력 ═══

    private void RefreshHistoryGrid()
    {
        if (_jobHistory == null) return;
        HistoryGrid.ItemsSource = null;
        HistoryGrid.ItemsSource = _jobHistory.Records;
    }

    private void RefreshHistory_Click(object sender, RoutedEventArgs e) => RefreshHistoryGrid();

    private JobRecord? GetSelectedJob()
    {
        return HistoryGrid.SelectedItem as JobRecord;
    }

    private void HistoryGrid_DoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        HistoryLoad_Click(sender, e);
    }

    private void HistoryLoad_Click(object sender, RoutedEventArgs e)
    {
        var job = GetSelectedJob();
        if (job == null) { MessageBox.Show("이력을 선택하세요.", "알림"); return; }

        if (!Directory.Exists(job.OutputRoot))
        {
            MessageBox.Show($"결과 폴더가 존재하지 않습니다.\n{job.OutputRoot}", "알림",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        _lastOutputRoot = job.OutputRoot;
        _lastOutputFile = job.OutputFile;
        OutputFileText.Text = job.OutputFile;

        OpenUploadExcelButton.IsEnabled = true;
        OpenOutputFolderButton.IsEnabled = true;
        Cafe24UploadButton.IsEnabled = true;
        Cafe24CreateButton.IsEnabled = true;

        Log($"이력 불러오기: {job.DisplaySource} ({job.DisplayTime})");
        StatusText.Text = $"이력 로드됨 — {job.DisplaySource}";
    }

    private void HistoryOpenFolder_Click(object sender, RoutedEventArgs e)
    {
        var job = GetSelectedJob();
        if (job == null) return;
        if (Directory.Exists(job.OutputRoot))
            Process.Start(new ProcessStartInfo("explorer.exe", job.OutputRoot));
        else
            MessageBox.Show("폴더가 존재하지 않습니다.", "알림");
    }

    private void HistoryCopy_Click(object sender, RoutedEventArgs e)
    {
        var job = GetSelectedJob();
        if (job == null) { MessageBox.Show("이력을 선택하세요.", "알림"); return; }
        _jobHistory?.Clone(job);
        RefreshHistoryGrid();
        Log($"이력 복사: {job.DisplaySource}");
    }

    private void HistoryEditMemo_Click(object sender, RoutedEventArgs e)
    {
        var job = GetSelectedJob();
        if (job == null) { MessageBox.Show("이력을 선택하세요.", "알림"); return; }

        var dlg = new MemoDialog(job.Memo) { Owner = this };
        if (dlg.ShowDialog() == true)
        {
            job.Memo = dlg.MemoText;
            _jobHistory?.Update(job);
            RefreshHistoryGrid();
        }
    }

    private void HistoryDelete_Click(object sender, RoutedEventArgs e)
    {
        var job = GetSelectedJob();
        if (job == null) { MessageBox.Show("이력을 선택하세요.", "알림"); return; }

        var confirm = MessageBox.Show(
            $"이력을 삭제하시겠습니까?\n{job.DisplaySource} ({job.DisplayTime})",
            "삭제 확인", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (confirm != MessageBoxResult.Yes) return;

        _jobHistory?.Delete(job.Id);
        RefreshHistoryGrid();
        Log($"이력 삭제: {job.DisplaySource}");
    }

    private bool _historyAllSelected = false;
    private void HistorySelectAll_Click(object sender, RoutedEventArgs e)
    {
        if (_historyAllSelected)
            HistoryGrid.UnselectAll();
        else
            HistoryGrid.SelectAll();
        _historyAllSelected = !_historyAllSelected;
    }

    private void HistoryBulkDelete_Click(object sender, RoutedEventArgs e)
    {
        var selected = HistoryGrid.SelectedItems.Cast<JobRecord>().ToList();
        if (selected.Count == 0) { MessageBox.Show("삭제할 이력을 선택하세요.", "알림"); return; }

        var confirm = MessageBox.Show(
            $"{selected.Count}개 이력을 삭제하시겠습니까?",
            "선택 삭제 확인", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (confirm != MessageBoxResult.Yes) return;

        foreach (var job in selected)
            _jobHistory?.Delete(job.Id);
        _historyAllSelected = false;
        RefreshHistoryGrid();
        Log($"이력 {selected.Count}개 일괄 삭제");
    }

    private void HistoryOpenCafe24Excel_Click(object sender, RoutedEventArgs e)
    {
        var job = GetSelectedJob();
        if (job == null) return;

        // 업로드용 엑셀 경로를 클립보드에 복사
        var uploadFile = FindLatestFile(job.OutputRoot, "업로드용_*.xlsx");
        if (uploadFile != null)
        {
            Clipboard.SetText(uploadFile);
            Log($"업로드용 엑셀 경로 클립보드 복사: {Path.GetFileName(uploadFile)}");
        }

        // Cafe24 상품 엑셀 관리 페이지 열기
        try
        {
            var store = new Cafe24ConfigStore(_v3Root, _legacyRoot);
            var state = store.LoadTokenState(
                string.IsNullOrWhiteSpace(SettingsTokenPath.Text) ? null : SettingsTokenPath.Text.Trim());
            var mallId = state.Config.MallId;
            if (!string.IsNullOrEmpty(mallId))
            {
                var url = $"https://{mallId}.cafe24.com/disp/admin/shop1/product/ProductExcelManage";
                Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
                Log("Cafe24 상품 엑셀 관리 페이지 열림");
            }
            else
            {
                MessageBox.Show("Mall ID가 설정되지 않았습니다. 설정 탭에서 토큰 파일을 확인하세요.", "알림");
            }
        }
        catch (Exception ex)
        {
            Log($"Cafe24 페이지 열기 오류: {ex.Message}");
        }
    }

    private async void HistoryCafe24Upload_Click(object sender, RoutedEventArgs e)
    {
        var job = GetSelectedJob();
        if (job == null) return;

        if (!TryGetSelectedCafe24Markets(out var runHomeMarket, out var runReadyMarket, out var marketLabel))
        {
            return;
        }

        if (!Directory.Exists(job.OutputRoot))
        {
            MessageBox.Show("결과 폴더가 존재하지 않습니다.", "알림"); return;
        }

        var uploadFile = FindLatestFile(job.OutputRoot, "업로드용_*.xlsx");
        if (uploadFile == null)
        {
            MessageBox.Show("업로드용 엑셀을 찾을 수 없습니다.", "알림"); return;
        }

        var confirm = MessageBox.Show(
            $"Cafe24에 이미지 + 옵션가격을 업로드합니다.\n\n" +
            $"대상 몰: {marketLabel}\n" +
            $"파일: {Path.GetFileName(uploadFile)}\n" +
            $"폴더: {job.OutputRoot}\n\n계속하시겠습니까?",
            "Cafe24 업로드 확인", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (confirm != MessageBoxResult.Yes) return;

        _lastOutputRoot = job.OutputRoot;
        _lastOutputFile = job.OutputFile;

        _cts = new CancellationTokenSource();
        Cafe24UploadButton.IsEnabled = false;

        try
        {
            StatusText.Text = "Cafe24 업로드 중...";
            ProgressBar.IsIndeterminate = true;

            SavePriceReviewJson();
            var options = BuildUploadOptions();
            options.ExportDir = job.OutputRoot;

            var priceDataPath = Path.Combine(job.OutputRoot, "cafe24_price_upload_data.json");
            if (File.Exists(priceDataPath))
                options.PriceDataPath = priceDataPath;

            var uploadService = new Cafe24UploadService(_v3Root, _legacyRoot);
            var progress = new Progress<string>(msg => Log(msg));
            var totalCount = 0;
            var totalSuccess = 0;
            var totalError = 0;
            var totalSkipped = 0;
            string? lastLogPath = null;

            void Accumulate(Cafe24UploadResult result)
            {
                totalCount += result.TotalCount;
                totalSuccess += result.SuccessCount;
                totalError += result.ErrorCount;
                totalSkipped += result.SkippedCount;
                if (!string.IsNullOrWhiteSpace(result.LogPath))
                {
                    lastLogPath = result.LogPath;
                }
            }

            async Task RunReadyMarketUploadAsync()
            {
                StatusText.Text = "준비몰 Cafe24 업로드 중...";
                var resultB = await uploadService.UploadBMarketAsync(uploadFile, job.OutputRoot, options, progress, _cts.Token, _bMarketTokenPath);
                Accumulate(resultB);
                if (resultB.TotalCount > 0)
                {
                    Log($"Cafe24 준비몰 업로드 완료: 성공 {resultB.SuccessCount} / 오류 {resultB.ErrorCount} / 스킵 {resultB.SkippedCount}");
                }
                else
                {
                    Log("Cafe24 준비몰 업로드 스킵: B마켓 시트 또는 대상 상품이 없습니다.");
                }
            }

            if (runHomeMarket)
            {
                StatusText.Text = "홈런마켓 Cafe24 업로드 중...";
                var result = await uploadService.UploadAsync(uploadFile, job.OutputRoot, options, progress, _cts.Token);
                Accumulate(result);
                Log($"Cafe24 홈런마켓 업로드 완료: 성공 {result.SuccessCount} / 오류 {result.ErrorCount} / 스킵 {result.SkippedCount}");
            }

            if (runReadyMarket)
            {
                if (runHomeMarket)
                {
                    try
                    {
                        await RunReadyMarketUploadAsync();
                    }
                    catch (Cafe24ReauthenticationRequiredException exB)
                    {
                        Log($"준비몰 Cafe24 토큰 오류: {exB.Message}");
                        MessageBox.Show("준비몰 토큰이 만료됐습니다. 설정 탭에서 토큰 파일을 교체해 주세요.", "준비몰 토큰 오류", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                    catch (Exception exB)
                    {
                        Log($"준비몰 업로드 오류 (홈런마켓은 성공): {exB.Message}");
                    }
                }
                else
                {
                    await RunReadyMarketUploadAsync();
                }
            }

            _lastUploadLogPath = lastLogPath;
            Log($"선택한 몰 Cafe24 업로드 완료: 대상 {marketLabel} / 성공 {totalSuccess} / 오류 {totalError} / 스킵 {totalSkipped}");
            StatusText.Text = $"업로드 완료 ({marketLabel}, 성공: {totalSuccess})";
            UploadSummaryText.Text = $"{marketLabel} | 총 {totalCount} | 성공 {totalSuccess} | 오류 {totalError} | 스킵 {totalSkipped}";
            OpenUploadLogButton.IsEnabled = !string.IsNullOrEmpty(lastLogPath);
            if (!string.IsNullOrEmpty(lastLogPath))
            {
                LoadUploadLog(lastLogPath);
            }
        }
        catch (OperationCanceledException) { Log("업로드 취소됨"); StatusText.Text = "취소됨"; }
        catch (Cafe24ReauthenticationRequiredException ex)
        {
            Log($"Cafe24 토큰 오류: {ex.Message}");
            StatusText.Text = "토큰 오류";
            MessageBox.Show("Cafe24 토큰이 만료됐습니다. 설정 탭에서 토큰 파일을 교체해 주세요.", "Cafe24 토큰 오류", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
        catch (Exception ex)
        {
            Log($"Cafe24 오류: {ex.Message}");
            StatusText.Text = "업로드 오류";
            MessageBox.Show(ex.Message, "Cafe24 업로드 오류", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            Cafe24UploadButton.IsEnabled = true;
            ProgressBar.IsIndeterminate = false;
        }
    }

    private async void HistoryCafe24Create_Click(object sender, RoutedEventArgs e)
    {
        var job = GetSelectedJob();
        if (job == null) return;

        if (!TryGetSelectedCafe24Markets(out var runHomeMarket, out var runReadyMarket, out var marketLabel))
        {
            return;
        }

        if (!Directory.Exists(job.OutputRoot))
        {
            MessageBox.Show("결과 폴더가 존재하지 않습니다.", "알림"); return;
        }

        var uploadFile = FindLatestFile(job.OutputRoot, "업로드용_*.xlsx");
        if (uploadFile == null)
        {
            MessageBox.Show("업로드용 엑셀을 찾을 수 없습니다.", "알림"); return;
        }

        var confirm = MessageBox.Show(
            $"Cafe24에 신규 상품을 등록합니다.\n\n" +
            $"대상 몰: {marketLabel}\n" +
            $"파일: {Path.GetFileName(uploadFile)}\n\n계속하시겠습니까?",
            "신규상품 등록 확인", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (confirm != MessageBoxResult.Yes) return;

        _lastOutputRoot = job.OutputRoot;
        _cts = new CancellationTokenSource();
        Cafe24CreateButton.IsEnabled = false;

        try
        {
            StatusText.Text = "신규상품 등록 중...";
            ProgressBar.IsIndeterminate = true;

            var createService = new Cafe24CreateProductService(_v3Root, _legacyRoot);
            var progress = new Progress<string>(msg => Log(msg));
            var totalCreated = 0;
            var totalError = 0;
            var totalSkipped = 0;

            async Task RunReadyMarketCreateAsync()
            {
                StatusText.Text = "준비몰 신규상품 등록 중...";
                Log("── [준비몰] 신규등록 시작 ──");
                var resultB = await createService.CreateBMarketAsync(uploadFile, job.OutputRoot, progress, _cts.Token, _bMarketTokenPath);
                totalCreated += resultB.CreatedCount;
                totalError += resultB.ErrorCount;
                totalSkipped += resultB.SkippedCount;
                if (resultB.TotalCount > 0)
                {
                    Log($"[준비몰] 신규등록 완료: 생성 {resultB.CreatedCount} / 오류 {resultB.ErrorCount} / 스킵 {resultB.SkippedCount}");
                }
                else
                {
                    Log("[준비몰] 신규등록 스킵: B마켓 시트에 등록 대상이 없습니다.");
                }
            }

            if (runHomeMarket)
            {
                StatusText.Text = "홈런마켓 신규상품 등록 중...";
                var aTokenPath = string.IsNullOrWhiteSpace(SettingsTokenPath.Text) ? null : SettingsTokenPath.Text.Trim();
                var result = await createService.CreateAsync(uploadFile, job.OutputRoot, progress, _cts.Token, tokenPath: aTokenPath);
                totalCreated += result.CreatedCount;
                totalError += result.ErrorCount;
                totalSkipped += result.SkippedCount;
                Log($"[홈런마켓] 신규등록 완료: 생성 {result.CreatedCount} / 오류 {result.ErrorCount} / 스킵 {result.SkippedCount}");
            }

            if (runReadyMarket)
            {
                if (runHomeMarket)
                {
                    try
                    {
                        await RunReadyMarketCreateAsync();
                    }
                    catch (Exception exB)
                    {
                        Log($"[준비몰] 신규등록 오류: {exB.Message}");
                    }
                }
                else
                {
                    await RunReadyMarketCreateAsync();
                }
            }

            StatusText.Text = $"등록 완료 ({marketLabel}, 생성: {totalCreated})";
        }
        catch (OperationCanceledException) { Log("등록 취소됨"); StatusText.Text = "취소됨"; }
        catch (Exception ex)
        {
            Log($"등록 오류: {ex.Message}");
            MessageBox.Show(ex.Message, "신규상품 등록 오류", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            Cafe24CreateButton.IsEnabled = true;
            ProgressBar.IsIndeterminate = false;
        }
    }

    private void HistoryOpenExcel_Click(object sender, RoutedEventArgs e)
    {
        var job = GetSelectedJob();
        if (job == null) return;
        var uploadFile = FindLatestFile(job.OutputRoot, "업로드용_*.xlsx");
        if (uploadFile != null && File.Exists(uploadFile))
            Process.Start(new ProcessStartInfo(uploadFile) { UseShellExecute = true });
        else
            MessageBox.Show("업로드용 엑셀을 찾을 수 없습니다.", "알림");
    }

    private void HistoryCopyPath_Click(object sender, RoutedEventArgs e)
    {
        var job = GetSelectedJob();
        if (job == null) return;
        var uploadFile = FindLatestFile(job.OutputRoot, "업로드용_*.xlsx");
        if (uploadFile != null)
        {
            Clipboard.SetText(uploadFile);
            Log($"클립보드 복사: {uploadFile}");
        }
        else
        {
            Clipboard.SetText(job.OutputRoot);
            Log($"클립보드 복사: {job.OutputRoot}");
        }
    }

    private void HistoryViewProducts_Click(object sender, RoutedEventArgs e)
    {
        var job = GetSelectedJob();
        if (job == null) { MessageBox.Show("이력을 선택하세요.", "알림"); return; }

        if (job.SelectedCodes.Count == 0)
        {
            MessageBox.Show("선택된 상품코드가 없습니다.", "알림");
            return;
        }

        var sb = new StringBuilder();
        sb.AppendLine($"작업: {job.DisplaySource} ({job.DisplayTime})");
        sb.AppendLine($"총 {job.SelectedCodes.Count}개 상품코드");
        sb.AppendLine(new string('─', 40));
        for (int i = 0; i < job.SelectedCodes.Count; i++)
            sb.AppendLine($"  {i + 1}. {job.SelectedCodes[i]}");

        MessageBox.Show(sb.ToString(), "상품목록", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void HistoryImageSelect_Click(object sender, RoutedEventArgs e)
    {
        var job = GetSelectedJob();
        if (job == null) { MessageBox.Show("이력을 선택하세요.", "알림"); return; }

        if (!Directory.Exists(job.OutputRoot))
        {
            MessageBox.Show($"결과 폴더가 존재하지 않습니다.\n{job.OutputRoot}", "알림",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // 해당 작업의 결과를 로드
        _lastOutputRoot = job.OutputRoot;
        _lastOutputFile = job.OutputFile;

        // 이미지 불러오기 실행
        LoadImages_Click(sender, e);

        // 이미지선택 탭으로 이동
        ImageSelectionTab.IsSelected = true;
    }

    #endregion

    #region ═══ 설정 탭 ═══

    private void BrowseLogoPath_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog { Filter = "이미지|*.png;*.jpg;*.jpeg|모든 파일|*.*", Title = "로고 파일 선택" };
        if (dlg.ShowDialog() == true)
            SettingsLogoPath.Text = dlg.FileName;
    }

    private void BrowseLogoPathB_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog { Filter = "이미지|*.png;*.jpg;*.jpeg|모든 파일|*.*", Title = "B마켓 로고 파일 선택" };
        if (dlg.ShowDialog() == true)
            SettingsLogoPathB.Text = dlg.FileName;
    }

    private void BrowseTokenPath_Click(object sender, RoutedEventArgs e)
    {
        var keyDir = DesktopKeyStore.DirectoryPath;
        var dlg = new OpenFileDialog
        {
            Filter = "JSON|*.json|모든 파일|*.*",
            Title = "홈런마켓 토큰 JSON 파일 선택",
            InitialDirectory = Directory.Exists(keyDir) ? keyDir : ""
        };
        if (dlg.ShowDialog() == true)
        {
            SettingsTokenPath.Text = dlg.FileName;
            LoadTokenInfo();
        }
    }

    private void LoadTokenInfo()
    {
        try
        {
            var store = new Cafe24ConfigStore(_v3Root, _legacyRoot);
            var tokenPath = string.IsNullOrWhiteSpace(SettingsTokenPath.Text)
                ? null : SettingsTokenPath.Text.Trim();
            var state = store.LoadTokenState(tokenPath);
            SettingsMallId.Text = state.Config.MallId;
            SettingsTokenStatus.Text = string.IsNullOrEmpty(state.Config.AccessToken)
                ? "토큰 없음" : $"토큰 로드됨 ({state.ConfigPath})";

            if (string.IsNullOrWhiteSpace(SettingsTokenPath.Text))
                SettingsTokenPath.Text = state.ConfigPath;
        }
        catch
        {
            SettingsMallId.Text = "";
            SettingsTokenStatus.Text = "토큰 파일을 찾을 수 없습니다.";
        }
    }

    // ─── 준비몰(B마켓) 토큰 관련 ───────────────────────────────────────────

    private void LoadTokenInfoB()
    {
        try
        {
            var store = new Cafe24ConfigStore(_v3Root, _legacyRoot);
            var path = string.IsNullOrWhiteSpace(_bMarketTokenPath) ? null : _bMarketTokenPath;
            var state = store.LoadTokenStateB(path);
            SettingsBMallId.Text = state.Config.MallId;
            SettingsBTokenStatus.Text = string.IsNullOrEmpty(state.Config.AccessToken)
                ? "토큰 없음" : $"토큰 로드됨 ({Path.GetFileName(state.ConfigPath)})";
            if (string.IsNullOrWhiteSpace(SettingsBTokenPath.Text))
                SettingsBTokenPath.Text = state.ConfigPath;
        }
        catch
        {
            SettingsBMallId.Text = "";
            SettingsBTokenStatus.Text = "토큰 파일을 찾을 수 없습니다.";
        }
    }

    private void BrowseTokenPathB_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog
        {
            Filter = "JSON|*.json|모든 파일|*.*",
            Title = "준비몰 토큰 JSON 파일 선택",
            InitialDirectory = DesktopKeyStore.DirectoryPath
        };
        if (dlg.ShowDialog() == true)
        {
            _bMarketTokenPath = dlg.FileName;
            SettingsBTokenPath.Text = dlg.FileName;
            LoadTokenInfoB();
            // 설정 저장
            var s = BuildListingSettings();
            SaveAppSettings(s);
            Log($"준비몰 토큰 파일 변경: {Path.GetFileName(dlg.FileName)}");
        }
    }

    private async void CheckTokenA_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            StatusText.Text = "토큰 확인 중...";
            var store = new Cafe24ConfigStore(_v3Root, _legacyRoot);
            var tokenPath = string.IsNullOrWhiteSpace(SettingsTokenPath.Text) ? null : SettingsTokenPath.Text.Trim();
            var state = store.LoadTokenState(tokenPath);
            var client = new Cafe24ApiClient();
            await client.CheckTokenAsync(state.Config, CancellationToken.None);
            SettingsTokenStatus.Text = $"사용 가능 ({DateTime.Now:HH:mm:ss})";
            Log("홈런마켓 토큰 확인 완료 — 정상");
            MessageBox.Show("홈런마켓 토큰이 정상입니다.", "토큰 확인", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            SettingsTokenStatus.Text = "토큰 오류";
            Log($"홈런마켓 토큰 확인 실패: {ex.Message}");
            MessageBox.Show($"토큰이 유효하지 않습니다. 토큰 파일을 교체해 주세요.\n\n{ex.Message}", "토큰 오류", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
        finally { StatusText.Text = "대기 중"; }
    }

    private async void CheckTokenB_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            StatusText.Text = "준비몰 토큰 확인 중...";
            var store = new Cafe24ConfigStore(_v3Root, _legacyRoot);
            var path = string.IsNullOrWhiteSpace(_bMarketTokenPath) ? null : _bMarketTokenPath;
            var state = store.LoadTokenStateB(path);
            var client = new Cafe24ApiClient();
            await client.CheckTokenAsync(state.Config, CancellationToken.None);
            SettingsBTokenStatus.Text = $"사용 가능 ({DateTime.Now:HH:mm:ss})";
            Log("준비몰 토큰 확인 완료 — 정상");
            MessageBox.Show("준비몰 토큰이 정상입니다.", "준비몰 토큰 확인", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            SettingsBTokenStatus.Text = "토큰 오류";
            Log($"준비몰 토큰 확인 실패: {ex.Message}");
            MessageBox.Show($"준비몰 토큰이 유효하지 않습니다. 토큰 파일을 교체해 주세요.\n\n{ex.Message}", "준비몰 토큰 오류", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
        finally { StatusText.Text = "대기 중"; }
    }

    #endregion

    #region ═══ 파일 열기 ═══

    private void OpenUploadExcel_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_lastOutputRoot)) return;
        var uploadFile = FindLatestFile(_lastOutputRoot, "업로드용_*.xlsx");
        if (uploadFile != null && File.Exists(uploadFile))
        {
            Process.Start(new ProcessStartInfo(uploadFile) { UseShellExecute = true });
            Log($"엑셀 열기: {Path.GetFileName(uploadFile)}");
        }
        else
            MessageBox.Show("업로드용 엑셀을 찾을 수 없습니다.", "알림", MessageBoxButton.OK, MessageBoxImage.Warning);
    }

    private void OpenOutputFolder_Click(object sender, RoutedEventArgs e)
    {
        if (!string.IsNullOrEmpty(_lastOutputRoot) && Directory.Exists(_lastOutputRoot))
            Process.Start(new ProcessStartInfo("explorer.exe", _lastOutputRoot));
    }

    #endregion

    #region ═══ 유틸 ═══

    private bool ValidateSource()
    {
        if (string.IsNullOrEmpty(_sourcePath) || !File.Exists(_sourcePath))
        {
            MessageBox.Show("파일을 먼저 선택하세요.", "알림", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }
        return true;
    }

    // ── 상품 선택 목록 ────────────────────────────────────────────────

    private void LoadCafe24ProductList(string? xlsxPath = null)
    {
        xlsxPath ??= (!string.IsNullOrEmpty(_lastOutputFile) && File.Exists(_lastOutputFile))
            ? _lastOutputFile
            : FindLatestFile(_lastOutputRoot, "업로드용_*.xlsx");

        _cafe24Items.Clear();
        _cafe24LastClickIndex = -1;

        if (string.IsNullOrEmpty(xlsxPath) || !File.Exists(xlsxPath))
        {
            Cafe24SelectCountText.Text = "(업로드용 엑셀 없음 — STEP 1 먼저 실행)";
            return;
        }

        try
        {
            var entries = Services.Cafe24CreateProductService.ExtractGsCodesFromWorkbook(xlsxPath);
            foreach (var (gsCode, productName) in entries)
            {
                var hist = _uploadHistory.Get(gsCode);
                _cafe24Items.Add(new UploadProductItem
                {
                    GsCode = gsCode,
                    ProductName = productName,
                    HomeMarketStatus = UploadProductItem.FormatDate(hist?.HomeMarket),
                    ReadyMarketStatus = UploadProductItem.FormatDate(hist?.ReadyMarket),
                });
            }
            UpdateCafe24SelectCount();
        }
        catch (Exception ex)
        {
            Cafe24SelectCountText.Text = $"읽기 실패: {ex.Message}";
        }
    }

    private void LoadCoupangProductList(string? xlsxPath = null)
    {
        xlsxPath ??= CoupangSourcePath?.Text?.Trim();
        if (string.IsNullOrEmpty(xlsxPath) && !string.IsNullOrEmpty(_lastOutputRoot))
            xlsxPath = FindLatestFile(_lastOutputRoot, "업로드용_*.xlsx");

        _coupangItems.Clear();
        _coupangLastClickIndex = -1;

        if (string.IsNullOrEmpty(xlsxPath) || !File.Exists(xlsxPath))
        {
            CoupangSelectCountText.Text = "(파일 없음)";
            return;
        }

        try
        {
            var rows = Services.CoupangProductBuilder.ReadSourceFile(xlsxPath);
            var gsRegex = new System.Text.RegularExpressions.Regex(@"(GS\d{7}[A-Z0-9]*)",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            foreach (var row in rows)
            {
                var rowNum = (int)row["_row_num"]!;
                var name = row.TryGetValue("상품명", out var n) ? n?.ToString() ?? "" : "";
                var codeField = row.TryGetValue("자체 상품코드", out var c) ? c?.ToString() ?? "" : "";
                var gsCode = gsRegex.Match(codeField).Success ? gsRegex.Match(codeField).Groups[1].Value.ToUpperInvariant()
                           : gsRegex.Match(name).Success ? gsRegex.Match(name).Groups[1].Value.ToUpperInvariant() : "";
                var hist = string.IsNullOrEmpty(gsCode) ? null : _uploadHistory.Get(gsCode);

                _coupangItems.Add(new UploadProductItem
                {
                    RowNum = rowNum,
                    GsCode = gsCode,
                    ProductName = name,
                    CoupangStatus = UploadProductItem.FormatDate(hist?.Coupang),
                });
            }
            UpdateCoupangSelectCount();
        }
        catch (Exception ex)
        {
            CoupangSelectCountText.Text = $"읽기 실패: {ex.Message}";
        }
    }

    private void UpdateCafe24SelectCount()
    {
        var total = _cafe24Items.Count;
        var selected = _cafe24Items.Count(i => i.IsChecked);
        Cafe24SelectCountText.Text = $"{selected}/{total} 선택";
    }

    private void UpdateCoupangSelectCount()
    {
        var total = _coupangItems.Count;
        var selected = _coupangItems.Count(i => i.IsChecked);
        CoupangSelectCountText.Text = $"{selected}/{total} 선택";
    }

    private void Cafe24SelectAll_Click(object sender, RoutedEventArgs e)
    {
        foreach (var item in _cafe24Items) item.IsChecked = true;
        UpdateCafe24SelectCount();
    }

    private void Cafe24DeselectAll_Click(object sender, RoutedEventArgs e)
    {
        foreach (var item in _cafe24Items) item.IsChecked = false;
        UpdateCafe24SelectCount();
    }

    private void Cafe24RefreshList_Click(object sender, RoutedEventArgs e) => LoadCafe24ProductList();

    private void CoupangSelectAll_Click(object sender, RoutedEventArgs e)
    {
        foreach (var item in _coupangItems) item.IsChecked = true;
        UpdateCoupangSelectCount();
    }

    private void CoupangDeselectAll_Click(object sender, RoutedEventArgs e)
    {
        foreach (var item in _coupangItems) item.IsChecked = false;
        UpdateCoupangSelectCount();
    }

    private void CoupangRefreshList_Click(object sender, RoutedEventArgs e) => LoadCoupangProductList();

    private void Cafe24ProductList_PreviewMouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        HandleProductListShiftClick(Cafe24ProductList, _cafe24Items, ref _cafe24LastClickIndex, e);
        UpdateCafe24SelectCount();
    }

    private void CoupangProductList_PreviewMouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        HandleProductListShiftClick(CoupangProductList, _coupangItems, ref _coupangLastClickIndex, e);
        UpdateCoupangSelectCount();
    }

    // ── 기본실행 탭 신규등록 목록 ─────────────────────────────────────────

    private void LoadBasicCafe24ProductList(string? xlsxPath = null)
    {
        xlsxPath ??= (!string.IsNullOrEmpty(_lastOutputFile) && File.Exists(_lastOutputFile))
            ? _lastOutputFile
            : FindLatestFile(_lastOutputRoot, "업로드용_*.xlsx");

        _basicCafe24Items.Clear();
        _basicCafe24LastClickIndex = -1;

        if (string.IsNullOrEmpty(xlsxPath) || !File.Exists(xlsxPath))
        {
            BasicCafe24CountText.Text = "(업로드용 엑셀 없음 — STEP 1 먼저 실행)";
            BasicCafe24RunButton.IsEnabled = false;
            return;
        }

        try
        {
            var entries = Services.Cafe24CreateProductService.ExtractGsCodesFromWorkbook(xlsxPath);
            foreach (var (gsCode, productName) in entries)
            {
                var hist = _uploadHistory.Get(gsCode);
                _basicCafe24Items.Add(new UploadProductItem
                {
                    GsCode = gsCode,
                    ProductName = productName,
                    HomeMarketStatus = UploadProductItem.FormatDate(hist?.HomeMarket),
                    ReadyMarketStatus = UploadProductItem.FormatDate(hist?.ReadyMarket),
                });
            }
            UpdateBasicCafe24Count();
            BasicCafe24RunButton.IsEnabled = _basicCafe24Items.Count > 0;
        }
        catch (Exception ex)
        {
            BasicCafe24CountText.Text = $"읽기 실패: {ex.Message}";
            BasicCafe24RunButton.IsEnabled = false;
        }
    }

    private void UpdateBasicCafe24Count()
    {
        var total = _basicCafe24Items.Count;
        var selected = _basicCafe24Items.Count(i => i.IsChecked);
        BasicCafe24CountText.Text = $"{selected}/{total} 선택";
    }

    private void BasicCafe24SelectAll_Click(object sender, RoutedEventArgs e)
    {
        foreach (var item in _basicCafe24Items) item.IsChecked = true;
        UpdateBasicCafe24Count();
    }

    private void BasicCafe24DeselectAll_Click(object sender, RoutedEventArgs e)
    {
        foreach (var item in _basicCafe24Items) item.IsChecked = false;
        UpdateBasicCafe24Count();
    }

    private void BasicCafe24Refresh_Click(object sender, RoutedEventArgs e)
    {
        var files = _testLlmResultFiles.Where(File.Exists).ToList();
        LoadBasicCafe24ProductList(files.Count > 0 ? files[0] : null);
    }

    private void BasicCafe24ProductList_PreviewMouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        HandleProductListShiftClick(BasicCafe24ProductGrid, _basicCafe24Items, ref _basicCafe24LastClickIndex, e);
        UpdateBasicCafe24Count();
    }

    private async void BasicCafe24Run_Click(object sender, RoutedEventArgs e)
    {
        var files = _testLlmResultFiles.Where(File.Exists).ToList();
        if (files.Count == 0)
        {
            MessageBox.Show("LLM 결과 파일을 먼저 선택하세요.", "알림", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var selectedGs = new HashSet<string>(
            _basicCafe24Items.Where(i => i.IsChecked).Select(i => i.GsCode),
            StringComparer.OrdinalIgnoreCase);

        if (selectedGs.Count == 0)
        {
            MessageBox.Show("등록할 상품을 선택하세요.", "알림", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var doHome = TestCafe24HomeCheckBox.IsChecked == true;
        var doReady = TestCafe24ReadyCheckBox.IsChecked == true;

        _lastOutputRoot = _testOutputRoot ?? Path.GetDirectoryName(files[0])!;
        _lastOutputFile = files[0];

        var uploadFile = FindUploadExcel();
        if (uploadFile == null)
        {
            MessageBox.Show("업로드용 엑셀을 찾을 수 없습니다. STEP 1을 먼저 실행하세요.", "알림", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // 네이버 중복 확인
        if (doHome)
        {
            StatusText.Text = "네이버 중복 확인 중...";
            var duplicateInfo = await CheckNaverDuplicatesAsync(uploadFile);
            if (duplicateInfo.Count > 0)
            {
                var dupLines = duplicateInfo.Select(d => $"  • {d.GsCode}  {d.ProductName}");
                var msg = $"다음 {duplicateInfo.Count}개 상품이 네이버에 이미 등록되어 있습니다:\n\n" +
                          string.Join("\n", dupLines) +
                          "\n\n포함하여 계속 진행하시겠습니까?";
                if (MessageBox.Show(msg, "네이버 중복 확인", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
                    return;
            }
        }

        BasicCafe24RunButton.IsEnabled = false;
        _cts = new CancellationTokenSource();

        try
        {
            StatusText.Text = "카테고리맵 자동 업로드 중...";
            ProgressBar.IsIndeterminate = true;

            var createService = new Cafe24CreateProductService(_v3Root, _legacyRoot);
            var progress = new Progress<string>(msg => Log(msg));
            int totalCreated = 0, totalError = 0, totalSkipped = 0;

            await TryUploadLatestMarketPlusCategoryMapAsync(uploadFile, files, _cts.Token);
            StatusText.Text = "Cafe24 신규등록 중...";

            if (doHome)
            {
                var aTokenPath = string.IsNullOrWhiteSpace(SettingsTokenPath.Text) ? null : SettingsTokenPath.Text.Trim();
                var result = await createService.CreateAsync(uploadFile, _lastOutputRoot, progress, _cts.Token,
                    tokenPath: aTokenPath, allowedGsCodes: selectedGs);
                totalCreated += result.CreatedCount;
                totalError += result.ErrorCount;
                totalSkipped += result.SkippedCount;
                foreach (var item in _basicCafe24Items.Where(i => i.IsChecked))
                {
                    _uploadHistory.Mark(item.GsCode, "homemarket");
                    item.HomeMarketStatus = UploadProductItem.FormatDate(DateTime.Now);
                }
                Log($"[홈런마켓] 신규등록 완료: 생성 {result.CreatedCount} / 오류 {result.ErrorCount} / 스킵 {result.SkippedCount}");
            }

            if (doReady)
            {
                StatusText.Text = "준비몰 신규상품 등록 중...";
                var resultB = await createService.CreateBMarketAsync(uploadFile, _lastOutputRoot, progress, _cts.Token,
                    _bMarketTokenPath, selectedGs);
                totalCreated += resultB.CreatedCount;
                totalError += resultB.ErrorCount;
                totalSkipped += resultB.SkippedCount;
                foreach (var item in _basicCafe24Items.Where(i => i.IsChecked))
                {
                    _uploadHistory.Mark(item.GsCode, "readymarket");
                    item.ReadyMarketStatus = UploadProductItem.FormatDate(DateTime.Now);
                }
                Log($"[준비몰] 신규등록 완료: 생성 {resultB.CreatedCount} / 오류 {resultB.ErrorCount} / 스킵 {resultB.SkippedCount}");
            }

            UpdateBasicCafe24Count();
            Log($"신규등록 완료: 생성 {totalCreated} / 오류 {totalError} / 스킵 {totalSkipped}");
            StatusText.Text = "신규등록 완료";
        }
        catch (OperationCanceledException)
        {
            Log("신규등록 취소됨");
            StatusText.Text = "취소됨";
        }
        catch (Exception ex)
        {
            Log($"신규등록 오류: {ex.Message}");
            StatusText.Text = "오류 발생";
            MessageBox.Show(ex.Message, "신규등록 오류", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            ProgressBar.IsIndeterminate = false;
            BasicCafe24RunButton.IsEnabled = true;
            _cts = null;
        }
    }

    private static void HandleProductListShiftClick(
        System.Windows.Controls.DataGrid grid,
        System.Collections.ObjectModel.ObservableCollection<UploadProductItem> items,
        ref int lastIndex,
        System.Windows.Input.MouseButtonEventArgs e)
    {
        var hit = grid.InputHitTest(e.GetPosition(grid)) as System.Windows.DependencyObject;
        if (hit is null) return;

        // 클릭한 DataGridRow 찾기
        while (hit is not null && hit is not System.Windows.Controls.DataGridRow)
            hit = System.Windows.Media.VisualTreeHelper.GetParent(hit);

        if (hit is not System.Windows.Controls.DataGridRow clickedRow) return;

        var item = clickedRow.Item as UploadProductItem;
        if (item is null) return;

        var clickedIndex = items.IndexOf(item);
        if (clickedIndex < 0) return;

        if (Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift))
        {
            if (lastIndex >= 0)
            {
                var start = Math.Min(lastIndex, clickedIndex);
                var end = Math.Max(lastIndex, clickedIndex);
                var targetState = items[clickedIndex].IsChecked;
                for (var i = start; i <= end; i++)
                    items[i].IsChecked = targetState;
                e.Handled = true;
                return;
            }
        }

        lastIndex = clickedIndex;
    }

    private static string? FindLatestFile(string? dir, string pattern)
    {
        if (string.IsNullOrEmpty(dir) || !Directory.Exists(dir)) return null;
        return Directory.GetFiles(dir, pattern)
            .OrderByDescending(File.GetLastWriteTime)
            .FirstOrDefault();
    }

    private async Task<List<(string GsCode, string ProductName)>> CheckNaverDuplicatesAsync(string uploadFile)
    {
        var result = new List<(string GsCode, string ProductName)>();
        try
        {
            var gsCodesInFile = Cafe24CreateProductService.ExtractGsCodesFromWorkbook(uploadFile);
            if (gsCodesInFile.Count == 0) return result;

            var naverClient = NaverCommerceApiClient.FromKeyFile();
            var existingCodes = await naverClient.GetExistingGsCodesAsync(CancellationToken.None);
            var existingSet = new HashSet<string>(existingCodes.Select(e => e.GsCode), StringComparer.OrdinalIgnoreCase);

            foreach (var (gsCode, productName) in gsCodesInFile)
            {
                if (existingSet.Contains(gsCode))
                    result.Add((gsCode, productName));
            }
        }
        catch (Exception ex)
        {
            Log($"네이버 중복 확인 실패 (스킵): {ex.Message}");
        }
        return result;
    }

    private static int FindCol(Dictionary<string, int> cols, string[] candidates)
    {
        foreach (var c in candidates)
            if (cols.TryGetValue(c, out var idx)) return idx;
        return -1;
    }

    private static decimal GetDecimal(IXLCell cell)
    {
        try { return (decimal)cell.GetDouble(); }
        catch { return decimal.TryParse(cell.GetString(), out var v) ? v : 0; }
    }

    private static int ParseInt(TextBox tb, int fallback)
    {
        return int.TryParse(tb.Text.Trim(), out var v) ? v : fallback;
    }

    private static double ParseDouble(TextBox tb, double fallback)
    {
        return double.TryParse(tb.Text.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v : fallback;
    }

    private void Log(string message)
    {
        var time = DateTime.Now.ToString("HH:mm:ss");
        Dispatcher.Invoke(() =>
        {
            LogBlock.AppendText($"[{time}] {message}\n");
            LogBlock.ScrollToEnd();
        });
    }

    private void ClearLog_Click(object sender, RoutedEventArgs e)
    {
        LogBlock.Text = "";
    }

    private void SetPipelineEnabled(bool enabled)
    {
        RunPipelineButton.IsEnabled = enabled;
        RunKeywordOnlyButton.IsEnabled = enabled;
        RunListingOnlyButton.IsEnabled = enabled;
        TestRunOcrOnlyButton.IsEnabled = enabled;
    }

    private void SetRunning(bool running)
    {
        SetPipelineEnabled(!running);
        RunPipelineButton.Content = running ? "실행 중..." : "전체 실행 (전처리+OCR+키워드+이미지)";
    }

    protected override void OnClosed(EventArgs e)
    {
        _cts?.Cancel();
        base.OnClosed(e);
    }

    #endregion

    #region ═══ 이미지 선택 ═══

    /// <summary>Phase 1 완료 후 이미지 로드 + 탭 전환</summary>
    private void LoadListingImagesFromRoot(string outputRoot)
    {
        _lastOutputRoot = outputRoot;
        LoadImages_Click(this, new RoutedEventArgs());
        ImageSelectionTab.Visibility = Visibility.Visible;
        ImageSelectionTab.IsSelected = true;
    }

    private void LoadImages_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_lastOutputRoot) || !Directory.Exists(_lastOutputRoot))
        {
            MessageBox.Show("먼저 파이프라인을 실행하거나 이력을 불러오세요.", "알림",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var listingRoot = Path.Combine(_lastOutputRoot, "listing_images");
        if (!Directory.Exists(listingRoot))
        {
            Log("listing_images 폴더를 찾을 수 없습니다.");
            return;
        }

        var dateDir = Directory.GetDirectories(listingRoot)
            .OrderByDescending(d => d).FirstOrDefault();
        if (dateDir == null)
        {
            Log("날짜 폴더를 찾을 수 없습니다.");
            return;
        }

        _imageListingRoot = dateDir;
        _imageGsCodes.Clear();
        _imageSelections.Clear();
        _imageThumbnails.Clear();

        // 기존 선택 불러오기
        var selectionsPath = Path.Combine(_lastOutputRoot, "image_selections.json");
        if (File.Exists(selectionsPath))
            LoadImageSelectionsFromJson(selectionsPath);

        var gsFolders = Cafe24UploadSupport.GetGsFolders(dateDir);
        foreach (var folder in gsFolders)
        {
            var gs = folder.Name.ToUpperInvariant();
            var gs9 = gs.Length >= 9 ? gs[..9] : gs;
            _imageGsCodes.Add(gs9);

            if (!_imageSelections.ContainsKey(gs9))
            {
                var fileCount = Directory.GetFiles(folder.FullName)
                    .Count(f => IsImageFile(f));
                var mainIdx = fileCount >= 2 ? 1 : 0;
                var addIndices = Enumerable.Range(2, Math.Max(0, fileCount - 2)).ToList();
                _imageSelections[gs9] = new ImageSelection(mainIdx, addIndices);
            }
        }

        // 파일 정보 표시
        var sourceName = _lastOutputFile != null ? Path.GetFileName(_lastOutputFile) : Path.GetFileName(_lastOutputRoot ?? "");
        var dateFolderName = Path.GetFileName(dateDir);
        ImageSourceFileText.Text = sourceName;
        ImageSourceDateText.Text = dateFolderName;
        ImageGsCountText.Text = $"{_imageGsCodes.Count}개";
        ImageSelectionStatus.Text = $"{_imageGsCodes.Count}개 상품";
        Log($"이미지 불러오기 완료: {_imageGsCodes.Count}개 상품 ({dateFolderName})");

        if (_imageGsCodes.Count > 0)
            ImageGsListBox.SelectedIndex = 0;
    }

    private void ImageGsListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        _imageThumbnails.Clear();
        _selectingBMarket = false;
        ImageSelectionStatus.Text = $"{_imageGsCodes.Count}개 상품 — A마켓 대표 선택";
        if (ImageGsListBox.SelectedItem is not string gs || _imageListingRoot == null) return;

        var folder = Path.Combine(_imageListingRoot, gs);
        if (!Directory.Exists(folder)) return;

        var files = Directory.GetFiles(folder)
            .Where(f => IsImageFile(f))
            .OrderBy(f => f)
            .ToList();

        _imageSelections.TryGetValue(gs, out var selection);

        for (int i = 0; i < files.Count; i++)
        {
            var item = new ImageThumbnailItem
            {
                Index = i,
                DisplayNumber = i + 1,
                FilePath = files[i],
                Thumbnail = LoadThumbnail(files[i]),
                IsMain = selection?.MainIndex == i,
                IsMainB = selection?.MainIndexB == i,
                IsAdditional = selection?.AdditionalIndices?.Contains(i) == true,
            };
            _imageThumbnails.Add(item);
        }
    }

    private void Thumbnail_LeftClick(object sender, MouseButtonEventArgs e)
    {
        if (sender is not FrameworkElement fe || fe.DataContext is not ImageThumbnailItem clicked) return;
        if (ImageGsListBox.SelectedItem is not string gs) return;

        if (_selectingBMarket)
        {
            // B마켓 대표이미지 선택
            foreach (var item in _imageThumbnails)
                item.IsMainB = false;
            clicked.IsMainB = true;
            clicked.IsAdditional = false;
            UpdateSelectionForGs(gs);

            // B마켓 선택 완료 → 다음 상품으로 이동
            _selectingBMarket = false;
            ImageSelectionStatus.Text = $"{_imageGsCodes.Count}개 상품 — A마켓 대표 선택";
            var currentIndex = ImageGsListBox.SelectedIndex;
            if (currentIndex < ImageGsListBox.Items.Count - 1)
            {
                ImageGsListBox.SelectedIndex = currentIndex + 1;
                ImageGsListBox.ScrollIntoView(ImageGsListBox.SelectedItem);
            }
            e.Handled = true;
            return;
        }

        // A마켓 대표이미지 선택
        foreach (var item in _imageThumbnails)
            item.IsMain = false;
        clicked.IsMain = true;
        clicked.IsAdditional = false;
        UpdateSelectionForGs(gs);

        // 더블클릭이면 B마켓 선택 모드로 전환
        if (e.ClickCount >= 2)
        {
            _selectingBMarket = true;
            ImageSelectionStatus.Text = $"{_imageGsCodes.Count}개 상품 — B마켓 대표 선택 (클릭하세요)";
            e.Handled = true;
        }
    }

    private void Thumbnail_RightClick(object sender, MouseButtonEventArgs e)
    {
        if (sender is not FrameworkElement fe || fe.DataContext is not ImageThumbnailItem clicked) return;
        if (ImageGsListBox.SelectedItem is not string gs) return;

        if (clicked.IsMain) return;
        clicked.IsAdditional = !clicked.IsAdditional;
        UpdateSelectionForGs(gs);
        e.Handled = true;
    }

    private void UpdateSelectionForGs(string gs)
    {
        var mainItem = _imageThumbnails.FirstOrDefault(t => t.IsMain);
        var mainBItem = _imageThumbnails.FirstOrDefault(t => t.IsMainB);
        var addItems = _imageThumbnails.Where(t => t.IsAdditional).Select(t => t.Index).ToList();
        _imageSelections[gs] = new ImageSelection(mainItem?.Index, addItems, mainBItem?.Index);
    }

    private void SaveImageSelection_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_lastOutputRoot))
        {
            MessageBox.Show("저장할 대상이 없습니다.", "알림"); return;
        }

        var dict = new Dictionary<string, object>();
        foreach (var kvp in _imageSelections)
        {
            dict[kvp.Key] = new { main = kvp.Value.MainIndex, mainB = kvp.Value.MainIndexB, additional = kvp.Value.AdditionalIndices };
        }

        var json = JsonSerializer.Serialize(dict, new JsonSerializerOptions { WriteIndented = true });
        var path = Path.Combine(_lastOutputRoot, "image_selections.json");
        File.WriteAllText(path, json, Encoding.UTF8);
        Log($"이미지 선택 저장 완료: {_imageSelections.Count}개 상품");

        // 실행이력에 이미지 선택 완료 표시
        if (_jobHistory != null)
        {
            var job = _jobHistory.Records.FirstOrDefault(r => r.OutputRoot == _lastOutputRoot);
            if (job != null && !job.ImageSelected)
            {
                job.ImageSelected = true;
                _jobHistory.Update(job);
                RefreshHistoryGrid();
            }
        }

        BasicRunTab.IsSelected = true;
        StatusText.Text = "이미지 선택 저장 완료";
    }

    private void LoadImageSelectionsFromJson(string path)
    {
        try
        {
            var json = File.ReadAllText(path, Encoding.UTF8);
            using var doc = JsonDocument.Parse(json);
            foreach (var prop in doc.RootElement.EnumerateObject())
            {
                int? mainIdx = null;
                var addIndices = new List<int>();

                if (prop.Value.TryGetProperty("main", out var mainEl) && mainEl.ValueKind == JsonValueKind.Number)
                    mainIdx = mainEl.GetInt32();
                int? mainIdxB = null;
                if (prop.Value.TryGetProperty("mainB", out var mainBEl) && mainBEl.ValueKind == JsonValueKind.Number)
                    mainIdxB = mainBEl.GetInt32();
                if (prop.Value.TryGetProperty("additional", out var addEl) && addEl.ValueKind == JsonValueKind.Array)
                {
                    foreach (var item in addEl.EnumerateArray())
                        if (item.ValueKind == JsonValueKind.Number) addIndices.Add(item.GetInt32());
                }

                _imageSelections[prop.Name] = new ImageSelection(mainIdx, addIndices, mainIdxB);
            }
            Log($"이미지 선택 불러옴: {_imageSelections.Count}개");
        }
        catch (Exception ex)
        {
            Log($"이미지 선택 로드 실패: {ex.Message}");
        }
    }

    private static BitmapImage LoadThumbnail(string path)
    {
        var bmp = new BitmapImage();
        bmp.BeginInit();
        bmp.UriSource = new Uri(path);
        bmp.DecodePixelWidth = 150;
        bmp.CacheOption = BitmapCacheOption.OnLoad;
        bmp.EndInit();
        bmp.Freeze();
        return bmp;
    }

    private static bool IsImageFile(string path)
    {
        var ext = Path.GetExtension(path).ToLowerInvariant();
        return ext is ".jpg" or ".jpeg" or ".png" or ".webp" or ".bmp";
    }

    #endregion
}

#region ═══ 데이터 모델 ═══

public class ProductItem : INotifyPropertyChanged
{
    private bool _isSelected = true;
    public string Code { get; set; } = "";
    public string Name { get; set; } = "";
    public DateTime? LastProcessedAt { get; set; }

    public bool IsSelected
    {
        get => _isSelected;
        set { if (_isSelected == value) return; _isSelected = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsSelected))); }
    }

    // 이력이 있으면 "(MM/dd HH:mm)" 형태로 표시, 없으면 ""
    public string HistoryText => LastProcessedAt.HasValue
        ? LastProcessedAt.Value.ToString("(MM/dd HH:mm)")
        : "";

    public event PropertyChangedEventHandler? PropertyChanged;
}

public class PriceRow : INotifyPropertyChanged
{
    private bool _isChecked = true;
    private decimal _additionalAmount;

    public bool IsChecked
    {
        get => _isChecked;
        set { if (_isChecked == value) return; _isChecked = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsChecked))); }
    }
    public string GsCode { get; set; } = "";
    public string OptionName { get; set; } = "";
    public decimal SupplyPrice { get; set; }
    public decimal SellingPrice { get; set; }
    public decimal AdditionalAmount
    {
        get => _additionalAmount;
        set { if (_additionalAmount == value) return; _additionalAmount = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(AdditionalAmount))); }
    }
    public decimal ConsumerPrice { get; set; }
    public event PropertyChangedEventHandler? PropertyChanged;
}

public class UploadResultRow
{
    public string Gs { get; set; } = "";
    public string ProductNo { get; set; } = "";
    public string Status { get; set; } = "";
    public string MainImage { get; set; } = "";
    public string AddCount { get; set; } = "";
    public string PriceStatus { get; set; } = "";
}

public class ImageThumbnailItem : INotifyPropertyChanged
{
    private bool _isMain;
    private bool _isAdditional;
    private bool _isMainB;

    public int Index { get; set; }
    public int DisplayNumber { get; set; }
    public string FilePath { get; set; } = "";
    public BitmapImage? Thumbnail { get; set; }

    public bool IsMain
    {
        get => _isMain;
        set
        {
            if (_isMain == value) return;
            _isMain = value;
            OnPropertyChanged(nameof(IsMain));
            OnPropertyChanged(nameof(StatusText));
            OnPropertyChanged(nameof(StatusColor));
        }
    }

    public bool IsAdditional
    {
        get => _isAdditional;
        set
        {
            if (_isAdditional == value) return;
            _isAdditional = value;
            OnPropertyChanged(nameof(IsAdditional));
            OnPropertyChanged(nameof(StatusText));
            OnPropertyChanged(nameof(StatusColor));
        }
    }

    public bool IsMainB
    {
        get => _isMainB;
        set
        {
            if (_isMainB == value) return;
            _isMainB = value;
            OnPropertyChanged(nameof(IsMainB));
            OnPropertyChanged(nameof(StatusText));
            OnPropertyChanged(nameof(StatusColor));
        }
    }

    public string StatusText => IsMain && IsMainB ? $"#{DisplayNumber} A+B대표"
        : IsMain ? $"#{DisplayNumber} A대표"
        : IsMainB ? $"#{DisplayNumber} B대표"
        : IsAdditional ? $"#{DisplayNumber} 추가"
        : $"#{DisplayNumber}";
    public string StatusColor => IsMain || IsMainB ? "#2196F3" : IsAdditional ? "#4CAF50" : "#888";

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

#endregion
