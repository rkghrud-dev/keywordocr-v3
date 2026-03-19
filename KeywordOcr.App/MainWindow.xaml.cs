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
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using ClosedXML.Excel;
using KeywordOcr.App.Services;
using Microsoft.Win32;

namespace KeywordOcr.App;

public partial class MainWindow : Window
{
    private readonly string _legacyRoot;
    private readonly string _v3Root;
    private string? _sourcePath;
    private string? _lastOutputRoot;
    private string? _lastOutputFile;
    private string? _lastUploadLogPath;
    private CancellationTokenSource? _cts;
    private readonly ObservableCollection<ProductItem> _products = new();
    private readonly ObservableCollection<PriceRow> _priceRows = new();
    private JobHistoryService? _jobHistory;

    public MainWindow()
    {
        InitializeComponent();

        _v3Root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
        _legacyRoot = Path.GetFullPath(Path.Combine(_v3Root, ".."));

        if (!Directory.Exists(Path.Combine(_legacyRoot, "app")))
            _legacyRoot = @"C:\Users\rkghr\Desktop\프로젝트\keywordocr";

        ProductList.ItemsSource = _products;
        PriceGrid.ItemsSource = _priceRows;

        // 설정 탭 초기값
        SettingsLegacyRoot.Text = _legacyRoot;
        SettingsV3Root.Text = _v3Root;
        Cafe24DateTag.Text = DateTime.Now.ToString("yyyyMMdd");
        LoadTokenInfo();

        _jobHistory = new JobHistoryService(_legacyRoot);
        RefreshHistoryGrid();

        Log("KeywordOCR v3 시작");
        Log($"Python 루트: {_legacyRoot}");
    }

    #region ═══ 드래그 앤 드롭 ═══

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
                SetPipelineEnabled(true);
                return;
            }

            foreach (var (code, name) in items)
                _products.Add(new ProductItem { Code = code, Name = name, IsSelected = true });

            ProductListPanel.Visibility = Visibility.Visible;
            UpdateProductCount();
            SetPipelineEnabled(true);
            Log($"상품 {items.Count}개 로드됨");
        }
        catch (Exception ex)
        {
            Log($"파일 읽기 오류: {ex.Message}");
            ProductListPanel.Visibility = Visibility.Collapsed;
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

        int codeCol = -1, nameCol = -1;
        var lastCol = headerRow.LastCellUsed()?.Address.ColumnNumber ?? 0;

        for (int c = 1; c <= lastCol; c++)
        {
            var header = headerRow.Cell(c).GetString().Trim();
            if (codeCol < 0 && CodeColumns.Any(h => h.Equals(header, StringComparison.OrdinalIgnoreCase)))
                codeCol = c;
            if (nameCol < 0 && NameColumns.Any(h => h.Equals(header, StringComparison.OrdinalIgnoreCase)))
                nameCol = c;
        }

        if (codeCol < 0) return results;

        var lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
        var seen = new HashSet<string>();

        for (int r = headerRow.RowNumber() + 1; r <= lastRow; r++)
        {
            var code = ws.Cell(r, codeCol).GetString().Trim();
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

    private void UpdateProductCount()
    {
        var selected = _products.Count(p => p.IsSelected);
        ProductCountText.Text = $"({selected}/{_products.Count} 선택)";
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
        return new ListingImageSettings(
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
            FlipLeftRight: SettingsFlipLR.IsChecked == true
        );
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
            var bridge = new PythonPipelineBridgeService(_v3Root, _legacyRoot);
            var progress = new Progress<string>(msg => Log(msg));

            Log("전체 파이프라인 실행 시작...");
            StatusText.Text = "실행 중...";
            ProgressBar.IsIndeterminate = true;

            var result = await bridge.RunPipelineAsync(inputFile, settings, progress, _cts.Token);
            OnPipelineComplete(result);
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
            var bridge = new PythonPipelineBridgeService(_v3Root, _legacyRoot);
            var progress = new Progress<string>(msg => Log(msg));

            Log("키워드만 생성 시작...");
            StatusText.Text = "키워드 생성 중...";
            ProgressBar.IsIndeterminate = true;

            var result = await bridge.RunPipelineAsync(inputFile, settings, progress, _cts.Token);
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
            var bridge = new PythonPipelineBridgeService(_v3Root, _legacyRoot);
            var progress = new Progress<string>(msg => Log(msg));

            Log("대표이미지만 생성 시작...");
            StatusText.Text = "대표이미지 생성 중...";
            ProgressBar.IsIndeterminate = true;

            var result = await bridge.RunPipelineAsync(inputFile, settings, progress, _cts.Token);
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
    }

    private void HandlePipelineError(Exception ex)
    {
        Log($"오류: {ex.Message}");
        StatusText.Text = "오류 발생";
        MessageBox.Show(ex.Message, "파이프라인 오류", MessageBoxButton.OK, MessageBoxImage.Error);
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

        var uploadFile = FindLatestFile(_lastOutputRoot, "업로드용_*.xlsx");
        if (uploadFile == null)
        {
            MessageBox.Show("업로드용 엑셀 파일을 찾을 수 없습니다.", "알림", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var confirm = MessageBox.Show(
            $"Cafe24에 이미지 업로드 + 옵션가격을 반영합니다.\n\n" +
            $"업로드 파일: {Path.GetFileName(uploadFile)}\n" +
            $"결과 폴더: {_lastOutputRoot}\n\n계속하시겠습니까?",
            "Cafe24 업로드 확인", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (confirm != MessageBoxResult.Yes) return;

        // 가격 데이터 저장
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

            var result = await uploadService.UploadAsync(
                uploadFile, _lastOutputRoot, options, progress, _cts.Token);

            _lastUploadLogPath = result.LogPath;
            Log($"Cafe24 업로드 완료: 성공 {result.SuccessCount} / 오류 {result.ErrorCount} / 스킵 {result.SkippedCount}");
            StatusText.Text = $"업로드 완료 (성공: {result.SuccessCount})";
            UploadSummaryText.Text = $"총 {result.TotalCount} | 성공 {result.SuccessCount} | 오류 {result.ErrorCount} | 스킵 {result.SkippedCount}";
            OpenUploadLogButton.IsEnabled = !string.IsNullOrEmpty(result.LogPath);

            LoadUploadLog(result.LogPath);
        }
        catch (OperationCanceledException) { Log("업로드 취소됨"); StatusText.Text = "취소됨"; }
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

        var uploadFile = FindLatestFile(_lastOutputRoot, "업로드용_*.xlsx");
        if (uploadFile == null)
        {
            MessageBox.Show("업로드용 엑셀 파일을 찾을 수 없습니다.", "알림", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var confirm = MessageBox.Show(
            $"Cafe24에 신규 상품을 등록합니다.\n\n" +
            $"업로드 파일: {Path.GetFileName(uploadFile)}\n\n계속하시겠습니까?",
            "신규상품 등록 확인", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (confirm != MessageBoxResult.Yes) return;

        Cafe24CreateButton.IsEnabled = false;
        _cts = new CancellationTokenSource();

        try
        {
            StatusText.Text = "신규상품 등록 중...";
            ProgressBar.IsIndeterminate = true;

            var options = BuildUploadOptions();
            var createService = new Cafe24CreateProductService(_v3Root, _legacyRoot);
            var progress = new Progress<string>(msg => Log(msg));

            var result = await createService.CreateAsync(
                uploadFile, _lastOutputRoot, progress, _cts.Token);

            Log($"신규등록 완료: 생성 {result.CreatedCount} / 오류 {result.ErrorCount} / 스킵 {result.SkippedCount}");
            StatusText.Text = $"등록 완료 (생성: {result.CreatedCount})";
        }
        catch (OperationCanceledException) { Log("등록 취소됨"); StatusText.Text = "취소됨"; }
        catch (Exception ex)
        {
            Log($"등록 오류: {ex.Message}");
            StatusText.Text = "등록 오류";
            MessageBox.Show(ex.Message, "신규상품 등록 오류", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            Cafe24CreateButton.IsEnabled = true;
            ProgressBar.IsIndeterminate = false;
        }
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
            $"파일: {Path.GetFileName(uploadFile)}\n" +
            $"폴더: {job.OutputRoot}\n\n계속하시겠습니까?",
            "Cafe24 업로드 확인", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (confirm != MessageBoxResult.Yes) return;

        // 이력에서 결과 로드
        _lastOutputRoot = job.OutputRoot;
        _lastOutputFile = job.OutputFile;

        // Cafe24 업로드 탭으로 이동 후 업로드 실행
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

            var result = await uploadService.UploadAsync(
                uploadFile, job.OutputRoot, options, progress, _cts.Token);

            _lastUploadLogPath = result.LogPath;
            Log($"Cafe24 업로드 완료: 성공 {result.SuccessCount} / 오류 {result.ErrorCount} / 스킵 {result.SkippedCount}");
            StatusText.Text = $"업로드 완료 (성공: {result.SuccessCount})";
            UploadSummaryText.Text = $"총 {result.TotalCount} | 성공 {result.SuccessCount} | 오류 {result.ErrorCount} | 스킵 {result.SkippedCount}";
            OpenUploadLogButton.IsEnabled = !string.IsNullOrEmpty(result.LogPath);
            LoadUploadLog(result.LogPath);
        }
        catch (OperationCanceledException) { Log("업로드 취소됨"); StatusText.Text = "취소됨"; }
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

            var result = await createService.CreateAsync(
                uploadFile, job.OutputRoot, progress, _cts.Token);

            Log($"신규등록 완료: 생성 {result.CreatedCount} / 오류 {result.ErrorCount} / 스킵 {result.SkippedCount}");
            StatusText.Text = $"등록 완료 (생성: {result.CreatedCount})";
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

    #endregion

    #region ═══ 설정 탭 ═══

    private void BrowseLogoPath_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog { Filter = "이미지|*.png;*.jpg;*.jpeg|모든 파일|*.*", Title = "로고 파일 선택" };
        if (dlg.ShowDialog() == true)
            SettingsLogoPath.Text = dlg.FileName;
    }

    private void BrowseTokenPath_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog { Filter = "텍스트|*.txt|모든 파일|*.*", Title = "Cafe24 토큰 파일 선택" };
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

    private async void RefreshToken_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            StatusText.Text = "토큰 갱신 중...";
            var store = new Cafe24ConfigStore(_v3Root, _legacyRoot);
            var tokenPath = string.IsNullOrWhiteSpace(SettingsTokenPath.Text) ? null : SettingsTokenPath.Text.Trim();
            var state = store.LoadTokenState(tokenPath);

            var client = new Cafe24ApiClient();
            await client.RefreshAccessTokenAsync(state.Config, CancellationToken.None);
            store.SaveTokenConfig(state.ConfigPath, state.Config);

            SettingsTokenStatus.Text = $"갱신 완료 ({DateTime.Now:HH:mm:ss})";
            Log("Cafe24 토큰 갱신 완료");
        }
        catch (Exception ex)
        {
            Log($"토큰 갱신 오류: {ex.Message}");
            SettingsTokenStatus.Text = "갱신 실패";
            MessageBox.Show(ex.Message, "토큰 갱신 오류", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            StatusText.Text = "대기 중";
        }
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

    private static string? FindLatestFile(string? dir, string pattern)
    {
        if (string.IsNullOrEmpty(dir) || !Directory.Exists(dir)) return null;
        return Directory.GetFiles(dir, pattern)
            .OrderByDescending(File.GetLastWriteTime)
            .FirstOrDefault();
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
            LogBlock.Text += $"[{time}] {message}\n";
            LogScroller.ScrollToEnd();
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
}

#region ═══ 데이터 모델 ═══

public class ProductItem : INotifyPropertyChanged
{
    private bool _isSelected = true;
    public string Code { get; set; } = "";
    public string Name { get; set; } = "";
    public bool IsSelected
    {
        get => _isSelected;
        set { if (_isSelected == value) return; _isSelected = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsSelected))); }
    }
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

#endregion
