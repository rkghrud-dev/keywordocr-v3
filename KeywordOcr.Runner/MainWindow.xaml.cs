using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using KeywordOcr.Core.Models;
using KeywordOcr.Core.Services;
using Microsoft.Win32;

namespace KeywordOcr.Runner;

public partial class MainWindow : Window
{
    private string? _sourcePath;
    private string? _lastOutputPath;
    private CancellationTokenSource? _cts;
    private bool _running;

    // ── 생성자 ────────────────────────────────────────────────────────────────

    public MainWindow()
    {
        InitializeComponent();
        TryLoadApiKey();
        AppendLog("준비 완료. 파일을 선택하고 실행 버튼을 누르세요.");
    }

    // ── API 키 자동 로드 ──────────────────────────────────────────────────────

    private void TryLoadApiKey()
    {
        var key = AnthropicApiClient.LoadApiKey();
        if (!string.IsNullOrEmpty(key))
        {
            ApiKeyBox.Password = key;
            AppendLog($"API 키 로드됨: {key[..8]}...");
        }
    }

    // ── 파일 드롭존 ───────────────────────────────────────────────────────────

    private void DropZone_DragOver(object sender, DragEventArgs e)
    {
        e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop)
            ? DragDropEffects.Copy
            : DragDropEffects.None;
        DropZone.BorderBrush = new SolidColorBrush(Color.FromRgb(37, 99, 235));
        e.Handled = true;
    }

    private void DropZone_DragLeave(object sender, DragEventArgs e)
    {
        DropZone.BorderBrush = new SolidColorBrush(Color.FromRgb(203, 213, 225));
    }

    private void DropZone_Drop(object sender, DragEventArgs e)
    {
        DropZone.BorderBrush = new SolidColorBrush(Color.FromRgb(203, 213, 225));
        if (e.Data.GetData(DataFormats.FileDrop) is string[] files && files.Length > 0)
            SetSourceFile(files[0]);
    }

    private void DropZone_Click(object sender, MouseButtonEventArgs e)
    {
        var dlg = new OpenFileDialog
        {
            Title = "입력 파일 선택",
            Filter = "Excel/CSV|*.xlsx;*.xls;*.csv|모든 파일|*.*",
        };
        if (_sourcePath != null)
            dlg.InitialDirectory = Path.GetDirectoryName(_sourcePath) ?? "";

        if (dlg.ShowDialog() == true)
            SetSourceFile(dlg.FileName);
    }

    private void SetSourceFile(string path)
    {
        _sourcePath = path;
        DropFileName.Text = Path.GetFileName(path);
        DropFileName.Visibility = Visibility.Visible;
        DropHint.Text = "✓ 파일 선택됨";
        DropHint.Foreground = new SolidColorBrush(Color.FromRgb(22, 163, 74));
        AppendLog($"파일: {path}");
    }

    // ── 폴더 찾기 ─────────────────────────────────────────────────────────────

    private void BrowseImgFolder_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFolderDialog { Title = "이미지 폴더 선택" };
        if (dlg.ShowDialog() == true)
            ImgFolderBox.Text = dlg.FolderName;
    }

    private void BrowseTess_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFolderDialog
        {
            Title = "Tesseract 설치 폴더 선택",
            InitialDirectory = TessPathBox.Text,
        };
        if (dlg.ShowDialog() == true)
            TessPathBox.Text = dlg.FolderName;
    }

    // ── 실행 버튼 ─────────────────────────────────────────────────────────────

    private void BtnRunAll_Click(object sender, RoutedEventArgs e)
        => StartPipeline("full");

    private void BtnRunKw_Click(object sender, RoutedEventArgs e)
        => StartPipeline("analysis");

    private void BtnRunImg_Click(object sender, RoutedEventArgs e)
        => StartPipeline("images");

    private void BtnStop_Click(object sender, RoutedEventArgs e)
    {
        _cts?.Cancel();
        AppendLog("중지 요청됨...");
    }

    private void BtnClearLog_Click(object sender, RoutedEventArgs e)
        => LogBox.Text = "";

    private void BtnOpenResult_Click(object sender, RoutedEventArgs e)
    {
        if (_lastOutputPath == null) return;
        var dir = Directory.Exists(_lastOutputPath)
            ? _lastOutputPath
            : Path.GetDirectoryName(_lastOutputPath) ?? "";
        if (Directory.Exists(dir))
            Process.Start("explorer.exe", dir);
    }

    // ── 파이프라인 실행 ───────────────────────────────────────────────────────

    private async void StartPipeline(string phase)
    {
        if (_running)
        {
            MessageBox.Show("이미 실행 중입니다.", "알림", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        if (string.IsNullOrEmpty(_sourcePath) || !File.Exists(_sourcePath))
        {
            MessageBox.Show("먼저 입력 파일을 선택해주세요.", "파일 없음", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        var apiKey = ApiKeyBox.Password.Trim();
        if (string.IsNullOrEmpty(apiKey))
        {
            MessageBox.Show("Anthropic API Key를 입력해주세요.", "키 없음", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        SetRunning(true);
        _cts = new CancellationTokenSource();
        SetProgress(0, "시작...");

        try
        {
            var cfg = BuildConfig(apiKey, phase);
            using var pipeline = new PipelineService(cfg);

            var progress = new Progress<string>(msg =>
            {
                Dispatcher.Invoke(() =>
                {
                    AppendLog(msg);
                    // 간단한 진행률 추정
                    if (msg.Contains("/"))
                    {
                        var m = System.Text.RegularExpressions.Regex.Match(msg, @"\[(\d+)/(\d+)\]");
                        if (m.Success
                            && int.TryParse(m.Groups[1].Value, out int cur)
                            && int.TryParse(m.Groups[2].Value, out int tot)
                            && tot > 0)
                        {
                            SetProgress((int)(cur * 100.0 / tot), $"{cur}/{tot}");
                        }
                    }
                });
            });

            var outputPath = await pipeline.RunAsync(progress, _cts.Token);
            _lastOutputPath = outputPath;

            SetProgress(100, "완료");
            AppendLog($"✅ 완료 → {outputPath}");
            BtnOpenResult.IsEnabled = true;

            // 결과 폴더 자동 열기
            var dir = Path.GetDirectoryName(outputPath) ?? "";
            if (Directory.Exists(dir))
                Process.Start("explorer.exe", dir);
        }
        catch (OperationCanceledException)
        {
            AppendLog("⚠ 중지됨.");
            SetProgress(0, "중지");
        }
        catch (Exception ex)
        {
            AppendLog($"❌ 오류: {ex.Message}");
            MessageBox.Show(ex.Message, "오류", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            SetRunning(false);
        }
    }

    // ── 설정 조립 ─────────────────────────────────────────────────────────────

    private PipelineConfig BuildConfig(string apiKey, string phase)
    {
        var model = ((ComboBoxItem)ModelCombo.SelectedItem)?.Tag?.ToString()
                    ?? "claude-haiku-4-5-20251001";

        return new PipelineConfig
        {
            FilePath          = _sourcePath!,
            AnthropicApiKey   = apiKey,
            ModelKeyword      = model,
            ModelLongtail     = model,
            Phase             = phase,
            UseLocalOcr       = ChkOcr.IsChecked == true,
            MakeListing       = ChkListing.IsChecked == true && phase != "analysis",
            EnableBMarket     = ChkBMarket.IsChecked == true,
            LocalImgDir       = ImgFolderBox.Text.Trim(),
            TesseractPath     = TessPathBox.Text.Trim(),
            Threads           = 4,
            Debug             = true,
        };
    }

    // ── UI 헬퍼 ───────────────────────────────────────────────────────────────

    private void SetRunning(bool running)
    {
        _running = running;
        BtnRunAll.IsEnabled  = !running;
        BtnRunKw.IsEnabled   = !running;
        BtnRunImg.IsEnabled  = !running;
        BtnStop.IsEnabled    = running;
    }

    private void SetProgress(int value, string label)
    {
        ProgressBar.Value   = value;
        ProgressLabel.Text  = label;
    }

    private void AppendLog(string message)
    {
        var ts   = DateTime.Now.ToString("HH:mm:ss");
        var line = $"[{ts}] {message}\n";
        LogBox.Text += line;
        LogScroller.ScrollToEnd();
    }
}
