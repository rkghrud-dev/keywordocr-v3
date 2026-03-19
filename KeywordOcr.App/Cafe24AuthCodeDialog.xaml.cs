using System;
using System.Diagnostics;
using System.Windows;

namespace KeywordOcr.App;

public partial class Cafe24AuthCodeDialog : Window
{
    private readonly string _authorizeUrl;

    public string InputText { get; private set; } = string.Empty;

    public Cafe24AuthCodeDialog(string authorizeUrl)
    {
        InitializeComponent();
        _authorizeUrl = authorizeUrl;
        AuthorizeUrlBox.Text = authorizeUrl;
        Loaded += (_, _) => CodeBox.Focus();
    }

    private void CopyUrl_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            Clipboard.SetText(_authorizeUrl);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "URL 복사 오류", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OpenBrowser_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            Process.Start(new ProcessStartInfo(_authorizeUrl) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "브라우저 열기 오류", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void Confirm_Click(object sender, RoutedEventArgs e)
    {
        InputText = CodeBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(InputText))
        {
            MessageBox.Show("callback URL 또는 code 값을 입력해 주세요.", "Cafe24 다시 인증", MessageBoxButton.OK, MessageBoxImage.Warning);
            CodeBox.Focus();
            return;
        }

        DialogResult = true;
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }
}