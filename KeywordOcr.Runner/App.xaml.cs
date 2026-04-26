using System.Windows;

namespace KeywordOcr.Runner;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        base.OnStartup(e);
    }
}
