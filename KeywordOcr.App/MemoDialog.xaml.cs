using System.Windows;

namespace KeywordOcr.App;

public partial class MemoDialog : Window
{
    public string MemoText { get; private set; } = "";

    public MemoDialog(string currentMemo)
    {
        InitializeComponent();
        MemoBox.Text = currentMemo;
        MemoText = currentMemo;
        MemoBox.Focus();
        MemoBox.SelectAll();
    }

    private void Save_Click(object sender, RoutedEventArgs e)
    {
        MemoText = MemoBox.Text;
        DialogResult = true;
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }
}
