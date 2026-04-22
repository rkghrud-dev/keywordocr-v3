using System.ComponentModel;

namespace KeywordOcr.App;

public sealed class UploadProductItem : INotifyPropertyChanged
{
    private bool _isChecked = true;

    public event PropertyChangedEventHandler? PropertyChanged;

    public bool IsChecked
    {
        get => _isChecked;
        set
        {
            if (_isChecked != value)
            {
                _isChecked = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsChecked)));
            }
        }
    }

    public string GsCode { get; init; } = "";
    public int RowNum { get; init; }
    public string ProductName { get; init; } = "";
    public string HomeMarketStatus { get; set; } = "";
    public string ReadyMarketStatus { get; set; } = "";
    public string CoupangStatus { get; set; } = "";

    public static string FormatDate(System.DateTime? dt) =>
        dt.HasValue ? dt.Value.ToString("MM/dd HH:mm") : "";
}
