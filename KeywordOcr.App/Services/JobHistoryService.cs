using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace KeywordOcr.App.Services;

public sealed class JobRecord
{
    [JsonPropertyName("id")]
    public string Id { get; set; } = Guid.NewGuid().ToString("N")[..8];

    [JsonPropertyName("timestamp")]
    public DateTime Timestamp { get; set; } = DateTime.Now;

    [JsonPropertyName("source_file")]
    public string SourceFile { get; set; } = "";

    [JsonPropertyName("output_root")]
    public string OutputRoot { get; set; } = "";

    [JsonPropertyName("output_file")]
    public string OutputFile { get; set; } = "";

    [JsonPropertyName("product_count")]
    public int ProductCount { get; set; }

    [JsonPropertyName("selected_codes")]
    public List<string> SelectedCodes { get; set; } = new();

    [JsonPropertyName("model")]
    public string Model { get; set; } = "";

    [JsonPropertyName("make_listing")]
    public bool MakeListing { get; set; }

    [JsonPropertyName("memo")]
    public string Memo { get; set; } = "";

    [JsonPropertyName("status")]
    public string Status { get; set; } = "완료";

    // UI 표시용
    [JsonIgnore]
    public string DisplayTime => Timestamp.ToString("MM/dd HH:mm");

    [JsonIgnore]
    public string DisplaySource => Path.GetFileName(SourceFile);

    [JsonIgnore]
    public string DisplaySummary =>
        $"{ProductCount}개 상품 | {Model} | {(MakeListing ? "이미지O" : "이미지X")}";
}

public sealed class JobHistoryService
{
    private readonly string _historyPath;
    private List<JobRecord> _records = new();

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    public JobHistoryService(string appRoot)
    {
        _historyPath = Path.Combine(appRoot, "job_history.json");
        Load();
    }

    public IReadOnlyList<JobRecord> Records => _records;

    public void Add(JobRecord record)
    {
        _records.Insert(0, record);
        Save();
    }

    public void Update(JobRecord record)
    {
        var idx = _records.FindIndex(r => r.Id == record.Id);
        if (idx >= 0)
        {
            _records[idx] = record;
            Save();
        }
    }

    public void Delete(string id)
    {
        _records.RemoveAll(r => r.Id == id);
        Save();
    }

    public JobRecord? FindById(string id) => _records.FirstOrDefault(r => r.Id == id);

    public JobRecord Clone(JobRecord source)
    {
        var clone = new JobRecord
        {
            SourceFile = source.SourceFile,
            OutputRoot = source.OutputRoot,
            OutputFile = source.OutputFile,
            ProductCount = source.ProductCount,
            SelectedCodes = new List<string>(source.SelectedCodes),
            Model = source.Model,
            MakeListing = source.MakeListing,
            Memo = source.Memo + " (복사)",
            Status = "대기",
        };
        Add(clone);
        return clone;
    }

    private void Load()
    {
        if (!File.Exists(_historyPath))
        {
            _records = new List<JobRecord>();
            return;
        }

        try
        {
            var json = File.ReadAllText(_historyPath);
            _records = JsonSerializer.Deserialize<List<JobRecord>>(json, JsonOptions)
                       ?? new List<JobRecord>();
        }
        catch
        {
            _records = new List<JobRecord>();
        }
    }

    private void Save()
    {
        try
        {
            var dir = Path.GetDirectoryName(_historyPath);
            if (!string.IsNullOrEmpty(dir))
                Directory.CreateDirectory(dir);

            var json = JsonSerializer.Serialize(_records, JsonOptions);
            File.WriteAllText(_historyPath, json);
        }
        catch { /* 저장 실패는 무시 */ }
    }
}
