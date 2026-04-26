using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using KeywordOcr.Core.Models;

namespace KeywordOcr.Core.Services;

/// <summary>
/// LLM + fallback 버킷 분류 → 마켓별 태그/검색어 생성 (Python market_keywords.py 완전 포팅)
/// </summary>
public class MarketKeywordsService
{
    private readonly AnthropicApiClient? _claude;

    public MarketKeywordsService(AnthropicApiClient? claude = null)
    {
        _claude = claude;
    }

    // ── 버킷 순서 ─────────────────────────────────────────────────────────────

    private static readonly string[] _bucketOrder =
    [
        "identity","usage_context","function","problem_solution",
        "material_spec","audience_scene","synonyms",
    ];

    // ── 금칙 / 힌트 ───────────────────────────────────────────────────────────

    private static readonly HashSet<string> _extraBan = new(StringComparer.Ordinal)
    {
        "마켓","스토어","쇼핑몰","샵","몰","상품","제품","정품","할인","배송","쿠폰",
        "당일","무료","특가","행사","사은품","추천","인기","선물","귀여운","예쁜",
        "고급진","럭셔리","힐링","인싸","필수품","데일리",
    };

    private static readonly HashSet<string> _usageHints = new(StringComparer.Ordinal)
    {
        "차량","본넷","보닛","트렁크","게이트","적재함","정원","전기박스","콘센트",
        "가구","도어","실내","실외","캠핑","현장","원예","호스","급수라인",
    };

    private static readonly HashSet<string> _functionHints = new(StringComparer.Ordinal)
    {
        "설치","장착","체결","연결","고정","거치","잠금","밀폐","방수","방진",
        "누수방지","회전","각도조절","분리","개폐","작업등","실링",
    };

    private static readonly HashSet<string> _problemHints = new(StringComparer.Ordinal)
    {
        "방지","차단","보호","보강","완화","해결","흔들림","누수","유입","처짐","밀폐",
    };

    private static readonly HashSet<string> _materialHints = new(StringComparer.Ordinal)
    {
        "스틸","철제","스테인리스","스텐","알루미늄","고무","플라스틱","실버","블랙",
        "화이트","304","ABS","니켈","아연합금",
    };

    private static readonly HashSet<string> _audienceHints = new(StringComparer.Ordinal)
    {
        "사용자","기사","운전자","시공","수리","작업","원예","DIY","튜닝","캠핑",
    };

    private static readonly HashSet<string> _identityHints = new(StringComparer.Ordinal)
    {
        "브라켓","브래킷","마운트","거치대","홀더","가스켓","가스킷","개스킷","패드",
        "힌지","경첩","커넥터","조인트","캐치","래치","고리","링","도어락","앵커포인트","조명",
    };

    private static readonly Regex _spaceKeepRe  = new(@"[^0-9A-Za-z가-힣\s]", RegexOptions.Compiled);
    private static readonly Regex _tokenRe       = new(@"[0-9A-Za-z가-힣]+", RegexOptions.Compiled);
    private static readonly Regex _badEndRe      = new(@"(하다|하는|되어|됨|하기|하고|하는데|이다|입니다)$", RegexOptions.Compiled);
    private static readonly Regex _badJosaRe     = new(@"(에|에서|으로|로|을|를|이|가|은|는|의|와|과)$", RegexOptions.Compiled);

    // ── 진입점 ────────────────────────────────────────────────────────────────

    public async Task<MarketKeywordPackages> GenerateAsync(
        string productName,
        string sourceText,
        string modelName = AnthropicApiClient.DefaultModel,
        IReadOnlyCollection<string>? anchors = null,
        IReadOnlyCollection<string>? baseline = null,
        string naverKeywordTable = "",
        string market = "A",
        IEnumerable<string>? avoidTerms = null,
        CancellationToken ct = default)
    {
        var anchorSet   = anchors?.Count > 0
            ? new HashSet<string>(anchors, StringComparer.Ordinal)
            : LegacyCoreService.BuildAnchorsFromName(productName);
        var baselineSet = baseline?.Count > 0
            ? new HashSet<string>(baseline, StringComparer.Ordinal)
            : LegacyCoreService.BuildBaselineTokensFromName(productName);
        var avoidKeys   = BuildAvoidKeys(avoidTerms);

        // LLM 버킷 분류
        var llmBucketed = _claude != null
            ? await GenerateLlmBucketsAsync(productName, sourceText, modelName, naverKeywordTable, ct)
            : EmptyBuckets();

        // fallback 버킷 분류 (규칙 기반)
        var fallbackBucketed = GenerateFallbackBuckets(productName, sourceText, naverKeywordTable);

        // 병합
        var bucketed = EmptyBuckets();
        foreach (var bucket in _bucketOrder)
        {
            bucketed[bucket].AddRange(llmBucketed[bucket]);
            bucketed[bucket].AddRange(fallbackBucketed[bucket]);
        }
        bucketed["synonyms"].AddRange(ExtractNaverCandidates(naverKeywordTable));

        bucketed = NormalizeBuckets(bucketed, anchorSet, baselineSet, market, avoidKeys);
        var candidatePool = FlattenBuckets(bucketed);

        var coupangTags = BuildCoupangTags(bucketed, candidatePool, productName, anchorSet, baselineSet, market, avoidKeys);
        var naverTags   = BuildNaverTags(bucketed, candidatePool, productName, anchorSet, baselineSet, market, avoidKeys);

        var searchSource = coupangTags.Count > 0
            ? coupangTags
            : candidatePool.Select(CompactPhrase).ToList();
        var searchKeywords = string.Join(" ", searchSource.Take(18)).Trim();

        return new MarketKeywordPackages
        {
            SearchKeywords = searchKeywords,
            CoupangTags    = coupangTags,
            NaverTags      = naverTags,
            CandidatePool  = candidatePool,
            BucketMap      = bucketed,
        };
    }

    // ── LLM 버킷 ─────────────────────────────────────────────────────────────

    private async Task<Dictionary<string, List<string>>> GenerateLlmBucketsAsync(
        string productName, string sourceText, string modelName,
        string naverKeywordTable, CancellationToken ct)
    {
        if (_claude == null || modelName == "없음") return EmptyBuckets();

        var source = NormalizePhrase(sourceText)[..Math.Min(1800, NormalizePhrase(sourceText).Length)];
        var naver  = NormalizePhrase(naverKeywordTable)[..Math.Min(900, NormalizePhrase(naverKeywordTable).Length)];

        var systemMsg =
            "당신은 국내 이커머스 키워드 구조 분류기다. " +
            "추상화하거나 새 카테고리를 창작하지 말고, 입력에서 근거가 있는 표현만 선택해 JSON으로 분류하라. " +
            "JSON만 반환하라. " +
            "각 후보는 2~20자, 명사구 중심, 조사/문장형/광고문구/배송문구 금지.";

        var userMsg =
            $$"""
아래 JSON 스키마로만 반환하라:
{
  "identity": string[],
  "usage_context": string[],
  "function": string[],
  "problem_solution": string[],
  "material_spec": string[],
  "audience_scene": string[],
  "synonyms": string[]
}

버킷별 규칙:
- identity: 제품 정체성, 제품유형, 상위/하위 카테고리, 2~6개
- usage_context: 사용 공간, 설치 위치, 사용 상황, 1~4개
- function: 기능, 동작, 장착/연결 방식, 1~4개
- problem_solution: 방지/차단/보호/정리 목적, 0~3개
- material_spec: 재질, 색상, 규격, 호환 힌트, 0~3개
- audience_scene: 사용자 유형, 현장 표현, 구매 문맥, 0~2개
- synonyms: 실무 유사어/띄어쓰기 변형만, 0~2개

절대 규칙:
- 감성/홍보어(귀여운,예쁜,고급진,럭셔리,힐링,인싸,필수품,데일리,추천,인기) 금지
- 자동차 차량 오토바이 바이크 퀵보드 자전거처럼 무관 확장 금지
- 네이버 검색 데이터는 같은 카테고리 여부 확인과 우선순위 참고용으로만 사용

상품명: {{productName}}

OCR_Vision요약: {{source}}

네이버검색데이터: {{naver}}
""";

        try
        {
            var raw = await _claude.CompleteAsync(systemMsg, userMsg, modelName, 900, 0.1, ct);
            if (string.IsNullOrWhiteSpace(raw)) return EmptyBuckets();

            var json = ExtractJsonBlock(raw);
            if (string.IsNullOrEmpty(json)) return EmptyBuckets();

            using var doc = JsonDocument.Parse(json);
            var out_ = EmptyBuckets();
            foreach (var bucket in _bucketOrder)
            {
                if (doc.RootElement.TryGetProperty(bucket, out var arr)
                    && arr.ValueKind == JsonValueKind.Array)
                {
                    foreach (var item in arr.EnumerateArray())
                    {
                        var s = item.GetString()?.Trim();
                        if (!string.IsNullOrEmpty(s)) out_[bucket].Add(s);
                    }
                }
            }
            return out_;
        }
        catch
        {
            return EmptyBuckets();
        }
    }

    // ── Fallback 버킷 ─────────────────────────────────────────────────────────

    private static Dictionary<string, List<string>> GenerateFallbackBuckets(
        string productName, string sourceText, string naverKeywordTable)
    {
        var out_ = EmptyBuckets();
        var phrases = new List<string>();
        phrases.AddRange(CollectAdjacentPhrases(productName, 16, 2));
        phrases.AddRange(CollectAdjacentPhrases(sourceText, 40, 1));
        phrases.AddRange(ExtractNaverCandidates(naverKeywordTable));

        foreach (var phrase in phrases)
            out_[GuessB(phrase)].Add(phrase);
        return out_;
    }

    private static string GuessB(string text)
    {
        var compact = CompactPhrase(text);
        if (_identityHints.Any(h => compact.Contains(h))) return "identity";
        if (_usageHints.Any(h => compact.Contains(h)))    return "usage_context";
        if (_problemHints.Any(h => compact.Contains(h)))  return "problem_solution";
        if (_materialHints.Any(h => compact.Contains(h))) return "material_spec";
        if (_audienceHints.Any(h => compact.Contains(h))) return "audience_scene";
        if (_functionHints.Any(h => compact.Contains(h))) return "function";
        return "synonyms";
    }

    // ── 정규화 ────────────────────────────────────────────────────────────────

    private static Dictionary<string, List<string>> NormalizeBuckets(
        Dictionary<string, List<string>> bucketed,
        HashSet<string> anchors, HashSet<string> baseline,
        string market, HashSet<string> avoidKeys)
    {
        var out_ = EmptyBuckets();
        var seen = new HashSet<string>(StringComparer.Ordinal);

        foreach (var bucket in _bucketOrder)
        {
            foreach (var raw in bucketed[bucket])
            {
                var phrase = NormalizePhrase(raw);
                if (IsBadPhrase(phrase)) continue;
                if (!PassesTopic(phrase, anchors, baseline)) continue;
                var key = SemanticKey(phrase);
                if (string.IsNullOrEmpty(key) || seen.Contains(key)) continue;
                if (market == "B" && bucket != "identity" && MatchesAvoid(key, avoidKeys)) continue;
                seen.Add(key);
                out_[bucket].Add(phrase);
                if (out_[bucket].Count >= 10) break;
            }
        }
        return out_;
    }

    // ── 쿠팡 태그 ─────────────────────────────────────────────────────────────

    private static List<string> BuildCoupangTags(
        Dictionary<string, List<string>> bucketed,
        List<string> candidatePool,
        string productName,
        HashSet<string> anchors, HashSet<string> baseline,
        string market, HashSet<string> avoidKeys)
    {
        (string bucket, int quota)[] plan = market == "B"
            ? [("identity",4),("function",3),("usage_context",2),("material_spec",2),
               ("problem_solution",1),("audience_scene",1),("synonyms",1)]
            : [("identity",6),("usage_context",4),("function",4),("problem_solution",3),
               ("material_spec",2),("audience_scene",1),("synonyms",2)];
        int maxTags = market == "B" ? 14 : 20;

        return BuildTagList(bucketed, candidatePool, productName, anchors, baseline,
            market, avoidKeys, plan, maxTags, compact: true);
    }

    // ── 네이버 태그 ───────────────────────────────────────────────────────────

    private static List<string> BuildNaverTags(
        Dictionary<string, List<string>> bucketed,
        List<string> candidatePool,
        string productName,
        HashSet<string> anchors, HashSet<string> baseline,
        string market, HashSet<string> avoidKeys)
    {
        (string bucket, int quota)[] plan = market == "B"
            ? [("identity",2),("function",2),("usage_context",1),("material_spec",1),("synonyms",1)]
            : [("identity",4),("usage_context",2),("function",2),("problem_solution",1),
               ("material_spec",1),("audience_scene",1),("synonyms",1)];
        int maxTags = market == "B" ? 7 : 10;

        return BuildTagList(bucketed, candidatePool, productName, anchors, baseline,
            market, avoidKeys, plan, maxTags, compact: false, charBudget: 100);
    }

    private static List<string> BuildTagList(
        Dictionary<string, List<string>> bucketed,
        List<string> candidatePool,
        string productName,
        HashSet<string> anchors, HashSet<string> baseline,
        string market, HashSet<string> avoidKeys,
        (string bucket, int quota)[] plan,
        int maxTags, bool compact, int charBudget = int.MaxValue)
    {
        var out_ = new List<string>();
        var seen = new HashSet<string>(StringComparer.Ordinal);

        bool Push(string value)
        {
            var phrase = compact ? CompactPhrase(value) : NormalizePhrase(value);
            if (IsBadPhrase(phrase)) return false;
            if (!PassesTopic(phrase, anchors, baseline)) return false;
            var key = SemanticKey(phrase);
            if (string.IsNullOrEmpty(key) || seen.Contains(key)) return false;
            if (!compact)   // 네이버: 글자 예산
            {
                var projected = string.Join("|", out_.Append(phrase)).Length;
                if (projected > charBudget) return false;
            }
            seen.Add(key);
            out_.Add(phrase);
            return true;
        }

        foreach (var (bucket, quota) in plan)
        {
            int added = 0;
            foreach (var value in bucketed[bucket])
            {
                if (market == "B" && bucket != "identity" && MatchesAvoid(SemanticKey(value), avoidKeys))
                    continue;
                if (Push(value)) added++;
                if (added >= quota || out_.Count >= maxTags) break;
            }
            if (out_.Count >= maxTags) return out_[..maxTags];
        }

        foreach (var value in candidatePool)
        {
            if (market == "B" && MatchesAvoid(SemanticKey(value), avoidKeys)) continue;
            Push(value);
            if (out_.Count >= maxTags) return out_[..maxTags];
        }

        foreach (var value in CollectAdjacentPhrases(productName, 16, 2))
        {
            if (market == "B" && MatchesAvoid(SemanticKey(value), avoidKeys)) continue;
            Push(value);
            if (out_.Count >= maxTags) break;
        }

        return out_[..Math.Min(out_.Count, maxTags)];
    }

    // ── 네이버 키워드 추출 ────────────────────────────────────────────────────

    private static List<string> ExtractNaverCandidates(string table)
    {
        var rows = new List<(string kw, int total)>();
        var text = (table ?? "").Trim();
        if (string.IsNullOrEmpty(text)) return [];

        foreach (var line in text.Split('\n'))
        {
            var l = line.Trim();
            if (string.IsNullOrEmpty(l) || l.StartsWith("키워드|")) continue;
            var parts = l.Split('|').Select(p => p.Trim()).ToArray();
            if (parts.Length >= 4 && !string.IsNullOrEmpty(parts[0]))
            {
                int.TryParse(parts[3], out var total);
                rows.Add((parts[0], total));
            }
        }

        if (rows.Count == 0)
        {
            foreach (var label in new[] { "PC5", "MO5" })
            {
                var m = Regex.Match(text, $@"{label}=([^|]+)");
                if (!m.Success) continue;
                foreach (var kw in m.Groups[1].Value.Split(',').Select(k => k.Trim()))
                    if (!string.IsNullOrEmpty(kw)) rows.Add((kw, 0));
            }
        }

        return rows.OrderByDescending(r => r.total).Select(r => r.kw).ToList();
    }

    // ── 유틸 ──────────────────────────────────────────────────────────────────

    private static Dictionary<string, List<string>> EmptyBuckets()
        => _bucketOrder.ToDictionary(b => b, _ => new List<string>(), StringComparer.Ordinal);

    private static List<string> FlattenBuckets(Dictionary<string, List<string>> bucketed)
    {
        var out_ = new List<string>();
        foreach (var b in _bucketOrder) out_.AddRange(bucketed[b]);
        return out_;
    }

    private static string NormalizePhrase(string text)
    {
        var c = _spaceKeepRe.Replace(text ?? "", " ");
        return Regex.Replace(c, @"\s+", " ").Trim();
    }

    private static string CompactPhrase(string text)
        => NormalizePhrase(text).Replace(" ", "");

    private static string SemanticKey(string text)
    {
        var key = LegacyCoreService.CleanOneKw(CompactPhrase(text)).ToLowerInvariant();
        key = key.Replace("차량용","차량").Replace("브래킷","브라켓").Replace("디링","d링")
                 .Replace("가스킷","가스켓").Replace("개스킷","가스켓")
                 .Replace("스텐","스테인리스").Replace("고정대","거치대");
        key = Regex.Replace(key, @"(용|형|식)$", "");
        return key;
    }

    private static bool IsBadPhrase(string text)
    {
        var compact = CompactPhrase(text);
        if (string.IsNullOrEmpty(compact) || compact.Length < 2 || compact.Length > 20) return true;
        if (Regex.IsMatch(compact, @"^\d+$")) return true;
        if (_badEndRe.IsMatch(compact) || _badJosaRe.IsMatch(compact)) return true;
        if (LegacyCoreService.Ban.Any(b => compact.Contains(b))) return true;
        if (_extraBan.Any(b => compact.Contains(b))) return true;
        return false;
    }

    private static bool PassesTopic(string text, HashSet<string> anchors, HashSet<string> baseline)
    {
        var compact = CompactPhrase(text);
        if (string.IsNullOrEmpty(compact)) return false;
        if (anchors.Count > 0 && baseline.Count > 0)
            return LegacyCoreService.IsOnTopic(compact, anchors, baseline);
        if (baseline.Count > 0)
            return LegacyCoreService.IsConsistentWithBaseline(compact, baseline);
        return true;
    }

    private static HashSet<string> BuildAvoidKeys(IEnumerable<string>? values)
    {
        if (values == null) return [];
        var keys = new HashSet<string>(StringComparer.Ordinal);
        foreach (var raw in values)
        {
            var phrase = NormalizePhrase(raw);
            if (string.IsNullOrEmpty(phrase)) continue;
            keys.Add(SemanticKey(phrase));
            foreach (Match m in _tokenRe.Matches(phrase))
                keys.Add(SemanticKey(m.Value));
        }
        return keys;
    }

    private static bool MatchesAvoid(string key, HashSet<string> avoidKeys)
    {
        if (string.IsNullOrEmpty(key) || avoidKeys.Count == 0) return false;
        foreach (var avoid in avoidKeys)
        {
            if (string.IsNullOrEmpty(avoid)) continue;
            if (key == avoid || avoid.Contains(key) || key.Contains(avoid)) return true;
        }
        return false;
    }

    private static List<string> CollectAdjacentPhrases(string text, int maxTokens, int maxSize)
    {
        var tokens = _tokenRe.Matches(NormalizePhrase(text ?? ""))
            .Cast<Match>()
            .Select(m => m.Value)
            .Where(w => w.Length >= 2 && w.Length <= 12
                && !LegacyCoreService.Stopwords.Contains(w)
                && !_extraBan.Contains(w))
            .Take(maxTokens)
            .ToList();

        var out_ = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        void Push(string value)
        {
            var phrase = NormalizePhrase(value);
            if (string.IsNullOrEmpty(phrase) || seen.Contains(phrase)) return;
            seen.Add(phrase);
            out_.Add(phrase);
        }

        foreach (var tok in tokens) Push(tok);
        for (int size = 2; size <= maxSize; size++)
            for (int i = 0; i <= tokens.Count - size; i++)
                Push(string.Join(" ", tokens[i..(i + size)]));

        return out_;
    }

    private static string ExtractJsonBlock(string text)
    {
        var m = Regex.Match(text, @"```json\s*([\s\S]*?)```", RegexOptions.IgnoreCase);
        if (m.Success) return m.Groups[1].Value.Trim();
        m = Regex.Match(text, @"```([\s\S]*?)```");
        if (m.Success) return m.Groups[1].Value.Trim();
        int start = text.IndexOf('{'), end = text.LastIndexOf('}');
        if (start >= 0 && end > start) return text[start..(end + 1)];
        return text.Trim();
    }
}
