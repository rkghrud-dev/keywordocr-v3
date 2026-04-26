using System.Text.Json;
using System.Text.RegularExpressions;

namespace KeywordOcr.Core.Services;

/// <summary>
/// Vision/OCR 증거 기반 규칙적 키워드 조립 (Python keyword_builder.py 완전 포팅)
/// </summary>
public static class KeywordBuilderService
{
    private const int TargetDefault = 20;
    private const int MinCharTargetA = 90;
    private const int MaxCharLimitA = 140;
    private const int MinCharTargetB = 63;
    private const int MaxCharLimitB = 98;
    private const int MaxTokenLen = 7;

    // ── 불용어 ────────────────────────────────────────────────────────────────

    private static readonly HashSet<string> _stopwords = new(StringComparer.OrdinalIgnoreCase)
    {
        "및","또는","에서","으로","같은","관련","용","용도","제품","상품","기타",
        "가능","활용","사용","적용","구성","기본","일반","세트","단품","옵션",
        "빠른","유지","작업","통해","위해","위한","따른","위의","후에","대한",
        "포함","약간","하여","있는","없는","있음","없음","됨","함","위치",
        "취급","용이","우수","뛰어난","다양한","선택","추천","가능한","적합한",
        "필요한","특징","장점","효과","방법","소형","대형","중형","라인업","사이즈",
        "경질","연질","고정형","이동형","겸용","내부","외부","다사이즈",
        "시리즈","용품","배관용품",
        "shaped","clamp","type","style","pipe","tube","hose","profile","point","option","con",
    };

    private static readonly HashSet<string> _oddSingleWords = new(StringComparer.Ordinal)
    {
        "펼침","접힘","열림","닫힘","분리","조절","가능","추천","강화","완성",
    };

    // ── 동의어 그룹 ───────────────────────────────────────────────────────────

    private static readonly Dictionary<string, string[]> _synonymGroups = new(StringComparer.Ordinal)
    {
        ["bracket_mount"]    = ["브라켓","브래킷","마운트","거치대","홀더","고정대","클램프"],
        ["install"]          = ["설치","장착","체결","부착","고정"],
        ["no_drill"]         = ["무타공","무천공","타공없음"],
        ["angle_rotate"]     = ["각도조절","각도조정","회전","회전형"],
        ["location_vehicle"] = ["본넷","보닛","트렁크","게이트"],
        ["material"]         = ["스틸","스텐","스테인리스","철제","알루미늄"],
        ["color_black"]      = ["블랙","검정","검은색"],
        ["color_silver"]     = ["실버","은색"],
    };

    // ── HEAD / CONTEXT 어휘 ───────────────────────────────────────────────────

    private static readonly string[] _headSuffixes =
    [
        "브라켓","브래킷","거치대","받침대","지지대","홀더","클립",
        "커넥터","조인트","클램프","노즐","테이프","커버","마개",
        "캡","패드","브러시","필터","밸브","후크","고리",
        "볼트","너트","핀","호스","파이프","케이블",
    ];

    private static readonly string[] _nameContextWords =
    [
        "하수구","배수구","세면대","싱크대","욕실","주방","차량","자동차",
        "스위치","호스","배관","파이프","관개","정원","조명","전선",
        "케이블","벽면","천장","창문","문","트렁크","본넷","밑창",
    ];

    private static readonly string[] _nameActionWords =
    [
        "세척","고정","연결","장착","설치","분사","배수","누수",
        "방지","보호","정리","거치","교체","수리","지지","보수",
    ];

    private static readonly HashSet<string> _contextSuffixHeads = new(StringComparer.Ordinal)
    {
        "핀","브라켓","브래킷","거치대","받침대","지지대","홀더",
        "클립","커넥터","조인트","클램프","테이프","커버","마개",
        "캡","패드","후크","고리","볼트","너트",
    };

    private static readonly HashSet<string> _allowedCompounds = new(StringComparer.Ordinal)
    { "고정핀" };

    private static readonly HashSet<string> _actionRedundantHeads = new(StringComparer.Ordinal)
    { "커넥터","조인트" };

    private static readonly HashSet<string> _genericWords = new(StringComparer.Ordinal)
    {
        "고정","자재","부품","소재","재료","도구","공구","용품","소품","제품","상품",
        "부자재","배관자재","설비","설치","연결","장착","정리","방지","보호",
        "옵션","사용","실내","실외","간편","강력","다양한","사이즈","작업효율",
        "편리","휴대용","다용도","구조","형태",
    };

    private static readonly Regex _josaRe = new(
        @"(을|를|에|의|은|는|가|로|와|과|에서|으로|하여|에도|까지)$", RegexOptions.Compiled);
    private static readonly Regex _gsCodeRe = new(@"GS\d{7}[A-Za-z]?", RegexOptions.Compiled);
    private static readonly Regex _cleanRe  = new(@"[^0-9A-Za-z가-힣\-\+ ]", RegexOptions.Compiled);
    private static readonly Regex _tokenRe  = new(@"[0-9A-Za-z가-힣]+", RegexOptions.Compiled);
    private static readonly Regex _colorFilterRe = new(
        @"^(흰색|검정|검은색|화이트|블랙|실버|은색|회색|그레이|빨간색|파란색|노란색|녹색|흰색화이트|블랙검정)$",
        RegexOptions.Compiled);

    // ── 진입점 ────────────────────────────────────────────────────────────────

    /// <summary>
    /// Vision/OCR 증거만 사용해 키워드 문자열을 조립한다.
    /// (Python build_keyword_string 완전 포팅)
    /// </summary>
    public static string BuildKeywordString(
        string ocrText,
        Dictionary<string, object?>? visionAnalysis,
        int targetCount = TargetDefault,
        string fallbackText = "",
        string market = "A")
    {
        try
        {
            if (ocrText?.Contains("OCR 텍스트 없음") == true) ocrText = "";
            ocrText ??= "";
            targetCount = Math.Max(1, targetCount);

            int minChar = market == "B" ? MinCharTargetB : MinCharTargetA;
            int maxChar = market == "B" ? MaxCharLimitB  : MaxCharLimitA;

            var analysis = visionAnalysis ?? [];
            var axis = ExtractRequiredAxes(analysis);

            var catWords  = DedupeNormalized(axis["category"]);
            var typeWords = DedupeNormalized(axis["product_type_correction"]);
            var coreWords = DedupeNormalized(
                axis["category"].Concat(axis["product_type_correction"]).ToList());

            var baseCore = PickBaseCore(catWords, typeWords);
            var nameIdentityTokens = ExtractNameOnlyTokens(fallbackText, market);

            // ── 검증 함수들 ─────────────────────────────────────────────────

            string SemKey(string token)
                => CoreForm(NormalizeToken(token));

            bool IsGenericToken(string token)
            {
                var tok = NormalizeToken(token);
                var key = SemKey(tok);
                if (string.IsNullOrEmpty(tok) || string.IsNullOrEmpty(key)) return true;
                if (_genericWords.Contains(key) || _stopwords.Contains(tok.ToLowerInvariant())) return true;
                if (key.EndsWith("용품") || key.EndsWith("부품") || key.EndsWith("도구")
                    || key.EndsWith("세트") || key.EndsWith("옵션")) return true;
                return false;
            }

            bool HasOverlap(string token, IReadOnlyList<string> refs)
            {
                var key = SemKey(token);
                if (string.IsNullOrEmpty(key)) return false;
                foreach (var r in refs)
                {
                    var rk = SemKey(r);
                    if (string.IsNullOrEmpty(rk)) continue;
                    if (key == rk) return true;
                    if (key.Length >= 3 && rk.Length >= 3 && (key.Contains(rk) || rk.Contains(key)))
                        return true;
                }
                return false;
            }

            List<string> DedupeSemanticLocal(IEnumerable<string> items)
            {
                var out2 = new List<string>();
                var seen2 = new HashSet<string>(StringComparer.Ordinal);
                foreach (var item in items)
                {
                    var tok = NormalizeToken(item);
                    var key = SemKey(tok);
                    if (string.IsNullOrEmpty(tok) || string.IsNullOrEmpty(key) || seen2.Contains(key)) continue;
                    seen2.Add(key);
                    out2.Add(tok);
                }
                return out2;
            }

            // ── identity terms ──────────────────────────────────────────────
            var identityTerms = new List<string>();
            if (!string.IsNullOrEmpty(baseCore) && !IsGenericToken(baseCore))
                identityTerms.Add(baseCore);
            identityTerms.AddRange(coreWords);
            identityTerms.AddRange(nameIdentityTokens);
            identityTerms = DedupeSemanticLocal(identityTerms)
                .Where(t => !IsGenericToken(t)).ToList();

            // ── usage terms ─────────────────────────────────────────────────
            var usageTerms = new List<string>();
            foreach (var raw in coreWords)
            {
                var tok = NormalizeToken(raw);
                if (string.IsNullOrEmpty(tok) || tok == baseCore || IsGenericToken(tok)) continue;
                usageTerms.Add(tok);
                if (tok.Length <= 3 && (tok + "용").Length <= MaxTokenLen)
                    usageTerms.Add(tok + "용");
            }
            usageTerms.AddRange(DedupeNormalized(
                axis["usage_location"].Concat(axis["space_keywords"]).ToList()));
            usageTerms = DedupeSemanticLocal(usageTerms).Where(t => !IsGenericToken(t)).ToList();

            // ── function terms ──────────────────────────────────────────────
            var actionVocab = new HashSet<string>(StringComparer.Ordinal)
            {
                "고정","정리","방지","설치","시공","조절","장착","연결","분리","교체",
                "보호","차단","밀봉","강화","지지","수납","거치","탈착","흔들림","내구성","내식성",
            };
            var knownWords = new HashSet<string>(identityTerms.Concat(usageTerms).Concat(actionVocab),
                StringComparer.Ordinal);
            var rawFunctions = DedupeNormalized(
                axis["problem_solving_keyword"].Concat(axis["usage_purpose"])
                    .Concat(axis["benefit_keywords"]).ToList());
            var functionTerms = new List<string>();
            foreach (var raw in rawFunctions)
            {
                var fn = NormalizeToken(raw);
                if (string.IsNullOrEmpty(fn) || IsGenericToken(fn)) continue;
                var split = fn.Length > 3 ? SplitCompoundOnce(fn, knownWords) : [fn];
                foreach (var part in split)
                {
                    var p = NormalizeToken(part);
                    if (!string.IsNullOrEmpty(p) && !IsGenericToken(p))
                        functionTerms.Add(p);
                }
            }
            functionTerms = DedupeSemanticLocal(functionTerms);

            // ── boost terms ─────────────────────────────────────────────────
            var evidenceRefs = identityTerms.Concat(usageTerms).Concat(functionTerms).ToList();
            var rawBoosts = DedupeNormalized(
                axis["installation_keywords"].Concat(axis["longtail_candidates"]).ToList());
            var boostTerms = new List<string>();
            foreach (var raw in rawBoosts)
            {
                foreach (var part in ExpandTerm(raw, knownWords))
                {
                    var p = NormalizeToken(part);
                    if (string.IsNullOrEmpty(p) || IsGenericToken(p)) continue;
                    if (!HasOverlap(p, evidenceRefs)) continue;
                    boostTerms.Add(p);
                }
            }
            boostTerms = DedupeSemanticLocal(boostTerms);

            // ── OCR terms ───────────────────────────────────────────────────
            evidenceRefs = identityTerms.Concat(usageTerms).Concat(functionTerms).Concat(boostTerms).ToList();
            var ocrTerms = new List<string>();
            foreach (var token in TokenizeOcr(ocrText))
            {
                if (IsGenericToken(token)) continue;
                if (!HasOverlap(token, evidenceRefs)) continue;
                ocrTerms.Add(token);
            }
            ocrTerms = DedupeSemanticLocal(ocrTerms);

            // ── fallback terms ──────────────────────────────────────────────
            var fallbackTerms = DedupeSemanticLocal(nameIdentityTokens)
                .Where(t => !IsGenericToken(t)).ToList();

            // ── 최종 조립 ────────────────────────────────────────────────────
            var out_ = new List<string>();
            var seen = new HashSet<string>(StringComparer.Ordinal);

            int CharLen() => out_.Sum(t => t.Length) + Math.Max(0, out_.Count - 1);
            bool IsFull() => out_.Count >= targetCount || CharLen() >= maxChar;

            bool TryAdd(string token)
            {
                var t = NormalizeToken(token);
                var key = SemKey(t);
                if (string.IsNullOrEmpty(t) || string.IsNullOrEmpty(key) || t.Length < 2
                    || _stopwords.Contains(t.ToLowerInvariant())) return false;
                if (t.EndsWith("소재") && t.Length > 2) return false;
                if (_colorFilterRe.IsMatch(t)) return false;
                if (t.Length > MaxTokenLen) return false;
                if (seen.Contains(key)) return false;
                seen.Add(key);
                out_.Add(t);
                return true;
            }

            foreach (var group in new[] { identityTerms, usageTerms, functionTerms,
                                           boostTerms, ocrTerms, fallbackTerms })
            {
                foreach (var token in group)
                {
                    if (IsFull()) break;
                    TryAdd(token);
                }
                if (IsFull()) break;
            }

            // 글자수 미달 시 재시도
            if (CharLen() < minChar)
            {
                var remaining = DedupeSemanticLocal(
                    identityTerms.Concat(usageTerms).Concat(functionTerms)
                        .Concat(boostTerms).Concat(ocrTerms).Concat(fallbackTerms));
                foreach (var token in remaining)
                {
                    if (IsFull()) break;
                    TryAdd(token);
                }
            }

            return string.Join(" ", out_.Take(targetCount)).Trim();
        }
        catch
        {
            return "";
        }
    }

    // ── 축 추출 ───────────────────────────────────────────────────────────────

    private static Dictionary<string, List<string>> ExtractRequiredAxes(
        Dictionary<string, object?> analysis)
    {
        string[] axisKeys = [
            "category","product_type_correction","structure","material_visual","color",
            "mount_type","installation_method","usage_location","usage_purpose","target_user",
            "usage_scenario","indoor_outdoor","primary_function","problem_solving_keyword",
            "convenience_feature","installation_keywords","space_keywords","benefit_keywords",
            "longtail_candidates",
        ];
        var result = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        foreach (var k in axisKeys)
            result[k] = [];

        foreach (var (section, sectionVal) in analysis)
        {
            if (sectionVal is not JsonElement je) continue;
            if (je.ValueKind != JsonValueKind.Object) continue;
            foreach (var prop in je.EnumerateObject())
            {
                if (result.TryGetValue(prop.Name, out var list))
                    list.AddRange(ToList(prop.Value));
            }
        }
        return result;
    }

    private static List<string> ToList(JsonElement el)
    {
        var out_ = new List<string>();
        switch (el.ValueKind)
        {
            case JsonValueKind.Array:
                foreach (var item in el.EnumerateArray())
                    out_.AddRange(ToList(item));
                break;
            case JsonValueKind.String:
                var s = CleanText(el.GetString() ?? "");
                if (!string.IsNullOrEmpty(s)) out_.Add(s);
                break;
            case JsonValueKind.Object:
                if (el.TryGetProperty("value", out var vEl))
                    out_.AddRange(ToList(vEl));
                else
                    foreach (var p in el.EnumerateObject())
                        out_.AddRange(ToList(p.Value));
                break;
        }
        return out_;
    }

    // ── 토큰 유틸 ─────────────────────────────────────────────────────────────

    public static string CleanText(string s)
    {
        s = Regex.Replace(s ?? "", @"[\t\r\n]+", " ");
        s = Regex.Replace(s, @"[|,;/·•‧]+", " ");
        return Regex.Replace(s, @"\s+", " ").Trim();
    }

    public static string NormalizeToken(string tok)
    {
        tok = CleanText(tok ?? "");
        tok = _cleanRe.Replace(tok, "");
        return Regex.Replace(tok, @"\s+", " ").Trim();
    }

    public static string CoreForm(string tok)
    {
        var t = tok.ToLowerInvariant();
        t = t.Replace("차량용", "차량")
             .Replace("브래킷", "브라켓")
             .Replace("디링", "d링");
        t = Regex.Replace(t, @"(용|형|식)$", "");
        return t;
    }

    private static string? SynGroup(string tok)
    {
        var t = CoreForm(tok);
        foreach (var (grp, words) in _synonymGroups)
            foreach (var w in words)
                if (CoreForm(w) == t) return grp;
        return null;
    }

    // ── DP 복합어 분해 (Python _split_compound_once 완전 포팅) ───────────────

    public static List<string> SplitCompoundOnce(string term, HashSet<string> vocab)
    {
        var s = term.Trim();
        if (string.IsNullOrEmpty(s) || s.Contains(' ')) return [s];
        if (!Regex.IsMatch(s, @"[가-힣]")) return [s];

        int n = s.Length;
        // dp[i] = (score, tokens) | null
        (int score, List<string> tokens)?[] dp = new (int, List<string>)?[n + 1];
        dp[0] = (0, []);

        for (int i = 0; i < n; i++)
        {
            if (dp[i] is null) continue;
            var (curScore, curTokens) = dp[i]!.Value;

            for (int j = i + 2; j <= Math.Min(n, i + 8); j++)
            {
                var piece = s[i..j];
                if (!vocab.Contains(piece)) continue;

                int candScore = curScore + piece.Length;
                var candTokens = new List<string>(curTokens) { piece };
                var old = dp[j];

                if (old is null
                    || candScore > old.Value.score
                    || (candScore == old.Value.score && candTokens.Count < old.Value.tokens.Count))
                {
                    dp[j] = (candScore, candTokens);
                }
            }
        }

        if (dp[n] is not null
            && dp[n]!.Value.tokens.Count >= 2
            && string.Concat(dp[n]!.Value.tokens) == s)
        {
            return dp[n]!.Value.tokens;
        }
        return [s];
    }

    private static List<string> ExpandTerm(string term, HashSet<string> vocab)
    {
        var t = NormalizeToken(term);
        if (string.IsNullOrEmpty(t)) return [];

        var outs = new List<string> { t };
        if (t.Contains(' '))
            outs.AddRange(t.Split(' ', StringSplitOptions.RemoveEmptyEntries));

        var splitOnce = SplitCompoundOnce(t, vocab);
        if (splitOnce.Count >= 2)
        {
            outs.AddRange(splitOnce);
            if (splitOnce[0] == "차량용") outs.Add("차량");
        }
        return outs.Where(x => !string.IsNullOrEmpty(NormalizeToken(x))).ToList();
    }

    private static List<string> TokenizeOcr(string ocrText)
    {
        var txt = NormalizeToken(ocrText ?? "");
        if (string.IsNullOrEmpty(txt)) return [];

        var out_ = new List<string>();
        foreach (Match m in _tokenRe.Matches(txt))
        {
            var w = m.Value;
            if (w.Length < 2 || w.Length > 14) continue;
            if (_stopwords.Contains(w.ToLowerInvariant())) continue;
            if (Regex.IsMatch(w, @"^\d+$")) continue;
            out_.Add(w);
        }
        return out_;
    }

    private static bool IsOddToken(string tok)
    {
        var t = CoreForm(tok);
        if (_oddSingleWords.Contains(t)) return true;
        if (t.Length <= 2 && Regex.IsMatch(t, @"(함|됨)$")) return true;
        return false;
    }

    private static List<string> DedupeNormalized(List<string> items)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var out_ = new List<string>();
        foreach (var item in items)
        {
            var t = NormalizeToken(item);
            if (string.IsNullOrEmpty(t)) continue;

            var words = t.Contains(' ') ? t.Split(' ', StringSplitOptions.RemoveEmptyEntries) : [t];
            foreach (var w in words)
            {
                var stripped = StripJosa(w.Trim());
                if (string.IsNullOrEmpty(stripped) || stripped.Length < 2 || seen.Contains(stripped)) continue;
                if (_stopwords.Contains(stripped.ToLowerInvariant())) continue;
                seen.Add(stripped);
                out_.Add(stripped);
            }
        }
        return out_;
    }

    private static string StripJosa(string w)
    {
        if (!Regex.IsMatch(w, @"[가-힣]")) return w;
        var cleaned = _josaRe.Replace(w, "");
        return cleaned.Length >= 2 ? cleaned : w;
    }

    private static string PickBaseCore(List<string> categoryWords, List<string> typeWords)
    {
        foreach (var w in Enumerable.Reverse(categoryWords))
            if (!_genericWords.Contains(w) && w.Length >= 2) return w;
        foreach (var w in Enumerable.Reverse(typeWords))
            if (!_genericWords.Contains(w) && w.Length >= 2) return w;
        if (categoryWords.Count > 0) return categoryWords[^1];
        if (typeWords.Count > 0) return typeWords[^1];
        return "";
    }

    private static List<string> ExtractNameOnlyTokens(string fallbackText, string market = "A")
    {
        var raw = NormalizeToken(_gsCodeRe.Replace(fallbackText ?? "", " ")).Trim();
        if (string.IsNullOrEmpty(raw)) return [];

        var rawParts = raw.Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Where(p => !Regex.IsMatch(p, @"\d")).ToList();
        raw = rawParts.Count > 0 ? string.Join(" ", rawParts) : raw;

        var joined = raw.Replace(" ", "");
        var out_ = new List<string>();
        var seen = new HashSet<string>(StringComparer.Ordinal);

        void Push(string token)
        {
            var t = NormalizeToken(token);
            if (string.IsNullOrEmpty(t) || t.Length < 2 || t.Length > MaxTokenLen || seen.Contains(t)) return;
            seen.Add(t);
            out_.Add(t);
        }

        string ContextToken(string word, bool useSuffix)
        {
            var token = NormalizeToken(word);
            if (string.IsNullOrEmpty(token)) return "";
            if (useSuffix && !token.EndsWith("용"))
            {
                var cand = token + "용";
                if (cand.Length <= MaxTokenLen) return cand;
            }
            return token;
        }

        string? head = null;
        foreach (var suffix in _headSuffixes)
            if (joined.EndsWith(suffix, StringComparison.Ordinal)) { head = suffix; break; }

        var stem = head != null ? joined[..^head.Length] : joined;

        var matchedContexts = _nameContextWords.Where(w => stem.Contains(w)).ToList();
        var matchedActions  = _nameActionWords.Where(w => stem.Contains(w)).ToList();
        var primaryAction   = matchedActions.Count > 0 ? matchedActions[0] : null;

        string? compactToken = null;
        if (primaryAction != null && head != null)
        {
            var cand = primaryAction + head;
            if (_allowedCompounds.Contains(cand) && cand.Length <= MaxTokenLen)
                compactToken = cand;
        }

        bool keepAction = primaryAction != null && compactToken == null && !_actionRedundantHeads.Contains(head ?? "");
        bool useContextSuffix = head != null && _contextSuffixHeads.Contains(head)
            && (compactToken != null || !keepAction);

        if (matchedContexts.Count > 0) Push(ContextToken(matchedContexts[0], useContextSuffix));
        foreach (var extra in matchedContexts.Skip(1)) Push(extra);

        if (compactToken != null)
            Push(compactToken);
        else
        {
            if (keepAction && primaryAction != null) Push(primaryAction);
            if (head != null) Push(head);
            else if (primaryAction != null) Push(primaryAction);
        }

        if (out_.Count == 0)
            foreach (var part in raw.Split(' ', StringSplitOptions.RemoveEmptyEntries))
                Push(part);
        if (out_.Count == 0 && !string.IsNullOrEmpty(joined))
            Push(joined);

        return out_.Take(4).ToList();
    }
}
