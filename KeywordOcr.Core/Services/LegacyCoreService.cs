using System.Text.RegularExpressions;

namespace KeywordOcr.Core.Services;

/// <summary>
/// 앵커/베이스라인/금칙어/토픽 필터 (Python legacy_core.py 핵심 포팅)
/// </summary>
public static class LegacyCoreService
{
    // ── 금칙어 ────────────────────────────────────────────────────────────────

    public static readonly HashSet<string> Ban = new(StringComparer.Ordinal)
    {
        "정품","국내발송","무료배송","행사","특가","세일","인기","추천","최고","프리미엄",
        "신제품","베스트","한정","할인","쿠폰","증정","사은품","당일발송","당일배송",
        "빠른배송","익일발송","선착순","한정수량","품절임박","스페셜","에디션",
        "가성비","가심비","대박","핫딜","타임세일","플래시","이벤트","혜택",
    };

    public static readonly HashSet<string> Stopwords = new(StringComparer.Ordinal)
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
        // SIZE_WORDS
        "mm","cm","m","L","리터","ml","미리","밀리","센치","센티",
        "인치","inch","피트","feet","foot","야드",
        // BAN_EXACT
        "안정적","안정적인","안정","사용법","탑재",
    };

    public static readonly HashSet<string> BanSubstrings = new(StringComparer.Ordinal)
    {
        "까지","최대","절단폭",
    };

    // ── HEAD / CONTEXT 어휘 (앵커 추출용) ────────────────────────────────────

    private static readonly string[] _headSuffixes =
    [
        "브라켓","브래킷","거치대","받침대","지지대","홀더","클립",
        "커넥터","조인트","클램프","노즐","테이프","커버","마개",
        "캡","패드","브러시","필터","밸브","후크","고리",
        "볼트","너트","핀","호스","파이프","케이블","고정대",
        "가스켓","가스킷","개스킷","힌지","경첩","캐치","래치",
        "링","앵커","마운트","조명","도어락",
    ];

    private static readonly Regex _gsCodeRe = new(@"GS\d{7}[A-Za-z]?", RegexOptions.Compiled);
    private static readonly Regex _tokenRe  = new(@"[0-9A-Za-z가-힣]+", RegexOptions.Compiled);
    private static readonly Regex _cleanKwRe = new(@"[^가-힣A-Za-z0-9]", RegexOptions.Compiled);
    private static readonly Regex _leadDigitRe = new(@"^\d+", RegexOptions.Compiled);

    // ── 공통 유틸 ─────────────────────────────────────────────────────────────

    /// <summary>Python _clean_one_kw 포팅 — 특수문자 제거, 선행 숫자 제거</summary>
    public static string CleanOneKw(string k)
    {
        k = (k ?? "").Trim();
        k = _cleanKwRe.Replace(k, "");
        k = _leadDigitRe.Replace(k, "");
        return k.Trim();
    }

    /// <summary>정규화 후 공백 제거한 키 반환</summary>
    public static string SemanticKey(string text)
    {
        var key = CleanOneKw(text.Replace(" ", "")).ToLowerInvariant();
        key = key.Replace("차량용", "차량")
                 .Replace("브래킷", "브라켓")
                 .Replace("디링", "d링")
                 .Replace("가스킷", "가스켓")
                 .Replace("개스킷", "가스켓")
                 .Replace("스텐", "스테인리스")
                 .Replace("고정대", "거치대");
        key = Regex.Replace(key, @"(용|형|식)$", "");
        return key;
    }

    // ── 앵커 / 베이스라인 ─────────────────────────────────────────────────────

    /// <summary>상품명에서 identity 중심 앵커 추출 (Python build_anchors_from_name 포팅)</summary>
    public static HashSet<string> BuildAnchorsFromName(string productName)
    {
        var tokens = CollectIdentityTokens(productName, maxMain: 4, maxTotal: 6);
        return tokens.Count > 0 ? tokens : [SemanticKey(productName)];
    }

    /// <summary>상품명에서 baseline 토큰 추출 (Python build_baseline_tokens_from_name 포팅)</summary>
    public static HashSet<string> BuildBaselineTokensFromName(string productName)
    {
        var tokens = CollectIdentityTokens(productName, maxMain: 4, maxTotal: 6);
        if (tokens.Count > 0) return tokens;

        // fallback: 상품명 전체 토큰화
        var fallback = new HashSet<string>(StringComparer.Ordinal);
        foreach (var m in _tokenRe.Matches(productName).Cast<Match>())
        {
            var t = m.Value;
            if (t.Length >= 2 && !Stopwords.Contains(t.ToLowerInvariant()))
                fallback.Add(SemanticKey(t));
        }
        return fallback;
    }

    private static HashSet<string> CollectIdentityTokens(string name, int maxMain, int maxTotal)
    {
        // GS코드 제거 후 토큰화
        var cleaned = _gsCodeRe.Replace(name ?? "", " ");
        var words = _tokenRe.Matches(cleaned)
            .Cast<Match>()
            .Select(m => m.Value)
            .Where(w => w.Length >= 2 && !Regex.IsMatch(w, @"^\d+$") && !Stopwords.Contains(w.ToLowerInvariant()))
            .ToList();

        var result = new HashSet<string>(StringComparer.Ordinal);

        // HEAD 접미사 추출 (대표 상품군)
        var joined = string.Concat(words);
        string? head = null;
        foreach (var suffix in _headSuffixes)
        {
            if (joined.EndsWith(suffix, StringComparison.Ordinal))
            { head = suffix; break; }
        }
        if (head != null) result.Add(SemanticKey(head));

        // 나머지 단어 중 명사만 추가
        foreach (var w in words)
        {
            if (result.Count >= maxTotal) break;
            var key = SemanticKey(w);
            if (!string.IsNullOrEmpty(key) && !result.Contains(key))
                result.Add(key);
        }

        return result;
    }

    // ── 온토픽 필터 ───────────────────────────────────────────────────────────

    /// <summary>금칙 도메인 배제 + 앵커/베이스라인 일관성 (Python is_on_topic 포팅)</summary>
    public static bool IsOnTopic(string keyword, HashSet<string> anchors, HashSet<string> baseline)
    {
        var k = CleanOneKw(keyword);
        if (string.IsNullOrEmpty(k)) return false;
        if (Ban.Contains(k)) return false;
        if (BanSubstrings.Any(bs => k.Contains(bs))) return false;
        if (anchors.Count > 0 && !SemanticOverlapCheck(k, anchors)) return false;
        if (baseline.Count > 0 && !IsConsistentWithBaseline(k, baseline)) return false;
        return true;
    }

    /// <summary>베이스라인 일관성 검사 (Python is_consistent_with_baseline 포팅)</summary>
    public static bool IsConsistentWithBaseline(string keyword, HashSet<string> baseline)
    {
        if (baseline.Count == 0) return true;
        var k = CleanOneKw(keyword);
        if (string.IsNullOrEmpty(k)) return false;

        // 토큰화
        var kwTokens = _tokenRe.Matches(k).Cast<Match>()
            .Select(m => SemanticKey(m.Value))
            .Where(t => t.Length >= 2)
            .ToList();

        int overlap = 0;
        foreach (var bt in baseline)
        {
            foreach (var kt in kwTokens)
            {
                if (kt == bt || (kt.Length >= 3 && bt.Length >= 3 && (kt.Contains(bt) || bt.Contains(kt))))
                {
                    overlap++;
                    break;
                }
            }
        }
        return overlap >= Math.Min(1, baseline.Count);
    }

    private static bool SemanticOverlapCheck(string k, HashSet<string> anchors)
    {
        var key = SemanticKey(k);
        if (string.IsNullOrEmpty(key)) return false;
        foreach (var a in anchors)
        {
            if (key == a || (key.Length >= 3 && a.Length >= 3 && (key.Contains(a) || a.Contains(key))))
                return true;
        }
        return false;
    }

    // ── 금칙어 검사 ───────────────────────────────────────────────────────────

    public static bool IsBanned(string keyword)
    {
        if (string.IsNullOrEmpty(keyword)) return true;
        var k = CleanOneKw(keyword);
        if (string.IsNullOrEmpty(k) || k.Length < 2 || k.Length > 20) return true;
        if (Regex.IsMatch(k, @"^\d+$")) return true;
        if (Ban.Contains(k)) return true;
        if (BanSubstrings.Any(bs => k.Contains(bs))) return true;
        if (Regex.IsMatch(k, @"(하다|하는|되어|됨|하기|하고|이다|입니다)$")) return true;
        if (Regex.IsMatch(k, @"(에|에서|으로|로|을|를|이|가|은|는|의|와|과)$")) return true;
        return false;
    }
}
