from __future__ import annotations

import re
from typing import Any


TARGET_DEFAULT = 20

_STOPWORDS = {
    "및", "또는", "에서", "으로", "같은", "관련", "용", "용도", "제품", "상품", "기타",
    "가능", "활용", "사용", "적용", "구성", "기본", "일반", "세트", "단품", "옵션",
    "빠른", "유지", "작업",
    # 조사/어미/잔여물 필터
    "통해", "위해", "위한", "따른", "위의", "후에", "대한", "포함", "약간",
    "하여", "있는", "없는", "있음", "없음", "됨", "함",
    "위치", "취급", "용이", "우수", "뛰어난", "다양한", "선택", "추천",
    "가능한", "적합한", "필요한", "특징", "장점", "효과", "방법",
    "소형", "대형", "중형", "라인업", "사이즈",
    # 재질/구조 설명 잔여물
    "경질", "연질", "고정형", "이동형", "겸용",
    # 노이즈
    "내부", "외부", "다사이즈",
    "시리즈", "세트", "용품", "배관용품",
    # 영문 잔여물
    "shaped", "clamp", "type", "style", "pipe", "tube", "hose",
    "profile", "point", "option", "con",
}

_ODD_SINGLE_WORDS = {
    "펼침", "접힘", "열림", "닫힘", "분리", "조절", "가능", "추천", "강화", "완성",
}

_SYNONYM_GROUPS = {
    "bracket_mount": ["브라켓", "브래킷", "마운트", "거치대", "홀더", "고정대", "클램프"],
    "install": ["설치", "장착", "체결", "부착", "고정"],
    "no_drill": ["무타공", "무천공", "타공없음"],
    "angle_rotate": ["각도조절", "각도조정", "회전", "회전형"],
    "location_vehicle": ["본넷", "보닛", "트렁크", "게이트"],
    "material": ["스틸", "스텐", "스테인리스", "철제", "알루미늄"],
    "color_black": ["블랙", "검정", "검은색"],
    "color_silver": ["실버", "은색"],
}



def _clean_text(s: Any) -> str:
    s = re.sub(r"[\t\r\n]+", " ", str(s or ""))
    s = re.sub(r"[|,;/·•‧]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _normalize_token(tok: str) -> str:
    tok = _clean_text(tok)
    tok = re.sub(r"[^0-9A-Za-z가-힣\-\+ ]", "", tok)
    tok = re.sub(r"\s+", " ", tok).strip()
    return tok


def _to_list(v: Any) -> list[str]:
    if v is None:
        return []
    if isinstance(v, list):
        return [_clean_text(x) for x in v if _clean_text(x)]
    if isinstance(v, dict):
        if "value" in v:
            return _to_list(v.get("value"))
        out = []
        for vv in v.values():
            out.extend(_to_list(vv))
        return out
    s = _clean_text(v)
    return [s] if s else []


def _extract_field(analysis: dict[str, Any], section: str, key: str) -> list[str]:
    sec = analysis.get(section, {})
    if not isinstance(sec, dict):
        return _to_list(sec)
    return _to_list(sec.get(key))


def _extract_required_axes(analysis: dict[str, Any]) -> dict[str, list[str]]:
    return {
        "category": _extract_field(analysis, "core_identity", "category"),
        "product_type_correction": _extract_field(analysis, "core_identity", "product_type_correction"),
        "structure": _extract_field(analysis, "core_identity", "structure"),
        "material_visual": _extract_field(analysis, "core_identity", "material_visual"),
        "color": _extract_field(analysis, "core_identity", "color"),
        "mount_type": _extract_field(analysis, "installation_and_physical", "mount_type"),
        "installation_method": _extract_field(analysis, "installation_and_physical", "installation_method"),
        "usage_location": _extract_field(analysis, "usage_context", "usage_location"),
        "usage_purpose": _extract_field(analysis, "usage_context", "usage_purpose"),
        "target_user": _extract_field(analysis, "usage_context", "target_user"),
        "usage_scenario": _extract_field(analysis, "usage_context", "usage_scenario"),
        "indoor_outdoor": _extract_field(analysis, "usage_context", "indoor_outdoor"),
        "primary_function": _extract_field(analysis, "functional_inference", "primary_function"),
        "problem_solving_keyword": _extract_field(analysis, "functional_inference", "problem_solving_keyword"),
        "convenience_feature": _extract_field(analysis, "functional_inference", "convenience_feature"),
        "installation_keywords": _extract_field(analysis, "search_boost_elements", "installation_keywords"),
        "space_keywords": _extract_field(analysis, "search_boost_elements", "space_keywords"),
        "benefit_keywords": _extract_field(analysis, "search_boost_elements", "benefit_keywords"),
        "longtail_candidates": _extract_field(analysis, "search_boost_elements", "longtail_candidates"),
    }


def _core_form(tok: str) -> str:
    t = tok.lower()
    t = t.replace("차량용", "차량")
    t = t.replace("브래킷", "브라켓")
    t = t.replace("디링", "d링")
    t = re.sub(r"(용|형|식)$", "", t)
    return t


def _syn_group(tok: str) -> str | None:
    t = _core_form(tok)
    for g, words in _SYNONYM_GROUPS.items():
        for w in words:
            if _core_form(w) == t:
                return g
    return None


def _split_compound_once(term: str, vocab: set[str]) -> list[str]:
    s = term.strip()
    if not s or " " in s:
        return [s] if s else []
    if not re.search(r"[가-힣]", s):
        return [s]
    n = len(s)
    dp: list[tuple[int, list[str]] | None] = [None] * (n + 1)
    dp[0] = (0, [])
    for i in range(n):
        if dp[i] is None:
            continue
        _, cur_tokens = dp[i]
        for j in range(i + 2, min(n, i + 8) + 1):
            piece = s[i:j]
            if piece in vocab:
                cand_score = (dp[i][0] + len(piece))
                cand_tokens = cur_tokens + [piece]
                old = dp[j]
                if old is None or cand_score > old[0] or (cand_score == old[0] and len(cand_tokens) < len(old[1])):
                    dp[j] = (cand_score, cand_tokens)
    if dp[n] and len(dp[n][1]) >= 2 and "".join(dp[n][1]) == s:
        return dp[n][1]
    return [s]


def _expand_term(term: str, vocab: set[str]) -> list[str]:
    t = _normalize_token(term)
    if not t:
        return []
    outs = [t]
    if " " in t:
        outs.extend([x for x in t.split(" ") if x])
    split_once = _split_compound_once(t, vocab)
    if len(split_once) >= 2:
        outs.extend(split_once)
        if split_once[0] == "차량용":
            outs.append("차량")
    return [x for x in outs if _normalize_token(x)]


def _tokenize_ocr(ocr_text: str) -> list[str]:
    txt = _normalize_token(ocr_text)
    if not txt:
        return []
    words = re.findall(r"[0-9A-Za-z가-힣]+", txt)
    out = []
    for w in words:
        if len(w) < 2 or len(w) > 14:
            continue
        if w.lower() in _STOPWORDS:
            continue
        if re.fullmatch(r"\d+", w):
            continue
        out.append(w)
    return out


def _is_odd_token(tok: str) -> bool:
    t = _core_form(tok)
    if t in _ODD_SINGLE_WORDS:
        return True
    if len(t) <= 2 and re.search(r"(함|됨)$", t):
        return True
    return False


def _collect_bucket_tokens(axis: dict[str, list[str]], ocr_text: str) -> dict[str, list[str]]:
    front_raw = (
        axis["category"]
        + axis["product_type_correction"]
        + axis["primary_function"]
        + axis["mount_type"]
        + axis["installation_method"]
        + axis["usage_location"]
        + axis["installation_keywords"]
        + axis["space_keywords"]
    )
    middle_raw = (
        axis["structure"]
        + axis["usage_purpose"]
        + axis["target_user"]
        + axis["problem_solving_keyword"]
        + axis["convenience_feature"]
        + axis["benefit_keywords"]
    )
    back_raw = (
        axis["material_visual"]
        + axis["color"]
        + axis["usage_scenario"]
        + axis["indoor_outdoor"]
        + axis["longtail_candidates"]
    )
    ocr_raw = _tokenize_ocr(ocr_text)
    return {
        "front": front_raw,
        "middle": middle_raw,
        "back": back_raw,
        "ocr": ocr_raw,
    }


def _add_tokens(
    source: list[str],
    out: list[str],
    seen_core: set[str],
    group_count: dict[str, int],
    vocab: set[str],
    group_limit: int = 2,
) -> None:
    for s in source:
        for cand in _expand_term(s, vocab):
            for part in cand.split(" "):
                tok = _normalize_token(part)
                if not tok:
                    continue
                if tok.lower() in _STOPWORDS or _is_odd_token(tok):
                    continue
                core = _core_form(tok)
                if not core:
                    continue
                grp = _syn_group(tok)
                if grp:
                    used = group_count.get(grp, 0)
                    if used >= group_limit:
                        continue
                if core in seen_core:
                    continue
                # 같은 그룹 내 과밀 완화: 대표어 + 보조어 1개까지만 허용
                if grp and group_count.get(grp, 0) >= 1:
                    if len(tok) <= 2:
                        continue
                seen_core.add(core)
                out.append(tok)
                if grp:
                    group_count[grp] = group_count.get(grp, 0) + 1


_MIN_CHAR_TARGET = 90
_MAX_CHAR_LIMIT = 140
_MAX_TOKEN_LEN = 7       # 토큰 최대 글자수 (합성어끼리 연결 방지)
_MAX_CORE_COMPOUNDS = 3  # 핵심어 합성어 최대 개수 (클램프 중복 제한)

# B마켓 글자수 규칙
_MIN_CHAR_TARGET_B = 63
_MAX_CHAR_LIMIT_B = 98


_JOSA_SUFFIXES = re.compile(r"(을|를|에|의|은|는|가|로|와|과|에서|으로|하여|에도|까지)$")


def _strip_josa(w: str) -> str:
    """한글 단어 끝의 조사를 제거."""
    if not re.search(r"[가-힣]", w):
        return w
    cleaned = _JOSA_SUFFIXES.sub("", w)
    return cleaned if len(cleaned) >= 2 else w


def _dedupe_normalized(items: list[str]) -> list[str]:
    """순서 유지, 정규화 후 중복 제거. 공백 포함 항목은 개별 단어로 분리."""
    seen: set[str] = set()
    out: list[str] = []
    for item in items:
        t = _normalize_token(item)
        if not t:
            continue
        # 공백 포함 → 개별 단어로 분리
        words = t.split() if " " in t else [t]
        for w in words:
            w = _strip_josa(w.strip())
            if not w or len(w) < 2 or w in seen:
                continue
            if w.lower() in _STOPWORDS:
                continue
            seen.add(w)
            out.append(w)
    return out


_GENERIC_WORDS = {
    "고정", "자재", "부품", "소재", "재료", "도구", "공구",
    "용품", "소품", "제품", "상품", "부자재", "배관자재", "설비",
}


def _pick_base_core(category_words: list[str], type_words: list[str]) -> str:
    """핵심 상품어 선택. 일반 명사(고정, 배관 등)를 제외하고 대표 상품명을 반환."""
    # 한국 상품명은 보통 마지막 명사가 제품 헤드라서 뒤에서부터 탐색한다.
    for w in reversed(category_words):
        if w not in _GENERIC_WORDS and len(w) >= 2:
            return w
    for w in reversed(type_words):
        if w not in _GENERIC_WORDS and len(w) >= 2:
            return w
    if category_words:
        return category_words[-1]
    if type_words:
        return type_words[-1]
    return ""


def _should_compact_core_phrase(modifier: str, base_core: str) -> bool:
    mod = _normalize_token(modifier)
    base = _normalize_token(base_core)
    if not mod or not base:
        return False
    # 한글 일반 명사는 붙이지 않고 띄어써야 자연스럽다. PVC/304 같은 짧은 표기만 합성 허용.
    if re.fullmatch(r"[A-Za-z0-9\-\+]{1,4}", mod):
        return True
    return False



_HEAD_SUFFIXES = [
    "브라켓", "브래킷", "거치대", "받침대", "지지대", "홀더", "클립",
    "커넥터", "조인트", "클램프", "노즐", "테이프", "커버", "마개",
    "캡", "패드", "브러시", "필터", "밸브", "후크", "고리",
    "볼트", "너트", "핀", "호스", "파이프", "케이블",
]

_NAME_CONTEXT_WORDS = [
    "하수구", "배수구", "세면대", "싱크대", "욕실", "주방", "차량", "자동차",
    "스위치", "호스", "배관", "파이프", "관개", "정원", "조명", "전선",
    "케이블", "벽면", "천장", "창문", "문", "트렁크", "본넷", "밑창",
]

_NAME_ACTION_WORDS = [
    "세척", "고정", "연결", "장착", "설치", "분사", "배수", "누수",
    "방지", "보호", "정리", "거치", "교체", "수리", "지지", "보수",
]

_CONTEXT_SUFFIX_HEADS = {
    "핀", "브라켓", "브래킷", "거치대", "받침대", "지지대", "홀더",
    "클립", "커넥터", "조인트", "클램프", "테이프", "커버", "마개",
    "캡", "패드", "후크", "고리", "볼트", "너트",
}

_ALLOWED_COMPOUNDS = {"고정핀"}

_ACTION_REDUNDANT_HEADS = {"커넥터", "조인트"}


def _extract_name_only_tokens(fallback_text: str, market: str = "A") -> list[str]:
    raw = _normalize_token(re.sub(r"GS\d{7}[A-Z]?", " ", str(fallback_text or ""))).strip()
    if not raw:
        return []

    raw_parts = [
        part for part in raw.split()
        if part and not re.search(r"\d", part)
    ]
    raw = " ".join(raw_parts).strip() or raw

    joined = raw.replace(" ", "")
    out: list[str] = []
    seen: set[str] = set()

    def _push(token: str) -> None:
        t = _normalize_token(token)
        if not t or len(t) < 2 or len(t) > _MAX_TOKEN_LEN or t in seen:
            return
        seen.add(t)
        out.append(t)

    def _context_token(word: str, use_suffix: bool) -> str:
        token = _normalize_token(word)
        if not token:
            return ""
        if use_suffix and not token.endswith("용"):
            candidate = token + "용"
            if len(candidate) <= _MAX_TOKEN_LEN:
                return candidate
        return token

    head = ""
    for suffix in _HEAD_SUFFIXES:
        if joined.endswith(suffix):
            head = suffix
            break
    stem = joined[:-len(head)] if head else joined

    matched_contexts = [word for word in _NAME_CONTEXT_WORDS if word in stem]
    matched_actions = [word for word in _NAME_ACTION_WORDS if word in stem]
    primary_action = matched_actions[0] if matched_actions else ""

    compact_token = ""
    if primary_action and head:
        candidate = primary_action + head
        if candidate in _ALLOWED_COMPOUNDS and len(candidate) <= _MAX_TOKEN_LEN:
            compact_token = candidate

    keep_action = bool(primary_action) and not compact_token and head not in _ACTION_REDUNDANT_HEADS
    use_context_suffix = bool(head in _CONTEXT_SUFFIX_HEADS and (compact_token or (not keep_action)))

    if matched_contexts:
        _push(_context_token(matched_contexts[0], use_context_suffix))
        for extra_context in matched_contexts[1:]:
            _push(extra_context)

    if compact_token:
        _push(compact_token)
    else:
        if keep_action:
            _push(primary_action)
        if head:
            _push(head)
        elif primary_action:
            _push(primary_action)

    if not out:
        for part in raw.split():
            _push(part)
    if not out and joined:
        _push(joined)

    return out[:4]


def build_keyword_string(
    ocr_text: str,
    vision_analysis: dict[str, Any] | None,
    target_count: int = TARGET_DEFAULT,
    fallback_text: str = "",
    market: str = "A",
) -> str:
    """Vision/OCR 증거만 사용해 키워드 문자열을 조립한다."""
    try:
        ocr_text = "" if "OCR 텍스트 없음" in str(ocr_text or "") else str(ocr_text or "")
        target_count = max(1, int(target_count or TARGET_DEFAULT))

        if market == "B":
            min_char = _MIN_CHAR_TARGET_B
            max_char = _MAX_CHAR_LIMIT_B
        else:
            min_char = _MIN_CHAR_TARGET
            max_char = _MAX_CHAR_LIMIT

        analysis = vision_analysis if isinstance(vision_analysis, dict) else {}
        axis = _extract_required_axes(analysis)

        cat_words = _dedupe_normalized(axis["category"])
        type_words = _dedupe_normalized(axis["product_type_correction"])
        core_words = _dedupe_normalized(axis["category"] + axis["product_type_correction"])
        base_core = _pick_base_core(cat_words, type_words)
        name_identity_tokens = _extract_name_only_tokens(fallback_text, market=market)

        generic_words = _GENERIC_WORDS | {
            "고정", "설치", "연결", "장착", "정리", "방지", "보호", "부품", "용품", "도구", "세트",
            "옵션", "사용", "실내", "실외", "간편", "강력", "다양한", "사이즈", "작업효율",
            "편리", "휴대용", "다용도", "구조", "형태",
        }

        def semantic_key(token: str) -> str:
            return _core_form(_normalize_token(token))

        def is_generic_token(token: str) -> bool:
            tok = _normalize_token(token)
            key = semantic_key(tok)
            if not tok or not key:
                return True
            if key in generic_words or tok.lower() in _STOPWORDS:
                return True
            if any(key.endswith(suffix) for suffix in ("용품", "부품", "도구", "세트", "옵션")):
                return True
            return False

        def has_overlap(token: str, refs: list[str]) -> bool:
            key = semantic_key(token)
            if not key:
                return False
            for ref in refs:
                ref_key = semantic_key(ref)
                if not ref_key:
                    continue
                if key == ref_key:
                    return True
                if len(key) >= 3 and len(ref_key) >= 3 and (key in ref_key or ref_key in key):
                    return True
            return False

        def dedupe_semantic(items: list[str]) -> list[str]:
            out = []
            seen = set()
            for item in items:
                tok = _normalize_token(item)
                key = semantic_key(tok)
                if not tok or not key or key in seen:
                    continue
                seen.add(key)
                out.append(tok)
            return out

        identity_terms = []
        if base_core and not is_generic_token(base_core):
            identity_terms.append(base_core)
        identity_terms.extend(core_words)
        identity_terms.extend(name_identity_tokens)
        identity_terms = [t for t in dedupe_semantic(identity_terms) if not is_generic_token(t)]

        usage_terms = []
        for raw in core_words:
            tok = _normalize_token(raw)
            if not tok or tok == base_core or is_generic_token(tok):
                continue
            usage_terms.append(tok)
            if len(tok) <= 3 and len(tok + "용") <= _MAX_TOKEN_LEN:
                usage_terms.append(tok + "용")
        usage_terms.extend(_dedupe_normalized(axis["usage_location"] + axis["space_keywords"]))
        usage_terms = [t for t in dedupe_semantic(usage_terms) if not is_generic_token(t)]

        _ACTION_VOCAB = {
            "고정", "정리", "방지", "설치", "시공", "조절", "장착",
            "연결", "분리", "교체", "보호", "차단", "밀봉", "강화",
            "지지", "수납", "거치", "탈착", "흔들림", "내구성", "내식성",
        }
        known_words = set(identity_terms + usage_terms) | _ACTION_VOCAB
        function_terms = []
        raw_functions = _dedupe_normalized(
            axis["problem_solving_keyword"] + axis["usage_purpose"] + axis["benefit_keywords"]
        )
        for raw in raw_functions:
            fn = _normalize_token(raw)
            if not fn or is_generic_token(fn):
                continue
            split = _split_compound_once(fn, known_words) if len(fn) > 3 else [fn]
            for part in split:
                part = _normalize_token(part)
                if part and not is_generic_token(part):
                    function_terms.append(part)
        function_terms = dedupe_semantic(function_terms)

        evidence_refs = identity_terms + usage_terms + function_terms
        boost_terms = []
        raw_boosts = _dedupe_normalized(axis["installation_keywords"] + axis["longtail_candidates"])
        for raw in raw_boosts:
            expanded = _expand_term(raw, known_words)
            for part in expanded:
                part = _normalize_token(part)
                if not part or is_generic_token(part):
                    continue
                if not has_overlap(part, evidence_refs):
                    continue
                boost_terms.append(part)
        boost_terms = dedupe_semantic(boost_terms)

        evidence_refs = identity_terms + usage_terms + function_terms + boost_terms
        ocr_terms = []
        for token in _tokenize_ocr(ocr_text):
            if is_generic_token(token):
                continue
            if not has_overlap(token, evidence_refs):
                continue
            ocr_terms.append(token)
        ocr_terms = dedupe_semantic(ocr_terms)

        fallback_terms = [t for t in dedupe_semantic(name_identity_tokens) if not is_generic_token(t)]

        out: list[str] = []
        seen: set[str] = set()

        def _char_len() -> int:
            return sum(len(t) for t in out) + max(0, len(out) - 1)

        def _is_full() -> bool:
            return len(out) >= target_count or _char_len() >= max_char

        def _try_add(token: str) -> bool:
            t = _normalize_token(token)
            key = semantic_key(t)
            if not t or not key or len(t) < 2 or t.lower() in _STOPWORDS:
                return False
            if t.endswith("소재") and len(t) > 2:
                return False
            if re.fullmatch(r"(흰색|검정|검은색|화이트|블랙|실버|은색|회색|그레이|빨간색|파란색|노란색|녹색|흰색화이트|블랙검정)", t):
                return False
            if len(t) > _MAX_TOKEN_LEN:
                return False
            if key in seen:
                return False
            seen.add(key)
            out.append(t)
            return True

        for group in (identity_terms, usage_terms, function_terms, boost_terms, ocr_terms, fallback_terms):
            for token in group:
                if _is_full():
                    break
                _try_add(token)
            if _is_full():
                break

        if _char_len() < min_char:
            remaining = dedupe_semantic(identity_terms + usage_terms + function_terms + boost_terms + ocr_terms + fallback_terms)
            for token in remaining:
                if _is_full():
                    break
                _try_add(token)

        return " ".join(out[:target_count]).strip()
    except Exception:
        return ""
