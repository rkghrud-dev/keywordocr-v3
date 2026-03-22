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

_COMMON_SEGMENTS = {
    "차량용", "차량", "조명", "브라켓", "브래킷", "마운트", "거치대", "홀더", "고정대",
    "무타공", "설치", "장착", "체결", "본넷", "보닛", "트렁크", "게이트", "각도조절",
    "회전형", "볼트체결", "스틸", "실버", "블랙", "작업등", "외부조명", "DIY", "튜닝",
    "틈새", "밀폐", "패드", "콘센트", "가스켓", "닭부리", "힌지", "플립", "도어",
    "트럭", "D링", "디링", "나사고리", "볼트고리", "관개", "커넥터", "호스", "연결",
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
    # category에서 비일반 명사 우선
    for w in category_words:
        if w not in _GENERIC_WORDS and len(w) >= 2:
            return w
    # product_type_correction에서 비일반 명사
    for w in type_words:
        if w not in _GENERIC_WORDS and len(w) >= 2:
            return w
    # fallback: category 첫번째
    if category_words:
        return category_words[0]
    if type_words:
        return type_words[0]
    return ""


def build_keyword_string(
    ocr_text: str,
    vision_analysis: dict[str, Any] | None,
    target_count: int = TARGET_DEFAULT,
    fallback_text: str = "",
) -> str:
    """Vision JSON + OCR 텍스트 기반 키워드 생성.

    핵심 철학:
    - 핵심어 합성어는 1개만 (PVC클램프)
    - 용도 수식은 "~용" 형태 (파이프용)
    - 기능/동작은 단독 단어 (고정, 지지, 정리)
    - 자연스러운 합성어만 유지 (흔들림방지 O, 파이프고정 → 쪼갬)
    - 사용처 단독 (천장, 벽면, 욕실)
    - 형용사 단독 (강력, 다양한)
    - 상품 특징 단독 (경량, 내식성)
    - 90~140자 목표
    """
    try:
        analysis = vision_analysis if isinstance(vision_analysis, dict) else {}
        axis = _extract_required_axes(analysis)

        # ── 축별 토큰 그룹 ──
        cat_words = _dedupe_normalized(axis["category"])
        type_words = _dedupe_normalized(axis["product_type_correction"])
        core_words = _dedupe_normalized(
            axis["category"] + axis["product_type_correction"]
        )

        # 핵심 상품어 (클램프, 호스 등)
        base_core = _pick_base_core(cat_words, type_words)

        # 용도/대상 단어: 핵심어와 결합해 "~용" 또는 합성어 생성
        purpose_words = [w for w in core_words
                         if w != base_core and w not in _GENERIC_WORDS]

        # 기능/동작: 단독 사용 (고정, 지지, 정리, 방지 등)
        functions = _dedupe_normalized(
            axis["problem_solving_keyword"]
            + axis["usage_purpose"]
            + axis["benefit_keywords"]
        )

        # 사용처: 단독 사용 (천장, 벽면, 욕실 등)
        locations = _dedupe_normalized(
            axis["usage_location"]
            + axis["space_keywords"]
        )

        # 부스트 키워드 (끼움형, 나사고정 등)
        boost_terms = _dedupe_normalized(
            axis["installation_keywords"] + axis["longtail_candidates"]
        )

        # 색상은 옵션이므로 키워드에서 제외

        out: list[str] = []
        seen: set[str] = set()

        def _try_add(token: str) -> bool:
            t = _normalize_token(token)
            if not t or len(t) < 2 or t.lower() in _STOPWORDS:
                return False
            # 불필요 패턴 제거
            if t.endswith("소재") and len(t) > 2:
                return False
            # 색상은 옵션이라 키워드 불필요
            if re.fullmatch(r"(흰색|검정|검은색|화이트|블랙|실버|은색|회색|그레이|빨간색|파란색|노란색|녹색|흰색화이트|블랙검정)", t):
                return False
            if len(t) > _MAX_TOKEN_LEN:
                return False
            if t in seen:
                return False
            seen.add(t)
            out.append(t)
            return True

        def _char_len() -> int:
            return sum(len(t) for t in out) + max(0, len(out) - 1)

        def _is_full() -> bool:
            return _char_len() >= _MAX_CHAR_LIMIT

        # ── Phase 1: 대표 합성어 1개 (PVC클램프 등) ──
        compound_mod_used = ""
        if base_core and purpose_words:
            best_mod = purpose_words[0]
            compound = best_mod + base_core
            if len(compound) <= _MAX_TOKEN_LEN:
                if _try_add(compound):
                    compound_mod_used = best_mod

        # ── Phase 2: 용도 "~용" (파이프용, 배관용 등) — 합성어에 쓰인 것 제외 ──
        for pw in purpose_words:
            if _is_full():
                break
            p = _normalize_token(pw)
            if not p or p == base_core or p == compound_mod_used:
                continue
            # 이미 "용"으로 끝나면 그대로, 4자 이상 합성어에는 "용" 안 붙임
            if p.endswith("용"):
                _try_add(p)
            elif len(p) <= 3 and len(p + "용") <= _MAX_TOKEN_LEN:
                _try_add(p + "용")

        # ── Phase 3: 사용처 단독 (천장, 벽면, 욕실 등) ──
        for loc in locations:
            if _is_full():
                break
            _try_add(loc)

        # ── Phase 4: 기능/동작 단독 (고정, 지지, 정리 등) ──
        # 합성어(파이프고정 등)는 분해해서 개별 단어로 추가
        # 알려진 단어 사전 구성 (분해용)
        _ACTION_VOCAB = {
            "고정", "정리", "방지", "설치", "시공", "조절", "장착",
            "연결", "분리", "교체", "보호", "차단", "밀봉", "강화",
            "지지", "수납", "거치", "탈착", "간편", "강력", "다양한",
            "흔들림", "사이즈", "내구성", "내식성",
        }
        _known_words = set(core_words + locations + purpose_words) | _ACTION_VOCAB
        for f in functions:
            if _is_full():
                break
            fn = _normalize_token(f)
            if not fn:
                continue
            # 3자 이하면 단일 개념어 → 그대로
            if len(fn) <= 3:
                _try_add(fn)
                continue
            # 4자 이상이면 분해 시도
            split = _split_compound_once(fn, _known_words)
            if len(split) >= 2:
                for part in split:
                    if not _is_full():
                        _try_add(part)
            else:
                _try_add(fn)

        # ── Phase 5: 부스트 키워드 (끼움형, 나사고정 등) ──
        for b in boost_terms:
            if _is_full():
                break
            bn = _normalize_token(b)
            if bn and (not base_core or base_core not in bn):
                _try_add(bn)

        # ── Phase 6: OCR 보충 ──
        for t in _tokenize_ocr(ocr_text):
            if _is_full():
                break
            _try_add(t)

        # ── Phase 7: 상품명 보충 ──
        if fallback_text and not _is_full():
            for m in re.findall(r"[0-9A-Za-z가-힣]{2,14}", _normalize_token(fallback_text)):
                if _is_full():
                    break
                if len(m) >= 2:
                    _try_add(m)

        # ── 글자수 보강 ──
        if _char_len() < _MIN_CHAR_TARGET:
            for cw in core_words:
                _try_add(cw)
                if _char_len() >= _MIN_CHAR_TARGET:
                    break

        if _char_len() < _MIN_CHAR_TARGET:
            for seg in sorted(_COMMON_SEGMENTS, key=lambda x: -len(x)):
                _try_add(seg)
                if _char_len() >= _MIN_CHAR_TARGET:
                    break

        return " ".join(out).strip()
    except Exception:
        return ""
