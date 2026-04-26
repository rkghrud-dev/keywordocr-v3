# -*- coding: utf-8 -*-



"""



전처리(old) + OCR(고급) + GPT키워드(정렬·루트중복제한) + 대표이미지(대표 배경색 일관 적용)



+ 초미세 변형(회전·이동·스케일·좌우반전) + 검색어설정(정확히 20개: 롱테일10 + 네이버 PC5/MO5) 통합 풀코드



(+ 네이버 Invalid Parameter 방지: hintKeywords 안전 생성기, dtype 경고 해결, 숫자 접두/접미 제거)







- 검색어설정 로직(업데이트):



  1) 롱테일키워드 10개: 상품명 기반 생성(GPT 또는 휴리스틱), 온토픽(앵커) + 상품명기반 키워드(baseline) 일관성 필터



  2) 네이버키워드 PC 상위 5개, 모바일 상위 5개: 월평균 CTR 5% 초과만 우선 채택(없으면 CTR 낮은 차순위에서 보충)



     * PC/MO 중복은 한 번만 쓰고, 모자란 수량은 각 플랫폼에서 추가 선별



  3) 최종 20개로 정확히 맞춤(중복 제거/보충 포함), 금칙어·길이(2~12자)·숫자 앞뒤 제거, 앵커/기준키워드 일관성 검증



"""







import os, re, io, traceback, threading, random, time, hmac, json, hashlib, base64



from datetime import datetime



from concurrent.futures import ThreadPoolExecutor, as_completed



from collections import defaultdict







import numpy as np



import pandas as pd



import chardet



import requests







from PIL import Image, ImageOps, ImageFilter, UnidentifiedImageError, ImageEnhance



from PIL import ImageChops



import pytesseract







try:



    import openpyxl



except Exception:



    pass







# === OpenAI (선택) ===



try:



    from openai import OpenAI



except ImportError:



    OpenAI = None



# === Anthropic / Claude (선택) ===

try:

    from .anthropic_wrapper import AnthropicClientWrapper

except ImportError:

    AnthropicClientWrapper = None



from .env_loader import ensure_env_loaded, get_env, key_file_candidates











# ------------------



def _app_root() -> str:

    here = os.path.abspath(os.path.dirname(__file__))

    return os.path.abspath(os.path.join(here, '..', '..'))



def _read_text_from_candidates(paths):

    for p in paths:

        try:

            with open(p, 'r', encoding='utf-8') as f:

                return f.read().strip()

        except FileNotFoundError:

            continue

    return None



# ------------------ 공통 유틸 ------------------







def load_api_key():

    base = _app_root()

    ensure_env_loaded(os.path.join(base, ".env"))

    env_key = get_env("OPENAI_API_KEY")

    if env_key:

        return env_key

    return _read_text_from_candidates(key_file_candidates("api_key.txt"))





def load_anthropic_api_key():

    base = _app_root()

    ensure_env_loaded(os.path.join(base, ".env"))

    env_key = get_env("ANTHROPIC_API_KEY")

    if env_key:

        return env_key

    return _read_text_from_candidates(key_file_candidates("anthropic_api_key.txt"))





def _is_claude_model(model_name: str) -> bool:

    """모델 이름이 Claude 계열인지 판별"""

    if not model_name:

        return False

    return any(k in model_name.lower() for k in ("claude", "haiku", "sonnet", "opus"))





def _create_client(model_name: str = ""):

    """모델 이름에 따라 적절한 클라이언트 반환 (OpenAI 또는 Anthropic 래퍼)"""

    if _is_claude_model(model_name):

        key = load_anthropic_api_key()

        if AnthropicClientWrapper and key:

            return AnthropicClientWrapper(api_key=key)

    key = load_api_key()

    if OpenAI and key:

        return OpenAI(api_key=key)

    return None





def _resolve_model_client(model_name: str = "", explicit_client=None):

    """모델명 기준으로 맞는 클라이언트를 우선 선택."""

    return _create_client(model_name) or explicit_client or client





_API_KEY = load_api_key()

_ANTHROPIC_API_KEY = load_anthropic_api_key()



# 기본 클라이언트: Anthropic 키가 있으면 Claude 우선, 없으면 OpenAI

if AnthropicClientWrapper and _ANTHROPIC_API_KEY:

    client = AnthropicClientWrapper(api_key=_ANTHROPIC_API_KEY)

elif OpenAI and _API_KEY:

    client = OpenAI(api_key=_API_KEY)

else:

    client = None





def refresh_openai_client():

    global _API_KEY, _ANTHROPIC_API_KEY, client

    _ANTHROPIC_API_KEY = load_anthropic_api_key()

    _API_KEY = load_api_key()

    if AnthropicClientWrapper and _ANTHROPIC_API_KEY:

        client = AnthropicClientWrapper(api_key=_ANTHROPIC_API_KEY)

        return True

    if OpenAI and _API_KEY:

        client = OpenAI(api_key=_API_KEY)

        return True

    return False







def setup_tesseract(custom_path=None):



    if custom_path and os.path.isfile(custom_path):



        pytesseract.pytesseract.tesseract_cmd = custom_path



        return custom_path



    for path in [r"C:\Program Files\Tesseract-OCR\tesseract.exe",



                 r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"]:



        if os.path.isfile(path):



            pytesseract.pytesseract.tesseract_cmd = path



            return path



    return None







STOPWORDS = {



    "정품","국내발송","무료배송","행사","특가","세일","인기","추천","최고","인증","브랜드",



    "프리미엄","신제품","베스트","전용","대박","핫딜","당일","당일발송","빠른","빠른배송","무료",



    "고급","기획","사은","증정","정가","한정","행사상품","총알","가성비","가심비","필수","완전",



    "생활용품","생활소품","생활가전",



    "국내창고","중국산","해외직구","해외배송","국내배송","당일출고","즉시발송",

    "무료반품","교환가능","환불","AS","품질보증","제조사","공장직영","도매",

    "패키지","구성품","세트구성","본품","단품",



    # OCR 메타/설명 단어 (키워드가 아닌 것들)

    "상품명","상품","제품","제품명","모델명","모델","브랜드명",

    "재질","소재","재료","원단","원재료",

    "사이즈","크기","규격","치수","높이","너비","폭","길이","두께","지름","직경",

    "색상","컬러","색깔",

    "특징","장점","특성","기능설명",

    "용도","적용","적용범위",

    "무게","중량","용량",

    "수량","단위","개입",

    "주의사항","참고","안내","문의","확인","참조","유의",

    "포함","별도","불포함","옵션","선택",

    "이미지","사진","상세","페이지","화면","참고사진",

    "상품은","제품은","본제품","해당제품","본상품","해당상품"



}



BAN_SUBSTRINGS = {"까지","최대","절단폭"}  # substring 매칭 (단어 안에 포함되면 제거)

BAN_EXACT = {"가능","안정적","안정적인","안정","사용법","사용","탑재"}  # 정확 매칭만



NUM_UNIT_PAT = re.compile(r"^\d+(매|mm|cm|g|kg|ml|l|L|개|팩|장|세트|set)?$")







# 사이즈/용량 계열 제외



SIZE_WORDS = {"대형","중형","소형","미니","대","중","소","대용량","소용량","초소형","초대형"}



STOPWORDS |= SIZE_WORDS





# ── 사용자 정의 제외 단어 (user_stopwords.json) ──



_USER_STOPWORDS_PATH = os.path.join(_app_root(), "user_stopwords.json")



def load_user_stopwords() -> set:

    """user_stopwords.json에서 사용자 제외 단어 로드"""

    try:

        with open(_USER_STOPWORDS_PATH, "r", encoding="utf-8") as f:

            data = json.load(f)

            if isinstance(data, list):

                return set(data)

    except (FileNotFoundError, json.JSONDecodeError):

        pass

    return set()



def save_user_stopwords(words) -> None:

    """사용자 제외 단어를 JSON으로 저장"""

    with open(_USER_STOPWORDS_PATH, "w", encoding="utf-8") as f:

        json.dump(sorted(words), f, ensure_ascii=False, indent=2)



def merge_user_stopwords() -> None:

    """사용자 제외 단어를 STOPWORDS에 합치기"""

    global STOPWORDS

    user_words = load_user_stopwords()

    if user_words:

        STOPWORDS |= user_words



# 모듈 로드 시 자동 합치기

merge_user_stopwords()





def normalize_space(s: str) -> str:



    return " ".join(str(s).split())







def _allow_digit_token(w: str) -> bool:



    if NUM_UNIT_PAT.match(w): return False



    if re.search(r"\d", w) and re.search(r"[가-힣]", w): return True



    if re.search(r"\d", w) and not re.search(r"[가-힣]", w): return False



    return True







def _filter_tokens_drop_digits(tokens: list) -> list:



    return [t for t in tokens if not re.search(r"\d", t)]







def postprocess_keywords_tokens(text: str, max_words=22, max_len=120):



    if not text: return "", []



    text = re.sub(r"[^0-9A-Za-z가-힣\s]", " ", str(text))



    text = normalize_space(text)



    words, seen = [], set()



    for w in text.split():



        if any(b in w for b in BAN_SUBSTRINGS): continue



        if w in BAN_EXACT: continue



        if w in STOPWORDS: continue



        if not _allow_digit_token(w): continue



        if w not in seen:



            seen.add(w); words.append(w)



        if len(words) >= max_words: break



    out = normalize_space(" ".join(words))[:max_len].strip()



    return out, words







def extract_text_from_html(html: str) -> str:



    if not html or not isinstance(html, str): return ""



    text = re.sub(r"<script.*?</script>", " ", html, flags=re.DOTALL|re.IGNORECASE)



    text = re.sub(r"<style.*?</style>", " ", text, flags=re.DOTALL|re.IGNORECASE)



    text = re.sub(r"<[^>]+>", " ", text)



    text = re.sub(r"&nbsp;|&amp;|&lt;|&gt;|&quot;|&#39;", " ", text)



    text = re.sub(r"[^0-9가-힣\s]", " ", text)



    return normalize_space(text)







def extract_img_srcs(html: str, max_images: int = 3):



    if not html or not isinstance(html, str): return []



    srcs = re.findall(r'<img[^>]+src=[\'"]([^\'"]+)[\'"]', html, flags=re.IGNORECASE)



    srcs = [s for s in srcs if s.lower().startswith("http")]



    seen, uniq = set(), []



    for s in srcs:



        if s not in seen:



            seen.add(s); uniq.append(s)



        if len(uniq) >= max_images: break



    return uniq







# ------------------ 키워드 정렬(별칭→특징→용도) ------------------







USAGE_TOKENS = {"사무실","가정","학교","카페","주방","욕실","세면대","싱크대","원룸","자취","야외","캠핑","차량","창고","정원","벽면","실내","여행","낚시","가방","자동차","학생","선물","인테리어","데일리","사무용","포인트","악세사리","결혼식","생일","졸업","입학","유아","어린이","남성","여성","커플","오피스"}



FEATURE_TOKENS = {



    "귀여운","심플","모던","내구성","고강도","강화","튼튼","견고","방수","생활방수","방진",



    "내열","내오염","미끄럼방지","위생","친환경","저소음","경량","휴대","호환","범용","고정밀",



    "투명","불투명","무광","유광","부드러운","거친","탄성","신축","항균","소음차단","충격흡수",



    "충전","자동","점등","에너지","절약","태양광","LED","풍력등","무드등","조명","램프",



    "감성","빈티지","레트로","미니멀","고급","프리미엄","세련","깔끔","화려한","예쁜","럭셔리","캐주얼"



}



ALIAS_HINTS = set()







def reorder_kw_tokens(tokens: list, base_name: str) -> list:



    if not tokens: return tokens



    toks = [t for t in tokens if t and t not in STOPWORDS and t not in SIZE_WORDS]



    base_tokens = set(normalize_space(re.sub(r"[^0-9가-힣\sA-Za-z]", " ", base_name)).split())



    alias, feats, usage, others = [], [], [], []



    for t in toks:



        if t in base_tokens: continue



        if t in ALIAS_HINTS: alias.append(t)



        elif t in FEATURE_TOKENS: feats.append(t)



        elif t in USAGE_TOKENS: usage.append(t)



        else: others.append(t)



    # 형용사/특징을 2-3번째에 배치: 별칭1개 → 형용사 → 나머지 → 용도

    alias_all = alias + others



    ordered, seen = [], set()



    # 첫 번째 별칭/일반 키워드

    if alias_all:

        t = alias_all[0]

        if t not in seen:

            seen.add(t); ordered.append(t)



    # 형용사/특징 (2-3번째 위치)

    for t in feats:

        if t not in seen:

            seen.add(t); ordered.append(t)



    # 나머지 별칭/일반 키워드

    for t in alias_all[1:]:

        if t not in seen:

            seen.add(t); ordered.append(t)



    # 용도/사용처 (맨 뒤)

    for t in usage:

        if t not in seen:

            seen.add(t); ordered.append(t)



    return ordered







def cap_root_repetition(tokens, root_caps):



    """토큰 내 루트어(부분문자열) 반복 상한 적용. 예: {'건축': 2}"""



    if not tokens or not root_caps:



        return tokens



    kept = []



    counts = defaultdict(int)



    for t in tokens:



        drop = False



        for root, max_n in root_caps.items():



            if root and (root in t):



                if counts[root] >= max_n:



                    drop = True



                    break



        if not drop:



            for root in root_caps:



                if root and (root in t):



                    counts[root] += 1



            kept.append(t)



    return kept





def dedup_compound_suffix(tokens: list, min_suffix_len: int = 2) -> list:

    """합성어 접미어 중복 제거.



    예: ['구름키링', '자동차키링', '가방키링'] → ['구름키링', '자동차', '가방']

    같은 접미어(키링, 케이스, 홀더 등)를 공유하는 합성어가 여러 개 있으면,

    첫 번째는 유지하고 나머지는 접미어를 제거하여 접두어만 남긴다.

    접두어가 이미 토큰 목록에 존재하면 해당 합성어는 제거한다.

    """

    if not tokens or len(tokens) < 2:

        return tokens



    # 한글 합성어만 대상 (3자 이상)

    korean_compounds = []

    for i, t in enumerate(tokens):

        if len(t) >= 3 and re.match(r"^[가-힣]+$", t):

            korean_compounds.append((i, t))



    if len(korean_compounds) < 2:

        return tokens



    # 접미어 후보 추출: 합성어 끝 2~4자 중 2개 이상 공유하는 접미어

    suffix_groups = defaultdict(list)  # {suffix: [(index, token, prefix), ...]}

    for i, t in korean_compounds:

        for slen in range(min_suffix_len, min(5, len(t))):

            suffix = t[-slen:]

            prefix = t[:-slen]

            if len(prefix) >= 1:

                suffix_groups[suffix].append((i, t, prefix))



    # 가장 긴 접미어 우선, 같은 길이면 가장 많은 합성어를 공유하는 접미어

    best_suffix = None

    best_count = 0

    best_slen = 0

    best_items = []

    for suffix, items in suffix_groups.items():

        if len(items) >= 2:

            slen = len(suffix)

            if slen > best_slen or (slen == best_slen and len(items) > best_count):

                best_count = len(items)

                best_slen = slen

                best_suffix = suffix

                best_items = items



    if not best_suffix or best_count < 2:

        return tokens



    # 첫 번째 합성어는 유지, 나머지는 접두어만 남김

    result = list(tokens)

    seen_set = set(tokens)

    first_kept = False

    remove_indices = set()



    for idx, tok, prefix in best_items:

        if not first_kept:

            first_kept = True

            continue  # 첫 번째는 유지

        # 접두어가 이미 존재하면 이 합성어는 제거

        if prefix in seen_set:

            remove_indices.add(idx)

        else:

            # 합성어를 접두어로 교체

            result[idx] = prefix

            seen_set.add(prefix)



    if remove_indices:

        result = [t for i, t in enumerate(result) if i not in remove_indices]



    return result





# ------------------ 검색어설정 보강: 앵커/기준키워드 일관성 ------------------







NEG_DOMAIN = {



    "기름","참기름","현수막","인쇄","프레스","유압","펀칭","천공","제본",



    "상장","상장만들기","복사기","플로터","라미네이팅","코팅기","제판"



}







def build_anchors_from_name(product_name: str) -> set:
    """상품명에서 정체성 중심 앵커를 추출한다."""
    anchors = _collect_identity_tokens_from_name(product_name, max_main=4, max_total=6)
    if not anchors:
        anchors = list(build_baseline_tokens_from_name(product_name))[:4]
    return set(anchors)


def _clean_one_kw(k: str) -> str:
    k = (k or "").strip()
    k = re.sub(r"[^가-힣A-Za-z0-9]", "", k)
    k = re.sub(r"^\d+", "", k)   # leading digits 제거
    k = re.sub(r"\d+$", "", k)   # trailing digits 제거
    return k


_IDENTITY_TOKEN_RE = re.compile(r"[0-9A-Za-z가-힣]+")


_STRONG_IDENTITY_RE = re.compile(
    r"^(?:[A-Za-z]{1,4}\d+[A-Za-z0-9]*|[A-Za-z]{2,5}|\d+(?:mm|cm|m|ml|l|kg|g|w|v|a|호|인치)|304|316|m\d+)$",
    re.IGNORECASE,
)


_WEAK_IDENTITY_WORDS = {
    "고정", "설치", "연결", "장착", "정리", "방지", "보호", "부품", "용품", "도구", "세트", "옵션",
    "사용", "사용처", "실내", "실외", "도어", "문", "작업", "현장", "시공", "수리", "교체", "체결",
    "부착", "결합", "문맥", "대상", "기능", "문제", "해결", "규격", "사이즈", "재질", "소재", "색상",
    "컬러", "형태", "구조", "전용", "일반", "기본", "다용도", "휴대용", "간편", "강력", "다양한",
    "기타", "호환", "필수", "상품", "제품", "액세서리", "악세서리",
}


_WEAK_IDENTITY_SUFFIXES = ("용품", "부품", "도구", "세트", "옵션")


def _normalize_identity_token(token: str) -> str:
    token = normalize_space(str(token or "")).replace(" ", "")
    token = re.sub(r"[^0-9A-Za-z가-힣]", "", token)
    if not token:
        return ""
    replacements = {
        "디링": "D링",
        "가스킷": "가스켓",
        "개스킷": "가스켓",
        "브래킷": "브라켓",
        "스텐": "스테인리스",
    }
    return replacements.get(token, token)


def _identity_semantic_key(token: str) -> str:
    token = _normalize_identity_token(token).lower()
    token = token.replace("차량용", "차량")
    token = re.sub(r"(용|형|식)$", "", token)
    return token


def _split_identity_name_tokens(product_name: str) -> list[str]:
    clean_name = re.sub(GS_CODE_PATTERN, " ", str(product_name or ""))
    out = []
    seen = set()
    for raw in _IDENTITY_TOKEN_RE.findall(clean_name):
        token = _normalize_identity_token(raw)
        if not token or not (2 <= len(token) <= 12):
            continue
        key = _identity_semantic_key(token)
        if key in seen:
            continue
        seen.add(key)
        out.append(token)
    return out


def _is_strong_identity_token(token: str) -> bool:
    token = _normalize_identity_token(token)
    if not token:
        return False
    if _STRONG_IDENTITY_RE.fullmatch(token):
        return True
    if re.search(r"\d", token):
        return True
    return bool(re.search(r"[A-Za-z]", token) and re.search(r"[가-힣]", token))


def _is_weak_identity_token(token: str) -> bool:
    token = _normalize_identity_token(token)
    key = _identity_semantic_key(token)
    if not key or len(key) < 2:
        return True
    if token in STOPWORDS or token in BAN:
        return True
    if key in STOPWORDS or key in BAN or key in _WEAK_IDENTITY_WORDS:
        return True
    if re.fullmatch(r"\d+", key):
        return True
    if any(key.endswith(suffix) for suffix in _WEAK_IDENTITY_SUFFIXES):
        return True
    return False


def _collect_identity_tokens_from_name(product_name: str, max_main: int = 4, max_total: int = 6) -> list[str]:
    raw_tokens = _split_identity_name_tokens(product_name)
    if not raw_tokens:
        return []

    main_scored = []
    specs = []
    fallback = []
    for idx, token in enumerate(raw_tokens):
        if _is_weak_identity_token(token):
            continue
        if _is_strong_identity_token(token):
            specs.append((idx, token))
        else:
            score = 40 - idx
            if idx < 2:
                score += 12
            elif idx < 4:
                score += 6
            if re.search(r"[가-힣]", token):
                score += 4
            main_scored.append((score, idx, token))
        fallback.append(token)

    main_scored.sort(key=lambda item: (-item[0], item[1]))
    ordered = [token for _, _, token in main_scored[:max_main]]
    for _, token in sorted(specs, key=lambda item: item[0]):
        if token not in ordered:
            ordered.append(token)
        if len(ordered) >= max_total:
            break
    if not ordered:
        ordered = fallback[:max_total]
    return ordered[:max_total]


def _semantic_overlap_count(source_tokens: list[str], reference_tokens: list[str]) -> int:
    refs = []
    for ref in reference_tokens:
        key = _identity_semantic_key(ref)
        if not key:
            continue
        refs.append(key)

    matched = set()
    for token in source_tokens:
        key = _identity_semantic_key(token)
        if not key:
            continue
        for ref_key in refs:
            if ref_key in matched:
                continue
            if key == ref_key:
                matched.add(ref_key)
                break
            if len(key) >= 3 and len(ref_key) >= 3 and (key in ref_key or ref_key in key):
                matched.add(ref_key)
                break
    return len(matched)


def semantic_overlap_count(source_tokens: list[str], reference_tokens: list[str]) -> int:
    return _semantic_overlap_count(source_tokens, reference_tokens)


def _has_anchor_overlap(keyword: str, anchors: set) -> bool:
    if not anchors:
        return True
    tokens = _split_identity_name_tokens(keyword)
    if not tokens:
        tokens = [_normalize_identity_token(keyword)]
    return _semantic_overlap_count(tokens, list(anchors)) >= 1


def build_baseline_tokens_from_name(product_name: str) -> set:
    """상품명에서 강한 identity token 중심 baseline을 만든다."""
    baseline = _collect_identity_tokens_from_name(product_name, max_main=4, max_total=6)
    if baseline:
        return set(baseline)

    fallback = []
    for token in _split_identity_name_tokens(product_name):
        if token in STOPWORDS or token in BAN:
            continue
        fallback.append(token)
        if len(fallback) >= 4:
            break
    return set(fallback)


def is_consistent_with_baseline(keyword: str, baseline: set) -> bool:
    """후보 키워드가 baseline identity와 충분히 겹치는지 검사한다."""
    if not baseline:
        return True

    baseline_tokens = [b for b in baseline if not _is_weak_identity_token(b)]
    if not baseline_tokens:
        return True

    keyword_tokens = _split_identity_name_tokens(keyword)
    compact_keyword = _normalize_identity_token(keyword)
    if compact_keyword and compact_keyword not in keyword_tokens:
        keyword_tokens.append(compact_keyword)

    overlap = _semantic_overlap_count(keyword_tokens, baseline_tokens)
    if overlap == 0:
        return False

    strong_overlap = 0
    for base in baseline_tokens:
        base_key = _identity_semantic_key(base)
        if not base_key:
            continue
        matched = False
        for tok in keyword_tokens:
            tok_key = _identity_semantic_key(tok)
            if not tok_key:
                continue
            if tok_key == base_key:
                matched = True
                break
            if len(base_key) >= 3 and len(tok_key) >= 3 and (base_key in tok_key or tok_key in base_key):
                matched = True
                break
        if matched and (_is_strong_identity_token(base) or len(base_key) >= 3):
            strong_overlap += 1

    if strong_overlap >= 1:
        return True
    return overlap >= min(2, len(baseline_tokens))


def is_on_topic(keyword: str, anchors: set, baseline: set) -> bool:
    """금칙 도메인 배제 + 앵커 교집합 + 상품명기반 baseline 일관성"""
    k = _clean_one_kw(keyword)
    if not k:
        return False
    if any(ng in k for ng in NEG_DOMAIN):  # 도메인 배제
        return False
    if anchors and not _has_anchor_overlap(k, anchors):
        return False
    if not is_consistent_with_baseline(k, baseline):
        return False
    return True

# ------------------ 검색어설정(R열) - GPT 5 + 네이버 통합 ------------------



BAN = {"정품","국내발송","무료배송","행사","특가","세일","인기","추천","최고","프리미엄","신제품","베스트",



       "전용","대박","핫딜","당일","빠른","무료","고급","기획","사은","증정","정가","한정","행사상품","총알",



       "가성비","가심비","필수","완전","까지","최대","가능","안정","안정적","안정적인","사용","사용법","탑재",



       "대형","중형","소형","미니","대","중","소","대용량","소용량","초소형","초대형","브랜드","인증"}







SAFE_ALIASES = [
    ("가스켓", "가스킷"),
    ("브라켓", "브래킷"),
    ("D링", "디링"),
]


SYSTEM_FOR_R = """너는 한국 이커머스 검색어 후보 생성기다.
목표: 상품명과 OCR 요약에서 근거가 있는 합성 키워드만 만든다.

출력: 최대 10개, 콤마(,)로만 구분, 공백 없이 붙여쓰기.

슬롯 구성:
1. 핵심상품군 합성어 3~4개
2. 실무 유사어/안전한 별칭 0~1개
3. 기능/문제해결 합성어 1~3개
4. 사용처/대상 합성어 0~2개
5. 재질/규격 합성어 0~1개

필수 규칙:
- evidence-first. 상품명/요약에 근거 없는 사용처/재질/기능 추가 금지.
- intentional typo, 맞춤법 변형, 혼동 유도형 표기변형 생성 금지.
- 길이를 채우기 위한 무관 토큰 추가 금지.
- 같은 뜻 중복 금지. 안전한 별칭은 최대 1개만 허용.
- 키워드만 반환, 설명 금지.

금지어: """ + ",".join(sorted(BAN))


def _compact_korean(s: str) -> str:
    s = re.sub(r"[^가-힣\sA-Za-z0-9]", " ", s or "")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _limit_safe_aliases(parts: list[str], max_aliases: int = 1) -> list[str]:
    alias_group = {}
    for idx, group in enumerate(SAFE_ALIASES):
        for item in group:
            alias_group[_identity_semantic_key(item)] = idx

    out = []
    seen = set()
    used_groups = set()
    alias_count = 0
    for part in parts:
        token = re.sub(r"\s+", "", str(part or ""))
        if not token:
            continue
        key = _identity_semantic_key(token)
        if not key or key in seen:
            continue
        group_id = alias_group.get(key)
        if group_id is not None:
            if group_id in used_groups or alias_count >= max_aliases:
                continue
            used_groups.add(group_id)
            alias_count += 1
        seen.add(key)
        out.append(token)
    return out


def _fallback_heuristic(product_name: str, summary: str, max_n=10) -> str:
    txt = _compact_korean(product_name + " " + summary)
    anchors = build_anchors_from_name(product_name)
    baseline = build_baseline_tokens_from_name(product_name)
    words = [w for w in txt.split() if w not in BAN and 2 <= len(w) <= 12]

    out, seen = [], set()

    def push(w):
        token = re.sub(r"\s+", "", str(w or ""))
        if not token or any(b in token for b in BAN):
            return
        if not (2 <= len(token) <= 12):
            return
        if not is_on_topic(token, anchors, baseline):
            return
        key = _identity_semantic_key(token)
        if not key or key in seen:
            return
        seen.add(key)
        out.append(token)

    for w in words:
        push(w)
        if len(out) >= max_n:
            break

    if len(out) < max_n:
        for i in range(len(words) - 1):
            push(words[i] + words[i + 1])
            if len(out) >= max_n:
                break

    if len(out) < max_n:
        for primary, alias in SAFE_ALIASES:
            if any(primary in item for item in out) and all(alias not in item for item in out):
                push(alias)
                break
            if any(alias in item for item in out) and all(primary not in item for item in out):
                push(primary)
                break

    return ",".join(_limit_safe_aliases(out, max_aliases=1)[:max_n])


def generate_longtail10(product_name: str, summary: str = "", client=None, model_name="gpt-4.1-mini") -> list:
    """상품명 기반 롱테일 10 생성(GPT 있으면 GPT, 없으면 휴리스틱)"""
    product_name = _compact_korean(product_name)
    summary = _compact_korean(summary)
    resolved_client = _resolve_model_client(model_name, explicit_client=client)

    if not resolved_client or not model_name or model_name == "없음":
        return [x for x in _fallback_heuristic(product_name, summary, max_n=10).split(",") if x]

    user_msg = f"""[입력]
상품명: "{product_name}"
요약: "{summary}"

[요청]
- 붙여쓴 핵심 합성어 5~10개를 콤마(,)로만 구분해 출력.
- 실무 유사어/안전한 별칭은 최대 1개만 허용.
- intentional typo, 맞춤법 변형, 길이 채우기용 무관 확장 금지."""

    try:
        resp = resolved_client.chat.completions.create(
            model=model_name,
            messages=[
                {"role": "system", "content": SYSTEM_FOR_R},
                {"role": "user", "content": user_msg},
            ],
            temperature=0.2,
            top_p=0.9,
            max_tokens=180,
        )
        raw = (resp.choices[0].message.content or "").strip().replace(" ", "")
        raw = re.sub(r"[^,가-힣A-Za-z0-9]", "", raw)
        parts = [p for p in raw.split(",") if p]
        if len(parts) < 5:
            parts.extend([x for x in _fallback_heuristic(product_name, summary, max_n=10).split(",") if x])
        parts = _limit_safe_aliases(parts, max_aliases=1)

        seen, out = set(), []
        for p in parts:
            key = _identity_semantic_key(p)
            if not key or key in seen:
                continue
            seen.add(key)
            out.append(p)
            if len(out) >= 10:
                break
        return out
    except Exception:
        return [x for x in _fallback_heuristic(product_name, summary, max_n=10).split(",") if x]


def generate_r_keywords_gpt5(product_name: str, summary: str, client=None, model_name="gpt-4.1") -> str:
    product_name = _compact_korean(product_name)
    summary = _compact_korean(summary)
    resolved_client = _resolve_model_client(model_name, explicit_client=client)

    if not resolved_client or not model_name or model_name == "없음":
        return _fallback_heuristic(product_name, summary, max_n=10)

    user_msg = f"""[입력]
상품명: "{product_name}"
요약: "{summary}"

[요청]
- 붙여쓴 핵심 합성어 5~10개를 콤마(,)로만 구분해 출력.
- 안전한 실무 별칭은 최대 1개만 허용.
- intentional typo, 맞춤법 변형, 무관 확장 금지."""

    try:
        resp = resolved_client.chat.completions.create(
            model=model_name,
            messages=[
                {"role": "system", "content": SYSTEM_FOR_R},
                {"role": "user", "content": user_msg},
            ],
            temperature=0.2,
            top_p=0.9,
            max_tokens=180,
        )
        raw = (resp.choices[0].message.content or "").strip()
    except Exception:
        return _fallback_heuristic(product_name, summary, max_n=10)

    raw = raw.replace(" ", "")
    raw = re.sub(r"[^,가-힣A-Za-z0-9]", "", raw)
    parts = [p for p in raw.split(",") if p]

    clean, seen = [], set()
    for p in _limit_safe_aliases(parts, max_aliases=1):
        if any(b in p for b in BAN):
            continue
        if not (2 <= len(p) <= 12):
            continue
        key = _identity_semantic_key(p)
        if not key or key in seen:
            continue
        seen.add(key)
        clean.append(p)
        if len(clean) >= 10:
            break

    if len(clean) < 5:
        fallback = [x for x in _fallback_heuristic(product_name, summary, max_n=10).split(",") if x]
        for item in fallback:
            key = _identity_semantic_key(item)
            if not key or key in seen:
                continue
            seen.add(key)
            clean.append(item)
            if len(clean) >= 10:
                break

    return ",".join(clean[:10])

# ------------------ (통합) 네이버 검색광고 API: 상위 키워드 + CTR ------------------



NAVER_KEY_FILE = "naver_api_key.txt"



NAVER_BASE_URL = "https://api.searchad.naver.com"



NAVER_TIMEOUT = 15







DRY_RUN = False



SLEEP_BETWEEN_CALLS = 0.7



USE_GPT_BACKFILL = True







def _loose_kv_parse_line(line: str):



    line = line.strip()



    if not line or line.startswith("#"):



        return None, None



    if "=" in line:



        k, v = line.split("=", 1)



    elif ":" in line:



        k, v = line.split(":", 1)



    else:



        return None, None



    k = k.strip().upper()



    v = v.strip().strip('"').strip("'")



    return k, v







def load_naver_keys():

    ensure_env_loaded()

    keys = {

        "ACCESS_LICENSE": get_env("NAVER_ACCESS_LICENSE", "ACCESS_LICENSE"),

        "SECRET_KEY": get_env("NAVER_SECRET_KEY", "SECRET_KEY"),

        "CUSTOMER_ID": get_env("NAVER_CUSTOMER_ID", "CUSTOMER_ID"),

    }



    file_path = next((p for p in key_file_candidates(NAVER_KEY_FILE) if os.path.isfile(p)), None)

    if not file_path:

        return keys



    with open(file_path, 'r', encoding='utf-8') as f:

        for line in f:

            k, v = _loose_kv_parse_line(line)

            if k in keys and not keys[k]:

                keys[k] = v

    return keys



def _sign(secret_key: str, method: str, path: str, timestamp: str) -> str:



    msg = f"{timestamp}.{method}.{path}".encode("utf-8")



    key = secret_key.encode("utf-8")



    digest = hmac.new(key, msg, hashlib.sha256).digest()



    return base64.b64encode(digest).decode("ascii")







def _build_headers(access, secret, customer, method, path):



    ts = str(int(time.time() * 1000))



    sig = _sign(secret, method, path, ts)



    return ts, {



        "X-Timestamp": ts,



        "X-API-KEY": access,



        "X-API-SECRET": secret,



        "X-Customer": str(customer),



        "X-Signature": sig,



        "Content-Type": "application/json; charset=UTF-8",



        "Accept": "application/json",



    }







def _validate_keys(access, secret, customer):



    if not access or not secret or not customer:



        raise RuntimeError("네이버 API 키가 비어있습니다. naver_api_key.txt에 ACCESS_LICENSE/SECRET_KEY/CUSTOMER_ID를 입력하세요.")



    if not str(customer).isdigit():



        raise RuntimeError(f"CUSTOMER_ID 형식 오류: '{customer}' (숫자만)")







def naver_keyword_tool(keys, hint_keywords_str, debug=False):



    """연관키워드 조회. DRY_RUN=True면 모의 데이터 반환. CTR 포함(showDetail=1)"""



    if DRY_RUN:



        base = re.sub(r"\s+", "", hint_keywords_str).split(",")[0][:6] or "키워드"



        mock = []



        for i in range(1, 40):



            mock.append({



                "relKeyword": f"{base}{i}",



                "monthlyPcQcCnt": max(0, 1500 - i * 53),



                "monthlyMobileQcCnt": max(0, 1700 - i * 61),



                "monthlyAvePcCtr": round(max(0.0, 0.12 - i*0.002), 4),      # 12% -> ...



                "monthlyAveMobileCtr": round(max(0.0, 0.14 - i*0.0025), 4)   # 14% -> ...



            })



        return mock



    access = keys.get("ACCESS_LICENSE", "")



    secret = keys.get("SECRET_KEY", "")



    customer = keys.get("CUSTOMER_ID", "")



    _validate_keys(access, secret, customer)



    method = "GET"; path = "/keywordstool"



    timestamp, headers = _build_headers(access, secret, customer, method, path)



    params = {"hintKeywords": hint_keywords_str, "showDetail": 1}



    url = NAVER_BASE_URL + path



    resp = requests.get(url, headers=headers, params=params, timeout=NAVER_TIMEOUT)



    if resp.status_code >= 400:



        raise RuntimeError(f"Naver API 오류 {resp.status_code}: {resp.text[:500]}")



    data = resp.json()



    items = data.get("keywordList") or data.get("result") or data.get("items") or []



    norm = []



    for it in items:



        rel = (it.get("relKeyword") or it.get("keyword") or "").strip()



        pc = it.get("monthlyPcQcCnt") or it.get("monthlyPcSearches") or 0



        mo = it.get("monthlyMobileQcCnt") or it.get("monthlyMobileSearches") or 0



        pc_ctr = it.get("monthlyAvePcCtr") or it.get("monthlyAvePcClickRate") or 0.0



        mo_ctr = it.get("monthlyAveMobileCtr") or it.get("monthlyAveMobileClickRate") or 0.0



        try: pc = int(pc)



        except: pc = 0



        try: mo = int(mo)



        except: mo = 0



        try: pc_ctr = float(pc_ctr)



        except: pc_ctr = 0.0



        try: mo_ctr = float(mo_ctr)



        except: mo_ctr = 0.0



        if rel:



            norm.append({



                "relKeyword": rel,



                "monthlyPcQcCnt": pc,



                "monthlyMobileQcCnt": mo,



                "monthlyAvePcCtr": pc_ctr,



                "monthlyAveMobileCtr": mo_ctr



            })



    return norm







def pick_top(items, topk, key_field):



    arr = sorted(items, key=lambda x: x.get(key_field, 0), reverse=True)



    seen, out = set(), []



    for it in arr:



        kw = (it.get("relKeyword") or "").strip()



        if not kw or kw in seen: continue



        seen.add(kw)



        out.append({"keyword": kw, "searches": int(it.get(key_field, 0))})



        if len(out) >= topk: break



    return out







def rank_and_pick_with_ctr(items, platform="pc", want=5, ctr_threshold=0.05):



    """



    platform: 'pc' or 'mobile'



    1) CTR >= threshold 먼저 상위 검색량 기준으로 선별



    2) 없으면 CTR 무관 상위 검색량에서 보충



    """



    if platform == "pc":



        ctr_key = "monthlyAvePcCtr"; vol_key = "monthlyPcQcCnt"



    else:



        ctr_key = "monthlyAveMobileCtr"; vol_key = "monthlyMobileQcCnt"







    arr = []



    for it in items:



        k = (it.get("relKeyword") or "").strip()



        if not k: continue



        vol = int(it.get(vol_key, 0) or 0)



        ctr = float(it.get(ctr_key, 0.0) or 0.0)



        arr.append((k, vol, ctr))







    # CTR 우선 구간



    high = [(k, v, c) for (k, v, c) in arr if c >= ctr_threshold]



    high.sort(key=lambda x: (x[1], x[2]), reverse=True)







    # 보충 구간



    low = [(k, v, c) for (k, v, c) in arr if (k, v, c) not in high]



    low.sort(key=lambda x: (x[1], x[2]), reverse=True)







    out, seen = [], set()



    for seq in (high, low):



        for k, v, c in seq:



            if k in seen: continue



            seen.add(k); out.append(k)



            if len(out) >= want: break



        if len(out) >= want: break



    return out[:want]







def clean_naver_kw_list(keywords, ban=BAN, anchors: set = None, baseline: set = None):



    """



    네이버/롱테일 키워드 정제:



    - 특문 제거/공백 제거



    - 앞/뒤 숫자 제거



    - 금칙어/길이/중복 필터



    - 온토픽(앵커 + 상품명기반 baseline) 필터



    """



    out = []



    seen = set()



    for k in keywords:



        k = _clean_one_kw(k)



        if not k: continue



        if any(b in k for b in ban): continue



        if not (2 <= len(k) <= 12): continue



        if not is_on_topic(k, anchors or set(), baseline or set()): continue



        if k in seen: continue



        seen.add(k)



        out.append(k)



    return out







# === GPT 키워드(상품명 확장용) ===



LAST_GPT_ERROR = ""





def _apply_keyword_feedback_rules(tokens: list) -> list:

    if not tokens:

        return []



    drop_exact = {

        "\uc18c\uc720\uc790", "\uc560\ud638\uac00", "\uc0ac\uc6a9\uc790", "\uc81c\uc791\uc790", "\uac00\ub4dc\ub108",

        "\uac00\ub2a5\ud55c", "\ud3ec\ud568\ub418\uc9c0", "\ubbf8\ud3ec\ud568", "\ud488\uc9c8", "\ud480", "\ud798",

        "diy", "\uad6c\uc870\ud615\ud0dc",

        "owner", "user", "enthusiast", "quality",

    }

    drop_contains = [

        "\uc218\ub3d9\uce21\uc815", "\ubaa8\ub2c8\ud130\uc0ac\uc815", "\uc624\ucc28", "\uc548\ub0b4",

        "\ubc30\uc1a1", "\uad50\ud658", "\ubc18\ud488", "\uc6d0\uc0b0\uc9c0", "\uc911\uad6d\uc0b0", "\uc218\uc785\uc6d0",

        "\uad6d\ub0b4\ucc3d\uace0", "\uc0ac\uc774\uc988\ucc38\uc870", "\uc0c1\uc138\ucc38\uc870", "\ud0dd1", "\uc635\uc158\uc120\ud0dd",

        "\uc608\uc2dc", "\uc804\uc6d0\ucc28\ub2e8",

        "manual", "monitor", "shipping", "return", "origin",

        "diy", "\uad6c\uc870\ud615\ud0dc",

    

        "손쉽게",

        "도와주는",

        "조절할수",

        "있으며",

        "모양의",

        "적용분야",

        "도톰한",

        "소재로",

        "연출",

        "간편한",

        "신속합니다",

        "높이고",

        "차단하는",

        "가능합니다",

    ]

    replacements = {

        "\uc870\uc808\uac01\ub3c4": "\uac01\ub3c4\uc870\uc808",

        "\ucc28\ub7c9\uc870\uba85\ube0c\ub77c\ucf13": "\ucc28\ub7c9\uc6a9\uc870\uba85\ube0c\ub77c\ucf13",

        "\ucc28\ub7c9\uc870\uba85\ube0c\ub77c\ud0b7": "\ucc28\ub7c9\uc6a9\uc870\uba85\ube0c\ub77c\ucf13",

        "\ud50c\ub9bd\ub3c4\uc5b4\uc9c0\uc9c0\ub300": "\ud50c\ub9bd\ub3c4\uc5b4\uc6a9\uc9c0\uc9c0\ub300",

        "\ucf58\uc13c\ud2b8\uac1c\uc2a4\ud0b7": "\ucf58\uc13c\ud2b8\uc6a9\uac1c\uc2a4\ud0b7",

        "\ucee8\uc13c\ud2b8\uac1c\uc2a4\ud0b7": "\ucf58\uc13c\ud2b8\uc6a9\uac1c\uc2a4\ud0b7",

        "\uad00\uac1c\ucee4\ub125\ud130": "\uad00\uac1c\uc6a9\ucee4\ub125\ud130",

        "\uc2ed\uc790": "\uc2ed\uc790\ud615",

        "\ud2b8\ub7ed\ud480\ub9c1": "\ud2b8\ub7ed\ud480\ub9c1\uc138\ud2b8",

        "\uc545\uc138\uc11c\ub9ac": "\uc561\uc138\uc11c\ub9ac",

        "\uc545\uc138\uc0ac\ub9ac": "\uc561\uc138\uc11c\ub9ac",

    }



    src = [normalize_space(str(x)).replace(" ", "") for x in tokens if x]

    has_vehicle = any(any(k in t for k in ["\ucc28\ub7c9", "\uc790\ub3d9\ucc28", "\ud2b8\ub7ed", "\ud2b8\ub808\uc77c\ub7ec"]) for t in src)

    out = []

    seen = set()



    for t in src:

        if not t:

            continue

        if t == "링":

            t = "디링"

        if len(t) <= 1 and t not in {"디링"}:

            continue

        if t in drop_exact:

            continue

        if any(x in t for x in drop_contains):

            continue

        # 조사/서술형 조각 제거

        if re.search(r"(입니다|합니다|하며|되어|되고|되는|가능)$", t):

            continue

        if re.search(r"[가-힣]+(은|는|을|를|에|에서|으로|로)$", t):

            continue

        t = replacements.get(t, t)

        if t == "\uc561\uc138\uc11c\ub9ac" and has_vehicle:

            t = "\uc790\ub3d9\ucc28\uc561\uc138\uc11c\ub9ac"

        if t not in seen:

            seen.add(t)

            out.append(t)

    # Remove likely OCR-truncated fragments.

    out = [

        t for t in out

        if not (

            len(t) <= 4

            and (t.endswith("\uc870") or t.endswith("\ub3c4") or t.endswith("\ud14c"))

            and t not in {"\uac01\ub3c4\uc870\uc808", "\ud15c\ud50c\ub9bf"}

        )

    ]



    # Compress near-duplicate families to representative tokens.

    families = [

        ["\ucc28\ub7c9\uc6a9\uc870\uba85\ube0c\ub77c\ucf13", "\ucc28\ub7c9\uc870\uba85\ube0c\ub77c\ucf13", "\ucc28\ub7c9\ub77c\uc774\ud2b8\ube0c\ub77c\ucf13"],

        ["\ucc28\ub7c9\uc870\uba85\ub9c8\uc6b4\ud2b8", "\ucc28\ub7c9\ub4f1\ub9c8\uc6b4\ud2b8", "\ub77c\uc774\ud2b8\ub9c8\uc6b4\ud2b8"],

        ["\ud2b8\ub7ed\ud480\ub9c1\uc138\ud2b8", "\ud2b8\ub7ed\ud480\ub9c1", "\ub9c1\ud480\ub9c1", "\ud480\ub9c1\uc138\ud2b8"],

        ["\ud50c\ub9bd\ub3c4\uc5b4\uc6a9\uc9c0\uc9c0\ub300", "\ud50c\ub9bd\uc5c5\ub3c4\uc5b4\uc9c0\uc9c0\ub300", "\ud50c\ub9bd\uc5c5\uc2a4\ud14c\uc774", "\ub3c4\uc5b4\uc2a4\ud14c\uc774"],

        ["\ucf58\uc13c\ud2b8\uc6a9\uac1c\uc2a4\ud0b7", "\ucf58\uc13c\ud2b8\uac1c\uc2a4\ud0b7", "\ucf58\uc13c\ud2b8\uac00\uc2a4\ucf13", "\ucf58\uc13c\ud2b8\uac00\uc2a4\ud0b7"],

        ["\uad00\uac1c\uc6a9\ucee4\ub125\ud130", "\uad00\uac1c\ucee4\ub125\ud130", "\ubaa8\uc138\uad00\ucee4\ub125\ud130"],

    ]

    fam_map = {}

    for fam in families:

        rep = fam[0]

        for k in fam:

            fam_map[k] = rep



    compressed = []

    compressed_seen = set()

    for t in out:

        rep = fam_map.get(t, t)

        if rep not in compressed_seen:

            compressed_seen.add(rep)

            compressed.append(rep)

    out = compressed



    # Hard cap to avoid verbose/repetitive lines.

    return out[:24]



def _finalize_keyword_candidate(tokens: list[str], max_words: int, max_len: int) -> tuple[str, list[str]]:
    seed_text = " ".join([normalize_space(str(token or "")).strip() for token in tokens if str(token or "").strip()])
    _, normalized = postprocess_keywords_tokens(seed_text, max_words=max_words, max_len=max_len)
    normalized = _apply_keyword_feedback_rules(normalized)
    if len(normalized) > max_words:
        normalized = normalized[:max_words]
    out = normalize_space(" ".join(normalized))[:max_len].strip()
    return out, normalized


def _extract_definition_context(definition: dict) -> tuple[str, list[str]]:
    field_specs = [
        ("official_name", "대표명", 1),
        ("category", "카테고리", 1),
        ("materials", "재질", 3),
        ("core_features", "핵심기능", 4),
        ("core_specs", "핵심규격", 3),
        ("use_cases", "용도", 3),
        ("use_context", "사용문맥", 3),
        ("aliases", "안전별칭", 1),
    ]
    lines = []
    flattened = []
    for field, label, limit in field_specs:
        value = definition.get(field)
        items = []
        if isinstance(value, list):
            items = [normalize_space(str(x)) for x in value if normalize_space(str(x))]
        elif isinstance(value, str) and normalize_space(value):
            items = [normalize_space(value)]
        if not items:
            continue
        deduped = []
        seen = set()
        for item in items:
            key = _identity_semantic_key(item)
            if not key or key in seen:
                continue
            seen.add(key)
            deduped.append(item)
        if not deduped:
            continue
        deduped = deduped[:limit]
        lines.append(f"- {label}: {', '.join(deduped)}")
        flattened.extend(deduped)
    return "\n".join(lines), flattened


def keyword_local_score(keyword: str, base_name: str = "", anchors=None, baseline=None) -> int:
    text = normalize_space(str(keyword or "")).strip()
    if not text:
        return -999

    tokens = [t for t in tokenize_korean_words(text) if len(_identity_semantic_key(t)) >= 2]
    if not tokens:
        return -999

    anchor_set = set(anchors or build_anchors_from_name(base_name))
    baseline_set = set(baseline or build_baseline_tokens_from_name(base_name))
    base_tokens = _split_identity_name_tokens(base_name)
    reference_tokens = list(anchor_set | baseline_set | set(base_tokens))

    score = 12 if is_on_topic(text, anchor_set, baseline_set) else -20
    score += min(len(tokens), 8)
    score += _semantic_overlap_count(tokens, list(anchor_set)) * 4
    score += _semantic_overlap_count(tokens, list(baseline_set)) * 5
    score += _semantic_overlap_count(tokens, base_tokens) * 3

    seen = set()
    dup_penalty = 0
    generic_penalty = 0
    off_topic_penalty = 0
    for token in tokens:
        key = _identity_semantic_key(token)
        if not key:
            continue
        if key in seen:
            dup_penalty += 1
        seen.add(key)
        if key in _WEAK_IDENTITY_WORDS:
            generic_penalty += 1
        elif reference_tokens and _semantic_overlap_count([token], reference_tokens) == 0 and not _is_strong_identity_token(token):
            off_topic_penalty += 1

    score -= dup_penalty * 4
    score -= generic_penalty * 2
    score -= off_topic_penalty * 3
    return score


def generate_keyword_gpt(product_name: str,
                         detail_summary: str,
                         model_name: str,
                         max_words: int,
                         max_len: int,
                         min_len: int,
                         vision_analysis: dict = None):
    global LAST_GPT_ERROR
    LAST_GPT_ERROR = ""

    _min_words_needed = int(min_len / 3.5) + 1 if min_len > 0 else max_words
    if max_words < _min_words_needed:
        max_words = min(_min_words_needed, 50)
    target_words = max(6, min(max_words, 24))

    cleaned_name = re.sub(r"GS\d{7}[A-Z0-9]*\s*", "", str(product_name)).strip()
    ocr_text = normalize_space(str(detail_summary or ""))
    base_name_for_score = cleaned_name or normalize_space(str(product_name or ""))
    base_anchors = build_anchors_from_name(base_name_for_score)
    base_baseline = build_baseline_tokens_from_name(base_name_for_score)
    resolved_client = _resolve_model_client(model_name)

    def finalize_text(text: str) -> tuple[str, list[str]]:
        _, tokens = postprocess_keywords_tokens(text, max_words=target_words, max_len=max_len)
        return _finalize_keyword_candidate(tokens, max_words=target_words, max_len=max_len)

    def rank(candidate: tuple[str, list[str]]) -> tuple[int, int, int, int]:
        out, tokens = candidate
        if not out:
            return (-999, -999, -999, -999)
        score = keyword_local_score(out, base_name=base_name_for_score, anchors=base_anchors, baseline=base_baseline)
        overlap = _semantic_overlap_count(tokens, list(base_baseline or base_anchors))
        unique = len({_identity_semantic_key(tok) for tok in tokens if _identity_semantic_key(tok)})
        generic = sum(1 for tok in tokens if _identity_semantic_key(tok) in _WEAK_IDENTITY_WORDS)
        return (score, overlap, unique - generic, -generic)

    def choose_better(current: tuple[str, list[str]], candidate: tuple[str, list[str]]) -> tuple[str, list[str]]:
        return candidate if rank(candidate) > rank(current) else current

    if not resolved_client or model_name == "없음":
        return finalize_text(f"{cleaned_name} {detail_summary}".strip())

    try:
        vision_context = ""
        if vision_analysis:
            lines = []
            for section, fields in vision_analysis.items():
                if isinstance(fields, dict):
                    for key, value in fields.items():
                        if isinstance(value, list):
                            vals = [str(x) for x in value if str(x).strip()]
                            if vals:
                                lines.append(f"- {section}.{key}: {', '.join(vals)}")
                        elif isinstance(value, str) and value.strip():
                            lines.append(f"- {section}.{key}: {value}")
            if lines:
                vision_context = "\n[이미지 분석 결과]\n" + "\n".join(lines)

        define_system = (
            "너는 한국 이커머스 상품 분석 전문가다. "
            "상품명과 OCR 텍스트에서 상품 정체성을 높은 정밀도로 추론하라. "
            + ("이미지 분석 결과가 보조 근거로 제공된다. 이미지 근거가 OCR/상품명과 일치할 때만 반영하라. " if vision_context else "")
            + "노이즈 보일러플레이트는 무시: 배송/반품/프로모션/정책/원산지/수입원/연락처/공지/쿠폰/이벤트/수동측정오차/모니터색상차이. "
            "반드시 JSON만 반환. 근거 없는 스펙/인증/호환성/새 카테고리 생성 금지. 오타/맞춤법 변형을 alias로 만들지 말라."
        )
        define_vision_suffix = "\n" + vision_context if vision_context else ""
        define_user = f"""아래 JSON 스키마를 정확히 채워 반환하라:
{{
  "official_name": string,
  "category": string,
  "brand": string,
  "materials": string array,
  "core_features": string array,
  "core_specs": string array,
  "use_cases": string array,
  "use_context": string array,
  "target_users": string array,
  "spec_terms": string array,
  "aliases": string array,
  "field_terms": string array,
  "compatibility": string array,
  "excluded_noise": string array
}}
Rules:
- official_name: 정확한 제품 대표명.
- category: 상세 카테고리 경로.
- materials/core_features/core_specs/use_cases/use_context: 근거 있는 값만.
- aliases: 안전한 실무 유사어만, 최대 2개. intentional typo/맞춤법 변형 금지.
ProductName: {cleaned_name}
OCRSummary: {ocr_text[:2500]}{define_vision_suffix}"""

        define_resp = resolved_client.chat.completions.create(
            model=model_name,
            messages=[
                {"role": "system", "content": define_system},
                {"role": "user", "content": define_user},
            ],
            temperature=0.1,
            top_p=0.9,
            max_tokens=900,
            response_format={"type": "json_object"},
        )
        definition_raw = (define_resp.choices[0].message.content or "").strip()
        try:
            definition = json.loads(definition_raw) if definition_raw else {}
        except Exception:
            definition = {}

        definition_context, definition_terms = _extract_definition_context(definition)

        keyword_system = f"""너는 한국 이커머스 상품 키워드 조립기다.
목표는 더 길게 쓰는 것이 아니라, 근거 있는 정보축만 남긴 한 줄 키워드를 만드는 것이다.

핵심 원칙:
- evidence-first: 상품명 > 구조화 정의 > OCR/Vision 순서로 사용한다.
- 근거 없는 사용처/재질/기능/대상 추가 금지.
- 검색 커버리지 목적의 오타/맞춤법변형/무관 동의어 생성 금지.
- 길이가 부족해도 무관 확장 금지. 근거가 부족하면 짧아도 된다.
- 같은 뜻 중복 금지.

출력 규칙:
- 한 줄, 공백 구분, 키워드만 출력.
- 목표 길이는 {min_len}~{max_len}자이지만 강제하지 않는다.
- 최대 {target_words}토큰까지만 사용한다.
- 순서는 정체성 -> 구조/규격 -> 핵심 기능 -> 사용처(근거 있을 때만) -> 재질/옵션."""

        definition_block = "\n[구조화된 정의]\n" + definition_context if definition_context else ""
        keyword_user = f"""상품명: {cleaned_name}

OCR요약: {ocr_text[:1800]}
{definition_block}
{vision_context if vision_context else ''}

작성 지시:
- official_name, category, materials, core_features, core_specs, use_cases, use_context, aliases 중 비어 있지 않은 값만 반영한다.
- aliases는 안전한 실무 유사어 0~1개만 허용한다.
- use_cases/use_context는 상품명/OCR와 교집합이 있을 때만 사용한다.
- 새 카테고리, 새 상품군, 새 사용처를 창작하지 말라.
- 짧더라도 정체성을 우선한다.

한 줄 키워드만 출력"""

        resp = resolved_client.chat.completions.create(
            model=model_name,
            messages=[
                {"role": "system", "content": keyword_system},
                {"role": "user", "content": keyword_user},
            ],
            temperature=0.1,
            top_p=0.8,
            max_tokens=520,
        )
        best = finalize_text((resp.choices[0].message.content or "").strip())
        if definition_terms:
            best = choose_better(best, finalize_text(f"{cleaned_name} {' '.join(definition_terms)}".strip()))

        for attempt in range(2):
            if rank(best)[0] >= 16 and len(best[1]) >= 5:
                break
            retry_msg = f"""현재 결과: {best[0] or '(없음)'}

다시 작성하라.
- 더 길게가 아니라 더 on-topic이고 더 정보축이 좋게 만든다.
- generic token, 중복 token, 무관 token을 제거한다.
- OCR에서 추가할 수 있는 토큰은 base identity와 겹치는 것만 허용한다.
- 근거가 부족하면 짧게 끝낸다.
- intentional typo/맞춤법변형/무관 사용처 추가 금지.

한 줄 키워드만 출력"""
            retry = resolved_client.chat.completions.create(
                model=model_name,
                messages=[
                    {"role": "system", "content": keyword_system},
                    {"role": "user", "content": retry_msg},
                    {"role": "user", "content": keyword_user},
                ],
                temperature=0.12 + attempt * 0.04,
                max_tokens=520,
            )
            best = choose_better(best, finalize_text((retry.choices[0].message.content or "").strip()))

        if ocr_text:
            _, ocr_tokens = postprocess_keywords_tokens(ocr_text, max_words=max(target_words + 6, 20), max_len=max(max_len * 2, 220))
            augmented = list(best[1])
            existing = {_identity_semantic_key(tok) for tok in augmented}
            for token in ocr_tokens:
                key = _identity_semantic_key(token)
                if not key or key in existing:
                    continue
                if not is_on_topic(token, base_anchors, base_baseline):
                    continue
                if _semantic_overlap_count([token], list(base_baseline or base_anchors)) == 0 and not _is_strong_identity_token(token):
                    continue
                augmented.append(token)
                existing.add(key)
                if len(augmented) >= target_words:
                    break
            best = choose_better(best, _finalize_keyword_candidate(augmented, max_words=target_words, max_len=max_len))

        if not best[0]:
            best = finalize_text(f"{cleaned_name} {detail_summary}".strip())
        return best
    except Exception as e:
        LAST_GPT_ERROR = f"{type(e).__name__}: {e}"
        return finalize_text(f"{product_name} {detail_summary}".strip())

def _dedupe_semantic_keyword_tokens(tokens: list[str]) -> list[str]:
    out = []
    seen = set()
    for token in tokens:
        key = _identity_semantic_key(token)
        if not key or key in seen:
            continue
        seen.add(key)
        out.append(token)
    return out


def _collect_stage2_reference_tokens(seed_tokens: list[str], naver_keyword_table: str, ocr_text: str) -> list[str]:
    anchors = build_anchors_from_name(" ".join(seed_tokens))
    baseline = build_baseline_tokens_from_name(" ".join(seed_tokens))
    refs = list(seed_tokens)

    for keyword, _ in _extract_naver_candidates_from_table(naver_keyword_table)[:12]:
        if not is_on_topic(keyword, anchors, baseline):
            continue
        for token in tokenize_korean_words(keyword):
            if _semantic_overlap_count([token], seed_tokens) >= 1 or _is_strong_identity_token(token):
                refs.append(token)

    for token in tokenize_korean_words(ocr_text):
        if not is_on_topic(token, anchors, baseline):
            continue
        if _semantic_overlap_count([token], seed_tokens) >= 1:
            refs.append(token)

    return _dedupe_semantic_keyword_tokens(refs)


def _filter_stage2_tokens(tokens: list[str], reference_tokens: list[str], anchors: set, baseline: set) -> list[str]:
    out = []
    seen = set()
    for token in tokens:
        key = _identity_semantic_key(token)
        if not key or key in seen:
            continue
        if _semantic_overlap_count([token], reference_tokens) >= 1:
            out.append(token)
            seen.add(key)
            continue
        if is_on_topic(token, anchors, baseline) and (_is_strong_identity_token(token) or _semantic_overlap_count([token], list(anchors)) >= 1):
            out.append(token)
            seen.add(key)
    return out


def generate_keyword_stage2(
    seed_keywords: str,
    naver_keyword_table: str,
    ocr_text: str = "",
    model_name: str = "gpt-4.1",
    min_len: int = 50,
    max_len: int = 100,
    max_words: int = 24,
):
    """2차 키워드는 생성이 아니라 재정렬/정제/우선순위 조정에만 사용한다."""
    global LAST_GPT_ERROR
    LAST_GPT_ERROR = ""

    seed_keywords = normalize_space(str(seed_keywords or "")).strip()
    naver_keyword_table = normalize_space(str(naver_keyword_table or "")).strip()
    ocr_text = normalize_space(str(ocr_text or "")).strip()
    target_words = max(6, min(max_words, 24))
    resolved_client = _resolve_model_client(model_name)

    if not seed_keywords:
        return "", []

    _, seed_tokens_raw = postprocess_keywords_tokens(seed_keywords, max_words=target_words, max_len=max_len)
    seed_out, seed_tokens = _finalize_keyword_candidate(seed_tokens_raw, max_words=target_words, max_len=max_len)
    if not seed_out:
        return "", []

    anchors = build_anchors_from_name(seed_keywords)
    baseline = build_baseline_tokens_from_name(seed_keywords)
    reference_tokens = _collect_stage2_reference_tokens(seed_tokens, naver_keyword_table, ocr_text)
    seed_score = keyword_local_score(seed_out, base_name=seed_keywords, anchors=anchors, baseline=baseline)

    if not resolved_client or model_name == "없음" or not naver_keyword_table:
        return seed_out, seed_tokens

    system_msg = """너는 한국 이커머스 상품명 정제기다.
역할은 새로 확장하는 것이 아니라, seed_keywords를 재정렬하고 불필요한 토큰을 제거하는 것이다.

절대 규칙:
- seed_keywords 안의 정체성을 유지한다.
- 새로운 상품군, 새로운 사용처, 새로운 문맥을 창작하지 않는다.
- intentional typo, 맞춤법 변형, 공격적 동의어 확장 금지.
- 카테고리 일반어(용품/도구/부품), 마케팅 수식어(간편/다용도/강력/휴대용/접이식) 추가 금지.
- 더 짧아져도 괜찮다. 더 깨끗하고 더 on-topic인 결과를 우선한다.
- 한 줄 키워드만 출력한다."""

    user_msg = f"""seed_keywords:
{seed_out}

참고 reference_tokens:
{', '.join(reference_tokens[:20])}

네이버 검색광고 데이터:
{naver_keyword_table[:1200]}

작업:
- seed_keywords를 재정렬/축약/정제한다.
- 네이버 데이터는 같은 카테고리의 우선순위 참고용으로만 사용한다.
- 새 카테고리, 새 상품군, 새 사용처를 만들지 않는다.
- 결과가 seed보다 더 약하거나 더 멀어지면 seed를 유지한다.

한 줄 키워드만 출력"""

    try:
        resp = resolved_client.chat.completions.create(
            model=model_name,
            messages=[
                {"role": "system", "content": system_msg},
                {"role": "user", "content": user_msg},
            ],
            temperature=0.1,
            max_tokens=320,
        )
        _, candidate_tokens_raw = postprocess_keywords_tokens((resp.choices[0].message.content or "").strip(), max_words=target_words, max_len=max_len)
        candidate_tokens = _filter_stage2_tokens(candidate_tokens_raw, reference_tokens, anchors, baseline)
        candidate_out, candidate_tokens = _finalize_keyword_candidate(candidate_tokens, max_words=target_words, max_len=max_len)

        if not candidate_out:
            return seed_out, seed_tokens
        if _semantic_overlap_count(candidate_tokens, seed_tokens) == 0:
            return seed_out, seed_tokens

        candidate_score = keyword_local_score(candidate_out, base_name=seed_keywords, anchors=anchors, baseline=baseline)
        if candidate_score < seed_score:
            return seed_out, seed_tokens
        return candidate_out, candidate_tokens
    except Exception as e:
        LAST_GPT_ERROR = f"{type(e).__name__}: {e}"
        return seed_out, seed_tokens

def _extract_naver_candidates_from_table(naver_keyword_table: str):

    rows = []

    text = str(naver_keyword_table or "").strip()

    if not text:

        return rows

    for line in text.splitlines():

        line = line.strip()

        if not line or line.startswith("키워드|"):

            continue

        parts = [p.strip() for p in line.split("|")]

        if len(parts) >= 4:

            kw = parts[0]

            try:

                total = int(parts[3])

            except Exception:

                try:

                    total = int(parts[1]) + int(parts[2])

                except Exception:

                    total = 0

            if kw:

                rows.append((kw, total))

    if rows:

        rows.sort(key=lambda x: x[1], reverse=True)

        return rows



    m = re.search(r"PC5=([^|]+)", text)

    if m:

        for kw in m.group(1).split(","):

            kw = _clean_one_kw(kw)

            if kw:

                rows.append((kw, 0))

    m = re.search(r"MO5=([^|]+)", text)

    if m:

        for kw in m.group(1).split(","):

            kw = _clean_one_kw(kw)

            if kw and kw not in [r[0] for r in rows]:

                rows.append((kw, 0))

    return rows





def generate_search_terms20(

    final_keyword: str,

    naver_keyword_table: str,

    model_name: str = "gpt-4.1-mini",

    anchors=None,

    baseline=None,

):

    """

    검색어설정 20개 생성:

    - 네이버 검색량 50% 규칙 우선

    - 부족분은 최종 키워드 핵심어 조합으로 보완

    """

    global LAST_GPT_ERROR

    LAST_GPT_ERROR = ""



    final_keyword = normalize_space(str(final_keyword or "")).strip()

    if not final_keyword:

        return []



    resolved_client = _resolve_model_client(model_name)

    anchor_set = set(anchors or [])

    baseline_set = set(baseline or [])

    candidates = _extract_naver_candidates_from_table(naver_keyword_table)



    def _normalize_kw_list(items):

        out, seen = [], set()

        for kw in items:

            kw = _clean_one_kw(str(kw or "")).replace(" ", "")

            if not kw:

                continue

            if kw in seen:

                continue

            if len(kw) < 2 or len(kw) > 20:

                continue

            if any(b in kw for b in BAN):

                continue

            if anchor_set and baseline_set and not is_on_topic(kw, anchor_set, baseline_set):

                continue

            seen.add(kw)

            out.append(kw)

            if len(out) >= 20:

                break

        return out



    rule_base = []

    if candidates:

        max_vol = max([v for _, v in candidates] or [0])

        thr = int(max_vol * 0.5) if max_vol > 0 else 0

        for kw, vol in candidates:

            if vol >= thr:

                rule_base.append(kw)

        rule_base = rule_base[:20]



    gpt_list = []

    if resolved_client is not None and model_name != "없음":

        system_msg = """너는 한국 이커머스(네이버/쿠팡) 검색어설정 생성기다.

목표: 쿠팡 searchTags 규격(각 20자 이내, 최대 20개)에 맞는 검색어 생성.

정합성(상품-카테고리 일치) 최우선. 노출 극대화보다 정책 안전성 우선.



슬롯 구성(이 순서로 채워라):

1. 핵심상품군 5~7개: 제품유형+핵심기능 조합

2. 기능/문제해결 4~5개: 구매 의도 직결 키워드

3. 사용처/대상 3~4개: 장소/상황/호환 대상

4. 재질/규격 2~3개: 근거 있는 스펙/재질

5. 동의어/별칭 1~2개: 표기변형 최소화



강제 규칙:

- 정확히 20개 생성. 콤마로 구분. 한 줄 출력.

- 각 키워드 2~20자.

- 상품명/카테고리에 이미 포함된 단어는 중복 입력 최소화.

- 각 키워드는 서로 다른 정보축(기능/용도/재질/대상 등) 담아 검색의도 다양화.

- 같은 뜻 중복 금지.

- 무관 브랜드/인기어/경쟁사 상표 금지.

- 근거 없는 최상급/과장 효능 금지.

- 감성/홍보 형용사 금지: 귀여운, 예쁜, 고급진, 럭셔리, 힐링, 인싸, 필수품, 데일리, 인기, 추천, 베스트, 프리미엄.

- 판매조건 금지: 무료배송, 당일발송, 특가, 세일, 할인.

- 무관 확장 금지.

- 색상/사이즈는 1~2개 이하.

- 네이버 데이터는 우선순위 참고용. 상품군 변경 금지.

- 설명 금지. 키워드만 출력."""

        user_msg = f"""최종 키워드:

{final_keyword}



네이버 검색광고 데이터:

{naver_keyword_table}



출력 전 자가점검:

1) 핵심상품군이 앞쪽에 충분히 포함됐는가?

2) 기능/문제해결이 구매 의도를 반영하는가?

3) 사용처/대상이 실제 사용 맥락인가?

4) 재질/규격에 근거가 있는가?

5) 무관 카테고리 확장이 없는가?

6) 동의어가 과다하지 않은가?

7) 각 키워드가 20자 이내인가?



정확히 20개를 콤마로만 구분해 출력"""

        try:

            resp = resolved_client.chat.completions.create(

                model=model_name,

                messages=[

                    {"role": "system", "content": system_msg},

                    {"role": "user", "content": user_msg},

                ],

                temperature=0.1,

                max_tokens=420,

            )

            raw = (resp.choices[0].message.content or "").strip()

            gpt_list = [x.strip() for x in raw.replace("\n", ",").split(",") if x.strip()]

        except Exception as e:

            LAST_GPT_ERROR = f"{type(e).__name__}: {e}"



    merged = _normalize_kw_list(rule_base + gpt_list)



    if len(merged) < 20:

        base_tokens = [t for t in re.sub(r"[^0-9가-힣A-Za-z\s]", " ", final_keyword).split() if len(t) >= 2]

        extras = []

        extras.extend(base_tokens)

        for i in range(len(base_tokens) - 1):

            extras.append((base_tokens[i] + base_tokens[i + 1]).replace(" ", ""))

        more = _normalize_kw_list(extras)

        seen = set(merged)

        for kw in more:

            if kw not in seen:

                merged.append(kw)

                seen.add(kw)

            if len(merged) >= 20:

                break



    return merged[:20]







def generate_search_keywords(product_name: str, ocr_text: str, model_name: str = "gpt-4.1-mini") -> str:

    """

    OCR 텍스트 기반으로 검색키워드(공백 구분 1줄)를 생성.

    """

    global LAST_GPT_ERROR

    LAST_GPT_ERROR = ""



    resolved_client = _resolve_model_client(model_name)



    if not resolved_client or model_name == "없음":

        return ""



    if not ocr_text or not ocr_text.strip():

        return ""



    try:

        cleaned_name = re.sub(r"GS\d{7}[A-Z0-9]*\s*", "", str(product_name)).strip()



        system_msg = """너는 한국 이커머스(네이버/쿠팡) 검색키워드 생성 전문가다.

목표: 정합성(상품-카테고리-속성 일치) 극대화 + 정책 위반 위험 최소화.

입력된 근거 데이터만 사용하라.



## 슬롯 구조 (이 순서로 채워라):

1. 핵심상품군 3~5개: 제품유형/카테고리 대표어

2. 핵심스펙 1~3개: 수치 기반 스펙(용량/길이/규격/인증). 근거 있는 값만.

3. 기능/문제해결 3~5개: 실제 구매 의도 직결 키워드

4. 사용처/대상 2~4개: 장소/상황/호환 대상

5. 재질 0~2개: 근거 있을 때만

6. 동의어/별칭 0~2개: 표기변형 최소화



## 절대 규칙:

- 상품과 무관한 유명 브랜드/인기어/경쟁사 상표 금지

- 근거 없는 최상급(최고/1위/완벽), 과장 효능, 오해 소지 표현 금지

- 감성/홍보 형용사 금지: 귀여운, 예쁜, 고급진, 럭셔리, 힐링, 인싸, 필수품, 데일리, 인기, 추천, 베스트, 프리미엄

- 판매조건 금지: 무료배송, 당일발송, 특가, 세일, 할인, 쿠폰, 이벤트

- 같은 뜻 단어 2개 초과 금지

- 문장형/조사/홍보/배송/할인 문구 금지

- 색상/사이즈 1~2개 이하

- 숫자 찌꺼기 금지 (실제 규격 숫자는 허용)



## 출력 형식:

- 한 줄, 공백 구분, 12~18개 토큰

- 명사구 중심, 불필요 특수문자 금지



## 제외 키워드:

원산지, 수입원, 제조국, 상세참조, 주의사항, 배송, 교환, 반품, 가격, 할인, 이벤트"""



        msgs = [

            {"role": "system", "content": system_msg},

            {

                "role": "user",

                "content": (

                    f"상품명: {cleaned_name}\n\n"

                    f"OCR 텍스트:\n{ocr_text[:1500]}\n\n"

                    "작업 순서:\n"

                    "1) 핵심상품군 3~5개 확정\n"

                    "2) 사용처/대상 2~4개 선택\n"

                    "3) 기능/문제해결 3~5개 선택\n"

                    "4) 재질/규격은 0~2개만 추가\n"

                    "5) 연관검색어/동의어는 0~2개만 추가\n"

                    "6) 감성어/홍보어/무관 확장어 제거\n\n"

                    "출력은 공백 구분 한 줄 키워드만 반환하라."

                ),

            },

        ]



        resp = resolved_client.chat.completions.create(

            model=model_name,

            messages=msgs,

            temperature=0.15,

            max_tokens=800,

        )



        result = (resp.choices[0].message.content or "").strip()

        result = re.sub(r"\s+", " ", result).strip()

        return result



    except Exception as e:

        LAST_GPT_ERROR = f"{type(e).__name__}: {e}"

        return ""





# === 湲곕낯?곹뭹紐??듭뀡 遺꾨━ & 寃고빀 ===



GS_CODE_PATTERN = re.compile(r"\bGS\d{7}[A-Z0-9]*\b", flags=re.IGNORECASE)







def extract_base_and_option(full_name: str):



    s = normalize_space(str(full_name).strip())



    m = GS_CODE_PATTERN.search(s)



    if m:



        base = s[:m.start()].strip()



        opt = s[m.end():].strip()



        return (base if base else s), opt



    return s, ""







SIZE_COMPOUND_PREFIX = re.compile(r"^((?:[ABab]\d{1,2})|(?:\d{1,2}[A-Za-z]?)|(?:대형|중형|소형))")







def _has_korean(s: str) -> bool: return bool(re.search(r"[가-힣]", s))



def _is_size_compound(tok: str) -> bool: return bool(SIZE_COMPOUND_PREFIX.match(tok)) or bool(re.search(r"\d", tok))







def _should_skip_due_to_existing_compound(candidate: str, existing_tokens: list) -> bool:



    for ex in existing_tokens:



        if ex == candidate: continue



        if len(ex) > len(candidate) and _has_korean(ex) and ex.endswith(candidate) and _is_size_compound(ex):



            return True



    return False







def tokenize_korean_words(s: str):



    s = re.sub(r"[^0-9가-힣\s]", " ", str(s))



    return normalize_space(s).split()







def _semantic_token_key(token: str) -> str:



    return re.sub(r"[^0-9가-힣]", "", str(token or "")).lower()







def _token_semantic_overlap(token: str, other_tokens: list[str]) -> bool:



    key = _semantic_token_key(token)

    if len(key) < 2:

        return False

    for other in other_tokens:

        other_key = _semantic_token_key(other)

        if len(other_key) < 2:

            continue

        if key == other_key or key in other_key or other_key in key:

            return True

    return False







def _looks_like_consumer_title(base_name: str, kw_tokens: list[str], ocr_text: str = "") -> bool:



    base_name_clean = normalize_space(re.sub(GS_CODE_PATTERN, "", str(base_name or ""))).strip()

    if not base_name_clean:

        return False



    base_tokens = [t for t in base_name_clean.split() if _semantic_token_key(t)]

    hangul_tokens = [t for t in base_tokens if re.search(r"[가-힣]", t)]

    if len(hangul_tokens) < 2:

        return False

    if len(base_tokens) < 3 and len(base_name_clean) < 10:

        return False

    if re.fullmatch(r"[A-Za-z0-9\s\-_/]+", base_name_clean):

        return False



    ref_tokens = list(kw_tokens or [])

    if not ref_tokens and ocr_text:

        ref_tokens = tokenize_korean_words(ocr_text)

    if not ref_tokens:

        return len(base_tokens) >= 3



    overlap = sum(1 for tok in hangul_tokens if _token_semantic_overlap(tok, ref_tokens))

    return overlap >= 1 and (len(base_tokens) >= 4 or len(base_name_clean) >= 12)







def merge_base_name_with_keywords(base_name: str, kw_line: str, max_words: int, max_len: int, option_tokens: set, ocr_text: str = ""):



    base_name_clean = normalize_space(re.sub(GS_CODE_PATTERN, "", str(base_name or ""))).strip()



    base_tokens = [t for t in base_name_clean.split() if t]



    kw_tokens = [t for t in normalize_space(kw_line).split() if t]



    _, ocr_tokens = postprocess_keywords_tokens(

        ocr_text or "",

        max_words=max(max_words + 6, 20),

        max_len=max(max_len * 2, 200),

    )



    keep_original = _looks_like_consumer_title(base_name_clean, kw_tokens, ocr_text)



    if keep_original:

        seed_tokens = [t for t in base_tokens if t not in option_tokens]

        extra_sources = [kw_tokens, ocr_tokens]

        extra_budget = 3

    else:

        seed_tokens = [t for t in kw_tokens if t not in option_tokens]

        if not seed_tokens:

            seed_tokens = [t for t in ocr_tokens if t not in option_tokens]

        if not seed_tokens:

            seed_tokens = [t for t in base_tokens if t not in option_tokens]

        extra_sources = [ocr_tokens, base_tokens]

        extra_budget = max_words



    seen = set(seed_tokens)



    out_tokens = list(seed_tokens)

    extra_added = 0



    option_order = []

    for tok in base_tokens + kw_tokens + ocr_tokens:

        if tok in option_tokens and tok not in option_order:

            option_order.append(tok)



    for source_tokens in extra_sources:

        for t in source_tokens:



            if not t: continue



            if t in option_tokens: continue



            if _should_skip_due_to_existing_compound(t, out_tokens): continue



            _is_substr_dup = False

            if len(t) >= 2 and re.match(r"^[가-힣]+$", t):

                for ex in out_tokens:

                    if ex == t: continue

                    if len(ex) >= 2 and re.match(r"^[가-힣]+$", ex):

                        if len(t) < len(ex) and t in ex:

                            _is_substr_dup = True

                            break

            if _is_substr_dup: continue



            if t not in seen:



                seen.add(t); out_tokens.append(t)

                extra_added += 1



            if len(out_tokens) >= max_words or extra_added >= extra_budget: break

        if len(out_tokens) >= max_words or extra_added >= extra_budget:

            break



    for t in option_order:

        if len(out_tokens) >= max_words:

            break

        if not t or t in seen:

            continue

        seen.add(t)

        out_tokens.append(t)



    out = " ".join(out_tokens[:max_words]).strip()



    return out[:max_len].rstrip() if len(out) > max_len else out







def insert_img_tag(html, img_tag):
    if not img_tag:
        return html

    html = "" if (html is None or (isinstance(html, float) and np.isnan(html))) else str(html)

    if img_tag in html: return html

    if html.startswith("<center>") and len(html) > 8:
        return html[:8] + img_tag + html[8:]

    # <div style="text-align: center;" ...> 형태 처리
    import re
    m = re.match(r'(<div\s+style="text-align:\s*center;?"[^>]*>)', html, re.IGNORECASE)
    if m:
        insert_pos = m.end()
        return html[:insert_pos] + img_tag + html[insert_pos:]

    return img_tag + html







def get_multiplier(value: float) -> float:



    try: v = float(value)



    except Exception: return 2.0



    if v >= 20000: return 1.6



    elif v >= 10000: return 1.8



    else: return 2.0







# === 파일 스캔 & CSV ===



def iter_files_with_depth(root_dir: str, max_depth: int):



    root_dir = os.path.abspath(root_dir); rsep = os.sep; rdepth = root_dir.count(rsep)



    for base, _, files in os.walk(root_dir):



        if max_depth >= 0 and (base.count(rsep) - rdepth > max_depth): continue



        yield base, files







def find_local_images_for_code(root_dir: str, code9: str, allow_folder_match: bool = True, max_depth: int = -1):



    if not root_dir or not os.path.isdir(root_dir) or not code9: return []



    exts = {".jpg", ".jpeg", ".png", ".webp", ".bmp"}



    hits = []; code9_lower = code9.lower()



    for base, files in iter_files_with_depth(root_dir, max_depth):



        base_low = base.lower(); folder_has_code = allow_folder_match and (code9_lower in base_low)



        for fn in files:



            _, ext = os.path.splitext(fn)



            if ext.lower() not in exts: continue



            fn_low = fn.lower()



            if folder_has_code or (code9_lower in fn_low):



                hits.append(os.path.join(base, fn))



    return hits







def safe_read_csv(path: str) -> pd.DataFrame:

    """CSV/Excel 파일을 안전하게 읽기 (확장자 자동 감지)"""



    # 파일 확장자 확인

    _, ext = os.path.splitext(path)

    ext_lower = ext.lower()



    # Excel 파일 처리 (.xlsx, .xls)

    if ext_lower in ['.xlsx', '.xls']:

        try:

            df = pd.read_excel(path, engine='openpyxl' if ext_lower == '.xlsx' else None)

            try:

                print(f"[파일] [OK] Excel 파일 읽기 성공: {os.path.basename(path)}")

            except UnicodeEncodeError:

                print(f"[파일] [OK] Excel 파일 읽기 성공")

            return df

        except Exception as e:

            try:

                print(f"[파일] [ERROR] Excel 파일 읽기 실패: {e}")

            except UnicodeEncodeError:

                print(f"[파일] [ERROR] Excel 파일 읽기 실패")

            raise ValueError(f"Excel 파일을 읽을 수 없습니다: {path}\n오류: {str(e)}")



    # CSV 파일 처리

    with open(path, "rb") as f: raw = f.read(100000)  # 더 많은 바이트 읽기



    detected = chardet.detect(raw)

    enc = (detected.get('encoding') or '').lower()

    confidence = detected.get('confidence', 0)



    try:

        print(f"[CSV] chardet 감지: encoding={enc}, confidence={confidence:.2f}")

    except UnicodeEncodeError:

        print(f"[CSV] chardet 감지 완료")



    tried = []



    # 더 많은 인코딩 후보 추가 (순서 중요)

    # confidence가 낮으면 chardet 결과를 신뢰하지 않고 일반적인 인코딩부터 시도

    if confidence < 0.7:

        encodings = ["utf-8-sig", "utf-8", "euc-kr", "cp949", enc, "utf-16", "latin-1", "iso-8859-1"]

    else:

        encodings = [enc, "utf-8-sig", "utf-8", "euc-kr", "cp949", "utf-16", "latin-1", "iso-8859-1"]



    for cand in encodings:



        if not cand: continue



        try:



            df = pd.read_csv(path, encoding=cand)

            try:

                print(f"[CSV] [OK] 성공: encoding={cand}")

            except UnicodeEncodeError:

                print(f"[CSV] [OK] 성공")

            return df



        except Exception as e:



            error_msg = str(e)[:100]  # 에러 메시지 짧게

            tried.append((cand, error_msg))

            try:

                print(f"[CSV] [FAIL] 실패: encoding={cand}, error={error_msg}")

            except UnicodeEncodeError:

                print(f"[CSV] [FAIL] 실패: {cand}")



    # 마지막 시도: pandas의 자동 감지 (encoding=None)

    try:

        df = pd.read_csv(path, encoding=None)

        try:

            print(f"[CSV] [OK] 성공: encoding=auto-detect")

        except UnicodeEncodeError:

            print(f"[CSV] [OK] 성공")

        return df

    except Exception as e:

        tried.append(("auto-detect", str(e)[:100]))

        try:

            print(f"[CSV] [FAIL] 실패: encoding=auto-detect")

        except UnicodeEncodeError:

            print(f"[CSV] [FAIL] 실패")



    raise ValueError(f"CSV 인코딩 감지 실패. 파일: {path}\n시도한 인코딩: {[t[0] for t in tried]}")







# ------------------ 이미지 튜닝(대표 배경색 + 초미세 변형 + 좌우반전) ------------------







def _load_logo(logo_path: str):



    if not logo_path or not os.path.isfile(logo_path): return None



    try: return Image.open(logo_path).convert("RGBA")



    except Exception: return None







def _apply_logo(canvas: Image.Image, logo_rgba: Image.Image,



                pos: str = "br", opacity: int = 65, margin: int = 20, logo_ratio: int = 14):



    if logo_rgba is None: return canvas



    W, H = canvas.size; base = canvas.convert("RGBA")



    target_w = max(1, int(min(W, H) * (logo_ratio / 100.0)))



    scale = target_w / float(logo_rgba.width)



    logo = logo_rgba.resize((target_w, max(1, int(logo_rgba.height * scale))), Image.LANCZOS)



    if opacity < 100:



        alpha = logo.split()[3] if logo.mode == "RGBA" else Image.new("L", logo.size, 255)



        alpha = ImageEnhance.Brightness(alpha).enhance(opacity / 100.0)



        logo.putalpha(alpha)



    if   pos == "br": x = W - logo.width - margin; y = H - logo.height - margin



    elif pos == "tr": x = W - logo.width - margin; y = margin



    elif pos == "bl": x = margin; y = H - logo.height - margin



    elif pos == "tl": x = margin; y = margin



    else:             x = (W - logo.width)//2; y = (H - logo.height)//2



    base.alpha_composite(logo, dest=(x, y))



    return base.convert("RGB")







def _dominant_edge_color(img: Image.Image, edge: int = 12, white_clip: int = 248, black_clip: int = 6):



    im = img.convert("RGB"); w, h = im.size



    strips = [im.crop((0,0,w,edge)), im.crop((0,h-edge,w,h)), im.crop((0,0,edge,h)), im.crop((w-edge,0,w,h))]



    arrs = [np.asarray(s, dtype=np.uint8).reshape(-1, 3) for s in strips]



    pix = np.concatenate(arrs, axis=0)



    Y = 0.2126*pix[:,0] + 0.7152*pix[:,1] + 0.0722*pix[:,2]



    keep = (Y < white_clip) & (Y > black_clip)



    core = pix[keep] if keep.any() else pix



    r, g, b = np.median(core, axis=0)



    return (int(r), int(g), int(b))







def _auto_trim_near_bg(img: Image.Image, bg=(255,255,255), tol=6):



    im = img.convert("RGB"); bg_im = Image.new("RGB", im.size, bg)



    diff = ImageChops.difference(im, bg_im).convert("L").point(lambda p: 0 if p <= tol else p)



    bbox = diff.getbbox()



    return im.crop(bbox) if bbox else im







def _to_square_canvas(img: Image.Image, size: int = 1200, bg=(255,255,255), pad: int = 20):



    try: img = ImageOps.exif_transpose(img)



    except Exception: pass



    img = img.convert("RGB"); w, h = img.size



    inner = max(1, size - pad*2)



    if w >= h:



        new_w = inner; new_h = max(1, int(h * (inner / float(w))))



    else:



        new_h = inner; new_w = max(1, int(w * (inner / float(h))))



    return img.resize((new_w, new_h), Image.LANCZOS)







def _affine_translate(img: Image.Image, tx: float, ty: float, fill):



    return img.transform(img.size, Image.AFFINE, (1,0,-tx, 0,1,-ty), resample=Image.BICUBIC, fillcolor=fill)







def _gentle_augment(img: Image.Image,



                    use_auto_contrast=True,



                    use_sharpen=True,



                    use_small_rotate=True,



                    rotate_zoom=1.04,



                    bg=(255,255,255),



                    ultra_angle_deg=0.35,



                    ultra_translate_px=0.6,



                    ultra_scale_pct=0.25,



                    trim_tol=8,



                    do_flip_lr=False):



    out = img.convert("RGB")



    if do_flip_lr:



        if random.random() < 0.5:



            out = ImageOps.mirror(out)



    # 색감 보존: autocontrast 대신 눈에 안 보이는 미세 노이즈(±2)로 픽셀값 변형

    if use_auto_contrast:

        arr = np.asarray(out, dtype=np.int16)

        noise = np.random.randint(-2, 3, size=arr.shape, dtype=np.int16)

        arr = np.clip(arr + noise, 0, 255).astype(np.uint8)

        out = Image.fromarray(arr, "RGB")



    # 색감 보존: 샤프닝을 아주 약하게 (percent 30, threshold 높게)

    if use_sharpen: out = out.filter(ImageFilter.UnsharpMask(radius=0.5, percent=30, threshold=4))



    bg_color = tuple(bg) if isinstance(bg, tuple) else _dominant_edge_color(out)



    w, h = out.size



    if use_small_rotate:



        angle = random.uniform(-ultra_angle_deg, ultra_angle_deg)



        if rotate_zoom and rotate_zoom > 1.0:



            out = out.resize((max(1,int(w*rotate_zoom)), max(1,int(h*rotate_zoom))), Image.LANCZOS)



            w2, h2 = out.size



        else:



            w2, h2 = w, h



        rot = out.rotate(angle, resample=Image.BICUBIC, expand=True, fillcolor=bg_color)



        out = ImageOps.fit(rot, (w2, h2), method=Image.LANCZOS, centering=(0.5, 0.5))



        out = ImageOps.fit(out, (w, h), method=Image.LANCZOS, centering=(0.5, 0.5))



    tx = random.uniform(-ultra_translate_px, ultra_translate_px)



    ty = random.uniform(-ultra_translate_px, ultra_translate_px)



    out = _affine_translate(out, tx, ty, bg_color)



    s = 1.0 + random.uniform(-ultra_scale_pct, ultra_scale_pct) / 100.0



    if abs(s - 1.0) > 0.0005:



        new_w = max(1, int(round(w * s))); new_h = max(1, int(round(h * s)))



        scaled = out.resize((new_w, new_h), Image.LANCZOS)



        out = ImageOps.fit(scaled, (w, h), method=Image.LANCZOS, centering=(0.5, 0.5))



    out = _auto_trim_near_bg(out, bg=bg_color, tol=trim_tol)



    out._matched_bg = bg_color



    return out







def _compose_on_square_canvas(img_resized: Image.Image, size: int = 1200, bg=(255,255,255), pad: int = 20):



    canvas = Image.new("RGB", (size, size), bg)



    w, h = img_resized.size



    x = (size - w)//2; y = (size - h)//2



    canvas.paste(img_resized, (x, y))



    return canvas







def _extract_gs_code_from_name(path_or_name: str):



    m = re.search(r'(GS\d{7})', os.path.basename(path_or_name), flags=re.IGNORECASE)



    return m.group(1).upper() if m else "NO_GS"







def process_listing_images_global(src_paths: list,

                                  base_out_root: str,

                                  logo_rgba: Image.Image,

                                  size: int,

                                  pad: int,

                                  bg_color: tuple,

                                  pos: str,

                                  opacity: int,

                                  logo_ratio: int,

                                  use_auto_contrast: bool,

                                  use_sharpen: bool,

                                  use_small_rotate: bool,

                                  rotate_zoom: float,

                                  max_images_per_code: int,

                                  ultra_angle_deg: float,

                                  ultra_translate_px: float,

                                  ultra_scale_pct: float,

                                  trim_tol: int,

                                  jpeg_q_min: int,

                                  jpeg_q_max: int,

                                  do_flip_lr: bool = False,

                                  progress_cb=None):

    by_code = {}



    for p in src_paths:



        base = os.path.basename(p)



        code9 = _extract_gs_code_from_name(base)



        if not code9 or code9 == "NO_GS":



            code9_folder = _extract_gs_code_from_name(p)



            code9 = code9_folder if code9_folder and code9_folder != "NO_GS" else None



        if not code9: continue



        by_code.setdefault(code9, []).append(p)







    results = []



    for code9, items in by_code.items():



        items_sorted = sorted(items)[:max_images_per_code] if max_images_per_code > 0 else sorted(items)



        out_dir = os.path.join(base_out_root, code9)



        os.makedirs(out_dir, exist_ok=True)



        for i, src in enumerate(items_sorted, start=1):



            try:



                with Image.open(src) as im:



                    im = im.convert("RGB")



                    matched_bg = _dominant_edge_color(im)



                    resized = _to_square_canvas(im, size=size, bg=matched_bg, pad=pad)



                    augmented = _gentle_augment(



                        resized,



                        use_auto_contrast, use_sharpen, use_small_rotate,



                        rotate_zoom=rotate_zoom, bg=matched_bg,



                        ultra_angle_deg=ultra_angle_deg,



                        ultra_translate_px=ultra_translate_px,



                        ultra_scale_pct=ultra_scale_pct,



                        trim_tol=trim_tol,



                        do_flip_lr=do_flip_lr



                    )



                    bg_use = getattr(augmented, "_matched_bg", matched_bg)



                    canvas = _compose_on_square_canvas(augmented, size=size, bg=bg_use, pad=pad)



                    final = _apply_logo(canvas, logo_rgba, pos=pos, opacity=opacity, margin=pad, logo_ratio=logo_ratio)



                    q = random.randint(jpeg_q_min, jpeg_q_max)



                    subs = random.choice([1, 2])



                    out_name = f"{code9}_{i}.jpg"



                    out_path = os.path.join(out_dir, out_name)



                    final.save(out_path, format="JPEG", quality=q, subsampling=subs, optimize=True)

                    if progress_cb:

                        try:

                            progress_cb()

                        except Exception:

                            pass

                    results.append(out_path)



            except Exception as e:



                print("[대표이미지] 변환 실패:", src, e)



    return results







# ------------------ NAVER hintKeywords 안전 생성기 ------------------







def build_hint_keywords_for_naver(base_text: str, fallback_text: str = "") -> str:



    """네이버 keywordstool용 hintKeywords (최대 5개, 각 2~20자, 한글/숫자만, 콤마 구분)."""



    def tokenize_kor_num(s):



        s = re.sub(r"[^\uAC00-\uD7A3 0-9]", " ", s or "")



        s = re.sub(r"\s+", " ", s).strip()



        toks = s.split()



        out = []



        for t in toks:



            t = re.sub(r"\s+", "", t)



            if 2 <= len(t) <= 20:



                out.append(t)



        return out







    pool = []



    seen = set()



    for cand in tokenize_kor_num(base_text) + tokenize_kor_num(fallback_text):



        if cand not in seen:



            seen.add(cand)



            pool.append(cand)



        if len(pool) >= 5:



            break



    if not pool:



        pool = ["키워드"]



    return ",".join(pool[:5])







# ------------------ 네이버 쇼핑 자동완성 ------------------



NAVER_AC_URL = "https://ac.search.naver.com/nx/ac"

NAVER_AC_TIMEOUT = 5



def naver_shopping_autocomplete(query: str, max_results: int = 10) -> list:

    """네이버 자동완성 키워드 조회. API 키 불필요."""

    if not query or not query.strip():

        return []

    try:

        params = {

            "q": query.strip(),

            "con": 1,

            "frm": "nv",

            "ans": 2,

            "r_format": "json",

            "r_enc": "UTF-8",

            "q_enc": "UTF-8",

            "st": 100,

        }

        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

        resp = requests.get(NAVER_AC_URL, params=params, headers=headers, timeout=NAVER_AC_TIMEOUT)

        if resp.status_code != 200:

            return []

        data = resp.json()

        items = data.get("items") or []

        keywords = []

        seen = set()

        # items = [[["키워드1"], ["키워드2"], ...]]

        for group in items:

            if not isinstance(group, list):

                continue

            for entry in group:

                if isinstance(entry, list) and len(entry) > 0:

                    kw = str(entry[0]).strip()

                elif isinstance(entry, str):

                    kw = entry.strip()

                else:

                    continue

                kw = re.sub(r"<[^>]+>", "", kw)  # HTML 태그 제거

                kw = re.sub(r"\s+", "", kw)       # 공백 제거(합성어)

                if not kw or len(kw) < 2 or len(kw) > 20:

                    continue

                if kw not in seen:

                    seen.add(kw)

                    keywords.append(kw)

                if len(keywords) >= max_results:

                    break

            if len(keywords) >= max_results:

                break

        return keywords

    except Exception:

        return []





def get_autocomplete_keywords_for_product(product_name: str, max_queries: int = 2, max_results: int = 10) -> list:

    """상품명에서 핵심어를 추출하고 자동완성 호출, 결과를 합쳐서 반환."""

    name_clean = re.sub(r"GS\d{7}[A-Z0-9]*", "", str(product_name))

    name_clean = re.sub(r"[^가-힣\s]", " ", name_clean)

    tokens = [t for t in name_clean.split() if len(t) >= 2 and t not in STOPWORDS and t not in SIZE_WORDS]

    if not tokens:

        return []

    queries = tokens[:max_queries]

    all_kw = []

    seen = set()

    for q in queries:

        results = naver_shopping_autocomplete(q, max_results=max_results)

        for kw in results:

            if kw not in seen:

                seen.add(kw)

                all_kw.append(kw)

        time.sleep(0.3)

    return all_kw





# ------------------ 구글 자동완성 ------------------



GOOGLE_AC_URL = "https://suggestqueries.google.com/complete/search"

GOOGLE_AC_TIMEOUT = 5





def google_autocomplete(query: str, max_results: int = 10) -> list:

    """구글 자동완성 키워드 조회 (한국어). API 키 불필요."""

    if not query or not query.strip():

        return []

    try:

        params = {

            "q": query.strip(),

            "client": "firefox",

            "hl": "ko",

            "gl": "kr",

        }

        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

        resp = requests.get(

            GOOGLE_AC_URL, params=params, headers=headers, timeout=GOOGLE_AC_TIMEOUT

        )

        if resp.status_code != 200:

            return []

        resp.encoding = "utf-8"

        data = resp.json()

        if not isinstance(data, list) or len(data) < 2:

            return []

        raw_keywords = data[1]

        keywords = []

        seen = set()

        for kw in raw_keywords:

            kw = str(kw).strip()

            kw = re.sub(r"<[^>]+>", "", kw)

            kw = re.sub(r"\s+", "", kw)

            if not kw or len(kw) < 2 or len(kw) > 20:

                continue

            if kw not in seen:

                seen.add(kw)

                keywords.append(kw)

            if len(keywords) >= max_results:

                break

        return keywords

    except Exception:

        return []





def get_google_autocomplete_for_product(

    product_name: str, max_queries: int = 2, max_results: int = 10

) -> list:

    """상품명에서 핵심어를 추출하고 구글 자동완성 호출, 결과를 합쳐서 반환."""

    name_clean = re.sub(r"GS\d{7}[A-Z0-9]*", "", str(product_name))

    name_clean = re.sub(r"[^가-힣\s]", " ", name_clean)

    tokens = [

        t

        for t in name_clean.split()

        if len(t) >= 2 and t not in STOPWORDS and t not in SIZE_WORDS

    ]

    if not tokens:

        return []

    queries = tokens[:max_queries]

    all_kw = []

    seen = set()

    for q in queries:

        results = google_autocomplete(q, max_results=max_results)

        for kw in results:

            if kw not in seen:

                seen.add(kw)

                all_kw.append(kw)

        time.sleep(0.3)

    return all_kw





# ------------------ GUI ------------------







# ------------------ OCR 개별 함수 ------------------



def _pil_ensure_rgb(img: Image.Image) -> Image.Image:



    try: return img.convert("RGB")



    except Exception: return img







def _only_korean(s: str) -> str:



    s = re.sub(r"[^가-힣\s]", " ", s or "")



    return normalize_space(s)







def _tess_config(psm: int, oem: int) -> str:



    return f"--psm {int(psm)} --oem {int(oem)}"







def _ocr_pil_image(img: Image.Image, lang="kor+eng", psm=3, oem=3, korean_only=False) -> str:



    try:



        g = ImageOps.grayscale(_pil_ensure_rgb(img))



        w, h = g.size



        max_w = 1800



        if w < 900: g = g.resize((w*2, h*2), Image.LANCZOS)



        elif w > max_w:



            scale = max_w / float(w)



            g = g.resize((int(w*scale), int(h*scale)), Image.LANCZOS)



        g = ImageOps.autocontrast(g, cutoff=1)



        g = g.filter(ImageFilter.UnsharpMask(radius=1.2, percent=140, threshold=3))



        text = pytesseract.image_to_string(g, lang=lang, config=_tess_config(psm,oem))



        text = normalize_space(re.sub(r"[^\s0-9A-Za-z가-힣]", " ", text))



        if korean_only: text = _only_korean(text)



        return text



    except Exception:



        return ""







def ocr_image_url(url: str, lang="kor+eng", timeout=10, psm=3, oem=3, korean_only=False) -> str:



    try:



        headers = {"User-Agent": "Mozilla/5.0 (KeywordOCR/1.0)"}



        r = requests.get(url, headers=headers, timeout=timeout); r.raise_for_status()



        img = Image.open(io.BytesIO(r.content))



        return _ocr_pil_image(img, lang=lang, psm=psm, oem=oem, korean_only=korean_only)



    except Exception:



        return ""







def ocr_image_file(path: str, lang="kor+eng", psm=3, oem=3, korean_only=False) -> str:



    try:



        img = Image.open(path)



        return _ocr_pil_image(img, lang=lang, psm=psm, oem=oem, korean_only=korean_only)



    except (UnidentifiedImageError, FileNotFoundError, OSError, ValueError):



        return ""







def summarize_features_tokens(text: str, max_len: int = 200):



    if not text: return "", []



    tokens = text.split()



    keep = []



    color_tokens = {"흰색","화이트","검정","블랙","회색","그레이","파랑","블루","초록","그린","빨강","레드",



                    "분홍","핑크","노랑","옐로우","브라운","베이지","실버","골드"}



    use_tokens = {"주방","욕실","가정","사무실","카페","세면대","싱크대","원룸","자취","청소","보관","정리",



                  "방수","미끄럼방지","위생","친환경","내열","내구성","튼튼","얇은","두꺼운","투명","불투명"} | color_tokens



    for t in tokens:



        if re.search(r"\d", t): keep.append(t); continue



        if t in use_tokens: keep.append(t)



    return normalize_space(" ".join(keep))[:max_len], keep







# ------------------ 엔트리 ------------------

