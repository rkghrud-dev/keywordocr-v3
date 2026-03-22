"""OCR 텍스트에서 반복되는 불필요한 문구를 필터링하는 모듈.

1. 하드코딩된 불용 문구 리스트 (배송안내, 주의사항, 제조국 등)
2. 자동 학습 DB (JSON) - 여러 상품에서 반복 등장하는 문구를 자동 축적
"""
from __future__ import annotations

import json
import os
import re
from collections import Counter
from datetime import date

# ── 프로젝트 루트 기준 data 폴더 ──
_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
_DATA_DIR = os.path.join(_PROJECT_ROOT, "data")
LEARNED_DB_PATH = os.path.join(_DATA_DIR, "ocr_learned_noise.json")

# ── 하드코딩 불용 문구 (포함되면 해당 문장 제거) ──
BOILERPLATE_PHRASES: set[str] = {
    # 면책/고지
    "구성이 다를 수 있",
    "색이 다를 수 있",
    "색상이 다를 수 있",
    "실제와 다를 수 있",
    "다를 수 있습니다",
    "다르게 보일 수 있",
    "다르게 보일수있",
    "모니터에 따라",
    "화면과 실제 색상",
    "촬영 환경에 따라",
    "실물과 차이",
    "사이즈에 약간의 오차",
    "오차가 발생할 수 있",
    "측정방법에 따라",
    "측정 방법에 따라",
    # 참고/주의/안내
    "참고사항",
    "참고 사항",
    "주의사항",
    "주의 사항",
    "취급주의",
    "취급 주의",
    "사용상 주의",
    "사용 시 주의",
    "안전사용",
    "주의 및 안내",
    "유의사항",
    "유의 사항",
    # 제조/원산지/수입
    "제조국",
    "제조일자",
    "원산지",
    "유통기한",
    "제조원",
    "판매원",
    "수입원",
    "수입사",
    "품질보증기준",
    "제조사",
    "제조업체",
    "수입업체",
    # 회사명/브랜드명
    "굿셀러스",
    "goodsellers",
    # 배송
    "배송안내",
    "배송 안내",
    "배송비",
    "무료배송",
    "출고 후",
    "택배비",
    "왕복 택배비",
    "배송기간",
    "영업일 이내",
    "일 이내 발송",
    "배송 지역",
    "도서산간",
    # 교환/반품
    "교환 및 반품",
    "교환및반품",
    "교환 반품",
    "반품 안내",
    "반품안내",
    "단순변심",
    "개봉 후에는",
    "반품 불가",
    "교환 불가",
    "환불 불가",
    "반품교환",
    "교환반품",
    # 법적/인증
    "전자상거래",
    "소비자분쟁해결",
    "공정거래위원회",
    "인증번호",
    "안전기준 적합",
    "KC인증",
    "소비자보호",
    # 고객지원
    "고객센터",
    "고객 센터",
    "불만족시",
    "문의사항",
    "문의 사항",
    "A/S",
    "AS안내",
    # 디스플레이/기종 면책
    "사용자의 기종",
    "디스플레이 설정",
    "사진은 제품의 실제",
    "사진은제품의 실제",
    # 불필요한 상세페이지 반복 라벨
    "용품 시리즈",
    "중선택",
    "중 선택",
    "옵션 중 선택",
    "낱개판매",
    "낱개로 판매",
    "낱개 판매",
    "판매구성에 포함되지",
    "구성에 포함되지 않",
    "비포함",
    # 수량/단위 라벨 (상세페이지 정보)
    "수량 개",
    "수량 세트",
    "수량 팩",
    # 원산지 값
    "중국",
    "중국산",
    "대한민국",
    "한국",
    "베트남",
    "인도네시아",
    "태국",
    "일본",
    "대만",
    # 기타 면책 잔여 문구
    "차이로 인해",
    "설정값 차이",
    "설정값차이",
    # 설치 안내/옵션 라벨
    "OPTION",
    "POINT",
    "좌우 방향 조절 방법",
    "먼저 나사를 풀어주세요",
    "잠금 고리를 빼세요",
    "방향을 바꿔서 다시 넣으세요",
    "나사를 다시 조여주면 완료",
    "불안하신가요",
    "남녀노소 누구나",
    "드라이버 하나면",
    "손쉽게 교체할 수",
    "좌우 구분 없는",
    "만능 설계",
    "DIY 초보자도",
    "부드러운 슬라이딩 작동",
    "Window screen lock hook",
}

# ── 정규식 기반 제거 패턴 ──
# 줄 단위로 매칭하여 해당 줄 전체 제거
NOISE_REGEX_PATTERNS: list[re.Pattern] = [
    re.compile(r"제조국\s*[:：]?\s*\S+", re.IGNORECASE),        # 제조국 중국, 제조국: 한국 등
    re.compile(r"수입사\s*[:：]?\s*\S+", re.IGNORECASE),        # 수입사 굿셀러스
    re.compile(r"원산지\s*[:：]?\s*\S+", re.IGNORECASE),        # 원산지 중국
    re.compile(r"제조사\s*[:：]?\s*\S+", re.IGNORECASE),        # 제조사 xxx
    re.compile(r"판매원\s*[:：]?\s*\S+", re.IGNORECASE),        # 판매원 xxx
    re.compile(r"수입원\s*[:：]?\s*\S+", re.IGNORECASE),        # 수입원 xxx
    re.compile(r"A/?S\s*[:：]?\s*\d", re.IGNORECASE),           # A/S 1544-xxxx
    re.compile(r"고객센터\s*[:：]?\s*\d", re.IGNORECASE),       # 고객센터 1544-xxxx
    re.compile(r"tel\s*[:：]?\s*\d", re.IGNORECASE),            # TEL: 02-xxxx
    re.compile(r"사이즈\s+\S+\s*[xX×]\s*\S+", re.IGNORECASE),  # 사이즈 10x20 (스펙 나열)
]

INLINE_NOISE_PATTERNS: list[re.Pattern] = [
    re.compile(r"\bOPTION\b", re.IGNORECASE),
    re.compile(r"\bPOINT\b", re.IGNORECASE),
    re.compile(r"[AB]\.\s*(?:블랙|화이트(?:\(미색\))?)", re.IGNORECASE),
    re.compile(r"좌우\s*방향(?:\s*조절)?\s*방법", re.IGNORECASE),
    re.compile(r"먼저\s*나사를?\s*풀어주세요", re.IGNORECASE),
    re.compile(r"잠금\s*고리를?\s*빼(?:세요)?", re.IGNORECASE),
    re.compile(r"방향을\s*바꿔서\s*다시\s*넣(?:으세요|으)?", re.IGNORECASE),
    re.compile(r"나사를?\s*다시\s*조여주면\s*완료", re.IGNORECASE),
    re.compile(r"방충망\(?창문\)?이\s*자꾸\s*열려서\s*불안하신가요", re.IGNORECASE),
    re.compile(r"남녀노소\s*누구나", re.IGNORECASE),
    re.compile(r"드라이버\s*하나면", re.IGNORECASE),
    re.compile(r"손쉽게\s*교체할\s*수\s*있(?:는|습니다)", re.IGNORECASE),
    re.compile(r"좌우\s*구분\s*없는", re.IGNORECASE),
    re.compile(r"만능\s*설계", re.IGNORECASE),
    re.compile(r"DIY\s*초보자도", re.IGNORECASE),
    re.compile(r"부드러운\s*슬라이딩\s*작동", re.IGNORECASE),
    re.compile(r"Window\s*screen\s*lock\s*hook", re.IGNORECASE),
    re.compile(r"기본장착된\s*금속고리의\s*나사를?\s*풀어", re.IGNORECASE),
    re.compile(r"집어\s*끼우기만\s*하면", re.IGNORECASE),
    re.compile(r"오랜기간\s*검증된", re.IGNORECASE),
    re.compile(r"힘들이지\s*않고", re.IGNORECASE),
]


# ────────────────────────────────────────
# 학습 DB 관리
# ────────────────────────────────────────

def _empty_db() -> dict:
    return {"version": 1, "phrases": {}, "total_products": 0}


def load_learned_db(path: str = LEARNED_DB_PATH) -> dict:
    if not os.path.isfile(path):
        return _empty_db()
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict) or "phrases" not in data:
            return _empty_db()
        return data
    except Exception:
        return _empty_db()


def save_learned_db(db: dict, path: str = LEARNED_DB_PATH) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(db, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def get_all_noise_phrases(db: dict | None = None) -> set[str]:
    """하드코딩 + 학습된 문구를 합친 전체 불용 문구 세트 반환."""
    phrases = set(BOILERPLATE_PHRASES)
    if db is None:
        db = load_learned_db()
    for phrase in db.get("phrases", {}):
        phrases.add(phrase)
    return phrases


# ────────────────────────────────────────
# 텍스트 문장 분할
# ────────────────────────────────────────

def _split_sentences(text: str) -> list[str]:
    """OCR 텍스트를 문장/줄 단위로 분할.

    분할 기준:
    - 줄바꿈
    - 공백 2개 이상
    - 마침표/느낌표/물음표 뒤 공백
    """
    # 줄바꿈으로 먼저 분할
    lines = re.split(r"[\n\r]+", text.strip())
    sentences = []
    for line in lines:
        # 공백 2개 이상으로 추가 분할
        parts = re.split(r"\s{2,}", line.strip())
        for part in parts:
            p = part.strip()
            if p:
                sentences.append(p)
    return sentences


# ────────────────────────────────────────
# 필터 함수
# ────────────────────────────────────────

def filter_ocr_text(text: str, noise_phrases: set[str] | None = None) -> str:
    """OCR 텍스트에서 불용 문구를 제거.

    2단계 필터:
    1) 문장 단위로 나뉘면 — 불용 문구 포함된 문장 전체 제거
    2) 문장이 하나로 합쳐져 있으면 — 불용 문구 부분만 직접 제거

    Args:
        text: 원본 OCR 텍스트
        noise_phrases: 불용 문구 세트. None이면 하드코딩+학습DB에서 로드.

    Returns:
        필터링된 텍스트
    """
    if not text or not text.strip():
        return ""

    if noise_phrases is None:
        db = load_learned_db()
        noise_phrases = get_all_noise_phrases(db)

    # 1단계: 문장 분할이 가능한 경우 문장 단위 필터링
    sentences = _split_sentences(text)
    if len(sentences) > 1:
        clean = []
        for sent in sentences:
            if any(phrase in sent for phrase in noise_phrases):
                continue
            if any(pat.search(sent) for pat in NOISE_REGEX_PATTERNS):
                continue
            if len(sent.strip()) <= 2:
                continue
            clean.append(sent)
        result = " ".join(clean)
    else:
        # 2단계: 문장이 하나로 합쳐진 경우 — 불용 문구를 직접 제거
        result = text

    # 항상: 불용 문구와 주변 컨텍스트를 직접 제거
    # 불용 문구 + 앞뒤로 붙은 관련 텍스트 제거
    for phrase in sorted(noise_phrases, key=len, reverse=True):
        if phrase in result:
            # 불용 문구 + 뒤에 이어지는 짧은 값(최대 20자)까지 제거
            # 예: "제조국 중국" → 전부 제거, "수입사 굿셀러스" → 전부 제거
            pattern = re.escape(phrase) + r"[\s:：]*[^\s]{0,20}"
            result = re.sub(pattern, " ", result)

    # 정규식 패턴도 직접 제거
    for pat in NOISE_REGEX_PATTERNS:
        result = pat.sub(" ", result)
    for pat in INLINE_NOISE_PATTERNS:
        result = pat.sub(" ", result)

    # 공백 정리
    result = re.sub(r"\s+", " ", result).strip()

    # 잔여 불필요 단어 제거
    _junk_words = {"인해", "있습니다", "됩니다", "않습니다", "드립니다", "바랍니다", "입니다"}
    words = result.split()
    words = [w for w in words if len(w) > 1 and w not in _junk_words]
    result = " ".join(words)

    return result


# ────────────────────────────────────────
# 자동 학습
# ────────────────────────────────────────

def preprocess_ocr_for_llm(ocr_raw_text: str) -> str:
    """Normalize OCR text for LLM while preserving product-identifying content."""
    if not ocr_raw_text:
        return ""

    text = str(ocr_raw_text).replace("\r", "\n")
    lines = [ln.strip() for ln in text.split("\n") if ln and ln.strip()]

    # Use unicode escapes to avoid source-encoding issues.
    drop_contains = [
        "\uc0ac/\uc774/\uc988/\ucc38/\uc870",
        "\uc0c9/\uc0c1/\ucc38/\uc870",
        "\uc635/\uc158/\uc120/\ud0dd",
        "PRODUCT DETAILS",
        "Installation",
        "\uac04\ub2e8 \uc124\uce58 \ubc29\ubc95",
        "\uc804\uc6d0 \ucc28\ub2e8",
        "\ub4dc\ub77c\uc774\ubc84",
        "\ucc28\ub2e8\uae30",
        "\uc2a4\uc704\uce58",
        "\uae30\uc874 \ud50c\ub808\uc774\ud2b8",
        "\uad6d\ub0b4\ucc3d\uace0",
        "\uc6d0\uc0b0\uc9c0",
        "\uc218\uc785\uc6d0",
        "\ud328\ud0a4\uc9c0",
        "\uc0c1\ud488\uc785\ub2c8\ub2e4",
        "\uc608\uc2dc",
        "\ubc30\uc1a1\ub429\ub2c8\ub2e4",
        "\uc218\ub3d9\uce21\uc815",
        "\uc624\ucc28",
    ]
    symbol_noise = ["C ||", "\u30b3", "oooo", "0000", "&//", "II", "Tz"]
    size_line_pat = re.compile(r"^\s*[\d\.,]+(?:\s*(?:mm|cm|\ub3c4))?\s*$", re.IGNORECASE)

    cleaned = []
    digit_run = 0
    for ln in lines:
        if any(x in ln for x in drop_contains):
            continue
        if any(x in ln for x in symbol_noise):
            continue

        pure = re.sub(r"\s+", "", ln)
        digit_ratio = (sum(ch.isdigit() for ch in pure) / len(pure)) if pure else 0.0

        # Drop size-table heavy lines.
        if size_line_pat.match(ln) or digit_ratio >= 0.7:
            digit_run += 1
            continue
        if re.search(r"\d+\s*(mm|cm|\ub3c4)", ln, re.IGNORECASE):
            digit_run += 1
            if digit_run >= 2:
                continue
        else:
            digit_run = 0

        if re.fullmatch(r"[\W_]+", ln):
            continue
        if len(ln) <= 1:
            continue

        cleaned.append(ln)

    out = " ".join(cleaned)
    out = re.sub(r"[|/\\{}\[\]<>]+", " ", out)
    out = re.sub(r"\s+", " ", out).strip()
    for pat in INLINE_NOISE_PATTERNS:
        out = pat.sub(" ", out)
    out = re.sub(r"\s+", " ", out).strip()

    # Safety fallback: never return empty for non-empty OCR input.
    if not out:
        out = re.sub(r"\s+", " ", str(ocr_raw_text)).strip()
    return out
def learn_from_batch(
    results: list[dict],
    db: dict | None = None,
    min_frequency_ratio: float = 0.5,
    min_phrase_len: int = 4,
    max_phrase_len: int = 30,
    min_products: int = 3,
) -> dict:
    """배치 OCR 결과에서 반복 문구를 감지하여 학습 DB에 추가.

    Args:
        results: OCR 파이프라인 결과 리스트 [{gs_code, raw_text, ...}, ...]
        db: 기존 학습 DB (None이면 파일에서 로드)
        min_frequency_ratio: 등장 비율 임계값 (0.5 = 50% 이상)
        min_phrase_len: 최소 문구 길이 (자)
        max_phrase_len: 최대 문구 길이 (자)
        min_products: 최소 상품 수 (이보다 적으면 학습 건너뜀)

    Returns:
        갱신된 학습 DB
    """
    if db is None:
        db = load_learned_db()

    # raw_text가 있는 결과만 수집
    texts = [r.get("raw_text", "") for r in results if r.get("raw_text", "").strip()]
    if len(texts) < min_products:
        return db

    # 상품별 문장 추출 (중복 제거: 같은 상품 내 동일 문장은 1회로)
    product_segments: list[set[str]] = []
    for txt in texts:
        segs = set()
        for seg in _split_sentences(txt):
            seg_norm = re.sub(r"\s+", " ", seg).strip()
            if min_phrase_len <= len(seg_norm) <= max_phrase_len:
                segs.add(seg_norm)
        product_segments.append(segs)

    # 문장별 등장 상품 수 카운트
    seg_counter: Counter = Counter()
    for segs in product_segments:
        for seg in segs:
            seg_counter[seg] += 1

    # 임계값 이상 등장하는 문장을 학습 DB에 추가
    threshold = max(min_products, int(len(texts) * min_frequency_ratio))
    today = date.today().isoformat()
    phrases_dict = db.setdefault("phrases", {})
    new_count = 0

    for seg, cnt in seg_counter.items():
        if cnt >= threshold:
            if seg in phrases_dict:
                phrases_dict[seg]["count"] = phrases_dict[seg].get("count", 0) + cnt
            else:
                phrases_dict[seg] = {"count": cnt, "first_seen": today}
                new_count += 1

    db["total_products"] = db.get("total_products", 0) + len(texts)

    if new_count > 0:
        save_learned_db(db)

    return db
