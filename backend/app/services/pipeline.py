from __future__ import annotations



from dataclasses import dataclass

from datetime import datetime

from concurrent.futures import ThreadPoolExecutor, as_completed

import base64

import json

import os

import re

import time



import numpy as np

import pandas as pd



from . import legacy_core as core

from .keyword_builder import build_keyword_string

from .market_keywords import generate_market_keyword_packages

from .env_loader import ensure_env_loaded, get_env





@dataclass

class PipelineConfig:

    file_path: str

    img_tag: str = ""

    tesseract_path: str = ""



    model_keyword: str = "gpt-4.1"        # 키워드 생성용

    model_longtail: str = "gpt-4.1-mini"      # 롱테일/R열 키워드용

    model_keyword_stage2: str = "" # 비우면 1차 모델 재사용



    max_words: int = 24

    max_len: int = 140

    min_len: int = 90



    use_html_ocr: bool = False

    use_local_ocr: bool = True

    merge_ocr_with_name: bool = True



    max_imgs: int = 999

    threads: int = 6

    max_depth: int = -1

    local_img_dir: str = ""

    allow_folder_match: bool = True



    korean_only: bool = True

    drop_digits: bool = True

    psm: int = 11

    oem: int = 3



    ocr_excel_path: str = ""             # 미리 처리된 OCR 결과 Excel 경로



    write_to_r: bool = True



    debug: bool = True



    naver_enabled: bool = False

    naver_dry_run: bool = False

    naver_retry: bool = False

    naver_retry_count: int = 2

    naver_retry_delay: float = 0.8



    naver_autocomplete: bool = False

    google_autocomplete: bool = True



    make_listing: bool = True

    listing_size: int = 1200

    listing_pad: int = 20

    listing_max: int = 20



    logo_path: str = ""

    logo_ratio: int = 14

    logo_opacity: int = 65

    logo_pos: str = "tr"



    use_auto_contrast: bool = True

    use_sharpen: bool = True

    use_small_rotate: bool = True

    rotate_zoom: float = 1.04



    ultra_angle_deg: float = 0.35

    ultra_translate_px: float = 0.6

    ultra_scale_pct: float = 0.25



    trim_tol: int = 8

    jpeg_q_min: int = 88

    jpeg_q_max: int = 92

    do_flip_lr: bool = True

    phase: str = "full"          # "full" | "images" | "analysis"
    export_root_override: str = ""  # phase=analysis 시 Phase1의 export_root 재사용



def _status(cb, msg: str) -> None:

    if cb:

        cb(msg)



def _progress(cb, value: int) -> None:

    if cb:

        cb(int(value))





def _format_naver_keyword_table(items: list, limit: int = 15) -> str:

    if not items:

        return ""

    rows = []

    for it in items:

        kw = str(it.get("relKeyword") or "").strip()

        if not kw:

            continue

        pc = int(it.get("monthlyPcQcCnt") or 0)

        mo = int(it.get("monthlyMobileQcCnt") or 0)

        total = pc + mo

        rows.append((kw, pc, mo, total))

    rows.sort(key=lambda x: x[3], reverse=True)

    lines = ["키워드|PC|MO|합계"]

    for kw, pc, mo, total in rows[: max(1, int(limit))]:

        lines.append(f"{kw}|{pc}|{mo}|{total}")

    return "\n".join(lines)

def run_pipeline(cfg: PipelineConfig, status_cb=None, progress_cb=None) -> tuple[str, str]:

    _status(status_cb, "🚀 run_pipeline 함수 시작!")



    if not cfg.file_path:

        raise ValueError("CSV/Excel 파일을 선택해 주세요.")



    csv_base = os.path.splitext(os.path.basename(cfg.file_path))[0]

    date_tag = datetime.now().strftime("%Y%m%d")

    # phase=analysis 시 Phase1의 export_root 재사용
    if cfg.export_root_override and os.path.isdir(cfg.export_root_override):
        export_root = cfg.export_root_override
    else:
        export_root = os.path.join("C:\\code", "exports", f"{date_tag}_{csv_base}")

    os.makedirs(export_root, exist_ok=True)

    _status(status_cb, f"📁 작업 폴더: {export_root} (phase={cfg.phase})")



    # Google Cloud Vision API 인증 설정 (통합 파이프라인에서 OCR 사용 시)

    _app_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

    ensure_env_loaded(os.path.join(_app_root, ".env"))

    _gv_cred_path = get_env("GOOGLE_APPLICATION_CREDENTIALS", "GOOGLE_VISION_CREDENTIALS")

    if not _gv_cred_path:

        _gv_cred_path = os.path.join(_app_root, "google_vision_key.json")

    if _gv_cred_path and os.path.isfile(_gv_cred_path):

        os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = _gv_cred_path

        _status(status_cb, f"Google Vision API 인증 설정 완료: {_gv_cred_path}")

    else:

        _status(status_cb, f"Google Vision 키 파일 없음: {_gv_cred_path}")



    if cfg.use_html_ocr or cfg.use_local_ocr:

        detected = core.setup_tesseract(cfg.tesseract_path or None)

        if not detected:

            _status(status_cb, "Tesseract 경로를 찾지 못했습니다. Google Vision API로 OCR 진행합니다.")



    # 네이버 검색광고 API/자동완성은 현재 완전 비활성화

    cfg.naver_enabled = False

    cfg.naver_autocomplete = False

    core.DRY_RUN = True

    naver_keys = {"ACCESS_LICENSE": "", "SECRET_KEY": "", "CUSTOMER_ID": ""}

    _status(status_cb, "네이버 API 비활성화: 검색어설정은 내부/GPT 로직만 사용")



    _status(status_cb, "처리중... (1/2) 전처리 + OCR + 키워드 생성")

    _progress(progress_cb, 10)

    if core.client is None:

        if core.refresh_openai_client():

            _status(status_cb, "AI client 재로드 완료")

        else:

            _status(status_cb, "AI client 없음: .env에 ANTHROPIC_API_KEY 또는 OPENAI_API_KEY 확인 필요")



    # 사용자 제외 단어 리로드 (실행 사이에 추가된 것 반영)

    core.merge_user_stopwords()



    max_words = max(5, int(cfg.max_words))

    max_len = max(30, int(cfg.max_len))

    min_len = max(0, int(cfg.min_len))

    if min_len > max_len:

        min_len = max_len // 2



    # min_len 달성에 필요한 최소 단어수 보정 (한국어 평균 ~3.8자/단어)

    _min_words_for_len = int(min_len / 3.5) + 1 if min_len > 0 else max_words

    if max_words < _min_words_for_len:

        max_words = min(_min_words_for_len, 50)  # 최대 50단어 상한



    max_imgs = max(0, int(cfg.max_imgs))

    threads = min(16, max(1, int(cfg.threads)))

    max_depth = int(cfg.max_depth)



    listing_size = max(200, int(cfg.listing_size))

    listing_pad = max(0, int(cfg.listing_pad))

    listing_max = max(0, int(cfg.listing_max))

    logo_ratio = max(1, min(60, int(cfg.logo_ratio)))

    logo_opacity = max(0, min(100, int(cfg.logo_opacity)))

    logo_pos = cfg.logo_pos or "tr"

    use_auto_contrast = bool(cfg.use_auto_contrast)

    use_sharpen = bool(cfg.use_sharpen)

    use_small_rotate = bool(cfg.use_small_rotate)

    rotate_zoom = float(cfg.rotate_zoom)

    logo_rgba = core._load_logo(cfg.logo_path.strip())



    ultra_angle_deg = float(cfg.ultra_angle_deg)

    ultra_translate_px = float(cfg.ultra_translate_px)

    ultra_scale_pct = float(cfg.ultra_scale_pct)

    trim_tol = int(cfg.trim_tol)

    jpeg_q_min = max(70, min(99, int(cfg.jpeg_q_min)))

    jpeg_q_max = max(jpeg_q_min, min(99, int(cfg.jpeg_q_max)))



    do_flip_lr = bool(cfg.do_flip_lr)



    psm = int(cfg.psm or 3)

    oem = int(cfg.oem or 3)

    korean_only = bool(cfg.korean_only)

    drop_digits = bool(cfg.drop_digits)

    tess_lang = "kor" if korean_only else "kor+eng"



    df = core.safe_read_csv(cfg.file_path)

    if df.empty:

        raise ValueError("CSV/Excel 내용이 비어 있습니다.")

    input_cols = list(df.columns)  # 업로드용 저장 시 원본 컬럼만 유지



    # 불필요한 컬럼 제거 (원본 파일에서 유입되는 CM, Unnamed 등)

    _drop_cols = [c for c in df.columns

                  if str(c).strip().upper() == "CM"

                  or str(c).startswith("Unnamed")]

    if _drop_cols:

        df.drop(columns=_drop_cols, inplace=True, errors="ignore")



    name_col = "상품명"

    if name_col not in df.columns:

        raise ValueError("'상품명' 컬럼이 없습니다.")



    code_col = None

    for c in df.columns:

        if str(c).strip() in ["자체상품코드", "자체 상품코드", "상품코드B", "코드", "코드B"]:

            code_col = c

            break



    if str(df.columns[0]) == "상품코드" and len(df) > 0:

        # 상품코드 열의 자동값 제거 (A2 Pxxxxxx 제거 포함)

        df.iloc[:, 0] = ""

    if df.shape[1] >= 5:

        df.iloc[:, 4] = 23



    detail_col = next((col for col in df.columns if "상세" in str(col)), None)

    if detail_col and cfg.img_tag:

        df[detail_col] = df[detail_col].apply(lambda x: core.insert_img_tag(x, cfg.img_tag))



    # AU열 (이미지등록(상세)) 컬럼 확인 - 대표이미지 순차 다운로드용

    listing_img_col = None

    for c in df.columns:

        if "이미지등록" in str(c) or str(c).strip().upper() == "AU":

            listing_img_col = c

            break



    공급가_col = next((col for col in df.columns if "공급가" in str(col)), None)

    판매가_col = next((col for col in df.columns if "판매가" in str(col)), None)

    소비자가_col = next((col for col in df.columns if "소비자가" in str(col)), None)

    if not all([공급가_col, 판매가_col, 소비자가_col]):

        raise ValueError("공급가/판매가/소비자가 컬럼명을 찾을 수 없습니다.")

    df[공급가_col] = pd.to_numeric(df[공급가_col], errors="coerce").fillna(0)

    df[판매가_col] = df[공급가_col].apply(lambda v: np.round(v * core.get_multiplier(v), -2))

    df[소비자가_col] = np.round(df[판매가_col].astype(float) * 1.2, -2)



    df["_code9_from_name"] = df[name_col].astype(str).str.extract(r"(GS\d{7})")

    df["_opt_from_name"] = (

        df[name_col].astype(str)

        .str.replace(r".*GS\d{7}[A-Z0-9]+", "", regex=True)

        .str.strip()

    )



    # ── 옵션 가격 분리: 추가금 > 기본판매가 시 상품 자동 분리 ──

    _price_split_log = []

    for _gs9, _grp in df.groupby("_code9_from_name", dropna=True):

        if _gs9 is None or pd.isna(_gs9) or len(_grp) <= 1:

            continue

        _sells = _grp[판매가_col].values

        _base = float(_sells.min())

        if _base <= 0 or float(_sells.max()) <= _base * 2:

            continue

        _sorted_idx = _grp[판매가_col].sort_values().index.tolist()

        _bands = []

        _current_band = [_sorted_idx[0]]

        _band_base = float(df.at[_sorted_idx[0], 판매가_col])

        for _si in _sorted_idx[1:]:

            _sell = float(df.at[_si, 판매가_col])

            if _sell <= _band_base * 2:

                _current_band.append(_si)

            else:

                _bands.append(_current_band)

                _current_band = [_si]

                _band_base = _sell

        _bands.append(_current_band)

        if len(_bands) <= 1:

            continue

        for _bi, _band in enumerate(_bands):

            _band_sells = [float(df.at[_idx, 판매가_col]) for _idx in _band]

            _min_p = int(min(_band_sells))

            _max_p = int(max(_band_sells))

            _price_tag = f"({_min_p:,}~{_max_p:,}원)" if _min_p != _max_p else f"({_min_p:,}원)"

            for _ri, _idx in enumerate(_band):

                _new_letter = chr(65 + _ri)

                df.at[_idx, "_code9_from_name"] = f"{_gs9}-{_bi+1}"

                _orig_name = str(df.at[_idx, name_col])

                _orig_name = re.sub(r'(GS\d{7})[A-Z0-9]+', rf'\g<1>{_new_letter}', _orig_name, count=1)

                if _price_tag not in _orig_name:

                    _orig_name = f"{_orig_name} {_price_tag}"

                df.at[_idx, name_col] = _orig_name

                if code_col and code_col in df.columns:

                    _old_code = str(df.at[_idx, code_col])

                    _new_code = re.sub(r'(GS\d{7})[A-Z0-9]+', rf'\g<1>{_new_letter}', _old_code, count=1)

                    df.at[_idx, code_col] = _new_code

        _price_split_log.append(f"{_gs9} → {len(_bands)}개 상품으로 분리")

    if _price_split_log:

        _status(status_cb, f"⚠️ 가격 분리: {len(_price_split_log)}건")

        for _msg in _price_split_log:

            _status(status_cb, f"  {_msg}")



    def _resolve_columns(frame, candidate_names):

        matched = [c for c in frame.columns if str(c).strip() in candidate_names]

        if not matched:

            primary = candidate_names[0]

            frame[primary] = pd.Series([""] * len(frame), index=frame.index, dtype="string")

            matched = [primary]

        return matched



    def _set_row_values(frame, row_idx, candidate_names, value):

        for c in frame.columns:

            if str(c).strip() in candidate_names:

                frame.at[row_idx, c] = value



    option_name_map = {

        "옵션사용": ["옵션사용"],

        "옵션 구성방식": ["옵션 구성방식", "옵션구성방식"],

        "옵션 표시방식": ["옵션 표시방식", "옵션표시방식"],

        "옵션입력": ["옵션입력"],

    }



    for key, names in option_name_map.items():

        cols = _resolve_columns(df, names)

        for col in cols:

            df[col] = df[col].astype("string").fillna("")

            df[col] = "N" if key == "옵션사용" else ""



    # 품목구성방식/품목 구성방식 모두 유지, 기본 공란 (옵션 상품만 T)

    def _apply_item_comp_fix(frame):

        item_cols = _resolve_columns(frame, ["품목구성방식", "품목 구성방식"])

        for col in item_cols:

            frame[col] = ""

        return frame



    df = _apply_item_comp_fix(df)



    grouped = df.groupby("_code9_from_name", dropna=True)

    for code, group in grouped:

        if code is None or pd.isna(code):

            continue

        opts = []

        for i, (idx, row) in enumerate(group.iterrows()):

            option_code = chr(65 + i)

            option_val = str(row["_opt_from_name"]).strip()

            if option_val:

                opts.append(f"{option_code} {option_val}")

        if opts:

            ak_val = "옵션{" + "|".join(opts) + "}"

            aidx = group.index[0]

            _set_row_values(df, aidx, ["품목구성방식", "품목 구성방식"], "T")

            _set_row_values(df, aidx, ["옵션사용"], "Y")

            _set_row_values(df, aidx, ["옵션 구성방식", "옵션구성방식"], "T")

            _set_row_values(df, aidx, ["옵션 표시방식", "옵션표시방식"], "S")

            _set_row_values(df, aidx, ["옵션입력"], ak_val)

            # 옵션 추가금액 계산: A옵션(기본가) 대비 차액
            _base_sell = float(df.at[aidx, 판매가_col])
            _additionals = []
            for _oi, (_oidx, _orow) in enumerate(group.iterrows()):
                _opt_sell = float(df.at[_oidx, 판매가_col])
                _additionals.append(str(int(_opt_sell - _base_sell)))
            if "옵션추가금" not in df.columns:
                df["옵션추가금"] = ""
            df.at[aidx, "옵션추가금"] = "|".join(_additionals)





    df.drop(columns=["_code9_from_name", "_opt_from_name"], inplace=True, errors="ignore")



    name_s = df[name_col].astype(str)

    code_s = df[code_col].astype(str) if (code_col and code_col in df.columns) else pd.Series("", index=df.index, dtype="string")

    # 자체상품코드/상품명 모두 "GS + 7자리 + A로 끝나는 코드"만 허용

    mask1 = code_s.str.contains(r"GS\d{7}A$", na=False, regex=True)

    mask2 = name_s.str.contains(r"GS\d{7}A\b", na=False, regex=True)

    rep_mask = (mask1 | mask2)

    df_after = df.loc[rep_mask].copy()

    if df_after.empty:

        fallback_mask = name_s.str.contains(r"GS\d{7}", na=False, regex=True) | code_s.str.contains(r"GS\d{7}", na=False, regex=True)

        df_after = df.loc[fallback_mask].copy()



    # NOTE:

    # df 단계에서 옵션상품의 품목구성방식=T를 세팅했으므로,

    # 필터링된 df_after에서 다시 초기화하면 T가 사라진다.

    # 여기서는 컬럼 존재만 보장하고 기존 값은 유지한다.

    _resolve_columns(df_after, ["품목구성방식", "품목 구성방식"])



    if "검색어설정" not in df_after.columns:

        df_after["검색어설정"] = pd.Series([""] * len(df_after), index=df_after.index, dtype="string")

    else:

        df_after["검색어설정"] = df_after["검색어설정"].astype("string").fillna("")

    if "쿠팡검색태그" not in df_after.columns:

        df_after["쿠팡검색태그"] = pd.Series([""] * len(df_after), index=df_after.index, dtype="string")

    else:

        df_after["쿠팡검색태그"] = df_after["쿠팡검색태그"].astype("string").fillna("")

    if "네이버태그" not in df_after.columns:

        df_after["네이버태그"] = pd.Series([""] * len(df_after), index=df_after.index, dtype="string")

    else:

        df_after["네이버태그"] = df_after["네이버태그"].astype("string").fillna("")



    if "검색키워드" not in df_after.columns:

        df_after["검색키워드"] = pd.Series([""] * len(df_after), index=df_after.index, dtype="string")

    else:

        df_after["검색키워드"] = df_after["검색키워드"].astype("string").fillna("")

    if "OCR요약" not in df_after.columns:

        df_after["OCR요약"] = pd.Series([""] * len(df_after), index=df_after.index, dtype="string")

    else:

        df_after["OCR요약"] = df_after["OCR요약"].astype("string").fillna("")

    if "네이버검색광고데이터" not in df_after.columns:

        df_after["네이버검색광고데이터"] = pd.Series([""] * len(df_after), index=df_after.index, dtype="string")

    else:

        df_after["네이버검색광고데이터"] = df_after["네이버검색광고데이터"].astype("string").fillna("")

    if "1차키워드" not in df_after.columns:

        df_after["1차키워드"] = pd.Series([""] * len(df_after), index=df_after.index, dtype="string")

    else:

        df_after["1차키워드"] = df_after["1차키워드"].astype("string").fillna("")

    if "최종키워드2차" not in df_after.columns:

        df_after["최종키워드2차"] = pd.Series([""] * len(df_after), index=df_after.index, dtype="string")

    else:

        df_after["최종키워드2차"] = df_after["최종키워드2차"].astype("string").fillna("")

    _vision_cols = [

        "Vision힌트",

        "Vision분석JSON",

        "Vision_core_identity",

        "Vision_installation_and_physical",

        "Vision_usage_context",

        "Vision_market_expansion",

        "Vision_compatibility",

        "Vision_functional_inference",

        "Vision_search_boost_elements",

    ]

    for _c in _vision_cols:

        if _c not in df_after.columns:

            df_after[_c] = pd.Series([""] * len(df_after), index=df_after.index, dtype="string")

        else:

            df_after[_c] = df_after[_c].astype("string").fillna("")



    debug_rows = []

    debug_on = bool(cfg.debug)



    local_root = cfg.local_img_dir

    allow_folder_match = bool(cfg.allow_folder_match)

    ocr_temp_root = os.path.join(export_root, "_ocr_tmp")

    os.makedirs(ocr_temp_root, exist_ok=True)

    use_local = cfg.use_local_ocr and bool(detail_col)



    def ocr_paths(paths):

        """Google Cloud Vision API로 이미지 OCR 처리"""

        texts, raw_pairs = [], []

        if not paths:

            return texts, raw_pairs



        # Google Cloud Vision OCR 사용

        from app.services.ocr_pipeline import _ocr_google_vision



        with ThreadPoolExecutor(max_workers=threads) as ex:

            futs = {ex.submit(_ocr_google_vision, p): p for p in paths}

            for fut in as_completed(futs):

                src = futs[fut]

                try:

                    t = fut.result()

                    if t:

                        texts.append(t)

                        raw_pairs.append((os.path.basename(src), t[:200]))

                except Exception as e:

                    # Google Vision 실패 시 에러 로그

                    _status(status_cb, f"[OCR 실패] {os.path.basename(src)}: {str(e)[:50]}")

        return texts, raw_pairs



    def analyze_product_images_local(image_paths, product_name, model_name, min_fill_ratio=0.5):

        """대표이미지 Vision 분석(JSON). 실패 시 {} 반환."""

        try:

            if not core.client or not image_paths:

                return {}



            _target_paths = [

                "core_identity.category",

                "core_identity.product_type_correction",

                "core_identity.structure",

                "core_identity.material_visual",

                "core_identity.color",

                "core_identity.size_context",

                "installation_and_physical.mount_type",

                "installation_and_physical.installation_method",

                "installation_and_physical.environment_resistance",

                "installation_and_physical.durability_hint",

                "installation_and_physical.weight_feel",

                "usage_context.usage_location",

                "usage_context.usage_purpose",

                "usage_context.target_user",

                "usage_context.usage_scenario",

                "usage_context.indoor_outdoor",

                "market_expansion.emotion_tone",

                "market_expansion.design_style",

                "market_expansion.shape_motif",

                "market_expansion.seasonal_context",

                "market_expansion.trend_alignment",

                "compatibility.compatible_with",

                "compatibility.size_compatibility",

                "compatibility.device_fit",

                "functional_inference.primary_function",

                "functional_inference.secondary_function",

                "functional_inference.problem_solving_keyword",

                "functional_inference.convenience_feature",

                "search_boost_elements.installation_keywords",

                "search_boost_elements.space_keywords",

                "search_boost_elements.benefit_keywords",

                "search_boost_elements.longtail_candidates",

            ]



            valid_paths = [p for p in image_paths if os.path.isfile(p)]

            if not valid_paths:

                return {}



            def _build_image_contents(paths):

                out = []

                for img_path in paths:

                    with open(img_path, "rb") as f:

                        b64 = base64.b64encode(f.read()).decode("utf-8")

                    ext = os.path.splitext(img_path)[1].lower()

                    mime = "image/png" if ext == ".png" else "image/jpeg"

                    out.append({

                        "type": "image_url",

                        "image_url": {"url": f"data:{mime};base64,{b64}", "detail": "low"},

                    })

                return out



            def _is_filled(v):

                if isinstance(v, str):

                    return bool(v.strip())

                if isinstance(v, list):

                    return any(bool(str(x).strip()) for x in v)

                if isinstance(v, dict):

                    return any(_is_filled(x) for x in v.values())

                return bool(v)



            def _get_path(d, path):

                cur = d

                for k in path.split("."):

                    if not isinstance(cur, dict) or k not in cur:

                        return None

                    cur = cur[k]

                return cur



            def _fill_ratio(d):

                if not isinstance(d, dict):

                    return 0.0

                filled = 0

                for p in _target_paths:

                    if _is_filled(_get_path(d, p)):

                        filled += 1

                return filled / max(1, len(_target_paths))



            def _merge_analysis(base, extra):

                if not isinstance(base, dict):

                    return extra if isinstance(extra, dict) else {}

                if not isinstance(extra, dict):

                    return base

                merged = dict(base)

                for k, v in extra.items():

                    if k not in merged:

                        merged[k] = v

                        continue

                    bv = merged[k]

                    if isinstance(bv, dict) and isinstance(v, dict):

                        merged[k] = _merge_analysis(bv, v)

                    elif isinstance(bv, list) and isinstance(v, list):

                        merged_list = []

                        seen = set()

                        for item in bv + v:

                            s = str(item).strip()

                            if not s:

                                continue

                            lk = s.lower()

                            if lk in seen:

                                continue

                            seen.add(lk)

                            merged_list.append(s)

                        merged[k] = merged_list

                    else:

                        if not _is_filled(bv) and _is_filled(v):

                            merged[k] = v

                return merged



            cleaned_name = re.sub(r"GS\d{7}[A-Z0-9]*\s*", "", str(product_name)).strip()

            system_prompt = (

                "너는 이커머스 상품 이미지 분석가다. JSON만 출력. 설명문/마크다운/코드블록 금지.\n"

                "모든 값은 검색 키워드로 쓸 수 있는 짧은 명사구(1~4단어)로 작성.\n"

                "문장형 금지. 추론 불가면 빈값(\"\") 또는 빈배열([])."

            )

            user_text = (

                f"상품명: {cleaned_name}\n\n"

                "이미지를 보고 아래 JSON을 채워라.\n"

                "규칙: 모든 값은 짧은 명사구/키워드만. 문장 금지. 모르면 비워라.\n\n"

                "{"

                "\"core_identity\":{\"category\":\"\",\"product_type_correction\":\"\",\"structure\":\"\",\"material_visual\":\"\",\"color\":\"\",\"size_context\":\"\"},"

                "\"installation_and_physical\":{\"mount_type\":\"\",\"installation_method\":\"\",\"environment_resistance\":[],\"durability_hint\":\"\",\"weight_feel\":\"\"},"

                "\"usage_context\":{\"usage_location\":[],\"usage_purpose\":[],\"target_user\":[],\"usage_scenario\":[],\"indoor_outdoor\":\"\"},"

                "\"market_expansion\":{\"emotion_tone\":[],\"design_style\":[],\"shape_motif\":\"\",\"seasonal_context\":\"\",\"trend_alignment\":\"\"},"

                "\"compatibility\":{\"compatible_with\":[],\"size_compatibility\":[],\"device_fit\":\"\"},"

                "\"functional_inference\":{\"primary_function\":\"\",\"secondary_function\":[],\"problem_solving_keyword\":[],\"convenience_feature\":[]},"

                "\"search_boost_elements\":{\"installation_keywords\":[],\"space_keywords\":[],\"benefit_keywords\":[],\"longtail_candidates\":[]}"

                "}\n\n"

                "예시(참고용):\n"

                "category→\"브라켓\" / product_type_correction→\"차량조명 마운트 브라켓\"\n"

                "mount_type→\"무타공 고정형\" / installation_method→\"클램프 체결\"\n"

                "usage_location→[\"차량 본넷\",\"트렁크\"] / usage_purpose→[\"조명 설치\",\"작업등 장착\"]\n"

                "primary_function→\"무타공 조명 고정 각도조절\" / problem_solving_keyword→[\"무타공\",\"간편설치\"]\n"

                "longtail_candidates→[\"무타공 차량조명 브라켓\",\"본넷 작업등 거치대\"]"

            )



            # 1차: 최대 3장, 2차: 부족하면 최대 5장으로 재분석 후 병합

            stages = []

            first_n = min(3, len(valid_paths))

            if first_n > 0:

                stages.append(first_n)

            if len(valid_paths) > first_n:

                second_n = min(5, len(valid_paths))

                if second_n > first_n:

                    stages.append(second_n)



            best = {}

            vision_client = core._create_client(model_name) or core.client

            if vision_client is None:

                return {}

            for n in stages:

                image_contents = _build_image_contents(valid_paths[:n])

                user_msg = [{"type": "text", "text": user_text}] + image_contents

                resp = vision_client.chat.completions.create(

                    model=model_name,

                    messages=[

                        {"role": "system", "content": system_prompt},

                        {"role": "user", "content": user_msg},

                    ],

                    temperature=0.1,

                    top_p=0.9,

                    max_tokens=2048,

                    response_format={"type": "json_object"},

                )

                raw = (resp.choices[0].message.content or "").strip()

                try:

                    cur = json.loads(raw) if raw else {}

                except json.JSONDecodeError:

                    cur = {}

                    for _closer in ['"}}', '"]}}', '"]}}}', '"}}}', '"]}}}}}']:

                        try:

                            cur = json.loads(raw + _closer)

                            break

                        except json.JSONDecodeError:

                            continue

                best = _merge_analysis(best, cur)

                if _fill_ratio(best) >= float(min_fill_ratio):

                    break



            return best if isinstance(best, dict) else {}

        except Exception as _ve:

            _status(status_cb, f"[Vision 오류] {type(_ve).__name__}: {str(_ve)[:200]}")

            return {}



    def _extract_vision_hints(analysis):

        """구/신 Vision 스키마에서 키워드 힌트 문자열만 안전하게 추출."""

        hints, filled = [], []



        def _add_hint(path, text):

            s = re.sub(r"\s+", " ", str(text or "")).strip()

            if s:

                hints.append(s)

                filled.append(path)



        def _walk(node, path):

            if isinstance(node, dict):

                # 구 스키마: {"value": "...", "confidence": 0~1, "evidence": "..."}

                if "value" in node and any(k in node for k in ("confidence", "evidence")):

                    conf_ok = True

                    try:

                        conf_ok = float(node.get("confidence", 1.0)) >= 0.5

                    except Exception:

                        conf_ok = True

                    if conf_ok:

                        _add_hint(path, node.get("value", ""))

                    return

                for k, v in node.items():

                    _walk(v, f"{path}.{k}" if path else k)

                return

            if isinstance(node, list):

                had = False

                for x in node:

                    sx = re.sub(r"\s+", " ", str(x or "")).strip()

                    if sx:

                        hints.append(sx)

                        had = True

                if had:

                    filled.append(path)

                return

            if isinstance(node, str):

                _add_hint(path, node)



        _walk(analysis or {}, "")

        uniq, seen = [], set()

        for h in hints:

            key = h.lower()

            if key in seen:

                continue

            seen.add(key)

            uniq.append(h)

        uniq_filled = list(dict.fromkeys([p for p in filled if p]))

        return uniq[:40], uniq_filled



    def _vision_excel_payload(analysis, hint_parts):

        """Vision 분석 결과를 엑셀 저장용 문자열 컬럼으로 변환."""

        def _as_json(v):

            if v is None:

                return ""

            try:

                return json.dumps(v, ensure_ascii=False, separators=(",", ":"))

            except Exception:

                return str(v)



        a = analysis or {}

        return {

            "Vision힌트": " | ".join([x for x in (hint_parts or []) if str(x).strip()])[:1000],

            "Vision분석JSON": _as_json(a),

            "Vision_core_identity": _as_json(a.get("core_identity", "")),

            "Vision_installation_and_physical": _as_json(a.get("installation_and_physical", "")),

            "Vision_usage_context": _as_json(a.get("usage_context", "")),

            "Vision_market_expansion": _as_json(a.get("market_expansion", "")),

            "Vision_compatibility": _as_json(a.get("compatibility", "")),

            "Vision_functional_inference": _as_json(a.get("functional_inference", "")),

            "Vision_search_boost_elements": _as_json(a.get("search_boost_elements", "")),

        }



    def _stamp_vision_to_ocr_results(gs_code9, payload):

        """같은 GS코드의 가장 최근 OCR 결과 레코드에 Vision 컬럼을 부착."""

        if not gs_code9 or not ocr_results_list or not payload:

            return

        for i in range(len(ocr_results_list) - 1, -1, -1):

            row = ocr_results_list[i]

            if str(row.get("GS코드", "")).strip() == str(gs_code9).strip():

                row.update(payload)

                return





    def _apply_market_keyword_packages(

        row_idx,

        product_name: str,

        source_text: str,

        naver_keyword_table: str,

        model_name: str,

        anchors,

        baseline,

    ):

        result = generate_market_keyword_packages(

            product_name=product_name,

            source_text=source_text,

            model_name=model_name,

            anchors=anchors,

            baseline=baseline,

            naver_keyword_table=naver_keyword_table,

        )

        if result.search_keywords:

            df_after.at[row_idx, "검색키워드"] = result.search_keywords

            _status(status_cb, f"검색키워드 생성: {str(product_name)[:20]} → {len(result.search_keywords)}자")

        if result.coupang_tags:

            coupang_line = ",".join(result.coupang_tags)

            df_after.at[row_idx, "검색어설정"] = coupang_line

            df_after.at[row_idx, "쿠팡검색태그"] = coupang_line

        if result.naver_tags:

            df_after.at[row_idx, "네이버태그"] = "|".join(result.naver_tags)

        return result



    # ── OCR 결과 Excel 로드 (미리 처리된 경우) ──

    ocr_lookup: dict = {}

    if cfg.ocr_excel_path and os.path.isfile(cfg.ocr_excel_path):

        from app.services.ocr_excel import read_ocr_results

        ocr_lookup, _ocr_meta = read_ocr_results(cfg.ocr_excel_path)

        _status(status_cb, f"OCR 결과 로드: {len(ocr_lookup)}개 상품 ({os.path.basename(cfg.ocr_excel_path)})")



    global_listing_sources = []

    naver_cache = {}

    no_detail_indices = []  # 상세 없는 상품 인덱스 수집

    ocr_results_list = []  # OCR 결과 수집 (저장용)



    def clean_base_name_for_naver(s: str) -> str:

        if not s:

            return ""

        s = re.sub(r"(GS\d{7}[A-Z0-9]*)", "", s)

        s = re.sub(r"[\[\]\(\)\-_|]+", " ", s)

        s = re.sub(r"\s+", " ", s).strip()

        return s



    CTR_THR = 0.05

    naver_stage_emitted = False

    fatal_gpt_error = ""



    def _is_fatal_gpt_error(msg: str) -> bool:

        m = str(msg or "").lower()

        return (

            "model_not_found" in m

            or "error code: 404" in m

        )



    def _fetch_naver_items_with_retry(hint_keywords: str):

        nonlocal naver_stage_emitted

        naver_err_local = ""

        if not hint_keywords:

            return [], naver_err_local

        if hint_keywords in naver_cache:

            return naver_cache.get(hint_keywords, []), naver_err_local

        try:

            if not naver_stage_emitted:

                _status(status_cb, "네이버 키워드 조회 중")

                _progress(progress_cb, 50)

                naver_stage_emitted = True

            retries = max(0, int(cfg.naver_retry_count)) if cfg.naver_retry else 0

            attempts = retries + 1

            items_local = []

            for attempt in range(attempts):

                try:

                    items_local = core.naver_keyword_tool(naver_keys, hint_keywords, debug=False)

                    naver_cache[hint_keywords] = items_local

                    if not core.DRY_RUN:

                        time.sleep(core.SLEEP_BETWEEN_CALLS)

                    break

                except Exception as e:

                    naver_err_local = str(e)

                    if attempt < attempts - 1:

                        _status(status_cb, f"네이버 재시도 {attempt + 1}/{attempts - 1}")

                        time.sleep(float(cfg.naver_retry_delay))

                    else:

                        naver_cache[hint_keywords] = []

        except Exception as e:

            naver_cache[hint_keywords] = []

            naver_err_local = str(e)

        return naver_cache.get(hint_keywords, []), naver_err_local



    def _query_naver_two_pass(final_line: str, base_name: str):

        return [], "", "네이버 API 비활성화"



    total_rows = max(1, len(df_after))

    for row_i, idx in enumerate(df_after.index, start=1):

        search_keywords = ""  # 검색 키워드 초기화

        try:

            full_pname = str(df_after.at[idx, name_col])

            base_name, option_text = core.extract_base_and_option(full_pname)

            option_tokens = set(core.tokenize_korean_words(option_text))

            prompt_product_name = base_name



            gs_code9 = None

            if code_col and code_col in df_after.columns:

                m1 = re.search(r"(GS\d{7})", str(df_after.at[idx, code_col]) or "")

                gs_code9 = m1.group(1) if m1 else None

            if not gs_code9:

                m2 = re.search(r"(GS\d{7})", full_pname)

                gs_code9 = m2.group(1) if m2 else None



            local_texts, matched_count = [], 0

            local_pairs = []



            # OCR 결과 Excel 에서 미리 처리된 데이터 사용

            # OCR Excel에 GSxxxxxxxA 형태로 저장되어 있으므로 두 가지 키로 매칭

            _ocr_match_key = None

            if ocr_lookup and gs_code9:

                for _mk in [gs_code9, f"{gs_code9}A"]:

                    if _mk in ocr_lookup:

                        _ocr_match_key = _mk

                        break

            # OCR Excel에서 매칭 시도 (raw 텍스트가 있는 경우만 사용)

            _used_ocr_excel = False

            if _ocr_match_key:

                _ocr_data = ocr_lookup[_ocr_match_key]

                if _ocr_data["raw"] and _ocr_data["raw"].strip():

                    local_texts = [_ocr_data["raw"]]

                    matched_count = _ocr_data["count"]

                    _used_ocr_excel = True



                    # OCR 결과 수집 (OCR Excel에서 불러온 경우)

                    ocr_results_list.append({

                        "GS코드": gs_code9,

                        "상품명": base_name,

                        "처리된이미지수": 0,

                        "전체이미지수": matched_count,

                        "OCR텍스트": _ocr_data["raw"][:500],

                        "이미지경로": "OCR Excel에서 로드"

                    })



                # OCR Excel의 이미지 경로로 대표이미지 소스 수집

                _gs_low = gs_code9.lower()

                _img_list = _ocr_data.get("images", [])

                _valid_imgs = [p for p in _img_list if p.strip() and p.strip().lower() != "nan" and os.path.isfile(p) and _gs_low in os.path.basename(p).lower()]

                if _valid_imgs:

                    global_listing_sources.extend(_valid_imgs)



            # phase=analysis: 기존 _ocr_tmp에서 이미지 재사용
            if cfg.phase == "analysis" and gs_code9:
                _gs_dir = os.path.join(ocr_temp_root, f"{gs_code9}A")
                if os.path.isdir(_gs_dir):
                    _existing = sorted([
                        os.path.join(_gs_dir, f) for f in os.listdir(_gs_dir)
                        if f.lower().endswith(('.jpg', '.jpeg', '.png', '.webp'))
                    ])
                    if _existing:
                        global_listing_sources.extend(_existing)
                        if not _used_ocr_excel:
                            all_hits_raw_reuse = _existing
                            ocr_hits_reuse = [p for p in all_hits_raw_reuse if os.path.splitext(os.path.basename(p))[0].isdigit()]
                            if not ocr_hits_reuse:
                                ocr_hits_reuse = all_hits_raw_reuse
                            sel = ocr_hits_reuse[:max_imgs] if max_imgs > 0 else []
                            local_texts, local_pairs = ocr_paths(sel)

            # OCR Excel에서 텍스트를 못 가져온 경우 → 로컬 이미지 OCR 실행

            if cfg.phase != "analysis" and not _used_ocr_excel and use_local and gs_code9:

                all_hits_raw = []



                # OCR용 이미지: 숫자 파일명만 (1.jpg, 2.jpg, 3.jpg...)

                ocr_hits = [

                    p for p in all_hits_raw

                    if os.path.splitext(os.path.basename(p))[0].isdigit()

                ]



                # 숫자파일이 없으면 → O열(상품상세설명)에서 이미지 URL 다운로드 시도

                if not ocr_hits and detail_col and detail_col in df_after.columns:

                    _detail_html_for_dl = str(df_after.at[idx, detail_col]) if pd.notna(df_after.at[idx, detail_col]) else ""

                    if _detail_html_for_dl and "<img" in _detail_html_for_dl.lower():

                        _status(status_cb, f"[{row_i}/{total_rows}] {gs_code9} — O열에서 상세이미지 다운로드 중...")

                        from app.services.ocr_pipeline import _download_and_save_images

                        _dl_paths = _download_and_save_images(_detail_html_for_dl, f"{gs_code9}A", ocr_temp_root)

                        if _dl_paths:

                            ocr_hits = _dl_paths

                            all_hits_raw = all_hits_raw + _dl_paths

                            _status(status_cb, f"[{row_i}/{total_rows}] {gs_code9} — 상세이미지 {len(_dl_paths)}개 다운로드 완료")

                        else:

                            _status(status_cb, f"[{row_i}/{total_rows}] {gs_code9} — O열 이미지 다운로드 실패")



                # 대표이미지용: GS코드가 파일명에 포함된 것만

                _gs_low2 = gs_code9.lower()

                listing_hits = [p for p in all_hits_raw if _gs_low2 in os.path.basename(p).lower()]

                global_listing_sources.extend(listing_hits)



                matched_count = len(all_hits_raw)

                if len(ocr_hits) > 0:

                    sel = ocr_hits[:max_imgs] if max_imgs > 0 else []

                    local_texts, local_pairs = ocr_paths(sel)



                    # OCR 결과 수집 (저장용) - 항상 수집

                    ocr_text_combined = " ".join(local_texts) if local_texts else "(OCR 텍스트 없음)"

                    ocr_results_list.append({

                        "GS코드": gs_code9,

                        "상품명": base_name,

                        "OCR처리이미지수": len(sel),

                        "전체이미지수": matched_count,

                        "대표이미지수": len(listing_hits),

                        "OCR텍스트": ocr_text_combined[:500],  # 500자로 제한

                        "OCR이미지": "; ".join([os.path.basename(p) for p in sel[:5]])  # 처음 5개만

                    })

                elif matched_count > 0:

                    # OCR용 이미지는 없지만 대표이미지는 있는 경우

                    ocr_results_list.append({

                        "GS코드": gs_code9,

                        "상품명": base_name,

                        "OCR처리이미지수": 0,

                        "전체이미지수": matched_count,

                        "대표이미지수": len(listing_hits),

                        "OCR텍스트": "(숫자 파일명 없음 - OCR 불가)",

                        "OCR이미지": "없음"

                    })



            # AU열 (이미지등록(상세)) 순차 다운로드 - 대표이미지용

            # 로컬에 이미지가 없고 AU열에 URL이 있으면 순차 다운로드 (phase=analysis에서는 건너뜀)

            if cfg.phase != "analysis" and listing_img_col and listing_img_col in df_after.columns and gs_code9:

                listing_url = str(df_after.at[idx, listing_img_col]) if pd.notna(df_after.at[idx, listing_img_col]) else ""

                if listing_url and listing_url.startswith("http"):

                    # 로컬에 GS코드 이미지가 없을 때만 다운로드

                    _gs_low_check = gs_code9.lower()

                    _existing_gs_imgs = [p for p in global_listing_sources if _gs_low_check in os.path.basename(p).lower()]

                    if not _existing_gs_imgs:

                        from app.services.ocr_pipeline import _download_sequential_images

                        downloaded = _download_sequential_images(

                            base_url=listing_url,

                            gs_code=f"{gs_code9}A",  # GS코드A 폴더에 저장

                            target_dir=ocr_temp_root,

                            max_fails=3,

                            max_images=100

                        )

                        if downloaded:

                            global_listing_sources.extend(downloaded)

                            _status(status_cb, f"[AU다운] {gs_code9}: {len(downloaded)}개 이미지 다운로드")


            # phase=images → 이미지 수집만 하고 OCR/Vision/키워드 건너뜀
            if cfg.phase == "images":
                continue

            detail_html = str(df_after.at[idx, detail_col]) if (detail_col and detail_col in df_after.columns) else ""

            html_text = re.sub(r"\s+", " ", core.extract_text_from_html(detail_html)) if detail_html else ""

            if korean_only and html_text:

                html_text = re.sub(r"[^가-힣\s]", " ", html_text)



            url_texts = []

            if cfg.merge_ocr_with_name and cfg.use_html_ocr and detail_html:

                srcs = core.extract_img_srcs(detail_html, max_images=max_imgs)

                for url in srcs:

                    ocr_txt = core.ocr_image_url(url, tess_lang, timeout=10, psm=psm, oem=oem, korean_only=korean_only)

                    if ocr_txt:

                        url_texts.append(ocr_txt)



            # OCR Excel 에서 raw 텍스트를 최우선으로 GPT에 전달

            ocr_raw_text = ""

            sum_text = ""

            _ocr_key = None

            if ocr_lookup and gs_code9:

                # GSxxxxxxxA 형태로도 매칭 시도

                for _try_key in [gs_code9, f"{gs_code9}A"]:

                    if _try_key in ocr_lookup:

                        _ocr_key = _try_key

                        break



            # OCR Excel이 로드된 상태에서 해당 상품의 OCR 데이터가 없으면 → 상세 없음 → 스킵

            if ocr_lookup and gs_code9:

                _has_ocr = False

                if _ocr_key:

                    _raw_check = ocr_lookup[_ocr_key].get("raw", "").strip()

                    _has_ocr = bool(_raw_check) and _raw_check != "(옵션상품 — A옵션 결과 참조)"

                if not _has_ocr:

                    no_detail_indices.append(idx)

                    _status(status_cb, f"[{row_i}/{total_rows}] {gs_code9} — 상세 없음, 스킵")

                    continue



            if _ocr_key and cfg.merge_ocr_with_name:

                ocr_raw_text = ocr_lookup[_ocr_key].get("raw", "")

                # 반복 문구 필터링 후 raw 텍스트를 최대 500자까지 GPT에 직접 전달

                if ocr_raw_text:

                    from app.services.ocr_noise_filter import filter_ocr_text, preprocess_ocr_for_llm

                    _raw0 = ocr_raw_text

                    _f = filter_ocr_text(_raw0)

                    _p = preprocess_ocr_for_llm(_f)

                    ocr_raw_text = _p if _p else (_f if _f else _raw0)

                sum_text = ocr_raw_text[:500] if ocr_raw_text else ""

            elif cfg.merge_ocr_with_name:

                detail_texts = []

                if local_texts:

                    detail_texts.extend(local_texts)

                if html_text:

                    detail_texts.append(html_text)

                if url_texts:

                    detail_texts.extend(url_texts)

                _combined_ocr = " ".join(detail_texts)

                # 반복 문구 필터링 (OCR Excel 경로와 동일하게 적용)

                if _combined_ocr:

                    try:

                        from app.services.ocr_noise_filter import filter_ocr_text, preprocess_ocr_for_llm

                        _raw0 = _combined_ocr

                        _f = filter_ocr_text(_raw0)

                        _p = preprocess_ocr_for_llm(_f)

                        _combined_ocr = _p if _p else (_f if _f else _raw0)

                    except Exception:

                        pass

                # OCR 원문을 500자까지 직접 GPT에 전달 (summarize_features_tokens 미사용)

                _combined_ocr = re.sub(r"\s+", " ", _combined_ocr).strip()

                sum_text = _combined_ocr[:500] if _combined_ocr else ""

            else:

                sum_text = ""





            # 디버그: GPT에 전달되는 OCR 텍스트 확인

            if sum_text:

                _status(status_cb, f"[{row_i}/{total_rows}] {gs_code9} — GPT에 OCR {len(sum_text)}자 전달")

            else:

                _status(status_cb, f"[{row_i}/{total_rows}] {gs_code9} — ⚠ GPT에 OCR 텍스트 없음! (local_texts={len(local_texts)}, html_text={len(html_text) if html_text else 0})")



            gpt_model_kw = cfg.model_keyword

            gpt_model_lt = cfg.model_longtail

            stage2_model = cfg.model_keyword_stage2 or gpt_model_kw



            # 대표이미지 Vision 분석 복원: category/구조/설치/재질/색상/유형교정을 추정

            _vision_analysis = {}

            _vision_hint_parts = []

            if gs_code9:

                _gs_low_v = gs_code9.lower()

                _vision_imgs = [p for p in global_listing_sources if _gs_low_v in os.path.basename(p).lower() and os.path.isfile(p)][:5]

                if _vision_imgs:

                    _status(status_cb, f"[{row_i}/{total_rows}] {gs_code9} — Vision 분석 중 ({len(_vision_imgs)}장)...")

                    _vision_analysis = analyze_product_images_local(_vision_imgs, prompt_product_name, gpt_model_kw)

                    if _vision_analysis:

                        _vision_hint_parts, _filled_paths = _extract_vision_hints(_vision_analysis)

                        _status(

                            status_cb,

                            f"[{row_i}/{total_rows}] {gs_code9} — Vision 채움 항목: {len(_filled_paths)}개"

                            + (f" ({', '.join(_filled_paths[:8])})" if _filled_paths else "")

                        )



            _sum_text_with_vision = sum_text

            if _vision_hint_parts:

                _vision_hint_text = " ".join(dict.fromkeys(_vision_hint_parts))[:300]

                _sum_text_with_vision = (f"{sum_text} {_vision_hint_text}").strip()



            _vision_payload = _vision_excel_payload(_vision_analysis, _vision_hint_parts)

            for _k, _v in _vision_payload.items():

                df_after.at[idx, _k] = _v

            _stamp_vision_to_ocr_results(gs_code9, _vision_payload)



            # Vision JSON이 실제로 채워졌을 때만 로컬 키워드 빌더 사용.

            # Vision이 비어 있으면 OCR 안내문이 그대로 토큰화될 수 있어 1차 GPT 경로를 우선한다.

            kw_line = ""

            kw_tokens = []

            gpt_err = ""

            # GPT 프롬프트(최적화됨)를 우선 사용. Vision 데이터는 context로 전달.
            kw_line, kw_tokens = core.generate_keyword_gpt(
                prompt_product_name, _sum_text_with_vision, gpt_model_kw, max_words, max_len, min_len,
                vision_analysis=_vision_analysis,
            )

            gpt_err = getattr(core, "LAST_GPT_ERROR", "")
            if gpt_err:
                _status(status_cb, f"GPT 오류: {gpt_err}")
                if _is_fatal_gpt_error(gpt_err):
                    fatal_gpt_error = gpt_err
                    _status(status_cb, "치명적 GPT 오류(404) 감지: 저장 없이 작업 중단")
                    break

            # GPT 실패 시 keyword_builder → heuristic fallback
            if not kw_line or len(kw_line) < 90:
                if _vision_analysis:
                    try:
                        kw_line = build_keyword_string(
                            ocr_text=_sum_text_with_vision,
                            vision_analysis=_vision_analysis,
                            target_count=20,
                            fallback_text=prompt_product_name,
                        )
                        if kw_line:
                            kw_tokens = [t for t in re.split(r"\s+", kw_line) if t]
                    except Exception:
                        kw_line = ""
                        kw_tokens = []

            if not kw_line:
                kw_line = core._fallback_heuristic(prompt_product_name, _sum_text_with_vision, max_n=max_words)
                kw_tokens = [t for t in re.split(r"\s+", kw_line) if t]



            search_keywords = ""

            kw_tokens = [t for t in kw_tokens if t not in core.SIZE_WORDS and t not in core.STOPWORDS]

            if drop_digits:

                kw_tokens = core._filter_tokens_drop_digits(kw_tokens)



            # 네이버 쇼핑 자동완성 키워드 조합

            ac_keywords_debug = []

            if cfg.naver_autocomplete:

                try:

                    ac_raw = core.get_autocomplete_keywords_for_product(base_name, max_queries=2, max_results=10)

                    if ac_raw:

                        anchors_for_ac = core.build_anchors_from_name(base_name)

                        baseline_for_ac = core.build_baseline_tokens_from_name(base_name)

                        ac_cleaned = core.clean_naver_kw_list(ac_raw, anchors=anchors_for_ac, baseline=baseline_for_ac)

                        ac_keywords_debug = list(ac_cleaned)

                        # 자동완성 키워드를 토큰화하여 기존 kw_tokens에 없는 것만 추가

                        existing = set(kw_tokens)

                        for ac_kw in ac_cleaned:

                            ac_toks = core.tokenize_korean_words(ac_kw)

                            for t in ac_toks:

                                if t not in existing and t not in core.STOPWORDS and t not in core.SIZE_WORDS:

                                    if len(t) >= 2:

                                        existing.add(t)

                                        kw_tokens.append(t)

                        _status(status_cb, f"네이버자동완성: {base_name[:20]} → {len(ac_cleaned)}개")

                except Exception as e:

                    _status(status_cb, f"네이버자동완성 오류: {e}")



            # 구글 자동완성 키워드 조합

            google_ac_debug = []

            if cfg.google_autocomplete:

                try:

                    gac_raw = core.get_google_autocomplete_for_product(base_name, max_queries=2, max_results=10)

                    if gac_raw:

                        anchors_for_gac = core.build_anchors_from_name(base_name)

                        baseline_for_gac = core.build_baseline_tokens_from_name(base_name)

                        gac_cleaned = core.clean_naver_kw_list(gac_raw, anchors=anchors_for_gac, baseline=baseline_for_gac)

                        google_ac_debug = list(gac_cleaned)

                        existing = set(kw_tokens)

                        for gac_kw in gac_cleaned:

                            gac_toks = core.tokenize_korean_words(gac_kw)

                            for t in gac_toks:

                                if t not in existing and t not in core.STOPWORDS and t not in core.SIZE_WORDS:

                                    if len(t) >= 2:

                                        existing.add(t)

                                        kw_tokens.append(t)

                        _status(status_cb, f"구글자동완성: {base_name[:20]} → {len(gac_cleaned)}개")

                except Exception as e:

                    _status(status_cb, f"구글자동완성 오류: {e}")



            kw_line = " ".join(kw_tokens)



            final_line = core.merge_base_name_with_keywords(base_name, kw_line, max_words, max_len, option_tokens=option_tokens, ocr_text=sum_text)



            # min_len 미달 시 OCR 텍스트에서 직접 보충

            if len(final_line) < min_len and sum_text:

                _ocr_sup, _ocr_toks = core.postprocess_keywords_tokens(sum_text, max_words=max_words, max_len=max_len)

                _existing = set(final_line.split())

                _final_tokens = final_line.split()

                for _t in _ocr_toks:

                    if _t not in _existing and len(_t) >= 2 and _t not in core.STOPWORDS and _t not in core.SIZE_WORDS:

                        _existing.add(_t)

                        _final_tokens.append(_t)

                    if len(" ".join(_final_tokens)) >= min_len or len(_final_tokens) >= max_words:

                        break

                final_line = " ".join(_final_tokens)[:max_len].rstrip()



            df_after.at[idx, name_col] = final_line

            df_after.at[idx, "1차키워드"] = final_line

            df_after.at[idx, "최종키워드2차"] = final_line

            df_after.at[idx, "OCR요약"] = (sum_text or "")[:500]

            df_after.at[idx, "네이버검색광고데이터"] = ""

            df_after.at[idx, "검색키워드"] = search_keywords  # 검색 키워드 추가



            if bool(cfg.write_to_r):

                TARGET_N = 20



                anchors = core.build_anchors_from_name(final_line)

                baseline = core.build_baseline_tokens_from_name(final_line)



                lt10_raw = core.generate_longtail10(final_line, sum_text, client=core.client, model_name=gpt_model_lt)

                lt10 = core.clean_naver_kw_list(lt10_raw, anchors=anchors, baseline=baseline)

                if len(lt10) < 10:

                    name_parts = [p for p in re.sub(r"[^0-9가-힣\sA-Za-z]", " ", final_line).split() if len(p) >= 2]

                    bigrams = [name_parts[i] + name_parts[i + 1] for i in range(len(name_parts) - 1)]

                    lt10_backup = core.clean_naver_kw_list(bigrams, anchors=anchors, baseline=baseline)

                    for k in lt10_backup:

                        if k not in lt10:

                            lt10.append(k)

                        if len(lt10) >= 10:

                            break

                lt10 = lt10[:10]



                naver_pc5, naver_mo5 = [], []

                items, hint_used, naver_err = _query_naver_two_pass(final_line, base_name)



                if items:

                    pc_list = core.rank_and_pick_with_ctr(items, platform="pc", want=5, ctr_threshold=CTR_THR)

                    mo_list = core.rank_and_pick_with_ctr(items, platform="mobile", want=5, ctr_threshold=CTR_THR)

                    pc_list = core.clean_naver_kw_list(pc_list, anchors=anchors, baseline=baseline)

                    mo_list = core.clean_naver_kw_list(mo_list, anchors=anchors, baseline=baseline)



                    inter = set(pc_list) & set(mo_list)

                    if inter:

                        pc_list = [k for k in pc_list if k not in inter]



                    if len(pc_list) < 5:

                        backup_pc = core.rank_and_pick_with_ctr(items, platform="pc", want=15, ctr_threshold=0.0)

                        backup_pc = core.clean_naver_kw_list(backup_pc, anchors=anchors, baseline=baseline)

                        for k in backup_pc:

                            if k not in pc_list and k not in mo_list:

                                pc_list.append(k)

                            if len(pc_list) >= 5:

                                break

                    if len(mo_list) < 5:

                        backup_mo = core.rank_and_pick_with_ctr(items, platform="mobile", want=15, ctr_threshold=0.0)

                        backup_mo = core.clean_naver_kw_list(backup_mo, anchors=anchors, baseline=baseline)

                        for k in backup_mo:

                            if k not in mo_list and k not in pc_list:

                                mo_list.append(k)

                            if len(mo_list) >= 5:

                                break



                    naver_pc5 = pc_list[:5]

                    naver_mo5 = mo_list[:5]



                final_kw, seen = [], set()



                def _clean_search_terms(lst):

                    cleaned = []

                    seen_local = set()

                    for w in (lst or []):

                        k = re.sub(r"\s+", "", str(w or ""))

                        if not k:

                            continue

                        # 조사/어미/문장형 파편 제거

                        if re.search(r"(하다|하는|되어|됨|하고|하기)$", k):

                            continue

                        if re.search(r"(에|에서|으로|로|을|를|이|가|은|는|의|와|과)$", k):

                            continue

                        if re.search(r"[가-힣](에|에서|으로|로|을|를|이|가|은|는|의|와|과)[가-힣]{2,}", k):

                            continue

                        if any(x in k for x in ["설치하고", "차량에", "조명을"]):

                            continue

                        # 판매처/브랜드 워터마크 파편 제거 (예: 홈런마켓브라켓)

                        if any(x in k for x in ["마켓", "스토어", "쇼핑몰", "샵", "몰"]):

                            continue

                        # 비자연 결합어 제거 (의미 약한 합성 파편)

                        if any(x in k for x in ["금속플라스틱", "플라스틱검정", "검정은색", "은색실외", "볼트외부", "외부금속"]):

                            continue

                        # 평가/형용 파편 결합어 제거

                        if re.search(r"(견고한|간편한|강력한|튼튼한|편리한)$", k):

                            continue

                        if re.search(r"(견고한|간편한|강력한|튼튼한|편리한)", k) and len(k) >= 6:

                            if not any(x in k for x in ["브라켓", "마운트", "거치대", "고정", "볼트", "조명"]):

                                continue

                        if "간편설치" in k and len(k) >= 6:

                            continue

                        if "외부간편" in k:

                            continue

                        if re.search(r"설치[가-힣]{2,}", k):

                            continue

                        if any(x in k for x in ["간편설치견고한", "견고한간편한", "간편한금속"]):

                            continue

                        # 재질/색상 단어 2개 이상을 붙여쓴 기계적 결합어 제거

                        mat_color = ["금속", "플라스틱", "스틸", "알루미늄", "철제", "검정", "은색", "실버", "블랙"]

                        hit = sum(1 for t in mat_color if t in k)

                        if hit >= 2 and len(k) >= 6:

                            continue

                        lk = k.lower()

                        if lk in seen_local:

                            continue

                        seen_local.add(lk)

                        cleaned.append(k)

                    return cleaned



                def push(lst):

                    for w in lst:

                        w = core._clean_one_kw(w)

                        if not w:

                            continue

                        if any(b in w for b in core.BAN):

                            continue

                        if not (2 <= len(w) <= 12):

                            continue

                        if not core.is_on_topic(w, anchors, baseline):

                            continue

                        if w in seen:

                            continue

                        seen.add(w)

                        final_kw.append(w)

                        if len(final_kw) >= TARGET_N:

                            break



                push(naver_pc5)

                push(naver_mo5)

                push(lt10)



                if len(final_kw) < TARGET_N and items:

                    backup_mix = core.rank_and_pick_with_ctr(items, platform="pc", want=30, ctr_threshold=0.0)

                    backup_mix += core.rank_and_pick_with_ctr(items, platform="mobile", want=30, ctr_threshold=0.0)

                    backup_mix = core.clean_naver_kw_list(backup_mix, anchors=anchors, baseline=baseline)

                    push(backup_mix)



                if len(final_kw) < TARGET_N:

                    name_parts = [p for p in re.sub(r"[^0-9가-힣\sA-Za-z]", " ", final_line).split() if len(p) >= 2]

                    extra = []

                    for i in range(len(name_parts) - 1):

                        extra.append(name_parts[i] + name_parts[i + 1])

                    for i in range(len(name_parts) - 2):

                        extra.append(name_parts[i] + name_parts[i + 1] + name_parts[i + 2])

                    extra = core.clean_naver_kw_list(extra, anchors=anchors, baseline=baseline)

                    push(extra)



                if len(final_kw) < TARGET_N and core.USE_GPT_BACKFILL:

                    fallback = core._fallback_heuristic(final_line, sum_text, max_n=TARGET_N)

                    push([x for x in fallback.split(",") if x])



                final_kw = _clean_search_terms(final_kw)

                df_after.at[idx, "검색어설정"] = ",".join(final_kw[:TARGET_N])

                naver_info = (

                    f"hint={hint_used} | "

                    f"PC5={','.join(naver_pc5)} | "

                    f"MO5={','.join(naver_mo5)}"

                )

                if naver_err:

                    naver_info += f" | 오류={naver_err[:120]}"

                naver_table = _format_naver_keyword_table(items, limit=15)

                df_after.at[idx, "네이버검색광고데이터"] = naver_table if naver_table else naver_info



                stage2_source = naver_table if naver_table else naver_info

                stage2_kw, _stage2_tokens = core.generate_keyword_stage2(

                    seed_keywords=final_line,

                    naver_keyword_table=stage2_source,

                    ocr_text=sum_text,

                    model_name=stage2_model,

                    min_len=50,

                    max_len=min(max_len, 90),

                    max_words=max_words,

                )

                stage2_err = getattr(core, "LAST_GPT_ERROR", "")

                if stage2_err and _is_fatal_gpt_error(stage2_err):

                    fatal_gpt_error = stage2_err

                    _status(status_cb, f"치명적 GPT 오류(2차): {stage2_err}")

                    break

                if stage2_kw:

                    _st2_ok = True

                    _st2_toks = set(core.tokenize_korean_words(stage2_kw))

                    _base_toks = set(core.tokenize_korean_words(final_line))

                    if len(_st2_toks) < 2:

                        _st2_ok = False

                    if _base_toks and len(_st2_toks & _base_toks) == 0:

                        _st2_ok = False

                    if not core.is_on_topic(stage2_kw, anchors, baseline):

                        _st2_ok = False

                    if _st2_ok:

                        df_after.at[idx, "최종키워드2차"] = stage2_kw

                        df_after.at[idx, name_col] = stage2_kw

                    else:

                        stage2_kw = ""

                final_for_search = stage2_kw if stage2_kw else final_line

                market_pkg = _apply_market_keyword_packages(

                    row_idx=idx,

                    product_name=final_for_search,

                    source_text=_sum_text_with_vision,

                    naver_keyword_table=stage2_source,

                    model_name=gpt_model_lt,

                    anchors=anchors,

                    baseline=baseline,

                )

                search_keywords = market_pkg.search_keywords or search_keywords



                if debug_on:

                    debug_rows.append({

                        "R열_타깃개수": TARGET_N,

                        "R열_최종개수": len(final_kw),

                        "네이버_DRY": core.DRY_RUN,

                        "네이버_사용": "Y" if cfg.naver_enabled else "N",

                        "네이버_오류": naver_err,

                        "네이버_hintKeywords": hint_used,

                        "네이버_PC5": ",".join(naver_pc5),

                        "네이버_MO5": ",".join(naver_mo5),

                        "롱테일10": ",".join(lt10),

                        "앵커": ",".join(sorted(anchors)),

                        "베이스라인": ",".join(sorted(baseline)),

                        "검색어설정(R)": df_after.at[idx, "검색어설정"],

                    })

            else:

                anchors = core.build_anchors_from_name(final_line)

                baseline = core.build_baseline_tokens_from_name(final_line)



                naver_pc5, naver_mo5 = [], []

                items, hint_used, naver_err = _query_naver_two_pass(final_line, base_name)



                if items:

                    pc_list = core.rank_and_pick_with_ctr(items, platform="pc", want=5, ctr_threshold=CTR_THR)

                    mo_list = core.rank_and_pick_with_ctr(items, platform="mobile", want=5, ctr_threshold=CTR_THR)

                    naver_pc5 = core.clean_naver_kw_list(pc_list, anchors=anchors, baseline=baseline)[:5]

                    naver_mo5 = core.clean_naver_kw_list(mo_list, anchors=anchors, baseline=baseline)[:5]



                naver_info = (

                    f"hint={hint_used} | "

                    f"PC5={','.join(naver_pc5)} | "

                    f"MO5={','.join(naver_mo5)}"

                )

                if naver_err:

                    naver_info += f" | 오류={naver_err[:120]}"

                naver_table = _format_naver_keyword_table(items, limit=15)

                df_after.at[idx, "네이버검색광고데이터"] = naver_table if naver_table else naver_info



                stage2_source = naver_table if naver_table else naver_info

                stage2_kw, _stage2_tokens = core.generate_keyword_stage2(

                    seed_keywords=final_line,

                    naver_keyword_table=stage2_source,

                    ocr_text=sum_text,

                    model_name=stage2_model,

                    min_len=50,

                    max_len=min(max_len, 90),

                    max_words=max_words,

                )

                stage2_err = getattr(core, "LAST_GPT_ERROR", "")

                if stage2_err and _is_fatal_gpt_error(stage2_err):

                    fatal_gpt_error = stage2_err

                    _status(status_cb, f"치명적 GPT 오류(2차): {stage2_err}")

                    break

                if stage2_kw:

                    _st2_ok = True

                    _st2_toks = set(core.tokenize_korean_words(stage2_kw))

                    _base_toks = set(core.tokenize_korean_words(final_line))

                    if len(_st2_toks) < 2:

                        _st2_ok = False

                    if _base_toks and len(_st2_toks & _base_toks) == 0:

                        _st2_ok = False

                    if not core.is_on_topic(stage2_kw, anchors, baseline):

                        _st2_ok = False

                    if _st2_ok:

                        df_after.at[idx, "최종키워드2차"] = stage2_kw

                        df_after.at[idx, name_col] = stage2_kw

                    else:

                        stage2_kw = ""

                final_for_search = stage2_kw if stage2_kw else final_line

                market_pkg = _apply_market_keyword_packages(

                    row_idx=idx,

                    product_name=final_for_search,

                    source_text=_sum_text_with_vision,

                    naver_keyword_table=stage2_source,

                    model_name=gpt_model_lt,

                    anchors=anchors,

                    baseline=baseline,

                )

                search_keywords = market_pkg.search_keywords or search_keywords



            if debug_on:

                ocr_samples = "; ".join([f"{fn}:{snip}" for fn, snip in (local_pairs[:5] if local_pairs else [])])

                debug_rows.append({

                    "상품명(원본)": full_pname,

                    "기본상품명(옵션제외)": base_name,

                    "옵션(본문)": option_text,

                    "옵션_토큰": " ".join(sorted(option_tokens)),

                    "적용 GS코드": gs_code9 or "",

                    "로컬_매칭수": matched_count,

                    "OCR샘플": ocr_samples,

                    "HTML텍스트": html_text[:200] if html_text else "",

                        "상세요약": sum_text[:220],

                        "GPT_ERROR": gpt_err,

                        "키워드정렬": kw_line,

                        "검색키워드": search_keywords,

                        "네이버자동완성": ",".join(ac_keywords_debug) if ac_keywords_debug else "",

                        "구글자동완성": ",".join(google_ac_debug) if google_ac_debug else "",

                        "최종상품명": final_line,

                    })

        except Exception:

            if debug_on:

                debug_rows.append({"오류행": int(idx), "오류": "행 처리 중 예외"})

            continue

        if row_i % 5 == 0 or row_i == total_rows:

            pct = 10 + int((row_i / total_rows) * 55)

            _progress(progress_cb, min(65, pct))



    if fatal_gpt_error:

        raise RuntimeError(f"치명적 GPT 오류로 저장 중단: {fatal_gpt_error}")



    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    save_path = os.path.join(export_root, f"상품전처리GPT_{timestamp}.xlsx")

    def _seq_path(directory: str, prefix: str, ext: str) -> str:

        """순차 번호 파일명 생성: prefix_01.ext, prefix_02.ext, ..."""

        seq = 1

        while True:

            name = f"{prefix}_{seq:02d}{ext}"

            path = os.path.join(directory, name)

            if not os.path.exists(path):

                return path

            seq += 1



    # 상세 없는 상품은 업로드용에서 제외

    df_upload = df_after.loc[~df_after.index.isin(no_detail_indices)].copy() if no_detail_indices else df_after

    upload_cols = [c for c in input_cols if c in df_upload.columns]

    extra_output_cols = [

        c for c in ["1차키워드", "최종키워드2차", "OCR요약", "검색키워드", "검색어설정", "쿠팡검색태그", "네이버태그", "네이버검색광고데이터", "옵션추가금"]

        if c in df_upload.columns and c not in upload_cols

    ]

    export_cols = upload_cols + extra_output_cols

    df_upload_export = df_upload.loc[:, export_cols].copy() if export_cols else df_upload.copy()



    upload_path = _seq_path(export_root, f"업로드용_{date_tag}", ".xlsx")

    with pd.ExcelWriter(save_path, engine="openpyxl") as writer:

        df.to_excel(writer, sheet_name="분리추출전", index=False)

        df_upload_export.to_excel(writer, sheet_name="분리추출후", index=False)

        if debug_on and len(debug_rows) > 0:

            pd.DataFrame(debug_rows).to_excel(writer, sheet_name="디버그", index=False)

    try:

        with pd.ExcelWriter(upload_path, engine="openpyxl") as writer:

            df_upload_export.to_excel(writer, sheet_name="분리추출후", index=False)

    except PermissionError:

        # file might be opened by Excel; write a timestamped file instead

        upload_path = _safe_path(upload_path)

        with pd.ExcelWriter(upload_path, engine="openpyxl") as writer:

            df_upload_export.to_excel(writer, sheet_name="분리추출후", index=False)



    # ── OCR 결과 별도 파일 저장 ──

    if ocr_results_list:

        # legacy-compatible fixed name (overwrite): OCR결과_YYYYMMDD_02.xlsx

        ocr_result_path = os.path.join(export_root, f"OCR결과_{date_tag}_02.xlsx")

        try:

            df_ocr_results = pd.DataFrame(ocr_results_list)

            with pd.ExcelWriter(ocr_result_path, engine="openpyxl") as writer:

                df_ocr_results.to_excel(writer, sheet_name="OCR결과", index=False)

            _status(status_cb, f"OCR 결과 {len(ocr_results_list)}개 → {os.path.basename(ocr_result_path)}")

        except PermissionError:

            # Excel opened/locked: fallback to sequential name

            ocr_result_path = _seq_path(export_root, f"OCR결과_{date_tag}", ".xlsx")

            df_ocr_results = pd.DataFrame(ocr_results_list)

            with pd.ExcelWriter(ocr_result_path, engine="openpyxl") as writer:

                df_ocr_results.to_excel(writer, sheet_name="OCR결과", index=False)

            _status(status_cb, f"OCR 결과(대체 저장) {len(ocr_results_list)}개 → {os.path.basename(ocr_result_path)}")

        except Exception as e:

            _status(status_cb, f"OCR 결과 저장 오류: {e}")



    # ── 상세 없는 상품 별도 파일 저장 (원본 형식 유지) ──

    if no_detail_indices:

        df_no_detail = df.loc[df.index.isin(no_detail_indices)].copy()

        if not df_no_detail.empty:

            no_detail_path = _seq_path(export_root, f"상세없음_{date_tag}", ".xlsx")

            try:

                with pd.ExcelWriter(no_detail_path, engine="openpyxl") as writer:

                    df_no_detail.to_excel(writer, sheet_name="상세없음", index=False)

                _status(status_cb, f"상세 없는 상품 {len(df_no_detail)}개 → {os.path.basename(no_detail_path)}")

            except Exception as e:

                _status(status_cb, f"상세없음 파일 저장 오류: {e}")



    # phase=analysis 이면 리스팅 이미지 처리 건너뜀 (Phase1에서 이미 완료)
    if cfg.phase == "analysis":
        _status(status_cb, "리스팅 이미지 처리 건너뜀 (Phase1에서 완료)")

    elif cfg.make_listing and len(global_listing_sources) > 0:

        _status(status_cb, "처리중... (2/2) 대표이미지 생성")

        _progress(progress_cb, 70)

        listing_out_root = os.path.join(export_root, "listing_images", date_tag)

        os.makedirs(listing_out_root, exist_ok=True)

        total_imgs = len(set(global_listing_sources))

        processed = 0

        def _progress_cb():

            nonlocal processed

            processed += 1

            pct = 70 + int((processed / max(1, total_imgs)) * 25)

            _progress(progress_cb, min(95, pct))

        core.process_listing_images_global(

            src_paths=list(set(global_listing_sources)),

            base_out_root=listing_out_root,

            logo_rgba=logo_rgba,

            size=listing_size,

            pad=listing_pad,

            bg_color=(255, 255, 255),

            pos=logo_pos,

            opacity=logo_opacity,

            logo_ratio=logo_ratio,

            use_auto_contrast=use_auto_contrast,

            use_sharpen=use_sharpen,

            use_small_rotate=use_small_rotate,

            rotate_zoom=rotate_zoom,

            max_images_per_code=listing_max,

            ultra_angle_deg=ultra_angle_deg,

            ultra_translate_px=ultra_translate_px,

            ultra_scale_pct=ultra_scale_pct,

            trim_tol=trim_tol,

            jpeg_q_min=jpeg_q_min,

            jpeg_q_max=jpeg_q_max,

            do_flip_lr=do_flip_lr,

            progress_cb=_progress_cb,

        )



    # phase=images 이면 이미지 처리만 하고 바로 반환 (OCR/키워드 건너뜀)
    if cfg.phase == "images":
        _status(status_cb, "이미지 처리 완료 (Phase 1)")
        _progress(progress_cb, 100)
        return export_root, ""

    _status(status_cb, "처리 완료")

    _progress(progress_cb, 100)

    return export_root, save_path





# ── 대표이미지만 생성하는 독립 함수 ──────────────────────────────────



@dataclass

class ListingOnlyConfig:

    """대표이미지만 생성할 때 필요한 설정."""

    local_img_dir: str = ""

    allow_folder_match: bool = True

    max_depth: int = -1



    listing_size: int = 1200

    listing_pad: int = 20

    listing_max: int = 20



    logo_path: str = ""

    logo_ratio: int = 14

    logo_opacity: int = 65

    logo_pos: str = "tr"



    use_auto_contrast: bool = True

    use_sharpen: bool = True

    use_small_rotate: bool = True

    rotate_zoom: float = 1.04



    ultra_angle_deg: float = 0.35

    ultra_translate_px: float = 0.6

    ultra_scale_pct: float = 0.25



    trim_tol: int = 8

    jpeg_q_min: int = 88

    jpeg_q_max: int = 92

    do_flip_lr: bool = True



    # 소스: CSV 또는 OCR Excel

    file_path: str = ""          # CSV/Excel (GS코드 목록)

    ocr_excel_path: str = ""     # OCR 결과 Excel (이미지 경로 포함)





def run_listing_only(cfg: ListingOnlyConfig, status_cb=None, progress_cb=None) -> str:

    """대표이미지만 생성 — GPT/키워드 없이 이미지 변환만 수행."""



    if not cfg.file_path:

        raise ValueError("CSV/Excel 파일을 선택해 주세요.")

    if not cfg.local_img_dir and not cfg.ocr_excel_path:

        raise ValueError("이미지 폴더 또는 OCR 결과 Excel이 필요합니다.")



    _status(status_cb, "대표이미지 생성 준비...")

    _progress(progress_cb, 5)



    # CSV에서 GS코드 목록 추출

    df = core.safe_read_csv(cfg.file_path)

    if df.empty:

        raise ValueError("CSV/Excel 내용이 비어 있습니다.")



    code_col = None

    for c in df.columns:

        if "코드" in str(c) or "code" in str(c).lower():

            code_col = c

            break



    name_col = None

    for c in df.columns:

        if "상품명" in str(c) or "name" in str(c).lower():

            name_col = c

            break



    # OCR Excel 로드 (이미지 경로)

    ocr_lookup: dict = {}

    if cfg.ocr_excel_path and os.path.isfile(cfg.ocr_excel_path):

        from app.services.ocr_excel import read_ocr_results

        ocr_lookup, _ = read_ocr_results(cfg.ocr_excel_path)

        _status(status_cb, f"OCR 결과 로드: {len(ocr_lookup)}개 ({os.path.basename(cfg.ocr_excel_path)})")



    local_root = cfg.local_img_dir

    allow_folder_match = cfg.allow_folder_match

    max_depth = cfg.max_depth



    # 이미지 소스 수집 — CSV의 모든 행에서 GS코드 추출 → 이미지 경로 수집

    _status(status_cb, f"이미지 소스 수집중... (CSV {len(df)}행, 코드컬럼={code_col}, 이미지폴더={local_root})")

    _progress(progress_cb, 10)

    global_listing_sources = []

    total_rows = len(df)

    found_codes = set()

    skipped_no_code = 0

    skipped_no_img = 0



    for row_i, idx in enumerate(df.index, start=1):

        # GS코드 추출 — 모든 컬럼에서 검색

        gs_code9 = None

        if code_col and code_col in df.columns:

            m = re.search(r"(GS\d{7})", str(df.at[idx, code_col]) or "")

            gs_code9 = m.group(1) if m else None

        if not gs_code9 and name_col and name_col in df.columns:

            m = re.search(r"(GS\d{7})", str(df.at[idx, name_col]) or "")

            gs_code9 = m.group(1) if m else None

        if not gs_code9:

            # 코드 컬럼/상품명 컬럼에 없으면 전체 행에서 검색

            for c in df.columns:

                m = re.search(r"(GS\d{7})", str(df.at[idx, c]) or "")

                if m:

                    gs_code9 = m.group(1)

                    break

        if not gs_code9:

            skipped_no_code += 1

            continue



        # 중복 코드 스킵 (B/C/D 옵션 등)

        if gs_code9 in found_codes:

            continue

        found_codes.add(gs_code9)



        # OCR Excel에서 이미지 경로 가져오기

        found = False

        if ocr_lookup:

            for _mk in [gs_code9, f"{gs_code9}A"]:

                if _mk in ocr_lookup:

                    _imgs = ocr_lookup[_mk].get("images", [])

                    _valid = [p for p in _imgs if p.strip() and p.strip().lower() != "nan" and os.path.isfile(p)]

                    if _valid:

                        global_listing_sources.extend(_valid)

                        found = True

                    break



        # 로컬 폴더에서 직접 검색 — 파일명에 GS코드가 포함된 것만 (대표이미지용)

        # 1.jpg, 2.jpg 등은 OCR 상세페이지용이므로 제외

        if not found and local_root:

            hits = core.find_local_images_for_code(local_root, gs_code9, allow_folder_match=allow_folder_match, max_depth=max_depth)

            gs_lower = gs_code9.lower()

            listing_hits = [p for p in hits if gs_lower in os.path.basename(p).lower()]

            if listing_hits:

                global_listing_sources.extend(listing_hits)

                found = True



        if not found:

            skipped_no_img += 1



        if row_i % 50 == 0:

            _status(status_cb, f"이미지 소스 수집중... {row_i}/{total_rows} (코드 {len(found_codes)}개, 이미지 {len(global_listing_sources)}개)")

            _progress(progress_cb, 10 + int(20 * row_i / max(1, total_rows)))



    _status(status_cb, f"이미지 소스 수집 완료: 코드 {len(found_codes)}개, 이미지 {len(set(global_listing_sources))}개 (코드없음={skipped_no_code}, 이미지없음={skipped_no_img})")

    _progress(progress_cb, 30)



    if not global_listing_sources:

        _status(status_cb, "이미지 소스가 없습니다. 이미지 폴더 경로를 확인해 주세요.")

        _progress(progress_cb, 100)

        return ""



    # 대표이미지 생성

    csv_base = os.path.splitext(os.path.basename(cfg.file_path))[0]

    date_tag = datetime.now().strftime("%Y%m%d")

    export_root = os.path.join("C:\\code", "exports", f"{date_tag}_{csv_base}")

    listing_out_root = os.path.join(export_root, "listing_images", date_tag)

    os.makedirs(listing_out_root, exist_ok=True)



    logo_rgba = core._load_logo(cfg.logo_path.strip())

    listing_size = max(200, int(cfg.listing_size))

    listing_pad = max(0, int(cfg.listing_pad))

    listing_max = max(0, int(cfg.listing_max))

    logo_ratio = max(1, min(60, int(cfg.logo_ratio)))

    logo_opacity = max(0, min(100, int(cfg.logo_opacity)))

    logo_pos = cfg.logo_pos or "tr"

    jpeg_q_min = max(70, min(99, int(cfg.jpeg_q_min)))

    jpeg_q_max = max(jpeg_q_min, min(99, int(cfg.jpeg_q_max)))



    total_imgs = len(set(global_listing_sources))

    processed = 0



    def _pcb():

        nonlocal processed

        processed += 1

        pct = 30 + int((processed / max(1, total_imgs)) * 65)

        _progress(progress_cb, min(95, pct))



    _status(status_cb, f"대표이미지 생성중... ({total_imgs}개)")



    results = core.process_listing_images_global(

        src_paths=list(set(global_listing_sources)),

        base_out_root=listing_out_root,

        logo_rgba=logo_rgba,

        size=listing_size,

        pad=listing_pad,

        bg_color=(255, 255, 255),

        pos=logo_pos,

        opacity=logo_opacity,

        logo_ratio=logo_ratio,

        use_auto_contrast=bool(cfg.use_auto_contrast),

        use_sharpen=bool(cfg.use_sharpen),

        use_small_rotate=bool(cfg.use_small_rotate),

        rotate_zoom=float(cfg.rotate_zoom),

        max_images_per_code=listing_max,

        ultra_angle_deg=float(cfg.ultra_angle_deg),

        ultra_translate_px=float(cfg.ultra_translate_px),

        ultra_scale_pct=float(cfg.ultra_scale_pct),

        trim_tol=int(cfg.trim_tol),

        jpeg_q_min=jpeg_q_min,

        jpeg_q_max=jpeg_q_max,

        do_flip_lr=bool(cfg.do_flip_lr),

        progress_cb=_pcb,

    )



    _status(status_cb, f"대표이미지 생성 완료 - {len(results)}개 -> {listing_out_root}")

    _progress(progress_cb, 100)

    return listing_out_root







