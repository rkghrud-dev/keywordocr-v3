"""OCR 전용 파이프라인 — CSV + 로컬 이미지 → OCR → Excel 출력.

지원 OCR 엔진:
  - Tesseract (무료, 로컬)
  - Google Cloud Vision API (유료, 고정확도)
"""
from __future__ import annotations

import os
import re
import requests
import pandas as pd
from dataclasses import dataclass, field
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

from . import legacy_core as core
from .env_loader import ensure_env_loaded, get_env
from .ocr_excel import write_ocr_results


# ── Google Cloud Vision OCR ───────────────────────────────────────────

_gv_client = None  # 싱글톤 클라이언트 (재사용)

def _get_gv_client():
    global _gv_client
    if _gv_client is None:
        from google.cloud import vision
        _gv_client = vision.ImageAnnotatorClient()
    return _gv_client

def _ocr_google_vision(image_path: str) -> str:
    """Google Cloud Vision API로 이미지에서 텍스트 추출."""
    try:
        from google.cloud import vision
        client = _get_gv_client()
        with open(image_path, "rb") as f:
            content = f.read()
        image = vision.Image(content=content)
        response = client.text_detection(image=image)
        if response.error.message:
            print(f"[GV 오류] {os.path.basename(image_path)}: {response.error.message}")
            return ""
        texts = response.text_annotations
        if texts:
            return texts[0].description.strip()
        return ""
    except Exception as e:
        print(f"[GV 예외] {os.path.basename(image_path)}: {type(e).__name__}: {e}")
        return ""


# ── URL 이미지 다운로드 ───────────────────────────────────────────────

def _extract_image_urls(html_text: str) -> list[str]:
    """HTML에서 <img src="..."> URL 추출"""
    if not html_text:
        return []
    pattern = r'<img\s+[^>]*src=["\']([^"\']+)["\']'
    urls = re.findall(pattern, html_text, flags=re.IGNORECASE)
    return urls


def _download_image(url: str, save_path: str) -> bool:
    """URL에서 이미지 다운로드"""
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        with open(save_path, "wb") as f:
            f.write(response.content)
        return True
    except Exception as e:
        print(f"[다운로드 실패] {url}: {e}")
        return False


def _download_and_save_images(html_text: str, gs_code: str, target_dir: str) -> list[str]:
    """HTML에서 이미지 URL 추출 → 다운로드 → 경로 리스트 반환

    Args:
        html_text: HTML 텍스트 (O열 내용)
        gs_code: GS 코드 (폴더명)
        target_dir: 저장할 루트 디렉토리 (예: D:\Pp)

    Returns:
        다운로드된 이미지 경로 리스트
    """
    urls = _extract_image_urls(html_text)
    if not urls:
        return []

    # GS 코드 폴더 생성
    gs_folder = os.path.join(target_dir, gs_code)
    os.makedirs(gs_folder, exist_ok=True)

    downloaded_paths = []
    for idx, url in enumerate(urls, start=1):
        # 확장자 추출 (없으면 .jpg)
        ext = os.path.splitext(url)[-1].lower()
        if ext not in [".jpg", ".jpeg", ".png", ".webp", ".bmp"]:
            ext = ".jpg"

        filename = f"{idx}{ext}"
        save_path = os.path.join(gs_folder, filename)

        if _download_image(url, save_path):
            downloaded_paths.append(save_path)

    return downloaded_paths


def _download_sequential_images(base_url: str, gs_code: str, target_dir: str, max_fails: int = 3, max_images: int = 100) -> list[str]:
    """AU열 URL에서 순차 이미지 다운로드 (A1, A2, A3... → 3회 연속 실패시 중단)

    Args:
        base_url: AU열 URL 샘플 (예: http://ai.esmplus.com/.../GS2100178A1.jpg)
        gs_code: GS 코드 (폴더명)
        target_dir: 저장할 루트 디렉토리 (예: D:\Pp)
        max_fails: 연속 실패 허용 횟수 (기본 3)
        max_images: 최대 다운로드 이미지 수 (기본 100)

    Returns:
        다운로드된 이미지 경로 리스트

    로직:
        - base_url에서 "A1.jpg" 패턴을 찾아 base URL 추출
        - A1, A2, A3... 순차 다운로드
        - 3회 연속 실패시 중단
        - 예: A1✅ A2❌ A3✅ A4❌❌❌ → A1, A3 다운로드 후 중단
    """
    if not base_url or not gs_code:
        return []

    # URL 패턴 분석: GS2100178A1.jpg → GS2100178A + 1.jpg
    # base_url 예: http://ai.esmplus.com/.../GS2100178A/GS2100178A1.jpg
    match = re.search(r"(.*?)(GS\d{7}A)(\d+)\.(jpg|jpeg|png|webp|bmp)", base_url, flags=re.IGNORECASE)
    if not match:
        print(f"[순차다운] URL 패턴 인식 실패: {base_url}")
        return []

    url_prefix = match.group(1)  # http://ai.esmplus.com/.../
    gs_pattern = match.group(2)  # GS2100178A
    ext = match.group(4)         # jpg

    # GS 코드 폴더 생성
    gs_folder = os.path.join(target_dir, gs_code)
    os.makedirs(gs_folder, exist_ok=True)

    downloaded_paths = []
    fail_count = 0

    for i in range(1, max_images + 1):
        # URL 생성: http://...../GS2100178A1.jpg, A2.jpg, ...
        url = f"{url_prefix}{gs_pattern}{i}.{ext}"
        filename = f"{gs_pattern}{i}.{ext}"
        save_path = os.path.join(gs_folder, filename)

        # 다운로드 시도
        if _download_image(url, save_path):
            downloaded_paths.append(save_path)
            fail_count = 0  # 성공시 연속 실패 카운터 리셋
        else:
            fail_count += 1
            if fail_count >= max_fails:
                print(f"[순차다운] {gs_code}: {fail_count}회 연속 실패 → 중단 (마지막 시도: A{i})")
                break

    if downloaded_paths:
        print(f"[순차다운] {gs_code}: {len(downloaded_paths)}개 다운로드 완료")

    return downloaded_paths


# ── 마지막 이미지(주의사항) 제외 ──────────────────────────────────────

def _exclude_last_numbered_image(paths: list[str]) -> list[str]:
    """숫자 파일명(1.jpg, 2.jpg...) 중 마지막 번호 파일을 제외.

    상세페이지 마지막 이미지는 보통 주의사항/배송안내 등 불필요한 내용.
    """
    if len(paths) <= 1:
        return paths

    # 숫자 파일명만 추출하여 최대 번호 파악
    numbered: dict[int, str] = {}
    non_numbered: list[str] = []
    for p in paths:
        stem = os.path.splitext(os.path.basename(p))[0]
        if stem.isdigit():
            numbered[int(stem)] = p
        else:
            non_numbered.append(p)

    if not numbered:
        return paths  # 숫자 파일이 없으면 원본 그대로

    max_num = max(numbered.keys())
    # 마지막 번호 제외
    filtered = [numbered[n] for n in sorted(numbered.keys()) if n != max_num]
    return filtered + non_numbered


@dataclass
class OcrPipelineConfig:
    csv_path: str
    local_img_dir: str
    output_dir: str = ""
    tesseract_path: str = ""
    korean_only: bool = True
    psm: int = 11
    oem: int = 3
    max_depth: int = -1
    allow_folder_match: bool = True
    threads: int = 6
    max_imgs_per_code: int = 999
    skip_last_image: bool = True           # 마지막 이미지(주의사항) 제외
    use_google_vision: bool = False        # Google Vision API 사용
    google_credentials_path: str = ""      # 서비스 계정 JSON 경로
    filter_noise: bool = True              # 반복 문구 자동 필터링


def _status(cb, msg: str) -> None:
    if cb:
        cb(msg)


def _progress(cb, value: int) -> None:
    if cb:
        cb(int(value))


def run_ocr_pipeline(
    cfg: OcrPipelineConfig,
    status_cb=None,
    progress_cb=None,
) -> str:
    """OCR 파이프라인 실행. 결과 Excel 경로를 반환."""

    if not cfg.csv_path:
        raise ValueError("CSV 파일을 선택해 주세요.")
    if not cfg.local_img_dir or not os.path.isdir(cfg.local_img_dir):
        raise ValueError("로컬 이미지 폴더를 확인해 주세요.")

    # ── OCR 엔진 설정 ──
    use_google = bool(cfg.use_google_vision)
    ocr_engine_name = "Google Vision"

    if use_google:
        # Google Cloud 인증 설정 (.env/환경변수 우선)
        ensure_env_loaded()
        cred_path = cfg.google_credentials_path or get_env("GOOGLE_APPLICATION_CREDENTIALS", "GOOGLE_VISION_CREDENTIALS")
        if cred_path and os.path.isfile(cred_path):
            os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = cred_path
        if not os.environ.get("GOOGLE_APPLICATION_CREDENTIALS"):
            raise ValueError("Google Cloud 서비스 계정 JSON 파일을 지정해 주세요 (.env의 GOOGLE_APPLICATION_CREDENTIALS 가능).")
        _status(status_cb, "Google Cloud Vision API 사용")
    else:
        ocr_engine_name = "Tesseract"
        _status(status_cb, "Tesseract 초기화...")
        detected = core.setup_tesseract(cfg.tesseract_path or None)
        if not detected:
            raise ValueError("Tesseract 경로를 찾지 못했습니다.")

    psm = int(cfg.psm or 3)
    oem = int(cfg.oem or 3)
    korean_only = bool(cfg.korean_only)
    tess_lang = "kor" if korean_only else "kor+eng"
    threads = min(16, max(1, int(cfg.threads)))
    max_imgs = max(0, int(cfg.max_imgs_per_code))
    max_depth = int(cfg.max_depth)
    allow_folder_match = bool(cfg.allow_folder_match)
    skip_last = bool(cfg.skip_last_image)

    # CSV 읽기
    _status(status_cb, "CSV 파일 읽는 중...")
    _progress(progress_cb, 5)
    df = core.safe_read_csv(cfg.csv_path)
    if df.empty:
        raise ValueError("CSV 파일이 비어 있습니다.")

    # 상품명 컬럼 확인
    name_col = "상품명"
    if name_col not in df.columns:
        raise ValueError("'상품명' 컬럼이 없습니다.")

    # 코드 컬럼 탐색
    code_col = None
    for c in df.columns:
        if str(c).strip() in ["자체상품코드", "자체 상품코드", "상품코드B", "코드", "코드B"]:
            code_col = c
            break

    # 상품 상세설명 컬럼 확인 (URL 이미지 다운로드용)
    detail_desc_col = None
    for c in df.columns:
        if str(c).strip() in ["상품 상세설명", "상품상세설명", "상세설명"]:
            detail_desc_col = c
            break

    # ── OCR 함수 (엔진별 분기) ──
    def _do_ocr_single(path: str) -> str:
        if use_google:
            raw = _ocr_google_vision(path)
        else:
            raw = core.ocr_image_file(path, tess_lang, psm, oem, korean_only)
        # 한글 외 문자 제거 (Google Vision은 모든 문자 반환하므로 후처리 필요)
        if raw and korean_only:
            raw = re.sub(r"[^가-힣\s]", " ", raw)
            raw = re.sub(r"\s+", " ", raw).strip()
        return raw

    def ocr_paths(paths: list[str]) -> list[str]:
        texts: list[str] = []
        if not paths:
            return texts
        # Google Vision은 API rate limit 방지를 위해 동시 요청 제한
        workers = min(threads, 2) if use_google else threads
        with ThreadPoolExecutor(max_workers=workers) as ex:
            futs = {ex.submit(_do_ocr_single, p): p for p in paths}
            for fut in as_completed(futs):
                src = futs[fut]
                try:
                    t = fut.result()
                    if t:
                        texts.append(t)
                    else:
                        _status(status_cb, f"  OCR 텍스트 없음: {os.path.basename(src)}")
                except Exception as e:
                    _status(status_cb, f"  OCR 실패: {os.path.basename(src)} — {e}")
        return texts

    # ── GS코드+A 필터: 기본옵션(A)만 처리 ──
    # GS코드 뒤에 알파벳이 붙는 구조 (GSxxxxxxxA, GSxxxxxxxB...)
    # A만 기본상품이므로 A가 아닌 행은 스킵
    def _extract_gs_code_with_option(text: str) -> tuple[str | None, str]:
        """GS코드 + 옵션알파벳 추출. (gs_code7, option_letter) 반환."""
        m = re.search(r"(GS\d{7})([A-Za-z])?", text or "")
        if not m:
            return None, ""
        return m.group(1), (m.group(2) or "A").upper()

    # 행별 처리
    results: list[dict] = []
    total_rows = max(1, len(df))
    ocr_count = 0
    skipped_option = 0
    a_option_has_images: dict[str, bool] = {}  # GS코드별 A옵션 이미지 유무 기록

    _status(status_cb, f"OCR 처리 시작 ({total_rows}개 상품, {ocr_engine_name})...")
    _progress(progress_cb, 10)

    for row_i, idx in enumerate(df.index, start=1):
        try:
            full_pname = str(df.at[idx, name_col])

            # GS 코드 + 옵션 추출
            gs_code9 = None
            option_letter = "A"

            if code_col and code_col in df.columns:
                gs_code9, option_letter = _extract_gs_code_with_option(
                    str(df.at[idx, code_col]) or "")
            if not gs_code9:
                gs_code9, option_letter = _extract_gs_code_with_option(full_pname)

            # GS코드 없는 행 스킵
            if not gs_code9:
                results.append({
                    "gs_code": "",
                    "product_name": full_pname,
                    "raw_text": "",
                    "summary_text": "",
                    "image_paths": [],
                    "image_count": 0,
                })
                continue

            # A옵션(기본상품)이 아닌 행은 스킵
            if option_letter != "A":
                skipped_option += 1
                # A옵션에 이미지가 있었는지 확인
                a_has_images = a_option_has_images.get(gs_code9, False)
                raw_text_value = "(옵션상품 — A옵션 결과 참조)" if a_has_images else ""

                results.append({
                    "gs_code": f"{gs_code9}{option_letter}",
                    "product_name": full_pname,
                    "raw_text": raw_text_value,
                    "summary_text": "",
                    "image_paths": [],
                    "image_count": 0,
                })
                pct = 10 + int(80 * row_i / total_rows)
                _progress(progress_cb, pct)
                _status(status_cb, f"[{row_i}/{total_rows}] {gs_code9}{option_letter} — 옵션상품 스킵")
                continue

            # 이미지 탐색
            all_hits_raw = core.find_local_images_for_code(
                cfg.local_img_dir, gs_code9,
                allow_folder_match=allow_folder_match,
                max_depth=max_depth,
            )

            # 숫자 파일명(1.jpg, 2.jpg...)만 필터 — 상세페이지 이미지만 OCR 대상
            all_hits = [
                p for p in all_hits_raw
                if os.path.splitext(os.path.basename(p))[0].isdigit()
            ]

            # 폴더매칭은 됐는데 숫자파일이 없는 경우 → 폴더 내 모든 숫자 파일 직접 탐색
            if not all_hits and all_hits_raw:
                folder = os.path.dirname(all_hits_raw[0])
                for fn in os.listdir(folder):
                    stem, ext = os.path.splitext(fn)
                    if stem.isdigit() and ext.lower() in {".jpg", ".jpeg", ".png", ".webp", ".bmp"}:
                        all_hits.append(os.path.join(folder, fn))
            elif not all_hits and not all_hits_raw:
                # 폴더매칭 자체가 안 된 경우 → GS코드 폴더 직접 탐색
                for try_name in [gs_code9, f"{gs_code9}A"]:
                    direct_folder = os.path.join(cfg.local_img_dir, try_name)
                    if os.path.isdir(direct_folder):
                        for fn in os.listdir(direct_folder):
                            stem, ext = os.path.splitext(fn)
                            if stem.isdigit() and ext.lower() in {".jpg", ".jpeg", ".png", ".webp", ".bmp"}:
                                all_hits.append(os.path.join(direct_folder, fn))
                        if all_hits:
                            break

            # 이미지가 없고 상세설명 컬럼이 있으면 URL 이미지 다운로드 시도
            if not all_hits and detail_desc_col and detail_desc_col in df.columns:
                html_text = str(df.at[idx, detail_desc_col]) if pd.notna(df.at[idx, detail_desc_col]) else ""
                if html_text and "<img" in html_text.lower():
                    _status(status_cb, f"[{row_i}/{total_rows}] {gs_code9}A — URL 이미지 다운로드 중...")
                    downloaded = _download_and_save_images(html_text, f"{gs_code9}A", cfg.local_img_dir)
                    if downloaded:
                        all_hits = downloaded
                        _status(status_cb, f"[{row_i}/{total_rows}] {gs_code9}A — URL 이미지 {len(downloaded)}개 다운로드 완료")

            # 마지막 이미지(주의사항) 제외
            if skip_last and all_hits:
                all_hits = _exclude_last_numbered_image(all_hits)

            matched_count = len(all_hits)
            sel = all_hits[:max_imgs] if max_imgs > 0 else []

            # OCR 실행
            raw_texts = ocr_paths(sel)
            raw_combined = " ".join(raw_texts) if raw_texts else ""

            sum_text, _ = core.summarize_features_tokens(raw_combined, max_len=220) if raw_combined else ("", [])

            if raw_texts:
                ocr_count += 1

            # A옵션 이미지 유무 기록 (이미지가 1개 이상 있으면 True)
            a_option_has_images[gs_code9] = matched_count > 0

            results.append({
                "gs_code": f"{gs_code9}A",
                "product_name": full_pname,
                "raw_text": raw_combined,
                "summary_text": sum_text,
                "image_paths": [os.path.abspath(p) for p in sel],
                "image_count": matched_count,
            })

            pct = 10 + int(80 * row_i / total_rows)
            _progress(progress_cb, pct)
            if raw_texts:
                _status(status_cb, f"[{row_i}/{total_rows}] {gs_code9}A — 이미지 {matched_count}장, OCR 완료")
            else:
                _status(status_cb, f"[{row_i}/{total_rows}] {gs_code9}A — ★ 이미지 없음 (탐색: {cfg.local_img_dir})")

        except Exception as e:
            _status(status_cb, f"[{row_i}] 오류: {e}")
            results.append({
                "gs_code": "",
                "product_name": str(df.at[idx, name_col]) if name_col in df.columns else "",
                "raw_text": "",
                "summary_text": "",
                "image_paths": [],
                "image_count": 0,
            })

    # ── 반복 문구 자동 학습 (다음 키워드 생성 시 활용) ──
    if cfg.filter_noise:
        try:
            from .ocr_noise_filter import learn_from_batch, load_learned_db
            _learn_db = load_learned_db()
            _learn_db = learn_from_batch(results, _learn_db)
            _learned_count = len(_learn_db.get("phrases", {}))
            if _learned_count:
                _status(status_cb, f"반복 문구 학습 완료: {_learned_count}개 축적됨")
        except Exception:
            pass

    # Excel 저장
    _status(status_cb, "Excel 파일 저장 중...")
    _progress(progress_cb, 92)

    csv_basename = os.path.splitext(os.path.basename(cfg.csv_path))[0]
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_filename = f"OCR결과_{csv_basename}_{ts}.xlsx"

    out_dir = cfg.output_dir or os.path.dirname(os.path.abspath(cfg.csv_path))
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, out_filename)

    metadata = {
        "원본CSV파일명": os.path.basename(cfg.csv_path),
        "이미지루트폴더": cfg.local_img_dir,
        "처리일시": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "OCR엔진": ocr_engine_name,
        "PSM": psm if not use_google else "N/A",
        "OEM": oem if not use_google else "N/A",
        "한글전용": korean_only,
        "마지막이미지제외": skip_last,
        "전체상품수": total_rows,
        "OCR처리상품수": ocr_count,
        "옵션스킵수": skipped_option,
    }

    write_ocr_results(results, metadata, out_path)

    _progress(progress_cb, 100)
    _status(status_cb, f"완료! OCR: {ocr_count}개 처리, 옵션 스킵: {skipped_option}개 → {out_filename}")

    return out_path



