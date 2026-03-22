"""OCR result Excel I/O utilities."""
from __future__ import annotations

import glob
import os
import re
from datetime import datetime

import pandas as pd
from openpyxl import Workbook


def _is_excel_lock_file(path: str) -> bool:
    name = os.path.basename(str(path or ""))
    return bool(name.startswith("~$"))


def _resolve_excel_path(path: str) -> str:
    p = str(path or "").strip()
    if not p:
        return p
    if _is_excel_lock_file(p):
        real_name = os.path.basename(p)[2:]
        real_path = os.path.join(os.path.dirname(p), real_name)
        if os.path.isfile(real_path):
            return real_path
    return p


def write_ocr_results(results: list[dict], metadata: dict, output_path: str) -> str:
    """Write OCR results as a single-sheet Excel file (OCR결과)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "OCR결과"
    ws.append([
        "GS코드",
        "상품명",
        "OCR텍스트_원본",
        "OCR텍스트_요약",
        "이미지파일목록",
        "이미지개수",
        "원본CSV파일",
    ])

    csv_basename = str(metadata.get("원본CSV파일명", ""))
    for r in results:
        gs = str(r.get("gs_code", "")).strip()
        if not gs or gs.lower() == "nan":
            continue
        raw = str(r.get("raw_text", "")).strip()
        summary = str(r.get("summary_text", "")).strip()
        paths = [str(x).strip() for x in (r.get("image_paths", []) or []) if str(x).strip()]
        try:
            cnt = int(r.get("image_count", 0) or 0)
        except Exception:
            cnt = 0

        ws.append([
            gs,
            str(r.get("product_name", "")),
            raw,
            summary,
            ";".join(paths),
            cnt,
            csv_basename,
        ])

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 50
    ws.column_dimensions["D"].width = 50
    ws.column_dimensions["E"].width = 40
    ws.column_dimensions["F"].width = 10
    ws.column_dimensions["G"].width = 25

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    wb.save(output_path)
    return output_path


def read_ocr_results(excel_path: str) -> tuple[dict, dict]:
    """Read OCR result Excel into lookup dict. Supports legacy and single-sheet schemas."""
    excel_path = _resolve_excel_path(excel_path)
    if not os.path.isfile(excel_path):
        raise FileNotFoundError(f"OCR Excel file not found: {excel_path}")
    if _is_excel_lock_file(excel_path):
        raise PermissionError(f"Excel lock file is not supported: {excel_path}")

    def _norm(s: str) -> str:
        return re.sub(r"[^0-9a-zA-Z\uAC00-\uD7A3]", "", str(s or "")).lower()

    def _pick_col(cols: list[str], aliases: list[str], contains: list[str] | None = None) -> str | None:
        norm_map = {_norm(c): c for c in cols}
        for a in aliases:
            k = _norm(a)
            if k in norm_map:
                return norm_map[k]
        if contains:
            for c in cols:
                nc = _norm(c)
                if all(tok in nc for tok in contains):
                    return c
        return None

    xls = pd.ExcelFile(excel_path)
    sheet = "OCR결과" if "OCR결과" in xls.sheet_names else (xls.sheet_names[0] if xls.sheet_names else None)
    if not sheet:
        return {}, {}

    df = pd.read_excel(excel_path, sheet_name=sheet)
    cols = [str(c) for c in df.columns]

    gs_col = _pick_col(cols, ["GS코드", "GS code", "gs_code"], contains=["gs", "코드"])
    raw_col = _pick_col(cols, ["OCR텍스트_원본", "OCR텍스트", "ocr_text_raw"], contains=["ocr", "텍스트"])
    summary_col = _pick_col(cols, ["OCR텍스트_요약", "OCR요약", "ocr_summary"], contains=["요약"])
    img_col = _pick_col(cols, ["이미지파일목록", "OCR이미지", "이미지경로", "ocr_image"], contains=["이미지"])
    count_col = _pick_col(cols, ["이미지개수", "전체이미지수", "OCR처리이미지수", "ocr_count"], contains=["이미지", "수"])

    lookup: dict[str, dict] = {}
    for _, row in df.iterrows():
        gs = str(row.get(gs_col, "") if gs_col else "").strip()
        if not gs or gs.lower() == "nan":
            continue

        raw = str(row.get(raw_col, "") if raw_col else "")
        if raw.lower() == "nan":
            raw = ""
        summary = str(row.get(summary_col, "") if summary_col else "")
        if summary.lower() == "nan":
            summary = ""

        imgs_raw = str(row.get(img_col, "") if img_col else "")
        if not imgs_raw or imgs_raw.lower() == "nan":
            images = []
        else:
            parts = re.split(r"[;|,]", imgs_raw)
            images = [str(x).strip() for x in parts if str(x).strip() and str(x).strip().lower() != "nan"]

        cnt_raw = row.get(count_col, 0) if count_col else 0
        try:
            count = int(cnt_raw)
        except Exception:
            count = 0

        lookup[gs] = {
            "raw": raw,
            "summary": summary,
            "images": images,
            "count": count,
        }

    metadata: dict[str, str] = {}
    for m_sheet in ["메타데이터", "metadata", "Metadata"]:
        if m_sheet in xls.sheet_names:
            try:
                df_meta = pd.read_excel(excel_path, sheet_name=m_sheet, header=None)
                for _, mrow in df_meta.iterrows():
                    if len(mrow) >= 2:
                        metadata[str(mrow.iloc[0])] = str(mrow.iloc[1])
            except Exception:
                pass
            break

    return lookup, metadata


def find_matching_ocr_file(csv_path: str, search_dirs: list[str] | None = None) -> str | None:
    """Find best matching OCR Excel for a given CSV path."""
    csv_basename = os.path.splitext(os.path.basename(csv_path))[0]

    strict_patterns = [
        f"OCR결과_{csv_basename}_*.xlsx",
        f"OCR_{csv_basename}_*.xlsx",
        f"*{csv_basename}*OCR*.xlsx",
    ]
    loose_patterns = ["OCR결과_*.xlsx", "OCR_*.xlsx", "*OCR*.xlsx"]

    dirs_to_search: list[str] = []
    csv_dir = os.path.dirname(os.path.abspath(csv_path))
    dirs_to_search.append(csv_dir)

    if search_dirs:
        dirs_to_search.extend(search_dirs)

    exports_root = r"C:\code\exports"
    if os.path.isdir(exports_root):
        for sub in os.listdir(exports_root):
            p = os.path.join(exports_root, sub)
            if os.path.isdir(p):
                dirs_to_search.append(p)

    seen = set()
    dirs_unique = []
    for d in dirs_to_search:
        d = os.path.abspath(d)
        if d not in seen and os.path.isdir(d):
            seen.add(d)
            dirs_unique.append(d)

    candidates: list[str] = []

    for d in dirs_unique:
        for pat in strict_patterns:
            candidates.extend(glob.glob(os.path.join(d, pat)))
    candidates = [p for p in candidates if os.path.isfile(p) and not _is_excel_lock_file(p)]

    if not candidates:
        for d in dirs_unique:
            for pat in loose_patterns:
                candidates.extend(glob.glob(os.path.join(d, pat)))
        candidates = [p for p in candidates if os.path.isfile(p) and not _is_excel_lock_file(p)]

    if not candidates:
        return None

    # Prefer *_02 fixed name first, then latest modified.
    preferred = [p for p in candidates if re.search(r"OCR결과_\d{8}_02\.xlsx$", os.path.basename(p))]
    pool = preferred or candidates
    pool.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return pool[0]
