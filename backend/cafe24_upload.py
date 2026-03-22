from __future__ import annotations

import os
import re
import time
from datetime import datetime

import pandas as pd

from app.services import cafe24
from app.services.env_loader import ensure_env_loaded, get_env


def load_config(path: str = "cafe24_upload_config.txt") -> dict:
    cfg = {
        "DATE_TAG": "",
        "MAIN_INDEX": "2",
        "ADD_START": "3",
        "ADD_MAX": "10",
        "EXPORT_DIR": "",
        "IMAGE_ROOT": "",
        "RETRY_COUNT": "1",
        "RETRY_DELAY": "1.0",
        "LOG_PATH": "",
        "MATCH_MODE": "PREFIX",
        "MATCH_PREFIX": "20",
        "GS_LIST": "",
        "PRICE_DATA": "",
    }

    ensure_env_loaded()
    env_map = {
        "DATE_TAG": ["CAFE24_UPLOAD_DATE_TAG"],
        "MAIN_INDEX": ["CAFE24_UPLOAD_MAIN_INDEX"],
        "ADD_START": ["CAFE24_UPLOAD_ADD_START"],
        "ADD_MAX": ["CAFE24_UPLOAD_ADD_MAX"],
        "EXPORT_DIR": ["CAFE24_UPLOAD_EXPORT_DIR"],
        "IMAGE_ROOT": ["CAFE24_UPLOAD_IMAGE_ROOT"],
        "RETRY_COUNT": ["CAFE24_UPLOAD_RETRY_COUNT"],
        "RETRY_DELAY": ["CAFE24_UPLOAD_RETRY_DELAY"],
        "LOG_PATH": ["CAFE24_UPLOAD_LOG_PATH"],
        "MATCH_MODE": ["CAFE24_UPLOAD_MATCH_MODE"],
        "MATCH_PREFIX": ["CAFE24_UPLOAD_MATCH_PREFIX"],
        "GS_LIST": ["CAFE24_UPLOAD_GS_LIST"],
        "PRICE_DATA": ["CAFE24_UPLOAD_PRICE_DATA"],
    }
    for k, names in env_map.items():
        v = get_env(*names)
        if v:
            cfg[k] = v

    try:
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                if "=" in line:
                    k, v = line.split("=", 1)
                    k = k.strip().upper()
                    v = v.strip().strip('"').strip("'")
                    if k in cfg and not cfg[k]:
                        cfg[k] = v
    except FileNotFoundError:
        return cfg
    return cfg

def pick_images(folder: str, main_index: int, add_start: int, add_max: int) -> tuple[str | None, list[str]]:
    files = [f for f in os.listdir(folder) if f.lower().endswith((".jpg", ".jpeg", ".png", ".webp"))]
    files.sort()
    if len(files) < main_index:
        return None, []
    main = os.path.join(folder, files[main_index - 1])
    start_idx = max(add_start - 1, 0)
    adds = [os.path.join(folder, f) for f in files[start_idx:start_idx + add_max]]
    return main, adds


def pick_images_by_selection(folder: str, selection: dict) -> tuple[str | None, list[str]]:
    """선택된 인덱스로 이미지 반환

    Args:
        folder: 이미지 폴더 경로
        selection: {"main": idx, "additional": [idx1, idx2, ...]}

    Returns:
        (main_image_path, [add_image_path1, add_image_path2, ...])
    """
    files = [f for f in os.listdir(folder) if f.lower().endswith((".jpg", ".jpeg", ".png", ".webp", ".bmp"))]
    files.sort()

    main_idx = selection.get("main")
    if main_idx is None or main_idx >= len(files):
        return None, []

    main = os.path.join(folder, files[main_idx])

    add_indices = selection.get("additional", [])
    adds = [os.path.join(folder, files[i]) for i in add_indices if i < len(files)]

    return main, adds


import math


def _get_multiplier(supply_price: float) -> float:
    if supply_price >= 20000:
        return 1.6
    elif supply_price >= 10000:
        return 1.8
    else:
        return 2.0


def _ceil10(v: float) -> int:
    """10원 단위 올림."""
    return int(math.ceil(v / 10)) * 10


def _ceil100(v: float) -> int:
    """100원 단위 올림."""
    return int(math.ceil(v / 100)) * 100


def calc_option_prices(supply_prices: list[float]) -> dict:
    """옵션별 공급가 리스트 → {base_selling, base_consumer, additional_amounts} 계산.

    로직:
    - 소액 구간 (최소 공급가 ≤ 100원): 고유 공급가별 10원 간격 강제 배분.
      같은 공급가면 같은 판매가, 다른 공급가면 +10원씩 증가.
    - 일반 구간 (최소 공급가 > 100원): 공급가 × 배율 → 100원 단위 올림.
      같은 판매가가 되는 서로 다른 공급가가 있으면 차이×2 기준으로 벌림
      (49원 이하 차이 → +50원, 50원 이상 → +100원).

    Returns:
        dict with keys:
            base_selling (int): 첫 번째 옵션(A)의 판매가
            base_consumer (int): 소비자가 (판매가 × 1.2, 100원 올림)
            additional_amounts (list[int]): 각 옵션의 추가금액 (A 기준 차이)
    """
    if not supply_prices:
        return {"base_selling": 0, "base_consumer": 0, "additional_amounts": []}

    min_sp = min(supply_prices)

    if min_sp <= 100:
        # === 소액 구간 ===
        unique_sorted = sorted(set(supply_prices))
        base = _ceil10(unique_sorted[0] * _get_multiplier(unique_sorted[0]))
        price_map = {}
        for i, usp in enumerate(unique_sorted):
            price_map[usp] = base + i * 10
        selling = [price_map[sp] for sp in supply_prices]
    else:
        # === 일반 구간 ===
        # 고유 공급가별 판매가 매핑
        unique_sps = sorted(set(supply_prices))
        sell_map = {}
        for usp in unique_sps:
            sell_map[usp] = _ceil100(usp * _get_multiplier(usp))

        # 충돌 해결: 같은 판매가에 다른 공급가가 매핑된 경우 벌리기
        rev_map: dict[int, list[float]] = {}
        for usp in unique_sps:
            rev_map.setdefault(sell_map[usp], []).append(usp)

        for sv, sp_group in rev_map.items():
            if len(sp_group) <= 1:
                continue
            sp_group.sort()
            for j in range(1, len(sp_group)):
                diff = (sp_group[j] - sp_group[0]) * 2
                adj = 50 if diff <= 49 else _ceil100(diff)
                sell_map[sp_group[j]] = sv + adj

        selling = [sell_map[sp] for sp in supply_prices]

    base_price = selling[0]
    additional = [s - base_price for s in selling]
    consumer = _ceil100(base_price * 1.2)
    return {
        "base_selling": base_price,
        "base_consumer": consumer,
        "additional_amounts": additional,
    }


def compute_split_groups(sell_prices: list[int], threshold_pct: float = 100.0) -> list[list[int]]:
    """판매가 리스트를 받아 추가금액이 threshold_pct% 이내가 되도록 그룹 분할.

    판매가를 오름차순 정렬 후 greedy 방식으로 그룹핑한다.
    각 그룹의 첫 번째 옵션이 base가 되고, 나머지 옵션의 추가금액이
    base × (threshold_pct / 100) 이내여야 한다.

    Args:
        sell_prices: 옵션별 판매가 리스트 (예: [1000, 2000, 3000, 4000, 5000])
        threshold_pct: 추가금액 한도 비율 (기본 100%)

    Returns:
        그룹별 원본 인덱스 리스트. 예: [[0, 1], [2, 3], [4]]
        각 그룹 내에서 판매가 오름차순 정렬되어 있음.
    """
    if not sell_prices:
        return []
    if len(sell_prices) == 1:
        return [[0]]

    threshold_ratio = threshold_pct / 100.0

    # 판매가 기준 오름차순 정렬 (원본 인덱스 보존)
    sorted_indices = sorted(range(len(sell_prices)), key=lambda i: sell_prices[i])

    groups: list[list[int]] = []
    current_group = [sorted_indices[0]]
    group_base = sell_prices[sorted_indices[0]]

    for idx in sorted_indices[1:]:
        additional = sell_prices[idx] - group_base
        # base가 0이면 분할 불가, 모두 같은 그룹
        if group_base > 0 and additional > group_base * threshold_ratio:
            groups.append(current_group)
            current_group = [idx]
            group_base = sell_prices[idx]
        else:
            current_group.append(idx)

    groups.append(current_group)
    return groups


def load_option_supply_prices(file_path: str) -> dict[str, list[tuple[str, float]]]:
    """분리추출전 시트에서 GS코드별 옵션 공급가 로드.

    Args:
        file_path: 상품전처리GPT 파일 또는 업로드용 파일 경로.
                   '분리추출전' 시트를 우선 시도하고, 없으면 '분리추출후'도 시도.

    Returns:
        {GS9자리: [(옵션명suffix, 공급가), ...]} 형태
        예: {"GS2600105": [("A 파랑", 4298), ("B 빨강", 5029), ("C 검정", 5628)]}
    """
    result = {}
    df = None
    for sheet in ("분리추출전", "분리추출후"):
        try:
            df = pd.read_excel(file_path, sheet_name=sheet)
            break
        except Exception:
            continue
    if df is None:
        print(f"  [가격] 시트를 읽지 못했습니다: {file_path}")
        return result

    # 공급가 컬럼 찾기
    supply_col = next((col for col in df.columns if "공급가" in str(col)), None)
    # 상품명 컬럼 찾기
    name_col = next((col for col in df.columns if str(col).strip() == "상품명"), None)
    if not supply_col or not name_col:
        print(f"  [가격] 공급가/상품명 컬럼 없음 (sheet cols: {list(df.columns)[:5]}...)")
        return result

    for _, row in df.iterrows():
        name = str(row.get(name_col, ""))
        m = re.search(r"(GS\d{7})", name, flags=re.IGNORECASE)
        if not m:
            continue
        gs9 = m.group(1).upper()
        # GS코드 뒤의 옵션 부분 (예: "A 파랑", "B 빨강" 등)
        suffix = name[m.end():].strip()
        try:
            sp = float(row[supply_col])
        except (ValueError, TypeError):
            sp = 0.0
        result.setdefault(gs9, []).append((suffix, sp))

    return result


def apply_variant_prices(mall_id: str, token: str, api_version: str,
                         product_no: int, option_data: list[tuple[str, float]],
                         shop_no: str = "1", cfg_path: str = "", cfg: dict = None) -> str:
    """상품의 variants를 조회한 뒤 옵션별 가격 차등 적용.

    Args:
        option_data: [(suffix, 공급가), ...] 옵션 순서대로

    Returns:
        상태 문자열 ("PRICE_OK", "PRICE_SKIP", "PRICE_ERROR: ...")
    """
    if not option_data or len(option_data) < 1:
        return "PRICE_SKIP"

    supply_prices = [sp for _, sp in option_data]

    # 공급가가 전부 동일하면 차등 불필요
    if len(set(supply_prices)) <= 1 and len(supply_prices) > 1:
        # 전부 같으면 additional_amount = 0이므로 업데이트 불필요
        return "PRICE_SKIP_SAME"

    prices = calc_option_prices(supply_prices)
    additional_amounts = prices["additional_amounts"]

    # 차등이 없으면 스킵
    if all(a == 0 for a in additional_amounts):
        return "PRICE_SKIP_NO_DIFF"

    try:
        variants = cafe24.get_variants(mall_id, token, api_version, product_no, shop_no=shop_no)
    except RuntimeError as e:
        if str(e) == "TOKEN_EXPIRED" and cfg_path and cfg is not None:
            cfg.update(cafe24.refresh_access_token(cfg_path, cfg))
            token = cfg.get("ACCESS_TOKEN", "")
            variants = cafe24.get_variants(mall_id, token, api_version, product_no, shop_no=shop_no)
        else:
            return f"PRICE_ERROR: {e}"
    except Exception as e:
        return f"PRICE_ERROR: variants조회실패 {e}"

    print(f"    variants({len(variants)}개): {[v.get('variant_code','?') for v in variants[:5]]}")

    if len(variants) != len(option_data):
        return f"PRICE_ERROR: variant수({len(variants)})≠옵션수({len(option_data)})"

    errors = []
    for i, var in enumerate(variants):
        vc = var.get("variant_code")
        amt = additional_amounts[i]
        amt_str = f"{amt:.2f}"
        try:
            cafe24.update_variant(
                mall_id, token, api_version, product_no, vc,
                {"additional_amount": amt_str}, shop_no=shop_no,
            )
        except RuntimeError as e:
            if str(e) == "TOKEN_EXPIRED" and cfg_path and cfg is not None:
                cfg.update(cafe24.refresh_access_token(cfg_path, cfg))
                token = cfg.get("ACCESS_TOKEN", "")
                cafe24.update_variant(
                    mall_id, token, api_version, product_no, vc,
                    {"additional_amount": amt_str}, shop_no=shop_no,
                )
            else:
                errors.append(f"{vc}:{e}")
        except Exception as e:
            errors.append(f"{vc}:{e}")

    if errors:
        return f"PRICE_PARTIAL: {'; '.join(errors)}"
    return "PRICE_OK"


def main():
    cfg_path = "cafe24_token.txt"
    cfg = cafe24.load_cafe24_config(cfg_path)
    upcfg = load_config("cafe24_upload_config.txt")

    mall_id = cfg.get("MALL_ID", "")
    token = cfg.get("ACCESS_TOKEN", "")
    shop_no = cfg.get("SHOP_NO", "1")
    api_version = cfg.get("API_VERSION", "2024-09-01")

    if not mall_id or not token:
        raise SystemExit("cafe24_token.txt에 MALL_ID/ACCESS_TOKEN이 필요합니다.")

    date_tag = upcfg.get("DATE_TAG", "") or datetime.now().strftime("%Y%m%d")
    main_index = int(upcfg.get("MAIN_INDEX", "2") or 2)
    add_start = int(upcfg.get("ADD_START", "3") or 3)
    add_max = int(upcfg.get("ADD_MAX", "10") or 10)
    export_dir = upcfg.get("EXPORT_DIR", "")
    image_root = upcfg.get("IMAGE_ROOT", "")
    retry_count = int(upcfg.get("RETRY_COUNT", "1") or 1)
    retry_delay = float(upcfg.get("RETRY_DELAY", "1.0") or 1.0)
    log_path = upcfg.get("LOG_PATH", "")
    match_mode = (upcfg.get("MATCH_MODE", "PREFIX") or "PREFIX").upper()
    match_prefix = int(upcfg.get("MATCH_PREFIX", "20") or 20)
    gs_list_path = upcfg.get("GS_LIST", "")
    price_data_path = upcfg.get("PRICE_DATA", "")

    # 가격 리뷰 데이터 로드 (GUI 옵션 가격 확인 탭에서 전달)
    price_review = None
    image_selections = {}  # {gs_code: {"main": idx, "additional": [...]}}
    if price_data_path and os.path.isfile(price_data_path):
        import json
        try:
            with open(price_data_path, "r", encoding="utf-8") as f:
                price_review = json.load(f)
            print(f"[가격리뷰] 로드 완료: {len(price_review.get('checked_gs', []))}개 선택됨")
            # 이미지 선택 정보 추출
            image_selections = price_review.get("image_selections", {})
            if image_selections:
                print(f"[이미지선택] {len(image_selections)}개 상품의 이미지 선택 정보 로드")
        except Exception as e:
            print(f"[가격리뷰] 로드 실패: {e}")

    if image_root:
        img_root = image_root
        latest = os.path.dirname(img_root)
    else:
        base_root = export_dir or os.path.join("C:\\code", "exports")
        # export_dir가 이미 listing_images 폴더를 가리키는 경우
        if os.path.basename(base_root.rstrip("/\\")) == "listing_images":
            img_root = os.path.join(base_root, date_tag)
            if not os.path.isdir(img_root):
                img_root = base_root
            latest = os.path.dirname(base_root)
        else:
            export_dirs = [d for d in os.listdir(base_root) if os.path.isdir(os.path.join(base_root, d))]
            export_dirs.sort(reverse=True)
            if not export_dirs:
                raise SystemExit("exports 폴더가 없습니다.")
            latest = os.path.join(base_root, export_dirs[0])
            img_root = os.path.join(latest, "listing_images", date_tag)
            if not os.path.isdir(img_root):
                img_root = os.path.join(latest, "listing_images")

    # 방어: listing_images 중복 경로 자동 정리
    img_root = os.path.normpath(img_root)
    dup = os.path.join("listing_images", "listing_images")
    while dup in img_root:
        img_root = img_root.replace(dup, "listing_images")

    # 경로가 없으면 주변 후보 경로로 폴백
    if not os.path.isdir(img_root):
        cands = [os.path.join(img_root, date_tag)]
        if os.path.basename(img_root).lower() == "listing_images":
            cands += [img_root, os.path.join(img_root, date_tag)]
        parent = os.path.dirname(img_root)
        if os.path.basename(parent).lower() == "listing_images":
            cands += [parent, os.path.join(parent, date_tag)]
        for c in cands:
            if c and os.path.isdir(c):
                img_root = os.path.normpath(c)
                break

    # img_root 기준으로 latest 재정렬
    p = os.path.normpath(img_root)
    if os.path.basename(p).lower() == "listing_images":
        latest = os.path.dirname(p)
    elif os.path.basename(os.path.dirname(p)).lower() == "listing_images":
        latest = os.path.dirname(os.path.dirname(p))

    print("[IMG ROOT]", img_root)

    # load upload file to get product names for matching (업로드용_YYYYMMDD_01.xlsx 등)
    import glob as _glob2
    upload_candidates = sorted(_glob2.glob(os.path.join(latest, f"업로드용_{date_tag}*.xlsx")))
    upload_file = upload_candidates[-1] if upload_candidates else os.path.join(latest, f"업로드용_{date_tag}.xlsx")
    if not os.path.isfile(upload_file):
        print("[WARN] 업로드용 파일을 찾지 못했습니다:", upload_file)
        upload_names = []
    else:
        try:
            df_upload = pd.read_excel(upload_file, sheet_name="분리추출후")
            if "상품명" in df_upload.columns:
                upload_names = [str(x) for x in df_upload["상품명"].fillna("").tolist() if str(x).strip()]
            else:
                upload_names = []
        except Exception:
            upload_names = []

    # 옵션별 공급가 로드: 상품전처리GPT 파일(분리추출전 시트 포함)에서 읽기
    import glob as _glob
    option_prices_map = {}
    gpt_files = sorted(_glob.glob(os.path.join(latest, "상품전처리GPT_*.xlsx")))
    if gpt_files:
        gpt_file = gpt_files[-1]  # 가장 최근 파일
        print(f"[옵션가격] GPT파일: {os.path.basename(gpt_file)}")
        option_prices_map = load_option_supply_prices(gpt_file)
    else:
        # fallback: 업로드용 파일에서 시도
        if os.path.isfile(upload_file):
            print(f"[옵션가격] GPT파일 없음, 업로드용 파일에서 시도")
            option_prices_map = load_option_supply_prices(upload_file)
    if option_prices_map:
        multi = sum(1 for v in option_prices_map.values() if len(v) > 1)
        print(f"[옵션가격] {len(option_prices_map)}개 GS코드 로드 (다중옵션: {multi}개)")
    else:
        print("[옵션가격] 옵션 공급가 데이터 없음")

    folders = [d for d in os.listdir(img_root) if os.path.isdir(os.path.join(img_root, d))]
    folders = [f for f in folders if re.match(r"GS\d{7}", f, flags=re.IGNORECASE)]


    # Guard: if img_root points to listing_images (parent), descend into single dated folder automatically.
    if not folders:
        subdirs = [d for d in os.listdir(img_root) if os.path.isdir(os.path.join(img_root, d))]
        if len(subdirs) == 1:
            nested_root = os.path.join(img_root, subdirs[0])
            nested_folders = [d for d in os.listdir(nested_root) if os.path.isdir(os.path.join(nested_root, d))]
            nested_folders = [f for f in nested_folders if re.match(r"GS\d{7}", f, flags=re.IGNORECASE)]
            if nested_folders:
                img_root = nested_root
                folders = nested_folders
                print(f"[IMG ROOT] 자동보정: {img_root}")
    # 가격 리뷰에서 checked_gs 목록 확인 (있으면 옵션상품, 없으면 단일상품)
    # ※ checked_gs에 있는 상품은 옵션상품으로 처리하고,
    #    나머지는 단일상품으로 업로드
    option_gs_set = set()  # 옵션 상품 GS 코드
    if price_review and price_review.get("checked_gs"):
        option_gs_set = {gs.upper()[:9] for gs in price_review["checked_gs"]}
        print(f"[옵션상품] {len(option_gs_set)}개 선택됨 (나머지는 단일상품으로 업로드)")
        # ※ 중요: 필터링하지 않고 모든 폴더를 유지 (단일 상품도 업로드하기 위해)
    elif gs_list_path and os.path.isfile(gs_list_path):
        with open(gs_list_path, "r", encoding="utf-8") as f:
            wanted = {line.strip().upper() for line in f if line.strip()}
        folders = [f for f in folders if f.upper() in wanted]

    products = []
    try:
        for p in cafe24.iter_products(mall_id, token, api_version, shop_no=shop_no):
            if str(p.get("selling", "")).upper() != "T":
                continue
            products.append(p)
    except RuntimeError as e:
        if str(e) == "TOKEN_EXPIRED":
            cfg = cafe24.refresh_access_token(cfg_path, cfg)
            token = cfg.get("ACCESS_TOKEN", "")
            for p in cafe24.iter_products(mall_id, token, api_version, shop_no=shop_no):
                if str(p.get("selling", "")).upper() != "T":
                    continue
                products.append(p)
        else:
            raise

    log_rows = []

    total = len(folders)
    done = 0

    # 옵션 상품 vs 단일 상품 분류
    option_count = sum(1 for f in folders if f.upper()[:9] in option_gs_set)
    single_count = total - option_count
    print(f"STATUS 총 {total}개 상품 처리 시작 (옵션: {option_count}개, 단일: {single_count}개)")
    for folder in folders:
        gs = folder.upper()
        gs9 = gs[:9]
        print(f"STATUS [{done+1}/{total}] {gs} 매칭 중...")
        folder_path = os.path.join(img_root, folder)

        # 이미지 선택: 사용자가 선택한 정보가 있으면 사용, 없으면 기본 로직
        if gs9 in image_selections:
            main_img, add_imgs = pick_images_by_selection(folder_path, image_selections[gs9])
            print(f"  [이미지] {gs9}: 사용자 선택 적용 (대표: {image_selections[gs9].get('main')}, 추가: {len(add_imgs)}장)")
        else:
            main_img, add_imgs = pick_images(folder_path, main_index, add_start, add_max)

        if not main_img:
            log_rows.append({"GS": gs, "STATUS": "NO_MAIN_IMAGE"})
            print(f"STATUS [{done+1}/{total}] {gs} 대표이미지 없음 → 스킵")
            continue

        matched = None
        # 1) custom_product_code 매칭 (자체 상품코드에 GS코드 포함 여부)
        for p in products:
            cpc = str(p.get("custom_product_code", "")).upper()
            if cpc and gs in cpc:
                matched = p
                break
        # 2) match using upload file product names
        if not matched:
            for up_name in upload_names:
                if gs in up_name:
                    for p in products:
                        if str(p.get("product_name", "")) == up_name:
                            matched = p
                            break
                    if matched:
                        break
        # 3) fallback to matching against product_name
        if not matched:
            for p in products:
                name = str(p.get("product_name", ""))
                if not name:
                    continue
                if gs in name:
                    matched = p
                    break
                n = cafe24.normalize_name(name)
                g = cafe24.normalize_name(gs)
                if match_mode == "PREFIX":
                    if n[:match_prefix] and g and g in n[:match_prefix]:
                        matched = p
                        break
                elif match_mode == "CONTAINS":
                    if g and g in n:
                        matched = p
                        break
                elif match_mode == "EXACT":
                    if g and g == n:
                        matched = p
                        break

        if not matched:
            log_rows.append({"GS": gs, "STATUS": "NO_PRODUCT_MATCH"})
            print(f"STATUS [{done+1}/{total}] {gs} 상품 매칭 실패 → 스킵")
            continue

        product_no = matched.get("product_no")
        print(f"STATUS [{done+1}/{total}] {gs} → 이미지 업로드 중 (대표+추가 {len(add_imgs)}장)...")
        ok = False
        err_msg = ""
        for attempt in range(retry_count + 1):
            try:
                cafe24.upload_main_image(mall_id, token, api_version, product_no, main_img, shop_no=shop_no)
                for img in add_imgs:
                    cafe24.upload_additional_image(mall_id, token, api_version, product_no, img, shop_no=shop_no)
                ok = True
                break
            except RuntimeError as e:
                if str(e) == "TOKEN_EXPIRED":
                    cfg = cafe24.refresh_access_token(cfg_path, cfg)
                    token = cfg.get("ACCESS_TOKEN", "")
                    continue
                err_msg = str(e)
            except Exception as e:
                err_msg = str(e)
            if attempt < retry_count:
                print(f"[RETRY] {gs} {attempt+1}/{retry_count}")
                time.sleep(retry_delay)

        # 옵션 가격 차등 적용
        price_status = ""
        gs9 = gs[:9]  # GS2600104 (GS + 7자리 = 9글자)
        opt_data = option_prices_map.get(gs9, [])

        # 가격리뷰 선택 목록에 있거나, 공급가 데이터가 다중옵션이면 옵션 처리
        is_option_product = (gs9 in option_gs_set) or (len(opt_data) > 1)

        # 가격 리뷰에서 수정된 추가금액이 있으면 사용
        edited_amounts = None
        if is_option_product and price_review and price_review.get("edited_amounts"):
            edited_amounts = price_review["edited_amounts"].get(gs9)

        if ok and is_option_product and edited_amounts is not None and len(edited_amounts) > 0:
            # GUI에서 수정된 추가금액 직접 적용
            print(f"  [가격] {gs9}: 리뷰 수정금액 적용 → {edited_amounts}")
            try:
                variants = cafe24.get_variants(mall_id, token, api_version, product_no, shop_no=shop_no)
            except RuntimeError as e:
                if str(e) == "TOKEN_EXPIRED":
                    cfg = cafe24.refresh_access_token(cfg_path, cfg)
                    token = cfg.get("ACCESS_TOKEN", "")
                    variants = cafe24.get_variants(mall_id, token, api_version, product_no, shop_no=shop_no)
                else:
                    variants = []
                    price_status = f"PRICE_ERROR: {e}"

            if variants and len(variants) == len(edited_amounts):
                errors = []
                for vi, var in enumerate(variants):
                    vc = var.get("variant_code")
                    amt = edited_amounts[vi]
                    amt_str = f"{amt:.2f}"
                    try:
                        cafe24.update_variant(
                            mall_id, token, api_version, product_no, vc,
                            {"additional_amount": amt_str}, shop_no=shop_no,
                        )
                    except RuntimeError as e:
                        if str(e) == "TOKEN_EXPIRED":
                            cfg = cafe24.refresh_access_token(cfg_path, cfg)
                            token = cfg.get("ACCESS_TOKEN", "")
                            cafe24.update_variant(
                                mall_id, token, api_version, product_no, vc,
                                {"additional_amount": amt_str}, shop_no=shop_no,
                            )
                        else:
                            errors.append(f"{vc}:{e}")
                    except Exception as e:
                        errors.append(f"{vc}:{e}")
                price_status = f"PRICE_PARTIAL: {'; '.join(errors)}" if errors else "PRICE_OK"
            elif variants:
                price_status = f"PRICE_ERROR: variant수({len(variants)})≠편집금액수({len(edited_amounts)})"
            token = cfg.get("ACCESS_TOKEN", token)
            print(f"  [가격] {gs9} → {price_status}")
        elif ok and is_option_product and opt_data and len(opt_data) > 1:
            supply_prices = [sp for _, sp in opt_data]
            prices_calc = calc_option_prices(supply_prices)
            print(f"  [가격] {gs9}: 옵션{len(opt_data)}개, "
                  f"공급가={supply_prices}, "
                  f"base={prices_calc['base_selling']}, "
                  f"추가금={prices_calc['additional_amounts']}")
            price_status = apply_variant_prices(
                mall_id, token, api_version, product_no, opt_data,
                shop_no=shop_no, cfg_path=cfg_path, cfg=cfg,
            )
            # token이 갱신됐을 수 있으므로 반영
            token = cfg.get("ACCESS_TOKEN", token)
            print(f"  [가격] {gs9} → {price_status}")
        elif ok and is_option_product and opt_data and len(opt_data) == 1:
            print(f"  [가격] {gs9}: 단일옵션 → 스킵")
        elif ok and not is_option_product:
            print(f"  [가격] {gs9}: 단일상품 → 이미지만 업로드")
        elif ok and not opt_data:
            print(f"  [가격] {gs9}: 공급가 데이터 없음")

        sel_info = image_selections.get(gs9, {}) if isinstance(image_selections, dict) else {}

        if ok:
            status_msg = "OK"
            if price_status and price_status.startswith("PRICE_OK"):
                status_msg = "OK (가격적용)"
            log_rows.append({
                "GS": gs,
                "PRODUCT_NO": product_no,
                "STATUS": "OK",
                "MAIN": os.path.basename(main_img),
                "ADD_COUNT": len(add_imgs),
                "ADD_FILES": ",".join([os.path.basename(x) for x in add_imgs]),
                "SELECT_MAIN_IDX": sel_info.get("main", ""),
                "SELECT_ADD_IDX": ",".join([str(x) for x in (sel_info.get("additional", []) or [])]),
                "PRICE": price_status,
            })
        else:
            status_msg = f"ERROR: {err_msg[:30]}"
            log_rows.append({
                "GS": gs,
                "PRODUCT_NO": product_no,
                "STATUS": "ERROR",
                "ERROR": err_msg,
                "SELECT_MAIN_IDX": sel_info.get("main", ""),
                "SELECT_ADD_IDX": ",".join([str(x) for x in (sel_info.get("additional", []) or [])]),
            })

        done += 1
        print(f"PROGRESS {done}/{total}")
        print(f"STATUS [{done}/{total}] {gs} → {status_msg}")

    # force log file creation (even if empty)
    out_dir = latest
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = log_path or os.path.join(out_dir, f"cafe24_upload_log_{ts}.xlsx")
    pd.DataFrame(log_rows).to_excel(out_path, index=False)
    print("[LOG]", out_path)


if __name__ == "__main__":
    main()





