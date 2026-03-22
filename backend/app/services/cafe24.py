from __future__ import annotations

import os
import re
import requests
from typing import Iterator

from app.services import cafe24_oauth
from app.services.env_loader import ensure_env_loaded, get_env


def _app_root() -> str:
    here = os.path.abspath(os.path.dirname(__file__))
    return os.path.abspath(os.path.join(here, '..', '..'))


def load_cafe24_config(path: str = "cafe24_token.txt") -> dict:
    cfg = {
        "MALL_ID": "",
        "ACCESS_TOKEN": "",
        "REFRESH_TOKEN": "",
        "CLIENT_ID": "",
        "CLIENT_SECRET": "",
        "REDIRECT_URI": "",
        "SHOP_NO": "1",
        "API_VERSION": "2025-12-01",
    }

    ensure_env_loaded()
    cfg["MALL_ID"] = get_env("CAFE24_MALL_ID", "MALL_ID")
    cfg["ACCESS_TOKEN"] = get_env("CAFE24_ACCESS_TOKEN", "ACCESS_TOKEN")
    cfg["REFRESH_TOKEN"] = get_env("CAFE24_REFRESH_TOKEN", "REFRESH_TOKEN")
    cfg["CLIENT_ID"] = get_env("CAFE24_CLIENT_ID", "CLIENT_ID")
    cfg["CLIENT_SECRET"] = get_env("CAFE24_CLIENT_SECRET", "CLIENT_SECRET")
    cfg["REDIRECT_URI"] = get_env("CAFE24_REDIRECT_URI", "REDIRECT_URI")
    cfg["SHOP_NO"] = get_env("CAFE24_SHOP_NO", "SHOP_NO") or cfg["SHOP_NO"]
    cfg["API_VERSION"] = get_env("CAFE24_API_VERSION", "API_VERSION") or cfg["API_VERSION"]

    candidates = [path, os.path.join(_app_root(), path)]
    file_path = next((p for p in candidates if os.path.isfile(p)), None)
    if not file_path:
        return cfg

    with open(file_path, "r", encoding="utf-8") as f:
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
    return cfg


def save_cafe24_config(path: str, cfg: dict) -> None:
    keys = ["MALL_ID", "ACCESS_TOKEN", "REFRESH_TOKEN", "CLIENT_ID", "CLIENT_SECRET", "REDIRECT_URI", "SHOP_NO", "API_VERSION"]
    with open(path, "w", encoding="utf-8") as f:
        for k in keys:
            if cfg.get(k, ""):
                f.write(f"{k}={cfg[k]}\\n")


def _headers(token: str, api_version: str) -> dict:
    return {
        "Authorization": f"Bearer {token}",
        "X-Cafe24-Api-Version": api_version,
    }


def _ensure_token(cfg_path: str, cfg: dict) -> dict:
    return cfg


def iter_products(mall_id: str, token: str, api_version: str, shop_no: str = "1",
                  limit: int = 100) -> Iterator[dict]:
    offset = 0
    while True:
        params = {"shop_no": shop_no, "limit": limit, "offset": offset}
        url = f"https://{mall_id}.cafe24api.com/api/v2/admin/products"
        r = requests.get(url, headers=_headers(token, api_version), params=params, timeout=30)
        if r.status_code == 401:
            raise RuntimeError("TOKEN_EXPIRED")
        r.raise_for_status()
        data = r.json()
        items = data.get("products", [])
        if not items:
            break
        for it in items:
            yield it
        if len(items) < limit:
            break
        offset += limit


def _read_image_base64(image_path: str) -> str:
    import base64
    with open(image_path, "rb") as f:
        return base64.b64encode(f.read()).decode()


def _json_headers(token: str, api_version: str) -> dict:
    return {
        "Authorization": f"Bearer {token}",
        "X-Cafe24-Api-Version": api_version,
        "Content-Type": "application/json",
    }


def upload_main_image(mall_id: str, token: str, api_version: str, product_no: int,
                      image_path: str, shop_no: str = "1") -> dict:
    url = f"https://{mall_id}.cafe24api.com/api/v2/admin/products/{product_no}/images"
    b64 = _read_image_base64(image_path)
    payload = {
        "request": {
            "detail_image": b64,
            "image_upload_type": "A",
        }
    }
    r = requests.post(url, headers=_json_headers(token, api_version), json=payload, timeout=60)
    if r.status_code == 401:
        raise RuntimeError("TOKEN_EXPIRED")
    r.raise_for_status()
    return r.json()


def upload_additional_image(mall_id: str, token: str, api_version: str, product_no: int,
                            image_path: str, shop_no: str = "1") -> dict:
    url = f"https://{mall_id}.cafe24api.com/api/v2/admin/products/{product_no}/additionalimages"
    b64 = _read_image_base64(image_path)
    payload = {
        "request": {
            "additional_image": [b64],
        }
    }
    r = requests.post(url, headers=_json_headers(token, api_version), json=payload, timeout=60)
    if r.status_code == 401:
        raise RuntimeError("TOKEN_EXPIRED")
    r.raise_for_status()
    return r.json()


def get_variants(mall_id: str, token: str, api_version: str, product_no: int,
                  shop_no: str = "1") -> list[dict]:
    url = f"https://{mall_id}.cafe24api.com/api/v2/admin/products/{product_no}/variants"
    params = {"shop_no": shop_no}
    r = requests.get(url, headers=_headers(token, api_version), params=params, timeout=30)
    if r.status_code == 401:
        raise RuntimeError("TOKEN_EXPIRED")
    r.raise_for_status()
    return r.json().get("variants", [])


def update_variant(mall_id: str, token: str, api_version: str, product_no: int,
                   variant_code: str, update_data: dict, shop_no: str = "1") -> dict:
    url = (f"https://{mall_id}.cafe24api.com/api/v2/admin/products/"
           f"{product_no}/variants/{variant_code}")
    payload = {"shop_no": int(shop_no), "request": update_data}
    r = requests.put(url, headers=_json_headers(token, api_version), json=payload, timeout=30)
    if r.status_code == 401:
        raise RuntimeError("TOKEN_EXPIRED")
    r.raise_for_status()
    return r.json()


def normalize_name(s: str) -> str:
    s = re.sub(r"[^0-9가-힣A-Za-z]", "", s or "")
    return s


def extract_gs_code(s: str) -> str | None:
    m = re.search(r"(GS\d{7})", s or "", flags=re.IGNORECASE)
    return m.group(1).upper() if m else None


def refresh_access_token(cfg_path: str, cfg: dict) -> dict:
    if not cfg.get("REFRESH_TOKEN") or not cfg.get("CLIENT_ID") or not cfg.get("CLIENT_SECRET"):
        raise RuntimeError("REFRESH_CREDENTIALS_MISSING")
    res = cafe24_oauth.refresh_token(
        cfg["MALL_ID"], cfg["CLIENT_ID"], cfg["CLIENT_SECRET"], cfg["REFRESH_TOKEN"], cfg.get("REDIRECT_URI", "")
    )
    cfg["ACCESS_TOKEN"] = res.get("access_token", cfg.get("ACCESS_TOKEN", ""))
    if res.get("refresh_token"):
        cfg["REFRESH_TOKEN"] = res.get("refresh_token")
    save_cafe24_config(cfg_path, cfg)
    return cfg
