from __future__ import annotations

import os
import re
import requests

from .env_loader import ensure_env_loaded, get_env, key_file_candidates


def _app_root() -> str:
    here = os.path.abspath(os.path.dirname(__file__))
    return os.path.abspath(os.path.join(here, '..', '..'))


def load_shopping_keys(path: str = "naver_shopping_api_key.txt") -> dict:
    keys = {
        "CLIENT_ID": "",
        "CLIENT_SECRET": "",
    }

    ensure_env_loaded()
    keys["CLIENT_ID"] = get_env("NAVER_SHOPPING_CLIENT_ID", "NAVER_CLIENT_ID")
    keys["CLIENT_SECRET"] = get_env("NAVER_SHOPPING_CLIENT_SECRET", "NAVER_CLIENT_SECRET")

    candidates = key_file_candidates(path)
    file_path = next((p for p in candidates if os.path.isfile(p)), None)
    if not file_path:
        return keys

    with open(file_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            if "=" in line:
                k, v = line.split("=", 1)
                k = k.strip().upper()
                v = v.strip().strip('"').strip("'")
                if k in keys and not keys[k]:
                    keys[k] = v
    return keys


def _strip_html(text: str) -> str:
    text = re.sub(r"<[^>]+>", " ", text or "")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def naver_shopping_titles(query: str,
                          client_id: str,
                          client_secret: str,
                          pages: int = 3,
                          display: int = 40,
                          timeout: int = 10) -> list[str]:
    if not query:
        return []
    if not client_id or not client_secret:
        raise RuntimeError("네이버 쇼핑 API 키가 없습니다. .env(NAVER_SHOPPING_CLIENT_ID/NAVER_SHOPPING_CLIENT_SECRET) 또는 naver_shopping_api_key.txt를 확인하세요.")

    headers = {
        "X-Naver-Client-Id": client_id,
        "X-Naver-Client-Secret": client_secret,
    }

    titles: list[str] = []
    pages = max(1, min(5, int(pages)))
    display = max(1, min(100, int(display)))

    for p in range(pages):
        start = 1 + p * display
        if start > 1000:
            break
        params = {
            "query": query,
            "display": display,
            "start": start,
            "sort": "sim",
        }
        r = requests.get("https://openapi.naver.com/v1/search/shopping.json",
                         headers=headers, params=params, timeout=timeout)
        r.raise_for_status()
        data = r.json()
        for item in data.get("items", []):
            title = _strip_html(item.get("title", ""))
            if title:
                titles.append(title)
    return titles

