from __future__ import annotations

import argparse
import base64
import os
import urllib.parse

import requests


TOKEN_FILE = "cafe24_token.txt"
DEFAULT_SCOPE = "mall.read_order,mall.write_order,mall.read_shipping,mall.write_shipping,mall.read_product,mall.write_product"


def load_cfg(path: str) -> dict[str, str]:
    cfg = {
        "MALL_ID": "",
        "CLIENT_ID": "",
        "CLIENT_SECRET": "",
        "REDIRECT_URI": "",
        "SCOPE": "",
        "ACCESS_TOKEN": "",
        "REFRESH_TOKEN": "",
        "SHOP_NO": "1",
        "API_VERSION": "2025-12-01",
    }
    if not os.path.isfile(path):
        return cfg
    with open(path, "r", encoding="utf-8") as f:
        for raw in f:
            line = raw.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            k = k.strip().upper()
            v = v.strip().strip('"').strip("'")
            if k in cfg:
                cfg[k] = v
    return cfg


def save_cfg(path: str, cfg: dict[str, str]) -> None:
    keys = [
        "MALL_ID",
        "ACCESS_TOKEN",
        "REFRESH_TOKEN",
        "CLIENT_ID",
        "CLIENT_SECRET",
        "REDIRECT_URI",
        "SCOPE",
        "SHOP_NO",
        "API_VERSION",
    ]
    with open(path, "w", encoding="utf-8") as f:
        for k in keys:
            v = str(cfg.get(k, "") or "")
            if v:
                f.write(f"{k}={v}\n")


def build_authorize_url(cfg: dict[str, str], state: str = "keywordocr") -> str:
    base = f"https://{cfg['MALL_ID']}.cafe24api.com/api/v2/oauth/authorize"
    scope = (cfg.get("SCOPE") or "").strip() or DEFAULT_SCOPE
    params = {
        "response_type": "code",
        "client_id": cfg["CLIENT_ID"],
        "redirect_uri": cfg["REDIRECT_URI"],
        "scope": scope,
        "state": state,
    }
    return base + "?" + urllib.parse.urlencode(params)


def exchange_code(cfg: dict[str, str], code: str) -> dict:
    url = f"https://{cfg['MALL_ID']}.cafe24api.com/api/v2/oauth/token"
    auth = base64.b64encode(f"{cfg['CLIENT_ID']}:{cfg['CLIENT_SECRET']}".encode()).decode()
    headers = {
        "Authorization": f"Basic {auth}",
        "Content-Type": "application/x-www-form-urlencoded",
    }
    data = {
        "grant_type": "authorization_code",
        "code": code,
        "redirect_uri": cfg["REDIRECT_URI"],
    }
    r = requests.post(url, headers=headers, data=data, timeout=30)
    r.raise_for_status()
    return r.json()


def main() -> int:
    parser = argparse.ArgumentParser(description="Cafe24 OAuth helper")
    parser.add_argument("--token-file", default=TOKEN_FILE)
    parser.add_argument("--state", default="keywordocr")
    parser.add_argument("--code", default="", help="authorization code")
    args = parser.parse_args()

    cfg = load_cfg(args.token_file)
    need = ["MALL_ID", "CLIENT_ID", "CLIENT_SECRET", "REDIRECT_URI"]
    missing = [k for k in need if not cfg.get(k)]
    if missing:
        print("Missing required fields in token file:", ", ".join(missing))
        return 1

    if not args.code:
        print("[1] Open this URL and login/consent:")
        print(build_authorize_url(cfg, state=args.state))
        print("\n[2] Copy `code` from callback URL and run:")
        print("python cafe24_auth_setup.py --code YOUR_CODE")
        return 0

    try:
        res = exchange_code(cfg, args.code)
    except requests.HTTPError as e:
        body = e.response.text if e.response is not None else str(e)
        print("Token exchange failed:", body)
        return 2
    except Exception as e:
        print("Token exchange error:", e)
        return 2

    if res.get("access_token"):
        cfg["ACCESS_TOKEN"] = res["access_token"]
    if res.get("refresh_token"):
        cfg["REFRESH_TOKEN"] = res["refresh_token"]
    save_cfg(args.token_file, cfg)
    print("Saved ACCESS_TOKEN/REFRESH_TOKEN to", args.token_file)
    print("ACCESS_TOKEN len:", len(cfg.get("ACCESS_TOKEN", "")))
    print("REFRESH_TOKEN len:", len(cfg.get("REFRESH_TOKEN", "")))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
