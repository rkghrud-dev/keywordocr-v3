from __future__ import annotations

import requests
import base64


def refresh_token(mall_id: str, client_id: str, client_secret: str, refresh_token: str, redirect_uri: str = "") -> dict:
    url = f"https://{mall_id}.cafe24api.com/api/v2/oauth/token"
    data = {
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
    }
    if redirect_uri:
        data["redirect_uri"] = redirect_uri
    auth = base64.b64encode(f"{client_id}:{client_secret}".encode()).decode()
    headers = {"Authorization": f"Basic {auth}", "Content-Type": "application/x-www-form-urlencoded"}
    r = requests.post(url, headers=headers, data=data, timeout=30)
    r.raise_for_status()
    return r.json()
