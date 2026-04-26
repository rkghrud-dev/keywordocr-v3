from __future__ import annotations

import os


def _app_root() -> str:
    here = os.path.abspath(os.path.dirname(__file__))
    return os.path.abspath(os.path.join(here, "..", ".."))


def _desktop_key_dir() -> str:
    home = os.path.expanduser("~")
    return os.path.join(home, "Desktop", "key")


def _strip_quotes(value: str) -> str:
    value = value.strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in ("'", '"'):
        return value[1:-1]
    return value


def _load_env_file(path: str) -> bool:
    if not os.path.isfile(path):
        return False

    try:
        with open(path, "r", encoding="utf-8") as f:
            for raw in f:
                line = raw.strip()
                if not line or line.startswith("#"):
                    continue
                if "=" not in line:
                    continue
                key, value = line.split("=", 1)
                key = key.strip()
                if not key:
                    continue
                value = _strip_quotes(value)
                if key not in os.environ or not (os.environ.get(key) or "").strip():
                    os.environ[key] = value
    except Exception:
        return False
    return True


def _key_search_dirs() -> list[str]:
    key_dir = _desktop_key_dir()
    dirs = [key_dir]
    try:
        for name in os.listdir(key_dir):
            sub = os.path.join(key_dir, name)
            if os.path.isdir(sub):
                dirs.append(sub)
                try:
                    for name2 in os.listdir(sub):
                        sub2 = os.path.join(sub, name2)
                        if os.path.isdir(sub2):
                            dirs.append(sub2)
                except Exception:
                    pass
    except Exception:
        pass

    dedup: list[str] = []
    seen = set()
    for d in dirs:
        n = os.path.normcase(os.path.normpath(d))
        if n in seen:
            continue
        seen.add(n)
        dedup.append(d)
    return dedup


def key_file_candidates(filename: str) -> list[str]:
    candidates: list[str] = []
    for d in _key_search_dirs():
        candidates.append(os.path.join(d, filename))

    dedup: list[str] = []
    seen = set()
    for p in candidates:
        n = os.path.normcase(os.path.normpath(p))
        if n in seen:
            continue
        seen.add(n)
        dedup.append(p)
    return dedup

def _load_from_external_key_files() -> None:
    key_dir = _desktop_key_dir()
    if not os.path.isdir(key_dir):
        return

    if not (os.getenv("OPENAI_API_KEY") or "").strip():
        for p in key_file_candidates("api_key.txt"):
            if not os.path.isfile(p):
                continue
            try:
                with open(p, "r", encoding="utf-8") as f:
                    v = f.read().strip()
                if v:
                    os.environ["OPENAI_API_KEY"] = v
                    break
            except Exception:
                continue

    # Anthropic (Claude) API 키 로딩
    if not (os.getenv("ANTHROPIC_API_KEY") or "").strip():
        for p in key_file_candidates("anthropic_api_key.txt"):
            if not os.path.isfile(p):
                continue
            try:
                with open(p, "r", encoding="utf-8") as f:
                    v = f.read().strip()
                if v:
                    os.environ["ANTHROPIC_API_KEY"] = v
                    break
            except Exception:
                continue

    naver_map = {
        "ACCESS_LICENSE": "NAVER_ACCESS_LICENSE",
        "SECRET_KEY": "NAVER_SECRET_KEY",
        "CUSTOMER_ID": "NAVER_CUSTOMER_ID",
    }
    if any(not (os.getenv(env) or "").strip() for env in naver_map.values()):
        for p in key_file_candidates("naver_api_key.txt"):
            if not os.path.isfile(p):
                continue
            try:
                with open(p, "r", encoding="utf-8") as f:
                    for raw in f:
                        line = raw.strip()
                        if not line or line.startswith("#") or "=" not in line:
                            continue
                        k, v = line.split("=", 1)
                        kk = k.strip().upper()
                        vv = _strip_quotes(v)
                        env_name = naver_map.get(kk)
                        if env_name and not (os.getenv(env_name) or "").strip() and vv:
                            os.environ[env_name] = vv
            except Exception:
                continue

    shopping_map = {
        "CLIENT_ID": "NAVER_SHOPPING_CLIENT_ID",
        "CLIENT_SECRET": "NAVER_SHOPPING_CLIENT_SECRET",
    }
    if any(not (os.getenv(env) or "").strip() for env in shopping_map.values()):
        for p in key_file_candidates("naver_shopping_api_key.txt"):
            if not os.path.isfile(p):
                continue
            try:
                with open(p, "r", encoding="utf-8") as f:
                    for raw in f:
                        line = raw.strip()
                        if not line or line.startswith("#") or "=" not in line:
                            continue
                        k, v = line.split("=", 1)
                        kk = k.strip().upper()
                        vv = _strip_quotes(v)
                        env_name = shopping_map.get(kk)
                        if env_name and not (os.getenv(env_name) or "").strip() and vv:
                            os.environ[env_name] = vv
            except Exception:
                continue

    if not (os.getenv("GOOGLE_APPLICATION_CREDENTIALS") or "").strip():
        google_candidates = key_file_candidates("google_vision_key.json") + key_file_candidates("credentials.json")
        for p in google_candidates:
            if os.path.isfile(p):
                os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = p
                break


def ensure_env_loaded(dotenv_path: str | None = None) -> str | None:
    loaded_from = None
    candidates = [
        os.path.join(_desktop_key_dir(), "keywordocr.env"),
        os.path.join(_desktop_key_dir(), ".env"),
    ]

    seen = set()
    for path in candidates:
        n = os.path.normcase(os.path.normpath(path))
        if n in seen:
            continue
        seen.add(n)
        if _load_env_file(path) and loaded_from is None:
            loaded_from = path

    _load_from_external_key_files()
    return loaded_from


def get_env(*names: str) -> str:
    for name in names:
        value = (os.getenv(name) or "").strip()
        if value:
            return value
    return ""
