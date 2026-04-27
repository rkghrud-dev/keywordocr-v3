"""
MarketPlus local helper server v3.

- Keeps the existing Naver shopping category proxy: http://localhost:5555/?q=상품명
- Adds a category-map upload GUI: http://localhost:5555/upload
- Looks up uploaded market category mappings by product name/code:
  http://localhost:5555/api/category-map?productName=...
"""
from __future__ import annotations

import base64
import difflib
import json
import os
import posixpath
import re
import time
import urllib.parse
import urllib.request
import zipfile
import xml.etree.ElementTree as ET
from http.server import BaseHTTPRequestHandler, HTTPServer
from io import BytesIO

NAVER_CLIENT_ID = os.environ.get("NAVER_CLIENT_ID", "")
NAVER_CLIENT_SECRET = os.environ.get("NAVER_CLIENT_SECRET", "")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STORE_PATH = os.path.join(BASE_DIR, "marketplus_category_map_store.json")
HELPER_JS_PATH = os.path.join(BASE_DIR, "marketplus-category-helper.user.js")
MAX_UPLOAD_BYTES = 20 * 1024 * 1024
HELPER_FEATURE_VERSION = 2

XML_NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def normalize_key(value: str) -> str:
    value = str(value or "").lower()
    value = re.sub(r"[a-z]{1,2}\d{5,}[a-z]?", "", value)
    value = re.sub(r"\d+(\.\d+)?\s*(cm|mm|m|g|kg|ml|l|개|매|장|ea)", "", value)
    value = re.sub(r"[^0-9a-z가-힣]+", "", value)
    return value.strip()


def normalize_code_key(value: str) -> str:
    value = str(value or "").lower()
    value = re.sub(r"[^0-9a-z가-힣]+", "", value)
    return value.strip()


def name_tokens(value: str) -> set[str]:
    value = str(value or "").lower()
    value = re.sub(r"[a-z]{1,2}\d{5,}[a-z]?", " ", value)
    tokens = set()
    for token in re.findall(r"[0-9a-z가-힣]+", value):
        if len(token) <= 1:
            continue
        if re.fullmatch(r"\d+", token):
            continue
        tokens.add(token)
    return tokens


def token_match_score(left: str, right: str) -> float:
    left_tokens = name_tokens(left)
    right_tokens = name_tokens(right)
    if not left_tokens or not right_tokens:
        return 0.0
    overlap = len(left_tokens & right_tokens)
    if overlap == 0:
        return 0.0
    containment = overlap / max(1, min(len(left_tokens), len(right_tokens)))
    jaccard = overlap / max(1, len(left_tokens | right_tokens))
    return max(jaccard * 0.94, containment * 0.78)


def to_float(value, default=0.0) -> float:
    try:
        return float(str(value).strip())
    except Exception:
        return default


def is_yes(value) -> bool:
    return str(value or "").strip().upper() in {"Y", "YES", "TRUE", "1"}


def col_to_idx(cell_ref: str) -> int | None:
    match = re.match(r"([A-Z]+)", cell_ref or "")
    if not match:
        return None
    n = 0
    for ch in match.group(1):
        n = n * 26 + ord(ch) - 64
    return n - 1


def cell_text(cell: ET.Element, shared: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(t.text or "" for t in cell.findall(".//a:t", XML_NS)).strip()

    value_node = cell.find("a:v", XML_NS)
    if value_node is None:
        return ""

    value = value_node.text or ""
    if cell_type == "s" and value:
        return shared[int(value)].strip()
    return value.strip()


def read_xlsx_rows(blob: bytes) -> dict[str, list[list[str]]]:
    sheets: dict[str, list[list[str]]] = {}
    with zipfile.ZipFile(BytesIO(blob)) as zf:
        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        relmap = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}

        shared: list[str] = []
        if "xl/sharedStrings.xml" in zf.namelist():
            root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in root.findall("a:si", XML_NS):
                shared.append("".join(t.text or "" for t in si.findall(".//a:t", XML_NS)))

        for sheet in workbook.findall("a:sheets/a:sheet", XML_NS):
            name = sheet.attrib["name"]
            rel_id = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
            target = relmap[rel_id]
            target = target.replace("\\", "/")
            if target.startswith("/"):
                sheet_path = target.lstrip("/")
            else:
                sheet_path = posixpath.normpath(posixpath.join("xl", target))
            root = ET.fromstring(zf.read(sheet_path))

            rows: list[list[str]] = []
            for row in root.findall("a:sheetData/a:row", XML_NS):
                sparse: dict[int, str] = {}
                max_idx = -1
                for cell in row.findall("a:c", XML_NS):
                    idx = col_to_idx(cell.attrib.get("r", ""))
                    if idx is None:
                        continue
                    sparse[idx] = cell_text(cell, shared)
                    max_idx = max(max_idx, idx)
                if max_idx >= 0:
                    rows.append([sparse.get(i, "") for i in range(max_idx + 1)])
            sheets[name] = rows
    return sheets


def rows_to_dicts(rows: list[list[str]]) -> list[dict[str, str]]:
    if not rows:
        return []
    headers = [str(x or "").strip() for x in rows[0]]
    result: list[dict[str, str]] = []
    for row in rows[1:]:
        item = {}
        for idx, header in enumerate(headers):
            if header:
                item[header] = row[idx] if idx < len(row) else ""
        if any(str(v).strip() for v in item.values()):
            result.append(item)
    return result


def make_record(
    row: dict[str, str],
    market: str,
    category_code: str = "",
    category_path: str = "",
    selector_json: str = "",
) -> dict:
    product_name = row.get("상품명", "")
    product_key = row.get("상품키", "") or row.get("상품코드", "")
    gs_code = row.get("GS코드", "")
    sku = row.get("SKU/자체상품코드", "") or row.get("SKU", "")
    confidence = to_float(row.get("confidence", row.get("카테고리확신도", row.get("확신도", ""))))
    review_needed = is_yes(row.get("review_needed", row.get("검수필요", "")))
    return {
        "productKey": product_key,
        "gsCode": gs_code,
        "sku": sku,
        "productName": product_name,
        "market": market,
        "categoryCode": str(category_code or "").strip(),
        "categoryPath": str(category_path or "").strip(),
        "selectorJson": str(selector_json or "").strip(),
        "confidence": confidence,
        "reviewNeeded": review_needed,
        "reason": row.get("reason", row.get("판단근거", row.get("매칭근거", ""))),
        "_productNameKey": normalize_key(product_name),
        "_productKeyKey": normalize_code_key(product_key),
        "_gsCodeKey": normalize_code_key(gs_code),
        "_skuKey": normalize_code_key(sku),
    }


def parse_category_workbook(blob: bytes, filename: str) -> dict:
    sheets = read_xlsx_rows(blob)
    records: list[dict] = []

    vertical_rows = rows_to_dicts(sheets.get("마켓별_세로형", []))
    if vertical_rows:
        for row in vertical_rows:
            records.append(make_record(
                row,
                row.get("market", ""),
                row.get("category_code", ""),
                row.get("category_path", ""),
                row.get("selector_json", ""),
            ))
    else:
        wide_rows = rows_to_dicts(sheets.get("상품_카테고리맵", []))
        if not wide_rows:
            for rows in sheets.values():
                candidate_rows = rows_to_dicts(rows)
                if candidate_rows and any(("상품명" in r and ("네이버카테고리코드" in r or "쿠팡카테고리코드" in r)) for r in candidate_rows):
                    wide_rows = candidate_rows
                    break
        for row in wide_rows:
            lotte_code = row.get("롯데ON전시카테고리코드", "") or row.get("롯데ON표준카테고리코드", "")
            lotte_path = row.get("롯데ON전시카테고리경로", "") or row.get("롯데ON표준카테고리경로", "")
            mappings = [
                ("naver", "네이버카테고리코드", "네이버카테고리경로", ""),
                ("coupang", "쿠팡카테고리코드", "쿠팡카테고리경로", ""),
                ("auction", "옥션카테고리코드", "옥션카테고리경로", "옥션드롭다운값"),
                ("11st", "11번가카테고리코드", "11번가카테고리경로", ""),
                ("smartstore", "스마트스토어카테고리코드", "스마트스토어카테고리경로", "스마트스토어드롭다운값"),
            ]
            for market, code_col, path_col, selector_col in mappings:
                code = row.get(code_col, "") if code_col else ""
                path = row.get(path_col, "")
                selector = row.get(selector_col, "") if selector_col else ""
                if code or path or selector:
                    records.append(make_record(row, market, code, path, selector))
            gmarket_path = row.get("G마켓카테고리경로", "") or row.get("ESM카테고리경로", "")
            gmarket_selector = row.get("G마켓드롭다운값", "")
            if gmarket_path or gmarket_selector:
                records.append(make_record(row, "gmarket", "", gmarket_path, gmarket_selector))
            if lotte_code or lotte_path:
                records.append(make_record(row, "lotteon", lotte_code, lotte_path, ""))

    learning_rows = rows_to_dicts(sheets.get("드롭다운값_학습", []))
    review_rows = rows_to_dicts(sheets.get("검수필요", []))
    product_keys = sorted({r["productKey"] or r["gsCode"] or r["productName"] for r in records if r["productName"]})

    return {
        "version": 1,
        "filename": filename,
        "uploadedAt": time.strftime("%Y-%m-%d %H:%M:%S"),
        "recordCount": len(records),
        "productCount": len(product_keys),
        "records": records,
        "learningRows": learning_rows,
        "reviewRows": review_rows,
    }


def alias_get(alias: dict, *names: str) -> str:
    for name in names:
        if name in alias and alias.get(name) is not None:
            return str(alias.get(name) or "").strip()
    return ""


def refresh_record_keys(record: dict) -> None:
    record["_productNameKey"] = normalize_key(record.get("productName", ""))
    record["_productKeyKey"] = normalize_code_key(record.get("productKey", ""))
    record["_gsCodeKey"] = normalize_code_key(record.get("gsCode", ""))
    record["_skuKey"] = normalize_code_key(record.get("sku", ""))


def apply_aliases_to_store(store: dict, aliases: list[dict] | None) -> int:
    if not aliases:
        store["aliasCount"] = 0
        return 0

    records = store.get("records") or []
    added = 0
    seen = {
        (
            str(r.get("market", "")).lower(),
            normalize_code_key(r.get("productKey", "") or r.get("gsCode", "") or r.get("sku", "")),
            normalize_key(r.get("productName", "")),
        )
        for r in records
    }

    for alias in aliases:
        alias_name = alias_get(alias, "productName", "ProductName", "name", "Name")
        if not alias_name:
            continue
        alias_code = alias_get(alias, "productKey", "ProductKey", "gsCode", "GsCode", "sku", "Sku", "code", "Code")
        alias_code_key = normalize_code_key(alias_code)
        alias_name_key = normalize_key(alias_name)

        matched_records = []
        for record in records:
            code_keys = {
                record.get("_productKeyKey", ""),
                record.get("_gsCodeKey", ""),
                record.get("_skuKey", ""),
            }
            if alias_code_key and alias_code_key in code_keys:
                matched_records.append(record)
            elif alias_name_key and alias_name_key == record.get("_productNameKey", ""):
                matched_records.append(record)

        for record in list(matched_records):
            market = str(record.get("market", "")).lower()
            group_code = normalize_code_key(record.get("productKey", "") or record.get("gsCode", "") or record.get("sku", "") or alias_code)
            key = (market, group_code, alias_name_key)
            if key in seen:
                continue

            cloned = dict(record)
            cloned["productName"] = alias_name
            if alias_code and not cloned.get("productKey"):
                cloned["productKey"] = alias_code
            if alias_code and not cloned.get("gsCode"):
                cloned["gsCode"] = alias_code
            refresh_record_keys(cloned)
            records.append(cloned)
            seen.add(key)
            added += 1

    store["records"] = records
    store["recordCount"] = len(records)
    store["productCount"] = len({
        r.get("productKey") or r.get("gsCode") or r.get("sku") or r.get("productName")
        for r in records
        if r.get("productName")
    })
    store["aliasCount"] = added
    return added


def save_store(store: dict) -> None:
    with open(STORE_PATH, "w", encoding="utf-8") as f:
        json.dump(store, f, ensure_ascii=False, indent=2)


def load_store() -> dict:
    if not os.path.isfile(STORE_PATH):
        return {}
    try:
        with open(STORE_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def product_match_score(records: list[dict], product_name: str, product_code: str) -> float:
    name_key = normalize_key(product_name)
    code_key = normalize_code_key(product_code)
    best = 0.0
    for record in records:
        code_keys = [record.get("_productKeyKey", ""), record.get("_gsCodeKey", ""), record.get("_skuKey", "")]
        name_candidate = record.get("_productNameKey", "")

        if code_key and code_key in code_keys:
            best = max(best, 1.0)
        if name_key and name_candidate:
            if name_key == name_candidate:
                best = max(best, 0.98)
            elif name_key in name_candidate or name_candidate in name_key:
                best = max(best, 0.90)
            else:
                best = max(best, difflib.SequenceMatcher(None, name_key, name_candidate).ratio() * 0.86)
                best = max(best, token_match_score(product_name, record.get("productName", "")))
    return best


def lookup_category_map(product_name: str, product_code: str = "") -> dict:
    store = load_store()
    records = store.get("records") or []
    if not records:
        return {"matched": False, "error": "category map not uploaded"}

    groups: dict[str, list[dict]] = {}
    for record in records:
        group_key = record.get("productKey") or record.get("gsCode") or record.get("sku") or record.get("productName")
        if group_key:
            groups.setdefault(group_key, []).append(record)

    best_key = ""
    best_score = 0.0
    for group_key, group_records in groups.items():
        score = product_match_score(group_records, product_name, product_code)
        if score > best_score:
            best_key = group_key
            best_score = score

    if not best_key or best_score < 0.58:
        return {"matched": False, "score": round(best_score, 4), "error": "no matching product"}

    group_records = groups[best_key]
    markets = {}
    for record in group_records:
        market = str(record.get("market", "")).strip().lower()
        if not market:
            continue
        markets[market] = {
            "categoryCode": record.get("categoryCode", ""),
            "categoryPath": record.get("categoryPath", ""),
            "selectorJson": record.get("selectorJson", ""),
            "confidence": record.get("confidence", 0),
            "reviewNeeded": record.get("reviewNeeded", False),
            "reason": record.get("reason", ""),
        }

    first = max(group_records, key=lambda r: product_match_score([r], product_name, product_code))
    confidence_values = [to_float(r.get("confidence")) for r in group_records if to_float(r.get("confidence")) > 0]
    confidence = min(confidence_values) if confidence_values else 0
    return {
        "matched": True,
        "score": round(best_score, 4),
        "productKey": first.get("productKey", ""),
        "gsCode": first.get("gsCode", ""),
        "sku": first.get("sku", ""),
        "productName": first.get("productName", ""),
        "confidence": confidence,
        "reviewNeeded": any(bool(r.get("reviewNeeded")) for r in group_records),
        "markets": markets,
        "source": {
            "filename": store.get("filename", ""),
            "uploadedAt": store.get("uploadedAt", ""),
            "recordCount": store.get("recordCount", 0),
            "productCount": store.get("productCount", 0),
        },
    }


UPLOAD_HTML = """<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <title>MarketPlus Category Map Upload</title>
  <style>
    body{font-family:Malgun Gothic,Arial,sans-serif;margin:32px;color:#1f2328}
    main{max-width:760px;margin:0 auto}
    h1{font-size:22px;margin:0 0 18px}
    .box{border:1px solid #d0d7de;border-radius:8px;padding:18px;margin:14px 0}
    input,button{font-size:14px}
    button{background:#2c6fbb;color:#fff;border:0;border-radius:6px;padding:9px 14px;cursor:pointer}
    pre{background:#f6f8fa;border-radius:6px;padding:12px;white-space:pre-wrap;line-height:1.5}
    .muted{color:#57606a;font-size:13px}
  </style>
</head>
<body>
<main>
  <h1>MarketPlus Category Map Upload</h1>
  <div class="box">
    <input id="file" type="file" accept=".xlsx">
    <button id="upload">업로드</button>
    <p class="muted">마켓별_카테고리맵 xlsx를 올리면 registerall 헬퍼가 상품명으로 조회해서 드롭다운값을 적용합니다.</p>
  </div>
  <div class="box">
    <strong>현재 상태</strong>
    <pre id="status">Loading...</pre>
  </div>
</main>
<script>
async function refreshStatus(){
  const res = await fetch('/api/map/status');
  document.getElementById('status').textContent = JSON.stringify(await res.json(), null, 2);
}
document.getElementById('upload').addEventListener('click', async () => {
  const file = document.getElementById('file').files[0];
  if(!file){ alert('xlsx 파일을 선택하세요.'); return; }
  const reader = new FileReader();
  reader.onload = async () => {
    const base64 = String(reader.result).split(',')[1];
    const res = await fetch('/api/upload-map', {
      method: 'POST',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify({filename:file.name, contentBase64:base64})
    });
    const data = await res.json();
    document.getElementById('status').textContent = JSON.stringify(data, null, 2);
  };
  reader.readAsDataURL(file);
});
refreshStatus();
</script>
</body>
</html>"""


class Handler(BaseHTTPRequestHandler):
    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "*")
        self.end_headers()

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        params = urllib.parse.parse_qs(parsed.query)

        if parsed.path == "/upload":
            self.respond_html(UPLOAD_HTML)
            return

        if parsed.path == "/helper.js":
            try:
                with open(HELPER_JS_PATH, "rb") as f:
                    self.respond_bytes(f.read(), "application/javascript; charset=utf-8")
            except Exception as e:
                self.respond({"ok": False, "error": str(e)}, status=404)
            return

        if parsed.path == "/api/map/status":
            store = load_store()
            self.respond({
                "helperVersion": HELPER_FEATURE_VERSION,
                "supportsAliases": True,
                "uploaded": bool(store.get("records")),
                "filename": store.get("filename", ""),
                "uploadedAt": store.get("uploadedAt", ""),
                "recordCount": store.get("recordCount", 0),
                "productCount": store.get("productCount", 0),
                "aliasCount": store.get("aliasCount", 0),
                "storePath": STORE_PATH,
            })
            return

        if parsed.path == "/api/category-map":
            product_name = params.get("productName", [""])[0]
            product_code = params.get("productCode", [""])[0]
            self.respond(lookup_category_map(product_name, product_code))
            return

        query = params.get("q", [""])[0]
        if query:
            try:
                result = self.fetch_naver_category(query)
                print(f'  [naver:{query}] -> {result.get("fullPath", "?")}')
                self.respond(result)
            except Exception as e:
                print(f"  ERROR: {e}")
                self.respond({"error": str(e), "category": "", "levels": []})
            return

        self.respond({
            "ok": True,
            "name": "MarketPlus local helper server",
            "upload": "http://localhost:5555/upload",
            "naverProxy": "http://localhost:5555/?q=상품명",
        })

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path != "/api/upload-map":
            self.send_error(404)
            return

        try:
            length = int(self.headers.get("Content-Length", "0"))
            if length <= 0 or length > MAX_UPLOAD_BYTES:
                self.respond({"ok": False, "error": "invalid upload size"}, status=400)
                return

            body = self.rfile.read(length)
            payload = json.loads(body.decode("utf-8"))
            filename = payload.get("filename") or "category_map.xlsx"
            content_b64 = payload.get("contentBase64") or ""
            blob = base64.b64decode(content_b64)
            store = parse_category_workbook(blob, filename)
            alias_count = apply_aliases_to_store(store, payload.get("aliases") or [])
            save_store(store)
            print(f'  [map-upload] {filename}: {store["productCount"]} products, {store["recordCount"]} records, {alias_count} aliases')
            self.respond({
                "ok": True,
                "filename": filename,
                "uploadedAt": store["uploadedAt"],
                "recordCount": store["recordCount"],
                "productCount": store["productCount"],
                "aliasCount": alias_count,
                "storePath": STORE_PATH,
            })
        except Exception as e:
            self.respond({"ok": False, "error": str(e)}, status=500)

    def fetch_naver_category(self, query: str) -> dict:
        if not NAVER_CLIENT_ID or not NAVER_CLIENT_SECRET:
            return {"category": "", "levels": [], "error": "NAVER_CLIENT_ID/NAVER_CLIENT_SECRET not configured"}

        url = "https://openapi.naver.com/v1/search/shop.json?" + urllib.parse.urlencode({
            "query": query,
            "display": 30,
            "sort": "sim",
        })

        req = urllib.request.Request(url)
        req.add_header("X-Naver-Client-Id", NAVER_CLIENT_ID)
        req.add_header("X-Naver-Client-Secret", NAVER_CLIENT_SECRET)

        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read().decode("utf-8"))

        items = data.get("items", [])
        if not items:
            return {"category": "", "levels": [], "error": "no results"}

        path_freq: dict[str, int] = {}
        for item in items:
            parts = [item.get(k, "") for k in ("category1", "category2", "category3", "category4")]
            path_key = " > ".join(x for x in parts if x)
            if path_key:
                path_freq[path_key] = path_freq.get(path_key, 0) + 1

        if not path_freq:
            return {"category": "", "levels": [], "error": "no category"}

        top_path = max(path_freq, key=path_freq.get)
        levels = top_path.split(" > ")
        return {
            "category": levels[-1] if levels else "",
            "fullPath": top_path,
            "levels": levels,
            "topLevel": levels[0] if levels else "",
            "all": [{"path": k, "count": v} for k, v in sorted(path_freq.items(), key=lambda x: -x[1])[:5]],
        }

    def respond(self, data: dict, status: int = 200):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def respond_html(self, html: str):
        body = html.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def respond_bytes(self, body: bytes, content_type: str):
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, format, *args):
        pass


if __name__ == "__main__":
    port = 5555
    print("=" * 56)
    print("  MarketPlus local helper server v3")
    print(f"  Upload UI : http://localhost:{port}/upload")
    print(f"  Naver API : http://localhost:{port}/?q=product")
    print("  Ctrl+C to stop")
    print("=" * 56)
    HTTPServer(("127.0.0.1", port), Handler).serve_forever()
