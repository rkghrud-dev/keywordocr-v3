"""
C# → Python 쿠팡 업로드 bridge
프로토콜: stdout 으로 진행상황 출력, __RESULT__ 접두사로 JSON 결과 전달
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path


def _status(msg: str) -> None:
    print(msg, flush=True)


def main() -> int:
    parser = argparse.ArgumentParser(description="쿠팡 업로드 bridge")
    parser.add_argument("--legacy-root", required=True, help="Python 백엔드 루트")
    parser.add_argument("--source", required=True, help="업로드용 엑셀 파일 경로")
    parser.add_argument("--row-start", type=int, default=0, help="시작 행 (0=전체)")
    parser.add_argument("--row-end", type=int, default=0, help="끝 행 (0=시작행과 동일)")
    parser.add_argument("--dry-run", default="true", help="DRY RUN 여부")
    args = parser.parse_args()

    legacy_root = Path(args.legacy_root).resolve()
    os.chdir(str(legacy_root))
    sys.path.insert(0, str(legacy_root))

    _status("쿠팡 업로드 bridge 시작")

    from app.services.coupang import CoupangUploadConfig, run_coupang_upload

    dry_run = args.dry_run.lower() in ("true", "1", "yes")

    cfg = CoupangUploadConfig(
        file_path=str(Path(args.source).resolve()),
        row_start=args.row_start,
        row_end=args.row_end,
        dry_run=dry_run,
    )

    results = run_coupang_upload(cfg, status_cb=_status, progress_cb=None)

    result_list = []
    success_count = 0
    fail_count = 0
    for r in results:
        entry = {
            "row": r.row,
            "name": r.name,
            "status": r.status,
            "category": r.category,
            "seller_product_id": r.seller_product_id,
            "error": r.error,
        }
        result_list.append(entry)
        if r.status in ("SUCCESS", "DRY_RUN"):
            success_count += 1
        else:
            fail_count += 1
            _status(f"  [실패] 행{r.row} {r.name[:30]} → {r.status}: {r.error[:100]}")

    payload = {
        "results": result_list,
        "success_count": success_count,
        "fail_count": fail_count,
        "total_count": len(result_list),
    }

    _status(f"완료: 성공 {success_count} / 실패 {fail_count} / 전체 {len(result_list)}")
    print("__RESULT__" + json.dumps(payload, ensure_ascii=False), flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
