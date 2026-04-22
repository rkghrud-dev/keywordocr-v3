from __future__ import annotations

import argparse
import json
import os
import sys
import traceback
from pathlib import Path


if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')
os.environ.setdefault('PYTHONIOENCODING', 'utf-8')
os.environ.setdefault('PYTHONUTF8', '1')


def _status(message: str) -> None:
    print(message, flush=True)


def _to_bool(value: str) -> bool:
    return str(value).strip().lower() not in {'false', '0', 'n', 'no'}


def _normalize_google_credentials(legacy_root: Path) -> None:
    from app.services.env_loader import ensure_env_loaded, key_file_candidates

    ensure_env_loaded(str(legacy_root / '.env'))
    current = (os.getenv('GOOGLE_APPLICATION_CREDENTIALS') or os.getenv('GOOGLE_VISION_CREDENTIALS') or '').strip()
    if current and Path(current).is_file() and _is_desktop_key_path(current):
        os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = str(Path(current).resolve())
        return

    candidates = key_file_candidates('google_vision_key.json') + key_file_candidates('credentials.json')
    for candidate in candidates:
        if Path(candidate).is_file():
            os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = str(Path(candidate).resolve())
            return

    os.environ.pop('GOOGLE_APPLICATION_CREDENTIALS', None)


def _is_desktop_key_path(path: str) -> bool:
    key_root = os.path.normcase(os.path.abspath(Path.home() / 'Desktop' / 'key'))
    target = os.path.normcase(os.path.abspath(path))
    return target == key_root or target.startswith(key_root + os.sep)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument('--legacy-root', required=True)
    parser.add_argument('--source', required=True)
    parser.add_argument('--make-listing', default='true')
    parser.add_argument('--listing-size', default='1200')
    parser.add_argument('--listing-pad', default='20')
    parser.add_argument('--listing-max', default='20')
    parser.add_argument('--logo-path', default='')
    parser.add_argument('--logo-ratio', default='14')
    parser.add_argument('--logo-opacity', default='65')
    parser.add_argument('--logo-pos', default='tr')
    parser.add_argument('--use-auto-contrast', default='true')
    parser.add_argument('--use-sharpen', default='true')
    parser.add_argument('--use-small-rotate', default='true')
    parser.add_argument('--rotate-zoom', default='1.04')
    parser.add_argument('--ultra-angle-deg', default='0.35')
    parser.add_argument('--ultra-translate-px', default='0.6')
    parser.add_argument('--ultra-scale-pct', default='0.25')
    parser.add_argument('--trim-tol', default='8')
    parser.add_argument('--jpeg-q-min', default='88')
    parser.add_argument('--jpeg-q-max', default='92')
    parser.add_argument('--flip-lr', default='true')
    parser.add_argument('--logo-path-b', default='')
    parser.add_argument('--img-tag', default="<img src='https://gi.esmplus.com/rkghrud/상세1.jpg' />")
    parser.add_argument('--img-tag-b', default='')
    parser.add_argument('--a-name-min', default='80')
    parser.add_argument('--a-name-max', default='100')
    parser.add_argument('--b-name-min', default='63')
    parser.add_argument('--b-name-max', default='98')
    parser.add_argument('--a-tag-count', default='20')
    parser.add_argument('--b-tag-count', default='14')
    parser.add_argument('--phase', default='full', choices=['full', 'images', 'analysis', 'ocr_only'])
    parser.add_argument('--export-root', default='')
    parser.add_argument('--model', default='claude-sonnet-4-6')
    parser.add_argument('--chunk-size', default='10')
    parser.add_argument('--keyword-version', default='2.0')
    args = parser.parse_args()

    legacy_root = Path(args.legacy_root).resolve()
    os.chdir(legacy_root)
    sys.path.insert(0, str(legacy_root))

    _normalize_google_credentials(legacy_root)

    from app.services.pipeline import PipelineConfig, run_pipeline

    cfg = PipelineConfig(
        file_path=str(Path(args.source).resolve()),
        img_tag=str(args.img_tag or '').strip(),
        tesseract_path='',
        model_keyword=args.model,
        model_longtail=args.model,
        max_words=24,
        max_len=140,
        min_len=90,
        use_html_ocr=False,
        use_local_ocr=True,
        merge_ocr_with_name=True,
        max_imgs=999,
        threads=6,
        max_depth=-1,
        local_img_dir='',
        allow_folder_match=True,
        korean_only=False,
        drop_digits=True,
        psm=11,
        oem=3,
        ocr_excel_path='',
        write_to_r=True,
        debug=True,
        naver_enabled=False,
        naver_dry_run=False,
        naver_retry=False,
        naver_retry_count=2,
        naver_retry_delay=0.8,
        naver_autocomplete=False,
        google_autocomplete=True,
        make_listing=_to_bool(args.make_listing),
        listing_size=int(args.listing_size),
        listing_pad=int(args.listing_pad),
        listing_max=int(args.listing_max),
        logo_path=str(args.logo_path or '').strip(),
        logo_ratio=int(args.logo_ratio),
        logo_opacity=int(args.logo_opacity),
        logo_pos=str(args.logo_pos or 'tr').strip(),
        use_auto_contrast=_to_bool(args.use_auto_contrast),
        use_sharpen=_to_bool(args.use_sharpen),
        use_small_rotate=_to_bool(args.use_small_rotate),
        rotate_zoom=float(args.rotate_zoom),
        ultra_angle_deg=float(args.ultra_angle_deg),
        ultra_translate_px=float(args.ultra_translate_px),
        ultra_scale_pct=float(args.ultra_scale_pct),
        trim_tol=int(args.trim_tol),
        jpeg_q_min=int(args.jpeg_q_min),
        jpeg_q_max=int(args.jpeg_q_max),
        do_flip_lr=_to_bool(args.flip_lr),
        phase=args.phase,
        export_root_override=str(args.export_root or '').strip(),
        chunk_size=int(args.chunk_size),
        keyword_version=str(args.keyword_version or '2.0').strip() or '2.0',
        enable_b_market=True,
        logo_path_b=str(getattr(args, 'logo_path_b', '') or '').strip(),
        img_tag_b=str(getattr(args, 'img_tag_b', '') or '').strip(),
        a_name_min=int(args.a_name_min),
        a_name_max=int(args.a_name_max),
        b_name_min=int(args.b_name_min),
        b_name_max=int(args.b_name_max),
        a_tag_count=int(args.a_tag_count),
        b_tag_count=int(args.b_tag_count),
    )

    _status('작업 시작')
    out_root, out_file = run_pipeline(cfg, status_cb=_status, progress_cb=None)
    _status('처리 완료')
    _status(f'완료: {out_file}')
    print('__RESULT__' + json.dumps({'output_root': out_root, 'output_file': out_file}, ensure_ascii=False), flush=True)
    return 0


if __name__ == '__main__':
    try:
        raise SystemExit(main())
    except Exception:
        traceback.print_exc()
        raise SystemExit(1)
