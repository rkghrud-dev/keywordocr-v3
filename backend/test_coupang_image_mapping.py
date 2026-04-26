import json
import os
import tempfile
import unittest

from app.services.coupang import (
    _IMAGE_SELECTION_CACHE,
    _build_fallback_image_urls,
    _pick_local_listing_images,
    build_option_attributes,
    build_coupang_product,
    parse_options,
)


class CoupangImageMappingTests(unittest.TestCase):
    def test_explicit_image_selection_does_not_append_remaining_listing_images(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            export_root = tmpdir
            image_dir = os.path.join(export_root, "listing_images", "20260402", "GS1700135")
            os.makedirs(image_dir, exist_ok=True)
            for idx in range(1, 6):
                with open(os.path.join(image_dir, f"GS1700135_{idx}.jpg"), "wb") as f:
                    f.write(b"x")

            with open(os.path.join(export_root, "image_selections.json"), "w", encoding="utf-8-sig") as f:
                json.dump({"GS1700135": {"main": 1, "additional": [3]}}, f)

            _IMAGE_SELECTION_CACHE.clear()
            row = {
                "_source_file_path": os.path.join(export_root, "llm_result", "chunk.xlsx"),
                "_export_root": export_root,
                "자체 상품코드": "GS1700135A",
            }

            picked = _pick_local_listing_images(row)

            self.assertEqual([os.path.basename(path) for path in picked], ["GS1700135_2.jpg", "GS1700135_4.jpg"])

    def test_fallback_images_ignore_detail_html_images(self):
        row = {
            "이미지등록(목록)": "http://example.com/list.jpg",
            "이미지등록(상세)": "http://example.com/detail-view.jpg",
            "이미지등록(추가)": "",
            "자체 상품코드": "GS1700135A",
        }
        detail_html = (
            '<center><img src="http://example.com/body-1.jpg">'
            '<img src="http://example.com/body-2.jpg"></center>'
        )

        image_urls = _build_fallback_image_urls(row, detail_html)

        self.assertEqual(image_urls, ["http://example.com/list.jpg"])

    def test_product_contents_use_text_type(self):
        row = {
            "상품명": "피규어 고정 테이프",
            "판매가": 1800,
            "소비자가": 2200,
            "이미지등록(목록)": "http://example.com/list.jpg",
            "상품 상세설명": '<center><img src="http://example.com/body-1.jpg"></center>',
            "검색어설정": "피규어고정테이프",
            "자체 상품코드": "GS1700135A",
        }

        product = build_coupang_product(row, 50007029, {"data": {"attributes": [], "noticeCategories": []}})
        contents = product["items"][0]["contents"][0]

        self.assertEqual(contents["contentsType"], "TEXT")
        self.assertEqual(contents["contentDetails"][0]["detailType"], "TEXT")

    def test_parse_options_supports_pipe_delimiter(self):
        options = parse_options("옵션{A 투명싱글|B 투명더블|C 스텐싱글|D 스텐더블}", "0|500|0|500")

        self.assertEqual(
            options,
            [
                {"name": "투명싱글", "price": 0},
                {"name": "투명더블", "price": 500},
                {"name": "스텐싱글", "price": 0},
                {"name": "스텐더블", "price": 500},
            ],
        )

    def test_build_option_attributes_maps_color_and_quantity(self):
        category_meta = {
            "data": {
                "attributes": [
                    {"attributeTypeName": "색상", "exposed": "EXPOSED", "basicUnit": "없음"},
                    {"attributeTypeName": "수량", "exposed": "EXPOSED", "basicUnit": "개"},
                ]
            }
        }

        attrs = build_option_attributes("스텐더블", category_meta)

        self.assertEqual(
            attrs,
            [
                {"attributeTypeName": "색상", "attributeValueName": "스텐"},
                {"attributeTypeName": "수량", "attributeValueName": "2개"},
            ],
        )


if __name__ == "__main__":
    unittest.main()
