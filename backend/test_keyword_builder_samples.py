import pytest

from app.services.keyword_builder import build_keyword_string


CASES = [
    {
        "name": "트럭 D링",
        "ocr": "트럭D링 고정고리 적재함 결속 앵커포인트 볼트체결 화물 스트랩 로프 체결",
        "vision": {
            "core_identity": {
                "category": "트럭 D링",
                "product_type_correction": "적재함 고정고리",
                "structure": "볼트 체결형",
                "material_visual": "철제",
                "color": "실버",
            },
            "installation_and_physical": {"mount_type": "앵커포인트", "installation_method": "볼트체결"},
            "usage_context": {
                "usage_location": "적재함",
                "usage_purpose": "화물 결속",
                "target_user": "화물차 사용자",
                "usage_scenario": "운송 고정",
                "indoor_outdoor": "실외",
            },
            "functional_inference": {
                "primary_function": "결속 고정",
                "problem_solving_keyword": "짐 흔들림 방지",
                "convenience_feature": "빠른 체결",
            },
            "search_boost_elements": {
                "installation_keywords": ["볼트", "체결", "고정"],
                "space_keywords": ["트럭", "적재함"],
                "benefit_keywords": ["안전고정", "내구성"],
                "longtail_candidates": ["트럭 적재함 D링 앵커포인트", "화물 결속 고정고리"],
            },
        },
        "forbidden": ["각도조절", "외부조명", "볼트고리", "나사고리"],
        "front": {"D링", "트럭", "적재함", "고정고리", "앵커포인트"},
    },
    {
        "name": "플립 도어 캐치",
        "ocr": "플립도어 캐치 잠금장치 도어락 스프링 래치 가구문 고정",
        "vision": {
            "core_identity": {
                "category": "플립 도어 캐치",
                "product_type_correction": "도어 잠금 래치",
                "structure": "스프링 래치",
                "material_visual": "스틸",
                "color": "실버",
            },
            "installation_and_physical": {"mount_type": "도어 캐치", "installation_method": "나사 체결"},
            "usage_context": {
                "usage_location": ["가구문", "수납장"],
                "usage_purpose": "문 닫힘 고정",
                "target_user": "가구 수리 사용자",
                "usage_scenario": "문 열림 방지",
                "indoor_outdoor": "실내",
            },
            "functional_inference": {
                "primary_function": "잠금 유지",
                "problem_solving_keyword": "문 흔들림 감소",
                "convenience_feature": "원터치 개폐",
            },
            "search_boost_elements": {
                "installation_keywords": ["나사", "체결", "설치"],
                "space_keywords": ["도어", "가구문"],
                "benefit_keywords": ["잠금", "고정"],
                "longtail_candidates": ["가구문 플립 도어 캐치", "스프링 래치 잠금장치"],
            },
        },
        "forbidden": ["각도조절", "외부조명", "볼트고리", "나사고리", "트럭"],
        "front": {"캐치", "도어", "플립", "래치", "잠금"},
    },
    {
        "name": "관개 커넥터",
        "ocr": "관개커넥터 호스연결 조인트 누수방지 원예 급수 라인 체결",
        "vision": {
            "core_identity": {
                "category": "관개 커넥터",
                "product_type_correction": "호스 연결 조인트",
                "structure": "원터치 체결형",
                "material_visual": "플라스틱",
                "color": "블랙",
            },
            "installation_and_physical": {"mount_type": "커넥터", "installation_method": "끼움 체결"},
            "usage_context": {
                "usage_location": ["급수 라인", "정원"],
                "usage_purpose": "호스 연결",
                "target_user": "원예 사용자",
                "usage_scenario": "관수 작업",
                "indoor_outdoor": "실외",
            },
            "functional_inference": {
                "primary_function": "관수 연결",
                "problem_solving_keyword": "누수 방지",
                "convenience_feature": "빠른 분리결합",
            },
            "search_boost_elements": {
                "installation_keywords": ["체결", "연결", "끼움"],
                "space_keywords": ["정원", "급수라인"],
                "benefit_keywords": ["누수방지", "작업효율"],
                "longtail_candidates": ["관개 호스 커넥터 조인트", "원예 급수 라인 연결"],
            },
        },
        "forbidden": ["각도조절", "외부조명", "볼트고리", "트럭"],
        "front": {"커넥터", "관개", "호스", "조인트", "급수"},
    },
    {
        "name": "콘센트 가스켓",
        "ocr": "콘센트가스켓 틈새밀폐패드 방수 방진 전기박스 커버 실링",
        "vision": {
            "core_identity": {
                "category": "콘센트 가스켓",
                "product_type_correction": "틈새 밀폐 패드",
                "structure": "실링 패킹형",
                "material_visual": "고무",
                "color": "블랙",
            },
            "installation_and_physical": {"mount_type": "패드", "installation_method": "부착 설치"},
            "usage_context": {
                "usage_location": ["콘센트", "전기박스"],
                "usage_purpose": "틈새 밀폐",
                "target_user": "실내 시공 사용자",
                "usage_scenario": "누수 차단",
                "indoor_outdoor": "실내외",
            },
            "functional_inference": {
                "primary_function": "방수 방진",
                "problem_solving_keyword": "먼지 유입 방지",
                "convenience_feature": "간편 부착",
            },
            "search_boost_elements": {
                "installation_keywords": ["부착", "설치", "실링"],
                "space_keywords": ["콘센트", "전기박스"],
                "benefit_keywords": ["밀폐", "누수방지"],
                "longtail_candidates": ["콘센트 틈새 밀폐 패드", "전기박스 가스켓 방수"],
            },
        },
        "forbidden": ["각도조절", "트럭", "D링"],
        "front": {"가스켓", "콘센트", "패드", "밀폐", "전기박스"},
    },
]


@pytest.mark.parametrize("case", CASES)
def test_build_keyword_string_blocks_cross_category_contamination(case) -> None:
    line = build_keyword_string(
        ocr_text=case["ocr"],
        vision_analysis=case["vision"],
        target_count=20,
        fallback_text=case["name"],
        market="A",
    )

    assert line, case["name"]
    tokens = [token for token in line.split() if token]
    assert tokens, line
    assert len(tokens) <= 20, line

    for forbidden in case["forbidden"]:
        assert forbidden not in line, line

    assert any(token in case["front"] for token in tokens[:3]), line
