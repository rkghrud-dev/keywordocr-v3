from app.services.keyword_builder import build_keyword_string


def _sample(name, ocr, vision):
    line = build_keyword_string(ocr_text=ocr, vision_analysis=vision, target_count=20, fallback_text=name)
    cnt = len([t for t in line.split() if t])
    print(f"[{name}] ({cnt}개)")
    print(line)
    print("-" * 80)


def main():
    samples = [
        (
            "차량 조명 브라켓",
            "차량용조명브라켓 무타공 본넷 트렁크 게이트 장착 각도조절 회전형 볼트체결 스틸 실버 블랙 작업등 외부조명 고정 DIY 튜닝",
            {
                "core_identity": {
                    "category": "차량조명 브라켓",
                    "product_type_correction": "차량용 마운트",
                    "structure": "회전형 볼트체결",
                    "material_visual": "스틸",
                    "color": ["실버", "블랙"],
                },
                "installation_and_physical": {"mount_type": "브라켓", "installation_method": ["무타공 설치", "볼트체결"]},
                "usage_context": {
                    "usage_location": ["본넷", "트렁크", "게이트"],
                    "usage_purpose": "작업등 장착",
                    "target_user": "DIY 튜닝 사용자",
                    "usage_scenario": "외부조명 고정",
                    "indoor_outdoor": "실외",
                },
                "functional_inference": {
                    "primary_function": "차량조명 거치",
                    "problem_solving_keyword": "흔들림 방지",
                    "convenience_feature": "각도조절 회전형",
                },
                "search_boost_elements": {
                    "installation_keywords": ["무타공", "장착", "설치"],
                    "space_keywords": ["본넷", "트렁크", "게이트"],
                    "benefit_keywords": ["고정", "진동완화"],
                    "longtail_candidates": ["차량 조명 브라켓 무타공", "트렁크 작업등 거치대"],
                },
            },
        ),
        (
            "트럭 D링",
            "트럭D링 고정고리 적재함 결속 앵커포인트 볼트체결 화물 스트랩 로프 체결",
            {
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
        ),
        (
            "플립 도어 캐치",
            "플립도어 캐치 잠금장치 도어락 스프링 래치 가구문 고정",
            {
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
        ),
        (
            "콘센트 가스켓",
            "콘센트가스켓 틈새밀폐패드 방수 방진 전기박스 커버 실링",
            {
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
        ),
        (
            "닭부리 힌지",
            "닭부리힌지 접이식 경첩 도어 힌지 고정 나사 체결 가구 철물",
            {
                "core_identity": {
                    "category": "닭부리 힌지",
                    "product_type_correction": "접이식 경첩",
                    "structure": "힌지 회전축",
                    "material_visual": "스틸",
                    "color": "실버",
                },
                "installation_and_physical": {"mount_type": "힌지", "installation_method": "나사 체결"},
                "usage_context": {
                    "usage_location": ["도어", "가구"],
                    "usage_purpose": "접이 구조 연결",
                    "target_user": "가구 제작 사용자",
                    "usage_scenario": "도어 여닫힘",
                    "indoor_outdoor": "실내",
                },
                "functional_inference": {
                    "primary_function": "회전 연결",
                    "problem_solving_keyword": "처짐 방지",
                    "convenience_feature": "부드러운 개폐",
                },
                "search_boost_elements": {
                    "installation_keywords": ["나사", "체결", "설치"],
                    "space_keywords": ["도어", "가구"],
                    "benefit_keywords": ["내구성", "안정고정"],
                    "longtail_candidates": ["닭부리 힌지 접이식 경첩", "도어 힌지 철물"],
                },
            },
        ),
        (
            "관개 커넥터",
            "관개커넥터 호스연결 조인트 누수방지 원예 급수 라인 체결",
            {
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
        ),
    ]

    for sample in samples:
        _sample(*sample)


if __name__ == "__main__":
    main()

