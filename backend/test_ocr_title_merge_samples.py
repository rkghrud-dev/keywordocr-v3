from app.services import legacy_core as core
from app.services import ocr_noise_filter


def _show_case(name: str, base_name: str, kw_line: str, ocr_text: str, option_tokens: set[str]) -> None:
    merged = core.merge_base_name_with_keywords(
        base_name=base_name,
        kw_line=kw_line,
        max_words=16,
        max_len=120,
        option_tokens=option_tokens,
        ocr_text=ocr_text,
    )
    print(f"[{name}]")
    print(f"base:   {base_name}")
    print(f"merged: {merged}")
    print('-' * 80)


def main() -> None:
    keep = core.merge_base_name_with_keywords(
        base_name='3D 프린터 황동 노즐 부품 1.75mm 필라멘트 전용 0.4 GS0100629A',
        kw_line='3D프린터노즐 황동노즐 프린터부품 필라멘트전용 시제품출력 1.75mm 0.4',
        max_words=16,
        max_len=120,
        option_tokens={'0.4'},
        ocr_text='3D 프린터 황동 노즐 1.75mm 필라멘트 전용 교체 부품 0.4',
    )
    assert keep.startswith('3D 프린터 황동 노즐 부품'), keep
    assert 'GS0100629A' not in keep, keep

    expand = core.merge_base_name_with_keywords(
        base_name='노즐 0.4 GS0100629A',
        kw_line='3D프린터노즐 황동노즐 프린터부품 필라멘트전용 시제품출력 1.75mm 0.4',
        max_words=16,
        max_len=120,
        option_tokens={'0.4'},
        ocr_text='3D 프린터 황동 노즐 부품 1.75mm 필라멘트 전용 교체 출력용 0.4',
    )
    assert expand.startswith('3D프린터노즐 황동노즐'), expand
    assert '노즐 0.4' != expand[:6], expand
    assert expand.endswith('0.4'), expand

    hook_filtered = ocr_noise_filter.filter_ocr_text(
        """A. 블랙
OPTION
B. 화이트(미색) 후크 좌우 방향 조절 방법
먼저 나사를 풀어주세요.
잠금 고리를 빼세요.
방향을 바꿔서 다시 넣으세요.
나사를 다시 조여주면 완료입니다.
창문·방충망 잠금 후크
미닫이 창문·방충망에 설치 가능
무단 침입 방지·아이 보호·간편한 설치"""
    )
    hook_pre = ocr_noise_filter.preprocess_ocr_for_llm(hook_filtered)
    assert 'OPTION' not in hook_pre, hook_pre
    assert '좌우 방향' not in hook_pre, hook_pre
    assert '풀어주세요' not in hook_pre, hook_pre
    assert '다시 넣으세요' not in hook_pre, hook_pre
    assert '창문·방충망 잠금 후크' in hook_pre, hook_pre

    _show_case(
        'KEEP',
        '3D 프린터 황동 노즐 부품 1.75mm 필라멘트 전용 0.4 GS0100629A',
        '3D프린터노즐 황동노즐 프린터부품 필라멘트전용 시제품출력 1.75mm 0.4',
        '3D 프린터 황동 노즐 1.75mm 필라멘트 전용 교체 부품 0.4',
        {'0.4'},
    )
    _show_case(
        'EXPAND',
        '노즐 0.4 GS0100629A',
        '3D프린터노즐 황동노즐 프린터부품 필라멘트전용 시제품출력 1.75mm 0.4',
        '3D 프린터 황동 노즐 부품 1.75mm 필라멘트 전용 교체 출력용 0.4',
        {'0.4'},
    )


if __name__ == '__main__':
    main()
