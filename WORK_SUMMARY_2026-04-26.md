# 작업 정리 - 2026-04-26

## 현재 목표

KeywordOCR v3를 실제 사용 흐름 중심으로 정리하고, 키워드 생성 품질을 한 단계 개선했다.

실사용 흐름:

1. 초기세팅에서 마켓별 로고 PNG, Cafe24 토큰 JSON 등 한 번만 설정
2. 기본실행에서 CSV 파일 선택
3. 1차 가공 실행
4. 대표이미지 선택 화면 자동 진입
5. PowerShell/Codex LLM 결과 생성 후 `LLM 결과 불러오기`
6. Cafe24 신규등록

## UI 변경

변경 파일:

- `KeywordOcr.App/MainWindow.xaml`
- `KeywordOcr.App/MainWindow.xaml.cs`

반영 내용:

- 메인 탭 이름을 `작업흐름`으로 변경
- 상단에 작업 순서 표시 추가
- 기본값 변경
  - Codex 분할 단위: `분할안함`
  - 키워드 버전: `3.0`
- 분할 설정, OCR 제외 같은 드문 설정은 `고급 옵션` 접힘 영역으로 이동
- 자주 쓰지 않는 화면은 메뉴로 숨김
  - 자동실행
  - 실행 이력
  - Cafe24 업로드
  - 쿠팡 업로드
  - 옵션 가격
- `초기세팅` 탭은 유지해서 로고/토큰/JSON 설정에 사용
- 1차 가공 후 대표이미지 화면이 열리도록 유지
- 대표이미지 저장 후 다시 `작업흐름` 탭으로 복귀
- LLM 결과를 불러오면 Cafe24 신규등록 대상 목록을 자동 로드
- `Cafe24 이미지+가격 업로드`, 쿠팡/네이버 마켓 업로드 영역은 기본 흐름에서 숨김

## 키워드 로직 변경

변경 파일:

- `backend/app/services/market_keywords.py`
- `backend/app/services/pipeline.py`
- `backend/test_market_keywords_b_split.py`

반영 내용:

- OCR 숫자 노이즈 필터 추가
  - 제거 예: `801`, `2024`, `19900원`, `2O2`
  - 유지 예: `35mm`, `M8`, `12V`, `500ml`
- 광고/과장형 금지어 확장
  - `프리미엄`, `고품질`, `최고급`, `베스트`, `핫딜`, `가성비` 등
- 사용처 힌트 확장
  - 욕실, 화장실, 주방, 싱크대, 배수구, 선반, 옷장, 서랍장, 책상 등
- 상품 정체성 힌트 확장
  - 클램프, 볼트, 너트, 나사, 핀, 밴드, 테이프, 커버, 노즐, 롤러 등
- 포함 중복 제거 추가
  - `볼트`와 `가구연결볼트`가 같이 들어가면 더 구체적인 조합을 우선
- 3.0 LLM 지시서에 저장 전 자체검사 규칙 추가
  - GS코드/가격/배송 문구 제거
  - 단위 없는 숫자 파편 제거
  - A/B 제목의 뒷부분 토큰 50% 이상 중복 방지
  - 검색어설정에서 공백변형/재조합/동의어 중복 제거
  - 근거 약한 태그로 목표 개수 채우지 않기

## 검증 결과

실행한 검증:

```powershell
dotnet build C:\Users\rkghr\Desktop\프로젝트\keywordocr-v3\KeywordOcr.App\KeywordOcr.App.csproj
```

결과:

- 성공
- 경고 0개
- 오류 0개

```powershell
py -3 -m py_compile backend/app/services/market_keywords.py backend/app/services/pipeline.py
```

결과:

- 성공

`pytest`는 현재 Python 환경에 설치되어 있지 않아 직접 테스트 함수 호출 방식으로 검증했다.

검증한 테스트:

- `backend/test_market_keywords_b_split.py`
- `backend/test_keyword_builder_samples.py`

추가한 테스트:

- OCR 숫자 노이즈 제거, 규격 숫자 유지
- 포함 중복에서 더 구체적인 조합 우선

직접 호출 결과:

- 모든 테스트 통과

배포 빌드:

```powershell
& C:\Users\rkghr\Desktop\프로젝트\keywordocr-v3\build-dist.ps1
```

결과:

- 성공
- `dist\KeywordOcr.exe` 갱신됨

실행 확인:

- 프로세스: `KeywordOcr`
- PID: `28744`
- 창 제목: `KeywordOCR v3`
- 상태: 응답 중

## 현재 Git 상태 참고

브랜치:

- `master`

원격:

- `origin https://github.com/rkghrud-dev/keywordocr-v3.git`

주의:

- 작업 시작 전부터 이미 변경/미추적 파일이 많이 있었다.
- 이번 작업에서 직접 수정한 핵심 파일은 위 `UI 변경`, `키워드 로직 변경` 섹션의 파일이다.

## 다음 컨텍스트에서 이어갈 일

1. 실제 CSV 1개로 `1차 가공 -> 대표이미지 저장 -> LLM 결과 불러오기 -> Cafe24 신규등록 목록 로드` 흐름을 한 번 더 사용자 기준으로 확인
2. 실제 LLM 결과 샘플 5~10개를 비교해서 키워드 품질 룰 추가 조정
3. 특히 확인할 항목
   - A마켓 제목이 너무 길거나 기능어가 과한지
   - B마켓 제목이 A와 충분히 다른 각도로 나오는지
   - OCR 숫자 필터가 필요한 규격까지 제거하지 않는지
   - 검색어설정이 목표 개수를 채우려고 억지 태그를 넣지 않는지

