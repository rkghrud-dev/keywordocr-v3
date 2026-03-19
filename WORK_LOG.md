# KeywordOCR v3 작업 로그

## 프로젝트 개요
- **위치**: `C:\Users\rkghr\Desktop\프로젝트\keywordocr\keywordocr-v3\KeywordOcr.App`
- **기술**: C# WPF (.NET 8) + Python 백엔드
- **목적**: CSV 입력 → 전처리/OCR/AI 키워드 생성 → 업로드용 Excel + 대표이미지 → Cafe24 업로드

## 완료된 작업

### 1. Claude API 통합 (v2에서 마이그레이션)
- `app/services/anthropic_wrapper.py` - OpenAI 호환 래퍼 (Vision 포함)
- `app/services/env_loader.py` - ANTHROPIC_API_KEY 로딩 (`~/Desktop/key/anthropic_api_key.txt`)
- `app/services/legacy_core.py` - 6개 프롬프트 전체 재작성 (한국 이커머스 키워드 규칙 기반)
- 기본 모델: `claude-haiku-4-5-20251001` (설정에서 변경 가능)

### 2. v3 WPF 앱 신규 생성
완전히 새로 만든 깔끔한 WPF 앱. 탭 구조:

#### Tab 1: 실행
- **드래그 앤 드롭** 파일 선택 (CSV/Excel)
- **상품코드 선택**: GS코드 A만 표시 (B/C/D 제외), 상품명 컬럼에서도 GS코드 추출
- **AI 모델 선택**: haiku(기본), sonnet, opus, gpt-4.1 등
- **실행 버튼**: 전체실행 / 키워드만 / 대표이미지만 / 결과폴더 열기
- 완료 시 결과 폴더 자동 열림 + 업로드용 엑셀 클립보드 복사

#### Tab 2: Cafe24 업로드
- 설정: 날짜태그, 매칭모드(PREFIX/CONTAINS/EXACT), 이미지 인덱스, 재시도
- **Cafe24 이미지+가격 업로드** / **신규상품 등록** 버튼
- 업로드 결과 DataGrid (GS코드, 상품번호, 상태, 가격)
- 업로드 로그 열기

#### Tab 3: 옵션 가격
- GPT 파일에서 가격 데이터 로드 (분리추출전 시트)
- DataGrid: 체크, GS코드, 옵션명, 공급가, 판매가, 추가금액(편집), 소비자가
- 가격 재계산 → `cafe24_price_upload_data.json` 저장

#### Tab 4: 실행 이력
- 파이프라인 실행 완료 시 자동 저장 (`job_history.json`)
- DataGrid: 일시, 파일, 요약, 상태, 메모
- **우클릭 팝업 메뉴**:
  - 결과 불러오기
  - Cafe24 상품 엑셀 업로드 (브라우저 열기 + 경로 클립보드 복사)
  - Cafe24 이미지+가격 업로드 (API 바로 실행)
  - Cafe24 신규상품 등록 (API 바로 실행)
  - 업로드용 엑셀 열기 / 결과 폴더 열기 / 경로 복사
  - 메모 수정 / 이력 복제 / 삭제
- 더블클릭으로 결과 바로 불러오기

#### Tab 5: 설정
- 대표이미지: 로고(경로/비율/투명도/위치), JPEG품질, 회전, 자동대비, 샤프닝, 좌우반전
- Cafe24 토큰: 파일 선택, Mall ID 표시, 토큰 갱신
- 경로: Python 루트, v3 루트

### 3. GS코드 A만 처리
- **상품 선택 UI**: 코드 컬럼 + 상품명 컬럼 둘 다에서 GS코드 검색, A접미사만 표시
- **필터링**: 선택된 상품만 필터링 시에도 상품명에서 GS코드 매칭
- **Cafe24 업로드**: `GetGsFolders()`에서 A 또는 접미사 없는 폴더만 처리

### 4. 토큰 검색 경로
Cafe24ConfigStore에서 토큰 파일 검색 순서:
1. 사용자 지정 경로
2. 설정파일에 지정된 경로
3. `v2Root/cafe24_token.txt`
4. `legacyRoot/cafe24_token.txt`
5. `Desktop/keywordocr/cafe24_token.txt` (추가됨)

## 파일 구조

```
keywordocr-v3/KeywordOcr.App/
├── KeywordOcr.App.csproj (net8.0-windows, ClosedXML)
├── App.xaml / App.xaml.cs
├── MainWindow.xaml / MainWindow.xaml.cs (메인 UI + 전체 로직)
├── MemoDialog.xaml / MemoDialog.xaml.cs (메모 수정 다이얼로그)
├── Bridge/
│   └── run_pipeline_bridge.py (Python 파이프라인 호출)
└── Services/
    ├── PythonPipelineBridgeService.cs (Python 프로세스 실행)
    ├── ListingImageSettings.cs (대표이미지 설정 record)
    ├── JobHistoryService.cs (실행 이력 관리)
    ├── Cafe24UploadService.cs (이미지+가격 업로드)
    ├── Cafe24CreateProductService.cs (신규상품 등록)
    ├── Cafe24ApiClient.cs (Cafe24 REST API)
    ├── Cafe24ConfigStore.cs (토큰/설정 관리)
    ├── Cafe24UploadSupport.cs (업로드 헬퍼 - GS폴더 A만 필터링)
    ├── Cafe24UploadModels.cs (데이터 모델)
    └── WorkbookFileLoader.cs (ClosedXML 래퍼)
```

## 레거시 Python 경로
- `C:\Users\rkghr\Desktop\프로젝트\keywordocr\app\services\` - 핵심 서비스
- `C:\Users\rkghr\Desktop\프로젝트\keywordocr\cafe24_upload.py` - Cafe24 업로드 스크립트
- API 키: `C:\Users\rkghr\Desktop\key\` 폴더

## 알려진 이슈 / 참고사항
- Python 실행: `py -3` 명령 사용 (py launcher)
- `py -m pip install anthropic` 으로 설치 필요
- Claude API: temperature와 top_p 동시 사용 불가 → anthropic_wrapper에서 처리
- 스마트 따옴표(U+201C 등) 주의 → Python 문법 오류 발생 가능

## 다음 작업 후보
- 이미지 선택 UI (대표/추가 이미지 미리보기 + 선택)
- 제외단어 관리 (user_stopwords.json)
- 모델 선택을 Bridge에 전달 (현재 bridge에서 하드코딩)
- 업로드 결과 필터링 (실패건만 보기, 재시도)
