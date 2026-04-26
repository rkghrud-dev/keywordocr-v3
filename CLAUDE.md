# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build & Run

```bash
# C# WPF 앱 빌드
dotnet build KeywordOcr.App/KeywordOcr.App.csproj

# 앱 실행
dotnet run --project KeywordOcr.App/KeywordOcr.App.csproj
# 또는 빌드된 exe 직접 실행
start "" "KeywordOcr.App/bin/Debug/net8.0-windows/KeywordOcr.App.exe"

# Python 백엔드 의존성 설치
pip install -r backend/requirements.txt

# Python 파이프라인 직접 테스트
py -3 KeywordOcr.App/Bridge/run_pipeline_bridge.py --legacy-root backend --source <file.xlsx> --phase ocr_only
```

## 아키텍처 개요

**C# WPF (.NET 8) + Python 백엔드의 두 레이어 구조**

```
CSV/Excel 입력
  → C# UI (MainWindow.xaml.cs)
    → PythonPipelineBridgeService.cs  [subprocess 실행]
      → Bridge/run_pipeline_bridge.py [CLI args → PipelineConfig]
        → backend/app/services/pipeline.py [메인 오케스트레이션]
          → OCR (Google Vision / local)
          → Claude API 키워드 생성 (legacy_core.py + keyword_builder.py)
          → 이미지 처리 (listing_images.py)
          → Excel 출력 (분리추출후 + B마켓 시트)
  → stdout에 __RESULT__<json> 출력 → C#이 파싱
  → Cafe24 API 직접 호출 (Cafe24ApiClient.cs)
```

### C#↔Python IPC

- C#이 `py -3 run_pipeline_bridge.py --arg value ...` 형태로 subprocess 실행
- Python은 `__RESULT__{"output_root": "...", "output_file": "..."}` 를 stdout에 출력
- `PythonPipelineBridgeService.cs`가 stdout을 읽고 나머지는 `progress?.Report(line)` 으로 로그 전달

### Pipeline Phases

| Phase | 동작 |
|-------|------|
| `full` (기본) | OCR + 키워드 생성 + 이미지 가공 |
| `images` | 이미지 가공만 (OCR/키워드 스킵) |
| `ocr_only` | OCR + 청크 분할만, 키워드는 LLM(Claude) 외부 처리용 |
| `analysis` | OCR + 키워드 생성 (이미지 스킵, phase=images 이후 병렬 실행) |

### A마켓 / B마켓 분리

- 단일 파이프라인 실행이 A마켓(홈런마켓)과 B마켓(준비몰) 결과를 **동시 생성**
- 출력 Excel: `분리추출후` 시트(A마켓) + `B마켓` 시트
- A마켓: 상품명 80~100자, 검색어 최대 20개
- B마켓: A마켓의 부분집합/축약, 63~98자, 검색어 최대 14개
- `pipeline.py`의 `enable_b_market=True`(기본값)일 때 B마켓 시트 생성

## 핵심 파일 위치

| 역할 | 파일 |
|------|------|
| 메인 UI + 전체 C# 로직 | `KeywordOcr.App/MainWindow.xaml(.cs)` |
| Python 프로세스 실행 | `KeywordOcr.App/Services/PythonPipelineBridgeService.cs` |
| Cafe24 API 호출 | `KeywordOcr.App/Services/Cafe24ApiClient.cs` |
| Cafe24 업로드 로직 | `KeywordOcr.App/Services/Cafe24UploadService.cs` (이미지/가격), `Cafe24CreateProductService.cs` (신규등록) |
| 이미지 설정 + B마켓 토큰 경로 저장 | `KeywordOcr.App/Services/ListingImageSettings.cs` → `app_settings.json` |
| 파이프라인 오케스트레이션 | `backend/app/services/pipeline.py` (~4000줄) |
| 키워드 생성 프롬프트 | `backend/app/services/legacy_core.py` |
| 키워드 후처리 | `backend/app/services/keyword_builder.py` |
| 마켓별 키워드 패키지 | `backend/app/services/market_keywords.py` |
| C#→Python 브릿지 스크립트 | `KeywordOcr.App/Bridge/run_pipeline_bridge.py` |

## 설정 파일 경로

| 파일 | 위치 | 내용 |
|------|------|------|
| `app_settings.json` | `backend/` (= `_legacyRoot`) | 이미지 설정, B마켓 토큰 경로 |
| `cafe24_upload_config.txt` | `backend/` | 날짜태그, 이미지 인덱스, 재시도 횟수, 매칭 모드 |
| `cafe24_price_upload_data.json` | `backend/` | 옵션 추가금 데이터 |
| `cafe24_token.json` | `~/Desktop/key/` | 홈런마켓 액세스/리프레시 토큰 |
| `cafe24_token_jb.json` | `~/Desktop/key/` | 준비몰(B마켓) 토큰 (기본 경로) |
| `anthropic_api_key.txt` | `~/Desktop/key/` | Claude API 키 |
| `job_history.json` | `_legacyRoot` (=`backend/`) | 실행 이력 |

## 키워드 생성 원칙 (AGENTS.md에서)

1. **evidence-first**: 상품명 → OCR/Vision → 검색데이터 순서로 근거 사용
2. **no global filler**: 카테고리 교차 필러 삽입 금지
3. **no typo-expansion**: 의도적 오타·커버리지 노이즈 금지
4. **base_name-centered**: 앵커/베이스라인은 확장 제목이 아닌 핵심 상품 정체성에서
5. 근거가 약하면 짧은 출력이 긴 억측보다 낫다
6. B마켓 출력은 검증된 A마켓 토큰의 부분집합/정제 — 새 토큰 추가 금지

## GS코드 처리 규칙

- UI 상품 목록: `GS코드 + A` 접미사만 표시 (B/C/D 변형 제외)
- `Cafe24UploadSupport.GetGsFolders()`: 이미지 폴더도 A 또는 접미사 없는 것만 처리
- 9자리 기준: `GS0000000` (7자리 숫자), 접미사 알파벳 별도 처리

## Cafe24 멀티마켓 등록 흐름

```
TryGetSelectedCafe24Markets() → runHomeMarket, runReadyMarket
  runHomeMarket → Cafe24CreateProductService.CreateAsync()       [분리추출후 시트]
  runReadyMarket → Cafe24CreateProductService.CreateBMarketAsync() [B마켓 시트]
```

- B마켓 토큰: `_bMarketTokenPath` 필드 → `Cafe24ConfigStore.LoadTokenStateB(preferredPath)` → 기본 `cafe24_token_jb.json`
- B마켓 시트 없으면 신규등록 다이얼로그에 `⚠` 경고 표시 후 스킵

## 실행 이력 & 상품 처리 이력

- `JobHistoryService`: `job_history.json`에 실행 레코드 저장, `SelectedCodes`(GS코드 목록) 포함
- 상품 목록 로드 시 `ApplyHistoryToProducts()` 호출 → 이력 있는 상품에 `(MM/dd HH:mm)` 녹색 표시 + 상단 정렬
