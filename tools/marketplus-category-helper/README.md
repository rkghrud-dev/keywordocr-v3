# 마켓플러스 카테고리 자동화 사용법

## 1. 카테고리맵 업로드

`upload_category_map.bat` 실행 후 열린 화면에서 `마켓별_카테고리맵.xlsx` 파일을 업로드합니다.

업로드된 데이터는 `marketplus_category_map_store.json`에 저장됩니다. 이 파일은 로컬 런타임 데이터라 Git에는 커밋하지 않습니다.

## 2. 마켓플러스 일괄등록

1. `start_marketplus_helper.bat`을 실행합니다.
2. 카페24 마켓플러스 상품 목록에서 상품을 선택합니다.
3. `일괄보내기`로 registerall 화면을 엽니다.
4. `Category Helper v4` 패널이 상품명으로 업로드된 매칭표를 조회합니다.
5. `selector_json`이 있는 마켓은 드롭다운 카테고리를 자동 적용합니다.
6. `selector_json`이 없으면 카테고리 경로의 마지막 leaf를 검색해 결과 링크 클릭을 시도합니다.

## 3. Chrome 확장 프로그램

Chrome 확장 프로그램은 아래 폴더를 압축해제 확장으로 로드합니다.

`MarketPlusCategoryHelperExtension`

폴더 위치를 옮겼으므로 Chrome 확장 프로그램 관리 화면에서 기존 항목이 깨지면 이 폴더를 다시 로드하세요.

## 주요 파일

- `naver_category_proxy.py`: 로컬 서버, 카테고리맵 업로드/조회, 네이버 카테고리 프록시
- `marketplus-category-helper.user.js`: Tampermonkey/콘솔용 헬퍼 스크립트
- `MarketPlusCategoryHelperExtension`: Chrome 로컬 확장 프로그램
- `marketplus_category_launcher.ps1`: 로컬 서버 시작 스크립트
- `marketplus_category_map_store.json`: 마지막으로 업로드한 카테고리맵 저장 데이터

## 네이버 프록시 설정

카테고리맵 업로드/조회 기능은 별도 API 키 없이 동작합니다.

네이버 쇼핑 카테고리 프록시 기능까지 쓰려면 실행 전에 환경변수를 설정하세요.

```powershell
$env:NAVER_CLIENT_ID="..."
$env:NAVER_CLIENT_SECRET="..."
```
