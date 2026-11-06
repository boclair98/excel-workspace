# 오픈소스 엑셀 디자인 계획

## 근거와 범위
작업 폴더는 비어 있었다. 참고 프로젝트는 사용자의 boclair98/rhwp-workspace(오픈소스 한글) README, package.json, design.md다. 한글의 실제 편집 공간, 간편 도구, 브라우저 보관함과 명시적인 파일 다운로드를 엑셀 작업에 적용한다. 한글 엔진 코드를 복사하지 않는다. 새 엑셀 편집기는 TypeScript/Vite, Node 정적 서버와 최소 AI 프록시로 참고 프로젝트의 배포 구조를 따른다. 문서 서버/DB는 만들지 않는다.

## 확정 입력
- Topic: 한국어 오픈소스 스프레드시트 작업 공간
- Offer: XLSX/CSV 편집, 수식 계산, 여러 시트, 로컬 보관과 AI 수식 도우미
- Audience: 엑셀 파일을 브라우저에서 편집하고 다운로드하려는 한국어 사용자
- Primary user action: 파일 열기 → 파일 입력 → 로컬 파싱 → 편집
- Deliverable: 실제 편집기와 보관함 디자인 1개
- Style direction: 한글 프로젝트의 담백한 타이포그래피와 간편 도구를 계승한 작업 중심 UI
- Tonal key: 밝은 흰색/회색, 엑셀 작업을 구분하는 차분한 초록색
- Icon policy: Lucide 단색 SVG, 브랜드 자산 복제 없음
- Proof assets: 직접 입력한 ‘예시’ 업무표와 실제 편집 표면만 사용
- Motion intensity: 120ms 조작 피드백, reduced-motion 준수

## Composition contract
- Hero archetype: 해당 없음, 작업 편집기
- First-viewport structure: 68px 문서 헤더, 48px 파일 도구, 접을 수 있는 224px 보관함, 나머지 공간에 실제 Univer 셀 그리드. 오른쪽 AI 패널은 요청할 때 연다.
- Primary visual/proof artifact: 편집·계산 가능한 실제 셀과 시트 탭
- Interaction behavior: 열기=importFile, 새 통합문서=createWorkbook, 보관=saveCurrent, 다운로드=exportXlsx, AI=질문과 선택 범위 명시 전송 → 제안 → 선택 셀 적용
- Reuse boundary: 한글 프로젝트의 정보 구조만 참고, 엑셀 모델/입출력/AI는 별도 구현

## 시각 토큰과 반응형
text #191f28, secondary #4e5968, muted #6b7684, surface #ffffff, workspace #f7f9f8, border #e5e9e7, primary #18794e, hover #10643f, selected #eaf5ef, danger #b42318. Noto Sans KR variable 번들 + system-ui, sans-serif. 제목 18px/700/1.4, 본문14px/400/1.5, 보조12px/400/1.5, 주요 버튼14px/600. 간격4/8/12/16/24px, 반경8/12px, 면 구분은 1px border. 헤더와 도구는 그리드 폭을 따르며 읽기 영역은 전체 폭이다.
모바일: 제목·상태 한 줄과 파일 도구 내부 가로 스크롤. 보관함/AI는 native dialog로 표시해 포커스를 가둔다. 기본 Univer 리본은 접고 별도 44px 간편 셀 입력과 서식 버튼 제공. 그리드의 내부 가로 이동은 허용하되 페이지 넘침은 없어야 한다. 한국어 keep-all/overflow-wrap:anywhere, flex 자식 min-width:0. 모든 대화상자 Escape/포커스 복귀, 입력 레이블 제공.

## 내용과 기능 계약
첫 진입은 ‘새 통합문서’의 빈 셀. 보관함에서 ‘예시 가계부’를 명시적으로 생성 가능. 자동 보관은 IndexedDB, 성공 트랜잭션 뒤만 저장 완료 표시. 파일 이름·저장 상태·파일 다운로드가 별개임을 안내한다. 파일은 기본적으로 서버 전송하지 않는다. AI 질문은 관리형 Claude에 전송하며 체크한 경우에만 선택 범위(최대200셀)를 추가한다. AI는 자동 실행하지 않으며 제안 적용은 사용자 클릭을 요구한다. 개인 API 키는 생성/등록하지 않는다. Coders 로그인과 사용량 한도가 적용됨을 표시한다. 키 없는 로컬 환경에서는 미연결 상태를 표시한다.
XLSX는 ExcelJS로 값·수식·기본 서식·병합·열 너비·행 높이·여러 시트를 변환한다. VBA/차트/피벗/개체의 완전 호환을 주장하지 않으며 가져오기 전에 손실 범위를 안내한다. CSV는 현재 시트만 값으로 내보내고 수식 주입 방지한다. 실제 원본 덮어쓰기 없음.

## 구현과 검증
src/main.ts 앱 셸과 이벤트, src/editor.ts Univer 어댑터, src/files.ts XLSX/CSV, src/storage.ts IndexedDB, src/style.css, server.mjs/ai.mjs, tests/*.test.ts, Dockerfile/coders.yaml, README/LICENSE/THIRD_PARTY_NOTICES. API 계약은 POST /api/ai, GET /api/status, GET /healthz. 문서 DB 없음. TypeScript build, 실제 XLSX 왕복/CSV/저장 검증과 AI 입력 경계 테스트. 앱을 실행해 4개 뷰포트·키보드·실제 편집/다운로드/보관·오류 상태·콘솔을 확인한다.

## 사용자 추가 요청: 엑셀에 익숙한 리본과 개인화
홈·삽입·데이터·보기의 실제 동작 리본 탭을 추가한다. 홈은 글꼴/색상/정렬/병합/표시 형식, 삽입은 행·열·시트, 데이터는 정렬·필터·찾기·내용 지우기, 보기는 틀 고정·확대·보관함/도구와 개인화다. PC 리본은 그룹 캡션으로 구분하고 모바일은 44px 조작을 유지하는 내부 스크롤이다. 수식 입력줄과 하단 시트 탭은 엑셀의 익숙한 위치를 유지한다. 세 가지 강조색과 화면 밀도 설정은 이 브라우저에만 저장한다. 문서 서식 색과 UI 강조색을 구분한다. 새 통합문서는 빈 표이며 예시 데이터는 선택해서 생성한다.

## 확정 생성 프롬프트와 출력
한국어 사용자가 파일을 열자마자 표를 편집하는 오픈소스 엑셀 작업 공간을 구현하라. 사용자 소유 한글 프로젝트의 간편 편집·보관·다운로드 흐름을 참고하되 별도 제품으로 구성하라. 실제 표를 주 화면으로 하고 흰색·회색·초록, 명확한 제목·서식·셀 위계와 44px 모바일 조작을 사용하라. 설명이나 홍보용 가짜 지표를 만들지 말라. Univer의 실제 편집·수식·여러 시트, ExcelJS 파일 변환과 로컬 보관을 연결하라. AI는 사용자 질문 및 동의한 선택 범위만 전송하며 검토 후 적용하라. 오류·로딩·빈 상태와 접근성을 제공하고 실제 화면/파일 왕복을 검증하라. Output target: src/main.ts, src/editor.ts, src/files.ts, src/storage.ts, src/style.css 및 서버/검증 파일.
