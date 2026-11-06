# 오픈소스 구성 요소 고지

오픈소스 엑셀의 직접 작성한 앱 코드는 MIT 라이선스입니다. 아래 구성 요소는 각 저작권과 라이선스를 유지합니다. 이 문서는 해당 라이브러리를 원저작자 대신 재라이선스하거나 공식 제휴를 표시하지 않습니다.

버전은 이 문서를 작성한 시점의 `package-lock.json`과 설치된 패키지 기준입니다. 정확한 재현에는 `npm ci`를 사용하세요. 업데이트 시 버전과 원문 고지도 함께 갱신합니다.

| 구성 요소 | 버전 | 라이선스 | 역할·원문 |
|---|---|---|---|
| [Univer](https://github.com/dream-num/univer) 공개 코어, presets, sheets-core/filter/sort와 관련 공개 패키지 | 1.3.1 | Apache-2.0 | 편집·렌더링·수식·시트·필터·정렬. [원문](public/licenses/UNIVER-LICENSE.txt) |
| [ExcelJS](https://github.com/exceljs/exceljs) | 4.4.0 | MIT | XLSX 파싱·생성. [원문](public/licenses/EXCELJS-LICENSE.txt) |
| [Papa Parse](https://github.com/mholt/PapaParse) | 5.7.0 | MIT | CSV 파싱·생성. [원문](public/licenses/PAPAPARSE-LICENSE.txt) |
| [idb](https://github.com/jakearchibald/idb) | 8.0.4 | ISC | IndexedDB Promise 래퍼. [원문](public/licenses/IDB-LICENSE.txt) |
| [Lucide](https://github.com/lucide-icons/lucide) | 1.70.0 | ISC, 일부 Feather 파생 아이콘 MIT | SVG 아이콘. Feather 파생 목록·저작권을 포함한 [원문](public/licenses/LUCIDE-LICENSE.txt) |
| [Noto Sans KR / Fontsource](https://fontsource.org/fonts/noto-sans-kr) | Fontsource 패키지 5.3.1 | OFL-1.1 | 번들 UI 글꼴. [원문](public/licenses/NOTO-SANS-KR-LICENSE.txt) |
| [React / React DOM](https://github.com/facebook/react) | 설치된 lockfile 참조 | MIT | Univer UI의 런타임. [React](public/licenses/REACT-LICENSE.txt), [React DOM](public/licenses/REACT-DOM-LICENSE.txt) |
| [RxJS](https://github.com/ReactiveX/rxjs) | 설치된 lockfile 참조 | Apache-2.0 | Univer의 반응형 런타임. [원문](public/licenses/RXJS-LICENSE.txt) |
| [uuid](https://github.com/uuidjs/uuid) | 11.1.1 | MIT | ExcelJS 전이 의존성. [원문](public/licenses/UUID-LICENSE.txt) |

Univer의 저작권자는 DreamNum Co., Ltd. 및 해당 기여자입니다. 공식 npm 배포 패키지의 공개 기능을 사용하며, 유료 import/export 서버, 협업 서버, 상용 AI SDK를 포함하지 않습니다. 자체 `src/files.ts`가 ExcelJS와 Univer 스냅샷을 변환합니다. 라이브러리 자체 소스에 변경을 가하지 않고 패키지 API를 통해 앱을 구성합니다.

글꼴 파일은 npm 패키지에서 그대로 번들하며 수정·재명명하지 않습니다. 문서가 사용하는 글꼴과 앱 UI 글꼴은 별개입니다. 아이콘은 Lucide 패키지에서 선택한 SVG를 사용합니다. Lucide 원문에 포함된 Feather MIT 고지도 보존합니다.

`public/licenses/`의 파일은 설치된 패키지에 포함된 원문을 변경 없이 복사한 것으로, Vite 빌드에 의해 운영 정적 파일에도 포함됩니다. 번들 전체에는 이 표 외의 전이 의존성이 포함될 수 있으므로 `package-lock.json`과 각 패키지의 LICENSE·NOTICE를 함께 확인하세요. 의존성 배포 시 라이선스 고지와 필요한 원본 헤더를 제거하지 마세요.

## 참고 프로젝트

[boclair98/rhwp-workspace](https://github.com/boclair98/rhwp-workspace)의 한국어 편집 공간·브라우저 보관함·명시적 다운로드 흐름을 참고했습니다. 해당 저장소의 앱 코드, RHWP Rust/WASM 엔진, 한컴 자산을 복사하거나 포함하지 않았습니다. 엑셀 편집기는 별도 프로젝트입니다.

Microsoft Excel과 그 상표는 각 권리자의 자산입니다. 오픈소스 엑셀은 독립 프로젝트이며 Microsoft의 승인·제휴·완전한 파일 호환을 의미하지 않습니다.
