# 아이린가구 Woodcut AI 목재 견적

AI 상담을 통해 **목재 종류 · 가로/세로 크기 · 추가가공 · 수량**을 입력받고
MS-SQL 단가표 및 가공비/이벤트 할인을 기반으로 견적을 계산합니다.

## 코드 구성

- `woodcut.html`: 이전 Google Apps Script 버전 (보관용, 현재 운영 아님)
- `woodcut_vm.html`: 최신 프론트엔드 (VM 백엔드 연동)
- `server/`: Node.js Express + Gemini Function Calling + MS-SQL 백엔드
- `server/src/quoteCalculator.js`: AI 견적 도구용 추가가공비 및 할인 계산
- `server/test/`: 가공 조합 및 단가 산출 단위 테스트

## 운영 환경 (2026-10-10 확인)

- 외부 홈페이지: https://34-68-55-157.sslip.io/woodcut/
- API: https://34-68-55-157.sslip.io/woodcut/api/prices
- 상태: https://34-68-55-157.sslip.io/woodcut/healthz
- 2026-08-12에 기존 단독 VM에서 통합 VM으로 이전했습니다.
- VM 운영 세부정보와 SSH 배포 경로는 비공개 Obsidian 운영 문서를 참조합니다.
- Nginx가 /woodcut/ 경로를 VM의 3000 포트로 프록시합니다.

## 로컬 테스트

```bash
cd server
npm ci
npm test
```

가공/기본단가는 분리되어 있습니다. 목재 기본단가는 MS-SQL
`ai_quote_price_table`에서만 조회하며, `quoteCalculator.js`가
가공비와 이벤트 할인을 계산합니다. 배송비는 주문목록별로 별도 계산합니다.

## 배포 유의사항

`server/.env`에는 DB/Gemini/알림톡 비밀정보가 있어 **절대 Git에 넣지 않습니다**.
고객 견적서 이미지가 쌓이는 `server/uploads/` 역시 저장소에 올리지 않습니다.

프론트 `woodcut_vm.html`에서 릴레이의 `/woodcut/` 경로를 감지해
API 요청도 `/woodcut/api/*`으로 보냅니다. VM 내부 3000 포트에서 직접
실행하면 기존 `/api/*` 경로가 사용됩니다.

기존 주문 데이터/단가표 스키마는 그대로 유지합니다.
