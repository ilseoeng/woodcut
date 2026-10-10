# woodcut-server

아이린가구 AI 목재 견적 채팅 — VM 백엔드 (Express + Gemini + MS-SQL).

기존 Google Apps Script 백엔드를 대체한다. AI는 Google Gemini function calling으로 동작하고, 주문 데이터는 MS-SQL에 저장된다.

---

## 구조

```
server/
├── package.json
├── .env.example          # 환경변수 템플릿 (복사해서 .env 생성)
├── .gitignore
├── uploads/              # 견적서 이미지 (base64 → 파일 저장)
└── src/
    ├── index.js          # Express 진입점
    ├── config.js         # 환경변수 로드/검증
    ├── db.js             # MS-SQL 커넥션 풀
    ├── gemini.js         # Gemini 클라이언트 + function-calling 루프
    ├── tools/
    │   ├── registry.js   # 도구 메타데이터 (단일 진실 공급원)
    │   └── schema.js     # Gemini function declarations 생성
    ├── routes/
    │   ├── prices.js     # GET  /api/prices
    │   ├── chat.js       # POST /api/chat
    │   ├── orders.js     # POST /api/orders
    │   └── load.js       # GET  /api/orders/load
    ├── sql/
    │   └── schema.sql    # ai_quote_* 테이블 DDL
    └── scripts/
        └── init-db.js    # npm run db:init 으로 실행
```

---

## API

| 메서드 | 경로 | 설명 |
|---|---|---|
| GET  | `/api/prices` | MS-SQL의 `ai_quote_price_table`에서 단가표 조회 (대체 단가 없음) |
| POST | `/api/chat` | AI 채팅. `{message, sessionId, state}` → `{reply?, tool_calls?}` |
| POST | `/api/orders` | 주문 저장. 고객정보 + orderList + 견적서 이미지 |
| GET  | `/api/orders/load?phone=&orderNo=` | 주문 내역 조회 |
| GET  | `/healthz` | 헬스체크 (DB 연결 상태 포함) |

---

## 로컬 개발

### 1) 의존성 설치

```bash
cd server
npm install
```

### 2) 환경변수 설정

`.env.example` 을 복사해 `.env` 를 만들고 실제 값을 채운다.

```bash
cp .env.example .env
```

필수값:
- `MSSQL_HOST`, `MSSQL_USERNAME`, `MSSQL_PASSWORD` (DB 연결 정보는 운영 환경변수에서 설정)
- `GEMINI_API_KEY` (Google AI Studio 에서 발급)

> 비밀번호/키는 절대 Git 에 커밋하지 않는다 (`.gitignore` 에 `.env` 포함됨).

### 3) DB 테이블 생성 (최초 1회)

```bash
npm run db:init
```

`ai_quote_sessions`, `ai_quote_orders`, `ai_quote_items` 세 테이블이 생성된다. 여러 번 실행해도 안전하다 (`IF NOT EXISTS`).

### 4) 서버 실행

```bash
npm start          # 일반 실행
npm run dev        # 파일 변경 시 자동 재시작 (node --watch)
```

서버가 `http://localhost:3000` 에서 실행된다.

### 5) 프론트엔드 연결

프론트엔드 `woodcut_vm.html` 는 기본적으로 `location.origin` 을 API 베이스로 사용한다.

- **VM에서 통합 서빙**: `index.js` 가 상위 폴더의 `woodcut_vm.html` 를 루트에서 서빙한다. `http://VM주소:3000/` 으로 접속하면 끝.
- **로컬에서 파일로 열 때**: `woodcut_vm.html` 상단의 `VM_BASE` 변수를 `'http://localhost:3000'` 로 설정한다.

---

## 구글 VM (Compute Engine) 배포

### 1) VM 생성 (Ubuntu 22.04)

```bash
gcloud compute instances create woodcut-server \
  --zone=asia-northeast3-a \
  --machine-type=e2-small \
  --image-family=ubuntu-2204-lts \
  --image-project=ubuntu-os-cloud \
  --tags=http-server
```

> 리전 `asia-northeast3-a` (서울). 기종 `e2-small` 이면 충분 (트래픽에 따라 조정).

### 2) 방화벽: 80/443 허용

```bash
gcloud compute firewall-rules create allow-http \
  --allow tcp:80,tcp:443 \
  --target-tags=http-server
```

### 3) VM 초기 설정

```bash
gcloud compute ssh woodcut-server --zone=asia-northeast3-a

# VM 안에서:
sudo apt update && sudo apt install -y nginx
curl -fsSL https://deb.nodesource.com/setup_20.x | sudo -E bash -
sudo apt install -y nodejs
sudo npm install -g pm2
```

### 4) 코드 배포

```bash
# 로컬에서 VM으로 업로드 (예: gcloud scp)
gcloud compute scp --recurse ./server woodcut-server:~/woodcut-server --zone=asia-northeast3-a

# VM 안에서:
cd ~/woodcut-server
npm ci
cp .env.example .env
nano .env          # 실제 값 입력
npm run db:init    # 테이블 생성
pm2 start src/index.js --name woodcut
pm2 save
pm2 startup        # 재부팅 시 자동 실행
```

### 5) (권장) Nginx 리버스프록시 + HTTPS

`/etc/nginx/sites-available/woodcut`:

```nginx
server {
    listen 80;
    server_name woodcut.example.com;   # 본인 도메인

    client_max_body_size 25M;          # 견적서 이미지

    location / {
        proxy_pass http://127.0.0.1:3000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
    }
}
```

```bash
sudo ln -s /etc/nginx/sites-available/woodcut /etc/nginx/sites-enabled/
sudo nginx -t && sudo systemctl reload nginx

# HTTPS (도메인 있을 때)
sudo apt install -y certbot python3-certbot-nginx
sudo certbot --nginx -d woodcut.example.com
```

---

## 도구(Function Calling) 아키텍처

```
프론트 (woodcut_vm.html)              백엔드 (server/)
┌──────────────────────┐            ┌─────────────────────────┐
│ 사용자 입력          │            │ /api/chat               │
│   ↓                  │            │   ↓                     │
│ sendToAI()           │── message ─→│ Gemini.chat()           │
│                      │   state     │   + functionDeclarations│
│                      │            │   ↓                     │
│                      │←─ reply ───│ Gemini 응답              │
│                      │   tool_calls│   (text + functionCall) │
│ executeToolCalls()   │            │                         │
│   ↓ TOOL_REGISTRY    │            │ 도구 '실행'은 프론트가 담당 │
│ UI 변경              │            │ (UI를 바꿔야 하므로)      │
└──────────────────────┘            └─────────────────────────┘
```

**핵심 원칙**: 도구의 *정의*(`tools/registry.js`)는 백엔드가, *실행*(`TOOL_REGISTRY` in `woodcut_vm.html`)은 프론트가 담당. 두 레지스트리의 키와 매개변수는 **반드시 일치**해야 한다.

도구를 추가하려면 양쪽을 모두 업데이트한다:
1. `server/src/tools/registry.js` — 설명/스키마 추가
2. `woodcut_vm.html` 의 `TOOL_REGISTRY` — 실행 함수 추가

---

## 문제 해결

| 증상 | 확인 |
|---|---|
| `[db] MS-SQL 연결 실패` | `.env` 의 `MSSQL_*` 값. 방화벽에서 VM→DB 서버의 SQL 포트 허용되었는지 |
| `[chat] Gemini 호출 오류` | `GEMINI_API_KEY` 유효성. quota |
| 도구가 호출 안 됨 | `tools/registry.js` 와 프론트 `TOOL_REGISTRY` 키 일치 여부 |
| 단가가 이상함 | `price_table.json` 또는 DB 단가표 테이블 확인 |

---

## 참고

- MS-SQL 연결 정보: 비공개 Obsidian DB 연결 안내
- 기존 연결 패턴(참고용): `scripts/wood-and-deco/register_wood_and_deco_orders.py` 의 `connect()`
- Gemini function calling 문서: https://ai.google.dev/gemini-api/docs/function-calling
