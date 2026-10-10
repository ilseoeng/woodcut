// ============================================================
// 환경변수 로드 및 검증
// ------------------------------------------------------------
// 우선순위: 시스템 환경변수 > .env 파일
//   - MS-SQL 자격증명은 Windows 환경변수에 이미 세팅되어 있으므로
//     dotenv의 override 옵션을 끄면 .env 의 플레이스홀더가 실제 값을 덮어쓰지 않는다.
//   - .env 는 개발용(로컬에서 환경변수가 없을 때) 보조 수단.
// 필수값이 없으면 명확한 에러로 종료시킨다 (침묵 실패 방지).
// ============================================================
require('dotenv').config({ override: false });

const required = ['MSSQL_HOST', 'MSSQL_USERNAME', 'MSSQL_PASSWORD', 'GEMINI_API_KEY'];
const missing = required.filter((k) => !process.env[k]);
if (missing.length) {
  console.error('[config] 필수 환경변수가 없습니다:', missing.join(', '));
  console.error('[config] 시스템 환경변수를 확인하거나 .env.example 을 참고해 .env 를 만드세요.');
  process.exit(1);
}

// MSSQL_HOST 가 "db.example.internal,1433" 처럼 포트를 포함할 수 있으므로 분리.
// (Windows 환경변수 세팅 방식 때문. mssql 패키지는 server/port 를 별도로 받는다.)
const rawHost = process.env.MSSQL_HOST || '';
const hostPortMatch = rawHost.match(/^(.+?),\s*(\d+)$/);
const server = hostPortMatch ? hostPortMatch[1].trim() : rawHost.trim();
const port = parseInt(
  hostPortMatch ? hostPortMatch[2] : (process.env.MSSQL_PORT || '1433'),
  10
);

const config = {
  port: parseInt(process.env.PORT || '3000', 10),
  corsOrigin: process.env.CORS_ORIGIN
    ? process.env.CORS_ORIGIN.split(',').map((s) => s.trim())
    : null, // null = 모든 출처 허용 (개발용)

  mssql: {
    server,
    port,
    database: process.env.MSSQL_DATABASE || 'irin',
    user: process.env.MSSQL_USERNAME,
    password: process.env.MSSQL_PASSWORD,
    encrypt: String(process.env.MSSQL_ENCRYPT || 'false').toLowerCase() === 'true',
    trustServerCertificate:
      String(process.env.MSSQL_TRUST_SERVER_CERTIFICATE || 'true').toLowerCase() === 'true',
    pool: {
      max: 10,
      min: 0,
      idleTimeoutMillis: 30000,
    },
    options: {
      encrypt: String(process.env.MSSQL_ENCRYPT || 'false').toLowerCase() === 'true',
      trustServerCertificate:
        String(process.env.MSSQL_TRUST_SERVER_CERTIFICATE || 'true').toLowerCase() === 'true',
    },
  },

  gemini: {
    apiKey: process.env.GEMINI_API_KEY,
    model: process.env.GEMINI_MODEL || 'gemini-2.5-flash',
  },

  solapi: {
    SOLAPI_API_KEY: process.env.SOLAPI_API_KEY || '',
    SOLAPI_API_SECRET: process.env.SOLAPI_API_SECRET || '',
    SOLAPI_SENDER: process.env.SOLAPI_SENDER || '',
    SOLAPI_KAKAO_SENDER_KEY: process.env.SOLAPI_KAKAO_SENDER_KEY || '',
    SOLAPI_KAKAO_PF_ID: process.env.SOLAPI_KAKAO_PF_ID || '',
    SOLAPI_KAKAO_TEMPLATE_CODE: process.env.SOLAPI_KAKAO_TEMPLATE_CODE || '',
  },

  // 참고: 메일 자격증명은 .env가 아닌 DB(거래처등록 테이블)에서 실시간 조회.
  // 비밀번호 변경 시 DB만 업데이트하면 자동 반영됨. mailer.js 참조.

  uploadsDir: require('path').join(__dirname, '..', 'uploads'),
};

module.exports = config;
