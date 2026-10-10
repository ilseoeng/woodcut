// ============================================================
// DB 초기화: schema.sql 을 실행해 ai_quote_* 테이블을 만든다.
// ------------------------------------------------------------
// 사용: npm run db:init
// 여러 번 실행해도 안전하다 (schema.sql 에 IF NOT EXISTS 포함).
// ============================================================
const fs = require('fs');
const path = require('path');
const { getPool, sql } = require('../db');

async function initDb() {
  const sqlPath = path.join(__dirname, '..', 'sql', 'schema.sql');
  const sqlText = fs.readFileSync(sqlPath, 'utf8');

  // MS-SQL는 한 번에 여러 문 실행이 까다로우므로 GO 기준 분할.
  // schema.sql은 IF NOT EXISTS ... BEGIN ... END 블록 단위로 분리되어 있다.
  const batches = sqlText
    .split(/^\s*GO\s*$/m)
    .map((b) => b.trim())
    .filter(Boolean);

  const pool = await getPool();
  console.log(`[init-db] ${batches.length}개 배치 실행 중...`);
  for (let i = 0; i < batches.length; i++) {
    try {
      await pool.request().query(batches[i]);
      console.log(`[init-db] 배치 ${i + 1}/${batches.length} 완료`);
    } catch (err) {
      // 이미 존재하는 객체 등 치명적이지 않은 에러는 경고로 넘김
      console.warn(`[init-db] 배치 ${i + 1} 경고:`, err.message);
    }
  }
  console.log('[init-db] 완료.');
  process.exit(0);
}

initDb().catch((err) => {
  console.error('[init-db] 실패:', err);
  process.exit(1);
});
