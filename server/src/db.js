// ============================================================
// MS-SQL 커넥션 풀
// ------------------------------------------------------------
// vault: scripts/wood-and-deco/register_wood_and_deco_orders.py 의 connect()
// 패턴을 Node(mssql 패키지)으로 이식.
//   - SQL 인증 (UID/PWD)
//   - TrustServerCertificate=yes (자체서명 인증서 환경 대응)
// ============================================================
const sql = require('mssql');
const config = require('./config');

let poolPromise = null;

/**
 * 싱글톤 커넥션 풀을 반환한다.
 * 최초 호출 시 풀을 생성하고 연결을 검증한다.
 * @returns {Promise<sql.ConnectionPool>}
 */
async function getPool() {
  if (!poolPromise) {
    poolPromise = new sql.ConnectionPool(config.mssql)
      .connect()
      .then((pool) => {
        console.log('[db] MS-SQL 연결 성공:', `${config.mssql.server}:${config.mssql.port}/${config.mssql.database}`);
        return pool;
      })
      .catch((err) => {
        poolPromise = null; // 실패 시 재시도 가능하도록 초기화
        console.error('[db] MS-SQL 연결 실패:', err.message);
        throw err;
      });
  }
  return poolPromise;
}

/**
 * 풀이 준비되었는지 핑 테스트. (헬스체크용)
 */
async function ping() {
  const pool = await getPool();
  await pool.request().query('SELECT 1 AS ok');
  return true;
}

module.exports = { sql, getPool, ping };
