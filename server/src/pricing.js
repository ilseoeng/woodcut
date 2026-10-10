// ============================================================
// 단가표 조회 로직
// ------------------------------------------------------------
// DB(ai_quote_price_table)에서 면적에 맞는 단가를 찾는다.
// 프론트의 getUnitPrice()와 동일한 규칙:
//   area <= areaMax 인 첫 번째 행을 사용.
// ============================================================
const { getPool, sql } = require('./db');

const WOOD_COLUMNS = ['misong18T', 'pyeonbaek18T', 'pvc18T', 'mulbau18T', 'mdf18T'];

/**
 * 면적에 해당하는 단가 행 전체를 반환.
 * @param {number} area - mm²
 * @returns {Promise<{areaMax, misong18T, pyeonbaek18T, pvc18T, mulbau18T, mdf18T}|null>}
 */
async function getPriceRow(area) {
  const pool = await getPool();
  const result = await pool
    .request()
    .input('area', sql.Int, Math.max(0, Math.floor(area))).query(
      `SELECT TOP 1 areaMax, misong18T, pyeonbaek18T, pvc18T, mulbau18T, mdf18T
       FROM ai_quote_price_table
       WHERE areaMax >= @area
       ORDER BY areaMax ASC`
    );
  return result.recordset[0] || null;
}

/**
 * 특정 목재의 면적당 단가.
 * @param {string} woodKey - mdf18T 등
 * @param {number} area - mm²
 * @returns {Promise<number>} 단가 (찾지 못하면 0)
 */
async function getUnitPrice(woodKey, area) {
  if (!WOOD_COLUMNS.includes(woodKey)) return 0;
  const row = await getPriceRow(area);
  return row ? Number(row[woodKey]) || 0 : 0;
}

/**
 * 모든 목재의 단가를 비교용으로 반환 (저렴한 순 정렬은 호출侧에서).
 * @param {number} area
 * @returns {Promise<Array<{woodKey, price}>>}
 */
async function getAllUnitPrices(area) {
  const row = await getPriceRow(area);
  if (!row) return [];
  return WOOD_COLUMNS.map((k) => ({ woodKey: k, price: Number(row[k]) || 0 }));
}

module.exports = { WOOD_COLUMNS, getPriceRow, getUnitPrice, getAllUnitPrices };
