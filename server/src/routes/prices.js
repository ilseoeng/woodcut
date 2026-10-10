// ============================================================
// GET /api/prices - 단가표 반환
// ------------------------------------------------------------
// MS-SQL ai_quote_price_table 에서 전체 단가표를 반환.
// 프론트 getUnitPrice() 호환 형식: [{area, mdf18T, pvc18T, ...}]
// ============================================================
const express = require('express');
const router = express.Router();
const { getPool, sql } = require('../db');

router.get('/', async (req, res) => {
  try {
    const pool = await getPool();
    const result = await pool.request().query(
      `SELECT areaMax AS area, misong18T, pyeonbaek18T, pvc18T, mulbau18T, mdf18T
       FROM ai_quote_price_table
       ORDER BY areaMax ASC`
    );
    if (result.recordset.length) {
      return res.json(result.recordset);
    }
    // DB가 비어있으면 빈 배열 (import-prices.js 로 데이터 넣어야 함)
    return res.json([]);
  } catch (err) {
    console.error('[prices] 오류:', err.message);
    return res.status(500).json({ result: 'error', message: '단가표 조회 실패' });
  }
});

module.exports = router;
