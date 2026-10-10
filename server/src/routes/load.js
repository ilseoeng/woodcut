// ============================================================
// GET /api/orders/load - 주문 내역 불러오기
// ------------------------------------------------------------
// 쿼리: ?phone=010-xxxx-xxxx&orderNo=YYYYMMDD-NNN
// 응답: { result:'success', data:{ orderNo, items:[...], shippingItems, totalPrice } }
//
// DB에서 읽어 프론트 applyLoadedOrderData 가 기대하는 정규화 형태로 변환.
// ============================================================
const express = require('express');
const router = express.Router();
const { getPool, sql } = require('../db');

function safeParse(s) {
  if (!s) return null;
  try {
    return JSON.parse(s);
  } catch {
    return null;
  }
}

router.get('/', async (req, res) => {
  const phone = (req.query.phone || '').trim();
  const orderNo = (req.query.orderNo || '').trim();

  if (!phone || !orderNo) {
    return res.status(400).json({ result: 'error', message: 'phone 과 orderNo 가 필요합니다.' });
  }

  let pool;
  try {
    pool = await getPool();
  } catch (err) {
    console.error('[load] DB 연결 실패:', err.message);
    return res.status(500).json({ result: 'error', message: '데이터베이스 연결에 실패했습니다.' });
  }

  try {
    const hdr = await pool
      .request()
      .input('orderNo', sql.NVarChar(40), orderNo)
      .input('phone', sql.NVarChar(30), phone)
      .query(
        `SELECT orderNo, custName, custPhone, itemTotal, shippingCost, grandTotal, quoteImage, status, createdAt
         FROM ai_quote_orders
         WHERE orderNo = @orderNo AND custPhone = @phone`
      );

    if (!hdr.recordset.length) {
      return res.json({ result: 'success', data: null });
    }

    const order = hdr.recordset[0];

    const itemsRs = await pool
      .request()
      .input('orderNo', sql.NVarChar(40), orderNo)
      .query(
        `SELECT itemType, woodType, woodName, [width], [length], quantity,
                processingSummary, selectedProcsJson, procDetailsJson,
                basePrice, additionalCost, discount, discountDetailsJson, estimate
         FROM ai_quote_items
         WHERE orderNo = @orderNo
         ORDER BY seq`
      );

    // 프론트 applyLoadedOrderData 가 기대하는 형태로 정규화
    const items = itemsRs.recordset.map((row) => ({
      type: row.itemType,
      woodType: row.woodType,
      woodName: row.woodName,
      name: row.itemType === 'accessory' ? row.woodName : undefined,
      width: row.width || 0,
      length: row.length || 0,
      quantity: row.quantity || 1,
      processingSummary: row.processingSummary || undefined,
      selectedProcs: safeParse(row.selectedProcsJson) || [],
      procDetails: safeParse(row.procDetailsJson) || {},
      discount: row.discount || 0,
      discountDetails: safeParse(row.discountDetailsJson) || [],
      estimate: row.estimate || 0,
    }));

    return res.json({
      result: 'success',
      data: {
        orderNo: order.orderNo,
        custName: order.custName,
        items,
        shippingCost: order.shippingCost,
        totalPrice: order.grandTotal,
        quoteImage: order.quoteImage,
        createdAt: order.createdAt,
      },
    });
  } catch (err) {
    console.error('[load] 조회 실패:', err);
    return res.status(500).json({
      result: 'error',
      message: '주문 내역 조회 중 오류가 발생했습니다.',
      detail: process.env.NODE_ENV === 'production' ? undefined : err.message,
    });
  }
});

module.exports = router;
