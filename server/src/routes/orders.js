// ============================================================
// POST /api/orders - 주문(견적) 저장
// ------------------------------------------------------------
// 요청 (프론트 submitOrder 가 보내는 형태):
// {
//   customer: { name, phone, email, memo },
//   orders: [ ...orderList ],
//   shippingItems: [...],
//   shippingCost, totalPrice,
//   quoteImage: "data:image/jpeg;base64,..." (또는 미포함),
//   sessionId, timestamp
// }
// 응답: { result:'success', orderNo }
//
// 처리:
//   1) 주문번호 생성 (YYYYMMDD-NNN 형식)
//   2) 세션 등록 (ai_quote_sessions)
//   3) 주문 헤더 저장 (ai_quote_orders)
//   4) 견적서 이미지 base64 → uploads/ 파일로 저장, 경로만 DB
//   5) 각 물품 라인 저장 (ai_quote_items)
// ============================================================
const express = require('express');
const fs = require('fs');
const path = require('path');
const router = express.Router();
const { getPool, sql } = require('../db');
const config = require('../config');
const { sendOrderNotification } = require('../solapi');
const { sendOrderEmail } = require('../mailer');

/**
 * 주문 완료 알림 발송 (알림톡 1순위, 실패 시 SMS 폴백).
 * 템플릿 변수 #{고객명} #{주문번호} #{총액} #{물품수} 치환.
 */
async function sendOrderConfirmation({ orderNo, customer, itemCount, totalPrice }) {
  const name = (customer && customer.name) || '고객';
  const phone = (customer && customer.phone) || '';
  if (!phone) {
    console.warn('[orders] 발송 생략: 전화번호 없음');
    return;
  }

  const priceStr = Number(totalPrice).toLocaleString() + '원';

  // 알림톡 템플릿 변수 - 실제 승인된 템플릿에 정의된 3개 변수만 매핑.
  // (불필요한 변수를 보내면 솔라피가 발송을 거부할 수 있음)
  const templateVars = {
    '고객명': name,
    '주문번호': orderNo,
    '합계': priceStr,
  };

  // SMS 폴백용 본문 (알림톡 실패 시 사용)
  // 결제 방법 안내 포함: 네이버/쿠팡에서 금액을 수량으로 환산해 결제
  const qtyForPayment = Math.max(1, Math.ceil(Number(totalPrice) / 1000)); // 1,000원 = 1개
  const storeUrl = 'https://naver.me/GJ5OEouG';
  const coupangUrl = 'https://link.coupang.com/a/cQ9keF';
  const smsText =
    `[아이린가구] 주문 접수 완료\n` +
    `주문번호: ${orderNo}\n` +
    `고객: ${name}님\n` +
    `합계: ${priceStr}\n\n` +
    `[결제 방법]\n` +
    `수량 ${qtyForPayment}개로 맞춰 결제 (${priceStr})\n` +
    `배송메시지에 주문번호(${orderNo}) 입력\n\n` +
    `결제: 스토어팜 ${storeUrl}\n` +
    `또는 쿠팡 ${coupangUrl}\n\n` +
    `문의: 031-541-4959`;

  console.log(`[orders] 알림 발송 시작 → ${phone} (${orderNo})`);
  const result = await sendOrderNotification({
    to: phone,
    text: smsText,
    templateVars,
  });
  console.log('[orders] 알림 발송 결과:', result.channel, result.ok ? '성공' : '실패');
  return result;
}

/**
 * data:image/jpeg;base64,....  → 파일로 저장. 파일명 반환.
 * 저장 실패해도 주문 자체는 진행되도록 try/catch.
 */
function saveQuoteImage(dataUrl, orderNo) {
  try {
    if (!dataUrl || typeof dataUrl !== 'string') return null;
    const m = dataUrl.match(/^data:image\/(\w+);base64,(.+)$/);
    if (!m) return null;
    const ext = m[1] === 'jpeg' ? 'jpg' : m[1];
    const buf = Buffer.from(m[2], 'base64');
    const filename = `${orderNo}.${ext}`;
    const dir = config.uploadsDir;
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(path.join(dir, filename), buf);
    return filename;
  } catch (err) {
    console.warn('[orders] 이미지 저장 실패 (주문은 계속):', err.message);
    return null;
  }
}

/**
 * 주문번호 생성: YYYYMMDD-NNN
 * 같은 날 같은 sessionId 의 주문 수 + 1.
 * (정합성이 중요하면 DB 시퀀스/트랜잭션 사용 권장)
 */
async function generateOrderNo(pool, sessionId) {
  const today = new Date();
  const ymd =
    today.getFullYear().toString() +
    String(today.getMonth() + 1).padStart(2, '0') +
    String(today.getDate()).padStart(2, '0');
  const prefix = ymd + '-'; // "20260703-"

  // 같은 날 생성된 주문 수를 세서 NNN 시퀀스 생성.
  // LIKE 매칭 대신 STARTSWITH 스타일로 prefix 길이를 고정 (NVarChar 길이 불일치 방지).
  const result = await pool
    .request()
    .input('prefix', sql.NVarChar(20), prefix)
    .input('prefixLen', sql.Int, prefix.length)
    .query('SELECT COUNT(*) AS cnt FROM ai_quote_orders WHERE LEFT(orderNo, @prefixLen) = @prefix');
  const seq = String((result.recordset[0]?.cnt || 0) + 1).padStart(3, '0');
  return `${ymd}-${seq}`;
}

router.post('/', async (req, res) => {
  const body = req.body || {};
  const { customer, orders, shippingItems, shippingCost, totalPrice, quoteImage, sessionId } = body;

  if (!customer || !customer.name || !customer.phone) {
    return res.status(400).json({ result: 'error', message: '고객 이름/연락처가 필요합니다.' });
  }
  if (!Array.isArray(orders) || orders.length === 0) {
    return res.status(400).json({ result: 'error', message: '주문 항목이 없습니다.' });
  }

  let pool;
  try {
    pool = await getPool();
  } catch (err) {
    console.error('[orders] DB 연결 실패:', err.message);
    return res.status(500).json({ result: 'error', message: '데이터베이스 연결에 실패했습니다.' });
  }

  const orderNo = await generateOrderNo(pool, sessionId);

  // 견적서 이미지 저장
  const imageFilename = saveQuoteImage(quoteImage, orderNo);

  const itemTotal = orders.reduce((s, o) => s + (Number(o.estimate) || 0), 0);

  const transaction = new sql.Transaction(pool);
  try {
    await transaction.begin();
    const req1 = new sql.Request(transaction);

    // 1) 세션 등록 (있으면 스킵되도록)
    await req1
      .input('sessionId', sql.NVarChar(64), sessionId || null)
      .query(
        `IF NOT EXISTS (SELECT 1 FROM ai_quote_sessions WHERE sessionId = @sessionId)
         AND @sessionId IS NOT NULL
         INSERT INTO ai_quote_sessions (sessionId) VALUES (@sessionId)`
      );

    // 2) 주문 헤더
    const req2 = new sql.Request(transaction);
    await req2
      .input('orderNo', sql.NVarChar(40), orderNo)
      .input('sessionId', sql.NVarChar(64), sessionId || null)
      .input('custName', sql.NVarChar(100), customer.name || null)
      .input('custPhone', sql.NVarChar(30), customer.phone || null)
      .input('custEmail', sql.NVarChar(200), customer.email || null)
      .input('custMemo', sql.NVarChar(1000), customer.memo || null)
      .input('itemTotal', sql.BigInt, itemTotal)
      .input('shippingCost', sql.BigInt, Number(shippingCost) || 0)
      .input('grandTotal', sql.BigInt, Number(totalPrice) || itemTotal + (Number(shippingCost) || 0))
      .input('quoteImage', sql.NVarChar(500), imageFilename)
      .query(
        `INSERT INTO ai_quote_orders
          (orderNo, sessionId, custName, custPhone, custEmail, custMemo,
           itemTotal, shippingCost, grandTotal, quoteImage, status)
         VALUES
          (@orderNo, @sessionId, @custName, @custPhone, @custEmail, @custMemo,
           @itemTotal, @shippingCost, @grandTotal, @quoteImage, 'pending')`
      );

    // 3) 물품 라인
    let seq = 0;
    for (const o of orders) {
      const itemType = o.type === 'accessory' ? 'accessory' : o.type === 'shipping' ? 'shipping' : o.type === 'discount' ? 'discount' : 'wood';
      const r = new sql.Request(transaction);
      await r
        .input('orderNo', sql.NVarChar(40), orderNo)
        .input('itemType', sql.NVarChar(20), itemType)
        .input('seq', sql.Int, seq++)
        .input('woodType', sql.NVarChar(40), o.woodTypeValue || null)
        .input('woodName', sql.NVarChar(200), o.woodName || o.name || null)
        .input('width', sql.Int, o.width || null)
        .input('length', sql.Int, o.length || null)
        .input('quantity', sql.Int, o.quantity || 1)
        .input('processingSummary', sql.NVarChar(500), o.processingSummary || null)
        .input('selectedProcsJson', sql.NVarChar(sql.MAX), JSON.stringify(o.selectedProcs || []))
        .input('procDetailsJson', sql.NVarChar(sql.MAX), JSON.stringify(o.procDetails || {}))
        .input('detailTextJson', sql.NVarChar(sql.MAX), JSON.stringify(o.detailText || {}))
        .input('basePrice', sql.BigInt, Number(o.basePrice) || null)
        .input('additionalCost', sql.BigInt, Number(o.additionalCost) || null)
        .input('discount', sql.BigInt, Number(o.discount) || null)
        .input('discountDetailsJson', sql.NVarChar(sql.MAX), JSON.stringify(o.discountDetails || []))
        .input('estimate', sql.BigInt, Number(o.estimate) || 0)
        .query(
          `INSERT INTO ai_quote_items
            (orderNo, itemType, seq, woodType, woodName, [width], [length], quantity,
             processingSummary, selectedProcsJson, procDetailsJson, detailTextJson,
             basePrice, additionalCost, discount, discountDetailsJson, estimate)
           VALUES
            (@orderNo, @itemType, @seq, @woodType, @woodName, @width, @length, @quantity,
             @processingSummary, @selectedProcsJson, @procDetailsJson, @detailTextJson,
             @basePrice, @additionalCost, @discount, @discountDetailsJson, @estimate)`
        );
    }

    await transaction.commit();
    console.log(`[orders] 주문 저장 완료: ${orderNo} (${orders.length}건, ${totalPrice}원)`);

    // 주문 완료 알림 발송 (비동기 - 주문 자체는 발송 결과와 무관하게 성공)
    sendOrderConfirmation({
      orderNo,
      customer,
      itemCount: orders.length,
      totalPrice: Number(totalPrice) || itemTotal + (Number(shippingCost) || 0),
    }).catch((e) => console.error('[orders] 알림 발송 중 오류:', e.message));

    // 고객 이메일이 있으면 주문 확인 메일 발송 (비동기)
    const customerEmail = customer && customer.email;
    if (customerEmail) {
      const grandTotal = Number(totalPrice) || itemTotal + (Number(shippingCost) || 0);
      sendOrderEmail({
        to: customerEmail,
        customerName: (customer && customer.name) || '고객',
        orderNo,
        items: orders,           // 주문 항목 전체 (목재/부자재/할인)
        shippingCost: Number(shippingCost) || 0,
        totalPrice: grandTotal,
      }).catch((e) => console.error('[orders] 이메일 발송 중 오류:', e.message));
    }

    return res.json({
      result: 'success',
      orderNo,
      orderNumber: orderNo, // 프론트 호환
    });
  } catch (err) {
    try {
      await transaction.rollback();
    } catch (_) {}
    console.error('[orders] 주문 저장 실패:', err);
    return res.status(500).json({
      result: 'error',
      message: '주문 저장 중 오류가 발생했습니다.',
      detail: process.env.NODE_ENV === 'production' ? undefined : err.message,
    });
  }
});

module.exports = router;
