// ============================================================
// 이메일 발송 (nodemailer + 네이트 메일 SMTP)
// ------------------------------------------------------------
// 핵심: 메일 자격증명을 .env에 두지 않고, 발송 시마다 DB의
// '거래처등록' 테이블에서 실시간 조회.
//   → 비밀번호가 바뀌어도 DB만 업데이트하면 자동 반영.
//
// DB 컬럼 (거래처등록 WHERE 거래처명='네이트'):
//   통신아이디, 통신비밀번호
// ============================================================
const nodemailer = require('nodemailer');
const { getPool, sql } = require('./db');

// 주문 내역 확인 페이지 (고객용)
const ORDER_CHECK_URL = process.env.ORDER_CHECK_URL || 'http://34.47.110.21:3000/';
// 결제 사이트
const STORE_URL = 'https://naver.me/GJ5OEouG';
const COUPANG_URL = 'https://link.coupang.com/a/cQ9keF';

// 캐시: 같은 비밀번호로 재사용 (DB 부하 감소). 값이 바뀌면 자동 갱신.
let _cachedCreds = null; // { user, pass, host, port }
let _transporter = null;

/**
 * DB에서 네이트 메일 자격증명을 실시간 조회.
 * 캐시가 있으면 재사용하되, 발송 실패 시 캐시 무효화해 재조회.
 */
async function getCredentials() {
  // 캐시가 있으면 그대로 사용
  if (_cachedCreds) return _cachedCreds;

  const pool = await getPool();
  const r = await pool
    .request()
    .input('name', sql.VarChar, '네이트')
    .query('SELECT 통신아이디, 통신비밀번호 FROM 거래처등록 WHERE 거래처명 = @name');

  const row = r.recordset[0];
  if (!row || !row.통신아이디 || !row.통신비밀번호) {
    throw new Error("DB에서 '네이트' 거래처 자격증명을 찾을 수 없습니다");
  }

  _cachedCreds = {
    user: row.통신아이디,
    pass: row.통신비밀번호,
    host: 'smtp.mail.nate.com',
    port: 465,
    from: `"아이린가구" <${row.통신아이디}@nate.com>`,
  };
  _transporter = null; // 자격증명 바뀌었으니 transporter도 새로 만들도록 무효화
  return _cachedCreds;
}

/**
 * 자격증명에 맞는 transporter 반환 (캐싱).
 * 인증 실패 등으로 발송이 실패하면 캐시를 무효화하도록 invalidateCache() 호출.
 */
async function getTransporter() {
  if (_transporter) return _transporter;
  const creds = await getCredentials();
  _transporter = nodemailer.createTransport({
    host: creds.host,
    port: creds.port,
    secure: creds.port === 465,
    auth: { user: creds.user, pass: creds.pass },
    tls: { rejectUnauthorized: false },
  });
  return _transporter;
}

/**
 * 캐시 무효화 - 인증 실패 시 호출하면 다음 발송 시 DB 재조회.
 */
function invalidateCache() {
  _cachedCreds = null;
  _transporter = null;
}

/**
 * 주문 항목을 HTML 테이블 행으로 변환.
 * 목재/부자재/할인/배송비 타입을 모두 처리.
 */
function itemsToTableRows(items) {
  if (!Array.isArray(items) || items.length === 0) {
    return '<tr><td colspan="5" style="padding:12px;text-align:center;color:#999;">주문 내역이 없습니다.</td></tr>';
  }
  return items.map((o, idx) => {
    const estimate = Number(o.estimate) || 0;
    // 타입별 표시
    let name = o.woodName || o.name || '미확정';
    let detail = '';
    const qty = o.quantity || 1;

    if (o.type === 'discount') {
      detail = '이벤트 할인';
      name = o.name || '이벤트 할인';
    } else if (o.type === 'shipping') {
      detail = '배송';
      name = o.name || '배송비';
    } else if (o.type === 'accessory') {
      detail = '부자재';
      name = '[부자재] ' + (o.name || '');
    } else {
      // 목재
      const size = (o.width && o.length) ? `${o.width} × ${o.length} mm` : '-';
      const proc = o.processingSummary && o.processingSummary !== '-' ? o.processingSummary : '가공 없음';
      detail = `${size}<br><span style="font-size:12px;color:#666;">${proc}</span>`;
    }

    const bgStyle = idx % 2 === 1 ? 'background:#fafafa;' : '';
    return `
      <tr>
        <td style="padding:10px;border:1px solid #ddd;${bgStyle}">${escapeHtml(name)}</td>
        <td style="padding:10px;border:1px solid #ddd;${bgStyle}">${detail}</td>
        <td style="padding:10px;border:1px solid #ddd;text-align:center;${bgStyle}">${qty}</td>
        <td style="padding:10px;border:1px solid #ddd;text-align:right;${bgStyle}">${estimate.toLocaleString()}원</td>
      </tr>`;
  }).join('');
}

function escapeHtml(s) {
  return String(s || '').replace(/[&<>"']/g, (c) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
  }[c]));
}

/**
 * 주문 확인 이메일 발송.
 * @param {Object} params
 * @param {string} params.to - 수신자 이메일
 * @param {string} params.customerName
 * @param {string} params.orderNo
 * @param {Array}  params.items - 주문 항목 배열
 * @param {number} params.shippingCost
 * @param {number} params.totalPrice
 * @returns {Promise<{ok:boolean, messageId?:string, error?:string}>}
 */
async function sendOrderEmail({ to, customerName, orderNo, items, shippingCost, totalPrice }) {
  if (!to) {
    return { ok: false, error: '수신 이메일이 없습니다' };
  }

  const total = Number(totalPrice) || 0;
  const ship = Number(shippingCost) || 0;
  const priceStr = total.toLocaleString() + '원';
  const qtyForPayment = Math.max(1, Math.ceil(total / 1000));
  const itemsHtml = itemsToTableRows(items);

  const html = `
<div style="max-width:650px;margin:0 auto;font-family:'Malgun Gothic',sans-serif;color:#333;">
  <h2 style="color:#5d4037;border-bottom:2px solid #5d4037;padding-bottom:10px;">아이린가구 주문 접수 완료</h2>

  <p><strong>${escapeHtml(customerName)}</strong>님, 주문해 주셔서 감사합니다.</p>
  <p style="color:#666;font-size:13px;">주문번호: <strong style="color:#5d4037;">${escapeHtml(orderNo)}</strong></p>

  <!-- 견적 내역서 -->
  <h3 style="margin-top:25px;color:#5d4037;">📋 주문 내역</h3>
  <table style="width:100%;border-collapse:collapse;margin:10px 0;font-size:14px;">
    <thead>
      <tr style="background:#5d4037;color:#fff;">
        <th style="padding:10px;border:1px solid #5d4037;text-align:left;">품목</th>
        <th style="padding:10px;border:1px solid #5d4037;text-align:left;">상세</th>
        <th style="padding:10px;border:1px solid #5d4037;text-align:center;width:60px;">수량</th>
        <th style="padding:10px;border:1px solid #5d4037;text-align:right;width:120px;">금액</th>
      </tr>
    </thead>
    <tbody>
      ${itemsHtml}
    </tbody>
    <tfoot>
      <tr style="background:#f5f5f5;font-weight:bold;">
        <td colspan="3" style="padding:12px;border:1px solid #ddd;text-align:right;">총 합계</td>
        <td style="padding:12px;border:1px solid #ddd;text-align:right;color:#c62828;font-size:16px;">${priceStr}</td>
      </tr>
    </tfoot>
  </table>

  <!-- 결제 방법 -->
  <div style="background:#fff3e0;border-left:4px solid #ff9800;padding:18px;margin:20px 0;">
    <h3 style="margin:0 0 12px;color:#e65100;">❗❗❗ 결제 방법 ❗❗❗</h3>
    <ol style="margin:0;padding-left:20px;line-height:1.8;">
      <li>상품 수량을 조절하여 <strong>견적 합계(${priceStr})</strong>와 금액을 맞춰주세요.
        <br><span style="color:#666;font-size:13px;">예: ${qtyForPayment}개 = ${priceStr} (1개 = 1,000원)</span></li>
      <li>'배송 메시지'란에 <strong>주문번호(${escapeHtml(orderNo)})</strong>를 입력해 주세요.
        <br><span style="color:#999;font-size:12px;">(주문번호를 잊으신 경우 생략하셔도 괜찮습니다)</span></li>
      <li>결제를 완료하면 주문이 정상적으로 접수됩니다.</li>
    </ol>
  </div>

  <!-- 주문 내역 확인 -->
  <div style="background:#e8f5e9;border-left:4px solid #4caf50;padding:15px;margin:20px 0;">
    <p style="margin:0 0 8px;"><strong>👉 주문 내역 다시 확인하기</strong></p>
    <p style="margin:0;font-size:13px;color:#555;">아래 페이지에서 요청하신 상세 내역을 언제든지 확인하실 수 있습니다.</p>
    <p style="margin:8px 0 0;"><a href="${ORDER_CHECK_URL}" style="color:#2e7d32;">${ORDER_CHECK_URL}</a></p>
  </div>

  <!-- 결제 사이트 -->
  <div style="background:#e3f2fd;border-left:4px solid #2196f3;padding:15px;margin:20px 0;">
    <p style="margin:0 0 10px;"><strong>👉 결제 사이트</strong> <span style="font-size:13px;color:#666;">(원하시는 곳에서 결제 진행)</span></p>
    <p style="margin:6px 0;">● <strong>스토어팜</strong><br><a href="https://naver.me/GJ5OEouG" style="color:#1565c0;">https://naver.me/GJ5OEouG</a></p>
    <p style="margin:6px 0;">● <strong>쿠팡</strong><br><a href="https://link.coupang.com/a/cQ9keF" style="color:#1565c0;">https://link.coupang.com/a/cQ9keF</a></p>
  </div>

  <!-- 문의 -->
  <div style="background:#f5f5f5;padding:15px;margin:20px 0;border-radius:5px;">
    <p style="margin:5px 0;">📞 문의: <strong>031-541-4959</strong></p>
    <p style="margin:5px 0;">💬 카카오톡: <a href="http://pf.kakao.com/_NKxaxcl/chat">아이린가구 채널</a></p>
  </div>

  <p style="text-align:center;color:#999;font-size:12px;margin-top:30px;">본 메일은 발신전용입니다. 문의는 위 연락처로 부탁드립니다.</p>
</div>`;

  // 인증 실패 시 캐시 무효화 후 1회 재시도
  for (let attempt = 0; attempt < 2; attempt++) {
    try {
      const transporter = await getTransporter();
      const creds = await getCredentials();
      const info = await transporter.sendMail({
        from: creds.from,
        to,
        subject: `[아이린가구] 주문 접수 완료 - ${orderNo}`,
        html,
      });
      console.log('[mail] 이메일 발송 성공:', info.messageId, '→', to);
      return { ok: true, messageId: info.messageId };
    } catch (e) {
      console.error(`[mail] 이메일 발송 실패 (시도 ${attempt + 1}/2):`, e.message);
      // 인증 관련 에러면 캐시 무효화 (DB에서 최신 비밀번호 재조회)
      if (/auth|EAUTH|invalid user|password|인증/i.test(e.message)) {
        console.warn('[mail] 인증 실패 - DB에서 자격증명 재조회');
        invalidateCache();
      } else {
        // 인증 외 에러면 재시도 의미 없음
        break;
      }
    }
  }
  return { ok: false, error: '이메일 발송 실패' };
}

module.exports = { sendOrderEmail };
