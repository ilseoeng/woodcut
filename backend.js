/**
 * @OnlyCurrentDoc
 */

/* ==================== 빌드/일반 설정 ==================== */
const BUILD_VERSION = '2025-02-13+FINAL-FIXED-COMPATIBLE';
const MAX_WIDTH_MM = 600;
const MAX_LENGTH_MM = 1800;
const SHIPPING_MODE = 'per-order'; // 'per-order' | 'per-unit'
const CACHE_TTL_SEC = 60 * 30;

/* ==================== 스프레드시트 설정 ==================== */
const PRICE_SHEET_ID = '1HSo0_5e5rbxFUGXVAavXZoDpzKKEVercB1Nd52GXGLw';
const PRICE_SHEET_NAME = '단가표';
const ORDER_SHEET_ID = '1J9Q2UybMsx6Kv_V7q2LXt1pbewtUF7TRQWcBymOrS9A';
const ORDER_SHEET_NAME = '견적요청';

/* ==================== SOLAPI(카카오 알림톡/LMS) 설정 ==================== */
const SOLAPI_API_KEY = 'NCSTLNWD19QCQIDK';
const SOLAPI_API_SECRET = '9ZMXYDBKJGMTYCF2MTDXGE2DUZTZ0SHY';
const SENDER_PHONE = '01029488203';
const KAKAO_PF_ID = 'KA01PF250717222607219jkItrsz0ptP';
const TEMPLATE_ID = 'KA01TP250916040322542jg37R2XFpbe';
const LMS_SUBJECT = '주문 안내';

/* ==================== OpenAI 설정 ==================== */
const OPENAI_API_KEY_TEMP = '';

function getOpenAiApiKey() {
    const k = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    if (k && k.startsWith('sk-')) return k;
    if (OPENAI_API_KEY_TEMP && OPENAI_API_KEY_TEMP.startsWith('sk-')) return OPENAI_API_KEY_TEMP;
    throw new Error('OpenAI 키 없음: setOpenAIKey()로 OPENAI_API_KEY 저장 필요');
}
function setOpenAIKey() { const KEY = 'sk-REPLACE_WITH_REAL_KEY'; PropertiesService.getScriptProperties().setProperty('OPENAI_API_KEY', KEY); }

/* ==================== 유틸리티 ==================== */
function L(rid, ...args) { Logger.log(`[RID ${rid}] ` + args.map(a => typeof a === 'object' ? JSON.stringify(a) : String(a)).join(' ')); }
function jsonOut(o) { return ContentService.createTextOutput(JSON.stringify(o || {})).setMimeType(ContentService.MimeType.JSON); }
function jsOut(s) { return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JAVASCRIPT); }
function _num(v) { if (v == null) return null; const s = String(v).trim().replace(/[, ]/g, ''); if (s === '') return null; const n = Number(s); return isNaN(n) ? null : n; }
function _cache() { return CacheService.getScriptCache(); }

/* ==================== 히스토리 관리 ==================== */
function loadHistory(cid) { if (!cid) return []; try { const v = _cache().get('chatv4:' + cid); return v ? JSON.parse(v) : []; } catch (e) { return []; } }
function saveHistory(cid, hist) { if (!cid) return; try { _cache().put('chatv4:' + cid, JSON.stringify((Array.isArray(hist) ? hist : []).slice(-20)), CACHE_TTL_SEC); } catch (e) { } }
function clearChatHistory(cid) { try { _cache().put('chatv4:' + cid, '[]', 1); } catch (e) { } }

function getWoodTypeName(v) {
    const map = { mdf18T: 'MDF 18T', pvc18T: 'PVC보드 18T+PET', misong18T: '미송집성목 18T', pyeonbaek18T: '편백집성목 18T', mulbau18T: '멀바우 집성목 18T' };
    return map[v] || v;
}

function generateOrderNumber(sheet) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
        const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
        const last = sheet.getLastRow(); if (last < 2) return `${today}-001`;
        const colA = sheet.getRange(2, 1, last - 1, 1).getValues();
        let max = 0;
        colA.forEach(r => { const v = String(r[0] || ''); if (v.startsWith(today)) { const n = parseInt(v.split('-')[1] || '0', 10); if (!isNaN(n) && n > max) max = n; } });
        return `${today}-${String(max + 1).padStart(3, '0')}`;
    } finally { lock.releaseLock(); }
}

/* ==================== 단가표 및 가격 계산 ==================== */
function readPriceTableSafe() {
    const ss = SpreadsheetApp.openById(PRICE_SHEET_ID);
    const sh = ss.getSheetByName(PRICE_SHEET_NAME);
    const vals = sh.getDataRange().getDisplayValues();
    const headers = vals[0].map(h => (h || '').trim());
    const rows = [];
    for (let r = 1; r < vals.length; r++) {
        const line = vals[r]; if (!line || line.join('').trim() === '') continue;
        const obj = {};
        headers.forEach((h, c) => {
            if (!h) return;
            const key = h.trim(); const raw = (line[c] || '').trim();
            if (key === 'area' || key === '배송비' || /18T$/i.test(key) || /단가|가격|fee|cost|price/i.test(key)) obj[key] = _num(raw);
            else obj[key] = raw === '' ? null : raw;
        });
        if (obj.area != null) rows.push(obj);
    }
    rows.sort((a, b) => (a.area || 0) - (b.area || 0));
    return { headers, rows };
}

function pickRowByArea(rows, area) {
    rows.sort((a, b) => (a.area || 0) - (b.area || 0));
    for (let i = 0; i < rows.length; i++) { if (area <= rows[i].area) return rows[i]; }
    return rows.length ? rows[rows.length - 1] : null;
}

function calculatePriceFromSheet(order, rid) {
    const table = readPriceTableSafe();
    const matKey = String(order.woodTypeValue).trim();
    const area = (Number(order.width) || 0) * (Number(order.length) || 0);
    const row = pickRowByArea(table.rows, area);
    if (!row) throw new Error('단가표에서 해당 면적을 찾을 수 없습니다.');
    const unitPrice = Number(row[matKey]) || 0;
    const shippingRaw = Number(row['배송비'] || 0);
    const qty = Number(order.quantity) || 1;
    const shipping = SHIPPING_MODE === 'per-unit' ? (shippingRaw * qty) : shippingRaw;
    return { total: (unitPrice * qty) + shipping, unitPrice, shipping, subtotal: unitPrice * qty };
}

function computeProcessingCost(order) {
    const mat = String(order.woodTypeValue || '').toLowerCase();
    const isMdfOrPvc = (mat === 'mdf18t' || mat === 'pvc18t');
    const qty = Number(order.quantity) || 1;
    const len = Number(order.length) || 0;

    const opts = normalizeOptions(order.options || []);
    let processingPerUnit = 0;
    const lines = [];

    if (!isMdfOrPvc) {
        if (opts.coating_single > 0) { const u = (len <= 600 ? 5000 : len <= 1200 ? 10000 : 15000); processingPerUnit += u; lines.push(`코팅(단면)`); }
        if (opts.coating_double > 0) { const u = (len <= 600 ? 10000 : len <= 1200 ? 20000 : 30000); processingPerUnit += u; lines.push(`코팅(양면)`); }
    }

    const FREE = { angled: 1, round: 1, circlecut: 1, squarecut: 1, hinge: 2 };
    const UNIT = { angled: 1000, round: 1000, circlecut: 3000, squarecut: 4000, hinge: 1000 };

    ['angled', 'round', 'circlecut', 'squarecut', 'hinge'].forEach(k => {
        const count = Math.max(0, Number(opts[k] || 0));
        if (count <= 0) return;
        const free = FREE[k] || 0;
        const paid = Math.max(0, count - free);
        if (paid > 0) processingPerUnit += (paid * UNIT[k]);
    });

    return { processingPerUnit, processingTotal: processingPerUnit * qty };
}

function normalizeOptions(arr) {
    const sum = { none: 0, sanding: 0, coating_single: 0, coating_double: 0, angled: 0, round: 0, circlecut: 0, squarecut: 0, hinge: 0 };
    if (!Array.isArray(arr)) return sum;
    arr.forEach(o => {
        if (!o) return;
        let type = (typeof o === 'string') ? o : String(o.type || o.key || '');
        type = type.toLowerCase().trim();
        const count = Math.max(0, Number((typeof o === 'string') ? 1 : o.count || o.qty || 1));
        if (type === 'none' || /없음/.test(type)) sum.none += count;
        else if (type === 'sanding') sum.sanding += count;
        else if (type === 'coating' || type === 'coating_single' || /단면/.test(type)) sum.coating_single += count;
        else if (type === 'coating_double' || /양면/.test(type)) sum.coating_double += count;
        else if (type === 'angled' || /사선/.test(type)) sum.angled += count;
        else if (type === 'round' || /라운드/.test(type)) sum.round += count;
        else if (type === 'circlecut' || /원형/.test(type)) sum.circlecut += count;
        else if (type === 'squarecut' || /사각/.test(type)) sum.squarecut += count;
        else if (type === 'hinge' || /경첩/.test(type)) sum.hinge += count;
    });
    return sum;
}

/* ==================== AI 채팅 최적화 로직 ==================== */
function buildChatSystemPrompt() {
    return `당신은 "아이린가구" 목재 재단 상담 AI입니다. 한국어로 친절히 답하세요.

[목재 5종 - 두께 18T 고정]
- mdf18T: MDF. 가성비. 습기취약. 샌딩/코팅 불가.
- pvc18T: PVC보드. 완전방수. 욕실/주방. 샌딩/코팅 불가.
- misong18T: 미송집성목. DIY표준. 초보 추천.
- pyeonbaek18T: 편백집성목. 항균/향기. 아이가구 추천.
- mulbau18T: 멀바우. 매우 단단. 고급 가구.

[사이즈 규칙]
- "NxM" 또는 "N*M" 입력 시: N=폭(width), M=길이(length)
- 폭 허용: 10~600mm, 길이 허용: 10~1800mm
- 범위 안이면 OK. 초과 시에만 수정 요청.

[추천]
- 아이가구 → 편백 or 미송
- 욕실/주방 → PVC보드
- 고급/단단한 → 멀바우

[대화 규칙]
- 이전에 목재/사이즈를 이미 말했으면 다시 묻지 않기.
- 정보 수집 순서: 목재 → 사이즈 → 수량 → show_processing
- 수량을 고객이 직접 말하지 않았으면 반드시 물어보기. 기본값 사용 금지.

[JSON 응답 형식 - 반드시 이 형식만 반환]
{ "reply": "설명(HTML OK)", "action": "액션명", "data": { "woodType":"키", "width":숫자, "length":숫자, "qty":숫자 } }

[action 종류]
- "chat": 일반 대화
- "show_woods": 목재 카드 표시
- "show_processing": 목재+유효사이즈+수량 모두 확인됨 → 가공옵션 표시. data 필수.
- "need_info": 정보 부족 또는 사이즈 초과 시

[예시]
입력: "mdf 400*200 3개" → action:"show_processing", data:{woodType:"mdf18T",width:400,length:200,qty:3}
입력: "800*500" → 폭 800>600이므로 거부, action:"need_info"
입력: "미송 300x600" → 수량 미확인, action:"need_info", 수량 질문`;
}

function handleChatRequest(message, rid, cid) {
    const apiKey = getOpenAiApiKey();
    let history = loadHistory(cid);
    const messages = [{ role: 'system', content: buildChatSystemPrompt() }];
    history.forEach(h => messages.push(h));
    messages.push({ role: 'user', content: message });

    const res = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
        method: 'post',
        contentType: 'application/json',
        headers: { 'Authorization': 'Bearer ' + apiKey },
        payload: JSON.stringify({ model: 'gpt-4o-mini', temperature: 0.3, response_format: { type: 'json_object' }, messages }),
        muteHttpExceptions: true
    });

    const json = JSON.parse(res.getContentText());
    const assistantMsg = json.choices[0].message.content;
    const result = JSON.parse(assistantMsg);

    history.push({ role: 'user', content: message });
    history.push({ role: 'assistant', content: assistantMsg });
    saveHistory(cid, history);

    if (result.action === 'estimate' && result.data && result.data.woodType) {
        try {
            const d = result.data;
            const order = { woodTypeValue: d.woodType, width: Number(d.width) || 0, length: Number(d.length) || 0, quantity: Number(d.qty) || 1, options: d.options || [] };
            const calc = calculatePriceFromSheet(order, rid);
            const proc = computeProcessingCost(order);
            result.data.unitPrice = calc.unitPrice + proc.processingPerUnit;
            result.data.totalPrice = calc.subtotal + proc.processingTotal;
            result.data.shippingEstimate = calc.shipping;
        } catch (e) { }
    }
    return { result: 'ok', ...result };
}

/* ==================== 주문 저장 (기존 doPostMessaging 등 하위 호환성 유지) ==================== */
function doPostMessaging(data) {
    const rid = Utilities.getUuid().slice(0, 8);
    try {
        const payload = (typeof data === 'string') ? JSON.parse(data) : data;
        const customer = payload.customer || {};
        const orders = payload.orders || [];
        if (!orders.length) return { result: 'error', message: '주문 항목이 없습니다.' };

        const ss = SpreadsheetApp.openById(ORDER_SHEET_ID);
        let sh = ss.getSheetByName(ORDER_SHEET_NAME);
        if (!sh) {
            sh = ss.insertSheet(ORDER_SHEET_NAME);
            sh.appendRow(['주문번호', '날짜', '성함', '전화번호', '이메일주소', '품명', '폭(mm)', '길이(mm)', '수량', '가공상세', '합계']);
        }

        const orderNumber = generateOrderNumber(sh);
        const ts = new Date();
        let totalAmount = 0;

        orders.forEach(order => {
            let rowData = null;
            totalAmount += Number(order.estimate) || 0;

            if (order.type === 'discount') {
                rowData = [orderNumber, ts, customer.name || '', customer.phone || '', customer.email || '', '[이벤트 할인]', '-', '-', '-', '', (Number(order.estimate) || 0)];
            } else if (order.type === 'accessory') {
                rowData = [orderNumber, ts, customer.name || '', customer.phone || '', customer.email || '', '[부자재] ' + (order.name || ''), '-', '-', (order.quantity || ''), '-', (Number(order.estimate) || 0)];
            } else {
                let processingDetails = '가공 없음';
                if (!order.noOption && order.detailText && Object.keys(order.detailText).length > 0) {
                    const summaryArray = [];
                    Object.values(order.detailText).forEach(detail => {
                        if (detail && Array.isArray(detail.summary)) detail.summary.forEach(line => summaryArray.push(String(line).replace(/<[^>]*>/g, '')));
                        else if (detail && typeof detail.summary === 'string') summaryArray.push(detail.summary);
                    });
                    if (summaryArray.length > 0) processingDetails = summaryArray.join(' | ');
                }
                const itemName = getWoodTypeName(order.woodTypeValue) || order.woodTypeValue || '[기타]';
                rowData = [orderNumber, ts, (customer.name || ''), customer.phone || '', customer.email || '', itemName, order.width || '-', order.length || '-', order.quantity || '-', processingDetails, order.estimate || 0];
            }
            if (rowData) sh.appendRow(rowData);
        });

        if (Array.isArray(payload.shippingItems)) {
            payload.shippingItems.forEach(item => {
                const shipTotal = Number(item.cost || 0) * Number(item.count || 0);
                totalAmount += shipTotal;
                sh.appendRow([orderNumber, ts, customer.name || '', customer.phone || '', customer.email || '', '예상 배송비', '-', '-', item.count || 0, (item.sizeCategory || ''), shipTotal]);
            });
        }

        SpreadsheetApp.flush();
        try { sendOrderNotifications(customer, orderNumber, totalAmount, payload.quoteImage); } catch (e) { }
        return { result: 'success', orderNo: orderNumber, totalAmount, rid };
    } catch (err) { return { result: 'error', message: err.toString(), rid }; }
}

function sendOrderNotifications(customer, orderNo, total, quoteImage) {
    const vars = { '#{고객명}': customer.name || '', '#{주문번호}': orderNo, '#{합계}': Number(total || 0).toLocaleString() };
    try { sendAlimtalk(customer.phone, vars); } catch (e) { }
    if (customer.email) try { sendOrderConfirmationEmail(customer.email, customer.name, orderNo, total, quoteImage); } catch (e) { }
}

/* ==================== HTTP 핸들러 (doGet, doPost) ==================== */
function doGet(e) {
    const cb = e.parameter.callback;
    const phone = e.parameter.phone;
    const orderNo = e.parameter.orderNo;
    if (phone && orderNo) {
        const res = findOrdersByPhoneAndOrderNo(phone, orderNo);
        return jsonpOutput(res, cb);
    }
    return jsonpOutput({ result: 'error', message: '파라미터 부족' }, cb);
}

function doPost(e) {
    const rid = Utilities.getUuid().slice(0, 8);
    const p = e.parameter || {};

    // 상황 1. 기존 current-woodcut.html 등의 일반 폼 전송 (data 파라미터)
    if (p.data) {
        return jsonOut(doPostMessaging(p.data));
    }

    // 상황 2. 신규 채팅 woodcut.html 등의 JSON 전송
    let bodyText = (e.postData && e.postData.contents) ? e.postData.contents.toString() : '';
    if (bodyText.startsWith('{')) {
        const body = JSON.parse(bodyText);
        const cid = body.cid || 'anon';
        if (body.type === 'chat') {
            if (body.message === 'reset') { clearChatHistory(cid); return jsonOut({ result: 'ok', reply: '대화가 초기화되었습니다.' }); }
            return jsonOut(handleChatRequest(body.message, rid, cid));
        }
        if (body.type === 'order') {
            return jsonOut(doPostMessaging(body.data));
        }
    }

    // 상황 3. 기타 텍스트 전송 (cid: 포함 가능)
    if (bodyText) {
        const cidMatch = bodyText.match(/\[cid:([^\]]+)\]/);
        let cid = cidMatch ? cidMatch[1] : 'anon';
        let cleanMsg = bodyText.replace(/\[cid:[^\]]+\]/, '').trim();
        return jsonOut(handleChatRequest(cleanMsg, rid, cid));
    }

    return jsonOut({ result: 'error', message: '처리할 수 없는 요청입니다.' });
}

/* ==================== 주문 조회 및 알림 (나머지 원복) ==================== */
function findOrdersByPhoneAndOrderNo(phone, orderNo) {
    const ss = SpreadsheetApp.openById(ORDER_SHEET_ID);
    const sh = ss.getSheetByName(ORDER_SHEET_NAME);
    const vals = sh.getDataRange().getValues();
    const headers = vals[0];
    const out = [];
    const cleanPhone = String(phone).replace(/-/g, '');
    for (let r = 1; r < vals.length; r++) {
        if (String(vals[r][0]) === String(orderNo) && String(vals[r][3]).replace(/-/g, '') === cleanPhone) {
            const obj = {};
            headers.forEach((h, i) => obj[h] = vals[r][i]);
            out.push(obj);
        }
    }
    return { result: 'success', data: out };
}

function jsonpOutput(obj, cb) {
    const json = JSON.stringify(obj);
    return cb ? jsOut(`${cb}(${json});`) : jsonOut(obj);
}

function sendAlimtalk(receiverPhone, variables) {
    const apiUrl = 'https://api.solapi.com/messages/v4/send';
    const now = new Date().toISOString();
    const salt = Utilities.getUuid();
    let sig = Utilities.computeHmacSha256Signature(now + salt, SOLAPI_API_SECRET);
    sig = sig.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
    const payload = { message: { to: String(receiverPhone).replace(/-/g, ''), kakaoOptions: { pfId: KAKAO_PF_ID, templateId: TEMPLATE_ID, variables } } };
    const options = { method: 'post', contentType: 'application/json', headers: { 'Authorization': `HMAC-SHA256 ApiKey=${SOLAPI_API_KEY}, Date=${now}, salt=${salt}, signature=${sig}` }, payload: JSON.stringify(payload), muteHttpExceptions: true };
    UrlFetchApp.fetch(apiUrl, options);
}

function sendOrderConfirmationEmail(email, name, orderNo, total, quoteImage) {
    const subject = `[아이린가구] ${name}님의 목재 견적 내역서가 도착했습니다. (${orderNo})`;
    const formattedTotal = Number(total).toLocaleString();

    let body = `안녕하세요, 아이린가구 목재재단 서비스입니다.\n` +
        `요청하신 견적의 주문 정보를 안내해 드립니다.\n\n` +
        `■ 주문 정보\n` +
        `고객명: ${name}\n` +
        `주문번호: ${orderNo}\n` +
        `최종 견적 합계: ${formattedTotal}원\n\n` +
        `■ 결제 방법\n` +
        `상품 수량을 조절하여 견적 합계(${formattedTotal}원)와 금액을 맞춰주세요.\n` +
        `(예: ${formattedTotal}원 → 수량 ${Number(total / 1000).toFixed(0)}개)\n` +
        `'배송 메시지'란에 주문번호(${orderNo})를 입력해주세요.\n` +
        `(주문번호를 잊으신 경우 생략 가능)\n` +
        `결제를 완료하면 주문이 정상 접수됩니다.\n\n` +
        `■ 결제 사이트\n` +
        `(원하시는 곳에서 결제 진행)\n` +
        `● 네이버 스토어 : https://naver.me/GJ5OEouG\n` +
        `● 쿠팡 : https://link.coupang.com/a/cQ9keF\n\n` +
        `상세 내역은 첨부된 이미지 파일을 확인해 주세요. 감사합니다.`;

    const options = {
        name: "아이린가구",
        attachments: []
    };

    if (quoteImage && quoteImage.indexOf("base64,") > -1) {
        try {
            const byteData = Utilities.base64Decode(quoteImage.split(",")[1]);
            const blob = Utilities.newBlob(byteData, "image/jpeg", "아이린가구_견적서_" + orderNo + ".jpg");
            options.attachments.push(blob);
        } catch (e) {
            body += "\n\n(※ 이미지 첨부 중 오류가 발생하여 텍스트로만 발송되었습니다. 고객센터로 문의 부탁드립니다.)";
        }
    }

    MailApp.sendEmail(email, subject, body, options);
}
