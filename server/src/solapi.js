// ============================================================
// 솔라피(SolAPI) 클라이언트 - SDK 기반 (SMS / 알림톡 발송)
// ------------------------------------------------------------
// 직접 fetch 대신 공식 SDK(solapi v6) 사용.
// SDK가 인증, date 필드, agent 필드 등을 자동 처리.
//
// 환경변수:
//   SOLAPI_API_KEY, SOLAPI_API_SECRET
//   SOLAPI_SENDER          - 발신번호
//   SOLAPI_KAKAO_PF_ID     - 알림톡 pfId (카카오 채널 프로필 ID)
//   SOLAPI_KAKAO_TEMPLATE_CODE - 승인된 알림톡 템플릿 코드
//
// 문서: https://solapi.com/developers/
// ============================================================
const Solapi = require('solapi');
const config = require('./config');

let _service = null;
function getService() {
  if (_service) return _service;
  const { SOLAPI_API_KEY, SOLAPI_API_SECRET } = config.solapi;
  if (!SOLAPI_API_KEY || !SOLAPI_API_SECRET) {
    throw new Error('솔라피 API 키가 설정되지 않았습니다 (.env의 SOLAPI_API_KEY/SOLAPI_API_SECRET)');
  }
  _service = new Solapi.SolapiMessageService(SOLAPI_API_KEY, SOLAPI_API_SECRET);
  return _service;
}

/**
 * 주문 완료 알림 발송 (알림톡 1순위, 실패 시 SDK가 SMS 자동 대체).
 *
 * @param {Object} params
 * @param {string} params.to - 수신번호 (하이픈 허용)
 * @param {string} params.text - SMS 대체발송용 본문
 * @param {Object} params.templateVars - 알림톡 변수 (키는 #{} 없이, 내부에서 붙임)
 * @returns {Promise<{ok:boolean, channel:string, data?:any, error?:string}>}
 */
async function sendOrderNotification({ to, text, templateVars }) {
  const to0 = String(to || '').replace(/[^0-9]/g, '');
  if (!to0) {
    return { ok: false, channel: 'none', error: '수신번호가 없습니다' };
  }

  const { SOLAPI_KAKAO_PF_ID, SOLAPI_KAKAO_TEMPLATE_CODE, SOLAPI_SENDER } = config.solapi;
  const service = getService();

  // 알림톡 설정이 있으면 알림톡 발송 (disableSms:false → 실패 시 SMS 자동 대체)
  if (SOLAPI_KAKAO_PF_ID && SOLAPI_KAKAO_TEMPLATE_CODE) {
    try {
      // variables 키에 #{} 붙이기 (공식 형식)
      const variables = {};
      for (const [k, v] of Object.entries(templateVars || {})) {
        const key = k.startsWith('#{') ? k : `#{${k}}`;
        variables[key] = String(v);
      }

      const message = {
        to: to0,
        from: SOLAPI_SENDER,
        text: text || '',
        kakaoOptions: {
          pfId: SOLAPI_KAKAO_PF_ID,
          templateId: SOLAPI_KAKAO_TEMPLATE_CODE,
          variables,
          disableSms: false, // 알림톡 실패 시 SMS 자동 대체 (솔라피가 처리)
        },
      };

      const data = await service.send([message]);
      console.log('[solapi] 알림톡 발송 성공:', JSON.stringify(data).slice(0, 200));
      return { ok: true, channel: 'kakao', data };
    } catch (e) {
      console.warn('[solapi] 알림톡 발송 실패, 수동 SMS 폴백:', e.message);
    }
  }

  // SMS 폴백
  try {
    const type = text && text.length > 45 ? 'LMS' : 'SMS';
    const message = {
      to: to0,
      from: SOLAPI_SENDER,
      text: text || '',
      type,
    };
    const data = await service.send([message]);
    console.log('[solapi] SMS 발송 성공:', JSON.stringify(data).slice(0, 200));
    return { ok: true, channel: 'sms', data };
  } catch (e) {
    console.error('[solapi] SMS 발송 실패:', e.message);
    return { ok: false, channel: 'sms', error: e.message };
  }
}

module.exports = { sendOrderNotification };
