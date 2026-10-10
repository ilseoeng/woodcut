// ============================================================
// Gemini 클라이언트 + Function-Calling Agentic 루프
// ------------------------------------------------------------
// 패키지: @google/generative-ai
// 흐름 (Agentic):
//   1) 사용자 메시지 + state + history + functionDeclarations 를 Gemini에 전송
//   2) Gemini가 functionCall을 반환하면:
//      a) 백엔드 도구(get_quote 등) → 백엔드에서 실행 → functionResponse로 회신 → 1)로 다시
//      b) 프론트 도구(select_wood 등) → tool_calls로 프론트에 전달 (실행은 프론트)
//   3) 최종 텍스트 응답이 오면 종료
//
// 주의: @google/generative-ai 의 Schema type 값은 대문자여야 한다.
// ============================================================
const { GoogleGenerativeAI } = require('@google/generative-ai');
const config = require('./config');
const { getFunctionDeclarations, BACKEND_TOOL_NAMES } = require('./tools/schema');
const { BACKEND_TOOL_EXECUTORS } = require('./tools/backendTools');

const genAI = new GoogleGenerativeAI(config.gemini.apiKey);

const MAX_AGENTIC_ROUNDS = 5; // 무한 루프 방지

/**
 * Schema 의 type 값을 Gemini가 요구하는 대문자로 정규화한다.
 */
function normalizeSchema(schema) {
  if (!schema || typeof schema !== 'object') return schema;
  const out = { ...schema };
  if (typeof out.type === 'string') out.type = out.type.toUpperCase();
  if (out.properties && typeof out.properties === 'object') {
    const props = {};
    for (const [k, v] of Object.entries(out.properties)) props[k] = normalizeSchema(v);
    out.properties = props;
  }
  if (out.items) out.items = normalizeSchema(out.items);
  return out;
}

let _model = null;
function getModel() {
  if (_model) return _model;
  const functionDeclarations = getFunctionDeclarations().map((fd) => ({
    ...fd,
    parameters: normalizeSchema(fd.parameters),
  }));
  _model = genAI.getGenerativeModel({
    model: config.gemini.model,
    tools: [{ functionDeclarations }],
    systemInstruction: buildSystemInstruction(),
  });
  return _model;
}

function buildSystemInstruction() {
  return [
    '당신은 "아이린가구"의 AI 목재 견적 상담 비서입니다. 친근하고 정직한 한국어로 응답합니다.',
    '',
    '[도구 사용 - 반드시 준수]',
    '- 추천/상담/비교 질문 → 반드시 먼저 도구를 호출한다. AI가 직접 추론하지 마라.',
    '  · "뭘 써야 해?" / "책장 만들 건데" → recommend_wood (함수가 용도 기반으로 객관적 추천)',
    '  · "A랑 B 비교해줘" / "어떤 게 더 싸?" → compare_woods (지정한 목재들 비교)',
    '  · "C얼마야?" / "가격 알려줘" → get_quote (MS-SQL 기본단가 + 추가 가공비 - 이벤트 할인, 추측 금지)',
    '- get_quote 호출 시 고객이 이번 메시지에서 지정한 추가가공은 processing 인수로 반드시 전달한다. 기존 UI 상태가 업데이트되기 전에도 AI가 가공 조건을 전달해야 한다.',
    '- 견적 답변에는 기본단가, 추가 가공비, 이벤트 할인, 수량별 소계(배송비 별도)를 구분한다.',
    '  · "X 장단점 알려줘" / "편백이 뭐가 좋아?" / "목재 특징?" → get_wood_info (장점·단점·주의점 조회)',
    '  · "어떤 목재 있어?" / "전체 목록" → get_wood_info (woodKey 생략, 전체 요약)',
    '- 가공 충돌 점검 → validate_processing.',
    '- 목재 선택·사이즈·가공 토글 등 UI 변경 → select_wood, set_size, toggle_processing 등.',
    '',
    '[범위 외 질문 대응 - 반드시 준수]',
    '- 당신은 목재 재단/견적 전문 비서다. 목재·단가·가공·사이즈·추천·비교 외의 질문은 아래 원칙으로 처리한다.',
    '- 아래 [FAQ]에 있는 정보는 정확하니 그대로 답한다. [FAQ]에 없는 정보는 절대 지어내지 말고 "고객센터로 문의해 주세요"로 답한다.',
    '- 목재와 무관한 일반 질문(날씨·코딩·잡담 등)은 정중히 거절하고 목재 상담으로 유도한다.',
    '',
    '[FAQ - 정확한 정보]',
    '- 배송: 기본 2~4일 소요. 단 코팅 가공 시 7일 소요.',
    '- 결제 방법: 네이버 또는 쿠팡에서 결제. 견적서에 나온 총액을 "수량"으로 환산해 결제한다.',
    '  · 예: 견적 총액 30,000원이면 수량 30개로 결제 (1개=1,000원 취급).',
    '  · 즉 결제 금액 = 견적 금액과 동일.',
    '- 목재 두께: 18T만 취급 (다른 두께 불가).',
    '- 결제·교환·반품·A/S 등 추가 문의는 아래 고객센터로 안내.',
    '',
    '[고객센터 안내]',
    '- 전화: 031-541-4959',
    '- 네이버톡톡: http://talk.naver.com/profile/wc80z0',
    '- 카카오톡: http://pf.kakao.com/_NKxaxcl/chat',
    '',
    '[AI의 역할]',
    '- 도구 결과를 받으면, 그것을 바탕으로 고객에게 자연스럽고 친절하게 설명한다.',
    '- 함수가 준 추천 이유/차이점을 풀어서 설명하고, 필요하면 한 번 더 질문한다.',
    '- 직접 "이 목재가 좋을 것 같아요"라고 추측하지 말고, 함수 결과를 근거로 답한다.',
    '',
    '[규칙] 폭≤600mm, 길이≤1800mm, 최소 10mm. 집성목(미송/편백/멀바우)은 폭>길이 시 파손 위험. 샌딩·코팅은 원목만. 샌딩+코팅 동시 불가.',
    '',
    '[목재]',
    '- mdf18T: 가성비 합판, 표면 매끄럼, 페인트 적합. 방수X, 샌딩/코팅X.',
    '- pvc18T: 완전방수, 욕실/주방/세탁실. 샌딩/코팅X.',
    '- misong18T: 미송집성목, 따뜻한 원목, DIY 인기. 샌딩/코팅O.',
    '- pyeonbaek18T: 편백집성목, 피톤치드/항균, 아이 방·침실 추천. 샌딩/코팅O.',
    '- mulbau18T: 멀바우 하드우드, 단단하고 내구성 강, 테이블 상판. 샌딩/코팅O.',
  ].join('\n');
}

/**
 * 채팅 처리 메인 함수 (Agentic).
 * 서버는 stateless. 대화 이력은 프론트가 localStorage에 저장해 매 요청마다 전송.
 *
 * @param {Object} params
 * @param {string} params.message
 * @param {string} params.sessionId
 * @param {Object} params.state
 * @param {Array} params.history - Gemini 형식 대화 이력
 * @returns {Promise<{reply?, tool_calls?}>}
 */
async function chat({ message, sessionId, state, history }) {
  const model = getModel();

  // history 정제: Gemini는 첫 항목이 반드시 role:'user' 여야 함.
  // 프론트 localStorage에서 온 이력이 model 로 시작하면 앞부분을 자른다.
  // 또한 연속된 같은 role 병합/빈 항목 제거.
  let cleanHistory = Array.isArray(history) ? history.filter((h) => h && h.role && h.parts && h.parts.length) : [];
  while (cleanHistory.length && cleanHistory[0].role !== 'user') {
    cleanHistory.shift();
  }

  const chatSession = model.startChat({
    history: cleanHistory,
    generationConfig: { temperature: 0.4 },
  });

  // 사용자 메시지에 현재 상태 컨텍스트를 함께 제공
  const contextualMessage = [
    `[현재 상태]`,
    JSON.stringify(state || {}, null, 2),
    ``,
    `[고객 메시지]`,
    message,
  ].join('\n');

  // Agentic 루프: 백엔드 도구는 자동 실행·회신, 프론트 도구는 모아서 반환
  const frontToolCalls = [];
  let lastText = '';
  let currentMessage = contextualMessage;

  for (let round = 0; round < MAX_AGENTIC_ROUNDS; round++) {
    const result = await chatSession.sendMessage(currentMessage);
    const response = result.response;
    const functionCalls = response.functionCalls() || [];
    const text = (response.text && response.text()) || '';

    if (text) lastText = text;

    if (functionCalls.length === 0) {
      // 더 이상 도구 호출 없음 → 종료
      break;
    }

    // functionCall 분류: 백엔드 실행 vs 프론트 전달
    const functionResponses = []; // 백엔드 도구 결과 (Gemini에게 회신)
    for (const fc of functionCalls) {
      if (BACKEND_TOOL_NAMES.has(fc.name)) {
        // 백엔드 도구: 여기서 직접 실행
        const executor = BACKEND_TOOL_EXECUTORS[fc.name];
        let toolResult;
        try {
          toolResult = await executor(fc.args || {}, { state, message, sessionId });
        } catch (e) {
          toolResult = { error: String(e && e.message || e) };
        }
        functionResponses.push({
          name: fc.name,
          response: toolResult, // 결과 객체 그대로 (SDK가 {result: ...} 로 래핑해 전송)
        });
      } else {
        // 프론트 도구: 프론트에 전달할 목록에 추가
        frontToolCalls.push({ name: fc.name, args: fc.args || {} });
      }
    }

    // 백엔드 도구 결과가 있으면 Gemini에게 functionResponse로 회신 → 다음 라운드
    if (functionResponses.length) {
      // SDK 형식: sendMessage([{ functionResponse: { name, response } }, ...])
      currentMessage = functionResponses.map((fr) => ({ functionResponse: fr }));
    } else {
      // 백엔드 도구 없이 프론트 도구만 있었으면 루프 종료 (프론트가 실행)
      break;
    }
  }

  return {
    reply: lastText || undefined,
    tool_calls: frontToolCalls.length ? frontToolCalls : undefined,
  };
}

module.exports = { chat };
