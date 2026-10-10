// ============================================================
// Gemini Function Declarations 생성
// ------------------------------------------------------------
// 두 종류의 도구를 합친다:
//   1) 프론트 실행 도구 (TOOL_REGISTRY) - UI 변경. AI가 부르면 프론트에서 실행
//   2) 백엔드 실행 도구 (BACKEND_TOOL_SCHEMAS) - DB 조회. AI가 부르면 백엔드에서 실행
// Gemini에게는 구분 없이 하나의 functionDeclarations 배열로 전달된다.
// ============================================================
const { TOOL_REGISTRY } = require('./registry');
const { BACKEND_TOOL_SCHEMAS, BACKEND_TOOL_EXECUTORS } = require('./backendTools');

// 백엔드에서 실행되는 도구 이름 집합 (Agentic 루프에서 사용)
const BACKEND_TOOL_NAMES = new Set(Object.keys(BACKEND_TOOL_EXECUTORS));

/**
 * @returns {Array<{name, description, parameters}>}
 *   Gemini SDK 의 tools: [{ functionDeclarations: [...] }] 형식에 사용.
 */
function getFunctionDeclarations() {
  // 프론트 실행 도구
  const front = Object.entries(TOOL_REGISTRY).map(([name, meta]) => ({
    name,
    description: meta.description,
    parameters: meta.parameters || { type: 'object', properties: {} },
  }));
  // 백엔드 실행 도구
  const back = Object.entries(BACKEND_TOOL_SCHEMAS).map(([name, meta]) => ({
    name,
    description: meta.description,
    parameters: meta.parameters || { type: 'object', properties: {} },
  }));
  return [...front, ...back];
}

/**
 * Gemini SDK용 tools 객체 반환.
 * @returns {{ tools: { functionDeclarations: Array } }}
 */
function getGeminiTools() {
  return {
    tools: {
      functionDeclarations: getFunctionDeclarations(),
    },
  };
}

module.exports = {
  getFunctionDeclarations,
  getGeminiTools,
  BACKEND_TOOL_NAMES,
};
