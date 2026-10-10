// ============================================================
// POST /api/chat - AI 채팅 (Gemini function calling)
// ------------------------------------------------------------
// 요청: { message, sessionId, state }
// 응답: { result: 'success', reply?, tool_calls? }
// ============================================================
const express = require('express');
const router = express.Router();
const gemini = require('../gemini');

router.post('/', async (req, res) => {
  const { message, sessionId, state, history } = req.body || {};

  // 임시 디버그 로그 (문제 파악 후 제거)
  console.log('[chat] 요청 수신:', { message: (message || '').slice(0, 50), sessionId, hasHistory: !!history });

  if (!message || typeof message !== 'string') {
    return res.status(400).json({ result: 'error', message: 'message 가 필요합니다.' });
  }

  try {
    const out = await gemini.chat({
      message,
      sessionId: sessionId || 'unknown',
      state: state || {},
      history: Array.isArray(history) ? history : [],
    });

    return res.json({
      result: 'success',
      reply: out.reply,
      tool_calls: out.tool_calls,
    });
  } catch (err) {
    console.error('[chat] Gemini 호출 오류:', err);
    return res.status(500).json({
      result: 'error',
      message: 'AI 응답 생성 중 오류가 발생했습니다.',
      detail: process.env.NODE_ENV === 'production' ? undefined : err.message,
    });
  }
});

module.exports = router;
