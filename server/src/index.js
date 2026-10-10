// ============================================================
// Express 진입점
// ------------------------------------------------------------
// - CORS 허용 (개발: 전체, 운영: CORS_ORIGIN 화이트리스트)
// - JSON 바디 파싱 (견적서 base64 이미지가 크므로 limit 상향)
// - /api/* 라우트 마운트
// - /uploads 정적 서빙 (저장된 견적서 이미지)
// - 루트 정적 호스팅 (woodcut_vm.html 서빙용)
// - 헬스체크 /healthz
// ============================================================
const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const config = require('./config');

const app = express();

// ---- CORS ----
const corsOptions = {};
if (config.corsOrigin) {
  corsOptions.origin = config.corsOrigin;
  corsOptions.credentials = true;
}
app.use(cors(corsOptions));

// ---- 바디 파싱 ----
// 견적서 이미지(base64 JPEG)가 포함될 수 있으므로 제한을 넉넉히.
app.use(express.json({ limit: '25mb' }));
app.use(express.urlencoded({ extended: true, limit: '25mb' }));

// ---- 헬스체크 ----
app.get('/healthz', async (req, res) => {
  let dbOk = false;
  try {
    const { ping } = require('./db');
    await ping();
    dbOk = true;
  } catch {
    dbOk = false;
  }
  res.json({ ok: true, db: dbOk, uptime: process.uptime() });
});

// ---- API 라우트 ----
app.use('/api/prices', require('./routes/prices'));
app.use('/api/chat', require('./routes/chat'));
app.use('/api/orders', require('./routes/orders'));
app.use('/api/orders/load', require('./routes/load'));

// ---- 업로드 폴더 정적 서빙 ----
if (!fs.existsSync(config.uploadsDir)) {
  fs.mkdirSync(config.uploadsDir, { recursive: true });
}
app.use('/uploads', express.static(config.uploadsDir));

// ---- 프론트엔드 정적 호스팅 (옵션) ----
// 서버 상위 폴더의 woodcut_vm.html 를 서빙하면 VM 하나로 통합 서비스 가능.
const publicDir = path.join(__dirname, '..');
app.use(express.static(publicDir));
app.get('/', (req, res) => {
  const vm = path.join(publicDir, '..', 'woodcut_vm.html');
  if (fs.existsSync(vm)) return res.sendFile(vm);
  res.sendFile(path.join(publicDir, '..', 'woodcut.html'));
});

// ---- 404 (API) ----
app.use('/api', (req, res) => {
  res.status(404).json({ result: 'error', message: '알 수 없는 API 경로: ' + req.path });
});

// ---- 전역 에러 핸들러 ----
app.use((err, req, res, next) => {
  console.error('[server] 처리되지 않은 에러:', err);
  if (res.headersSent) return next(err);
  res.status(500).json({ result: 'error', message: '서버 내부 오류' });
});

app.listen(config.port, () => {
  console.log(`[server] woodcut-server 가 포트 ${config.port} 에서 실행 중입니다.`);
  console.log(`[server]   프론트: http://localhost:${config.port}/`);
  console.log(`[server]   헬스:   http://localhost:${config.port}/healthz`);
});
