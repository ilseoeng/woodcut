// ============================================================
// 단가표 import: 엑셀(목재.xlsx) → MS-SQL ai_quote_price_table
// ------------------------------------------------------------
// 사용:
//   node src/scripts/import-prices.js [엑셀경로]
//   기본 경로: C:\Users\ilseoeng\Downloads\목재.xlsx
//
// 엑셀 형식:
//   R1: area | misong18T | pyeonbaek18T | pvc18T | mulbau18T | mdf18T
//   R2~ : 면적(mm²) | 각 목재별 단가
//
// 기존 데이터가 있으면 모두 지우고 다시 넣는다 (idempotent).
// ============================================================
const fs = require('fs');
const path = require('path');
const { getPool, sql } = require('../db');

// 엑셀을 읽기 위해 임시로 파이썬을 사용할 수도 있지만,
// 여기서는 CSV로 변환된 단가표를 직접 사용할 수 있도록
// 두 가지 소스를 지원한다:
//   1) 엑셀 파일 (python openpyxl 로 읽어 JSON 생성)
//   2) JSON 파일 (price_table.json)

const FALLBACK_TABLE = [
  { area: 20000, mdf18T: 1000, pvc18T: 3000, misong18T: 1000, pyeonbaek18T: 2000, mulbau18T: 2000 },
  { area: 50000, mdf18T: 1000, pvc18T: 5000, misong18T: 2000, pyeonbaek18T: 4000, mulbau18T: 3000 },
  { area: 70000, mdf18T: 2000, pvc18T: 8000, misong18T: 3000, pyeonbaek18T: 6000, mulbau18T: 5000 },
  { area: 100000, mdf18T: 2000, pvc18T: 10000, misong18T: 4000, pyeonbaek18T: 8000, mulbau18T: 6000 },
  { area: 120000, mdf18T: 3000, pvc18T: 13000, misong18T: 5000, pyeonbaek18T: 10000, mulbau18T: 8000 },
  { area: 150000, mdf18T: 3000, pvc18T: 15000, misong18T: 6000, pyeonbaek18T: 12000, mulbau18T: 9000 },
  { area: 170000, mdf18T: 4000, pvc18T: 18000, misong18T: 7000, pyeonbaek18T: 14000, mulbau18T: 11000 },
  { area: 200000, mdf18T: 4000, pvc18T: 20000, misong18T: 8000, pyeonbaek18T: 16000, mulbau18T: 12000 },
  { area: 220000, mdf18T: 5000, pvc18T: 23000, misong18T: 9000, pyeonbaek18T: 18000, mulbau18T: 14000 },
  { area: 250000, mdf18T: 5000, pvc18T: 25000, misong18T: 10000, pyeonbaek18T: 20000, mulbau18T: 15000 },
  { area: 270000, mdf18T: 6000, pvc18T: 28000, misong18T: 11000, pyeonbaek18T: 22000, mulbau18T: 17000 },
  { area: 300000, mdf18T: 6000, pvc18T: 30000, misong18T: 12000, pyeonbaek18T: 24000, mulbau18T: 18000 },
  { area: 320000, mdf18T: 7000, pvc18T: 33000, misong18T: 13000, pyeonbaek18T: 26000, mulbau18T: 20000 },
  { area: 350000, mdf18T: 7000, pvc18T: 35000, misong18T: 14000, pyeonbaek18T: 28000, mulbau18T: 21000 },
  { area: 370000, mdf18T: 8000, pvc18T: 38000, misong18T: 15000, pyeonbaek18T: 30000, mulbau18T: 23000 },
  { area: 400000, mdf18T: 8000, pvc18T: 40000, misong18T: 16000, pyeonbaek18T: 32000, mulbau18T: 24000 },
  { area: 420000, mdf18T: 9000, pvc18T: 43000, misong18T: 17000, pyeonbaek18T: 34000, mulbau18T: 26000 },
  { area: 450000, mdf18T: 9000, pvc18T: 45000, misong18T: 18000, pyeonbaek18T: 36000, mulbau18T: 27000 },
  { area: 470000, mdf18T: 10000, pvc18T: 48000, misong18T: 19000, pyeonbaek18T: 38000, mulbau18T: 29000 },
  { area: 500000, mdf18T: 10000, pvc18T: 50000, misong18T: 20000, pyeonbaek18T: 40000, mulbau18T: 30000 },
  { area: 520000, mdf18T: 11000, pvc18T: 53000, misong18T: 21000, pyeonbaek18T: 42000, mulbau18T: 32000 },
  { area: 550000, mdf18T: 11000, pvc18T: 55000, misong18T: 22000, pyeonbaek18T: 44000, mulbau18T: 33000 },
  { area: 570000, mdf18T: 12000, pvc18T: 58000, misong18T: 23000, pyeonbaek18T: 46000, mulbau18T: 35000 },
  { area: 600000, mdf18T: 12000, pvc18T: 60000, misong18T: 24000, pyeonbaek18T: 48000, mulbau18T: 36000 },
  { area: 620000, mdf18T: 13000, pvc18T: 63000, misong18T: 25000, pyeonbaek18T: 50000, mulbau18T: 38000 },
  { area: 650000, mdf18T: 13000, pvc18T: 65000, misong18T: 26000, pyeonbaek18T: 52000, mulbau18T: 39000 },
  { area: 670000, mdf18T: 14000, pvc18T: 68000, misong18T: 27000, pyeonbaek18T: 54000, mulbau18T: 41000 },
  { area: 700000, mdf18T: 14000, pvc18T: 70000, misong18T: 28000, pyeonbaek18T: 56000, mulbau18T: 42000 },
  { area: 720000, mdf18T: 14500, pvc18T: 73000, misong18T: 29000, pyeonbaek18T: 58000, mulbau18T: 44000 },
  { area: 750000, mdf18T: 15000, pvc18T: 75000, misong18T: 30000, pyeonbaek18T: 60000, mulbau18T: 45000 },
  { area: 770000, mdf18T: 15500, pvc18T: 77500, misong18T: 31000, pyeonbaek18T: 62000, mulbau18T: 47000 },
  { area: 800000, mdf18T: 16000, pvc18T: 80000, misong18T: 32000, pyeonbaek18T: 64000, mulbau18T: 48000 },
  { area: 820000, mdf18T: 16500, pvc18T: 83000, misong18T: 33000, pyeonbaek18T: 66000, mulbau18T: 50000 },
  { area: 850000, mdf18T: 17000, pvc18T: 85000, misong18T: 34000, pyeonbaek18T: 68000, mulbau18T: 51000 },
  { area: 870000, mdf18T: 17500, pvc18T: 87500, misong18T: 35000, pyeonbaek18T: 70000, mulbau18T: 53000 },
  { area: 900000, mdf18T: 18000, pvc18T: 90000, misong18T: 36000, pyeonbaek18T: 72000, mulbau18T: 54000 },
  { area: 920000, mdf18T: 18500, pvc18T: 93000, misong18T: 37000, pyeonbaek18T: 74000, mulbau18T: 56000 },
  { area: 950000, mdf18T: 19000, pvc18T: 95000, misong18T: 38000, pyeonbaek18T: 76000, mulbau18T: 57000 },
  { area: 970000, mdf18T: 19500, pvc18T: 98000, misong18T: 39000, pyeonbaek18T: 78000, mulbau18T: 59000 },
  { area: 1000000, mdf18T: 20000, pvc18T: 100000, misong18T: 40000, pyeonbaek18T: 80000, mulbau18T: 60000 },
  { area: 1020000, mdf18T: 20500, pvc18T: 103000, misong18T: 41000, pyeonbaek18T: 82000, mulbau18T: 62000 },
  { area: 1050000, mdf18T: 21000, pvc18T: 105000, misong18T: 42000, pyeonbaek18T: 84000, mulbau18T: 63000 },
  { area: 1070000, mdf18T: 21500, pvc18T: 108000, misong18T: 43000, pyeonbaek18T: 86000, mulbau18T: 65000 },
  { area: 1100000, mdf18T: 22000, pvc18T: 110000, misong18T: 44000, pyeonbaek18T: 88000, mulbau18T: 66000 },
];

/**
 * 엑셀에서 단가표 읽기 (python openpyxl 사용, 별도 .py 파일 경유)
 * @param {string} xlsxPath
 * @returns {Promise<Array>} 단가표 배열
 */
async function readFromExcel(xlsxPath) {
  const { execSync } = require('child_process');
  const os = require('os');
  // 인라인 파이썬은 개행/따옴표 이스케이프가 깨지므로 임시 파일로 분리.
  const pyFile = path.join(os.tmpdir(), 'woodcut_read_prices.py');
  const py = `# -*- coding: utf-8 -*-
import json, sys, openpyxl
wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
ws = wb[wb.sheetnames[0]]
out = []
headers = None
for r in range(1, ws.max_row + 1):
    row = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
    if r == 1:
        headers = row
        continue
    area = row[0]
    if area is None:
        break
    rec = {"area": int(area)}
    for ci in range(1, len(headers)):
        h = headers[ci]
        if h and row[ci] is not None:
            rec[h] = int(row[ci])
    out.append(rec)
print(json.dumps(out, ensure_ascii=False))
`;
  fs.writeFileSync(pyFile, py, 'utf8');
  const result = execSync(`python "${pyFile}" "${xlsxPath}"`, {
    encoding: 'utf8',
    maxBuffer: 1024 * 1024 * 10,
  });
  fs.unlinkSync(pyFile);
  return JSON.parse(result.trim());
}

async function main() {
  const xlsxArg = process.argv[2];
  let table;

  if (xlsxArg && fs.existsSync(xlsxArg)) {
    console.log(`[import-prices] 엑셀에서 읽는 중: ${xlsxArg}`);
    table = await readFromExcel(xlsxArg);
    console.log(`[import-prices] ${table.length}행 읽음`);
  } else if (fs.existsSync(path.join(__dirname, '..', '..', 'price_table.json'))) {
    console.log('[import-prices] price_table.json 에서 읽는 중');
    table = JSON.parse(
      fs.readFileSync(path.join(__dirname, '..', '..', 'price_table.json'), 'utf8')
    );
  } else {
    console.log('[import-prices] 엑셀/JSON 없음 → 내장 기본 단가표 사용');
    table = FALLBACK_TABLE;
  }

  const pool = await getPool();

  // 기존 데이터 삭제 (idempotent)
  await pool.request().query('DELETE FROM ai_quote_price_table');

  // prepared insert (bulk는 IDENTITY 컬럼 충돌로 불안정)
  const ps = new sql.PreparedStatement(pool);
  ps.input('areaMax', sql.Int);
  ps.input('misong18T', sql.Int);
  ps.input('pyeonbaek18T', sql.Int);
  ps.input('pvc18T', sql.Int);
  ps.input('mulbau18T', sql.Int);
  ps.input('mdf18T', sql.Int);
  await ps.prepare(
    'INSERT INTO ai_quote_price_table (areaMax, misong18T, pyeonbaek18T, pvc18T, mulbau18T, mdf18T) ' +
      'VALUES (@areaMax, @misong18T, @pyeonbaek18T, @pvc18T, @mulbau18T, @mdf18T)'
  );

  let count = 0;
  for (const row of table) {
    await ps.execute({
      areaMax: Number(row.area),
      misong18T: Number(row.misong18T),
      pyeonbaek18T: Number(row.pyeonbaek18T),
      pvc18T: Number(row.pvc18T),
      mulbau18T: Number(row.mulbau18T),
      mdf18T: Number(row.mdf18T),
    });
    count++;
  }
  await ps.unprepare();

  const r = await pool
    .request()
    .query('SELECT COUNT(*) AS cnt, MIN(areaMax) AS minA, MAX(areaMax) AS maxA FROM ai_quote_price_table');
  const row = r.recordset[0];
  console.log(`[import-prices] ${count}행 import 완료.`);
  console.log(`[import-prices] 확인: ${row.cnt}행, 면적 ${row.minA}~${row.maxA} mm²`);
  process.exit(0);
}

main().catch((err) => {
  console.error('[import-prices] 실패:', err);
  process.exit(1);
});
