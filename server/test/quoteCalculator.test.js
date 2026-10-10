const test = require('node:test');
const assert = require('node:assert/strict');
const { calculateQuote, calculateProcessingPrice, normalizeProcessing } = require('../src/quoteCalculator');

const woodKey = 'misong18T';
const basePrice = 10000;
function quote({ quantity = 1, length = 600, selectedProcs = [], procDetails = {}, wood = woodKey } = {}) {
  return calculateQuote({ woodKey: wood, length, quantity, basePrice, selectedProcs, procDetails });
}

test('목재 단가만 있을 때 기본단가 × 수량', () => {
  assert.equal(quote({ quantity: 3 }).total, 30000);
});
test('무료 샌딩 이벤트가 적용된다', () => {
  assert.equal(quote({ selectedProcs: ['sanding'] }).total, 10000);
});
test('길이별 단면 및 양면 코팅 가공비가 다르다', () => {
  assert.equal(quote({ selectedProcs: ['coating'], procDetails: { coating: { type: '단면' } } }).total, 15000);
  assert.equal(quote({ length: 1200, quantity: 2, selectedProcs: ['coating'], procDetails: { coating: { type: '양면' } } }).total, 60000);
  assert.equal(quote({ length: 1500, selectedProcs: ['coating'] }).total, 25000);
});
test('사선 2면과 라운드 4개는 첫 1개/면이 무료', () => {
  const q = quote({ selectedProcs: ['angled', 'round'], procDetails: {
    angled: { sides: ['A', 'B'] }, round: { corners: ['1', '2', '3', '4'] },
  } });
  assert.equal(q.additionalCost, 6000);
  assert.equal(q.eventDiscount, 2000);
  assert.equal(q.total, 14000);
});
test('원형·사각 타공과 경첩 무료 이벤트가 적용된다', () => {
  const q = quote({ selectedProcs: ['circlecut', 'squarecut', 'hinge'], procDetails: {
    circlecut: { count: 3 }, squarecut: { count: 2 }, hinge: { count: 3 },
  } });
  assert.equal(q.additionalCost, 20000);
  assert.equal(q.eventDiscount, 9000);
  assert.equal(q.total, 21000);
});
test('수량이 증가하면 가공비와 할인 모두 수량에 따라 적용된다', () => {
  assert.equal(quote({ selectedProcs: ['circlecut'], procDetails: { circlecut: { count: 2 } }, quantity: 3 }).total, 39000);
});
test('MDF/PVC에 샌딩·코팅은 불가, 샌딩과 코팅 동시 선택도 불가', () => {
  assert.throws(() => quote({ wood: 'mdf18T', selectedProcs: ['sanding'] }), /원목/);
  assert.throws(() => quote({ wood: 'pvc18T', selectedProcs: ['coating'] }), /원목/);
  assert.throws(() => quote({ selectedProcs: ['sanding', 'coating'] }), /동시에/);
});
test('AI processing 값을 프론트와 동일한 구조로 정규화한다', () => {
  const x = normalizeProcessing([{ key: 'angled', count: 2 }, { key: 'coating', type: '양면' }]);
  assert.deepEqual(x.selectedProcs, ['angled', 'coating']);
  assert.deepEqual(x.procDetails, { angled: { count: 2 }, coating: { type: '양면' } });
  assert.equal(calculateProcessingPrice({ woodKey, length: 600, ...x }).additionalCost, 12000);
  assert.deepEqual(normalizeProcessing(undefined, { selectedProcs: ['round'], procDetails: { round: { corners: ['1'] } } }).selectedProcs, ['round']);
  assert.deepEqual(normalizeProcessing([], { selectedProcs: ['round'] }).selectedProcs, []);
});
