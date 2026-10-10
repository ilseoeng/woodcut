// 목재 추가가공 견적 규칙. woodcut_vm.html의 calculatePrice()와 동기화한다.
// 기본 목재단가는 항상 MS-SQL에서 조회하고, 아래에는 가공비 및 이벤트 할인만 정의한다.
const KNOWN_PROCESSES = new Set([
  'none', 'sanding', 'coating', 'angled', 'round', 'circlecut', 'squarecut', 'hinge',
]);
const NATURAL_WOODS = new Set(['misong18T', 'pyeonbaek18T', 'mulbau18T']);

function toValidCount(value, fallback) {
  const n = Number(value);
  return Number.isFinite(n) && n > 0 ? Math.floor(n) : fallback;
}

function calculateProcessingPrice({ woodKey, length, selectedProcs = [], procDetails = {} }) {
  const keys = new Set(selectedProcs || []);
  keys.delete('none');
  for (const key of keys) {
    if (!KNOWN_PROCESSES.has(key)) throw new Error('지원하지 않는 가공 옵션: ' + key);
  }
  if ((keys.has('coating') || keys.has('sanding')) && !NATURAL_WOODS.has(woodKey)) {
    throw new Error('샌딩과 투명코팅은 미송·편백·멀바우 원목에만 가능합니다.');
  }
  if (keys.has('sanding') && keys.has('coating')) {
    throw new Error('샌딩과 투명코팅은 동시에 선택할 수 없습니다.');
  }

  let additionalCost = 0;
  let eventDiscount = 0;
  const discountDetails = [];

  for (const key of keys) {
    const d = procDetails[key] || {};
    switch (key) {
      case 'sanding':
        additionalCost += 5000;
        eventDiscount += 5000;
        discountDetails.push('샌딩 무료');
        break;
      case 'coating': {
        const l = Number(length);
        const single = l <= 600 ? 5000 : l <= 1200 ? 10000 : 15000;
        additionalCost += d.type === '양면' ? single * 2 : single;
        break;
      }
      case 'angled': {
        const n = Array.isArray(d.sides) ? d.sides.length : toValidCount(d.count, 0);
        additionalCost += n * 1000;
        if (n > 0) {
          eventDiscount += 1000;
          discountDetails.push('사선 1면 무료');
        }
        break;
      }
      case 'round': {
        const n = Array.isArray(d.corners) ? d.corners.length : toValidCount(d.count, 0);
        additionalCost += n * 1000;
        if (n > 0) {
          eventDiscount += 1000;
          discountDetails.push('라운드 1개 무료');
        }
        break;
      }
      case 'circlecut': {
        const n = toValidCount(d.count, 1);
        additionalCost += n * 3000;
        eventDiscount += 3000;
        discountDetails.push('원형타공 1개 무료');
        break;
      }
      case 'squarecut': {
        const n = toValidCount(d.count, 1);
        additionalCost += n * 4000;
        eventDiscount += 4000;
        discountDetails.push('사각타공 1개 무료');
        break;
      }
      case 'hinge': {
        const n = toValidCount(d.count, 2);
        additionalCost += n * 1000;
        const free = Math.min(n, 2);
        eventDiscount += free * 1000;
        discountDetails.push(`경첩 ${free}개 무료`);
        break;
      }
    }
  }
  return { additionalCost, eventDiscount, discountDetails };
}

function calculateQuote({ woodKey, length, quantity = 1, basePrice, selectedProcs = [], procDetails = {} }) {
  const qty = Number(quantity);
  const base = Number(basePrice);
  if (!Number.isSafeInteger(qty) || qty < 1) throw new Error('수량은 1 이상의 정수여야 합니다.');
  if (!Number.isFinite(base) || base < 0) throw new Error('목재 기본단가가 올바르지 않습니다.');

  const charges = calculateProcessingPrice({ woodKey, length, selectedProcs, procDetails });
  const finalUnitPrice = base + charges.additionalCost - charges.eventDiscount;
  return {
    basePrice: base,
    additionalCost: charges.additionalCost,
    eventDiscount: charges.eventDiscount,
    discountDetails: charges.discountDetails,
    finalUnitPrice,
    quantity: qty,
    total: finalUnitPrice * qty, // 배송비는 주문목록 묶음배송 계산에서 별도
  };
}

// Gemini function-call의 processing 배열을 기존 프론트의 selectedProcs/procDetails로 바꾼다.
function normalizeProcessing(processing, state = {}) {
  if (!Array.isArray(processing)) {
    return {
      selectedProcs: Array.isArray(state.selectedProcs) ? state.selectedProcs : [],
      procDetails: state.procDetails && typeof state.procDetails === 'object' ? state.procDetails : {},
    };
  }
  const selectedProcs = [];
  const procDetails = {};
  for (const item of processing) {
    if (!item || !item.key || selectedProcs.includes(item.key)) continue;
    selectedProcs.push(item.key);
    procDetails[item.key] = {};
    for (const prop of ['count', 'sides', 'corners', 'type']) {
      if (item[prop] !== undefined) procDetails[item.key][prop] = item[prop];
    }
  }
  return { selectedProcs, procDetails };
}

module.exports = { calculateProcessingPrice, calculateQuote, normalizeProcessing };
