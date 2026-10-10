// ============================================================
// 백엔드 실행 도구 (DB 조회 필요)
// ------------------------------------------------------------
// 프론트 실행 도구(TOOL_REGISTRY)는 UI를 변경하고,
// 백엔드 실행 도구(이 파일)는 DB를 조회해 결과를 AI에게 돌려준다.
//
// AI가 functionCall로 이 도구들을 부르면:
//   1) gemini.js 의 Agentic 루프가 이 파일의 executor 를 실행
//   2) 결과를 functionResponse 로 Gemini에게 돌려보냄
//   3) Gemini가 결과를 보고 다시 응답(추천/설명/추가 도구 호출)
// ============================================================
const pricing = require('../pricing');
const quoteCalculator = require('../quoteCalculator');

// 목재 메타데이터 (프론트 WOOD_DATA 와 동기화 필요)
// 각 항목의 pros/cons/cautions 는 get_wood_info 도구가 반환하는 풍부한 설명 자료.
const WOOD_INFO = {
  mdf18T: {
    name: 'MDF 18T',
    desc: '가성비 좋은 기본 합판',
    features: '목재 섬유를 압축해 표면이 매끄럽고 페인트 칠에 좋음. 선반이나 수납장에 가장 많이 쓰임.',
    pros: ['가격이 가장 저렴', '표면이 고르고 매끄러워 마감재·페인트 도장에 유리', '가공이 쉬워 초보자도 다루기 편함', '수납장·선반 등 범용 가구에 적합'],
    cons: ['방수가 안 됨 (물에 약함)', '나무 결(원목 느낌)이 없음', '샌딩·코팅 불가', '수분 먹으면 부풀어 오를 수 있음'],
    cautions: ['물이 닿는 곳(욕실/주방/세탁실)에는 부적합', '원목 질감을 원하면 다른 목재 선택', '코팅 가공 불가하니 마감재로 마무리'],
    canSand: false,
    canCoat: false,
    waterproof: false,
    priceTier: '최저가',
    useCases: ['선반', '수납장', '가성비 가구', '페인트 도장 가구'],
  },
  pvc18T: {
    name: 'PVC보드 18T + PET',
    desc: '완전방수 · 욕실/주방 추천',
    features: '완전 방수. 물에 닿아도 썩지 않아 욕실·주방·세탁실·발코니에 적합. PET 코팅으로 깔끔함.',
    pros: ['완전 방수 (물에 완전히 강함)', '습기·곰팡이에 강해 위생적', '청소가 쉬움', '욕실·주방·세탁실·발코니 등 습한 곳에 최적'],
    cons: ['원목 느낌이 아님 (플라스틱 질감)', '가격이 MDF보다 비쌈', '샌딩·코팅 불가', '친환경 이미지는 약함'],
    cautions: ['원목 질감을 원하면 부적합', '페인트 도장이 어려움 (PET 코팅면)', '투명코팅 등 추가 가공 불가'],
    canSand: false,
    canCoat: false,
    waterproof: true,
    priceTier: '중간',
    useCases: ['욕실 선반', '주방 싱크대', '세탁실', '발코니', '물이 닿는 곳'],
  },
  misong18T: {
    name: '미송집성목 18T',
    desc: '부드러운 원목 느낌',
    features: '나무 본연의 따뜻한 느낌. 결이 부드러워 DIY 선호. 샌딩+코팅 시 고급스러움.',
    pros: ['따뜻한 원무늬 결이 아늑함', '결이 부드러워 가공·DIY 적합', '샌딩·코팅 가능 (고급 마감)', '거실·침실 가구에 인기'],
    cons: ['방수 아님 (물에 약함)', 'MDF보다 비쌈', '무거운 하중에는 멀바우가 더 적합', '가공 안 하면 거칠 수 있어 샌딩 권장'],
    cautions: ['물이 닿는 곳 부적합', '폭>길이 방향은 파손 위험 (집성목 공통)', '코팅과 샌딩 동시 선택 불가'],
    canSand: true,
    canCoat: true,
    waterproof: false,
    priceTier: '중간',
    useCases: ['원무늬 선반', 'DIY 가구', '거실/침실 가구', '코팅 가공 가구'],
  },
  pyeonbaek18T: {
    name: '편백집성목 18T',
    desc: '피톤치드 · 프리미엄 원목',
    features: '피톤치드 방출로 항균 효과. 아이 방/침실 가구에 추천. 향기와 건강까지 고려한 프리미엄.',
    pros: ['피톤치드 방출 → 항균·탈취 효과', '아이·호흡기 약한 분에게 추천', '은은한 나무 향', '침실·아이 방에 건강 이점'],
    cons: ['가격이 비교적 높은 편', '방수 아님', '무거운 하중에는 부적합', '향에 민감하면 확인 필요'],
    cautions: ['물이 닿는 곳 부적합', '폭>길이 방향은 파손 위험', '코팅 시 향이 다소 가려질 수 있음'],
    canSand: true,
    canCoat: true,
    waterproof: false,
    priceTier: '고급',
    useCases: ['아이 방 가구', '침실 가구', '항균이 필요한 곳', '프리미엄 선반'],
  },
  mulbau18T: {
    name: '멀바우 집성목 18T',
    desc: '고급 하드우드 · 강한 내구성',
    features: '단단하고 묵직한 하드우드. 짙은 색감으로 고급스러움. 무거운 물건을 올리는 테이블에 적합.',
    pros: ['하드우드라 매우 단단하고 내구성 강함', '무거운 물건(책·기기) 하중에 적합', '짙은 색감으로 고급스러운 분위기', '테이블 상판·작업대로 최적'],
    cons: ['가장 비쌈', '무거워 다루기 힘들 수 있음', '방수 아님', '결이 거칠어 샌딩 필요할 수 있음'],
    cautions: ['물이 닿는 곳 부적합', '폭>길이 방향은 파손 위험', '투명코팅 권장 (거친 결 보호)'],
    canSand: true,
    canCoat: true,
    waterproof: false,
    priceTier: '최고가',
    useCases: ['테이블 상판', '무거운 물건 선반', '고품질 인테리어', '내구성 필수 가구'],
  },
};

const LAMINATED_WOODS = ['misong18T', 'pyeonbaek18T', 'mulbau18T'];

/**
 * 백엔드 도구 스키마 (Gemini functionDeclarations 에 추가될 항목)
 */
const BACKEND_TOOL_SCHEMAS = {
  get_quote: {
    description:
      '목재 종류·사이즈·수량·추가 가공을 반영한 실제 견적을 조회한다. 목재 기본단가는 MS-SQL에서 조회하고, 가공비와 이벤트 할인을 적용한다. 견적·가격·단가 질문에서는 추측하지 말고 반드시 호출한다. 고객이 이번 메시지에서 가공을 지정하면 processing에 포함한다.',
    parameters: {
      type: 'object',
      properties: {
        woodKey: {
          type: 'string',
          enum: ['mdf18T', 'pvc18T', 'misong18T', 'pyeonbaek18T', 'mulbau18T'],
          description: '목재 키. 여러 개 비교 시 compare=true',
        },
        width: { type: 'number', description: '폭(mm)' },
        length: { type: 'number', description: '길이(mm)' },
        quantity: { type: 'number', description: '수량(기본 1)' },
        compare: {
          type: 'boolean',
          description: '모든 목재를 같은 규격·가공 기준으로 가격 비교',
        },
        processing: {
          type: 'array',
          description: '이번 메시지에서 고객이 말한 가공 옵션을 반드시 담는다. 생략하면 현재 UI 가공 상태를 쓴다. 가공 없음은 빈 배열을 보낸다. 사선/라운드는 sides/corners 또는 count, 코팅은 type, 타공/경첩은 count를 설정.',
          items: {
            type: 'object',
            properties: {
              key: {
                type: 'string',
                enum: ['sanding', 'coating', 'angled', 'round', 'circlecut', 'squarecut', 'hinge'],
              },
              type: { type: 'string', enum: ['단면', '양면'] },
              count: { type: 'number', description: '가공 개수/면 수. 기본 수량이 아닌 해당 가공 횟수' },
              sides: { type: 'array', items: { type: 'string', enum: ['A', 'B', 'C', 'D'] } },
              corners: { type: 'array', items: { type: 'string', enum: ['1', '2', '3', '4'] } },
            },
            required: ['key'],
          },
        },
      },
      required: ['width', 'length'],
    },
  },
  recommend_wood: {
    description:
      '고객의 용도/조건에 맞는 목재를 추천한다. "책장 만들 건데 뭘 써?", "방습 필요해", "아이 방인데" 같은 상담성 질문일 때 사용. 필요하면 사이즈도 받아 단가까지 함께 제시.',
    parameters: {
      type: 'object',
      properties: {
        useCase: {
          type: 'string',
          description: '용도/상황. 예: "책장", "욕실 선반", "아이 방 가구", "테이블 상판", "방습"',
        },
        requirements: {
          type: 'array',
          items: { type: 'string' },
          description: '추가 요구조건. 예: ["방수", "가성비", "원목 느낌", "내구성"]',
        },
        width: { type: 'number', description: '폭(mm). 있으면 단가까지 제시' },
        length: { type: 'number', description: '길이(mm)' },
        budget: { type: 'string', description: '예산 힌트. 예: "저렴하게", "가성비", "프리미엄"' },
      },
      required: ['useCase'],
    },
  },

  compare_woods: {
    description:
      '특정 목재 2개 이상을 지정된 사이즈로 직접 비교한다. "멀바우랑 미송으로 100x1000 얼마야?", "미송이랑 편백이랑 뭐가 달라?", "MDF랑 PVC 비교해줘"처럼 특정 목재들을 비교하는 질문에 사용. 각 목재의 단가·특징·차이점을 함께 반환.',
    parameters: {
      type: 'object',
      properties: {
        woodKeys: {
          type: 'array',
          items: { type: 'string', enum: ['mdf18T', 'pvc18T', 'misong18T', 'pyeonbaek18T', 'mulbau18T'] },
          description: '비교할 목재 키 배열 (2개 이상). 예: ["mulbau18T","misong18T"]',
        },
        width: { type: 'number', description: '폭(mm). 있으면 각 목재별 단가까지 비교' },
        length: { type: 'number', description: '길이(mm)' },
        quantity: { type: 'number', description: '수량(기본 1)' },
      },
      required: ['woodKeys'],
    },
  },

  validate_processing: {
    description:
      '선택된 가공 옵션组合이 유효한지 검증한다. 샌딩/코팅 충돌, 원목 여부, 집성목 나뭇결 방향 등 비즈니스 규칙을 점검해 경고/제안을 반환.',
    parameters: {
      type: 'object',
      properties: {
        woodKey: { type: 'string', enum: ['mdf18T', 'pvc18T', 'misong18T', 'pyeonbaek18T', 'mulbau18T'] },
        width: { type: 'number', description: '폭(mm). 집성목 방향 경고용' },
        length: { type: 'number', description: '길이(mm)' },
        processing: {
          type: 'array',
          items: { type: 'string' },
          description: '가공 키 배열. 예: ["sanding","coating","angled"]',
        },
      },
      required: ['woodKey', 'processing'],
    },
  },

  get_wood_info: {
    description:
      '특정 목재의 상세 정보를 반환한다. 장점·단점·주의점·추천 용도·가격대·가공 가능 여부를 포함. "미송 장단점 알려줘", "편백이 뭐가 좋아?", "멀바우 특징?" 같은 목재 자체에 대한 질문에 사용. woodKey 생략 시 전체 목재 요약 반환.',
    parameters: {
      type: 'object',
      properties: {
        woodKey: {
          type: 'string',
          enum: ['mdf18T', 'pvc18T', 'misong18T', 'pyeonbaek18T', 'mulbau18T'],
          description: '상세 조회할 목재 키. 생략하면 5개 전체 요약 반환',
        },
      },
    },
  },
};

/**
 * 백엔드 도구 실행기. 도구 이름 → 실행 함수.
 * 각 함수는 args 를 받아 결과 객체를 반환 (AI에게 functionResponse로 전달).
 */
const BACKEND_TOOL_EXECUTORS = {
  async get_quote(args, context = {}) {
    const { woodKey, width, length, quantity, compare, processing } = args;
    const w = Number(width), l = Number(length), qty = quantity === undefined ? 1 : Number(quantity);
    if (!Number.isFinite(w) || !Number.isFinite(l) || w < 10 || l < 10) {
      return { error: '폭과 길이는 각각 10mm 이상이어야 합니다.' };
    }
    if (w > 600) return { error: `폭은 최대 600mm입니다 (입력: ${w}mm)` };
    if (l > 1800) return { error: `길이는 최대 1800mm입니다 (입력: ${l}mm)` };
    if (!Number.isSafeInteger(qty) || qty < 1) return { error: '수량은 1 이상의 정수여야 합니다.' };
    if (woodKey && !WOOD_INFO[woodKey]) return { error: '지원하지 않는 목재입니다: ' + woodKey };

    const area = w * l;
    const { selectedProcs, procDetails } = quoteCalculator.normalizeProcessing(processing, context.state);
    const quoteFor = (key, basePrice) => {
      const price = quoteCalculator.calculateQuote({
        woodKey: key,
        length: l,
        quantity: qty,
        basePrice,
        selectedProcs,
        procDetails,
      });
      return {
        woodKey: key, woodName: WOOD_INFO[key].name, unitPrice: price.finalUnitPrice,
        basePrice: price.basePrice, additionalCost: price.additionalCost,
        eventDiscount: price.eventDiscount, discountDetails: price.discountDetails,
        total: price.total,
      };
    };

    if (compare || !woodKey) {
      const all = await pricing.getAllUnitPrices(area);
      const comparison = all.map(({ woodKey: key, price }) => {
        try {
          return { ...quoteFor(key, price), desc: WOOD_INFO[key].desc };
        } catch (e) {
          return { woodKey: key, woodName: WOOD_INFO[key].name, error: e.message };
        }
      });
      comparison.sort((a, b) => (a.error ? 1 : 0) - (b.error ? 1 : 0) || ((a.unitPrice || 0) - (b.unitPrice || 0)));
      const cheapest = comparison.find((p) => !p.error && p.unitPrice > 0);
      return {
        width: w, length: l, quantity: qty, area,
        selectedProcs, procDetails,
        cheapest: cheapest ? cheapest.woodKey : null,
        comparison,
        note: '배송비는 주문목록의 묶음배송 기준으로 별도 계산합니다.',
      };
    }

    const unitPrice = await pricing.getUnitPrice(woodKey, area);
    if (!unitPrice) return { error: `${woodKey}의 ${area}mm² 구간 단가를 찾을 수 없습니다` };
    try {
      return {
        ...quoteFor(woodKey, unitPrice),
        width: w, length: l, quantity: qty, area,
        selectedProcs, procDetails,
        note: '단가 = 목재 기본단가 + 추가 가공비 - 이벤트 할인. 배송비는 별도입니다.',
      };
    } catch (e) {
      return { error: e.message };
    }
  },

  async recommend_wood(args) {
    const { useCase, requirements = [], width, length, budget } = args;
    const reqLower = (requirements || []).map((r) => String(r).toLowerCase());
    const useLower = String(useCase || '').toLowerCase();

    // 점수 기반 추천
    const scored = Object.entries(WOOD_INFO).map(([key, info]) => {
      let score = 0;
      const reasons = [];

      // 용도 매칭
      if (info.useCases.some((u) => useLower.includes(u))) {
        score += 3;
        reasons.push(`용도(${useCase})에 적합`);
      }

      // 요구조건 매칭
      if (reqLower.includes('방수') || reqLower.includes('방습') || useLower.includes('욕실') || useLower.includes('주방') || useLower.includes('세탁')) {
        if (info.waterproof) { score += 5; reasons.push('완전방수'); }
        else { score -= 2; reasons.push('방수 아님'); }
      }
      if (reqLower.includes('가성비') || budget === '저렴하게' || budget === '가성비') {
        if (key === 'mdf18T') { score += 3; reasons.push('가성비 최고'); }
      }
      if (reqLower.includes('원목') || useLower.includes('아이') || useLower.includes('침실')) {
        if (key === 'pyeonbaek18T') { score += 3; reasons.push('피톤치드·항균(아이/침실 추천)'); }
        else if (info.canCoat && LAMINATED_WOODS.includes(key)) { score += 1; }
      }
      if (reqLower.includes('내구성') || useLower.includes('테이블') || useLower.includes('무거운')) {
        if (key === 'mulbau18T') { score += 3; reasons.push('하드우드·강한 내구성'); }
      }

      return { key, name: info.name, desc: info.desc, score, reasons, features: info.features };
    });

    scored.sort((a, b) => b.score - a.score);
    const best = scored[0];

    // 사이즈 있으면 단가까지
    let pricingInfo = null;
    if (width && length) {
      const w = Number(width), l = Number(length);
      const area = w * l;
      const unitPrice = await pricing.getUnitPrice(best.key, area);
      pricingInfo = { area, unitPrice, total: unitPrice * 1 };
    }

    return {
      useCase,
      recommended: { woodKey: best.key, name: best.name, reasons: best.reasons, features: best.features },
      alternatives: scored.slice(1, 3).map((s) => ({ woodKey: s.key, name: s.name, reasons: s.reasons })),
      pricing: pricingInfo,
    };
  },

  async validate_processing(args) {
    const { woodKey, width, length, processing } = args;
    const info = WOOD_INFO[woodKey];
    if (!info) return { error: '알 수 없는 목재: ' + woodKey };

    const warnings = [];
    const errors = [];
    const procs = processing || [];

    // 1) 샌딩 가능 여부
    if (procs.includes('sanding') && !info.canSand) {
      errors.push(`${info.name}은(는) 샌딩이 불가능합니다 (원목만 가능).`);
    }
    // 2) 코팅 가능 여부
    if (procs.includes('coating') && !info.canCoat) {
      errors.push(`${info.name}은(는) 코팅이 불가능합니다 (원목만 가능).`);
    }
    // 3) 샌딩+코팅 충돌
    if (procs.includes('sanding') && procs.includes('coating')) {
      warnings.push('샌딩과 코팅은 동시 선택할 수 없습니다. 코팅에는 별도 샌딩이 포함되지 않습니다.');
    }
    // 4) 집성목 나뭇결 방향 (폭 > 길이)
    if (width && length && LAMINATED_WOODS.includes(woodKey) && Number(width) > Number(length)) {
      warnings.push(
        `집성목(${info.name})은 폭(${width}mm)이 길이(${length}mm)보다 넓으면 결대로 쪼개지거나 배송 중 파손 위험이 있습니다. ${length}x${width}mm 로 변경을 권장합니다.`
      );
    }

    return {
      valid: errors.length === 0,
      errors,
      warnings,
      woodKey,
      woodName: info.name,
      processing: procs,
    };
  },

  async compare_woods(args) {
    const { woodKeys, width, length, quantity } = args;
    const keys = Array.isArray(woodKeys) ? woodKeys.filter((k) => WOOD_INFO[k]) : [];
    const qty = Number(quantity) || 1;

    if (keys.length < 2) {
      return { error: '비교하려면 최소 2개 이상의 목재가 필요합니다.' };
    }

    // 사이즈가 있으면 각 목재의 단가까지 함께 비교
    let area = null;
    let priceInfo = null;
    if (width && length && Number(width) > 0 && Number(length) > 0) {
      const w = Number(width), l = Number(length);
      if (w > 600) return { error: `폭은 최대 600mm입니다 (입력: ${w}mm)` };
      if (l > 1800) return { error: `길이는 최대 1800mm입니다 (입력: ${l}mm)` };
      area = w * l;
      priceInfo = [];
      for (const key of keys) {
        const unitPrice = await pricing.getUnitPrice(key, area);
        priceInfo.push({
          woodKey: key,
          woodName: WOOD_INFO[key].name,
          unitPrice,
          total: unitPrice * qty,
        });
      }
      priceInfo.sort((a, b) => a.unitPrice - b.unitPrice);
    }

    // 특징 비교 표 (모든 목재 공통 항목)
    const features = keys.map((key) => {
      const info = WOOD_INFO[key];
      return {
        woodKey: key,
        woodName: info.name,
        desc: info.desc,
        features: info.features,
        useCases: info.useCases,
        canSand: info.canSand,
        canCoat: info.canCoat,
        waterproof: info.waterproof,
        isLaminated: LAMINATED_WOODS.includes(key),
      };
    });

    // 차이점 요약 (AI가 설명하기 쉽도록 핵심 차이 도출)
    const differences = [];
    const sandable = features.filter((f) => f.canSand).map((f) => f.woodName);
    const notSandable = features.filter((f) => !f.canSand).map((f) => f.woodName);
    if (sandable.length && notSandable.length) {
      differences.push(`샌딩/코팅: ${sandable.join(', ')} 가능 vs ${notSandable.join(', ')} 불가`);
    }
    const waterproof = features.filter((f) => f.waterproof).map((f) => f.woodName);
    if (waterproof.length && waterproof.length < features.length) {
      const nonWater = features.filter((f) => !f.waterproof).map((f) => f.woodName);
      differences.push(`방수: ${waterproof.join(', ')} 완전방수 vs ${nonWater.join(', ')} 방수 아님`);
    }
    if (priceInfo && priceInfo.length >= 2) {
      const cheapest = priceInfo[0];
      const expensive = priceInfo[priceInfo.length - 1];
      const diff = expensive.total - cheapest.total;
      differences.push(
        `가격: ${cheapest.woodName} ${cheapest.total.toLocaleString()}원(최저) vs ${expensive.woodName} ${expensive.total.toLocaleString()}원(최고) → ${diff.toLocaleString()}원 차이`
      );
    }

    return {
      woodKeys: keys,
      width: width ? Number(width) : null,
      length: length ? Number(length) : null,
      quantity: qty,
      area,
      features,
      pricing: priceInfo,
      cheapest: priceInfo && priceInfo.length ? priceInfo[0].woodKey : null,
      differences,
    };
  },

  async get_wood_info(args) {
    const { woodKey } = args;

    // 특정 목재 상세 조회
    if (woodKey) {
      const info = WOOD_INFO[woodKey];
      if (!info) return { error: '알 수 없는 목재: ' + woodKey };
      return {
        woodKey,
        woodName: info.name,
        desc: info.desc,
        features: info.features,
        pros: info.pros,
        cons: info.cons,
        cautions: info.cautions,
        useCases: info.useCases,
        canSand: info.canSand,
        canCoat: info.canCoat,
        waterproof: info.waterproof,
        isLaminated: LAMINATED_WOODS.includes(woodKey),
        priceTier: info.priceTier,
      };
    }

    // woodKey 생략 → 전체 요약
    const summary = Object.entries(WOOD_INFO).map(([key, info]) => ({
      woodKey: key,
      woodName: info.name,
      desc: info.desc,
      priceTier: info.priceTier,
      waterproof: info.waterproof,
      canSand: info.canSand,
      canCoat: info.canCoat,
      useCases: info.useCases,
    }));

    // 가격대 순 정렬 (직관적 비교용)
    const tierOrder = { 최저가: 0, 중간: 1, 고급: 2, '최고가': 3 };
    summary.sort((a, b) => (tierOrder[a.priceTier] ?? 9) - (tierOrder[b.priceTier] ?? 9));

    return { count: summary.length, woods: summary };
  },
};

module.exports = {
  WOOD_INFO,
  LAMINATED_WOODS,
  BACKEND_TOOL_SCHEMAS,
  BACKEND_TOOL_EXECUTORS,
};
