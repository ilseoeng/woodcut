// ============================================================
// 도구 메타데이터 (단일 진실 공급원)
// ------------------------------------------------------------
// 도구의 '실행'은 프론트엔드(TOOL_REGISTRY)가 담당한다.
// 이 파일은 도구의 '설명/매개변수 스키마'만 정의하며,
// schema.js 가 이것을 Gemini function declarations 로 변환한다.
//
// 중요: 프론트엔드 TOOL_REGISTRY 의 키/매개변수와 반드시 일치해야 한다.
//       (woodcut_vm.html 의 TOOL_REGISTRY 참조)
// ============================================================

// 가공 키 (프론트 PROC_OPTIONS 와 동일)
const PROC_KEYS = [
  'none', 'sanding', 'coating', 'angled', 'round', 'circlecut', 'squarecut', 'hinge',
];

const TOOL_REGISTRY = {
  // ---- 목재 선택/정보 ----
  select_wood: {
    description: '목재를 선택한다. 목재 키로 mdf18T, pvc18T, misong18T, pyeonbaek18T, mulbau18T 중 하나.',
    parameters: {
      type: 'object',
      properties: {
        woodKey: { type: 'string', enum: ['mdf18T', 'pvc18T', 'misong18T', 'pyeonbaek18T', 'mulbau18T'], description: '목재 키' },
      },
      required: ['woodKey'],
    },
  },
  show_woods: {
    description: '목재 선택 카드 목록을 화면에 표시한다. 목재 종류를 다시 보여줘야 할 때 호출.',
    parameters: { type: 'object', properties: {} },
  },
  show_wood_detail: {
    description: '특정 목재의 상세 설명(특징, 샌딩/코팅 가능 여부, 영상)을 표시한다.',
    parameters: {
      type: 'object',
      properties: {
        woodKey: { type: 'string', enum: ['mdf18T', 'pvc18T', 'misong18T', 'pyeonbaek18T', 'mulbau18T'] },
      },
      required: ['woodKey'],
    },
  },
  compare_wood_prices: {
    description: '특정 사이즈의 목재 가격을 비교해서 저렴한 순으로 보여준다. 가격 비교/추천 질문일 때 호출.',
    parameters: {
      type: 'object',
      properties: {
        width: { type: 'number', description: '폭(mm)' },
        length: { type: 'number', description: '길이(mm)' },
        quantity: { type: 'number', description: '수량(선택)' },
      },
      required: ['width', 'length'],
    },
  },

  // ---- 사이즈/수량 ----
  set_size: {
    description: '목재 사이즈(폭/길이 mm)와 수량을 설정한다. 폭은 최대 600, 길이는 최대 1800, 최소 10.',
    parameters: {
      type: 'object',
      properties: {
        width: { type: 'number', description: '폭(mm, 10~600)' },
        length: { type: 'number', description: '길이(mm, 10~1800)' },
        quantity: { type: 'number', description: '수량(선택). 없으면 수량 입력 요청.' },
      },
      required: ['width', 'length'],
    },
  },
  set_quantity: {
    description: '수량만 설정한다. 사이즈는 이미 정해진 상태여야 한다.',
    parameters: {
      type: 'object',
      properties: { quantity: { type: 'number', description: '수량(1 이상)' } },
      required: ['quantity'],
    },
  },

  // ---- 가공 옵션 ----
  show_processing_options: {
    description: '가공 옵션 선택 그리드를 표시한다.',
    parameters: { type: 'object', properties: {} },
  },
  toggle_processing: {
    description: '가공 옵션을 켜거나 끈다. 이미 켜져 있으면 끄고, 꺼져 있으면 켠다.',
    parameters: {
      type: 'object',
      properties: { key: { type: 'string', enum: PROC_KEYS, description: '가공 키' } },
      required: ['key'],
    },
  },
  set_processing_detail: {
    description: '활성화된 가공 옵션의 상세 설정을 지정한다. 해당 가공이 먼저 toggle_processing 으로 켜져 있어야 한다.',
    parameters: {
      type: 'object',
      properties: {
        key: { type: 'string', enum: ['angled', 'round', 'coating', 'hinge', 'circlecut', 'squarecut'], description: '가공 키' },
        detail: {
          type: 'object',
          description: '가공별 상세. angled:{sides:[A,B,C,D]} round:{corners:[1,2,3,4]} coating:{type:단면|양면} hinge:{count,positions,mode} circlecut:{count,holes:[{x,y,r}]} squarecut:{count,holes:[{x,y,w,h}]}',
          properties: {
            sides: { type: 'array', items: { type: 'string', enum: ['A', 'B', 'C', 'D'] } },
            corners: { type: 'array', items: { type: 'string', enum: ['1', '2', '3', '4'] } },
            type: { type: 'string', enum: ['단면', '양면'] },
            count: { type: 'number' },
            holes: { type: 'array', items: { type: 'object' } },
            positions: { type: 'array', items: { type: 'object' } },
            mode: { type: 'string', enum: ['Y_INPUT', 'X_INPUT'] },
          },
        },
      },
      required: ['key', 'detail'],
    },
  },

  // ---- 견적/주문 ----
  show_estimate: {
    description: '현재 선택된 목재/사이즈/가공으로 견적서를 표시한다.',
    parameters: { type: 'object', properties: {} },
  },
  add_to_order: {
    description: '현재 견적을 주문목록에 추가한다.',
    parameters: { type: 'object', properties: {} },
  },
  add_accessory: {
    description: '부자재(피스/경첩/사포/엣지테이프)를 주문목록에 추가한다.',
    parameters: {
      type: 'object',
      properties: {
        categoryKey: { type: 'string', enum: ['screws', 'hinges', 'sandpaper', 'tape'] },
        itemIndex: { type: 'number', description: '카테고리 내 항목 인덱스(0부터)' },
        quantity: { type: 'number', description: '수량(기본 1)' },
      },
      required: ['categoryKey', 'itemIndex'],
    },
  },
  show_accessories: {
    description: '부자재 카테고리/항목을 표시한다. categoryKey 생략 시 전체 카테고리 표시.',
    parameters: {
      type: 'object',
      properties: {
        categoryKey: { type: 'string', enum: ['screws', 'hinges', 'sandpaper', 'tape'] },
      },
    },
  },
  initiate_order: {
    description: '주문 진행 패널을 띄운다. 주문하기/결제 진행 의도일 때 호출.',
    parameters: { type: 'object', properties: {} },
  },

  // ---- 초기화/조회 ----
  reset_all: {
    description: '모든 선택과 주문목록을 초기화하고 첫 화면으로 돌아간다.',
    parameters: { type: 'object', properties: {} },
  },
  get_state: {
    description: '현재 선택된 목재/사이즈/수량/가공/주문목록 상태를 반환한다. 상황 파악이 필요할 때 호출. (부작용 없음)',
    parameters: { type: 'object', properties: {} },
  },
};

module.exports = { TOOL_REGISTRY, PROC_KEYS };
