-- ============================================================
-- 아이린가구 AI 목재 견적 - 신규 테이블 (ai_quote_*)
-- ------------------------------------------------------------
-- 기존 주문등록 테이블(주문기본정보, 주문물품 등)과 충돌하지 않도록
-- 별도의 접두어(ai_quote_)를 사용한다.
-- MS-SQL용. id 는 IDENTITY(자동증가)로 사용한다.
-- 이 스크립트는 여러 번 실행해도 안전하다 (IF NOT EXISTS).
-- ============================================================

-- 0) 단가표: 면적(mm²) 구간별 목재 단가
--    원본: 목재.xlsx (area | misong18T | pyeonbaek18T | pvc18T | mulbau18T | mdf18T)
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'ai_quote_price_table')
BEGIN
  CREATE TABLE ai_quote_price_table (
    areaTier      INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_ai_quote_price PRIMARY KEY,
    areaMax       INT NOT NULL,            -- 이 구간의 상한 면적(mm²). area <= areaMax 인 첫 행 사용.
    misong18T     INT NOT NULL,
    pyeonbaek18T  INT NOT NULL,
    pvc18T        INT NOT NULL,
    mulbau18T     INT NOT NULL,
    mdf18T        INT NOT NULL
  );
  CREATE UNIQUE INDEX UX_ai_quote_area ON ai_quote_price_table(areaMax);
END

-- 1) 세션: 채팅 세션 단위 식별
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'ai_quote_sessions')
BEGIN
  CREATE TABLE ai_quote_sessions (
    sessionId     NVARCHAR(64)  NOT NULL CONSTRAINT PK_ai_quote_sessions PRIMARY KEY,
    createdAt     DATETIME2     NOT NULL CONSTRAINT DF_aiq_sessions_created DEFAULT (SYSDATETIME())
  );
END

-- 2) 주문(견적서 단위): 고객 정보 + 총액 + 이미지
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'ai_quote_orders')
BEGIN
  CREATE TABLE ai_quote_orders (
    orderNo       NVARCHAR(40)  NOT NULL CONSTRAINT PK_ai_quote_orders PRIMARY KEY,
    sessionId     NVARCHAR(64)  NULL CONSTRAINT FK_aiq_orders_session
                    REFERENCES ai_quote_sessions(sessionId) ON DELETE SET NULL,
    custName      NVARCHAR(100) NULL,
    custPhone     NVARCHAR(30)  NULL,
    custEmail     NVARCHAR(200) NULL,
    custMemo      NVARCHAR(1000) NULL,
    itemTotal     BIGINT        NOT NULL CONSTRAINT DF_aiq_orders_itemTotal DEFAULT (0),  -- 물품 소계
    shippingCost  BIGINT        NOT NULL CONSTRAINT DF_aiq_orders_ship DEFAULT (0),       -- 배송비
    grandTotal    BIGINT        NOT NULL CONSTRAINT DF_aiq_orders_grand DEFAULT (0),      -- 총액
    quoteImage    NVARCHAR(500) NULL,   -- uploads/ 내 파일 경로 (예: 20260703_xxx.jpg)
    status        NVARCHAR(20)  NOT NULL CONSTRAINT DF_aiq_orders_status DEFAULT ('pending'),  -- pending|paid|cancelled
    createdAt     DATETIME2     NOT NULL CONSTRAINT DF_aiq_orders_created DEFAULT (SYSDATETIME())
  );
  CREATE INDEX IX_aiq_orders_phone ON ai_quote_orders(custPhone);
END

-- 3) 주문 물품: 한 주문에 여러 라인 (목재/부자재/배송비/할인 모두 포함)
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'ai_quote_items')
BEGIN
  CREATE TABLE ai_quote_items (
    id                  BIGINT IDENTITY(1,1) NOT NULL CONSTRAINT PK_ai_quote_items PRIMARY KEY,
    orderNo             NVARCHAR(40) NOT NULL CONSTRAINT FK_aiq_items_order
                          REFERENCES ai_quote_orders(orderNo) ON DELETE CASCADE,
    itemType            NVARCHAR(20) NOT NULL CONSTRAINT DF_aiq_items_type DEFAULT ('wood'),
                          -- wood | accessory | shipping | discount
    seq                 INT NOT NULL CONSTRAINT DF_aiq_items_seq DEFAULT (0),

    woodType            NVARCHAR(40) NULL,    -- mdf18T 등 키. 부자재/배송비는 NULL
    woodName            NVARCHAR(200) NULL,   -- 표시명
    [width]             INT NULL,             -- 폭(mm). 예약어라 대괄호 처리
    [length]            INT NULL,             -- 길이(mm)
    quantity            INT NOT NULL CONSTRAINT DF_aiq_items_qty DEFAULT (1),
    processingSummary   NVARCHAR(500) NULL,   -- "사선가공(A,B) + 코팅(단면)"
    selectedProcsJson   NVARCHAR(MAX) NULL,   -- 가공 키 배열 JSON (["angled","coating"])
    procDetailsJson     NVARCHAR(MAX) NULL,   -- 가공 상세 객체 JSON
    detailTextJson      NVARCHAR(MAX) NULL,   -- buildDetailText 결과 JSON

    basePrice           BIGINT NULL,          -- 판재 단가
    additionalCost      BIGINT NULL,          -- 가공비
    discount            BIGINT NULL,          -- 이벤트 할인
    discountDetailsJson NVARCHAR(MAX) NULL,   -- 할인 내역 배열 JSON
    estimate            BIGINT NOT NULL CONSTRAINT DF_aiq_items_est DEFAULT (0)  -- 라인 합계
  );
  CREATE INDEX IX_aiq_items_order ON ai_quote_items(orderNo);
END


-- 2) 주문(견적서 단위): 고객 정보 + 총액 + 이미지
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'ai_quote_orders')
BEGIN
  CREATE TABLE ai_quote_orders (
    orderNo       NVARCHAR(40)  NOT NULL CONSTRAINT PK_ai_quote_orders PRIMARY KEY,
    sessionId     NVARCHAR(64)  NULL CONSTRAINT FK_aiq_orders_session
                    REFERENCES ai_quote_sessions(sessionId) ON DELETE SET NULL,
    custName      NVARCHAR(100) NULL,
    custPhone     NVARCHAR(30)  NULL,
    custEmail     NVARCHAR(200) NULL,
    custMemo      NVARCHAR(1000) NULL,
    itemTotal     BIGINT        NOT NULL CONSTRAINT DF_aiq_orders_itemTotal DEFAULT (0),  -- 물품 소계
    shippingCost  BIGINT        NOT NULL CONSTRAINT DF_aiq_orders_ship DEFAULT (0),       -- 배송비
    grandTotal    BIGINT        NOT NULL CONSTRAINT DF_aiq_orders_grand DEFAULT (0),      -- 총액
    quoteImage    NVARCHAR(500) NULL,   -- uploads/ 내 파일 경로 (예: 20260703_xxx.jpg)
    status        NVARCHAR(20)  NOT NULL CONSTRAINT DF_aiq_orders_status DEFAULT ('pending'),  -- pending|paid|cancelled
    createdAt     DATETIME2     NOT NULL CONSTRAINT DF_aiq_orders_created DEFAULT (SYSDATETIME())
  );
  CREATE INDEX IX_aiq_orders_phone ON ai_quote_orders(custPhone);
END

-- 3) 주문 물품: 한 주문에 여러 라인 (목재/부자재/배송비/할인 모두 포함)
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'ai_quote_items')
BEGIN
  CREATE TABLE ai_quote_items (
    id                  BIGINT IDENTITY(1,1) NOT NULL CONSTRAINT PK_ai_quote_items PRIMARY KEY,
    orderNo             NVARCHAR(40) NOT NULL CONSTRAINT FK_aiq_items_order
                          REFERENCES ai_quote_orders(orderNo) ON DELETE CASCADE,
    itemType            NVARCHAR(20) NOT NULL CONSTRAINT DF_aiq_items_type DEFAULT ('wood'),
                          -- wood | accessory | shipping | discount
    seq                 INT NOT NULL CONSTRAINT DF_aiq_items_seq DEFAULT (0),

    woodType            NVARCHAR(40) NULL,    -- mdf18T 등 키. 부자재/배송비는 NULL
    woodName            NVARCHAR(200) NULL,   -- 표시명
    [width]             INT NULL,             -- 폭(mm). 예약어라 대괄호 처리
    [length]            INT NULL,             -- 길이(mm)
    quantity            INT NOT NULL CONSTRAINT DF_aiq_items_qty DEFAULT (1),
    processingSummary   NVARCHAR(500) NULL,   -- "사선가공(A,B) + 코팅(단면)"
    selectedProcsJson   NVARCHAR(MAX) NULL,   -- 가공 키 배열 JSON (["angled","coating"])
    procDetailsJson     NVARCHAR(MAX) NULL,   -- 가공 상세 객체 JSON
    detailTextJson      NVARCHAR(MAX) NULL,   -- buildDetailText 결과 JSON

    basePrice           BIGINT NULL,          -- 판재 단가
    additionalCost      BIGINT NULL,          -- 가공비
    discount            BIGINT NULL,          -- 이벤트 할인
    discountDetailsJson NVARCHAR(MAX) NULL,   -- 할인 내역 배열 JSON
    estimate            BIGINT NOT NULL CONSTRAINT DF_aiq_items_est DEFAULT (0)  -- 라인 합계
  );
  CREATE INDEX IX_aiq_items_order ON ai_quote_items(orderNo);
END
