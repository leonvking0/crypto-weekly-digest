const pptxgen = require("pptxgenjs");

const C = {
  navy: "1A2744", darkNavy: "0F1A2E", gold: "C9A96E", white: "FFFFFF",
  offWhite: "F5F6F8", lightGray: "E8EBF0", medGray: "8C95A6",
  bodyText: "2D3748", red: "C53030", green: "2F855A",
};
const makeShadow = () => ({ type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.12 });
const WEEK = "2026.03.23\u201303.29";
const FOOTER_TEXT = `Crypto Weekly | ${WEEK}`;

async function generate() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "42";
  pres.title = `Crypto Weekly ${WEEK}`;

  // ============ SLIDE 1: Title ============
  const s1 = pres.addSlide();
  s1.background = { color: C.darkNavy };
  s1.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.6, w: 1.2, h: 0.04, fill: { color: C.gold } });
  s1.addText("Crypto Market Weekly", {
    x: 0.8, y: 1.8, w: 8, h: 1.0,
    fontSize: 42, fontFace: "Arial", bold: true, color: C.white, margin: 0
  });
  s1.addText("March 23 \u2013 29, 2026", {
    x: 0.8, y: 2.7, w: 8, h: 0.5,
    fontSize: 20, fontFace: "Arial", color: C.gold, margin: 0
  });
  s1.addText("Sources: Messari \u00B7 The Block \u00B7 CoinGecko \u00B7 DefiLlama", {
    x: 0.8, y: 3.4, w: 8, h: 0.4,
    fontSize: 11, fontFace: "Arial", color: C.medGray, margin: 0
  });
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.gold } });
  s1.addText("CONFIDENTIAL \u2014 FOR LP DISTRIBUTION ONLY", {
    x: 0, y: 5.25, w: 10, h: 0.375,
    fontSize: 9, fontFace: "Arial", color: C.darkNavy, align: "center", valign: "middle", bold: true
  });
  s1.addNotes("This weekly report covers March 23-29, 2026. Markets operated under extreme duress from geopolitical risks (Israel-Iran conflict escalation, Strait of Hormuz shipping disruptions) and Trump tariff policy escalation (global tariffs raised to 15%). Fear & Greed Index dropped to 8-9, lowest since the Luna crash in 2022.\n\nKey drivers:\n1. Fed revised 2026 PCE inflation forecast up to 2.7%, pushing rate cut expectations further out\n2. Iran threatened to blockade the Strait of Hormuz, oil prices surged 25%+ vs pre-conflict\n3. BTC/ETH/SOL ETFs posted simultaneous net outflows on the same day\n\nHowever, a landmark regulatory development occurred: SEC-CFTC joint rules formally classified 16 crypto assets as digital commodities, ending a decade of regulatory ambiguity. This clears the path for SOL/XRP spot ETFs.");

  // ============ SLIDE 2: Market Overview ============
  const s2 = pres.addSlide();
  s2.background = { color: C.offWhite };
  s2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy } });
  s2.addText("\uD83D\uDCCA  Market Overview", {
    x: 0.6, y: 0, w: 9, h: 0.7,
    fontSize: 20, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0
  });

  const cardY = 0.95, cardH = 1.05, cardW = 2.8, cardGap = 0.3, cardStartX = 0.6;

  // BTC
  s2.addShape(pres.shapes.RECTANGLE, { x: cardStartX, y: cardY, w: cardW, h: cardH, fill: { color: C.white }, shadow: makeShadow() });
  s2.addShape(pres.shapes.RECTANGLE, { x: cardStartX, y: cardY, w: 0.06, h: cardH, fill: { color: C.gold } });
  s2.addText("BTC", { x: cardStartX + 0.25, y: cardY + 0.1, w: 2, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.medGray, bold: true, margin: 0 });
  s2.addText("$66,445", { x: cardStartX + 0.25, y: cardY + 0.35, w: 2, h: 0.4, fontSize: 28, fontFace: "Arial", color: C.bodyText, bold: true, margin: 0 });
  s2.addText("-3.5%", { x: cardStartX + 0.25, y: cardY + 0.72, w: 2, h: 0.25, fontSize: 14, fontFace: "Arial", color: C.red, bold: true, margin: 0 });

  // ETH
  const c2x = cardStartX + cardW + cardGap;
  s2.addShape(pres.shapes.RECTANGLE, { x: c2x, y: cardY, w: cardW, h: cardH, fill: { color: C.white }, shadow: makeShadow() });
  s2.addShape(pres.shapes.RECTANGLE, { x: c2x, y: cardY, w: 0.06, h: cardH, fill: { color: C.gold } });
  s2.addText("ETH", { x: c2x + 0.25, y: cardY + 0.1, w: 2, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.medGray, bold: true, margin: 0 });
  s2.addText("$1,996", { x: c2x + 0.25, y: cardY + 0.35, w: 2, h: 0.4, fontSize: 28, fontFace: "Arial", color: C.bodyText, bold: true, margin: 0 });
  s2.addText("-4.1%", { x: c2x + 0.25, y: cardY + 0.72, w: 2, h: 0.25, fontSize: 14, fontFace: "Arial", color: C.red, bold: true, margin: 0 });

  // Fear & Greed
  const c3x = c2x + cardW + cardGap;
  s2.addShape(pres.shapes.RECTANGLE, { x: c3x, y: cardY, w: cardW, h: cardH, fill: { color: C.white }, shadow: makeShadow() });
  s2.addShape(pres.shapes.RECTANGLE, { x: c3x, y: cardY, w: 0.06, h: cardH, fill: { color: C.red } });
  s2.addText("Fear & Greed", { x: c3x + 0.25, y: cardY + 0.1, w: 2.4, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.medGray, bold: true, margin: 0 });
  s2.addText("9", { x: c3x + 0.25, y: cardY + 0.35, w: 1, h: 0.4, fontSize: 28, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
  s2.addText("Extreme Fear", { x: c3x + 1.1, y: cardY + 0.42, w: 1.5, h: 0.3, fontSize: 13, fontFace: "Arial", color: C.red, margin: 0 });

  // Price table
  const tableY = 2.25;
  const headerRow = [
    { text: "Asset", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 11, fontFace: "Arial", align: "center" } },
    { text: "Price", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 11, fontFace: "Arial", align: "center" } },
    { text: "7d Change", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 11, fontFace: "Arial", align: "center" } },
    { text: "Notes", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 11, fontFace: "Arial", align: "center" } },
  ];
  const makeRow = (asset, price, chg, note, isNeg) => [
    { text: asset, options: { fontSize: 11, fontFace: "Arial", bold: true, align: "center", color: C.bodyText } },
    { text: price, options: { fontSize: 11, fontFace: "Arial", align: "right", color: C.bodyText } },
    { text: chg, options: { fontSize: 11, fontFace: "Arial", align: "center", color: isNeg ? C.red : C.green, bold: true } },
    { text: note, options: { fontSize: 10, fontFace: "Arial", color: C.medGray } },
  ];
  const tableData = [
    headerRow,
    makeRow("SOL", "$81.63", "-6.7%", "SEC commodity classification not yet priced in", true),
    makeRow("HYPE", "$38.91", "+2.3%", "Perp DEX leader; weekly fees $14M", false),
    makeRow("PENDLE", "$1.12", "-6.2%", "Yield token sector under pressure", true),
    makeRow("ENA", "$0.09", "-8.4%", "Weakening with broader market", true),
    makeRow("DYDX", "$0.09", "+11.6%", "Buyback expanded to 75% net revenue", false),
  ];
  s2.addTable(tableData, {
    x: 0.6, y: tableY, w: 8.8, colW: [1.3, 1.8, 1.5, 4.2],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.35, 0.33, 0.33, 0.33, 0.33, 0.33],
    autoPage: false,
  });

  s2.addText("Macro: Israel-Iran conflict + Trump 15% tariffs + Fed PCE revision to 2.7% \u2192 risk-off across all assets; BTC/ETH/SOL ETFs same-day net outflows", {
    x: 0.6, y: 4.55, w: 8.8, h: 0.4,
    fontSize: 10, fontFace: "Arial", color: C.medGray, italic: true, margin: 0
  });

  s2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s2.addText(FOOTER_TEXT, { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  s2.addNotes("Market Overview Deep Dive:\n\n1. BTC fell to $66,445, -3.5% WoW: Primary catalyst was the $14.16B BTC options expiry on March 27, with heavy short bets paying off. ETF weekly net outflows of $296M signal institutional defensive positioning. Technically, $66,000 is critical support; a daily close below opens the path to $50,000.\n\n2. ETH broke below $2,000 psychological level: Post-Fusaka upgrade L1 fee revenue plummeted, L2 activity siphoning accelerated. A founder transferred ~79,000 ETH to Kraken adding direct sell pressure. Down 60% from August high. ETH/BTC ratio continues weakening.\n\n3. Fear & Greed at 8-9, extreme fear zone: Historically this level often indicates a medium-term bottom area, but requires macro cooperation to confirm reversal. Current Israel-Iran conflict shows no signs of de-escalation.\n\n4. Total MCap $2.37T, BTC dominance 56.1%: Altcoins broadly underperforming BTC as capital concentrates in safe havens. SOL commodity classification upside not yet reflected in price.\n\n5. HYPE bucked the trend at +2.3%: Hyperliquid\u2019s weekly fee revenue of $14M (new record, annualized >$600M) validates the thesis that protocols with real revenue are more resilient in bear markets.\n\nPortfolio implications: Consider defensive positioning. HYPE showing relative strength. SOL/XRP ETF catalyst could be a Q2 inflection point if macro improves.");

  // ============ SLIDE 3: Top Stories + Industry ============
  const s3 = pres.addSlide();
  s3.background = { color: C.offWhite };
  s3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy } });
  s3.addText("\uD83D\uDD25  Top Stories & Industry", {
    x: 0.6, y: 0, w: 9, h: 0.7,
    fontSize: 20, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0
  });

  s3.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 4.1, fill: { color: C.white }, shadow: makeShadow() });
  s3.addText("TOP 5 STORIES", {
    x: 0.6, y: 0.95, w: 4, h: 0.35,
    fontSize: 13, fontFace: "Arial", bold: true, color: C.navy, margin: 0
  });

  const stories = [
    ["1", "SEC-CFTC: 16 Crypto Assets Get Commodity Status", "Ends decade of regulatory ambiguity; clears SOL/XRP ETF path"],
    ["2", "CLARITY Act Bans Stablecoin Yield; Circle -20%", "$2B market cap wiped; stablecoin business model at risk"],
    ["3", "Resolv USR Exploited: $25M Stolen, 97.5% Depeg", "AWS key compromised; collapsed in 17 minutes on Curve"],
    ["4", "Geopolitical + Tariff Shock Crushes Market", "Israel-Iran war + Trump 15% tariffs; ETFs all net outflows"],
    ["5", "BTC Miners Pivot to AI: Mining Cost $80K vs $68K Price", "$70B+ AI contracts signed; 15,000+ BTC sold for funding"],
  ];

  stories.forEach((s, i) => {
    const itemY = 1.38 + i * 0.72;
    s3.addShape(pres.shapes.OVAL, { x: 0.65, y: itemY, w: 0.3, h: 0.3, fill: { color: C.gold } });
    s3.addText(s[0], { x: 0.65, y: itemY, w: 0.3, h: 0.3, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, align: "center", valign: "middle", margin: 0 });
    s3.addText(s[1], { x: 1.1, y: itemY - 0.02, w: 3.6, h: 0.25, fontSize: 11, fontFace: "Arial", bold: true, color: C.bodyText, margin: 0 });
    s3.addText(s[2], { x: 1.1, y: itemY + 0.22, w: 3.6, h: 0.25, fontSize: 9, fontFace: "Arial", color: C.medGray, margin: 0 });
  });

  // Right: L1/L2
  s3.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 4.5, h: 1.95, fill: { color: C.white }, shadow: makeShadow() });
  s3.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s3.addText("Layer 1 & Layer 2", { x: 5.3, y: 0.9, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  const l1Items = [
    "16 tokens get SEC-CFTC commodity classification (BTC/ETH/SOL etc)",
    "Base TVL $4.15B, ~50% of L2 DEX volume",
    "Linea Yield Boost: ETH auto-staked via Lido V3",
    "L2/L1 DAU ratio dropped to 1.12 from peak 10.43",
  ];
  s3.addText(l1Items.map((t, i) => ({
    text: t,
    options: { bullet: true, breakLine: i < l1Items.length - 1, fontSize: 10, fontFace: "Arial", color: C.bodyText }
  })), { x: 5.3, y: 1.32, w: 4.1, h: 1.45, margin: 0, paraSpaceAfter: 4 });

  // DeFi
  s3.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 3.05, w: 4.5, h: 1.95, fill: { color: C.white }, shadow: makeShadow() });
  s3.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 3.05, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s3.addText("DeFi & Perp DEXes", { x: 5.3, y: 3.05, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  const defiItems = [
    "Aave cumulative lending exceeds $1T; TVL $27.3B (62.8% share)",
    "DeFi TVL rose 4.4% despite extreme fear sentiment",
    "Hyperliquid: 54% perp DEX share; weekly fees $14M (+56%)",
    "dYdX buyback to 75% net rev; synthetic equity perps planned",
  ];
  s3.addText(defiItems.map((t, i) => ({
    text: t,
    options: { bullet: true, breakLine: i < defiItems.length - 1, fontSize: 10, fontFace: "Arial", color: C.bodyText }
  })), { x: 5.3, y: 3.47, w: 4.1, h: 1.45, margin: 0, paraSpaceAfter: 4 });

  s3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s3.addText(FOOTER_TEXT, { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  s3.addNotes("Top Stories Deep Dive:\n\n1. SEC-CFTC Joint Rules: The single most important regulatory development in crypto history. Formally classified 16 tokens as digital commodities under CFTC oversight. SEC Chair Atkins acknowledged most crypto assets are not securities. Creates safe harbors for mining, staking, and no-consideration airdrops. Immediate impact: clears SOL/XRP spot ETF path. Long-term: dramatically reduces project legal risk and institutional hesitancy.\n\n2. CLARITY Act: Senate Banking Committee draft bans stablecoin passive yield (\"economically equivalent to interest\"). Circle derives 95% of revenue from reserve interest. 20% stock crash reflects existential threat to their business model. Coinbase down 10% on collateral damage. Watch: full committee markup expected late April.\n\n3. Resolv USR: Attacker compromised AWS KMS privileged key, minted 80M unbacked USR with <$200K collateral. Collapsed 97.5% in 17 minutes on Curve. Cascaded into Morpho lending vaults (oracles still valued USR at $1.13). Lessons: centralized cloud key management is a critical vulnerability; DeFi composability amplifies contagion.\n\n4. Geopolitical shock: Iran-Israel conflict + Strait of Hormuz threats + oil +25% + Trump 15% tariffs created perfect storm for risk assets. First time in 2026 that BTC/ETH/SOL ETFs posted simultaneous net outflows.\n\n5. Mining economics inverted: All-in cost $79,995/BTC vs $68K market price = $19K loss per coin. Miners pivoting aggressively: $70B+ in AI/HPC contracts signed. Mining difficulty dropped 7.76% on 3/21 (2nd largest drop of the year). Public miners sold 15,000+ BTC to fund AI transition.\n\nDeFi: Aave\u2019s $1T cumulative lending is a milestone. But ACI governance exit exposes DAO centralization issues. ECB report named Aave, MakerDAO, Uniswap: top 100 holders control >80% voting power. Hyperliquid\u2019s crude oil perps doing >$2.2B daily \u2014 perp DEX growth now extends beyond crypto into commodities.");

  // ============ SLIDE 4: Regulation + Fundraising ============
  const s4 = pres.addSlide();
  s4.background = { color: C.offWhite };
  s4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy } });
  s4.addText("\u2696\uFE0F  Regulation & Fundraising", {
    x: 0.6, y: 0, w: 9, h: 0.7,
    fontSize: 20, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0
  });

  s4.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 4.1, fill: { color: C.white }, shadow: makeShadow() });
  s4.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 0.06, h: 4.1, fill: { color: C.gold } });
  s4.addText("Regulation & Policy", { x: 0.7, y: 0.95, w: 4, h: 0.35, fontSize: 14, fontFace: "Arial", bold: true, color: C.navy, margin: 0 });

  const regItems = [
    { title: "SEC-CFTC Joint Commodity Rules", desc: "16 tokens classified as commodities; ends regulation-by-enforcement" },
    { title: "CLARITY Act Draft", desc: "Bans stablecoin passive yield; Circle stock crashes 20%" },
    { title: "CFTC Innovation Task Force", desc: "Covers digital assets, AI, and prediction markets oversight" },
    { title: "House Tokenization Hearing", desc: "Financial Services Committee on blockchain + capital markets" },
    { title: "NYSE \u00D7 Securitize Platform", desc: "SEC approves tokenized equities 24/7 trading + on-chain settlement" },
  ];
  regItems.forEach((item, i) => {
    const iy = 1.45 + i * 0.65;
    s4.addText(item.title, { x: 0.7, y: iy, w: 4, h: 0.25, fontSize: 11, fontFace: "Arial", bold: true, color: C.bodyText, margin: 0 });
    s4.addText(item.desc, { x: 0.7, y: iy + 0.23, w: 4, h: 0.25, fontSize: 10, fontFace: "Arial", color: C.medGray, margin: 0 });
  });

  s4.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 4.5, h: 4.1, fill: { color: C.white }, shadow: makeShadow() });
  s4.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 0.06, h: 4.1, fill: { color: C.gold } });
  s4.addText("Fundraising & M&A", { x: 5.4, y: 0.95, w: 4, h: 0.35, fontSize: 14, fontFace: "Arial", bold: true, color: C.navy, margin: 0 });

  const fundHeader = [
    { text: "Project", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial" } },
    { text: "Amount", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial", align: "right" } },
    { text: "Notes", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial" } },
  ];
  const fundRow = (name, amt, note) => [
    { text: name, options: { fontSize: 10, fontFace: "Arial", color: C.bodyText, bold: true } },
    { text: amt, options: { fontSize: 10, fontFace: "Arial", color: C.bodyText, align: "right" } },
    { text: note, options: { fontSize: 9, fontFace: "Arial", color: C.medGray } },
  ];
  const fundData = [
    fundHeader,
    fundRow("MC \u00D7 BVNK", "$1.8B", "TradFi deep into crypto infra"),
    fundRow("Kalshi Series E", "$1B", "Prediction market record valuation"),
    fundRow("Polymarket", "$600M", "ICE-led; TradFi endorsement"),
    fundRow("Core Scientific", "$1B", "JPM+MS credit; AI chip expansion"),
    fundRow("GSR Acquires", "$57M", "Autonomous + Architech"),
  ];
  s4.addTable(fundData, {
    x: 5.3, y: 1.45, w: 4.1, colW: [1.3, 0.8, 2.0],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.3, 0.3, 0.3, 0.3, 0.3, 0.3],
    autoPage: false,
  });

  s4.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 3.4, w: 4.1, h: 0.7, fill: { color: "FFF8E7" }, shadow: makeShadow() });
  s4.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 3.4, w: 0.06, h: 0.7, fill: { color: C.gold } });
  s4.addText([
    { text: "Highlight: ", options: { bold: true, fontSize: 10, fontFace: "Arial", color: C.navy } },
    { text: "22 deals totaling ~$3.28B disclosed. 2026 crypto M&A projected to exceed $37B record. TradFi accelerating crypto infrastructure acquisitions.", options: { fontSize: 10, fontFace: "Arial", color: C.bodyText } }
  ], { x: 5.55, y: 3.45, w: 3.7, h: 0.6, margin: 0, valign: "middle" });

  s4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s4.addText(FOOTER_TEXT, { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  s4.addNotes("Regulation & Fundraising Analysis:\n\nRegulatory landscape:\n1. SEC-CFTC joint commodity classification is the most consequential US crypto regulation ever. Creates five-category taxonomy (digital commodities, collectibles, tools, stablecoins, securities) with explicit safe harbors for mining, staking, and airdrops. Projects no longer need to guess whether their token will be sued as a security. Long-term: reduces compliance costs, attracts more traditional capital.\n\n2. CLARITY Act stablecoin yield ban is existential for Circle (95% revenue from reserve interest). If passed, Circle must pivot business model or face revenue collapse. Banks welcome this as stablecoin yield threatens traditional deposit business. Outcome determines whether stablecoins become \"digital cash\" or \"digital savings accounts.\"\n\n3. NYSE-Securitize digital trading platform: SEC approved tokenized US equities with 24/7 trading + on-chain stablecoin settlement. Traditional finance infrastructure actively embracing blockchain.\n\nFundraising:\n1. Mastercard $1.8B BVNK acquisition: largest TradFi-crypto acquisition. Signal that traditional finance is buying, not just watching.\n2. Kalshi $1B Series E + Polymarket $600M (ICE-led): prediction market sector receives unprecedented institutional validation. ICE is NYSE's parent company.\n3. 22 deals, ~$3.28B disclosed. 2026 crypto M&A on track to exceed $37B.\n\nRisk assessment: regulatory clarity is bullish medium-term but CLARITY Act stablecoin yield provisions create near-term uncertainty for portfolio holdings in stablecoin ecosystem (Ethena, potentially Pendle).");

  // ============ SLIDE 5: Security + Sectors ============
  const s5 = pres.addSlide();
  s5.background = { color: C.offWhite };
  s5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy } });
  s5.addText("\uD83D\uDD13  Security & Hot Sectors", {
    x: 0.6, y: 0, w: 9, h: 0.7,
    fontSize: 20, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0
  });

  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 2.35, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 0.35, fill: { color: C.red } });
  s5.addText("Security  |  YTD Loss: ~$480M", { x: 0.6, y: 0.9, w: 4.2, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  const secTable = [
    [
      { text: "Incident", options: { bold: true, fontSize: 9, fontFace: "Arial", fill: { color: C.lightGray }, color: C.bodyText } },
      { text: "Loss", options: { bold: true, fontSize: 9, fontFace: "Arial", fill: { color: C.lightGray }, color: C.bodyText, align: "right" } },
      { text: "Vector", options: { bold: true, fontSize: 9, fontFace: "Arial", fill: { color: C.lightGray }, color: C.bodyText } },
    ],
    [{ text: "Iran Exchange", options: { fontSize: 9, fontFace: "Arial" } }, { text: "~$90M", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "Suspected state actor", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
    [{ text: "Resolv USR", options: { fontSize: 9, fontFace: "Arial" } }, { text: "~$25M", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "AWS key compromise", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
    [{ text: "Infiniti/GhostClaw", options: { fontSize: 9, fontFace: "Arial" } }, { text: "Undisclosed", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "Malicious npm + wallet drain", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
    [{ text: "Fake stablecoins", options: { fontSize: 9, fontFace: "Arial" } }, { text: "54,000+", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "Approval phishing", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
  ];
  s5.addTable(secTable, {
    x: 0.55, y: 1.35, w: 4.2, colW: [1.4, 1.1, 1.7],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.28, 0.28, 0.28, 0.28, 0.28],
    autoPage: false,
  });

  // Stablecoins
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.5, w: 4.5, h: 1.5, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.5, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s5.addText("Stablecoins", { x: 0.6, y: 3.5, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });
  s5.addText([
    { text: "USDC volume surpasses USDT (64% share); 86% institutions hold USDC", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "CLARITY Act threatens yield model; Tether launches full audit counter", options: { bullet: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
  ], { x: 0.6, y: 3.95, w: 4.1, h: 0.9, margin: 0, paraSpaceAfter: 4 });

  // AI x Crypto
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 4.5, h: 1.85, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s5.addText("AI \u00D7 Crypto", { x: 5.3, y: 0.9, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });
  s5.addText([
    { text: "NVIDIA GTC aftershock: Huang declares \"AGI-level\" reached", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "AI crypto sector MCap >$28B; TAO leads at $3.4B", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "Mining cost $80K vs $68K price; $70B+ AI contracts signed", options: { bullet: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
  ], { x: 5.3, y: 1.35, w: 4.1, h: 1.3, margin: 0, paraSpaceAfter: 4 });

  // Prediction Markets
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 2.95, w: 4.5, h: 2.05, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 2.95, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s5.addText("Prediction Markets", { x: 5.3, y: 2.95, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });
  s5.addText([
    { text: "Polymarket: 11,176 active markets; 840K monthly wallets", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "Industry monthly volume exceeds $21B; Kalshi $1B Series E", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "CFTC integrates Polymarket; anti-insider trading rules go live", options: { bullet: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
  ], { x: 5.3, y: 3.4, w: 4.1, h: 1.5, margin: 0, paraSpaceAfter: 4 });

  s5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s5.addText(FOOTER_TEXT, { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  s5.addNotes("Security & Sector Analysis:\n\nSecurity incidents:\n1. Iran exchange $90M hack: Geopolitical context \u2014 state-level hacking activity increases during Israel-Iran conflict. Not just a financial loss but reflects crypto infrastructure vulnerability in geopolitical hotspots.\n\n2. Resolv USR exploit \u2014 critical lessons: Attacker compromised AWS KMS privileged key, exposing over-reliance on centralized cloud services. The cascading effect through Morpho vaults (oracles still pricing USR at $1.13) demonstrates how DeFi composability becomes a systemic risk conductor during crises.\n\n3. YTD: 103 security incidents + 36 phishing scams, ~$480M total losses. Social engineering has surpassed code vulnerabilities as the primary attack vector. CertiK warns full-year losses may exceed $500M.\n\nAI x Crypto:\nPost-NVIDIA GTC sector MCap >$28B. The real story is BTC miner pivots: all-in cost $79,995/BTC vs $68K market price = $19K loss per coin. Miners have signed $70B+ in AI/HPC contracts. By year-end, AI revenue expected to grow from <15% to 70%+ of miner revenue. This is a fundamental industry inflection point.\n\nPrediction Markets:\nPolymarket monthly volume surpasses $21B (7.5x YoY). CFTC self-certification is a milestone \u2014 first on-chain prediction market formally integrated into US regulatory framework. Anti-insider trading rules signal sector maturation.\n\nStablecoins:\nUSDC surpassing USDT in transaction volume (64% share) is a watershed moment. CLARITY Act creates existential risk for yield-dependent stablecoin businesses. Tether hiring Big Four auditors is a strategic counter-move. Portfolio impact: ENA directly exposed to stablecoin yield narrative; monitor CLARITY Act progress closely.");

  // ============ SLIDE 6: Catalysts + Outlook ============
  const s6 = pres.addSlide();
  s6.background = { color: C.offWhite };
  s6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy } });
  s6.addText("\uD83D\uDD2E  Catalysts & Outlook", {
    x: 0.6, y: 0, w: 9, h: 0.7,
    fontSize: 20, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0
  });

  const catHeader = [
    { text: "Date", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial", align: "center" } },
    { text: "Event", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial" } },
    { text: "Asset", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial", align: "center" } },
    { text: "Expected Impact", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial" } },
  ];
  const catRow = (date, event, asset, impact, highlight = false) => [
    { text: date, options: { fontSize: 10, fontFace: "Arial", align: "center", color: C.bodyText, fill: highlight ? { color: "FFF3E0" } : undefined } },
    { text: event, options: { fontSize: 10, fontFace: "Arial", color: C.bodyText, bold: highlight, fill: highlight ? { color: "FFF3E0" } : undefined } },
    { text: asset, options: { fontSize: 10, fontFace: "Arial", align: "center", color: C.navy, bold: true, fill: highlight ? { color: "FFF3E0" } : undefined } },
    { text: impact, options: { fontSize: 9, fontFace: "Arial", color: C.medGray, fill: highlight ? { color: "FFF3E0" } : undefined } },
  ];
  const catData = [
    catHeader,
    catRow("4/3", "Wormhole 600M token unlock (6% supply)", "W", "Supply pressure >$90M"),
    catRow("4/6", "Hyperliquid core contributor unlock", "HYPE", "Historically high-risk sell category", true),
    catRow("4/15", "Starknet 127M token unlock", "STRK", "Supply pressure"),
    catRow("Late Apr", "CLARITY Act full committee markup", "Stables", "Stablecoin yield ban decision", true),
    catRow("4/28-29", "Powell\u2019s final FOMC meeting", "All", "Rate guidance + successor signals"),
    catRow("4/29-30", "TOKEN2049 Dubai", "All", "Major industry summit"),
    catRow("Q2", "SOL/XRP spot ETF progress", "SOL/XRP", "Commodity classification opens ETF path"),
  ];
  s6.addTable(catData, {
    x: 0.5, y: 0.9, w: 9, colW: [0.9, 3.2, 1.2, 3.7],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.32, 0.32, 0.32, 0.32, 0.32, 0.32, 0.32, 0.32],
    autoPage: false,
  });

  s6.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.65, w: 9, h: 1.3, fill: { color: C.white }, shadow: makeShadow() });
  s6.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.65, w: 0.06, h: 1.3, fill: { color: C.gold } });
  s6.addText("Weekly Summary", { x: 0.8, y: 3.7, w: 8, h: 0.3, fontSize: 13, fontFace: "Arial", bold: true, color: C.navy, margin: 0 });
  s6.addText([
    { text: "Extreme fear (FGI=9) driven by geopolitical risks and tariff policy. SEC-CFTC commodity classification is a major long-term positive but failed to offset near-term macro headwinds.", options: { breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "Key watch: ", options: { bold: true, fontSize: 10, fontFace: "Arial", color: C.navy } },
    { text: "4/6 HYPE unlock (portfolio-relevant), CLARITY Act markup impact on stablecoin sector, Israel-Iran conflict trajectory as macro overhang.", options: { fontSize: 10, fontFace: "Arial", color: C.bodyText } },
  ], { x: 0.8, y: 4.0, w: 8.5, h: 0.85, margin: 0 });

  s6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s6.addText(FOOTER_TEXT + "  |  CONFIDENTIAL", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  s6.addNotes("Catalysts & Outlook Deep Dive:\n\nCritical catalysts ranked by portfolio impact:\n\n1. 4/6 HYPE Core Contributor Unlock: Directly portfolio-relevant. Historically, core team unlocks carry high sell pressure risk. However, Hyperliquid fundamentals are exceptionally strong ($14M weekly fees, $600M+ annualized). If team holds or provides liquidity rather than sells, could be a confidence catalyst. Monitor on-chain large transfers pre/post unlock.\n\n2. CLARITY Act Full Committee Markup (late April): Determines the fate of stablecoin yield. If yield ban passes, major headwind for Circle, Coinbase, and the broader stablecoin yield ecosystem (including Ethena). If amended or removed, stablecoin sector could rally hard.\n\n3. Powell\u2019s Final FOMC (4/28-29): Last meeting of Powell\u2019s tenure. Market will scrutinize signals about successor and rate path. Dovish surprise could trigger risk-on rally.\n\n4. Wormhole 6% Unlock (4/3): >$90M supply pressure. Cross-chain bridge sector is competitive; watch for team mitigation mechanisms.\n\nWeekly Summary:\nMarkets are in a multi-factor headwind zone (geopolitical conflict + tariffs + inflation revision + ETF outflows). Fear index at extreme levels. However, regulatory fundamentals improved dramatically with the commodity classification milestone. This combination of improving fundamentals + extreme sentiment historically indicates a medium-term buying opportunity. The key variable is when macro improves \u2014 Israel-Iran trajectory and Fed stance are the determining factors.\n\nPortfolio action items:\n- Monitor HYPE unlock risk; consider partial hedge if available\n- Track SOL/XRP ETF catalysts in Q2\n- Watch CLARITY Act markup for stablecoin exposure (ENA, PENDLE)\n- Maintain defensive positioning until geopolitical risks show de-escalation signals");

  const outPath = "/Users/lx/Github/crypto-weekly-digest/reports/2026-03-23/slides-en.pptx";
  await pres.writeFile({ fileName: outPath });
  console.log("\u2705 Written to " + outPath);
}

generate().catch(e => { console.error(e); process.exit(1); });
