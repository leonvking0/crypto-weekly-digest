const pptxgen = require("pptxgenjs");

const C = {
  navy: "1A2744", darkNavy: "0F1A2E", gold: "C9A96E", white: "FFFFFF",
  offWhite: "F5F6F8", lightGray: "E8EBF0", medGray: "8C95A6",
  bodyText: "2D3748", red: "C53030", green: "2F855A",
};
const makeShadow = () => ({ type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.12 });
const WEEK = "2026.03.16\u201303.22";
const footer = (s, pres) => {
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s.addText(`Crypto Weekly Digest | ${WEEK}`, { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });
};

async function generate() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "42";
  pres.title = `Crypto Weekly Digest ${WEEK}`;

  // ===== SLIDE 1: Title =====
  const s1 = pres.addSlide();
  s1.background = { color: C.darkNavy };
  s1.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.6, w: 1.2, h: 0.04, fill: { color: C.gold } });
  s1.addText("Crypto Weekly Digest", { x: 0.8, y: 1.8, w: 8, h: 1.0, fontSize: 42, fontFace: "Arial", bold: true, color: C.white, margin: 0 });
  s1.addText("March 16 \u2014 22, 2026", { x: 0.8, y: 2.7, w: 8, h: 0.5, fontSize: 20, fontFace: "Arial", color: C.gold, margin: 0 });
  s1.addText("Sources: Messari \u00B7 The Block \u00B7 CoinGecko \u00B7 DefiLlama", { x: 0.8, y: 3.4, w: 8, h: 0.4, fontSize: 11, fontFace: "Arial", color: C.medGray, margin: 0 });
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.gold } });
  s1.addText("CONFIDENTIAL \u2014 FOR LP DISTRIBUTION ONLY", { x: 0, y: 5.25, w: 10, h: 0.375, fontSize: 9, fontFace: "Arial", color: C.darkNavy, align: "center", valign: "middle", bold: true });
  s1.addNotes("This week the crypto market experienced a rollercoaster ride. BTC rallied to a 6-week high of $76K mid-week driven by landmark SEC/CFTC regulatory guidance, then flash-crashed to $68.8K on Saturday due to Middle East geopolitical crisis (Strait of Hormuz escalation, oil spiking to $110/barrel). Over $1B in liquidations occurred within 24 hours (85% longs). Fear & Greed Index dropped to 10 (Extreme Fear), lowest since 2024. Despite short-term volatility, the regulatory breakthrough (18 tokens classified as digital commodities) lays an extremely positive foundation for the medium-to-long term.");

  // ===== SLIDE 2: Market Overview =====
  const s2 = pres.addSlide();
  s2.background = { color: C.offWhite };
  s2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy } });
  s2.addText("\uD83D\uDCCA  Market Overview", { x: 0.6, y: 0, w: 9, h: 0.7, fontSize: 20, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  const cardY = 0.95, cardH = 1.05, cardW = 2.8, cardGap = 0.3, cardStartX = 0.6;
  // BTC
  s2.addShape(pres.shapes.RECTANGLE, { x: cardStartX, y: cardY, w: cardW, h: cardH, fill: { color: C.white }, shadow: makeShadow() });
  s2.addShape(pres.shapes.RECTANGLE, { x: cardStartX, y: cardY, w: 0.06, h: cardH, fill: { color: C.gold } });
  s2.addText("BTC", { x: cardStartX + 0.25, y: cardY + 0.1, w: 2, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.medGray, bold: true, margin: 0 });
  s2.addText("$68,811", { x: cardStartX + 0.25, y: cardY + 0.35, w: 2, h: 0.4, fontSize: 28, fontFace: "Arial", color: C.bodyText, bold: true, margin: 0 });
  s2.addText("-3.9%", { x: cardStartX + 0.25, y: cardY + 0.72, w: 2, h: 0.25, fontSize: 14, fontFace: "Arial", color: C.red, bold: true, margin: 0 });

  const c2x = cardStartX + cardW + cardGap;
  s2.addShape(pres.shapes.RECTANGLE, { x: c2x, y: cardY, w: cardW, h: cardH, fill: { color: C.white }, shadow: makeShadow() });
  s2.addShape(pres.shapes.RECTANGLE, { x: c2x, y: cardY, w: 0.06, h: cardH, fill: { color: C.gold } });
  s2.addText("ETH", { x: c2x + 0.25, y: cardY + 0.1, w: 2, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.medGray, bold: true, margin: 0 });
  s2.addText("$2,082", { x: c2x + 0.25, y: cardY + 0.35, w: 2, h: 0.4, fontSize: 28, fontFace: "Arial", color: C.bodyText, bold: true, margin: 0 });
  s2.addText("-0.6%", { x: c2x + 0.25, y: cardY + 0.72, w: 2, h: 0.25, fontSize: 14, fontFace: "Arial", color: C.red, bold: true, margin: 0 });

  const c3x = c2x + cardW + cardGap;
  s2.addShape(pres.shapes.RECTANGLE, { x: c3x, y: cardY, w: cardW, h: cardH, fill: { color: C.white }, shadow: makeShadow() });
  s2.addShape(pres.shapes.RECTANGLE, { x: c3x, y: cardY, w: 0.06, h: cardH, fill: { color: C.red } });
  s2.addText("Fear & Greed", { x: c3x + 0.25, y: cardY + 0.1, w: 2.4, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.medGray, bold: true, margin: 0 });
  s2.addText("10", { x: c3x + 0.25, y: cardY + 0.35, w: 1, h: 0.4, fontSize: 28, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
  s2.addText("Extreme Fear", { x: c3x + 1.1, y: cardY + 0.42, w: 1.5, h: 0.3, fontSize: 13, fontFace: "Arial", color: C.red, margin: 0 });

  const headerRow = [
    { text: "Asset", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 11, fontFace: "Arial", align: "center" } },
    { text: "Price", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 11, fontFace: "Arial", align: "center" } },
    { text: "7d Chg", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 11, fontFace: "Arial", align: "center" } },
    { text: "Notes", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 11, fontFace: "Arial", align: "center" } },
  ];
  const makeRow = (asset, price, chg, note, isNeg) => [
    { text: asset, options: { fontSize: 11, fontFace: "Arial", bold: true, align: "center", color: C.bodyText } },
    { text: price, options: { fontSize: 11, fontFace: "Arial", align: "right", color: C.bodyText } },
    { text: chg, options: { fontSize: 11, fontFace: "Arial", align: "center", color: isNeg ? C.red : C.green, bold: true } },
    { text: note, options: { fontSize: 10, fontFace: "Arial", color: C.medGray } },
  ];
  s2.addTable([
    headerRow,
    makeRow("SOL", "$87.49", "-0.5%", "SOL ETF AUM crosses $1B", true),
    makeRow("HYPE", "$38.00", "+1.5%", "Grayscale files HYPE ETF", false),
    makeRow("PENDLE", "$1.19", "-6.4%", "TVL hits $4B milestone", true),
    makeRow("ENA", "$0.10", "-9.7%", "sUSDe dynamic unstaking", true),
    makeRow("DYDX", "$0.08", "-10.0%", "Shuts down 16 markets", true),
  ], {
    x: 0.6, y: 2.25, w: 8.8, colW: [1.3, 1.8, 1.5, 4.2],
    border: { pt: 0.5, color: C.lightGray }, rowH: [0.35, 0.33, 0.33, 0.33, 0.33, 0.33], autoPage: false,
  });

  s2.addText("Macro: Fed holds rates (3.50-3.75%), Middle East crisis pushes oil to $110 triggering Saturday flash crash; BTC from $76K to $68.8K, $1B+ liquidated in 24h", {
    x: 0.6, y: 4.55, w: 8.8, h: 0.4, fontSize: 10, fontFace: "Arial", color: C.medGray, italic: true, margin: 0
  });
  footer(s2, pres);
  s2.addNotes("Market Overview Deep Dive:\n\n1. BTC experienced a classic \"wash\" pattern this week. Mid-week rally to $76K (6-week high) driven by regulatory tailwinds, but institutional buyers were absorbing at 2.8x new supply rate while ETFs saw $219M net outflow on Wednesday - showing institutional divergence.\n\n2. ETH was relatively resilient (-0.6%) supported by Glamsterdam upgrade expectations and SEC explicitly classifying ETH as a \"digital commodity\" rather than a security.\n\n3. Fear & Greed at 10 (Extreme Fear) is the lowest since 2024. Historically, extreme fear zones have been medium-to-long-term buying opportunities.\n\n4. $438M in token unlocks this week added sell pressure: ASTER ($55.9M), LayerZero ($50.3M), Lombard ($34.8M).\n\n5. Positive signals: Metaplanet raised $255M to expand BTC treasury, Bullish reported 24,400 BTC holdings (~$1.67B), corporate buying pace at 3x new supply.\n\nPortfolio implications: PENDLE TVL milestone ($4B) is bullish for fundamentals but price declined 6.4% showing negative short-term sentiment. ENA yield compression to 3.5% APY is a key risk signal - negative funding rates directly compress USDe model's revenue.");

  // ===== SLIDE 3: Top Stories + Industry =====
  const s3 = pres.addSlide();
  s3.background = { color: C.offWhite };
  s3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy } });
  s3.addText("\uD83D\uDD25  Top Stories & Industry", { x: 0.6, y: 0, w: 9, h: 0.7, fontSize: 20, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  s3.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 4.1, fill: { color: C.white }, shadow: makeShadow() });
  s3.addText("TOP 5 STORIES", { x: 0.6, y: 0.95, w: 4, h: 0.35, fontSize: 13, fontFace: "Arial", bold: true, color: C.navy, margin: 0 });

  const stories = [
    ["1", "Strait of Hormuz Flash Crash & $1B Liq.", "BTC crashed from $76K to $68.2K; oil at $110 triggered panic"],
    ["2", "SEC/CFTC Joint Regulatory Guidance", "18 tokens classified as digital commodities; ends reg ambiguity"],
    ["3", "Bhutan Sovereign Fund Moves $72M BTC", "Arkham on-chain data reveals sovereign-level position change"],
    ["4", "20 Millionth Bitcoin Mined", "Only 1M BTC left to mine; BTC dominance steady at 57%"],
    ["5", "Record Token Unlocks in March 2026", "$6B+ monthly; LayerZero $55.5M, Lombard $32.4M this week"],
  ];
  stories.forEach((s, i) => {
    const itemY = 1.38 + i * 0.72;
    s3.addShape(pres.shapes.OVAL, { x: 0.65, y: itemY, w: 0.3, h: 0.3, fill: { color: C.gold } });
    s3.addText(s[0], { x: 0.65, y: itemY, w: 0.3, h: 0.3, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, align: "center", valign: "middle", margin: 0 });
    s3.addText(s[1], { x: 1.1, y: itemY - 0.02, w: 3.6, h: 0.25, fontSize: 11, fontFace: "Arial", bold: true, color: C.bodyText, margin: 0 });
    s3.addText(s[2], { x: 1.1, y: itemY + 0.22, w: 3.6, h: 0.25, fontSize: 9, fontFace: "Arial", color: C.medGray, margin: 0 });
  });

  s3.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 4.5, h: 1.95, fill: { color: C.white }, shadow: makeShadow() });
  s3.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s3.addText("Layer 1 & Layer 2", { x: 5.3, y: 0.9, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });
  s3.addText([
    "Aster Chain privacy L1 mainnet genesis launch",
    "Linea upgrades to Type-1 zkEVM (100% EVM compat.)",
    "Base dominates L2: 46% TVL, >60% fee revenue",
    "Solana Alpenglow in final testing; RWA crosses $1.7B",
  ].map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < 3, fontSize: 10, fontFace: "Arial", color: C.bodyText } })), { x: 5.3, y: 1.32, w: 4.1, h: 1.45, margin: 0, paraSpaceAfter: 4 });

  s3.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 3.05, w: 4.5, h: 1.95, fill: { color: C.white }, shadow: makeShadow() });
  s3.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 3.05, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s3.addText("DeFi & Perp DEX", { x: 5.3, y: 3.05, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });
  s3.addText([
    "Pendle TVL $4B, cumulative volume $10B milestone",
    "DEX perp market share hits 20% globally",
    "Hyperliquid oil $1.77B/day; S&P 500 officially licensed",
    "Aave v4 + Lido v3 major upgrades go live",
  ].map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < 3, fontSize: 10, fontFace: "Arial", color: C.bodyText } })), { x: 5.3, y: 3.47, w: 4.1, h: 1.45, margin: 0, paraSpaceAfter: 4 });
  footer(s3, pres);
  s3.addNotes("Top Stories & Industry Deep Dive:\n\n1. The Strait of Hormuz flash crash is a textbook geopolitical black swan. Trump's ultimatum triggered oil to $110/barrel, causing global risk-off. BTC dropped from $76K to $68.2K with $1B+ liquidated in 24h (85% longs). This shows crypto's sensitivity to geopolitical risk remains high, but historically extreme fear levels have been buying opportunities.\n\n2. The SEC/CFTC joint guidance is arguably the most important US crypto regulatory development ever. The 68-page document classifies BTC, ETH, SOL, XRP and 14 other tokens as \"digital commodities,\" and confirms that mining, staking, airdrops, and wrapped tokens are NOT securities transactions. This fundamentally ends the era of \"regulation by enforcement.\"\n\n3. On L1/L2: Base's dominance (46% L2 TVL, 60%+ fee revenue) shows consumer app aggregation effects accelerating. Solana's Alpenglow upgrade targeting 100-150ms finality would further cement its competitive advantage in high-frequency trading.\n\n4. The S&P 500 officially licensing to Hyperliquid via Trade.xyz is a landmark - first time a major traditional equity index is sanctioned for DEX perpetual trading. DEX reaching 20% global perp market share represents an accelerating structural shift.");

  // ===== SLIDE 4: Regulation + Fundraising =====
  const s4 = pres.addSlide();
  s4.background = { color: C.offWhite };
  s4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy } });
  s4.addText("\u2696\uFE0F  Regulation & Fundraising", { x: 0.6, y: 0, w: 9, h: 0.7, fontSize: 20, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  s4.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 4.1, fill: { color: C.white }, shadow: makeShadow() });
  s4.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 0.06, h: 4.1, fill: { color: C.gold } });
  s4.addText("Regulation & Policy", { x: 0.7, y: 0.95, w: 4, h: 0.35, fontSize: 14, fontFace: "Arial", bold: true, color: C.navy, margin: 0 });

  const regItems = [
    { title: "SEC/CFTC Joint Guidance", desc: "68-page taxonomy: 18 tokens as digital commodities" },
    { title: "SEC Startup Safe Harbor", desc: "Grace period for crypto projects to raise and build" },
    { title: "CFTC Self-Custody No-Action", desc: "Non-custodial wallets exempt from intermediary rules" },
    { title: "Arizona Charges Kalshi", desc: "20 criminal counts; first state-level criminal enforcement" },
    { title: "Nevada Halts Kalshi Ops", desc: "14-day restraining order; federal/state jurisdictional war" },
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
  s4.addTable([
    fundHeader,
    fundRow("Kalshi", "$1B", "Coatue-led; $22B valuation"),
    fundRow("Cryptio", "$45M", "Series B; enterprise accounting"),
    fundRow("MetaComp", "$35M", "Alibaba; Asia-MENA payments"),
    fundRow("GSR x Architech", "M&A", "Tokenomics & liquidity advisory"),
    fundRow("Backpack ($BP)", "TGE 3/23", "Staked tokens convertible to equity"),
  ], {
    x: 5.3, y: 1.45, w: 4.1, colW: [1.3, 0.8, 2.0],
    border: { pt: 0.5, color: C.lightGray }, rowH: [0.3, 0.3, 0.3, 0.3, 0.3, 0.3], autoPage: false,
  });

  s4.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 3.4, w: 4.1, h: 0.7, fill: { color: "FFF8E7" }, shadow: makeShadow() });
  s4.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 3.4, w: 0.06, h: 0.7, fill: { color: C.gold } });
  s4.addText([
    { text: "Watch: ", options: { bold: true, fontSize: 10, fontFace: "Arial", color: C.navy } },
    { text: "Kalshi's $1B raise signals massive institutional conviction in prediction markets, but multi-state regulatory opposition highlights compliance risk", options: { fontSize: 10, fontFace: "Arial", color: C.bodyText } }
  ], { x: 5.55, y: 3.45, w: 3.7, h: 0.6, margin: 0, valign: "middle" });
  footer(s4, pres);
  s4.addNotes("Regulation & Fundraising Deep Dive:\n\n1. The SEC/CFTC joint guidance is a watershed moment for US crypto regulation. The 68-page document establishes a five-category digital asset taxonomy, explicitly classifying BTC, ETH, SOL, XRP and 14 others as \"digital commodities.\" Key implications:\n   - Ends \"regulation by enforcement\" era, clearing the biggest legal barrier for institutional entry\n   - Staking, mining, airdrops, wrapped tokens confirmed NOT securities transactions\n   - Paves the road for ETFs, institutional custody products\n\n2. The CFTC's self-custody no-action letter is equally important - it establishes the legal right to self-custody and ensures wallet developers won't be regulated as financial intermediaries.\n\n3. However, the regulatory environment isn't uniformly positive. Prediction markets face state-level opposition: Arizona criminal charges against Kalshi, Nevada restraining order, Argentina blocking Polymarket. This shows federal vs. state jurisdictional conflict is intensifying.\n\n4. On fundraising, Kalshi's $1B raise at $22B valuation is the biggest highlight, signaling massive institutional conviction in prediction markets. Backpack's TGE model innovation (staked tokens convertible to equity) is noteworthy - a new experiment in token-equity fusion.\n\nPortfolio implications: Backpack ($BP) TGE on 3/23 is directly relevant as a holding. The token-to-equity conversion model could set industry precedent.");

  // ===== SLIDE 5: Security + Sectors =====
  const s5 = pres.addSlide();
  s5.background = { color: C.offWhite };
  s5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy } });
  s5.addText("\uD83D\uDD13  Security & Hot Sectors", { x: 0.6, y: 0, w: 9, h: 0.7, fontSize: 20, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 2.35, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 0.35, fill: { color: C.red } });
  s5.addText("Security Incidents  |  Weekly Loss: ~$31.7M", { x: 0.6, y: 0.9, w: 4.2, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  s5.addTable([
    [
      { text: "Incident", options: { bold: true, fontSize: 9, fontFace: "Arial", fill: { color: C.lightGray }, color: C.bodyText } },
      { text: "Loss", options: { bold: true, fontSize: 9, fontFace: "Arial", fill: { color: C.lightGray }, color: C.bodyText, align: "right" } },
      { text: "Method", options: { bold: true, fontSize: 9, fontFace: "Arial", fill: { color: C.lightGray }, color: C.bodyText } },
    ],
    [{ text: "Resolv Labs", options: { fontSize: 9, fontFace: "Arial" } }, { text: "$25M", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "USR minting exploit", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
    [{ text: "Venus Protocol", options: { fontSize: 9, fontFace: "Arial" } }, { text: "$3.7M", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "Price manipulation", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
    [{ text: "Solv Protocol", options: { fontSize: 9, fontFace: "Arial" } }, { text: "$2.7M", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "Targeted breach", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
    [{ text: "Bitrefill (Lazarus)", options: { fontSize: 9, fontFace: "Arial" } }, { text: "N/A", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "NK state-sponsored APT", options: { fontSize: 9, fontFace: "Arial", color: C.red, bold: true } }],
  ], {
    x: 0.55, y: 1.35, w: 4.2, colW: [1.5, 0.9, 1.8],
    border: { pt: 0.5, color: C.lightGray }, rowH: [0.28, 0.28, 0.28, 0.28, 0.28], autoPage: false,
  });

  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.5, w: 4.5, h: 1.5, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.5, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s5.addText("Stablecoins", { x: 0.6, y: 3.5, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });
  s5.addText([
    "USDC surpasses USDT in YTD transfer volume ($2.2T vs $1.3T)",
    "Stablecoin market cap hits $316B all-time high",
    "Ethena sUSDe yield compresses to 3.49-3.72% APY",
  ].map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < 2, fontSize: 10, fontFace: "Arial", color: C.bodyText } })), { x: 0.6, y: 3.95, w: 4.1, h: 0.9, margin: 0, paraSpaceAfter: 4 });

  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 4.5, h: 1.85, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s5.addText("AI \u00D7 Crypto", { x: 5.3, y: 0.9, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });
  s5.addText([
    "Bittensor completes 72B-param decentralized LLM; TAO +56%",
    "Grayscale files TAO spot ETP (NYSE Arca)",
    "Akash BME hard fork: deflationary AKT + WASM; +20%",
    "MoonPay+Ledger launch first hardware-signed AI agent",
  ].map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < 3, fontSize: 10, fontFace: "Arial", color: C.bodyText } })), { x: 5.3, y: 1.35, w: 4.1, h: 1.3, margin: 0, paraSpaceAfter: 4 });

  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 2.95, w: 4.5, h: 2.05, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 2.95, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s5.addText("Prediction Markets", { x: 5.3, y: 2.95, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });
  s5.addText([
    "Polymarket signs $150M exclusive MLB partnership",
    "Kalshi raises $1B at $22B valuation",
    "Arizona criminal charges + Nevada restraining order",
    "Argentina blocks Polymarket nationwide",
  ].map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < 3, fontSize: 10, fontFace: "Arial", color: C.bodyText } })), { x: 5.3, y: 3.4, w: 4.1, h: 1.5, margin: 0, paraSpaceAfter: 4 });
  footer(s5, pres);
  s5.addNotes("Security & Hot Sectors Deep Dive:\n\n1. Security: Total weekly losses ~$31.7M, dominated by Resolv Labs $25M exploit. The USR stablecoin minting vulnerability allowed an attacker to convert 200K USDC into $25M USR, causing an 80% depeg. This again warns of smart contract risks in new stablecoin protocols. The Bitrefill incident confirmed as North Korea's Lazarus Group shows nation-state threats to crypto persist. The new \"Coruna\" iOS exploit targeting iPhone wallet seed phrases is concerning for mobile-first users.\n\n2. AI x Crypto: Bittensor's 72B parameter decentralized LLM training is a historic milestone - 70 global contributors using commodity hardware proved decentralized compute viability. Grayscale's TAO ETP filing brings institutional legitimacy. Akash's BME hard fork makes AKT deflationary tied to network usage. MoonPay+Ledger's hardware-signed AI agent solves the critical security bottleneck for autonomous fund management.\n\n3. Prediction Markets: Polymarket's $150M MLB exclusive deal is historic - first deep integration of prediction markets with mainstream sports. But multi-state regulatory opposition (Arizona criminal charges, Nevada restraining order, Argentina blocking) shows the compliance risk. The federal vs. state jurisdictional war could be a key variable.\n\n4. Stablecoins: USDC surpassing USDT in YTD transfer volume is a critical inflection point, reflecting regulatory compliance importance for institutional users. Ethena yield compression to ~3.5% is worth monitoring - negative funding rates directly compress the USDe model.");

  // ===== SLIDE 6: Catalysts + Outlook =====
  const s6 = pres.addSlide();
  s6.background = { color: C.offWhite };
  s6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy } });
  s6.addText("\uD83D\uDD2E  Upcoming Catalysts & Outlook", { x: 0.6, y: 0, w: 9, h: 0.7, fontSize: 20, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  const catHeader = [
    { text: "Date", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial", align: "center" } },
    { text: "Event", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial" } },
    { text: "Asset", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial", align: "center" } },
    { text: "Expected Impact", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial" } },
  ];
  const catRow = (date, event, asset, impact, hl = false) => [
    { text: date, options: { fontSize: 10, fontFace: "Arial", align: "center", color: C.bodyText, fill: hl ? { color: "FFF3E0" } : undefined } },
    { text: event, options: { fontSize: 10, fontFace: "Arial", color: C.bodyText, bold: hl, fill: hl ? { color: "FFF3E0" } : undefined } },
    { text: asset, options: { fontSize: 10, fontFace: "Arial", align: "center", color: C.navy, bold: true, fill: hl ? { color: "FFF3E0" } : undefined } },
    { text: impact, options: { fontSize: 9, fontFace: "Arial", color: C.medGray, fill: hl ? { color: "FFF3E0" } : undefined } },
  ];
  s6.addTable([
    catHeader,
    catRow("3/23", "Backpack ($BP) TGE", "BP", "Staked tokens convertible to equity", true),
    catRow("3/23", "Akash BME Hard Fork", "AKT", "Deflationary tokenomics + WASM"),
    catRow("3/24-26", "Digital Asset Summit (NYC)", "RWA", "Institutional adoption signal"),
    catRow("3/30-4/2", "EthCC 9 (Cannes)", "ETH", "Largest European ETH conference"),
    catRow("4/3", "CLARITY Act Signing?", "All", "Legal commodity/security definition milestone", true),
    catRow("Apr", "Uniswap Fee Switch", "UNI", "Fee revenue on 8 L2 chains"),
    catRow("Early Apr", "Polkadot 2.1.1 Upgrade", "DOT", "Dynamic Allocation Pool; end Treasury burns"),
  ], {
    x: 0.5, y: 0.9, w: 9, colW: [0.9, 3.2, 1.2, 3.7],
    border: { pt: 0.5, color: C.lightGray }, rowH: [0.32, 0.32, 0.32, 0.32, 0.32, 0.32, 0.32, 0.32], autoPage: false,
  });

  s6.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.65, w: 9, h: 1.3, fill: { color: C.white }, shadow: makeShadow() });
  s6.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.65, w: 0.06, h: 1.3, fill: { color: C.gold } });
  s6.addText("Weekly Summary", { x: 0.8, y: 3.7, w: 8, h: 0.3, fontSize: 13, fontFace: "Arial", bold: true, color: C.navy, margin: 0 });
  s6.addText([
    { text: "Market sentiment crashed to Extreme Fear (FGI=10) as BTC flash-crashed from $76K to $68.8K. However, landmark SEC/CFTC regulatory guidance provides an extremely positive medium-term foundation.", options: { breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "Portfolio Watch: ", options: { bold: true, fontSize: 10, fontFace: "Arial", color: C.navy } },
    { text: "Backpack TGE (3/23), HYPE ETF filing progress, Pendle $4B TVL milestone, ENA yield compression risk, Middle East geopolitical developments.", options: { fontSize: 10, fontFace: "Arial", color: C.bodyText } },
  ], { x: 0.8, y: 4.0, w: 8.5, h: 0.85, margin: 0 });

  s6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s6.addText(`Crypto Weekly Digest | ${WEEK}  |  CONFIDENTIAL`, { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });
  s6.addNotes("Catalysts & Outlook Deep Dive:\n\n1. Backpack TGE (3/23) is the most immediately relevant event as a portfolio holding. The \"staked tokens convertible to equity\" model is an innovative experiment in token-equity fusion that could set industry precedent. Monitor closely for launch dynamics.\n\n2. CLARITY Act potential signing on 4/3 would be one of the most important US crypto legislative milestones ever, providing definitive legal definitions for digital commodity vs. security classification. This would directly impact the entire market.\n\n3. EthCC 9 (Cannes) historically features major technical and ecosystem announcements. Combined with Glamsterdam upgrade expectations, could provide positive catalysts for ETH.\n\n4. Portfolio Management Perspective:\n   - PENDLE: $4B TVL is a critical milestone for fundamentals, but price declined 6.4% showing negative short-term sentiment. The divergence between fundamentals and price may present opportunity.\n   - ENA: sUSDe yield compression to 3.5% is a key risk signal. Negative funding rates directly compress USDe model revenue. Need to monitor funding rate environment.\n   - HYPE: Grayscale ETF filing is a major medium-term positive, but approval timeline uncertain. The S&P 500 licensing deal validates the platform.\n   - The Middle East geopolitical crisis is the biggest uncertainty factor. If tensions escalate further, expect continued risk-off across all assets. If tensions de-escalate, the regulatory tailwinds + extreme fear positioning could catalyze a strong recovery.");

  const outPath = "/Users/lx/Github/crypto-weekly-digest/reports/2026-03-16/slides-en.pptx";
  await pres.writeFile({ fileName: outPath });
  console.log("\u2705 Written to " + outPath);
}

generate().catch(e => { console.error(e); process.exit(1); });
