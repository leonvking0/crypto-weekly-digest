const pptxgen = require("pptxgenjs");

const C = {
  navy: "1A2744",
  darkNavy: "0F1A2E",
  gold: "C9A96E",
  white: "FFFFFF",
  offWhite: "F5F6F8",
  lightGray: "E8EBF0",
  medGray: "8C95A6",
  bodyText: "2D3748",
  red: "C53030",
  green: "2F855A",
};

const makeShadow = () => ({ type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.12 });

async function generate() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "42";
  pres.title = "Crypto Market Weekly Report — Feb 20-27, 2026";

  // ============ SLIDE 1: Title ============
  const s1 = pres.addSlide();
  s1.background = { color: C.darkNavy };
  s1.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.6, w: 1.2, h: 0.04, fill: { color: C.gold } });
  s1.addText("Crypto Market Weekly", {
    x: 0.8, y: 1.8, w: 8, h: 1.0,
    fontSize: 42, fontFace: "Arial", bold: true, color: C.white, margin: 0
  });
  s1.addText("February 20 \u2013 27, 2026", {
    x: 0.8, y: 2.7, w: 8, h: 0.5,
    fontSize: 20, fontFace: "Arial", color: C.gold, margin: 0
  });
  s1.addText("Sources: Messari \u00B7 The Block \u00B7 CoinGecko \u00B7 CoinGlass \u00B7 DefiLlama", {
    x: 0.8, y: 3.4, w: 8, h: 0.4,
    fontSize: 11, fontFace: "Arial", color: C.medGray, margin: 0
  });
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.gold } });
  s1.addText("CONFIDENTIAL \u2014 FOR LP DISTRIBUTION ONLY", {
    x: 0, y: 5.25, w: 10, h: 0.375,
    fontSize: 9, fontFace: "Arial", color: C.darkNavy, align: "center", valign: "middle", bold: true
  });
  s1.addNotes("Welcome to the weekly crypto market report covering February 20-27, 2026. This report is prepared for LP distribution and covers market performance, key developments, regulatory updates, security incidents, and upcoming catalysts. All price data sourced from CoinGecko API; news from Messari, The Block, CoinDesk, and other major outlets.");

  // ============ SLIDE 2: Market Overview ============
  const s2 = pres.addSlide();
  s2.background = { color: C.offWhite };
  s2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy } });
  s2.addText("\uD83D\uDCCA  Market Overview", {
    x: 0.6, y: 0, w: 9, h: 0.7,
    fontSize: 20, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0
  });

  const cardY = 0.95, cardH = 1.05, cardW = 2.8, cardGap = 0.3, cardStartX = 0.6;

  // Card 1: BTC
  s2.addShape(pres.shapes.RECTANGLE, { x: cardStartX, y: cardY, w: cardW, h: cardH, fill: { color: C.white }, shadow: makeShadow() });
  s2.addShape(pres.shapes.RECTANGLE, { x: cardStartX, y: cardY, w: 0.06, h: cardH, fill: { color: C.gold } });
  s2.addText("BTC", { x: cardStartX + 0.25, y: cardY + 0.1, w: 2, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.medGray, bold: true, margin: 0 });
  s2.addText("$67,119", { x: cardStartX + 0.25, y: cardY + 0.35, w: 2, h: 0.4, fontSize: 28, fontFace: "Arial", color: C.bodyText, bold: true, margin: 0 });
  s2.addText("+0.3%", { x: cardStartX + 0.25, y: cardY + 0.72, w: 2, h: 0.25, fontSize: 14, fontFace: "Arial", color: C.green, bold: true, margin: 0 });

  // Card 2: ETH
  const c2x = cardStartX + cardW + cardGap;
  s2.addShape(pres.shapes.RECTANGLE, { x: c2x, y: cardY, w: cardW, h: cardH, fill: { color: C.white }, shadow: makeShadow() });
  s2.addShape(pres.shapes.RECTANGLE, { x: c2x, y: cardY, w: 0.06, h: cardH, fill: { color: C.gold } });
  s2.addText("ETH", { x: c2x + 0.25, y: cardY + 0.1, w: 2, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.medGray, bold: true, margin: 0 });
  s2.addText("$1,939", { x: c2x + 0.25, y: cardY + 0.35, w: 2, h: 0.4, fontSize: 28, fontFace: "Arial", color: C.bodyText, bold: true, margin: 0 });
  s2.addText("-2.6%", { x: c2x + 0.25, y: cardY + 0.72, w: 2, h: 0.25, fontSize: 14, fontFace: "Arial", color: C.red, bold: true, margin: 0 });

  // Card 3: Fear & Greed
  const c3x = c2x + cardW + cardGap;
  s2.addShape(pres.shapes.RECTANGLE, { x: c3x, y: cardY, w: cardW, h: cardH, fill: { color: C.white }, shadow: makeShadow() });
  s2.addShape(pres.shapes.RECTANGLE, { x: c3x, y: cardY, w: 0.06, h: cardH, fill: { color: C.red } });
  s2.addText("Fear & Greed Index", { x: c3x + 0.25, y: cardY + 0.1, w: 2.4, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.medGray, bold: true, margin: 0 });
  s2.addText("13", { x: c3x + 0.25, y: cardY + 0.35, w: 1, h: 0.4, fontSize: 28, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
  s2.addText("Extreme Fear", { x: c3x + 1.1, y: cardY + 0.42, w: 1.5, h: 0.3, fontSize: 13, fontFace: "Arial", color: C.red, margin: 0 });

  // Price table
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
  s2.addTable([
    headerRow,
    makeRow("SOL", "$81.73", "-3.0%", "MCap ~$40B", true),
    makeRow("HYPE", "$30.25", "-3.7%", "3/6 contributor unlock $316M", true),
    makeRow("PENDLE", "$1.18", "-8.3%", "", true),
    makeRow("ENA", "$0.099", "-5.4%", "", true),
    makeRow("DYDX", "$0.080", "-12.4%", "Worst performer", true),
  ], {
    x: 0.6, y: 2.25, w: 8.8, colW: [1.3, 1.8, 1.5, 4.2],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.35, 0.33, 0.33, 0.33, 0.33, 0.33],
    autoPage: false,
  });

  s2.addText("Macro driver: US Supreme Court overturned executive tariff order on 2/20 \u2192 brief risk-on; overall sentiment remains in Extreme Fear.", {
    x: 0.6, y: 4.55, w: 8.8, h: 0.4,
    fontSize: 10, fontFace: "Arial", color: C.medGray, italic: true, margin: 0
  });

  s2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s2.addText("Crypto Market Weekly | 2026.02.20\u201302.27", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  s2.addNotes("BTC traded in a $62,400-$70,000 range this week, ending roughly flat at $67,119 (+0.3%). ETH underperformed at -2.6%. Fear & Greed Index at 13 signals extreme fear \u2014 the lowest since the FTX collapse era. Key macro event: US Supreme Court overturned executive tariff order on 2/20, providing a brief risk-on bounce. Alt performance was broadly negative with DYDX leading losses at -12.4%. Notable: HYPE has a major $316M contributor unlock on March 6 \u2014 a key risk for our position.");

  // ============ SLIDE 3: Top Stories + DeFi/L1 ============
  const s3 = pres.addSlide();
  s3.background = { color: C.offWhite };
  s3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy } });
  s3.addText("\uD83D\uDD25  Top Stories & Industry Developments", {
    x: 0.6, y: 0, w: 9, h: 0.7,
    fontSize: 20, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0
  });

  // Left: Top 5
  s3.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 4.1, fill: { color: C.white }, shadow: makeShadow() });
  s3.addText("TOP 5 STORIES", {
    x: 0.6, y: 0.95, w: 4, h: 0.35,
    fontSize: 13, fontFace: "Arial", bold: true, color: C.navy, margin: 0
  });

  const stories = [
    ["1", "Block Inc. cuts 40%, pivots to AI ops", "Dorsey replaces traditional structure with AI Agent workflows"],
    ["2", "Uniswap Fee Switch vote launched", "8 L2 chains activate protocol fees; est. $27M annual revenue"],
    ["3", "SEC token classification + OCC rules", "Securities vs commodities clarity; stablecoin $5M min capital"],
    ["4", "Aave cumulative lending hits $1 trillion", "RWA assets also surpass $1B; DeFi credit leader"],
    ["5", "Grayscale files TAO spot ETF", "First institutional vehicle for decentralized AI compute"],
  ];

  stories.forEach((s, i) => {
    const itemY = 1.38 + i * 0.72;
    s3.addShape(pres.shapes.OVAL, { x: 0.65, y: itemY, w: 0.3, h: 0.3, fill: { color: C.gold } });
    s3.addText(s[0], { x: 0.65, y: itemY, w: 0.3, h: 0.3, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, align: "center", valign: "middle", margin: 0 });
    s3.addText(s[1], { x: 1.1, y: itemY - 0.02, w: 3.6, h: 0.25, fontSize: 11, fontFace: "Arial", bold: true, color: C.bodyText, margin: 0 });
    s3.addText(s[2], { x: 1.1, y: itemY + 0.22, w: 3.6, h: 0.25, fontSize: 9, fontFace: "Arial", color: C.medGray, margin: 0 });
  });

  // Right top: L1/L2
  s3.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 4.5, h: 1.95, fill: { color: C.white }, shadow: makeShadow() });
  s3.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s3.addText("Layer 1 & Layer 2", { x: 5.3, y: 0.9, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  const l1Items = [
    "MegaETH mainnet: 2.1M txns day 1, sub-10ms blocks",
    "Base leaves OP Stack for custom architecture",
    "On-chain RWA market cap breaks $15B (YoY +200%)",
    "US spot BTC ETFs: $1.1B net inflows in 3 days",
  ];
  s3.addText(l1Items.map((t, i) => ({
    text: t,
    options: { bullet: true, breakLine: i < l1Items.length - 1, fontSize: 10, fontFace: "Arial", color: C.bodyText }
  })), { x: 5.3, y: 1.32, w: 4.1, h: 1.45, margin: 0, paraSpaceAfter: 4 });

  // Right bottom: DeFi
  s3.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 3.05, w: 4.5, h: 1.95, fill: { color: C.white }, shadow: makeShadow() });
  s3.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 3.05, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s3.addText("DeFi & Perp DEXes", { x: 5.3, y: 3.05, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  const defiItems = [
    "Aave: $1T cumulative lending + $1B RWA milestone",
    "DeFi TVL at $96.8B (-0.52%); Solana bucked trend +3.46%",
    "Hyperliquid weekly volume $35.4B, daily peak $6B",
    "Perp DEX market share reaches 10.2% (was 2% in Jan 2024)",
  ];
  s3.addText(defiItems.map((t, i) => ({
    text: t,
    options: { bullet: true, breakLine: i < defiItems.length - 1, fontSize: 10, fontFace: "Arial", color: C.bodyText }
  })), { x: 5.3, y: 3.47, w: 4.1, h: 1.45, margin: 0, paraSpaceAfter: 4 });

  s3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s3.addText("Crypto Market Weekly | 2026.02.20\u201302.27", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  s3.addNotes("Top story: Block Inc. (Jack Dorsey) cutting 40% of workforce (~4,000 people) to pivot entirely to AI Agent workflows \u2014 a dramatic signal of AI replacing corporate functions at scale.\n\nUniswap's Fee Switch is a landmark moment for DeFi sustainability \u2014 first major protocol to activate revenue sharing with token holders across 8 L2s.\n\nSEC token classification guidance submitted to White House is the most concrete step toward regulatory clarity we've seen. Combined with OCC's GENIUS Act framework for stablecoins ($5M minimum capital, no interest payments allowed), this week saw more regulatory progress than all of Q4 2025.\n\nAave hitting $1T cumulative lending is a DeFi milestone. Hyperliquid continues to dominate perp DEXes with $35.4B weekly volume. The structural shift of perp DEX market share from 2% to 10.2% in two years is a key thesis confirmation for our portfolio.");

  // ============ SLIDE 4: Regulation + Fundraising ============
  const s4 = pres.addSlide();
  s4.background = { color: C.offWhite };
  s4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy } });
  s4.addText("\u2696\uFE0F  Regulation & Fundraising", {
    x: 0.6, y: 0, w: 9, h: 0.7,
    fontSize: 20, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0
  });

  // Left: Regulation
  s4.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 4.1, fill: { color: C.white }, shadow: makeShadow() });
  s4.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 0.06, h: 4.1, fill: { color: C.gold } });
  s4.addText("Regulation & Policy", { x: 0.7, y: 0.95, w: 4, h: 0.35, fontSize: 14, fontFace: "Arial", bold: true, color: C.navy, margin: 0 });

  const regItems = [
    { title: "SEC Token Classification", desc: "Guidance submitted to White House; securities vs commodities" },
    { title: "OCC GENIUS Act Framework", desc: "Stablecoin min $5M capital; no interest payments allowed" },
    { title: "CFTC Innovation Committee", desc: "35-member advisory; major Crypto/DeFi leaders included" },
    { title: "CLARITY Act Deadlock", desc: "Stablecoin interest payments: banking vs crypto divide" },
    { title: "UK FCA Stablecoin Sandbox", desc: "Revolut + 3 others selected; formal rules by end of 2026" },
  ];

  regItems.forEach((item, i) => {
    const iy = 1.45 + i * 0.65;
    s4.addText(item.title, { x: 0.7, y: iy, w: 4, h: 0.25, fontSize: 11, fontFace: "Arial", bold: true, color: C.bodyText, margin: 0 });
    s4.addText(item.desc, { x: 0.7, y: iy + 0.23, w: 4, h: 0.25, fontSize: 10, fontFace: "Arial", color: C.medGray, margin: 0 });
  });

  // Right: Fundraising
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
    fundRow("Flying Tulip", "$206M", "Token Sale"),
    fundRow("Whop", "$200M", "Tether; val $1.6B"),
    fundRow("Gold.com", "$150M", "Tether; digital gold"),
    fundRow("BTC Inc.", "$107M", "Acquired by Nakamoto"),
    fundRow("Anchorage", "$100M", "Tether; val $4.2B"),
  ], {
    x: 5.3, y: 1.45, w: 4.1, colW: [1.1, 0.9, 2.1],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.3, 0.3, 0.3, 0.3, 0.3, 0.3],
    autoPage: false,
  });

  // Tether callout
  s4.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 3.4, w: 4.1, h: 0.7, fill: { color: "FFF8E7" }, shadow: makeShadow() });
  s4.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 3.4, w: 0.06, h: 0.7, fill: { color: C.gold } });
  s4.addText([
    { text: "Key: ", options: { bold: true, fontSize: 10, fontFace: "Arial", color: C.navy } },
    { text: "Tether deployed $450M+ this week, expanding from stablecoins into e-commerce, gold, and custody banking", options: { fontSize: 10, fontFace: "Arial", color: C.bodyText } }
  ], { x: 5.55, y: 3.45, w: 3.7, h: 0.6, margin: 0, valign: "middle" });

  s4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s4.addText("Crypto Market Weekly | 2026.02.20\u201302.27", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  s4.addNotes("Regulation: This was arguably the most significant regulatory week in recent memory. The SEC submitting token classification guidance to the White House signals we're moving from enforcement-by-litigation to actual rulemaking. The OCC GENIUS Act framework sets concrete requirements for stablecoin issuers \u2014 $5M minimum capital, no crypto-asset reserves, and crucially, no interest payments allowed. The no-interest provision is a key battleground between traditional banking and crypto lobbies.\n\nFundraising: Tether is the story here \u2014 $450M+ deployed across Whop (social commerce), Gold.com (digital gold), and Anchorage Digital (federal-chartered crypto bank at $4.2B valuation). This is a clear strategic diversification beyond stablecoins into the broader financial infrastructure stack. Flying Tulip's $206M token sale was the largest single raise of the week.");

  // ============ SLIDE 5: Security + AI + Prediction Markets ============
  const s5 = pres.addSlide();
  s5.background = { color: C.offWhite };
  s5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy } });
  s5.addText("\uD83D\uDD13  Security & Sector Highlights", {
    x: 0.6, y: 0, w: 9, h: 0.7,
    fontSize: 20, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0
  });

  // Security table - left
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 2.35, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 0.35, fill: { color: C.red } });
  s5.addText("Security Incidents  |  Weekly Loss: ~$21M\u2013$26M", { x: 0.6, y: 0.9, w: 4.2, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  s5.addTable([
    [
      { text: "Incident", options: { bold: true, fontSize: 9, fontFace: "Arial", fill: { color: C.lightGray }, color: C.bodyText } },
      { text: "Loss", options: { bold: true, fontSize: 9, fontFace: "Arial", fill: { color: C.lightGray }, color: C.bodyText, align: "right" } },
      { text: "Vector", options: { bold: true, fontSize: 9, fontFace: "Arial", fill: { color: C.lightGray }, color: C.bodyText } },
    ],
    [{ text: "YieldBlox", options: { fontSize: 9, fontFace: "Arial" } }, { text: "~$10M", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "Oracle manipulation", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
    [{ text: "IoTeX/ioTube", options: { fontSize: 9, fontFace: "Arial" } }, { text: "$4.4\u20139M", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "Private key leak", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
    [{ text: "Korea NTS", options: { fontSize: 9, fontFace: "Arial" } }, { text: "~$4.8M", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "Mnemonic exposed", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
    [{ text: "Moonwell", options: { fontSize: 9, fontFace: "Arial" } }, { text: "$1.78M", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "AI-generated code flaw", options: { fontSize: 9, fontFace: "Arial", color: C.red, bold: true } }],
  ], {
    x: 0.55, y: 1.35, w: 4.2, colW: [1.4, 1.1, 1.7],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.28, 0.28, 0.28, 0.28, 0.28],
    autoPage: false,
  });

  // AI x Crypto - right top
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 4.5, h: 1.85, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s5.addText("AI \u00D7 Crypto", { x: 5.3, y: 0.9, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  s5.addText([
    { text: "Grayscale files TAO spot ETF \u2014 institutional DeMI", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "Bitdeer liquidates BTC treasury, pivots to AI compute", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "On-chain autonomous AI Agents surpass 11,000 (ERC-8004)", options: { bullet: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
  ], { x: 5.3, y: 1.35, w: 4.1, h: 1.3, margin: 0, paraSpaceAfter: 4 });

  // Prediction Markets - right bottom
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 2.95, w: 4.5, h: 1.85, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 2.95, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s5.addText("Prediction Markets", { x: 5.3, y: 2.95, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  s5.addText([
    { text: "Polymarket daily peak $425M; monthly $7B (YoY 7.5x)", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "Feb total: Kalshi $9.9B > Polymarket $7.9B; industry $23.4B", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "ICE integrates Polymarket signals; CFTC eyes regulation", options: { bullet: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
  ], { x: 5.3, y: 3.4, w: 4.1, h: 1.3, margin: 0, paraSpaceAfter: 4 });

  // Stablecoins - bottom left
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.5, w: 4.5, h: 1.5, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.5, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s5.addText("Stablecoins", { x: 0.6, y: 3.5, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });
  s5.addText([
    { text: "Total MCap $297.75B; USDT minted $1.6B on Tron in Feb", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "Visa \u00D7 Stripe stablecoin cards expand to 100+ countries", options: { bullet: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
  ], { x: 0.6, y: 3.95, w: 4.1, h: 0.9, margin: 0, paraSpaceAfter: 4 });

  s5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s5.addText("Crypto Market Weekly | 2026.02.20\u201302.27", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  s5.addNotes("Security: Total weekly losses of $21-26M \u2014 a relatively mild week with no mega-hacks. The Moonwell incident is notable because it was caused by AI-generated smart contract code \u2014 a new attack vector category that will likely grow as more protocols use AI for development.\n\nAI x Crypto: Grayscale filing a TAO spot ETF is a first for decentralized AI compute networks and could unlock institutional capital into this narrative. Bitdeer's full pivot from BTC mining to AI compute infrastructure is a bellwether for the mining industry.\n\nPrediction Markets: Polymarket's $7B monthly volume (7.5x YoY) shows the sector is exploding. Kalshi has actually overtaken Polymarket in monthly volume ($9.9B vs $7.9B) \u2014 the regulated venue is winning. ICE (Intercontinental Exchange) integrating Polymarket signals is a major institutional validation moment.\n\nStablecoins: $297.75B total market cap is near all-time highs. Visa x Stripe expanding stablecoin cards to 100+ countries is a massive distribution play.");

  // ============ SLIDE 6: Catalysts + Outlook ============
  const s6 = pres.addSlide();
  s6.background = { color: C.offWhite };
  s6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy } });
  s6.addText("\uD83D\uDD2E  Upcoming Catalysts & Outlook", {
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
  s6.addTable([
    catHeader,
    catRow("3/1", "HyperEVM Mainnet Launch", "HYPE", "EVM smart contracts go live"),
    catRow("3/1", "SUI Token Unlock (53.82M tokens)", "SUI", "Supply pressure ~$50M+"),
    catRow("3/4", "Polygon Lisovo Hard Fork", "POL", "Contract optimization + cross-chain"),
    catRow("3/5", "POWER Token Unlock", "POWER", "Supply pressure $23M"),
    catRow("3/6", "HYPE Contributor Unlock $316.6M", "HYPE", "Major liquidity event", true),
    catRow("3/12", "Polkadot Tokenomics Reform", "DOT", "Emissions cut 53.6%"),
    catRow("3/12", "APT Token Unlock (11.31M tokens)", "APT", "Supply pressure ~$100M+"),
  ], {
    x: 0.5, y: 0.9, w: 9, colW: [0.9, 3.2, 1.2, 3.7],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.32, 0.32, 0.32, 0.32, 0.32, 0.32, 0.32, 0.32],
    autoPage: false,
  });

  // Outlook box
  s6.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.65, w: 9, h: 1.3, fill: { color: C.white }, shadow: makeShadow() });
  s6.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.65, w: 0.06, h: 1.3, fill: { color: C.gold } });
  s6.addText("Weekly Summary", { x: 0.8, y: 3.7, w: 8, h: 0.3, fontSize: 13, fontFace: "Arial", bold: true, color: C.navy, margin: 0 });
  s6.addText([
    { text: "Market sentiment at Extreme Fear (FGI=13). BTC range-bound; alts broadly down. Positive regulatory signals (SEC classification, OCC stablecoin framework) have not yet translated into price action.", options: { breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "Key watches: ", options: { bold: true, fontSize: 10, fontFace: "Arial", color: C.navy } },
    { text: "3/6 HYPE $316M unlock (portfolio exposure), Tether\u2019s strategic expansion pace, DeFi protocol revenue sustainability progress.", options: { fontSize: 10, fontFace: "Arial", color: C.bodyText } },
  ], { x: 0.8, y: 4.0, w: 8.5, h: 0.85, margin: 0 });

  s6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s6.addText("Crypto Market Weekly | 2026.02.20\u201302.27  |  CONFIDENTIAL", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  s6.addNotes("Catalyst focus for next two weeks:\n\n1. HYPE Contributor Unlock (3/6) \u2014 $316.6M is the largest single unlock event. This is directly relevant to our portfolio. HyperEVM mainnet on 3/1 could provide a positive narrative offset, but the unlock size is significant.\n\n2. SUI and APT unlocks add aggregate supply pressure of ~$150M+ to the market in a fear-driven environment.\n\n3. Polkadot's 53.6% emissions reduction is structurally bullish for DOT supply dynamics.\n\nOverall assessment: Market is in a capitulation-adjacent zone with FGI at 13. Regulatory progress is real but hasn't catalyzed buying yet. The disconnect between improving fundamentals (Aave $1T, perp DEX growth, stablecoin adoption) and depressed sentiment suggests we're in a late-cycle fear phase. Maintain current positioning; the HYPE unlock on 3/6 is the primary near-term risk to monitor.");

  const outPath = "/Users/lx/Github/crypto-weekly-digest/reports/2026-02-24-en.pptx";
  await pres.writeFile({ fileName: outPath });
  console.log("\u2705 Written to " + outPath);
}

generate().catch(e => { console.error(e); process.exit(1); });
