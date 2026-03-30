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
  pres.title = "Crypto Market Weekly Report \u2014 Mar 23-29, 2026";

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
  s1.addText("Sources: Messari \u00B7 The Block \u00B7 CoinGecko \u00B7 CoinGlass \u00B7 DefiLlama", {
    x: 0.8, y: 3.4, w: 8, h: 0.4,
    fontSize: 11, fontFace: "Arial", color: C.medGray, margin: 0
  });
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.gold } });
  s1.addText("CONFIDENTIAL \u2014 FOR LP DISTRIBUTION ONLY", {
    x: 0, y: 5.25, w: 10, h: 0.375,
    fontSize: 9, fontFace: "Arial", color: C.darkNavy, align: "center", valign: "middle", bold: true
  });
  s1.addNotes("Welcome to the weekly crypto market report covering March 23-29, 2026. This week saw extreme fear (FGI=8), driven by Iran-Israel war escalation, Trump tariff increases to 15%, and the Fed raising its PCE inflation forecast to 2.7%. However, the most significant development was the SEC-CFTC joint rule classifying 16 crypto assets as digital commodities \u2014 the most important regulatory clarity event in a decade. The stablecoin sector was rocked by the CLARITY Act draft banning passive yield, sending Circle stock down 20%.");

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
  s2.addText("$66,668", { x: cardStartX + 0.25, y: cardY + 0.35, w: 2, h: 0.4, fontSize: 28, fontFace: "Arial", color: C.bodyText, bold: true, margin: 0 });
  s2.addText("-1.5%", { x: cardStartX + 0.25, y: cardY + 0.72, w: 2, h: 0.25, fontSize: 14, fontFace: "Arial", color: C.red, bold: true, margin: 0 });

  // Card 2: ETH
  const c2x = cardStartX + cardW + cardGap;
  s2.addShape(pres.shapes.RECTANGLE, { x: c2x, y: cardY, w: cardW, h: cardH, fill: { color: C.white }, shadow: makeShadow() });
  s2.addShape(pres.shapes.RECTANGLE, { x: c2x, y: cardY, w: 0.06, h: cardH, fill: { color: C.gold } });
  s2.addText("ETH", { x: c2x + 0.25, y: cardY + 0.1, w: 2, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.medGray, bold: true, margin: 0 });
  s2.addText("$2,012", { x: c2x + 0.25, y: cardY + 0.35, w: 2, h: 0.4, fontSize: 28, fontFace: "Arial", color: C.bodyText, bold: true, margin: 0 });
  s2.addText("-1.5%", { x: c2x + 0.25, y: cardY + 0.72, w: 2, h: 0.25, fontSize: 14, fontFace: "Arial", color: C.red, bold: true, margin: 0 });

  // Card 3: Fear & Greed
  const c3x = c2x + cardW + cardGap;
  s2.addShape(pres.shapes.RECTANGLE, { x: c3x, y: cardY, w: cardW, h: cardH, fill: { color: C.white }, shadow: makeShadow() });
  s2.addShape(pres.shapes.RECTANGLE, { x: c3x, y: cardY, w: 0.06, h: cardH, fill: { color: C.red } });
  s2.addText("Fear & Greed Index", { x: c3x + 0.25, y: cardY + 0.1, w: 2.4, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.medGray, bold: true, margin: 0 });
  s2.addText("8", { x: c3x + 0.25, y: cardY + 0.35, w: 1, h: 0.4, fontSize: 28, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
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
    makeRow("SOL", "$82.73", "-4.0%", "SEC commodity classification not yet priced in", true),
    makeRow("HYPE", "$38.13", "+0.3%", "Perp DEX leader; weekly fees $14M (+56%)", false),
    makeRow("PENDLE", "$1.14", "-2.9%", "Yield token sector under pressure", true),
    makeRow("ENA", "$0.09", "-5.2%", "Tracking broad market weakness", true),
    makeRow("DYDX", "$0.10", "+21.0%", "Technical bounce + buyback expansion to 75%", false),
  ], {
    x: 0.6, y: 2.25, w: 8.8, colW: [1.3, 1.8, 1.5, 4.2],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.35, 0.33, 0.33, 0.33, 0.33, 0.33],
    autoPage: false,
  });

  s2.addText("Macro: Iran-Israel war escalation + Trump tariffs at 15% + Fed PCE forecast raised to 2.7% \u2192 risk assets broadly pressured; FGI at 8", {
    x: 0.6, y: 4.55, w: 8.8, h: 0.4,
    fontSize: 10, fontFace: "Arial", color: C.medGray, italic: true, margin: 0
  });

  s2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s2.addText("Crypto Market Weekly | 2026.03.23\u201303.29", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  s2.addNotes("BTC dropped from $68,378 at week open to $66,668, hitting its March low. ETH broke below $2,000 for the first time since mid-2024 \u2014 now 60% below its August 2025 high of $4,953. Fear & Greed Index at 8 is the lowest in recent memory.\n\nThree macro headwinds converged: (1) Iran threatened to block the Bab el-Mandeb Strait, sending Brent crude to $120/bbl (+25% since conflict start); (2) Trump raised global tariffs to 15% under Section 122; (3) Fed revised 2026 PCE inflation forecast to 2.7% on March 18, pushing rate cut expectations further out with the 10Y yield near 4.5%.\n\nETF flows turned sharply negative: BTC ETFs saw $296M weekly outflow, reversing a 4-week $2.2B inflow streak. BlackRock IBIT alone shed $201.67M on Friday. ETH ETFs posted outflows every session since March 18. March 26 was the first session where BTC, ETH, and SOL ETFs all posted negative net flows simultaneously.\n\nFor LPs: Short-term sentiment is extremely bearish, but the SEC-CFTC digital commodity classification (effective 3/23) is a historic positive that hasn't been priced in. It opens SOL/XRP spot ETF pathways \u2014 a medium-term catalyst.");

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
    ["1", "SEC-CFTC: 16 cryptos classified as commodities", "Decade of regulatory ambiguity ends; SOL/XRP ETF path cleared"],
    ["2", "Circle stock crashes 20% on CLARITY Act", "Senate draft bans passive stablecoin yield; $2B market cap erased"],
    ["3", "Resolv USR exploit: $25M stolen, 97.5% depeg", "AWS key compromised; 80M unbacked tokens minted"],
    ["4", "Iran-Israel war + Trump 15% tariff", "ETFs all net negative; oil +25%; risk assets crushed"],
    ["5", "BTC miners mass-pivot to AI compute", "Mining cost $80K vs $68K price; $70B AI contracts signed"],
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
    "SEC-CFTC digital commodity rule effective 3/23; 16 tokens classified",
    "Base leads L2: $4.15B TVL, 3.3B+ txns YTD",
    "Linea Yield Boost: bridged ETH auto-staked via Lido V3",
    "L2/L1 DAU ratio fell to 1.12 (peak: 10.43 in Jun 2025)",
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
    "Aave: $1T cumulative lending; TVL $27.3B; 62.8% market share",
    "DeFi TVL +4.4% despite extreme fear sentiment",
    "Hyperliquid: 54% perp DEX share; weekly fees $14M (+56%)",
    "dYdX buyback to 75% net revenue; synthetic equity perps planned",
  ];
  s3.addText(defiItems.map((t, i) => ({
    text: t,
    options: { bullet: true, breakLine: i < defiItems.length - 1, fontSize: 10, fontFace: "Arial", color: C.bodyText }
  })), { x: 5.3, y: 3.47, w: 4.1, h: 1.45, margin: 0, paraSpaceAfter: 4 });

  s3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s3.addText("Crypto Market Weekly | 2026.03.23\u201303.29", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  s3.addNotes("The SEC-CFTC joint rule is the single most important regulatory event this week and arguably this year. Classifying 16 tokens as digital commodities \u2014 including BTC, ETH, SOL, XRP, ADA \u2014 ends the Howey Test litigation overhang and opens pathways for ETFs, institutional custody, and banking partnerships. XRP spot ETFs already attracted $1.4B in Q1 inflows; SOL ETFs are the next logical step.\n\nCircle's 20% stock crash is the market repricing stablecoin business models. The CLARITY Act draft bans passive yield, which would fundamentally undermine Circle's revenue (95% from reserve interest). Tether countered by hiring a Big Four firm for a full audit.\n\nAave hitting $1T cumulative lending is a DeFi milestone. Hyperliquid now commands 54% of perp DEX market share, with commodity contracts (crude oil daily volume >$2.2B, second only to BTC) emerging as a new growth vector. This signals on-chain derivatives expanding beyond crypto-native assets into traditional commodity markets.\n\nDeFi TVL rose 4.4% despite extreme fear \u2014 a remarkable divergence that suggests on-chain capital is more resilient than headline sentiment. dYdX's plan to launch synthetic equity perpetuals (starting with Tesla) could open another large addressable market.");

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
    { title: "SEC-CFTC Joint Digital Commodity Rule", desc: "16 tokens classified; token taxonomy established" },
    { title: "CLARITY Act Revised Draft", desc: "Bans passive stablecoin yield; Circle -20%" },
    { title: "CFTC Innovation Task Force Formed", desc: "Digital assets, AI, prediction market frameworks" },
    { title: "House Tokenization Hearing", desc: "Financial Services Committee on capital market modernization" },
    { title: "Outlook: CLARITY Act Markup", desc: "Full committee review expected late April" },
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
    fundRow("Mastercard/BVNK", "$1.8B", "Acquisition; TradFi into digital infra"),
    fundRow("Kalshi", "$1.0B", "Series E; event trading platform"),
    fundRow("GSR", "$57M", "Acquired Autonomous + Architech"),
    fundRow("Polymarket/Brahma", "N/A", "DeFi + prediction market merger"),
    fundRow("Rarible/Impossible", "N/A", "NFT marketplace M&A"),
  ], {
    x: 5.3, y: 1.45, w: 4.1, colW: [1.4, 0.8, 1.9],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.3, 0.3, 0.3, 0.3, 0.3, 0.3],
    autoPage: false,
  });

  // Callout
  s4.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 3.4, w: 4.1, h: 0.7, fill: { color: "FFF8E7" }, shadow: makeShadow() });
  s4.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 3.4, w: 0.06, h: 0.7, fill: { color: C.gold } });
  s4.addText([
    { text: "Key: ", options: { bold: true, fontSize: 10, fontFace: "Arial", color: C.navy } },
    { text: "Mastercard's $1.8B BVNK acquisition signals TradFi accelerating digital asset infrastructure buildout; 2026 crypto M&A on pace to exceed $37B record", options: { fontSize: 10, fontFace: "Arial", color: C.bodyText } }
  ], { x: 5.55, y: 3.45, w: 3.7, h: 0.6, margin: 0, valign: "middle" });

  s4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s4.addText("Crypto Market Weekly | 2026.03.23\u201303.29", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  s4.addNotes("Regulation: The SEC-CFTC joint rule is the most important regulatory breakthrough in crypto's history. Classifying BTC, ETH, SOL, and 13 other tokens as digital commodities ends the Howey Test grey zone and opens up ETFs, institutional custody, and banking partnerships. XRP spot ETFs already saw $1.4B in Q1 inflows. SOL ETFs are the next logical product.\n\nThe CLARITY Act draft is the week's negative regulatory surprise. Banning passive stablecoin yield directly threatens Circle's business model (95% of revenue from reserve interest). This will also reshape the broader stablecoin incentive structure. If passed, it would force issuers to find alternative revenue models.\n\nFundraising: Mastercard's $1.8B BVNK acquisition is the biggest deal this week, signaling traditional finance's commitment to digital asset infrastructure. Kalshi's $1B Series E validates the prediction market thesis institutionally. Total disclosed fundraising was ~$3.28B across 22 deals, showing capital deployment into crypto infrastructure hasn't slowed despite market fear.\n\n2026 crypto M&A is projected to exceed $37B \u2014 a new record driven by TradFi entering the space and consolidation among crypto-native firms.");

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
  s5.addText("Security Incidents  |  YTD Loss: ~$480M", { x: 0.6, y: 0.9, w: 4.2, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  s5.addTable([
    [
      { text: "Incident", options: { bold: true, fontSize: 9, fontFace: "Arial", fill: { color: C.lightGray }, color: C.bodyText } },
      { text: "Loss", options: { bold: true, fontSize: 9, fontFace: "Arial", fill: { color: C.lightGray }, color: C.bodyText, align: "right" } },
      { text: "Vector", options: { bold: true, fontSize: 9, fontFace: "Arial", fill: { color: C.lightGray }, color: C.bodyText } },
    ],
    [{ text: "Resolv USR", options: { fontSize: 9, fontFace: "Arial" } }, { text: "~$25M", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "AWS key compromise; 80M unbacked tokens", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
    [{ text: "Step Finance", options: { fontSize: 9, fontFace: "Arial" } }, { text: "~$27M", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "Largest YTD incident", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
    [{ text: "Truebit", options: { fontSize: 9, fontFace: "Arial" } }, { text: "~$26.6M", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "Third largest YTD", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
    [{ text: "Phishing (36 cases)", options: { fontSize: 9, fontFace: "Arial" } }, { text: "Rising", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "Social engineering now dominant vector", options: { fontSize: 9, fontFace: "Arial", color: C.red, bold: true } }],
  ], {
    x: 0.55, y: 1.35, w: 4.2, colW: [1.2, 0.8, 2.2],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.28, 0.35, 0.28, 0.28, 0.35],
    autoPage: false,
  });

  // AI x Crypto - right top
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 4.5, h: 1.85, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s5.addText("AI \u00D7 Crypto", { x: 5.3, y: 0.9, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  s5.addText([
    { text: "NVIDIA GTC: Huang declares \u201CAGI reached\u201D; $1T chip demand", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "AI crypto sector MCap >$28B; TAO leads at $3.4B", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "Miners pivot: cost $80K vs price $68K; difficulty -7.76%", options: { bullet: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
  ], { x: 5.3, y: 1.35, w: 4.1, h: 1.3, margin: 0, paraSpaceAfter: 4 });

  // Prediction Markets - right bottom
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 2.95, w: 4.5, h: 1.85, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 2.95, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s5.addText("Prediction Markets", { x: 5.3, y: 2.95, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  s5.addText([
    { text: "Polymarket: 11,176 active markets; 840K monthly wallets", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "Kalshi $1B Series E validates event trading at scale", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "Both platforms announce anti-insider trading rules (3/23)", options: { bullet: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
  ], { x: 5.3, y: 3.4, w: 4.1, h: 1.3, margin: 0, paraSpaceAfter: 4 });

  // Stablecoins - bottom left
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.5, w: 4.5, h: 1.5, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.5, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s5.addText("Stablecoins", { x: 0.6, y: 3.5, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });
  s5.addText([
    { text: "USDC volume overtakes USDT (64% share); Tether hires Big Four audit", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "CLARITY Act draft bans passive yield: Circle -20%, Coinbase -10%", options: { bullet: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
  ], { x: 0.6, y: 3.95, w: 4.1, h: 0.9, margin: 0, paraSpaceAfter: 4 });

  s5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s5.addText("Crypto Market Weekly | 2026.03.23\u201303.29", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  s5.addNotes("Security: The Resolv USR exploit is the week's major security event. An attacker compromised a privileged private key in AWS KMS and minted 80M unbacked USR tokens with <$200K collateral, causing USR to crash 97.5% in 17 minutes on Curve. Root causes were structural: single-key control, no oracle checks, no max mint limits. This highlights the risk of off-chain infrastructure dependencies in DeFi protocols.\n\nYTD 2026: 103 security incidents + 36 phishing scams totaling ~$480M. CertiK warns losses could exceed $500M for the full year. Social engineering has replaced code exploits as the dominant attack vector.\n\nAI x Crypto: NVIDIA's GTC 2026 AGI declaration and $1T chip demand projection sent AI tokens surging 10-20%. The more significant structural story is BTC miners' mass pivot to AI \u2014 mining costs of $80K/BTC far exceed the ~$68K market price, and miners have signed $70B+ in AI/HPC contracts. Mining difficulty dropped 7.76% on March 21. This could pose medium-term risks to BTC's hashrate security but creates opportunities for decentralized AI compute networks.\n\nPrediction Markets: Kalshi's $1B Series E is a strong institutional signal. Both Kalshi and Polymarket simultaneously announcing anti-insider trading rules shows proactive self-regulation ahead of expected CFTC oversight.");

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
    catRow("4/3", "Wormhole 600M token unlock (6% supply)", "W", "Supply pressure >$90M", true),
    catRow("4/15", "Starknet 127M token unlock", "STRK", "Supply pressure"),
    catRow("Q2", "Ethereum Glamsterdam Upgrade", "ETH", "78% gas reduction; may pull activity from L2s"),
    catRow("Late Apr", "CLARITY Act Committee Markup", "CRCL", "Stablecoin yield ban determination"),
    catRow("4/28-29", "Powell's Final FOMC Meeting", "All", "Rate forward guidance"),
    catRow("4/29-30", "TOKEN2049 Dubai", "All", "Year's biggest industry summit"),
    catRow("Q2", "SOL/XRP Spot ETF Progress", "SOL/XRP", "Commodity classification clears ETF path"),
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
    { text: "Extreme Fear (FGI=8). BTC hit March low; ETH broke $2,000. Triple macro pressure: Iran-Israel war, Trump 15% tariffs, Fed inflation forecast upgrade. ETF flows turned sharply negative.", options: { breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "Key watches: ", options: { bold: true, fontSize: 10, fontFace: "Arial", color: C.navy } },
    { text: "4/3 Wormhole $90M unlock, CLARITY Act markup impact on stablecoin positions, Iran-Israel trajectory on oil/inflation. SEC-CFTC commodity classification is a major medium-term positive.", options: { fontSize: 10, fontFace: "Arial", color: C.bodyText } },
  ], { x: 0.8, y: 4.0, w: 8.5, h: 0.85, margin: 0 });

  s6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s6.addText("Crypto Market Weekly | 2026.03.23\u201303.29  |  CONFIDENTIAL", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  s6.addNotes("Catalyst focus for next two weeks:\n\n1. Wormhole Token Unlock (4/3) \u2014 600M tokens (6% of supply, >$90M value) is the largest near-term unlock event. In a fear-driven market, supply events hit harder.\n\n2. CLARITY Act Committee Markup (late April) \u2014 The outcome will determine the future of stablecoin yield products and directly impact Circle, Coinbase, and any stablecoin position in the portfolio.\n\n3. Powell's Final FOMC (4/28-29) \u2014 With PCE inflation forecast at 2.7% and tariff/oil pressures ongoing, rate cut expectations are being pushed further out. This FOMC sets the tone for Q2-Q3 monetary policy.\n\n4. Ethereum Glamsterdam \u2014 A 78% gas reduction could be a game-changer for mainnet activity, potentially reversing the L2 migration trend.\n\nOverall assessment: The market is in a capitulation zone (FGI=8) driven by geopolitical and macro headwinds. But the fundamental backdrop is divergent: DeFi TVL +4.4%, Aave $1T lending, Hyperliquid fees $14M/week, stablecoin adoption accelerating, and the historic SEC-CFTC commodity classification not yet priced in. This suggests we may be in a late-cycle fear phase where sentiment overshoots fundamentals.\n\nFor LP positioning: Maintain exposure; the SEC-CFTC classification is a structural unlock. The primary risks to monitor are the CLARITY Act (stablecoin positions), Wormhole unlock (W exposure), and Iran-Israel trajectory (broad market risk).");

  const outPath = "/Users/lx/Github/crypto-weekly-digest/reports/2026-03-23/report-en.pptx";
  await pres.writeFile({ fileName: outPath });
  console.log("\u2705 Written to " + outPath);
}

generate().catch(e => { console.error(e); process.exit(1); });
