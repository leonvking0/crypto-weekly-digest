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
  pres.title = "Crypto Market Weekly Report \u2014 Mar 9-15, 2026";

  // ============ SLIDE 1: Title ============
  const s1 = pres.addSlide();
  s1.background = { color: C.darkNavy };
  s1.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.6, w: 1.2, h: 0.04, fill: { color: C.gold } });
  s1.addText("Crypto Market Weekly", {
    x: 0.8, y: 1.8, w: 8, h: 1.0,
    fontSize: 42, fontFace: "Arial", bold: true, color: C.white, margin: 0
  });
  s1.addText("March 9 \u2013 15, 2026", {
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
  s1.addNotes("Weekly crypto market report for March 9-15, 2026. This was a volatile week dominated by a US-Iran geopolitical flash crash and subsequent recovery, alongside landmark regulatory developments including the SEC-CFTC joint regulatory MOU.");

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
  s2.addText("$71,434", { x: cardStartX + 0.25, y: cardY + 0.35, w: 2, h: 0.4, fontSize: 28, fontFace: "Arial", color: C.bodyText, bold: true, margin: 0 });
  s2.addText("+6.3%", { x: cardStartX + 0.25, y: cardY + 0.72, w: 2, h: 0.25, fontSize: 14, fontFace: "Arial", color: C.green, bold: true, margin: 0 });

  // Card 2: ETH
  const c2x = cardStartX + cardW + cardGap;
  s2.addShape(pres.shapes.RECTANGLE, { x: c2x, y: cardY, w: cardW, h: cardH, fill: { color: C.white }, shadow: makeShadow() });
  s2.addShape(pres.shapes.RECTANGLE, { x: c2x, y: cardY, w: 0.06, h: cardH, fill: { color: C.gold } });
  s2.addText("ETH", { x: c2x + 0.25, y: cardY + 0.1, w: 2, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.medGray, bold: true, margin: 0 });
  s2.addText("$2,092", { x: c2x + 0.25, y: cardY + 0.35, w: 2, h: 0.4, fontSize: 28, fontFace: "Arial", color: C.bodyText, bold: true, margin: 0 });
  s2.addText("+7.5%", { x: c2x + 0.25, y: cardY + 0.72, w: 2, h: 0.25, fontSize: 14, fontFace: "Arial", color: C.green, bold: true, margin: 0 });

  // Card 3: Fear & Greed
  const c3x = c2x + cardW + cardGap;
  s2.addShape(pres.shapes.RECTANGLE, { x: c3x, y: cardY, w: cardW, h: cardH, fill: { color: C.white }, shadow: makeShadow() });
  s2.addShape(pres.shapes.RECTANGLE, { x: c3x, y: cardY, w: 0.06, h: cardH, fill: { color: C.red } });
  s2.addText("Fear & Greed Index", { x: c3x + 0.25, y: cardY + 0.1, w: 2.4, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.medGray, bold: true, margin: 0 });
  s2.addText("15", { x: c3x + 0.25, y: cardY + 0.35, w: 1, h: 0.4, fontSize: 28, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
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
    makeRow("SOL", "$87.61", "+6.8%", "Alpenglow upgrade approved", false),
    makeRow("HYPE", "$37.13", "+21.7%", "Best performer; HyperEVM ecosystem", false),
    makeRow("PENDLE", "$1.27", "+7.4%", "Recovery with broader market", false),
    makeRow("ENA", "$0.11", "+9.2%", "", false),
    makeRow("DYDX", "$0.09", "+13.4%", "Oversold bounce", false),
  ], {
    x: 0.6, y: 2.25, w: 8.8, colW: [1.3, 1.8, 1.5, 4.2],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.35, 0.33, 0.33, 0.33, 0.33, 0.33],
    autoPage: false,
  });

  s2.addText("Macro: US-Iran geopolitical escalation triggered flash crash (BTC hit $62,400) \u2192 recovered on institutional inflows ($53.8M BTC ETF) + regulatory tailwinds | Total MCap $2.51T | BTC Dom 57.0%", {
    x: 0.6, y: 4.55, w: 8.8, h: 0.4,
    fontSize: 10, fontFace: "Arial", color: C.medGray, italic: true, margin: 0
  });

  s2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s2.addText("Crypto Market Weekly | 2026.03.09\u201303.15", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  s2.addNotes("Volatile week: US-Iran military escalation on 3/9-10 triggered a broad risk-off selloff. BTC flash-crashed to $62,400 intraday before rapidly recovering to $68,366. By week end, BTC settled at $71,434 (+6.3%). HYPE was the standout performer at +21.7%, driven by HyperEVM ecosystem momentum. Despite the strong recovery, Fear & Greed Index remains at 15 (Extreme Fear), highlighting a stark divergence between institutional behavior (continued ETF inflows of $53.8M over 3 days) and retail sentiment. AHR999 indicator at 0.33-0.36 suggests we\u2019re in a \u201Cbottom accumulation\u201D zone. All portfolio positions (PENDLE, ENA, HYPE) posted solid weekly gains.");

  // ============ SLIDE 3: Top Stories + Industry ============
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
    ["1", "SEC-CFTC Joint Regulatory MOU (3/11)", "Ends years of jurisdictional disputes; unified oversight era begins"],
    ["2", "Ripple $750M buyback at $50B val", "Fortune 500-level capital ops from a private crypto company"],
    ["3", "Geopolitical flash crash & recovery", "US-Iran escalation sent BTC to $62,400; rapid rebound to $68K+"],
    ["4", "Solana Alpenglow approved (98.27%)", "Finality from 12.8s to 100-150ms \u2014 generational upgrade"],
    ["5", "SEC drops BitClout/DeSo case (3/14)", "Signals shift from litigation enforcement to compliance guidance"],
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

  s3.addText([
    { text: "Solana Alpenglow: 98.27% staker approval; finality \u2192 100ms", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "Avalanche TVL breaks $2.1B; Q1 doubled (Octane upgrade)", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "Monad mainnet: sustained 10,000 TPS + sub-second finality", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "Base accelerates custom arch, leaving OP Stack", options: { bullet: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
  ], { x: 5.3, y: 1.32, w: 4.1, h: 1.45, margin: 0, paraSpaceAfter: 4 });

  // Right bottom: DeFi
  s3.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 3.05, w: 4.5, h: 1.95, fill: { color: C.white }, shadow: makeShadow() });
  s3.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 3.05, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s3.addText("DeFi & Perp DEXes", { x: 5.3, y: 3.05, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  s3.addText([
    { text: "Aave V4 \u201CHub and Spoke\u201D cross-chain unified liquidity", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "BlackRock staking ETH ETF (ETHB) launched on Nasdaq", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "Hyperliquid OI peak $6.48B; enters global Top 10", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "Perp DEX market share at 10.2%; structural shift confirmed", options: { bullet: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
  ], { x: 5.3, y: 3.47, w: 4.1, h: 1.45, margin: 0, paraSpaceAfter: 4 });

  s3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s3.addText("Crypto Market Weekly | 2026.03.09\u201303.15", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  s3.addNotes("The SEC-CFTC joint MOU is arguably the most significant US regulatory development in crypto history. It formally ends the \u201Cis it a security or commodity?\u201D jurisdictional turf war that has plagued the industry. This creates a framework for joint review of new digital asset products and information sharing.\n\nRipple\u2019s $750M buyback at a $50B valuation demonstrates the maturity of large-cap private crypto companies. This is Fortune 500-level corporate finance.\n\nSolana\u2019s Alpenglow upgrade is a generational leap \u2014 reducing finality from 12.8 seconds to 100-150 milliseconds. The 98.27% staker approval shows overwhelming network consensus. This positions SOL competitively against even the fastest new L1s like Monad.\n\nHyperliquid entering the global Top 10 exchanges (including CEXes) by open interest at $6.48B is a watershed moment for on-chain derivatives. The commodity perp breakout \u2014 crude oil perps surpassing ETH volume during the geopolitical crisis \u2014 signals DEXes expanding beyond crypto-native assets.");

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
    { title: "SEC-CFTC Joint MOU (3/11)", desc: "Info sharing, joint review, unified definitions framework" },
    { title: "SEC Token Classification", desc: "Submitted to White House; shift from litigation to rulemaking" },
    { title: "CFTC Project Crypto (3/9)", desc: "Unified crypto classification system + DeFi registration rules" },
    { title: "Treasury DeFi Risk Report", desc: "Proposes \u201Cwithholding law\u201D for AML/CFT compliance" },
    { title: "Stablecoin Yield Ban Debate", desc: "Senate reviews 130+ amendments; may ban retail yield stables" },
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
    fundRow("Ripple Buyback", "$750M", "Valuation $50B"),
    fundRow("ASTER Unlock", "~$56M", "3/17: 78.14M tokens"),
    fundRow("LayerZero ZRO", "~$45M", "3/20 unlock; 2.47% of MCap"),
    fundRow("ETH Foundation", "$10M", "OTC sale of 5,000 ETH"),
  ], {
    x: 5.3, y: 1.45, w: 4.1, colW: [1.3, 0.9, 1.9],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.3, 0.3, 0.3, 0.3, 0.3],
    autoPage: false,
  });

  // Callout
  s4.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 3.1, w: 4.1, h: 0.85, fill: { color: "FFF8E7" }, shadow: makeShadow() });
  s4.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 3.1, w: 0.06, h: 0.85, fill: { color: C.gold } });
  s4.addText([
    { text: "Trend: ", options: { bold: true, fontSize: 10, fontFace: "Arial", color: C.navy } },
    { text: "Last quarter fundraising volume down 62%, but avg deal size up 272% to $34M. Industry M&A expected to exceed $37B in 2026. Kraken, BitGo, Chainalysis top IPO candidates.", options: { fontSize: 10, fontFace: "Arial", color: C.bodyText } }
  ], { x: 5.55, y: 3.15, w: 3.7, h: 0.75, margin: 0, valign: "middle" });

  s4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s4.addText("Crypto Market Weekly | 2026.03.09\u201303.15", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  s4.addNotes("Regulation: The SEC-CFTC MOU is the single most important regulatory development. Combined with the SEC removing crypto from its 2026 examination priorities and dropping the BitClout case, we\u2019re seeing a clear pivot from adversarial enforcement to constructive rulemaking under Atkins\u2019 leadership.\n\nThe stablecoin yield ban debate in the Senate is a key risk to watch \u2014 particularly relevant for our Ethena (ENA) position. If retail yield-bearing stablecoins are banned, USDe\u2019s distribution model would need significant restructuring.\n\nFundraising: Ripple\u2019s $750M buyback at $50B valuation is the headline. The shift toward larger but fewer deals (avg $34M, up 272%) signals institutional-grade capital entering the space. The $37B M&A forecast for 2026 would represent a paradigm shift in industry consolidation.");

  // ============ SLIDE 5: Security + AI + Prediction ============
  const s5 = pres.addSlide();
  s5.background = { color: C.offWhite };
  s5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy } });
  s5.addText("\uD83D\uDD13  Security & Sector Highlights", {
    x: 0.6, y: 0, w: 9, h: 0.7,
    fontSize: 20, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0
  });

  // Security - left
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 2.35, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 0.35, fill: { color: C.red } });
  s5.addText("Security Incidents  |  Weekly Loss: ~$81M", { x: 0.6, y: 0.9, w: 4.2, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  s5.addTable([
    [
      { text: "Incident", options: { bold: true, fontSize: 9, fontFace: "Arial", fill: { color: C.lightGray }, color: C.bodyText } },
      { text: "Loss", options: { bold: true, fontSize: 9, fontFace: "Arial", fill: { color: C.lightGray }, color: C.bodyText, align: "right" } },
      { text: "Vector", options: { bold: true, fontSize: 9, fontFace: "Arial", fill: { color: C.lightGray }, color: C.bodyText } },
    ],
    [{ text: "Aave Slippage", options: { fontSize: 9, fontFace: "Arial" } }, { text: "~$50M", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "Extreme slippage/MEV", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
    [{ text: "Aave Oracle", options: { fontSize: 9, fontFace: "Arial" } }, { text: "$27.78M", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "wstETH config error (refunded)", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
    [{ text: "Solv Protocol", options: { fontSize: 9, fontFace: "Arial" } }, { text: "$2.73M", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "Unaudited contract", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
    [{ text: "Bonk.fun DNS", options: { fontSize: 9, fontFace: "Arial" } }, { text: "N/A", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "DNS hijack + phishing", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
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
    { text: "Messari: 2026 is AI-blockchain \u201Csystemic integration\u201D year", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "TRON joins Agentic AI Foundation (first major L1)", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "ElizaOS becomes standard AI Agent personality framework", options: { bullet: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
  ], { x: 5.3, y: 1.35, w: 4.1, h: 1.3, margin: 0, paraSpaceAfter: 4 });

  // Prediction Markets - right bottom
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 2.95, w: 4.5, h: 1.85, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 2.95, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s5.addText("Prediction Markets", { x: 5.3, y: 2.95, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  s5.addText([
    { text: "Polymarket & Kalshi eyeing $20B valuations", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "Ultra-short contracts hit $70M daily; Vitalik criticizes", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "Regulatory pushback: insider trading bill + Utah gambling ban", options: { bullet: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
  ], { x: 5.3, y: 3.4, w: 4.1, h: 1.3, margin: 0, paraSpaceAfter: 4 });

  // Stablecoins
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.5, w: 4.5, h: 1.5, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.5, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s5.addText("Stablecoins", { x: 0.6, y: 3.5, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });
  s5.addText([
    { text: "Yield stablecoins MCap $22.7B; growing 15x faster than market", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "Visa expands USDC settlement; Hashdex forecasts $600B by EOY", options: { bullet: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
  ], { x: 0.6, y: 3.95, w: 4.1, h: 0.9, margin: 0, paraSpaceAfter: 4 });

  s5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s5.addText("Crypto Market Weekly | 2026.03.09\u201303.15", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  s5.addNotes("Security: Aave dominated security incidents this week with ~$78M in losses across two events. The oracle misconfiguration ($27.78M) was a config error in the wstETH oracle cap \u2014 DAO has committed to full reimbursement. The $50M slippage/MEV event highlights systemic risks in DeFi during high-volatility periods. The Solv Protocol exploit ($2.73M from unaudited contracts) reinforces the importance of audit coverage.\n\nAI x Crypto: Messari\u2019s \u201CCrypto Theses 2026\u201D positioning this as the year of \u201Csystemic integration\u201D between AI and blockchain is a major institutional signal. TRON joining the Agentic AI Foundation is the first major L1 to formally commit to autonomous AI agent infrastructure.\n\nPrediction Markets: Both Polymarket and Kalshi targeting $20B valuations shows the sector\u2019s explosive growth trajectory. However, regulatory headwinds are mounting \u2014 insider trading legislation, state-level gambling reclassification (Utah), and the \u201Cdeath contract\u201D class action against Kalshi all represent risks.\n\nStablecoins: The yield stablecoin boom ($22.7B MCap, 15x market growth rate) is directly relevant to our ENA position. Senate debate on banning retail yield stablecoins is a critical risk factor.");

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
    catRow("3/16", "Neutron Upgrade", "NTRN", "Enhanced cross-chain capabilities"),
    catRow("3/17", "ASTER Token Unlock ($56M)", "ASTER", "Supply pressure"),
    catRow("3/18", "FOMC Rate Decision", "ALL", "Key macro event", true),
    catRow("3/20", "LayerZero ZRO Unlock ($45M)", "ZRO", "Supply pressure (2.47% of MCap)"),
    catRow("3/24-26", "Digital Asset Summit (NYC)", "Institutional", "TradFi integration + AI convergence"),
    catRow("3/27", "Crypto ETP Final Deadline", "Multiple", "Approval could unlock institutional wave"),
    catRow("Late Mar", "Midnight Mainnet Launch", "ADA", "Cardano privacy sidechain"),
  ], {
    x: 0.5, y: 0.9, w: 9, colW: [0.9, 3.2, 1.2, 3.7],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.32, 0.32, 0.32, 0.32, 0.32, 0.32, 0.32, 0.32],
    autoPage: false,
  });

  // Outlook
  s6.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.65, w: 9, h: 1.3, fill: { color: C.white }, shadow: makeShadow() });
  s6.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.65, w: 0.06, h: 1.3, fill: { color: C.gold } });
  s6.addText("Weekly Summary", { x: 0.8, y: 3.7, w: 8, h: 0.3, fontSize: 13, fontFace: "Arial", bold: true, color: C.navy, margin: 0 });
  s6.addText([
    { text: "BTC rebounded +6.3% after geopolitical flash crash; HYPE led at +21.7%. SEC-CFTC joint MOU marks a historic regulatory milestone, but FGI remains at Extreme Fear (15), showing stark retail-institutional divergence.", options: { breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "Key watches: ", options: { bold: true, fontSize: 10, fontFace: "Arial", color: C.navy } },
    { text: "3/18 FOMC decision (portfolio-wide), ASTER/ZRO unlock pressure, Aave oracle incident\u2019s implications for DeFi safety mechanisms.", options: { fontSize: 10, fontFace: "Arial", color: C.bodyText } },
  ], { x: 0.8, y: 4.0, w: 8.5, h: 0.85, margin: 0 });

  s6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s6.addText("Crypto Market Weekly | 2026.03.09\u201303.15  |  CONFIDENTIAL", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  s6.addNotes("Near-term catalyst focus:\n\n1. FOMC Rate Decision (3/18) \u2014 The single most important macro event. Markets are pricing in a hold, but any hawkish surprise could trigger another selloff given the fragile Extreme Fear sentiment. Dovish hints would likely catalyze a breakout above $72K.\n\n2. ASTER ($56M, 3/17) and ZRO ($45M, 3/20) unlocks add ~$100M of combined supply pressure. While not directly portfolio-relevant, they contribute to overall market selling pressure.\n\n3. Digital Asset Summit (3/24-26, NYC) could produce institutional announcements. The crypto ETP deadline on 3/27 is a binary event \u2014 approvals would unlock significant new institutional capital flows.\n\nPortfolio assessment: All positions posted strong weekly gains (PENDLE +7.4%, ENA +9.2%, HYPE +21.7%). The AHR999 indicator at 0.33-0.36 suggests we remain in a \u201Cbottom accumulation\u201D zone despite the recovery. The institutional-retail sentiment divergence (ETF inflows vs Extreme Fear) typically resolves with a sharp move \u2014 positioning suggests upward resolution but FOMC risk is real.\n\nRisk flag: Senate stablecoin yield ban debate is a medium-term threat to ENA thesis. Monitor closely.");

  const outPath = "/Users/lx/Github/crypto-weekly-digest/reports/2026-03-09/slides-en.pptx";
  await pres.writeFile({ fileName: outPath });
  console.log("\u2705 Written to " + outPath);
}

generate().catch(e => { console.error(e); process.exit(1); });
