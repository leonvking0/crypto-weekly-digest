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
  pres.title = "Crypto Market Weekly Report \u2014 Mar 30\u2013Apr 5, 2026";

  // ============ SLIDE 1: Title ============
  const s1 = pres.addSlide();
  s1.background = { color: C.darkNavy };
  s1.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.6, w: 1.2, h: 0.04, fill: { color: C.gold } });
  s1.addText("Crypto Market Weekly", {
    x: 0.8, y: 1.8, w: 8, h: 1.0,
    fontSize: 42, fontFace: "Arial", bold: true, color: C.white, margin: 0
  });
  s1.addText("March 30 \u2013 April 5, 2026", {
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
  s1.addNotes("Welcome to the weekly crypto market report covering March 30 \u2013 April 5, 2026. This report is prepared for LP distribution and covers market performance, key developments, regulatory updates, security incidents, and upcoming catalysts. This week was dominated by the $285M Drift Protocol exploit (the largest single hack of 2026), significant US regulatory progress (GENIUS Act implementation rules, CLARITY Act compromise), and Aave V4 mainnet launch. The Fear & Greed Index sits at 12 \u2014 Extreme Fear \u2014 reflecting deep market pessimism despite improving fundamentals. All price data sourced from CoinGecko API.");

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
  s2.addText("$67,197", { x: cardStartX + 0.25, y: cardY + 0.35, w: 2, h: 0.4, fontSize: 28, fontFace: "Arial", color: C.bodyText, bold: true, margin: 0 });
  s2.addText("+1.1%", { x: cardStartX + 0.25, y: cardY + 0.72, w: 2, h: 0.25, fontSize: 14, fontFace: "Arial", color: C.green, bold: true, margin: 0 });

  // Card 2: ETH
  const c2x = cardStartX + cardW + cardGap;
  s2.addShape(pres.shapes.RECTANGLE, { x: c2x, y: cardY, w: cardW, h: cardH, fill: { color: C.white }, shadow: makeShadow() });
  s2.addShape(pres.shapes.RECTANGLE, { x: c2x, y: cardY, w: 0.06, h: cardH, fill: { color: C.gold } });
  s2.addText("ETH", { x: c2x + 0.25, y: cardY + 0.1, w: 2, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.medGray, bold: true, margin: 0 });
  s2.addText("$2,050", { x: c2x + 0.25, y: cardY + 0.35, w: 2, h: 0.4, fontSize: 28, fontFace: "Arial", color: C.bodyText, bold: true, margin: 0 });
  s2.addText("+3.0%", { x: c2x + 0.25, y: cardY + 0.72, w: 2, h: 0.25, fontSize: 14, fontFace: "Arial", color: C.green, bold: true, margin: 0 });

  // Card 3: Fear & Greed
  const c3x = c2x + cardW + cardGap;
  s2.addShape(pres.shapes.RECTANGLE, { x: c3x, y: cardY, w: cardW, h: cardH, fill: { color: C.white }, shadow: makeShadow() });
  s2.addShape(pres.shapes.RECTANGLE, { x: c3x, y: cardY, w: 0.06, h: cardH, fill: { color: C.red } });
  s2.addText("Fear & Greed Index", { x: c3x + 0.25, y: cardY + 0.1, w: 2.4, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.medGray, bold: true, margin: 0 });
  s2.addText("12", { x: c3x + 0.25, y: cardY + 0.35, w: 1, h: 0.4, fontSize: 28, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
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
    makeRow("SOL", "$79.45", "-2.8%", "Drift hack impact; -11% MAA", true),
    makeRow("HYPE", "$35.47", "-9.0%", "Q1 vol $492.7B; global top 10", true),
    makeRow("PENDLE", "$0.98", "-12.4%", "Yield sector under pressure", true),
    makeRow("ENA", "$0.08", "-14.4%", "Broad alt weakness", true),
    makeRow("DYDX", "$0.10", "+7.3%", "Technical bounce", false),
  ], {
    x: 0.6, y: 2.25, w: 8.8, colW: [1.3, 1.8, 1.5, 4.2],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.35, 0.33, 0.33, 0.33, 0.33, 0.33],
    autoPage: false,
  });

  s2.addText("Macro: Fed holds 3.50\u20133.75%, inflation forecast raised; Drift $285M hack rattles Solana ecosystem; sentiment at Extreme Fear (FGI=12).", {
    x: 0.6, y: 4.55, w: 8.8, h: 0.4,
    fontSize: 10, fontFace: "Arial", color: C.medGray, italic: true, margin: 0
  });

  s2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s2.addText("Crypto Market Weekly | 2026.03.30\u201304.05", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  s2.addNotes("BTC traded in a tight $65,000\u2013$70,000 range, briefly testing $69,000 resistance on Thursday before settling at $67,197 (+1.1% WoW). ETH outperformed slightly at +3.0%, oscillating between $2,020\u2013$2,140, though spot ETH ETFs saw $42M in net outflows indicating continued institutional caution.\n\nAlt performance was broadly negative: HYPE -9.0%, ENA -14.4%, PENDLE -12.4%. The yield token sector (PENDLE, ENA) continues to face headwinds as base staking yields compress. DYDX was the sole positive outlier at +7.3%, driven by a technical bounce rather than fundamental improvement.\n\nFear & Greed Index at 12 marks one of the lowest readings since the FTX collapse era. Historically, extreme fear readings have been contrarian buy signals, but the current macro backdrop (Fed holding rates high, geopolitical tensions from Iran-Israel conflict) provides legitimate reasons for caution.\n\nThe Drift Protocol $285M exploit was the week's black swan, directly causing SOL to underperform as the worst major asset. Network monthly active addresses dropped 11%, and the hack raised fundamental questions about DeFi security models \u2014 specifically, the vulnerability of oracle-dependent protocols to sophisticated social engineering + oracle manipulation attack vectors.\n\nSpot BTC ETF flows were barely positive (+$22M), suggesting institutional conviction remains low despite BTC holding its range.");

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
    ["1", "Drift Protocol $285M exploit", "Lazarus Group; social engineering + oracle manipulation; 12 min"],
    ["2", "US Treasury: GENIUS Act rules", "1:1 reserves + 12-month operational backstop for stablecoins"],
    ["3", "CLARITY Act yield compromise nears", "Bans passive yield; allows activity-based rewards"],
    ["4", "Aave V4 mainnet + BGD Labs exit", "Modular hub-and-spoke; Prime/Core/Plus liquidity hubs"],
    ["5", "Solana Alpenglow: 98% validator approval", "Finality from ~400ms to 100\u2013150ms"],
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
    "Base: 50% L2 DEX volume, $4.15B TVL, 3.3B txns YTD",
    "\"Stablechains\" narrative: a16z leads $10M for Better Money",
    "Monad TVL >$350M; L1 FDV reset trend (Messari thesis)",
    "Cosmos setback: Leap Wallet shuts down (4/3)",
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
    "Aave V4: modular hub-and-spoke; Lido + Ethena partners",
    "Lido V3 stVaults: EarnETH + EarnUSD multi-yield products",
    "Hyperliquid: Q1 $492.7B vol; enters global top-10 deriv",
    "DeFi TVL ~$92\u201393B; RWA sector grows to ~$17B",
  ];
  s3.addText(defiItems.map((t, i) => ({
    text: t,
    options: { bullet: true, breakLine: i < defiItems.length - 1, fontSize: 10, fontFace: "Arial", color: C.bodyText }
  })), { x: 5.3, y: 3.47, w: 4.1, h: 1.45, margin: 0, paraSpaceAfter: 4 });

  s3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s3.addText("Crypto Market Weekly | 2026.03.30\u201304.05", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  s3.addNotes("Top story analysis:\n\n1. Drift Protocol $285M exploit: Executed by North Korea's Lazarus Group after a 6-month social engineering campaign. Attackers posed as a quant trading firm to trick multisig signers into pre-signing transactions, then created a fictitious token and manipulated oracle prices to use it as collateral. This exposes a critical vulnerability in the DeFi security model \u2014 on-chain code audits alone are insufficient; off-chain operational security (personnel, key management, multisig processes) is equally critical.\n\n2. GENIUS Act implementation rules represent the first comprehensive US stablecoin regulatory framework. The 12-month operational expense backstop requirement will price out smaller issuers while being manageable for Tether and Circle.\n\n3. CLARITY Act yield compromise is a major positive \u2014 banning passive hold yields but allowing activity-based rewards preserves DeFi protocol revenue models. This is better than the outright ban that crashed Circle stock 20% last week.\n\n4. Aave V4's modular architecture is a landmark moment for DeFi institutionalization. BGD Labs exit highlights structural governance tensions in DAOs.\n\n5. Hyperliquid entering global top-10 derivatives venues ($492.7B Q1 volume) is a thesis-confirming milestone for on-chain perpetuals. The platform now competes directly with centralized exchanges. Grayscale's HYPE spot ETF S-1 filing accelerates institutional recognition.\n\nBase continues to dominate L2s with 50% DEX volume share. The \"Stablechains\" narrative (purpose-built chains for stablecoin settlement) is a significant new development worth monitoring for portfolio implications.");

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
    { title: "GENIUS Act Implementation", desc: "Treasury 4/1: 1:1 reserves + 12-month op backstop" },
    { title: "CLARITY Act Compromise", desc: "Bans passive yield; allows activity rewards; markup Apr" },
    { title: "CFTC Policy Pivot", desc: "New enforcement dir: no more regulation by litigation" },
    { title: "Banking Infrastructure", desc: "Kraken: Fed master acct; Ripple/Crypto.com: OCC charters" },
    { title: "Schwab Opens Crypto Trading", desc: "4/4 waitlist for BTC/ETH; full launch expected Q2" },
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
    fundRow("ZeroStack/0G", "$107M", "Ecosystem; 21% supply"),
    fundRow("Midas", "$50M", "Series A; RWA liquidity"),
    fundRow("Keyrock", "Series C", "Val $1.1B; Ripple backed"),
    fundRow("Franklin/250D", "M&A", "Forms Franklin Crypto"),
    fundRow("MicroStrategy", "44,377 BTC", "Mar; 94% of corp net adds"),
  ], {
    x: 5.3, y: 1.45, w: 4.1, colW: [1.3, 0.9, 1.9],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.3, 0.3, 0.3, 0.3, 0.3, 0.3],
    autoPage: false,
  });

  // Mining callout
  s4.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 3.4, w: 4.1, h: 0.7, fill: { color: "FFF8E7" }, shadow: makeShadow() });
  s4.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 3.4, w: 0.06, h: 0.7, fill: { color: C.gold } });
  s4.addText([
    { text: "Key: ", options: { bold: true, fontSize: 10, fontFace: "Arial", color: C.navy } },
    { text: "Miners accelerate BTC sales for AI pivot: MARA sold 15,133 BTC; Riot sold $290M in BTC for HPC buildout", options: { fontSize: 10, fontFace: "Arial", color: C.bodyText } }
  ], { x: 5.55, y: 3.45, w: 3.7, h: 0.6, margin: 0, valign: "middle" });

  s4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s4.addText("Crypto Market Weekly | 2026.03.30\u201304.05", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  s4.addNotes("Regulation: This week represents a watershed moment for US crypto regulatory clarity. Three major developments occurred simultaneously:\n\n1. GENIUS Act implementation rules from Treasury establish the first comprehensive federal stablecoin framework. The 1:1 high-quality liquid asset backing requirement and 12-month operational expense backstop are strict but reasonable for major issuers. Smaller issuers (sub-$10B circulation) will be regulated at state level.\n\n2. CLARITY Act yield compromise resolves the most contentious issue that crashed Circle stock 20% last week. The Tillis-Alsobrooks language bans passive hold yields but allows activity-based rewards \u2014 this is structurally bullish for DeFi protocols that offer yield through active participation rather than passive holding.\n\n3. CFTC's new enforcement director explicitly stating they will no longer use litigation to set policy marks the formal end of \"regulation by enforcement\" in the US. Combined with Kraken securing a Fed master account (first crypto-native bank to do so), the institutional infrastructure for crypto is rapidly maturing.\n\nCharles Schwab opening a BTC/ETH trading waitlist signals that mainstream retail brokerage crypto access is imminent.\n\nFundraising: ZeroStack's $107M commitment to the 0G ecosystem and Franklin Templeton's acquisition of 250 Digital reflect TradFi's accelerating push into crypto infrastructure. MicroStrategy accumulated 44,377 BTC in March (94% of all corporate net additions), while miners continue liquidating BTC holdings to fund AI/HPC transitions \u2014 a structural supply-side shift worth monitoring.");

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
  s5.addText("Security Incidents  |  Weekly Loss: >$310M", { x: 0.6, y: 0.9, w: 4.2, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  s5.addTable([
    [
      { text: "Incident", options: { bold: true, fontSize: 9, fontFace: "Arial", fill: { color: C.lightGray }, color: C.bodyText } },
      { text: "Loss", options: { bold: true, fontSize: 9, fontFace: "Arial", fill: { color: C.lightGray }, color: C.bodyText, align: "right" } },
      { text: "Vector", options: { bold: true, fontSize: 9, fontFace: "Arial", fill: { color: C.lightGray }, color: C.bodyText } },
    ],
    [{ text: "Drift Protocol", options: { fontSize: 9, fontFace: "Arial" } }, { text: "$285M", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red, bold: true } }, { text: "Social eng + oracle manip", options: { fontSize: 9, fontFace: "Arial", color: C.red } }],
    [{ text: "Resolv contagion", options: { fontSize: 9, fontFace: "Arial" } }, { text: "$25M", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "AWS key + cascading bad debt", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
    [{ text: "SIREN rug pull", options: { fontSize: 9, fontFace: "Arial" } }, { text: ">$100M", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "50% supply concentration", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
    [{ text: "Venus Protocol", options: { fontSize: 9, fontFace: "Arial" } }, { text: "$2.18M", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "Hybrid on/off-chain", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
  ], {
    x: 0.55, y: 1.35, w: 4.2, colW: [1.4, 1.0, 1.8],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.28, 0.28, 0.28, 0.28, 0.28],
    autoPage: false,
  });

  // AI x Crypto - right top
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 4.5, h: 1.85, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s5.addText("AI \u00D7 Crypto", { x: 5.3, y: 0.9, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  s5.addText([
    { text: "AI tokens rally: FET/Virtuals +35\u201345%, TAO +73%", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "ZeroStack commits $107M to 0G AI ecosystem", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "Render integrates NVIDIA Blackwell B200 architecture", options: { bullet: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
  ], { x: 5.3, y: 1.35, w: 4.1, h: 1.3, margin: 0, paraSpaceAfter: 4 });

  // Prediction Markets - right bottom
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 2.95, w: 4.5, h: 1.85, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 2.95, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s5.addText("Prediction Markets", { x: 5.3, y: 2.95, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  s5.addText([
    { text: "March total: $25.7B (+10.6%); Kalshi $13.1B leads", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "Polymarket: $9B valuation; ICE $2B investment", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "CFTC sues 3 states over federal jurisdiction claims", options: { bullet: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
  ], { x: 5.3, y: 3.4, w: 4.1, h: 1.3, margin: 0, paraSpaceAfter: 4 });

  // Stablecoins - bottom left
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.5, w: 4.5, h: 1.5, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.5, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s5.addText("Stablecoins", { x: 0.6, y: 3.5, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });
  s5.addText([
    { text: "MCap ~$325B; Mar vol $7.5T surpasses US ACH network", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "USDC supply +$2B vs USDT -$3B; divergence accelerates", options: { bullet: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
  ], { x: 0.6, y: 3.95, w: 4.1, h: 0.9, margin: 0, paraSpaceAfter: 4 });

  s5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s5.addText("Crypto Market Weekly | 2026.03.30\u201304.05", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  s5.addNotes("Security: This was the worst week for crypto security losses in 2026, with total damages exceeding $310M. The Drift Protocol $285M hack is the year's largest single incident and the most sophisticated attack we've seen \u2014 combining 6-month social engineering infiltration, oracle price manipulation via a fictitious token, and multisig compromise. Key lesson: DeFi security cannot rely solely on smart contract audits; operational security (personnel vetting, key management, multisig governance) is equally critical.\n\nThe Resolv \"shadow contagion\" is equally concerning from a systemic risk perspective. A single stablecoin de-peg cascaded through Morpho Blue, Euler, and Fluid via collateral chains, demonstrating that DeFi composability risk remains significantly underpriced.\n\nAI x Crypto was the standout sector, with AI tokens rallying 35\u201373% against an Extreme Fear backdrop. TAO's 73% surge reflects enterprise adoption of decentralized ML subnets. ZeroStack's $107M commitment to 0G represents TradFi institutional capital flowing into decentralized compute \u2014 a structural validation of the DePIN thesis. Render's NVIDIA Blackwell integration positions it as the go-to network for high-performance AI inference.\n\nPrediction markets hit $25.7B monthly volume in March. The CFTC suing 3 states over jurisdiction is unprecedented and will define the regulatory landscape for event trading. Polymarket's $9B valuation (ICE $2B investment) signals deep institutional conviction. The platform's expansion into sports (MLB $150M deal, LaLiga) diversifies beyond political/geopolitical markets.\n\nStablecoins: March transaction volume of $7.5T surpassing the US ACH network is a historic milestone \u2014 stablecoins have evolved from crypto-native tools to mainstream payment infrastructure. The USDC/USDT supply divergence (USDC +$2B vs USDT -$3B in Q1) reflects a market preference shift toward transparency and regulatory compliance.");

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
    catRow("4/6", "HYPE Token Unlock (~2.65%)", "HYPE", "Supply pressure; portfolio relevant", true),
    catRow("4/12", "APT Unlock (11.3M tokens)", "APT", "Supply pressure ~0.68%"),
    catRow("4/15-16", "Paris Blockchain Week", "Market", "MiCA discussion; institutional"),
    catRow("Mid-Apr", "CLARITY Act Committee Markup", "Stables", "Stablecoin yield framework", true),
    catRow("Apr", "91 ETF Applications Follow-up", "Multi", "SOL/XRP/LTC approval dynamics"),
    catRow("4/20", "LayerZero Unlock (25.7M ZRO)", "ZRO", "Supply pressure"),
    catRow("4/20", "Hong Kong Web3 Festival", "Market", "Asia\u2019s premier industry summit"),
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
    { text: "Extreme Fear (FGI=12). BTC range-bound $65\u201370K; alts broadly down. Drift $285M hack is 2026\u2019s largest security event. Regulatory framework accelerating (GENIUS Act, CLARITY Act, Schwab crypto).", options: { breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "Key watches: ", options: { bold: true, fontSize: 10, fontFace: "Arial", color: C.navy } },
    { text: "4/6 HYPE unlock (portfolio), CLARITY markup, AI sector momentum, Drift fallout trajectory.", options: { fontSize: 10, fontFace: "Arial", color: C.bodyText } },
  ], { x: 0.8, y: 4.0, w: 8.5, h: 0.85, margin: 0 });

  s6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s6.addText("Crypto Market Weekly | 2026.03.30\u201304.05  |  CONFIDENTIAL", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  s6.addNotes("Catalyst analysis for the next two weeks:\n\n1. HYPE Token Unlock (4/6) \u2014 ~2.65% of circulating supply. Directly relevant to our portfolio positioning. In the current Extreme Fear environment, unlock pressure could be amplified. However, Hyperliquid entering global top-10 derivatives venues and Grayscale's ETF S-1 filing provide positive narrative offsets. Monitor sell pressure vs. holder conviction.\n\n2. CLARITY Act Committee Markup (mid-April) \u2014 The final shape of the stablecoin yield framework will have deep implications for Circle, Coinbase, and the broader DeFi ecosystem. The current compromise (ban passive yields, allow activity-based rewards) is constructive, but the committee process could introduce amendments. Watch for any changes to the Tillis-Alsobrooks language.\n\n3. 91 ETF Applications Follow-up \u2014 The SEC's March 27 deadline for 91 crypto ETF filings (including SOL, XRP, LTC, DOGE spot ETFs) means approval/rejection dynamics will play out over the coming weeks. Any positive developments would be structurally transformative for altcoin markets.\n\nOverall assessment: The market sits at FGI=12 \u2014 deep Extreme Fear \u2014 but fundamental signals are improving rapidly. Regulatory clarity is accelerating (GENIUS Act, CLARITY Act, CFTC policy pivot, Schwab crypto access). TradFi institutional entry continues (Franklin Templeton, Schwab, ZeroStack). DeFi protocols are maturing (Aave V4 institutionalization). AI x Crypto is decoupling from broader market weakness.\n\nThe disconnect between deep pessimism in sentiment and accelerating fundamental improvement is historically a late-cycle fear signature. Maintain current positioning; the HYPE unlock on 4/6 and Drift hack fallout are the primary near-term risks. Consider opportunistic additions if sentiment remains depressed while regulatory catalysts deliver positively.");

  const outPath = "/Users/lx/Github/crypto-weekly-digest/reports/2026-03-30/slides-en.pptx";
  await pres.writeFile({ fileName: outPath });
  console.log("\u2705 Written to " + outPath);
}

generate().catch(e => { console.error(e); process.exit(1); });
