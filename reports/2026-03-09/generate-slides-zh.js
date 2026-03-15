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
  pres.title = "\u52A0\u5BC6\u5E02\u573A\u5468\u62A5 2026.03.09\u201303.15";

  // ============ SLIDE 1: Title ============
  const s1 = pres.addSlide();
  s1.background = { color: C.darkNavy };
  s1.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.6, w: 1.2, h: 0.04, fill: { color: C.gold } });
  s1.addText("\u52A0\u5BC6\u5E02\u573A\u5468\u62A5", {
    x: 0.8, y: 1.8, w: 8, h: 1.0,
    fontSize: 42, fontFace: "Arial", bold: true, color: C.white, margin: 0
  });
  s1.addText("2026\u5E743\u67089\u65E5 \u2014 3\u670815\u65E5", {
    x: 0.8, y: 2.7, w: 8, h: 0.5,
    fontSize: 20, fontFace: "Arial", color: C.gold, margin: 0
  });
  s1.addText("\u6570\u636E\u6765\u6E90\uFF1AMessari \u00B7 The Block \u00B7 CoinGecko \u00B7 CoinGlass \u00B7 DefiLlama", {
    x: 0.8, y: 3.4, w: 8, h: 0.4,
    fontSize: 11, fontFace: "Arial", color: C.medGray, margin: 0
  });
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.gold } });
  s1.addText("CONFIDENTIAL \u2014 FOR LP DISTRIBUTION ONLY", {
    x: 0, y: 5.25, w: 10, h: 0.375,
    fontSize: 9, fontFace: "Arial", color: C.darkNavy, align: "center", valign: "middle", bold: true
  });

  // ============ SLIDE 2: Market Overview ============
  const s2 = pres.addSlide();
  s2.background = { color: C.offWhite };
  s2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy } });
  s2.addText("\uD83D\uDCCA  \u5E02\u573A\u6982\u89C8", {
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
  s2.addText("\u6050\u60E7\u8D2A\u5A6A\u6307\u6570", { x: c3x + 0.25, y: cardY + 0.1, w: 2.4, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.medGray, bold: true, margin: 0 });
  s2.addText("15", { x: c3x + 0.25, y: cardY + 0.35, w: 1, h: 0.4, fontSize: 28, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
  s2.addText("\u6781\u5EA6\u6050\u60E7", { x: c3x + 1.1, y: cardY + 0.42, w: 1.5, h: 0.3, fontSize: 13, fontFace: "Arial", color: C.red, margin: 0 });

  // Price table
  const headerRow = [
    { text: "\u8D44\u4EA7", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 11, fontFace: "Arial", align: "center" } },
    { text: "\u4EF7\u683C", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 11, fontFace: "Arial", align: "center" } },
    { text: "\u5468\u6DA8\u8DCC\u5E45", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 11, fontFace: "Arial", align: "center" } },
    { text: "\u5907\u6CE8", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 11, fontFace: "Arial", align: "center" } },
  ];
  const makeRow = (asset, price, chg, note, isNeg) => [
    { text: asset, options: { fontSize: 11, fontFace: "Arial", bold: true, align: "center", color: C.bodyText } },
    { text: price, options: { fontSize: 11, fontFace: "Arial", align: "right", color: C.bodyText } },
    { text: chg, options: { fontSize: 11, fontFace: "Arial", align: "center", color: isNeg ? C.red : C.green, bold: true } },
    { text: note, options: { fontSize: 10, fontFace: "Arial", color: C.medGray } },
  ];
  s2.addTable([
    headerRow,
    makeRow("SOL", "$87.61", "+6.8%", "Alpenglow \u5347\u7EA7\u5229\u597D", false),
    makeRow("HYPE", "$37.13", "+21.7%", "\u8868\u73B0\u6700\u5F3A\uFF0CHyperEVM\u751F\u6001", false),
    makeRow("PENDLE", "$1.27", "+7.4%", "\u8DDF\u968F\u5927\u76D8\u56DE\u6696", false),
    makeRow("ENA", "$0.11", "+9.2%", "", false),
    makeRow("DYDX", "$0.09", "+13.4%", "\u8D85\u8DCC\u53CD\u5F39", false),
  ], {
    x: 0.6, y: 2.25, w: 8.8, colW: [1.3, 1.8, 1.5, 4.2],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.35, 0.33, 0.33, 0.33, 0.33, 0.33],
    autoPage: false,
  });

  s2.addText("\u5B8F\u89C2\u9A71\u52A8\uFF1A\u7F8E\u4F0A\u5730\u7F18\u51B2\u7A81\u89E6\u53D1\u95EA\u5D29\uFF08BTC\u89E6\u53CA$62,400\uFF09\u2192 \u673A\u6784\u8D44\u91D1\u6D41\u5165+\u76D1\u7BA1\u5229\u597D\u53CC\u91CD\u9A71\u52A8\u4FEE\u590D | \u603B\u5E02\u503C $2.51T | BTC\u4E3B\u5BFC\u7387 57.0%", {
    x: 0.6, y: 4.55, w: 8.8, h: 0.4,
    fontSize: 10, fontFace: "Arial", color: C.medGray, italic: true, margin: 0
  });

  s2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s2.addText("\u52A0\u5BC6\u5E02\u573A\u5468\u62A5 | 2026.03.09\u201303.15", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  // ============ SLIDE 3: Top Stories + Industry ============
  const s3 = pres.addSlide();
  s3.background = { color: C.offWhite };
  s3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy } });
  s3.addText("\uD83D\uDD25  \u672C\u5468\u8981\u95FB & \u884C\u4E1A\u52A8\u6001", {
    x: 0.6, y: 0, w: 9, h: 0.7,
    fontSize: 20, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0
  });

  // Left: Top 5
  s3.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 4.1, fill: { color: C.white }, shadow: makeShadow() });
  s3.addText("\u672C\u5468\u8981\u95FB TOP 5", {
    x: 0.6, y: 0.95, w: 4, h: 0.35,
    fontSize: 13, fontFace: "Arial", bold: true, color: C.navy, margin: 0
  });

  const stories = [
    ["1", "SEC\u4E0ECFTC\u7B7E\u7F72\u8054\u5408\u76D1\u7BA1MOU", "\u7ED3\u675F\u591A\u5E74\u7BA1\u8F96\u6743\u4E4B\u4E89\uFF0C\u5F00\u542F\u7EDF\u4E00\u76D1\u7BA1\u65F6\u4EE3"],
    ["2", "Ripple $7.5\u4EBF\u56DE\u8D2D\uFF0C\u4F30\u503C$500\u4EBF", "\u79C1\u8425\u516C\u53F8\u8D22\u5BCC500\u5F3A\u7EA7\u522B\u8D44\u672C\u8FD0\u8425"],
    ["3", "\u5730\u7F18\u653F\u6CBB\u95EA\u5D29\u4E0E\u4FEE\u590D", "\u7F8E\u4F0A\u51B2\u7A81\u89E6\u53D1BTC\u6025\u8DCC\u81F3$62,400\u540E\u8FC5\u901F\u53CD\u5F39"],
    ["4", "Solana Alpenglow\u5347\u7EA7\u83B7\u6279", "\u4EA4\u6613\u6700\u7EC8\u6027\u4ECE12.8\u79D2\u7F29\u77ED\u81F3100-150\u6BEB\u79D2"],
    ["5", "SEC\u64A4\u9500BitClout/DeSo\u8BC9\u8BBC", "\u76D1\u7BA1\u7B56\u7565\u8F6C\u5411\u5408\u89C4\u6307\u5BFC"],
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
    "Solana Alpenglow: 98.27%\u8D28\u62BC\u8005\u6279\u51C6\uFF0C\u6700\u7EC8\u6027\u2192100ms",
    "Avalanche TVL\u7A81\u7834$2.1B\uFF0CQ1\u7FFB\u500D\u589E\u957F",
    "Monad\u4E3B\u7F51\u5B9E\u6D4B\u6301\u7EED10,000 TPS",
    "Base\u52A0\u901F\u81EA\u7814\u67B6\u6784\u8F6C\u578B\uFF0C\u8131\u79BBOP Stack",
  ];
  s3.addText(l1Items.map((t, i) => ({
    text: t,
    options: { bullet: true, breakLine: i < l1Items.length - 1, fontSize: 10, fontFace: "Arial", color: C.bodyText }
  })), { x: 5.3, y: 1.32, w: 4.1, h: 1.45, margin: 0, paraSpaceAfter: 4 });

  // Right bottom: DeFi
  s3.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 3.05, w: 4.5, h: 1.95, fill: { color: C.white }, shadow: makeShadow() });
  s3.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 3.05, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s3.addText("DeFi & \u6C38\u7EED\u5408\u7EA6", { x: 5.3, y: 3.05, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  const defiItems = [
    "Aave V4 \u201CHub and Spoke\u201D\u8DE8\u94FE\u7EDF\u4E00\u6D41\u52A8\u6027\u67B6\u6784",
    "BlackRock\u8D28\u62BCETH ETF (ETHB)\u4E0A\u7EBFNasdaq",
    "Hyperliquid OI\u5CF0\u503C$6.48B\uFF0C\u8FDB\u5165\u5168\u7403Top10",
    "\u6C38\u7EEDEX\u5E02\u573A\u4EFD\u989D\u8FBE10.2%\uFF0C\u7ED3\u6784\u6027\u63D0\u5347\u786E\u8BA4",
  ];
  s3.addText(defiItems.map((t, i) => ({
    text: t,
    options: { bullet: true, breakLine: i < defiItems.length - 1, fontSize: 10, fontFace: "Arial", color: C.bodyText }
  })), { x: 5.3, y: 3.47, w: 4.1, h: 1.45, margin: 0, paraSpaceAfter: 4 });

  s3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s3.addText("\u52A0\u5BC6\u5E02\u573A\u5468\u62A5 | 2026.03.09\u201303.15", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  // ============ SLIDE 4: Regulation + Fundraising ============
  const s4 = pres.addSlide();
  s4.background = { color: C.offWhite };
  s4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy } });
  s4.addText("\u2696\uFE0F  \u76D1\u7BA1\u653F\u7B56 & \u6295\u878D\u8D44", {
    x: 0.6, y: 0, w: 9, h: 0.7,
    fontSize: 20, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0
  });

  // Left: Regulation
  s4.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 4.1, fill: { color: C.white }, shadow: makeShadow() });
  s4.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 0.06, h: 4.1, fill: { color: C.gold } });
  s4.addText("\u76D1\u7BA1\u4E0E\u653F\u7B56", { x: 0.7, y: 0.95, w: 4, h: 0.35, fontSize: 14, fontFace: "Arial", bold: true, color: C.navy, margin: 0 });

  const regItems = [
    { title: "SEC\u4E0ECFTC\u8054\u5408\u76D1\u7BA1MOU", desc: "\u4FE1\u606F\u5171\u4EAB\u3001\u8054\u5408\u5BA1\u67E5\u3001\u7EDF\u4E00\u5B9A\u4E49\u6846\u67B6" },
    { title: "SEC\u4EE3\u5E01\u5206\u7C7B\u6307\u5F15", desc: "\u63D0\u4EA4\u767D\u5BAB\uFF0C\u4ECE\u8BC9\u8BBC\u6267\u6CD5\u8F6C\u5411\u89C4\u5219\u5236\u5B9A" },
    { title: "CFTC\u542F\u52A8Project Crypto", desc: "\u4E0ESEC\u5408\u4F5C\u5EFA\u7ACB\u7EDF\u4E00\u52A0\u5BC6\u8D44\u4EA7\u5206\u7C7B" },
    { title: "\u8D22\u653F\u90E8DeFi\u98CE\u9669\u62A5\u544A", desc: "\u5EFA\u8BAE\u201C\u6682\u6263\u6CD5\u201D\uFF0C\u660E\u786EDeFi\u53C2\u4E0E\u8005AML\u8D23\u4EFB" },
    { title: "\u7A33\u5B9A\u5E01\u6536\u76CA\u7981\u4EE4\u8FA9\u8BBA", desc: "\u53C2\u8BAE\u9662\u5BA1\u8BAE130+\u4FEE\u6B63\u6848\uFF0C\u62DF\u7981\u6563\u6237\u6536\u76CA\u578B\u7A33\u5B9A\u5E01" },
  ];

  regItems.forEach((item, i) => {
    const iy = 1.45 + i * 0.65;
    s4.addText(item.title, { x: 0.7, y: iy, w: 4, h: 0.25, fontSize: 11, fontFace: "Arial", bold: true, color: C.bodyText, margin: 0 });
    s4.addText(item.desc, { x: 0.7, y: iy + 0.23, w: 4, h: 0.25, fontSize: 10, fontFace: "Arial", color: C.medGray, margin: 0 });
  });

  // Right: Fundraising
  s4.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 4.5, h: 4.1, fill: { color: C.white }, shadow: makeShadow() });
  s4.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 0.06, h: 4.1, fill: { color: C.gold } });
  s4.addText("\u6295\u878D\u8D44\u4E0E\u5E76\u8D2D", { x: 5.4, y: 0.95, w: 4, h: 0.35, fontSize: 14, fontFace: "Arial", bold: true, color: C.navy, margin: 0 });

  const fundHeader = [
    { text: "\u9879\u76EE", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial" } },
    { text: "\u91D1\u989D", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial", align: "right" } },
    { text: "\u5907\u6CE8", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial" } },
  ];
  const fundRow = (name, amt, note) => [
    { text: name, options: { fontSize: 10, fontFace: "Arial", color: C.bodyText, bold: true } },
    { text: amt, options: { fontSize: 10, fontFace: "Arial", color: C.bodyText, align: "right" } },
    { text: note, options: { fontSize: 9, fontFace: "Arial", color: C.medGray } },
  ];
  s4.addTable([
    fundHeader,
    fundRow("Ripple\u56DE\u8D2D", "$750M", "\u4F30\u503C$50B"),
    fundRow("ASTER\u89E3\u9501", "~$56M", "3/17\u91CA\u653E7,814\u4E07\u679A"),
    fundRow("LayerZero ZRO", "~$45M", "3/20\u89E3\u9501\uFF0C\u5360\u5E02\u503C2.47%"),
    fundRow("ETH Foundation", "$10M", "OTC\u51FA\u552E5,000 ETH"),
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
    { text: "\u5173\u6CE8\uFF1A", options: { bold: true, fontSize: 10, fontFace: "Arial", color: C.navy } },
    { text: "\u4E0A\u5B63\u5EA6\u878D\u8D44\u989D\u4E0B\u964D62%\uFF0C\u4F46\u5355\u7B14\u5E73\u5747\u98D9\u5347272%\u81F3$3,400\u4E07\u3002\u884C\u4E1AM&A\u9884\u8BA12026\u5E74\u7A81\u7834$370\u4EBF", options: { fontSize: 10, fontFace: "Arial", color: C.bodyText } }
  ], { x: 5.55, y: 3.15, w: 3.7, h: 0.75, margin: 0, valign: "middle" });

  s4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s4.addText("\u52A0\u5BC6\u5E02\u573A\u5468\u62A5 | 2026.03.09\u201303.15", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  // ============ SLIDE 5: Security + AI + Prediction ============
  const s5 = pres.addSlide();
  s5.background = { color: C.offWhite };
  s5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy } });
  s5.addText("\uD83D\uDD13  \u5B89\u5168\u4E8B\u4EF6 & \u70ED\u70B9\u8D5B\u9053", {
    x: 0.6, y: 0, w: 9, h: 0.7,
    fontSize: 20, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0
  });

  // Security - left
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 2.35, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 0.35, fill: { color: C.red } });
  s5.addText("\u5B89\u5168\u4E8B\u4EF6  |  \u5468\u635F\u5931: ~$81M", { x: 0.6, y: 0.9, w: 4.2, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  s5.addTable([
    [
      { text: "\u4E8B\u4EF6", options: { bold: true, fontSize: 9, fontFace: "Arial", fill: { color: C.lightGray }, color: C.bodyText } },
      { text: "\u635F\u5931", options: { bold: true, fontSize: 9, fontFace: "Arial", fill: { color: C.lightGray }, color: C.bodyText, align: "right" } },
      { text: "\u65B9\u5F0F", options: { bold: true, fontSize: 9, fontFace: "Arial", fill: { color: C.lightGray }, color: C.bodyText } },
    ],
    [{ text: "Aave\u6ED1\u70B9\u4EA4\u6613", options: { fontSize: 9, fontFace: "Arial" } }, { text: "~$50M", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "\u6781\u7AEF\u6ED1\u70B9/MEV", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
    [{ text: "Aave\u9884\u8A00\u673A\u6545\u969C", options: { fontSize: 9, fontFace: "Arial" } }, { text: "$27.78M", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "\u914D\u7F6E\u9519\u8BEF(\u5DF2\u8D54\u507F)", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
    [{ text: "Solv Protocol", options: { fontSize: 9, fontFace: "Arial" } }, { text: "$2.73M", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "\u672A\u5BA1\u8BA1\u5408\u7EA6", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
    [{ text: "Bonk.fun DNS\u52AB\u6301", options: { fontSize: 9, fontFace: "Arial" } }, { text: "\u672A\u62AB\u9732", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "DNS+\u9493\u9C7C\u811A\u672C", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
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
    { text: "Messari\u5B9A\u4E492026\u5E74\u4E3AAI\u4E0E\u533A\u5757\u94FE\u201C\u7CFB\u7EDF\u7EA7\u6574\u5408\u201D\u5143\u5E74", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "TRON\u52A0\u5165Agentic AI Foundation\uFF0C\u5927\u578BL1\u9996\u6B21\u5E03\u5C40", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "ElizaOS\u6210\u4E3AAI Agent\u4EBA\u683C\u6807\u51C6\u6846\u67B6", options: { bullet: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
  ], { x: 5.3, y: 1.35, w: 4.1, h: 1.3, margin: 0, paraSpaceAfter: 4 });

  // Prediction Markets - right bottom
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 2.95, w: 4.5, h: 1.85, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 2.95, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s5.addText("\u9884\u6D4B\u5E02\u573A", { x: 5.3, y: 2.95, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  s5.addText([
    { text: "Polymarket & Kalshi\u7784\u51C6$200\u4EBF\u4F30\u503C", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "\u8D85\u77ED\u671F\u5408\u7EA6\u65E5\u4EA4\u6613\u91CF$7,000\u4E07\uFF0CVitalik\u6279\u8BC4\u8FC7\u5EA6\u6295\u673A", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "\u76D1\u7BA1\u53CD\u5F39\uFF1A\u5185\u5E55\u4EA4\u6613\u7ACB\u6CD5+\u72B9\u4ED6\u5DDE\u62DF\u91CD\u65B0\u5206\u7C7B\u4E3A\u8D4C\u535A", options: { bullet: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
  ], { x: 5.3, y: 3.4, w: 4.1, h: 1.3, margin: 0, paraSpaceAfter: 4 });

  // Stablecoins
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.5, w: 4.5, h: 1.5, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.5, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s5.addText("\u7A33\u5B9A\u5E01", { x: 0.6, y: 3.5, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });
  s5.addText([
    { text: "\u6536\u76CA\u578B\u7A33\u5B9A\u5E01\u5E02\u503C$22.7B\uFF0C\u589E\u901F\u8D85\u5E02\u573A15\u500D", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "Visa\u6269\u5C55USDC\u7ED3\u7B97\uFF0CHashdex\u9884\u6D4B\u5E74\u5E95\u7FFB\u500D\u81F3$6000\u4EBF", options: { bullet: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
  ], { x: 0.6, y: 3.95, w: 4.1, h: 0.9, margin: 0, paraSpaceAfter: 4 });

  s5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s5.addText("\u52A0\u5BC6\u5E02\u573A\u5468\u62A5 | 2026.03.09\u201303.15", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  // ============ SLIDE 6: Catalysts + Outlook ============
  const s6 = pres.addSlide();
  s6.background = { color: C.offWhite };
  s6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy } });
  s6.addText("\uD83D\uDD2E  \u672A\u6765\u50AC\u5316\u5242 & \u5C55\u671B", {
    x: 0.6, y: 0, w: 9, h: 0.7,
    fontSize: 20, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0
  });

  const catHeader = [
    { text: "\u65E5\u671F", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial", align: "center" } },
    { text: "\u4E8B\u4EF6", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial" } },
    { text: "\u5F71\u54CD\u8D44\u4EA7", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial", align: "center" } },
    { text: "\u9884\u671F\u5F71\u54CD", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial" } },
  ];
  const catRow = (date, event, asset, impact, highlight = false) => [
    { text: date, options: { fontSize: 10, fontFace: "Arial", align: "center", color: C.bodyText, fill: highlight ? { color: "FFF3E0" } : undefined } },
    { text: event, options: { fontSize: 10, fontFace: "Arial", color: C.bodyText, bold: highlight, fill: highlight ? { color: "FFF3E0" } : undefined } },
    { text: asset, options: { fontSize: 10, fontFace: "Arial", align: "center", color: C.navy, bold: true, fill: highlight ? { color: "FFF3E0" } : undefined } },
    { text: impact, options: { fontSize: 9, fontFace: "Arial", color: C.medGray, fill: highlight ? { color: "FFF3E0" } : undefined } },
  ];
  s6.addTable([
    catHeader,
    catRow("3/16", "Neutron\u5347\u7EA7", "NTRN", "\u8DE8\u94FE\u80FD\u529B\u589E\u5F3A"),
    catRow("3/17", "ASTER Token Unlock ($56M)", "ASTER", "\u4F9B\u7ED9\u538B\u529B"),
    catRow("3/18", "FOMC\u5229\u7387\u51B3\u8BAE", "\u5168\u5E02\u573A", "\u5173\u952E\u5B8F\u89C2\u4E8B\u4EF6", true),
    catRow("3/20", "LayerZero ZRO\u89E3\u9501 ($45M)", "ZRO", "\u4F9B\u7ED9\u538B\u529B\uFF0C\u5360\u5E02\u503C2.47%"),
    catRow("3/24-26", "Digital Asset Summit (\u7EBD\u7EA6)", "\u673A\u6784\u8D44\u4EA7", "TradFi\u6574\u5408+AI\u878D\u5408"),
    catRow("3/27", "\u52A0\u5BC6ETP\u6700\u7EC8\u622A\u6B62\u65E5", "\u591A\u79CD\u4EE3\u5E01", "\u6279\u51C6\u5C06\u5F00\u542F\u65B0\u673A\u6784\u8D44\u91D1"),
    catRow("3\u6708\u5E95", "Midnight\u4E3B\u7F51\u4E0A\u7EBF", "ADA", "Cardano\u9690\u79C1\u4FA7\u94FE"),
  ], {
    x: 0.5, y: 0.9, w: 9, colW: [0.9, 3.2, 1.2, 3.7],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.32, 0.32, 0.32, 0.32, 0.32, 0.32, 0.32, 0.32],
    autoPage: false,
  });

  // Outlook
  s6.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.65, w: 9, h: 1.3, fill: { color: C.white }, shadow: makeShadow() });
  s6.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.65, w: 0.06, h: 1.3, fill: { color: C.gold } });
  s6.addText("\u672C\u5468\u5C0F\u7ED3", { x: 0.8, y: 3.7, w: 8, h: 0.3, fontSize: 13, fontFace: "Arial", bold: true, color: C.navy, margin: 0 });
  s6.addText([
    { text: "\u5730\u7F18\u51B2\u7A81\u89E6\u53D1\u95EA\u5D29\u540EBTC\u5F3A\u52BF\u53CD\u5F39+6.3%\uFF0CHYPE\u9886\u6DA8+21.7%\u3002SEC/CFTC\u8054\u5408\u76D1\u7BA1MOU\u4E3A\u884C\u4E1A\u5E26\u6765\u91CD\u5927\u6B63\u9762\u4FE1\u53F7\uFF0C\u4F46FGI\u4ECD\u5904\u201C\u6781\u5EA6\u6050\u60E7\u201D\uFF0C\u663E\u793A\u6563\u6237\u4E0E\u673A\u6784\u4E25\u91CD\u5206\u5316\u3002", options: { breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "\u91CD\u70B9\u5173\u6CE8\uFF1A", options: { bold: true, fontSize: 10, fontFace: "Arial", color: C.navy } },
    { text: " 3/18 FOMC\u5229\u7387\u51B3\u8BAE\uFF08\u6301\u4ED3\u76F8\u5173\uFF09\u3001ASTER/ZRO\u89E3\u9501\u538B\u529B\u3001Aave\u9884\u8A00\u673A\u6545\u969C\u540E\u884C\u4E1A\u5B89\u5168\u673A\u5236\u53CD\u601D\u3002", options: { fontSize: 10, fontFace: "Arial", color: C.bodyText } },
  ], { x: 0.8, y: 4.0, w: 8.5, h: 0.85, margin: 0 });

  s6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s6.addText("\u52A0\u5BC6\u5E02\u573A\u5468\u62A5 | 2026.03.09\u201303.15  |  CONFIDENTIAL", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  const outPath = "/Users/lx/Github/crypto-weekly-digest/reports/2026-03-09/slides-zh.pptx";
  await pres.writeFile({ fileName: outPath });
  console.log("\u2705 Written to " + outPath);
}

generate().catch(e => { console.error(e); process.exit(1); });
