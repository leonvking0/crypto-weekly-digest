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
  s.addText(`\u52A0\u5BC6\u5E02\u573A\u5468\u62A5 | ${WEEK}`, { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });
};

async function generate() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "42";
  pres.title = `\u52A0\u5BC6\u5E02\u573A\u5468\u62A5 ${WEEK}`;

  // ===== SLIDE 1: Title =====
  const s1 = pres.addSlide();
  s1.background = { color: C.darkNavy };
  s1.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.6, w: 1.2, h: 0.04, fill: { color: C.gold } });
  s1.addText("\u52A0\u5BC6\u5E02\u573A\u5468\u62A5", { x: 0.8, y: 1.8, w: 8, h: 1.0, fontSize: 42, fontFace: "Arial", bold: true, color: C.white, margin: 0 });
  s1.addText("2026\u5E743\u670816\u65E5 \u2014 3\u670822\u65E5", { x: 0.8, y: 2.7, w: 8, h: 0.5, fontSize: 20, fontFace: "Arial", color: C.gold, margin: 0 });
  s1.addText("\u6570\u636E\u6765\u6E90\uFF1AMessari \u00B7 The Block \u00B7 CoinGecko \u00B7 DefiLlama", { x: 0.8, y: 3.4, w: 8, h: 0.4, fontSize: 11, fontFace: "Arial", color: C.medGray, margin: 0 });
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.gold } });
  s1.addText("CONFIDENTIAL \u2014 FOR LP DISTRIBUTION ONLY", { x: 0, y: 5.25, w: 10, h: 0.375, fontSize: 9, fontFace: "Arial", color: C.darkNavy, align: "center", valign: "middle", bold: true });
  s1.addNotes("\u672C\u5468\u52A0\u5BC6\u5E02\u573A\u7ECF\u5386\u4E86\u201C\u8FC7\u5C71\u8F66\u201D\u884C\u60C5\u3002\u5468\u521D\u56E0SEC/CFTC\u8054\u5408\u76D1\u7BA1\u6307\u5F15\u53CA\u673A\u6784\u4E70\u76D8\u63A8\u52A8BTC\u53CD\u5F39\u81F3$76K\u516D\u5468\u65B0\u9AD8\uFF0C\u4F46\u5468\u516D\u56E0\u4E2D\u4E1C\u5730\u7F18\u653F\u6CBB\u5371\u673A\uFF08\u970D\u5C14\u6728\u5179\u6D77\u5CE1\u5C40\u52BF\u5347\u7EA7\u3001\u539F\u6CB9\u7A81\u7834$110\uFF09\u89E6\u53D1\u95EA\u5D29\uFF0CBTC\u56DE\u843D\u81F3$68.8K\u3002\n\n\u5173\u952E\u80CC\u666F\uFF1A\n1. \u7F8E\u8054\u50A8\u7EF4\u6301\u5229\u7387\u4E0D\u53D8\uFF083.50-3.75%\uFF09\uFF0CCPI\u6301\u7EED2.4%\uFF0C\u964D\u606F\u9884\u671F\u843D\u7A7A\n2. \u6050\u60E7\u8D2A\u5A6A\u6307\u6570\u8DCC\u81F310\uFF08\u6781\u5EA6\u6050\u60E7\uFF09\uFF0C\u4E3A2024\u5E74\u4EE5\u6765\u6700\u4F4E\u6C34\u5E73\n3. \u672C\u5468\u6E05\u7B97\u8D85$10\u4EBF\uFF0C85%\u4E3A\u591A\u5934\u4ED3\u4F4D\n4. \u5C3D\u7BA1\u77ED\u671F\u6CE2\u52A8\u5267\u70C8\uFF0C\u76D1\u7BA1\u7AEF\u7684\u91CD\u5927\u7A81\u7834\uFF08SEC/CFTC\u8054\u5408\u6307\u5F15\u660E\u786E18\u79CD\u4EE3\u5E01\u4E3A\u6570\u5B57\u5546\u54C1\uFF09\u4E3A\u4E2D\u957F\u671F\u5960\u5B9A\u4E86\u6781\u4E3A\u79EF\u6781\u7684\u57FA\u7840");

  // ===== SLIDE 2: Market Overview =====
  const s2 = pres.addSlide();
  s2.background = { color: C.offWhite };
  s2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy } });
  s2.addText("\uD83D\uDCCA  \u5E02\u573A\u6982\u89C8", { x: 0.6, y: 0, w: 9, h: 0.7, fontSize: 20, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  // 3 metric cards
  const cardY = 0.95, cardH = 1.05, cardW = 2.8, cardGap = 0.3, cardStartX = 0.6;
  // BTC
  s2.addShape(pres.shapes.RECTANGLE, { x: cardStartX, y: cardY, w: cardW, h: cardH, fill: { color: C.white }, shadow: makeShadow() });
  s2.addShape(pres.shapes.RECTANGLE, { x: cardStartX, y: cardY, w: 0.06, h: cardH, fill: { color: C.gold } });
  s2.addText("BTC", { x: cardStartX + 0.25, y: cardY + 0.1, w: 2, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.medGray, bold: true, margin: 0 });
  s2.addText("$68,811", { x: cardStartX + 0.25, y: cardY + 0.35, w: 2, h: 0.4, fontSize: 28, fontFace: "Arial", color: C.bodyText, bold: true, margin: 0 });
  s2.addText("-3.9%", { x: cardStartX + 0.25, y: cardY + 0.72, w: 2, h: 0.25, fontSize: 14, fontFace: "Arial", color: C.red, bold: true, margin: 0 });

  // ETH
  const c2x = cardStartX + cardW + cardGap;
  s2.addShape(pres.shapes.RECTANGLE, { x: c2x, y: cardY, w: cardW, h: cardH, fill: { color: C.white }, shadow: makeShadow() });
  s2.addShape(pres.shapes.RECTANGLE, { x: c2x, y: cardY, w: 0.06, h: cardH, fill: { color: C.gold } });
  s2.addText("ETH", { x: c2x + 0.25, y: cardY + 0.1, w: 2, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.medGray, bold: true, margin: 0 });
  s2.addText("$2,082", { x: c2x + 0.25, y: cardY + 0.35, w: 2, h: 0.4, fontSize: 28, fontFace: "Arial", color: C.bodyText, bold: true, margin: 0 });
  s2.addText("-0.6%", { x: c2x + 0.25, y: cardY + 0.72, w: 2, h: 0.25, fontSize: 14, fontFace: "Arial", color: C.red, bold: true, margin: 0 });

  // Fear & Greed
  const c3x = c2x + cardW + cardGap;
  s2.addShape(pres.shapes.RECTANGLE, { x: c3x, y: cardY, w: cardW, h: cardH, fill: { color: C.white }, shadow: makeShadow() });
  s2.addShape(pres.shapes.RECTANGLE, { x: c3x, y: cardY, w: 0.06, h: cardH, fill: { color: C.red } });
  s2.addText("\u6050\u60E7\u8D2A\u5A6A\u6307\u6570", { x: c3x + 0.25, y: cardY + 0.1, w: 2.4, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.medGray, bold: true, margin: 0 });
  s2.addText("10", { x: c3x + 0.25, y: cardY + 0.35, w: 1, h: 0.4, fontSize: 28, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
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
    makeRow("SOL", "$87.49", "-0.5%", "SOL ETF AUM\u7A01$10\u4EBF", true),
    makeRow("HYPE", "$38.00", "+1.5%", "Grayscale\u63D0\u4EA4HYPE ETF\u7533\u8BF7", false),
    makeRow("PENDLE", "$1.19", "-6.4%", "TVL\u8FBE$40\u4EBF\u91CC\u7A0B\u7891", true),
    makeRow("ENA", "$0.10", "-9.7%", "sUSDe\u52A8\u6001\u89E3\u8D28\u62BC\u671F\u8C03\u6574", true),
    makeRow("DYDX", "$0.08", "-10.0%", "\u5173\u505C16\u4E2A\u4F4E\u6D3B\u8DC3\u5E02\u573A", true),
  ], {
    x: 0.6, y: 2.25, w: 8.8, colW: [1.3, 1.8, 1.5, 4.2],
    border: { pt: 0.5, color: C.lightGray }, rowH: [0.35, 0.33, 0.33, 0.33, 0.33, 0.33], autoPage: false,
  });

  s2.addText("\u5B8F\u89C2\u9A71\u52A8\uFF1A\u7F8E\u8054\u50A8\u7EF4\u6301\u5229\u7387\u4E0D\u53D8(3.50-3.75%)\uFF0C\u4E2D\u4E1C\u5730\u7F18\u5371\u673A\u5347\u7EA7\u539F\u6CB9\u7A81\u7834$110\u89E6\u53D1\u5468\u516D\u95EA\u5D29\uFF1BBTC\u4ECE$76K\u66B4\u8DCC\u81F3$68.8K\uFF0C24h\u6E05\u7B97$10\u4EBF+", {
    x: 0.6, y: 4.55, w: 8.8, h: 0.4, fontSize: 10, fontFace: "Arial", color: C.medGray, italic: true, margin: 0
  });
  footer(s2, pres);
  s2.addNotes("\u5E02\u573A\u6982\u89C8\u6DF1\u5EA6\u5206\u6790\uFF1A\n\n1. BTC\u672C\u5468\u7ECF\u5386\u4E86\u5178\u578B\u7684\u201C\u6D17\u76D8\u201D\u884C\u60C5\u3002\u5468\u521D\u56E0\u76D1\u7BA1\u5229\u597D\u63A8\u52A8\u81F3$76K\u516D\u5468\u65B0\u9AD8\uFF0C\u4F46\u673A\u6784\u4E70\u76D8\u4EE52.8\u500D\u65B0\u4F9B\u5E94\u901F\u5EA6\u5438\u6536\u7684\u540C\u65F6\uFF0CETF\u5728\u5468\u4E09\u5374\u51FA\u73B0$2.19\u4EBF\u51C0\u6D41\u51FA\uFF0C\u663E\u793A\u673A\u6784\u77ED\u671F\u5206\u6B67\u3002\n\n2. ETH\u8868\u73B0\u76F8\u5BF9\u6297\u8DCC\uFF08-0.6%\uFF09\uFF0C\u4E3B\u8981\u56E0\u4E3A\u201CGlamsterdam\u201D\u5347\u7EA7\u9884\u671F\u63D0\u4F9B\u652F\u6491\uFF0C\u4EE5\u53CASEC\u660E\u786E\u5C06ETH\u5F52\u7C7B\u4E3A\u201C\u6570\u5B57\u5546\u54C1\u201D\u800C\u975E\u8BC1\u5238\u3002\n\n3. \u6050\u60E7\u8D2A\u5A6A\u6307\u6570\u8DCC\u81F310\uFF08\u6781\u5EA6\u6050\u60E7\uFF09\u662F2024\u5E74\u4EE5\u6765\u6700\u4F4E\u6C34\u5E73\uFF0C\u5386\u53F2\u4E0A\u6781\u5EA6\u6050\u60E7\u533A\u95F4\u5F80\u5F80\u662F\u4E2D\u957F\u671F\u8D2D\u4E70\u673A\u4F1A\u3002\n\n4. \u672C\u5468$438M\u4EE3\u5E01\u89E3\u9501\u52A0\u5267\u4E86\u5356\u538B\uFF0C\u5305\u62ECASTER($55.9M)\u3001LayerZero($50.3M)\u3001Lombard($34.8M)\u7B49\u3002\n\n5. \u4F46\u4E5F\u6709\u79EF\u6781\u4FE1\u53F7\uFF1AMetaplanet\u52DF\u96C6$2.55\u4EBF\u7528\u4E8E\u589E\u6301BTC\uFF0CBullish\u62A5\u544A\u6301\u6709\u91CF\u8FBE24,400 BTC\uFF08~$16.7\u4EBF\uFF09\uFF0C\u4F01\u4E1A\u8D2D\u4E70\u901F\u5EA6\u8FBE\u65B0\u4F9B\u5E94\u76843\u500D\u3002");

  // ===== SLIDE 3: Top Stories + Industry =====
  const s3 = pres.addSlide();
  s3.background = { color: C.offWhite };
  s3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy } });
  s3.addText("\uD83D\uDD25  \u672C\u5468\u8981\u95FB & \u884C\u4E1A\u52A8\u6001", { x: 0.6, y: 0, w: 9, h: 0.7, fontSize: 20, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  // Left: Top 5
  s3.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 4.1, fill: { color: C.white }, shadow: makeShadow() });
  s3.addText("\u672C\u5468\u8981\u95FB TOP 5", { x: 0.6, y: 0.95, w: 4, h: 0.35, fontSize: 13, fontFace: "Arial", bold: true, color: C.navy, margin: 0 });

  const stories = [
    ["1", "\u201C\u970D\u5C14\u6728\u5179\u6D77\u5CE1\u201D\u95EA\u5D29\u4E0E$10\u4EBF\u6E05\u7B97", "BTC\u4ECE$76K\u66B4\u8DCC\u81F3$68.2K\uFF0C\u539F\u6CB9$110\u89E6\u53D1\u6050\u6148"],
    ["2", "SEC/CFTC\u8054\u5408\u76D1\u7BA1\u6307\u5F15", "18\u79CD\u4EE3\u5E01\u660E\u786E\u4E3A\u201C\u6570\u5B57\u5546\u54C1\u201D\uFF0C\u7ED3\u675F\u76D1\u7BA1\u6A21\u7CCA\u5730\u5E26"],
    ["3", "\u4E0D\u4E39\u4E3B\u6743\u57FA\u91D1\u8F6C\u51FA$7200\u4E07BTC", "Arkham\u94FE\u4E0A\u6570\u636E\u663E\u793A\u4E3B\u6743\u7EA7\u6301\u4ED3\u53D8\u52A8"],
    ["4", "\u7B2C2000\u4E07\u679ABTC\u88AB\u6316\u51FA", "\u4EC5\u5269100\u4E07\u679A\u5F85\u6316\uFF0CBTC\u4E3B\u5BFC\u7387\u7A33\u5B8857%"],
    ["5", "\u4EE3\u5E01\u89E3\u9501\u521B2026\u5E74\u65B0\u9AD8", "\u672C\u6708$60\u4EBF+\uFF0CLayerZero $55.5M\u3001Lombard $32.4M"],
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
    "Aster Chain\u4E3B\u7F51\u521B\u4E16\u4E0A\u7EBF\uFF0CZK\u9690\u79C1L1",
    "Linea\u5347\u7EA7\u4E3AType-1 zkEVM\uFF0C100% EVM\u517C\u5BB9",
    "Base\u5360L2 TVL 46%\u3001\u8D39\u7528\u6536\u5165>60%",
    "Solana Alpenglow\u8FDB\u5165\u6700\u7EC8\u6D4B\u8BD5\uFF0CRWA\u7A81\u7834$17\u4EBF",
  ];
  s3.addText(l1Items.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < l1Items.length - 1, fontSize: 10, fontFace: "Arial", color: C.bodyText } })), { x: 5.3, y: 1.32, w: 4.1, h: 1.45, margin: 0, paraSpaceAfter: 4 });

  // Right bottom: DeFi & Perp DEX
  s3.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 3.05, w: 4.5, h: 1.95, fill: { color: C.white }, shadow: makeShadow() });
  s3.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 3.05, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s3.addText("DeFi & \u6C38\u7EED\u5408\u7EA6", { x: 5.3, y: 3.05, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });
  const defiItems = [
    "Pendle TVL $40\u4EBF\u3001\u7D2F\u8BA1\u4EA4\u6613\u91CF$100\u4EBF",
    "DEX\u5360\u5168\u7403\u6C38\u7EED\u5408\u7EA6\u5E02\u573A\u4EFD\u989D\u8FBE20%",
    "Hyperliquid\u539F\u6CB9\u5355\u65E5$17.7\u4EBF\uFF0CS&P 500\u5B98\u65B9\u6388\u6743",
    "Aave v4 + Lido v3 \u91CD\u5927\u5347\u7EA7\u4E0A\u7EBF",
  ];
  s3.addText(defiItems.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < defiItems.length - 1, fontSize: 10, fontFace: "Arial", color: C.bodyText } })), { x: 5.3, y: 3.47, w: 4.1, h: 1.45, margin: 0, paraSpaceAfter: 4 });
  footer(s3, pres);
  s3.addNotes("\u672C\u5468\u8981\u95FB\u6DF1\u5EA6\u5206\u6790\uFF1A\n\n1. \u970D\u5C14\u6728\u5179\u6D77\u5CE1\u95EA\u5D29\uFF1A\u8FD9\u662F\u5178\u578B\u7684\u5730\u7F18\u653F\u6CBB\u9ED1\u5929\u9E45\u4E8B\u4EF6\u3002\u7279\u6717\u666E\u5BF9\u970D\u5C14\u6728\u5179\u6D77\u5CE1\u53D1\u51FA\u6700\u540E\u901A\u789D\uFF0C\u539F\u6CB9\u5347\u81F3$110/\u6876\uFF0C\u89E6\u53D1\u5168\u7403\u98CE\u9669\u8D44\u4EA7\u629B\u552E\u3002BTC\u4ECE$76K\u66B4\u8DCC\u81F3$68.2K\uFF0C24\u5C0F\u65F6\u5185\u6E05\u7B97\u8D85$10\u4EBF\uFF0C85%\u4E3A\u591A\u5934\u3002\u8FD9\u663E\u793A\u52A0\u5BC6\u5E02\u573A\u5BF9\u5730\u7F18\u653F\u6CBB\u98CE\u9669\u7684\u654F\u611F\u5EA6\u4F9D\u7136\u5F88\u9AD8\uFF0C\u4F46\u4E5F\u53EF\u80FD\u63D0\u4F9B\u4E86\u201C\u6050\u6148\u8D2A\u5A6A\u201D\u7684\u4E70\u5165\u673A\u4F1A\u3002\n\n2. SEC/CFTC\u8054\u5408\u6307\u5F15\uFF1A\u8FD9\u662F\u7F8E\u56FD\u52A0\u5BC6\u76D1\u7BA1\u6700\u91CD\u8981\u7684\u7A81\u7834\u4E4B\u4E00\u300268\u9875\u6587\u4EF6\u660E\u786E\u5C06BTC\u3001ETH\u3001SOL\u3001XRP\u7B4918\u79CD\u4EE3\u5E01\u5F52\u7C7B\u4E3A\u201C\u6570\u5B57\u5546\u54C1\u201D\uFF0C\u5E76\u786E\u8BA4\u6316\u77FF\u3001\u8D28\u62BC\u3001\u7A7A\u6295\u548C\u5305\u88C5\u4EE3\u5E01\u4E0D\u6784\u6210\u8BC1\u5238\u4EA4\u6613\u3002\u8FD9\u4ECE\u6839\u672C\u4E0A\u7ED3\u675F\u4E86\u7F8E\u56FD\u201C\u6267\u6CD5\u5F0F\u76D1\u7BA1\u201D\u65F6\u4EE3\uFF0C\u4E3A\u673A\u6784\u5165\u573A\u626B\u6E05\u4E86\u6700\u5927\u969C\u788D\u3002\n\n3. L1/L2\u65B9\u9762\uFF0CBase\u7684\u4E3B\u5BFC\u5730\u4F4D\u7EE7\u7EED\u5F3A\u5316\uFF0846% TVL + 60%\u8D39\u7528\u6536\u5165\uFF09\uFF0C\u663E\u793A\u6D88\u8D39\u7EA7\u5E94\u7528\u7684\u805A\u96C6\u6548\u5E94\u3002Solana\u7684Alpenglow\u5347\u7EA7\u82E5\u80FD\u5C06\u6700\u7EC8\u6027\u964D\u81F3100-150ms\uFF0C\u5C06\u8FDB\u4E00\u6B65\u5DE9\u56FA\u5176\u5728\u9AD8\u9891\u4EA4\u6613\u573A\u666F\u7684\u7ADE\u4E89\u4F18\u52BF\u3002\n\n4. DeFi/Perp DEX\u65B9\u9762\uFF0CHyperliquid\u83B7\u5F97S&P 500\u5B98\u65B9\u6388\u6743\u662F\u91CC\u7A0B\u7891\u4E8B\u4EF6\uFF0C\u8FD9\u662F\u4F20\u7EDF\u80A1\u6307\u9996\u6B21\u5728DEX\u4E0A\u5408\u89C4\u4EA4\u6613\u3002DEX\u5360\u5168\u7403\u6C38\u7EED\u5408\u7EA6\u5E02\u573A\u4EFD\u989D\u8FBE20%\uFF0C\u8FD9\u4E00\u7ED3\u6784\u6027\u8F6C\u53D8\u6B63\u5728\u52A0\u901F\u3002");

  // ===== SLIDE 4: Regulation + Fundraising =====
  const s4 = pres.addSlide();
  s4.background = { color: C.offWhite };
  s4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy } });
  s4.addText("\u2696\uFE0F  \u76D1\u7BA1\u653F\u7B56 & \u6295\u878D\u8D44", { x: 0.6, y: 0, w: 9, h: 0.7, fontSize: 20, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  // Left: Regulation
  s4.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 4.1, fill: { color: C.white }, shadow: makeShadow() });
  s4.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 0.06, h: 4.1, fill: { color: C.gold } });
  s4.addText("\u76D1\u7BA1\u4E0E\u653F\u7B56", { x: 0.7, y: 0.95, w: 4, h: 0.35, fontSize: 14, fontFace: "Arial", bold: true, color: C.navy, margin: 0 });

  const regItems = [
    { title: "SEC/CFTC\u8054\u5408\u6307\u5F15", desc: "68\u9875\u6587\u4EF6\u5EFA\u7ACB5\u7C7B\u6570\u5B57\u8D44\u4EA7\u5206\u7C7B\u6CD5\uFF0C18\u79CD\u4EE3\u5E01\u786E\u8BA4\u4E3A\u6570\u5B57\u5546\u54C1" },
    { title: "SEC\u521B\u4E1A\u516C\u53F8\u5B89\u5168\u6E2F", desc: "\u7ED9\u4E88\u52A0\u5BC6\u9879\u76EE\u878D\u8D44\u548C\u5EFA\u8BBE\u7684\u5BBD\u9650\u671F" },
    { title: "CFTC\u81EA\u6258\u7BA1\u94B1\u5305\u65E0\u884C\u52A8\u51FD", desc: "\u975E\u6258\u7BA1\u94B1\u5305\u4E0D\u53D7\u4F20\u7EDF\u91D1\u878D\u4E2D\u4ECB\u76D1\u7BA1" },
    { title: "\u4E9A\u5229\u6851\u90A3\u5DDE\u5BF9Kalshi\u63D0\u8D77\u5211\u4E8B\u6307\u63A7", desc: "20\u9879\u8F7B\u7F6A\u6307\u63A7\uFF0C\u9996\u6B21\u5DDE\u7EA7\u5211\u4E8B\u6267\u6CD5" },
    { title: "\u5185\u534E\u8FBE\u5DDE\u53D1\u5E03Kalshi\u4E34\u65F6\u7981\u4EE4", desc: "14\u5929\u7981\u6B62\u8FD0\u8425\uFF0C\u8054\u90A6/\u5DDE\u7BA1\u8F96\u6743\u4E4B\u4E89\u5347\u7EA7" },
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
    fundRow("Kalshi", "$1B", "Coatue\u9886\u6295\uFF0C\u4F30\u503C$220\u4EBF"),
    fundRow("Cryptio", "$45M", "B\u8F6E\uFF0C\u4F01\u4E1A\u7EA7\u6570\u5B57\u8D44\u4EA7\u4F1A\u8BA1"),
    fundRow("MetaComp", "$35M", "\u963F\u91CC\u5DF4\u5DF4\u53C2\u6295\uFF0C\u4E9A\u6D32-\u4E2D\u4E1C\u652F\u4ED8"),
    fundRow("GSR\u6536\u8D2DArchitech", "M&A", "\u6574\u5408\u4EE3\u5E01\u7ECF\u6D4E\u5B66\u548C\u6D41\u52A8\u6027\u7B56\u7565"),
    fundRow("Backpack ($BP)", "TGE 3/23", "\u8D28\u62BC\u4EE3\u5E01\u53EF\u8F6C\u516C\u53F8\u80A1\u6743"),
  ], {
    x: 5.3, y: 1.45, w: 4.1, colW: [1.3, 0.8, 2.0],
    border: { pt: 0.5, color: C.lightGray }, rowH: [0.3, 0.3, 0.3, 0.3, 0.3, 0.3], autoPage: false,
  });

  // Callout
  s4.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 3.4, w: 4.1, h: 0.7, fill: { color: "FFF8E7" }, shadow: makeShadow() });
  s4.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 3.4, w: 0.06, h: 0.7, fill: { color: C.gold } });
  s4.addText([
    { text: "\u5173\u6CE8\uFF1A", options: { bold: true, fontSize: 10, fontFace: "Arial", color: C.navy } },
    { text: "Kalshi $10\u4EBF\u878D\u8D44\u5F70\u663E\u9884\u6D4B\u5E02\u573A\u8D5B\u9053\u673A\u6784\u4FE1\u5FC3\uFF0C\u4F46\u591A\u5DDE\u76D1\u7BA1\u53CD\u5BF9\u663E\u793A\u5408\u89C4\u98CE\u9669\u4E0D\u5BB9\u5FFD\u89C6", options: { fontSize: 10, fontFace: "Arial", color: C.bodyText } }
  ], { x: 5.55, y: 3.45, w: 3.7, h: 0.6, margin: 0, valign: "middle" });
  footer(s4, pres);
  s4.addNotes("\u76D1\u7BA1\u4E0E\u6295\u878D\u8D44\u6DF1\u5EA6\u5206\u6790\uFF1A\n\n1. SEC/CFTC\u8054\u5408\u6307\u5F15\u662F\u7F8E\u56FD\u52A0\u5BC6\u76D1\u7BA1\u53F2\u4E0A\u7684\u5206\u6C34\u5CAD\u3002\u8FD9\u4EFD68\u9875\u6587\u4EF6\u5EFA\u7ACB\u4E86\u4E94\u7C7B\u6570\u5B57\u8D44\u4EA7\u5206\u7C7B\u6CD5\uFF0C\u660E\u786E\u5C06BTC\u3001ETH\u3001SOL\u3001XRP\u7B4918\u79CD\u4EE3\u5E01\u5F52\u7C7B\u4E3A\u201C\u6570\u5B57\u5546\u54C1\u201D\u3002\u8FD9\u610F\u5473\u7740\uFF1A\n   - \u7ED3\u675F\u4E86\u201C\u6267\u6CD5\u5F0F\u76D1\u7BA1\u201D\u65F6\u4EE3\uFF0C\u4E3A\u673A\u6784\u5165\u573A\u626B\u6E05\u4E86\u6700\u5927\u6CD5\u5F8B\u969C\u788D\n   - \u8D28\u62BC\u3001\u6316\u77FF\u3001\u7A7A\u6295\u3001\u5305\u88C5\u4EE3\u5E01\u660E\u786E\u4E0D\u6784\u6210\u8BC1\u5238\u4EA4\u6613\n   - \u4E3AETF\u3001\u673A\u6784\u6258\u7BA1\u7B49\u4EA7\u54C1\u94FA\u5E73\u9053\u8DEF\n\n2. CFTC\u81EA\u6258\u7BA1\u94B1\u5305\u65E0\u884C\u52A8\u51FD\u540C\u6837\u91CD\u8981\uFF0C\u786E\u7ACB\u4E86\u81EA\u6258\u7BA1\u7684\u6CD5\u5F8B\u6743\u5229\uFF0C\u786E\u4FDD\u94B1\u5305\u5F00\u53D1\u8005\u4E0D\u4F1A\u88AB\u5F53\u4F5C\u91D1\u878D\u4E2D\u4ECB\u76D1\u7BA1\u3002\n\n3. \u4F46\u76D1\u7BA1\u73AF\u5883\u5E76\u975E\u5168\u9762\u5229\u597D\u3002\u9884\u6D4B\u5E02\u573A\u9762\u4E34\u5DDE\u7EA7\u76D1\u7BA1\u53CD\u5BF9\uFF1A\u4E9A\u5229\u6851\u90A3\u5DDE\u5BF9Kalshi\u63D0\u8D77\u5211\u4E8B\u6307\u63A7\u3001\u5185\u534E\u8FBE\u5DDE\u53D1\u5E03\u4E34\u65F6\u7981\u4EE4\u3001\u963F\u6839\u5EF7\u5C01\u9501Polymarket\u3002\u8FD9\u663E\u793A\u8054\u90A6\u4E0E\u5DDE\u7EA7\u76D1\u7BA1\u7684\u51B2\u7A81\u6B63\u5728\u52A0\u5267\u3002\n\n4. \u6295\u878D\u8D44\u65B9\u9762\uFF0CKalshi\u7684$10\u4EBF\u878D\u8D44\u662F\u672C\u5468\u6700\u5927\u4EAE\u70B9\uFF0C\u4F30\u503C$220\u4EBF\u663E\u793A\u9884\u6D4B\u5E02\u573A\u8D5B\u9053\u7684\u5DE8\u5927\u6F5C\u529B\u3002Backpack\u7684TGE\u6A21\u5F0F\u521B\u65B0\uFF08\u8D28\u62BC\u4EE3\u5E01\u53EF\u8F6C\u80A1\u6743\uFF09\u503C\u5F97\u5173\u6CE8\uFF0C\u8FD9\u662F\u4EE3\u5E01\u4E0E\u80A1\u6743\u878D\u5408\u7684\u65B0\u5C1D\u8BD5\u3002");

  // ===== SLIDE 5: Security + Sectors =====
  const s5 = pres.addSlide();
  s5.background = { color: C.offWhite };
  s5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy } });
  s5.addText("\uD83D\uDD13  \u5B89\u5168\u4E8B\u4EF6 & \u70ED\u70B9\u8D5B\u9053", { x: 0.6, y: 0, w: 9, h: 0.7, fontSize: 20, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  // Security - left
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 2.35, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 0.35, fill: { color: C.red } });
  s5.addText("\u5B89\u5168\u4E8B\u4EF6  |  \u5468\u635F\u5931: ~$31.7M", { x: 0.6, y: 0.9, w: 4.2, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  s5.addTable([
    [
      { text: "\u4E8B\u4EF6", options: { bold: true, fontSize: 9, fontFace: "Arial", fill: { color: C.lightGray }, color: C.bodyText } },
      { text: "\u635F\u5931", options: { bold: true, fontSize: 9, fontFace: "Arial", fill: { color: C.lightGray }, color: C.bodyText, align: "right" } },
      { text: "\u65B9\u5F0F", options: { bold: true, fontSize: 9, fontFace: "Arial", fill: { color: C.lightGray }, color: C.bodyText } },
    ],
    [{ text: "Resolv Labs", options: { fontSize: 9, fontFace: "Arial" } }, { text: "$25M", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "USR\u94F8\u9020\u6F0F\u6D1E", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
    [{ text: "Venus Protocol", options: { fontSize: 9, fontFace: "Arial" } }, { text: "$3.7M", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "\u4EF7\u683C\u64CD\u7EB5", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
    [{ text: "Solv Protocol", options: { fontSize: 9, fontFace: "Arial" } }, { text: "$2.7M", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "\u5B9A\u5411\u5B89\u5168\u6F0F\u6D1E", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
    [{ text: "Bitrefill (Lazarus)", options: { fontSize: 9, fontFace: "Arial" } }, { text: "N/A", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "\u6734\u9C9C\u9ED1\u5BA2APT\u653B\u51FB", options: { fontSize: 9, fontFace: "Arial", color: C.red, bold: true } }],
  ], {
    x: 0.55, y: 1.35, w: 4.2, colW: [1.5, 0.9, 1.8],
    border: { pt: 0.5, color: C.lightGray }, rowH: [0.28, 0.28, 0.28, 0.28, 0.28], autoPage: false,
  });

  // Stablecoins - left bottom
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.5, w: 4.5, h: 1.5, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.5, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s5.addText("\u7A33\u5B9A\u5E01", { x: 0.6, y: 3.5, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });
  s5.addText([
    "USDC\u5E74\u5EA6\u8F6C\u8D26\u91CF\u8D85\u8D8AUSDT ($2.2T vs $1.3T)",
    "\u7A33\u5B9A\u5E01\u603B\u5E02\u503C\u8FBE$316B\u65B0\u9AD8",
    "Ethena sUSDe\u6536\u76CA\u538B\u7F29\u81F33.49-3.72% APY",
  ].map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < 2, fontSize: 10, fontFace: "Arial", color: C.bodyText } })), { x: 0.6, y: 3.95, w: 4.1, h: 0.9, margin: 0, paraSpaceAfter: 4 });

  // AI x Crypto - right top
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 4.5, h: 1.85, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s5.addText("AI \u00D7 Crypto", { x: 5.3, y: 0.9, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });
  s5.addText([
    "Bittensor\u5B8C\u621072B\u53C2\u6570\u53BB\u4E2D\u5FC3\u5316LLM\u8BAD\u7EC3\uFF0CTAO\u5468\u6DA856%",
    "Grayscale\u63D0\u4EA4TAO\u73B0\u8D27ETP\u7533\u8BF7",
    "Akash Network BME\u786C\u5206\u53C9\uFF0CAKT\u6DA820%",
    "MoonPay+Ledger\u63A8\u51FA\u9996\u4E2A\u786C\u4EF6\u7B7E\u540D\u5668AI\u4EE3\u7406",
  ].map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < 3, fontSize: 10, fontFace: "Arial", color: C.bodyText } })), { x: 5.3, y: 1.35, w: 4.1, h: 1.3, margin: 0, paraSpaceAfter: 4 });

  // Prediction Markets - right bottom
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 2.95, w: 4.5, h: 2.05, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 2.95, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s5.addText("\u9884\u6D4B\u5E02\u573A", { x: 5.3, y: 2.95, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });
  s5.addText([
    "Polymarket\u4E0EMLB\u7B7E$1.5\u4EBF\u72EC\u5BB6\u5408\u4F5C",
    "Kalshi\u878D\u8D44$10\u4EBF\uFF0C\u4F30\u503C$220\u4EBF",
    "\u4E9A\u5229\u6851\u90A3\u5DDE\u63D0\u8D77\u5211\u4E8B\u6307\u63A7\uFF0C\u5185\u534E\u8FBE\u53D1\u4E34\u65F6\u7981\u4EE4",
    "\u963F\u6839\u5EF7\u5168\u56FD\u5C01\u9501Polymarket",
  ].map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < 3, fontSize: 10, fontFace: "Arial", color: C.bodyText } })), { x: 5.3, y: 3.4, w: 4.1, h: 1.5, margin: 0, paraSpaceAfter: 4 });
  footer(s5, pres);
  s5.addNotes("\u5B89\u5168\u4E8B\u4EF6\u4E0E\u70ED\u70B9\u8D5B\u9053\u6DF1\u5EA6\u5206\u6790\uFF1A\n\n1. \u5B89\u5168\u4E8B\u4EF6\uFF1A\u672C\u5468\u603B\u635F\u5931~$31.7M\uFF0C\u4EE5Resolv Labs $25M\u6700\u4E3A\u4E25\u91CD\u3002USR\u7A33\u5B9A\u5E01\u94F8\u9020\u6F0F\u6D1E\u5BFC\u81F4\u653B\u51FB\u8005\u5C06200K USDC\u8F6C\u4E3A$25M USR\uFF0C\u7A33\u5B9A\u5E01\u8131\u9528\u8FBE80%\u3002\u8FD9\u518D\u6B21\u8B66\u793A\u65B0\u578B\u7A33\u5B9A\u5E01\u534F\u8BAE\u7684\u667A\u80FD\u5408\u7EA6\u98CE\u9669\u3002Bitrefill\u4E8B\u4EF6\u786E\u8BA4\u4E3A\u6734\u9C9C Lazarus Group \u6240\u4E3A\uFF0C\u663E\u793A\u56FD\u5BB6\u7EA7\u9ED1\u5BA2\u7EC4\u7EC7\u5BF9\u52A0\u5BC6\u884C\u4E1A\u7684\u5A01\u80C1\u6301\u7EED\u5B58\u5728\u3002\n\n2. AI\u00D7Crypto\uFF1A\u672C\u5468\u6700\u5927\u4EAE\u70B9\u662FBittensor\u6210\u529F\u5B8C\u621072B\u53C2\u6570\u6A21\u578B\u7684\u53BB\u4E2D\u5FC3\u5316\u8BAD\u7EC3\uFF0C\u8FD9\u662F\u53BB\u4E2D\u5FC3\u5316AI\u8BAD\u7EC3\u89C4\u6A21\u7684\u5386\u53F2\u6027\u7A81\u7834\u300270\u4E2A\u5168\u7403\u8D21\u732E\u8005\u4F7F\u7528\u5546\u54C1\u786C\u4EF6\u5B8C\u6210\uFF0C\u8BC1\u660E\u4E86\u53BB\u4E2D\u5FC3\u5316\u8BA1\u7B97\u7684\u53EF\u884C\u6027\u3002Grayscale\u63D0\u4EA4TAO ETP\u7533\u8BF7\u662F\u673A\u6784\u5316\u7684\u91CD\u8981\u4FE1\u53F7\u3002Akash\u7684BME\u786C\u5206\u53C9\u5F15\u5165\u6C38\u4E45\u6027\u4EE3\u5E01\u9500\u6BC1\u673A\u5236\uFF0C\u4F7F\u4EE3\u5E01\u4F9B\u5E94\u4E0E\u7F51\u7EDC\u4F7F\u7528\u76F4\u63A5\u6302\u94A9\u3002\n\n3. \u9884\u6D4B\u5E02\u573A\uFF1APolymarket\u4E0EMLB\u7684$1.5\u4EBF\u72EC\u5BB6\u5408\u4F5C\u662F\u5386\u53F2\u6027\u91CC\u7A0B\u7891\uFF0C\u8FD9\u662F\u9884\u6D4B\u5E02\u573A\u4E0E\u4E3B\u6D41\u4F53\u80B2\u7684\u9996\u6B21\u6DF1\u5EA6\u878D\u5408\u3002\u4F46\u540C\u65F6\u591A\u5DDE\u76D1\u7BA1\u53CD\u5BF9\uFF08\u4E9A\u5229\u6851\u90A3\u5DDE\u5211\u4E8B\u6307\u63A7\u3001\u5185\u534E\u8FBE\u4E34\u65F6\u7981\u4EE4\u3001\u963F\u6839\u5EF7\u5C01\u9501\uFF09\u663E\u793A\u8D5B\u9053\u7684\u5408\u89C4\u98CE\u9669\u4ECD\u5728\u3002\u8054\u90A6\u4E0E\u5DDE\u7684\u7BA1\u8F96\u6743\u4E4B\u4E89\u53EF\u80FD\u6210\u4E3A\u5F71\u54CD\u8D5B\u9053\u53D1\u5C55\u7684\u5173\u952E\u53D8\u91CF\u3002\n\n4. \u7A33\u5B9A\u5E01\uFF1AUSDC\u5E74\u5EA6\u8F6C\u8D26\u91CF\u8D85\u8D8AUSDT\u662F\u91CD\u8981\u8F6C\u6298\u70B9\uFF0C\u53CD\u6620\u76D1\u7BA1\u5408\u89C4\u6027\u5BF9\u673A\u6784\u7528\u6237\u7684\u91CD\u8981\u6027\u3002Ethena\u6536\u76CA\u538B\u7F29\u81F33.5%\u5DE6\u53F3\u503C\u5F97\u5173\u6CE8\uFF0C\u56E0\u4E3A\u8FD9\u53CD\u6620\u4E86\u5F53\u524D\u8D1F\u8D44\u91D1\u8D39\u7387\u73AF\u5883\u5BF9USDe\u6A21\u5F0F\u7684\u538B\u529B\u3002");

  // ===== SLIDE 6: Catalysts + Outlook =====
  const s6 = pres.addSlide();
  s6.background = { color: C.offWhite };
  s6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy } });
  s6.addText("\uD83D\uDD2E  \u672A\u6765\u50AC\u5316\u5242 & \u5C55\u671B", { x: 0.6, y: 0, w: 9, h: 0.7, fontSize: 20, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  const catHeader = [
    { text: "\u65E5\u671F", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial", align: "center" } },
    { text: "\u4E8B\u4EF6", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial" } },
    { text: "\u5F71\u54CD\u8D44\u4EA7", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial", align: "center" } },
    { text: "\u9884\u671F\u5F71\u54CD", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial" } },
  ];
  const catRow = (date, event, asset, impact, hl = false) => [
    { text: date, options: { fontSize: 10, fontFace: "Arial", align: "center", color: C.bodyText, fill: hl ? { color: "FFF3E0" } : undefined } },
    { text: event, options: { fontSize: 10, fontFace: "Arial", color: C.bodyText, bold: hl, fill: hl ? { color: "FFF3E0" } : undefined } },
    { text: asset, options: { fontSize: 10, fontFace: "Arial", align: "center", color: C.navy, bold: true, fill: hl ? { color: "FFF3E0" } : undefined } },
    { text: impact, options: { fontSize: 9, fontFace: "Arial", color: C.medGray, fill: hl ? { color: "FFF3E0" } : undefined } },
  ];
  s6.addTable([
    catHeader,
    catRow("3/23", "Backpack ($BP) TGE", "BP", "\u8D28\u62BC\u4EE3\u5E01\u53EF\u8F6C\u80A1\u6743\u521B\u65B0\u6A21\u5F0F", true),
    catRow("3/23", "Akash BME\u786C\u5206\u53C9\u4E0A\u7EBF", "AKT", "\u4EE3\u5E01\u901A\u7F29+WASM\u652F\u6301"),
    catRow("3/24-26", "Digital Asset Summit (NYC)", "RWA", "\u673A\u6784\u91C7\u7528\u4FE1\u53F7"),
    catRow("3/30-4/2", "EthCC 9 (Cannes)", "ETH", "\u6B27\u6D32\u6700\u5927ETH\u4F1A\u8BAE"),
    catRow("4/3", "CLARITY\u6CD5\u6848\u53EF\u80FD\u7B7E\u7F72", "\u5168\u5E02\u573A", "\u5546\u54C1vs\u8BC1\u5238\u6CD5\u5F8B\u5B9A\u4E49\u91CC\u7A0B\u7891", true),
    catRow("4\u6708", "Uniswap\u8D39\u7528\u5F00\u5173\u5B9E\u65BD", "UNI", "8\u6761L2\u542F\u52A8\u624B\u7EED\u8D39\u6536\u5165"),
    catRow("4\u6708\u521D", "Polkadot 2.1.1\u5347\u7EA7", "DOT", "\u52A8\u6001\u5206\u914D\u6C60\uFF0C\u7EC8\u6B62Treasury\u9500\u6BC1"),
  ], {
    x: 0.5, y: 0.9, w: 9, colW: [0.9, 3.2, 1.2, 3.7],
    border: { pt: 0.5, color: C.lightGray }, rowH: [0.32, 0.32, 0.32, 0.32, 0.32, 0.32, 0.32, 0.32], autoPage: false,
  });

  // Outlook box
  s6.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.65, w: 9, h: 1.3, fill: { color: C.white }, shadow: makeShadow() });
  s6.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.65, w: 0.06, h: 1.3, fill: { color: C.gold } });
  s6.addText("\u672C\u5468\u5C0F\u7ED3", { x: 0.8, y: 3.7, w: 8, h: 0.3, fontSize: 13, fontFace: "Arial", bold: true, color: C.navy, margin: 0 });
  s6.addText([
    { text: "\u5E02\u573A\u60C5\u7EEA\u8DCC\u5165\u201C\u6781\u5EA6\u6050\u60E7\u201D\uFF08FGI=10\uFF09\uFF0CBTC\u4ECE$76K\u95EA\u5D29\u81F3$68.8K\u3002\u4F46\u76D1\u7BA1\u7AEF\u51FA\u73B0\u5386\u53F2\u6027\u7A81\u7834\uFF08SEC/CFTC\u8054\u5408\u6307\u5F15\uFF09\uFF0C\u4E2D\u957F\u671F\u6781\u4E3A\u79EF\u6781\u3002", options: { breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "\u7EC4\u5408\u5173\u6CE8\uFF1A", options: { bold: true, fontSize: 10, fontFace: "Arial", color: C.navy } },
    { text: " Backpack TGE (3/23)\u3001HYPE ETF\u7533\u8BF7\u8FDB\u5C55\u3001Pendle TVL\u7A81\u7834$40\u4EBF\u3001ENA\u6536\u76CA\u538B\u7F29\u9700\u89C2\u5BDF\u3001\u4E2D\u4E1C\u5730\u7F18\u5371\u673A\u540E\u7EED\u6F14\u53D8\u3002", options: { fontSize: 10, fontFace: "Arial", color: C.bodyText } },
  ], { x: 0.8, y: 4.0, w: 8.5, h: 0.85, margin: 0 });

  s6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s6.addText(`\u52A0\u5BC6\u5E02\u573A\u5468\u62A5 | ${WEEK}  |  CONFIDENTIAL`, { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });
  s6.addNotes("\u672A\u6765\u50AC\u5316\u5242\u4E0E\u5C55\u671B\u6DF1\u5EA6\u5206\u6790\uFF1A\n\n1. Backpack TGE\uFF083/23\uFF09\u662F\u672C\u5468\u6700\u503C\u5F97\u5173\u6CE8\u7684\u4E8B\u4EF6\uFF0C\u4F5C\u4E3A\u6301\u4ED3\u8D44\u4EA7\u9700\u5BC6\u5207\u8DDF\u8E2A\u3002\u5176\u201C\u8D28\u62BC\u4EE3\u5E01\u53EF\u8F6C\u80A1\u6743\u201D\u6A21\u5F0F\u662F\u4EE3\u5E01\u4E0E\u80A1\u6743\u878D\u5408\u7684\u521B\u65B0\u5C1D\u8BD5\uFF0C\u53EF\u80FD\u5F15\u53D1\u884C\u4E1A\u6548\u4EFF\u3002\n\n2. CLARITY\u6CD5\u6848\u82E5\u57284/3\u7B7E\u7F72\uFF0C\u5C06\u662F\u7F8E\u56FD\u52A0\u5BC6\u76D1\u7BA1\u6700\u91CD\u8981\u7684\u7ACB\u6CD5\u91CC\u7A0B\u7891\u4E4B\u4E00\uFF0C\u4E3A\u6570\u5B57\u8D44\u4EA7\u7684\u5546\u54C1vs\u8BC1\u5238\u5F52\u7C7B\u63D0\u4F9B\u660E\u786E\u6CD5\u5F8B\u5B9A\u4E49\u3002\u8FD9\u5C06\u76F4\u63A5\u5F71\u54CD\u5168\u5E02\u573A\u3002\n\n3. EthCC 9 (Cannes)\u662F\u6B27\u6D32\u6700\u5927\u7684\u4EE5\u592A\u574A\u4F1A\u8BAE\uFF0C\u5386\u53F2\u4E0A\u5E38\u6709\u91CD\u5927\u6280\u672F\u548C\u751F\u6001\u7CFB\u7EDF\u516C\u544A\u3002\u7ED3\u5408Glamsterdam\u5347\u7EA7\u9884\u671F\uFF0C\u53EF\u80FD\u5BF9ETH\u4EF7\u683C\u4EA7\u751F\u79EF\u6781\u5F71\u54CD\u3002\n\n4. \u7EC4\u5408\u7BA1\u7406\u89D2\u5EA6\uFF1A\n   - PENDLE: TVL\u7A81\u7834$40\u4EBF\u662F\u91CD\u8981\u91CC\u7A0B\u7891\uFF0C\u4F46\u4EF7\u683C\u4E0B\u8DCC6.4%\u663E\u793A\u77ED\u671F\u5E02\u573A\u60C5\u7EEA\u8D1F\u9762\n   - ENA: sUSDe\u6536\u76CA\u538B\u7F29\u81F33.5%\u662F\u5173\u952E\u98CE\u9669\u4FE1\u53F7\uFF0C\u56E0\u4E3A\u8D1F\u8D44\u91D1\u8D39\u7387\u73AF\u5883\u76F4\u63A5\u538B\u7F29\u4E86USDe\u6A21\u5F0F\u7684\u6536\u76CA\u7A7A\u95F4\n   - HYPE: Grayscale\u63D0\u4EA4ETF\u7533\u8BF7\u662F\u4E2D\u957F\u671F\u91CD\u5927\u5229\u597D\uFF0C\u4F46\u9700\u89C2\u5BDF\u5BA1\u6279\u8FDB\u5EA6\n   - \u4E2D\u4E1C\u5730\u7F18\u5371\u673A\u540E\u7EED\u6F14\u53D8\u662F\u6700\u5927\u4E0D\u786E\u5B9A\u6027\u56E0\u7D20");

  const outPath = "/Users/lx/Github/crypto-weekly-digest/reports/2026-03-16/slides-zh.pptx";
  await pres.writeFile({ fileName: outPath });
  console.log("\u2705 Written to " + outPath);
}

generate().catch(e => { console.error(e); process.exit(1); });
