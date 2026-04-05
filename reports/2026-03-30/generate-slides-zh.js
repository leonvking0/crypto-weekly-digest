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
  pres.title = "\u52A0\u5BC6\u5E02\u573A\u5468\u62A5 2026.03.30\u201304.05";

  // ============ SLIDE 1: Title ============
  const s1 = pres.addSlide();
  s1.background = { color: C.darkNavy };
  s1.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.6, w: 1.2, h: 0.04, fill: { color: C.gold } });
  s1.addText("\u52A0\u5BC6\u5E02\u573A\u5468\u62A5", {
    x: 0.8, y: 1.8, w: 8, h: 1.0,
    fontSize: 42, fontFace: "Arial", bold: true, color: C.white, margin: 0
  });
  s1.addText("2026\u5E743\u670830\u65E5 \u2014 4\u67085\u65E5", {
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
  s1.addNotes("\u672C\u5468\u52A0\u5BC6\u5E02\u573A\u5468\u62A5\u8986\u76D6 2026\u5E743\u670830\u65E5\u81F34\u67085\u65E5\u3002\u672C\u5468\u5E02\u573A\u5728\u201C\u6781\u5EA6\u6050\u60E7\u201D\u60C5\u7EEA\u4E2D\u5EF6\u7EED\u4F4E\u6CE2\u52A8\u6574\u56FA\uFF0CBTC \u7EF4\u6301 $65K\u2013$70K \u533A\u95F4\u9707\u8361\u3002\u672C\u5468\u6700\u91CD\u5927\u4E8B\u4EF6\u4E3A Drift Protocol \u906D\u53D7 $2.85 \u4EBF\u653B\u51FB\uFF0C\u4E3A 2026 \u5E74\u8FC4\u4ECA\u6700\u5927\u5355\u7B14\u5B89\u5168\u4E8B\u4EF6\u3002\u76D1\u7BA1\u65B9\u9762\uFF0C\u7F8E\u56FD\u8D22\u653F\u90E8\u53D1\u5E03 GENIUS Act \u5B9E\u65BD\u7EC6\u5219\uFF0CCLARITY Act \u7A33\u5B9A\u5E01\u6536\u76CA\u59A5\u534F\u65B9\u6848\u63A5\u8FD1\u8FBE\u6210\uFF0C\u76D1\u7BA1\u6846\u67B6\u6B63\u5728\u52A0\u901F\u843D\u5730\u3002Aave V4 \u4E3B\u7F51\u4E0A\u7EBF\u548C Solana Alpenglow \u5347\u7EA7\u5219\u4EE3\u8868\u4E86\u6280\u672F\u5C42\u9762\u7684\u91CD\u8981\u8FDB\u5C55\u3002\u6050\u60E7\u8D2A\u5A6A\u6307\u6570\u4F4E\u81F3 12\uFF0C\u53CD\u6620\u5E02\u573A\u60C5\u7EEA\u6781\u5EA6\u4F4E\u8FF7\uFF0C\u4F46\u57FA\u672C\u9762\u6539\u5584\u4FE1\u53F7\u6B63\u5728\u7D2F\u79EF\u3002");

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
  s2.addText("\u6050\u60E7\u8D2A\u5A6A\u6307\u6570", { x: c3x + 0.25, y: cardY + 0.1, w: 2.4, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.medGray, bold: true, margin: 0 });
  s2.addText("12", { x: c3x + 0.25, y: cardY + 0.35, w: 1, h: 0.4, fontSize: 28, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
  s2.addText("\u6781\u5EA6\u6050\u60E7", { x: c3x + 1.1, y: cardY + 0.42, w: 1.5, h: 0.3, fontSize: 13, fontFace: "Arial", color: C.red, margin: 0 });

  // Price table
  const tableY = 2.25;
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
  const tableData = [
    headerRow,
    makeRow("SOL", "$79.45", "-2.8%", "Drift\u653B\u51FB\u51B2\u51FB\u751F\u6001", true),
    makeRow("HYPE", "$35.47", "-9.0%", "Q1\u603B\u91CF $4927\u4EBF", true),
    makeRow("PENDLE", "$0.98", "-12.4%", "\u6536\u76CA\u4EE3\u5E01\u8D5B\u9053\u627F\u538B", true),
    makeRow("ENA", "$0.08", "-14.4%", "\u968F\u5927\u76D8\u8D70\u5F31", true),
    makeRow("DYDX", "$0.10", "+7.3%", "\u6280\u672F\u53CD\u5F39", false),
  ];
  s2.addTable(tableData, {
    x: 0.6, y: tableY, w: 8.8, colW: [1.3, 1.8, 1.5, 4.2],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.35, 0.33, 0.33, 0.33, 0.33, 0.33],
    autoPage: false,
  });

  s2.addText("\u5B8F\u89C2\u9A71\u52A8\uFF1A\u7F8E\u8054\u50A8\u7EF4\u6301\u5229\u7387 3.50%\u20133.75%\uFF0C\u4E0A\u8C03\u901A\u80C0\u9884\u6D4B\uFF1BDrift $2.85\u4EBF\u653B\u51FB\u6253\u51FB Solana \u751F\u6001\u4FE1\u5FC3\uFF1B\u5E02\u573A\u60C5\u7EEA\u7EF4\u6301\u201C\u6781\u5EA6\u6050\u60E7\u201D", {
    x: 0.6, y: 4.55, w: 8.8, h: 0.4,
    fontSize: 10, fontFace: "Arial", color: C.medGray, italic: true, margin: 0
  });

  s2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s2.addText("\u52A0\u5BC6\u5E02\u573A\u5468\u62A5 | 2026.03.30\u201304.05", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  s2.addNotes("BTC \u672C\u5468\u5728 $65,000\u2013$70,000 \u533A\u95F4\u5185\u7A84\u5E45\u9707\u8361\uFF0C\u5468\u56DB\u4E00\u5EA6\u6D4B\u8BD5 $69,000 \u963B\u529B\u540E\u56DE\u843D\uFF0C\u5468\u6DA8\u5E45\u4EC5 +1.1%\u3002ETH \u5728 $2,020\u2013$2,140 \u7A84\u5E45\u6CE2\u52A8\uFF0C\u5468\u6DA8 +3.0% \u8868\u73B0\u7565\u4F18\u4E8E BTC\uFF0C\u4F46 ETF \u51C0\u6D41\u51FA $42M \u663E\u793A\u673A\u6784\u5BF9 ETH \u7684\u6001\u5EA6\u4ECD\u504F\u8C28\u614E\u3002\n\n\u5C71\u5BE8\u5E01\u666E\u904D\u56DE\u8C03\uFF1AHYPE -9.0%\u3001ENA -14.4%\u3001PENDLE -12.4%\uFF0C\u53CD\u6620\u98CE\u9669\u504F\u597D\u6781\u5EA6\u6536\u7F29\u3002DYDX \u662F\u5C11\u6570\u6B63\u6536\u76CA\u8D44\u4EA7\uFF08+7.3%\uFF09\uFF0C\u4E3B\u8981\u53D7\u6280\u672F\u53CD\u5F39\u9A71\u52A8\u3002\n\n\u6050\u60E7\u8D2A\u5A6A\u6307\u6570\u4F4E\u81F3 12\uFF0C\u5904\u4E8E\u201C\u6781\u5EA6\u6050\u60E7\u201D\u533A\u95F4\u3002\u8FD9\u662F\u81EA FTX \u5D29\u76D8\u4EE5\u6765\u6700\u4F4E\u6C34\u5E73\u4E4B\u4E00\u3002\u4F46\u503C\u5F97\u6CE8\u610F\u7684\u662F\uFF0C\u6781\u5EA6\u6050\u60E7\u5386\u53F2\u4E0A\u5F80\u5F80\u662F\u8F83\u597D\u7684\u903E\u671F\u4E70\u5165\u70B9\u3002\u5B8F\u89C2\u73AF\u5883\u4E0D\u5229\u2014\u2014\u7F8E\u8054\u50A8\u7EF4\u6301\u9AD8\u5229\u7387\u4E14\u4E0A\u8C03\u901A\u80C0\u9884\u6D4B\uFF0C\u53E0\u52A0\u5730\u7F18\u653F\u6CBB\u98CE\u9669\uFF08\u4F0A\u4EE5\u51B2\u7A81\u5347\u7EA7\uFF09\uFF0C\u538B\u5236\u4E86\u98CE\u9669\u8D44\u4EA7\u7684\u4EF7\u683C\u8868\u73B0\u3002\n\nDrift Protocol $2.85 \u4EBF\u653B\u51FB\u662F\u672C\u5468\u6700\u5927\u9ED1\u5929\u9E45\u4E8B\u4EF6\uFF0C\u76F4\u63A5\u5BFC\u81F4 SOL \u6210\u4E3A\u4E3B\u8981\u8D44\u4EA7\u4E2D\u8868\u73B0\u6700\u5DEE\u8005\uFF0C\u7F51\u7EDC\u6708\u6D3B\u5730\u5740\u4E0B\u964D 11%\u3002\u8FD9\u4E5F\u5F15\u53D1\u4E86\u5BF9 DeFi \u5B89\u5168\u6027\u548C\u9884\u8A00\u673A\u64CD\u7EB5\u98CE\u9669\u7684\u5E7F\u6CDB\u8BA8\u8BBA\u3002");

  // ============ SLIDE 3: Top Stories + DeFi/L1 ============
  const s3 = pres.addSlide();
  s3.background = { color: C.offWhite };
  s3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy } });
  s3.addText("\uD83D\uDD25  \u672C\u5468\u8981\u95FB & \u884C\u4E1A\u52A8\u6001", {
    x: 0.6, y: 0, w: 9, h: 0.7,
    fontSize: 20, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0
  });

  // Left column: Top 5 stories
  s3.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 4.1, fill: { color: C.white }, shadow: makeShadow() });
  s3.addText("\u672C\u5468\u8981\u95FB TOP 5", {
    x: 0.6, y: 0.95, w: 4, h: 0.35,
    fontSize: 13, fontFace: "Arial", bold: true, color: C.navy, margin: 0
  });

  const stories = [
    ["1", "Drift Protocol \u906D $2.85\u4EBF\u653B\u51FB", "Lazarus\u96C6\u56E2\u793E\u5DE5+\u9884\u8A00\u673A\u64CD\u7EB5\uFF0C12\u5206\u949F\u5B8C\u6210"],
    ["2", "\u8D22\u653F\u90E8\u53D1\u5E03 GENIUS Act \u5B9E\u65BD\u7EC6\u5219", "\u7A33\u5B9A\u5E01 1:1 \u50A8\u5907 + 12\u4E2A\u6708\u8FD0\u8425\u8D39\u7528\u7F13\u51B2"],
    ["3", "CLARITY Act \u7A33\u5B9A\u5E01\u6536\u76CA\u59A5\u534F\u63A5\u8FD1\u8FBE\u6210", "\u7981\u6B62\u88AB\u52A8\u6536\u76CA\uFF0C\u5141\u8BB8\u57FA\u4E8E\u6D3B\u52A8\u7684\u6FC0\u52B1"],
    ["4", "Aave V4 \u4E3B\u7F51\u4E0A\u7EBF + BGD Labs \u9000\u51FA", "\u6A21\u5757\u5316 hub-and-spoke \u67B6\u6784\uFF0C\u673A\u6784\u4FE1\u8D37\u4E09\u4E2D\u5FC3"],
    ["5", "Solana Alpenglow \u83B7 98% \u9A8C\u8BC1\u8005\u6279\u51C6", "\u7EC8\u5C40\u6027 ~400ms \u2192 100\u2013150ms"],
  ];

  stories.forEach((s, i) => {
    const itemY = 1.38 + i * 0.72;
    s3.addShape(pres.shapes.OVAL, { x: 0.65, y: itemY, w: 0.3, h: 0.3, fill: { color: C.gold } });
    s3.addText(s[0], { x: 0.65, y: itemY, w: 0.3, h: 0.3, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, align: "center", valign: "middle", margin: 0 });
    s3.addText(s[1], { x: 1.1, y: itemY - 0.02, w: 3.6, h: 0.25, fontSize: 11, fontFace: "Arial", bold: true, color: C.bodyText, margin: 0 });
    s3.addText(s[2], { x: 1.1, y: itemY + 0.22, w: 3.6, h: 0.25, fontSize: 9, fontFace: "Arial", color: C.medGray, margin: 0 });
  });

  // Right column: L1/L2
  s3.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 4.5, h: 1.95, fill: { color: C.white }, shadow: makeShadow() });
  s3.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s3.addText("Layer 1 & Layer 2", { x: 5.3, y: 0.9, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  const l1Items = [
    "Base \u5360 L2 DEX \u4EA4\u6613\u91CF 50%\uFF0CTVL $41.5\u4EBF",
    "\u201CStablechains\u201D\u65B0\u53D9\u4E8B\uFF1Aa16z \u9886\u6295 Better Money $1000\u4E07",
    "Monad TVL \u7A81\u7834 $3.5\u4EBF\uFF0C\u4F46 FDV \u4E0B\u884C",
    "Cosmos \u751F\u6001\u53D7\u632B\uFF1ALeap Wallet \u5BA3\u5E03\u5173\u95ED",
  ];
  s3.addText(l1Items.map((t, i) => ({
    text: t,
    options: { bullet: true, breakLine: i < l1Items.length - 1, fontSize: 10, fontFace: "Arial", color: C.bodyText }
  })), { x: 5.3, y: 1.32, w: 4.1, h: 1.45, margin: 0, paraSpaceAfter: 4 });

  // DeFi box
  s3.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 3.05, w: 4.5, h: 1.95, fill: { color: C.white }, shadow: makeShadow() });
  s3.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 3.05, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s3.addText("DeFi & \u6C38\u7EED\u5408\u7EA6", { x: 5.3, y: 3.05, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  const defiItems = [
    "Aave V4 \u4E3B\u7F51\u4E0A\u7EBF\uFF0C\u6A21\u5757\u5316\u4E09\u4E2D\u5FC3\u67B6\u6784",
    "Lido V3 stVaults\uFF1AEarnETH + EarnUSD \u591A\u5143\u6536\u76CA",
    "Hyperliquid Q1 $4927\u4EBF\uFF0C\u8FDB\u5165\u5168\u7403\u884D\u751F\u54C1\u524D\u5341",
    "DeFi TVL ~$92\u201393B\uFF1BRWA \u6301\u7EED\u589E\u957F\u81F3 $170\u4EBF",
  ];
  s3.addText(defiItems.map((t, i) => ({
    text: t,
    options: { bullet: true, breakLine: i < defiItems.length - 1, fontSize: 10, fontFace: "Arial", color: C.bodyText }
  })), { x: 5.3, y: 3.47, w: 4.1, h: 1.45, margin: 0, paraSpaceAfter: 4 });

  s3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s3.addText("\u52A0\u5BC6\u5E02\u573A\u5468\u62A5 | 2026.03.30\u201304.05", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  s3.addNotes("Drift Protocol $2.85 \u4EBF\u653B\u51FB\u662F\u672C\u5468\u6700\u91CD\u5927\u7684\u9ED1\u5929\u9E45\u4E8B\u4EF6\u3002\u653B\u51FB\u8005\u4E3A\u671D\u9C9C Lazarus \u96C6\u56E2\uFF0C\u901A\u8FC7\u516D\u4E2A\u6708\u7684\u793E\u4F1A\u5DE5\u7A0B\u5B66\u6E17\u900F\uFF0C\u4F2A\u88C5\u91CF\u5316\u4EA4\u6613\u516C\u53F8\u8BF1\u5BFC\u591A\u7B7E\u6301\u6709\u8005\u9884\u7B7E\u4EA4\u6613\uFF0C\u7136\u540E\u4F2A\u9020\u8D44\u4EA7\u64CD\u7EB5\u9884\u8A00\u673A\u4EF7\u683C\uFF0C12\u5206\u949F\u5185\u5B8C\u6210\u653B\u51FB\u3002\u8FD9\u66B4\u9732\u4E86\u591A\u7B7E\u94B1\u5305\u5B89\u5168\u6027\u548C\u9884\u8A00\u673A\u64CD\u7EB5\u7684\u53CC\u91CD\u98CE\u9669\u3002\n\n\u76D1\u7BA1\u65B9\u9762\uFF0CGENIUS Act \u5B9E\u65BD\u7EC6\u5219\u548C CLARITY Act \u59A5\u534F\u65B9\u6848\u4EE3\u8868\u4E86\u7F8E\u56FD\u52A0\u5BC6\u76D1\u7BA1\u6846\u67B6\u7684\u5B9E\u8D28\u6027\u8FDB\u5C55\u3002\u8D22\u653F\u90E8\u8981\u6C42\u7A33\u5B9A\u5E01\u53D1\u884C\u65B9\u7EF4\u6301 12 \u4E2A\u6708\u8FD0\u8425\u8D39\u7528\u7F13\u51B2\u7684\u8981\u6C42\u76F8\u5F53\u4E25\u683C\uFF0C\u53EF\u80FD\u5C06\u6DE1\u51FA\u5C0F\u578B\u53D1\u884C\u65B9\u3002CLARITY Act \u7A33\u5B9A\u5E01\u6536\u76CA\u7981\u4EE4\u7684\u59A5\u534F\u2014\u2014\u7981\u6B62\u88AB\u52A8\u6301\u6709\u6536\u76CA\u4F46\u5141\u8BB8\u57FA\u4E8E\u6D3B\u52A8\u7684\u6FC0\u52B1\u2014\u2014\u662F\u5BF9 Circle \u7B49\u4F9D\u8D56\u50A8\u5907\u5229\u606F\u6536\u5165\u7684\u516C\u53F8\u7684\u91CD\u5927\u5229\u597D\u3002\n\nAave V4 \u7684\u6A21\u5757\u5316\u67B6\u6784\u662F DeFi \u534F\u8BAE\u5411\u673A\u6784\u5316\u8FDB\u5316\u7684\u6807\u5FD7\u6027\u4E8B\u4EF6\u3002BGD Labs \u9000\u51FA\u53CD\u6620\u4E86 DeFi \u6CBB\u7406\u4E2D\u6838\u5FC3\u8D21\u732E\u8005\u4E0E DAO \u4E4B\u95F4\u7684\u7ED3\u6784\u6027\u7D27\u5F20\u5173\u7CFB\u3002\n\nHyperliquid \u8FDB\u5165\u5168\u7403\u884D\u751F\u54C1\u524D\u5341\u662F\u6C38\u7EED DEX \u8D5B\u9053\u7684\u91CC\u7A0B\u7891\u4E8B\u4EF6\uFF0CQ1 $4927\u4EBF\u7684\u4EA4\u6613\u91CF\u8BC1\u660E\u94FE\u4E0A\u884D\u751F\u54C1\u5DF2\u4ECE\u8FB9\u7F18\u4EA7\u54C1\u6210\u957F\u4E3A\u4E3B\u6D41\u4EA4\u6613\u573A\u6240\u3002");

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
    { title: "GENIUS Act \u5B9E\u65BD\u7EC6\u5219", desc: "\u8D22\u653F\u90E8 4/1 \u53D1\u5E03\uFF0C1:1 \u50A8\u5907 + 12\u6708\u8FD0\u8425\u7F13\u51B2" },
    { title: "CLARITY Act \u59A5\u534F\u65B9\u6848", desc: "\u7981\u88AB\u52A8\u6536\u76CA\uFF0C\u5141\u6D3B\u52A8\u6FC0\u52B1\uFF1B\u59D4\u5458\u4F1A\u5BA1\u8BAE4\u6708\u4E0B\u65EC" },
    { title: "CFTC \u653F\u7B56\u8F6C\u5411", desc: "\u65B0\u4EFB\u6267\u6CD5\u4E3B\u7BA1\u5BA3\u5E03\u4E0D\u518D\u4EE5\u8BC9\u8BBC\u8BBE\u5B9A\u653F\u7B56" },
    { title: "\u94F6\u884C\u57FA\u7840\u8BBE\u65BD\u7A81\u7834", desc: "Kraken \u83B7\u7F8E\u8054\u50A8\u4E3B\u8D26\u6237\uFF1BRipple/Crypto.com \u83B7OCC\u724C\u7167" },
    { title: "Charles Schwab \u5F00\u653E\u52A0\u5BC6\u4EA4\u6613", desc: "4/4 \u5F00\u653E BTC/ETH \u7B49\u5F85\u540D\u5355\uFF0CQ2\u6B63\u5F0F\u4E0A\u7EBF" },
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
  const fundData = [
    fundHeader,
    fundRow("ZeroStack/0G", "$107M", "\u751F\u6001\u6295\u8D44\uFF0C21%\u4F9B\u5E94\u91CF"),
    fundRow("Midas", "$50M", "Series A\uFF0CRWA\u6D41\u52A8\u6027\u5C42"),
    fundRow("Keyrock", "Series C", "\u4F30\u503C$11\u4EBF\uFF0CRipple\u53C2\u6295"),
    fundRow("Franklin/250D", "M&A", "\u7EC4\u5EFA Franklin Crypto"),
    fundRow("MicroStrategy", "44,377 BTC", "3\u6708\u589E\u6301\uFF0C\u5360\u4F01\u4E1A\u51C094%"),
  ];
  s4.addTable(fundData, {
    x: 5.3, y: 1.45, w: 4.1, colW: [1.3, 0.9, 1.9],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.3, 0.3, 0.3, 0.3, 0.3, 0.3],
    autoPage: false,
  });

  // Mining callout
  s4.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 3.4, w: 4.1, h: 0.7, fill: { color: "FFF8E7" }, shadow: makeShadow() });
  s4.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 3.4, w: 0.06, h: 0.7, fill: { color: C.gold } });
  s4.addText([
    { text: "\u5173\u6CE8\uFF1A", options: { bold: true, fontSize: 10, fontFace: "Arial", color: C.navy } },
    { text: "\u77FF\u4F01\u52A0\u901F\u629B\u552E BTC \u4E3A AI \u8F6C\u578B\u878D\u8D44\uFF1AMARA \u51FA\u552E 15,133 BTC\uFF0CRiot \u51FA\u552E $2.9\u4EBF BTC", options: { fontSize: 10, fontFace: "Arial", color: C.bodyText } }
  ], { x: 5.55, y: 3.45, w: 3.7, h: 0.6, margin: 0, valign: "middle" });

  s4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s4.addText("\u52A0\u5BC6\u5E02\u573A\u5468\u62A5 | 2026.03.30\u201304.05", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  s4.addNotes("\u76D1\u7BA1\u65B9\u9762\uFF0C\u672C\u5468\u662F\u7F8E\u56FD\u52A0\u5BC6\u76D1\u7BA1\u6846\u67B6\u843D\u5730\u7684\u5173\u952E\u4E00\u5468\u3002GENIUS Act \u5B9E\u65BD\u7EC6\u5219\u8981\u6C42\u7A33\u5B9A\u5E01\u53D1\u884C\u65B9\u7EF4\u6301 1:1 \u50A8\u5907\u548C 12 \u4E2A\u6708\u8FD0\u8425\u8D39\u7528\u7F13\u51B2\uFF0C\u8FD9\u662F\u7F8E\u56FD\u9996\u4E2A\u5168\u9762\u7684\u7A33\u5B9A\u5E01\u76D1\u7BA1\u6846\u67B6\u3002\u5BF9\u4E8E Tether \u548C Circle \u7B49\u5927\u578B\u53D1\u884C\u65B9\u800C\u8A00\u5408\u89C4\u6210\u672C\u53EF\u63A7\uFF0C\u4F46\u5C0F\u578B\u53D1\u884C\u65B9\u53EF\u80FD\u88AB\u8FEB\u9000\u51FA\u5E02\u573A\u3002\n\nCLARITY Act \u7A33\u5B9A\u5E01\u6536\u76CA\u59A5\u534F\u65B9\u6848\u662F\u91CD\u5927\u5229\u597D\uFF0C\u56E0\u4E3A\u5B83\u4FDD\u7559\u4E86\u57FA\u4E8E\u6D3B\u52A8\u7684\u6FC0\u52B1\u673A\u5236\uFF0C\u8FD9\u610F\u5473\u7740 DeFi \u534F\u8BAE\u4ECD\u53EF\u4EE5\u901A\u8FC7\u5408\u89C4\u65B9\u5F0F\u63D0\u4F9B\u7A33\u5B9A\u5E01\u6536\u76CA\u3002Circle \u80A1\u4EF7\u5728\u4E0A\u5468\u66B4\u8DCC 20% \u540E\uFF0C\u8FD9\u4E00\u59A5\u534F\u6D88\u606F\u53EF\u80FD\u63D0\u4F9B\u4E00\u5B9A\u7F13\u89E3\u3002\n\nCFTC \u65B0\u4EFB\u6267\u6CD5\u4E3B\u7BA1\u5BA3\u5E03\u4E0D\u518D\u4EE5\u8BC9\u8BBC\u8BBE\u5B9A\u653F\u7B56\uFF0C\u6807\u5FD7\u7740\u7F8E\u56FD\u6B63\u5F0F\u4ECE\u201C\u6267\u6CD5\u4EE3\u76D1\u7BA1\u201D\u6A21\u5F0F\u8F6C\u5411\u89C4\u5219\u5236\u5B9A\u6A21\u5F0F\u3002\n\nKraken \u83B7\u7F8E\u8054\u50A8\u4E3B\u8D26\u6237\u662F\u91CC\u7A0B\u7891\u4E8B\u4EF6\uFF0C\u610F\u5473\u7740\u52A0\u5BC6\u539F\u751F\u673A\u6784\u9996\u6B21\u83B7\u5F97\u4E0E\u4F20\u7EDF\u94F6\u884C\u540C\u7B49\u7684\u652F\u4ED8\u7CFB\u7EDF\u8BBF\u95EE\u6743\u9650\u3002Charles Schwab \u5F00\u653E\u52A0\u5BC6\u4EA4\u6613\u5219\u4EE3\u8868\u4F20\u7EDF\u5238\u5546\u5F00\u59CB\u5207\u5165\u96F6\u552E\u5E02\u573A\u3002\n\n\u6295\u878D\u8D44\u65B9\u9762\uFF0CZeroStack \u5BF9 0G \u751F\u6001\u7684 $1.07 \u4EBF\u6295\u8D44\u548C Franklin Templeton \u6536\u8D2D 250 Digital \u53CD\u6620\u4E86\u4F20\u7EDF\u91D1\u878D\u52A0\u901F\u5E03\u5C40\u52A0\u5BC6\u8D44\u4EA7\u57FA\u7840\u8BBE\u65BD\u7684\u8D8B\u52BF\u3002MicroStrategy 3\u6708\u589E\u6301 44,377 BTC\u5360\u4F01\u4E1A\u51C0\u589E\u91CF\u768494%\uFF0C\u663E\u793A\u5176\u6FC0\u8FDB\u7684 BTC \u7D2F\u79EF\u7B56\u7565\u4ECD\u5728\u7EE7\u7EED\u3002");

  // ============ SLIDE 5: Security + AI + Prediction Markets ============
  const s5 = pres.addSlide();
  s5.background = { color: C.offWhite };
  s5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy } });
  s5.addText("\uD83D\uDD13  \u5B89\u5168\u4E8B\u4EF6 & \u70ED\u70B9\u8D5B\u9053", {
    x: 0.6, y: 0, w: 9, h: 0.7,
    fontSize: 20, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0
  });

  // Security incidents - left
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 2.35, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 0.35, fill: { color: C.red } });
  s5.addText("\u5B89\u5168\u4E8B\u4EF6  |  \u5468\u635F\u5931: >$3.1\u4EBF", { x: 0.6, y: 0.9, w: 4.2, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  const secTable = [
    [
      { text: "\u4E8B\u4EF6", options: { bold: true, fontSize: 9, fontFace: "Arial", fill: { color: C.lightGray }, color: C.bodyText } },
      { text: "\u635F\u5931", options: { bold: true, fontSize: 9, fontFace: "Arial", fill: { color: C.lightGray }, color: C.bodyText, align: "right" } },
      { text: "\u65B9\u5F0F", options: { bold: true, fontSize: 9, fontFace: "Arial", fill: { color: C.lightGray }, color: C.bodyText } },
    ],
    [{ text: "Drift Protocol", options: { fontSize: 9, fontFace: "Arial" } }, { text: "$2.85\u4EBF", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red, bold: true } }, { text: "\u793E\u5DE5+\u9884\u8A00\u673A\u64CD\u7EB5", options: { fontSize: 9, fontFace: "Arial", color: C.red } }],
    [{ text: "Resolv \u5F71\u5B50\u4F20\u67D3", options: { fontSize: 9, fontFace: "Arial" } }, { text: "$2500\u4E07", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "AWS\u5BC6\u94A5\u6CC4\u9732+\u8FDE\u9501\u574F\u8D26", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
    [{ text: "SIREN Rug Pull", options: { fontSize: 9, fontFace: "Arial" } }, { text: ">$1\u4EBF", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "50%\u4F9B\u5E94\u96C6\u4E2D\u64CD\u63A7", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
    [{ text: "Venus Protocol", options: { fontSize: 9, fontFace: "Arial" } }, { text: "$218\u4E07", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "\u6DF7\u5408\u94FE\u4E0A/\u94FE\u4E0B\u653B\u51FB", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
  ];
  s5.addTable(secTable, {
    x: 0.55, y: 1.35, w: 4.2, colW: [1.4, 1.0, 1.8],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.28, 0.28, 0.28, 0.28, 0.28],
    autoPage: false,
  });

  // AI x Crypto - right top
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 4.5, h: 1.85, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s5.addText("AI \u00D7 Crypto", { x: 5.3, y: 0.9, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  const aiItems = [
    "AI\u4EE3\u5E01\u9006\u52BF\u4E0A\u6DA8\uFF1AFET/Virtuals +35\u201345%\uFF0CTAO +73%",
    "ZeroStack \u6295\u5165 $1.07\u4EBF\u5E03\u5C40 0G AI\u751F\u6001",
    "Render\u63A5\u5165 NVIDIA Blackwell B200\u67B6\u6784",
  ];
  s5.addText(aiItems.map((t, i) => ({
    text: t,
    options: { bullet: true, breakLine: i < aiItems.length - 1, fontSize: 10, fontFace: "Arial", color: C.bodyText }
  })), { x: 5.3, y: 1.35, w: 4.1, h: 1.3, margin: 0, paraSpaceAfter: 4 });

  // Prediction Markets - right bottom
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 2.95, w: 4.5, h: 1.85, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 2.95, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s5.addText("\u9884\u6D4B\u5E02\u573A", { x: 5.3, y: 2.95, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  const pmItems = [
    "\u884C\u4E1A3\u6708\u603B\u91CF $257\u4EBF\uFF0CKalshi $131\u4EBF\u9886\u5148",
    "Polymarket\u4F30\u503C$90\u4EBF\uFF0C\u83B7ICE $20\u4EBF\u6295\u8D44",
    "CFTC\u8D77\u8BC9\u4E09\u5DDE\uFF0C\u4E89\u593A\u9884\u6D4B\u5E02\u573A\u8054\u90A6\u7BA1\u8F96\u6743",
  ];
  s5.addText(pmItems.map((t, i) => ({
    text: t,
    options: { bullet: true, breakLine: i < pmItems.length - 1, fontSize: 10, fontFace: "Arial", color: C.bodyText }
  })), { x: 5.3, y: 3.4, w: 4.1, h: 1.3, margin: 0, paraSpaceAfter: 4 });

  // Stablecoin callout bar at bottom
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.5, w: 4.5, h: 1.5, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.5, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s5.addText("\u7A33\u5B9A\u5E01", { x: 0.6, y: 3.5, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });
  const stableItems = [
    "\u603B\u5E02\u503C ~$3250\u4EBF\uFF1B3\u6708\u4EA4\u6613\u91CF $7.5\u4E07\u4EBF\u8D85\u8D8A ACH",
    "USDC\u4F9B\u5E94+$20\u4EBF vs USDT-$30\u4EBF\uFF0C\u5206\u5316\u52A0\u5267",
  ];
  s5.addText(stableItems.map((t, i) => ({
    text: t,
    options: { bullet: true, breakLine: i < stableItems.length - 1, fontSize: 10, fontFace: "Arial", color: C.bodyText }
  })), { x: 0.6, y: 3.95, w: 4.1, h: 0.9, margin: 0, paraSpaceAfter: 4 });

  s5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s5.addText("\u52A0\u5BC6\u5E02\u573A\u5468\u62A5 | 2026.03.30\u201304.05", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  s5.addNotes("\u5B89\u5168\u4E8B\u4EF6\uFF1A\u672C\u5468\u662F 2026 \u5E74\u8FC4\u4ECA\u635F\u5931\u6700\u4E25\u91CD\u7684\u4E00\u5468\uFF0C\u603B\u635F\u5931\u8D85\u8FC7 $3.1 \u4EBF\u3002Drift Protocol $2.85 \u4EBF\u653B\u51FB\u662F\u5E74\u5EA6\u6700\u5927\u5355\u7B14\u4E8B\u4EF6\uFF0C\u653B\u51FB\u65B9\u5F0F\u6781\u4E3A\u590D\u6742\uFF1A\u516D\u4E2A\u6708\u793E\u5DE5\u6E17\u900F + \u9884\u8A00\u673A\u64CD\u7EB5 + \u591A\u7B7E\u5165\u4FB5\u3002\u8FD9\u63ED\u793A\u4E86\u5F53\u524D DeFi \u5B89\u5168\u6A21\u578B\u7684\u6839\u672C\u6027\u7F3A\u9677\u2014\u2014\u4EC5\u4F9D\u8D56\u94FE\u4E0A\u5B89\u5168\u5BA1\u8BA1\u662F\u4E0D\u591F\u7684\uFF0C\u94FE\u4E0B\u8FD0\u8425\u5B89\u5168\uFF08\u4EBA\u5458\u3001\u5BC6\u94A5\u7BA1\u7406\u3001\u591A\u7B7E\u6D41\u7A0B\uFF09\u540C\u6837\u5173\u952E\u3002\n\nResolv \u7684\u201C\u5F71\u5B50\u4F20\u67D3\u201D\u6548\u5E94\u66F4\u503C\u5F97\u8B66\u60D5\u2014\u2014\u5355\u4E2A\u7A33\u5B9A\u5E01\u7684\u8131\u951A\u901A\u8FC7\u62B5\u62BC\u54C1\u94FE\u5F15\u53D1 Morpho Blue\u3001Euler\u3001Fluid \u7B49\u591A\u4E2A\u534F\u8BAE\u7684\u574F\u8D26\uFF0C\u663E\u793A DeFi \u7EC4\u5408\u6027\u98CE\u9669\u4ECD\u7136\u88AB\u4F4E\u4F30\u3002\n\nAI \u00D7 Crypto \u8D5B\u9053\u672C\u5468\u8868\u73B0\u4EAE\u773C\uFF0C\u5728\u5927\u76D8\u6781\u5EA6\u6050\u60E7\u7684\u80CC\u666F\u4E0B AI \u4EE3\u5E01\u9006\u52BF\u4E0A\u6DA8\u3002TAO +73% \u7684\u6DA8\u5E45\u53CD\u6620\u4E86\u53BB\u4E2D\u5FC3\u5316\u673A\u5668\u5B66\u4E60\u83B7\u5F97\u4F01\u4E1A\u7EA7\u91C7\u7528\u7684\u53D9\u4E8B\u52A8\u529B\u3002ZeroStack \u5BF9 0G \u7684 $1.07 \u4EBF\u6295\u8D44\u662F\u4F20\u7EDF\u91D1\u878D\u673A\u6784\u5E03\u5C40\u53BB\u4E2D\u5FC3\u5316\u7B97\u529B\u7684\u91CC\u7A0B\u7891\u3002\n\n\u9884\u6D4B\u5E02\u573A\u884C\u4E1A\u89C4\u6A21\u7206\u53D1\u5F0F\u589E\u957F\uFF0C3\u6708\u603B\u91CF\u8FBE $257 \u4EBF\u3002\u4F46 CFTC \u4E0E\u5DDE\u653F\u5E9C\u7684\u7BA1\u8F96\u6743\u4E89\u593A\u5C06\u5728\u672A\u6765\u6570\u6708\u5185\u6301\u7EED\u5F71\u54CD\u884C\u4E1A\u53D1\u5C55\u65B9\u5411\u3002Polymarket \u83B7 ICE $20 \u4EBF\u6295\u8D44\u4F30\u503C\u8FBE $90 \u4EBF\uFF0C\u663E\u793A\u4F20\u7EDF\u91D1\u878D\u5BF9\u9884\u6D4B\u5E02\u573A\u7684\u8BA4\u53EF\u3002\n\n\u7A33\u5B9A\u5E01\u4EA4\u6613\u91CF 3\u6708\u8FBE $7.5 \u4E07\u4EBF\u8D85\u8D8A ACH \u7F51\u7EDC\u662F\u5386\u53F2\u6027\u65F6\u523B\uFF0C\u8BC1\u660E\u7A33\u5B9A\u5E01\u5DF2\u4ECE\u52A0\u5BC6\u539F\u751F\u5DE5\u5177\u5347\u7EA7\u4E3A\u4E3B\u6D41\u652F\u4ED8\u57FA\u7840\u8BBE\u65BD\u3002USDC \u4E0E USDT \u7684\u4F9B\u5E94\u5206\u5316\u52A0\u5267\uFF0C\u53CD\u6620\u4E86\u5E02\u573A\u5BF9\u900F\u660E\u5EA6\u548C\u5408\u89C4\u6027\u7684\u504F\u597D\u8F6C\u53D8\u3002");

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
  const catData = [
    catHeader,
    catRow("4/6", "Hyperliquid \u4EE3\u5E01\u89E3\u9501 ~2.65%", "HYPE", "\u4F9B\u7ED9\u538B\u529B", true),
    catRow("4/12", "Aptos \u89E3\u9501 1130\u4E07 APT", "APT", "\u4F9B\u7ED9\u538B\u529B ~0.68%"),
    catRow("4/15-16", "Paris Blockchain Week", "\u5168\u5E02\u573A", "MiCA\u8BA8\u8BBA+\u673A\u6784\u91C7\u7528"),
    catRow("4\u6708\u4E2D", "CLARITY Act \u59D4\u5458\u4F1A\u5BA1\u8BAE", "\u7A33\u5B9A\u5E01", "\u6536\u76CA\u7981\u4EE4\u5B9A\u8C03", true),
    catRow("4\u6708", "91\u9879 ETF \u7533\u8BF7\u540E\u7EED", "\u591A\u8D44\u4EA7", "SOL/XRP/LTC \u5BA1\u6279\u52A8\u6001"),
    catRow("4/20", "LayerZero \u89E3\u9501 2570\u4E07 ZRO", "ZRO", "\u4F9B\u7ED9\u538B\u529B"),
    catRow("4/20", "\u9999\u6E2F Web3 Festival", "\u5168\u5E02\u573A", "\u4E9A\u6D32\u5E74\u5EA6\u5CF0\u4F1A"),
  ];
  s6.addTable(catData, {
    x: 0.5, y: 0.9, w: 9, colW: [0.9, 3.2, 1.2, 3.7],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.32, 0.32, 0.32, 0.32, 0.32, 0.32, 0.32, 0.32],
    autoPage: false,
  });

  // Outlook box
  s6.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.65, w: 9, h: 1.3, fill: { color: C.white }, shadow: makeShadow() });
  s6.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.65, w: 0.06, h: 1.3, fill: { color: C.gold } });
  s6.addText("\u672C\u5468\u5C0F\u7ED3", { x: 0.8, y: 3.7, w: 8, h: 0.3, fontSize: 13, fontFace: "Arial", bold: true, color: C.navy, margin: 0 });
  s6.addText([
    { text: "\u5E02\u573A\u60C5\u7EEA\u7EF4\u6301\u201C\u6781\u5EA6\u6050\u60E7\u201D\uFF08FGI=12\uFF09\uFF0CBTC \u7A84\u5E45\u9707\u8361\uFF0C\u5C71\u5BE8\u5E01\u666E\u8DCC\u3002Drift $2.85\u4EBF\u653B\u51FB\u4E3A\u5E74\u5EA6\u6700\u5927\u5B89\u5168\u4E8B\u4EF6\u3002\u76D1\u7BA1\u6846\u67B6\u52A0\u901F\u843D\u5730\uFF08GENIUS Act\u3001CLARITY Act\u3001Schwab\u5F00\u653E\u52A0\u5BC6\u4EA4\u6613\uFF09\u3002", options: { breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "\u91CD\u70B9\u5173\u6CE8\uFF1A", options: { bold: true, fontSize: 10, fontFace: "Arial", color: C.navy } },
    { text: " 4/6 HYPE \u89E3\u9501\uFF08\u6301\u4ED3\u76F8\u5173\uFF09\u3001CLARITY Act \u59D4\u5458\u4F1A\u5BA1\u8BAE\u3001AI\u8D5B\u9053\u9006\u52BF\u8868\u73B0\u3001Drift\u653B\u51FB\u540E\u7EED\u5F71\u54CD\u3002", options: { fontSize: 10, fontFace: "Arial", color: C.bodyText } },
  ], { x: 0.8, y: 4.0, w: 8.5, h: 0.85, margin: 0 });

  s6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s6.addText("\u52A0\u5BC6\u5E02\u573A\u5468\u62A5 | 2026.03.30\u201304.05  |  CONFIDENTIAL", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  s6.addNotes("\u672A\u6765\u4E24\u5468\u50AC\u5316\u5242\u5206\u6790\uFF1A\n\n1. HYPE \u89E3\u9501\uFF084/6\uFF09\u2014\u2014 \u7EA6 2.65% \u6D41\u901A\u91CF\u89E3\u9501\uFF0C\u5BF9\u6211\u4EEC\u7684\u6301\u4ED3\u76F4\u63A5\u76F8\u5173\u3002\u5728\u5F53\u524D\u6781\u5EA6\u6050\u60E7\u7684\u5E02\u573A\u73AF\u5883\u4E0B\uFF0C\u89E3\u9501\u538B\u529B\u53EF\u80FD\u88AB\u653E\u5927\u3002\u4F46 Hyperliquid \u8FDB\u5165\u5168\u7403\u884D\u751F\u54C1\u524D\u5341\u548C Grayscale ETF S-1 \u7533\u62A5\u63D0\u4F9B\u4E86\u6B63\u9762\u53D9\u4E8B\u5BF9\u51B2\u3002\n\n2. CLARITY Act \u59D4\u5458\u4F1A\u5BA1\u8BAE\uFF084\u6708\u4E2D\u65EC\uFF09\u2014\u2014 \u7A33\u5B9A\u5E01\u6536\u76CA\u7981\u4EE4\u7684\u6700\u7EC8\u5F62\u6001\u5C06\u5BF9 Circle\u3001Coinbase \u7B49\u4F01\u4E1A\u4EA7\u751F\u6DF1\u8FDC\u5F71\u54CD\u3002\u5982\u679C\u5982\u5E02\u573A\u9884\u671F\u822C\u901A\u8FC7\uFF0C\u5C06\u4E3A\u7A33\u5B9A\u5E01\u884C\u4E1A\u63D0\u4F9B\u671F\u5F85\u5DF2\u4E45\u7684\u76D1\u7BA1\u786E\u5B9A\u6027\u3002\n\n3. 91\u9879 ETF \u7533\u8BF7\u540E\u7EED\u2014\u2014 3/27 \u622A\u6B62\u540E\uFF0CSOL\u3001XRP\u3001LTC\u3001DOGE \u73B0\u8D27 ETF \u7684\u5BA1\u6279\u52A8\u6001\u5C06\u663E\u8457\u5F71\u54CD\u76F8\u5173\u8D44\u4EA7\u4EF7\u683C\u3002\u673A\u6784\u8D44\u91D1\u901A\u9053\u6253\u5F00\u5C06\u4E3A\u5C71\u5BE8\u5E01\u5E02\u573A\u5E26\u6765\u7ED3\u6784\u6027\u53D8\u5316\u3002\n\n\u603B\u4F53\u8BC4\u4F30\uFF1A\u5E02\u573A\u5904\u4E8E FGI=12 \u7684\u6781\u5EA6\u6050\u60E7\u533A\u95F4\uFF0C\u4F46\u57FA\u672C\u9762\u6539\u5584\u4FE1\u53F7\u660E\u786E\u2014\u2014\u76D1\u7BA1\u6846\u67B6\u52A0\u901F\u843D\u5730\u3001\u4F20\u7EDF\u91D1\u878D\u673A\u6784\u52A0\u901F\u5165\u573A\uFF08Schwab\u3001Franklin Templeton\uFF09\u3001DeFi \u534F\u8BAE\u5411\u673A\u6784\u5316\u8FDB\u5316\uFF08Aave V4\uFF09\u3001AI \u8D5B\u9053\u7206\u53D1\u3002\u60C5\u7EEA\u4E0E\u57FA\u672C\u9762\u7684\u80CC\u79BB\u8868\u660E\u6211\u4EEC\u53EF\u80FD\u5904\u4E8E\u5468\u671F\u540E\u671F\u6050\u60E7\u9636\u6BB5\uFF0C\u5386\u53F2\u4E0A\u8FD9\u5F80\u5F80\u662F\u8F83\u597D\u7684\u903E\u671F\u5E03\u5C40\u7A97\u53E3\u3002\u4FDD\u6301\u5F53\u524D\u4ED3\u4F4D\uFF0C\u91CD\u70B9\u76D1\u63A7 HYPE \u89E3\u9501\u548C Drift \u653B\u51FB\u540E\u7EED\u5F71\u54CD\u3002");

  const outPath = "/Users/lx/Github/crypto-weekly-digest/reports/2026-03-30/slides-zh.pptx";
  await pres.writeFile({ fileName: outPath });
  console.log("\u2705 Written to " + outPath);
}

generate().catch(e => { console.error(e); process.exit(1); });
