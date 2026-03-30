const pptxgen = require("pptxgenjs");

const C = {
  navy: "1A2744", darkNavy: "0F1A2E", gold: "C9A96E", white: "FFFFFF",
  offWhite: "F5F6F8", lightGray: "E8EBF0", medGray: "8C95A6",
  bodyText: "2D3748", red: "C53030", green: "2F855A",
};
const makeShadow = () => ({ type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.12 });
const WEEK = "2026.03.23\u201303.29";
const FOOTER_TEXT = `\u52A0\u5BC6\u5E02\u573A\u5468\u62A5 | ${WEEK}`;

async function generate() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "42";
  pres.title = `\u52A0\u5BC6\u5E02\u573A\u5468\u62A5 ${WEEK}`;

  // ============ SLIDE 1: Title ============
  const s1 = pres.addSlide();
  s1.background = { color: C.darkNavy };
  s1.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.6, w: 1.2, h: 0.04, fill: { color: C.gold } });
  s1.addText("\u52A0\u5BC6\u5E02\u573A\u5468\u62A5", {
    x: 0.8, y: 1.8, w: 8, h: 1.0,
    fontSize: 42, fontFace: "Arial", bold: true, color: C.white, margin: 0
  });
  s1.addText("2026\u5E743\u670823\u65E5 \u2014 3\u670829\u65E5", {
    x: 0.8, y: 2.7, w: 8, h: 0.5,
    fontSize: 20, fontFace: "Arial", color: C.gold, margin: 0
  });
  s1.addText("\u6570\u636E\u6765\u6E90\uFF1AMessari \u00B7 The Block \u00B7 CoinGecko \u00B7 DefiLlama", {
    x: 0.8, y: 3.4, w: 8, h: 0.4,
    fontSize: 11, fontFace: "Arial", color: C.medGray, margin: 0
  });
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.gold } });
  s1.addText("CONFIDENTIAL \u2014 FOR LP DISTRIBUTION ONLY", {
    x: 0, y: 5.25, w: 10, h: 0.375,
    fontSize: 9, fontFace: "Arial", color: C.darkNavy, align: "center", valign: "middle", bold: true
  });
  s1.addNotes("本周报覆盖2026年3月23日至29日。市场在地缘政治风险（伊以冲突升级、霍尔木兹海峡航运受阻）和特朗普关税政策加码的双重压力下，整体风险偏好极度低迷。恐惧贪婪指数降至8，为2022年Luna崩盘以来最低水平。\n\n核心驱动因素：\n1. 美联储将2026年PCE通胀预测上调至2.7%，降息预期进一步推迟\n2. 伊朗威胁封锁曼德海峡，油价较冲突前暴涨25%+\n3. 特朗普将全球关税提至15%\n4. BTC/ETH/SOL ETF同日全线净流出，机构避险情绪明显\n\n但监管端出现重大利好：SEC-CFTC联合规则正式将16种加密资产认定为数字商品，终结了长达十年的监管模糊地带。这为SOL/XRP现货ETF扫清了路径障碍。");

  // ============ SLIDE 2: Market Overview ============
  const s2 = pres.addSlide();
  s2.background = { color: C.offWhite };
  s2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy } });
  s2.addText("\uD83D\uDCCA  \u5E02\u573A\u6982\u89C8", {
    x: 0.6, y: 0, w: 9, h: 0.7,
    fontSize: 20, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0
  });

  // 3 metric cards
  const cardY = 0.95, cardH = 1.05, cardW = 2.8, cardGap = 0.3, cardStartX = 0.6;

  // BTC card
  s2.addShape(pres.shapes.RECTANGLE, { x: cardStartX, y: cardY, w: cardW, h: cardH, fill: { color: C.white }, shadow: makeShadow() });
  s2.addShape(pres.shapes.RECTANGLE, { x: cardStartX, y: cardY, w: 0.06, h: cardH, fill: { color: C.gold } });
  s2.addText("BTC", { x: cardStartX + 0.25, y: cardY + 0.1, w: 2, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.medGray, bold: true, margin: 0 });
  s2.addText("$66,445", { x: cardStartX + 0.25, y: cardY + 0.35, w: 2, h: 0.4, fontSize: 28, fontFace: "Arial", color: C.bodyText, bold: true, margin: 0 });
  s2.addText("-3.5%", { x: cardStartX + 0.25, y: cardY + 0.72, w: 2, h: 0.25, fontSize: 14, fontFace: "Arial", color: C.red, bold: true, margin: 0 });

  // ETH card
  const c2x = cardStartX + cardW + cardGap;
  s2.addShape(pres.shapes.RECTANGLE, { x: c2x, y: cardY, w: cardW, h: cardH, fill: { color: C.white }, shadow: makeShadow() });
  s2.addShape(pres.shapes.RECTANGLE, { x: c2x, y: cardY, w: 0.06, h: cardH, fill: { color: C.gold } });
  s2.addText("ETH", { x: c2x + 0.25, y: cardY + 0.1, w: 2, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.medGray, bold: true, margin: 0 });
  s2.addText("$1,996", { x: c2x + 0.25, y: cardY + 0.35, w: 2, h: 0.4, fontSize: 28, fontFace: "Arial", color: C.bodyText, bold: true, margin: 0 });
  s2.addText("-4.1%", { x: c2x + 0.25, y: cardY + 0.72, w: 2, h: 0.25, fontSize: 14, fontFace: "Arial", color: C.red, bold: true, margin: 0 });

  // Fear & Greed card
  const c3x = c2x + cardW + cardGap;
  s2.addShape(pres.shapes.RECTANGLE, { x: c3x, y: cardY, w: cardW, h: cardH, fill: { color: C.white }, shadow: makeShadow() });
  s2.addShape(pres.shapes.RECTANGLE, { x: c3x, y: cardY, w: 0.06, h: cardH, fill: { color: C.red } });
  s2.addText("\u6050\u60E7\u8D2A\u5A6A\u6307\u6570", { x: c3x + 0.25, y: cardY + 0.1, w: 2.4, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.medGray, bold: true, margin: 0 });
  s2.addText("9", { x: c3x + 0.25, y: cardY + 0.35, w: 1, h: 0.4, fontSize: 28, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
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
    makeRow("SOL", "$81.63", "-6.7%", "SEC\u5546\u54C1\u5206\u7C7B\u5229\u597D\u5C1A\u672A\u53CD\u6620", true),
    makeRow("HYPE", "$38.91", "+2.3%", "\u6C38\u7EED DEX \u9F99\u5934\uFF1B\u5468\u8D39\u7528$14M", false),
    makeRow("PENDLE", "$1.12", "-6.2%", "\u6536\u76CA\u4EE3\u5E01\u8D5B\u9053\u627F\u538B", true),
    makeRow("ENA", "$0.09", "-8.4%", "\u968F\u5927\u76D8\u8D70\u5F31", true),
    makeRow("DYDX", "$0.09", "+11.6%", "\u56DE\u8D2D\u8BA1\u5212\u6269\u81F375%\u51C0\u6536\u5165", false),
  ];
  s2.addTable(tableData, {
    x: 0.6, y: tableY, w: 8.8, colW: [1.3, 1.8, 1.5, 4.2],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.35, 0.33, 0.33, 0.33, 0.33, 0.33],
    autoPage: false,
  });

  // Macro note
  s2.addText("\u5B8F\u89C2\u9A71\u52A8\uFF1A\u4F0A\u4EE5\u51B2\u7A81\u5347\u7EA7 + \u7279\u6717\u666E\u5173\u7A0E15% + \u7F8E\u8054\u50A8PCE\u4E0A\u8C032.7% \u2192 \u98CE\u9669\u8D44\u4EA7\u666E\u904D\u627F\u538B\uFF1BBTC/ETH/SOL ETF\u540C\u65E5\u5168\u7EBF\u51C0\u6D41\u51FA", {
    x: 0.6, y: 4.55, w: 8.8, h: 0.4,
    fontSize: 10, fontFace: "Arial", color: C.medGray, italic: true, margin: 0
  });

  s2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s2.addText(FOOTER_TEXT, { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  s2.addNotes("市场概览深度分析：\n\n1. BTC跌至$66,445，周跌3.5%：核心驱动是3月27日$141.6亿BTC期权到期，大量空头押注得到兑现。ETF单周净流出$296M，机构资金转为防御姿态。技术面上，$66,000是关键支撑位，若日线收盘跌破，可能打开向$50,000的下行通道。\n\n2. ETH跌破$2,000心理关口：Fusaka升级后L1费用收入大幅下降，L2对主网的活动虹吸效应加剧。某创始人向Kraken转入约79,000 ETH增加了直接抛压。距8月高点已下跌60%，ETH/BTC汇率持续走弱。\n\n3. 恐惧贪婪指数降至8-9，为极端恐惧区间：这一水平在历史上通常预示着中期底部区域，但需要宏观环境配合才能确认反转。当前伊以冲突未见缓和迹象，关税政策悬而未决，短期内可能维持低迷。\n\n4. 总市值$2.37T，BTC主导率56.1%：山寨币表现普遍弱于BTC，资金向BTC集中避险。SOL虽获SEC商品分类利好，但价格尚未反映，说明宏观风险暂时压过了基本面改善。\n\n5. HYPE逆势微涨2.3%：Hyperliquid作为永续DEX龙头表现出韧性，周费用收入$14M创新高，年化超$6亿。这验证了\u201C有真实收入的协议在熊市中更抗跌\u201D的逻辑。");

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
    ["1", "SEC-CFTC\u8054\u5408\u89C4\u5219\uFF1A16\u79CD\u52A0\u5BC6\u8D44\u4EA7\u83B7\u6570\u5B57\u5546\u54C1\u5206\u7C7B", "\u7EC8\u7ED3\u5341\u5E74\u76D1\u7BA1\u6A21\u7CCA\u5730\u5E26\uFF0C\u4E3ASOL/XRP ETF\u626B\u6E05\u969C\u788D"],
    ["2", "CLARITY Act\u7981\u6B62\u7A33\u5B9A\u5E01\u6536\u76CA\uFF0CCircle\u80A1\u4EF7\u66B4\u8DCC20%", "\u5E02\u503C\u84B8\u53D1$20\u4EBF\uFF0C\u7A33\u5B9A\u5E01\u5546\u4E1A\u6A21\u5F0F\u9762\u4E34\u91CD\u5851"],
    ["3", "Resolv USR\u88AB\u653B\u51FB\uFF1A$2500\u4E07\u88AB\u76D7\u3001\u8131\u951A97.5%", "AWS\u5BC6\u94A5\u88AB\u5165\u4FB5\uFF0C17\u5206\u949F\u5185\u5D29\u76D8"],
    ["4", "\u5730\u7F18\u51B2\u7A81+\u5173\u7A0E\u538B\u5757\u5E02\u573A", "\u4F0A\u4EE5\u6218\u4E89+\u7279\u6717\u666E15%\u5173\u7A0E\uFF0CETF\u540C\u65E5\u5168\u7EBF\u51C0\u6D41\u51FA"],
    ["5", "BTC\u77FF\u4F01\u96C6\u4F53\u8F6C\u578BAI\uFF1A\u6316\u77FF\u6210\u672C$80K vs \u5E01\u4EF7$68K", "\u5DF2\u7B7E\u7EA6$700\u4EBFAI\u5408\u540C\uFF0C\u51FA\u552E15,000+BTC\u878D\u8D44"],
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
    "16\u79CD\u4EE3\u5E01\u83B7SEC-CFTC\u6570\u5B57\u5546\u54C1\u8BA4\u5B9A\uFF08BTC/ETH/SOL\u7B49\uFF09",
    "Base TVL $41.5\u4EBF\uFF0C\u5360L2 DEX\u4EA4\u6613\u91CF\u7EA650%",
    "Linea \u63A8\u51FA Yield Boost\uFF0CETH\u81EA\u52A8\u901A\u8FC7Lido V3\u8D28\u62BC",
    "L2/L1\u65E5\u6D3B\u6BD4\u964D\u81F31.12\uFF0C\u8F83\u5CE8\u503C10.43\u5927\u5E45\u56DE\u843D",
  ];
  s3.addText(l1Items.map((t, i) => ({
    text: t,
    options: { bullet: true, breakLine: i < l1Items.length - 1, fontSize: 10, fontFace: "Arial", color: C.bodyText }
  })), { x: 5.3, y: 1.32, w: 4.1, h: 1.45, margin: 0, paraSpaceAfter: 4 });

  // Right bottom: DeFi + Perp DEX
  s3.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 3.05, w: 4.5, h: 1.95, fill: { color: C.white }, shadow: makeShadow() });
  s3.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 3.05, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s3.addText("DeFi & \u6C38\u7EED\u5408\u7EA6", { x: 5.3, y: 3.05, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  const defiItems = [
    "Aave\u7D2F\u8BA1\u501F\u8D37\u7A81\u7834$1\u4E07\u4EBF\uFF0CTVL $272.9\u4EBF\uFF0C\u5360\u501F\u8D3762.8%",
    "DeFi TVL\u9006\u52BF\u4E0A\u6DA84.4%\uFF0C\u94FE\u4E0A\u8D44\u91D1\u97E7\u6027\u5F3A",
    "Hyperliquid\u536054%\u6C38\u7EEDDEX\u5E02\u573A\uFF0C\u5468\u8D39$14M\uFF08+56%\uFF09",
    "dYdX\u56DE\u8D2D\u6269\u81F375%\u51C0\u6536\u5165\uFF0C\u62DF\u63A8\u5408\u6210\u80A1\u7968\u5408\u7EA6",
  ];
  s3.addText(defiItems.map((t, i) => ({
    text: t,
    options: { bullet: true, breakLine: i < defiItems.length - 1, fontSize: 10, fontFace: "Arial", color: C.bodyText }
  })), { x: 5.3, y: 3.47, w: 4.1, h: 1.45, margin: 0, paraSpaceAfter: 4 });

  s3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s3.addText(FOOTER_TEXT, { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  s3.addNotes("本周要闻深度解读：\n\n1. SEC-CFTC联合规则是加密行业监管里程碑。该规则于3月17日发布、3月23日生效，将BTC、ETH、SOL、XRP、ADA等16种代币认定为数字商品，由CFTC监管。SEC主席Paul Atkins正式承认\u201C大多数加密资产本身不是证券\u201D。这不仅终结了\u201C以执法代监管\u201D的时代，更为SOL和XRP现货ETF铺平了道路。对行业影响：项目方法律风险大幅降低，机构配置意愿提升。\n\n2. CLARITY Act修订草案冲击稳定币赛道。参议院银行委员会拟禁止稳定币被动收益（\u201Ceconomically equivalent to interest\u201D），直接打击Circle的商业模式——Circle 95%收入来自储备利息。股价暴跌20%，Coinbase连带下跌10%。后续发展：4月下旬全面委员会审议，如果通过将重塑整个稳定币行业的盈利模式。\n\n3. Resolv USR攻击暴露了DeFi私钥管理的根本性风险。攻击者入侵AWS KMS中的特权私钥，以不足$20万抵押铸造8000万USR，17分钟内在Curve上脱锚至$0.025。此事件还引发Morpho上的连锁反应——多个借贷金库因预言机仍以$1.13估值USR而被掏空。\n\nDeFi方面：Aave累计借贷突破$1万亿是DeFi信贷里程碑。但ACI退出治理事件反映了DAO治理模式的深层矛盾——投票权集中化问题可能成为DeFi下一个需要解决的系统性问题。ECB点名Aave、MakerDAO、Uniswap前100持币者控制超80%投票权，可能面临MiCA式监管。\n\nHyperliquid周费用$14M（+56%），年化超$6亿，原油永续合约日交易量超$22亿。这说明永续DEX的增长已不再局限于加密资产，正在向大宗商品和传统金融资产扩展。");

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
    { title: "SEC-CFTC\u8054\u5408\u6570\u5B57\u5546\u54C1\u89C4\u5219", desc: "16\u79CD\u4EE3\u5E01\u83B7\u5546\u54C1\u8BA4\u5B9A\uFF0C\u7EC8\u7ED3\u201C\u4EE5\u6267\u6CD5\u4EE3\u76D1\u7BA1\u201D\u6A21\u5F0F" },
    { title: "CLARITY Act\u4FEE\u8BA2\u8349\u6848", desc: "\u7981\u6B62\u7A33\u5B9A\u5E01\u88AB\u52A8\u6536\u76CA\uFF0CCircle\u80A1\u4EF7\u66B4\u8DCC20%" },
    { title: "CFTC\u521B\u65B0\u5DE5\u4F5C\u7EC4\u6210\u7ACB", desc: "\u6DB5\u76D6\u6570\u5B57\u8D44\u4EA7\u3001AI\u3001\u9884\u6D4B\u5E02\u573A\u76D1\u7BA1\u6846\u67B6" },
    { title: "\u4F17\u8BAE\u9662\u4EE3\u5E01\u5316\u542C\u8BC1\u4F1A", desc: "\u91D1\u878D\u670D\u52A1\u59D4\u8BA8\u8BBA\u533A\u5757\u94FE\u5BF9\u8D44\u672C\u5E02\u573A\u7684\u73B0\u4EE3\u5316\u4F5C\u7528" },
    { title: "NYSE\u00D7Securitize\u6570\u5B57\u4EA4\u6613\u5E73\u53F0", desc: "SEC\u6279\u51C6\u4EE3\u5E01\u5316\u7F8E\u80A1\u768424/7\u4EA4\u6613+\u94FE\u4E0A\u7A33\u5B9A\u5E01\u7ED3\u7B97" },
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
    fundRow("Mastercard\u00D7BVNK", "$18\u4EBF", "\u4F20\u7EDF\u91D1\u878D\u6DF1\u5165\u52A0\u5BC6\u57FA\u5EFA"),
    fundRow("Kalshi E\u8F6E", "$10\u4EBF", "\u4E8B\u4EF6\u4EA4\u6613\u5E73\u53F0\u4F30\u503C\u521B\u65B0\u9AD8"),
    fundRow("Polymarket", "$6\u4EBF", "ICE\u9886\u6295\uFF0C\u4F20\u7EDF\u91D1\u878D\u80CC\u4E66"),
    fundRow("Core Scientific", "$10\u4EBF", "JPM+\u5927\u6469\u6839\u4FE1\u8D37\uFF0CAI\u82AF\u7247\u6269\u5F20"),
    fundRow("GSR\u6536\u8D2D", "$5700\u4E07", "\u6536\u8D2DAutonomous+Architech"),
  ];
  s4.addTable(fundData, {
    x: 5.3, y: 1.45, w: 4.1, colW: [1.4, 0.8, 1.9],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.3, 0.3, 0.3, 0.3, 0.3, 0.3],
    autoPage: false,
  });

  // Callout
  s4.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 3.4, w: 4.1, h: 0.7, fill: { color: "FFF8E7" }, shadow: makeShadow() });
  s4.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 3.4, w: 0.06, h: 0.7, fill: { color: C.gold } });
  s4.addText([
    { text: "\u5173\u6CE8\uFF1A", options: { bold: true, fontSize: 10, fontFace: "Arial", color: C.navy } },
    { text: "\u672C\u5468\u62AB\u9732\u603B\u989D\u7EA6$32.8\u4EBF\uFF0C\u4F20\u7EDF\u91D1\u878D\u673A\u6784\u52A0\u901F\u5E03\u5C40\u52A0\u5BC6\u57FA\u5EFA\uFF0C\u9884\u8BA12026\u5E74M&A\u6709\u671B\u7A81\u7834$370\u4EBF", options: { fontSize: 10, fontFace: "Arial", color: C.bodyText } }
  ], { x: 5.55, y: 3.45, w: 3.7, h: 0.6, margin: 0, valign: "middle" });

  s4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s4.addText(FOOTER_TEXT, { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  s4.addNotes("监管与投融资深度分析：\n\n监管方面：\n1. SEC-CFTC联合规则是美国加密监管史上最重要的里程碑之一。它建立了五类资产分类体系（数字商品、数字收藏品、数字工具、稳定币、数字证券），并为挖矿、质押、无对价空投创建了明确的安全港条款。这意味着项目方不再需要猜测自己的代币是否会被SEC起诉为证券。对行业的长期影响是降低合规成本、吸引更多传统资金入场。\n\n2. CLARITY Act修订草案拟禁止稳定币被动收益，这对Circle是毁灭性打击——其95%收入来自储备利息。如果通过，Circle要么彻底改变商业模式，要么面临收入归零。Coinbase作为Circle的主要分销渠道也受到牵连。但银行业对此乐见其成，因为稳定币收益直接威胁传统存款业务。这场博弈的结果将决定稳定币是成为\u201C数字现金\u201D还是\u201C数字储蓄账户\u201D。\n\n3. NYSE与Securitize合作推出数字交易平台，SEC批准代币化美股24/7交易+链上结算，标志着传统金融基础设施正在主动拥抱区块链技术。\n\n投融资方面：\n1. Mastercard以约$18亿收购BVNK是本周最大交易，也是传统支付巨头最大规模的加密收购之一。这说明传统金融不再只是\u201C观望\u201D，而是在积极并购加密基础设施。\n\n2. Kalshi $10亿E轮和Polymarket $6亿融资（ICE领投）表明预测市场赛道获得了空前的机构认可。ICE是纽交所的母公司，其投资Polymarket具有重大信号意义。\n\n3. 本周22笔交易披露金额约$32.8亿，2026年加密M&A预计将突破$370亿创纪录。这反映了行业整合加速和传统金融资本大规模涌入。");

  // ============ SLIDE 5: Security + Hot Sectors ============
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
  s5.addText("\u5B89\u5168\u4E8B\u4EF6  |  YTD\u603B\u635F\u5931: ~$4.8\u4EBF", { x: 0.6, y: 0.9, w: 4.2, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  const secTable = [
    [
      { text: "\u4E8B\u4EF6", options: { bold: true, fontSize: 9, fontFace: "Arial", fill: { color: C.lightGray }, color: C.bodyText } },
      { text: "\u635F\u5931", options: { bold: true, fontSize: 9, fontFace: "Arial", fill: { color: C.lightGray }, color: C.bodyText, align: "right" } },
      { text: "\u65B9\u5F0F", options: { bold: true, fontSize: 9, fontFace: "Arial", fill: { color: C.lightGray }, color: C.bodyText } },
    ],
    [{ text: "\u4F0A\u6717\u4EA4\u6613\u6240", options: { fontSize: 9, fontFace: "Arial" } }, { text: "~$9000\u4E07", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "\u7591\u56FD\u5BB6\u7EA7\u653B\u51FB", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
    [{ text: "Resolv USR", options: { fontSize: 9, fontFace: "Arial" } }, { text: "~$2500\u4E07", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "AWS\u5BC6\u94A5\u5165\u4FB5", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
    [{ text: "Infiniti/GhostClaw", options: { fontSize: 9, fontFace: "Arial" } }, { text: "\u672A\u62AB\u9732", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "\u6076\u610Fnpm\u5305+\u94B1\u5305\u6392\u5E72", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
    [{ text: "\u865A\u5047\u7A33\u5B9A\u5E01", options: { fontSize: 9, fontFace: "Arial" } }, { text: "54,000+\u679A", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "\u94F6\u884C\u6388\u6743\u8BC8\u9A97", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
  ];
  s5.addTable(secTable, {
    x: 0.55, y: 1.35, w: 4.2, colW: [1.4, 1.1, 1.7],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.28, 0.28, 0.28, 0.28, 0.28],
    autoPage: false,
  });

  // Stablecoins - left bottom
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.5, w: 4.5, h: 1.5, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.5, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s5.addText("\u7A33\u5B9A\u5E01", { x: 0.6, y: 3.5, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });
  const stableItems = [
    "USDC\u4EA4\u6613\u91CF\u9996\u8D85USDT\uFF08\u5360\u6BD464%\uFF09\uFF1B86%\u673A\u6784\u6301\u6709USDC",
    "CLARITY Act\u62DF\u7981\u7A33\u5B9A\u5E01\u6536\u76CA\uFF0CTether\u542F\u52A8\u5168\u9762\u5BA1\u8BA1\u53CD\u51FB",
  ];
  s5.addText(stableItems.map((t, i) => ({
    text: t,
    options: { bullet: true, breakLine: i < stableItems.length - 1, fontSize: 10, fontFace: "Arial", color: C.bodyText }
  })), { x: 0.6, y: 3.95, w: 4.1, h: 0.9, margin: 0, paraSpaceAfter: 4 });

  // AI x Crypto - right top
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 4.5, h: 1.85, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s5.addText("AI \u00D7 Crypto", { x: 5.3, y: 0.9, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  const aiItems = [
    "NVIDIA GTC\u4F59\u9707\uFF1A\u9EC4\u4EC1\u52CB\u5BA3\u5E03\u201C\u5DF2\u8FBE AGI \u6C34\u5E73\u201D",
    "AI\u52A0\u5BC6\u677F\u5757\u5E02\u503C\u8D85$280\u4EBF\uFF0CTAO $34\u4EBF\u9886\u8DD1",
    "\u77FF\u4F01\u632F\u52A8\u6210\u672C$80K vs \u5E01\u4EF7$68K\uFF0C\u5DF2\u7B7E$700\u4EBFAI\u5408\u540C",
  ];
  s5.addText(aiItems.map((t, i) => ({
    text: t,
    options: { bullet: true, breakLine: i < aiItems.length - 1, fontSize: 10, fontFace: "Arial", color: C.bodyText }
  })), { x: 5.3, y: 1.35, w: 4.1, h: 1.3, margin: 0, paraSpaceAfter: 4 });

  // Prediction Markets - right bottom
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 2.95, w: 4.5, h: 2.05, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 2.95, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s5.addText("\u9884\u6D4B\u5E02\u573A", { x: 5.3, y: 2.95, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  const pmItems = [
    "Polymarket: 11,176\u4E2A\u6D3B\u8DC3\u5E02\u573A\uFF0C\u670884\u4E07\u6D3B\u8DC3\u94B1\u5305",
    "\u6708\u5EA6\u603B\u91CF\u7A81\u7834$210\u4EBF\uFF0CKalshi\u5B8C\u6210$10\u4EBBE\u8F6E",
    "CFTC\u5C06Polymarket\u7EB3\u5165\u76D1\u7BA1\u6846\u67B6\uFF0C\u53CD\u5185\u5E55\u4EA4\u6613\u65B0\u89C4\u751F\u6548",
  ];
  s5.addText(pmItems.map((t, i) => ({
    text: t,
    options: { bullet: true, breakLine: i < pmItems.length - 1, fontSize: 10, fontFace: "Arial", color: C.bodyText }
  })), { x: 5.3, y: 3.4, w: 4.1, h: 1.5, margin: 0, paraSpaceAfter: 4 });

  s5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s5.addText(FOOTER_TEXT, { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  s5.addNotes("安全事件与热点赛道深度分析：\n\n安全事件：\n1. 伊朗交易所$9000万被盗事件的地缘政治背景：伊以冲突升级期间，国家级黑客活动显著增加。这不仅是金额问题，更反映了加密行业在地缘冲突中的脆弱性。中东地区的加密基础设施安全需要特别关注。\n\n2. Resolv USR攻击的教训极为深刻：攻击者入侵的是AWS密钥管理服务（KMS）中的特权私钥，这暴露了DeFi协议过度依赖中心化云服务的风险。更严重的是，此事件在Morpho上引发了连锁反应——多个借贷金库因预言机价格滞后（仍以$1.13估值已崩盘的USR）而被掏空。这提醒我们：DeFi的可组合性在危机时刻会变成系统性风险的传导器。\n\n3. 2026年YTD已发生103起安全事件+36起钓鱼诈骗，总损失约$4.8亿。社会工程学已超越代码漏洞成为主要攻击向量——这意味着技术审计不够，团队的安全运营和人员培训同样关键。\n\nAI x Crypto：\nNVIDIA GTC 2026后AI加密板块市值超$280亿。更重要的趋势是BTC矿企的集体转型：挖矿成本$79,995/BTC已远超市价$68K，每枚亏损约$19K。矿企已签署超$700亿AI/HPC合同，预计到年底AI收入占比将从不足15%升至70%+。这是加密挖矿行业的根本性转折——从能源消耗者转变为AI算力提供者。\n\n预测市场：\nPolymarket月度总量突破$210亿，是去年同期的7.5倍。CFTC将其纳入监管框架（自认证新市场规则）是重要里程碑——这意味着链上预测市场首次获得美国监管体系的正式认可。反内幕交易新规的出台也表明该赛道正在走向成熟和合规。");

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
    catRow("4/3", "Wormhole 6\u4EBF\u679A\u89E3\u9501\uFF086%\u4F9B\u5E94\u91CF\uFF09", "W", "\u4F9B\u7ED9\u538B\u529B >$9000\u4E07"),
    catRow("4/6", "Hyperliquid\u6838\u5FC3\u8D21\u732E\u8005\u89E3\u9501", "HYPE", "\u5386\u53F2\u4E0A\u9AD8\u98CE\u9669\u629B\u538B\u7C7B\u522B", true),
    catRow("4/15", "Starknet 1.27\u4EBF\u679A\u89E3\u9501", "STRK", "\u4F9B\u7ED9\u538B\u529B"),
    catRow("4\u6708\u4E0B\u65EC", "CLARITY Act\u5168\u9762\u59D4\u5458\u4F1A\u5BA1\u8BAE", "\u7A33\u5B9A\u5E01", "\u7A33\u5B9A\u5E01\u6536\u76CA\u7981\u4EE4\u5B9A\u8C03", true),
    catRow("4/28-29", "\u9C8D\u5A01\u5C14\u6700\u540E\u4E00\u6B21FOMC\u4F1A\u8BAE", "\u5168\u5E02\u573A", "\u5229\u7387\u524D\u77BB\u6307\u5F15"),
    catRow("4/29-30", "TOKEN2049 Dubai", "\u5168\u5E02\u573A", "\u884C\u4E1A\u5E74\u5EA6\u6700\u5927\u5CF0\u4F1A\u4E4B\u4E00"),
    catRow("Q2", "SOL/XRP\u73B0\u8D27ETF\u65B0\u8FDB\u5C55", "SOL/XRP", "\u5546\u54C1\u5206\u7C7B\u540EETF\u901A\u9053\u6253\u5F00"),
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
    { text: "\u5E02\u573A\u60C5\u7EEA\u5904\u4E8E\u201C\u6781\u5EA6\u6050\u60E7\u201D\uFF08FGI=9\uFF09\uFF0C\u5730\u7F18\u653F\u6CBB\u98CE\u9669\u548C\u5173\u7A0E\u653F\u7B56\u538B\u5236\u98CE\u9669\u504F\u597D\u3002SEC-CFTC\u8054\u5408\u89C4\u5219\u662F\u91CD\u5927\u957F\u671F\u5229\u597D\uFF0C\u4F46\u77ED\u671F\u672A\u80FD\u8F6C\u5316\u4E3A\u5E02\u573A\u52A8\u529B\u3002", options: { breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "\u91CD\u70B9\u5173\u6CE8\uFF1A", options: { bold: true, fontSize: 10, fontFace: "Arial", color: C.navy } },
    { text: " 4/6 HYPE\u8D21\u732E\u8005\u89E3\u9501\uFF08\u6301\u4ED3\u76F8\u5173\uFF09\u3001CLARITY Act\u5BA1\u8BAE\u5BF9\u7A33\u5B9A\u5E01\u8D5B\u9053\u5F71\u54CD\u3001\u4F0A\u4EE5\u51B2\u7A81\u8D70\u5411\u5BF9\u5E02\u573A\u60C5\u7EEA\u7684\u6301\u7EED\u538B\u5236\u3002", options: { fontSize: 10, fontFace: "Arial", color: C.bodyText } },
  ], { x: 0.8, y: 4.0, w: 8.5, h: 0.85, margin: 0 });

  s6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s6.addText(FOOTER_TEXT + "  |  CONFIDENTIAL", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  s6.addNotes("催化剂与展望深度分析：\n\n最需要关注的催化剂：\n\n1. 4/6 Hyperliquid核心贡献者解锁：这是持仓直接相关的事件。历史上，核心团队解锁往往伴随短期抛压。但Hyperliquid的基本面极强（周费用$14M、年化$6亿+），如果团队选择锁仓或做市而非抛售，可能成为信心催化剂。建议密切关注解锁前后的链上大额转账。\n\n2. 4月下旬CLARITY Act全面委员会审议：这将决定稳定币赛道的命运。如果禁止收益条款通过，Circle、Coinbase等严重依赖储备利息收入的公司将面临商业模式重塑。反之，如果条款被修改或删除，稳定币板块可能迎来强劲反弹。\n\n3. 4/28-29鲍威尔最后一次FOMC：作为鲍威尔任期内最后一次会议，市场将密切关注其对继任者的暗示和利率前瞻指引。如果释放鸽派信号，可能触发风险资产反弹。\n\n4. Wormhole 6亿枚解锁（4/3）：占总供应量6%，价值超$9000万的供给压力。跨链桥赛道竞争激烈，需关注团队是否有缓冲机制。\n\n本周总结：\n市场处于多重利空共振期（地缘冲突+关税+通胀上修+ETF流出），恐惧指数降至极端区域。但监管端出现了加密行业历史上最清晰的利好（16种代币获商品认定）。这种基本面改善+短期情绪恶化的组合，历史上往往预示着中期的买入机会。关键是宏观环境何时改善——伊以冲突的走向和美联储的态度将是决定因素。\n\n组合建议关注点：HYPE解锁风险管理、SOL现货ETF进展跟踪、PENDLE在收益代币赛道承压时的估值安全边际。");

  const outPath = "/Users/lx/Github/crypto-weekly-digest/reports/2026-03-23/slides-zh.pptx";
  await pres.writeFile({ fileName: outPath });
  console.log("\u2705 Written to " + outPath);
}

generate().catch(e => { console.error(e); process.exit(1); });
