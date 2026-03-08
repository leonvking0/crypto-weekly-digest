const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");

// Icons
const { FaBitcoin, FaShieldAlt, FaChartLine, FaGavel, FaRocket, FaHandshake } = require("react-icons/fa");
const { MdTrendingDown, MdTrendingUp, MdWarning } = require("react-icons/md");

async function iconToBase64Png(IconComponent, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComponent, { color, size: String(size) })
  );
  const pngBuffer = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + pngBuffer.toString("base64");
}

// Color Palette — 国企商务风 (SOE Business Style)
const C = {
  navy: "1A2744",       // primary dark
  darkNavy: "0F1A2E",   // slide bg dark
  gold: "C9A96E",       // accent gold
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
  pres.title = "加密市场周报 2026.02.20–02.27";

  // ============ SLIDE 1: Title ============
  const s1 = pres.addSlide();
  s1.background = { color: C.darkNavy };
  
  // Gold accent line
  s1.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.6, w: 1.2, h: 0.04, fill: { color: C.gold } });
  
  s1.addText("加密市场周报", {
    x: 0.8, y: 1.8, w: 8, h: 1.0,
    fontSize: 42, fontFace: "Arial", bold: true, color: C.white, margin: 0
  });
  s1.addText("2026年2月20日 — 2月27日", {
    x: 0.8, y: 2.7, w: 8, h: 0.5,
    fontSize: 20, fontFace: "Arial", color: C.gold, margin: 0
  });
  s1.addText("数据来源：Messari · The Block · CoinGecko · CoinGlass · DefiLlama", {
    x: 0.8, y: 3.4, w: 8, h: 0.4,
    fontSize: 11, fontFace: "Arial", color: C.medGray, margin: 0
  });
  
  // Bottom bar
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.gold } });
  s1.addText("CONFIDENTIAL — FOR LP DISTRIBUTION ONLY", {
    x: 0, y: 5.25, w: 10, h: 0.375,
    fontSize: 9, fontFace: "Arial", color: C.darkNavy, align: "center", valign: "middle", bold: true
  });

  // ============ SLIDE 2: Market Overview ============
  const s2 = pres.addSlide();
  s2.background = { color: C.offWhite };
  
  // Header bar
  s2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy } });
  s2.addText("📊  市场概览", {
    x: 0.6, y: 0, w: 9, h: 0.7,
    fontSize: 20, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0
  });

  // Key metrics row — 3 cards
  const cardY = 0.95;
  const cardH = 1.05;
  const cardW = 2.8;
  const cardGap = 0.3;
  const cardStartX = 0.6;

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
  s2.addText("恐惧贪婪指数", { x: c3x + 0.25, y: cardY + 0.1, w: 2.4, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.medGray, bold: true, margin: 0 });
  s2.addText("13", { x: c3x + 0.25, y: cardY + 0.35, w: 1, h: 0.4, fontSize: 28, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
  s2.addText("极度恐惧", { x: c3x + 1.1, y: cardY + 0.42, w: 1.5, h: 0.3, fontSize: 13, fontFace: "Arial", color: C.red, margin: 0 });

  // Price table
  const tableY = 2.25;
  const headerRow = [
    { text: "资产", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 11, fontFace: "Arial", align: "center" } },
    { text: "价格", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 11, fontFace: "Arial", align: "center" } },
    { text: "周涨跌幅", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 11, fontFace: "Arial", align: "center" } },
    { text: "备注", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 11, fontFace: "Arial", align: "center" } },
  ];
  const makeRow = (asset, price, chg, note, isNeg) => [
    { text: asset, options: { fontSize: 11, fontFace: "Arial", bold: true, align: "center", color: C.bodyText } },
    { text: price, options: { fontSize: 11, fontFace: "Arial", align: "right", color: C.bodyText } },
    { text: chg, options: { fontSize: 11, fontFace: "Arial", align: "center", color: isNeg ? C.red : C.green, bold: true } },
    { text: note, options: { fontSize: 10, fontFace: "Arial", color: C.medGray } },
  ];
  const tableData = [
    headerRow,
    makeRow("SOL", "$81.73", "-3.0%", "市值 ~$40B", true),
    makeRow("HYPE", "$30.25", "-3.7%", "3/6 解锁 $316M", true),
    makeRow("PENDLE", "$1.18", "-8.3%", "", true),
    makeRow("ENA", "$0.099", "-5.4%", "", true),
    makeRow("DYDX", "$0.080", "-12.4%", "跌幅领先", true),
  ];
  s2.addTable(tableData, {
    x: 0.6, y: tableY, w: 8.8, colW: [1.3, 1.8, 1.5, 4.2],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.35, 0.33, 0.33, 0.33, 0.33, 0.33],
    autoPage: false,
  });

  // Macro note
  s2.addText("宏观驱动：美最高法院 2/20 推翻行政关税令 → 风险偏好短暂回暖；整体仍处\u201C极度恐惧\u201D区间", {
    x: 0.6, y: 4.55, w: 8.8, h: 0.4,
    fontSize: 10, fontFace: "Arial", color: C.medGray, italic: true, margin: 0
  });

  // Footer
  s2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s2.addText("加密市场周报 | 2026.02.20–02.27", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  // ============ SLIDE 3: Top Stories + DeFi/L1 ============
  const s3 = pres.addSlide();
  s3.background = { color: C.offWhite };
  s3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy } });
  s3.addText("🔥  本周要闻 & 行业动态", {
    x: 0.6, y: 0, w: 9, h: 0.7,
    fontSize: 20, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0
  });

  // Left column: Top 5 stories
  s3.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 4.1, fill: { color: C.white }, shadow: makeShadow() });
  s3.addText("本周要闻 TOP 5", {
    x: 0.6, y: 0.95, w: 4, h: 0.35,
    fontSize: 13, fontFace: "Arial", bold: true, color: C.navy, margin: 0
  });

  const stories = [
    ["1", "Block Inc. 裁员40%，全面转向AI运营", "Jack Dorsey 宣布以 AI Agent 替代传统企业架构"],
    ["2", "Uniswap Fee Switch 投票启动", "8条L2激活协议费，预估年化 $27M 收入"],
    ["3", "SEC 代币分类 + OCC 稳定币规则", "证券vs商品界限明确；稳定币最低$5M资本金"],
    ["4", "Aave 累计借贷量破 $1 万亿", "RWA资产同步超$1B，DeFi信贷龙头"],
    ["5", "Grayscale 申报 TAO 现货 ETF", "机构资金首次触及去中心化AI计算网络"],
  ];

  stories.forEach((s, i) => {
    const itemY = 1.38 + i * 0.72;
    // Number circle
    s3.addShape(pres.shapes.OVAL, { x: 0.65, y: itemY, w: 0.3, h: 0.3, fill: { color: C.gold } });
    s3.addText(s[0], { x: 0.65, y: itemY, w: 0.3, h: 0.3, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, align: "center", valign: "middle", margin: 0 });
    s3.addText(s[1], { x: 1.1, y: itemY - 0.02, w: 3.6, h: 0.25, fontSize: 11, fontFace: "Arial", bold: true, color: C.bodyText, margin: 0 });
    s3.addText(s[2], { x: 1.1, y: itemY + 0.22, w: 3.6, h: 0.25, fontSize: 9, fontFace: "Arial", color: C.medGray, margin: 0 });
  });

  // Right column: L1/L2 + DeFi highlights
  s3.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 4.5, h: 1.95, fill: { color: C.white }, shadow: makeShadow() });
  s3.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s3.addText("Layer 1 & Layer 2", { x: 5.3, y: 0.9, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });
  
  const l1Items = [
    "MegaETH 主网上线：首日 210万笔交易，亚10ms出块",
    "Base 脱离 OP Stack，转向自研架构",
    "RWA 链上市值破 $15B（YoY +200%）",
    "BTC ETF 三日净流入 $1.1B",
  ];
  s3.addText(l1Items.map((t, i) => ({
    text: t,
    options: { bullet: true, breakLine: i < l1Items.length - 1, fontSize: 10, fontFace: "Arial", color: C.bodyText }
  })), { x: 5.3, y: 1.32, w: 4.1, h: 1.45, margin: 0, paraSpaceAfter: 4 });

  // DeFi box
  s3.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 3.05, w: 4.5, h: 1.95, fill: { color: C.white }, shadow: makeShadow() });
  s3.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 3.05, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s3.addText("DeFi & 永续合约", { x: 5.3, y: 3.05, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  const defiItems = [
    "Aave 借贷 $1T + RWA $1B 历史性里程碑",
    "DeFi TVL $96.8B（-0.52%），Solana逆势+3.46%",
    "Hyperliquid 周交易量 $35.4B，单日峰值 $6B",
    "永续DEX市场份额达10.2%（2024年初仅2%）",
  ];
  s3.addText(defiItems.map((t, i) => ({
    text: t,
    options: { bullet: true, breakLine: i < defiItems.length - 1, fontSize: 10, fontFace: "Arial", color: C.bodyText }
  })), { x: 5.3, y: 3.47, w: 4.1, h: 1.45, margin: 0, paraSpaceAfter: 4 });

  // Footer
  s3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s3.addText("加密市场周报 | 2026.02.20–02.27", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  // ============ SLIDE 4: Regulation + Fundraising ============
  const s4 = pres.addSlide();
  s4.background = { color: C.offWhite };
  s4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy } });
  s4.addText("⚖️  监管政策 & 投融资", {
    x: 0.6, y: 0, w: 9, h: 0.7,
    fontSize: 20, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0
  });

  // Left: Regulation
  s4.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 4.1, fill: { color: C.white }, shadow: makeShadow() });
  s4.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 0.06, h: 4.1, fill: { color: C.gold } });
  s4.addText("监管与政策", { x: 0.7, y: 0.95, w: 4, h: 0.35, fontSize: 14, fontFace: "Arial", bold: true, color: C.navy, margin: 0 });

  const regItems = [
    { title: "SEC 代币分类指引", desc: "提交白宫，明确证券 vs 商品界限" },
    { title: "OCC GENIUS Act 框架", desc: "稳定币最低$5M资本金，禁止付息" },
    { title: "CFTC 创新顾问委员会", desc: "35人，涵盖主流Crypto/DeFi领袖" },
    { title: "CLARITY Act 僵局", desc: "稳定币付息成银行业vs加密业分歧" },
    { title: "英国FCA稳定币沙盒", desc: "Revolut等4家入选，2026年底出规" },
  ];

  regItems.forEach((item, i) => {
    const iy = 1.45 + i * 0.65;
    s4.addText(item.title, { x: 0.7, y: iy, w: 4, h: 0.25, fontSize: 11, fontFace: "Arial", bold: true, color: C.bodyText, margin: 0 });
    s4.addText(item.desc, { x: 0.7, y: iy + 0.23, w: 4, h: 0.25, fontSize: 10, fontFace: "Arial", color: C.medGray, margin: 0 });
  });

  // Right: Fundraising
  s4.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 4.5, h: 4.1, fill: { color: C.white }, shadow: makeShadow() });
  s4.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 0.06, h: 4.1, fill: { color: C.gold } });
  s4.addText("投融资与并购", { x: 5.4, y: 0.95, w: 4, h: 0.35, fontSize: 14, fontFace: "Arial", bold: true, color: C.navy, margin: 0 });

  // Fundraising table
  const fundHeader = [
    { text: "项目", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial" } },
    { text: "金额", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial", align: "right" } },
    { text: "备注", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial" } },
  ];
  const fundRow = (name, amt, note) => [
    { text: name, options: { fontSize: 10, fontFace: "Arial", color: C.bodyText, bold: true } },
    { text: amt, options: { fontSize: 10, fontFace: "Arial", color: C.bodyText, align: "right" } },
    { text: note, options: { fontSize: 9, fontFace: "Arial", color: C.medGray } },
  ];
  const fundData = [
    fundHeader,
    fundRow("Flying Tulip", "$206M", "Token Sale"),
    fundRow("Whop", "$200M", "Tether投资, 估值$1.6B"),
    fundRow("Gold.com", "$150M", "Tether投资, 数字黄金"),
    fundRow("BTC Inc.", "$107M", "Nakamoto收购"),
    fundRow("Anchorage", "$100M", "Tether投资, 估值$4.2B"),
  ];
  s4.addTable(fundData, {
    x: 5.3, y: 1.45, w: 4.1, colW: [1.1, 0.9, 2.1],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.3, 0.3, 0.3, 0.3, 0.3, 0.3],
    autoPage: false,
  });

  // Tether callout
  s4.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 3.4, w: 4.1, h: 0.7, fill: { color: "FFF8E7" }, shadow: makeShadow() });
  s4.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 3.4, w: 0.06, h: 0.7, fill: { color: C.gold } });
  s4.addText([
    { text: "关注：", options: { bold: true, fontSize: 10, fontFace: "Arial", color: C.navy } },
    { text: "Tether 本周密集出手 $450M+，战略布局从稳定币延伸至电商、黄金、托管银行", options: { fontSize: 10, fontFace: "Arial", color: C.bodyText } }
  ], { x: 5.55, y: 3.45, w: 3.7, h: 0.6, margin: 0, valign: "middle" });

  // Footer
  s4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s4.addText("加密市场周报 | 2026.02.20–02.27", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  // ============ SLIDE 5: Security + AI + Prediction Markets ============
  const s5 = pres.addSlide();
  s5.background = { color: C.offWhite };
  s5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy } });
  s5.addText("🔓  安全事件 & 热点赛道", {
    x: 0.6, y: 0, w: 9, h: 0.7,
    fontSize: 20, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0
  });

  // Security incidents - left
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 2.35, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 0.35, fill: { color: C.red } });
  s5.addText("安全事件  |  周损失: ~$21M–$26M", { x: 0.6, y: 0.9, w: 4.2, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  const secTable = [
    [
      { text: "事件", options: { bold: true, fontSize: 9, fontFace: "Arial", fill: { color: C.lightGray }, color: C.bodyText } },
      { text: "损失", options: { bold: true, fontSize: 9, fontFace: "Arial", fill: { color: C.lightGray }, color: C.bodyText, align: "right" } },
      { text: "方式", options: { bold: true, fontSize: 9, fontFace: "Arial", fill: { color: C.lightGray }, color: C.bodyText } },
    ],
    [{ text: "YieldBlox", options: { fontSize: 9, fontFace: "Arial" } }, { text: "~$10M", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "预言机操纵", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
    [{ text: "IoTeX/ioTube", options: { fontSize: 9, fontFace: "Arial" } }, { text: "$4.4–9M", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "私钥泄露", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
    [{ text: "韩国国税厅", options: { fontSize: 9, fontFace: "Arial" } }, { text: "~$4.8M", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "助记词泄露", options: { fontSize: 9, fontFace: "Arial", color: C.medGray } }],
    [{ text: "Moonwell", options: { fontSize: 9, fontFace: "Arial" } }, { text: "$1.78M", options: { fontSize: 9, fontFace: "Arial", align: "right", color: C.red } }, { text: "AI代码缺陷", options: { fontSize: 9, fontFace: "Arial", color: C.red, bold: true } }],
  ];
  s5.addTable(secTable, {
    x: 0.55, y: 1.35, w: 4.2, colW: [1.4, 1.1, 1.7],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.28, 0.28, 0.28, 0.28, 0.28],
    autoPage: false,
  });

  // AI x Crypto - right top
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 4.5, h: 1.85, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s5.addText("AI × Crypto", { x: 5.3, y: 0.9, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  const aiItems = [
    "Grayscale 申报 TAO 现货ETF — 机构化去中心化ML",
    "Bitdeer 清仓BTC，全力转向AI算力",
    "链上自主AI Agent突破11,000个（ERC-8004）",
  ];
  s5.addText(aiItems.map((t, i) => ({
    text: t,
    options: { bullet: true, breakLine: i < aiItems.length - 1, fontSize: 10, fontFace: "Arial", color: C.bodyText }
  })), { x: 5.3, y: 1.35, w: 4.1, h: 1.3, margin: 0, paraSpaceAfter: 4 });

  // Prediction Markets - right bottom
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 2.95, w: 4.5, h: 1.85, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 2.95, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s5.addText("预测市场", { x: 5.3, y: 2.95, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });

  const pmItems = [
    "Polymarket 日峰值 $425M，月度 $7B（YoY ×7.5）",
    "2月行业总量 $23.4B — Kalshi领先Polymarket",
    "ICE整合Polymarket信号；CFTC拟纳入监管",
  ];
  s5.addText(pmItems.map((t, i) => ({
    text: t,
    options: { bullet: true, breakLine: i < pmItems.length - 1, fontSize: 10, fontFace: "Arial", color: C.bodyText }
  })), { x: 5.3, y: 3.4, w: 4.1, h: 1.3, margin: 0, paraSpaceAfter: 4 });

  // Stablecoin callout bar at bottom
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.5, w: 4.5, h: 1.5, fill: { color: C.white }, shadow: makeShadow() });
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.5, w: 4.5, h: 0.35, fill: { color: C.navy } });
  s5.addText("稳定币", { x: 0.6, y: 3.5, w: 4, h: 0.35, fontSize: 11, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0 });
  const stableItems = [
    "总市值 $297.75B，USDT 2月Tron增发$1.6B",
    "Visa×Stripe稳定币卡扩至100+国家",
  ];
  s5.addText(stableItems.map((t, i) => ({
    text: t,
    options: { bullet: true, breakLine: i < stableItems.length - 1, fontSize: 10, fontFace: "Arial", color: C.bodyText }
  })), { x: 0.6, y: 3.95, w: 4.1, h: 0.9, margin: 0, paraSpaceAfter: 4 });

  // Footer
  s5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s5.addText("加密市场周报 | 2026.02.20–02.27", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  // ============ SLIDE 6: Catalysts + Outlook ============
  const s6 = pres.addSlide();
  s6.background = { color: C.offWhite };
  s6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy } });
  s6.addText("🔮  未来催化剂 & 展望", {
    x: 0.6, y: 0, w: 9, h: 0.7,
    fontSize: 20, fontFace: "Arial", bold: true, color: C.white, valign: "middle", margin: 0
  });

  // Catalysts table
  const catHeader = [
    { text: "日期", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial", align: "center" } },
    { text: "事件", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial" } },
    { text: "影响资产", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial", align: "center" } },
    { text: "预期影响", options: { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial" } },
  ];
  const catRow = (date, event, asset, impact, highlight = false) => [
    { text: date, options: { fontSize: 10, fontFace: "Arial", align: "center", color: C.bodyText, fill: highlight ? { color: "FFF3E0" } : undefined } },
    { text: event, options: { fontSize: 10, fontFace: "Arial", color: C.bodyText, bold: highlight, fill: highlight ? { color: "FFF3E0" } : undefined } },
    { text: asset, options: { fontSize: 10, fontFace: "Arial", align: "center", color: C.navy, bold: true, fill: highlight ? { color: "FFF3E0" } : undefined } },
    { text: impact, options: { fontSize: 9, fontFace: "Arial", color: C.medGray, fill: highlight ? { color: "FFF3E0" } : undefined } },
  ];
  const catData = [
    catHeader,
    catRow("3/1", "HyperEVM 主网上线", "HYPE", "EVM智能合约启用"),
    catRow("3/1", "SUI Token Unlock (5382万枚)", "SUI", "供给压力 ~$50M+"),
    catRow("3/4", "Polygon Lisovo 硬分叉", "POL", "合约优化+跨链增强"),
    catRow("3/5", "POWER Token Unlock", "POWER", "供给压力 $23M"),
    catRow("3/6", "HYPE 贡献者解锁 $316.6M", "HYPE", "重大流动性事件", true),
    catRow("3/12", "Polkadot 代币经济改革", "DOT", "减排53.6%"),
    catRow("3/12", "APT Token Unlock (1131万枚)", "APT", "供给压力 ~$100M+"),
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
  s6.addText("本周小结", { x: 0.8, y: 3.7, w: 8, h: 0.3, fontSize: 13, fontFace: "Arial", bold: true, color: C.navy, margin: 0 });
  s6.addText([
    { text: "市场情绪处于\u201C极度恐惧\u201D（FGI=13），BTC 窄幅震荡，山寨币普跌。监管端出现积极信号（SEC代币分类、OCC稳定币框架），但短期未转化为市场动力。", options: { breakLine: true, fontSize: 10, fontFace: "Arial", color: C.bodyText } },
    { text: "重点关注：", options: { bold: true, fontSize: 10, fontFace: "Arial", color: C.navy } },
    { text: " 3/6 HYPE $316M 解锁（持仓相关）、Tether 战略扩张节奏、DeFi 协议收入可持续模型进展。", options: { fontSize: 10, fontFace: "Arial", color: C.bodyText } },
  ], { x: 0.8, y: 4.0, w: 8.5, h: 0.85, margin: 0 });

  // Footer
  s6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy } });
  s6.addText("加密市场周报 | 2026.02.20–02.27  |  CONFIDENTIAL", { x: 0.5, y: 5.25, w: 9, h: 0.375, fontSize: 8, fontFace: "Arial", color: C.medGray, valign: "middle", margin: 0 });

  const outPath = "/Users/lx/Github/crypto-weekly-digest/reports/2026-02-24.pptx";
  await pres.writeFile({ fileName: outPath });
  console.log("✅ Written to " + outPath);
}

generate().catch(e => { console.error(e); process.exit(1); });
