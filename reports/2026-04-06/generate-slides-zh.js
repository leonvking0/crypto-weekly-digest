const pptxgen = require('pptxgenjs');

const C = {
  navy: '1A2744',
  darkNavy: '0F1A2E',
  gold: 'C9A96E',
  white: 'FFFFFF',
  offWhite: 'F5F6F8',
  lightGray: 'E8EBF0',
  medGray: '8C95A6',
  bodyText: '2D3748',
  red: 'C53030',
  green: '2F855A',
  paleGold: 'FFF8E7',
  paleRed: 'FDECEC'
};

const makeShadow = () => ({ type: 'outer', blur: 3, offset: 1, angle: 45, color: '000000', opacity: 0.12 });
const addFooter = (slide, text) => {
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy }, line: { color: C.navy } });
  slide.addText(text, { x: 0.45, y: 5.25, w: 9.1, h: 0.375, fontFace: 'Arial', fontSize: 8, color: C.medGray, valign: 'middle', margin: 0 });
};
const addHeader = (slide, title) => {
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy }, line: { color: C.navy } });
  slide.addText(title, { x: 0.6, y: 0, w: 8.8, h: 0.7, fontFace: 'Arial', fontSize: 20, bold: true, color: C.white, valign: 'middle', margin: 0 });
};

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_16x9';
pptx.author = '42';
pptx.company = '42';
pptx.subject = 'Crypto Weekly Digest';
pptx.title = '加密市场周报 2026.04.06–04.12';
pptx.lang = 'zh-CN';

// Slide 1
{
  const s = pptx.addSlide();
  s.background = { color: C.darkNavy };
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.8, y: 1.55, w: 1.2, h: 0.04, fill: { color: C.gold }, line: { color: C.gold } });
  s.addText('加密市场周报', { x: 0.8, y: 1.8, w: 5.5, h: 0.8, fontFace: 'Arial', fontSize: 42, bold: true, color: C.white, margin: 0 });
  s.addText('2026年4月6日 — 4月12日', { x: 0.8, y: 2.7, w: 5.5, h: 0.45, fontFace: 'Arial', fontSize: 20, color: C.gold, margin: 0 });
  s.addText('数据来源：Messari · The Block · CoinGecko · DefiLlama', { x: 0.8, y: 3.4, w: 6.5, h: 0.3, fontFace: 'Arial', fontSize: 11, color: C.medGray, margin: 0 });
  s.addText('本周关键词：宏观反弹、监管关键周、Hyperliquid 份额提升、AI Agent 商业闭环', { x: 0.8, y: 4.15, w: 8.3, h: 0.5, fontFace: 'Arial', fontSize: 14, color: C.white, margin: 0 });
  s.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.gold }, line: { color: C.gold } });
  s.addText('CONFIDENTIAL — FOR LP DISTRIBUTION ONLY', { x: 0, y: 5.25, w: 10, h: 0.375, align: 'center', valign: 'middle', fontFace: 'Arial', fontSize: 9, bold: true, color: C.darkNavy, margin: 0 });
  s.addNotes(`本页最重要的三件事：
1. 本周主线不是链上单点爆发，而是宏观风险偏好修复。美伊冲突出现 45 天停火窗口后，油价快速回落，市场把“高油价推高通胀、压制降息”的担忧先做了一轮修正，所以 BTC、ETH 的反弹更像是宏观 relief rally，而不是新牛市确认。
2. 第二条主线是制度预期。CLARITY Act 与稳定币收益规则进入关键推进期，机构现在最关心的不是一句口头利好，而是美国是否真的开始形成可执行的市场结构框架。这会决定 Q2 之后哪些资产拿到估值重估权。
3. 第三条主线是行业内部的结构分化。Hyperliquid、Base、AI Agent 支付轨道这些有真实使用增长的方向在变强，而 Solana 生态仍在消化 Drift 攻击余波。也就是说，资金开始重新奖励“有业务闭环”的子赛道，而不是无差别炒作。`);
}

// Slide 2
{
  const s = pptx.addSlide();
  s.background = { color: C.offWhite };
  addHeader(s, '📊  市场概览');
  const x0 = 0.6, y = 0.95, w = 2.8, h = 1.05, gap = 0.3;
  const card = (x, label, value, delta, deltaColor) => {
    s.addShape(pptx.shapes.RECTANGLE, { x, y, w, h, fill: { color: C.white }, line: { color: C.white }, shadow: makeShadow() });
    s.addShape(pptx.shapes.RECTANGLE, { x, y, w: 0.06, h, fill: { color: C.gold }, line: { color: C.gold } });
    s.addText(label, { x: x + 0.22, y: y + 0.1, w: 2.1, h: 0.25, fontFace: 'Arial', fontSize: 12, bold: true, color: C.medGray, margin: 0 });
    s.addText(value, { x: x + 0.22, y: y + 0.35, w: 2.2, h: 0.35, fontFace: 'Arial', fontSize: 28, bold: true, color: C.bodyText, margin: 0 });
    s.addText(delta, { x: x + 0.22, y: y + 0.73, w: 2.2, h: 0.2, fontFace: 'Arial', fontSize: 14, bold: true, color: deltaColor, margin: 0 });
  };
  card(x0, 'BTC', '$70,950', '+5.4%', C.green);
  card(x0 + w + gap, 'ETH', '$2,187.94', '+6.6%', C.green);
  s.addShape(pptx.shapes.RECTANGLE, { x: x0 + (w + gap) * 2, y, w, h, fill: { color: C.white }, line: { color: C.white }, shadow: makeShadow() });
  s.addShape(pptx.shapes.RECTANGLE, { x: x0 + (w + gap) * 2, y, w: 0.06, h, fill: { color: C.red }, line: { color: C.red } });
  s.addText('恐惧贪婪指数', { x: x0 + (w + gap) * 2 + 0.22, y: y + 0.1, w: 2.2, h: 0.25, fontFace: 'Arial', fontSize: 12, bold: true, color: C.medGray, margin: 0 });
  s.addText('16', { x: x0 + (w + gap) * 2 + 0.22, y: y + 0.35, w: 0.8, h: 0.35, fontFace: 'Arial', fontSize: 28, bold: true, color: C.red, margin: 0 });
  s.addText('极度恐惧', { x: x0 + (w + gap) * 2 + 1.0, y: y + 0.43, w: 1.5, h: 0.25, fontFace: 'Arial', fontSize: 13, color: C.red, margin: 0 });
  s.addTable([
    [
      { text: '资产', options: { bold: true, color: C.white, fill: { color: C.navy }, align: 'center', fontFace: 'Arial', fontSize: 11 } },
      { text: '价格', options: { bold: true, color: C.white, fill: { color: C.navy }, align: 'center', fontFace: 'Arial', fontSize: 11 } },
      { text: '7d', options: { bold: true, color: C.white, fill: { color: C.navy }, align: 'center', fontFace: 'Arial', fontSize: 11 } },
      { text: '备注', options: { bold: true, color: C.white, fill: { color: C.navy }, align: 'center', fontFace: 'Arial', fontSize: 11 } }
    ],
    ['SOL', '$81.87', '+3.0%', '反弹偏弱，仍受安全事件折价'],
    ['HYPE', '$41.05', '+15.7%', '永续 DEX 龙头，份额持续扩张'],
    ['ENA', '$0.09', '+16.6%', '高 beta 反弹'],
    ['PENDLE', '$1.05', '+5.8%', '收益赛道情绪修复'],
    ['DYDX', '$0.09', '-4.6%', '份额继续下滑']
  ], {
    x: 0.6, y: 2.25, w: 8.8, colW: [1.2, 1.7, 1.2, 4.7], rowH: [0.35, 0.33, 0.33, 0.33, 0.33, 0.33],
    border: { pt: 0.4, color: C.lightGray }, fontFace: 'Arial', fontSize: 10, color: C.bodyText
  });
  s.addText('总市值 $2.49T，BTC 主导率 56.9%，24h 成交额 $77.0B。宏观停火窗口与油价下跌带动全市场 short covering。', { x: 0.6, y: 4.55, w: 8.8, h: 0.35, fontFace: 'Arial', fontSize: 10, italic: true, color: C.medGray, margin: 0 });
  addFooter(s, '加密市场周报 | 2026.04.06–04.12');
  s.addNotes(`本页最重要的三件事：
1. BTC、ETH 同时上涨，而且 ETH 涨幅高于 BTC，这说明资金并不只是被动回补最安全资产，而是开始愿意承担一点额外风险去追逐有明确催化剂的资产。ETH 的催化来自升级预期与潜在机构配置回流，所以它的强势比单纯的反弹更有信息量。
2. 恐惧贪婪指数虽然回升到 16，但仍处在极度恐惧区间。这意味着市场情绪修复还远远没完成，现在更像“从极端悲观回到谨慎观望”，并不是全面 risk-on。对于仓位管理，这通常意味着可以做结构性进攻，但不适合把短期反弹误判成单边趋势。
3. 组合相关资产里，HYPE 的弹性最强，PENDLE/ENA 跟随反弹，DYDX 继续掉队。这里的核心含义是，资金在永续 DEX 龙头和能产生现金流想象的协议之间重新集中，而不是均匀扩散到所有老牌 DeFi 资产。后续若总量资金继续回暖，强者恒强概率会更高。`);
}

// Slide 3
{
  const s = pptx.addSlide();
  s.background = { color: C.offWhite };
  addHeader(s, '🔥  本周要闻 & 行业动态');
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 4.1, fill: { color: C.white }, line: { color: C.white }, shadow: makeShadow() });
  s.addText('本周要闻 TOP 5', { x: 0.6, y: 0.96, w: 3.5, h: 0.3, fontFace: 'Arial', fontSize: 13, bold: true, color: C.navy, margin: 0 });
  const stories = [
    ['1', '停火窗口触发风险资产反弹', '油价下跌 + 空头回补推高 BTC/ETH'],
    ['2', 'CLARITY Act 进入关键周', '稳定币收益与 DeFi 定性成焦点'],
    ['3', 'Hyperliquid 份额继续提升', 'Perp DEX/CEX 比率升至约 21%'],
    ['4', 'AI Agent 链上化加速', 'x402 与 AgentFi 闭环逐渐成形'],
    ['5', 'Solana 继续消化 Drift 余波', '安全与治理折价仍未消退']
  ];
  stories.forEach((it, i) => {
    const iy = 1.38 + i * 0.72;
    s.addShape(pptx.shapes.OVAL, { x: 0.64, y: iy, w: 0.3, h: 0.3, fill: { color: C.gold }, line: { color: C.gold } });
    s.addText(it[0], { x: 0.64, y: iy, w: 0.3, h: 0.3, fontFace: 'Arial', fontSize: 11, bold: true, color: C.white, align: 'center', valign: 'middle', margin: 0 });
    s.addText(it[1], { x: 1.08, y: iy - 0.02, w: 3.55, h: 0.22, fontFace: 'Arial', fontSize: 10.5, bold: true, color: C.bodyText, margin: 0 });
    s.addText(it[2], { x: 1.08, y: iy + 0.22, w: 3.55, h: 0.2, fontFace: 'Arial', fontSize: 9, color: C.medGray, margin: 0 });
  });
  const panel = (y, title, items) => {
    s.addShape(pptx.shapes.RECTANGLE, { x: 5.1, y, w: 4.5, h: 1.95, fill: { color: C.white }, line: { color: C.white }, shadow: makeShadow() });
    s.addShape(pptx.shapes.RECTANGLE, { x: 5.1, y, w: 4.5, h: 0.35, fill: { color: C.navy }, line: { color: C.navy } });
    s.addText(title, { x: 5.3, y, w: 3.9, h: 0.35, fontFace: 'Arial', fontSize: 11, bold: true, color: C.white, valign: 'middle', margin: 0 });
    s.addText(items.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < items.length - 1, fontFace: 'Arial', fontSize: 10, color: C.bodyText } })), { x: 5.3, y: y + 0.42, w: 4.0, h: 1.35, margin: 0, paraSpaceAfterPt: 5 });
  };
  panel(0.9, 'Layer 1 & Layer 2', [
    'Circle 预告量子抗性 L1 Arc，瞄准合规金融底层',
    'Base TVL 约 $5.6B，L2 三强格局更稳固',
    'Ether.fi 将约 $5.7B 生态迁移至 OP Mainnet',
    'SKALE on Base 上线，承接 AI Agent 需求'
  ]);
  panel(3.05, 'DeFi & 永续合约', [
    'RWA TVL 升至约 $8B，成为收益资金锚点',
    'Aave 市占率约 30%，流动性进一步集中',
    'Hyperliquid 30 日成交额 $185.5B，市占率约 34%',
    'Aster 成为第二梯队领先者，赛道头部化加速'
  ]);
  addFooter(s, '加密市场周报 | 2026.04.06–04.12');
  s.addNotes(`本页最重要的三件事：
1. 这页反映的是行业内部的“再集中化”。L2 里是 Base、Arbitrum、Optimism 三强格局更加稳固，DeFi 里是 Aave 把更多信用流动性吸到头部，永续 DEX 里是 Hyperliquid 继续吃份额。熊市或震荡市里，资金通常先回流到最有网络效应和最强执行力的平台。
2. Hyperliquid 的意义已经不只是一个涨得快的代币，而是链上衍生品对中心化交易所的替代速度正在被市场重新估值。Perp DEX/CEX 比率升至约 21%，说明用户越来越接受把主交易行为放到链上执行，这对整个 DeFi 现金流逻辑都是正面的。
3. Circle 做 Arc、SKALE on Base 承接 AI Agent、Ether.fi 大规模迁移到 OP，这些事件共同说明基础设施竞争已经进入“谁能承接真实工作负载”的阶段。未来胜负手不是再讲 TPS 故事，而是谁能更低摩擦地承接稳定币结算、AI 交易代理、机构资金与 RWA 发行。`);
}

// Slide 4
{
  const s = pptx.addSlide();
  s.background = { color: C.offWhite };
  addHeader(s, '⚖️  监管政策 & 投融资');
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 4.1, fill: { color: C.white }, line: { color: C.white }, shadow: makeShadow() });
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 0.06, h: 4.1, fill: { color: C.gold }, line: { color: C.gold } });
  s.addText('监管与政策', { x: 0.7, y: 0.96, w: 3.5, h: 0.3, fontFace: 'Arial', fontSize: 14, bold: true, color: C.navy, margin: 0 });
  const regs = [
    ['CLARITY Act 关键推进', '稳定币收益与 DeFi 边界是争议核心'],
    ['SEC 批准 Nasdaq 代币化证券试点', '传统证券上链进入更实质阶段'],
    ['香港首批稳定币牌照发放', '亚洲合规稳定币竞争升温'],
    ['英国禁止加密政治捐款', '政治资金监管趋严'],
    ['SEC/CFTC 非证券框架开始执行', '现货与协议行为边界更清晰']
  ];
  regs.forEach((r, i) => {
    const yy = 1.45 + i * 0.65;
    s.addText(r[0], { x: 0.7, y: yy, w: 3.9, h: 0.22, fontFace: 'Arial', fontSize: 11, bold: true, color: C.bodyText, margin: 0 });
    s.addText(r[1], { x: 0.7, y: yy + 0.23, w: 3.9, h: 0.2, fontFace: 'Arial', fontSize: 9.5, color: C.medGray, margin: 0 });
  });
  s.addShape(pptx.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 4.5, h: 4.1, fill: { color: C.white }, line: { color: C.white }, shadow: makeShadow() });
  s.addShape(pptx.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 0.06, h: 4.1, fill: { color: C.gold }, line: { color: C.gold } });
  s.addText('投融资与并购', { x: 5.4, y: 0.96, w: 3.5, h: 0.3, fontFace: 'Arial', fontSize: 14, bold: true, color: C.navy, margin: 0 });
  s.addTable([
    [
      { text: '项目', options: { bold: true, color: C.white, fill: { color: C.navy }, fontFace: 'Arial', fontSize: 10 } },
      { text: '金额', options: { bold: true, color: C.white, fill: { color: C.navy }, fontFace: 'Arial', fontSize: 10, align: 'right' } },
      { text: '备注', options: { bold: true, color: C.white, fill: { color: C.navy }, fontFace: 'Arial', fontSize: 10 } }
    ],
    ['Pharos', '$44M', 'RWA / DePIN 导向 L1'],
    ['BVNK + Mastercard', '$1.8B', '稳定币支付整合'],
    ['Ozak AI', '$6.65M', 'AI 金融分析代理'],
    ['Kulipa', '$6.2M', '非托管钱包支付卡']
  ], {
    x: 5.3, y: 1.45, w: 4.1, colW: [1.2, 0.9, 2.0], rowH: [0.3, 0.32, 0.32, 0.32, 0.32],
    border: { pt: 0.4, color: C.lightGray }, fontFace: 'Arial', fontSize: 9.5, color: C.bodyText
  });
  s.addShape(pptx.shapes.RECTANGLE, { x: 5.3, y: 3.35, w: 4.1, h: 0.85, fill: { color: C.paleGold }, line: { color: C.paleGold }, shadow: makeShadow() });
  s.addShape(pptx.shapes.RECTANGLE, { x: 5.3, y: 3.35, w: 0.06, h: 0.85, fill: { color: C.gold }, line: { color: C.gold } });
  s.addText('一级市场仍在集中投向三条主线：稳定币支付、RWA 底层、AI+Crypto。传统金融并购开始直接买“有现金流的加密基础设施”。', { x: 5.52, y: 3.48, w: 3.6, h: 0.55, fontFace: 'Arial', fontSize: 10, color: C.bodyText, margin: 0, valign: 'mid' });
  addFooter(s, '加密市场周报 | 2026.04.06–04.12');
  s.addNotes(`本页最重要的三件事：
1. CLARITY Act 是这周最重要的制度变量，因为它不只是影响稳定币，而是会影响 DeFi、交易所、钱包和资本市场基础设施的合规边界。一旦法案推进明确，市场会开始给“可以合法扩张”的赛道更高估值，而把监管灰区资产继续打折。
2. Nasdaq 代币化证券试点的象征意义很大。过去几年大家一直在讲 RWA，但真正的门槛不在技术，而在证券交易、清算和托管体系是否愿意上链。现在 SEC 允许以试点形式前进，说明传统金融的资产发行与流通链路开始真的松动。
3. 融资与并购也在验证同一件事：资本不再追空洞叙事，而是追有落地收入场景的基础设施。BVNK 与 Mastercard 的大额交易说明稳定币支付已经不只是 crypto native 机会，而是在吸引全球支付网络亲自下场整合。AI 和 RWA 方向持续拿钱，则说明一级市场押注的是“下一代金融操作系统”。`);
}

// Slide 5
{
  const s = pptx.addSlide();
  s.background = { color: C.offWhite };
  addHeader(s, '🛡️  安全事件 & 重点赛道');
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 4.1, fill: { color: C.white }, line: { color: C.white }, shadow: makeShadow() });
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 0.35, fill: { color: C.red }, line: { color: C.red } });
  s.addText('安全事件', { x: 0.65, y: 0.9, w: 3.5, h: 0.35, fontFace: 'Arial', fontSize: 12, bold: true, color: C.white, valign: 'middle', margin: 0 });
  s.addTable([
    [
      { text: '项目', options: { bold: true, color: C.white, fill: { color: C.red }, fontFace: 'Arial', fontSize: 10 } },
      { text: '损失', options: { bold: true, color: C.white, fill: { color: C.red }, fontFace: 'Arial', fontSize: 10 } },
      { text: '说明', options: { bold: true, color: C.white, fill: { color: C.red }, fontFace: 'Arial', fontSize: 10 } }
    ],
    ['Drift', '$280M–$285M', '社工 + durable nonce + 假资产白名单'],
    ['Resolv Labs', '~$25M', 'AWS KMS 被攻破后触发铸币漏洞'],
    ['Stabble', '预警', '疑似 DPRK 内部人风险，团队提示撤资']
  ], {
    x: 0.6, y: 1.35, w: 4.05, colW: [0.95, 1.0, 2.1], rowH: [0.28, 0.34, 0.34, 0.34],
    border: { pt: 0.4, color: C.lightGray }, fontFace: 'Arial', fontSize: 9.2, color: C.bodyText
  });
  s.addText('稳定币', { x: 0.6, y: 3.0, w: 2, h: 0.25, fontFace: 'Arial', fontSize: 12, bold: true, color: C.navy, margin: 0 });
  s.addText([
    { text: '• ', options: { color: C.gold } }, { text: '总市值突破 $250B，行业体量进入核心金融基础设施区间', options: { color: C.bodyText } },
    { text: '\n• ', options: { color: C.gold } }, { text: '香港发放首批稳定币发行牌照，亚洲合规竞争升温', options: { color: C.bodyText } },
    { text: '\n• ', options: { color: C.gold } }, { text: '美国对收益包装的监管边界仍未最终落定', options: { color: C.bodyText } }
  ], { x: 0.7, y: 3.35, w: 3.9, h: 1.15, fontFace: 'Arial', fontSize: 10, margin: 0 });

  const box = (y, title, items) => {
    s.addShape(pptx.shapes.RECTANGLE, { x: 5.1, y, w: 4.5, h: 1.95, fill: { color: C.white }, line: { color: C.white }, shadow: makeShadow() });
    s.addShape(pptx.shapes.RECTANGLE, { x: 5.1, y, w: 4.5, h: 0.35, fill: { color: C.navy }, line: { color: C.navy } });
    s.addText(title, { x: 5.3, y, w: 3.9, h: 0.35, fontFace: 'Arial', fontSize: 11, bold: true, color: C.white, valign: 'middle', margin: 0 });
    s.addText(items.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < items.length - 1, fontFace: 'Arial', fontSize: 10, color: C.bodyText } })), { x: 5.3, y: y + 0.42, w: 4.0, h: 1.35, margin: 0, paraSpaceAfterPt: 5 });
  };
  box(0.9, 'AI × Crypto', [
    'AgentFi 从叙事进入真实资金管理与交易执行',
    'AlphaTON 为 TON AI 算力网络融资 $43M',
    'Covenant AI 退出 Bittensor，暴露治理争议',
    'x402 被视为 Agent 商业闭环关键支付标准'
  ]);
  box(3.05, '预测市场', [
    'Polymarket / Kalshi 被重新理解为数据基础设施',
    'BTC 相关市场对现货走势反应极快',
    '地缘政治与能源合约吸收大量新增流动性',
    'Robinhood、IBKR、Crypto.com 持续扩张入口'
  ]);
  addFooter(s, '加密市场周报 | 2026.04.06–04.12');
  s.addNotes(`本页最重要的三件事：
1. 安全事件依然是这个周期里最被低估的系统性风险。Drift 的损失规模和攻击方式说明，问题不只是智能合约漏洞，而是团队治理、权限设计、签名流程、云基础设施都可能成为突破口。也就是说，未来市场会对“治理复杂但安全预算不足”的协议持续打折。
2. 稳定币与 AI × Crypto 其实在同一条逻辑线上：都在争夺“链上真实经济活动的底层结算与执行权”。稳定币负责结算，Agent 负责自动化执行。如果 x402 这类标准普及，链上支付和 AI 自动调用服务之间会出现非常强的飞轮。
3. 预测市场的意义也在升级。过去它更像投机工具，现在越来越像实时定价和情绪数据层。对交易者和研究员来说，它提供了一个比社媒更可量化、比传统民调更实时的预期观察窗口。未来它和稳定币支付、AI 代理结合，可能形成新的链上信息市场基础设施。`);
}

// Slide 6
{
  const s = pptx.addSlide();
  s.background = { color: C.offWhite };
  addHeader(s, '🔮  催化剂 & 周度展望');
  s.addTable([
    [
      { text: '日期', options: { bold: true, color: C.white, fill: { color: C.navy }, fontFace: 'Arial', fontSize: 10 } },
      { text: '事件', options: { bold: true, color: C.white, fill: { color: C.navy }, fontFace: 'Arial', fontSize: 10 } },
      { text: '影响资产', options: { bold: true, color: C.white, fill: { color: C.navy }, fontFace: 'Arial', fontSize: 10 } },
      { text: '预期影响', options: { bold: true, color: C.white, fill: { color: C.navy }, fontFace: 'Arial', fontSize: 10 } }
    ],
    ['4/13–4/18', 'CLARITY Act 关键推进', '稳定币 / DeFi / 全市场', '监管预期重定价'],
    ['4/15', 'STRK 解锁', 'STRK', '中等供给压力'],
    ['4/16', 'ARB 解锁', 'ARB', 'L2 板块波动放大'],
    ['4/17', 'deBridge 大额解锁', 'DBR', '高供给冲击'],
    ['4/17–4/24', 'Flare FIP.16 投票', 'FLR', '影响通胀与质押逻辑'],
    ['4/27–4/29', 'Bitcoin 2026 大会', 'BTC / BTCFi', '潜在机构催化']
  ], {
    x: 0.45, y: 1.0, w: 6.2, colW: [1.05, 2.2, 1.35, 1.6], rowH: [0.32, 0.34, 0.3, 0.3, 0.3, 0.3, 0.32],
    border: { pt: 0.4, color: C.lightGray }, fontFace: 'Arial', fontSize: 9.3, color: C.bodyText
  });
  s.addShape(pptx.shapes.RECTANGLE, { x: 6.95, y: 1.0, w: 2.6, h: 3.9, fill: { color: C.white }, line: { color: C.white }, shadow: makeShadow() });
  s.addShape(pptx.shapes.RECTANGLE, { x: 6.95, y: 1.0, w: 2.6, h: 0.35, fill: { color: C.navy }, line: { color: C.navy } });
  s.addText('周度总结', { x: 7.15, y: 1.0, w: 2.1, h: 0.35, fontFace: 'Arial', fontSize: 11, bold: true, color: C.white, valign: 'middle', margin: 0 });
  s.addText([
    { text: '• ', options: { color: C.gold } }, { text: '市场从系统性恐慌切回结构性修复', options: { color: C.bodyText } },
    { text: '\n• ', options: { color: C.gold } }, { text: '监管是未来两周最重要外部变量', options: { color: C.bodyText } },
    { text: '\n• ', options: { color: C.gold } }, { text: '组合相关赛道里，HYPE 继续最强，PENDLE/ENA 随风险偏好修复', options: { color: C.bodyText } },
    { text: '\n• ', options: { color: C.gold } }, { text: '若法案推进顺利，稳定币、DeFi、RWA 有望先获估值重估', options: { color: C.bodyText } }
  ], { x: 7.18, y: 1.48, w: 2.0, h: 1.9, fontFace: 'Arial', fontSize: 10, margin: 0 });
  s.addShape(pptx.shapes.RECTANGLE, { x: 7.15, y: 3.55, w: 2.05, h: 0.9, fill: { color: C.paleGold }, line: { color: C.paleGold } });
  s.addText('组合观察：\nHYPE 维持强势主线，PENDLE/ENA 属于风险偏好回暖下的弹性修复。', { x: 7.3, y: 3.72, w: 1.75, h: 0.5, fontFace: 'Arial', fontSize: 9.5, bold: true, color: C.navy, margin: 0, align: 'center' });
  addFooter(s, '加密市场周报 | 2026.04.06–04.12');
  s.addNotes(`本页最重要的三件事：
1. 接下来一到两周最重要的事情不是价格本身，而是“监管与供给”两个变量谁先主导市场。CLARITY Act 如果顺利推进，会让稳定币、DeFi、RWA 这些长期受折价的赛道先得到修复；反之如果推进受阻，市场会重新回到保守定价框架。
2. 中短期的交易噪音主要来自解锁。STRK、ARB、DBR 这类供给事件不会改变长期逻辑，但会显著影响局部流动性和板块风险偏好，尤其是在整体成交额仍未完全恢复的环境下。也就是说，方向判断正确不代表时间点正确，需要更重视仓位节奏。
3. 组合角度，HYPE 仍是最强结构性资产，因为它背后对应的是链上衍生品份额提升这个长期趋势。PENDLE 与 ENA 则更像风险偏好修复下的弹性仓位。如果未来两周监管预期转暖，这两类资产可能继续补涨；如果宏观再度恶化，资金大概率仍先回到 BTC 和少数现金流逻辑更清晰的龙头。`);
}

pptx.writeFile('/Users/lx/Github/crypto-weekly-digest/reports/2026-04-06/slides-zh.pptx');
