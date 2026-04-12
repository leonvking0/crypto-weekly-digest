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
  paleGold: 'FFF8E7'
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
pptx.title = 'Crypto Market Weekly Report 2026.04.06–04.12';
pptx.lang = 'en-US';

// 1
{
  const s = pptx.addSlide();
  s.background = { color: C.darkNavy };
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.8, y: 1.55, w: 1.2, h: 0.04, fill: { color: C.gold }, line: { color: C.gold } });
  s.addText('Crypto Market Weekly', { x: 0.8, y: 1.8, w: 6, h: 0.8, fontFace: 'Arial', fontSize: 42, bold: true, color: C.white, margin: 0 });
  s.addText('April 6 – April 12, 2026', { x: 0.8, y: 2.7, w: 5.5, h: 0.45, fontFace: 'Arial', fontSize: 20, color: C.gold, margin: 0 });
  s.addText('Sources: Messari · The Block · CoinGecko · DefiLlama', { x: 0.8, y: 3.4, w: 6.5, h: 0.3, fontFace: 'Arial', fontSize: 11, color: C.medGray, margin: 0 });
  s.addText('Themes: macro relief rally, regulatory inflection, Hyperliquid share gains, AI agent payment rails', { x: 0.8, y: 4.15, w: 8.2, h: 0.55, fontFace: 'Arial', fontSize: 14, color: C.white, margin: 0 });
  s.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.gold }, line: { color: C.gold } });
  s.addText('CONFIDENTIAL — FOR LP DISTRIBUTION ONLY', { x: 0, y: 5.25, w: 10, h: 0.375, align: 'center', valign: 'middle', fontFace: 'Arial', fontSize: 9, bold: true, color: C.darkNavy, margin: 0 });
  s.addNotes(`本页最重要的三件事：
1. 本周英文版封面要传达的核心不是“市场全面转牛”，而是一次典型的 macro relief rally。停火窗口带来的油价回落，先修复的是风险偏好和空头仓位，而不是所有基本面问题都解决了。
2. 监管进入关键时点，是这份周报最需要 LP 理解的背景。现在市场在重新评估美国是否会真正给稳定币、DeFi、代币化证券一套更可执行的框架，这会直接影响资金愿不愿意从观望转向配置。
3. 行业内部最值得关注的是结构性赢家已经开始浮现，尤其是 Hyperliquid、Base、以及 AI Agent 支付轨道。这些方向背后都有真实用户和真实工作负载，不再只是叙事。`);
}

// 2
{
  const s = pptx.addSlide();
  s.background = { color: C.offWhite };
  addHeader(s, '📊  Market Overview');
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
  s.addText('Fear & Greed', { x: x0 + (w + gap) * 2 + 0.22, y: y + 0.1, w: 2.2, h: 0.25, fontFace: 'Arial', fontSize: 12, bold: true, color: C.medGray, margin: 0 });
  s.addText('16', { x: x0 + (w + gap) * 2 + 0.22, y: y + 0.35, w: 0.8, h: 0.35, fontFace: 'Arial', fontSize: 28, bold: true, color: C.red, margin: 0 });
  s.addText('Extreme Fear', { x: x0 + (w + gap) * 2 + 1.0, y: y + 0.43, w: 1.5, h: 0.25, fontFace: 'Arial', fontSize: 13, color: C.red, margin: 0 });
  s.addTable([
    [
      { text: 'Asset', options: { bold: true, color: C.white, fill: { color: C.navy }, align: 'center', fontFace: 'Arial', fontSize: 11 } },
      { text: 'Price', options: { bold: true, color: C.white, fill: { color: C.navy }, align: 'center', fontFace: 'Arial', fontSize: 11 } },
      { text: '7d', options: { bold: true, color: C.white, fill: { color: C.navy }, align: 'center', fontFace: 'Arial', fontSize: 11 } },
      { text: 'Notes', options: { bold: true, color: C.white, fill: { color: C.navy }, align: 'center', fontFace: 'Arial', fontSize: 11 } }
    ],
    ['SOL', '$81.87', '+3.0%', 'Still trading with security overhang'],
    ['HYPE', '$41.05', '+15.7%', 'Perp DEX leader, share still rising'],
    ['ENA', '$0.09', '+16.6%', 'High beta rebound'],
    ['PENDLE', '$1.05', '+5.8%', 'Yield sector recovery bounce'],
    ['DYDX', '$0.09', '-4.6%', 'Continued share loss']
  ], {
    x: 0.6, y: 2.25, w: 8.8, colW: [1.2, 1.7, 1.2, 4.7], rowH: [0.35, 0.33, 0.33, 0.33, 0.33, 0.33],
    border: { pt: 0.4, color: C.lightGray }, fontFace: 'Arial', fontSize: 10, color: C.bodyText
  });
  s.addText('Total market cap at $2.49T, BTC dominance 56.9%, 24h volume $77.0B. Macro de-escalation and lower oil drove short covering.', { x: 0.6, y: 4.55, w: 8.8, h: 0.35, fontFace: 'Arial', fontSize: 10, italic: true, color: C.medGray, margin: 0 });
  addFooter(s, 'Crypto Market Weekly | 2026.04.06–04.12');
  s.addNotes(`本页最重要的三件事：
1. BTC 与 ETH 同时走强，尤其 ETH 跑赢 BTC，说明这轮上涨并非只有避险式回补，而是部分资金开始提前布局具有后续催化剂的核心资产。市场愿意为“有升级、有机构配置空间”的资产支付更高弹性。
2. Fear & Greed 虽然仍在 Extreme Fear，但已经从更低的极端位置回升。这种状态通常对应“市场不再恐慌性抛售，但还没有重建趋势信仰”。因此策略上更适合结构性配置，不适合把单周反弹外推成趋势行情。
3. Portfolio watchlist 里，HYPE 明显最强，PENDLE 和 ENA 属于风险偏好修复带来的弹性回升，DYDX 继续相对弱势。这说明资金在赛道内部也在集中，未来如果市场再有增量，强者可能继续吃掉大部分贝塔。`);
}

// 3
{
  const s = pptx.addSlide();
  s.background = { color: C.offWhite };
  addHeader(s, '🔥  Top Stories & Industry');
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 4.1, fill: { color: C.white }, line: { color: C.white }, shadow: makeShadow() });
  s.addText('TOP 5 STORIES', { x: 0.6, y: 0.96, w: 3.5, h: 0.3, fontFace: 'Arial', fontSize: 13, bold: true, color: C.navy, margin: 0 });
  const stories = [
    ['1', 'Macro ceasefire rally', 'Oil down, short squeeze lifts BTC and ETH'],
    ['2', 'CLARITY Act enters critical week', 'Stablecoin rewards and DeFi status in focus'],
    ['3', 'Hyperliquid extends leadership', 'Perp DEX/CEX ratio rises to ~21%'],
    ['4', 'AI agents move on-chain', 'x402 and AgentFi rails start to close the loop'],
    ['5', 'Solana still digesting Drift fallout', 'Security and governance remain key discounts']
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
    'Circle previews quantum-resistant L1 Arc',
    'Base TVL around $5.6B, L2 market further consolidates',
    'Ether.fi migrates roughly $5.7B ecosystem to OP Mainnet',
    'SKALE on Base launches for AI agent workloads'
  ]);
  panel(3.05, 'DeFi & Perp DEXes', [
    'RWA TVL rises to roughly $8B as yield capital rotates',
    'Aave commands about 30% of DeFi market share',
    'Hyperliquid hits $185.5B in 30d volume, about 34% share',
    'Aster emerges as the strongest second-tier venue'
  ]);
  addFooter(s, 'Crypto Market Weekly | 2026.04.06–04.12');
  s.addNotes(`本页最重要的三件事：
1. 这页的重点是行业内部的“胜者集中”。Base、Aave、Hyperliquid 都在各自子赛道继续强化网络效应。震荡市里，用户和流动性更倾向于向头部平台收缩，这种集中化通常会让龙头的商业模式变得更清晰。
2. Hyperliquid 的持续扩张说明链上衍生品正在从边缘市场变成主流市场基础设施。Perp DEX/CEX 比率升到约 21%，意味着越来越多交易者愿意把高频核心交易放在链上完成，这对交易费收入、清算逻辑和长期估值都很关键。
3. L1/L2 竞争也在变化。市场不再只看性能参数，而是看谁能承接稳定币、RWA、AI agent 这些真实需求。Circle 的 Arc、SKALE on Base、Ether.fi 的迁移，本质上都是在争夺“下一阶段真实工作负载会落在哪条链”。`);
}

// 4
{
  const s = pptx.addSlide();
  s.background = { color: C.offWhite };
  addHeader(s, '⚖️  Regulation & Fundraising');
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 4.1, fill: { color: C.white }, line: { color: C.white }, shadow: makeShadow() });
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 0.06, h: 4.1, fill: { color: C.gold }, line: { color: C.gold } });
  s.addText('Regulation & Policy', { x: 0.7, y: 0.96, w: 3.5, h: 0.3, fontFace: 'Arial', fontSize: 14, bold: true, color: C.navy, margin: 0 });
  const regs = [
    ['CLARITY Act critical window', 'Stablecoin rewards and DeFi classification remain key issues'],
    ['SEC approves Nasdaq tokenized securities pilot', 'Tokenized capital markets move closer to reality'],
    ['Hong Kong grants first stablecoin licenses', 'Asian regulated competition is accelerating'],
    ['UK bans crypto political donations', 'Election-related oversight is tightening'],
    ['SEC/CFTC non-security framework begins execution', 'Spot market rules are becoming clearer']
  ];
  regs.forEach((r, i) => {
    const yy = 1.45 + i * 0.65;
    s.addText(r[0], { x: 0.7, y: yy, w: 3.9, h: 0.22, fontFace: 'Arial', fontSize: 11, bold: true, color: C.bodyText, margin: 0 });
    s.addText(r[1], { x: 0.7, y: yy + 0.23, w: 3.9, h: 0.2, fontFace: 'Arial', fontSize: 9.5, color: C.medGray, margin: 0 });
  });
  s.addShape(pptx.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 4.5, h: 4.1, fill: { color: C.white }, line: { color: C.white }, shadow: makeShadow() });
  s.addShape(pptx.shapes.RECTANGLE, { x: 5.1, y: 0.9, w: 0.06, h: 4.1, fill: { color: C.gold }, line: { color: C.gold } });
  s.addText('Fundraising & M&A', { x: 5.4, y: 0.96, w: 3.5, h: 0.3, fontFace: 'Arial', fontSize: 14, bold: true, color: C.navy, margin: 0 });
  s.addTable([
    [
      { text: 'Project', options: { bold: true, color: C.white, fill: { color: C.navy }, fontFace: 'Arial', fontSize: 10 } },
      { text: 'Amount', options: { bold: true, color: C.white, fill: { color: C.navy }, fontFace: 'Arial', fontSize: 10, align: 'right' } },
      { text: 'Notes', options: { bold: true, color: C.white, fill: { color: C.navy }, fontFace: 'Arial', fontSize: 10 } }
    ],
    ['Pharos', '$44M', 'RWA / DePIN-focused L1'],
    ['BVNK + Mastercard', '$1.8B', 'Stablecoin payments consolidation'],
    ['Ozak AI', '$6.65M', 'AI financial agents'],
    ['Kulipa', '$6.2M', 'Non-custodial wallet cards']
  ], {
    x: 5.3, y: 1.45, w: 4.1, colW: [1.2, 0.9, 2.0], rowH: [0.3, 0.32, 0.32, 0.32, 0.32],
    border: { pt: 0.4, color: C.lightGray }, fontFace: 'Arial', fontSize: 9.5, color: C.bodyText
  });
  s.addShape(pptx.shapes.RECTANGLE, { x: 5.3, y: 3.35, w: 4.1, h: 0.85, fill: { color: C.paleGold }, line: { color: C.paleGold }, shadow: makeShadow() });
  s.addText('Primary-market capital remains concentrated in three themes: stablecoin payments, RWA rails, and AI-integrated crypto infrastructure.', { x: 5.55, y: 3.52, w: 3.6, h: 0.45, fontFace: 'Arial', fontSize: 10, color: C.bodyText, margin: 0, valign: 'mid' });
  addFooter(s, 'Crypto Market Weekly | 2026.04.06–04.12');
  s.addNotes(`本页最重要的三件事：
1. Regulation is the key near-term valuation driver. 如果 CLARITY Act 能在未来几周形成更明确的进展，市场会优先重估稳定币、DeFi 和代币化证券相关资产，因为这些资产过去一直承受合规折价。
2. Nasdaq tokenized securities pilot 的价值不在于短期成交量，而在于监管层第一次允许传统证券市场以更具体的方式试水上链。这会让市场开始认真讨论“哪条链、哪类托管、哪类合规基础设施”能承接未来的真实证券流通。
3. Fundraising and M&A are confirming the same pattern: capital is no longer chasing empty narratives. 传统支付公司和金融基础设施玩家开始直接买入 stablecoin rails，而一级市场持续给 RWA 和 AI+Crypto 供血，说明行业的增长逻辑正从 token beta 迁移到 infrastructure alpha。`);
}

// 5
{
  const s = pptx.addSlide();
  s.background = { color: C.offWhite };
  addHeader(s, '🛡️  Security & Key Sectors');
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 4.1, fill: { color: C.white }, line: { color: C.white }, shadow: makeShadow() });
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 4.5, h: 0.35, fill: { color: C.red }, line: { color: C.red } });
  s.addText('Security Incidents', { x: 0.65, y: 0.9, w: 3.5, h: 0.35, fontFace: 'Arial', fontSize: 12, bold: true, color: C.white, valign: 'middle', margin: 0 });
  s.addTable([
    [
      { text: 'Project', options: { bold: true, color: C.white, fill: { color: C.red }, fontFace: 'Arial', fontSize: 10 } },
      { text: 'Loss', options: { bold: true, color: C.white, fill: { color: C.red }, fontFace: 'Arial', fontSize: 10 } },
      { text: 'Comment', options: { bold: true, color: C.white, fill: { color: C.red }, fontFace: 'Arial', fontSize: 10 } }
    ],
    ['Drift', '$280M–$285M', 'Social engineering + durable nonce + fake asset whitelist'],
    ['Resolv Labs', '~$25M', 'AWS KMS compromise enabled mint exploit'],
    ['Stabble', 'Warning', 'Team flagged suspected DPRK insider risk']
  ], {
    x: 0.6, y: 1.35, w: 4.05, colW: [0.95, 1.0, 2.1], rowH: [0.28, 0.34, 0.34, 0.34],
    border: { pt: 0.4, color: C.lightGray }, fontFace: 'Arial', fontSize: 9.2, color: C.bodyText
  });
  s.addText('Stablecoins', { x: 0.6, y: 3.0, w: 2, h: 0.25, fontFace: 'Arial', fontSize: 12, bold: true, color: C.navy, margin: 0 });
  s.addText([
    { text: '• ', options: { color: C.gold } }, { text: 'Total market cap above $250B', options: { color: C.bodyText } },
    { text: '\n• ', options: { color: C.gold } }, { text: 'Hong Kong grants first issuer licenses', options: { color: C.bodyText } },
    { text: '\n• ', options: { color: C.gold } }, { text: 'US reward-wrapper debate remains unresolved', options: { color: C.bodyText } }
  ], { x: 0.7, y: 3.35, w: 3.9, h: 1.15, fontFace: 'Arial', fontSize: 10, margin: 0 });
  const box = (y, title, items) => {
    s.addShape(pptx.shapes.RECTANGLE, { x: 5.1, y, w: 4.5, h: 1.95, fill: { color: C.white }, line: { color: C.white }, shadow: makeShadow() });
    s.addShape(pptx.shapes.RECTANGLE, { x: 5.1, y, w: 4.5, h: 0.35, fill: { color: C.navy }, line: { color: C.navy } });
    s.addText(title, { x: 5.3, y, w: 3.9, h: 0.35, fontFace: 'Arial', fontSize: 11, bold: true, color: C.white, valign: 'middle', margin: 0 });
    s.addText(items.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < items.length - 1, fontFace: 'Arial', fontSize: 10, color: C.bodyText } })), { x: 5.3, y: y + 0.42, w: 4.0, h: 1.35, margin: 0, paraSpaceAfterPt: 5 });
  };
  box(0.9, 'AI × Crypto', [
    'AgentFi moves into real treasury and execution use cases',
    'AlphaTON raises $43M for TON-focused decentralized compute',
    'Covenant AI exits Bittensor, exposing governance concerns',
    'x402 remains the key payment rail to watch'
  ]);
  box(3.05, 'Prediction Markets', [
    'Polymarket and Kalshi increasingly treated as data infra',
    'BTC-linked markets reprice quickly with spot moves',
    'Energy and geopolitics absorb major new flow',
    'Robinhood, IBKR, Crypto.com continue expanding access'
  ]);
  addFooter(s, 'Crypto Market Weekly | 2026.04.06–04.12');
  s.addNotes(`本页最重要的三件事：
1. Security is still the most underpriced systemic risk in crypto. Drift 的案例说明，真正危险的不只是 smart contract bug，而是权限设计、人员安全、签名流程、云基础设施等“链下攻击面”。未来市场会继续给安全架构薄弱的协议打折。
2. Stablecoins and AI × Crypto are converging into the same long-term story: who controls settlement and execution for real on-chain economic activity. 稳定币提供结算层，AI agent 提供自动执行层，如果支付标准成熟，两者会形成很强的复利网络效应。
3. Prediction markets are evolving from speculative products into real-time information infrastructure. 它们越来越像可以量化的预期数据层。对投资研究而言，这种市场化概率信号未来会越来越重要，特别是在宏观、监管和事件驱动交易中。`);
}

// 6
{
  const s = pptx.addSlide();
  s.background = { color: C.offWhite };
  addHeader(s, '🔮  Catalysts & Outlook');
  s.addTable([
    [
      { text: 'Date', options: { bold: true, color: C.white, fill: { color: C.navy }, fontFace: 'Arial', fontSize: 10 } },
      { text: 'Event', options: { bold: true, color: C.white, fill: { color: C.navy }, fontFace: 'Arial', fontSize: 10 } },
      { text: 'Assets', options: { bold: true, color: C.white, fill: { color: C.navy }, fontFace: 'Arial', fontSize: 10 } },
      { text: 'Expected Impact', options: { bold: true, color: C.white, fill: { color: C.navy }, fontFace: 'Arial', fontSize: 10 } }
    ],
    ['Apr 13–18', 'CLARITY Act push', 'Stablecoins / DeFi / broad market', 'Regulatory repricing'],
    ['Apr 15', 'STRK unlock', 'STRK', 'Moderate supply pressure'],
    ['Apr 16', 'ARB unlock', 'ARB', 'Higher L2 volatility'],
    ['Apr 17', 'deBridge unlock', 'DBR', 'Major supply shock'],
    ['Apr 17–24', 'Flare FIP.16 vote', 'FLR', 'Changes inflation and staking logic'],
    ['Apr 27–29', 'Bitcoin 2026', 'BTC / BTCfi', 'Potential institutional catalyst']
  ], {
    x: 0.45, y: 1.0, w: 6.2, colW: [1.05, 2.2, 1.35, 1.6], rowH: [0.32, 0.34, 0.3, 0.3, 0.3, 0.3, 0.32],
    border: { pt: 0.4, color: C.lightGray }, fontFace: 'Arial', fontSize: 9.3, color: C.bodyText
  });
  s.addShape(pptx.shapes.RECTANGLE, { x: 6.95, y: 1.0, w: 2.6, h: 3.9, fill: { color: C.white }, line: { color: C.white }, shadow: makeShadow() });
  s.addShape(pptx.shapes.RECTANGLE, { x: 6.95, y: 1.0, w: 2.6, h: 0.35, fill: { color: C.navy }, line: { color: C.navy } });
  s.addText('Weekly Takeaway', { x: 7.12, y: 1.0, w: 2.1, h: 0.35, fontFace: 'Arial', fontSize: 11, bold: true, color: C.white, valign: 'middle', margin: 0 });
  s.addText([
    { text: '• ', options: { color: C.gold } }, { text: 'Market shifted from panic to selective repair', options: { color: C.bodyText } },
    { text: '\n• ', options: { color: C.gold } }, { text: 'Regulation is the key external variable for the next 2 weeks', options: { color: C.bodyText } },
    { text: '\n• ', options: { color: C.gold } }, { text: 'HYPE remains the strongest portfolio-linked trend', options: { color: C.bodyText } },
    { text: '\n• ', options: { color: C.gold } }, { text: 'PENDLE / ENA are recovery-beta trades if sentiment improves', options: { color: C.bodyText } }
  ], { x: 7.18, y: 1.48, w: 2.0, h: 1.9, fontFace: 'Arial', fontSize: 10, margin: 0 });
  s.addShape(pptx.shapes.RECTANGLE, { x: 7.15, y: 3.55, w: 2.05, h: 0.9, fill: { color: C.paleGold }, line: { color: C.paleGold } });
  s.addText('Portfolio lens:\nHYPE stays structurally strongest; PENDLE and ENA are leverage to improving risk appetite.', { x: 7.28, y: 3.72, w: 1.8, h: 0.5, fontFace: 'Arial', fontSize: 9.5, bold: true, color: C.navy, margin: 0, align: 'center' });
  addFooter(s, 'Crypto Market Weekly | 2026.04.06–04.12');
  s.addNotes(`本页最重要的三件事：
1. The next 1–2 weeks are mainly about two forces: regulation and supply. 如果 CLARITY Act 推进顺利，市场会优先修复那些长期被政策不确定性压制的赛道；如果法案受阻，市场很可能重新回到“只买最安全资产”的状态。
2. Unlocks are important because the market is still fragile. STRK、ARB、DBR 这些供给事件本身未必改变长期方向，但在成交量尚未完全恢复的环境里，会显著放大短期波动。所以 timing 和 position sizing 仍然很重要。
3. Portfolio-wise, HYPE remains the clearest structural winner because it is tied to the long-term rise of on-chain derivatives. PENDLE and ENA are more sentiment-sensitive. 如果监管和宏观环境继续改善，它们可能继续补涨；若环境再次恶化，资金大概率会重新回到 BTC 和少数龙头现金流资产。`);
}

pptx.writeFile('/Users/lx/Github/crypto-weekly-digest/reports/2026-04-06/slides-en.pptx');
