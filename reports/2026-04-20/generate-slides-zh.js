const pptxgen = require('pptxgenjs');

const C = {
  navy: '1A2744', darkNavy: '0F1A2E', gold: 'C9A96E', white: 'FFFFFF', offWhite: 'F5F6F8',
  lightGray: 'E8EBF0', medGray: '8C95A6', body: '2D3748', red: 'C53030', green: '2F855A', paleGold: 'FFF8E7'
};
const shadow = { type: 'outer', blur: 3, offset: 1, angle: 45, color: '000000', opacity: 0.12 };
const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = '42';
pres.title = '加密市场周报 2026.04.20-04.26';
pres.subject = 'Crypto Weekly Digest';
pres.company = '42';
pres.lang = 'zh-CN';

function header(slide, title) {
  slide.background = { color: C.offWhite };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy }, line: { color: C.navy } });
  slide.addText(title, { x: 0.6, y: 0, w: 8.8, h: 0.7, fontFace: 'Arial', fontSize: 20, bold: true, color: C.white, valign: 'mid', margin: 0 });
}
function footer(slide) {
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy }, line: { color: C.navy } });
  slide.addText('加密市场周报 | 2026.04.20–04.26', { x: 0.5, y: 5.25, w: 9, h: 0.375, fontFace: 'Arial', fontSize: 8, color: C.medGray, valign: 'mid', margin: 0 });
}
function card(slide, x, y, w, h, title, body, opts = {}) {
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: opts.fill || C.white }, line: { color: opts.line || C.lightGray, pt: 0.5 }, shadow });
  if (opts.bar !== false) slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.06, h, fill: { color: opts.barColor || C.gold }, line: { color: opts.barColor || C.gold } });
  slide.addText(title, { x: x + 0.2, y: y + 0.08, w: w - 0.3, h: 0.25, fontFace: 'Arial', fontSize: 12, bold: true, color: C.navy, margin: 0 });
  slide.addText(body, { x: x + 0.2, y: y + 0.35, w: w - 0.3, h: h - 0.45, fontFace: 'Arial', fontSize: opts.fontSize || 10, color: opts.color || C.body, margin: 0.02, breakLine: false, valign: 'top' });
}

// Slide 1
let s = pres.addSlide();
s.background = { color: C.darkNavy };
s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.6, w: 1.2, h: 0.04, fill: { color: C.gold }, line: { color: C.gold } });
s.addText('加密市场周报', { x: 0.8, y: 1.8, w: 8, h: 0.8, fontFace: 'Arial', fontSize: 42, bold: true, color: C.white, margin: 0 });
s.addText('2026年4月20日 — 4月26日', { x: 0.8, y: 2.7, w: 8, h: 0.4, fontFace: 'Arial', fontSize: 20, color: C.gold, margin: 0 });
s.addText('数据来源：Messari · The Block · CoinGecko · 另类恐惧贪婪指数', { x: 0.8, y: 3.35, w: 8.5, h: 0.3, fontFace: 'Arial', fontSize: 11, color: C.medGray, margin: 0 });
s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.gold }, line: { color: C.gold } });
s.addText('CONFIDENTIAL — FOR LP DISTRIBUTION ONLY', { x: 0, y: 5.25, w: 10, h: 0.375, fontFace: 'Arial', fontSize: 9, bold: true, color: C.darkNavy, align: 'center', valign: 'mid', margin: 0 });
s.addNotes(`本页重点第一，过去一周市场的核心不是全面 risk-on，而是资金重新集中到最确定的主线，尤其是 BTC 与少数基础设施龙头。Strategy 再度大额买入 BTC，让企业资产负债表配置比特币的叙事继续深化。

第二，监管与传统金融基础设施都在靠近链上资产。稳定币立法、链上证券豁免、交易所规则修改这些动作，说明市场正在从“能否合规存在”转向“如何在合规框架内扩张”。

第三，本周仍然不能忽视安全阴影。Kelp DAO / rsETH 事件提醒我们，哪怕宏观和监管边际改善，跨链与共享安全的结构性漏洞依然会反复冲击 DeFi 风险偏好。`);

// Slide 2
s = pres.addSlide();
header(s, '📊  市场概览');
const metrics = [
  ['BTC', '$77,934', '+3.2%', C.green],
  ['ETH', '$2,343', '+1.3%', C.green],
  ['恐惧贪婪指数', '33', 'Fear', C.red]
];
metrics.forEach((m, i) => {
  const x = 0.6 + i * 3.1;
  s.addShape(pres.shapes.RECTANGLE, { x, y: 0.95, w: 2.8, h: 1.0, fill: { color: C.white }, line: { color: C.lightGray, pt: 0.5 }, shadow });
  s.addShape(pres.shapes.RECTANGLE, { x, y: 0.95, w: 0.06, h: 1.0, fill: { color: i === 2 ? C.red : C.gold }, line: { color: i === 2 ? C.red : C.gold } });
  s.addText(m[0], { x: x + 0.2, y: 1.06, w: 2.2, h: 0.2, fontFace: 'Arial', fontSize: 11, color: C.medGray, bold: true, margin: 0 });
  s.addText(m[1], { x: x + 0.2, y: 1.3, w: 2.1, h: 0.28, fontFace: 'Arial', fontSize: 26, bold: true, color: C.body, margin: 0 });
  s.addText(m[2], { x: x + 0.2, y: 1.63, w: 1.8, h: 0.2, fontFace: 'Arial', fontSize: 13, bold: true, color: m[3], margin: 0 });
});
const table = [
  [{text:'资产',options:{bold:true,color:C.white,fill:C.navy,align:'center',fontSize:11}}, {text:'价格',options:{bold:true,color:C.white,fill:C.navy,align:'center',fontSize:11}}, {text:'7d',options:{bold:true,color:C.white,fill:C.navy,align:'center',fontSize:11}}, {text:'备注',options:{bold:true,color:C.white,fill:C.navy,align:'center',fontSize:11}}],
  [{text:'SOL',options:{bold:true,align:'center'}},{text:'$86.42',options:{align:'right'}},{text:'+0.5%',options:{align:'center',color:C.green,bold:true}},{text:'Alpenglow 预期支撑',options:{fontSize:10,color:C.medGray}}],
  [{text:'HYPE',options:{bold:true,align:'center'}},{text:'$41.30',options:{align:'right'}},{text:'-4.4%',options:{align:'center',color:C.red,bold:true}},{text:'龙头强势但波动放大',options:{fontSize:10,color:C.medGray}}],
  [{text:'ENA',options:{bold:true,align:'center'}},{text:'$0.11',options:{align:'right'}},{text:'-6.8%',options:{align:'center',color:C.red,bold:true}},{text:'稳定币监管博弈压制',options:{fontSize:10,color:C.medGray}}],
  [{text:'PENDLE',options:{bold:true,align:'center'}},{text:'$1.31',options:{align:'right'}},{text:'+0.8%',options:{align:'center',color:C.green,bold:true}},{text:'收益叙事仍在',options:{fontSize:10,color:C.medGray}}],
  [{text:'DYDX',options:{bold:true,align:'center'}},{text:'$0.16',options:{align:'right'}},{text:'+9.8%',options:{align:'center',color:C.green,bold:true}},{text:'低位反弹',options:{fontSize:10,color:C.medGray}}]
];
s.addTable(table, { x:0.6, y:2.2, w:8.8, colW:[1.2,1.6,1.1,4.9], rowH:[0.35,0.32,0.32,0.32,0.32,0.32], border:{pt:0.5,color:C.lightGray}, autoPage:false });
s.addText('总市值 $2.69T | BTC 主导率 58.1% | 24h 交易量 $50.2B。市场仍是“BTC 强、山寨弱”的结构性修复。', { x:0.6, y:4.52, w:8.8, h:0.35, fontFace:'Arial', fontSize:10, color:C.medGray, italic:true, margin:0 });
footer(s);
s.addNotes(`本页最重要的第一点，是 BTC 的强势来自真实资金结构而不只是情绪反弹。ETF 流入、Strategy 增持与企业级配置形成了更稳定的买盘来源，所以 BTC 表现明显强于 ETH 和多数山寨。

第二点，恐惧贪婪指数虽然回到 33，但仍然处于 Fear 区间，这意味着市场修复远未完成。资金愿意回到最有共识的资产，却还不愿意普遍下沉到高 beta 长尾资产。

第三点，对组合管理而言，HYPE、ENA、PENDLE、DYDX 的表现分化提醒我们，行业内部已经不是简单的 beta 交易。接下来更需要区分“被叙事驱动的反弹”与“有持续现金流或市场份额改善支撑的反弹”。`);

// Slide 3
s = pres.addSlide();
header(s, '🔥  本周要闻 & 行业动态');
card(s, 0.4, 0.9, 4.5, 4.1, '本周要闻 TOP 5',
  '1. Strategy 再买入 $2.54B BTC，企业配置叙事继续强化。\n\n2. CLARITY / GENIUS 进入收益分配博弈期。\n\n3. SEC + NYSE 同步推进链上证券基础设施。\n\n4. Hyperliquid 巩固永续 DEX 龙头。\n\n5. Kelp DAO / rsETH 攻击继续压制 DeFi 风险偏好。', { fontSize: 10 });
card(s, 5.1, 0.9, 4.5, 1.95, 'Layer 1 & Layer 2',
  '• Fluent 主网上线，尝试 blended execution。\n• Base 拿下约 50% L2 DEX 成交量。\n• TRON Q1 处理 $2T USDT 转账。\n• Stellar 借 RWA 借贷把 TVL 推到新高。', { fontSize: 10 });
card(s, 5.1, 3.05, 4.5, 1.95, 'DeFi & 永续合约',
  '• Aave 因 rsETH 事件面临约 $123M–$230M 坏账压力。\n• DeFi United 形成联合救助。\n• Hyperliquid 控制超 70% 链上永续 OI。\n• GMX 用 RWA 合成永续与 MegaETH 尝试反击。', { fontSize: 10 });
footer(s);
s.addNotes(`本页第一件最重要的事，是行业主线正在从“总量增长”切回“结构分化”。同样是加密市场，机构资金、监管资源和用户流动性都越来越集中到少数具备确定性的资产与平台。

第二，L1/L2 的竞争逻辑也在变化。现在真正决定份额的，不只是 TPS 或最终性，而是谁拥有分发入口、稳定币流量和生态协同。Base 的份额提升就是典型例子。

第三，永续 DEX 的赢家通吃已经非常明显。Hyperliquid 的意义不只是业绩更好，而是它说明链上交易基础设施已经开始具备对中心化衍生品平台的真实替代性。对组合而言，这一赛道仍是最值得跟踪的 alpha 来源之一。`);

// Slide 4
s = pres.addSlide();
header(s, '⚖️  监管政策 & 投融资');
card(s, 0.4, 0.9, 4.5, 4.1, '监管与政策',
  '• 美国市场结构法案推进放缓，焦点落在稳定币收益分配。\n\n• SEC 对部分 DeFi 界面给出更清晰口径。\n\n• Treasury / FDIC / FinCEN 开始推进 GENIUS 相关规则。\n\n• NYSE 推进 tokenized securities 规则变更。', { fontSize: 10 });
const fund = [
  [{text:'项目',options:{bold:true,color:C.white,fill:C.navy,fontSize:10}},{text:'金额',options:{bold:true,color:C.white,fill:C.navy,fontSize:10,align:'right'}},{text:'备注',options:{bold:true,color:C.white,fill:C.navy,fontSize:10}}],
  [{text:'Blockchain Capital',options:{bold:true,fontSize:10}},{text:'$700M',options:{align:'right',fontSize:10}},{text:'双基金募资',options:{fontSize:9,color:C.medGray}}],
  [{text:'Kraken / Bitnomial',options:{bold:true,fontSize:10}},{text:'$550M',options:{align:'right',fontSize:10}},{text:'扩张美國合规衍生品',options:{fontSize:9,color:C.medGray}}],
  [{text:'Polymarket',options:{bold:true,fontSize:10}},{text:'$400M',options:{align:'right',fontSize:10}},{text:'估值或达 $15B',options:{fontSize:9,color:C.medGray}}],
  [{text:'Kaio',options:{bold:true,fontSize:10}},{text:'$8M',options:{align:'right',fontSize:10}},{text:'Tether 参投的 RWA 项目',options:{fontSize:9,color:C.medGray}}]
];
s.addShape(pres.shapes.RECTANGLE, { x:5.1, y:0.9, w:4.5, h:4.1, fill:{color:C.white}, line:{color:C.lightGray, pt:0.5}, shadow });
s.addShape(pres.shapes.RECTANGLE, { x:5.1, y:0.9, w:0.06, h:4.1, fill:{color:C.gold}, line:{color:C.gold} });
s.addText('投融资与并购', { x:5.35, y:0.98, w:3.8, h:0.25, fontFace:'Arial', fontSize:12, bold:true, color:C.navy, margin:0 });
s.addTable(fund, { x:5.3, y:1.38, w:4.1, colW:[1.5,0.8,1.8], rowH:[0.3,0.3,0.3,0.3,0.3], border:{pt:0.5,color:C.lightGray}, autoPage:false });
s.addShape(pres.shapes.RECTANGLE, { x:5.3, y:3.35, w:4.1, h:0.95, fill:{color:C.paleGold}, line:{color:C.gold, pt:0.5} });
s.addText('关键信号：资本并未离场，而是在向更后期、更合规、能承接机构流量的平台集中。', { x:5.5, y:3.55, w:3.6, h:0.45, fontFace:'Arial', fontSize:10, color:C.body, bold:true, margin:0 });
footer(s);
s.addNotes(`这一页的核心，是监管与融资都在指向同一个方向：行业的“制度化”正在加速。立法虽然有拉扯，但争论点已经不是是否允许稳定币存在，而是收益应该如何分配、哪些主体可以分配，这本身就是成熟化信号。

第二，从投融资结构看，资金明显向大平台和合规基础设施集中。Kraken 收购 Bitnomial、Polymarket 高估值融资、头部 VC 募更大基金，说明资本市场正在挑选下一轮周期的关键入口，而不是做分散式撒网。

第三，这对市场的影响是估值分层会进一步加剧。真正能受益的，不是所有加密项目，而是那些处在监管友好区域、具备分发能力或拥有基础设施护城河的公司与协议。`);

// Slide 5
s = pres.addSlide();
header(s, '🛡️  安全事件 & 重点赛道');
card(s, 0.4, 0.9, 4.5, 2.45, '安全事件',
  '• Kelp DAO / rsETH：约 $293M，LayerZero DVN 脆弱性被利用，约 30,766 ETH 已冻结。\n\n• Volo：约 $3.5M。\n\n• VaultX：约 $2.1M。\n\n• BridgeNet：约 $0.8M。\n\n本周总损失约 $299M。', { fontSize: 10, barColor: C.red });
card(s, 0.4, 3.55, 4.5, 1.45, '稳定币',
  '• Morgan Stanley 切入储备管理。\n• Tether 冻结约 $344M USDT。\n• 收益型稳定币增速约为行业平均 15 倍。', { fontSize: 10 });
card(s, 5.1, 0.9, 4.5, 1.95, 'AI × Crypto',
  '• 资金继续流向 DeAI 基础设施。\n• Grayscale 申请 TAO ETF。\n• AI 既放大增长，也放大攻击能力。', { fontSize: 10 });
card(s, 5.1, 3.05, 4.5, 1.95, '预测市场',
  '• 监管诉讼与海外禁令同步升温。\n• Polymarket 与 Kalshi 计划切入 perps。\n• Polymarket 月成交量据报升至约 $10.57B。', { fontSize: 10 });
footer(s);
s.addNotes(`本页最关键的是，安全问题依然是阻碍 DeFi 和链上交易进一步机构化的最大现实门槛。Kelp DAO 事件不是单一协议失误，而是跨链验证、抵押品管理、借贷清算等多个环节耦合后的系统性事故。

第二，稳定币和 AI 两个赛道都在走向更深的基础设施化。稳定币的竞争正在从发行端转向储备管理、收益设计和合规分发；AI x Crypto 的竞争则从故事型 token 转向算力、验证和安全。

第三，预测市场的意义在于它开始从“事件投注平台”进化成更广义的信息定价基础设施。一旦它与 perps、AI agent 和更强的 oracle 体系结合，市场空间会被重新打开，但监管冲突也会更激烈。`);

// Slide 6
s = pres.addSlide();
header(s, '🔮  催化剂 & 下周展望');
s.addTable([
  [{text:'日期',options:{bold:true,color:C.white,fill:C.navy,fontSize:10}},{text:'事件',options:{bold:true,color:C.white,fill:C.navy,fontSize:10}},{text:'资产',options:{bold:true,color:C.white,fill:C.navy,fontSize:10}},{text:'影响',options:{bold:true,color:C.white,fill:C.navy,fontSize:10}}],
  [{text:'4/28',options:{fontSize:10}},{text:'SIGN 解锁',options:{fontSize:10,bold:true}},{text:'SIGN',options:{fontSize:10}},{text:'供给压力大',options:{fontSize:9,color:C.red,bold:true}}],
  [{text:'4/28',options:{fontSize:10}},{text:'JUP 解锁',options:{fontSize:10,bold:true}},{text:'JUP / SOL',options:{fontSize:10}},{text:'关注承接',options:{fontSize:9}}],
  [{text:'4/29-30',options:{fontSize:10}},{text:'TOKEN2049 Dubai',options:{fontSize:10,bold:true}},{text:'广义山寨',options:{fontSize:10}},{text:'叙事与合作密集',options:{fontSize:9,color:C.green,bold:true}}],
  [{text:'5/1',options:{fontSize:10}},{text:'SUI 解锁',options:{fontSize:10,bold:true}},{text:'SUI',options:{fontSize:10}},{text:'中等供给冲击',options:{fontSize:9}}],
  [{text:'5/5-7',options:{fontSize:10}},{text:'Consensus Miami',options:{fontSize:10,bold:true}},{text:'全市场',options:{fontSize:10}},{text:'机构表态窗口',options:{fontSize:9,color:C.green,bold:true}}],
  [{text:'5月',options:{fontSize:10}},{text:'ETH / BTCFi 技术催化',options:{fontSize:10,bold:true}},{text:'ETH / BTC',options:{fontSize:10}},{text:'看价值捕获修复',options:{fontSize:9}}]
], { x:0.45, y:0.95, w:6.0, colW:[0.8,2.1,1.1,2.0], rowH:[0.32,0.32,0.32,0.32,0.32,0.32,0.32], border:{pt:0.5,color:C.lightGray}, autoPage:false });
card(s, 6.7, 0.95, 2.85, 3.95, '周度结论',
  '1. BTC 仍是最清晰的机构共识交易。\n\n2. 监管与 TradFi 正在把稳定币、RWA、链上证券推向主流化。\n\n3. Perp DEX 龙头化继续强化，值得持续跟踪 HYPE 与相关基础设施。\n\n4. 安全风险仍压制 DeFi beta，择时上要更谨慎。', { fontSize: 10 });
footer(s);
s.addNotes(`下周最需要关注的第一类变量，是供给冲击。SIGN、JUP、SUI 这类解锁会测试市场对山寨资产的真实承接能力，如果承接弱，说明风险偏好并没有随着 BTC 反弹而同步恢复。

第二类变量是会议与叙事窗口。TOKEN2049 和 Consensus 往往不是直接决定价格的事件，但它们会显著影响资金接下来追逐的赛道，特别是 AI、稳定币、RWA 和 perp 相关基础设施。

第三类变量是技术与制度的中期催化。Ethereum 的扩容与价值捕获修复、BTCFi 主网上线、稳定币立法继续推进，这些因素决定了市场是否会从短期修复走向更持续的结构性重估。`);

pres.writeFile('/Users/lx/Github/crypto-weekly-digest/reports/2026-04-20/slides-zh.pptx');
