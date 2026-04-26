const pptxgen = require('pptxgenjs');

const C = {
  navy: '1A2744', darkNavy: '0F1A2E', gold: 'C9A96E', white: 'FFFFFF', offWhite: 'F5F6F8',
  lightGray: 'E8EBF0', medGray: '8C95A6', body: '2D3748', red: 'C53030', green: '2F855A', paleGold: 'FFF8E7'
};
const shadow = { type: 'outer', blur: 3, offset: 1, angle: 45, color: '000000', opacity: 0.12 };
const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = '42';
pres.title = 'Crypto Market Weekly Report 2026.04.20-04.26';
pres.subject = 'Crypto Weekly Digest';
pres.company = '42';
pres.lang = 'en-US';

function header(slide, title) {
  slide.background = { color: C.offWhite };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navy }, line: { color: C.navy } });
  slide.addText(title, { x: 0.6, y: 0, w: 8.8, h: 0.7, fontFace: 'Arial', fontSize: 20, bold: true, color: C.white, valign: 'mid', margin: 0 });
}
function footer(slide) {
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.navy }, line: { color: C.navy } });
  slide.addText('Crypto Market Weekly | 2026.04.20–04.26', { x: 0.5, y: 5.25, w: 9, h: 0.375, fontFace: 'Arial', fontSize: 8, color: C.medGray, valign: 'mid', margin: 0 });
}
function card(slide, x, y, w, h, title, body, opts = {}) {
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: opts.fill || C.white }, line: { color: opts.line || C.lightGray, pt: 0.5 }, shadow });
  if (opts.bar !== false) slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.06, h, fill: { color: opts.barColor || C.gold }, line: { color: opts.barColor || C.gold } });
  slide.addText(title, { x: x + 0.2, y: y + 0.08, w: w - 0.3, h: 0.25, fontFace: 'Arial', fontSize: 12, bold: true, color: C.navy, margin: 0 });
  slide.addText(body, { x: x + 0.2, y: y + 0.35, w: w - 0.3, h: h - 0.45, fontFace: 'Arial', fontSize: opts.fontSize || 10, color: opts.color || C.body, margin: 0.02, breakLine: false, valign: 'top' });
}

let s = pres.addSlide();
s.background = { color: C.darkNavy };
s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.6, w: 1.2, h: 0.04, fill: { color: C.gold }, line: { color: C.gold } });
s.addText('Crypto Market Weekly', { x: 0.8, y: 1.8, w: 8.2, h: 0.8, fontFace: 'Arial', fontSize: 42, bold: true, color: C.white, margin: 0 });
s.addText('April 20 – April 26, 2026', { x: 0.8, y: 2.7, w: 8, h: 0.4, fontFace: 'Arial', fontSize: 20, color: C.gold, margin: 0 });
s.addText('Sources: Messari · The Block · CoinGecko · Alternative.me Fear & Greed Index', { x: 0.8, y: 3.35, w: 8.8, h: 0.3, fontFace: 'Arial', fontSize: 11, color: C.medGray, margin: 0 });
s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.375, fill: { color: C.gold }, line: { color: C.gold } });
s.addText('CONFIDENTIAL — FOR LP DISTRIBUTION ONLY', { x: 0, y: 5.25, w: 10, h: 0.375, fontFace: 'Arial', fontSize: 9, bold: true, color: C.darkNavy, align: 'center', valign: 'mid', margin: 0 });
s.addNotes(`本页重点第一，过去一周市场的核心不是全面 risk-on，而是资金重新集中到最确定的主线，尤其是 BTC 与少数基础设施龙头。Strategy 再度大额买入 BTC，让企业资产负债表配置比特币的叙事继续深化。

第二，监管与传统金融基础设施都在靠近链上资产。稳定币立法、链上证券豁免、交易所规则修改这些动作，说明市场正在从“能否合规存在”转向“如何在合规框架内扩张”。

第三，本周仍然不能忽视安全阴影。Kelp DAO / rsETH 事件提醒我们，哪怕宏观和监管边际改善，跨链与共享安全的结构性漏洞依然会反复冲击 DeFi 风险偏好。`);

s = pres.addSlide();
header(s, '📊  Market Overview');
[['BTC','$77,934','+3.2%',C.green],['ETH','$2,343','+1.3%',C.green],['Fear & Greed','33','Fear',C.red]].forEach((m, i) => {
  const x = 0.6 + i * 3.1;
  s.addShape(pres.shapes.RECTANGLE, { x, y: 0.95, w: 2.8, h: 1.0, fill: { color: C.white }, line: { color: C.lightGray, pt: 0.5 }, shadow });
  s.addShape(pres.shapes.RECTANGLE, { x, y: 0.95, w: 0.06, h: 1.0, fill: { color: i === 2 ? C.red : C.gold }, line: { color: i === 2 ? C.red : C.gold } });
  s.addText(m[0], { x: x + 0.2, y: 1.06, w: 2.2, h: 0.2, fontFace: 'Arial', fontSize: 11, color: C.medGray, bold: true, margin: 0 });
  s.addText(m[1], { x: x + 0.2, y: 1.3, w: 2.1, h: 0.28, fontFace: 'Arial', fontSize: 26, bold: true, color: C.body, margin: 0 });
  s.addText(m[2], { x: x + 0.2, y: 1.63, w: 1.8, h: 0.2, fontFace: 'Arial', fontSize: 13, bold: true, color: m[3], margin: 0 });
});
s.addTable([
  [{text:'Asset',options:{bold:true,color:C.white,fill:C.navy,align:'center',fontSize:11}}, {text:'Price',options:{bold:true,color:C.white,fill:C.navy,align:'center',fontSize:11}}, {text:'7d',options:{bold:true,color:C.white,fill:C.navy,align:'center',fontSize:11}}, {text:'Notes',options:{bold:true,color:C.white,fill:C.navy,align:'center',fontSize:11}}],
  [{text:'SOL',options:{bold:true,align:'center'}},{text:'$86.42',options:{align:'right'}},{text:'+0.5%',options:{align:'center',color:C.green,bold:true}},{text:'Supported by Alpenglow expectations',options:{fontSize:10,color:C.medGray}}],
  [{text:'HYPE',options:{bold:true,align:'center'}},{text:'$41.30',options:{align:'right'}},{text:'-4.4%',options:{align:'center',color:C.red,bold:true}},{text:'Leader remains strong but more volatile',options:{fontSize:10,color:C.medGray}}],
  [{text:'ENA',options:{bold:true,align:'center'}},{text:'$0.11',options:{align:'right'}},{text:'-6.8%',options:{align:'center',color:C.red,bold:true}},{text:'Stablecoin regulation overhang',options:{fontSize:10,color:C.medGray}}],
  [{text:'PENDLE',options:{bold:true,align:'center'}},{text:'$1.31',options:{align:'right'}},{text:'+0.8%',options:{align:'center',color:C.green,bold:true}},{text:'Yield narrative still intact',options:{fontSize:10,color:C.medGray}}],
  [{text:'DYDX',options:{bold:true,align:'center'}},{text:'$0.16',options:{align:'right'}},{text:'+9.8%',options:{align:'center',color:C.green,bold:true}},{text:'Low-base rebound',options:{fontSize:10,color:C.medGray}}]
], { x:0.6, y:2.2, w:8.8, colW:[1.2,1.6,1.1,4.9], rowH:[0.35,0.32,0.32,0.32,0.32,0.32], border:{pt:0.5,color:C.lightGray}, autoPage:false });
s.addText('Total mcap $2.69T | BTC dominance 58.1% | 24h volume $50.2B. The tape still shows a BTC-led repair rather than broad altseason.', { x:0.6, y:4.52, w:8.8, h:0.35, fontFace:'Arial', fontSize:10, color:C.medGray, italic:true, margin:0 });
footer(s);
s.addNotes(`本页最重要的第一点，是 BTC 的强势来自真实资金结构而不只是情绪反弹。ETF 流入、Strategy 增持与企业级配置形成了更稳定的买盘来源，所以 BTC 表现明显强于 ETH 和多数山寨。

第二点，恐惧贪婪指数虽然回到 33，但仍然处于 Fear 区间，这意味着市场修复远未完成。资金愿意回到最有共识的资产，却还不愿意普遍下沉到高 beta 长尾资产。

第三点，对组合管理而言，HYPE、ENA、PENDLE、DYDX 的表现分化提醒我们，行业内部已经不是简单的 beta 交易。接下来更需要区分“被叙事驱动的反弹”与“有持续现金流或市场份额改善支撑的反弹”。`);

s = pres.addSlide();
header(s, '🔥  Top Stories & Industry');
card(s, 0.4, 0.9, 4.5, 4.1, 'Top 5 Stories', '1. Strategy bought another $2.54B of BTC.\n\n2. CLARITY / GENIUS moved deeper into yield-distribution debates.\n\n3. SEC and NYSE both advanced tokenized securities infrastructure.\n\n4. Hyperliquid strengthened perp DEX leadership.\n\n5. The Kelp DAO / rsETH exploit kept DeFi risk appetite under pressure.', { fontSize: 10 });
card(s, 5.1, 0.9, 4.5, 1.95, 'Layer 1 & Layer 2', '• Fluent mainnet launched blended execution.\n• Base captured about 50% of L2 DEX volume.\n• TRON processed $2T of USDT in Q1.\n• Stellar hit a new TVL high via RWA lending.', { fontSize: 10 });
card(s, 5.1, 3.05, 4.5, 1.95, 'DeFi & Perp DEXes', '• Aave faces roughly $123M-$230M bad debt pressure from rsETH fallout.\n• DeFi United formed a coordinated rescue plan.\n• Hyperliquid controls 70%+ of on-chain perp OI.\n• GMX is countering with RWA synthetics and MegaETH.', { fontSize: 10 });
footer(s);
s.addNotes(`本页第一件最重要的事，是行业主线正在从“总量增长”切回“结构分化”。同样是加密市场，机构资金、监管资源和用户流动性都越来越集中到少数具备确定性的资产与平台。

第二，L1/L2 的竞争逻辑也在变化。现在真正决定份额的，不只是 TPS 或最终性，而是谁拥有分发入口、稳定币流量和生态协同。Base 的份额提升就是典型例子。

第三，永续 DEX 的赢家通吃已经非常明显。Hyperliquid 的意义不只是业绩更好，而是它说明链上交易基础设施已经开始具备对中心化衍生品平台的真实替代性。对组合而言，这一赛道仍是最值得跟踪的 alpha 来源之一。`);

s = pres.addSlide();
header(s, '⚖️  Regulation & Fundraising');
card(s, 0.4, 0.9, 4.5, 4.1, 'Regulation & Policy', '• US market structure legislation slowed, with stablecoin yield distribution as the key fault line.\n\n• The SEC gave clearer signals on some DeFi interfaces.\n\n• Treasury, FDIC, and FinCEN moved into GENIUS implementation work.\n\n• NYSE pushed rule changes for tokenized securities.', { fontSize: 10 });
s.addShape(pres.shapes.RECTANGLE, { x:5.1, y:0.9, w:4.5, h:4.1, fill:{color:C.white}, line:{color:C.lightGray, pt:0.5}, shadow });
s.addShape(pres.shapes.RECTANGLE, { x:5.1, y:0.9, w:0.06, h:4.1, fill:{color:C.gold}, line:{color:C.gold} });
s.addText('Fundraising & M&A', { x:5.35, y:0.98, w:3.8, h:0.25, fontFace:'Arial', fontSize:12, bold:true, color:C.navy, margin:0 });
s.addTable([
  [{text:'Project',options:{bold:true,color:C.white,fill:C.navy,fontSize:10}},{text:'Amount',options:{bold:true,color:C.white,fill:C.navy,fontSize:10,align:'right'}},{text:'Notes',options:{bold:true,color:C.white,fill:C.navy,fontSize:10}}],
  [{text:'Blockchain Capital',options:{bold:true,fontSize:10}},{text:'$700M',options:{align:'right',fontSize:10}},{text:'New dual-fund raise',options:{fontSize:9,color:C.medGray}}],
  [{text:'Kraken / Bitnomial',options:{bold:true,fontSize:10}},{text:'$550M',options:{align:'right',fontSize:10}},{text:'US regulated derivatives expansion',options:{fontSize:9,color:C.medGray}}],
  [{text:'Polymarket',options:{bold:true,fontSize:10}},{text:'$400M',options:{align:'right',fontSize:10}},{text:'Reported at up to $15B valuation',options:{fontSize:9,color:C.medGray}}],
  [{text:'Kaio',options:{bold:true,fontSize:10}},{text:'$8M',options:{align:'right',fontSize:10}},{text:'RWA round with Tether participation',options:{fontSize:9,color:C.medGray}}]
], { x:5.3, y:1.38, w:4.1, colW:[1.5,0.8,1.8], rowH:[0.3,0.3,0.3,0.3,0.3], border:{pt:0.5,color:C.lightGray}, autoPage:false });
s.addShape(pres.shapes.RECTANGLE, { x:5.3, y:3.35, w:4.1, h:0.95, fill:{color:C.paleGold}, line:{color:C.gold, pt:0.5} });
s.addText('Key takeaway: capital is not leaving. It is concentrating into later-stage, more compliant, institution-ready platforms.', { x:5.5, y:3.52, w:3.6, h:0.45, fontFace:'Arial', fontSize:10, color:C.body, bold:true, margin:0 });
footer(s);
s.addNotes(`这一页的核心，是监管与融资都在指向同一个方向：行业的“制度化”正在加速。立法虽然有拉扯，但争论点已经不是是否允许稳定币存在，而是收益应该如何分配、哪些主体可以分配，这本身就是成熟化信号。

第二，从投融资结构看，资金明显向大平台和合规基础设施集中。Kraken 收购 Bitnomial、Polymarket 高估值融资、头部 VC 募更大基金，说明资本市场正在挑选下一轮周期的关键入口，而不是做分散式撒网。

第三，这对市场的影响是估值分层会进一步加剧。真正能受益的，不是所有加密项目，而是那些处在监管友好区域、具备分发能力或拥有基础设施护城河的公司与协议。`);

s = pres.addSlide();
header(s, '🛡️  Security & Focus Sectors');
card(s, 0.4, 0.9, 4.5, 2.45, 'Security Incidents', '• Kelp DAO / rsETH: about $293M, exploiting LayerZero DVN weakness; ~30,766 ETH has been frozen.\n\n• Volo: about $3.5M.\n\n• VaultX: about $2.1M.\n\n• BridgeNet: about $0.8M.\n\nTotal weekly losses: about $299M.', { fontSize: 10, barColor: C.red });
card(s, 0.4, 3.55, 4.5, 1.45, 'Stablecoins', '• Morgan Stanley entered reserve management.\n• Tether froze about $344M of USDT.\n• Yield-bearing stablecoins are growing roughly 15x faster than the market.', { fontSize: 10 });
card(s, 5.1, 0.9, 4.5, 1.95, 'AI × Crypto', '• Capital keeps rotating toward DeAI infrastructure.\n• Grayscale filed for a TAO ETF.\n• AI is amplifying both growth and attack surfaces.', { fontSize: 10 });
card(s, 5.1, 3.05, 4.5, 1.95, 'Prediction Markets', '• Lawsuits and overseas bans both intensified.\n• Polymarket and Kalshi plan to enter perps.\n• Polymarket monthly volume reportedly reached about $10.57B.', { fontSize: 10 });
footer(s);
s.addNotes(`本页最关键的是，安全问题依然是阻碍 DeFi 和链上交易进一步机构化的最大现实门槛。Kelp DAO 事件不是单一协议失误，而是跨链验证、抵押品管理、借贷清算等多个环节耦合后的系统性事故。

第二，稳定币和 AI 两个赛道都在走向更深的基础设施化。稳定币的竞争正在从发行端转向储备管理、收益设计和合规分发；AI x Crypto 的竞争则从故事型 token 转向算力、验证和安全。

第三，预测市场的意义在于它开始从“事件投注平台”进化成更广义的信息定价基础设施。一旦它与 perps、AI agent 和更强的 oracle 体系结合，市场空间会被重新打开，但监管冲突也会更激烈。`);

s = pres.addSlide();
header(s, '🔮  Catalysts & Outlook');
s.addTable([
  [{text:'Date',options:{bold:true,color:C.white,fill:C.navy,fontSize:10}},{text:'Event',options:{bold:true,color:C.white,fill:C.navy,fontSize:10}},{text:'Assets',options:{bold:true,color:C.white,fill:C.navy,fontSize:10}},{text:'Impact',options:{bold:true,color:C.white,fill:C.navy,fontSize:10}}],
  [{text:'Apr 28',options:{fontSize:10}},{text:'SIGN unlock',options:{fontSize:10,bold:true}},{text:'SIGN',options:{fontSize:10}},{text:'Large supply pressure',options:{fontSize:9,color:C.red,bold:true}}],
  [{text:'Apr 28',options:{fontSize:10}},{text:'JUP unlock',options:{fontSize:10,bold:true}},{text:'JUP / SOL',options:{fontSize:10}},{text:'Watch absorption',options:{fontSize:9}}],
  [{text:'Apr 29-30',options:{fontSize:10}},{text:'TOKEN2049 Dubai',options:{fontSize:10,bold:true}},{text:'Alt / VC ecosystem',options:{fontSize:10}},{text:'Dense narrative and partnership window',options:{fontSize:9,color:C.green,bold:true}}],
  [{text:'May 1',options:{fontSize:10}},{text:'SUI unlock',options:{fontSize:10,bold:true}},{text:'SUI',options:{fontSize:10}},{text:'Medium supply shock',options:{fontSize:9}}],
  [{text:'May 5-7',options:{fontSize:10}},{text:'Consensus Miami',options:{fontSize:10,bold:true}},{text:'Broad market',options:{fontSize:10}},{text:'Institutional messaging window',options:{fontSize:9,color:C.green,bold:true}}],
  [{text:'May',options:{fontSize:10}},{text:'ETH / BTCFi technical catalysts',options:{fontSize:10,bold:true}},{text:'ETH / BTC',options:{fontSize:10}},{text:'Watch for value-capture repair',options:{fontSize:9}}]
], { x:0.45, y:0.95, w:6.0, colW:[0.8,2.1,1.1,2.0], rowH:[0.32,0.32,0.32,0.32,0.32,0.32,0.32], border:{pt:0.5,color:C.lightGray}, autoPage:false });
card(s, 6.7, 0.95, 2.85, 3.95, 'Weekly Conclusion', '1. BTC remains the cleanest institutional consensus trade.\n\n2. Regulation and TradFi are pulling stablecoins, RWAs, and tokenized securities closer to the mainstream.\n\n3. Perp DEX leadership continues to consolidate, keeping HYPE and related infra in focus.\n\n4. Security risk still caps DeFi beta, so timing needs to stay selective.', { fontSize: 10 });
footer(s);
s.addNotes(`下周最需要关注的第一类变量，是供给冲击。SIGN、JUP、SUI 这类解锁会测试市场对山寨资产的真实承接能力，如果承接弱，说明风险偏好并没有随着 BTC 反弹而同步恢复。

第二类变量是会议与叙事窗口。TOKEN2049 和 Consensus 往往不是直接决定价格的事件，但它们会显著影响资金接下来追逐的赛道，特别是 AI、稳定币、RWA 和 perp 相关基础设施。

第三类变量是技术与制度的中期催化。Ethereum 的扩容与价值捕获修复、BTCFi 主网上线、稳定币立法继续推进，这些因素决定了市场是否会从短期修复走向更持续的结构性重估。`);

pres.writeFile('/Users/lx/Github/crypto-weekly-digest/reports/2026-04-20/slides-en.pptx');
