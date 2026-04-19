const pptxgen = require('pptxgenjs');

const C = {
  navy: '1A2744', darkNavy: '0F1A2E', gold: 'C9A96E', white: 'FFFFFF', offWhite: 'F5F6F8',
  lightGray: 'E8EBF0', medGray: '8C95A6', bodyText: '2D3748', red: 'C53030', green: '2F855A', paleGold: 'FFF8E7'
};
const pptx = new pptxgen();
pptx.layout = 'LAYOUT_16x9'; pptx.author = '42'; pptx.company = '42'; pptx.subject = 'Crypto Weekly Digest'; pptx.title = 'Crypto Market Weekly 2026.04.13–04.19'; pptx.lang = 'en-US';
const makeShadow = () => ({ type: 'outer', blur: 3, offset: 1, angle: 45, color: '000000', opacity: 0.12 });
const addFooter = (s, t) => { s.addShape(pptx.shapes.RECTANGLE,{x:0,y:5.25,w:10,h:0.375,fill:{color:C.navy},line:{color:C.navy}}); s.addText(t,{x:0.45,y:5.25,w:9.1,h:0.375,fontFace:'Arial',fontSize:8,color:C.medGray,valign:'middle',margin:0}); };
const addHeader = (s, t) => { s.addShape(pptx.shapes.RECTANGLE,{x:0,y:0,w:10,h:0.7,fill:{color:C.navy},line:{color:C.navy}}); s.addText(t,{x:0.6,y:0,w:8.8,h:0.7,fontFace:'Arial',fontSize:20,bold:true,color:C.white,valign:'middle',margin:0}); };

{
  const s = pptx.addSlide(); s.background = { color: C.darkNavy };
  s.addShape(pptx.shapes.RECTANGLE,{x:0.8,y:1.55,w:1.2,h:0.04,fill:{color:C.gold},line:{color:C.gold}});
  s.addText('Crypto Market Weekly',{x:0.8,y:1.8,w:6,h:0.8,fontFace:'Arial',fontSize:42,bold:true,color:C.white,margin:0});
  s.addText('April 13 – April 19, 2026',{x:0.8,y:2.7,w:6,h:0.45,fontFace:'Arial',fontSize:20,color:C.gold,margin:0});
  s.addText('Sources: Messari · The Block · CoinGecko · DefiLlama',{x:0.8,y:3.4,w:6.8,h:0.3,fontFace:'Arial',fontSize:11,color:C.medGray,margin:0});
  s.addText('Themes: Kelp exploit, institutional inflows, CLARITY Act, Hyperliquid concentration, agent commerce',{x:0.8,y:4.15,w:8.5,h:0.55,fontFace:'Arial',fontSize:14,color:C.white,margin:0});
  s.addShape(pptx.shapes.RECTANGLE,{x:0,y:5.25,w:10,h:0.375,fill:{color:C.gold},line:{color:C.gold}});
  s.addText('CONFIDENTIAL — FOR LP DISTRIBUTION ONLY',{x:0,y:5.25,w:10,h:0.375,align:'center',valign:'middle',fontFace:'Arial',fontSize:9,bold:true,color:C.darkNavy,margin:0});
  s.addNotes(`本页最重要的三件事：
1. 英文版封面同样要强调，本周不是全面 risk-on，而是一次修复中的再分化。机构回流和仓位挤压让市场看起来很强，但 Kelp DAO 的大额攻击提醒投资人，底层风险还在，所以不应把这轮反弹简单理解成趋势翻转。
2. 对海外 LP 来说，最值得解释的是美国监管正在变得更“可操作”。DeFi front-end safe harbor 与 CLARITY Act 的推进，意味着过去纯靠执法裁量的灰区开始被重新定义，这会直接影响未来哪些赛道更容易被机构配置。
3. 结构性赢家已经在浮现。Hyperliquid、Plasma、Virtuals、Polymarket 这些项目共同反映出一个方向，就是资金正在重新奖励有真实用户、有现金流潜力、有产品闭环的平台。`);
}

{
  const s = pptx.addSlide(); s.background = { color: C.offWhite }; addHeader(s,'📊  Market Overview');
  const card=(x,l,v,d,c)=>{s.addShape(pptx.shapes.RECTANGLE,{x,y:0.95,w:2.8,h:1.05,fill:{color:C.white},line:{color:C.white},shadow:makeShadow()});s.addShape(pptx.shapes.RECTANGLE,{x,y:0.95,w:0.06,h:1.05,fill:{color:C.gold},line:{color:C.gold}});s.addText(l,{x:x+0.22,y:1.05,w:2.1,h:0.25,fontFace:'Arial',fontSize:12,bold:true,color:C.medGray,margin:0});s.addText(v,{x:x+0.22,y:1.3,w:2.2,h:0.35,fontFace:'Arial',fontSize:28,bold:true,color:C.bodyText,margin:0});s.addText(d,{x:x+0.22,y:1.68,w:2.2,h:0.2,fontFace:'Arial',fontSize:14,bold:true,color:c,margin:0});};
  card(0.6,'BTC','$75,480','+6.5%',C.green); card(3.7,'ETH','$2,312.20','+5.7%',C.green);
  s.addShape(pptx.shapes.RECTANGLE,{x:6.8,y:0.95,w:2.8,h:1.05,fill:{color:C.white},line:{color:C.white},shadow:makeShadow()}); s.addShape(pptx.shapes.RECTANGLE,{x:6.8,y:0.95,w:0.06,h:1.05,fill:{color:C.red},line:{color:C.red}});
  s.addText('Fear & Greed',{x:7.02,y:1.05,w:2.1,h:0.25,fontFace:'Arial',fontSize:12,bold:true,color:C.medGray,margin:0}); s.addText('27',{x:7.02,y:1.3,w:0.8,h:0.35,fontFace:'Arial',fontSize:28,bold:true,color:C.red,margin:0}); s.addText('Fear',{x:7.78,y:1.38,w:1.2,h:0.2,fontFace:'Arial',fontSize:13,color:C.red,margin:0});
  s.addTable([[{text:'Asset',options:{bold:true,color:C.white,fill:{color:C.navy},align:'center',fontFace:'Arial',fontSize:11}},{text:'Price',options:{bold:true,color:C.white,fill:{color:C.navy},align:'center',fontFace:'Arial',fontSize:11}},{text:'7d',options:{bold:true,color:C.white,fill:{color:C.navy},align:'center',fontFace:'Arial',fontSize:11}},{text:'Notes',options:{bold:true,color:C.white,fill:{color:C.navy},align:'center',fontFace:'Arial',fontSize:11}}],['SOL','$85.86','+5.2%','Recovered from the $81.5 support area'],['HYPE','$43.18','+5.8%','Perp DEX leader keeps consolidating share'],['ENA','$0.12','+28.9%','Best high-beta rebound'],['PENDLE','$1.30','+23.1%','Yield-theme recovery'],['DYDX','$0.14','+57.3%','Sharp low-base bounce, but share still weak']],{x:0.6,y:2.25,w:8.8,colW:[1.2,1.7,1.2,4.7],rowH:[0.35,0.33,0.33,0.33,0.33,0.33],border:{pt:0.4,color:C.lightGray},fontFace:'Arial',fontSize:10,color:C.bodyText});
  s.addText('Total market cap at $2.63T, BTC dominance 57.5%, 24h volume $109.0B. The market moved from panic toward cautious repair.',{x:0.6,y:4.55,w:8.8,h:0.35,fontFace:'Arial',fontSize:10,italic:true,color:C.medGray,margin:0}); addFooter(s,'Crypto Market Weekly | 2026.04.13–04.19');
  s.addNotes(`本页最重要的三件事：
1. 这页需要向英文受众解释，价格上涨背后的主要驱动力仍是仓位修复。连续负资金费率意味着前期市场做空非常拥挤，因此当宏观情绪略有缓和、资金重新流入时，价格会表现出超预期的弹性，但这并不等于基本面已经全面改善。
2. ETH 的资金流质量比表面涨幅更重要。ETH 产品净流入改善，加上市场对 staking ETF 最终期限的关注升温，使 ETH 重新具备“制度型催化”属性。对机构来说，这种资产通常比单纯 beta 资产更容易获得持续配置。
3. 在 portfolio lens 下，HYPE 仍是最有结构性逻辑的强者，ENA 和 PENDLE 则更像流动性与利率预期修复的高弹性表达。DYDX 的涨幅虽然大，但如果没有市场份额回升配合，它更像交易性反弹而不是基本面反转。`);
}

{
  const s = pptx.addSlide(); s.background = { color: C.offWhite }; addHeader(s,'🔥  Top Stories & Industry');
  s.addShape(pptx.shapes.RECTANGLE,{x:0.4,y:0.9,w:4.5,h:4.1,fill:{color:C.white},line:{color:C.white},shadow:makeShadow()}); s.addText('TOP 5 STORIES',{x:0.6,y:0.96,w:3.5,h:0.3,fontFace:'Arial',fontSize:13,bold:true,color:C.navy,margin:0});
  const stories=[['1','Kelp DAO $294M exploit','Shared bridge and LRT security risk exposed'],['2','SEC + CLARITY progress','Regulatory perimeter becomes more actionable'],['3','$2.2B institutional inflows','BTC and ETH lead allocation rebound'],['4','Hyperliquid at ~44% share','Perp DEX market concentrates around the leader'],['5','Virtuals ACP / Polymarket USD','Real usage rails continue to expand']];
  stories.forEach((it,i)=>{const y=1.38+i*0.72;s.addShape(pptx.shapes.OVAL,{x:0.64,y,w:0.3,h:0.3,fill:{color:C.gold},line:{color:C.gold}});s.addText(it[0],{x:0.64,y,w:0.3,h:0.3,fontFace:'Arial',fontSize:11,bold:true,color:C.white,align:'center',valign:'middle',margin:0});s.addText(it[1],{x:1.08,y:y-0.02,w:3.55,h:0.22,fontFace:'Arial',fontSize:10.5,bold:true,color:C.bodyText,margin:0});s.addText(it[2],{x:1.08,y:y+0.22,w:3.55,h:0.2,fontFace:'Arial',fontSize:9,color:C.medGray,margin:0});});
  const panel=(y,title,items)=>{s.addShape(pptx.shapes.RECTANGLE,{x:5.1,y,w:4.5,h:1.95,fill:{color:C.white},line:{color:C.white},shadow:makeShadow()});s.addShape(pptx.shapes.RECTANGLE,{x:5.1,y,w:4.5,h:0.35,fill:{color:C.navy},line:{color:C.navy}});s.addText(title,{x:5.3,y,w:3.9,h:0.35,fontFace:'Arial',fontSize:11,bold:true,color:C.white,valign:'middle',margin:0});s.addText(items.map((t,i)=>({text:t,options:{bullet:true,breakLine:i<items.length-1,fontFace:'Arial',fontSize:10,color:C.bodyText}})),{x:5.3,y:y+0.42,w:4.0,h:1.35,margin:0,paraSpaceAfterPt:5});};
  panel(0.9,'Layer 1 & Layer 2',['TON activates Catchain 2.0 with sub-second confirmations','Tempo launches Zones for permissioned private execution','Solana TVL mix shifts further toward lending yield','Ethereum Higota roadmap discussion stays focused on Verkle Trees']);
  panel(3.05,'DeFi & Perp DEXes',['SEC gives DeFi interfaces a 5-year interim safe harbor path','Plasma TVL tops $2B after Tether-linked rollout','Hyperliquid keeps taking share with ~ $6.5B daily volume','Sector volume cools, but leader concentration rises']);
  addFooter(s,'Crypto Market Weekly | 2026.04.13–04.19');
  s.addNotes(`本页最重要的三件事：
1. 对 LP 来说，这页最该看的不是某个 headline 本身，而是它们共同指向的产业结构变化。市场总量在去杠杆后没有回到高点，但头部平台仍在增强，这说明行业已经进入“低增速但高集中度”的阶段，赢家的护城河反而可能更深。
2. Hyperliquid 的表现尤其重要，因为它证明链上衍生品不是昙花一现的交易热度，而是在持续蚕食中心化交易所的份额。只要这个趋势继续，围绕链上撮合、保证金、流动性、做市和清算的整条价值链都会被重估。
3. TON、Tempo、Plasma 这些基础设施方向则说明，下一轮增长不会只由单一公链带动，而会来自不同类型的专用执行环境。消费级支付、机构隐私链、稳定币结算层、Agent commerce 这些场景可能分别孕育出新的赢家。`);
}

{
  const s = pptx.addSlide(); s.background = { color: C.offWhite }; addHeader(s,'⚖️  Regulation & Fundraising');
  s.addShape(pptx.shapes.RECTANGLE,{x:0.4,y:0.9,w:4.5,h:4.1,fill:{color:C.white},line:{color:C.white},shadow:makeShadow()}); s.addShape(pptx.shapes.RECTANGLE,{x:0.4,y:0.9,w:0.06,h:4.1,fill:{color:C.gold},line:{color:C.gold}});
  s.addText('Regulation & Policy',{x:0.7,y:0.96,w:3.5,h:0.3,fontFace:'Arial',fontSize:14,bold:true,color:C.navy,margin:0});
  [['SEC no-action for DeFi interfaces','Non-custodial and neutral front-ends benefit most'],['Senate advances CLARITY Act','BTC, ETH, SOL taxonomy becomes clearer'],['Tokenized securities framework evolves','Digital format does not remove securities status'],['Pakistan launches PVARA','A notable national shift from ban to regulated adoption'],['Stablecoin yield debate unresolved','Banks and crypto firms still at odds']].forEach((r,i)=>{const y=1.45+i*0.65;s.addText(r[0],{x:0.7,y,w:3.9,h:0.22,fontFace:'Arial',fontSize:11,bold:true,color:C.bodyText,margin:0});s.addText(r[1],{x:0.7,y:y+0.23,w:3.9,h:0.2,fontFace:'Arial',fontSize:9.5,color:C.medGray,margin:0});});
  s.addShape(pptx.shapes.RECTANGLE,{x:5.1,y:0.9,w:4.5,h:4.1,fill:{color:C.white},line:{color:C.white},shadow:makeShadow()}); s.addShape(pptx.shapes.RECTANGLE,{x:5.1,y:0.9,w:0.06,h:4.1,fill:{color:C.gold},line:{color:C.gold}});
  s.addText('Fundraising & M&A',{x:5.4,y:0.96,w:3.5,h:0.3,fontFace:'Arial',fontSize:14,bold:true,color:C.navy,margin:0});
  s.addTable([[{text:'Project',options:{bold:true,color:C.white,fill:{color:C.navy},fontFace:'Arial',fontSize:10}},{text:'Amount',options:{bold:true,color:C.white,fill:{color:C.navy},fontFace:'Arial',fontSize:10,align:'right'}},{text:'Notes',options:{bold:true,color:C.white,fill:{color:C.navy},fontFace:'Arial',fontSize:10}}],['Kraken','$200M','Deutsche Börse secondary investment'],['Spektr','$20M','AI-driven KYC and risk tooling'],['Paxos Labs','$12M','Amplify stablecoin toolkit'],['MicroStrategy','$1B','Fresh capital for BTC accumulation']],{x:5.3,y:1.45,w:4.1,colW:[1.2,0.9,2.0],rowH:[0.3,0.32,0.32,0.32,0.32],border:{pt:0.4,color:C.lightGray},fontFace:'Arial',fontSize:9.5,color:C.bodyText});
  s.addShape(pptx.shapes.RECTANGLE,{x:5.3,y:3.35,w:4.1,h:0.85,fill:{color:C.paleGold},line:{color:C.paleGold},shadow:makeShadow()}); s.addText('Capital remains focused on compliant trading access, stablecoin infrastructure, and AI-powered operating tools.',{x:5.52,y:3.48,w:3.6,h:0.55,fontFace:'Arial',fontSize:10,color:C.bodyText,margin:0,valign:'mid'});
  addFooter(s,'Crypto Market Weekly | 2026.04.13–04.19');
  s.addNotes(`本页最重要的三件事：
1. 对英文受众需要强调，当前监管变化的关键在于角色拆分。市场过去常把“crypto”当成一个整体被监管，但现在 SEC 和立法机构开始区分 front-end、custody、settlement、tokenized securities、stablecoin yield 等不同环节。谁的业务模式更清晰，谁就更容易拿到机构资本。
2. CLARITY Act 之所以重要，是因为它可能把临时指引变成更长期的制度基础。一旦数字商品和证券型代币的边界更稳定，传统金融机构就能更放心地推出 ETF、结构化产品、托管和交易服务，这会是长期增量资金进入的必要条件。
3. 融资与并购也说明资本现在更偏好“卖铲子”业务。Kraken 的二级交易、Paxos Labs 的企业工具、Spektr 的合规自动化，背后共同的逻辑是，市场正在为帮助机构上链的基础设施付更高估值，而不是单纯追逐高波动 beta。`);
}

{
  const s = pptx.addSlide(); s.background = { color: C.offWhite }; addHeader(s,'🛡️  Security & Key Sectors');
  s.addShape(pptx.shapes.RECTANGLE,{x:0.4,y:0.9,w:4.5,h:4.1,fill:{color:C.white},line:{color:C.white},shadow:makeShadow()}); s.addShape(pptx.shapes.RECTANGLE,{x:0.4,y:0.9,w:4.5,h:0.35,fill:{color:C.red},line:{color:C.red}}); s.addText('Security Incidents',{x:0.65,y:0.9,w:3.5,h:0.35,fontFace:'Arial',fontSize:12,bold:true,color:C.white,valign:'middle',margin:0});
  s.addTable([[{text:'Project',options:{bold:true,color:C.white,fill:{color:C.red},fontFace:'Arial',fontSize:10}},{text:'Loss',options:{bold:true,color:C.white,fill:{color:C.red},fontFace:'Arial',fontSize:10}},{text:'Comment',options:{bold:true,color:C.white,fill:{color:C.red},fontFace:'Arial',fontSize:10}}],['Kelp DAO','$294M','Forged rsETH bridge message exploit'],['Grinex','$13.7M','Exchange breach and service suspension'],['Rhea Finance','$18.4M','Slippage logic vulnerability'],['Kraken','Internal incident','Approx. 2,000 accounts viewed']],{x:0.6,y:1.35,w:4.05,colW:[0.95,1.0,2.1],rowH:[0.28,0.34,0.34,0.34,0.34],border:{pt:0.4,color:C.lightGray},fontFace:'Arial',fontSize:9.2,color:C.bodyText});
  s.addText('Stablecoins / AI / Prediction',{x:0.6,y:3.3,w:3.4,h:0.25,fontFace:'Arial',fontSize:12,bold:true,color:C.navy,margin:0});
  s.addText([{text:'• ',options:{color:C.gold}},{text:'Yield-bearing stablecoins are growing roughly 15x faster than the market',options:{color:C.bodyText}},{text:'\n• ',options:{color:C.gold}},{text:'Virtuals ACP expands agent-held capital and cross-chain execution',options:{color:C.bodyText}},{text:'\n• ',options:{color:C.gold}},{text:'Polymarket USD + CTF V2 improve matching efficiency and fee capture',options:{color:C.bodyText}}],{x:0.7,y:3.62,w:3.9,h:1.0,fontFace:'Arial',fontSize:10,margin:0});
  const box=(y,title,items)=>{s.addShape(pptx.shapes.RECTANGLE,{x:5.1,y,w:4.5,h:1.95,fill:{color:C.white},line:{color:C.white},shadow:makeShadow()});s.addShape(pptx.shapes.RECTANGLE,{x:5.1,y,w:4.5,h:0.35,fill:{color:C.navy},line:{color:C.navy}});s.addText(title,{x:5.3,y,w:3.9,h:0.35,fontFace:'Arial',fontSize:11,bold:true,color:C.white,valign:'middle',margin:0});s.addText(items.map((t,i)=>({text:t,options:{bullet:true,breakLine:i<items.length-1,fontFace:'Arial',fontSize:10,color:C.bodyText}})),{x:5.3,y:y+0.42,w:4.0,h:1.35,margin:0,paraSpaceAfterPt:5});};
  box(0.9,'AI × Crypto',['Messari sees AI agents becoming major on-chain users','ETHGas and ether.fi announce a $3B-scale partnership','Utility Net Mainnet V2 revives DePIN compute narrative','Ledger pushes human-in-the-loop agent security']);
  box(3.05,'Prediction Markets',['Election-related volume remains above $2B cumulative','Oil and geopolitics markets absorb major new flow','BTC-linked contracts remain the primary liquidity magnet','Regulatory scrutiny is rising with mainstream adoption']);
  addFooter(s,'Crypto Market Weekly | 2026.04.13–04.19');
  s.addNotes(`本页最重要的三件事：
1. 对外部投资人来说，Kelp DAO 事件的真正含义在于“模块化安全风险”的再定价。很多协议现在并不是单点故障，而是把桥、预言机、再质押、外部消息验证耦合在一起，一旦其中一个环节失效，损失会快速传导到多个协议和市场。
2. 稳定币、AI Agent、预测市场这三条线值得放在一起看，因为它们共同描绘的是 crypto 如何承接真实经济活动。稳定币解决支付和结算，AI Agent 解决自动化执行，预测市场解决概率和信息定价，这三层一旦打通，会比传统单点协议拥有更强的规模效应。
3. 预测市场的 mainstream 化已经开始产生第二层影响。随着它越来越被媒体、终端和研究平台引用，监管自然会跟上。这既带来合规压力，也可能反过来提升头部平台的壁垒，因为只有少数玩家能承受长期合规成本。`);
}

{
  const s = pptx.addSlide(); s.background = { color: C.offWhite }; addHeader(s,'🔮  Catalysts & Outlook');
  s.addTable([[{text:'Date',options:{bold:true,color:C.white,fill:{color:C.navy},fontFace:'Arial',fontSize:10}},{text:'Event',options:{bold:true,color:C.white,fill:{color:C.navy},fontFace:'Arial',fontSize:10}},{text:'Assets',options:{bold:true,color:C.white,fill:{color:C.navy},fontFace:'Arial',fontSize:10}},{text:'Expected Impact',options:{bold:true,color:C.white,fill:{color:C.navy},fontFace:'Arial',fontSize:10}}],['Late Apr','ETHA staking ETF final deadline','ETH / LDO / ETHFI','Repricing of staking-yield assets'],['Late Apr','CLARITY Act committee process','Broad market / stablecoins / DeFi','Regulatory repricing'],['Apr 27–29','Bitcoin 2026','BTC / Bitcoin L2s','Institutional and BTCFi catalyst'],['Apr 29–30','TOKEN2049 Dubai','Broad alt / VC ecosystem','Partnership and fundraising catalyst'],['Apr 30','KMNO unlock','KMNO / Solana DeFi','Moderate supply pressure'],['May 4','CME SUI futures','SUI','Institutional derivatives access']],{x:0.45,y:1.0,w:6.2,colW:[1.05,2.2,1.35,1.6],rowH:[0.32,0.34,0.3,0.3,0.3,0.3,0.32],border:{pt:0.4,color:C.lightGray},fontFace:'Arial',fontSize:9.3,color:C.bodyText});
  s.addShape(pptx.shapes.RECTANGLE,{x:6.95,y:1.0,w:2.6,h:3.9,fill:{color:C.white},line:{color:C.white},shadow:makeShadow()}); s.addShape(pptx.shapes.RECTANGLE,{x:6.95,y:1.0,w:2.6,h:0.35,fill:{color:C.navy},line:{color:C.navy}}); s.addText('Weekly Takeaway',{x:7.1,y:1.0,w:2.1,h:0.35,fontFace:'Arial',fontSize:11,bold:true,color:C.white,valign:'middle',margin:0});
  s.addText([{text:'• ',options:{color:C.gold}},{text:'The market repaired positioning, but remains fragile',options:{color:C.bodyText}},{text:'\n• ',options:{color:C.gold}},{text:'Regulation and security will shape the next leg of risk appetite',options:{color:C.bodyText}},{text:'\n• ',options:{color:C.gold}},{text:'HYPE remains structurally strongest in the portfolio lens',options:{color:C.bodyText}},{text:'\n• ',options:{color:C.gold}},{text:'ENA and PENDLE stay more sensitive to liquidity recovery',options:{color:C.bodyText}}],{x:7.18,y:1.48,w:2.0,h:1.9,fontFace:'Arial',fontSize:10,margin:0});
  s.addShape(pptx.shapes.RECTANGLE,{x:7.15,y:3.55,w:2.05,h:0.9,fill:{color:C.paleGold},line:{color:C.paleGold}}); s.addText('Portfolio lens:\nWatch HYPE for structural strength, and ENA / PENDLE for liquidity-driven upside beta.',{x:7.28,y:3.72,w:1.8,h:0.5,fontFace:'Arial',fontSize:9.5,bold:true,color:C.navy,margin:0,align:'center'});
  addFooter(s,'Crypto Market Weekly | 2026.04.13–04.19');
  s.addNotes(`本页最重要的三件事：
1. 接下来两周的市场方向，很大程度取决于制度催化能否兑现。ETH staking ETF 的最终期限和 CLARITY Act 的推进，都是能改变长期资金配置逻辑的事件。如果这两件事有正面进展，市场将更愿意重新给 ETH、质押赛道、稳定币和 DeFi infrastructure 提高估值。
2. 同时不能忽视供给和事件风险。像 KMNO 解锁这类供给事件，在成交量尚未完全恢复的环境中仍可能显著放大波动。也就是说，中期方向和短期节奏依然可能分离，仓位管理会比单纯方向判断更重要。
3. 组合角度，HYPE 依然最有结构性确定性，因为它背后是链上衍生品持续抢占市场份额的长期趋势。ENA 和 PENDLE 更像在流动性回暖、稳定币收益扩张、风险偏好改善时会表现更好的高弹性工具。`);
}

pptx.writeFile('/Users/lx/Github/crypto-weekly-digest/reports/2026-04-13/slides-en.pptx');