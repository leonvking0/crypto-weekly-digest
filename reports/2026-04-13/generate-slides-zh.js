const pptxgen = require('pptxgenjs');

const C = {
  navy: '1A2744', darkNavy: '0F1A2E', gold: 'C9A96E', white: 'FFFFFF', offWhite: 'F5F6F8',
  lightGray: 'E8EBF0', medGray: '8C95A6', bodyText: '2D3748', red: 'C53030', green: '2F855A', paleGold: 'FFF8E7'
};
const pptx = new pptxgen();
pptx.layout = 'LAYOUT_16x9'; pptx.author = '42'; pptx.company = '42'; pptx.subject = 'Crypto Weekly Digest'; pptx.title = '加密市场周报 2026.04.13–04.19'; pptx.lang = 'zh-CN';
const makeShadow = () => ({ type: 'outer', blur: 3, offset: 1, angle: 45, color: '000000', opacity: 0.12 });
const addFooter = (s, t) => { s.addShape(pptx.shapes.RECTANGLE,{x:0,y:5.25,w:10,h:0.375,fill:{color:C.navy},line:{color:C.navy}}); s.addText(t,{x:0.45,y:5.25,w:9.1,h:0.375,fontFace:'Arial',fontSize:8,color:C.medGray,valign:'middle',margin:0}); };
const addHeader = (s, t) => { s.addShape(pptx.shapes.RECTANGLE,{x:0,y:0,w:10,h:0.7,fill:{color:C.navy},line:{color:C.navy}}); s.addText(t,{x:0.6,y:0,w:8.8,h:0.7,fontFace:'Arial',fontSize:20,bold:true,color:C.white,valign:'middle',margin:0}); };

{
  const s = pptx.addSlide(); s.background = { color: C.darkNavy };
  s.addShape(pptx.shapes.RECTANGLE,{x:0.8,y:1.55,w:1.2,h:0.04,fill:{color:C.gold},line:{color:C.gold}});
  s.addText('加密市场周报',{x:0.8,y:1.8,w:5.5,h:0.8,fontFace:'Arial',fontSize:42,bold:true,color:C.white,margin:0});
  s.addText('2026年4月13日 — 4月19日',{x:0.8,y:2.7,w:5.8,h:0.45,fontFace:'Arial',fontSize:20,color:C.gold,margin:0});
  s.addText('数据来源：Messari · The Block · CoinGecko · DefiLlama',{x:0.8,y:3.4,w:6.8,h:0.3,fontFace:'Arial',fontSize:11,color:C.medGray,margin:0});
  s.addText('本周关键词：Kelp 攻击、机构净流入、CLARITY Act、Hyperliquid 集中、Agent Commerce',{x:0.8,y:4.15,w:8.5,h:0.55,fontFace:'Arial',fontSize:14,color:C.white,margin:0});
  s.addShape(pptx.shapes.RECTANGLE,{x:0,y:5.25,w:10,h:0.375,fill:{color:C.gold},line:{color:C.gold}});
  s.addText('CONFIDENTIAL — FOR LP DISTRIBUTION ONLY',{x:0,y:5.25,w:10,h:0.375,align:'center',valign:'middle',fontFace:'Arial',fontSize:9,bold:true,color:C.darkNavy,margin:0});
  s.addNotes(`本页最重要的三件事：
1. 本周封面要传达的核心是，市场主线从“恐慌”转成了“修复中的分化”。BTC 和 ETH 在机构资金流入与空头回补推动下显著反弹，但周末 Kelp DAO 大额攻击提醒大家，行业底层风险并没有消失，所以这仍不是全面风险偏好扩张的阶段。
2. 第二条主线是监管终于从口头讨论进入可执行框架。SEC 对 DeFi 前端和非托管钱包给出阶段性豁免，参议院又推进 CLARITY Act，意味着资本市场开始看到一个更清晰的合规扩张路径，这对稳定币、DeFi、交易基础设施的估值都很关键。
3. 第三条主线是行业结构继续集中。Hyperliquid、Plasma、Virtuals、Polymarket 这类有明确产品闭环的平台继续走强，说明资金越来越愿意为真实使用场景而不是纯叙事付费。`);
}

{
  const s = pptx.addSlide(); s.background = { color: C.offWhite }; addHeader(s,'📊  市场概览');
  const card=(x,l,v,d,c)=>{s.addShape(pptx.shapes.RECTANGLE,{x,y:0.95,w:2.8,h:1.05,fill:{color:C.white},line:{color:C.white},shadow:makeShadow()});s.addShape(pptx.shapes.RECTANGLE,{x,y:0.95,w:0.06,h:1.05,fill:{color:C.gold},line:{color:C.gold}});s.addText(l,{x:x+0.22,y:1.05,w:2.1,h:0.25,fontFace:'Arial',fontSize:12,bold:true,color:C.medGray,margin:0});s.addText(v,{x:x+0.22,y:1.3,w:2.2,h:0.35,fontFace:'Arial',fontSize:28,bold:true,color:C.bodyText,margin:0});s.addText(d,{x:x+0.22,y:1.68,w:2.2,h:0.2,fontFace:'Arial',fontSize:14,bold:true,color:c,margin:0});};
  card(0.6,'BTC','$75,480','+6.5%',C.green); card(3.7,'ETH','$2,312.20','+5.7%',C.green);
  s.addShape(pptx.shapes.RECTANGLE,{x:6.8,y:0.95,w:2.8,h:1.05,fill:{color:C.white},line:{color:C.white},shadow:makeShadow()});
  s.addShape(pptx.shapes.RECTANGLE,{x:6.8,y:0.95,w:0.06,h:1.05,fill:{color:C.red},line:{color:C.red}});
  s.addText('恐惧贪婪指数',{x:7.02,y:1.05,w:2.1,h:0.25,fontFace:'Arial',fontSize:12,bold:true,color:C.medGray,margin:0});
  s.addText('27',{x:7.02,y:1.3,w:0.8,h:0.35,fontFace:'Arial',fontSize:28,bold:true,color:C.red,margin:0});
  s.addText('Fear',{x:7.78,y:1.38,w:1.2,h:0.2,fontFace:'Arial',fontSize:13,color:C.red,margin:0});
  s.addTable([
    [{text:'资产',options:{bold:true,color:C.white,fill:{color:C.navy},align:'center',fontFace:'Arial',fontSize:11}},{text:'价格',options:{bold:true,color:C.white,fill:{color:C.navy},align:'center',fontFace:'Arial',fontSize:11}},{text:'7d',options:{bold:true,color:C.white,fill:{color:C.navy},align:'center',fontFace:'Arial',fontSize:11}},{text:'备注',options:{bold:true,color:C.white,fill:{color:C.navy},align:'center',fontFace:'Arial',fontSize:11}}],
    ['SOL','$85.86','+5.2%','自 $81.5 支撑位企稳反弹'],['HYPE','$43.18','+5.8%','永续 DEX 龙头继续集中'],['ENA','$0.12','+28.9%','高 beta 反弹最强'],['PENDLE','$1.30','+23.1%','收益赛道修复'],['DYDX','$0.14','+57.3%','低基数反弹，但份额仍弱']
  ],{x:0.6,y:2.25,w:8.8,colW:[1.2,1.7,1.2,4.7],rowH:[0.35,0.33,0.33,0.33,0.33,0.33],border:{pt:0.4,color:C.lightGray},fontFace:'Arial',fontSize:10,color:C.bodyText});
  s.addText('总市值 $2.63T，BTC 主导率 57.5%，24h 成交额 $109.0B。市场从极端恐慌回到谨慎修复。',{x:0.6,y:4.55,w:8.8,h:0.35,fontFace:'Arial',fontSize:10,italic:true,color:C.medGray,margin:0}); addFooter(s,'加密市场周报 | 2026.04.13–04.19');
  s.addNotes(`本页最重要的三件事：
1. 这轮上涨的性质更像 short squeeze + 机构回补，而不是新的单边趋势起点。BTC 连续 46 天负资金费率后重新上冲 $76K，本质上说明市场在此前过度悲观，资金先修正的是仓位结构，而不一定是长期增长预期。
2. ETH 虽然周涨幅略低于 BTC，但资金流结构更值得重视。单周近 5 亿美元的 ETH 产品净流入和链上活跃度改善，说明资金正在为“质押、ETF、合规收益化”这条主线提前定价。如果后续 ETH 质押型 ETF 真有进展，ETH 的估值弹性会继续提升。
3. 组合相关资产里，ENA 和 PENDLE 的弹性明显更大，说明市场在回暖时最先追逐的是与稳定币收益、链上利率和资金效率有关的资产。DYDX 反弹虽然剧烈，但更多是低基数修复，基本面并未显示它重新夺回赛道主导权。`);
}

{
  const s = pptx.addSlide(); s.background = { color: C.offWhite }; addHeader(s,'🔥  本周要闻 & 行业动态');
  s.addShape(pptx.shapes.RECTANGLE,{x:0.4,y:0.9,w:4.5,h:4.1,fill:{color:C.white},line:{color:C.white},shadow:makeShadow()});
  s.addText('本周要闻 TOP 5',{x:0.6,y:0.96,w:3.5,h:0.3,fontFace:'Arial',fontSize:13,bold:true,color:C.navy,margin:0});
  const stories=[['1','Kelp DAO $294M 攻击','LRT 与跨链共享安全风险集中暴露'],['2','SEC + CLARITY 同步推进','DeFi 与数字商品监管边界逐渐清晰'],['3','机构净流入 $2.2B','BTC 与 ETH 成为主要受益者'],['4','Hyperliquid 市占率约 44%','永续 DEX 进入龙头集聚阶段'],['5','Virtuals ACP / Polymarket USD','真实使用场景继续扩张']];
  stories.forEach((it,i)=>{const y=1.38+i*0.72;s.addShape(pptx.shapes.OVAL,{x:0.64,y,w:0.3,h:0.3,fill:{color:C.gold},line:{color:C.gold}});s.addText(it[0],{x:0.64,y,w:0.3,h:0.3,fontFace:'Arial',fontSize:11,bold:true,color:C.white,align:'center',valign:'middle',margin:0});s.addText(it[1],{x:1.08,y:y-0.02,w:3.55,h:0.22,fontFace:'Arial',fontSize:10.5,bold:true,color:C.bodyText,margin:0});s.addText(it[2],{x:1.08,y:y+0.22,w:3.55,h:0.2,fontFace:'Arial',fontSize:9,color:C.medGray,margin:0});});
  const panel=(y,title,items)=>{s.addShape(pptx.shapes.RECTANGLE,{x:5.1,y,w:4.5,h:1.95,fill:{color:C.white},line:{color:C.white},shadow:makeShadow()});s.addShape(pptx.shapes.RECTANGLE,{x:5.1,y,w:4.5,h:0.35,fill:{color:C.navy},line:{color:C.navy}});s.addText(title,{x:5.3,y,w:3.9,h:0.35,fontFace:'Arial',fontSize:11,bold:true,color:C.white,valign:'middle',margin:0});s.addText(items.map((t,i)=>({text:t,options:{bullet:true,breakLine:i<items.length-1,fontFace:'Arial',fontSize:10,color:C.bodyText}})),{x:5.3,y:y+0.42,w:4.0,h:1.35,margin:0,paraSpaceAfterPt:5});};
  panel(0.9,'Layer 1 & Layer 2',['TON Catchain 2.0 激活，确认时间压缩到亚秒级','Tempo 推出 Zones，切入可许可隐私执行链','Solana TVL 更偏向借贷收益，DEX 占比下滑','以太坊 Higota 路线继续围绕 Verkle Trees']);
  panel(3.05,'DeFi & 永续合约',['SEC 对 DeFi 前端给出 5 年临时 safe harbor','Plasma TVL 升破 $2B，Tether 生态继续外扩','Hyperliquid 24h 成交额约 $6.5B，持续吃份额','Perp DEX 总量降温，但头部更强']);
  addFooter(s,'加密市场周报 | 2026.04.13–04.19');
  s.addNotes(`本页最重要的三件事：
1. 这页最重要的结论是“行业在去泡沫，但没有停创新”。Perp DEX 总量相比高峰回落，说明投机杠杆在降温，但 Hyperliquid 仍在吃份额，证明用户不会因为周期降温就停止使用好产品，反而会更加集中到体验最好、流动性最强的平台。
2. L1/L2 竞争焦点也正在变化。TON 和 Tempo 代表的是两种不同方向，一个追求更快的消费级和支付型体验，一个追求机构需要的隐私执行环境。市场不再只比 TPS，而是比谁能真正承接稳定币、Agent、机构资产这些真实工作负载。
3. Plasma、Hyperliquid、Virtuals 这些项目共同指向一个趋势，行业增量越来越来自“可执行的应用层闭环”。也就是资金、结算、交易、自动执行能在一条链路里完成。这会让没有真实使用需求支撑的资产越来越难拿到溢价。`);
}

{
  const s = pptx.addSlide(); s.background = { color: C.offWhite }; addHeader(s,'⚖️  监管政策 & 投融资');
  s.addShape(pptx.shapes.RECTANGLE,{x:0.4,y:0.9,w:4.5,h:4.1,fill:{color:C.white},line:{color:C.white},shadow:makeShadow()}); s.addShape(pptx.shapes.RECTANGLE,{x:0.4,y:0.9,w:0.06,h:4.1,fill:{color:C.gold},line:{color:C.gold}});
  s.addText('监管与政策',{x:0.7,y:0.96,w:3.5,h:0.3,fontFace:'Arial',fontSize:14,bold:true,color:C.navy,margin:0});
  [['SEC 为 DeFi 前端提供临时 no-action','非托管、不中介客户资产的平台最受益'],['参议院推进 CLARITY Act','BTC、ETH、SOL 归类更清晰'],['Tokenized securities 路径继续澄清','链上证券仍是证券，但小额融资豁免在成形'],['Pakistan 设立 PVARA','国家层面由禁转管'],['稳定币收益争议仍未解决','银行与加密平台博弈加剧']].forEach((r,i)=>{const y=1.45+i*0.65;s.addText(r[0],{x:0.7,y,w:3.9,h:0.22,fontFace:'Arial',fontSize:11,bold:true,color:C.bodyText,margin:0});s.addText(r[1],{x:0.7,y:y+0.23,w:3.9,h:0.2,fontFace:'Arial',fontSize:9.5,color:C.medGray,margin:0});});
  s.addShape(pptx.shapes.RECTANGLE,{x:5.1,y:0.9,w:4.5,h:4.1,fill:{color:C.white},line:{color:C.white},shadow:makeShadow()}); s.addShape(pptx.shapes.RECTANGLE,{x:5.1,y:0.9,w:0.06,h:4.1,fill:{color:C.gold},line:{color:C.gold}});
  s.addText('投融资与并购',{x:5.4,y:0.96,w:3.5,h:0.3,fontFace:'Arial',fontSize:14,bold:true,color:C.navy,margin:0});
  s.addTable([[{text:'项目',options:{bold:true,color:C.white,fill:{color:C.navy},fontFace:'Arial',fontSize:10}},{text:'金额',options:{bold:true,color:C.white,fill:{color:C.navy},fontFace:'Arial',fontSize:10,align:'right'}},{text:'备注',options:{bold:true,color:C.white,fill:{color:C.navy},fontFace:'Arial',fontSize:10}}],['Kraken','$200M','Deutsche Börse 入股，估值约 $13.3B'],['Spektr','$20M','AI KYC / 风控基础设施'],['Paxos Labs','$12M','Amplify 稳定币工具链'],['MicroStrategy','$1B','继续加杠杆买 BTC']],{x:5.3,y:1.45,w:4.1,colW:[1.2,0.9,2.0],rowH:[0.3,0.32,0.32,0.32,0.32],border:{pt:0.4,color:C.lightGray},fontFace:'Arial',fontSize:9.5,color:C.bodyText});
  s.addShape(pptx.shapes.RECTANGLE,{x:5.3,y:3.35,w:4.1,h:0.85,fill:{color:C.paleGold},line:{color:C.paleGold},shadow:makeShadow()});
  s.addText('一级市场和并购继续押注三条主线：合规交易入口、稳定币基础设施、AI 驱动的风控与运营工具。',{x:5.52,y:3.48,w:3.6,h:0.55,fontFace:'Arial',fontSize:10,color:C.bodyText,margin:0,valign:'mid'});
  addFooter(s,'加密市场周报 | 2026.04.13–04.19');
  s.addNotes(`本页最重要的三件事：
1. 监管现在的变化不只是“态度更友好”，而是开始把不同角色拆分对待。SEC 给 DeFi 前端的 no-action 口径，实际上是在告诉市场，只要不托管客户资产、不提供投资建议、只做中立技术界面，就有机会在现有框架下生存。这会让真正的非托管基础设施明显受益。
2. CLARITY Act 的意义在于把过去依赖临时解释和执法裁量的边界，变成更可预期的法律分类。一旦 BTC、ETH、SOL 这类资产的监管属性被更明确地锁定，机构就更容易做长期产品设计，稳定币、交易接口、托管和衍生品都可能获得估值重估。
3. 融资和并购也在验证这一点。Kraken 引入传统交易所资本、Paxos Labs 继续扩张稳定币工具链、Spektr 获得 AI 风控融资，说明资本现在买的是“合规放量能力”和“真实现金流基础设施”，而不是上一轮那种纯故事型增长。`);
}

{
  const s = pptx.addSlide(); s.background = { color: C.offWhite }; addHeader(s,'🛡️  安全事件 & 重点赛道');
  s.addShape(pptx.shapes.RECTANGLE,{x:0.4,y:0.9,w:4.5,h:4.1,fill:{color:C.white},line:{color:C.white},shadow:makeShadow()}); s.addShape(pptx.shapes.RECTANGLE,{x:0.4,y:0.9,w:4.5,h:0.35,fill:{color:C.red},line:{color:C.red}});
  s.addText('安全事件',{x:0.65,y:0.9,w:3.5,h:0.35,fontFace:'Arial',fontSize:12,bold:true,color:C.white,valign:'middle',margin:0});
  s.addTable([[{text:'项目',options:{bold:true,color:C.white,fill:{color:C.red},fontFace:'Arial',fontSize:10}},{text:'损失',options:{bold:true,color:C.white,fill:{color:C.red},fontFace:'Arial',fontSize:10}},{text:'说明',options:{bold:true,color:C.white,fill:{color:C.red},fontFace:'Arial',fontSize:10}}],['Kelp DAO','$294M','rsETH 跨链桥伪造消息攻击'],['Grinex','$13.7M','交易所被盗后暂停服务'],['Rhea Finance','$18.4M','滑点逻辑漏洞'],['Kraken','内部事件','约 2,000 账户信息被查看']],{x:0.6,y:1.35,w:4.05,colW:[0.95,1.0,2.1],rowH:[0.28,0.34,0.34,0.34,0.34],border:{pt:0.4,color:C.lightGray},fontFace:'Arial',fontSize:9.2,color:C.bodyText});
  s.addText('稳定币 / AI / 预测市场',{x:0.6,y:3.3,w:3.2,h:0.25,fontFace:'Arial',fontSize:12,bold:true,color:C.navy,margin:0});
  s.addText([{text:'• ',options:{color:C.gold}},{text:'收益型稳定币近半年增速约为大盘 15 倍',options:{color:C.bodyText}},{text:'\n• ',options:{color:C.gold}},{text:'Virtuals ACP 推进 Agent 跨链持资与执行',options:{color:C.bodyText}},{text:'\n• ',options:{color:C.gold}},{text:'Polymarket USD + CTF V2 改善手续费与撮合效率',options:{color:C.bodyText}}],{x:0.7,y:3.62,w:3.9,h:1.0,fontFace:'Arial',fontSize:10,margin:0});
  const box=(y,title,items)=>{s.addShape(pptx.shapes.RECTANGLE,{x:5.1,y,w:4.5,h:1.95,fill:{color:C.white},line:{color:C.white},shadow:makeShadow()});s.addShape(pptx.shapes.RECTANGLE,{x:5.1,y,w:4.5,h:0.35,fill:{color:C.navy},line:{color:C.navy}});s.addText(title,{x:5.3,y,w:3.9,h:0.35,fontFace:'Arial',fontSize:11,bold:true,color:C.white,valign:'middle',margin:0});s.addText(items.map((t,i)=>({text:t,options:{bullet:true,breakLine:i<items.length-1,fontFace:'Arial',fontSize:10,color:C.bodyText}})),{x:5.3,y:y+0.42,w:4.0,h:1.35,margin:0,paraSpaceAfterPt:5});};
  box(0.9,'AI × Crypto',['Messari 判断 AI Agent 将成主要链上用户','ETHGas 与 ether.fi 达成 $3B 级合作','Utility Net Mainnet V2 推动 DePIN 算力叙事','Ledger 强化 human-in-the-loop Agent 安全']);
  box(3.05,'预测市场',['2026 选举相关成交累计超 $2B','油价和地缘政治合约热度显著抬升','BTC 相关市场成为最集中流量入口','监管审视开始跟上产品扩张']);
  addFooter(s,'加密市场周报 | 2026.04.13–04.19');
  s.addNotes(`本页最重要的三件事：
1. Kelp DAO 攻击是本周最关键的风险事件，因为它提醒市场，DeFi 现在最大的脆弱点往往不在单一智能合约，而在跨链消息、共享安全模型和外部基础设施的耦合。随着协议越来越模块化，安全风险也会越来越系统化，这会显著提高市场对桥、LRT、再质押赛道的风险贴现率。
2. 稳定币、AI Agent 和预测市场表面上是三个不同赛道，本质上其实都在争夺“链上真实经济活动的底层入口”。稳定币解决结算，Agent 解决自动执行，预测市场解决实时概率定价。如果这三层继续打通，会形成一个比单一 DeFi 协议更大的网络效应。
3. 预测市场的变化尤其值得重视。它正在从投机工具转向信息基础设施，用户不是只拿它下注，而是把它当成对事件概率、宏观情绪和价格预期的实时信号层。未来它和 AI Agent、自动交易、稳定币支付结合后，可能成为新的链上金融操作界面。`);
}

{
  const s = pptx.addSlide(); s.background = { color: C.offWhite }; addHeader(s,'🔮  催化剂 & 周度展望');
  s.addTable([[{text:'日期',options:{bold:true,color:C.white,fill:{color:C.navy},fontFace:'Arial',fontSize:10}},{text:'事件',options:{bold:true,color:C.white,fill:{color:C.navy},fontFace:'Arial',fontSize:10}},{text:'影响资产',options:{bold:true,color:C.white,fill:{color:C.navy},fontFace:'Arial',fontSize:10}},{text:'预期影响',options:{bold:true,color:C.white,fill:{color:C.navy},fontFace:'Arial',fontSize:10}}],['4月下旬','ETHA Staking ETF 最终截止期','ETH / LDO / ETHFI','质押收益资产重估'],['4月下旬','CLARITY Act 审议','全市场 / 稳定币 / DeFi','监管预期重定价'],['4/27–4/29','Bitcoin 2026','BTC / BTC L2','机构与 BTCFi 催化'],['4/29–4/30','TOKEN2049 Dubai','广义山寨 / VC','合作与融资催化'],['4/30','KMNO 解锁','KMNO / Solana','中度供给压力'],['5/4','CME 上线 SUI Futures','SUI','机构衍生品入口']],{x:0.45,y:1.0,w:6.2,colW:[1.05,2.2,1.35,1.6],rowH:[0.32,0.34,0.3,0.3,0.3,0.3,0.32],border:{pt:0.4,color:C.lightGray},fontFace:'Arial',fontSize:9.3,color:C.bodyText});
  s.addShape(pptx.shapes.RECTANGLE,{x:6.95,y:1.0,w:2.6,h:3.9,fill:{color:C.white},line:{color:C.white},shadow:makeShadow()}); s.addShape(pptx.shapes.RECTANGLE,{x:6.95,y:1.0,w:2.6,h:0.35,fill:{color:C.navy},line:{color:C.navy}});
  s.addText('周度总结',{x:7.15,y:1.0,w:2.1,h:0.35,fontFace:'Arial',fontSize:11,bold:true,color:C.white,valign:'middle',margin:0});
  s.addText([{text:'• ',options:{color:C.gold}},{text:'市场完成一轮仓位修复，但并未脱离脆弱区间',options:{color:C.bodyText}},{text:'\n• ',options:{color:C.gold}},{text:'监管与安全将决定下一阶段风险偏好',options:{color:C.bodyText}},{text:'\n• ',options:{color:C.gold}},{text:'组合相关方向里，HYPE 仍最强，ENA/PENDLE 对流动性更敏感',options:{color:C.bodyText}},{text:'\n• ',options:{color:C.gold}},{text:'若 ETH 质押 ETF 与 CLARITY 同时推进，ETH beta 可能重新领先',options:{color:C.bodyText}}],{x:7.18,y:1.48,w:2.0,h:1.9,fontFace:'Arial',fontSize:10,margin:0});
  s.addShape(pptx.shapes.RECTANGLE,{x:7.15,y:3.55,w:2.05,h:0.9,fill:{color:C.paleGold},line:{color:C.paleGold}});
  s.addText('组合观察：\n关注 HYPE 的结构性强势，以及 ENA / PENDLE 在流动性继续回暖时的弹性。',{x:7.3,y:3.72,w:1.75,h:0.5,fontFace:'Arial',fontSize:9.5,bold:true,color:C.navy,margin:0,align:'center'});
  addFooter(s,'加密市场周报 | 2026.04.13–04.19');
  s.addNotes(`本页最重要的三件事：
1. 接下来一到两周的关键不是单纯看价格，而是看“监管利好能否落地”与“安全冲击是否继续外溢”这两个变量谁占上风。如果 CLARITY Act 和 ETH 质押 ETF 出现实质进展，市场会优先重估 ETH、稳定币、DeFi 和链上基础设施；如果安全事件继续扩散，资金又会重新回到更保守的位置。
2. 会议和产品催化也值得重视。Bitcoin 2026 和 TOKEN2049 这类大会不只是情绪场，它们往往是机构合作、融资公告、BTCFi 新叙事落地的发布窗口。尤其在监管转暖背景下，会议更容易放大原本就在形成的趋势。
3. 组合角度，HYPE 仍然是结构性最强的资产，因为它代表的是链上衍生品持续蚕食中心化份额这一长期趋势。ENA 和 PENDLE 则更依赖风险偏好与利率预期回暖，属于更高弹性但也更受市场情绪影响的仓位。`);
}

pptx.writeFile('/Users/lx/Github/crypto-weekly-digest/reports/2026-04-13/slides-zh.pptx');