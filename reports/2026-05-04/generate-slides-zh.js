const pptxgen = require('pptxgenjs');
const C = { navy:'1A2744', darkNavy:'0F1A2E', gold:'C9A96E', white:'FFFFFF', offWhite:'F5F6F8', lightGray:'E8EBF0', medGray:'8C95A6', bodyText:'2D3748', red:'C53030', green:'2F855A', paleGold:'FFF8E7' };
const pptx = new pptxgen();
pptx.layout = 'LAYOUT_16x9'; pptx.author = '42'; pptx.company='42'; pptx.subject='Crypto Weekly Digest'; pptx.title='加密市场周报 2026.05.04–05.10'; pptx.lang='zh-CN';
const makeShadow = () => ({ type:'outer', blur:3, offset:1, angle:45, color:'000000', opacity:0.12 });
const addFooter = (s,t)=>{ s.addShape(pptx.shapes.RECTANGLE,{x:0,y:5.25,w:10,h:0.375,fill:{color:C.navy},line:{color:C.navy}}); s.addText(t,{x:0.45,y:5.25,w:9.1,h:0.375,fontFace:'Arial',fontSize:8,color:C.medGray,valign:'middle',margin:0}); };
const addHeader = (s,t)=>{ s.addShape(pptx.shapes.RECTANGLE,{x:0,y:0,w:10,h:0.7,fill:{color:C.navy},line:{color:C.navy}}); s.addText(t,{x:0.6,y:0,w:8.8,h:0.7,fontFace:'Arial',fontSize:20,bold:true,color:C.white,valign:'middle',margin:0}); };

{
 const s=pptx.addSlide(); s.background={color:C.darkNavy};
 s.addShape(pptx.shapes.RECTANGLE,{x:0.8,y:1.55,w:1.2,h:0.04,fill:{color:C.gold},line:{color:C.gold}});
 s.addText('加密市场周报',{x:0.8,y:1.8,w:5.5,h:0.8,fontFace:'Arial',fontSize:42,bold:true,color:C.white,margin:0});
 s.addText('2026年5月4日 — 5月10日',{x:0.8,y:2.7,w:5.8,h:0.45,fontFace:'Arial',fontSize:20,color:C.gold,margin:0});
 s.addText('数据来源：Messari · The Block · CoinGecko · Reuters · CoinDesk',{x:0.8,y:3.4,w:7.3,h:0.3,fontFace:'Arial',fontSize:11,color:C.medGray,margin:0});
 s.addText('本周关键词：CLARITY Act、Meta稳定币、a16z Fund V、Hyperliquid、桥安全',{x:0.8,y:4.15,w:8.5,h:0.55,fontFace:'Arial',fontSize:14,color:C.white,margin:0});
 s.addShape(pptx.shapes.RECTANGLE,{x:0,y:5.25,w:10,h:0.375,fill:{color:C.gold},line:{color:C.gold}});
 s.addText('CONFIDENTIAL — FOR LP DISTRIBUTION ONLY',{x:0,y:5.25,w:10,h:0.375,align:'center',valign:'middle',fontFace:'Arial',fontSize:9,bold:true,color:C.darkNavy,margin:0});
 s.addNotes(`本页最重要的三件事：
1. 本周市场的核心不是单纯上涨，而是“主线越来越清楚”。政策、稳定币支付和 AI Agent 基础设施开始同时指向同一件事，也就是 crypto 正在从投机市场逐渐长成一套可用的金融与结算网络。
2. 封面上列出的几个关键词里，CLARITY Act 和 Meta 稳定币试点最值得重视。前者决定美国会不会给行业一个更可预期的监管框架，后者则代表 Web2 巨头终于开始把稳定币当成真实支付轨道，而不是概念实验。
3. 对组合而言，Hyperliquid、ENA、PENDLE 这类能承接真实流量和利率主题的资产，比单纯公链 beta 更值得跟踪。市场正在奖励有现金流捕获和场景闭环的方向。`);
}

{
 const s=pptx.addSlide(); s.background={color:C.offWhite}; addHeader(s,'📊  市场概览');
 const card=(x,l,v,d,c)=>{s.addShape(pptx.shapes.RECTANGLE,{x,y:0.95,w:2.8,h:1.05,fill:{color:C.white},line:{color:C.white},shadow:makeShadow()});s.addShape(pptx.shapes.RECTANGLE,{x,y:0.95,w:0.06,h:1.05,fill:{color:C.gold},line:{color:C.gold}});s.addText(l,{x:x+0.22,y:1.05,w:2.1,h:0.25,fontFace:'Arial',fontSize:12,bold:true,color:C.medGray,margin:0});s.addText(v,{x:x+0.22,y:1.3,w:2.2,h:0.35,fontFace:'Arial',fontSize:28,bold:true,color:C.bodyText,margin:0});s.addText(d,{x:x+0.22,y:1.68,w:2.2,h:0.2,fontFace:'Arial',fontSize:14,bold:true,color:c,margin:0});};
 card(0.6,'BTC','$81,347','+3.4%',C.green); card(3.7,'ETH','$2,349','+0.9%',C.green);
 s.addShape(pptx.shapes.RECTANGLE,{x:6.8,y:0.95,w:2.8,h:1.05,fill:{color:C.white},line:{color:C.white},shadow:makeShadow()}); s.addShape(pptx.shapes.RECTANGLE,{x:6.8,y:0.95,w:0.06,h:1.05,fill:{color:C.gold},line:{color:C.gold}});
 s.addText('恐惧贪婪指数',{x:7.02,y:1.05,w:2.1,h:0.25,fontFace:'Arial',fontSize:12,bold:true,color:C.medGray,margin:0}); s.addText('47',{x:7.02,y:1.3,w:0.8,h:0.35,fontFace:'Arial',fontSize:28,bold:true,color:C.bodyText,margin:0}); s.addText('Neutral',{x:7.72,y:1.38,w:1.4,h:0.2,fontFace:'Arial',fontSize:13,color:C.bodyText,margin:0});
 s.addTable([[{text:'资产',options:{bold:true,color:C.white,fill:{color:C.navy},align:'center',fontFace:'Arial',fontSize:11}},{text:'价格',options:{bold:true,color:C.white,fill:{color:C.navy},align:'center',fontFace:'Arial',fontSize:11}},{text:'7d',options:{bold:true,color:C.white,fill:{color:C.navy},align:'center',fontFace:'Arial',fontSize:11}},{text:'备注',options:{bold:true,color:C.white,fill:{color:C.navy},align:'center',fontFace:'Arial',fontSize:11}}],['SOL','$94.83','+12.7%','主流 beta 中最强'],['HYPE','$43.16','+5.9%','龙头地位稳固，解锁临近'],['ENA','$0.1325','+31.3%','稳定币收益主线受益'],['PENDLE','$1.99','+26.6%','利率交易继续升温'],['DYDX','$0.1720','+19.5%','低基数修复']],{x:0.6,y:2.25,w:8.8,colW:[1.2,1.7,1.2,4.7],rowH:[0.35,0.33,0.33,0.33,0.33,0.33],border:{pt:0.4,color:C.lightGray},fontFace:'Arial',fontSize:10,color:C.bodyText});
 s.addText('总市值 $2.79T，BTC 主导率 58.3%，24h 成交额 $60.4B。上涨延续，但量能未全面跟上。',{x:0.6,y:4.55,w:8.8,h:0.35,fontFace:'Arial',fontSize:10,italic:true,color:C.medGray,margin:0}); addFooter(s,'加密市场周报 | 2026.05.04–05.10');
 s.addNotes(`本页最重要的三件事：
1. BTC 重新站上 8 万美元上方，本质上意味着市场愿意为“监管预期改善 + 机构配置回流”重新付价。值得注意的是，这轮上涨并没有伴随特别夸张的成交量放大，所以它更像结构性重估，而不是全面狂热。
2. ETH 明显跑输 SOL、ENA、PENDLE 这类更高弹性资产，说明市场当前偏好的不是“以太坊大盘叙事”，而是能够承接稳定币、收益率、交易活动和 AI 主题的更直接资产。这种分化对选币比对大方向更重要。
3. 恐惧贪婪指数回到 47，代表市场已从防守区回到中性区。接下来若政策催化继续兑现，风险偏好有继续改善空间；但如果监管表述或稳定币奖励条款低于预期，当前这波修复也可能重新降温。`);
}

{
 const s=pptx.addSlide(); s.background={color:C.offWhite}; addHeader(s,'🔥  本周要闻 & 行业动态');
 s.addShape(pptx.shapes.RECTANGLE,{x:0.4,y:0.9,w:4.5,h:4.1,fill:{color:C.white},line:{color:C.white},shadow:makeShadow()}); s.addText('本周要闻 TOP 5',{x:0.6,y:0.96,w:3.5,h:0.3,fontFace:'Arial',fontSize:13,bold:true,color:C.navy,margin:0});
 const stories=[['1','a16z 募资 22 亿美元','一级市场继续押注稳定币与 AI Agent'],['2','5/14 审议 CLARITY Act','市场结构法案进入关键窗口'],['3','Meta × Stripe 稳定币试点','Web2 开始把稳定币用于创作者分发'],['4','Hyperliquid + HYPE 解锁','强基本面与供给扩张并存'],['5','桥与交易基础设施安全再被重估','资产迁移与风控产品化加速']];
 stories.forEach((it,i)=>{const y=1.38+i*0.72;s.addShape(pptx.shapes.OVAL,{x:0.64,y,w:0.3,h:0.3,fill:{color:C.gold},line:{color:C.gold}});s.addText(it[0],{x:0.64,y,w:0.3,h:0.3,fontFace:'Arial',fontSize:11,bold:true,color:C.white,align:'center',valign:'middle',margin:0});s.addText(it[1],{x:1.08,y:y-0.02,w:3.55,h:0.22,fontFace:'Arial',fontSize:10.5,bold:true,color:C.bodyText,margin:0});s.addText(it[2],{x:1.08,y:y+0.22,w:3.55,h:0.2,fontFace:'Arial',fontSize:9,color:C.medGray,margin:0});});
 const panel=(y,title,items)=>{s.addShape(pptx.shapes.RECTANGLE,{x:5.1,y,w:4.5,h:1.95,fill:{color:C.white},line:{color:C.white},shadow:makeShadow()});s.addShape(pptx.shapes.RECTANGLE,{x:5.1,y,w:4.5,h:0.35,fill:{color:C.navy},line:{color:C.navy}});s.addText(title,{x:5.3,y,w:3.9,h:0.35,fontFace:'Arial',fontSize:11,bold:true,color:C.white,valign:'middle',margin:0});s.addText(items.map((t,i)=>({text:t,options:{bullet:true,breakLine:i<items.length-1,fontFace:'Arial',fontSize:10,color:C.bodyText}})),{x:5.3,y:y+0.42,w:4.0,h:1.35,margin:0,paraSpaceAfterPt:5});};
 panel(0.9,'Layer 1 & Layer 2',['本周无压倒性大升级，主线让位于支付与监管','Pi Network 5/11 Protocol 23 值得观察','NEAR 强化 AI + 后量子安全叙事','SOL 继续是最强主流 beta']);
 panel(3.05,'DeFi & 永续合约',['Kelp 余波仍压制可组合杠杆偏好','Aave/Compound 更强调风控优先','Hyperliquid 龙头优势不改','收益型资产继续跑赢']);
 addFooter(s,'加密市场周报 | 2026.05.04–05.10');
 s.addNotes(`本页最重要的三件事：
1. 这页最关键的结论是，市场的中心已经从“哪条链更快”转到“谁能承接真实需求”。a16z 的募资方向、Meta 的稳定币试点、Hyperliquid 的交易扩张，背后共同指向的是支付、执行和结算，而不是单纯的公链性能竞赛。
2. L1/L2 本周相对平淡不是坏事，反而说明市场在自然筛选更成熟的主线。过去行业常常被新链叙事主导，但现在资金更愿意追逐能够直接产生交易量、结算流量或收费能力的基础设施，这种迁移通常意味着行业进入更成熟阶段。
3. DeFi 和永续合约的对比也很鲜明。一边是 Kelp 余波让市场重新审视可组合风险，另一边是 Hyperliquid 继续拿走更多交易份额。这说明用户并没有离开链上交易，只是更偏向简单、清晰、深流动性的头部平台。`);
}

{
 const s=pptx.addSlide(); s.background={color:C.offWhite}; addHeader(s,'⚖️  监管政策 & 投融资');
 s.addShape(pptx.shapes.RECTANGLE,{x:0.4,y:0.9,w:4.5,h:4.1,fill:{color:C.white},line:{color:C.white},shadow:makeShadow()}); s.addShape(pptx.shapes.RECTANGLE,{x:0.4,y:0.9,w:0.06,h:4.1,fill:{color:C.gold},line:{color:C.gold}});
 s.addText('监管与政策',{x:0.7,y:0.96,w:3.5,h:0.3,fontFace:'Arial',fontSize:14,bold:true,color:C.navy,margin:0});
 [['5/14 审议 CLARITY Act','市场结构与稳定币规则进入关键窗口'],['稳定币奖励出现折中版本','限制与银行存款等价的被动收益'],['SEC/CFTC 分工预期升温','交易所、钱包、DeFi 估值折价或下降'],['Meta 稳定币试点受监管关注','平台型支付扩张进入政策视野'],['政策重心转向“如何允许创新”','从粗放执法走向分类监管']].forEach((r,i)=>{const y=1.45+i*0.65;s.addText(r[0],{x:0.7,y,w:3.9,h:0.22,fontFace:'Arial',fontSize:11,bold:true,color:C.bodyText,margin:0});s.addText(r[1],{x:0.7,y:y+0.23,w:3.9,h:0.2,fontFace:'Arial',fontSize:9.5,color:C.medGray,margin:0});});
 s.addShape(pptx.shapes.RECTANGLE,{x:5.1,y:0.9,w:4.5,h:4.1,fill:{color:C.white},line:{color:C.white},shadow:makeShadow()}); s.addShape(pptx.shapes.RECTANGLE,{x:5.1,y:0.9,w:0.06,h:4.1,fill:{color:C.gold},line:{color:C.gold}});
 s.addText('投融资与并购',{x:5.4,y:0.96,w:3.5,h:0.3,fontFace:'Arial',fontSize:14,bold:true,color:C.navy,margin:0});
 s.addTable([[{text:'项目',options:{bold:true,color:C.white,fill:{color:C.navy},fontFace:'Arial',fontSize:10}},{text:'金额',options:{bold:true,color:C.white,fill:{color:C.navy},fontFace:'Arial',fontSize:10,align:'right'}},{text:'备注',options:{bold:true,color:C.white,fill:{color:C.navy},fontFace:'Arial',fontSize:10}}],['a16z Crypto Fund V','$2.2B','稳定币、Agent、crypto rails'],['Bullish × Equiniti','$4.2B','tokenized securities 基建整合'],['Payward × Reap','$600M','稳定币跨境支付扩张'],['OpenTrade / Squads','$17M / $18M','机构收益与企业多签工具']],{x:5.3,y:1.45,w:4.1,colW:[1.25,0.9,1.95],rowH:[0.3,0.32,0.32,0.32,0.32],border:{pt:0.4,color:C.lightGray},fontFace:'Arial',fontSize:9.3,color:C.bodyText});
 s.addShape(pptx.shapes.RECTANGLE,{x:5.3,y:3.35,w:4.1,h:0.85,fill:{color:C.paleGold},line:{color:C.paleGold},shadow:makeShadow()});
 s.addText('资本继续追逐三类资产：合规入口、稳定币分发、AI/Agent 金融基础设施。',{x:5.52,y:3.48,w:3.6,h:0.55,fontFace:'Arial',fontSize:10,color:C.bodyText,margin:0,valign:'mid'});
 addFooter(s,'加密市场周报 | 2026.05.04–05.10');
 s.addNotes(`本页最重要的三件事：
1. 监管现在进入了最有价值的阶段，也就是不再讨论“要不要管”，而是在讨论“怎么分类、谁来管、允许哪些商业模式”。对市场来说，这比单纯口头友好更重要，因为它直接决定机构敢不敢上线产品、敢不敢配置资金。
2. 稳定币奖励条款之争的本质，是银行存款和链上美元在争夺同一种用户余额。监管不愿意让 crypto 公司直接复制银行的收益型存款功能，所以未来真正能跑出来的模式，很可能是支付返利、生态积分、手续费返还等边界更清晰的设计。
3. 一级市场和并购也验证了这件事。a16z、Bullish、Payward 这些交易都不是在追逐短期热点，而是在提前卡位下一轮机构上链所需的底层设施。谁掌握合规入口、支付流量和资产登记能力，谁就更有可能拿走下一阶段的利润池。`);
}

{
 const s=pptx.addSlide(); s.background={color:C.offWhite}; addHeader(s,'🛡️  安全事件 & 重点赛道');
 s.addShape(pptx.shapes.RECTANGLE,{x:0.4,y:0.9,w:4.5,h:4.1,fill:{color:C.white},line:{color:C.white},shadow:makeShadow()}); s.addShape(pptx.shapes.RECTANGLE,{x:0.4,y:0.9,w:4.5,h:0.35,fill:{color:C.red},line:{color:C.red}}); s.addText('安全事件',{x:0.65,y:0.9,w:3.5,h:0.35,fontFace:'Arial',fontSize:12,bold:true,color:C.white,valign:'middle',margin:0});
 s.addTable([[{text:'项目',options:{bold:true,color:C.white,fill:{color:C.red},fontFace:'Arial',fontSize:10}},{text:'损失',options:{bold:true,color:C.white,fill:{color:C.red},fontFace:'Arial',fontSize:10}},{text:'说明',options:{bold:true,color:C.white,fill:{color:C.red},fontFace:'Arial',fontSize:10}}],['TrustedVolumes','$5.9M','RFQ 代理设计缺陷'],['Cryptomax','$200M+','交易所被盗，待更多核实'],['DSJEX','冻结 $41.5M','多方协作追缴进展'],['Solv','迁移 $700M','因桥安全担忧主动切换']],{x:0.6,y:1.35,w:4.05,colW:[1.05,0.95,2.05],rowH:[0.28,0.34,0.34,0.34,0.34],border:{pt:0.4,color:C.lightGray},fontFace:'Arial',fontSize:9.2,color:C.bodyText});
 s.addText('稳定币 / AI / 预测市场',{x:0.6,y:3.3,w:3.2,h:0.25,fontFace:'Arial',fontSize:12,bold:true,color:C.navy,margin:0});
 s.addText([{text:'• ',options:{color:C.gold}},{text:'Meta 与 Stripe 把稳定币推向创作者支付场景',options:{color:C.bodyText}},{text:'\n• ',options:{color:C.gold}},{text:'a16z、AWS、Google 继续强化 Agent 支付叙事',options:{color:C.bodyText}},{text:'\n• ',options:{color:C.gold}},{text:'预测市场正成为交易与信息定价的中间层',options:{color:C.bodyText}}],{x:0.7,y:3.62,w:3.9,h:1.0,fontFace:'Arial',fontSize:10,margin:0});
 const box=(y,title,items)=>{s.addShape(pptx.shapes.RECTANGLE,{x:5.1,y,w:4.5,h:1.95,fill:{color:C.white},line:{color:C.white},shadow:makeShadow()});s.addShape(pptx.shapes.RECTANGLE,{x:5.1,y,w:4.5,h:0.35,fill:{color:C.navy},line:{color:C.navy}});s.addText(title,{x:5.3,y,w:3.9,h:0.35,fontFace:'Arial',fontSize:11,bold:true,color:C.white,valign:'middle',margin:0});s.addText(items.map((t,i)=>({text:t,options:{bullet:true,breakLine:i<items.length-1,fontFace:'Arial',fontSize:10,color:C.bodyText}})),{x:5.3,y:y+0.42,w:4.0,h:1.35,margin:0,paraSpaceAfterPt:5});};
 box(0.9,'AI × Crypto',['a16z Fund V 明确押注 Agent 与稳定币','企业和 Agent 可能成为下一波稳定币主需求方','NEAR / HPC / compute 叙事继续受益','支付、身份、KYA 成为新中间件']);
 box(3.05,'预测市场',['Hyperliquid HIP-4 带来新交易层','Polymarket 继续向机构级监控靠拢','概率层正在被媒体、交易与 AI 共用','头部流动性进一步集中']);
 addFooter(s,'加密市场周报 | 2026.05.04–05.10');
 s.addNotes(`本页最重要的三件事：
1. 安全这周最值得注意的不是某一个漏洞本身，而是市场行为发生了变化。Solv 主动迁移 7 亿美元 tokenized BTC，说明桥安全已经从“技术人讨论的话题”变成了大资金的实际配置约束。未来桥、预言机、托管路径的可信度会直接影响 TVL 和估值。
2. 稳定币、AI Agent、预测市场三条线放在一起看，会更容易理解下一阶段行业增长从哪里来。稳定币负责结算，Agent 负责自动执行，预测市场提供实时概率和信息定价。如果这三层打通，行业会从单点协议竞争升级为完整金融操作系统竞争。
3. 预测市场尤其值得重视，因为它开始从“下注平台”变成“信息中间层”。一旦交易员、媒体、研究机构和 AI Agent 都开始引用同一套链上概率市场，头部平台会天然形成数据和流动性的双重网络效应。`);
}

{
 const s=pptx.addSlide(); s.background={color:C.offWhite}; addHeader(s,'🔮  催化剂 & 周度展望');
 s.addTable([[{text:'日期',options:{bold:true,color:C.white,fill:{color:C.navy},fontFace:'Arial',fontSize:10}},{text:'事件',options:{bold:true,color:C.white,fill:{color:C.navy},fontFace:'Arial',fontSize:10}},{text:'影响资产',options:{bold:true,color:C.white,fill:{color:C.navy},fontFace:'Arial',fontSize:10}},{text:'预期影响',options:{bold:true,color:C.white,fill:{color:C.navy},fontFace:'Arial',fontSize:10}}],['5/11','Pi Network Protocol 23','PI / 新公链','智能合约主题交易'],['5/12','APT 解锁约 1130 万枚','APT','中等供给压力'],['5/13–5/14','Digital Assets Week USA','RWA / 基建','机构合作催化'],['5/14','参议院审议 CLARITY Act','全市场 / 稳定币 / DeFi','最关键监管催化'],['5/15','STRK / SEI / AEVO 解锁窗口','L2 / 交易类资产','多资产波动抬升'],['5/16','ARB 解锁约 9265 万枚','ARB','供给压力']],{x:0.45,y:1.0,w:6.2,colW:[1.0,2.2,1.35,1.65],rowH:[0.32,0.3,0.3,0.3,0.3,0.3,0.32],border:{pt:0.4,color:C.lightGray},fontFace:'Arial',fontSize:9.3,color:C.bodyText});
 s.addShape(pptx.shapes.RECTANGLE,{x:6.95,y:1.0,w:2.6,h:3.9,fill:{color:C.white},line:{color:C.white},shadow:makeShadow()}); s.addShape(pptx.shapes.RECTANGLE,{x:6.95,y:1.0,w:2.6,h:0.35,fill:{color:C.navy},line:{color:C.navy}}); s.addText('周度总结',{x:7.1,y:1.0,w:2.1,h:0.35,fontFace:'Arial',fontSize:11,bold:true,color:C.white,valign:'middle',margin:0});
 s.addText([{text:'• ',options:{color:C.gold}},{text:'市场修复仍由少数高质量主线驱动',options:{color:C.bodyText}},{text:'\n• ',options:{color:C.gold}},{text:'5/14 CLARITY Act 是最重要观察点',options:{color:C.bodyText}},{text:'\n• ',options:{color:C.gold}},{text:'ENA / PENDLE 保持高弹性，HYPE 看供给消化',options:{color:C.bodyText}},{text:'\n• ',options:{color:C.gold}},{text:'桥安全与支付基础设施会继续影响资金分配',options:{color:C.bodyText}}],{x:7.18,y:1.48,w:2.0,h:1.9,fontFace:'Arial',fontSize:10,margin:0});
 s.addShape(pptx.shapes.RECTANGLE,{x:7.15,y:3.55,w:2.05,h:0.9,fill:{color:C.paleGold},line:{color:C.paleGold}}); s.addText('组合观察：\n继续跟踪政策兑现、稳定币分发和交易基础设施龙头。',{x:7.28,y:3.72,w:1.8,h:0.5,fontFace:'Arial',fontSize:9.5,bold:true,color:C.navy,margin:0,align:'center'});
 addFooter(s,'加密市场周报 | 2026.05.04–05.10');
 s.addNotes(`本页最重要的三件事：
1. 接下来一周最重要的事件毫无疑问是 5 月 14 日的 CLARITY Act 审议。它不一定立刻带来价格单边上涨，但会显著影响市场愿意给稳定币、交易平台、DeFi 前端和钱包类资产多少监管折价，因此对结构性定价非常关键。
2. 供给端也不能忽视。APT、ARB、STRK、AEVO 等一系列解锁集中出现，意味着即使宏观和政策层面继续改善，细分资产之间也会因为解锁压力产生明显分化。所以未来一到两周更像“选主线、控供给”的市场，而不是无差别上涨。
3. 组合层面，我更关注三类东西。第一是政策落地后谁最先受益，第二是稳定币分发和 Agent 支付是否继续扩张，第三是交易基础设施龙头是否能在波动中继续吸走流量。如果这三条线继续兑现，市场会更偏向质量资产而不是纯情绪资产。`);
}

pptx.writeFile('/Users/lx/Github/crypto-weekly-digest/reports/2026-05-04/slides-zh.pptx');
