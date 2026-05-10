const pptxgen = require('pptxgenjs');
const C = { navy:'1A2744', darkNavy:'0F1A2E', gold:'C9A96E', white:'FFFFFF', offWhite:'F5F6F8', lightGray:'E8EBF0', medGray:'8C95A6', bodyText:'2D3748', red:'C53030', green:'2F855A', paleGold:'FFF8E7' };
const pptx = new pptxgen();
pptx.layout = 'LAYOUT_16x9'; pptx.author = '42'; pptx.company='42'; pptx.subject='Crypto Weekly Digest'; pptx.title='Crypto Market Weekly 2026.05.04–05.10'; pptx.lang='en-US';
const makeShadow = () => ({ type:'outer', blur:3, offset:1, angle:45, color:'000000', opacity:0.12 });
const addFooter = (s,t)=>{ s.addShape(pptx.shapes.RECTANGLE,{x:0,y:5.25,w:10,h:0.375,fill:{color:C.navy},line:{color:C.navy}}); s.addText(t,{x:0.45,y:5.25,w:9.1,h:0.375,fontFace:'Arial',fontSize:8,color:C.medGray,valign:'middle',margin:0}); };
const addHeader = (s,t)=>{ s.addShape(pptx.shapes.RECTANGLE,{x:0,y:0,w:10,h:0.7,fill:{color:C.navy},line:{color:C.navy}}); s.addText(t,{x:0.6,y:0,w:8.8,h:0.7,fontFace:'Arial',fontSize:20,bold:true,color:C.white,valign:'middle',margin:0}); };

{
 const s=pptx.addSlide(); s.background={color:C.darkNavy};
 s.addShape(pptx.shapes.RECTANGLE,{x:0.8,y:1.55,w:1.2,h:0.04,fill:{color:C.gold},line:{color:C.gold}});
 s.addText('Crypto Market Weekly',{x:0.8,y:1.8,w:6,h:0.8,fontFace:'Arial',fontSize:42,bold:true,color:C.white,margin:0});
 s.addText('May 4 – May 10, 2026',{x:0.8,y:2.7,w:6,h:0.45,fontFace:'Arial',fontSize:20,color:C.gold,margin:0});
 s.addText('Sources: Messari · The Block · CoinGecko · Reuters · CoinDesk',{x:0.8,y:3.4,w:7.3,h:0.3,fontFace:'Arial',fontSize:11,color:C.medGray,margin:0});
 s.addText('Themes: CLARITY Act, Meta stablecoins, a16z Fund V, Hyperliquid, bridge security',{x:0.8,y:4.15,w:8.5,h:0.55,fontFace:'Arial',fontSize:14,color:C.white,margin:0});
 s.addShape(pptx.shapes.RECTANGLE,{x:0,y:5.25,w:10,h:0.375,fill:{color:C.gold},line:{color:C.gold}});
 s.addText('CONFIDENTIAL — FOR LP DISTRIBUTION ONLY',{x:0,y:5.25,w:10,h:0.375,align:'center',valign:'middle',fontFace:'Arial',fontSize:9,bold:true,color:C.darkNavy,margin:0});
 s.addNotes(`本页最重要的三件事：
1. 这一周对英文受众最需要解释的是，市场并不是简单地 risk-on 了，而是开始围绕少数高确定性主线重新定价。监管清晰度、稳定币支付和 AI Agent 基础设施正在互相强化，说明 crypto 正在从纯交易市场向真实金融网络过渡。
2. Meta 通过 Stripe 做稳定币创作者支付试点，是一个比价格波动更值得注意的信号。因为这代表大型互联网平台开始把稳定币看成实际的跨境分发工具，而不是只和 crypto native 用户相关的东西。
3. a16z 新基金和 Hyperliquid 的持续强势则说明资本市场的偏好已经很清楚，大家更愿意为结算、支付、执行和流动性基础设施买单，而不是为缺少场景支撑的叙事资产买单。`);
}

{
 const s=pptx.addSlide(); s.background={color:C.offWhite}; addHeader(s,'📊  Market Overview');
 const card=(x,l,v,d,c)=>{s.addShape(pptx.shapes.RECTANGLE,{x,y:0.95,w:2.8,h:1.05,fill:{color:C.white},line:{color:C.white},shadow:makeShadow()});s.addShape(pptx.shapes.RECTANGLE,{x,y:0.95,w:0.06,h:1.05,fill:{color:C.gold},line:{color:C.gold}});s.addText(l,{x:x+0.22,y:1.05,w:2.1,h:0.25,fontFace:'Arial',fontSize:12,bold:true,color:C.medGray,margin:0});s.addText(v,{x:x+0.22,y:1.3,w:2.2,h:0.35,fontFace:'Arial',fontSize:28,bold:true,color:C.bodyText,margin:0});s.addText(d,{x:x+0.22,y:1.68,w:2.2,h:0.2,fontFace:'Arial',fontSize:14,bold:true,color:c,margin:0});};
 card(0.6,'BTC','$81,347','+3.4%',C.green); card(3.7,'ETH','$2,349','+0.9%',C.green);
 s.addShape(pptx.shapes.RECTANGLE,{x:6.8,y:0.95,w:2.8,h:1.05,fill:{color:C.white},line:{color:C.white},shadow:makeShadow()}); s.addShape(pptx.shapes.RECTANGLE,{x:6.8,y:0.95,w:0.06,h:1.05,fill:{color:C.gold},line:{color:C.gold}});
 s.addText('Fear & Greed',{x:7.02,y:1.05,w:2.1,h:0.25,fontFace:'Arial',fontSize:12,bold:true,color:C.medGray,margin:0}); s.addText('47',{x:7.02,y:1.3,w:0.8,h:0.35,fontFace:'Arial',fontSize:28,bold:true,color:C.bodyText,margin:0}); s.addText('Neutral',{x:7.72,y:1.38,w:1.4,h:0.2,fontFace:'Arial',fontSize:13,color:C.bodyText,margin:0});
 s.addTable([[{text:'Asset',options:{bold:true,color:C.white,fill:{color:C.navy},align:'center',fontFace:'Arial',fontSize:11}},{text:'Price',options:{bold:true,color:C.white,fill:{color:C.navy},align:'center',fontFace:'Arial',fontSize:11}},{text:'7d',options:{bold:true,color:C.white,fill:{color:C.navy},align:'center',fontFace:'Arial',fontSize:11}},{text:'Notes',options:{bold:true,color:C.white,fill:{color:C.navy},align:'center',fontFace:'Arial',fontSize:11}}],['SOL','$94.83','+12.7%','Strongest major-beta performer'],['HYPE','$43.16','+5.9%','Leadership intact, unlock pressure'],['ENA','$0.1325','+31.3%','Stablecoin-yield theme'],['PENDLE','$1.99','+26.6%','Rates/yield trade strengthens'],['DYDX','$0.1720','+19.5%','Low-base rebound']],{x:0.6,y:2.25,w:8.8,colW:[1.2,1.7,1.2,4.7],rowH:[0.35,0.33,0.33,0.33,0.33,0.33],border:{pt:0.4,color:C.lightGray},fontFace:'Arial',fontSize:10,color:C.bodyText});
 s.addText('Total market cap $2.79T, BTC dominance 58.3%, 24h volume $60.4B. Rally extended, but broad participation is still incomplete.',{x:0.6,y:4.55,w:8.8,h:0.35,fontFace:'Arial',fontSize:10,italic:true,color:C.medGray,margin:0}); addFooter(s,'Crypto Market Weekly | 2026.05.04–05.10');
 s.addNotes(`本页最重要的三件事：
1. BTC 站稳 8 万美元上方，说明政策预期和机构资金流向正在提供新的定价锚。但成交量没有同步放大，意味着这轮上涨更多是集中在高质量资产上的重估，而不是全市场同步扩张。
2. ETH 跑输 SOL、ENA、PENDLE，说明本周市场追逐的是更直接的叙事承接者，比如稳定币收益、链上交易活动、支付和 AI 相关基础设施，而不是广义的大盘 beta。对 LP 来说，这种分化意味着选赛道和选资产的重要性高于单纯判断方向。
3. Fear & Greed 回到 47 代表情绪从防守切回中性，但还没有进入过热区域。如果下周 CLARITY Act 带来积极进展，情绪有可能继续上修；反过来，如果法案细节令市场失望，当前这波修复也容易出现回撤。`);
}

{
 const s=pptx.addSlide(); s.background={color:C.offWhite}; addHeader(s,'🔥  Top Stories & Industry');
 s.addShape(pptx.shapes.RECTANGLE,{x:0.4,y:0.9,w:4.5,h:4.1,fill:{color:C.white},line:{color:C.white},shadow:makeShadow()}); s.addText('TOP 5 STORIES',{x:0.6,y:0.96,w:3.5,h:0.3,fontFace:'Arial',fontSize:13,bold:true,color:C.navy,margin:0});
 const stories=[['1','a16z raises $2.2B Fund V','Fresh conviction in stablecoins and AI agents'],['2','CLARITY Act markup set for May 14','Market structure enters a key week'],['3','Meta × Stripe stablecoin pilot','Stablecoins move deeper into Web2 payouts'],['4','Hyperliquid + HYPE unlock','Strong fundamentals meet supply expansion'],['5','Bridge/security risk repriced','Infrastructure trust becomes a portfolio issue']];
 stories.forEach((it,i)=>{const y=1.38+i*0.72;s.addShape(pptx.shapes.OVAL,{x:0.64,y,w:0.3,h:0.3,fill:{color:C.gold},line:{color:C.gold}});s.addText(it[0],{x:0.64,y,w:0.3,h:0.3,fontFace:'Arial',fontSize:11,bold:true,color:C.white,align:'center',valign:'middle',margin:0});s.addText(it[1],{x:1.08,y:y-0.02,w:3.55,h:0.22,fontFace:'Arial',fontSize:10.5,bold:true,color:C.bodyText,margin:0});s.addText(it[2],{x:1.08,y:y+0.22,w:3.55,h:0.2,fontFace:'Arial',fontSize:9,color:C.medGray,margin:0});});
 const panel=(y,title,items)=>{s.addShape(pptx.shapes.RECTANGLE,{x:5.1,y,w:4.5,h:1.95,fill:{color:C.white},line:{color:C.white},shadow:makeShadow()});s.addShape(pptx.shapes.RECTANGLE,{x:5.1,y,w:4.5,h:0.35,fill:{color:C.navy},line:{color:C.navy}});s.addText(title,{x:5.3,y,w:3.9,h:0.35,fontFace:'Arial',fontSize:11,bold:true,color:C.white,valign:'middle',margin:0});s.addText(items.map((t,i)=>({text:t,options:{bullet:true,breakLine:i<items.length-1,fontFace:'Arial',fontSize:10,color:C.bodyText}})),{x:5.3,y:y+0.42,w:4.0,h:1.35,margin:0,paraSpaceAfterPt:5});};
 panel(0.9,'Layer 1 & Layer 2',['No dominant L1/L2 headline this week','Pi Protocol 23 is the next catalyst to watch','NEAR keeps linking AI and post-quantum security','SOL remains the strongest major-beta chain']);
 panel(3.05,'DeFi & Perp DEXes',['Kelp aftershocks still curb leverage appetite','Aave/Compound risk management back in focus','Hyperliquid keeps absorbing share','Yield-bearing assets continue to outperform']);
 addFooter(s,'Crypto Market Weekly | 2026.05.04–05.10');
 s.addNotes(`本页最重要的三件事：
1. 英文版这一页最想传达的是，行业注意力正在从基础设施性能竞争转向实际需求承接能力。a16z、Meta 和 Hyperliquid 代表的并不是三个分散事件，而是同一个趋势，也就是支付、执行、交易和结算正成为新一轮价值捕获中心。
2. L1/L2 本周相对安静，反而凸显出行业成熟。以前很多周报会被“新链、新升级、新 TPS”主导，但现在真正驱动市场的是谁能承接稳定币流量、企业支付、Agent 行为和持续交易活动，这种变化通常对长期投资者更重要。
3. DeFi 与 perp DEX 的结构也很清楚。可组合杠杆和桥风险让市场更谨慎，但用户并没有退出链上交易，而是在向更简单、更深流动性、更成熟风控的平台集中。Hyperliquid 的持续领先正好说明了这一点。`);
}

{
 const s=pptx.addSlide(); s.background={color:C.offWhite}; addHeader(s,'⚖️  Regulation & Fundraising');
 s.addShape(pptx.shapes.RECTANGLE,{x:0.4,y:0.9,w:4.5,h:4.1,fill:{color:C.white},line:{color:C.white},shadow:makeShadow()}); s.addShape(pptx.shapes.RECTANGLE,{x:0.4,y:0.9,w:0.06,h:4.1,fill:{color:C.gold},line:{color:C.gold}});
 s.addText('Regulation & Policy',{x:0.7,y:0.96,w:3.5,h:0.3,fontFace:'Arial',fontSize:14,bold:true,color:C.navy,margin:0});
 [['May 14 Senate markup','CLARITY Act enters a decisive window'],['Stablecoin rewards compromise','Bank-like passive yield still restricted'],['SEC/CFTC split in focus','Could reduce discount on infra assets'],['Meta pilot draws policy attention','Platform-scale stablecoin expansion begins'],['Debate shifts to permitted innovation','From blanket enforcement to classification']].forEach((r,i)=>{const y=1.45+i*0.65;s.addText(r[0],{x:0.7,y,w:3.9,h:0.22,fontFace:'Arial',fontSize:11,bold:true,color:C.bodyText,margin:0});s.addText(r[1],{x:0.7,y:y+0.23,w:3.9,h:0.2,fontFace:'Arial',fontSize:9.5,color:C.medGray,margin:0});});
 s.addShape(pptx.shapes.RECTANGLE,{x:5.1,y:0.9,w:4.5,h:4.1,fill:{color:C.white},line:{color:C.white},shadow:makeShadow()}); s.addShape(pptx.shapes.RECTANGLE,{x:5.1,y:0.9,w:0.06,h:4.1,fill:{color:C.gold},line:{color:C.gold}});
 s.addText('Fundraising & M&A',{x:5.4,y:0.96,w:3.5,h:0.3,fontFace:'Arial',fontSize:14,bold:true,color:C.navy,margin:0});
 s.addTable([[{text:'Project',options:{bold:true,color:C.white,fill:{color:C.navy},fontFace:'Arial',fontSize:10}},{text:'Amount',options:{bold:true,color:C.white,fill:{color:C.navy},fontFace:'Arial',fontSize:10,align:'right'}},{text:'Notes',options:{bold:true,color:C.white,fill:{color:C.navy},fontFace:'Arial',fontSize:10}}],['a16z Crypto Fund V','$2.2B','Stablecoins, agents, crypto rails'],['Bullish × Equiniti','$4.2B','Tokenized securities infrastructure'],['Payward × Reap','$600M','Stablecoin cross-border payments'],['OpenTrade / Squads','$17M / $18M','Yield infra and enterprise tooling']],{x:5.3,y:1.45,w:4.1,colW:[1.25,0.9,1.95],rowH:[0.3,0.32,0.32,0.32,0.32],border:{pt:0.4,color:C.lightGray},fontFace:'Arial',fontSize:9.3,color:C.bodyText});
 s.addShape(pptx.shapes.RECTANGLE,{x:5.3,y:3.35,w:4.1,h:0.85,fill:{color:C.paleGold},line:{color:C.paleGold},shadow:makeShadow()});
 s.addText('Capital keeps clustering around compliant access, stablecoin distribution, and AI/agent finance infrastructure.',{x:5.52,y:3.48,w:3.6,h:0.55,fontFace:'Arial',fontSize:10,color:C.bodyText,margin:0,valign:'mid'});
 addFooter(s,'Crypto Market Weekly | 2026.05.04–05.10');
 s.addNotes(`本页最重要的三件事：
1. 监管的价值不在于短期 headline，而在于它是否能降低长期不确定性。CLARITY Act 一旦推动 SEC/CFTC 权限边界更明确，交易平台、钱包、DeFi 前端和稳定币业务的估值基础就会更稳，因为机构终于可以据此设计产品和风控框架。
2. 稳定币奖励条款的争议反映了一个更深层的冲突，也就是链上美元到底能不能复制银行存款的经济属性。如果监管最终不允许直接做“类存款收益”，市场就会转向更间接的用户激励和支付返利模式，这会改变稳定币商业模式的利润分配。
3. Fundraising 与 M&A 也很能说明问题。a16z、Bullish、Payward 这几笔交易分别覆盖了风险资本、交易基础设施和支付网络，说明聪明资金正在沿着同一条路线布局，也就是把 crypto 嵌入更大范围的金融和互联网分发系统。`);
}

{
 const s=pptx.addSlide(); s.background={color:C.offWhite}; addHeader(s,'🛡️  Security & Key Sectors');
 s.addShape(pptx.shapes.RECTANGLE,{x:0.4,y:0.9,w:4.5,h:4.1,fill:{color:C.white},line:{color:C.white},shadow:makeShadow()}); s.addShape(pptx.shapes.RECTANGLE,{x:0.4,y:0.9,w:4.5,h:0.35,fill:{color:C.red},line:{color:C.red}}); s.addText('Security Incidents',{x:0.65,y:0.9,w:3.5,h:0.35,fontFace:'Arial',fontSize:12,bold:true,color:C.white,valign:'middle',margin:0});
 s.addTable([[{text:'Project',options:{bold:true,color:C.white,fill:{color:C.red},fontFace:'Arial',fontSize:10}},{text:'Loss',options:{bold:true,color:C.white,fill:{color:C.red},fontFace:'Arial',fontSize:10}},{text:'Comment',options:{bold:true,color:C.white,fill:{color:C.red},fontFace:'Arial',fontSize:10}}],['TrustedVolumes','$5.9M','RFQ proxy design flaw'],['Cryptomax','$200M+','Exchange breach, still being verified'],['DSJEX','$41.5M frozen','Recovery progress'],['Solv','$700M migrated','Security-driven bridge switch']],{x:0.6,y:1.35,w:4.05,colW:[1.05,0.95,2.05],rowH:[0.28,0.34,0.34,0.34,0.34],border:{pt:0.4,color:C.lightGray},fontFace:'Arial',fontSize:9.2,color:C.bodyText});
 s.addText('Stablecoins / AI / Prediction',{x:0.6,y:3.3,w:3.4,h:0.25,fontFace:'Arial',fontSize:12,bold:true,color:C.navy,margin:0});
 s.addText([{text:'• ',options:{color:C.gold}},{text:'Meta and Stripe push stablecoins into creator payouts',options:{color:C.bodyText}},{text:'\n• ',options:{color:C.gold}},{text:'a16z, AWS, and Google reinforce the agent-payments thesis',options:{color:C.bodyText}},{text:'\n• ',options:{color:C.gold}},{text:'Prediction markets keep becoming an information layer',options:{color:C.bodyText}}],{x:0.7,y:3.62,w:3.9,h:1.0,fontFace:'Arial',fontSize:10,margin:0});
 const box=(y,title,items)=>{s.addShape(pptx.shapes.RECTANGLE,{x:5.1,y,w:4.5,h:1.95,fill:{color:C.white},line:{color:C.white},shadow:makeShadow()});s.addShape(pptx.shapes.RECTANGLE,{x:5.1,y,w:4.5,h:0.35,fill:{color:C.navy},line:{color:C.navy}});s.addText(title,{x:5.3,y,w:3.9,h:0.35,fontFace:'Arial',fontSize:11,bold:true,color:C.white,valign:'middle',margin:0});s.addText(items.map((t,i)=>({text:t,options:{bullet:true,breakLine:i<items.length-1,fontFace:'Arial',fontSize:10,color:C.bodyText}})),{x:5.3,y:y+0.42,w:4.0,h:1.35,margin:0,paraSpaceAfterPt:5});};
 box(0.9,'AI × Crypto',['a16z Fund V validates agent/stablecoin rails','Enterprises and agents may drive next stablecoin demand','NEAR / HPC / compute continue to benefit','Payments, identity, and KYA become key middleware']);
 box(3.05,'Prediction Markets',['HIP-4 expands Hyperliquid’s trading stack','Polymarket keeps moving toward institutional oversight','Probability markets are becoming shared data rails','Liquidity remains heavily concentrated at the top']);
 addFooter(s,'Crypto Market Weekly | 2026.05.04–05.10');
 s.addNotes(`本页最重要的三件事：
1. 安全问题现在越来越像配置问题，而不只是工程问题。Solv 因为桥安全担忧迁移 7 亿美元资产，说明大资金已经把跨链路径、消息验证和托管方式视为决定是否配置某个协议的重要因素，这会持续改变资本流向。
2. 稳定币、AI Agent 和预测市场放在一起看，能更清楚地看到下一阶段基础设施长什么样。稳定币提供价值转移，Agent 负责自动执行，预测市场提供实时概率和信息定价，这三层如果结合，会构成新的 crypto-native operating stack。
3. 预测市场的价值正在从“下注”升级为“概率 API”。一旦交易员、研究机构、媒体乃至 AI Agent 都把同一套链上概率市场当成参考数据，头部平台就会积累强大的网络效应，这对长期估值非常重要。`);
}

{
 const s=pptx.addSlide(); s.background={color:C.offWhite}; addHeader(s,'🔮  Catalysts & Outlook');
 s.addTable([[{text:'Date',options:{bold:true,color:C.white,fill:{color:C.navy},fontFace:'Arial',fontSize:10}},{text:'Event',options:{bold:true,color:C.white,fill:{color:C.navy},fontFace:'Arial',fontSize:10}},{text:'Assets',options:{bold:true,color:C.white,fill:{color:C.navy},fontFace:'Arial',fontSize:10}},{text:'Expected Impact',options:{bold:true,color:C.white,fill:{color:C.navy},fontFace:'Arial',fontSize:10}}],['5/11','Pi Network Protocol 23','PI / new-chain theme','Smart-contract catalyst'],['5/12','APT unlock ~11.3M','APT','Moderate supply pressure'],['5/13–5/14','Digital Assets Week USA','RWA / infra','Institutional partnership catalyst'],['5/14','Senate CLARITY Act markup','Broad market / stablecoins / DeFi','Most important policy catalyst'],['5/15','STRK / SEI / AEVO unlock window','L2s / trading assets','Cross-asset volatility'],['5/16','ARB unlock ~92.65M','ARB','Supply pressure']],{x:0.45,y:1.0,w:6.2,colW:[1.0,2.2,1.35,1.65],rowH:[0.32,0.3,0.3,0.3,0.3,0.3,0.32],border:{pt:0.4,color:C.lightGray},fontFace:'Arial',fontSize:9.3,color:C.bodyText});
 s.addShape(pptx.shapes.RECTANGLE,{x:6.95,y:1.0,w:2.6,h:3.9,fill:{color:C.white},line:{color:C.white},shadow:makeShadow()}); s.addShape(pptx.shapes.RECTANGLE,{x:6.95,y:1.0,w:2.6,h:0.35,fill:{color:C.navy},line:{color:C.navy}}); s.addText('Weekly Takeaway',{x:7.1,y:1.0,w:2.1,h:0.35,fontFace:'Arial',fontSize:11,bold:true,color:C.white,valign:'middle',margin:0});
 s.addText([{text:'• ',options:{color:C.gold}},{text:'The rally is still concentrated in quality themes',options:{color:C.bodyText}},{text:'\n• ',options:{color:C.gold}},{text:'May 14 is the key policy date to watch',options:{color:C.bodyText}},{text:'\n• ',options:{color:C.gold}},{text:'ENA/PENDLE remain high-beta winners; HYPE must absorb supply',options:{color:C.bodyText}},{text:'\n• ',options:{color:C.gold}},{text:'Bridge trust and payment rails still shape flows',options:{color:C.bodyText}}],{x:7.18,y:1.48,w:2.0,h:1.9,fontFace:'Arial',fontSize:10,margin:0});
 s.addShape(pptx.shapes.RECTANGLE,{x:7.15,y:3.55,w:2.05,h:0.9,fill:{color:C.paleGold},line:{color:C.paleGold}}); s.addText('Portfolio lens:\nStay focused on policy delivery, stablecoin distribution, and infra leaders.',{x:7.28,y:3.72,w:1.8,h:0.5,fontFace:'Arial',fontSize:9.5,bold:true,color:C.navy,margin:0,align:'center'});
 addFooter(s,'Crypto Market Weekly | 2026.05.04–05.10');
 s.addNotes(`本页最重要的三件事：
1. 下周的最核心催化就是 5 月 14 日的 CLARITY Act 审议。它未必马上决定价格方向，但会决定市场是否愿意进一步压缩监管折价，尤其是对稳定币、交易平台、钱包和 DeFi 基础设施这几类资产来说，影响会比对纯 beta 资产更直接。
2. 供给端会让市场继续分化。APT、ARB、STRK、AEVO 等解锁提醒我们，即便大方向偏积极，单个资产也可能因为供给释放而表现落后。所以未来一到两周是一个非常典型的“选主线、躲供给、看政策”的阶段。
3. 组合层面，我会继续盯三件事。第一，政策是否真的往更清晰的分类监管迈进；第二，稳定币是否继续向 Web2 分发和 Agent 支付扩张；第三，Hyperliquid 这类基础设施龙头能否在供给和波动中继续吃到更多流量。如果答案大多是肯定的，市场会越来越偏好质量而不是噪音。`);
}

pptx.writeFile('/Users/lx/Github/crypto-weekly-digest/reports/2026-05-04/slides-en.pptx');
