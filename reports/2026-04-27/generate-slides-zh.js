const pptxgen = require('pptxgenjs');
const C = { navy:'1A2744', dark:'0F1A2E', gold:'C9A96E', white:'FFFFFF', off:'F5F6F8', gray:'8C95A6', text:'2D3748', light:'E8EBF0', red:'C53030', green:'2F855A' };
const shadow = { type:'outer', blur:3, offset:1, angle:45, color:'000000', opacity:0.12 };
const week='2026年4月27日 – 5月3日';
const footer='加密市场周报 | 2026.04.27–05.03';
let p;
function header(s,t){ s.addShape(p.shapes.RECTANGLE,{x:0,y:0,w:10,h:0.7,fill:{color:C.navy}}); s.addText(t,{x:0.55,y:0,w:9,h:0.7,fontFace:'Arial',fontSize:20,bold:true,color:C.white,valign:'mid'}); }
function footerBar(s,txt=footer){ s.addShape(p.shapes.RECTANGLE,{x:0,y:5.25,w:10,h:0.375,fill:{color:C.navy}}); s.addText(txt,{x:0.45,y:5.25,w:9.1,h:0.375,fontFace:'Arial',fontSize:8,color:C.gray,valign:'mid'}); }
function bullets(s,items,x,y,w,h,fs=10,color=C.text){ s.addText(items.map((t,i)=>({text:t,options:{bullet:true,breakLine:i<items.length-1,fontFace:'Arial',fontSize:fs,color}})),{x,y,w,h,margin:0,paraSpaceAfter:4}); }
function note(s,txt){ s.addNotes(txt); }
async function main(){ p = new pptxgen(); p.layout='LAYOUT_16x9'; p.author='42'; p.subject='crypto weekly digest'; p.title='加密市场周报';
let s=p.addSlide(); s.background={color:C.dark};
s.addShape(p.shapes.RECTANGLE,{x:0.8,y:1.6,w:1.2,h:0.04,fill:{color:C.gold}});
s.addText('加密市场周报',{x:0.8,y:1.85,w:7,h:0.8,fontFace:'Arial',fontSize:32,bold:true,color:C.white});
s.addText(week,{x:0.8,y:2.55,w:7,h:0.45,fontFace:'Arial',fontSize:19,color:C.gold});
s.addText('数据来源：CoinGecko · Messari · The Block · 历史周报模板整理',{x:0.8,y:3.2,w:8.3,h:0.3,fontFace:'Arial',fontSize:11,color:C.gray});
s.addShape(p.shapes.RECTANGLE,{x:0,y:5.25,w:10,h:0.375,fill:{color:C.gold}});
s.addText('CONFIDENTIAL — FOR LP DISTRIBUTION ONLY',{x:0,y:5.25,w:10,h:0.375,fontFace:'Arial',fontSize:9,bold:true,color:C.dark,align:'center',valign:'mid'});
note(s,'第1页最重要的三点：\n1. 这一周的核心不是全面牛市回归，而是情绪从恐慌区继续修复。恐惧贪婪指数回到 47，说明市场已经不再像上一周那样被系统性安全冲击主导，但资金依然高度挑剔。\n2. 本周最大的结构变化是 BTC 主导率继续升到 58.4%。这背后的原因，是机构资金仍然更愿意配置最清晰、流动性最深、监管路径也最明确的资产，比特币继续扮演整个市场的风险锚。\n3. 对 LP 来说，这份周报的主线不是追逐全面 beta，而是识别资金真正愿意付估值的方向，包括 BTC、稳定币基础设施、链上交易基础设施，以及能提供明确现金流的收益类协议。');

s=p.addSlide(); s.background={color:C.off}; header(s,'📊  市场概览');
const cards=[['BTC','$78,643','+0.8%',C.green],['ETH','$2,327','-0.8%',C.red],['恐惧贪婪','47 Neutral',null,C.gold]];
[0,1,2].forEach(i=>{ const x=0.6+i*3.05; s.addShape(p.shapes.RECTANGLE,{x,y:0.95,w:2.8,h:1.05,fill:{color:C.white},shadow}); s.addShape(p.shapes.RECTANGLE,{x,y:0.95,w:0.06,h:1.05,fill:{color:cards[i][3]}}); s.addText(cards[i][0],{x:x+0.22,y:1.08,w:2,h:0.2,fontFace:'Arial',fontSize:12,bold:true,color:C.gray}); s.addText(cards[i][1],{x:x+0.22,y:1.35,w:2.2,h:0.28,fontFace:'Arial',fontSize:i===2?24:28,bold:true,color:C.text}); if(cards[i][2]) s.addText(cards[i][2],{x:x+0.22,y:1.72,w:1.7,h:0.2,fontFace:'Arial',fontSize:13,bold:true,color:cards[i][3]}); });
const rows=[['资产','价格','7d','备注'],['SOL','$84.13','-2.8%','高 beta 修复不足'],['HYPE','$40.98','-0.9%','龙头强于板块'],['PENDLE','$1.58','+21.1%','本周最强'],['ENA','$0.10','-8.0%','收益稳定币承压'],['DYDX','$0.14','-11.6%','份额叙事弱化']].map((r,idx)=>r.map((t,j)=>({text:t,options: idx===0?{fill:{color:C.navy},color:C.white,bold:true,fontFace:'Arial',fontSize:11,align:'center'}:{fontFace:'Arial',fontSize:10.5,bold:j===0,align:j===1?'right':'left',color:j===2?(t.includes('-')?C.red:C.green):C.text}})));
s.addTable(rows,{x:0.6,y:2.2,w:8.8,colW:[1.25,1.7,1.2,4.65],rowH:[0.35,0.32,0.32,0.32,0.32,0.32],border:{pt:0.5,color:C.light}});
s.addText('总市值 $2.70T，BTC 主导率 58.4%，24h 成交额 $51.7B。市场修复更多体现为恐慌消退，而非风险偏好的全面回归。',{x:0.6,y:4.55,w:8.8,h:0.4,fontFace:'Arial',fontSize:10,color:C.gray,italic:true}); footerBar(s);
note(s,'第2页最重要的三点：\n1. BTC、ETH、SOL 的表现继续分化。BTC 周涨 0.8% 并维持在 7.8 万美元上方，ETH 与 SOL 仍然偏弱，说明市场并没有进入典型的山寨轮动阶段。资金当前更在乎确定性，而不是弹性。\n2. PENDLE 单周上涨 21.1% 很关键。它说明在风险偏好没有全面回来的情况下，市场开始重新给“收益率、现金流、利率曲线交易”更高权重，这对 DeFi 的风格切换很有参考价值。\n3. 恐惧贪婪指数从上周 33 修复到 47，代表情绪层面最坏的阶段暂时过去，但离真正的 risk-on 还差一段距离。后续如果 ETH 继续跑输，说明这轮修复仍以避险和龙头集中为主。');

s=p.addSlide(); s.background={color:C.off}; header(s,'🔥  本周要闻 & 行业动态');
s.addShape(p.shapes.RECTANGLE,{x:0.4,y:0.9,w:4.5,h:4.1,fill:{color:C.white},shadow}); s.addText('本周要闻 TOP 5',{x:0.6,y:0.96,w:3,h:0.25,fontFace:'Arial',fontSize:13,bold:true,color:C.navy});
const top=[['1','BTC 主导率升至 58.4%','资金继续集中在最确定资产'],['2','情绪修复至 Neutral 47','恐慌缓解但尚未全面 risk-on'],['3','Hyperliquid 龙头格局未变','链上永续继续向头部集中'],['4','PENDLE 周涨 21.1%','收益率叙事重新受追捧'],['5','监管与机构化主线延续','更像中期利多而非短期催化']];
top.forEach((r,i)=>{ const y=1.36+i*0.72; s.addShape(p.shapes.OVAL,{x:0.65,y,w:0.3,h:0.3,fill:{color:C.gold}}); s.addText(r[0],{x:0.65,y,w:0.3,h:0.3,fontFace:'Arial',fontSize:11,bold:true,color:C.white,align:'center',valign:'mid'}); s.addText(r[1],{x:1.08,y:y-0.02,w:3.55,h:0.22,fontFace:'Arial',fontSize:11,bold:true,color:C.text}); s.addText(r[2],{x:1.08,y:y+0.22,w:3.55,h:0.2,fontFace:'Arial',fontSize:9,color:C.gray}); });
[['Layer 1 & Layer 2',['L2 竞争继续从技术竞赛走向生态分发战','RWA 与稳定币结算支撑高吞吐链基本面','链上证券预期提升，机构流量成下一阶段焦点']],['DeFi & 永续合约',['收益类协议相对跑赢，PENDLE 表现突出','Hyperliquid 维持流动性优势，HYPE 抗跌','尾部协议与老牌 perp 平台继续被压估值']]].forEach((box,idx)=>{ const x=5.1,y=idx?3.05:0.9; s.addShape(p.shapes.RECTANGLE,{x,y,w:4.5,h:1.95,fill:{color:C.white},shadow}); s.addShape(p.shapes.RECTANGLE,{x,y,w:4.5,h:0.35,fill:{color:C.navy}}); s.addText(box[0],{x:x+0.2,y,w:4,h:0.35,fontFace:'Arial',fontSize:11,bold:true,color:C.white,valign:'mid'}); bullets(s,box[1],x+0.2,y+0.43,4.05,1.35,10); }); footerBar(s);
note(s,'第3页最重要的三点：\n1. 本周真正的主线是“集中”。BTC 吸走全市场最稳的机构资金，Hyperliquid 吸走链上永续最强的流动性，DeFi 里则是收益率资产吸走更有限的风险预算。市场并没有同时奖励所有叙事。\n2. L1/L2 赛道的变化也很明确，单纯讲 TPS 或模块化已经不够了。未来更重要的是谁能承接稳定币、RWA、链上证券和真实机构流量，这会决定长期估值中枢。\n3. 对组合而言，PENDLE 的超额收益提示我们，资金开始重新偏好“有现金流解释、有收益率锚”的资产。这比追逐高波动新概念更符合当前市场阶段。');

s=p.addSlide(); s.background={color:C.off}; header(s,'⚖️  监管政策 & 投融资');
s.addShape(p.shapes.RECTANGLE,{x:0.4,y:0.9,w:4.5,h:4.1,fill:{color:C.white},shadow}); s.addShape(p.shapes.RECTANGLE,{x:0.4,y:0.9,w:0.06,h:4.1,fill:{color:C.gold}}); s.addText('监管与政策',{x:0.7,y:0.98,w:3,h:0.25,fontFace:'Arial',fontSize:14,bold:true,color:C.navy});
const reg=['稳定币收益分配边界仍是法案博弈核心','链上证券与 tokenized assets 继续升温','监管进展更像中期估值利多，不是短线点火器','市场开始从“是否监管”转向“谁将受益”']; reg.forEach((t,i)=>{s.addText(t,{x:0.72,y:1.45+i*0.72,w:3.9,h:0.25,fontFace:'Arial',fontSize:11,bold:i<2,color:C.text});});
s.addShape(p.shapes.RECTANGLE,{x:5.1,y:0.9,w:4.5,h:4.1,fill:{color:C.white},shadow}); s.addShape(p.shapes.RECTANGLE,{x:5.1,y:0.9,w:0.06,h:4.1,fill:{color:C.gold}}); s.addText('投融资与并购',{x:5.38,y:0.98,w:3,h:0.25,fontFace:'Arial',fontSize:14,bold:true,color:C.navy});
const tbl=[['方向','观察','意义'],['交易/托管','机构更偏好合规基础设施','承接真实资金流'],['RWA','继续获得高关注','更接近真实现金流'],['稳定币分发','支付/卡/钱包层价值凸显','利润不一定留在 token'],['链上清算层','整合趋势延续','强者更强']].map((r,idx)=>r.map((t)=>({text:t,options:idx===0?{fill:{color:C.navy},color:C.white,bold:true,fontFace:'Arial',fontSize:10}:{fontFace:'Arial',fontSize:9.5,color:C.text}})));
s.addTable(tbl,{x:5.3,y:1.42,w:4.05,colW:[1.1,1.3,1.65],rowH:[0.3,0.32,0.32,0.32,0.32],border:{pt:0.5,color:C.light}});
s.addShape(p.shapes.RECTANGLE,{x:5.3,y:3.45,w:4.05,h:0.72,fill:{color:'FFF8E7'},shadow}); s.addShape(p.shapes.RECTANGLE,{x:5.3,y:3.45,w:0.06,h:0.72,fill:{color:C.gold}}); s.addText('关注点：监管框架越清晰，真正受益的往往是发行、分发、托管、交易与清算层，而不是所有相关 token 同步受益。',{x:5.48,y:3.55,w:3.65,h:0.45,fontFace:'Arial',fontSize:10,color:C.text}); footerBar(s);
note(s,'第4页最重要的三点：\n1. 监管线索没有消失，反而越来越具体。现在市场讨论的重点，已经从“美国会不会监管”变成“稳定币收益能不能分配、谁有资格分配、资本金怎么定”，这说明行业在走向制度化深水区。\n2. 这类监管进展为什么暂时没有明显推高价格？因为它更偏中期估值框架改善，而不是短期交易催化。市场需要看到规则落地之后，真实收入和真实流量究竟沉淀在哪一层。\n3. 所以投融资也越来越集中在基础设施。未来更值得配置的，未必是叙事最热的 token，而是能承接支付、托管、清算、RWA 发行与机构交易的底层平台。');

s=p.addSlide(); s.background={color:C.off}; header(s,'🔓  安全事件 & 热点赛道');
s.addShape(p.shapes.RECTANGLE,{x:0.4,y:0.9,w:4.5,h:2.2,fill:{color:C.white},shadow}); s.addShape(p.shapes.RECTANGLE,{x:0.4,y:0.9,w:4.5,h:0.35,fill:{color:C.red}}); s.addText('安全事件 | 本周未见新增系统性冲击',{x:0.58,y:0.9,w:4.1,h:0.35,fontFace:'Arial',fontSize:11,bold:true,color:C.white,valign:'mid'});
const sec=[['观察','解读'],['无超大级别新增事故','情绪修复的重要原因'],['跨链/预言机/权限管理','仍是最敏感风险点'],['AI 辅助开发风险','审计压力继续上升'],['市场更看系统传染性','不只看单次损失金额']].map((r,idx)=>r.map(t=>({text:t,options:idx===0?{fill:{color:C.light},bold:true,fontFace:'Arial',fontSize:9,color:C.text}:{fontFace:'Arial',fontSize:9,color:C.text}})));
s.addTable(sec,{x:0.55,y:1.35,w:4.2,colW:[1.45,2.75],rowH:[0.28,0.3,0.3,0.3,0.3],border:{pt:0.5,color:C.light}});
[['稳定币',['监管博弈压制 ENA 等收益稳定币相关资产','链上美元扩张仍在继续','价值捕获更可能留在分发和结算层']],['AI × Crypto',['机构更偏好算力与推理基础设施','赛道从概念炒作转向质量筛选','AI 同时提升增长潜力与安全复杂度']],['预测市场',['交易化品类继续扩张','监管边界仍是最大变量','若规则清晰，容量上限可能上修']]].forEach((box,idx)=>{const x=5.1,y=0.9+idx*1.35,h=1.15; s.addShape(p.shapes.RECTANGLE,{x,y,w:4.5,h,fill:{color:C.white},shadow}); s.addShape(p.shapes.RECTANGLE,{x,y,w:4.5,h:0.3,fill:{color:C.navy}}); s.addText(box[0],{x:x+0.2,y,w:4,h:0.3,fontFace:'Arial',fontSize:11,bold:true,color:C.white,valign:'mid'}); bullets(s,box[1],x+0.2,y+0.34,4.05,0.72,9.5);}); footerBar(s);
note(s,'第5页最重要的三点：\n1. 本周没有出现类似上周 Kelp DAO / rsETH 那样的大型系统性安全事故，这是情绪能够继续修复的重要原因之一。市场在经历大事故后，最怕的不是单点损失，而是风险继续向借贷、清算和稳定币层传染。\n2. 稳定币、AI 和预测市场这三个赛道都还在，但它们的交易逻辑已经比几个月前更严格。稳定币看监管与分发，AI 看真实需求与基础设施，预测市场看政策边界是否放松。\n3. 安全问题并没有过去，尤其是 AI 辅助开发让代码产出更快，但也可能让错误更快扩散。后续如果再出现因权限、跨链、预言机或 AI 代码缺陷导致的事故，市场惩罚会非常直接。');

s=p.addSlide(); s.background={color:C.off}; header(s,'🔮  未来催化剂 & 展望');
const cat=[['日期','事件','影响资产','预期影响'],['5/5–5/7','Consensus 2026 Miami','全市场','合作、监管表态与融资窗口'],['5月上旬','稳定币/市场结构法案推进','稳定币/交易基建','重塑行业估值框架'],['5月上旬','ETH 扩容与 Gas Limit 测试','ETH/OP/ARB/STRK','影响扩容效率预期'],['5月','Bitcoin Hyper 主网上线目标期','BTCFi/BTC 生态','强化 BTC 新叙事'],['持续','Hyperliquid 生态扩张跟踪','HYPE/perp DEX','验证龙头是否继续集中']].map((r,idx)=>r.map((t,j)=>({text:t,options:idx===0?{fill:{color:C.navy},color:C.white,bold:true,fontFace:'Arial',fontSize:10,align:j===0||j===2?'center':'left'}:{fontFace:'Arial',fontSize:9.5,bold:j===2,color:C.text,fill: idx===5?{color:'FFF3E0'}:undefined,align:j===0||j===2?'center':'left'}})));
s.addTable(cat,{x:0.5,y:0.9,w:9,colW:[1.0,3.1,1.3,3.6],rowH:[0.32,0.32,0.32,0.32,0.32,0.32],border:{pt:0.5,color:C.light}});
s.addShape(p.shapes.RECTANGLE,{x:0.5,y:3.65,w:9,h:1.28,fill:{color:C.white},shadow}); s.addShape(p.shapes.RECTANGLE,{x:0.5,y:3.65,w:0.06,h:1.28,fill:{color:C.gold}}); s.addText('本周小结',{x:0.78,y:3.72,w:2,h:0.25,fontFace:'Arial',fontSize:13,bold:true,color:C.navy}); s.addText('当前市场更像“高质量资产的防守型修复”，而不是全面 risk-on。BTC 继续提供主锚，Hyperliquid 代表链上交易基础设施的集中趋势，PENDLE 则提示收益率资产重新获得定价权。未来两周重点观察监管落地速度、ETH 扩容预期，以及 HYPE 生态扩张是否继续强化 perp DEX 龙头逻辑。',{x:0.78,y:4.02,w:8.45,h:0.72,fontFace:'Arial',fontSize:10,color:C.text}); footerBar(s,footer+' | CONFIDENTIAL');
note(s,'第6页最重要的三点：\n1. 接下来两周最值得看的，不是短线价格波动，而是几个会改变估值框架的催化剂。Consensus 2026 可能带来监管、合作与融资层面的新信号，尤其适合观察机构叙事是否继续强化。\n2. 组合层面需要重点盯住三条线：第一是 BTC 是否继续吸收风险预算；第二是 Hyperliquid 是否继续证明链上永续的集中化趋势；第三是 ETH 扩容预期能否重新带动 ETH beta 修复。\n3. 总体判断仍偏谨慎乐观。最坏的恐慌阶段已经过去，但真正的新一轮全面上涨，需要看到 ETH 与高质量山寨同步转强，以及监管利好开始转化为可见的资金流入和收入增长。');
await p.writeFile({fileName:'/Users/lx/Github/crypto-weekly-digest/reports/2026-04-27/slides-zh.pptx'});
}
main().catch(err=>{ console.error(err); process.exit(1); });
