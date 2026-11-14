/* Standalone reproduction of workbook P&L; all monetary values in KRW. */
(function(root){
"use strict";
const assumptions=[{"row": 10, "label": "본체 실결제가", "values": [[129000.0, 129000.0, 129000.0], [149000.0, 149000.0, 149000.0], [169000.0, 169000.0, 169000.0]]}, {"row": 14, "label": "완제품 키트 제조원가", "values": [[70000.0, 65000.0, 60000.0], [55000.0, 50000.0, 46000.0], [45000.0, 40000.0, 37000.0]]}, {"row": 18, "label": "유료 유입 클릭당 비용", "values": [[1200.0, 1100.0, 1000.0], [900.0, 800.0, 750.0], [700.0, 650.0, 600.0]]}, {"row": 22, "label": "유효 클릭의 출고 전환율", "values": [[0.015, 0.018, 0.02], [0.025, 0.03, 0.035], [0.04, 0.045, 0.05]]}, {"row": 26, "label": "본체 DTC 전액 환불률", "values": [[0.15, 0.12, 0.1], [0.08, 0.07, 0.06], [0.05, 0.04, 0.03]]}, {"row": 30, "label": "첫 리필 유료 참여율", "values": [[0.1, 0.12, 0.15], [0.25, 0.3, 0.35], [0.4, 0.45, 0.5]]}, {"row": 34, "label": "리필 구매자 월 유지율", "values": [[0.75, 0.78, 0.8], [0.9, 0.91, 0.92], [0.94, 0.95, 0.96]]}, {"row": 38, "label": "성숙 월 자연 유입", "values": [[500.0, 1000.0, 2000.0], [1000.0, 3000.0, 5000.0], [1500.0, 5000.0, 10000.0]]}, {"row": 42, "label": "성숙 월 광고비", "values": [[1500000.0, 3000000.0, 5000000.0], [2000000.0, 6000000.0, 12000000.0], [3000000.0, 10000000.0, 25000000.0]]}, {"row": 46, "label": "첫 출시월 번호", "values": [[10.0, 10.0, 10.0], [7.0, 7.0, 7.0], [5.0, 5.0, 5.0]]}, {"row": 50, "label": "리테일 성숙 월 잠재 출고", "values": [[0.0, 0.0, 100.0], [0.0, 100.0, 300.0], [0.0, 200.0, 800.0]]}, {"row": 54, "label": "리테일 공급가 비율", "values": [[0.55, 0.55, 0.55], [0.6, 0.6, 0.6], [0.65, 0.65, 0.65]]}, {"row": 58, "label": "경제적 인력 규모", "values": [[3.0, 3.0, 3.0], [3.0, 3.0, 4.0], [3.0, 4.0, 6.0]]}, {"row": 62, "label": "인당 월 총인건비", "values": [[4500000.0, 4700000.0, 5000000.0], [4500000.0, 4700000.0, 5000000.0], [4500000.0, 4700000.0, 5000000.0]]}, {"row": 66, "label": "월 고정 운영비", "values": [[2000000.0, 2500000.0, 3000000.0], [2000000.0, 3000000.0, 4000000.0], [2000000.0, 4000000.0, 6000000.0]]}, {"row": 70, "label": "리필 24개팩 실결제가", "values": [[9900.0, 9900.0, 9900.0], [12900.0, 12900.0, 12900.0], [14900.0, 14900.0, 14900.0]]}, {"row": 74, "label": "리필 24개팩 원가", "values": [[4500.0, 4300.0, 4000.0], [3500.0, 3200.0, 3000.0], [2800.0, 2500.0, 2300.0]]}];
const common={"r85": 0.1, "r86": 0.033, "r87": 4500.0, "r88": 6000.0, "r89": 2000.0, "r90": 1000.0, "r91": 3000.0, "r92": 0.033, "r93": 0.02, "r94": 1.0, "r95": 0.0, "r96": 2.0, "r97": 2000.0, "r98": 0.03, "r99": 300.0, "r100": 100.0, "r101": 2.0, "r102": 0.5, "r103": 25000000.0, "r104": 36.0, "r105": 300000000.0, "r106": 30000000.0, "r107": 0.0, "r108": 15000000.0, "r109": 10000000.0, "r110": 1000000.0, "r111": 0.2, "r112": 0.3, "r113": 0.1, "r114": 3.0, "r115": 20000.0, "r116": 12.0, "r117": 24.0, "r118": 3.0, "rd": [15000000.0, 1500000.0, 1500000.0]};
const sum=(o,a,b)=>Array.from({length:b-a+1},(_,i)=>o['r'+(a+i)]||0).reduce((a,b)=>a+b,0);
const ratio=(a,b)=>b>0?a/b:null;
function simulate(scenario=2,overrides={}){
  if(![1,2,3].includes(scenario))throw new RangeError('scenario must be 1, 2, or 3');
  const factors={priceFactor:1,costFactor:1,cpcFactor:1,conversionFactor:1,refundFactor:1,refillFactor:1,...overrides};
  for(const [key,value] of Object.entries(factors))if(!Number.isFinite(value)||value<0)throw new RangeError('Invalid factor: '+key);
  const inputs=[0,1,2].map(y=>Object.fromEntries(assumptions.map(a=>['r'+a.row,a.values[scenario-1][y]])));
  for(const a of inputs){
    a.r10*=factors.priceFactor;a.r14*=factors.costFactor;a.r18*=factors.cpcFactor;
    a.r22=Math.min(1,a.r22*factors.conversionFactor);a.r26=Math.min(1,a.r26*factors.refundFactor);a.r30=Math.min(1,a.r30*factors.refillFactor);
  }
  const c=common;
  const unit=inputs.map((a,y)=>{
    const u={year:y+1,r8:a.r10,r9:a.r10/(1+c.r85),r10:1-a.r26,r12:a.r14,r14:c.r87,r15:a.r26*c.r88,r16:c.r89,r17:c.r90};
    u.r11=u.r9*u.r10;u.r13=u.r11*c.r86;u.r18=u.r11-sum(u,12,17);u.r19=ratio(a.r18,a.r22);u.r20=u.r19===null?null:u.r18-u.r19;u.r21=u.r20===null?null:ratio(u.r20,u.r11);
    u.r22=c.r115;u.r23=u.r18-u.r22;u.r24=ratio(a.r18,u.r23);u.r25=a.r58*a.r62+a.r66;u.r26=u.r20>0?Math.ceil(u.r25/u.r20):null;
    u.r30=a.r10/(1+c.r85)*a.r54;u.r31=1-c.r98;u.r32=u.r30*u.r31;u.r33=a.r14;u.r34=c.r97;u.r35=c.r98*c.r88;u.r36=c.r89;u.r37=c.r90;u.r38=u.r32-sum(u,33,37);
    u.r43=a.r70;u.r44=a.r70/(1+c.r85)*(1-c.r93);u.r45=a.r74;u.r46=u.r44*c.r92;u.r47=c.r91;u.r48=c.r93*c.r88;u.r49=u.r44-sum(u,45,48);
    u.r50=a.r30;u.r51=a.r34;u.r52=u.r10*u.r50*c.r94*(u.r51===1?11:(1-u.r51**11)/(1-u.r51));u.r53=u.r52*u.r49;u.r54=u.r18+u.r53;u.r55=u.r19===null?null:u.r54-u.r19;u.r56=u.r19===null?null:ratio(u.r54,u.r19);
    return u;
  });
  const launch=inputs[0].r46,months=[];
  for(let i=0;i<36;i++){
    const year=Math.floor(i/12)+1,month=i+1,a=inputs[year-1],p=months[i-1];
    const m={year,month,r8:month>=launch?1:0};
    m.r9=month<launch?0:month===launch?.5:month===launch+1?.75:1;
    m.r10=a.r42*m.r9;m.r11=a.r18;m.r12=m.r11>0?m.r10/m.r11:0;m.r13=a.r38*m.r9;m.r14=a.r22;
    m.r15=Math.floor(m.r12*m.r14);m.r16=Math.floor(m.r13*m.r14);m.r17=m.r15+m.r16;m.r18=Math.floor(a.r50*m.r9*c.r95);m.r19=m.r17+m.r18;
    m.r20=m.r17*a.r26;m.r21=m.r17-m.r20;m.r22=m.r18*c.r98;m.r23=m.r21+m.r18-m.r22;m.r24=a.r30;m.r25=a.r34;
    m.r26=p?p.r21*p.r24:0;m.r27=p?p.r28*m.r25:0;m.r28=m.r26+m.r27;m.r29=m.r28*c.r94;
    months.push(m);
  }
  let inventory=0,inventoryValue=0;
  for(let i=0;i<36;i++){
    const m=months[i],a=inputs[m.year-1],u=unit[m.year-1];
    const next3=months.slice(i,i+3).reduce((s,x)=>s+x.r19,0);
    const received=m.month>=launch&&(m.month-launch)%3===0?Math.max(0,Math.ceil((next3+c.r100-inventory)/c.r99)*c.r99):0;
    const avgCost=inventory+received>0?(inventoryValue+received*a.r14)/(inventory+received):0;
    m.r30=m.r21*a.r10/(1+c.r85);m.r31=(m.r18-m.r22)*a.r10/(1+c.r85)*a.r54;m.r32=m.r29*u.r44;m.r33=sum(m,30,32);
    m.r34=m.r19*avgCost;m.r35=m.r29*a.r74;m.r36=m.r33-m.r34-m.r35;
    inventory+=received-m.r19;inventoryValue+=received*a.r14-m.r34;
    m.inventory={received,endingUnits:inventory,endingValue:inventoryValue,avgCost};
    m.r37=m.r30*c.r86;m.r38=m.r17*c.r87;m.r39=(m.r20+m.r22)*c.r88;m.r40=m.r19*c.r89;m.r41=m.r19*c.r90;m.r42=m.r18*c.r97;m.r43=m.r32*c.r92;m.r44=m.r29*c.r91;m.r45=m.r29*c.r93*c.r88;
    m.r46=m.r36-sum(m,37,45);m.r47=m.r10;m.r48=m.r46-m.r47;m.r49=a.r58*a.r62;m.r50=a.r66;
    m.r51=m.month<=4?c.r109*[.2,.3,.3,.2][i]:0;
    if(m.month>=4&&m.month<=6)m.r51+=c.rd[0];
    if(m.year>1)m.r51=c.rd[m.year-1];
    m.r52=m.r48-sum(m,49,51);
    const depStart=Math.max(1,launch-2);
    m.r53=m.month>=depStart&&m.month<depStart+c.r104?c.r103/c.r104:0;
    m.r54=m.r52-m.r53;m.r55=ratio(m.r48,m.r33);m.r56=ratio(m.r47,m.r17);m.r57=m.r49+m.r50;
  }
  const years=[1,2,3].map(year=>{
    const group=months.slice((year-1)*12,year*12),r={year};
    for(let n=8;n<=57;n++)r['r'+n]=group.reduce((s,m)=>s+(m['r'+n]||0),0);
    // These annual cells are intentionally blank in Excel: rates are not summed.
    for(const n of [8,9,11,14,24,25])r['r'+n]=null;
    r.r28=group[11].r28;r.r55=ratio(r.r48,r.r33);r.r56=ratio(r.r47,r.r17);
    return r;
  });
  return {scenario,factors,inputs,common:c,months,years,unit};
}
const api={simulate,assumptions,common};
if(typeof module!=='undefined'&&module.exports)module.exports=api;
root.BusinessModel=api;
})(globalThis);
