// Full-formula Excel model generator, readable and auditable build.
// Runs identically in Node (test) and browser (ExcelJS). buildFullModel(ExcelJSlib, S) -> workbook.
//
// Layout standard (one geometry on every sheet):
//  - Column A carries the section number, B the sub-number, C the row label, D the subtotal label.
//  - Column E carries the unit, plain and short. Column F carries every single value.
//  - Columns H, I and J reproduce the inputs a line's calculation uses, so F2 on any year cell
//    highlights its drivers on the same row. Green = linked from Inputs, black = derived here.
//  - The timeline starts in column K. The year header sits in row 2 and is frozen together with
//    columns A to J (panes lock at K3). Gridlines are off on every sheet.
//  - Percentages are formatted as percentages, years as plain numbers without separators.
//  - NO named ranges. NO IF(): conditions are 0/1 flags built by boolean multiplication;
//    caps and floors use MIN/MAX. IFERROR only prints "n/m" for an undefined IRR.
//  - Helper rows are free: every intermediate step is its own labelled row.
//  - Two blank rows breathe between sections. Tie-out check rows compare against the dashboard.
function buildFullModel(ExcelJS, S){
 const wb=new ExcelJS.Workbook(); wb.creator='Nexwell'; wb.created=new Date();
 wb.calcProperties.fullCalcOnLoad=true;
 const Y0=2026, YN=2060, NY=YN-Y0+1, c0=11;                    // timeline starts in column K
 function colL(n){let s='';while(n>0){const m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=(n-m-1)/26;}return s;}
 const lastC=colL(c0+NY-1);
 // styles
 const YEL={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFF2CC'}};   // editable input fill
 const HDR={type:'pattern',pattern:'solid',fgColor:{argb:'FF1F3864'}};   // year header navy
 const CHK={type:'pattern',pattern:'solid',fgColor:{argb:'FFE2EFDA'}};   // check green fill
 const BLUE={color:{argb:'FF0000FF'}};                                   // editable input font
 const GRN ={color:{argb:'FF008000'}};                                   // linked-from-Inputs font
 const bold={bold:true}, white={color:{argb:'FFFFFFFF'},bold:true};
 const SECF={bold:true,size:11,color:{argb:'FF1F7A33'}};
 const F={num:'#,##0.00;[Red]-#,##0.00',eur:'#,##0.0',int:'#,##0',yr:'0',x:'0.00"x"',
   pct0:'0%',pct1:'0.0%',pct2:'0.00%',fac:'0.0000',price:'#,##0.0',date:'yyyy-mm-dd'};

 // ================= INPUTS =================
 const wi=wb.addWorksheet('Inputs');
 wi.views=[{state:'frozen',xSplit:6,showGridLines:false}];               // columns frozen, rows free
 [2.5,3,46,2,14,14,2].forEach((w,i)=>wi.getColumn(i+1).width=w);
 wi.getCell(1,2).value='Project Burgenland'; wi.getCell(1,2).font={bold:true,size:15,color:{argb:'FF1F7A33'}};
 wi.getCell(2,3).value='Inputs | Energy SPV model | exported '+S.today; wi.getCell(2,3).font={italic:true,size:9,color:{argb:'FF808080'}};
 const IN={};
 let r=4, isNo=0;
 function isect(txt){wi.getCell(r,1).value=++isNo; wi.getCell(r,1).font=SECF;
  const c=wi.getCell(r,3); c.value=txt; c.font=SECF; wi.getRow(r).height=17; r++;}
 function inp(label,value,unit,key,numfmt){
  wi.getCell(r,3).value=label;
  wi.getCell(r,5).value=unit||''; wi.getCell(r,5).font={size:10,color:{argb:'FF7A7A7A'}};
  const c=wi.getCell(r,6); c.value=value; c.fill=YEL; c.font=BLUE;
  c.border={top:{style:'hair'},left:{style:'hair'},bottom:{style:'hair'},right:{style:'hair'}};
  c.numFmt=numfmt||(typeof value==='number'&&Math.abs(value)<1&&value!==0?'0.####':(Number.isInteger(value)?'#,##0':'#,##0.####'));
  IN[key]=r; r++;
 }
 isect('MACRO / FINANCING');
 inp('Inflation (CPI)',S.macro.infl,'per yr','INFL',F.pct1);
 inp('Tax rate',S.macro.tax,'%','TAXR',F.pct0);
 inp('Gearing (wind / solar / line)',S.macro.gearing,'% of capex','GEAR',F.pct0);
 inp('Debt tenor',S.macro.tenor,'years','TENOR',F.yr);
 inp('Amortisation (1 = annuity, 0 = flat)',S.macro.amort==='annuity'?1:0,'1/0','AMORT',F.yr);
 inp('All-in senior debt rate',S.macro.allInRate,'per yr','RATE',F.pct2);
 inp('PPA term',S.macro.ppaTermY,'years','PPAT',F.yr);
 inp('Merchant power price',S.macro.merchReal,'EUR/MWh','MERCH',F.price);
 r++; isect('PPA TRANCHES');
 inp('Tranche pricing on (1 = auto blend)',S.tranche&&S.tranche.on!=null?S.tranche.on:1,'1/0','T_ON',F.yr);
 inp('Tranche 1 size, wind',S.tranche?S.tranche.t1W:100,'MW','T1W',F.int);
 inp('Tranche 1 size, solar',S.tranche?S.tranche.t1S:100,'MWp','T1S',F.int);
 inp('Tranche 1 price',S.tranche?S.tranche.p1:100,'EUR/MWh','TP1',F.price);
 inp('Tranche 2 price',S.tranche?S.tranche.p2:80,'EUR/MWh','TP2',F.price);
 r++; isect('CONSTRUCTION FUNDING (Energy SPV)');
 inp('Funding coupon',S.sub?S.sub.subRate:0.08,'per yr','SUBR',F.pct1);
 inp('Burgenland share while outstanding',S.sub?S.sub.beShare:0.20,'% of cash','SUBBE',F.pct0);
 inp('Coupon step-up',S.sub?S.sub.stepAdd:0,'per yr','SUBSTEP',F.pct1);
 inp('Step-up after',S.sub?S.sub.stepY:5,'op. years','SUBSTEPY',F.yr);
 r++; isect('WIND');
 inp('Capacity',S.wind.mw,'MW','W_MW',F.int);
 inp('Capex',S.wind.capexPerMW,'EURm/MW','W_CAPEX','0.000');
 inp('Gross capacity factor',S.wind.grossCF,'%','W_GCF',F.pct2);
 inp('Plant losses',S.wind.loss,'%','W_LOSS',F.pct1);
 inp('Direct-line losses',S.wind.lineLoss,'%','W_LLOSS',F.pct1);
 inp('Degradation',S.wind.degr,'per yr','W_DEGR',F.pct2);
 inp('Opex',S.wind.opexPerMW,'EURm/MW/yr','W_OPEX','0.000');
 inp('PPA price, manual override',S.wind.ppa,'EUR/MWh','W_PPA',F.price);
 inp('Contracted share',S.wind.contr,'%','W_CONTR',F.pct0);
 inp('COD (first generation)',S.wind.codY,'year','W_COD',F.yr);
 inp('Useful life',S.wind.lifeY,'years','W_LIFE',F.yr);
 inp('Merchant tail price',S.CAP7w,'EUR/MWh','W_TAIL',F.price);
 r++; isect('SOLAR');
 inp('Capacity',S.solar.mw,'MWp','S_MW',F.int);
 inp('Capex',S.solar.capexPerMW,'EURm/MWp','S_CAPEX','0.000');
 inp('Gross capacity factor',S.solar.grossCF,'%','S_GCF',F.pct2);
 inp('Plant losses',S.solar.loss,'%','S_LOSS',F.pct1);
 inp('Direct-line losses',S.solar.lineLoss,'%','S_LLOSS',F.pct1);
 inp('Degradation',S.solar.degr,'per yr','S_DEGR',F.pct2);
 inp('Opex',S.solar.opexPerMW,'EURm/MWp/yr','S_OPEX','0.000');
 inp('PPA price, manual override',S.solar.ppa,'EUR/MWh','S_PPA',F.price);
 inp('Contracted share',S.solar.contr,'%','S_CONTR',F.pct0);
 inp('COD (first generation)',S.solar.codY,'year','S_COD',F.yr);
 inp('Useful life',S.solar.lifeY,'years','S_LIFE',F.yr);
 inp('Merchant tail price',S.CAP7s,'EUR/MWh','S_TAIL',F.price);
 r++; isect('BATTERY');
 inp('Power',S.battery.powerMW,'MW','B_MW',F.int);
 inp('Duration',S.battery.durationH,'hours','B_DUR',F.yr);
 inp('Cell capex',S.battery.capexPerKWh,'EUR/kWh','B_CKWH',F.int);
 inp('Grid-interface allowance',S.battery.substation,'EURm','B_SUB',F.eur);
 inp('Campus connection allowance',S.battery.interconnect,'EURm','B_INT',F.eur);
 inp('Round-trip efficiency',S.battery.rte,'%','B_RTE',F.pct0);
 inp('SoC floor (outage reserve)',S.battery.socFloor,'%','B_SOC',F.pct0);
 inp('Cycles per day',S.battery.cyclesDay,'x','B_CYC',F.yr);
 inp('Capture factor',S.battery.captureFactor,'x','B_CAPF','0.00');
 inp('Opex',S.battery.opexPct,'% of cell capex','B_OPEXP',F.pct1);
 inp('Ancillary revenue',S.battery.ancPerMW,'EURk/MW/yr','B_ANC',F.int);
 inp('DC reliability charge',S.battery.capChargeMWyr,'EURk/MW/yr','B_CCH',F.int);
 inp('First merchant year',S.battery.gridYear,'year','B_GY',F.yr);
 inp('Revenue compression',S.battery.compression,'per yr','B_COMP',F.pct1);
 inp('Degradation',S.battery.degr,'per yr','B_DEGR',F.pct1);
 inp('Gearing',S.battery.gearing,'% of capex','B_GEAR',F.pct0);
 inp('Debt rate',S.battery.debtRate,'per yr','B_RATE',F.pct2);
 inp('Grid fee: capacity, manual',S.battery.gridCapFee,'EUR/kW/yr','B_GFC',F.eur);
 inp('Apply DC NE3 tariff when merchant',S.battery.mktCapFee?1:0,'1/0','B_MKT',F.yr);
 inp('Grid fee: energy',S.battery.gridEnergyFee,'EUR/MWh','B_GFE',F.eur);
 inp('Backtest average buy price',S.battery.backtestAvgBuy,'EUR/MWh','B_AVGBUY',F.price);
 inp('Backtest average sell price',S.battery.backtestAvgSell,'EUR/MWh','B_AVGSELL',F.price);
 inp('Backtest self-charge share',S.battery.backtestSelfShare,'%','B_FSELF',F.pct0);
 inp('Self-charge cost',S.battery.btmCharge||0,'EUR/MWh','B_BTM',F.price);
 inp('Battery COD (capex year)',S.COD,'year','B_COD',F.yr);
 inp('Useful life',S.battery.lifeY,'years','B_LIFE',F.yr);
 r++; isect('DATA CENTER / GRID BALANCING');
 inp('DC contracted load',S.dc.firmMW,'MW','DC_MW',F.int);
 inp('DC sale price (today)',S.dc.dcPrice,'EUR/MWh','DC_P',F.price);
 inp('Residual market price',S.dc.resFix,'EUR/MWh','RES_P',F.price);
 inp('Energy-trading margin',S.dc.beMargin,'%','RES_M',F.pct1);
 inp('NE3 energy fee (2028)',S.dc.gridEnergyFee,'EUR/MWh','FEE_E',F.price);
 inp('NE3 capacity fee (2028)',S.dc.gridCapFeeKW,'EUR/kW/yr','FEE_C',F.price);
 inp('Grid-fee escalation',S.dc.feeEsc,'per yr','FEESC',F.pct1);
 inp('RES sourcing (1 = owned, 0 = PPA)',(S.dc.resMode||'lcoe')==='lcoe'?1:0,'1/0','RES_OWN',F.yr);
 inp('Direct line cost',S.linePer100,'EURm/100MW','LINE_C','0.00');
 inp('Solar DC/AC ratio',S.dcac,'x','DCAC','0.00');
 inp('Billing (1 = pass-through, 0 = fixed)',(S.dc.spvMode||'pass')==='pass'?1:0,'1/0','SPV_MODE',F.yr);
 inp('SPV margin, percent form',S.dc.spvMargin!=null?S.dc.spvMargin:0.03,'% of cost','SPV_M',F.pct1);
 inp('Margin form (1 = EUR/MWh, 0 = %)',S.dc.marginMode==='flat'?1:0,'1/0','MGN_MODE',F.yr);
 inp('SPV margin, euro form',S.dc.marginEur!=null?S.dc.marginEur:3.5,'EUR/MWh','MGN_E',F.price);
 inp('SPV first revenue year',S.FF,'year','SPV_FF',F.yr);
 r++; isect('REFERENCE AVERAGE-DAY PRICE CURVE | '+S.priceYear);
 const PH0=r;
 for(let h=0;h<24;h++){ wi.getCell(r,3).value='Hour '+h; wi.getCell(r,5).value='EUR/MWh';
  wi.getCell(r,5).font={size:10,color:{argb:'FF7A7A7A'}};
  const c=wi.getCell(r,6); c.value=S.ph[h]; c.fill=YEL; c.font=BLUE; c.numFmt=F.price; r++; }
 IN.PH0=PH0;
 const inCell=key=>`Inputs!$F$${IN[key]}`;

 // ================= SHARED SHEET HELPERS =================
 function calcSheet(name,subtitle){
  const ws=wb.addWorksheet(name);
  ws.views=[{state:'frozen',xSplit:10,ySplit:2,showGridLines:false}];    // panes lock at K3
  [2.5,3,46,20,12,13,2,11,11,11].forEach((w,i)=>ws.getColumn(i+1).width=w);
  for(let i=0;i<NY;i++)ws.getColumn(c0+i).width=10.5;
  ws.getCell(1,2).value='Project Burgenland'; ws.getCell(1,2).font={bold:true,size:13,color:{argb:'FF1F7A33'}};
  ws.getCell(1,4).value=subtitle; ws.getCell(1,4).font={italic:true,size:9,color:{argb:'FF808080'}};
  ws.getCell(2,3).value='Year'; ws.getCell(2,3).font=bold;
  for(let i=0;i<NY;i++){const c=ws.getCell(2,c0+i); c.value=Y0+i; c.font=white; c.fill=HDR; c.numFmt=F.yr;}
  return ws;
 }
 const YROW=2, yr=X=>`${X}$${YROW}`;
 function sect(ws,rr,num,txt){ws.getCell(rr,1).value=num; ws.getCell(rr,1).font=SECF;
  const c=ws.getCell(rr,3); c.value=txt; c.font=SECF; return rr+1;}
 // audit cells: reproduce a row's scalar drivers in H, I, J. Green = linked from Inputs.
 function audit(ws,row,aud){
  const cols=[8,9,10];
  (aud||[]).slice(0,3).forEach((a,i)=>{
   const c=ws.getCell(row,cols[i]);
   if(a==null)return;
   if(a.in){c.value={formula:inCell(a.in)};c.font=GRN;}
   else if(a.f){c.value={formula:a.f};}
   else if(a.v!=null){c.value=a.v;}
   c.numFmt=a.fmt||'General';
  });
  return ['$H$'+row,'$I$'+row,'$J$'+row];
 }
 // time-series row: label in C (or D for subtotals), unit in E, formula filled K..; aud drives H/I/J.
 function tsRow(ws,row,label,unit,fmtStr,fFn,opt){
  opt=opt||{};
  const lc=ws.getCell(row,opt.subtotal?4:3); lc.value=label; if(opt.subtotal||opt.bold)lc.font=bold;
  if(unit){const u=ws.getCell(row,5); u.value=unit; u.font={size:9.5,color:{argb:'FF7A7A7A'}};}
  const refs=audit(ws,row,opt.aud);
  for(let i=0;i<NY;i++){
   const X=colL(c0+i), pX=i?colL(c0+i-1):null;
   const c=ws.getCell(row,c0+i); c.value={formula:fFn(X,pX,refs)}; c.numFmt=fmtStr;
   if(opt.subtotal||opt.bold)c.font=bold;
   if(opt.subtotal)c.border={top:{style:'thin'}};
   if(opt.fill)c.fill=opt.fill;
  }
  return row;
 }
 // derived single value: label C, unit E, value F. Green only when the formula just links Inputs.
 function derRow(ws,row,D,key,label,formula,fmtStr,unit,green){
  ws.getCell(row,3).value=label;
  if(unit){const u=ws.getCell(row,5); u.value=unit; u.font={size:9.5,color:{argb:'FF7A7A7A'}};}
  const c=ws.getCell(row,2+4); c.value={formula:formula}; c.numFmt=fmtStr||F.num; if(green)c.font=GRN;
  D[key]=row; return row+1;
 }
 const lnk=(ws,row,D,key,label,inKey,fmtStr,unit)=>derRow(ws,row,D,key,label,inCell(inKey),fmtStr,unit,true);

 // ================= WIND / SOLAR builder =================
 function assetSheet(name,P,chk,chkWf){
  const ws=calcSheet(name,name+' | full-formula model | green = linked from Inputs, black = calculated here');
  const D={}, R={};
  let rr=4;
  rr=sect(ws,rr,1,'ASSUMPTIONS (linked from Inputs)');
  [['INFL','Inflation (CPI)','per yr',F.pct1],['TAXR','Tax rate','%',F.pct0],['GEAR','Gearing','% of capex',F.pct0],
   ['TENOR','Debt tenor','years',F.yr],['AMORT','Amortisation (1 = annuity)','1/0',F.yr],['RATE','Senior debt rate','per yr',F.pct2],
   ['PPAT','PPA term','years',F.yr],['MERCH','Merchant power price','EUR/MWh',F.price],
   [P+'_MW','Capacity',P==='W'?'MW':'MWp',F.int],[P+'_CAPEX','Capex','EURm/MW','0.000'],[P+'_GCF','Gross capacity factor','%',F.pct2],
   [P+'_LOSS','Plant losses','%',F.pct1],[P+'_LLOSS','Direct-line losses','%',F.pct1],[P+'_DEGR','Degradation','per yr',F.pct2],
   [P+'_OPEX','Opex','EURm/MW/yr','0.000'],[P+'_PPA','PPA price, manual override','EUR/MWh',F.price],[P+'_CONTR','Contracted share','%',F.pct0],
   [P+'_COD','COD (first generation)','year',F.yr],[P+'_LIFE','Useful life','years',F.yr],[P+'_TAIL','Merchant tail price','EUR/MWh',F.price],
   ['T_ON','Tranche pricing on','1/0',F.yr],[P==='W'?'T1W':'T1S','Tranche 1 size',P==='W'?'MW':'MWp',F.int],
   ['TP1','Tranche 1 price','EUR/MWh',F.price],['TP2','Tranche 2 price','EUR/MWh',F.price],
   ['SUBR','Funding coupon','per yr',F.pct1],['SUBBE','Burgenland share while outstanding','%',F.pct0],
   ['SUBSTEP','Coupon step-up','per yr',F.pct1],['SUBSTEPY','Step-up after','op. years',F.yr]
  ].forEach(x=>{rr=lnk(ws,rr,D,x[0]==='T1W'||x[0]==='T1S'?'T1':x[0],x[1],x[0],x[3],x[2]);});
  const A=k=>`$F$${D[k]}`;

  rr+=2; rr=sect(ws,rr,2,'DERIVED (computed once)');
  rr=derRow(ws,rr,D,'T1MW','Tranche 1 capacity',`MIN(${A('T1')},${A(P+'_MW')})`,F.int,P==='W'?'MW':'MWp');
  rr=derRow(ws,rr,D,'T2MW','Tranche 2 capacity',`MAX(0,${A(P+'_MW')}-${A('T1')})`,F.int,P==='W'?'MW':'MWp');
  rr=derRow(ws,rr,D,'BLEND','Blended tranche price',`($F$${D.T1MW}*${A('TP1')}+$F$${D.T2MW}*${A('TP2')})/MAX(0.0001,${A(P+'_MW')})`,F.price,'EUR/MWh');
  rr=derRow(ws,rr,D,'EFFPPA','PPA price in force',`${A('T_ON')}*$F$${D.BLEND}+(1-${A('T_ON')})*${A(P+'_PPA')}`,F.price,'EUR/MWh');
  rr=derRow(ws,rr,D,'GROSS','Gross generation at 100% availability',`${A(P+'_MW')}*8760*${A(P+'_GCF')}`,F.int,'MWh');
  rr=derRow(ws,rr,D,'NETFAC','Net delivery factor',`(1-${A(P+'_LOSS')})*(1-${A(P+'_LLOSS')})`,F.fac,'x');
  rr=derRow(ws,rr,D,'TCAPEX','Total capex',`${A(P+'_MW')}*${A(P+'_CAPEX')}`,F.eur,'EURm');
  rr=derRow(ws,rr,D,'EQUITY','Construction funding (junior leg)',`(1-${A('GEAR')})*$F$${D.TCAPEX}`,F.eur,'EURm');
  rr=derRow(ws,rr,D,'DEPY','Depreciation period',`MIN(20,${A(P+'_LIFE')})`,F.yr,'years');
  rr=derRow(ws,rr,D,'ANNDEP','Annual depreciation',`$F$${D.TCAPEX}/$F$${D.DEPY}`,F.num,'EURm');
  rr=derRow(ws,rr,D,'REPY','Debt amortisation period',`MIN(${A('TENOR')},${A(P+'_LIFE')})`,F.yr,'years');
  rr=derRow(ws,rr,D,'PPAEND','PPA end year',`${A(P+'_COD')}+MIN(${A('PPAT')},${A(P+'_LIFE')})`,F.yr,'year');
  rr=derRow(ws,rr,D,'DRAW1','Senior draw at COD-2',`0.3*$F$${D.TCAPEX}*${A('GEAR')}`,F.num,'EURm');
  rr=derRow(ws,rr,D,'IDC1','Construction interest, year 1',`$F$${D.DRAW1}/2*${A('RATE')}`,F.num,'EURm');
  rr=derRow(ws,rr,D,'BAL1','Senior debt after COD-2',`$F$${D.DRAW1}+$F$${D.IDC1}`,F.num,'EURm');
  rr=derRow(ws,rr,D,'DRAW2','Senior draw at COD-1',`0.7*$F$${D.TCAPEX}*${A('GEAR')}`,F.num,'EURm');
  rr=derRow(ws,rr,D,'IDC2','Construction interest, year 2',`($F$${D.BAL1}+$F$${D.DRAW2}/2)*${A('RATE')}`,F.num,'EURm');
  rr=derRow(ws,rr,D,'DEBTCOD','Senior debt drawn by COD',`$F$${D.BAL1}+$F$${D.DRAW2}+$F$${D.IDC2}`,F.num,'EURm');
  rr=derRow(ws,rr,D,'ANNPRIN','Annual principal, flat',`$F$${D.DEBTCOD}/$F$${D.REPY}`,F.num,'EURm');
  rr=derRow(ws,rr,D,'ANNDS','Annual debt service, annuity',`PMT(${A('RATE')},$F$${D.REPY},-$F$${D.DEBTCOD})`,F.num,'EURm');

  rr+=2; rr=sect(ws,rr,3,'PRODUCTION');
  R.op  =tsRow(ws,rr++,'Operating flag','1/0',F.yr,(X,p,a)=>`(${yr(X)}>=${a[0]})*(${yr(X)}<${a[0]}+${a[1]})`,{aud:[{in:P+'_COD',fmt:F.yr},{in:P+'_LIFE',fmt:F.yr}]});
  R.ysc =tsRow(ws,rr++,'Years since COD','years',F.yr,(X,p,a)=>`${yr(X)}-${a[0]}`,{aud:[{in:P+'_COD',fmt:F.yr}]});
  R.cpi =tsRow(ws,rr++,'CPI index vs 2023','x',F.fac,(X,p,a)=>`(1+${a[0]})^(${yr(X)}-2023)`,{aud:[{in:'INFL',fmt:F.pct1}]});
  R.degf=tsRow(ws,rr++,'Degradation factor','x',F.fac,(X,p,a)=>`(1-${a[0]})^MAX(0,${X}$${R.ysc})`,{aud:[{in:P+'_DEGR',fmt:F.pct2}]});
  R.prod=tsRow(ws,rr++,'Production','MWh',F.int,(X,p,a)=>`${X}$${R.op}*${a[0]}*${a[1]}*${X}$${R.degf}`,{aud:[{f:A('GROSS'),fmt:F.int},{f:A('NETFAC'),fmt:F.fac}]});

  rr+=2; rr=sect(ws,rr,4,'REVENUE (tranches, merchant and tail)');
  R.ppaf=tsRow(ws,rr++,'In-PPA-period flag','1/0',F.yr,(X,p,a)=>`(${yr(X)}<${a[0]})`,{aud:[{f:A('PPAEND'),fmt:F.yr}]});
  R.cmwh=tsRow(ws,rr++,'Contracted volume','MWh',F.int,(X,p,a)=>`${X}$${R.op}*${X}$${R.ppaf}*${X}$${R.prod}*${a[0]}`,{aud:[{in:P+'_CONTR',fmt:F.pct0}]});
  R.t1sh=tsRow(ws,rr++,'Tranche 1 volume','MWh',F.int,(X,p,a)=>`${X}$${R.cmwh}*${a[0]}/MAX(0.0001,${a[1]})`,{aud:[{f:A('T1MW'),fmt:F.int},{in:P+'_MW',fmt:F.int}]});
  R.t2sh=tsRow(ws,rr++,'Tranche 2 volume','MWh',F.int,X=>`${X}$${R.cmwh}-${X}$${R.t1sh}`);
  R.rt1 =tsRow(ws,rr++,'Revenue: tranche 1','EURm',F.num,(X,p,a)=>`${a[2]}*${X}$${R.t1sh}*${a[0]}/10^6`,{aud:[{in:'TP1',fmt:F.price},null,{in:'T_ON',fmt:F.yr}]});
  R.rt2 =tsRow(ws,rr++,'Revenue: tranche 2','EURm',F.num,(X,p,a)=>`${a[2]}*${X}$${R.t2sh}*${a[0]}/10^6`,{aud:[{in:'TP2',fmt:F.price},null,{in:'T_ON',fmt:F.yr}]});
  R.rman=tsRow(ws,rr++,'Revenue: manual PPA (tranches off)','EURm',F.num,(X,p,a)=>`(1-${a[2]})*${X}$${R.cmwh}*${a[0]}/10^6`,{aud:[{in:P+'_PPA',fmt:F.price},null,{in:'T_ON',fmt:F.yr}]});
  R.rppa=tsRow(ws,rr++,'Revenue: contracted','EURm',F.num,X=>`${X}$${R.rt1}+${X}$${R.rt2}+${X}$${R.rman}`,{subtotal:true});
  R.mmwh=tsRow(ws,rr++,'Merchant volume (in PPA period)','MWh',F.int,(X,p,a)=>`${X}$${R.op}*${X}$${R.ppaf}*${X}$${R.prod}*(1-${a[0]})`,{aud:[{in:P+'_CONTR',fmt:F.pct0}]});
  R.tmwh=tsRow(ws,rr++,'Tail volume (post-PPA)','MWh',F.int,X=>`${X}$${R.op}*(1-${X}$${R.ppaf})*${X}$${R.prod}`);
  R.mpr =tsRow(ws,rr++,'Merchant price','EUR/MWh',F.price,(X,p,a)=>`${a[0]}*${X}$${R.cpi}`,{aud:[{in:'MERCH',fmt:F.price}]});
  R.rmer=tsRow(ws,rr++,'Revenue: merchant','EURm',F.num,X=>`${X}$${R.mmwh}*${X}$${R.mpr}/10^6`);
  R.rtail=tsRow(ws,rr++,'Revenue: tail','EURm',F.num,(X,p,a)=>`${X}$${R.tmwh}*${a[0]}/10^6`,{aud:[{in:P+'_TAIL',fmt:F.price}]});
  R.rev =tsRow(ws,rr++,'Revenue: total','EURm',F.num,X=>`${X}$${R.rppa}+${X}$${R.rmer}+${X}$${R.rtail}`,{subtotal:true});

  rr+=2; rr=sect(ws,rr,5,'COSTS & EBITDA');
  R.opex=tsRow(ws,rr++,'Opex','EURm',F.num,(X,p,a)=>`${X}$${R.op}*${a[0]}*${a[1]}*${X}$${R.cpi}`,{aud:[{in:P+'_OPEX',fmt:'0.000'},{in:P+'_MW',fmt:F.int}]});
  R.ebit=tsRow(ws,rr++,'EBITDA','EURm',F.num,X=>`${X}$${R.rev}-${X}$${R.opex}`,{subtotal:true});

  rr+=2; rr=sect(ws,rr,6,'CAPEX & SENIOR DEBT');
  R.cshr=tsRow(ws,rr++,'Capex draw share','x','0.00',(X,p,a)=>`0.3*(${yr(X)}=${a[0]}-2)+0.7*(${yr(X)}=${a[0]}-1)`,{aud:[{in:P+'_COD',fmt:F.yr}]});
  R.capex=tsRow(ws,rr++,'Capex','EURm',F.num,(X,p,a)=>`${X}$${R.cshr}*${a[0]}`,{aud:[{f:A('TCAPEX'),fmt:F.eur}]});
  R.draw=tsRow(ws,rr++,'Senior debt draw','EURm',F.num,(X,p,a)=>`${X}$${R.capex}*${a[0]}`,{aud:[{in:'GEAR',fmt:F.pct0}]});
  R.cflag=tsRow(ws,rr++,'Construction flag','1/0',F.yr,(X,p,a)=>`(${yr(X)}>=${a[0]}-2)*(${yr(X)}<${a[0]})`,{aud:[{in:P+'_COD',fmt:F.yr}]});
  R.idc=rr++; R.intr=rr++; R.prin=rr++; R.bal=rr++; R.dep=rr++; R.ebt=rr++; R.nol=rr++; R.tax=rr++; R.fcfe=rr++; R.date=rr++; R.chk=rr++; R.diff=rr++; R.xcf=rr++;
  tsRow(ws,R.idc,'Construction interest (IDC)','EURm',F.num,(X,pX,a)=>`${X}$${R.cflag}*((${pX?pX+'$'+R.bal:'0'})+${X}$${R.draw}/2)*${a[0]}`,{aud:[{in:'RATE',fmt:F.pct2}]});
  tsRow(ws,R.intr,'Senior interest','EURm',F.num,(X,pX,a)=>`(${yr(X)}>=${a[1]})*(${pX?pX+'$'+R.bal:'0'})*${a[0]}`,{aud:[{in:'RATE',fmt:F.pct2},{in:P+'_COD',fmt:F.yr}]});
  tsRow(ws,R.prin,'Senior principal repaid','EURm',F.num,(X,pX,a)=>`(${yr(X)}>=${a[1]})*(${yr(X)}<${a[1]}+$F$${D.REPY})*MIN((${pX?pX+'$'+R.bal:'0'}),${a[0]}*MAX(0,$F$${D.ANNDS}-${X}$${R.intr})+(1-${a[0]})*$F$${D.ANNPRIN})`,{aud:[{in:'AMORT',fmt:F.yr},{in:P+'_COD',fmt:F.yr}]});
  tsRow(ws,R.bal,'Senior debt balance, end of year','EURm',F.num,(X,pX,a)=>`(${yr(X)}<${a[0]})*((${pX?pX+'$'+R.bal:'0'})+${X}$${R.draw}+${X}$${R.idc})+(${yr(X)}>=${a[0]})*MAX(0,(${pX?pX+'$'+R.bal:'0'})-${X}$${R.prin})`,{aud:[{in:P+'_COD',fmt:F.yr}]});
  tsRow(ws,R.dep,'Depreciation','EURm',F.num,(X,p,a)=>`(${yr(X)}>=${a[0]})*(${yr(X)}<${a[0]}+$F$${D.DEPY})*$F$${D.ANNDEP}`,{aud:[{in:P+'_COD',fmt:F.yr},{f:A('ANNDEP'),fmt:F.num}]});
  tsRow(ws,R.ebt,'EBT','EURm',F.num,X=>`${X}$${R.ebit}-${X}$${R.dep}-${X}$${R.intr}`);
  tsRow(ws,R.nol,'Tax-loss balance','EURm',F.num,(X,pX,a)=>`(${yr(X)}>=${a[0]})*MIN(0,(${pX?pX+'$'+R.nol:'0'})+${X}$${R.ebt})`,{aud:[{in:P+'_COD',fmt:F.yr}]});
  tsRow(ws,R.tax,'Tax','EURm',F.num,(X,pX,a)=>`(${yr(X)}>=${a[1]})*MAX(0,${X}$${R.ebt}+(${pX?pX+'$'+R.nol:'0'}))*${a[0]}`,{aud:[{in:'TAXR',fmt:F.pct0},{in:P+'_COD',fmt:F.yr}]});
  tsRow(ws,R.fcfe,'Cash after senior debt service','EURm',F.num,(X,p,a)=>`(${yr(X)}<${a[0]})*(-${X}$${R.capex}+${X}$${R.draw})+(${yr(X)}>=${a[0]})*(${X}$${R.ebit}-${X}$${R.tax}-${X}$${R.prin}-${X}$${R.intr})`,{subtotal:true,aud:[{in:P+'_COD',fmt:F.yr}]});
  ws.getCell(R.date,3).value='Date (for XIRR)';
  for(let i=0;i<NY;i++){const c=ws.getCell(R.date,c0+i);c.value={formula:`DATE(${colL(c0+i)}$${YROW},12,31)`};c.numFmt=F.date;}
  ws.getCell(R.chk,3).value='Check: dashboard cash flow';
  for(let i=0;i<NY;i++){const y=Y0+i;const c=ws.getCell(R.chk,c0+i);c.value=(chk&&chk[y]!==undefined?chk[y]:0);c.fill=CHK;c.numFmt=F.num;}
  tsRow(ws,R.diff,'Check difference (should be 0)','EURm',F.num,X=>`${X}$${R.fcfe}-${X}$${R.chk}`);
  ws.getCell(R.xcf,3).value='Equity CF for XIRR (seed at t0, immaterial)';
  ws.getCell(R.xcf,c0).value=-0.01; ws.getCell(R.xcf,c0).numFmt=F.num;
  for(let i=1;i<NY;i++){const c=ws.getCell(R.xcf,c0+i);c.value={formula:`${colL(c0+i)}${R.fcfe}`};c.numFmt=F.num;}

  rr=R.xcf+3; rr=sect(ws,rr,7,'FINANCING WATERFALL (senior first, then the split)');
  ws.getCell(rr,3).value='The junior 30% is construction funding advanced by the Energy SPV. It accrues the coupon from drawdown; unpaid coupon capitalises. Each year the cash after senior service splits: the sweep share repays the funding, the rest is Burgenland Energie\'s. Once repaid, everything is Burgenland Energie\'s.';
  ws.getCell(rr,3).font={italic:true,size:9,color:{argb:'FF808080'}}; rr++;
  R.inj =tsRow(ws,rr++,'Funding drawn (junior leg)','EURm',F.num,X=>`${X}$${R.capex}-${X}$${R.draw}`);
  R.srat=tsRow(ws,rr++,'Coupon in force','per yr',F.pct2,(X,p,a)=>`${a[0]}+${a[1]}*(${yr(X)}>=${a[2]}+$F$${D.SUBSTEPY})`,{aud:[{in:'SUBR',fmt:F.pct1},{in:'SUBSTEP',fmt:F.pct1},{in:P+'_COD',fmt:F.yr}]});
  R.wacc=rr++; R.wowed=rr++; R.wdist=rr++; R.wsub=rr++; R.wbe=rr++; R.wbal=rr++; R.wbec=rr++; R.wscf=rr++; R.wchkb=rr++; R.wdifb=rr++; R.wchke=rr++; R.wdife=rr++;
  tsRow(ws,R.wacc,'Coupon accrued','EURm',F.num,(X,pX)=>`((${pX?pX+'$'+R.wbal:'0'})+${X}$${R.inj}/2)*${X}$${R.srat}`);
  tsRow(ws,R.wowed,'Funding owed before payment','EURm',F.num,(X,pX)=>`(${pX?pX+'$'+R.wbal:'0'})+${X}$${R.inj}+${X}$${R.wacc}`);
  tsRow(ws,R.wdist,'Cash available to split','EURm',F.num,(X,p,a)=>`(${yr(X)}>=${a[0]})*MAX(0,${X}$${R.fcfe})`,{aud:[{in:P+'_COD',fmt:F.yr}]});
  tsRow(ws,R.wsub,'Paid to Energy SPV funding','EURm',F.num,(X,p,a)=>`MIN(${X}$${R.wowed},MAX(0,${X}$${R.wdist}*(1-${a[0]})))`,{subtotal:true,aud:[{in:'SUBBE',fmt:F.pct0}]});
  tsRow(ws,R.wbe,'Paid to Burgenland Energie','EURm',F.num,X=>`${X}$${R.wdist}-${X}$${R.wsub}`,{subtotal:true});
  tsRow(ws,R.wbal,'Funding balance, end of year','EURm',F.num,X=>`${X}$${R.wowed}-${X}$${R.wsub}`);
  tsRow(ws,R.wbec,'Burgenland cumulative cash','EURm',F.num,(X,pX)=>`(${pX?pX+'$'+R.wbec:'0'})+${X}$${R.wbe}`);
  tsRow(ws,R.wscf,'Funding cash flow (for XIRR)','EURm',F.num,X=>`${X}$${R.wsub}-${X}$${R.inj}`);
  ws.getCell(R.wchkb,3).value='Check: dashboard funding balance';
  ws.getCell(R.wchke,3).value='Check: dashboard Burgenland cash';
  for(let i=0;i<NY;i++){const y=Y0+i;
   const b=ws.getCell(R.wchkb,c0+i); b.value=(chkWf&&chkWf.bal[y]!==undefined?chkWf.bal[y]:0); b.fill=CHK; b.numFmt=F.num;
   const e=ws.getCell(R.wchke,c0+i); e.value=(chkWf&&chkWf.be[y]!==undefined?chkWf.be[y]:0); e.fill=CHK; e.numFmt=F.num;}
  tsRow(ws,R.wdifb,'Check difference, balance (0)','EURm',F.num,X=>`${X}$${R.wbal}-${X}$${R.wchkb}`);
  tsRow(ws,R.wdife,'Check difference, Burgenland (0)','EURm',F.num,X=>`${X}$${R.wbe}-${X}$${R.wchke}`);

  rr=R.wdife+3; rr=sect(ws,rr,8,'LCOE, STEP BY STEP');
  R.dfac=tsRow(ws,rr++,'Discount factor to COD','x',F.fac,(X,p,a)=>`(1+${a[0]})^-(${yr(X)}-${a[1]})`,{aud:[{in:'RATE',fmt:F.pct2},{in:P+'_COD',fmt:F.yr}]});
  R.pvc =tsRow(ws,rr++,'PV of capex and opex','EURm',F.num,X=>`(${X}$${R.capex}+${X}$${R.opex})*${X}$${R.dfac}`);
  R.pve =tsRow(ws,rr++,'PV of production','MWh',F.int,X=>`${X}$${R.prod}*${X}$${R.dfac}`);

  rr+=2; rr=sect(ws,rr,9,'RESULTS');
  const put=(row,label,formula,fmtStr,unit,fill)=>{ws.getCell(row,3).value=label;
   if(unit){const u=ws.getCell(row,5);u.value=unit;u.font={size:9.5,color:{argb:'FF7A7A7A'}};}
   const c=ws.getCell(row,6);c.value={formula:formula};c.numFmt=fmtStr;if(fill)c.fill=fill;return row;};
  let k=rr;
  const IRR=k; put(k++,'Equity IRR (XIRR, junior leg unsplit)',`XIRR(${colL(c0)}${R.xcf}:${lastC}${R.xcf},${colL(c0)}${R.date}:${lastC}${R.date})`,F.pct2,'per yr'); ws.getCell(IRR,6).font=bold;
  const MOIC=k; put(k++,'MOIC',`SUMIF(${colL(c0)}${R.fcfe}:${lastC}${R.fcfe},">0")/-SUMIF(${colL(c0)}${R.fcfe}:${lastC}${R.fcfe},"<0")`,F.x,'x');
  put(k++,'Total capex',`$F$${D.TCAPEX}`,F.eur,'EURm');
  put(k++,'Senior debt at drawdown',`$F$${D.DEBTCOD}`,F.eur,'EURm');
  put(k++,'Construction funding advanced',`SUMIF(${colL(c0)}${R.inj}:${lastC}${R.inj},">0")`,F.eur,'EURm');
  const PVC=k; put(k++,'PV of costs',`SUM(${colL(c0)}${R.pvc}:${lastC}${R.pvc})`,F.eur,'EURm');
  const PVE=k; put(k++,'PV of production',`SUM(${colL(c0)}${R.pve}:${lastC}${R.pve})`,F.int,'MWh');
  const LC=k; put(k++,'LCOE = PV costs / PV production',`$F$${PVC}*10^6/MAX(1,$F$${PVE})`,F.price,'EUR/MWh');
  const BL=k; put(k++,'Blended PPA in force',`$F$${D.EFFPPA}`,F.price,'EUR/MWh');
  k++;
  const SIRR=k; put(k++,'Funding IRR (XIRR)',`IFERROR(XIRR(${colL(c0)}${R.wscf}:${lastC}${R.wscf},${colL(c0)}${R.date}:${lastC}${R.date}),"n/m")`,F.pct2,'per yr'); ws.getCell(SIRR,6).font=bold;
  const SMOIC=k; put(k++,'Funding MOIC',`SUMIF(${colL(c0)}${R.wscf}:${lastC}${R.wscf},">0")/MAX(0.0001,-SUMIF(${colL(c0)}${R.wscf}:${lastC}${R.wscf},"<0"))`,F.x,'x');
  const PAYY=k; put(k++,'Funding repaid in year',`SUMPRODUCT(MAX((${colL(c0)}${R.wbal}:${lastC}${R.wbal}>0.0001)*(${colL(c0)}${YROW}:${lastC}${YROW})))+1`,F.yr,'year');
  const PAYN=k; put(k++,'Payback, years of operation',`$F$${PAYY}-${A(P+'_COD')}+1`,F.yr,'years');
  const BE10=k; put(k++,'Burgenland cash, first 10 operating years',`SUMPRODUCT((${colL(c0)}${YROW}:${lastC}${YROW}>=${A(P+'_COD')})*(${colL(c0)}${YROW}:${lastC}${YROW}<${A(P+'_COD')}+10)*(${colL(c0)}${R.wbe}:${lastC}${R.wbe}))`,F.eur,'EURm');
  const BETOT=k; put(k++,'Burgenland cash, total',`SUM(${colL(c0)}${R.wbe}:${lastC}${R.wbe})`,F.eur,'EURm');
  k++;
  const MX=k; put(k++,'Max check difference, cash flow',`MAX(ABS(MIN(${colL(c0)}${R.diff}:${lastC}${R.diff})),ABS(MAX(${colL(c0)}${R.diff}:${lastC}${R.diff})))`,F.num,'EURm',CHK);
  const MXW=k; put(k++,'Max check difference, waterfall',`MAX(ABS(MIN(${colL(c0)}${R.wdifb}:${lastC}${R.wdifb})),ABS(MAX(${colL(c0)}${R.wdifb}:${lastC}${R.wdifb})),ABS(MIN(${colL(c0)}${R.wdife}:${lastC}${R.wdife})),ABS(MAX(${colL(c0)}${R.wdife}:${lastC}${R.wdife})))`,F.num,'EURm',CHK);
  put(k++,'Tie-out',`CHOOSE(1+(MAX($F$${MX},$F$${MXW})>0.001),"OK ties to dashboard","CHECK differences")`,'General');
  return {irr:`${name}!$F$${IRR}`,moic:`${name}!$F$${MOIC}`,lcoe:`${name}!$F$${LC}`,blend:`${name}!$F$${BL}`,
    subIRR:`${name}!$F$${SIRR}`,subMOIC:`${name}!$F$${SMOIC}`,payY:`${name}!$F$${PAYY}`,payN:`${name}!$F$${PAYN}`,
    be10:`${name}!$F$${BE10}`,beTot:`${name}!$F$${BETOT}`,
    tcapex:`${name}!$F$${D.TCAPEX}`,debt:`${name}!$F$${D.DEBTCOD}`,equity:`${name}!$F$${D.EQUITY}`,
    prodRow:R.prod,lcoeCell:`${name}!$F$${LC}`};
 }
 const windRef =assetSheet('Wind','W',S.chkWind,S.chkWfWind);
 const solarRef=assetSheet('Solar','S',S.chkSolar,S.chkWfSolar);

 // ================= BATTERY =================
 function batterySheet(){
  const ws=calcSheet('Battery','Battery | arbitrage from the day-by-day backtest summary; the 24h curve is a reference');
  const D={}, R={};
  let rr=4; rr=sect(ws,rr,1,'ASSUMPTIONS (linked from Inputs)');
  [['INFL','Inflation','per yr',F.pct1],['TAXR','Tax rate','%',F.pct0],['TENOR','Debt tenor','years',F.yr],['AMORT','Amortisation (1 = annuity)','1/0',F.yr],['FEESC','Grid-fee escalation','per yr',F.pct1],
   ['B_MW','Power','MW',F.int],['B_DUR','Duration','hours',F.yr],['B_CKWH','Cell capex','EUR/kWh',F.int],['B_SUB','Grid-interface allowance','EURm',F.eur],['B_INT','Campus connection allowance','EURm',F.eur],
   ['B_RTE','Round-trip efficiency','%',F.pct0],['B_SOC','SoC floor','%',F.pct0],['B_CYC','Cycles per day','x',F.yr],['B_CAPF','Capture factor','x','0.00'],
   ['B_OPEXP','Opex','% of cell capex',F.pct1],['B_ANC','Ancillary revenue','EURk/MW/yr',F.int],['B_CCH','DC reliability charge','EURk/MW/yr',F.int],
   ['B_GY','First merchant year','year',F.yr],['B_COMP','Revenue compression','per yr',F.pct1],['B_DEGR','Degradation','per yr',F.pct1],
   ['B_GEAR','Gearing','%',F.pct0],['B_RATE','Debt rate','per yr',F.pct2],['B_GFC','Grid fee capacity, manual','EUR/kW/yr',F.eur],['B_MKT','Apply DC NE3 tariff','1/0',F.yr],['FEE_C','DC NE3 capacity tariff','EUR/kW/yr',F.price],['B_GFE','Grid fee energy','EUR/MWh',F.eur],
   ['B_AVGBUY','Backtest average buy','EUR/MWh',F.price],['B_AVGSELL','Backtest average sell','EUR/MWh',F.price],['B_FSELF','Self-charge share','%',F.pct0],['B_BTM','Self-charge cost','EUR/MWh',F.price],
   ['B_COD','COD (capex year)','year',F.yr],['B_LIFE','Useful life','years',F.yr]
  ].forEach(x=>{rr=lnk(ws,rr,D,x[0],x[1],x[0],x[3],x[2]);});
  const A=k=>`$F$${D[k]}`;

  rr+=2; rr=sect(ws,rr,2,'REFERENCE 24h CURVE (not used in the financial formulas)');
  const CST=rr;
  for(let h=0;h<24;h++){
   ws.getCell(rr,3).value='Hour '+h; ws.getCell(rr,5).value='EUR/MWh'; ws.getCell(rr,5).font={size:9.5,color:{argb:'FF7A7A7A'}};
   const p=ws.getCell(rr,6); p.value={formula:`Inputs!$F$${IN.PH0+h}`}; p.font=GRN; p.numFmt=F.price;
   rr++;
  }
  const CEN=rr-1;
  for(let h=0;h<24;h++){const rk=ws.getCell(CST+h,8); rk.value={formula:`RANK($F$${CST+h},$F$${CST}:$F$${CEN},1)`}; rk.numFmt=F.int;}
  ws.getCell(CST-1,8).value='rank (1 = cheapest)'; ws.getCell(CST-1,8).font={size:9.5,color:{argb:'FF7A7A7A'}};

  rr+=2; rr=sect(ws,rr,3,'DERIVED DISPATCH & CAPEX (computed once)');
  rr=derRow(ws,rr,D,'NCH','Tradeable hours per day',`MIN(12,MAX(0,${A('B_DUR')}*(1-${A('B_SOC')})*${A('B_CYC')}))`,F.num,'hours');
  rr=derRow(ws,rr,D,'SCHEAP','Daily charge cost',`$F$${D.NCH}*(${A('B_AVGBUY')}*(1-${A('B_FSELF')})+${A('B_FSELF')}*${A('B_BTM')})`,F.num,'EUR/MW');
  rr=derRow(ws,rr,D,'SPRICE','Daily sale value before RTE',`$F$${D.NCH}*${A('B_AVGSELL')}`,F.num,'EUR/MW');
  rr=derRow(ws,rr,D,'DAYARB','Day arbitrage',`MAX(0,${A('B_MW')}*(${A('B_RTE')}*$F$${D.SPRICE}-$F$${D.SCHEAP}))*${A('B_CAPF')}`,F.num,'EUR/day');
  rr=derRow(ws,rr,D,'ARBRR','Arbitrage revenue run-rate',`$F$${D.DAYARB}*365/10^6`,F.num,'EURm/yr');
  rr=derRow(ws,rr,D,'ANC','Ancillary revenue',`${A('B_ANC')}*MIN(${A('B_MW')},225)/1000`,F.num,'EURm/yr');
  rr=derRow(ws,rr,D,'DCCH','DC reliability charge',`${A('B_CCH')}*${A('B_MW')}/1000`,F.num,'EURm/yr');
  rr=derRow(ws,rr,D,'GFCAP','Grid fee, capacity (2028 level)',`${A('B_MW')}*MAX(${A('B_GFC')},${A('B_MKT')}*${A('FEE_C')})/1000`,F.num,'EURm/yr');
  rr=derRow(ws,rr,D,'THRU','Grid throughput',`$F$${D.NCH}*${A('B_MW')}*${A('B_RTE')}*365`,F.int,'MWh/yr');
  rr=derRow(ws,rr,D,'GFENE','Grid fee, energy (2028 level)',`$F$${D.THRU}*${A('B_GFE')}/10^6`,F.num,'EURm/yr');
  rr=derRow(ws,rr,D,'CELLCX','Cell capex',`${A('B_MW')}*${A('B_DUR')}*1000*${A('B_CKWH')}/10^6`,F.eur,'EURm');
  rr=derRow(ws,rr,D,'TCAPEX','Total capex',`$F$${D.CELLCX}+${A('B_SUB')}+${A('B_INT')}`,F.eur,'EURm');
  rr=derRow(ws,rr,D,'EQUITY','Equity',`(1-${A('B_GEAR')})*$F$${D.TCAPEX}`,F.eur,'EURm');
  rr=derRow(ws,rr,D,'DEPY','Depreciation period',`MIN(15,${A('B_LIFE')}-1)`,F.yr,'years');
  rr=derRow(ws,rr,D,'ANNDEP','Annual depreciation',`$F$${D.TCAPEX}/$F$${D.DEPY}`,F.num,'EURm');
  rr=derRow(ws,rr,D,'REPY','Debt amortisation period',`MIN(${A('TENOR')},${A('B_LIFE')}-1)`,F.yr,'years');
  rr=derRow(ws,rr,D,'ANNPRIN','Annual principal, flat',`${A('B_GEAR')}*$F$${D.TCAPEX}/$F$${D.REPY}`,F.num,'EURm');
  rr=derRow(ws,rr,D,'ANNDS','Annual debt service, annuity',`PMT(${A('B_RATE')},$F$${D.REPY},-${A('B_GEAR')}*$F$${D.TCAPEX})`,F.num,'EURm');

  rr+=2; rr=sect(ws,rr,4,'YEARLY MODEL');
  R.mflag=tsRow(ws,rr++,'Merchant flag','1/0',F.yr,(X,p,a)=>`(${yr(X)}>=${a[0]})`,{aud:[{in:'B_GY',fmt:F.yr}]});
  R.olf =tsRow(ws,rr++,'Operating flag','1/0',F.yr,(X,p,a)=>`(${yr(X)}>${a[0]})*(${yr(X)}<${a[0]}+${a[1]})`,{aud:[{in:'B_COD',fmt:F.yr},{in:'B_LIFE',fmt:F.yr}]});
  R.cpi =tsRow(ws,rr++,'CPI index vs 2023','x',F.fac,(X,p,a)=>`(1+${a[0]})^(${yr(X)}-2023)`,{aud:[{in:'INFL',fmt:F.pct1}]});
  R.merch=tsRow(ws,rr++,'Merchant revenue','EURm',F.num,(X,p,a)=>`${X}$${R.olf}*${X}$${R.mflag}*($F$${D.ARBRR}*(1-${a[0]})^MAX(0,${yr(X)}-${a[2]})+$F$${D.ANC})*(1-${a[1]})^MAX(0,${yr(X)}-${a[2]})*(1+${A('INFL')})^MAX(0,${yr(X)}-${a[2]})`,{aud:[{in:'B_DEGR',fmt:F.pct1},{in:'B_COMP',fmt:F.pct1},{in:'B_GY',fmt:F.yr}]});
  R.cap =tsRow(ws,rr++,'DC reliability charge','EURm',F.num,(X,p,a)=>`${X}$${R.olf}*$F$${D.DCCH}*(1+${A('INFL')})^MAX(0,${yr(X)}-(${a[0]}+1))`,{aud:[{in:'B_COD',fmt:F.yr}]});
  R.rev =tsRow(ws,rr++,'Revenue: total','EURm',F.num,X=>`${X}$${R.merch}+${X}$${R.cap}`,{subtotal:true});
  R.opxc=tsRow(ws,rr++,'Opex: cell O&M','EURm',F.num,(X,p,a)=>`${X}$${R.olf}*$F$${D.CELLCX}*${a[0]}*${X}$${R.cpi}`,{aud:[{in:'B_OPEXP',fmt:F.pct1}]});
  R.gfee=tsRow(ws,rr++,'Grid fees (from merchant year)','EURm',F.num,(X,p,a)=>`${X}$${R.olf}*${X}$${R.mflag}*($F$${D.GFCAP}+$F$${D.GFENE})*(1+${a[0]})^MAX(0,${yr(X)}-2028)`,{aud:[{in:'FEESC',fmt:F.pct1}]});
  R.opex=tsRow(ws,rr++,'Opex: total','EURm',F.num,X=>`${X}$${R.opxc}+${X}$${R.gfee}`,{subtotal:true});
  R.codf=tsRow(ws,rr++,'COD flag','1/0',F.yr,(X,p,a)=>`(${yr(X)}=${a[0]})`,{aud:[{in:'B_COD',fmt:F.yr}]});
  R.intr=rr++; R.prin=rr++; R.bal=rr++; R.dep=rr++; R.ebt=rr++; R.nol=rr++; R.tax=rr++; R.fcfe=rr++; R.date=rr++; R.chk=rr++; R.diff=rr++; R.xcf=rr++;
  tsRow(ws,R.intr,'Interest','EURm',F.num,(X,pX,a)=>`(${yr(X)}>${a[1]})*(${pX?pX+'$'+R.bal:'0'})*${a[0]}`,{aud:[{in:'B_RATE',fmt:F.pct2},{in:'B_COD',fmt:F.yr}]});
  tsRow(ws,R.prin,'Principal repaid','EURm',F.num,(X,pX,a)=>`(${yr(X)}>${a[1]})*(${yr(X)}<=${a[1]}+$F$${D.REPY})*MIN((${pX?pX+'$'+R.bal:'0'}),${a[0]}*MAX(0,$F$${D.ANNDS}-${X}$${R.intr})+(1-${a[0]})*$F$${D.ANNPRIN})`,{aud:[{in:'AMORT',fmt:F.yr},{in:'B_COD',fmt:F.yr}]});
  tsRow(ws,R.bal,'Debt balance, end of year','EURm',F.num,(X,pX,a)=>`${X}$${R.codf}*${a[0]}*$F$${D.TCAPEX}+(${yr(X)}>${a[1]})*MAX(0,(${pX?pX+'$'+R.bal:'0'})-${X}$${R.prin})`,{aud:[{in:'B_GEAR',fmt:F.pct0},{in:'B_COD',fmt:F.yr}]});
  tsRow(ws,R.dep,'Depreciation','EURm',F.num,(X,p,a)=>`(${yr(X)}>${a[0]})*(${yr(X)}<=${a[0]}+$F$${D.DEPY})*$F$${D.ANNDEP}`,{aud:[{in:'B_COD',fmt:F.yr}]});
  tsRow(ws,R.ebt,'EBT','EURm',F.num,X=>`${X}$${R.rev}-${X}$${R.opex}-${X}$${R.dep}-${X}$${R.intr}`);
  tsRow(ws,R.nol,'Tax-loss balance','EURm',F.num,(X,pX,a)=>`(${yr(X)}>${a[0]})*MIN(0,(${pX?pX+'$'+R.nol:'0'})+${X}$${R.ebt})`,{aud:[{in:'B_COD',fmt:F.yr}]});
  tsRow(ws,R.tax,'Tax','EURm',F.num,(X,pX,a)=>`(${yr(X)}>${a[1]})*MAX(0,${X}$${R.ebt}+(${pX?pX+'$'+R.nol:'0'}))*${a[0]}`,{aud:[{in:'TAXR',fmt:F.pct0},{in:'B_COD',fmt:F.yr}]});
  tsRow(ws,R.fcfe,'Equity cash flow','EURm',F.num,(X,p,a)=>`${X}$${R.codf}*(-(1-${a[0]})*$F$${D.TCAPEX})+${X}$${R.olf}*(${X}$${R.rev}-${X}$${R.opex}-${X}$${R.tax}-${X}$${R.prin}-${X}$${R.intr})`,{subtotal:true,aud:[{in:'B_GEAR',fmt:F.pct0}]});
  ws.getCell(R.date,3).value='Date (for XIRR)';
  for(let i=0;i<NY;i++){const c=ws.getCell(R.date,c0+i);c.value={formula:`DATE(${colL(c0+i)}$${YROW},12,31)`};c.numFmt=F.date;}
  ws.getCell(R.chk,3).value='Check: dashboard equity CF';
  for(let i=0;i<NY;i++){const y=Y0+i;const c=ws.getCell(R.chk,c0+i);c.value=(S.chkBatt&&S.chkBatt[y]!==undefined?S.chkBatt[y]:0);c.fill=CHK;c.numFmt=F.num;}
  tsRow(ws,R.diff,'Check difference (should be 0)','EURm',F.num,X=>`${X}$${R.fcfe}-${X}$${R.chk}`);
  ws.getCell(R.xcf,3).value='Equity CF for XIRR (seed at t0, immaterial)';
  ws.getCell(R.xcf,c0).value=-0.01; ws.getCell(R.xcf,c0).numFmt=F.num;
  for(let i=1;i<NY;i++){const c=ws.getCell(R.xcf,c0+i);c.value={formula:`${colL(c0+i)}${R.fcfe}`};c.numFmt=F.num;}

  let k=rr+3; k=sect(ws,k,5,'RESULTS');
  const put=(row,label,formula,fmtStr,unit,fill)=>{ws.getCell(row,3).value=label;
   if(unit){const u=ws.getCell(row,5);u.value=unit;u.font={size:9.5,color:{argb:'FF7A7A7A'}};}
   const c=ws.getCell(row,6);c.value={formula:formula};c.numFmt=fmtStr;if(fill)c.fill=fill;return row;};
  const IRR=k; put(k++,'Equity IRR (XIRR)',`IFERROR(XIRR(${colL(c0)}${R.xcf}:${lastC}${R.xcf},${colL(c0)}${R.date}:${lastC}${R.date}),"n/m")`,F.pct2,'per yr'); ws.getCell(IRR,6).font=bold;
  put(k++,'Equity',`$F$${D.EQUITY}`,F.eur,'EURm');
  put(k++,'Total capex',`$F$${D.TCAPEX}`,F.eur,'EURm');
  const MX=k; put(k++,'Max check difference',`MAX(ABS(MIN(${colL(c0)}${R.diff}:${lastC}${R.diff})),ABS(MAX(${colL(c0)}${R.diff}:${lastC}${R.diff})))`,F.num,'EURm',CHK);
  put(k++,'Tie-out',`CHOOSE(1+($F$${MX}>0.001),"OK ties to dashboard","CHECK differences")`,'General');
  return {irr:`Battery!$F$${IRR}`,cells:{ARBRR:D.ARBRR,ANC:D.ANC,DCCH:D.DCCH,GFCAP:D.GFCAP,GFENE:D.GFENE,CELLCX:D.CELLCX,TCAPEX:D.TCAPEX}};
 }
 const batteryRef=batterySheet();
 const BX=k=>`Battery!$F$${batteryRef.cells[k]}`;

 // ================= SPV =================
 function spvSheet(){
  const ws=calcSheet('SPV','Consolidated Energy SPV | owns RES (per mode), battery and private electrical infrastructure');
  const D={}, R={};
  let rr=4; rr=sect(ws,rr,1,'ASSUMPTIONS (linked from Inputs)');
  [['INFL','Inflation','per yr',F.pct1],['TAXR','Tax rate','%',F.pct0],['GEAR','RES gearing','%',F.pct0],['TENOR','Debt tenor','years',F.yr],['AMORT','Amortisation (1 = annuity)','1/0',F.yr],['RATE','RES debt rate','per yr',F.pct2],['PPAT','PPA term','years',F.yr],
   ['W_MW','Wind capacity','MW',F.int],['W_CAPEX','Wind capex','EURm/MW','0.000'],['W_COD','Wind COD','year',F.yr],['W_LIFE','Wind life','years',F.yr],['W_OPEX','Wind opex','EURm/MW','0.000'],['W_TAIL','Wind tail price','EUR/MWh',F.price],
   ['S_MW','Solar capacity','MWp',F.int],['S_CAPEX','Solar capex','EURm/MWp','0.000'],['S_COD','Solar COD','year',F.yr],['S_LIFE','Solar life','years',F.yr],['S_OPEX','Solar opex','EURm/MWp','0.000'],['S_TAIL','Solar tail price','EUR/MWh',F.price],
   ['B_GEAR','Battery gearing','%',F.pct0],['B_RATE','Battery debt rate','per yr',F.pct2],['B_GY','Battery merchant year','year',F.yr],['B_COD','Battery COD','year',F.yr],['B_LIFE','Battery life','years',F.yr],['B_DEGR','Battery degradation','per yr',F.pct1],['B_COMP','Battery compression','per yr',F.pct1],['B_OPEXP','Battery opex','% of cell capex',F.pct1],
   ['DC_MW','DC load','MW',F.int],['DC_P','DC price','EUR/MWh',F.price],['RES_P','Residual price','EUR/MWh',F.price],['RES_M','Trading margin','%',F.pct1],['FEE_E','NE3 energy fee','EUR/MWh',F.price],['FEE_C','NE3 capacity fee','EUR/kW/yr',F.price],['FEESC','Grid-fee escalation','per yr',F.pct1],['RES_OWN','RES owned (1/0)','1/0',F.yr],['LINE_C','Direct line cost','EURm/100MW','0.00'],['DCAC','Solar DC/AC ratio','x','0.00'],
   ['SPV_MODE','Pass-through (1/0)','1/0',F.yr],['SPV_M','SPV margin, percent form','% of cost',F.pct1],['MGN_MODE','Margin form (1 = EUR/MWh)','1/0',F.yr],['MGN_E','SPV margin, euro form','EUR/MWh',F.price],['SPV_FF','SPV first revenue year','year',F.yr]
  ].forEach(x=>{rr=lnk(ws,rr,D,x[0],x[1],x[0],x[3],x[2]);});
  const A=k=>`$F$${D[k]}`;

  rr+=2; rr=sect(ws,rr,2,'LINKED FROM THE BATTERY SHEET');
  const BL={};
  [['ARBRR','Battery arbitrage run-rate','EURm/yr'],['ANC','Battery ancillary','EURm/yr'],['DCCH','Battery DC charge','EURm/yr'],
   ['GFCAP','Battery grid fee, capacity','EURm/yr'],['GFENE','Battery grid fee, energy','EURm/yr'],['CELLCX','Battery cell capex','EURm'],['TCAPEX','Battery total capex','EURm']
  ].forEach(x=>{ws.getCell(rr,3).value=x[1]; const c=ws.getCell(rr,6); c.value={formula:BX(x[0])}; c.font=GRN; c.numFmt=F.num;
   ws.getCell(rr,5).value=x[2]; ws.getCell(rr,5).font={size:9.5,color:{argb:'FF7A7A7A'}}; BL[x[0]]=rr; rr++;});
  const BXL=k=>`$F$${BL[k]}`;

  rr+=2; rr=sect(ws,rr,3,'DERIVED (computed once)');
  rr=derRow(ws,rr,D,'RESCX','RES capex owned',`${A('RES_OWN')}*(${A('W_MW')}*${A('W_CAPEX')}+${A('S_MW')}*${A('S_CAPEX')})`,F.eur,'EURm');
  rr=derRow(ws,rr,D,'LINECX','Line capex (export capacity base)',`(${A('W_MW')}+${A('S_MW')}/${A('DCAC')})*${A('LINE_C')}/100`,F.eur,'EURm');
  rr=derRow(ws,rr,D,'SPVCX','SPV total capex',`$F$${D.RESCX}+${BXL('TCAPEX')}+$F$${D.LINECX}`,F.eur,'EURm');
  rr=derRow(ws,rr,D,'IRRCX','Financed SPV capex, incl. line',`$F$${D.SPVCX}`,F.eur,'EURm');
  rr=derRow(ws,rr,D,'SENDBT','Senior RES + line debt base',`${A('GEAR')}*($F$${D.RESCX}+$F$${D.LINECX})`,F.eur,'EURm');
  rr=derRow(ws,rr,D,'BATDBT','Battery debt base',`${A('B_GEAR')}*${BXL('TCAPEX')}`,F.eur,'EURm');
  rr=derRow(ws,rr,D,'BGEAR','Blended SPV gearing',`($F$${D.SENDBT}+$F$${D.BATDBT})/MAX(0.0001,$F$${D.IRRCX})`,F.fac,'x');
  rr=derRow(ws,rr,D,'BRATE','Blended SPV debt rate',`(${A('GEAR')}*($F$${D.RESCX}+$F$${D.LINECX})*${A('RATE')}+${A('B_GEAR')}*${BXL('TCAPEX')}*${A('B_RATE')})/MAX(0.0001,$F$${D.SENDBT}+$F$${D.BATDBT})`,F.pct2,'per yr');
  rr=derRow(ws,rr,D,'RESLIFE','RES life',`MAX(${A('W_LIFE')},${A('S_LIFE')})`,F.yr,'years');
  rr=derRow(ws,rr,D,'DEPY','Depreciation period',`MIN(20,$F$${D.RESLIFE})`,F.yr,'years');
  rr=derRow(ws,rr,D,'ANNDEP','Annual SPV depreciation',`$F$${D.IRRCX}/$F$${D.DEPY}`,F.num,'EURm');
  rr=derRow(ws,rr,D,'REPY','Debt amortisation period',`MIN(${A('TENOR')},$F$${D.RESLIFE})`,F.yr,'years');
  rr=derRow(ws,rr,D,'DRAW1','Construction draw at FF-2',`0.3*($F$${D.RESCX}+$F$${D.LINECX})*$F$${D.BGEAR}`,F.num,'EURm');
  rr=derRow(ws,rr,D,'IDC1','Construction interest, year 1',`$F$${D.DRAW1}/2*$F$${D.BRATE}`,F.num,'EURm');
  rr=derRow(ws,rr,D,'BAL1','Debt after FF-2',`$F$${D.DRAW1}+$F$${D.IDC1}`,F.num,'EURm');
  rr=derRow(ws,rr,D,'CAPFF1','SPV capex at FF-1',`0.7*($F$${D.RESCX}+$F$${D.LINECX})+${BXL('TCAPEX')}`,F.num,'EURm');
  rr=derRow(ws,rr,D,'DRAW2','Construction draw at FF-1',`$F$${D.CAPFF1}*$F$${D.BGEAR}`,F.num,'EURm');
  rr=derRow(ws,rr,D,'IDC2','Construction interest, year 2',`($F$${D.BAL1}+$F$${D.DRAW2}/2)*$F$${D.BRATE}`,F.num,'EURm');
  rr=derRow(ws,rr,D,'DEBTFF','Debt drawn by first revenue year',`$F$${D.BAL1}+$F$${D.DRAW2}+$F$${D.IDC2}`,F.num,'EURm');
  rr=derRow(ws,rr,D,'ANNPRIN','Annual principal, flat',`$F$${D.DEBTFF}/$F$${D.REPY}`,F.num,'EURm');
  rr=derRow(ws,rr,D,'ANNDS','Annual debt service, annuity',`PMT($F$${D.BRATE},$F$${D.REPY},-$F$${D.DEBTFF})`,F.num,'EURm');

  rr+=2; rr=sect(ws,rr,4,'YEARLY MODEL');
  R.op  =tsRow(ws,rr++,'Operating flag','1/0',F.yr,(X,p,a)=>`(${yr(X)}>=${a[0]})*(${yr(X)}<${a[0]}+$F$${D.RESLIFE})`,{aud:[{in:'SPV_FF',fmt:F.yr}]});
  R.dcrev=rr++;
  R.bmf =tsRow(ws,rr++,'Battery merchant flag','1/0',F.yr,(X,p,a)=>`(${yr(X)}>=${a[0]})*(${yr(X)}<${a[1]}+${A('B_LIFE')})`,{aud:[{in:'B_GY',fmt:F.yr},{in:'B_COD',fmt:F.yr}]});
  R.bcf =tsRow(ws,rr++,'Battery operating flag','1/0',F.yr,(X,p,a)=>`(${yr(X)}>${a[0]})*(${yr(X)}<${a[0]}+${A('B_LIFE')})`,{aud:[{in:'B_COD',fmt:F.yr}]});
  R.bmer=tsRow(ws,rr++,'Battery merchant revenue','EURm',F.num,(X,p,a)=>`${X}$${R.bmf}*(${BXL('ARBRR')}*(1-${a[0]})^MAX(0,${yr(X)}-${a[2]})+${BXL('ANC')})*(1-${a[1]})^MAX(0,${yr(X)}-${a[2]})*(1+${A('INFL')})^MAX(0,${yr(X)}-${a[2]})`,{aud:[{in:'B_DEGR',fmt:F.pct1},{in:'B_COMP',fmt:F.pct1},{in:'B_GY',fmt:F.yr}]});
  R.bcap=tsRow(ws,rr++,'Battery DC charge','EURm',F.num,(X,p,a)=>`${X}$${R.bcf}*${BXL('DCCH')}*(1+${A('INFL')})^MAX(0,${yr(X)}-(${a[0]}+1))`,{aud:[{in:'B_COD',fmt:F.yr}]});
  R.brev=tsRow(ws,rr++,'Battery revenue','EURm',F.num,X=>`${X}$${R.bmer}+${X}$${R.bcap}`,{subtotal:true});
  R.wprod=tsRow(ws,rr++,'Wind to DC','MWh',F.int,X=>`Wind!${X}$${windRef.prodRow}`);
  R.sprod=tsRow(ws,rr++,'Solar to DC','MWh',F.int,X=>`Solar!${X}$${solarRef.prodRow}`);
  R.resm=tsRow(ws,rr++,'Residual from grid','MWh',F.int,(X,p,a)=>`${X}$${R.op}*MAX(0,${a[0]}*8760-${X}$${R.wprod}-${X}$${R.sprod})`,{aud:[{in:'DC_MW',fmt:F.int}]});
  R.resp=tsRow(ws,rr++,'Residual energy price','EUR/MWh',F.price,(X,p,a)=>`${a[0]}*(1+${A('INFL')})^(${yr(X)}-2025)*(1+${a[1]})+${a[2]}*(1+${A('FEESC')})^MAX(0,${yr(X)}-2028)`,{aud:[{in:'RES_P',fmt:F.price},{in:'RES_M',fmt:F.pct1},{in:'FEE_E',fmt:F.price}]});
  R.resec=tsRow(ws,rr++,'Residual energy cost','EURm',F.num,X=>`${X}$${R.op}*${X}$${R.resm}*${X}$${R.resp}/10^6`);
  R.rescc=tsRow(ws,rr++,'Residual capacity fee','EURm',F.num,(X,p,a)=>`${X}$${R.op}*${a[0]}*${a[1]}/1000*(1+${A('FEESC')})^MAX(0,${yr(X)}-2028)`,{aud:[{in:'DC_MW',fmt:F.int},{in:'FEE_C',fmt:F.price}]});
  R.resc=tsRow(ws,rr++,'Residual cost: total','EURm',F.num,X=>`${X}$${R.resec}+${X}$${R.rescc}`,{subtotal:true});
  R.wpf =tsRow(ws,rr++,'Wind in-PPA flag','1/0',F.yr,(X,p,a)=>`(${yr(X)}<${a[0]}+MIN(${A('PPAT')},${A('W_LIFE')}))`,{aud:[{in:'W_COD',fmt:F.yr}]});
  R.spf =tsRow(ws,rr++,'Solar in-PPA flag','1/0',F.yr,(X,p,a)=>`(${yr(X)}<${a[0]}+MIN(${A('PPAT')},${A('S_LIFE')}))`,{aud:[{in:'S_COD',fmt:F.yr}]});
  R.wpr =tsRow(ws,rr++,'Wind price to SPV','EUR/MWh',F.price,(X,p,a)=>`${X}$${R.wpf}*${a[0]}+(1-${X}$${R.wpf})*${a[1]}`,{aud:[{f:windRef.blend,fmt:F.price},{in:'W_TAIL',fmt:F.price}]});
  R.spr =tsRow(ws,rr++,'Solar price to SPV','EUR/MWh',F.price,(X,p,a)=>`${X}$${R.spf}*${a[0]}+(1-${X}$${R.spf})*${a[1]}`,{aud:[{f:solarRef.blend,fmt:F.price},{in:'S_TAIL',fmt:F.price}]});
  R.resppa=tsRow(ws,rr++,'RES bought at PPA (mode = 0)','EURm',F.num,(X,p,a)=>`${X}$${R.op}*(1-${a[0]})*(${X}$${R.wpr}*${X}$${R.wprod}+${X}$${R.spr}*${X}$${R.sprod})/10^6`,{aud:[{in:'RES_OWN',fmt:F.yr}]});
  R.resown=tsRow(ws,rr++,'RES billed at own LCOE (mode = 1)','EURm',F.num,(X,p,a)=>`${X}$${R.op}*${a[0]}*(${a[1]}*${X}$${R.wprod}+${a[2]}*${X}$${R.sprod})/10^6`,{aud:[{in:'RES_OWN',fmt:F.yr},{f:windRef.lcoeCell,fmt:F.price},{f:solarRef.lcoeCell,fmt:F.price}]});
  R.pasee=tsRow(ws,rr++,'Pass-through energy base','EURm',F.num,X=>`${X}$${R.resppa}+${X}$${R.resown}+${X}$${R.resec}`);
  R.mgn =tsRow(ws,rr++,'SPV margin','EURm',F.num,(X,p,a)=>`${X}$${R.op}*(${a[0]}*${A('DC_MW')}*8760*${a[1]}*(1+${A('INFL')})^(${yr(X)}-2026)/10^6+(1-${a[0]})*${X}$${R.pasee}*${a[2]})`,{aud:[{in:'MGN_MODE',fmt:F.yr},{in:'MGN_E',fmt:F.price},{in:'SPV_M',fmt:F.pct1}]});
  tsRow(ws,R.dcrev,'DC revenue','EURm',F.num,(X,p)=>`${X}$${R.op}*(${A('SPV_MODE')}*(${X}$${R.pasee}+${X}$${R.rescc}+${X}$${R.mgn})+(1-${A('SPV_MODE')})*${A('DC_MW')}*8760*${A('DC_P')}*(1+${A('INFL')})^(${yr(X)}-2026)/10^6)`,{subtotal:true});
  R.cpi =tsRow(ws,rr++,'CPI index vs 2023','x',F.fac,(X,p,a)=>`(1+${a[0]})^(${yr(X)}-2023)`,{aud:[{in:'INFL',fmt:F.pct1}]});
  R.resox=tsRow(ws,rr++,'RES opex owned','EURm',F.num,(X,p,a)=>`${X}$${R.op}*${a[0]}*(${A('W_OPEX')}*${A('W_MW')}+${A('S_OPEX')}*${A('S_MW')})*${X}$${R.cpi}`,{aud:[{in:'RES_OWN',fmt:F.yr}]});
  R.batox=tsRow(ws,rr++,'Battery cell opex','EURm',F.num,X=>`${X}$${R.op}*${BXL('CELLCX')}*${A('B_OPEXP')}*${X}$${R.cpi}`);
  R.batgf=tsRow(ws,rr++,'Battery grid fees','EURm',F.num,(X,p,a)=>`${X}$${R.op}*(${yr(X)}>=${a[0]})*(${BXL('GFCAP')}+${BXL('GFENE')})*(1+${A('FEESC')})^MAX(0,${yr(X)}-2028)`,{aud:[{in:'B_GY',fmt:F.yr}]});
  R.opex=tsRow(ws,rr++,'Opex: total','EURm',F.num,X=>`${X}$${R.resox}+${X}$${R.batox}+${X}$${R.batgf}`,{subtotal:true});
  R.ebit=tsRow(ws,rr++,'EBITDA','EURm',F.num,X=>`${X}$${R.dcrev}+${X}$${R.brev}-${X}$${R.resc}-${X}$${R.resppa}-${X}$${R.opex}`,{subtotal:true});
  R.cshr=tsRow(ws,rr++,'RES + line capex draw share','x','0.00',(X,p,a)=>`0.3*(${yr(X)}=${a[0]}-2)+0.7*(${yr(X)}=${a[0]}-1)`,{aud:[{in:'SPV_FF',fmt:F.yr}]});
  R.capex=tsRow(ws,rr++,'SPV capex, incl. direct line','EURm',F.num,(X,p,a)=>`${X}$${R.cshr}*($F$${D.RESCX}+$F$${D.LINECX})+(${yr(X)}=${a[0]}-1)*${BXL('TCAPEX')}`,{aud:[{in:'SPV_FF',fmt:F.yr}]});
  R.draw=tsRow(ws,rr++,'Debt draw','EURm',F.num,X=>`${X}$${R.capex}*$F$${D.BGEAR}`);
  R.cflag=tsRow(ws,rr++,'Construction flag','1/0',F.yr,(X,p,a)=>`(${yr(X)}>=${a[0]}-2)*(${yr(X)}<${a[0]})`,{aud:[{in:'SPV_FF',fmt:F.yr}]});
  R.idc=rr++; R.intr=rr++; R.prin=rr++; R.bal=rr++; R.dep=rr++; R.ebt=rr++; R.nol=rr++; R.tax=rr++; R.fcfe=rr++; R.date=rr++; R.xcf=rr++;
  tsRow(ws,R.idc,'Construction interest','EURm',F.num,(X,pX)=>`${X}$${R.cflag}*((${pX?pX+'$'+R.bal:'0'})+${X}$${R.draw}/2)*$F$${D.BRATE}`);
  tsRow(ws,R.intr,'Interest','EURm',F.num,(X,pX,a)=>`(${yr(X)}>=${a[0]})*(${pX?pX+'$'+R.bal:'0'})*$F$${D.BRATE}`,{aud:[{in:'SPV_FF',fmt:F.yr}]});
  tsRow(ws,R.prin,'Principal repaid','EURm',F.num,(X,pX,a)=>`(${yr(X)}>=${a[1]})*(${yr(X)}<${a[1]}+$F$${D.REPY})*MIN((${pX?pX+'$'+R.bal:'0'}),${a[0]}*MAX(0,$F$${D.ANNDS}-${X}$${R.intr})+(1-${a[0]})*$F$${D.ANNPRIN})`,{aud:[{in:'AMORT',fmt:F.yr},{in:'SPV_FF',fmt:F.yr}]});
  tsRow(ws,R.bal,'Debt balance, end of year','EURm',F.num,(X,pX,a)=>`(${yr(X)}<${a[0]})*((${pX?pX+'$'+R.bal:'0'})+${X}$${R.draw}+${X}$${R.idc})+(${yr(X)}>=${a[0]})*MAX(0,(${pX?pX+'$'+R.bal:'0'})-${X}$${R.prin})`,{aud:[{in:'SPV_FF',fmt:F.yr}]});
  tsRow(ws,R.dep,'Depreciation','EURm',F.num,(X,p,a)=>`(${yr(X)}>=${a[0]})*(${yr(X)}<${a[0]}+$F$${D.DEPY})*$F$${D.ANNDEP}`,{aud:[{in:'SPV_FF',fmt:F.yr}]});
  tsRow(ws,R.ebt,'EBT','EURm',F.num,X=>`${X}$${R.ebit}-${X}$${R.dep}-${X}$${R.intr}`);
  tsRow(ws,R.nol,'Tax-loss balance','EURm',F.num,(X,pX,a)=>`(${yr(X)}>=${a[0]})*MIN(0,(${pX?pX+'$'+R.nol:'0'})+${X}$${R.ebt})`,{aud:[{in:'SPV_FF',fmt:F.yr}]});
  tsRow(ws,R.tax,'Tax','EURm',F.num,(X,pX,a)=>`(${yr(X)}>=${a[1]})*MAX(0,${X}$${R.ebt}+(${pX?pX+'$'+R.nol:'0'}))*${a[0]}`,{aud:[{in:'TAXR',fmt:F.pct0},{in:'SPV_FF',fmt:F.yr}]});
  tsRow(ws,R.fcfe,'Equity cash flow','EURm',F.num,(X,p,a)=>`(${yr(X)}<${a[0]})*(-${X}$${R.capex}+${X}$${R.draw})+(${yr(X)}>=${a[0]})*(${X}$${R.ebit}-${X}$${R.tax}-${X}$${R.prin}-${X}$${R.intr})`,{subtotal:true,aud:[{in:'SPV_FF',fmt:F.yr}]});
  ws.getCell(R.date,3).value='Date (for XIRR)';
  for(let i=0;i<NY;i++){const c=ws.getCell(R.date,c0+i);c.value={formula:`DATE(${colL(c0+i)}$${YROW},12,31)`};c.numFmt=F.date;}
  ws.getCell(R.xcf,3).value='Equity CF for XIRR (seed at t0, immaterial)';
  ws.getCell(R.xcf,c0).value=-0.01; ws.getCell(R.xcf,c0).numFmt=F.num;
  for(let i=1;i<NY;i++){const c=ws.getCell(R.xcf,c0+i);c.value={formula:`${colL(c0+i)}${R.fcfe}`};c.numFmt=F.num;}

  let k=rr+3; k=sect(ws,k,5,'RESULTS');
  const put=(row,label,formula,fmtStr,unit,fill)=>{ws.getCell(row,3).value=label;
   if(unit){const u=ws.getCell(row,5);u.value=unit;u.font={size:9.5,color:{argb:'FF7A7A7A'}};}
   const c=ws.getCell(row,6);c.value={formula:formula};c.numFmt=fmtStr;if(fill)c.fill=fill;return row;};
  const IRR=k; put(k++,'SPV equity IRR, incl. direct-line capex (XIRR)',`IFERROR(XIRR(${colL(c0)}${R.xcf}:${lastC}${R.xcf},${colL(c0)}${R.date}:${lastC}${R.date}),"n/m")`,F.pct2,'per yr'); ws.getCell(IRR,6).font=bold;
  ws.getCell(k,3).value='Dashboard SPV IRR at export'; ws.getCell(k,6).value=S.spvIRR; ws.getCell(k,6).numFmt=F.pct2; ws.getCell(k,6).fill=CHK; k++;
  put(k++,'Financed SPV capex, incl. direct line',`$F$${D.IRRCX}`,F.eur,'EURm');
  put(k++,'SPV total capex, incl. direct line',`$F$${D.SPVCX}`,F.eur,'EURm');
  ws.getCell(k,3).value='Method note'; ws.getCell(k,6).value='Asset-only returns and the financing waterfall live on the Wind, Solar and Battery sheets. Battery arbitrage uses exported day-by-day backtest averages for the selected duration and year; re-export after changing those. A non-zero tie-out identifies a reconciliation issue.'; k++;
  return {irr:`SPV!$F$${IRR}`};
 }
 const spvRef=spvSheet();

 // ================= DASHBOARD =================
 const wo=wb.addWorksheet('Dashboard');
 wo.views=[{state:'frozen',xSplit:6,showGridLines:false}];
 [2.5,3,46,2,14,16,2].forEach((w,i)=>wo.getColumn(i+1).width=w);
 wo.getCell(1,2).value='Project Burgenland'; wo.getCell(1,2).font={bold:true,size:15,color:{argb:'FF1F7A33'}};
 wo.getCell(2,3).value='Dashboard | every figure recalculates from the Inputs sheet'; wo.getCell(2,3).font={italic:true,size:9,color:{argb:'FF808080'}};
 let ro=4, dsNo=0;
 function dsect(txt){wo.getCell(ro,1).value=++dsNo; wo.getCell(ro,1).font=SECF;
  const c=wo.getCell(ro,3); c.value=txt; c.font=SECF; ro++;}
 function dput(label,formula,fmtStr,unit,boldV){wo.getCell(ro,3).value=label;
  if(unit){const u=wo.getCell(ro,5);u.value=unit;u.font={size:10,color:{argb:'FF7A7A7A'}};}
  const c=wo.getCell(ro,6); c.value={formula:formula}; c.numFmt=fmtStr; c.font=boldV?{bold:true,color:{argb:'FF008000'}}:GRN; ro++;}
 dsect('RETURNS');
 dput('Wind equity IRR',windRef.irr,F.pct2,'per yr',true);
 dput('Wind MOIC',windRef.moic,F.x,'x');
 dput('Wind LCOE',windRef.lcoe,F.price,'EUR/MWh');
 dput('Solar equity IRR',solarRef.irr,F.pct2,'per yr',true);
 dput('Solar MOIC',solarRef.moic,F.x,'x');
 dput('Solar LCOE',solarRef.lcoe,F.price,'EUR/MWh');
 dput('Battery equity IRR',batteryRef.irr,F.pct2,'per yr');
 dput('Energy SPV equity IRR, incl. direct line',spvRef.irr,F.pct2,'per yr',true);
 ro++; dsect('PPA PRICING (per-asset tranches)');
 dput('Wind blended PPA in force',windRef.blend,F.price,'EUR/MWh');
 dput('Solar blended PPA in force',solarRef.blend,F.price,'EUR/MWh');
 dput('Tranche 1 price',inCell('TP1'),F.price,'EUR/MWh');
 dput('Tranche 2 price',inCell('TP2'),F.price,'EUR/MWh');
 ro++; dsect('FINANCING & PAYOUT | WIND');
 dput('Total capex',windRef.tcapex,F.eur,'EURm');
 dput('Senior debt at drawdown',windRef.debt,F.eur,'EURm');
 dput('Construction funding (junior leg)',windRef.equity,F.eur,'EURm');
 dput('Funding IRR',windRef.subIRR,F.pct2,'per yr',true);
 dput('Funding MOIC',windRef.subMOIC,F.x,'x');
 dput('Funding repaid in year',windRef.payY,F.yr,'year',true);
 dput('Payback, years of operation',windRef.payN,F.yr,'years');
 dput('Burgenland cash, first 10 operating years',windRef.be10,F.eur,'EURm',true);
 dput('Burgenland cash, total',windRef.beTot,F.eur,'EURm');
 ro++; dsect('FINANCING & PAYOUT | SOLAR');
 dput('Total capex',solarRef.tcapex,F.eur,'EURm');
 dput('Senior debt at drawdown',solarRef.debt,F.eur,'EURm');
 dput('Construction funding (junior leg)',solarRef.equity,F.eur,'EURm');
 dput('Funding IRR',solarRef.subIRR,F.pct2,'per yr',true);
 dput('Funding MOIC',solarRef.subMOIC,F.x,'x');
 dput('Funding repaid in year',solarRef.payY,F.yr,'year',true);
 dput('Payback, years of operation',solarRef.payN,F.yr,'years');
 dput('Burgenland cash, first 10 operating years',solarRef.be10,F.eur,'EURm',true);
 dput('Burgenland cash, total',solarRef.beTot,F.eur,'EURm');
 ro++;
 wo.getCell(ro,3).value='All inputs live on the Inputs sheet (yellow, blue font). Green cells are linked; change an input and every sheet recalculates.'; ro++;
 wo.getCell(ro,3).value='Each calc sheet reproduces a row\'s drivers in columns H to J, so F2 on any year cell shows them beside the formula.'; ro++;
 return wb;
}
if(typeof module!=='undefined')module.exports={buildFullModel};
