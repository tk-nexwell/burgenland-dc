// Full-formula Excel model generator, readable and auditable build.
// Runs identically in Node (test) and browser (ExcelJS). buildFullModel(ExcelJSlib, S) -> workbook.
//
// Design rules (Formula & Layout Standards):
//  - NO named ranges. Each calc sheet carries its own LOCAL ASSUMPTIONS block whose cells are
//    linked (=Inputs!$B$n) and shown in RED, so F2 traces to a nearby labelled cell.
//  - NO IF(). Conditions are 0/1 flags built by boolean multiplication ((y>=cod)*(y<cod+life));
//    caps/floors use MIN/MAX. IFERROR is used only to print "n/m" for an undefined battery IRR.
//  - HELPER ROWS ARE FREE. Every intermediate step is its own labelled row (label in col A,
//    unit in col C on single-value rows; time series filled across with $-anchored refs).
//  - ONE FORMULA PER ROW, filled across; business numbers live only in the yellow Inputs cells.
//  - Tie-out check row + a ✓/⚠ flag (via CHOOSE, not IF) next to each asset's equity IRR.
function buildFullModel(ExcelJS, S){
 const wb=new ExcelJS.Workbook(); wb.creator='Nexwell Power'; wb.created=new Date();
 wb.calcProperties.fullCalcOnLoad=true;
 const Y0=2026, YN=2060, NY=YN-Y0+1, c0=2;
 const lastC=colL(c0+NY-1);
 // styles
 const YEL={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFF2CC'}};   // input fill
 const HDR={type:'pattern',pattern:'solid',fgColor:{argb:'FF1F3864'}};   // header navy
 const CHK={type:'pattern',pattern:'solid',fgColor:{argb:'FFE2EFDA'}};   // check green
 const SUBT={type:'pattern',pattern:'solid',fgColor:{argb:'FFF2F2F2'}};  // section band
 const BLUE={color:{argb:'FF0000FF'}};                 // input font
 const RED ={color:{argb:'FFC00000'}};                 // linked-from-Inputs font
 const bold={bold:true}, white={color:{argb:'FFFFFFFF'},bold:true};
 const numF='#,##0.00;[Red]-#,##0.00', pctF='0.00%', eurF='#,##0.0', xF='0.00"x"', intF='#,##0';
 function colL(n){let s='';while(n>0){const m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=(n-m-1)/26;}return s;}
 function sect(ws,r,txt){const c=ws.getCell(r,1);c.value=txt;c.font=bold;c.fill=SUBT;ws.getCell(r,2).fill=SUBT;ws.getCell(r,3).fill=SUBT;return r+1;}

 // ================= INPUTS =================
 const wi=wb.addWorksheet('Inputs');
 wi.getColumn(1).width=2.4; wi.getColumn(2).width=2.6; wi.getColumn(3).width=42; wi.getColumn(4).width=24; wi.getColumn(5).width=14; wi.getColumn(6).width=42; wi.views=[{state:'frozen',ySplit:2}];
 wi.getCell(1,3).value='GDC NICKELSDORF | Power SPV model'; wi.getCell(1,3).font={bold:true,size:15,color:{argb:'FF2D7D32'}};
 wi.getCell(2,3).value='Yellow = editable inputs (blue). Calc sheets link these in red. Exported '+S.today+'.'; wi.getCell(2,3).font={italic:true,size:9,color:{argb:'FF808080'}};
 const IN={};                 // name -> Inputs row number
 let r=4;
 function isect(txt){const c=wi.getCell(r,2);c.value=txt;c.font={bold:true,size:11,color:{argb:'FF2D7D32'}};for(let k=1;k<=6;k++)wi.getCell(r,k).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFEFF5EE'}};wi.getRow(r).height=17;r++;}
 function inp(label,value,unit,key){
  wi.getCell(r,3).value=label;
  wi.getCell(r,4).value=unit||''; wi.getCell(r,4).font={size:10,color:{argb:'FF7A7A7A'}};
  const c=wi.getCell(r,5); c.value=value; c.fill=YEL; c.font=BLUE; c.border={outline:{style:'thin'}};
  c.numFmt = (typeof value==='number' && Math.abs(value)<1 && value!==0)?'0.####':(Number.isInteger(value)?'#,##0':'#,##0.####');
  IN[key]=r; r++;
 }
 isect('MACRO / FINANCING');
 inp('Inflation (CPI)',S.macro.infl,'per yr','INFL');
 inp('Tax rate',S.macro.tax,'','TAXR');
 inp('Gearing (wind/solar/line)',S.macro.gearing,'debt share','GEAR');
 inp('Debt tenor',S.macro.tenor,'years','TENOR');
 inp('Amortisation (1 = annuity, 0 = flat principal)',S.macro.amort==='annuity'?1:0,'level total payment, or level principal','AMORT');
 inp('All-in debt rate',S.macro.allInRate,'','RATE');
 inp('PPA term',S.macro.ppaTermY,'years','PPAT');
 inp('Merchant power (uncontracted RES)',S.macro.merchReal,'€/MWh 2023-real, CPI-indexed','MERCH');
 r++; isect('WIND');
 inp('Capacity',S.wind.mw,'MW AC','W_MW');
 inp('Capex',S.wind.capexPerMW,'€m per MW','W_CAPEX');
 inp('Gross capacity factor',S.wind.grossCF,'','W_GCF');
 inp('Plant losses (wake, electrical)',S.wind.loss,'','W_LOSS');
 inp('Direct-line losses',S.wind.lineLoss,'editable model input','W_LLOSS');
 inp('Degradation',S.wind.degr,'per yr','W_DEGR');
 inp('Opex',S.wind.opexPerMW,'€m/MW/yr, CPI-indexed','W_OPEX');
 inp('PPA price',S.wind.ppa,'€/MWh fixed','W_PPA');
 inp('Contracted share',S.wind.contr,'','W_CONTR');
 inp('COD (first generation year)',S.wind.codY,'model input','W_COD');
 inp('Useful life',S.wind.lifeY,'years','W_LIFE');
 inp('Merchant tail price',S.CAP7w,'€/MWh flat (capture avg ex-2022)','W_TAIL');
 r++; isect('SOLAR');
 inp('Capacity',S.solar.mw,'MWp DC','S_MW');
 inp('Capex',S.solar.capexPerMW,'€m per MWp','S_CAPEX');
 inp('Gross capacity factor',S.solar.grossCF,'','S_GCF');
 inp('Plant losses (soiling, inverter)',S.solar.loss,'','S_LOSS');
 inp('Direct-line losses',S.solar.lineLoss,'','S_LLOSS');
 inp('Degradation',S.solar.degr,'per yr','S_DEGR');
 inp('Opex',S.solar.opexPerMW,'€m/MWp/yr, CPI-indexed','S_OPEX');
 inp('PPA price',S.solar.ppa,'€/MWh fixed','S_PPA');
 inp('Contracted share',S.solar.contr,'','S_CONTR');
 inp('COD (first generation year)',S.solar.codY,'','S_COD');
 inp('Useful life',S.solar.lifeY,'years','S_LIFE');
 inp('Merchant tail price',S.CAP7s,'€/MWh flat','S_TAIL');
 r++; isect('BATTERY');
 inp('Power',S.battery.powerMW,'MW','B_MW');
 inp('Duration',S.battery.durationH,'hours','B_DUR');
 inp('Cell capex',S.battery.capexPerKWh,'€/kWh','B_CKWH');
 inp('Grid-interface allowance',S.battery.substation,'€m, model input','B_SUB');
 inp('Campus connection allowance',S.battery.interconnect,'€m, model input','B_INT');
 inp('Round-trip efficiency',S.battery.rte,'','B_RTE');
 inp('SoC floor (outage reserve)',S.battery.socFloor,'not traded','B_SOC');
 inp('Cycles per day',S.battery.cyclesDay,'','B_CYC');
 inp('Capture factor',S.battery.captureFactor,'× volatility uplift','B_CAPF');
 inp('Opex',S.battery.opexPct,'% of cell capex/yr, CPI-indexed','B_OPEXP');
 inp('Ancillary revenue',S.battery.ancPerMW,'€k/MW/yr (from merchant year)','B_ANC');
 inp('DC reliability charge (paid by DC)',S.battery.capChargeMWyr,'€k/MW/yr from COD+1','B_CCH');
 inp('First merchant year',S.battery.gridYear,'model input','B_GY');
 inp('Revenue compression',S.battery.compression,'per yr on merchant','B_COMP');
 inp('Degradation',S.battery.degr,'per yr','B_DEGR');
 inp('Gearing',S.battery.gearing,'','B_GEAR');
 inp('Debt rate',S.battery.debtRate,'','B_RATE');
 inp('Grid fee: capacity (NE3), manual override',S.battery.gridCapFee,'€/kW/yr, esc from 2028','B_GFC');
 inp('Apply DC NE3 capacity tariff when merchant',S.battery.mktCapFee?1:0,'1=yes, 0=no','B_MKT');
 inp('Grid fee: energy (NE3)',S.battery.gridEnergyFee,'€/MWh on throughput','B_GFE');
 inp('Backtest average buy price',S.battery.backtestAvgBuy,'€/MWh · selected duration/year','B_AVGBUY');
 inp('Backtest average sell price',S.battery.backtestAvgSell,'€/MWh · selected duration/year','B_AVGSELL');
 inp('Backtest self-charge share',S.battery.backtestSelfShare,'share of charging','B_FSELF');
 inp('Self-charge cost',S.battery.btmCharge||0,'€/MWh','B_BTM');
 inp('Battery COD (capex year)',S.COD,'','B_COD');
 inp('Useful life',S.battery.lifeY,'years','B_LIFE');
 r++; isect('DATA CENTER / GRID BALANCING');
 inp('DC contracted load',S.dc.firmMW,'MW','DC_MW');
 inp('DC sale price (today)',S.dc.dcPrice,'€/MWh, CPI-indexed','DC_P');
 inp('Residual market price',S.dc.resFix,'€/MWh 2025-real, CPI-indexed','RES_P');
 inp('Energy-trading margin',S.dc.beMargin,'','RES_M');
 inp('NE3 energy fee',S.dc.gridEnergyFee,'€/MWh (2028 level)','FEE_E');
 inp('NE3 capacity fee',S.dc.gridCapFeeKW,'€/kW/yr (2028 level)','FEE_C');
 inp('Grid-fee escalation',S.dc.feeEsc,'per yr from 2028','FEESC');
 inp('RES sourcing (1=owned, 0=PPA)',(S.dc.resMode||'lcoe')==='lcoe'?1:0,'drives SPV sheet','RES_OWN');
 inp('Direct line cost',S.linePer100,'€m per 100 MW','LINE_C');
 inp('Solar DC/AC ratio',S.dcac,'line is sized on export capacity','DCAC');
 inp('Billing (1=pass-through, 0=fixed)',(S.dc.spvMode||'pass')==='pass'?1:0,'cost plus margin, or a flat €/MWh','SPV_MODE');
 inp('SPV margin (% of energy cost)',S.dc.spvMargin!=null?S.dc.spvMargin:0.03,'used when margin form = 0','SPV_M');
 inp('Margin form (1=€/MWh, 0=%)',S.dc.marginMode==='flat'?1:0,'€/MWh is neutral to the market','MGN_MODE');
 inp('SPV margin (€/MWh)',S.dc.marginEur!=null?S.dc.marginEur:3.5,'CPI-indexed, used when form = 1','MGN_E');
 inp('SPV first revenue year',S.FF,'','SPV_FF');
 r++; isect('REFERENCE AVERAGE-DAY CURVE | price profile reference | '+S.priceYear);
 const PH0=r;
 for(let h=0;h<24;h++){ wi.getCell(r,3).value='Hour '+h; const c=wi.getCell(r,5); c.value=S.ph[h]; c.fill=YEL; c.font=BLUE; c.numFmt='#,##0.0'; wi.getCell(r,4).value=(h===0?'→ Battery sheet':''); r++; }
 IN.PH0=PH0;

 const inCell = key => `Inputs!$E$${IN[key]}`;   // absolute link to an input

 // ---- helper: write a filled-across time-series row -> returns the row number used
 function tsRow(ws, row, label, Yrow, fmt, fFn, opt){
  opt=opt||{};
  const a=ws.getCell(row,1); a.value=label; if(opt.bold){a.font=bold;} if(opt.band){a.fill=SUBT;}
  for(let i=0;i<NY;i++){
   const X=colL(c0+i), pX=i?colL(c0+i-1):null;
   const c=ws.getCell(row,c0+i); c.value={formula:fFn(X,pX)}; c.numFmt=fmt;
   if(opt.bold) c.font=bold;
  }
  return row;
 }
 // ---- helper: local linked-assumption row (red). Returns row number. Records into map L.
 function linkRow(ws, row, L, key, label, unit){
  ws.getCell(row,1).value=label;
  const c=ws.getCell(row,2); c.value={formula:inCell(key)}; c.font=RED; c.numFmt='General';
  ws.getCell(row,3).value=unit||''; ws.getCell(row,3).font={italic:true,color:{argb:'FF808080'}};
  L[key]=row; return row+1;
 }
 // ---- helper: local derived single-value row (black). Returns row number. Records into D.
 function derRow(ws, row, D, key, label, formula, fmt, unit){
  ws.getCell(row,1).value=label;
  const c=ws.getCell(row,2); c.value={formula:formula}; c.numFmt=fmt||numF;
  ws.getCell(row,3).value=unit||'';
  D[key]=row; return row+1;
 }
 function yearHeader(ws,row){ ws.getCell(row,1).value='Year'; ws.getCell(row,1).font=bold;
  for(let i=0;i<NY;i++){const c=ws.getCell(row,c0+i); c.value=Y0+i; c.font=white; c.fill=HDR;} return row; }

 // ================= WIND / SOLAR builder =================
 function assetSheet(name,P,chk){
  const ws=wb.addWorksheet(name); ws.getColumn(1).width=38; ws.getColumn(2).width=13; ws.getColumn(3).width=30; ws.views=[{state:'frozen',xSplit:1}];
  ws.getCell(1,1).value=name.toUpperCase()+' | full-formula model'; ws.getCell(1,1).font={bold:true,size:12};
  ws.getCell(2,1).value='Red = linked from Inputs (F2 to trace). Flags are 0/1 (no IF); caps use MIN/MAX. One step per row.';
  ws.getCell(2,1).font={italic:true,size:9,color:{argb:'FF808080'}};
  const L={}, D={}, R={};
  let rr=4; rr=sect(ws,rr,'LOCAL ASSUMPTIONS  (linked from Inputs, shown red)');
  [['INFL','Inflation (CPI)','per yr'],['TAXR','Tax rate',''],['GEAR','Gearing',''],['TENOR','Debt tenor','yrs'],['AMORT','Amortisation (1=annuity)',''],
   ['RATE','Debt rate',''],['PPAT','PPA term','yrs'],['MERCH','Merchant (uncontracted)','€/MWh 23-real'],
   [P+'_MW','Capacity','MW'],[P+'_CAPEX','Capex','€m/MW'],[P+'_GCF','Gross capacity factor',''],
   [P+'_LOSS','Plant losses',''],[P+'_LLOSS','Direct-line losses',''],[P+'_DEGR','Degradation','per yr'],
   [P+'_OPEX','Opex','€m/MW/yr'],[P+'_PPA','PPA price','€/MWh'],[P+'_CONTR','Contracted share',''],
   [P+'_COD','COD (first gen year)',''],[P+'_LIFE','Useful life','yrs'],[P+'_TAIL','Merchant tail price','€/MWh']
  ].forEach(x=>{ rr=linkRow(ws,rr,L,x[0],x[1],x[2]); });
  const A=k=>`$B$${L[k]}`;                         // local assumption cell (abs)

  rr++; rr=sect(ws,rr,'DERIVED (computed once)');
  rr=derRow(ws,rr,D,'GROSS',   'Gross generation @100% avail (MWh)', `${A(P+'_MW')}*8760*${A(P+'_GCF')}`, intF,'MW×8760×CF');
  rr=derRow(ws,rr,D,'NETFAC',  'Net delivery factor', `(1-${A(P+'_LOSS')})*(1-${A(P+'_LLOSS')})`, '0.0000','(1−plant)(1−line)');
  rr=derRow(ws,rr,D,'TCAPEX',  'Total capex (€m)', `${A(P+'_MW')}*${A(P+'_CAPEX')}`, eurF);
  rr=derRow(ws,rr,D,'EQUITY',  'Equity (€m)', `(1-${A('GEAR')})*$B$${D.TCAPEX}`, eurF,'(1−gearing)×capex');
  rr=derRow(ws,rr,D,'DEPY',    'Depreciation period (yrs)', `MIN(20,${A(P+'_LIFE')})`, intF);
  rr=derRow(ws,rr,D,'ANNDEP',  'Annual depreciation (€m)', `$B$${D.TCAPEX}/$B$${D.DEPY}`, numF,'straight line');
  rr=derRow(ws,rr,D,'REPY',    'Debt amortisation period (yrs)', `MIN(${A('TENOR')},${A(P+'_LIFE')})`, intF);
  rr=derRow(ws,rr,D,'ANNF',    'Annuity factor (LCOE)', `${A('RATE')}/(1-(1+${A('RATE')})^-${A(P+'_LIFE')})`, '0.0000','r / (1−(1+r)^−life)');
  rr=derRow(ws,rr,D,'CAPANN',  'Capex annuity (€m/yr)', `$B$${D.TCAPEX}*$B$${D.ANNF}`, numF);
  rr=derRow(ws,rr,D,'PPAEND',  'PPA end year', `${A(P+'_COD')}+MIN(${A('PPAT')},${A(P+'_LIFE')})`, intF);
  rr=derRow(ws,rr,D,'DRAW1',   'Constr. draw @ COD-2 (€m)', `0.3*$B$${D.TCAPEX}*${A('GEAR')}`, numF,'30% of capex × gearing');
  rr=derRow(ws,rr,D,'IDC1',    'Constr. IDC year 1 (€m)', `$B$${D.DRAW1}/2*${A('RATE')}`, numF);
  rr=derRow(ws,rr,D,'BAL1',    'Debt after COD-2 (€m)', `$B$${D.DRAW1}+$B$${D.IDC1}`, numF);
  rr=derRow(ws,rr,D,'DRAW2',   'Constr. draw @ COD-1 (€m)', `0.7*$B$${D.TCAPEX}*${A('GEAR')}`, numF,'70% of capex × gearing');
  rr=derRow(ws,rr,D,'IDC2',    'Constr. IDC year 2 (€m)', `($B$${D.BAL1}+$B$${D.DRAW2}/2)*${A('RATE')}`, numF);
  rr=derRow(ws,rr,D,'DEBTCOD', 'Debt drawn by COD (€m)', `$B$${D.BAL1}+$B$${D.DRAW2}+$B$${D.IDC2}`, numF,'construction debt at COD-1');
  rr=derRow(ws,rr,D,'ANNPRIN', 'Annual principal, flat (€m)', `$B$${D.DEBTCOD}/$B$${D.REPY}`, numF);
  rr=derRow(ws,rr,D,'ANNDS',   'Annual debt service, annuity (€m)', `PMT(${A('RATE')},$B$${D.REPY},-$B$${D.DEBTCOD})`, numF,'level total payment');

  rr++; const Yrow=rr; yearHeader(ws,Yrow);
  rr++; rr=sect(ws,rr,'PRODUCTION');
  const yr=X=>`${X}$${Yrow}`;
  R.op   =tsRow(ws,rr++,'Operating flag (1/0)',Yrow,intF,X=>`(${yr(X)}>=${A(P+'_COD')})*(${yr(X)}<${A(P+'_COD')}+${A(P+'_LIFE')})`);
  R.ysc  =tsRow(ws,rr++,'Years since COD',Yrow,intF,X=>`${yr(X)}-${A(P+'_COD')}`);
  R.cpi  =tsRow(ws,rr++,'CPI index vs 2023',Yrow,'0.0000',X=>`(1+${A('INFL')})^(${yr(X)}-2023)`);
  R.degf =tsRow(ws,rr++,'Degradation factor',Yrow,'0.0000',X=>`(1-${A(P+'_DEGR')})^MAX(0,${X}$${R.ysc})`);
  R.prod =tsRow(ws,rr++,'Production (MWh)',Yrow,intF,X=>`${X}$${R.op}*$B$${D.GROSS}*$B$${D.NETFAC}*${X}$${R.degf}`);
  rr=sect(ws,rr,'REVENUE  (split contracted / merchant / tail)');
  R.ppaf =tsRow(ws,rr++,'In-PPA-period flag',Yrow,intF,X=>`(${yr(X)}<$B$${D.PPAEND})`);
  R.cmwh =tsRow(ws,rr++,'Contracted MWh',Yrow,intF,X=>`${X}$${R.op}*${X}$${R.ppaf}*${X}$${R.prod}*${A(P+'_CONTR')}`);
  R.mmwh =tsRow(ws,rr++,'Merchant MWh (in PPA period)',Yrow,intF,X=>`${X}$${R.op}*${X}$${R.ppaf}*${X}$${R.prod}*(1-${A(P+'_CONTR')})`);
  R.tmwh =tsRow(ws,rr++,'Tail MWh (post-PPA)',Yrow,intF,X=>`${X}$${R.op}*(1-${X}$${R.ppaf})*${X}$${R.prod}`);
  R.mpr  =tsRow(ws,rr++,'Merchant price (€/MWh)',Yrow,eurF,X=>`${A('MERCH')}*${X}$${R.cpi}`);
  R.rppa =tsRow(ws,rr++,'Revenue: contracted (€m)',Yrow,numF,X=>`${X}$${R.cmwh}*${A(P+'_PPA')}/10^6`);
  R.rmer =tsRow(ws,rr++,'Revenue: merchant (€m)',Yrow,numF,X=>`${X}$${R.mmwh}*${X}$${R.mpr}/10^6`);
  R.rtail=tsRow(ws,rr++,'Revenue: tail (€m)',Yrow,numF,X=>`${X}$${R.tmwh}*${A(P+'_TAIL')}/10^6`);
  R.rev  =tsRow(ws,rr++,'Revenue: total (€m)',Yrow,numF,X=>`${X}$${R.rppa}+${X}$${R.rmer}+${X}$${R.rtail}`);
  rr=sect(ws,rr,'COSTS & EBITDA');
  R.opex =tsRow(ws,rr++,'Opex (€m)',Yrow,numF,X=>`${X}$${R.op}*${A(P+'_OPEX')}*${A(P+'_MW')}*${X}$${R.cpi}`);
  R.ebit =tsRow(ws,rr++,'EBITDA (€m)',Yrow,numF,X=>`${X}$${R.rev}-${X}$${R.opex}`);
  rr=sect(ws,rr,'CAPEX & DEBT');
  R.cshr =tsRow(ws,rr++,'Capex draw share (0/0.3/0.7)',Yrow,'0.00',X=>`0.3*(${yr(X)}=${A(P+'_COD')}-2)+0.7*(${yr(X)}=${A(P+'_COD')}-1)`);
  R.capex=tsRow(ws,rr++,'Capex (€m)',Yrow,numF,X=>`${X}$${R.cshr}*$B$${D.TCAPEX}`);
  R.draw =tsRow(ws,rr++,'Debt draw (€m)',Yrow,numF,X=>`${X}$${R.capex}*${A('GEAR')}`);
  R.cflag=tsRow(ws,rr++,'Construction flag',Yrow,intF,X=>`(${yr(X)}>=${A(P+'_COD')}-2)*(${yr(X)}<${A(P+'_COD')})`);
  R.idc  =rr++;   // reserve interdependent rows first; fill after all row numbers are known
  R.intr =rr++;
  R.prin =rr++;
  R.bal  =rr++;
  R.dep  =rr++;
  R.ebt  =rr++;
  R.nol  =rr++;
  R.tax  =rr++;
  R.fcfe =rr++;
  R.date =rr++;
  R.chk  =rr++;
  R.diff =rr++;
  R.xcf  =rr++;
  // now fill the reserved rows with correct references (bal known)
  // re-do IDC with proper prevBal ref:
  tsRow(ws,R.idc,'IDC: interest during constr. (€m)',Yrow,numF,(X,pX)=>`${X}$${R.cflag}*((${pX?pX+'$'+R.bal:'0'})+${X}$${R.draw}/2)*${A('RATE')}`);
  tsRow(ws,R.intr,'Interest (€m)',Yrow,numF,(X,pX)=>`(${yr(X)}>=${A(P+'_COD')})*(${pX?pX+'$'+R.bal:'0'})*${A('RATE')}`);
  R.rflag=null; // repayment flag folded into principal for brevity
  tsRow(ws,R.prin,'Principal repay (€m)',Yrow,numF,(X,pX)=>`(${yr(X)}>=${A(P+'_COD')})*(${yr(X)}<${A(P+'_COD')}+$B$${D.REPY})*MIN((${pX?pX+'$'+R.bal:'0'}),${A('AMORT')}*MAX(0,$B$${D.ANNDS}-${X}$${R.intr})+(1-${A('AMORT')})*$B$${D.ANNPRIN})`);
  tsRow(ws,R.bal,'Debt balance EOY (€m)',Yrow,numF,(X,pX)=>`(${yr(X)}<${A(P+'_COD')})*((${pX?pX+'$'+R.bal:'0'})+${X}$${R.draw}+${X}$${R.idc})+(${yr(X)}>=${A(P+'_COD')})*MAX(0,(${pX?pX+'$'+R.bal:'0'})-${X}$${R.prin})`);
  tsRow(ws,R.dep,'Depreciation (€m)',Yrow,numF,X=>`(${yr(X)}>=${A(P+'_COD')})*(${yr(X)}<${A(P+'_COD')}+$B$${D.DEPY})*$B$${D.ANNDEP}`);
  tsRow(ws,R.ebt,'EBT (€m)',Yrow,numF,X=>`${X}$${R.ebit}-${X}$${R.dep}-${X}$${R.intr}`);
  tsRow(ws,R.nol,'NOL balance (€m)',Yrow,numF,(X,pX)=>`(${yr(X)}>=${A(P+'_COD')})*MIN(0,(${pX?pX+'$'+R.nol:'0'})+${X}$${R.ebt})`);
  tsRow(ws,R.tax,'Tax (€m)',Yrow,numF,(X,pX)=>`(${yr(X)}>=${A(P+'_COD')})*MAX(0,${X}$${R.ebt}+(${pX?pX+'$'+R.nol:'0'}))*${A('TAXR')}`);
  tsRow(ws,R.fcfe,'EQUITY CASH FLOW (€m)',Yrow,numF,X=>`(${yr(X)}<${A(P+'_COD')})*(-${X}$${R.capex}+${X}$${R.draw})+(${yr(X)}>=${A(P+'_COD')})*(${X}$${R.ebit}-${X}$${R.tax}-${X}$${R.prin}-${X}$${R.intr})`,{bold:true});
  for(let i=0;i<NY;i++){const c=ws.getCell(R.date,c0+i);c.value={formula:`DATE(${colL(c0+i)}$${Yrow},12,31)`};c.numFmt='yyyy-mm-dd';}
  ws.getCell(R.date,1).value='Date (XIRR)';
  ws.getCell(R.chk,1).value='Check: dashboard equity CF';
  for(let i=0;i<NY;i++){const y=Y0+i;const c=ws.getCell(R.chk,c0+i);c.value=(chk&&chk[y]!==undefined?chk[y]:0);c.fill=CHK;c.numFmt=numF;}
  tsRow(ws,R.diff,'Check diff (≈0)',Yrow,numF,X=>`${X}$${R.fcfe}-${X}$${R.chk}`);
  ws.getCell(R.fcfe,1).fill=HDR; ws.getCell(R.fcfe,1).font=white;
  ws.getCell(R.xcf,1).value='Equity CF for XIRR (dummy seed − t0, immaterial)';
  ws.getCell(R.xcf,c0).value=-0.01; ws.getCell(R.xcf,c0).numFmt=numF;
  for(let i=1;i<NY;i++){const c=ws.getCell(R.xcf,c0+i);c.value={formula:`${colL(c0+i)}${R.fcfe}`};c.numFmt=numF;}

  // ---- RESULTS
  let k=rr+1; k=sect(ws,k,'RESULTS');
  const put=(row,label,formula,fmt,fill)=>{ws.getCell(row,1).value=label;const c=ws.getCell(row,2);c.value={formula:formula};c.numFmt=fmt;if(fill)c.fill=fill;return row;};
  const IRR=k; put(k++,'Equity IRR (XIRR)',`XIRR(B${R.xcf}:${lastC}${R.xcf},B${R.date}:${lastC}${R.date})`,pctF); ws.getCell(IRR,2).font=bold;
  const MOIC=k; put(k++,'MOIC',`SUMIF(B${R.fcfe}:${lastC}${R.fcfe},">0")/-SUMIF(B${R.fcfe}:${lastC}${R.fcfe},"<0")`,xF);
  put(k++,'Total capex (€m)',`$B$${D.TCAPEX}`,eurF);
  put(k++,'Equity (€m)',`$B$${D.EQUITY}`,eurF);
  const FP=k; put(k++,'First-year production (MWh)',`SUMIF(B${Yrow}:${lastC}${Yrow},${A(P+'_COD')},B${R.prod}:${lastC}${R.prod})`,intF);
  const FO=k; put(k++,'First-year opex (€m)',`SUMIF(B${Yrow}:${lastC}${Yrow},${A(P+'_COD')},B${R.opex}:${lastC}${R.opex})`,numF);
  const LC=k; put(k++,'LCOE (€/MWh)',`SUMPRODUCT((B${R.capex}:${lastC}${R.capex}+B${R.opex}:${lastC}${R.opex})*(1+${A('RATE')})^-(B${Yrow}:${lastC}${Yrow}-${A(P+'_COD')}))*10^6/SUMPRODUCT(B${R.prod}:${lastC}${R.prod}*(1+${A('RATE')})^-(B${Yrow}:${lastC}${Yrow}-${A(P+'_COD')}))`,eurF);
  const MX=k; put(k++,'Max |check diff| (€m)',`MAX(ABS(MIN(B${R.diff}:${lastC}${R.diff})),ABS(MAX(B${R.diff}:${lastC}${R.diff})))`,numF,CHK);
  put(k++,'Tie-out',`CHOOSE(1+($B$${MX}>0.001),"✓ ties to dashboard","⚠ check diff")`,'General');
  return {irr:`${name}!B${IRR}`, moic:`${name}!B${MOIC}`, lcoe:`${name}!B${LC}`, prodRow:R.prod};
 }
 const windRef =assetSheet('Wind','W',S.chkWind);
 const solarRef=assetSheet('Solar','S',S.chkSolar);

 // ================= BATTERY =================
 function batterySheet(){
  const ws=wb.addWorksheet('Battery'); ws.getColumn(1).width=38; ws.getColumn(2).width=13; ws.getColumn(3).width=26; ws.getColumn(4).width=8; ws.views=[{state:'frozen',xSplit:1}];
  ws.getCell(1,1).value='BATTERY | full-formula model'; ws.getCell(1,1).font={bold:true,size:12};
  ws.getCell(2,1).value='Red = linked from Inputs. Arbitrage uses the dashboard day-by-day backtest summary; the 24h curve is a reference view.';
  ws.getCell(2,1).font={italic:true,size:9,color:{argb:'FF808080'}};
  const L={}, D={}, R={};
  let rr=4; rr=sect(ws,rr,'LOCAL ASSUMPTIONS (linked from Inputs, red)');
  [['INFL','Inflation','per yr'],['TAXR','Tax rate',''],['TENOR','Debt tenor','yrs'],['AMORT','Amortisation (1=annuity)',''],['FEESC','Grid-fee escalation','per yr'],
   ['B_MW','Power','MW'],['B_DUR','Duration','h'],['B_CKWH','Cell capex','€/kWh'],['B_SUB','Grid-interface allowance','€m'],['B_INT','Campus connection allowance','€m'],
   ['B_RTE','Round-trip eff.',''],['B_SOC','SoC floor',''],['B_CYC','Cycles/day',''],['B_CAPF','Capture factor',''],
   ['B_OPEXP','Opex %/yr',''],['B_ANC','Ancillary','€k/MW/yr'],['B_CCH','DC reliability charge','€k/MW/yr'],
   ['B_GY','First merchant year',''],['B_COMP','Compression','per yr'],['B_DEGR','Degradation','per yr'],
   ['B_GEAR','Gearing',''],['B_RATE','Debt rate',''],['B_GFC','Grid fee capacity, manual','€/kW/yr'],['B_MKT','Apply DC NE3 tariff','1/0'],['FEE_C','DC NE3 capacity tariff','€/kW/yr'],['B_GFE','Grid fee energy','€/MWh'],
   ['B_AVGBUY','Backtest average buy','€/MWh'],['B_AVGSELL','Backtest average sell','€/MWh'],['B_FSELF','Self-charge share',''],['B_BTM','Self-charge cost','€/MWh'],
   ['B_COD','COD (capex year)',''],['B_LIFE','Useful life','yrs']
  ].forEach(x=>{ rr=linkRow(ws,rr,L,x[0],x[1],x[2]); });
  const A=k=>`$B$${L[k]}`;

  rr++; rr=sect(ws,rr,'REFERENCE 24h CURVE (not used in financial formulas)');
  const CST=rr;
  for(let h=0;h<24;h++){
   ws.getCell(rr,1).value='Hour '+h;
   const p=ws.getCell(rr,2); p.value={formula:`Inputs!$E$${IN.PH0+h}`}; p.font=RED; p.numFmt='#,##0.0';
   rr++;
  }
  const CEN=rr-1;
  for(let h=0;h<24;h++){ const rk=ws.getCell(CST+h,3); rk.value={formula:`RANK($B$${CST+h},$B$${CST}:$B$${CEN},1)`}; rk.numFmt=intF; }
  ws.getCell(CST-1,3).value='rank (1=cheapest)';
  const PB=`$B$${CST}:$B$${CEN}`, RB=`$C$${CST}:$C$${CEN}`;

  rr++; rr=sect(ws,rr,'DERIVED DISPATCH & CAPEX (once)');
  rr=derRow(ws,rr,D,'NCH',   'Tradeable hours nCh', `MIN(12,MAX(0,${A('B_DUR')}*(1-${A('B_SOC')})*${A('B_CYC')}))`, numF,'= dur×(1−SoC)×cycles, preserving fractional tradeable duration');
  rr=derRow(ws,rr,D,'SCHEAP','Effective daily charge cost (€/MW)', `$B$${D.NCH}*(${A('B_AVGBUY')}*(1-${A('B_FSELF')})+${A('B_FSELF')}*${A('B_BTM')})`, numF,'day-by-day backtest average plus self-charge share');
  rr=derRow(ws,rr,D,'SPRICE','Daily sale value before RTE (€/MW)', `$B$${D.NCH}*${A('B_AVGSELL')}`, numF,'day-by-day backtest average');
  rr=derRow(ws,rr,D,'DAYARB','Day arbitrage (€/day)', `MAX(0,${A('B_MW')}*(${A('B_RTE')}*$B$${D.SPRICE}-$B$${D.SCHEAP}))*${A('B_CAPF')}`, numF);
  rr=derRow(ws,rr,D,'ARBRR', 'Arbitrage revenue run-rate (€m/yr)', `$B$${D.DAYARB}*365/10^6`, numF);
  rr=derRow(ws,rr,D,'ANC',   'Ancillary (€m/yr)', `${A('B_ANC')}*MIN(${A('B_MW')},225)/1000`, numF,'public planning cap on addressable capacity');
  rr=derRow(ws,rr,D,'DCCH',  'DC reliability charge (€m/yr)', `${A('B_CCH')}*${A('B_MW')}/1000`, numF);
  rr=derRow(ws,rr,D,'GFCAP', 'Grid fee capacity (€m/yr, 2028)', `${A('B_MW')}*MAX(${A('B_GFC')},${A('B_MKT')}*${A('FEE_C')})/1000`, numF,'manual override or DC NE3 fallback, matching dashboard');
  rr=derRow(ws,rr,D,'THRU',  'Grid throughput (MWh/yr)', `$B$${D.NCH}*${A('B_MW')}*${A('B_RTE')}*365`, intF);
  rr=derRow(ws,rr,D,'GFENE', 'Grid fee energy (€m/yr, 2028)', `$B$${D.THRU}*${A('B_GFE')}/10^6`, numF);
  rr=derRow(ws,rr,D,'CELLCX','Cell capex (€m)', `${A('B_MW')}*${A('B_DUR')}*1000*${A('B_CKWH')}/10^6`, eurF);
  rr=derRow(ws,rr,D,'TCAPEX','Total capex (€m)', `$B$${D.CELLCX}+${A('B_SUB')}+${A('B_INT')}`, eurF);
  rr=derRow(ws,rr,D,'EQUITY','Equity (€m)', `(1-${A('B_GEAR')})*$B$${D.TCAPEX}`, eurF);
  rr=derRow(ws,rr,D,'DEPY',  'Depreciation period (yrs)', `MIN(15,${A('B_LIFE')}-1)`, intF);
  rr=derRow(ws,rr,D,'ANNDEP','Annual depreciation (€m)', `$B$${D.TCAPEX}/$B$${D.DEPY}`, numF);
  rr=derRow(ws,rr,D,'REPY',  'Debt amortisation period (yrs)', `MIN(${A('TENOR')},${A('B_LIFE')}-1)`, intF);
  rr=derRow(ws,rr,D,'ANNPRIN','Annual principal, flat (€m)', `${A('B_GEAR')}*$B$${D.TCAPEX}/$B$${D.REPY}`, numF);
  rr=derRow(ws,rr,D,'ANNDS',  'Annual debt service, annuity (€m)', `PMT(${A('B_RATE')},$B$${D.REPY},-${A('B_GEAR')}*$B$${D.TCAPEX})`, numF,'level total payment');

  rr++; const Yrow=rr; yearHeader(ws,Yrow); const yr=X=>`${X}$${Yrow}`;
  rr++; rr=sect(ws,rr,'YEARLY MODEL');
  R.mflag=tsRow(ws,rr++,'Merchant flag (y≥ban)',Yrow,intF,X=>`(${yr(X)}>=${A('B_GY')})`);
  R.olf  =tsRow(ws,rr++,'Operating flag (COD<y<COD+life)',Yrow,intF,X=>`(${yr(X)}>${A('B_COD')})*(${yr(X)}<${A('B_COD')}+${A('B_LIFE')})`);
  R.cpi  =tsRow(ws,rr++,'CPI index vs 2023',Yrow,'0.0000',X=>`(1+${A('INFL')})^(${yr(X)}-2023)`);
  R.merch=tsRow(ws,rr++,'Merchant revenue (€m)',Yrow,numF,X=>`${X}$${R.olf}*${X}$${R.mflag}*($B$${D.ARBRR}*(1-${A('B_DEGR')})^MAX(0,${yr(X)}-${A('B_GY')})+$B$${D.ANC})*(1-${A('B_COMP')})^MAX(0,${yr(X)}-${A('B_GY')})*(1+${A('INFL')})^MAX(0,${yr(X)}-${A('B_GY')})`);
  R.cap  =tsRow(ws,rr++,'DC reliability charge (€m)',Yrow,numF,X=>`${X}$${R.olf}*$B$${D.DCCH}*(1+${A('INFL')})^MAX(0,${yr(X)}-(${A('B_COD')}+1))`);
  R.rev  =tsRow(ws,rr++,'Revenue: total (€m)',Yrow,numF,X=>`${X}$${R.merch}+${X}$${R.cap}`);
  R.opxc =tsRow(ws,rr++,'Opex: cell O&M (€m)',Yrow,numF,X=>`${X}$${R.olf}*$B$${D.CELLCX}*${A('B_OPEXP')}*${X}$${R.cpi}`);
  R.gfee =tsRow(ws,rr++,'Grid fees (€m, from ban)',Yrow,numF,X=>`${X}$${R.olf}*${X}$${R.mflag}*($B$${D.GFCAP}+$B$${D.GFENE})*(1+${A('FEESC')})^MAX(0,${yr(X)}-2028)`);
  R.opex =tsRow(ws,rr++,'Opex: total (€m)',Yrow,numF,X=>`${X}$${R.opxc}+${X}$${R.gfee}`);
  R.codf =tsRow(ws,rr++,'COD flag (y=COD)',Yrow,intF,X=>`(${yr(X)}=${A('B_COD')})`);
  R.intr =rr++; R.prin=rr++; R.bal=rr++; R.dep=rr++; R.ebt=rr++; R.nol=rr++; R.tax=rr++; R.fcfe=rr++; R.date=rr++; R.chk=rr++; R.diff=rr++; R.xcf=rr++;
  tsRow(ws,R.intr,'Interest (€m)',Yrow,numF,(X,pX)=>`(${yr(X)}>${A('B_COD')})*(${pX?pX+'$'+R.bal:'0'})*${A('B_RATE')}`);
  tsRow(ws,R.prin,'Principal repay (€m)',Yrow,numF,(X,pX)=>`(${yr(X)}>${A('B_COD')})*(${yr(X)}<=${A('B_COD')}+$B$${D.REPY})*MIN((${pX?pX+'$'+R.bal:'0'}),${A('AMORT')}*MAX(0,$B$${D.ANNDS}-${X}$${R.intr})+(1-${A('AMORT')})*$B$${D.ANNPRIN})`);
  tsRow(ws,R.bal,'Debt balance EOY (€m)',Yrow,numF,(X,pX)=>`${X}$${R.codf}*${A('B_GEAR')}*$B$${D.TCAPEX}+(${yr(X)}>${A('B_COD')})*MAX(0,(${pX?pX+'$'+R.bal:'0'})-${X}$${R.prin})`);
  tsRow(ws,R.dep,'Depreciation (€m)',Yrow,numF,X=>`(${yr(X)}>${A('B_COD')})*(${yr(X)}<=${A('B_COD')}+$B$${D.DEPY})*$B$${D.ANNDEP}`);
  tsRow(ws,R.ebt,'EBT (€m)',Yrow,numF,X=>`${X}$${R.rev}-${X}$${R.opex}-${X}$${R.dep}-${X}$${R.intr}`);
  tsRow(ws,R.nol,'NOL balance (€m)',Yrow,numF,(X,pX)=>`(${yr(X)}>${A('B_COD')})*MIN(0,(${pX?pX+'$'+R.nol:'0'})+${X}$${R.ebt})`);
  tsRow(ws,R.tax,'Tax (€m)',Yrow,numF,(X,pX)=>`(${yr(X)}>${A('B_COD')})*MAX(0,${X}$${R.ebt}+(${pX?pX+'$'+R.nol:'0'}))*${A('TAXR')}`);
  tsRow(ws,R.fcfe,'EQUITY CASH FLOW (€m)',Yrow,numF,X=>`${X}$${R.codf}*(-(1-${A('B_GEAR')})*$B$${D.TCAPEX})+${X}$${R.olf}*(${X}$${R.rev}-${X}$${R.opex}-${X}$${R.tax}-${X}$${R.prin}-${X}$${R.intr})`,{bold:true});
  for(let i=0;i<NY;i++){const c=ws.getCell(R.date,c0+i);c.value={formula:`DATE(${colL(c0+i)}$${Yrow},12,31)`};c.numFmt='yyyy-mm-dd';}
  ws.getCell(R.date,1).value='Date (XIRR)';
  ws.getCell(R.chk,1).value='Check: dashboard equity CF';
  for(let i=0;i<NY;i++){const y=Y0+i;const c=ws.getCell(R.chk,c0+i);c.value=(S.chkBatt&&S.chkBatt[y]!==undefined?S.chkBatt[y]:0);c.fill=CHK;c.numFmt=numF;}
  tsRow(ws,R.diff,'Check diff (≈0)',Yrow,numF,X=>`${X}$${R.fcfe}-${X}$${R.chk}`);
  ws.getCell(R.fcfe,1).fill=HDR; ws.getCell(R.fcfe,1).font=white;
  ws.getCell(R.xcf,1).value='Equity CF for XIRR (dummy seed − t0, immaterial)';
  ws.getCell(R.xcf,c0).value=-0.01; ws.getCell(R.xcf,c0).numFmt=numF;
  for(let i=1;i<NY;i++){const c=ws.getCell(R.xcf,c0+i);c.value={formula:`${colL(c0+i)}${R.fcfe}`};c.numFmt=numF;}

  let k=rr+1; k=sect(ws,k,'RESULTS');
  const IRR=k; ws.getCell(k,1).value='Equity IRR (XIRR)'; ws.getCell(k,2).value={formula:`IFERROR(XIRR(B${R.xcf}:${lastC}${R.xcf},B${R.date}:${lastC}${R.date}),"n/m")`}; ws.getCell(k,2).numFmt=pctF; ws.getCell(k,2).font=bold; k++;
  ws.getCell(k,1).value='Equity (€m)'; ws.getCell(k,2).value={formula:`$B$${D.EQUITY}`}; ws.getCell(k,2).numFmt=eurF; k++;
  ws.getCell(k,1).value='Total capex (€m)'; ws.getCell(k,2).value={formula:`$B$${D.TCAPEX}`}; ws.getCell(k,2).numFmt=eurF; k++;
  const MX=k; ws.getCell(k,1).value='Max |check diff| (€m)'; ws.getCell(k,2).value={formula:`MAX(ABS(MIN(B${R.diff}:${lastC}${R.diff})),ABS(MAX(B${R.diff}:${lastC}${R.diff})))`}; ws.getCell(k,2).numFmt=numF; ws.getCell(k,2).fill=CHK; k++;
  ws.getCell(k,1).value='Tie-out'; ws.getCell(k,2).value={formula:`CHOOSE(1+($B$${MX}>0.001),"✓ ties to dashboard","⚠ check diff")`}; k++;
  // expose derived cells for the SPV sheet
  return {irr:`Battery!B${IRR}`, cells:{ARBRR:D.ARBRR,ANC:D.ANC,DCCH:D.DCCH,GFCAP:D.GFCAP,GFENE:D.GFENE,CELLCX:D.CELLCX,TCAPEX:D.TCAPEX,OPEXP:L.B_OPEXP,DEGR:L.B_DEGR,COMP:L.B_COMP,GY:L.B_GY,COD:L.B_COD,LIFE:L.B_LIFE}};
 }
 const batteryRef=batterySheet();
 const BX=k=>`Battery!$B$${batteryRef.cells[k]}`;   // link a Battery derived/assumption cell

 // ================= SPV =================
 function spvSheet(){
  const ws=wb.addWorksheet('SPV'); ws.getColumn(1).width=40; ws.getColumn(2).width=13; ws.getColumn(3).width=24; ws.views=[{state:'frozen',xSplit:1}];
  ws.getCell(1,1).value='CONSOLIDATED POWER SPV | owns RES (per mode), battery and private electrical infrastructure'; ws.getCell(1,1).font={bold:true,size:12};
  ws.getCell(2,1).value='Red = linked. The consolidated return includes every asset assigned to the Power SPV, including the private line.';
  ws.getCell(2,1).font={italic:true,size:9,color:{argb:'FF808080'}};
  const L={}, D={}, R={};
  let rr=4; rr=sect(ws,rr,'LOCAL ASSUMPTIONS (linked from Inputs, red)');
  [['INFL','Inflation','per yr'],['TAXR','Tax rate',''],['GEAR','RES gearing',''],['TENOR','Debt tenor','yrs'],['AMORT','Amortisation (1=annuity)',''],['RATE','RES debt rate',''],['PPAT','PPA term','yrs'],
   ['W_MW','Wind MW',''],['W_CAPEX','Wind capex','€m/MW'],['W_COD','Wind COD',''],['W_LIFE','Wind life','yrs'],['W_OPEX','Wind opex','€m/MW'],['W_PPA','Wind PPA','€/MWh'],['W_TAIL','Wind tail','€/MWh'],
   ['S_MW','Solar MW',''],['S_CAPEX','Solar capex','€m/MW'],['S_COD','Solar COD',''],['S_LIFE','Solar life','yrs'],['S_OPEX','Solar opex','€m/MW'],['S_PPA','Solar PPA','€/MWh'],['S_TAIL','Solar tail','€/MWh'],
   ['B_GEAR','Battery gearing',''],['B_RATE','Battery debt rate',''],['B_GY','Battery merchant year',''],['B_COD','Battery COD',''],['B_LIFE','Battery life','yrs'],['B_DEGR','Battery degradation',''],['B_COMP','Battery compression',''],['B_OPEXP','Battery opex %',''],
   ['DC_MW','DC load','MW'],['DC_P','DC price','€/MWh'],['RES_P','Residual price','€/MWh'],['RES_M','Trading margin',''],['FEE_E','NE3 energy fee','€/MWh'],['FEE_C','NE3 capacity fee','€/kW/yr'],['FEESC','Grid-fee esc.','per yr'],['RES_OWN','RES owned (1/0)',''],['LINE_C','Direct line','€m/100MW'],['DCAC','Solar DC/AC',''],
   ['SPV_MODE','Pass-through (1/0)',''],['SPV_M','SPV margin','% of cost'],['MGN_MODE','Margin form (1=€/MWh)',''],['MGN_E','SPV margin','€/MWh'],['SPV_FF','SPV first year','']
  ].forEach(x=>{ rr=linkRow(ws,rr,L,x[0],x[1],x[2]); });
  const A=k=>`$B$${L[k]}`;

  rr++; rr=sect(ws,rr,'LINKED FROM BATTERY SHEET (red)');
  const bl=(key,label,unit)=>{ ws.getCell(rr,1).value=label; const c=ws.getCell(rr,2); c.value={formula:BX(key)}; c.font=RED; c.numFmt=numF; ws.getCell(rr,3).value=unit||''; L['BX_'+key]=rr; rr++; };
  bl('ARBRR','Battery arbitrage run-rate','€m/yr'); bl('ANC','Battery ancillary','€m/yr'); bl('DCCH','Battery DC charge','€m/yr');
  bl('GFCAP','Battery grid fee cap.','€m/yr'); bl('GFENE','Battery grid fee ene.','€m/yr'); bl('CELLCX','Battery cell capex','€m'); bl('TCAPEX','Battery total capex','€m');
  const BXL=k=>`$B$${L['BX_'+k]}`;

  rr++; rr=sect(ws,rr,'DERIVED (computed once)');
  rr=derRow(ws,rr,D,'RESCX', 'RES capex owned (€m)', `${A('RES_OWN')}*(${A('W_MW')}*${A('W_CAPEX')}+${A('S_MW')}*${A('S_CAPEX')})`, eurF);
  rr=derRow(ws,rr,D,'LINECX','Line capex (€m)', `(${A('W_MW')}+${A('S_MW')}/${A('DCAC')})*${A('LINE_C')}/100`, eurF,'export capacity: solar enters at AC');
  rr=derRow(ws,rr,D,'SPVCX', 'SPV total capex (€m)', `$B$${D.RESCX}+${BXL('TCAPEX')}+$B$${D.LINECX}`, eurF);
  rr=derRow(ws,rr,D,'IRRCX', 'Financed SPV capex, incl. line (€m)', `$B$${D.SPVCX}`, eurF);
  rr=derRow(ws,rr,D,'SENDBT','Senior RES + line debt base (€m)', `${A('GEAR')}*($B$${D.RESCX}+$B$${D.LINECX})`, eurF);
  rr=derRow(ws,rr,D,'BATDBT','Battery debt base (€m)', `${A('B_GEAR')}*${BXL('TCAPEX')}`, eurF);
  rr=derRow(ws,rr,D,'BGEAR', 'Blended SPV gearing, incl. line', `($B$${D.SENDBT}+$B$${D.BATDBT})/MAX(0.0001,$B$${D.IRRCX})`, '0.0000');
  rr=derRow(ws,rr,D,'BRATE', 'Blended SPV debt rate, incl. line', `(${A('GEAR')}*($B$${D.RESCX}+$B$${D.LINECX})*${A('RATE')}+${A('B_GEAR')}*${BXL('TCAPEX')}*${A('B_RATE')})/MAX(0.0001,$B$${D.SENDBT}+$B$${D.BATDBT})`, '0.0000');
  rr=derRow(ws,rr,D,'RESLIFE','RES life (yrs)', `MAX(${A('W_LIFE')},${A('S_LIFE')})`, intF);
  rr=derRow(ws,rr,D,'DEPY',  'Depreciation period (yrs)', `MIN(20,$B$${D.RESLIFE})`, intF);
  rr=derRow(ws,rr,D,'ANNDEP','Annual SPV depreciation (€m)', `$B$${D.IRRCX}/$B$${D.DEPY}`, numF);
  rr=derRow(ws,rr,D,'REPY',  'Debt amortisation period (yrs)', `MIN(${A('TENOR')},$B$${D.RESLIFE})`, intF);
  rr=derRow(ws,rr,D,'DRAW1', 'Constr. draw @ FF-2 (€m)', `0.3*($B$${D.RESCX}+$B$${D.LINECX})*$B$${D.BGEAR}`, numF);
  rr=derRow(ws,rr,D,'IDC1',  'Constr. IDC year 1 (€m)', `$B$${D.DRAW1}/2*$B$${D.BRATE}`, numF);
  rr=derRow(ws,rr,D,'BAL1',  'Debt after FF-2 (€m)', `$B$${D.DRAW1}+$B$${D.IDC1}`, numF);
  rr=derRow(ws,rr,D,'CAPFF1','SPV capex @ FF-1 (€m)', `0.7*($B$${D.RESCX}+$B$${D.LINECX})+${BXL('TCAPEX')}`, numF);
  rr=derRow(ws,rr,D,'DRAW2', 'Constr. draw @ FF-1 (€m)', `$B$${D.CAPFF1}*$B$${D.BGEAR}`, numF);
  rr=derRow(ws,rr,D,'IDC2',  'Constr. IDC year 2 (€m)', `($B$${D.BAL1}+$B$${D.DRAW2}/2)*$B$${D.BRATE}`, numF);
  rr=derRow(ws,rr,D,'DEBTFF','Debt drawn by SPV_FF (€m)', `$B$${D.BAL1}+$B$${D.DRAW2}+$B$${D.IDC2}`, numF);
  rr=derRow(ws,rr,D,'ANNPRIN','Annual principal, flat (€m)', `$B$${D.DEBTFF}/$B$${D.REPY}`, numF);
  rr=derRow(ws,rr,D,'ANNDS',  'Annual debt service, annuity (€m)', `PMT($B$${D.BRATE},$B$${D.REPY},-$B$${D.DEBTFF})`, numF,'level total payment');

  rr++; const Yrow=rr; yearHeader(ws,Yrow); const yr=X=>`${X}$${Yrow}`;
  rr++; rr=sect(ws,rr,'YEARLY MODEL');
  R.op   =tsRow(ws,rr++,'Operating flag (FF≤y<FF+RESlife)',Yrow,intF,X=>`(${yr(X)}>=${A('SPV_FF')})*(${yr(X)}<${A('SPV_FF')}+$B$${D.RESLIFE})`);
  R.dcrev=rr++; // filled below, once the pass-through cost base exists
  // battery revenue (mirror Battery sheet, referenced via battery links)
  R.bmf  =tsRow(ws,rr++,'Battery merchant flag',Yrow,intF,X=>`(${yr(X)}>=${A('B_GY')})*(${yr(X)}<${A('B_COD')}+${A('B_LIFE')})`);
  R.bcf  =tsRow(ws,rr++,'Battery op flag (y>COD)',Yrow,intF,X=>`(${yr(X)}>${A('B_COD')})*(${yr(X)}<${A('B_COD')}+${A('B_LIFE')})`);
  R.bmer =tsRow(ws,rr++,'Battery merchant rev (€m)',Yrow,numF,X=>`${X}$${R.bmf}*(${BXL('ARBRR')}*(1-${A('B_DEGR')})^MAX(0,${yr(X)}-${A('B_GY')})+${BXL('ANC')})*(1-${A('B_COMP')})^MAX(0,${yr(X)}-${A('B_GY')})*(1+${A('INFL')})^MAX(0,${yr(X)}-${A('B_GY')})`);
  R.bcap =tsRow(ws,rr++,'Battery DC charge (€m)',Yrow,numF,X=>`${X}$${R.bcf}*${BXL('DCCH')}*(1+${A('INFL')})^MAX(0,${yr(X)}-(${A('B_COD')}+1))`);
  R.brev =tsRow(ws,rr++,'Battery revenue (€m)',Yrow,numF,X=>`${X}$${R.bmer}+${X}$${R.bcap}`);
  R.wprod=tsRow(ws,rr++,'Wind → DC (MWh)',Yrow,intF,X=>`Wind!${X}$${windRef.prodRow}`);
  R.sprod=tsRow(ws,rr++,'Solar → DC (MWh)',Yrow,intF,X=>`Solar!${X}$${solarRef.prodRow}`);
  R.resm =tsRow(ws,rr++,'Residual from grid (MWh)',Yrow,intF,X=>`${X}$${R.op}*MAX(0,${A('DC_MW')}*8760-${X}$${R.wprod}-${X}$${R.sprod})`);
  R.resp =tsRow(ws,rr++,'Residual energy price (€/MWh)',Yrow,eurF,X=>`${A('RES_P')}*(1+${A('INFL')})^(${yr(X)}-2025)*(1+${A('RES_M')})+${A('FEE_E')}*(1+${A('FEESC')})^MAX(0,${yr(X)}-2028)`);
  R.resec=tsRow(ws,rr++,'Residual energy cost (€m)',Yrow,numF,X=>`${X}$${R.op}*${X}$${R.resm}*${X}$${R.resp}/10^6`);
  R.rescc=tsRow(ws,rr++,'Residual capacity fee (€m)',Yrow,numF,X=>`${X}$${R.op}*${A('DC_MW')}*${A('FEE_C')}/1000*(1+${A('FEESC')})^MAX(0,${yr(X)}-2028)`);
  R.resc =tsRow(ws,rr++,'Residual cost: total (€m)',Yrow,numF,X=>`${X}$${R.resec}+${X}$${R.rescc}`);
  R.wpf  =tsRow(ws,rr++,'Wind in-PPA flag',Yrow,intF,X=>`(${yr(X)}<${A('W_COD')}+MIN(${A('PPAT')},${A('W_LIFE')}))`);
  R.spf  =tsRow(ws,rr++,'Solar in-PPA flag',Yrow,intF,X=>`(${yr(X)}<${A('S_COD')}+MIN(${A('PPAT')},${A('S_LIFE')}))`);
  R.wpr  =tsRow(ws,rr++,'Wind price to SPV (€/MWh)',Yrow,eurF,X=>`${X}$${R.wpf}*${A('W_PPA')}+(1-${X}$${R.wpf})*${A('W_TAIL')}`);
  R.spr  =tsRow(ws,rr++,'Solar price to SPV (€/MWh)',Yrow,eurF,X=>`${X}$${R.spf}*${A('S_PPA')}+(1-${X}$${R.spf})*${A('S_TAIL')}`);
  R.resppa=tsRow(ws,rr++,'RES bought at PPA (€m, mode=0)',Yrow,numF,X=>`${X}$${R.op}*(1-${A('RES_OWN')})*(${X}$${R.wpr}*${X}$${R.wprod}+${X}$${R.spr}*${X}$${R.sprod})/10^6`);
  R.resown=tsRow(ws,rr++,'RES billed at own LCOE (€m, mode=1)',Yrow,numF,X=>`${X}$${R.op}*${A('RES_OWN')}*(${windRef.lcoe}*${X}$${R.wprod}+${solarRef.lcoe}*${X}$${R.sprod})/10^6`);
  R.pasee=tsRow(ws,rr++,'Pass-through energy base (€m)',Yrow,numF,X=>`${X}$${R.resppa}+${X}$${R.resown}+${X}$${R.resec}`);
  R.mgn  =tsRow(ws,rr++,'SPV margin (€m)',Yrow,numF,X=>`${X}$${R.op}*(${A('MGN_MODE')}*${A('DC_MW')}*8760*${A('MGN_E')}*(1+${A('INFL')})^(${yr(X)}-2026)/10^6+(1-${A('MGN_MODE')})*${X}$${R.pasee}*${A('SPV_M')})`);
  // Pass-through: energy at cost + regulated network charge at cost + the margin. Fixed: a flat €/MWh.
  tsRow(ws,R.dcrev,'DC revenue (€m)',Yrow,numF,X=>`${X}$${R.op}*(${A('SPV_MODE')}*(${X}$${R.pasee}+${X}$${R.rescc}+${X}$${R.mgn})+(1-${A('SPV_MODE')})*${A('DC_MW')}*8760*${A('DC_P')}*(1+${A('INFL')})^(${yr(X)}-2026)/10^6)`);
  R.cpi  =tsRow(ws,rr++,'CPI index vs 2023',Yrow,'0.0000',X=>`(1+${A('INFL')})^(${yr(X)}-2023)`);
  R.resox=tsRow(ws,rr++,'RES opex owned (€m)',Yrow,numF,X=>`${X}$${R.op}*${A('RES_OWN')}*(${A('W_OPEX')}*${A('W_MW')}+${A('S_OPEX')}*${A('S_MW')})*${X}$${R.cpi}`);
  R.batox=tsRow(ws,rr++,'Battery cell opex (€m)',Yrow,numF,X=>`${X}$${R.op}*${BXL('CELLCX')}*${A('B_OPEXP')}*${X}$${R.cpi}`);
  R.batgf=tsRow(ws,rr++,'Battery grid fees (€m)',Yrow,numF,X=>`${X}$${R.op}*(${yr(X)}>=${A('B_GY')})*(${BXL('GFCAP')}+${BXL('GFENE')})*(1+${A('FEESC')})^MAX(0,${yr(X)}-2028)`);
  R.opex =tsRow(ws,rr++,'Opex: total (€m)',Yrow,numF,X=>`${X}$${R.resox}+${X}$${R.batox}+${X}$${R.batgf}`);
  R.ebit =tsRow(ws,rr++,'EBITDA (€m)',Yrow,numF,X=>`${X}$${R.dcrev}+${X}$${R.brev}-${X}$${R.resc}-${X}$${R.resppa}-${X}$${R.opex}`);
  R.cshr =tsRow(ws,rr++,'RES + line capex draw share',Yrow,'0.00',X=>`0.3*(${yr(X)}=${A('SPV_FF')}-2)+0.7*(${yr(X)}=${A('SPV_FF')}-1)`);
  R.capex=tsRow(ws,rr++,'SPV capex, incl. direct line (€m)',Yrow,numF,X=>`${X}$${R.cshr}*($B$${D.RESCX}+$B$${D.LINECX})+(${yr(X)}=${A('SPV_FF')}-1)*${BXL('TCAPEX')}`);
  R.draw =tsRow(ws,rr++,'Debt draw (€m)',Yrow,numF,X=>`${X}$${R.capex}*$B$${D.BGEAR}`);
  R.cflag=tsRow(ws,rr++,'Construction flag',Yrow,intF,X=>`(${yr(X)}>=${A('SPV_FF')}-2)*(${yr(X)}<${A('SPV_FF')})`);
  R.idc=rr++; R.intr=rr++; R.prin=rr++; R.bal=rr++; R.dep=rr++; R.ebt=rr++; R.nol=rr++; R.tax=rr++; R.fcfe=rr++; R.date=rr++; R.xcf=rr++;
  tsRow(ws,R.idc,'IDC: constr. interest (€m)',Yrow,numF,(X,pX)=>`${X}$${R.cflag}*((${pX?pX+'$'+R.bal:'0'})+${X}$${R.draw}/2)*$B$${D.BRATE}`);
  tsRow(ws,R.intr,'Interest (€m)',Yrow,numF,(X,pX)=>`(${yr(X)}>=${A('SPV_FF')})*(${pX?pX+'$'+R.bal:'0'})*$B$${D.BRATE}`);
  tsRow(ws,R.prin,'Principal repay (€m)',Yrow,numF,(X,pX)=>`(${yr(X)}>=${A('SPV_FF')})*(${yr(X)}<${A('SPV_FF')}+$B$${D.REPY})*MIN((${pX?pX+'$'+R.bal:'0'}),${A('AMORT')}*MAX(0,$B$${D.ANNDS}-${X}$${R.intr})+(1-${A('AMORT')})*$B$${D.ANNPRIN})`);
  tsRow(ws,R.bal,'Debt balance EOY (€m)',Yrow,numF,(X,pX)=>`(${yr(X)}<${A('SPV_FF')})*((${pX?pX+'$'+R.bal:'0'})+${X}$${R.draw}+${X}$${R.idc})+(${yr(X)}>=${A('SPV_FF')})*MAX(0,(${pX?pX+'$'+R.bal:'0'})-${X}$${R.prin})`);
  tsRow(ws,R.dep,'Depreciation (€m)',Yrow,numF,X=>`(${yr(X)}>=${A('SPV_FF')})*(${yr(X)}<${A('SPV_FF')}+$B$${D.DEPY})*$B$${D.ANNDEP}`);
  tsRow(ws,R.ebt,'EBT (€m)',Yrow,numF,X=>`${X}$${R.ebit}-${X}$${R.dep}-${X}$${R.intr}`);
  tsRow(ws,R.nol,'NOL balance (€m)',Yrow,numF,(X,pX)=>`(${yr(X)}>=${A('SPV_FF')})*MIN(0,(${pX?pX+'$'+R.nol:'0'})+${X}$${R.ebt})`);
  tsRow(ws,R.tax,'Tax (€m)',Yrow,numF,(X,pX)=>`(${yr(X)}>=${A('SPV_FF')})*MAX(0,${X}$${R.ebt}+(${pX?pX+'$'+R.nol:'0'}))*${A('TAXR')}`);
  tsRow(ws,R.fcfe,'EQUITY CASH FLOW (€m)',Yrow,numF,X=>`(${yr(X)}<${A('SPV_FF')})*(-${X}$${R.capex}+${X}$${R.draw})+(${yr(X)}>=${A('SPV_FF')})*(${X}$${R.ebit}-${X}$${R.tax}-${X}$${R.prin}-${X}$${R.intr})`,{bold:true});
  for(let i=0;i<NY;i++){const c=ws.getCell(R.date,c0+i);c.value={formula:`DATE(${colL(c0+i)}$${Yrow},12,31)`};c.numFmt='yyyy-mm-dd';}
  ws.getCell(R.date,1).value='Date (XIRR)';
  ws.getCell(R.fcfe,1).fill=HDR; ws.getCell(R.fcfe,1).font=white;
  ws.getCell(R.xcf,1).value='Equity CF for XIRR (dummy seed − t0, immaterial)';
  ws.getCell(R.xcf,c0).value=-0.01; ws.getCell(R.xcf,c0).numFmt=numF;
  for(let i=1;i<NY;i++){const c=ws.getCell(R.xcf,c0+i);c.value={formula:`${colL(c0+i)}${R.fcfe}`};c.numFmt=numF;}

  let k=rr+1; k=sect(ws,k,'RESULTS');
  const IRR=k; ws.getCell(k,1).value='SPV equity IRR, incl. direct-line capex (XIRR)'; ws.getCell(k,2).value={formula:`IFERROR(XIRR(B${R.xcf}:${lastC}${R.xcf},B${R.date}:${lastC}${R.date}),"n/m")`}; ws.getCell(k,2).numFmt=pctF; ws.getCell(k,2).font=bold; k++;
  ws.getCell(k,1).value='Dashboard SPV IRR at export'; ws.getCell(k,2).value=S.spvIRR; ws.getCell(k,2).numFmt=pctF; ws.getCell(k,2).fill=CHK; k++;
  ws.getCell(k,1).value='Financed SPV capex, incl direct line (€m)'; ws.getCell(k,2).value={formula:`$B$${D.IRRCX}`}; ws.getCell(k,2).numFmt=eurF; k++;
  ws.getCell(k,1).value='SPV total capex, incl direct line (€m)'; ws.getCell(k,2).value={formula:`$B$${D.SPVCX}`}; ws.getCell(k,2).numFmt=eurF; k++;
  ws.getCell(k,1).value='Method note'; ws.getCell(k,2).value='Asset-only returns are shown on the Wind, Solar and Battery sheets. Battery arbitrage uses exported day-by-day backtest averages from the selected duration and year. Re-export after changing those dimensions. Any non-zero tie-out identifies a model reconciliation issue.';
  return {irr:`SPV!B${IRR}`};
 }
 // Asset builders must expose their production row for the SPV pull; patch refs:
 // windRef/solarRef were created above. Attach prodRow by re-reading; storing during build is simpler.
 const spvRef=spvSheet();

 // ================= OUTPUT =================
 const wo=wb.addWorksheet('Output'); wo.getColumn(1).width=30; wo.getColumn(2).width=16;
 wo.getCell(1,1).value='OUTPUT DASHBOARD'; wo.getCell(1,1).font={bold:true,size:13};
 let ro=3;
 [['Wind equity IRR',windRef.irr,pctF],['Wind MOIC',windRef.moic,xF],['Wind LCOE (€/MWh)',windRef.lcoe,eurF],
  ['Solar equity IRR',solarRef.irr,pctF],['Solar MOIC',solarRef.moic,xF],['Solar LCOE (€/MWh)',solarRef.lcoe,eurF],
  ['Battery equity IRR',batteryRef.irr,pctF],['SPV equity IRR, incl. direct-line capex',spvRef.irr,pctF]
 ].forEach(o=>{ wo.getCell(ro,1).value=o[0]; wo.getCell(ro,2).value={formula:o[1]}; wo.getCell(ro,2).numFmt=o[2]; wo.getCell(ro,2).font=bold; ro++; });
 wo.getCell(ro+1,1).value='All inputs live on the Inputs sheet (yellow). Change one → every sheet recalculates.';
 wo.getCell(ro+2,1).value='Each calc sheet: red rows = linked inputs (F2 to trace); helper rows show every step; no IF, no named ranges.';
 return wb;
}
if(typeof module!=='undefined')module.exports={buildFullModel};
