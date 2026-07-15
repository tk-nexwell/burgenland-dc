// Full-formula Excel model generator — runs identically in Node (test) and browser (ExcelJS).
// buildFullModel(ExcelJSlib, S) -> workbook. S = state snapshot + check rows from the live dashboard.
function buildFullModel(ExcelJS, S){
 const wb=new ExcelJS.Workbook(); wb.creator='Nexwell Power'; wb.created=new Date();
 const Y0=2026, YN=2060, NY=YN-Y0+1;
 const YEL={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFF2CC'}};   // input yellow
 const HDR={type:'pattern',pattern:'solid',fgColor:{argb:'FF1F3864'}};   // header navy
 const CHK={type:'pattern',pattern:'solid',fgColor:{argb:'FFE2EFDA'}};   // check green
 const bold={bold:true}, white={color:{argb:'FFFFFFFF'},bold:true};
 const numF='#,##0.00;[Red]-#,##0.00', pctF='0.00%', eurF='#,##0.0';
 function inputRow(ws,r,label,value,unit,name){
  ws.getCell(r,1).value=label; ws.getCell(r,2).value=value; ws.getCell(r,2).fill=YEL; ws.getCell(r,2).border={outline:{style:'thin'}};
  ws.getCell(r,3).value=unit||'';
  if(name) wb.definedNames.add(`Inputs!$B$${r}`,name);
  return r+1;
 }
 // ---------------- INPUTS ----------------
 const wi=wb.addWorksheet('Inputs'); wi.getColumn(1).width=34; wi.getColumn(2).width=14; wi.getColumn(3).width=40;
 wi.getCell(1,1).value='GDC NICKELSDORF — POWER SPV MODEL'; wi.getCell(1,1).font={bold:true,size:14};
 wi.getCell(2,1).value='Yellow cells = inputs. All calculations are live Excel formulas. Exported from the dashboard '+S.today+'.';
 let r=4;
 wi.getCell(r,1).value='MACRO / FINANCING'; wi.getCell(r,1).font=bold; r++;
 r=inputRow(wi,r,'Inflation (CPI)',S.macro.infl,'per yr','INFL');
 r=inputRow(wi,r,'Tax rate',S.macro.tax,'','TAXR');
 r=inputRow(wi,r,'Gearing (wind/solar/line)',S.macro.gearing,'','GEAR');
 r=inputRow(wi,r,'Debt tenor',S.macro.tenor,'years','TENOR');
 r=inputRow(wi,r,'All-in debt rate',S.macro.allInRate,'','RATE');
 r=inputRow(wi,r,'PPA term',S.macro.ppaTermY,'years','PPAT');
 r=inputRow(wi,r,'Merchant power (uncontracted RES)',S.macro.merchReal,'€/MWh 2023-real, CPI-indexed','MERCH');
 r++; wi.getCell(r,1).value='WIND'; wi.getCell(r,1).font=bold; r++;
 r=inputRow(wi,r,'Capacity',S.wind.mw,'MW AC','W_MW');
 r=inputRow(wi,r,'Capex',S.wind.capexPerMW,'€m per MW','W_CAPEX');
 r=inputRow(wi,r,'Gross capacity factor',S.wind.grossCF,'','W_GCF');
 r=inputRow(wi,r,'Plant losses (wake, electrical)',S.wind.loss,'','W_LOSS');
 r=inputRow(wi,r,'Direct-line losses',S.wind.lineLoss,'BE 15-Jul: 4.0% @ 12.5 km','W_LLOSS');
 r=inputRow(wi,r,'Degradation',S.wind.degr,'per yr','W_DEGR');
 r=inputRow(wi,r,'Opex',S.wind.opexPerMW,'€m/MW/yr, CPI-indexed (excl land rent TBC)','W_OPEX');
 r=inputRow(wi,r,'PPA price',S.wind.ppa,'€/MWh fixed','W_PPA');
 r=inputRow(wi,r,'Contracted share',S.wind.contr,'','W_CONTR');
 r=inputRow(wi,r,'COD (first generation year)',S.wind.codY,'capacity-weighted, BE list','W_COD');
 r=inputRow(wi,r,'Useful life',S.wind.lifeY,'years','W_LIFE');
 r=inputRow(wi,r,'Merchant tail price',S.CAP7w,'€/MWh flat (capture avg ex-2022)','W_TAIL');
 r++; wi.getCell(r,1).value='SOLAR'; wi.getCell(r,1).font=bold; r++;
 r=inputRow(wi,r,'Capacity',S.solar.mw,'MWp DC','S_MW');
 r=inputRow(wi,r,'Capex',S.solar.capexPerMW,'€m per MWp','S_CAPEX');
 r=inputRow(wi,r,'Gross capacity factor',S.solar.grossCF,'','S_GCF');
 r=inputRow(wi,r,'Plant losses (soiling, inverter)',S.solar.loss,'','S_LOSS');
 r=inputRow(wi,r,'Direct-line losses',S.solar.lineLoss,'','S_LLOSS');
 r=inputRow(wi,r,'Degradation',S.solar.degr,'per yr','S_DEGR');
 r=inputRow(wi,r,'Opex',S.solar.opexPerMW,'€m/MWp/yr, CPI-indexed','S_OPEX');
 r=inputRow(wi,r,'PPA price',S.solar.ppa,'€/MWh fixed','S_PPA');
 r=inputRow(wi,r,'Contracted share',S.solar.contr,'','S_CONTR');
 r=inputRow(wi,r,'COD (first generation year)',S.solar.codY,'','S_COD');
 r=inputRow(wi,r,'Useful life',S.solar.lifeY,'years','S_LIFE');
 r=inputRow(wi,r,'Merchant tail price',S.CAP7s,'€/MWh flat','S_TAIL');
 r++; wi.getCell(r,1).value='BATTERY'; wi.getCell(r,1).font=bold; r++;
 r=inputRow(wi,r,'Power',S.battery.powerMW,'MW','B_MW');
 r=inputRow(wi,r,'Duration',S.battery.durationH,'hours','B_DUR');
 r=inputRow(wi,r,'Cell capex',S.battery.capexPerKWh,'€/kWh','B_CKWH');
 r=inputRow(wi,r,'NEB Netzzutritt (switching station)',S.battery.substation,'€m — NEB est., actual-cost','B_SUB');
 r=inputRow(wi,r,'DC connection (customer 3×3.5 km)',S.battery.interconnect,'€m','B_INT');
 r=inputRow(wi,r,'Round-trip efficiency',S.battery.rte,'','B_RTE');
 r=inputRow(wi,r,'SoC floor (outage reserve)',S.battery.socFloor,'not traded','B_SOC');
 r=inputRow(wi,r,'Cycles per day',S.battery.cyclesDay,'','B_CYC');
 r=inputRow(wi,r,'Capture factor',S.battery.captureFactor,'x volatility uplift','B_CAPF');
 r=inputRow(wi,r,'Opex',S.battery.opexPct,'% of cell capex, CPI-indexed','B_OPEXP');
 r=inputRow(wi,r,'Ancillary revenue',S.battery.ancPerMW,'€k/MW/yr (from merchant year)','B_ANC');
 r=inputRow(wi,r,'DC capacity charge (TW pays)',S.battery.capChargeMWyr,'€k/MW/yr from COD+1 — internal lever','B_CCH');
 r=inputRow(wi,r,'First merchant year (export ban)',S.battery.gridYear,'BE: ban to 2035','B_GY');
 r=inputRow(wi,r,'Revenue compression',S.battery.compression,'per yr on merchant','B_COMP');
 r=inputRow(wi,r,'Degradation',S.battery.degr,'per yr','B_DEGR');
 r=inputRow(wi,r,'Gearing',S.battery.gearing,'','B_GEAR');
 r=inputRow(wi,r,'Debt rate',S.battery.debtRate,'','B_RATE');
 r=inputRow(wi,r,'Grid fee — capacity (NE3)',S.battery.gridCapFee,'€/kW/yr, esc w/ FEESC from 2028','B_GFC');
 r=inputRow(wi,r,'Grid fee — energy (NE3)',S.battery.gridEnergyFee,'€/MWh on grid throughput','B_GFE');
 r=inputRow(wi,r,'Battery COD (capex year)',S.COD,'','B_COD');
 r=inputRow(wi,r,'Useful life',S.battery.lifeY,'years','B_LIFE');
 r++; wi.getCell(r,1).value='DATA CENTER / RESIDUAL (via BE Trading)'; wi.getCell(r,1).font=bold; r++;
 r=inputRow(wi,r,'DC contracted load',S.dc.firmMW,'MW','DC_MW');
 r=inputRow(wi,r,'DC sale price (today)',S.dc.dcPrice,'€/MWh, CPI-indexed','DC_P');
 r=inputRow(wi,r,'Residual market price',S.dc.resFix,'€/MWh 2025-real, CPI-indexed','RES_P');
 r=inputRow(wi,r,'BE Trading margin',S.dc.beMargin,'','RES_M');
 r=inputRow(wi,r,'NE3 energy fee',S.dc.gridEnergyFee,'€/MWh (2028 level)','FEE_E');
 r=inputRow(wi,r,'NE3 capacity fee',S.dc.gridCapFeeKW,'€/kW/yr (2028 level)','FEE_C');
 r=inputRow(wi,r,'Grid-fee escalation',S.dc.feeEsc,'per yr from 2028 (BE Trading path)','FEESC');
 r=inputRow(wi,r,'RES sourcing mode (1=owned, 0=PPA)',(S.dc.resMode||'lcoe')==='lcoe'?1:0,'drives SPV sheet','RES_OWN');
 r=inputRow(wi,r,'Direct line cost',S.linePer100,'€m per 100 MW','LINE_C');
 r=inputRow(wi,r,'SPV first revenue year',S.FF,'','SPV_FF');
 // battery price curve block
 r++; wi.getCell(r,1).value='BATTERY ARBITRAGE CURVE — hourly wholesale, year '+S.priceYear+' (€/MWh)'; wi.getCell(r,1).font=bold; r++;
 const cst=r;
 for(let h=0;h<24;h++){wi.getCell(r,1).value='Hour '+h; wi.getCell(r,2).value=S.ph[h]; wi.getCell(r,2).fill=YEL;
  wi.getCell(r,3).value={formula:`RANK($B$${r},$B$${cst}:$B$${cst+23},1)`}; r++;}
 wb.definedNames.add(`Inputs!$B$${cst}:$B$${cst+23}`,'B_PH'); wb.definedNames.add(`Inputs!$C$${cst}:$C$${cst+23}`,'B_RANK');
 wi.getCell(cst-1,3).value='rank (1=cheapest)';

 // ---------------- ASSET SHEET BUILDER (wind/solar) ----------------
 function assetSheet(name,P,chk){ // P: prefix 'W' or 'S'
  const ws=wb.addWorksheet(name); ws.getColumn(1).width=30;
  const c0=2; // first year column
  ws.getCell(1,1).value=name.toUpperCase()+' — full formula model'; ws.getCell(1,1).font=bold;
  ws.getCell(2,1).value='Year'; for(let i=0;i<NY;i++){const c=ws.getCell(2,c0+i); c.value=Y0+i; c.font=white; c.fill=HDR;}
  const R={yr:2,op:3,prod:4,rev:5,opex:6,ebitda:7,capex:8,dep:9,draw:10,idc:11,bal:12,intr:13,rep:14,ebt:15,nol:16,tax:17,fcfe:18,date:19,chk:20,diff:21};
  const L={op:'Operating flag',prod:'Production (MWh)',rev:'Revenue (€m)',opex:'Opex (€m)',ebitda:'EBITDA (€m)',capex:'Capex (€m)',dep:'Depreciation (€m)',draw:'Debt draw (€m)',idc:'IDC (€m)',bal:'Debt balance EOY (€m)',intr:'Interest (€m)',rep:'Principal repay (€m)',ebt:'EBT (€m)',nol:'NOL balance (€m)',tax:'Tax (€m)',fcfe:'EQUITY CASH FLOW (€m)',date:'Date (XIRR)',chk:'Check: dashboard equity CF',diff:'Check diff (should be ~0)'};
  Object.keys(L).forEach(k=>{ws.getCell(R[k],1).value=L[k]; if(k==='fcfe')ws.getCell(R[k],1).font=bold;});
  for(let i=0;i<NY;i++){
   const C=ws.getColumn(c0+i).letter, y=Y0+i, P_=P+'_';
   const cell=(row,f)=>{ws.getCell(row,c0+i).value={formula:f}; ws.getCell(row,c0+i).numFmt=numF;};
   cell(R.op,`IF(AND(${C}2>=${P_}COD,${C}2<${P_}COD+${P_}LIFE),1,0)`);
   cell(R.prod,`${C}${R.op}*${P_}MW*8760*${P_}GCF*(1-${P_}LOSS)*(1-${P_}LLOSS)*(1-${P_}DEGR)^(${C}2-${P_}COD)`);
   cell(R.rev,`${C}${R.op}*IF(${C}2<${P_}COD+MIN(PPAT,${P_}LIFE),(${P_}CONTR*${C}${R.prod}*${P_}PPA+(1-${P_}CONTR)*${C}${R.prod}*MERCH*(1+INFL)^(${C}2-2023)),${C}${R.prod}*${P_}TAIL)/10^6`);
   cell(R.opex,`${C}${R.op}*${P_}OPEX*${P_}MW*(1+INFL)^(${C}2-2023)`);
   cell(R.ebitda,`${C}${R.rev}-${C}${R.opex}`);
   cell(R.capex,`IF(${C}2=${P_}COD-2,0.3,IF(${C}2=${P_}COD-1,0.7,0))*${P_}MW*${P_}CAPEX`);
   cell(R.dep,`IF(AND(${C}2>=${P_}COD,${C}2<${P_}COD+MIN(20,${P_}LIFE)),${P_}MW*${P_}CAPEX/MIN(20,${P_}LIFE),0)`);
   cell(R.draw,`IF(${C}2<${P_}COD,GEAR*${C}${R.capex},0)`);
   const prevBal=i===0?'0':ws.getColumn(c0+i-1).letter+R.bal;
   cell(R.idc,`IF(AND(${C}2<${P_}COD,${C}${R.capex}>0),(${prevBal}+${C}${R.draw}/2)*RATE,0)`);
   cell(R.bal,`IF(${C}2<${P_}COD,${prevBal}+${C}${R.draw}+${C}${R.idc},MAX(0,${prevBal}-${C}${R.rep}))`);
   cell(R.intr,`IF(${C}2>=${P_}COD,${prevBal}*RATE,0)`);
   // ppy = balance at end of COD-1 divided by min(tenor,life)
   cell(R.rep,`IF(AND(${C}2>=${P_}COD,${C}2<${P_}COD+MIN(TENOR,${P_}LIFE)),MIN(${prevBal},INDEX($B$${R.bal}:$${ws.getColumn(c0+NY-1).letter}$${R.bal},MATCH(${P_}COD-1,$B$2:$${ws.getColumn(c0+NY-1).letter}$2,0))/MIN(TENOR,${P_}LIFE)),0)`);
   cell(R.ebt,`${C}${R.ebitda}-${C}${R.dep}-${C}${R.intr}`);
   const prevNol=i===0?'0':ws.getColumn(c0+i-1).letter+R.nol;
   cell(R.nol,`IF(${C}2<${P_}COD,0,MIN(0,${prevNol}+${C}${R.ebt}))`);
   cell(R.tax,`IF(${C}2<${P_}COD,0,MAX(0,${C}${R.ebt}+${prevNol})*TAXR)`);
   cell(R.fcfe,`IF(${C}2<${P_}COD,-${C}${R.capex}+${C}${R.draw},${C}${R.ebitda}-${C}${R.tax}-${C}${R.rep}-${C}${R.intr})`);
   ws.getCell(R.date,c0+i).value={formula:`DATE(${C}2,12,31)`}; ws.getCell(R.date,c0+i).numFmt='yyyy-mm-dd';
   ws.getCell(R.chk,c0+i).value=(chk[y]!==undefined?chk[y]:0); ws.getCell(R.chk,c0+i).fill=CHK; ws.getCell(R.chk,c0+i).numFmt=numF;
   cell(R.diff,`${C}${R.fcfe}-${C}${R.chk}`);
  }
  ws.getCell(R.fcfe,1).fill=HDR; ws.getCell(R.fcfe,1).font=white;
  const lastC=ws.getColumn(c0+NY-1).letter;
  const K=23;
  ws.getCell(K,1).value='RESULTS'; ws.getCell(K,1).font=bold;
  ws.getCell(K+1,1).value='Equity IRR (XIRR)'; ws.getCell(K+1,2).value={formula:`XIRR(B${R.fcfe}:${lastC}${R.fcfe},B${R.date}:${lastC}${R.date})`}; ws.getCell(K+1,2).numFmt=pctF; ws.getCell(K+1,2).font=bold;
  ws.getCell(K+2,1).value='MOIC'; ws.getCell(K+2,2).value={formula:`SUMIF(B${R.fcfe}:${lastC}${R.fcfe},">0")/-SUMIF(B${R.fcfe}:${lastC}${R.fcfe},"<0")`}; ws.getCell(K+2,2).numFmt='0.00"x"';
  ws.getCell(K+3,1).value='Total capex (€m)'; ws.getCell(K+3,2).value={formula:`${P}_MW*${P}_CAPEX`}; ws.getCell(K+3,2).numFmt=eurF;
  ws.getCell(K+4,1).value='Equity (€m)'; ws.getCell(K+4,2).value={formula:`(1-GEAR)*${P}_MW*${P}_CAPEX`}; ws.getCell(K+4,2).numFmt=eurF;
  ws.getCell(K+5,1).value='LCOE (€/MWh)'; ws.getCell(K+5,2).value={formula:`(${P}_MW*${P}_CAPEX*RATE/(1-(1+RATE)^-20)+INDEX(B${R.opex}:${lastC}${R.opex},MATCH(${P}_COD,B2:${lastC}2,0)))*10^6/INDEX(B${R.prod}:${lastC}${R.prod},MATCH(${P}_COD,B2:${lastC}2,0))`}; ws.getCell(K+5,2).numFmt=eurF;
  ws.getCell(K+6,1).value='Max |check diff| (€m)'; ws.getCell(K+6,2).value={formula:`MAX(ABS(MIN(B${R.diff}:${lastC}${R.diff})),ABS(MAX(B${R.diff}:${lastC}${R.diff})))`}; ws.getCell(K+6,2).numFmt=numF; ws.getCell(K+6,2).fill=CHK;
  return ws;
 }
 assetSheet('Wind','W',S.chkWind);
 assetSheet('Solar','S',S.chkSolar);

 // ---------------- BATTERY SHEET ----------------
 (function(){
  const ws=wb.addWorksheet('Battery'); ws.getColumn(1).width=32; const c0=2;
  ws.getCell(1,1).value='BATTERY — full formula model'; ws.getCell(1,1).font=bold;
  // dispatch economics (single block, from inputs + curve)
  const D=3;
  ws.getCell(D,1).value='Tradeable hours nCh = ROUND(duration×(1−SoC floor)×cycles)'; ws.getCell(D,2).value={formula:'MIN(12,MAX(0,ROUND(B_DUR*(1-B_SOC)*B_CYC,0)))'};
  wb.definedNames.add(`Battery!$B$${D}`,'B_NCH');
  ws.getCell(D+1,1).value='Sum of cheapest nCh prices (€)'; ws.getCell(D+1,2).value={formula:'SUMIF(B_RANK,"<="&B_NCH,B_PH)'};
  ws.getCell(D+2,1).value='Sum of priciest nCh prices (€)'; ws.getCell(D+2,2).value={formula:'SUMIF(B_RANK,">"&(24-B_NCH),B_PH)'};
  ws.getCell(D+3,1).value='Day arbitrage (€/day)'; ws.getCell(D+3,2).value={formula:`MAX(0,B_MW*(B_RTE*B$${D+2}-B$${D+1}))*B_CAPF`};
  ws.getCell(D+4,1).value='Arbitrage revenue (€m/yr, run-rate)'; ws.getCell(D+4,2).value={formula:`B${D+3}*365/10^6`};
  ws.getCell(D+5,1).value='Ancillary (€m/yr)'; ws.getCell(D+5,2).value={formula:'B_ANC*B_MW/1000'};
  ws.getCell(D+6,1).value='DC capacity charge (€m/yr)'; ws.getCell(D+6,2).value={formula:'B_CCH*B_MW/1000'};
  ws.getCell(D+7,1).value='Grid fee capacity (€m/yr, 2028 base)'; ws.getCell(D+7,2).value={formula:'B_MW*B_GFC/1000'};
  ws.getCell(D+8,1).value='Grid throughput (MWh/yr)'; ws.getCell(D+8,2).value={formula:'B_NCH*B_MW*B_RTE*365'};
  ws.getCell(D+9,1).value='Grid fee energy (€m/yr, 2028 base)'; ws.getCell(D+9,2).value={formula:`B${D+8}*B_GFE/10^6`};
  ws.getCell(D+10,1).value='Total capex (€m)'; ws.getCell(D+10,2).value={formula:'B_MW*B_DUR*1000*B_CKWH/10^6+B_SUB+B_INT'};
  wb.definedNames.add(`Battery!$B$${D+10}`,'B_CPX');
  for(let i=0;i<=10;i++)ws.getCell(D+i,2).numFmt=numF;
  // finance rows
  const F0=D+13;
  ws.getCell(F0-1,1).value='Year'; for(let i=0;i<NY;i++){const c=ws.getCell(F0-1,c0+i); c.value=Y0+i; c.font=white; c.fill=HDR;}
  const R={yr:F0-1,rev:F0,opex:F0+1,dep:F0+2,bal:F0+3,intr:F0+4,rep:F0+5,ebt:F0+6,nol:F0+7,tax:F0+8,fcfe:F0+9,date:F0+10,chk:F0+11,diff:F0+12};
  const L={rev:'Revenue (€m)',opex:'Opex + grid fees (€m)',dep:'Depreciation (€m)',bal:'Debt balance EOY (€m)',intr:'Interest (€m)',rep:'Principal repay (€m)',ebt:'EBT (€m)',nol:'NOL balance',tax:'Tax (€m)',fcfe:'EQUITY CASH FLOW (€m)',date:'Date (XIRR)',chk:'Check: dashboard equity CF',diff:'Check diff'};
  Object.keys(L).forEach(k=>{ws.getCell(R[k],1).value=L[k]; if(k==='fcfe'){ws.getCell(R[k],1).font=white;ws.getCell(R[k],1).fill=HDR;}});
  for(let i=0;i<NY;i++){
   const C=ws.getColumn(c0+i).letter, cell=(row,f)=>{ws.getCell(row,c0+i).value={formula:f}; ws.getCell(row,c0+i).numFmt=numF;};
   const prevBal=i===0?'0':ws.getColumn(c0+i-1).letter+R.bal, prevNol=i===0?'0':ws.getColumn(c0+i-1).letter+R.nol;
   cell(R.rev,`IF(AND(${C}${R.yr}>B_COD,${C}${R.yr}<B_COD+B_LIFE),IF(${C}${R.yr}>=B_GY,(B$${D+4}*(1-B_DEGR)^(${C}${R.yr}-B_GY)+B$${D+5})*(1-B_COMP)^(${C}${R.yr}-B_GY)*(1+INFL)^(${C}${R.yr}-B_GY),0)+B$${D+6}*(1+INFL)^(${C}${R.yr}-(B_COD+1)),0)`);
   cell(R.opex,`IF(AND(${C}${R.yr}>B_COD,${C}${R.yr}<B_COD+B_LIFE),B_MW*B_DUR*1000*B_CKWH/10^6*B_OPEXP*(1+INFL)^(${C}${R.yr}-2023)+IF(${C}${R.yr}>=B_GY,(B$${D+7}+B$${D+9})*(1+FEESC)^MAX(0,${C}${R.yr}-2028),0),0)`);
   cell(R.dep,`IF(AND(${C}${R.yr}>B_COD,${C}${R.yr}<=B_COD+MIN(15,B_LIFE-1)),B_CPX/MIN(15,B_LIFE-1),0)`);
   cell(R.bal,`IF(${C}${R.yr}<B_COD,0,IF(${C}${R.yr}=B_COD,B_GEAR*B_CPX,MAX(0,${prevBal}-${C}${R.rep})))`);
   cell(R.intr,`IF(${C}${R.yr}>B_COD,${prevBal}*B_RATE,0)`);
   cell(R.rep,`IF(AND(${C}${R.yr}>B_COD,${C}${R.yr}<=B_COD+MIN(TENOR,B_LIFE-1)),MIN(${prevBal},B_GEAR*B_CPX/MIN(TENOR,B_LIFE-1)),0)`);
   cell(R.ebt,`${C}${R.rev}-${C}${R.opex}-${C}${R.dep}-${C}${R.intr}`);
   cell(R.nol,`IF(${C}${R.yr}<=B_COD,0,MIN(0,${prevNol}+${C}${R.ebt}))`);
   cell(R.tax,`IF(${C}${R.yr}<=B_COD,0,MAX(0,${C}${R.ebt}+${prevNol})*TAXR)`);
   cell(R.fcfe,`IF(${C}${R.yr}=B_COD,-B_CPX+B_GEAR*B_CPX,IF(AND(${C}${R.yr}>B_COD,${C}${R.yr}<B_COD+B_LIFE),${C}${R.rev}-${C}${R.opex}-${C}${R.tax}-${C}${R.rep}-${C}${R.intr},0))`);
   ws.getCell(R.date,c0+i).value={formula:`DATE(${C}${R.yr},12,31)`}; ws.getCell(R.date,c0+i).numFmt='yyyy-mm-dd';
   ws.getCell(R.chk,c0+i).value=(S.chkBatt[Y0+i]!==undefined?S.chkBatt[Y0+i]:0); ws.getCell(R.chk,c0+i).fill=CHK; ws.getCell(R.chk,c0+i).numFmt=numF;
   cell(R.diff,`${C}${R.fcfe}-${C}${R.chk}`);
  }
  const lastC=ws.getColumn(c0+NY-1).letter, K=R.diff+2;
  ws.getCell(K,1).value='Equity IRR (XIRR)'; ws.getCell(K,2).value={formula:`IFERROR(XIRR(B${R.fcfe}:${lastC}${R.fcfe},B${R.date}:${lastC}${R.date}),"n/m")`}; ws.getCell(K,2).numFmt=pctF; ws.getCell(K,2).font=bold;
  ws.getCell(K+1,1).value='Equity (€m)'; ws.getCell(K+1,2).value={formula:'(1-B_GEAR)*B_CPX'}; ws.getCell(K+1,2).numFmt=eurF;
  ws.getCell(K+2,1).value='Max |check diff| (€m)'; ws.getCell(K+2,2).value={formula:`MAX(ABS(MIN(B${R.diff}:${lastC}${R.diff})),ABS(MAX(B${R.diff}:${lastC}${R.diff})))`}; ws.getCell(K+2,2).fill=CHK; ws.getCell(K+2,2).numFmt=numF;
 })();

 // ---------------- SPV SHEET ----------------
 (function(){
  const ws=wb.addWorksheet('SPV'); ws.getColumn(1).width=34; const c0=2;
  ws.getCell(1,1).value='CONSOLIDATED POWER SPV — owns RES (per mode), battery & line; buys residual; sells to DC'; ws.getCell(1,1).font=bold;
  const R={yr:2,dcrev:3,brev:4,wprod:5,sprod:6,resmwh:7,gridc:8,resppa:9,opex:10,ebitda:11,capex:12,dep:13,bal:14,intr:15,rep:16,ebt:17,nol:18,tax:19,fcfe:20,date:21};
  const L={dcrev:'DC revenue (€m)',brev:'Battery revenue (€m)',wprod:'Wind → DC (MWh)',sprod:'Solar → DC (MWh)',resmwh:'Residual from grid (MWh)',gridc:'Residual cost incl fees (€m)',resppa:'RES bought at PPA (€m, mode=0)',opex:'Opex (€m)',ebitda:'EBITDA (€m)',capex:'Capex (€m)',dep:'Depreciation (€m)',bal:'Debt balance (€m)',intr:'Interest (€m)',rep:'Principal (€m)',ebt:'EBT (€m)',nol:'NOL',tax:'Tax (€m)',fcfe:'EQUITY CASH FLOW (€m)',date:'Date'};
  ws.getCell(2,1).value='Year'; for(let i=0;i<NY;i++){const c=ws.getCell(2,c0+i);c.value=Y0+i;c.font=white;c.fill=HDR;}
  Object.keys(L).forEach(k=>{ws.getCell(R[k],1).value=L[k]; if(k==='fcfe'){ws.getCell(R[k],1).font=white;ws.getCell(R[k],1).fill=HDR;}});
  const resLife='MAX(W_LIFE,S_LIFE)';
  for(let i=0;i<NY;i++){
   const C=ws.getColumn(c0+i).letter, cell=(row,f)=>{ws.getCell(row,c0+i).value={formula:f}; ws.getCell(row,c0+i).numFmt=numF;};
   const prevBal=i===0?'0':ws.getColumn(c0+i-1).letter+R.bal, prevNol=i===0?'0':ws.getColumn(c0+i-1).letter+R.nol;
   const OP=`IF(AND(${C}2>=SPV_FF,${C}2<SPV_FF+${resLife}),1,0)`;
   cell(R.dcrev,`${OP}*DC_MW*8760*DC_P*(1+INFL)^(${C}2-2026)/10^6`);
   cell(R.brev,`IF(AND(${C}2>=B_GY,${C}2<B_COD+B_LIFE),(Battery!B$7*(1-B_DEGR)^(${C}2-B_GY)+Battery!B$8)*(1-B_COMP)^(${C}2-B_GY)*(1+INFL)^(${C}2-B_GY),0)+IF(AND(${C}2>B_COD,${C}2<B_COD+B_LIFE),Battery!B$9*(1+INFL)^(${C}2-(B_COD+1)),0)`);
   cell(R.wprod,`Wind!${C}4`); cell(R.sprod,`Solar!${C}4`);
   cell(R.resmwh,`${OP}*MAX(0,DC_MW*8760-${C}${R.wprod}-${C}${R.sprod})`);
   cell(R.gridc,`${OP}*(${C}${R.resmwh}*(RES_P*(1+INFL)^(${C}2-2025)*(1+RES_M)+FEE_E*(1+FEESC)^MAX(0,${C}2-2028))/10^6+DC_MW*FEE_C/1000*(1+FEESC)^MAX(0,${C}2-2028))`);
   cell(R.resppa,`${OP}*(1-RES_OWN)*(W_PPA*${C}${R.wprod}+S_PPA*${C}${R.sprod})/10^6`);
   cell(R.opex,`${OP}*(RES_OWN*(W_OPEX*W_MW+S_OPEX*S_MW)*(1+INFL)^(${C}2-2023)+B_MW*B_DUR*1000*B_CKWH/10^6*B_OPEXP)`);
   cell(R.ebitda,`${C}${R.dcrev}+${C}${R.brev}-${C}${R.gridc}-${C}${R.resppa}-${C}${R.opex}`);
   cell(R.capex,`IF(${C}2=SPV_FF-2,0.3*RES_OWN*(W_MW*W_CAPEX+S_MW*S_CAPEX),IF(${C}2=SPV_FF-1,0.7*RES_OWN*(W_MW*W_CAPEX+S_MW*S_CAPEX)+B_CPX+(W_MW+S_MW)*LINE_C/100,0))`);
   cell(R.dep,`IF(AND(${C}2>=SPV_FF,${C}2<SPV_FF+MIN(20,${resLife})),(RES_OWN*(W_MW*W_CAPEX+S_MW*S_CAPEX)+B_CPX+(W_MW+S_MW)*LINE_C/100)/MIN(20,${resLife}),0)`);
   const blendG=`((GEAR*(RES_OWN*(W_MW*W_CAPEX+S_MW*S_CAPEX)+(W_MW+S_MW)*LINE_C/100)+B_GEAR*B_CPX)/(RES_OWN*(W_MW*W_CAPEX+S_MW*S_CAPEX)+B_CPX+(W_MW+S_MW)*LINE_C/100))`;
   cell(R.bal,`IF(${C}2<SPV_FF,${prevBal}+${C}${R.capex}*${blendG}+IF(${C}${R.capex}>0,(${prevBal}+${C}${R.capex}*${blendG}/2)*RATE,0),MAX(0,${prevBal}-${C}${R.rep}))`);
   cell(R.intr,`IF(${C}2>=SPV_FF,${prevBal}*RATE,0)`);
   cell(R.rep,`IF(AND(${C}2>=SPV_FF,${C}2<SPV_FF+MIN(TENOR,${resLife})),MIN(${prevBal},INDEX($B$${R.bal}:$${ws.getColumn(c0+NY-1).letter}$${R.bal},MATCH(SPV_FF-1,$B$2:$${ws.getColumn(c0+NY-1).letter}$2,0))/MIN(TENOR,${resLife})),0)`);
   cell(R.ebt,`${C}${R.ebitda}-${C}${R.dep}-${C}${R.intr}`);
   cell(R.nol,`IF(${C}2<SPV_FF,0,MIN(0,${prevNol}+${C}${R.ebt}))`);
   cell(R.tax,`IF(${C}2<SPV_FF,0,MAX(0,${C}${R.ebt}+${prevNol})*TAXR)`);
   cell(R.fcfe,`IF(${C}2<SPV_FF,-${C}${R.capex}+${C}${R.capex}*${blendG},${C}${R.ebitda}-${C}${R.tax}-${C}${R.rep}-${C}${R.intr})`);
   ws.getCell(R.date,c0+i).value={formula:`DATE(${C}2,12,31)`}; ws.getCell(R.date,c0+i).numFmt='yyyy-mm-dd';
  }
  const lastC=ws.getColumn(c0+NY-1).letter, K=R.date+2;
  ws.getCell(K,1).value='SPV equity IRR (XIRR)'; ws.getCell(K,2).value={formula:`IFERROR(XIRR(B${R.fcfe}:${lastC}${R.fcfe},B${R.date}:${lastC}${R.date}),"n/m")`}; ws.getCell(K,2).numFmt=pctF; ws.getCell(K,2).font=bold;
  ws.getCell(K+1,1).value='Dashboard SPV IRR at export'; ws.getCell(K+1,2).value=S.spvIRR; ws.getCell(K+1,2).numFmt=pctF; ws.getCell(K+1,2).fill=CHK;
  ws.getCell(K+2,1).value='Note'; ws.getCell(K+2,2).value='Small diffs vs dashboard are timing conventions; use Goal Seek on Inputs!DC_P for target IRR.';
 })();

 // ---------------- OUTPUT ----------------
 const wo=wb.addWorksheet('Output'); wo.getColumn(1).width=36; wo.getColumn(2).width=16;
 wo.getCell(1,1).value='OUTPUT DASHBOARD'; wo.getCell(1,1).font={bold:true,size:13};
 const out=[["Wind equity IRR","=Wind!B24"],["Wind MOIC","=Wind!B25"],["Solar equity IRR","=Solar!B24"],["Solar MOIC","=Solar!B25"],
  ["Battery equity IRR","=Battery!B"+(3+13+12+2-13+13)] ,["SPV equity IRR","=SPV!B23"]];
 let ro=3;
 [["Wind equity IRR","Wind!B24",pctF],["Wind MOIC","Wind!B25",'0.00"x"'],["Solar equity IRR","Solar!B24",pctF],["Solar MOIC","Solar!B25",'0.00"x"'],["SPV equity IRR","SPV!B23",pctF]].forEach(o=>{
  wo.getCell(ro,1).value=o[0]; wo.getCell(ro,2).value={formula:o[1]}; wo.getCell(ro,2).numFmt=o[2]; wo.getCell(ro,2).font=bold; ro++;});
 wo.getCell(ro+1,1).value='All inputs on the Inputs sheet (yellow). Change any input → every sheet recalculates.';
 wo.getCell(ro+2,1).value='Check rows (green) hold the dashboard values at export; diff rows should be ≈ 0.';
 return wb;
}
if(typeof module!=='undefined')module.exports={buildFullModel};
