/* ═══════════════════════════════════════════════════════
   iMagnum FP&A Dashboard — app.js
   Dynamic · SheetJS + Chart.js · GitHub Pages ready
   Data auto-refreshes when Excel files are replaced.
═══════════════════════════════════════════════════════ */
'use strict';

/* ════════════════════════════════════════
   CONFIGURATION LAYER
   Update this mapping if Excel layout changes —
   no other code changes required.
════════════════════════════════════════ */
const CFG = {
  files: {
    mis:    './data/MIS Dashboard working.xlsx',
    budget: './data/Imagnum Budget CY 2026_V10 - Forecast2.xlsx'
  },
  sheets: {
    mis: {
      consolidated: 'Sheet1',
      /* partCol / indiaCol / usCol / consCol are ALL 0-indexed (sheet_to_json array positions) */
      monthMap: {
        'Jan-26': { sheet: 'Consolidated PnL-Jan',  partCol:1, indiaCol:3, usCol:2, consCol:4 },
        'Feb-26': { sheet: 'Consolidated- Mar&Feb', partCol:1, indiaCol:6, usCol:7, consCol:8 },
        'Mar-26': { sheet: 'Consolidated- Mar&Feb', partCol:1, indiaCol:2, usCol:3, consCol:4 },
        'Apr-26': { sheet: 'Consolidted PnL-April', partCol:1, indiaCol:2, usCol:3, consCol:5 },
        'May-26': { sheet: 'Consolidted PnL-May',   partCol:1, indiaCol:2, usCol:3, consCol:5 }
      }
    },
    budget: { plSummary: '3. P&L Summary' }
  },
  /* Row-label search patterns for MIS sheets (first match wins) */
  labels: {
    revenue:      ['Total Revenue', 'Direct Incomes'],
    cogs:         ['Total COGS', 'Total Cogs'],
    grossProfit:  ['GM ','GP ','Gross Profit'],
    manpower:     ['Manpower'],
    facility:     ['Facility Cost'],
    telecom:      ['Telcom/Data', 'Telecom/Data'],
    transport:    ['Transport'],
    otherCogs:    ['Other COGS', 'Other Cogs'],
    sellingExp:   ['Selling Expenses'],
    mgmtSal:      ['Salaries-Management', 'Management Salaries'],
    legalProf:    ['Legal & Professional Fees', 'Legal & Professinal'],
    gaSal:        ['G&A salaries', 'G&A Salaries'],
    adminExp:     ['Gen & Admin'],
    totalSGA:     ['Overall SG&A', 'Overall Sg&A'],
    ebitda:       ['EBITDA', 'Ebidta'],
    financeCost:  ['Finance Cost', 'Interest/Finance', 'Finance Charges'],
    depreciation: ['Depreciation'],
    taxes:        ['Rates & Taxes'],
    pbt:          ['PBT']
  },
  /* Budget sheet label patterns — use LAST match (consolidated section) */
  budgetLabels: {
    revenue:     [' Revenue '],
    salaries:    [' Salaries '],
    facilities:  [' Facilities '],
    telecom:     [' Telcom/Data', ' Telecom'],
    transport:   [' Transport '],
    otherCogs:   [' Other Cogs ', ' Other COGS'],
    totalCogs:   [' Total Cogs', ' Total COGS'],
    gm:          [' GM '],
    totalSGA:    [' Overall SG&A', ' Overall Sg&A'],
    ebitda:      [' Ebidta ', ' EBITDA '],
    financeCost: [' Interest/Finance', ' Finance'],
    pbt:         [' PBT ']
  },
  expenseCategories: [
    { key:'manpower',  label:'Employee Cost',        color:'#2563eb' },
    { key:'facility',  label:'Facility Cost',         color:'#7c3aed' },
    { key:'telecom',   label:'Telecom & Software',    color:'#0891b2' },
    { key:'transport', label:'Transport',             color:'#059669' },
    { key:'sellingExp',label:'Selling & Marketing',   color:'#d97706' },
    { key:'mgmtSal',   label:'Management Salaries',   color:'#dc2626' },
    { key:'legalProf', label:'Legal & Professional',  color:'#4338ca' },
    { key:'adminExp',  label:'Admin & G&A',           color:'#64748b' }
  ]
};

/* ════════════════════════════════════════
   COLOUR PALETTE
════════════════════════════════════════ */
const PAL = {
  blue:'#2563eb', green:'#059669', purple:'#7c3aed', orange:'#d97706',
  red:'#dc2626', teal:'#0891b2', indigo:'#4338ca', amber:'#d97706',
  list:['#2563eb','#059669','#7c3aed','#d97706','#0891b2','#dc2626','#4338ca','#65a30d','#db2777','#06b6d4']
};
const alpha = (hex, a) => {
  const r=parseInt(hex.slice(1,3),16),g=parseInt(hex.slice(3,5),16),b=parseInt(hex.slice(5,7),16);
  return `rgba(${r},${g},${b},${a})`;
};

/* ════════════════════════════════════════
   UTILITIES
════════════════════════════════════════ */
function parseNum(v) {
  if (v===null||v===undefined||v===''||v==='-'||v==='—') return 0;
  let s = String(v).trim().replace(/\$/g,'').replace(/,/g,'').replace(/\s/g,'');
  if (s.startsWith('(') && s.endsWith(')')) return -(parseFloat(s.slice(1,-1))||0);
  const n=parseFloat(s); return isNaN(n)?0:n;
}
function fmtUSD(n, short=false) {
  if (n===null||n===undefined) return '—';
  const abs=Math.abs(n);
  if (short) {
    if (abs>=1e6) return `$${(n/1e6).toFixed(1)}M`;
    if (abs>=1e3) return `$${(n/1e3).toFixed(1)}K`;
    return `$${n.toFixed(0)}`;
  }
  const s=Math.abs(Math.round(n)).toLocaleString('en-US');
  return n<0?`($${s})`:`$${s}`;
}
function fmtPct(n) {
  if (n===null||n===undefined||isNaN(n)) return '—';
  return `${n>=0?'+':''}${n.toFixed(1)}%`;
}
function mom(curr,prev) { if(!prev||prev===0)return null; return((curr-prev)/Math.abs(prev))*100; }
function pct(part,whole) { if(!whole||whole===0)return 0; return(part/whole)*100; }

function labelMatch(cell, patterns) {
  const c=String(cell||'').trim().toLowerCase();
  return patterns.some(p=>c.includes(p.trim().toLowerCase()));
}
/* Find FIRST row matching patterns in the given column */
function findRowIdx(rows, patterns, colIdx=0) {
  for (let i=0;i<rows.length;i++) {
    if (labelMatch(rows[i][colIdx], patterns)) return i;
  }
  return -1;
}
/* Find LAST row matching patterns — used for budget consolidated section */
function findLastRowIdx(rows, patterns, colIdx=0) {
  let last=-1;
  for (let i=0;i<rows.length;i++) {
    if (labelMatch(rows[i][colIdx], patterns)) last=i;
  }
  return last;
}
function getVal(rows, patterns, col, partCol=0) {
  const idx=findRowIdx(rows, patterns, partCol);
  return idx===-1?0:parseNum(rows[idx][col]);
}
function getValLast(rows, patterns, col, partCol=0) {
  const idx=findLastRowIdx(rows, patterns, partCol);
  return idx===-1?0:parseNum(rows[idx][col]);
}
function setStatus(msg) { const e=document.getElementById('loading-status'); if(e)e.textContent=msg; }

/* ════════════════════════════════════════
   CHART.JS GLOBAL DEFAULTS
════════════════════════════════════════ */
function applyChartDefaults() {
  /* Intentionally empty — Chart.js 4.4.x has a Proxy recursion bug
     triggered by ANY modification to Chart.defaults when axes are present.
     All styling is passed inline per-chart via lineOpts(). */
}
const CHARTS={};
function makeChart(id,cfg) {
  if(CHARTS[id]){CHARTS[id].destroy();}
  const ctx=document.getElementById(id);
  if(!ctx)return null;
  CHARTS[id]=new Chart(ctx,cfg);
  return CHARTS[id];
}

/* ════════════════════════════════════════
   DATA LOADER
════════════════════════════════════════ */
let WB={mis:null,budget:null};
let DATA={};

async function loadWorkbook(url,label) {
  setStatus(`Loading ${label}…`);
  const resp=await fetch(url+'?t='+Date.now());
  if(!resp.ok)throw new Error(`Cannot fetch ${url}: ${resp.status}`);
  const buf=await resp.arrayBuffer();
  return XLSX.read(new Uint8Array(buf),{type:'array',cellText:true,cellNF:true});
}
function sheetToArr(wb,name) {
  const ws=wb.Sheets[name];
  if(!ws)return[];
  return XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:false});
}

/* ════════════════════════════════════════
   EXTRACT — MIS Sheet1 (Consolidated Multi-Month)
════════════════════════════════════════ */
function extractMIS() {
  const rows=sheetToArr(WB.mis,CFG.sheets.mis.consolidated);
  /* Header row — Particulars at col 0 */
  const hdrIdx=findRowIdx(rows,['Particulars','particulars'],0);
  if(hdrIdx===-1){console.warn('Sheet1 header not found');return{months:[]};}
  const hdr=rows[hdrIdx];
  const months=[];
  for(let c=1;c<hdr.length;c++){
    const v=String(hdr[c]||'').trim();
    if(/\w+-\d{2}/.test(v)) months.push({label:v,col:c});
  }
  function ser(patterns){return months.map(m=>getVal(rows,patterns,m.col,0));}
  return {
    months:months.map(m=>m.label),
    revenue:     ser(CFG.labels.revenue),
    cogs:        ser(CFG.labels.cogs),
    grossProfit: ser(CFG.labels.grossProfit),
    ebitda:      ser(CFG.labels.ebitda),
    financeCost: ser(CFG.labels.financeCost),
    depreciation:ser(CFG.labels.depreciation),
    taxes:       ser(CFG.labels.taxes),
    pbt:         ser(CFG.labels.pbt),
    sellingExp:  ser(CFG.labels.sellingExp),
    mgmtSal:     ser(CFG.labels.mgmtSal),
    legalProf:   ser(CFG.labels.legalProf),
    gaSal:       ser(CFG.labels.gaSal),
    adminExp:    ser(CFG.labels.adminExp),
    manpower:    ser(CFG.labels.manpower),
    facility:    ser(CFG.labels.facility),
    telecom:     ser(CFG.labels.telecom),
    transport:   ser(CFG.labels.transport),
    otherCogs:   ser(CFG.labels.otherCogs),
    totalSGA:    ser(CFG.labels.totalSGA)
  };
}

/* ════════════════════════════════════════
   EXTRACT — Individual Month Sheets (India/US/Consol)
   Uses CFG column indices DIRECTLY — no dynamic override
   (avoids bug on combined Feb/Mar sheet where both months
    share one sheet with different column positions)
════════════════════════════════════════ */
function extractMonthPL(monthLabel) {
  const mc=CFG.sheets.mis.monthMap[monthLabel];
  if(!mc)return null;
  const rows=sheetToArr(WB.mis,mc.sheet);
  if(!rows.length)return null;
  const {partCol,indiaCol,usCol,consCol}=mc;

  function gv(patterns){
    return {
      india:getVal(rows,patterns,indiaCol,partCol),
      us:   getVal(rows,patterns,usCol,   partCol),
      cons: getVal(rows,patterns,consCol, partCol)
    };
  }

  /* Collect all detail rows for drill-down */
  const result={month:monthLabel,rows:[]};
  for(let i=0;i<rows.length;i++){
    const label=String(rows[i][partCol]||'').trim();
    if(!label)continue;
    const india=parseNum(rows[i][indiaCol]);
    const us   =parseNum(rows[i][usCol]);
    const cons =parseNum(rows[i][consCol]);
    const allZero=india===0&&us===0&&cons===0;
    const isTotalLine=/total|cogs|gm$|gp$|ebitda|pbt|sga|sg&a/i.test(label);
    result.rows.push({
      label,india,us,cons,
      type:allZero&&!isTotalLine?'header':'data'
    });
  }
  result.summary={
    revenue:    gv(CFG.labels.revenue),
    cogs:       gv(CFG.labels.cogs),
    grossProfit:gv(CFG.labels.grossProfit),
    manpower:   gv(CFG.labels.manpower),
    facility:   gv(CFG.labels.facility),
    telecom:    gv(CFG.labels.telecom),
    transport:  gv(CFG.labels.transport),
    otherCogs:  gv(CFG.labels.otherCogs),
    sellingExp: gv(CFG.labels.sellingExp),
    mgmtSal:    gv(CFG.labels.mgmtSal),
    legalProf:  gv(CFG.labels.legalProf),
    gaSal:      gv(CFG.labels.gaSal),
    adminExp:   gv(CFG.labels.adminExp),
    totalSGA:   gv(CFG.labels.totalSGA),
    ebitda:     gv(CFG.labels.ebitda),
    financeCost:gv(CFG.labels.financeCost),
    pbt:        gv(CFG.labels.pbt)
  };
  return result;
}

/* ════════════════════════════════════════
   EXTRACT — Budget Workbook
   Uses LAST match for ambiguous labels so we get the
   Consolidated section (rows 38-46) not Current (row 6)
════════════════════════════════════════ */
function extractBudget() {
  const rows=sheetToArr(WB.budget,CFG.sheets.budget.plSummary);
  /* Find header row — col 0 says "Particulars" */
  const hdrIdx=findRowIdx(rows,['Particulars','particulars'],0);
  const hdr=hdrIdx>=0?rows[hdrIdx]:rows[3]||[];
  const months=[];
  for(let c=0;c<hdr.length;c++){
    const v=String(hdr[c]||'').trim();
    if(/\w+-\d{2}/.test(v)) months.push({label:v,col:c});
  }
  /* LAST match for consolidated section */
  function serLast(patterns){return months.map(m=>getValLast(rows,patterns,m.col,0));}
  /* First match for unique labels */
  function ser(patterns){return months.map(m=>getVal(rows,patterns,m.col,0));}
  return {
    months:      months.map(m=>m.label),
    revenue:     serLast(CFG.budgetLabels.revenue),
    salaries:    serLast(CFG.budgetLabels.salaries),
    facilities:  serLast(CFG.budgetLabels.facilities),
    telecom:     serLast(CFG.budgetLabels.telecom),
    transport:   serLast(CFG.budgetLabels.transport),
    otherCogs:   serLast(CFG.budgetLabels.otherCogs),
    totalCogs:   serLast(CFG.budgetLabels.totalCogs),
    gm:          serLast(CFG.budgetLabels.gm),
    totalSGA:    ser(CFG.budgetLabels.totalSGA),  /* unique */
    ebitda:      ser(CFG.budgetLabels.ebitda),    /* unique */
    financeCost: ser(CFG.budgetLabels.financeCost),
    pbt:         ser(CFG.budgetLabels.pbt)        /* unique */
  };
}

/* ════════════════════════════════════════
   PAGE 1 — EXECUTIVE DASHBOARD
════════════════════════════════════════ */
function renderExecutive() {
  const mis=DATA.mis; const months=mis.months; const n=months.length;
  if(!n)return;
  const li=n-1, pi=n>1?n-2:0;

  /* Animated KPI cards */
  function kpi(vId,bId,spId,series,color=PAL.blue){
    const curr=series[li]||0, prev=series[pi]||0;
    const chg=n>1?mom(curr,prev):null;
    animateCounter(document.getElementById(vId),curr);
    const b=document.getElementById(bId);
    if(b){
      b.textContent=chg===null?'—':fmtPct(chg);
      b.className='kpi-badge'+(chg===null?' neutral':chg>=0?'':' negative');
    }
    makeChart(spId,{type:'line',data:{labels:months,datasets:[{data:series,borderColor:color,borderWidth:2,fill:true,backgroundColor:alpha(color,0.1),tension:0.4,pointRadius:0}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{enabled:false}},scales:{x:{display:false},y:{display:false}},animation:{duration:900}}});
  }
  kpi('v-revenue','b-revenue','spark-revenue',mis.revenue,PAL.blue);
  kpi('v-gp',     'b-gp',     'spark-gp',     mis.grossProfit,PAL.green);
  kpi('v-ebitda', 'b-ebitda', 'spark-ebitda', mis.ebitda,PAL.purple);
  kpi('v-pbt',    'b-pbt',    'spark-pbt',    mis.pbt,PAL.orange);

  /* Margin strips */
  const rev=mis.revenue[li]||1, gp=mis.grossProfit[li]||0, eb=mis.ebitda[li]||0, pb=mis.pbt[li]||0;
  setBar('ms-gm',    pct(gp,rev),'ms-gm-val');
  setBar('ms-ebitda',pct(eb,rev),'ms-ebitda-val');
  setBar('ms-net',   pct(pb,rev),'ms-net-val');

  /* Revenue + GP Trend */
  makeChart('chart-rev-trend',{type:'bar',data:{labels:months,datasets:[
    {label:'Revenue',data:mis.revenue,backgroundColor:alpha(PAL.blue,0.2),borderColor:PAL.blue,borderWidth:2,borderRadius:4,type:'bar'},
    {label:'Gross Profit',data:mis.grossProfit,borderColor:PAL.green,borderWidth:2.5,fill:false,tension:0.4,type:'line',pointRadius:4,pointBackgroundColor:PAL.green}
  ]},options:lineOpts({scales:{y:{ticks:{callback:v=>fmtUSD(v,true)}}}})});

  /* Expense Distribution donut */
  const expL=['Manpower','Facility','Telecom','Transport','Other COGS','Selling','Mgmt Sal','Legal','Admin'];
  const expV=[mis.manpower[li],mis.facility[li],mis.telecom[li],mis.transport[li],mis.otherCogs[li],mis.sellingExp[li],mis.mgmtSal[li],mis.legalProf[li],mis.adminExp[li]].map(v=>Math.abs(v||0));
  makeChart('chart-exp-dist',{type:'doughnut',data:{labels:expL,datasets:[{data:expV,backgroundColor:PAL.list,borderWidth:0,hoverOffset:6}]},options:{responsive:true,maintainAspectRatio:false,cutout:'60%',plugins:{legend:{position:'right',labels:{font:{size:11},padding:8}},tooltip:{callbacks:{label:c=>` ${c.label}: ${fmtUSD(c.raw)}`}}}}});

  /* Waterfall */
  renderWaterfall('chart-waterfall',months[li],mis,li);

  /* EBITDA Performance */
  const ebPcts=mis.revenue.map((r,i)=>r?pct(mis.ebitda[i],r):0);
  makeChart('chart-ebitda-perf',{type:'bar',data:{labels:months,datasets:[
    {label:'EBITDA ($)',data:mis.ebitda,backgroundColor:mis.ebitda.map(v=>alpha(v>=0?PAL.purple:PAL.red,0.7)),borderColor:mis.ebitda.map(v=>v>=0?PAL.purple:PAL.red),borderWidth:1,borderRadius:4,yAxisID:'y'},
    {label:'EBITDA %',data:ebPcts,type:'line',borderColor:PAL.orange,borderWidth:2,fill:false,tension:0.4,pointRadius:4,yAxisID:'y1'}
  ]},options:lineOpts({scales:{y:{ticks:{callback:v=>fmtUSD(v,true)},position:'left'},y1:{ticks:{callback:v=>`${v.toFixed(0)}%`},position:'right',grid:{display:false}}}})});

  /* Margin Trend */
  const gmP=mis.revenue.map((r,i)=>r?pct(mis.grossProfit[i],r):0);
  const ebP=mis.revenue.map((r,i)=>r?pct(mis.ebitda[i],r):0);
  const ntP=mis.revenue.map((r,i)=>r?pct(mis.pbt[i],r):0);
  makeChart('chart-margin-trend',{type:'line',data:{labels:months,datasets:[
    {label:'Gross Margin %',data:gmP,borderColor:PAL.blue,backgroundColor:alpha(PAL.blue,0.05),fill:true,tension:0.4,borderWidth:2.5,pointRadius:4},
    {label:'EBITDA %',data:ebP,borderColor:PAL.purple,borderDash:[5,3],fill:false,tension:0.4,borderWidth:2,pointRadius:3},
    {label:'Net %',data:ntP,borderColor:PAL.orange,borderDash:[2,3],fill:false,tension:0.4,borderWidth:2,pointRadius:3}
  ]},options:lineOpts({scales:{y:{ticks:{callback:v=>`${v.toFixed(0)}%`}}}})});

  const dr=document.getElementById('exec-date-range');
  if(dr)dr.textContent=`${months[0]} – ${months[n-1]}`;
}

function renderWaterfall(cId,label,mis,idx){
  const rev=mis.revenue[idx]||0, cogs=Math.abs(mis.cogs[idx]||0),gp=mis.grossProfit[idx]||0;
  const sga=Math.abs(mis.totalSGA[idx]||0), eb=mis.ebitda[idx]||0;
  const fc=Math.abs(mis.financeCost[idx]||0), pbt=mis.pbt[idx]||0;
  const el=document.getElementById('waterfall-label');
  if(el)el.textContent=label;
  const steps=[
    {l:'Revenue',    v:rev,  b:0,      c:PAL.blue,   total:false},
    {l:'- COGS',     v:-cogs,b:rev-cogs,c:PAL.red,    total:false},
    {l:'Gross Profit',v:gp, b:0,       c:PAL.green,  total:true},
    {l:'- SG&A',     v:-sga, b:gp,     c:PAL.orange, total:false},
    {l:'EBITDA',     v:eb,   b:0,       c:PAL.purple, total:true},
    {l:'- Fin Cost', v:-fc,  b:Math.max(0,eb), c:PAL.red, total:false},
    {l:'Net PBT',    v:pbt,  b:0,       c:pbt>=0?PAL.teal:PAL.red, total:true}
  ];
  const bases =steps.map(s=>s.total?0:Math.min(s.b,s.b+s.v));
  const values=steps.map(s=>Math.abs(s.v));
  makeChart(cId,{type:'bar',data:{labels:steps.map(s=>s.l),datasets:[
    {data:bases, backgroundColor:'transparent',borderWidth:0},
    {data:values,backgroundColor:steps.map(s=>alpha(s.c,0.8)),borderColor:steps.map(s=>s.c),borderWidth:1.5,borderRadius:3}
  ]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:(c)=>{if(c.datasetIndex===0)return null;return ` ${steps[c.dataIndex].l}: ${fmtUSD(steps[c.dataIndex].v)}`;}}}},scales:{x:{stacked:true},y:{stacked:true,ticks:{callback:v=>fmtUSD(v,true)}}}}});
}

/* ════════════════════════════════════════
   PAGE 2 — P&L DASHBOARD
════════════════════════════════════════ */
function renderPL(monthLabel) {
  const plData=extractMonthPL(monthLabel);
  const s=plData?.summary;
  const tbody=document.getElementById('pl-tbody');

  const el=document.getElementById('pl-table-period');
  if(el)el.textContent=monthLabel||'—';

  if(!plData||!s){
    tbody.innerHTML='<tr><td colspan="4" class="table-placeholder">Data not available for this month</td></tr>';
    return;
  }

  /* P&L structure rows */
  const n=(i,u,c)=>({india:i||0,us:u||0,cons:c||0});
  const structure=[
    {sec:'REVENUE'},
    {label:'Service Revenue',  ...n(s.revenue.india,    s.revenue.us,    s.revenue.cons),    cls:'row-subtotal'},
    {sec:'COST OF GOODS SOLD'},
    {label:'Employee / Manpower',...n(s.manpower.india,  s.manpower.us,   s.manpower.cons),   indent:1},
    {label:'Facility Cost',    ...n(s.facility.india,   s.facility.us,   s.facility.cons),   indent:1},
    {label:'Telecom & Data',   ...n(s.telecom.india,    s.telecom.us,    s.telecom.cons),    indent:1},
    {label:'Transport',        ...n(s.transport.india,  s.transport.us,  s.transport.cons),  indent:1},
    {label:'Other COGS',       ...n(s.otherCogs.india,  s.otherCogs.us,  s.otherCogs.cons),  indent:1},
    {label:'Total COGS',       ...n(s.cogs.india,       s.cogs.us,       s.cogs.cons),       cls:'row-subtotal'},
    {label:'GROSS PROFIT',     ...n(s.grossProfit.india,s.grossProfit.us,s.grossProfit.cons),cls:'row-total'},
    {sec:'SG&A EXPENSES'},
    {label:'Selling & Marketing',...n(s.sellingExp.india,s.sellingExp.us, s.sellingExp.cons), indent:1},
    {label:'Management Salaries',...n(s.mgmtSal.india,  s.mgmtSal.us,   s.mgmtSal.cons),    indent:1},
    {label:'Legal & Professional',...n(s.legalProf.india,s.legalProf.us, s.legalProf.cons),  indent:1},
    {label:'G&A Salaries',     ...n(s.gaSal.india,      s.gaSal.us,      s.gaSal.cons),      indent:1},
    {label:'General & Admin',  ...n(s.adminExp.india,   s.adminExp.us,   s.adminExp.cons),   indent:1},
    {label:'Total SG&A',       ...n(s.totalSGA.india,   s.totalSGA.us,   s.totalSGA.cons),   cls:'row-subtotal'},
    {label:'EBITDA',           ...n(s.ebitda.india,     s.ebitda.us,     s.ebitda.cons),     cls:'row-grand'},
    {sec:'BELOW THE LINE'},
    {label:'Finance Cost',     ...n(s.financeCost.india,s.financeCost.us,s.financeCost.cons),indent:1},
    {label:'NET PROFIT (PBT)', ...n(s.pbt.india,        s.pbt.us,        s.pbt.cons),        cls:'row-grand'}
  ];

  tbody.innerHTML=structure.map(row=>{
    if(row.sec) return `<tr class="row-header"><td colspan="4">${row.sec}</td></tr>`;
    const nc=v=>v<0?'col-num num-val negative':v>0?'col-num num-val':'col-num num-val text-muted';
    return `<tr class="${row.cls||''}">
      <td class="${row.indent?'pl-indent-1':''}">${row.label}</td>
      <td class="${nc(row.india)}">${fmtUSD(row.india)}</td>
      <td class="${nc(row.us)}">${fmtUSD(row.us)}</td>
      <td class="${nc(row.cons)} highlight-col">${fmtUSD(row.cons)}</td>
    </tr>`;
  }).join('');

  /* Summary chips */
  const chips=document.getElementById('pl-chips');
  if(chips){
    chips.innerHTML=[
      {l:'Revenue',   v:s.revenue.cons,    p:null},
      {l:'Gross Profit',v:s.grossProfit.cons,p:pct(s.grossProfit.cons,s.revenue.cons||1)},
      {l:'EBITDA',    v:s.ebitda.cons,     p:pct(s.ebitda.cons,s.revenue.cons||1)},
      {l:'Net PBT',   v:s.pbt.cons,        p:pct(s.pbt.cons,s.revenue.cons||1)}
    ].map(c=>{
      const pc=c.p!==null?(c.p>=0?'chip-positive':'chip-negative'):'';
      return `<div class="pl-chip"><span class="pl-chip-label">${c.l}</span><span class="pl-chip-value">${fmtUSD(c.v)}</span>${c.p!==null?`<span class="pl-chip-pct ${pc}">${fmtPct(c.p)}</span>`:''}</div>`;
    }).join('');
  }

  /* Charts */
  const cats=['Revenue','COGS','Gross Profit','SG&A','EBITDA','Net PBT'];
  const indV=[s.revenue.india,Math.abs(s.cogs.india||0),s.grossProfit.india,Math.abs(s.totalSGA.india||0),s.ebitda.india,s.pbt.india];
  const usV =[s.revenue.us,   Math.abs(s.cogs.us||0),   s.grossProfit.us,   Math.abs(s.totalSGA.us||0),   s.ebitda.us,   s.pbt.us];
  const coV =[s.revenue.cons, Math.abs(s.cogs.cons||0), s.grossProfit.cons, Math.abs(s.totalSGA.cons||0), s.ebitda.cons, s.pbt.cons];

  makeChart('chart-pl-bar',{type:'bar',data:{labels:cats,datasets:[
    {label:'India',       data:indV,backgroundColor:alpha(PAL.blue,0.7),  borderRadius:3},
    {label:'US',          data:usV, backgroundColor:alpha(PAL.purple,0.7),borderRadius:3},
    {label:'Consolidated',data:coV, backgroundColor:alpha(PAL.teal,0.7),  borderRadius:3}
  ]},options:lineOpts({scales:{y:{ticks:{callback:v=>fmtUSD(v,true)}}}})});

  makeChart('chart-pl-pie',{type:'doughnut',data:{labels:['Manpower','Facility','Telecom','Transport','Other COGS','Selling','Mgmt Sal','Legal','Admin'],datasets:[{data:[Math.abs(s.manpower.cons||0),Math.abs(s.facility.cons||0),Math.abs(s.telecom.cons||0),Math.abs(s.transport.cons||0),Math.abs(s.otherCogs.cons||0),Math.abs(s.sellingExp.cons||0),Math.abs(s.mgmtSal.cons||0),Math.abs(s.legalProf.cons||0),Math.abs(s.adminExp.cons||0)],backgroundColor:PAL.list,borderWidth:0,hoverOffset:5}]},options:{responsive:true,maintainAspectRatio:false,cutout:'55%',plugins:{legend:{position:'right',labels:{font:{size:11},padding:8}},tooltip:{callbacks:{label:c=>` ${c.label}: ${fmtUSD(c.raw)}`}}}}});
}

/* ════════════════════════════════════════
   PAGE 3 — EXPENSE ANALYSIS
════════════════════════════════════════ */
function renderExpenseAnalysis(monthLabel, entity) {
  const mis=DATA.mis, months=mis.months;
  const mIdx=months.indexOf(monthLabel);
  const si=mIdx>=0?mIdx:months.length-1;
  const totalExp=CFG.expenseCategories.reduce((a,c)=>a+Math.abs(mis[c.key]?.[si]||0),0)||1;

  /* Expense Mix donut */
  const mixLabels=CFG.expenseCategories.map(c=>c.label);
  const mixVals  =CFG.expenseCategories.map(c=>Math.abs(mis[c.key]?.[si]||0));
  makeChart('chart-exp-mix',{type:'doughnut',data:{labels:mixLabels,datasets:[{data:mixVals,backgroundColor:CFG.expenseCategories.map(c=>c.color),borderWidth:0,hoverOffset:6}]},options:{responsive:true,maintainAspectRatio:false,cutout:'58%',plugins:{legend:{position:'right',labels:{font:{size:11},padding:8}},tooltip:{callbacks:{label:c=>` ${c.label}: ${fmtUSD(c.raw)} (${pct(c.raw,totalExp).toFixed(1)}%)`}}}}});

  /* Stacked trend */
  makeChart('chart-exp-trend',{type:'line',data:{labels:months,datasets:CFG.expenseCategories.map((c,i)=>({label:c.label,data:(mis[c.key]||[]).map(v=>Math.abs(v||0)),borderColor:c.color,backgroundColor:alpha(c.color,0.05+(i===0?0.1:0)),fill:i===0?'origin':false,tension:0.4,borderWidth:2,pointRadius:2}))},options:lineOpts({scales:{y:{ticks:{callback:v=>fmtUSD(v,true)}}}})});

  /* Category drill-down cards */
  const container=document.getElementById('expense-categories');
  if(!container)return;
  const plData=DATA.monthPL&&DATA.monthPL[months[si]];

  container.innerHTML=CFG.expenseCategories.map((cat,ci)=>{
    const val=Math.abs(mis[cat.key]?.[si]||0);
    const share=pct(val,totalExp);
    const subs=getSubItems(plData,cat.key);
    const subHTML=subs.length
      ?`<table class="exp-sub-table"><thead><tr><th>Line Item</th><th class="num">Amount</th><th>% Share</th><th class="pct-bar-cell"></th></tr></thead><tbody>${subs.map(s=>`<tr><td>${s.label}</td><td class="num">${fmtUSD(s.val)}</td><td class="num">${pct(s.val,val||1).toFixed(1)}%</td><td><div class="mini-pct-bar"><div class="mini-pct-fill" style="width:${Math.min(100,pct(s.val,val||1))}%;background:${cat.color}"></div></div></td></tr>`).join('')}</tbody></table>`
      :`<p style="font-size:12px;color:var(--text-muted);padding:8px 0">Sub-items available in individual month view</p>`;
    return `<div class="exp-cat-card">
      <div class="exp-cat-header" onclick="toggleExpCat(this)">
        <div class="exp-cat-dot" style="background:${cat.color}"></div>
        <span class="exp-cat-name">${cat.label}</span>
        <span class="exp-cat-amount" style="color:${cat.color}">${fmtUSD(val)}</span>
        <span class="exp-cat-pct">${share.toFixed(1)}%</span>
        <span class="exp-cat-chevron">▼</span>
      </div>
      <div class="exp-cat-body">
        <div class="exp-cat-charts">
          <div class="exp-cat-chart-wrap" style="height:170px"><canvas id="exp-pie-${ci}"></canvas></div>
          <div class="exp-cat-chart-wrap" style="height:170px"><canvas id="exp-bar-${ci}"></canvas></div>
        </div>
        ${subHTML}
      </div>
    </div>`;
  }).join('');

  /* Render mini charts */
  CFG.expenseCategories.forEach((cat,ci)=>{
    const subs=getSubItems(plData,cat.key);
    const val=Math.abs(mis[cat.key]?.[si]||0);
    if(subs.length){
      makeChart(`exp-pie-${ci}`,{type:'doughnut',data:{labels:subs.map(s=>s.label),datasets:[{data:subs.map(s=>Math.abs(s.val||0)),backgroundColor:PAL.list,borderWidth:0,hoverOffset:4}]},options:{responsive:true,maintainAspectRatio:false,cutout:'55%',plugins:{legend:{display:false}}}});
    } else {
      makeChart(`exp-pie-${ci}`,{type:'doughnut',data:{labels:[cat.label,'Other'],datasets:[{data:[val,Math.max(0,totalExp-val)],backgroundColor:[cat.color,alpha(cat.color,0.15)],borderWidth:0}]},options:{responsive:true,maintainAspectRatio:false,cutout:'60%',plugins:{legend:{display:false}}}});
    }
    makeChart(`exp-bar-${ci}`,{type:'bar',data:{labels:months,datasets:[{label:cat.label,data:(mis[cat.key]||[]).map(v=>Math.abs(v||0)),backgroundColor:alpha(cat.color,0.75),borderRadius:3}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{ticks:{font:{size:9}}},y:{ticks:{callback:v=>fmtUSD(v,true),font:{size:9}}}}}});
  });
}

window.toggleExpCat=function(h){
  h.classList.toggle('open');
  h.nextElementSibling.classList.toggle('open');
};

function getSubItems(plData,catKey){
  if(!plData)return[];
  const pats={
    manpower:  ['Operation salaries','Ops Support Salaries','Consultancy fee','Contract employee','Coding charges','Employee ESI','EPF Admin','Employee PF','Employer ESI','Employer PF','Professional tax','Gratuity'],
    facility:  ['Electricity Expense','Janani Property','Laptop Rental','Thiruttani','Max Office','Rent Expense','Trichy rental'],
    telecom:   ['Telephone Expense','Dues and Subscriptions','IT and Internet','Telecom Charges','Telephone Rental','License/Subscription','Software Expenses','Software expense','Software development'],
    transport: ['Staff Transportation'],
    sellingExp:['Travel Expense','Meals and Entertainment','Sales and marketing','S&M Salaries','Advertisement','Conference','Business Promotion'],
    mgmtSal:   ['Management Salaries','Bonus and Incentives'],
    legalProf: ['FPO Charges','Legal & Professional Services','Audit'],
    adminExp:  ['Consultancy Fees','Fuel/Mileage','Payroll Processing','Travelling Expenses','Insurance Expense','Postage','Office Supplies','Housekeeping','Blocked Credit','Payroll taxes','HR Salaries','Admin Salaries','IT Salaries']
  };
  const ps=pats[catKey]||[];
  const res=[];
  for(const row of plData.rows){
    if(row.type==='header')continue;
    if(ps.some(p=>row.label.toLowerCase().includes(p.toLowerCase()))){
      const v=Math.abs(row.cons||row.india||row.us||0);
      if(v>0) res.push({label:row.label,val:v});
    }
  }
  return res.slice(0,10);
}

/* ════════════════════════════════════════
   PAGE 4 — BUDGET vs ACTUAL
════════════════════════════════════════ */
function renderBudgetVsActual(monthLabel) {
  const mis=DATA.mis, bud=DATA.budget;
  if(!bud||!bud.months.length){
    document.getElementById('bva-tbody').innerHTML='<tr><td colspan="7" class="table-placeholder">Budget workbook not loaded</td></tr>';
    return;
  }
  const aI=mis.months.indexOf(monthLabel);
  const bI=bud.months.findIndex(m=>m.trim()===monthLabel.trim());

  document.getElementById('bva-table-period').textContent=monthLabel||'—';

  if(aI===-1){
    document.getElementById('bva-tbody').innerHTML=`<tr><td colspan="7" class="table-placeholder">No actual data for ${monthLabel}</td></tr>`;
    return;
  }

  /* BvA structure — map budget labels to actual MIS keys */
  const rows=[
    {sec:'REVENUE'},
    {key:'revenue',    label:'Total Revenue',     bud:bud.revenue[bI]||0,   act:mis.revenue[aI]||0,   fav:v=>v>=0, cls:'row-subtotal'},
    {sec:'COST OF GOODS SOLD'},
    {key:'salaries',   label:'Salaries',           bud:bud.salaries[bI]||0,  act:mis.manpower[aI]||0,  fav:v=>v<=0, indent:true},
    {key:'facility',   label:'Facilities',         bud:bud.facilities[bI]||0,act:mis.facility[aI]||0,  fav:v=>v<=0, indent:true},
    {key:'telecom',    label:'Telecom & Data',      bud:bud.telecom[bI]||0,  act:mis.telecom[aI]||0,  fav:v=>v<=0, indent:true},
    {key:'transport',  label:'Transport',           bud:bud.transport[bI]||0, act:mis.transport[aI]||0, fav:v=>v<=0, indent:true},
    {key:'otherCogs',  label:'Other COGS',         bud:bud.otherCogs[bI]||0, act:mis.otherCogs[aI]||0, fav:v=>v<=0, indent:true},
    {key:'totalCogs',  label:'Total COGS',         bud:bud.totalCogs[bI]||0, act:mis.cogs[aI]||0,      fav:v=>v<=0, cls:'row-subtotal'},
    {sec:'PROFITABILITY'},
    {key:'gm',         label:'Gross Profit',        bud:bud.gm[bI]||0,        act:mis.grossProfit[aI]||0,fav:v=>v>=0, cls:'row-total'},
    {key:'totalSGA',   label:'Total SG&A',          bud:bud.totalSGA[bI]||0, act:mis.totalSGA[aI]||0,  fav:v=>v<=0, indent:true},
    {key:'ebitda',     label:'EBITDA',              bud:bud.ebitda[bI]||0,   act:mis.ebitda[aI]||0,   fav:v=>v>=0, cls:'row-grand'},
    {key:'financeCost',label:'Finance Charges',     bud:bud.financeCost[bI]||0,act:mis.financeCost[aI]||0,fav:v=>v>=0, indent:true},
    {key:'pbt',        label:'Net Profit (PBT)',    bud:bud.pbt[bI]||0,      act:mis.pbt[aI]||0,      fav:v=>v>=0, cls:'row-grand'}
  ];

  /* KPI summary cards */
  const kpiRow=document.getElementById('bva-kpi-row');
  if(kpiRow){
    const kpis=[
      {l:'Budget Revenue',   v:fmtUSD(bud.revenue[bI]||0)},
      {l:'Actual Revenue',   v:fmtUSD(mis.revenue[aI]||0)},
      {l:'Revenue Variance', v:(()=>{const d=(mis.revenue[aI]||0)-(bud.revenue[bI]||0);return{val:d,txt:fmtUSD(d)};})()},
      {l:'EBITDA Variance',  v:(()=>{const d=(mis.ebitda[aI]||0)-(bud.ebitda[bI]||0);return{val:d,txt:fmtUSD(d)};})()}
    ];
    kpiRow.innerHTML=kpis.map((k,i)=>{
      const isVar=i>=2;
      const cls=isVar?(k.v.val>=0?'favorable-text':'unfavorable-text'):'';
      return `<div class="bva-kpi-card"><div class="bva-kpi-label">${k.l}</div><div class="bva-kpi-actual ${cls}">${isVar?k.v.txt:k.v}</div></div>`;
    }).join('');
  }

  /* Variance table */
  const tbody=document.getElementById('bva-tbody');
  tbody.innerHTML=rows.map(row=>{
    if(row.sec) return `<tr class="row-header"><td colspan="7">${row.sec}</td></tr>`;
    if(bI===-1) return `<tr class="${row.cls||''}"><td class="${row.indent?'pl-indent-1':''}">${row.label}</td><td colspan="5" style="text-align:center;color:var(--text-muted);font-size:12px">No budget data</td><td></td></tr>`;
    const varV=row.act-row.bud;
    const varPct=row.bud?((varV/Math.abs(row.bud))*100):0;
    const fav=row.fav(varV);
    const within=Math.abs(varPct)<=5;
    const vc=within?'neutral-var':fav?'favorable':'unfavorable';
    const tl=within?'🟡':fav?'🟢':'🔴';
    return `<tr class="${row.cls||''}">
      <td class="${row.indent?'pl-indent-1':''}">${row.label}</td>
      <td class="col-num">${fmtUSD(row.bud)}</td>
      <td class="col-num">${fmtUSD(row.act)}</td>
      <td class="col-num ${vc}">${fmtUSD(varV)}</td>
      <td class="col-num ${vc}">${fmtPct(varPct)}</td>
      <td class="col-status">${tl}</td>
      <td><button class="drill-btn" onclick="openDrilldown('${row.key}','${monthLabel}')" title="Drill down">⊕</button></td>
    </tr>`;
  }).join('');

  /* BvA Grouped Bar */
  const mRows=rows.filter(r=>!r.sec&&['revenue','totalCogs','gm','ebitda','pbt'].includes(r.key));
  makeChart('chart-bva-grouped',{type:'bar',data:{labels:mRows.map(r=>r.label),datasets:[
    {label:'Budget',data:mRows.map(r=>Math.abs(r.bud||0)),backgroundColor:alpha(PAL.blue,0.6),borderRadius:4},
    {label:'Actual',data:mRows.map(r=>Math.abs(r.act||0)),backgroundColor:alpha(PAL.green,0.7),borderRadius:4}
  ]},options:lineOpts({scales:{y:{ticks:{callback:v=>fmtUSD(v,true)}}}})});

  /* Variance Waterfall */
  const vRows=rows.filter(r=>!r.sec&&r.bud&&r.act);
  makeChart('chart-bva-waterfall',{type:'bar',data:{labels:vRows.map(r=>r.label),datasets:[{label:'Variance ($)',data:vRows.map(r=>r.act-r.bud),backgroundColor:vRows.map(r=>alpha(r.fav(r.act-r.bud)?PAL.green:PAL.red,0.75)),borderRadius:4}]},options:lineOpts({plugins:{legend:{display:false}},scales:{y:{ticks:{callback:v=>fmtUSD(v,true)}}}})});

  /* Monthly variance trend */
  const common=mis.months.filter(m=>bud.months.some(b=>b.trim()===m));
  const revV=common.map(m=>{const a=mis.months.indexOf(m),b=bud.months.findIndex(bm=>bm.trim()===m);return a>=0&&b>=0?(mis.revenue[a]||0)-(bud.revenue[b]||0):0;});
  const ebV =common.map(m=>{const a=mis.months.indexOf(m),b=bud.months.findIndex(bm=>bm.trim()===m);return a>=0&&b>=0?(mis.ebitda[a]||0)-(bud.ebitda[b]||0):0;});
  makeChart('chart-variance-trend',{type:'bar',data:{labels:common,datasets:[
    {label:'Revenue Variance',data:revV,backgroundColor:revV.map(v=>alpha(v>=0?PAL.blue:PAL.red,0.65)),borderRadius:4},
    {label:'EBITDA Variance', data:ebV,type:'line',borderColor:PAL.orange,borderWidth:2.5,fill:false,tension:0.4,pointRadius:4}
  ]},options:lineOpts({scales:{y:{ticks:{callback:v=>fmtUSD(v,true)}}}})});
}

/* ════════════════════════════════════════
   DRILL-DOWN MODAL
════════════════════════════════════════ */
window.openDrilldown=function(key,monthLabel){
  const mis=DATA.mis, bud=DATA.budget;
  const aI=mis.months.indexOf(monthLabel);
  const bI=bud?bud.months.findIndex(m=>m.trim()===monthLabel):-1;
  const pl=DATA.monthPL?.[monthLabel];
  const labels={revenue:'Revenue',cogs:'Total COGS',gm:'Gross Profit',totalCogs:'Total COGS',ebitda:'EBITDA',pbt:'Net Profit',salaries:'Salaries',facility:'Facilities',telecom:'Telecom',transport:'Transport',otherCogs:'Other COGS',totalSGA:'Total SG&A',financeCost:'Finance Charges',grossProfit:'Gross Profit',manpower:'Employee Cost',sellingExp:'Selling Expenses',mgmtSal:'Mgmt Salaries',legalProf:'Legal & Prof',adminExp:'Admin Expenses'};
  const modal=document.getElementById('drilldown-modal');
  const overlay=document.getElementById('drilldown-overlay');
  document.getElementById('drilldown-title').textContent=`${labels[key]||key} — ${monthLabel}`;

  /* Map budget keys → MIS keys (names differ between workbooks) */
  const misKeyMap={salaries:'manpower',facilities:'facility',totalCogs:'cogs',gm:'grossProfit'};
  const misKey=misKeyMap[key]||key;
  /* Map key to MIS series and budget series */
  const misSeries=mis[misKey]||mis['revenue']||[];
  const budSeries=bud?.[key]||[];
  const actVal=aI>=0?(misSeries[aI]||0):0;
  const budVal=bI>=0?(budSeries[bI]||0):0;
  const varVal=actVal-budVal;
  const varPct=budVal?((varVal/Math.abs(budVal))*100):0;

  const subs=pl?getSubItems(pl,key):[];
  let html=`<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:20px">
    <div class="bva-kpi-card"><div class="bva-kpi-label">Budget</div><div class="bva-kpi-actual">${fmtUSD(budVal)}</div></div>
    <div class="bva-kpi-card"><div class="bva-kpi-label">Actual</div><div class="bva-kpi-actual">${fmtUSD(actVal)}</div></div>
    <div class="bva-kpi-card"><div class="bva-kpi-label">Variance</div><div class="bva-kpi-actual ${varVal>=0?'favorable-text':'unfavorable-text'}">${fmtUSD(varVal)} (${fmtPct(varPct)})</div></div>
  </div>
  <div style="margin-bottom:20px">
    <div class="chart-card-header"><h3>Monthly Trend</h3></div>
    <div style="position:relative;height:200px"><canvas id="modal-chart"></canvas></div>
  </div>`;
  if(subs.length){
    html+=`<h4 style="font-size:13px;font-weight:700;margin-bottom:10px">Breakdown</h4>
    <table class="fin-table"><thead><tr><th>Item</th><th class="col-num">Amount</th><th class="col-num">% of Total</th></tr></thead>
    <tbody>${subs.map(s=>`<tr><td>${s.label}</td><td class="col-num">${fmtUSD(s.val)}</td><td class="col-num">${pct(s.val,Math.abs(actVal)||1).toFixed(1)}%</td></tr>`).join('')}</tbody></table>`;
  }
  document.getElementById('drilldown-body').innerHTML=html;
  modal.style.display='flex'; overlay.style.display='block';
  makeChart('modal-chart',{type:'bar',data:{labels:mis.months,datasets:[
    {label:'Budget',data:mis.months.map((_,i)=>{const bi=bud?.months?.findIndex(m=>m.trim()===mis.months[i]);return bi>=0?Math.abs(budSeries[bi]||0):0;}),backgroundColor:alpha(PAL.blue,0.4),borderRadius:3},
    {label:'Actual',data:misSeries.map(v=>Math.abs(v||0)),backgroundColor:alpha(PAL.green,0.7),borderRadius:3}
  ]},options:lineOpts({scales:{y:{ticks:{callback:v=>fmtUSD(v,true)}}}})});
};

document.addEventListener('click',e=>{
  if(e.target.id==='drilldown-close'||e.target.id==='drilldown-overlay'){
    document.getElementById('drilldown-modal').style.display='none';
    document.getElementById('drilldown-overlay').style.display='none';
  }
});

/* ════════════════════════════════════════
   CHART HELPERS
════════════════════════════════════════ */
function lineOpts(extra={}){
  const dark=document.documentElement.dataset.theme==='dark';
  const gridColor=dark?'rgba(255,255,255,0.06)':'rgba(0,0,0,0.06)';
  const tickColor=dark?'#94a3b8':'#475569';
  const tooltipBg=dark?'#1e2533':'#0f172a';
  const borderCol=dark?'rgba(255,255,255,0.1)':'rgba(0,0,0,0.1)';
  const{scales={},plugins={}, ...rest}=extra;
  const baseX={grid:{display:false},ticks:{maxRotation:0,color:tickColor}};
  const baseY={grid:{color:gridColor,drawBorder:false},ticks:{color:tickColor,callback:v=>fmtUSD(v,true)}};
  const mergedScales={};
  const allKeys=new Set(['x','y',...Object.keys(scales)]);
  allKeys.forEach(k=>{
    const base=k==='x'?baseX:k==='y'?baseY:{};
    const ov=scales[k]||{};
    mergedScales[k]={
      ...base,...ov,
      grid:{...base.grid,...(ov.grid||{})},
      ticks:{...base.ticks,...(ov.ticks||{})}
    };
  });
  return{
    responsive:true,maintainAspectRatio:false,
    interaction:{mode:'index',intersect:false},
    plugins:{
      legend:{position:'top',align:'end',labels:{usePointStyle:true,pointStyleWidth:8,padding:14,color:tickColor}},
      tooltip:{backgroundColor:tooltipBg,titleColor:'#f1f5f9',bodyColor:'#94a3b8',borderColor:borderCol,borderWidth:1,padding:12,cornerRadius:8,callbacks:{label:c=>` ${c.dataset.label}: ${fmtUSD(c.raw)}`}},
      ...plugins
    },
    scales:mergedScales,
    ...rest
  };
}
function setBar(barId,value,labelId){
  const bar=document.getElementById(barId),lbl=document.getElementById(labelId);
  const v=isNaN(value)?0:value;
  if(bar) setTimeout(()=>{bar.style.width=Math.min(100,Math.max(0,v))+'%';},100);
  if(lbl) lbl.textContent=fmtPct(v);
}

/* ════════════════════════════════════════
   ANIMATED COUNTER
════════════════════════════════════════ */
function animateCounter(el,target){
  if(!el)return;
  const dur=900,start=performance.now();
  function upd(now){
    const p=Math.min((now-start)/dur,1);
    const e=1-Math.pow(1-p,3);
    el.textContent=fmtUSD(target*e,true);
    if(p<1)requestAnimationFrame(upd);
    else el.textContent=fmtUSD(target,true);
  }
  requestAnimationFrame(upd);
}

/* ════════════════════════════════════════
   SEARCH
════════════════════════════════════════ */
function buildSearchIndex(){
  const idx=[];
  const mis=DATA.mis;
  mis.months.forEach((m,i)=>{
    const items=[['Revenue',mis.revenue[i],'executive'],['Gross Profit',mis.grossProfit[i],'pl'],['EBITDA',mis.ebitda[i],'executive'],['Net PBT',mis.pbt[i],'pl'],['COGS',mis.cogs[i],'pl'],['Employee Cost',mis.manpower[i],'expense'],['Selling Exp',mis.sellingExp[i],'expense'],['Telecom/SW',mis.telecom[i],'expense'],['Admin Exp',mis.adminExp[i],'expense'],['Legal Prof',mis.legalProf[i],'expense']];
    items.forEach(([type,val,page])=>idx.push({type,label:`${type} — ${m}`,value:val||0,page}));
  });
  DATA.searchIndex=idx;
}

function setupSearch(){
  const input=document.getElementById('global-search');
  const dd=document.getElementById('search-dropdown');
  const bar=document.getElementById('search-highlight-bar');
  if(!input)return;
  input.addEventListener('input',()=>{
    const q=input.value.trim().toLowerCase();
    if(!q){dd.innerHTML='';dd.classList.remove('open');return;}
    const results=(DATA.searchIndex||[]).filter(r=>r.label.toLowerCase().includes(q)||r.type.toLowerCase().includes(q)).slice(0,10);
    if(!results.length){dd.innerHTML='<div class="search-result-item" style="color:var(--text-muted)">No results</div>';dd.classList.add('open');return;}
    dd.innerHTML=results.map((r,i)=>`<div class="search-result-item" data-i="${i}"><span class="sri-type">${r.type}</span><span>${r.label}</span><span class="sri-value">${fmtUSD(r.value,true)}</span></div>`).join('');
    dd.classList.add('open');
    dd.querySelectorAll('.search-result-item').forEach((el,i)=>{
      el.addEventListener('click',()=>{
        dd.classList.remove('open');input.value='';
        navigateTo(results[i].page);
        if(bar){bar.style.display='flex';document.getElementById('search-result-info').textContent=`"${results[i].label}" = ${fmtUSD(results[i].value)}`;}
      });
    });
  });
  document.addEventListener('click',e=>{if(!e.target.closest('.search-box'))dd.classList.remove('open');});
  document.getElementById('search-clear-btn')?.addEventListener('click',()=>{if(bar)bar.style.display='none';input.value='';});
  /* P&L inline search */
  document.getElementById('pl-table-search')?.addEventListener('input',function(){
    const q=this.value.toLowerCase();
    document.querySelectorAll('#pl-tbody tr').forEach(tr=>{
      tr.style.display=q&&!tr.textContent.toLowerCase().includes(q)?'none':'';
      tr.classList.toggle('search-highlight',!!(q&&tr.textContent.toLowerCase().includes(q)));
    });
  });
}

/* ════════════════════════════════════════
   EXPORTS
════════════════════════════════════════ */
function setupExports(){
  document.getElementById('btn-print')?.addEventListener('click',()=>window.print());
  document.getElementById('btn-export-pdf')?.addEventListener('click',()=>window.print());
  document.getElementById('btn-export-excel')?.addEventListener('click',()=>{
    const mis=DATA.mis;
    const wb2=XLSX.utils.book_new();
    const sumData=[['Particulars',...mis.months],['Revenue',...mis.revenue],['Gross Profit',...mis.grossProfit],['EBITDA',...mis.ebitda],['Net PBT',...mis.pbt],['COGS',...mis.cogs],['Manpower',...mis.manpower],['SG&A',...mis.totalSGA]];
    XLSX.utils.book_append_sheet(wb2,XLSX.utils.aoa_to_sheet(sumData),'Summary');
    if(DATA.budget){
      const bud=DATA.budget;
      const bvaData=[['Month','Metric','Budget','Actual','Variance','Var%']];
      mis.months.forEach((m,aI)=>{
        const bI=bud.months.findIndex(bm=>bm.trim()===m);
        if(bI>=0){
          [['Revenue',bud.revenue[bI],mis.revenue[aI]],['EBITDA',bud.ebitda[bI],mis.ebitda[aI]],['PBT',bud.pbt[bI],mis.pbt[aI]]].forEach(([lbl,b,a])=>bvaData.push([m,lbl,b,a,a-b,b?((a-b)/Math.abs(b)*100).toFixed(1)+'%':'—']));
        }
      });
      XLSX.utils.book_append_sheet(wb2,XLSX.utils.aoa_to_sheet(bvaData),'Budget vs Actual');
    }
    XLSX.writeFile(wb2,'iMagnum_FPA_Dashboard.xlsx');
  });
}

/* ════════════════════════════════════════
   NAVIGATION
════════════════════════════════════════ */
let currentPage='executive';
function navigateTo(pageId){
  currentPage=pageId;
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n=>n.classList.remove('active'));
  const page=document.getElementById(`page-${pageId}`);
  if(page)page.classList.add('active');
  const nav=document.querySelector(`.nav-item[data-page="${pageId}"]`);
  if(nav)nav.classList.add('active');
  const bc=document.getElementById('breadcrumb-page');
  if(bc)bc.textContent=page?.dataset.title||pageId;
  if(window.innerWidth<=900){document.getElementById('sidebar').classList.remove('open');document.getElementById('sidebar-overlay').classList.remove('visible');}
}

function setupNavigation(){
  document.querySelectorAll('.nav-item').forEach(btn=>btn.addEventListener('click',()=>navigateTo(btn.dataset.page)));
  document.getElementById('menu-toggle')?.addEventListener('click',()=>{document.getElementById('sidebar').classList.toggle('open');document.getElementById('sidebar-overlay').classList.toggle('visible');});
  document.getElementById('sidebar-close')?.addEventListener('click',()=>{document.getElementById('sidebar').classList.remove('open');document.getElementById('sidebar-overlay').classList.remove('visible');});
  document.getElementById('sidebar-overlay')?.addEventListener('click',()=>{document.getElementById('sidebar').classList.remove('open');document.getElementById('sidebar-overlay').classList.remove('visible');});
}

/* ════════════════════════════════════════
   POPULATE MONTH DROPDOWNS
════════════════════════════════════════ */
function populateSelects(){
  const months=DATA.mis.months;
  const latest=months[months.length-1];
  ['pl-month-select','exp-month-select','bva-month-select'].forEach(id=>{
    const el=document.getElementById(id);
    if(!el)return;
    el.innerHTML=months.map(m=>`<option value="${m}"${m===latest?' selected':''}>${m}</option>`).join('');
  });
  document.getElementById('pl-month-select')?.addEventListener('change',e=>renderPL(e.target.value));
  document.getElementById('exp-month-select')?.addEventListener('change',e=>renderExpenseAnalysis(e.target.value,document.getElementById('exp-entity-select')?.value));
  document.getElementById('exp-entity-select')?.addEventListener('change',e=>renderExpenseAnalysis(document.getElementById('exp-month-select')?.value,e.target.value));
  document.getElementById('bva-month-select')?.addEventListener('change',e=>renderBudgetVsActual(e.target.value));
}

/* ════════════════════════════════════════
   DARK MODE
════════════════════════════════════════ */
function setupTheme(){
  const saved=localStorage.getItem('imag-theme')||'light';
  document.documentElement.dataset.theme=saved;
  document.getElementById('theme-toggle')?.addEventListener('click',()=>{
    const next=document.documentElement.dataset.theme==='dark'?'light':'dark';
    document.documentElement.dataset.theme=next;
    localStorage.setItem('imag-theme',next);
    applyChartDefaults();
    Object.values(CHARTS).forEach(c=>{if(c)c.update();});
  });
}

/* ════════════════════════════════════════
   MAIN INIT
════════════════════════════════════════ */
async function init(){
  try{
    applyChartDefaults();
    WB.mis   =await loadWorkbook(CFG.files.mis,   'MIS Dashboard');
    WB.budget=await loadWorkbook(CFG.files.budget, 'Budget');
    setStatus('Extracting financial data…');
    DATA.mis   =extractMIS();
    DATA.budget=extractBudget();
    setStatus('Processing monthly P&L…');
    DATA.monthPL={};
    for(const m of DATA.mis.months) DATA.monthPL[m]=extractMonthPL(m);
    buildSearchIndex();
    setStatus('Rendering dashboard…');
    populateSelects();
    const latest=DATA.mis.months[DATA.mis.months.length-1];
    renderExecutive();
    renderPL(latest);
    renderExpenseAnalysis(latest,'consolidated');
    renderBudgetVsActual(latest);
    setupNavigation();
    setupSearch();
    setupExports();
    setupTheme();
    /* Freshness */
    const ft=document.getElementById('freshness-text');
    if(ft)ft.textContent='Updated '+new Date().toLocaleTimeString('en-US',{hour:'2-digit',minute:'2-digit'});
    /* Show app */
    const ls=document.getElementById('loading-screen');
    ls.style.opacity='0'; ls.style.transition='opacity 0.5s';
    setTimeout(()=>ls.style.display='none',500);
    const app=document.getElementById('app');
    app.classList.remove('app-hidden'); app.classList.add('app-visible');
  } catch(err){
    console.error('Dashboard error:',err);
    document.getElementById('loading-screen').innerHTML=`<div style="text-align:center;padding:40px;color:#f1f5f9"><div style="font-size:48px;margin-bottom:16px">⚠️</div><h2 style="margin-bottom:10px">Unable to load data</h2><p style="color:#94a3b8;margin-bottom:6px">Ensure Excel files are in the <code style="background:rgba(255,255,255,0.1);padding:2px 6px;border-radius:4px">data/</code> folder</p><p style="color:#64748b;font-size:13px">${err.message}</p><button onclick="location.reload()" style="margin-top:20px;padding:10px 24px;background:#2563eb;color:#fff;border:none;border-radius:8px;cursor:pointer;font-size:14px">↺ Retry</button></div>`;
  }
}

if(document.readyState==='loading') document.addEventListener('DOMContentLoaded',init);
else init();
