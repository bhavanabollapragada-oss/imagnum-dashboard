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
    mis:    './data/Imagnum_MIS_Consolidated_Jan_May_2026.xlsx',
    budget: './data/Imagnum_MIS_Consolidated_Jan_May_2026.xlsx'
  },
  sheets: {
    mis: {
      consolidated: 'MIS Actuals Jan-Jun',
      ytdLabel: 'Jan-Jun Total',
      /* partCol / indiaCol / usCol / consCol are ALL 0-indexed (sheet_to_json array positions) */
      monthMap: {
        'Jan-26': { sheet: 'Jan',         partCol:1, indiaCol:3, usCol:2, consCol:4 },
        'Feb-26': { sheet: 'March-Feb',   partCol:1, indiaCol:6, usCol:7, consCol:8 },
        'Mar-26': { sheet: 'March-Feb',   partCol:1, indiaCol:2, usCol:3, consCol:4 },
        'Apr-26': { sheet: 'April-March', partCol:1, indiaCol:2, usCol:3, consCol:5 },
        'May-26': { sheet: 'May-April',   partCol:1, indiaCol:2, usCol:3, consCol:5 },
        'Jun-26': { sheet: "Jun'26",      partCol:1, indiaCol:2, usCol:3, consCol:5 }
      }
    },
    budget: { plSummary: 'P&L - Bud.', bva: 'Bud vs Act - May & Jun' },
    /* Budget vs Actual is sourced from the Proforma P&L of each month.
       Budget  = Proforma forecast column · Actual = MIS Actuals Jan-Jun.
       Only the months listed here appear on the Budget vs Actual page. */
    /* Proforma sheets are auto-discovered by name at load time — see
       discoverProforma(). Any sheet called "Proforma-<Month>" (spacing and
       punctuation are ignored, so "Proforma- July" resolves too) becomes a
       Budget vs Actual month, provided MIS actuals exist for it. */
    proforma: {}
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
    empStat:      ['Employee statutory'],
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
    revenue:     ['Revenue'],
    salaries:    ['Salaries'],
    facilities:  ['Facilities'],
    telecom:     ['Telcom/Data', 'Telecom'],
    transport:   ['Transport'],
    otherCogs:   ['Other Cogs', 'Other COGS'],
    totalCogs:   ['Total Cogs', 'Total COGS'],
    gm:          ['GM'],
    totalSGA:    ['Overall SG&A', 'Overall Sg&A'],
    ebitda:      ['Ebidta', 'EBITDA'],
    financeCost: ['Interest/Finance Charges'],
    pbt:         ['PBT']
  },
  /* Every cost group carried on "MIS Actuals Jan-Jun".
     COGS members sum to Total COGS; SG&A members sum to Overall SG&A.
     The tie-out is asserted at render time and shown under the mix chart. */
  expenseCategories: [
    { key:'manpower',  label:'Manpower',             color:'#2563eb', grp:'COGS' },
    { key:'empStat',   label:'Employee Statutory',   color:'#3b82f6', grp:'COGS' },
    { key:'facility',  label:'Facility Cost',        color:'#7c3aed', grp:'COGS' },
    { key:'telecom',   label:'Telcom/Data & others', color:'#0891b2', grp:'COGS' },
    { key:'transport', label:'Transport',            color:'#059669', grp:'COGS' },
    { key:'otherCogs', label:'Other COGS',           color:'#65a30d', grp:'COGS' },
    { key:'sellingExp',label:'Selling Expenses',     color:'#d97706', grp:'SG&A' },
    { key:'mgmtSal',   label:'Salaries-Management',  color:'#dc2626', grp:'SG&A' },
    { key:'legalProf', label:'Legal & Professional', color:'#4338ca', grp:'SG&A' },
    { key:'gaSal',     label:'G&A Salaries',         color:'#db2777', grp:'SG&A' },
    { key:'adminExp',  label:'Gen & Admin Expenses', color:'#64748b', grp:'SG&A' }
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
/* Find FIRST row matching patterns in the given column.
   Patterns are tried in the order given — a pattern listed first wins even
   if a later pattern matches a row higher up the sheet. Without this,
   CFG.labels.financeCost's "Finance Charges" fallback matched a zero line
   in Gen & Admin and shadowed the real "Finance Cost" group below it. */
function findRowIdx(rows, patterns, colIdx=0) {
  for (const p of patterns) {
    for (let i=0;i<rows.length;i++) {
      if (labelMatch(rows[i][colIdx], [p])) return i;
    }
  }
  return -1;
}
/* Exact (not substring) label match — needed where a component line
   contains the group's name, e.g. "Other Cogs- MD" vs "Other COGS". */
function findRowIdxExact(rows, label, colIdx=0) {
  const t=String(label).replace(/\s+/g,' ').trim().toLowerCase();
  for (let i=0;i<rows.length;i++) {
    if (cleanLabel(rows[i][colIdx]).toLowerCase()===t) return i;
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
   GROUP-BLOCK PARSER
   In every MIS / month pack a cost GROUP header carries leading
   whitespace ("        Facility Cost") while its constituent lines
   do not. That indentation — not label text — defines the hierarchy,
   so drill-downs survive line items being added, removed or renamed.
   A block ends at the next indented header or at a total/section row.
════════════════════════════════════════ */
const BLOCK_STOP=new Set(['total cogs','total cost','total','gm','gm%','gp','gp%',
  'gross profit','sg&a','g&a expenses','overall sg&a','ebitda','ebitda%','pbt',
  'service income','overall operational cost','below ebitda','indirect incomes',
  'total revenue','particulars']);
function cleanLabel(v){
  return String(v===null||v===undefined?'':v).replace(/\u00a0/g,' ').replace(/\s+/g,' ').trim();
}
function isIndented(v){
  const s=String(v===null||v===undefined?'':v).replace(/\u00a0/g,' ');
  return /^\s/.test(s) && s.trim()!=='';
}
function isStopRow(v){ return BLOCK_STOP.has(cleanLabel(v).toLowerCase()); }

/* → {hdr, subs:[rowIdx,…]} for the group whose header matches `patterns` */
function groupBlock(rows, patterns, partCol){
  const hdr=findRowIdx(rows,patterns,partCol);
  if(hdr===-1) return {hdr:-1,subs:[]};
  /* No indented header ⇒ this pack carries summary lines only (January).
     Without this guard the next sibling summary lines get adopted as
     children and the page shows a hierarchy the workbook doesn't have. */
  if(!isIndented(rows[hdr][partCol])) return {hdr,subs:[]};
  const subs=[];
  for(let i=hdr+1;i<rows.length;i++){
    const cell=rows[i][partCol];
    if(cleanLabel(cell)==='') continue;
    if(isIndented(cell)||isStopRow(cell)) break;
    subs.push(i);
  }
  return {hdr,subs};
}
function blockSum(rows,block,col){
  return block.subs.reduce((a,i)=>a+parseNum(rows[i][col]),0);
}

/* Manpower absorbs the statutory block on some packs (January) and sits
   beside it on others (Feb–Jun). Detect which, so the two are never
   double-counted and never silently dropped. */
function splitManpower(rows,partCol,col){
  const mp=groupBlock(rows,CFG.labels.manpower,partCol);
  const st=groupBlock(rows,CFG.labels.empStat, partCol);
  const mpHdr = mp.hdr>=0?parseNum(rows[mp.hdr][col]):0;
  const stHdr = st.hdr>=0?parseNum(rows[st.hdr][col]):0;
  const mpSubs=blockSum(rows,mp,col);
  const stSubs=blockSum(rows,st,col);
  /* Values arrive as display text (whole USD), so a sum of sub-lines drifts
     a dollar or two from its group total. An absolute tolerance therefore
     misfires — instead ask which hypothesis explains the header better:
     Manpower = subs + statutory, or Manpower = subs alone. */
  const absorbed = stSubs!==0 &&
    Math.abs(mpHdr-(mpSubs+stSubs)) < Math.abs(mpHdr-mpSubs);
  return absorbed
    ? {manpower:mpHdr-stSubs, empStat:stSubs,          absorbed:true}
    : {manpower:mpHdr,        empStat:(stHdr||stSubs), absorbed:false};
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
  /* YTD column — labelled "Jan-Jun Total"; falls back to summing the months */
  let ytdCol=-1;
  for(let c=1;c<hdr.length;c++){
    if(/total/i.test(String(hdr[c]||''))) { ytdCol=c; break; }
  }
  function ytdVal(patterns,series){
    /* Deliberately NOT read from the sheet's own Total column. That column is
       a fixed-range formula and is not widened when a month is inserted, so
       after July was added it still summed Jan–Jun. Summing the detected
       months keeps YTD consistent with the per-month figures automatically. */
    return series.reduce((a,b)=>a+(b||0),0);
  }
  /* Cross-check against the sheet's Total column and warn on a mismatch */
  function ytdSheet(patterns){ return ytdCol>=0?getVal(rows,patterns,ytdCol,0):null; }
  const out={
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
  out.ytdLabel = months.length
    ? `${months[0].label}–${months[months.length-1].label} Total`
    : 'YTD Total';
  /* Surface a stale Total column rather than silently disagreeing with it */
  const shRev=ytdSheet(CFG.labels.revenue);
  out.ytdSheetRevenue=shRev;
  out.ytdStale = shRev!==null && Math.abs(shRev-(out.revenue||[]).reduce((a,b)=>a+(b||0),0))>2;
  out.ytd = {};
  Object.keys(CFG.labels).forEach(k=>{ out.ytd[k]=ytdVal(CFG.labels[k], out[k]||[]); });

  /* Manpower / Employee-Statutory de-overlap, month by month.
     The YTD column cannot be read directly for these two because the
     January pack folds statutory into Manpower and later packs do not —
     so YTD is summed from the corrected monthly series instead. */
  const split=months.map(m=>splitManpower(rows,0,m.col));
  out.manpower = split.map(s=>s.manpower);
  out.empStat  = split.map(s=>s.empStat);
  out.absorbed = split.map(s=>s.absorbed);
  out.ytd.manpower = out.manpower.reduce((a,b)=>a+b,0);
  out.ytd.empStat  = out.empStat.reduce((a,b)=>a+b,0);

  /* Sub-line detail for every group, every month — drives Expense Analysis */
  out.groups={};
  CFG.expenseCategories.forEach(cat=>{
    const b=groupBlock(rows,CFG.labels[cat.key],0);
    out.groups[cat.key]=b.subs
      .map(i=>({label:cleanLabel(rows[i][0]), values:months.map(m=>parseNum(rows[i][m.col]))}))
      .filter(s=>s.values.some(v=>v!==0));
  });
  return out;
}

/* ════════════════════════════════════════
   EXTRACT — Individual Month Sheets (India/US/Consol)
   Uses CFG column indices DIRECTLY — no dynamic override
   (avoids bug on combined Feb/Mar sheet where both months
    share one sheet with different column positions)
════════════════════════════════════════ */
function extractMonthPL(monthLabel) {
  /* Explicit mapping wins (it disambiguates the paired packs, where two
     months share one sheet); otherwise discover the pack by name. */
  let mc=CFG.sheets.mis.monthMap[monthLabel];
  if(!mc){
    mc=discoverMonthPack(WB.mis,monthLabel);
    if(mc) CFG.sheets.mis.monthMap[monthLabel]=mc;   /* cache */
  }
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
  /* Sub-line detail per group, split India / US / Consolidated.
     January's pack carries summary lines only, so every block comes back
     empty there and the P&L page renders those rows non-expandable. */
  result.groups={};
  ['manpower','empStat','facility','telecom','transport','otherCogs',
   'sellingExp','mgmtSal','legalProf','gaSal','adminExp'].forEach(k=>{
    const b=groupBlock(rows,CFG.labels[k],partCol);
    result.groups[k]=b.subs.map(i=>({
      label:cleanLabel(rows[i][partCol]),
      india:parseNum(rows[i][indiaCol]),
      us:   parseNum(rows[i][usCol]),
      cons: parseNum(rows[i][consCol])
    })).filter(r=>r.india||r.us||r.cons);
  });
  const spI=splitManpower(rows,partCol,indiaCol),
        spU=splitManpower(rows,partCol,usCol),
        spC=splitManpower(rows,partCol,consCol);

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
  result.summary.manpower={india:spI.manpower,us:spU.manpower,cons:spC.manpower};
  result.summary.empStat ={india:spI.empStat, us:spU.empStat, cons:spC.empStat};
  return result;
}

/* ════════════════════════════════════════
   SHEET DISCOVERY
   Month packs and Proforma sheets are named inconsistently across the
   workbook ("Jun'26", "May-April", "Proforma-June", "Proforma- July").
   Rather than hard-code each new one, resolve them by pattern so a month
   added to the workbook flows through with no code change.
════════════════════════════════════════ */
const MONTH_NAMES=['january','february','march','april','may','june',
                   'july','august','september','october','november','december'];
const MON_ABBR=['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];

/* "Jun-26" → 5 (0-based month index) */
function monthIndexOf(label){
  const m=String(label||'').trim().toLowerCase().slice(0,3);
  return MON_ABBR.indexOf(m);
}

/* Find the P&L pack for a month label, e.g. "Jul-26" → sheet "Jul'26".
   Returns {sheet, partCol, indiaCol, usCol, consCol} or null. Columns are
   read from the pack's own header row, because their order differs between
   packs (some put Total before Elimination). */
function discoverMonthPack(wb, monthLabel){
  const mi=monthIndexOf(monthLabel);
  if(mi<0) return null;
  const abbr=MON_ABBR[mi], full=MONTH_NAMES[mi];
  const yy=(String(monthLabel).match(/(\d{2})$/)||[])[1]||'';
  const norm=n=>String(n).toLowerCase().replace(/[^a-z0-9]/g,'');
  const want=[norm(abbr+yy), norm(abbr+"'"+yy), norm(full+yy), norm(abbr), norm(full)];
  let name=null;
  for(const w of want){
    name=(wb.SheetNames||[]).find(n=>norm(n)===w);
    if(name) break;
  }
  if(!name) return null;

  const rows=sheetToArr(wb,name);
  /* header row = the one containing "particulars" */
  let hr=-1, partCol=0;
  for(let i=0;i<Math.min(rows.length,20)&&hr<0;i++){
    for(let c=0;c<(rows[i]||[]).length;c++){
      if(cleanLabel(rows[i][c]).toLowerCase()==='particulars'){ hr=i; partCol=c; break; }
    }
  }
  if(hr<0) return null;
  const hdr=rows[hr];
  const findCol=(re,from)=>{
    for(let c=from;c<hdr.length;c++) if(re.test(cleanLabel(hdr[c]))) return c;
    return -1;
  };
  const indiaCol=findCol(/india/i,partCol+1);
  const usCol   =findCol(/\bus\b|u\.s\./i,partCol+1);
  const consCol =findCol(/^total$/i,partCol+1);
  if(indiaCol<0||usCol<0||consCol<0) return null;
  return {sheet:name,partCol,indiaCol,usCol,consCol,discovered:true};
}

/* Every "Proforma-<Month>" sheet in the workbook → {'Jul-26':'Proforma- July'} */
function discoverProforma(wb, monthLabels){
  const found={};
  (wb.SheetNames||[]).forEach(n=>{
    const t=String(n).toLowerCase().replace(/[^a-z]/g,'');
    if(t.indexOf('proforma')!==0) return;
    const rest=t.slice('proforma'.length);
    const mi=MONTH_NAMES.findIndex(m=>rest===m||rest===m.slice(0,3));
    if(mi<0) return;
    const label=monthLabels.find(L=>monthIndexOf(L)===mi);
    if(label) found[label]=n;
  });
  return found;
}

/* ════════════════════════════════════════
   EXTRACT — Proforma P&L (BUDGET side of Budget vs Actual)
   Budget is read from that month's own Proforma pack, forecast column.
   Two reclassifications are applied so the lines are comparable with
   the MIS actuals they are measured against:
     · Software Expenses → out of Telcom/Data, into Other COGS
     · Interest on OD    → out of Other COGS,  into Finance Cost
   Without these, Telecom and Other COGS both show large offsetting
   variances that cancel at the Total COGS line and mean nothing.
════════════════════════════════════════ */
function extractProforma(){
  const out={};
  const map=CFG.sheets.proforma;
  Object.keys(map).forEach(month=>{
    const name=map[month];
    const rows=sheetToArr(WB.budget,name);
    if(!rows.length){ console.warn('Proforma sheet missing:',name); return; }

    /* Forecast column = the column headed "Forecast" (Actuals sits left of it) */
    let fc=-1;
    for(let r=0;r<Math.min(rows.length,12)&&fc<0;r++){
      for(let c=1;c<(rows[r]||[]).length;c++){
        if(cleanLabel(rows[r][c]).toLowerCase()==='forecast'){ fc=c; break; }
      }
    }
    if(fc<0){ console.warn('Forecast column not found on',name); return; }

    const g=pats=>getVal(rows,pats,fc,0);
    /* Every figure below is a GROUP TOTAL, and each Proforma pack contains
       component lines whose names embed the group's name — "Chennai New
       Facility Cost" above "Facility Cost", "Other Cogs- MD" above
       "Other COGS". A substring match grabs the component and silently
       understates the budget, so match the exact label first and only fall
       back to substring when no exact label exists. */
    const gx=(label,fallback)=>{
      const i=findRowIdxExact(rows,label,0);
      return i>=0?parseNum(rows[i][fc]):g(fallback||[label]);
    };
    const revenue    = gx('Total Revenue');
    const manpower   = gx('Manpower');
    const facility   = gx('Facility Cost');
    const telecomRaw = gx('Telcom/Data & others',['Telcom/Data','Telecom/Data']);
    const software   = gx('Software Expenses');
    const transport  = gx('Transport');
    const otherRaw   = gx('Other COGS',['Other COGS','Other Cogs']);
    const interestOD = gx('Interest on OD');
    const totalOpRaw = gx('Total Operational Expenses');
    const sellingExp = gx('Selling Expenses');
    const mgmtSal    = gx('Salaries-Management & Others',['Salaries-Management']);
    const fpo        = gx('FPO Charges - SRF',['FPO Charges']);
    const legalRaw   = gx('Legal & Professional Fees');
    const gaSal      = gx('G&A salaries');
    const adminExp   = gx('Gen & Admin. Expenses',['Gen & Admin. Expenses','Gen & Admin Expenses']);
    const totalSGA   = gx('Total SG&A Expenses');
    const ebitdaRaw  = gx('EBITDA');

    const telecom  = telecomRaw - software;              /* software moved out */
    const otherCogs= otherRaw - interestOD + software;   /* software in, OD out */
    const totalCogs= totalOpRaw - interestOD;
    const gm       = revenue - totalCogs;

    out[month]={
      _sheet:name,
      revenue, manpower, facility, telecom, transport, otherCogs,
      totalCogs, gm, sellingExp, mgmtSal,
      legalProf: fpo + legalRaw,                         /* FPO + Legal & Prof */
      gaSal, adminExp, totalSGA,
      ebitda:      ebitdaRaw + interestOD,               /* OD out of EBITDA */
      financeCost: interestOD,
      pbt:         ebitdaRaw,
      _reclass:{software, interestOD}
    };
  });
  return out;
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
function renderExecutive(sel) {
  const mis=DATA.mis; const months=mis.months; const n=months.length;
  if(!n)return;

  /* `sel` is a month label ("Mar-26") or 'ytd'. Anything unknown falls
     back to the latest closed month. */
  const ytdMode = (sel==='ytd');
  let li = ytdMode ? -1 : months.indexOf(sel);
  if(!ytdMode && li===-1) li = n-1;
  const pi = (!ytdMode && li>0) ? li-1 : -1;
  const periodLabel = ytdMode ? (mis.ytdLabel||'YTD') : months[li];

  /* Value of a series for the selected period */
  const at = (key) => ytdMode
    ? (mis.ytd&&mis.ytd[key]!==undefined ? mis.ytd[key]
       : (mis[key]||[]).reduce((a,b)=>a+(b||0),0))
    : ((mis[key]||[])[li]||0);

  /* Animated KPI cards — sparkline always shows the full period */
  function kpi(vId,bId,spId,key,color=PAL.blue){
    const series=mis[key]||[];
    const curr=at(key);
    const chg=(pi>=0)?mom(curr,series[pi]||0):null;
    animateCounter(document.getElementById(vId),curr);
    const b=document.getElementById(bId);
    if(b){
      b.textContent=chg===null?'—':fmtPct(chg);
      b.className='kpi-badge'+(chg===null?' neutral':chg>=0?'':' negative');
    }
    makeChart(spId,{type:'line',data:{labels:months,datasets:[{data:series,borderColor:color,borderWidth:2,fill:true,backgroundColor:alpha(color,0.1),tension:0.4,pointRadius:0}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{enabled:false}},scales:{x:{display:false},y:{display:false}},animation:{duration:900}}});
  }
  kpi('v-revenue','b-revenue','spark-revenue','revenue',    PAL.blue);
  kpi('v-gp',     'b-gp',     'spark-gp',     'grossProfit',PAL.green);
  kpi('v-ebitda', 'b-ebitda', 'spark-ebitda', 'ebitda',     PAL.purple);

  /* MoM caption under each KPI */
  document.querySelectorAll('.kpi-card .kpi-badge').forEach(b=>{
    const cap=b.parentElement&&b.parentElement.querySelector('.kpi-change-label');
    if(cap) cap.textContent = pi>=0 ? `vs ${months[pi]}` : (ytdMode?'cumulative':'no prior month');
  });

  /* Margin strips — selected period */
  const rev=at('revenue')||1, gp=at('grossProfit'), eb=at('ebitda');
  setBar('ms-gm',    pct(gp,rev),'ms-gm-val');
  setBar('ms-ebitda',pct(eb,rev),'ms-ebitda-val');

  /* Revenue + GP Trend — full period, unaffected by the selector */
  makeChart('chart-rev-trend',{type:'bar',data:{labels:months,datasets:[
    {label:'Revenue',data:mis.revenue,backgroundColor:alpha(PAL.blue,0.2),borderColor:PAL.blue,borderWidth:2,borderRadius:4,type:'bar'},
    {label:'Gross Profit',data:mis.grossProfit,borderColor:PAL.green,borderWidth:2.5,fill:false,tension:0.4,type:'line',pointRadius:4,pointBackgroundColor:PAL.green}
  ]},options:lineOpts({scales:{y:{ticks:{callback:v=>fmtUSD(v,true)}}}})});

  /* Expense Distribution donut — selected period, every cost group */
  const cats=CFG.expenseCategories;
  const expL=cats.map(c=>c.label);
  const expV=cats.map(c=>Math.abs(at(c.key)||0));
  makeChart('chart-exp-dist',{type:'doughnut',data:{labels:expL,datasets:[{data:expV,backgroundColor:cats.map(c=>c.color),borderWidth:0,hoverOffset:6}]},options:{responsive:true,maintainAspectRatio:false,cutout:'60%',plugins:{legend:{position:'right',labels:{font:{size:10},padding:6,boxWidth:10}},tooltip:{callbacks:{label:c=>` ${c.label}: ${fmtUSD(c.raw)}`}}}}});
  const expSub=document.getElementById('exec-expdist-sub');
  if(expSub) expSub.textContent=periodLabel;

  /* Waterfall — selected period */
  renderWaterfall('chart-waterfall',periodLabel,{
    revenue:at('revenue'), cogs:at('cogs'), grossProfit:at('grossProfit'),
    totalSGA:at('totalSGA'), ebitda:at('ebitda')
  });

  /* EBITDA Performance — full period */
  const ebPcts=mis.revenue.map((r,i)=>r?pct(mis.ebitda[i],r):0);
  makeChart('chart-ebitda-perf',{type:'bar',data:{labels:months,datasets:[
    {label:'EBITDA ($)',data:mis.ebitda,backgroundColor:mis.ebitda.map(v=>alpha(v>=0?PAL.purple:PAL.red,0.7)),borderColor:mis.ebitda.map(v=>v>=0?PAL.purple:PAL.red),borderWidth:1,borderRadius:4,yAxisID:'y'},
    {label:'EBITDA %',data:ebPcts,type:'line',borderColor:PAL.orange,borderWidth:2,fill:false,tension:0.4,pointRadius:4,yAxisID:'y1'}
  ]},options:lineOpts({scales:{y:{ticks:{callback:v=>fmtUSD(v,true)},position:'left'},y1:{ticks:{callback:v=>`${v.toFixed(0)}%`},position:'right',grid:{display:false}}}})});

  /* Margin Trend — full period */
  const gmP=mis.revenue.map((r,i)=>r?pct(mis.grossProfit[i],r):0);
  const ebP=mis.revenue.map((r,i)=>r?pct(mis.ebitda[i],r):0);
  makeChart('chart-margin-trend',{type:'line',data:{labels:months,datasets:[
    {label:'Gross Margin %',data:gmP,borderColor:PAL.blue,backgroundColor:alpha(PAL.blue,0.05),fill:true,tension:0.4,borderWidth:2.5,pointRadius:4},
    {label:'EBITDA %',data:ebP,borderColor:PAL.purple,borderDash:[5,3],fill:false,tension:0.4,borderWidth:2,pointRadius:3}
  ]},options:lineOpts({scales:{y:{ticks:{callback:v=>`${v.toFixed(0)}%`}}}})});

  const dr=document.getElementById('exec-date-range');
  if(dr)dr.textContent = ytdMode
    ? `${months[0]} – ${months[n-1]}`
    : `${periodLabel} (period ${li+1} of ${n})`;
}

function renderWaterfall(cId,label,v){
  const rev=v.revenue||0, cogs=Math.abs(v.cogs||0), gp=v.grossProfit||0;
  const sga=Math.abs(v.totalSGA||0), eb=v.ebitda||0;
  const el=document.getElementById('waterfall-label');
  if(el)el.textContent=label;
  const steps=[
    {l:'Revenue',    v:rev,  b:0,      c:PAL.blue,   total:false},
    {l:'- COGS',     v:-cogs,b:rev-cogs,c:PAL.red,    total:false},
    {l:'Gross Profit',v:gp, b:0,       c:PAL.green,  total:true},
    {l:'- SG&A',     v:-sga, b:gp,     c:PAL.orange, total:false},
    {l:'EBITDA',     v:eb,   b:0,       c:eb>=0?PAL.purple:PAL.red, total:true}
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

  /* P&L structure rows. `grp` marks a row whose sub-lines can be
     expanded in place; January's pack has no sub-lines so those rows
     render flat. */
  const n=(i,u,c)=>({india:i||0,us:u||0,cons:c||0});
  const G=DATA.monthPL&&DATA.monthPL[monthLabel]?DATA.monthPL[monthLabel].groups:(plData.groups||{});
  const structure=[
    {sec:'REVENUE'},
    {label:'Service Revenue',  ...n(s.revenue.india,    s.revenue.us,    s.revenue.cons),    cls:'row-subtotal'},
    {sec:'COST OF GOODS SOLD'},
    {label:'Manpower',           ...n(s.manpower.india, s.manpower.us,  s.manpower.cons),  indent:1, grp:'manpower'},
    {label:'Employee Statutory', ...n(s.empStat.india,  s.empStat.us,   s.empStat.cons),   indent:1, grp:'empStat'},
    {label:'Facility Cost',    ...n(s.facility.india,   s.facility.us,   s.facility.cons),   indent:1, grp:'facility'},
    {label:'Telcom/Data & others',...n(s.telecom.india, s.telecom.us,    s.telecom.cons),    indent:1, grp:'telecom'},
    {label:'Transport',        ...n(s.transport.india,  s.transport.us,  s.transport.cons),  indent:1, grp:'transport'},
    {label:'Other COGS',       ...n(s.otherCogs.india,  s.otherCogs.us,  s.otherCogs.cons),  indent:1, grp:'otherCogs'},
    {label:'Total COGS',       ...n(s.cogs.india,       s.cogs.us,       s.cogs.cons),       cls:'row-subtotal'},
    {label:'GROSS PROFIT',     ...n(s.grossProfit.india,s.grossProfit.us,s.grossProfit.cons),cls:'row-total'},
    {sec:'SG&A EXPENSES'},
    {label:'Selling Expenses',   ...n(s.sellingExp.india,s.sellingExp.us, s.sellingExp.cons), indent:1, grp:'sellingExp'},
    {label:'Salaries-Management',...n(s.mgmtSal.india,  s.mgmtSal.us,   s.mgmtSal.cons),    indent:1, grp:'mgmtSal'},
    {label:'Legal & Professional',...n(s.legalProf.india,s.legalProf.us, s.legalProf.cons),  indent:1, grp:'legalProf'},
    {label:'G&A Salaries',     ...n(s.gaSal.india,      s.gaSal.us,      s.gaSal.cons),      indent:1, grp:'gaSal'},
    {label:'Gen & Admin Expenses',...n(s.adminExp.india,s.adminExp.us,   s.adminExp.cons),   indent:1, grp:'adminExp'},
    {label:'Total SG&A',       ...n(s.totalSGA.india,   s.totalSGA.us,   s.totalSGA.cons),   cls:'row-subtotal'},
    {label:'EBITDA',           ...n(s.ebitda.india,     s.ebitda.us,     s.ebitda.cons),     cls:'row-grand'}
  ];

  const nc=v=>v<0?'col-num num-val negative':v>0?'col-num num-val':'col-num num-val text-muted';
  let gi=0;
  tbody.innerHTML=structure.filter(row=>{
    if(row.sec||!row.grp) return true;
    /* drop a group that is zero across all three columns and has no detail —
       e.g. Employee Statutory in January, where it is folded into Manpower */
    const has=(G[row.grp]||[]).length;
    return has||row.india||row.us||row.cons;
  }).map(row=>{
    if(row.sec) return `<tr class="row-header"><td colspan="4">${row.sec}</td></tr>`;
    const subs=(row.grp&&G[row.grp])?G[row.grp]:[];
    const canOpen=subs.length>0;
    const id=`plg-${gi++}`;
    const caret=canOpen?`<span class="pl-caret">▸</span>`:'';
    const main=`<tr class="${row.cls||''}${canOpen?' pl-expandable':''}"${canOpen?` onclick="togglePLGroup('${id}',this)"`:''}>
      <td class="${row.indent?'pl-indent-1':''}">${caret}${row.label}${canOpen?`<span class="pl-sub-count">${subs.length}</span>`:''}</td>
      <td class="${nc(row.india)}">${fmtUSD(row.india)}</td>
      <td class="${nc(row.us)}">${fmtUSD(row.us)}</td>
      <td class="${nc(row.cons)} highlight-col">${fmtUSD(row.cons)}</td>
    </tr>`;
    if(!canOpen) return main;
    const kids=subs.map(sb=>`<tr class="pl-sub-row" data-plg="${id}" style="display:none">
      <td class="pl-indent-2">${sb.label}</td>
      <td class="${nc(sb.india)}">${fmtUSD(sb.india)}</td>
      <td class="${nc(sb.us)}">${fmtUSD(sb.us)}</td>
      <td class="${nc(sb.cons)} highlight-col">${fmtUSD(sb.cons)}</td>
    </tr>`).join('');
    return main+kids;
  }).join('');

  /* Note when a month has no expandable detail (January) */
  const dn=document.getElementById('pl-detail-note');
  if(dn){
    const any=Object.keys(G).some(k=>(G[k]||[]).length);
    dn.textContent=any
      ? 'Click any cost line to expand its constituent items.'
      : `The ${monthLabel} pack carries summary lines only — no sub-line detail available for this month.`;
  }

  /* Summary chips */
  const chips=document.getElementById('pl-chips');
  if(chips){
    chips.innerHTML=[
      {l:'Revenue',   v:s.revenue.cons,    p:null},
      {l:'Gross Profit',v:s.grossProfit.cons,p:pct(s.grossProfit.cons,s.revenue.cons||1)},
      {l:'EBITDA',    v:s.ebitda.cons,     p:pct(s.ebitda.cons,s.revenue.cons||1)}
    ].map(c=>{
      const pc=c.p!==null?(c.p>=0?'chip-positive':'chip-negative'):'';
      return `<div class="pl-chip"><span class="pl-chip-label">${c.l}</span><span class="pl-chip-value">${fmtUSD(c.v)}</span>${c.p!==null?`<span class="pl-chip-pct ${pc}">${fmtPct(c.p)}</span>`:''}</div>`;
    }).join('');
  }

  /* Charts */
  const cats=['Revenue','COGS','Gross Profit','SG&A','EBITDA'];
  const indV=[s.revenue.india,Math.abs(s.cogs.india||0),s.grossProfit.india,Math.abs(s.totalSGA.india||0),s.ebitda.india];
  const usV =[s.revenue.us,   Math.abs(s.cogs.us||0),   s.grossProfit.us,   Math.abs(s.totalSGA.us||0),   s.ebitda.us];
  const coV =[s.revenue.cons, Math.abs(s.cogs.cons||0), s.grossProfit.cons, Math.abs(s.totalSGA.cons||0), s.ebitda.cons];

  makeChart('chart-pl-bar',{type:'bar',data:{labels:cats,datasets:[
    {label:'India',       data:indV,backgroundColor:alpha(PAL.blue,0.7),  borderRadius:3},
    {label:'US',          data:usV, backgroundColor:alpha(PAL.purple,0.7),borderRadius:3},
    {label:'Consolidated',data:coV, backgroundColor:alpha(PAL.teal,0.7),  borderRadius:3}
  ]},options:lineOpts({scales:{y:{ticks:{callback:v=>fmtUSD(v,true)}}}})});

  const pcats=CFG.expenseCategories;
  makeChart('chart-pl-pie',{type:'doughnut',data:{labels:pcats.map(c=>c.label),datasets:[{data:pcats.map(c=>Math.abs((s[c.key]&&s[c.key].cons)||0)),backgroundColor:pcats.map(c=>c.color),borderWidth:0,hoverOffset:5}]},options:{responsive:true,maintainAspectRatio:false,cutout:'55%',plugins:{legend:{position:'right',labels:{font:{size:10},padding:6,boxWidth:10}},tooltip:{callbacks:{label:c=>` ${c.label}: ${fmtUSD(c.raw)}`}}}}});
}

/* ════════════════════════════════════════
   PAGE 3 — EXPENSE ANALYSIS
════════════════════════════════════════ */
function renderExpenseAnalysis(monthLabel, entity) {
  const mis=DATA.mis, months=mis.months;
  const mIdx=months.indexOf(monthLabel);
  const si=mIdx>=0?mIdx:months.length-1;
  const ent=(entity||'consolidated').toLowerCase();
  const cats=CFG.expenseCategories;

  /* Entity view. Consolidated reads the MIS Actuals sheet, which is the
     only place carrying full line detail for all six months. India / US
     read that month's own pack; January's pack has no sub-lines. */
  const mp=DATA.monthPL&&DATA.monthPL[months[si]];
  const entVal=(key)=>{
    if(ent==='consolidated') return Math.abs((mis[key]||[])[si]||0);
    const s=mp&&mp.summary&&mp.summary[key];
    return s?Math.abs(s[ent==='india'?'india':'us']||0):0;
  };
  const entSubs=(key)=>{
    if(ent==='consolidated'){
      const g=(mis.groups&&mis.groups[key])||[];
      return g.map(s=>({label:s.label,val:Math.abs(s.values[si]||0)})).filter(s=>s.val>0);
    }
    const g=(mp&&mp.groups&&mp.groups[key])||[];
    return g.map(s=>({label:s.label,val:Math.abs(s[ent==='india'?'india':'us']||0)})).filter(s=>s.val>0);
  };

  const vals=cats.map(c=>entVal(c.key));
  const totalExp=vals.reduce((a,b)=>a+b,0)||1;

  /* Tie-out — categories must reconcile to Total COGS + Overall SG&A */
  const tie=document.getElementById('exp-tieout');
  if(tie){
    if(ent==='consolidated'){
      const cogsCat=cats.reduce((a,c,i)=>a+(c.grp==='COGS'?vals[i]:0),0);
      const sgaCat =cats.reduce((a,c,i)=>a+(c.grp==='SG&A'?vals[i]:0),0);
      const cogsSh=Math.abs(mis.cogs[si]||0), sgaSh=Math.abs(mis.totalSGA[si]||0);
      const d1=cogsCat-cogsSh, d2=sgaCat-sgaSh;
      const ok=Math.abs(d1)<2&&Math.abs(d2)<2;
      tie.innerHTML=`<span class="${ok?'favorable-text':'unfavorable-text'}">${ok?'✓':'⚠'}</span> `+
        `COGS categories ${fmtUSD(cogsCat)} vs Total COGS ${fmtUSD(cogsSh)}`+
        (Math.abs(d1)>=2?` <b>(${fmtUSD(d1)} out)</b>`:'')+
        ` · SG&A categories ${fmtUSD(sgaCat)} vs Overall SG&A ${fmtUSD(sgaSh)}`+
        (Math.abs(d2)>=2?` <b>(${fmtUSD(d2)} out)</b>`:'');
    } else {
      tie.textContent=`${ent==='india'?'India':'US'} entity view — sourced from the ${months[si]} pack.`;
    }
  }

  /* Expense Mix donut */
  makeChart('chart-exp-mix',{type:'doughnut',data:{labels:cats.map(c=>c.label),datasets:[{data:vals,backgroundColor:cats.map(c=>c.color),borderWidth:0,hoverOffset:6}]},options:{responsive:true,maintainAspectRatio:false,cutout:'58%',plugins:{legend:{position:'right',labels:{font:{size:10},padding:6,boxWidth:10}},tooltip:{callbacks:{label:c=>` ${c.label}: ${fmtUSD(c.raw)} (${pct(c.raw,totalExp).toFixed(1)}%)`}}}}});

  /* Category drill-down cards */
  const container=document.getElementById('expense-categories');
  if(!container)return;

  container.innerHTML=cats.map((cat,ci)=>{
    const val=vals[ci];
    const share=pct(val,totalExp);
    const subs=entSubs(cat.key);
    const subHTML=subs.length
      ?`<table class="exp-sub-table"><thead><tr><th>Line Item</th><th class="num">Amount</th><th>% Share</th><th class="pct-bar-cell"></th></tr></thead><tbody>${subs.map(s=>`<tr><td>${s.label}</td><td class="num">${fmtUSD(s.val)}</td><td class="num">${pct(s.val,val||1).toFixed(1)}%</td><td><div class="mini-pct-bar"><div class="mini-pct-fill" style="width:${Math.min(100,pct(s.val,val||1))}%;background:${cat.color}"></div></div></td></tr>`).join('')}</tbody></table>`
      :`<p style="font-size:12px;color:var(--text-muted);padding:8px 0">No sub-line detail in the ${months[si]} pack for this group.</p>`;
    const subTot=subs.reduce((a,s)=>a+s.val,0);
    const gap=val-subTot;
    const gapNote=(subs.length&&Math.abs(gap)>=2)
      ?`<p style="font-size:11px;color:var(--text-muted);padding:4px 0">Sub-lines total ${fmtUSD(subTot)} · ${fmtUSD(gap)} sits on the group line itself.</p>`:'';
    return `<div class="exp-cat-card">
      <div class="exp-cat-header" onclick="toggleExpCat(this)">
        <div class="exp-cat-dot" style="background:${cat.color}"></div>
        <span class="exp-cat-name">${cat.label}<span class="exp-cat-grp">${cat.grp}</span></span>
        <span class="exp-cat-amount" style="color:${cat.color}">${fmtUSD(val)}</span>
        <span class="exp-cat-pct">${share.toFixed(1)}%</span>
        <span class="exp-cat-chevron">▼</span>
      </div>
      <div class="exp-cat-body">
        <div class="exp-cat-charts">
          <div class="exp-cat-chart-wrap" style="height:170px"><canvas id="exp-pie-${ci}"></canvas></div>
          <div class="exp-cat-chart-wrap" style="height:170px"><canvas id="exp-bar-${ci}"></canvas></div>
        </div>
        ${subHTML}${gapNote}
      </div>
    </div>`;
  }).join('');

  cats.forEach((cat,ci)=>{
    const subs=entSubs(cat.key);
    const val=vals[ci];
    if(subs.length){
      makeChart(`exp-pie-${ci}`,{type:'doughnut',data:{labels:subs.map(s=>s.label),datasets:[{data:subs.map(s=>s.val),backgroundColor:PAL.list,borderWidth:0,hoverOffset:4}]},options:{responsive:true,maintainAspectRatio:false,cutout:'55%',plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>` ${c.label}: ${fmtUSD(c.raw)}`}}}}});
    } else {
      makeChart(`exp-pie-${ci}`,{type:'doughnut',data:{labels:[cat.label,'Other'],datasets:[{data:[val,Math.max(0,totalExp-val)],backgroundColor:[cat.color,alpha(cat.color,0.15)],borderWidth:0}]},options:{responsive:true,maintainAspectRatio:false,cutout:'60%',plugins:{legend:{display:false}}}});
    }
    makeChart(`exp-bar-${ci}`,{type:'bar',data:{labels:months,datasets:[{label:cat.label,data:(mis[cat.key]||[]).map(v=>Math.abs(v||0)),backgroundColor:months.map((_,i)=>alpha(cat.color,i===si?0.95:0.45)),borderRadius:3}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{ticks:{font:{size:9}}},y:{ticks:{callback:v=>fmtUSD(v,true),font:{size:9}}}}}});
  });
}

window.toggleExpCat=function(h){
  h.classList.toggle('open');
  h.nextElementSibling.classList.toggle('open');
};

/* Expand / collapse a P&L cost group in place */
window.togglePLGroup=function(id,tr){
  const open=tr.classList.toggle('open');
  document.querySelectorAll(`tr.pl-sub-row[data-plg="${id}"]`)
    .forEach(r=>{ r.style.display = open ? '' : 'none'; });
};

function getSubItems(plData,catKey){
  if(!plData||!plData.groups)return[];
  return (plData.groups[catKey]||[])
    .map(r=>({label:r.label,val:Math.abs(r.cons||r.india||r.us||0)}))
    .filter(r=>r.val>0);
}

/* ════════════════════════════════════════
   PAGE 4 — BUDGET vs ACTUAL
════════════════════════════════════════ */
function renderBudgetVsActual(monthLabel) {
  const mis=DATA.mis, pf=DATA.proforma||{};
  const tbody=document.getElementById('bva-tbody');
  const b=pf[monthLabel];
  const aI=mis.months.indexOf(monthLabel);

  const pEl=document.getElementById('bva-table-period');
  if(pEl)pEl.textContent=monthLabel||'—';

  if(!b||aI===-1){
    const have=Object.keys(pf).join(', ')||'none';
    tbody.innerHTML=`<tr><td colspan="7" class="table-placeholder">Budget vs Actual is published for ${have} only.</td></tr>`;
    ['bva-kpi-row'].forEach(id=>{const e=document.getElementById(id); if(e)e.innerHTML='';});
    return;
  }

  const A=k=>(mis[k]||[])[aI]||0;
  const actManpower=A('manpower')+A('empStat');   /* budget carries the two combined */

  /* fav(v) — is a variance of v favourable?  v = actual − budget */
  const up  =v=>v>=0;   /* more is better  — revenue, profit  */
  const down=v=>v<=0;   /* less is better  — every cost line  */

  const rows=[
    {sec:'REVENUE'},
    {key:'revenue',    label:'Total Revenue',        bud:b.revenue,   act:A('revenue'),    fav:up,   cls:'row-subtotal'},
    {sec:'COST OF GOODS SOLD'},
    {key:'manpower',   label:'Manpower (incl. statutory)', bud:b.manpower, act:actManpower, fav:down, indent:true},
    {key:'facility',   label:'Facility Cost',        bud:b.facility,  act:A('facility'),   fav:down, indent:true},
    {key:'telecom',    label:'Telcom/Data & others', bud:b.telecom,   act:A('telecom'),    fav:down, indent:true},
    {key:'transport',  label:'Transport',            bud:b.transport, act:A('transport'),  fav:down, indent:true},
    {key:'otherCogs',  label:'Other COGS',           bud:b.otherCogs, act:A('otherCogs'),  fav:down, indent:true},
    {key:'totalCogs',  label:'Total COGS',           bud:b.totalCogs, act:A('cogs'),       fav:down, cls:'row-subtotal'},
    {sec:'PROFITABILITY'},
    {key:'gm',         label:'Gross Profit',         bud:b.gm,        act:A('grossProfit'),fav:up,   cls:'row-total'},
    {key:'sellingExp', label:'Selling Expenses',     bud:b.sellingExp,act:A('sellingExp'), fav:down, indent:true},
    {key:'mgmtSal',    label:'Salaries-Management',  bud:b.mgmtSal,   act:A('mgmtSal'),    fav:down, indent:true},
    {key:'legalProf',  label:'Legal & Professional', bud:b.legalProf, act:A('legalProf'),  fav:down, indent:true},
    {key:'gaSal',      label:'G&A Salaries',         bud:b.gaSal,     act:A('gaSal'),      fav:down, indent:true},
    {key:'adminExp',   label:'Gen & Admin Expenses', bud:b.adminExp,  act:A('adminExp'),   fav:down, indent:true},
    {key:'totalSGA',   label:'Overall SG&A',         bud:b.totalSGA,  act:A('totalSGA'),   fav:down, cls:'row-subtotal'},
    {key:'ebitda',     label:'EBITDA',               bud:b.ebitda,    act:A('ebitda'),     fav:up,   cls:'row-grand'}
  ];

  const srcNote=document.getElementById('bva-source-note');
  if(srcNote){
    srcNote.innerHTML=`Budget: <b>${b._sheet}</b> (forecast column) · Actual: <b>${CFG.sheets.mis.consolidated}</b>. `+
      `Two budget lines are reclassified to match the actuals' treatment — `+
      `Software Expenses ${fmtUSD(b._reclass.software)} moved from Telcom/Data into Other COGS, `+
      `Interest on OD ${fmtUSD(b._reclass.interestOD)} moved from Other COGS into Finance Cost.`;
  }

  /* KPI summary cards */
  const kpiRow=document.getElementById('bva-kpi-row');
  if(kpiRow){
    const rV=A('revenue')-b.revenue, eV=A('ebitda')-b.ebitda;
    const kpis=[
      {l:'Budget Revenue', t:fmtUSD(b.revenue),  v:null},
      {l:'Actual Revenue', t:fmtUSD(A('revenue')),v:null},
      {l:'Revenue Variance',t:fmtUSD(rV),        v:rV},
      {l:'EBITDA Variance', t:fmtUSD(eV),        v:eV}
    ];
    kpiRow.innerHTML=kpis.map(k=>{
      const cls=k.v===null?'':(k.v>=0?'favorable-text':'unfavorable-text');
      return `<div class="bva-kpi-card"><div class="bva-kpi-label">${k.l}</div><div class="bva-kpi-actual ${cls}">${k.t}</div></div>`;
    }).join('');
  }

  /* Variance table */
  tbody.innerHTML=rows.map(row=>{
    if(row.sec) return `<tr class="row-header"><td colspan="7">${row.sec}</td></tr>`;
    const varV=row.act-row.bud;
    const varPct=row.bud?((varV/Math.abs(row.bud))*100):0;
    const fav=row.fav(varV);
    const within=row.bud?Math.abs(varPct)<=5:false;
    const vc=within?'neutral-var':fav?'favorable':'unfavorable';
    const tl=within?'🟡':fav?'🟢':'🔴';
    return `<tr class="${row.cls||''}">
      <td class="${row.indent?'pl-indent-1':''}">${row.label}</td>
      <td class="col-num">${fmtUSD(row.bud)}</td>
      <td class="col-num">${fmtUSD(row.act)}</td>
      <td class="col-num ${vc}">${fmtUSD(varV)}</td>
      <td class="col-num ${vc}">${row.bud?fmtPct(varPct):'—'}</td>
      <td class="col-status">${tl}</td>
      <td><button class="drill-btn" onclick="openDrilldown('${row.key}','${monthLabel}')" title="Drill down">⊕</button></td>
    </tr>`;
  }).join('');

  /* Grouped bar */
  const mRows=rows.filter(r=>!r.sec&&['revenue','totalCogs','gm','totalSGA','ebitda'].includes(r.key));
  makeChart('chart-bva-grouped',{type:'bar',data:{labels:mRows.map(r=>r.label),datasets:[
    {label:'Budget',data:mRows.map(r=>Math.abs(r.bud||0)),backgroundColor:alpha(PAL.blue,0.6),borderRadius:4},
    {label:'Actual',data:mRows.map(r=>Math.abs(r.act||0)),backgroundColor:alpha(PAL.green,0.7),borderRadius:4}
  ]},options:lineOpts({scales:{y:{ticks:{callback:v=>fmtUSD(v,true)}}}})});

  /* Variance waterfall */
  const vRows=rows.filter(r=>!r.sec);
  makeChart('chart-bva-waterfall',{type:'bar',data:{labels:vRows.map(r=>r.label),datasets:[{label:'Variance ($)',data:vRows.map(r=>r.act-r.bud),backgroundColor:vRows.map(r=>alpha(r.fav(r.act-r.bud)?PAL.green:PAL.red,0.75)),borderRadius:4}]},options:lineOpts({plugins:{legend:{display:false}},scales:{x:{ticks:{font:{size:9},maxRotation:60,minRotation:30}},y:{ticks:{callback:v=>fmtUSD(v,true)}}}})});

  /* Variance trend — only the months that have a Proforma pack */
  const pm=Object.keys(pf).filter(m=>mis.months.indexOf(m)>=0);
  const revV=pm.map(m=>((mis.revenue[mis.months.indexOf(m)]||0)-pf[m].revenue));
  const ebV =pm.map(m=>((mis.ebitda [mis.months.indexOf(m)]||0)-pf[m].ebitda));
  makeChart('chart-variance-trend',{type:'bar',data:{labels:pm,datasets:[
    {label:'Revenue Variance',data:revV,backgroundColor:revV.map(v=>alpha(v>=0?PAL.blue:PAL.red,0.65)),borderRadius:4},
    {label:'EBITDA Variance', data:ebV,type:'line',borderColor:PAL.orange,borderWidth:2.5,fill:false,tension:0.4,pointRadius:5}
  ]},options:lineOpts({scales:{y:{ticks:{callback:v=>fmtUSD(v,true)}}}})});
}

/* ════════════════════════════════════════
   DRILL-DOWN MODAL
════════════════════════════════════════ */
window.openDrilldown=function(key,monthLabel){
  const mis=DATA.mis;
  const pf=(DATA.proforma||{})[monthLabel]||null;
  const aI=mis.months.indexOf(monthLabel);
  const pl=DATA.monthPL?.[monthLabel];
  const labels={revenue:'Revenue',cogs:'Total COGS',gm:'Gross Profit',totalCogs:'Total COGS',ebitda:'EBITDA',pbt:'Net Profit',salaries:'Salaries',facility:'Facilities',telecom:'Telecom',transport:'Transport',otherCogs:'Other COGS',totalSGA:'Total SG&A',financeCost:'Finance Charges',grossProfit:'Gross Profit',manpower:'Employee Cost',sellingExp:'Selling Expenses',mgmtSal:'Mgmt Salaries',legalProf:'Legal & Prof',adminExp:'Admin Expenses'};
  const modal=document.getElementById('drilldown-modal');
  const overlay=document.getElementById('drilldown-overlay');
  document.getElementById('drilldown-title').textContent=`${labels[key]||key} — ${monthLabel}`;

  /* Map budget keys → MIS keys (names differ between the two packs) */
  const misKeyMap={salaries:'manpower',facilities:'facility',totalCogs:'cogs',gm:'grossProfit'};
  const misKey=misKeyMap[key]||key;
  const misSeries=mis[misKey]||mis['revenue']||[];
  let actVal=aI>=0?(misSeries[aI]||0):0;
  if(key==='manpower'&&aI>=0) actVal=(mis.manpower[aI]||0)+(mis.empStat[aI]||0);
  const budVal=pf?(pf[key]||0):0;
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
    {label:'Budget',data:mis.months.map(m=>{const p=(DATA.proforma||{})[m];return p?Math.abs(p[key]||0):0;}),backgroundColor:alpha(PAL.blue,0.4),borderRadius:3},
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
    const sumData=[['Particulars',...mis.months],['Revenue',...mis.revenue],['Gross Profit',...mis.grossProfit],['EBITDA',...mis.ebitda],['COGS',...mis.cogs],['Manpower',...mis.manpower],['SG&A',...mis.totalSGA]];
    XLSX.utils.book_append_sheet(wb2,XLSX.utils.aoa_to_sheet(sumData),'Summary');
    /* Budget vs Actual export mirrors the page exactly — Proforma budget
       against MIS actuals, for the Proforma months only. */
    const pf=DATA.proforma||{};
    if(Object.keys(pf).length){
      const bvaData=[['Month','Metric','Budget','Actual','Variance','Var%']];
      const LINES=[['Total Revenue','revenue','revenue'],['Manpower (incl. statutory)','manpower','manpower'],
        ['Facility Cost','facility','facility'],['Telcom/Data & others','telecom','telecom'],
        ['Transport','transport','transport'],['Other COGS','otherCogs','otherCogs'],
        ['Total COGS','totalCogs','cogs'],['Gross Profit','gm','grossProfit'],
        ['Selling Expenses','sellingExp','sellingExp'],['Salaries-Management','mgmtSal','mgmtSal'],
        ['Legal & Professional','legalProf','legalProf'],['G&A Salaries','gaSal','gaSal'],
        ['Gen & Admin Expenses','adminExp','adminExp'],['Overall SG&A','totalSGA','totalSGA'],
        ['EBITDA','ebitda','ebitda']];
      Object.keys(pf).forEach(m=>{
        const aI=mis.months.indexOf(m); if(aI<0) return;
        LINES.forEach(([lbl,bk,ak])=>{
          const b=pf[m][bk]||0;
          const a=(ak==='manpower')?((mis.manpower[aI]||0)+(mis.empStat[aI]||0)):((mis[ak]||[])[aI]||0);
          bvaData.push([m,lbl,b,a,a-b,b?((a-b)/Math.abs(b)*100).toFixed(1)+'%':'—']);
        });
      });
      XLSX.utils.book_append_sheet(wb2,XLSX.utils.aoa_to_sheet(bvaData),'Budget vs Actual');
    }
    XLSX.writeFile(wb2,'iMagnum_FPA_Dashboard.xlsx');
  });
}

/* ════════════════════════════════════════
   PAGE 1b — KEY METRICS (all months + YTD)
════════════════════════════════════════ */
function renderKeyMetrics(){
  const mis=DATA.mis, months=mis.months, y=mis.ytd||{};
  const host=document.getElementById('key-metrics-body');
  if(!host||!months.length) return;
  const pctRow=(num,den)=>months.map((_,i)=>den[i]?pct(num[i],den[i]):0);
  const ytdPct=(a,b)=>y[b]?pct(y[a],y[b]):0;

  const rows=[
    {label:'Total Revenue',   s:mis.revenue,     ytd:y.revenue,     strong:true},
    {label:'Total COGS',      s:mis.cogs,        ytd:y.cogs},
    {label:'Gross Profit',    s:mis.grossProfit, ytd:y.grossProfit, strong:true},
    {label:'Gross Margin %',  s:pctRow(mis.grossProfit,mis.revenue), ytd:ytdPct('grossProfit','revenue'), isPct:true},
    {label:'Overall SG&A',    s:mis.totalSGA,    ytd:y.totalSGA},
    {label:'EBITDA',          s:mis.ebitda,      ytd:y.ebitda,      strong:true},
    {label:'EBITDA Margin %', s:pctRow(mis.ebitda,mis.revenue), ytd:ytdPct('ebitda','revenue'), isPct:true},

    {label:'Manpower',            s:mis.manpower,   ytd:y.manpower},
    {label:'Employee Statutory',  s:mis.empStat,    ytd:y.empStat},
    {label:'Facility Cost',       s:mis.facility,   ytd:y.facility},
    {label:'Telcom/Data & others',s:mis.telecom,    ytd:y.telecom},
    {label:'Transport',           s:mis.transport,  ytd:y.transport},
    {label:'Other COGS',          s:mis.otherCogs,  ytd:y.otherCogs},
    {label:'Selling Expenses',    s:mis.sellingExp, ytd:y.sellingExp},
    {label:'Salaries-Management', s:mis.mgmtSal,    ytd:y.mgmtSal},
    {label:'Legal & Professional',s:mis.legalProf,  ytd:y.legalProf},
    {label:'G&A Salaries',        s:mis.gaSal,      ytd:y.gaSal},
    {label:'Gen & Admin Expenses',s:mis.adminExp,   ytd:y.adminExp}
  ];
  const cell=(v,isPct)=>isPct?fmtPct(v):fmtUSD(v);
  host.innerHTML=`
    <table class="fin-table km-table">
      <thead><tr>
        <th>Particulars</th>
        ${months.map(m=>`<th class="col-num">${m}</th>`).join('')}
        <th class="col-num km-ytd">${mis.ytdLabel||'YTD'}</th>
      </tr></thead>
      <tbody>
        ${rows.map(r=>`<tr class="${r.strong?'row-strong':''}">
          <td>${r.label}</td>
          ${r.s.map(v=>`<td class="col-num ${(v||0)<0?'unfavorable-text':''}">${cell(v,r.isPct)}</td>`).join('')}
          <td class="col-num km-ytd ${(r.ytd||0)<0?'unfavorable-text':''}">${cell(r.ytd,r.isPct)}</td>
        </tr>`).join('')}
      </tbody>
    </table>`;
}

/* ════════════════════════════════════════
   PAGE — FULL-YEAR P&L
   Renders the consolidated P&L exactly as it is laid out on the MIS
   Actuals sheet: every month side by side, plus a YTD column. Row type is
   inferred from the sheet itself — indented labels are cost groups, the
   known total/section labels are totals, everything else between them is a
   constituent line — so added or renamed line items flow through without
   a code change. Stops at EBITDA.
════════════════════════════════════════ */
const PL_SECTIONS=new Set(['service income','overall operational cost','sg&a','g&a expenses']);

function extractFullPL(){
  const rows=sheetToArr(WB.mis,CFG.sheets.mis.consolidated);
  const hIdx=findRowIdx(rows,['Particulars'],0);
  if(hIdx<0) return null;
  const hdr=rows[hIdx]||[];
  const months=[];
  for(let c=1;c<hdr.length;c++){
    const v=String(hdr[c]||'').trim();
    if(/\w+-\d{2}/.test(v)) months.push({label:v,col:c});
  }
  if(!months.length) return null;

  const stop=findRowIdx(rows,['EBITDA'],0);          /* render down to EBITDA only */
  const last=stop>=0?stop:rows.length-1;

  const out=[]; let inGroup=false; let gid=0; let curGroup=null;
  for(let i=hIdx+1;i<=last;i++){
    const raw=rows[i][0];
    const lab=cleanLabel(raw);
    if(!lab) continue;
    if(lab.endsWith('%')) continue;                   /* ratios recomputed below */
    const vals=months.map(m=>parseNum(rows[i][m.col]));
    const ytd=vals.reduce((x,y)=>x+y,0);
    const low=lab.toLowerCase();

    let type;
    if(PL_SECTIONS.has(low))      { type='section'; inGroup=false; curGroup=null; }
    else if(isStopRow(raw))       { type='total';   inGroup=false; curGroup=null; }
    else if(isIndented(raw))      { type='group';   inGroup=true;  curGroup='fpg'+(gid++); }
    else if(inGroup)              { type='sub'; }
    else                          { type='sub'; }

    if(type==='sub'&&vals.every(v=>v===0)) continue;  /* drop dormant line items */
    out.push({label:lab,type,vals,ytd,group:type==='sub'?curGroup:(type==='group'?curGroup:null)});
  }

  /* Reconcile each group to its itemised lines. On this workbook the group
     headers for January–March are pasted totals that do not foot to the line
     items beneath them (April onward do), so rather than display itemisation
     that silently fails to add up, the difference is surfaced as an explicit
     residual line. Group totals stay the source of truth — they tie to
     Total COGS. */
  const withKids=new Set(out.filter(r=>r.type==='sub'&&r.group).map(r=>r.group));
  out.forEach(r=>{ if(r.type==='group') r.hasKids=withKids.has(r.group); });

  const nMon=months.length;
  for(let i=out.length-1;i>=0;i--){
    const g=out[i];
    if(g.type!=='group'||!g.hasKids) continue;
    const kids=out.filter(r=>r.type==='sub'&&r.group===g.group);
    const resid=[]; let material=false;
    for(let m=0;m<nMon;m++){
      const d=g.vals[m]-kids.reduce((a,k)=>a+k.vals[m],0);
      resid.push(d);
      if(Math.abs(d)>2) material=true;
    }
    if(!material) continue;
    const lastKid=out.lastIndexOf(kids[kids.length-1]);
    out.splice(lastKid+1,0,{
      label:'Not itemised on source sheet',
      type:'sub', residual:true, group:g.group,
      vals:resid, ytd:resid.reduce((a,b)=>a+b,0)
    });
    g.hasResidual=true;
  }

  /* computed margin rows, inserted after GM and after EBITDA */
  const find=n=>out.find(r=>r.label.toLowerCase()===n);
  const rev=find('total revenue'), gm=find('gm'), eb=find('ebitda');
  const ratio=(num,den)=>({
    label:'', type:'ratio',
    vals:num.vals.map((v,i)=>den.vals[i]?v/den.vals[i]*100:0),
    ytd:den.ytd?num.ytd/den.ytd*100:0
  });
  if(rev&&gm){ const r=ratio(gm,rev); r.label='GM %';     out.splice(out.indexOf(gm)+1,0,r); }
  if(rev&&eb){ const r=ratio(eb,rev); r.label='EBITDA %'; out.splice(out.indexOf(eb)+1,0,r); }

  return {months:months.map(m=>m.label),rows:out};
}

function renderFullPL(mode){
  const D=DATA.fullPL;
  const head=document.getElementById('fullpl-head');
  const body=document.getElementById('fullpl-tbody');
  if(!head||!body) return;
  if(!D){ body.innerHTML='<tr><td class="table-placeholder">P&L could not be read from the workbook.</td></tr>'; return; }

  const showAll = (mode==='full');
  head.innerHTML='<th>Particulars</th>'+
    D.months.map(m=>`<th class="col-num">${m}</th>`).join('')+
    '<th class="col-num highlight-col">YTD</th>';

  const rng=document.getElementById('fullpl-range');
  if(rng) rng.textContent=`${D.months[0]} – ${D.months[D.months.length-1]} · YTD summed from ${D.months.length} months`;
  const note=document.getElementById('fullpl-note');
  const gaps=D.rows.filter(r=>r.type==='group'&&r.hasResidual).map(r=>r.label);
  if(note) note.innerHTML=
    (showAll?'All line items shown. Switch to Summary to collapse cost groups. '
            :'Click any cost group to expand its constituent line items. ')+
    (gaps.length?`<b>Note:</b> on the source sheet the group totals for `+
       `${gaps.join(', ')} exceed the sum of their itemised lines in the earlier months. `+
       `The difference is shown as \u201cNot itemised on source sheet\u201d so each group still foots.`
     :'');

  const num=v=>{
    const cls=v<0?'col-num num-val negative':v>0?'col-num num-val':'col-num num-val text-muted';
    return {cls,txt:fmtUSD(v)};
  };
  body.innerHTML=D.rows.map(r=>{
    if(r.type==='section')
      return `<tr class="row-header"><td colspan="${D.months.length+2}">${r.label}</td></tr>`;

    if(r.type==='ratio'){
      const cells=r.vals.map(v=>`<td class="col-num num-val${v<0?' negative':''}">${v.toFixed(1)}%</td>`).join('');
      return `<tr class="row-ratio"><td class="pl-indent-1">${r.label}</td>${cells}`+
             `<td class="col-num num-val highlight-col${r.ytd<0?' negative':''}">${r.ytd.toFixed(1)}%</td></tr>`;
    }

    const cells=r.vals.map(v=>{const n=num(v);return `<td class="${n.cls}">${n.txt}</td>`;}).join('');
    const y=num(r.ytd);
    const ytdCell=`<td class="${y.cls} highlight-col">${y.txt}</td>`;

    if(r.type==='total'){
      const grand=/^(ebitda|gm|total revenue)$/i.test(r.label);
      return `<tr class="${grand?'row-grand':'row-subtotal'}"><td>${r.label}</td>${cells}${ytdCell}</tr>`;
    }
    if(r.type==='group'){
      const open=r.hasKids&&showAll;
      const caret=r.hasKids?`<span class="pl-caret">▸</span>`:'';
      return `<tr class="${r.hasKids?'pl-expandable':''}${open?' open':''}"`+
             `${r.hasKids?` onclick="toggleFullPLGroup('${r.group}',this)"`:''}>`+
             `<td class="pl-indent-1">${caret}${r.label}</td>${cells}${ytdCell}</tr>`;
    }
    /* sub-line */
    return `<tr class="pl-sub-row${r.residual?' pl-residual':''}" data-fpg="${r.group||''}" `+
           `style="display:${showAll?'':'none'}"${r.residual?' title="Group total on the source sheet exceeds the sum of its itemised lines for these months."':''}>`+
           `<td class="pl-indent-2">${r.label}</td>${cells}${ytdCell}</tr>`;
  }).join('');
}

window.toggleFullPLGroup=function(gid,tr){
  const open=tr.classList.toggle('open');
  document.querySelectorAll(`tr.pl-sub-row[data-fpg="${gid}"]`)
    .forEach(x=>{ x.style.display = open ? '' : 'none'; });
};

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

  /* Executive — every month, plus a cumulative option */
  const execSel=document.getElementById('exec-period-select');
  if(execSel){
    execSel.innerHTML=
      months.map(m=>`<option value="${m}"${m===latest?' selected':''}>${m}</option>`).join('')+
      `<option value="ytd">${DATA.mis.ytdLabel||'YTD (All Months)'}</option>`;
    execSel.addEventListener('change',e=>renderExecutive(e.target.value));
  }

  ['pl-month-select','exp-month-select'].forEach(id=>{
    const el=document.getElementById(id);
    if(!el)return;
    el.innerHTML=months.map(m=>`<option value="${m}"${m===latest?' selected':''}>${m}</option>`).join('');
  });

  /* Budget vs Actual — only months with a Proforma pack */
  const bvaMonths=Object.keys(DATA.proforma||{}).filter(m=>months.indexOf(m)>=0);
  const bvaSel=document.getElementById('bva-month-select');
  if(bvaSel){
    const bvaLatest=bvaMonths[bvaMonths.length-1];
    bvaSel.innerHTML=bvaMonths.length
      ? bvaMonths.map(m=>`<option value="${m}"${m===bvaLatest?' selected':''}>${m}</option>`).join('')
      : '<option value="">No budget months</option>';
  }

  document.getElementById('pl-month-select')?.addEventListener('change',e=>renderPL(e.target.value));
  document.getElementById('fullpl-detail-select')?.addEventListener('change',e=>renderFullPL(e.target.value));
  document.getElementById('exp-month-select')?.addEventListener('change',e=>renderExpenseAnalysis(e.target.value,document.getElementById('exp-entity-select')?.value));
  document.getElementById('exp-entity-select')?.addEventListener('change',e=>renderExpenseAnalysis(document.getElementById('exp-month-select')?.value,e.target.value));
  document.getElementById('bva-month-select')?.addEventListener('change',e=>renderBudgetVsActual(e.target.value));
}

/* ════════════════════════════════════════
   DARK MODE
════════════════════════════════════════ */
function setupTheme(){
  const saved=localStorage.getItem('imag-theme')||'dark';
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
    WB.mis   =await loadWorkbook(CFG.files.mis, 'MIS Consolidated');
    WB.budget=(CFG.files.budget===CFG.files.mis)
      ? WB.mis
      : await loadWorkbook(CFG.files.budget, 'Budget');
    setStatus('Extracting financial data…');
    DATA.mis     =extractMIS();
    CFG.sheets.proforma=discoverProforma(WB.budget,DATA.mis.months);
    DATA.budget  =extractBudget();
    DATA.proforma=extractProforma();
    setStatus('Processing monthly P&L…');
    DATA.monthPL={};
    for(const m of DATA.mis.months) DATA.monthPL[m]=extractMonthPL(m);
    buildSearchIndex();
    setStatus('Rendering dashboard…');
    populateSelects();
    const latest=DATA.mis.months[DATA.mis.months.length-1];
    renderExecutive(latest);
    renderKeyMetrics();
    DATA.fullPL=extractFullPL();
    renderFullPL('summary');
    renderPL(latest);
    renderExpenseAnalysis(latest,'consolidated');
    const bvaFirst=document.getElementById('bva-month-select')?.value
                 || Object.keys(DATA.proforma||{})[0];
    renderBudgetVsActual(bvaFirst);
    setupNavigation();
    setupSearch();
    setupExports();
    setupTheme();
    /* Freshness + data health — surfaces a broken workbook rather than
       letting it render silent zeros */
    const ft=document.getElementById('freshness-text');
    if(ft){
      const missing=Object.keys(CFG.labels).filter(k=>{
        const s=DATA.mis[k];
        return Array.isArray(s)&&s.length&&s.every(v=>v===0);
      });
      const pfMissing=Object.keys(CFG.sheets.proforma).filter(m=>!DATA.proforma[m]);
      const bits=[`${DATA.mis.months.length} months: ${DATA.mis.months.join(', ')}`];
      if(missing.length)   bits.push(`⚠ unresolved: ${missing.join(', ')}`);
      if(DATA.mis.ytdStale) bits.push(`⚠ workbook Total column is stale (sums fewer months) — dashboard YTD is summed from the months instead`);
      if(pfMissing.length) bits.push(`⚠ no proforma: ${pfMissing.join(', ')}`);
      ft.textContent=bits.join(' · ');
      ft.title='Updated '+new Date().toLocaleTimeString('en-US',{hour:'2-digit',minute:'2-digit'});
      if(missing.length||pfMissing.length) console.warn('Data health:',{missing,pfMissing});
    }
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
