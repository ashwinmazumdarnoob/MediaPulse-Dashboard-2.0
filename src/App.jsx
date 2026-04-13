import { useState, useMemo, useCallback, useRef } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  PieChart, Pie, Cell, BarChart, Bar, XAxis, YAxis, Tooltip,
  LineChart, Line, CartesianGrid, Legend, ResponsiveContainer,
} from "recharts";
import {
  Upload, CheckCircle2, AlertTriangle, ChevronRight,
  Brain, FileText, Columns3, LayoutDashboard,
  TrendingUp, TrendingDown, Copy, Info, Eye,
  Table2, Zap, Database, FileBarChart
} from "lucide-react";

/* ═══ DESIGN TOKENS ═══ */
const C = {
  bg0:"#080c18",bg1:"#0f1525",bg2:"#161d30",bg3:"#1e2844",
  border:"#253050",p:"#7c5cfc",pL:"#a78bfa",pD:"#5b3fd6",pBg:"rgba(124,92,252,0.12)",
  g:"#34d399",r:"#f87171",a:"#fbbf24",b:"#60a5fa",cy:"#22d3ee",
  t0:"#f1f5f9",t1:"#94a3b8",t2:"#64748b",
  font:"'DM Sans',system-ui,sans-serif",mono:"'JetBrains Mono','Fira Code',monospace",
};
const COLORS=[C.p,C.g,C.b,C.a,C.r,C.cy,"#f472b6","#a3e635","#fb923c","#8b5cf6"];

/* ═══ ALIASES — maps metric names to possible CSV headers ═══ */
const ALIASES = {
  Date:["date","day","report_date","reporting_date"],
  Campaign:["campaign","campaign_name","insertion order","insertion_order","line item","ad group","adgroup","ad_group"],
  Platform:["platform","source","media_source","network","channel"],
  Spend:["spend","cost","amount_spent","budget","revenue (adv currency)","revenue (advertiser currency)","revenue","media cost","total cost","cost (advertiser currency)"],
  Impressions:["impressions","impr","imps","impression","total_impressions"],
  Clicks:["clicks","click","link_clicks","total_clicks"],
  Installs:["installs","install","conversions","total_attributions","attributed installs"],
  "Complete Views":["complete views (video)","complete views","video completions","completed views","trueview views","video views"],
  Creative:["creative","creative_name","youtube ad","ad_creative","ad name","ad"],
  Advertiser:["advertiser","advertiser_name","account","account_name"],
  Currency:["advertiser currency","currency"],
  Device:["device","device_type","os"],
  City:["city","geo","location","region"],
  Placement:["placement","position"],
  CTR:["ctr","click_through_rate"],CPC:["cpc","cost_per_click"],CPI:["cpi","cost_per_install"],
  CPM:["cpm","cost_per_mille"],ROAS:["roas"],CVR:["cvr","conversion_rate"],
  VTR:["vtr","view_through_rate","video_view_rate"],
  Sessions:["sessions","session"],Revenue:["in_app_revenue","total_revenue","app_revenue"],
  // AppsFlyer event columns
  "Add to Cart":["event_count_instamart_add_to_cart","event_count_im_add_to_cart","add_to_cart","atc"],
  "First Order":["event_count_instamart_first_order","event_count_im_first_order_purchase","first_order","first_purchase"],
  "Purchase":["event_count_instamart_purchase","event_count_im_purchase","purchase","purchases"],
};

// Only Date + Campaign + at least one metric are truly required
const REQ_COLS = ["Date","Campaign"];

/* ═══ AUTO MAP ═══ */
function autoMap(headers) {
  const m = {};
  // Pass 1: exact match
  for (const metric of Object.keys(ALIASES)) {
    for (const h of headers) {
      if (h.toLowerCase().trim() === metric.toLowerCase()) { m[metric]=h; break; }
    }
  }
  // Pass 2: alias match
  for (const [metric, aliasList] of Object.entries(ALIASES)) {
    if (m[metric]) continue;
    for (const h of headers) {
      if (aliasList.includes(h.toLowerCase().trim())) { m[metric]=h; break; }
    }
  }
  return m;
}

/* ═══ DV360 FOOTER DETECTION ═══ */
function isDV360FooterRow(row) {
  const vals = Object.values(row).map(v => String(v||"").trim());
  const first = vals[0]?.toLowerCase() || "";
  if (/^(date range|group by|filter by|mrc accredited|reporting numbers)/.test(first)) return true;
  if (first === "" && vals.filter(v=>v).length <= 1) return true;
  return false;
}

/* ═══ GEO EXTRACTION from DV360 campaign names ═══ */
function extractGeo(campaignName) {
  if (!campaignName) return "Other";
  const s = campaignName.toUpperCase();
  const geos = ["TAMILNADU","TAMIL NADU","ANDHRA PRADESH","TELANGANA","KERALA","HINDI-CITIES","ENGLISH-GEOS","KARNATAKA","MAHARASHTRA","MH & DNCR","DNCR","BENGALURU","MUMBAI","DELHI","KOLKATA","CHENNAI","HYDERABAD","PUNE","GOA","CHANDIGARH","AHMEDABAD","KOCHI"];
  for (const g of geos) { if (s.includes(g)) return g.charAt(0)+g.slice(1).toLowerCase(); }
  // Try extracting from AppsFlyer format: ..._LC_Mum-Del_NCR-Hyd...
  const lcMatch = s.match(/_LC_([^_]+)/);
  if (lcMatch) return lcMatch[1];
  return "Other";
}

function extractAdFormat(campaignName) {
  if (!campaignName) return "Other";
  const s = campaignName.toUpperCase();
  if (s.includes("INNS-CTV") || s.includes("INSTREAM")) return "In-Stream CTV";
  if (s.includes("VVC-CTV") || s.includes("SHORTS")) return "Shorts / VVC";
  if (s.includes("SKIP")) return "Skippable";
  if (s.includes("BUMPER")) return "Bumper";
  if (s.includes("DISPLAY")) return "Display";
  if (s.includes("VIDEO")) return "Video";
  if (s.includes("SOCIAL")) return "Social";
  return "Other";
}

/* ═══ DATE PARSING ═══ */
function parseDate(v) {
  if (!v) return "";
  let s = String(v).trim();
  if (v instanceof Date && !isNaN(v.getTime())) return v.toISOString().slice(0,10);
  // Already YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  // YYYY/MM/DD (DV360)
  let m = s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})/);
  if (m) return `${m[1]}-${m[2].padStart(2,"0")}-${m[3].padStart(2,"0")}`;
  // DD/MM/YYYY
  m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/);
  if (m) return `${m[3]}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}`;
  // Fallback
  const d = new Date(s);
  if (!isNaN(d.getTime()) && d.getFullYear()>2000) return d.toISOString().slice(0,10);
  return s;
}

/* ═══ NUMBER PARSING (handles commas, quotes, symbols) ═══ */
function parseNum(v) {
  if (v===null||v===undefined) return 0;
  if (typeof v === "number") return isNaN(v)?0:v;
  let s = String(v).trim().replace(/[₹$€£"']/g,"").replace(/,/g,"").replace(/%/g,"");
  if (s===""||s==="—"||s==="-"||s==="N/A"||s==="null") return 0;
  const n = parseFloat(s);
  return isNaN(n)?0:n;
}

/* ═══ VALIDATE ═══ */
function validate(rows, cm, dataType) {
  const errs=[], warns=[];
  if (!rows?.length) { errs.push("File is empty."); return {errs,warns,clean:[]}; }
  const hdr = Object.keys(rows[0]);

  // Check minimum columns
  for (const r of REQ_COLS) {
    if (!cm[r]||!hdr.includes(cm[r])) errs.push(`Required column "${r}" not mapped. Your headers: ${hdr.slice(0,8).join(", ")}${hdr.length>8?"...":""}`);
  }
  // For media data, need at least Spend or Impressions
  if (dataType==="media" && !cm.Spend && !cm.Impressions) {
    errs.push(`Need at least Spend or Impressions column for media data.`);
  }
  if (errs.length) return {errs,warns,clean:[]};

  const clean=[];
  let stripped=0, badDate=0, blank=0;
  const dateRe=/^\d{4}-\d{2}-\d{2}$/;

  for (let i=0;i<rows.length;i++) {
    const raw=rows[i];
    // Skip DV360 footer
    if (isDV360FooterRow(raw)) { stripped++; continue; }
    // Skip blank
    const vals=Object.values(raw).filter(v=>v!==null&&v!==undefined&&String(v).trim()!=="");
    if (!vals.length) { blank++; continue; }
    // Skip totals
    const campVal = String(raw[cm.Campaign]||"").trim();
    if (/^(total|sum|grand|subtotal|average)\b/i.test(campVal)) { stripped++; continue; }

    const row={_i:i+2};
    // Date
    const dv = parseDate(raw[cm.Date]);
    if (!dateRe.test(dv)) { badDate++; if(badDate<=2) warns.push(`Row ${i+2}: Date "${raw[cm.Date]}" could not be parsed.`); }
    row.Date=dv;
    row.Campaign=campVal;
    // Platform: optional — use column if available, else "DV360" or "Unknown"
    row.Platform = cm.Platform ? String(raw[cm.Platform]||"").trim() : "";
    if (!row.Platform) {
      // Try to infer from campaign name or advertiser
      const cLower = (row.Campaign + " " + String(raw[cm.Advertiser]||"")).toLowerCase();
      if (cLower.includes("dv360")) row.Platform="DV360";
      else if (cLower.includes("meta")||cLower.includes("facebook")||cLower.includes("fbig")) row.Platform="Meta";
      else if (cLower.includes("google")||cLower.includes("uac")) row.Platform="Google";
      else if (cLower.includes("programmatic")) row.Platform="Programmatic";
      else row.Platform=dataType==="appsflyer"?"AppsFlyer":"DV360";
    }

    // All numeric fields
    const numF=["Spend","Impressions","Clicks","Installs","Complete Views","CTR","CPC","CPI","CPM","ROAS","CVR","VTR","Sessions","Add to Cart","First Order","Purchase"];
    for (const f of numF) {
      if (cm[f]) row[f]=parseNum(raw[cm[f]]);
      else row[f]=0;
    }
    // String fields
    for (const f of ["Creative","Advertiser","Currency","Device","City","Placement"]) {
      if (cm[f]) row[f]=String(raw[cm[f]]||"").trim();
    }
    // Derived metrics
    row.Geo = extractGeo(row.Campaign);
    row.AdFormat = extractAdFormat(row.Campaign + " " + (row.Creative||""));
    // Calculated: VTR, CPM, CPV
    if (!row.VTR && row["Complete Views"] && row.Impressions) row.VTR = (row["Complete Views"]/row.Impressions*100);
    if (!row.CPM && row.Spend && row.Impressions) row.CPM = (row.Spend/row.Impressions*1000);
    if (row["Complete Views"] && row.Spend) row.CPV = row.Spend/row["Complete Views"];

    clean.push(row);
  }
  if (stripped) warns.push(`${stripped} metadata/footer/total rows auto-removed.`);
  if (blank) warns.push(`${blank} blank rows removed.`);
  if (badDate>2) warns.push(`${badDate} date format issues total.`);

  // Platform consistency
  const plats=[...new Set(clean.map(r=>r.Platform).filter(Boolean))];
  const lm={};
  plats.forEach(p=>{const k=p.toLowerCase();(lm[k]=lm[k]||[]).push(p);});
  Object.values(lm).filter(v=>v.length>1).forEach(v=>warns.push(`Inconsistent names: ${v.join(" vs ")}`));

  if (!clean.length) errs.push("No valid data rows after cleaning.");
  return {errs,warns,clean};
}

/* ═══ FILE PARSING ═══ */
function parseFile(file) {
  return new Promise((resolve,reject)=>{
    const ext=file.name.split(".").pop().toLowerCase();
    if (!["csv","tsv","xlsx","xls"].includes(ext)) { reject(new Error(`Unsupported file: .${ext}`)); return; }
    if (ext==="csv"||ext==="tsv") {
      Papa.parse(file,{header:true,skipEmptyLines:true,dynamicTyping:false,
        complete:r=>{
          if(!r.data?.length){reject(new Error("No data rows"));return;}
          resolve({headers:r.meta.fields||Object.keys(r.data[0]),rows:r.data});
        },error:e=>reject(new Error(e.message))});
    } else {
      const reader=new FileReader();
      reader.onload=e=>{try{
        const wb=XLSX.read(e.target.result,{type:"array",cellDates:true});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const data=XLSX.utils.sheet_to_json(ws,{defval:""});
        if(!data.length){reject(new Error("Sheet empty"));return;}
        resolve({headers:Object.keys(data[0]),rows:data});
      }catch(e){reject(new Error(`Excel error: ${e.message}`));}};
      reader.onerror=()=>reject(new Error("Read failed"));
      reader.readAsArrayBuffer(file);
    }
  });
}

/* ═══ FORMATTERS ═══ */
const fmtN=n=>{if(n==null||isNaN(n))return"—";if(n>=1e7)return(n/1e7).toFixed(1)+"Cr";if(n>=1e5)return(n/1e5).toFixed(1)+"L";if(n>=1e3)return(n/1e3).toFixed(1)+"K";return Math.round(n).toLocaleString("en-IN");};
const fmtC=n=>"₹"+fmtN(n);
const fmtDec=(n,d=2)=>(n==null||isNaN(n))?"—":n.toFixed(d);
const sumK=(a,k)=>a.reduce((s,r)=>s+(Number(r[k])||0),0);
const avgK=(a,k)=>{const v=a.filter(r=>r[k]&&!isNaN(r[k])&&r[k]!==0);return v.length?sumK(v,k)/v.length:0;};

/* ═══ SIDEBAR ═══ */
function Sidebar({active,goTo,maxStep}){
  const items=[
    {id:"upload",label:"Data Upload",Icon:Upload,step:1},
    {id:"mapping",label:"Column Mapping",Icon:Columns3,step:2},
    {id:"dashboard",label:"Dashboard",Icon:LayoutDashboard,step:3},
    {id:"analytics",label:"Analytics Builder",Icon:Table2,step:4},
    {id:"insights",label:"AI Insights",Icon:Brain,step:5},
    {id:"reports",label:"Reports",Icon:FileText,step:6},
  ];
  return(
    <div style={{width:210,background:C.bg1,borderRight:`1px solid ${C.border}`,padding:"20px 0",display:"flex",flexDirection:"column",position:"fixed",top:0,left:0,height:"100vh",zIndex:50,overflowY:"auto"}}>
      <div style={{padding:"0 20px 24px",borderBottom:`1px solid ${C.border}`}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <div style={{width:32,height:32,borderRadius:8,background:`linear-gradient(135deg,${C.p},${C.b})`,display:"flex",alignItems:"center",justifyContent:"center"}}><Database size={16} color="#fff"/></div>
          <div><div style={{fontWeight:700,fontSize:15,color:C.t0}}>MediaPulse</div><div style={{fontSize:10,color:C.t2,letterSpacing:1,textTransform:"uppercase"}}>Analytics Engine</div></div>
        </div>
      </div>
      <div style={{padding:"12px 0",flex:1}}>{items.map(({id,label,Icon,step})=>{
        const ok=step<=maxStep,act=active===id;
        return <button key={id} onClick={()=>ok&&goTo(id)} style={{display:"flex",alignItems:"center",gap:10,width:"100%",padding:"11px 20px",border:"none",background:act?C.pBg:"transparent",color:act?C.pL:ok?C.t1:C.t2+"60",cursor:ok?"pointer":"not-allowed",fontFamily:C.font,fontSize:13,fontWeight:act?600:400,borderLeft:act?`3px solid ${C.p}`:"3px solid transparent",opacity:ok?1:.35,transition:"all .15s",textAlign:"left"}}><Icon size={16}/>{label}</button>;
      })}</div>
      <div style={{padding:"16px 20px",borderTop:`1px solid ${C.border}`,fontSize:11,color:C.t2}}>v1.0 · 100% Browser</div>
    </div>
  );
}

/* ═══ SHARED UI ═══ */
function KpiCard({label,value,sub,color}){
  return <div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:"18px 20px",borderTop:`3px solid ${color||C.p}`,flex:1,minWidth:140}}>
    <div style={{fontSize:10,fontWeight:600,color:C.t2,textTransform:"uppercase",letterSpacing:1,marginBottom:8}}>{label}</div>
    <div style={{fontSize:22,fontWeight:700,color:C.t0,marginBottom:4}}>{value}</div>
    {sub && <div style={{fontSize:11,color:C.t1}}>{sub}</div>}
  </div>;
}
function AlertBox({type,items,title}){
  if(!items?.length) return null;
  const isErr=type==="error",col=isErr?C.r:C.a;
  return <div style={{background:col+"12",border:`1px solid ${col}40`,borderRadius:10,padding:16,marginBottom:16}}>
    <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>{isErr?<AlertTriangle size={16} color={col}/>:<Info size={16} color={col}/>}<span style={{fontWeight:600,color:col,fontSize:13}}>{title||(isErr?"Errors":"Warnings")}</span></div>
    {items.map((m,i)=><div key={i} style={{fontSize:12,color:C.t1,paddingLeft:24,marginTop:4,whiteSpace:"pre-wrap"}}>{m}</div>)}
  </div>;
}
const tt={background:C.bg1,border:`1px solid ${C.border}`,borderRadius:8,padding:"8px 12px",fontSize:12,color:C.t0,boxShadow:"0 8px 20px rgba(0,0,0,.4)"};

/* ═══ PAGE: UPLOAD ═══ */
function UploadPage({onFile,onDemo,mediaFile,afFile,valResult,rawRows,onGoMapping,onAfFile}){
  const ref1=useRef(),ref2=useRef();
  const [drag,setDrag]=useState(false);
  return <div>
    <h1 style={{fontSize:24,fontWeight:700,color:C.t0,marginBottom:4}}>Data Upload</h1>
    <p style={{color:C.t2,marginBottom:28,fontSize:14}}>Upload your media performance data and AppsFlyer attribution data.</p>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:20,marginBottom:24}}>
      <div onDragOver={e=>{e.preventDefault();setDrag(true)}} onDragLeave={()=>setDrag(false)} onDrop={e=>{e.preventDefault();setDrag(false);onFile(e.dataTransfer.files[0])}} onClick={()=>ref1.current?.click()} style={{border:`2px dashed ${drag?C.p:C.border}`,borderRadius:14,padding:"48px 32px",textAlign:"center",cursor:"pointer",background:drag?C.pBg:C.bg2,transition:"all .2s"}}>
        <Upload size={28} color={C.p} style={{margin:"0 auto 12px",display:"block"}}/>
        <div style={{fontWeight:600,fontSize:15,color:C.t0,marginBottom:6}}>Media Platform Data</div>
        <div style={{fontSize:12,color:C.t2,marginBottom:12}}>Meta, Google, DV360, Programmatic data</div>
        <div style={{display:"inline-block",background:C.p,color:"#fff",borderRadius:6,padding:"6px 16px",fontSize:12,fontWeight:600}}>CSV or Excel file</div>
        {mediaFile&&<div style={{marginTop:12,fontSize:12,color:C.g,display:"flex",alignItems:"center",gap:6,justifyContent:"center"}}><CheckCircle2 size={14}/>{mediaFile}</div>}
        <input ref={ref1} type="file" accept=".csv,.xlsx,.xls,.tsv" style={{display:"none"}} onChange={e=>{onFile(e.target.files[0]);e.target.value="";}}/>
      </div>
      <div onClick={()=>ref2.current?.click()} style={{border:`2px dashed ${C.border}`,borderRadius:14,padding:"48px 32px",textAlign:"center",cursor:"pointer",background:C.bg2,transition:"all .2s"}}>
        <Upload size={28} color={C.cy} style={{margin:"0 auto 12px",display:"block"}}/>
        <div style={{fontWeight:600,fontSize:15,color:C.t0,marginBottom:6}}>AppsFlyer Data</div>
        <div style={{fontSize:12,color:C.t2,marginBottom:12}}>Attribution, installs, events, retention</div>
        <div style={{display:"inline-block",background:C.cy+"30",color:C.cy,borderRadius:6,padding:"6px 16px",fontSize:12,fontWeight:600}}>CSV or Excel file</div>
        {afFile&&<div style={{marginTop:12,fontSize:12,color:C.g,display:"flex",alignItems:"center",gap:6,justifyContent:"center"}}><CheckCircle2 size={14}/>{afFile}</div>}
        <input ref={ref2} type="file" accept=".csv,.xlsx,.xls,.tsv" style={{display:"none"}} onChange={e=>{onAfFile(e.target.files[0]);e.target.value="";}}/>
      </div>
    </div>
    <div style={{textAlign:"center",marginBottom:24}}>
      <div style={{color:C.t2,fontSize:12,marginBottom:12}}>or</div>
      <button onClick={onDemo} style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:8,padding:"10px 28px",color:C.t0,cursor:"pointer",fontFamily:C.font,fontSize:13,fontWeight:500,display:"inline-flex",alignItems:"center",gap:8}}><Database size={14}/>Load Demo Data</button>
    </div>
    {valResult&&<><AlertBox type="error" items={valResult.errs} title="Fix These Issues"/><AlertBox type="warn" items={valResult.warns} title="Auto-Fixed"/></>}
    {rawRows&&rawRows.length>0&&(()=>{
      const hdrs=Object.keys(rawRows[0]).filter(h=>!h.startsWith("_")).slice(0,8);
      return <div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:20,marginTop:16}}>
        <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:12}}><Eye size={14} color={C.pL}/><span style={{fontWeight:600,fontSize:13,color:C.t0}}>Preview ({rawRows.length} rows, {Object.keys(rawRows[0]).length} columns)</span></div>
        <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
          <thead><tr>{hdrs.map(h=><th key={h} style={{padding:"8px 12px",textAlign:"left",fontSize:10,fontWeight:600,color:C.t2,borderBottom:`1px solid ${C.border}`,whiteSpace:"nowrap",textTransform:"uppercase"}}>{h.length>25?h.slice(0,25)+"…":h}</th>)}</tr></thead>
          <tbody>{rawRows.slice(0,4).map((r,i)=><tr key={i} style={{background:i%2?C.bg1+"80":"transparent"}}>{hdrs.map(h=><td key={h} style={{padding:"7px 12px",color:C.t1,borderBottom:`1px solid ${C.border}15`,whiteSpace:"nowrap",maxWidth:160,overflow:"hidden",textOverflow:"ellipsis"}}>{String(r[h]??"").slice(0,30)}</td>)}</tr>)}</tbody>
        </table></div>
        <div style={{textAlign:"right",marginTop:14}}><button onClick={onGoMapping} style={{background:C.p,color:"#fff",border:"none",borderRadius:8,padding:"10px 24px",cursor:"pointer",fontFamily:C.font,fontSize:14,fontWeight:600,display:"inline-flex",alignItems:"center",gap:8}}>Continue to Mapping <ChevronRight size={16}/></button></div>
      </div>;
    })()}
    {!rawRows&&<div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:20,marginTop:8}}>
      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:12}}><Info size={16} color={C.a}/><span style={{fontWeight:600,fontSize:13,color:C.a}}>Supported Formats</span></div>
      <div style={{fontSize:12,color:C.t1,lineHeight:1.8}}>
        <strong style={{color:C.t0}}>DV360 exports:</strong> Insertion Order, Revenue (Adv Currency), Impressions, Clicks, Complete Views — footer rows auto-stripped<br/>
        <strong style={{color:C.t0}}>Meta/Google/Programmatic:</strong> Standard campaign exports with Spend, Impressions, Clicks<br/>
        <strong style={{color:C.t0}}>AppsFlyer:</strong> Overview reports with Installs, Sessions, Events — comma-formatted numbers OK<br/>
        <strong style={{color:C.t0}}>Dates:</strong> YYYY-MM-DD, YYYY/MM/DD, DD/MM/YYYY all auto-detected<br/>
        <strong style={{color:C.t0}}>Platform column:</strong> Optional — auto-detected from campaign names if missing
      </div>
    </div>}
  </div>;
}

/* ═══ PAGE: MAPPING ═══ */
function MappingPage({headers,colMap,setColMap,onLaunch,valResult,platforms}){
  const sel={background:C.bg1,border:`1px solid ${C.border}`,borderRadius:8,padding:"9px 12px",color:C.t0,fontSize:13,fontFamily:C.font,outline:"none",width:"100%",cursor:"pointer"};
  const allMappable = ["Date","Campaign","Platform","Spend","Impressions","Clicks","Installs","Complete Views","Creative","Advertiser","CTR","CPC","CPI","CPM","ROAS","VTR","Sessions","Device","City","Add to Cart","First Order","Purchase"];
  return <div>
    <h1 style={{fontSize:24,fontWeight:700,color:C.t0,marginBottom:4}}>Column Mapping</h1>
    <p style={{color:C.t2,marginBottom:28,fontSize:14}}>Map each metric to the correct column. Auto-detected where possible.</p>
    <div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:24,marginBottom:20}}>
      <div style={{fontWeight:600,fontSize:14,color:C.t0,marginBottom:16,display:"flex",alignItems:"center",gap:8}}><Columns3 size={16} color={C.p}/>Metric Mapping</div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:16}}>
        {allMappable.map(m=><div key={m}>
          <label style={{fontSize:11,color:REQ_COLS.includes(m)?C.pL:C.t2,marginBottom:4,display:"block",fontWeight:600,textTransform:"uppercase",letterSpacing:.5}}>{m}{REQ_COLS.includes(m)?" *":""}</label>
          <div style={{position:"relative"}}>
            <select value={colMap[m]||""} onChange={e=>setColMap(p=>({...p,[m]:e.target.value||undefined}))} style={{...sel,paddingRight:32}}>
              <option value="">— Select —</option>
              {headers.map(h=><option key={h} value={h}>{h}</option>)}
            </select>
            {colMap[m]&&<CheckCircle2 size={14} color={C.g} style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",pointerEvents:"none"}}/>}
          </div>
        </div>)}
      </div>
    </div>
    {platforms.length>0&&<div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:20,marginBottom:20}}>
      <div style={{fontWeight:600,fontSize:13,color:C.t0,marginBottom:10}}>Platforms Detected</div>
      <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>{platforms.map(p=><span key={p} style={{background:C.p+"30",color:C.pL,border:`1px solid ${C.p}`,borderRadius:20,padding:"5px 14px",fontSize:12,fontWeight:500}}>{p}</span>)}</div>
    </div>}
    {valResult&&<><AlertBox type="error" items={valResult.errs}/><AlertBox type="warn" items={valResult.warns} title="Auto-Fixed"/></>}
    <div style={{textAlign:"right"}}><button onClick={onLaunch} style={{background:C.p,color:"#fff",border:"none",borderRadius:8,padding:"12px 28px",cursor:"pointer",fontFamily:C.font,fontSize:14,fontWeight:600,display:"inline-flex",alignItems:"center",gap:8}}>Launch Dashboard <ChevronRight size={16}/></button></div>
  </div>;
}

/* ═══ PAGE: DASHBOARD ═══ */
function DashboardPage({data,afData}){
  const [range,setRange]=useState("All");
  const filtered=useMemo(()=>{
    if(range==="All")return data;
    const dates=[...new Set(data.map(r=>r.Date))].sort();
    const n=range==="7D"?7:range==="14D"?14:30;
    return data.filter(r=>dates.slice(-n).includes(r.Date));
  },[data,range]);

  const s=sumK(filtered,"Spend"),imp=sumK(filtered,"Impressions"),clk=sumK(filtered,"Clicks"),cv=sumK(filtered,"Complete Views"),inst=sumK(filtered,"Installs");
  const vtr=imp?(cv/imp*100):0, cpm=imp?(s/imp*1000):0, cpv=cv?(s/cv):0, ctr=imp?(clk/imp*100):0;

  const byDate=useMemo(()=>{
    const m={};filtered.forEach(r=>{if(!m[r.Date])m[r.Date]={Date:r.Date,Spend:0,Impressions:0,"Complete Views":0};m[r.Date].Spend+=r.Spend||0;m[r.Date].Impressions+=r.Impressions||0;m[r.Date]["Complete Views"]+=r["Complete Views"]||0;});
    return Object.values(m).sort((a,b)=>a.Date.localeCompare(b.Date));
  },[filtered]);

  const byCampaign=useMemo(()=>{
    const m={};filtered.forEach(r=>{const k=r.Campaign||"Unknown";if(!m[k])m[k]={Campaign:k,Spend:0,Impressions:0,Clicks:0,"Complete Views":0};const c=m[k];c.Spend+=r.Spend||0;c.Impressions+=r.Impressions||0;c.Clicks+=r.Clicks||0;c["Complete Views"]+=r["Complete Views"]||0;});
    return Object.values(m).map(c=>({...c,VTR:c.Impressions?(c["Complete Views"]/c.Impressions*100):0,CPM:c.Impressions?(c.Spend/c.Impressions*1000):0})).sort((a,b)=>b.Spend-a.Spend);
  },[filtered]);

  const byGeo=useMemo(()=>{
    const m={};filtered.forEach(r=>{const g=r.Geo||"Other";if(!m[g])m[g]={name:g,value:0,views:0,impr:0};m[g].value+=r.Spend||0;m[g].views+=r["Complete Views"]||0;m[g].impr+=r.Impressions||0;});
    return Object.values(m).sort((a,b)=>b.value-a.value);
  },[filtered]);

  const byFormat=useMemo(()=>{
    const m={};filtered.forEach(r=>{const f=r.AdFormat||"Other";if(!m[f])m[f]={name:f,Spend:0,"Complete Views":0,Impressions:0};m[f].Spend+=r.Spend||0;m[f]["Complete Views"]+=r["Complete Views"]||0;m[f].Impressions+=r.Impressions||0;});
    return Object.values(m).map(f=>({...f,VTR:f.Impressions?(f["Complete Views"]/f.Impressions*100):0}));
  },[filtered]);

  return <div>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:24}}>
      <div><h1 style={{fontSize:24,fontWeight:700,color:C.t0,marginBottom:2}}>Performance Dashboard</h1><span style={{fontSize:12,color:C.t2}}>{filtered.length.toLocaleString()} records</span></div>
      <div style={{display:"flex",gap:6}}>{["7D","14D","30D","All"].map(r=><button key={r} onClick={()=>setRange(r)} style={{background:range===r?C.p:"transparent",color:range===r?"#fff":C.t1,border:`1px solid ${range===r?C.p:C.border}`,borderRadius:20,padding:"5px 14px",cursor:"pointer",fontSize:12,fontFamily:C.font,fontWeight:500}}>{r}</button>)}</div>
    </div>
    <div style={{display:"flex",gap:14,marginBottom:24,flexWrap:"wrap"}}>
      <KpiCard label="Total Spend" value={fmtC(s)} color={C.p}/>
      <KpiCard label="Impressions" value={fmtN(imp)} color={C.b}/>
      <KpiCard label="Complete Views" value={fmtN(cv)} color={C.g}/>
      <KpiCard label="VTR" value={fmtDec(vtr)+"%"} sub="Views / Impressions" color={C.cy}/>
      <KpiCard label="CPM" value={"₹"+fmtDec(cpm)} color={C.a}/>
      <KpiCard label="CPV" value={"₹"+fmtDec(cpv,3)} sub="Cost per View" color={C.r}/>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"2fr 1fr",gap:16,marginBottom:20}}>
      <div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:20}}>
        <div style={{fontWeight:600,fontSize:14,color:C.t0,marginBottom:16}}>Spend & Views Trend</div>
        {byDate.length>0?<ResponsiveContainer width="100%" height={220}><LineChart data={byDate}><CartesianGrid stroke={C.border} strokeDasharray="3 3"/><XAxis dataKey="Date" tick={{fill:C.t2,fontSize:10}} tickFormatter={d=>d.slice(5)}/><YAxis yAxisId="l" tick={{fill:C.t2,fontSize:10}} tickFormatter={fmtN}/><YAxis yAxisId="r" orientation="right" tick={{fill:C.t2,fontSize:10}} tickFormatter={fmtN}/><Tooltip contentStyle={tt}/><Legend wrapperStyle={{fontSize:11}}/><Line yAxisId="l" type="monotone" dataKey="Spend" stroke={C.p} strokeWidth={2} dot={false}/><Line yAxisId="r" type="monotone" dataKey="Complete Views" stroke={C.g} strokeWidth={2} dot={false}/></LineChart></ResponsiveContainer>:<div style={{color:C.t2,textAlign:"center",padding:40}}>No trend data</div>}
      </div>
      <div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:20}}>
        <div style={{fontWeight:600,fontSize:14,color:C.t0,marginBottom:16}}>Spend by Region</div>
        {byGeo.length>0?<ResponsiveContainer width="100%" height={220}><PieChart><Pie data={byGeo.slice(0,8)} cx="50%" cy="50%" innerRadius={50} outerRadius={78} dataKey="value" nameKey="name" paddingAngle={2}>{byGeo.slice(0,8).map((_,i)=><Cell key={i} fill={COLORS[i%COLORS.length]}/>)}</Pie><Tooltip contentStyle={tt} formatter={v=>fmtC(v)}/><Legend wrapperStyle={{fontSize:10}}/></PieChart></ResponsiveContainer>:<div style={{color:C.t2,textAlign:"center",padding:40}}>No geo data</div>}
      </div>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:20}}>
      <div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:20}}>
        <div style={{fontWeight:600,fontSize:14,color:C.t0,marginBottom:16}}>Ad Format Performance</div>
        {byFormat.length>0?<ResponsiveContainer width="100%" height={200}><BarChart data={byFormat}><XAxis dataKey="name" tick={{fill:C.t1,fontSize:10}}/><YAxis tick={{fill:C.t2,fontSize:10}} tickFormatter={v=>fmtDec(v,0)+"%"}/><Tooltip contentStyle={tt} formatter={v=>fmtDec(v)+"%"}/><Bar dataKey="VTR" radius={[4,4,0,0]} barSize={40}>{byFormat.map((_,i)=><Cell key={i} fill={COLORS[i%COLORS.length]}/>)}</Bar></BarChart></ResponsiveContainer>:<div style={{color:C.t2,textAlign:"center",padding:40}}>No format data</div>}
      </div>
      <div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:20}}>
        <div style={{fontWeight:600,fontSize:14,color:C.t0,marginBottom:16}}>Spend by Ad Format</div>
        {byFormat.length>0?<ResponsiveContainer width="100%" height={200}><BarChart data={byFormat} layout="vertical"><XAxis type="number" tick={{fill:C.t2,fontSize:10}} tickFormatter={fmtN}/><YAxis dataKey="name" type="category" tick={{fill:C.t1,fontSize:10}} width={100}/><Tooltip contentStyle={tt} formatter={v=>fmtC(v)}/><Bar dataKey="Spend" fill={C.p} radius={[0,4,4,0]} barSize={16}/></BarChart></ResponsiveContainer>:<div style={{color:C.t2,textAlign:"center",padding:40}}>No data</div>}
      </div>
    </div>

    {/* AppsFlyer section */}
    {afData&&afData.length>0&&<div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:20,marginBottom:20}}>
      <div style={{fontWeight:600,fontSize:14,color:C.cy,marginBottom:16}}>AppsFlyer Attribution Data</div>
      <div style={{display:"flex",gap:14,marginBottom:16,flexWrap:"wrap"}}>
        <KpiCard label="Installs" value={fmtN(sumK(afData,"Installs"))} color={C.cy}/>
        <KpiCard label="Sessions" value={fmtN(sumK(afData,"Sessions"))} color={C.g}/>
        <KpiCard label="Add to Cart" value={fmtN(sumK(afData,"Add to Cart"))} color={C.a}/>
        <KpiCard label="First Orders" value={fmtN(sumK(afData,"First Order"))} color={C.p}/>
        <KpiCard label="Purchases" value={fmtN(sumK(afData,"Purchase"))} color={C.b}/>
      </div>
    </div>}

    {/* Campaign Table */}
    <div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:20}}>
      <div style={{fontWeight:600,fontSize:14,color:C.t0,marginBottom:16}}>Campaign Breakdown</div>
      <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse"}}><thead><tr>
        {["Campaign","Spend","Impressions","Complete Views","VTR","CPM"].map(h=><th key={h} style={{padding:"10px 14px",textAlign:h==="Campaign"?"left":"right",fontSize:11,fontWeight:600,color:C.t2,textTransform:"uppercase",letterSpacing:.5,borderBottom:`1px solid ${C.border}`,whiteSpace:"nowrap"}}>{h}</th>)}
      </tr></thead><tbody>{byCampaign.map((r,i)=><tr key={i} style={{background:i%2?C.bg1+"60":"transparent"}}>
        <td style={{padding:"10px 14px",fontSize:12,color:C.t0,borderBottom:`1px solid ${C.border}15`,maxWidth:320,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={r.Campaign}>{r.Campaign.length>60?r.Campaign.slice(0,60)+"…":r.Campaign}</td>
        <td style={{padding:"10px 14px",fontSize:12,color:C.t0,textAlign:"right",fontFamily:C.mono,borderBottom:`1px solid ${C.border}15`}}>{fmtC(r.Spend)}</td>
        <td style={{padding:"10px 14px",fontSize:12,color:C.t1,textAlign:"right",fontFamily:C.mono,borderBottom:`1px solid ${C.border}15`}}>{fmtN(r.Impressions)}</td>
        <td style={{padding:"10px 14px",fontSize:12,color:C.t0,textAlign:"right",fontFamily:C.mono,borderBottom:`1px solid ${C.border}15`}}>{fmtN(r["Complete Views"])}</td>
        <td style={{padding:"10px 14px",fontSize:12,textAlign:"right",fontFamily:C.mono,borderBottom:`1px solid ${C.border}15`,color:r.VTR>80?C.g:r.VTR>50?C.a:C.r}}>{fmtDec(r.VTR)}%</td>
        <td style={{padding:"10px 14px",fontSize:12,color:C.t1,textAlign:"right",fontFamily:C.mono,borderBottom:`1px solid ${C.border}15`}}>₹{fmtDec(r.CPM)}</td>
      </tr>)}</tbody></table></div>
    </div>
  </div>;
}

/* ═══ PAGE: ANALYTICS ═══ */
function AnalyticsPage({data}){
  const [dims,setDims]=useState(["Campaign","Geo"]);
  const [metrics,setMetrics]=useState(["Spend","Impressions","Complete Views","VTR"]);
  const allDims=["Campaign","Platform","Geo","AdFormat","Creative","Date"].filter(d=>data.some(r=>r[d]&&String(r[d]).trim()));
  const allMetrics=["Spend","Impressions","Clicks","Complete Views","VTR","CPM","CPV","Installs"];
  const toggle=(arr,set,v)=>set(arr.includes(v)?arr.filter(x=>x!==v):[...arr,v]);

  const pivotData=useMemo(()=>{
    if(!dims.length)return[];
    const m={};data.forEach(r=>{
      const key=dims.map(d=>r[d]||"—").join(" | ");
      if(!m[key])m[key]={_key:key,Spend:0,Impressions:0,Clicks:0,"Complete Views":0,Installs:0};
      const g=m[key];g.Spend+=r.Spend||0;g.Impressions+=r.Impressions||0;g.Clicks+=r.Clicks||0;g["Complete Views"]+=r["Complete Views"]||0;g.Installs+=r.Installs||0;
    });
    return Object.values(m).map(g=>({...g,VTR:g.Impressions?(g["Complete Views"]/g.Impressions*100):0,CPM:g.Impressions?(g.Spend/g.Impressions*1000):0,CPV:g["Complete Views"]?(g.Spend/g["Complete Views"]):0})).sort((a,b)=>b.Spend-a.Spend);
  },[data,dims]);

  const fmtCell=(v,m)=>{
    if(["Spend"].includes(m))return fmtC(v);if(["Impressions","Clicks","Complete Views","Installs"].includes(m))return fmtN(v);
    if(m==="VTR")return fmtDec(v)+"%";if(m==="CPM"||m==="CPV")return"₹"+fmtDec(v);return String(v);
  };

  return <div>
    <h1 style={{fontSize:24,fontWeight:700,color:C.t0,marginBottom:4}}>Analytics Builder</h1>
    <p style={{color:C.t2,marginBottom:24,fontSize:14}}>Build custom pivot tables.</p>
    <div style={{display:"grid",gridTemplateColumns:"220px 1fr",gap:20}}>
      <div>
        <div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:20,marginBottom:16}}>
          <div style={{fontWeight:600,fontSize:13,color:C.t0,marginBottom:12}}>Dimensions</div>
          {allDims.map(d=><button key={d} onClick={()=>toggle(dims,setDims,d)} style={{display:"block",width:"100%",marginBottom:6,background:dims.includes(d)?C.p+"30":"transparent",color:dims.includes(d)?C.pL:C.t1,border:`1px solid ${dims.includes(d)?C.p:C.border}`,borderRadius:8,padding:"8px 14px",cursor:"pointer",fontSize:12,fontFamily:C.font,fontWeight:500,textAlign:"left"}}>{d}</button>)}
        </div>
        <div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:20}}>
          <div style={{fontWeight:600,fontSize:13,color:C.t0,marginBottom:12}}>Metrics</div>
          {allMetrics.map(m=><button key={m} onClick={()=>toggle(metrics,setMetrics,m)} style={{display:"block",width:"100%",marginBottom:6,background:metrics.includes(m)?C.g+"30":"transparent",color:metrics.includes(m)?C.g:C.t1,border:`1px solid ${metrics.includes(m)?C.g+"60":C.border}`,borderRadius:8,padding:"8px 14px",cursor:"pointer",fontSize:12,fontFamily:C.font,fontWeight:500,textAlign:"left"}}>{m}</button>)}
        </div>
      </div>
      <div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:20}}>
        <div style={{display:"flex",justifyContent:"space-between",marginBottom:16}}><span style={{fontWeight:600,fontSize:14,color:C.t0}}>Results</span><span style={{fontSize:11,color:C.t2}}>{pivotData.length} rows</span></div>
        {!dims.length?<div style={{color:C.t2,textAlign:"center",padding:60}}>Select at least one dimension</div>:
        <div style={{overflowX:"auto",maxHeight:520,overflowY:"auto"}}><table style={{width:"100%",borderCollapse:"collapse"}}>
          <thead style={{position:"sticky",top:0,background:C.bg2,zIndex:2}}><tr>
            <th style={{padding:"10px 14px",textAlign:"left",fontSize:11,fontWeight:600,color:C.t2,borderBottom:`1px solid ${C.border}`,textTransform:"uppercase"}}>{dims.join(" / ")}</th>
            {metrics.map(m=><th key={m} style={{padding:"10px 14px",textAlign:"right",fontSize:11,fontWeight:600,color:C.t2,borderBottom:`1px solid ${C.border}`,textTransform:"uppercase"}}>{m}</th>)}
          </tr></thead>
          <tbody>{pivotData.map((r,i)=><tr key={i} style={{background:i%2?C.bg1+"60":"transparent"}}>
            <td style={{padding:"10px 14px",fontSize:12,color:C.t0,borderBottom:`1px solid ${C.border}15`,whiteSpace:"nowrap",maxWidth:300,overflow:"hidden",textOverflow:"ellipsis"}} title={r._key}>{r._key}</td>
            {metrics.map(m=><td key={m} style={{padding:"10px 14px",fontSize:12,color:C.t1,textAlign:"right",fontFamily:C.mono,borderBottom:`1px solid ${C.border}15`}}>{fmtCell(r[m],m)}</td>)}
          </tr>)}</tbody>
        </table></div>}
      </div>
    </div>
  </div>;
}

/* ═══ PAGE: INSIGHTS ═══ */
function InsightsPage({data,afData}){
  const insights=useMemo(()=>{
    const res=[];
    const totalSpend=sumK(data,"Spend")||1;
    const totalCV=sumK(data,"Complete Views")||1;
    const totalImpr=sumK(data,"Impressions")||1;
    const overallVTR=totalCV/totalImpr*100;
    const overallCPV=totalSpend/totalCV;

    // By Geo
    const byGeo={};data.forEach(r=>{const g=r.Geo||"Other";if(!byGeo[g])byGeo[g]={Spend:0,CV:0,Impr:0};byGeo[g].Spend+=r.Spend||0;byGeo[g].CV+=r["Complete Views"]||0;byGeo[g].Impr+=r.Impressions||0;});
    const geoArr=Object.entries(byGeo).map(([name,d])=>({name,vtr:d.Impr?(d.CV/d.Impr*100):0,cpv:d.CV?(d.Spend/d.CV):0,spend:d.Spend,pctSpend:d.Spend/totalSpend*100})).sort((a,b)=>a.cpv-b.cpv);

    if(geoArr.length>0){
      const best=geoArr[0];
      res.push({type:"success",title:`${best.name} delivers best CPV at ₹${best.cpv.toFixed(3)}`,desc:`VTR: ${best.vtr.toFixed(1)}% with ${best.pctSpend.toFixed(0)}% of total budget. Most efficient region for video completions.`,action:`Scale ${best.name} budget by 15-20% to maximize completed views.`});
    }
    if(geoArr.length>1){
      const worst=geoArr[geoArr.length-1];
      if(worst.cpv>overallCPV*1.3){
        res.push({type:"error",title:`${worst.name} underperforming — CPV ₹${worst.cpv.toFixed(3)}`,desc:`${((worst.cpv-overallCPV)/overallCPV*100).toFixed(0)}% above average CPV. VTR: ${worst.vtr.toFixed(1)}%.`,action:`Review ${worst.name} targeting and creative relevance. Consider reducing allocation by 20-30%.`});
      }
    }

    // By Ad Format
    const byFmt={};data.forEach(r=>{const f=r.AdFormat||"Other";if(!byFmt[f])byFmt[f]={Spend:0,CV:0,Impr:0};byFmt[f].Spend+=r.Spend||0;byFmt[f].CV+=r["Complete Views"]||0;byFmt[f].Impr+=r.Impressions||0;});
    const fmtArr=Object.entries(byFmt).map(([name,d])=>({name,vtr:d.Impr?(d.CV/d.Impr*100):0,cpv:d.CV?(d.Spend/d.CV):0})).sort((a,b)=>b.vtr-a.vtr);
    if(fmtArr.length>=2){
      res.push({type:"info",title:`${fmtArr[0].name} achieves highest VTR at ${fmtArr[0].vtr.toFixed(1)}%`,desc:`Compared to ${fmtArr[fmtArr.length-1].name} at ${fmtArr[fmtArr.length-1].vtr.toFixed(1)}% VTR. Format selection significantly impacts view completion rates.`,action:`Prioritize ${fmtArr[0].name} format for awareness campaigns. Use ${fmtArr[fmtArr.length-1].name} for reach goals.`});
    }

    // AppsFlyer insights
    if(afData&&afData.length>0){
      const totalInstalls=sumK(afData,"Installs");
      const totalPurchases=sumK(afData,"Purchase");
      const cvr=totalInstalls?(totalPurchases/totalInstalls*100):0;
      res.push({type:"success",title:`Attribution: ${fmtN(totalInstalls)} installs → ${fmtN(totalPurchases)} purchases`,desc:`Install-to-purchase conversion rate: ${cvr.toFixed(1)}%. AppsFlyer data shows downstream impact of media spend.`,action:`Focus on campaigns driving highest purchase conversion, not just installs.`});
    }

    if(!res.length) res.push({type:"info",title:"Data looks healthy",desc:"No major issues. Continue monitoring.",action:"Upload weekly data for trend detection."});
    return res;
  },[data,afData]);

  const colorMap={success:C.g,warning:C.a,error:C.r,info:C.b};
  return <div>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:24}}>
      <div><h1 style={{fontSize:24,fontWeight:700,color:C.t0,marginBottom:2}}>AI Insights Engine</h1><p style={{color:C.t2,fontSize:14}}>Automated performance diagnosis and recommendations.</p></div>
      <button style={{background:C.p,color:"#fff",border:"none",borderRadius:8,padding:"10px 24px",cursor:"pointer",fontFamily:C.font,fontSize:13,fontWeight:600,display:"flex",alignItems:"center",gap:8}}><Zap size={14}/>Deep AI Analysis</button>
    </div>
    <div style={{display:"flex",flexDirection:"column",gap:16}}>{insights.map((ins,i)=><div key={i} style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:20,borderLeft:`4px solid ${colorMap[ins.type]}`}}>
      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:8}}><div style={{width:12,height:12,borderRadius:6,background:colorMap[ins.type],flexShrink:0}}/><span style={{fontWeight:600,fontSize:14,color:C.t0}}>{ins.title}</span></div>
      <p style={{fontSize:13,color:C.t1,marginBottom:12,paddingLeft:22,lineHeight:1.5}}>{ins.desc}</p>
      <div style={{background:C.bg1,borderRadius:8,padding:"10px 16px",marginLeft:22,fontSize:12,color:C.pL}}>→ {ins.action}</div>
    </div>)}</div>
  </div>;
}

/* ═══ PAGE: REPORTS ═══ */
function ReportsPage({data,afData}){
  const [tab,setTab]=useState("completion");
  const [copied,setCopied]=useState(false);
  const s=sumK(data,"Spend"),imp=sumK(data,"Impressions"),cv=sumK(data,"Complete Views"),clk=sumK(data,"Clicks");
  const vtr=imp?(cv/imp*100):0,cpm=imp?(s/imp*1000):0,cpv=cv?(s/cv):0,ctr=imp?(clk/imp*100):0;
  const dates=[...new Set(data.map(r=>r.Date))].sort();
  const dateRange=dates.length?`${dates[0]} to ${dates[dates.length-1]}`:"—";
  const today=new Date().toLocaleDateString("en-IN",{day:"numeric",month:"long",year:"numeric"});
  const inst=afData?sumK(afData,"Installs"):sumK(data,"Installs");
  const purch=afData?sumK(afData,"Purchase"):0;

  const reports={
    completion:`━━━━━━━━━━━━━━━━━━━━━━━━━━
📊 CAMPAIGN COMPLETION REPORT
━━━━━━━━━━━━━━━━━━━━━━━━━━

📅 Period: ${dateRange}
📅 Generated: ${today}

—— MEDIA PERFORMANCE ——

  Total Spend         ${fmtC(s)}
  Impressions         ${fmtN(imp)}
  Complete Views      ${fmtN(cv)}
  Clicks              ${fmtN(clk)}
  VTR                 ${fmtDec(vtr)}%
  CPM                 ₹${fmtDec(cpm)}
  CPV                 ₹${fmtDec(cpv,3)}
  CTR                 ${fmtDec(ctr)}%
${inst?`\n—— ATTRIBUTION (AppsFlyer) ——\n\n  Installs            ${fmtN(inst)}${purch?`\n  Purchases           ${fmtN(purch)}`:""}\n`:""}
—— RECOMMENDATIONS ——

1. Scale top-performing geo regions by 15-20%
2. Refresh creatives approaching fatigue threshold
3. Optimize In-Stream vs Shorts mix based on VTR data`,
    weekly:`━━━━━━━━━━━━━━━━━━━━━━━━━━
📧 WEEKLY PERFORMANCE EMAIL
━━━━━━━━━━━━━━━━━━━━━━━━━━

Hi Team,

Weekly performance snapshot:

📊 Period: ${dateRange}
💰 Spend: ${fmtC(s)} | 👁 Impressions: ${fmtN(imp)} | ▶ Views: ${fmtN(cv)}
📈 VTR: ${fmtDec(vtr)}% | 💵 CPM: ₹${fmtDec(cpm)} | 🎬 CPV: ₹${fmtDec(cpv,3)}
${inst?`📲 Installs: ${fmtN(inst)}`:""}

Action Items:
• Review underperforming geo regions
• Monitor creative fatigue signals
• Optimize format mix (In-Stream vs Shorts)

Best,
Media Team`,
    executive:`━━━━━━━━━━━━━━━━━━━━━━━━━━
📋 EXECUTIVE SUMMARY
━━━━━━━━━━━━━━━━━━━━━━━━━━

Campaign: DV360 Video — Awareness
Period: ${dateRange}
Budget Utilized: ${fmtC(s)}

KEY METRICS
━━━━━━━━━━
Impressions:      ${fmtN(imp)}
Complete Views:   ${fmtN(cv)}
VTR:              ${fmtDec(vtr)}%
CPM:              ₹${fmtDec(cpm)}
CPV:              ₹${fmtDec(cpv,3)}
${inst?`Installs:         ${fmtN(inst)}`:""}

STRATEGIC OUTLOOK
━━━━━━━━━━━━━━━━
Video completion rates are ${vtr>70?"strong":"moderate"} at ${fmtDec(vtr)}%.
${vtr>70?"Recommend maintaining current strategy with incremental scaling.":"Recommend creative optimization and audience refinement to improve VTR."}`
  };

  const handleCopy=()=>{navigator.clipboard.writeText(reports[tab]).then(()=>{setCopied(true);setTimeout(()=>setCopied(false),2000);}).catch(()=>{const t=document.createElement("textarea");t.value=reports[tab];document.body.appendChild(t);t.select();document.execCommand("copy");document.body.removeChild(t);setCopied(true);setTimeout(()=>setCopied(false),2000);});};

  return <div>
    <h1 style={{fontSize:24,fontWeight:700,color:C.t0,marginBottom:4}}>Report Generator</h1>
    <p style={{color:C.t2,marginBottom:24,fontSize:14}}>Auto-generated reports with your actual KPIs.</p>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20,flexWrap:"wrap",gap:12}}>
      <div style={{display:"flex",gap:8}}>{[{id:"completion",label:"Completion"},{id:"weekly",label:"Weekly Email"},{id:"executive",label:"Executive"}].map(t=><button key={t.id} onClick={()=>setTab(t.id)} style={{background:tab===t.id?C.p:"transparent",color:tab===t.id?"#fff":C.t1,border:`1px solid ${tab===t.id?C.p:C.border}`,borderRadius:8,padding:"8px 20px",cursor:"pointer",fontSize:13,fontFamily:C.font,fontWeight:500}}>{t.label}</button>)}</div>
      <button onClick={handleCopy} style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:8,padding:"8px 18px",cursor:"pointer",color:copied?C.g:C.t1,fontFamily:C.font,fontSize:13,display:"flex",alignItems:"center",gap:6}}>{copied?<CheckCircle2 size={14}/>:<Copy size={14}/>}{copied?"Copied!":"Copy to Clipboard"}</button>
    </div>
    <div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:32}}><pre style={{fontFamily:C.mono,fontSize:13,color:C.t1,lineHeight:1.7,whiteSpace:"pre-wrap",margin:0}}>{reports[tab]}</pre></div>
  </div>;
}

/* ═══ ROOT APP ═══ */
export default function App(){
  const [tab,setTab]=useState("upload");
  const [rawRows,setRawRows]=useState(null);
  const [headers,setHeaders]=useState([]);
  const [colMap,setColMap]=useState({});
  const [data,setData]=useState(null);
  const [afData,setAfData]=useState(null);
  const [valResult,setValResult]=useState(null);
  const [mediaFile,setMediaFile]=useState("");
  const [afFile,setAfFile]=useState("");

  const maxStep=!rawRows?1:!data?2:6;
  const platforms=useMemo(()=>data?[...new Set(data.map(r=>r.Platform).filter(Boolean))]:rawRows&&colMap.Platform?[...new Set(rawRows.map(r=>String(r[colMap.Platform]||"")).filter(Boolean))]:[], [data,rawRows,colMap]);

  const handleFile=useCallback(async(file)=>{
    if(!file)return;
    setMediaFile(file.name);setData(null);setValResult(null);setRawRows(null);setHeaders([]);
    try{
      const{headers:h,rows}=await parseFile(file);
      setHeaders(h);setRawRows(rows);
      const am=autoMap(h);setColMap(am);
      const vr=validate(rows,am,"media");
      if(vr.errs.length)setValResult(vr);
      else setValResult(vr.warns.length?vr:null);
    }catch(e){setValResult({errs:[e.message],warns:[],clean:[]});}
  },[]);

  const handleAfFile=useCallback(async(file)=>{
    if(!file)return;
    setAfFile(file.name);
    try{
      const{headers:h,rows}=await parseFile(file);
      const am=autoMap(h);
      const vr=validate(rows,am,"appsflyer");
      if(!vr.errs.length&&vr.clean.length)setAfData(vr.clean);
      else console.warn("AppsFlyer validation:",vr.errs);
    }catch(e){console.warn("AppsFlyer parse error:",e.message);}
  },[]);

  const handleDemo=useCallback(()=>{
    // Demo with DV360-like data
    const demo=[
      {Date:"2026/03/12","Insertion Order":"HM_SWIGGY_DV360_VIDEO_YT_RS9_NU_HINDI-CITIES(MH & DNCR ONLY)_AWA_INNS-CTV","YouTube Ad":"HM_SWIGGY_INNS-CTV_RS9_15SEC_HIN","Advertiser Currency":"INR","Revenue (Adv Currency)":22570.66,Impressions:295536,Clicks:116,"Complete Views (Video)":260507},
      {Date:"2026/03/12","Insertion Order":"HM_SWIGGY_DV360_VIDEO_YT_RS9_NU_ENGLISH-GEOS_AWA_INNS-CTV","YouTube Ad":"HM_SWIGGY_INNS-CTV_RS9_15SEC_ENG","Advertiser Currency":"INR","Revenue (Adv Currency)":23990.66,Impressions:279788,Clicks:89,"Complete Views (Video)":245829},
      {Date:"2026/03/12","Insertion Order":"HM_SWIGGY_DV360_VIDEO_YT_RS9_NU_TAMILNADU_AWA_INNS-CTV","YouTube Ad":"HM_SWIGGY_INNS-CTV_RS9_15SEC_TAM","Advertiser Currency":"INR","Revenue (Adv Currency)":5200,Impressions:65000,Clicks:20,"Complete Views (Video)":58000},
      {Date:"2026/03/13","Insertion Order":"HM_SWIGGY_DV360_VIDEO_YT_RS9_NU_HINDI-CITIES(MH & DNCR ONLY)_AWA_VVC-CTV","YouTube Ad":"HM_SWIGGY_SHORTS_RS9_15SEC_HIN","Advertiser Currency":"INR","Revenue (Adv Currency)":2231.07,Impressions:33177,Clicks:40,"Complete Views (Video)":13582},
      {Date:"2026/03/13","Insertion Order":"HM_SWIGGY_DV360_VIDEO_YT_RS9_NU_KERALA_AWA_INNS-CTV","YouTube Ad":"HM_SWIGGY_INNS-CTV_RS9_15SEC_MAL","Advertiser Currency":"INR","Revenue (Adv Currency)":3426.89,Impressions:49457,Clicks:16,"Complete Views (Video)":40380},
      {Date:"2026/03/14","Insertion Order":"HM_SWIGGY_DV360_VIDEO_YT_RS9_NU_ANDHRA PRADESH/TELANGANA_AWA_VVC-CTV","YouTube Ad":"HM_SWIGGY_SKIP_RS9_15SEC_TEL","Advertiser Currency":"INR","Revenue (Adv Currency)":2967.68,Impressions:49766,Clicks:58,"Complete Views (Video)":27486},
    ];
    const h=Object.keys(demo[0]);setHeaders(h);setRawRows(demo);
    const am=autoMap(h);setColMap(am);setMediaFile("demo-dv360.csv");
    const vr=validate(demo,am,"media");setData(vr.clean);setValResult(null);setTab("dashboard");
  },[]);

  const handleLaunch=useCallback(()=>{
    const vr=validate(rawRows,colMap,"media");
    if(vr.errs.length){setValResult(vr);return;}
    setData(vr.clean);setValResult(vr.warns.length?vr:null);setTab("dashboard");
  },[rawRows,colMap]);

  return <div style={{minHeight:"100vh",background:`linear-gradient(160deg,${C.bg0} 0%,#0d1222 40%,${C.bg0} 100%)`,color:C.t0,fontFamily:C.font,fontSize:14}}>
    <Sidebar active={tab} goTo={setTab} maxStep={maxStep}/>
    <div style={{marginLeft:210,padding:"32px 40px",minHeight:"100vh"}}>
      {tab==="upload"&&<UploadPage onFile={handleFile} onDemo={handleDemo} mediaFile={mediaFile} afFile={afFile} valResult={valResult} rawRows={rawRows} onGoMapping={()=>setTab("mapping")} onAfFile={handleAfFile}/>}
      {tab==="mapping"&&headers.length>0&&<MappingPage headers={headers} colMap={colMap} setColMap={setColMap} onLaunch={handleLaunch} valResult={valResult} platforms={platforms}/>}
      {tab==="dashboard"&&data&&<DashboardPage data={data} afData={afData}/>}
      {tab==="analytics"&&data&&<AnalyticsPage data={data}/>}
      {tab==="insights"&&data&&<InsightsPage data={data} afData={afData}/>}
      {tab==="reports"&&data&&<ReportsPage data={data} afData={afData}/>}
    </div>
  </div>;
}
