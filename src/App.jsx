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

/* ═══ TOKENS ═══ */
const C={bg0:"#080c18",bg1:"#0f1525",bg2:"#161d30",bg3:"#1e2844",border:"#253050",
p:"#7c5cfc",pL:"#a78bfa",pBg:"rgba(124,92,252,0.12)",g:"#34d399",r:"#f87171",
a:"#fbbf24",b:"#60a5fa",cy:"#22d3ee",t0:"#f1f5f9",t1:"#94a3b8",t2:"#64748b",
font:"'DM Sans',system-ui,sans-serif",mono:"'JetBrains Mono','Fira Code',monospace"};
const CLR=[C.p,C.g,C.b,C.a,C.r,C.cy,"#f472b6","#a3e635","#fb923c","#8b5cf6","#6ee7b7","#fca5a5"];

/* ═══ CAMPAIGN NAME PARSER ═══
   HM_SWIGGY_DV360_VIDEO_YT_RS9_NU_TAMILNADU_AWA_INNS-CTV
   [0]Agency [1]Client [2]Platform [3]AssetType [4]Channel [5]Code [6]UserType [7]Geo [8]Obj [9]Placement
   
   Creative: HM_SWIGGY_INNS-CTV_RS9_15SEC_TAM → Language from last token
   
   AppsFlyer: HM_SWIGGY_SWIM_NOICE_BRN_FBIG_SOCIAL_LC_Geos_RU_DATE_ASSET
*/

const LANG_MAP={ENG:"English",HIN:"Hindi",TAM:"Tamil",TEL:"Telugu",MAL:"Malayalam",KAN:"Kannada",BEN:"Bengali",MAR:"Marathi"};

const GEO_CLEAN=s=>{
  if(!s) return "Other";
  if(/ENGLISH.?GEO/i.test(s)) return "English Geos";
  if(/HINDI.?CITIES.*MH.*DNCR.*ONLY/i.test(s)) return "Hindi (MH & DNCR)";
  if(/HINDI.?CITIES.*EX.?MH/i.test(s)) return "Hindi (Ex-MH-DNCR)";
  if(/HINDI.?CITIES/i.test(s)) return "Hindi Cities";
  if(/TAMILNADU|TAMIL\s?NADU/i.test(s)) return "Tamil Nadu";
  if(/ANDHRA|TELANGANA/i.test(s)) return "AP / Telangana";
  if(/KERALA/i.test(s)) return "Kerala";
  if(/KARNATAKA/i.test(s)) return "Karnataka";
  if(/MAHARASHTRA/i.test(s)) return "Maharashtra";
  // AppsFlyer city patterns
  if(/Mum.*Del|NCR.*Hyd|Blr.*Che/i.test(s)) return "Metro Cities";
  if(/KOLKATA.*KOCHI|T2/i.test(s)) return "Tier 2 Cities";
  return s.length>30?s.slice(0,28)+"…":s;
};

function parseCampaignName(campaign, creative) {
  const c = campaign||"", cr = creative||"";
  const parts = c.split("_");
  const result = {Platform:"Unknown",AssetType:"Other",UserType:"Unknown",Geo:"Other",Placement:"Other",Language:"Other"};

  // --- DV360 pattern: HM_SWIGGY_DV360_VIDEO_YT_RS9_NU_GEO_AWA_PLACEMENT ---
  const knownPlatforms = ["DV360","META","FBIG","FB","GOOGLE","GAD","MEDIASMART","MS","PROG","PROGRAMMATIC","TWITTER","LINKEDIN"];
  const p2raw = (parts[2]||"").toUpperCase().trim();
  const isDV360Pattern = parts.length >= 8 && /^HM$/i.test(parts[0]) && knownPlatforms.includes(p2raw);

  if (isDV360Pattern) {
    // Platform from position 2
    const p2 = p2raw;
    if (p2==="DV360") result.Platform="DV360";
    else if (p2==="FBIG"||p2==="FB"||p2==="META") result.Platform="Meta";
    else if (p2==="GOOGLE"||p2==="GAD") result.Platform="Google";
    else if (p2==="MEDIASMART"||p2==="MS") result.Platform="Mediasmart";
    else if (p2==="PROG"||p2==="PROGRAMMATIC") result.Platform="Programmatic";
    else result.Platform=p2;

    // Asset Type from position 3
    const a3 = (parts[3]||"").toUpperCase();
    if (a3==="VIDEO") result.AssetType="Video";
    else if (a3==="DISPLAY") result.AssetType="Display";
    else if (a3==="SOCIAL") result.AssetType="Social";
    else if (/VIDEO.*DISPLAY|DISPLAY.*VIDEO/i.test(a3)) result.AssetType="Video & Display";
    else result.AssetType=a3||"Other";

    // User Type: scan for NU or RU
    for (const p of parts) {
      const pu = p.toUpperCase().trim();
      if (pu==="NU") { result.UserType="New User"; break; }
      if (pu==="RU") { result.UserType="Returning User"; break; }
    }

    // Geo: everything between UserType (NU/RU) and objective (AWA/PERF/CONV)
    const nuIdx = parts.findIndex(p=>/^(NU|RU)$/i.test(p.trim()));
    const objIdx = parts.findIndex((p,i)=>i>nuIdx&&/^(AWA|PERF|CONV|ACQ|RET)/i.test(p.trim()));
    if (nuIdx>=0 && objIdx>nuIdx) {
      const geoStr = parts.slice(nuIdx+1, objIdx).join("_");
      result.Geo = GEO_CLEAN(geoStr);
    } else if (nuIdx>=0 && nuIdx+1<parts.length) {
      result.Geo = GEO_CLEAN(parts.slice(nuIdx+1).join("_"));
    }

    // Placement: last token
    const lastToken = parts[parts.length-1]||"";
    if (/INNS.?CTV|INSTREAM/i.test(lastToken)) result.Placement="In-Stream CTV";
    else if (/VVC.?CTV/i.test(lastToken)) result.Placement="Shorts / VVC CTV";
    else if (/BUMPER/i.test(lastToken)) result.Placement="Bumper";
    else if (/CTV/i.test(lastToken)) result.Placement="CTV";
    else result.Placement=lastToken;
  }

  // --- AppsFlyer pattern: HM_SWIGGY_SWIM_NOICE_BRN_FBIG_SOCIAL_LC_Geos_RU_... ---
  else if (/FBIG|SWIM/i.test(c)) {
    result.Platform = /FBIG/i.test(c)?"Meta":/GOOGLE/i.test(c)?"Google":"Other";
    result.AssetType = /VIDEO.*DISPLAY/i.test(c)?"Video & Display":/VIDEO/i.test(c)?"Video":/DISPLAY/i.test(c)?"Display":"Social";
    for (const p of parts) {
      const pu=p.toUpperCase().trim();
      if(pu==="NU"){result.UserType="New User";break;}
      if(pu==="RU"){result.UserType="Returning User";break;}
    }
    // Geo: extract between LC and RU/NU
    const lcIdx = parts.findIndex(p=>/^LC$/i.test(p.trim()));
    const ruIdx = parts.findIndex((p,i)=>i>lcIdx&&/^(RU|NU)$/i.test(p.trim()));
    if (lcIdx>=0 && ruIdx>lcIdx) {
      const geoStr = parts.slice(lcIdx+1, ruIdx).join(" ");
      if (/T2|Tier/i.test(geoStr)) result.Geo = "Tier 2 Cities";
      else if (/Mum|Del|NCR|Hyd|Blr|Che|Pun/i.test(geoStr)) result.Geo = "Metro Cities";
      else result.Geo = GEO_CLEAN(geoStr);
    } else result.Geo = GEO_CLEAN(c);
    result.Placement = "Social";
  }

  // --- Language from creative name ---
  if (cr) {
    const crParts = cr.split("_");
    const lastCr = (crParts[crParts.length-1]||"").toUpperCase().trim();
    if (LANG_MAP[lastCr]) result.Language = LANG_MAP[lastCr];
    // Also get ad format from creative
    if (/INNS.?CTV|INSTREAM/i.test(cr)) result.AdFormat = "In-Stream";
    else if (/SHORTS/i.test(cr)) result.AdFormat = "Shorts";
    else if (/SKIP/i.test(cr)) result.AdFormat = "Skippable";
    else if (/BUMPER/i.test(cr)) result.AdFormat = "Bumper";
    else result.AdFormat = "Other";
  } else {
    result.AdFormat = result.Placement.includes("In-Stream")?"In-Stream":result.Placement.includes("Shorts")?"Shorts":"Other";
  }

  return result;
}

/* ═══ ALIASES ═══ */
const ALIASES={
  Date:["date","day","report_date"],
  Campaign:["campaign","campaign_name","insertion order","insertion_order","line item","ad group"],
  Platform:["platform","source","media_source","network","channel"],
  Spend:["spend","cost","amount_spent","budget","revenue (adv currency)","revenue (advertiser currency)","revenue","media cost","total cost"],
  Impressions:["impressions","impr","imps","impression"],
  Clicks:["clicks","click","link_clicks"],
  Installs:["installs","install","conversions","total_attributions"],
  "Complete Views":["complete views (video)","complete views","video completions","trueview views"],
  Creative:["creative","creative_name","youtube ad","ad_creative","ad name"],
  Advertiser:["advertiser","advertiser_name","account"],
  Sessions:["sessions","session"],
  "Add to Cart":["event_count_instamart_add_to_cart","event_count_im_add_to_cart","add_to_cart"],
  "First Order":["event_count_instamart_first_order","event_count_im_first_order_purchase","first_order"],
  Purchase:["event_count_instamart_purchase","event_count_im_purchase","purchase","purchases"],
};
const REQ_COLS=["Date","Campaign"];

function autoMap(h){const m={};for(const k of Object.keys(ALIASES))for(const hd of h){if(hd.toLowerCase().trim()===k.toLowerCase()){m[k]=hd;break;}}for(const[k,al]of Object.entries(ALIASES)){if(m[k])continue;for(const hd of h){if(al.includes(hd.toLowerCase().trim())){m[k]=hd;break;}}}return m;}

/* ═══ UTILITIES ═══ */
function isDV360Footer(row){const v=String(Object.values(row)[0]||"").toLowerCase().trim();return/^(date range|group by|filter by|mrc accredited|reporting numbers)/.test(v)||(v===""&&Object.values(row).filter(x=>String(x||"").trim()).length<=1);}
function parseDate(v){if(!v)return"";let s=String(v).trim();if(v instanceof Date&&!isNaN(v))return v.toISOString().slice(0,10);if(/^\d{4}-\d{2}-\d{2}$/.test(s))return s;let m=s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})/);if(m)return`${m[1]}-${m[2].padStart(2,"0")}-${m[3].padStart(2,"0")}`;m=s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/);if(m)return`${m[3]}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}`;const d=new Date(s);return(!isNaN(d)&&d.getFullYear()>2000)?d.toISOString().slice(0,10):s;}
function parseNum(v){if(v==null)return 0;if(typeof v==="number")return isNaN(v)?0:v;let s=String(v).trim().replace(/[₹$€£"']/g,"").replace(/,/g,"").replace(/%/g,"");if(!s||s==="—"||s==="-"||s==="N/A")return 0;return parseFloat(s)||0;}
function parseFile(file){return new Promise((res,rej)=>{const ext=file.name.split(".").pop().toLowerCase();if(!["csv","tsv","xlsx","xls"].includes(ext)){rej(new Error(`Unsupported: .${ext}`));return;}if(ext==="csv"||ext==="tsv"){Papa.parse(file,{header:true,skipEmptyLines:true,dynamicTyping:false,complete:r=>{if(!r.data?.length){rej(new Error("No data"));return;}res({headers:r.meta.fields||Object.keys(r.data[0]),rows:r.data});},error:e=>rej(new Error(e.message))});}else{const rd=new FileReader();rd.onload=e=>{try{const wb=XLSX.read(e.target.result,{type:"array",cellDates:true});const data=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:""});if(!data.length){rej(new Error("Empty sheet"));return;}res({headers:Object.keys(data[0]),rows:data});}catch(e){rej(new Error(e.message));}};rd.onerror=()=>rej(new Error("Read failed"));rd.readAsArrayBuffer(file);}});}

/* ═══ VALIDATE + ENRICH ═══ */
function validate(rows,cm,type){
  const errs=[],warns=[];
  if(!rows?.length){errs.push("File empty.");return{errs,warns,clean:[]};}
  const hdr=Object.keys(rows[0]);
  for(const r of REQ_COLS)if(!cm[r]||!hdr.includes(cm[r]))errs.push(`Required "${r}" not mapped. Headers: ${hdr.slice(0,8).join(", ")}`);
  if(type==="media"&&!cm.Spend&&!cm.Impressions)errs.push("Need Spend or Impressions.");
  if(errs.length)return{errs,warns,clean:[]};

  const clean=[];let stripped=0,blank=0,badDate=0;
  for(let i=0;i<rows.length;i++){
    const raw=rows[i];
    if(isDV360Footer(raw)){stripped++;continue;}
    const vals=Object.values(raw).filter(v=>v!=null&&String(v).trim());
    if(!vals.length){blank++;continue;}
    const campVal=String(raw[cm.Campaign]||"").trim();
    if(/^(total|sum|grand|subtotal|average)\b/i.test(campVal)){stripped++;continue;}

    const row={_i:i+2};
    const dv=parseDate(raw[cm.Date]);
    if(!/^\d{4}-\d{2}-\d{2}$/.test(dv)){badDate++;if(badDate<=2)warns.push(`Row ${i+2}: Date "${raw[cm.Date]}" unparseable.`);}
    row.Date=dv;
    row.Campaign=campVal;

    // Numeric fields
    for(const f of["Spend","Impressions","Clicks","Installs","Complete Views","Sessions","Add to Cart","First Order","Purchase"])
      row[f]=cm[f]?parseNum(raw[cm[f]]):0;
    // String fields
    row.Creative=cm.Creative?String(raw[cm.Creative]||"").trim():"";
    row.Advertiser=cm.Advertiser?String(raw[cm.Advertiser]||"").trim():"";

    // ─── PARSE CAMPAIGN NAME ───
    const parsed = parseCampaignName(row.Campaign, row.Creative);
    row.Platform = cm.Platform?String(raw[cm.Platform]||"").trim():parsed.Platform;
    if(!row.Platform||row.Platform==="Unknown") row.Platform=parsed.Platform;
    row.AssetType = parsed.AssetType;
    row.UserType = parsed.UserType;
    row.Geo = parsed.Geo;
    row.Placement = parsed.Placement;
    row.Language = parsed.Language;
    row.AdFormat = parsed.AdFormat;

    // ─── CALCULATED METRICS ───
    const sp=row.Spend,imp=row.Impressions,clk=row.Clicks,cv=row["Complete Views"],sess=row.Sessions;
    row.CPM = imp?(sp/imp*1000):0;
    row.CTR = imp?(clk/imp*100):0;
    row.VTR = imp?(cv/imp*100):0;
    row.CPV = cv?(sp/cv):0;
    row.CPC = clk?(sp/clk):0;
    row.CPS = sess?(sp/sess):0; // Cost per Session

    clean.push(row);
  }
  if(stripped)warns.push(`${stripped} metadata/footer/total rows removed.`);
  if(blank)warns.push(`${blank} blank rows removed.`);
  if(badDate>2)warns.push(`${badDate} date issues total.`);
  if(!clean.length)errs.push("No valid rows after cleaning.");
  return{errs,warns,clean};
}

/* ═══ FORMATTERS ═══ */
const fmtN=n=>{if(n==null||isNaN(n))return"—";if(n>=1e7)return(n/1e7).toFixed(1)+"Cr";if(n>=1e5)return(n/1e5).toFixed(1)+"L";if(n>=1e3)return(n/1e3).toFixed(1)+"K";return Math.round(n).toLocaleString("en-IN");};
const fmtC=n=>"₹"+fmtN(n);
const fmtD=(n,d=2)=>(n==null||isNaN(n))?"—":n.toFixed(d);
const sumK=(a,k)=>a.reduce((s,r)=>s+(Number(r[k])||0),0);

/* ═══ SHARED UI ═══ */
const tt={background:C.bg1,border:`1px solid ${C.border}`,borderRadius:8,padding:"8px 12px",fontSize:12,color:C.t0,boxShadow:"0 8px 20px rgba(0,0,0,.4)"};
function KpiCard({label,value,sub,color}){return<div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:"16px 18px",borderTop:`3px solid ${color||C.p}`,flex:1,minWidth:130}}><div style={{fontSize:10,fontWeight:600,color:C.t2,textTransform:"uppercase",letterSpacing:1,marginBottom:6}}>{label}</div><div style={{fontSize:21,fontWeight:700,color:C.t0,marginBottom:3}}>{value}</div>{sub&&<div style={{fontSize:10,color:C.t1}}>{sub}</div>}</div>;}
function AlertBox({type,items,title}){if(!items?.length)return null;const isE=type==="error",col=isE?C.r:C.a;return<div style={{background:col+"12",border:`1px solid ${col}40`,borderRadius:10,padding:16,marginBottom:16}}><div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>{isE?<AlertTriangle size={16} color={col}/>:<Info size={16} color={col}/>}<span style={{fontWeight:600,color:col,fontSize:13}}>{title}</span></div>{items.map((m,i)=><div key={i} style={{fontSize:12,color:C.t1,paddingLeft:24,marginTop:4}}>{m}</div>)}</div>;}
function Sidebar({active,goTo,maxStep}){const items=[{id:"upload",label:"Data Upload",Icon:Upload,step:1},{id:"mapping",label:"Column Mapping",Icon:Columns3,step:2},{id:"dashboard",label:"Dashboard",Icon:LayoutDashboard,step:3},{id:"analytics",label:"Analytics Builder",Icon:Table2,step:4},{id:"insights",label:"AI Insights",Icon:Brain,step:5},{id:"reports",label:"Reports",Icon:FileText,step:6}];return<div style={{width:210,background:C.bg1,borderRight:`1px solid ${C.border}`,padding:"20px 0",display:"flex",flexDirection:"column",position:"fixed",top:0,left:0,height:"100vh",zIndex:50,overflowY:"auto"}}><div style={{padding:"0 20px 24px",borderBottom:`1px solid ${C.border}`}}><div style={{display:"flex",alignItems:"center",gap:10}}><div style={{width:32,height:32,borderRadius:8,background:`linear-gradient(135deg,${C.p},${C.b})`,display:"flex",alignItems:"center",justifyContent:"center"}}><Database size={16} color="#fff"/></div><div><div style={{fontWeight:700,fontSize:15,color:C.t0}}>MediaPulse</div><div style={{fontSize:10,color:C.t2,letterSpacing:1,textTransform:"uppercase"}}>Analytics Engine</div></div></div></div><div style={{padding:"12px 0",flex:1}}>{items.map(({id,label,Icon,step})=>{const ok=step<=maxStep,act=active===id;return<button key={id} onClick={()=>ok&&goTo(id)} style={{display:"flex",alignItems:"center",gap:10,width:"100%",padding:"11px 20px",border:"none",background:act?C.pBg:"transparent",color:act?C.pL:ok?C.t1:C.t2+"60",cursor:ok?"pointer":"not-allowed",fontFamily:C.font,fontSize:13,fontWeight:act?600:400,borderLeft:act?`3px solid ${C.p}`:"3px solid transparent",opacity:ok?1:.35,textAlign:"left"}}><Icon size={16}/>{label}</button>;})}</div><div style={{padding:"16px 20px",borderTop:`1px solid ${C.border}`,fontSize:11,color:C.t2}}>v2.0 · Browser</div></div>;}

/* ═══ MINI BAR CHART COMPONENT ═══ */
function DimBreakdown({data,title,metricKey,metricLabel,formatter}){
  if(!data||!data.length)return null;
  return<div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:20}}>
    <div style={{fontWeight:600,fontSize:14,color:C.t0,marginBottom:16}}>{title}</div>
    <ResponsiveContainer width="100%" height={Math.max(160,data.length*32+40)}>
      <BarChart data={data.slice(0,10)} layout="vertical" margin={{left:10,right:20}}>
        <XAxis type="number" tick={{fill:C.t2,fontSize:10}} tickFormatter={formatter||fmtN}/>
        <YAxis dataKey="name" type="category" tick={{fill:C.t1,fontSize:11}} width={120}/>
        <Tooltip contentStyle={tt} formatter={v=>(formatter||fmtN)(v)}/>
        <Bar dataKey={metricKey} radius={[0,4,4,0]} barSize={16}>{data.slice(0,10).map((_,i)=><Cell key={i} fill={CLR[i%CLR.length]}/>)}</Bar>
      </BarChart>
    </ResponsiveContainer>
  </div>;
}

/* ═══ AGGREGATE HELPER ═══ */
function aggByDim(data, dimKey){
  const m={};
  data.forEach(r=>{
    const k=r[dimKey]||"Other";
    if(!m[k])m[k]={name:k,Spend:0,Impressions:0,Clicks:0,"Complete Views":0,Sessions:0,Installs:0};
    const g=m[k];g.Spend+=r.Spend||0;g.Impressions+=r.Impressions||0;g.Clicks+=r.Clicks||0;g["Complete Views"]+=r["Complete Views"]||0;g.Sessions+=r.Sessions||0;g.Installs+=r.Installs||0;
  });
  return Object.values(m).map(g=>({...g,
    CPM:g.Impressions?(g.Spend/g.Impressions*1000):0,
    CTR:g.Impressions?(g.Clicks/g.Impressions*100):0,
    VTR:g.Impressions?(g["Complete Views"]/g.Impressions*100):0,
    CPV:g["Complete Views"]?(g.Spend/g["Complete Views"]):0,
    CPC:g.Clicks?(g.Spend/g.Clicks):0,
    CPS:g.Sessions?(g.Spend/g.Sessions):0,
  })).sort((a,b)=>b.Spend-a.Spend);
}

/* ═══ PAGE: UPLOAD ═══ */
function UploadPage({onFile,onDemo,mediaFile,afFile,valResult,rawRows,onGoMapping,onAfFile}){
  const ref1=useRef(),ref2=useRef();const[drag,setDrag]=useState(false);
  return<div>
    <h1 style={{fontSize:24,fontWeight:700,color:C.t0,marginBottom:4}}>Data Upload</h1>
    <p style={{color:C.t2,marginBottom:28,fontSize:14}}>Upload media performance data and AppsFlyer attribution data.</p>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:20,marginBottom:24}}>
      <div onDragOver={e=>{e.preventDefault();setDrag(true)}} onDragLeave={()=>setDrag(false)} onDrop={e=>{e.preventDefault();setDrag(false);onFile(e.dataTransfer.files[0])}} onClick={()=>ref1.current?.click()} style={{border:`2px dashed ${drag?C.p:C.border}`,borderRadius:14,padding:"48px 32px",textAlign:"center",cursor:"pointer",background:drag?C.pBg:C.bg2}}>
        <Upload size={28} color={C.p} style={{margin:"0 auto 12px",display:"block"}}/><div style={{fontWeight:600,fontSize:15,color:C.t0,marginBottom:6}}>Media Platform Data</div><div style={{fontSize:12,color:C.t2,marginBottom:12}}>DV360, Meta, Google, Programmatic</div><div style={{display:"inline-block",background:C.p,color:"#fff",borderRadius:6,padding:"6px 16px",fontSize:12,fontWeight:600}}>CSV or Excel</div>
        {mediaFile&&<div style={{marginTop:12,fontSize:12,color:C.g,display:"flex",alignItems:"center",gap:6,justifyContent:"center"}}><CheckCircle2 size={14}/>{mediaFile}</div>}
        <input ref={ref1} type="file" accept=".csv,.xlsx,.xls,.tsv" style={{display:"none"}} onChange={e=>{onFile(e.target.files[0]);e.target.value="";}}/>
      </div>
      <div onClick={()=>ref2.current?.click()} style={{border:`2px dashed ${C.border}`,borderRadius:14,padding:"48px 32px",textAlign:"center",cursor:"pointer",background:C.bg2}}>
        <Upload size={28} color={C.cy} style={{margin:"0 auto 12px",display:"block"}}/><div style={{fontWeight:600,fontSize:15,color:C.t0,marginBottom:6}}>AppsFlyer Data</div><div style={{fontSize:12,color:C.t2,marginBottom:12}}>Attribution, installs, events</div><div style={{display:"inline-block",background:C.cy+"30",color:C.cy,borderRadius:6,padding:"6px 16px",fontSize:12,fontWeight:600}}>CSV or Excel</div>
        {afFile&&<div style={{marginTop:12,fontSize:12,color:C.g,display:"flex",alignItems:"center",gap:6,justifyContent:"center"}}><CheckCircle2 size={14}/>{afFile}</div>}
        <input ref={ref2} type="file" accept=".csv,.xlsx,.xls,.tsv" style={{display:"none"}} onChange={e=>{onAfFile(e.target.files[0]);e.target.value="";}}/>
      </div>
    </div>
    <div style={{textAlign:"center",marginBottom:24}}><div style={{color:C.t2,fontSize:12,marginBottom:12}}>or</div><button onClick={onDemo} style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:8,padding:"10px 28px",color:C.t0,cursor:"pointer",fontFamily:C.font,fontSize:13,fontWeight:500,display:"inline-flex",alignItems:"center",gap:8}}><Database size={14}/>Load Demo Data</button></div>
    {valResult&&<><AlertBox type="error" items={valResult.errs} title="Fix These"/><AlertBox type="warn" items={valResult.warns} title="Auto-Fixed"/></>}
    {rawRows&&rawRows.length>0&&(()=>{const hdrs=Object.keys(rawRows[0]).slice(0,8);return<div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:20,marginTop:16}}><div style={{display:"flex",alignItems:"center",gap:8,marginBottom:12}}><Eye size={14} color={C.pL}/><span style={{fontWeight:600,fontSize:13,color:C.t0}}>Preview ({rawRows.length} rows)</span></div><div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}><thead><tr>{hdrs.map(h=><th key={h} style={{padding:"8px 12px",textAlign:"left",fontSize:10,fontWeight:600,color:C.t2,borderBottom:`1px solid ${C.border}`,whiteSpace:"nowrap",textTransform:"uppercase"}}>{h.length>20?h.slice(0,20)+"…":h}</th>)}</tr></thead><tbody>{rawRows.slice(0,4).map((r,i)=><tr key={i} style={{background:i%2?C.bg1+"80":"transparent"}}>{hdrs.map(h=><td key={h} style={{padding:"7px 12px",color:C.t1,borderBottom:`1px solid ${C.border}15`,whiteSpace:"nowrap",maxWidth:150,overflow:"hidden",textOverflow:"ellipsis"}}>{String(r[h]??"").slice(0,25)}</td>)}</tr>)}</tbody></table></div><div style={{textAlign:"right",marginTop:14}}><button onClick={onGoMapping} style={{background:C.p,color:"#fff",border:"none",borderRadius:8,padding:"10px 24px",cursor:"pointer",fontFamily:C.font,fontSize:14,fontWeight:600,display:"inline-flex",alignItems:"center",gap:8}}>Continue to Mapping <ChevronRight size={16}/></button></div></div>;})()}
  </div>;
}

/* ═══ PAGE: MAPPING ═══ */
function MappingPage({headers,colMap,setColMap,onLaunch,valResult,platforms}){
  const sel={background:C.bg1,border:`1px solid ${C.border}`,borderRadius:8,padding:"9px 12px",color:C.t0,fontSize:13,fontFamily:C.font,outline:"none",width:"100%",cursor:"pointer"};
  const allM=["Date","Campaign","Platform","Spend","Impressions","Clicks","Installs","Complete Views","Creative","Advertiser","Sessions","Add to Cart","First Order","Purchase"];
  return<div>
    <h1 style={{fontSize:24,fontWeight:700,color:C.t0,marginBottom:4}}>Column Mapping</h1>
    <p style={{color:C.t2,marginBottom:12,fontSize:14}}>Auto-detected mappings shown. Adjust if needed.</p>
    <p style={{color:C.t2,marginBottom:28,fontSize:12}}>Platform, Geo, User Type, Language, Placement, Asset Type are auto-parsed from campaign names — no column needed.</p>
    <div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:24,marginBottom:20}}>
      <div style={{fontWeight:600,fontSize:14,color:C.t0,marginBottom:16,display:"flex",alignItems:"center",gap:8}}><Columns3 size={16} color={C.p}/>Metric Mapping</div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:16}}>{allM.map(m=><div key={m}><label style={{fontSize:11,color:REQ_COLS.includes(m)?C.pL:C.t2,marginBottom:4,display:"block",fontWeight:600,textTransform:"uppercase",letterSpacing:.5}}>{m}{REQ_COLS.includes(m)?" *":""}</label><div style={{position:"relative"}}><select value={colMap[m]||""} onChange={e=>setColMap(p=>({...p,[m]:e.target.value||undefined}))} style={{...sel,paddingRight:32}}><option value="">— Select —</option>{headers.map(h=><option key={h} value={h}>{h}</option>)}</select>{colMap[m]&&<CheckCircle2 size={14} color={C.g} style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",pointerEvents:"none"}}/>}</div></div>)}</div>
    </div>
    {valResult&&<><AlertBox type="error" items={valResult.errs} title="Fix These"/><AlertBox type="warn" items={valResult.warns} title="Auto-Fixed"/></>}
    <div style={{textAlign:"right"}}><button onClick={onLaunch} style={{background:C.p,color:"#fff",border:"none",borderRadius:8,padding:"12px 28px",cursor:"pointer",fontFamily:C.font,fontSize:14,fontWeight:600,display:"inline-flex",alignItems:"center",gap:8}}>Launch Dashboard <ChevronRight size={16}/></button></div>
  </div>;
}

/* ═══ PAGE: DASHBOARD ═══ */
function DashboardPage({data,afData}){
  const[range,setRange]=useState("All");
  const filtered=useMemo(()=>{if(range==="All")return data;const dates=[...new Set(data.map(r=>r.Date))].sort();const n=range==="7D"?7:range==="14D"?14:30;return data.filter(r=>dates.slice(-n).includes(r.Date));},[data,range]);

  const s=sumK(filtered,"Spend"),imp=sumK(filtered,"Impressions"),clk=sumK(filtered,"Clicks"),cv=sumK(filtered,"Complete Views"),sess=sumK(filtered,"Sessions"),inst=sumK(filtered,"Installs");
  const cpm=imp?(s/imp*1000):0,ctr=imp?(clk/imp*100):0,vtr=imp?(cv/imp*100):0,cpv=cv?(s/cv):0,cpc=clk?(s/clk):0,cps=sess?(s/sess):0;

  const byDate=useMemo(()=>{const m={};filtered.forEach(r=>{if(!m[r.Date])m[r.Date]={Date:r.Date,Spend:0,Impressions:0,"Complete Views":0};const g=m[r.Date];g.Spend+=r.Spend||0;g.Impressions+=r.Impressions||0;g["Complete Views"]+=r["Complete Views"]||0;});return Object.values(m).sort((a,b)=>a.Date.localeCompare(b.Date));},[filtered]);

  const byGeo=useMemo(()=>aggByDim(filtered,"Geo"),[filtered]);
  const byPlat=useMemo(()=>aggByDim(filtered,"Platform"),[filtered]);
  const byAsset=useMemo(()=>aggByDim(filtered,"AssetType"),[filtered]);
  const byUser=useMemo(()=>aggByDim(filtered,"UserType"),[filtered]);
  const byLang=useMemo(()=>aggByDim(filtered,"Language"),[filtered]);
  const byPlace=useMemo(()=>aggByDim(filtered,"Placement"),[filtered]);
  const byFormat=useMemo(()=>aggByDim(filtered,"AdFormat"),[filtered]);

  const byCampaign=useMemo(()=>aggByDim(filtered,"Campaign"),[filtered]);

  return<div>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:24}}>
      <div><h1 style={{fontSize:24,fontWeight:700,color:C.t0,marginBottom:2}}>Performance Dashboard</h1><span style={{fontSize:12,color:C.t2}}>{filtered.length.toLocaleString()} records</span></div>
      <div style={{display:"flex",gap:6}}>{["7D","14D","30D","All"].map(r=><button key={r} onClick={()=>setRange(r)} style={{background:range===r?C.p:"transparent",color:range===r?"#fff":C.t1,border:`1px solid ${range===r?C.p:C.border}`,borderRadius:20,padding:"5px 14px",cursor:"pointer",fontSize:12,fontFamily:C.font,fontWeight:500}}>{r}</button>)}</div>
    </div>

    {/* KPIs */}
    <div style={{display:"flex",gap:12,marginBottom:24,flexWrap:"wrap"}}>
      <KpiCard label="Total Spend" value={fmtC(s)} color={C.p}/>
      <KpiCard label="Impressions" value={fmtN(imp)} color={C.b}/>
      <KpiCard label="Complete Views" value={fmtN(cv)} color={C.g}/>
      <KpiCard label="CPM" value={"₹"+fmtD(cpm)} sub="Cost/1000 Impr" color={C.a}/>
      <KpiCard label="CTR" value={fmtD(ctr)+"%"} sub="Clicks/Impr" color={C.cy}/>
      <KpiCard label="VTR" value={fmtD(vtr)+"%"} sub="Views/Impr" color={C.g}/>
      <KpiCard label="CPV" value={"₹"+fmtD(cpv,3)} sub="Cost/View" color={C.r}/>
      {clk>0&&<KpiCard label="CPC" value={"₹"+fmtD(cpc)} sub="Cost/Click" color={C.b}/>}
      {sess>0&&<KpiCard label="CPS" value={"₹"+fmtD(cps)} sub="Cost/Session" color="#f472b6"/>}
    </div>

    {/* Trend */}
    <div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:20,marginBottom:20}}>
      <div style={{fontWeight:600,fontSize:14,color:C.t0,marginBottom:16}}>Spend & Complete Views Trend</div>
      {byDate.length>1?<ResponsiveContainer width="100%" height={220}><LineChart data={byDate}><CartesianGrid stroke={C.border} strokeDasharray="3 3"/><XAxis dataKey="Date" tick={{fill:C.t2,fontSize:10}} tickFormatter={d=>d.slice(5)}/><YAxis yAxisId="l" tick={{fill:C.t2,fontSize:10}} tickFormatter={fmtN}/><YAxis yAxisId="r" orientation="right" tick={{fill:C.t2,fontSize:10}} tickFormatter={fmtN}/><Tooltip contentStyle={tt}/><Legend wrapperStyle={{fontSize:11}}/><Line yAxisId="l" type="monotone" dataKey="Spend" stroke={C.p} strokeWidth={2} dot={false}/><Line yAxisId="r" type="monotone" dataKey="Complete Views" stroke={C.g} strokeWidth={2} dot={false}/></LineChart></ResponsiveContainer>:<div style={{color:C.t2,textAlign:"center",padding:40}}>Need 2+ dates for trend</div>}
    </div>

    {/* Dimension Breakdowns — 2x3 grid */}
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:20}}>
      <DimBreakdown data={byGeo} title="Spend by Geo / Region" metricKey="Spend" formatter={fmtC}/>
      <DimBreakdown data={byLang} title="Spend by Language" metricKey="Spend" formatter={fmtC}/>
      <DimBreakdown data={byPlace} title="VTR by Placement" metricKey="VTR" formatter={v=>fmtD(v)+"%"}/>
      <DimBreakdown data={byFormat} title="VTR by Ad Format" metricKey="VTR" formatter={v=>fmtD(v)+"%"}/>
      <DimBreakdown data={byUser} title="Spend by User Type" metricKey="Spend" formatter={fmtC}/>
      <DimBreakdown data={byPlat} title="Spend by Platform" metricKey="Spend" formatter={fmtC}/>
    </div>
    {byAsset.length>1&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:20}}>
      <DimBreakdown data={byAsset} title="Spend by Asset Type" metricKey="Spend" formatter={fmtC}/>
      <DimBreakdown data={byAsset} title="CPM by Asset Type" metricKey="CPM" formatter={v=>"₹"+fmtD(v)}/>
    </div>}

    {/* AppsFlyer */}
    {afData&&afData.length>0&&<div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:20,marginBottom:20}}>
      <div style={{fontWeight:600,fontSize:14,color:C.cy,marginBottom:16}}>AppsFlyer Attribution</div>
      <div style={{display:"flex",gap:12,flexWrap:"wrap"}}>
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
        {["Campaign","Spend","Impr","Views","CPM","CTR","VTR","CPV"].map(h=><th key={h} style={{padding:"10px 12px",textAlign:h==="Campaign"?"left":"right",fontSize:10,fontWeight:600,color:C.t2,textTransform:"uppercase",letterSpacing:.5,borderBottom:`1px solid ${C.border}`,whiteSpace:"nowrap"}}>{h}</th>)}
      </tr></thead><tbody>{byCampaign.map((r,i)=><tr key={i} style={{background:i%2?C.bg1+"60":"transparent"}}>
        <td style={{padding:"10px 12px",fontSize:11,color:C.t0,borderBottom:`1px solid ${C.border}15`,maxWidth:280,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={r.name}>{r.name.length>55?r.name.slice(0,55)+"…":r.name}</td>
        <td style={{padding:"10px 12px",fontSize:11,color:C.t0,textAlign:"right",fontFamily:C.mono,borderBottom:`1px solid ${C.border}15`}}>{fmtC(r.Spend)}</td>
        <td style={{padding:"10px 12px",fontSize:11,color:C.t1,textAlign:"right",fontFamily:C.mono,borderBottom:`1px solid ${C.border}15`}}>{fmtN(r.Impressions)}</td>
        <td style={{padding:"10px 12px",fontSize:11,color:C.t0,textAlign:"right",fontFamily:C.mono,borderBottom:`1px solid ${C.border}15`}}>{fmtN(r["Complete Views"])}</td>
        <td style={{padding:"10px 12px",fontSize:11,color:C.t1,textAlign:"right",fontFamily:C.mono,borderBottom:`1px solid ${C.border}15`}}>₹{fmtD(r.CPM)}</td>
        <td style={{padding:"10px 12px",fontSize:11,color:C.t1,textAlign:"right",fontFamily:C.mono,borderBottom:`1px solid ${C.border}15`}}>{fmtD(r.CTR)}%</td>
        <td style={{padding:"10px 12px",fontSize:11,textAlign:"right",fontFamily:C.mono,borderBottom:`1px solid ${C.border}15`,color:r.VTR>80?C.g:r.VTR>50?C.a:C.r}}>{fmtD(r.VTR)}%</td>
        <td style={{padding:"10px 12px",fontSize:11,color:C.t1,textAlign:"right",fontFamily:C.mono,borderBottom:`1px solid ${C.border}15`}}>₹{fmtD(r.CPV,3)}</td>
      </tr>)}</tbody></table></div>
    </div>
  </div>;
}

/* ═══ PAGE: ANALYTICS ═══ */
function AnalyticsPage({data}){
  const[dims,setDims]=useState(["Geo","Placement"]);
  const[metrics,setMetrics]=useState(["Spend","Complete Views","VTR","CPM"]);
  const allDims=["Campaign","Platform","Geo","Language","Placement","AdFormat","AssetType","UserType","Creative","Date"].filter(d=>data.some(r=>r[d]&&String(r[d]).trim()&&r[d]!=="Other"&&r[d]!=="Unknown"));
  const allMetrics=["Spend","Impressions","Clicks","Complete Views","CPM","CTR","VTR","CPV","CPC","CPS","Installs","Sessions"];
  const toggle=(a,s,v)=>s(a.includes(v)?a.filter(x=>x!==v):[...a,v]);

  const pivotData=useMemo(()=>{
    if(!dims.length)return[];
    const m={};data.forEach(r=>{const key=dims.map(d=>r[d]||"—").join(" | ");if(!m[key])m[key]={_key:key,Spend:0,Impressions:0,Clicks:0,"Complete Views":0,Sessions:0,Installs:0};const g=m[key];g.Spend+=r.Spend||0;g.Impressions+=r.Impressions||0;g.Clicks+=r.Clicks||0;g["Complete Views"]+=r["Complete Views"]||0;g.Sessions+=r.Sessions||0;g.Installs+=r.Installs||0;});
    return Object.values(m).map(g=>({...g,CPM:g.Impressions?(g.Spend/g.Impressions*1000):0,CTR:g.Impressions?(g.Clicks/g.Impressions*100):0,VTR:g.Impressions?(g["Complete Views"]/g.Impressions*100):0,CPV:g["Complete Views"]?(g.Spend/g["Complete Views"]):0,CPC:g.Clicks?(g.Spend/g.Clicks):0,CPS:g.Sessions?(g.Spend/g.Sessions):0})).sort((a,b)=>b.Spend-a.Spend);
  },[data,dims]);

  const fmtCell=(v,m)=>{if(["Spend"].includes(m))return fmtC(v);if(["Impressions","Clicks","Complete Views","Sessions","Installs"].includes(m))return fmtN(v);if(["CTR","VTR"].includes(m))return fmtD(v)+"%";if(["CPM","CPC","CPS"].includes(m))return v?"₹"+fmtD(v):"—";if(m==="CPV")return v?"₹"+fmtD(v,3):"—";return String(v);};

  return<div>
    <h1 style={{fontSize:24,fontWeight:700,color:C.t0,marginBottom:4}}>Analytics Builder</h1>
    <p style={{color:C.t2,marginBottom:24,fontSize:14}}>Custom pivot tables with parsed dimensions.</p>
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
        {!dims.length?<div style={{color:C.t2,textAlign:"center",padding:60}}>Select a dimension</div>:
        <div style={{overflowX:"auto",maxHeight:520,overflowY:"auto"}}><table style={{width:"100%",borderCollapse:"collapse"}}>
          <thead style={{position:"sticky",top:0,background:C.bg2,zIndex:2}}><tr><th style={{padding:"10px 14px",textAlign:"left",fontSize:11,fontWeight:600,color:C.t2,borderBottom:`1px solid ${C.border}`,textTransform:"uppercase"}}>{dims.join(" / ")}</th>{metrics.map(m=><th key={m} style={{padding:"10px 14px",textAlign:"right",fontSize:11,fontWeight:600,color:C.t2,borderBottom:`1px solid ${C.border}`,textTransform:"uppercase",whiteSpace:"nowrap"}}>{m}</th>)}</tr></thead>
          <tbody>{pivotData.map((r,i)=><tr key={i} style={{background:i%2?C.bg1+"60":"transparent"}}><td style={{padding:"10px 14px",fontSize:12,color:C.t0,borderBottom:`1px solid ${C.border}15`,whiteSpace:"nowrap",maxWidth:280,overflow:"hidden",textOverflow:"ellipsis"}} title={r._key}>{r._key}</td>{metrics.map(m=><td key={m} style={{padding:"10px 14px",fontSize:12,color:C.t1,textAlign:"right",fontFamily:C.mono,borderBottom:`1px solid ${C.border}15`,whiteSpace:"nowrap"}}>{fmtCell(r[m],m)}</td>)}</tr>)}</tbody>
        </table></div>}
      </div>
    </div>
  </div>;
}

/* ═══ PAGE: INSIGHTS ═══ */
function InsightsPage({data,afData}){
  const insights=useMemo(()=>{
    const res=[];const ts=sumK(data,"Spend")||1,tcv=sumK(data,"Complete Views")||1,ti=sumK(data,"Impressions")||1;
    const avgVTR=tcv/ti*100,avgCPV=ts/tcv,avgCPM=ts/ti*1000;

    // Geo analysis
    const byGeo=aggByDim(data,"Geo").filter(g=>g.name!=="Other");
    if(byGeo.length>0){const best=byGeo.reduce((a,b)=>a.CPV<b.CPV?a:b);res.push({type:"success",title:`${best.name} — best CPV at ₹${fmtD(best.CPV,3)}`,desc:`VTR ${fmtD(best.VTR)}%, CPM ₹${fmtD(best.CPM)}. ${fmtD(best.Spend/ts*100,0)}% of budget.`,action:`Scale ${best.name} budget 15-20%.`});}
    if(byGeo.length>1){const worst=byGeo.reduce((a,b)=>a.CPV>b.CPV?a:b);if(worst.CPV>avgCPV*1.3)res.push({type:"error",title:`${worst.name} — high CPV ₹${fmtD(worst.CPV,3)}`,desc:`${fmtD((worst.CPV-avgCPV)/avgCPV*100,0)}% above avg. VTR ${fmtD(worst.VTR)}%.`,action:`Reduce ${worst.name} allocation 20-30% or refresh creatives.`});}

    // Placement analysis
    const byPlace=aggByDim(data,"Placement").filter(p=>p.name!=="Other");
    if(byPlace.length>=2){const best=byPlace.reduce((a,b)=>a.VTR>b.VTR?a:b);const worst=byPlace.reduce((a,b)=>a.VTR<b.VTR?a:b);res.push({type:"info",title:`${best.name}: ${fmtD(best.VTR)}% VTR vs ${worst.name}: ${fmtD(worst.VTR)}%`,desc:`Placement significantly impacts view completion. CPV: ${best.name} ₹${fmtD(best.CPV,3)} vs ${worst.name} ₹${fmtD(worst.CPV,3)}.`,action:`Prioritize ${best.name} for awareness. Use ${worst.name} for incremental reach.`});}

    // Language
    const byLang=aggByDim(data,"Language").filter(l=>l.name!=="Other");
    if(byLang.length>=2){const best=byLang.reduce((a,b)=>a.VTR>b.VTR?a:b);res.push({type:"info",title:`${best.name} language content leads with ${fmtD(best.VTR)}% VTR`,desc:`CPM ₹${fmtD(best.CPM)}, CPV ₹${fmtD(best.CPV,3)}. Language-market fit driving completions.`,action:`Invest in more ${best.name} creative variants.`});}

    // User type
    const byUser=aggByDim(data,"UserType").filter(u=>u.name!=="Unknown");
    if(byUser.length>=2){const nu=byUser.find(u=>u.name==="New User"),ru=byUser.find(u=>u.name==="Returning User");if(nu&&ru)res.push({type:nu.CPV<ru.CPV?"success":"warning",title:`New Users CPV ₹${fmtD(nu.CPV,3)} vs Returning ₹${fmtD(ru?.CPV||0,3)}`,desc:`NU VTR ${fmtD(nu.VTR)}%, RU VTR ${fmtD(ru?.VTR||0)}%. ${nu.CPV<ru.CPV?"New users more cost-efficient.":"Returning users more cost-efficient."}`,action:nu.CPV<ru.CPV?"Maintain NU acquisition focus.":"Consider increasing RU retargeting budget."});}

    // AppsFlyer
    if(afData?.length>0){const inst=sumK(afData,"Installs"),purch=sumK(afData,"Purchase");if(inst>0)res.push({type:"success",title:`Attribution: ${fmtN(inst)} installs → ${fmtN(purch)} purchases`,desc:`Install→purchase CVR: ${fmtD(purch/inst*100)}%.`,action:`Optimize for downstream conversions, not just installs.`});}

    if(!res.length)res.push({type:"info",title:"Data looks healthy",desc:"No major issues.",action:"Upload more data for trends."});
    return res;
  },[data,afData]);
  const colMap={success:C.g,warning:C.a,error:C.r,info:C.b};
  return<div><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:24}}><div><h1 style={{fontSize:24,fontWeight:700,color:C.t0,marginBottom:2}}>AI Insights Engine</h1><p style={{color:C.t2,fontSize:14}}>Automated performance diagnosis.</p></div><button style={{background:C.p,color:"#fff",border:"none",borderRadius:8,padding:"10px 24px",cursor:"pointer",fontFamily:C.font,fontSize:13,fontWeight:600,display:"flex",alignItems:"center",gap:8}}><Zap size={14}/>Deep AI Analysis</button></div>
    <div style={{display:"flex",flexDirection:"column",gap:16}}>{insights.map((ins,i)=><div key={i} style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:20,borderLeft:`4px solid ${colMap[ins.type]}`}}><div style={{display:"flex",alignItems:"center",gap:10,marginBottom:8}}><div style={{width:12,height:12,borderRadius:6,background:colMap[ins.type]}}/><span style={{fontWeight:600,fontSize:14,color:C.t0}}>{ins.title}</span></div><p style={{fontSize:13,color:C.t1,marginBottom:12,paddingLeft:22,lineHeight:1.5}}>{ins.desc}</p><div style={{background:C.bg1,borderRadius:8,padding:"10px 16px",marginLeft:22,fontSize:12,color:C.pL}}>→ {ins.action}</div></div>)}</div>
  </div>;
}

/* ═══ PAGE: REPORTS ═══ */
function ReportsPage({data,afData}){
  const[tab,setTab]=useState("completion");const[copied,setCopied]=useState(false);
  const s=sumK(data,"Spend"),imp=sumK(data,"Impressions"),clk=sumK(data,"Clicks"),cv=sumK(data,"Complete Views"),sess=sumK(data,"Sessions");
  const cpm=imp?(s/imp*1000):0,ctr=imp?(clk/imp*100):0,vtr=imp?(cv/imp*100):0,cpv=cv?(s/cv):0,cpc=clk?(s/clk):0,cps=sess?(s/sess):0;
  const dates=[...new Set(data.map(r=>r.Date))].sort();const dr=dates.length?`${dates[0]} to ${dates[dates.length-1]}`:"—";
  const today=new Date().toLocaleDateString("en-IN",{day:"numeric",month:"long",year:"numeric"});
  const inst=afData?sumK(afData,"Installs"):0,purch=afData?sumK(afData,"Purchase"):0;
  const byGeo=aggByDim(data,"Geo");const bestGeo=byGeo.length?byGeo.reduce((a,b)=>a.CPV<b.CPV&&a.CPV>0?a:b):null;

  const reports={
    completion:`━━━━━━━━━━━━━━━━━━━━━━━━━━
📊 CAMPAIGN COMPLETION REPORT
━━━━━━━━━━━━━━━━━━━━━━━━━━

📅 Period: ${dr}
📅 Generated: ${today}

—— MEDIA PERFORMANCE ——

  Total Spend         ${fmtC(s)}
  Impressions         ${fmtN(imp)}
  Clicks              ${fmtN(clk)}
  Complete Views      ${fmtN(cv)}
  CPM                 ₹${fmtD(cpm)}
  CTR                 ${fmtD(ctr)}%
  VTR                 ${fmtD(vtr)}%
  CPV                 ₹${fmtD(cpv,3)}
  CPC                 ₹${fmtD(cpc)}${sess?`\n  Sessions            ${fmtN(sess)}\n  CPS                 ₹${fmtD(cps)}`:""}
${inst?`\n—— ATTRIBUTION ——\n  Installs            ${fmtN(inst)}${purch?`\n  Purchases           ${fmtN(purch)}\n  Install→Purchase    ${fmtD(inst?(purch/inst*100):0)}%`:""}`:"\n"}
—— GEO PERFORMANCE ——
${byGeo.slice(0,6).map(g=>`  ${g.name.padEnd(22)} Spend: ${fmtC(g.Spend).padEnd(10)} VTR: ${fmtD(g.VTR)}%  CPV: ₹${fmtD(g.CPV,3)}`).join("\n")}

—— RECOMMENDATIONS ——
${bestGeo?`1. Scale ${bestGeo.name} — best CPV at ₹${fmtD(bestGeo.CPV,3)}`:"1. Review geo performance"}
2. Optimize creative mix by language
3. Monitor VTR trends weekly`,

    weekly:`━━━━━━━━━━━━━━━━━━━━━━━━━━
📧 WEEKLY PERFORMANCE EMAIL
━━━━━━━━━━━━━━━━━━━━━━━━━━

Hi Team,

📊 Period: ${dr}
💰 Spend: ${fmtC(s)} | 👁 ${fmtN(imp)} impr | ▶ ${fmtN(cv)} views
📈 CPM: ₹${fmtD(cpm)} | CTR: ${fmtD(ctr)}% | VTR: ${fmtD(vtr)}% | CPV: ₹${fmtD(cpv,3)}${sess?`\n📲 Sessions: ${fmtN(sess)} | CPS: ₹${fmtD(cps)}`:""}${inst?`\n🎯 Installs: ${fmtN(inst)}${purch?` | Purchases: ${fmtN(purch)}`:""}`:"\n"}

Best,
Media Team`,

    executive:`━━━━━━━━━━━━━━━━━━━━━━━━━━
📋 EXECUTIVE SUMMARY
━━━━━━━━━━━━━━━━━━━━━━━━━━

Period: ${dr} | Budget: ${fmtC(s)}

Impressions: ${fmtN(imp)} | Views: ${fmtN(cv)} | VTR: ${fmtD(vtr)}%
CPM: ₹${fmtD(cpm)} | CPV: ₹${fmtD(cpv,3)} | CTR: ${fmtD(ctr)}%${inst?`\nInstalls: ${fmtN(inst)}${purch?` | Purchases: ${fmtN(purch)}`:""}`:"\n"}

Video completions ${vtr>70?"strong":"need attention"} at ${fmtD(vtr)}% VTR.
${bestGeo?`Best performing region: ${bestGeo.name} (CPV ₹${fmtD(bestGeo.CPV,3)}).`:""}`
  };

  const cp=()=>{navigator.clipboard.writeText(reports[tab]).then(()=>{setCopied(true);setTimeout(()=>setCopied(false),2000);}).catch(()=>{const t=document.createElement("textarea");t.value=reports[tab];document.body.appendChild(t);t.select();document.execCommand("copy");document.body.removeChild(t);setCopied(true);setTimeout(()=>setCopied(false),2000);});};
  return<div><h1 style={{fontSize:24,fontWeight:700,color:C.t0,marginBottom:4}}>Report Generator</h1><p style={{color:C.t2,marginBottom:24,fontSize:14}}>Auto-generated with your KPIs.</p>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20,flexWrap:"wrap",gap:12}}><div style={{display:"flex",gap:8}}>{[{id:"completion",l:"Completion"},{id:"weekly",l:"Weekly Email"},{id:"executive",l:"Executive"}].map(t=><button key={t.id} onClick={()=>setTab(t.id)} style={{background:tab===t.id?C.p:"transparent",color:tab===t.id?"#fff":C.t1,border:`1px solid ${tab===t.id?C.p:C.border}`,borderRadius:8,padding:"8px 20px",cursor:"pointer",fontSize:13,fontFamily:C.font,fontWeight:500}}>{t.l}</button>)}</div><button onClick={cp} style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:8,padding:"8px 18px",cursor:"pointer",color:copied?C.g:C.t1,fontFamily:C.font,fontSize:13,display:"flex",alignItems:"center",gap:6}}>{copied?<CheckCircle2 size={14}/>:<Copy size={14}/>}{copied?"Copied!":"Copy"}</button></div>
    <div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:32}}><pre style={{fontFamily:C.mono,fontSize:13,color:C.t1,lineHeight:1.7,whiteSpace:"pre-wrap",margin:0}}>{reports[tab]}</pre></div>
  </div>;
}

/* ═══ ROOT ═══ */
export default function App(){
  const[tab,setTab]=useState("upload");const[rawRows,setRawRows]=useState(null);const[headers,setHeaders]=useState([]);
  const[colMap,setColMap]=useState({});const[data,setData]=useState(null);const[afData,setAfData]=useState(null);
  const[valResult,setValResult]=useState(null);const[mediaFile,setMediaFile]=useState("");const[afFile,setAfFile]=useState("");
  const maxStep=!rawRows?1:!data?2:6;
  const platforms=useMemo(()=>data?[...new Set(data.map(r=>r.Platform).filter(Boolean))]:[], [data]);

  const handleFile=useCallback(async f=>{if(!f)return;setMediaFile(f.name);setData(null);setValResult(null);setRawRows(null);setHeaders([]);try{const{headers:h,rows}=await parseFile(f);setHeaders(h);setRawRows(rows);const am=autoMap(h);setColMap(am);const vr=validate(rows,am,"media");setValResult(vr.errs.length||vr.warns.length?vr:null);}catch(e){setValResult({errs:[e.message],warns:[],clean:[]});}},[]);
  const handleAfFile=useCallback(async f=>{if(!f)return;setAfFile(f.name);try{const{headers:h,rows}=await parseFile(f);const am=autoMap(h);const vr=validate(rows,am,"appsflyer");if(!vr.errs.length&&vr.clean.length)setAfData(vr.clean);}catch(e){console.warn("AF error:",e.message);}},[]);
  const handleDemo=useCallback(()=>{const demo=[{Date:"2026/03/12","Insertion Order":"HM_SWIGGY_DV360_VIDEO_YT_RS9_NU_HINDI-CITIES(MH & DNCR ONLY)_AWA_INNS-CTV","YouTube Ad":"HM_SWIGGY_INNS-CTV_RS9_15SEC_HIN","Advertiser Currency":"INR","Revenue (Adv Currency)":22570.66,Impressions:295536,Clicks:116,"Complete Views (Video)":260507},{Date:"2026/03/12","Insertion Order":"HM_SWIGGY_DV360_VIDEO_YT_RS9_NU_ENGLISH-GEOS_AWA_INNS-CTV","YouTube Ad":"HM_SWIGGY_INNS-CTV_RS9_15SEC_ENG","Advertiser Currency":"INR","Revenue (Adv Currency)":23990.66,Impressions:279788,Clicks:89,"Complete Views (Video)":245829},{Date:"2026/03/13","Insertion Order":"HM_SWIGGY_DV360_VIDEO_YT_RS9_NU_TAMILNADU_AWA_VVC-CTV","YouTube Ad":"HM_SWIGGY_SHORTS_RS9_15SEC_TAM","Advertiser Currency":"INR","Revenue (Adv Currency)":1800,Impressions:28000,Clicks:45,"Complete Views (Video)":11200},{Date:"2026/03/13","Insertion Order":"HM_SWIGGY_DV360_VIDEO_YT_RS9_NU_KERALA_AWA_INNS-CTV","YouTube Ad":"HM_SWIGGY_INNS-CTV_RS9_15SEC_MAL","Advertiser Currency":"INR","Revenue (Adv Currency)":3426.89,Impressions:49457,Clicks:16,"Complete Views (Video)":40380},{Date:"2026/03/14","Insertion Order":"HM_SWIGGY_DV360_VIDEO_YT_RS9_NU_ANDHRA PRADESH/TELANGANA_AWA_VVC-CTV","YouTube Ad":"HM_SWIGGY_SKIP_RS9_15SEC_TEL","Advertiser Currency":"INR","Revenue (Adv Currency)":2967.68,Impressions:49766,Clicks:58,"Complete Views (Video)":27486},{Date:"2026/03/14","Insertion Order":"HM_SWIGGY_DV360_VIDEO_YT_RS9_NU_HINDI-CITIES(EX-MH-DNCR)_AWA_INNS-CTV","YouTube Ad":"HM_SWIGGY_INNS-CTV_RS9_15SEC_HIN","Advertiser Currency":"INR","Revenue (Adv Currency)":27946.4,Impressions:403885,Clicks:184,"Complete Views (Video)":354564}];const h=Object.keys(demo[0]);setHeaders(h);setRawRows(demo);const am=autoMap(h);setColMap(am);setMediaFile("demo-dv360.csv");const vr=validate(demo,am,"media");setData(vr.clean);setValResult(null);setTab("dashboard");},[]);
  const handleLaunch=useCallback(()=>{const vr=validate(rawRows,colMap,"media");if(vr.errs.length){setValResult(vr);return;}setData(vr.clean);setValResult(vr.warns.length?vr:null);setTab("dashboard");},[rawRows,colMap]);

  return<div style={{minHeight:"100vh",background:`linear-gradient(160deg,${C.bg0} 0%,#0d1222 40%,${C.bg0} 100%)`,color:C.t0,fontFamily:C.font,fontSize:14}}>
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
