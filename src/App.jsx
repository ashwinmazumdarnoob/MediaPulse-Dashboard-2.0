import { useState, useMemo, useCallback, useRef, useEffect } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  PieChart, Pie, Cell, BarChart, Bar, XAxis, YAxis, Tooltip,
  LineChart, Line, CartesianGrid, Legend, ResponsiveContainer,
} from "recharts";
import {
  Upload, CheckCircle2, AlertTriangle, ChevronRight, Brain, FileText,
  Columns3, LayoutDashboard, Copy, Info, Eye, Table2, Zap, Database,
  Calendar, ChevronDown, X, Lock, LogOut, User, Settings, Target,
  TrendingUp, TrendingDown, BarChart3, ClipboardPaste
} from "lucide-react";

/* ═══ TOKENS ═══ */
const C={bg0:"#080c18",bg1:"#0f1525",bg2:"#161d30",bg3:"#1e2844",border:"#253050",
p:"#7c5cfc",pL:"#a78bfa",pBg:"rgba(124,92,252,0.12)",g:"#34d399",r:"#f87171",
a:"#fbbf24",b:"#60a5fa",cy:"#22d3ee",t0:"#f1f5f9",t1:"#94a3b8",t2:"#64748b",
font:"'DM Sans',system-ui,sans-serif",mono:"'JetBrains Mono','Fira Code',monospace"};
const CLR=[C.p,C.g,C.b,C.a,C.r,C.cy,"#f472b6","#a3e635","#fb923c","#8b5cf6","#6ee7b7","#fca5a5"];

/* ═══════════════════════════════════════
   LOGIN SYSTEM — 5 accounts, localStorage
   ═══════════════════════════════════════ */
const DEFAULT_USERS=[
  {id:"user1",name:"User 1",pass:"media123"},{id:"user2",name:"User 2",pass:"media123"},
  {id:"user3",name:"User 3",pass:"media123"},{id:"user4",name:"User 4",pass:"media123"},
  {id:"user5",name:"User 5",pass:"media123"},
];
function getUsers(){try{const s=localStorage.getItem("mp_users");return s?JSON.parse(s):DEFAULT_USERS;}catch{return DEFAULT_USERS;}}
function saveUsers(u){try{localStorage.setItem("mp_users",JSON.stringify(u));}catch{}}
function getSession(){try{return localStorage.getItem("mp_session")||null;}catch{return null;}}
function setSession(id){try{localStorage.setItem("mp_session",id);}catch{}}
function clearSession(){try{localStorage.removeItem("mp_session");}catch{}}

function LoginPage({onLogin}){
  const[uid,setUid]=useState("");const[pw,setPw]=useState("");const[err,setErr]=useState("");
  const handle=()=>{const users=getUsers();const u=users.find(u=>u.id===uid&&u.pass===pw);
    if(u){setSession(u.id);onLogin(u);}else setErr("Invalid username or password.");};
  return(
    <div style={{minHeight:"100vh",background:`linear-gradient(160deg,${C.bg0} 0%,#0d1222 40%,${C.bg0} 100%)`,display:"flex",alignItems:"center",justifyContent:"center",fontFamily:C.font}}>
      <div style={{width:380,background:C.bg1,border:`1px solid ${C.border}`,borderRadius:16,padding:40}}>
        <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:32}}>
          <div style={{width:40,height:40,borderRadius:10,background:`linear-gradient(135deg,${C.p},${C.b})`,display:"flex",alignItems:"center",justifyContent:"center"}}><Database size={20} color="#fff"/></div>
          <div><div style={{fontWeight:700,fontSize:20,color:C.t0}}>MediaPulse</div><div style={{fontSize:11,color:C.t2,letterSpacing:1,textTransform:"uppercase"}}>Sign In</div></div>
        </div>
        <div style={{marginBottom:16}}>
          <label style={{fontSize:11,color:C.t2,marginBottom:4,display:"block",fontWeight:600,textTransform:"uppercase",letterSpacing:.5}}>Username</label>
          <input value={uid} onChange={e=>setUid(e.target.value)} onKeyDown={e=>e.key==="Enter"&&handle()} placeholder="e.g. user1" style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:8,padding:"10px 14px",color:C.t0,fontSize:14,fontFamily:C.font,outline:"none",width:"100%"}}/>
        </div>
        <div style={{marginBottom:24}}>
          <label style={{fontSize:11,color:C.t2,marginBottom:4,display:"block",fontWeight:600,textTransform:"uppercase",letterSpacing:.5}}>Password</label>
          <input type="password" value={pw} onChange={e=>setPw(e.target.value)} onKeyDown={e=>e.key==="Enter"&&handle()} placeholder="••••••" style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:8,padding:"10px 14px",color:C.t0,fontSize:14,fontFamily:C.font,outline:"none",width:"100%"}}/>
        </div>
        {err&&<div style={{color:C.r,fontSize:12,marginBottom:12}}>{err}</div>}
        <button onClick={handle} style={{background:C.p,color:"#fff",border:"none",borderRadius:8,padding:"12px 0",cursor:"pointer",fontFamily:C.font,fontSize:14,fontWeight:600,width:"100%",display:"flex",alignItems:"center",justifyContent:"center",gap:8}}><Lock size={16}/>Sign In</button>
        <div style={{marginTop:20,fontSize:11,color:C.t2,textAlign:"center"}}>Default accounts: user1–user5 / media123</div>
      </div>
    </div>
  );
}

function ChangePasswordModal({user,onClose}){
  const[cur,setCur]=useState("");const[nw,setNw]=useState("");const[conf,setConf]=useState("");const[msg,setMsg]=useState("");
  const handle=()=>{
    const users=getUsers();const u=users.find(u=>u.id===user.id);
    if(!u||u.pass!==cur){setMsg("Current password is wrong.");return;}
    if(nw.length<4){setMsg("New password must be 4+ characters.");return;}
    if(nw!==conf){setMsg("Passwords don't match.");return;}
    u.pass=nw;saveUsers(users);setMsg("");onClose("Password changed!");
  };
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.7)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:200}}>
      <div style={{width:360,background:C.bg1,border:`1px solid ${C.border}`,borderRadius:14,padding:32}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
          <div style={{fontWeight:700,fontSize:16,color:C.t0}}>Change Password</div>
          <button onClick={()=>onClose(null)} style={{background:"none",border:"none",cursor:"pointer",color:C.t2}}><X size={18}/></button>
        </div>
        {["Current Password","New Password","Confirm New Password"].map((label,i)=>(
          <div key={i} style={{marginBottom:14}}>
            <label style={{fontSize:11,color:C.t2,marginBottom:4,display:"block",fontWeight:600,textTransform:"uppercase",letterSpacing:.5}}>{label}</label>
            <input type="password" value={[cur,nw,conf][i]} onChange={e=>[setCur,setNw,setConf][i](e.target.value)}
              style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:8,padding:"10px 14px",color:C.t0,fontSize:14,fontFamily:C.font,outline:"none",width:"100%"}}/>
          </div>
        ))}
        {msg&&<div style={{color:C.r,fontSize:12,marginBottom:10}}>{msg}</div>}
        <button onClick={handle} style={{background:C.p,color:"#fff",border:"none",borderRadius:8,padding:"10px 0",cursor:"pointer",fontFamily:C.font,fontSize:13,fontWeight:600,width:"100%"}}>Update Password</button>
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════
   CAMPAIGN NAME PARSER (unchanged)
   ═══════════════════════════════════════ */
const LANG_MAP={ENG:"English",HIN:"Hindi",TAM:"Tamil",TEL:"Telugu",MAL:"Malayalam",KAN:"Kannada",BEN:"Bengali",MAR:"Marathi"};
const KNOWN_PLAT=["DV360","META","FBIG","FB","GOOGLE","GAD","MEDIASMART","MS","PROG","PROGRAMMATIC","TWITTER","LINKEDIN","SNAPCHAT","SNAP","SPOTIFY","SPOT","JIOSTAR","JIO"];
function cleanGeo(s){if(!s)return"Other";if(/ENGLISH.?GEO/i.test(s))return"English Geos";if(/HINDI.?CITIES.*MH.*DNCR.*ONLY/i.test(s))return"Hindi (MH & DNCR)";if(/HINDI.?CITIES.*EX.?MH/i.test(s))return"Hindi (Ex-MH-DNCR)";if(/HINDI.?CITIES/i.test(s))return"Hindi Cities";if(/TAMILNADU|TAMIL\s?NADU/i.test(s))return"Tamil Nadu";if(/ANDHRA|TELANGANA/i.test(s))return"AP / Telangana";if(/KERALA/i.test(s))return"Kerala";if(/KARNATAKA/i.test(s))return"Karnataka";if(/MAHARASHTRA/i.test(s))return"Maharashtra";if(/Mum.*Del|NCR.*Hyd|Blr.*Che/i.test(s))return"Metro Cities";if(/KOLKATA.*KOCHI|^T2$/i.test(s))return"Tier 2 Cities";return s.length>30?s.slice(0,28)+"…":s;}
function parseCampaignName(c,cr){const parts=(c||"").split("_");const r={Platform:"Unknown",AssetType:"Other",UserType:"Unknown",Geo:"Other",Placement:"Other",Language:"Other",AdFormat:"Other"};const p2=(parts[2]||"").toUpperCase().trim();if(parts.length>=8&&/^HM$/i.test(parts[0])&&KNOWN_PLAT.includes(p2)){r.Platform=p2==="DV360"?"DV360":["FBIG","FB","META"].includes(p2)?"Meta":["GOOGLE","GAD"].includes(p2)?"Google":["MEDIASMART","MS"].includes(p2)?"Mediasmart":["SNAPCHAT","SNAP"].includes(p2)?"Snapchat":["SPOTIFY","SPOT"].includes(p2)?"Spotify":["JIOSTAR","JIO"].includes(p2)?"JioStar":p2;const a3=(parts[3]||"").toUpperCase();r.AssetType=a3==="VIDEO"?"Video":a3==="DISPLAY"?"Display":a3==="AUDIO"?"Audio":a3==="SOCIAL"?"Social":a3||"Other";for(const p of parts){const pu=p.toUpperCase().trim();if(pu==="NU"){r.UserType="New User";break;}if(pu==="RU"){r.UserType="Returning User";break;}}const nuI=parts.findIndex(p=>/^(NU|RU)$/i.test(p.trim()));const objI=parts.findIndex((p,i)=>i>nuI&&/^(AWA|PERF|CONV|ACQ|RET)/i.test(p.trim()));if(nuI>=0&&objI>nuI)r.Geo=cleanGeo(parts.slice(nuI+1,objI).join("_"));const last=parts[parts.length-1]||"";if(/INNS.?CTV|INSTREAM/i.test(last))r.Placement="In-Stream CTV";else if(/VVC.?CTV/i.test(last))r.Placement="Shorts / VVC CTV";else if(/BUMPER/i.test(last))r.Placement="Bumper";else if(/CTV/i.test(last))r.Placement="CTV";else if(/FEED/i.test(last))r.Placement="Feed";else if(/STORIES/i.test(last))r.Placement="Stories";else if(/REELS/i.test(last))r.Placement="Reels";else if(/NATIVE/i.test(last))r.Placement="Native";else r.Placement=last;}else if(/FBIG|SWIM/i.test(c)){r.Platform=/FBIG/i.test(c)?"Meta":"Other";r.AssetType=/VIDEO.*DISPLAY/i.test(c)?"Video & Display":/VIDEO/i.test(c)?"Video":"Social";for(const p of parts){const pu=p.toUpperCase().trim();if(pu==="NU"){r.UserType="New User";break;}if(pu==="RU"){r.UserType="Returning User";break;}}const lcI=parts.findIndex(p=>/^LC$/i.test(p.trim())),ruI=parts.findIndex((p,i)=>i>lcI&&/^(RU|NU)$/i.test(p.trim()));if(lcI>=0&&ruI>lcI){const gs=parts.slice(lcI+1,ruI).join(" ");r.Geo=/T2|Tier/i.test(gs)?"Tier 2 Cities":/Mum|Del|NCR|Hyd|Blr|Che|Pun/i.test(gs)?"Metro Cities":cleanGeo(gs);}r.Placement="Social";}if(cr){const lastCr=(cr.split("_").pop()||"").toUpperCase().trim();if(LANG_MAP[lastCr])r.Language=LANG_MAP[lastCr];if(/INNS/i.test(cr))r.AdFormat="In-Stream";else if(/SHORTS/i.test(cr))r.AdFormat="Shorts";else if(/SKIP/i.test(cr))r.AdFormat="Skippable";else if(/BUMPER/i.test(cr))r.AdFormat="Bumper";}else{r.AdFormat=r.Placement.includes("In-Stream")?"In-Stream":r.Placement.includes("Shorts")?"Shorts":"Other";}return r;}

/* ═══ ALIASES + PARSE + VALIDATE (unchanged core, fixed) ═══ */
const ALIASES={Date:["date","day","report_date"],Campaign:["campaign","campaign_name","insertion order","insertion_order","line item","ad group"],Platform:["platform","source","media_source","network","channel"],Spend:["spend","cost","amount_spent","budget","revenue (adv currency)","revenue (advertiser currency)","revenue","media cost","total cost"],Impressions:["impressions","impr","imps"],Clicks:["clicks","click","link_clicks"],Installs:["installs","install","conversions","total_attributions"],"Complete Views":["complete views (video)","complete views","video completions","trueview views"],Creative:["creative","creative_name","youtube ad","ad_creative","ad name"],Advertiser:["advertiser","advertiser_name"],Sessions:["sessions","session"],"ATC_1":["event_count_instamart_add_to_cart"],"ATC_2":["event_count_im_add_to_cart"],"FO_1":["event_count_instamart_first_order"],"FO_2":["event_count_im_first_order_purchase"],"P_1":["event_count_instamart_purchase"],"P_2":["event_count_im_purchase"]};
const REQ_COLS=["Date","Campaign"];
function autoMap(h){const m={};for(const k of Object.keys(ALIASES))for(const hd of h){if(hd.toLowerCase().trim()===k.toLowerCase()){m[k]=hd;break;}}for(const[k,al]of Object.entries(ALIASES)){if(m[k])continue;for(const hd of h){if(al.includes(hd.toLowerCase().trim())){m[k]=hd;break;}}}return m;}
function isDV360Footer(row){const vals=Object.entries(row);const first=String(vals[0]?.[1]||"").toLowerCase().trim();if(/^(date range|group by|filter by|mrc accredited|reporting numbers|report time)/.test(first))return true;if(first===""&&Object.values(row).filter(x=>String(x||"").trim()&&String(x).toLowerCase()!=="undefined").length<=1)return true;const dateVal=String(row.Date||row.date||vals[1]?.[1]||"").trim();const campVal=String(row["Insertion Order"]||row.Campaign||row.campaign||vals[2]?.[1]||"").trim();if(!dateVal&&!campVal){const hasLargeNum=Object.values(row).some(v=>{const n=parseFloat(String(v||"").replace(/,/g,""));return!isNaN(n)&&n>10000;});if(hasLargeNum)return true;}if(Object.values(row).filter(v=>String(v).toLowerCase()==="undefined").length>=3)return true;return false;}
function parseDate(v){if(!v)return"";if(v instanceof Date&&!isNaN(v))return v.toISOString().slice(0,10);let s=String(v).trim();if(/^\d{4}-\d{2}-\d{2}$/.test(s))return s;let m=s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})/);if(m)return`${m[1]}-${m[2].padStart(2,"0")}-${m[3].padStart(2,"0")}`;m=s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/);if(m)return`${m[3]}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}`;const d=new Date(s);return(!isNaN(d)&&d.getFullYear()>2000)?d.toISOString().slice(0,10):s;}
function parseNum(v){if(v==null)return 0;if(typeof v==="number")return isNaN(v)?0:v;let s=String(v).trim();if(!s||s==="—"||s==="-"||s==="N/A"||s==="null"||s==="undefined"||s==="NaN"||s==="#N/A"||s==="#VALUE!"||s==="#DIV/0!")return 0;s=s.replace(/["'₹$€£\s]/g,"").replace(/,/g,"").replace(/%/g,"");if(!s)return 0;const n=parseFloat(s);return isNaN(n)?0:n;}
function parseFile(file){return new Promise((res,rej)=>{const ext=file.name.split(".").pop().toLowerCase();if(!["csv","tsv","xlsx","xls"].includes(ext)){rej(new Error(`Unsupported: .${ext}`));return;}if(ext==="csv"||ext==="tsv"){Papa.parse(file,{header:true,skipEmptyLines:true,dynamicTyping:false,complete:r=>{if(!r.data?.length){rej(new Error("No data"));return;}res({headers:r.meta.fields||Object.keys(r.data[0]),rows:r.data});},error:e=>rej(new Error(e.message))});}else{const rd=new FileReader();rd.onload=e=>{try{const wb=XLSX.read(e.target.result,{type:"array",cellDates:true});const data=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:""});if(!data.length){rej(new Error("Empty"));return;}res({headers:Object.keys(data[0]),rows:data});}catch(e){rej(new Error(e.message));}};rd.onerror=()=>rej(new Error("Read failed"));rd.readAsArrayBuffer(file);}});}
function validate(rows,cm,type){const errs=[],warns=[];if(!rows?.length){errs.push("File empty.");return{errs,warns,clean:[]};}const hdr=Object.keys(rows[0]);for(const r of REQ_COLS)if(!cm[r]||!hdr.includes(cm[r]))errs.push(`Required "${r}" not mapped. Headers: ${hdr.slice(0,8).join(", ")}`);if(type==="media"&&!cm.Spend&&!cm.Impressions)errs.push("Need Spend or Impressions.");if(errs.length)return{errs,warns,clean:[]};const clean=[];let stripped=0,blank=0,badDate=0;for(let i=0;i<rows.length;i++){const raw=rows[i];if(isDV360Footer(raw)){stripped++;continue;}const vals=Object.values(raw).filter(v=>v!=null&&String(v).trim()&&String(v).toLowerCase()!=="undefined");if(!vals.length){blank++;continue;}const campVal=String(raw[cm.Campaign]||"").trim();const dateRaw=String(raw[cm.Date]||"").trim();if(!dateRaw&&!campVal){stripped++;continue;}if(/^(total|sum|grand|subtotal|average)\b/i.test(campVal)){stripped++;continue;}if(campVal.toLowerCase()==="undefined"){stripped++;continue;}const dv=parseDate(dateRaw);if(!/^\d{4}-\d{2}-\d{2}$/.test(dv)){badDate++;if(badDate<=2)warns.push(`Row ${i+2}: Date "${dateRaw}" skipped.`);continue;}const row={_i:i+2,Date:dv,Campaign:campVal};for(const f of["Spend","Impressions","Clicks","Installs","Complete Views","Sessions","ATC_1","ATC_2","FO_1","FO_2","P_1","P_2"])row[f]=cm[f]?parseNum(raw[cm[f]]):0;row["Add to Cart"]=(row.ATC_1||0)+(row.ATC_2||0);row["First Order"]=(row.FO_1||0)+(row.FO_2||0);row.Purchase=(row.P_1||0)+(row.P_2||0);row.Creative=cm.Creative?String(raw[cm.Creative]||"").trim():"";row.Advertiser=cm.Advertiser?String(raw[cm.Advertiser]||"").trim():"";const parsed=parseCampaignName(row.Campaign,row.Creative);row.Platform=cm.Platform?String(raw[cm.Platform]||"").trim():parsed.Platform;if(!row.Platform||row.Platform==="Unknown")row.Platform=parsed.Platform;Object.assign(row,{AssetType:parsed.AssetType,UserType:parsed.UserType,Geo:parsed.Geo,Placement:parsed.Placement,Language:parsed.Language,AdFormat:parsed.AdFormat});const sp=row.Spend,imp=row.Impressions,clk=row.Clicks,cv=row["Complete Views"];row.CPM=imp?(sp/imp*1000):0;row.CTR=imp?(clk/imp*100):0;row.VTR=imp?(cv/imp*100):0;row.CPV=cv?(sp/cv):0;row.CPC=clk?(sp/clk):0;clean.push(row);}if(stripped)warns.push(`${stripped} metadata/footer/total rows removed.`);if(blank)warns.push(`${blank} blank rows removed.`);if(badDate>2)warns.push(`${badDate} date issues total.`);if(!clean.length)errs.push("No valid rows after cleaning.");return{errs,warns,clean};}

/* ═══ FORMATTERS ═══ */
const fmtN=n=>{if(n==null||isNaN(n))return"—";if(n>=1e7)return(n/1e7).toFixed(2)+"Cr";if(n>=1e5)return(n/1e5).toFixed(1)+"L";if(n>=1e3)return(n/1e3).toFixed(1)+"K";return Math.round(n).toLocaleString("en-IN");};
const fmtC=n=>"₹"+fmtN(n);const fmtD=(n,d=2)=>(n==null||isNaN(n))?"—":n.toFixed(d);
const sumK=(a,k)=>a.reduce((s,r)=>s+(Number(r[k])||0),0);
const tt={background:C.bg1,border:`1px solid ${C.border}`,borderRadius:8,padding:"8px 12px",fontSize:12,color:C.t0,boxShadow:"0 8px 20px rgba(0,0,0,.4)"};
function aggByDim(data,dk){const m={};data.forEach(r=>{const k=r[dk]||"Other";if(!m[k])m[k]={name:k,Spend:0,Impressions:0,Clicks:0,"Complete Views":0,Sessions:0,Installs:0};const g=m[k];g.Spend+=r.Spend||0;g.Impressions+=r.Impressions||0;g.Clicks+=r.Clicks||0;g["Complete Views"]+=r["Complete Views"]||0;g.Sessions+=r.Sessions||0;g.Installs+=r.Installs||0;});return Object.values(m).map(g=>({...g,CPM:g.Impressions?(g.Spend/g.Impressions*1000):0,CTR:g.Impressions?(g.Clicks/g.Impressions*100):0,VTR:g.Impressions?(g["Complete Views"]/g.Impressions*100):0,CPV:g["Complete Views"]?(g.Spend/g["Complete Views"]):0,CPC:g.Clicks?(g.Spend/g.Clicks):0})).sort((a,b)=>b.Spend-a.Spend);}

/* ═══ OBJECTIVES CONFIG ═══ */
const OBJECTIVES=[
  {id:"video_views",label:"Video Views",icon:"▶",primary:["Complete Views","VTR","CPV"],secondary:["Impressions","CPM","Spend"]},
  {id:"rnf",label:"Reach & Frequency",icon:"📡",primary:["Impressions","CPM","Spend"],secondary:["Clicks","CTR","Complete Views"]},
  {id:"sessions",label:"Sessions / Traffic",icon:"📱",primary:["Sessions","CPS","Clicks"],secondary:["CTR","CPC","Spend"]},
  {id:"installs",label:"App Installs",icon:"📲",primary:["Installs","CPI","Sessions"],secondary:["CPS","Clicks","Spend"]},
  {id:"awareness",label:"Awareness",icon:"🎯",primary:["Impressions","CPM","VTR"],secondary:["Complete Views","Spend","CTR"]},
  {id:"performance",label:"Performance / Conversions",icon:"💰",primary:["Clicks","CPC","CTR"],secondary:["Installs","CPI","Spend"]},
];

/* ═══ SHARED UI ═══ */
function KpiCard({label,value,sub,color,benchmark,actual}){
  const hasBm = benchmark!=null && !isNaN(benchmark) && benchmark>0 && actual!=null;
  const isEfficiency = ["CPM","CTR","VTR","CPV","CPC","CPS","CPI","CPFO"].some(m=>label.includes(m));
  // For cost metrics (CPM, CPV, CPC, CPS, CPI): lower actual = better (green). For rate metrics (CTR, VTR): higher = better.
  const isCost = ["CPM","CPV","CPC","CPS","CPI","CPFO"].some(m=>label.includes(m));
  const diff = hasBm ? ((actual-benchmark)/benchmark*100) : null;
  const isGood = diff!=null ? (isCost ? diff<0 : diff>0) : null;
  return <div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:"14px 16px",borderTop:`3px solid ${color||C.p}`,flex:1,minWidth:125}}>
    <div style={{fontSize:9,fontWeight:600,color:C.t2,textTransform:"uppercase",letterSpacing:1,marginBottom:5}}>{label}</div>
    <div style={{fontSize:19,fontWeight:700,color:C.t0,marginBottom:2}}>{value}</div>
    {sub&&<div style={{fontSize:9,color:C.t1}}>{sub}</div>}
    {hasBm&&<div style={{marginTop:4,fontSize:9,display:"flex",alignItems:"center",gap:4}}>
      <span style={{color:C.t2}}>Plan: {typeof benchmark==="string"?benchmark:fmtD(benchmark)}</span>
      <span style={{color:isGood?C.g:C.r,fontWeight:600}}>{diff>0?"+":""}{fmtD(diff,1)}%</span>
    </div>}
  </div>;
}
function AlertBox({type,items,title}){if(!items?.length)return null;const isE=type==="error",col=isE?C.r:C.a;return<div style={{background:col+"12",border:`1px solid ${col}40`,borderRadius:10,padding:16,marginBottom:16}}><div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>{isE?<AlertTriangle size={16} color={col}/>:<Info size={16} color={col}/>}<span style={{fontWeight:600,color:col,fontSize:13}}>{title}</span></div>{items.map((m,i)=><div key={i} style={{fontSize:12,color:C.t1,paddingLeft:24,marginTop:4}}>{m}</div>)}</div>;}
function DimBreakdown({data,title,metricKey,formatter}){if(!data?.length)return null;return<div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:20}}><div style={{fontWeight:600,fontSize:14,color:C.t0,marginBottom:16}}>{title}</div><ResponsiveContainer width="100%" height={Math.max(150,data.length*30+40)}><BarChart data={data.slice(0,10)} layout="vertical" margin={{left:10,right:20}}><XAxis type="number" tick={{fill:C.t2,fontSize:10}} tickFormatter={formatter||fmtN}/><YAxis dataKey="name" type="category" tick={{fill:C.t1,fontSize:10}} width={120}/><Tooltip contentStyle={tt} formatter={v=>(formatter||fmtN)(v)}/><Bar dataKey={metricKey} radius={[0,4,4,0]} barSize={14}>{data.slice(0,10).map((_,i)=><Cell key={i} fill={CLR[i%CLR.length]}/>)}</Bar></BarChart></ResponsiveContainer></div>;}

/* ═══ SIDEBAR ═══ */
function Sidebar({active,goTo,maxStep,user,onLogout,onChangePw}){
  const items=[{id:"upload",label:"Data Upload",Icon:Upload,step:1},{id:"mapping",label:"Column Mapping",Icon:Columns3,step:2},{id:"dashboard",label:"Dashboard",Icon:LayoutDashboard,step:3},{id:"analytics",label:"Analytics Builder",Icon:Table2,step:4},{id:"insights",label:"AI Insights",Icon:Brain,step:5},{id:"reports",label:"Reports",Icon:FileText,step:6}];
  return<div style={{width:210,background:C.bg1,borderRight:`1px solid ${C.border}`,padding:"20px 0",display:"flex",flexDirection:"column",position:"fixed",top:0,left:0,height:"100vh",zIndex:50,overflowY:"auto"}}>
    <div style={{padding:"0 20px 20px",borderBottom:`1px solid ${C.border}`}}>
      <div style={{display:"flex",alignItems:"center",gap:10}}>
        <div style={{width:32,height:32,borderRadius:8,background:`linear-gradient(135deg,${C.p},${C.b})`,display:"flex",alignItems:"center",justifyContent:"center"}}><Database size={16} color="#fff"/></div>
        <div><div style={{fontWeight:700,fontSize:15,color:C.t0}}>MediaPulse</div><div style={{fontSize:10,color:C.t2,letterSpacing:1,textTransform:"uppercase"}}>Analytics Engine</div></div>
      </div>
    </div>
    <div style={{padding:"12px 0",flex:1}}>{items.map(({id,label,Icon,step})=>{const ok=step<=maxStep,act=active===id;return<button key={id} onClick={()=>ok&&goTo(id)} style={{display:"flex",alignItems:"center",gap:10,width:"100%",padding:"11px 20px",border:"none",background:act?C.pBg:"transparent",color:act?C.pL:ok?C.t1:C.t2+"60",cursor:ok?"pointer":"not-allowed",fontFamily:C.font,fontSize:13,fontWeight:act?600:400,borderLeft:act?`3px solid ${C.p}`:"3px solid transparent",opacity:ok?1:.35,textAlign:"left"}}><Icon size={16}/>{label}</button>;})}</div>
    <div style={{padding:"12px 20px",borderTop:`1px solid ${C.border}`}}>
      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}><User size={14} color={C.pL}/><span style={{fontSize:12,color:C.t0}}>{user?.name||"User"}</span></div>
      <div style={{display:"flex",gap:6}}>
        <button onClick={onChangePw} style={{flex:1,background:C.bg2,border:`1px solid ${C.border}`,borderRadius:6,padding:"6px 0",cursor:"pointer",color:C.t1,fontSize:10,fontFamily:C.font,display:"flex",alignItems:"center",justifyContent:"center",gap:4}}><Settings size={12}/>Password</button>
        <button onClick={onLogout} style={{flex:1,background:C.bg2,border:`1px solid ${C.border}`,borderRadius:6,padding:"6px 0",cursor:"pointer",color:C.r,fontSize:10,fontFamily:C.font,display:"flex",alignItems:"center",justifyContent:"center",gap:4}}><LogOut size={12}/>Logout</button>
      </div>
    </div>
  </div>;
}

/* ═══ DATE RANGE PICKER ═══ */
function DateRangePicker({allDates,from,to,onChange}){const[open,setOpen]=useState(false);const[cF,setCF]=useState(from);const[cT,setCT]=useState(to);const ref=useRef();
  useEffect(()=>{const h=e=>{if(ref.current&&!ref.current.contains(e.target))setOpen(false);};document.addEventListener("mousedown",h);return()=>document.removeEventListener("mousedown",h);},[]);
  const sorted=[...allDates].sort();const minD=sorted[0]||"",maxD=sorted[sorted.length-1]||"";
  const dAgo=n=>{const d=new Date();d.setDate(d.getDate()-n);return d.toISOString().slice(0,10);};const mS=()=>{const d=new Date();return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-01`;};const pMS=()=>{const d=new Date();d.setDate(0);d.setDate(1);return d.toISOString().slice(0,10);};const pME=()=>{const d=new Date();d.setDate(0);return d.toISOString().slice(0,10);};const wS=()=>{const d=new Date();d.setDate(d.getDate()-d.getDay()+1);return d.toISOString().slice(0,10);};const pWS=()=>{const d=new Date();d.setDate(d.getDate()-d.getDay()-6);return d.toISOString().slice(0,10);};const pWE=()=>{const d=new Date();d.setDate(d.getDate()-d.getDay());return d.toISOString().slice(0,10);};const qS=()=>{const d=new Date();return`${d.getFullYear()}-${String(Math.floor(d.getMonth()/3)*3+1).padStart(2,"0")}-01`;};const yS=()=>`${new Date().getFullYear()}-01-01`;const td=new Date().toISOString().slice(0,10);
  const cl=d=>d<minD?minD:d>maxD?maxD:d;
  const presets=[{l:"Today",f:cl(td),t:cl(td)},{l:"Yesterday",f:cl(dAgo(1)),t:cl(dAgo(1))},{l:"Last 7 days",f:cl(dAgo(6)),t:cl(td)},{l:"Last 14 days",f:cl(dAgo(13)),t:cl(td)},{l:"Last 28 days",f:cl(dAgo(27)),t:cl(td)},{l:"Last 30 days",f:cl(dAgo(29)),t:cl(td)},{l:"Last 90 days",f:cl(dAgo(89)),t:cl(td)},{l:"This week",f:cl(wS()),t:cl(td)},{l:"Last week",f:cl(pWS()),t:cl(pWE())},{l:"This month",f:cl(mS()),t:cl(td)},{l:"Last month",f:cl(pMS()),t:cl(pME())},{l:"This quarter",f:cl(qS()),t:cl(td)},{l:"This year",f:cl(yS()),t:cl(td)},{l:"All time",f:minD,t:maxD}];
  const actL=presets.find(p=>p.f===from&&p.t===to)?.l||"Custom";
  const fmtL=d=>{if(!d)return"";const p=d.split("-");return`${p[2]} ${["","Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][+p[1]]} ${p[0]}`;};
  return<div ref={ref} style={{position:"relative"}}><button onClick={()=>setOpen(!open)} style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:8,padding:"8px 16px",cursor:"pointer",fontFamily:C.font,fontSize:12,color:C.t0,display:"flex",alignItems:"center",gap:8,whiteSpace:"nowrap"}}><Calendar size={14} color={C.pL}/><span>{fmtL(from)} — {fmtL(to)}</span><span style={{background:C.p+"30",color:C.pL,borderRadius:4,padding:"2px 8px",fontSize:10,fontWeight:600}}>{actL}</span><ChevronDown size={14} color={C.t2}/></button>
    {open&&<div style={{position:"absolute",top:"100%",right:0,marginTop:8,background:C.bg1,border:`1px solid ${C.border}`,borderRadius:12,boxShadow:"0 20px 60px rgba(0,0,0,.6)",zIndex:100,display:"flex",width:540,overflow:"hidden"}}><div style={{width:180,borderRight:`1px solid ${C.border}`,padding:"12px 0",maxHeight:400,overflowY:"auto"}}>{presets.map(p=><button key={p.l} onClick={()=>{onChange(p.f,p.t);setOpen(false);}} style={{display:"block",width:"100%",padding:"8px 16px",border:"none",background:actL===p.l?C.pBg:"transparent",color:actL===p.l?C.pL:C.t1,cursor:"pointer",fontFamily:C.font,fontSize:11,fontWeight:actL===p.l?600:400,textAlign:"left",borderLeft:actL===p.l?`3px solid ${C.p}`:"3px solid transparent"}}>{p.l}</button>)}</div><div style={{flex:1,padding:20}}><div style={{fontWeight:600,fontSize:13,color:C.t0,marginBottom:14}}>Custom Range</div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:14}}><div><label style={{fontSize:10,color:C.t2,display:"block",marginBottom:3,fontWeight:600}}>START</label><input type="date" value={cF} min={minD} max={maxD} onChange={e=>setCF(e.target.value)} style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:6,padding:"8px",color:C.t0,fontSize:12,fontFamily:C.font,outline:"none",width:"100%",colorScheme:"dark"}}/></div><div><label style={{fontSize:10,color:C.t2,display:"block",marginBottom:3,fontWeight:600}}>END</label><input type="date" value={cT} min={minD} max={maxD} onChange={e=>setCT(e.target.value)} style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:6,padding:"8px",color:C.t0,fontSize:12,fontFamily:C.font,outline:"none",width:"100%",colorScheme:"dark"}}/></div></div><button onClick={()=>{if(cF&&cT){onChange(cF<=cT?cF:cT,cF<=cT?cT:cF);setOpen(false);}}} style={{background:C.p,color:"#fff",border:"none",borderRadius:8,padding:"8px 0",cursor:"pointer",fontFamily:C.font,fontSize:12,fontWeight:600,width:"100%"}}>Apply</button><div style={{marginTop:10,fontSize:10,color:C.t2}}>Data: {fmtL(minD)} — {fmtL(maxD)} · {allDates.filter(d=>d>=from&&d<=to).length} days selected</div></div></div>}
  </div>;
}

/* ═══ BENCHMARK INPUT ═══ */
function parseBenchmarkText(text){
  const lines=text.trim().split("\n").map(l=>l.split(/[\t,]+/).map(c=>c.trim()));
  if(lines.length<2)return null;
  const hdr=lines[0].map(h=>h.toLowerCase());
  const platIdx=hdr.findIndex(h=>/platform/i.test(h));
  if(platIdx<0)return null;
  const bm={};
  const metricMap={spend:"Spend",impressions:"Impressions",cpm:"CPM",ctr:"CTR",vtr:"VTR",cpv:"CPV",cpc:"CPC",cps:"CPS",cpi:"CPI"};
  lines.slice(1).forEach(row=>{
    const plat=row[platIdx];if(!plat)return;
    bm[plat]={};
    hdr.forEach((h,i)=>{
      if(i===platIdx)return;
      const clean=h.replace(/planned[_ ]?/i,"").replace(/benchmark[_ ]?/i,"").trim().toLowerCase();
      const mk=Object.entries(metricMap).find(([k])=>clean.includes(k));
      if(mk)bm[plat][mk[1]]=parseNum(row[i]);
    });
  });
  return Object.keys(bm).length?bm:null;
}

function BenchmarkInput({benchmarks,setBenchmarks}){
  const[mode,setMode]=useState("paste");const[text,setText]=useState("");const[msg,setMsg]=useState("");
  const ref=useRef();
  const handlePaste=()=>{const bm=parseBenchmarkText(text);if(bm){setBenchmarks(bm);setMsg(`✅ Loaded benchmarks for ${Object.keys(bm).length} platforms.`);}else setMsg("❌ Could not parse. Need 'Platform' column + metric columns (CPM, VTR, etc.).");};
  const handleFile=async(file)=>{if(!file)return;try{const{rows}=await new Promise((res,rej)=>{Papa.parse(file,{header:true,skipEmptyLines:true,dynamicTyping:false,complete:r=>res({rows:r.data}),error:e=>rej(e)});});const lines=[Object.keys(rows[0]).join("\t"),...rows.map(r=>Object.values(r).join("\t"))].join("\n");const bm=parseBenchmarkText(lines);if(bm){setBenchmarks(bm);setMsg(`✅ Loaded benchmarks for ${Object.keys(bm).length} platforms.`);}else setMsg("❌ Need 'Platform' column + metric columns.");}catch{setMsg("❌ File error.");}};
  const sampleText="Platform\tPlanned Spend\tPlanned Impressions\tPlanned CPM\tPlanned CTR\tPlanned VTR\tPlanned CPV\tPlanned CPC\nDV360\t6000000\t90000000\t65\t0.04\t60\t0.10\t150\nMeta\t2000000\t5000000\t400\t3.5\t\t\t3\nGoogle\t1500000\t10000000\t150\t1.0\t40\t0.15\t15";
  return<div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:20,marginBottom:20}}>
    <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14}}>
      <div style={{fontWeight:600,fontSize:14,color:C.t0,display:"flex",alignItems:"center",gap:8}}><Target size={16} color={C.a}/>Planned Benchmarks (Platform Level)</div>
      <div style={{display:"flex",gap:6}}>{[{id:"paste",l:"Paste Table",I:ClipboardPaste},{id:"upload",l:"Upload CSV",I:Upload}].map(({id,l,I})=><button key={id} onClick={()=>setMode(id)} style={{background:mode===id?C.p+"30":"transparent",color:mode===id?C.pL:C.t1,border:`1px solid ${mode===id?C.p:C.border}`,borderRadius:6,padding:"5px 12px",cursor:"pointer",fontSize:11,fontFamily:C.font,display:"flex",alignItems:"center",gap:4}}><I size={12}/>{l}</button>)}</div>
    </div>
    {mode==="paste"&&<>
      <textarea value={text} onChange={e=>setText(e.target.value)} placeholder={sampleText} rows={5} style={{background:C.bg1,border:`1px solid ${C.border}`,borderRadius:8,padding:12,color:C.t0,fontSize:11,fontFamily:C.mono,outline:"none",width:"100%",resize:"vertical",lineHeight:1.5}}/>
      <div style={{display:"flex",gap:10,marginTop:10,alignItems:"center"}}>
        <button onClick={handlePaste} style={{background:C.p,color:"#fff",border:"none",borderRadius:6,padding:"8px 20px",cursor:"pointer",fontFamily:C.font,fontSize:12,fontWeight:600}}>Parse Benchmarks</button>
        <button onClick={()=>setText(sampleText)} style={{background:"transparent",border:`1px solid ${C.border}`,borderRadius:6,padding:"8px 16px",cursor:"pointer",color:C.t1,fontFamily:C.font,fontSize:11}}>Load Example</button>
      </div>
    </>}
    {mode==="upload"&&<div><input ref={ref} type="file" accept=".csv,.xlsx,.xls,.tsv" style={{display:"none"}} onChange={e=>{handleFile(e.target.files[0]);e.target.value="";}}/><button onClick={()=>ref.current?.click()} style={{background:C.bg1,border:`2px dashed ${C.border}`,borderRadius:8,padding:"20px 0",cursor:"pointer",width:"100%",color:C.t1,fontFamily:C.font,fontSize:12,textAlign:"center"}}>Click to upload benchmark CSV</button></div>}
    {msg&&<div style={{marginTop:8,fontSize:12,color:msg.startsWith("✅")?C.g:C.r}}>{msg}</div>}
    {benchmarks&&Object.keys(benchmarks).length>0&&<div style={{marginTop:12,overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:10}}>
      <thead><tr>{["Platform","Spend","Impr","CPM","CTR","VTR","CPV","CPC","CPS","CPI"].map(h=><th key={h} style={{padding:"6px 8px",textAlign:h==="Platform"?"left":"right",color:C.t2,borderBottom:`1px solid ${C.border}`,fontWeight:600,textTransform:"uppercase",fontSize:9}}>{h}</th>)}</tr></thead>
      <tbody>{Object.entries(benchmarks).map(([plat,bm],i)=><tr key={plat} style={{background:i%2?C.bg1+"60":"transparent"}}><td style={{padding:"6px 8px",color:C.t0,borderBottom:`1px solid ${C.border}15`}}>{plat}</td>{["Spend","Impressions","CPM","CTR","VTR","CPV","CPC","CPS","CPI"].map(m=><td key={m} style={{padding:"6px 8px",textAlign:"right",color:C.t1,fontFamily:C.mono,borderBottom:`1px solid ${C.border}15`}}>{bm[m]!=null?fmtD(bm[m]):"—"}</td>)}</tr>)}</tbody>
    </table></div>}
  </div>;
}

/* ═══ PAGE: UPLOAD ═══ */
function UploadPage({onFile,onDemo,mediaFile,afFile,valResult,rawRows,onGoMapping,onAfFile}){const ref1=useRef(),ref2=useRef();const[drag,setDrag]=useState(false);
  return<div><h1 style={{fontSize:24,fontWeight:700,color:C.t0,marginBottom:4}}>Data Upload</h1><p style={{color:C.t2,marginBottom:28,fontSize:14}}>Upload media performance data and AppsFlyer attribution data.</p>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:20,marginBottom:24}}><div onDragOver={e=>{e.preventDefault();setDrag(true)}} onDragLeave={()=>setDrag(false)} onDrop={e=>{e.preventDefault();setDrag(false);onFile(e.dataTransfer.files[0])}} onClick={()=>ref1.current?.click()} style={{border:`2px dashed ${drag?C.p:C.border}`,borderRadius:14,padding:"48px 32px",textAlign:"center",cursor:"pointer",background:drag?C.pBg:C.bg2}}><Upload size={28} color={C.p} style={{margin:"0 auto 12px",display:"block"}}/><div style={{fontWeight:600,fontSize:15,color:C.t0,marginBottom:6}}>Media Platform Data</div><div style={{fontSize:12,color:C.t2,marginBottom:12}}>DV360, Meta, Google, Snapchat, any platform</div><div style={{display:"inline-block",background:C.p,color:"#fff",borderRadius:6,padding:"6px 16px",fontSize:12,fontWeight:600}}>CSV or Excel</div>{mediaFile&&<div style={{marginTop:12,fontSize:12,color:C.g,display:"flex",alignItems:"center",gap:6,justifyContent:"center"}}><CheckCircle2 size={14}/>{mediaFile}</div>}<input ref={ref1} type="file" accept=".csv,.xlsx,.xls,.tsv" style={{display:"none"}} onChange={e=>{onFile(e.target.files[0]);e.target.value="";}}/></div>
      <div onClick={()=>ref2.current?.click()} style={{border:`2px dashed ${C.border}`,borderRadius:14,padding:"48px 32px",textAlign:"center",cursor:"pointer",background:C.bg2}}><Upload size={28} color={C.cy} style={{margin:"0 auto 12px",display:"block"}}/><div style={{fontWeight:600,fontSize:15,color:C.t0,marginBottom:6}}>AppsFlyer Data</div><div style={{fontSize:12,color:C.t2,marginBottom:12}}>Attribution, installs, events</div><div style={{display:"inline-block",background:C.cy+"30",color:C.cy,borderRadius:6,padding:"6px 16px",fontSize:12,fontWeight:600}}>CSV or Excel</div>{afFile&&<div style={{marginTop:12,fontSize:12,color:C.g,display:"flex",alignItems:"center",gap:6,justifyContent:"center"}}><CheckCircle2 size={14}/>{afFile}</div>}<input ref={ref2} type="file" accept=".csv,.xlsx,.xls,.tsv" style={{display:"none"}} onChange={e=>{onAfFile(e.target.files[0]);e.target.value="";}}/></div></div>
    <div style={{textAlign:"center",marginBottom:24}}><button onClick={onDemo} style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:8,padding:"10px 28px",color:C.t0,cursor:"pointer",fontFamily:C.font,fontSize:13,fontWeight:500,display:"inline-flex",alignItems:"center",gap:8}}><Database size={14}/>Load Demo Data</button></div>
    {valResult&&<><AlertBox type="error" items={valResult.errs} title="Fix These"/><AlertBox type="warn" items={valResult.warns} title="Auto-Fixed"/></>}
    {rawRows&&rawRows.length>0&&(()=>{const hdrs=Object.keys(rawRows[0]).slice(0,8);return<div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:20,marginTop:16}}><div style={{display:"flex",alignItems:"center",gap:8,marginBottom:12}}><Eye size={14} color={C.pL}/><span style={{fontWeight:600,fontSize:13,color:C.t0}}>Preview ({rawRows.length} rows)</span></div><div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}><thead><tr>{hdrs.map(h=><th key={h} style={{padding:"8px 12px",textAlign:"left",fontSize:10,fontWeight:600,color:C.t2,borderBottom:`1px solid ${C.border}`,whiteSpace:"nowrap",textTransform:"uppercase"}}>{h.length>20?h.slice(0,20)+"…":h}</th>)}</tr></thead><tbody>{rawRows.slice(0,4).map((r,i)=><tr key={i} style={{background:i%2?C.bg1+"80":"transparent"}}>{hdrs.map(h=><td key={h} style={{padding:"7px 12px",color:C.t1,borderBottom:`1px solid ${C.border}15`,whiteSpace:"nowrap",maxWidth:150,overflow:"hidden",textOverflow:"ellipsis"}}>{String(r[h]??"").slice(0,25)}</td>)}</tr>)}</tbody></table></div><div style={{textAlign:"right",marginTop:14}}><button onClick={onGoMapping} style={{background:C.p,color:"#fff",border:"none",borderRadius:8,padding:"10px 24px",cursor:"pointer",fontFamily:C.font,fontSize:14,fontWeight:600,display:"inline-flex",alignItems:"center",gap:8}}>Continue to Mapping <ChevronRight size={16}/></button></div></div>;})()}
  </div>;
}

/* ═══ PAGE: MAPPING (with Objectives + Benchmarks) ═══ */
function MappingPage({headers,colMap,setColMap,onLaunch,valResult,objectives,setObjectives,benchmarks,setBenchmarks}){
  const sel={background:C.bg1,border:`1px solid ${C.border}`,borderRadius:8,padding:"9px 12px",color:C.t0,fontSize:13,fontFamily:C.font,outline:"none",width:"100%",cursor:"pointer"};
  const allM=["Date","Campaign","Platform","Spend","Impressions","Clicks","Installs","Complete Views","Creative","Advertiser","Sessions","ATC_1","ATC_2","FO_1","FO_2","P_1","P_2"];
  const labels={ATC_1:"Add to Cart (Instamart)",ATC_2:"Add to Cart (IM)",FO_1:"First Order (Instamart)",FO_2:"First Order (IM)",P_1:"Purchase (Instamart)",P_2:"Purchase (IM)"};
  const toggleObj=id=>setObjectives(prev=>prev.includes(id)?prev.filter(x=>x!==id):prev.length<2?[...prev,id]:prev);
  return<div><h1 style={{fontSize:24,fontWeight:700,color:C.t0,marginBottom:4}}>Column Mapping &amp; Configuration</h1><p style={{color:C.t2,marginBottom:24,fontSize:14}}>Map columns, set objectives, and load benchmarks.</p>

    {/* Objectives */}
    <div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:20,marginBottom:20}}>
      <div style={{fontWeight:600,fontSize:14,color:C.t0,marginBottom:4,display:"flex",alignItems:"center",gap:8}}><Target size={16} color={C.a}/>Campaign Objectives</div>
      <div style={{fontSize:12,color:C.t2,marginBottom:14}}>Select 1–2 objectives. Dashboard will prioritize relevant metrics.</div>
      <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>{OBJECTIVES.map(o=>{const act=objectives.includes(o.id);return<button key={o.id} onClick={()=>toggleObj(o.id)} style={{background:act?C.p+"30":"transparent",color:act?C.pL:C.t1,border:`1px solid ${act?C.p:C.border}`,borderRadius:8,padding:"10px 18px",cursor:"pointer",fontFamily:C.font,fontSize:13,fontWeight:act?600:400,display:"flex",alignItems:"center",gap:6}}><span>{o.icon}</span>{o.label}{act&&<CheckCircle2 size={14} color={C.g}/>}</button>;})}</div>
    </div>

    {/* Benchmarks */}
    <BenchmarkInput benchmarks={benchmarks} setBenchmarks={setBenchmarks}/>

    {/* Column Mapping */}
    <div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:24,marginBottom:20}}>
      <div style={{fontWeight:600,fontSize:14,color:C.t0,marginBottom:16,display:"flex",alignItems:"center",gap:8}}><Columns3 size={16} color={C.p}/>Column Mapping</div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:16}}>{allM.map(m=><div key={m}><label style={{fontSize:11,color:REQ_COLS.includes(m)?C.pL:C.t2,marginBottom:4,display:"block",fontWeight:600,textTransform:"uppercase",letterSpacing:.5}}>{labels[m]||m}{REQ_COLS.includes(m)?" *":""}</label><div style={{position:"relative"}}><select value={colMap[m]||""} onChange={e=>setColMap(p=>({...p,[m]:e.target.value||undefined}))} style={{...sel,paddingRight:32}}><option value="">— Select —</option>{headers.map(h=><option key={h} value={h}>{h}</option>)}</select>{colMap[m]&&<CheckCircle2 size={14} color={C.g} style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",pointerEvents:"none"}}/>}</div></div>)}</div>
    </div>

    {valResult&&<><AlertBox type="error" items={valResult.errs} title="Fix These"/><AlertBox type="warn" items={valResult.warns} title="Auto-Fixed"/></>}
    <div style={{textAlign:"right"}}><button onClick={onLaunch} style={{background:C.p,color:"#fff",border:"none",borderRadius:8,padding:"12px 28px",cursor:"pointer",fontFamily:C.font,fontSize:14,fontWeight:600,display:"inline-flex",alignItems:"center",gap:8}}>Launch Dashboard <ChevronRight size={16}/></button></div>
  </div>;
}

/* ═══ DASHBOARD ═══ */
function DashboardPage({data,afData,objectives,benchmarks}){
  const allDates=useMemo(()=>[...new Set(data.map(r=>r.Date))].sort(),[data]);
  const[from,setFrom]=useState(allDates[0]||"");const[to,setTo]=useState(allDates[allDates.length-1]||"");
  const filtered=useMemo(()=>data.filter(r=>r.Date>=from&&r.Date<=to),[data,from,to]);
  const filteredAf=useMemo(()=>afData?afData.filter(r=>r.Date>=from&&r.Date<=to):null,[afData,from,to]);

  const spend=sumK(filtered,"Spend"),imp=sumK(filtered,"Impressions"),clk=sumK(filtered,"Clicks"),cv=sumK(filtered,"Complete Views");
  const cpm=imp?(spend/imp*1000):0,ctr=imp?(clk/imp*100):0,vtr=imp?(cv/imp*100):0,cpv=cv?(spend/cv):0,cpc=clk?(spend/clk):0;
  const afI=filteredAf?sumK(filteredAf,"Installs"):0,afS=filteredAf?sumK(filteredAf,"Sessions"):0,afATC=filteredAf?sumK(filteredAf,"Add to Cart"):0,afFO=filteredAf?sumK(filteredAf,"First Order"):0,afP=filteredAf?sumK(filteredAf,"Purchase"):0;
  const cps=afS?(spend/afS):0,cpi=afI?(spend/afI):0,cpfo=afFO?(spend/afFO):0;

  // Aggregate benchmarks across platforms weighted by spend
  const getBm=(metric)=>{if(!benchmarks)return null;const byPlat=aggByDim(filtered,"Platform");let total=0,wSum=0;byPlat.forEach(p=>{const bm=benchmarks[p.name];if(bm&&bm[metric]!=null){total+=p.Spend;wSum+=bm[metric]*p.Spend;}});return total>0?wSum/total:null;};

  const activeObj=OBJECTIVES.filter(o=>objectives.includes(o.id));
  const primaryMetrics=activeObj.length?[...new Set(activeObj.flatMap(o=>o.primary))]:null;

  const byDate=useMemo(()=>{const m={};filtered.forEach(r=>{if(!m[r.Date])m[r.Date]={Date:r.Date,Spend:0,"Complete Views":0};m[r.Date].Spend+=r.Spend||0;m[r.Date]["Complete Views"]+=r["Complete Views"]||0;});return Object.values(m).sort((a,b)=>a.Date.localeCompare(b.Date));},[filtered]);
  const byGeo=useMemo(()=>aggByDim(filtered,"Geo"),[filtered]);
  const byLang=useMemo(()=>aggByDim(filtered,"Language"),[filtered]);
  const byPlace=useMemo(()=>aggByDim(filtered,"Placement"),[filtered]);
  const byFormat=useMemo(()=>aggByDim(filtered,"AdFormat"),[filtered]);
  const byUser=useMemo(()=>aggByDim(filtered,"UserType"),[filtered]);
  const byPlat=useMemo(()=>aggByDim(filtered,"Platform"),[filtered]);
  const byCamp=useMemo(()=>aggByDim(filtered,"Campaign"),[filtered]);

  const isPrimary=m=>primaryMetrics?primaryMetrics.includes(m):true;

  return<div>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:24,flexWrap:"wrap",gap:12}}>
      <div><h1 style={{fontSize:24,fontWeight:700,color:C.t0,marginBottom:2}}>Performance Dashboard</h1><div style={{display:"flex",gap:8,alignItems:"center"}}><span style={{fontSize:12,color:C.t2}}>{filtered.length} records</span>{activeObj.map(o=><span key={o.id} style={{background:C.a+"20",color:C.a,borderRadius:4,padding:"2px 8px",fontSize:10,fontWeight:600}}>{o.icon} {o.label}</span>)}</div></div>
      <DateRangePicker allDates={allDates} from={from} to={to} onChange={(f,t)=>{setFrom(f);setTo(t);}}/>
    </div>

    {/* Objective highlight */}
    {primaryMetrics&&<div style={{background:C.a+"12",border:`1px solid ${C.a}30`,borderRadius:10,padding:"10px 16px",marginBottom:16,fontSize:12,color:C.a}}>
      <strong>Focus metrics:</strong> {primaryMetrics.join(", ")} — highlighted below based on your {activeObj.map(o=>o.label).join(" + ")} objective{activeObj.length>1?"s":""}.
    </div>}

    {/* KPI Row 1: Volume */}
    <div style={{fontSize:10,fontWeight:600,color:C.t2,textTransform:"uppercase",letterSpacing:1,marginBottom:6}}>Media Performance</div>
    <div style={{display:"flex",gap:10,marginBottom:10,flexWrap:"wrap"}}>
      <KpiCard label="Total Spend" value={fmtC(spend)} color={isPrimary("Spend")?C.p:C.border} benchmark={getBm("Spend")} actual={spend}/>
      <KpiCard label="Impressions" value={fmtN(imp)} color={isPrimary("Impressions")?C.b:C.border} benchmark={getBm("Impressions")} actual={imp}/>
      <KpiCard label="Clicks" value={fmtN(clk)} color={isPrimary("Clicks")?C.cy:C.border}/>
      <KpiCard label="Complete Views" value={fmtN(cv)} color={isPrimary("Complete Views")?C.g:C.border}/>
    </div>
    {/* KPI Row 2: Efficiency */}
    <div style={{display:"flex",gap:10,marginBottom:20,flexWrap:"wrap"}}>
      <KpiCard label="CPM" value={"₹"+fmtD(cpm)} sub="Cost/1000 Impr" color={isPrimary("CPM")?C.a:C.border} benchmark={getBm("CPM")} actual={cpm}/>
      <KpiCard label="CTR" value={fmtD(ctr,4)+"%"} sub="Clicks/Impr" color={isPrimary("CTR")?C.cy:C.border} benchmark={getBm("CTR")} actual={ctr}/>
      <KpiCard label="VTR" value={fmtD(vtr)+"%"} sub="Views/Impr" color={isPrimary("VTR")?C.g:C.border} benchmark={getBm("VTR")} actual={vtr}/>
      <KpiCard label="CPV" value={"₹"+fmtD(cpv,4)} sub="Cost/View" color={isPrimary("CPV")?C.r:C.border} benchmark={getBm("CPV")} actual={cpv}/>
      {clk>0&&<KpiCard label="CPC" value={"₹"+fmtD(cpc)} sub="Cost/Click" color={isPrimary("CPC")?C.b:C.border} benchmark={getBm("CPC")} actual={cpc}/>}
    </div>

    {/* AppsFlyer */}
    {filteredAf&&filteredAf.length>0&&<>
      <div style={{fontSize:10,fontWeight:600,color:C.cy,textTransform:"uppercase",letterSpacing:1,marginBottom:6}}>AppsFlyer Attribution</div>
      <div style={{display:"flex",gap:10,marginBottom:10,flexWrap:"wrap"}}>
        <KpiCard label="Installs" value={fmtN(afI)} color={isPrimary("Installs")?C.cy:C.border}/>
        <KpiCard label="Sessions" value={fmtN(afS)} color={isPrimary("Sessions")?C.g:C.border}/>
        <KpiCard label="Add to Cart" value={fmtN(afATC)} sub="Summed" color={C.border}/>
        <KpiCard label="First Orders" value={fmtN(afFO)} sub="Summed" color={C.border}/>
        <KpiCard label="Purchases" value={fmtN(afP)} sub="Summed" color={C.border}/>
      </div>
      <div style={{fontSize:10,fontWeight:600,color:"#f472b6",textTransform:"uppercase",letterSpacing:1,marginBottom:6}}>Blended Cost Metrics</div>
      <div style={{display:"flex",gap:10,marginBottom:24,flexWrap:"wrap"}}>
        <KpiCard label="CPS" value={"₹"+fmtD(cps)} sub="Cost/Session" color={isPrimary("CPS")?"#f472b6":C.border} benchmark={getBm("CPS")} actual={cps}/>
        <KpiCard label="CPI" value={"₹"+fmtD(cpi)} sub="Cost/Install" color={isPrimary("CPI")?C.r:C.border} benchmark={getBm("CPI")} actual={cpi}/>
        <KpiCard label="CPFO" value={"₹"+fmtD(cpfo)} sub="Cost/First Order" color={C.border}/>
        <KpiCard label="Install→Purchase" value={afI?fmtD(afP/afI*100)+"%":"—"} sub="CVR" color={C.border}/>
        <KpiCard label="Session→Purchase" value={afS?fmtD(afP/afS*100)+"%":"—"} sub="CVR" color={C.border}/>
      </div>
    </>}

    {/* Platform vs Benchmark Table */}
    {benchmarks&&Object.keys(benchmarks).length>0&&byPlat.length>0&&<div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:20,marginBottom:20}}>
      <div style={{fontWeight:600,fontSize:14,color:C.t0,marginBottom:16,display:"flex",alignItems:"center",gap:8}}><Target size={16} color={C.a}/>Actual vs Planned (by Platform)</div>
      <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse"}}><thead><tr>{["Platform","Spend","Impr","CPM","Plan CPM","Δ","VTR","Plan VTR","Δ","CPV","Plan CPV","Δ"].map(h=><th key={h} style={{padding:"8px 10px",textAlign:h==="Platform"?"left":"right",fontSize:9,fontWeight:600,color:C.t2,borderBottom:`1px solid ${C.border}`,textTransform:"uppercase",whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
        <tbody>{byPlat.map((p,i)=>{const bm=benchmarks[p.name]||{};const dCPM=bm.CPM?((p.CPM-bm.CPM)/bm.CPM*100):null;const dVTR=bm.VTR?((p.VTR-bm.VTR)/bm.VTR*100):null;const dCPV=bm.CPV?((p.CPV-bm.CPV)/bm.CPV*100):null;
          return<tr key={p.name} style={{background:i%2?C.bg1+"60":"transparent"}}><td style={{padding:"8px 10px",fontSize:11,color:C.t0,borderBottom:`1px solid ${C.border}15`}}>{p.name}</td>
            <td style={{padding:"8px 10px",fontSize:11,color:C.t1,textAlign:"right",fontFamily:C.mono,borderBottom:`1px solid ${C.border}15`}}>{fmtC(p.Spend)}</td>
            <td style={{padding:"8px 10px",fontSize:11,color:C.t1,textAlign:"right",fontFamily:C.mono,borderBottom:`1px solid ${C.border}15`}}>{fmtN(p.Impressions)}</td>
            <td style={{padding:"8px 10px",fontSize:11,color:C.t0,textAlign:"right",fontFamily:C.mono,borderBottom:`1px solid ${C.border}15`}}>₹{fmtD(p.CPM)}</td>
            <td style={{padding:"8px 10px",fontSize:11,color:C.t2,textAlign:"right",fontFamily:C.mono,borderBottom:`1px solid ${C.border}15`}}>{bm.CPM?`₹${fmtD(bm.CPM)}`:"—"}</td>
            <td style={{padding:"8px 10px",fontSize:11,color:dCPM!=null?(dCPM<0?C.g:C.r):C.t2,textAlign:"right",fontWeight:600,borderBottom:`1px solid ${C.border}15`}}>{dCPM!=null?`${dCPM>0?"+":""}${fmtD(dCPM,1)}%`:"—"}</td>
            <td style={{padding:"8px 10px",fontSize:11,color:C.t0,textAlign:"right",fontFamily:C.mono,borderBottom:`1px solid ${C.border}15`}}>{fmtD(p.VTR)}%</td>
            <td style={{padding:"8px 10px",fontSize:11,color:C.t2,textAlign:"right",fontFamily:C.mono,borderBottom:`1px solid ${C.border}15`}}>{bm.VTR?`${fmtD(bm.VTR)}%`:"—"}</td>
            <td style={{padding:"8px 10px",fontSize:11,color:dVTR!=null?(dVTR>0?C.g:C.r):C.t2,textAlign:"right",fontWeight:600,borderBottom:`1px solid ${C.border}15`}}>{dVTR!=null?`${dVTR>0?"+":""}${fmtD(dVTR,1)}%`:"—"}</td>
            <td style={{padding:"8px 10px",fontSize:11,color:C.t0,textAlign:"right",fontFamily:C.mono,borderBottom:`1px solid ${C.border}15`}}>₹{fmtD(p.CPV,4)}</td>
            <td style={{padding:"8px 10px",fontSize:11,color:C.t2,textAlign:"right",fontFamily:C.mono,borderBottom:`1px solid ${C.border}15`}}>{bm.CPV?`₹${fmtD(bm.CPV,4)}`:"—"}</td>
            <td style={{padding:"8px 10px",fontSize:11,color:dCPV!=null?(dCPV<0?C.g:C.r):C.t2,textAlign:"right",fontWeight:600,borderBottom:`1px solid ${C.border}15`}}>{dCPV!=null?`${dCPV>0?"+":""}${fmtD(dCPV,1)}%`:"—"}</td>
          </tr>})}</tbody></table></div>
    </div>}

    {/* Trend */}
    <div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:20,marginBottom:20}}>
      <div style={{fontWeight:600,fontSize:14,color:C.t0,marginBottom:16}}>Spend &amp; Complete Views Trend</div>
      {byDate.length>1?<ResponsiveContainer width="100%" height={220}><LineChart data={byDate}><CartesianGrid stroke={C.border} strokeDasharray="3 3"/><XAxis dataKey="Date" tick={{fill:C.t2,fontSize:10}} tickFormatter={d=>d.slice(5)}/><YAxis yAxisId="l" tick={{fill:C.t2,fontSize:10}} tickFormatter={fmtN}/><YAxis yAxisId="r" orientation="right" tick={{fill:C.t2,fontSize:10}} tickFormatter={fmtN}/><Tooltip contentStyle={tt}/><Legend wrapperStyle={{fontSize:11}}/><Line yAxisId="l" type="monotone" dataKey="Spend" stroke={C.p} strokeWidth={2} dot={false}/><Line yAxisId="r" type="monotone" dataKey="Complete Views" stroke={C.g} strokeWidth={2} dot={false}/></LineChart></ResponsiveContainer>:<div style={{color:C.t2,textAlign:"center",padding:40}}>Need 2+ dates</div>}
    </div>

    {/* Dimension Breakdowns */}
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:20}}>
      <DimBreakdown data={byGeo} title="Spend by Geo" metricKey="Spend" formatter={fmtC}/>
      <DimBreakdown data={byLang} title="Spend by Language" metricKey="Spend" formatter={fmtC}/>
      <DimBreakdown data={byPlace} title="VTR by Placement" metricKey="VTR" formatter={v=>fmtD(v)+"%"}/>
      <DimBreakdown data={byFormat} title="VTR by Ad Format" metricKey="VTR" formatter={v=>fmtD(v)+"%"}/>
      <DimBreakdown data={byUser} title="Spend by User Type" metricKey="Spend" formatter={fmtC}/>
      <DimBreakdown data={byPlat} title="Spend by Platform" metricKey="Spend" formatter={fmtC}/>
    </div>

    {/* Campaign Table */}
    <div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:20}}>
      <div style={{fontWeight:600,fontSize:14,color:C.t0,marginBottom:16}}>Campaign Breakdown</div>
      <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse"}}><thead><tr>{["Campaign","Spend","Impr","Clicks","Views","CPM","CTR","VTR","CPV","CPC"].map(h=><th key={h} style={{padding:"10px 10px",textAlign:h==="Campaign"?"left":"right",fontSize:9,fontWeight:600,color:C.t2,textTransform:"uppercase",borderBottom:`1px solid ${C.border}`,whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
        <tbody>{byCamp.map((r,i)=><tr key={i} style={{background:i%2?C.bg1+"60":"transparent"}}><td style={{padding:"8px 10px",fontSize:11,color:C.t0,borderBottom:`1px solid ${C.border}15`,maxWidth:240,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={r.name}>{r.name.length>50?r.name.slice(0,50)+"…":r.name}</td>
          <td style={{padding:"8px 10px",fontSize:11,color:C.t0,textAlign:"right",fontFamily:C.mono,borderBottom:`1px solid ${C.border}15`}}>{fmtC(r.Spend)}</td>
          <td style={{padding:"8px 10px",fontSize:11,color:C.t1,textAlign:"right",fontFamily:C.mono,borderBottom:`1px solid ${C.border}15`}}>{fmtN(r.Impressions)}</td>
          <td style={{padding:"8px 10px",fontSize:11,color:C.t1,textAlign:"right",fontFamily:C.mono,borderBottom:`1px solid ${C.border}15`}}>{fmtN(r.Clicks)}</td>
          <td style={{padding:"8px 10px",fontSize:11,color:C.t0,textAlign:"right",fontFamily:C.mono,borderBottom:`1px solid ${C.border}15`}}>{fmtN(r["Complete Views"])}</td>
          <td style={{padding:"8px 10px",fontSize:11,color:C.t1,textAlign:"right",fontFamily:C.mono,borderBottom:`1px solid ${C.border}15`}}>₹{fmtD(r.CPM)}</td>
          <td style={{padding:"8px 10px",fontSize:11,color:C.t1,textAlign:"right",fontFamily:C.mono,borderBottom:`1px solid ${C.border}15`}}>{fmtD(r.CTR,4)}%</td>
          <td style={{padding:"8px 10px",fontSize:11,textAlign:"right",fontFamily:C.mono,borderBottom:`1px solid ${C.border}15`,color:r.VTR>80?C.g:r.VTR>50?C.a:C.r}}>{fmtD(r.VTR)}%</td>
          <td style={{padding:"8px 10px",fontSize:11,color:C.t1,textAlign:"right",fontFamily:C.mono,borderBottom:`1px solid ${C.border}15`}}>₹{fmtD(r.CPV,4)}</td>
          <td style={{padding:"8px 10px",fontSize:11,color:C.t1,textAlign:"right",fontFamily:C.mono,borderBottom:`1px solid ${C.border}15`}}>₹{fmtD(r.CPC)}</td>
        </tr>)}</tbody></table></div>
    </div>
  </div>;
}

/* ═══ ANALYTICS ═══ */
function AnalyticsPage({data}){const[dims,setDims]=useState(["Geo","Placement"]);const[metrics,setMetrics]=useState(["Spend","Complete Views","VTR","CPM","CTR"]);const allDims=["Campaign","Platform","Geo","Language","Placement","AdFormat","AssetType","UserType","Creative","Date"].filter(d=>data.some(r=>r[d]&&r[d]!=="Other"&&r[d]!=="Unknown"));const allMetrics=["Spend","Impressions","Clicks","Complete Views","CPM","CTR","VTR","CPV","CPC","Installs","Sessions"];const toggle=(a,s,v)=>s(a.includes(v)?a.filter(x=>x!==v):[...a,v]);
  const pivot=useMemo(()=>{if(!dims.length)return[];const m={};data.forEach(r=>{const k=dims.map(d=>r[d]||"—").join(" | ");if(!m[k])m[k]={_key:k,Spend:0,Impressions:0,Clicks:0,"Complete Views":0,Sessions:0,Installs:0};const g=m[k];g.Spend+=r.Spend||0;g.Impressions+=r.Impressions||0;g.Clicks+=r.Clicks||0;g["Complete Views"]+=r["Complete Views"]||0;g.Sessions+=r.Sessions||0;g.Installs+=r.Installs||0;});return Object.values(m).map(g=>({...g,CPM:g.Impressions?(g.Spend/g.Impressions*1000):0,CTR:g.Impressions?(g.Clicks/g.Impressions*100):0,VTR:g.Impressions?(g["Complete Views"]/g.Impressions*100):0,CPV:g["Complete Views"]?(g.Spend/g["Complete Views"]):0,CPC:g.Clicks?(g.Spend/g.Clicks):0})).sort((a,b)=>b.Spend-a.Spend);},[data,dims]);
  const fc=(v,m)=>{if(m==="Spend")return fmtC(v);if(["Impressions","Clicks","Complete Views","Sessions","Installs"].includes(m))return fmtN(v);if(m==="CTR")return fmtD(v,4)+"%";if(m==="VTR")return fmtD(v)+"%";if(["CPM","CPC"].includes(m))return v?"₹"+fmtD(v):"—";if(m==="CPV")return v?"₹"+fmtD(v,4):"—";return String(v);};
  return<div><h1 style={{fontSize:24,fontWeight:700,color:C.t0,marginBottom:24}}>Analytics Builder</h1><div style={{display:"grid",gridTemplateColumns:"220px 1fr",gap:20}}><div><div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:20,marginBottom:16}}><div style={{fontWeight:600,fontSize:13,color:C.t0,marginBottom:12}}>Dimensions</div>{allDims.map(d=><button key={d} onClick={()=>toggle(dims,setDims,d)} style={{display:"block",width:"100%",marginBottom:6,background:dims.includes(d)?C.p+"30":"transparent",color:dims.includes(d)?C.pL:C.t1,border:`1px solid ${dims.includes(d)?C.p:C.border}`,borderRadius:8,padding:"8px 14px",cursor:"pointer",fontSize:12,fontFamily:C.font,fontWeight:500,textAlign:"left"}}>{d}</button>)}</div><div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:20}}><div style={{fontWeight:600,fontSize:13,color:C.t0,marginBottom:12}}>Metrics</div>{allMetrics.map(m=><button key={m} onClick={()=>toggle(metrics,setMetrics,m)} style={{display:"block",width:"100%",marginBottom:6,background:metrics.includes(m)?C.g+"30":"transparent",color:metrics.includes(m)?C.g:C.t1,border:`1px solid ${metrics.includes(m)?C.g+"60":C.border}`,borderRadius:8,padding:"8px 14px",cursor:"pointer",fontSize:12,fontFamily:C.font,fontWeight:500,textAlign:"left"}}>{m}</button>)}</div></div>
    <div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:20}}><div style={{display:"flex",justifyContent:"space-between",marginBottom:16}}><span style={{fontWeight:600,fontSize:14,color:C.t0}}>Results</span><span style={{fontSize:11,color:C.t2}}>{pivot.length} rows</span></div>{!dims.length?<div style={{color:C.t2,textAlign:"center",padding:60}}>Select a dimension</div>:<div style={{overflowX:"auto",maxHeight:520,overflowY:"auto"}}><table style={{width:"100%",borderCollapse:"collapse"}}><thead style={{position:"sticky",top:0,background:C.bg2,zIndex:2}}><tr><th style={{padding:"10px 14px",textAlign:"left",fontSize:11,fontWeight:600,color:C.t2,borderBottom:`1px solid ${C.border}`,textTransform:"uppercase"}}>{dims.join(" / ")}</th>{metrics.map(m=><th key={m} style={{padding:"10px 14px",textAlign:"right",fontSize:11,fontWeight:600,color:C.t2,borderBottom:`1px solid ${C.border}`,textTransform:"uppercase",whiteSpace:"nowrap"}}>{m}</th>)}</tr></thead><tbody>{pivot.map((r,i)=><tr key={i} style={{background:i%2?C.bg1+"60":"transparent"}}><td style={{padding:"10px 14px",fontSize:12,color:C.t0,borderBottom:`1px solid ${C.border}15`,whiteSpace:"nowrap",maxWidth:280,overflow:"hidden",textOverflow:"ellipsis"}} title={r._key}>{r._key}</td>{metrics.map(m=><td key={m} style={{padding:"10px 14px",fontSize:12,color:C.t1,textAlign:"right",fontFamily:C.mono,borderBottom:`1px solid ${C.border}15`,whiteSpace:"nowrap"}}>{fc(r[m],m)}</td>)}</tr>)}</tbody></table></div>}</div></div></div>;
}

/* ═══ INSIGHTS ═══ */
function InsightsPage({data,afData}){const insights=useMemo(()=>{const res=[];const ts=sumK(data,"Spend")||1,tcv=sumK(data,"Complete Views")||1,ti=sumK(data,"Impressions")||1;const avgCPV=ts/tcv;const byGeo=aggByDim(data,"Geo").filter(g=>g.name!=="Other");if(byGeo.length>0){const best=byGeo.reduce((a,b)=>a.CPV<b.CPV?a:b);res.push({type:"success",title:`${best.name} — best CPV ₹${fmtD(best.CPV,4)}`,desc:`VTR ${fmtD(best.VTR)}%, CPM ₹${fmtD(best.CPM)}.`,action:`Scale ${best.name} budget 15-20%.`});}if(byGeo.length>1){const worst=byGeo.reduce((a,b)=>a.CPV>b.CPV?a:b);if(worst.CPV>avgCPV*1.3)res.push({type:"error",title:`${worst.name} — high CPV ₹${fmtD(worst.CPV,4)}`,desc:`${fmtD((worst.CPV-avgCPV)/avgCPV*100,0)}% above avg.`,action:`Reduce ${worst.name} 20-30%.`});}const byP=aggByDim(data,"Placement").filter(p=>p.name!=="Other");if(byP.length>=2){const best=byP.reduce((a,b)=>a.VTR>b.VTR?a:b);const worst=byP.reduce((a,b)=>a.VTR<b.VTR?a:b);res.push({type:"info",title:`${best.name}: ${fmtD(best.VTR)}% VTR vs ${worst.name}: ${fmtD(worst.VTR)}%`,desc:`CPV: ₹${fmtD(best.CPV,4)} vs ₹${fmtD(worst.CPV,4)}.`,action:`Prioritize ${best.name} for awareness.`});}const byL=aggByDim(data,"Language").filter(l=>l.name!=="Other");if(byL.length>=2){const best=byL.reduce((a,b)=>a.VTR>b.VTR?a:b);res.push({type:"info",title:`${best.name} leads with ${fmtD(best.VTR)}% VTR`,desc:`CPM ₹${fmtD(best.CPM)}, CPV ₹${fmtD(best.CPV,4)}.`,action:`Invest in more ${best.name} variants.`});}if(afData?.length>0){const inst=sumK(afData,"Installs"),sess=sumK(afData,"Sessions"),purch=sumK(afData,"Purchase");res.push({type:"success",title:`Blended CPS: ₹${fmtD(ts/sess)} | CPI: ₹${fmtD(ts/inst)}`,desc:`${fmtN(inst)} installs, ${fmtN(sess)} sessions, ${fmtN(purch)} purchases.`,action:`Optimize for downstream conversion.`});}if(!res.length)res.push({type:"info",title:"Data looks healthy",desc:"No major issues.",action:"Upload more data."});return res;},[data,afData]);
  const cm={success:C.g,warning:C.a,error:C.r,info:C.b};
  return<div><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:24}}><div><h1 style={{fontSize:24,fontWeight:700,color:C.t0,marginBottom:2}}>AI Insights Engine</h1><p style={{color:C.t2,fontSize:14}}>Automated performance diagnosis.</p></div><button style={{background:C.p,color:"#fff",border:"none",borderRadius:8,padding:"10px 24px",cursor:"pointer",fontFamily:C.font,fontSize:13,fontWeight:600,display:"flex",alignItems:"center",gap:8}}><Zap size={14}/>Deep AI Analysis</button></div><div style={{display:"flex",flexDirection:"column",gap:16}}>{insights.map((ins,i)=><div key={i} style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:20,borderLeft:`4px solid ${cm[ins.type]}`}}><div style={{display:"flex",alignItems:"center",gap:10,marginBottom:8}}><div style={{width:12,height:12,borderRadius:6,background:cm[ins.type]}}/><span style={{fontWeight:600,fontSize:14,color:C.t0}}>{ins.title}</span></div><p style={{fontSize:13,color:C.t1,marginBottom:12,paddingLeft:22,lineHeight:1.5}}>{ins.desc}</p><div style={{background:C.bg1,borderRadius:8,padding:"10px 16px",marginLeft:22,fontSize:12,color:C.pL}}>→ {ins.action}</div></div>)}</div></div>;
}

/* ═══ REPORTS ═══ */
function ReportsPage({data,afData}){const[tab,setTab]=useState("completion");const[copied,setCopied]=useState(false);const s=sumK(data,"Spend"),imp=sumK(data,"Impressions"),clk=sumK(data,"Clicks"),cv=sumK(data,"Complete Views");const cpm=imp?(s/imp*1000):0,ctr=imp?(clk/imp*100):0,vtr=imp?(cv/imp*100):0,cpv=cv?(s/cv):0,cpc=clk?(s/clk):0;const dates=[...new Set(data.map(r=>r.Date))].sort();const dr=dates.length?`${dates[0]} to ${dates[dates.length-1]}`:"—";const afI=afData?sumK(afData,"Installs"):0,afS=afData?sumK(afData,"Sessions"):0,afP=afData?sumK(afData,"Purchase"):0,afFO=afData?sumK(afData,"First Order"):0;const byGeo=aggByDim(data,"Geo");
  const reports={completion:`📊 CAMPAIGN REPORT — ${dr}\n\nMEDIA: Spend ${fmtC(s)} | Impr ${fmtN(imp)} | Clicks ${fmtN(clk)} | Views ${fmtN(cv)}\nCPM ₹${fmtD(cpm)} | CTR ${fmtD(ctr,4)}% | VTR ${fmtD(vtr)}% | CPV ₹${fmtD(cpv,4)} | CPC ₹${fmtD(cpc)}${afI?`\n\nATTRIBUTION: Installs ${fmtN(afI)} | Sessions ${fmtN(afS)} | First Orders ${fmtN(afFO)} | Purchases ${fmtN(afP)}\nCPS ₹${fmtD(s/afS)} | CPI ₹${fmtD(s/afI)} | CPFO ₹${fmtD(s/afFO)}`:""}\n\nGEO:\n${byGeo.slice(0,6).map(g=>`  ${g.name.padEnd(22)} ${fmtC(g.Spend).padEnd(10)} VTR ${fmtD(g.VTR)}% CPV ₹${fmtD(g.CPV,4)}`).join("\n")}`,weekly:`📧 WEEKLY — ${dr}\n💰 ${fmtC(s)} | 👁 ${fmtN(imp)} | 🖱 ${fmtN(clk)} | ▶ ${fmtN(cv)}\n📈 CPM ₹${fmtD(cpm)} | CTR ${fmtD(ctr,4)}% | VTR ${fmtD(vtr)}% | CPV ₹${fmtD(cpv,4)}${afI?`\n📲 Installs ${fmtN(afI)} | Sessions ${fmtN(afS)} | CPS ₹${fmtD(s/afS)} | CPI ₹${fmtD(s/afI)}`:""}`,executive:`📋 EXECUTIVE — ${dr}\nBudget ${fmtC(s)} | Views ${fmtN(cv)} | VTR ${fmtD(vtr)}% | CPV ₹${fmtD(cpv,4)}${afI?` | CPI ₹${fmtD(s/afI)} | CPS ₹${fmtD(s/afS)}`:""}`};
  const cp=()=>{navigator.clipboard.writeText(reports[tab]).then(()=>{setCopied(true);setTimeout(()=>setCopied(false),2000);}).catch(()=>{const t=document.createElement("textarea");t.value=reports[tab];document.body.appendChild(t);t.select();document.execCommand("copy");document.body.removeChild(t);setCopied(true);setTimeout(()=>setCopied(false),2000);});};
  return<div><h1 style={{fontSize:24,fontWeight:700,color:C.t0,marginBottom:24}}>Report Generator</h1><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20,flexWrap:"wrap",gap:12}}><div style={{display:"flex",gap:8}}>{[{id:"completion",l:"Completion"},{id:"weekly",l:"Weekly"},{id:"executive",l:"Executive"}].map(t=><button key={t.id} onClick={()=>setTab(t.id)} style={{background:tab===t.id?C.p:"transparent",color:tab===t.id?"#fff":C.t1,border:`1px solid ${tab===t.id?C.p:C.border}`,borderRadius:8,padding:"8px 20px",cursor:"pointer",fontSize:13,fontFamily:C.font,fontWeight:500}}>{t.l}</button>)}</div><button onClick={cp} style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:8,padding:"8px 18px",cursor:"pointer",color:copied?C.g:C.t1,fontFamily:C.font,fontSize:13,display:"flex",alignItems:"center",gap:6}}>{copied?<CheckCircle2 size={14}/>:<Copy size={14}/>}{copied?"Copied!":"Copy"}</button></div><div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:12,padding:32}}><pre style={{fontFamily:C.mono,fontSize:13,color:C.t1,lineHeight:1.7,whiteSpace:"pre-wrap",margin:0}}>{reports[tab]}</pre></div></div>;
}

/* ═══ ROOT ═══ */
export default function App(){
  const[user,setUser]=useState(()=>{const sid=getSession();if(sid){const u=getUsers().find(u=>u.id===sid);return u||null;}return null;});
  const[showPwModal,setShowPwModal]=useState(false);const[pwMsg,setPwMsg]=useState("");
  const[tab,setTab]=useState("upload");const[rawRows,setRawRows]=useState(null);const[headers,setHeaders]=useState([]);
  const[colMap,setColMap]=useState({});const[data,setData]=useState(null);const[afData,setAfData]=useState(null);
  const[valResult,setValResult]=useState(null);const[mediaFile,setMediaFile]=useState("");const[afFile,setAfFile]=useState("");
  const[objectives,setObjectives]=useState(["video_views"]);const[benchmarks,setBenchmarks]=useState(null);
  const maxStep=!rawRows?1:!data?2:6;

  const handleFile=useCallback(async f=>{if(!f)return;setMediaFile(f.name);setData(null);setValResult(null);setRawRows(null);setHeaders([]);try{const{headers:h,rows}=await parseFile(f);setHeaders(h);setRawRows(rows);const am=autoMap(h);setColMap(am);const vr=validate(rows,am,"media");setValResult(vr.errs.length||vr.warns.length?vr:null);}catch(e){setValResult({errs:[e.message],warns:[],clean:[]});}},[]);
  const handleAfFile=useCallback(async f=>{if(!f)return;setAfFile(f.name);try{const{headers:h,rows}=await parseFile(f);const am=autoMap(h);const vr=validate(rows,am,"appsflyer");if(!vr.errs.length&&vr.clean.length)setAfData(vr.clean);}catch(e){console.warn("AF:",e.message);}},[]);
  const handleDemo=useCallback(()=>{const demo=[{Date:"2026/03/12","Insertion Order":"HM_SWIGGY_DV360_VIDEO_YT_RS9_NU_HINDI-CITIES(MH & DNCR ONLY)_AWA_INNS-CTV","YouTube Ad":"HM_SWIGGY_INNS-CTV_RS9_15SEC_HIN","Revenue (Adv Currency)":"22570.66",Impressions:"295536",Clicks:"116","Complete Views (Video)":"260507"},{Date:"2026/03/12","Insertion Order":"HM_SWIGGY_DV360_VIDEO_YT_RS9_NU_ENGLISH-GEOS_AWA_INNS-CTV","YouTube Ad":"HM_SWIGGY_INNS-CTV_RS9_15SEC_ENG","Revenue (Adv Currency)":"23990.66",Impressions:"279788",Clicks:"89","Complete Views (Video)":"245829"},{Date:"2026/03/13","Insertion Order":"HM_SWIGGY_DV360_VIDEO_YT_RS9_NU_TAMILNADU_AWA_VVC-CTV","YouTube Ad":"HM_SWIGGY_SHORTS_RS9_15SEC_TAM","Revenue (Adv Currency)":"1800",Impressions:"28000",Clicks:"45","Complete Views (Video)":"11200"},{Date:"2026/03/13","Insertion Order":"HM_SWIGGY_DV360_VIDEO_YT_RS9_NU_KERALA_AWA_INNS-CTV","YouTube Ad":"HM_SWIGGY_INNS-CTV_RS9_15SEC_MAL","Revenue (Adv Currency)":"3426.89",Impressions:"49457",Clicks:"16","Complete Views (Video)":"40380"},{Date:"2026/03/14","Insertion Order":"HM_SWIGGY_DV360_VIDEO_YT_RS9_NU_ANDHRA PRADESH/TELANGANA_AWA_VVC-CTV","YouTube Ad":"HM_SWIGGY_SKIP_RS9_15SEC_TEL","Revenue (Adv Currency)":"2967.68",Impressions:"49766",Clicks:"58","Complete Views (Video)":"27486"},{Date:"2026/03/14","Insertion Order":"HM_SWIGGY_DV360_VIDEO_YT_RS9_NU_HINDI-CITIES(EX-MH-DNCR)_AWA_INNS-CTV","YouTube Ad":"HM_SWIGGY_INNS-CTV_RS9_15SEC_HIN","Revenue (Adv Currency)":"27946.40",Impressions:"403885",Clicks:"184","Complete Views (Video)":"354564"}];const h=Object.keys(demo[0]);setHeaders(h);setRawRows(demo);const am=autoMap(h);setColMap(am);setMediaFile("demo-dv360.csv");const vr=validate(demo,am,"media");setData(vr.clean);setValResult(null);setTab("dashboard");},[]);
  const handleLaunch=useCallback(()=>{const vr=validate(rawRows,colMap,"media");if(vr.errs.length){setValResult(vr);return;}setData(vr.clean);setValResult(vr.warns.length?vr:null);setTab("dashboard");},[rawRows,colMap]);
  const handleLogout=()=>{clearSession();setUser(null);setTab("upload");setData(null);setRawRows(null);};

  // Show password change toast
  useEffect(()=>{if(pwMsg){const t=setTimeout(()=>setPwMsg(""),3000);return()=>clearTimeout(t);}},[pwMsg]);

  if(!user) return<LoginPage onLogin={setUser}/>;

  return<div style={{minHeight:"100vh",background:`linear-gradient(160deg,${C.bg0} 0%,#0d1222 40%,${C.bg0} 100%)`,color:C.t0,fontFamily:C.font,fontSize:14}}>
    <Sidebar active={tab} goTo={setTab} maxStep={maxStep} user={user} onLogout={handleLogout} onChangePw={()=>setShowPwModal(true)}/>
    {showPwModal&&<ChangePasswordModal user={user} onClose={(msg)=>{setShowPwModal(false);if(msg)setPwMsg(msg);}}/>}
    {pwMsg&&<div style={{position:"fixed",top:20,right:20,background:C.g+"20",border:`1px solid ${C.g}40`,borderRadius:8,padding:"10px 20px",color:C.g,fontSize:13,zIndex:300,fontFamily:C.font}}>{pwMsg}</div>}
    <div style={{marginLeft:210,padding:"32px 40px",minHeight:"100vh"}}>
      {tab==="upload"&&<UploadPage onFile={handleFile} onDemo={handleDemo} mediaFile={mediaFile} afFile={afFile} valResult={valResult} rawRows={rawRows} onGoMapping={()=>setTab("mapping")} onAfFile={handleAfFile}/>}
      {tab==="mapping"&&headers.length>0&&<MappingPage headers={headers} colMap={colMap} setColMap={setColMap} onLaunch={handleLaunch} valResult={valResult} objectives={objectives} setObjectives={setObjectives} benchmarks={benchmarks} setBenchmarks={setBenchmarks}/>}
      {tab==="dashboard"&&data&&<DashboardPage data={data} afData={afData} objectives={objectives} benchmarks={benchmarks}/>}
      {tab==="analytics"&&data&&<AnalyticsPage data={data}/>}
      {tab==="insights"&&data&&<InsightsPage data={data} afData={afData}/>}
      {tab==="reports"&&data&&<ReportsPage data={data} afData={afData}/>}
    </div></div>;
}
