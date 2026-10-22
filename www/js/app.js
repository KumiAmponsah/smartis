let sheetId='',selRC=null,selDD=null,selComp=null,calM=new Date().getMonth(),calY=new Date().getFullYear(),rTimer=null;
let D={rc:[],dd:[],comp:[],lv:[],wo:[],spares:[],kpis:[],techs:[],costs:[]};let selLV=null;
const TD=[{p:"Jan",a:88,m:285,r:30},{p:"Feb",a:85,m:268,r:36},{p:"Mar",a:90,m:300,r:28},{p:"Apr",a:83,m:278,r:34},{p:"May",a:86,m:290,r:31},{p:"Jun",a:88,m:298,r:33}];
const sC=s=>s==="DRILLING"||s==="RUNNING"||s==="AVAILABLE"||s==="ON ROAD"?"var(--success)":s==="BREAKDOWN"||s==="DOWN"?"var(--danger)":s==="PM DUE"||s==="IN WORKSHOP"?"var(--warning)":"var(--neutral)";
const hC=v=>v<=2?"var(--success)":v<=3?"var(--warning)":"var(--danger)";
const hL=v=>v===1?"EXCELLENT":v===2?"GOOD":v===3?"MONITOR":v===4?"DEGRADED":"CRITICAL";
const wTC=t=>t==="REACTIVE"?"var(--danger)":t==="PM"?"var(--success)":t==="PdM"?"var(--info)":"var(--warning)";
const pC=p=>p==="HIGH"?"var(--danger)":p==="MEDIUM"?"var(--warning)":"var(--neutral)";
const wSC=s=>s==="COMPLETED"?"var(--success)":s==="IN PROGRESS"?"var(--info)":"var(--warning)";
const kC=s=>s==="green"?"var(--success)":s==="red"?"var(--danger)":"var(--warning)";
const aTC=t=>"var(--muted)";
function fade(c,p){return `color-mix(in srgb, ${c} ${p}%, transparent)`;}
function pl(l,c){return `<span class="pill" style="background:${fade(c,14)};color:${c};border:1px solid ${fade(c,34)};">${l}</span>`;}
function $(i){return document.getElementById(i);}
function all(){return [...D.rc,...D.dd,...D.comp];} function allLV(){return D.lv||[];}
function typeOf(a){return D.rc.includes(a)?"RC":D.dd.includes(a)?"DD":D.comp.includes(a)?"COMPRESSOR":"LV";}
function updateClock(){const n=new Date();const c=$("clock"),d=$("date-disp"),e=$("exec-date"),s=$("shift-date");if(c)c.textContent=n.toLocaleTimeString("en-GB");if(d)d.textContent=n.toLocaleDateString("en-GB",{weekday:"short",day:"2-digit",month:"short",year:"numeric"});if(e)e.textContent=n.toLocaleDateString("en-GB",{day:"2-digit",month:"long",year:"numeric"});if(s)s.textContent="Date: "+n.toLocaleDateString("en-GB")+" | "+n.toLocaleTimeString("en-GB");}
updateClock();setInterval(updateClock,1000);
function drawSpark(si,li,f,col){const svg=$(si),lbl=$(li);if(!svg)return;const vals=TD.map(d=>d[f]);const mn=Math.min(...vals)-5,mx=Math.max(...vals)+5,W=260,H=48;const pts=vals.map((v,i)=>{const x=(i/(vals.length-1))*W,y=H-((v-mn)/(mx-mn))*H;return x+","+y;}).join(" ");const li2=vals.length-1,lx=(li2/(vals.length-1))*W,ly=H-((vals[li2]-mn)/(mx-mn))*H;svg.innerHTML=`<defs><linearGradient id="g${si}" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="${col}" stop-opacity=".3"/><stop offset="100%" stop-color="${col}" stop-opacity=".02"/></linearGradient></defs><polygon points="0,${H} ${pts} ${W},${H}" fill="url(#g${si})"/><polyline points="${pts}" fill="none" stroke="${col}" stroke-width="2.5" stroke-linejoin="round" stroke-linecap="round"/><circle cx="${lx}" cy="${ly}" r="4" fill="${col}" stroke="var(--bg)" stroke-width="2"/>`;const sf=f==="m"?" hrs":"%";if(lbl)lbl.innerHTML=TD.map((d,i)=>`<span style="font-size:8px;font-family:'JetBrains Mono',monospace;color:${i===TD.length-1?col:"var(--dim)"};">${d.p}<br>${d[f]}${sf}</span>`).join("");}
async function refreshData(){if(sheetId==="FILE"){loadStoredFile();return;}await refreshLive();}

const TAB_MAP=[["RC_Rigs","rc"],["DD_Rigs","dd"],["Compressors","comp"],["LightVehicles","lv"],["WorkOrders","wo"],["Spares","spares"],["KPITargets","kpis"],["Technicians","techs"]];
function upStatus(msg,col){const e=$("upload-status");if(e){e.innerHTML=msg;e.style.color=col||"var(--muted)";}}
function _normName(s){return String(s).toLowerCase().replace(/[^a-z0-9]/g,"");}
function _fmtDate(d){const m=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const x=new Date(Math.round(d.getTime()/86400000)*86400000);
  return String(x.getUTCDate()).padStart(2,"0")+" "+m[x.getUTCMonth()]+" "+x.getUTCFullYear();}
function _normRow(o){const out={};Object.keys(o).forEach(k=>{const key=String(k).trim();if(!key)return;let v=o[k];if(v instanceof Date)v=_fmtDate(v);else if(typeof v==="number"){/*keep*/}else v=String(v).trim();out[key]=v;});return out;}
function _sheetRows(ws){const rows=XLSX.utils.sheet_to_json(ws,{defval:"",raw:true});return rows.map(_normRow).filter(r=>Object.values(r).some(v=>v!==""&&v!=null));}
function _addWorkbook(bucket,wb,fileKey){wb.SheetNames.forEach(sn=>{const rows=_sheetRows(wb.Sheets[sn]);if(rows.length)bucket[fileKey||_normName(sn)]=rows;});return bucket;}
function _bucketToData(bucket){
  const nd={rc:[],dd:[],comp:[],lv:[],wo:[],spares:[],kpis:[],techs:[],costs:[]};
  const loaded=[],missing=[];
  TAB_MAP.forEach(([tab,key])=>{const rows=bucket[_normName(tab)];if(rows&&rows.length){nd[key]=rows;loaded.push(tab);}else missing.push(tab);});
  return {nd,loaded,missing};
}
function setSourceBadge(label,col){const dl=$("data-label"),ld=$("live-dot");if(dl){dl.textContent=label;dl.style.color=col;}if(ld)ld.style.background=col;}
async function handleFileUpload(files){
  if(!files||!files.length)return;
  if(typeof XLSX==="undefined"){upStatus("✕ Spreadsheet reader not loaded. Check your internet connection and reload the page.","var(--red)");return;}
  upStatus("🔄 Reading file...","var(--accent)");
  const bucket={};
  try{
    for(const f of files){
      const wb=XLSX.read(await f.arrayBuffer(),{type:"array",cellDates:true});
      _addWorkbook(bucket,wb,/\.csv$/i.test(f.name)?_normName(f.name.replace(/\.[^.]+$/,"")):null);
    }
  }catch(e){upStatus("✕ Could not read file: "+e.message,"var(--red)");return;}
  const {nd,loaded,missing}=_bucketToData(bucket);
  if(!loaded.length){upStatus("✕ No matching tabs found in that file. Tabs must be named exactly: "+TAB_MAP.map(t=>t[0]).join(", "),"var(--red)");return;}
  stopLive();D=nd;sheetId="FILE";
  try{localStorage.setItem("smartis-file-data",JSON.stringify(nd));localStorage.setItem("smartis-file-when",new Date().toISOString());}catch(e){}
  const dl=$("data-label"),ld=$("live-dot");if(dl){dl.textContent="FILE DATA";dl.style.color="var(--accent)";}if(ld)ld.style.background="var(--accent)";
  const lr=$("last-refresh");if(lr)lr.textContent=(missing.length?"⚠ "+loaded.length+"/8 tabs":"✓ All 8 tabs")+" loaded from file — "+new Date().toLocaleTimeString();
  if(missing.length)upStatus("✓ Loaded "+loaded.length+"/8 tabs. Missing/empty: "+missing.join(", ")+".","var(--amber)");
  hideSetup();renderAll();calcROI();
}
function loadStoredFile(){
  try{const fd=localStorage.getItem("smartis-file-data");if(!fd){upStatus("No previously uploaded file found on this device.","var(--amber)");return;}
    stopLive();D=JSON.parse(fd);sheetId="FILE";
    const dl=$("data-label"),ld=$("live-dot");if(dl){dl.textContent="FILE DATA";dl.style.color="var(--accent)";}if(ld)ld.style.background="var(--accent)";
    const lr=$("last-refresh");if(lr)lr.textContent="Restored last uploaded file — "+new Date().toLocaleTimeString();
    hideSetup();renderAll();calcROI();
  }catch(e){upStatus("Could not restore saved data: "+e.message,"var(--red)");}
}

const LIVE_EVERY=300000;
let liveLink=null,liveAt=0,liveBusy=false,connecting=false;
function sheetStatus(msg,col){const e=$("sheet-status");if(e){e.innerHTML=msg;e.style.color=col||"var(--muted)";}}
function parseSheetLink(v){
  v=String(v||"").trim();let m=v.match(/\/spreadsheets\/d\/e\/([\w-]+)/);
  if(m)return {raw:v,url:`https://docs.google.com/spreadsheets/d/e/${m[1]}/pub?output=xlsx`};
  m=v.match(/\/spreadsheets\/d\/([\w-]+)/)||v.match(/^([\w-]{25,})$/);
  return m?{raw:v,url:`https://docs.google.com/spreadsheets/d/${m[1]}/export?format=xlsx`}:null;
}
const LIVE_HELP="Check the link, and that the sheet is shared as <strong>Anyone with the link → Viewer</strong>.";
async function pullLive(link){
  if(typeof XLSX==="undefined")throw new Error("Spreadsheet reader not loaded. Check your internet connection and reload the page.");
  let r;
  try{r=await fetch(link.url,{cache:"no-store"});}
  catch(e){throw new Error(navigator.onLine===false?"You appear to be offline.":"Couldn't reach the sheet. "+LIVE_HELP);}
  if(!r.ok||/html/i.test(r.headers.get("content-type")||""))throw new Error((r.status===404?"Sheet not found. ":"Google wouldn't share this sheet. ")+LIVE_HELP);
  const res=_bucketToData(_addWorkbook({},XLSX.read(await r.arrayBuffer(),{type:"array",cellDates:true})));
  if(!res.loaded.length)throw new Error("Connected, but no tab names matched. Tabs must be named exactly: "+TAB_MAP.map(t=>t[0]).join(", "));
  return res;
}
function applyLive({nd,loaded,missing},when){
  D=nd;sheetId="LIVE";liveAt=when?new Date(when).getTime():Date.now();
  if(!when){try{localStorage.setItem("smartis-live-data",JSON.stringify(nd));localStorage.setItem("smartis-live-when",new Date(liveAt).toISOString());}catch(e){}}
  setSourceBadge("LIVE DATA","var(--green)");
  const t=new Date(liveAt).toLocaleTimeString(),lr=$("last-refresh");
  if(lr)lr.textContent=(missing.length?"⚠ "+loaded.length+"/8 tabs (missing: "+missing.join(", ")+")":"✓ Live sheet")+" — updated "+t;
  renderAll();
}
function startLiveTimer(){if(rTimer)clearInterval(rTimer);rTimer=setInterval(refreshData,LIVE_EVERY);}
function stopLive(){if(rTimer){clearInterval(rTimer);rTimer=null;}try{localStorage.setItem("smartis-source","OTHER");}catch(e){}}
async function connectSheet(){
  const link=parseSheetLink($("sheet-url-input").value);
  if(!link){sheetStatus("✕ That doesn't look like a Google Sheets link. Copy it from your browser's address bar while the sheet is open.","var(--red)");return;}
  if(connecting)return;connecting=true;
  sheetStatus("🔄 Connecting to Google Sheets...","var(--accent)");
  try{
    const res=await pullLive(link);
    liveLink=link;
    try{localStorage.setItem("smartis-sheet-link",link.raw);localStorage.setItem("smartis-source","LIVE");}catch(e){}
    sheetStatus(res.missing.length?"✓ Connected. Missing/empty tabs: "+res.missing.join(", ")+".":"✓ Connected. All 8 tabs found.",res.missing.length?"var(--amber)":"var(--green)");
    hideSetup();applyLive(res);calcROI();startLiveTimer();
  }catch(e){sheetStatus("✕ "+e.message,"var(--red)");}
  finally{connecting=false;}
}
async function refreshLive(){
  if(!liveLink||liveBusy)return;liveBusy=true;
  const lr=$("last-refresh");if(lr)lr.textContent="🔄 Refreshing from Google Sheets...";
  try{applyLive(await pullLive(liveLink));}
  catch(e){setSourceBadge("NOT UPDATING","var(--amber)");if(lr)lr.innerHTML="⚠ "+e.message+(liveAt?" Showing data from "+new Date(liveAt).toLocaleString("en-GB")+".":"");}
  finally{liveBusy=false;}
}
document.addEventListener("visibilitychange",()=>{if(!document.hidden&&sheetId==="LIVE"&&Date.now()-liveAt>60000)refreshLive();});
function resumeLive(){
  let saved=null,src=null;
  try{
    saved = localStorage.getItem("smartis-sheet-link") || localStorage.getItem("smartis-sid");
    src   = localStorage.getItem("smartis-source");
  }catch(e){}

  if(saved) $("sheet-url-input").value = saved;

  const link = parseSheetLink(saved);

  // No saved link at all → make sure the Load Data page is visible.
  if(!link || src==="OTHER"){
    // If a saved FILE exists, loadStoredFile() will be tried below by the caller.
    if(!localStorage.getItem("smartis-file-data")) showSetup();
    return;
  }

  liveLink = link;
  sheetId  = "LIVE";
  startLiveTimer();

  let cached=null, when=null;
  try{
    cached = JSON.parse(localStorage.getItem("smartis-live-data")||"null");
    when   = localStorage.getItem("smartis-live-when");
  }catch(e){}

  if(cached){
    hideSetup();
    applyLive({nd:cached, loaded:TAB_MAP.map(t=>t[0]), missing:[]}, when);
    calcROI();
    refreshLive();
    return;
  }

  sheetStatus("🔄 Connecting to your saved Google Sheet...","var(--accent)");
  pullLive(link)
    .then(res => { hideSetup(); applyLive(res); calcROI(); })
    .catch(e  => {
      // Could not pull → don't leave the user on a blank dashboard.
      showSetup();
      sheetStatus("✕ Couldn't load your saved sheet: " + e.message, "var(--red)");
    });
}
function showSetup(){$("setup-modal").classList.remove("hidden");$("main-app").classList.add("hidden");}
function hideSetup(){$("setup-modal").classList.add("hidden");$("main-app").classList.remove("hidden");}
function renderAll(){buildPills();buildExec();buildOverview();buildRC();buildDD();buildComp();buildLV();buildWOs();buildCal();buildKPIs();buildAnalytics();buildTechs();buildSpares();buildCostTracker();buildShift();calcROI();updateNotifCount();buildCascadeVisual();}
function buildPills(){const bd=all().filter(a=>a.Status==="BREAKDOWN"||a.Status==="DOWN").length;const po=all().filter(a=>Number(a.HrsToNextPM)<0).length;const so=D.spares.filter(s=>s.Status==="STOCK-OUT").length;const lvExp=(D.lv||[]).filter(v=>daysUntil(v.RoadworthyExpiry)<0||daysUntil(v.InsuranceExpiry)<0||daysUntil(v.FireExtExpiry)<0||daysUntil(v.FirstAidExpiry)<0).length;let h="";if(bd>0)h+=`<div class="apill" style="background:rgba(255,51,85,.12);border:1px solid rgba(255,51,85,.3);color:var(--red);"><div style="width:6px;height:6px;border-radius:50%;background:var(--red);animation:pulse 1.8s infinite;"></div>${bd} DOWN</div>`;if(po>0)h+=`<div class="apill" style="background:rgba(245,168,0,.12);border:1px solid rgba(245,168,0,.3);color:var(--amber);"><div style="width:6px;height:6px;border-radius:50%;background:var(--amber);animation:pulse 1.8s infinite;"></div>${po} PM OVERDUE</div>`;if(so>0)h+=`<div class="apill" style="background:rgba(255,51,85,.12);border:1px solid rgba(255,51,85,.3);color:var(--red);"><div style="width:6px;height:6px;border-radius:50%;background:var(--red);animation:pulse 1.8s infinite;"></div>${so} STOCK-OUT${so>1?"S":""}</div>`;if(lvExp>0)h+=`<div class="apill" style="background:rgba(255,51,85,.12);border:1px solid rgba(255,51,85,.3);color:var(--red);"><div style="width:6px;height:6px;border-radius:50%;background:var(--red);animation:pulse 1.8s infinite;"></div>${lvExp} LV COMPLIANCE EXPIRED</div>`;$("alert-pills").innerHTML=h;}
function sumCard(v,l,c){return `<div class="card-sm" style="text-align:center;border-top:2px solid ${c};"><div style="font-size:clamp(22px,3.4vw,30px);font-weight:800;color:var(--text);line-height:1.1;">${v}</div><div style="font-size:9px;color:var(--muted);letter-spacing:1.5px;text-transform:uppercase;font-family:'JetBrains Mono',monospace;margin-top:5px;">${l}</div></div>`;}function assetSummary(arr){const dr=arr.filter(a=>a.Status==="DRILLING"||a.Status==="RUNNING").length;const dn=arr.filter(a=>a.Status==="BREAKDOWN"||a.Status==="DOWN").length;const pm=arr.filter(a=>Number(a.HrsToNextPM)<0).length;const av=arr.length?Math.round(arr.reduce((s,a)=>s+Number(a.Availability),0)/arr.length):0;return sumCard(dr,"Operating","#00e087")+sumCard(dn,"Down","#ff3355")+sumCard(pm,"PM Overdue","#f5a800")+sumCard(av+"%","Avg Availability",av>=85?"#00e087":av>=75?"#f5a800":"#ff3355");}
function buildExec(){const a=all();const dr=D.rc.filter(r=>r.Status==="DRILLING").length+D.dd.filter(d=>d.Status==="DRILLING").length;const dn=a.filter(x=>x.Status==="BREAKDOWN"||x.Status==="DOWN").length;const av=a.length?Math.round(a.reduce((s,x)=>s+Number(x.Availability),0)/a.length):0;const ow=D.wo.filter(w=>w.Status!=="COMPLETED").length;const dc=D.wo.reduce((s,w)=>s+Number(w.DowntimeCost||0),0);$("exec-big").innerHTML=[{l:"Overall Fleet Availability",v:av+"%",c:av>=85?"#00e087":av>=75?"#f5a800":"#ff3355",s:`${D.rc.length} RC | ${D.dd.length} DD | ${D.comp.length} Compressors`},{l:"Assets Drilling / Running",v:dr,c:"#00e087",s:`${dn} asset(s) offline | ${a.length} total`},{l:"Open Work Orders",v:ow,c:ow>10?"#ff3355":ow>5?"#f5a800":"#00e087",s:`${D.wo.filter(w=>w.Priority==="HIGH"&&w.Status!=="COMPLETED").length} HIGH priority`},{l:"Downtime Cost (MTD)",v:"$"+dc.toLocaleString(),c:dc>150000?"#ff3355":dc>80000?"#f5a800":"#00e087",s:"All asset types combined"}].map(k=>`<div class="card-sm" style="--kpi-color:${k.c};text-align:center;"><div style="color:${k.c};">${k.v}</div><div style="font-size:10px;font-weight:700;margin-bottom:4px;letter-spacing:.5px;color:var(--text);">${k.l}</div><div style="font-size:9px;color:var(--muted);">${k.s}</div></div>`).join("");
let strip=`<div style="font-size:9px;color:var(--rc);font-family:'JetBrains Mono',monospace;font-weight:700;letter-spacing:2px;margin-bottom:8px;">RC DRILLING RIGS</div><div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:14px;">`;
strip+=D.rc.map(r=>`<div onclick="goA('rc','${r.ID}')" style="background:var(--surface2);border:1px solid ${fade(sC(r.Status),30)};border-radius:8px;padding:8px 12px;cursor:pointer;min-width:120px;"><div style="display:flex;align-items:center;gap:5px;margin-bottom:3px;"><div style="width:6px;height:6px;border-radius:50%;background:${sC(r.Status)};"></div><span style="font-size:9px;font-weight:700;color:var(--rc);font-family:'JetBrains Mono',monospace;">${r.ID}</span></div><div style="font-size:10px;font-weight:600;">${r.Name}</div><div style="font-size:16px;font-weight:900;color:${Number(r.Availability)>=90?"#00e087":Number(r.Availability)>=80?"#f5a800":"#ff3355"};">${r.Availability}%</div></div>`).join("");
strip+=`</div><div style="font-size:9px;color:var(--dd);font-family:'JetBrains Mono',monospace;font-weight:700;letter-spacing:2px;margin-bottom:8px;">DIAMOND DRILL RIGS</div><div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:14px;">`;
strip+=D.dd.map(d=>`<div onclick="goA('dd','${d.ID}')" style="background:var(--surface2);border:1px solid ${fade(sC(d.Status),30)};border-radius:8px;padding:8px 12px;cursor:pointer;min-width:120px;"><div style="display:flex;align-items:center;gap:5px;margin-bottom:3px;"><div style="width:6px;height:6px;border-radius:50%;background:${sC(d.Status)};"></div><span style="font-size:9px;font-weight:700;color:var(--dd);font-family:'JetBrains Mono',monospace;">${d.ID}</span></div><div style="font-size:10px;font-weight:600;">${d.Name}</div><div style="font-size:16px;font-weight:900;color:${Number(d.Availability)>=90?"#00e087":Number(d.Availability)>=80?"#f5a800":"#ff3355"};">${d.Availability}%</div>${d.CoreRecovery?`<div style="font-size:9px;color:var(--muted);">CR: ${d.CoreRecovery}%</div>`:""}</div>`).join("");
strip+=`</div><div style="font-size:9px;color:var(--comp);font-family:'JetBrains Mono',monospace;font-weight:700;letter-spacing:2px;margin-bottom:8px;">EXTERNAL COMPRESSORS</div><div style="display:flex;gap:8px;flex-wrap:wrap;">`;
strip+=D.comp.map(c=>`<div onclick="goA('comp','${c.ID}')" style="background:var(--surface2);border:1px solid ${fade(sC(c.Status),30)};border-radius:8px;padding:8px 12px;cursor:pointer;min-width:140px;"><div style="display:flex;align-items:center;gap:5px;margin-bottom:3px;"><div style="width:6px;height:6px;border-radius:50%;background:${sC(c.Status)};"></div><span style="font-size:9px;font-weight:700;color:var(--comp);font-family:'JetBrains Mono',monospace;">${c.ID}</span></div><div style="font-size:10px;font-weight:600;">${c.Name}</div><div style="font-size:9px;color:var(--muted);">→ ${c.AssignedTo}</div><div style="font-size:16px;font-weight:900;color:${Number(c.Availability)>=90?"#00e087":Number(c.Availability)>=80?"#f5a800":"#ff3355"};">${c.Availability}%</div></div>`).join("");
strip+=`</div>`;
strip+=`<div style="font-size:9px;color:var(--lv);font-family:'JetBrains Mono',monospace;font-weight:700;letter-spacing:2px;margin:14px 0 8px;">LIGHT VEHICLES</div><div style="display:flex;gap:8px;flex-wrap:wrap;">`;
strip+=(D.lv||[]).map(v=>`<div onclick="goA('lv','${v.ID}')" style="background:var(--surface2);border:1px solid ${lvStatusCol(v.Status)}44;border-radius:8px;padding:8px 12px;cursor:pointer;min-width:110px;"><div style="display:flex;align-items:center;gap:5px;margin-bottom:3px;"><div style="width:6px;height:6px;border-radius:50%;background:${lvStatusCol(v.Status)};"></div><span style="font-size:9px;font-weight:700;color:var(--lv);font-family:'JetBrains Mono',monospace;">${v.ID}</span></div><div style="font-size:10px;font-weight:600;">${v.Name}</div><div style="font-size:9px;color:var(--muted);">${v.Type}</div><div style="font-size:14px;font-weight:900;color:${Number(v.Availability)>=90?"#00e087":Number(v.Availability)>=80?"#f5a800":"#ff3355"};">${v.Availability}%</div></div>`).join("");
strip+=`</div>`;
$("exec-strip").innerHTML=strip;
const crit=[];a.filter(x=>x.Status==="BREAKDOWN"||x.Status==="DOWN").forEach(x=>crit.push({m:`${x.ID} ${x.Name} — OFFLINE. ${x.OpenWOs} open WO(s). Immediate action.`,c:"#ff3355"}));a.filter(x=>Number(x.HrsToNextPM)<0).forEach(x=>crit.push({m:`${x.Name} — PM overdue by ${Math.abs(x.HrsToNextPM)} hrs.`,c:"#f5a800"}));D.spares.filter(s=>s.Status==="STOCK-OUT").forEach(s=>crit.push({m:`STOCK-OUT: ${s.Name} (${s.PartNumber}). Class ${s.Class}. Lead: ${s.LeadTimeDays} day(s).`,c:"#ff3355"}));D.comp.filter(c=>c.Status==="DOWN").forEach(c=>{const rc=D.rc.find(r=>r.ConnectedCompressor===c.ID);if(rc)crit.push({m:`CASCADE: ${c.Name} DOWN → ${rc.Name} has no air supply. Both assets affected.`,c:"#ff3355"});});
$("exec-critical").innerHTML=crit.length?crit.map(i=>`<div style="display:flex;gap:8px;padding:8px 0;border-bottom:1px solid var(--border);"><div style="width:3px;background:${i.c};border-radius:2px;flex-shrink:0;min-height:32px;"></div><div style="font-size:11px;line-height:1.5;">${i.m}</div></div>`).join(""):'<div style="font-size:11px;color:var(--green);padding:14px 0;text-align:center;">✓ No critical issues</div>';
const wins=[];const done=D.wo.filter(w=>w.Status==="COMPLETED").length;if(done>0)wins.push(`${done} work orders completed this period`);[...D.rc,...D.dd].filter(x=>Number(x.Availability)>=95).forEach(x=>wins.push(`${x.Name} — ${x.Availability}% availability. Excellent.`));D.dd.filter(d=>Number(d.CoreRecovery)>=96).forEach(d=>wins.push(`${d.Name} — Core recovery ${d.CoreRecovery}%. Outstanding drilling performance.`));D.comp.filter(c=>Number(c.HrsToNextPM)>200).forEach(c=>wins.push(`${c.Name} — ${c.HrsToNextPM} hrs to next PM. Well ahead.`));
$("exec-wins").innerHTML=wins.length?wins.map(w=>`<div style="display:flex;gap:8px;padding:7px 0;border-bottom:1px solid var(--border);"><div style="color:var(--green);flex-shrink:0;">✓</div><div style="font-size:11px;">${w}</div></div>`).join(""):'<div style="font-size:11px;color:var(--muted);padding:14px 0;">No data yet</div>';
$("exec-recs").innerHTML=["Prioritise PM compliance — 76% vs 90% target. Schedule all overdue services immediately.","Compressor availability (79.2%) below target — increase PM frequency, focus on oil separator service.","Resolve all Class A spare stock-outs within 24 hours — bean pump seals and compressor separators are critical.","Plan major overhaul for DGL-DD04 RAPTOR — 8,900 hrs, all systems degraded, core recovery at 0%.","Implement dedicated compressor PM register — track compressor hours separately from RC rig hours.","Track core recovery % as a primary KPI for DD rigs — low core recovery = lost geological data and revenue."].map((r,i)=>`<div style="display:flex;gap:8px;padding:7px 0;border-bottom:1px solid var(--border);"><span style="font-size:9px;font-weight:700;color:var(--accent);font-family:'JetBrains Mono',monospace;flex-shrink:0;margin-top:1px;">${i+1}.</span><div style="font-size:11px;">${r}</div></div>`).join("");}
function buildOverview(){$("ov-kpis").innerHTML=D.kpis.slice(0,4).map(k=>`<div class="card-sm"><div class="kl">${k.Metric}</div><div class="kv" style="color:${kC(k.Status)};">${k.Value}</div><div class="kf"><span class="kt">TGT ${k.Target}</span><span class="ktr" style="color:${kC(k.Status)};">${k.Trend}</span></div></div>`).join("");$("ov-fleet").innerHTML=all().map(a=>{const t=typeOf(a);return `<div onclick="goA('${t.toLowerCase().substring(0,4)}','${a.ID}')" style="display:flex;align-items:center;justify-content:space-between;padding:8px 0;border-bottom:1px solid var(--border);cursor:pointer;"><div style="display:flex;align-items:center;gap:8px;"><div style="width:7px;height:7px;border-radius:50%;background:${sC(a.Status)};box-shadow:0 0 6px ${sC(a.Status)};"></div><div><div style="font-size:11px;font-weight:600;">${a.Name}</div><div style="font-size:9px;color:var(--muted);">${a.Site}</div></div></div><div style="display:flex;gap:5px;">${pl(t==="COMPRESSOR"?"COMP":t,aTC(t))} ${pl(a.Status,sC(a.Status))}</div></div>`;}).join("");const sm={};D.wo.filter(w=>w.Status!=="COMPLETED").forEach(w=>{sm[w.System]=(sm[w.System]||0)+Number(w.Hours||0);});const tot=Object.values(sm).reduce((s,v)=>s+v,0)||1;const sr=Object.entries(sm).sort((a,b)=>b[1]-a[1]).slice(0,6);const pc=["var(--info)","var(--info)","var(--info)","var(--info)","var(--info)","var(--info)"];$("ov-pareto").innerHTML=sr.length?sr.map(([s,h],i)=>`<div style="margin-bottom:9px;"><div style="display:flex;justify-content:space-between;margin-bottom:3px;"><span style="font-size:11px;">${s}</span><span style="font-size:9px;color:${pc[i]};font-family:'JetBrains Mono',monospace;font-weight:700;">${h}h (${Math.round(h/tot*100)}%)</span></div><div class="bt"><div class="bf" style="width:${Math.round(h/tot*100)}%;background:${pc[i]};"></div></div></div>`).join(""):'<div style="font-size:11px;color:var(--muted);padding:16px;text-align:center;">No open WOs</div>';const alerts=[];all().filter(a=>a.Status==="BREAKDOWN"||a.Status==="DOWN").forEach(a=>alerts.push({s:"CRITICAL",m:`${a.Name} — OFFLINE. ${a.OpenWOs} WO(s) open.`,c:"#ff3355"}));all().filter(a=>Number(a.HrsToNextPM)<0).forEach(a=>alerts.push({s:"HIGH",m:`${a.Name} — PM overdue ${Math.abs(a.HrsToNextPM)} hrs`,c:"#f5a800"}));D.spares.filter(s=>s.Status==="STOCK-OUT").forEach(s=>alerts.push({s:"STOCK-OUT",m:`${s.Name} (${s.PartNumber}) Class ${s.Class}`,c:"#ff3355"}));D.comp.filter(c=>c.Status==="DOWN").forEach(c=>{const rc=D.rc.find(r=>r.ConnectedCompressor===c.ID);if(rc)alerts.push({s:"CASCADE",m:`${c.Name} DOWN → ${rc.Name} has no air supply`,c:"#ff3355"});});$("ov-alerts").innerHTML=alerts.length?alerts.slice(0,8).map(a=>`<div class="ar"><div class="astr" style="background:${a.c};"></div><div><div style="font-size:9px;font-weight:700;letter-spacing:1.5px;font-family:'JetBrains Mono',monospace;color:${a.c};margin-bottom:2px;">${a.s}</div><div style="font-size:11px;line-height:1.4;">${a.m}</div></div></div>`).join(""):'<div style="font-size:11px;color:var(--green);padding:16px;text-align:center;">✓ No active alerts</div>';drawSpark("sp-avail","sp-avail-l","a","#22c55e");drawSpark("sp-mtbf","sp-mtbf-l","m","#5b8ef0");drawSpark("sp-react","sp-react-l","r","#ef4444");}
function goA(t,id){const tMap={rc:2,dd:3,comp:4,lv:5};showTab(t,document.querySelectorAll(".tab")[tMap[t]]);setTimeout(()=>{if(t==="rc")selRC=id;else if(t==="dd")selDD=id;else if(t==="comp")selComp=id;else selLV=id;if(t==="rc")renderRCDet();else if(t==="dd")renderDDDet();else if(t==="comp")renderCompDet();else renderLVDet();},80);}
function buildRC(){$("rc-summary").innerHTML=assetSummary(D.rc);$("rc-cards").innerHTML=D.rc.map(r=>`<div class="ac${selRC===r.ID?" sel":""}" id="rcc-${r.ID}" onclick="selRCf('${r.ID}')" style="border:1px solid ${fade(selRC===r.ID?"var(--rc)":sC(r.Status),30)};"><div style="display:flex;justify-content:space-between;margin-bottom:7px;"><span class="mono" style="font-size:9px;font-weight:700;color:var(--rc);">${r.ID}</span>${pl(r.Status,sC(r.Status))}</div><div style="font-size:12px;font-weight:700;margin-bottom:2px;">${r.Name}</div><div style="font-size:9px;color:var(--muted);margin-bottom:10px;">${r.Site}</div><div style="display:flex;justify-content:space-between;margin-bottom:8px;"><div><div style="font-size:20px;font-weight:900;color:${Number(r.Availability)>=90?"#00e087":Number(r.Availability)>=80?"#f5a800":"#ff3355"};">${r.Availability}%</div><div style="font-size:8px;color:var(--muted);font-family:'JetBrains Mono',monospace;">AVAIL</div></div><div style="text-align:right;"><div style="font-size:20px;font-weight:900;">${Number(r.Hours).toLocaleString()}</div><div style="font-size:8px;color:var(--muted);font-family:'JetBrains Mono',monospace;">HRS</div></div></div><div class="mono" style="font-size:8px;font-weight:700;color:${r.HrsToNextPM<0?"#ff3355":r.HrsToNextPM<100?"#f5a800":"#00e087"};">${r.HrsToNextPM<0?"⚠ PM OVERDUE "+Math.abs(r.HrsToNextPM)+"h":"NEXT PM: "+r.HrsToNextPM+"h"}</div><div class="mono" style="font-size:8px;color:var(--comp);margin-top:4px;">📡 COMP: ${r.ConnectedCompressor||"—"}</div>${r.MetresDrilled?`<div class="mono" style="font-size:8px;color:var(--muted);margin-top:2px;">${r.MetresDrilled}m / ${r.ShiftTarget}m</div>`:""}</div>`).join("");$("rc-comp-table").innerHTML=D.rc.map(r=>{const c=D.comp.find(x=>x.ID===r.ConnectedCompressor);return `<tr><td style="font-weight:700;">${r.Name}</td><td class="mono" style="color:var(--comp);">${r.ConnectedCompressor||"—"}</td><td style="color:var(--muted);">${c?c.Make+" "+c.Model:"Not linked"}</td><td>${c?pl(c.Status,sC(c.Status)):"—"}</td><td class="mono">${c?c.CFM+" CFM":"—"}</td><td class="mono">${c?c.WorkingPressure+" psi":"—"}</td><td style="font-size:10px;color:${c&&(c.Status==="DOWN")?"#ff3355":"var(--muted)"};">${c&&c.Status==="DOWN"?"⚠ NO AIR — CANNOT DRILL":"Normal supply"}</td></tr>`;}).join("");}
function selRCf(id){selRC=selRC===id?null:id;document.querySelectorAll("[id^='rcc-']").forEach(c=>c.classList.remove("sel"));if(selRC)document.getElementById("rcc-"+selRC)?.classList.add("sel");renderRCDet();}
function renderRCDet(){const el=$("rc-detail");if(!selRC){el.innerHTML="";return;}const r=D.rc.find(x=>x.ID===selRC);if(!r){el.innerHTML="";return;}const comp=D.comp.find(c=>c.ID===r.ConnectedCompressor);const sys=[["Engine",r.Engine],["Hydraulics",r.Hydraulics],["Track",r.Track],["Rotation Head",r.RotationHead],["RC Sample System",r.RCSampleSystem],["Electrical",r.Electrical]];const rWOs=D.wo.filter(w=>w.AssetID===r.ID);el.innerHTML=`<div class="dp fade-in" style="border:1px solid rgba(0,200,240,.25);"><div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:16px;"><div><div style="font-size:16px;font-weight:800;">${r.Name}</div><div style="font-size:11px;color:var(--muted);">${r.Site} | ${r.ID} | ${Number(r.Hours).toLocaleString()} hrs | $${r.CostPerHour}/hr</div></div>${pl(r.Status,sC(r.Status))}</div><div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:18px;"><div><div class="st"><div class="sb" style="background:var(--rc);"></div><span class="sl">RC Rig Systems</span></div>${sys.map(([n,v])=>`<div class="hr2"><div class="hs">${n}</div><div class="bt"><div class="bf" style="width:${(v/6)*100}%;background:${hC(v)};"></div></div>${pl(hL(v),hC(v))}</div>`).join("")}</div><div><div class="st"><div class="sb" style="background:var(--comp);"></div><span class="sl">Connected Compressor</span></div>${comp?`<div style="background:var(--surface2);border:1px solid ${fade(sC(comp.Status),30)};border-radius:8px;padding:12px;"><div style="font-size:12px;font-weight:700;margin-bottom:3px;">${comp.Name}</div><div style="font-size:9px;color:var(--muted);margin-bottom:8px;">${comp.Make} ${comp.Model} | ${comp.ID}</div>${[["Status",pl(comp.Status,sC(comp.Status))],["CFM",comp.CFM+" CFM"],["Pressure",comp.WorkingPressure+" psi"],["Hours",Number(comp.Hours).toLocaleString()+"h"],["Next PM",comp.HrsToNextPM<0?`<span style="color:#ff3355;">OVERDUE ${Math.abs(comp.HrsToNextPM)}h</span>`:comp.HrsToNextPM+"h"]].map(([l,v])=>`<div class="ir"><span style="font-size:10px;color:var(--muted);">${l}</span><span style="font-size:11px;font-weight:600;">${v}</span></div>`).join("")}${comp.Status==="DOWN"?'<div style="background:rgba(255,51,85,.1);border:1px solid rgba(255,51,85,.3);border-radius:6px;padding:8px;margin-top:8px;font-size:10px;color:#ff3355;font-weight:700;">⚠ COMPRESSOR DOWN — RC RIG HAS NO AIR SUPPLY. DRILLING NOT POSSIBLE.</div>':""}</div>`:'<div style="font-size:11px;color:var(--muted);padding:12px;">No compressor linked to this rig</div>'}</div><div><div class="st"><div class="sb" style="background:var(--amber);"></div><span class="sl">Active Work Orders</span></div>${rWOs.length?rWOs.map(w=>`<div style="padding:6px 0;border-bottom:1px solid var(--border);"><div style="display:flex;justify-content:space-between;margin-bottom:2px;"><span class="mono" style="font-size:9px;color:var(--accent);">${w.ID}</span>${pl(w.Status,wSC(w.Status))}</div><div style="font-size:11px;">${w.Description.substring(0,45)}...</div></div>`).join(""):'<div style="font-size:11px;color:var(--muted);padding:8px 0;">No active WOs</div>'}</div></div></div>`;}
function buildDD(){$("dd-summary").innerHTML=assetSummary(D.dd);$("dd-cards").innerHTML=D.dd.map(d=>`<div class="ac${selDD===d.ID?" sel":""}" id="ddc-${d.ID}" onclick="selDDf('${d.ID}')" style="border:1px solid ${fade(selDD===d.ID?"var(--dd)":sC(d.Status),30)};"><div style="display:flex;justify-content:space-between;margin-bottom:7px;"><span class="mono" style="font-size:9px;font-weight:700;color:var(--dd);">${d.ID}</span>${pl(d.Status,sC(d.Status))}</div><div style="font-size:12px;font-weight:700;margin-bottom:2px;">${d.Name}</div><div style="font-size:9px;color:var(--muted);margin-bottom:10px;">${d.Site}</div><div style="display:flex;justify-content:space-between;margin-bottom:6px;"><div><div style="font-size:20px;font-weight:900;color:${Number(d.Availability)>=90?"#00e087":Number(d.Availability)>=80?"#f5a800":"#ff3355"};">${d.Availability}%</div><div style="font-size:8px;color:var(--muted);font-family:'JetBrains Mono',monospace;">AVAIL</div></div><div style="text-align:right;"><div style="font-size:20px;font-weight:900;">${Number(d.Hours).toLocaleString()}</div><div style="font-size:8px;color:var(--muted);font-family:'JetBrains Mono',monospace;">HRS</div></div></div>${d.CoreRecovery?`<div class="mono" style="font-size:8px;font-weight:700;color:${Number(d.CoreRecovery)>=95?"#00e087":Number(d.CoreRecovery)>=85?"#f5a800":"#ff3355"};">CORE RECOVERY: ${d.CoreRecovery}%</div>`:""}  <div class="mono" style="font-size:8px;font-weight:700;color:${d.HrsToNextPM<0?"#ff3355":d.HrsToNextPM<100?"#f5a800":"#00e087"};margin-top:4px;">${d.HrsToNextPM<0?"⚠ PM OVERDUE "+Math.abs(d.HrsToNextPM)+"h":"NEXT PM: "+d.HrsToNextPM+"h"}</div>${d.MetresDrilled?`<div class="mono" style="font-size:8px;color:var(--muted);margin-top:2px;">${d.MetresDrilled}m / ${d.ShiftTarget}m</div>`:""}</div>`).join("");}
function selDDf(id){selDD=selDD===id?null:id;document.querySelectorAll("[id^='ddc-']").forEach(c=>c.classList.remove("sel"));if(selDD)document.getElementById("ddc-"+selDD)?.classList.add("sel");renderDDDet();}
function renderDDDet(){const el=$("dd-detail");if(!selDD){el.innerHTML="";return;}const d=D.dd.find(x=>x.ID===selDD);if(!d){el.innerHTML="";return;}const sys=[["Engine",d.Engine],["Hydraulics",d.Hydraulics],["Rotation Head",d.RotationHead],["Bean Pump",d.BeanPump],["Wireline Winch",d.WirelineWinch],["Water Pump",d.WaterPump],["Rod Handling",d.RodHandling],["Core Barrel",d.CoreBarrel],["Mast & Feed",d.MastFeed],["Electrical",d.Electrical]];const dWOs=D.wo.filter(w=>w.AssetID===d.ID);el.innerHTML=`<div class="dp fade-in" style="border:1px solid rgba(168,85,247,.25);"><div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:16px;"><div><div style="font-size:16px;font-weight:800;">${d.Name}</div><div style="font-size:11px;color:var(--muted);">${d.Site} | ${d.ID} | ${Number(d.Hours).toLocaleString()} hrs | Bean Pump — No external compressor</div></div>${pl(d.Status,sC(d.Status))}</div><div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:18px;"><div><div class="st"><div class="sb" style="background:var(--dd);"></div><span class="sl">DD Rig Systems</span></div>${sys.map(([n,v])=>`<div class="hr2"><div class="hs">${n}</div><div class="bt"><div class="bf" style="width:${(v/6)*100}%;background:${hC(v)};"></div></div>${pl(hL(v),hC(v))}</div>`).join("")}</div><div><div class="st"><div class="sb" style="background:var(--amber);"></div><span class="sl">Drilling Performance</span></div>${[["Asset ID",d.ID],["Meter Hours",Number(d.Hours).toLocaleString()+" hrs"],["Core Recovery",d.CoreRecovery?d.CoreRecovery+"%":"N/A"],["Metres (shift)",d.MetresDrilled?Number(d.MetresDrilled).toLocaleString()+" m":"N/A"],["Shift Target",d.ShiftTarget?Number(d.ShiftTarget).toLocaleString()+" m":"N/A"],["Cost/Hr","$"+d.CostPerHour+"/hr"],["Last PM",d.LastPM],["Next PM",d.NextPM],["Hrs to PM",d.HrsToNextPM<0?"OVERDUE "+Math.abs(d.HrsToNextPM)+" hrs":d.HrsToNextPM+" hrs"]].map(([l,v])=>`<div class="ir"><span style="font-size:10px;color:var(--muted);">${l}</span><span style="font-size:11px;font-weight:600;color:${l.includes("Hrs")&&d.HrsToNextPM<0?"#ff3355":l.includes("Core")&&Number(d.CoreRecovery)<85?"#ff3355":"var(--text)"};">${v}</span></div>`).join("")}</div><div><div class="st"><div class="sb" style="background:var(--purple);"></div><span class="sl">Active Work Orders</span></div>${dWOs.length?dWOs.map(w=>`<div style="padding:6px 0;border-bottom:1px solid var(--border);"><div style="display:flex;justify-content:space-between;margin-bottom:2px;"><span class="mono" style="font-size:9px;color:var(--accent);">${w.ID}</span>${pl(w.Status,wSC(w.Status))}</div><div style="font-size:11px;">${w.Description.substring(0,45)}...</div></div>`).join(""):'<div style="font-size:11px;color:var(--muted);padding:8px 0;">No active WOs</div>'}</div></div></div>`;}
function buildComp(){$("comp-summary").innerHTML=assetSummary(D.comp);$("comp-cards").innerHTML=D.comp.map(c=>`<div class="ac${selComp===c.ID?" sel":""}" id="cpc-${c.ID}" onclick="selCF('${c.ID}')" style="border:1px solid ${fade(selComp===c.ID?"var(--comp)":sC(c.Status),30)};"><div style="display:flex;justify-content:space-between;margin-bottom:7px;"><span class="mono" style="font-size:9px;font-weight:700;color:var(--comp);">${c.ID}</span>${pl(c.Status,sC(c.Status))}</div><div style="font-size:12px;font-weight:700;margin-bottom:1px;">${c.Name}</div><div style="font-size:9px;color:var(--muted);margin-bottom:2px;">${c.Make} ${c.Model}</div><div style="font-size:9px;color:var(--rc);font-family:'JetBrains Mono',monospace;margin-bottom:9px;">→ RC: ${c.AssignedTo||"Unassigned"}</div><div style="display:flex;justify-content:space-between;margin-bottom:6px;"><div><div style="font-size:20px;font-weight:900;color:${Number(c.Availability)>=90?"#00e087":Number(c.Availability)>=80?"#f5a800":"#ff3355"};">${c.Availability}%</div><div style="font-size:8px;color:var(--muted);font-family:'JetBrains Mono',monospace;">AVAIL</div></div><div style="text-align:right;"><div style="font-size:20px;font-weight:900;">${Number(c.Hours).toLocaleString()}</div><div style="font-size:8px;color:var(--muted);font-family:'JetBrains Mono',monospace;">HRS</div></div></div><div class="mono" style="font-size:8px;color:var(--muted);">${c.CFM} CFM | ${c.WorkingPressure} psi</div><div class="mono" style="font-size:8px;font-weight:700;color:${c.HrsToNextPM<0?"#ff3355":c.HrsToNextPM<100?"#f5a800":"#00e087"};margin-top:4px;">${c.HrsToNextPM<0?"⚠ PM OVERDUE "+Math.abs(c.HrsToNextPM)+"h":"NEXT PM: "+c.HrsToNextPM+"h"}</div></div>`).join("");}
function selCF(id){selComp=selComp===id?null:id;document.querySelectorAll("[id^='cpc-']").forEach(c=>c.classList.remove("sel"));if(selComp)document.getElementById("cpc-"+selComp)?.classList.add("sel");renderCompDet();}
function renderCompDet(){const el=$("comp-detail");if(!selComp){el.innerHTML="";return;}const c=D.comp.find(x=>x.ID===selComp);if(!c){el.innerHTML="";return;}const rc=D.rc.find(r=>r.ConnectedCompressor===c.ID);const sys=[["Engine",c.Engine],["Air End / Element",c.AirEnd],["Oil Separator",c.OilSeparator],["Cooling System",c.CoolingSystem],["Pressure & Safety Valves",c.PressureValves],["Electrical & Controls",c.Electrical]];const cWOs=D.wo.filter(w=>w.AssetID===c.ID);el.innerHTML=`<div class="dp fade-in" style="border:1px solid rgba(245,168,0,.25);"><div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:16px;"><div><div style="font-size:16px;font-weight:800;">${c.Name}</div><div style="font-size:11px;color:var(--muted);">${c.Make} ${c.Model} | ${c.ID} | ${Number(c.Hours).toLocaleString()} hrs | Assigned: ${c.AssignedTo||"Unassigned"}</div></div>${pl(c.Status,sC(c.Status))}</div><div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:18px;"><div><div class="st"><div class="sb" style="background:var(--comp);"></div><span class="sl">Compressor Systems</span></div>${sys.map(([n,v])=>`<div class="hr2"><div class="hs">${n}</div><div class="bt"><div class="bf" style="width:${(v/6)*100}%;background:${hC(v)};"></div></div>${pl(hL(v),hC(v))}</div>`).join("")}</div><div><div class="st"><div class="sb" style="background:var(--amber);"></div><span class="sl">Asset Details</span></div>${[["Asset ID",c.ID],["Make / Model",c.Make+" "+c.Model],["Output",c.CFM+" CFM"],["Working Pressure",c.WorkingPressure+" psi"],["Meter Hours",Number(c.Hours).toLocaleString()+" hrs"],["Cost/Hr","$"+c.CostPerHour+"/hr"],["Last PM",c.LastPM],["Next PM",c.NextPM],["Hrs to Next PM",c.HrsToNextPM<0?"OVERDUE "+Math.abs(c.HrsToNextPM)+" hrs":c.HrsToNextPM+" hrs"]].map(([l,v])=>`<div class="ir"><span style="font-size:10px;color:var(--muted);">${l}</span><span style="font-size:11px;font-weight:600;color:${l.includes("Hrs")&&c.HrsToNextPM<0?"#ff3355":"var(--text)"};">${v}</span></div>`).join("")}${rc?`<div style="margin-top:12px;background:rgba(0,200,240,.08);border:1px solid rgba(0,200,240,.2);border-radius:7px;padding:10px;"><div style="font-size:9px;color:var(--muted);margin-bottom:3px;font-family:'JetBrains Mono',monospace;">CONNECTED RC RIG</div><div style="font-size:12px;font-weight:700;color:var(--rc);">${rc.Name}</div><div style="font-size:9px;color:var(--muted);">${rc.Site}</div>${c.Status==="DOWN"?'<div style="font-size:10px;color:#ff3355;font-weight:700;margin-top:5px;">⚠ THIS RC RIG CANNOT DRILL — NO AIR SUPPLY</div>':""}</div>`:""}</div><div><div class="st"><div class="sb" style="background:var(--purple);"></div><span class="sl">Active Work Orders</span></div>${cWOs.length?cWOs.map(w=>`<div style="padding:6px 0;border-bottom:1px solid var(--border);"><div style="display:flex;justify-content:space-between;margin-bottom:2px;"><span class="mono" style="font-size:9px;color:var(--accent);">${w.ID}</span>${pl(w.Status,wSC(w.Status))}</div><div style="font-size:11px;">${w.Description.substring(0,45)}...</div></div>`).join(""):'<div style="font-size:11px;color:var(--muted);padding:8px 0;">No active WOs</div>'}</div></div></div>`;}
function buildWOs(){const op=D.wo.filter(w=>w.Status!=="COMPLETED").length;const pr=D.wo.filter(w=>w.Status==="IN PROGRESS").length;const hi=D.wo.filter(w=>w.Priority==="HIGH"&&w.Status!=="COMPLETED").length;const dn=D.wo.filter(w=>w.Status==="COMPLETED").length;$("wo-summary").innerHTML=[{l:"Open WOs",v:op,c:"#f5a800"},{l:"In Progress",v:pr,c:"#00c8f0"},{l:"High Priority",v:hi,c:"#ff3355"},{l:"Completed",v:dn,c:"#00e087"}].map(s=>sumCard(s.v,s.l,s.c)).join("");renderWOT(D.wo);}
function filterWOs(){const at=$("wof-at")?.value||"",wt=$("wof-wt")?.value||"",pr=$("wof-pr")?.value||"",st=$("wof-st")?.value||"";renderWOT(D.wo.filter(w=>(!at||w.AssetType===at)&&(!wt||w.Type===wt)&&(!pr||w.Priority===pr)&&(!st||w.Status===st)));}
function clearWOF(){["wof-at","wof-wt","wof-pr","wof-st"].forEach(id=>{const e=$(id);if(e)e.value="";});renderWOT(D.wo);}

function lvStatusCol(s){return s==="AVAILABLE"||s==="ON ROAD"?"#00e087":s==="IN WORKSHOP"?"#f5a800":"#5a6e8a";}
function lvHealthCol(v){return v<=2?"#00e087":v<=3?"#f5a800":"#ff3355";}
function daysUntil(dateStr){try{const d=new Date(dateStr);const n=new Date();return Math.round((d-n)/(1000*60*60*24));}catch(e){return 999;}}
function expiryStatus(dateStr){const days=daysUntil(dateStr);return days<0?"EXPIRED":days<=30?"EXPIRING SOON":"VALID";}
function expiryCol(dateStr){const days=daysUntil(dateStr);return days<0?"#ff3355":days<=30?"#f5a800":"#00e087";}

function buildLV(){
  const lvs=D.lv||[];
  const avail=lvs.filter(v=>v.Status==="AVAILABLE"||v.Status==="ON ROAD").length;
  const ws=lvs.filter(v=>v.Status==="IN WORKSHOP").length;
  const comp_exp=lvs.filter(v=>daysUntil(v.RoadworthyExpiry)<0||daysUntil(v.InsuranceExpiry)<0||daysUntil(v.FireExtExpiry)<0||daysUntil(v.FirstAidExpiry)<0).length;
  const svc_due=lvs.filter(v=>Number(v.KmToNextService)<=0).length;
  $("lv-summary").innerHTML=[
    {l:"Available",v:avail,c:"#00e087"},
    {l:"In Workshop",v:ws,c:"#f5a800"},
    {l:"Compliance Expired",v:comp_exp,c:comp_exp>0?"#ff3355":"#00e087"},
    {l:"Service Overdue",v:svc_due,c:svc_due>0?"#ff3355":"#00e087"},
  ].map(s=>sumCard(s.v,s.l,s.c)).join("");

  const expired=lvs.filter(v=>daysUntil(v.RoadworthyExpiry)<0||daysUntil(v.InsuranceExpiry)<0||daysUntil(v.FireExtExpiry)<0||daysUntil(v.FirstAidExpiry)<0);
  const expiring=lvs.filter(v=>{const d=[v.RoadworthyExpiry,v.InsuranceExpiry,v.FireExtExpiry,v.FirstAidExpiry];return d.some(x=>daysUntil(x)>=0&&daysUntil(x)<=30);});
  let banner="";
  if(expired.length>0) banner+=`<div style="background:rgba(255,51,85,.1);border:1px solid rgba(255,51,85,.3);border-radius:8px;padding:10px 14px;font-size:11px;margin-bottom:8px;"><strong style="color:#ff3355;">⚠ COMPLIANCE EXPIRED:</strong> ${expired.map(v=>v.Name).join(", ")} — Remove from service immediately. Mine site access will be denied.</div>`;
  if(expiring.length>0) banner+=`<div style="background:rgba(245,168,0,.08);border:1px solid rgba(245,168,0,.25);border-radius:8px;padding:10px 14px;font-size:11px;"><strong style="color:#f5a800;">⚠ EXPIRING WITHIN 30 DAYS:</strong> ${expiring.map(v=>v.Name).join(", ")} — Renew compliance documents urgently.</div>`;
  $("lv-compliance-banner").innerHTML=banner;

  $("lv-cards").innerHTML=lvs.map(v=>`
    <div class="ac${selLV===v.ID?" sel":""}" id="lvc-${v.ID}" onclick="selLVf('${v.ID}')" style="border:1px solid ${fade(selLV===v.ID?"var(--lv)":lvStatusCol(v.Status),30)};">
      <div style="display:flex;justify-content:space-between;margin-bottom:7px;">
        <span class="mono" style="font-size:9px;font-weight:700;color:var(--lv);">${v.ID}</span>
        ${pl(v.Status,lvStatusCol(v.Status))}
      </div>
      <div style="font-size:11px;font-weight:700;margin-bottom:1px;">${v.Name}</div>
      <div style="font-size:9px;color:var(--muted);margin-bottom:9px;">${v.Type} | ${v.Site}</div>
      <div style="display:flex;justify-content:space-between;margin-bottom:8px;">
        <div><div style="font-size:18px;font-weight:900;color:${Number(v.Availability)>=90?"#00e087":Number(v.Availability)>=75?"#f5a800":"#ff3355"};">${v.Availability}%</div><div style="font-size:8px;color:var(--muted);font-family:'JetBrains Mono',monospace;">AVAIL</div></div>
        <div style="text-align:right;"><div style="font-size:18px;font-weight:900;">${Number(v.Odometer).toLocaleString()}</div><div style="font-size:8px;color:var(--muted);font-family:'JetBrains Mono',monospace;">KM</div></div>
      </div>
      <div class="mono" style="font-size:8px;font-weight:700;color:${Number(v.KmToNextService)<=0?"#ff3355":Number(v.KmToNextService)<=1000?"#f5a800":"#00e087"};">
        ${Number(v.KmToNextService)<=0?"⚠ SERVICE OVERDUE "+Math.abs(v.KmToNextService)+"km":"NEXT SVC: "+Number(v.KmToNextService).toLocaleString()+"km"}
      </div>
      <div class="mono" style="font-size:8px;margin-top:4px;color:${expiryCol(v.RoadworthyExpiry)};">
        RW: ${expiryStatus(v.RoadworthyExpiry)}
      </div>
      <div class="mono" style="font-size:8px;margin-top:2px;color:${v.PreStartCompliance>=95?"#00e087":v.PreStartCompliance>=85?"#f5a800":"#ff3355"};">
        Pre-start: ${v.PreStartCompliance}%
      </div>
      ${Number(v.TyreTreadMm)<3?`<div class="mono" style="font-size:8px;margin-top:2px;color:#ff3355;">⚠ TYRE TREAD: ${v.TyreTreadMm}mm — BELOW MIN</div>`:""}
    </div>`).join("");

  $("lv-prestart").innerHTML=`<table class="dt"><thead><tr><th>VEHICLE</th><th>DRIVER</th><th>COMPLIANCE %</th><th>STATUS</th></tr></thead><tbody>${lvs.map(v=>{const col=v.PreStartCompliance>=95?"#00e087":v.PreStartCompliance>=85?"#f5a800":"#ff3355";const stat=v.PreStartCompliance>=95?"ON TARGET":v.PreStartCompliance>=85?"BELOW TARGET":"ACTION REQUIRED";return `<tr><td style="font-weight:600;">${v.Name}</td><td>${v.Driver}</td><td><div style="display:flex;align-items:center;gap:10px;"><span class="mono" style="font-weight:700;color:${col};">${v.PreStartCompliance}%</span><div class="bt"><div class="bf" style="width:${v.PreStartCompliance}%;background:${col};"></div></div></div></td><td>${pl(stat,col)}</td></tr>`;}).join("")}</tbody></table>`;

  $("lv-expiry").innerHTML=`<table class="dt"><thead><tr><th>VEHICLE</th><th>ROADWORTHY</th><th>INSURANCE</th><th>FIRE EXT</th><th>FIRST AID</th></tr></thead><tbody>${lvs.map(v=>`<tr><td style="font-weight:600;">${v.Name}</td><td style="font-size:10px;color:${expiryCol(v.RoadworthyExpiry)};font-family:'JetBrains Mono',monospace;">${v.RoadworthyExpiry}<br><span style="font-size:9px;">${expiryStatus(v.RoadworthyExpiry)}</span></td><td style="font-size:10px;color:${expiryCol(v.InsuranceExpiry)};font-family:'JetBrains Mono',monospace;">${v.InsuranceExpiry}<br><span style="font-size:9px;">${expiryStatus(v.InsuranceExpiry)}</span></td><td style="font-size:10px;color:${expiryCol(v.FireExtExpiry)};font-family:'JetBrains Mono',monospace;">${v.FireExtExpiry}<br><span style="font-size:9px;">${expiryStatus(v.FireExtExpiry)}</span></td><td style="font-size:10px;color:${expiryCol(v.FirstAidExpiry)};font-family:'JetBrains Mono',monospace;">${v.FirstAidExpiry}<br><span style="font-size:9px;">${expiryStatus(v.FirstAidExpiry)}</span></td></tr>`).join("")}</tbody></table>`;

  const avgAvail=lvs.length?Math.round(lvs.reduce((s,v)=>s+Number(v.Availability),0)/lvs.length):0;
  const avgFuel=lvs.length?Math.round(lvs.reduce((s,v)=>s+Number(v.FuelL100km),0)/lvs.length*10)/10:0;
  const avgPSC=lvs.length?Math.round(lvs.reduce((s,v)=>s+Number(v.PreStartCompliance),0)/lvs.length):0;
  const avgCPK=lvs.length?Math.round(lvs.reduce((s,v)=>s+Number(v.CostPerKm),0)/lvs.length*100)/100:0;
  $("lv-kpis").innerHTML=[
    {l:"Fleet Availability",v:avgAvail+"%",c:avgAvail>=90?"#00e087":avgAvail>=80?"#f5a800":"#ff3355",t:"Target ≥ 90%"},
    {l:"Pre-Start Compliance",v:avgPSC+"%",c:avgPSC>=95?"#00e087":avgPSC>=85?"#f5a800":"#ff3355",t:"Target 100%"},
    {l:"Avg Fuel (L/100km)",v:avgFuel+" L",c:avgFuel<=15?"#00e087":avgFuel<=22?"#f5a800":"#ff3355",t:"Monitor trend"},
    {l:"Avg Cost per km",v:"$"+avgCPK+"/km",c:avgCPK<=2.5?"#00e087":avgCPK<=3.5?"#f5a800":"#ff3355",t:"Monitor trend"},
    {l:"Vehicles Service Overdue",v:svc_due,c:svc_due===0?"#00e087":"#ff3355",t:"Target: 0"},
    {l:"Compliance Documents Expired",v:comp_exp,c:comp_exp===0?"#00e087":"#ff3355",t:"Must be 0 at all times"},
    {l:"Tyres Below 3mm",v:lvs.filter(v=>Number(v.TyreTreadMm)<3).length,c:lvs.filter(v=>Number(v.TyreTreadMm)<3).length===0?"#00e087":"#ff3355",t:"Mine minimum: 3mm"},
    {l:"Vehicles in Workshop",v:ws,c:"#f5a800",t:"Track downtime impact"},
  ].map(k=>`<div style="display:flex;justify-content:space-between;align-items:center;padding:8px 0;border-bottom:1px solid var(--border);"><div><div style="font-size:11px;font-weight:600;">${k.l}</div><div style="font-size:9px;color:var(--muted);">${k.t}</div></div><div style="font-size:16px;font-weight:900;color:${k.c};">${k.v}</div></div>`).join("");

  $("lv-register").innerHTML=lvs.map(v=>`<tr>
    <td class="mono" style="color:var(--lv);">${v.ID}</td>
    <td>${pl(v.Type,"#00b4a0")}</td>
    <td style="font-size:10px;">${v.Make} ${v.Model}</td>
    <td class="mono">${Number(v.Odometer).toLocaleString()} km</td>
    <td class="mono" style="color:${Number(v.KmToNextService)<=0?"#ff3355":Number(v.KmToNextService)<=1000?"#f5a800":"#00e087"};">${Number(v.KmToNextService)<=0?"OVERDUE "+Math.abs(v.KmToNextService)+"km":Number(v.KmToNextService).toLocaleString()+"km"}</td>
    <td style="font-size:10px;color:${expiryCol(v.RoadworthyExpiry)};">${v.RoadworthyExpiry}</td>
    <td class="mono" style="color:${Number(v.Availability)>=90?"#00e087":Number(v.Availability)>=80?"#f5a800":"#ff3355"};">${v.Availability}%</td>
    <td>${pl(v.Status,lvStatusCol(v.Status))}</td>
  </tr>`).join("");
}

function selLVf(id){selLV=selLV===id?null:id;document.querySelectorAll("[id^='lvc-']").forEach(c=>c.classList.remove("sel"));if(selLV)document.getElementById("lvc-"+selLV)?.classList.add("sel");renderLVDet();}

function renderLVDet(){
  const el=$("lv-detail");if(!selLV){el.innerHTML="";return;}
  const v=(D.lv||[]).find(x=>x.ID===selLV);if(!v){el.innerHTML="";return;}
  const sys=[["Engine",v.Engine],["Tyres",v.Tyres],["Brakes",v.Brakes],["Electrical",v.Electrical],["Body & Chassis",v.BodyChassis]];
  const lWOs=D.wo.filter(w=>w.AssetID===v.ID);
  el.innerHTML=`<div class="dp fade-in" style="border:1px solid rgba(0,180,160,.25);">
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:16px;">
      <div><div style="font-size:16px;font-weight:800;">${v.Name}</div><div style="font-size:11px;color:var(--muted);">${v.Type} | ${v.Site} | ${Number(v.Odometer).toLocaleString()} km | Driver: ${v.Driver}</div></div>
      ${pl(v.Status,lvStatusCol(v.Status))}
    </div>
    <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:18px;">
      <div>
        <div class="st"><div class="sb" style="background:var(--lv);"></div><span class="sl">Vehicle System Health</span></div>
        ${sys.map(([n,val])=>`<div class="hr2"><div class="hs">${n}</div><div class="bt"><div class="bf" style="width:${(val/6)*100}%;background:${lvHealthCol(val)};"></div></div>${pl(lvHealthCol(val)==="#00e087"?"GOOD":lvHealthCol(val)==="#f5a800"?"MONITOR":"ATTENTION",lvHealthCol(val))}</div>`).join("")}
        <div style="margin-top:12px;">
          <div class="st"><div class="sb" style="background:var(--amber);"></div><span class="sl">Tyre Status</span></div>
          <div style="display:flex;justify-content:space-between;padding:6px 0;border-bottom:1px solid var(--border);"><span style="font-size:11px;color:var(--muted);">Tread Depth</span><span style="font-size:12px;font-weight:700;color:${Number(v.TyreTreadMm)>=4?"#00e087":Number(v.TyreTreadMm)>=3?"#f5a800":"#ff3355"};">${v.TyreTreadMm}mm ${Number(v.TyreTreadMm)<3?"⚠ REPLACE":Number(v.TyreTreadMm)<4?"MONITOR":""}</span></div>
        </div>
      </div>
      <div>
        <div class="st"><div class="sb" style="background:var(--amber);"></div><span class="sl">Service & Compliance</span></div>
        ${[
          ["Odometer",Number(v.Odometer).toLocaleString()+" km"],
          ["Last Service",v.LastService],
          ["Next Service Due",Number(v.NextServiceKm).toLocaleString()+" km"],
          ["Km to Next Service",Number(v.KmToNextService)<=0?"OVERDUE "+Math.abs(v.KmToNextService)+"km":Number(v.KmToNextService).toLocaleString()+" km"],
          ["Fuel Consumption",v.FuelL100km+" L/100km"],
          ["Cost per km","$"+v.CostPerKm+"/km"],
          ["Roadworthy Expiry",v.RoadworthyExpiry+" — "+expiryStatus(v.RoadworthyExpiry)],
          ["Insurance Expiry",v.InsuranceExpiry+" — "+expiryStatus(v.InsuranceExpiry)],
          ["Fire Extinguisher",v.FireExtExpiry+" — "+expiryStatus(v.FireExtExpiry)],
          ["First Aid Kit",v.FirstAidExpiry+" — "+expiryStatus(v.FirstAidExpiry)],
          ["Pre-Start Compliance",v.PreStartCompliance+"%"],
        ].map(([l,val])=>`<div class="ir"><span style="font-size:10px;color:var(--muted);">${l}</span><span style="font-size:11px;font-weight:600;color:${(l.includes("km")&&Number(v.KmToNextService)<=0)||(l.includes("Expiry")&&val.includes("EXPIRED"))||(l.includes("Pre-Start")&&v.PreStartCompliance<85)?"#ff3355":"var(--text)"};">${val}</span></div>`).join("")}
      </div>
      <div>
        <div class="st"><div class="sb" style="background:var(--purple);"></div><span class="sl">Active Work Orders</span></div>
        ${lWOs.length?lWOs.map(w=>`<div style="padding:6px 0;border-bottom:1px solid var(--border);"><div style="display:flex;justify-content:space-between;margin-bottom:2px;"><span class="mono" style="font-size:9px;color:var(--accent);">${w.ID}</span>${pl(w.Status,wSC(w.Status))}</div><div style="font-size:11px;">${w.Description.substring(0,45)}...</div></div>`).join(""):'<div style="font-size:11px;color:var(--muted);padding:8px 0;">No active WOs</div>'}
      </div>
    </div>
  </div>`;
}

function renderWOT(wos){const c=$("wo-count");if(c)c.textContent=`(${wos.length} records)`;$("wo-tbody").innerHTML=wos.map(w=>`<tr><td class="mono" style="color:var(--accent);">${w.ID}</td><td style="font-weight:600;">${w.AssetID}</td><td>${pl(w.AssetType||"",aTC(w.AssetType||""))}</td><td>${pl(w.Type,wTC(w.Type))}</td><td>${pl(w.Priority,pC(w.Priority))}</td><td style="color:var(--muted);">${w.System}</td><td style="max-width:180px;">${w.Description}</td><td>${w.Technician}</td><td class="mono">${w.DueDate}</td><td class="mono" style="text-align:center;">${w.Hours}h</td><td class="mono" style="color:${Number(w.DowntimeCost)>0?"#ff3355":"var(--muted)"};">${Number(w.DowntimeCost)>0?"$"+Number(w.DowntimeCost).toLocaleString():"—"}</td><td>${pl(w.Status,wSC(w.Status))}</td></tr>`).join("");}
function buildCal(){renderCal();const a=all();$("cal-summary").innerHTML=[{l:"Due This Month",v:a.filter(x=>Number(x.HrsToNextPM)<=30&&Number(x.HrsToNextPM)>0).length,c:"#f5a800"},{l:"Overdue",v:a.filter(x=>Number(x.HrsToNextPM)<0).length,c:"#ff3355"},{l:"PMs Completed",v:D.wo.filter(w=>w.Type==="PM"&&w.Status==="COMPLETED").length,c:"#00e087"}].map(s=>sumCard(s.v,s.l,s.c)).join("");$("cal-list").innerHTML=a.map(x=>{const t=typeOf(x);const col=x.HrsToNextPM<0?"#ff3355":x.HrsToNextPM<100?"#f5a800":"#00e087";const stat=x.HrsToNextPM<0?"OVERDUE":x.HrsToNextPM<100?"DUE SOON":"SCHEDULED";return `<tr><td class="mono" style="color:${aTC(t)};">${x.ID}</td><td style="font-weight:600;">${x.Name}</td><td>${pl(t==="COMPRESSOR"?"COMP":t,aTC(t))}</td><td class="mono">${x.NextPM||"—"}</td><td class="mono" style="color:${col};">${x.HrsToNextPM<0?"OVERDUE "+Math.abs(x.HrsToNextPM)+"h":x.HrsToNextPM+"h"}</td><td class="mono">${x.LastPM||"—"}</td><td>${pl(stat,col)}</td></tr>`;}).join("");}
function renderCal(){const days=["SUN","MON","TUE","WED","THU","FRI","SAT"];const months=["January","February","March","April","May","June","July","August","September","October","November","December"];const el=$("cal-label");if(el)el.textContent=months[calM]+" "+calY;const hdr=$("cal-hdrs");if(hdr)hdr.innerHTML=days.map(d=>`<div class="cal-hdr">${d}</div>`).join("");const grid=$("cal-grid");if(!grid)return;const first=new Date(calY,calM,1).getDay(),dim=new Date(calY,calM+1,0).getDate(),today=new Date();let html="";for(let i=0;i<first;i++)html+=`<div></div>`;const a=all();for(let d=1;d<=dim;d++){const isT=d===today.getDate()&&calM===today.getMonth()&&calY===today.getFullYear();let ev="";a.forEach(x=>{try{const nd=new Date(x.NextPM);if(nd.getDate()===d&&nd.getMonth()===calM&&nd.getFullYear()===calY){const t=typeOf(x);const col=aTC(t==="COMPRESSOR"?"COMPRESSOR":t);ev+=`<div class="cal-event" style="background:${fade(col,16)};color:${col};border-left:2px solid ${col};">${x.ID.replace("DGL-","")}</div>`;}}catch(e){}});html+=`<div class="cal-day" style="background:${isT?"var(--surface3)":"var(--surface2)"};border:1px solid ${isT?"var(--accent)":"var(--border)"};"><div style="font-size:9px;font-weight:${isT?"800":"500"};color:${isT?"var(--accent)":"var(--muted)"};">${d}</div>${ev}</div>`;}grid.innerHTML=html;}
function prevMonth(){calM--;if(calM<0){calM=11;calY--;}renderCal();}
function nextMonth(){calM++;if(calM>11){calM=0;calY++;}renderCal();}
function goToday(){const n=new Date();calM=n.getMonth();calY=n.getFullYear();renderCal();}
function buildKPIs(){$("kpi-grid").innerHTML=D.kpis.map(k=>`<div class="card-sm"><div class="kl">${k.Metric}</div><div class="kv" style="color:${kC(k.Status)};">${k.Value}</div><div class="kf"><span class="kt">TGT ${k.Target}</span><span class="ktr" style="color:${kC(k.Status)};">${k.Trend}</span></div><div style="margin-top:8px;height:3px;background:var(--border);border-radius:2px;"><div style="height:100%;width:${k.Status==="green"?"85%":k.Status==="amber"?"60%":"35%"};background:${kC(k.Status)};border-radius:2px;"></div></div></div>`).join("");const tot=D.wo.length||1;const pm=D.wo.filter(w=>w.Type==="PM").length;const pdm=D.wo.filter(w=>w.Type==="PdM").length;const cm=D.wo.filter(w=>w.Type==="CORRECTIVE").length;const rx=D.wo.filter(w=>w.Type==="REACTIVE").length;$("maint-bd").innerHTML=[{l:"Preventive (PM)",v:Math.round(pm/tot*100),c:"var(--success)"},{l:"Predictive (PdM)",v:Math.round(pdm/tot*100),c:"var(--info)"},{l:"Corrective",v:Math.round(cm/tot*100),c:"var(--warning)"},{l:"Reactive",v:Math.round(rx/tot*100),c:"var(--danger)"}].map(b=>`<div style="margin-bottom:12px;"><div style="display:flex;justify-content:space-between;margin-bottom:3px;"><span style="font-size:11px;">${b.l}</span><span style="font-size:10px;font-weight:700;color:${b.c};font-family:'JetBrains Mono',monospace;">${b.v}%</span></div><div class="bt"><div class="bf" style="width:${b.v}%;background:${b.c};"></div></div></div>`).join("");const rcAv=D.rc.length?Math.round(D.rc.reduce((s,r)=>s+Number(r.Availability),0)/D.rc.length):0;const ddAv=D.dd.length?Math.round(D.dd.reduce((s,d)=>s+Number(d.Availability),0)/D.dd.length):0;const cAv=D.comp.length?Math.round(D.comp.reduce((s,c)=>s+Number(c.Availability),0)/D.comp.length):0;$("avail-type").innerHTML=[{l:"RC Drilling Rigs",v:rcAv,c:"var(--info)",t:85},{l:"Diamond Drill Rigs",v:ddAv,c:"var(--info)",t:85},{l:"External Compressors",v:cAv,c:"var(--info)",t:90}].map(b=>`<div style="margin-bottom:16px;"><div style="display:flex;justify-content:space-between;margin-bottom:4px;"><span style="font-size:12px;font-weight:600;">${b.l}</span><span style="font-size:14px;font-weight:900;color:${b.v>=b.t?"#00e087":b.v>=b.t-10?"#f5a800":"#ff3355"};">${b.v}%</span></div><div class="bt" style="height:8px;"><div class="bf" style="width:${b.v}%;background:${b.c};height:8px;"></div></div><div style="font-size:9px;color:var(--muted);margin-top:3px;">Target ≥ ${b.t}% — ${b.v>=b.t?"ON TARGET":"BELOW TARGET"}</div></div>`).join("");$("cph-chart").innerHTML=`<div style="display:flex;gap:6px;align-items:flex-end;height:100px;overflow-x:auto;padding-bottom:4px;">${all().map(a=>{const h=Math.min(Number(a.CostPerHour)||0,100);const t="var(--info)";return `<div style="min-width:34px;flex:1;display:flex;flex-direction:column;align-items:center;gap:3px;"><div style="font-size:8px;color:${t};font-family:'JetBrains Mono',monospace;">$${a.CostPerHour}</div><div style="width:100%;height:${h}px;background:${t};border-radius:3px 3px 0 0;"></div><div style="font-size:7px;color:var(--muted);text-align:center;font-family:'JetBrains Mono',monospace;">${a.ID.replace("DGL-","")}</div></div>`;}).join("")}</div><div style="font-size:9px;color:var(--muted);margin-top:8px;">Cost per operating hour — lower = better</div>`;}
function buildAnalytics(){const svg=$("bath-svg");if(svg){const W=500,H=170,p=38;const pts=[];for(let x=0;x<=W-p*2;x+=5){const t=x/(W-p*2);const y=H-p-(3*Math.exp(-t*8)+0.5+2*Math.exp((t-1)*6))*16;pts.push(`${x+p},${Math.max(p,y)}`);}const ps=pts.join(" ");svg.innerHTML=`<defs><linearGradient id="bg2" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#a855f7" stop-opacity=".25"/><stop offset="100%" stop-color="#a855f7" stop-opacity=".02"/></linearGradient></defs><line x1="${p}" y1="${H-p}" x2="${W-p}" y2="${H-p}" stroke="var(--border2)" stroke-width="1"/><line x1="${p}" y1="${p}" x2="${p}" y2="${H-p}" stroke="var(--border2)" stroke-width="1"/><text x="${W/2}" y="${H-6}" text-anchor="middle" fill="var(--muted)" font-size="9" font-family="JetBrains Mono">Operating Hours →</text><text x="14" y="${H/2}" text-anchor="middle" fill="var(--muted)" font-size="9" font-family="JetBrains Mono" transform="rotate(-90,14,${H/2})">Failure Rate</text><polygon points="${p},${H-p} ${ps} ${W-p},${H-p}" fill="url(#bg2)"/><polyline points="${ps}" fill="none" stroke="#a855f7" stroke-width="2.5" stroke-linejoin="round"/><line x1="${p+55}" y1="${p}" x2="${p+55}" y2="${H-p}" stroke="var(--red)" stroke-width="1" stroke-dasharray="4,3"/><line x1="${W-p-75}" y1="${p}" x2="${W-p-75}" y2="${H-p}" stroke="var(--amber)" stroke-width="1" stroke-dasharray="4,3"/>${all().map((a)=>{const xP=Math.min(p+((Number(a.Hours)/10000)*(W-p*2)),W-p-8);const col="var(--info)";return `<circle cx="${xP}" cy="${H-p-18}" r="5" fill="${col}" stroke="var(--bg)" stroke-width="2"/><text x="${xP}" y="${H-p-26}" text-anchor="middle" fill="${col}" font-size="7" font-family="JetBrains Mono">${a.ID.replace("DGL-","")}</text>`;}).join("")}`;}
const fm={};D.wo.filter(w=>w.System).forEach(w=>{fm[w.System]=(fm[w.System]||0)+1;});const fmS=Object.entries(fm).sort((a,b)=>b[1]-a[1]);const fmM=fmS[0]?.[1]||1;$("fail-freq").innerHTML=fmS.length?fmS.map(([s,c],i)=>{const col=["var(--info)"][i]||"var(--info)";return `<div style="margin-bottom:8px;"><div style="display:flex;justify-content:space-between;margin-bottom:3px;"><span style="font-size:11px;">${s}</span><span style="font-size:10px;font-weight:700;color:${col};font-family:'JetBrains Mono',monospace;">${c} events</span></div><div class="bt"><div class="bf" style="width:${(c/fmM)*100}%;background:${col};"></div></div></div>`;}).join(""):'<div style="font-size:11px;color:var(--muted);padding:16px;text-align:center;">No failure data yet</div>';
$("pf-track").innerHTML=[{c:"Track Idler Bearing (RC)",pf:"250 hrs",a:"Greasing — nipple check",i:"250 hrs",s:"Within window"},{c:"Bean Pump Seal (DD)",pf:"300 hrs",a:"Visual leak check + seal inspection",i:"250 hrs",s:"Within window"},{c:"Compressor Air End (Comp)",pf:"200 hrs",a:"Oil analysis + vibration check",i:"250 hrs",s:"Review — P-F shorter than interval"},{c:"Wireline Cable (DD)",pf:"500 hrs",a:"Visual inspection + diameter measure",i:"500 hrs",s:"Within window"},{c:"Oil Separator Blockage (Comp)",pf:"150 hrs",a:"Differential pressure monitoring",i:"250 hrs",s:"⚠ Interval too wide — shorten to 150hr"},{c:"Hydraulic Hoses (All)",pf:"400 hrs",a:"Visual + pressure test",i:"500 hrs",s:"Review interval"},{c:"Core Barrel Wear (DD)",pf:"300 hrs",a:"Core barrel diameter measurement",i:"500 hrs",s:"Review — P-F shorter than interval"}].map(p=>`<div style="display:flex;justify-content:space-between;padding:7px 0;border-bottom:1px solid var(--border);"><div><div style="font-size:11px;font-weight:600;">${p.c}</div><div style="font-size:9px;color:var(--muted);">${p.a}</div></div><div style="text-align:right;"><div style="font-size:9px;color:var(--accent);font-family:'JetBrains Mono',monospace;">P-F: ${p.pf}</div><div style="font-size:9px;color:${p.s.includes("⚠")||p.s.includes("Review")?"#f5a800":"#00e087"};">${p.s}</div></div></div>`).join("");
$("rul-track").innerHTML=all().map(a=>{const kl=D.rc.includes(a)?["Engine","Hydraulics","Track","RotationHead","RCSampleSystem","Electrical"]:D.dd.includes(a)?["Engine","Hydraulics","RotationHead","BeanPump","WirelineWinch","WaterPump","RodHandling","CoreBarrel","MastFeed","Electrical"]:["Engine","AirEnd","OilSeparator","CoolingSystem","PressureValves","Electrical"];const avg=Math.round(kl.reduce((s,k)=>s+Number(a[k]||3),0)/kl.length*10)/10;const rul=avg<=2?"> 3,000 hrs":avg<=3?"1,000–3,000 hrs":avg<=4?"500–1,000 hrs":"< 500 hrs — PLAN OVERHAUL";const col=avg<=2?"#00e087":avg<=3?"#f5a800":"#ff3355";const t=D.rc.includes(a)?"RC":D.dd.includes(a)?"DD":"COMP";return `<div style="display:flex;justify-content:space-between;align-items:center;padding:6px 0;border-bottom:1px solid var(--border);"><div><div style="display:flex;align-items:center;gap:6px;"><span style="font-size:11px;font-weight:600;">${a.Name}</span>${pl(t,aTC(t==="COMP"?"COMPRESSOR":t))}</div><div style="font-size:9px;color:var(--muted);">${Number(a.Hours).toLocaleString()} hrs | Avg health: ${avg}/6</div></div><div style="font-size:9px;color:${col};font-family:'JetBrains Mono',monospace;font-weight:700;text-align:right;">${rul}</div></div>`;}).join("");}
function buildTechs(){
  const dn=D.techs.reduce((s,t)=>s+Number(t.WOsCompleted||0),0);
  const op=D.techs.reduce((s,t)=>s+Number(t.WOsOpen||0),0);
  const on=D.techs.filter(t=>t.Status==="ON SHIFT").length;
  $("tech-summary").innerHTML=[
    {l:"On Shift",v:on,c:"#00e087"},
    {l:"Off Shift",v:D.techs.length-on,c:"#5a6e8a"},
    {l:"WOs Completed",v:dn,c:"#00c8f0"},
    {l:"WOs Open",v:op,c:"#f5a800"}
  ].map(s=>sumCard(s.v,s.l,s.c)).join("");
  $("tech-board").innerHTML=`<div style="overflow-x:auto;-webkit-overflow-scrolling:touch;"><table class="dt" style="min-width:640px;"><thead><tr><th>NAME</th><th>ROLE</th><th>SPECIALISATION</th><th>WOs DONE</th><th>WOs OPEN</th><th>HRS</th><th>STATUS</th></tr></thead><tbody>${D.techs.map(t=>`<tr><td style="font-weight:700;">${t.Name}</td><td style="color:var(--muted);">${t.Role}</td><td>${t.Specialisation}</td><td class="mono" style="color:var(--green);">${t.WOsCompleted}</td><td class="mono" style="color:var(--amber);">${t.WOsOpen}</td><td class="mono">${t.HoursLogged}h</td><td>${pl(t.Status,t.Status==="ON SHIFT"?"#00e087":"#5a6e8a")}</td></tr>`).join("")}</tbody></table></div>`;
  const mx=Math.max(...D.techs.map(t=>Number(t.WOsCompleted||0)))||1;
  $("tech-wl").innerHTML=D.techs.map(t=>{
    const pct=Math.round((Number(t.WOsCompleted||0)/mx)*100);
    const col=pct>=80?"#00e087":pct>=50?"#00c8f0":"#f5a800";
    return `<div style="margin-bottom:12px;"><div style="display:flex;justify-content:space-between;margin-bottom:3px;"><span style="font-size:12px;font-weight:600;">${t.Name}</span><span style="font-size:10px;color:${col};font-family:'JetBrains Mono',monospace;">${t.WOsCompleted} WOs | ${t.HoursLogged}h</span></div><div class="bt"><div class="bf" style="width:${pct}%;background:${col};"></div></div><div style="font-size:9px;color:var(--muted);margin-top:2px;">${t.Specialisation}</div></div>`;
  }).join("");
}
function buildSpares(){const so=D.spares.filter(s=>s.Status==="STOCK-OUT").length;const bm=D.spares.filter(s=>Number(s.CurrentStock)<Number(s.MinStock)).length;const tv=D.spares.reduce((s,x)=>s+Number(x.CurrentStock||0)*Number(x.UnitCost||0),0);$("spare-summary").innerHTML=[{l:"Stock-Outs",v:so,c:"#ff3355",d:"Order immediately"},{l:"Below Minimum",v:bm,c:"#f5a800",d:"Reorder required"},{l:"Total Stock Value",v:"$"+tv.toLocaleString(),c:"#00e087",d:"Current inventory"}].map(s=>`<div class="card-sm" style="border-color:${s.c!=="#00e087"&&s.v>0?s.c+"44":"var(--border)"}"><div style="font-size:32px;font-weight:900;color:${s.c};">${s.v}</div><div style="font-size:11px;font-weight:600;margin:4px 0 2px;">${s.l}</div><div style="font-size:9px;color:var(--muted);">${s.d}</div></div>`).join("");const al=D.spares.filter(s=>Number(s.CurrentStock)<Number(s.MinStock));$("spare-alerts").innerHTML=al.length?al.map(s=>{const col=s.Status==="STOCK-OUT"?"#ff3355":"#f5a800";return `<div style="display:flex;justify-content:space-between;align-items:center;padding:13px 0;border-bottom:1px solid var(--border);flex-wrap:wrap;gap:8px;"><div style="display:flex;align-items:center;"><div style="width:4px;height:44px;border-radius:2px;background:${col};margin-right:12px;flex-shrink:0;"></div><div><div style="font-size:12px;font-weight:600;">${s.Name}</div><div style="font-size:9px;color:var(--muted);font-family:'JetBrains Mono',monospace;">Part: ${s.PartNumber} | Class ${s.Class} | ${s.Supplier||"TBD"} | Lead: ${s.LeadTimeDays||"?"} day(s)</div></div></div><div style="display:flex;gap:18px;align-items:center;"><div style="text-align:center;"><div style="font-size:18px;font-weight:900;color:${Number(s.CurrentStock)===0?"#ff3355":"#f5a800"};">${s.CurrentStock}</div><div style="font-size:8px;color:var(--muted);font-family:'JetBrains Mono',monospace;">IN STOCK</div></div><div style="text-align:center;"><div style="font-size:18px;font-weight:900;color:var(--muted);">${s.MinStock}</div><div style="font-size:8px;color:var(--muted);font-family:'JetBrains Mono',monospace;">MINIMUM</div></div>${pl(s.Status,col)}<button class="btn-sm">RAISE PO ↗</button></div></div>`;}).join(""):'<div style="font-size:12px;color:var(--green);padding:16px;text-align:center;">✓ All critical spares above minimum stock</div>';$("cost-bars").innerHTML=D.costs.map(c=>{const ov=Number(c.Actual)>Number(c.Budget);return `<div style="margin-bottom:12px;"><div style="display:flex;justify-content:space-between;margin-bottom:3px;"><span style="font-size:11px;">${c.Category}</span><span style="font-size:10px;font-family:'JetBrains Mono',monospace;font-weight:700;color:${ov?"#ff3355":"#00e087"};">$${Number(c.Actual).toLocaleString()} / $${Number(c.Budget).toLocaleString()}${ov?" ▲":" ✓"}</span></div><div style="position:relative;height:6px;background:var(--border);border-radius:3px;overflow:hidden;"><div style="height:100%;width:${Math.min(Number(c.Actual)/Number(c.Budget)*100,100)}%;background:${ov?"#ff3355":Number(c.Actual)/Number(c.Budget)>0.85?"#f5a800":"#00e087"};border-radius:3px;"></div></div></div>`;}).join("");$("spares-tbl").innerHTML=D.spares.map(s=>`<tr><td style="font-weight:600;">${s.Name}</td><td class="mono" style="color:var(--muted);">${s.PartNumber}</td><td>${pl("CLASS "+s.Class,s.Class==="A"?"#ff3355":s.Class==="B"?"#f5a800":s.Class==="X"?"#a855f7":"#5a6e8a")}</td><td class="mono" style="text-align:center;">${s.CurrentStock}</td><td class="mono" style="text-align:center;">${s.MinStock}</td><td class="mono">$${Number(s.UnitCost).toLocaleString()}</td><td class="mono" style="color:var(--accent);font-weight:700;">$${(Number(s.CurrentStock)*Number(s.UnitCost)).toLocaleString()}</td><td>${pl(s.Status||"OK",s.Status==="STOCK-OUT"?"#ff3355":s.Status==="BELOW MIN"?"#f5a800":"#00e087")}</td></tr>`).join("");}
function calcROI(){const ch=Number($("roi-ch")?.value||8000),br=Number($("roi-br")?.value||30),ah=Number($("roi-ah")?.value||6),red=Number($("roi-red")?.value||40)/100,ic=Number($("roi-ic")?.value||15000),rc2=Number($("roi-rc")?.value||5000);const cc=ch*br*ah,sh=Math.round(br*ah*red),ss=Math.round(cc*red),net=ss-rc2,pb=net>0?Math.round((ic/net)*12):999;const sv=$("roi-savings");if(sv)sv.textContent="$"+net.toLocaleString();const hv=$("roi-hours");if(hv)hv.textContent=sh+" hrs";const pv=$("roi-payback");if(pv)pv.textContent=pb<=24?pb+" months":"> 2 years";const st=$("roi-stmt");if(st)st.innerHTML=`<div style="background:var(--surface2);border:1px solid var(--border2);border-radius:10px;padding:18px;line-height:1.9;font-size:12px;"><div style="font-size:13px;font-weight:700;color:var(--accent);margin-bottom:10px;">SMARTIS Reliability Program — ROI Justification — Deeprock Ghana Limited</div><p>Deeprock Ghana Limited operates a mixed fleet of RC drilling rigs (requiring external compressors), diamond drill rigs (using bean pump water circulation systems), and standalone external compressors. The fleet is currently experiencing approximately <strong style="color:var(--amber);">${br} unplanned breakdown events per year</strong>, averaging <strong style="color:var(--amber);">${ah} hours</strong> per event at a cost of <strong style="color:var(--red);">$${ch.toLocaleString()} per hour</strong>.</p><br><p>This represents an <strong style="color:var(--red);">annual unplanned downtime cost of $${cc.toLocaleString()}</strong>. Critically, compressor failures create cascade events — when an external compressor goes down, the connected RC rig also stops drilling, compounding the impact. Similarly, bean pump failures on DD rigs halt water circulation and stop all drilling operations immediately.</p><br><p>The SMARTIS Reliability Program addresses all three asset types through improved PM compliance, structured root cause analysis, predictive monitoring and real-time fleet visibility. A projected <strong style="color:var(--green);">${Math.round(red*100)}% reduction</strong> in breakdown frequency delivers:</p><br><div style="display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin:10px 0;"><div style="background:rgba(0,224,135,.08);border:1px solid rgba(0,224,135,.2);border-radius:8px;padding:12px;text-align:center;"><div style="font-size:18px;font-weight:900;color:var(--green);">${sh} hrs</div><div style="font-size:9px;color:var(--muted);">Downtime hours prevented per year</div></div><div style="background:rgba(0,224,135,.08);border:1px solid rgba(0,224,135,.2);border-radius:8px;padding:12px;text-align:center;"><div style="font-size:18px;font-weight:900;color:var(--green);">$${net.toLocaleString()}</div><div style="font-size:9px;color:var(--muted);">Net annual savings after program costs</div></div><div style="background:rgba(168,85,247,.08);border:1px solid rgba(168,85,247,.2);border-radius:8px;padding:12px;text-align:center;"><div style="font-size:18px;font-weight:900;color:var(--purple);">${pb} months</div><div style="font-size:9px;color:var(--muted);">Full program payback period</div></div></div><br><p>At an implementation cost of <strong>$${ic.toLocaleString()}</strong> and annual running cost of <strong>$${rc2.toLocaleString()}</strong>, SMARTIS delivers a <strong style="color:var(--green);">first-year return of ${Math.round(net/ic*100)}%</strong>. Prepared by Shaddy, Maintenance Officer — Deeprock Ghana Limited.</p></div>`;}
function buildShift(){const dr=D.rc.filter(r=>r.Status==="DRILLING").length+D.dd.filter(d=>d.Status==="DRILLING").length;const dn=D.wo.filter(w=>w.Status==="COMPLETED").length;const op=D.wo.filter(w=>w.Status!=="COMPLETED").length;const hi=D.wo.filter(w=>w.Priority==="HIGH"&&w.Status!=="COMPLETED").length;$("shift-kpis").innerHTML=[{l:"Rigs Drilling",v:dr,c:"#00e087"},{l:"WOs Completed",v:dn,c:"#00c8f0"},{l:"Outstanding WOs",v:op,c:"#f5a800"},{l:"High Priority",v:hi,c:"#ff3355"}].map(s=>sumCard(s.v,s.l,s.c)).join("");$("shift-done").innerHTML=D.wo.filter(w=>w.Status==="COMPLETED").map(w=>`<div style="padding:7px 0;border-bottom:1px solid var(--border);"><div style="display:flex;justify-content:space-between;"><span class="mono" style="font-size:9px;color:var(--accent);">${w.ID}</span>${pl("COMPLETED","#00e087")}</div><div style="font-size:11px;margin-top:2px;">${w.AssetID} — ${w.Description.substring(0,45)}</div><div style="font-size:9px;color:var(--muted);">Tech: ${w.Technician} | ${w.Hours}h</div></div>`).join("")||'<div style="font-size:11px;color:var(--muted);padding:12px 0;">No completed WOs</div>';$("shift-out").innerHTML=D.wo.filter(w=>w.Status!=="COMPLETED").map(w=>`<div style="padding:7px 0;border-bottom:1px solid var(--border);"><div style="display:flex;justify-content:space-between;">${pl(w.Priority,pC(w.Priority))} ${pl(w.Status,wSC(w.Status))}</div><div style="font-size:11px;margin-top:2px;"><span class="mono" style="color:var(--accent);">${w.ID}</span> — ${w.AssetID} — ${w.Description.substring(0,40)}</div><div style="font-size:9px;color:var(--muted);">Tech: ${w.Technician} | Due: ${w.DueDate}</div></div>`).join("")||'<div style="font-size:11px;color:var(--green);padding:12px 0;">All WOs complete ✓</div>';$("shift-fleet").innerHTML=[...all(),...(D.lv||[])].map(a=>{const isLV=(D.lv||[]).includes(a);const t=isLV?"LV":typeOf(a);return `<div style="display:flex;align-items:center;justify-content:space-between;padding:6px 0;border-bottom:1px solid var(--border);"><div style="display:flex;align-items:center;gap:7px;"><div style="width:6px;height:6px;border-radius:50%;background:${isLV?lvStatusCol(a.Status):sC(a.Status)};"></div><div><div style="font-size:11px;font-weight:600;">${a.Name}</div><div style="font-size:9px;color:var(--muted);">${a.Site}</div></div></div><div style="display:flex;align-items:center;gap:5px;">${pl(t,aTC(t))} ${pl(a.Status,isLV?lvStatusCol(a.Status):sC(a.Status))}</div></div>`;}).join("");}

// ══ THEME TOGGLE ═══════════════════════════════════════════════════════════
function toggleTheme(){
  const b=document.body,btn=$("theme-btn");
  b.classList.toggle("light-mode");
  const isL=b.classList.contains("light-mode");
  const lbl=$("theme-label");
  const icon=btn.querySelector("i");
  if(lbl) lbl.textContent = isL ? "LIGHT" : "DARK";
  if(icon) icon.className = isL ? "bi bi-sun-fill" : "bi bi-moon-stars";
  localStorage.setItem("smartis-theme", isL ? "light" : "dark");
}
(function(){
  try {
    const t = localStorage.getItem("smartis-theme");
    if (t === "light") {
      document.body.classList.add("light-mode");
      const b = $("theme-btn");
      if (b) {
        const icon = b.querySelector("i");
        const lbl = $("theme-label");
        if (icon) icon.className = "bi bi-sun-fill";
        if (lbl) lbl.textContent = "LIGHT";
      }
    }
  } catch(e) {}
})();

// ══ NOTIFICATIONS ══════════════════════════════════════════════════════════
let notifOpen=false;
function toggleNotifPanel(){
  notifOpen=!notifOpen;
  const p=$("notif-panel");
  if(p)p.style.display=notifOpen?"block":"none";
  if(notifOpen)buildNotifPanel();
}
function buildNotifPanel(){
  const items=[];
  all().filter(a=>a.Status==="BREAKDOWN"||a.Status==="DOWN").forEach(a=>{items.push({l:"CRITICAL",m:`${a.Name} — OFFLINE. Immediate action required.`,c:"#ff3355",t:new Date().toLocaleTimeString()});});
  D.comp.filter(c=>c.Status==="DOWN").forEach(c=>{const rc=D.rc.find(r=>r.ConnectedCompressor===c.ID);if(rc)items.push({l:"CASCADE",m:`${c.Name} DOWN → ${rc.Name} has no air. Both assets affected.`,c:"#ff3355",t:new Date().toLocaleTimeString()});});
  all().filter(a=>Number(a.HrsToNextPM)<0).forEach(a=>{items.push({l:"PM OVERDUE",m:`${a.Name} — PM overdue by ${Math.abs(a.HrsToNextPM)} hrs`,c:"#f5a800",t:new Date().toLocaleTimeString()});});
  (D.lv||[]).filter(v=>daysUntil(v.RoadworthyExpiry)<0||daysUntil(v.InsuranceExpiry)<0||daysUntil(v.FireExtExpiry)<0||daysUntil(v.FirstAidExpiry)<0).forEach(v=>{items.push({l:"LV COMPLIANCE",m:`${v.Name} — compliance document expired`,c:"#ff3355",t:new Date().toLocaleTimeString()});});
  D.spares.filter(s=>s.Status==="STOCK-OUT").forEach(s=>{items.push({l:"STOCK-OUT",m:`${s.Name} (${s.PartNumber}) — Class ${s.Class}. Lead: ${s.LeadTimeDays}d`,c:"#ff3355",t:new Date().toLocaleTimeString()});});
  all().filter(a=>Number(a.HrsToNextPM)>=0&&Number(a.HrsToNextPM)<=50).forEach(a=>{items.push({l:"PM DUE SOON",m:`${a.Name} — PM due in ${a.HrsToNextPM} hrs`,c:"#f5a800",t:new Date().toLocaleTimeString()});});
  (D.lv||[]).filter(v=>{const d=[v.RoadworthyExpiry,v.InsuranceExpiry,v.FireExtExpiry,v.FirstAidExpiry];return d.some(x=>daysUntil(x)>=0&&daysUntil(x)<=30);}).forEach(v=>{items.push({l:"EXPIRING SOON",m:`${v.Name} — compliance document expiring within 30 days`,c:"#f5a800",t:new Date().toLocaleTimeString()});});
  const cnt=$("notif-count");if(cnt){cnt.textContent=items.length>0?items.length:"";cnt.style.display=items.length>0?"flex":"none";}
  const list=$("notif-list");
  if(list)list.innerHTML=items.length?items.map(i=>`<div style="display:flex;gap:10px;padding:10px 14px;border-bottom:1px solid var(--border);"><div style="flex-shrink:0;"><div style="font-size:8px;font-weight:700;font-family:'JetBrains Mono',monospace;color:${i.c};">${i.l}</div><div style="font-size:8px;color:var(--muted);margin-top:2px;">${i.t}</div></div><div style="font-size:11px;line-height:1.5;">${i.m}</div></div>`).join(""):'<div style="padding:20px;text-align:center;font-size:11px;color:var(--green);">✓ No active alerts</div>';
}
function updateNotifCount(){
  const bd=all().filter(a=>a.Status==="BREAKDOWN"||a.Status==="DOWN").length;
  const po=all().filter(a=>Number(a.HrsToNextPM)<0).length;
  const so=D.spares.filter(s=>s.Status==="STOCK-OUT").length;
  const lvExp=(D.lv||[]).filter(v=>daysUntil(v.RoadworthyExpiry)<0||daysUntil(v.InsuranceExpiry)<0||daysUntil(v.FireExtExpiry)<0||daysUntil(v.FirstAidExpiry)<0).length;
  const total=bd+po+so+lvExp;
  const cnt=$("notif-count");if(cnt){cnt.textContent=total>0?total:"";cnt.style.display=total>0?"flex":"none";}
}

// ══ WO MODAL ════════════════════════════════════════════════════════════════
let woCounter=9000;
function openWOModal(){
  const m=$("wo-modal");if(!m)return;
  m.style.display="flex";
  const existing=D.wo.map(w=>parseInt((w.ID||"WO-0").replace("WO-",""))||0);
  woCounter=Math.max(...existing,woCounter)+1;
  $("wo-id").value="WO-"+woCounter;
  const today=new Date().toISOString().split("T")[0];
  $("wo-raised").value=today;
  const sel=$("wo-asset");sel.innerHTML="";
  [...D.rc.map(a=>({id:a.ID,t:"RC"})),...D.dd.map(a=>({id:a.ID,t:"DD"})),...D.comp.map(a=>({id:a.ID,t:"COMPRESSOR"})),...(D.lv||[]).map(a=>({id:a.ID,t:"LV"}))].forEach(a=>{
    const o=document.createElement("option");o.value=a.id;o.dataset.type=a.t;o.textContent=a.id+" ("+a.t+")";sel.appendChild(o);
  });
  woAssetChange();
  const ts=$("wo-tech");ts.innerHTML="";
  D.techs.forEach(t=>{const o=document.createElement("option");o.value=t.Name;o.textContent=t.Name+" — "+t.Role;ts.appendChild(o);});
  $("wo-preview").style.display="none";
}
function closeWOModal(){const m=$("wo-modal");if(m)m.style.display="none";}
function woAssetChange(){
  const sel=$("wo-asset");
  const opt=sel.options[sel.selectedIndex];
  if(opt)$("wo-atype").value=opt.dataset.type||"";
}
function previewWO(){
  const row=[
    $("wo-id").value,$("wo-asset").value,$("wo-atype").value,$("wo-type").value,
    $("wo-pri").value,$("wo-sys").value,$("wo-desc").value,$("wo-tech").value,
    $("wo-raised").value,$("wo-due").value,$("wo-hrs").value,$("wo-stat").value,
    $("wo-rc").value,$("wo-cost").value
  ].join(" | ");
  const p=$("wo-preview");
  p.style.display="block";
  p.innerHTML='<div style="font-size:9px;color:var(--muted);margin-bottom:5px;">PASTE THIS ROW INTO YOUR GOOGLE SHEET — WorkOrders tab:</div>'+row.split(" | ").map((v,i)=>{
    const hdrs=["ID","AssetID","AssetType","Type","Priority","System","Description","Technician","RaisedDate","DueDate","Hours","Status","RootCause","DowntimeCost"];
    return `<span style="display:inline-block;background:var(--surface);border:1px solid var(--border2);border-radius:3px;padding:2px 6px;margin:2px;font-size:9px;"><span style="color:var(--muted);">${hdrs[i]}: </span>${v||"—"}</span>`;
  }).join("");
}
function copyWORow(){
  const row=[
    $("wo-id").value,$("wo-asset").value,$("wo-atype").value,$("wo-type").value,
    $("wo-pri").value,$("wo-sys").value,$("wo-desc").value,$("wo-tech").value,
    $("wo-raised").value,$("wo-due").value,$("wo-hrs").value,$("wo-stat").value,
    $("wo-rc").value,$("wo-cost").value
  ].join("\t");
  navigator.clipboard.writeText(row).then(()=>{
    const btn=event.target;const orig=btn.textContent;
    btn.textContent="✓ COPIED!";btn.style.background="linear-gradient(135deg,#00e087,#00b060)";
    setTimeout(()=>{btn.textContent=orig;btn.style.background="";},2200);
  });
  previewWO();
}

// ══ CASCADE VISUAL ══════════════════════════════════════════════════════════
function buildCascadeVisual(){
  const el=$("cascade-visual");if(!el)return;
  const pairs=D.rc.map(r=>{const comp=D.comp.find(c=>c.ID===r.ConnectedCompressor);return{rc:r,comp:comp};});
  el.innerHTML=`<div style="display:flex;gap:20px;flex-wrap:wrap;justify-content:center;">${pairs.map(({rc,comp})=>{
    const ok=comp&&comp.Status!=="DOWN";
    const lineCol=ok?"#00e087":"#ff3355";
    const lineLabel=ok?"AIR SUPPLY OK":"NO AIR — CASCADE";
    return `<div style="display:flex;flex-direction:column;align-items:center;gap:0;min-width:160px;">
      <div style="background:var(--surface2);border:2px solid ${comp?lineCol+"66":"#5a6e8a"};border-radius:10px;padding:10px 14px;width:100%;text-align:center;">
        <div style="font-size:8px;font-weight:700;color:var(--comp);font-family:'JetBrains Mono',monospace;letter-spacing:1px;margin-bottom:4px;">COMPRESSOR</div>
        <div style="font-size:10px;font-weight:700;">${comp?comp.Name:"UNLINKED"}</div>
        <div style="font-size:8px;color:var(--muted);margin:2px 0;">${comp?comp.Make+" "+comp.Model:""}</div>
        ${comp?`<div style="display:inline-block;background:${lineCol}22;border:1px solid ${lineCol}55;border-radius:4px;padding:1px 6px;font-size:8px;font-weight:700;color:${lineCol};font-family:'JetBrains Mono',monospace;margin-top:4px;">${comp.Status}</div>`:""}
      </div>
      <div style="display:flex;flex-direction:column;align-items:center;gap:0;width:100%;">
        <div style="width:2px;height:12px;background:${lineCol};"></div>
        <div style="background:${lineCol}22;border:1px solid ${lineCol}55;border-radius:4px;padding:2px 8px;font-size:7px;font-weight:700;color:${lineCol};font-family:'JetBrains Mono',monospace;white-space:nowrap;">${lineLabel}</div>
        <div style="width:0;height:0;border-left:6px solid transparent;border-right:6px solid transparent;border-top:8px solid ${lineCol};"></div>
      </div>
      <div style="background:var(--surface2);border:2px solid ${fade(sC(rc.Status),40)};border-radius:10px;padding:10px 14px;width:100%;text-align:center;">
        <div style="font-size:8px;font-weight:700;color:var(--rc);font-family:'JetBrains Mono',monospace;letter-spacing:1px;margin-bottom:4px;">RC DRILLING RIG</div>
        <div style="font-size:10px;font-weight:700;">${rc.Name}</div>
        <div style="font-size:8px;color:var(--muted);margin:2px 0;">${rc.Site}</div>
        <div style="display:inline-block;background:${fade(sC(rc.Status),16)};border:1px solid ${fade(sC(rc.Status),34)};border-radius:4px;padding:1px 6px;font-size:8px;font-weight:700;color:${sC(rc.Status)};font-family:'JetBrains Mono',monospace;margin-top:4px;">${rc.Status}</div>
      </div>
    </div>`;
  }).join("")}</div>`;
}

// ══ COST TRACKER ════════════════════════════════════════════════════════════
function buildCostTracker(){
  const costs=D.costs||[];if(!costs.length)return;
  const totalBudget=costs.reduce((s,c)=>s+Number(c.Budget||0),0);
  const totalActual=costs.reduce((s,c)=>s+Number(c.Actual||0),0);
  const variance=totalActual-totalBudget;
  const overruns=costs.filter(c=>Number(c.Actual)>Number(c.Budget)).length;
  const cp=$("cost-period-badge");if(cp)cp.textContent="PERIOD: "+(costs[0]?.Period||"Current");
  $("cost-big").innerHTML=[
    {l:"Total Budget",v:"$"+totalBudget.toLocaleString(),c:"#00c8f0"},
    {l:"Total Actual Spend",v:"$"+totalActual.toLocaleString(),c:totalActual>totalBudget?"#ff3355":"#00e087"},
    {l:"Variance",v:(variance>=0?"+$":"-$")+Math.abs(variance).toLocaleString(),c:variance>0?"#ff3355":"#00e087"},
    {l:"Categories Over Budget",v:overruns,c:overruns>0?"#ff3355":"#00e087"},
  ].map(k=>sumCard(k.v,k.l,k.c)).join("");
  $("cost-avb").innerHTML=costs.map(c=>{
    const pct=Math.min(Math.round(Number(c.Actual)/Number(c.Budget)*100),150);
    const ov=Number(c.Actual)>Number(c.Budget);
    const col=ov?"#ff3355":pct>85?"#f5a800":"#00e087";
    return `<div style="margin-bottom:14px;"><div style="display:flex;justify-content:space-between;margin-bottom:3px;"><span style="font-size:11px;">${c.Category}</span><span style="font-size:10px;font-family:'JetBrains Mono',monospace;font-weight:700;color:${col};">$${Number(c.Actual).toLocaleString()} / $${Number(c.Budget).toLocaleString()} ${ov?"▲ OVER":"✓"}</span></div><div style="position:relative;height:8px;background:var(--border);border-radius:4px;overflow:hidden;"><div style="height:100%;width:${Math.min(pct,100)}%;background:${col};border-radius:4px;"></div></div><div style="font-size:9px;color:var(--muted);margin-top:2px;">${pct}% of budget used</div></div>`;
  }).join("");
  const rcShare=Math.round(totalActual*0.38),ddShare=Math.round(totalActual*0.27),compShare=Math.round(totalActual*0.23),lvShare=Math.round(totalActual*0.12);
  $("cost-by-type").innerHTML=[{l:"RC Drilling Rigs",v:rcShare,c:"var(--info)"},{l:"DD Rigs",v:ddShare,c:"var(--info)"},{l:"Compressors",v:compShare,c:"var(--info)"},{l:"Light Vehicles",v:lvShare,c:"var(--info)"}].map(b=>`<div style="margin-bottom:10px;"><div style="display:flex;justify-content:space-between;margin-bottom:3px;"><span style="font-size:11px;">${b.l}</span><span style="font-size:10px;font-weight:700;color:${b.c};font-family:'JetBrains Mono',monospace;">$${Number(b.v).toLocaleString()} (${Math.round(b.v/totalActual*100)}%)</span></div><div class="bt"><div class="bf" style="width:${Math.round(b.v/totalActual*100)}%;background:${b.c};"></div></div></div>`).join("");
  const pmHrs=D.wo.filter(w=>w.Type==="PM").reduce((s,w)=>s+Number(w.Hours||0),0);
  const rxHrs=D.wo.filter(w=>w.Type==="REACTIVE").reduce((s,w)=>s+Number(w.Hours||0),0);
  const cmHrs=D.wo.filter(w=>w.Type==="CORRECTIVE").reduce((s,w)=>s+Number(w.Hours||0),0);
  const pdHrs=D.wo.filter(w=>w.Type==="PdM").reduce((s,w)=>s+Number(w.Hours||0),0);
  const tHrs=pmHrs+rxHrs+cmHrs+pdHrs||1;
  $("cost-maint-split").innerHTML=[{l:"Preventive (PM)",v:pmHrs,c:"#00e087"},{l:"Reactive",v:rxHrs,c:"#ff3355"},{l:"Corrective",v:cmHrs,c:"#00c8f0"},{l:"Predictive (PdM)",v:pdHrs,c:"#a855f7"}].map(b=>`<div style="margin-bottom:8px;"><div style="display:flex;justify-content:space-between;margin-bottom:2px;"><span style="font-size:11px;">${b.l}</span><span style="font-size:10px;font-weight:700;color:${b.c};font-family:'JetBrains Mono',monospace;">${b.v}h (${Math.round(b.v/tHrs*100)}%)</span></div><div class="bt"><div class="bf" style="width:${Math.round(b.v/tHrs*100)}%;background:${b.c};"></div></div></div>`).join("");
  $("cost-cph").innerHTML=`<div style="display:flex;gap:5px;align-items:flex-end;height:110px;overflow-x:auto;">${[...all(),...(D.lv||[])].map(a=>{
    const isLV=(D.lv||[]).includes(a);
    const h=Math.min(Number(a.CostPerHour||a.CostPerKm||0),100);
    const avgCPH=all().reduce((s,x)=>s+Number(x.CostPerHour||0),0)/all().length;
    const over=Number(a.CostPerHour||a.CostPerKm||0)>avgCPH;
    const col=over?"var(--danger)":"var(--info)";
    return `<div style="min-width:44px;flex:1;display:flex;flex-direction:column;align-items:center;gap:3px;"><div style="font-size:8px;color:${col};font-family:'JetBrains Mono',monospace;font-weight:700;">${isLV?"$"+a.CostPerKm+"/km":"$"+a.CostPerHour+"/h"}</div><div style="width:100%;height:${h}px;background:var(--surface3);border:1px solid var(--border2);border-radius:3px 3px 0 0;display:flex;align-items:flex-end;"><div style="width:100%;height:${h}px;background:${col};border-radius:3px 3px 0 0;"></div></div><div style="font-size:7px;color:var(--muted);text-align:center;font-family:'JetBrains Mono',monospace;">${a.ID.replace("DGL-","")}</div></div>`;
  }).join("")}</div><div style="font-size:9px;color:var(--muted);margin-top:8px;">Red bars = above fleet average cost per hour.</div>`;
  const ov=costs.filter(c=>Number(c.Actual)>Number(c.Budget));
  $("cost-overrun").innerHTML=ov.length?ov.map(c=>{
    const diff=Number(c.Actual)-Number(c.Budget);const pct=Math.round(diff/Number(c.Budget)*100);
    return `<div style="display:flex;justify-content:space-between;align-items:center;padding:12px 0;border-bottom:1px solid var(--border);"><div><div style="font-size:12px;font-weight:700;">${c.Category}</div><div style="font-size:10px;color:var(--muted);">Budget: $${Number(c.Budget).toLocaleString()} | Actual: $${Number(c.Actual).toLocaleString()}</div></div><div style="text-align:right;"><div style="font-size:16px;font-weight:900;color:#ff3355;">+$${diff.toLocaleString()}</div><div style="font-size:9px;color:#ff3355;font-family:'JetBrains Mono',monospace;">${pct}% OVER</div></div></div>`;
  }).join(""):'<div style="font-size:12px;color:var(--green);padding:16px;text-align:center;">✓ All categories within budget</div>';
}


function showTab(id,btn){document.querySelectorAll("[id^='tab-']").forEach(e=>e.classList.add("hidden"));const el=$("tab-"+id);if(el){el.classList.remove("hidden");el.classList.add("fade-in");}document.querySelectorAll(".tab").forEach(t=>t.classList.remove("active"));if(btn)btn.classList.add("active");}
(function(){
  try{
    if(localStorage.getItem("smartis-file-data")){
      const b=$("reload-file-btn");
      if(b) b.style.display="";
    }
  }catch(e){}
})();

// Only auto-resume if we're actually going to the dashboard.
// If the auth-skip block above left the setup modal visible, we're done.
(function(){
  const setupVisible = !document.getElementById('setup-modal')?.classList.contains('hidden');
  if(setupVisible) return;                 // user must pick a data source first
  resumeLive();
})();

// ─── Animated number count-up on Executive tab ───
(function(){
  const origShowTab = window.showTab;
  window.showTab = function(id, btn) {
    origShowTab(id, btn);
    if (id === 'exec') {
      setTimeout(() => {
        document.querySelectorAll('#exec-big > div > div:first-child').forEach(el => {
          const text = el.textContent;
          const match = text.match(/^([$]?)([\d,]+)(.*)$/);
          if (!match) return;
          const prefix = match[1], numStr = match[2].replace(/,/g,''), suffix = match[3];
          const target = parseInt(numStr, 10);
          if (isNaN(target)) return;
          const dur = 800;
          const start = performance.now();
          const from = 0;
          function tick(now){
            const p = Math.min((now - start) / dur, 1);
            const ease = 1 - Math.pow(1 - p, 3);
            const v = Math.floor(from + (target - from) * ease);
            el.textContent = prefix + v.toLocaleString() + suffix;
            if (p < 1) requestAnimationFrame(tick);
          }
          requestAnimationFrame(tick);
        });
      }, 60);
    }
  };
})();

// ─── Bridge window.print() to the native Android print service ───
(function waitForCapacitor() {
  if (window.Capacitor && window.Capacitor.Plugins && window.Capacitor.Plugins.Printer) {
    window.print = async function() {
      try {
        await window.Capacitor.Plugins.Printer.printWebView({ name: 'SMARTIS Report' });
      } catch (e) {
        console.error('Print failed:', e);
        try { await window.Capacitor.Plugins.Printer.print(); }
        catch (e2) { console.error('Print also failed:', e2); }
      }
    };
    console.log('✅ Native printer bridge installed');
  } else {
    if (!waitForCapacitor._tries) waitForCapacitor._tries = 0;
    if (waitForCapacitor._tries++ < 50) setTimeout(waitForCapacitor, 100);
    else console.warn('⚠️ Capacitor Printer plugin never became available — print will use the browser fallback.');
  }
})();

// ─── Safe-area guard: keeps content out of status bar & Android nav buttons ───
(function(){
  const root = document.documentElement;
  function applySafeAreas(){
    const probe = document.createElement('div');
    probe.style.cssText = 'position:fixed;top:0;left:0;visibility:hidden;pointer-events:none;' +
      'padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px);';
    document.body.appendChild(probe);
    const cs = getComputedStyle(probe);
    let top = parseFloat(cs.paddingTop) || 0;
    let bottom = parseFloat(cs.paddingBottom) || 0;
    probe.remove();

    const isAppWebView = !!window.Capacitor || /; wv\)/.test(navigator.userAgent);
    const isPhone = window.matchMedia('(max-width: 992px)').matches;
    if (isAppWebView && isPhone) {
      if (!top)    top    = 32;
      if (!bottom) bottom = 48;
    }
    root.style.setProperty('--safe-top',    top + 'px');
    root.style.setProperty('--safe-bottom', bottom + 'px');
  }
  applySafeAreas();
  window.addEventListener('resize', applySafeAreas);
  window.addEventListener('orientationchange', () => setTimeout(applySafeAreas, 250));
})();

// ─── Sign-out confirmation ───
function askLogout(){
  document.body.classList.remove('sb-open');
  var m=document.getElementById('logout-modal');if(m)m.classList.remove('hidden');
}
function closeLogout(){
  var m=document.getElementById('logout-modal');if(m)m.classList.add('hidden');
}

// ══════════════════════════════════════════════════════════════════════════
// ══ SMARTIS LOGIN — SUPABASE AUTH ═════════════════════════════════════════
// ══════════════════════════════════════════════════════════════════════════
var LG_KEY='smartis-auth';
var LG_USER_KEY='smartis-user-name';

function smartisLogout(){
  try{
    sessionStorage.removeItem(LG_KEY);
    localStorage.removeItem(LG_KEY);
    localStorage.removeItem(LG_USER_KEY);
  }catch(e){}
  location.reload();
}

(function(){
  var L=document.getElementById('login-screen');if(!L)return;
  var g=function(i){return document.getElementById(i);};

  // Already signed in? Skip the login screen.
  // Already signed in? Skip the login screen — but only reveal the dashboard
// if there is actually data to show. Otherwise show the setup/load page.
try{
  if(sessionStorage.getItem(LG_KEY)==='1' || localStorage.getItem(LG_KEY)==='1'){

    // Decide whether we already have a data source:
    //   – a saved live Google Sheet link AND cached live data, OR
    //   – a previously uploaded file
    var hasData = false;
    try{
      var savedLink  = localStorage.getItem('smartis-sheet-link') ||
                       localStorage.getItem('smartis-sid');
      var cachedLive = localStorage.getItem('smartis-live-data');
      var cachedFile = localStorage.getItem('smartis-file-data');
      var src        = localStorage.getItem('smartis-source');
      if (src !== 'OTHER' && savedLink && cachedLive) hasData = true;
      if (cachedFile) hasData = true;
    }catch(e){}

    L.remove();
    document.body.classList.remove('lg-locked');

    var setupM = g('setup-modal');
    var mainA  = g('main-app');

    if (hasData){
      // Dashboard path — resumeLive() below will hydrate it.
      if (setupM) setupM.classList.add('hidden');
      if (mainA)  mainA.classList.remove('hidden');
    } else {
      // No data yet → show the Load Data page.
      if (mainA)  mainA.classList.add('hidden');
      if (setupM) setupM.classList.remove('hidden');
    }
    return;
  }
}catch(e){}
  document.body.classList.add('lg-locked');

  var user=g('lg-user'),pass=g('lg-pass'),card=g('lg-card'),err=g('lg-err'),btn=g('lg-btn');
  try{var ru=localStorage.getItem('smartis-user');if(ru){user.value=ru;g('lg-keep').checked=true;}}catch(e){}

  // Show / hide password
  g('lg-eye').addEventListener('click',function(){
    var show=pass.type==='password';
    pass.type=show?'text':'password';
    this.innerHTML='<i class="bi bi-eye'+(show?'-slash':'')+'"></i>';
  });

  // Caps lock warning
  function caps(e){if(e.getModifierState){g('lg-caps').classList.toggle('show',e.getModifierState('CapsLock'));}}
  pass.addEventListener('keydown',caps);pass.addEventListener('keyup',caps);
  pass.addEventListener('blur',function(){g('lg-caps').classList.remove('show');});

  function shake(msg,el){
    err.textContent=msg;
    card.classList.remove('lg-shake');void card.offsetWidth;card.classList.add('lg-shake');
    if(el)el.focus();
  }
  [user,pass].forEach(function(i){i.addEventListener('input',function(){err.textContent='';});});

  // ── SUBMIT: verify credentials against Supabase ──
  g('lg-form').addEventListener('submit', async function(e){
    e.preventDefault();
    var u=user.value.trim(), p=pass.value;
    if(!u){shake('Please enter your username',user);return;}
    if(!p){shake('Please enter your password',pass);return;}

    // Guard: Supabase must be loaded
    if(!window.sb){
      shake('Login service unavailable. Please reload the page.',pass);
      return;
    }

    btn.disabled=true;
    g('lg-btn-txt').textContent='VERIFYING…';
    g('lg-btn-ico').outerHTML='<span class="lg-spin" id="lg-btn-ico"></span>';

    try{
      // Call our Postgres function `smartis_login(username, password)`
      const { data, error } = await window.sb.rpc('smartis_login', {
        p_username: u,
        p_password: p
      });

      if (error) throw new Error(error.message || 'Login failed');
      if (!data || data.length === 0) {
        throw new Error('Invalid username or password');
      }

      // ✔ Credentials OK
      const account = data[0];

      btn.classList.add('ok');
      g('lg-btn-txt').textContent='ACCESS GRANTED';
      g('lg-btn-ico').outerHTML='<i class="bi bi-check-circle-fill" id="lg-btn-ico"></i>';

      try{
        var keep=g('lg-keep').checked;
        (keep?localStorage:sessionStorage).setItem(LG_KEY,'1');
        if(keep)localStorage.setItem('smartis-user',u);else localStorage.removeItem('smartis-user');
        localStorage.setItem(LG_USER_KEY, account.full_name || account.username);
      }catch(e){}

      setTimeout(function(){
        // Show welcome overlay with the real name
        g('lg-wname').textContent = account.full_name || account.username;
        card.classList.add('lg-hide');
        g('lg-welcome').classList.add('show');
        setTimeout(function(){
          L.classList.add('lg-out');
          document.body.classList.remove('lg-locked');
          running=false;
          setTimeout(function(){L.remove();},800);
        },1700);
      },650);

    }catch(ex){
      // Restore button
      btn.disabled=false;
      btn.classList.remove('ok');
      g('lg-btn-txt').textContent='SIGN IN';
      g('lg-btn-ico').outerHTML='<i class="bi bi-arrow-right" id="lg-btn-ico"></i>';
      shake(ex.message || 'Login failed. Check your username and password.');
    }
  });

  // Animated particle network background
  var cv=g('lg-canvas'),cx=cv.getContext('2d'),W=0,H=0,D=1,pts=[],running=true;
  function size(){
    D=window.devicePixelRatio||1;
    W=cv.width=innerWidth*D;H=cv.height=innerHeight*D;
    var n=Math.min(70,Math.floor(innerWidth*innerHeight/15000));pts=[];
    for(var i=0;i<n;i++)pts.push({x:Math.random()*W,y:Math.random()*H,vx:(Math.random()-.5)*.4*D,vy:(Math.random()-.5)*.4*D,r:(Math.random()*1.4+.6)*D});
  }
  function draw(){
    if(!running)return;
    cx.clearRect(0,0,W,H);
    var link=130*D;
    for(var i=0;i<pts.length;i++){
      var a=pts[i];a.x+=a.vx;a.y+=a.vy;
      if(a.x<0||a.x>W)a.vx*=-1;if(a.y<0||a.y>H)a.vy*=-1;
      cx.beginPath();cx.arc(a.x,a.y,a.r,0,6.283);cx.fillStyle='rgba(150,180,255,.75)';cx.fill();
      for(var j=i+1;j<pts.length;j++){
        var b=pts[j],dx=a.x-b.x,dy=a.y-b.y,d=Math.sqrt(dx*dx+dy*dy);
        if(d<link){cx.strokeStyle='rgba(120,160,255,'+(0.22*(1-d/link))+')';cx.lineWidth=D*.8;cx.beginPath();cx.moveTo(a.x,a.y);cx.lineTo(b.x,b.y);cx.stroke();}
      }
    }
    requestAnimationFrame(draw);
  }
  size();window.addEventListener('resize',size);draw();
})();


// ══════════ WEB APP EDITION ══════════
function toast(msg, type){
  var w = document.getElementById('toast-wrap');
  if(!w){ w = document.createElement('div'); w.id = 'toast-wrap'; document.body.appendChild(w); }
  var t = document.createElement('div');
  t.className = 'toast ' + (type || '');
  t.innerHTML = '<i class="bi bi-' + (type==='ok'?'check-circle-fill':type==='warn'?'exclamation-triangle-fill':'info-circle-fill') + '"></i><span>' + msg + '</span>';
  w.appendChild(t);
  setTimeout(function(){ t.classList.add('out'); setTimeout(function(){ t.remove(); }, 320); }, 3200);
}
(function(){
  var b = document.body;

  // Menu button: collapses the sidebar on desktop, slides it in on phones
  var mt = document.querySelector('.menu-toggle');
  if(mt){
    mt.removeAttribute('onclick');
    mt.addEventListener('click', function(){
      if(window.innerWidth > 992){
        b.classList.toggle('sb-collapsed');
        try{ localStorage.setItem('smartis-sb', b.classList.contains('sb-collapsed') ? '1' : '0'); }catch(e){}
      } else { b.classList.toggle('sb-open'); }
    });
  }
  try{ if(localStorage.getItem('smartis-sb') === '1') b.classList.add('sb-collapsed'); }catch(e){}

  // Choosing a tab closes the phone menu and returns to the top
  document.querySelectorAll('.sidebar .tab').forEach(function(t){
    t.addEventListener('click', function(){ b.classList.remove('sb-open'); window.scrollTo({top:0}); });
  });

  // Esc closes menus and popups
  document.addEventListener('keydown', function(e){
    if(e.key !== 'Escape') return;
    b.classList.remove('sb-open');
    var n = document.getElementById('notif-panel');
    if(n && n.style.display === 'block') toggleNotifPanel();
    closeLogout(); closeWOModal();
  });

  // Click outside the notification panel closes it
  document.addEventListener('click', function(e){
    var n = document.getElementById('notif-panel');
    if(n && n.style.display === 'block' && !n.contains(e.target) && !e.target.closest('#notif-btn')) toggleNotifPanel();
  });

  // Connection feedback
  window.addEventListener('offline', function(){ toast('You are offline. Showing the last saved data.', 'warn'); });
  window.addEventListener('online',  function(){ toast('Back online.', 'ok'); if(typeof refreshData === 'function') refreshData(); });
})();