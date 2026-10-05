const VERSION="2.3";
const MONTHS=["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"];
const STORAGE={
  timeline:"mh_v1_consumption_timeline",stock:"mh_v1_stock",prevStock:"mh_v1_prev_stock",lastStock:"mh_v1_last_stock",
  warehouse:"mh_v1_warehouse",arsenalAdd:"mh_v1_arsenal_add",arsenalOverride:"mh_v1_arsenal_override",arsenalRemoved:"mh_v1_arsenal_removed",
  weekly:"mh_v1_weekly_inventory",generalDates:"mh_v1_general_dates",generalHistory:"mh_v1_general_history",
  participants:"mh_v1_participants",orders:"mh_v1_orders",expiries:"mh_v1_expiries",stockouts:"mh_v1_stockouts",
  orderDates:"mh_v1_order_dates",medicineMeta:"mh_v1_medicine_meta",cloudCode:"mh_v2_cloud_code",cloudMigrated:"mh_v2_cloud_migrated"
};
const DEFAULT_ORDER_DATES=[];
let baseMaster=[],master=[],masterVersion="",pedidoTemplate={};
let timeline=load(STORAGE.timeline,{}),stockMap=load(STORAGE.stock,{}),prevStockMap=load(STORAGE.prevStock,{});
let warehouse=load(STORAGE.warehouse,{}),arsenalAdd=load(STORAGE.arsenalAdd,[]),arsenalOverride=load(STORAGE.arsenalOverride,{});
let arsenalRemoved=load(STORAGE.arsenalRemoved,[]),weeklyHistory=load(STORAGE.weekly,[]),generalDates=load(STORAGE.generalDates,[{date:"2026-09-01",status:"Postergado"},{date:"2026-12-19",status:"Programado"}]);
let generalHistory=load(STORAGE.generalHistory,[]),participants=load(STORAGE.participants,{tens1:"",tens2:"",tens3:"",qf:""});
let orderDates=load(STORAGE.orderDates,DEFAULT_ORDER_DATES);
let ordersHistory=load(STORAGE.orders,[]),expiries=load(STORAGE.expiries,[]),stockouts=load(STORAGE.stockouts,{});
let medicineMeta=load(STORAGE.medicineMeta,{});
let currentWeekly=[],currentOrder=[],activeQuick="",activeMedicineDetail=null;

let annualDraftOrders=[],annualDraftGeneral=[];

const SUPABASE_URL="https://upbwdypklxnxkhbyzhdu.supabase.co";
const SUPABASE_KEY="sb_publishable_NDDt53ODUhDOvmgHLsCCOg_rTLOkuap";
const INTERNAL_MH_CLOUD_CODE="MH-F87BQCC9D2";
let supabaseClient=null, cloudConnected=false, cloudCode="", cloudProductByNorm=new Map(), cloudProductById=new Map();
let supabaseJsPromise=null;


const $=id=>document.getElementById(id);
let sheetJsPromise=null,excelJsPromise=null;
function loadExternalScript(src){
  return new Promise((resolve,reject)=>{
    const s=document.createElement("script");
    s.src=src;
    s.async=true;
    s.onload=resolve;
    s.onerror=()=>reject(new Error("No se pudo cargar la librería necesaria."));
    document.head.appendChild(s);
  });
}
async function ensureSheetJS(){
  if(typeof XLSX!=="undefined")return;
  if(!sheetJsPromise)sheetJsPromise=loadExternalScript("https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js");
  await sheetJsPromise;
}
async function ensureExcelJS(){
  if(typeof ExcelJS!=="undefined")return;
  if(!excelJsPromise)excelJsPromise=loadExternalScript("https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js");
  await excelJsPromise;
}

async function ensureSupabaseJS(){
  if(window.supabase?.createClient)return;
  if(!supabaseJsPromise)supabaseJsPromise=loadExternalScript("https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2");
  await supabaseJsPromise;
}
async function makeSupabaseClient(code){
  await ensureSupabaseJS();
  return window.supabase.createClient(SUPABASE_URL,SUPABASE_KEY,{
    auth:{persistSession:false,autoRefreshToken:false},
    global:{headers:{"x-mh-code":code}}
  });
}
function setCloudBadge(text,ok=false){
  const b=$("cloudBadge"); if(!b)return;
  b.textContent=(ok?"☁️ ":"☁️ ")+text;
  b.classList.toggle("connected",ok);
}
async function verifyCloudCode(code){
  const client=await makeSupabaseClient(code);
  const {data,error}=await client.rpc("mh_access_ok");
  if(error)throw error;
  if(data!==true)throw new Error("Código de acceso incorrecto.");
  return client;
}
async function connectCloud(code,remember=true,quiet=false){
  if(!code)return false;
  try{
    setCloudBadge("Conectando…",false);
    const client=await verifyCloudCode(code);
    supabaseClient=client; cloudCode=code; cloudConnected=true;
    // v2.1: la conexión usa un código interno; no se guarda ni se solicita al usuario.
    await syncMasterToCloud();
    await migrateLocalDataIfCloudEmpty();
    await loadCloudState();
    setCloudBadge("Sincronizado",true);
    if($("cloudLoginStatus")){ $("cloudLoginStatus").className="status-strip ok"; $("cloudLoginStatus").textContent="✅ Conectado. Los datos se sincronizan en línea."; }
    if(!quiet)alert("Conexión en línea activa. Stock, consumos e inventarios quedarán disponibles desde otros dispositivos.");
    return true;
  }catch(e){
    console.error("Cloud:",e);
    cloudConnected=false; supabaseClient=null;
    setCloudBadge("Conectar nube",false);
    if($("cloudLoginStatus")){ $("cloudLoginStatus").className="status-strip warning"; $("cloudLoginStatus").textContent="⚠️ "+(e.message||"No fue posible conectar."); }
    if(!quiet)alert(e.message||"No fue posible conectar a la base en línea.");
    return false;
  }
}
function disconnectCloud(){
  cloudConnected=false; supabaseClient=null; cloudCode="";
  localStorage.removeItem(STORAGE.cloudCode);
  setCloudBadge("Conectar nube",false);
  if($("cloudLoginStatus")){ $("cloudLoginStatus").className="status-strip"; $("cloudLoginStatus").textContent="Dispositivo desconectado. Los datos locales se mantienen."; }
}
async function syncMasterToCloud(){
  if(!cloudConnected)return;
  const rows=master.map(m=>({
    establishment_code:"MH",
    local_name:m.glosaABG,
    warehouse_name:m.glosaBodega||null,
    category:m.categoria||null,
    is_active:true,
    is_controlled:!!m.controlado,
    dispensable:m.dispensableMH!==false,
    warehouse_available:m.disponibleBodega===null||m.disponibleBodega===undefined?null:!!m.disponibleBodega,
    notes:m.dispensableMH===false?"Solo registro / no dispensable en MH":null
  }));
  const {error}=await supabaseClient.from("products").upsert(rows,{onConflict:"establishment_code,local_name"});
  if(error)throw error;
  const {data,error:err}=await supabaseClient.from("products").select("id,local_name").eq("establishment_code","MH").eq("is_active",true);
  if(err)throw err;
  cloudProductByNorm=new Map((data||[]).map(p=>[norm(p.local_name),p]));
  cloudProductById=new Map((data||[]).map(p=>[p.id,p]));
}
async function cloudConsumptionCount(){
  if(!cloudConnected)return 0;
  const {count,error}=await supabaseClient.from("monthly_consumption").select("*",{count:"exact",head:true}).eq("establishment_code","MH");
  if(error)throw error;
  return count||0;
}
async function cloudStockCount(){
  if(!cloudConnected)return 0;
  const {count,error}=await supabaseClient.from("stock_snapshots").select("*",{count:"exact",head:true}).eq("establishment_code","MH");
  if(error)throw error;
  return count||0;
}
async function cloudInventoryCount(){
  if(!cloudConnected)return 0;
  const {count,error}=await supabaseClient.from("inventory_runs").select("*",{count:"exact",head:true}).eq("establishment_code","MH");
  if(error)throw error;
  return count||0;
}
async function cloudUpsertConsumption(key,filename="Carga local"){
  if(!cloudConnected)return;
  const month=key+"-01";
  const rows=[];
  master.forEach(m=>{
    const p=cloudProductByNorm.get(norm(m.glosaABG)); const v=getTimeline(m)[key];
    if(p&&num(v)!==null)rows.push({establishment_code:"MH",product_id:p.id,month,quantity:num(v),source_filename:filename});
  });
  if(!rows.length)return;
  const {error}=await supabaseClient.from("monthly_consumption").upsert(rows,{onConflict:"establishment_code,product_id,month"});
  if(error)throw error;
}
async function cloudSaveCurrentStock(filename="Stock Rayen",stockDate=null){
  if(!cloudConnected)return;
  const date=(stockDate||todayISO()).slice(0,10);
  const snapRow={establishment_code:"MH",stock_date:date,source:"rayen",filename};
  let {data:snap,error}=await supabaseClient.from("stock_snapshots").upsert(snapRow,{onConflict:"establishment_code,stock_date,source"}).select("id").single();
  if(error)throw error;
  await supabaseClient.from("stock_items").delete().eq("snapshot_id",snap.id);
  const rows=[];
  master.forEach(m=>{
    const p=cloudProductByNorm.get(norm(m.glosaABG)),q=stockFor(m);
    if(p&&q!==null)rows.push({snapshot_id:snap.id,product_id:p.id,quantity:q});
  });
  if(!rows.length)throw new Error("No había stock local válido para sincronizar.");
  const {error:e2}=await supabaseClient.from("stock_items").insert(rows);
  if(e2)throw e2;
  return {snapshotId:snap.id,date,count:rows.length};
}
async function syncLocalStockIfNeeded(){
  if(!cloudConnected || !Object.keys(stockMap).length)return {uploaded:false,reason:"no_local_stock"};
  const localStamp=localStorage.getItem(STORAGE.lastStock);
  if(!localStamp)return {uploaded:false,reason:"no_local_date"};
  const localDate=String(localStamp).slice(0,10);
  const {data,error}=await supabaseClient.from("stock_snapshots")
    .select("stock_date,created_at")
    .eq("establishment_code","MH")
    .order("stock_date",{ascending:false})
    .order("created_at",{ascending:false})
    .limit(1);
  if(error)throw error;
  const latest=data?.[0]?.stock_date||null;
  if(!latest || localDate>latest){
    const r=await cloudSaveCurrentStock("Migración automática del último stock local",localDate);
    return {uploaded:true,...r};
  }
  if(localDate===latest){
    const {data:snap,error:se}=await supabaseClient.from("stock_snapshots")
      .select("id").eq("establishment_code","MH").eq("stock_date",localDate).eq("source","rayen").maybeSingle();
    if(se)throw se;
    if(snap?.id){
      const {count,error:ce}=await supabaseClient.from("stock_items").select("*",{count:"exact",head:true}).eq("snapshot_id",snap.id);
      if(ce)throw ce;
      if(!count){
        const r=await cloudSaveCurrentStock("Reparación automática del último stock local",localDate);
        return {uploaded:true,...r};
      }
    }
  }
  return {uploaded:false,reason:"cloud_same_or_newer"};
}

async function cloudSaveInventory(type,date,responsible,items,qf=""){
  if(!cloudConnected)return;
  const run={establishment_code:"MH",inventory_type:type,inventory_date:date,status:"realizado",qf_responsible:qf||null,participants:responsible?[responsible]:[]};
  const {data:r,error}=await supabaseClient.from("inventory_runs").insert(run).select("id").single();
  if(error)throw error;
  const rows=[];
  (items||[]).forEach(i=>{
    const p=cloudProductByNorm.get(norm(i.product));
    if(!p)return;
    const system=i.system??null,physical=i.physical??null;
    rows.push({inventory_run_id:r.id,product_id:p.id,counter_name:i.responsible||responsible||null,system_stock:system,physical_count:physical,qf_recount:i.qf_recount??null,adjusted_stock:i.adjusted_stock??null,observation:i.observation||null});
  });
  if(rows.length){
    const {error:e2}=await supabaseClient.from("inventory_items").insert(rows);
    if(e2)throw e2;
  }
}
async function migrateLocalDataIfCloudEmpty(){
  if(!cloudConnected)return;
  if(await cloudConsumptionCount()===0){
    const keys=[...new Set(Object.values(timeline).flatMap(x=>Object.keys(x||{})))].filter(k=>/^\d{4}-\d{2}$/.test(k)).sort();
    for(const key of keys)await cloudUpsertConsumption(key,"Migración inicial desde MH");
  }
  await syncLocalStockIfNeeded();
  if(await cloudInventoryCount()===0){
    for(const h of [...weeklyHistory].reverse())await cloudSaveInventory("rotativo",h.date,h.responsible,h.items||[]);
    for(const h of [...generalHistory].reverse())await cloudSaveInventory("general",h.date,"",h.items||[],h.qf||"");
  }
  localStorage.setItem(STORAGE.cloudMigrated,new Date().toISOString());
}
async function loadCloudState(){
  if(!cloudConnected)return;
  const {data:cons,error:ce}=await supabaseClient.from("monthly_consumption").select("product_id,month,quantity").eq("establishment_code","MH").order("month");
  if(ce)throw ce;
  if(cons?.length){
    timeline={};
    for(const c of cons){
      const p=cloudProductById.get(c.product_id); if(!p)continue;
      const k=norm(p.local_name),mk=String(c.month).slice(0,7);
      timeline[k]=timeline[k]||{}; timeline[k][mk]=Number(c.quantity);
    }
    save(STORAGE.timeline,timeline);
  }

  const {data:snaps,error:se}=await supabaseClient.from("stock_snapshots").select("id,stock_date,created_at").eq("establishment_code","MH").order("stock_date",{ascending:false}).order("created_at",{ascending:false}).limit(2);
  if(se)throw se;
  if(snaps?.length){
    const current=snaps[0],previous=snaps[1];
    const ids=snaps.map(s=>s.id);
    const {data:si,error:sie}=await supabaseClient.from("stock_items").select("snapshot_id,product_id,quantity").in("snapshot_id",ids);
    if(sie)throw sie;
    stockMap={}; prevStockMap={};
    for(const x of si||[]){
      const p=cloudProductById.get(x.product_id); if(!p)continue;
      if(x.snapshot_id===current.id)stockMap[norm(p.local_name)]=Number(x.quantity);
      if(previous&&x.snapshot_id===previous.id)prevStockMap[norm(p.local_name)]=Number(x.quantity);
    }
    save(STORAGE.stock,stockMap); save(STORAGE.prevStock,prevStockMap);
    localStorage.setItem(STORAGE.lastStock,current.stock_date+"T12:00:00");
  }

  const {data:runs,error:re}=await supabaseClient.from("inventory_runs").select("id,inventory_type,inventory_date,qf_responsible,participants,created_at").eq("establishment_code","MH").order("inventory_date",{ascending:false}).limit(60);
  if(re)throw re;
  if(runs?.length){
    const ids=runs.map(r=>r.id);
    const {data:it,error:ie}=await supabaseClient.from("inventory_items").select("inventory_run_id,product_id,counter_name,system_stock,physical_count,difference,qf_recount,adjusted_stock,observation").in("inventory_run_id",ids);
    if(ie)throw ie;
    const byRun=new Map();
    for(const x of it||[]){
      const p=cloudProductById.get(x.product_id); if(!p)continue;
      const a=byRun.get(x.inventory_run_id)||[];
      a.push({product:p.local_name,system:x.system_stock===null?null:Number(x.system_stock),physical:x.physical_count===null?null:Number(x.physical_count),difference:x.difference===null?null:Number(x.difference),responsible:x.counter_name||"",qf_recount:x.qf_recount===null?null:Number(x.qf_recount),adjusted_stock:x.adjusted_stock===null?null:Number(x.adjusted_stock),observation:x.observation||""});
      byRun.set(x.inventory_run_id,a);
    }
    weeklyHistory=runs.filter(r=>r.inventory_type==="rotativo").map(r=>({date:r.inventory_date,responsible:Array.isArray(r.participants)?(r.participants[0]||""):"",items:byRun.get(r.id)||[]}));
    generalHistory=runs.filter(r=>r.inventory_type==="general").map(r=>({date:r.inventory_date,qf:r.qf_responsible||"",items:byRun.get(r.id)||[]}));
    save(STORAGE.weekly,weeklyHistory); save(STORAGE.generalHistory,generalHistory);
  }
  updateStockouts();
  renderAll();
  // Si la ficha de un producto estaba abierta antes de terminar la sincronización,
  // refrescarla con los consumos recién descargados desde Supabase.
  if(activeMedicineDetail && $("medicineDetailDialog")?.open){
    openMedicineDetail(activeMedicineDetail.glosaABG);
  }
}
async function tryAutoConnectCloud(){
  setCloudBadge("Conectando…",false);
  const ok=await connectCloud(INTERNAL_MH_CLOUD_CODE,false,true);
  if(ok){
    setCloudBadge("Sincronizado",true);
  }else{
    setCloudBadge("Sin conexión",false);
  }
}

function load(k,f){try{return JSON.parse(localStorage.getItem(k))??f}catch{return f}}
function save(k,v){localStorage.setItem(k,JSON.stringify(v))}
function norm(v){return String(v??"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").toUpperCase().replace(/,/g,".").replace(/[^\w.%/+\- ]/g," ").replace(/\s+/g," ").trim()}
function num(v){if(v===null||v===undefined||v==="")return null;if(typeof v==="number")return Number.isFinite(v)?v:null;let s=String(v).trim().replace(/\s/g,"");if(s.includes(",")&&!s.includes("."))s=s.replace(",",".");else if(s.includes(",")&&s.includes(".")){if(s.lastIndexOf(",")>s.lastIndexOf("."))s=s.replace(/\./g,"").replace(",",".");else s=s.replace(/,/g,"")}const x=Number(s);return Number.isFinite(x)?x:null}
function esc(v){return String(v??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}[c]))}
function fmt(v,d=1){return v===null||v===undefined||Number.isNaN(v)?"—":Number(v).toLocaleString("es-CL",{maximumFractionDigits:d})}
function todayISO(){return new Date().toISOString().slice(0,10)}
function sameDay(iso){return iso&&new Date(iso).toISOString().slice(0,10)===todayISO()}
function previousMonthKey(){const d=new Date();d.setDate(1);d.setMonth(d.getMonth()-1);return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`}
function monthLabel(key){const [y,m]=key.split("-").map(Number);return `${MONTHS[m-1][0].toUpperCase()+MONTHS[m-1].slice(1)} ${y}`}
function monthKeyFromName(name,year=2026){return `${year}-${String(MONTHS.indexOf(name)+1).padStart(2,"0")}`}
function badge(v){let c="neutral";if(v==="CRÍTICO")c="critical";else if(["REVISAR","VIGILAR"].includes(v))c="warning";else if(v==="OK")c="good";else if(["REVISAR BODEGA","PEDIR"].includes(v))c="storage";else if(v==="INFORMAR")c="notify";else if(v==="ALTO")c="risk";else if(v==="SOLO REGISTRO"||v==="DERIVAR ABG"||v==="NO APLICA")c="neutral";return `<span class="badge ${c}">${esc(v)}</span>`}
function program(m){if(m.iaaps&&m.fofar)return"IAAPS/FOFAR";if(m.iaaps)return"IAAPS";if(m.fofar)return"FOFAR";return"NO TRAZADOR"}
function applyArsenal(){const removed=new Set(arsenalRemoved);let x=[];baseMaster.forEach(m=>{const k=norm(m.glosaABG);if(!removed.has(k))x.push({...m,...(arsenalOverride[k]||{}),source:"maestro"})});arsenalAdd.forEach(m=>{const k=norm(m.glosaABG);if(!removed.has(k)){const i=x.findIndex(a=>norm(a.glosaABG)===k);const r={...m,...(arsenalOverride[k]||{}),source:"local"};if(i>=0)x[i]=r;else x.push(r)}});master=x.sort((a,b)=>a.glosaABG.localeCompare(b.glosaABG,"es"))}
function ensureTimelineSeed(){let changed=false;master.forEach(m=>{const k=norm(m.glosaABG);if(!timeline[k])timeline[k]={};const init=m.initialConsumption||{};Object.entries(init).forEach(([mn,v])=>{const mk=monthKeyFromName(mn,2026);if(timeline[k][mk]===undefined&&num(v)!==null){timeline[k][mk]=num(v);changed=true}})});if(changed)save(STORAGE.timeline,timeline)}
function getTimeline(m){return timeline[norm(m.glosaABG)]||{}}
function lastMonthValues(m,count=3){const entries=Object.entries(getTimeline(m)).filter(([,v])=>num(v)!==null).sort((a,b)=>a[0].localeCompare(b[0]));return entries.slice(-count).map(([key,v])=>({key,value:num(v)}))}
function cpm3(m){const vals=lastMonthValues(m,3);return vals.length?vals.reduce((s,x)=>s+x.value,0)/vals.length:null}
function trend(m){const vals=lastMonthValues(m,3);if(vals.length<2)return"Sin datos";const first=vals[0].value,last=vals[vals.length-1].value;if(first===0&&last===0)return"Estable";const pct=(last-first)/Math.max(first,1);if(pct>0.15)return"↑ Al alza";if(pct<-0.15)return"↓ A la baja";return"→ Estable"}
function stockFor(m){const k=norm(m.glosaABG);return Object.prototype.hasOwnProperty.call(stockMap,k)?stockMap[k]:null}
function prevStockFor(m){const k=norm(m.glosaABG);return Object.prototype.hasOwnProperty.call(prevStockMap,k)?prevStockMap[k]:null}
function stockDelta(m){const a=stockFor(m),b=prevStockFor(m);return a===null||b===null?null:a-b}
function coverageMonths(m){const c=cpm3(m),s=stockFor(m);return c&&s!==null?s/c:null}
function state(m){if(m.dispensableMH===false)return"SOLO REGISTRO";const c=cpm3(m),s=stockFor(m);if(s===null)return"SIN STOCK";if(c===null||c<=0)return"SIN CPM";if(s<c*.2)return"CRÍTICO";if(s<c*.4)return"REVISAR";return"OK"}
function action(m){if(m.dispensableMH===false)return"DERIVAR ABG";const st=state(m);if(st!=="CRÍTICO")return st==="REVISAR"?"VIGILAR":"—";if(m.disponibleBodega===false)return"INFORMAR";const d=warehouse[norm(m.glosaABG)];if(d==="available")return"PEDIR";if(d==="unavailable")return"INFORMAR";return"REVISAR BODEGA"}
function risk(m){if(m.dispensableMH===false)return"NO APLICA";const cov=coverageMonths(m),tr=trend(m),st=state(m);if(st==="CRÍTICO"||cov!==null&&cov<.4)return"ALTO";if(cov!==null&&(cov<1||tr.includes("al alza")&&cov<1.5))return"MEDIO";if(cov===null)return"SIN DATOS";return"BAJO"}
function isOverstock(m){const cov=coverageMonths(m);return cov!==null&&cov>3}
function noRecentConsumption(m){const vals=lastMonthValues(m,2);return vals.length<2||vals.every(x=>x.value===0)}
function expiryStatus(e){const days=Math.ceil((new Date(e.date)-new Date(todayISO()))/86400000);if(days<=30)return"≤30 días";if(days<=90)return"31–90 días";if(days<=180)return"91–180 días";return">180 días"}
function expiryFor(m){return expiries.filter(e=>norm(e.product)===norm(m.glosaABG)).sort((a,b)=>a.date.localeCompare(b.date))[0]||null}
function exceptionsFor(m){
  const ex=[];
  const hasStockLoaded=!!localStorage.getItem(STORAGE.lastStock);
  if(cpm3(m)===null)ex.push("Sin CPM");
  if(!m.glosaBodega)ex.push("Sin glosa bodega");
  if(hasStockLoaded && stockFor(m)===null)ex.push("Sin stock leído");
  return ex
}
async function readWorkbook(file){await ensureSheetJS();return new Promise((res,rej)=>{const r=new FileReader();r.onload=e=>{try{res(XLSX.read(new Uint8Array(e.target.result),{type:"array"}))}catch(x){rej(x)}};r.onerror=rej;r.readAsArrayBuffer(file)})}
function sheetRows(wb,name){const n=name&&wb.SheetNames.find(x=>norm(x)===norm(name))||wb.SheetNames[0];return XLSX.utils.sheet_to_json(wb.Sheets[n],{defval:null,raw:true})}
function findHeader(hs,cands){const ns=hs.map(norm);for(const c of cands){const t=norm(c);let i=ns.findIndex(x=>x===t);if(i>=0)return hs[i];i=ns.findIndex(x=>x.includes(t));if(i>=0)return hs[i]}return null}

async function init(){
  $("masterBadge").textContent="Cargando maestro…";
  try{
    const mr=await fetch("./maestro.json",{cache:"no-store"});
    if(!mr.ok)throw new Error(`No se pudo abrir maestro.json (${mr.status})`);
    const m=await mr.json();
    baseMaster=m.medicamentos||[];
    masterVersion=m.version||"";
    if(!baseMaster.length)throw new Error("El maestro está vacío.");

    try{
      const pr=await fetch("./pedido_template.json",{cache:"no-store"});
      pedidoTemplate=pr.ok?await pr.json():{sheets:{}};
    }catch{
      pedidoTemplate={sheets:{}};
    }

    applyArsenal();
    ensureTimelineSeed();
    populateMonths();
    populateDatalist();
    renderAll();
    tryAutoConnectCloud();
  }catch(e){
    $("masterBadge").textContent="Error al cargar maestro";
    $("summaryText").textContent=e.message;
    console.error(e);
    throw e;
  }
}
function populateMonths(){const opts=[];const currentYear=new Date().getFullYear();for(let y=currentYear-1;y<=currentYear+4;y++)MONTHS.forEach((m,i)=>opts.push(`<option value="${y}-${String(i+1).padStart(2,"0")}">${m[0].toUpperCase()+m.slice(1)} ${y}</option>`));$("monthSelect").innerHTML=opts.join("");$("singleMonth").innerHTML=opts.join("");$("monthSelect").value=previousMonthKey();$("singleMonth").value=previousMonthKey()}
function populateDatalist(){$("medicineOptions").innerHTML=master.map(m=>`<option value="${esc(m.glosaABG)}"></option>`).join("")}

async function loadStock(file){
  const wb=await readWorkbook(file),rows=sheetRows(wb);if(!rows.length)throw Error("Archivo de stock vacío");
  const hs=Object.keys(rows[0]),art=findHeader(hs,["ARTÍCULO","ARTICULO","PRODUCTO","MEDICAMENTO"]),stk=findHeader(hs,["TOTAL EN BODEGA","STOCK ACTUAL","STOCK INSTITUCIONAL","STOCK"]);
  if(!art||!stk)throw Error("No encontré las columnas de producto y stock.");
  prevStockMap={...stockMap};save(STORAGE.prevStock,prevStockMap);stockMap={};
  rows.forEach(r=>{if(r[art]){const v=num(r[stk]);if(v!==null)stockMap[norm(r[art])]=v}});
  save(STORAGE.stock,stockMap);localStorage.setItem(STORAGE.lastStock,new Date().toISOString());updateStockouts();renderAll();if(cloudConnected)await cloudSaveCurrentStock(file.name||"Stock Rayen",todayISO());
}
function updateStockouts(){
  const now=new Date().toISOString();
  master.filter(m=>m.dispensableMH!==false).forEach(m=>{const k=norm(m.glosaABG),s=stockFor(m),rec=stockouts[k];if(s===0&&!rec?.active)stockouts[k]={active:true,start:now,end:null};if(s!==null&&s>0&&rec?.active){rec.active=false;rec.end=now}});
  save(STORAGE.stockouts,stockouts)
}
async function loadMonthlyConsumption(file,key){
  // Antes de leer y sobrescribir, revisar si este mes ya tiene datos guardados.
  const existing=master.filter(m=>Object.prototype.hasOwnProperty.call(getTimeline(m),key));
  if(existing.length){
    const total=master.length;
    const msg=existing.length===total
      ? `El consumo de ${monthLabel(key)} ya fue cargado para todo el arsenal (${existing.length} medicamentos).\n\n¿Deseas reemplazar los valores existentes con este archivo?`
      : `El consumo de ${monthLabel(key)} ya tiene datos guardados para ${existing.length} de ${total} productos.\n\nSi continúas, los productos incluidos en el archivo reemplazarán los valores existentes de ese mes. ¿Deseas continuar?`;
    if(!confirm(msg))return;
  }

  const wb=await readWorkbook(file),rows=sheetRows(wb);if(!rows.length)throw Error("Archivo de consumo vacío");
  const hs=Object.keys(rows[0]),art=findHeader(hs,["GLOSA ABG","ARTÍCULO","ARTICULO","MEDICAMENTO","PRODUCTO","GLOSA"]);if(!art)throw Error("No encontré la columna de medicamento.");
  let ch=findHeader(hs,[monthLabel(key).split(" ")[0],"CONSUMO","CONSUMO MES","CONSUMO MENSUAL","CANTIDAD","TOTAL"]);
  if(!ch){const cand=hs.filter(h=>h!==art).filter(h=>rows.slice(0,30).map(r=>num(r[h])).filter(v=>v!==null).length>=3);if(cand.length===1)ch=cand[0]}
  if(!ch)throw Error("No pude identificar la columna de consumo.");

  const valid=new Set(master.map(m=>norm(m.glosaABG)));
  let nup=0,replaced=0,newValues=0;
  rows.forEach(r=>{
    if(!r[art])return;
    const k=norm(r[art]),v=num(r[ch]);
    if(valid.has(k)&&v!==null){
      timeline[k]=timeline[k]||{};
      if(Object.prototype.hasOwnProperty.call(timeline[k],key))replaced++;
      else newValues++;
      timeline[k][key]=v;
      nup++;
    }
  });

  save(STORAGE.timeline,timeline);
  renderAll();
  if(cloudConnected)await cloudUpsertConsumption(key,file.name||"Consumo mensual");
  alert(`Consumo ${monthLabel(key)} guardado para ${nup} productos.${replaced?`\nValores reemplazados: ${replaced}.`:""}${newValues?`\nValores nuevos: ${newValues}.`:""}`);
}
async function saveSingleConsumption(){
  const name=$("singleProduct").value,key=$("singleMonth").value,v=num($("singleValue").value);
  const m=master.find(x=>norm(x.glosaABG)===norm(name));if(!m)return alert("Producto no encontrado en arsenal.");if(v===null||v<0)return alert("Ingresa un valor válido.");
  const k=norm(m.glosaABG);timeline[k]=timeline[k]||{};
  if(timeline[k][key]!==undefined&&!confirm(`Ya existe ${timeline[k][key]} para ${monthLabel(key)}. ¿Reemplazar?`))return;
  timeline[k][key]=v;save(STORAGE.timeline,timeline);
  if(cloudConnected)await cloudUpsertConsumption(key,"Ingreso individual");
  $("singleConsumptionDialog").close();renderAll();
}

function renderMonthlyStatus(){
  const key=previousMonthKey(),missing=master.filter(m=>!Object.prototype.hasOwnProperty.call(getTimeline(m),key));
  const box=$("monthlyStatus");if(!missing.length){box.className="status-strip ok";box.innerHTML=`✅ Consumo de ${monthLabel(key)} cargado para los ${master.length} productos.`}
  else{box.className="status-strip warning";box.innerHTML=`⚠️ Consumo de ${monthLabel(key)} incompleto: ${master.length-missing.length} de ${master.length}.<details class="month-missing"><summary>Ver ${missing.length} productos faltantes</summary><ol>${missing.map(m=>`<li>${esc(m.glosaABG)}</li>`).join("")}</ol></details>`}
}
function nextOrderInfo(){
  const today=new Date(todayISO()+"T00:00:00");for(const d of [...orderDates].sort()){const od=new Date(d+"T00:00:00"),diff=Math.round((od-today)/86400000);if(diff>=0)return{date:d,diff,eligible:diff<=3}}return null
}
function nextGeneralInfo(){
  return generalDates.filter(x=>x.status==="Programado"&&x.date>=todayISO()).sort((a,b)=>a.date.localeCompare(b.date))[0]||null
}
function renderAlerts(){
  const a=[];
  const currentYear=String(new Date().getFullYear());
  const hasGeneralThisYear=generalDates.some(x=>String(x.date||"").startsWith(currentYear+"-"));
  if(!hasGeneralThisYear)a.push(`<div class="alert info">📋 No hay inventarios generales configurados para ${currentYear}.</div>`);
  const g=nextGeneralInfo();if(g)a.push(`<div class="alert info">📋 Próximo inventario general: ${new Date(g.date+"T00:00:00").toLocaleDateString("es-CL")}.</div>`);
  const lastStock=localStorage.getItem(STORAGE.lastStock);
  if(!lastStock)a.push(`<div class="alert info">📦 Aún no se ha cargado stock Rayen de MH. Los consumos históricos y CPM ya están disponibles; las alertas de stock se activarán cuando cargues el archivo.</div>`);
  else if(!sameDay(lastStock))a.push(`<div class="alert warning">Último stock Rayen MH: ${new Date(lastStock).toLocaleString("es-CL")}. </div>`);
  a.push(`<div class="alert info">🔒 Los medicamentos controlados se mantienen para registro/indicación, pero no se consideran dispensables ni generan quiebres de stock en MH.</div>`);
  $("alertsArea").innerHTML=a.join("")
}
function renderKPIs(){
  const xs=master.map(m=>({m,st:state(m),ac:action(m),ri:risk(m),ex:exceptionsFor(m)}));
  $("kpiCritical").textContent=xs.filter(x=>x.st==="CRÍTICO").length;$("kpiReview").textContent=xs.filter(x=>x.st==="REVISAR").length;$("kpiWarehouse").textContent=xs.filter(x=>x.ac==="REVISAR BODEGA").length;$("kpiNotify").textContent=xs.filter(x=>x.ac==="INFORMAR").length;$("kpiRisk").textContent=xs.filter(x=>x.ri==="ALTO").length;$("kpiOverstock").textContent=xs.filter(x=>isOverstock(x.m)).length;$("kpiExpiry").textContent=expiries.filter(e=>expiryStatus(e)!==">180 días").length;$("kpiException").textContent=xs.filter(x=>x.ex.length).length
}
function filteredMaster(){const q=norm($("searchInput").value),p=$("programFilter").value,s=$("stateFilter").value;return master.filter(m=>{if(q&&!norm(m.glosaABG+" "+(m.glosaBodega||"")).includes(q))return false;if(p&&program(m)!==p)return false;if(s&&state(m)!==s)return false;if(activeQuick==="critical"&&state(m)!=="CRÍTICO")return false;if(activeQuick==="review"&&state(m)!=="REVISAR")return false;if(activeQuick==="warehouse"&&action(m)!=="REVISAR BODEGA")return false;if(activeQuick==="notify"&&action(m)!=="INFORMAR")return false;if(activeQuick==="risk"&&risk(m)!=="ALTO")return false;if(activeQuick==="overstock"&&!isOverstock(m))return false;if(activeQuick==="expiry"&&!expiryFor(m))return false;if(activeQuick==="exception"&&!exceptionsFor(m).length)return false;return true})}

function medicineMetaFor(m){return medicineMeta[norm(m.glosaABG)]||{}}
function monthEntries(m){return Object.entries(getTimeline(m)).filter(([k,v])=>/^\d{4}-\d{2}$/.test(k)&&num(v)!==null).sort((a,b)=>a[0].localeCompare(b[0])).map(([key,value])=>({key,value:num(value)}))}
function movingAverage3(entries,index){const vals=entries.slice(Math.max(0,index-2),index+1).map(x=>x.value).filter(v=>v!==null);return vals.length?vals.reduce((a,b)=>a+b,0)/vals.length:null}
function trendForEntries(entries){if(entries.length<2)return{label:"Sin tendencia suficiente",pct:null,dir:"stable"};const a=entries.at(-2).value,b=entries.at(-1).value;if(a===0)return{label:b>0?"Al alza":"Estable",pct:null,dir:b>0?"up":"stable"};const pct=(b-a)/a*100;return pct>5?{label:"Al alza",pct,dir:"up"}:pct<-5?{label:"A la baja",pct,dir:"down"}:{label:"Estable",pct,dir:"stable"}}
function lastPhysicalCountFor(m){const k=norm(m.glosaABG),r=[];weeklyHistory.forEach(h=>(h.items||[]).forEach(it=>{if(norm(it.product)===k)r.push({date:h.date,type:"Rotativo",responsible:h.responsible||it.responsible||""})}));generalHistory.forEach(h=>(h.items||[]).forEach(it=>{if(norm(it.product)===k)r.push({date:h.date,type:"General",responsible:it.responsible||""})}));return r.sort((a,b)=>String(b.date).localeCompare(String(a.date)))[0]||null}
function drawConsumptionChart(entries){const canvas=$("consumptionChart"),ctx=canvas.getContext("2d"),w=Math.max(520,canvas.parentElement.clientWidth-20),h=320,dpr=window.devicePixelRatio||1;canvas.width=w*dpr;canvas.height=h*dpr;canvas.style.width=w+"px";canvas.style.height=h+"px";ctx.scale(dpr,dpr);ctx.clearRect(0,0,w,h);ctx.font="12px sans-serif";if(!entries.length){ctx.fillStyle="#5f6d76";ctx.fillText("Sin consumos históricos para graficar.",20,40);return}const p={l:55,r:20,t:20,b:55},vals=entries.map(e=>e.value),ma=entries.map((e,i)=>movingAverage3(entries,i)),max=Math.max(...vals,...ma.filter(v=>v!==null),1)*1.1,x=i=>p.l+(entries.length===1?0:i*(w-p.l-p.r)/(entries.length-1)),y=v=>h-p.b-v*(h-p.t-p.b)/max;ctx.strokeStyle="#d7dee3";ctx.fillStyle="#5f6d76";ctx.lineWidth=1;for(let g=0;g<=4;g++){const yy=p.t+g*(h-p.t-p.b)/4;ctx.beginPath();ctx.moveTo(p.l,yy);ctx.lineTo(w-p.r,yy);ctx.stroke();ctx.fillText(Math.round(max-g*max/4).toLocaleString("es-CL"),4,yy+4)}ctx.strokeStyle="#1f6f8b";ctx.lineWidth=2.5;ctx.beginPath();entries.forEach((e,i)=>i?ctx.lineTo(x(i),y(e.value)):ctx.moveTo(x(i),y(e.value)));ctx.stroke();ctx.fillStyle="#1f6f8b";entries.forEach((e,i)=>{ctx.beginPath();ctx.arc(x(i),y(e.value),3.5,0,Math.PI*2);ctx.fill()});ctx.strokeStyle="#7c8a93";ctx.lineWidth=2;ctx.setLineDash([6,5]);ctx.beginPath();ma.forEach((v,i)=>i?ctx.lineTo(x(i),y(v)):ctx.moveTo(x(i),y(v)));ctx.stroke();ctx.setLineDash([]);ctx.fillStyle="#5f6d76";entries.forEach((e,i)=>{ctx.save();ctx.translate(x(i),h-p.b+18);ctx.rotate(-Math.PI/6);ctx.fillText(monthLabel(e.key),0,0);ctx.restore()})}
function openMedicineDetail(name){const m=master.find(x=>norm(x.glosaABG)===norm(name));if(!m)return;activeMedicineDetail=m;const meta=medicineMetaFor(m),entries=monthEntries(m),tr=trendForEntries(entries),s=stockFor(m),c=cpm3(m),cov=c&&s!==null?s/c:null,last=lastPhysicalCountFor(m),exp=expiryFor(m);$("medicineDetailTitle").textContent=m.glosaABG;
const latestEntry=entries.at(-1);
$("medicineDetailSubtitle").textContent=(m.glosaBodega?`Glosa bodega: ${m.glosaBodega}`:"Sin glosa de bodega")+(latestEntry?` · Último consumo: ${monthLabel(latestEntry.key)}`:"");$("medicineDetailSummary").innerHTML=`<div class="detail-card"><span>Stock actual</span><strong>${s===null?"—":fmt(s,0)}</strong></div><div class="detail-card"><span>CPM 3 meses</span><strong>${c===null?"—":fmt(c)}</strong></div><div class="detail-card"><span>Cobertura</span><strong>${cov===null?"—":fmt(cov,1)+" meses"}</strong></div><div class="detail-card"><span>Tendencia</span><strong>${tr.label}</strong></div><div class="detail-card"><span>Crítico 20%</span><strong>${c===null?"—":fmt(c*.2,0)}</strong></div><div class="detail-card"><span>Mínimo 40%</span><strong>${c===null?"—":fmt(c*.4,0)}</strong></div><div class="detail-card"><span>Controlado</span><strong>${m.controlado?"Sí · solo registro MH":"No"}</strong></div><div class="detail-card"><span>Disponible bodega</span><strong>${String(m.disponibleBodega).toLowerCase()==="no"?"No":"Sí"}</strong></div>`;$("medicineHistoryTable").innerHTML=entries.length?entries.map((e,i)=>`<tr><td>${monthLabel(e.key)}</td><td>${fmt(e.value,0)}</td><td>${fmt(movingAverage3(entries,i),1)}</td></tr>`).join(""):`<tr><td colspan="3">Sin consumos históricos.</td></tr>`;$("detailGES").value=meta.ges||"";$("detailProgram").value=meta.program||m.programa||"";$("detailGeneralArsenal").checked=meta.generalArsenal!==undefined?!!meta.generalArsenal:true;$("detailIAAPS").checked=meta.iaaps!==undefined?!!meta.iaaps:!!m.iaaps;$("detailFOFAR").checked=meta.fofar!==undefined?!!meta.fofar:!!m.fofar;const pct=tr.pct===null?"":` · variación último mes ${tr.pct>0?"+":""}${fmt(tr.pct,1)}%`;$("consumptionTrendSummary").textContent=`${tr.label}${pct}${cov!==null?` · cobertura actual ${fmt(cov,1)} meses`:""}`;$("medicineOperationalHistory").innerHTML=`<div class="history-card"><strong>Último conteo físico:</strong> ${last?`${new Date(last.date+"T00:00:00").toLocaleDateString("es-CL")} · ${last.type}${last.responsible?` · ${esc(last.responsible)}`:""}`:"Sin registro"}</div><div class="history-card"><strong>Vencimiento registrado:</strong> ${exp?`${new Date(exp.date+"T00:00:00").toLocaleDateString("es-CL")} · ${fmt(exp.qty,0)} unidades`:"Sin registro"}</div>`;drawConsumptionChart(entries);$("medicineDetailDialog").showModal()}
function saveMedicineClassification(){if(!activeMedicineDetail)return;const k=norm(activeMedicineDetail.glosaABG);medicineMeta[k]={ges:$("detailGES").value.trim(),program:$("detailProgram").value.trim(),generalArsenal:$("detailGeneralArsenal").checked,iaaps:$("detailIAAPS").checked,fofar:$("detailFOFAR").checked};save(STORAGE.medicineMeta,medicineMeta);alert("Clasificación guardada.")}

function renderTable(){
  const rows=filteredMaster();$("summaryText").textContent=`${rows.length} de ${master.length} productos mostrados.`;
  $("medTable").innerHTML=rows.map(m=>{const c=cpm3(m),s=stockFor(m),d=stockDelta(m),cov=coverageMonths(m),ac=action(m);return `<tr><td><button type="button" class="medicine-link" data-med="${encodeURIComponent(m.glosaABG)}"><strong>${esc(m.glosaABG)}</strong></button>${m.controlado?'<br><span class="badge neutral">Controlado · solo registro</span>':''}</td><td>${badge(program(m))}</td><td>${fmt(s)}</td><td>${d===null?"—":(d>0?"+":"")+fmt(d,0)}</td><td>${fmt(c)}</td><td>${cov===null?"—":fmt(cov,2)+" meses"}</td><td>${esc(trend(m))}</td><td>${badge(risk(m))}</td><td>${fmt(c===null?null:c*.2)}</td><td>${fmt(c===null?null:c*.4)}</td><td>${badge(state(m))}</td><td>${ac==="REVISAR BODEGA"?`<button class="secondary warehouse-btn" data-med="${encodeURIComponent(m.glosaABG)}">Revisar bodega</button>`:badge(ac)}</td><td>${esc(m.glosaBodega||"—")}</td></tr>`}).join("");
  document.querySelectorAll(".warehouse-btn").forEach(b=>b.onclick=()=>openWarehouse(decodeURIComponent(b.dataset.med)));
  document.querySelectorAll(".medicine-link").forEach(b=>b.onclick=()=>openMedicineDetail(decodeURIComponent(b.dataset.med)));
}

function renderAll(){
  $("masterBadge").textContent=`Maestro MH: ${master.length} productos · ${masterVersion}`;const ls=localStorage.getItem(STORAGE.lastStock);$("stockBadge").textContent=ls?`Stock: ${new Date(ls).toLocaleString("es-CL")}`:"Sin stock cargado";renderMonthlyStatus();renderAlerts();renderKPIs();renderTable()
}

function openWarehouse(name){const m=master.find(x=>x.glosaABG===name);$("warehouseTitle").textContent=m.glosaABG;$("warehouseGlosa").innerHTML=`<strong>Buscar en bodega como:</strong><br>${esc(m.glosaBodega||"Sin glosa registrada")}`;$("warehouseDialog").dataset.med=name;$("warehouseDialog").showModal()}
$("warehouseDialog").addEventListener("close",()=>{const r=$("warehouseDialog").returnValue;if(!["available","unavailable"].includes(r))return;const name=$("warehouseDialog").dataset.med,k=norm(name);warehouse[k]=r;save(STORAGE.warehouse,warehouse);if(r==="available"){ordersHistory.unshift({type:"bodega",date:new Date().toISOString(),product:name,status:"Pendiente"});save(STORAGE.orders,ordersHistory)}renderAll()});

function renderArsenal(){const q=norm($("arsenalSearch").value);$("arsenalTable").innerHTML=master.filter(m=>!q||norm(m.glosaABG).includes(q)).map(m=>`<tr><td><strong>${esc(m.glosaABG)}</strong></td><td>${badge(program(m))}</td><td>${m.controlado?"Sí":"No"}</td><td><button type="button" class="secondary edit-med" data-med="${encodeURIComponent(m.glosaABG)}">Editar</button> <button type="button" class="secondary remove-med" data-med="${encodeURIComponent(m.glosaABG)}">Retirar</button></td></tr>`).join("");document.querySelectorAll(".edit-med").forEach(b=>b.onclick=()=>editMed(decodeURIComponent(b.dataset.med)));document.querySelectorAll(".remove-med").forEach(b=>b.onclick=()=>removeMed(decodeURIComponent(b.dataset.med)))}
function clearArsenalForm(){$("arsenalABG").value="";$("arsenalABG").dataset.key="";$("arsenalBodega").value="";$("arsenalProgram").value="";$("arsenalControlled").checked=false}
function editMed(name){const m=master.find(x=>x.glosaABG===name);$("arsenalABG").value=m.glosaABG;$("arsenalABG").dataset.key=norm(m.glosaABG);$("arsenalBodega").value=m.glosaBodega||"";$("arsenalProgram").value=program(m)==="NO TRAZADOR"?"":program(m);$("arsenalControlled").checked=!!m.controlado}
function removeMed(name){if(!confirm(`¿Retirar ${name} del arsenal?`))return;const k=norm(name);if(!arsenalRemoved.includes(k))arsenalRemoved.push(k);arsenalAdd=arsenalAdd.filter(m=>norm(m.glosaABG)!==k);delete arsenalOverride[k];save(STORAGE.arsenalRemoved,arsenalRemoved);save(STORAGE.arsenalAdd,arsenalAdd);save(STORAGE.arsenalOverride,arsenalOverride);applyArsenal();populateDatalist();renderArsenal();renderAll()}
function saveMed(){const name=$("arsenalABG").value.trim();if(!name)return alert("Glosa ABG obligatoria.");const p=$("arsenalProgram").value,flags={iaaps:p==="IAAPS"||p==="IAAPS/FOFAR",fofar:p==="FOFAR"||p==="IAAPS/FOFAR",trazador:!!p};const rec={glosaABG:name,glosaBodega:$("arsenalBodega").value.trim(),categoria:"MEDICAMENTO",...flags,programa:p,controlado:$("arsenalControlled").checked,disponibleBodega:!$("arsenalControlled").checked,initialConsumption:{}};const old=$("arsenalABG").dataset.key,newk=norm(name),base=baseMaster.some(m=>norm(m.glosaABG)===old||norm(m.glosaABG)===newk);if(old&&old!==newk){if(!arsenalRemoved.includes(old))arsenalRemoved.push(old);delete arsenalOverride[old];arsenalAdd=arsenalAdd.filter(m=>norm(m.glosaABG)!==old)}if(base&&(!old||old===newk))arsenalOverride[newk]=rec;else{const i=arsenalAdd.findIndex(m=>norm(m.glosaABG)===newk);if(i>=0)arsenalAdd[i]=rec;else arsenalAdd.push(rec)}arsenalRemoved=arsenalRemoved.filter(k=>k!==newk);save(STORAGE.arsenalRemoved,arsenalRemoved);save(STORAGE.arsenalAdd,arsenalAdd);save(STORAGE.arsenalOverride,arsenalOverride);applyArsenal();timeline[newk]=timeline[newk]||{};save(STORAGE.timeline,timeline);clearArsenalForm();populateDatalist();renderArsenal();renderAll()}

function classifyABC(){const arr=master.map(m=>({m,c:cpm3(m)||0})).filter(x=>x.c>0).sort((a,b)=>b.c-a.c),total=arr.reduce((s,x)=>s+x.c,0);let a=0,map={};arr.forEach(x=>{a+=x.c;const sh=a/total;map[norm(x.m.glosaABG)]=sh<=.8?"A":sh<=.95?"B":"C"});master.forEach(m=>map[norm(m.glosaABG)]=map[norm(m.glosaABG)]||"C");return map}
function weightedPick(pool,n,abc,recent){const c=[...pool],out=[];while(out.length<n&&c.length){const ws=c.map(m=>{let w=(abc[norm(m.glosaABG)]==="A"?5:abc[norm(m.glosaABG)]==="B"?3:1);if(m.iaaps||m.fofar)w*=2.2;if(recent.has(norm(m.glosaABG)))w*=.18;return w});let r=Math.random()*ws.reduce((s,x)=>s+x,0),i=0;for(;i<c.length;i++){r-=ws[i];if(r<=0)break}i=Math.min(i,c.length-1);out.push(c[i]);c.splice(i,1)}return out}
function generateWeekly(){const abc=classifyABC(),recent=new Set();weeklyHistory.slice(0,4).forEach(h=>(h.items||[]).forEach(i=>recent.add(norm(i.product))));const g={A:master.filter(m=>abc[norm(m.glosaABG)]==="A"),B:master.filter(m=>abc[norm(m.glosaABG)]==="B"),C:master.filter(m=>abc[norm(m.glosaABG)]==="C")};let sel=[...weightedPick(g.A,Math.min(9,g.A.length),abc,recent),...weightedPick(g.B,Math.min(4,g.B.length),abc,recent),...weightedPick(g.C,Math.min(2,g.C.length),abc,recent)];if(sel.length<15){const used=new Set(sel.map(m=>norm(m.glosaABG)));sel=sel.concat(weightedPick(master.filter(m=>!used.has(norm(m.glosaABG))),15-sel.length,abc,recent))}currentWeekly=sel.slice(0,15).map((m,i)=>({n:i+1,product:m.glosaABG,abc:abc[norm(m.glosaABG)],tracer:program(m),system:stockFor(m),physical:null}));renderWeekly()}
function renderWeekly(){if(!currentWeekly.length){$("weeklyTable").innerHTML="";$("weeklySummary").textContent="Aún no se genera listado.";return}const counts=currentWeekly.reduce((x,i)=>(x[i.abc]=(x[i.abc]||0)+1,x),{});$("weeklySummary").textContent=`15 productos · A ${counts.A||0} · B ${counts.B||0} · C ${counts.C||0} · IAAPS/FOFAR ${currentWeekly.filter(i=>i.tracer!=="NO TRAZADOR").length}`;$("weeklyTable").innerHTML=currentWeekly.map((i,idx)=>{const d=i.physical===null||i.system===null?null:i.physical-i.system;return `<tr><td>${i.n}</td><td>${esc(i.product)}</td><td>${i.abc}</td><td>${i.tracer==="NO TRAZADOR"?"No":i.tracer}</td><td>${fmt(i.system)}</td><td><input class="inventory-input weekly-input" data-i="${idx}" inputmode="numeric" value="${i.physical??""}"></td><td data-d="${idx}">${fmt(d,0)}</td><td data-r="${idx}" class="${d===0?"result-ok":d!==null?"result-diff":""}">${d===null?"—":d===0?"Concordante":"Diferencia"}</td></tr>`}).join("");document.querySelectorAll(".weekly-input").forEach(inp=>inp.addEventListener("input",()=>{const i=Number(inp.dataset.i),v=inp.value.replace(/\D/g,"");inp.value=v;currentWeekly[i].physical=v===""?null:Number(v);const d=currentWeekly[i].physical===null||currentWeekly[i].system===null?null:currentWeekly[i].physical-currentWeekly[i].system;$(`[data-d="${i}"]`).textContent=fmt(d,0);const rc=$(`[data-r="${i}"]`);rc.textContent=d===null?"—":d===0?"Concordante":"Diferencia";rc.className=d===0?"result-ok":d!==null?"result-diff":""}))}
async function saveWeekly(){
  const resp=$("weeklyResponsible").value.trim();if(!resp)return alert("Indica quién realizó el inventario.");if(!currentWeekly.length)return alert("Genera el listado.");
  const date=$("weeklyDate").value||todayISO();
  weeklyHistory.unshift({date,responsible:resp,items:currentWeekly.map(x=>({...x}))});weeklyHistory=weeklyHistory.slice(0,52);
  save(STORAGE.weekly,weeklyHistory);
  if(cloudConnected)await cloudSaveInventory("rotativo",date,resp,currentWeekly);
  renderWeeklyHistory();alert("Inventario guardado"+(cloudConnected?" y sincronizado en línea.":"."));
}
function renderWeeklyHistory(){$("weeklyHistory").innerHTML=weeklyHistory.slice(0,8).map(h=>{const dif=(h.items||[]).filter(i=>i.physical!==null&&i.system!==null&&i.physical!==i.system).length;const done=(h.items||[]).filter(i=>i.physical!==null).length;const acc=done?((done-dif)/done*100):null;return `<div class="history-card"><strong>${new Date(h.date+"T00:00:00").toLocaleDateString("es-CL")}</strong> · ${esc(h.responsible)} · Exactitud ${acc===null?"—":fmt(acc,1)+"%"} · ${dif} diferencias</div>`}).join("")||'<div class="history-card">Sin inventarios guardados.</div>'}
function exportWeekly(){if(!currentWeekly.length)return alert("Genera inventario.");const rows=currentWeekly.map(i=>({"N°":i.n,"Medicamento":i.product,"Clase ABC":i.abc,"Programa":i.tracer,"Stock sistema":i.system,"Stock físico":i.physical,"Diferencia":i.physical===null||i.system===null?null:i.physical-i.system,"Responsable":$("weeklyResponsible").value}));exportWorkbook({"Inventario rotativo":rows},`Inventario_rotativo_MH_${todayISO()}.xlsx`)}

function renderGeneralCalendar(){$("generalCalendar").innerHTML=generalDates.sort((a,b)=>a.date.localeCompare(b.date)).map((x,i)=>`<div class="calendar-item"><div><strong>${new Date(x.date+"T00:00:00").toLocaleDateString("es-CL")}</strong> · ${esc(x.status)}</div><div class="button-row"><select class="general-status" data-i="${i}"><option ${x.status==="Programado"?"selected":""}>Programado</option><option ${x.status==="Postergado"?"selected":""}>Postergado</option><option ${x.status==="Realizado"?"selected":""}>Realizado</option></select><input class="general-date" data-i="${i}" type="date" value="${x.date}"><button type="button" class="secondary general-remove" data-i="${i}">Eliminar</button></div></div>`).join("");document.querySelectorAll(".general-status").forEach(e=>e.onchange=()=>{generalDates[e.dataset.i].status=e.value;save(STORAGE.generalDates,generalDates);renderAlerts()});document.querySelectorAll(".general-date").forEach(e=>e.onchange=()=>{generalDates[e.dataset.i].date=e.value;save(STORAGE.generalDates,generalDates);renderAlerts()});document.querySelectorAll(".general-remove").forEach(e=>e.onclick=()=>{generalDates.splice(Number(e.dataset.i),1);save(STORAGE.generalDates,generalDates);renderGeneralCalendar();renderAlerts()})}
function generateGeneral(){const names=[$("tens1").value.trim()||"TENS 1",$("tens2").value.trim()||"TENS 2",$("tens3").value.trim()||"TENS 3",$("qfName").value.trim()||"QF"];participants={tens1:names[0],tens2:names[1],tens3:names[2],qf:names[3]};save(STORAGE.participants,participants);const rows=master.map((m,i)=>({"N°":i+1,"Medicamento":m.glosaABG,"Glosa Bodega":m.glosaBodega||"","Responsable conteo":names[i%4],"Stock sistema":stockFor(m),"Conteo físico":"","Diferencia":"","Recuento QF":"","Stock ajustado":"","Observación":"","Vencimiento más próximo":expiryFor(m)?.date||"","Cantidad próxima a vencer":expiryFor(m)?.qty||""}));exportWorkbook({"Inventario general":rows},`Inventario_general_MH_${todayISO()}.xlsx`)}
async function importGeneral(file){
  const wb=await readWorkbook(file),rows=sheetRows(wb,"Inventario general");if(!rows.length)return alert("Archivo sin datos.");
  const hs=Object.keys(rows[0]),prod=findHeader(hs,["MEDICAMENTO","PRODUCTO"]),sys=findHeader(hs,["STOCK SISTEMA"]),phy=findHeader(hs,["CONTEO FÍSICO","CONTEO FISICO"]),resp=findHeader(hs,["RESPONSABLE CONTEO"]);
  if(!prod||!phy)return alert("No reconozco la planilla de inventario general.");
  const items=rows.filter(r=>r[prod]).map(r=>({product:r[prod],system:num(r[sys]),physical:num(r[phy]),responsible:r[resp]||""}));
  const qf=$("qfName").value||participants.qf,date=todayISO();
  generalHistory.unshift({date,qf,items});save(STORAGE.generalHistory,generalHistory);
  if(cloudConnected)await cloudSaveInventory("general",date,"",items,qf);
  renderGeneralHistory();alert("Inventario general incorporado al historial"+(cloudConnected?" y sincronizado en línea.":"."));
}
function renderOrderHistory(){$("orderHistory").innerHTML=ordersHistory.filter(x=>x.type==="mensual").slice(0,6).map(x=>`<div class="history-card"><strong>${new Date(x.scheduled+"T00:00:00").toLocaleDateString("es-CL")}</strong> · ${esc(x.status)} · ${x.items||0} líneas</div>`).join("")||'<div class="history-card">Sin pedidos mensuales generados.</div>'}







function tracerReportValidation(){
  const key=previousMonthKey();
  const tracers=master.filter(m=>m.trazador);
  const missingConsumption=tracers.filter(m=>!Object.prototype.hasOwnProperty.call(getTimeline(m),key));
  const missingCpm=tracers.filter(m=>cpm3(m)===null);
  const missingStock=tracers.filter(m=>stockFor(m)===null);
  const ls=localStorage.getItem(STORAGE.lastStock);
  const stockToday=sameDay(ls);

  return {
    key,
    stockToday,
    missingConsumption,
    missingCpm,
    missingStock,
    ok: stockToday && !missingConsumption.length && !missingCpm.length && !missingStock.length
  };
}

function renderTracerValidation(){
  const v=tracerReportValidation();
  const box=$("tracerValidationStatus");
  const details=$("tracerValidationDetails");
  const btn=$("generateTracerReport");
  const issues=[];

  if(!v.stockToday)issues.push("Debes cargar el stock Rayen actualizado hoy.");
  if(v.missingConsumption.length)issues.push(`Falta consumo de ${monthLabel(v.key)} para ${v.missingConsumption.length} trazador(es).`);
  if(v.missingCpm.length)issues.push(`${v.missingCpm.length} trazador(es) no tienen CPM de los últimos 3 meses.`);
  if(v.missingStock.length)issues.push(`${v.missingStock.length} trazador(es) no fueron reconocidos en el stock actual.`);

  if(v.ok){
    box.className="status-strip ok";
    box.textContent="✅ Datos completos. El informe está listo para generar.";
    details.innerHTML="";
    btn.disabled=false;
  }else{
    box.className="status-strip warning";
    box.textContent="⚠️ Hay datos pendientes antes de generar el informe.";
    btn.disabled=true;

    const sections=[];
    if(v.missingConsumption.length)sections.push(
      `<details class="month-missing"><summary>Consumo faltante (${v.missingConsumption.length})</summary><ol>${v.missingConsumption.map(m=>`<li>${esc(m.glosaABG)}</li>`).join("")}</ol></details>`
    );
    if(v.missingCpm.length)sections.push(
      `<details class="month-missing"><summary>Sin CPM 3 meses (${v.missingCpm.length})</summary><ol>${v.missingCpm.map(m=>`<li>${esc(m.glosaABG)}</li>`).join("")}</ol></details>`
    );
    if(v.missingStock.length)sections.push(
      `<details class="month-missing"><summary>Sin stock reconocido (${v.missingStock.length})</summary><ol>${v.missingStock.map(m=>`<li>${esc(m.glosaABG)}</li>`).join("")}</ol></details>`
    );
    details.innerHTML=`<div class="validation-list">${issues.map(x=>`<div class="exception-card">${esc(x)}</div>`).join("")}</div>${sections.join("")}`;
  }
}

function openTracerReport(){
  renderTracerValidation();
  $("tracerReportDialog").showModal();
}



function renderExpiries(){$("expiryTable").innerHTML=expiries.sort((a,b)=>a.date.localeCompare(b.date)).map((e,i)=>`<tr><td>${esc(e.product)}</td><td>${new Date(e.date+"T00:00:00").toLocaleDateString("es-CL")}</td><td>${fmt(e.qty,0)}</td><td class="${expiryStatus(e)==="≤30 días"?"expiry-high":expiryStatus(e)==="31–90 días"?"expiry-med":"expiry-low"}">${expiryStatus(e)}</td><td><button type="button" class="secondary exp-remove" data-i="${i}">Eliminar</button></td></tr>`).join("");document.querySelectorAll(".exp-remove").forEach(b=>b.onclick=()=>{expiries.splice(Number(b.dataset.i),1);save(STORAGE.expiries,expiries);renderExpiries();renderAll()})}
function saveExpiry(){const p=$("expiryProduct").value.trim(),d=$("expiryDate").value,q=num($("expiryQty").value);if(!master.some(m=>norm(m.glosaABG)===norm(p)))return alert("Medicamento no encontrado.");if(!d||q===null)return alert("Completa fecha y cantidad.");expiries.push({product:p,date:d,qty:q,updated:todayISO()});save(STORAGE.expiries,expiries);renderExpiries();renderAll()}

function renderExceptions(){const cards=[];master.forEach(m=>{const ex=exceptionsFor(m);if(ex.length)cards.push(`<div class="exception-card"><strong>${esc(m.glosaABG)}</strong><br>${ex.map(x=>badge(x)).join(" ")}</div>`)});$("exceptionsList").innerHTML=cards.join("")||'<div class="history-card">Sin excepciones.</div>'}

function repeatedDifferences(m){const k=norm(m.glosaABG);let n=0;weeklyHistory.forEach(h=>(h.items||[]).forEach(i=>{if(norm(i.product)===k&&i.physical!==null&&i.system!==null&&i.physical!==i.system)n++}));generalHistory.forEach(h=>(h.items||[]).forEach(i=>{if(norm(i.product)===k&&i.physical!==null&&i.system!==null&&i.physical!==i.system)n++}));return n}
function exportManagement(){const sheets={Críticos:[],Bodega:[],Informar:[],Sobrestock:[],Excepciones:[],Vencimientos:[]};master.forEach(m=>{const row={"Medicamento":m.glosaABG,"Programa":program(m),"Stock":stockFor(m),"CPM 3M":cpm3(m),"Cobertura meses":coverageMonths(m),"Tendencia":trend(m),"Riesgo":risk(m),"Estado":state(m),"Acción":action(m),"Diferencias repetidas":repeatedDifferences(m)};if(state(m)==="CRÍTICO")sheets.Críticos.push(row);if(["REVISAR BODEGA","PEDIR"].includes(action(m)))sheets.Bodega.push(row);if(action(m)==="INFORMAR")sheets.Informar.push(row);if(isOverstock(m))sheets.Sobrestock.push(row);if(exceptionsFor(m).length)sheets.Excepciones.push({...row,"Excepciones":exceptionsFor(m).join(", ")});const e=expiryFor(m);if(e)sheets.Vencimientos.push({...row,"Vencimiento":e.date,"Cantidad":e.qty,"Alerta":expiryStatus(e)})});exportWorkbook(sheets,`Gestion_inventario_ABG_${todayISO()}.xlsx`)}
async function exportWorkbook(sheets,filename){await ensureSheetJS();const wb=XLSX.utils.book_new();Object.entries(sheets).forEach(([name,rows])=>XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),name.slice(0,31)));XLSX.writeFile(wb,filename)}

function backup(){downloadJSON({version:VERSION,date:new Date().toISOString(),timeline,warehouse,arsenalAdd,arsenalOverride,arsenalRemoved,weeklyHistory,generalDates,generalHistory,participants,ordersHistory,expiries,stockouts,medicineMeta},`Respaldo_MH_${todayISO()}.json`)}
function downloadJSON(obj,name){const b=new Blob([JSON.stringify(obj,null,2)],{type:"application/json"}),a=document.createElement("a");a.href=URL.createObjectURL(b);a.download=name;a.click();URL.revokeObjectURL(a.href)}
async function importBackup(file){const x=JSON.parse(await file.text());if(!confirm("Esto reemplazará los datos guardados localmente. ¿Continuar?"))return;timeline=x.timeline||{};warehouse=x.warehouse||{};arsenalAdd=x.arsenalAdd||[];arsenalOverride=x.arsenalOverride||{};arsenalRemoved=x.arsenalRemoved||[];weeklyHistory=x.weeklyHistory||[];generalDates=x.generalDates||generalDates;generalHistory=x.generalHistory||[];participants=x.participants||participants;ordersHistory=x.ordersHistory||[];expiries=x.expiries||[];stockouts=x.stockouts||{};medicineMeta=x.medicineMeta||{};Object.entries({[STORAGE.timeline]:timeline,[STORAGE.warehouse]:warehouse,[STORAGE.arsenalAdd]:arsenalAdd,[STORAGE.arsenalOverride]:arsenalOverride,[STORAGE.arsenalRemoved]:arsenalRemoved,[STORAGE.weekly]:weeklyHistory,[STORAGE.generalDates]:generalDates,[STORAGE.generalHistory]:generalHistory,[STORAGE.participants]:participants,[STORAGE.orders]:ordersHistory,[STORAGE.expiries]:expiries,[STORAGE.stockouts]:stockouts,[STORAGE.medicineMeta]:medicineMeta}).forEach(([k,v])=>save(k,v));applyArsenal();populateDatalist();renderAll();alert("Respaldo importado.")}

$("stockFile").onchange=async e=>{try{if(e.target.files[0])await loadStock(e.target.files[0])}catch(x){alert(x.message)}e.target.value=""};
$("consumptionFile").onchange=async e=>{try{if(e.target.files[0])await loadMonthlyConsumption(e.target.files[0],$("monthSelect").value)}catch(x){alert(x.message)}e.target.value=""};
$("singleConsumptionBtn").onclick=()=>{$("singleProduct").value="";$("singleValue").value="";$("singleMonth").value=previousMonthKey();$("singleConsumptionDialog").showModal()};
$("saveSingleConsumption").onclick=saveSingleConsumption;$("arsenalBtn").onclick=()=>{clearArsenalForm();renderArsenal();$("arsenalDialog").showModal()};$("arsenalSearch").oninput=renderArsenal;$("saveArsenal").onclick=saveMed;$("clearArsenal").onclick=clearArsenalForm;
$("backupBtn").onclick=backup;$("backupFile").onchange=async e=>{if(e.target.files[0])await importBackup(e.target.files[0]);e.target.value=""};
$("weeklyInventoryBtn").onclick=()=>{$("weeklyDate").value=todayISO();renderWeekly();renderWeeklyHistory();$("weeklyInventoryDialog").showModal()};$("generateWeekly").onclick=generateWeekly;$("saveWeekly").onclick=saveWeekly;$("exportWeekly").onclick=exportWeekly;

function annualYearOptions(){
  const y=new Date().getFullYear(),years=[];
  for(let i=y-1;i<=y+4;i++)years.push(i);
  $("annualYear").innerHTML=years.map(v=>`<option value="${v}">${v}</option>`).join("");
  $("annualYear").value=String(y);
}

function loadAnnualDraft(){
  const y=$("annualYear").value;
  annualDraftOrders=orderDates.filter(d=>d.startsWith(y+"-")).sort().map(d=>({date:d}));
  annualDraftGeneral=generalDates.filter(x=>String(x.date||"").startsWith(y+"-"))
    .sort((a,b)=>a.date.localeCompare(b.date)).map(x=>({...x}));
  renderAnnualDraft();
}

function renderAnnualDraft(){
  const y=$("annualYear").value;
  $("annualYearStatus").textContent=`${annualDraftOrders.length} pedido(s) · ${annualDraftGeneral.length} inventario(s) general(es)`;

  $("annualOrderDates").innerHTML=annualDraftOrders.length
    ? annualDraftOrders.map((x,i)=>`<div class="annual-row">
        <input class="annual-order-date" data-i="${i}" type="date" value="${x.date}">
        <span></span>
        <button type="button" class="secondary annual-order-remove" data-i="${i}">Eliminar</button>
      </div>`).join("")
    : `<div class="annual-empty">Sin fechas de pedidos para ${y}.</div>`;

  $("annualGeneralDates").innerHTML=annualDraftGeneral.length
    ? annualDraftGeneral.map((x,i)=>`<div class="annual-row">
        <input class="annual-general-date" data-i="${i}" type="date" value="${x.date}">
        <select class="annual-general-status" data-i="${i}">
          <option ${x.status==="Programado"?"selected":""}>Programado</option>
          <option ${x.status==="Postergado"?"selected":""}>Postergado</option>
          <option ${x.status==="Realizado"?"selected":""}>Realizado</option>
        </select>
        <button type="button" class="secondary annual-general-remove" data-i="${i}">Eliminar</button>
      </div>`).join("")
    : `<div class="annual-empty">Sin inventarios generales para ${y}.</div>`;

  document.querySelectorAll(".annual-order-date").forEach(el=>el.onchange=()=>{annualDraftOrders[Number(el.dataset.i)].date=el.value});
  document.querySelectorAll(".annual-order-remove").forEach(el=>el.onclick=()=>{annualDraftOrders.splice(Number(el.dataset.i),1);renderAnnualDraft()});
  document.querySelectorAll(".annual-general-date").forEach(el=>el.onchange=()=>{annualDraftGeneral[Number(el.dataset.i)].date=el.value});
  document.querySelectorAll(".annual-general-status").forEach(el=>el.onchange=()=>{annualDraftGeneral[Number(el.dataset.i)].status=el.value});
  document.querySelectorAll(".annual-general-remove").forEach(el=>el.onclick=()=>{annualDraftGeneral.splice(Number(el.dataset.i),1);renderAnnualDraft()});
}

function openAnnualConfig(){
  annualYearOptions();
  loadAnnualDraft();
  $("annualConfigDialog").showModal();
}

function saveAnnualConfiguration(){
  const y=$("annualYear").value;
  const validOrders=annualDraftOrders.map(x=>x.date).filter(d=>d&&d.startsWith(y+"-"));
  const validGeneral=annualDraftGeneral.filter(x=>x.date&&x.date.startsWith(y+"-"))
    .map(x=>({date:x.date,status:x.status||"Programado"}));

  orderDates=orderDates.filter(d=>!d.startsWith(y+"-")).concat(validOrders).sort();
  generalDates=generalDates.filter(x=>!String(x.date||"").startsWith(y+"-")).concat(validGeneral)
    .sort((a,b)=>a.date.localeCompare(b.date));

  save(STORAGE.orderDates,orderDates);
  save(STORAGE.generalDates,generalDates);
  renderAlerts();
  renderGeneralCalendar();
  loadAnnualDraft();
  alert(`Configuración ${y} guardada.`);
}

$("annualConfigBtn").onclick=openAnnualConfig;
$("annualYear").onchange=loadAnnualDraft;
$("addOrderDate").onclick=()=>{const y=$("annualYear").value;annualDraftOrders.push({date:`${y}-01-01`});renderAnnualDraft()};
$("addAnnualGeneralDate").onclick=()=>{const y=$("annualYear").value;annualDraftGeneral.push({date:`${y}-01-01`,status:"Programado"});renderAnnualDraft()};
$("saveAnnualConfig").onclick=saveAnnualConfiguration;

$("generalInventoryBtn").onclick=()=>{$("tens1").value=participants.tens1;$("tens2").value=participants.tens2;$("tens3").value=participants.tens3;$("qfName").value=participants.qf;renderGeneralCalendar();renderGeneralHistory();$("generalInventoryDialog").showModal()};$("addGeneralDate").onclick=()=>{generalDates.push({date:todayISO(),status:"Programado"});save(STORAGE.generalDates,generalDates);renderGeneralCalendar()};$("generateGeneral").onclick=generateGeneral;$("generalCompletedFile").onchange=async e=>{if(e.target.files[0])await importGeneral(e.target.files[0]);e.target.value=""};
$("ordersBtn").onclick=()=>{const e=orderEligibility();$("orderStatus").textContent=e.msg;renderOrderTable();renderOrderHistory();$("ordersDialog").showModal()};$("prepareOrder").onclick=prepareOrder;$("exportOrder").onclick=exportOrder;
$("expiriesBtn").onclick=()=>{renderExpiries();$("expiriesDialog").showModal()};$("saveExpiry").onclick=saveExpiry;$("exceptionsBtn").onclick=()=>{renderExceptions();$("exceptionsDialog").showModal()};$("exportManagementBtn").onclick=exportManagement;
$("searchInput").oninput=()=>{activeQuick="";renderTable()};$("programFilter").onchange=()=>{activeQuick="";renderTable()};$("stateFilter").onchange=()=>{activeQuick="";renderTable()};$("clearFilters").onclick=()=>{$("searchInput").value="";$("programFilter").value="";$("stateFilter").value="";activeQuick="";renderTable()};
$("saveMedicineClassification").onclick=saveMedicineClassification;
document.querySelectorAll("[data-kpi]").forEach(b=>b.onclick=()=>{activeQuick=activeQuick===b.dataset.kpi?"":b.dataset.kpi;renderTable()});

$("cloudBadge").onclick=async()=>{
  if(cloudConnected){
    setCloudBadge("Sincronizando…",true);
    try{
      await syncMasterToCloud();
      await loadCloudState();
      setCloudBadge("Sincronizado",true);
    }catch(e){
      console.error("Sincronización manual:",e);
      setCloudBadge("Sin conexión",false);
    }
  }else{
    await tryAutoConnectCloud();
  }
};
init().catch(e=>{$("masterBadge").textContent="Error al cargar maestro";$("summaryText").textContent=e.message;console.error(e)});
