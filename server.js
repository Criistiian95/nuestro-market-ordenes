const http=require("http");
const fs=require("fs");
const path=require("path");
const XLSX=require("xlsx");
const PDFDocument=require("pdfkit");
const {Pool}=require("pg");

const PORT=process.env.PORT||3000;
const ADMIN_PIN=String(process.env.ADMIN_PIN||"").trim();
const DATABASE_URL=String(process.env.DATABASE_URL||"").trim();
const LOCAL_STATE=path.join(__dirname,"state.local.json");
const pool=DATABASE_URL?new Pool({connectionString:DATABASE_URL,ssl:process.env.NODE_ENV==="production"?{rejectUnauthorized:false}:false}):null;

const emptyState=()=>({costs:null,stock:null,sales:null,dataset:[],meta:{coverage:7,safety:2}});
async function initStore(){
  if(!pool)return;
  await pool.query(`CREATE TABLE IF NOT EXISTS app_state (
    id TEXT PRIMARY KEY,
    payload JSONB NOT NULL,
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
  )`);
}
async function load(){
  if(pool){
    const r=await pool.query("SELECT payload FROM app_state WHERE id=$1",["main"]);
    return r.rows[0]?.payload||emptyState();
  }
  try{return JSON.parse(fs.readFileSync(LOCAL_STATE,"utf8"))}catch(e){return emptyState()}
}
async function save(s){
  if(pool){
    await pool.query(`INSERT INTO app_state(id,payload,updated_at)
      VALUES($1,$2::jsonb,NOW())
      ON CONFLICT(id) DO UPDATE SET payload=EXCLUDED.payload,updated_at=NOW()`,["main",JSON.stringify(s)]);
    return;
  }
  fs.writeFileSync(LOCAL_STATE,JSON.stringify(s));
}

function json(res,status,obj){res.writeHead(status,{"Content-Type":"application/json; charset=utf-8","Cache-Control":"no-store"});res.end(JSON.stringify(obj));}
function file(res,name,type){const p=path.join(__dirname,name);if(!fs.existsSync(p)){res.writeHead(404);return res.end("No encontrado")}res.writeHead(200,{"Content-Type":type,"Cache-Control":"no-cache"});fs.createReadStream(p).pipe(res)}
function body(req,limit=28*1024*1024){return new Promise((resolve,reject)=>{let a=[],n=0;req.on("data",c=>{n+=c.length;if(n>limit){reject(new Error("Archivo demasiado grande"));req.destroy();return}a.push(c)});req.on("end",()=>resolve(Buffer.concat(a)));req.on("error",reject)})}
function norm(v){return String(v??"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").trim().toUpperCase().replace(/\s+/g," ")}
function num(v,d=0){if(v===null||v===undefined||v==="")return d;let s=String(v).trim();if(s.includes(",")&&s.includes("."))s=s.replace(/\./g,"").replace(",",".");else s=s.replace(",",".");const n=Number(s);return Number.isFinite(n)?n:d}
function code(v){if(typeof v==="number"&&Number.isInteger(v))return String(v);return String(v??"").trim().replace(/\.0$/,"")}
function dateKey(v){
  if(v instanceof Date&&!isNaN(v))return [v.getFullYear(),String(v.getMonth()+1).padStart(2,"0"),String(v.getDate()).padStart(2,"0")].join("-");
  if(typeof v==="number"&&v>30000){const d=XLSX.SSF.parse_date_code(v);if(d)return [d.y,String(d.m).padStart(2,"0"),String(d.d).padStart(2,"0")].join("-")}
  const s=String(v??"").trim();const m=s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})$/);if(m){let y=+m[3];if(y<100)y+=2000;return [y,String(+m[2]).padStart(2,"0"),String(+m[1]).padStart(2,"0")].join("-")}
  const d=new Date(s);return isNaN(d)?"":[d.getFullYear(),String(d.getMonth()+1).padStart(2,"0"),String(d.getDate()).padStart(2,"0")].join("-")
}
function parseRows(buf){
  const zip=buf[0]===0x50&&buf[1]===0x4b,ole=buf[0]===0xd0&&buf[1]===0xcf;
  if(zip||ole){const wb=XLSX.read(buf,{type:"buffer",cellDates:true});return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,defval:null,raw:true})}
  let txt=buf.toString("latin1");const lines=txt.split(/\r?\n/).filter(x=>x.trim());const sep=(lines[0]||"").includes("\t")?"\t":";";return lines.map(l=>l.split(sep))
}
function headerRow(rows,groups){for(let i=0;i<Math.min(rows.length,50);i++){const h=(rows[i]||[]).map(norm);if(groups.every(g=>g.some(x=>h.includes(norm(x)))))return i}return -1}
function idx(h,variants){const n=h.map(norm);for(const v of variants){const i=n.indexOf(norm(v));if(i>=0)return i}return -1}
function dayCount(a,b){if(!a||!b)return 0;return Math.round((new Date(b+"T00:00:00")-new Date(a+"T00:00:00"))/86400000)+1}

function rebuild(s){
  if(!s.costs||!s.stock||!s.sales){s.dataset=[];s.meta={...(s.meta||{}),coverage:7,safety:2};return}
  const dates=[];let d=new Date(s.sales.minDate+"T00:00:00"),end=new Date(s.sales.maxDate+"T00:00:00");
  while(d<=end){dates.push([d.getFullYear(),String(d.getMonth()+1).padStart(2,"0"),String(d.getDate()).padStart(2,"0")].join("-"));d.setDate(d.getDate()+1)}
  const out=[];const cov=7,safe=2;
  for(const [c,m] of Object.entries(s.costs)){
    const st=s.stock[c];if(!st)continue;
    const map=s.sales.byCode[c]||{},daily=dates.map(k=>Number(map[k]||0)),l7=daily.slice(-7),l14=daily.slice(-14);
    const avg=a=>a.length?a.reduce((x,y)=>x+y,0)/a.length:0;
    const a7=avg(l7),a14=avg(l14),aa=avg(daily),dem=.5*a7+.3*a14+.2*aa;
    const stockCalc=Math.max(0,Number(st.stock||0)),target=dem*(cov+safe),need=Math.max(0,target-stockCalc),uxb=Math.max(1,Number(st.uxb||1));
    const suggested=dem>0&&need>0?Math.ceil((need-1e-9)/uxb)*uxb:0;
    out.push({code:c,desc:m.desc,rubro:m.rubro,provider:m.provider,stock:Number(st.stock||0),uxb,sold7:l7.reduce((x,y)=>x+y,0),demand:dem,coverage:dem>0?stockCalc/dem:null,suggested,soldTotal:daily.reduce((x,y)=>x+y,0)})
  }
  const salesDays=dayCount(s.sales.minDate,s.sales.maxDate);
  s.dataset=out;
  s.meta={articles:out.length,providers:new Set(out.map(x=>x.provider).filter(Boolean)).size,salesFrom:s.sales.minDate,salesTo:s.sales.maxDate,salesDays,coverage:cov,safety:safe,updatedAt:new Date().toISOString(),salesWindowOk:salesDays>=28&&salesDays<=32}
}
function parseUpload(type,buf){
  const rows=parseRows(buf);
  if(type==="costs"){
    const hi=headerRow(rows,[["Codigo","Código"],["Descripcion","Descripción"],["Rubro"],["Prove.","Proveedor"]]);if(hi<0)throw new Error("No encontré Código, Descripción, Rubro y Proveedor");
    const h=rows[hi],ic=idx(h,["Codigo","Código"]),id=idx(h,["Descripcion","Descripción"]),ir=idx(h,["Rubro"]),ip=idx(h,["Prove.","Proveedor","Prove"]),o={};
    for(let r=hi+1;r<rows.length;r++){const c=code(rows[r]?.[ic]);if(!/^\d+$/.test(c))continue;o[c]={desc:String(rows[r]?.[id]??"").trim(),rubro:String(rows[r]?.[ir]??"").trim(),provider:String(rows[r]?.[ip]??"").trim()}}
    return o
  }
  if(type==="stock"){
    const hi=headerRow(rows,[["Codigo","Código"],["UxB","U X B","UXB"],["Stock xU","Stock XU","Stock"]]);if(hi<0)throw new Error("No encontré Código, UxB y Stock xU");
    const h=rows[hi],ic=idx(h,["Codigo","Código"]),iu=idx(h,["UxB","U X B","UXB"]),is=idx(h,["Stock xU","Stock XU","Stock"]),o={};
    for(let r=hi+1;r<rows.length;r++){const c=code(rows[r]?.[ic]);if(!/^\d+$/.test(c))continue;o[c]={uxb:Math.max(1,num(rows[r]?.[iu],1)),stock:num(rows[r]?.[is],0)}}
    return o
  }
  if(type==="sales"){
    const hi=headerRow(rows,[["Fecha"],["Codigo","Código"],["Cantidad","Cant.","Cant"]]);if(hi<0)throw new Error("No encontré Fecha, Código y Cantidad");
    const h=rows[hi],idf=idx(h,["Fecha"]),ic=idx(h,["Codigo","Código"]),iq=idx(h,["Cantidad","Cant.","Cant"]);
    let cur="",min="",max="";const by={};
    for(let r=hi+1;r<rows.length;r++){const row=rows[r]||[];if(row[idf]!==null&&row[idf]!==undefined&&String(row[idf]).trim()!==""){const k=dateKey(row[idf]);if(k)cur=k}const c=code(row[ic]);if(!cur||!/^\d+$/.test(c))continue;if(!by[c])by[c]={};by[c][cur]=(by[c][cur]||0)+num(row[iq],0);if(!min||cur<min)min=cur;if(!max||cur>max)max=cur}
    if(!min)throw new Error("No encontré ventas con fecha válida");
    return {byCode:by,minDate:min,maxDate:max}
  }
  throw new Error("Tipo inválido")
}
function authorized(pin){return !!ADMIN_PIN&&String(pin||"")===ADMIN_PIN}
function makePdf(res,payload,s){
  const provider=String(payload.provider||""),edits=payload.items||[],um=new Map(edits.map(x=>[String(x.code),Number(x.units)||0]));
  const rows=(s.dataset||[]).filter(x=>x.provider===provider&&um.get(x.code)>0).sort((a,b)=>a.desc.localeCompare(b.desc,"es"));
  if(!rows.length)return json(res,400,{error:"No hay artículos para imprimir"});
  const doc=new PDFDocument({size:"A4",layout:"landscape",margin:28});res.writeHead(200,{"Content-Type":"application/pdf","Content-Disposition":'attachment; filename="pedido.pdf"'});doc.pipe(res);
  const providerName=provider.replace(/^\[[^\]]+\]\s*/,""),totalU=rows.reduce((a,x)=>a+um.get(x.code),0);
  const header=()=>{doc.fillColor("#0d4a34").font("Helvetica-Bold").fontSize(19).text("Órdenes de compra",28,25);doc.fillColor("#333").fontSize(11).text(providerName,430,28,{width:380,align:"right"});doc.font("Helvetica").fontSize(9).text("Fecha: "+new Date().toLocaleDateString("es-AR"),430,46,{width:380,align:"right"});doc.text("Artículos: "+rows.length+"   Unidades: "+totalU,28,58);doc.moveTo(28,78).lineTo(814,78).strokeColor("#b8c9c2").stroke()};
  header();let y=90;const drawHead=()=>{doc.font("Helvetica-Bold").fontSize(8).fillColor("#234").text("Código",30,y);doc.text("Descripción",85,y);doc.text("Rubro",355,y);doc.text("Stock",500,y);doc.text("UxB",550,y);doc.text("Sugerido",590,y);doc.text("Unidades",660,y);doc.text("Cobertura",735,y);y+=18};drawHead();
  for(const x of rows){if(y>545){doc.addPage();header();y=90;drawHead()}const u=um.get(x.code),pc=x.demand>0?(Math.max(0,x.stock)+u)/x.demand:null;doc.font("Helvetica").fontSize(7.5).fillColor("#111").text(x.code,30,y,{width:50});doc.text(x.desc,85,y,{width:260,height:12,ellipsis:true});doc.text(String(x.rubro||"").replace(/^\[[^\]]+\]\s*/,""),355,y,{width:135,height:12,ellipsis:true});doc.text(String(Math.round(x.stock*10)/10),500,y,{width:45,align:"right"});doc.text(String(x.uxb),550,y,{width:35,align:"right"});doc.text(String(x.suggested),590,y,{width:55,align:"right"});doc.font("Helvetica-Bold").text(String(u),660,y,{width:55,align:"right"});doc.font("Helvetica").text(pc===null?"—":pc.toFixed(1)+" d",735,y,{width:65,align:"right"});y+=18;doc.moveTo(28,y-5).lineTo(814,y-5).strokeColor("#e5ece9").stroke()}
  doc.end()
}

const server=http.createServer(async(req,res)=>{const u=new URL(req.url,"http://localhost");
  try{
    if(req.method==="GET"&&(u.pathname==="/"||u.pathname==="/index.html"))return file(res,"index.html","text/html; charset=utf-8");
    if(req.method==="GET"&&u.pathname==="/admin")return file(res,"admin.html","text/html; charset=utf-8");
    if(req.method==="GET"&&u.pathname==="/logo.jpg")return file(res,"logo.jpg","image/jpeg");
    if(req.method==="GET"&&u.pathname==="/health"){res.writeHead(200,{"Content-Type":"text/plain"});return res.end("ok")}
    if(req.method==="GET"&&u.pathname==="/api/data"){const s=await load();return json(res,200,{dataset:s.dataset||[],meta:s.meta||{},ready:!!(s.costs&&s.stock&&s.sales)})}
    if(req.method==="GET"&&u.pathname==="/api/status"){const s=await load();return json(res,200,{costs:s.costs?Object.keys(s.costs).length:0,stock:s.stock?Object.keys(s.stock).length:0,sales:s.sales?Object.keys(s.sales.byCode).length:0,meta:s.meta||{},database:pool?"postgres":"local"})}
    if(req.method==="POST"&&u.pathname==="/api/upload"){
      if(process.env.RENDER&&!pool)return json(res,503,{error:"Falta conectar DATABASE_URL en Render antes de cargar las bases"});
      if(!ADMIN_PIN)return json(res,503,{error:"ADMIN_PIN no configurado en el servidor"});
      const p=JSON.parse((await body(req)).toString("utf8"));if(!authorized(p.pin))return json(res,403,{error:"PIN incorrecto"});
      const buf=Buffer.from(p.dataBase64||"","base64");if(!buf.length)throw new Error("Archivo vacío");
      const parsed=parseUpload(p.type,buf),s=await load();s[p.type]=parsed;rebuild(s);await save(s);
      const count=p.type==="sales"?Object.keys(parsed.byCode).length:Object.keys(parsed).length;
      return json(res,200,{ok:true,type:p.type,count,ready:!!(s.costs&&s.stock&&s.sales),articles:s.dataset.length,meta:s.meta})
    }
    if(req.method==="POST"&&u.pathname==="/api/pdf"){const p=JSON.parse((await body(req,2*1024*1024)).toString("utf8"));return makePdf(res,p,await load())}
    res.writeHead(404,{"Content-Type":"text/plain; charset=utf-8"});res.end("No encontrado")
  }catch(e){console.error(e);if(!res.headersSent)return json(res,400,{error:String(e.message||e)});try{res.end()}catch(_){}}
});

initStore().then(()=>server.listen(PORT,"0.0.0.0",()=>console.log("Nuestro Market:",PORT))).catch(e=>{console.error("No se pudo iniciar:",e);process.exit(1)});
