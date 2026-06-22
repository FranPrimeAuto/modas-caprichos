import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

// ═══════════════════════════════════════════════════════════════════════════════
// CONFIGURACIÓN
// ═══════════════════════════════════════════════════════════════════════════════
const CONFIG = {
  API_URL: "https://api.anthropic.com/v1/messages",
  MODEL:   "claude-sonnet-4-5",
  MAX_IMG_SIZE: 5 * 1024 * 1024,
  LS_KEY: "mc_api_key",
};

const getApiKey = () => localStorage.getItem(CONFIG.LS_KEY) || "";
const saveApiKey = k => localStorage.setItem(CONFIG.LS_KEY, k.trim());

// ═══════════════════════════════════════════════════════════════════════════════
// TEMA
// ═══════════════════════════════════════════════════════════════════════════════
const T = {
  bg:"#FAF7F2", surface:"#FFFFFF", card:"#FFFDF9",
  border:"#EDE8DF", borderD:"#D9D0C0",
  rose:"#C9736A", roseDark:"#A85C54", roseLight:"#F2E8E6",
  gold:"#B8973A", goldLight:"#FBF4E3",
  sage:"#6A8C72", sageLight:"#EDF3EE",
  indigo:"#5B6ED4", indigoLight:"#ECEFFE",
  text:"#2C2416", textMid:"#6B5E4E", textSoft:"#A8998A", cream:"#F5F0E8",
  shadow:"rgba(44,36,22,0.08)",
};

// ═══════════════════════════════════════════════════════════════════════════════
// USUARIOS
// ═══════════════════════════════════════════════════════════════════════════════
const USUARIOS = [
  { user:"caprichos",  pass:"modas2025",   rol:"cliente",   nombre:"Modas Capricho's" },
  { user:"oraclya",    pass:"oraclya2025", rol:"consultor", nombre:"Oraclya Consultor" },
];

// ═══════════════════════════════════════════════════════════════════════════════
// CONSTANTES
// ═══════════════════════════════════════════════════════════════════════════════
const CATS    = ["Todas","Vestidos","Pantalones","Camisas","Chaquetas","Faldas","Abrigos","Accesorios","Otros"];
const PASOS_IA = ["Leyendo el documento…","Identificando artículos…","Extrayendo precios…","Calculando PVP…","Preparando catálogo…"];

// ═══════════════════════════════════════════════════════════════════════════════
// HELPERS FECHA
// ═══════════════════════════════════════════════════════════════════════════════
const desgFecha = s => { if(!s)return null; const p=s.split("/"); return p.length===3?{d:+p[0],m:+p[1],y:+p[2]}:null; };
const desgISO   = s => { if(!s)return null; const p=s.split("-"); return p.length===3?{d:+p[2],m:+p[1],y:+p[0]}:null; };
const enRango   = (vs,d,h) => {
  const v=desgFecha(vs); if(!v)return true;
  if(d){const di=desgISO(d);if(di&&(v.y<di.y||(v.y===di.y&&v.m<di.m)||(v.y===di.y&&v.m===di.m&&v.d<di.d)))return false;}
  if(h){const hi=desgISO(h);if(hi&&(v.y>hi.y||(v.y===hi.y&&v.m>hi.m)||(v.y===hi.y&&v.m===hi.m&&v.d>hi.d)))return false;}
  return true;
};
const fmt      = n  => new Intl.NumberFormat("es-ES",{style:"currency",currency:"EUR"}).format(n);
const hoyISO   = () => { const h=new Date(); return `${h.getFullYear()}-${String(h.getMonth()+1).padStart(2,"0")}-${String(h.getDate()).padStart(2,"0")}`; };
const hoyLocal = () => new Date().toLocaleDateString("es-ES");
const dAtras   = n  => { const d=new Date(); d.setDate(d.getDate()-n); return d.toLocaleDateString("es-ES"); };

// ═══════════════════════════════════════════════════════════════════════════════
// DATOS DEMO
// ═══════════════════════════════════════════════════════════════════════════════
const STOCK0 = [
  {id:1, nombre:"Vestido Floral Primavera",cat:"Vestidos",  talla:"M",    color:"Rosa",      pvp:49.95,coste:22.00,iva:21,stock:3,vendidas:8, proveedor:"Moda Sur S.L.",      fechaEntrada:"10/03/2025",emoji:"👗"},
  {id:2, nombre:"Pantalón Lino Beige",     cat:"Pantalones",talla:"L",    color:"Beige",     pvp:39.95,coste:16.00,iva:21,stock:5,vendidas:12,proveedor:"Tejidos Córdoba",    fechaEntrada:"15/03/2025",emoji:"👖"},
  {id:3, nombre:"Camisa Rayas Azul",       cat:"Camisas",   talla:"S",    color:"Azul",      pvp:29.95,coste:11.00,iva:21,stock:2,vendidas:5, proveedor:"Moda Sur S.L.",      fechaEntrada:"20/03/2025",emoji:"👚"},
  {id:4, nombre:"Chaqueta Punto Rosa",     cat:"Chaquetas", talla:"M",    color:"Rosa",      pvp:65.00,coste:28.00,iva:21,stock:4,vendidas:3, proveedor:"Punto España",        fechaEntrada:"01/04/2025",emoji:"🧥"},
  {id:5, nombre:"Falda Midi Estampada",    cat:"Faldas",    talla:"S",    color:"Verde",     pvp:35.95,coste:14.00,iva:21,stock:6,vendidas:15,proveedor:"Moda Sur S.L.",      fechaEntrada:"05/04/2025",emoji:"👗"},
  {id:6, nombre:"Abrigo Camel Clásico",    cat:"Abrigos",   talla:"L",    color:"Camel",     pvp:129.0,coste:55.00,iva:21,stock:1,vendidas:2, proveedor:"Tejidos Córdoba",    fechaEntrada:"10/02/2025",emoji:"🧥"},
  {id:7, nombre:"Bolso Piel Marrón",       cat:"Accesorios",talla:"Única",color:"Marrón",    pvp:55.00,coste:20.00,iva:21,stock:3,vendidas:6, proveedor:"Complementos España",fechaEntrada:"12/04/2025",emoji:"👜"},
  {id:8, nombre:"Vestido Negro Elegante",  cat:"Vestidos",  talla:"M",    color:"Negro",     pvp:79.95,coste:32.00,iva:21,stock:2,vendidas:4, proveedor:"Moda Sur S.L.",      fechaEntrada:"18/04/2025",emoji:"👗"},
  {id:9, nombre:"Pantalón Vaquero Slim",   cat:"Pantalones",talla:"S",    color:"Azul",      pvp:45.95,coste:18.00,iva:21,stock:7,vendidas:20,proveedor:"Denim España",        fechaEntrada:"22/03/2025",emoji:"👖"},
  {id:10,nombre:"Pañuelo Seda Flores",     cat:"Accesorios",talla:"Única",color:"Multicolor",pvp:22.00,coste:8.00, iva:21,stock:8,vendidas:9, proveedor:"Complementos España",fechaEntrada:"28/04/2025",emoji:"🧣"},
];
const VENTAS0 = [
  {id:1,fecha:dAtras(0),hora:"10:23",cliente:"María García",   prendas:[{id:1,nombre:"Vestido Floral Primavera",pvp:49.95,qty:1}],total:49.95,iva:8.68,base:41.27,metodo:"Tarjeta"},
  {id:2,fecha:dAtras(0),hora:"12:45",cliente:"Ana Martínez",   prendas:[{id:5,nombre:"Falda Midi Estampada",pvp:35.95,qty:1},{id:3,nombre:"Camisa Rayas Azul",pvp:29.95,qty:1}],total:65.90,iva:11.40,base:54.50,metodo:"Efectivo"},
  {id:3,fecha:dAtras(1),hora:"17:10",cliente:"Carmen López",   prendas:[{id:4,nombre:"Chaqueta Punto Rosa",pvp:65.00,qty:1}],total:65.00,iva:11.27,base:53.73,metodo:"Tarjeta"},
  {id:4,fecha:dAtras(2),hora:"11:30",cliente:"Rosa Fernández", prendas:[{id:9,nombre:"Pantalón Vaquero Slim",pvp:45.95,qty:1},{id:10,nombre:"Pañuelo Seda Flores",pvp:22.00,qty:1}],total:67.95,iva:11.77,base:56.18,metodo:"Bizum"},
  {id:5,fecha:dAtras(3),hora:"16:00",cliente:"Pilar Ruiz",     prendas:[{id:7,nombre:"Bolso Piel Marrón",pvp:55.00,qty:1}],total:55.00,iva:9.54,base:45.46,metodo:"Tarjeta"},
  {id:6,fecha:dAtras(5),hora:"09:45",cliente:"Lola Jiménez",   prendas:[{id:8,nombre:"Vestido Negro Elegante",pvp:79.95,qty:1}],total:79.95,iva:13.87,base:66.08,metodo:"Tarjeta"},
  {id:7,fecha:dAtras(8),hora:"18:00",cliente:"María García",   prendas:[{id:2,nombre:"Pantalón Lino Beige",pvp:39.95,qty:1}],total:39.95,iva:6.93,base:33.02,metodo:"Efectivo"},
  {id:8,fecha:dAtras(12),hora:"11:00",cliente:"Encarna Moreno",prendas:[{id:6,nombre:"Abrigo Camel Clásico",pvp:129.00,qty:1}],total:129.0,iva:22.37,base:106.63,metodo:"Tarjeta"},
  {id:9,fecha:dAtras(15),hora:"16:30",cliente:"Ana Martínez",  prendas:[{id:1,nombre:"Vestido Floral Primavera",pvp:49.95,qty:1},{id:10,nombre:"Pañuelo Seda Flores",pvp:22.00,qty:1}],total:71.95,iva:12.47,base:59.48,metodo:"Bizum"},
  {id:10,fecha:dAtras(20),hora:"10:00",cliente:"Carmen López", prendas:[{id:5,nombre:"Falda Midi Estampada",pvp:35.95,qty:2}],total:71.90,iva:12.46,base:59.44,metodo:"Efectivo"},
];

// ═══════════════════════════════════════════════════════════════════════════════
// CSS
// ═══════════════════════════════════════════════════════════════════════════════
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,600;0,700;1,600&family=Nunito:wght@400;500;600;700&display=swap');
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
html{-webkit-text-size-adjust:100%;touch-action:manipulation;}
body{background:${T.bg};color:${T.text};font-family:'Nunito',sans-serif;font-size:15px;}
input,select,textarea,button{font-family:'Nunito',sans-serif;-webkit-appearance:none;}
::-webkit-scrollbar{width:4px;height:4px;}
::-webkit-scrollbar-thumb{background:${T.borderD};border-radius:2px;}
@keyframes fadeUp{from{opacity:0;transform:translateY(8px);}to{opacity:1;transform:none;}}
@keyframes fadeIn{from{opacity:0;}to{opacity:1;}}
@keyframes spin{to{transform:rotate(360deg);}}
@keyframes pulse{0%,100%{opacity:1;}50%{opacity:.4;}}
@keyframes slideUp{from{opacity:0;transform:translateY(24px);}to{opacity:1;transform:none;}}
@keyframes slideIn{from{transform:translateX(110%);opacity:0;}to{transform:none;opacity:1;}}
@keyframes pop{0%{transform:scale(.92);}70%{transform:scale(1.03);}100%{transform:scale(1);}}
.fade-up{animation:fadeUp .28s ease both;}
.btn-rose{background:${T.rose};color:#fff;border:none;border-radius:12px;padding:12px 20px;font-weight:700;font-size:14px;cursor:pointer;transition:all .18s;-webkit-tap-highlight-color:transparent;}
.btn-rose:hover{background:${T.roseDark};transform:translateY(-1px);}
.btn-rose:active{transform:scale(.96);}
.btn-rose:disabled{opacity:.4;cursor:not-allowed;transform:none;}
.btn-sage{background:${T.sage};color:#fff;border:none;border-radius:10px;padding:10px 16px;font-weight:600;font-size:13px;cursor:pointer;transition:all .18s;-webkit-tap-highlight-color:transparent;}
.btn-sage:hover{filter:brightness(1.08);transform:translateY(-1px);}
.btn-sage:active{transform:scale(.96);}
.btn-indigo{background:${T.indigo};color:#fff;border:none;border-radius:12px;padding:11px 18px;font-weight:700;font-size:13px;cursor:pointer;transition:all .18s;-webkit-tap-highlight-color:transparent;}
.btn-indigo:hover{filter:brightness(1.08);transform:translateY(-1px);}
.btn-indigo:disabled{opacity:.4;cursor:not-allowed;transform:none;}
.btn-ghost{background:${T.cream};color:${T.textMid};border:1.5px solid ${T.border};border-radius:10px;padding:9px 14px;font-weight:500;font-size:13px;cursor:pointer;transition:all .18s;-webkit-tap-highlight-color:transparent;}
.btn-ghost:hover{border-color:${T.borderD};background:${T.border};}
.card{background:${T.card};border:1.5px solid ${T.border};border-radius:16px;box-shadow:0 2px 12px ${T.shadow};}
.tab-btn{background:none;border:none;padding:8px 14px;border-radius:9px;font-size:13px;font-weight:500;color:${T.textMid};cursor:pointer;white-space:nowrap;transition:all .15s;-webkit-tap-highlight-color:transparent;}
.tab-btn:hover{background:${T.cream};color:${T.text};}
.tab-btn.active{background:${T.rose};color:#fff;}
.input-field{width:100%;padding:11px 13px;border:1.5px solid ${T.border};border-radius:10px;background:${T.surface};color:${T.text};font-size:15px;outline:none;transition:border-color .15s;}
.input-field:focus{border-color:${T.rose};}
.date-input{padding:9px 11px;border:1.5px solid ${T.border};border-radius:9px;background:${T.surface};color:${T.text};font-size:13px;outline:none;font-family:'Nunito',sans-serif;-webkit-appearance:none;}
.date-input:focus{border-color:${T.rose};}
table{width:100%;border-collapse:collapse;font-size:13px;}
th{text-align:left;padding:9px 12px;font-size:10px;font-weight:700;color:${T.textSoft};letter-spacing:.07em;text-transform:uppercase;border-bottom:1.5px solid ${T.border};background:${T.cream};white-space:nowrap;}
td{padding:10px 12px;border-bottom:1px solid ${T.border};color:${T.textMid};vertical-align:middle;}
tr:last-child td{border-bottom:none;}
tr:hover td{background:${T.cream};}
.prenda-card{background:${T.card};border:1.5px solid ${T.border};border-radius:14px;padding:13px;cursor:pointer;transition:all .18s;-webkit-tap-highlight-color:transparent;}
.prenda-card:hover{border-color:${T.rose};transform:translateY(-2px);}
.prenda-card:active{transform:scale(.96);}
.drop-zone{border:2px dashed ${T.borderD};border-radius:12px;padding:18px;text-align:center;transition:all .18s;background:${T.cream};}
.drop-zone.over,.drop-zone:hover{border-color:${T.rose};background:${T.roseLight};}
.modal-overlay{position:fixed;inset:0;background:rgba(44,36,22,.55);z-index:500;display:flex;align-items:center;justify-content:center;padding:14px;backdrop-filter:blur(4px);-webkit-backdrop-filter:blur(4px);}
.modal-box{background:${T.surface};border-radius:20px;box-shadow:0 20px 60px rgba(44,36,22,.25);max-width:720px;width:100%;max-height:90vh;overflow-y:auto;animation:slideUp .28s ease;}
@media(max-width:600px){.hide-m{display:none!important;}.tab-btn{padding:7px 10px;font-size:12px;}}
`;

// ═══════════════════════════════════════════════════════════════════════════════
// ATOMS
// ═══════════════════════════════════════════════════════════════════════════════
const Spin = ({s=18,c=T.rose}) => <span style={{display:"inline-block",width:s,height:s,border:`2.5px solid ${T.border}`,borderTopColor:c,borderRadius:"50%",animation:"spin .75s linear infinite",flexShrink:0}}/>;
const Badge = ({children,color="rose"}) => {
  const m={rose:{bg:T.roseLight,c:T.rose},gold:{bg:T.goldLight,c:T.gold},sage:{bg:T.sageLight,c:T.sage},indigo:{bg:T.indigoLight,c:T.indigo},gray:{bg:T.cream,c:T.textSoft}};
  const s=m[color]||m.gray;
  return <span style={{display:"inline-block",padding:"2px 9px",borderRadius:20,fontSize:11,fontWeight:700,background:s.bg,color:s.c,whiteSpace:"nowrap"}}>{children}</span>;
};
const Toast = ({msg,type="success",onClose}) => {
  useEffect(()=>{const t=setTimeout(onClose,4500);return()=>clearTimeout(t);},[]);
  return <div style={{position:"fixed",bottom:20,left:12,right:12,zIndex:9999,background:type==="error"?T.rose:T.sage,color:"#fff",padding:"13px 16px",borderRadius:13,fontSize:14,fontWeight:600,animation:"slideIn .28s ease",boxShadow:"0 8px 28px rgba(0,0,0,.2)",textAlign:"center",maxWidth:420,margin:"0 auto"}}>{type==="error"?"⚠ ":"✓ "}{msg}</div>;
};

// ═══════════════════════════════════════════════════════════════════════════════
// LOGO
// ═══════════════════════════════════════════════════════════════════════════════
const Logo = ({big=false}) => (
  <div style={{display:"flex",alignItems:"center",gap:big?14:9}}>
    <svg width={big?58:40} height={big?58:40} viewBox="0 0 46 46" fill="none">
      <circle cx="23" cy="23" r="22" fill="white" stroke="#D6E8E4" strokeWidth="1.2"/>
      <ellipse cx="14" cy="10" rx="4"   ry="2"   fill="#7AAB8A" transform="rotate(-30 14 10)" opacity=".85"/>
      <ellipse cx="23" cy="7"  rx="4"   ry="1.5" fill="#8AB89A" transform="rotate(5 23 7)"   opacity=".8"/>
      <ellipse cx="32" cy="10" rx="4"   ry="2"   fill="#7AAB8A" transform="rotate(30 32 10)" opacity=".85"/>
      <ellipse cx="10" cy="15" rx="3"   ry="1.2" fill="#9BBD7A" transform="rotate(-50 10 15)"opacity=".7"/>
      <ellipse cx="36" cy="15" rx="3"   ry="1.2" fill="#9BBD7A" transform="rotate(50 36 15)" opacity=".7"/>
      <circle  cx="19" cy="8"  r="1.5" fill="#B8D4E8" opacity=".9"/>
      <circle  cx="27" cy="7"  r="1.2" fill="#C4B8D8" opacity=".85"/>
      <circle  cx="33" cy="12" r="1.3" fill="#B8D4E8" opacity=".8"/>
      <ellipse cx="10" cy="33" rx="3.5" ry="1.3" fill="#7AAB8A" transform="rotate(40 10 33)"  opacity=".75"/>
      <ellipse cx="36" cy="33" rx="3.5" ry="1.3" fill="#7AAB8A" transform="rotate(-40 36 33)" opacity=".75"/>
      <ellipse cx="23" cy="36" rx="2.5" ry="1.5" fill="#A8CFC8" transform="rotate(-15 23 36)" opacity=".7"/>
      <text x="23" y="27" textAnchor="middle" fontFamily="Georgia,serif" fontSize="9.5" fontStyle="italic" fontWeight="600" fill="#6AABA0" letterSpacing="0.3">Capricho's</text>
      <line x1="11" y1="29.5" x2="35" y2="29.5" stroke="#6AABA0" strokeWidth="0.8" opacity=".6"/>
    </svg>
    <div>
      <div style={{fontFamily:"Georgia,serif",fontStyle:"italic",fontWeight:600,fontSize:big?26:18,color:"#6AABA0",lineHeight:1.1}}>Capricho's</div>
      <div style={{fontSize:big?10:8,color:T.textSoft,letterSpacing:".1em",opacity:.65,marginTop:1}}>powered by ORACLYA</div>
    </div>
  </div>
);

// ═══════════════════════════════════════════════════════════════════════════════
// MODAL API KEY
// ═══════════════════════════════════════════════════════════════════════════════
function ModalApiKey({onSave,onClose,canClose=true}) {
  const [k,setK]=useState(getApiKey());
  const save=()=>{if(!k.trim())return;saveApiKey(k);onSave();};
  return(
    <div className="modal-overlay" onClick={e=>{if(canClose&&e.target===e.currentTarget)onClose();}}>
      <div className="modal-box" style={{maxWidth:480}}>
        <div style={{padding:"22px 24px 18px",borderBottom:`1.5px solid ${T.border}`}}>
          <div style={{fontFamily:"Georgia,serif",fontStyle:"italic",fontSize:20,fontWeight:700,color:T.text,marginBottom:4}}>🔑 Clave de la IA</div>
          <div style={{fontSize:13,color:T.textSoft}}>Necesitas una clave de Anthropic para usar el escáner de documentos.</div>
        </div>
        <div style={{padding:"20px 24px"}}>
          <div style={{background:T.indigoLight,border:`1px solid ${T.indigo}30`,borderRadius:10,padding:"12px 14px",marginBottom:18}}>
            <div style={{fontSize:12,fontWeight:700,color:T.indigo,marginBottom:4}}>¿Cómo obtenerla?</div>
            <ol style={{fontSize:12,color:T.textMid,paddingLeft:16,lineHeight:1.8}}>
              <li>Ve a <strong>console.anthropic.com</strong></li>
              <li>Crea una cuenta (gratis)</li>
              <li>Menú <strong>API Keys → Create Key</strong></li>
              <li>Copia la clave y pégala aquí</li>
            </ol>
            <div style={{marginTop:8,fontSize:11,color:T.textSoft}}>Coste aproximado: <strong>0,003€ por escaneo</strong>. La clave se guarda solo en tu navegador.</div>
          </div>
          <label style={{fontSize:10,fontWeight:700,color:T.textSoft,letterSpacing:".06em",display:"block",marginBottom:5}}>TU CLAVE API</label>
          <input className="input-field" type="password" placeholder="sk-ant-api03-…" value={k} onChange={e=>setK(e.target.value)}
            onKeyDown={e=>e.key==="Enter"&&save()} style={{fontFamily:"monospace",fontSize:13,letterSpacing:".03em"}}/>
          <div style={{display:"flex",gap:9,marginTop:16}}>
            {canClose&&<button className="btn-ghost" style={{flex:1}} onClick={onClose}>Cancelar</button>}
            <button className="btn-rose" style={{flex:2}} disabled={!k.trim()} onClick={save}>Guardar y continuar →</button>
          </div>
          {k&&<div style={{marginTop:10,fontSize:11,color:T.textSoft,textAlign:"center"}}>La clave se guarda en localStorage de este navegador. Nunca sale a nuestros servidores.</div>}
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// LOGIN
// ═══════════════════════════════════════════════════════════════════════════════
function LoginView({onLogin}) {
  const [u,setU]=useState(""); const [p,setP]=useState(""); const [err,setErr]=useState(""); const [loading,setL]=useState(false);
  const doLogin = () => {
    setErr(""); setL(true);
    setTimeout(()=>{
      const found = USUARIOS.find(x=>x.user===u.trim().toLowerCase()&&x.pass===p);
      found ? onLogin(found) : setErr("Usuario o contraseña incorrectos");
      setL(false);
    },500);
  };
  return (
    <div style={{minHeight:"100vh",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:"20px 16px",background:`linear-gradient(150deg,${T.roseLight},${T.bg} 50%,${T.sageLight})`}}>
      <div style={{width:"100%",maxWidth:360,animation:"fadeUp .4s ease"}}>
        <div style={{textAlign:"center",marginBottom:28}}><Logo big/></div>
        <div className="card" style={{padding:26}}>
          <div style={{fontFamily:"Georgia,serif",fontSize:20,fontWeight:700,color:T.text,textAlign:"center",marginBottom:20}}>Accede a tu tienda</div>
          {[["Usuario","text",u,setU,"username"],["Contraseña","password",p,setP,"current-password"]].map(([lab,type,val,set,ac])=>(
            <div key={lab} style={{marginBottom:13}}>
              <label style={{fontSize:10,fontWeight:700,color:T.textSoft,letterSpacing:".06em",display:"block",marginBottom:4}}>{lab.toUpperCase()}</label>
              <input className="input-field" type={type} autoCapitalize="none" autoCorrect="off" autoComplete={ac}
                placeholder={type==="password"?"••••••••":lab.toLowerCase()} value={val}
                onChange={e=>set(e.target.value)} onKeyDown={e=>e.key==="Enter"&&doLogin()}/>
            </div>
          ))}
          {err&&<div style={{fontSize:13,color:T.rose,background:T.roseLight,padding:"8px 12px",borderRadius:8,marginBottom:13,textAlign:"center"}}>{err}</div>}
          <button className="btn-rose" style={{width:"100%",fontSize:15,marginTop:4}} onClick={doLogin} disabled={loading||!u||!p}>
            {loading?<span style={{display:"flex",alignItems:"center",justifyContent:"center",gap:8}}><Spin s={15} c="#fff"/> Entrando…</span>:"Entrar →"}
          </button>
        </div>
        <div style={{textAlign:"center",marginTop:18,fontSize:11,color:T.textSoft,lineHeight:1.6}}>Sistema de gestión · Modas Capricho's<br/><span style={{opacity:.6}}>powered by Oraclya</span></div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// MOTOR ESCANEO IA
// Llama directamente a Anthropic API desde el navegador usando el header
// anthropic-dangerous-direct-browser-access: true (igual que Prime Auto)
// La API key se lee de localStorage — nunca pasa por nuestros servidores
// ═══════════════════════════════════════════════════════════════════════════════
const getMime = f => {
  const t = f.type||"";
  if(t && t!=="application/octet-stream") return t;
  const ext = (f.name||"").split(".").pop().toLowerCase();
  return ({jpg:"image/jpeg",jpeg:"image/jpeg",png:"image/png",webp:"image/webp",
    heic:"image/jpeg",heif:"image/jpeg",pdf:"application/pdf",gif:"image/gif",bmp:"image/bmp"})[ext]||"image/jpeg";
};

const toB64 = f => new Promise((ok,fail) => {
  const r = new FileReader();
  r.onload  = e => ok(e.target.result.split(",")[1]);
  r.onerror = () => fail(new Error("No se pudo leer el archivo. Inténtalo de nuevo."));
  r.readAsDataURL(f);
});

const parseJSON = str => {
  if(!str) return null;
  const clean = str.replace(/```json|```/gi,"").trim();
  for(const s of [clean, str]) {
    if(s.startsWith("[")) try{ return JSON.parse(s); }catch{}
    const m = s.match(/\[[\s\S]*\]/);
    if(m) try{ return JSON.parse(m[0]); }catch{}
  }
  return null;
};

const PROMPT_SCAN = `Eres un contable experto en moda española. Analiza el documento y extrae cada artículo de ropa o complemento.

PASO 1 — TIPO DE DOCUMENTO
Determina si es: albarán, factura, ticket, pedido, lista manual o email.

PASO 2 — COLUMNAS (pueden tener muchos nombres distintos)
• Referencia: Ref, Cód, SKU, Art.nº, Referencia
• Artículo: Artículo, Descripción, Producto, Concepto
• Cantidad: Cantid, Cant, Uds, Unidades, Q → campo "stock"
• Precio unitario: Pre/U, P.Unit, Precio, Tarifa, €/ud → campo "coste" (precio que paga la tienda)
• Descuento: Dto%, Desc → si >0, aplica: coste_real = precio × (1 - dto/100)
• Importe: Total, Subtotal → úsalo solo para verificar

PASO 3 — PARA CADA ARTÍCULO
• nombre: texto del artículo. Si es genérico (BLUSA, PANTALON) añade la ref: "Blusa ref.34026". Máx 50 chars.
• stock: cantidad exacta de unidades
• coste: precio unitario con descuento aplicado. 2 decimales.
• pvp: si el doc lo indica úsalo; si no → coste × 2.35
• cat: clasifica por nombre:
  Vestidos→vestido,mono,conjunto | Pantalones→pantalon,vaquero,legging,bermuda
  Camisas→camisa,camiseta,polo,blusa,top | Chaquetas→chaqueta,blazer,cardigan,rebeca
  Faldas→falda | Abrigos→abrigo,cazadora,anorak,trench
  Accesorios→bolso,cinturon,pañuelo,bufanda,gorro,collar,pulsera
  Otros→resto
• talla: extrae si aparece; si no → "Única"
• color: extrae si aparece; si no → "Varios"
• emoji: 👗Vestidos 👖Pantalones 👚Camisas 🧥Chaquetas 🩱Faldas 🧥Abrigos 👜Accesorios 🛍Otros
• ref: referencia numérica o alfanumérica
• proveedor_detectado: nombre del proveedor si aparece claramente; si no → ""

PASO 4 — IGNORA
TOTAL, SUMA, BASE IMPONIBLE, IVA, PORTES, TRANSPORTE, cabeceras, datos fiscales del cliente.

PASO 5 — SALIDA
Responde SOLO con el array JSON. Sin texto antes. Sin texto después. Sin markdown. Empieza con [ termina con ].
Si no hay artículos de ropa: []

Formato de cada objeto:
{"nombre":"","cat":"Vestidos|Pantalones|Camisas|Chaquetas|Faldas|Abrigos|Accesorios|Otros","talla":"XS|S|M|L|XL|XXL|Única","color":"","coste":0.00,"pvp":0.00,"iva":21,"stock":1,"emoji":"","ref":"","proveedor_detectado":""}`;

async function callScanAPI(content) {
  const apiKey = getApiKey();
  if(!apiKey) throw new Error("Configura tu clave de API de Anthropic primero.");

  const res = await fetch(CONFIG.API_URL, {
    method:"POST",
    headers:{
      "Content-Type":"application/json",
      "x-api-key": apiKey,
      "anthropic-version":"2023-06-01",
      "anthropic-dangerous-direct-browser-access":"true",
    },
    body: JSON.stringify({model:CONFIG.MODEL, max_tokens:2000, messages:[{role:"user",content}]})
  });
  if(!res.ok) {
    let msg = `Error ${res.status}`;
    try{ const e=await res.json(); msg=e?.error?.message||msg; }catch{}
    throw new Error(msg);
  }
  const data = await res.json();
  if(data.error) throw new Error(data.error.message||"Error de API");
  if(!data.content?.length) throw new Error("La API no devolvió contenido");
  return data.content.map(b=>b.text||"").join("").trim();
}

async function escanearArchivo(file, prov) {
  if(file.size > CONFIG.MAX_IMG_SIZE)
    throw new Error(`El archivo pesa ${(file.size/1024/1024).toFixed(1)}MB. Usa una foto de menor resolución.`);

  const mime   = getMime(file);
  const sufijo = prov ? `\n\nProveedor conocido: ${prov}` : "";
  const prompt = PROMPT_SCAN + sufijo;
  const b64    = await toB64(file);

  const mimeApi = mime==="application/pdf" ? "application/pdf"
    : (mime==="image/heic"||mime==="image/heif") ? "image/jpeg"
    : mime.startsWith("image/") ? mime : "image/jpeg";

  const content = mime==="application/pdf"
    ? [{type:"document",source:{type:"base64",media_type:"application/pdf",data:b64}},{type:"text",text:prompt}]
    : [{type:"image",  source:{type:"base64",media_type:mimeApi,            data:b64}},{type:"text",text:prompt}];

  const texto = await callScanAPI(content);
  const result = parseJSON(texto);
  if(!result) throw new Error(`Respuesta inesperada. Comprueba que la foto sea legible y bien iluminada.`);
  return result;
}

async function escanearTexto(texto, prov) {
  const prompt  = PROMPT_SCAN + (prov?`\n\nProveedor conocido: ${prov}`:"") + `\n\nTEXTO DEL DOCUMENTO:\n${texto}`;
  const rawText = await callScanAPI([{type:"text",text:prompt}]);
  const result  = parseJSON(rawText);
  if(!result) throw new Error("No se pudo interpretar el texto. Comprueba el formato.");
  return result;
}

const enriquecer = (items, prov) => items.map((p,i)=>({
  ...p, id:Date.now()+i,
  proveedor: p.proveedor_detectado||prov||"Proveedor",
  fechaEntrada: hoyLocal(), vendidas:0,
  pvp: p.pvp||parseFloat((p.coste*2.35).toFixed(2)),
  coste: p.coste||0,
  stock: p.stock||1,
}));

// ═══════════════════════════════════════════════════════════════════════════════
// MOTOR ANÁLISIS NEGOCIO (streaming)
// ═══════════════════════════════════════════════════════════════════════════════
async function analizarNegocio(stock, ventas, onChunk) {
  const apiKey = getApiKey();
  if(!apiKey) throw new Error("Configura tu clave de API de Anthropic primero.");

  const hoy=new Date(); const mes=hoy.getMonth()+1; const dia=hoy.getDate();
  const tv=ventas.reduce((s,v)=>s+v.total,0);
  const top5=[...stock].sort((a,b)=>b.vendidas-a.vendidas).slice(0,5);
  const bajo=stock.filter(p=>p.stock<3);
  const mg=(stock.reduce((s,p)=>s+(p.pvp-p.coste)/p.pvp*100,0)/stock.length).toFixed(1);
  const cliM={}; ventas.forEach(v=>{if(!cliM[v.cliente])cliM[v.cliente]={n:v.cliente,c:0,t:0};cliM[v.cliente].c++;cliM[v.cliente].t+=v.total;});
  const top3c=Object.values(cliM).sort((a,b)=>b.t-a.t).slice(0,3);
  const catM={}; stock.forEach(p=>{if(!catM[p.cat])catM[p.cat]={c:p.cat,u:0,e:0};catM[p.cat].u+=p.vendidas;catM[p.cat].e+=p.pvp*p.vendidas;});
  const cats=Object.values(catM).sort((a,b)=>b.e-a.e);
  const prompt=`Eres consultor de retail de moda especializado en pequeño comercio andaluz con conocimiento de Posadas (Córdoba) y su calendario festivo.

DATOS REALES — MODAS CAPRICHO'S, POSADAS (CÓRDOBA):
Fecha: ${dia}/${mes}/${hoy.getFullYear()} | Facturación: ${fmt(tv)} | ${ventas.length} ventas | Margen medio: ${mg}%
TOP 5: ${top5.map((p,i)=>`${i+1}.${p.nombre}(${p.cat}) ${p.vendidas}uds PVP${fmt(p.pvp)} stock:${p.stock}`).join(" | ")}
STOCK BAJO: ${bajo.length?bajo.map(p=>`${p.nombre}:${p.stock}`).join(", "):"Sin alertas"}
CATEGORÍAS: ${cats.map(c=>`${c.c}:${c.u}uds/${fmt(c.e)}`).join(" | ")}
TOP CLIENTAS: ${top3c.map((c,i)=>`${i+1}.${c.n}(${c.c}v/${fmt(c.t)})`).join(" | ")}
CONTEXTO: Mayo-Junio=bodas/comuniones/bautizos en Córdoba. Feria Posadas junio. Verano extremo→ropa ligera. Clientela local fidelizada.

Redacta un informe profesional:
## 🎯 Diagnóstico general
## 📈 Puntos fuertes
## 🔧 4 Acciones para el próximo mes
### 1. [título] **Qué hacer:** ... **Cómo:** ... **Resultado esperado:** ...
### 2. [título] [igual]
### 3. [título] [igual]
### 4. [título] [igual]
## 🗓️ Oportunidad inmediata — ${new Intl.DateTimeFormat("es-ES",{month:"long"}).format(hoy)}
## 💰 Impacto potencial estimado

Usa datos reales, menciona prendas concretas. Tono profesional pero cercano.`;

  const resp=await fetch(CONFIG.API_URL,{method:"POST",
    headers:{
      "Content-Type":"application/json",
      "x-api-key": apiKey,
      "anthropic-version":"2023-06-01",
      "anthropic-dangerous-direct-browser-access":"true",
    },
    body:JSON.stringify({model:CONFIG.MODEL,max_tokens:2000,stream:true,messages:[{role:"user",content:prompt}]})});
  const reader=resp.body.getReader(); const dec=new TextDecoder(); let buf="";
  while(true){
    const{done,value}=await reader.read(); if(done)break;
    buf+=dec.decode(value,{stream:true});
    const lines=buf.split("\n"); buf=lines.pop();
    for(const l of lines){
      if(!l.startsWith("data:"))continue;
      const d=l.slice(5).trim(); if(d==="[DONE]")return;
      try{const ev=JSON.parse(d);if(ev.type==="content_block_delta"&&ev.delta?.text)onChunk(ev.delta.text);}catch{}
    }
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// ZONA ESCANEO
// ═══════════════════════════════════════════════════════════════════════════════
function ZonaEscaneo({onFile}) {
  const [over,setOver]=useState(false);
  const onDrop=e=>{e.preventDefault();setOver(false);if(e.dataTransfer.files[0])onFile(e.dataTransfer.files[0]);};
  const idCam  = useRef(`cam-${Date.now()}`).current;
  const idFile = useRef(`file-${Date.now()}`).current;

  return (
    <div style={{display:"flex",flexDirection:"column",gap:10}}>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
        <label htmlFor={idCam} style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",
          gap:8,padding:"20px 10px",background:T.roseLight,border:`2px solid ${T.rose}`,borderRadius:14,
          cursor:"pointer",textAlign:"center",userSelect:"none",WebkitTapHighlightColor:"transparent"}}>
          <input id={idCam} type="file" accept="image/*" capture="environment"
            style={{position:"absolute",width:1,height:1,overflow:"hidden",opacity:0,zIndex:-1}}
            onChange={e=>{if(e.target.files?.[0]){onFile(e.target.files[0]);e.target.value="";}}}/>
          <span style={{fontSize:32,pointerEvents:"none"}}>📷</span>
          <div style={{pointerEvents:"none"}}>
            <div style={{fontWeight:700,fontSize:13,color:T.rose}}>Hacer foto</div>
            <div style={{fontSize:11,color:T.textSoft,marginTop:1}}>Ticket · Albarán · Factura</div>
          </div>
        </label>
        <label htmlFor={idFile} style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",
          gap:8,padding:"20px 10px",background:T.sageLight,border:`2px solid ${T.sage}`,borderRadius:14,
          cursor:"pointer",textAlign:"center",userSelect:"none",WebkitTapHighlightColor:"transparent"}}>
          <input id={idFile} type="file" accept="image/*,application/pdf,.heic,.heif"
            style={{position:"absolute",width:1,height:1,overflow:"hidden",opacity:0,zIndex:-1}}
            onChange={e=>{if(e.target.files?.[0]){onFile(e.target.files[0]);e.target.value="";}}}/>
          <span style={{fontSize:32,pointerEvents:"none"}}>📁</span>
          <div style={{pointerEvents:"none"}}>
            <div style={{fontWeight:700,fontSize:13,color:T.sage}}>Subir archivo</div>
            <div style={{fontSize:11,color:T.textSoft,marginTop:1}}>PDF · Galería · Drive</div>
          </div>
        </label>
      </div>
      <div className={`drop-zone hide-m ${over?"over":""}`}
        onDragOver={e=>{e.preventDefault();setOver(true);}}
        onDragLeave={()=>setOver(false)} onDrop={onDrop}>
        <div style={{fontSize:20,marginBottom:4}}>🖱️</div>
        <div style={{fontSize:12,color:T.textSoft}}>{over?"¡Suelta aquí!":"Arrastra el archivo aquí (ordenador)"}</div>
        <div style={{display:"flex",gap:4,justifyContent:"center",flexWrap:"wrap",marginTop:6}}>
          {["PDF","JPG","PNG","HEIC","Ticket","Albarán","Factura"].map(t=>(
            <span key={t} style={{fontSize:9,padding:"2px 6px",borderRadius:20,background:T.surface,border:`1px solid ${T.border}`,color:T.textSoft}}>{t}</span>
          ))}
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// MODAL ANÁLISIS CONSULTOR
// ═══════════════════════════════════════════════════════════════════════════════
function ModalAnalisis({stock,ventas,onClose}) {
  const [txt,setTxt]=useState(""); const [busy,setBusy]=useState(false); const [done,setDone]=useState(false);
  const ref=useRef();
  useEffect(()=>{run();},[]);
  const run=async()=>{setBusy(true);setTxt("");setDone(false);
    try{await analizarNegocio(stock,ventas,c=>{setTxt(t=>t+c);if(ref.current)ref.current.scrollTop=ref.current.scrollHeight;});setDone(true);}
    catch(e){setTxt("⚠️ Error al conectar. Inténtalo de nuevo.");setDone(true);}
    finally{setBusy(false);}
  };
  const renderMD=t=>t.split("\n").map((l,i)=>{
    if(l.startsWith("## "))  return <div key={i} style={{fontFamily:"Georgia,serif",fontSize:18,fontWeight:700,color:T.rose,marginTop:18,marginBottom:6,borderBottom:`2px solid ${T.roseLight}`,paddingBottom:4}}>{l.slice(3)}</div>;
    if(l.startsWith("### ")) return <div key={i} style={{fontFamily:"Georgia,serif",fontSize:15,fontWeight:700,color:T.indigo,marginTop:12,marginBottom:3}}>{l.slice(4)}</div>;
    const parts=l.split(/(\*\*[^*]+\*\*)/g).map((p,j)=>p.startsWith("**")?<strong key={j}>{p.slice(2,-2)}</strong>:p);
    if(l.startsWith("- ")) return <div key={i} style={{display:"flex",gap:6,marginBottom:3}}><span style={{color:T.rose,flexShrink:0}}>▸</span><span style={{fontSize:13,color:T.textMid,lineHeight:1.6}}>{parts.slice(1)}</span></div>;
    if(!l.trim()) return <div key={i} style={{height:4}}/>;
    return <p key={i} style={{fontSize:14,color:T.textMid,lineHeight:1.7,marginBottom:2}}>{parts}</p>;
  });
  return (
    <div className="modal-overlay" onClick={e=>{if(e.target===e.currentTarget)onClose();}}>
      <div className="modal-box">
        <div style={{padding:"18px 22px 14px",borderBottom:`1.5px solid ${T.border}`,background:`linear-gradient(135deg,${T.indigoLight},${T.surface})`,borderRadius:"20px 20px 0 0",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <div>
            <div style={{display:"flex",alignItems:"center",gap:8}}><span style={{fontSize:20}}>🧠</span><div style={{fontFamily:"Georgia,serif",fontWeight:700,fontSize:19,color:T.indigo}}>Análisis Consultor · Oraclya</div></div>
            <div style={{fontSize:12,color:T.textSoft,marginTop:2}}>Modas Capricho's · {hoyLocal()}</div>
          </div>
          <div style={{display:"flex",gap:8}}>
            {done&&<button className="btn-indigo" style={{fontSize:12,padding:"6px 12px"}} onClick={run}>↻</button>}
            <button onClick={onClose} style={{background:"none",border:"none",fontSize:22,cursor:"pointer",color:T.textSoft}}>×</button>
          </div>
        </div>
        <div ref={ref} style={{padding:"20px 22px",overflowY:"auto",maxHeight:"62vh"}}>
          {busy&&!txt&&<div style={{textAlign:"center",padding:"32px 0"}}><Spin s={26} c={T.indigo}/><div style={{fontFamily:"Georgia,serif",fontSize:16,color:T.indigo,marginTop:10}}>Analizando el negocio…</div></div>}
          {txt&&<div>{renderMD(txt)}{busy&&<span style={{display:"inline-block",width:7,height:15,background:T.indigo,marginLeft:2,borderRadius:2,animation:"pulse .7s infinite"}}/>}</div>}
        </div>
        {done&&<div style={{padding:"13px 22px",borderTop:`1.5px solid ${T.border}`,display:"flex",justifyContent:"flex-end",background:T.cream,borderRadius:"0 0 20px 20px"}}><button className="btn-ghost" onClick={onClose}>Cerrar</button></div>}
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// EXPORTAR EXCEL
// ═══════════════════════════════════════════════════════════════════════════════
function exportXLS(ventas,stock,etiq){
  const wb=XLSX.utils.book_new();
  const ws1=XLSX.utils.aoa_to_sheet([
    ["MODAS CAPRICHO'S — Ventas","","","","","","",""],
    [`Período: ${etiq}`],[`Generado: ${hoyLocal()}`],[],
    ["Fecha","Hora","Cliente","Artículos","Base(€)","IVA(€)","Total(€)","Cobro"],
    ...ventas.map(v=>[v.fecha,v.hora,v.cliente,v.prendas.map(p=>`${p.nombre}${p.qty>1?` x${p.qty}`:""}`).join(" | "),v.base.toFixed(2),v.iva.toFixed(2),v.total.toFixed(2),v.metodo]),
    [],["","","","TOTALES",ventas.reduce((s,v)=>s+v.base,0).toFixed(2),ventas.reduce((s,v)=>s+v.iva,0).toFixed(2),ventas.reduce((s,v)=>s+v.total,0).toFixed(2)],
  ]);
  ws1["!cols"]=[{wch:12},{wch:7},{wch:20},{wch:42},{wch:11},{wch:11},{wch:11},{wch:11}];
  XLSX.utils.book_append_sheet(wb,ws1,"Ventas");
  const cM={}; ventas.forEach(v=>{if(!cM[v.cliente])cM[v.cliente]={n:v.cliente,c:0,t:0,b:0,i:0};cM[v.cliente].c++;cM[v.cliente].t+=v.total;cM[v.cliente].b+=v.base;cM[v.cliente].i+=v.iva;});
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([["CLIENTAS"],[`Período: ${etiq}`],[],["Cliente","Visitas","Base(€)","IVA(€)","Total(€)"],...Object.values(cM).sort((a,b)=>b.t-a.t).map(c=>[c.n,c.c,c.b.toFixed(2),c.i.toFixed(2),c.t.toFixed(2)])]),"Clientas");
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([["STOCK"],[],["Artículo","Cat","Talla","Color","Coste(€)","PVP(€)","Oferta(€)","Margen%","Stock","Vendidas","Proveedor"],...stock.map(p=>[p.nombre,p.cat,p.talla,p.color,p.coste.toFixed(2),p.pvp.toFixed(2),p.pvpOferta?p.pvpOferta.toFixed(2):"",`${((p.pvp-p.coste)/p.pvp*100).toFixed(1)}%`,p.stock,p.vendidas,p.proveedor])]),"Stock");
  const tv=ventas.reduce((s,v)=>s+v.total,0);
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([["RESUMEN GESTOR"],[`Período: ${etiq}`],[`Generado: ${hoyLocal()}`],[],["Nº ventas",ventas.length],["Base imponible(€)",ventas.reduce((s,v)=>s+v.base,0).toFixed(2)],["IVA 21%(€)",ventas.reduce((s,v)=>s+v.iva,0).toFixed(2)],["TOTAL(€)",tv.toFixed(2)],[],["Tarjeta(€)",ventas.filter(v=>v.metodo==="Tarjeta").reduce((s,v)=>s+v.total,0).toFixed(2)],["Efectivo(€)",ventas.filter(v=>v.metodo==="Efectivo").reduce((s,v)=>s+v.total,0).toFixed(2)],["Bizum(€)",ventas.filter(v=>v.metodo==="Bizum").reduce((s,v)=>s+v.total,0).toFixed(2)],[],["Prendas catálogo",stock.length],["Uds. stock",stock.reduce((s,p)=>s+p.stock,0)],["Stock bajo(<3)",stock.filter(p=>p.stock<3).length]]),"Resumen Gestor");
  XLSX.writeFile(wb,`ModasCaprichos_${etiq.replace(/[\s\/]/g,"_")}.xlsx`);
}

// ═══════════════════════════════════════════════════════════════════════════════
// FILTRO FECHAS
// ═══════════════════════════════════════════════════════════════════════════════
function FiltroFechas({desde,hasta,onD,onH}){
  const hoy=hoyISO();
  const iso=d=>`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
  const ap=k=>{const h=new Date();
    if(k==="hoy"){onD(hoy);onH(hoy);}
    else if(k==="sem"){const l=new Date(h);l.setDate(h.getDate()-((h.getDay()+6)%7));onD(iso(l));onH(hoy);}
    else if(k==="mes"){onD(`${h.getFullYear()}-${String(h.getMonth()+1).padStart(2,"0")}-01`);onH(hoy);}
    else if(k==="prev"){const p=new Date(h.getFullYear(),h.getMonth()-1,1);const f=new Date(h.getFullYear(),h.getMonth(),0);onD(iso(p));onH(iso(f));}
    else{onD("");onH("");}
  };
  return(
    <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap",padding:"10px 13px",background:T.goldLight,borderRadius:11,border:`1px solid ${T.gold}30`,marginBottom:13}}>
      <span style={{fontSize:11,fontWeight:700,color:T.gold}}>📅</span>
      <div style={{display:"flex",gap:4,flexWrap:"wrap"}}>
        {[["todo","Todo"],["hoy","Hoy"],["sem","Semana"],["mes","Mes"],["prev","Mes ant."]].map(([k,l])=>(
          <button key={k} className="btn-ghost" style={{fontSize:11,padding:"4px 9px",borderColor:`${T.gold}44`,color:T.gold}} onClick={()=>ap(k)}>{l}</button>
        ))}
      </div>
      <div style={{display:"flex",alignItems:"center",gap:5,marginLeft:"auto",flexWrap:"wrap"}}>
        <input type="date" className="date-input" value={desde} max={hoy} onChange={e=>onD(e.target.value)}/>
        <span style={{fontSize:12,color:T.textSoft}}>—</span>
        <input type="date" className="date-input" value={hasta} max={hoy} onChange={e=>onH(e.target.value)}/>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// VISTA: CAJA
// ═══════════════════════════════════════════════════════════════════════════════
function CajaView({stock,onVenta}){
  const [cat,setCat]=useState("Todas");const [q,setQ]=useState("");const [cart,setCart]=useState([]);
  const [cli,setCli]=useState("");const [met,setMet]=useState("Tarjeta");const [ticket,setTicket]=useState(null);
  const pvpE=p=>p.pvpOferta??p.pvp;
  const add=p=>setCart(prev=>{const ex=prev.find(c=>c.id===p.id);if(ex)return prev.map(c=>c.id===p.id?{...c,qty:c.qty+1}:c);return[...prev,{...p,pvp:pvpE(p),qty:1}];});
  const rm=id=>setCart(prev=>prev.filter(c=>c.id!==id));
  const tot=cart.reduce((s,c)=>s+c.pvp*c.qty,0);
  const iva=tot*21/121; const base=tot-iva;
  const filtradas=stock.filter(p=>p.stock>0&&(cat==="Todas"||p.cat===cat)&&(q===""||p.nombre.toLowerCase().includes(q.toLowerCase())||p.color.toLowerCase().includes(q.toLowerCase())));
  const cobrar=()=>{
    if(!cart.length)return;
    const v={id:Date.now(),fecha:hoyLocal(),hora:new Date().toLocaleTimeString("es-ES",{hour:"2-digit",minute:"2-digit"}),
      cliente:cli||"Cliente ocasional",prendas:cart.map(c=>({id:c.id,nombre:c.nombre,pvp:c.pvp,qty:c.qty})),total:tot,iva,base,metodo:met};
    onVenta(v,cart);setTicket(v);setCart([]);setCli("");
  };
  if(ticket)return(
    <div style={{maxWidth:400,margin:"24px auto",animation:"pop .4s ease"}}>
      <div className="card" style={{padding:24,textAlign:"center"}}>
        <div style={{fontSize:40,marginBottom:8}}>🎉</div>
        <div style={{fontFamily:"Georgia,serif",fontSize:22,fontWeight:700,color:T.rose,marginBottom:3}}>¡Venta realizada!</div>
        <div style={{color:T.textSoft,fontSize:12,marginBottom:16}}>{ticket.fecha} · {ticket.hora}</div>
        <div style={{background:T.cream,borderRadius:11,padding:14,marginBottom:16,textAlign:"left"}}>
          <div style={{fontWeight:700,marginBottom:8}}>👤 {ticket.cliente}</div>
          {ticket.prendas.map((p,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",fontSize:13,color:T.textMid,marginBottom:3}}><span>{p.nombre}{p.qty>1?` ×${p.qty}`:""}</span><span>{fmt(p.pvp*p.qty)}</span></div>)}
          <div style={{borderTop:`1px solid ${T.border}`,marginTop:10,paddingTop:10}}>
            <div style={{display:"flex",justifyContent:"space-between",fontSize:11,color:T.textSoft,marginBottom:2}}><span>Base imponible</span><span>{fmt(ticket.base)}</span></div>
            <div style={{display:"flex",justifyContent:"space-between",fontSize:11,color:T.textSoft,marginBottom:8}}><span>IVA 21%</span><span>{fmt(ticket.iva)}</span></div>
            <div style={{display:"flex",justifyContent:"space-between",fontWeight:700,fontSize:18,color:T.rose}}><span>TOTAL</span><span>{fmt(ticket.total)}</span></div>
          </div>
          <div style={{marginTop:10,textAlign:"center"}}><Badge color="sage">{ticket.metodo}</Badge></div>
        </div>
        <button className="btn-rose" style={{width:"100%"}} onClick={()=>setTicket(null)}>Nueva venta →</button>
      </div>
    </div>
  );
  return(
    <div style={{display:"flex",gap:13,height:"calc(100dvh - 128px)",minHeight:400}}>
      <div style={{flex:1,display:"flex",flexDirection:"column",gap:9,minWidth:0}}>
        <div style={{position:"relative"}}><span style={{position:"absolute",left:11,top:"50%",transform:"translateY(-50%)",color:T.textSoft,pointerEvents:"none"}}>🔍</span><input className="input-field" style={{paddingLeft:34,fontSize:14}} placeholder="Buscar prenda, color…" value={q} onChange={e=>setQ(e.target.value)}/></div>
        <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>{CATS.map(c=><button key={c} className={`tab-btn ${cat===c?"active":""}`} style={{fontSize:12,padding:"5px 10px"}} onClick={()=>setCat(c)}>{c}</button>)}</div>
        <div style={{flex:1,overflowY:"auto",display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(145px,1fr))",gap:8,alignContent:"start",paddingBottom:8}}>
          {filtradas.map(p=>(
            <div key={p.id} className="prenda-card" onClick={()=>add(p)}>
              <div style={{fontSize:27,marginBottom:4,textAlign:"center"}}>{p.emoji}</div>
              <div style={{fontWeight:600,fontSize:12,color:T.text,marginBottom:2,lineHeight:1.3}}>{p.nombre}</div>
              <div style={{fontSize:11,color:T.textSoft,marginBottom:5}}>{p.talla} · {p.color}</div>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-end"}}>
                <div>{p.pvpOferta&&<div style={{fontSize:10,color:T.textSoft,textDecoration:"line-through"}}>{fmt(p.pvp)}</div>}<span style={{fontFamily:"Georgia,serif",fontSize:16,fontWeight:700,color:p.pvpOferta?T.gold:T.rose}}>{fmt(pvpE(p))}</span></div>
                <span style={{fontSize:10,color:p.stock<3?T.rose:T.textSoft,fontWeight:p.stock<3?700:400}}>{p.stock}ud{p.stock!==1?"s":""}</span>
              </div>
            </div>
          ))}
          {!filtradas.length&&<div style={{gridColumn:"1/-1",textAlign:"center",padding:28,color:T.textSoft}}><div style={{fontSize:26,marginBottom:4}}>🔍</div>Sin resultados</div>}
        </div>
      </div>
      <div style={{width:275,display:"flex",flexDirection:"column",gap:9,flexShrink:0}}>
        <div className="card" style={{flex:1,padding:14,display:"flex",flexDirection:"column",gap:9}}>
          <div style={{fontFamily:"Georgia,serif",fontWeight:700,fontSize:18,color:T.text}}>🛒 Venta actual</div>
          <div><label style={{fontSize:10,color:T.textSoft,fontWeight:700,letterSpacing:".06em",display:"block",marginBottom:4}}>CLIENTE</label><input className="input-field" style={{fontSize:13}} placeholder="Nombre (opcional)" value={cli} onChange={e=>setCli(e.target.value)}/></div>
          <div style={{flex:1,overflowY:"auto"}}>
            {!cart.length?<div style={{textAlign:"center",padding:"20px 0",color:T.textSoft,fontSize:13}}><div style={{fontSize:22,marginBottom:4}}>👆</div>Toca una prenda</div>
            :cart.map(c=><div key={c.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 0",borderBottom:`1px solid ${T.border}`}}>
              <div style={{flex:1,minWidth:0}}><div style={{fontSize:12,fontWeight:600,color:T.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.emoji} {c.nombre}</div><div style={{fontSize:11,color:T.textSoft}}>{c.talla} · {fmt(c.pvp)}</div></div>
              <div style={{display:"flex",alignItems:"center",gap:4,marginLeft:4}}><span style={{fontSize:12,fontWeight:700,color:T.rose}}>{fmt(c.pvp*c.qty)}</span><button onClick={()=>rm(c.id)} style={{background:"none",border:"none",color:T.textSoft,fontSize:15,cursor:"pointer",padding:"0 2px"}}>×</button></div>
            </div>)}
          </div>
          {cart.length>0&&<div style={{borderTop:`1.5px solid ${T.border}`,paddingTop:9}}>
            <div style={{display:"flex",justifyContent:"space-between",fontSize:11,color:T.textSoft,marginBottom:2}}><span>Base</span><span>{fmt(base)}</span></div>
            <div style={{display:"flex",justifyContent:"space-between",fontSize:11,color:T.textSoft,marginBottom:7}}><span>IVA 21%</span><span>{fmt(iva)}</span></div>
            <div style={{display:"flex",justifyContent:"space-between",fontWeight:700,fontSize:18,fontFamily:"Georgia,serif",color:T.text,marginBottom:9}}><span>Total</span><span style={{color:T.rose}}>{fmt(tot)}</span></div>
            <div style={{display:"flex",gap:4,marginBottom:9}}>
              {["Tarjeta","Efectivo","Bizum"].map(m=><button key={m} onClick={()=>setMet(m)} style={{flex:1,padding:"6px 2px",borderRadius:8,fontSize:11,fontWeight:600,border:`1.5px solid ${met===m?T.rose:T.border}`,background:met===m?T.roseLight:"transparent",color:met===m?T.rose:T.textSoft,cursor:"pointer",WebkitTapHighlightColor:"transparent"}}>{m==="Tarjeta"?"💳":m==="Efectivo"?"💵":"📱"} {m}</button>)}
            </div>
            <button className="btn-rose" style={{width:"100%",padding:"12px"}} onClick={cobrar}>Cobrar {fmt(tot)} →</button>
          </div>}
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// VISTA: STOCK
// ═══════════════════════════════════════════════════════════════════════════════
function StockView({stock,onUpdate,setToast,onNeedKey}){
  const [modo,setModo]=useState("lista");const [prov,setProv]=useState("");const [txt,setTxt]=useState("");
  const [busy,setBusy]=useState(false);const [paso,setPaso]=useState(0);const [prev,setPrev]=useState([]);
  const [archInfo,setArchInfo]=useState(null);const [edit,setEdit]=useState(null);

  const guardar=()=>{
    if(!edit)return;
    const v=parseFloat(edit.val);
    if(isNaN(v)||v<=0){setEdit(null);return;}
    onUpdate(s=>s.map(p=>p.id===edit.id?{...p,[edit.campo]:v}:p));
    setToast({msg:"Precio actualizado"});setEdit(null);
  };

  const procesarArchivo=async file=>{
    if(!getApiKey()){onNeedKey();return;}
    setArchInfo({nombre:file.name||"archivo",tipo:getMime(file)==="application/pdf"?"pdf":"img",src:null});
    if(getMime(file)!=="application/pdf"){
      const r=new FileReader(); r.onload=e=>setArchInfo(a=>({...a,src:e.target.result})); r.readAsDataURL(file);
    }
    setBusy(true);setPaso(0);
    const iv=setInterval(()=>setPaso(p=>Math.min(p+1,PASOS_IA.length-1)),900);
    try{
      const res=await escanearArchivo(file,prov);
      clearInterval(iv);
      if(!res.length){setToast({msg:"No se detectaron artículos. Prueba con mejor iluminación.",type:"error"});setBusy(false);return;}
      setPrev(enriquecer(res,prov));setModo("preview");
    }catch(e){clearInterval(iv);setToast({msg:e.message,type:"error"});}
    finally{setBusy(false);}
  };

  const procesarTexto=async()=>{
    if(!txt.trim())return;
    if(!getApiKey()){onNeedKey();return;}
    setBusy(true);setPaso(0);
    const iv=setInterval(()=>setPaso(p=>Math.min(p+1,PASOS_IA.length-1)),900);
    try{
      const res=await escanearTexto(txt,prov);
      clearInterval(iv);
      if(!res.length){setToast({msg:"No se detectaron artículos.",type:"error"});setBusy(false);return;}
      setPrev(enriquecer(res,prov));setModo("preview");
    }catch(e){clearInterval(iv);setToast({msg:e.message,type:"error"});}
    finally{setBusy(false);}
  };

  const confirmar=()=>{onUpdate(s=>[...s,...prev]);setPrev([]);setTxt("");setArchInfo(null);setModo("lista");setToast({msg:`${prev.length} artículos añadidos al stock`});};
  const actualizarPrev=(i,campo,valor)=>setPrev(p=>p.map((x,j)=>j===i?{...x,[campo]:valor}:x));
  const quitarOferta=id=>onUpdate(s=>s.map(p=>p.id===id?{...p,pvpOferta:undefined}:p));

  return(
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:15,gap:8,flexWrap:"wrap"}}>
        <div style={{fontFamily:"Georgia,serif",fontSize:21,fontWeight:700,color:T.text}}>Gestión de Stock</div>
        <div style={{display:"flex",gap:8}}><button className="btn-ghost" onClick={()=>setModo("lista")}>Ver stock</button><button className="btn-rose" onClick={()=>{setModo("ia");setArchInfo(null);}}>✨ Añadir con IA</button></div>
      </div>

      {modo==="ia"&&(
        <div className="card fade-up" style={{padding:20,marginBottom:15}}>
          <div style={{fontFamily:"Georgia,serif",fontSize:18,fontWeight:700,color:T.rose,marginBottom:3}}>✨ Añadir stock desde documento</div>
          <div style={{color:T.textSoft,fontSize:13,marginBottom:13}}>Haz una foto o sube el archivo. La IA lee cualquier albarán, factura, ticket o pedido.</div>
          {!getApiKey()&&(
            <div style={{background:T.goldLight,border:`1.5px solid ${T.gold}40`,borderRadius:10,padding:"12px 14px",marginBottom:14,display:"flex",alignItems:"center",gap:10}}>
              <span style={{fontSize:20}}>🔑</span>
              <div style={{flex:1}}>
                <div style={{fontSize:13,fontWeight:700,color:T.gold}}>Clave de API no configurada</div>
                <div style={{fontSize:12,color:T.textMid}}>Necesitas configurarla para usar el escáner IA.</div>
              </div>
              <button className="btn-ghost" style={{fontSize:12,borderColor:`${T.gold}44`,color:T.gold}} onClick={onNeedKey}>Configurar</button>
            </div>
          )}
          <div style={{marginBottom:12}}><label style={{fontSize:10,color:T.textSoft,fontWeight:700,letterSpacing:".06em",display:"block",marginBottom:4}}>PROVEEDOR (la IA lo detecta sola si aparece en el documento)</label><input className="input-field" style={{maxWidth:280,fontSize:14}} placeholder="Ej: Moda Sur S.L." value={prov} onChange={e=>setProv(e.target.value)}/></div>
          {busy?(
            <div style={{background:T.cream,borderRadius:11,padding:18}}>
              <div style={{display:"flex",alignItems:"center",gap:9,marginBottom:11}}><Spin/><span style={{fontWeight:600,color:T.rose}}>IA analizando el documento…</span></div>
              {PASOS_IA.map((p,i)=>(
                <div key={i} style={{display:"flex",alignItems:"center",gap:8,marginBottom:6,opacity:i<=paso?1:.2,transition:"opacity .4s",animation:i===paso?"pulse 1.4s infinite":"none"}}>
                  <span style={{width:19,height:19,borderRadius:"50%",flexShrink:0,background:i<paso?T.sage:i===paso?T.rose:T.border,display:"flex",alignItems:"center",justifyContent:"center",fontSize:9,color:"#fff",fontWeight:700}}>{i<paso?"✓":i+1}</span>
                  <span style={{fontSize:13,color:i===paso?T.text:T.textSoft}}>{p}</span>
                </div>
              ))}
            </div>
          ):(
            <div style={{display:"flex",flexDirection:"column",gap:10}}>
              <ZonaEscaneo onFile={procesarArchivo}/>
              {archInfo&&<div style={{display:"flex",alignItems:"center",gap:9,padding:"8px 12px",background:T.sageLight,borderRadius:9,border:`1px solid ${T.sage}30`}}>
                {archInfo.src&&<img src={archInfo.src} style={{width:40,height:40,objectFit:"cover",borderRadius:6,flexShrink:0}}/>}
                {archInfo.tipo==="pdf"&&<span style={{fontSize:24}}>📄</span>}
                <div><div style={{fontSize:13,fontWeight:600,color:T.text}}>{archInfo.nombre}</div><div style={{fontSize:11,color:T.sage}}>✓ Cargado</div></div>
              </div>}
              <div style={{display:"flex",alignItems:"center",gap:8}}><div style={{flex:1,height:1,background:T.border}}/><span style={{fontSize:11,color:T.textSoft}}>o pega el texto</span><div style={{flex:1,height:1,background:T.border}}/></div>
              <textarea className="input-field" style={{minHeight:100,resize:"vertical",lineHeight:1.6,fontSize:13}} placeholder={"Pega aquí el texto del albarán, factura o email del proveedor…"} value={txt} onChange={e=>setTxt(e.target.value)}/>
              <div style={{display:"flex",gap:8}}><button className="btn-ghost" onClick={()=>setModo("lista")}>Cancelar</button><button className="btn-rose" disabled={!txt.trim()} onClick={procesarTexto}>✨ Analizar texto</button></div>
            </div>
          )}
        </div>
      )}

      {modo==="preview"&&(
        <div className="card fade-up" style={{padding:18,marginBottom:15}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12,flexWrap:"wrap",gap:8}}>
            <div><div style={{fontFamily:"Georgia,serif",fontSize:18,fontWeight:700,color:T.sage}}>✓ {prev.length} artículos detectados</div><div style={{fontSize:12,color:T.textSoft,marginTop:1}}>Revisa y confirma</div></div>
            <div style={{display:"flex",gap:8}}><button className="btn-ghost" onClick={()=>setModo("ia")}>← Repetir</button><button className="btn-rose" onClick={confirmar}>✓ Añadir al stock</button></div>
          </div>
          <div style={{fontSize:12,color:T.textSoft,marginBottom:10,background:T.goldLight,padding:"8px 12px",borderRadius:8,border:`1px solid ${T.gold}30`}}>
            ✏️ Revisa y edita el <strong>PVP</strong> de cada artículo antes de añadirlo al stock.
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(185px,1fr))",gap:9}}>
            {prev.map((p,i)=>(
              <div key={i} style={{background:T.sageLight,borderRadius:11,padding:12,border:`1px solid ${T.sage}30`,animation:`fadeUp .3s ease ${i*.04}s both`}}>
                <div style={{fontSize:22,marginBottom:4}}>{p.emoji||"👗"}</div>
                <div style={{fontWeight:600,fontSize:12,color:T.text,marginBottom:2}}>{p.nombre}</div>
                <div style={{fontSize:11,color:T.textSoft,marginBottom:6}}>{p.cat} · {p.talla} · {p.color}</div>
                <div style={{fontSize:11,color:T.textSoft,marginBottom:6}}>Coste: <strong>{fmt(p.coste)}</strong> · {p.stock} uds</div>
                <label style={{fontSize:10,fontWeight:700,color:T.rose,letterSpacing:".05em",display:"block",marginBottom:3}}>PVP (€)</label>
                <input
                  type="number" min="0" step="0.01"
                  value={p.pvp}
                  onChange={e=>actualizarPrev(i,"pvp",parseFloat(e.target.value)||0)}
                  style={{width:"100%",padding:"6px 9px",border:`2px solid ${T.rose}`,borderRadius:7,fontSize:14,fontWeight:700,color:T.rose,fontFamily:"monospace",outline:"none",background:"#fff"}}
                />
              </div>
            ))}
          </div>
        </div>
      )}

      <div className="card" style={{padding:0,overflow:"hidden"}}>
        <div style={{padding:"11px 15px",borderBottom:`1.5px solid ${T.border}`,display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:6}}>
          <span style={{fontWeight:700,color:T.text}}>Catálogo · {stock.length} artículos</span>
          <span style={{fontSize:12,color:T.textSoft}}>Valor stock: {fmt(stock.reduce((s,p)=>s+p.coste*p.stock,0))}</span>
        </div>
        <div style={{overflowX:"auto"}}>
          <table>
            <thead><tr>
              <th></th><th>Artículo</th><th className="hide-m">Cat.</th><th>Talla</th><th className="hide-m">Color</th>
              <th>Coste</th><th>PVP</th><th>Oferta</th><th className="hide-m">Margen</th>
              <th style={{textAlign:"center"}}>Stock</th><th className="hide-m" style={{textAlign:"center"}}>Vend.</th>
            </tr></thead>
            <tbody>
              {stock.map(p=>{
                const mg=((p.pvp-p.coste)/p.pvp*100).toFixed(0);
                const alerta=p.stock<3;
                const editP=edit?.id===p.id&&edit?.campo==="pvpOferta";
                return(
                  <tr key={p.id} style={{background:alerta?`${T.rose}06`:"transparent"}}>
                    <td style={{fontSize:16,textAlign:"center"}}>{p.emoji}</td>
                    <td style={{fontWeight:600,color:T.text,maxWidth:150,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.nombre}</td>
                    <td className="hide-m"><Badge color="gray">{p.cat}</Badge></td>
                    <td style={{fontWeight:600}}>{p.talla}</td>
                    <td className="hide-m" style={{color:T.textSoft}}>{p.color}</td>
                    <td style={{fontFamily:"monospace",color:T.textSoft,whiteSpace:"nowrap"}}>{fmt(p.coste)}</td>
                    <td style={{whiteSpace:"nowrap"}}>{p.pvpOferta?<span style={{textDecoration:"line-through",color:T.textSoft,fontSize:11}}>{fmt(p.pvp)}</span>:<span style={{fontWeight:700,color:T.rose}}>{fmt(p.pvp)}</span>}</td>
                    <td style={{minWidth:120}}>
                      {editP?(
                        <div style={{display:"flex",alignItems:"center",gap:3}}>
                          <input type="number" min="0" step="0.01" autoFocus style={{width:65,padding:"3px 6px",border:`1.5px solid ${T.rose}`,borderRadius:6,fontSize:13,fontFamily:"monospace",outline:"none"}}
                            value={edit.val} onChange={e=>setEdit(v=>({...v,val:e.target.value}))}
                            onKeyDown={e=>{if(e.key==="Enter")guardar();if(e.key==="Escape")setEdit(null);}}/>
                          <button onClick={guardar} style={{background:T.sage,border:"none",color:"#fff",borderRadius:5,padding:"3px 6px",cursor:"pointer",fontSize:12,fontWeight:700}}>✓</button>
                          <button onClick={()=>setEdit(null)} style={{background:"none",border:"none",color:T.textSoft,cursor:"pointer",fontSize:13}}>×</button>
                        </div>
                      ):(
                        <div style={{display:"flex",alignItems:"center",gap:4,flexWrap:"wrap"}}>
                          {p.pvpOferta&&<span style={{fontWeight:800,color:T.gold,fontSize:13,fontFamily:"monospace"}}>{fmt(p.pvpOferta)}</span>}
                          <button onClick={()=>setEdit({id:p.id,campo:"pvpOferta",val:p.pvpOferta??p.pvp})}
                            style={{background:"none",border:`1px dashed ${T.borderD}`,color:T.textSoft,borderRadius:5,padding:"2px 6px",cursor:"pointer",fontSize:11}}>
                            {p.pvpOferta?"✏":"🏷 oferta"}
                          </button>
                          {p.pvpOferta&&<button onClick={()=>quitarOferta(p.id)} style={{background:"none",border:"none",color:T.textSoft,cursor:"pointer",fontSize:12}}>×</button>}
                        </div>
                      )}
                    </td>
                    <td className="hide-m"><Badge color={mg>=45?"sage":mg>=35?"gold":"rose"}>{mg}%</Badge></td>
                    <td style={{textAlign:"center"}}>{alerta?<span style={{display:"inline-flex",alignItems:"center",gap:3,background:T.rose,color:"#fff",fontWeight:800,fontSize:11,padding:"3px 9px",borderRadius:7,animation:"pulse 1.8s infinite"}}>⚠ {p.stock}</span>:<span style={{fontWeight:600,color:T.text}}>{p.stock}</span>}</td>
                    <td className="hide-m" style={{textAlign:"center",color:T.textSoft}}>{p.vendidas}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// VISTA: VENTAS
// ═══════════════════════════════════════════════════════════════════════════════
function VentasView({ventas}){
  const tv=ventas.reduce((s,v)=>s+v.total,0);
  return(
    <div>
      <div style={{display:"flex",gap:10,marginBottom:14,flexWrap:"wrap"}}>
        {[["💶","Facturación",fmt(tv),T.rose],["🧾","Base",fmt(ventas.reduce((s,v)=>s+v.base,0)),T.gold],["🏛","IVA",fmt(ventas.reduce((s,v)=>s+v.iva,0)),T.textMid],["🛍","Ventas",ventas.length,T.sage]].map(([ico,lab,val,col])=>(
          <div key={lab} className="card fade-up" style={{flex:1,minWidth:120,padding:15}}>
            <div style={{fontSize:17,marginBottom:3}}>{ico}</div>
            <div style={{fontSize:10,color:T.textSoft,fontWeight:700,letterSpacing:".06em",marginBottom:2}}>{lab}</div>
            <div style={{fontFamily:"Georgia,serif",fontSize:18,fontWeight:700,color:col}}>{val}</div>
          </div>
        ))}
      </div>
      <div className="card" style={{padding:0,overflow:"hidden"}}>
        <div style={{padding:"11px 15px",borderBottom:`1.5px solid ${T.border}`}}><span style={{fontWeight:700,color:T.text}}>Historial · {ventas.length} ventas</span></div>
        <div style={{overflowX:"auto"}}>
          <table>
            <thead><tr><th>Fecha</th><th className="hide-m">Hora</th><th>Cliente</th><th className="hide-m">Artículos</th><th className="hide-m">Base</th><th className="hide-m">IVA</th><th>Total</th><th>Cobro</th></tr></thead>
            <tbody>
              {[...ventas].reverse().map(v=>(
                <tr key={v.id}>
                  <td style={{fontWeight:600,whiteSpace:"nowrap"}}>{v.fecha}</td>
                  <td className="hide-m" style={{color:T.textSoft}}>{v.hora}</td>
                  <td style={{fontWeight:600,color:T.text}}>👤 {v.cliente}</td>
                  <td className="hide-m" style={{fontSize:12,color:T.textSoft,maxWidth:180,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{v.prendas.map(p=>`${p.nombre}${p.qty>1?` ×${p.qty}`:""}`).join(", ")}</td>
                  <td className="hide-m" style={{fontFamily:"monospace"}}>{fmt(v.base)}</td>
                  <td className="hide-m" style={{fontFamily:"monospace",color:T.gold}}>{fmt(v.iva)}</td>
                  <td style={{fontWeight:700,color:T.rose,fontFamily:"monospace",whiteSpace:"nowrap"}}>{fmt(v.total)}</td>
                  <td><Badge color={v.metodo==="Tarjeta"?"rose":v.metodo==="Efectivo"?"gold":"sage"}>{v.metodo==="Tarjeta"?"💳":v.metodo==="Efectivo"?"💵":"📱"} {v.metodo}</Badge></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// VISTA: DASHBOARD
// ═══════════════════════════════════════════════════════════════════════════════
function DashboardView({stock,ventas,esConsultor,onNeedKey}){
  const [desde,setD]=useState("");const [hasta,setH]=useState("");const [modalIA,setModal]=useState(false);
  const vf=ventas.filter(v=>enRango(v.fecha,desde,hasta));
  const etiq=desde||hasta?`${desde?new Date(desde+"T12:00").toLocaleDateString("es-ES"):"inicio"} — ${hasta?new Date(hasta+"T12:00").toLocaleDateString("es-ES"):"hoy"}`:"Todo el histórico";
  const tv=vf.reduce((s,v)=>s+v.total,0);const tiv=vf.reduce((s,v)=>s+v.iva,0);const tb=vf.reduce((s,v)=>s+v.base,0);const tk=vf.length?tv/vf.length:0;
  const vxp={};vf.forEach(v=>v.prendas.forEach(p=>{if(!vxp[p.nombre])vxp[p.nombre]={u:0,e:0};vxp[p.nombre].u+=p.qty;vxp[p.nombre].e+=p.pvp*p.qty;}));
  const scp=stock.map(p=>({...p,uP:vxp[p.nombre]?.u||0,eP:vxp[p.nombre]?.e||0}));
  const rV=[...scp].sort((a,b)=>b.uP-a.uP).slice(0,5);
  const rM=[...stock].sort((a,b)=>(b.pvp-b.coste)/b.pvp-(a.pvp-a.coste)/a.pvp).slice(0,5);
  const sb=stock.filter(p=>p.stock<3);
  const pC=CATS.filter(c=>c!=="Todas").map(c=>{const ps=scp.filter(p=>p.cat===c);return{cat:c,u:ps.reduce((s,p)=>s+p.uP,0),e:ps.reduce((s,p)=>s+p.eP,0)};}).filter(c=>c.e>0).sort((a,b)=>b.e-a.e);
  const cM={};vf.forEach(v=>{if(!cM[v.cliente])cM[v.cliente]={n:v.cliente,c:0,t:0};cM[v.cliente].c++;cM[v.cliente].t+=v.total;});
  const rC=Object.values(cM).sort((a,b)=>b.t-a.t);
  const mxV=rV[0]?.uP||1;const mxC=pC[0]?.e||1;

  const abrirAnalisis=()=>{
    if(!getApiKey()){onNeedKey();return;}
    setModal(true);
  };

  return(
    <div>
      {modalIA&&<ModalAnalisis stock={stock} ventas={vf.length?vf:ventas} onClose={()=>setModal(false)}/>}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:11,flexWrap:"wrap",gap:8}}>
        <div style={{fontFamily:"Georgia,serif",fontSize:21,fontWeight:700,color:T.text}}>Cuadro de Mandos</div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          {esConsultor&&<button className="btn-indigo" onClick={abrirAnalisis} style={{display:"flex",alignItems:"center",gap:6,fontSize:13}}>🧠 Análisis IA</button>}
          <button className="btn-sage" onClick={()=>exportXLS(vf.length?vf:ventas,stock,etiq)} style={{display:"flex",alignItems:"center",gap:5,fontSize:13}}>⬇ Excel</button>
        </div>
      </div>
      <FiltroFechas desde={desde} hasta={hasta} onD={setD} onH={setH}/>
      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:12,flexWrap:"wrap"}}>
        <span style={{fontSize:12,color:T.textSoft}}>Mostrando:</span>
        <span style={{fontSize:13,fontWeight:700,color:T.gold}}>{etiq}</span>
        <span style={{fontSize:12,color:T.textSoft}}>· {vf.length} ventas · {fmt(tv)}</span>
      </div>
      <div style={{display:"flex",gap:9,flexWrap:"wrap",marginBottom:12}}>
        {[["💶","Facturación",fmt(tv),T.rose],["🧾","Base",fmt(tb),T.gold],["🏛","IVA",fmt(tiv),T.textMid],["🛍","Ventas",vf.length,T.sage],["🎯","Ticket medio",fmt(tk),T.gold],["👥","Clientas",rC.length,T.sage]].map(([ico,lab,val,col],i)=>(
          <div key={lab} className="card fade-up" style={{flex:1,minWidth:105,padding:13,animationDelay:`${i*.05}s`}}>
            <div style={{fontSize:16,marginBottom:3}}>{ico}</div>
            <div style={{fontSize:10,color:T.textSoft,fontWeight:700,letterSpacing:".07em",marginBottom:2}}>{lab}</div>
            <div style={{fontFamily:"Georgia,serif",fontSize:17,fontWeight:700,color:col}}>{val}</div>
          </div>
        ))}
      </div>
      <div style={{display:"flex",gap:11,marginBottom:11,flexWrap:"wrap"}}>
        <div className="card" style={{flex:1,minWidth:240,padding:17}}>
          <div style={{fontFamily:"Georgia,serif",fontWeight:700,fontSize:16,color:T.text,marginBottom:11}}>🏆 Más vendidas</div>
          {rV.every(p=>p.uP===0)?<div style={{color:T.textSoft,fontSize:13,textAlign:"center",padding:14}}>Sin ventas en este período</div>
          :rV.filter(p=>p.uP>0).map((p,i)=>(
            <div key={p.id} style={{marginBottom:10}}>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:3,alignItems:"center"}}>
                <div style={{display:"flex",alignItems:"center",gap:6}}>
                  <span style={{width:19,height:19,borderRadius:"50%",flexShrink:0,background:i===0?T.gold:i===1?"#C0C0C0":i===2?"#CD7F32":T.cream,display:"flex",alignItems:"center",justifyContent:"center",fontSize:9,fontWeight:800,color:i<3?"#fff":T.textSoft}}>{i+1}</span>
                  <span style={{fontSize:12,fontWeight:600,color:T.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:150}}>{p.emoji} {p.nombre}</span>
                </div>
                <span style={{fontSize:12,fontWeight:700,color:T.rose,flexShrink:0,marginLeft:4}}>{p.uP} uds</span>
              </div>
              <div style={{height:4,background:T.cream,borderRadius:2}}><div style={{width:`${(p.uP/mxV)*100}%`,height:"100%",background:i===0?T.gold:T.rose,borderRadius:2}}/></div>
            </div>
          ))}
        </div>
        <div className="card" style={{flex:1,minWidth:220,padding:17}}>
          <div style={{fontFamily:"Georgia,serif",fontWeight:700,fontSize:16,color:T.text,marginBottom:11}}>💎 Más rentables</div>
          {rM.map((p,i)=>{const mg=((p.pvp-p.coste)/p.pvp*100).toFixed(1);return(
            <div key={p.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 0",borderBottom:`1px solid ${T.border}`}}>
              <div style={{overflow:"hidden",flex:1,minWidth:0}}><div style={{fontSize:12,fontWeight:600,color:T.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.emoji} {p.nombre}</div><div style={{fontSize:11,color:T.textSoft}}>{fmt(p.coste)} → {fmt(p.pvp)}</div></div>
              <Badge color={mg>=50?"sage":mg>=40?"gold":"rose"}>{mg}%</Badge>
            </div>
          );})}
        </div>
      </div>
      <div style={{display:"flex",gap:11,marginBottom:11,flexWrap:"wrap"}}>
        <div className="card" style={{flex:1,minWidth:220,padding:17}}>
          <div style={{fontFamily:"Georgia,serif",fontWeight:700,fontSize:16,color:T.text,marginBottom:11}}>📊 Por categoría</div>
          {!pC.length?<div style={{color:T.textSoft,fontSize:13,textAlign:"center",padding:14}}>Sin ventas en este período</div>
          :pC.map(c=>(
            <div key={c.cat} style={{marginBottom:9}}>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}><span style={{fontSize:12,fontWeight:600,color:T.text}}>{c.cat}</span><div style={{display:"flex",gap:8}}><span style={{fontSize:11,color:T.textSoft}}>{c.u} uds</span><span style={{fontSize:12,fontWeight:700,color:T.rose}}>{fmt(c.e)}</span></div></div>
              <div style={{height:4,background:T.cream,borderRadius:2}}><div style={{width:`${(c.e/mxC)*100}%`,height:"100%",background:T.rose,borderRadius:2,opacity:.75}}/></div>
            </div>
          ))}
        </div>
        <div className="card" style={{flex:1,minWidth:200,padding:17}}>
          <div style={{fontFamily:"Georgia,serif",fontWeight:700,fontSize:16,color:T.text,marginBottom:11}}>👑 Mejores clientas</div>
          {!rC.length?<div style={{color:T.textSoft,fontSize:13,textAlign:"center",padding:14}}>Sin ventas en este período</div>
          :rC.map((c,i)=>(
            <div key={c.n} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"7px 0",borderBottom:`1px solid ${T.border}`}}>
              <div style={{display:"flex",alignItems:"center",gap:7}}>
                <span style={{width:23,height:23,borderRadius:"50%",flexShrink:0,background:i===0?T.goldLight:T.cream,display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:800,color:i===0?T.gold:T.textSoft}}>{i===0?"👑":i+1}</span>
                <div><div style={{fontSize:13,fontWeight:600,color:T.text}}>{c.n}</div><div style={{fontSize:11,color:T.textSoft}}>{c.c} visita{c.c!==1?"s":""}</div></div>
              </div>
              <span style={{fontWeight:700,color:T.rose,fontFamily:"Georgia,serif",fontSize:14,flexShrink:0,marginLeft:6}}>{fmt(c.t)}</span>
            </div>
          ))}
        </div>
      </div>
      {sb.length>0&&(
        <div className="card" style={{padding:15,border:`1.5px solid ${T.rose}40`,background:T.roseLight}}>
          <div style={{fontFamily:"Georgia,serif",fontWeight:700,fontSize:15,color:T.rose,marginBottom:9}}>⚠️ Reponer pronto</div>
          <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
            {sb.map(p=><div key={p.id} style={{background:"#fff",borderRadius:8,padding:"7px 11px",border:`1px solid ${T.rose}30`,display:"flex",alignItems:"center",gap:6}}><span style={{fontSize:15}}>{p.emoji}</span><div><div style={{fontSize:12,fontWeight:600,color:T.text}}>{p.nombre}</div><div style={{fontSize:11,color:T.rose,fontWeight:700}}>Solo {p.stock} ud{p.stock!==1?"s":""} · {p.talla}</div></div></div>)}
          </div>
        </div>
      )}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// APP ROOT
// ═══════════════════════════════════════════════════════════════════════════════
export default function App(){
  const [sesion,setSesion]=useState(null);
  const [tab,setTab]=useState("caja");
  const [stock,setStock]=useState(STOCK0);
  const [ventas,setVentas]=useState(VENTAS0);
  const [toast,setToastS]=useState(null);
  const [modalKey,setModalKey]=useState(false);

  const setToast=p=>{setToastS(typeof p==="string"?{msg:p,type:"success"}:p);setTimeout(()=>setToastS(null),4500);};
  const onVenta=(v,cart)=>{
    setVentas(p=>[...p,v]);
    setStock(p=>p.map(s=>{const it=cart.find(c=>c.id===s.id);if(!it)return s;return{...s,stock:Math.max(0,s.stock-it.qty),vendidas:s.vendidas+it.qty};}));
  };

  if(!sesion)return<><style>{CSS}</style><LoginView onLogin={setSesion}/></>;

  const esCons=sesion.rol==="consultor";
  const vHoy=ventas.filter(v=>v.fecha===hoyLocal()).length;
  const tieneKey=!!getApiKey();
  const TABS=[{id:"caja",l:"🛒 Caja"},{id:"stock",l:"👗 Stock"},{id:"ventas",l:"📋 Ventas"},{id:"dashboard",l:"📊 Análisis"}];

  return(
    <>
      <style>{CSS}</style>
      {toast&&<Toast msg={toast.msg} type={toast.type} onClose={()=>setToastS(null)}/>}
      {modalKey&&<ModalApiKey onSave={()=>{setModalKey(false);setToast({msg:"Clave guardada. ¡Ya puedes escanear!"});}} onClose={()=>setModalKey(false)}/>}
      <div style={{minHeight:"100dvh",display:"flex",flexDirection:"column"}}>
        <div style={{background:T.surface,borderBottom:`1.5px solid ${T.border}`,padding:"0 14px",position:"sticky",top:0,zIndex:100,boxShadow:`0 2px 12px ${T.shadow}`}}>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",maxWidth:1280,margin:"0 auto",height:56,gap:8}}>
            <Logo/>
            <div style={{display:"flex",gap:2,overflowX:"auto"}}>{TABS.map(t=><button key={t.id} className={`tab-btn ${tab===t.id?"active":""}`} onClick={()=>setTab(t.id)}>{t.l}</button>)}</div>
            <div style={{display:"flex",alignItems:"center",gap:8,flexShrink:0}}>
              <div className="hide-m" style={{fontSize:11,color:vHoy>0?T.sage:T.textSoft,textAlign:"right",lineHeight:1.4}}><div style={{fontWeight:600}}>{vHoy} venta{vHoy!==1?"s":""} hoy</div><div style={{opacity:.7}}>{sesion.nombre}</div></div>
              {esCons&&<Badge color="indigo">CONSULTOR</Badge>}
              <button title={tieneKey?"Cambiar clave API":"Configurar clave API"} onClick={()=>setModalKey(true)}
                style={{background:tieneKey?T.sageLight:T.goldLight,border:`1.5px solid ${tieneKey?T.sage:T.gold}40`,borderRadius:8,padding:"5px 9px",cursor:"pointer",fontSize:14,lineHeight:1}}>
                {tieneKey?"🔑":"⚠️"}
              </button>
              <button className="btn-ghost" style={{fontSize:12,padding:"6px 10px"}} onClick={()=>setSesion(null)}>Salir</button>
            </div>
          </div>
        </div>
        <div style={{flex:1,padding:"14px",maxWidth:1280,width:"100%",margin:"0 auto"}}>
          {tab==="caja"&&<CajaView stock={stock} onVenta={onVenta}/>}
          {tab==="stock"&&<StockView stock={stock} onUpdate={setStock} setToast={setToast} onNeedKey={()=>setModalKey(true)}/>}
          {tab==="ventas"&&<VentasView ventas={ventas}/>}
          {tab==="dashboard"&&<DashboardView stock={stock} ventas={ventas} esConsultor={esCons} onNeedKey={()=>setModalKey(true)}/>}
        </div>
        <div style={{borderTop:`1px solid ${T.border}`,padding:"8px 14px",display:"flex",justifyContent:"space-between",background:T.surface,flexWrap:"wrap",gap:4}}>
          <span style={{fontSize:10,color:T.textSoft}}>Modas Capricho's · Sistema de gestión</span>
          <span style={{fontSize:10,color:T.textSoft,letterSpacing:".08em"}}>powered by ORACLYA</span>
        </div>
      </div>
    </>
  );
}
