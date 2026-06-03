import { useState, useMemo, useEffect, useCallback } from "react";

/* ══════════════════════════════════════════════════════════
   CONSTANTES FIBA
═══════════════════════════════════════════════════════════ */
const MSAL_CLIENT_ID   = "48c17191-0f37-422c-8c54-dcdfe41142e2";
const MSAL_REDIRECT    = "https://basketball-swart-nine.vercel.app/";
const ONEDRIVE_CURRENT = "basketball_arbitral_data.json";
const MSAL_SCOPES      = ["Files.ReadWrite","User.Read"];
const MSAL_AUTH_URL    = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";

const makeFileName = (config) => {
  const eq1   = (config.equipo1||"EquipoA").replace(/\s+/g,"");
  const eq2   = (config.equipo2||"EquipoB").replace(/\s+/g,"");
  const fecha = (config.fecha||"").replace(/-/g,"");
  return `Evaluacion_Arbitral_${eq1}_vs_${eq2}_${fecha}.json`;
};

const CATEGORIAS = {
  falta:    { label:"Falta",            color:"#e91e63", icon:"🚨" },
  violacion:{ label:"Violación",        color:"#ff9100", icon:"⚠️" },
  salida:   { label:"Salida de Pelota", color:"#29b6f6", icon:"📤" },
};

const TIPOS_POR_CATEGORIA = {
  falta: [
    "Falta Personal Ofensiva","Falta Personal Defensiva","Falta en Tiro",
    "Doble Foul","Falta Técnica","Falta Antideportiva","Falta Descalificadora",
  ],
  violacion: [
    "Violación de Pasos","Doble Drible","Violación de 3 Segundos",
    "Violación de 5 Segundos","Violación de 8 Segundos","Violación de 24 Segundos",
    "Bola Devuelta","Interferencia Ofensiva","Interferencia Defensiva","Invasión en TL",
  ],
  salida: ["Fuera de Banda"],
};

/* Veredictos según categoría */
const VEREDICTOS_FALTA = [
  { value:"acertado",    label:"Acertado",            color:"#00e676", bg:"rgba(0,230,118,0.12)",  positivo:true  },
  { value:"aceptable",   label:"Aceptable / Marginal", color:"#ffd740", bg:"rgba(255,215,64,0.12)", positivo:true  },
  { value:"marginal",    label:"Contacto Marginal",   color:"#ff9100", bg:"rgba(255,145,0,0.12)",  positivo:false },
  { value:"fantasioso",  label:"Fantasioso",          color:"#ff1744", bg:"rgba(255,23,68,0.12)",  positivo:false },
  { value:"correcto_nc", label:"Correcto No Llamado", color:"#00bcd4", bg:"rgba(0,188,212,0.12)",  positivo:true  },
  { value:"nocall_error",label:"No Call (Error)",     color:"#aa00ff", bg:"rgba(170,0,255,0.12)",  positivo:false },
];
const VEREDICTOS_SIMPLE = [
  { value:"acertado",  label:"Acertado",   color:"#00e676", bg:"rgba(0,230,118,0.12)", positivo:true  },
  { value:"no_acertado",label:"No Acertado",color:"#ff1744", bg:"rgba(255,23,68,0.12)",positivo:false },
];
const ALL_VEREDICTOS = [
  ...VEREDICTOS_FALTA,
  { value:"no_acertado", label:"No Acertado", color:"#ff1744", bg:"rgba(255,23,68,0.12)", positivo:false },
];

const getVeredictos = (cat) => (cat === "falta" ? VEREDICTOS_FALTA : VEREDICTOS_SIMPLE);
const vdInfo = (v) => ALL_VEREDICTOS.find(x => x.value === v) || ALL_VEREDICTOS[0];

/* Rangos de efectividad acordados en reunión 28-may-2026 */
const SEMAFORO = [
  { min:95, max:100, label:"Excelente",  color:"#00e676" },
  { min:85, max:94,  label:"Muy Bueno",  color:"#69f0ae" },
  { min:75, max:84,  label:"Bueno",      color:"#ffd740" },
  { min:61, max:74,  label:"Regular",    color:"#ff9100" },
  { min:0,  max:60,  label:"Deficiente", color:"#ff1744" },
];

const PERIODOS   = ["1er Cuarto","2do Cuarto","3er Cuarto","4to Cuarto","Tiempo Extra"];
const ROLES      = ["Crew Chief","Umpire 1","Umpire 2"];
const FISICO_OPT = ["Excelente","Bueno","Regular","Deficiente"];
const POS_ARB    = ["Líder","Center","Seguidor"];
const ZONAS      = ["Zona Primaria","Zona Secundaria","Zona Terciaria"];



const makeId = () => Date.now().toString(36) + Math.random().toString(36).slice(2);
const pct    = (n,t) => t===0 ? "0%" : Math.round((n/t)*100)+"%";
const pctNum = (n,t) => t===0 ? 0 : Math.round((n/t)*100);
const getSemaforo = (ef) => SEMAFORO.find(s => ef>=s.min && ef<=s.max) || SEMAFORO[4];
const categoriaDe = (tipo) => {
  for (const [cat,tipos] of Object.entries(TIPOS_POR_CATEGORIA))
    if (tipos.includes(tipo)) return cat;
  return "falta";
};

const DEFAULT_CONFIG = {
  liga:"Liga Señal Colombia de Baloncesto",
  equipo1:"Caimanes del Llano", equipo2:"Paisas de Antioquia",
  fecha:"2026-04-28", notaGrupal:"",
  arbitros:[
    { id:"a1", nombre:"Carlos Gonzalez", rol:"Crew Chief", foto:"" },
    { id:"a2", nombre:"Laura Niño",      rol:"Umpire 1",   foto:"" },
    { id:"a3", nombre:"Noe Diaz",        rol:"Umpire 2",   foto:"" },
  ],
};

function makeForm(arbitros, equipo1) {
  return {
    periodo:"1er Cuarto", minuto:"",
    arbitro: arbitros[0]?.id || "",
    posArbitro:"Líder", zona:"Zona Primaria",
    categoria:"falta", tipo:"Falta Personal Defensiva",
    equipo:equipo1, aceptacion:"Aceptado",
    veredicto:"acertado", pitazo:"simple",
    companeros:[], irs:false, irsResultado:"", descripcion:"",
  };
}
function makeFisico(arbitros) {
  return Object.fromEntries(arbitros.map(a => [a.id,{ nivel:"Bueno", notas:"" }]));
}

/* ── MSAL ── */
function buildAuthUrl() {
  const p = new URLSearchParams({
    client_id:MSAL_CLIENT_ID, response_type:"token",
    redirect_uri:MSAL_REDIRECT, scope:MSAL_SCOPES.join(" "),
    response_mode:"fragment", prompt:"select_account",
  });
  return `${MSAL_AUTH_URL}?${p.toString()}`;
}
function parseTokenFromHash() {
  const p = new URLSearchParams(window.location.hash.substring(1));
  const token = p.get("access_token");
  if (token) {
    const expiry = Date.now() + parseInt(p.get("expires_in")||"3600")*1000;
    localStorage.setItem("msft_token", token);
    localStorage.setItem("msft_expiry", expiry.toString());
    window.history.replaceState({}, document.title, window.location.pathname);
    return token;
  }
  return null;
}
function getStoredToken() {
  const token  = localStorage.getItem("msft_token");
  const expiry = parseInt(localStorage.getItem("msft_expiry")||"0");
  return (token && Date.now() < expiry-60000) ? token : null;
}

/* ── OneDrive ── */
async function saveToOneDrive(token, data) {
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/me/drive/root:/${fileName}:/content`,
    { method:"PUT", headers:{ Authorization:`Bearer ${token}`, "Content-Type":"application/json" }, body:JSON.stringify(data,null,2) }
  );
  if (!res.ok) throw new Error(`OneDrive save error: ${res.status}`);
}
async function loadFromOneDrive(token) {
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/me/drive/root:/${fileName}:/content`,
    { headers:{ Authorization:`Bearer ${token}` } }
  );
  if (res.status===404) return null;
  if (!res.ok) throw new Error(`OneDrive load error: ${res.status}`);
  return res.json();
}
async function getUserInfo(token) {
  const res = await fetch("https://graph.microsoft.com/v1.0/me",{ headers:{ Authorization:`Bearer ${token}` } });
  return res.ok ? res.json() : null;
}

/* ══════════════════════════════════════════════════════════
   CSS
═══════════════════════════════════════════════════════════ */
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@400;600;700;800&family=Barlow:wght@300;400;500&display=swap');
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
body{background:#090d18}
::-webkit-scrollbar{width:5px}
::-webkit-scrollbar-track{background:#090d18}
::-webkit-scrollbar-thumb{background:#e91e63;border-radius:3px}
input,select,textarea{background:#0f1623!important;color:#e8eaf6!important;border:1px solid #1d2840!important;border-radius:7px!important;padding:9px 12px!important;font-family:'Barlow',sans-serif!important;font-size:14px!important;width:100%;outline:none;transition:border-color .2s}
input:focus,select:focus,textarea:focus{border-color:#e91e63!important}
select option{background:#0f1623}
.btn{cursor:pointer;border:none;border-radius:7px;font-family:'Barlow Condensed',sans-serif;font-weight:700;letter-spacing:1px;font-size:15px;padding:10px 22px;transition:all .18s}
.btn-red{background:linear-gradient(135deg,#e91e63,#c62828);color:#fff}
.btn-red:hover{filter:brightness(1.15);transform:translateY(-1px);box-shadow:0 4px 18px rgba(233,30,99,.4)}
.btn-blue{background:linear-gradient(135deg,#1565c0,#0d47a1);color:#fff}
.btn-blue:hover{filter:brightness(1.15);transform:translateY(-1px)}
.btn-ghost{background:#1d2840;color:#90caf9}
.btn-ghost:hover{background:#253352}
.btn-warn{background:#1d2840;color:#ffd740}
.btn-del{background:#1d2840;color:#ff5252}
.btn-sm{padding:6px 12px;font-size:13px}
.card{background:#0f1623;border-radius:13px;border:1px solid #1d2840}
.tab{cursor:pointer;padding:10px 24px;border-radius:7px 7px 0 0;font-family:'Barlow Condensed',sans-serif;font-weight:700;font-size:15px;letter-spacing:1px;border:none;transition:all .2s}
.tab-on{background:#e91e63;color:#fff}
.tab-off{background:#0f1623;color:#546e7a}
.tab-off:hover{color:#90caf9}
.sec{font-family:'Barlow Condensed',sans-serif;font-size:20px;font-weight:800;letter-spacing:2px;text-transform:uppercase;color:#e91e63;margin-bottom:14px;border-bottom:2px solid rgba(233,30,99,.25);padding-bottom:5px}
.lbl{font-family:'Barlow',sans-serif;font-size:11px;color:#546e7a;text-transform:uppercase;letter-spacing:.9px;margin-bottom:4px}
.badge{display:inline-block;padding:3px 10px;border-radius:20px;font-size:12px;font-weight:700;font-family:'Barlow Condensed',sans-serif;white-space:nowrap}
.row-item{display:flex;align-items:flex-start;gap:10px;padding:11px 14px;border-radius:9px;margin-bottom:6px;border:1px solid #1d2840;transition:background .15s}
.row-item:hover{background:#131d2e}
.prog-bg{background:#1d2840;border-radius:20px;height:7px;overflow:hidden}
.prog-fill{height:7px;border-radius:20px;transition:width .5s}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:14px}
.grid3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:14px}
@media(max-width:700px){.grid2{grid-template-columns:1fr}.grid3{grid-template-columns:1fr}}
.tgl{cursor:pointer;border:2px solid #1d2840;border-radius:8px;padding:8px 14px;font-family:'Barlow Condensed',sans-serif;font-weight:700;font-size:13px;background:#1d2840;color:#546e7a;transition:all .15s}
.tgl:hover{color:#90caf9;border-color:#90caf9}
.tgl-blue{border-color:#29b6f6!important;background:rgba(41,182,246,0.12)!important;color:#29b6f6!important}
.tgl-orange{border-color:#ff9100!important;background:rgba(255,145,0,0.12)!important;color:#ff9100!important}
.tgl-green{border-color:#00e676!important;background:rgba(0,230,118,0.12)!important;color:#00e676!important}
.tgl-red{border-color:#ff1744!important;background:rgba(255,23,68,0.12)!important;color:#ff1744!important}
.tgl-pink{border-color:#e91e63!important;background:rgba(233,30,99,0.12)!important;color:#e91e63!important}
.sync-dot{width:8px;height:8px;border-radius:50%;display:inline-block;margin-right:6px}
.sync-ok{background:#00e676}
.sync-spin{background:#ffd740;animation:pulse 1s infinite}
.sync-err{background:#ff5252}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.4}}
`;

/* ══════════════════════════════════════════════════════════
   GRÁFICO TORTA SVG
═══════════════════════════════════════════════════════════ */
function PieChart({ data }) {
  if (!data || data.length === 0) return <div style={{color:"#546e7a",textAlign:"center",padding:20,fontFamily:"Barlow,sans-serif",fontSize:13}}>Sin datos</div>;
  const total = data.reduce((s,d)=>s+d.value,0);
  if (total === 0) return <div style={{color:"#546e7a",textAlign:"center",padding:20,fontFamily:"Barlow,sans-serif",fontSize:13}}>Sin datos</div>;
  const cx=110,cy=110,r=90;
  let startAngle = -Math.PI/2;
  const slices = data.map(d => {
    const angle = (d.value/total)*2*Math.PI;
    const x1=cx+r*Math.cos(startAngle), y1=cy+r*Math.sin(startAngle);
    startAngle += angle;
    const x2=cx+r*Math.cos(startAngle), y2=cy+r*Math.sin(startAngle);
    const large = angle>Math.PI?1:0;
    return { ...d, path:`M${cx},${cy} L${x1},${y1} A${r},${r} 0 ${large},1 ${x2},${y2} Z`, pct:Math.round((d.value/total)*100) };
  });
  return (
    <div style={{display:"flex",alignItems:"center",gap:16,flexWrap:"wrap"}}>
      <svg width={220} height={220} style={{flexShrink:0}}>
        {slices.map((s,i)=><path key={i} d={s.path} fill={s.color} stroke="#090d18" strokeWidth={2}/>)}
        <circle cx={cx} cy={cy} r={40} fill="#090d18"/>
        <text x={cx} y={cy+5} textAnchor="middle" fill="#e8eaf6" fontSize={13} fontFamily="Barlow Condensed,sans-serif" fontWeight="700">{total}</text>
      </svg>
      <div style={{flex:1,display:"flex",flexDirection:"column",gap:7}}>
        {slices.map((s,i)=>(
          <div key={i} style={{display:"flex",alignItems:"center",gap:8}}>
            <div style={{width:12,height:12,borderRadius:3,background:s.color,flexShrink:0}}/>
            <div style={{flex:1,fontSize:13,color:"#e8eaf6",fontFamily:"Barlow,sans-serif"}}>{s.label}</div>
            <div style={{fontSize:13,fontWeight:700,color:s.color,fontFamily:"Barlow Condensed,sans-serif"}}>{s.value} <span style={{color:"#546e7a",fontWeight:400}}>({s.pct}%)</span></div>
          </div>
        ))}
      </div>
    </div>
  );
}

/* ══════════════════════════════════════════════════════════
   GRÁFICO BARRAS SVG
═══════════════════════════════════════════════════════════ */
function BarChart({ data }) {
  if (!data || data.length === 0) return <div style={{color:"#546e7a",textAlign:"center",padding:20,fontFamily:"Barlow,sans-serif",fontSize:13}}>Sin datos</div>;
  const max = Math.max(...data.map(d=>d.value),1);
  const W=480,H=180,pad=40,barW=Math.min(50,Math.floor((W-pad*2)/data.length)-8);
  return (
    <svg width="100%" viewBox={`0 0 ${W} ${H+60}`} style={{overflow:"visible"}}>
      {[0,0.25,0.5,0.75,1].map(f=>{
        const y=pad+(1-f)*(H-pad);
        return <g key={f}><line x1={pad} y1={y} x2={W-10} y2={y} stroke="#1d2840" strokeWidth={1}/><text x={pad-5} y={y+4} textAnchor="end" fill="#546e7a" fontSize={10} fontFamily="Barlow,sans-serif">{Math.round(max*f)}</text></g>;
      })}
      {data.map((d,i)=>{
        const x=pad+i*((W-pad*2)/data.length)+(W-pad*2)/data.length/2-barW/2;
        const bh=((d.value/max)*(H-pad));
        const y=pad+(H-pad)-bh;
        return (
          <g key={i}>
            <rect x={x} y={y} width={barW} height={bh} rx={4} fill={d.color} opacity={0.85}/>
            <text x={x+barW/2} y={y-5} textAnchor="middle" fill={d.color} fontSize={12} fontWeight="700" fontFamily="Barlow Condensed,sans-serif">{d.value}</text>
            <text x={x+barW/2} y={H+pad+14} textAnchor="middle" fill="#90caf9" fontSize={11} fontFamily="Barlow,sans-serif">{d.label.length>10?d.label.substring(0,10)+"…":d.label}</text>
          </g>
        );
      })}
    </svg>
  );
}

/* ══════════════════════════════════════════════════════════
   PALETA DE COLORES
═══════════════════════════════════════════════════════════ */
const PALETTE = ["#e91e63","#29b6f6","#00e676","#ffd740","#ff9100","#aa00ff","#00bcd4","#ff5252","#69f0ae","#f06292"];

/* ══════════════════════════════════════════════════════════
   PESTAÑA GRÁFICOS — FILTROS COMBINABLES
═══════════════════════════════════════════════════════════ */
function VistaGraficos({ llamados, config }) {
  const [tipoGrafico,   setTipoGrafico]   = useState("barras");
  const [filtroCat,     setFiltroCat]     = useState("todas");   // todas | falta | violacion | salida
  const [filtroArbitros,setFiltroArbitros]= useState([]);        // [] = todos
  const [eje,           setEje]           = useState("equipo");  // dimensión del eje X

  const EJES = [
    { value:"equipo",     label:"Equipo sancionado"  },
    { value:"posArbitro", label:"Posición del árbitro"},
    { value:"veredicto",  label:"Veredicto"           },
    { value:"aceptacion", label:"Aceptación"          },
    { value:"pitazo",     label:"Tipo de pitazo"      },
    { value:"periodo",    label:"Período"             },
    { value:"irs",        label:"Revisión IRS"        },
    { value:"arbitro",    label:"Árbitro"             },
  ];

  const arbNombre = id => config.arbitros.find(a=>a.id===id)?.nombre || id;
  const vdLabel   = v  => ALL_VEREDICTOS.find(x=>x.value===v)?.label || v;
  const vdColor   = v  => ALL_VEREDICTOS.find(x=>x.value===v)?.color || "#90caf9";

  const toggleArbitro = id => setFiltroArbitros(prev =>
    prev.includes(id) ? prev.filter(x=>x!==id) : [...prev, id]
  );

  /* Aplicar filtros */
  const llamadosFiltrados = useMemo(() => {
    let data = llamados;
    if (filtroCat !== "todas")
      data = data.filter(l => (l.categoria || categoriaDe(l.tipo)) === filtroCat);
    if (filtroArbitros.length > 0)
      data = data.filter(l => filtroArbitros.includes(l.arbitro));
    return data;
  }, [llamados, filtroCat, filtroArbitros]);

  /* Construir datos del gráfico según eje */
  const chartData = useMemo(() => {
    const counts = {};
    llamadosFiltrados.forEach(l => {
      let key = "";
      switch(eje) {
        case "equipo":     key = l.equipo; break;
        case "posArbitro": key = l.posArbitro; break;
        case "veredicto":  key = vdLabel(l.veredicto); break;
        case "aceptacion": key = l.aceptacion; break;
        case "pitazo":     key = l.pitazo==="simple"?"Simple":l.pitazo==="doble"?"Doble":"Triple"; break;
        case "periodo":    key = l.periodo; break;
        case "irs":        key = l.irs?(l.irsResultado==="sostenida"?"IRS Sostenida":"IRS Cambiada"):"Sin IRS"; break;
        case "arbitro":    key = arbNombre(l.arbitro); break;
        default: key = "Otro";
      }
      counts[key] = (counts[key]||0)+1;
    });
    return Object.entries(counts).map(([label,value],i) => ({
      label, value,
      color: eje==="veredicto"
        ? (vdColor(llamadosFiltrados.find(l=>vdLabel(l.veredicto)===label)?.veredicto) || PALETTE[i%PALETTE.length])
        : eje==="equipo"
          ? (label===config.equipo1?"#29b6f6":label===config.equipo2?"#f06292":PALETTE[i%PALETTE.length])
          : eje==="posArbitro"
            ? (label==="Líder"?"#00e676":label==="Center"?"#ffd740":"#ff9100")
            : eje==="irs"
              ? (label==="Sin IRS"?"#546e7a":label==="IRS Sostenida"?"#00e676":"#ff5252")
              : eje==="aceptacion"
                ? (label==="Aceptado"?"#00e676":"#ff1744")
                : PALETTE[i%PALETTE.length],
    }));
  }, [llamadosFiltrados, eje, config]);

  /* Descripción del filtro activo */
  const descFiltro = () => {
    const cat = filtroCat==="todas" ? "Todos los llamados" : CATEGORIAS[filtroCat]?.label+"s";
    const arb = filtroArbitros.length===0 ? "toda la terna"
      : filtroArbitros.map(id=>arbNombre(id)).join(", ");
    return `${cat} · ${arb}`;
  };

  if (llamados.length === 0) return (
    <div className="card" style={{padding:40,textAlign:"center",color:"#546e7a",fontFamily:"Barlow,sans-serif"}}>
      Registra llamados primero para ver los gráficos.
    </div>
  );

  return (
    <div>
      <div className="sec">📈 Gráficos con Filtros Combinables</div>

      {/* Panel de filtros */}
      <div className="card" style={{padding:22,marginBottom:16}}>
        <div style={{display:"grid",gap:18}}>

          {/* Tipo de gráfico */}
          <div>
            <div className="lbl" style={{marginBottom:8}}>TIPO DE GRÁFICO</div>
            <div style={{display:"flex",gap:8}}>
              {[["barras","📊 Barras"],["torta","🥧 Torta"]].map(([v,lbl])=>(
                <button key={v} onClick={()=>setTipoGrafico(v)}
                  className={`tgl ${tipoGrafico===v?"tgl-blue":""}`}>{lbl}</button>
              ))}
            </div>
          </div>

          {/* Filtro categoría */}
          <div>
            <div className="lbl" style={{marginBottom:8}}>FILTRAR POR CATEGORÍA</div>
            <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
              <button onClick={()=>setFiltroCat("todas")}
                className={`tgl ${filtroCat==="todas"?"tgl-blue":""}`}>
                🏀 Todas
              </button>
              {Object.entries(CATEGORIAS).map(([key,cat])=>(
                <button key={key} onClick={()=>setFiltroCat(key)} style={{
                  cursor:"pointer",borderRadius:8,padding:"8px 16px",
                  fontFamily:"'Barlow Condensed',sans-serif",fontWeight:700,fontSize:13,
                  border:`2px solid ${filtroCat===key?cat.color:"transparent"}`,
                  background:filtroCat===key?`${cat.color}1a`:"#1d2840",
                  color:filtroCat===key?cat.color:"#546e7a",transition:"all .15s",
                }}>{cat.icon} {cat.label}</button>
              ))}
            </div>
          </div>

          {/* Filtro árbitros */}
          <div>
            <div className="lbl" style={{marginBottom:8}}>
              FILTRAR POR ÁRBITRO
              <span style={{color:"#546e7a",fontWeight:400,marginLeft:8}}>(vacío = toda la terna)</span>
            </div>
            <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
              {config.arbitros.map(a=>(
                <button key={a.id} onClick={()=>toggleArbitro(a.id)}
                  className={`tgl ${filtroArbitros.includes(a.id)?"tgl-pink":""}`}>
                  🧑‍⚖️ {a.nombre} <span style={{fontSize:11,opacity:.7}}>({a.rol})</span>
                </button>
              ))}
              {filtroArbitros.length>0&&(
                <button onClick={()=>setFiltroArbitros([])}
                  className="tgl" style={{color:"#ff5252",borderColor:"#ff5252"}}>
                  ✕ Limpiar
                </button>
              )}
            </div>
          </div>

          {/* Eje del gráfico */}
          <div>
            <div className="lbl" style={{marginBottom:8}}>VER DISTRIBUCIÓN POR</div>
            <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
              {EJES.map(e=>(
                <button key={e.value} onClick={()=>setEje(e.value)}
                  className={`tgl ${eje===e.value?"tgl-orange":""}`}
                  style={{fontSize:12,padding:"6px 14px"}}>
                  {e.label}
                </button>
              ))}
            </div>
          </div>
        </div>
      </div>

      {/* Gráfico */}
      <div className="card" style={{padding:24,marginBottom:16}}>
        <div style={{marginBottom:4}}>
          <div style={{fontSize:16,fontWeight:800,color:"#e91e63",letterSpacing:1,fontFamily:"'Barlow Condensed',sans-serif"}}>
            {tipoGrafico==="torta"?"🥧":"📊"} {EJES.find(e=>e.value===eje)?.label}
          </div>
          <div style={{fontSize:12,color:"#546e7a",fontFamily:"Barlow,sans-serif",marginTop:4}}>
            {descFiltro()} &nbsp;·&nbsp;
            <span style={{color:"#90caf9",fontWeight:700}}>{llamadosFiltrados.length}</span> llamados
          </div>
        </div>
        <div style={{marginTop:20}}>
          {llamadosFiltrados.length===0
            ? <div style={{textAlign:"center",color:"#546e7a",fontFamily:"Barlow,sans-serif",padding:30}}>
                Sin datos para los filtros seleccionados.
              </div>
            : tipoGrafico==="torta"
              ? <PieChart data={chartData}/>
              : <BarChart data={chartData}/>
          }
        </div>
      </div>

      {/* Tarjetas resumen */}
      {chartData.length>0&&(
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(130px,1fr))",gap:10}}>
          {chartData.map((d,i)=>(
            <div key={i} style={{background:"#0f1623",border:`1px solid ${d.color}33`,borderRadius:10,padding:"12px 14px",textAlign:"center"}}>
              <div style={{fontSize:11,color:d.color,fontWeight:700,fontFamily:"Barlow,sans-serif",marginBottom:4,wordBreak:"break-word"}}>{d.label}</div>
              <div style={{fontSize:26,fontWeight:800,color:d.color,lineHeight:1}}>{d.value}</div>
              <div style={{fontSize:11,color:"#546e7a",fontFamily:"Barlow,sans-serif"}}>{pct(d.value,llamadosFiltrados.length)}</div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

/* ══════════════════════════════════════════════════════════
   APP PRINCIPAL
═══════════════════════════════════════════════════════════ */
export default function App() {
  const [token,      setToken]     = useState(()=>getStoredToken());
  const [userInfo,   setUserInfo]  = useState(null);
  const [syncState,  setSyncState] = useState("idle");
  const [lastSync,   setLastSync]  = useState(null);
  const [dataLoaded,    setDataLoaded]   = useState(false);
  const [modalFinalizar,setModalFinalizar]= useState(false);
  const [modalNuevo,    setModalNuevo]    = useState(false);
  const [guardandoFinal,setGuardandoFinal]= useState(false);

  const [config,   setConfig]   = useState(DEFAULT_CONFIG);
  const [editCfg,  setEditCfg]  = useState(false);
  const [tmpCfg,   setTmpCfg]   = useState(DEFAULT_CONFIG);

  const [llamados, setLlamados] = useState([]);
  const [form,     setForm]     = useState(()=>makeForm(DEFAULT_CONFIG.arbitros,DEFAULT_CONFIG.equipo1));
  const [fisico,   setFisico]   = useState(()=>makeFisico(DEFAULT_CONFIG.arbitros));
  const [vista,    setVista]    = useState("registro");
  const [editId,   setEditId]   = useState(null);

  useEffect(()=>{ const t=parseTokenFromHash(); if(t) setToken(t); },[]);

  useEffect(()=>{
    if(!token) return;
    getUserInfo(token).then(u=>u&&setUserInfo(u));
    loadFromOneDrive(token).then(data=>{
      if(data){
        if(data.config)  { setConfig(data.config); setForm(makeForm(data.config.arbitros,data.config.equipo1)); setFisico(makeFisico(data.config.arbitros)); }
        if(data.llamados) setLlamados(data.llamados);
        if(data.fisico)   setFisico(data.fisico);
        setLastSync(new Date());
      }
      setDataLoaded(true); // habilitar guardado solo despues de cargar
    }).catch(()=>{ setDataLoaded(true); }); // si falla carga, igual habilitar
  },[token]);

  const syncData = useCallback(async(nl,nc,nf)=>{
    if(!token || !dataLoaded) return;
    setSyncState("saving");
    try {
      await saveToOneDrive(token,{config:nc,llamados:nl,fisico:nf},ONEDRIVE_CURRENT);
      setSyncState("ok"); setLastSync(new Date());
      setTimeout(()=>setSyncState("idle"),2000);
    } catch { setSyncState("error"); }
  },[token,dataLoaded]);

  const fv = (k,v) => setForm(f=>({...f,[k]:v}));
  const arbInfo = id => config.arbitros.find(a=>a.id===id);
  const companeroOpts = config.arbitros.filter(a=>a.id!==form.arbitro);
  const toggleCompanero = id => setForm(f=>({...f,companeros:f.companeros.includes(id)?f.companeros.filter(x=>x!==id):[...f.companeros,id]}));

  const saveCfg = () => {
    const nf={};
    tmpCfg.arbitros.forEach(a=>{nf[a.id]=fisico[a.id]||{nivel:"Bueno",notas:""};});
    setConfig(tmpCfg); setFisico(nf);
    setForm(f=>({...makeForm(tmpCfg.arbitros,tmpCfg.equipo1),periodo:f.periodo}));
    setEditCfg(false); syncData(llamados,tmpCfg,nf);
  };

  const submitLlamado = () => {
    if(!form.minuto){alert("Ingresa el minuto del llamado.");return;}
    let nl;
    if(editId){ nl=llamados.map(x=>x.id===editId?{...form,id:editId}:x); setEditId(null); }
    else { nl=[...llamados,{...form,id:makeId()}]; }
    setLlamados(nl);
    setForm(f=>({...makeForm(config.arbitros,config.equipo1),periodo:f.periodo,arbitro:f.arbitro}));
    syncData(nl,config,fisico);
  };

  const startEdit     = item=>{setForm({...item});setEditId(item.id);setVista("registro");window.scrollTo(0,0);};
  const cancelEdit    = ()=>{setForm(f=>({...makeForm(config.arbitros,config.equipo1),periodo:f.periodo,arbitro:f.arbitro}));setEditId(null);};
  const deleteLlamado = id=>{const nl=llamados.filter(x=>x.id!==id);setLlamados(nl);syncData(nl,config,fisico);};
  const updateFisico  = nf=>{setFisico(nf);syncData(llamados,config,nf);};

  /* Cuando cambia categoría, resetear veredicto */
  const cambiarCategoria = (cat) => {
    fv("categoria",cat);
    fv("tipo",TIPOS_POR_CATEGORIA[cat][0]);
    fv("veredicto", cat==="falta"?"acertado":"acertado");
  };

  /* ── Estadísticas ── */
  const stats = useMemo(()=>{
    const r={};
    config.arbitros.forEach(({id})=>{
      const m=llamados.filter(l=>l.arbitro===id);
      const positivos=m.filter(x=>ALL_VEREDICTOS.find(v=>v.value===x.veredicto)?.positivo).length;
      const ef=m.length?Math.round((positivos/m.length)*100):0;
      const solofalta=m.filter(x=>(x.categoria||categoriaDe(x.tipo))==="falta");
      r[id]={
        total:m.length, efectividad:ef, semaforo:getSemaforo(ef),
        acertado:    m.filter(x=>x.veredicto==="acertado").length,
        aceptable:   m.filter(x=>x.veredicto==="aceptable").length,
        marginal:    m.filter(x=>x.veredicto==="marginal").length,
        fantasioso:  m.filter(x=>x.veredicto==="fantasioso").length,
        correcto_nc: m.filter(x=>x.veredicto==="correcto_nc").length,
        nocall_error:m.filter(x=>x.veredicto==="nocall_error").length,
        no_acertado: m.filter(x=>x.veredicto==="no_acertado").length,
        localCount:  solofalta.filter(x=>x.equipo===config.equipo1).length,
        visitCount:  solofalta.filter(x=>x.equipo===config.equipo2).length,
        aceptados:   m.filter(x=>x.aceptacion==="Aceptado").length,
        noAceptados: m.filter(x=>x.aceptacion==="No Aceptado").length,
        porCategoria:{
          falta:    m.filter(x=>(x.categoria||categoriaDe(x.tipo))==="falta").length,
          violacion:m.filter(x=>(x.categoria||categoriaDe(x.tipo))==="violacion").length,
          salida:   m.filter(x=>(x.categoria||categoriaDe(x.tipo))==="salida").length,
        },
        pitazoDoble: m.filter(x=>x.pitazo==="doble").length,
        pitazoTriple:m.filter(x=>x.pitazo==="triple").length,
        irsCount:    m.filter(x=>x.irs).length,
        irsSostenida:m.filter(x=>x.irs&&x.irsResultado==="sostenida").length,
        irsCambiada: m.filter(x=>x.irs&&x.irsResultado==="cambiada").length,
        cuestionables:m.filter(x=>["fantasioso","marginal","nocall_error","no_acertado"].includes(x.veredicto)),
      };
    });
    return r;
  },[llamados,config]);

  /* Globales */
  const totalFaltas   = llamados.filter(x=>(x.categoria||categoriaDe(x.tipo))==="falta");
  const totalLocal    = totalFaltas.filter(x=>x.equipo===config.equipo1).length;
  const totalVisit    = totalFaltas.filter(x=>x.equipo===config.equipo2).length;
  const totalAcept    = llamados.filter(x=>x.aceptacion==="Aceptado").length;
  const totalNoAcept  = llamados.filter(x=>x.aceptacion==="No Aceptado").length;
  const totalDoble    = llamados.filter(x=>x.pitazo==="doble").length;
  const totalTriple   = llamados.filter(x=>x.pitazo==="triple").length;
  const totalIRS      = llamados.filter(x=>x.irs).length;
  const totalIRSSost  = llamados.filter(x=>x.irs&&x.irsResultado==="sostenida").length;
  const totalIRSCamb  = llamados.filter(x=>x.irs&&x.irsResultado==="cambiada").length;
  const positivosGlobal = llamados.filter(x=>ALL_VEREDICTOS.find(v=>v.value===x.veredicto)?.positivo).length;
  const efGlobal      = llamados.length?Math.round((positivosGlobal/llamados.length)*100):0;
  const globalSem     = getSemaforo(efGlobal);
  const veredictos    = getVeredictos(form.categoria);

  /* ── Exportar JSON local ── */
  const exportarJSON = () => {
    const data = { config, llamados, fisico };
    const blob = new Blob([JSON.stringify(data,null,2)], { type:"application/json" });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement("a");
    a.href = url;
    a.download = makeFileName(config);
    a.click();
    URL.revokeObjectURL(url);
  };

  /* ── Finalizar partido ── */
  const finalizarPartido = async () => {
    setGuardandoFinal(true);
    const data     = { config, llamados, fisico };
    const fileName = makeFileName(config);
    try {
      await saveToOneDrive(token, data, fileName);
      await saveToOneDrive(token, data, ONEDRIVE_CURRENT);
      setModalFinalizar(false);
      alert(`✅ Partido guardado como:\n${fileName}`);
    } catch(e) {
      alert("❌ Error al guardar en OneDrive. Usa Exportar JSON como respaldo.");
    }
    setGuardandoFinal(false);
  };

  /* ── Nuevo juego ── */
  const nuevoJuego = () => {
    const empty = {
      liga:"", equipo1:"", equipo2:"", fecha:"", notaGrupal:"",
      arbitros:[
        { id:makeId(), nombre:"", rol:"Crew Chief", foto:"" },
        { id:makeId(), nombre:"", rol:"Umpire 1",   foto:"" },
        { id:makeId(), nombre:"", rol:"Umpire 2",   foto:"" },
      ],
    };
    setConfig(empty);
    setLlamados([]);
    setFisico(makeFisico(empty.arbitros));
    setForm(makeForm(empty.arbitros,""));
    setEditId(null);
    setVista("registro");
    setModalNuevo(false);
    // Limpiar archivo actual en OneDrive
    saveToOneDrive(token, { config:empty, llamados:[], fisico:makeFisico(empty.arbitros) }, ONEDRIVE_CURRENT).catch(()=>{});
  };

  /* ── LOGIN ── */
  if(!token) return (
    <div style={{minHeight:"100vh",background:"#090d18",display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <style>{CSS}</style>
      <div style={{textAlign:"center",maxWidth:420}}>
        <div style={{fontSize:72,marginBottom:16}}>🏀</div>
        <div style={{fontSize:32,fontWeight:800,color:"#fff",letterSpacing:3,textTransform:"uppercase",fontFamily:"'Barlow Condensed',sans-serif"}}>Análisis Arbitral</div>
        <div style={{fontSize:14,color:"#546e7a",fontFamily:"Barlow,sans-serif",marginBottom:32,marginTop:8}}>Liga Señal Colombia de Baloncesto</div>
        <div className="card" style={{padding:32,marginBottom:20}}>
          <div style={{fontSize:16,fontWeight:700,color:"#90caf9",marginBottom:8,fontFamily:"'Barlow Condensed',sans-serif",letterSpacing:1}}>🔐 SINCRONIZACIÓN CON ONEDRIVE</div>
          <div style={{fontSize:13,color:"#546e7a",fontFamily:"Barlow,sans-serif",marginBottom:24,lineHeight:1.7}}>Tus datos se guardan automáticamente en tu OneDrive. Inicia sesión para continuar.</div>
          <button className="btn btn-blue" style={{width:"100%",fontSize:16,padding:"14px 22px"}}
            onClick={()=>window.location.href=buildAuthUrl()}>
            Iniciar sesión con Microsoft
          </button>
        </div>
        <div style={{fontSize:12,color:"#37474f",fontFamily:"Barlow,sans-serif"}}>Los datos se guardan en tu OneDrive personal como archivo JSON.</div>
      </div>
    </div>
  );

  /* ── APP PRINCIPAL ── */
  return (
    <div style={{minHeight:"100vh",background:"#090d18",color:"#e8eaf6",fontFamily:"'Barlow Condensed',sans-serif"}}>
      <style>{CSS}</style>

      {/* HEADER */}
      <div style={{background:"linear-gradient(135deg,#0b1020,#180826)",borderBottom:"3px solid #e91e63",padding:"22px 28px 18px"}}>
        <div style={{maxWidth:1100,margin:"0 auto"}}>
          <div style={{display:"flex",alignItems:"center",gap:14,flexWrap:"wrap"}}>
            <span style={{fontSize:46,lineHeight:1}}>🏀</span>
            <div style={{flex:1}}>
              <div style={{fontSize:26,fontWeight:800,letterSpacing:3,textTransform:"uppercase",color:"#fff"}}>Análisis Arbitral</div>
              <div style={{fontSize:13,color:"#546e7a",fontFamily:"Barlow,sans-serif"}}>
                {config.liga} &nbsp;|&nbsp;
                <span style={{color:"#90caf9"}}>{config.equipo1} vs {config.equipo2}</span>
                &nbsp;|&nbsp;{config.fecha}
              </div>
              <div style={{marginTop:5,display:"flex",gap:12,flexWrap:"wrap",alignItems:"center"}}>
                {config.arbitros.map(a=>(
                  <span key={a.id} style={{fontSize:13,color:"#b0bec5",fontFamily:"Barlow,sans-serif"}}>
                    <span style={{color:"#e91e63",fontWeight:700}}>{a.rol}:</span> {a.nombre}
                  </span>
                ))}
                <span style={{fontSize:12,color:"#546e7a",fontFamily:"Barlow,sans-serif",marginLeft:8}}>
                  <span className={`sync-dot ${syncState==="saving"?"sync-spin":syncState==="error"?"sync-err":"sync-ok"}`}/>
                  {syncState==="saving"?"Guardando...":syncState==="error"?"Error al guardar":lastSync?`Guardado ${lastSync.toLocaleTimeString()}`:"OneDrive conectado"}
                </span>
              </div>
            </div>
            <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
              {userInfo&&<span style={{fontSize:12,color:"#546e7a",fontFamily:"Barlow,sans-serif"}}>👤 {userInfo.displayName||userInfo.userPrincipalName}</span>}
              <button className="btn btn-ghost" style={{fontSize:13,padding:"7px 16px"}} onClick={()=>{setTmpCfg(config);setEditCfg(true);}}>⚙️ Configurar</button>
              <button className="btn btn-green btn-sm" onClick={()=>setModalFinalizar(true)}>🏁 Finalizar Partido</button>
              <button className="btn btn-warn btn-sm" onClick={exportarJSON}>💾 Exportar JSON</button>
              <button className="btn btn-ghost btn-sm" onClick={()=>setModalNuevo(true)}>🆕 Nuevo Juego</button>
              <button className="btn btn-del btn-sm" onClick={()=>{localStorage.removeItem("msft_token");localStorage.removeItem("msft_expiry");window.location.reload();}}>Cerrar sesión</button>
            </div>
          </div>
        </div>
      </div>

      {/* MODAL CONFIG */}
      {editCfg&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.85)",zIndex:200,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
          <div className="card" style={{width:"100%",maxWidth:640,maxHeight:"90vh",overflowY:"auto",padding:28}}>
            <div className="sec">⚙️ Configuración del Partido</div>
            <div style={{display:"grid",gap:14}}>
              <div><div className="lbl">Liga / Torneo</div><input value={tmpCfg.liga} onChange={e=>setTmpCfg(c=>({...c,liga:e.target.value}))}/></div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
                <div><div className="lbl">Equipo Local</div><input value={tmpCfg.equipo1} onChange={e=>setTmpCfg(c=>({...c,equipo1:e.target.value}))}/></div>
                <div><div className="lbl">Equipo Visitante</div><input value={tmpCfg.equipo2} onChange={e=>setTmpCfg(c=>({...c,equipo2:e.target.value}))}/></div>
              </div>
              <div><div className="lbl">Fecha</div><input type="date" value={tmpCfg.fecha} onChange={e=>setTmpCfg(c=>({...c,fecha:e.target.value}))}/></div>
              <div><div className="lbl">Nota Grupal — Trabajo en Equipo</div>
                <textarea rows={3} placeholder="Evaluación general de la terna como equipo..." value={tmpCfg.notaGrupal||""} onChange={e=>setTmpCfg(c=>({...c,notaGrupal:e.target.value}))} style={{resize:"vertical"}}/>
              </div>
              <div style={{borderTop:"1px solid #1d2840",paddingTop:14}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
                  <div style={{fontSize:15,fontWeight:700,color:"#e91e63"}}>TERNA ARBITRAL</div>
                  <button className="btn btn-ghost btn-sm" onClick={()=>setTmpCfg(c=>({...c,arbitros:[...c.arbitros,{id:makeId(),nombre:"",rol:"Crew Chief",foto:""}]}))}>+ Árbitro</button>
                </div>
                {tmpCfg.arbitros.map(a=>(
                  <div key={a.id} style={{display:"grid",gridTemplateColumns:"1fr 1fr auto",gap:8,marginBottom:8,alignItems:"center"}}>
                    <input placeholder="Nombre" value={a.nombre} onChange={e=>setTmpCfg(c=>({...c,arbitros:c.arbitros.map(x=>x.id===a.id?{...x,nombre:e.target.value}:x)}))}/>
                    <select value={a.rol} onChange={e=>setTmpCfg(c=>({...c,arbitros:c.arbitros.map(x=>x.id===a.id?{...x,rol:e.target.value}:x)}))}>
                      {ROLES.map(r=><option key={r}>{r}</option>)}
                    </select>
                    <button className="btn btn-del btn-sm" disabled={tmpCfg.arbitros.length<=1} onClick={()=>setTmpCfg(c=>({...c,arbitros:c.arbitros.filter(x=>x.id!==a.id)}))}>🗑</button>
                  </div>
                ))}
              </div>
            </div>
            <div style={{display:"flex",gap:10,marginTop:20}}>
              <button className="btn btn-red" onClick={saveCfg}>💾 GUARDAR</button>
              <button className="btn btn-ghost" onClick={()=>setEditCfg(false)}>Cancelar</button>
            </div>
          </div>
        </div>
      )}

      {/* TABS */}
      <div style={{background:"#090d18",padding:"0 28px",borderBottom:"1px solid #1d2840"}}>
        <div style={{maxWidth:1100,margin:"0 auto",display:"flex",gap:4,paddingTop:14,flexWrap:"wrap"}}>
          {[["registro","📋 Registrar"],["lista",`📑 Lista (${llamados.length})`],["reporte","📊 Reporte"],["graficos","📈 Gráficos"]].map(([v,lbl])=>(
            <button key={v} className={`tab ${vista===v?"tab-on":"tab-off"}`} onClick={()=>setVista(v)}>{lbl}</button>
          ))}
        </div>
      </div>

      <div style={{maxWidth:1100,margin:"0 auto",padding:"26px 28px"}}>

        {/* ═══ REGISTRO ═══ */}
        {vista==="registro"&&(
          <div>
            <div className="sec">{editId?"✏️ Editar Llamado":"➕ Registrar Llamado"}</div>
            <div style={{marginBottom:16}}>
              <div className="lbl" style={{marginBottom:8}}>PERÍODO ACTIVO</div>
              <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                {PERIODOS.map(p=><button key={p} onClick={()=>fv("periodo",p)} className={`tgl ${form.periodo===p?"tgl-blue":""}`}>{p}</button>)}
              </div>
            </div>
            <div className="card" style={{padding:22}}>
              {/* Categoría */}
              <div style={{marginBottom:18}}>
                <div className="lbl" style={{marginBottom:10}}>CATEGORÍA DEL LLAMADO</div>
                <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                  {Object.entries(CATEGORIAS).map(([key,cat])=>(
                    <button key={key} onClick={()=>cambiarCategoria(key)} style={{
                      cursor:"pointer",borderRadius:8,padding:"10px 18px",
                      fontFamily:"'Barlow Condensed',sans-serif",fontWeight:700,fontSize:14,
                      border:`2px solid ${form.categoria===key?cat.color:"transparent"}`,
                      background:form.categoria===key?`${cat.color}1a`:"#1d2840",
                      color:form.categoria===key?cat.color:"#546e7a",transition:"all .15s",
                    }}>{cat.icon} {cat.label}</button>
                  ))}
                </div>
              </div>

              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(185px,1fr))",gap:14}}>
                <div><div className="lbl">Minuto</div><input placeholder="Ej: 3:45" value={form.minuto} onChange={e=>fv("minuto",e.target.value)}/></div>
                <div><div className="lbl">Árbitro que Pita</div>
                  <select value={form.arbitro} onChange={e=>fv("arbitro",e.target.value)}>
                    {config.arbitros.map(a=><option key={a.id} value={a.id}>{a.nombre} ({a.rol})</option>)}
                  </select>
                </div>
                <div><div className="lbl">Posición del Árbitro</div>
                  <select value={form.posArbitro} onChange={e=>fv("posArbitro",e.target.value)}>
                    {POS_ARB.map(p=><option key={p}>{p}</option>)}
                  </select>
                </div>
                <div><div className="lbl">Zona del Llamado</div>
                  <select value={form.zona} onChange={e=>fv("zona",e.target.value)}>
                    {ZONAS.map(z=><option key={z}>{z}</option>)}
                  </select>
                </div>
                <div><div className="lbl">Tipo de Llamado</div>
                  <select value={form.tipo} onChange={e=>fv("tipo",e.target.value)}>
                    {TIPOS_POR_CATEGORIA[form.categoria].map(t=><option key={t}>{t}</option>)}
                  </select>
                </div>
                <div><div className="lbl">Equipo Sancionado</div>
                  <select value={form.equipo} onChange={e=>fv("equipo",e.target.value)}>
                    <option value={config.equipo1}>{config.equipo1}</option>
                    <option value={config.equipo2}>{config.equipo2}</option>
                    <option value="Ambos (Doble Foul)">Ambos (Doble Foul)</option>
                  </select>
                </div>
                <div><div className="lbl">Aceptación</div>
                  <div style={{display:"flex",gap:8,marginTop:4}}>
                    <button onClick={()=>fv("aceptacion","Aceptado")} className={`tgl ${form.aceptacion==="Aceptado"?"tgl-green":""}`} style={{flex:1}}>✔ Aceptado</button>
                    <button onClick={()=>fv("aceptacion","No Aceptado")} className={`tgl ${form.aceptacion==="No Aceptado"?"tgl-red":""}`} style={{flex:1}}>✘ No Aceptado</button>
                  </div>
                </div>

                {/* Pitazo */}
                <div style={{gridColumn:"1 / -1"}}>
                  <div className="lbl" style={{marginBottom:8}}>Tipo de Pitazo</div>
                  <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:10}}>
                    {[["simple","🔔 Simple"],["doble","🔔🔔 Doble"],["triple","🔔🔔🔔 Triple"]].map(([v,lbl])=>(
                      <button key={v} onClick={()=>fv("pitazo",v)} className={`tgl ${form.pitazo===v?"tgl-orange":""}`}>{lbl}</button>
                    ))}
                  </div>
                  {(form.pitazo==="doble"||form.pitazo==="triple")&&(
                    <div>
                      <div className="lbl" style={{marginBottom:7}}>{form.pitazo==="doble"?"Compañero que también pitó:":"Compañeros que también pitaron:"}</div>
                      <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                        {companeroOpts.map(a=>(
                          <button key={a.id} onClick={()=>toggleCompanero(a.id)} className={`tgl ${form.companeros.includes(a.id)?"tgl-orange":""}`}>
                            {a.nombre} ({a.rol})
                          </button>
                        ))}
                      </div>
                    </div>
                  )}
                </div>

                {/* IRS */}
                <div style={{gridColumn:"1 / -1"}}>
                  <div className="lbl" style={{marginBottom:8}}>REVISIÓN IRS</div>
                  <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
                    <button onClick={()=>fv("irs",!form.irs)} className={`tgl ${form.irs?"tgl-pink":""}`}>
                      {form.irs?"✅ IRS Revisado":"☐ Marcar revisión IRS"}
                    </button>
                    {form.irs&&<>
                      <button onClick={()=>fv("irsResultado","sostenida")} className={`tgl ${form.irsResultado==="sostenida"?"tgl-green":""}`}>✔ Decisión Sostenida</button>
                      <button onClick={()=>fv("irsResultado","cambiada")}  className={`tgl ${form.irsResultado==="cambiada"?"tgl-red":""}`}>↩ Decisión Cambiada</button>
                    </>}
                  </div>
                </div>

                {/* Veredicto — dinámico según categoría */}
                <div style={{gridColumn:"1 / -1"}}>
                  <div className="lbl" style={{marginBottom:8}}>
                    Veredicto del Llamado
                    {form.categoria!=="falta"&&<span style={{color:"#546e7a",fontWeight:400,fontSize:11,marginLeft:8}}>(Violaciones y Salidas: Acertado / No Acertado)</span>}
                  </div>
                  <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                    {veredictos.map(v=>(
                      <button key={v.value} onClick={()=>fv("veredicto",v.value)} style={{
                        cursor:"pointer",
                        border:`2px solid ${form.veredicto===v.value?v.color:"transparent"}`,
                        background:form.veredicto===v.value?v.bg:"#1d2840",
                        color:form.veredicto===v.value?v.color:"#546e7a",
                        borderRadius:8,padding:"8px 14px",
                        fontFamily:"'Barlow Condensed',sans-serif",fontWeight:700,fontSize:13,transition:"all .15s",
                      }}>{form.veredicto===v.value?"● ":""}{v.label}</button>
                    ))}
                  </div>
                  {form.categoria==="falta"&&(
                    <div style={{marginTop:8,fontSize:12,color:"#546e7a",fontFamily:"Barlow,sans-serif"}}>
                      <span style={{color:"#00e676"}}>Positivos:</span> Acertado · Aceptable · Correcto No Llamado &nbsp;|&nbsp;
                      <span style={{color:"#ff5252"}}>Negativos:</span> Marginal · Fantasioso · No Call (Error)
                    </div>
                  )}
                </div>

                <div style={{gridColumn:"1 / -1"}}>
                  <div className="lbl">Observación (opcional)</div>
                  <textarea rows={2} placeholder="Jugada, posición de jugadores, contexto..." value={form.descripcion} onChange={e=>fv("descripcion",e.target.value)} style={{resize:"vertical"}}/>
                </div>
              </div>

              <div style={{marginTop:18,display:"flex",gap:10,flexWrap:"wrap"}}>
                <button className="btn btn-red" onClick={submitLlamado}>{editId?"💾 GUARDAR CAMBIOS":"➕ AGREGAR LLAMADO"}</button>
                {editId&&<button className="btn btn-ghost" onClick={cancelEdit}>Cancelar</button>}
              </div>
            </div>

            {/* Estado físico */}
            <div style={{marginTop:26}}>
              <div className="sec">🏃 Estado Físico por Árbitro</div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(290px,1fr))",gap:14}}>
                {config.arbitros.map(a=>(
                  <div className="card" key={a.id} style={{padding:16}}>
                    <div style={{fontSize:16,fontWeight:700,color:"#90caf9",marginBottom:10}}>{a.nombre} <span style={{color:"#546e7a",fontSize:12}}>— {a.rol}</span></div>
                    <div className="lbl">Nivel Físico</div>
                    <select value={fisico[a.id]?.nivel||"Bueno"} onChange={e=>updateFisico({...fisico,[a.id]:{...fisico[a.id],nivel:e.target.value}})}>
                      {FISICO_OPT.map(f=><option key={f}>{f}</option>)}
                    </select>
                    <div className="lbl" style={{marginTop:10}}>Observaciones Físicas</div>
                    <textarea rows={2} placeholder="Movilidad, transiciones, cansancio..." value={fisico[a.id]?.notas||""} onChange={e=>updateFisico({...fisico,[a.id]:{...fisico[a.id],notas:e.target.value}})} style={{resize:"vertical"}}/>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

        {/* ═══ LISTA ═══ */}
        {vista==="lista"&&(
          <div>
            <div className="sec">📑 Todos los Llamados Registrados</div>
            {llamados.length===0
              ?<div className="card" style={{padding:40,textAlign:"center",color:"#546e7a",fontFamily:"Barlow,sans-serif"}}>Aún no hay llamados registrados.</div>
              :[...llamados].reverse().map(item=>{
                const a=arbInfo(item.arbitro);
                const vv=vdInfo(item.veredicto);
                const cat=CATEGORIAS[item.categoria||categoriaDe(item.tipo)];
                return(
                  <div key={item.id} className="row-item">
                    <div style={{minWidth:78,fontSize:12,color:"#546e7a",fontFamily:"Barlow,sans-serif",flexShrink:0}}>
                      {item.periodo}<br/><span style={{color:"#e91e63",fontWeight:700,fontSize:14}}>{item.minuto}</span>
                    </div>
                    <div style={{minWidth:130,flexShrink:0}}>
                      <div style={{fontSize:13,fontWeight:700,color:"#90caf9"}}>{a?.nombre||"?"}</div>
                      <div style={{fontSize:11,color:"#546e7a"}}>{a?.rol} · {item.posArbitro}</div>
                      <div style={{fontSize:11,color:"#ff9100"}}>{item.pitazo==="doble"?"🔔🔔":item.pitazo==="triple"?"🔔🔔🔔":"🔔"}</div>
                      {item.irs&&<div style={{fontSize:11,color:item.irsResultado==="cambiada"?"#ff5252":"#00e676"}}>IRS: {item.irsResultado==="cambiada"?"↩ Cambiada":"✔ Sostenida"}</div>}
                    </div>
                    <div style={{flex:1,minWidth:0}}>
                      <span style={{fontSize:11,background:`${cat?.color}22`,color:cat?.color,borderRadius:4,padding:"2px 7px",fontFamily:"Barlow,sans-serif",marginRight:6}}>{cat?.icon} {cat?.label}</span>
                      <div style={{fontSize:14,fontWeight:600,marginTop:4}}>{item.tipo}</div>
                      <div style={{fontSize:12,color:"#546e7a",fontFamily:"Barlow,sans-serif"}}>{item.equipo}</div>
                      {item.descripcion&&<div style={{fontSize:12,color:"#90a4ae",fontFamily:"Barlow,sans-serif",marginTop:2}}>{item.descripcion}</div>}
                    </div>
                    <div style={{display:"flex",flexDirection:"column",gap:5,flexShrink:0,alignItems:"flex-end"}}>
                      <span className="badge" style={{background:vv.bg,color:vv.color,border:`1px solid ${vv.color}`}}>{vv.label}</span>
                      <span className="badge" style={{background:item.aceptacion==="Aceptado"?"rgba(0,230,118,0.1)":"rgba(255,23,68,0.1)",color:item.aceptacion==="Aceptado"?"#00e676":"#ff1744",border:`1px solid ${item.aceptacion==="Aceptado"?"#00e676":"#ff1744"}`}}>
                        {item.aceptacion==="Aceptado"?"✔ Aceptado":"✘ No Aceptado"}
                      </span>
                    </div>
                    <div style={{display:"flex",gap:5,flexShrink:0}}>
                      <button className="btn btn-warn btn-sm" onClick={()=>startEdit(item)}>✏️</button>
                      <button className="btn btn-del  btn-sm" onClick={()=>deleteLlamado(item.id)}>🗑</button>
                    </div>
                  </div>
                );
              })
            }
          </div>
        )}

        {/* ═══ REPORTE ═══ */}
        {vista==="reporte"&&(
          <div>
            <div className="sec">📊 Reporte Final — Análisis Arbitral</div>
            {llamados.length===0
              ?<div className="card" style={{padding:40,textAlign:"center",color:"#546e7a",fontFamily:"Barlow,sans-serif"}}>Registra llamados primero para generar el reporte.</div>
              :(<>
                {/* Nota grupal */}
                {config.notaGrupal&&(
                  <div className="card" style={{padding:22,marginBottom:18,borderLeft:"4px solid #e91e63"}}>
                    <div style={{fontSize:15,fontWeight:800,color:"#e91e63",marginBottom:8,letterSpacing:1}}>🤝 TRABAJO EN EQUIPO — NOTA GRUPAL</div>
                    <div style={{fontSize:14,color:"#b0bec5",fontFamily:"Barlow,sans-serif",lineHeight:1.7}}>{config.notaGrupal}</div>
                  </div>
                )}

                {/* Resumen global */}
                <div className="card" style={{padding:22,marginBottom:22}}>
                  <div style={{fontSize:16,fontWeight:700,color:"#e91e63",marginBottom:16,letterSpacing:1}}>RESUMEN GLOBAL DEL PARTIDO</div>

                  {/* Semáforo global */}
                  <div style={{display:"flex",alignItems:"center",gap:16,marginBottom:20,background:"#090d18",borderRadius:10,padding:"14px 18px",border:`1px solid ${globalSem.color}33`}}>
                    <div style={{fontSize:42,fontWeight:800,color:globalSem.color,lineHeight:1}}>{efGlobal}%</div>
                    <div>
                      <div style={{fontSize:22,fontWeight:800,color:globalSem.color,letterSpacing:2}}>{globalSem.label.toUpperCase()}</div>
                      <div style={{fontSize:12,color:"#546e7a",fontFamily:"Barlow,sans-serif"}}>Efectividad global de la terna</div>
                    </div>
                    <div style={{flex:1}}>
                      <div className="prog-bg"><div className="prog-fill" style={{width:`${efGlobal}%`,background:globalSem.color}}/></div>
                      <div style={{display:"flex",gap:8,marginTop:8,flexWrap:"wrap"}}>
                        {SEMAFORO.map(s=><span key={s.label} style={{fontSize:10,color:s.color,fontFamily:"Barlow,sans-serif",opacity:.7}}>{s.label} {s.min}–{s.max}%</span>)}
                      </div>
                    </div>
                  </div>

                  {/* Veredictos globales */}
                  <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(95px,1fr))",gap:10,marginBottom:20}}>
                    {[
                      {label:"Total",        val:llamados.length,                                               color:"#90caf9"},
                      {label:"Acertados",    val:llamados.filter(x=>x.veredicto==="acertado").length,           color:"#00e676"},
                      {label:"Aceptables",   val:llamados.filter(x=>x.veredicto==="aceptable").length,          color:"#ffd740"},
                      {label:"Correcto NC",  val:llamados.filter(x=>x.veredicto==="correcto_nc").length,        color:"#00bcd4"},
                      {label:"Marginales",   val:llamados.filter(x=>x.veredicto==="marginal").length,           color:"#ff9100"},
                      {label:"Fantasiosos",  val:llamados.filter(x=>x.veredicto==="fantasioso").length,         color:"#ff1744"},
                      {label:"No Call Err",  val:llamados.filter(x=>x.veredicto==="nocall_error").length,       color:"#aa00ff"},
                      {label:"No Acertados", val:llamados.filter(x=>x.veredicto==="no_acertado").length,        color:"#ff5252"},
                    ].map(s=>(
                      <div key={s.label} style={{background:"#090d18",border:`1px solid ${s.color}22`,borderRadius:10,padding:"12px 10px",textAlign:"center"}}>
                        <div className="lbl">{s.label}</div>
                        <div style={{fontSize:24,fontWeight:800,color:s.color,lineHeight:1}}>{s.val}</div>
                      </div>
                    ))}
                  </div>

                  {/* Pitazos dobles y triples */}
                  <div style={{marginBottom:20}}>
                    <div className="lbl" style={{marginBottom:10}}>PITAZOS DOBLES Y TRIPLES</div>
                    <div className="grid3">
                      {[
                        {label:"Simples", val:llamados.filter(x=>x.pitazo==="simple").length,  color:"#90caf9"},
                        {label:"Dobles",  val:totalDoble,  color:"#ff9100"},
                        {label:"Triples", val:totalTriple, color:"#e91e63"},
                      ].map(e=>(
                        <div key={e.label} style={{background:"#090d18",border:`1px solid ${e.color}33`,borderRadius:10,padding:"14px 16px"}}>
                          <div style={{fontSize:13,color:e.color,fontWeight:700}}>{e.label}</div>
                          <div style={{fontSize:30,fontWeight:800,color:e.color,lineHeight:1.2}}>{e.val}</div>
                          <div style={{fontSize:11,color:"#546e7a",fontFamily:"Barlow,sans-serif"}}>{pct(e.val,llamados.length)} del total</div>
                        </div>
                      ))}
                    </div>
                  </div>

                  {/* Desglose por categoría */}
                  <div style={{marginBottom:20}}>
                    <div className="lbl" style={{marginBottom:10}}>DESGLOSE POR CATEGORÍA</div>
                    <div className="grid3">
                      {Object.entries(CATEGORIAS).map(([key,cat])=>{
                        const count=llamados.filter(x=>(x.categoria||categoriaDe(x.tipo))===key).length;
                        return(
                          <div key={key} style={{background:"#090d18",border:`1px solid ${cat.color}33`,borderRadius:10,padding:"14px 16px"}}>
                            <div style={{fontSize:13,color:cat.color,fontWeight:700}}>{cat.icon} {cat.label}</div>
                            <div style={{fontSize:30,fontWeight:800,color:cat.color,lineHeight:1.2}}>{count}</div>
                            <div style={{fontSize:11,color:"#546e7a",fontFamily:"Barlow,sans-serif"}}>{pct(count,llamados.length)} del total</div>
                            <div className="prog-bg" style={{marginTop:6}}><div className="prog-fill" style={{width:pct(count,llamados.length),background:cat.color}}/></div>
                          </div>
                        );
                      })}
                    </div>
                  </div>

                  {/* Sanciones por equipo — solo faltas */}
                  <div style={{marginBottom:20}}>
                    <div className="lbl" style={{marginBottom:6}}>SANCIONES POR EQUIPO <span style={{color:"#e91e63",fontWeight:700}}>(solo faltas)</span></div>
                    <div style={{fontSize:11,color:"#546e7a",fontFamily:"Barlow,sans-serif",marginBottom:10}}>
                      Total faltas: <span style={{color:"#e91e63",fontWeight:700}}>{totalFaltas.length}</span> &nbsp;|&nbsp; {config.equipo1}: {totalLocal} · {config.equipo2}: {totalVisit}
                    </div>
                    <div className="grid2">
                      {[
                        {label:config.equipo1,val:totalLocal,color:"#29b6f6"},
                        {label:config.equipo2,val:totalVisit,color:"#f06292"},
                      ].map(e=>(
                        <div key={e.label} style={{background:"#090d18",border:`1px solid ${e.color}33`,borderRadius:10,padding:"16px 18px"}}>
                          <div style={{fontSize:13,color:e.color,fontWeight:700,marginBottom:4}}>{e.label}</div>
                          <div style={{fontSize:34,fontWeight:800,color:e.color,lineHeight:1}}>{e.val}</div>
                          <div style={{fontSize:12,color:"#546e7a",fontFamily:"Barlow,sans-serif",marginTop:4}}>{pct(e.val,totalFaltas.length)} de las faltas</div>
                          <div className="prog-bg" style={{marginTop:8}}><div className="prog-fill" style={{width:pct(e.val,totalFaltas.length),background:e.color}}/></div>
                        </div>
                      ))}
                    </div>
                  </div>

                  {/* IRS Global */}
                  {totalIRS>0&&(
                    <div style={{marginBottom:20}}>
                      <div className="lbl" style={{marginBottom:10}}>REVISIONES IRS</div>
                      <div className="grid3">
                        {[
                          {label:"Total IRS",  val:totalIRS,     color:"#90caf9"},
                          {label:"Sostenidas", val:totalIRSSost, color:"#00e676"},
                          {label:"Cambiadas",  val:totalIRSCamb, color:"#ff5252"},
                        ].map(x=>(
                          <div key={x.label} style={{background:"#090d18",border:`1px solid ${x.color}33`,borderRadius:10,padding:"14px 16px"}}>
                            <div style={{fontSize:13,color:x.color,fontWeight:700}}>{x.label}</div>
                            <div style={{fontSize:30,fontWeight:800,color:x.color,lineHeight:1.2}}>{x.val}</div>
                            {x.label!=="Total IRS"&&<div style={{fontSize:11,color:"#546e7a",fontFamily:"Barlow,sans-serif"}}>{pct(x.val,totalIRS)} de las revisiones</div>}
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

                  {/* Aceptación global */}
                  <div>
                    <div className="lbl" style={{marginBottom:10}}>ACEPTACIÓN GLOBAL</div>
                    <div className="grid2">
                      {[
                        {label:"Aceptados",    val:totalAcept,   color:"#00e676"},
                        {label:"No Aceptados", val:totalNoAcept, color:"#ff1744"},
                      ].map(e=>(
                        <div key={e.label} style={{background:"#090d18",border:`1px solid ${e.color}33`,borderRadius:10,padding:"16px 18px"}}>
                          <div style={{fontSize:13,color:e.color,fontWeight:700,marginBottom:4}}>{e.label}</div>
                          <div style={{fontSize:34,fontWeight:800,color:e.color,lineHeight:1}}>{e.val}</div>
                          <div style={{fontSize:12,color:"#546e7a",fontFamily:"Barlow,sans-serif",marginTop:4}}>{pct(e.val,llamados.length)} del total</div>
                          <div className="prog-bg" style={{marginTop:8}}><div className="prog-fill" style={{width:pct(e.val,llamados.length),background:e.color}}/></div>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>

                {/* Detalle por árbitro */}
                {config.arbitros.map(a=>{
                  const s=stats[a.id];
                  const fis=fisico[a.id]||{nivel:"Bueno",notas:""};
                  const sem=s.semaforo;
                  return(
                    <div key={a.id} className="card" style={{padding:22,marginBottom:20}}>
                      {/* Cabecera */}
                      <div style={{display:"flex",alignItems:"center",gap:14,marginBottom:18,flexWrap:"wrap"}}>
                        <div style={{width:52,height:52,borderRadius:"50%",background:"linear-gradient(135deg,#e91e63,#880e4f)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:24,flexShrink:0}}>🧑‍⚖️</div>
                        <div style={{flex:1}}>
                          <div style={{fontSize:24,fontWeight:800,color:"#fff"}}>{a.nombre}</div>
                          <div style={{fontSize:12,color:"#546e7a",fontFamily:"Barlow,sans-serif"}}>{a.rol} &nbsp;|&nbsp; Físico: <span style={{color:"#90caf9"}}>{fis.nivel}</span></div>
                        </div>
                        <div style={{textAlign:"center",background:`${sem.color}15`,borderRadius:10,padding:"10px 18px",border:`1px solid ${sem.color}44`}}>
                          <div style={{fontSize:38,fontWeight:800,color:sem.color,lineHeight:1}}>{s.efectividad}%</div>
                          <div style={{fontSize:14,fontWeight:800,color:sem.color,letterSpacing:1}}>{sem.label.toUpperCase()}</div>
                          <div className="lbl" style={{marginTop:2}}>Efectividad</div>
                        </div>
                      </div>

                      {/* Barra efectividad */}
                      <div style={{marginBottom:16}}>
                        <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
                          <span style={{fontSize:11,color:"#546e7a",fontFamily:"Barlow,sans-serif"}}>EFECTIVIDAD</span>
                          <span style={{fontSize:11,color:sem.color,fontFamily:"Barlow,sans-serif",fontWeight:700}}>{s.efectividad}% — {sem.label}</span>
                        </div>
                        <div className="prog-bg"><div className="prog-fill" style={{width:`${s.efectividad}%`,background:sem.color}}/></div>
                      </div>

                      {/* Veredictos */}
                      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(85px,1fr))",gap:9,marginBottom:16}}>
                        {[
                          {label:"Total",       val:s.total,       color:"#90caf9"},
                          {label:"Acertados",   val:s.acertado,    color:"#00e676"},
                          {label:"Aceptables",  val:s.aceptable,   color:"#ffd740"},
                          {label:"Correcto NC", val:s.correcto_nc, color:"#00bcd4"},
                          {label:"Marginales",  val:s.marginal,    color:"#ff9100"},
                          {label:"Fantasiosos", val:s.fantasioso,  color:"#ff1744"},
                          {label:"No Call Err", val:s.nocall_error,color:"#aa00ff"},
                          {label:"No Acertado", val:s.no_acertado, color:"#ff5252"},
                        ].map(st=>(
                          <div key={st.label} style={{background:"#090d18",border:`1px solid ${st.color}33`,borderRadius:9,padding:"9px 8px",textAlign:"center"}}>
                            <div className="lbl">{st.label}</div>
                            <div style={{fontSize:20,fontWeight:800,color:st.color,lineHeight:1}}>{st.val}</div>
                          </div>
                        ))}
                      </div>

                      {/* Pitazos dobles/triples */}
                      {(s.pitazoDoble>0||s.pitazoTriple>0)&&(
                        <div style={{marginBottom:16}}>
                          <div className="lbl" style={{marginBottom:10}}>PITAZOS DOBLES Y TRIPLES</div>
                          <div className="grid2">
                            {[
                              {label:"Dobles",  val:s.pitazoDoble,  color:"#ff9100"},
                              {label:"Triples", val:s.pitazoTriple, color:"#e91e63"},
                            ].map(e=>(
                              <div key={e.label} style={{background:"#090d18",border:`1px solid ${e.color}33`,borderRadius:9,padding:"12px 14px"}}>
                                <div style={{fontSize:12,color:e.color,fontWeight:700}}>{e.label}</div>
                                <div style={{fontSize:26,fontWeight:800,color:e.color,lineHeight:1.2}}>{e.val}</div>
                                <div style={{fontSize:11,color:"#546e7a",fontFamily:"Barlow,sans-serif"}}>{pct(e.val,s.total)} del total</div>
                              </div>
                            ))}
                          </div>
                        </div>
                      )}

                      {/* Categorías */}
                      <div style={{marginBottom:16}}>
                        <div className="lbl" style={{marginBottom:10}}>LLAMADOS POR CATEGORÍA</div>
                        <div className="grid3">
                          {Object.entries(CATEGORIAS).map(([key,cat])=>(
                            <div key={key} style={{background:"#090d18",border:`1px solid ${cat.color}33`,borderRadius:9,padding:"12px 14px"}}>
                              <div style={{fontSize:12,color:cat.color,fontWeight:700}}>{cat.icon} {cat.label}</div>
                              <div style={{fontSize:26,fontWeight:800,color:cat.color,lineHeight:1.2}}>{s.porCategoria[key]}</div>
                              <div style={{fontSize:11,color:"#546e7a",fontFamily:"Barlow,sans-serif"}}>{pct(s.porCategoria[key],s.total)}</div>
                            </div>
                          ))}
                        </div>
                      </div>

                      {/* IRS por árbitro */}
                      {s.irsCount>0&&(
                        <div style={{marginBottom:16}}>
                          <div className="lbl" style={{marginBottom:10}}>REVISIONES IRS</div>
                          <div className="grid3">
                            {[
                              {label:"Total IRS",  val:s.irsCount,     color:"#90caf9"},
                              {label:"Sostenidas", val:s.irsSostenida, color:"#00e676"},
                              {label:"Cambiadas",  val:s.irsCambiada,  color:"#ff5252"},
                            ].map(x=>(
                              <div key={x.label} style={{background:"#090d18",border:`1px solid ${x.color}33`,borderRadius:9,padding:"12px 14px"}}>
                                <div style={{fontSize:12,color:x.color,fontWeight:700}}>{x.label}</div>
                                <div style={{fontSize:26,fontWeight:800,color:x.color,lineHeight:1.2}}>{x.val}</div>
                              </div>
                            ))}
                          </div>
                        </div>
                      )}

                      {/* Sanciones por equipo — solo faltas */}
                      <div style={{marginBottom:16}}>
                        <div className="lbl" style={{marginBottom:6}}>SANCIONES POR EQUIPO <span style={{color:"#e91e63",fontWeight:700}}>(solo faltas)</span></div>
                        <div className="grid2">
                          {[
                            {label:config.equipo1,val:s.localCount,color:"#29b6f6"},
                            {label:config.equipo2,val:s.visitCount,color:"#f06292"},
                          ].map(e=>(
                            <div key={e.label} style={{background:"#090d18",border:`1px solid ${e.color}33`,borderRadius:9,padding:"12px 14px"}}>
                              <div style={{fontSize:12,color:e.color,fontWeight:700}}>{e.label}</div>
                              <div style={{fontSize:28,fontWeight:800,color:e.color,lineHeight:1.1}}>{e.val}</div>
                              <div style={{fontSize:11,color:"#546e7a",fontFamily:"Barlow,sans-serif"}}>{pct(e.val,s.porCategoria.falta)}</div>
                              <div className="prog-bg" style={{marginTop:6}}><div className="prog-fill" style={{width:pct(e.val,s.porCategoria.falta),background:e.color}}/></div>
                            </div>
                          ))}
                        </div>
                      </div>

                      {/* Aceptación */}
                      <div style={{marginBottom:16}}>
                        <div className="lbl" style={{marginBottom:10}}>ACEPTACIÓN DE SUS LLAMADOS</div>
                        <div className="grid2">
                          {[
                            {label:"Aceptados",    val:s.aceptados,   color:"#00e676"},
                            {label:"No Aceptados", val:s.noAceptados, color:"#ff1744"},
                          ].map(e=>(
                            <div key={e.label} style={{background:"#090d18",border:`1px solid ${e.color}33`,borderRadius:9,padding:"12px 14px"}}>
                              <div style={{fontSize:12,color:e.color,fontWeight:700}}>{e.label}</div>
                              <div style={{fontSize:28,fontWeight:800,color:e.color,lineHeight:1.1}}>{e.val}</div>
                              <div style={{fontSize:11,color:"#546e7a",fontFamily:"Barlow,sans-serif"}}>{pct(e.val,s.total)}</div>
                              <div className="prog-bg" style={{marginTop:6}}><div className="prog-fill" style={{width:pct(e.val,s.total),background:e.color}}/></div>
                            </div>
                          ))}
                        </div>
                      </div>

                      {/* Cuestionables */}
                      {s.cuestionables.length>0&&(
                        <div style={{marginBottom:16}}>
                          <div className="lbl" style={{color:"#ff9100",marginBottom:8}}>⚠️ LLAMADOS CUESTIONABLES</div>
                          {s.cuestionables.map(item=>{
                            const vv=vdInfo(item.veredicto);
                            const cat=CATEGORIAS[item.categoria||categoriaDe(item.tipo)];
                            return(
                              <div key={item.id} style={{background:"#090d18",borderRadius:8,padding:"10px 14px",marginBottom:5,borderLeft:`3px solid ${vv.color}`}}>
                                <div style={{display:"flex",gap:10,alignItems:"center",flexWrap:"wrap"}}>
                                  <span style={{fontSize:13,color:"#e91e63",fontWeight:700}}>{item.periodo} · {item.minuto}</span>
                                  <span style={{fontSize:11,background:`${cat?.color}22`,color:cat?.color,borderRadius:4,padding:"2px 7px"}}>{cat?.icon} {cat?.label}</span>
                                  <span style={{fontSize:13,color:"#fff",fontWeight:600}}>{item.tipo}</span>
                                  <span className="badge" style={{background:vv.bg,color:vv.color,border:`1px solid ${vv.color}`}}>{vv.label}</span>
                                </div>
                                {item.descripcion&&<div style={{fontSize:12,color:"#78909c",fontFamily:"Barlow,sans-serif",marginTop:3}}>{item.descripcion}</div>}
                              </div>
                            );
                          })}
                        </div>
                      )}

                      {/* Notas físico */}
                      {fis.notas&&(
                        <div style={{background:"#090d18",borderRadius:8,padding:"12px 14px",marginBottom:16,borderLeft:"3px solid #29b6f6"}}>
                          <div style={{fontSize:12,color:"#29b6f6",fontWeight:700,marginBottom:3}}>🏃 ESTADO FÍSICO</div>
                          <div style={{fontSize:13,color:"#b0bec5",fontFamily:"Barlow,sans-serif"}}>{fis.notas}</div>
                        </div>
                      )}

                      {/* Fortalezas y oportunidades */}
                      <div className="grid2">
                        <div style={{background:"rgba(0,230,118,.06)",border:"1px solid rgba(0,230,118,.2)",borderRadius:10,padding:16}}>
                          <div style={{fontSize:14,fontWeight:800,color:"#00e676",marginBottom:10,letterSpacing:1}}>✅ FORTALEZAS</div>
                          <ul style={{paddingLeft:17,fontFamily:"Barlow,sans-serif",fontSize:13,color:"#b0bec5",lineHeight:1.9}}>
                            {s.acertado>0      &&<li>{s.acertado} llamado(s) bien acertado(s).</li>}
                            {s.correcto_nc>0   &&<li>{s.correcto_nc} correcto(s) no llamado(s).</li>}
                            {s.efectividad>=80 &&<li>Alta efectividad ({s.efectividad}% — {sem.label}).</li>}
                            {s.fantasioso===0  &&<li>Sin llamados fantasiosos.</li>}
                            {s.nocall_error===0&&<li>Sin No Calls por error.</li>}
                            {s.irsSostenida>0  &&<li>{s.irsSostenida} revisión/es IRS sostenida(s).</li>}
                            {s.aceptados>=s.noAceptados&&s.total>0&&<li>Mayor % aceptado ({pct(s.aceptados,s.total)}).</li>}
                            {s.total===0       &&<li>Sin datos registrados aún.</li>}
                          </ul>
                        </div>
                        <div style={{background:"rgba(255,23,68,.06)",border:"1px solid rgba(255,23,68,.2)",borderRadius:10,padding:16}}>
                          <div style={{fontSize:14,fontWeight:800,color:"#ff5252",marginBottom:10,letterSpacing:1}}>🎯 OPORTUNIDADES DE MEJORA</div>
                          <ul style={{paddingLeft:17,fontFamily:"Barlow,sans-serif",fontSize:13,color:"#b0bec5",lineHeight:1.9}}>
                            {s.fantasioso>0    &&<li>{s.fantasioso} llamado(s) fantasioso(s).</li>}
                            {s.marginal>1      &&<li>{s.marginal} marginales — afinar criterio.</li>}
                            {s.nocall_error>0  &&<li>{s.nocall_error} No Call(s) por error.</li>}
                            {s.irsCambiada>0   &&<li>{s.irsCambiada} decisión/es IRS cambiada(s).</li>}
                            {s.noAceptados>s.aceptados&&s.total>0&&<li>Alto % no aceptación ({pct(s.noAceptados,s.total)}).</li>}
                            {s.efectividad<75&&s.total>0&&<li>Efectividad menor al 75% ({sem.label}).</li>}
                            {s.total===0       &&<li>Sin datos registrados aún.</li>}
                          </ul>
                        </div>
                      </div>
                    </div>
                  );
                })}
              </>)
            }
          </div>
        )}

        {/* ═══ GRÁFICOS ═══ */}
        {vista==="graficos"&&<VistaGraficos llamados={llamados} config={config}/>}

      </div>

      {/* MODAL FINALIZAR PARTIDO */}
      {modalFinalizar&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.88)",zIndex:300,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
          <div className="card" style={{width:"100%",maxWidth:480,padding:32,textAlign:"center"}}>
            <div style={{fontSize:48,marginBottom:12}}>🏁</div>
            <div style={{fontSize:22,fontWeight:800,color:"#fff",letterSpacing:2,fontFamily:"'Barlow Condensed',sans-serif",marginBottom:8}}>FINALIZAR PARTIDO</div>
            <div style={{fontSize:13,color:"#546e7a",fontFamily:"Barlow,sans-serif",lineHeight:1.7,marginBottom:20}}>
              Se guardará el partido en OneDrive con el nombre:<br/>
              <span style={{color:"#00e676",fontWeight:700,fontFamily:"monospace",fontSize:12,wordBreak:"break-all"}}>{makeFileName(config)}</span>
            </div>
            <div style={{background:"#090d18",borderRadius:8,padding:"12px 16px",marginBottom:20,fontSize:13,color:"#b0bec5",fontFamily:"Barlow,sans-serif"}}>
              <div>📅 {config.fecha}</div>
              <div>🏀 {config.equipo1} vs {config.equipo2}</div>
              <div>📋 {llamados.length} llamados registrados</div>
            </div>
            <div style={{display:"flex",gap:10,justifyContent:"center",flexWrap:"wrap"}}>
              <button className="btn btn-green" onClick={finalizarPartido} disabled={guardandoFinal}>
                {guardandoFinal?"⏳ Guardando...":"✅ GUARDAR PARTIDO"}
              </button>
              <button className="btn btn-ghost" onClick={()=>setModalFinalizar(false)}>Cancelar</button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL NUEVO JUEGO */}
      {modalNuevo&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.88)",zIndex:300,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
          <div className="card" style={{width:"100%",maxWidth:480,padding:32,textAlign:"center"}}>
            <div style={{fontSize:48,marginBottom:12}}>🆕</div>
            <div style={{fontSize:22,fontWeight:800,color:"#fff",letterSpacing:2,fontFamily:"'Barlow Condensed',sans-serif",marginBottom:8}}>NUEVO JUEGO</div>
            <div style={{fontSize:13,color:"#ff9100",fontFamily:"Barlow,sans-serif",lineHeight:1.7,marginBottom:20}}>
              ⚠️ Esta acción borrará <strong>toda</strong> la información del partido actual:<br/>
              equipos, árbitros, llamados, estado físico y nota grupal.
            </div>
            <div style={{background:"#090d18",borderRadius:8,padding:"12px 16px",marginBottom:20,fontSize:13,color:"#b0bec5",fontFamily:"Barlow,sans-serif"}}>
              <div>🏀 {config.equipo1} vs {config.equipo2}</div>
              <div>📋 {llamados.length} llamados se perderán</div>
              <div style={{color:"#ff5252",marginTop:6}}>¿Finalizaste y guardaste el partido primero?</div>
            </div>
            <div style={{display:"flex",gap:10,justifyContent:"center",flexWrap:"wrap"}}>
              <button className="btn btn-red" onClick={nuevoJuego}>🗑 SÍ, INICIAR NUEVO JUEGO</button>
              <button className="btn btn-ghost" onClick={()=>setModalNuevo(false)}>Cancelar</button>
            </div>
          </div>
        </div>
      )}

      <div style={{borderTop:"1px solid #1d2840",padding:"14px 28px",textAlign:"center",color:"#263238",fontSize:11,fontFamily:"Barlow,sans-serif"}}>
        Herramienta de Análisis Arbitral · Liga Señal Colombia de Baloncesto · Uso interno
      </div>
    </div>
  );
}
