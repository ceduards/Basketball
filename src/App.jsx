import { useState, useMemo, useEffect, useCallback } from "react";

/* ══════════════════════════════════════════════════════════
   CONSTANTES FIBA
═══════════════════════════════════════════════════════════ */
const MSAL_CLIENT_ID   = "48c17191-0f37-422c-8c54-dcdfe41142e2";
const MSAL_REDIRECT    = "https://basketball-swart-nine.vercel.app/";
const ONEDRIVE_FILE    = "basketball_arbitral_data.json";
const MSAL_SCOPES      = ["Files.ReadWrite", "User.Read"];
const MSAL_AUTH_URL    = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize`;

const CATEGORIAS = {
  falta:    { label: "Falta",           color: "#e91e63", icon: "🚨" },
  violacion:{ label: "Violación",       color: "#ff9100", icon: "⚠️" },
  salida:   { label: "Salida de Pelota",color: "#29b6f6", icon: "📤" },
};

const TIPOS_POR_CATEGORIA = {
  falta: [
    "Falta Personal Ofensiva",
    "Falta Personal Defensiva",
    "Falta en Tiro",
    "Doble Foul",
    "Falta Técnica",
    "Falta Antideportiva",
    "Falta Descalificadora",
  ],
  violacion: [
    "Violación de Pasos",
    "Doble Drible",
    "Violación de 3 Segundos",
    "Violación de 5 Segundos",
    "Violación de 8 Segundos",
    "Violación de 24 Segundos",
  ],
  salida: [
    "Fuera de Banda",
  ],
};

const TODOS_TIPOS = Object.values(TIPOS_POR_CATEGORIA).flat();

const VEREDICTOS = [
  { value:"acertado",    label:"Acertado",              color:"#00e676", bg:"rgba(0,230,118,0.12)",  positivo: true  },
  { value:"aceptable",   label:"Aceptable / Marginal",  color:"#ffd740", bg:"rgba(255,215,64,0.12)", positivo: true  },
  { value:"marginal",    label:"Contacto Marginal",     color:"#ff9100", bg:"rgba(255,145,0,0.12)",  positivo: false },
  { value:"fantasioso",  label:"Fantasioso",            color:"#ff1744", bg:"rgba(255,23,68,0.12)",  positivo: false },
  { value:"correcto_nc", label:"Correcto No Llamado",   color:"#00bcd4", bg:"rgba(0,188,212,0.12)",  positivo: true  },
  { value:"nocall_error",label:"No Call (Error)",       color:"#aa00ff", bg:"rgba(170,0,255,0.12)",  positivo: false },
];

/* Rangos de efectividad acordados en reunión 28-may-2026 */
const SEMAFORO = [
  { min:95, max:100, label:"Excelente",  color:"#00e676", bg:"rgba(0,230,118,0.15)"  },
  { min:85, max:94,  label:"Muy Bueno", color:"#69f0ae", bg:"rgba(105,240,174,0.15)" },
  { min:75, max:84,  label:"Bueno",     color:"#ffd740", bg:"rgba(255,215,64,0.15)"  },
  { min:61, max:74,  label:"Regular",   color:"#ff9100", bg:"rgba(255,145,0,0.15)"   },
  { min:0,  max:60,  label:"Deficiente",color:"#ff1744", bg:"rgba(255,23,68,0.15)"   },
];

const PERIODOS   = ["1er Cuarto","2do Cuarto","3er Cuarto","4to Cuarto","Tiempo Extra"];
const ROLES      = ["Crew Chief","Umpire 1","Umpire 2"];
const FISICO_OPT = ["Excelente","Bueno","Regular","Deficiente"];
const POS_ARB    = ["Líder","Center","Seguidor"];
const ZONAS      = ["Zona Primaria","Zona Secundaria","Zona Terciaria"];

const makeId  = () => Date.now().toString(36) + Math.random().toString(36).slice(2);
const pct     = (n, t) => t === 0 ? "0%" : Math.round((n / t) * 100) + "%";
const getSemaforo = (ef) => SEMAFORO.find(s => ef >= s.min && ef <= s.max) || SEMAFORO[4];
const categoriaDe = (tipo) => {
  for (const [cat, tipos] of Object.entries(TIPOS_POR_CATEGORIA)) {
    if (tipos.includes(tipo)) return cat;
  }
  return "falta";
};

const DEFAULT_CONFIG = {
  liga:     "Liga Señal Colombia de Baloncesto",
  equipo1:  "Caimanes del Llano",
  equipo2:  "Paisas de Antioquia",
  fecha:    "2026-04-28",
  notaGrupal: "",
  arbitros: [
    { id:"a1", nombre:"Carlos Gonzalez", rol:"Crew Chief", foto:"" },
    { id:"a2", nombre:"Laura Niño",      rol:"Umpire 1",   foto:"" },
    { id:"a3", nombre:"Noe Diaz",        rol:"Umpire 2",   foto:"" },
  ],
};

function makeForm(arbitros, equipo1) {
  return {
    periodo:    "1er Cuarto",
    minuto:     "",
    arbitro:    arbitros[0]?.id || "",
    posArbitro: "Líder",
    zona:       "Zona Primaria",
    categoria:  "falta",
    tipo:       "Falta Personal Defensiva",
    equipo:     equipo1,
    aceptacion: "Aceptado",
    veredicto:  "acertado",
    pitazo:     "simple",
    companeros: [],
    irs:        false,
    irsResultado: "",
    descripcion:"",
  };
}

function makeFisico(arbitros) {
  return Object.fromEntries(arbitros.map(a => [a.id, { nivel:"Bueno", notas:"" }]));
}

/* ══════════════════════════════════════════════════════════
   MSAL — AUTENTICACIÓN MICROSOFT
═══════════════════════════════════════════════════════════ */
function buildAuthUrl() {
  const params = new URLSearchParams({
    client_id:     MSAL_CLIENT_ID,
    response_type: "token",
    redirect_uri:  MSAL_REDIRECT,
    scope:         MSAL_SCOPES.join(" "),
    response_mode: "fragment",
    prompt:        "select_account",
  });
  return `${MSAL_AUTH_URL}?${params.toString()}`;
}

function parseTokenFromHash() {
  const hash = window.location.hash.substring(1);
  const params = new URLSearchParams(hash);
  const token   = params.get("access_token");
  const expires = params.get("expires_in");
  if (token) {
    const expiry = Date.now() + parseInt(expires || "3600") * 1000;
    localStorage.setItem("msft_token",  token);
    localStorage.setItem("msft_expiry", expiry.toString());
    window.history.replaceState({}, document.title, window.location.pathname);
    return token;
  }
  return null;
}

function getStoredToken() {
  const token  = localStorage.getItem("msft_token");
  const expiry = parseInt(localStorage.getItem("msft_expiry") || "0");
  if (token && Date.now() < expiry - 60000) return token;
  return null;
}

/* ══════════════════════════════════════════════════════════
   ONEDRIVE API
═══════════════════════════════════════════════════════════ */
async function saveToOneDrive(token, data) {
  const json = JSON.stringify(data, null, 2);
  const res  = await fetch(
    `https://graph.microsoft.com/v1.0/me/drive/root:/${ONEDRIVE_FILE}:/content`,
    {
      method:  "PUT",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body:    json,
    }
  );
  if (!res.ok) throw new Error(`OneDrive save error: ${res.status}`);
  return true;
}

async function loadFromOneDrive(token) {
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/me/drive/root:/${ONEDRIVE_FILE}:/content`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (res.status === 404) return null;
  if (!res.ok) throw new Error(`OneDrive load error: ${res.status}`);
  return await res.json();
}

async function getUserInfo(token) {
  const res = await fetch("https://graph.microsoft.com/v1.0/me", {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) return null;
  return res.json();
}

/* ══════════════════════════════════════════════════════════
   CSS
═══════════════════════════════════════════════════════════ */
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@400;600;700;800&family=Barlow:wght@300;400;500&display=swap');
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
body { background: #090d18; }
::-webkit-scrollbar { width: 5px; }
::-webkit-scrollbar-track { background: #090d18; }
::-webkit-scrollbar-thumb { background: #e91e63; border-radius: 3px; }
input, select, textarea {
  background: #0f1623 !important; color: #e8eaf6 !important;
  border: 1px solid #1d2840 !important; border-radius: 7px !important;
  padding: 9px 12px !important; font-family: 'Barlow', sans-serif !important;
  font-size: 14px !important; width: 100%; outline: none; transition: border-color .2s;
}
input:focus, select:focus, textarea:focus { border-color: #e91e63 !important; }
select option { background: #0f1623; }
.btn { cursor: pointer; border: none; border-radius: 7px; font-family: 'Barlow Condensed', sans-serif; font-weight: 700; letter-spacing: 1px; font-size: 15px; padding: 10px 22px; transition: all .18s; }
.btn-red   { background: linear-gradient(135deg,#e91e63,#c62828); color: #fff; }
.btn-red:hover { filter: brightness(1.15); transform: translateY(-1px); box-shadow: 0 4px 18px rgba(233,30,99,.4); }
.btn-blue  { background: linear-gradient(135deg,#1565c0,#0d47a1); color: #fff; }
.btn-blue:hover { filter: brightness(1.15); transform: translateY(-1px); }
.btn-ghost { background: #1d2840; color: #90caf9; }
.btn-ghost:hover { background: #253352; }
.btn-warn  { background: #1d2840; color: #ffd740; }
.btn-del   { background: #1d2840; color: #ff5252; }
.btn-sm    { padding: 6px 12px; font-size: 13px; }
.btn-green { background: linear-gradient(135deg,#00897b,#004d40); color: #fff; }
.card { background: #0f1623; border-radius: 13px; border: 1px solid #1d2840; }
.tab { cursor: pointer; padding: 10px 24px; border-radius: 7px 7px 0 0; font-family: 'Barlow Condensed', sans-serif; font-weight: 700; font-size: 15px; letter-spacing: 1px; border: none; transition: all .2s; }
.tab-on  { background: #e91e63; color: #fff; }
.tab-off { background: #0f1623; color: #546e7a; }
.tab-off:hover { color: #90caf9; }
.sec { font-family: 'Barlow Condensed', sans-serif; font-size: 20px; font-weight: 800; letter-spacing: 2px; text-transform: uppercase; color: #e91e63; margin-bottom: 14px; border-bottom: 2px solid rgba(233,30,99,.25); padding-bottom: 5px; }
.lbl { font-family: 'Barlow', sans-serif; font-size: 11px; color: #546e7a; text-transform: uppercase; letter-spacing: .9px; margin-bottom: 4px; }
.badge { display: inline-block; padding: 3px 10px; border-radius: 20px; font-size: 12px; font-weight: 700; font-family: 'Barlow Condensed', sans-serif; white-space: nowrap; }
.row-item { display: flex; align-items: flex-start; gap: 10px; padding: 11px 14px; border-radius: 9px; margin-bottom: 6px; border: 1px solid #1d2840; transition: background .15s; }
.row-item:hover { background: #131d2e; }
.prog-bg   { background: #1d2840; border-radius: 20px; height: 7px; overflow: hidden; }
.prog-fill { height: 7px; border-radius: 20px; transition: width .5s; }
.grid2 { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }
.grid3 { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 14px; }
@media (max-width: 700px) { .grid2 { grid-template-columns: 1fr; } .grid3 { grid-template-columns: 1fr; } }
.tgl { cursor: pointer; border: 2px solid #1d2840; border-radius: 8px; padding: 8px 14px; font-family: 'Barlow Condensed', sans-serif; font-weight: 700; font-size: 13px; background: #1d2840; color: #546e7a; transition: all .15s; }
.tgl:hover { color: #90caf9; border-color: #90caf9; }
.tgl-blue   { border-color: #29b6f6 !important; background: rgba(41,182,246,0.12) !important; color: #29b6f6 !important; }
.tgl-orange { border-color: #ff9100 !important; background: rgba(255,145,0,0.12) !important; color: #ff9100 !important; }
.tgl-green  { border-color: #00e676 !important; background: rgba(0,230,118,0.12) !important; color: #00e676 !important; }
.tgl-red    { border-color: #ff1744 !important; background: rgba(255,23,68,0.12) !important; color: #ff1744 !important; }
.tgl-pink   { border-color: #e91e63 !important; background: rgba(233,30,99,0.12) !important; color: #e91e63 !important; }
.sync-dot { width:8px; height:8px; border-radius:50%; display:inline-block; margin-right:6px; }
.sync-ok   { background:#00e676; }
.sync-spin { background:#ffd740; animation: pulse 1s infinite; }
.sync-err  { background:#ff5252; }
@keyframes pulse { 0%,100%{opacity:1} 50%{opacity:.4} }
.cat-tab { cursor:pointer; padding:8px 16px; border-radius:6px; font-family:'Barlow Condensed',sans-serif; font-weight:700; font-size:13px; border:2px solid transparent; transition:all .15s; }
`;

/* ══════════════════════════════════════════════════════════
   APP
═══════════════════════════════════════════════════════════ */
export default function App() {
  const [token,    setToken]    = useState(() => getStoredToken());
  const [userInfo, setUserInfo] = useState(null);
  const [syncState,setSyncState]= useState("idle"); // idle | saving | ok | error
  const [lastSync, setLastSync] = useState(null);

  const [config,   setConfig]   = useState(DEFAULT_CONFIG);
  const [editCfg,  setEditCfg]  = useState(false);
  const [tmpCfg,   setTmpCfg]   = useState(DEFAULT_CONFIG);

  const [llamados, setLlamados] = useState([]);
  const [form,     setForm]     = useState(() => makeForm(DEFAULT_CONFIG.arbitros, DEFAULT_CONFIG.equipo1));
  const [fisico,   setFisico]   = useState(() => makeFisico(DEFAULT_CONFIG.arbitros));
  const [vista,    setVista]    = useState("registro");
  const [editId,   setEditId]   = useState(null);

  /* ── Token desde redirect hash ── */
  useEffect(() => {
    const t = parseTokenFromHash();
    if (t) { setToken(t); }
  }, []);

  /* ── Cargar datos de OneDrive al conectar ── */
  useEffect(() => {
    if (!token) return;
    getUserInfo(token).then(u => u && setUserInfo(u));
    loadFromOneDrive(token).then(data => {
      if (!data) return;
      if (data.config)   { setConfig(data.config);   setForm(makeForm(data.config.arbitros, data.config.equipo1)); setFisico(makeFisico(data.config.arbitros)); }
      if (data.llamados) setLlamados(data.llamados);
      if (data.fisico)   setFisico(data.fisico);
      setLastSync(new Date());
    }).catch(() => {});
  }, [token]);

  /* ── Auto-guardar en OneDrive ── */
  const syncData = useCallback(async (newLlamados, newConfig, newFisico) => {
    if (!token) return;
    setSyncState("saving");
    try {
      await saveToOneDrive(token, { config: newConfig, llamados: newLlamados, fisico: newFisico });
      setSyncState("ok");
      setLastSync(new Date());
      setTimeout(() => setSyncState("idle"), 2000);
    } catch {
      setSyncState("error");
    }
  }, [token]);

  /* ── Helpers ── */
  const fv      = (k, v) => setForm(f => ({ ...f, [k]: v }));
  const vdInfo  = v => VEREDICTOS.find(x => x.value === v);
  const arbInfo = id => config.arbitros.find(a => a.id === id);
  const companeroOpts = config.arbitros.filter(a => a.id !== form.arbitro);
  const toggleCompanero = (id) => setForm(f => ({
    ...f,
    companeros: f.companeros.includes(id) ? f.companeros.filter(x => x !== id) : [...f.companeros, id],
  }));

  /* ── Guardar configuración ── */
  const saveCfg = () => {
    const nf = {};
    tmpCfg.arbitros.forEach(a => { nf[a.id] = fisico[a.id] || { nivel:"Bueno", notas:"" }; });
    setConfig(tmpCfg);
    setFisico(nf);
    setForm(f => ({ ...makeForm(tmpCfg.arbitros, tmpCfg.equipo1), periodo: f.periodo }));
    setEditCfg(false);
    syncData(llamados, tmpCfg, nf);
  };

  /* ── CRUD llamados ── */
  const submitLlamado = () => {
    if (!form.minuto) { alert("Ingresa el minuto del llamado."); return; }
    let newLlamados;
    if (editId) {
      newLlamados = llamados.map(x => x.id === editId ? { ...form, id: editId } : x);
      setEditId(null);
    } else {
      newLlamados = [...llamados, { ...form, id: makeId() }];
    }
    setLlamados(newLlamados);
    setForm(f => ({ ...makeForm(config.arbitros, config.equipo1), periodo: f.periodo, arbitro: f.arbitro }));
    syncData(newLlamados, config, fisico);
  };

  const startEdit     = (item) => { setForm({ ...item }); setEditId(item.id); setVista("registro"); window.scrollTo(0,0); };
  const cancelEdit    = () => { setForm(f => ({ ...makeForm(config.arbitros, config.equipo1), periodo: f.periodo, arbitro: f.arbitro })); setEditId(null); };
  const deleteLlamado = (id) => {
    const nl = llamados.filter(x => x.id !== id);
    setLlamados(nl);
    syncData(nl, config, fisico);
  };

  const updateFisico = (newFisico) => { setFisico(newFisico); syncData(llamados, config, newFisico); };

  /* ── Estadísticas ── */
  const stats = useMemo(() => {
    const r = {};
    config.arbitros.forEach(({ id }) => {
      const m = llamados.filter(l => l.arbitro === id);
      const positivos = m.filter(x => VEREDICTOS.find(v => v.value === x.veredicto)?.positivo).length;
      const ef = m.length ? Math.round((positivos / m.length) * 100) : 0;
      r[id] = {
        total:       m.length,
        acertado:    m.filter(x => x.veredicto === "acertado").length,
        aceptable:   m.filter(x => x.veredicto === "aceptable").length,
        marginal:    m.filter(x => x.veredicto === "marginal").length,
        fantasioso:  m.filter(x => x.veredicto === "fantasioso").length,
        correcto_nc: m.filter(x => x.veredicto === "correcto_nc").length,
        nocall_error:m.filter(x => x.veredicto === "nocall_error").length,
        efectividad: ef,
        semaforo:    getSemaforo(ef),
        localCount:  m.filter(x => x.equipo === config.equipo1).length,
        visitCount:  m.filter(x => x.equipo === config.equipo2).length,
        aceptados:   m.filter(x => x.aceptacion === "Aceptado").length,
        noAceptados: m.filter(x => x.aceptacion === "No Aceptado").length,
        porCategoria:{
          falta:    m.filter(x => categoriaDe(x.tipo) === "falta").length,
          violacion:m.filter(x => categoriaDe(x.tipo) === "violacion").length,
          salida:   m.filter(x => categoriaDe(x.tipo) === "salida").length,
        },
        irsCount:    m.filter(x => x.irs).length,
        irsSostenida:m.filter(x => x.irs && x.irsResultado === "sostenida").length,
        irsCambiada: m.filter(x => x.irs && x.irsResultado === "cambiada").length,
        cuestionables: m.filter(x => ["fantasioso","marginal","nocall_error"].includes(x.veredicto)),
      };
    });
    return r;
  }, [llamados, config]);

  const totalLocal   = llamados.filter(x => x.equipo === config.equipo1).length;
  const totalVisit   = llamados.filter(x => x.equipo === config.equipo2).length;
  const totalAcept   = llamados.filter(x => x.aceptacion === "Aceptado").length;
  const totalNoAcept = llamados.filter(x => x.aceptacion === "No Aceptado").length;
  const globalSem    = getSemaforo(llamados.length > 0
    ? Math.round(llamados.filter(x => VEREDICTOS.find(v => v.value === x.veredicto)?.positivo).length / llamados.length * 100)
    : 0);

  /* ══════════════════════════════════════════════════════
     RENDER LOGIN
  ═══════════════════════════════════════════════════════ */
  if (!token) return (
    <div style={{ minHeight:"100vh", background:"#090d18", display:"flex", alignItems:"center", justifyContent:"center", padding:20 }}>
      <style>{CSS}</style>
      <div style={{ textAlign:"center", maxWidth:420 }}>
        <div style={{ fontSize:72, marginBottom:16 }}>🏀</div>
        <div style={{ fontSize:32, fontWeight:800, color:"#fff", letterSpacing:3, textTransform:"uppercase", fontFamily:"'Barlow Condensed',sans-serif" }}>
          Análisis Arbitral
        </div>
        <div style={{ fontSize:14, color:"#546e7a", fontFamily:"Barlow,sans-serif", marginBottom:32, marginTop:8 }}>
          Liga Señal Colombia de Baloncesto
        </div>
        <div className="card" style={{ padding:32, marginBottom:20 }}>
          <div style={{ fontSize:16, fontWeight:700, color:"#90caf9", marginBottom:8, fontFamily:"'Barlow Condensed',sans-serif", letterSpacing:1 }}>
            🔐 SINCRONIZACIÓN CON ONEDRIVE
          </div>
          <div style={{ fontSize:13, color:"#546e7a", fontFamily:"Barlow,sans-serif", marginBottom:24, lineHeight:1.7 }}>
            Tus datos se guardan automáticamente en tu OneDrive de Microsoft. Inicia sesión para continuar.
          </div>
          <button className="btn btn-blue" style={{ width:"100%", fontSize:16, padding:"14px 22px" }}
            onClick={() => window.location.href = buildAuthUrl()}>
            Iniciar sesión con Microsoft
          </button>
        </div>
        <div style={{ fontSize:12, color:"#37474f", fontFamily:"Barlow,sans-serif" }}>
          Los datos se guardan en tu OneDrive personal como archivo JSON.
        </div>
      </div>
    </div>
  );

  /* ══════════════════════════════════════════════════════
     RENDER PRINCIPAL
  ═══════════════════════════════════════════════════════ */
  return (
    <div style={{ minHeight:"100vh", background:"#090d18", color:"#e8eaf6", fontFamily:"'Barlow Condensed',sans-serif" }}>
      <style>{CSS}</style>

      {/* ═══ HEADER ═══ */}
      <div style={{ background:"linear-gradient(135deg,#0b1020,#180826)", borderBottom:"3px solid #e91e63", padding:"22px 28px 18px" }}>
        <div style={{ maxWidth:1100, margin:"0 auto" }}>
          <div style={{ display:"flex", alignItems:"center", gap:14, flexWrap:"wrap" }}>
            <span style={{ fontSize:46, lineHeight:1 }}>🏀</span>
            <div style={{ flex:1 }}>
              <div style={{ fontSize:26, fontWeight:800, letterSpacing:3, textTransform:"uppercase", color:"#fff" }}>Análisis Arbitral</div>
              <div style={{ fontSize:13, color:"#546e7a", fontFamily:"Barlow,sans-serif" }}>
                {config.liga} &nbsp;|&nbsp;
                <span style={{ color:"#90caf9" }}>{config.equipo1} vs {config.equipo2}</span>
                &nbsp;|&nbsp;{config.fecha}
              </div>
              <div style={{ marginTop:5, display:"flex", gap:12, flexWrap:"wrap", alignItems:"center" }}>
                {config.arbitros.map(a => (
                  <span key={a.id} style={{ fontSize:13, color:"#b0bec5", fontFamily:"Barlow,sans-serif" }}>
                    <span style={{ color:"#e91e63", fontWeight:700 }}>{a.rol}:</span> {a.nombre}
                  </span>
                ))}
                {/* Indicador de sync */}
                <span style={{ fontSize:12, color:"#546e7a", fontFamily:"Barlow,sans-serif", marginLeft:8 }}>
                  <span className={`sync-dot ${syncState==="saving"?"sync-spin":syncState==="ok"?"sync-ok":syncState==="error"?"sync-err":"sync-ok"}`} />
                  {syncState==="saving" ? "Guardando..." : syncState==="error" ? "Error al guardar" : lastSync ? `Guardado ${lastSync.toLocaleTimeString()}` : "OneDrive conectado"}
                </span>
              </div>
            </div>
            <div style={{ display:"flex", gap:8, flexWrap:"wrap", alignItems:"center" }}>
              {userInfo && (
                <span style={{ fontSize:12, color:"#546e7a", fontFamily:"Barlow,sans-serif" }}>
                  👤 {userInfo.displayName || userInfo.userPrincipalName}
                </span>
              )}
              <button className="btn btn-ghost" style={{ fontSize:13, padding:"7px 16px" }}
                onClick={() => { setTmpCfg(config); setEditCfg(true); }}>
                ⚙️ Configurar
              </button>
              <button className="btn btn-del btn-sm"
                onClick={() => { localStorage.removeItem("msft_token"); localStorage.removeItem("msft_expiry"); window.location.reload(); }}>
                Cerrar sesión
              </button>
            </div>
          </div>
        </div>
      </div>

      {/* ═══ MODAL CONFIGURACIÓN ═══ */}
      {editCfg && (
        <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,.85)", zIndex:200, display:"flex", alignItems:"center", justifyContent:"center", padding:20 }}>
          <div className="card" style={{ width:"100%", maxWidth:640, maxHeight:"90vh", overflowY:"auto", padding:28 }}>
            <div className="sec">⚙️ Configuración del Partido</div>
            <div style={{ display:"grid", gap:14 }}>
              <div><div className="lbl">Liga / Torneo</div>
                <input value={tmpCfg.liga} onChange={e => setTmpCfg(c => ({ ...c, liga:e.target.value }))} />
              </div>
              <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12 }}>
                <div><div className="lbl">Equipo Local</div>
                  <input value={tmpCfg.equipo1} onChange={e => setTmpCfg(c => ({ ...c, equipo1:e.target.value }))} />
                </div>
                <div><div className="lbl">Equipo Visitante</div>
                  <input value={tmpCfg.equipo2} onChange={e => setTmpCfg(c => ({ ...c, equipo2:e.target.value }))} />
                </div>
              </div>
              <div><div className="lbl">Fecha</div>
                <input type="date" value={tmpCfg.fecha} onChange={e => setTmpCfg(c => ({ ...c, fecha:e.target.value }))} />
              </div>
              <div><div className="lbl">Nota Grupal — Trabajo en Equipo (aparece en el reporte)</div>
                <textarea rows={3} placeholder="Evaluación general de la terna como equipo..."
                  value={tmpCfg.notaGrupal || ""}
                  onChange={e => setTmpCfg(c => ({ ...c, notaGrupal:e.target.value }))}
                  style={{ resize:"vertical" }} />
              </div>
              <div style={{ borderTop:"1px solid #1d2840", paddingTop:14 }}>
                <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:12 }}>
                  <div style={{ fontSize:15, fontWeight:700, color:"#e91e63" }}>TERNA ARBITRAL</div>
                  <button className="btn btn-ghost btn-sm"
                    onClick={() => setTmpCfg(c => ({ ...c, arbitros:[...c.arbitros,{ id:makeId(), nombre:"", rol:"Crew Chief", foto:"" }] }))}>
                    + Árbitro
                  </button>
                </div>
                {tmpCfg.arbitros.map(a => (
                  <div key={a.id} style={{ display:"grid", gridTemplateColumns:"1fr 1fr auto", gap:8, marginBottom:8, alignItems:"center" }}>
                    <input placeholder="Nombre" value={a.nombre}
                      onChange={e => setTmpCfg(c => ({ ...c, arbitros:c.arbitros.map(x => x.id===a.id?{...x,nombre:e.target.value}:x) }))} />
                    <select value={a.rol}
                      onChange={e => setTmpCfg(c => ({ ...c, arbitros:c.arbitros.map(x => x.id===a.id?{...x,rol:e.target.value}:x) }))}>
                      {ROLES.map(r => <option key={r}>{r}</option>)}
                    </select>
                    <button className="btn btn-del btn-sm" disabled={tmpCfg.arbitros.length<=1}
                      onClick={() => setTmpCfg(c => ({ ...c, arbitros:c.arbitros.filter(x=>x.id!==a.id) }))}>🗑</button>
                  </div>
                ))}
              </div>
            </div>
            <div style={{ display:"flex", gap:10, marginTop:20 }}>
              <button className="btn btn-red" onClick={saveCfg}>💾 GUARDAR</button>
              <button className="btn btn-ghost" onClick={() => setEditCfg(false)}>Cancelar</button>
            </div>
          </div>
        </div>
      )}

      {/* ═══ TABS ═══ */}
      <div style={{ background:"#090d18", padding:"0 28px", borderBottom:"1px solid #1d2840" }}>
        <div style={{ maxWidth:1100, margin:"0 auto", display:"flex", gap:4, paddingTop:14, flexWrap:"wrap" }}>
          {[
            ["registro","📋 Registrar"],
            ["lista",`📑 Lista (${llamados.length})`],
            ["reporte","📊 Reporte"],
          ].map(([v,lbl]) => (
            <button key={v} className={`tab ${vista===v?"tab-on":"tab-off"}`} onClick={() => setVista(v)}>{lbl}</button>
          ))}
        </div>
      </div>

      <div style={{ maxWidth:1100, margin:"0 auto", padding:"26px 28px" }}>

        {/* ═══ VISTA: REGISTRO ═══ */}
        {vista === "registro" && (
          <div>
            <div className="sec">{editId ? "✏️ Editar Llamado" : "➕ Registrar Llamado"}</div>

            {/* Período */}
            <div style={{ marginBottom:16 }}>
              <div className="lbl" style={{ marginBottom:8 }}>PERÍODO ACTIVO</div>
              <div style={{ display:"flex", gap:8, flexWrap:"wrap" }}>
                {PERIODOS.map(p => (
                  <button key={p} onClick={() => fv("periodo",p)} className={`tgl ${form.periodo===p?"tgl-blue":""}`}>{p}</button>
                ))}
              </div>
            </div>

            <div className="card" style={{ padding:22 }}>

              {/* CATEGORÍA — tabs visuales */}
              <div style={{ marginBottom:18 }}>
                <div className="lbl" style={{ marginBottom:10 }}>CATEGORÍA DEL LLAMADO</div>
                <div style={{ display:"flex", gap:8, flexWrap:"wrap" }}>
                  {Object.entries(CATEGORIAS).map(([key,cat]) => (
                    <button key={key} onClick={() => {
                      fv("categoria", key);
                      fv("tipo", TIPOS_POR_CATEGORIA[key][0]);
                    }} style={{
                      cursor:"pointer", borderRadius:8, padding:"10px 18px",
                      fontFamily:"'Barlow Condensed',sans-serif", fontWeight:700, fontSize:14,
                      border:`2px solid ${form.categoria===key ? cat.color : "transparent"}`,
                      background: form.categoria===key ? `${cat.color}1a` : "#1d2840",
                      color: form.categoria===key ? cat.color : "#546e7a",
                      transition:"all .15s",
                    }}>
                      {cat.icon} {cat.label}
                    </button>
                  ))}
                </div>
              </div>

              <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(185px,1fr))", gap:14 }}>

                <div><div className="lbl">Minuto</div>
                  <input placeholder="Ej: 3:45" value={form.minuto} onChange={e => fv("minuto",e.target.value)} />
                </div>

                <div><div className="lbl">Árbitro que Pita</div>
                  <select value={form.arbitro} onChange={e => fv("arbitro",e.target.value)}>
                    {config.arbitros.map(a => <option key={a.id} value={a.id}>{a.nombre} ({a.rol})</option>)}
                  </select>
                </div>

                <div><div className="lbl">Posición del Árbitro</div>
                  <select value={form.posArbitro} onChange={e => fv("posArbitro",e.target.value)}>
                    {POS_ARB.map(p => <option key={p}>{p}</option>)}
                  </select>
                </div>

                <div><div className="lbl">Zona del Llamado</div>
                  <select value={form.zona} onChange={e => fv("zona",e.target.value)}>
                    {ZONAS.map(z => <option key={z}>{z}</option>)}
                  </select>
                </div>

                <div><div className="lbl">Tipo de Llamado</div>
                  <select value={form.tipo} onChange={e => fv("tipo",e.target.value)}>
                    {TIPOS_POR_CATEGORIA[form.categoria].map(t => <option key={t}>{t}</option>)}
                  </select>
                </div>

                <div><div className="lbl">Equipo Sancionado</div>
                  <select value={form.equipo} onChange={e => fv("equipo",e.target.value)}>
                    <option value={config.equipo1}>{config.equipo1}</option>
                    <option value={config.equipo2}>{config.equipo2}</option>
                    <option value="Ambos (Doble Foul)">Ambos (Doble Foul)</option>
                  </select>
                </div>

                {/* Aceptación */}
                <div><div className="lbl">Aceptación del Llamado</div>
                  <div style={{ display:"flex", gap:8, marginTop:4 }}>
                    <button onClick={() => fv("aceptacion","Aceptado")}
                      className={`tgl ${form.aceptacion==="Aceptado"?"tgl-green":""}`} style={{ flex:1 }}>✔ Aceptado</button>
                    <button onClick={() => fv("aceptacion","No Aceptado")}
                      className={`tgl ${form.aceptacion==="No Aceptado"?"tgl-red":""}`} style={{ flex:1 }}>✘ No Aceptado</button>
                  </div>
                </div>

                {/* Pitazo */}
                <div style={{ gridColumn:"1 / -1" }}>
                  <div className="lbl" style={{ marginBottom:8 }}>Tipo de Pitazo</div>
                  <div style={{ display:"flex", gap:8, flexWrap:"wrap", marginBottom:10 }}>
                    {[["simple","🔔 Simple"],["doble","🔔🔔 Doble"],["triple","🔔🔔🔔 Triple"]].map(([v,lbl]) => (
                      <button key={v} onClick={() => fv("pitazo",v)}
                        className={`tgl ${form.pitazo===v?"tgl-orange":""}`}>{lbl}</button>
                    ))}
                  </div>
                  {(form.pitazo==="doble"||form.pitazo==="triple") && (
                    <div>
                      <div className="lbl" style={{ marginBottom:7 }}>
                        {form.pitazo==="doble"?"Compañero que también pitó:":"Compañeros que también pitaron:"}
                      </div>
                      <div style={{ display:"flex", gap:8, flexWrap:"wrap" }}>
                        {companeroOpts.map(a => (
                          <button key={a.id} onClick={() => toggleCompanero(a.id)}
                            className={`tgl ${form.companeros.includes(a.id)?"tgl-orange":""}`}>
                            {a.nombre} ({a.rol})
                          </button>
                        ))}
                      </div>
                    </div>
                  )}
                </div>

                {/* IRS */}
                <div style={{ gridColumn:"1 / -1" }}>
                  <div className="lbl" style={{ marginBottom:8 }}>REVISIÓN IRS</div>
                  <div style={{ display:"flex", gap:8, alignItems:"center", flexWrap:"wrap" }}>
                    <button onClick={() => fv("irs", !form.irs)}
                      className={`tgl ${form.irs?"tgl-pink":""}`}>
                      {form.irs ? "✅ IRS Revisado" : "☐ Marcar revisión IRS"}
                    </button>
                    {form.irs && (
                      <>
                        <button onClick={() => fv("irsResultado","sostenida")}
                          className={`tgl ${form.irsResultado==="sostenida"?"tgl-green":""}`}>
                          ✔ Decisión Sostenida
                        </button>
                        <button onClick={() => fv("irsResultado","cambiada")}
                          className={`tgl ${form.irsResultado==="cambiada"?"tgl-red":""}`}>
                          ↩ Decisión Cambiada
                        </button>
                      </>
                    )}
                  </div>
                </div>

                {/* Veredicto */}
                <div style={{ gridColumn:"1 / -1" }}>
                  <div className="lbl" style={{ marginBottom:8 }}>Veredicto del Llamado</div>
                  <div style={{ display:"flex", gap:8, flexWrap:"wrap" }}>
                    {VEREDICTOS.map(v => (
                      <button key={v.value} onClick={() => fv("veredicto",v.value)} style={{
                        cursor:"pointer",
                        border:`2px solid ${form.veredicto===v.value?v.color:"transparent"}`,
                        background:form.veredicto===v.value?v.bg:"#1d2840",
                        color:form.veredicto===v.value?v.color:"#546e7a",
                        borderRadius:8, padding:"8px 14px",
                        fontFamily:"'Barlow Condensed',sans-serif",
                        fontWeight:700, fontSize:13, transition:"all .15s",
                      }}>
                        {form.veredicto===v.value?"● ":""}{v.label}
                      </button>
                    ))}
                  </div>
                  <div style={{ marginTop:8, fontSize:12, color:"#546e7a", fontFamily:"Barlow,sans-serif" }}>
                    <span style={{ color:"#00e676" }}>Positivos:</span> Acertado · Aceptable · Correcto No Llamado &nbsp;|&nbsp;
                    <span style={{ color:"#ff5252" }}>Negativos:</span> Marginal · Fantasioso · No Call (Error)
                  </div>
                </div>

                {/* Descripción */}
                <div style={{ gridColumn:"1 / -1" }}>
                  <div className="lbl">Observación (opcional)</div>
                  <textarea rows={2} placeholder="Jugada, posición de jugadores, contexto..."
                    value={form.descripcion} onChange={e => fv("descripcion",e.target.value)}
                    style={{ resize:"vertical" }} />
                </div>
              </div>

              <div style={{ marginTop:18, display:"flex", gap:10, flexWrap:"wrap" }}>
                <button className="btn btn-red" onClick={submitLlamado}>
                  {editId?"💾 GUARDAR CAMBIOS":"➕ AGREGAR LLAMADO"}
                </button>
                {editId && <button className="btn btn-ghost" onClick={cancelEdit}>Cancelar</button>}
              </div>
            </div>

            {/* Estado físico */}
            <div style={{ marginTop:26 }}>
              <div className="sec">🏃 Estado Físico por Árbitro</div>
              <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(290px,1fr))", gap:14 }}>
                {config.arbitros.map(a => (
                  <div className="card" key={a.id} style={{ padding:16 }}>
                    <div style={{ fontSize:16, fontWeight:700, color:"#90caf9", marginBottom:10 }}>
                      {a.nombre} <span style={{ color:"#546e7a", fontSize:12 }}>— {a.rol}</span>
                    </div>
                    <div className="lbl">Nivel Físico</div>
                    <select value={fisico[a.id]?.nivel||"Bueno"}
                      onChange={e => updateFisico({ ...fisico, [a.id]:{ ...fisico[a.id], nivel:e.target.value } })}>
                      {FISICO_OPT.map(f => <option key={f}>{f}</option>)}
                    </select>
                    <div className="lbl" style={{ marginTop:10 }}>Observaciones Físicas</div>
                    <textarea rows={2} placeholder="Movilidad, transiciones, cansancio..."
                      value={fisico[a.id]?.notas||""}
                      onChange={e => updateFisico({ ...fisico, [a.id]:{ ...fisico[a.id], notas:e.target.value } })}
                      style={{ resize:"vertical" }} />
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

        {/* ═══ VISTA: LISTA ═══ */}
        {vista === "lista" && (
          <div>
            <div className="sec">📑 Todos los Llamados Registrados</div>
            {llamados.length === 0
              ? <div className="card" style={{ padding:40, textAlign:"center", color:"#546e7a", fontFamily:"Barlow,sans-serif" }}>
                  Aún no hay llamados registrados.
                </div>
              : [...llamados].reverse().map(item => {
                  const a   = arbInfo(item.arbitro);
                  const vv  = vdInfo(item.veredicto);
                  const cat = CATEGORIAS[item.categoria || categoriaDe(item.tipo)];
                  return (
                    <div key={item.id} className="row-item">
                      <div style={{ minWidth:78, fontSize:12, color:"#546e7a", fontFamily:"Barlow,sans-serif", flexShrink:0 }}>
                        {item.periodo}<br/>
                        <span style={{ color:"#e91e63", fontWeight:700, fontSize:14 }}>{item.minuto}</span>
                      </div>
                      <div style={{ minWidth:130, flexShrink:0 }}>
                        <div style={{ fontSize:13, fontWeight:700, color:"#90caf9" }}>{a?.nombre||"?"}</div>
                        <div style={{ fontSize:11, color:"#546e7a" }}>{a?.rol} · {item.posArbitro}</div>
                        <div style={{ fontSize:11, color:"#546e7a" }}>{item.zona}</div>
                        <div style={{ fontSize:11, color:"#ff9100" }}>
                          {item.pitazo==="doble"?"🔔🔔":item.pitazo==="triple"?"🔔🔔🔔":"🔔"}
                        </div>
                        {item.irs && (
                          <div style={{ fontSize:11, color: item.irsResultado==="cambiada"?"#ff5252":"#00e676" }}>
                            IRS: {item.irsResultado==="cambiada"?"↩ Cambiada":"✔ Sostenida"}
                          </div>
                        )}
                      </div>
                      <div style={{ flex:1, minWidth:0 }}>
                        <span style={{ fontSize:11, background:`${cat?.color}22`, color:cat?.color, borderRadius:4, padding:"2px 7px", fontFamily:"Barlow,sans-serif", marginRight:6 }}>
                          {cat?.icon} {cat?.label}
                        </span>
                        <div style={{ fontSize:14, fontWeight:600, marginTop:4 }}>{item.tipo}</div>
                        <div style={{ fontSize:12, color:"#546e7a", fontFamily:"Barlow,sans-serif" }}>{item.equipo}</div>
                        {item.descripcion && <div style={{ fontSize:12, color:"#90a4ae", fontFamily:"Barlow,sans-serif", marginTop:2 }}>{item.descripcion}</div>}
                      </div>
                      <div style={{ display:"flex", flexDirection:"column", gap:5, flexShrink:0, alignItems:"flex-end" }}>
                        <span className="badge" style={{ background:vv.bg, color:vv.color, border:`1px solid ${vv.color}` }}>{vv.label}</span>
                        <span className="badge" style={{
                          background: item.aceptacion==="Aceptado"?"rgba(0,230,118,0.1)":"rgba(255,23,68,0.1)",
                          color:      item.aceptacion==="Aceptado"?"#00e676":"#ff1744",
                          border:     `1px solid ${item.aceptacion==="Aceptado"?"#00e676":"#ff1744"}`,
                        }}>
                          {item.aceptacion==="Aceptado"?"✔ Aceptado":"✘ No Aceptado"}
                        </span>
                      </div>
                      <div style={{ display:"flex", gap:5, flexShrink:0 }}>
                        <button className="btn btn-warn btn-sm" onClick={() => startEdit(item)}>✏️</button>
                        <button className="btn btn-del  btn-sm" onClick={() => deleteLlamado(item.id)}>🗑</button>
                      </div>
                    </div>
                  );
                })
            }
          </div>
        )}

        {/* ═══ VISTA: REPORTE ═══ */}
        {vista === "reporte" && (
          <div>
            <div className="sec">📊 Reporte Final — Análisis Arbitral</div>
            {llamados.length === 0
              ? <div className="card" style={{ padding:40, textAlign:"center", color:"#546e7a", fontFamily:"Barlow,sans-serif" }}>
                  Registra llamados primero para generar el reporte.
                </div>
              : (<>
                  {/* ─ Nota grupal ─ */}
                  {config.notaGrupal && (
                    <div className="card" style={{ padding:22, marginBottom:18, borderLeft:"4px solid #e91e63" }}>
                      <div style={{ fontSize:15, fontWeight:800, color:"#e91e63", marginBottom:8, letterSpacing:1 }}>🤝 TRABAJO EN EQUIPO — NOTA GRUPAL</div>
                      <div style={{ fontSize:14, color:"#b0bec5", fontFamily:"Barlow,sans-serif", lineHeight:1.7 }}>{config.notaGrupal}</div>
                    </div>
                  )}

                  {/* ─ Resumen global ─ */}
                  <div className="card" style={{ padding:22, marginBottom:22 }}>
                    <div style={{ fontSize:16, fontWeight:700, color:"#e91e63", marginBottom:16, letterSpacing:1 }}>RESUMEN GLOBAL DEL PARTIDO</div>

                    {/* Semáforo global */}
                    <div style={{ display:"flex", alignItems:"center", gap:16, marginBottom:20, background:"#090d18", borderRadius:10, padding:"14px 18px", border:`1px solid ${globalSem.color}33` }}>
                      <div style={{ fontSize:42, fontWeight:800, color:globalSem.color, lineHeight:1 }}>
                        {llamados.length > 0 ? Math.round(llamados.filter(x=>VEREDICTOS.find(v=>v.value===x.veredicto)?.positivo).length/llamados.length*100) : 0}%
                      </div>
                      <div>
                        <div style={{ fontSize:22, fontWeight:800, color:globalSem.color, letterSpacing:2 }}>{globalSem.label.toUpperCase()}</div>
                        <div style={{ fontSize:12, color:"#546e7a", fontFamily:"Barlow,sans-serif" }}>Efectividad global de la terna</div>
                      </div>
                      <div style={{ flex:1 }}>
                        <div className="prog-bg">
                          <div className="prog-fill" style={{ width:`${llamados.length>0?Math.round(llamados.filter(x=>VEREDICTOS.find(v=>v.value===x.veredicto)?.positivo).length/llamados.length*100):0}%`, background:globalSem.color }} />
                        </div>
                        <div style={{ display:"flex", gap:6, marginTop:8, flexWrap:"wrap" }}>
                          {SEMAFORO.map(s => (
                            <span key={s.label} style={{ fontSize:10, color:s.color, fontFamily:"Barlow,sans-serif", opacity:0.7 }}>
                              {s.label} {s.min}–{s.max}%
                            </span>
                          ))}
                        </div>
                      </div>
                    </div>

                    {/* Veredictos globales */}
                    <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(100px,1fr))", gap:10, marginBottom:20 }}>
                      {[
                        { label:"Total",          val:llamados.length,                                               color:"#90caf9" },
                        { label:"Acertados",       val:llamados.filter(x=>x.veredicto==="acertado").length,           color:"#00e676" },
                        { label:"Aceptables",      val:llamados.filter(x=>x.veredicto==="aceptable").length,          color:"#ffd740" },
                        { label:"Correcto NC",     val:llamados.filter(x=>x.veredicto==="correcto_nc").length,        color:"#00bcd4" },
                        { label:"Marginales",      val:llamados.filter(x=>x.veredicto==="marginal").length,           color:"#ff9100" },
                        { label:"Fantasiosos",     val:llamados.filter(x=>x.veredicto==="fantasioso").length,         color:"#ff1744" },
                        { label:"No Call Error",   val:llamados.filter(x=>x.veredicto==="nocall_error").length,       color:"#aa00ff" },
                      ].map(s => (
                        <div key={s.label} style={{ background:"#090d18", border:`1px solid ${s.color}22`, borderRadius:10, padding:"12px 10px", textAlign:"center" }}>
                          <div className="lbl">{s.label}</div>
                          <div style={{ fontSize:26, fontWeight:800, color:s.color, lineHeight:1 }}>{s.val}</div>
                        </div>
                      ))}
                    </div>

                    {/* Desglose por categoría global */}
                    <div style={{ marginBottom:20 }}>
                      <div className="lbl" style={{ marginBottom:10 }}>DESGLOSE POR CATEGORÍA</div>
                      <div className="grid3">
                        {Object.entries(CATEGORIAS).map(([key,cat]) => {
                          const count = llamados.filter(x => (x.categoria||categoriaDe(x.tipo)) === key).length;
                          return (
                            <div key={key} style={{ background:"#090d18", border:`1px solid ${cat.color}33`, borderRadius:10, padding:"14px 16px" }}>
                              <div style={{ fontSize:13, color:cat.color, fontWeight:700 }}>{cat.icon} {cat.label}</div>
                              <div style={{ fontSize:30, fontWeight:800, color:cat.color, lineHeight:1.2 }}>{count}</div>
                              <div style={{ fontSize:11, color:"#546e7a", fontFamily:"Barlow,sans-serif" }}>{pct(count,llamados.length)} del total</div>
                              <div className="prog-bg" style={{ marginTop:6 }}>
                                <div className="prog-fill" style={{ width:pct(count,llamados.length), background:cat.color }} />
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </div>

                    {/* Faltas por equipo */}
                    <div style={{ marginBottom:20 }}>
                      <div className="lbl" style={{ marginBottom:10 }}>SANCIONES POR EQUIPO</div>
                      <div className="grid2">
                        {[
                          { label:config.equipo1, val:totalLocal, color:"#29b6f6" },
                          { label:config.equipo2, val:totalVisit, color:"#f06292" },
                        ].map(e => (
                          <div key={e.label} style={{ background:"#090d18", border:`1px solid ${e.color}33`, borderRadius:10, padding:"16px 18px" }}>
                            <div style={{ fontSize:13, color:e.color, fontWeight:700, marginBottom:4 }}>{e.label}</div>
                            <div style={{ fontSize:34, fontWeight:800, color:e.color, lineHeight:1 }}>{e.val}</div>
                            <div style={{ fontSize:12, color:"#546e7a", fontFamily:"Barlow,sans-serif", marginTop:4 }}>{pct(e.val,llamados.length)} del total</div>
                            <div className="prog-bg" style={{ marginTop:8 }}>
                              <div className="prog-fill" style={{ width:pct(e.val,llamados.length), background:e.color }} />
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>

                    {/* Aceptación global */}
                    <div>
                      <div className="lbl" style={{ marginBottom:10 }}>ACEPTACIÓN GLOBAL</div>
                      <div className="grid2">
                        {[
                          { label:"Aceptados",    val:totalAcept,   color:"#00e676" },
                          { label:"No Aceptados", val:totalNoAcept, color:"#ff1744" },
                        ].map(e => (
                          <div key={e.label} style={{ background:"#090d18", border:`1px solid ${e.color}33`, borderRadius:10, padding:"16px 18px" }}>
                            <div style={{ fontSize:13, color:e.color, fontWeight:700, marginBottom:4 }}>{e.label}</div>
                            <div style={{ fontSize:34, fontWeight:800, color:e.color, lineHeight:1 }}>{e.val}</div>
                            <div style={{ fontSize:12, color:"#546e7a", fontFamily:"Barlow,sans-serif", marginTop:4 }}>{pct(e.val,llamados.length)} del total</div>
                            <div className="prog-bg" style={{ marginTop:8 }}>
                              <div className="prog-fill" style={{ width:pct(e.val,llamados.length), background:e.color }} />
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  </div>

                  {/* ─ Detalle por árbitro ─ */}
                  {config.arbitros.map(a => {
                    const s   = stats[a.id];
                    const fis = fisico[a.id] || { nivel:"Bueno", notas:"" };
                    const sem = s.semaforo;
                    return (
                      <div key={a.id} className="card" style={{ padding:22, marginBottom:20 }}>

                        {/* Cabecera */}
                        <div style={{ display:"flex", alignItems:"center", gap:14, marginBottom:18, flexWrap:"wrap" }}>
                          <div style={{ width:52, height:52, borderRadius:"50%", background:"linear-gradient(135deg,#e91e63,#880e4f)", display:"flex", alignItems:"center", justifyContent:"center", fontSize:24, flexShrink:0 }}>🧑‍⚖️</div>
                          <div style={{ flex:1 }}>
                            <div style={{ fontSize:24, fontWeight:800, color:"#fff" }}>{a.nombre}</div>
                            <div style={{ fontSize:12, color:"#546e7a", fontFamily:"Barlow,sans-serif" }}>
                              {a.rol} &nbsp;|&nbsp; Estado físico: <span style={{ color:"#90caf9" }}>{fis.nivel}</span>
                            </div>
                          </div>
                          <div style={{ textAlign:"center", background:`${sem.color}15`, borderRadius:10, padding:"10px 18px", border:`1px solid ${sem.color}44` }}>
                            <div style={{ fontSize:38, fontWeight:800, color:sem.color, lineHeight:1 }}>{s.efectividad}%</div>
                            <div style={{ fontSize:14, fontWeight:800, color:sem.color, letterSpacing:1 }}>{sem.label.toUpperCase()}</div>
                            <div className="lbl" style={{ marginTop:2 }}>Efectividad</div>
                          </div>
                        </div>

                        {/* Barra efectividad */}
                        <div style={{ marginBottom:16 }}>
                          <div style={{ display:"flex", justifyContent:"space-between", marginBottom:4 }}>
                            <span style={{ fontSize:11, color:"#546e7a", fontFamily:"Barlow,sans-serif" }}>EFECTIVIDAD (positivos / total)</span>
                            <span style={{ fontSize:11, color:sem.color, fontFamily:"Barlow,sans-serif", fontWeight:700 }}>{s.efectividad}% — {sem.label}</span>
                          </div>
                          <div className="prog-bg">
                            <div className="prog-fill" style={{ width:`${s.efectividad}%`, background:sem.color }} />
                          </div>
                        </div>

                        {/* Veredictos por árbitro */}
                        <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(85px,1fr))", gap:9, marginBottom:16 }}>
                          {[
                            { label:"Total",        val:s.total,       color:"#90caf9" },
                            { label:"Acertados",    val:s.acertado,    color:"#00e676" },
                            { label:"Aceptables",   val:s.aceptable,   color:"#ffd740" },
                            { label:"Correcto NC",  val:s.correcto_nc, color:"#00bcd4" },
                            { label:"Marginales",   val:s.marginal,    color:"#ff9100" },
                            { label:"Fantasiosos",  val:s.fantasioso,  color:"#ff1744" },
                            { label:"No Call Err",  val:s.nocall_error,color:"#aa00ff" },
                          ].map(st => (
                            <div key={st.label} style={{ background:"#090d18", border:`1px solid ${st.color}33`, borderRadius:9, padding:"9px 8px", textAlign:"center" }}>
                              <div className="lbl">{st.label}</div>
                              <div style={{ fontSize:22, fontWeight:800, color:st.color, lineHeight:1 }}>{st.val}</div>
                            </div>
                          ))}
                        </div>

                        {/* Desglose por categoría — árbitro */}
                        <div style={{ marginBottom:16 }}>
                          <div className="lbl" style={{ marginBottom:10 }}>LLAMADOS POR CATEGORÍA</div>
                          <div className="grid3">
                            {Object.entries(CATEGORIAS).map(([key,cat]) => (
                              <div key={key} style={{ background:"#090d18", border:`1px solid ${cat.color}33`, borderRadius:9, padding:"12px 14px" }}>
                                <div style={{ fontSize:12, color:cat.color, fontWeight:700 }}>{cat.icon} {cat.label}</div>
                                <div style={{ fontSize:26, fontWeight:800, color:cat.color, lineHeight:1.2 }}>{s.porCategoria[key]}</div>
                                <div style={{ fontSize:11, color:"#546e7a", fontFamily:"Barlow,sans-serif" }}>{pct(s.porCategoria[key],s.total)}</div>
                              </div>
                            ))}
                          </div>
                        </div>

                        {/* IRS */}
                        {s.irsCount > 0 && (
                          <div style={{ marginBottom:16 }}>
                            <div className="lbl" style={{ marginBottom:10 }}>REVISIONES IRS</div>
                            <div className="grid3">
                              {[
                                { label:"Total IRS", val:s.irsCount,     color:"#90caf9" },
                                { label:"Sostenidas",val:s.irsSostenida, color:"#00e676" },
                                { label:"Cambiadas", val:s.irsCambiada,  color:"#ff5252" },
                              ].map(x => (
                                <div key={x.label} style={{ background:"#090d18", border:`1px solid ${x.color}33`, borderRadius:9, padding:"12px 14px" }}>
                                  <div style={{ fontSize:12, color:x.color, fontWeight:700 }}>{x.label}</div>
                                  <div style={{ fontSize:26, fontWeight:800, color:x.color, lineHeight:1.2 }}>{x.val}</div>
                                </div>
                              ))}
                            </div>
                          </div>
                        )}

                        {/* Sanciones por equipo */}
                        <div style={{ marginBottom:16 }}>
                          <div className="lbl" style={{ marginBottom:10 }}>SANCIONES POR EQUIPO</div>
                          <div className="grid2">
                            {[
                              { label:config.equipo1, val:s.localCount, color:"#29b6f6" },
                              { label:config.equipo2, val:s.visitCount, color:"#f06292" },
                            ].map(e => (
                              <div key={e.label} style={{ background:"#090d18", border:`1px solid ${e.color}33`, borderRadius:9, padding:"12px 14px" }}>
                                <div style={{ fontSize:12, color:e.color, fontWeight:700 }}>{e.label}</div>
                                <div style={{ fontSize:28, fontWeight:800, color:e.color, lineHeight:1.1 }}>{e.val}</div>
                                <div style={{ fontSize:11, color:"#546e7a", fontFamily:"Barlow,sans-serif" }}>{pct(e.val,s.total)}</div>
                                <div className="prog-bg" style={{ marginTop:6 }}>
                                  <div className="prog-fill" style={{ width:pct(e.val,s.total), background:e.color }} />
                                </div>
                              </div>
                            ))}
                          </div>
                        </div>

                        {/* Aceptación */}
                        <div style={{ marginBottom:16 }}>
                          <div className="lbl" style={{ marginBottom:10 }}>ACEPTACIÓN DE SUS LLAMADOS</div>
                          <div className="grid2">
                            {[
                              { label:"Aceptados",    val:s.aceptados,   color:"#00e676" },
                              { label:"No Aceptados", val:s.noAceptados, color:"#ff1744" },
                            ].map(e => (
                              <div key={e.label} style={{ background:"#090d18", border:`1px solid ${e.color}33`, borderRadius:9, padding:"12px 14px" }}>
                                <div style={{ fontSize:12, color:e.color, fontWeight:700 }}>{e.label}</div>
                                <div style={{ fontSize:28, fontWeight:800, color:e.color, lineHeight:1.1 }}>{e.val}</div>
                                <div style={{ fontSize:11, color:"#546e7a", fontFamily:"Barlow,sans-serif" }}>{pct(e.val,s.total)}</div>
                                <div className="prog-bg" style={{ marginTop:6 }}>
                                  <div className="prog-fill" style={{ width:pct(e.val,s.total), background:e.color }} />
                                </div>
                              </div>
                            ))}
                          </div>
                        </div>

                        {/* Llamados cuestionables */}
                        {s.cuestionables.length > 0 && (
                          <div style={{ marginBottom:16 }}>
                            <div className="lbl" style={{ color:"#ff9100", marginBottom:8 }}>⚠️ LLAMADOS CUESTIONABLES</div>
                            {s.cuestionables.map(item => {
                              const vv  = vdInfo(item.veredicto);
                              const cat = CATEGORIAS[item.categoria||categoriaDe(item.tipo)];
                              return (
                                <div key={item.id} style={{ background:"#090d18", borderRadius:8, padding:"10px 14px", marginBottom:5, borderLeft:`3px solid ${vv.color}` }}>
                                  <div style={{ display:"flex", gap:10, alignItems:"center", flexWrap:"wrap" }}>
                                    <span style={{ fontSize:13, color:"#e91e63", fontWeight:700 }}>{item.periodo} · {item.minuto}</span>
                                    <span style={{ fontSize:11, background:`${cat?.color}22`, color:cat?.color, borderRadius:4, padding:"2px 7px" }}>{cat?.icon} {cat?.label}</span>
                                    <span style={{ fontSize:13, color:"#fff", fontWeight:600 }}>{item.tipo}</span>
                                    <span className="badge" style={{ background:vv.bg, color:vv.color, border:`1px solid ${vv.color}` }}>{vv.label}</span>
                                  </div>
                                  {item.descripcion && <div style={{ fontSize:12, color:"#78909c", fontFamily:"Barlow,sans-serif", marginTop:3 }}>{item.descripcion}</div>}
                                </div>
                              );
                            })}
                          </div>
                        )}

                        {/* Notas físico */}
                        {fis.notas && (
                          <div style={{ background:"#090d18", borderRadius:8, padding:"12px 14px", marginBottom:16, borderLeft:"3px solid #29b6f6" }}>
                            <div style={{ fontSize:12, color:"#29b6f6", fontWeight:700, marginBottom:3 }}>🏃 ESTADO FÍSICO</div>
                            <div style={{ fontSize:13, color:"#b0bec5", fontFamily:"Barlow,sans-serif" }}>{fis.notas}</div>
                          </div>
                        )}

                        {/* Fortalezas y oportunidades */}
                        <div className="grid2">
                          <div style={{ background:"rgba(0,230,118,.06)", border:"1px solid rgba(0,230,118,.2)", borderRadius:10, padding:16 }}>
                            <div style={{ fontSize:14, fontWeight:800, color:"#00e676", marginBottom:10, letterSpacing:1 }}>✅ FORTALEZAS</div>
                            <ul style={{ paddingLeft:17, fontFamily:"Barlow,sans-serif", fontSize:13, color:"#b0bec5", lineHeight:1.9 }}>
                              {s.acertado > 0       && <li>{s.acertado} llamado(s) bien acertado(s).</li>}
                              {s.correcto_nc > 0    && <li>{s.correcto_nc} correcto(s) no llamado(s) — buena paciencia.</li>}
                              {s.efectividad >= 80  && <li>Alta efectividad ({s.efectividad}% — {sem.label}).</li>}
                              {s.fantasioso === 0   && <li>Sin llamados fantasiosos.</li>}
                              {s.nocall_error === 0 && <li>Sin No Calls por error detectados.</li>}
                              {s.irsSostenida > 0   && <li>{s.irsSostenida} revisión/es IRS sostenida(s).</li>}
                              {s.aceptados >= s.noAceptados && s.total > 0 && <li>Mayor porcentaje aceptado ({pct(s.aceptados,s.total)}).</li>}
                              {s.total === 0        && <li>Sin datos registrados aún.</li>}
                            </ul>
                          </div>
                          <div style={{ background:"rgba(255,23,68,.06)", border:"1px solid rgba(255,23,68,.2)", borderRadius:10, padding:16 }}>
                            <div style={{ fontSize:14, fontWeight:800, color:"#ff5252", marginBottom:10, letterSpacing:1 }}>🎯 OPORTUNIDADES DE MEJORA</div>
                            <ul style={{ paddingLeft:17, fontFamily:"Barlow,sans-serif", fontSize:13, color:"#b0bec5", lineHeight:1.9 }}>
                              {s.fantasioso > 0     && <li>{s.fantasioso} llamado(s) fantasioso(s) — revisar umbral.</li>}
                              {s.marginal > 1       && <li>{s.marginal} marginales — afinar criterio de contacto.</li>}
                              {s.nocall_error > 0   && <li>{s.nocall_error} No Call(s) por error — mejorar ángulo y anticipación.</li>}
                              {s.irsCambiada > 0    && <li>{s.irsCambiada} decisión/es IRS cambiada(s) — revisar criterio IRS.</li>}
                              {s.noAceptados > s.aceptados && s.total > 0 && <li>Alto % no aceptación ({pct(s.noAceptados,s.total)}) — revisar comunicación.</li>}
                              {s.efectividad < 75 && s.total > 0 && <li>Efectividad menor al 75% ({sem.label}) — revisión general.</li>}
                              {s.total === 0        && <li>Sin datos registrados aún.</li>}
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

      </div>

      <div style={{ borderTop:"1px solid #1d2840", padding:"14px 28px", textAlign:"center", color:"#263238", fontSize:11, fontFamily:"Barlow,sans-serif" }}>
        Herramienta de Análisis Arbitral · Liga Señal Colombia de Baloncesto · Uso interno
      </div>
    </div>
  );
}
