import { useState, useMemo, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import { Capacitor } from "@capacitor/core";
import { Filesystem, Directory } from "@capacitor/filesystem";
import { Share } from "@capacitor/share";
import { db, auth, FIREBASE_LISTO } from "./firebase";
import {
  collection, doc, onSnapshot, writeBatch, getDoc, setDoc,
} from "firebase/firestore";
import { signInAnonymously, onAuthStateChanged } from "firebase/auth";
import {
  LayoutDashboard, Package, ArrowDownLeft, ArrowUpRight,
  Users, Truck, FlaskConical, Tag, FileSpreadsheet,
  Plus, Search, Edit2, Trash2, X, AlertTriangle,
  Menu, Upload, Download, TrendingUp, RefreshCw, ChevronDown, ChevronUp
} from "lucide-react";

/* ══════════════════════════════════════════════
   EXPORTAR EXCEL — funciona dentro de la app (APK)
   En Android: guarda el archivo y abre "Compartir/Guardar".
   En navegador: descarga normal.
══════════════════════════════════════════════ */
async function saveWorkbook(wb, filename, showToast){
  try{
    if (Capacitor?.isNativePlatform?.()){
      const b64 = XLSX.write(wb, { type:"base64", bookType:"xlsx" });
      const res = await Filesystem.writeFile({
        path: filename,
        data: b64,
        directory: Directory.Cache,
      });
      await Share.share({
        title: filename,
        text: `Reporte de Sirope: ${filename}`,
        url: res.uri,
        dialogTitle: "Guardar o compartir el Excel",
      });
      showToast("Excel listo ✓");
    } else {
      XLSX.writeFile(wb, filename);
      showToast("Excel descargado ✓");
    }
  }catch(err){
    const msg = (err && err.message) ? err.message : String(err);
    if (/cancel/i.test(msg)) return;          // el usuario cerró el menú de compartir
    showToast("No se pudo exportar: " + msg, "error");
  }
}

/* ══════════════════════════════════════════════
   GRUPOS (11 categorías reales de Sirope)
══════════════════════════════════════════════ */
const GRP = {
  jarabes:      { label:"JARABES",           color:"#B45309", bg:"#FEF3C7", emoji:"🍯",  defUnit:"L"    },
  glitter:      { label:"GLITTER",           color:"#9D174D", bg:"#FCE7F3", emoji:"✨",  defUnit:"pzas" },
  concentrados: { label:"CONCENTRADOS",      color:"#6D28D9", bg:"#EDE9FE", emoji:"🫙",  defUnit:"pzas" },
  neon:         { label:"NEON",              color:"#0369A1", bg:"#E0F2FE", emoji:"💡",  defUnit:"pzas" },
  chile_hot:    { label:"CHILE HOT",         color:"#B91C1C", bg:"#FEE2E2", emoji:"🌶️", defUnit:"pzas" },
  chile_gourmet:{ label:"CHILE HOT GOURMET", color:"#C2410C", bg:"#FFF7ED", emoji:"🔥",  defUnit:"pzas" },
  pami_chela:   { label:"PA'MI CHELA",       color:"#1E40AF", bg:"#DBEAFE", emoji:"🍺",  defUnit:"pzas" },
  herencia_azul:{ label:"HERENCIA AZUL",     color:"#3730A3", bg:"#EEF2FF", emoji:"💙",  defUnit:"pzas" },
  ice_especiales:{ label:"ICE ESPECIALES",   color:"#0E7490", bg:"#CFFAFE", emoji:"🧊",  defUnit:"pzas" },
  base_polvo:   { label:"BASE EN POLVO",     color:"#78350F", bg:"#FEF9C3", emoji:"🥛",  defUnit:"kg"   },
  salsa_choco:  { label:"SALSA CHOCOLATE",   color:"#451A03", bg:"#FDE68A", emoji:"🍫",  defUnit:"pzas" },
};

const UNITS   = ["L","pzas","tambo","kg"];
const SIDEBAR_W = 220;
const BRAND     = "#FF6B35";

/* ══════════════════════════════════════════════
   SINCRONIZACIÓN CON LA NUBE (Firebase Firestore)
   - Tiempo real entre los 10 usuarios
   - Funciona offline y sincroniza al volver el internet
   La función "set" acepta el mismo estilo que useState
   (set(prev => ...)), así las pantallas no cambian.
══════════════════════════════════════════════ */
function useCloud(name){
  const [items,setItems]=useState([]);
  const ref=useRef([]);

  useEffect(()=>{
    if(!FIREBASE_LISTO) return;
    const unsub=onSnapshot(collection(db,name),
      snap=>{
        const arr=snap.docs.map(d=>({id:d.id,...d.data()}));
        ref.current=arr;
        setItems(arr);
      },
      err=>console.error("Firestore("+name+")",err)
    );
    return ()=>unsub();
  },[name]);

  const set=(updater)=>{
    const prev=ref.current;
    const next=typeof updater==="function"?updater(prev):updater;
    const prevMap=new Map(prev.map(i=>[i.id,i]));
    const nextMap=new Map(next.map(i=>[i.id,i]));
    const batch=writeBatch(db);
    // crear / actualizar lo que cambió
    next.forEach(it=>{
      const old=prevMap.get(it.id);
      if(!old || JSON.stringify(old)!==JSON.stringify(it)){
        const {id,...data}=it;
        batch.set(doc(db,name,String(id)),data);
      }
    });
    // borrar lo que ya no está
    prev.forEach(it=>{ if(!nextMap.has(it.id)) batch.delete(doc(db,name,String(it.id))); });
    batch.commit().catch(e=>console.error("sync "+name,e));
    // actualización optimista: la pantalla reacciona al instante
    ref.current=next; setItems(next);
  };

  return [items,set];
}

/* ══════════════════════════════════════════════
   ID FACTORY
══════════════════════════════════════════════ */
let _seq = 0;
const nid = k => `${k}_${Date.now().toString(36)}_${(_seq++).toString(36)}`;

/* ══════════════════════════════════════════════
   PRODUCTOS INICIALES (del Excel Sirope_Almacen)
══════════════════════════════════════════════ */
let _pid = 1;
const mk = (grp, names) => names.map(name => ({
  id: String(_pid++), name: name.trim(), grp,
  stock: 0, min: 5, unit: GRP[grp].defUnit, cost: 0, price: 0,
}));

const PRODS_INIT = [
  ...mk("jarabes",[
    "Algodón de Azúcar Azul","Algodón de Azúcar Rosa","Amaretto","Avellana",
    "Black Berry","Black Cherry","Brandy","Café","Cajeta","Capuchino",
    "Cereza","Chicle Azul","Chicle Rosa","Chocolate Blanco","Chocolate Suizo",
    "Coco","Crema Irlandesa","Curacao","Durazno","Frambuesa","Fresa",
    "Frutas de la Pasión","Frutos Rojos","Grosella","Guaraná","Horchata Clásica",
    "Horchata Coco","Horchata Fresa","Horchata Mazapán","Jamaica","Kiwi",
    "Lavanda","Lichi","Licor de Café","Limón","Limonada Rosa","Malvavisco",
    "Mandarina","Mango","Mantequilla","Manzana Verde","Menta Azul","Menta Blanca",
    "Menta Verde","Mocachino","Mojito","Mora Azul","Mora Blue","Naranja",
    "Nuez","Pepino Limón","Piña","Pitahaya","Rompope","Rosas","Tamarindo",
    "Taro","Té Chai","Té Cítricos","Té Durazno","Té Limón","Té Negro",
    "Té Sandía","Tequila","Uva","Vainilla","Vainilla Francesa","Violeta",
  ]),
  ...mk("glitter",[
    "Algodón de Azúcar","Black Berry","Black Cherry","Bubalo","Cereza",
    "Chicle Azul","Chicle Rosa","Coco","Curacao","Frambuesa","Fresa",
    "Frutas de la Pasión","Frutos Rojos","Guaraná","Kiwi","Limonada",
    "Mango","Manzana Verde","Mojito","Mora Azul","Piña","Tamarindo","Uva",
  ]),
  ...mk("concentrados",[
    "Arándano","Bailyes","Blue Berry","Calabaza y Especias","Cereza","Coco",
    "Durazno","Frambuesa","Fresa","Fresa Coco","Fresa Kiwi","Frutos Rojos",
    "Guanábana","Guayaba","Kiwi","Jamaica con Canela y Anís","Lichi","Limón",
    "Mandarina","Mango","Manzana Verde","Maracuyá","Mazapán","Mojito",
    "Pepino Limón","Pica Fresa","Piña Colada","Pitahaya","Ponche","Rompope",
    "Sandía","Tamarindo","Zarzamora",
  ]),
  ...mk("neon",[
    "Bubalo Azul","Bubalo Morado","Cereza","Frambuesa","Fresa","Frutos Rojos",
    "Guaraná","Limonada Amarilla","Limonada Rosa","Mango","Manzana Verde",
    "Mora Azul","Uva",
  ]),
  ...mk("chile_hot",[
    "Arándano Chipotle","BBQ Clásica","BBQ Diabla","BBQ Hawaiana","BBQ Mezquite",
    "Búfalo","Cajún","Cereza Habanero","Fantasma","Flamin Hot","Frambuesa Morita",
    "Fresa Hot","Frutos Rojos Chipotle","Guayaba Habanero","Jalapeño Hot",
    "Jamaica Chipotle","Lemon Pepper","Mango Guajillo","Mango Habanero",
    "Maracuyá Habanero","Naranja Hot","Original","Pastor","Pelón Pelo Rico",
    "Piña Guajillo","Piña Habanero","Tamarindo 3 Chiles","Tamarindo Chipotle",
    "Tamarindo Habanero","Teriyaki","Zarzamora Chipotle",
  ]),
  ...mk("chile_gourmet",[
    "Arándano Chipotle","Blue Berry","Cereza Cascabel","Extrem","Fresa Tajín",
    "Mango Habanero","Manzana Habanero","Maracuyá Habanero","Naranja Mezcal",
    "Piña Habanero","Piña Mango Habanero","Tamarindo Chipotle","Tamarindo Habanero",
  ]),
  ...mk("pami_chela",[
    "Cereza","Chamoy","Clamato","Cubana","Fresa","Mango","Maracuyá",
    "Mora Azul","Pelón Pelo Rico","Pepino Limón","Piña","Tamarindo",
  ]),
  ...mk("herencia_azul",[
    "Herencia Glitter Frutos Rojos","Herencia Glitter Mango",
    "Herencia Glitter Manzana Verde","Herencia Glitter Mora Azul",
    "Herencia Especial Crema Irlandesa","Herencia Especial Horchata Coco",
    "Licor Agave Reposado","Licor Agave Cristalino","Vodka Biskra",
  ]),
  ...mk("ice_especiales",[
    "Icee Cereza","Icee Frambuesa Azul","Icee Lima Limón","Icee Mango",
    "Icee Manzana Verde","Icee Mora Azul","Icee Pink (Frutos Rojos)",
    "Icee Tutifruti","Icee Uva",
    "Cheetos","Cheetos Fleming","Chips Habanero","Chips Jalapeño",
    "Doritos Incógnito","Doritos Nachos","Hot Nuts","Mangomitas",
    "Pica Fresa","Pulparindo","Rancheritos","Rockaleta","Ruffles",
    "Sabritas Adobadas","Skwinkles","Takis Blue","Takis Fuego","Tutsi Pop",
  ]),
  ...mk("base_polvo",[
    "Capuchino Clásico","Chai Late","Frapuchino Base Cristal",
    "Frapuchino Base Neutra","Frapuchino Caramelo",
    "Frapuchino Chocolate Blanco","Frapuchino Chocolate Oaxaca",
    "Frapuchino Chocolate Obscuro","Frapuchino Chocomenta",
    "Frapuchino Clásico Base Latte","Frapuchino Cookies and Crema",
    "Matcha","Smoothie Fresas con Crema","Smoothie Mazapán","Smoothie Red Velvet",
  ]),
  ...mk("salsa_choco",[
    "Bubulubu","Canastita","Chocomenta","Chocorrol","Clásica",
    "Delice","Ferrero","Gansito",
  ]),
];

/* ══════════════════════════════════════════════
   SAMPLE DATA (entradas, salidas, clientes, proveedores)
══════════════════════════════════════════════ */
const CLIENTS_INIT = [
  {id:"1",name:"María García",     phone:"555-1234",email:"maria@gmail.com",  address:"Col. Centro", notes:"Cliente frecuente"},
  {id:"2",name:"Carlos Eventos MX",phone:"555-5678",email:"carlos@eventos.mx",address:"Col. Roma",   notes:"Pedidos grandes"},
  {id:"3",name:"Dulcería La Abuela",phone:"555-9012",email:"abuela@dulce.com",address:"Mercado #5",  notes:""},
  {id:"4",name:"Cafetería Bloom",  phone:"555-3456",email:"bloom@cafe.mx",    address:"Col. Nápoles",notes:"Pago semanal"},
];
const PROVS_INIT = [
  {id:"1",name:"Distribuidora Dulce+", contact:"Luis Ramos", phone:"555-0001",email:"ventas@dulceplus.mx",  address:"Zona Industrial"},
  {id:"2",name:"Ingredientes Pro MX",  contact:"Ana Torres", phone:"555-0002",email:"ana@ipromx.com",       address:"Col. Industrial"},
  {id:"3",name:"Abastecedora Central", contact:"Pedro Díaz", phone:"555-0003",email:"pedro@abasto.mx",      address:"Bodega Central"},
  {id:"4",name:"GourmetSupply",        contact:"Sara Núñez", phone:"555-0004",email:"sara@gourmetsupply.com",address:"Parque Empresarial"},
];
const ENTS_INIT  = [];
const SALS_INIT  = [];
const PROD_INIT  = [];

/* ══════════════════════════════════════════════
   HELPERS
══════════════════════════════════════════════ */
const fmt  = n => `$${Number(n).toLocaleString("es-MX",{minimumFractionDigits:2,maximumFractionDigits:2})}`;
const fmtD = d => { try{ return new Date(d+"T12:00:00").toLocaleDateString("es-MX"); }catch{ return d; }};
const tdayStr = () => new Date().toISOString().split("T")[0];
const stockSt = (s,m) => {
  if(s===0)  return{label:"Sin stock",color:"#DC2626",bg:"#FEE2E2"};
  if(s<m)    return{label:"Stock bajo",color:"#D97706",bg:"#FEF3C7"};
  return         {label:"OK",         color:"#059669",bg:"#D1FAE5"};
};

/* ══════════════════════════════════════════════
   BASE COMPONENTS
══════════════════════════════════════════════ */
const GrpBadge = ({grp}) => {
  const g=GRP[grp]; if(!g) return null;
  return(
    <span style={{background:g.bg,color:g.color,fontSize:10,fontWeight:800,
      padding:"2px 8px",borderRadius:99,whiteSpace:"nowrap",display:"inline-flex",
      alignItems:"center",gap:3}}>{g.emoji} {g.label}</span>
  );
};
const StockBadge=({stock,min})=>{
  const s=stockSt(stock,min);
  return (
    <span style={{background:s.bg,color:s.color,fontSize:10,fontWeight:800,
      padding:"2px 8px",borderRadius:99,whiteSpace:"nowrap"}}>{s.label}</span>
  );
};

const Modal=({title,onClose,children,wide})=>(
  <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.6)",zIndex:1000,
    display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
    <div style={{background:"#fff",borderRadius:16,width:"100%",maxWidth:wide?680:520,
      maxHeight:"92vh",overflow:"auto",boxShadow:"0 25px 60px rgba(0,0,0,.3)"}}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",
        padding:"18px 22px",borderBottom:"1px solid #F1F5F9",
        position:"sticky",top:0,background:"#fff",zIndex:1}}>
        <h3 style={{margin:0,fontSize:16,fontWeight:900,color:"#1E293B"}}>{title}</h3>
        <button onClick={onClose} style={{background:"#F1F5F9",border:"none",borderRadius:8,
          width:30,height:30,cursor:"pointer",display:"flex",alignItems:"center",
          justifyContent:"center",color:"#64748B"}}><X size={15}/></button>
      </div>
      <div style={{padding:22}}>{children}</div>
    </div>
  </div>
);

const Confirm=({msg,onYes,onNo})=>(
  <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.5)",zIndex:2000,
    display:"flex",alignItems:"center",justifyContent:"center"}}>
    <div style={{background:"#fff",borderRadius:14,padding:28,maxWidth:320,textAlign:"center",
      boxShadow:"0 20px 50px rgba(0,0,0,.3)"}}>
      <div style={{fontSize:28,marginBottom:10}}>🗑️</div>
      <p style={{margin:"0 0 20px",fontSize:14,color:"#374151",fontWeight:600}}>{msg}</p>
      <div style={{display:"flex",gap:10,justifyContent:"center"}}>
        <Btn variant="secondary" onClick={onNo}>Cancelar</Btn>
        <Btn variant="danger" onClick={onYes}>Eliminar</Btn>
      </div>
    </div>
  </div>
);

const Toast=({msg,type,onClose})=>(
  <div style={{position:"fixed",bottom:20,right:20,zIndex:9999,
    background:type==="error"?"#DC2626":"#059669",color:"#fff",
    padding:"12px 18px",borderRadius:10,fontWeight:700,fontSize:13,
    boxShadow:"0 8px 30px rgba(0,0,0,.25)",display:"flex",alignItems:"center",gap:10}}>
    {type==="error"?"⚠️":"✅"} {msg}
    <button onClick={onClose} style={{background:"rgba(255,255,255,.2)",border:"none",
      borderRadius:6,color:"#fff",cursor:"pointer",padding:"2px 7px",fontWeight:900}}>×</button>
  </div>
);

const Field=({label,children,half})=>(
  <div style={{marginBottom:13,flex:half?"1 1 45%":"1 1 100%"}}>
    <label style={{display:"block",fontSize:10,fontWeight:800,color:"#94A3B8",
      marginBottom:5,textTransform:"uppercase",letterSpacing:.7}}>{label}</label>
    {children}
  </div>
);

const inp={width:"100%",padding:"9px 12px",border:"1.5px solid #E2E8F0",borderRadius:8,
  fontSize:13,color:"#1E293B",outline:"none",boxSizing:"border-box",background:"#fff"};
const FInput =(p)=><input  {...p} style={{...inp,...p.style}}/>;
const FSelect=({children,...p})=><select {...p} style={{...inp,...p.style}}>{children}</select>;
const FArea  =(p)=><textarea {...p} style={{...inp,resize:"vertical",minHeight:60,...p.style}}/>;
const FormRow=({children})=><div style={{display:"flex",gap:12,flexWrap:"wrap"}}>{children}</div>;

const Btn=({children,onClick,variant="primary",small,style:ext={},type="button",disabled})=>{
  const base={border:"none",borderRadius:8,cursor:disabled?"not-allowed":"pointer",fontWeight:700,
    display:"inline-flex",alignItems:"center",gap:6,
    padding:small?"5px 10px":"9px 16px",fontSize:small?11:13,
    opacity:disabled?.6:1,transition:"opacity .15s"};
  const V={
    primary:  {background:BRAND,color:"#fff"},
    secondary:{background:"#F1F5F9",color:"#475569"},
    danger:   {background:"#FEE2E2",color:"#DC2626"},
    success:  {background:"#D1FAE5",color:"#059669"},
    ghost:    {background:"transparent",color:"#64748B"},
  };
  return(
    <button type={type} onClick={!disabled?onClick:undefined} disabled={disabled}
      style={{...base,...V[variant],...ext}}
      onMouseEnter={e=>!disabled&&(e.currentTarget.style.opacity=".8")}
      onMouseLeave={e=>e.currentTarget.style.opacity="1"}>{children}</button>
  );
};

const StatCard=({label,value,sub,icon,color})=>(
  <div style={{background:"#fff",borderRadius:14,padding:"18px 20px",flex:1,minWidth:145,
    boxShadow:"0 1px 3px rgba(0,0,0,.07)"}}>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
      <div>
        <p style={{margin:0,fontSize:10,fontWeight:800,color:"#94A3B8",
          textTransform:"uppercase",letterSpacing:.7}}>{label}</p>
        <p style={{margin:"7px 0 3px",fontSize:24,fontWeight:900,color:"#1E293B",lineHeight:1}}>{value}</p>
        {sub&&<p style={{margin:0,fontSize:11,color:"#94A3B8"}}>{sub}</p>}
      </div>
      <div style={{background:color+"22",borderRadius:10,padding:10,color,display:"flex"}}>{icon}</div>
    </div>
  </div>
);

const SearchBar=({value,onChange,placeholder="Buscar..."})=>(
  <div style={{position:"relative",flex:1,minWidth:180}}>
    <Search size={14} style={{position:"absolute",left:11,top:"50%",
      transform:"translateY(-50%)",color:"#CBD5E1"}}/>
    <input value={value} onChange={e=>onChange(e.target.value)} placeholder={placeholder}
      style={{...inp,paddingLeft:32}}/>
  </div>
);

const Card=({children,style})=>(
  <div style={{background:"#fff",borderRadius:14,boxShadow:"0 1px 3px rgba(0,0,0,.07)",
    overflow:"hidden",...style}}>{children}</div>
);

const Tbl=({cols,children})=>(
  <div style={{overflowX:"auto"}}>
    <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
      <thead>
        <tr style={{background:"#F8FAFC"}}>
          {cols.map((c,i)=>(
            <th key={i} style={{padding:"9px 13px",textAlign:"left",fontWeight:800,
              color:"#94A3B8",fontSize:9,textTransform:"uppercase",letterSpacing:.8,
              borderBottom:"2px solid #F1F5F9",whiteSpace:"nowrap"}}>{c}</th>
          ))}
        </tr>
      </thead>
      <tbody>{children}</tbody>
    </table>
  </div>
);
const Tr=({children})=>(
  <tr style={{borderBottom:"1px solid #F8FAFC",transition:"background .1s"}}
    onMouseEnter={e=>e.currentTarget.style.background="#F8FAFC"}
    onMouseLeave={e=>e.currentTarget.style.background="transparent"}>{children}</tr>
);
const Td=({children,style})=>(
  <td style={{padding:"10px 13px",color:"#374151",verticalAlign:"middle",...style}}>{children}</td>
);
const EmptyRow=({cols,msg="Sin registros"})=>(
  <tr><td colSpan={cols} style={{textAlign:"center",padding:"36px 20px",color:"#CBD5E1",fontSize:13}}>{msg}</td></tr>
);

/* ══════════════════════════════════════════════
   GROUP CARD (clickable panel)
══════════════════════════════════════════════ */
const GrpCard=({grpKey,grp,count,totalStock,selected,onClick})=>(
  <div onClick={onClick} style={{
    background:selected?grp.bg:"#fff",
    border:`2px solid ${selected?grp.color:"#E2E8F0"}`,
    borderRadius:12,padding:"12px 14px",cursor:"pointer",
    transition:"all .15s",minWidth:0,
    boxShadow:selected?`0 0 0 3px ${grp.color}22`:"0 1px 3px rgba(0,0,0,.05)",
  }}>
    <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:6}}>
      <span style={{fontSize:18}}>{grp.emoji}</span>
      <span style={{fontSize:10,fontWeight:800,color:grp.color,
        textTransform:"uppercase",letterSpacing:.5,lineHeight:1.2}}>{grp.label}</span>
    </div>
    <p style={{margin:0,fontSize:20,fontWeight:900,color:"#1E293B",lineHeight:1}}>
      {count} <span style={{fontSize:10,fontWeight:500,color:"#94A3B8"}}>productos</span>
    </p>
    {totalStock>0&&(
      <p style={{margin:"3px 0 0",fontSize:10,color:grp.color,fontWeight:700}}>
        Stock: {totalStock}
      </p>
    )}
  </div>
);

/* ══════════════════════════════════════════════
   DASHBOARD
══════════════════════════════════════════════ */
function Dashboard({products,entries,exits,setView}){
  const alerts=products.filter(p=>p.stock<p.min);
  const alertsShown=alerts.slice(0,12);
  const totalVal=products.reduce((s,p)=>s+p.stock*p.cost,0);
  const totalSales=exits.reduce((s,e)=>s+e.qty*e.price,0);
  const byGrp={};
  Object.keys(GRP).forEach(k=>{byGrp[k]=products.filter(p=>p.grp===k).length});

  return(
    <div>
      {/* BRAND HERO */}
      <div style={{background:"#0D1629",borderRadius:16,padding:"22px 24px",
        marginBottom:22,display:"flex",alignItems:"center",gap:20,flexWrap:"wrap",
        boxShadow:"0 4px 20px rgba(13,22,41,.4)"}}>
        <div style={{flex:1}}>
          <p style={{margin:0,fontSize:11,fontWeight:700,color:"rgba(255,255,255,.4)",
            textTransform:"uppercase",letterSpacing:2,marginBottom:4}}>Sistema de Inventario</p>
          <div style={{fontFamily:"'Georgia',serif",fontSize:42,fontWeight:900,
            color:"#fff",lineHeight:1,letterSpacing:-1,
            textShadow:`0 0 30px ${BRAND}88`}}>
            <span style={{fontStyle:"italic",fontFamily:"'Georgia',cursive"}}>Sirope</span>
          </div>
          <p style={{margin:"6px 0 0",fontSize:12,color:"rgba(255,255,255,.45)"}}>
            {Object.keys(GRP).length} líneas · {products.length} productos en catálogo
          </p>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:6}}>
          {Object.entries(GRP).slice(0,6).map(([k,g])=>(
            <div key={k} style={{background:`${g.color}22`,borderRadius:8,padding:"5px 8px",
              textAlign:"center"}}>
              <span style={{fontSize:16}}>{g.emoji}</span>
              <p style={{margin:0,fontSize:8,color:g.color,fontWeight:800,
                textTransform:"uppercase",letterSpacing:.3}}>{g.label.split(" ")[0]}</p>
            </div>
          ))}
        </div>
      </div>

      <div style={{display:"flex",gap:12,flexWrap:"wrap",marginBottom:20}}>
        <StatCard label="Productos"   value={products.length} sub="en catálogo"
          icon={<Package size={18}/>} color={BRAND}/>
        <StatCard label="Líneas"      value={Object.keys(GRP).length} sub="grupos activos"
          icon={<TrendingUp size={18}/>} color="#6D28D9"/>
        <StatCard label="Alertas"     value={alerts.length} sub="stock bajo"
          icon={<AlertTriangle size={18}/>} color="#D97706"/>
        <StatCard label="Ventas"      value={fmt(totalSales)} sub="acumulado"
          icon={<RefreshCw size={18}/>} color="#059669"/>
      </div>

      {alerts.length>0&&(
        <div style={{background:"#FFF7ED",border:"1.5px solid #FED7AA",borderRadius:14,
          padding:16,marginBottom:20}}>
          <h3 style={{margin:"0 0 12px",fontSize:13,fontWeight:800,color:"#C2410C",
            display:"flex",alignItems:"center",gap:8}}>
            <AlertTriangle size={15}/> Alertas de Stock ({alerts.length})
          </h3>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(190px,1fr))",gap:8}}>
            {alertsShown.map(p=>{
              const s=stockSt(p.stock,p.min);
              const g=GRP[p.grp];
              return(
                <div key={p.id} style={{background:"#fff",borderRadius:9,padding:"10px 13px",
                  borderLeft:`4px solid ${s.color}`}}>
                  <p style={{margin:"0 0 2px",fontWeight:800,fontSize:12}}>{p.name}</p>
                  <p style={{margin:"0 0 5px",fontSize:11,color:s.color,fontWeight:700}}>
                    {p.stock} {p.unit} / mín {p.min}
                  </p>
                  <GrpBadge grp={p.grp}/>
                </div>
              );
            })}
          </div>
          {alerts.length>12&&(
            <p onClick={()=>setView&&setView("inventario")}
              style={{margin:"12px 0 0",fontSize:12,color:"#C2410C",fontWeight:700,
                cursor:"pointer",textDecoration:"underline"}}>
              + {alerts.length-12} productos más con stock bajo — ver en Inventario
            </p>
          )}
        </div>
      )}

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
        <Card style={{padding:16}}>
          <h3 style={{margin:"0 0 12px",fontSize:13,fontWeight:800,color:"#059669",
            display:"flex",alignItems:"center",gap:7}}><ArrowDownLeft size={15}/> Últimas Entradas</h3>
          {entries.length===0&&<p style={{color:"#CBD5E1",fontSize:12}}>Sin movimientos aún</p>}
          {[...entries].reverse().slice(0,5).map(e=>{
            const p=products.find(x=>x.id===e.prodId);
            return(
              <div key={e.id} style={{display:"flex",justifyContent:"space-between",
                alignItems:"center",padding:"7px 0",borderBottom:"1px solid #F8FAFC"}}>
                <span style={{fontSize:12,fontWeight:600,color:"#374151"}}>{p?.name||"—"}</span>
                <span style={{fontSize:12,fontWeight:800,color:"#059669"}}>+{e.qty} {p?.unit}</span>
              </div>
            );
          })}
        </Card>
        <Card style={{padding:16}}>
          <h3 style={{margin:"0 0 12px",fontSize:13,fontWeight:800,color:"#DC2626",
            display:"flex",alignItems:"center",gap:7}}><ArrowUpRight size={15}/> Últimas Salidas</h3>
          {exits.length===0&&<p style={{color:"#CBD5E1",fontSize:12}}>Sin movimientos aún</p>}
          {[...exits].reverse().slice(0,5).map(s=>{
            const p=products.find(x=>x.id===s.prodId);
            return(
              <div key={s.id} style={{display:"flex",justifyContent:"space-between",
                alignItems:"center",padding:"7px 0",borderBottom:"1px solid #F8FAFC"}}>
                <span style={{fontSize:12,fontWeight:600,color:"#374151"}}>{p?.name||"—"}</span>
                <span style={{fontSize:12,fontWeight:800,color:"#DC2626"}}>-{s.qty} {p?.unit}</span>
              </div>
            );
          })}
        </Card>
      </div>
    </div>
  );
}

/* ══════════════════════════════════════════════
   INVENTARIO VIEW  (grupos clickeables)
══════════════════════════════════════════════ */
function InventarioView({products,setProducts,showToast}){
  const [selGrp,setSelGrp]=useState(null);
  const [search,setSearch]=useState("");
  const [modal,setModal]=useState(null);
  const [form,setForm]=useState({});
  const [confirm,setConfirm]=useState(null);
  const [page,setPage]=useState(0);
  const PAGE=40;

  const f=k=>e=>setForm(prev=>({...prev,[k]:e.target.value}));

  const filtered=useMemo(()=>products.filter(p=>{
    const q=search.trim().toLowerCase();
    const mQ=!q||p.name.toLowerCase().includes(q);
    const mG=!selGrp||p.grp===selGrp;
    return mQ&&mG;
  }),[products,search,selGrp]);

  // Reset to first page whenever the filter changes
  useEffect(()=>{setPage(0)},[selGrp,search]);

  // Only show the product list once a group is picked or a search is typed.
  const showList = !!selGrp || search.trim().length>0;
  const totalPages = Math.max(1, Math.ceil(filtered.length/PAGE));
  const pageItems = filtered.slice(page*PAGE,(page+1)*PAGE);

  const openAdd=()=>{
    const grp=selGrp||"jarabes";
    setForm({name:"",grp,stock:0,min:5,unit:GRP[grp].defUnit,cost:0,price:0});
    setModal("add");
  };
  const openEdit=p=>{setForm({...p});setModal("edit");};

  const save=()=>{
    if(!form.name.trim()) return showToast("El nombre es requerido","error");
    const item={...form,stock:+form.stock,min:+form.min,cost:+form.cost,price:+form.price};
    if(modal==="add"){
      setProducts(prev=>[...prev,{...item,id:nid("p")}]);
      showToast("Producto agregado ✓");
    }else{
      setProducts(prev=>prev.map(p=>p.id===form.id?item:p));
      showToast("Producto actualizado ✓");
    }
    setModal(null);
  };

  const del=id=>{
    setProducts(prev=>prev.filter(p=>p.id!==id));
    showToast("Eliminado");setConfirm(null);
  };

  const grpCounts=useMemo(()=>{
    const r={};
    Object.keys(GRP).forEach(k=>{
      const ps=products.filter(p=>p.grp===k);
      r[k]={count:ps.length,stock:ps.reduce((s,p)=>s+p.stock,0)};
    });
    return r;
  },[products]);

  return(
    <div>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:16}}>
        <div>
          <h2 style={{margin:0,fontSize:19,fontWeight:900,color:"#1E293B"}}>Inventario</h2>
          <p style={{margin:"2px 0 0",fontSize:11,color:"#94A3B8"}}>
            {selGrp?`${GRP[selGrp].label} · `:"Todos los grupos · "}{filtered.length} productos
          </p>
        </div>
        <Btn onClick={openAdd}><Plus size={14}/> Nuevo Producto</Btn>
      </div>

      {/* GROUP GRID */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(145px,1fr))",
        gap:10,marginBottom:18}}>
        {Object.entries(GRP).map(([k,g])=>(
          <GrpCard key={k} grpKey={k} grp={g}
            count={grpCounts[k]?.count||0}
            totalStock={grpCounts[k]?.stock||0}
            selected={selGrp===k}
            onClick={()=>setSelGrp(selGrp===k?null:k)}/>
        ))}
      </div>

      <div style={{display:"flex",gap:10,marginBottom:14}}>
        <SearchBar value={search} onChange={setSearch} placeholder="Buscar producto en el grupo..."/>
        {selGrp&&(
          <Btn variant="secondary" onClick={()=>setSelGrp(null)}>
            <X size={12}/> {GRP[selGrp].label}
          </Btn>
        )}
      </div>

      {!showList ? (
        <Card style={{padding:"40px 24px",textAlign:"center"}}>
          <Package size={32} style={{color:"#CBD5E1",marginBottom:10}}/>
          <p style={{margin:"0 0 4px",fontWeight:800,fontSize:15,color:"#374151"}}>
            Elige un grupo para ver sus productos
          </p>
          <p style={{margin:0,fontSize:12,color:"#94A3B8"}}>
            Toca una de las {Object.keys(GRP).length} líneas de arriba, o busca un producto por nombre.
          </p>
        </Card>
      ) : (
        <Card>
          <Tbl cols={["Producto","Grupo","Stock","Mín.","Unidad","Costo","Precio","Margen","Estado",""]}>
            {pageItems.length===0&&<EmptyRow cols={10} msg="Sin productos para este filtro"/>}
            {pageItems.map(p=>{
              const m=p.cost>0?Math.round(((p.price-p.cost)/p.cost)*100):0;
              return(
                <Tr key={p.id}>
                  <Td style={{fontWeight:700,maxWidth:180}}>{p.name}</Td>
                  <Td><GrpBadge grp={p.grp}/></Td>
                  <Td><span style={{fontWeight:900,fontSize:14}}>{p.stock}</span></Td>
                  <Td style={{color:"#94A3B8"}}>{p.min}</Td>
                  <Td style={{color:"#64748B"}}>{p.unit}</Td>
                  <Td>{fmt(p.cost)}</Td>
                  <Td style={{fontWeight:700,color:BRAND}}>{fmt(p.price)}</Td>
                  <Td><span style={{color:m>=30?"#059669":m>=10?"#D97706":"#94A3B8",fontWeight:700}}>{m}%</span></Td>
                  <Td><StockBadge stock={p.stock} min={p.min}/></Td>
                  <Td>
                    <div style={{display:"flex",gap:4}}>
                      <Btn small variant="secondary" onClick={()=>openEdit(p)}><Edit2 size={11}/></Btn>
                      <Btn small variant="danger" onClick={()=>setConfirm(p.id)}><Trash2 size={11}/></Btn>
                    </div>
                  </Td>
                </Tr>
              );
            })}
          </Tbl>

          {filtered.length>PAGE&&(
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",
              padding:"12px 16px",borderTop:"1px solid #F1F5F9",flexWrap:"wrap",gap:8}}>
              <span style={{fontSize:12,color:"#94A3B8",fontWeight:600}}>
                {page*PAGE+1}–{Math.min((page+1)*PAGE,filtered.length)} de {filtered.length}
              </span>
              <div style={{display:"flex",gap:6}}>
                <Btn small variant="secondary" disabled={page===0}
                  onClick={()=>setPage(p=>Math.max(0,p-1))}>Anterior</Btn>
                <span style={{fontSize:12,fontWeight:700,color:"#475569",
                  alignSelf:"center",padding:"0 4px"}}>{page+1} / {totalPages}</span>
                <Btn small variant="secondary" disabled={page>=totalPages-1}
                  onClick={()=>setPage(p=>Math.min(totalPages-1,p+1))}>Siguiente</Btn>
              </div>
            </div>
          )}
        </Card>
      )}

      {modal&&(
        <Modal title={modal==="add"?"Nuevo Producto":"Editar Producto"} onClose={()=>setModal(null)}>
          <Field label="Nombre del producto">
            <FInput value={form.name} onChange={f("name")} placeholder="ej. Fresa"/>
          </Field>
          <FormRow>
            <Field label="Grupo / Línea" half>
              <FSelect value={form.grp} onChange={e=>{
                setForm(prev=>({...prev,grp:e.target.value,unit:GRP[e.target.value].defUnit}));
              }}>
                {Object.entries(GRP).map(([k,g])=><option key={k} value={k}>{g.emoji} {g.label}</option>)}
              </FSelect>
            </Field>
            <Field label="Unidad de medida" half>
              <FSelect value={form.unit} onChange={f("unit")}>
                {UNITS.map(u=><option key={u} value={u}>{u}</option>)}
              </FSelect>
            </Field>
            <Field label="Stock actual" half>
              <FInput type="number" value={form.stock} onChange={f("stock")} min={0}/>
            </Field>
            <Field label="Stock mínimo" half>
              <FInput type="number" value={form.min} onChange={f("min")} min={0}/>
            </Field>
            <Field label="Costo ($)" half>
              <FInput type="number" value={form.cost} onChange={f("cost")} min={0}/>
            </Field>
            <Field label="Precio venta ($)" half>
              <FInput type="number" value={form.price} onChange={f("price")} min={0}/>
            </Field>
          </FormRow>
          <div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:6}}>
            <Btn variant="secondary" onClick={()=>setModal(null)}>Cancelar</Btn>
            <Btn onClick={save}>Guardar</Btn>
          </div>
        </Modal>
      )}
      {confirm&&<Confirm msg="¿Eliminar este producto?" onYes={()=>del(confirm)} onNo={()=>setConfirm(null)}/>}
    </div>
  );
}

/* ══════════════════════════════════════════════
   BUSCADOR DE PRODUCTO  (reemplaza el <select> de 250
   opciones que se atora en Android). Permite teclear,
   filtrar y CREAR un producto nuevo sobre la marcha.
══════════════════════════════════════════════ */
function createProductInline(setProducts, showToast, grp, rawName){
  const g = GRP[grp] ? grp : "jarabes";
  const np = { id:nid("p"), name:String(rawName).trim(), grp:g,
    stock:0, min:5, unit:GRP[g].defUnit, cost:0, price:0 };
  setProducts(prev=>[...prev, np]);
  showToast(`Producto «${np.name}» creado en ${GRP[g].label}`);
  return np.id;
}

function ProductPicker({products, value, onChange, grp, onCreate, placeholder}){
  const [open,setOpen]=useState(false);
  const [q,setQ]=useState("");
  const selected=products.find(p=>p.id===value);

  const list=useMemo(()=>{
    const qq=q.trim().toLowerCase();
    return products
      .filter(p=>!grp||p.grp===grp)
      .filter(p=>!qq||p.name.toLowerCase().includes(qq))
      .slice(0,80);
  },[products,grp,q]);

  const qq=q.trim();
  const exact=products.some(p=>p.name.toLowerCase()===qq.toLowerCase() && (!grp||p.grp===grp));

  const pick=(id)=>{ onChange(id); setOpen(false); setQ(""); };
  const create=()=>{ if(!qq||!onCreate) return; const id=onCreate(qq); if(id) pick(id); };

  return(
    <div>
      {/* Caja con la selección actual */}
      <div onClick={()=>setOpen(o=>!o)}
        style={{...inp, display:"flex", alignItems:"center", justifyContent:"space-between",
          cursor:"pointer", gap:8}}>
        <span style={{color:selected?"#1E293B":"#94A3B8", overflow:"hidden",
          whiteSpace:"nowrap", textOverflow:"ellipsis", fontSize:13}}>
          {selected ? `${GRP[selected.grp]?.emoji||""} ${selected.name}`
                    : (placeholder||"Toca para buscar un producto")}
        </span>
        <Search size={14} style={{color:"#94A3B8",flexShrink:0}}/>
      </div>

      {open&&(
        <div style={{marginTop:6, border:"1.5px solid #E2E8F0", borderRadius:10,
          overflow:"hidden", background:"#fff", boxShadow:"0 8px 24px rgba(0,0,0,.10)"}}>
          <div style={{padding:8, borderBottom:"1px solid #F1F5F9"}}>
            <input autoFocus value={q} onChange={e=>setQ(e.target.value)}
              placeholder="Escribe el nombre…"
              style={{...inp, padding:"8px 10px"}}/>
          </div>
          <div style={{maxHeight:240, overflowY:"auto"}}>
            {list.length===0 && !qq && (
              <div style={{padding:14, fontSize:12, color:"#94A3B8", textAlign:"center"}}>
                Escribe para buscar tu producto…
              </div>
            )}
            {list.map(p=>(
              <div key={p.id} onClick={()=>pick(p.id)}
                style={{padding:"10px 12px", cursor:"pointer", display:"flex",
                  justifyContent:"space-between", gap:8, fontSize:13,
                  borderBottom:"1px solid #F8FAFC",
                  background:p.id===value?"#FFF7ED":"#fff"}}>
                <span style={{overflow:"hidden",whiteSpace:"nowrap",textOverflow:"ellipsis"}}>
                  {GRP[p.grp]?.emoji} {p.name}
                </span>
                <span style={{color:"#94A3B8",flexShrink:0,fontSize:11}}>{p.stock} {p.unit}</span>
              </div>
            ))}
            {qq && !exact && onCreate && (
              <div onClick={create}
                style={{padding:"12px", cursor:"pointer", display:"flex",
                  alignItems:"center", gap:8, fontSize:13, fontWeight:800,
                  color:BRAND, background:"#FFF7ED"}}>
                <Plus size={15}/> Crear nuevo producto: «{qq}»
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

/* ══════════════════════════════════════════════
   ENTRADAS VIEW
══════════════════════════════════════════════ */
function EntradasView({entries,setEntries,products,setProducts,suppliers,showToast}){
  const [modal,setModal]=useState(false);
  const [form,setForm]=useState({});
  const [search,setSearch]=useState("");
  const [selGrp,setSelGrp]=useState(null);
  const f=k=>e=>setForm(prev=>({...prev,[k]:e.target.value}));

  const filtProds=selGrp?products.filter(p=>p.grp===selGrp):products;

  const openAdd=()=>{
    setForm({date:tdayStr(),prodId:"",qty:1,cost:"",provId:suppliers[0]?.id||"",notes:""});
    setModal(true);
  };
  const save=()=>{
    if(!form.prodId) return showToast("Selecciona un producto","error");
    if(+form.qty<=0) return showToast("Cantidad inválida","error");
    setEntries(prev=>[...prev,{...form,id:nid("ent"),qty:+form.qty,cost:+form.cost}]);
    setProducts(prev=>prev.map(p=>p.id===form.prodId?{...p,stock:p.stock+(+form.qty)}:p));
    showToast(`+${form.qty} unidades registradas`);setModal(false);
  };

  const filtered=useMemo(()=>entries.filter(e=>{
    const p=products.find(x=>x.id===e.prodId);
    const mS=!search||p?.name.toLowerCase().includes(search.toLowerCase());
    const mG=!selGrp||p?.grp===selGrp;
    return mS&&mG;
  }),[entries,products,search,selGrp]);

  return(
    <div>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14}}>
        <div>
          <h2 style={{margin:0,fontSize:19,fontWeight:900,color:"#1E293B"}}>Entradas</h2>
          <p style={{margin:"2px 0 0",fontSize:11,color:"#94A3B8"}}>Recepciones de mercancía</p>
        </div>
        <Btn onClick={openAdd}><Plus size={14}/> Registrar Entrada</Btn>
      </div>

      <div style={{display:"flex",gap:8,marginBottom:14,flexWrap:"wrap"}}>
        <SearchBar value={search} onChange={setSearch} placeholder="Buscar producto..."/>
        <FSelect value={selGrp||""} onChange={e=>setSelGrp(e.target.value||null)} style={{width:"auto",minWidth:180}}>
          <option value="">Todos los grupos</option>
          {Object.entries(GRP).map(([k,g])=><option key={k} value={k}>{g.emoji} {g.label}</option>)}
        </FSelect>
      </div>

      <Card>
        <Tbl cols={["Fecha","Producto","Grupo","Cantidad","Costo Unit.","Total","Proveedor","Notas"]}>
          {filtered.length===0&&<EmptyRow cols={8} msg="Sin entradas registradas"/>}
          {[...filtered].reverse().map(e=>{
            const p=products.find(x=>x.id===e.prodId);
            const prov=suppliers.find(x=>x.id===e.provId);
            return(
              <Tr key={e.id}>
                <Td style={{color:"#64748B",whiteSpace:"nowrap"}}>{fmtD(e.date)}</Td>
                <Td style={{fontWeight:700}}>{p?.name||"—"}</Td>
                <Td><GrpBadge grp={p?.grp}/></Td>
                <Td><span style={{fontWeight:800,color:"#059669"}}>+{e.qty}</span> <span style={{color:"#94A3B8",fontSize:11}}>{p?.unit}</span></Td>
                <Td>{e.cost?fmt(e.cost):"—"}</Td>
                <Td style={{fontWeight:700}}>{e.cost?fmt(e.qty*e.cost):"—"}</Td>
                <Td style={{color:"#64748B",fontSize:12}}>{prov?.name||"—"}</Td>
                <Td style={{color:"#94A3B8",fontSize:11}}>{e.notes||"—"}</Td>
              </Tr>
            );
          })}
        </Tbl>
      </Card>

      {modal&&(
        <Modal title="Registrar Entrada" onClose={()=>setModal(false)}>
          <FormRow>
            <Field label="Fecha" half><FInput type="date" value={form.date} onChange={f("date")}/></Field>
            <Field label="Proveedor" half>
              <FSelect value={form.provId} onChange={f("provId")}>
                <option value="">Sin proveedor</option>
                {suppliers.map(s=><option key={s.id} value={s.id}>{s.name}</option>)}
              </FSelect>
            </Field>
          </FormRow>
          <Field label="Grupo / Línea">
            <FSelect value={selGrp||""} onChange={e=>setSelGrp(e.target.value||null)}
              style={{marginBottom:8}}>
              <option value="">Todos los grupos</option>
              {Object.entries(GRP).map(([k,g])=><option key={k} value={k}>{g.emoji} {g.label}</option>)}
            </FSelect>
          </Field>
          <Field label="Producto">
            <ProductPicker products={products} value={form.prodId} grp={selGrp}
              onChange={id=>setForm(prev=>({...prev,prodId:id}))}
              onCreate={name=>createProductInline(setProducts,showToast,selGrp,name)}/>
          </Field>
          <FormRow>
            <Field label="Cantidad" half><FInput type="number" value={form.qty} onChange={f("qty")} min={1}/></Field>
            <Field label="Costo unitario ($)" half><FInput type="number" value={form.cost} onChange={f("cost")} min={0} placeholder="0.00"/></Field>
          </FormRow>
          <Field label="Notas"><FArea value={form.notes} onChange={f("notes")} placeholder="Info adicional..."/></Field>
          <div style={{display:"flex",gap:10,justifyContent:"flex-end"}}>
            <Btn variant="secondary" onClick={()=>setModal(false)}>Cancelar</Btn>
            <Btn onClick={save} variant="success"><ArrowDownLeft size={13}/> Confirmar</Btn>
          </div>
        </Modal>
      )}
    </div>
  );
}

/* ══════════════════════════════════════════════
   SALIDAS VIEW
══════════════════════════════════════════════ */
function SalidasView({exits,setExits,products,setProducts,clients,showToast}){
  const [modal,setModal]=useState(false);
  const [form,setForm]=useState({});
  const [search,setSearch]=useState("");
  const [selGrp,setSelGrp]=useState(null);
  const f=k=>e=>setForm(prev=>({...prev,[k]:e.target.value}));

  const openAdd=()=>{
    setForm({date:tdayStr(),prodId:"",qty:1,price:"",cliId:clients[0]?.id||"",notes:""});
    setModal(true);
  };
  const save=()=>{
    if(!form.prodId) return showToast("Selecciona un producto","error");
    const prod=products.find(p=>p.id===form.prodId);
    if(!prod) return showToast("Producto no encontrado","error");
    if(+form.qty<=0) return showToast("Cantidad inválida","error");
    if(prod.stock<+form.qty) return showToast(`Stock insuficiente. Disponible: ${prod.stock} ${prod.unit}`,"error");
    setExits(prev=>[...prev,{...form,id:nid("sal"),qty:+form.qty,price:+form.price}]);
    setProducts(prev=>prev.map(p=>p.id===form.prodId?{...p,stock:p.stock-(+form.qty)}:p));
    showToast(`-${form.qty} unidades registradas`);setModal(false);
  };

  const filtered=useMemo(()=>exits.filter(e=>{
    const p=products.find(x=>x.id===e.prodId);
    const mS=!search||p?.name.toLowerCase().includes(search.toLowerCase());
    const mG=!selGrp||p?.grp===selGrp;
    return mS&&mG;
  }),[exits,products,search,selGrp]);

  return(
    <div>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14}}>
        <div>
          <h2 style={{margin:0,fontSize:19,fontWeight:900,color:"#1E293B"}}>Salidas</h2>
          <p style={{margin:"2px 0 0",fontSize:11,color:"#94A3B8"}}>Ventas y consumos</p>
        </div>
        <Btn onClick={openAdd}><Plus size={14}/> Registrar Salida</Btn>
      </div>
      <div style={{display:"flex",gap:8,marginBottom:14,flexWrap:"wrap"}}>
        <SearchBar value={search} onChange={setSearch} placeholder="Buscar producto..."/>
        <FSelect value={selGrp||""} onChange={e=>setSelGrp(e.target.value||null)} style={{width:"auto",minWidth:180}}>
          <option value="">Todos los grupos</option>
          {Object.entries(GRP).map(([k,g])=><option key={k} value={k}>{g.emoji} {g.label}</option>)}
        </FSelect>
      </div>
      <Card>
        <Tbl cols={["Fecha","Producto","Grupo","Cantidad","Precio Unit.","Total","Cliente","Notas"]}>
          {filtered.length===0&&<EmptyRow cols={8} msg="Sin salidas registradas"/>}
          {[...filtered].reverse().map(e=>{
            const p=products.find(x=>x.id===e.prodId);
            const cli=clients.find(x=>x.id===e.cliId);
            return(
              <Tr key={e.id}>
                <Td style={{color:"#64748B",whiteSpace:"nowrap"}}>{fmtD(e.date)}</Td>
                <Td style={{fontWeight:700}}>{p?.name||"—"}</Td>
                <Td><GrpBadge grp={p?.grp}/></Td>
                <Td><span style={{fontWeight:800,color:"#DC2626"}}>-{e.qty}</span> <span style={{color:"#94A3B8",fontSize:11}}>{p?.unit}</span></Td>
                <Td>{e.price?fmt(e.price):"—"}</Td>
                <Td style={{fontWeight:700,color:BRAND}}>{e.price?fmt(e.qty*e.price):"—"}</Td>
                <Td style={{color:"#64748B",fontSize:12}}>{cli?.name||"—"}</Td>
                <Td style={{color:"#94A3B8",fontSize:11}}>{e.notes||"—"}</Td>
              </Tr>
            );
          })}
        </Tbl>
      </Card>

      {modal&&(
        <Modal title="Registrar Salida" onClose={()=>setModal(false)}>
          <FormRow>
            <Field label="Fecha" half><FInput type="date" value={form.date} onChange={f("date")}/></Field>
            <Field label="Cliente" half>
              <FSelect value={form.cliId} onChange={f("cliId")}>
                <option value="">Sin cliente</option>
                {clients.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}
              </FSelect>
            </Field>
          </FormRow>
          <Field label="Grupo / Línea">
            <FSelect value={selGrp||""} onChange={e=>setSelGrp(e.target.value||null)} style={{marginBottom:8}}>
              <option value="">Todos los grupos</option>
              {Object.entries(GRP).map(([k,g])=><option key={k} value={k}>{g.emoji} {g.label}</option>)}
            </FSelect>
          </Field>
          <Field label="Producto">
            <ProductPicker products={products} value={form.prodId} grp={selGrp}
              onChange={id=>setForm(prev=>({...prev,prodId:id}))}
              onCreate={name=>createProductInline(setProducts,showToast,selGrp,name)}/>
          </Field>
          <FormRow>
            <Field label="Cantidad" half><FInput type="number" value={form.qty} onChange={f("qty")} min={1}/></Field>
            <Field label="Precio unitario ($)" half><FInput type="number" value={form.price} onChange={f("price")} min={0} placeholder="0.00"/></Field>
          </FormRow>
          <Field label="Notas"><FArea value={form.notes} onChange={f("notes")} placeholder="Info adicional..."/></Field>
          <div style={{display:"flex",gap:10,justifyContent:"flex-end"}}>
            <Btn variant="secondary" onClick={()=>setModal(false)}>Cancelar</Btn>
            <Btn onClick={save} variant="danger"><ArrowUpRight size={13}/> Confirmar</Btn>
          </div>
        </Modal>
      )}
    </div>
  );
}

/* ══════════════════════════════════════════════
   CRUD VIEW (Clientes / Proveedores)
══════════════════════════════════════════════ */
function CRUDView({title,items,setItems,fields,idKey,showToast}){
  const [search,setSearch]=useState("");
  const [modal,setModal]=useState(null);
  const [form,setForm]=useState({});
  const [confirm,setConfirm]=useState(null);
  const f=k=>e=>setForm(prev=>({...prev,[k]:e.target.value}));

  const filtered=useMemo(()=>items.filter(i=>{
    const q=search.toLowerCase();
    return fields.some(fl=>String(i[fl.key]||"").toLowerCase().includes(q));
  }),[items,search,fields]);

  const openAdd=()=>{const b={};fields.forEach(fl=>b[fl.key]="");setForm(b);setModal("add");};
  const openEdit=i=>{setForm({...i});setModal("edit");};

  const save=()=>{
    const req=fields.find(fl=>fl.required&&!form[fl.key]?.trim());
    if(req) return showToast(`${req.label} es requerido`,"error");
    if(modal==="add"){setItems(prev=>[...prev,{...form,id:nid(idKey)}]);showToast("Agregado ✓");}
    else{setItems(prev=>prev.map(i=>i.id===form.id?{...form}:i));showToast("Actualizado ✓");}
    setModal(null);
  };
  const del=id=>{setItems(prev=>prev.filter(i=>i.id!==id));showToast("Eliminado");setConfirm(null);};

  return(
    <div>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14}}>
        <h2 style={{margin:0,fontSize:19,fontWeight:900,color:"#1E293B"}}>{title}</h2>
        <Btn onClick={openAdd}><Plus size={14}/> Agregar</Btn>
      </div>
      <div style={{marginBottom:14}}><SearchBar value={search} onChange={setSearch} placeholder={`Buscar ${title.toLowerCase()}...`}/></div>
      <Card>
        <Tbl cols={[...fields.map(fl=>fl.label),""]}>
          {filtered.length===0&&<EmptyRow cols={fields.length+1}/>}
          {filtered.map(item=>(
            <Tr key={item.id}>
              {fields.map(fl=>(
                <Td key={fl.key} style={fl.bold?{fontWeight:700}:{}}>{item[fl.key]||"—"}</Td>
              ))}
              <Td>
                <div style={{display:"flex",gap:4}}>
                  <Btn small variant="secondary" onClick={()=>openEdit(item)}><Edit2 size={11}/></Btn>
                  <Btn small variant="danger" onClick={()=>setConfirm(item.id)}><Trash2 size={11}/></Btn>
                </div>
              </Td>
            </Tr>
          ))}
        </Tbl>
      </Card>
      {modal&&(
        <Modal title={modal==="add"?`Nuevo ${title.slice(0,-1)}`:`Editar ${title.slice(0,-1)}`} onClose={()=>setModal(null)}>
          <FormRow>
            {fields.map(fl=>(
              <Field key={fl.key} label={fl.label} half={fl.half}>
                {fl.area
                  ?<FArea value={form[fl.key]||""} onChange={f(fl.key)} placeholder={fl.ph||""}/>
                  :<FInput type={fl.type||"text"} value={form[fl.key]||""} onChange={f(fl.key)} placeholder={fl.ph||""}/>
                }
              </Field>
            ))}
          </FormRow>
          <div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:6}}>
            <Btn variant="secondary" onClick={()=>setModal(null)}>Cancelar</Btn>
            <Btn onClick={save}>Guardar</Btn>
          </div>
        </Modal>
      )}
      {confirm&&<Confirm msg={`¿Eliminar este ${title.slice(0,-1).toLowerCase()}?`} onYes={()=>del(confirm)} onNo={()=>setConfirm(null)}/>}
    </div>
  );
}

/* ══════════════════════════════════════════════
   PRODUCCIÓN VIEW
══════════════════════════════════════════════ */
function ProduccionView({prodLogs,setProdLogs,products,setProducts,showToast}){
  const [selGrp,setSelGrp]=useState(null);
  const [modal,setModal]=useState(false);
  const [form,setForm]=useState({});
  const f=k=>e=>setForm(prev=>({...prev,[k]:e.target.value}));

  const filtProds=selGrp?products.filter(p=>p.grp===selGrp):products;

  const openAdd=()=>{
    setForm({date:tdayStr(),grp:selGrp||"jarabes",prodId:"",
      qty:1,unit:GRP[selGrp||"jarabes"].defUnit,notes:""});
    setModal(true);
  };
  const save=()=>{
    if(+form.qty<=0) return showToast("Cantidad inválida","error");
    setProdLogs(prev=>[...prev,{...form,id:nid("prod"),qty:+form.qty}]);
    showToast("Lote registrado ✓");setModal(false);
  };

  const filtered=useMemo(()=>prodLogs.filter(l=>!selGrp||l.grp===selGrp),[prodLogs,selGrp]);

  const grpTotals=useMemo(()=>{
    const r={};
    Object.keys(GRP).forEach(k=>{r[k]=prodLogs.filter(l=>l.grp===k).reduce((s,l)=>s+l.qty,0)});
    return r;
  },[prodLogs]);

  return(
    <div>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14}}>
        <div>
          <h2 style={{margin:0,fontSize:19,fontWeight:900,color:"#1E293B"}}>Producción</h2>
          <p style={{margin:"2px 0 0",fontSize:11,color:"#94A3B8"}}>Registro por sección</p>
        </div>
        <Btn onClick={openAdd}><Plus size={14}/> Registrar Lote</Btn>
      </div>

      {/* GROUP CARDS con totales */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(145px,1fr))",
        gap:10,marginBottom:18}}>
        {Object.entries(GRP).map(([k,g])=>(
          <div key={k} onClick={()=>setSelGrp(selGrp===k?null:k)}
            style={{background:selGrp===k?g.bg:"#fff",
              border:`2px solid ${selGrp===k?g.color:"#E2E8F0"}`,
              borderRadius:12,padding:"12px 14px",cursor:"pointer",transition:"all .15s",
              boxShadow:selGrp===k?`0 0 0 3px ${g.color}22`:"0 1px 3px rgba(0,0,0,.05)"}}>
            <div style={{display:"flex",alignItems:"center",gap:5,marginBottom:5}}>
              <span style={{fontSize:16}}>{g.emoji}</span>
              <span style={{fontSize:9,fontWeight:800,color:g.color,
                textTransform:"uppercase",letterSpacing:.4,lineHeight:1.2}}>{g.label}</span>
            </div>
            <p style={{margin:0,fontSize:19,fontWeight:900,color:"#1E293B",lineHeight:1}}>
              {grpTotals[k]||0} <span style={{fontSize:10,fontWeight:400,color:"#94A3B8"}}>und.</span>
            </p>
          </div>
        ))}
      </div>

      <Card>
        <Tbl cols={["Fecha","Grupo","Producto","Cant. Producida","Unidad","Notas"]}>
          {filtered.length===0&&<EmptyRow cols={6} msg="Sin lotes registrados"/>}
          {[...filtered].reverse().map(l=>{
            const p=products.find(x=>x.id===l.prodId);
            const g=GRP[l.grp];
            return(
              <Tr key={l.id}>
                <Td style={{color:"#64748B",whiteSpace:"nowrap"}}>{fmtD(l.date)}</Td>
                <Td><GrpBadge grp={l.grp}/></Td>
                <Td style={{fontWeight:700}}>{p?.name||"—"}</Td>
                <Td><span style={{fontWeight:900,color:"#6D28D9",fontSize:14}}>{l.qty}</span></Td>
                <Td style={{color:"#64748B"}}>{l.unit||p?.unit||"—"}</Td>
                <Td style={{color:"#94A3B8",fontSize:11}}>{l.notes||"—"}</Td>
              </Tr>
            );
          })}
        </Tbl>
      </Card>

      {modal&&(
        <Modal title="Registrar Lote de Producción" onClose={()=>setModal(false)}>
          <Field label="Grupo / Línea">
            <FSelect value={form.grp} onChange={e=>{
              const g=e.target.value;
              const first=products.find(p=>p.grp===g);
              setForm(prev=>({...prev,grp:g,prodId:first?.id||"",unit:GRP[g].defUnit}));
            }}>
              {Object.entries(GRP).map(([k,g])=><option key={k} value={k}>{g.emoji} {g.label}</option>)}
            </FSelect>
          </Field>
          <Field label="Producto elaborado">
            <ProductPicker products={products} value={form.prodId} grp={form.grp}
              onChange={id=>setForm(prev=>({...prev,prodId:id}))}
              onCreate={name=>createProductInline(setProducts,showToast,form.grp,name)}/>
          </Field>
          <FormRow>
            <Field label="Fecha" half><FInput type="date" value={form.date} onChange={f("date")}/></Field>
            <Field label="Unidad de medida" half>
              <FSelect value={form.unit} onChange={f("unit")}>
                {UNITS.map(u=><option key={u} value={u}>{u}</option>)}
              </FSelect>
            </Field>
            <Field label="Cantidad producida" half>
              <FInput type="number" value={form.qty} onChange={f("qty")} min={1}/>
            </Field>
          </FormRow>
          <Field label="Notas"><FArea value={form.notes} onChange={f("notes")} placeholder="Detalles del lote..."/></Field>
          <div style={{display:"flex",gap:10,justifyContent:"flex-end"}}>
            <Btn variant="secondary" onClick={()=>setModal(false)}>Cancelar</Btn>
            <Btn onClick={save}><FlaskConical size={13}/> Registrar</Btn>
          </div>
        </Modal>
      )}
    </div>
  );
}

/* ══════════════════════════════════════════════
   PRECIOS VIEW
══════════════════════════════════════════════ */
function PreciosView({products,setProducts,showToast}){
  const [selGrp,setSelGrp]=useState(null);
  const [search,setSearch]=useState("");
  const [editing,setEditing]=useState(null);
  const [tp,setTp]=useState("");
  const [tc,setTc]=useState("");
  const [page,setPage]=useState(0);
  const PAGE=40;

  const filtered=useMemo(()=>products.filter(p=>{
    const q=search.trim().toLowerCase();
    return (!selGrp||p.grp===selGrp) && (!q||p.name.toLowerCase().includes(q));
  }),[products,selGrp,search]);
  useEffect(()=>{setPage(0)},[selGrp,search]);

  const totalPages=Math.max(1,Math.ceil(filtered.length/PAGE));
  const pageItems=filtered.slice(page*PAGE,(page+1)*PAGE);

  const startEdit=p=>{setEditing(p.id);setTp(String(p.price));setTc(String(p.cost));};
  const saveEdit=id=>{
    setProducts(prev=>prev.map(p=>p.id===id?{...p,price:+tp,cost:+tc}:p));
    setEditing(null);showToast("Precios actualizados ✓");
  };

  const totalCost=filtered.reduce((s,p)=>s+p.stock*p.cost,0);
  const totalPrice=filtered.reduce((s,p)=>s+p.stock*p.price,0);

  return(
    <div>
      <h2 style={{margin:"0 0 14px",fontSize:19,fontWeight:900,color:"#1E293B"}}>Precios</h2>
      <div style={{display:"flex",gap:12,flexWrap:"wrap",marginBottom:18}}>
        <StatCard label="Valor a Costo"  value={fmt(totalCost)}        sub="inventario" icon={<Tag size={16}/>} color="#64748B"/>
        <StatCard label="Valor a Precio" value={fmt(totalPrice)}       sub="a venta"    icon={<TrendingUp size={16}/>} color={BRAND}/>
        <StatCard label="Ganancia Pot."  value={fmt(totalPrice-totalCost)} sub="margen bruto" icon={<TrendingUp size={16}/>} color="#059669"/>
      </div>
      <div style={{display:"flex",gap:8,marginBottom:14,flexWrap:"wrap"}}>
        <SearchBar value={search} onChange={setSearch} placeholder="Buscar producto..."/>
        <FSelect value={selGrp||""} onChange={e=>setSelGrp(e.target.value||null)} style={{width:"auto",minWidth:180}}>
          <option value="">Todos los grupos</option>
          {Object.entries(GRP).map(([k,g])=><option key={k} value={k}>{g.emoji} {g.label}</option>)}
        </FSelect>
      </div>
      <Card>
        <Tbl cols={["Producto","Grupo","Unidad","Costo","Precio Venta","Margen $","Margen %",""]}>
          {pageItems.length===0&&<EmptyRow cols={8}/>}
          {pageItems.map(p=>{
            const isEd=editing===p.id;
            const m=p.price-p.cost;
            const mp=p.cost>0?Math.round((m/p.cost)*100):0;
            return(
              <Tr key={p.id}>
                <Td style={{fontWeight:700,maxWidth:160}}>{p.name}</Td>
                <Td><GrpBadge grp={p.grp}/></Td>
                <Td style={{color:"#64748B"}}>{p.unit}</Td>
                <Td>{isEd?<FInput type="number" value={tc} onChange={e=>setTc(e.target.value)} style={{width:85}}/>:fmt(p.cost)}</Td>
                <Td>{isEd?<FInput type="number" value={tp} onChange={e=>setTp(e.target.value)} style={{width:85}}/>:<span style={{fontWeight:700,color:BRAND}}>{fmt(p.price)}</span>}</Td>
                <Td style={{color:m>=0?"#059669":"#DC2626",fontWeight:700}}>{fmt(m)}</Td>
                <Td><span style={{fontWeight:800,color:mp>=30?"#059669":mp>=10?"#D97706":"#DC2626"}}>{mp}%</span></Td>
                <Td>{isEd
                  ?<div style={{display:"flex",gap:4}}>
                      <Btn small variant="success" onClick={()=>saveEdit(p.id)}>✓</Btn>
                      <Btn small variant="secondary" onClick={()=>setEditing(null)}>✕</Btn>
                    </div>
                  :<Btn small variant="secondary" onClick={()=>startEdit(p)}><Edit2 size={11}/></Btn>
                }</Td>
              </Tr>
            );
          })}
        </Tbl>
        {filtered.length>PAGE&&(
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",
            padding:"12px 16px",borderTop:"1px solid #F1F5F9",flexWrap:"wrap",gap:8}}>
            <span style={{fontSize:12,color:"#94A3B8",fontWeight:600}}>
              {page*PAGE+1}–{Math.min((page+1)*PAGE,filtered.length)} de {filtered.length}
            </span>
            <div style={{display:"flex",gap:6}}>
              <Btn small variant="secondary" disabled={page===0}
                onClick={()=>setPage(p=>Math.max(0,p-1))}>Anterior</Btn>
              <span style={{fontSize:12,fontWeight:700,color:"#475569",
                alignSelf:"center",padding:"0 4px"}}>{page+1} / {totalPages}</span>
              <Btn small variant="secondary" disabled={page>=totalPages-1}
                onClick={()=>setPage(p=>Math.min(totalPages-1,p+1))}>Siguiente</Btn>
            </div>
          </div>
        )}
      </Card>
    </div>
  );
}

/* ══════════════════════════════════════════════
   REPORTES VIEW
══════════════════════════════════════════════ */
function ReportesView({products,entries,exits,clients,suppliers,prodLogs,setProducts,showToast}){
  const fileRef=useRef();

  const exportSheet=async(name,data,cols)=>{
    // Construye filas legibles: grupos con nombre, fechas con formato
    const rows=data.map(row=>{
      const r={};
      cols.forEach(c=>{
        let v=row[c.k];
        if(c.k==="grp")      v=GRP[v]?.label ?? v;          // "jarabes" -> "JARABES"
        else if(c.k==="date")v=v?fmtD(v):"";                // fecha legible
        else if(v==null)     v="";
        r[c.h]=v;
      });
      return r;
    });
    const ws=XLSX.utils.json_to_sheet(rows,{header:cols.map(c=>c.h)});
    // Ancho de columnas automático según el contenido
    ws["!cols"]=cols.map(c=>{
      const dataMax=rows.reduce((m,row)=>Math.max(m,String(row[c.h]??"").length),0);
      return { wch: Math.min(42, Math.max(12, c.h.length, dataMax)+2) };
    });
    // Congela la fila de encabezados
    ws["!freeze"]={xSplit:0,ySplit:1};
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,name);
    await saveWorkbook(wb,`Sirope_${name}_${tdayStr()}.xlsx`,showToast);
  };

  const exportInv=()=>exportSheet("Inventario",products,[
    {k:"name",h:"Producto"},{k:"grp",h:"Grupo"},{k:"stock",h:"Stock"},
    {k:"min",h:"Stock Mínimo"},{k:"unit",h:"Unidad"},{k:"cost",h:"Costo"},{k:"price",h:"Precio"},
  ]);

  const exportEnt=()=>{
    const d=entries.map(e=>({...e,
      producto:products.find(p=>p.id===e.prodId)?.name||"",
      proveedor:suppliers.find(p=>p.id===e.provId)?.name||"",
      total:e.qty*e.cost}));
    exportSheet("Entradas",d,[
      {k:"date",h:"Fecha"},{k:"producto",h:"Producto"},{k:"qty",h:"Cantidad"},
      {k:"cost",h:"Costo Unit."},{k:"total",h:"Total"},{k:"proveedor",h:"Proveedor"},{k:"notes",h:"Notas"},
    ]);
  };

  const exportSal=()=>{
    const d=exits.map(e=>({...e,
      producto:products.find(p=>p.id===e.prodId)?.name||"",
      cliente:clients.find(c=>c.id===e.cliId)?.name||"",
      total:e.qty*e.price}));
    exportSheet("Salidas",d,[
      {k:"date",h:"Fecha"},{k:"producto",h:"Producto"},{k:"qty",h:"Cantidad"},
      {k:"price",h:"Precio Unit."},{k:"total",h:"Total"},{k:"cliente",h:"Cliente"},{k:"notes",h:"Notas"},
    ]);
  };

  const importProds=async e=>{
    const file=e.target.files[0];if(!file)return;
    try{
      const buf=await file.arrayBuffer();
      const wb=XLSX.read(buf);
      const ws=wb.Sheets[wb.SheetNames[0]];
      const rows=XLSX.utils.sheet_to_json(ws);
      let added=0;
      rows.forEach(row=>{
        const name=row["Producto"]||row["nombre"]||row["Name"]||"";if(!name)return;
        const grpLabel=String(row["Grupo"]||row["Categoría"]||"").toLowerCase();
        const grp=Object.entries(GRP).find(([,g])=>g.label.toLowerCase()===grpLabel)?.[0]||"jarabes";
        setProducts(prev=>[...prev,{
          id:nid("p"),name:name.trim(),grp,
          stock:+row["Stock"]||0,min:+row["Stock Mínimo"]||5,
          unit:row["Unidad"]||GRP[grp].defUnit,
          cost:+row["Costo"]||0,price:+row["Precio"]||0,
        }]);
        added++;
      });
      showToast(`${added} productos importados ✓`);
    }catch{showToast("Error al leer el archivo","error");}
    e.target.value="";
  };

  const ExpBtn=({label,icon,onClick,color})=>(
    <div style={{background:"#fff",borderRadius:12,padding:"16px 18px",
      boxShadow:"0 1px 3px rgba(0,0,0,.07)",display:"flex",alignItems:"center",
      justifyContent:"space-between",gap:12}}>
      <div style={{display:"flex",alignItems:"center",gap:12}}>
        <div style={{background:color+"18",color,borderRadius:10,padding:9,display:"flex"}}>{icon}</div>
        <span style={{fontWeight:700,fontSize:13,color:"#1E293B"}}>{label}</span>
      </div>
      <Btn onClick={onClick} small><Download size={11}/> Exportar</Btn>
    </div>
  );

  return(
    <div>
      <h2 style={{margin:"0 0 18px",fontSize:19,fontWeight:900,color:"#1E293B"}}>Reportes</h2>
      <h3 style={{fontSize:11,fontWeight:800,color:"#94A3B8",textTransform:"uppercase",
        letterSpacing:.7,marginBottom:10}}>Exportar a Excel</h3>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(290px,1fr))",gap:10,marginBottom:26}}>
        <ExpBtn label={`Inventario (${products.length} productos)`} icon={<Package size={16}/>} color={BRAND} onClick={exportInv}/>
        <ExpBtn label={`Entradas (${entries.length})`} icon={<ArrowDownLeft size={16}/>} color="#059669" onClick={exportEnt}/>
        <ExpBtn label={`Salidas (${exits.length})`} icon={<ArrowUpRight size={16}/>} color="#DC2626" onClick={exportSal}/>
        <ExpBtn label={`Clientes (${clients.length})`} icon={<Users size={16}/>} color="#1E40AF" onClick={()=>exportSheet("Clientes",clients,[
          {k:"name",h:"Nombre"},{k:"phone",h:"Teléfono"},{k:"email",h:"Email"},{k:"address",h:"Dirección"},{k:"notes",h:"Notas"}])}/>
        <ExpBtn label={`Proveedores (${suppliers.length})`} icon={<Truck size={16}/>} color="#D97706" onClick={()=>exportSheet("Proveedores",suppliers,[
          {k:"name",h:"Empresa"},{k:"contact",h:"Contacto"},{k:"phone",h:"Teléfono"},{k:"email",h:"Email"},{k:"address",h:"Dirección"}])}/>
        <ExpBtn label={`Producción (${prodLogs.length})`} icon={<FlaskConical size={16}/>} color="#6D28D9" onClick={()=>exportSheet("Produccion",prodLogs.map(l=>({...l,producto:products.find(p=>p.id===l.prodId)?.name||""})),[
          {k:"date",h:"Fecha"},{k:"grp",h:"Grupo"},{k:"producto",h:"Producto"},{k:"qty",h:"Cantidad"},{k:"unit",h:"Unidad"},{k:"notes",h:"Notas"}])}/>
      </div>
      <h3 style={{fontSize:11,fontWeight:800,color:"#94A3B8",textTransform:"uppercase",
        letterSpacing:.7,marginBottom:10}}>Importar desde Excel</h3>
      <div style={{background:"#F8FAFC",border:"2px dashed #E2E8F0",borderRadius:14,
        padding:26,textAlign:"center"}}>
        <Upload size={26} style={{color:"#CBD5E1",marginBottom:8}}/>
        <p style={{margin:"0 0 5px",fontWeight:700,fontSize:13,color:"#374151"}}>Importar Productos</p>
        <p style={{margin:"0 0 14px",fontSize:11,color:"#94A3B8"}}>
          Columnas: <strong>Producto, Grupo, Stock, Stock Mínimo, Unidad, Costo, Precio</strong>
        </p>
        <input type="file" accept=".xlsx,.xls,.csv" ref={fileRef} onChange={importProds} style={{display:"none"}}/>
        <Btn onClick={()=>fileRef.current?.click()}><Upload size={13}/> Seleccionar archivo</Btn>
      </div>
    </div>
  );
}

/* ══════════════════════════════════════════════
   MAIN APP
══════════════════════════════════════════════ */
const NAV=[
  {id:"dashboard",  label:"Dashboard",   icon:<LayoutDashboard size={16}/>},
  {id:"inventario", label:"Inventario",  icon:<Package size={16}/>},
  {id:"entradas",   label:"Entradas",    icon:<ArrowDownLeft size={16}/>},
  {id:"salidas",    label:"Salidas",     icon:<ArrowUpRight size={16}/>},
  {id:"clientes",   label:"Clientes",    icon:<Users size={16}/>},
  {id:"proveedores",label:"Proveedores", icon:<Truck size={16}/>},
  {id:"produccion", label:"Producción",  icon:<FlaskConical size={16}/>},
  {id:"precios",    label:"Precios",     icon:<Tag size={16}/>},
  {id:"reportes",   label:"Reportes",    icon:<FileSpreadsheet size={16}/>},
];

const CLI_FIELDS=[
  {key:"name",    label:"Nombre",    required:true,bold:true},
  {key:"phone",   label:"Teléfono",  half:true},
  {key:"email",   label:"Email",     half:true,type:"email"},
  {key:"address", label:"Dirección", half:true},
  {key:"notes",   label:"Notas",     half:true},
];
const PROV_FIELDS=[
  {key:"name",    label:"Empresa",   required:true,bold:true},
  {key:"contact", label:"Contacto",  half:true},
  {key:"phone",   label:"Teléfono",  half:true},
  {key:"email",   label:"Email",     half:true,type:"email"},
  {key:"address", label:"Dirección", half:true},
];

/* ══════════════════════════════════════════════
   PANTALLAS DE NUBE
══════════════════════════════════════════════ */
function CenterBox({children}){
  return(
    <div style={{height:"100vh",display:"flex",alignItems:"center",justifyContent:"center",
      background:"#0D1629",fontFamily:"system-ui,-apple-system,sans-serif",padding:24}}>
      <div style={{background:"#fff",borderRadius:18,padding:"32px 26px",maxWidth:420,
        textAlign:"center",boxShadow:"0 20px 60px rgba(0,0,0,.4)"}}>{children}</div>
    </div>
  );
}

function SetupScreen(){
  return(
    <CenterBox>
      <div style={{fontSize:34,marginBottom:8}}>☁️</div>
      <h2 style={{margin:"0 0 8px",fontSize:19,fontWeight:900,color:"#1E293B"}}>Falta conectar la nube</h2>
      <p style={{margin:"0 0 14px",fontSize:13,color:"#475569",lineHeight:1.5}}>
        Abre el archivo <b>firebase.js</b> y pega los datos de tu proyecto de Firebase
        donde dice <code>PEGA_AQUI_…</code>. Sigue la guía
        <b> CONECTAR-LA-NUBE.md</b> paso a paso.
      </p>
      <p style={{margin:0,fontSize:11,color:"#94A3B8"}}>Es un solo archivo. Toma ~10 minutos.</p>
    </CenterBox>
  );
}

function LoadingScreen({error}){
  return(
    <CenterBox>
      {error ? (
        <>
          <div style={{fontSize:32,marginBottom:8}}>⚠️</div>
          <h2 style={{margin:"0 0 8px",fontSize:17,fontWeight:900,color:"#DC2626"}}>No se pudo conectar</h2>
          <p style={{margin:0,fontSize:12,color:"#475569",lineHeight:1.5}}>{error}</p>
          <p style={{margin:"12px 0 0",fontSize:11,color:"#94A3B8"}}>
            Revisa que activaste <b>Authentication → Anónimo</b> en Firebase.
          </p>
        </>
      ) : (
        <>
          <div style={{width:38,height:38,border:"4px solid #E2E8F0",borderTopColor:BRAND,
            borderRadius:"50%",margin:"0 auto 14px",animation:"spin 1s linear infinite"}}/>
          <style>{"@keyframes spin{to{transform:rotate(360deg)}}"}</style>
          <p style={{margin:0,fontSize:14,fontWeight:700,color:"#1E293B"}}>Conectando a la nube…</p>
          <p style={{margin:"4px 0 0",fontSize:11,color:"#94A3B8"}}>Sincronizando inventario</p>
        </>
      )}
    </CenterBox>
  );
}

export default function SiroperApp(){
  const [view,      setView]      = useState("dashboard");
  const [products,  setProducts]  = useCloud("products");
  const [entries,   setEntries]   = useCloud("entries");
  const [exits,     setExits]     = useCloud("exits");
  const [clients,   setClients]   = useCloud("clients");
  const [suppliers, setSuppliers] = useCloud("suppliers");
  const [prodLogs,  setProdLogs]  = useCloud("prodlogs");
  const [toast,     setToast]     = useState(null);
  const [authReady, setAuthReady] = useState(false);
  const [authError, setAuthError] = useState(null);
  const [isMobile,  setIsMobile]  = useState(
    typeof window!=="undefined" && window.innerWidth<820);
  const [sideOpen,  setSideOpen]  = useState(
    !(typeof window!=="undefined" && window.innerWidth<820));

  // Inicia sesión anónima en la nube y siembra el catálogo la primera vez
  useEffect(()=>{
    if(!FIREBASE_LISTO) return;
    const unsub=onAuthStateChanged(auth,async(user)=>{
      if(user){
        setAuthReady(true);
        // Sembrar catálogo de productos solo una vez
        try{
          const metaRef=doc(db,"meta","app");
          const metaSnap=await getDoc(metaRef);
          if(!(metaSnap.exists() && metaSnap.data().seeded)){
            const batch=writeBatch(db);
            PRODS_INIT.forEach(p=>{ const {id,...d}=p; batch.set(doc(db,"products",String(id)),d); });
            batch.set(metaRef,{seeded:true,seededAt:Date.now()});
            await batch.commit();
          }
        }catch(e){ console.error("seed",e); }
      }
    });
    signInAnonymously(auth).catch(e=>setAuthError(e.message||String(e)));
    return ()=>unsub();
  },[]);

  // Track viewport so the sidebar behaves as a drawer on phones
  useEffect(()=>{
    const onResize=()=>{
      const mobile=window.innerWidth<820;
      setIsMobile(mobile);
      if(mobile) setSideOpen(false); else setSideOpen(true);
    };
    window.addEventListener("resize",onResize);
    return ()=>window.removeEventListener("resize",onResize);
  },[]);

  // Estado de conexión (para mostrar online / sin conexión)
  const [online,setOnline]=useState(
    typeof navigator!=="undefined" ? navigator.onLine : true);
  useEffect(()=>{
    const up=()=>setOnline(true), down=()=>setOnline(false);
    window.addEventListener("online",up);
    window.addEventListener("offline",down);
    return ()=>{ window.removeEventListener("online",up); window.removeEventListener("offline",down); };
  },[]);

  const go=(id)=>{ setView(id); if(isMobile) setSideOpen(false); };

  const showToast=(msg,type="success")=>{
    setToast({msg,type});setTimeout(()=>setToast(null),3500);
  };

  // Pantalla de aviso si Firebase aún no está configurado
  if(!FIREBASE_LISTO) return <SetupScreen/>;
  // Pantalla de carga mientras conecta a la nube
  if(!authReady) return <LoadingScreen error={authError}/>;

  const alerts=products.filter(p=>p.stock<p.min).length;
  const props={products,setProducts,entries,setEntries,exits,setExits,
    clients,setClients,suppliers,setSuppliers,prodLogs,setProdLogs,showToast};

  const renderView=()=>{
    switch(view){
      case "dashboard":   return <Dashboard {...props} setView={setView}/>;
      case "inventario":  return <InventarioView {...props}/>;
      case "entradas":    return <EntradasView {...props}/>;
      case "salidas":     return <SalidasView {...props}/>;
      case "clientes":    return <CRUDView title="Clientes"    items={clients}   setItems={setClients}   fields={CLI_FIELDS}  idKey="cli"  showToast={showToast}/>;
      case "proveedores": return <CRUDView title="Proveedores" items={suppliers} setItems={setSuppliers} fields={PROV_FIELDS} idKey="prov" showToast={showToast}/>;
      case "produccion":  return <ProduccionView {...props}/>;
      case "precios":     return <PreciosView {...props}/>;
      case "reportes":    return <ReportesView {...props}/>;
      default: return null;
    }
  };

  return(
    <div style={{display:"flex",height:"100vh",fontFamily:"system-ui,-apple-system,sans-serif",
      background:"#F1F5F9",overflow:"hidden"}}>

      {/* Backdrop (mobile, when drawer open) */}
      {isMobile&&sideOpen&&(
        <div onClick={()=>setSideOpen(false)}
          style={{position:"fixed",inset:0,background:"rgba(0,0,0,.45)",zIndex:90}}/>
      )}

      {/* ── SIDEBAR ── */}
      <aside style={{
        width: isMobile ? 230 : (sideOpen?SIDEBAR_W:58),
        background:"#0D1629",display:"flex",flexDirection:"column",
        transition:"transform .2s, width .2s",overflow:"hidden",flexShrink:0,
        boxShadow:"4px 0 24px rgba(0,0,0,.35)",
        ...(isMobile ? {
          position:"fixed",top:0,left:0,bottom:0,zIndex:100,
          transform: sideOpen?"translateX(0)":"translateX(-100%)",
        } : {})
      }}>

        {/* Logo Sirope */}
        <div style={{padding:"16px 12px 14px",borderBottom:"1px solid rgba(255,255,255,.07)",
          display:"flex",alignItems:"center",gap:10,minHeight:60,flexShrink:0}}>
          <div style={{width:34,height:34,background:"#fff",borderRadius:10,flexShrink:0,
            display:"flex",alignItems:"center",justifyContent:"center",
            boxShadow:`0 0 16px ${BRAND}55`}}>
            <span style={{fontSize:18,fontStyle:"italic",fontFamily:"Georgia,serif",
              fontWeight:900,color:"#0D1629",lineHeight:1}}>S</span>
          </div>
          {(sideOpen||isMobile)&&(
            <div style={{overflow:"hidden"}}>
              <p style={{margin:0,fontFamily:"Georgia,serif",fontStyle:"italic",
                fontSize:19,fontWeight:900,color:"#fff",lineHeight:1,letterSpacing:-.5}}>Sirope</p>
              <p style={{margin:0,fontSize:9,color:"rgba(255,255,255,.35)",
                fontWeight:700,textTransform:"uppercase",letterSpacing:1}}>Inventario</p>
            </div>
          )}
          <button onClick={()=>setSideOpen(p=>!p)}
            style={{marginLeft:"auto",background:"none",border:"none",
              color:"rgba(255,255,255,.35)",cursor:"pointer",display:"flex",
              padding:4,flexShrink:0}}><X size={isMobile?18:17}/></button>
        </div>

        {/* Nav items */}
        <nav style={{flex:1,padding:"8px 6px",overflowY:"auto"}}>
          {NAV.map(item=>{
            const active=view===item.id;
            const showLabel = sideOpen||isMobile;
            return(
              <button key={item.id} onClick={()=>go(item.id)}
                style={{width:"100%",display:"flex",alignItems:"center",gap:10,
                  padding:"10px 9px",borderRadius:9,border:"none",cursor:"pointer",
                  marginBottom:2,textAlign:"left",transition:"all .15s",
                  background:active?"rgba(255,107,53,.14)":"transparent",
                  color:active?BRAND:"rgba(255,255,255,.5)"}}>
                <span style={{flexShrink:0,color:active?BRAND:"rgba(255,255,255,.38)",display:"flex"}}>
                  {item.icon}
                </span>
                {showLabel&&(
                  <span style={{fontSize:12.5,fontWeight:active?800:500,whiteSpace:"nowrap"}}>
                    {item.label}
                    {item.id==="dashboard"&&alerts>0&&(
                      <span style={{background:"#DC2626",color:"#fff",borderRadius:99,
                        fontSize:9,fontWeight:900,padding:"1px 5px",marginLeft:5}}>{alerts}</span>
                    )}
                  </span>
                )}
              </button>
            );
          })}
        </nav>

        {/* Footer */}
        <div style={{padding:"10px 12px",borderTop:"1px solid rgba(255,255,255,.05)",
          display:"flex",alignItems:"center",gap:7,flexShrink:0}}>
          <div style={{width:7,height:7,borderRadius:99,
            background:online?"#059669":"#F59E0B",
            boxShadow:`0 0 6px ${online?"#059669":"#F59E0B"}`,flexShrink:0}}/>
          {(sideOpen||isMobile)&&<span style={{fontSize:10,color:"rgba(255,255,255,.35)",fontWeight:700}}>
            {online ? `En línea · ${products.length} productos` : "Sin conexión · se sincronizará"}
          </span>}
        </div>
      </aside>

      {/* ── MAIN ── */}
      <main style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden",minWidth:0}}>
        {/* Header */}
        <header style={{background:"#fff",borderBottom:"1px solid #F1F5F9",
          padding:"0 16px",height:52,display:"flex",alignItems:"center",
          justifyContent:"space-between",flexShrink:0,gap:10}}>
          <div style={{display:"flex",alignItems:"center",gap:10,minWidth:0}}>
            {(isMobile||!sideOpen)&&(
              <button onClick={()=>setSideOpen(true)}
                style={{background:"#F1F5F9",border:"none",borderRadius:8,width:34,height:34,
                  cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",
                  color:"#475569",flexShrink:0}}><Menu size={18}/></button>
            )}
            <h1 style={{margin:0,fontSize:14,fontWeight:900,color:"#1E293B",
              whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>
              {NAV.find(n=>n.id===view)?.label}
            </h1>
          </div>
          <div style={{display:"flex",alignItems:"center",gap:10,flexShrink:0}}>
            {alerts>0&&(
              <button onClick={()=>setView("inventario")}
                style={{background:"#FEF3C7",border:"none",borderRadius:7,
                  padding:"4px 10px",cursor:"pointer",display:"flex",alignItems:"center",
                  gap:5,color:"#D97706",fontWeight:800,fontSize:11,whiteSpace:"nowrap"}}>
                <AlertTriangle size={12}/> {alerts}{!isMobile&&" alertas"}
              </button>
            )}
            <div style={{width:30,height:30,borderRadius:99,
              background:"#0D1629",display:"flex",alignItems:"center",
              justifyContent:"center",fontFamily:"Georgia,serif",fontStyle:"italic",
              fontSize:14,color:"#fff",fontWeight:900,flexShrink:0}}>S</div>
          </div>
        </header>

        {/* Content */}
        <div style={{flex:1,overflow:"auto",padding:isMobile?14:20}}>
          {renderView()}
        </div>
      </main>

      {toast&&<Toast msg={toast.msg} type={toast.type} onClose={()=>setToast(null)}/>}
    </div>
  );
}
