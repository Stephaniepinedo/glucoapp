import React, { useState, useEffect } from "react";

const CLIENT_ID = "56cd2e49-dddb-4dc1-8e66-d9e7b9338f6c";
const SCOPES = ["User.Read", "Files.ReadWrite", "offline_access"];
const REDIRECT_URI = window.location.origin;
const FILE_NAME = "GlucoApp.xlsx";

// ── Token helpers ──
const getToken = () => localStorage.getItem("ms_token");
const setToken = (t) => localStorage.setItem("ms_token", t);
const clearToken = () => localStorage.removeItem("ms_token");

// ── PKCE helpers ──
const generateCodeVerifier = () => {
  const array = new Uint8Array(32);
  crypto.getRandomValues(array);
  return btoa(String.fromCharCode(...array)).replace(/\+/g,'-').replace(/\//g,'_').replace(/=/g,'');
};
const generateCodeChallenge = async (verifier) => {
  const data = new TextEncoder().encode(verifier);
  const digest = await crypto.subtle.digest('SHA-256', data);
  return btoa(String.fromCharCode(...new Uint8Array(digest))).replace(/\+/g,'-').replace(/\//g,'_').replace(/=/g,'');
};
const exchangeCodeForToken = async (code) => {
  const verifier = sessionStorage.getItem('pkce_verifier');
  const res = await fetch('https://login.microsoftonline.com/common/oauth2/v2.0/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: CLIENT_ID, grant_type: 'authorization_code',
      code, redirect_uri: REDIRECT_URI, code_verifier: verifier,
    }),
  });
  const data = await res.json();
  return data.access_token;
};

// ── Login Screen ──
function AuthScreen() {
  const C2 = { bg:"#f8fafc", card:"#ffffff", border:"#e5e7eb", blue:"#1d4ed8", text:"#111827", muted:"#6b7280" };

  const handleMicrosoftLogin = async () => {
    const verifier = generateCodeVerifier();
    const challenge = await generateCodeChallenge(verifier);
    sessionStorage.setItem("pkce_verifier", verifier);
    const params = new URLSearchParams({
      client_id: CLIENT_ID, response_type: "code",
      redirect_uri: REDIRECT_URI, scope: SCOPES.join(" "),
      code_challenge: challenge, code_challenge_method: "S256",
      response_mode: "query",
    });
    window.location.href = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?${params}`;
  };

  return (
    <div style={{fontFamily:"system-ui,sans-serif",background:"#eff6ff",minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{width:"100%",maxWidth:360}}>
        <div style={{textAlign:"center",marginBottom:36}}>
          <div style={{width:72,height:72,background:"#1d4ed8",borderRadius:20,display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 16px",fontSize:36}}>💉</div>
          <div style={{fontSize:11,color:C2.blue,fontWeight:700,letterSpacing:3,marginBottom:6}}>GLUCOAPP</div>
          <div style={{fontSize:26,fontWeight:700,color:C2.text,marginBottom:6}}>Control Diabetes</div>
          <div style={{fontSize:14,color:C2.muted,lineHeight:1.6}}>Calcula dosis, registra comidas<br/>y mejora cada día</div>
        </div>
        <div style={{background:C2.card,borderRadius:20,padding:28,boxShadow:"0 2px 12px rgba(0,0,0,0.08)"}}>
          <div style={{fontSize:15,fontWeight:600,color:C2.text,textAlign:"center",marginBottom:8}}>Bienvenida</div>
          <div style={{fontSize:13,color:C2.muted,textAlign:"center",marginBottom:24,lineHeight:1.6}}>
            Inicia sesión con tu cuenta Microsoft para acceder a tus datos desde cualquier dispositivo
          </div>
          <button onClick={handleMicrosoftLogin}
            style={{width:"100%",display:"flex",alignItems:"center",justifyContent:"center",gap:12,background:"#ffffff",border:"1.5px solid #e5e7eb",borderRadius:14,padding:"14px 0",fontSize:15,fontWeight:600,color:C2.text,cursor:"pointer",boxShadow:"0 1px 3px rgba(0,0,0,0.08)"}}>
            <svg width="20" height="20" viewBox="0 0 21 21" fill="none" xmlns="http://www.w3.org/2000/svg">
              <rect x="1" y="1" width="9" height="9" fill="#f25022"/>
              <rect x="11" y="1" width="9" height="9" fill="#7fba00"/>
              <rect x="1" y="11" width="9" height="9" fill="#00a4ef"/>
              <rect x="11" y="11" width="9" height="9" fill="#ffb900"/>
            </svg>
            Entrar con Microsoft
          </button>
          <div style={{marginTop:20,padding:"12px 14px",background:"#f8fafc",borderRadius:10,fontSize:12,color:C2.muted,lineHeight:1.7,textAlign:"center"}}>
            🔒 Tus datos se guardan en <strong>tu propio OneDrive</strong>.<br/>
            Cada cuenta Microsoft tiene sus datos completamente separados.
          </div>
        </div>
        <div style={{textAlign:"center",marginTop:20,fontSize:12,color:C2.muted}}>
          ⚠️ Herramienta de referencia. Confirma las dosis con tu endocrinólogo.
        </div>
      </div>
    </div>
  );
}

const FOODS = [
  { name: "Arroz", portion: "1/3 taza", carbs: 15, protein: 2, kcal: 67, cat: "Cereales" },
  { name: "Arroz", portion: "1/2 taza", carbs: 22, protein: 3, kcal: 98, cat: "Cereales" },
  { name: "Arroz", portion: "1 taza", carbs: 45, protein: 5, kcal: 200, cat: "Cereales" },
  { name: "Arepa de maíz", portion: "1 unidad CD", carbs: 30, protein: 4, kcal: 180, cat: "Cereales" },
  { name: "Arepa paisa", portion: "1 unidad delgada", carbs: 15, protein: 2, kcal: 90, cat: "Cereales" },
  { name: "Pan tajado", portion: "1 tajada", carbs: 15, protein: 3, kcal: 75, cat: "Cereales" },
  { name: "Pasta", portion: "1/3 taza", carbs: 15, protein: 2, kcal: 70, cat: "Cereales" },
  { name: "Pasta", portion: "1/2 taza", carbs: 23, protein: 4, kcal: 110, cat: "Cereales" },
  { name: "Pasta", portion: "1 taza", carbs: 45, protein: 7, kcal: 220, cat: "Cereales" },
  { name: "Avena cruda", portion: "3 cdas", carbs: 15, protein: 3, kcal: 75, cat: "Cereales" },
  { name: "Granola", portion: "1/4 taza", carbs: 15, protein: 2, kcal: 110, cat: "Cereales" },
  { name: "Granola", portion: "1/2 taza", carbs: 30, protein: 4, kcal: 220, cat: "Cereales" },
  { name: "Empanada", portion: "1 unidad mediana", carbs: 30, protein: 8, kcal: 280, cat: "Cereales" },
  { name: "Buñuelo", portion: "1 unidad pequeña", carbs: 12, protein: 3, kcal: 120, cat: "Cereales" },
  { name: "Pandebono", portion: "1 unidad mediana", carbs: 20, protein: 5, kcal: 160, cat: "Cereales" },
  { name: "Pan de yuca", portion: "1 unidad grande", carbs: 18, protein: 3, kcal: 140, cat: "Cereales" },
  { name: "Almojábana", portion: "1 unidad mediana", carbs: 22, protein: 5, kcal: 180, cat: "Cereales" },
  { name: "Pancake", portion: "1 unidad", carbs: 15, protein: 3, kcal: 90, cat: "Cereales" },
  { name: "Waffle", portion: "1 unidad", carbs: 15, protein: 3, kcal: 95, cat: "Cereales" },
  { name: "Quinua", portion: "7 cdas", carbs: 15, protein: 4, kcal: 80, cat: "Cereales" },
  { name: "Papa común", portion: "1 mediana", carbs: 15, protein: 2, kcal: 77, cat: "Tubérculos" },
  { name: "Papa común", portion: "1 grande", carbs: 30, protein: 4, kcal: 154, cat: "Tubérculos" },
  { name: "Papa criolla", portion: "1 pequeña", carbs: 5, protein: 1, kcal: 28, cat: "Tubérculos" },
  { name: "Papa criolla", portion: "3 pequeñas", carbs: 15, protein: 2, kcal: 75, cat: "Tubérculos" },
  { name: "Papa francesa", portion: "10 unidades", carbs: 15, protein: 2, kcal: 120, cat: "Tubérculos" },
  { name: "Plátano hartón", portion: "1/4 unidad", carbs: 15, protein: 1, kcal: 65, cat: "Tubérculos" },
  { name: "Plátano hartón", portion: "1/2 unidad", carbs: 30, protein: 1, kcal: 130, cat: "Tubérculos" },
  { name: "Guinéo", portion: "1 unidad pequeña", carbs: 15, protein: 1, kcal: 65, cat: "Tubérculos" },
  { name: "Yuca", portion: "1 astilla pequeña", carbs: 15, protein: 1, kcal: 80, cat: "Tubérculos" },
  { name: "Ñame", portion: "1 trozo mediano", carbs: 15, protein: 2, kcal: 70, cat: "Tubérculos" },
  { name: "Banano", portion: "1/2 mediano", carbs: 15, protein: 1, kcal: 60, cat: "Frutas" },
  { name: "Banano", portion: "1 mediano", carbs: 30, protein: 1, kcal: 120, cat: "Frutas" },
  { name: "Manzana", portion: "1 pequeña", carbs: 15, protein: 0, kcal: 65, cat: "Frutas" },
  { name: "Naranja", portion: "1 mediana", carbs: 15, protein: 1, kcal: 60, cat: "Frutas" },
  { name: "Mango", portion: "1 pequeño", carbs: 15, protein: 1, kcal: 65, cat: "Frutas" },
  { name: "Papaya", portion: "1 taza picada", carbs: 15, protein: 1, kcal: 55, cat: "Frutas" },
  { name: "Fresa", portion: "13 unidades", carbs: 15, protein: 1, kcal: 50, cat: "Frutas" },
  { name: "Aguacate", portion: "1/4 mediano", carbs: 3.5, protein: 1, kcal: 80, cat: "Frutas" },
  { name: "Melón", portion: "1 taza", carbs: 15, protein: 1, kcal: 55, cat: "Frutas" },
  { name: "Sandía", portion: "1 tajada", carbs: 15, protein: 1, kcal: 50, cat: "Frutas" },
  { name: "Mandarina", portion: "1 grande", carbs: 15, protein: 1, kcal: 60, cat: "Frutas" },
  { name: "Frijol", portion: "1/2 taza", carbs: 15, protein: 7, kcal: 110, cat: "Leguminosas" },
  { name: "Lenteja", portion: "1/2 taza", carbs: 15, protein: 9, kcal: 115, cat: "Leguminosas" },
  { name: "Garbanzo", portion: "1/2 taza", carbs: 15, protein: 7, kcal: 115, cat: "Leguminosas" },
  { name: "Leche entera", portion: "1 vaso 200ml", carbs: 10, protein: 6, kcal: 120, cat: "Lácteos" },
  { name: "Leche entera", portion: "1/2 vaso 100ml", carbs: 5, protein: 3, kcal: 60, cat: "Lácteos" },
  { name: "Yogurt finesse", portion: "1 vaso 180ml", carbs: 11, protein: 5, kcal: 90, cat: "Lácteos" },
  { name: "Carne de res magra", portion: "100g", carbs: 0, protein: 26, kcal: 190, cat: "Carnes" },
  { name: "Carne de res magra", portion: "150g", carbs: 0, protein: 39, kcal: 285, cat: "Carnes" },
  { name: "Hígado de res", portion: "100g", carbs: 0, protein: 26, kcal: 135, cat: "Carnes" },
  { name: "Costilla de res", portion: "100g", carbs: 0, protein: 22, kcal: 250, cat: "Carnes" },
  { name: "Pechuga de pollo", portion: "100g", carbs: 0, protein: 31, kcal: 165, cat: "Carnes" },
  { name: "Pechuga de pollo", portion: "150g", carbs: 0, protein: 46, kcal: 248, cat: "Carnes" },
  { name: "Pierna/muslo pollo", portion: "100g", carbs: 0, protein: 25, kcal: 190, cat: "Carnes" },
  { name: "Lomo de cerdo", portion: "100g", carbs: 0, protein: 20, kcal: 185, cat: "Carnes" },
  { name: "Chicharrón asado", portion: "100g", carbs: 0, protein: 30, kcal: 500, cat: "Carnes" },
  { name: "Tilapia roja", portion: "100g", carbs: 0, protein: 26, kcal: 128, cat: "Pescados" },
  { name: "Trucha arco iris", portion: "100g", carbs: 0, protein: 22, kcal: 148, cat: "Pescados" },
  { name: "Cachama", portion: "100g", carbs: 0, protein: 18, kcal: 140, cat: "Pescados" },
  { name: "Bocachico", portion: "100g", carbs: 0, protein: 20, kcal: 110, cat: "Pescados" },
  { name: "Mojarra", portion: "100g", carbs: 0, protein: 22, kcal: 120, cat: "Pescados" },
  { name: "Atún enlatado", portion: "100g", carbs: 0, protein: 28, kcal: 130, cat: "Pescados" },
  { name: "Salmón", portion: "100g", carbs: 0, protein: 20, kcal: 208, cat: "Pescados" },
  { name: "Verdura", portion: "1/2 taza", carbs: 5, protein: 2, kcal: 25, cat: "Verduras" },
  { name: "Verdura cruda", portion: "1 taza", carbs: 5, protein: 1, kcal: 20, cat: "Verduras" },
  { name: "Piña colada", portion: "1 vaso", carbs: 35, protein: 0, kcal: 230, cat: "Bebidas" },
  { name: "Pan Masa Madre con Proteína", portion: "1 tajada ~50g", carbs: 22.5, protein: 7.1, kcal: 133, cat: "Cereales" },
  { name: "Pan Masa Madre con Proteína", portion: "2 tajadas 100g", carbs: 45, protein: 14.2, kcal: 266, cat: "Cereales" },
  { name: "DelOrigen Almendras", portion: "1 unidad 18g", carbs: 4.1, protein: 1.3, kcal: 54, cat: "Personalizados" },
  { name: "Mini Arepa Yuca Queso DonMaíz", portion: "1 unidad 50g", carbs: 20, protein: 7.5, kcal: 145, cat: "Personalizados" },
  { name: "DelOrigen Minis", portion: "1 unidad 50g", carbs: 10, protein: 5, kcal: 123, cat: "Personalizados" },
  { name: "Not Milk", portion: "1 vaso 200ml", carbs: 7.8, protein: 2.8, kcal: 61, cat: "Personalizados" },
  { name: "Pan de Arroz", portion: "1 tajada 22g", carbs: 12.8, protein: 1.8, kcal: 59, cat: "Personalizados" },
  { name: "Yogurt Snack Fit Me", portion: "1 vaso 150g", carbs: 9.2, protein: 5.1, kcal: 97, cat: "Lácteos" },
  { name: "Empanada Maíz Tradición Sierra Flor", portion: "1 unidad 45g", carbs: 8, protein: 3, kcal: 83, cat: "Personalizados" },
  { name: "Deditos Pakeeto", portion: "1 unidad 43g", carbs: 1.7, protein: 11, kcal: 197, cat: "Personalizados" },
  { name: "Pan de Bono DelOrigen", portion: "1 unidad 18g", carbs: 5, protein: 1.5, kcal: 45, cat: "Personalizados" },
  { name: "Chocolate Luker", portion: "1 porción", carbs: 2, protein: 1, kcal: 30, cat: "Personalizados" },
  { name: "Whey Protein Nutramerican", portion: "1/2 scoop 26g", carbs: 1.2, protein: 20, kcal: 93, cat: "Personalizados" },
  { name: "Whey Protein Nutramerican", portion: "1 scoop 40g", carbs: 2.4, protein: 40, kcal: 186, cat: "Personalizados" },
  { name: "Cereal Crispies The Protein Choice", portion: "1/4 sobre 50g", carbs: 23, protein: 10, kcal: 165, cat: "Personalizados" },
  { name: "Cereal Crispies The Protein Choice", portion: "100g", carbs: 46, protein: 20, kcal: 330, cat: "Personalizados" },
  { name: "Yogurt Griego San Martín", portion: "1/2 vaso 110ml", carbs: 4.6, protein: 7, kcal: 46, cat: "Personalizados" },
  { name: "Yogurt Griego San Martín", portion: "1 vaso 220ml", carbs: 9.2, protein: 14, kcal: 93, cat: "Personalizados" },
  { name: "Granola Amande Macadamia Chocolate", portion: "1 unidad 25g", carbs: 13, protein: 2.7, kcal: 120, cat: "Personalizados" },
  { name: "Granola Amande Macadamia Chocolate", portion: "2 unidades 50g", carbs: 26, protein: 5.4, kcal: 240, cat: "Personalizados" },
];

const CATS = ["Recientes","Todos","Cereales","Tubérculos","Frutas","Leguminosas","Lácteos","Carnes","Pescados","Verduras","Bebidas","Personalizados"];

const C = {
  bg:"#f8fafc", card:"#ffffff", border:"#e5e7eb",
  blue:"#1d4ed8", sky:"#0284c7", green:"#16a34a",
  red:"#dc2626", yellow:"#d97706", text:"#111827",
  muted:"#6b7280", purple:"#7c3aed", orange:"#ea580c",
  headerBg:"#ffffff", headerBorder:"#e5e7eb",
  inputBg:"#f1f5f9",
};

const inp = {
  background:"#f1f5f9", border:"1px solid #e5e7eb", outline:"none",
  color:"#111827", borderRadius:12, padding:"12px 14px",
  fontSize:16, width:"100%", boxSizing:"border-box", WebkitAppearance:"none",
};

const DEFAULT_SETTINGS = {
  ratios:[
    { label:"🌅 Mañana", from:"00:00", to:"12:00", ratio:3.5 },
    { label:"☀️ Tarde", from:"12:00", to:"19:00", ratio:3.5 },
    { label:"🌙 Noche", from:"19:00", to:"00:00", ratio:3.5 },
  ],
  sensitivity:30,
  hipoglucemia:54, glucemiaBaja:70, objetivo:100, glucemiaAlta:180, hiperglucemia:250,
  sexo:"Femenino", pesoKg:55, alturaCm:155, fechaNacimiento:"1991-07-01",
  pesoMeta:50,
  toujeoDosis:17, insulinaRapida:"Apidra®", insulinaLenta:"Toujeo®",
  otrosMedicamentos:["Atorvastatina 10mg","Dapaglifosina"],
  // Metas nutricionales diarias
  metaCarbs:130, metaProtein:85, metaKcal:1350,
};

const getCurrentRatio = (ratios) => {
  const now = new Date();
  const cur = now.getHours()*60 + now.getMinutes();
  for (const r of ratios) {
    const [fh,fm] = r.from.split(":").map(Number);
    const [th,tm] = r.to.split(":").map(Number);
    const from = fh*60+fm;
    const to = th===0&&tm===0 ? 24*60 : th*60+tm;
    if (cur>=from && cur<to) return r.ratio;
  }
  return ratios[0].ratio;
};

const graphGet = async (url) => {
  const res = await fetch(`https://graph.microsoft.com/v1.0${url}`, {
    headers: { Authorization: `Bearer ${getToken()}` }
  });
  if (!res.ok) throw new Error(await res.text());
  return res.json();
};

const graphPost = async (url, body) => {
  const res = await fetch(`https://graph.microsoft.com/v1.0${url}`, {
    method:"POST",
    headers:{ Authorization:`Bearer ${getToken()}`, "Content-Type":"application/json" },
    body: body ? JSON.stringify(body) : undefined,
  });
  if (!res.ok) throw new Error(await res.text());
  return res.json();
};

const graphPatch = async (url, body) => {
  const res = await fetch(`https://graph.microsoft.com/v1.0${url}`, {
    method:"PATCH",
    headers:{ Authorization:`Bearer ${getToken()}`, "Content-Type":"application/json" },
    body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(await res.text());
  return res.json();
};

const graphPut = async (url, body, contentType) => {
  const res = await fetch(`https://graph.microsoft.com/v1.0${url}`, {
    method:"PUT",
    headers:{ Authorization:`Bearer ${getToken()}`, "Content-Type": contentType || "application/json" },
    body,
  });
  if (!res.ok) throw new Error(await res.text());
  return res.json();
};

const HEADER_ROW = ["Fecha","Hora","Glucemia","Carbs","Proteina","Kcal","Insulina","Alimentos","Toujeo"];

// Minimal valid empty .xlsx file, base64-encoded
const EMPTY_XLSX_BASE64 = "UEsDBBQAAAAIAAAAIQAAAAAAAAAAAAAAAAAVAAAAZG9jUHJvcHMvY29yZS54bWxQSwECFAAUAAAACAAAACEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAZG9jUHJvcHMvY29yZS54bWxQSwUGAAAAAAEAAQAjAAAAGwAAAAAA";

const findOrCreateFile = async () => {
  // Search for an existing GlucoApp.xlsx
  const data = await graphGet(`/me/drive/root/search(q='${FILE_NAME}')`);
  let file = data.value?.find(f => f.name === FILE_NAME);
  if (file) return { file, created:false };

  // Not found — create it via a simple upload of an empty workbook,
  // then add the "Registros" worksheet with headers.
  const bin = atob(EMPTY_XLSX_BASE64);
  const bytes = new Uint8Array(bin.length);
  for (let i=0;i<bin.length;i++) bytes[i] = bin.charCodeAt(i);
  file = await graphPut(
    `/me/drive/root:/${FILE_NAME}:/content`,
    bytes,
    "application/octet-stream"
  );
  return { file, created:true };
};

const ensureSheet = async (fileId) => {
  const wb = await graphGet(`/me/drive/items/${fileId}/workbook/worksheets`);
  let sheet = wb.value?.find(s => s.name === "Registros");
  if (!sheet) {
    sheet = await graphPost(`/me/drive/items/${fileId}/workbook/worksheets/add`, { name:"Registros" });
    await graphPatch(
      `/me/drive/items/${fileId}/workbook/worksheets/Registros/range(address='A1:I1')`,
      { values:[HEADER_ROW] }
    );
  } else {
    try {
      const used = await graphGet(`/me/drive/items/${fileId}/workbook/worksheets/Registros/usedRange`);
      if (!used.values || used.values.length===0) {
        await graphPatch(
          `/me/drive/items/${fileId}/workbook/worksheets/Registros/range(address='A1:I1')`,
          { values:[HEADER_ROW] }
        );
      }
    } catch {
      await graphPatch(
        `/me/drive/items/${fileId}/workbook/worksheets/Registros/range(address='A1:I1')`,
        { values:[HEADER_ROW] }
      );
    }
  }

  // Ensure "Configuracion" sheet exists
  let configSheet = wb.value?.find(s => s.name === "Configuracion");
  if (!configSheet) {
    await graphPost(`/me/drive/items/${fileId}/workbook/worksheets/add`, { name:"Configuracion" });
    await graphPatch(
      `/me/drive/items/${fileId}/workbook/worksheets/Configuracion/range(address='A1:B1')`,
      { values:[["clave","valor"]] }
    );
  }

  // Ensure "Alimentos" sheet exists
  let alimentosSheet = wb.value?.find(s => s.name === "Alimentos");
  if (!alimentosSheet) {
    await graphPost(`/me/drive/items/${fileId}/workbook/worksheets/add`, { name:"Alimentos" });
    await graphPatch(
      `/me/drive/items/${fileId}/workbook/worksheets/Alimentos/range(address='A1:B1')`,
      { values:[["clave","valor"]] }
    );
  }

  return sheet;
};

// Save settings to OneDrive "Configuracion" sheet
const saveSettingsToOneDrive = async (fileId, settings) => {
  try {
    const json = JSON.stringify(settings);
    // Store as a single row: key="settings", value=JSON
    await graphPatch(
      `/me/drive/items/${fileId}/workbook/worksheets/Configuracion/range(address='A2:B2')`,
      { values:[["settings", json]] }
    );
  } catch {}
};

// Load settings from OneDrive "Configuracion" sheet
const loadSettingsFromOneDrive = async (fileId) => {
  try {
    const used = await graphGet(`/me/drive/items/${fileId}/workbook/worksheets/Configuracion/usedRange`);
    const rows = used.values || [];
    for (const row of rows) {
      if (row[0] === "settings" && row[1]) {
        return JSON.parse(row[1]);
      }
    }
  } catch {}
  return null;
};

// Save custom foods to OneDrive "Alimentos" sheet
const saveCustomFoodsToOneDrive = async (fileId, foods) => {
  try {
    const json = JSON.stringify(foods);
    await graphPatch(
      `/me/drive/items/${fileId}/workbook/worksheets/Alimentos/range(address='A2:B2')`,
      { values:[["customFoods", json]] }
    );
  } catch {}
};

// Load custom foods from OneDrive "Alimentos" sheet
const loadCustomFoodsFromOneDrive = async (fileId) => {
  try {
    const used = await graphGet(`/me/drive/items/${fileId}/workbook/worksheets/Alimentos/usedRange`);
    const rows = used.values || [];
    for (const row of rows) {
      if (row[0] === "customFoods" && row[1]) {
        return JSON.parse(row[1]);
      }
    }
  } catch {}
  return null;
};
const readAllRecords = async (fileId) => {
  const used = await graphGet(`/me/drive/items/${fileId}/workbook/worksheets/Registros/usedRange`);
  const rows = used.values || [];
  if (rows.length <= 1) return []; // only header or empty
  const dataRows = rows.slice(1);
  return dataRows
    .filter(r => r[0]) // must have a date
    .map((r,i) => ({
      id: `${r[0]}-${r[1]}-${i}`,
      date: String(r[0]||""),
      time: String(r[1]||""),
      glucose: String(r[2]??"-"),
      carbs: String(r[3]??"0"),
      protein: String(r[4]??"0"),
      kcal: String(r[5]??"0"),
      insulin: Number(r[6])||0,
      foods: String(r[7]??"-"),
      toujeo: Number(r[8])||0,
    }))
    .reverse(); // most recent first, matching how we display things
};

const appendRow = async (fileId, row) => {
  try {
    await graphPost(
      `/me/drive/items/${fileId}/workbook/worksheets/Registros/tables/Table1/rows/add`,
      { values: [row] }
    );
  } catch {
    const used = await graphGet(`/me/drive/items/${fileId}/workbook/worksheets/Registros/usedRange`);
    const nextRow = (used.rowCount || 1) + 1;
    await graphPatch(
      `/me/drive/items/${fileId}/workbook/worksheets/Registros/range(address='A${nextRow}:I${nextRow}')`,
      { values: [row] }
    );
  }
};

// ── Local cache (best-effort only — NOT the source of truth) ──
const STORAGE_KEY = "glucoapp-records";
const cacheRecords = (recs) => {
  try { localStorage.setItem(STORAGE_KEY, JSON.stringify(recs)); } catch {}
};
const loadCachedRecords = () => {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch { return []; }
};

// ── Group records by date ──
const groupByDate = (records) => {
  const map = {};
  for (const r of records) {
    if (!map[r.date]) map[r.date] = [];
    map[r.date].push(r);
  }
  return Object.entries(map).sort((a,b) => {
    // parse "DD/MM/YYYY"
    const parse = (s) => { const [d,m,y] = s.split("/"); return new Date(y,m-1,d); };
    return parse(b[0]) - parse(a[0]);
  });
};

const dayTotals = (recs) => ({
  carbs:   recs.reduce((s,r) => s + parseFloat(r.carbs  ||0), 0),
  protein: recs.reduce((s,r) => s + parseFloat(r.protein||0), 0),
  kcal:    recs.reduce((s,r) => s + parseFloat(r.kcal   ||0), 0),
  insulin: recs.reduce((s,r) => s + (r.insulin||0), 0),
  toujeo:  recs.reduce((s,r) => s + (r.toujeo||0), 0),
});

const CUSTOM_FOODS_KEY = "glucoapp-custom-foods";
const loadCustomFoods = () => {
  try {
    const raw = localStorage.getItem(CUSTOM_FOODS_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch { return []; }
};
const saveCustomFoods = (foods) => {
  try { localStorage.setItem(CUSTOM_FOODS_KEY, JSON.stringify(foods)); } catch {}
};

export default function Root() {
  const [msToken, setMsTokenState] = React.useState(getToken());
  const [userInfo, setUserInfo] = React.useState(null);

  // Handle OAuth PKCE redirect
  React.useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const code = params.get("code");
    if (code) {
      window.history.replaceState({}, "", window.location.pathname);
      exchangeCodeForToken(code).then(token => {
        if (token) { setToken(token); setMsTokenState(token); }
      }).catch(() => {});
    }
  }, []);

  // Get Microsoft user profile
  React.useEffect(() => {
    if (!msToken || userInfo) return;
    fetch("https://graph.microsoft.com/v1.0/me", {
      headers: { Authorization: `Bearer ${msToken}` }
    }).then(r => r.json()).then(data => setUserInfo(data)).catch(() => {});
  }, [msToken]);

  const handleLogout = () => { clearToken(); setMsTokenState(null); setUserInfo(null); };

  if (!msToken) return <AuthScreen />;
  return <App msToken={msToken} setMsToken={setMsTokenState} userInfo={userInfo} onLogout={handleLogout} />;
}

function App({ msToken, setMsToken, userInfo, onLogout }) {
  const currentUser = userInfo?.userPrincipalName || userInfo?.mail || "usuario";
  const displayName = userInfo?.displayName || currentUser;
  const [tab, setTab] = useState("home");
  const [glucose, setGlucose] = useState("");
  const [foods, setFoods] = useState([]);
  const [recentFoods, setRecentFoods] = useState(() => {
    try { return JSON.parse(localStorage.getItem("glucoapp-recent-foods")||"[]"); } catch { return []; }
  });
  const [search, setSearch] = useState("");
  const [cat, setCat] = useState("Todos");
  const [modal, setModal] = useState(false);
  const [records, setRecords] = useState([]);
  const [customFoods, setCustomFoods] = useState([]);
  const [customFoodsReady, setCustomFoodsReady] = useState(false);
  const [newFood, setNewFood] = useState({ name:"", portion:"", carbs:"", protein:"", kcal:"" });
  const [scanning, setScanning] = useState(false);
  const [scanResult, setScanResult] = useState(null);
  const [scanError, setScanError] = useState("");
  const [editingFood, setEditingFood] = useState(null); // {index, food} for editing
  const [editFood, setEditFood] = useState({name:"",portion:"",carbs:"",protein:"",kcal:""});
  const [xlsxImporting, setXlsxImporting] = useState(false);
  const [xlsxMsg, setXlsxMsg] = useState("");
  const [settings, setSettings] = useState(() => {
    try {
      const saved = localStorage.getItem(`glucoapp-${currentUser}-settings`);
      return saved ? JSON.parse(saved) : DEFAULT_SETTINGS;
    } catch { return DEFAULT_SETTINGS; }
  });
  const [saved, setSaved] = useState(false);
  const [settingsSaved, setSettingsSaved] = useState(false);
  const [withToujeo, setWithToujeo] = useState(false);
  const [newMed, setNewMed] = useState("");
  const [syncing, setSyncing] = useState(false);
  const [syncMsg, setSyncMsg] = useState("");
  const [fileId, setFileId] = useState(null);
  const [expandedDate, setExpandedDate] = useState(null);
  const [odStatus, setOdStatus] = useState(msToken ? "connecting" : "disconnected");
  // odStatus: 'disconnected' | 'connecting' | 'ready' | 'error'
  const [odError, setOdError] = useState("");
  // Load custom foods from local cache on mount (best effort, not source of truth)
  // Load cached records immediately on mount so historial isn't empty while OneDrive connects
  useEffect(() => {
    const cached = loadCachedRecords();
    if (cached.length) setRecords(cached);
  }, []);

  useEffect(() => {
    setCustomFoods(loadCustomFoods());
    setCustomFoodsReady(true);
  }, []);

  useEffect(() => {
    if (customFoodsReady) {
      saveCustomFoods(customFoods);
      if (fileId && odStatus === "ready") saveCustomFoodsToOneDrive(fileId, customFoods);
    }
  }, [customFoods, customFoodsReady]);

  // Token from OAuth popup is handled by msLogin callback — nothing to read from hash
  useEffect(() => {}, []);

  // Connect to OneDrive: find/create GlucoApp.xlsx, ensure "Registros" sheet+headers,
  // then download the full history. This is the SOURCE OF TRUTH for records.
  useEffect(() => {
    if (!msToken) { setOdStatus("disconnected"); return; }
    let cancelled = false;
    (async () => {
      setOdStatus("connecting");
      setSyncMsg("☁️ Conectando a OneDrive...");
      try {
        const { file, created } = await findOrCreateFile();
        if (cancelled) return;
        setFileId(file.id);
        await ensureSheet(file.id);
        setSyncMsg(created ? "📄 GlucoApp.xlsx creado, cargando..." : "📥 Cargando historial...");
        const recs = await readAllRecords(file.id);
        if (cancelled) return;
        setRecords(recs);
        cacheRecords(recs);
        // Load settings from OneDrive (overrides localStorage)
        const cloudSettings = await loadSettingsFromOneDrive(file.id);
        if (cloudSettings) {
          setSettings(cloudSettings);
          try { localStorage.setItem(`glucoapp-${currentUser}-settings`, JSON.stringify(cloudSettings)); } catch {}
        }
        // Load custom foods from OneDrive (overrides localStorage)
        const cloudFoods = await loadCustomFoodsFromOneDrive(file.id);
        if (cloudFoods) {
          setCustomFoods(cloudFoods);
          try { localStorage.setItem(CUSTOM_FOODS_KEY, JSON.stringify(cloudFoods)); } catch {}
        }
        setOdStatus("ready");
        setSyncMsg("✅ Conectado a OneDrive");
        setTimeout(() => setSyncMsg(""), 2500);
      } catch (e) {
        if (cancelled) return;
        setOdStatus("error");
        setOdError(String(e?.message || e));
        setSyncMsg("⚠️ Error conectando a OneDrive");
        // Fall back to whatever was cached locally so the user isn't stuck empty-handed
        const cached = loadCachedRecords(); if (cached.length) setRecords(cached);
      }
    })();
    return () => { cancelled = true; };
  }, [msToken]);

  const allFoods = [...FOODS, ...customFoods];
  const filtered = (() => {
    if (cat === "Recientes") {
      return recentFoods
        .filter(f => f.name.toLowerCase().includes(search.toLowerCase()))
        .sort((a,b) => b.count - a.count || b.last - a.last);
    }
    return allFoods.filter(f =>
      f.name.toLowerCase().includes(search.toLowerCase()) &&
      (cat === "Todos" || f.cat === cat)
    );
  })();

  const totalCarbs   = foods.reduce((s,f) => s + f.carbs*f.qty, 0);
  const totalProtein = foods.reduce((s,f) => s + (f.protein||0)*f.qty, 0);
  const totalKcal    = foods.reduce((s,f) => s + (f.kcal||0)*f.qty, 0);
  const currentRatio = getCurrentRatio(settings.ratios);

  // ── Edad, IMC y TMB (Mifflin-St Jeor) a partir de datos personales ──
  const getAge = () => {
    if (!settings.fechaNacimiento) return null;
    const dob = new Date(settings.fechaNacimiento);
    const now = new Date();
    let age = now.getFullYear() - dob.getFullYear();
    const monthDiff = now.getMonth() - dob.getMonth();
    if (monthDiff < 0 || (monthDiff === 0 && now.getDate() < dob.getDate())) age--;
    return age;
  };

  const imc = (() => {
    const h = settings.alturaCm/100;
    if (!h || !settings.pesoKg) return { value:"-", label:"", color:C.muted };
    const val = settings.pesoKg / (h*h);
    let label, color;
    if (val < 18.5)      { label = "Bajo peso"; color = C.yellow; }
    else if (val < 25)   { label = "Normal";    color = C.green; }
    else if (val < 30)   { label = "Sobrepeso"; color = C.yellow; }
    else                 { label = "Obesidad";  color = C.red; }
    return { value: val.toFixed(1), label, color };
  })();

  const tmb = (() => {
    const age = getAge();
    if (!age || !settings.pesoKg || !settings.alturaCm) return "-";
    // Mifflin-St Jeor
    const base = 10*settings.pesoKg + 6.25*settings.alturaCm - 5*age;
    const val = settings.sexo === "Femenino" ? base - 161 : base + 5;
    return Math.round(val);
  })();

  const calc = () => {
    const g = parseFloat(glucose);
    if (!g && totalCarbs===0) return null;
    const corr = g ? Math.max(0,(g-settings.objetivo)/settings.sensitivity) : 0;
    const meal = totalCarbs/currentRatio;
    return { corr:corr.toFixed(1), meal:meal.toFixed(1), raw:(corr+meal).toFixed(1), total:Math.round(corr+meal), ratio:currentRatio };
  };
  const result = calc();

  // Clasifica la glucemia en 5 niveles, igual que la app de referencia:
  // Hipoglucemia < Glucemia baja < (rango objetivo) < Glucemia alta < Hiperglucemia
  const gLevel = () => {
    const g = parseFloat(glucose);
    if (!g) return null;
    const { hipoglucemia, glucemiaBaja, glucemiaAlta, hiperglucemia } = settings;
    if (g < hipoglucemia) return "hipo";
    if (g < glucemiaBaja) return "baja";
    if (g <= glucemiaAlta) return "normal";
    if (g <= hiperglucemia) return "alta";
    return "hiper";
  };

  const gColor = () => {
    const lvl = gLevel();
    if (!lvl) return C.muted;
    return { hipo:C.red, baja:C.yellow, normal:C.green, alta:C.yellow, hiper:C.red }[lvl];
  };

  const gLabel = () => {
    const lvl = gLevel();
    if (!lvl) return null;
    const map = {
      hipo:   { t:"Hipoglucemia ⚠️", c:C.red },
      baja:   { t:"Glucemia baja",   c:C.yellow },
      normal: { t:"En rango ✓",      c:C.green },
      alta:   { t:"Glucemia alta",   c:C.yellow },
      hiper:  { t:"Hiperglucemia ⚠️", c:C.red },
    };
    return map[lvl];
  };

  const addFood = (f) => {
    // Track usage frequency for Recientes
    setRecentFoods(prev => {
      const key = `${f.name}||${f.portion}`;
      const existing = prev.find(r => r.key===key);
      const updated = existing
        ? prev.map(r => r.key===key ? {...r,count:r.count+1,last:Date.now()} : r)
        : [...prev, {key,name:f.name,portion:f.portion,carbs:f.carbs,protein:f.protein,kcal:f.kcal,cat:f.cat,count:1,last:Date.now()}];
      const sorted = updated.sort((a,b)=>b.count-a.count||b.last-a.last).slice(0,20);
      try { localStorage.setItem("glucoapp-recent-foods", JSON.stringify(sorted)); } catch {}
      return sorted;
    });
    setFoods(prev => {
      const ex = prev.find(x => x.name===f.name && x.portion===f.portion);
      if (ex) return prev.map(x => x.name===f.name && x.portion===f.portion ? {...x,qty:x.qty+1} : x);
      return [...prev, {...f,qty:1}];
    });
  };

  const remFood = (f) => setFoods(prev => {
    const ex = prev.find(x => x.name===f.name && x.portion===f.portion);
    if (ex?.qty>1) return prev.map(x => x.name===f.name && x.portion===f.portion ? {...x,qty:x.qty-1} : x);
    return prev.filter(x => !(x.name===f.name && x.portion===f.portion));
  });

  const doSave = async () => {
    if (!result) return;
    const now = new Date();
    const r = {
      id: Date.now(),
      time:  now.toLocaleTimeString("es-CO",{hour:"2-digit",minute:"2-digit"}),
      date:  now.toLocaleDateString("es-CO"),
      glucose: glucose||"-",
      carbs:   totalCarbs.toFixed(1),
      protein: totalProtein.toFixed(1),
      kcal:    totalKcal.toFixed(0),
      insulin: result.total,
      ratio:   result.ratio,
      foods:   foods.map(f=>`${f.name} (${f.portion}) x${f.qty}`).join(", ")||"-",
      toujeo:  withToujeo ? settings.toujeoDosis : 0,
    };

    if (msToken && fileId && odStatus==="ready") {
      setSyncing(true); setSyncMsg("💾 Guardando en OneDrive...");
      try {
        await appendRow(fileId, [r.date,r.time,r.glucose,r.carbs,r.protein,r.kcal,r.insulin,r.foods,r.toujeo]);
        const updated = [r,...records].slice(0,500);
        setRecords(updated);
        cacheRecords(updated);
        setSyncMsg("✅ Guardado en OneDrive");
        setSyncing(false);
        setTimeout(() => setSyncMsg(""), 2500);
        setSaved(true);
        setTimeout(() => { setGlucose(""); setFoods([]); setWithToujeo(false); setSaved(false); }, 1800);
      } catch (e) {
        setSyncing(false);
        setSyncMsg("⚠️ No se pudo guardar en OneDrive. Intenta de nuevo.");
        // Do NOT clear the form or mark as saved — the user needs to retry.
      }
      return;
    }

    // Not connected to OneDrive: save to session memory directly (no confirm dialog)
    const updated = [r,...records].slice(0,300);
    setRecords(updated);
    cacheRecords(updated);
    setSaved(true);
    setTimeout(() => { setGlucose(""); setFoods([]); setWithToujeo(false); setSaved(false); }, 1800);
  };

  const todayStr  = new Date().toLocaleDateString("es-CO");
  const todayRecs = records.filter(r => r.date===todayStr);
  const todayTot  = dayTotals(todayRecs);
  const grouped   = groupByDate(records);

  const updateRatio = (i,key,val) => {
    setSettings(p => ({...p, ratios: p.ratios.map((r,idx) => idx===i ? {...r,[key]:val} : r)}));
  };

  const scanLabel = async (file) => {
    if (!file) return;
    setScanning(true); setScanResult(null); setScanError("");
    try {
      // Use Tesseract.js loaded via script tag in index.html
      const Tesseract = window.Tesseract;
      if (!Tesseract) {
        setScanError("El escáner no está disponible. Ingresa los datos manualmente.");
        setScanning(false); return;
      }
      const { data: { text } } = await Tesseract.recognize(file, "spa+eng", {
        logger: () => {}
      });

      // Parse nutritional values from OCR text
      const normalize = (s) => s.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"");
      const lines = text.split("\n").map(l => l.trim()).filter(Boolean);
      const fullText = normalize(text);

      // Helper: find number after keyword
      const findVal = (keywords) => {
        for (const kw of keywords) {
          const idx = fullText.indexOf(normalize(kw));
          if (idx === -1) continue;
          const after = fullText.slice(idx + kw.length, idx + kw.length + 30);
          const match = after.match(/[\d]+[.,]?[\d]*/);
          if (match) return parseFloat(match[0].replace(",","."));
        }
        return 0;
      };

      // Find portion size
      const findPortion = () => {
        for (const line of lines) {
          const n = normalize(line);
          if (n.includes("porcion") || n.includes("porción") || n.includes("serving size") || n.includes("tamano")) {
            const match = line.match(/[\d]+\s*[gGmMlL]/);
            if (match) return match[0].trim();
            return line.replace(/.*[:\/]/,"").trim().slice(0,30);
          }
        }
        return "1 porción";
      };

      // Find product name (first non-nutritional line)
      const findName = () => {
        for (const line of lines) {
          const n = normalize(line);
          if (n.length < 3) continue;
          if (n.includes("informacion") || n.includes("nutricional") || n.includes("caloria") || n.includes("serving")) continue;
          if (/^\d/.test(line)) continue;
          return line.slice(0, 40);
        }
        return "Producto escaneado";
      };

      const carbs   = findVal(["carbohidratos totales","total carbohydrates","carbohidratos","carbohydrate","hidratos"]);
      const protein = findVal(["proteina","proteína","protein"]);
      const kcal    = findVal(["calorias","calorías","calories","kcal","energy"]);
      const portion = findPortion();
      const name    = findName();

      if (carbs === 0 && protein === 0 && kcal === 0) {
        setScanError("No se pudieron leer los valores nutricionales. Intenta con mejor iluminación o ingresa los datos manualmente.");
        setScanning(false); return;
      }

      const parsed = { name, portion, carbs, protein, kcal };
      setScanResult(parsed);
      setNewFood({
        name: parsed.name,
        portion: parsed.portion,
        carbs: String(parsed.carbs),
        protein: String(parsed.protein),
        kcal: String(parsed.kcal),
      });
    } catch(e) {
      setScanError("No se pudo leer la etiqueta. Intenta con mejor iluminación o ingresa los datos manualmente.");
    }
    setScanning(false);
  };

  const importFromExcel = async (file) => {
    if (!file) return;
    setXlsxImporting(true); setXlsxMsg("");
    try {
      // Read file as ArrayBuffer
      const buffer = await file.arrayBuffer();
      // Dynamically import SheetJS from CDN
      const XLSX = await import('xlsx');
      const wb = XLSX.read(buffer, { type:"array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval:"" });

      if (rows.length === 0) {
        setXlsxMsg("⚠️ El archivo está vacío o no tiene el formato esperado.");
        setXlsxImporting(false); return;
      }

      // Flexible column detection — acepta headers en español e inglés, con o sin tildes
      const normalize = (s) => String(s).toLowerCase()
        .normalize("NFD").replace(/[\u0300-\u036f]/g,"").trim();

      const findCol = (row, options) => {
        const keys = Object.keys(row);
        for (const opt of options) {
          const match = keys.find(k => normalize(k).includes(opt));
          if (match) return match;
        }
        return null;
      };

      const firstRow = rows[0];
      const colName    = findCol(firstRow, ["nombre","name","alimento","food","producto"]);
      const colPortion = findCol(firstRow, ["porcion","porción","portion","cantidad","amount","unidad"]);
      const colCarbs   = findCol(firstRow, ["carb","ch","hidr","carbohi"]);
      const colProt    = findCol(firstRow, ["prot","protein"]);
      const colKcal    = findCol(firstRow, ["kcal","cal","energia","energ"]);

      if (!colName) {
        setXlsxMsg("⚠️ No encontré columna de nombre. Asegúrate de tener una columna 'Nombre' o 'Alimento'.");
        setXlsxImporting(false); return;
      }

      const toNum = (v) => parseFloat(String(v).replace(",",".")) || 0;

      const imported = [];
      const duplicates = [];

      for (const row of rows) {
        const name = String(row[colName]||"").trim();
        if (!name) continue;
        const portion = colPortion ? String(row[colPortion]||"1 porción").trim() : "1 porción";
        const carbs   = colCarbs ? toNum(row[colCarbs]) : 0;
        const protein = colProt  ? toNum(row[colProt])  : 0;
        const kcal    = colKcal  ? toNum(row[colKcal])  : 0;

        // Check for duplicates against existing custom foods AND FOODS array
        const allExisting = [...FOODS, ...customFoods];
        const isDuplicate = allExisting.some(
          f => f.name.toLowerCase()===name.toLowerCase() &&
               f.portion.toLowerCase()===portion.toLowerCase()
        );
        if (isDuplicate) { duplicates.push(name); continue; }
        imported.push({ name, portion, carbs, protein, kcal, cat:"Personalizados" });
      }

      if (imported.length > 0) {
        setCustomFoods(p => [...p, ...imported]);
      }

      let msg = `✅ ${imported.length} alimento${imported.length!==1?"s":""} importado${imported.length!==1?"s":""}`;
      if (duplicates.length > 0) msg += ` · ${duplicates.length} duplicado${duplicates.length!==1?"s":""} omitido${duplicates.length!==1?"s":""}`;
      setXlsxMsg(msg);
    } catch(e) {
      setXlsxMsg("⚠️ Error leyendo el archivo. Asegúrate de que sea un .xlsx o .csv válido.");
    }
    setXlsxImporting(false);
    setTimeout(() => setXlsxMsg(""), 5000);
  };

  const label = gLabel();

  // ── Stat pill component ──
  const StatRow = ({items}) => (
    <div style={{display:"flex"}}>
      {items.map((x,i) => (
        <div key={i} style={{flex:1,textAlign:"center",borderRight:i<items.length-1?`1px solid ${C.border}`:"none"}}>
          <div style={{fontSize:9,color:C.muted,textTransform:"uppercase",letterSpacing:1,marginBottom:3}}>{x.l}</div>
          <div style={{fontSize:15,fontWeight:700,fontFamily:"monospace",color:x.c}}>{x.v}</div>
        </div>
      ))}
    </div>
  );

  return (
    <div style={{fontFamily:"system-ui,sans-serif",background:C.bg,minHeight:"100vh",color:C.text,maxWidth:430,margin:"0 auto"}}>

      {/* Header */}
      <div style={{background:"#ffffff",borderBottom:"1px solid #e5e7eb",padding:"16px 20px",position:"sticky",top:0,zIndex:30}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <div>
            <div style={{fontSize:10,color:C.blue,fontWeight:700,letterSpacing:2}}>GLUCOAPP</div>
            <div style={{fontSize:17,fontWeight:700,color:C.text}}>{displayName && displayName!=="usuario" ? `Hola, ${displayName.split(" ")[0]} 👋` : "Control Diabetes"}</div>
          </div>
          <div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:4}}>
            <div style={{background:"#eff6ff",border:"0.5px solid #bfdbfe",borderRadius:20,padding:"4px 10px",fontSize:11,fontWeight:600,color:C.blue}}>
              Ratio: {currentRatio}g/U
            </div>
          </div>
        </div>
      </div>

      {/* Acumulado del día — siempre visible, en todas las pestañas */}
      <div style={{background:C.card,borderBottom:`1px solid ${C.border}`,padding:"10px 16px",position:"sticky",top:0,zIndex:25}}>
        <div style={{fontSize:10,color:C.muted,fontWeight:700,textTransform:"uppercase",letterSpacing:1,marginBottom:8}}>📅 Acumulado de hoy</div>
        <div style={{display:"flex"}}>
          {[
            {l:"Carbs",   v:todayTot.carbs.toFixed(0)+"g",   c:C.sky},
            {l:"Proteína",v:todayTot.protein.toFixed(0)+"g", c:C.green},
            {l:"Calorías",v:todayTot.kcal.toFixed(0),        c:C.orange},
            {l:settings.insulinaRapida||"Rápida", v:todayTot.insulin+"U", c:C.purple},
            {l:settings.insulinaLenta||"Lenta",   v:todayTot.toujeo+"U",  c:"#7c3aed"},
          ].map((x,i) => (
            <div key={i} style={{flex:1,textAlign:"center",borderRight:i<4?`1px solid ${C.border}`:"none"}}>
              <div style={{fontSize:9,color:C.muted,textTransform:"uppercase",letterSpacing:1,marginBottom:3}}>{x.l}</div>
              <div style={{fontSize:15,fontWeight:700,fontFamily:"monospace",color:x.c}}>{x.v}</div>
            </div>
          ))}
        </div>
      </div>

      <div style={{padding:"0 0 90px"}}>

        {/* ── HOME ── */}
        {tab==="home" && (
          <div style={{padding:16}}>

            {/* Glucemia */}
            <div style={{background:C.card,borderRadius:20,padding:20,marginBottom:12,border:`1.5px solid ${glucose?gColor()+"60":C.border}`}}>
              <div style={{display:"flex",gap:8,alignItems:"center",marginBottom:10}}>
                <div style={{background:C.blue,color:"white",borderRadius:"50%",width:26,height:26,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:700,fontSize:13}}>1</div>
                <div style={{fontWeight:700}}>¿Cuál es tu glucemia?</div>
              </div>
              <div style={{display:"flex",alignItems:"baseline",gap:8}}>
                <input type="number" value={glucose} onChange={e=>setGlucose(e.target.value)} placeholder="000"
                  style={{flex:1,background:"transparent",border:"none",outline:"none",fontSize:52,fontWeight:700,color:glucose?gColor():"#1e3a5f",fontFamily:"monospace",WebkitAppearance:"none"}} />
                <span style={{color:C.muted,fontSize:16}}>mg/dL</span>
              </div>
              {label && (
                <div style={{display:"inline-flex",alignItems:"center",gap:6,background:label.c+"25",borderRadius:20,padding:"4px 12px",marginTop:4}}>
                  <div style={{width:6,height:6,borderRadius:"50%",background:label.c}} />
                  <span style={{fontSize:13,color:label.c,fontWeight:600}}>{label.t}</span>
                </div>
              )}
            </div>

            {/* Alimentos */}
            <div style={{background:C.card,borderRadius:20,padding:20,marginBottom:12,border:`1.5px solid ${foods.length>0?C.sky+"40":C.border}`}}>
              <div style={{display:"flex",gap:8,alignItems:"center",marginBottom:14}}>
                <div style={{background:C.blue,color:"white",borderRadius:"50%",width:26,height:26,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:700,fontSize:13}}>2</div>
                <div style={{fontWeight:700}}>¿Qué vas a comer?</div>
                {totalCarbs>0 && <div style={{marginLeft:"auto",background:C.sky+"20",borderRadius:20,padding:"3px 10px",fontSize:13,color:C.sky,fontWeight:700}}>{totalCarbs.toFixed(1)}g</div>}
              </div>

              {foods.map((f,i) => (
                <div key={i} style={{display:"flex",alignItems:"center",gap:8,padding:"8px 0",borderBottom:`1px solid ${C.bg}`}}>
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{fontSize:14,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{f.name}</div>
                    <div style={{fontSize:11,marginTop:2}}>
                      <span style={{color:C.sky}}>{(f.carbs*f.qty).toFixed(1)}g C</span>
                      {f.protein>0 && <span style={{color:C.green}}> · {(f.protein*f.qty).toFixed(1)}g P</span>}
                      {f.kcal>0 && <span style={{color:C.orange}}> · {(f.kcal*f.qty).toFixed(0)} kcal</span>}
                    </div>
                  </div>
                  <div style={{display:"flex",alignItems:"center",gap:4}}>
                    <button onClick={()=>remFood(f)} style={{background:C.red+"25",border:"none",color:C.red,borderRadius:8,width:30,height:30,fontSize:20,cursor:"pointer"}}>−</button>
                    <span style={{fontWeight:700,width:22,textAlign:"center"}}>{f.qty}</span>
                    <button onClick={()=>addFood(f)} style={{background:C.green+"25",border:"none",color:C.green,borderRadius:8,width:30,height:30,fontSize:20,cursor:"pointer"}}>+</button>
                  </div>
                </div>
              ))}

              {foods.length>0 && (
                <div style={{marginTop:10,marginBottom:12,background:C.bg,borderRadius:12,padding:12,display:"flex",justifyContent:"space-around"}}>
                  {[
                    {l:"Carbs",   v:totalCarbs.toFixed(1)+"g",   c:C.sky},
                    {l:"Proteína",v:totalProtein.toFixed(1)+"g", c:C.green},
                    {l:"Calorías",v:totalKcal.toFixed(0),        c:C.orange},
                  ].map((x,i) => (
                    <div key={i} style={{textAlign:"center"}}>
                      <div style={{fontSize:9,color:C.muted,textTransform:"uppercase",letterSpacing:1}}>{x.l}</div>
                      <div style={{fontSize:18,fontWeight:700,color:x.c,fontFamily:"monospace"}}>{x.v}</div>
                    </div>
                  ))}
                </div>
              )}

              <button onClick={()=>{ setModal(true); setCat(recentFoods.length>0?"Recientes":"Todos"); }} style={{width:"100%",background:C.bg,border:`1.5px dashed ${C.border}`,color:C.muted,borderRadius:14,padding:"13px 0",fontSize:16,fontWeight:600,cursor:"pointer"}}>
                + Buscar alimento
              </button>
            </div>

            {/* Resultado */}
            {result && !saved && (
              <div style={{background:"#eff6ff",border:"1.5px solid #bfdbfe",borderRadius:20,padding:20,marginBottom:12}}>
                <div style={{fontSize:11,color:C.blue,fontWeight:700,letterSpacing:2,textTransform:"uppercase",marginBottom:8}}>💉 DOSIS APIDRA</div>
                <div style={{display:"flex",alignItems:"baseline",gap:6,marginBottom:4}}>
                  <span style={{fontSize:68,fontWeight:700,color:C.blue,fontFamily:"monospace",lineHeight:1}}>{result.total}</span>
                  <span style={{fontSize:24,color:C.blue,paddingBottom:8}}>U</span>
                </div>
                <div style={{fontSize:12,color:C.muted,marginBottom:16}}>Exacto: {result.raw}U → {result.total}U · Ratio: {result.ratio}g/U</div>
                <div style={{display:"flex",gap:10,marginBottom:16}}>
                  {[
                    {label:"Corrección",val:result.corr+"U",sub:"por glucemia"},
                    {label:"Comida",val:result.meal+"U",sub:`${totalCarbs.toFixed(1)}g ÷ ${result.ratio}`},
                  ].map((x,i) => (
                    <div key={i} style={{flex:1,background:"#dbeafe",borderRadius:14,padding:"10px 14px"}}>
                      <div style={{fontSize:9,color:C.blue,fontWeight:700,textTransform:"uppercase",letterSpacing:1,marginBottom:4}}>{x.label}</div>
                      <div style={{fontSize:22,fontWeight:700,fontFamily:"monospace",color:C.text}}>{x.val}</div>
                      <div style={{fontSize:10,color:C.muted}}>{x.sub}</div>
                    </div>
                  ))}
                </div>
                <button onClick={()=>setWithToujeo(v=>!v)}
                  style={{width:"100%",display:"flex",alignItems:"center",gap:10,background:withToujeo?C.purple+"30":"#0f172a50",border:`1.5px solid ${withToujeo?C.purple:"transparent"}`,borderRadius:14,padding:"12px 14px",marginBottom:14,cursor:"pointer"}}>
                  <div style={{width:22,height:22,borderRadius:6,border:`2px solid ${withToujeo?C.purple:C.muted}`,background:withToujeo?C.purple:"transparent",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
                    {withToujeo && <span style={{color:"white",fontSize:14,fontWeight:900}}>✓</span>}
                  </div>
                  <div style={{textAlign:"left",flex:1}}>
                    <div style={{fontSize:14,fontWeight:700,color:withToujeo?C.purple:C.text}}>💉 También aplicar {settings.insulinaLenta||"insulina lenta"}</div>
                    <div style={{fontSize:11,color:C.muted}}>{settings.toujeoDosis}U basal — registrar junto a esta dosis</div>
                  </div>
                </button>
                <button onClick={doSave} style={{width:"100%",background:C.blue,border:"none",color:"white",borderRadius:14,padding:"14px 0",fontSize:16,fontWeight:700,cursor:"pointer"}}>
                  ✓ Registrar {msToken?"y guardar en OneDrive ☁️":""}
                </button>
              </div>
            )}

            {saved && (
              <div style={{background:"#f0fdf4",border:"1.5px solid #bbf7d0",borderRadius:20,padding:30,textAlign:"center"}}>
                <div style={{fontSize:48}}>✅</div>
                <div style={{fontSize:17,fontWeight:700,color:C.green,marginTop:10}}>
                  {msToken ? "¡Guardado en OneDrive! ☁️" : "¡Guardado!"}
                </div>
              </div>
            )}

            {!result && !saved && (
              <div style={{textAlign:"center",color:C.muted,fontSize:13,padding:8}}>
                Ingresa glucemia y/o alimentos para calcular
              </div>
            )}
          </div>
        )}

        {/* ── ANÁLISIS ── */}
        {tab==="analysis" && (() => {
          // Build last 7 weeks of data grouped by week
          const parseDate = (s) => {
            const [d,m,y] = s.split("/");
            return new Date(parseInt(y), parseInt(m)-1, parseInt(d));
          };
          const getWeekKey = (date) => {
            const d = new Date(date);
            d.setHours(0,0,0,0);
            d.setDate(d.getDate() - d.getDay()); // Sunday
            return d.toLocaleDateString("es-CO");
          };

          // Group records by week
          const weekMap = {};
          for (const r of records) {
            try {
              const wk = getWeekKey(parseDate(r.date));
              if (!weekMap[wk]) weekMap[wk] = [];
              weekMap[wk].push(r);
            } catch {}
          }

          // Get last 8 weeks sorted
          const weeks = Object.entries(weekMap)
            .sort((a,b) => new Date(a[0]) - new Date(b[0]))
            .slice(-8);

          // Calculate weekly averages
          const weeklyData = weeks.map(([wk, recs]) => {
            const days = [...new Set(recs.map(r => r.date))].length;
            return {
              label: wk.slice(0,5), // "DD/MM"
              carbs:   Math.round(recs.reduce((s,r)=>s+(parseFloat(r.carbs)||0),0) / days),
              protein: Math.round(recs.reduce((s,r)=>s+(parseFloat(r.protein)||0),0) / days),
              kcal:    Math.round(recs.reduce((s,r)=>s+(parseFloat(r.kcal)||0),0) / days),
              insulin: Math.round((recs.reduce((s,r)=>s+(r.insulin||0),0) / days) * 10) / 10,
              days,
            };
          });

          const metaCarbs   = settings.metaCarbs   || 130;
          const metaProtein = settings.metaProtein || 85;
          const metaKcal    = settings.metaKcal    || 1350;
          const pesoMeta    = settings.pesoMeta    || 50;
          const pesoActual  = settings.pesoKg      || 55;

          // Progress bar component - without %
          const ProgressBar = ({ label, value, meta, color, unit="" }) => {
            const pct = Math.min(Math.round((value/meta)*100), 100);
            const over = value > meta;
            return (
              <div style={{marginBottom:14}}>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:5}}>
                  <span style={{fontSize:13,color:C.text,fontWeight:600}}>{label}</span>
                  <span style={{fontSize:12,color:C.muted}}>{value}{unit} / {meta}{unit}</span>
                </div>
                <div style={{height:10,background:"#f1f5f9",borderRadius:99,overflow:"hidden"}}>
                  <div style={{height:"100%",width:`${pct}%`,background:over?"#f97316":color,borderRadius:99,transition:"width 0.5s"}} />
                </div>
              </div>
            );
          };

          return (
            <div style={{padding:16,paddingBottom:80}}>
              <div style={{fontSize:18,fontWeight:700,marginBottom:4}}>📈 Análisis</div>
              <div style={{fontSize:12,color:C.muted,marginBottom:16}}>Hoy · {todayStr}</div>

              {/* Daily donut charts */}
              <div style={{background:C.card,borderRadius:16,padding:16,marginBottom:12}}>
                <div style={{fontSize:13,fontWeight:700,color:C.muted,marginBottom:16}}>🍽️ MACROS DE HOY</div>
                <div style={{display:"flex",justifyContent:"space-around",alignItems:"flex-start"}}>
                  {[
                    {label:"Carbohidratos", value:todayTot.carbs,   meta:metaCarbs,   color:"#38bdf8", unit:"g"},
                    {label:"Proteína",      value:todayTot.protein, meta:metaProtein, color:"#f97316", unit:"g"},
                    {label:"Calorías",      value:todayTot.kcal,    meta:metaKcal,    color:"#16a34a", unit:""},
                  ].map(({label, value, meta, color, unit}) => {
                    const pct = Math.min((value/meta), 1);
                    const r = 34; const circ = 2*Math.PI*r;
                    const dash = pct*circ;
                    const over = value > meta;
                    const remaining = Math.max(meta - value, 0);
                    return (
                      <div key={label} style={{textAlign:"center",flex:1}}>
                        <div style={{fontSize:12,fontWeight:700,color,marginBottom:8}}>{label}</div>
                        <div style={{position:"relative",display:"inline-block"}}>
                          <svg width="88" height="88" viewBox="0 0 88 88">
                            <circle cx="44" cy="44" r={r} fill="none" stroke="#f1f5f9" strokeWidth="8"/>
                            <circle cx="44" cy="44" r={r} fill="none"
                              stroke={color} strokeWidth="8"
                              strokeDasharray={`${dash} ${circ}`}
                              strokeLinecap="round"
                              transform="rotate(-90 44 44)"
                            />
                          </svg>
                          <div style={{position:"absolute",top:"50%",left:"50%",transform:"translate(-50%,-50%)",textAlign:"center",lineHeight:1.1}}>
                            <div style={{fontSize:18,fontWeight:700,color:C.text}}>{value}</div>
                            <div style={{fontSize:10,color:C.muted}}>/{meta}{unit}</div>
                          </div>
                        </div>
                        <div style={{fontSize:11,color:C.muted,marginTop:6,fontWeight:500}}>
                          {over ? `+${Math.round(value-meta)}${unit} extra` : `${Math.round(remaining)}${unit} restante`}
                        </div>
                        <div style={{fontSize:11,color,fontWeight:700,marginTop:2}}>{Math.min(Math.round((value/meta)*100),100)}%</div>
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* Macro distribution pie chart */}
              {(todayTot.carbs > 0 || todayTot.protein > 0) && (() => {
                const total = todayTot.carbs + todayTot.protein;
                const carbsPct = total > 0 ? Math.round((todayTot.carbs/total)*100) : 0;
                const protPct  = total > 0 ? Math.round((todayTot.protein/total)*100) : 0;
                const cx = 60, cy = 60, r = 55;
                const slice = (start, pct, color) => {
                  if (pct <= 0) return null;
                  if (pct >= 100) return <circle key={color} cx={cx} cy={cy} r={r} fill={color}/>;
                  const a1 = (start/100)*2*Math.PI - Math.PI/2;
                  const a2 = ((start+pct)/100)*2*Math.PI - Math.PI/2;
                  const x1 = cx+r*Math.cos(a1), y1 = cy+r*Math.sin(a1);
                  const x2 = cx+r*Math.cos(a2), y2 = cy+r*Math.sin(a2);
                  const large = pct > 50 ? 1 : 0;
                  const mid = ((start+pct/2)/100)*2*Math.PI - Math.PI/2;
                  const lx = cx+(r*0.65)*Math.cos(mid), ly = cy+(r*0.65)*Math.sin(mid);
                  return (
                    <g key={color}>
                      <path d={`M${cx},${cy} L${x1},${y1} A${r},${r} 0 ${large},1 ${x2},${y2} Z`} fill={color}/>
                      {pct >= 8 && <text x={lx} y={ly} textAnchor="middle" dominantBaseline="middle" fontSize="11" fontWeight="700" fill="white">{pct}%</text>}
                    </g>
                  );
                };
                return (
                  <div style={{background:C.card,borderRadius:16,padding:16,marginBottom:12}}>
                    <div style={{fontSize:13,fontWeight:700,color:C.muted,marginBottom:16}}>🥗 DISTRIBUCIÓN DE MACROS</div>
                    <div style={{display:"flex",alignItems:"center",gap:20}}>
                      <svg width="120" height="120" viewBox="0 0 120 120">
                        {slice(0, carbsPct, "#38bdf8")}
                        {slice(carbsPct, protPct, "#f97316")}
                      </svg>
                      <div style={{flex:1}}>
                        {[
                          {label:"Carbohidratos", g:todayTot.carbs,   pct:carbsPct, color:"#38bdf8"},
                          {label:"Proteína",      g:todayTot.protein, pct:protPct,  color:"#f97316"},
                        ].map(({label,g,pct,color})=>(
                          <div key={label} style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12}}>
                            <div style={{display:"flex",alignItems:"center",gap:8}}>
                              <div style={{width:12,height:12,borderRadius:3,background:color,flexShrink:0}}/>
                              <div style={{fontSize:12,color:C.text}}>{label} ({g}g)</div>
                            </div>
                            <div style={{fontSize:13,fontWeight:700,color}}>{pct}%</div>
                          </div>
                        ))}
                      </div>
                    </div>
                  </div>
                );
              })()}

              {weeklyData.length === 0 ? (
                <div style={{textAlign:"center",padding:40,color:C.muted}}>
                  <div style={{fontSize:40,marginBottom:12}}>📊</div>
                  <div style={{fontSize:15,fontWeight:600}}>Sin datos aún</div>
                  <div style={{fontSize:13,marginTop:6}}>Registra tus comidas e insulina para ver el análisis</div>
                </div>
              ) : (
                <>
                  {/* Macros semanales — barras de progreso */}
                  <div style={{background:C.card,borderRadius:16,padding:16,marginBottom:12}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
                      <div style={{fontSize:13,fontWeight:700,color:C.muted}}>📊 PROMEDIO SEMANAL</div>
                      <div style={{fontSize:11,color:C.muted}}>{weeklyData.length>0 ? weeklyData[weeklyData.length-1].label : ""}</div>
                    </div>
                    {weeklyData.length > 0 && (() => {
                      const last = weeklyData[weeklyData.length-1];
                      return (
                        <>
                          <ProgressBar label="Carbohidratos" value={last.carbs} meta={metaCarbs} color={C.sky} unit="g" />
                          <ProgressBar label="Proteína" value={last.protein} meta={metaProtein} color={C.green} unit="g" />
                          <ProgressBar label="Calorías" value={last.kcal} meta={metaKcal} color={C.orange} unit="" />
                          <div style={{borderTop:`1px solid ${C.border}`,marginTop:4,paddingTop:14}}>
                            <div style={{fontSize:12,fontWeight:700,color:C.muted,marginBottom:10}}>💉 {settings.insulinaRapida||"Insulina rápida"}</div>
                            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8}}>
                              {(() => {
                                const allInsulin = weeklyData.map(w=>w.insulin);
                                const lastWeekRecords = records.filter(r => {
                                  try { return getWeekKey(parseDate(r.date)) === weeks[weeks.length-1][0]; } catch { return false; }
                                });
                                const weekTotal = Math.round(lastWeekRecords.reduce((s,r)=>s+(r.insulin||0),0));
                                return [
                                  {label:"Promedio", value: Math.round(last.insulin*10)/10+"U"},
                                  {label:"Mínimo",   value: Math.round(Math.min(...allInsulin)*10)/10+"U"},
                                  {label:"Máximo",   value: Math.round(Math.max(...allInsulin)*10)/10+"U"},
                                  {label:"Total sem",value: weekTotal+"U"},
                                ].map(({label,value})=>(
                                  <div key={label} style={{background:"#f8fafc",borderRadius:10,padding:"10px 8px",textAlign:"center"}}>
                                    <div style={{fontSize:10,color:C.muted,marginBottom:4}}>{label}</div>
                                    <div style={{fontSize:18,fontWeight:700,color:C.purple}}>{value}</div>
                                  </div>
                                ));
                              })()}
                            </div>
                          </div>
                        </>
                      );
                    })()}
                  </div>

                  {/* Historial por semanas */}
                  {weeklyData.length > 1 && (
                    <div style={{background:C.card,borderRadius:16,padding:16,marginBottom:12}}>
                      <div style={{fontSize:13,fontWeight:700,color:C.muted,marginBottom:14}}>📅 HISTORIAL SEMANAL</div>
                      {weeklyData.map((w,i) => (
                        <div key={i} style={{marginBottom:16}}>
                          <div style={{fontSize:12,fontWeight:600,color:C.text,marginBottom:8}}>Semana del {w.label} <span style={{color:C.muted,fontWeight:400}}>· {w.days} día{w.days!==1?"s":""}</span></div>
                          <ProgressBar label="Carbs" value={w.carbs} meta={metaCarbs} color={C.sky} unit="g" />
                          <ProgressBar label="Proteína" value={w.protein} meta={metaProtein} color={C.green} unit="g" />
                          <ProgressBar label="Calorías" value={w.kcal} meta={metaKcal} color={C.orange} unit="" />
                        </div>
                      ))}
                    </div>
                  )}

                  {/* Peso — al final */}
                  <div style={{background:C.card,borderRadius:16,padding:16,marginBottom:12}}>
                    <div style={{fontSize:13,fontWeight:700,color:C.muted,marginBottom:12}}>⚖️ PESO</div>
                    <div style={{display:"flex",gap:12}}>
                      <div style={{flex:1,background:"#f8fafc",borderRadius:12,padding:12,textAlign:"center"}}>
                        <div style={{fontSize:11,color:C.muted,marginBottom:4}}>Actual</div>
                        <div style={{fontSize:24,fontWeight:700,color:C.blue}}>{pesoActual}<span style={{fontSize:12}}> kg</span></div>
                      </div>
                      <div style={{flex:1,background:"#f0fdf4",borderRadius:12,padding:12,textAlign:"center"}}>
                        <div style={{fontSize:11,color:C.muted,marginBottom:4}}>Meta</div>
                        <div style={{fontSize:24,fontWeight:700,color:C.green}}>{pesoMeta}<span style={{fontSize:12}}> kg</span></div>
                      </div>
                      <div style={{flex:1,background:"#fff7ed",borderRadius:12,padding:12,textAlign:"center"}}>
                        <div style={{fontSize:11,color:C.muted,marginBottom:4}}>Diferencia</div>
                        <div style={{fontSize:24,fontWeight:700,color:pesoActual>pesoMeta?C.orange:C.green}}>
                          {pesoActual>pesoMeta?"+":""}{(pesoActual-pesoMeta).toFixed(1)}<span style={{fontSize:12}}> kg</span>
                        </div>
                      </div>
                    </div>
                  </div>

                  <div style={{background:"#fffbeb",border:"1px solid #fde68a",borderRadius:12,padding:12,fontSize:12,color:"#92400e"}}>
                    ⚠️ Las metas son orientativas. Ajústalas según indicación de tu nutricionista en ⚙️ Config.
                  </div>
                </>
              )}
            </div>
          );
        })()}

        {/* ── HISTORIAL ── */}
        {tab==="records" && (
          <div style={{padding:16}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
              <div style={{fontSize:18,fontWeight:700}}>📊 Historial</div>
              {records.length>0 && (
                <button onClick={()=>{ if(window.confirm("¿Borrar todo el historial?")){ setRecords([]); }}}
                  style={{background:C.red+"20",border:"none",color:C.red,borderRadius:10,padding:"5px 12px",fontSize:12,cursor:"pointer"}}>
                  Borrar todo
                </button>
              )}
            </div>

            {records.length===0 ? (
              <div style={{textAlign:"center",color:C.muted,padding:60}}>
                <div style={{fontSize:48}}>📋</div>
                <div style={{marginTop:12}}>Sin registros aún</div>
              </div>
            ) : grouped.map(([date, recs]) => {
              const tot = dayTotals(recs);
              const isToday = date === todayStr;
              const isOpen = expandedDate === date;

              return (
                <div key={date} style={{marginBottom:12}}>
                  {/* Day header — always visible */}
                  <button
                    onClick={() => setExpandedDate(isOpen ? null : date)}
                    style={{width:"100%",background:isToday?"#eff6ff":C.card,border:`1px solid ${isToday?C.blue+"60":C.border}`,borderRadius:16,padding:16,cursor:"pointer",textAlign:"left",marginBottom:isOpen?2:0}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
                      <div>
                        <div style={{fontSize:13,fontWeight:700,color:isToday?C.sky:C.text}}>
                          {isToday ? "📅 Hoy" : "📅 "+date}
                        </div>
                        <div style={{fontSize:11,color:C.muted,marginTop:2}}>{recs.length} registro{recs.length!==1?"s":""}</div>
                      </div>
                      <div style={{display:"flex",alignItems:"center",gap:10}}>
                        <div style={{textAlign:"right"}}>
                          <div style={{fontSize:11,color:C.muted}}>{settings.insulinaRapida||"Rápida"} · {settings.insulinaLenta||"Lenta"}</div>
                          <div style={{fontSize:18,fontWeight:700,color:C.purple,fontFamily:"monospace"}}>{tot.insulin}U <span style={{color:"#7c3aed",fontSize:14}}>· {tot.toujeo}U</span></div>
                        </div>
                        <div style={{fontSize:16,color:C.muted}}>{isOpen?"▲":"▼"}</div>
                      </div>
                    </div>
                    {/* Day macro summary */}
                    <div style={{display:"flex",gap:8}}>
                      {[
                        {l:"🌾 Carbs",   v:tot.carbs.toFixed(0)+"g",   c:C.sky},
                        {l:"💪 Proteína",v:tot.protein.toFixed(0)+"g", c:C.green},
                        {l:"🔥 Calorías",v:tot.kcal.toFixed(0)+" kcal",c:C.orange},
                      ].map((x,i) => (
                        <div key={i} style={{flex:1,background:"#f1f5f9",borderRadius:10,padding:"8px 6px",textAlign:"center"}}>
                          <div style={{fontSize:9,color:C.muted,marginBottom:3}}>{x.l}</div>
                          <div style={{fontSize:14,fontWeight:700,fontFamily:"monospace",color:x.c}}>{x.v}</div>
                        </div>
                      ))}
                    </div>
                  </button>

                  {/* Individual records — visible when expanded */}
                  {isOpen && (
                    <div style={{background:C.bg,border:`1px solid ${C.border}`,borderTop:"none",borderRadius:"0 0 16px 16px",padding:"8px 12px 12px"}}>
                      {recs.map((r,i) => (
                        <div key={i} style={{background:C.card,borderRadius:12,padding:12,marginTop:8}}>
                          <div style={{display:"flex",justifyContent:"space-between",marginBottom:8}}>
                            <span style={{fontSize:12,color:C.muted}}>🕐 {r.time}</span>
                            <span style={{fontSize:15,fontWeight:700,color:C.sky,fontFamily:"monospace"}}>
                              {r.insulin}U {settings.insulinaRapida||"Rápida"}{r.toujeo>0 && <span style={{color:"#7c3aed"}}> · {r.toujeo}U {settings.insulinaLenta||"Lenta"}</span>}
                            </span>
                          </div>
                          <div style={{display:"flex",gap:6,marginBottom:r.foods!=="-"?8:0}}>
                            {[
                              {l:"Glucemia",v:r.glucose,        c:C.text},
                              {l:"Carbs",   v:r.carbs+"g",      c:C.sky},
                              {l:"Prot",    v:(r.protein||"0")+"g",c:C.green},
                              {l:"kcal",    v:r.kcal||"0",      c:C.orange},
                            ].map((x,j) => (
                              <div key={j} style={{flex:1,background:C.bg,borderRadius:8,padding:"5px 0",textAlign:"center"}}>
                                <div style={{fontSize:8,color:C.muted,marginBottom:2}}>{x.l}</div>
                                <div style={{fontSize:12,fontWeight:700,fontFamily:"monospace",color:x.c}}>{x.v}</div>
                              </div>
                            ))}
                          </div>
                          {r.foods!=="-" && (
                            <div style={{fontSize:11,color:C.muted,lineHeight:1.6,marginTop:4}}>🍽️ {r.foods}</div>
                          )}
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        )}

        {/* ── ALIMENTOS ── */}
        {tab==="foodsPage" && (
          <div style={{padding:16}}>
            <div style={{fontSize:18,fontWeight:700,marginBottom:16}}>➕ Mis alimentos</div>
            {customFoods.length>0 && (
              <div style={{marginBottom:20}}>
                <div style={{fontSize:14,fontWeight:700,color:C.muted,marginBottom:12}}>Guardados permanentemente ({customFoods.length})</div>
                {customFoods.map((f,i) => (
                  <div key={i} style={{background:C.card,borderRadius:12,padding:"12px 14px",marginBottom:8}}>
                    {editingFood?.index===i ? (
                      // Edit mode
                      <div>
                        <input value={editFood.name} onChange={e=>setEditFood(p=>({...p,name:e.target.value}))}
                          placeholder="Nombre" style={{...inp,marginBottom:6,fontSize:13}} />
                        <input value={editFood.portion} onChange={e=>setEditFood(p=>({...p,portion:e.target.value}))}
                          placeholder="Porción" style={{...inp,marginBottom:6,fontSize:13}} />
                        <div style={{display:"flex",gap:6,marginBottom:8}}>
                          <input type="number" value={editFood.carbs} onChange={e=>setEditFood(p=>({...p,carbs:e.target.value}))}
                            placeholder="Carbs" style={{...inp,flex:1,fontSize:13}} />
                          <input type="number" value={editFood.protein} onChange={e=>setEditFood(p=>({...p,protein:e.target.value}))}
                            placeholder="Prot" style={{...inp,flex:1,fontSize:13}} />
                          <input type="number" value={editFood.kcal} onChange={e=>setEditFood(p=>({...p,kcal:e.target.value}))}
                            placeholder="Kcal" style={{...inp,flex:1,fontSize:13}} />
                        </div>
                        <div style={{display:"flex",gap:8}}>
                          <button onClick={()=>{
                            setCustomFoods(p=>p.map((x,j)=>j===i?{...x,...editFood,carbs:parseFloat(editFood.carbs)||0,protein:parseFloat(editFood.protein)||0,kcal:parseFloat(editFood.kcal)||0}:x));
                            setEditingFood(null);
                          }} style={{flex:1,background:C.green,border:"none",color:"white",borderRadius:10,padding:"8px 0",fontSize:13,fontWeight:700,cursor:"pointer"}}>
                            ✓ Guardar
                          </button>
                          <button onClick={()=>setEditingFood(null)}
                            style={{flex:1,background:"#f1f5f9",border:"none",color:C.muted,borderRadius:10,padding:"8px 0",fontSize:13,cursor:"pointer"}}>
                            Cancelar
                          </button>
                        </div>
                      </div>
                    ) : (
                      // View mode
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                        <div>
                          <div style={{fontSize:14,fontWeight:600}}>{f.name}</div>
                          <div style={{fontSize:12,color:C.muted}}>{f.portion}</div>
                          <div style={{fontSize:12,marginTop:2}}>
                            <span style={{color:C.sky}}>{f.carbs}g C</span>
                            {f.protein>0 && <span style={{color:C.green}}> · {f.protein}g P</span>}
                            {f.kcal>0 && <span style={{color:C.orange}}> · {f.kcal} kcal</span>}
                          </div>
                        </div>
                        <div style={{display:"flex",gap:6}}>
                          <button onClick={()=>{setEditingFood({index:i,food:f});setEditFood({name:f.name,portion:f.portion,carbs:String(f.carbs),protein:String(f.protein),kcal:String(f.kcal)});}}
                            style={{background:"#eff6ff",border:"none",color:C.blue,borderRadius:8,padding:"6px 10px",fontSize:13,cursor:"pointer"}}>✏️</button>
                          <button onClick={()=>{ if(window.confirm(`¿Borrar "${f.name}"?`)) setCustomFoods(p=>p.filter((_,j)=>j!==i)); }}
                            style={{background:C.red+"20",border:"none",color:C.red,borderRadius:8,padding:"6px 10px",fontSize:13,cursor:"pointer"}}>🗑️</button>
                        </div>
                      </div>
                    )}
                  </div>
                ))}
              </div>
            )}
            <div style={{background:C.card,borderRadius:16,padding:16}}>
              <div style={{fontSize:14,fontWeight:700,marginBottom:4}}>Agregar alimento</div>
              <div style={{fontSize:12,color:C.muted,marginBottom:14}}>📷 Escanea la etiqueta o ingrésalo manualmente</div>

              {/* Scanner button */}
              <label style={{display:"block",marginBottom:14,cursor:"pointer"}}>
                <div style={{
                  background: scanning ? C.bg : "linear-gradient(135deg,#1d4ed8,#0369a1)",
                  borderRadius:14, padding:"14px 0", textAlign:"center",
                  fontSize:15, fontWeight:700, color:"white",
                  opacity: scanning ? 0.7 : 1,
                  display:"flex", alignItems:"center", justifyContent:"center", gap:8
                }}>
                  {scanning ? (
                    <>⏳ Analizando etiqueta...</>
                  ) : (
                    <>📷 Escanear etiqueta nutricional</>
                  )}
                </div>
                <input type="file" accept="image/*" capture="environment"
                  style={{display:"none"}}
                  onChange={e=>{ if(e.target.files[0]) scanLabel(e.target.files[0]); e.target.value=""; }}
                  disabled={scanning} />
              </label>

              {/* Scan error */}
              {scanError && (
                <div style={{background:C.red+"20",borderRadius:10,padding:"10px 12px",marginBottom:12,fontSize:12,color:C.red}}>
                  ⚠️ {scanError}
                </div>
              )}

              {/* Scan result preview */}
              {scanResult && !scanning && (
                <div style={{background:"#f0fdf4",borderRadius:12,padding:12,marginBottom:14,border:`1px solid ${C.green}40`}}>
                  <div style={{fontSize:11,color:C.green,fontWeight:700,textTransform:"uppercase",letterSpacing:1,marginBottom:8}}>✅ Datos detectados — revisa y confirma</div>
                  <div style={{fontSize:13,color:C.text,lineHeight:1.8}}>
                    <div><span style={{color:C.muted}}>Producto:</span> {scanResult.name}</div>
                    <div><span style={{color:C.muted}}>Porción:</span> {scanResult.portion}</div>
                    <div>
                      <span style={{color:C.sky}}>{scanResult.carbs}g C</span>
                      {" · "}<span style={{color:C.green}}>{scanResult.protein}g P</span>
                      {" · "}<span style={{color:C.orange}}>{scanResult.kcal} kcal</span>
                    </div>
                  </div>
                </div>
              )}

              {/* Manual fields (pre-filled by scanner or editable) */}
              <div style={{fontSize:12,color:C.muted,marginBottom:10}}>
                {scanResult ? "✏️ Edita si es necesario:" : "O ingresa manualmente:"}
              </div>
              {[
                {key:"name",placeholder:"Nombre del alimento",type:"text"},
                {key:"portion",placeholder:"Porción (ej: 1 unidad 45g)",type:"text"},
                {key:"carbs",placeholder:"Carbohidratos (g)",type:"number"},
                {key:"protein",placeholder:"Proteínas (g)",type:"number"},
                {key:"kcal",placeholder:"Calorías (kcal)",type:"number"},
              ].map(f => (
                <input key={f.key} type={f.type} value={newFood[f.key]}
                  onChange={e=>setNewFood(p=>({...p,[f.key]:e.target.value}))}
                  placeholder={f.placeholder}
                  style={{...inp,marginBottom:10}} />
              ))}
              <button onClick={()=>{
                if(!newFood.name||!newFood.carbs) return;
                setCustomFoods(p=>[...p,{...newFood,carbs:parseFloat(newFood.carbs)||0,protein:parseFloat(newFood.protein)||0,kcal:parseFloat(newFood.kcal)||0,cat:"Personalizados"}]);
                setNewFood({name:"",portion:"",carbs:"",protein:"",kcal:""});
                setScanResult(null); setScanError("");
              }} style={{width:"100%",background:C.blue,border:"none",color:"white",borderRadius:12,padding:"14px 0",fontSize:16,fontWeight:700,cursor:"pointer"}}>
                Agregar ✓
              </button>
            </div>

            <div style={{background:C.card,borderRadius:16,padding:16,marginTop:12}}>
              <div style={{fontSize:14,fontWeight:700,marginBottom:4}}>📊 Importar desde Excel</div>
              <div style={{fontSize:12,color:C.muted,marginBottom:12}}>
                Sube un archivo <strong>.xlsx</strong> o <strong>.csv</strong> con columnas: <span style={{color:C.blue}}>Nombre · Porción · Carbs</span> (Proteína y Calorías opcionales). Los duplicados se omiten automáticamente.
              </div>
              <label style={{display:"block",cursor:"pointer"}}>
                <div style={{
                  background:xlsxImporting?"#f1f5f9":C.bg,
                  border:`1.5px dashed ${C.border}`,
                  borderRadius:12, padding:"14px 0", textAlign:"center",
                  fontSize:14, fontWeight:600, color:xlsxImporting?C.muted:C.blue,
                  display:"flex", alignItems:"center", justifyContent:"center", gap:8
                }}>
                  {xlsxImporting ? "⏳ Importando..." : "📂 Seleccionar archivo Excel o CSV"}
                </div>
                <input type="file" accept=".xlsx,.xls,.csv"
                  style={{display:"none"}}
                  onChange={e=>{ if(e.target.files[0]) importFromExcel(e.target.files[0]); e.target.value=""; }}
                  disabled={xlsxImporting} />
              </label>
              {xlsxMsg && (
                <div style={{
                  marginTop:10, padding:"10px 12px", borderRadius:10, fontSize:13, fontWeight:600,
                  background: xlsxMsg.startsWith("✅") ? "#f0fdf4" : "#fff7ed",
                  color: xlsxMsg.startsWith("✅") ? C.green : C.orange,
                  border: `1px solid ${xlsxMsg.startsWith("✅") ? "#bbf7d0" : "#fed7aa"}`
                }}>
                  {xlsxMsg}
                </div>
              )}
              <div style={{marginTop:12,background:"#f8fafc",borderRadius:10,padding:10}}>
                <div style={{fontSize:11,color:C.muted,fontWeight:600,marginBottom:6}}>Ejemplo de formato:</div>
                <table style={{width:"100%",fontSize:11,borderCollapse:"collapse"}}>
                  <thead>
                    <tr style={{background:"#eff6ff"}}>
                      {["Nombre","Porción","Carbs","Proteína","Calorías"].map(h=>(
                        <th key={h} style={{padding:"4px 6px",textAlign:"left",color:C.blue,fontWeight:600,border:`1px solid ${C.border}`}}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {[
                      ["Arroz blanco","1 taza","45","5","200"],
                      ["Manzana","1 unidad","15","0","65"],
                      ["Pechuga pollo","100g","0","31","165"],
                    ].map((row,i)=>(
                      <tr key={i} style={{background:i%2===0?"#ffffff":"#f8fafc"}}>
                        {row.map((cell,j)=>(
                          <td key={j} style={{padding:"4px 6px",color:C.text,border:`1px solid ${C.border}`}}>{cell}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* ── CONFIG ── */}
        {tab==="settings" && (
          <div style={{padding:16}}>
            <div style={{fontSize:18,fontWeight:700,marginBottom:16}}>⚙️ Configuración</div>

            <div style={{background:C.card,borderRadius:16,padding:16,marginBottom:14}}>
              <div style={{fontSize:12,color:C.muted,fontWeight:700,textTransform:"uppercase",letterSpacing:1,marginBottom:12}}>☁️ OneDrive</div>
              {msToken ? (
                <div>
                  <div style={{fontSize:14,color:C.green,fontWeight:600,marginBottom:8}}>✅ Conectado</div>
                  <div style={{fontSize:12,color:C.muted}}>Los registros se guardan automáticamente en GlucoApp.xlsx</div>
                  <button onClick={()=>{clearToken();setMsToken(null);setFileId(null);setOdStatus("disconnected");setRecords([]);onLogout();}}
                    style={{marginTop:12,background:C.red+"20",border:"none",color:C.red,borderRadius:10,padding:"6px 14px",fontSize:12,cursor:"pointer"}}>
                    Desconectar
                  </button>
                </div>
              ) : (
                <div>
                  <div style={{fontSize:14,color:C.muted,marginBottom:12}}>Conecta tu OneDrive para guardar los registros permanentemente</div>
                  <button onClick={()=>{ const f = async()=>{ const v=generateCodeVerifier(); const c=await generateCodeChallenge(v); sessionStorage.setItem("pkce_verifier",v); const p=new URLSearchParams({client_id:CLIENT_ID,response_type:"code",redirect_uri:REDIRECT_URI,scope:SCOPES.join(" "),code_challenge:c,code_challenge_method:"S256",response_mode:"query"}); window.location.href=`https://login.microsoftonline.com/common/oauth2/v2.0/authorize?${p}`;}; f(); }} style={{width:"100%",background:C.blue,border:"none",color:"white",borderRadius:12,padding:"12px 0",fontSize:15,fontWeight:700,cursor:"pointer"}}>
                    Conectar OneDrive ☁️
                  </button>
                </div>
              )}
            </div>

            <div style={{background:C.card,borderRadius:20,padding:20,marginBottom:14}}>
              <div style={{fontSize:12,color:C.muted,fontWeight:700,textTransform:"uppercase",letterSpacing:1,marginBottom:16}}>Ratio I:C por horario</div>
              {settings.ratios.map((r,i) => (
                <div key={i} style={{marginBottom:i<2?16:0,background:C.bg,borderRadius:14,padding:14}}>
                  <div style={{fontSize:14,fontWeight:700,marginBottom:12}}>{r.label}</div>
                  <div style={{display:"flex",gap:8,marginBottom:10}}>
                    <div style={{flex:1}}>
                      <div style={{fontSize:11,color:C.muted,marginBottom:6}}>Desde</div>
                      <input type="time" value={r.from} onChange={e=>updateRatio(i,"from",e.target.value)} style={{...inp,padding:"10px 12px"}} />
                    </div>
                    <div style={{flex:1}}>
                      <div style={{fontSize:11,color:C.muted,marginBottom:6}}>Hasta</div>
                      <input type="time" value={r.to} onChange={e=>updateRatio(i,"to",e.target.value)} style={{...inp,padding:"10px 12px"}} />
                    </div>
                  </div>
                  <div style={{fontSize:11,color:C.muted,marginBottom:6}}>Ratio (g CH por 1U)</div>
                  <input type="number" value={r.ratio} onChange={e=>updateRatio(i,"ratio",parseFloat(e.target.value))}
                    style={{...inp,color:C.sky,fontSize:24,fontWeight:700,fontFamily:"monospace"}} />
                </div>
              ))}
            </div>

            <div style={{background:C.card,borderRadius:20,padding:20,marginBottom:14}}>
              <div style={{fontSize:12,color:C.muted,fontWeight:700,textTransform:"uppercase",letterSpacing:1,marginBottom:16}}>Sensibilidad</div>
              <div style={{fontSize:14,fontWeight:600,marginBottom:8}}>Factor de sensibilidad (mg/dL por 1U)</div>
              <input type="number" value={settings.sensitivity}
                onChange={e=>setSettings(p=>({...p,sensitivity:parseFloat(e.target.value)}))}
                style={{...inp,color:C.sky,fontSize:24,fontWeight:700,fontFamily:"monospace"}} />
            </div>

            <div style={{background:C.card,borderRadius:20,padding:20,marginBottom:14}}>
              <div style={{fontSize:12,color:C.muted,fontWeight:700,textTransform:"uppercase",letterSpacing:1,marginBottom:6}}>Rangos de glucemia (mg/dL)</div>
              <div style={{fontSize:11,color:C.muted,marginBottom:16}}>De más bajo a más alto: Hipoglucemia → Glucemia baja → Objetivo → Glucemia alta → Hiperglucemia</div>
              {[
                {key:"hipoglucemia",  label:"🔴 Hipoglucemia",   c:C.red},
                {key:"glucemiaBaja",  label:"🟡 Glucemia baja",  c:C.yellow},
                {key:"objetivo",      label:"🟢 Objetivo",       c:C.green},
                {key:"glucemiaAlta",  label:"🟡 Glucemia alta",  c:C.yellow},
                {key:"hiperglucemia", label:"🔴 Hiperglucemia",  c:C.red},
              ].map((x,i) => (
                <div key={x.key} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 0",borderBottom:i<4?`1px solid ${C.bg}`:"none"}}>
                  <div style={{fontSize:14,fontWeight:600,color:x.c}}>{x.label}</div>
                  <input type="number" value={settings[x.key]}
                    onChange={e=>setSettings(p=>({...p,[x.key]:parseFloat(e.target.value)}))}
                    style={{...inp,width:90,color:x.c,fontSize:16,fontWeight:700,fontFamily:"monospace",padding:"8px 10px"}} />
                </div>
              ))}
              <div style={{fontSize:11,color:C.muted,marginTop:10}}>La corrección de {settings.insulinaRapida||"insulina rápida"} se calcula respecto al valor "Objetivo".</div>
            </div>

            <div style={{background:C.card,borderRadius:20,padding:20,marginBottom:14}}>
              <div style={{fontSize:12,color:C.muted,fontWeight:700,textTransform:"uppercase",letterSpacing:1,marginBottom:16}}>👤 Datos personales</div>

              <div style={{marginBottom:14}}>
                <div style={{fontSize:13,color:C.muted,marginBottom:6}}>Sexo</div>
                <div style={{display:"flex",gap:8}}>
                  {["Femenino","Masculino"].map(s => (
                    <button key={s} onClick={()=>setSettings(p=>({...p,sexo:s}))}
                      style={{flex:1,background:settings.sexo===s?C.blue:C.bg,border:"none",color:settings.sexo===s?"white":C.muted,borderRadius:10,padding:"10px 0",fontSize:14,fontWeight:600,cursor:"pointer"}}>
                      {s}
                    </button>
                  ))}
                </div>
              </div>

              <div style={{display:"flex",gap:10,marginBottom:14}}>
                <div style={{flex:1}}>
                  <div style={{fontSize:13,color:C.muted,marginBottom:6}}>Peso (kg)</div>
                  <input type="number" value={settings.pesoKg}
                    onChange={e=>setSettings(p=>({...p,pesoKg:parseFloat(e.target.value)}))}
                    style={{...inp,color:C.sky,fontSize:18,fontWeight:700,fontFamily:"monospace"}} />
                </div>
                <div style={{flex:1}}>
                  <div style={{fontSize:13,color:C.muted,marginBottom:6}}>Altura (cm)</div>
                  <input type="number" value={settings.alturaCm}
                    onChange={e=>setSettings(p=>({...p,alturaCm:parseFloat(e.target.value)}))}
                    style={{...inp,color:C.sky,fontSize:18,fontWeight:700,fontFamily:"monospace"}} />
                </div>
              </div>

              <div style={{marginBottom:16}}>
                <div style={{fontSize:13,color:C.muted,marginBottom:6}}>Fecha de nacimiento</div>
                <input type="date" value={settings.fechaNacimiento}
                  onChange={e=>setSettings(p=>({...p,fechaNacimiento:e.target.value}))}
                  style={{...inp,padding:"10px 12px"}} />
              </div>

              {/* IMC y TMB calculados */}
              <div style={{background:C.bg,borderRadius:14,padding:14,display:"flex",gap:10}}>
                <div style={{flex:1,textAlign:"center"}}>
                  <div style={{fontSize:9,color:C.muted,textTransform:"uppercase",letterSpacing:1,marginBottom:4}}>IMC</div>
                  <div style={{fontSize:20,fontWeight:700,fontFamily:"monospace",color:imc.color}}>{imc.value}</div>
                  <div style={{fontSize:11,color:imc.color,marginTop:2}}>{imc.label}</div>
                </div>
                <div style={{flex:1,textAlign:"center",borderLeft:`1px solid ${C.border}`}}>
                  <div style={{fontSize:9,color:C.muted,textTransform:"uppercase",letterSpacing:1,marginBottom:4}}>TMB</div>
                  <div style={{fontSize:20,fontWeight:700,fontFamily:"monospace",color:C.orange}}>{tmb}</div>
                  <div style={{fontSize:11,color:C.muted,marginTop:2}}>kcal/día</div>
                </div>
              </div>
              <div style={{fontSize:10,color:C.muted,marginTop:8}}>IMC y Tasa Metabólica Basal (Mifflin-St Jeor) calculados a partir de tus datos.</div>
            </div>

            <div style={{background:C.card,borderRadius:20,padding:20,marginBottom:14}}>
              <div style={{fontSize:12,color:C.muted,fontWeight:700,textTransform:"uppercase",letterSpacing:1,marginBottom:16}}>💉 Insulinas</div>

              <div style={{marginBottom:14}}>
                <div style={{fontSize:13,color:C.muted,marginBottom:6}}>Insulina rápida</div>
                <input type="text" value={settings.insulinaRapida}
                  onChange={e=>setSettings(p=>({...p,insulinaRapida:e.target.value}))}
                  placeholder="Ej: Apidra®, NovoRapid®, Humalog®"
                  style={{...inp}} />
              </div>

              <div style={{marginBottom:14}}>
                <div style={{fontSize:13,color:C.muted,marginBottom:6}}>Insulina lenta (basal)</div>
                <input type="text" value={settings.insulinaLenta}
                  onChange={e=>setSettings(p=>({...p,insulinaLenta:e.target.value}))}
                  placeholder="Ej: Toujeo®, Lantus®, Tresiba®"
                  style={{...inp}} />
              </div>

              <div style={{fontSize:13,color:C.muted,marginBottom:6}}>Dosis diaria de {settings.insulinaLenta||"insulina lenta"} (U)</div>
              <input type="number" value={settings.toujeoDosis}
                onChange={e=>setSettings(p=>({...p,toujeoDosis:parseFloat(e.target.value)}))}
                style={{...inp,color:C.purple,fontSize:24,fontWeight:700,fontFamily:"monospace"}} />
              <div style={{fontSize:11,color:C.muted,marginTop:8}}>Se puede registrar junto a tus dosis de insulina rápida.</div>
            </div>

            <div style={{background:C.card,borderRadius:20,padding:20,marginBottom:14}}>
              <div style={{fontSize:12,color:C.muted,fontWeight:700,textTransform:"uppercase",letterSpacing:1,marginBottom:16}}>💊 Otros medicamentos</div>

              {(settings.otrosMedicamentos||[]).map((med,i) => (
                <div key={i} style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
                  <div style={{flex:1,background:"#f8fafc",borderRadius:10,padding:"10px 14px",fontSize:14,color:C.text,border:`1px solid ${C.border}`}}>
                    {med}
                  </div>
                  <button onClick={()=>setSettings(p=>({...p,otrosMedicamentos:p.otrosMedicamentos.filter((_,j)=>j!==i)}))}
                    style={{background:C.red+"20",border:"none",color:C.red,borderRadius:8,width:34,height:34,fontSize:18,cursor:"pointer",flexShrink:0}}>
                    ✕
                  </button>
                </div>
              ))}

              <div style={{display:"flex",gap:8,marginTop:8}}>
                <input type="text" value={newMed}
                  onChange={e=>setNewMed(e.target.value)}
                  onKeyDown={e=>{
                    if(e.key==="Enter"&&newMed.trim()){
                      setSettings(p=>({...p,otrosMedicamentos:[...(p.otrosMedicamentos||[]),newMed.trim()]}));
                      setNewMed("");
                    }
                  }}
                  placeholder="Ej: Metformina 500mg"
                  style={{...inp,flex:1}} />
                <button onClick={()=>{
                  if(!newMed.trim()) return;
                  setSettings(p=>({...p,otrosMedicamentos:[...(p.otrosMedicamentos||[]),newMed.trim()]}));
                  setNewMed("");
                }} style={{background:C.blue,border:"none",color:"white",borderRadius:12,padding:"0 18px",fontSize:20,fontWeight:700,cursor:"pointer",flexShrink:0}}>
                  +
                </button>
              </div>
            </div>

            <div style={{background:C.card,borderRadius:20,padding:20,marginBottom:14}}>
              <div style={{fontSize:12,color:C.muted,fontWeight:700,textTransform:"uppercase",letterSpacing:1,marginBottom:16}}>🎯 Metas nutricionales diarias</div>
              <div style={{fontSize:12,color:C.muted,marginBottom:14,lineHeight:1.6}}>
                Valores sugeridos según tu perfil. Ajústalos según indicación de tu nutricionista.
              </div>

              <div style={{marginBottom:12}}>
                <div style={{fontSize:13,color:C.muted,marginBottom:6}}>Peso meta (kg)</div>
                <input type="number" value={settings.pesoMeta||50}
                  onChange={e=>setSettings(p=>({...p,pesoMeta:parseFloat(e.target.value)||0}))}
                  style={{...inp,color:C.green,fontSize:22,fontWeight:700,fontFamily:"monospace"}} />
              </div>

              <div style={{marginBottom:12}}>
                <div style={{fontSize:13,color:C.muted,marginBottom:6}}>Carbohidratos diarios (g)</div>
                <input type="number" value={settings.metaCarbs||130}
                  onChange={e=>setSettings(p=>({...p,metaCarbs:parseFloat(e.target.value)||0}))}
                  style={{...inp,color:C.sky,fontSize:22,fontWeight:700,fontFamily:"monospace"}} />
              </div>

              <div style={{marginBottom:12}}>
                <div style={{fontSize:13,color:C.muted,marginBottom:6}}>Proteína diaria (g)</div>
                <input type="number" value={settings.metaProtein||85}
                  onChange={e=>setSettings(p=>({...p,metaProtein:parseFloat(e.target.value)||0}))}
                  style={{...inp,color:C.green,fontSize:22,fontWeight:700,fontFamily:"monospace"}} />
              </div>

              <div style={{marginBottom:4}}>
                <div style={{fontSize:13,color:C.muted,marginBottom:6}}>Calorías diarias (kcal)</div>
                <input type="number" value={settings.metaKcal||1350}
                  onChange={e=>setSettings(p=>({...p,metaKcal:parseFloat(e.target.value)||0}))}
                  style={{...inp,color:C.orange,fontSize:22,fontWeight:700,fontFamily:"monospace"}} />
              </div>
            </div>

            <button onClick={async ()=>{
                try { localStorage.setItem(`glucoapp-${currentUser}-settings`, JSON.stringify(settings)); } catch {}
                if (fileId) await saveSettingsToOneDrive(fileId, settings);
                setSettingsSaved(true);
                setTimeout(()=>setSettingsSaved(false), 2500);
              }}
              style={{width:"100%",background:settingsSaved?"#16a34a":C.blue,border:"none",color:"white",borderRadius:14,padding:"14px 0",fontSize:16,fontWeight:700,cursor:"pointer",marginBottom:12,transition:"background 0.3s"}}>
              {settingsSaved ? "✅ Configuración guardada" : "💾 Guardar configuración"}
            </button>

            <div style={{background:C.card,borderRadius:14,padding:14}}>
              <div style={{fontSize:12,color:C.muted,lineHeight:1.7}}>
                ⚠️ Herramienta de referencia. Confirma siempre las dosis con tu endocrinólogo.
              </div>
            </div>
          </div>
        )}
      </div>

      {/* ── MODAL ── */}
      {modal && (
        <div style={{position:"fixed",inset:0,zIndex:50,background:C.bg,display:"flex",flexDirection:"column"}}>
          <div style={{background:"#ffffff",borderBottom:"1px solid #e5e7eb",padding:"16px 20px",flexShrink:0}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div style={{fontSize:16,fontWeight:700,color:C.text}}>¿Qué comiste? 🍽️</div>
              <button onClick={()=>{setModal(false);setSearch("");setCat("Todos");}}
                style={{background:C.bg,border:`1px solid ${C.border}`,color:C.text,borderRadius:20,padding:"7px 16px",fontSize:14,fontWeight:700,cursor:"pointer"}}>
                Listo ✓
              </button>
            </div>
          </div>

          <div style={{padding:"12px 16px 8px",background:C.bg,flexShrink:0}}>
            <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 Buscar alimento..."
              style={{...inp}} />
          </div>

          <div style={{padding:"0 16px 10px",display:"flex",gap:6,overflowX:"auto",background:C.bg,flexShrink:0}}>
            {CATS.map(c => (
              <button key={c} onClick={()=>setCat(c)}
                style={{background:cat===c?C.blue:C.card,border:"none",color:cat===c?"white":C.muted,borderRadius:20,padding:"6px 14px",fontSize:12,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}>
                {c}
              </button>
            ))}
          </div>

          <div style={{flex:1,overflowY:"auto",padding:"0 16px"}}>
            {filtered.length===0 && (
              cat==="Recientes"
                ? <div style={{textAlign:"center",color:C.muted,padding:40,fontSize:14}}>
                    <div style={{fontSize:32,marginBottom:8}}>🕐</div>
                    Aquí aparecerán los alimentos que más uses
                  </div>
                : <div style={{textAlign:"center",color:C.muted,padding:40,fontSize:14}}>Sin resultados para "{search}"</div>
            )}
            {filtered.map((food,i) => {
              const sel = foods.find(x=>x.name===food.name&&x.portion===food.portion);
              return (
                <div key={i} style={{display:"flex",alignItems:"center",gap:10,padding:"12px 0",borderBottom:`1px solid ${C.border}30`}}>
                  <div style={{flex:1}}>
                    <div style={{display:"flex",alignItems:"center",gap:6}}>
                      <div style={{fontSize:15,fontWeight:sel?700:500,color:C.text}}>{food.name}</div>
                      {cat==="Recientes" && food.count>1 && (
                        <div style={{background:"#eff6ff",borderRadius:20,padding:"1px 7px",fontSize:10,fontWeight:700,color:C.blue}}>
                          ×{food.count}
                        </div>
                      )}
                    </div>
                    <div style={{fontSize:12,marginTop:2}}>
                      <span style={{color:C.sky,fontWeight:600}}>{food.carbs}g C</span>
                      {food.protein>0 && <span style={{color:C.green}}> · {food.protein}g P</span>}
                      {food.kcal>0 && <span style={{color:C.orange}}> · {food.kcal} kcal</span>}
                      <span style={{color:C.muted}}> · {food.portion}</span>
                    </div>
                  </div>
                  <div style={{display:"flex",alignItems:"center",gap:6}}>
                    {sel && (
                      <>
                        <button onClick={()=>remFood(food)} style={{background:C.red+"20",border:"none",color:C.red,borderRadius:8,width:34,height:34,fontSize:20,cursor:"pointer"}}>−</button>
                        <span style={{fontWeight:700,width:24,textAlign:"center",fontSize:16}}>{sel.qty}</span>
                      </>
                    )}
                    <button onClick={()=>addFood(food)}
                      style={{background:sel?C.green:C.card,border:"none",color:sel?"white":C.muted,borderRadius:10,width:38,height:38,fontSize:22,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"}}>
                      +
                    </button>
                  </div>
                </div>
              );
            })}
            <div style={{height:24}} />
          </div>

          {foods.length>0 && (
            <div style={{padding:"14px 16px 20px",borderTop:`1px solid ${C.border}`,background:C.card,flexShrink:0}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <div>
                  <div style={{fontSize:12,color:C.muted}}>{foods.length} items</div>
                  <div style={{fontSize:13,marginTop:2}}>
                    <span style={{color:C.sky,fontWeight:700}}>{totalCarbs.toFixed(1)}g C</span>
                    <span style={{color:C.green,fontWeight:700}}> · {totalProtein.toFixed(1)}g P</span>
                    <span style={{color:C.orange,fontWeight:700}}> · {totalKcal.toFixed(0)} kcal</span>
                  </div>
                </div>
                <button onClick={()=>{setModal(false);setSearch("");}}
                  style={{background:C.blue,border:"none",color:"white",borderRadius:12,padding:"11px 22px",fontSize:15,fontWeight:700,cursor:"pointer"}}>
                  Ver cálculo →
                </button>
              </div>
            </div>
          )}
        </div>
      )}

      {/* Nav */}
      <div style={{position:"fixed",bottom:0,left:"50%",transform:"translateX(-50%)",width:"100%",maxWidth:430,background:C.bg,borderTop:`1px solid ${C.border}`,display:"flex",padding:"8px 0 16px",zIndex:40}}>
        {[
          {id:"home",     icon:"🏠", label:"Inicio"},
          {id:"records",  icon:"📋", label:"Historial"},
          {id:"analysis", icon:"📈", label:"Análisis"},
          {id:"foodsPage",icon:"➕", label:"Alimentos"},
          {id:"settings", icon:"⚙️", label:"Config"},
        ].map(t => (
          <button key={t.id} onClick={()=>setTab(t.id)} style={{flex:1,background:"none",border:"none",cursor:"pointer",display:"flex",flexDirection:"column",alignItems:"center",gap:2}}>
            <div style={{fontSize:20}}>{t.icon}</div>
            <div style={{fontSize:9,fontWeight:600,color:tab===t.id?C.blue:C.muted}}>{t.label}</div>
            {tab===t.id && <div style={{width:20,height:2,background:C.blue,borderRadius:2}} />}
          </button>
        ))}
      </div>
    </div>
  );
}
