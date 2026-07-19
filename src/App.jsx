import React, { useState, useEffect } from "react";
const CLIENT_ID = "56cd2e49-dddb-4dc1-8e66-d9e7b9338f6c";
const SCOPES = ["User.Read", "Files.ReadWrite", "offline_access"];
const REDIRECT_URI = window.location.origin;
const FILE_NAME = "GlucoApp.xlsx";
// ── Token helpers ──
const getToken = () => localStorage.getItem("ms_token");
const setToken = (t) => localStorage.setItem("ms_token", t);
const getRefreshToken = () => localStorage.getItem("ms_refresh_token");
const setRefreshToken = (t) => { if (t) localStorage.setItem("ms_refresh_token", t); };
const clearToken = () => {
  localStorage.removeItem("ms_token");
  localStorage.removeItem("ms_refresh_token");
};
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
  // Capture the refresh_token (issued because we request the offline_access
  // scope) so we can silently renew the access_token once it expires,
  // instead of every Graph call starting to fail with 401 after ~1 hour.
  if (data.refresh_token) setRefreshToken(data.refresh_token);
  return data.access_token;
};
// Exchange a stored refresh_token for a new access_token. Access tokens
// expire (~1h) and there's no reliable client-side way to know exactly
// when, so this is called reactively whenever a Graph API call comes back
// 401 (see graphFetch below) rather than on a timer.
const refreshAccessToken = async () => {
  const refreshToken = getRefreshToken();
  if (!refreshToken) return null;
  try {
    const res = await fetch('https://login.microsoftonline.com/common/oauth2/v2.0/token', {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: CLIENT_ID, grant_type: 'refresh_token',
        refresh_token: refreshToken, scope: SCOPES.join(" "),
      }),
    });
    if (!res.ok) return null;
    const data = await res.json();
    if (!data.access_token) return null;
    setToken(data.access_token);
    if (data.refresh_token) setRefreshToken(data.refresh_token);
    return data.access_token;
  } catch { return null; }
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
  { name: "Platano hartón", portion: "1/4 unidad", carbs: 15, protein: 1, kcal: 65, cat: "Tubérculos" },
  { name: "Platano hartón", portion: "1/2 unidad", carbs: 30, protein: 1, kcal: 130, cat: "Tubérculos" },
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
  { name: "Kefir", portion: "1/2 vaso 110ml", carbs: 4, protein: 5, kcal: 34, cat: "Lácteos" },
  { name: "Kefir", portion: "1 vaso 220ml", carbs: 8, protein: 11, kcal: 34, cat: "Lácteos" },
  { name: "Milo", portion: "2 cucharadas", carbs: 13, protein: 4, kcal: 73, cat: "Bebidas" },
  { name: "Huevo", portion: "1 unidad", carbs: 0.5, protein: 6, kcal: 70, cat: "Carnes" },
  { name: "Huevo", portion: "2 unidades", carbs: 1, protein: 12, kcal: 140, cat: "Carnes" },
  { name: "Huevo", portion: "3 unidades", carbs: 1.5, protein: 18, kcal: 210, cat: "Carnes" },
  { name: "Helado Patti", portion: "1 vaso pequeño", carbs: 15, protein: 6, kcal: 143, cat: "Personalizados" },
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
  { name: "Verdura cocida", portion: "1 taza 185g", carbs: 10, protein: 0, kcal: 0, cat: "Verduras" },
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
  { name: "Salchicha de Pavo Gwalthney", portion: "1/2 salchicha 28g", carbs: 2, protein: 3, kcal: 52, cat: "Personalizados" },
  { name: "Salchicha de Pavo Gwalthney", portion: "1 salchicha 56g", carbs: 4, protein: 6, kcal: 104, cat: "Personalizados" },
  { name: "Pasta El Dorado", portion: "1/4 taza 64g", carbs: 52, protein: 4.3, kcal: 224, cat: "Personalizados" },
  { name: "Atún a las finas hierbas", portion: "1 porción 52g", carbs: 0.7, protein: 13, kcal: 99, cat: "Personalizados" },
  { name: "Maní Mix La Especial", portion: "1 sobre 35g", carbs: 13, protein: 6.1, kcal: 191, cat: "Personalizados" },
  { name: "Jamón de Cerdo Pietrán", portion: "2 tajadas 42g", carbs: 3.8, protein: 6, kcal: 49, cat: "Personalizados" },
  { name: "Queso Finesse Alpina", portion: "2 tajadas 30g", carbs: 0.6, protein: 7.8, kcal: 88, cat: "Personalizados" },
  { name: "Tortilla Rap Bimbo", portion: "1 unidad 29g", carbs: 16, protein: 2.6, kcal: 83, cat: "Personalizados" },
  { name: "Cereal Crispies The Protein Choice", portion: "1/4 sobre 50g", carbs: 23, protein: 10, kcal: 165, cat: "Personalizados" },
  { name: "Cereal Crispies The Protein Choice", portion: "100g", carbs: 46, protein: 20, kcal: 330, cat: "Personalizados" },
  { name: "Yogurt Griego San Martín", portion: "1/2 vaso 110ml", carbs: 4.6, protein: 7, kcal: 46, cat: "Personalizados" },
  { name: "Yogurt Griego San Martín", portion: "1 vaso 220ml", carbs: 9.2, protein: 14, kcal: 93, cat: "Personalizados" },
  { name: "Granola Amande Macadamia Chocolate", portion: "1 unidad 25g", carbs: 13, protein: 2.7, kcal: 120, cat: "Personalizados" },
  { name: "Granola Amande Macadamia Chocolate", portion: "2 unidades 50g", carbs: 26, protein: 5.4, kcal: 240, cat: "Personalizados" },
  // ── Cereales (Medtronic) ──
  { name: "Almojábana", portion: "1 unidad mediana", carbs: 30, protein: 0, kcal: 0, cat: "Cereales" },
  { name: "Amaranto cocido", portion: "8 cdas soperas", carbs: 30, protein: 0, kcal: 0, cat: "Cereales" },
  { name: "Arepa maíz blanco/amarillo", portion: "1 unidad CD", carbs: 30, protein: 0, kcal: 0, cat: "Cereales" },
  { name: "Arepa redonda pequeña", portion: "1 unidad", carbs: 10, protein: 0, kcal: 0, cat: "Cereales" },
  { name: "Cereal Corn Flakes", portion: "1/2 taza", carbs: 23, protein: 0, kcal: 0, cat: "Cereales" },
  { name: "Cereal Fitness Nestlé", portion: "1 bolsa 35g", carbs: 29, protein: 0, kcal: 0, cat: "Cereales" },
  { name: "Cebada perlada cocida", portion: "1 taza", carbs: 45, protein: 0, kcal: 0, cat: "Cereales" },
  { name: "Cuscús cocido", portion: "3 cdas soperas", carbs: 33, protein: 0, kcal: 0, cat: "Cereales" },
  { name: "Croissant", portion: "1 unidad mediana", carbs: 30, protein: 0, kcal: 0, cat: "Cereales" },
  { name: "Envuelto de mazorca", portion: "1 unidad mediana", carbs: 30, protein: 0, kcal: 0, cat: "Cereales" },
  { name: "Fécula de maíz", portion: "4 cdas soperas", carbs: 24, protein: 0, kcal: 0, cat: "Cereales" },
  { name: "Galleta soda/saltina/integral", portion: "3 unidades", carbs: 15, protein: 0, kcal: 0, cat: "Cereales" },
  { name: "Harina de trigo", portion: "3 cdas soperas", carbs: 15, protein: 0, kcal: 0, cat: "Cereales" },
  { name: "Maíz pira preparado", portion: "1 taza", carbs: 5, protein: 0, kcal: 0, cat: "Cereales" },
  { name: "Mantecada/Ponqué/Torta", portion: "1 tajada mediana", carbs: 35, protein: 0, kcal: 0, cat: "Cereales" },
  { name: "Mazorca cocida", portion: "1 unidad grande", carbs: 30, protein: 0, kcal: 0, cat: "Cereales" },
  { name: "Mazorca desgranada", portion: "1 cda sopera", carbs: 3, protein: 0, kcal: 0, cat: "Cereales" },
  { name: "Pan tajado blanco/centeno/integral", portion: "1 tajada", carbs: 15, protein: 0, kcal: 0, cat: "Cereales" },
  { name: "Pan blandito/rollo panadería", portion: "1 unidad pequeña", carbs: 20, protein: 0, kcal: 0, cat: "Cereales" },
  { name: "Pan perro", portion: "1 unidad mediana", carbs: 35, protein: 0, kcal: 0, cat: "Cereales" },
  { name: "Pan hamburguesa", portion: "1 unidad mediana", carbs: 45, protein: 0, kcal: 0, cat: "Cereales" },
  { name: "Pan de bono", portion: "1 unidad mediana", carbs: 30, protein: 0, kcal: 0, cat: "Cereales" },
  { name: "Pan pita", portion: "1 unidad CD", carbs: 30, protein: 0, kcal: 0, cat: "Cereales" },
  { name: "Tostada/calado", portion: "1 unidad pequeña", carbs: 10, protein: 0, kcal: 0, cat: "Cereales" },
  { name: "Tortilla mexicana", portion: "1 unidad mediana", carbs: 15, protein: 0, kcal: 0, cat: "Cereales" },
  { name: "Taco mexicano vacío", portion: "1 unidad", carbs: 8, protein: 0, kcal: 0, cat: "Cereales" },
  // ── Tubérculos nuevos ──
  { name: "Arracacha", portion: "1/2 unidad mediana", carbs: 15, protein: 0, kcal: 0, cat: "Tubérculos" },
  { name: "Guineo verde", portion: "1 unidad pequeña", carbs: 15, protein: 0, kcal: 0, cat: "Tubérculos" },
  { name: "Ouyuco/Chuguas", portion: "7 unidades grandes", carbs: 15, protein: 0, kcal: 0, cat: "Tubérculos" },
  // ── Leguminosas nuevas ──
  { name: "Arveja seca cocida", portion: "1 taza", carbs: 30, protein: 0, kcal: 0, cat: "Leguminosas" },
  { name: "Frijol rojo cocido", portion: "1 taza", carbs: 30, protein: 0, kcal: 0, cat: "Leguminosas" },
  // ── Frutas nuevas ──
  { name: "Anón", portion: "1 unidad pequeña", carbs: 15, protein: 0, kcal: 0, cat: "Frutas" },
  { name: "Bananano bocadillo", portion: "1/2 unidad mediana", carbs: 15, protein: 0, kcal: 0, cat: "Frutas" },
  { name: "Chirimoya", portion: "1/2 unidad mediana", carbs: 15, protein: 0, kcal: 0, cat: "Frutas" },
  { name: "Chontaduro", portion: "1 unidad mediana", carbs: 2, protein: 0, kcal: 0, cat: "Frutas" },
  { name: "Ciruela común", portion: "3 unidades pequeñas", carbs: 15, protein: 0, kcal: 0, cat: "Frutas" },
  { name: "Ciruela importada", portion: "3 cdas postreras", carbs: 15, protein: 0, kcal: 0, cat: "Frutas" },
  { name: "Ciruela pasa", portion: "6 unidades medianas", carbs: 15, protein: 0, kcal: 0, cat: "Frutas" },
  { name: "Coco", portion: "2 unidades pequeñas", carbs: 15, protein: 0, kcal: 0, cat: "Frutas" },
  { name: "Curuba", portion: "1 unidad mediana", carbs: 15, protein: 0, kcal: 0, cat: "Frutas" },
  { name: "Durazno nacional", portion: "2 unidades medianas", carbs: 15, protein: 0, kcal: 0, cat: "Frutas" },
  { name: "Durazno importado/Nectarín", portion: "13 unidades pequeñas", carbs: 15, protein: 0, kcal: 0, cat: "Frutas" },
  { name: "Feijoa", portion: "2 unidades medianas", carbs: 15, protein: 0, kcal: 0, cat: "Frutas" },
  { name: "Granadilla", portion: "2 unidades medianas", carbs: 15, protein: 0, kcal: 0, cat: "Frutas" },
  { name: "Guayaba", portion: "6 cdas soperas", carbs: 15, protein: 0, kcal: 0, cat: "Frutas" },
  { name: "Guanábana", portion: "3 unidades medianas", carbs: 15, protein: 0, kcal: 0, cat: "Frutas" },
  { name: "Higo", portion: "1 grande o 2 pequeñas", carbs: 15, protein: 0, kcal: 0, cat: "Frutas" },
  { name: "Mango de azúcar", portion: "3/4 taza", carbs: 15, protein: 0, kcal: 0, cat: "Frutas" },
  { name: "Mango Tommy", portion: "1 unidad pequeña", carbs: 15, protein: 0, kcal: 0, cat: "Frutas" },
  { name: "Moras", portion: "2 unidades pequeñas", carbs: 15, protein: 0, kcal: 0, cat: "Frutas" },
  { name: "Níspero sin semillas", portion: "1 unidad mediana", carbs: 15, protein: 0, kcal: 0, cat: "Frutas" },
  { name: "Pera", portion: "1/2 taza picada", carbs: 15, protein: 0, kcal: 0, cat: "Frutas" },
  { name: "Pitahaya", portion: "1 tajada delgada", carbs: 15, protein: 0, kcal: 0, cat: "Frutas" },
  { name: "Tamarindo", portion: "2 unidades grandes", carbs: 15, protein: 0, kcal: 0, cat: "Frutas" },
  { name: "Tomate de árbol", portion: "1 unidad grande", carbs: 15, protein: 0, kcal: 0, cat: "Frutas" },
  { name: "Toronja", portion: "13 unidades", carbs: 15, protein: 0, kcal: 0, cat: "Frutas" },
  { name: "Uchuva", portion: "20 unidades", carbs: 15, protein: 0, kcal: 0, cat: "Frutas" },
  { name: "Uva nacional", portion: "1 unidad mediana", carbs: 15, protein: 0, kcal: 0, cat: "Frutas" },
  { name: "Zapote", portion: "1 unidad mediana", carbs: 15, protein: 0, kcal: 0, cat: "Frutas" },
  // ── Lácteos nuevos ──
  { name: "Leche semi/descremada", portion: "1 vaso 200ml", carbs: 10, protein: 0, kcal: 0, cat: "Lácteos" },
  { name: "Leche deslactosada", portion: "1 vaso 200ml", carbs: 10, protein: 0, kcal: 0, cat: "Lácteos" },
  { name: "Leche en polvo entera", portion: "3 cdas colmadas", carbs: 10, protein: 0, kcal: 0, cat: "Lácteos" },
  { name: "Yogurt Slight Colanta", portion: "1 vaso 200ml", carbs: 12, protein: 0, kcal: 0, cat: "Lácteos" },
  { name: "Yogurt Benecol", portion: "1 unidad 100ml", carbs: 4, protein: 0, kcal: 0, cat: "Lácteos" },
  { name: "Avena Finesse", portion: "1 vaso 250ml", carbs: 15, protein: 0, kcal: 0, cat: "Lácteos" },
  { name: "Yogurt Pasco sin azúcar", portion: "1 vaso 150ml", carbs: 5, protein: 0, kcal: 0, cat: "Lácteos" },
  { name: "Yogurt San Fernando Light", portion: "1 vaso 160ml", carbs: 10, protein: 0, kcal: 0, cat: "Lácteos" },
  { name: "Yogurt Slim Natural Doña Leche", portion: "1 vaso 200ml", carbs: 12, protein: 0, kcal: 0, cat: "Lácteos" },
  { name: "Yogurt Dietético Coolechera", portion: "1 vaso 150ml", carbs: 10, protein: 0, kcal: 0, cat: "Lácteos" },
  { name: "Helado Light Mimos", portion: "1 bola 70g", carbs: 18, protein: 0, kcal: 0, cat: "Lácteos" },
  { name: "Paleta Light Robin Hood", portion: "1 unidad 60g", carbs: 20, protein: 0, kcal: 0, cat: "Lácteos" },
  // ── Alimentos procesados nuevos ──
  { name: "Arroz con pollo", portion: "1 taza", carbs: 38, protein: 0, kcal: 0, cat: "Personalizados" },
  { name: "Pizza Zenú", portion: "1 tajada mediana", carbs: 35, protein: 0, kcal: 0, cat: "Personalizados" },
  { name: "Lasaña Zenú", portion: "1 porción mediana", carbs: 47, protein: 0, kcal: 0, cat: "Personalizados" },
  { name: "Hamburguesa pan Bimbo", portion: "1 unidad mediana", carbs: 45, protein: 0, kcal: 0, cat: "Personalizados" },
  { name: "Nuggets de pollo Zenú", portion: "5 unidades", carbs: 14, protein: 0, kcal: 0, cat: "Personalizados" },
  { name: "Aborrajado sin bocadillo Kalisto", portion: "1 unidad pequeña", carbs: 20, protein: 0, kcal: 0, cat: "Personalizados" },
  { name: "Tamal santandereano sin hojas", portion: "1 unidad mediana", carbs: 75, protein: 0, kcal: 0, cat: "Personalizados" },
  { name: "Tamal tolimense sin hojas", portion: "1 unidad mediana", carbs: 60, protein: 0, kcal: 0, cat: "Personalizados" },
  { name: "Salvado de trigo", portion: "5 cdas rasas", carbs: 30, protein: 0, kcal: 0, cat: "Personalizados" },
  { name: "All Bran", portion: "6 cdas", carbs: 21, protein: 0, kcal: 0, cat: "Personalizados" },
  { name: "Tostada Bimbo", portion: "2 unidades", carbs: 4, protein: 0, kcal: 0, cat: "Personalizados" },
  { name: "Paleta en agua Mimos Light", portion: "1 unidad 25g", carbs: 12, protein: 0, kcal: 0, cat: "Personalizados" },
  { name: "Papas fritas Margarita", portion: "1 paquete 19g", carbs: 13, protein: 0, kcal: 0, cat: "Personalizados" },
  { name: "Rosquitas Cronch", portion: "1 paquete 28g", carbs: 20, protein: 0, kcal: 0, cat: "Personalizados" },
  { name: "Tajaditas plátano verde Natuchips", portion: "1 paquete", carbs: 3, protein: 0, kcal: 0, cat: "Personalizados" },
  { name: "Insta Crem", portion: "1 sobre", carbs: 3, protein: 0, kcal: 0, cat: "Personalizados" },
  { name: "Choco Lyne", portion: "1 pastilla", carbs: 2, protein: 0, kcal: 0, cat: "Personalizados" },
  { name: "Choco Lyne/Choco Express", portion: "1 paquete 50g", carbs: 10, protein: 0, kcal: 0, cat: "Personalizados" },
  { name: "Maní", portion: "1 paquete 40g", carbs: 10, protein: 0, kcal: 0, cat: "Personalizados" },
  { name: "Marañón", portion: "1 paquete 50g", carbs: 11, protein: 0, kcal: 0, cat: "Personalizados" },
  { name: "Mezcla de nueces", portion: "1 paquete 23g", carbs: 16, protein: 0, kcal: 0, cat: "Personalizados" },
  { name: "Barra granola Sport Lyne", portion: "1 unidad 25g", carbs: 13.5, protein: 0, kcal: 0, cat: "Personalizados" },
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
// Defensive time parser: normally "from"/"to" are "HH:MM" strings, but if
// Excel auto-converted the cell to its Time type, Graph API returns a
// fraction-of-a-day number instead (e.g. 0.5 = 12:00) — handle both so a
// malformed value never throws and blanks the whole app.
const parseRatioTime = (t) => {
  if (typeof t === "number") {
    const totalMin = Math.round(t*24*60);
    return [Math.floor(totalMin/60)%24, totalMin%60];
  }
  const parts = String(t||"0:0").split(":");
  return [parseInt(parts[0])||0, parseInt(parts[1])||0];
};
const getCurrentRatio = (ratios) => {
  try {
    const now = new Date();
    const cur = now.getHours()*60 + now.getMinutes();
    for (const r of ratios) {
      const [fh,fm] = parseRatioTime(r.from);
      const [th,tm] = parseRatioTime(r.to);
      const from = fh*60+fm;
      const to = th===0&&tm===0 ? 24*60 : th*60+tm;
      if (cur>=from && cur<to) return r.ratio;
    }
    return ratios[0]?.ratio ?? DEFAULT_SETTINGS.ratios[0].ratio;
  } catch {
    return DEFAULT_SETTINGS.ratios[0].ratio;
  }
};
// Shared request helper: attaches the current access token, and if Graph
// responds 401 (expired token), silently refreshes once via the stored
// refresh_token and retries before giving up. This is what actually fixes
// "conecta pero no lee nada" once the ~1h access token has expired.
const graphFetch = async (url, options = {}) => {
  const doFetch = (token) => fetch(`https://graph.microsoft.com/v1.0${url}`, {
    ...options,
    headers: { ...(options.headers||{}), Authorization: `Bearer ${token}` },
  });
  let res = await doFetch(getToken());
  if (res.status === 401) {
    const newToken = await refreshAccessToken();
    if (newToken) res = await doFetch(newToken);
  }
  if (!res.ok) throw new Error(await res.text());
  return res.json();
};
const graphGet = (url) => graphFetch(url);
const graphPost = (url, body) => graphFetch(url, {
  method:"POST",
  headers:{ "Content-Type":"application/json" },
  body: body ? JSON.stringify(body) : undefined,
});
const graphPatch = (url, body) => graphFetch(url, {
  method:"PATCH",
  headers:{ "Content-Type":"application/json" },
  body: JSON.stringify(body),
});
const graphPut = (url, body, contentType) => graphFetch(url, {
  method:"PUT",
  headers:{ "Content-Type": contentType || "application/json" },
  body,
});
const HEADER_ROW = ["Fecha","Hora","Glucemia","Carbs","Proteina","Kcal","Insulina","Alimentos","Toujeo"];
// ── Configuracion sheet: one column per setting (not a JSON blob) ──
// Unmodified fields simply keep whatever value they already had, because
// we always write the FULL current settings object — nothing is dropped.
const SETTINGS_HEADER = [
  "Sensibilidad","Hipoglucemia","Glucemia baja","Objetivo","Glucemia alta","Hiperglucemia",
  "Sexo","Peso (kg)","Altura (cm)","Fecha nacimiento","Peso meta (kg)",
  "Insulina rápida","Insulina lenta","Dosis insulina lenta (U)","Otros medicamentos",
  "Meta carbs (g)","Meta proteína (g)","Meta calorías",
  "Ratio Mañana desde","Ratio Mañana hasta","Ratio Mañana (g/U)",
  "Ratio Tarde desde","Ratio Tarde hasta","Ratio Tarde (g/U)",
  "Ratio Noche desde","Ratio Noche hasta","Ratio Noche (g/U)",
];
const settingsToRow = (s) => {
  const r = (s.ratios && s.ratios.length===3) ? s.ratios : DEFAULT_SETTINGS.ratios;
  return [
    s.sensitivity, s.hipoglucemia, s.glucemiaBaja, s.objetivo, s.glucemiaAlta, s.hiperglucemia,
    s.sexo, s.pesoKg, s.alturaCm, s.fechaNacimiento, s.pesoMeta,
    s.insulinaRapida, s.insulinaLenta, s.toujeoDosis, (s.otrosMedicamentos||[]).join("; "),
    s.metaCarbs, s.metaProtein, s.metaKcal,
    r[0].from, r[0].to, r[0].ratio,
    r[1].from, r[1].to, r[1].ratio,
    r[2].from, r[2].to, r[2].ratio,
  ];
};
// If Excel auto-converted a "HH:MM" text cell to its Time type, Graph API
// returns a fraction-of-a-day number instead — turn it back into "HH:MM".
const normalizeExcelTimeStr = (val) => {
  if (val === null || val === undefined || val === "") return "";
  if (typeof val === "number") {
    const totalMin = Math.round(val*24*60);
    const hh = Math.floor(totalMin/60)%24;
    const mm = totalMin%60;
    return `${String(hh).padStart(2,"0")}:${String(mm).padStart(2,"0")}`;
  }
  return String(val);
};
// Same idea for a "YYYY-MM-DD" text cell that got auto-converted to a Date.
const normalizeExcelDateISO = (val) => {
  if (val === null || val === undefined || val === "") return "";
  if (typeof val === "number") {
    const excelEpoch = Date.UTC(1899, 11, 30);
    const d = new Date(excelEpoch + val*86400000);
    return `${d.getUTCFullYear()}-${String(d.getUTCMonth()+1).padStart(2,"0")}-${String(d.getUTCDate()).padStart(2,"0")}`;
  }
  return String(val);
};
const rowToSettings = (row) => {
  const num = (v) => (v===undefined||v===null||v==="") ? undefined : Number(v);
  const labels = ["🌅 Mañana","☀️ Tarde","🌙 Noche"];
  const ratios = [0,1,2].map(i => ({
    label: labels[i],
    from: normalizeExcelTimeStr(row[18+i*3]) || DEFAULT_SETTINGS.ratios[i].from,
    to:   normalizeExcelTimeStr(row[19+i*3]) || DEFAULT_SETTINGS.ratios[i].to,
    ratio: num(row[20+i*3]) ?? DEFAULT_SETTINGS.ratios[i].ratio,
  }));
  return {
    sensitivity: num(row[0]) ?? DEFAULT_SETTINGS.sensitivity,
    hipoglucemia: num(row[1]) ?? DEFAULT_SETTINGS.hipoglucemia,
    glucemiaBaja: num(row[2]) ?? DEFAULT_SETTINGS.glucemiaBaja,
    objetivo: num(row[3]) ?? DEFAULT_SETTINGS.objetivo,
    glucemiaAlta: num(row[4]) ?? DEFAULT_SETTINGS.glucemiaAlta,
    hiperglucemia: num(row[5]) ?? DEFAULT_SETTINGS.hiperglucemia,
    sexo: row[6] || DEFAULT_SETTINGS.sexo,
    pesoKg: num(row[7]) ?? DEFAULT_SETTINGS.pesoKg,
    alturaCm: num(row[8]) ?? DEFAULT_SETTINGS.alturaCm,
    fechaNacimiento: normalizeExcelDateISO(row[9]) || DEFAULT_SETTINGS.fechaNacimiento,
    pesoMeta: num(row[10]) ?? DEFAULT_SETTINGS.pesoMeta,
    insulinaRapida: row[11] || DEFAULT_SETTINGS.insulinaRapida,
    insulinaLenta: row[12] || DEFAULT_SETTINGS.insulinaLenta,
    toujeoDosis: num(row[13]) ?? DEFAULT_SETTINGS.toujeoDosis,
    otrosMedicamentos: row[14] ? String(row[14]).split(";").map(x=>x.trim()).filter(Boolean) : DEFAULT_SETTINGS.otrosMedicamentos,
    metaCarbs: num(row[15]) ?? DEFAULT_SETTINGS.metaCarbs,
    metaProtein: num(row[16]) ?? DEFAULT_SETTINGS.metaProtein,
    metaKcal: num(row[17]) ?? DEFAULT_SETTINGS.metaKcal,
    ratios,
  };
};
// ── Alimentos sheet: one row per food, one column per field ──
const ALIMENTOS_HEADER = ["Nombre","Porción","Carbs (g)","Proteína (g)","Kcal","Categoría"];
// Minimal valid empty .xlsx file, base64-encoded
const EMPTY_XLSX_BASE64 = "UEsDBBQAAAAIAPux8VxGx01IlQAAAM0AAAAQAAAAZG9jUHJvcHMvYXBwLnhtbE3PTQvCMAwG4L9SdreZih6kDkQ9ip68zy51hbYpbYT67+0EP255ecgboi6JIia2mEXxLuRtMzLHDUDWI/o+y8qhiqHke64x3YGMsRoPpB8eA8OibdeAhTEMOMzit7Dp1C5GZ3XPlkJ3sjpRJsPiWDQ6sScfq9wcChDneiU+ixNLOZcrBf+LU8sVU57mym/8ZAW/B7oXUEsDBBQAAAAIAPux8Vzt8pUt7wAAACsCAAARAAAAZG9jUHJvcHMvY29yZS54bWzNksFOwzAMhl8F5d46bdlAUZcLiBNISEwCcYsSb4tomigxavf2pGXrhOABOMb+8/mz5FYHoX3E5+gDRrKYrkbX9UnosGEHoiAAkj6gU6nMiT43dz46RfkZ9xCU/lB7hJrzNTgkZRQpmIBFWIhMtkYLHVGRjye80Qs+fMZuhhkN2KHDnhJUZQVMThPDcexauAAmGGF06buAZiHO1T+xcwfYKTkmu6SGYSiHZs7lHSp4e3p8mdctbJ9I9Rrzr2QFHQNu2Hnya3N3v31gsub1uuA3RXW75Y2oVmJ1/T65/vC7CDtv7M7+Y+OzoGzh113IL1BLAwQUAAAACAD7sfFcmVycIxAGAACcJwAAEwAAAHhsL3RoZW1lL3RoZW1lMS54bWztWltz2jgUfu+v0Hhn9m0LxjaBtrQTc2l227SZhO1OH4URWI1seWSRhH+/RzYQy5YN7ZJNups8BCzp+85FR+foOHnz7i5i6IaIlPJ4YNkv29a7ty/e4FcyJBFBMBmnr/DACqVMXrVaaQDDOH3JExLD3IKLCEt4FMvWXOBbGi8j1uq0291WhGlsoRhHZGB9XixoQNBUUVpvXyC05R8z+BXLVI1lowETV0EmuYi08vlsxfza3j5lz+k6HTKBbjAbWCB/zm+n5E5aiOFUwsTAamc/VmvH0dJIgILJfZQFukn2o9MVCDINOzqdWM52fPbE7Z+Mytp0NG0a4OPxeDi2y9KLcBwE4FG7nsKd9Gy/pEEJtKNp0GTY9tqukaaqjVNP0/d93+ubaJwKjVtP02t33dOOicat0HgNvvFPh8Ouicar0HTraSYn/a5rpOkWaEJG4+t6EhW15UDTIABYcHbWzNIDll4p+nWUGtkdu91BXPBY7jmJEf7GxQTWadIZljRGcp2QBQ4AN8TRTFB8r0G2iuDCktJckNbPKbVQGgiayIH1R4Ihxdyv/fWXu8mkM3qdfTrOa5R/aasBp+27m8+T/HPo5J+nk9dNQs5wvCwJ8fsjW2GHJ247E3I6HGdCfM/29pGlJTLP7/kK6048Zx9WlrBdz8/knoxyI7vd9lh99k9HbiPXqcCzIteURiRFn8gtuuQROLVJDTITPwidhphqUBwCpAkxlqGG+LTGrBHgE323vgjI342I96tvmj1XoVhJ2oT4EEYa4pxz5nPRbPsHpUbR9lW83KOXWBUBlxjfNKo1LMXWeJXA8a2cPB0TEs2UCwZBhpckJhKpOX5NSBP+K6Xa/pzTQPCULyT6SpGPabMjp3QmzegzGsFGrxt1h2jSPHr+BfmcNQockRsdAmcbs0YhhGm78B6vJI6arcIRK0I+Yhk2GnK1FoG2camEYFoSxtF4TtK0EfxZrDWTPmDI7M2Rdc7WkQ4Rkl43Qj5izouQEb8ehjhKmu2icVgE/Z5ew0nB6ILLZv24fobVM2wsjvdH1BdK5A8mpz/pMjQHo5pZCb2EVmqfqoc0PqgeMgoF8bkePuV6eAo3lsa8UK6CewH/0do3wqv4gsA5fy59z6XvufQ9odK3NyN9Z8HTi1veRm5bxPuuMdrXNC4oY1dyzcjHVK+TKdg5n8Ds/Wg+nvHt+tkkhK+aWS0jFpBLgbNBJLj8i8rwKsQJ6GRbJQnLVNNlN4oSnkIbbulT9UqV1+WvuSi4PFvk6a+hdD4sz/k8X+e0zQszQ7dyS+q2lL61JjhK9LHMcE4eyww7ZzySHbZ3oB01+/ZdduQjpTBTl0O4GkK+A226ndw6OJ6YkbkK01KQb8P56cV4GuI52QS5fZhXbefY0dH758FRsKPvPJYdx4jyoiHuoYaYz8NDh3l7X5hnlcZQNBRtbKwkLEa3YLjX8SwU4GRgLaAHg69RAvJSVWAxW8YDK5CifEyMRehw55dcX+PRkuPbpmW1bq8pdxltIlI5wmmYE2eryt5lscFVHc9VW/Kwvmo9tBVOz/5ZrcifDBFOFgsSSGOUF6ZKovMZU77nK0nEVTi/RTO2EpcYvOPmx3FOU7gSdrYPAjK5uzmpemUxZ6by3y0MCSxbiFkS4k1d7dXnm5yueiJ2+pd3wWDy/XDJRw/lO+df9F1Drn723eP6bpM7SEycecURAXRFAiOVHAYWFzLkUO6SkAYTAc2UyUTwAoJkphyAmPoLvfIMuSkVzq0+OX9FLIOGTl7SJRIUirAMBSEXcuPv75Nqd4zX+iyBbYRUMmTVF8pDicE9M3JD2FQl867aJguF2+JUzbsaviZgS8N6bp0tJ//bXtQ9tBc9RvOjmeAes4dzm3q4wkWs/1jWHvky3zlw2zreA17mEyxDpH7BfYqKgBGrYr66r0/5JZw7tHvxgSCb/NbbpPbd4Ax81KtapWQrET9LB3wfkgZjjFv0NF+PFGKtprGtxtoxDHmAWPMMoWY434dFmhoz1YusOY0Kb0HVQOU/29QNaPYNNByRBV4xmbY2o+ROCjzc/u8NsMLEjuHti78BUEsDBBQAAAAIAPux8VyVniUOEwEAAMwBAAAYAAAAeGwvd29ya3NoZWV0cy9zaGVldDEueG1sTVFdT8MgFP0rhB8wOpOpWdom24zRB5NmRn1m621LBtwKt1b/vUDXZk+ccz8O50A+orv4DoDYr9HWF7wj6rdC+HMHRvoV9mBDp0FnJAXqWuF7B7JOS0aLuyy7F0Yqy8s81SpX5jiQVhYqx/xgjHR/e9A4FnzN58JRtR3FgijzXrbwDvTRVy4wsajUyoD1Ci1z0BR8t97u0nwa+FQw+hvMYpIT4iWS17rgWTQEGs4UFWQ4fuAAWkehYOP7qsmXK+PiLZ7Vn1P2kOUkPRxQf6mauoI/clZDIwdNRxxf4Jpnsxh8kiRnuQnHnG/Stcp6pqEJ49nqYcOZm3YnQtindzohEZoEu/Dc4OJA6DeINJNoffnA8h9QSwMEFAAAAAgA+7HxXHzzo9xRAgAA9gkAAA0AAAB4bC9zdHlsZXMueG1s3VbbitswEP0V4Q+ok5g1cUnyUENgoS0Luw99VWI5EejiyvKS9Os7Izl2s6tZKH2rTfDMHJ25G2fT+6sSz2chPLtoZfptdva++5zn/fEsNO8/2U4YQFrrNPegulPed07wpkeSVvlqsShzzaXJdhsz6L32PTvawfhttsjy3aa1ZrYss2iAo1wL9srVNqu5kgcnw1mupbpG8woNR6usYx5SEUgGS/8rwsuoYZajHy2NdWjMY4Tw6MGpVGpKYJVFw27Tce+FM3tQAicY30FslF+uHWRwcvy6XD1kMyE8IMjBuka4uzqjabdRovVAcPJ0xqe3XY6g91aD0Eh+soaHHG6MUQC3R6HUM47oR3vn+9Ky2OvHBtvMsNSbCAmNYnQTFfT/p7fo+5/dsk6+Wv9lgGpM0H8O1osnJ1p5CfqlvY8/hQ6J3EWfrAyXY5t9x51Tswt2GKTy0ozaWTaNMO9qA/eeH2Cp7/zD+Ua0fFD+ZQK32Sx/E40cdDWdesKyxlOz/BVnuCynzYRY0jTiIpp6VN3pEEQGAkQdLyS8RfbhSiMUJ2JpBDEqDpUBxYksKs7/VM+arCdiVG7rJLImOWuSE1kppA43FSfNqeBKV1pVRVGWVEfrOplBTfWtLPGX9kblhgwqDkb6u17T06Y35OM9oGb60YZQldKbSFVK9xqRdN+QUVXpaVNxkEFNgdodjJ+OgzuV5hQFTpXKjXqDaaSqKAR3Mb2jZUl0p8Q7PR/qLSmKqkojiKUzKAoKwbeRRqgMMAcKKYrwHXzzPcpv36l8/qe3+w1QSwMEFAAAAAgA+7HxXJeKuxzAAAAAEwIAAAsAAABfcmVscy8ucmVsc52SuW7DMAxAf8XQnjAH0CGIM2XxFgT5AVaiD9gSBYpFnb+v2qVxkAsZeT08EtweaUDtOKS2i6kY/RBSaVrVuAFItiWPac6RQq7ULB41h9JARNtjQ7BaLD5ALhlmt71kFqdzpFeIXNedpT3bL09Bb4CvOkxxQmlISzMO8M3SfzL38ww1ReVKI5VbGnjT5f524EnRoSJYFppFydOiHaV/Hcf2kNPpr2MitHpb6PlxaFQKjtxjJYxxYrT+NYLJD+x+AFBLAwQUAAAACAD7sfFc1BOG9zUBAAAmAgAADwAAAHhsL3dvcmtib29rLnhtbI1R0W7CMAz8lSofsBa0IQ1RXoa2IU0bgon30LrUIokrx4WNr5/bqhrSXvaU3Nm63F0WF+LTgeiUfHkXYm5qkWaeprGowdt4Rw0EnVTE3opCPqaxYbBlrAHEu3SaZbPUWwxmuRi1NpzeAhIoBCko2RF7hEv8nXcwOWPEAzqU79z0dwcm8RjQ4xXK3GQmiTVdXonxSkGs2xVMzuVmMgz2wILFH3rXmfy0h9gzYg9bq0ZyM8tUsEKO0m/0+lY9nkGXB9QKPaMT4JUVeGFqGwzHTkZTpDcx+h7Gcyhxzv+pkaoKC1hR0XoIMvTI4DqDIdbYRJME6yE3WzhiFKbYZdJH1uWQT9TYTVs8Rx3wuhwsjr5KqDBA+a5SUXntqNhw0h29zvT+YfKoXbTOPSn3Ed7IlmPM8YuWP1BLAwQUAAAACAD7sfFcJB6boq0AAAD4AQAAGgAAAHhsL19yZWxzL3dvcmtib29rLnhtbC5yZWxztZE9DoMwDIWvEuUANVCpQwVMXVgrLhAF8yMSEsWuCrcvhQGQOnRhsp4tf+/JTp9oFHduoLbzJEZrBspky+zvAKRbtIouzuMwT2oXrOJZhga80r1qEJIoukHYM2Se7pminDz+Q3R13Wl8OP2yOPAPMLxd6KlFZClKFRrkTMJotjbBUuLLTJaiqDIZiiqWcFog4skgbWlWfbBPTrTneRc390WuzeMJrt8McHh0/gFQSwMEFAAAAAgA+7HxXGWQeZIZAQAAzwMAABMAAABbQ29udGVudF9UeXBlc10ueG1srZNNTsMwEIWvEmVbJS4sWKCmG2ALXXABY08aq/6TZ1rS2zNO2kqgEhWFTax43rzPnpes3o8RsOid9diUHVF8FAJVB05iHSJ4rrQhOUn8mrYiSrWTWxD3y+WDUMETeKooe5Tr1TO0cm+peOl5G03wTZnAYlk8jcLMakoZozVKEtfFwesflOpEqLlz0GBnIi5YUIqrhFz5HXDqeztASkZDsZGJXqVjleitQDpawHra4soZQ9saBTqoveOWGmMCqbEDIGfr0XQxTSaeMIzPu9n8wWYKyMpNChE5sQR/x50jyd1VZCNIZKaveCGy9ez7QU5bg76RzeP9DGk35IFiWObP+HvGF/8bzvERwu6/P7G81k4af+aL4T9efwFQSwECFAMUAAAACAD7sfFcRsdNSJUAAADNAAAAEAAAAAAAAAAAAAAAgAEAAAAAZG9jUHJvcHMvYXBwLnhtbFBLAQIUAxQAAAAIAPux8Vzt8pUt7wAAACsCAAARAAAAAAAAAAAAAACAAcMAAABkb2NQcm9wcy9jb3JlLnhtbFBLAQIUAxQAAAAIAPux8VyZXJwjEAYAAJwnAAATAAAAAAAAAAAAAACAAeEBAAB4bC90aGVtZS90aGVtZTEueG1sUEsBAhQDFAAAAAgA+7HxXJWeJQ4TAQAAzAEAABgAAAAAAAAAAAAAAICBIggAAHhsL3dvcmtzaGVldHMvc2hlZXQxLnhtbFBLAQIUAxQAAAAIAPux8Vx886PcUQIAAPYJAAANAAAAAAAAAAAAAACAAWsJAAB4bC9zdHlsZXMueG1sUEsBAhQDFAAAAAgA+7HxXJeKuxzAAAAAEwIAAAsAAAAAAAAAAAAAAIAB5wsAAF9yZWxzLy5yZWxzUEsBAhQDFAAAAAgA+7HxXNQThvc1AQAAJgIAAA8AAAAAAAAAAAAAAIAB0AwAAHhsL3dvcmtib29rLnhtbFBLAQIUAxQAAAAIAPux8VwkHpuirQAAAPgBAAAaAAAAAAAAAAAAAACAATIOAAB4bC9fcmVscy93b3JrYm9vay54bWwucmVsc1BLAQIUAxQAAAAIAPux8VxlkHmSGQEAAM8DAAATAAAAAAAAAAAAAACAARcPAABbQ29udGVudF9UeXBlc10ueG1sUEsFBgAAAAAJAAkAPgIAAGEQAAAAAA==";
const findOrCreateFile = async () => {
  // Look up the file directly by path — this is strongly consistent.
  // The previous implementation used /search, which is backed by an
  // index that can lag behind recently created/renamed files. That lag
  // could make the app think GlucoApp.xlsx didn't exist yet and create
  // a second, separate (empty) file — leaving the user's real data in
  // one file while the app silently reads/writes an entirely different one.
  try {
    const file = await graphGet(`/me/drive/root:/${FILE_NAME}`);
    if (file && file.id) return { file, created:false };
  } catch {
    // 404 (or similar) — fall through and create it below.
  }
  // Not found — create it via a simple upload of an empty workbook,
  // then add the "Registros" worksheet with headers.
  const bin = atob(EMPTY_XLSX_BASE64);
  const bytes = new Uint8Array(bin.length);
  for (let i=0;i<bin.length;i++) bytes[i] = bin.charCodeAt(i);
  const file = await graphPut(
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
    // Check the actual value of A1 rather than usedRange — usedRange can
    // report a sheet as "non-empty" due to formatting/dimension metadata
    // even when no cell actually has text in it, which previously caused
    // the header row to be silently skipped on freshly created files.
    try {
      const a1 = await graphGet(`/me/drive/items/${fileId}/workbook/worksheets/Registros/range(address='A1')`);
      const a1Value = a1?.values?.[0]?.[0];
      if (!a1Value) {
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
  // Ensure "Configuracion" sheet exists — wrapped so a failure here
  // (e.g. an already-damaged workbook) doesn't abort the whole connection.
  try {
    let configSheet = wb.value?.find(s => s.name === "Configuracion");
    if (!configSheet) {
      await graphPost(`/me/drive/items/${fileId}/workbook/worksheets/add`, { name:"Configuracion" });
      await graphPatch(
        `/me/drive/items/${fileId}/workbook/worksheets/Configuracion/range(address='A1:AA1')`,
        { values:[SETTINGS_HEADER] }
      );
    } else {
      // Migrate the old "Fecha y hora / Clave / Valor" (JSON blob) layout
      // to the new one-column-per-setting layout, carrying over the most
      // recent saved settings so nothing is lost.
      const a1 = await graphGet(`/me/drive/items/${fileId}/workbook/worksheets/Configuracion/range(address='A1')`);
      const a1Value = a1?.values?.[0]?.[0];
      if (a1Value !== SETTINGS_HEADER[0]) {
        let migrated = null;
        try {
          const used = await graphGet(`/me/drive/items/${fileId}/workbook/worksheets/Configuracion/usedRange`);
          const rows = used.values || [];
          const settingsRows = rows.filter(r => r[1] === "settings" && r[2]);
          if (settingsRows.length > 0) migrated = JSON.parse(settingsRows[settingsRows.length-1][2]);
        } catch {}
        await graphPatch(
          `/me/drive/items/${fileId}/workbook/worksheets/Configuracion/range(address='A1:AA1')`,
          { values:[SETTINGS_HEADER] }
        );
        if (migrated) {
          await graphPatch(
            `/me/drive/items/${fileId}/workbook/worksheets/Configuracion/range(address='A2:AA2')`,
            { values:[settingsToRow({...DEFAULT_SETTINGS, ...migrated})] }
          );
        }
      }
    }
  } catch {}
  // Ensure "Alimentos" sheet exists
  try {
    let alimentosSheet = wb.value?.find(s => s.name === "Alimentos");
    if (!alimentosSheet) {
      await graphPost(`/me/drive/items/${fileId}/workbook/worksheets/add`, { name:"Alimentos" });
      await graphPatch(
        `/me/drive/items/${fileId}/workbook/worksheets/Alimentos/range(address='A1:F1')`,
        { values:[ALIMENTOS_HEADER] }
      );
    } else {
      // Migrate the old "clave/valor" (JSON blob) layout to one row per food.
      const a1 = await graphGet(`/me/drive/items/${fileId}/workbook/worksheets/Alimentos/range(address='A1')`);
      const a1Value = a1?.values?.[0]?.[0];
      if (a1Value !== ALIMENTOS_HEADER[0]) {
        let migrated = null;
        try {
          const used = await graphGet(`/me/drive/items/${fileId}/workbook/worksheets/Alimentos/usedRange`);
          const rows = used.values || [];
          for (const row of rows) {
            if (row[0] === "customFoods" && row[1]) { try { migrated = JSON.parse(row[1]); } catch {} break; }
          }
        } catch {}
        await graphPatch(
          `/me/drive/items/${fileId}/workbook/worksheets/Alimentos/range(address='A1:F1')`,
          { values:[ALIMENTOS_HEADER] }
        );
        if (migrated && migrated.length > 0) {
          await graphPatch(
            `/me/drive/items/${fileId}/workbook/worksheets/Alimentos/range(address='A2:F${1+migrated.length}')`,
            { values: migrated.map(f => [f.name, f.portion, f.carbs, f.protein, f.kcal, f.cat||"Personalizados"]) }
          );
        }
      }
    }
  } catch {}
  // Ensure "Peso" sheet exists
  try {
  let pesoSheet = wb.value?.find(s => s.name === "Peso");
  if (!pesoSheet) {
    await graphPost(`/me/drive/items/${fileId}/workbook/worksheets/add`, { name:"Peso" });
    await graphPatch(
      `/me/drive/items/${fileId}/workbook/worksheets/Peso/range(address='A1:B1')`,
      { values:[["Fecha","Peso (kg)"]] }
    );
  }
  } catch {}
  return sheet;
};
// Save settings to OneDrive "Configuracion" sheet — updates row 2 IN PLACE,
// one column per setting. We always write the full current settings object,
// so any field the user didn't just change simply keeps its existing value.
const saveSettingsToOneDrive = async (fileId, settings) => {
  try {
    await graphPatch(
      `/me/drive/items/${fileId}/workbook/worksheets/Configuracion/range(address='A2:AA2')`,
      { values:[settingsToRow(settings)] }
    );
  } catch {}
};
// Load settings from OneDrive — row 2 of the columnar layout
const loadSettingsFromOneDrive = async (fileId) => {
  try {
    const used = await graphGet(`/me/drive/items/${fileId}/workbook/worksheets/Configuracion/usedRange`);
    const rows = used.values || [];
    if (rows.length < 2) return null;
    // Back-compat: old "Clave/Valor" layout stored full JSON snapshots
    const oldRows = rows.filter(r => r[1] === "settings" && r[2]);
    if (oldRows.length > 0) {
      try { return JSON.parse(oldRows[oldRows.length-1][2]); } catch {}
    }
    return rowToSettings(rows[1]);
  } catch {}
  return null;
};
// Save custom foods to OneDrive "Alimentos" sheet — one row per food.
const saveCustomFoodsToOneDrive = async (fileId, foods) => {
  try {
    let prevRows = 0;
    try {
      const used = await graphGet(`/me/drive/items/${fileId}/workbook/worksheets/Alimentos/usedRange`);
      prevRows = Math.max((used.values?.length || 1) - 1, 0);
    } catch {}
    if (foods.length > 0) {
      await graphPatch(
        `/me/drive/items/${fileId}/workbook/worksheets/Alimentos/range(address='A2:F${1+foods.length}')`,
        { values: foods.map(f => [f.name, f.portion, f.carbs, f.protein, f.kcal, f.cat||"Personalizados"]) }
      );
    }
    // Clear any leftover rows if the list got shorter
    if (prevRows > foods.length) {
      const clearCount = prevRows - foods.length;
      const blank = Array.from({length:clearCount}, () => ["","","","","",""]);
      await graphPatch(
        `/me/drive/items/${fileId}/workbook/worksheets/Alimentos/range(address='A${2+foods.length}:F${1+prevRows}')`,
        { values: blank }
      );
    }
  } catch {}
};
// Load custom foods from OneDrive "Alimentos" sheet
const loadCustomFoodsFromOneDrive = async (fileId) => {
  try {
    const used = await graphGet(`/me/drive/items/${fileId}/workbook/worksheets/Alimentos/usedRange`);
    const rows = used.values || [];
    // Back-compat: old single-blob format
    for (const row of rows) {
      if (row[0] === "customFoods" && row[1]) {
        return JSON.parse(row[1]);
      }
    }
    // New columnar format: one row per food
    const dataRows = rows.slice(1);
    const foods = dataRows
      .filter(r => r[0])
      .map(r => ({
        name: String(r[0]),
        portion: String(r[1]??""),
        carbs: Number(r[2])||0,
        protein: Number(r[3])||0,
        kcal: Number(r[4])||0,
        cat: String(r[5]||"Personalizados"),
      }));
    return foods.length > 0 ? foods : null;
  } catch {}
  return null;
};
// Excel may store the "Fecha" cell either as plain text ("13/7/2026") or,
// if the cell got auto-converted to a real date, as a serial number
// (days since 1899-12-30, per Excel's date system). Normalize both to the
// same "D/M/YYYY" (non-zero-padded) format the app uses everywhere else,
// matching toLocaleDateString("es-CO").
const normalizeExcelDate = (val) => {
  if (val === null || val === undefined || val === "") return "";
  if (typeof val === "number") {
    const excelEpoch = Date.UTC(1899, 11, 30);
    const d = new Date(excelEpoch + val * 86400000);
    return `${d.getUTCDate()}/${d.getUTCMonth()+1}/${d.getUTCFullYear()}`;
  }
  return String(val);
};
// Same problem, same fix, for the "Hora" column: Excel may have converted a
// pasted "08:14 p.m." cell into its Time type (a fraction-of-a-day number).
// calcIOB()'s regex expects a string, so this prevents another silent crash.
const normalizeExcelTimeCell = (val) => {
  if (val === null || val === undefined || val === "") return "";
  if (typeof val === "number") {
    const totalMin = Math.round(val*24*60);
    const hh24 = Math.floor(totalMin/60)%24;
    const mm = totalMin%60;
    const isPM = hh24 >= 12;
    let hh12 = hh24 % 12; if (hh12 === 0) hh12 = 12;
    return `${String(hh12).padStart(2,"0")}:${String(mm).padStart(2,"0")} ${isPM?"p":"a"}. m.`;
  }
  return String(val);
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
      date: normalizeExcelDate(r[0]),
      time: normalizeExcelTimeCell(r[1]),
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
// Rewrite the full Registros sheet from the current in-memory records —
// used when deleting a single record (e.g. an accidental duplicate), since
// there's no reliable single-row-delete against a plain range. `recordsList`
// is most-recent-first (as kept in React state); write back oldest-first to
// match the original append order, and blank out any leftover rows if the
// list got shorter.
const rewriteAllRecordsToOneDrive = async (fileId, recordsList) => {
  try {
    const chronological = [...recordsList].reverse();
    let prevRows = 0;
    try {
      const used = await graphGet(`/me/drive/items/${fileId}/workbook/worksheets/Registros/usedRange`);
      prevRows = Math.max((used.values?.length || 1) - 1, 0);
    } catch {}
    if (chronological.length > 0) {
      await graphPatch(
        `/me/drive/items/${fileId}/workbook/worksheets/Registros/range(address='A2:I${1+chronological.length}')`,
        { values: chronological.map(r => [r.date, r.time, r.glucose, r.carbs, r.protein, r.kcal, r.insulin, r.foods, r.toujeo]) }
      );
    }
    if (prevRows > chronological.length) {
      const clearCount = prevRows - chronological.length;
      const blank = Array.from({length:clearCount}, () => ["","","","","","","","",""]);
      await graphPatch(
        `/me/drive/items/${fileId}/workbook/worksheets/Registros/range(address='A${2+chronological.length}:I${1+prevRows}')`,
        { values: blank }
      );
    }
  } catch {}
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
// ── Registro de peso ──
const WEIGHT_LOG_KEY = "glucoapp-weight-log";
const loadWeightLogCache = () => {
  try {
    const raw = localStorage.getItem(WEIGHT_LOG_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch { return []; }
};
const cacheWeightLog = (log) => {
  try { localStorage.setItem(WEIGHT_LOG_KEY, JSON.stringify(log)); } catch {}
};
// Save one weight entry to OneDrive "Peso" sheet (appends a row)
const saveWeightToOneDrive = async (fileId, date, kg) => {
  try {
    let nextRow = 2;
    try {
      const used = await graphGet(`/me/drive/items/${fileId}/workbook/worksheets/Peso/usedRange`);
      nextRow = (used.values?.length || 1) + 1;
    } catch {}
    await graphPatch(
      `/me/drive/items/${fileId}/workbook/worksheets/Peso/range(address='A${nextRow}:B${nextRow}')`,
      { values:[[date, kg]] }
    );
  } catch {}
};
// Load full weight history from OneDrive "Peso" sheet
const loadWeightFromOneDrive = async (fileId) => {
  try {
    const used = await graphGet(`/me/drive/items/${fileId}/workbook/worksheets/Peso/usedRange`);
    const rows = (used.values || []).slice(1); // skip header
    return rows
      .filter(r => r[0] && r[1]!==undefined && r[1]!=="")
      .map(r => ({ date:String(r[0]), kg:parseFloat(r[1])||0 }));
  } catch {}
  return null;
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
  const [exercise, setExercise] = useState(0); // 0=ninguno, 10=leve, 20=moderado, 30=intenso
  const [newMed, setNewMed] = useState("");
  const [syncing, setSyncing] = useState(false);
  const [syncMsg, setSyncMsg] = useState("");
  const [fileId, setFileId] = useState(null);
  const [expandedDate, setExpandedDate] = useState(null);
  const [weightLog, setWeightLog] = useState([]);
  const [weightInput, setWeightInput] = useState("");
  const [weightSaved, setWeightSaved] = useState(false);
  const [odStatus, setOdStatus] = useState(msToken ? "connecting" : "disconnected");
  // odStatus: 'disconnected' | 'connecting' | 'ready' | 'error'
  const [odError, setOdError] = useState("");
  // Load custom foods from local cache on mount (best effort, not source of truth)
  // Load cached records immediately on mount so historial isn't empty while OneDrive connects
  useEffect(() => {
    const cached = loadCachedRecords();
    if (cached.length) setRecords(cached);
    const cachedWeight = loadWeightLogCache();
    if (cachedWeight.length) setWeightLog(cachedWeight);
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
  }, [customFoods, customFoodsReady, fileId, odStatus]);
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
        // Load weight history from OneDrive (overrides localStorage)
        const cloudWeight = await loadWeightFromOneDrive(file.id);
        if (cloudWeight) {
          setWeightLog(cloudWeight);
          cacheWeightLog(cloudWeight);
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
    ).sort((a,b) => a.name.localeCompare(b.name, 'es'));
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
  // ── IOB: Insulin On Board (Insulina Activa) ──
  // Apidra DIA = 4 horas. Curva de acción trapezoidal estándar.
  // % activo según minutos transcurridos desde la dosis
  const insulinActivePercent = (minutesAgo) => {
    const DIA = 240; // 4 horas en minutos
    if (minutesAgo >= DIA) return 0;
    if (minutesAgo <= 0) return 1;
    // Curva trapezoidal: sube rápido, meseta, baja gradual
    const peak = 75; // pico a los 75 min
    if (minutesAgo <= peak) {
      return 1 - (minutesAgo / peak) * 0.3; // baja 30% hasta el pico
    }
    // Después del pico baja linealmente hasta 0
    return 0.7 * (1 - (minutesAgo - peak) / (DIA - peak));
  };
  const calcIOB = () => {
    const now = new Date();
    let iob = 0;
    for (const r of records) {
      if (!r.insulin || r.insulin <= 0) continue;
      try {
        const dateParts = r.date ? r.date.split("/") : [];
        if (dateParts.length < 3 || !r.time) continue;
        const [day, month, year] = dateParts;
        // r.time is stored in 12-hour format (e.g. "08:14 p.m.") — parse
        // it AM/PM-aware, since a naive split(":") previously ignored
        // a.m./p.m. entirely and treated every dose as if taken in the
        // morning, throwing IOB off by up to 12 hours.
        const timeMatch = r.time.match(/(\d{1,2}):(\d{2})\s*([ap])\.?\s*m\.?/i);
        let hours, minutes;
        if (timeMatch) {
          hours = parseInt(timeMatch[1]);
          minutes = parseInt(timeMatch[2]);
          const isPM = timeMatch[3].toLowerCase() === "p";
          if (isPM && hours !== 12) hours += 12;
          if (!isPM && hours === 12) hours = 0;
        } else {
          const timeParts = r.time.split(":");
          if (timeParts.length < 2) continue;
          hours = parseInt(timeParts[0]);
          minutes = parseInt(timeParts[1]);
        }
        const doseTime = new Date(
          parseInt(year), parseInt(month)-1, parseInt(day),
          hours, minutes
        );
        if (isNaN(doseTime.getTime())) continue;
        const minutesAgo = (now - doseTime) / 60000;
        if (minutesAgo < 0 || minutesAgo >= 240) continue;
        const active = insulinActivePercent(minutesAgo);
        iob += r.insulin * active;
      } catch {}
    }
    return Math.max(0, Math.round(iob * 10) / 10);
  };
  const iob = calcIOB();
  const calc = () => {
    const g = parseFloat(glucose);
    if (!g && totalCarbs===0) return null;
    const corr = g ? Math.max(0,(g-settings.objetivo)/settings.sensitivity) : 0;
    const meal = totalCarbs/currentRatio;
    const rawTotal = corr + meal;
    const safeIob = isNaN(iob) ? 0 : iob;
    const afterIOB = Math.max(0, rawTotal - safeIob);
    const exerciseReduction = exercise > 0 ? afterIOB * (exercise/100) : 0;
    const netTotal = Math.max(0, afterIOB - exerciseReduction);
    if (isNaN(netTotal)) return null;
    return {
      corr:corr.toFixed(1),
      meal:meal.toFixed(1),
      raw:rawTotal.toFixed(1),
      total:Math.round(netTotal),
      netExact:netTotal.toFixed(1),
      iob:safeIob,
      iobDeducted: safeIob > 0 && rawTotal > 0,
      exercise,
      exerciseReduction: exerciseReduction.toFixed(1),
      ratio:currentRatio
    };
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
        setTimeout(() => { setGlucose(""); setFoods([]); setWithToujeo(false); setExercise(0); setSaved(false); }, 1800);
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
  const doSaveWeight = async () => {
    const kg = parseFloat(weightInput);
    if (!kg || kg<=0) return;
    const dateStr = new Date().toLocaleDateString("es-CO");
    // Replace today's entry if it already exists, otherwise append
    const updated = [...weightLog.filter(w => w.date!==dateStr), { date:dateStr, kg }];
    setWeightLog(updated);
    cacheWeightLog(updated);
    // Keep pesoKg (used for IMC/TMB) in sync with the latest weigh-in
    const newSettings = {...settings, pesoKg:kg};
    setSettings(newSettings);
    try { localStorage.setItem(`glucoapp-${currentUser}-settings`, JSON.stringify(newSettings)); } catch {}
    if (msToken && fileId && odStatus==="ready") {
      try {
        await saveWeightToOneDrive(fileId, dateStr, kg);
        await saveSettingsToOneDrive(fileId, newSettings);
      } catch {}
    }
    setWeightInput("");
    setWeightSaved(true);
    setTimeout(() => setWeightSaved(false), 2000);
  };
  // Delete a single record (e.g. an accidental duplicate) from Historial.
  // Updates local state immediately, then rewrites the Registros sheet in
  // OneDrive so the deletion sticks (OneDrive is the source of truth).
  const deleteRecord = (id) => {
    const updated = records.filter(r => r.id !== id);
    setRecords(updated);
    cacheRecords(updated);
    if (msToken && fileId && odStatus==="ready") {
      rewriteAllRecordsToOneDrive(fileId, updated).catch(()=>{});
    }
  };
  const todayStr  = new Date().toLocaleDateString("es-CO");
  const todayRecs = records.filter(r => r.date===todayStr);
  const todayTot  = dayTotals(todayRecs);
  const grouped   = groupByDate(records);
  const updateRatio = (i,key,val) => {
    setSettings(p => ({...p, ratios: p.ratios.map((r,idx) => idx===i ? {...r,[key]:val} : r)}));
  };
  // Improve OCR accuracy on real phone photos of nutrition labels: upscale,
  // grayscale, and boost contrast before handing the image to Tesseract.
  // Low-contrast, small, glare-heavy label photos are the #1 cause of the
  // "no se pudieron leer los valores" failure — this directly targets that.
  const preprocessImage = (file) => new Promise((resolve) => {
    try {
      const img = new Image();
      const url = URL.createObjectURL(file);
      img.onload = () => {
        try {
          const scale = Math.max(1, 1400 / img.width);
          const w = Math.round(img.width * scale);
          const h = Math.round(img.height * scale);
          const canvas = document.createElement("canvas");
          canvas.width = w; canvas.height = h;
          const ctx = canvas.getContext("2d");
          ctx.drawImage(img, 0, 0, w, h);
          const imgData = ctx.getImageData(0, 0, w, h);
          const d = imgData.data;
          for (let i = 0; i < d.length; i += 4) {
            const gray = 0.299*d[i] + 0.587*d[i+1] + 0.114*d[i+2];
            const contrasted = Math.min(255, Math.max(0, (gray - 128) * 1.5 + 128));
            d[i] = d[i+1] = d[i+2] = contrasted;
          }
          ctx.putImageData(imgData, 0, 0);
          canvas.toBlob((blob) => {
            URL.revokeObjectURL(url);
            resolve(blob || file);
          }, "image/jpeg", 0.92);
        } catch { URL.revokeObjectURL(url); resolve(file); }
      };
      img.onerror = () => { URL.revokeObjectURL(url); resolve(file); };
      img.src = url;
    } catch { resolve(file); }
  });
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
      const ocrInput = await preprocessImage(file);
      const { data: { text } } = await Tesseract.recognize(ocrInput, "spa+eng", {
        logger: () => {}
      });
      // Parse nutritional values from OCR text
      const normalize = (s) => s.toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g,"");
      const lines = text.split("\n").map(l => l.trim()).filter(Boolean);
      // Join lines with spaces (not the raw text) so a keyword phrase that
      // got OCR'd/wrapped across two lines — e.g. "Carbohidratos" / "totales"
      // on separate lines — is still found as one contiguous phrase.
      const fullText = normalize(lines.join(" "));
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
      const carbs   = findVal(["carbohidratos totales","total carbohydrates","total carbohydrate","carbohidratos","carbohidrato","carbohydrate","hidratos de carbono","hidratos"]);
      const protein = findVal(["proteinas","proteínas","proteina","proteína","protein"]);
      const kcal    = findVal(["valor energetico","valor energético","calorias","calorías","calories","kcal","energia","energía","energy"]);
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
        .normalize("NFD").replace(/[̀-ͯ]/g,"").trim();
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
            <div style={{background:iob>0?"#f5f3ff":"#f8fafc",border:`0.5px solid ${iob>0?"#ddd6fe":C.border}`,borderRadius:20,padding:"4px 10px",fontSize:11,fontWeight:600,color:iob>0?C.purple:C.muted}}>
              Ins Act: {iob}U
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
            {/* Recordatorio de peso — domingos */}
            {new Date().getDay()===0 && !weightLog.some(w=>w.date===todayStr) && (
              <div style={{background:"#fffbeb",border:"1.5px solid #fde68a",borderRadius:16,padding:16,marginBottom:12}}>
                <div style={{fontSize:14,fontWeight:700,color:"#92400e",marginBottom:4}}>⚖️🔔 ¡Hoy es domingo!</div>
                <div style={{fontSize:12,color:"#92400e",marginBottom:10}}>Es día de pesarte. Registra tu peso de hoy:</div>
                <div style={{display:"flex",gap:8}}>
                  <input type="number" value={weightInput} onChange={e=>setWeightInput(e.target.value)}
                    placeholder={`Peso actual (kg)`}
                    style={{...inp,flex:1,background:"#ffffff"}} />
                  <button onClick={doSaveWeight}
                    style={{background:weightSaved?"#16a34a":"#d97706",border:"none",color:"white",borderRadius:12,padding:"0 18px",fontSize:14,fontWeight:700,cursor:"pointer",flexShrink:0}}>
                    {weightSaved ? "✓" : "Guardar"}
                  </button>
                </div>
              </div>
            )}
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
            {/* Ejercicio */}
            <div style={{background:C.card,borderRadius:16,padding:"12px 16px",marginBottom:12}}>
              <div style={{fontSize:13,fontWeight:700,color:C.text,marginBottom:10}}>🏃 ¿Vas a hacer ejercicio?</div>
              <div style={{display:"flex",gap:8}}>
                {[
                  {label:"No",         value:0,  color:"#6b7280"},
                  {label:"Leve −10%",  value:10, color:"#0ea5e9"},
                  {label:"Moderado −20%", value:20, color:"#f97316"},
                  {label:"Intenso −30%",  value:30, color:"#ef4444"},
                ].map(opt => (
                  <button key={opt.value} onClick={()=>setExercise(opt.value)}
                    style={{
                      flex:1, border:`1.5px solid ${exercise===opt.value ? opt.color : C.border}`,
                      background: exercise===opt.value ? opt.color+"18" : C.bg,
                      color: exercise===opt.value ? opt.color : C.muted,
                      borderRadius:10, padding:"7px 4px", fontSize:10, fontWeight:700, cursor:"pointer",
                      textAlign:"center"
                    }}>
                    {opt.label}
                  </button>
                ))}
              </div>
            </div>
            {/* Resultado */}
            {result && !saved && (
              <div style={{background:"#eff6ff",border:"1.5px solid #bfdbfe",borderRadius:20,padding:20,marginBottom:12}}>
                <div style={{fontSize:11,color:C.blue,fontWeight:700,letterSpacing:2,textTransform:"uppercase",marginBottom:8}}>💉 DOSIS {(settings.insulinaRapida||"APIDRA").toUpperCase()}</div>
                <div style={{display:"flex",alignItems:"baseline",gap:6,marginBottom:4}}>
                  <span style={{fontSize:68,fontWeight:700,color:C.blue,fontFamily:"monospace",lineHeight:1}}>{result.total}</span>
                  <span style={{fontSize:24,color:C.blue,paddingBottom:8}}>U</span>
                </div>
                <div style={{fontSize:12,color:C.muted,marginBottom:result.iobDeducted?8:16}}>
                  Exacto: {result.raw}U → {result.total}U · Ratio: {result.ratio}g/U
                </div>
                {result.iobDeducted && (
                  <div style={{background:"#f5f3ff",border:"1px solid #e9d5ff",borderRadius:10,padding:"8px 12px",marginBottom:16,fontSize:12}}>
                    <span style={{color:C.purple,fontWeight:700}}>−{result.iob}U IOB descontado</span>
                    <span style={{color:C.muted}}> · Sin IOB serían {result.raw}U</span>
                  </div>
                )}
                <div style={{display:"flex",gap:6,marginBottom:16,flexWrap:"wrap"}}>
                  {[
                    {label:"Corrección", val:result.corr+"U", sub:"glucemia",    bg:"#dbeafe", color:C.blue},
                    {label:"Comida",     val:result.meal+"U", sub:`${totalCarbs.toFixed(1)}g÷${result.ratio}`, bg:"#dbeafe", color:C.blue},
                    ...(result.iobDeducted ? [{label:"IOB",val:"−"+result.iob+"U",sub:"activo",bg:"#f5f3ff",color:C.purple}] : []),
                    ...(result.exercise>0  ? [{label:`Ejr −${result.exercise}%`,val:"−"+result.exerciseReduction+"U",sub:"actividad",bg:"#fff7ed",color:"#f97316"}] : []),
                  ].map((x,i)=>(
                    <div key={i} style={{flex:"1 1 80px",background:x.bg,borderRadius:12,padding:"8px 10px"}}>
                      <div style={{fontSize:9,color:x.color,fontWeight:700,textTransform:"uppercase",letterSpacing:0.8,marginBottom:3}}>{x.label}</div>
                      <div style={{fontSize:15,fontWeight:700,color:x.color,fontFamily:"monospace"}}>{x.val}</div>
                      <div style={{fontSize:9,color:"#9ca3af",marginTop:1}}>{x.sub}</div>
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
                    {label:"Carbohidratos", value:todayTot.carbs,   meta:metaCarbs,   color:"#38bdf8", unit:"g", good: todayTot.carbs<=metaCarbs},
                    {label:"Proteína",      value:todayTot.protein, meta:metaProtein, color:"#f97316", unit:"g", good: todayTot.protein>=metaProtein},
                    {label:"Calorías",      value:todayTot.kcal,    meta:metaKcal,    color:"#16a34a", unit:"", good: todayTot.kcal<=metaKcal},
                  ].map(({label, value, meta, color, unit, good}) => {
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
                        <div style={{fontSize:16,marginTop:4}}>
                          {good ? <span style={{color:"#16a34a",fontWeight:700}}>✓</span> : <span style={{color:"#ef4444",fontWeight:700}}>✗</span>}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
              {/* Carbs y proteína por momento del día (Mañana/Tarde/Noche), solo de
                  HOY — usa los mismos horarios definidos en Ratio I:C, y se rellena
                  a lo largo del día a medida que registras comidas. */}
              {(() => {
                const periods = settings.ratios.map(r => {
                  const [fh,fm] = parseRatioTime(r.from);
                  const [th,tm] = parseRatioTime(r.to);
                  const from = fh*60+fm;
                  const to = (th===0 && tm===0) ? 24*60 : th*60+tm;
                  return { label:r.label, from, to, carbs:0, protein:0 };
                });
                const parseTimeToMinutes = (t) => {
                  if (!t) return null;
                  const m = String(t).match(/(\d{1,2}):(\d{2})\s*([ap])\.?\s*m\.?/i);
                  if (m) {
                    let h = parseInt(m[1]); const mi = parseInt(m[2]);
                    const isPM = m[3].toLowerCase()==="p";
                    if (isPM && h!==12) h+=12;
                    if (!isPM && h===12) h=0;
                    return h*60+mi;
                  }
                  const parts = String(t).split(":");
                  if (parts.length<2) return null;
                  const h = parseInt(parts[0]), mi = parseInt(parts[1]);
                  if (isNaN(h)||isNaN(mi)) return null;
                  return h*60+mi;
                };
                for (const r of todayRecs) {
                  const mins = parseTimeToMinutes(r.time);
                  if (mins === null) continue;
                  const period = periods.find(p => mins>=p.from && mins<p.to) || periods[periods.length-1];
                  period.carbs += parseFloat(r.carbs||0);
                  period.protein += parseFloat(r.protein||0);
                }
                const hasData = periods.some(p => p.carbs>0 || p.protein>0);
                if (!hasData) return null;
                const maxVal = Math.max(...periods.map(p=>Math.max(p.carbs,p.protein)), 1);
                const chartH = 100;
                return (
                  <div style={{background:C.card,borderRadius:16,padding:16,marginBottom:12}}>
                    <div style={{fontSize:13,fontWeight:700,color:C.muted,marginBottom:4}}>🍽️ CARBS Y PROTEÍNA POR MOMENTO DEL DÍA</div>
                    <div style={{fontSize:11,color:C.muted,marginBottom:14}}>Hoy — para ver qué comida ajustar</div>
                    <div style={{display:"flex",gap:14,marginBottom:10,fontSize:10}}>
                      <div style={{display:"flex",alignItems:"center",gap:4}}>
                        <div style={{width:8,height:8,borderRadius:2,background:C.sky}}/>
                        <span style={{color:C.muted}}>Carbohidratos (g)</span>
                      </div>
                      <div style={{display:"flex",alignItems:"center",gap:4}}>
                        <div style={{width:8,height:8,borderRadius:2,background:C.green}}/>
                        <span style={{color:C.muted}}>Proteína (g)</span>
                      </div>
                    </div>
                    <div style={{display:"flex",alignItems:"flex-end",gap:16,height:chartH}}>
                      {periods.map((p,i) => (
                        <div key={i} style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",height:"100%",justifyContent:"flex-end"}}>
                          <div style={{display:"flex",alignItems:"flex-end",gap:4,width:"100%",height:"100%",justifyContent:"center"}}>
                            <div style={{display:"flex",flexDirection:"column",alignItems:"center",width:"38%",height:"100%",justifyContent:"flex-end"}}>
                              {p.carbs>0 && <div style={{fontSize:9,color:C.sky,fontWeight:700,marginBottom:2}}>{Math.round(p.carbs)}</div>}
                              <div style={{width:"100%",height:p.carbs>0?`${Math.max((p.carbs/maxVal)*100,4)}%`:"2px",background:C.sky,borderRadius:"3px 3px 0 0",transition:"height 0.4s"}}/>
                            </div>
                            <div style={{display:"flex",flexDirection:"column",alignItems:"center",width:"38%",height:"100%",justifyContent:"flex-end"}}>
                              {p.protein>0 && <div style={{fontSize:9,color:C.green,fontWeight:700,marginBottom:2}}>{Math.round(p.protein)}</div>}
                              <div style={{width:"100%",height:p.protein>0?`${Math.max((p.protein/maxVal)*100,4)}%`:"2px",background:C.green,borderRadius:"3px 3px 0 0",transition:"height 0.4s"}}/>
                            </div>
                          </div>
                        </div>
                      ))}
                    </div>
                    <div style={{display:"flex",gap:16,marginTop:6}}>
                      {periods.map((p,i) => (
                        <div key={i} style={{flex:1,textAlign:"center",fontSize:10,color:C.text,fontWeight:600}}>{p.label}</div>
                      ))}
                    </div>
                  </div>
                );
              })()}
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
              {weeklyData.length === 0 && weightLog.length === 0 ? (
                <div style={{textAlign:"center",padding:40,color:C.muted}}>
                  <div style={{fontSize:40,marginBottom:12}}>📊</div>
                  <div style={{fontSize:15,fontWeight:600}}>Sin datos aún</div>
                  <div style={{fontSize:13,marginTop:6}}>Registra tus comidas, insulina o peso para ver el análisis</div>
                </div>
              ) : (
                <>
                  {/* Macros semanales — barras de progreso */}
                  {weeklyData.length > 0 && (
                  <div style={{background:C.card,borderRadius:16,padding:16,marginBottom:12}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
                      <div style={{fontSize:13,fontWeight:700,color:C.muted}}>📊 PROMEDIO SEMANAL</div>
                      <div style={{fontSize:11,color:C.muted}}>{weeklyData.length>0 ? weeklyData[weeklyData.length-1].label : ""}</div>
                    </div>
                    {weeklyData.length > 0 && (() => {
                      const last = weeklyData[weeklyData.length-1];
                      return (
                        <>
                          <div style={{fontSize:12,fontWeight:700,color:C.muted,marginBottom:8}}>💉 Insulina — últimos 7 días</div>
                          {(() => {
                            // Get last 7 days
                            const today = new Date();
                            const days7 = Array.from({length:7}, (_,i) => {
                              const d = new Date(today);
                              d.setDate(d.getDate() - (6-i));
                              return {
                                date: `${d.getDate()}/${d.getMonth()+1}/${d.getFullYear()}`,
                                label: ["D","L","M","X","J","V","S"][d.getDay()],
                                isToday: i===6,
                              };
                            });
                            const dayMapRapida = {};
                            const dayMapLenta = {};
                            for (const r of records) {
                              dayMapRapida[r.date] = (dayMapRapida[r.date]||0) + (r.insulin||0);
                              dayMapLenta[r.date]  = (dayMapLenta[r.date]||0)  + (r.toujeo||0);
                            }
                            const vals = days7.map(d => ({
                              ...d,
                              rapida: Math.round((dayMapRapida[d.date]||0)*10)/10,
                              lenta:  Math.round((dayMapLenta[d.date]||0)*10)/10,
                            }));
                            const maxVal = Math.max(...vals.map(v=>Math.max(v.rapida,v.lenta)), 1);
                            const avgRapida = Math.round((vals.reduce((s,v)=>s+v.rapida,0)/7)*10)/10;
                            const avgLenta  = Math.round((vals.reduce((s,v)=>s+v.lenta,0)/7)*10)/10;
                            const chartH = 90;
                            return (
                              <div>
                                <div style={{display:"flex",gap:14,marginBottom:8,fontSize:10,flexWrap:"wrap"}}>
                                  <div style={{display:"flex",alignItems:"center",gap:4}}>
                                    <div style={{width:8,height:8,borderRadius:2,background:C.purple}}/>
                                    <span style={{color:C.muted}}>{settings.insulinaRapida||"Rápida"} · prom {avgRapida}U</span>
                                  </div>
                                  <div style={{display:"flex",alignItems:"center",gap:4}}>
                                    <div style={{width:8,height:8,borderRadius:2,background:"#c4b5fd"}}/>
                                    <span style={{color:C.muted}}>{settings.insulinaLenta||"Lenta"} · prom {avgLenta}U</span>
                                  </div>
                                </div>
                                <div style={{position:"relative",height:chartH}}>
                                  <div style={{position:"absolute",left:0,right:0,bottom:`${Math.min((avgRapida/maxVal)*100,100)}%`,borderTop:`1.5px dashed ${C.purple}`,zIndex:2}} />
                                  <div style={{position:"absolute",left:0,right:0,bottom:`${Math.min((avgLenta/maxVal)*100,100)}%`,borderTop:`1.5px dashed #c4b5fd`,zIndex:2}} />
                                  <div style={{display:"flex",alignItems:"flex-end",gap:6,height:"100%"}}>
                                    {vals.map((d,i) => (
                                      <div key={i} style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",height:"100%",justifyContent:"flex-end"}}>
                                        <div style={{display:"flex",alignItems:"flex-end",gap:2,width:"100%",height:"100%",justifyContent:"center"}}>
                                          <div style={{display:"flex",flexDirection:"column",alignItems:"center",width:"38%",height:"100%",justifyContent:"flex-end"}}>
                                            {d.rapida>0 && <div style={{fontSize:8,color:C.purple,fontWeight:700,marginBottom:2}}>{d.rapida}</div>}
                                            <div title={`${d.rapida}U`} style={{
                                              width:"100%",
                                              height: d.rapida>0 ? `${Math.max((d.rapida/maxVal)*100,4)}%` : "2px",
                                              background: d.isToday ? C.purple : "#a78bfa",
                                              borderRadius:"3px 3px 0 0",
                                              transition:"height 0.4s"
                                            }}/>
                                          </div>
                                          <div style={{display:"flex",flexDirection:"column",alignItems:"center",width:"38%",height:"100%",justifyContent:"flex-end"}}>
                                            {d.lenta>0 && <div style={{fontSize:8,color:"#8b5cf6",fontWeight:700,marginBottom:2}}>{d.lenta}</div>}
                                            <div title={`${d.lenta}U`} style={{
                                              width:"100%",
                                              height: d.lenta>0 ? `${Math.max((d.lenta/maxVal)*100,4)}%` : "2px",
                                              background: d.isToday ? "#8b5cf6" : "#ddd6fe",
                                              borderRadius:"3px 3px 0 0",
                                              transition:"height 0.4s"
                                            }}/>
                                          </div>
                                        </div>
                                      </div>
                                    ))}
                                  </div>
                                </div>
                                <div style={{display:"flex",gap:6,marginTop:4}}>
                                  {vals.map((d,i) => (
                                    <div key={i} style={{flex:1,textAlign:"center",fontSize:9,color:d.isToday?C.purple:C.muted,fontWeight:d.isToday?700:400}}>{d.label}</div>
                                  ))}
                                </div>
                              </div>
                            );
                          })()}
                        </>
                      );
                    })()}
                  </div>
                  )}
                  {/* Macros por día — 3 gráficas separadas (Carbs / Proteína / Calorías) */}
                  {weeklyData.length > 0 && (() => {
                    const today = new Date();
                    const days7 = Array.from({length:7}, (_,i) => {
                      const d = new Date(today);
                      d.setDate(d.getDate() - (6-i));
                      return {
                        date: `${d.getDate()}/${d.getMonth()+1}/${d.getFullYear()}`,
                        label: ["D","L","M","X","J","V","S"][d.getDay()],
                        isToday: i===6,
                      };
                    });
                    const dayMacros = {};
                    for (const r of records) {
                      if (!dayMacros[r.date]) dayMacros[r.date] = {carbs:0,protein:0,kcal:0};
                      dayMacros[r.date].carbs   += parseFloat(r.carbs||0);
                      dayMacros[r.date].protein += parseFloat(r.protein||0);
                      dayMacros[r.date].kcal    += parseFloat(r.kcal||0);
                    }
                    const vals = days7.map(d => ({
                      ...d,
                      carbs:   Math.round(dayMacros[d.date]?.carbs||0),
                      protein: Math.round(dayMacros[d.date]?.protein||0),
                      kcal:    Math.round(dayMacros[d.date]?.kcal||0),
                    }));
                    const MacroChart = ({title, dataKey, color, meta, unit, isGood}) => {
                      const maxVal = Math.max(...vals.map(v=>v[dataKey]), meta, 1);
                      const chartH = 90;
                      return (
                        <div style={{background:C.card,borderRadius:16,padding:16,marginBottom:12}}>
                          <div style={{fontSize:13,fontWeight:700,color,marginBottom:14}}>{title} — últimos 7 días</div>
                          <div style={{position:"relative",height:chartH}}>
                            <div style={{position:"absolute",left:0,right:0,bottom:`${Math.min((meta/maxVal)*100,100)}%`,borderTop:`1.5px dashed ${C.muted}`,zIndex:2}} />
                            <div style={{display:"flex",alignItems:"flex-end",gap:6,height:"100%"}}>
                              {vals.map((d,i) => (
                                <div key={i} style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",height:"100%",justifyContent:"flex-end"}}>
                                  {d[dataKey]>0 && <div style={{fontSize:8,color,fontWeight:700,marginBottom:2}}>{d[dataKey]}</div>}
                                  <div title={`${d[dataKey]}${unit}`} style={{
                                    width:"70%",
                                    height: d[dataKey]>0 ? `${Math.max((d[dataKey]/maxVal)*100,4)}%` : "2px",
                                    background: d.isToday ? color : color+"80",
                                    borderRadius:"3px 3px 0 0",
                                    transition:"height 0.4s"
                                  }}/>
                                </div>
                              ))}
                            </div>
                          </div>
                          <div style={{display:"flex",gap:6,marginTop:4}}>
                            {vals.map((d,i) => (
                              <div key={i} style={{flex:1,textAlign:"center",fontSize:9,color:d.isToday?color:C.muted,fontWeight:d.isToday?700:400}}>{d.label}</div>
                            ))}
                          </div>
                          <div style={{display:"flex",gap:6,marginTop:2}}>
                            {vals.map((d,i) => (
                              <div key={i} style={{flex:1,textAlign:"center",fontSize:11}}>
                                {d[dataKey]>0 ? (isGood(d[dataKey]) ? <span style={{color:"#16a34a",fontWeight:700}}>✓</span> : <span style={{color:"#ef4444",fontWeight:700}}>✗</span>) : ""}
                              </div>
                            ))}
                          </div>
                          <div style={{fontSize:10,color:C.muted,marginTop:8}}>Meta: {meta}{unit}</div>
                        </div>
                      );
                    };
                    return (
                      <>
                        <MacroChart title="🍞 Carbohidratos" dataKey="carbs" color={C.sky} meta={metaCarbs} unit="g" isGood={v=>v<=metaCarbs} />
                        <MacroChart title="🍗 Proteína" dataKey="protein" color={C.green} meta={metaProtein} unit="g" isGood={v=>v>=metaProtein} />
                        <MacroChart title="🔥 Calorías" dataKey="kcal" color={C.orange} meta={metaKcal} unit="" isGood={v=>v<=metaKcal} />
                      </>
                    );
                  })()}
                  {/* Historial semanal — solo la semana actual, no semanas pasadas */}
                  {weeklyData.length > 0 && (() => {
                    const w = weeklyData[weeklyData.length-1];
                    return (
                      <div style={{background:C.card,borderRadius:16,padding:16,marginBottom:12}}>
                        <div style={{fontSize:13,fontWeight:700,color:C.muted,marginBottom:14}}>📅 SEMANA ACTUAL</div>
                        <div style={{fontSize:12,fontWeight:600,color:C.text,marginBottom:8}}>Semana del {w.label} <span style={{color:C.muted,fontWeight:400}}>· {w.days} día{w.days!==1?"s":""}</span></div>
                        <ProgressBar label="Carbs" value={w.carbs} meta={metaCarbs} color={C.sky} unit="g" />
                        <ProgressBar label="Proteína" value={w.protein} meta={metaProtein} color={C.green} unit="g" />
                        <ProgressBar label="Calorías" value={w.kcal} meta={metaKcal} color={C.orange} unit="" />
                      </div>
                    );
                  })()}
                  {/* Peso — al final */}
                  <div style={{background:C.card,borderRadius:16,padding:16,marginBottom:12}}>
                    <div style={{fontSize:13,fontWeight:700,color:C.muted,marginBottom:12}}>⚖️ PESO</div>
                    <div style={{display:"flex",gap:12,marginBottom:16}}>
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
                    {/* Registrar peso de hoy */}
                    <div style={{display:"flex",gap:8,marginBottom:16}}>
                      <input type="number" value={weightInput} onChange={e=>setWeightInput(e.target.value)}
                        placeholder="Registrar peso de hoy (kg)"
                        style={{...inp,flex:1}} />
                      <button onClick={doSaveWeight}
                        style={{background:weightSaved?"#16a34a":C.blue,border:"none",color:"white",borderRadius:12,padding:"0 18px",fontSize:14,fontWeight:700,cursor:"pointer",flexShrink:0}}>
                        {weightSaved ? "✓" : "Guardar"}
                      </button>
                    </div>
                    {/* Gráfica de peso promedio por semana */}
                    {(() => {
                      const weightWeekMap = {};
                      for (const w of weightLog) {
                        try {
                          const wk = getWeekKey(parseDate(w.date));
                          if (!weightWeekMap[wk]) weightWeekMap[wk] = [];
                          weightWeekMap[wk].push(w.kg);
                        } catch {}
                      }
                      const weightWeeks = Object.entries(weightWeekMap)
                        .sort((a,b) => new Date(a[0]) - new Date(b[0]))
                        .slice(-8)
                        .map(([wk, kgs]) => ({
                          label: wk.slice(0,5),
                          avg: Math.round((kgs.reduce((s,v)=>s+v,0)/kgs.length)*10)/10,
                        }));
                      if (weightWeeks.length === 0) {
                        return (
                          <div style={{fontSize:12,color:C.muted,textAlign:"center",padding:12}}>
                            Aún no hay registros de peso semanales. Regístrate arriba para empezar a ver tu progreso.
                          </div>
                        );
                      }
                      const values = weightWeeks.map(w=>w.avg);
                      const minVal = Math.min(...values, pesoMeta) - 2;
                      const maxVal = Math.max(...values, pesoMeta) + 2;
                      const range = Math.max(maxVal - minVal, 1);
                      const chartH = 100;
                      const metaBottom = Math.min(Math.max(((pesoMeta-minVal)/range)*100,0),100);
                      return (
                        <div>
                          <div style={{fontSize:12,fontWeight:700,color:C.muted,marginBottom:8}}>Promedio semanal (últimas {weightWeeks.length} semanas)</div>
                          <div style={{position:"relative",height:chartH}}>
                            <div style={{position:"absolute",left:0,right:0,bottom:`${metaBottom}%`,borderTop:`1.5px dashed ${C.green}`,zIndex:2}} />
                            <div style={{position:"absolute",right:0,bottom:`${Math.min(metaBottom+3,94)}%`,fontSize:9,color:C.green,fontWeight:700}}>Meta {pesoMeta}kg</div>
                            <div style={{display:"flex",alignItems:"flex-end",gap:8,height:"100%"}}>
                              {weightWeeks.map((w,i) => (
                                <div key={i} style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",height:"100%",justifyContent:"flex-end"}}>
                                  <div style={{fontSize:9,color:C.blue,fontWeight:600,marginBottom:2}}>{w.avg}</div>
                                  <div style={{
                                    width:"70%",
                                    height:`${Math.max(((w.avg-minVal)/range)*100,4)}%`,
                                    background:C.blue,
                                    borderRadius:"4px 4px 0 0",
                                    transition:"height 0.4s"
                                  }}/>
                                </div>
                              ))}
                            </div>
                          </div>
                          <div style={{display:"flex",gap:8,marginTop:4}}>
                            {weightWeeks.map((w,i) => (
                              <div key={i} style={{flex:1,textAlign:"center",fontSize:9,color:C.muted}}>{w.label}</div>
                            ))}
                          </div>
                        </div>
                      );
                    })()}
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
                          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
                            <span style={{fontSize:12,color:C.muted}}>🕐 {r.time}</span>
                            <div style={{display:"flex",alignItems:"center",gap:8}}>
                              <span style={{fontSize:15,fontWeight:700,color:C.sky,fontFamily:"monospace"}}>
                                {r.insulin}U {settings.insulinaRapida||"Rápida"}{r.toujeo>0 && <span style={{color:"#7c3aed"}}> · {r.toujeo}U {settings.insulinaLenta||"Lenta"}</span>}
                              </span>
                              <button onClick={()=>{ if(window.confirm("¿Borrar este registro?")) deleteRecord(r.id); }}
                                style={{background:C.red+"20",border:"none",color:C.red,borderRadius:8,width:26,height:26,fontSize:13,cursor:"pointer",flexShrink:0}}>🗑️</button>
                            </div>
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
                {customFoods.map((f,idx)=>({f,idx})).sort((a,b)=>a.f.name.localeCompare(b.f.name,'es')).map(({f,idx:i}) => (
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
