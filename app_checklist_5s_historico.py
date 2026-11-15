import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import gspread
from google.oauth2.service_account import Credentials
import datetime
import io
import openpyxl
st.set_page_config(page_title="Sistema Integral 5S - Santgar", layout="wide", page_icon="??")
st.title("?? Sistema Integral de Auditoría 5S — Santgar")
st.caption("Captura consolidada (1 a 4 checklists), Dictamen General, Radar 5S y Control Histórico de Avance")
# --- Función para generar el Radar 5S idéntico al tablero físico de Santgar ---
def generar_radar_santgar(niveles_dict):
fig, ax = plt.subplots(figsize=(6, 6), dpi=200)
ax.set_aspect('equal')
ax.axis('off')
fig.patch.set_facecolor('white')
NAVY = '#1D3273'
LIGHT_BLUE_PENT = '#9BA4C4'
ax.add_patch(patches.Rectangle((-5.5, 5.2), 11, 0.35, color=NAVY))
ax.text(0, 4.5, "RADAR 5S - GENERAL", ha='center', va='center', fontsize=18, fontweight='bold', color=NAVY)
arms = [
("ORDENAR", 126, niveles_dict.get("ORDENAR", 0.0)),
("LIMPIAR", 54, niveles_dict.get("LIMPIAR", 0.0)),
("MANTENER", 342, niveles_dict.get("MANTENER", 0.0)),
("DISCIPLINAR", 270, niveles_dict.get("DISCIPLINAR", 0.0)),
("CLASIFICAR", 198, niveles_dict.get("CLASIFICAR", 0.0)),
]
pent_angles_rad = np.deg2rad([90, 18, 306, 234, 162])
for lvl in [1, 2, 3, 4]:
r_pent = 0.8 + lvl * 0.78
xs = r_pent * np.cos(pent_angles_rad)
ys = r_pent * np.sin(pent_angles_rad)
ax.add_patch(patches.Polygon(np.column_stack([xs, ys]), closed=True, fill=False, edgecolor=LIGHT_BLUE_PENT, linewidth=1.6, zorder=1))
arm_width = 0.65
r_start = 0.72
step = 0.72
for label, deg, score in arms:
rad = np.deg2rad(deg)
ux, uy = np.cos(rad), np.sin(rad)
nx, ny = -uy, ux
for box_idx in range(1, 5):
r0 = r_start + (box_idx - 1) * step
r1 = r_start + box_idx * step
p1 = (r0 * ux - (arm_width/2) * nx, r0 * uy - (arm_width/2) * ny)
p2 = (r1 * ux - (arm_width/2) * nx, r1 * uy - (arm_width/2) * ny)
p3 = (r1 * ux + (arm_width/2) * nx, r1 * uy + (arm_width/2) * ny)
p4 = (r0 * ux + (arm_width/2) * nx, r0 * uy + (arm_width/2) * ny)
if score >= box_idx:
facecolor = '#3B66BC'
elif score > (box_idx - 1):
facecolor = '#B8C9EC'
else:
facecolor = 'white'
ax.add_patch(patches.Polygon([p1, p2, p3, p4], closed=True, facecolor=facecolor, edgecolor=NAVY, linewidth=1.6, zorder=3))
r_mid = (r0 + r1) / 2
offset_dir = -1 if deg in [126, 198, 270] else 1
num_x = r_mid * ux + offset_dir * (arm_width/2 + 0.25) * nx
num_y = r_mid * uy + offset_dir * (arm_width/2 + 0.25) * ny
ax.text(num_x, num_y, str(box_idx), ha='center', va='center', fontsize=8, fontweight='bold', color=NAVY, zorder=4)
r_tip = r_start + 4 * step + 0.35
tx, ty = r_tip * ux, r_tip * uy
rot = -72 if deg == 198 else (72 if deg == 342 else (36 if deg == 126 else (-36 if deg == 54 else 0)))
ax.text(tx, ty, label, ha='center', va='center', fontsize=10.5, fontweight='bold', color=NAVY, rotation=rot, zorder=5)
cxs = 0.78 * np.cos(pent_angles_rad)
cys = 0.78 * np.sin(pent_angles_rad)
ax.add_patch(patches.Polygon(np.column_stack([cxs, cys]), closed=True, facecolor=NAVY, edgecolor=NAVY, linewidth=1.5, zorder=6))
ax.text(0, -0.04, "5S", ha='center', va='center', fontsize=16, fontweight='bold', color='white', zorder=7)
ax.text(3.8, -4.6, "Santgar", ha='right', va='center', fontsize=14, fontweight='bold', color=NAVY)
ax.set_xlim(-5.2, 5.2)
ax.set_ylim(-5.2, 5.5)
return fig
secciones = {
"1ª S — Seleccionar": [
(1, "Sobre los escritorios solo hay lo necesario para el trabajo del día."),
(2, "No hay documentos de asuntos ya cerrados en escritorios ni bandejas."),
(3, "No hay equipos, mobiliario ni accesorios descompuestos u obsoletos."),
(4, "Papelería y consumibles en cantidad razonable, sin acumulación 'por si acaso'."),
(5, "No hay versiones obsoletas ni copias no controladas de PNO, formatos o especificaciones."),
(6, "No hay cajas, carpetas u objetos sin identificar en el piso, sobre archiveros o bajo escritorios.")
],
"2ª S — Ordenar": [
(7, "Cada archivero, cajón y repisa está identificado con su contenido."),
(8, "Las carpetas y expedientes tienen identificación uniforme en el lomo."),
(9, "La documentación sigue un criterio de orden claro (proceso, producto, año)."),
(10, "Los documentos en revisión tienen un lugar definido e identificado por estatus."),
(11, "El material de uso común tiene un lugar definido y está en él."),
(12, "Las carpetas compartidas tienen una estructura acordada y nombres de archivo uniformes."),
(13, "Prueba del cronómetro: los 5 documentos buscados se encontraron en el tiempo meta.")
],
"3ª S — Limpiar": [
(14, "Escritorios, superficies y equipos de cómputo están limpios."),
(15, "Piso, rincones y debajo de los muebles sin polvo, basura ni papeles."),
(16, "Hay botes de basura y de destrucción confidencial, sin desbordarse."),
(17, "La zona de impresora está limpia y sin impresiones abandonadas."),
(18, "La sala de juntas y las áreas comunes están en orden."),
(19, "Mobiliario, sillas e iluminación en buen estado.")
],
"4ª S — Estandarizar": [
(20, "Cada zona tiene su estándar visible (foto del estado ideal)."),
(21, "Existe un plan de limpieza con responsables y frecuencia, y se cumple."),
(22, "Existe un criterio acordado para nombrar y guardar archivos electrónicos."),
(23, "Las personas del área conocen las reglas acordadas (preguntar a 2 personas).")
],
"5ª S — Auditar y sostener": [
(24, "Existe un plan de auditorías 5S con fechas, y se está cumpliendo."),
(25, "Las acciones de la auditoría anterior están cerradas en tiempo."),
(26, "El panel 5S del área está actualizado."),
(27, "El estado se mantiene en el día a día (no se ve 'arreglado para la auditoría').")
],
"Seguridad": [
(28, "Pasillos y salidas de emergencia libres."),
(29, "Extintor accesible, señalizado y con revisión vigente."),
(30, "Cables canalizados, sin riesgo de tropiezo ni contactos sobrecargados."),
(31, "No hay objetos pesados en las partes altas de repisas o archiveros.")
]
}
# --- Datos Generales de la Auditoría ---
with st.expander("?? Datos Generales de la Auditoría", expanded=True):
col1, col2, col3 = st.columns(3)
with col1:
area_general = st.text_input("Área General Auditada:", "Oficinas de Calidad")
fecha = st.date_input("Fecha de Auditoría")
with col2:
auditores = st.text_input("Auditores:", "O. Montiel")
tipo = st.selectbox("Tipo de Auditoría:", ["Inicial", "Periódica", "Cierre de Piloto"])
with col3:
num_checklists = st.slider("Cantidad de Checklists (Zonas) a integrar:", min_value=1, max_value=4, value=2)
# --- Pestañas de Captura y Consolidado ---
tab_names = [f"?? Checklist {i+1}" for i in range(num_checklists)] + ["?? Dictamen General", "?? Histórico y Tendencia de Avance"]
tabs = st.tabs(tab_names)
all_responses = {}
subareas = {}
for i in range(num_checklists):
with tabs[i]:
subareas[i] = st.text_input(f"Nombre de Sub-área / Zona (Checklist {i+1}):", value=f"Zona de Trabajo {i+1}", key=f"sub_{i}")
all_responses[i] = {}
for sec_name, preguntas in secciones.items():
st.markdown(f"#### {sec_name}")
for idx, preg in preguntas:
c1, c2, c3 = st.columns([3.5, 1.2, 2.3])
with c1:
st.write(f"**{idx}.** {preg}")
with c2:
all_responses[i][idx] = st.selectbox("Evaluación", options=["-", "C", "NC", "NA"], key=f"c{i}_q{idx}")
with c3:
st.text_input("Observación / Ubicación", key=f"c{i}_obs{idx}")
st.divider()
# --- Pestaña de Dictamen General ---
dictamen_tab_idx = num_checklists
with tabs[dictamen_tab_idx]:
st.subheader("?? Dictamen General Consolidado")
resumen_checklists = []
for i in range(num_checklists):
c_tot = sum(1 for q in range(1, 32) if all_responses[i].get(q) == "C")
nc_tot = sum(1 for q in range(1, 32) if all_responses[i].get(q) == "NC")
app_tot = c_tot + nc_tot
pct_i = (c_tot / app_tot * 100) if app_tot > 0 else 0.0
est_i = "Sin evaluar" if app_tot == 0 else ("APROBADO (=85%)" if pct_i >= 85 else "NO APROBADO (<85%)")
resumen_checklists.append({
"Checklist": f"Checklist {i+1}",
"Sub-área / Zona": subareas[i],
"Aplicables (C+NC)": app_tot,
"Conformes (C)": c_tot,
"% Cumplimiento": f"{pct_i:.1f}%",
"Dictamen Individual": est_i
})
st.markdown("#### 1. Desglose por Checklist Evaluado")
st.dataframe(pd.DataFrame(resumen_checklists), use_container_width=True)
mapeo_radar = {
"1ª S — Seleccionar": "CLASIFICAR",
"2ª S — Ordenar": "ORDENAR",
"3ª S — Limpiar": "LIMPIAR",
"4ª S — Estandarizar": "MANTENER",
"5ª S — Auditar y sostener": "DISCIPLINAR"
}
niveles_radar = {}
resumen_etapas = []
global_c, global_nc, global_app = 0, 0, 0
for sec_name, preguntas in secciones.items():
sec_c = 0
sec_nc = 0
for i in range(num_checklists):
sec_c += sum(1 for idx, _ in preguntas if all_responses[i].get(idx) == "C")
sec_nc += sum(1 for idx, _ in preguntas if all_responses[i].get(idx) == "NC")
sec_app = sec_c + sec_nc
sec_pct = (sec_c / sec_app * 100) if sec_app > 0 else 0.0
nivel = round((sec_pct / 100.0) * 4.0, 1)
if sec_name in mapeo_radar:
niveles_radar[mapeo_radar[sec_name]] = nivel
global_c += sec_c
global_nc += sec_nc
global_app += sec_app
resumen_etapas.append({
"Bloque / Etapa 5S": sec_name,
"Pilar Radar": mapeo_radar.get(sec_name, "—"),
"Aplicables Totales": sec_app,
"Conformes Totales": sec_c,
"% Global": f"{sec_pct:.1f}%",
"Nivel Radar (0-4)": nivel if sec_name in mapeo_radar else "—",
"Estatus": "Sin evaluar" if sec_app == 0 else ("CUMPLE (=85%)" if sec_pct >= 85 else "NO CUMPLE (<85%)")
})
global_pct = (global_c / global_app * 100) if global_app > 0 else 0.0
if global_app == 0:
    dictamen_final = "PENDIENTE DE EVALUACION"
elif global_pct >= 85:
    dictamen_final = "DICTAMEN FAVORABLE (>= 85%)"
else:
    dictamen_final = "DICTAMEN NO FAVORABLE (< 85%)"
st.markdown("#### 2. Resultado Consolidado y Dictamen Final")
col_t, col_r = st.columns([1.3, 1])
with col_t:
st.dataframe(pd.DataFrame(resumen_etapas), use_container_width=True)
m1, m2, m3 = st.columns(3)
m1.metric("Total Aplicables", global_app)
m2.metric("Total Conformes", global_c)
m3.metric("% Cumplimiento Global", f"{global_pct:.1f}%")
st.info(f"**DICTAMEN GENERAL DE AUDITORÍA:** {dictamen_final}")
with col_r:
st.pyplot(generar_radar_santgar(niveles_radar))
# --- Botón para Guardar y Respaldar Auditoría ---
st.markdown("---")
if st.button("?? Guardar Auditoría en Historial (Nube / Local) y Generar Respaldo"):
nuevo_registro = {
"Fecha": str(fecha),
"Área": area_general,
"Tipo": tipo,
"Auditores": auditores,
"Puntos Aplicables": global_app,
"Puntos Conformes": global_c,
"% Cumplimiento": round(global_pct, 2),
"Dictamen": dictamen_final,
"Clasificar (Nivel)": niveles_radar.get("CLASIFICAR", 0),
"Ordenar (Nivel)": niveles_radar.get("ORDENAR", 0),
"Limpiar (Nivel)": niveles_radar.get("LIMPIAR", 0),
"Mantener (Nivel)": niveles_radar.get("MANTENER", 0),
"Disciplinar (Nivel)": niveles_radar.get("DISCIPLINAR", 0)
}
if "historial_auditorias" not in st.session_state:
st.session_state["historial_auditorias"] = []
st.session_state["historial_auditorias"].append(nuevo_registro)
try:
scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
creds_dict = dict(st.secrets["gcp_service_account"])
creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
client = gspread.authorize(creds)
sheet_name = st.secrets["database"]["sheet_name"]
sheet = client.open(sheet_name).worksheet("Historial")
fila = list(nuevo_registro.values())
sheet.append_row(fila)
st.success("¡Auditoría guardada exitosamente en Google Sheets y respaldada!")
except Exception as e:
st.warning(f"Auditoría guardada en sesión local (Google Sheets no configurado aún o error: {e})")
st.success("¡Auditoría respaldada en sesión local correctamente!")
# --- Pestaña de Histórico y Tendencia ---
historico_tab_idx = num_checklists + 1
with tabs[historico_tab_idx]:
st.subheader("?? Histórico de Auditorías y Tendencia de Avance 5S")
st.caption("Consulta el progreso de implementación de las 5S a lo largo del tiempo para medir la mejora continua.")
historial_list = st.session_state.get("historial_auditorias", [])
if len(historial_list) == 0:
historial_list = [
{"Fecha": "2026-08-15", "Área": "Oficinas de Calidad", "Tipo": "Inicial", "Auditores": "O. Montiel", "Puntos Aplicables": 28, "Puntos Conformes":
10, "% Cumplimiento": 35.7, "Dictamen": "NO FAVORABLE", "Clasificar (Nivel)": 1.2, "Ordenar (Nivel)": 1.0, "Limpiar (Nivel)": 1.5, "Mantener (Nivel)":
0.4, "Disciplinar (Nivel)": 0.0},
{"Fecha": "2026-09-15", "Área": "Oficinas de Calidad", "Tipo": "Periódica", "Auditores": "O. Montiel", "Puntos Aplicables": 30, "Puntos
Conformes": 21, "% Cumplimiento": 70.0, "Dictamen": "NO FAVORABLE", "Clasificar (Nivel)": 2.4, "Ordenar (Nivel)": 2.2, "Limpiar (Nivel)": 2.8,
"Mantener (Nivel)": 1.6, "Disciplinar (Nivel)": 0.8},
{"Fecha": "2026-10-15", "Área": "Oficinas de Calidad", "Tipo": "Periódica", "Auditores": "B. Cayetano", "Puntos Aplicables": 31, "Puntos
Conformes": 27, "% Cumplimiento": 87.1, "Dictamen": "FAVORABLE", "Clasificar (Nivel)": 3.2, "Ordenar (Nivel)": 3.0, "Limpiar (Nivel)": 3.6, "Mantener
(Nivel)": 3.1, "Disciplinar (Nivel)": 2.5},
{"Fecha": "2026-11-06", "Área": "Oficinas de Calidad", "Tipo": "Cierre de Piloto", "Auditores": "O. Montiel", "Puntos Aplicables": 31, "Puntos
Conformes": 29, "% Cumplimiento": 93.5, "Dictamen": "FAVORABLE", "Clasificar (Nivel)": 3.8, "Ordenar (Nivel)": 3.7, "Limpiar (Nivel)": 3.9, "Mantener
(Nivel)": 3.5, "Disciplinar (Nivel)": 3.4}
]
df_hist = pd.DataFrame(historial_list)
st.markdown("#### Tabla Histórica de Auditorías Registradas")
st.dataframe(df_hist, use_container_width=True)
if len(df_hist) > 0:
st.markdown("#### Gráfica de Tendencia de Cumplimiento General (%)")
df_chart = df_hist.set_index("Fecha")[["% Cumplimiento"]]
st.line_chart(df_chart)
st.markdown("#### Evolución por Pilar 5S (Niveles 0 a 4)")
pillars_df = df_hist.set_index("Fecha")[["Clasificar (Nivel)", "Ordenar (Nivel)", "Limpiar (Nivel)", "Mantener (Nivel)", "Disciplinar (Nivel)"]]
st.line_chart(pillars_df)