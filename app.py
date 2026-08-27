import streamlit as st
import openpyxl
import os
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Me Latte Café — Control de Asistencias WhatsApp", page_icon="☕", layout="wide")

EXCEL_PATH = "registro_asistencias_me_latte_cafe.xlsx"

def asegurar_archivo_excel():
    if not os.path.exists(EXCEL_PATH):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Registro Semanal"
        ws.append(["ME LATTE CAFÉ — CONTROL DE ASISTENCIAS"])
        ws.append(["Semana en curso"])
        ws.append([])
        ws.append([])
        ws.append([])
        headers = ["Día", "Fecha", "Nombre", "Puesto", "H. Entrada", "F. Entrada", "H. Salida", "F. Salida", "Retardo", "Observaciones"]
        ws.append(headers)
        wb.save(EXCEL_PATH)

asegurar_archivo_excel()

st.title("☕ Me Latte Café — Panel de Control de Asistencias por WhatsApp")
st.markdown("Esta aplicación simula y gestiona la recepción de fotografías de asistencia enviadas por los colaboradores.")

st.sidebar.header("📱 Simulador de Entrada (WhatsApp)")
empleados_db = {
    "Carlos Ruiz (Barista)": {"tel": "5219931234567", "puesto": "Barista"},
    "Ana Gómez (Cajera)": {"tel": "5219939876543", "puesto": "Cajera"},
    "Luis Pérez (Cocinero)": {"tel": "5219935554433", "puesto": "Cocinero (Chilaquiles)"},
    "María López (Mesera)": {"tel": "5219932221100", "puesto": "Mesera"}
}

selected_emp = st.sidebar.selectbox("Seleccionar Empleado:", list(empleados_db.keys()))
uploaded_photo = st.sidebar.file_uploader("Sube la fotografía (Selfie de asistencia)", type=["jpg", "jpeg", "png"])

if st.sidebar.button("Simular Envío de Asistencia", type="primary"):
    if uploaded_photo is not None:
        now = datetime.now()
        date_str = now.strftime("%d/%m/%Y")
        time_str = now.strftime("%H:%M:%S")
        dia_nombre = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"][now.weekday()]
        
        emp_info = empleados_db[selected_emp]
        nombre = selected_emp.split(" (")[0]
        puesto = emp_info["puesto"]
        
        os.makedirs("asistencias_fotos", exist_ok=True)
        filename = f"asistencias_fotos/{emp_info['tel']}_{now.strftime('%Y%m%d_%H%M%S')}.jpg"
        with open(filename, "wb") as f:
            f.write(uploaded_photo.getbuffer())
            
        wb = openpyxl.load_workbook(EXCEL_PATH)
        ws = wb["Registro Semanal"]
        
        target_row = None
        for r in range(7, 100):
            val = ws.cell(row=r, column=3).value
            if val is None or val == "":
                target_row = r
                break
                
        if target_row:
            while ws.max_row < target_row:
                ws.append([])
                
            ws.cell(row=target_row, column=1, value=dia_nombre)
            ws.cell(row=target_row, column=2, value=date_str)
            ws.cell(row=target_row, column=3, value=nombre)
            ws.cell(row=target_row, column=4, value=puesto)
            ws.cell(row=target_row, column=5, value=time_str)
            ws.cell(row=target_row, column=6, value="✔ Verificado (WhatsApp)")
            ws.cell(row=target_row, column=10, value=f"Evidencia: {filename}")
            wb.save(EXCEL_PATH)
            
            st.sidebar.success(f"¡Asistencia de {nombre} registrada a las {time_str}!")
        else:
            st.sidebar.error("No se encontró espacio disponible.")
    else:
        st.sidebar.warning("Por favor, sube una fotografía.")

st.subheader("📊 Estado Actual del Registro de Asistencias")
if os.path.exists(EXCEL_PATH):
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    ws = wb["Registro Semanal"]
    
    data = []
    for r in range(7, ws.max_row + 1):
        row_vals = [ws.cell(row=r, column=c).value for c in range(1, 11)]
        if any(row_vals[:3]):
            data.append(row_vals)
            
    if data:
        df = pd.DataFrame(data, columns=["Día", "Fecha", "Nombre", "Puesto", "H. Entrada", "F. Entrada", "H. Salida", "F. Salida", "Retardo", "Observaciones"])
        st.dataframe(df, use_container_width=True)
    else:
        st.info("Aún no hay registros en la semana.")
