"""
Generador de Plantilla RIPS - version web (Streamlit).

Paso 1: sube los archivos AF0/US0/AP0/AD0 (y opcionalmente CT0 y el Excel
de pago moderador), llena la plantilla oficial RIPS (.xlsm) usando
rips_core.py (sin ninguna modificacion) y permite descargar el resultado.

Los Pasos 2 (Convertir a JSON en SISPRO) y 3 (Ajustar XML) se agregan
en versiones siguientes de este mismo archivo.
"""
import os
import tempfile
from datetime import date

import streamlit as st

import rips_core

# ---------------------------------------------------------------------------
# Recursos: en Streamlit Cloud la app corre desde la raiz del repositorio,
# asi que estas rutas son relativas a esa raiz (misma carpeta "resources"
# que usabas con PyInstaller, solo que ahora no hace falta resource_path()
# porque no hay empaquetado en .exe).
# ---------------------------------------------------------------------------
TEMPLATE_PATH = os.path.join("resources", "plantilla_fev-rips_v_1_0.xlsm")

st.set_page_config(page_title="Generador RIPS - MICRODIAG", page_icon="🧾", layout="wide")

st.title("🧾 Generador de RIPS Electronicos - Laboratorio Especializado MICRODIAG")
st.caption("Paso 1: Generar plantilla RIPS a partir de los archivos planos")

if not os.path.exists(TEMPLATE_PATH):
    st.error(
        f"No se encontro la plantilla oficial en `{TEMPLATE_PATH}`. "
        "Sube el archivo `plantilla_fev-rips_v_1_0.xlsm` dentro de una carpeta "
        "`resources/` en el repositorio de GitHub."
    )
    st.stop()

# ---------------------------------------------------------------------------
# Datos de la factura
# ---------------------------------------------------------------------------
st.subheader("Datos de la factura")
col1, col2, col3 = st.columns(3)
with col1:
    prefijo = st.text_input("Prefijo", placeholder="Ej: FEDI")
with col2:
    consecutivo = st.text_input("Consecutivo", placeholder="Ej: 38841")
with col3:
    tipo_usuario_label = st.selectbox(
        "Tipo de usuario", list(rips_core.TIPO_USUARIO_OPCIONES.keys())
    )
tipo_usuario = rips_core.TIPO_USUARIO_OPCIONES[tipo_usuario_label]
num_factura = f"{prefijo.strip()}{consecutivo.strip()}"

# ---------------------------------------------------------------------------
# Archivos RIPS (TXT) - carga combinada, se identifican por prefijo del
# nombre igual que en la version de escritorio (_browse_todos_txt)
# ---------------------------------------------------------------------------
st.subheader("Archivos RIPS (TXT)")
st.caption("Sube AF0/US0/AP0/AD0 juntos (obligatorios) y CT0 si lo tienes (opcional). Se identifican por el prefijo del nombre del archivo.")

PREFIJOS_TXT = {"AF": "af0", "US": "us0", "AP": "ap0", "AD": "ad0", "CT": "ct0"}

archivos_txt = st.file_uploader(
    "Selecciona los archivos TXT",
    type=["txt"],
    accept_multiple_files=True,
)

asignados = {}
no_reconocidos = []
if archivos_txt:
    for archivo in archivos_txt:
        nombre = archivo.name.upper()
        prefijo_archivo = next((p for p in PREFIJOS_TXT if nombre.startswith(p)), None)
        if prefijo_archivo is None:
            no_reconocidos.append(archivo.name)
            continue
        asignados[PREFIJOS_TXT[prefijo_archivo]] = archivo

    col_ok, col_warn = st.columns(2)
    with col_ok:
        for key, label in (("af0", "AF0"), ("us0", "US0"), ("ap0", "AP0"), ("ad0", "AD0"), ("ct0", "CT0")):
            if key in asignados:
                st.success(f"{label}: {asignados[key].name}")
            elif key != "ct0":
                st.warning(f"{label}: no cargado (obligatorio)")
            else:
                st.caption("CT0: no cargado (opcional)")
    if no_reconocidos:
        with col_warn:
            st.error("No se reconocio el prefijo de: " + ", ".join(no_reconocidos))

# ---------------------------------------------------------------------------
# Pagos compartidos (opcional)
# ---------------------------------------------------------------------------
st.subheader("Copagos / cuotas moderadoras / pagos compartidos (opcional)")
archivo_pago = st.file_uploader(
    "Reporte de ordenes por empresa (Excel)", type=["xls", "xlsx"], key="pago"
)

# ---------------------------------------------------------------------------
# Generar
# ---------------------------------------------------------------------------
st.divider()

if st.button("Generar plantilla", type="primary"):
    faltantes = [k for k in ("af0", "us0", "ap0", "ad0") if k not in asignados]
    if faltantes:
        st.error("Debes cargar: " + ", ".join(f.upper() for f in faltantes))
    elif not prefijo.strip() or not consecutivo.strip():
        st.error("Debes ingresar el prefijo y el consecutivo de la factura.")
    else:
        with st.spinner("Procesando..."):
            with tempfile.TemporaryDirectory() as tmp:
                # Streamlit entrega los archivos subidos como objetos en
                # memoria; rips_core.py espera rutas de archivo, asi que
                # se guardan primero en una carpeta temporal.
                rutas = {}
                for key, archivo in asignados.items():
                    ruta = os.path.join(tmp, archivo.name)
                    with open(ruta, "wb") as f:
                        f.write(archivo.getbuffer())
                    rutas[key] = ruta

                ruta_pago = None
                if archivo_pago is not None:
                    ruta_pago = os.path.join(tmp, archivo_pago.name)
                    with open(ruta_pago, "wb") as f:
                        f.write(archivo_pago.getbuffer())

                output_path = os.path.join(tmp, f"RIPS_{num_factura}.xlsm")

                try:
                    resumen, warnings = rips_core.procesar(
                        af0_path=rutas["af0"],
                        us0_path=rutas["us0"],
                        ap0_path=rutas["ap0"],
                        ad0_path=rutas["ad0"],
                        template_path=TEMPLATE_PATH,
                        output_path=output_path,
                        num_factura=num_factura,
                        tipo_usuario=tipo_usuario,
                        ct0_path=rutas.get("ct0"),
                        pago_path=ruta_pago,
                        today=date.today(),
                    )
                except rips_core.RipsError as exc:
                    st.error(str(exc))
                except Exception as exc:
                    st.error(f"Error inesperado: {exc}")
                else:
                    st.success("Plantilla generada con exito.")

                    st.write(f"**Factura:** {resumen['numFactura']}")
                    st.write(f"**NIT obligado:** {resumen['numDocumentoIdObligado']}")
                    st.write(f"**Usuarios cargados:** {resumen['totalUsuarios']}")
                    st.write(f"**Registros de consultas cargados:** {resumen['totalConsultas']}")
                    st.write(f"**Suma vrServicio:** {resumen['sumaVrServicio']}")
                    st.write(f"**Valor total AD0:** {resumen['valorTotalAD0']}")
                    if resumen["cuadreOk"]:
                        st.success("Cuadre AD0 vs vrServicio: OK")
                    else:
                        st.warning("Cuadre AD0 vs vrServicio: NO CUADRA - revisar")

                    if warnings:
                        st.write("**Advertencias:**")
                        for w in warnings:
                            st.write(f"- {w}")

                    with open(output_path, "rb") as f:
                        st.download_button(
                            "Descargar plantilla generada",
                            data=f.read(),
                            file_name=f"RIPS_{num_factura}.xlsm",
                            mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                        )

st.divider()
st.caption("Creado por: Christian Duque - Contador Público - 3134380447")

