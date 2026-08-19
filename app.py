"""
Generador de Plantilla RIPS - version web (Streamlit).

Replica las dos rutas del menu principal de la version de escritorio:
  A) Crear plantilla RIPS desde archivos planos (3 pasos: generar,
     convertir a JSON, ajustar XML).
  B) Convertir una plantilla ya diligenciada a JSON (SISPRO), sin pasar
     por la generacion desde archivos planos.

Toda la logica de negocio vive en rips_core.py, sin modificar.
"""
import os
import tempfile
from datetime import date

import streamlit as st

import rips_core

TEMPLATE_PATH = os.path.join("resources", "plantilla_fev-rips_v_1_0.xlsm")

st.set_page_config(page_title="Generador RIPS - MICRODIAG", page_icon="🧾", layout="wide")

# ---------------------------------------------------------------------------
# Estado persistente entre pasos (Streamlit vuelve a correr todo el script
# en cada clic, asi que lo que se genera en un paso se guarda aqui para que
# el siguiente paso lo pueda usar).
# ---------------------------------------------------------------------------
if "xlsm_bytes" not in st.session_state:
    st.session_state.xlsm_bytes = None
    st.session_state.xlsm_name = None
    st.session_state.resumen = None
    st.session_state.warnings = None
    st.session_state.json_bytes = None
    st.session_state.json_name = None

# ---------------------------------------------------------------------------
# Menu principal
# ---------------------------------------------------------------------------
st.sidebar.title("🧾 Generador RIPS")
st.sidebar.caption("Laboratorio Especializado MICRODIAG")
opcion = st.sidebar.radio(
    "¿Qué deseas hacer?",
    [
        "Crear plantilla desde archivos planos",
        "Convertir plantilla ya diligenciada a JSON",
    ],
)
st.sidebar.divider()
st.sidebar.caption("Creado por: Christian Duque - Contador Público - 3134380447")

if not os.path.exists(TEMPLATE_PATH):
    st.error(
        f"No se encontro la plantilla oficial en `{TEMPLATE_PATH}`. "
        "Sube el archivo `plantilla_fev-rips_v_1_0.xlsm` dentro de una carpeta "
        "`resources/` en el repositorio de GitHub."
    )
    st.stop()

PREFIJOS_TXT = {"AF": "af0", "US": "us0", "AP": "ap0", "AD": "ad0", "CT": "ct0"}


def _guardar_temp(carpeta, archivo):
    ruta = os.path.join(carpeta, archivo.name)
    with open(ruta, "wb") as f:
        f.write(archivo.getbuffer())
    return ruta


# ===========================================================================
# RUTA A: Crear plantilla desde archivos planos (3 pasos)
# ===========================================================================
if opcion == "Crear plantilla desde archivos planos":
    st.title("🧾 Crear plantilla RIPS desde archivos planos")

    tab1, tab2, tab3 = st.tabs(
        ["1️⃣ Generar plantilla", "2️⃣ Convertir a JSON (SISPRO)", "3️⃣ Ajustar XML factura"]
    )

    # -----------------------------------------------------------------
    # PASO 1: Generar plantilla
    # -----------------------------------------------------------------
    with tab1:
        st.subheader("Datos de la factura")
        col1, col2, col3 = st.columns(3)
        with col1:
            prefijo = st.text_input("Prefijo", placeholder="Ej: FEDI", key="prefijo")
        with col2:
            consecutivo = st.text_input("Consecutivo", placeholder="Ej: 38841", key="consecutivo")
        with col3:
            tipo_usuario_label = st.selectbox(
                "Tipo de usuario", list(rips_core.TIPO_USUARIO_OPCIONES.keys()), key="tipo_usuario"
            )
        tipo_usuario = rips_core.TIPO_USUARIO_OPCIONES[tipo_usuario_label]
        num_factura = f"{prefijo.strip()}{consecutivo.strip()}"

        st.subheader("Archivos RIPS (TXT)")
        st.caption(
            "Sube AF0/US0/AP0/AD0 juntos (obligatorios) y CT0 si lo tienes (opcional). "
            "Se identifican por el prefijo del nombre del archivo."
        )
        archivos_txt = st.file_uploader(
            "Selecciona los archivos TXT", type=["txt"], accept_multiple_files=True, key="txt_files"
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

            for key, label in (("af0", "AF0"), ("us0", "US0"), ("ap0", "AP0"), ("ad0", "AD0"), ("ct0", "CT0")):
                if key in asignados:
                    st.success(f"{label}: {asignados[key].name}")
                elif key != "ct0":
                    st.warning(f"{label}: no cargado (obligatorio)")
                else:
                    st.caption("CT0: no cargado (opcional)")
            if no_reconocidos:
                st.error("No se reconocio el prefijo de: " + ", ".join(no_reconocidos))

        st.subheader("Copagos / cuotas moderadoras / pagos compartidos (opcional)")
        archivo_pago = st.file_uploader(
            "Reporte de ordenes por empresa (Excel)", type=["xls", "xlsx"], key="pago"
        )

        st.divider()
        if st.button("Generar plantilla", type="primary", key="btn_generar"):
            faltantes = [k for k in ("af0", "us0", "ap0", "ad0") if k not in asignados]
            if faltantes:
                st.error("Debes cargar: " + ", ".join(f.upper() for f in faltantes))
            elif not prefijo.strip() or not consecutivo.strip():
                st.error("Debes ingresar el prefijo y el consecutivo de la factura.")
            else:
                with st.spinner("Procesando..."):
                    with tempfile.TemporaryDirectory() as tmp:
                        rutas = {k: _guardar_temp(tmp, a) for k, a in asignados.items()}
                        ruta_pago = _guardar_temp(tmp, archivo_pago) if archivo_pago else None
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
                            with open(output_path, "rb") as f:
                                st.session_state.xlsm_bytes = f.read()
                            st.session_state.xlsm_name = f"RIPS_{num_factura}.xlsm"
                            st.session_state.resumen = resumen
                            st.session_state.warnings = warnings
                            # un archivo nuevo invalida el JSON del paso anterior
                            st.session_state.json_bytes = None
                            st.session_state.json_name = None

        if st.session_state.resumen:
            resumen = st.session_state.resumen
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
            if st.session_state.warnings:
                st.write("**Advertencias:**")
                for w in st.session_state.warnings:
                    st.write(f"- {w}")
            st.download_button(
                "Descargar plantilla generada",
                data=st.session_state.xlsm_bytes,
                file_name=st.session_state.xlsm_name,
                mime="application/vnd.ms-excel.sheet.macroEnabled.12",
            )

    # -----------------------------------------------------------------
    # PASO 2: Convertir a JSON (SISPRO)
    # -----------------------------------------------------------------
    with tab2:
        if not st.session_state.xlsm_bytes:
            st.info("Primero genera la plantilla en el Paso 1.")
        else:
            st.write(f"Plantilla lista para convertir: **{st.session_state.xlsm_name}**")
            if st.button("Convertir plantilla a JSON", type="primary", key="btn_json"):
                with st.spinner("Subiendo a SISPRO... puede tardar hasta un minuto."):
                    with tempfile.TemporaryDirectory() as tmp:
                        xlsm_path = os.path.join(tmp, st.session_state.xlsm_name)
                        with open(xlsm_path, "wb") as f:
                            f.write(st.session_state.xlsm_bytes)
                        try:
                            output_json = rips_core.convertir_json_sispro(xlsm_path)
                        except rips_core.RipsError as exc:
                            st.error(str(exc))
                        except Exception as exc:
                            st.error(f"Error inesperado: {exc}")
                        else:
                            with open(output_json, "rb") as f:
                                st.session_state.json_bytes = f.read()
                            st.session_state.json_name = os.path.basename(output_json)

            if st.session_state.json_bytes:
                st.success(f"JSON generado: {st.session_state.json_name}")
                st.download_button(
                    "Descargar JSON",
                    data=st.session_state.json_bytes,
                    file_name=st.session_state.json_name,
                    mime="application/json",
                )

    # -----------------------------------------------------------------
    # PASO 3: Ajustar XML de la factura electronica
    # -----------------------------------------------------------------
    with tab3:
        st.write(
            "Ajusta en el XML de la factura electronica los campos MODALIDAD_PAGO y "
            "COBERTURA_PLAN_BENEFICIOS segun el tipo de usuario seleccionado en el Paso 1."
        )
        xml_factura = st.file_uploader("XML de la factura electronica", type=["xml"], key="xml_factura")

        if st.button("Ajustar XML", type="primary", key="btn_xml"):
            if xml_factura is None:
                st.error("Debes cargar el XML de la factura electronica.")
            elif not st.session_state.get("prefijo", "").strip() or not st.session_state.get("consecutivo", "").strip():
                st.error("Debes ingresar el prefijo y el consecutivo en el Paso 1.")
            else:
                with st.spinner("Ajustando XML..."):
                    with tempfile.TemporaryDirectory() as tmp:
                        xml_path = _guardar_temp(tmp, xml_factura)
                        nombre_salida = os.path.splitext(xml_factura.name)[0] + "_ajustado.xml"
                        xml_destino = os.path.join(tmp, nombre_salida)
                        tipo_usuario = rips_core.TIPO_USUARIO_OPCIONES[st.session_state["tipo_usuario"]]
                        try:
                            rips_core.ajustar_xml_factura(xml_path, tipo_usuario, xml_destino)
                        except rips_core.RipsError as exc:
                            st.error(str(exc))
                        except Exception as exc:
                            st.error(f"Error inesperado: {exc}")
                        else:
                            with open(xml_destino, "rb") as f:
                                xml_bytes = f.read()
                            st.success("XML ajustado con exito.")
                            st.download_button(
                                "Descargar XML ajustado",
                                data=xml_bytes,
                                file_name=nombre_salida,
                                mime="application/xml",
                            )
                            if st.session_state.json_bytes:
                                st.info(
                                    "Recuerda descargar tambien el JSON del Paso 2 — "
                                    "en la version de escritorio ambos quedaban juntos en la misma carpeta."
                                )

# ===========================================================================
# RUTA B: Convertir plantilla ya diligenciada a JSON (independiente)
# ===========================================================================
else:
    st.title("🔄 Convertir plantilla RIPS a JSON (SISPRO)")

    st.subheader("1. Descarga la plantilla oficial (si aún no la tienes)")
    st.caption("Descarga directamente el archivo oficial de la plantilla RIPS MINSALUD vigente desde SISPRO.")
    if st.button("⬇️ Descargar plantilla desde SISPRO"):
        with st.spinner("Descargando..."):
            with tempfile.TemporaryDirectory() as tmp:
                destino = os.path.join(tmp, "plantilla_fev-rips_v_1_0.xlsm")
                try:
                    rips_core.descargar_plantilla_sispro(destino)
                except rips_core.RipsError as exc:
                    st.error(str(exc))
                else:
                    with open(destino, "rb") as f:
                        st.download_button(
                            "Guardar plantilla descargada",
                            data=f.read(),
                            file_name="plantilla_fev-rips_v_1_0.xlsm",
                            mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                        )

    st.divider()
    st.subheader("2. Carga la plantilla ya diligenciada")
    plantilla_subida = st.file_uploader("Plantilla RIPS (.xlsm o .xlsx)", type=["xlsm", "xlsx"], key="plantilla_b")

    st.subheader("3. Convierte la plantilla a JSON")
    if st.button("Convertir plantilla a JSON", type="primary", key="btn_json_b"):
        if plantilla_subida is None:
            st.error("Debes cargar la plantilla que quieres convertir.")
        else:
            with st.spinner("Subiendo a SISPRO... puede tardar hasta un minuto."):
                with tempfile.TemporaryDirectory() as tmp:
                    ruta = _guardar_temp(tmp, plantilla_subida)
                    try:
                        output_json = rips_core.convertir_json_sispro(ruta)
                    except rips_core.RipsError as exc:
                        st.error(str(exc))
                    except Exception as exc:
                        st.error(f"Error inesperado: {exc}")
                    else:
                        with open(output_json, "rb") as f:
                            json_bytes = f.read()
                        st.success("JSON generado con exito.")
                        st.download_button(
                            "Descargar JSON",
                            data=json_bytes,
                            file_name=os.path.basename(output_json),
                            mime="application/json",
                        )
