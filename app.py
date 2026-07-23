import streamlit as st
from datetime import datetime
from consolidador import consolidar_desde_bytes

st.set_page_config(page_title="Consolidador de Solicitudes — Subalcaldías", page_icon="🗂️", layout="centered")

st.title("🗂️ Consolidador de Solicitudes — Subalcaldías")
st.caption("Dirección de Coordinación Distrital · GAM Santa Cruz de la Sierra")

st.markdown(
    "Sube tu **Matriz maestra** actual y las **plantillas llenadas** por las subalcaldías. "
    "La herramienta agrega solo las solicitudes nuevas (no duplica nada) y te devuelve el "
    "Excel listo para descargar."
)

st.divider()

col1, col2 = st.columns(2)
with col1:
    st.subheader("1. Matriz maestra")
    matriz_file = st.file_uploader("Tu archivo Matriz_Seguimiento_Solicitudes_Subalcaldias.xlsx", type=["xlsx"], key="matriz")

with col2:
    st.subheader("2. Plantillas llenadas")
    plantillas_files = st.file_uploader(
        "Puedes subir varias a la vez (una por subalcaldía)",
        type=["xlsx"], accept_multiple_files=True, key="plantillas"
    )

st.divider()

if st.button("🔄 Consolidar", type="primary", use_container_width=True, disabled=not (matriz_file and plantillas_files)):
    with st.spinner("Leyendo plantillas y actualizando la matriz..."):
        plantillas_bytes = [(pf.name, pf.getvalue()) for pf in plantillas_files]
        try:
            resultado_bytes, resumen = consolidar_desde_bytes(matriz_file.getvalue(), plantillas_bytes)
        except Exception as e:
            st.error(f"Algo falló al procesar los archivos: {e}")
            st.stop()

    st.success(f"Listo — {resumen['nuevas']} solicitud(es) nueva(s) agregada(s).")

    c1, c2, c3 = st.columns(3)
    c1.metric("Nuevas agregadas", resumen["nuevas"])
    c2.metric("Ya existían (omitidas)", resumen["saltadas"])
    c3.metric("Recalculado con LibreOffice", "Sí" if resumen["recalculado"] else "No — se calcula al abrir en Excel")

    if resumen["avisos"]:
        for a in resumen["avisos"]:
            st.warning(a)

    if resumen["agregadas"]:
        with st.expander(f"Ver el detalle de las {len(resumen['agregadas'])} solicitudes agregadas"):
            st.table(resumen["agregadas"])

    nombre_salida = f"Matriz_Consolidada_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
    st.download_button(
        "⬇️ Descargar Matriz consolidada",
        data=resultado_bytes,
        file_name=nombre_salida,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

    st.caption(
        "Si al abrir el Excel las columnas 'Solicitante' o 'Plazo Compromiso' aparecen vacías, "
        "presiona F9 (o Fórmulas → Calcular ahora) — Excel las recalcula solo."
    )

elif not matriz_file or not plantillas_files:
    st.info("Sube la Matriz maestra y al menos una plantilla llenada para poder consolidar.")

with st.expander("¿Cómo funciona esto? (ver detalle)"):
    st.markdown(
        """
        - Cada solicitud nueva se identifica por **archivo + distrito + N° interno**, así que si
          subes la misma plantilla dos veces, no se duplica nada.
        - Las columnas automáticas de la Matriz (Solicitante, Plazo Compromiso, Días Transcurridos)
          **no se tocan** — siguen calculándose solas con sus fórmulas.
        - Los desplegables de Distrito y Destino se reconstruyen automáticamente al guardar
          (Excel los guarda en un formato que las librerías de edición pueden perder).
        - Nada se guarda en un servidor: los archivos solo existen mientras dura esta sesión del
          navegador. Si cierras la pestaña sin descargar, se pierde el resultado.
        """
    )
