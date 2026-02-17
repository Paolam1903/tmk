import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import os
from datetime import datetime

st.set_page_config(layout="wide")

# -------------------------
# USUARIOS
# -------------------------
USUARIOS = {
    "director": {"password": "1234", "rol": "director"},
    "admin": {"password": "2026+", "rol": "administrativo"}
}

# -------------------------
# LOGIN
# -------------------------
if "usuario" not in st.session_state:
    st.session_state.usuario = None
    st.session_state.rol = None

if st.session_state.usuario is None:
    st.title("🔐 Acceso sistema TMK - SER Comunicaciones")

    usuario = st.text_input("Usuario")
    password = st.text_input("Contraseña", type="password")

    if st.button("Ingresar"):
        if usuario in USUARIOS and USUARIOS[usuario]["password"] == password:
            st.session_state.usuario = usuario
            st.session_state.rol = USUARIOS[usuario]["rol"]
            st.rerun()
        else:
            st.error("Usuario o contraseña incorrectos")

    st.stop()

# -------------------------
# SIDEBAR CON LOGO
# -------------------------
st.sidebar.image("logo_ser.png", use_container_width=True)
st.sidebar.markdown("### Panel de filtros")

st.sidebar.write(f"👤 Usuario: {st.session_state.usuario}")

if st.sidebar.button("Cerrar sesión"):
    st.session_state.usuario = None
    st.session_state.rol = None
    st.rerun()

# -------------------------
# CARGAR DATOS
# -------------------------
df = pd.read_excel("presupuesto_TMK.xlsx", sheet_name="Datos")
df.columns = df.columns.str.strip()

# Si no existe la columna Pago cumplimiento la creamos
if "Pago cumplimiento" not in df.columns:
    df["Pago cumplimiento"] = "0%"

# -------------------------
# FILTROS EN SIDEBAR
# -------------------------
año = st.sidebar.selectbox("Año", sorted(df["Año"].unique()))

mes = st.sidebar.selectbox(
    "Mes",
    df[df["Año"] == año]["Mes"].unique()
)

df_filtrado = df[(df["Año"] == año) & (df["Mes"] == mes)].copy()

nombres = st.sidebar.multiselect(
    "Nombre asesoras",
    df_filtrado["Nombre"].unique(),
    default=df_filtrado["Nombre"].unique()
)

df_filtrado = df_filtrado[df_filtrado["Nombre"].isin(nombres)]

# -------------------------
# TÍTULO PRINCIPAL
# -------------------------
st.title(f"📊 Presupuesto TMK - {mes} {año}")

# -------------------------
# CÁLCULO CUMPLIMIENTO
# -------------------------
def calcular_cumplimiento(row):
    if row["Concepto"] in ["Pyme Móvil - salarial", "CLOUND - salarial"]:
        return "Sin meta"
    elif row["Meta"] == 0:
        return "0%"
    else:
        valor = round((row["Ejecutado"] / row["Meta"]) * 100, 1)
        return f"{valor}%"

df_filtrado["% Cumplimiento"] = df_filtrado.apply(calcular_cumplimiento, axis=1)

opciones_pago = ["100%", "90%", "0%"]


# -------------------------
# TABLA PRINCIPAL
# -------------------------
if st.session_state.rol == "director":
    st.success("Modo Director: puedes editar pagos")

    df_editado = st.data_editor(
        df_filtrado,
        column_config={
            "Pago cumplimiento": st.column_config.SelectboxColumn(
                "Pago cumplimiento",
                options=opciones_pago
            )
        },
        use_container_width=True,
        num_rows="fixed"
    )

    if st.button("💾 Guardar cambios"):

        archivo_base = "presupuesto_TMK.xlsx"
        archivo_hist = "historico_cambios_TMK.xlsx"

        # Cargar base completa
        df_base = pd.read_excel(archivo_base, sheet_name="Datos")

        # Actualizar solo los registros filtrados
        for index, row in df_editado.iterrows():

            condicion = (
                (df_base["Año"] == row["Año"]) &
                (df_base["Mes"] == row["Mes"]) &
                (df_base["Nombre"] == row["Nombre"]) &
                (df_base["Concepto"] == row["Concepto"])
            )

            df_base.loc[condicion, "Pago cumplimiento"] = row["Pago cumplimiento"]

        # Guardar archivo base actualizado
        df_base.to_excel(archivo_base, sheet_name="Datos", index=False)

        # Guardar histórico
        df_hist = df_editado.copy()
        df_hist["Usuario cambio"] = st.session_state.usuario
        df_hist["Fecha cambio"] = datetime.now()

        if os.path.exists(archivo_hist):
            hist = pd.read_excel(archivo_hist)
            hist = pd.concat([hist, df_hist], ignore_index=True)
        else:
            hist = df_hist

        hist.to_excel(archivo_hist, index=False)

        st.success("Cambios guardados y actualizados correctamente")

# -------------------------
# DESCARGAR HISTÓRICO
# -------------------------
if os.path.exists("historico_cambios_TMK.xlsx"):
    with open("historico_cambios_TMK.xlsx", "rb") as file:
        st.download_button(
            "📥 Descargar histórico",
            data=file,
            file_name="historico_cambios_TMK.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# -------------------------
# GRÁFICOS
# -------------------------
# -------------------------
# GRÁFICOS MEJORADOS
# -------------------------
st.markdown("## 📈 Competencia asesoras")

aux = df_editado[df_editado["Rol"] == "Auxiliar"]

def grafico_con_tendencia(data, titulo):

    if data.empty:
        st.warning("No hay datos disponibles")
        return

    fig, ax = plt.subplots()

    nombres = data["Nombre"]
    valores = data["Ejecutado"]

    barras = ax.bar(nombres, valores)

    # Tamaño letra pequeño
    ax.tick_params(axis='x', labelsize=8)
    ax.tick_params(axis='y', labelsize=8)
    ax.set_title(titulo, fontsize=10)

    # Etiquetas de datos
    for bar in barras:
        height = bar.get_height()
        ax.text(
            bar.get_x() + bar.get_width()/2,
            height,
            f'{int(height)}',
            ha='center',
            va='bottom',
            fontsize=8
        )

    # Línea de tendencia
    x = range(len(valores))
    z = pd.Series(valores).rolling(window=2, min_periods=1).mean()
    ax.plot(nombres, z, marker='o')

    st.pyplot(fig)


# ---- Ventas directas
venta_directa = aux[aux["Concepto"] == "Venta directa - salarial"]
vd = venta_directa.groupby("Nombre")["Ejecutado"].sum().reset_index()

grafico_con_tendencia(vd, "Ventas directas por asesora")

# ---- Referidos
referido = aux[aux["Concepto"] == "Referido - bono"]
ref = referido.groupby("Nombre")["Ejecutado"].sum().reset_index()

grafico_con_tendencia(ref, "Referidos por asesora")


# -------------------------
# GRÁFICO CORPORATIVO POR PRODUCTO
# -------------------------
st.markdown("## 🏢 Comportamiento corporativo por producto")

lider = df_editado[df_editado["Rol"] == "Líder"]

corp_producto = (
    lider.groupby("Concepto")[["Meta", "Ejecutado"]]
    .sum()
    .reset_index()
)

if not corp_producto.empty:

    fig2, ax2 = plt.subplots()

    x = range(len(corp_producto))

    barras1 = ax2.bar(
        [i - 0.2 for i in x],
        corp_producto["Meta"],
        width=0.4,
        label="Meta"
    )

    barras2 = ax2.bar(
        [i + 0.2 for i in x],
        corp_producto["Ejecutado"],
        width=0.4,
        label="Ejecutado"
    )

    ax2.set_xticks(x)
    ax2.set_xticklabels(corp_producto["Concepto"], fontsize=8)
    ax2.set_title("Meta vs Ejecutado por producto", fontsize=10)
    ax2.tick_params(axis='y', labelsize=8)

    # Etiquetas de datos
    for barras in [barras1, barras2]:
        for bar in barras:
            height = bar.get_height()
            ax2.text(
                bar.get_x() + bar.get_width()/2,
                height,
                f'{int(height)}',
                ha='center',
                va='bottom',
                fontsize=8
            )

    # Línea de tendencia sobre ejecutado
    tendencia = corp_producto["Ejecutado"].rolling(window=2, min_periods=1).mean()
    ax2.plot(x, tendencia, marker='o')

    ax2.legend(fontsize=8)

    st.pyplot(fig2)