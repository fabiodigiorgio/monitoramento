
import streamlit as st
import pandas as pd
import os
from datetime import datetime

st.title("📊 Monitoramento de Status - Simples")

CSV_BANCO = "banco_monitoramento.csv"

# Carregar banco existente ou criar novo
if os.path.exists(CSV_BANCO):
    banco_df = pd.read_csv(CSV_BANCO, parse_dates=["Data_Arquivo"])
else:
    banco_df = pd.DataFrame()

uploaded_files = st.file_uploader("📤 Envie as planilhas de monitoramento direcionado", type="xlsx", accept_multiple_files=True)

if uploaded_files:
    st.success(f"{len(uploaded_files)} arquivo(s) carregado(s). Processando...")

    for file in uploaded_files:
        nome = file.name
        try:
            data_ref = pd.to_datetime(nome.split()[-1].replace(".xlsx", ""), dayfirst=True).date()
        except:
            st.error(f"Não foi possível extrair data do nome do arquivo: {nome}")
            continue
        df = pd.read_excel(file, sheet_name="Base")
        df["Data_Arquivo"] = data_ref
        banco_df = pd.concat([banco_df, df], ignore_index=True)

    banco_df.drop_duplicates(subset=["OS ID", "Data_Arquivo"], keep="last", inplace=True)
    banco_df.to_csv(CSV_BANCO, index=False)
    st.success("📁 Banco de dados atualizado com sucesso!")

df_total = banco_df.copy()

if not df_total.empty:
    modelo_base = pd.read_excel("modelo_base.xlsx", sheet_name="monitoramento")
    colunas_fixas = ['OS ID', 'TAT', 'Cod Autorizada', 'Modelo', 'Número de Série',
                     'Status OS', 'Meta', 'Fora do Prazo', 'Entrega da Peça']
    colunas_status = [col for col in modelo_base.columns if col not in colunas_fixas]

    df_m = modelo_base.copy()
    df_m = df_m[df_m["OS ID"].isin(df_total["OS ID"])]
    df_long = df_m.melt(id_vars=["OS ID", "Status OS", "Meta", "Cod Autorizada"],
                        value_vars=colunas_status, var_name="Status", value_name="Dias")
    df_long["Status"] = df_long["Status"].str.strip().str.upper()
    df_long["Status Atual"] = df_long["Status OS"].str.strip().str.upper()
    df_long["Dias"] = pd.to_numeric(df_long["Dias"], errors="coerce").fillna(0)
    df_long["Meta"] = pd.to_numeric(df_long["Meta"], errors="coerce").fillna(0)
    df_long["Dias em Atraso"] = df_long.apply(
        lambda x: x["Dias"] - x["Meta"] if x["Status"] == x["Status Atual"] and x["Dias"] > x["Meta"] else 0, axis=1
    )
    df_atrasos = df_long[df_long["Dias em Atraso"] > 0].copy()

    ranking_postos = df_atrasos.groupby("Cod Autorizada").agg(
        Quantidade_OS=("OS ID", "nunique"),
        Total_Dias_Atraso=("Dias em Atraso", "sum")
    ).reset_index().sort_values("Total_Dias_Atraso", ascending=False)

    ranking_os = df_atrasos.groupby("OS ID").agg(
        Total_Dias_Atraso=("Dias em Atraso", "sum")
    ).reset_index().sort_values("Total_Dias_Atraso", ascending=False)

    ranking_status = df_atrasos.groupby("Status").agg(
        Total_Dias_Atraso=("Dias em Atraso", "sum"),
        Quantidade_OS=("OS ID", "nunique")
    ).reset_index().sort_values("Total_Dias_Atraso", ascending=False)

    data_hoje = datetime.today().strftime("%d-%m-%Y")
    output_excel_path = f"Monitoramento STATUS {data_hoje}.xlsx"
    with pd.ExcelWriter(output_excel_path, engine="xlsxwriter") as writer:
        df_m.to_excel(writer, sheet_name="monitoramento", index=False)
        ranking_postos.to_excel(writer, sheet_name="Ranking Postos Atraso", index=False)
        ranking_os.to_excel(writer, sheet_name="Ranking OS Atraso", index=False)
        ranking_status.to_excel(writer, sheet_name="Ranking Status Atraso", index=False)

    with open(output_excel_path, "rb") as f:
        st.download_button("📥 Baixar Relatório Final", f, file_name=output_excel_path)
else:
    st.info("Nenhum dado disponível no momento. Faça upload de planilhas para iniciar.")
