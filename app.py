import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta

st.set_page_config(layout="wide")
st.title("📊 Monitoramento STATUS - Somatório Diário com Histórico")

st.markdown("Este sistema **acumula os dias úteis por status** com base na **última planilha gerada** (`Monitoramento STATUS *.xlsx`) e soma com os novos arquivos enviados.")

uploaded_files = st.file_uploader("📄 Envie os novos arquivos diários (aba 'Base')", type=["xlsx"], accept_multiple_files=True)

# Carregar a planilha de metas do diretório atual
try:
    df_metas = pd.read_excel("meta.xlsx", sheet_name="meta")
    df_metas.columns = ["STATUS OS", "Meta"]
    df_metas["STATUS OS"] = df_metas["STATUS OS"].astype(str).str.strip().str.upper()
    mapa_metas = dict(zip(df_metas["STATUS OS"], df_metas["Meta"]))
except Exception as e:
    st.error(f"❌ Erro ao carregar a planilha de metas 'meta.xlsx': {e}")
    st.stop()

arquivos_existentes = [f for f in os.listdir() if f.startswith("Monitoramento STATUS") and f.endswith(".xlsx")]
arquivo_base = sorted(arquivos_existentes)[-1] if arquivos_existentes else None

if not arquivo_base:
    st.warning("⚠️ Nenhuma planilha anterior 'Monitoramento STATUS' encontrada. O sistema irá iniciar do zero.")

if uploaded_files:
    novas_planilhas = []
    data_plan_map = {}

    for file in uploaded_files:
        try:
            df = pd.read_excel(file, sheet_name="Base")
            df["Arquivo"] = file.name
            df["Data_Referencia"] = pd.to_datetime(df["Data Plan"], errors="coerce")
            df["Modelo"] = df.iloc[:, 5]
            df["Número de Série"] = df.iloc[:, 6]
            df["Data de Entrega Peça"] = df.iloc[:, 17]
            novas_planilhas.append(df)
            data_plan_map[file.name] = df["Data_Referencia"].max()
        except Exception as e:
            st.error(f"Erro ao processar {file.name}: {e}")

    if novas_planilhas:
        novas_planilhas.sort(key=lambda x: x["Data_Referencia"].iloc[0])
        historico_os = {}
        transicoes = []

        if arquivo_base:
            try:
                df_anterior = pd.read_excel(arquivo_base, sheet_name="Consolidado")
            except ValueError:
                df_anterior = pd.read_excel(arquivo_base)

            for _, row in df_anterior.iterrows():
                os_id = row["OS ID"]
                status_atual = row["Status OS"]
                dias_em_status = {col.replace("Dias em: ", ""): int(row[col]) for col in row.index if col.startswith("Dias em: ")}
                historico_os[os_id] = {
                    "Cod Autorizada": row["Cod Autorizada"],
                    "TAT": row.get("TAT", None),
                    "Status OS": status_atual,
                    "Ultimo Status": status_atual,
                    "Ultima Data": row.get("Ultima Data", None),
                    "Dias em Status": dias_em_status,
                    "Modelo": row.get("Modelo", None),
                    "Número de Série": row.get("Número de Série", None),
                    "Data de Entrega Peça": row.get("Data de Entrega Peça", None),
                    "Historico": []
                }

        for df in novas_planilhas:
            data_atual = df["Data_Referencia"].iloc[0].date()
            for _, row in df.iterrows():
                os_id = row["OS ID"]
                status = row["Status OS"]
                cod = row["Cod Autorizada"]
                tat = row.get("TAT", None)
                modelo = row.get("Modelo", None)
                serie = row.get("Número de Série", None)
                entrega = row.get("Data de Entrega Peça", None)

                if os_id not in historico_os:
                    historico_os[os_id] = {
                        "Cod Autorizada": cod,
                        "TAT": tat,
                        "Status OS": status,
                        "Ultimo Status": status,
                        "Ultima Data": data_atual,
                        "Dias em Status": {status: 1},
                        "Modelo": modelo,
                        "Número de Série": serie,
                        "Data de Entrega Peça": entrega,
                        "Historico": [(status, data_atual, None)]
                    }
                else:
                    registro = historico_os[os_id]
                    if registro["Ultimo Status"] == status:
                        dias_uteis = pd.bdate_range(end=data_atual, periods=2).size - 1
                        registro["Dias em Status"][status] = registro["Dias em Status"].get(status, 0) + dias_uteis
                    else:
                        if registro["Historico"]:
                            ultimo = registro["Historico"][-1]
                            registro["Historico"][-1] = (ultimo[0], ultimo[1], data_atual - timedelta(days=1))
                        registro["Historico"].append((status, data_atual, None))
                        registro["Dias em Status"][status] = 1
                        registro["Ultimo Status"] = status
                    registro["Status OS"] = status
                    registro["Ultima Data"] = data_atual
                    if tat and tat != registro["TAT"]:
                        registro["TAT"] = tat
                    if modelo: registro["Modelo"] = modelo
                    if serie: registro["Número de Série"] = serie
                    if entrega: registro["Data de Entrega Peça"] = entrega

        dados_final = []
        historico_status = []

        for os_id, dados in historico_os.items():
            status_atual = dados["Status OS"]
            dias_no_status = dados["Dias em Status"].get(status_atual, 1)
            meta = mapa_metas.get(status_atual.strip().upper(), None)

            base = {
                "OS ID": os_id,
                "Cod Autorizada": dados["Cod Autorizada"],
                "Modelo": dados["Modelo"],
                "Número de Série": dados["Número de Série"],
                "Data de Entrega Peça": dados["Data de Entrega Peça"],
                "TAT": int(dados["TAT"]) if dados["TAT"] is not None else None,
                "Status OS": status_atual,
                "Dias no Status": int(dias_no_status),
                "Meta": int(meta) if meta is not None else None,
                "Fora do Prazo": int(max(dias_no_status - meta, 0)) if meta is not None else None,
                "Ultima Data": dados["Ultima Data"]
            }

            for status, dias in dados["Dias em Status"].items():
                base[f"Dias em: {status}"] = int(dias)

            dados_final.append(base)

            for status, inicio, fim in dados["Historico"]:
                dias_uteis = len(pd.bdate_range(inicio, fim or dados["Ultima Data"]))
                meta_status = mapa_metas.get(status.strip().upper(), None)
                atraso = max(dias_uteis - meta_status, 0) if meta_status is not None else None
                historico_status.append({
                    "OS ID": os_id,
                    "Cod Autorizada": dados["Cod Autorizada"],
                    "Status": status,
                    "Data Início": inicio,
                    "Data Fim": fim or dados["Ultima Data"],
                    "Dias Úteis": int(dias_uteis),
                    "Meta": int(meta_status) if meta_status is not None else None,
                    "Dias em Atraso": int(atraso) if atraso is not None else None
                })

        df_resultado = pd.DataFrame(dados_final).fillna(0)
        df_historico = pd.DataFrame(historico_status)

        arquivo_mais_recente = max(data_plan_map.items(), key=lambda x: x[1])[0]
        df_mais_recente = [df for df in novas_planilhas if df["Arquivo"].iloc[0] == arquivo_mais_recente][0]
        os_ativas = set(df_mais_recente["OS ID"])
        data_mais_recente = max(data_plan_map.values()).strftime("%d-%m-%Y")

        df_resultado["Fechada"] = df_resultado["OS ID"].apply(lambda x: "Não" if x in os_ativas else "Sim")

        nome_saida = f"Monitoramento STATUS {data_mais_recente}.xlsx"
        
        # Gerar Rankings (somente OS ativas com atraso)
        df_abertas = df_resultado[(df_resultado["Fechada"] == "Não") & (df_resultado["Fora do Prazo"] > 0)]

        ranking_postos = df_abertas.groupby("Cod Autorizada")["OS ID"].count().reset_index()
        ranking_postos = ranking_postos.rename(columns={"OS ID": "Total OS em Atraso"})
        ranking_postos = ranking_postos.sort_values(by="Total OS em Atraso", ascending=False)

        ranking_os = df_abertas[["OS ID", "Cod Autorizada", "Status OS", "Fora do Prazo"]]
        ranking_os = ranking_os.sort_values(by="Fora do Prazo", ascending=False)

        ranking_status = df_abertas.groupby("Status OS")["OS ID"].count().reset_index()
        ranking_status = ranking_status.rename(columns={"OS ID": "Total OS em Atraso"})
        ranking_status = ranking_status.sort_values(by="Total OS em Atraso", ascending=False)

        with pd.ExcelWriter(nome_saida, engine="xlsxwriter") as writer:
            df_resultado.to_excel(writer, sheet_name="Consolidado", index=False)
            df_historico.to_excel(writer, sheet_name="Histórico por Status", index=False)
            ranking_postos.to_excel(writer, sheet_name="Ranking Postos", index=False)
            ranking_os.to_excel(writer, sheet_name="Ranking OS", index=False)
            ranking_status.to_excel(writer, sheet_name="Ranking Status", index=False)


        st.success(f"✅ Relatório atualizado com sucesso: {nome_saida}")
        st.dataframe(df_resultado)

        with open(nome_saida, "rb") as f:
            st.download_button("📅 Baixar Planilha Final", f, file_name=nome_saida)
