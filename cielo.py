import os
import pandas as pd
import streamlit as st
import logging
from datetime import datetime
from rapidfuzz import fuzz
from openpyxl import load_workbook


# Configuração de logging
logging.basicConfig(
    level=logging.DEBUG,  # ou DEBUG para mais detalhes
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler("conciliacao.log", encoding="utf-8"),  # grava em arquivo
        logging.StreamHandler()  # mostra no console
    ]
)


# =========================
# Função de limpeza ERP
# =========================
def limpar_erp(df):
    try:
        with st.spinner("🧹 Limpando dados do ERP..."):
            df["Emissão"] = pd.to_datetime(df["Emissão"], dayfirst=True, errors="coerce")
            parcelas = df["Numero"].str.extract(r"-(\d+)/(\d+)")
            df["Numero da Parcela"] = parcelas[0].astype(float).fillna(1).astype(int)
            df["Total Parcelas"] = parcelas[1].astype(float).fillna(1).astype(int)

            df["Valor"] = (
                df["Valor"].astype(str).str.replace(",", ".", regex=False).astype(float)
            )

    except Exception as e:
        logging.error(f"Erro ao limpar dados ERP: {e}", exc_info=True)
        raise

    return df

# =========================
# Função de limpeza Cielo
# =========================
def limpar_cielo(df):
    try:
        with st.spinner("🧹 Limpando dados da Cielo..."):
            df = df.iloc[8:].reset_index(drop=True)
            df.columns = df.iloc[0]
            df = df[1:].reset_index(drop=True)
            df.columns = df.columns.str.strip().str.lower()

            df = df.rename(columns={
                "valor bruto": "VALOR DA PARCELA",
                "valor líquido": "VALOR LÍQUIDO",
                "número da parcela": "PARCELA",
                "quantidade total de parcelas": "TOTAL_PARCELAS",
                "código da autorização": "AUTORIZAÇÃO",
                "nsu/doc": "NSU/DOC",
                "data da venda": "DATA DA VENDA",
                "data prevista de pagamento": "DATA DE VENCIMENTO",
                "tipo de lançamento": "TIPO DE LANÇAMENTO",
            })

            for col in ["VALOR DA PARCELA", "VALOR LÍQUIDO"]:
                df[col] = (
                    df[col].astype(str).str.replace(",", ".", regex=False).astype(float)
                )

            df["PARCELA"] = pd.to_numeric(df["PARCELA"], errors="coerce").fillna(1).astype(int)
            df["TOTAL_PARCELAS"] = pd.to_numeric(df["TOTAL_PARCELAS"], errors="coerce").fillna(1).astype(int)

            for col in ["DATA DA VENDA", "DATA DE VENCIMENTO"]:
                df[col] = pd.to_datetime(df[col], dayfirst=True, errors="coerce")
                
                
            # Mantém apenas as colunas mencionadas acima:
            colunas_manter = [
                "VALOR DA PARCELA",
                "VALOR LÍQUIDO",
                "PARCELA",
                "TOTAL_PARCELAS",
                "AUTORIZAÇÃO",
                "NSU/DOC",
                "DATA DA VENDA",
                "DATA DE VENCIMENTO",
                "TIPO DE LANÇAMENTO",
            ]
            df = df[colunas_manter]
    except Exception as e:
        logging.error(f"Erro ao limpar dados Cielo: {e}", exc_info=True)
        raise
    return df





# =========================
# ==Função de conciliação==
# =========================

def conciliar_cielo_erp(df_cielo, df_erp, tolerancia_dias=5, tolerancia_valor=0.20):
    df_cielo = df_cielo.copy()
    df_erp = df_erp.copy()

    # Normalizar chaves
    df_erp["Chave"] = pd.to_numeric(df_erp["Chave"], errors="coerce").astype("string")
    df_erp["Usada"] = False

    # Adiciona colunas de resultado na df_cielo
    df_cielo["Autorização ERP"] = None
    df_cielo["NSU ERP"] = None
    df_cielo["Chave ERP"] = None
    df_cielo["Valor ERP"] = None
    df_cielo["Emissão ERP"] = None
    df_cielo["Parcela ERP"] = None
    df_cielo["Total Parcelas ERP"] = None
    df_cielo["Pessoa do Título"] = None
    df_cielo["Pessoa do Título"] = None
    df_cielo["Status"] = "Não conciliado"
    df_cielo["Pontuação"] = 999

    progress_text = st.empty()  # cria um espaço que podemos atualizar
    progress_bar = st.progress(0)
    total = len(df_cielo)

    for i, row in df_cielo.iterrows():
        progresso = (i + 1) / total

        # Atualiza o texto com os registros já processados
        progress_text.text(f"🔄 Conciliando ({i + 1}/{total}) registros...")

        # Atualiza a barra
        progress_bar.progress(progresso)

        # seu processamento da linha aqui

    for i, row in df_cielo.iterrows():
        if pd.isna(row["AUTORIZAÇÃO"]) or pd.isna(row["NSU/DOC"]):
            logging.warning(f"⚠️ Linha {i} ignorada por dados ausentes.")
            continue

        logging.debug(f"🔍 Linha {i} - Aut: {row['AUTORIZAÇÃO']}, NSU: {row['NSU/DOC']}, Parcela: {row['PARCELA']}")

        candidatos = df_erp[
            (~df_erp["Usada"]) &
            (abs((df_erp["Emissão"] - row["DATA DA VENDA"]).dt.days) <= tolerancia_dias) &
            (abs(df_erp["Valor"] - row["VALOR DA PARCELA"]) <= tolerancia_valor) &
            (df_erp["Numero da Parcela"] == row["PARCELA"]) &
            (df_erp["Total Parcelas"] == row["TOTAL_PARCELAS"])
        ]

        logging.debug(f"🔎 {len(candidatos)} candidatos encontrados para a linha {i} da Cielo.")

        melhor = None
        menor_pontuacao = float("inf")

        for _, linha in candidatos.iterrows():
            dias_dif = abs((linha["Emissão"] - row["DATA DA VENDA"]).days)
            valor_dif = abs(linha["Valor"] - row["VALOR DA PARCELA"])
            sim_aut = fuzz.ratio(str(linha["Autorização"]), str(row["AUTORIZAÇÃO"]))
            sim_nsu = fuzz.ratio(str(linha["NSU"]), str(row["NSU/DOC"]))

            pontuacao = dias_dif * 10 + valor_dif * 100 + (100 - sim_aut) + (100 - sim_nsu)
            if "Pessoa do Título" in linha and linha["Pessoa do Título"] != "Cielo":
                    pontuacao += 101

            logging.debug(f"➡️ Testando Chave {linha['Chave']} | Dias: {dias_dif}, Valor: {valor_dif}, Aut: {sim_aut}, NSU: {sim_nsu}, Pontuação: {pontuacao:.2f}")

            if pontuacao < menor_pontuacao:
                menor_pontuacao = pontuacao
                melhor = linha

        if melhor is not None:
            idx_erp = df_erp.index[df_erp["Chave"] == melhor["Chave"]].tolist()
            if idx_erp:
                df_erp.at[idx_erp[0], "Usada"] = True

            df_cielo.at[i, "Autorização ERP"] = melhor["Autorização"]
            df_cielo.at[i, "NSU ERP"] = melhor["NSU"]
            df_cielo.at[i, "Chave ERP"] = melhor["Chave"]
            df_cielo.at[i, "Valor ERP"] = melhor["Valor"]
            df_cielo.at[i, "Emissão ERP"] = melhor["Emissão"]
            df_cielo.at[i, "Parcela ERP"] = melhor["Numero da Parcela"]
            df_cielo.at[i, "Total Parcelas ERP"] = melhor["Total Parcelas"]
            df_cielo.at[i, "Pessoa do Título"] = melhor.get("Pessoa do Título", None)
            df_cielo.at[i, "Status"] = "Conciliado"
            df_cielo.at[i, "Pontuação"] = round(menor_pontuacao, 0)
            logging.info(f"✅ Linha {i} conciliada com chave {melhor['Chave']} (Pontuação: {round(menor_pontuacao, 0)})")
        else:
            logging.info(f"❌ Linha {i} não conciliada (sem candidatos adequados)")

    return df_cielo, df_erp 






def main():

    # === BARRA LATERAL ===
    with st.sidebar:
        st.markdown("# App Conciliação Bancária")
        st.markdown("### Carregar planilhas")
        caminho_erp = st.file_uploader("ERP (CSV)", type=["csv"], key="erp_uploader")
        caminho_cielo = st.file_uploader("Cielo (XLSX)", type=["xlsx"], key="cielo_uploader")

    # === TELA INICIAL ===
    if caminho_erp is None or caminho_cielo is None:
        st.subheader("Bem-vindo ao Sistema de Conciliação")
        st.markdown("""
        <div style='text-align: center; margin-bottom: 20px;'>
            <p>Este sistema realiza a conciliação automática entre:</p>
            <p>•  Cielo</p>
            <p>• ERP</p>
        </div>
        """, unsafe_allow_html=True)
        st.warning("⚠️ Por favor, faça upload de ambos os arquivos para iniciar a conciliação")
        st.stop()

    def carregar_planilha(caminho):
        if caminho.name.lower().endswith(".csv"):
            return pd.read_csv(caminho, sep=";", encoding="latin1")
        elif caminho.name.lower().endswith(".xlsx") or caminho.name.lower().endswith(".xls"):
            return pd.read_excel(caminho, engine="openpyxl")
        else:
            raise ValueError("❌ Formato de arquivo não suportado. Só aceitamos CSV e XLSX.")

    try:
        with st.spinner("📂 Carregando planilhas..."):
            df_erp = carregar_planilha(caminho_erp)
            df_cielo = carregar_planilha(caminho_cielo)

        with st.spinner("🔧 Iniciando limpeza e conciliação dos dados..."):
            df_erp = limpar_erp(df_erp)
            df_cielo = limpar_cielo(df_cielo)
            df_conciliado, df_erp = conciliar_cielo_erp(df_cielo, df_erp)
            df_aba_conciliados = df_conciliado[df_conciliado["Status"] == "Conciliado"].copy()
            df_aba_nao_conciliados = df_conciliado[df_conciliado["Status"] != "Conciliado"].copy()

        totais_conc = {
            "liquido": df_aba_conciliados["VALOR LÍQUIDO"].sum(),
            "parcela": df_aba_conciliados["VALOR DA PARCELA"].sum(),
            "qtd": len(df_aba_conciliados)
        }
        totais_nao = {
            "liquido": df_aba_nao_conciliados["VALOR LÍQUIDO"].sum(),
            "parcela": df_aba_nao_conciliados["VALOR DA PARCELA"].sum(),
            "qtd": len(df_aba_nao_conciliados)
        }

        relatorio_linhas = [
            ["RELATÓRIO DE CONCILIAÇÃO", "", ""],
            ["CONCILIADO", "", ""],
            ["- Valor Líquido Total", "", f"R$ {totais_conc['liquido']:,.2f}"],
            ["- Valor da Parcela Total", "", f"R$ {totais_conc['parcela']:,.2f}"],
            ["- Quantidade de Títulos", "", f"{totais_conc['qtd']}"],
            ["", "", ""],
            ["NÃO CONCILIADO", "", ""],
            ["- Valor Líquido Total", "", f"R$ {totais_nao['liquido']:,.2f}"],
            ["- Valor da Parcela Total", "", f"R$ {totais_nao['parcela']:,.2f}"],
            ["- Quantidade de Títulos", "", f"{totais_nao['qtd']}"]
        ]
        relatorio_df = pd.DataFrame(relatorio_linhas, columns=["Categoria", "Descrição", "Valor"])

        # =====================================================================
        # EXCLUSÃO FINAL DAS COLUNAS (APÓS TODO O PROCESSAMENTO)
        # =====================================================================
        # Definir colunas a serem excluídas pelos nomes reais
        colunas_para_excluir = [
            "TIPO DE LANÇAMENTO",   # Coluna I
            "Parcela ERP",          # Coluna O
            "Total Parcelas ERP"    # Coluna P
        ]

        # Aplicar exclusão apenas se as colunas existirem
        for col in colunas_para_excluir:
            if col in df_aba_conciliados.columns:
                df_aba_conciliados = df_aba_conciliados.drop(columns=[col])
            if col in df_aba_nao_conciliados.columns:
                df_aba_nao_conciliados = df_aba_nao_conciliados.drop(columns=[col])

        # Agora gerar o Excel com as colunas já excluídas


        
        output_path = "Conciliação_final.xlsx"
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df_aba_conciliados.to_excel(writer, sheet_name="Conciliados", index=False)
            df_aba_nao_conciliados.to_excel(writer, sheet_name="Não conciliados", index=False)
            relatorio_df.to_excel(writer, sheet_name="Resumo", index=False)

            # Tratar abas especiais (aluguel e estornos) - também remover coluna I
            if "TIPO DE LANÇAMENTO" in df_cielo.columns:
                # Criar cópias para não alterar o original
                df_cielo_sem_coluna = df_cielo.drop(columns=["TIPO DE LANÇAMENTO"], errors="ignore")
                
                df_aluguel = df_cielo_sem_coluna[df_cielo["TIPO DE LANÇAMENTO"].str.lower().str.contains("aluguel", na=False)]
                if not df_aluguel.empty:
                    df_aluguel.to_excel(writer, sheet_name="Aluguel de máquina", index=False)
                
                df_estornos = df_cielo_sem_coluna[df_cielo["TIPO DE LANÇAMENTO"].str.lower().str.contains("estorno", na=False)]
                if not df_estornos.empty:
                    df_estornos.to_excel(writer, sheet_name="Estornos", index=False)

        # === INSERIR CHAVES ERP EM BLOCOS NA ABA RESUMO ===
        try:
            wb = load_workbook(output_path)
            ws_conciliados = wb["Conciliados"]
            ws_resumo = wb["Resumo"]

            # Detecta a coluna da Chave ERP
            header = [cell.value for cell in ws_conciliados[1]]
            if "Chave ERP" in header:
                idx_chave = header.index("Chave ERP")
                letra_coluna = chr(65 + idx_chave)

                chaves = [str(cell.value) for cell in ws_conciliados[letra_coluna][1:] if cell.value is not None]

                blocos = [chaves[i:i+2000] for i in range(0, len(chaves), 2000)]
                blocos_concat = [", ".join(bloco) for bloco in blocos]

                start_row = ws_resumo.max_row + 2
                for i, texto in enumerate(blocos_concat, start=1):
                    ws_resumo.cell(row=start_row + i - 1, column=1, value=f"Grupo {i}")
                    ws_resumo.cell(row=start_row + i - 1, column=2, value=texto)

                wb.save(output_path)
            else:
                st.warning("Coluna 'Chave ERP' não encontrada na aba Conciliados")

        except Exception as e:
            st.error(f"❌ Erro ao adicionar blocos de Chave ERP: {e}")

        # === INTERFACE FINAL ===
        with st.container():
            st.header("Resultados da Conciliação")
            col1, col2 = st.columns(2)
            with col1:
                st.metric("✅ Conciliados", 
                        f"R$ {totais_conc['liquido']:,.2f}", 
                        f"{totais_conc['qtd']} títulos")
            with col2:
                st.metric("⚠ Não Conciliados", 
                        f"R$ {totais_nao['liquido']:,.2f}", 
                        f"{totais_nao['qtd']} títulos")

            with st.expander("📊 Ver relatório completo"):
                st.dataframe(relatorio_df, hide_index=True)

        if os.path.exists(output_path):
            with open(output_path, "rb") as f:
                st.download_button(
                    label="📥 Baixar Planilha de Conciliação",
                    data=f,
                    file_name="Conciliação_final_cielo.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"❌ Erro ao carregar arquivos: {e}")
        st.stop()