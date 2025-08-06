import streamlit as st
import pandas as pd
import pdfplumber
import re
import os
from io import BytesIO

# --- LÓGICA DE EXTRAÇÃO DO PDF (a mesma função de antes) ---
def extrair_dados_demonstrativo(arquivo_pdf):
    unidades_consumidoras = []

    try:
        with pdfplumber.open(arquivo_pdf) as pdf:
            for i, pagina in enumerate(pdf.pages):
                texto_pagina = pagina.extract_text()
                if not texto_pagina:
                    continue

                dados_uc = {
                    "UC": "Não encontrado", "Nome": "Não encontrado", "Cidade": "Não encontrado",
                    "Tipo": "Não encontrado", "Ref. Mês": "03/2025", "Crédito Recebido (kWh)": "N/A",
                    "Energia Injetada (kWh)": "N/A", "Crédito Utilizado (kWh)": "N/A",
                    "Saldo Final (kWh)": "N/A", "Página": i + 1
                }

                uc_match = re.search(r"UC\s*:\s*(\d+)", texto_pagina)
                if uc_match: dados_uc["UC"] = uc_match.group(1)

                nome_match = re.search(r"Nome\s*:\s*(.*?)\n", texto_pagina)
                if nome_match: dados_uc["Nome"] = nome_match.group(1).strip()

                cidade_match = re.search(r"Cidade\s*:\s*(.*?)\s*-", texto_pagina)
                if cidade_match: dados_uc["Cidade"] = cidade_match.group(1).strip()

                if "UC Geradora" in texto_pagina: dados_uc["Tipo"] = "Geradora"
                elif "UC Beneficiária" in texto_pagina: dados_uc["Tipo"] = "Beneficiária"

                tabelas = pagina.extract_tables()
                tabela_dados = None
                for tabela in tabelas:
                    if tabela and tabela[0] and tabela[0][0] and 'Referência' in tabela[0][0]:
                        tabela_dados = tabela
                        break
                
                if not tabela_dados:
                    unidades_consumidoras.append(dados_uc)
                    continue

                cabecalho = tabela_dados[1]
                linha_dados = None
                for linha in tabela_dados:
                    if linha and linha[0] is not None and '03/2025' in linha[0]:
                        linha_dados = linha
                        break
                
                if not linha_dados:
                    unidades_consumidoras.append(dados_uc)
                    continue
                
                if 'TP' in cabecalho:
                    try:
                        idx_cred_recebido, idx_energia_injetada, idx_cred_utilizado, idx_saldo_final = 2, 3, 5, 8
                        
                        dados_uc["Crédito Recebido (kWh)"] = linha_dados[idx_cred_recebido].replace('.', '') if linha_dados[idx_cred_recebido] else "0"
                        dados_uc["Energia Injetada (kWh)"] = linha_dados[idx_energia_injetada].replace('.', '') if linha_dados[idx_energia_injetada] else "0"
                        dados_uc["Crédito Utilizado (kWh)"] = linha_dados[idx_cred_utilizado].replace('.', '') if linha_dados[idx_cred_utilizado] else "0"
                        dados_uc["Saldo Final (kWh)"] = linha_dados[idx_saldo_final].replace('.', '') if linha_dados[idx_saldo_final] else "0"
                    except (ValueError, IndexError):
                        pass

                unidades_consumidoras.append(dados_uc)
        
        return pd.DataFrame(unidades_consumidoras)

    except Exception as e:
        st.error(f"Ocorreu um erro ao processar o PDF: {e}")
        return None

# --- INTERFACE DA APLICAÇÃO WEB (STREAMLIT) ---

# Título da Aplicação
st.set_page_config(page_title="Leitor de Demonstrativos", layout="centered")
st.title("Leitor e Conversor de Demonstrativos de Energia")
st.write("Faça o upload de um arquivo PDF para extrair os dados e gerar uma planilha Excel.")

# Componente de Upload de Arquivo
arquivo_pdf_anexado = st.file_uploader(
    "Anexe o arquivo PDF aqui",
    type="pdf",
    help="Apenas arquivos PDF são aceitos"
)

# Botão para iniciar o processamento
if arquivo_pdf_anexado is not None:
    st.info(f"Arquivo anexado: **{arquivo_pdf_anexado.name}**")
    
    if st.button("Ler Demonstrativo e Gerar Planilha", type="primary"):
        with st.spinner("Processando o PDF... Por favor, aguarde."):
            # Chama a função de extração
            df_resultado = extrair_dados_demonstrativo(arquivo_pdf_anexado)

            if df_resultado is not None and not df_resultado.empty:
                st.success("PDF processado com sucesso!")
                
                # Mostra uma prévia da tabela na tela
                st.write("### Pré-visualização dos Dados")
                st.dataframe(df_resultado)

                # --- Lógica para Download do Excel ---
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_resultado.to_excel(writer, index=False, sheet_name='Demonstrativo')
                
                # Pega o nome do arquivo original e troca a extensão para .xlsx
                nome_arquivo_excel = os.path.splitext(arquivo_pdf_anexado.name)[0] + '.xlsx'

                st.download_button(
                    label="📥 Baixar Planilha Excel",
                    data=output.getvalue(),
                    file_name=nome_arquivo_excel,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("Não foi possível extrair dados do PDF. Verifique o arquivo.")