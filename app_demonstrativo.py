import streamlit as st
import pandas as pd
import pdfplumber
import re
import os
from io import BytesIO

# --- LÓGICA DE EXTRAÇÃO DO PDF (VERSÃO APRIMORADA COM CONTROLE DE PÁGINA) ---
def extrair_dados_demonstrativo(arquivo_pdf):
    """
    Função aprimorada para ler um PDF de demonstrativo, extrair dados de todos os meses de cada página
    e organizar em um DataFrame detalhado, com controle de páginas não processadas.
    """
    todos_os_registros = []
    paginas_nao_processadas = [] # Lista para rastrear páginas que falharam

    try:
        # Abre o arquivo PDF com o pdfplumber
        with pdfplumber.open(arquivo_pdf) as pdf:
            # Itera sobre cada página do documento
            for i, pagina in enumerate(pdf.pages):
                num_pagina = i + 1 # Número real da página (começando em 1)
                try:
                    texto_pagina = pagina.extract_text(x_tolerance=2, layout=True)
                    if not texto_pagina:
                        st.write(f"-> Página {num_pagina}: Nenhum texto encontrado, pulando.")
                        paginas_nao_processadas.append(num_pagina)
                        continue

                    # --- 1. Extrai os dados estáticos que se aplicam a toda a página ---
                    dados_pagina = {
                        "UC": "Não encontrado", "Nome": "Não encontrado", "Cidade": "Não encontrado",
                        "Tipo": "Não Identificado", "Custo de Disp. (kWh)": "N/A", "Página": num_pagina
                    }

                    # Extrai UC (Unidade Consumidora)
                    uc_match = re.search(r"UC\s*:\s*(\d+)", texto_pagina)
                    if uc_match: dados_pagina["UC"] = uc_match.group(1)

                    # --- LÓGICA CORRIGIDA PARA EXTRAIR O NOME ---
                    nome_match = re.search(r"Nome\s*:\s*(.*?)(?:\n\s*Endereço|\n\s*Bairro|\n\s*1\.\s*Demonstrativos)", texto_pagina, re.DOTALL)
                    if nome_match:
                        nome_bruto = nome_match.group(1)
                    else:
                        nome_match = re.search(r"Nome\s*:\s*(.*?)\n", texto_pagina)
                        nome_bruto = nome_match.group(1) if nome_match else ""

                    custo_texto_match = re.search(r"Valor do Custo de Disp", nome_bruto)
                    if custo_texto_match:
                        nome_bruto = nome_bruto[:custo_texto_match.start()]
                    
                    dados_pagina["Nome"] = nome_bruto.replace('\n', ' ').strip()
                    
                    # Extrai Cidade
                    cidade_match = re.search(r"Cidade\s*:\s*(.*?)\s*-", texto_pagina)
                    if cidade_match: dados_pagina["Cidade"] = cidade_match.group(1).strip()

                    # Extrai Tipo
                    if "Demonstrativos de Créditos Utilizados - UC Geradora" in texto_pagina:
                        dados_pagina["Tipo"] = "Geradora"
                    elif "Demonstrativos de Créditos Utilizados - UC Beneficiária" in texto_pagina:
                        dados_pagina["Tipo"] = "Beneficiária"

                    # --- LÓGICA CORRIGIDA PARA EXTRAIR O CUSTO DE DISPONIBILIDADE ---
                    custo_match = re.search(r"Valor do Custo de Disp\.\s*Kwh\s*[:\n,]*\s*\"?(\d+)", texto_pagina)
                    if custo_match: dados_pagina["Custo de Disp. (kWh)"] = custo_match.group(1)

                    # --- 2. Extrai a tabela de dados mensais da página ---
                    tabela_dados = pagina.extract_table()
                    if not tabela_dados or len(tabela_dados) < 2:
                        st.write(f"-> Página {num_pagina}: Nenhuma tabela de dados encontrada, pulando.")
                        paginas_nao_processadas.append(num_pagina)
                        continue

                    # --- 3. Itera sobre as linhas da tabela para extrair os dados de CADA MÊS ---
                    linhas_processadas_na_pagina = 0
                    for linha in tabela_dados:
                        if not linha or not linha[0]: continue
                        
                        ref_mes = str(linha[0]).strip()
                        if re.match(r"^\d{2}/\d{4}$", ref_mes):
                            dados_mes = dados_pagina.copy()
                            dados_mes['Referência'] = ref_mes
                            
                            def clean_cell(cell_value):
                                if cell_value is None: return "0"
                                return str(cell_value).replace('.', '').replace('\n', ' ').strip() or "0"

                            try:
                                if len(linha) > 15: # Tabela complexa
                                    dados_mes['Saldo Anterior (kWh)'] = clean_cell(linha[3])
                                    dados_mes['Créd. Receb. Outra UC (kWh)'] = clean_cell(linha[6])
                                    dados_mes['Energia Injetada (kWh)'] = clean_cell(linha[9])
                                    dados_mes['Energia Ativa (kWh)'] = clean_cell(linha[12])
                                    dados_mes['Crédito Utilizado (kWh)'] = clean_cell(linha[15])
                                    dados_mes['Saldo Mês (kWh)'] = clean_cell(linha[18])
                                    dados_mes['Saldo Transferido (kWh)'] = clean_cell(linha[21])
                                    dados_mes['Saldo Final (kWh)'] = clean_cell(linha[24])
                                else: # Tabela simples
                                    dados_mes['Saldo Anterior (kWh)'] = clean_cell(linha[1])
                                    dados_mes['Créd. Receb. Outra UC (kWh)'] = clean_cell(linha[2])
                                    dados_mes['Energia Injetada (kWh)'] = clean_cell(linha[3])
                                    dados_mes['Energia Ativa (kWh)'] = clean_cell(linha[4])
                                    dados_mes['Crédito Utilizado (kWh)'] = clean_cell(linha[5])
                                    dados_mes['Saldo Mês (kWh)'] = clean_cell(linha[6])
                                    dados_mes['Saldo Transferido (kWh)'] = clean_cell(linha[7])
                                    dados_mes['Saldo Final (kWh)'] = clean_cell(linha[8])
                                
                                todos_os_registros.append(dados_mes)
                                linhas_processadas_na_pagina += 1
                            except (IndexError, TypeError):
                                continue
                    
                    if linhas_processadas_na_pagina == 0:
                         st.write(f"-> Página {num_pagina}: Tabela encontrada, mas sem linhas de dados válidas.")
                         if num_pagina not in paginas_nao_processadas:
                            paginas_nao_processadas.append(num_pagina)

                except Exception as page_error:
                    st.write(f"-> Página {num_pagina}: Ocorreu um erro inesperado: {page_error}")
                    if num_pagina not in paginas_nao_processadas:
                        paginas_nao_processadas.append(num_pagina)
                    continue

        if not todos_os_registros:
            return pd.DataFrame(), list(set(paginas_nao_processadas))

        df = pd.DataFrame(todos_os_registros)
        
        ordem_colunas = [
            'Página', 'UC', 'Nome', 'Cidade', 'Tipo', 'Custo de Disp. (kWh)', 'Referência',
            'Saldo Anterior (kWh)', 'Créd. Receb. Outra UC (kWh)', 'Energia Injetada (kWh)',
            'Energia Ativa (kWh)', 'Crédito Utilizado (kWh)', 'Saldo Mês (kWh)',
            'Saldo Transferido (kWh)', 'Saldo Final (kWh)'
        ]
        
        df = df.reindex(columns=ordem_colunas)
        return df, list(set(paginas_nao_processadas))

    except Exception as e:
        st.error(f"Ocorreu um erro crítico ao processar o PDF: {e}")
        return None, []

# --- INTERFACE DA APLICAÇÃO WEB (STREAMLIT) ---

st.set_page_config(page_title="Leitor de Demonstrativos", layout="centered")
st.title("Leitor e Conversor de Demonstrativos de Energia")
st.write("Faça o upload de um arquivo PDF para extrair os dados e gerar uma planilha Excel.")

st.markdown("""
<style>
    div.stButton > button:first-child {
        background-color: #28a745;
        color: white;
        border: none;
        border-radius: 5px;
        padding: 10px 24px;
        font-size: 16px;
    }
    div.stButton > button:first-child:hover {
        background-color: #218838;
        color: white;
    }
</style>""", unsafe_allow_html=True)

arquivo_pdf_anexado = st.file_uploader(
    "Anexe o arquivo PDF aqui",
    type="pdf",
    help="Apenas arquivos PDF são aceitos"
)

if arquivo_pdf_anexado is not None:
    st.info(f"Arquivo anexado: **{arquivo_pdf_anexado.name}**")
    
    if st.button("Ler Demonstrativo e Gerar Planilha"):
        with st.spinner("Processando o PDF... Por favor, aguarde."):
            df_resultado, paginas_com_problema = extrair_dados_demonstrativo(arquivo_pdf_anexado)

            if df_resultado is not None and not df_resultado.empty:
                st.success("PDF processado com sucesso!")

                # CONTROLE DE PÁGINAS: informa ao usuário se alguma página foi pulada
                if paginas_com_problema:
                    paginas_str = ", ".join(map(str, sorted(paginas_com_problema)))
                    st.warning(f"**Atenção:** Não foi possível extrair tabelas das seguintes páginas: **{paginas_str}**. Verifique se elas contêm dados no formato esperado.")

                st.write("### Pré-visualização dos Dados")
                st.dataframe(df_resultado)

                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_resultado.to_excel(writer, index=False, sheet_name='Demonstrativo')
                
                nome_arquivo_excel = os.path.splitext(arquivo_pdf_anexado.name)[0] + '_completo.xlsx'

                st.download_button(
                    label="📥 Baixar Planilha Excel Completa",
                    data=output.getvalue(),
                    file_name=nome_arquivo_excel,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("Não foi possível extrair dados do PDF. Verifique se o formato do arquivo é o esperado ou se ele não está vazio.")