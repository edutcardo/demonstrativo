import streamlit as st
import pandas as pd
import pdfplumber
import re
import os
from io import BytesIO

# --- LÓGICA DE EXTRAÇÃO DO PDF (VERSÃO CORRIGIDA E OTIMIZADA) ---
def extrair_dados_demonstrativo(arquivo_pdf):
    """
    Função otimizada para ler um PDF, extrair dados com mais flexibilidade na identificação
    das linhas da tabela, e controlar páginas não processadas.

    Alterações principais:
    1.  Melhoria na extração da tabela com `table_settings`.
    2.  Pré-processamento para separar linhas que foram mescladas incorretamente (problema da pág. 6).
    3.  Uso de uma função `safe_get` para evitar erros de `IndexError` ao acessar colunas,
        tornando o código mais robusto a variações no layout da tabela.
    """
    todos_os_registros = []
    paginas_nao_processadas = []

    # Função auxiliar para obter valores da lista de forma segura
    def safe_get(data_list, index, default=None):
        return data_list[index] if index < len(data_list) else default

    try:
        with pdfplumber.open(arquivo_pdf) as pdf:
            for i, pagina in enumerate(pdf.pages):
                num_pagina = i + 1
                try:
                    texto_pagina = pagina.extract_text(x_tolerance=2) or ""

                    # --- 1. Extração de dados estáticos (cabeçalho da página) ---
                    dados_pagina = {
                        "UC": "Não encontrado", "Nome": "Não encontrado", "Cidade": "Não encontrado",
                        "Tipo": "Não Identificado", "Custo de Disp. (kWh)": "N/A", "Página": num_pagina
                    }
                    uc_match = re.search(r"UC\s*:\s*(\d+)", texto_pagina)
                    if uc_match: dados_pagina["UC"] = uc_match.group(1)

                    nome_match = re.search(r"Nome\s*:\s*(.*?)(?:\n\s*Endereço|\n\s*Bairro|\n\s*1\.\s*Demonstrativos)", texto_pagina, re.DOTALL)
                    if nome_match:
                        nome_bruto = nome_match.group(1)
                    else: # Tenta uma regex mais simples como fallback
                        nome_match = re.search(r"Nome\s*:\s*(.*?)\n", texto_pagina)
                        nome_bruto = nome_match.group(1) if nome_match else ""

                    # Remove texto indesejado que pode ser capturado junto com o nome
                    custo_texto_match = re.search(r"Valor do Custo de Disp", nome_bruto)
                    if custo_texto_match:
                        nome_bruto = nome_bruto[:custo_texto_match.start()]
                    dados_pagina["Nome"] = nome_bruto.replace('\n', ' ').strip()
                    
                    cidade_match = re.search(r"Cidade\s*:\s*(.*?)\s*-", texto_pagina)
                    if cidade_match: dados_pagina["Cidade"] = cidade_match.group(1).strip()

                    if "UC Geradora" in texto_pagina: dados_pagina["Tipo"] = "Geradora"
                    elif "UC Beneficiária" in texto_pagina: dados_pagina["Tipo"] = "Beneficiária"
                    
                    custo_match = re.search(r"Valor do Custo de Disp\.\s*Kwh\s*[:\n,]*\s*\"?(\d+)", texto_pagina)
                    if custo_match: dados_pagina["Custo de Disp. (kWh)"] = custo_match.group(1)

                    # --- 2. Extração da tabela ---
                    # Melhorando a extração da tabela para lidar com layouts sem linhas verticais claras
                    tabela_bruta = pagina.extract_table(table_settings={"vertical_strategy": "text"})
                    
                    if not tabela_bruta or len(tabela_bruta) < 2:
                        st.write(f"-> Página {num_pagina}: Nenhuma tabela de dados encontrada.")
                        paginas_nao_processadas.append(num_pagina)
                        continue

                    # --- 3. Pré-processamento da tabela para separar linhas mescladas ---
                    tabela_dados = []
                    for linha_bruta in tabela_bruta:
                        if not linha_bruta or not linha_bruta[0] or not isinstance(linha_bruta[0], str):
                            tabela_dados.append(linha_bruta)
                            continue
                        
                        # Heurística para detectar células com múltiplas linhas de dados (comum na pág. 6)
                        if linha_bruta[0].count('/') > 1 and '\n' in linha_bruta[0]:
                            celulas_divididas = [str(c or '').split('\n') for c in linha_bruta]
                            num_sub_linhas = max(len(c) for c in celulas_divididas if c)
                            
                            for idx_sub_linha in range(num_sub_linhas):
                                nova_linha = [safe_get(col, idx_sub_linha, '') for col in celulas_divididas]
                                tabela_dados.append(nova_linha)
                        else:
                            tabela_dados.append(linha_bruta)

                    # --- 4. Processamento flexível das linhas da tabela ---
                    linhas_processadas_na_pagina = 0
                    for linha in tabela_dados:
                        if not linha: continue
                        
                        date_offset = -1
                        ref_mes = None

                        # Procura pela data nas 3 primeiras células da linha
                        for idx, cell in enumerate(linha[:3]):
                            if cell:
                                match = re.search(r"(\d{2}/\d{4})", str(cell))
                                if match:
                                    date_offset = idx
                                    ref_mes = match.group(1)
                                    break
                        
                        if date_offset != -1:
                            dados_mes = dados_pagina.copy()
                            dados_mes['Referência'] = ref_mes
                            
                            def clean_cell(cell_value):
                                if cell_value is None: return "0"
                                return str(cell_value).replace('.', '').replace('\n', ' ').strip() or "0"

                            # A lógica de extração agora é mais segura contra IndexErrors
                            # Mantendo a estrutura original (simples vs complexa) mas com acesso seguro
                            try:
                                if len(linha) > 15: # Tabela complexa
                                    dados_mes['Saldo Anterior (kWh)'] = clean_cell(safe_get(linha, 3))
                                    dados_mes['Créd. Receb. Outra UC (kWh)'] = clean_cell(safe_get(linha, 6))
                                    dados_mes['Energia Injetada (kWh)'] = clean_cell(safe_get(linha, 9))
                                    dados_mes['Energia Ativa (kWh)'] = clean_cell(safe_get(linha, 12))
                                    dados_mes['Crédito Utilizado (kWh)'] = clean_cell(safe_get(linha, 15))
                                    dados_mes['Saldo Mês (kWh)'] = clean_cell(safe_get(linha, 18))
                                    dados_mes['Saldo Transferido (kWh)'] = clean_cell(safe_get(linha, 21))
                                    dados_mes['Saldo Final (kWh)'] = clean_cell(safe_get(linha, 24))
                                else: # Tabela simples
                                    dados_mes['Saldo Anterior (kWh)'] = clean_cell(safe_get(linha, date_offset + 1))
                                    dados_mes['Créd. Receb. Outra UC (kWh)'] = clean_cell(safe_get(linha, date_offset + 2))
                                    dados_mes['Energia Injetada (kWh)'] = clean_cell(safe_get(linha, date_offset + 3))
                                    dados_mes['Energia Ativa (kWh)'] = clean_cell(safe_get(linha, date_offset + 4))
                                    dados_mes['Crédito Utilizado (kWh)'] = clean_cell(safe_get(linha, date_offset + 5))
                                    dados_mes['Saldo Mês (kWh)'] = clean_cell(safe_get(linha, date_offset + 6))
                                    dados_mes['Saldo Transferido (kWh)'] = clean_cell(safe_get(linha, date_offset + 7))
                                    dados_mes['Saldo Final (kWh)'] = clean_cell(safe_get(linha, date_offset + 8))
                                
                                todos_os_registros.append(dados_mes)
                                linhas_processadas_na_pagina += 1
                            except TypeError: # Pode ocorrer se alguma célula tiver um tipo inesperado
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
            return pd.DataFrame(), sorted(list(set(paginas_nao_processadas)))

        df = pd.DataFrame(todos_os_registros)
        ordem_colunas = [
            'Página', 'UC', 'Nome', 'Cidade', 'Tipo', 'Custo de Disp. (kWh)', 'Referência',
            'Saldo Anterior (kWh)', 'Créd. Receb. Outra UC (kWh)', 'Energia Injetada (kWh)',
            'Energia Ativa (kWh)', 'Crédito Utilizado (kWh)', 'Saldo Mês (kWh)',
            'Saldo Transferido (kWh)', 'Saldo Final (kWh)'
        ]
        df = df.reindex(columns=ordem_colunas)
        return df, sorted(list(set(paginas_nao_processadas)))

    except Exception as e:
        st.error(f"Ocorreu um erro crítico ao processar o PDF: {e}")
        return pd.DataFrame(), []

# --- INTERFACE DA APLICAÇÃO WEB (STREAMLIT) ---

st.set_page_config(page_title="Leitor de Demonstrativos", layout="centered")
st.title("Leitor e Conversor de Demonstrativos de Energia")
st.write("Faça o upload de um arquivo PDF para extrair os dados e gerar uma planilha Excel.")

st.markdown("""
<style>
    div.stButton > button:first-child {
        background-color: #28a745; color: white; border: none;
        border-radius: 5px; padding: 10px 24px; font-size: 16px;
    }
    div.stButton > button:first-child:hover {
        background-color: #218838; color: white;
    }
</style>""", unsafe_allow_html=True)

arquivo_pdf_anexado = st.file_uploader(
    "Anexe o arquivo PDF aqui", type="pdf", help="Apenas arquivos PDF são aceitos"
)

if arquivo_pdf_anexado is not None:
    st.info(f"Arquivo anexado: **{arquivo_pdf_anexado.name}**")
    
    if st.button("Ler Demonstrativo e Gerar Planilha"):
        with st.spinner("Processando o PDF... Por favor, aguarde."):
            df_resultado, paginas_com_problema = extrair_dados_demonstrativo(arquivo_pdf_anexado)

            if df_resultado is not None and not df_resultado.empty:
                st.success("PDF processado com sucesso!")
                if paginas_com_problema:
                    paginas_str = ", ".join(map(str, sorted(paginas_com_problema)))
                    st.warning(f"**Atenção:** Não foi possível extrair tabelas das seguintes páginas: **{paginas_str}**. Verifique se elas contêm dados no formato esperado.")

                st.write("### Pré-visualização dos Dados")
                st.dataframe(df_resultado)

                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_resultado.to_excel(writer, index=False, sheet_name='Demonstrativo')
                
                nome_base = os.path.splitext(arquivo_pdf_anexado.name)[0]
                nome_arquivo_excel = f'{nome_base}_extraido.xlsx'

                st.download_button(
                    label="📥 Baixar Planilha Excel Completa",
                    data=output.getvalue(),
                    file_name=nome_arquivo_excel,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("Não foi possível extrair nenhum dado válido do PDF. Verifique se o formato do arquivo é o esperado ou se ele não está vazio.")
