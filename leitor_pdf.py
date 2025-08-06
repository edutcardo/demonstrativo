import pdfplumber
import pandas as pd
import re
import os

def extrair_dados_demonstrativo(caminho_pdf):
    """
    Função para ler um PDF de demonstrativo de energia, extrair e organizar os dados.

    Args:
        caminho_pdf (str): O caminho para o arquivo PDF.

    Returns:
        pandas.DataFrame: Um DataFrame com os dados extraídos de cada unidade consumidora.
    """
    if not os.path.exists(caminho_pdf):
        print(f"Erro: O arquivo '{caminho_pdf}' não foi encontrado.")
        print("Por favor, verifique se o caminho da pasta e o nome do arquivo estão corretos.")
        return None

    unidades_consumidoras = []

    print(f"Lendo o arquivo: {caminho_pdf}")

    with pdfplumber.open(caminho_pdf) as pdf:
        # Itera por cada página do documento
        for i, pagina in enumerate(pdf.pages):
            print(f"Processando página {i+1}...")
            
            texto_pagina = pagina.extract_text()
            if not texto_pagina:
                print(f"  - Aviso: Não foi possível extrair texto da página {i+1}.")
                continue

            # Dicionário para armazenar os dados da página atual
            dados_uc = {
                "UC": "Não encontrado",
                "Nome": "Não encontrado",
                "Cidade": "Não encontrado",
                "Tipo": "Não encontrado",
                "Ref. Mês": "03/2025",
                "Crédito Recebido (kWh)": "N/A",
                "Energia Injetada (kWh)": "N/A",
                "Crédito Utilizado (kWh)": "N/A",
                "Saldo Final (kWh)": "N/A",
                "Página": i + 1
            }

            # Extrai informações básicas com expressões regulares (Regex)
            uc_match = re.search(r"UC\s*:\s*(\d+)", texto_pagina)
            if uc_match:
                dados_uc["UC"] = uc_match.group(1)

            nome_match = re.search(r"Nome\s*:\s*(.*?)\n", texto_pagina)
            if nome_match:
                dados_uc["Nome"] = nome_match.group(1).strip()

            cidade_match = re.search(r"Cidade\s*:\s*(.*?)\s*-", texto_pagina)
            if cidade_match:
                dados_uc["Cidade"] = cidade_match.group(1).strip()
            
            if "UC Geradora" in texto_pagina:
                dados_uc["Tipo"] = "Geradora"
            elif "UC Beneficiária" in texto_pagina:
                dados_uc["Tipo"] = "Beneficiária"

            # Extrai a tabela principal da página
            tabelas = pagina.extract_tables()
            tabela_dados = None
            for tabela in tabelas:
                if tabela and tabela[0] and tabela[0][0] and 'Referência' in tabela[0][0]:
                    tabela_dados = tabela
                    break
            
            if not tabela_dados:
                print(f"  - Aviso: Nenhuma tabela de dados encontrada na página {i+1}.")
                unidades_consumidoras.append(dados_uc)
                continue

            # Encontra a linha de dados de 03/2025
            cabecalho = tabela_dados[1]
            linha_dados = None
            for linha in tabela_dados:
                if linha and linha[0] is not None and '03/2025' in linha[0]:
                    linha_dados = linha
                    break
            
            if not linha_dados:
                print(f"  - Aviso: Dados para 03/2025 não encontrados na tabela da página {i+1}.")
                unidades_consumidoras.append(dados_uc)
                continue
                
            # Mapeia os dados da linha para as colunas corretas
            if 'TP' in cabecalho:
                try:
                    idx_cred_recebido = 2
                    idx_energia_injetada = 3
                    idx_cred_utilizado = 5
                    idx_saldo_final = 8
                    
                    if len(linha_dados) < 9:
                       idx_cred_recebido = cabecalho.index('TP', 1)
                       idx_energia_injetada = cabecalho.index('TP', idx_cred_recebido + 1)
                       idx_cred_utilizado = cabecalho.index('TP', idx_energia_injetada + 1)
                       idx_saldo_final = cabecalho.index('TP', idx_cred_utilizado + 3)

                    dados_uc["Crédito Recebido (kWh)"] = linha_dados[idx_cred_recebido].replace('.', '') if linha_dados[idx_cred_recebido] else "0"
                    dados_uc["Energia Injetada (kWh)"] = linha_dados[idx_energia_injetada].replace('.', '') if linha_dados[idx_energia_injetada] else "0"
                    dados_uc["Crédito Utilizado (kWh)"] = linha_dados[idx_cred_utilizado].replace('.', '') if linha_dados[idx_cred_utilizado] else "0"
                    dados_uc["Saldo Final (kWh)"] = linha_dados[idx_saldo_final].replace('.', '') if linha_dados[idx_saldo_final] else "0"
                except (ValueError, IndexError) as e:
                     print(f"  - Erro ao processar a tabela na página {i+1}: {e}")

            unidades_consumidoras.append(dados_uc)

    df = pd.DataFrame(unidades_consumidoras)
    return df

# --- INÍCIO DA EXECUÇÃO ---
if __name__ == "__main__":
    # --- CONFIGURAÇÃO DO ARQUIVO ALVO ---
    diretorio_alvo = r"C:\Users\CANAL VERDE\Documents\demonstrativo\arquivos"
    nome_do_arquivo_pdf = "DemonstrativoMicroMiniGeracao (53).pdf"
    caminho_completo_pdf = os.path.join(diretorio_alvo, nome_do_arquivo_pdf)
    
    # --- FIM DA CONFIGURAÇÃO ---

    # Chama a função principal com o caminho completo do arquivo.
    df_resultado = extrair_dados_demonstrativo(caminho_completo_pdf)

    # Exibe o resultado e salva em Excel se a extração foi bem-sucedida.
    if df_resultado is not None and not df_resultado.empty:
        print("\n--- RESUMO DOS DADOS EXTRAÍDOS ---")
        pd.set_option('display.max_rows', 500)
        pd.set_option('display.max_columns', 500)
        pd.set_option('display.width', 1000)
        print(df_resultado)

        # --- NOVA SEÇÃO PARA SALVAR EM EXCEL ---
        try:
            nome_arquivo_excel = "resumo_demonstrativo.xlsx"
            caminho_completo_excel = os.path.join(diretorio_alvo, nome_arquivo_excel)
            
            # Salva o DataFrame em um arquivo Excel, sem a coluna de índice (index=False)
            df_resultado.to_excel(caminho_completo_excel, index=False, engine='openpyxl')
            
            print("\n--- SUCESSO! ---")
            print(f"A planilha foi salva com sucesso em:")
            print(caminho_completo_excel)
        except Exception as e:
            print(f"\n--- ERRO AO SALVAR A PLANILHA ---")
            print(f"Ocorreu um erro: {e}")
            print("Verifique se você tem permissão para escrever na pasta de destino.")
            
    elif df_resultado is not None:
         print("\nO script foi executado, mas nenhum dado foi extraído para ser salvo na planilha.")