import os
import pandas as pd


caminho_arquivo = 'Inventário 2024.xlsx'
arquivo_saida = 'lista_hostnames_ativos.txt'

def extrair_hostnames_procurando_texto(caminho_xlsx, arquivo_txt):
    # Verifica se o arquivo existe
    if not os.path.exists(caminho_xlsx):
        print(f"Erro: O arquivo '{caminho_xlsx}' não foi encontrado.")
        return

    print(f"Arquivo encontrado: {caminho_xlsx}")
    
    try:
        # Carrega todas as abas da planilha
        planilhas = pd.read_excel(caminho_xlsx, sheet_name=None, header=None)
        print(f"{len(planilhas)} abas encontradas na planilha.")

        # Conjunto (set) para armazenar hostnames/equipamentos únicos
        hostnames_unicos = set()

        # Percorre cada aba da planilha
        for nome_aba, df in planilhas.items():
            # Verifica se a aba deve ser ignorada
            if nome_aba.strip().lower() == 'estoque':
                print(f"Aba '{nome_aba}' ignorada conforme solicitado.")
                continue  # Pula a aba "ESTOQUE"

            print(f"Processando aba: '{nome_aba}'")
            # Percorre todas as células procurando pelas palavras "Hostname" ou "Equipamentos"
            for linha in range(df.shape[0]):
                for coluna in range(df.shape[1]):
                    # Normaliza o texto da célula para comparação
                    celula = str(df.iloc[linha, coluna]).strip().lower()
                    if celula in ['hostname', 'equipamentos']:
                        print(f"'{celula.capitalize()}' encontrado na aba '{nome_aba}', linha {linha + 1}, coluna {coluna + 1}")
                        
                        # Coleta os valores abaixo da célula até encontrar uma célula vazia
                        linha_atual = linha + 1
                        while linha_atual < df.shape[0]:
                            valor = df.iloc[linha_atual, coluna]
                            if pd.isna(valor):  # Para se encontrar uma célula vazia
                                break
                            # Ignora cabeçalhos repetidos e adiciona ao conjunto
                            if str(valor).strip().lower() not in ['hostname', 'equipamentos']:
                                hostnames_unicos.add(str(valor).strip())
                            linha_atual += 1
        
        # Salva os valores únicos encontrados em um arquivo TXT
        if hostnames_unicos:
            with open(arquivo_txt, 'w') as arquivo:
                for hostname in sorted(hostnames_unicos):  # Ordena os nomes antes de salvar
                    arquivo.write(f"{hostname}\n")
            print(f"Extração concluída! {len(hostnames_unicos)} entradas únicas encontradas e salvas em '{arquivo_txt}'.")
        else:
            print("Nenhum dado foi encontrado na planilha.")
    
    except Exception as e:
        print(f"Erro ao processar o arquivo: {e}")

# Executa a função
extrair_hostnames_procurando_texto(caminho_arquivo, arquivo_saida)