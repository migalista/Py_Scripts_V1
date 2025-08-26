# NOME DO ARQUIVO: conversor_excel.py

import pandas as pd
import os

def run():
    """Função principal que executa o fluxo do conversor."""
    print("=" * 50)
    print("--- Conversor Rápido de Excel para CSV ---")
    print("=" * 50)

    # --- PASSO 1: ESCOLHER ARQUIVO OU PASTA ---
    while True:
        modo = input("""O que você quer converter?
  [1] Um único arquivo Excel
  [2] Todos os arquivos Excel de uma pasta
Escolha (1 ou 2): """)
        if modo in ['1', '2']:
            break
        print("\n*** Opção inválida! Digite 1 ou 2. ***\n")

    # --- PASSO 2: INDICAR O LOCAL DOS ARQUIVOS ---
    lista_arquivos = obter_lista_de_arquivos(modo)
    if not lista_arquivos:
        print("\nNenhum arquivo Excel encontrado. O programa será encerrado.")
        return

    # --- PASSO 3: DEFINIR ONDE SALVAR OS ARQUIVOS CSV ---
    caminho_saida_base = obter_caminho_saida()

    # --- PASSO 4 (OPCIONAL): ORGANIZAR EM SUBPASTAS ---
    usar_subpastas = False
    if len(lista_arquivos) > 1: # Só faz sentido perguntar se tiver mais de um arquivo
        resposta = input("\nCriar uma subpasta para cada arquivo Excel? (s/n): ").lower()
        usar_subpastas = (resposta == 's')

    # --- EXECUÇÃO DA CONVERSÃO ---
    processar_arquivos(lista_arquivos, caminho_saida_base, usar_subpastas)

def obter_lista_de_arquivos(modo):
    """Pede o caminho e retorna a lista de arquivos Excel a serem processados."""
    while True:
        if modo == '1':
            caminho = input("\nDigite o nome do arquivo Excel (ex: Relatorio.xlsx): ")
            if os.path.isfile(caminho):
                return [caminho]
            print(f"\n*** ERRO: Arquivo '{caminho}' não encontrado. ***")
        else: # modo == '2'
            caminho = input("\nDigite o caminho da pasta (ex: C:\\Users\\Joao\\Documentos): ")
            if os.path.isdir(caminho):
                arquivos = [os.path.join(caminho, f) for f in os.listdir(caminho) if f.lower().endswith(('.xlsx', '.xls'))]
                if not arquivos:
                    print(f"\n*** AVISO: Nenhum arquivo .xlsx ou .xls encontrado na pasta. ***")
                return arquivos
            print(f"\n*** ERRO: Pasta '{caminho}' não encontrada. ***")

def obter_caminho_saida():
    """Define a pasta de destino para os arquivos CSV."""
    while True:
        print("\nOnde salvar os arquivos CSV?")
        print("  [1] Na pasta atual")
        print("  [2] Em outra pasta (criar ou escolher uma existente)")
        escolha = input("Escolha (1 ou 2): ")

        if escolha == '1':
            return os.getcwd()
        if escolha == '2':
            caminho_destino = input("Digite o nome da pasta de destino (se não existir, será criada): ")
            try:
                os.makedirs(caminho_destino, exist_ok=True)
                return caminho_destino
            except OSError as e:
                print(f"\n*** ERRO: Nome de pasta inválido. {e} ***")
        print("\n*** Opção inválida! Digite 1 ou 2. ***")

def processar_arquivos(lista_arquivos, caminho_saida_base, usar_subpastas):
    """Realiza a conversão dos arquivos, mostrando o progresso."""
    print("\n" + "="*50)
    print("--- INICIANDO CONVERSÃO ---")
    sucessos = 0
    falhas = 0
    
    for caminho_excel in lista_arquivos:
        nome_base_arquivo = os.path.splitext(os.path.basename(caminho_excel))[0]
        print(f"\n[*] Lendo arquivo: '{nome_base_arquivo}'...")
        
        try:
            caminho_final = caminho_saida_base
            if usar_subpastas:
                caminho_final = os.path.join(caminho_saida_base, nome_base_arquivo)
                os.makedirs(caminho_final, exist_ok=True)

            xls = pd.ExcelFile(caminho_excel)
            for nome_aba in xls.sheet_names:
                df = xls.parse(nome_aba)
                nome_csv = f"{nome_aba}.csv"
                caminho_csv_completo = os.path.join(caminho_final, nome_csv)
                df.to_csv(caminho_csv_completo, index=False, encoding='utf-8-sig')
                print(f"  -> Aba '{nome_aba}' salva como '{nome_csv}'")
            sucessos += 1
        except Exception as e:
            print(f"  *** FALHA ao processar '{nome_base_arquivo}': {e} ***")
            falhas += 1

    print("\n" + "="*50)
    print("--- CONVERSÃO CONCLUÍDA ---")
    print(f"Arquivos processados com sucesso: {sucessos}")
    print(f"Arquivos com falha: {falhas}")
    print(f"Os resultados estão em: '{os.path.abspath(caminho_saida_base)}'")
    print("="*50)

# --- INICIA O SCRIPT ---
if __name__ == "__main__":
    run()
    # Mantém a janela aberta no final para o usuário ver o resultado
    input("\nPressione ENTER para fechar.")