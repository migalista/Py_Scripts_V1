# -*- coding: utf-8 -*-
import pandas as pd
import os

def processar_dados(pasta_trabalho):
    """Processa o arquivo bruto do LISTCUBE e gera um arquivo de análise com abas."""
    try:
        caminho_listcube = os.path.join(pasta_trabalho, "ExtracaoSAP", "LISTCUBE_Export.xlsx")
        print(f"[INFO] Lendo dados brutos de: {caminho_listcube}")
        df_listcube = pd.read_excel(caminho_listcube)

        print("[INFO] Criando tabela dinâmica principal (base para 'Draft')...")
        colunas_linhas = [
            'Ship to Region', 'S.Org(Company Code)', 'Ship to Country', 
            'YA_SBUT', 'L1 Demand Planner', 'YA_MATIDH', 'YA_SLDTO', 'Delivering Plant'
        ]
        colunas_valores = [
            'Uncleansed Sales History / Orders', 
            'History uncleansed in Base unit of measu'
        ]
        
        df_draft = pd.pivot_table(
            df_listcube,
            index=[col for col in colunas_linhas if col in df_listcube.columns],
            values=[col for col in colunas_valores if col in df_listcube.columns],
            aggfunc='sum',
            fill_value=0
        ).reset_index()

        print("[INFO] Simulando a lógica de #N/D para separar dados...")
        caminho_ref_apps = os.path.join(pasta_trabalho, "Referencias", "referencia_apps.xlsx")
        chave_de_busca = 'Ship to Country'
        
        try:
            df_ref_apps = pd.read_excel(caminho_ref_apps)
            df_merged = pd.merge(df_draft, df_ref_apps, on=chave_de_busca, how='left', indicator=True)
            df_ush = df_merged[df_merged['_merge'] == 'left_only'].drop(columns=['_merge'])
            print(f"[INFO] {len(df_ush)} linhas separadas para a aba 'USH'.")
            df_fca_lag = df_ush.copy()
            print(f"[INFO] {len(df_fca_lag)} linhas separadas para a aba 'FCA_LAG_1_STC'.")

        except FileNotFoundError:
            print(f"[AVISO] Arquivo de referência '{caminho_ref_apps}' não encontrado. Abas 'USH' e 'FCA_LAG_1_STC' ficarão vazias.")
            df_ush = pd.DataFrame(columns=df_draft.columns)
            df_fca_lag = pd.DataFrame(columns=df_draft.columns)

        caminho_saida = os.path.join(pasta_trabalho, "previa_analise.xlsx")
        print(f"[INFO] Salvando arquivo de análise em: {caminho_saida}")
        with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
            df_draft.to_excel(writer, sheet_name='Draft', index=False)
            df_ush.to_excel(writer, sheet_name='USH', index=False)
            df_fca_lag.to_excel(writer, sheet_name='FCA_LAG_1_STC', index=False)
        
        print("[SUCESSO] Processamento de dados concluído.")
        return True

    except Exception as e:
        print(f"[ERRO] Falha no processamento dos dados. Detalhe: {e}")
        return False

if __name__ == '__main__':
    if not processar_dados(os.getcwd()):
        exit(1)