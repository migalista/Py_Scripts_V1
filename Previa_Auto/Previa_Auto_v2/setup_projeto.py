import os
import getpass

def verificar_e_criar_pasta(caminho_da_pasta):
    """
    Verifica se a pasta existe. Se não, a cria e retorna uma mensagem clara.
    Retorna uma tupla: (status_da_operacao, caminho_da_pasta)
    """
    if os.path.exists(caminho_da_pasta):
        return "Já Existia", caminho_da_pasta
    else:
        os.makedirs(caminho_da_pasta)
        return "Criada Agora", caminho_da_pasta

# --- Configurações de Caminho ---
# Usuário atual do Windows
usuario = getpass.getuser()

# Caminhos locais (mesma pasta do projeto)
base_local = os.path.dirname(os.path.abspath(__file__))
pastas_locais = [
    os.path.join(base_local, "ExtracaoSAP"),
    os.path.join(base_local, "Referencias")
]

# Caminhos do SharePoint (OneDrive da empresa)
base_sharepoint = rf"C:\Users\{usuario}\OneDrive - Henkel\IB Plan LATAM - Documents"
pastas_sharepoint = [
    os.path.join(base_sharepoint, "PBI", "FCA")
]

# --- Execução ---
print("Iniciando verificação e criação da estrutura de pastas...\n")

# Processar pastas locais
print("## Status das Pastas Locais ##")
for pasta in pastas_locais:
    status, caminho = verificar_e_criar_pasta(pasta)
    print(f"[{status.upper()}] - {caminho}")

print("\n" + "="*60 + "\n")

# Processar pastas do SharePoint
print("## Status das Pastas do SharePoint ##")
for pasta in pastas_sharepoint:
    status, caminho = verificar_e_criar_pasta(pasta)
    print(f"[{status.upper()}] - {caminho}")

print("\n" + "="*60 + "\n")

# --- Resumo e Instruções Finais ---
print("✅ A verificação da estrutura de pastas foi concluída com sucesso!")
print("\nInstruções Adicionais:")
print("- Para 'ExtracaoSAP': coloque o arquivo 'LISTCUBE_Export.xlsx'")
print("- Para 'Referencias': coloque os arquivos 'referencia_apps.xlsx' e 'referencia_fca_lag.xlsx'")
print("- Para 'SharePoint/PBI/FCA': coloque o arquivo 'Prévia FCA 2023.xlsx'")