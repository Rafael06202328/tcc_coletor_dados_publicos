# Tratamento_Dados_Portal_SSP.py

# Autor: Rafael Freitas
# Data: 01/06/2025

# Bibliotecas necessárias:

# pip install pandas

# Contexto do projeto:

# Tratar os dados coletados do site da SSP-SP (Secretaria de Segurança Pública do Estado de São Paulo)
# para análise posterior.
# O arquivo em questão é: Dados_Portal_SSP.xlsx (coletado via Web Scraping)
# O objetivo é limpar e organizar os dados de criminalidade por bairro na cidade de São Paulo

import pandas as pd
import re

# Caminho dos arquivos
caminho_dados_portal = r"C:\Users\rafae\OneDrive\Área de Trabalho\Projetos\TCC\Dados_Portal_SSP.xlsx"
caminho_bairros_dp = r"C:\Users\rafae\OneDrive\Área de Trabalho\Projetos\TCC\Bairros_DP.xlsx"

# Carregar os arquivos Excel
df_portal = pd.read_excel(caminho_dados_portal)
df_bairros = pd.read_excel(caminho_bairros_dp)

# Garantir que os nomes das colunas estejam padronizados
df_portal.columns = df_portal.columns.str.strip()
df_bairros.columns = df_bairros.columns.str.strip()

# Lista de palavras/frases que devem causar a remoção da linha inteira
excluir_linha = [
    "Delegacias", "DDM", "Del. Pol. Atendimento ao Turista", "Idoso", "Santos",
    "DELPOL Metropolitano", "OUTRAS ESPECIALIZADAS", "01 DP Pessoa com Deficiência",
    "Central de Flagrantes II - 91 DP", "Del. Aeroporto Int. Viracopos - CPS (SP)"
]

# Remover linhas onde a coluna "Bairro" contém qualquer palavra/frase da lista de remoção completa
df_portal = df_portal[~df_portal['Bairro'].isin(excluir_linha)]

# Função para limpar texto na coluna "Bairro", mantendo apenas número e "DP"
def limpar_bairro(texto):
    if pd.isna(texto):  # Verifica se a célula está vazia
        return texto
    
    # Captura a numeração e "DP" no início e remove o restante
    match = re.match(r'^(\d{3}\s*DP)', texto)
    if match:
        return match.group(1)  # Retorna apenas a parte encontrada
    
    return ""  # Retorna uma string vazia caso não corresponda ao padrão

# Aplicar a limpeza na coluna "Bairro"
df_portal['Bairro'] = df_portal['Bairro'].apply(limpar_bairro)

# Renomear a coluna "Bairro" para "Código DP"
df_portal.rename(columns={'Bairro': 'Código DP'}, inplace=True)

# Realizar o merge usando 'Código DP' como chave
df_final = df_portal.merge(df_bairros[['Código DP', 'Bairro']], on='Código DP', how='left')

# Remover linhas que ficaram sem informação de 'Código DP' e 'Bairro'
df_final = df_final.dropna(subset=['Código DP', 'Bairro'])

# Para melhor entendimento dessa etapa do script, devido a existência de várias naturezas
# que representam a mesma coisa, como por exemplo:
# "ESTUPRO", "ESTUPRO DE VULNERÁVEL" e "TOTAL DE ESTUPRO (4)", todas essas naturezas representam o crime de "ESTUPRO".
# Por este motivo foi determinada a junção dessas naturezas em uma só.
# Facilitando a complição dos dados posteriormente e o compreendimento dos mesmos.
mapeamento_natureza = {
    "ESTUPRO": "ESTUPRO",
    "ESTUPRO DE VULNERÁVEL": "ESTUPRO",
    "TOTAL DE ESTUPRO (4)": "ESTUPRO",
    "FURTO - OUTROS": "FURTO",
    "FURTO DE VEÍCULO": "FURTO",
    "HOMICÍDIO CULPOSO OUTROS": "HOMICÍDIO",
    "HOMICÍDIO CULPOSO POR ACIDENTE DE TRÂNSITO": "HOMICÍDIO",
    "HOMICÍDIO DOLOSO (2)": "HOMICÍDIO",
    "HOMICÍDIO DOLOSO POR ACIDENTE DE TRÂNSITO": "HOMICÍDIO",
    "Nº DE VÍTIMAS EM HOMICÍDIO DOLOSO (3)": "HOMICÍDIO",
    "Nº DE VÍTIMAS EM HOMICÍDIO DOLOSO POR ACIDENTE DE TRÂNSITO": "HOMICÍDIO",
    "Nº DE VÍTIMAS EM LATROCÍNIO": "HOMICÍDIO",
    "TENTATIVA DE HOMICÍDIO": "HOMICÍDIO",
    "LATROCÍNIO": "LATROCÍNIO",
    "LESÃO CORPORAL CULPOSA - OUTRAS": "LESÃO CORPORAL",
    "LESÃO CORPORAL CULPOSA POR ACIDENTE DE TRÂNSITO": "LESÃO CORPORAL",
    "LESÃO CORPORAL DOLOSA": "LESÃO CORPORAL",
    "LESÃO CORPORAL SEGUIDA DE MORTE": "LESÃO CORPORAL",
    "ROUBO - OUTROS": "ROUBO",
    "ROUBO A BANCO": "ROUBO",
    "ROUBO DE CARGA": "ROUBO",
    "ROUBO DE VEÍCULO": "ROUBO",
    "TOTAL DE ROUBO - OUTROS (1)": "ROUBO"
}

# Aplicando a substituição diretamente na coluna 'Natureza'
df_final['Natureza'] = df_final['Natureza'].replace(mapeamento_natureza)

# Reordenar colunas para que 'Ano' fique antes de 'Código DP' e 'Bairro'
colunas_ordem = ['Ano', 'Código DP', 'Bairro', 'Natureza'] + [col for col in df_final.columns if col not in ['Ano', 'Código DP', 'Bairro', 'Natureza']]
df_final = df_final[colunas_ordem]

# Agora, faça o groupby e aplique a soma corretamente
df_final = df_final.groupby(['Ano', 'Código DP', 'Bairro', 'Natureza'], as_index=False).sum()

# Salvar o arquivo atualizado
df_final.to_excel(caminho_dados_portal, index=False)

print("Processo concluído! Bairros ajustados, dados consolidados corretamente até a coluna 'Natureza' e arquivo atualizado.")