# Tratamento_Dados_Desligamentos.py

# Autor: Rafael Freitas
# Data: 01/06/2025  

# Bibliotecas necessárias:

# pip install pandas

# Contexto do projeto:

# Tratar os dados coletados do site do CAGED (Cadastro Geral de Empregados e Desempregados)
# para análise posterior.
# O arquivo em questão é: Dados_Desemprego.xlsx (coletado via Web Scraping)
# O objetivo é limpar e organizar os dados de admissões e desligamentos na cidade de São Paulo
# para facilitar a análise posterior.

import pandas as pd
import re

# Caminhos dos arquivos
caminho_dados_desemprego = r"C:\Users\rafae\OneDrive\Área de Trabalho\Rafael\TCC\Entrega\Dados Coletados\Dados_Desemprego.xlsx"
caminho_mapeamento_regiao = r"C:\Users\rafae\OneDrive\Área de Trabalho\Rafael\TCC\Entrega\Dados Coletados\Mapeamento_Regiao.xlsx"

# Carregar os arquivos
df_desemprego = pd.read_excel(caminho_dados_desemprego)
df_mapeamento = pd.read_excel(caminho_mapeamento_regiao)

# **Etapa 1: Renomear colunas**
colunas_renomeadas = {
    'competencia': 'Competência',
    'municipio': 'Município',
    'secao': 'Seção',
    'admissoes': 'Admissões',
    'desligamentos': 'Desligamentos',
    'saldomovimentacao': 'Saldomovimentação'
}

df_desemprego = df_desemprego.rename(columns=colunas_renomeadas)
print("Colunas renomeadas com sucesso!")

# **Etapa 2: Processamento da coluna "Competência"**
df_desemprego['Ano'] = df_desemprego['Competência'].astype(str).str[:4]  # Ano
df_desemprego['Mês'] = df_desemprego['Competência'].astype(str).str[4:6]  # Mês
df_desemprego = df_desemprego.drop(columns=['Competência'])
print("Colunas 'Ano' e 'Mês' criadas e 'Competência' removida!")

# **Etapa 3: Excluir colunas indesejadas**
colunas_excluir = ['Seção', 'Saldomovimentação']
df_desemprego = df_desemprego.drop(columns=[col for col in colunas_excluir if col in df_desemprego.columns])

# **Etapa 4: Cruzamento com Mapeamento_Região**
if 'Cód. Município' not in df_desemprego.columns:
    if 'Município' in df_desemprego.columns and 'Município' in df_mapeamento.columns:
        df_mapeamento = df_mapeamento.rename(columns={'Município': 'Cód. Município', 'Nome_Município': 'Município'})
        df_desemprego = df_desemprego.merge(df_mapeamento[['Cód. Município', 'Município']], left_on='Município', right_on='Cód. Município', how='left')
        
        # Remover a coluna antiga "Município" usada para cruzamento
        df_desemprego = df_desemprego.drop(columns=['Município_x']).rename(columns={'Município_y': 'Município'})

        print("Cruzamento realizado e colunas renomeadas com sucesso!")
    else:
        print("Aviso: A coluna 'Município' não foi encontrada em uma das tabelas. O cruzamento não foi realizado.")
else:
    print("Aviso: A coluna 'Cód. Município' já existe. O cruzamento não foi feito.")

# **Etapa 5: Filtrar apenas 'São Paulo'**
if 'Município' in df_desemprego.columns:
    df_desemprego = df_desemprego[df_desemprego['Município'] == "São Paulo"]
    print("Filtragem concluída: Apenas 'São Paulo' foi mantido.")
else:
    print("Aviso: A coluna 'Município' não foi encontrada. O filtro não foi aplicado.")

# **Etapa 6: Reorganizar colunas**
colunas_ordenadas = ['Ano', 'Mês', 'Cód. Município', 'Município'] + [col for col in df_desemprego.columns if col not in ['Ano', 'Mês', 'Cód. Município', 'Município']]
df_desemprego = df_desemprego[colunas_ordenadas]

# **Etapa 7: Agrupar por Ano e Mês somando Admissões e Desligamentos**
if 'Admissões' in df_desemprego.columns and 'Desligamentos' in df_desemprego.columns:
    df_desemprego = df_desemprego.groupby(['Ano', 'Mês', 'Cód. Município', 'Município'], as_index=False).agg({'Admissões': 'sum', 'Desligamentos': 'sum'})
    print("Soma das colunas 'Admissões' e 'Desligamentos' realizada com sucesso!")
else:
    print("Aviso: As colunas 'Admissões' e/ou 'Desligamentos' não foram encontradas. A soma não foi realizada.")

# **Etapa Final: Salvar o arquivo**
df_desemprego.to_excel(caminho_dados_desemprego, index=False)

print("Script finalizado com sucesso! Colunas renomeadas, filtradas e dados agrupados corretamente.")