# %%
import pandas as pd
from openpyxl import Workbook
pin = r'C:\POWERBI'

# %%
df_agroatacado = pd.read_csv(pin + r'\AgroAtacado\Valid' + '\\AgroAtacadoDev.csv',sep=';',encoding='utf-8-sig')
df_agroatacado['esteira'] = 'Agro Atacado'
df_agroatacado = df_agroatacado.loc[df_agroatacado['GrpSegmento'].isin(['Corporate', 'Scib'])]
df_agroatacado = df_agroatacado[['esteira','case_id', 'retrabalho_jornada', 'status_caso_fix', 'GrpSegmento', 'segmento', 'valor', 'dt_inicio', 'descricao', 'ofensor', 'Grpvalor', 'nome']]
df_agroatacado.GrpSegmento.unique()

# %%
df_ativosatacado = pd.read_csv(pin + r'\AtivosAtacado\Valid' + '\\Ativos_AtacadoDev.csv',sep=';',encoding='utf-8-sig')
df_ativosatacado['esteira'] = 'Ativos Atacado'
df_ativosatacado = df_ativosatacado.loc[df_ativosatacado['GrpSegmento'].isin(['Corporate', 'Scib'])]
df_ativosatacado = df_ativosatacado[['esteira','case_id', 'retrabalho_jornada', 'status_caso_fix', 'GrpSegmento', 'segmento', 'valor', 'dt_inicio', 'descricao', 'ofensor', 'Grpvalor', 'nome']]
df_ativosatacado.GrpSegmento.unique()

# %%
df_tradefinance = pd.read_csv(pin + r'\TradeFinance\Valid' + '\\TradeFinanceDev.csv',sep=';',encoding='utf-8-sig')
df_tradefinance['esteira'] = 'Trade Finance'
df_tradefinance = df_tradefinance.loc[df_tradefinance['GrpSegmento'].isin(['Corporate', 'Scib'])]
df_tradefinance = df_tradefinance[['esteira','case_id', 'retrabalho_jornada', 'status_caso_fix', 'GrpSegmento', 'segmento', 'valor', 'dt_inicio', 'descricao', 'ofensor', 'Grpvalor', 'nome']]
df_tradefinance.GrpSegmento.unique()


# %%
df_tesouraria = pd.read_csv(pin + r'\Tesouraria\Valid' + '\\TesourariaDev.csv',sep=',',encoding='utf-8-sig')
df_tesouraria['esteira'] = 'Tesouraria'
df_tesouraria = df_tesouraria.loc[df_tesouraria['GrpSegmento'].isin(['Corporate', 'SCIB'])]
df_tesouraria = df_tesouraria[['esteira','case_id', 'retrabalho_jornada', 'status_caso_fix', 'GrpSegmento', 'segmento', 'valor', 'dt_inicio', 'descricao', 'ofensor', 'Grpvalor', 'nome']]
df_tesouraria.GrpSegmento.unique()

# %%
df = pd.concat([df_agroatacado, df_ativosatacado, df_tradefinance, df_tesouraria]) #

df = df.loc[df['status_caso_fix'] == 'Concluído']
df.status_caso_fix.unique()

# %%
df['dt_inicio'] = pd.to_datetime(df['dt_inicio'])
# Filtrar o DataFrame para manter apenas as datas de setembro de 2024
start_date = '2024-09-01'
end_date = '2024-10-01'
df = df[df['dt_inicio'].between(start_date, end_date)]

# %%
print(df.shape)
df = df.drop_duplicates()
print(df.shape)

# %%
# Calcular total de retrabalho
total_retrabalho = df['retrabalho_jornada'].sum()

# Calcular total de jornadas
total_jornadas = df['case_id'].nunique()

# Calcular percentual de retrabalho
df['percentual_retrabalho'] = df['retrabalho_jornada'] / total_jornadas

# Agrupar por descrição e calcular a soma do percentual de retrabalho
descricao_retrabalho = df.groupby('descricao')['percentual_retrabalho'].sum().reset_index()

# Ordenar e encontrar as 3 principais descrições
top_3_descricao = descricao_retrabalho.sort_values(by='percentual_retrabalho', ascending=False).head(3)

# Exibir as 3 principais descrições
print(top_3_descricao)

# %%
df.to_excel('Pedido_Ralph_07_10_2024.xlsx', index=False)


