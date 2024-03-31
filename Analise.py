# Primeiro importa-se as bibliotecas que serão utilizadas (e se for conveniente cria-se um alias para elas)
# O pandas é uma biblioteca muito utilizada e famosa, usada para criação de dataframes, e análise de dados
# Já o potly é uma biblioteca comumente usada para desenvolvimento de gráficos.
import pandas as pd
import plotly.express as px

# Neste bloco lemos as cplanilhas do arquivo Excel e definimos uma variável para cada planilha.
dfPrincipal = pd.read_excel('./ImersaoPython.xlsx',sheet_name='Principal')
dfTotalAcoes = pd.read_excel('./ImersaoPython.xlsx',sheet_name='Total_de_acoes')
dfTicker = pd.read_excel('./ImersaoPython.xlsx',sheet_name='Ticker')
dfSegmentos = pd.read_excel('./ImersaoPython.xlsx',sheet_name='Planilha4')

# Neste bloco optei por excluir as colunas que não iria utilizar. Também seria possível selecionar as que gostaria de utilizar. 
# Além disso, a função merge me permite fazer o que seria feito com o Procv no excel, juntando informações de duas planilhas pela chave, removo uma das colunas para não 
# manter informações repetidas 
dfPrincipal = dfPrincipal.drop(columns=['Var. Dia (%)','Var. Ano (%)','Var. 12M (%)','Var. Sem. (%)','Volume'])
dfPrincipal = dfPrincipal.rename(columns={'Último (R$)': 'ValorFinal','Var. Mês (%)': 'VariacaoMensal','Val. Mín': 'ValorMinimo','Val. Máx': 'ValorMaximo'})
dfPrincipal = dfPrincipal.merge(dfTicker, left_on='Ativo', right_on='Ticker', how='left').drop(columns=['Ticker'])
dfPrincipal = dfPrincipal.merge(dfTotalAcoes, left_on='Ativo', right_on='Código', how='left').drop(columns=['Código']).rename(columns={'Qtde. Teórica': 'QtdAcoes'})
dfPrincipal = dfPrincipal.merge(dfSegmentos, left_on='Nome', right_on='Empresa', how='left').drop(columns=['Empresa'])

# Esta linha cria uma nova coluna no dataframe, que primeiramente calcula a variação em reais de uma ação e depois multiplica pela quantidade de ações que a empresa possui
dfPrincipal['VarMensalReais'] = dfPrincipal['ValorFinal'] * ((dfPrincipal['VariacaoMensal'] / 100) + 1) * dfPrincipal['QtdAcoes']

# Criei um novo dataframe, para poder fazer uma análise gráfica da variação em reais por cada segmento. (Soma-se os valores obtidos anteriormente agrupando por segmento)
df_analise_segmento = dfPrincipal.groupby('Segmento')['VarMensalReais'].sum().reset_index()

# A linha abaixo foi comentada porque o gráfico gerado era um gráfico de barras, mas este já havia sido exibido na aula do curso
# Optei então por buscar uma forma de gerar um gráfico de setores (pizza)
# fig = px.bar(df_analise_segmento, x='Segmento', y='VarMensalReais', text='VarMensalReais', title='Variação Reais por Segmento')
fig = px.pie(df_analise_segmento, values='VarMensalReais', names='Segmento', title='Variação Reais por Segmento')

# Mantive o print do dataframe gerado para que possa conferir valores. 
# Além disso, é possível como trabalho futuro formatar esse dado para que fique em formato float com duas casas decimais para melhorar a leitura
print(df_analise_segmento)
# Apesar de não poder ser exibido no console, o comando fig.show permite que o python abra em meu navegador o gráfico gerado.
# Não foi necessário usar nenhuma identação visto que não há comandos sendo executados dentro de outras funções como condicionais ou laços de repetição
fig.show()