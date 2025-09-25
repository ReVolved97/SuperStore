import streamlit as st
import pandas as pd
import plotly.express as px


# 1. Carregar os dados (Usando pd.read_excel para arquivos .xls)
@st.cache_data
def load_data():
    # OBSERVAÇÃO IMPORTANTE: Use o caminho do seu arquivo com 'r' e use pd.read_excel
    # Substitua pelo seu caminho real:
    # df = pd.read_excel(r'C:\Users\Rafael\Downloads\Sample - Superstore.xls')

    # Para fins de demonstração, vou simular o carregamento correto do Excel.
    # Por favor, use a linha acima com o seu caminho correto.
    try:
        df = pd.read_excel(r'C:\Users\Rafael\Downloads\Sample - Superstore.xls')

        # 2. Conversão e Limpeza de Datas na função de carregamento
        # É uma boa prática garantir que 'Order Date' seja datetime aqui.
        if 'Order Date' in df.columns:
            df['Order Date'] = pd.to_datetime(df['Order Date'], errors='coerce')

        # Remove linhas com datas inválidas, se houver
        df.dropna(subset=['Order Date'], inplace=True)

        return df
    except FileNotFoundError:
        st.error(
            "Erro: Arquivo não encontrado. Verifique o caminho. Certifique-se de que o arquivo 'Sample - Superstore.xls' está no local correto.")
        return pd.DataFrame()  # Retorna DataFrame vazio em caso de erro


data = load_data()

# Verifica se o DataFrame foi carregado (útil após a checagem de erro)
if data.empty:
    st.stop()

# Título do Dashboard
st.title('📊 Dashboard de Vendas Superstore')
st.markdown('### Análise Interativa de Vendas, Lucros e Categorias')

# --- Sidebar para os filtros ---
st.sidebar.header('Filtros')

# Criar um filtro de Categoria
categorias = data['Category'].unique().tolist()
categoria_selecionada = st.sidebar.multiselect('Selecione a Categoria:', categorias, default=categorias)

# Criar um filtro de Segmento do Cliente
segmentos = data['Segment'].unique().tolist()
segmento_selecionado = st.sidebar.multiselect('Selecione o Segmento:', segmentos, default=segmentos)

# Filtrar o DataFrame com base na seleção do usuário
df_filtrado = data[data['Category'].isin(categoria_selecionada) & data['Segment'].isin(segmento_selecionado)]

# Se o filtro resultar em um DataFrame vazio, mostre uma mensagem e pare a execução
if df_filtrado.empty:
    st.warning("Nenhum dado encontrado com os filtros selecionados. Por favor, ajuste os filtros.")
    st.stop()

# --- Métricas Principais (KPIs) ---
total_vendas = df_filtrado['Sales'].sum()
total_lucro = df_filtrado['Profit'].sum()

# Formatação dos KPIs
col1, col2 = st.columns(2)
with col1:
    st.metric(label='💰 Vendas Totais', value=f'R${total_vendas:,.2f}')
with col2:
    st.metric(label='📈 Lucro Total', value=f'R${total_lucro:,.2f}')

# --- Visualizações ---

# 1. Vendas por Sub-Categoria (Gráfico de Barras)
st.header('Vendas por Sub-Categoria')
vendas_por_subcategoria = df_filtrado.groupby('Sub-Category')['Sales'].sum().reset_index()
# Ordena o gráfico para melhor visualização
vendas_por_subcategoria = vendas_por_subcategoria.sort_values(by='Sales', ascending=False)

fig1 = px.bar(
    vendas_por_subcategoria,
    x='Sub-Category',
    y='Sales',
    title='Vendas por Sub-Categoria de Produto',
    color='Sub-Category'
)
st.plotly_chart(fig1, use_container_width=True)

# 2. Tendência de Vendas ao Longo do Tempo (Gráfico de Linha)
st.header('Tendência de Vendas ao Longo do Tempo')

# Correção da lógica: Agrupar por mês no DataFrame filtrado
# Evita o "SettingWithCopyWarning" e garante que a data já foi tratada em load_data()
vendas_por_mes = df_filtrado.groupby(df_filtrado['Order Date'].dt.to_period('M'))['Sales'].sum().reset_index()
vendas_por_mes['Order Date'] = vendas_por_mes['Order Date'].astype(str)

fig2 = px.line(
    vendas_por_mes,
    x='Order Date',
    y='Sales',
    title='Tendência de Vendas Mensal',
    markers=True  # Adiciona pontos para facilitar a visualização
)
st.plotly_chart(fig2, use_container_width=True)

# 3. Lucro por Categoria (Gráfico de Pizza/Donut para Proporção)
st.header('Proporção de Lucro por Categoria')
lucro_por_categoria = df_filtrado.groupby('Category')['Profit'].sum().reset_index()

fig3 = px.pie(
    lucro_por_categoria,
    values='Profit',
    names='Category',
    title='Distribuição de Lucro por Categoria',
    hole=.3  # Cria um gráfico de Donut
)
st.plotly_chart(fig3, use_container_width=True)

# Opcional: Tabela de Dados
st.markdown("---")
if st.checkbox('Mostrar Tabela de Dados Filtrados'):
    st.subheader('Dados Filtrados')
    st.dataframe(df_filtrado)
