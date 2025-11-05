from flask import Flask, render_template, request 
import random 
import pandas as pd
import plotly.graph_objs as go
import plotly.utils
import json
import os
import base64
from charts import *

app = Flask(__name__)

def generate_random_color():
    r = random.randint(0, 255)
    g = random.randint(0, 255)
    b = random.randint(0, 255)
    return f"#{r:02x}{g:02x}{b:02x}"

def load_spreadsheets():
    """Carrega todas as planilhas da pasta dados_teste"""
    data_folder = '../dados_teste'
    spreadsheets = {}
    
    # Lista todos os arquivos .xlsx na pasta
    for filename in os.listdir(data_folder):
        if filename.endswith('.xlsx') and not filename.startswith('~'):
            filepath = os.path.join(data_folder, filename)
            try:
                df = pd.read_excel(filepath)
                spreadsheet_name = filename.replace('.xlsx', '')
                spreadsheets[spreadsheet_name] = df
            except Exception as e:
                print(f"Erro ao carregar {filename}: {e}")
    
    return spreadsheets

def create_safras_charts(df):
    """Cria gráficos de barras empilhadas separados por cultivo para dados de Safras"""
    charts = []
    
    # Normalizar os tipos de cultivo (Grão e Grãos são o mesmo)
    df = df.copy()
    df['Cultivo'] = df['Cultivo'].replace('Grãos', 'Grão')
    
    # Cores para cada medida
    medida_colors = {
        'Área(milhão ha)': 'rgba(55, 128, 191, 0.8)',
        'Produção(milhão ton)': 'rgba(128, 0, 128, 0.8)',
        'Produção(milhão sc)': 'rgba(128, 0, 128, 0.8)',
        'Produtividade(mil kg/há)': 'rgba(255, 165, 0, 0.8)',
        'Produtividade(sc/há)': 'rgba(255, 165, 0, 0.8)'
    }
    
    # Cores de fundo para cada cultivo
    cultivo_bg_colors = {
        'Café': 'rgba(255, 140, 0, 0.1)',
        'Cana': 'rgba(70, 130, 180, 0.1)',
        'Grão': 'rgba(70, 130, 180, 0.1)'
    }
    
    # Criar um gráfico separado para cada cultivo
    for cultivo in ['Café', 'Cana', 'Grão']:
        df_cultivo = df[df['Cultivo'] == cultivo]
        
        if df_cultivo.empty:
            continue
            
        # Criar figura para este cultivo
        fig = go.Figure()
        
        # Obter medidas únicas para este cultivo
        medidas = df_cultivo['Medida'].unique()
        
        # Adicionar uma barra para cada medida
        for medida in medidas:
            df_medida = df_cultivo[df_cultivo['Medida'] == medida]
            
            fig.add_trace(go.Bar(
                x=df_medida['Safra'],
                y=df_medida['Valor'],
                name=medida,
                marker_color=medida_colors.get(medida, 'rgba(128, 128, 128, 0.8)'),
                text=df_medida['Valor'].round(1),
                textposition='inside',
                textfont=dict(color='white', size=10)
            ))
        
        # Configurar layout específico para cada cultivo
            fig.update_layout(
            title=f'Safra {cultivo}',
            xaxis_title='',
            yaxis_title='Medida',
            template='plotly_white',
            height=400,
            barmode='group',
            plot_bgcolor=cultivo_bg_colors.get(cultivo, 'white'),
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=-0.4,
                xanchor="center",
                x=0.5
            ),
            
        )
        
        charts.append({
            'title': f'Safra {cultivo}',
            'chart': json.dumps(fig, cls=plotly.utils.PlotlyJSONEncoder)
        })
    
    return charts

def create_vbp_charts(df):
    """Cria gráficos de barras e linhas para dados de VBP"""
    charts = []
    
    # Converter valores para trilhões para melhor visualização
    df['Variacao_Trilhoes'] = df['Variação(R$)'] / 1e12
    
    # Gráfico de Barras - VBP por Ano
    df_anual = df.groupby('Ano')['Variacao_Trilhoes'].mean().reset_index()
    
    bar_chart = go.Figure(data=[
        go.Bar(
            x=df_anual['Ano'],
            y=df_anual['Variacao_Trilhoes'],
            name='VBP Médio (R$ Trilhões)',
            marker_color='rgba(219, 64, 82, 0.7)',
            text=df_anual['Variacao_Trilhoes'].round(2),
            textposition='auto'
        )
    ])
    
    bar_chart.update_layout(
        title='VBP - Valor Médio por Ano',
        xaxis_title='Ano',
        yaxis_title='VBP (R$ Trilhões)',
        template='plotly_white',
        height=400
    )
    
    charts.append({
        'title': 'Gráfico de Barras - VBP Médio Anual',
        'chart': json.dumps(bar_chart, cls=plotly.utils.PlotlyJSONEncoder)
    })
    
    # Gráfico de Linhas - Evolução Mensal do VBP
    # Pegar os últimos 24 meses para melhor visualização
    df_recent = df.tail(24).copy()
    df_recent['Periodo'] = df_recent['Mês'] + '/' + df_recent['Ano'].astype(str)
    
    line_chart = go.Figure()
    
    line_chart.add_trace(go.Scatter(
        x=df_recent['Periodo'],
        y=df_recent['Variacao_Trilhoes'],
        mode='lines+markers',
        name='VBP (R$ Trilhões)',
        line=dict(width=3, color='rgba(50, 171, 96, 1)'),
        marker=dict(size=8, color='rgba(50, 171, 96, 0.8)')
    ))
    
    line_chart.update_layout(
        title='VBP - Evolução Mensal (Últimos 24 Meses)',
        xaxis_title='Período',
        yaxis_title='VBP (R$ Trilhões)',
        template='plotly_white',
        height=400,
        xaxis=dict(tickangle=45)
    )
    
    charts.append({
        'title': 'Gráfico de Linhas - Evolução Mensal',
        'chart': json.dumps(line_chart, cls=plotly.utils.PlotlyJSONEncoder)
    })
    
    return charts

def create_dolar_charts(df):
    """Cria gráfico de linha de tendência para dados do Dólar"""
    charts = []
    
    # Criar uma coluna de período combinando Ano e Mês para melhor visualização
    df_sorted = df.copy()
    df_sorted = df_sorted.sort_values(['Ano', 'Nº Mês'])
    

    # Criar rótulos de período (Mês/Ano)
    df_sorted['Período'] = df_sorted['Mês'] + '/' + df_sorted['Ano'].astype(str)
    
    # Gráfico de linha de tendência
    line_chart = go.Figure()
    
    line_chart.add_trace(go.Scatter(
        x=df_sorted['Período'],
        y=df_sorted['Preço'],
        mode='lines+markers',
        name='Cotação USD/BRL',
        line=dict(width=3, color='rgba(0, 128, 0, 1)'),
        marker=dict(size=8, color='rgba(0, 128, 0, 0.8)'),
        text=df_sorted['Preço'].round(3),
        textposition='top center',
        hovertemplate='<b>%{x}</b><br>Cotação: R$ %{y:.2f}<extra></extra>'
    ))
    
    line_chart.update_layout(
        title='Dólar - Tendência de Cotação (USD/BRL)',
        xaxis_title='Período',
        yaxis_title='Cotação (R$)',
        template='plotly_white',
        height=400,
        xaxis=dict(tickangle=45),
        showlegend=True,
        hovermode='x unified'
    )
    
    charts.append({
        'title': 'Gráfico de Linhas - Tendência de Cotação',
        'chart': json.dumps(line_chart, cls=plotly.utils.PlotlyJSONEncoder)
    })
    
    # Gráfico de barras para comparação mensal
    bar_chart = go.Figure()
    
    bar_chart.add_trace(go.Bar(
        x=df_sorted['Período'],
        y=df_sorted['Preço'],
        name='Cotação USD/BRL',
        marker_color='rgba(0, 128, 0, 0.7)',
        text=df_sorted['Preço'].round(2),
        textposition='outside',
        hovertemplate='<b>%{x}</b><br>Cotação: R$ %{y:.2f}<extra></extra>'
    ))
    
    bar_chart.update_layout(
        title='Dólar - Comparação Mensal (USD/BRL)',
        xaxis_title='Período',
        yaxis_title='Cotação (R$)',
        template='plotly_white',
        height=400,
        xaxis=dict(tickangle=45)
    )
    
    charts.append({
        'title': 'Gráfico de Barras - Comparação Mensal',
        'chart': json.dumps(bar_chart, cls=plotly.utils.PlotlyJSONEncoder)
    })
    
    return charts

def create_agrotoxicos_charts(df):
    """Cria gráficos de evolução mensal e anual para dados de Agrotóxicos"""
    charts = []
    
    # Filtrar dados válidos (remover NaN)
    df_clean = df.dropna(subset=['Ano', 'Sub-Grupo', 'Preço']).copy()
    
    # Cores para cada sub-grupo
    subgrupo_colors = {
        'Inseticida': 'rgba(255, 99, 132, 0.8)',
        'Herbicida': 'rgba(54, 162, 235, 0.8)',
        'Fungicida': 'rgba(255, 206, 86, 0.8)'
    }
    
    # 1. Gráfico de Evolução Anual - Preço Médio por Sub-Grupo
    df_anual = df_clean.groupby(['Ano', 'Sub-Grupo'])['Preço'].mean().reset_index()
    
    annual_chart = go.Figure()
    
    for subgrupo in df_anual['Sub-Grupo'].unique():
        df_sub = df_anual[df_anual['Sub-Grupo'] == subgrupo]
        annual_chart.add_trace(go.Scatter(
            x=df_sub['Ano'],
            y=df_sub['Preço'],
            mode='lines+markers',
            name=subgrupo,
            line=dict(width=3, color=subgrupo_colors.get(subgrupo, 'rgba(128, 128, 128, 0.8)')),
            marker=dict(size=8),
            text=df_sub['Preço'].round(2),
            textposition='top center',
            hovertemplate='<b>%{fullData.name}</b><br>Ano: %{x}<br>Preço Médio: R$ %{y:.2f}<extra></extra>'
        ))
    
    annual_chart.update_layout(
        title='Agrotóxicos - Evolução Anual do Preço Médio por Categoria',
        xaxis_title='Ano',
        yaxis_title='Preço Médio (R$)',
        template='plotly_white',
        height=450,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.2,
            xanchor="center",
            x=0.5
        ),
        hovermode='x unified'
    )
    
    charts.append({
        'title': 'Evolução Anual - Preço Médio por Categoria',
        'chart': json.dumps(annual_chart, cls=plotly.utils.PlotlyJSONEncoder)
    })
    
    # 2. Gráfico de Barras - Comparação de Preços por Sub-Grupo (Ano Mais Recente)
    ano_recente = df_clean['Ano'].max()
    df_recente = df_clean[df_clean['Ano'] == ano_recente]
    df_barras = df_recente.groupby('Sub-Grupo')['Preço'].agg(['mean', 'count']).reset_index()
    
    bar_chart = go.Figure()
    
    for i, row in df_barras.iterrows():
        subgrupo = row['Sub-Grupo']
        bar_chart.add_trace(go.Bar(
            x=[subgrupo],
            y=[row['mean']],
            name=subgrupo,
            marker_color=subgrupo_colors.get(subgrupo, 'rgba(128, 128, 128, 0.8)'),
            text=f'R$ {row["mean"]:.2f}',
            textposition='outside',
            hovertemplate=f'<b>{subgrupo}</b><br>Preço Médio: R$ %{{y:.2f}}<br>Produtos: {int(row["count"])}<extra></extra>'
        ))
    
    bar_chart.update_layout(
        title=f'Agrotóxicos - Preço Médio por Categoria ({ano_recente})',
        xaxis_title='Categoria',
        yaxis_title='Preço Médio (R$)',
        template='plotly_white',
        height=400,
        showlegend=False
    )
    
    charts.append({
        'title': f'Preço Médio por Categoria ({ano_recente})',
        'chart': json.dumps(bar_chart, cls=plotly.utils.PlotlyJSONEncoder)
    })
    
    # 3. Gráfico de Linha - Evolução de Preços por Produto
    # Selecionar os produtos mais comuns
    produtos_mais_comuns = df_clean['Produto'].value_counts().index.tolist()
    df_produtos = df_clean[df_clean['Produto'].isin(produtos_mais_comuns)]
    
    # Agrupar por produto e ano para calcular preço médio
    df_produto_ano = df_produtos.groupby(['Produto', 'Ano'])['Preço'].mean().reset_index()
    
    # Cores para os produtos
    produto_colors = [
        '#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7',
        '#DDA0DD', '#98D8C8', '#F7DC6F', '#BB8FCE', '#85C1E9',
        '#F8C471', '#82E0AA', '#F1948A', '#85C1E9', '#D7BDE2'
    ]
    
    product_line_chart = go.Figure()
    
    # Adicionar linhas para cada produto (inicialmente mostrar apenas os primeiros 5)
    for i, produto in enumerate(produtos_mais_comuns):
        df_produto = df_produto_ano[df_produto_ano['Produto'] == produto]
        
        if not df_produto.empty:
            # Truncar nome do produto se for muito longo
            produto_nome = produto[:25] + '...' if len(produto) > 25 else produto
            
            # Mostrar apenas os primeiros 5 produtos inicialmente
            visible = True if i < 5 else 'legendonly'
            
            product_line_chart.add_trace(go.Scatter(
                x=df_produto['Ano'],
                y=df_produto['Preço'],
                mode='lines+markers',
                name=produto_nome,
                line=dict(width=2, color=generate_random_color()),
                marker=dict(size=6),
                visible=visible,
                hovertemplate=f'<b>{produto_nome}</b><br>Ano: %{{x}}<br>Preço Médio: R$ %{{y:.2f}}<extra></extra>'
            ))
    
    product_line_chart.update_layout(
        title='Agrotóxicos - Evolução de Preços por Produto',
        xaxis_title='Ano',
        yaxis_title='Preço Médio (R$)',
        template='plotly_white',
        height=500,
        legend=dict(
            orientation="v",
            yanchor="top",
            y=1,
            xanchor="left",
            x=1.02
        ),
        hovermode='x unified',
        margin=dict(l=50, r=180, t=80, b=50),
        sliders=[
            dict(
                active=4,  # Valor inicial (5 produtos)
                currentvalue={"prefix": "Produtos: "},
                pad={"t": 20},
                steps=[
                    dict(
                        label=str(i+1),
                        method="restyle",
                        args=[{"visible": [True if j <= i else 'legendonly' for j in range(len(produtos_mais_comuns))]}]
                    ) for i in range(len(produtos_mais_comuns))
                ],
                x=0.1,
                len=0.8,
                y=0,
                yanchor="top",
                ticklen=0,
                tickwidth=0
            )
        ]
    )
    
    charts.append({
        'title': 'Evolução de Preços por Produto',
        'chart': json.dumps(product_line_chart, cls=plotly.utils.PlotlyJSONEncoder)
    })
  
    return charts

def create_credito_rural_charts(df):
    """Cria gráficos para dados de Crédito Rural"""
    charts = []
    
    # Filtrar dados válidos
    df_clean = df.dropna(subset=['Valor'])
    
    # 1. Gráfico por Ano (dados anuais)
    df_anos = df_clean[df_clean['Mês'].isna()].copy()
    if len(df_anos) > 0:
        df_anos = df_anos.sort_values('Ano')
        
        # Preparar textos de variação para exibir nos pontos
        variacao_texts = []
        for _, row in df_anos.iterrows():
            if pd.notna(row['Variação']):
                variacao_texts.append(f"{row['Variação']*100:.1f}%")
            else:
                variacao_texts.append("")
        
        annual_chart = go.Figure()
        annual_chart.add_trace(go.Bar(
            x=df_anos['Ano'],
            y=df_anos['Valor'],
            name='Valor Anual',
            marker_color='rgba(55, 128, 191, 0.8)',
            text=variacao_texts,
            textposition='outside',
            textfont=dict(size=12, color='black'),
            hovertemplate='<b>Ano: %{x}</b><br>' +
                         'Valor: R$ %{y:,.0f}<br>' +
                         'Variação: %{text}<br>' +
                         '<extra></extra>'
        ))
        
        annual_chart.update_layout(
            title='Crédito Rural - Valores por Ano',
            xaxis_title='Ano',
            yaxis_title='Valor (R$)',
            template='plotly_white',
            height=400
        )
        
        charts.append({
            'title': 'Valores por Ano',
            'chart': json.dumps(annual_chart, cls=plotly.utils.PlotlyJSONEncoder)
        })
    
    # 2. Gráfico por Mês (últimos 12 meses)
    df_meses = df_clean[df_clean['Mês'].notna()].copy()
    if len(df_meses) > 0:
        df_meses = df_meses.sort_values(['Ano', 'Nº Mês']).tail(12)
        df_meses['Label'] = df_meses['Mês'] + '/' + df_meses['Ano'].astype(str)
        
        # Preparar textos de variação mensal para exibir nos pontos
        variacao_mensal_texts = []
        for _, row in df_meses.iterrows():
            if pd.notna(row['Variação Mês']):
                variacao_mensal_texts.append(f"{row['Variação Mês']*100:.1f}%")
            else:
                variacao_mensal_texts.append("")
        
        monthly_chart = go.Figure()
        monthly_chart.add_trace(go.Scatter(
            x=df_meses['Label'],
            y=df_meses['Valor'],
            mode='lines+markers+text',
            name='Valor Mensal',
            line=dict(width=3, color='rgba(255, 99, 132, 0.8)'),
            marker=dict(size=8),
            text=variacao_mensal_texts,
            textposition='top center',
            textfont=dict(size=10, color='black'),
            hovertemplate='<b>Período: %{x}</b><br>' +
                         'Valor: R$ %{y:,.0f}<br>' +
                         'Variação Mensal: %{text}<br>' +
                         '<extra></extra>'
        ))
        
        monthly_chart.update_layout(
            title='Crédito Rural - Últimos 12 Meses',
            xaxis_title='Período',
            yaxis_title='Valor (R$)',
            template='plotly_white',
            height=400,
            xaxis=dict(tickangle=45)
        )
        
        charts.append({
            'title': 'Últimos 12 Meses',
            'chart': json.dumps(monthly_chart, cls=plotly.utils.PlotlyJSONEncoder)
        })
    
    return charts

def create_indicadores_podas_charts(df):
    
    """Cria gráficos para dados de Indicadores de Podas"""
    charts = []

    print('aknskjabflcjsjhsldbvsbdaljvbhçvbdçvosBV~dnvidN')

    
    # Renomear colunas para facilitar o uso e evitar caracteres especiais
    df.columns = df.columns.str.strip().str.replace(' ', '_').str.replace('(', '').str.replace(')', '').str.lower()
    print(df.columns)
   
    # Filtrar dados válidos (remover NaN) apenas para colunas essenciais para agrupamento
    df_clean = df.dropna(subset=['year', 'item']).copy()
    
    # Cores para cada item (geradas dinamicamente)
    items = df_clean['item'].unique()
    item_colors = {item: generate_random_color() for item in items}

    # 1. Gráfico de Barras Agrupadas - Produção(t) por Item e Ano
    bar_chart_producao = go.Figure()
    for item in items:
        df_item = df_clean[df_clean['item'] == item]
        bar_chart_producao.add_trace(go.Bar(
            x=df_item['year'],
            y=df_item['produçãot'],
            name=item,
            marker_color=item_colors.get(item, 'rgba(128, 128, 128, 0.8)'),
            hovertemplate='<b>Item: %{fullData.name}</b><br>Ano: %{x}<br>Produção: %{y:,.0f} t<extra></extra>'
        ))

    bar_chart_producao.update_layout(
        title='Produção (toneladas) por Item e Ano',
        xaxis_title='Ano',
        yaxis_title='Produção (toneladas)',
        template='plotly_white',
        height=450,
        barmode='group',
      
        hovermode='x unified'
    )
    charts.append({
        'title': 'Produção (toneladas) por Item e Ano',
        'chart': base64.b64encode(json.dumps(bar_chart_producao, cls=plotly.utils.PlotlyJSONEncoder).encode('utf-8')).decode('utf-8')
    })

    # 2. Gráfico de Linha - Custo por Hectare por Item e Ano
    line_chart_custo = go.Figure()
    for i, item in enumerate(items):
        df_item = df_clean[df_clean['item'] == item]
        if not df_item.empty:
            visible = True if i < 5 else 'legendonly' # Mostrar apenas os primeiros 5 itens por padrão
            line_chart_custo.add_trace(go.Scatter(
                x=df_item['year'],
                y=df_item['custo_por_ha'],
                mode='lines+markers',
                name=item,
                line=dict(width=2, color=item_colors.get(item)),
                marker=dict(size=6),
                visible=visible,
                hovertemplate=f'<b>{item}</b><br>Ano: %{{x}}<br>Custo por HA: R$ %{{y:.2f}}<extra></extra>'
            ))

    line_chart_custo.update_layout(
        title='Custo por Hectare por Item e Ano',
        xaxis_title='Ano',
        yaxis_title='Custo por HA (R$)',
        template='plotly_white',
        height=500,
        legend=dict(
            orientation="v",
            yanchor="top",
            y=1,
            xanchor="left",
            x=1.02
        ),
        hovermode='x unified',
        margin=dict(l=50, r=180, t=80, b=50),
        sliders=[
            dict(
                active=min(4, len(items) - 1),  # Valor inicial (até 5 itens)
                currentvalue={"prefix": "Itens: "},
                pad={"t": 20},
                steps=[
                    dict(
                        label=str(i+1),
                        method="restyle",
                        args=[{"visible": [True if j <= i else 'legendonly' for j in range(len(items))]}]
                    ) for i in range(len(items))
                ],
                x=0.1,
                len=0.8,
                y=0,
                yanchor="top",
                ticklen=0,
                tickwidth=0
            )
        ]
    )
    charts.append({
        'title': 'Custo por Hectare por Item e Ano',
        'chart': base64.b64encode(json.dumps(line_chart_custo, cls=plotly.utils.PlotlyJSONEncoder).encode('utf-8')).decode('utf-8')
    })

    # 3. Gráfico de Linha - Área Colhida por Item e Ano
    line_chart_area = go.Figure()
    for i, item in enumerate(items):
        df_item = df_clean[df_clean['item'] == item]
        if not df_item.empty:
            visible = True if i < 5 else 'legendonly' # Mostrar apenas os primeiros 5 itens por padrão
            line_chart_area.add_trace(go.Scatter(
                x=df_item['year'],
                y=df_item['área_colhida'],
                mode='lines+markers',
                name=item,
                line=dict(width=2, color=item_colors.get(item)),
                marker=dict(size=6),
                visible=visible,
                hovertemplate=f'<b>{item}</b><br>Ano: %{{x}}<br>Área Colhida: %{{y:,.0f}} ha<extra></extra>'
            ))

    line_chart_area.update_layout(
        title='Área Colhida (hectares) por Item e Ano',
        xaxis_title='Ano',
        yaxis_title='Área Colhida (hectares)',
        template='plotly_white',
        height=500,
        legend=dict(
            orientation="v",
            yanchor="top",
            y=1,
            xanchor="left",
            x=1.02
        ),
        hovermode='x unified',
        margin=dict(l=50, r=180, t=80, b=50),
        sliders=[
            dict(
                active=min(4, len(items) - 1),  # Valor inicial (até 5 itens)
                currentvalue={"prefix": "Itens: "},
                pad={"t": 20},
                steps=[
                    dict(
                        label=str(i+1),
                        method="restyle",
                        args=[{"visible": [True if j <= i else 'legendonly' for j in range(len(items))]}]
                    ) for i in range(len(items))
                ],
                x=0.1,
                len=0.8,
                y=0,
                yanchor="top",
                ticklen=0,
                tickwidth=0
            )
        ]
    )
    charts.append({
        'title': 'Área Colhida (hectares) por Item e Ano',
        'chart': base64.b64encode(json.dumps(line_chart_area, cls=plotly.utils.PlotlyJSONEncoder).encode('utf-8')).decode('utf-8')
    })

    return charts

def create_generic_charts(df, name):
    """Cria gráficos genéricos para planilhas não reconhecidas"""
    charts = []
    print('CRIOU OS GRÀFICOS GENÈRICOS')
    
    # Tentar identificar colunas numéricas e categóricas
    numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
    categorical_cols = df.select_dtypes(include=['object']).columns.tolist()
    
    if len(numeric_cols) > 0 and len(categorical_cols) > 0:
        # Gráfico de Barras
        x_col = categorical_cols[0]
        y_col = numeric_cols[0]
        
        # Agrupar dados se necessário
        if len(df) > 20:
            df_grouped = df.groupby(x_col)[y_col].mean().reset_index()
        else:
            df_grouped = df[[x_col, y_col]].copy()
        
        bar_chart = go.Figure(data=[
            go.Bar(
                x=df_grouped[x_col],
                y=df_grouped[y_col],
                name=y_col,
                marker_color='rgba(158, 185, 243, 0.7)'
            )
        ])
        
        bar_chart.update_layout(
            title=f'{name} - Gráfico de Barras',
            xaxis_title=x_col,
            yaxis_title=y_col,
            template='plotly_white',
            height=400
        )
        
        charts.append({
            'title': 'Gráfico de Barras',
            'chart': json.dumps(bar_chart, cls=plotly.utils.PlotlyJSONEncoder)
        })
        
        # Gráfico de Linhas
        line_chart = go.Figure()
        
        line_chart.add_trace(go.Scatter(
            x=df_grouped[x_col],
            y=df_grouped[y_col],
            mode='lines+markers',
            name=y_col,
            line=dict(width=3)
        ))
        
        line_chart.update_layout(
            title=f'{name} - Gráfico de Linhas',
            xaxis_title=x_col,
            yaxis_title=y_col,
            template='plotly_white',
            height=400
        )
        
        charts.append({
            'title': 'Gráfico de Linhas',
            'chart': json.dumps(line_chart, cls=plotly.utils.PlotlyJSONEncoder)
        })
    
    return charts

@app.route('/')
def index():
    """Página principal com gráficos de todas as planilhas"""
    spreadsheets = load_spreadsheets()
    all_charts = {}
    

    
    for name, df in spreadsheets.items():
        print(name)
        if name.lower() == 'safras':
            charts = create_safras_charts(df)
        elif name.lower() == 'vbp':
            charts = create_vbp_charts(df)
        elif name.lower() == 'dolar':
            charts = create_dolar_charts(df)
        elif name.lower() == 'agrotóxicos':
            charts = create_agrotoxicos_charts(df)
        elif name.lower() == 'credito_rural':
            charts = create_credito_rural_charts(df)
        elif name.lower() == 'indicadores_poda':
            charts = create_indicadores_podas_charts(df)
        else:
            charts = create_generic_charts(df, name)
        
        all_charts[name] = {
            'charts': charts,
            'info': {
                'rows': len(df),
                'columns': len(df.columns),
                'column_names': df.columns.tolist()
            }
        }
    
    return render_template('dashboard.html', all_charts=all_charts)

request_counts={}
ALLOWED_ENDPOINTS = {'index', 'powerbi_dashboard', 'powerbi_sheet'}

@app.before_request
def before_request():
    global request_counts
    endpoint = request.endpoint
    if endpoint in ALLOWED_ENDPOINTS:
        key = endpoint
        if endpoint == 'powerbi_sheet' and request.view_args:
            sheet = request.view_args.get('sheet_name')
            if sheet:
                key = f'{endpoint}:{sheet}'
        request_counts[key] = request_counts.get(key, 0) + 1
        print(f"Request to {key}, Current counts: {request_counts}")

@app.route('/stats')
def stats():
    return render_template('stats.html', request_counts=request_counts)

@app.route('/powerbi')
def powerbi_dashboard():
    """Dashboard estilo PowerBI com paginação"""
    spreadsheets = load_spreadsheets()
    funcs = [principal_tab,  pulverizadores_tab, podas_tab, fert_tab, safras_tab, cultivos_tab]
    return render_template('home.html', current_sheet='None', img_name='',funcs=funcs)

@app.route('/powerbi/testes/skeleton')
def teste_esqueleto():
    return render_template('./components/skeleton.html', current_sheet='', img_name='', title='')

@app.route('/powerbi/paginas/<sheet_name>')
def skeleton_waiter(sheet_name):
    sheet_name = sheet_name.lower()
    print(sheet_name)
    title = sheet_name.capitalize()
    image = f"{sheet_name.lower()}.jpg"
    return render_template('./components/skeleton.html', current_sheet=sheet_name, img_name=image, title=title)

@app.route('/powerbi/principal-data')
def principal_tab():
    """Aba principal com KPIs e gráficos resumidos (VBP, Crédito Rural, PIB Agro e Import/Export)."""
    # Filtros via query string: view=anual|mensal (default: anual)
    view = request.args.get('view', 'anual')
    month = request.args.get('mes')
    year = request.args.get('ano')

    charts = create_principal_charts(view=view, year=year, month=month)

    # Para manter a navegação padrão em abas
    spreadsheets = load_spreadsheets()
    sheet_names = list(spreadsheets.keys())
    
    return render_template('charts.html',
                           img_name='principal.jpg',
                           title='Aba Principal',
                           charts=charts, 
                           current_sheet='Principal')

@app.route('/powerbi/pulverizadores-data')
def pulverizadores_tab():
    spreadsheets = load_spreadsheets()
    sheet_names = list(spreadsheets.keys())
    charts = create_pulv_charts()
    print(charts)
    return render_template('charts.html',
                            img_name='pulv.jpg',
                            title='Pulverizador',
                            charts=charts,
                            sheet_names=sheet_names,
                            current_sheet='Pulverização')

@app.route('/powerbi/podas-data')
def podas_tab():
    spreadsheets = load_spreadsheets()
    sheet_names = list(spreadsheets.keys())
    charts = create_podas_charts()
    return render_template('charts.html',
                            title='Poda',
                            img_name='podas.jpg',
                            charts=charts,
                            sheet_names=sheet_names,
                            current_sheet='Poda')

@app.route('/powerbi/fertilizantes-data')
def fert_tab():
    spreadsheets = load_spreadsheets()
    sheet_names = list(spreadsheets.keys())
    charts = create_fert_charts()
    return render_template('charts.html',
                            img_name='fert.jpg',
                            title='Fertilizante',
                            charts=charts,
                            sheet_names=sheet_names,
                            current_sheet='Fertilizante')
                        
@app.route('/powerbi/safras-data')
def safras_tab():
    spreadsheets = load_spreadsheets()
    sheet_names = list(spreadsheets.keys())
    charts = create_safra_charts()
    return render_template('charts.html',
    img_name='safra.jpg',
                            title='Safra',
                            charts=charts,
                            sheet_names=sheet_names,
                            current_sheet='Safra')

@app.route('/powerbi/cultivos-data')
def cultivos_tab():
    spreadsheets = load_spreadsheets()
    sheet_names = list(spreadsheets.keys())
    charts = create_cultivos_charts()
    return render_template('charts.html',
                            charts=charts,
                            sheet_names=sheet_names,
                            )

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)