import random 
import pandas as pd
import plotly.graph_objs as go
import plotly.utils
import json
import os
import base64
from datetime import datetime
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.express as px

BLUE = "#15458f"
ORANGE = "#EC6608"

def create_principal_charts(view: str = 'anual', year: str | int | None = None, month: str | None = None) -> list[dict]:
    """Para a página principal, precisa do VBP, Credito Rural, Pib AGRO e Importação e exportação de Máquinas agricolas"""
    
    all_charts = []
    
    actual_year = datetime.now().year
    selected_year = int(year) if year is not None and str(year).isdigit() else actual_year
    selected_month = month

    # --- 1. Gráfico VBP ---
    
        # Tenta ler o arquivo original
    vbp = pd.read_excel('../dados_teste/VBP.xlsx')

    vbp['Variação Mensal'] = vbp['Variação(R$)'].pct_change() 
    
      # Agrupamento para VBP Anual: último valor do ano
    vbp_anual = vbp.groupby('Ano')['Variação(R$)'].last().reset_index()

    vbp_atual = vbp[vbp['Ano'] == selected_year]
    
    if view == 'mensal' and selected_month is not None:
        vbp_atual = vbp_atual[vbp_atual['Mês'] == selected_month]
    

    print("Dados = ", vbp_atual.head(5))
    fig_vbp = make_subplots(
        rows=1, cols=2,
        subplot_titles=("VBP (Anual)", "Variação Mensal"),
        horizontal_spacing=0.08
    )
    y_vals_bi = vbp_anual['Variação(R$)'] / 1e9 # Valores em Bilhões (Bi)
    fig_vbp.add_trace(
        go.Bar(
            x=vbp_anual['Ano'].tolist(),
            y=y_vals_bi.tolist(),
            name="VBP Anual",
            marker=dict(color=BLUE,cornerradius=20),
            text=[f"R$ {v:,.2f} bi" for v in y_vals_bi],
            texttemplate='%{text}',
            textposition='outside',
            cliponaxis=False,
        ),
        row=1, col=1 # Explicitação de row e col
    )
    fig_vbp.add_trace(
        go.Scatter(
            x=vbp_atual['Mês'].to_list(), 
            y=vbp_atual['Variação Mensal'].to_list(), 
            name="Variação Mensal",
            mode='lines+markers+text',
            line=dict(color=ORANGE, width=3),
            marker=dict(size=8, symbol='circle'),
            text=[f'{p:+.2%}' for p in vbp_atual['Variação Mensal']], 
            textposition='top center',
        ),
        row=1, col=2 # Explicitação de row e col
    )
    fig_vbp.update_yaxes(
        zeroline=True,
        tickformat='.2%', 
        row=1, col=2
    )
    fig_vbp.update_layout(
        title_text='VBP',
        template="plotly_white",
        yaxis1=dict(
            title="R$ (Bilhões)",
            tickprefix='R$ ',
            ticksuffix=' Bi',
            showgrid=True,
        ),
        yaxis2=dict(
            visible=False
        ),
        xaxis1=dict(
                    title="Ano",
                    type='category'),
        xaxis2=dict(title="Mês"),
        showlegend=False
    )
    chart_json_vbp = fig_vbp.to_json()
    all_charts.append({
        'title': 'VBP',
        'chart': chart_json_vbp
    })

    # --- 2. Gráfico Crédito Rural ---
    try:
        credito_rural = pd.read_excel('../dados_teste/Credito_Rural.xlsx')
    except FileNotFoundError:
        credito_rural = pd.DataFrame({'Ano': [2022, 2023], 'Mês': ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez', 'Jan'], 
                                      'Valor': [4e10, 4.2e10, 4.5e10, 4.4e10, 4.8e10, 5e10, 5.2e10, 5.5e10, 5.8e10, 6e10, 6.2e10, 6.5e10, 6.8e10]})
        credito_rural['Ano'] = credito_rural['Ano'].apply(lambda x: 2023 if x == 'Jan' else 2022)

    credito_rural['Variação Mensal'] = credito_rural['Valor'].pct_change()
    
    credito_rural_anual = credito_rural.groupby('Ano')['Valor'].last().reset_index()


    credito_rural_atual = credito_rural[credito_rural['Ano'] == selected_year]
    if view == 'mensal' and selected_month is not None:
        credito_rural_atual = credito_rural_atual[credito_rural_atual['Mês'] == selected_month]

    fig_crdt_rural = make_subplots(
        rows=1, cols=2,
        subplot_titles=("Crédito Rural (Anual)", "Variação Mensal"),
        horizontal_spacing=0.08
    )

    milhoes_crdt_rural = credito_rural_anual['Valor'] / 1e9 # Ajustado para Bilhões (Bi)
    fig_crdt_rural.add_trace(
        go.Bar(
            x=credito_rural_anual['Ano'],
            y=milhoes_crdt_rural,
            name="Crédito Rural Anual",
            marker=dict(color=BLUE,cornerradius=20),
            text=[f"R$ {v:,.2f} Bi" for v in milhoes_crdt_rural],
            texttemplate='%{text}',
            textposition='outside',
            cliponaxis=False
        ),
        row=1, col=1 # Explicitação de row e col
    )

    fig_crdt_rural.add_trace(
        go.Scatter(
            x=credito_rural_atual['Mês'], 
            y=credito_rural_atual['Variação Mensal'], 
            name="Variação Mensal",
            mode='lines+markers+text',
            line=dict(color=ORANGE, width=3),
            marker=dict(size=8, symbol='circle'),
            text=[f'{p:+.2%}' for p in credito_rural_atual['Variação Mensal']], 
            textposition='top center',
        ),
        row=1, col=2 # Explicitação de row e col
    )

    fig_crdt_rural.update_yaxes(
        tickformat='.2%', 
        row=1, col=2,
        title='Variação Mensal (%)', 
        showgrid=True 
    )

    fig_crdt_rural.update_layout(
        title_text='Crédito Rural',
        template="plotly_white",
        height=520,

        yaxis1=dict(
            title="R$ (Bilhões)", 
            tickprefix='R$ ',
            ticksuffix=' Bi',
            showgrid=True
        ),
        yaxis2=dict(visible=False),   
        xaxis1=dict(title="Ano"),
        xaxis2=dict(title="Mês"),

        showlegend=False
    )
    chart_json_cr = json.dumps(fig_crdt_rural, cls=plotly.utils.PlotlyJSONEncoder)
    all_charts.append({
        'title': 'Crédito Rural',
        'chart': chart_json_cr
    })
    
    pib_agro = pd.read_excel('../dados_teste/PIB.xlsx')
    pib_agro_anual = pib_agro.groupby('Ano')['Variação(R$ Bilhão)'].mean().reset_index()
    pib_agro_anual['Variação(R$ Bilhão)'] = pib_agro_anual['Variação(R$ Bilhão)'] / 1e9

    # fig_pib = px.bar(pib_agro_anual, 
    #              x='Ano', 
    #              y='Variação(R$ Bilhão)',
    #              title='PIB do Agronegócio por Ano'
                 
    #             )

    # fig_pib.update_traces(texttemplate=f'R$ {y} Bi' for , textposition='outside', marker=dict(color=BLUE,cornerradius=20))

    # fig_pib.update_layout(
    #     xaxis_type='category',
    #     xaxis_title='Ano',
    #     yaxis_title='PIB (em R$)', 
    #     yaxis_tickformat='.2s',
    #     template='plotly_white',
 
    # )
    # all_charts.append({
    #     'title': 'PIB Agro',
    #     'chart':json.dumps(fig_pib, cls=plotly.utils.PlotlyJSONEncoder)
    # })
    all_charts.append({
        'title': 'Importação e Exportação de Máquinas Agrícolas (building)',
        'chart': json.dumps(go.Figure(layout=dict(
            annotations=[dict(text='🚧 building 🚧', showarrow=False, font=dict(size=24))],
            template='plotly_white', height=300
        )), cls=plotly.utils.PlotlyJSONEncoder)
    })

    return all_charts

def create_pulv_charts():
    all_charts = []
    
    try:
        agrotoxicos = pd.read_excel('../dados_teste/Agrotóxicos.xlsx')
    except FileNotFoundError:
        agrotoxicos = pd.DataFrame({'Ano': [2020, 2021, 2022, 2023], 'Preço': [140.00, 150.00, 165.50, 170.00]}) 

    indicadores_linhas_anual = agrotoxicos.groupby('Ano')['Preço'].mean().reset_index()
    
    variation = []
    for i, r in indicadores_linhas_anual.iterrows():
        print(i,r)
        if i == 0:
            v = 0
        else:
            v = round(((indicadores_linhas_anual.at[i,'Preço'] - indicadores_linhas_anual.at[i-1,'Preço']) / indicadores_linhas_anual.at[i-1,'Preço'])*100,2) 
        variation.append(v)


    fig_ind_pulv = make_subplots(
        rows=1, cols=1,
        horizontal_spacing=0.05,
        specs=[[{"secondary_y": True}]]  # <--- ADICIONE ISSO
    )

    fig_ind_pulv.add_trace(
        go.Bar(
            x=indicadores_linhas_anual['Ano'],
            y=indicadores_linhas_anual['Preço'],
            name="Preço Médio",
            marker=dict(color=BLUE,cornerradius=20), # Uso de variável de cor
            text=[f"R$ {v:,.2f}" for v in indicadores_linhas_anual['Preço']],
            texttemplate='%{text}',
            textposition='outside',
            cliponaxis=False
        ),
        row=1, col=1 # Explicitação de row e col
    )
    
    fig_ind_pulv.add_trace(
        go.Scatter(
            x=indicadores_linhas_anual['Ano'],
            y=variation,
            name="Preço Médio (Linha)", # Nome ajustado
            mode='lines+markers',
            line=dict(color=ORANGE, width=3), # Uso de variável de cor
            marker=dict(size=8, symbol='circle'),
            text=[f"% {v:,.2f}" for v in variation],
            texttemplate='%{text}',
            textposition='top center',
            cliponaxis=False
        ),
        row=1, col=1, 
        secondary_y=True# Explicitação de row e col
    )
    
    fig_ind_pulv.update_layout(
        title_text='Preço Médio de Agrotóxicos',
        template="plotly_white",
        height=520,

        yaxis1=dict(
            title="Valor (R$)",
            tickprefix='R$ ',
            showgrid=True,
        ),
        yaxis2=dict(
        title="Variação (%)",
        ticksuffix=' %',        # Adiciona um '%' ao lado dos números do eixo
        showgrid=False,        # Desliga a grade para não poluir
        overlaying='y',        # Sobrepõe este eixo ao eixo 'y' (yaxis1)
        side='right'           # Posiciona o eixo à direita
        ),
        xaxis1=dict(title="Ano", type='category'), # 'category' para anos discretos
        showlegend=False
    )
    chart_json_ind_pulv = json.dumps(fig_ind_pulv, cls=plotly.utils.PlotlyJSONEncoder)
    all_charts.append({
        'title': 'Indicadores Linhas',
        'chart': chart_json_ind_pulv
    })

    # Gráfico de Barras de Preços de Pulverizadores (px.bar)
    try:
        dados_pulverizadores = pd.read_excel('../dados_apoio/Preços_pulverizadores.xlsx')
    except FileNotFoundError:
        dados_pulverizadores = pd.DataFrame({'Modelo': ['Modelo A', 'Modelo B', 'Modelo C'], 'Mediana': [500000, 750000, 600000]})

    fig_bar = px.bar(
        dados_pulverizadores,
        x='Mediana',
        y='Modelo',
        orientation='h',
        title='Preço dos Modelos'
    )
    # Atualização de traços e layout para melhor visualização
    fig_bar.update_traces(marker_color=BLUE, texttemplate='R$ %{x:,.0f}', textposition='outside')
    fig_bar.update_layout(xaxis_title='Preço Mediano (R$)', yaxis_title='Modelo', template='plotly_white')

    chart_json_dados_pulverizadores = json.dumps(fig_bar, cls=plotly.utils.PlotlyJSONEncoder)
    all_charts.append({
        'title': 'Preço por Modelo',
        'chart': chart_json_dados_pulverizadores
    })

    pib_agro = pd.read_excel('../dados_teste/PIB.xlsx')
    pib_agro_anual = pib_agro.groupby('Ano')['Variação(R$ Bilhão)'].mean().reset_index()
    pib_agro_anual['Variação(R$ Bilhão)'] = pib_agro_anual['Variação(R$ Bilhão)'] / 1e9

    vendas_jacto = pd.read_excel('../dados_teste/Vendas_jacto.xlsx')
    vendas_jacto = vendas_jacto.loc[vendas_jacto['Product '].str.contains('Sprayers')]
    vendas_jacto['Sales USD'] = pd.to_numeric(vendas_jacto['Sales USD'])
    vendas_jacto_ano = vendas_jacto.groupby('Year')['Sales USD'].mean().reset_index()
    
    print(pib_agro_anual)
    print(vendas_jacto_ano)

    fig_pib_vendas = make_subplots(
        rows=1,cols=1,
        horizontal_spacing=0.05,
        specs=[[{"secondary_y":True}]]
    )
    fig_pib_vendas.add_trace(
        go.Bar(
            x=pib_agro_anual['Ano'],
            y=pib_agro_anual['Variação(R$ Bilhão)'],
            name='PIB Anual',
            marker=dict(color=BLUE,cornerradius=20),
            text=[f"R$ {v} Bi".replace('.0','') for v in pib_agro_anual['Variação(R$ Bilhão)']],
            texttemplate="%{text}",
            textposition='outside',
            cliponaxis=False
        ), row=1,col=1 
    )
    fig_pib_vendas.add_trace(
        go.Scatter(
            x=vendas_jacto_ano['Year'],
            y=vendas_jacto_ano["Sales USD"],
            name="Vendas Jacto", # Nome ajustado
            mode='lines+markers',
            line=dict(color=ORANGE, width=3), # Uso de variável de cor
            marker=dict(size=8, symbol='circle'),
            text=[f"R$ {v:.2f}" for v in vendas_jacto_ano["Sales USD"]],
            texttemplate='%{text}',
            textposition='top center',
            cliponaxis=False
        ),
        row=1, col=1, 
        secondary_y=True# Explicitação de row e col
    )
    
    fig_pib_vendas.update_layout(
        title_text='Pib Agro x Vendas Jacto',
        template="plotly_white",
        height=520,

        yaxis1=dict(
            title="Valor (R$)",
            tickprefix='R$ ',
            showgrid=True,
        ),
        yaxis2=dict(
        title="Variação (%)",
        ticksuffix=' %',        # Adiciona um '%' ao lado dos números do eixo
        showgrid=False,        # Desliga a grade para não poluir
        overlaying='y',        # Sobrepõe este eixo ao eixo 'y' (yaxis1)
        side='right'           # Posiciona o eixo à direita
        ),
        xaxis1=dict(title="Ano", type='category'), # 'category' para anos discretos
        showlegend=False
    )
    
    all_charts.append({
        'title': 'PIB Agro x Faturamento Jacto',
        'chart':json.dumps(fig_pib_vendas, cls=plotly.utils.PlotlyJSONEncoder)
    })

    fig_faturamento = px.bar(
        vendas_jacto_ano,
        x='Year',
        y='Sales USD',
        title='Faturamento Jacto'
    )
    fig_faturamento.update_traces(
        marker=dict(color=BLUE,cornerradius=20),
        texttemplate='R$ %{y:.2f}',
        textposition='outside'
    )
    fig_faturamento.update_layout(
        template="plotly_white"
    )
    all_charts.append({
        'title':'Faturamento Jacto',
        'chart':json.dumps(fig_faturamento, cls=plotly.utils.PlotlyJSONEncoder)
    })
    return all_charts

def create_podas_charts():
    all_charts = []
    
    try:
        podas = pd.read_excel('../dados_teste/Indicadores_Poda.xlsx')
    except FileNotFoundError:
        podas = pd.DataFrame({'Year': [2021, 2022, 2023], 'Área Colhida': [1000, 1500, 1200], 'CUSTO POR HA': [500, 550, 600]})

    # Agrupamento para Gráfico de Podas
    podas_anual = podas.groupby('Year')['Área Colhida'].mean().reset_index()
    podas_anual_mean = podas.groupby('Year')['CUSTO POR HA'].mean().reset_index()

    
    # 1. Gráfico de Podas (Eixos Secundários) - CORRIGIDO
    fig_podas = make_subplots(
        rows=1, cols=1,
        specs=[[{"secondary_y":True}]],
        horizontal_spacing=0.05,
       
    )
    
    # Trace 1: Área Colhida (Eixo Y Primário, Barra)
    fig_podas.add_trace(
        go.Bar(
            x=podas_anual['Year'],
            y=podas_anual['Área Colhida'],
            name="Área Colhida (ha)",
            marker=dict(color=BLUE,cornerradius=20),
            text=[f"{v:,.0f}" for v in podas_anual['Área Colhida']],
            texttemplate='%{text}',
            textposition='outside',
            cliponaxis=False
        ),
        row=1, col=1, # ESSENCIAL: Explicitação de row e col
        secondary_y=False 
    )
    
    # Trace 2: Custo por HA (Eixo Y Secundário, Linha) - CORRIGIDO
    fig_podas.add_trace(
        go.Scatter(
            x=podas_anual_mean['Year'],
            y=podas_anual_mean['CUSTO POR HA'],
            name="CUSTO POR HA (R$)",
            mode='lines+markers',
            line=dict(color=ORANGE, width=3),
            marker=dict(size=8, symbol='circle'),
            text=[f"R$ {v:,.2f}" for v in podas_anual_mean['CUSTO POR HA']],
            texttemplate='%{text}',
            textposition='top center',
            cliponaxis=False
        ),
        row=1, col=1, # ESSENCIAL: Explicitação de row e col
        secondary_y=True
    )
    
    fig_podas.update_layout(
        title_text='Indicadores de Poda',
        template="plotly_white",
        height=520,
        xaxis=dict(title="Ano", type='category'),
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5)
    )

    # Configuração dos eixos Y
    fig_podas.update_yaxes(title_text="Área Colhida (Hectares)", secondary_y=False, tickformat=',.0f')
    fig_podas.update_yaxes(title_text="Custo por HA (R$)", tickprefix='R$ ', secondary_y=True, showgrid=False)

    chart_json_podas = json.dumps(fig_podas, cls=plotly.utils.PlotlyJSONEncoder)
    all_charts.append({
        'title': 'Indicadores de Poda',
        'chart': chart_json_podas
    })

    # 2. Gráfico PIB Agro (CORRIGIDO: Não sobrescreve mais o all_charts)
    try:
        pib_agro = pd.read_excel('../dados_teste/PIB.xlsx')
    except FileNotFoundError:
        pib_agro = pd.DataFrame({'Ano': [2021, 2022, 2023], 'Variação(R$ Bilhão)': [1.5e12, 1.8e12, 2.0e12]})

    df_agg = pib_agro.groupby('Ano')['Variação(R$ Bilhão)'].first().reset_index()
    
    fig_pib = px.bar(df_agg, 
                 x='Ano', 
                 y='Variação(R$ Bilhão)',
                 title='PIB do Agronegócio por Ano'
                 
                )

    fig_pib.update_layout(
        xaxis_type='category',
        xaxis_title='Ano',
        yaxis_title='PIB (em R$)', 
        yaxis_tickformat='.2s',
        template='plotly_white',
 
    )

    fig_pib.update_traces(texttemplate='R$ %{y:.2s}', textposition='outside', marker=dict(color=BLUE,cornerradius=20))

    chart_json_pib = json.dumps(fig_pib, cls=plotly.utils.PlotlyJSONEncoder)
    all_charts.append({
        'title': 'PIB Agro',
        'chart': chart_json_pib
    })
    
    pib_agro = pd.read_excel('../dados_teste/PIB.xlsx')
    pib_agro_anual = pib_agro.groupby('Ano')['Variação(R$ Bilhão)'].mean().reset_index()
    pib_agro_anual['Variação(R$ Bilhão)'] = pib_agro_anual['Variação(R$ Bilhão)'] / 1e9

    vendas_jacto = pd.read_excel('../dados_teste/Vendas_jacto.xlsx')
    vendas_jacto = vendas_jacto.loc[vendas_jacto['Product '].str.contains('Prunning')]
    vendas_jacto['Sales USD'] = pd.to_numeric(vendas_jacto['Sales USD'])
    vendas_jacto_ano = vendas_jacto.groupby('Year')['Sales USD'].mean().reset_index()
    
    print(pib_agro_anual)
    print(vendas_jacto_ano)

    fig_pib_vendas = make_subplots(
        rows=1,cols=1,
        horizontal_spacing=0.05,
        specs=[[{"secondary_y":True}]]
    )
    fig_pib_vendas.add_trace(
        go.Bar(
            x=pib_agro_anual['Ano'],
            y=pib_agro_anual['Variação(R$ Bilhão)'],
            name='PIB Anual',
            marker=dict(color=BLUE,cornerradius=20),
            text=[f"R$ {v} Bi".replace('.0','') for v in pib_agro_anual['Variação(R$ Bilhão)']],
            texttemplate="%{text}",
            textposition='outside',
            cliponaxis=False
        ), row=1,col=1 
    )
    fig_pib_vendas.add_trace(
        go.Scatter(
            x=vendas_jacto_ano['Year'],
            y=vendas_jacto_ano["Sales USD"],
            name="Vendas Jacto", # Nome ajustado
            mode='lines+markers',
            line=dict(color=ORANGE, width=3), # Uso de variável de cor
            marker=dict(size=8, symbol='circle'),
            text=[f"R$ {v:.2f}" for v in vendas_jacto_ano["Sales USD"]],
            texttemplate='%{text}',
            textposition='top center',
            cliponaxis=False
        ),
        row=1, col=1, 
        secondary_y=True# Explicitação de row e col
    )
    
    fig_pib_vendas.update_layout(
        title_text='Pib Agro x Vendas Jacto',
        template="plotly_white",
        height=520,

        yaxis1=dict(
            title="Valor (R$)",
            tickprefix='R$ ',
            showgrid=True,
        ),
        yaxis2=dict(
        title="Variação (%)",
        ticksuffix=' %',        # Adiciona um '%' ao lado dos números do eixo
        showgrid=False,        # Desliga a grade para não poluir
        overlaying='y',        # Sobrepõe este eixo ao eixo 'y' (yaxis1)
        side='right'           # Posiciona o eixo à direita
        ),
        xaxis1=dict(title="Ano", type='category'), # 'category' para anos discretos
        showlegend=False
    )
    
    all_charts.append({
        'title': 'PIB Agro x Faturamento Jacto',
        'chart':json.dumps(fig_pib_vendas, cls=plotly.utils.PlotlyJSONEncoder)
    })



    return all_charts
    
def create_fert_charts():
    # Esta função era uma cópia de create_pulv_charts, mantida a estrutura e aplicada a correção de row/col
    all_charts = []
    
    try:
        agrotoxicos = pd.read_excel('../dados_teste/Agrotóxicos.xlsx')
    except FileNotFoundError:
        agrotoxicos = pd.DataFrame({'Ano': [2020, 2021, 2022, 2023], 'Preço': [140.00, 150.00, 165.50, 170.00]}) 

    indicadores_linhas_anual = agrotoxicos.groupby('Ano')['Preço'].mean().reset_index()
    
    variation = []
    for i, r in indicadores_linhas_anual.iterrows():
        print(i,r)
        if i == 0:
            v = 0
        else:
            v = round(((indicadores_linhas_anual.at[i,'Preço'] - indicadores_linhas_anual.at[i-1,'Preço']) / indicadores_linhas_anual.at[i-1,'Preço'])*100,2) 
        variation.append(v)
    fig_ind_pulv = make_subplots(
        rows=1, cols=1,
        horizontal_spacing=0.05,
        specs=[[{"secondary_y": True}]]
    )

    fig_ind_pulv.add_trace(
        go.Bar(
            x=indicadores_linhas_anual['Ano'],
            y=indicadores_linhas_anual['Preço'],
            name="Preço Médio",
            marker=dict(color=BLUE,cornerradius=20),
            text=[f"R$ {v:,.2f}" for v in indicadores_linhas_anual['Preço']],
            texttemplate='%{text}',
            textposition='outside',
            cliponaxis=False
        ),
        row=1, col=1 # Explicitação de row e col
    )
    fig_ind_pulv.add_trace(
        go.Scatter(
            x=indicadores_linhas_anual['Ano'],
            y=variation,
            name="Preço Médio (Linha)",
            mode='lines+markers',
            line=dict(color=ORANGE, width=3),
            marker=dict(size=8, symbol='circle'),
            text=[f"R$ {v:,.2f}" for v in variation],
            texttemplate='%{text}',
            textposition='top center',
            cliponaxis=False
        ),
        row=1, col=1, # Explicitação de row e col
        secondary_y=True
    )
    fig_ind_pulv.update_layout(
        title_text='Preço Médio de Agrotóxicos (Fertilizantes)',
        template="plotly_white",

        yaxis1=dict(
            title="Valor (R$)",
            tickprefix='R$ ',
            showgrid=True,
        ),
        yaxis2=dict(
        title="Variação (%)",
        ticksuffix=' %',        # Adiciona um '%' ao lado dos números do eixo
        showgrid=False,        # Desliga a grade para não poluir
        overlaying='y',        # Sobrepõe este eixo ao eixo 'y' (yaxis1)
        side='right'           # Posiciona o eixo à direita
        ), 
        xaxis1=dict(title="Ano", type='category'),
        showlegend=False
    )
    chart_json_ind_pulv = json.dumps(fig_ind_pulv, cls=plotly.utils.PlotlyJSONEncoder)
    all_charts.append({
        'title': 'Indicadores Linhas',
        'chart': chart_json_ind_pulv
    })

    pib_agro = pd.read_excel('../dados_teste/PIB.xlsx')
    pib_agro_anual = pib_agro.groupby('Ano')['Variação(R$ Bilhão)'].mean().reset_index()
    pib_agro_anual['Variação(R$ Bilhão)'] = pib_agro_anual['Variação(R$ Bilhão)'] / 1e9

    vendas_jacto = pd.read_excel('../dados_teste/Vendas_jacto.xlsx')
    vendas_jacto = vendas_jacto.loc[vendas_jacto['Product '].str.contains('Fertilizer')]
    vendas_jacto['Sales USD'] = pd.to_numeric(vendas_jacto['Sales USD'])
    vendas_jacto_ano = vendas_jacto.groupby('Year')['Sales USD'].mean().reset_index()
    
    print(pib_agro_anual)
    print(vendas_jacto_ano)

    fig_pib_vendas = make_subplots(
        rows=1,cols=1,
        horizontal_spacing=0.05,
        specs=[[{"secondary_y":True}]]
    )
    fig_pib_vendas.add_trace(
        go.Bar(
            x=pib_agro_anual['Ano'],
            y=pib_agro_anual['Variação(R$ Bilhão)'],
            name='PIB Anual',
            marker=dict(color=BLUE,cornerradius=20),
            text=[f"R$ {v} Bi".replace('.0','') for v in pib_agro_anual['Variação(R$ Bilhão)']],
            texttemplate="%{text}",
            textposition='outside',
            cliponaxis=False
        ), row=1,col=1 
    )
    fig_pib_vendas.add_trace(
        go.Scatter(
            x=vendas_jacto_ano['Year'],
            y=vendas_jacto_ano["Sales USD"],
            name="Vendas Jacto", # Nome ajustado
            mode='lines+markers',
            line=dict(color=ORANGE, width=3), # Uso de variável de cor
            marker=dict(size=8, symbol='circle'),
            text=[f"R$ {v:.2f}" for v in vendas_jacto_ano["Sales USD"]],
            texttemplate='%{text}',
            textposition='top center',
            cliponaxis=False
        ),
        row=1, col=1, 
        secondary_y=True# Explicitação de row e col
    )
    
    fig_pib_vendas.update_layout(
        title_text='Pib Agro x Vendas Jacto',
        template="plotly_white",
        height=520,

        yaxis1=dict(
            title="Valor (R$)",
            tickprefix='R$ ',
            showgrid=True,
        ),
        yaxis2=dict(
        title="Variação (%)",
        ticksuffix=' %',        # Adiciona um '%' ao lado dos números do eixo
        showgrid=False,        # Desliga a grade para não poluir
        overlaying='y',        # Sobrepõe este eixo ao eixo 'y' (yaxis1)
        side='right'           # Posiciona o eixo à direita
        ),
        xaxis1=dict(title="Ano", type='category'), # 'category' para anos discretos
        showlegend=False
    )
    
    all_charts.append({
        'title': 'PIB Agro x Faturamento Jacto',
        'chart':json.dumps(fig_pib_vendas, cls=plotly.utils.PlotlyJSONEncoder)
    })

    all_charts.append({
        'title': 'Importação e Exportação de Máquinas Agrícolas (building)',
        'chart': json.dumps(go.Figure(layout=dict(
            annotations=[dict(text='🚧 building 🚧', showarrow=False, font=dict(size=24))],
            template='plotly_white', height=300
        )), cls=plotly.utils.PlotlyJSONEncoder)
    })
    return all_charts

def create_safra_charts():
    """Cria gráficos de barras agrupadas por cultivo para dados de Safras"""
    charts = []
    
    try:
        df = pd.read_excel('../dados_teste/Safras.xlsx')
    except FileNotFoundError:
        df = pd.DataFrame({
            'Cultivo': ['Café', 'Café', 'Cana', 'Cana', 'Grão', 'Grão'],
            'Safra': ['2020/21', '2021/22', '2020/21', '2021/22', '2020/21', '2021/22'],
            'Medida': ['Área(milhão ha)', 'Área(milhão ha)', 'Produção(milhão ton)', 'Produção(milhão ton)', 'Produtividade(mil kg/há)', 'Produtividade(mil kg/há)'],
            'Valor': [1.5, 1.6, 600, 650, 4.5, 4.8]
        })

    # Normalizar os tipos de cultivo
    df['Cultivo'] = df['Cultivo'].replace('Grãos', 'Grão')
    
    # Cores para cada medida
    medida_colors = {
        'Área(milhão ha)': BLUE, # Uso de variável de cor
        'Produção(milhão ton)': 'rgba(128, 0, 128, 0.8)',
        'Produção(milhão sc)': 'rgba(128, 0, 128, 0.8)',
        'Produtividade(mil kg/há)': ORANGE, # Uso de variável de cor
        'Produtividade(sc/há)': ORANGE
    }
    
    cultivo_bg_colors = {
        'Café': 'rgba(255, 140, 0, 0.1)',
        'Cana': 'rgba(70, 130, 180, 0.1)',
        'Grão': 'rgba(70, 130, 180, 0.1)'
    }
    
    for cultivo in ['Café', 'Cana', 'Grão']:
        df_cultivo = df[df['Cultivo'] == cultivo]
        
        if df_cultivo.empty:
            continue
            
        fig = go.Figure()
        medidas = df_cultivo['Medida'].unique()
        
        for medida in medidas:
            df_medida = df_cultivo[df_cultivo['Medida'] == medida]
            
            fig.add_trace(go.Bar(
                x=df_medida['Safra'],
                y=df_medida['Valor'],
                name=medida,
                marker=dict(color=medida_colors.get(medida, 'rgba(128, 128, 128, 0.8)'), cornerradius=10),
                text=df_medida['Valor'].round(2),
                textposition='outside', # Alterado para outside para não depender do tamanho da barra
                textfont=dict(size=10) # Retirado cor branca para melhor contraste
            ))
            
        fig.update_layout(
            title=f'Safra {cultivo}',
            xaxis_title='Safra',
            yaxis_title='Valor',
            template='plotly_white',
            barmode='group', # Barras agrupadas
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

def create_cultivos_charts():
    all_charts = []
    
    try:
        agrotoxics = pd.read_excel('../dados_teste/Agrotóxicos.xlsx')
    except FileNotFoundError:
        agrotoxics = pd.DataFrame({'Sub-Grupo': ['Herbicida', 'Fungicida', 'Herbicida'], 'Ano': [2022, 2022, 2023], 'Preço': [150.00, 200.00, 160.00]})

    # 1. Gráficos de barra por Sub-Grupo
    agrotoxics_sub_groups = agrotoxics.groupby(['Sub-Grupo','Ano'])['Preço'].mean().reset_index()
    for sub_group in agrotoxics_sub_groups['Sub-Grupo'].unique():
        df = agrotoxics_sub_groups.loc[agrotoxics_sub_groups['Sub-Grupo'] == sub_group]
        fig = px.bar(
            df,
            x='Ano',
            y='Preço',
            title=f'Preço {sub_group}'
        )
        fig.update_traces(marker=dict(color=BLUE,cornerradius=20), texttemplate='R$ %{y:,.2f}', textposition='outside')

        fig.update_layout(
            title=f'Preço por Sub-Grupo: {sub_group}',
            xaxis_title='Ano',
            yaxis_title='Preço',
            yaxis_tickprefix='R$ ',
            template='plotly_white',
            showlegend=False
        )
        all_charts.append({
            'title': f'{sub_group}',
            'chart': json.dumps(fig, cls=plotly.utils.PlotlyJSONEncoder)
        })
    
    # 2. Gráfico de Linhas de Preço Médio de Cultivos
    try:
        df_cultivos = pd.read_excel("../dados_teste/Preço Cultivos.xlsx")
    except FileNotFoundError:
        df_cultivos = pd.DataFrame({'Produto': ['Soja', 'Soja', 'Milho', 'Milho'], 'Ano': [2022, 2023, 2022, 2023], 'Preço medio(R$/Kg)': [2.5, 2.7, 1.8, 1.9]})

    df_cultivos = df_cultivos.sort_values(by='Ano')  
    df_agregado = df_cultivos.groupby(['Ano', 'Produto']).agg(
        Preco_Medio_Anual=('Preço medio(R$/Kg)', 'mean')
    ).reset_index()
    fig = px.line(
        df_agregado,
        x='Ano',
        y='Preco_Medio_Anual',
        color='Produto', 
        markers=True, 
        title='Preço Médio de Cultivos por Ano',
        log_y=True,
        labels={
            "Ano": "Ano",
            "Preco_Medio_Anual": "Preço Médio (R$/Kg)",
            "Produto": "Produto"
            }
    )
    fig.update_layout(
        yaxis_tickprefix='R$ ',
        yaxis_tickformat='.2f',
        xaxis_type='category',
        template='plotly_white',
        height=450
    )

    all_charts.append({
        'title':"Preço dos cultivos por ano",
        'chart':json.dumps(fig, cls=plotly.utils.PlotlyJSONEncoder)
    })

    return all_charts