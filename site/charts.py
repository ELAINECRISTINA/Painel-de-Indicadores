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

def format_currency_brl(value, decimals=2):
    """Formata valores em reais com separadores brasileiros"""
    return f"R$ {value:,.{decimals}f}".replace(",", "X").replace(".", ",").replace("X", ".")

def format_currency_usd(value, decimals=2):
    """Formata valores em dólares com separadores brasileiros"""
    return f"US$ {value:,.{decimals}f}".replace(",", "X").replace(".", ",").replace("X", ".")

def format_percentage(value, decimals=2, show_sign=True):
    """Formata porcentagens com sinal opcional"""
    if show_sign:
        return f"{value:+,.{decimals}%}".replace(",", "X").replace(".", ",").replace("X", ".")
    else:
        return f"{value:,.{decimals}%}".replace(",", "X").replace(".", ",").replace("X", ".")

def format_number_br(value, decimals=2):
    """Formata números grandes com separadores brasileiros"""
    return f"{value:,.{decimals}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def create_principal_charts(view: str = 'anual', year: str | int | None = None, month: str | None = None) -> list[dict]:
    """Para a página principal: VBP, Crédito Rural, PIB Agro e Importação/Exportação."""

    all_charts = []

    actual_year = datetime.now().year
    selected_year = int(year) if year is not None and str(year).isdigit() else actual_year
    selected_month = month

    # --- 1. Gráfico VBP ---
    try:
        # Aba 'VBP': valores anuais por produto. Cabeçalho real está na linha 7 (header=6).
        vbp_raw = pd.read_excel('dados_teste/VBP.xlsx', sheet_name='VBP', header=6)
        vbp_total = vbp_raw[vbp_raw['LAVOURAS'] == 'TOTAL LAVOURAS'].iloc[0]

        # Colunas de ano: 2022, 2023, 2024, 2025, '2026**' (mistura int e str)
        year_cols = [c for c in vbp_raw.columns if str(c).strip('*').isdigit()]
        anos = [int(str(c).strip('*')) for c in year_cols]
        valores_anuais = [vbp_total[c] for c in year_cols]
        vbp_anual = pd.DataFrame({'Ano': anos, 'Variação(R$)': valores_anuais})

        # Aba 'Variação': últimos 6 meses + variação % mês a mês, por produto
        variacao_raw = pd.read_excel('dados_teste/VBP.xlsx', sheet_name='Variação', header=6)
        variacao_total = variacao_raw[variacao_raw['LAVOURAS'] == 'TOTAL LAVOURAS'].iloc[0]

        # Colunas de variação % são as que não são datas nem 'LAVOURAS'
        pct_cols = [c for c in variacao_raw.columns if isinstance(c, str) and '/' in c]
        vbp_atual = pd.DataFrame({
            'Mês': pct_cols,
            'Variação Mensal': [variacao_total[c] / 100 for c in pct_cols]  # já vem em %, divide por 100 pra ficar fração
        })
    except (FileNotFoundError, KeyError, IndexError) as e:
        print(f"⚠️ Usando dados de exemplo para VBP (motivo: {e})")
        vbp_anual = pd.DataFrame({
            'Ano': [2022, 2023, 2024],
            'Variação(R$)': [1.2e12, 1.4e12, 1.5e12]
        })
        vbp_atual = pd.DataFrame({
            'Mês': ['Out', 'Nov', 'Dez'],
            'Variação Mensal': [0.05, 0.08, 0.03]
        })

    print("Dados VBP anual = ", vbp_anual)
    print("Dados VBP variação mensal = ", vbp_atual)
    fig_vbp = make_subplots(
        rows=1, cols=2,
        subplot_titles=("VBP (Anual)", "Variação Mensal"),
        horizontal_spacing=0.08
    )
    y_vals_bi = vbp_anual['Variação(R$)'] / 1e9
    fig_vbp.add_trace(
        go.Bar(
            x=vbp_anual['Ano'].tolist(),
            y=y_vals_bi.tolist(),
            name="VBP Anual",
            marker=dict(color=BLUE, cornerradius=20),
            text=[format_currency_brl(v, 2) + " Bi" for v in y_vals_bi],
            texttemplate='%{text}',
            textposition='outside',
            cliponaxis=False,
        ),
        row=1, col=1
    )
    fig_vbp.add_trace(
        go.Scatter(
            x=vbp_atual['Mês'].to_list(),
            y=vbp_atual['Variação Mensal'].to_list(),
            name="Variação Mensal",
            mode='lines+markers+text',
            line=dict(color=ORANGE, width=3),
            marker=dict(size=8, symbol='circle'),
            text=[f'{p:+,.2%}' for p in vbp_atual['Variação Mensal']],
            textposition='top center',
        ),
        row=1, col=2
    )
    fig_vbp.update_yaxes(zeroline=True, tickformat='.2%', row=1, col=2)
    fig_vbp.update_layout(
        template="plotly_white",
        yaxis1=dict(title="R$ (Bilhões)", tickprefix='R$ ', ticksuffix=' Bi', showgrid=True),
        yaxis2=dict(visible=False),
        xaxis1=dict(title="Ano", type='category'),
        xaxis2=dict(title="Mês"),
        showlegend=False
    )
    all_charts.append({'title': 'VBP', 'chart': fig_vbp.to_json()})

    # --- 2. Gráfico Crédito Rural ---
    # --- 2. Gráfico Crédito Rural ---
    try:
        credito_rural = pd.read_excel('dados_teste/Credito_Rural.xlsx', header=3)
        
        # Renomeia para nomes curtos internamente
        credito_rural = credito_rural.rename(columns={
            'Mês/Ano Protocolo': 'MesAno',
            'Valor comprometido R$': 'Valor'
        })
        
        # Converte MesAno para datetime e extrai Ano e Mês
        credito_rural['MesAno'] = pd.to_datetime(credito_rural['MesAno'], format='%m/%Y')
        credito_rural['Ano'] = credito_rural['MesAno'].dt.year
        credito_rural['Mês'] = credito_rural['MesAno'].dt.strftime('%b/%Y')  # Ex: 'Jul/2024'
        
        # Agrega por mês (soma todos os estados/linhas do mesmo mês)
        credito_rural = credito_rural.groupby(['Ano', 'Mês', 'MesAno'])['Valor'].sum().reset_index()
        credito_rural = credito_rural.sort_values('MesAno')
        
        # Variação mensal
        credito_rural['Variação Mensal'] = credito_rural['Valor'].pct_change()

    except FileNotFoundError:
        credito_rural = pd.DataFrame({
            'Ano':  [2024]*12,
            'Mês':  ['Jan/2024','Fev/2024','Mar/2024','Abr/2024','Mai/2024','Jun/2024',
                    'Jul/2024','Ago/2024','Set/2024','Out/2024','Nov/2024','Dez/2024'],
            'Valor': [4e10,4.2e10,4.5e10,4.4e10,4.8e10,5e10,5.2e10,5.5e10,5.8e10,6e10,6.2e10,6.5e10],
            'Variação Mensal': [None,0.05,0.07,-0.02,0.09,0.04,0.04,0.06,0.05,0.03,0.03,0.05]
        })

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
    milhoes_crdt_rural = credito_rural_anual['Valor'] / 1e9
    fig_crdt_rural.add_trace(
        go.Bar(
            x=credito_rural_anual['Ano'],
            y=milhoes_crdt_rural,
            name="Crédito Rural Anual",
            marker=dict(color=BLUE, cornerradius=20),
            text=[f"R$ {v:.2f} Bi".replace('.', ',') for v in milhoes_crdt_rural],
            texttemplate='%{text}',
            textposition='outside',
            cliponaxis=False
        ),
        row=1, col=1
    )
    fig_crdt_rural.add_trace(
        go.Scatter(
            x=credito_rural_atual['Mês'],
            y=credito_rural_atual['Variação Mensal'],
            name="Variação Mensal",
            mode='lines+markers+text',
            line=dict(color=ORANGE, width=3),
            marker=dict(size=8, symbol='circle'),
            text=[f'{p:+,.2%}'.replace('.', ',') for p in credito_rural_atual['Variação Mensal']],
            textposition='top center',
        ),
        row=1, col=2
    )
    fig_crdt_rural.update_yaxes(tickformat='.2%', row=1, col=2, title='Variação Mensal (%)', showgrid=True)
    fig_crdt_rural.update_layout(
        template="plotly_white",
        height=520,
        yaxis1=dict(title="R$ (Bilhões)", tickprefix='R$ ', ticksuffix=' Bi', showgrid=True),
        yaxis2=dict(visible=False),
        xaxis1=dict(title="Ano"),
        xaxis2=dict(title="Mês"),
        showlegend=False
    )
    all_charts.append({'title': 'Crédito Rural', 'chart': json.dumps(fig_crdt_rural, cls=plotly.utils.PlotlyJSONEncoder)})

    # --- 3. Gráfico PIB Agro ---
    try:
        pib_agro = pd.read_excel('dados_teste/PIB.xlsx')
    except FileNotFoundError:
        pib_agro = pd.DataFrame({'Ano': [2021, 2022, 2023], 'Variação(R$ Bilhão)': [1.5e12, 1.8e12, 2.0e12]})

    pib_agro['Variação(R$ Bilhão)'] = (
        pib_agro['Variação(R$ Bilhão)']
        .astype(str)
        .str.replace('.', '', regex=False)
        .str.replace(',', '.', regex=False)
        .str.replace(r'R\$ ?', '', regex=True)
    )
    pib_agro['Variação(R$ Bilhão)'] = pd.to_numeric(pib_agro['Variação(R$ Bilhão)'], errors='coerce')
    pib_agro_anual = pib_agro.groupby('Ano')['Variação(R$ Bilhão)'].mean().reset_index()
    pib_agro_anual['Variação(R$ Bilhão)'] = pib_agro_anual['Variação(R$ Bilhão)'] / 1e9
    print(pib_agro_anual)

    fig_pib = px.bar(
        pib_agro_anual,
        x='Ano',
        y='Variação(R$ Bilhão)',
        title='PIB do Agronegócio por Ano',
        text=[f'{v} Bi'.replace('.0', '') for v in pib_agro_anual['Variação(R$ Bilhão)']]
    )
    fig_pib.update_traces(textposition='outside', marker=dict(color=BLUE, cornerradius=20))
    fig_pib.update_layout(
        xaxis_type='category',
        xaxis_title='Ano',
        yaxis_title='PIB (em R$)',
        template='plotly_white',
    )
    all_charts.append({'title': 'PIB Agro', 'chart': json.dumps(fig_pib, cls=plotly.utils.PlotlyJSONEncoder)})

    # --- 4. Gráfico Exportação vs Importação ---
    try:
        exportacao = pd.read_excel('dados_teste/Exportacao.xlsx')
        exportacao.columns = [col.strip() for col in exportacao.columns]
        exportacao_anual = exportacao.groupby(['Ano', 'Fluxo'])['Valor (US$)'].sum().reset_index()
        exportacao_pivot = exportacao_anual.pivot(index='Ano', columns='Fluxo', values='Valor (US$)').fillna(0).reset_index()
    except FileNotFoundError:
        exportacao_pivot = pd.DataFrame({
            'Ano': [2022, 2023, 2024],
            'Exportação': [2.5e9, 3.2e9, 3.8e9],
            'Importação': [1.8e9, 2.1e9, 2.3e9]
        })

    fig_exportacao = go.Figure()
    if 'Exportação' in exportacao_pivot.columns:
        valores_exportacao = exportacao_pivot['Exportação'] / 1e6
        fig_exportacao.add_trace(go.Scatter(
            x=exportacao_pivot['Ano'], y=valores_exportacao,
            name="Exportação", mode='lines+markers+text',
            line=dict(color=BLUE, width=3), marker=dict(size=8, symbol='circle'),
            text=[f"US$ {v:.0f} Mi" for v in valores_exportacao],
            textposition='top center',
        ))
    if 'Importação' in exportacao_pivot.columns:
        valores_importacao = exportacao_pivot['Importação'] / 1e6
        fig_exportacao.add_trace(go.Scatter(
            x=exportacao_pivot['Ano'], y=valores_importacao,
            name="Importação", mode='lines+markers+text',
            line=dict(color=ORANGE, width=3), marker=dict(size=8, symbol='triangle-up'),
            text=[f"US$ {v:.0f} Mi" for v in valores_importacao],
            textposition='bottom center',
        ))
    fig_exportacao.update_layout(
        template="plotly_white", height=520,
        yaxis=dict(title="US$ (Milhões)", tickprefix='US$ ', ticksuffix=' Mi', showgrid=True),
        xaxis=dict(title="Ano", type='category'),
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5)
    )
    all_charts.append({'title': 'Importação vs Exportação de Máquinas Agrícolas', 'chart': json.dumps(fig_exportacao, cls=plotly.utils.PlotlyJSONEncoder)})

    return all_charts


def create_pulv_charts():
    all_charts = []

    try:
        agrotoxicos = pd.read_excel('dados_teste/Agrotóxicos.xlsx')
    except FileNotFoundError:
        agrotoxicos = pd.DataFrame({'Ano': [2020, 2021, 2022, 2023], 'Preço': [140.00, 150.00, 165.50, 170.00]})

    indicadores_linhas_anual = agrotoxicos.groupby('Ano')['Preço'].mean().reset_index()

    variation = []
    for i, r in indicadores_linhas_anual.iterrows():
        if i == 0:
            v = 0
        else:
            v = round(((indicadores_linhas_anual.at[i, 'Preço'] - indicadores_linhas_anual.at[i-1, 'Preço']) / indicadores_linhas_anual.at[i-1, 'Preço']) * 100, 2)
        variation.append(v)

    fig_ind_pulv = make_subplots(rows=1, cols=1, horizontal_spacing=0.05, specs=[[{"secondary_y": True}]])
    fig_ind_pulv.add_trace(
        go.Bar(
            x=indicadores_linhas_anual['Ano'], y=indicadores_linhas_anual['Preço'],
            name="Preço Médio", marker=dict(color=BLUE, cornerradius=20),
            text=[f"R$ {v:,.2f}".replace('.', ',') for v in indicadores_linhas_anual['Preço']],
            texttemplate='%{text}', textposition='outside', cliponaxis=False
        ), row=1, col=1
    )
    fig_ind_pulv.add_trace(
        go.Scatter(
            x=indicadores_linhas_anual['Ano'], y=variation,
            name="Variação (%)", mode='lines+markers',
            line=dict(color=ORANGE, width=3), marker=dict(size=8, symbol='circle'),
            text=[f"% {v:,.2f}" for v in variation],
            texttemplate='%{text}', textposition='top center', cliponaxis=False
        ), row=1, col=1, secondary_y=True
    )
    fig_ind_pulv.update_layout(
        template="plotly_white", height=520,
        yaxis1=dict(title="Valor (R$)", tickprefix='R$ ', showgrid=True),
        yaxis2=dict(title="Variação (%)", ticksuffix=' %', showgrid=False, overlaying='y', side='right'),
        xaxis1=dict(title="Ano", type='category'),
        showlegend=False
    )
    all_charts.append({'title': 'Indicadores Linhas', 'chart': json.dumps(fig_ind_pulv, cls=plotly.utils.PlotlyJSONEncoder)})

    try:
        dados_pulverizadores = pd.read_excel('dados_apoio/Preços_pulverizadores.xlsx')
    except FileNotFoundError:
        dados_pulverizadores = pd.DataFrame({'Modelo': ['Modelo A', 'Modelo B', 'Modelo C'], 'Mediana': [500000, 750000, 600000]})

    fig_bar = px.bar(dados_pulverizadores, x='Mediana', y='Modelo', orientation='h')
    fig_bar.update_traces(marker_color=BLUE, texttemplate='R$ %{x:.0f}', textposition='outside')
    fig_bar.update_layout(xaxis_title='Preço Mediano (R$)', yaxis_title='Modelo', template='plotly_white')
    all_charts.append({'title': 'Preço por Modelo', 'chart': json.dumps(fig_bar, cls=plotly.utils.PlotlyJSONEncoder)})

    try:
        pib_agro = pd.read_excel('dados_teste/PIB.xlsx')
    except FileNotFoundError:
        pib_agro = pd.DataFrame({'Ano': [2021, 2022, 2023], 'Variação(R$ Bilhão)': [1.5e12, 1.8e12, 2.0e12]})

    pib_agro['Variação(R$ Bilhão)'] = (
        pib_agro['Variação(R$ Bilhão)'].astype(str)
        .str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
        .str.replace(r'R\$ ?', '', regex=True)
    )
    pib_agro['Variação(R$ Bilhão)'] = pd.to_numeric(pib_agro['Variação(R$ Bilhão)'], errors='coerce')
    pib_agro_anual = pib_agro.groupby('Ano')['Variação(R$ Bilhão)'].mean().reset_index()
    pib_agro_anual['Variação(R$ Bilhão)'] = pib_agro_anual['Variação(R$ Bilhão)'] / 1e9

    try:
        vendas_jacto = pd.read_excel('dados_teste/Vendas_jacto.xlsx')
        vendas_jacto = vendas_jacto.loc[vendas_jacto['Product '].str.contains('Sprayers')]
        vendas_jacto['Sales USD'] = pd.to_numeric(vendas_jacto['Sales USD'])
        vendas_jacto_ano = vendas_jacto.groupby('Year')['Sales USD'].mean().reset_index()
    except FileNotFoundError:
        vendas_jacto_ano = pd.DataFrame({'Year': [2021, 2022, 2023], 'Sales USD': [1e6, 1.2e6, 1.4e6]})

    fig_pib_vendas = make_subplots(rows=1, cols=1, horizontal_spacing=0.05, specs=[[{"secondary_y": True}]])
    fig_pib_vendas.add_trace(
        go.Bar(
            x=pib_agro_anual['Ano'], y=pib_agro_anual['Variação(R$ Bilhão)'],
            name='PIB Anual', marker=dict(color=BLUE, cornerradius=20),
            text=[f"R$ {v} Bi".replace('.0', '') for v in pib_agro_anual['Variação(R$ Bilhão)']],
            texttemplate="%{text}", textposition='outside', cliponaxis=False
        ), row=1, col=1
    )
    fig_pib_vendas.add_trace(
        go.Scatter(
            x=vendas_jacto_ano['Year'], y=vendas_jacto_ano["Sales USD"],
            name="Vendas Jacto", mode='lines+markers',
            line=dict(color=ORANGE, width=3), marker=dict(size=8, symbol='circle'),
            text=[f"US$ {v:.2f}" for v in vendas_jacto_ano["Sales USD"]],
            texttemplate='%{text}', textposition='top center', cliponaxis=False
        ), row=1, col=1, secondary_y=True
    )
    fig_pib_vendas.update_layout(
        template="plotly_white", height=520,
        yaxis1=dict(title="PIB (R$ Bilhões)", tickprefix='R$ ', showgrid=True),
        yaxis2=dict(title="Vendas (US$)", showgrid=False, overlaying='y', side='right'),
        xaxis1=dict(title="Ano", type='category'),
        showlegend=False
    )
    all_charts.append({'title': 'PIB Agro x Faturamento Jacto', 'chart': json.dumps(fig_pib_vendas, cls=plotly.utils.PlotlyJSONEncoder)})

    fig_faturamento = px.bar(vendas_jacto_ano, x='Year', y='Sales USD')
    fig_faturamento.update_traces(marker=dict(color=BLUE, cornerradius=20), texttemplate='$ %{y:.2f}', textposition='outside')
    fig_faturamento.update_layout(template="plotly_white")
    all_charts.append({'title': 'Faturamento Jacto', 'chart': json.dumps(fig_faturamento, cls=plotly.utils.PlotlyJSONEncoder)})

    return all_charts


def create_podas_charts():
    all_charts = []

    try:
        podas = pd.read_excel('dados_teste/Indicadores_Poda.xlsx')
    except FileNotFoundError:
        podas = pd.DataFrame({'Year': [2021, 2022, 2023], 'Área Colhida': [1000, 1500, 1200], 'CUSTO POR HA': [500, 550, 600]})

    podas_anual = podas.groupby('Year')['Área Colhida'].mean().reset_index()
    podas_anual_mean = podas.groupby('Year')['CUSTO POR HA'].mean().reset_index()

    fig_podas = make_subplots(rows=1, cols=1, specs=[[{"secondary_y": True}]], horizontal_spacing=0.05)
    fig_podas.add_trace(
        go.Bar(
            x=podas_anual['Year'], y=podas_anual['Área Colhida'],
            name="Área Colhida (ha)", marker=dict(color=BLUE, cornerradius=20),
            text=[f"{v:,.0f}" for v in podas_anual['Área Colhida']],
            texttemplate='%{text}', textposition='outside', cliponaxis=False
        ), row=1, col=1, secondary_y=False
    )
    fig_podas.add_trace(
        go.Scatter(
            x=podas_anual_mean['Year'], y=podas_anual_mean['CUSTO POR HA'],
            name="CUSTO POR HA (R$)", mode='lines+markers',
            line=dict(color=ORANGE, width=3), marker=dict(size=8, symbol='circle'),
            text=[f"R$ {v:,.2f}" for v in podas_anual_mean['CUSTO POR HA']],
            texttemplate='%{text}', textposition='top center', cliponaxis=False
        ), row=1, col=1, secondary_y=True
    )
    fig_podas.update_layout(
        template="plotly_white", height=520,
        xaxis=dict(title="Ano", type='category'),
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5)
    )
    fig_podas.update_yaxes(secondary_y=False, tickformat=',.0f')
    fig_podas.update_yaxes(tickprefix='R$ ', secondary_y=True, showgrid=False)
    all_charts.append({'title': 'Indicadores de Poda', 'chart': json.dumps(fig_podas, cls=plotly.utils.PlotlyJSONEncoder)})

    try:
        pib_agro = pd.read_excel('dados_teste/PIB.xlsx')
    except FileNotFoundError:
        pib_agro = pd.DataFrame({'Ano': [2021, 2022, 2023], 'Variação(R$ Bilhão)': [1.5e12, 1.8e12, 2.0e12]})

    df_agg = pib_agro.groupby('Ano')['Variação(R$ Bilhão)'].first().reset_index()
    fig_pib = px.bar(df_agg, x='Ano', y='Variação(R$ Bilhão)', title='PIB do Agronegócio por Ano')
    fig_pib.update_layout(xaxis_type='category', xaxis_title='Ano', yaxis_title='PIB (em R$)', yaxis_tickformat='.2s', template='plotly_white')
    fig_pib.update_traces(texttemplate='R$ %{y:.2s}', textposition='outside', marker=dict(color=BLUE, cornerradius=20))
    all_charts.append({'title': 'PIB Agro', 'chart': json.dumps(fig_pib, cls=plotly.utils.PlotlyJSONEncoder)})

    pib_agro2 = pd.read_excel('dados_teste/PIB.xlsx')
    pib_agro2['Variação(R$ Bilhão)'] = (
        pib_agro2['Variação(R$ Bilhão)'].astype(str)
        .str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
        .str.replace(r'R\$ ?', '', regex=True)
    )
    pib_agro2['Variação(R$ Bilhão)'] = pd.to_numeric(pib_agro2['Variação(R$ Bilhão)'], errors='coerce')
    pib_agro_anual = pib_agro2.groupby('Ano')['Variação(R$ Bilhão)'].mean().reset_index()
    pib_agro_anual['Variação(R$ Bilhão)'] = pib_agro_anual['Variação(R$ Bilhão)'] / 1e9

    try:
        vendas_jacto = pd.read_excel('dados_teste/Vendas_jacto.xlsx')
        vendas_jacto = vendas_jacto.loc[vendas_jacto['Product '].str.contains('Prunning')]
        vendas_jacto['Sales USD'] = pd.to_numeric(vendas_jacto['Sales USD'])
        vendas_jacto_ano = vendas_jacto.groupby('Year')['Sales USD'].mean().reset_index()
    except FileNotFoundError:
        vendas_jacto_ano = pd.DataFrame({'Year': [2021, 2022, 2023], 'Sales USD': [1e6, 1.2e6, 1.4e6]})

    fig_pib_vendas = make_subplots(rows=1, cols=1, horizontal_spacing=0.05, specs=[[{"secondary_y": True}]])
    fig_pib_vendas.add_trace(
        go.Bar(
            x=pib_agro_anual['Ano'], y=pib_agro_anual['Variação(R$ Bilhão)'],
            name='PIB Anual', marker=dict(color=BLUE, cornerradius=20),
            text=[f"R$ {v} Bi".replace('.0', '') for v in pib_agro_anual['Variação(R$ Bilhão)']],
            texttemplate="%{text}", textposition='outside', cliponaxis=False
        ), row=1, col=1
    )
    fig_pib_vendas.add_trace(
        go.Scatter(
            x=vendas_jacto_ano['Year'], y=vendas_jacto_ano["Sales USD"],
            name="Vendas Jacto", mode='lines+markers',
            line=dict(color=ORANGE, width=3), marker=dict(size=8, symbol='circle'),
            text=[f"US$ {v:.2f}" for v in vendas_jacto_ano["Sales USD"]],
            texttemplate='%{text}', textposition='top center', cliponaxis=False
        ), row=1, col=1, secondary_y=True
    )
    fig_pib_vendas.update_layout(
        template="plotly_white", height=520,
        yaxis1=dict(title="PIB (R$ Bilhões)", tickprefix='R$ ', showgrid=True),
        yaxis2=dict(title="Vendas (US$)", showgrid=False, overlaying='y', side='right'),
        xaxis1=dict(title="Ano", type='category'),
        showlegend=False
    )
    all_charts.append({'title': 'PIB Agro x Faturamento Jacto', 'chart': json.dumps(fig_pib_vendas, cls=plotly.utils.PlotlyJSONEncoder)})

    return all_charts


def create_fert_charts():
    all_charts = []

    try:
        agrotoxicos = pd.read_excel('dados_teste/Agrotóxicos.xlsx')
    except FileNotFoundError:
        agrotoxicos = pd.DataFrame({'Ano': [2020, 2021, 2022, 2023], 'Preço': [140.00, 150.00, 165.50, 170.00]})

    indicadores_linhas_anual = agrotoxicos.groupby('Ano')['Preço'].mean().reset_index()

    variation = []
    for i, r in indicadores_linhas_anual.iterrows():
        if i == 0:
            v = 0
        else:
            v = round(((indicadores_linhas_anual.at[i, 'Preço'] - indicadores_linhas_anual.at[i-1, 'Preço']) / indicadores_linhas_anual.at[i-1, 'Preço']) * 100, 2)
        variation.append(v)

    fig_ind_pulv = make_subplots(rows=1, cols=1, horizontal_spacing=0.05, specs=[[{"secondary_y": True}]])
    fig_ind_pulv.add_trace(
        go.Bar(
            x=indicadores_linhas_anual['Ano'], y=indicadores_linhas_anual['Preço'],
            name="Preço Médio", marker=dict(color=BLUE, cornerradius=20),
            text=[f"R$ {v:,.2f}" for v in indicadores_linhas_anual['Preço']],
            texttemplate='%{text}', textposition='outside', cliponaxis=False
        ), row=1, col=1
    )
    fig_ind_pulv.add_trace(
        go.Scatter(
            x=indicadores_linhas_anual['Ano'], y=variation,
            name="Variação (%)", mode='lines+markers',
            line=dict(color=ORANGE, width=3), marker=dict(size=8, symbol='circle'),
            text=[f"% {v:,.2f}" for v in variation],
            texttemplate='%{text}', textposition='top center', cliponaxis=False
        ), row=1, col=1, secondary_y=True
    )
    fig_ind_pulv.update_layout(
        title_text='Preço Médio de Agrotóxicos (Fertilizantes)',
        template="plotly_white",
        yaxis1=dict(title="Valor (R$)", tickprefix='R$ ', showgrid=True),
        yaxis2=dict(title="Variação (%)", ticksuffix=' %', showgrid=False, overlaying='y', side='right'),
        xaxis1=dict(title="Ano", type='category'),
        showlegend=False
    )
    all_charts.append({'title': 'Indicadores Linhas', 'chart': json.dumps(fig_ind_pulv, cls=plotly.utils.PlotlyJSONEncoder)})

    try:
        pib_agro = pd.read_excel('dados_teste/PIB.xlsx')
    except FileNotFoundError:
        pib_agro = pd.DataFrame({'Ano': [2021, 2022, 2023], 'Variação(R$ Bilhão)': [1.5e12, 1.8e12, 2.0e12]})

    pib_agro['Variação(R$ Bilhão)'] = (
        pib_agro['Variação(R$ Bilhão)'].astype(str)
        .str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
        .str.replace(r'R\$ ?', '', regex=True)
    )
    pib_agro['Variação(R$ Bilhão)'] = pd.to_numeric(pib_agro['Variação(R$ Bilhão)'], errors='coerce')
    pib_agro_anual = pib_agro.groupby('Ano')['Variação(R$ Bilhão)'].mean().reset_index()
    pib_agro_anual['Variação(R$ Bilhão)'] = pib_agro_anual['Variação(R$ Bilhão)'] / 1e9

    try:
        vendas_jacto = pd.read_excel('dados_teste/Vendas_jacto.xlsx')
        vendas_jacto = vendas_jacto.loc[vendas_jacto['Product '].str.contains('Fertilizer')]
        vendas_jacto['Sales USD'] = pd.to_numeric(vendas_jacto['Sales USD'])
        vendas_jacto_ano = vendas_jacto.groupby('Year')['Sales USD'].mean().reset_index()
    except FileNotFoundError:
        vendas_jacto_ano = pd.DataFrame({'Year': [2021, 2022, 2023], 'Sales USD': [1e6, 1.2e6, 1.4e6]})

    fig_pib_vendas = make_subplots(rows=1, cols=1, horizontal_spacing=0.05, specs=[[{"secondary_y": True}]])
    fig_pib_vendas.add_trace(
        go.Bar(
            x=pib_agro_anual['Ano'], y=pib_agro_anual['Variação(R$ Bilhão)'],
            name='PIB Anual', marker=dict(color=BLUE, cornerradius=20),
            text=[f"R$ {v} Bi".replace('.0', '') for v in pib_agro_anual['Variação(R$ Bilhão)']],
            texttemplate="%{text}", textposition='outside', cliponaxis=False
        ), row=1, col=1
    )
    fig_pib_vendas.add_trace(
        go.Scatter(
            x=vendas_jacto_ano['Year'], y=vendas_jacto_ano["Sales USD"],
            name="Vendas Jacto", mode='lines+markers',
            line=dict(color=ORANGE, width=3), marker=dict(size=8, symbol='circle'),
            text=[f"US$ {v:.2f}" for v in vendas_jacto_ano["Sales USD"]],
            texttemplate='%{text}', textposition='top center', cliponaxis=False
        ), row=1, col=1, secondary_y=True
    )
    fig_pib_vendas.update_layout(
        title_text='PIB Agro x Vendas Jacto',
        template="plotly_white", height=520,
        yaxis1=dict(title="PIB (R$ Bilhões)", tickprefix='R$ ', showgrid=True),
        yaxis2=dict(title="Vendas (US$)", showgrid=False, overlaying='y', side='right'),
        xaxis1=dict(title="Ano", type='category'),
        showlegend=False
    )
    all_charts.append({'title': 'PIB Agro x Faturamento Jacto', 'chart': json.dumps(fig_pib_vendas, cls=plotly.utils.PlotlyJSONEncoder)})

    return all_charts


def create_safra_charts():
    """Cria gráficos de área plantada por produto a partir do arquivo Safras.xlsx (CONAB)"""
    charts = []

    # --- Produtos principais a exibir (filtra sub-itens e totalizadores) ---
    PRODUTOS_VERAO = [
        'ALGODÃO', 'AMENDOIM TOTAL', 'ARROZ', 'FEIJÃO TOTAL',
        'GERGELIM', 'GIRASSOL', 'MAMONA', 'MILHO TOTAL', 'SOJA', 'SORGO'
    ]
    PRODUTOS_INVERNO = ['AVEIA', 'CANOLA', 'CENTEIO', 'CEVADA', 'TRIGO', 'TRITICALE']

    def parse_area_sheet(filepath):
        """Lê a aba GRA - Área_Brasil e retorna DataFrames estruturados."""
        try:
            df_raw = pd.read_excel(filepath, sheet_name='GRA - Área_Brasil', header=None)
        except Exception as e:
            print(f"⚠️ Erro ao ler Safras.xlsx: {e}")
            return None, None

        # Safras verão: linhas 4 a 35 (índice 4..35 inclusive)
        # Safras inverno: linhas 40 a 45 (índice 40..45 inclusive)
        # Colunas: 0=Produto, 1=24/25(a), 2=25/26 mai(b), 3=25/26 jun(c)
        SAFRA_A = '24/25'
        SAFRA_B = '25/26 (mai)'
        SAFRA_C = '25/26 (jun)'

        def extrair_linhas(linhas_raw, produtos_filtro):
            rows = []
            for _, row in linhas_raw.iterrows():
                prod = str(row.iloc[0]).strip()
                if prod == 'nan' or prod == '':
                    continue
                if prod.upper() not in [p.upper() for p in produtos_filtro]:
                    continue
                for safra_label, col_idx in [(SAFRA_A, 1), (SAFRA_B, 2), (SAFRA_C, 3)]:
                    try:
                        val = float(row.iloc[col_idx])
                        rows.append({'Produto': prod.upper(), 'Safra': safra_label, 'Área (mil ha)': val})
                    except (ValueError, TypeError):
                        pass
            return pd.DataFrame(rows)

        df_verao = extrair_linhas(df_raw.iloc[4:36], PRODUTOS_VERAO)
        df_inverno = extrair_linhas(df_raw.iloc[40:46], PRODUTOS_INVERNO)
        return df_verao, df_inverno

    df_verao, df_inverno = parse_area_sheet('dados_teste/Safras.xlsx')

    if df_verao is None:
        # Fallback com dados de exemplo
        df_verao = pd.DataFrame({
            'Produto': ['SOJA', 'SOJA', 'MILHO TOTAL', 'MILHO TOTAL'],
            'Safra': ['24/25', '25/26 (jun)', '24/25', '25/26 (jun)'],
            'Área (mil ha)': [47346, 48563, 21838, 22582]
        })
        df_inverno = pd.DataFrame({
            'Produto': ['TRIGO', 'TRIGO', 'AVEIA', 'AVEIA'],
            'Safra': ['2025', '25/26 (jun)', '2025', '25/26 (jun)'],
            'Área (mil ha)': [2446, 2117, 546, 536]
        })

    # --- Gráfico 1: Área plantada - Grãos de Verão (barras agrupadas por safra) ---
    if not df_verao.empty:
        safras = df_verao['Safra'].unique().tolist()
        cores = [BLUE, '#5b8dd9', ORANGE]

        fig_verao = go.Figure()
        for i, safra in enumerate(safras):
            df_s = df_verao[df_verao['Safra'] == safra]
            fig_verao.add_trace(go.Bar(
                x=df_s['Produto'],
                y=df_s['Área (mil ha)'],
                name=safra,
                marker=dict(color=cores[i % len(cores)], cornerradius=8),
                text=df_s['Área (mil ha)'].apply(lambda v: f"{v:,.0f}".replace(',', '.')),
                textposition='outside',
                cliponaxis=False,
            ))

        fig_verao.update_layout(
            title='Área Plantada — Grãos de Verão (mil ha)',
            xaxis_title='Produto',
            yaxis_title='Área (mil ha)',
            barmode='group',
            template='plotly_white',
            height=500,
            legend=dict(orientation='h', yanchor='bottom', y=-0.35, xanchor='center', x=0.5),
            xaxis=dict(tickangle=-30),
            yaxis=dict(showgrid=True),
        )
        charts.append({
            'title': 'Área Plantada — Grãos de Verão',
            'chart': json.dumps(fig_verao, cls=plotly.utils.PlotlyJSONEncoder)
        })

    # --- Gráfico 2: Área plantada - Culturas de Inverno ---
    if not df_inverno.empty:
        safras = df_inverno['Safra'].unique().tolist()
        cores = [BLUE, '#5b8dd9', ORANGE]

        fig_inverno = go.Figure()
        for i, safra in enumerate(safras):
            df_s = df_inverno[df_inverno['Safra'] == safra]
            fig_inverno.add_trace(go.Bar(
                x=df_s['Produto'],
                y=df_s['Área (mil ha)'],
                name=safra,
                marker=dict(color=cores[i % len(cores)], cornerradius=8),
                text=df_s['Área (mil ha)'].apply(lambda v: f"{v:,.0f}".replace(',', '.')),
                textposition='outside',
                cliponaxis=False,
            ))

        fig_inverno.update_layout(
            title='Área Plantada — Culturas de Inverno (mil ha)',
            xaxis_title='Produto',
            yaxis_title='Área (mil ha)',
            barmode='group',
            template='plotly_white',
            height=450,
            legend=dict(orientation='h', yanchor='bottom', y=-0.35, xanchor='center', x=0.5),
            yaxis=dict(showgrid=True),
        )
        charts.append({
            'title': 'Área Plantada — Culturas de Inverno',
            'chart': json.dumps(fig_inverno, cls=plotly.utils.PlotlyJSONEncoder)
        })

    # --- Gráfico 3: Variação % entre 25/26 mai e 25/26 jun (top culturas) ---
    try:
        df_raw = pd.read_excel('dados_teste/Safras.xlsx', sheet_name='GRA - Área_Brasil', header=None)
        # col 4 = variação % (c/b) = jun vs mai
        var_rows = []
        for _, row in df_raw.iloc[4:36].iterrows():
            prod = str(row.iloc[0]).strip().upper()
            if prod in PRODUTOS_VERAO:
                try:
                    var_rows.append({'Produto': prod, 'Var % (jun/mai)': float(row.iloc[4])})
                except (ValueError, TypeError):
                    pass
        for _, row in df_raw.iloc[40:46].iterrows():
            prod = str(row.iloc[0]).strip().upper()
            if prod in PRODUTOS_INVERNO:
                try:
                    var_rows.append({'Produto': prod, 'Var % (jun/mai)': float(row.iloc[4])})
                except (ValueError, TypeError):
                    pass

        df_var = pd.DataFrame(var_rows).sort_values('Var % (jun/mai)', ascending=True)

        cores_var = [ORANGE if v < 0 else BLUE for v in df_var['Var % (jun/mai)']]
        fig_var = go.Figure(go.Bar(
            x=df_var['Var % (jun/mai)'],
            y=df_var['Produto'],
            orientation='h',
            marker=dict(color=cores_var, cornerradius=6),
            text=df_var['Var % (jun/mai)'].apply(lambda v: f"{v:+.1f}%"),
            textposition='outside',
            cliponaxis=False,
        ))
        fig_var.update_layout(
            title='Variação da Área Plantada: jun/2026 vs mai/2026 (%)',
            xaxis_title='Variação (%)',
            yaxis_title='',
            template='plotly_white',
            height=480,
            xaxis=dict(showgrid=True, zeroline=True, zerolinewidth=2, zerolinecolor='gray'),
        )
        charts.append({
            'title': 'Variação Área Plantada (jun vs mai)',
            'chart': json.dumps(fig_var, cls=plotly.utils.PlotlyJSONEncoder)
        })
    except Exception as e:
        print(f"⚠️ Gráfico de variação não gerado: {e}")

    return charts


def create_cultivos_charts():
    all_charts = []

    try:
        df_cultivos = pd.read_excel("dados_teste/Preço Cultivos.xlsx", header=12)
        
        df_cultivos = df_cultivos.drop(columns=['Unnamed: 0'], errors='ignore')
        df_cultivos['Produto'] = df_cultivos['Produto'].ffill()
        df_cultivos['Nível de comercialização'] = df_cultivos['Nível de comercialização'].ffill()
        df_cultivos['UF'] = df_cultivos['UF'].ffill()
        
        df_cultivos = df_cultivos.dropna(subset=['Período', 'Preço medio'])
        
        df_cultivos['Período'] = pd.to_datetime(df_cultivos['Período'], format='%m/%Y')
        df_cultivos['Ano'] = df_cultivos['Período'].dt.year
        
        df_cultivos['Preço medio(R$/Kg)'] = (
            df_cultivos['Preço medio']
            .str.replace('R$', '', regex=False)
            .str.replace('.', '', regex=False)
            .str.replace(',', '.', regex=False)
            .str.strip()
        )
        df_cultivos['Preço medio(R$/Kg)'] = pd.to_numeric(df_cultivos['Preço medio(R$/Kg)'], errors='coerce')
        df_cultivos['Produto'] = df_cultivos['Produto'].str.replace(r'\s*\(.*?\)', '', regex=True).str.strip()

    except FileNotFoundError:
        df_cultivos = pd.DataFrame({
            'Produto': ['Soja', 'Soja', 'Milho', 'Milho'],
            'Ano': [2025, 2026, 2025, 2026],
            'Preço medio(R$/Kg)': [2.5, 2.7, 1.8, 1.9]
        })

    df_cultivos = df_cultivos.sort_values(by='Ano')
    df_agregado = df_cultivos.groupby(['Ano', 'Produto']).agg(
        Preco_Medio_Anual=('Preço medio(R$/Kg)', 'mean')
    ).reset_index()

    fig = px.line(
        df_agregado, x='Ano', y='Preco_Medio_Anual', color='Produto',
        markers=True, title='Preço Médio de Cultivos por Ano',
        labels={"Ano": "Ano", "Preco_Medio_Anual": "Preço Médio (R$/Kg)", "Produto": "Produto"}
    )
    fig.update_layout(
        yaxis_tickprefix='R$ ', yaxis_tickformat='.2f',
        xaxis_type='category', template='plotly_white', height=450
    )
    all_charts.append({'title': "Preço dos cultivos por ano", 'chart': json.dumps(fig, cls=plotly.utils.PlotlyJSONEncoder)})

    return all_charts