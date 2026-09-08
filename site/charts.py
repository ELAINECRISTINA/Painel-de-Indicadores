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

    all_charts = []

    actual_year = datetime.now().year
    selected_year = int(year) if year is not None and str(year).isdigit() else actual_year
    selected_month = month

    # ─────────────────────────────────────────────
    # 1. VBP
    # ─────────────────────────────────────────────
    try:
        df_raw = pd.read_excel('dados_teste/VBP.xlsx', sheet_name='Sheet1', header=0)
        # Renomeia sem caracteres especiais desde o início
        df_raw.columns = ['Ano', 'Num', 'ValorRS', 'Variacao', 'Mes', 'NumMes', 'VariacaoMes']
        df_raw = df_raw[df_raw['Ano'] != 'Ano']
        df_raw['Ano']        = pd.to_numeric(df_raw['Ano'],        errors='coerce')
        df_raw['ValorRS']    = pd.to_numeric(df_raw['ValorRS'],    errors='coerce')
        df_raw['VariacaoMes']= pd.to_numeric(df_raw['VariacaoMes'],errors='coerce')
        df_raw = df_raw.dropna(subset=['Ano', 'ValorRS'])
        df_raw['Ano'] = df_raw['Ano'].astype(int)
        vbp = df_raw[['Ano', 'Mes', 'ValorRS', 'VariacaoMes']].copy()
        vbp.rename(columns={'VariacaoMes': 'Variacao Mensal', 'Mes': 'Mês', 'ValorRS': 'Variação(R$)'}, inplace=True)
    except Exception as e:
        print(f"⚠️ Erro ao carregar VBP: {e}")
        vbp = pd.DataFrame({'Ano': [], 'Mês': [], 'Variação(R$)': [], 'Variacao Mensal': []})

    vbp_anual = vbp.groupby('Ano')['Variação(R$)'].last().reset_index() if not vbp.empty else pd.DataFrame()
    vbp_atual  = vbp[vbp['Ano'] == selected_year].copy()             if not vbp.empty else pd.DataFrame()

    if not vbp.empty and vbp_atual.empty:
        ultimo_ano = vbp['Ano'].max()
        vbp_atual  = vbp[vbp['Ano'] == ultimo_ano].copy()

    if view == 'mensal' and selected_month is not None and not vbp_atual.empty:
        vbp_atual = vbp_atual[vbp_atual['Mês'] == selected_month]

    fig_vbp = make_subplots(
        rows=1, cols=2,
        subplot_titles=("VBP (Anual)", "Variação Mensal"),
        horizontal_spacing=0.08
    )

    if not vbp_anual.empty:
        y_vals_bi = vbp_anual['Variação(R$)'] / 1e9
        fig_vbp.add_trace(go.Bar(
            x=vbp_anual['Ano'].astype(str).tolist(),
            y=y_vals_bi.tolist(),
            name="VBP Anual",
            marker=dict(color=BLUE, cornerradius=20),
            text=[format_currency_brl(v, 2) + " Bi" for v in y_vals_bi],
            texttemplate='%{text}',
            textposition='outside',
            cliponaxis=False,
        ), row=1, col=1)

    if not vbp_atual.empty:
        fig_vbp.add_trace(go.Scatter(
            x=vbp_atual['Mês'].tolist(),
            y=vbp_atual['Variação Mensal'].tolist(),
            name="Variação Mensal",
            mode='lines+markers+text',
            line=dict(color=ORANGE, width=3),
            marker=dict(size=8, symbol='circle'),
            text=[f'{p:+,.2%}' if pd.notna(p) else '' for p in vbp_atual['Variação Mensal']],
            textposition='top center',
        ), row=1, col=2)

    fig_vbp.update_yaxes(zeroline=True, tickformat='.2%', row=1, col=2)
    fig_vbp.update_layout(
    template="plotly_white",
    yaxis1=dict(
        title="R$ (Bilhões)",
        tickprefix="R$ ",
        ticksuffix=" Bi",
        showgrid=True
    ),
    yaxis2=dict(visible=False),
    xaxis1=dict(title="Ano", type="category"),
    xaxis2=dict(title="Mês"),
    showlegend=False
)
    all_charts.append({
        'title': 'VBP',
        'chart': json.dumps(fig_vbp, cls=plotly.utils.PlotlyJSONEncoder)
    })

    # ─────────────────────────────────────────────
    # 2. Crédito Rural
    # ─────────────────────────────────────────────
    try:
        credito_rural = pd.read_excel('dados_teste/Credito_Rural.xlsx')
        credito_rural.columns = credito_rural.columns.str.strip()

        # DEBUG — remover após confirmar
        print(">>> COLUNAS CREDITO RURAL:", credito_rural.columns.tolist())
        print(credito_rural.head(3).to_string())

        # Detecta coluna de valor
        col_valor = next(
            (c for c in credito_rural.columns if 'valor' in c.lower()),
            None
        )
        col_ano = next(
            (c for c in credito_rural.columns if 'ano' in c.lower()),
            None
        )
        col_mes = next(
            (c for c in credito_rural.columns if 'mes' in c.lower()
                                            or 'mês' in c.lower()),
            None
        )

        print(f">>> col_valor={col_valor} | col_ano={col_ano} | col_mes={col_mes}")

        # Se não encontrou coluna de valor, mostra todas e lança erro claro
        if col_valor is None:
            raise ValueError(
                f"Nenhuma coluna com 'valor' encontrada. "
                f"Colunas disponíveis: {credito_rural.columns.tolist()}"
            )

        # Renomeia para padrão
        rename_map = {}
        if col_valor: rename_map[col_valor] = 'Valor'
        if col_ano:   rename_map[col_ano]   = 'Ano'
        if col_mes:   rename_map[col_mes]   = 'Mês'
        credito_rural = credito_rural.rename(columns=rename_map)

        credito_rural['Valor'] = pd.to_numeric(credito_rural['Valor'], errors='coerce')
        credito_rural = credito_rural.dropna(subset=['Valor'])
        credito_rural['Variação Mensal'] = credito_rural['Valor'].pct_change()

    except FileNotFoundError:
        print("⚠️ Credito_Rural.xlsx não encontrado, usando dados de exemplo.")
        credito_rural = pd.DataFrame({
            'Ano':  [2022] * 12 + [2023] * 12,
            'Mês':  ['Jan','Fev','Mar','Abr','Mai','Jun',
                    'Jul','Ago','Set','Out','Nov','Dez'] * 2,
            'Valor': [4.0e10, 4.2e10, 4.5e10, 4.4e10, 4.8e10, 5.0e10,
                    5.2e10, 5.5e10, 5.8e10, 6.0e10, 6.2e10, 6.5e10,
                    6.8e10, 7.0e10, 7.2e10, 7.5e10, 7.8e10, 8.0e10,
                    8.2e10, 8.5e10, 8.8e10, 9.0e10, 9.2e10, 9.5e10]
        })
        credito_rural['Variação Mensal'] = credito_rural['Valor'].pct_change()

    except Exception as e:
        import traceback
        print(f"❌ Erro ao carregar Credito_Rural: {e}")
        print(traceback.format_exc())
        credito_rural = pd.DataFrame({
            'Ano': [],
            'Mês': [],
            'Valor': [],
            'Variação Mensal': []
        })


    credito_rural_anual = (
        credito_rural.groupby('Ano')['Valor'].last().reset_index()
        if not credito_rural.empty else pd.DataFrame()
    )
    credito_rural_atual = (
        credito_rural[credito_rural['Ano'] == selected_year]
        if 'Ano' in credito_rural.columns else pd.DataFrame()
    )

    if view == 'mensal' and selected_month is not None and not credito_rural_atual.empty:
        credito_rural_atual = credito_rural_atual[credito_rural_atual['Mês'] == selected_month]

    fig_crdt_rural = make_subplots(
        rows=1, cols=2,
        subplot_titles=("Crédito Rural (Anual)", "Variação Mensal"),
        horizontal_spacing=0.08
    )

    if not credito_rural_anual.empty:
        milhoes_crdt_rural = credito_rural_anual['Valor'] / 1e9
        fig_crdt_rural.add_trace(go.Bar(
            x=credito_rural_anual['Ano'],
            y=milhoes_crdt_rural,
            name="Crédito Rural Anual",
            marker=dict(color=BLUE, cornerradius=20),
            text=[f"R$ {v:.2f} Bi".replace('.', ',') for v in milhoes_crdt_rural],
            texttemplate='%{text}',
            textposition='outside',
            cliponaxis=False
        ), row=1, col=1)

    if not credito_rural_atual.empty:
        fig_crdt_rural.add_trace(go.Scatter(
            x=credito_rural_atual['Mês'],
            y=credito_rural_atual['Variação Mensal'],
            name="Variação Mensal",
            mode='lines+markers+text',
            line=dict(color=ORANGE, width=3),
            marker=dict(size=8, symbol='circle'),
            text=[
                f'{p:+,.2%}'.replace('.', ',') if pd.notna(p) else ''
                for p in credito_rural_atual['Variação Mensal']
            ],
            textposition='top center',
        ), row=1, col=2)

    fig_crdt_rural.update_yaxes(tickformat='.2%', row=1, col=2, showgrid=True)
    fig_crdt_rural.update_layout(
    template="plotly_white",
    height=520,
    yaxis1=dict(
        title="R$ (Bilhões)",
        tickprefix="R$ ",
        ticksuffix=" Bi",
        showgrid=True
    ),
    yaxis2=dict(visible=False),
    xaxis1=dict(title="Ano"),
    xaxis2=dict(title="Mês"),
    showlegend=False
)
    all_charts.append({
        'title': 'Crédito Rural',
        'chart': json.dumps(fig_crdt_rural, cls=plotly.utils.PlotlyJSONEncoder)
    })

    # ─────────────────────────────────────────────
    # 3. PIB Agro
    # ─────────────────────────────────────────────
    try:
        pib_agro = pd.read_excel('dados_teste/PIB.xlsx')
        pib_agro.columns = pib_agro.columns.str.strip()

        col_pib     = next((c for c in pib_agro.columns if 'varia' in c.lower()
                                                         or 'pib'   in c.lower()
                                                         or 'valor' in c.lower()), None)
        col_ano_pib = next((c for c in pib_agro.columns if 'ano' in c.lower()), None)

        print(f"✅ Colunas PIB: {pib_agro.columns.tolist()}")

        if col_pib is None or col_ano_pib is None:
            raise ValueError(
                f"Colunas necessárias não encontradas no PIB. Colunas: {pib_agro.columns.tolist()}"
            )

        pib_agro = pib_agro.rename(columns={col_pib: 'Variação(R$ Bilhão)', col_ano_pib: 'Ano'})

        pib_agro['Variação(R$ Bilhão)'] = (
            pib_agro['Variação(R$ Bilhão)']
            .astype(str)
            .str.replace('.', '',   regex=False)
            .str.replace(',', '.',  regex=False)
            .str.replace(r'R\$ ?', '', regex=True)
        )
        pib_agro['Variação(R$ Bilhão)'] = pd.to_numeric(
            pib_agro['Variação(R$ Bilhão)'], errors='coerce'
        )

        pib_agro_anual = pib_agro.groupby('Ano')['Variação(R$ Bilhão)'].mean().reset_index()
        pib_agro_anual['Variação(RBilhão)&#x27;</span>] = pib_agro_anual[<span class="hljs-string">&#x27;Variação(R Bilhão)'] / 1e9

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
            template='plotly_white'
        )
        all_charts.append({
            'title': 'PIB Agro',
            'chart': json.dumps(fig_pib, cls=plotly.utils.PlotlyJSONEncoder)
        })

    except Exception as e:
        import traceback
        print(f"❌ Erro ao carregar PIB: {e}")
        print(traceback.format_exc())

    # ─────────────────────────────────────────────
    # 4. Exportação vs Importação
    # ─────────────────────────────────────────────
    try:
        exportacao = pd.read_excel('dados_teste/Exportacao.xlsx')
        exportacao.columns = exportacao.columns.str.strip()

        print(f"✅ Colunas Exportacao: {exportacao.columns.tolist()}")

        col_ano_exp   = next((c for c in exportacao.columns if 'ano'   in c.lower()), None)
        col_fluxo     = next((c for c in exportacao.columns if 'fluxo' in c.lower()), None)
        col_valor_exp = next((c for c in exportacao.columns if 'valor' in c.lower()), None)

        if not all([col_ano_exp, col_fluxo, col_valor_exp]):
            raise ValueError(
                f"Colunas necessárias não encontradas. Colunas: {exportacao.columns.tolist()}"
            )

        exportacao = exportacao.rename(columns={
            col_ano_exp:   'Ano',
            col_fluxo:     'Fluxo',
            col_valor_exp: 'Valor (US$)'
        })

        exportacao_anual = exportacao.groupby(['Ano', 'Fluxo'])['Valor (US$)'].sum().reset_index()
        exportacao_pivot = (
            exportacao_anual
            .pivot(index='Ano', columns='Fluxo', values='Valor (US$)')
            .fillna(0)
            .reset_index()
        )

    except FileNotFoundError:
        print("⚠️ Exportacao.xlsx não encontrado, usando dados de exemplo.")
        exportacao_pivot = pd.DataFrame({
            'Ano':        [2022,    2023,    2024],
            'Exportação': [2.5e9,   3.2e9,   3.8e9],
            'Importação': [1.8e9,   2.1e9,   2.3e9]
        })
    except Exception as e:
        import traceback
        print(f"❌ Erro ao carregar Exportacao: {e}")
        print(traceback.format_exc())
        exportacao_pivot = pd.DataFrame(columns=['Ano'])

    fig_exportacao = go.Figure()

    if 'Exportação' in exportacao_pivot.columns:
        valores_exp = exportacao_pivot['Exportação'] / 1e6
        fig_exportacao.add_trace(go.Scatter(
            x=exportacao_pivot['Ano'],
            y=valores_exp,
            name="Exportação",
            mode='lines+markers+text',
            line=dict(color=BLUE, width=3),
            marker=dict(size=8, symbol='circle'),
            text=[f"US$ {v:.0f} Mi" for v in valores_exp],
            textposition='top center',
        ))

    if 'Importação' in exportacao_pivot.columns:
        valores_imp = exportacao_pivot['Importação'] / 1e6
        fig_exportacao.add_trace(go.Scatter(
            x=exportacao_pivot['Ano'],
            y=valores_imp,
            name="Importação",
            mode='lines+markers+text',
            line=dict(color=ORANGE, width=3),
            marker=dict(size=8, symbol='triangle-up'),
            text=[f"US$ {v:.0f} Mi" for v in valores_imp],
            textposition='bottom center',
        ))

    fig_exportacao.update_layout(
    template="plotly_white",
    height=520,
    yaxis=dict(
        title="US$ (Milhões)",
        tickprefix="US$ ",
        ticksuffix=" Mi",
        showgrid=True
    ),
    xaxis=dict(title="Ano", type="category"),
    showlegend=True,
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=1.02,
        xanchor="center",
        x=0.5
    )
)
    all_charts.append({
        'title': 'Importação vs Exportação de Máquinas Agrícolas',
        'chart': json.dumps(fig_exportacao, cls=plotly.utils.PlotlyJSONEncoder)
    })

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
            text=[f"R$ {v:,.2f}".replace('.',',') for v in indicadores_linhas_anual['Preço']],
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
        dados_pulverizadores = pd.read_excel('dados_apoio/Preços_pulverizadores.xlsx')
    except FileNotFoundError:
        dados_pulverizadores = pd.DataFrame({'Modelo': ['Modelo A', 'Modelo B', 'Modelo C'], 'Mediana': [500000, 750000, 600000]})

    fig_bar = px.bar(
        dados_pulverizadores,
        x='Mediana',
        y='Modelo',
        orientation='h',
  
    )
    # Atualização de traços e layout para melhor visualização
    fig_bar.update_traces(marker_color=BLUE, texttemplate='R$ %{x:.0f}', textposition='outside')
    fig_bar.update_layout(xaxis_title='Preço Mediano (R$)', yaxis_title='Modelo', template='plotly_white')

    chart_json_dados_pulverizadores = json.dumps(fig_bar, cls=plotly.utils.PlotlyJSONEncoder)
    all_charts.append({
        'title': 'Preço por Modelo',
        'chart': chart_json_dados_pulverizadores
    })

    pib_agro = pd.read_excel('dados_teste/PIB.xlsx')
    pib_agro['Variação(R$ Bilhão)'] = pib_agro['Variação(R$ Bilhão)'].astype(str).str.replace('.', '', regex=False)
    # 2. Substitui vírgula (decimal) por ponto
    pib_agro['Variação(R$ Bilhão)'] = pib_agro['Variação(R$ Bilhão)'].str.replace(',', '.', regex=False)
    # 3. Remove "R$" se houver
    pib_agro['Variação(R$ Bilhão)'] = pib_agro['Variação(R$ Bilhão)'].str.replace(r'R\$ ?', '', regex=True)
    # 4. Converte para numérico (erros viram NaN)
    pib_agro['Variação(R$ Bilhão)'] = pd.to_numeric(pib_agro['Variação(R$ Bilhão)'], errors='coerce')
    pib_agro_anual = pib_agro.groupby('Ano')['Variação(R$ Bilhão)'].mean().reset_index()
    pib_agro_anual['Variação(R$ Bilhão)'] = pib_agro_anual['Variação(R$ Bilhão)'] / 1e9

    vendas_jacto = pd.read_excel('dados_teste/Vendas_jacto.xlsx')
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
    )
    fig_faturamento.update_traces(
        marker=dict(color=BLUE,cornerradius=20),
        texttemplate='$ %{y:.2f}',
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

    # ─────────────────────────────────────────────
    # CARREGAMENTO E DETECÇÃO DE COLUNAS
    # ─────────────────────────────────────────────
    podas = None

    try:
        podas_raw = pd.read_excel('dados_teste/Indicadores_Poda.xlsx')
        podas_raw.columns = podas_raw.columns.str.strip()

        # DEBUG — mostra colunas reais do arquivo
        print("=" * 60)
        print(">>> COLUNAS REAIS DO ARQUIVO:")
        for i, c in enumerate(podas_raw.columns.tolist()):
            print(f"    [{i}] '{c}'")
        print(">>> PRIMEIRAS LINHAS:")
        print(podas_raw.head(3).to_string())
        print("=" * 60)

        # Detecta coluna de ano
        col_year = next(
            (c for c in podas_raw.columns
             if 'year' in c.lower() or 'ano' in c.lower()),
            None
        )
        # Detecta coluna de área
        col_area = next(
            (c for c in podas_raw.columns
             if 'area' in c.lower() or 'área' in c.lower()),
            None
        )
        # Detecta coluna de custo
        col_custo = next(
            (c for c in podas_raw.columns
             if 'custo' in c.lower()),
            None
        )

        print(f">>> DETECTADO: col_year='{col_year}' | col_area='{col_area}' | col_custo='{col_custo}'")

        # Se não detectou, mostra todas as colunas e para
        if col_year is None or col_area is None or col_custo is None:
            print("❌ COLUNAS NÃO DETECTADAS AUTOMATICAMENTE!")
            print("   Colunas disponíveis:", podas_raw.columns.tolist())
            raise ValueError(
                f"Colunas não encontradas.\n"
                f"  Ano   → '{col_year}'\n"
                f"  Área  → '{col_area}'\n"
                f"  Custo → '{col_custo}'\n"
                f"  Disponíveis: {podas_raw.columns.tolist()}"
            )

        # Renomeia e converte
        podas = podas_raw.rename(columns={
            col_year:  'Year',
            col_area:  'Área Colhida',
            col_custo: 'CUSTO POR HA'
        })

        podas['Year']         = pd.to_numeric(podas['Year'],         errors='coerce')
        podas['Área Colhida'] = pd.to_numeric(podas['Área Colhida'], errors='coerce')
        podas['CUSTO POR HA'] = pd.to_numeric(podas['CUSTO POR HA'], errors='coerce')
        podas = podas.dropna(subset=['Year', 'Área Colhida', 'CUSTO POR HA'])
        podas['Year'] = podas['Year'].astype(int)

        print(f">>> APÓS LIMPEZA: {len(podas)} linhas válidas")
        print(podas.head(3).to_string())

    except FileNotFoundError:
        print("⚠️ Indicadores_Poda.xlsx não encontrado — usando dados de exemplo.")
        podas = pd.DataFrame({
            'Year':         [2021, 2022, 2023],
            'Área Colhida': [1000.0, 1500.0, 1200.0],
            'CUSTO POR HA': [500.0,  550.0,  600.0]
        })

    except Exception as e:
        import traceback
        print("❌ ERRO AO CARREGAR PODAS:")
        print(traceback.format_exc())
        podas = pd.DataFrame({
            'Year':         [2021, 2022, 2023],
            'Área Colhida': [1000.0, 1500.0, 1200.0],
            'CUSTO POR HA': [500.0,  550.0,  600.0]
        })

    # ─────────────────────────────────────────────
    # AGRUPAMENTOS — garantidos após try/except
    # ─────────────────────────────────────────────
    print(f">>> COLUNAS DO PODAS FINAL: {podas.columns.tolist()}")
    print(f">>> VAZIO? {podas.empty}")

    if podas.empty:
        print("⚠️ DataFrame vazio — sem gráficos.")
        return all_charts

    podas_anual      = podas.groupby('Year')['Área Colhida'].mean().reset_index()
    podas_anual_mean = podas.groupby('Year')['CUSTO POR HA'].mean().reset_index()

    # ─────────────────────────────────────────────
    # GRÁFICO — Área Colhida + Custo por HA
    # ─────────────────────────────────────────────
    fig_podas = make_subplots(
        rows=1, cols=1,
        specs=[[{"secondary_y": True}]]
    )

    fig_podas.add_trace(
        go.Bar(
            x=podas_anual['Year'],
            y=podas_anual['Área Colhida'],
            name="Área Colhida (ha)",
            marker=dict(color=BLUE, cornerradius=20),
            text=[f"{v:,.0f}" for v in podas_anual['Área Colhida']],
            texttemplate='%{text}',
            textposition='outside',
            cliponaxis=False
        ),
        row=1, col=1,
        secondary_y=False
    )

    fig_podas.add_trace(
        go.Scatter(
            x=podas_anual_mean['Year'],
            y=podas_anual_mean['CUSTO POR HA'],
            name="Custo por HA (R$)",
            mode='lines+markers',
            line=dict(color=ORANGE, width=3),
            marker=dict(size=8, symbol='circle'),
            text=[f"R$ {v:,.2f}" for v in podas_anual_mean['CUSTO POR HA']],
            texttemplate='%{text}',
            textposition='top center',
            cliponaxis=False
        ),
        row=1, col=1,
        secondary_y=True
    )

    fig_podas.update_layout(
        template="plotly_white",
        height=520,
        xaxis=dict(title="Ano", type='category'),
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.2,
            xanchor="center",
            x=0.5
        )
    )

    fig_podas.update_yaxes(secondary_y=False, tickformat=',.0f')
    fig_podas.update_yaxes(tickprefix='R$ ', secondary_y=True, showgrid=False)

    all_charts.append({
        'title': 'Indicadores de Poda',
        'chart': json.dumps(fig_podas, cls=plotly.utils.PlotlyJSONEncoder)
    })

    
    try:
        pib_agro = pd.read_excel('dados_teste/PIB.xlsx')
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
    
    pib_agro = pd.read_excel('dados_teste/PIB.xlsx')
    pib_agro['Variação(R$ Bilhão)'] = pib_agro['Variação(R$ Bilhão)'].astype(str).str.replace('.', '', regex=False)
    # 2. Substitui vírgula (decimal) por ponto
    pib_agro['Variação(R$ Bilhão)'] = pib_agro['Variação(R$ Bilhão)'].str.replace(',', '.', regex=False)
    # 3. Remove "R$" se houver
    pib_agro['Variação(R$ Bilhão)'] = pib_agro['Variação(R$ Bilhão)'].str.replace(r'R\$ ?', '', regex=True)
    # 4. Converte para numérico (erros viram NaN)
    pib_agro['Variação(R$ Bilhão)'] = pd.to_numeric(pib_agro['Variação(R$ Bilhão)'], errors='coerce')
    pib_agro_anual = pib_agro.groupby('Ano')['Variação(R$ Bilhão)'].mean().reset_index()
    pib_agro_anual['Variação(R$ Bilhão)'] = pib_agro_anual['Variação(R$ Bilhão)'] / 1e9

    vendas_jacto = pd.read_excel('dados_teste/Vendas_jacto.xlsx')
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
        agrotoxicos = pd.read_excel('dados_teste/Agrotóxicos.xlsx')
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

    pib_agro = pd.read_excel('dados_teste/PIB.xlsx')
    pib_agro['Variação(R$ Bilhão)'] = pib_agro['Variação(R$ Bilhão)'].astype(str).str.replace('.', '', regex=False)
    # 2. Substitui vírgula (decimal) por ponto
    pib_agro['Variação(R$ Bilhão)'] = pib_agro['Variação(R$ Bilhão)'].str.replace(',', '.', regex=False)
    # 3. Remove "R$" se houver
    pib_agro['Variação(R$ Bilhão)'] = pib_agro['Variação(R$ Bilhão)'].str.replace(r'R\$ ?', '', regex=True)
    # 4. Converte para numérico (erros viram NaN)
    pib_agro['Variação(R$ Bilhão)'] = pd.to_numeric(pib_agro['Variação(R$ Bilhão)'], errors='coerce')
    pib_agro_anual = pib_agro.groupby('Ano')['Variação(R$ Bilhão)'].mean().reset_index()
    pib_agro_anual['Variação(R$ Bilhão)'] = pib_agro_anual['Variação(R$ Bilhão)'] / 1e9

    vendas_jacto = pd.read_excel('dados_teste/Vendas_jacto.xlsx')
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

    return all_charts

def create_safra_charts():
    charts = []

    try:
        # Lê pulando as 3 primeiras linhas de cabeçalho mesclado
        df_raw = pd.read_excel(
            'dados_teste/Safras.xlsx',
            header=None  # lê tudo sem cabeçalho automático
        )

        print(">>> SHAPE RAW:", df_raw.shape)
        print(df_raw.head(6).to_string())

        # Detecta a linha onde começa o cabeçalho real (contém 'PRODUTO')
        header_row = None
        for i, row in df_raw.iterrows():
            if row.astype(str).str.contains('PRODUTO', case=False, na=False).any():
                header_row = i
                break

        if header_row is None:
            raise ValueError("Linha de cabeçalho com 'PRODUTO' não encontrada.")

        print(f">>> HEADER ROW: {header_row}")

        # Detecta linha onde começa os dados reais (primeira após cabeçalho com valor numérico)
        data_start = None
        for i in range(header_row + 1, len(df_raw)):
            row = df_raw.iloc[i]
            # Linha de dados tem pelo menos 2 valores numéricos
            nums = pd.to_numeric(row, errors='coerce').notna().sum()
            if nums >= 2:
                data_start = i
                break

        if data_start is None:
            raise ValueError("Linha de início dos dados não encontrada.")

        print(f">>> DATA START ROW: {data_start}")

        # Monta colunas manualmente baseado nas linhas de cabeçalho
        # Estrutura CONAB: PRODUTO | Safra anterior | jul/2026 | ago/2026 | Var% | ...
        df = df_raw.iloc[data_start:].copy()
        df = df.reset_index(drop=True)

        # Renomeia colunas pela posição (mais seguro que pelo nome)
        col_names = df_raw.iloc[header_row:data_start].ffill(axis=0).iloc[-1].tolist()
        print(f">>> NOMES DAS COLUNAS DETECTADOS: {col_names}")

        df.columns = range(len(df.columns))

        # Coluna 0 = PRODUTO, coluna 1 = safra anterior, col 2 = jul, col 3 = ago
        df = df.rename(columns={
            0: 'Produto',
            1: 'Safra_Anterior',
            2: 'Jul_2026',
            3: 'Ago_2026',
            4: 'Var_Percentual',
        })

        # Mantém só colunas relevantes
        df = df[['Produto', 'Safra_Anterior', 'Jul_2026', 'Ago_2026', 'Var_Percentual']].copy()

        # Converte para string e limpa produto
        df['Produto'] = df['Produto'].astype(str).str.strip()

        # Remove linhas inválidas (totais, fontes, NaN, cabeçalhos repetidos)
        linhas_invalidas = [
            'nan', 'PRODUTO', 'SUBTOTAL', 'BRASIL',
            'Fonte:', 'Nota:', 'CULTURAS DE INVERNO', ''
        ]
        df = df[~df['Produto'].isin(linhas_invalidas)]
        df = df[~df['Produto'].str.startswith('Fonte', na=True)]
        df = df[~df['Produto'].str.startswith('Nota', na=True)]

        # Converte valores numéricos
        for col in ['Safra_Anterior', 'Jul_2026', 'Ago_2026', 'Var_Percentual']:
            df[col] = pd.to_numeric(df[col], errors='coerce')

        df = df.dropna(subset=['Ago_2026'])
        df = df.reset_index(drop=True)

        print(f">>> PRODUTOS ENCONTRADOS ({len(df)}):")
        print(df[['Produto', 'Ago_2026']].to_string())

    except FileNotFoundError:
        print("⚠️ Safras.xlsx não encontrado — usando dados de exemplo.")
        df = pd.DataFrame({
            'Produto':        ['SOJA', 'MILHO TOTAL', 'ALGODÃO', 'TRIGO'],
            'Safra_Anterior': [45000, 116000, 2085, 10000],
            'Jul_2026':       [46000, 118000, 2018, 10500],
            'Ago_2026':       [46500, 119000, 2018, 10800],
            'Var_Percentual': [1.1, 0.8, 0.0, 2.9],
        })

    except Exception as e:
        import traceback
        print(f"❌ Erro ao carregar Safras: {e}")
        print(traceback.format_exc())
        return charts

    if df.empty:
        print("⚠️ DataFrame vazio após limpeza.")
        return charts

    # ── Gráfico 1: Área por Produto (Ago/2026) ──────────────────────────
    df_sorted = df.sort_values('Ago_2026', ascending=True).tail(15)  # top 15

    fig_area = go.Figure()
    fig_area.add_trace(go.Bar(
        x=df_sorted['Ago_2026'],
        y=df_sorted['Produto'],
        orientation='h',
        marker=dict(color=BLUE, cornerradius=10),
        text=[f"{v:,.1f} mil ha" for v in df_sorted['Ago_2026']],
        textposition='outside',
        cliponaxis=False,
    ))
    fig_area.update_layout(
        title='Área Plantada por Produto — Ago/2026 (mil ha)',
        xaxis_title='Área (mil hectares)',
        yaxis_title='Produto',
        template='plotly_white',
        height=520,
        margin=dict(l=180),
    )
    charts.append({
        'title': 'Área Plantada (Ago/2026)',
        'chart': json.dumps(fig_area, cls=plotly.utils.PlotlyJSONEncoder)
    })

    # ── Gráfico 2: Variação % por Produto ───────────────────────────────
    df_var = df[df['Var_Percentual'].notna()].copy()
    df_var = df_var.sort_values('Var_Percentual', ascending=True)

    colors_var = [BLUE if v >= 0 else ORANGE for v in df_var['Var_Percentual']]

    fig_var = go.Figure()
    fig_var.add_trace(go.Bar(
        x=df_var['Var_Percentual'],
        y=df_var['Produto'],
        orientation='h',
        marker=dict(color=colors_var, cornerradius=10),
        text=[f"{v:+.1f}%" for v in df_var['Var_Percentual']],
        textposition='outside',
        cliponaxis=False,
    ))
    fig_var.update_layout(
        title='Variação % — jul/2026 vs ago/2026',
        xaxis_title='Variação (%)',
        yaxis_title='Produto',
        template='plotly_white',
        height=520,
        margin=dict(l=180),
    )
    charts.append({
        'title': 'Variação % por Produto',
        'chart': json.dumps(fig_var, cls=plotly.utils.PlotlyJSONEncoder)
    })

    # ── Gráfico 3: Comparativo Safra Anterior vs Ago/2026 ───────────────
    df_comp = df.dropna(subset=['Safra_Anterior', 'Ago_2026']).copy()
    df_comp = df_comp.sort_values('Ago_2026', ascending=False).head(10)

    fig_comp = go.Figure()
    fig_comp.add_trace(go.Bar(
        name='Safra Anterior',
        x=df_comp['Produto'],
        y=df_comp['Safra_Anterior'],
        marker=dict(color=ORANGE, cornerradius=10),
        text=[f"{v:,.0f}" for v in df_comp['Safra_Anterior']],
        textposition='outside',
        cliponaxis=False,
    ))
    fig_comp.add_trace(go.Bar(
        name='Ago/2026',
        x=df_comp['Produto'],
        y=df_comp['Ago_2026'],
        marker=dict(color=BLUE, cornerradius=10),
        text=[f"{v:,.0f}" for v in df_comp['Ago_2026']],
        textposition='outside',
        cliponaxis=False,
    ))
    fig_comp.update_layout(
        title='Comparativo: Safra Anterior vs Ago/2026 (Top 10)',
        xaxis_title='Produto',
        yaxis_title='Área (mil ha)',
        template='plotly_white',
        barmode='group',
        height=520,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="center",
            x=0.5
        ),
    )
    charts.append({
        'title': 'Comparativo Safras',
        'chart': json.dumps(fig_comp, cls=plotly.utils.PlotlyJSONEncoder)
    })

    print(f">>> TOTAL CHARTS GERADOS: {len(charts)}")
    return charts



def create_cultivos_charts():
    all_charts = []

    try:
        agrotoxics = pd.read_excel('dados_teste/Agrotóxicos.xlsx')
        agrotoxics.columns = agrotoxics.columns.str.strip()
    except FileNotFoundError:
        agrotoxics = pd.DataFrame({
            'Sub-Grupo': ['Herbicida', 'Fungicida', 'Herbicida'],
            'Ano': [2022, 2022, 2023],
            'Preço': [150.00, 200.00, 160.00]
        })

    agrotoxics_sub_groups = agrotoxics.groupby(['Sub-Grupo', 'Ano'])['Preço'].mean().reset_index()
    for sub_group in agrotoxics_sub_groups['Sub-Grupo'].unique():
        df = agrotoxics_sub_groups.loc[agrotoxics_sub_groups['Sub-Grupo'] == sub_group]
        fig = px.bar(df, x='Ano', y='Preço', title=f'Preço {sub_group}')
        fig.update_traces(
            marker=dict(color=BLUE, cornerradius=20),
            texttemplate='R$ %{y:,.2f}',
            textposition='outside'
        )
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

    # ── Gráfico de Preço Médio de Cultivos ──────────────────────────────
    try:
        df_cultivos = pd.read_excel("dados_teste/Preço Cultivos.xlsx")
        df_cultivos.columns = df_cultivos.columns.str.strip()

        # DEBUG
        print("=" * 60)
        print(">>> COLUNAS REAIS Preço Cultivos.xlsx:")
        for i, c in enumerate(df_cultivos.columns.tolist()):
            print(f"    [{i}] '{c}'")
        print(df_cultivos.head(3).to_string())
        print("=" * 60)

        # Detecta colunas dinamicamente
        col_ano = next(
            (c for c in df_cultivos.columns
             if 'ano' in c.lower() or 'year' in c.lower() or 'data' in c.lower()),
            None
        )
        col_produto = next(
            (c for c in df_cultivos.columns
             if any(x in c.lower() for x in ['produto', 'cultivo', 'cultura', 'item'])),
            None
        )
        col_preco = next(
            (c for c in df_cultivos.columns
             if any(x in c.lower() for x in ['preço', 'preco', 'valor', 'price'])),
            None
        )

        print(f">>> col_ano='{col_ano}' | col_produto='{col_produto}' | col_preco='{col_preco}'")

        if not all([col_ano, col_produto, col_preco]):
            raise ValueError(
                f"Colunas não detectadas.\n"
                f"  Ano     → '{col_ano}'\n"
                f"  Produto → '{col_produto}'\n"
                f"  Preço   → '{col_preco}'\n"
                f"  Disponíveis: {df_cultivos.columns.tolist()}"
            )

        df_cultivos = df_cultivos.rename(columns={
            col_ano:     'Ano',
            col_produto: 'Produto',
            col_preco:   'Preço medio(R$/Kg)',
        })

        df_cultivos['Preço medio(R$/Kg)'] = pd.to_numeric(
            df_cultivos['Preço medio(R$/Kg)'], errors='coerce'
        )
        df_cultivos['Ano'] = pd.to_numeric(df_cultivos['Ano'], errors='coerce')
        df_cultivos = df_cultivos.dropna(subset=['Ano', 'Preço medio(R$/Kg)'])
        df_cultivos['Ano'] = df_cultivos['Ano'].astype(int)

    except FileNotFoundError:
        print("⚠️ Preço Cultivos.xlsx não encontrado — usando dados de exemplo.")
        df_cultivos = pd.DataFrame({
            'Produto': ['Soja', 'Soja', 'Milho', 'Milho'],
            'Ano':     [2022, 2023, 2022, 2023],
            'Preço medio(R$/Kg)': [2.5, 2.7, 1.8, 1.9]
        })

    except Exception as e:
        import traceback
        print(f"❌ Erro ao carregar Preço Cultivos: {e}")
        print(traceback.format_exc())
        df_cultivos = pd.DataFrame(columns=['Produto', 'Ano', 'Preço medio(R$/Kg)'])

    if not df_cultivos.empty:
        df_cultivos = df_cultivos.sort_values(by='Ano')  # ✅ seguro agora

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
            'title': "Preço dos cultivos por ano",
            'chart': json.dumps(fig, cls=plotly.utils.PlotlyJSONEncoder)
        })
    else:
        print("⚠️ df_cultivos vazio — gráfico de cultivos não gerado.")

    return all_charts