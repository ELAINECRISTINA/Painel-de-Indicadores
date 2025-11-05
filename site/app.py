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