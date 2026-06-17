from flask import Flask, render_template, request, jsonify, flash, redirect, url_for
import random 
import pandas as pd
import plotly.graph_objs as go
import plotly.utils
import json
import os
import base64
import io
import sys
from datetime import datetime
from charts import *
import subprocess
import threading
from pathlib import Path


app = Flask(__name__)

def load_spreadsheets():
    """Carrega todas as planilhas da pasta dados_teste"""
    data_folder = resolve_path('dados_teste')  # Remove a barra inicial
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

class LogCapture:
    def __init__(self):
        self.logs = []
        self.max_logs = 100  # Máximo de logs para manter na memória
    
    def write(self, message):
        if message.strip():  # Só adiciona se não for string vazia
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            log_entry = f"[{timestamp}] {message.strip()}"
            self.logs.append(log_entry)
            # Mantém apenas os últimos logs
            if len(self.logs) > self.max_logs:
                self.logs = self.logs[-self.max_logs:]
    
    def flush(self):
        pass
    
    def get_logs(self):
        return self.logs.copy()

# Instancia o capturador de logs
log_capture = LogCapture()

# Redireciona stdout para capturar prints
original_stdout = sys.stdout
sys.stdout = log_capture

def get_project_root():
    """Retorna o caminho absoluto para a raiz do projeto (pasta automate)"""
    # Caminho do arquivo atual (app.py)
    current_file = Path(__file__).resolve()
    # Sobe dois níveis: site/app.py -> automate/
    project_root = current_file.parent.parent
    return project_root

# Função para resolver caminhos relativos para absolutos
def resolve_path(relative_path):
    """Converte um caminho relativo para absoluto baseado na raiz do projeto"""
    project_root = get_project_root()
    
    # Se já é um caminho absoluto válido, retorna como está
    if os.path.isabs(relative_path) and os.path.exists(relative_path):
        return Path(relative_path)
    
    # Trata caminhos que começam com / ou \ como relativos à raiz do projeto
    if relative_path.startswith('/') or relative_path.startswith('\\'):
        relative_path = relative_path[1:]  # Remove a barra inicial
    
    # Remove pontos iniciais se houver
    if relative_path.startswith('../'):
        relative_path = relative_path[3:]  # Remove '../'
    elif relative_path.startswith('./'):
        relative_path = relative_path[2:]  # Remove './'
    
    # Junta com a raiz do projeto
    return project_root / relative_path

# Configura o diretório de trabalho para a raiz do projeto
project_root = get_project_root()
os.chdir(project_root)
print(f"📁 Diretório de trabalho definido para: {project_root}")

request_counts={}
ALLOWED_ENDPOINTS = {'index', 'powerbi_dashboard', 'powerbi_sheet'}

@app.route('/')
def index():
    return redirect(url_for('powerbi_dashboard'))



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

@app.route('/powerbi/admin')
def admin():
    """Página admin para visualizar logs do terminal"""
    logs = log_capture.get_logs()
    return render_template('admin.html', logs=logs, log_count=len(logs))

@app.route('/powerbi/admin/logs')
def get_logs():
    """API endpoint para obter logs em JSON (para updates em tempo real)"""
    logs = log_capture.get_logs()
    return jsonify({'logs': logs, 'count': len(logs)})

@app.route('/powerbi/admin/logs', methods=['DELETE'])
def clear_logs():
    """Endpoint para limpar logs"""
    log_capture.logs.clear()
    return jsonify({'status': 'success', 'message': 'Logs cleared'})

# Variável global para controlar se a pipeline está rodando
pipeline_running = False
pipeline_status = "idle"

@app.route('/powerbi/admin/run-pipeline', methods=['POST'])
def run_pipeline():
    """Executa a pipeline completa de dados"""
    global pipeline_running, pipeline_status
    
    if pipeline_running:
        return jsonify({
            'status': 'error', 
            'message': 'Pipeline já está em execução'
        }), 409
    
    pipeline_running = True
    pipeline_status = "running"
    
    def execute_pipeline():
        global pipeline_running, pipeline_status
        try:
            print("🚀 Iniciando pipeline completa de dados...")

            parent_dir = project_root
            if str(parent_dir) not in sys.path:
                sys.path.insert(0, str(parent_dir))
            # Importa as funções de atualização
            import atualiza_arquivos
            import scraper
            import data_eng
            
            # 1. Executa o scraper principal
            print("📊 Executando scraper...")
            pipeline_status = "scraping"
            
            # Aqui você pode adicionar chamadas específicas do scraper
            # scraper.main()  # Se houver uma função main no scraper
            
            # 2. Atualiza VBP
            print("💰 Atualizando VBP...")
            pipeline_status = "updating_vbp"
            atualiza_arquivos.atualiza_VBP()
            
            # 3. Atualiza Safras
            print("🌾 Atualizando Safras...")
            pipeline_status = "updating_safras"
            atualiza_arquivos.atualiza_safras()
            
            # 4. Atualiza Dólar
            print("💵 Atualizando Dólar...")
            pipeline_status = "updating_dolar"
            atualiza_arquivos.atualiza_dolar()
            
            # 5. Atualiza Agrotóxicos
            print("🧪 Atualizando Agrotóxicos...")
            pipeline_status = "updating_agrotoxicos"
            atualiza_arquivos.atualiza_agrotoxicos()
            
            # 6. Atualiza Crédito Rural
            print("🏦 Atualizando Crédito Rural...")
            pipeline_status = "updating_credito_rural"
            atualiza_arquivos.atualiza_credito_rural()
            
            # 7. Atualiza Indicadores de Poda
            print("✂️ Atualizando Indicadores de Poda...")
            pipeline_status = "updating_poda"
            atualiza_arquivos.atualiza_indicadores_poda()
            
            # 8. Atualiza Preços
            print("💰 Atualizando Indicadores de Preços...")
            pipeline_status = "updating_precos"
            atualiza_arquivos.atualiza_indicadores_preços()
            
            # 9. Atualiza PIB Agro
            print("📈 Atualizando PIB Agro...")
            pipeline_status = "updating_pib"
            atualiza_arquivos.atualiza_pib_agro()
            
            print("✅ Pipeline completa executada com sucesso!")
            pipeline_status = "completed"
            
        except Exception as e:
            print(f"❌ Erro na pipeline: {str(e)}")
            pipeline_status = f"error: {str(e)}"
        finally:
            pipeline_running = False
    
    # Executa em uma thread separada para não bloquear a resposta
    thread = threading.Thread(target=execute_pipeline)
    thread.daemon = True
    thread.start()
    
    return jsonify({
        'status': 'started', 
        'message': 'Pipeline iniciada em background'
    })

@app.route('/powerbi/admin/pipeline-status')
def get_pipeline_status():
    """Retorna o status atual da pipeline"""
    return jsonify({
        'running': pipeline_running,
        'status': pipeline_status
    })

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)