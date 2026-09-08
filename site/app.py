from flask import Flask, render_template, request, jsonify, flash, redirect, url_for, session, Blueprint
from flask_login import LoginManager, UserMixin, login_user, logout_user, login_required, current_user
import pandas as pd
import plotly.graph_objs as go
import plotly.utils
import json
import os
import sys
from datetime import datetime
from charts import *
import threading
from pathlib import Path
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "jacto-admin-2026-secreto")

# ── Flask-Login ───────────────────────────────────────────────
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = "login"
login_manager.login_message = "Por favor, faça login para acessar esta página."

ADMIN_USERNAME = os.environ.get("ADMIN_USER", "admin")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASS", "jacto2026")

class AdminUser(UserMixin):
    def __init__(self):
        self.id = "admin"

admin_user = AdminUser()

@login_manager.user_loader
def load_user(user_id):
    if user_id == "admin":
        return admin_user
    return None

# ── Log Capture ───────────────────────────────────────────────
class LogCapture:
    def __init__(self):
        self.logs = []
        self.max_logs = 100

    def write(self, message):
        if message.strip():
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            self.logs.append(f"[{timestamp}] {message.strip()}")
            if len(self.logs) > self.max_logs:
                self.logs = self.logs[-self.max_logs:]

    def flush(self):
        pass

    def get_logs(self):
        return self.logs.copy()

log_capture = LogCapture()
original_stdout = sys.stdout
sys.stdout = log_capture

# ── Caminhos ──────────────────────────────────────────────────
def get_project_root():
    current_file = Path(__file__).resolve()
    return current_file.parent.parent

def resolve_path(relative_path):
    project_root = get_project_root()
    if os.path.isabs(relative_path) and os.path.exists(relative_path):
        return Path(relative_path)
    if relative_path.startswith('/') or relative_path.startswith('\\'):
        relative_path = relative_path[1:]
    if relative_path.startswith('../'):
        relative_path = relative_path[3:]
    elif relative_path.startswith('./'):
        relative_path = relative_path[2:]
    return project_root / relative_path

project_root = get_project_root()
os.chdir(project_root)
print(f"📁 Diretório de trabalho definido para: {project_root}")

# ── Planilhas ─────────────────────────────────────────────────
def load_spreadsheets():
    data_folder = resolve_path('dados_teste')
    spreadsheets = {}
    for filename in os.listdir(data_folder):
        if filename.endswith('.xlsx') and not filename.startswith('~'):
            filepath = os.path.join(data_folder, filename)
            try:
                df = pd.read_excel(filepath)
                df.columns = df.columns.str.strip().str.title()
                spreadsheets[filename.replace('.xlsx', '')] = df
            except Exception as e:
                print(f"Erro ao carregar {filename}: {e}")
    return spreadsheets

def create_generic_charts(df, name):
    charts = []
    numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
    categorical_cols = df.select_dtypes(include=['object']).columns.tolist()
    if len(numeric_cols) > 0 and len(categorical_cols) > 0:
        x_col = categorical_cols[0]
        y_col = numeric_cols[0]
        df_grouped = df.groupby(x_col)[y_col].mean().reset_index() if len(df) > 20 else df[[x_col, y_col]].copy()

        bar_chart = go.Figure(data=[go.Bar(
            x=df_grouped[x_col], y=df_grouped[y_col],
            name=y_col, marker_color='rgba(158, 185, 243, 0.7)'
        )])
        bar_chart.update_layout(title=f'{name} - Gráfico de Barras',
                                xaxis_title=x_col, yaxis_title=y_col,
                                template='plotly_white', height=400)
        charts.append({'title': 'Gráfico de Barras',
                       'chart': json.dumps(bar_chart, cls=plotly.utils.PlotlyJSONEncoder)})

        line_chart = go.Figure()
        line_chart.add_trace(go.Scatter(
            x=df_grouped[x_col], y=df_grouped[y_col],
            mode='lines+markers', name=y_col, line=dict(width=3)
        ))
        line_chart.update_layout(title=f'{name} - Gráfico de Linhas',
                                 xaxis_title=x_col, yaxis_title=y_col,
                                 template='plotly_white', height=400)
        charts.append({'title': 'Gráfico de Linhas',
                       'chart': json.dumps(line_chart, cls=plotly.utils.PlotlyJSONEncoder)})
    return charts

# ── Contagem de requests ──────────────────────────────────────
request_counts = {}
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

# ── Rotas públicas ────────────────────────────────────────────
@app.route('/')
def index():
    return redirect(url_for('powerbi_dashboard'))

@app.route('/stats')
def stats():
    return render_template('stats.html', request_counts=request_counts)

@app.route('/powerbi')
def powerbi_dashboard():
    spreadsheets = load_spreadsheets()
    funcs = [principal_tab, pulverizadores_tab, podas_tab, fert_tab, safras_tab, cultivos_tab]
    return render_template('home.html', current_sheet='None', img_name='', funcs=funcs)

@app.route('/powerbi/testes/skeleton')
def teste_esqueleto():
    return render_template('./components/skeleton.html', current_sheet='', img_name='', title='')

@app.route('/powerbi/paginas/<sheet_name>')
def skeleton_waiter(sheet_name):
    sheet_name = sheet_name.lower()
    return render_template('./components/skeleton.html',
                           current_sheet=sheet_name,
                           img_name=f"{sheet_name}.jpg",
                           title=sheet_name.capitalize())

@app.route('/powerbi/principal-data')
def principal_tab():
    view  = request.args.get('view', 'anual')
    month = request.args.get('mes')
    year  = request.args.get('ano')
    charts = create_principal_charts(view=view, year=year, month=month)
    spreadsheets = load_spreadsheets()
    return render_template('charts.html', img_name='principal.jpg',
                           title='Aba Principal', charts=charts, current_sheet='Principal')

@app.route('/powerbi/pulverizadores-data')
def pulverizadores_tab():
    spreadsheets = load_spreadsheets()
    sheet_names = list(spreadsheets.keys())
    charts = create_pulv_charts()
    return render_template('charts.html', img_name='pulv.jpg', title='Pulverizador',
                           charts=charts, sheet_names=sheet_names, current_sheet='Pulverização')

@app.route('/powerbi/podas-data')
def podas_tab():
    spreadsheets = load_spreadsheets()
    sheet_names = list(spreadsheets.keys())
    charts = create_podas_charts()
    return render_template('charts.html', img_name='podas.jpg', title='Poda',
                           charts=charts, sheet_names=sheet_names, current_sheet='Poda')

@app.route('/powerbi/fertilizantes-data')
def fert_tab():
    spreadsheets = load_spreadsheets()
    sheet_names = list(spreadsheets.keys())
    charts = create_fert_charts()
    return render_template('charts.html', img_name='fert.jpg', title='Fertilizante',
                           charts=charts, sheet_names=sheet_names, current_sheet='Fertilizante')

@app.route('/powerbi/safras-data')
def safras_tab():
    try:
        charts = create_safra_charts()
        print(f">>> SAFRA CHARTS COUNT: {len(charts)}")
        for c in charts:
            print(f"    → {c['title']}")
    except Exception as e:
        import traceback
        print("❌ ERRO NA ROTA:")
        print(traceback.format_exc())
        charts = []

    spreadsheets = load_spreadsheets()
    sheet_names = list(spreadsheets.keys())

    return render_template(
        'charts.html',
        img_name='safra.jpg',
        title='Safra',
        charts=charts,
        sheet_names=sheet_names,
        current_sheet='Safra'
    )

@app.route('/powerbi/cultivos-data')
def cultivos_tab():
    spreadsheets = load_spreadsheets()
    sheet_names = list(spreadsheets.keys())
    charts = create_cultivos_charts()
    return render_template('charts.html', charts=charts, sheet_names=sheet_names)

# ── Login / Logout ────────────────────────────────────────────
@app.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('admin.admin_data'))  # 👈 prefixo "admin."
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '').strip()
        if username == ADMIN_USERNAME and password == ADMIN_PASSWORD:
            login_user(admin_user, remember=True)
            next_page = request.args.get('next')
            return redirect(next_page or url_for('admin.admin_data'))  # 👈 prefixo "admin."
        flash("Usuário ou senha incorretos.", "error")
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))

# ═══════════════════════════════════════════════════════════════
# ── Blueprint Admin ────────────────────────────────────────────
# ═══════════════════════════════════════════════════════════════
admin_bp = Blueprint('admin', __name__, url_prefix='/powerbi/admin')

@admin_bp.before_request
def require_login():
    """Protege TODAS as rotas do blueprint automaticamente."""
    if not current_user.is_authenticated:
        return login_manager.unauthorized()

# ── Rotas Admin ───────────────────────────────────────────────
pipeline_running = False
pipeline_status  = "idle"

@admin_bp.route('/data')                        # → /powerbi/admin/data
def admin_data():
    logs = log_capture.get_logs()
    return jsonify({
        'logs': logs,
        'count': len(logs),
        'pipeline_running': pipeline_running,
        'pipeline_status': pipeline_status
    })

@admin_bp.route('/logs', methods=['GET'])       # → /powerbi/admin/logs
def get_logs():
    logs = log_capture.get_logs()
    return jsonify({'logs': logs, 'count': len(logs)})

@admin_bp.route('/logs', methods=['DELETE'])    # → /powerbi/admin/logs
def clear_logs():
    log_capture.logs.clear()
    return jsonify({'status': 'success', 'message': 'Logs cleared'})

@admin_bp.route('/run-pipeline', methods=['POST'])  # → /powerbi/admin/run-pipeline
def run_pipeline():
    global pipeline_running, pipeline_status

    if pipeline_running:
        return jsonify({'status': 'error', 'message': 'Pipeline já está em execução'}), 409

    pipeline_running = True
    pipeline_status  = "running"

    def execute_pipeline():
        global pipeline_running, pipeline_status
        try:
            print("🚀 Iniciando pipeline completa de dados...")
            if str(project_root) not in sys.path:
                sys.path.insert(0, str(project_root))

            import atualiza_arquivos
            import scraper
            import data_eng

            print("📊 Executando scraper...")
            pipeline_status = "scraping"

            print("💰 Atualizando VBP...")
            pipeline_status = "updating_vbp"
            atualiza_arquivos.atualiza_VBP()

            print("🌾 Atualizando Safras...")
            pipeline_status = "updating_safras"
            atualiza_arquivos.atualiza_safras()

            print("💵 Atualizando Dólar...")
            pipeline_status = "updating_dolar"
            atualiza_arquivos.atualiza_dolar()

            print("🧪 Atualizando Agrotóxicos...")
            pipeline_status = "updating_agrotoxicos"
            atualiza_arquivos.atualiza_agrotoxicos()

            print("🏦 Atualizando Crédito Rural...")
            pipeline_status = "updating_credito_rural"
            atualiza_arquivos.atualiza_credito_rural()

            print("✂️ Atualizando Indicadores de Poda...")
            pipeline_status = "updating_poda"
            atualiza_arquivos.atualiza_indicadores_poda()

            print("💰 Atualizando Indicadores de Preços...")
            pipeline_status = "updating_precos"
            atualiza_arquivos.atualiza_indicadores_preços()

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

    thread = threading.Thread(target=execute_pipeline)
    thread.daemon = True
    thread.start()

    return jsonify({'status': 'started', 'message': 'Pipeline iniciada em background'})

@admin_bp.route('/pipeline-status')            # → /powerbi/admin/pipeline-status
def get_pipeline_status():
    return jsonify({'running': pipeline_running, 'status': pipeline_status})

# ── Registro do Blueprint ─────────────────────────────────────
app.register_blueprint(admin_bp)

# ── Inicialização ─────────────────────────────────────────────
if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000, use_reloader=False)