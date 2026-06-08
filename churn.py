import pandas as pd
import os
import streamlit as st
import plotly.express as px
from datetime import datetime, date
import io
import json
try:
    from streamlit_gsheets import GSheetsConnection
    _GSHEETS_OK = True
except Exception:
    _GSHEETS_OK = False
    class GSheetsConnection:  # placeholder para nao quebrar o import
        pass

class _DummyConn:
    """Conexao falsa usada quando o Google Sheets nao esta instalado/configurado (modo local)."""
    def read(self, *args, **kwargs):
        return pd.DataFrame()

def _get_gsheets_conn():
    if _GSHEETS_OK:
        try:
            return st.connection("gsheets", type=GSheetsConnection)
        except Exception:
            return _DummyConn()
    return _DummyConn()
# (English-only panel; reads the Global taxonomy columns directly from the data files)

# --- 1. Configurações, Caminhos e API ---
data_dir = os.path.dirname(os.path.abspath(__file__))
# ATUALIZAÇÃO DE ANOS: 2025 e 2026
file_previous_year = 'churn_2025.xlsx'
file_current_year = 'churn_2026.xlsx'
file_backlog_churn = 'backlog_churn.xlsx'
file_active_base = 'base_ativa_clientes.xlsx'
file_creation_previous = 'criação_2025.xlsx'
file_creation_current = 'criação_2026.xlsx'
otl_projections_file = 'otl_churn.xlsx'
projecao_file = os.path.join(data_dir, 'projecao_churn.json')
update_date_file = 'data_atualizacao.txt'
update_date_path = os.path.join(data_dir, update_date_file)

# --- FUNÇÃO DE ESTILO PARA GRÁFICOS ---
def aplicar_tema_moderno(fig, cores_azuis=None):
    """Applies a modern blue corporate visual theme to a Plotly chart."""
    if cores_azuis is None:
        cores_azuis = ['#0A2A66', '#1E5FCC', '#3B82F6', '#60A5FA', '#93C5FD', '#1E40AF', '#2563EB']

    fig.update_layout(
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font=dict(family="Inter, sans-serif", size=12, color='#16233F'),
        title_font=dict(family="Inter, sans-serif", size=16, color='#0A2A66'),
        colorway=cores_azuis,
        legend=dict(font=dict(color='#16233F')),
        hoverlabel=dict(font=dict(family="Inter, sans-serif"), bgcolor='#0A2A66'),
        margin=dict(t=48, l=10, r=10, b=10),
    )
    if fig.layout.title.text is None:
        fig.update_layout(title_text="")
    fig.update_traces(marker_cornerradius=8, selector=dict(type='bar'))
    fig.update_yaxes(showgrid=True, gridwidth=1, gridcolor='#E4EBF6', zeroline=False)
    fig.update_xaxes(showgrid=False)
    return fig

# --- Função para checagem de senha ---
def check_password():
    """Returns True if the correct password is entered, otherwise shows the password field."""
    if "password_correct" not in st.session_state:
        st.session_state.password_correct = False

    if st.session_state.password_correct:
        return True

    with st.form("password_form"):
        password_attempt = st.text_input("Enter password to access:", type="password")
        submitted = st.form_submit_button("Entrar")

        if submitted:
            # Compara a tentativa com a senha armazenada em st.secrets
            if "app_password" in st.secrets and password_attempt == st.secrets["app_password"]:
                st.session_state.password_correct = True
                st.rerun()
            else:
                st.error("Incorrect password. Please try again.")
    return False

# --- FUNÇÕES DE FORMATAÇÃO ---
def format_BRL(value):
    """Formats a numeric value as full BRL currency."""
    return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def format_BRL_abbreviated(value):
    """Formats a numeric value as abbreviated BRL currency."""
    if abs(value) < 1_000_000: # Menor que 1 Milhão, usa K (mil)
        val_k = value / 1_000
        # Formato: R$ 1.166K (sem casas decimais)
        return f"R$ {val_k:,.0f}K".replace(",", "X").replace(".", ",").replace("X", ".")
    elif abs(value) < 1_000_000_000: # Entre 1M e 1B, usa M (milhão)
        val_m = value / 1_000_000
        # Formato: R$ 1,17M (com 2 casas decimais)
        return f"R$ {val_m:,.2f}M".replace(",", "X").replace(".", ",").replace("X", ".")
    else: # Acima de 1B, usa B (bilhão)
        val_b = value / 1_000_000_000
        return f"R$ {val_b:,.2f}B".replace(",", "X").replace(".", ",").replace("X", ".")

# --- Funções de Carregamento de Dados Churn (COM CACHE) ---

@st.cache_data
def load_otl_projections_from_excel(filepath):
    """Loads OTL projections from an Excel file (cached)."""
    otl_values = {'OTL Churn': 0, 'OTL Churn Operacional': 0, 'OTL Backlog': 0}
    if os.path.exists(filepath):
        try:
            df_otl = pd.read_excel(filepath)
            df_otl.columns = [col.strip() for col in df_otl.columns]
            if 'OTL' in df_otl.columns and 'Valores' in df_otl.columns:
                df_otl.set_index('OTL', inplace=True)
                for key_expected in otl_values.keys():
                    if key_expected in df_otl.index:
                        otl_values[key_expected] = int(df_otl.loc[key_expected, 'Valores'])
                    elif key_expected == 'OTL Churn Operacional' and 'OTL Churn Op' in df_otl.index:
                        otl_values['OTL Churn Operacional'] = int(df_otl.loc['OTL Churn Op', 'Valores'])
            else:
                st.warning(f"NOTE: Columns 'OTL' or 'Valores' not found in file '{filepath}'.")
        except Exception as e:
            st.error(f"ERROR loading Budget file data: {e}")
    else:
        st.warning(f"NOTE: Budget projections file '{filepath}' not found.")
    return otl_values

@st.cache_data
def load_projecao(filepath):
    """Lê o JSON de projeção gerado pelo projecao_churn.py (ou None se não existir)."""
    if not os.path.exists(filepath):
        return None
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        st.warning(f"NOTE: não consegui ler a projeção em '{filepath}': {e}")
        return None


_MES_EN = {"Janeiro": "January", "Fevereiro": "February", "Março": "March",
           "Abril": "April", "Maio": "May", "Junho": "June", "Julho": "July",
           "Agosto": "August", "Setembro": "September", "Outubro": "October",
           "Novembro": "November", "Dezembro": "December"}


def _mes_en(nome):
    return _MES_EN.get(nome, nome)


def _proj_card_dual(title, d, pct=False):
    """Card branco: Realizado + Budget (Budget em destaque)."""
    real = d.get("real", "-"); budg = d.get("budget", "-")
    rp = f" <span class='proj-pct'>({d['real_pct']})</span>" if pct and d.get("real_pct") else ""
    bp = f" <span class='proj-pct budget'>({d['budget_pct']})</span>" if pct and d.get("budget_pct") else ""
    return (f"<div class='proj-card'><div class='proj-title'>{title}</div>"
            f"<div class='proj-row'><span class='lbl'>Actual</span>"
            f"<span class='val'>{real}{rp}</span></div>"
            f"<div class='proj-row budget'><span class='lbl'>Budget</span>"
            f"<span class='val'>{budg}{bp}</span></div></div>")


def _proj_card_single(title, value):
    """Card branco com um único número em destaque."""
    return (f"<div class='proj-card'><div class='proj-title'>{title}</div>"
            f"<div class='proj-row budget' style='justify-content:center;border-top:none;margin-top:0;padding-top:10px;'>"
            f"<span class='val' style='font-size:1.25rem;'>{value}</span></div></div>")


@st.cache_data
def load_and_transform_data(data_folder, file_prev_churn, file_curr_churn, file_active_base_name, file_backlog_churn_name):
    """Loads and transforms the main (Executed) data, cached."""
    df_churn = pd.DataFrame()
    df_combined = pd.DataFrame()
    df_active_processed = pd.DataFrame()
    df_backlog_processed = pd.DataFrame()
    try:
        df_prev = pd.read_excel(os.path.join(data_folder, file_prev_churn))
        df_curr = pd.read_excel(os.path.join(data_folder, file_curr_churn))
        df_combined = pd.concat([df_prev, df_curr], ignore_index=True)
    except FileNotFoundError as e: st.error(f"ERROR: CHURN .xlsx file not found. Details: {e}"); st.stop()
    except Exception as e: st.error(f"ERROR: Problem loading or combining CHURN data: {e}"); st.stop()
    try:
        df_active_raw = pd.read_excel(os.path.join(data_folder, file_active_base_name))
        df_active_raw.rename(columns={'Data': 'Data Base Ativa', 'Tipo Cliente': 'Tipo de Cliente Base Ativa Raw', 'Volume Clientes Ativos': 'Volume Base Ativa'}, inplace=True)
        df_active_raw['Data Base Ativa'] = pd.to_datetime(df_active_raw['Data Base Ativa'], errors='coerce')
        df_active_raw['Volume Base Ativa'] = pd.to_numeric(df_active_raw['Volume Base Ativa'], errors='coerce')
        df_active_raw.dropna(subset=['Data Base Ativa', 'Volume Base Ativa'], inplace=True)
        df_active_raw['Ano Base Ativa'] = df_active_raw['Data Base Ativa'].dt.year.astype(int)
        df_active_raw['Mes Base Ativa'] = df_active_raw['Data Base Ativa'].dt.month.astype(int)
        month_names_map = {1: "January", 2: "February", 3: "March", 4: "April", 5: "May", 6: "June", 7: "July", 8: "August", 9: "September", 10: "October", 11: "November", 12: "December"}
        df_active_raw['Nome Mes Ativa'] = df_active_raw['Mes Base Ativa'].map(month_names_map)
        df_active_raw['Tipo de Cliente Base Ativa Raw'] = df_active_raw['Tipo de Cliente Base Ativa Raw'].astype(str).str.strip().str.replace('\xa0', ' ')
        
        # Uso de .map ao inves de .apply em Series
        df_active_raw['Tipo de Cliente Base Ativa'] = df_active_raw['Tipo de Cliente Base Ativa Raw'].map(map_tipo_cliente)
        df_active_processed = df_active_raw.drop(columns=['Tipo de Cliente Base Ativa Raw'])
    except Exception as e:
        st.info("Note: 'base_ativa_clientes.xlsx' not found — churn-rate KPIs are limited (optional for local testing).")
        df_active_processed = pd.DataFrame()
    try:
        df_backlog_raw = pd.read_excel(os.path.join(data_folder, file_backlog_churn_name))
        df_backlog_raw.rename(columns={
            'Mês': 'Nome Mes Backlog',
            'Ano': 'Ano Backlog',
            'Backlog (Geral)': 'Volume Backlog',
            'Tipo de Cliente': 'Tipo de Cliente Backlog',
            'Tipo de Churn': 'Tipo de Churn Backlog'
        }, inplace=True)

        df_backlog_raw['Ano Backlog'] = df_backlog_raw['Ano Backlog'].astype(int)
        month_to_num_map = {
            "January": 1, "February": 2, "March": 3, "April": 4, "May": 5, "June": 6,
            "July": 7, "August": 8, "September": 9, "October": 10, "November": 11, "December": 12
        }
        df_backlog_raw['Mes Backlog'] = df_backlog_raw['Nome Mes Backlog'].map(month_to_num_map)
        df_backlog_raw['Volume Backlog'] = pd.to_numeric(df_backlog_raw['Volume Backlog'], errors='coerce').fillna(0).astype(int)
        df_backlog_raw['Tipo de Cliente Backlog'] = df_backlog_raw['Tipo de Cliente Backlog'].map(map_tipo_cliente)
        df_backlog_raw['Tipo de Churn Backlog'] = df_backlog_raw['Tipo de Churn Backlog'].astype(str).map(normalize_churn_type)

        df_backlog_processed = df_backlog_raw[['Ano Backlog', 'Mes Backlog', 'Nome Mes Backlog', 'Volume Backlog', 'Tipo de Cliente Backlog', 'Tipo de Churn Backlog']].copy()
        df_backlog_processed.dropna(subset=['Ano Backlog', 'Mes Backlog', 'Volume Backlog'], inplace=True)

    except Exception as e:
        st.info("Note: 'backlog_churn.xlsx' not found — the backlog analysis is empty (optional for local testing).")
        df_backlog_processed = pd.DataFrame()

    df_combined.rename(columns={'Datacriacaoos': 'Data de Criacao da OS', 'Statusos': 'Status da OS', 'DATADESINSTALACAO': 'Data de Desinstalacao', 'Formajuridica': 'Forma Juridica Original', 'tipoChurn': 'Tipo de Churn', 'Filialos': 'Filial'}, inplace=True)
    if 'Categoria4' in df_combined.columns:
        df_combined['Categoria4_Motivo'] = df_combined['Categoria4'].astype(str)
    else:
        df_combined['Categoria4_Motivo'] = None
        
    if 'Preçovalidado' in df_combined.columns:
        df_combined['Preçovalidado'] = pd.to_numeric(df_combined['Preçovalidado'], errors='coerce').fillna(0)
    else:
        df_combined['Preçovalidado'] = 0
        
    if 'Tipo de Churn' in df_combined.columns: df_combined = df_combined[df_combined['Tipo de Churn'].astype(str).str.strip().str.lower() != 'desconsiderar'].copy()
    # >>> Painel em ingles: tipo e motivo vem da TAXONOMIA GLOBAL (colunas ja existentes nos arquivos)
    if 'Tipo de Churn (Global Insight)' in df_combined.columns:
        df_combined['Tipo de Churn'] = df_combined['Tipo de Churn (Global Insight)'].astype(str)
    if 'Churn Sub-Reason (Global)' in df_combined.columns:
        df_combined['Categoria4_Motivo'] = df_combined['Churn Sub-Reason (Global)'].astype(str)
    
    df_combined['Data de Criacao da OS'] = pd.to_datetime(df_combined['Data de Criacao da OS'], errors='coerce')
    df_combined['Data de Desinstalacao'] = pd.to_datetime(df_combined['Data de Desinstalacao'], errors='coerce')

    df_churn = df_combined[df_combined['Status da OS'].astype(str).str.strip().str.contains('Concluído', na=False, case=False)].copy()
    if not df_churn.empty:
        # Uso de .map ao inves de .apply em Series
        df_churn['Tipo de Cliente'] = df_churn['Forma Juridica Original'].map(map_tipo_cliente)
        df_churn.dropna(subset=['Data de Desinstalacao'], inplace=True)
        df_churn['Ano Churn'] = df_churn['Data de Desinstalacao'].dt.year.astype(int)
        df_churn['Mes Churn'] = df_churn['Data de Desinstalacao'].dt.month.astype(int)
        month_names_map = {1: "January", 2: "February", 3: "March", 4: "April", 5: "May", 6: "June", 7: "July", 8: "August", 9: "September", 10: "October", 11: "November", 12: "December"}
        df_churn['Nome Mes Churn'] = df_churn['Mes Churn'].map(month_names_map)
        df_churn['Volume'] = 1
    for col in df_churn.select_dtypes(include=['object']).columns: df_churn[col] = df_churn[col].astype(str)
    for col in df_active_processed.select_dtypes(include=['object']).columns: df_active_processed[col] = df_active_processed[col].astype(str)
    return df_churn, df_active_processed, df_backlog_processed

@st.cache_data
def load_created_churn_data(data_folder, file_prev, file_curr):
    """Loads and transforms the Created churn data, cached."""
    df_created = pd.DataFrame()
    try:
        df_prev_data = pd.read_excel(os.path.join(data_folder, file_prev))
        df_curr_data = pd.read_excel(os.path.join(data_folder, file_curr))
        df_created = pd.concat([df_prev_data, df_curr_data], ignore_index=True)
    except FileNotFoundError as e: st.error(f"ERROR: CREATED CHURN .xlsx file not found. Details: {e}"); return pd.DataFrame()
    except Exception as e: st.error(f"ERROR: Problem loading CREATED CHURN data: {e}"); return pd.DataFrame()
    df_created.rename(columns={'Datacriacaoos': 'Data de Referencia', 'Statusos': 'Status da OS', 'DATADESINSTALACAO': 'Data de Desinstalacao', 'Formajuridica': 'Forma Juridica Original', 'tipoChurn': 'Tipo de Churn', 'Filialos': 'Filial'}, inplace=True)
    if 'Tipo de Churn' in df_created.columns: df_created = df_created[df_created['Tipo de Churn'].astype(str).str.strip().str.lower() != 'desconsiderar'].copy()
    # >>> Painel em ingles: tipo e motivo vem da TAXONOMIA GLOBAL (colunas ja existentes nos arquivos)
    if 'Tipo de Churn (Global Insight)' in df_created.columns:
        df_created['Tipo de Churn'] = df_created['Tipo de Churn (Global Insight)'].astype(str)
    if 'Churn Sub-Reason (Global)' in df_created.columns:
        df_created['Categoria4_Motivo'] = df_created['Churn Sub-Reason (Global)'].astype(str)
    if 'Categoria4' in df_created.columns:
        df_created['Categoria4_Motivo'] = df_created['Categoria4'].astype(str)
    else:
        df_created['Categoria4_Motivo'] = None
        
    if 'Preçovalidado' in df_created.columns:
        df_created['Preçovalidado'] = pd.to_numeric(df_created['Preçovalidado'], errors='coerce').fillna(0)
    else:
        df_created['Preçovalidado'] = 0
        
    df_created['Data de Referencia'] = pd.to_datetime(df_created['Data de Referencia'], errors='coerce')
    df_created.dropna(subset=['Data de Referencia'], inplace=True)
    
    # Uso de .map ao inves de .apply em Series
    df_created['Tipo de Cliente'] = df_created['Forma Juridica Original'].map(map_tipo_cliente)
    
    df_created['Ano Churn'] = df_created['Data de Referencia'].dt.year.astype(int)
    df_created['Mes Churn'] = df_created['Data de Referencia'].dt.month.astype(int)
    month_names_map = {1: "January", 2: "February", 3: "March", 4: "April", 5: "May", 6: "June", 7: "July", 8: "August", 9: "September", 10: "October", 11: "November", 12: "December"}
    df_created['Nome Mes Churn'] = df_created['Mes Churn'].map(month_names_map)
    df_created['Volume'] = 1
    for col in df_created.select_dtypes(include=['object']).columns: df_created[col] = df_created[col].astype(str)
    return df_created

def map_tipo_cliente(forma_juridica):
    # 3 segmentos: House Hold (HH = Small Corporate; PF/individual tambem),
    # Medium Corporate e Large Corporate. Desconhecido/vazio -> House Hold (sem 'Other').
    u = '' if pd.isna(forma_juridica) else str(forma_juridica).strip().upper()
    if u in ['MEDIUM CORPORATE', 'PME', 'P1', 'SME']: return 'Medium Corporate'
    if u in ['LARGE CORPORATE', 'C1', 'CORPORATIVO', 'CORPORATE']: return 'Large Corporate'
    return 'House Hold'

_CHURN_TYPE_NORM = {
    'voluntário':'Voluntary Churn','voluntario':'Voluntary Churn','voluntary':'Voluntary Churn','voluntary churn':'Voluntary Churn',
    'involuntário':'Involuntary Churn','involuntario':'Involuntary Churn','involuntary':'Involuntary Churn','involuntary churn':'Involuntary Churn',
    'baixa de ativo':'Involuntary Churn','asset write-off':'Involuntary Churn',
    'adjustments':'Adjustments','ajustes':'Adjustments',
}
def normalize_churn_type(v):
    return _CHURN_TYPE_NORM.get(str(v).strip().lower(), str(v).strip())

# --- Funções de Análise Avançada ---
def get_ai_analysis(_df_filtered_for_ai, _churn_view, _selected_years, _selected_months, _selected_client_types, all_client_types, _selected_churn_types, all_churn_types):
    summary = f"Churn View Analyzed: {_churn_view}\n"
    summary += f"Analysis Period: {', '.join(_selected_months)} of {', '.join(map(str, _selected_years))}\n"
    focus_area = "general"
    if len(_selected_client_types) == 1 and len(_selected_client_types) < len(all_client_types):
        focus_area = f"fully focused on the segment '{_selected_client_types[0]}'"
    elif len(_selected_churn_types) == 1 and len(_selected_churn_types) < len(all_churn_types):
         focus_area = f"fully focused on the churn type '{_selected_churn_types[0]}'"
    summary += f"Applied Filter: {focus_area} analysis\n\n"
    df_prev = _df_filtered_for_ai[_df_filtered_for_ai['Ano Churn'] == 2025]
    df_curr = _df_filtered_for_ai[_df_filtered_for_ai['Ano Churn'] == 2026]
    vol_prev = df_prev['Volume'].sum()
    vol_curr = df_curr['Volume'].sum()
    summary += "### Total Churn Volume (2026 vs 2025)\n"
    summary += f"- **2025:** {vol_prev} cancellations\n"
    summary += f"- **2026:** {vol_curr} cancellations\n\n"
    comp_cliente_prev = df_prev.groupby('Tipo de Cliente')['Volume'].sum()
    comp_cliente_curr = df_curr.groupby('Tipo de Cliente')['Volume'].sum()
    df_comp_cliente = pd.concat([comp_cliente_prev, comp_cliente_curr], axis=1, keys=['2025', '2026']).fillna(0)
    summary += "### Churn Comparison by Client Type\n"
    summary += df_comp_cliente.to_markdown()
    summary += "\n\n"
    if 'Categoria4_Motivo' in _df_filtered_for_ai.columns and _df_filtered_for_ai['Categoria4_Motivo'].nunique() > 1:
        comp_motivo_prev = df_prev['Categoria4_Motivo'].value_counts().head(5)
        comp_motivo_curr = df_curr['Categoria4_Motivo'].value_counts().head(5)
        df_comp_motivo = pd.concat([comp_motivo_prev, comp_motivo_curr], axis=1, keys=['2025', '2026']).fillna(0)
        summary += "### Top 5 Churn Reasons Comparison\n"
        summary += df_comp_motivo.to_markdown()
        summary += "\n\n"
    prompt = f"""
    **Context:** You are a **strategy consultant**. Your analysis must be sharp and prioritized. {focus_area} analysis.
    **Dados:**\n{summary}
    **Mission:** Produce an executive diagnosis with: 1. Key Insight, 2. Cause-and-Effect Diagnosis, 3. Risks and Opportunities, 4. Action Plan.
    """
    return prompt

# --- FUNÇÃO ATUALIZADA: Comparativo Customizado com Receita, Impacto, Produtos e Backlog ---
def get_ai_custom_comparison(df_full, df_backlog, year1, month1, year2, month2, client_types, churn_types, churn_view):
    month_order_num_pt = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
    month_to_num_map = {name: i+1 for i, name in enumerate(month_order_num_pt)}
    
    month1_num = month_to_num_map.get(month1, 0)
    month2_num = month_to_num_map.get(month2, 0)

    # Filtros Churn
    df1 = df_full[(df_full['Ano Churn'] == year1) & (df_full['Mes Churn'] == month1_num) & (df_full['Tipo de Cliente'].isin(client_types))]
    df2 = df_full[(df_full['Ano Churn'] == year2) & (df_full['Mes Churn'] == month2_num) & (df_full['Tipo de Cliente'].isin(client_types))]
    
    if churn_types:
        if 'Tipo de Churn' in df1.columns: df1 = df1[df1['Tipo de Churn'].isin(churn_types)]
        if 'Tipo de Churn' in df2.columns: df2 = df2[df2['Tipo de Churn'].isin(churn_types)]

    if df1.empty and df2.empty: return "Insufficient data for both periods."

    def create_complete_summary(df, df_bl, year, month_name, month_num):
        summary = ""
        vol = df['Volume'].sum()
        summary += f"### General Data ({month_name}/{year})\n"
        summary += f"- **Churn Volume:** {vol}\n"
        
        # --- Receita e Impacto ---
        if 'Preçovalidado' in df.columns:
            receita_nominal = df['Preçovalidado'].sum()
            # Impacto = Valor * (13 - Mês)
            df['Receita_Impacto_Calc'] = df['Preçovalidado'] * (13 - month_num)
            impacto_financeiro = df['Receita_Impacto_Calc'].sum()
            
            summary += f"- **Nominal Revenue Lost:** {format_BRL(receita_nominal)}\n"
            summary += f"- **Financial Impact (Year):** {format_BRL(impacto_financeiro)}\n"
        
        # --- Backlog ---
        vol_backlog = 0
        if not df_bl.empty:
            # Filtra backlog pelo ano, mês e tipos selecionados
            df_bl_filt = df_bl[
                (df_bl['Ano Backlog'] == year) & 
                (df_bl['Mes Backlog'] == month_num) & 
                (df_bl['Tipo de Cliente Backlog'].isin(client_types))
            ]
            # Se houver filtro de tipo de churn, aplica também (assumindo coluna mapeada)
            if churn_types and 'Tipo de Churn Backlog' in df_bl_filt.columns:
                 df_bl_filt = df_bl_filt[df_bl_filt['Tipo de Churn Backlog'].isin(churn_types)]
            
            vol_backlog = df_bl_filt['Volume Backlog'].sum()
        
        summary += f"- **Backlog Volume:** {vol_backlog}\n"

        # --- Distribuições ---
        client_dist = df.groupby('Tipo de Cliente')['Volume'].sum().sort_values(ascending=False)
        if not client_dist.empty: 
            summary += "\n**By Client Segment:**\n" + client_dist.to_markdown() + "\n"
        
        if 'Família' in df.columns:
            prod_dist = df['Família'].value_counts().head(5)
            if not prod_dist.empty:
                summary += "\n**Top 5 Product Families:**\n" + prod_dist.to_markdown() + "\n"

        if 'Categoria4_Motivo' in df.columns:
            reason_dist = df['Categoria4_Motivo'].value_counts().head(5)
            if not reason_dist.empty: 
                summary += "\n**Top 5 Reasons:**\n" + reason_dist.to_markdown() + "\n"
                
        return summary

    summary1 = create_complete_summary(df1, df_backlog, year1, month1, month1_num)
    summary2 = create_complete_summary(df2, df_backlog, year2, month2, month2_num)

    prompt = f"""
    Act as a Senior FP&A and Customer Experience Specialist.
    **Objective:** Perform a deep comparative churn analysis ({churn_view}).
    
    **PERIOD 1 (Base): {month1}/{year1}**
    {summary1}
    
    **PERIOD 2 (Comparison): {month2}/{year2}**
    {summary2}
    
    **Request:**
    Produce an executive report comparing the two periods. Your analysis must include:
    1. **Volume and Financial Variation:** Compare not only the number of customers, but nominal revenue loss and annualized financial impact. Did the average ticket change?
    2. **Backlog Analysis:** How did the backlog behave between periods? Does it indicate a growing or shrinking operational problem?
    3. **Product Mix:** Did the most-cancelled product families change? Did any specific product get worse?
    4. **Root Cause and Segments:** Identify whether the customer profile or the main reasons changed.
    5. **Conclusion:** Did the picture get worse or better financially and operationally?
    """
    return prompt

def get_ai_lead_time_analysis(df_lead_time):
    summary_stats = df_lead_time['Lead_Time_Days'].describe()
    lead_time_by_client = df_lead_time.groupby('Tipo de Cliente')['Lead_Time_Days'].mean().sort_values(ascending=False)
    data_summary = f"**Statistics (days):**\n{summary_stats.to_markdown()}\n**Averages by Client:**\n{lead_time_by_client.to_markdown()}"
    prompt = f"""
    **Context:** Churn Lead Time Analysis.
    **Dados:**\n{data_summary}
    **Mission:** 1. Quick Diagnosis, 2. Hypothesis (Opportunity vs Friction), 3. Recommendation.
    """
    return prompt

def get_ai_root_cause_deep_dive(df_churn_slice, reason, segment, churn_view):
    if df_churn_slice.empty: return "No data."
    total_volume = len(df_churn_slice)
    data_summary = f"Analysis: **{total_volume}** cancellations of **'{segment}'** due to **'{reason}'**.\n"
    if 'Filial' in df_churn_slice.columns: data_summary += "**Top Branches:**\n" + (df_churn_slice['Filial'].value_counts(normalize=True).head(5)*100).map('{:.1f}%'.format).to_markdown() + "\n"
    vol_prev = len(df_churn_slice[df_churn_slice['Ano Churn'] == 2025])
    vol_curr = len(df_churn_slice[df_churn_slice['Ano Churn'] == 2026])
    data_summary += f"**Trend (26 vs 25):** 2025: {vol_prev} | 2026: {vol_curr}\n"
    prompt = f"""
    **Context:** Root Cause Investigation.
    **Dossier:**\n{data_summary}
    **Mission:** 1. Executive Headline, 2. Deeper Analysis, 3. Consolidated Diagnosis, 4. Next Step.
    """
    return prompt

# --- DIALOG DO RELATÓRIO EXECUTIVO ---
@st.dialog("Executive Summary for the Board")
def show_executive_report(
    date_str, 
    total_churn, otl_churn, churn_rate_str,
    churn_op, otl_churn_op,
    backlog, otl_backlog,
    rev_jan_otl, tkm_jan, rev_anual_otl, rev_ytd_otl,
    intencoes_pf, retidos_pf, nao_retidos_pf, conv_pf,
    intencoes_corp, retidos_corp, nao_retidos_corp, conv_corp
):
    
    report_text = f"""Updated churn projections {date_str}

*Executed Churn:* OTL {otl_churn:,.0f} ({churn_rate_str} Ecohouse)
*Operational Churn:* {churn_op:,.0f} (OTL {otl_churn_op:,.0f})
*Total Backlog:* {backlog:,.0f} (OTL {otl_backlog:,.0f})

*OTL Churn Revenue Current Month:* {format_BRL_abbreviated(rev_jan_otl)} ({format_BRL(tkm_jan)} Tkm Ecohouse)
*OTL Churn Revenue Annual (Proj.):* {format_BRL_abbreviated(rev_anual_otl)} (Ecohouse)
*OTL Churn Revenue YTD:* {format_BRL_abbreviated(rev_ytd_otl)}

*Cancellation Intentions Individual&SME:* {intencoes_pf} Retained: {retidos_pf} Not Retained: {nao_retidos_pf} Conversion: {conv_pf:.2f}%
*Cancellation Intentions Corp:* {intencoes_corp} Retained: {retidos_corp} Not Retained: {nao_retidos_corp} Conversion: {conv_corp:.2f}%
""".replace(",", "X").replace(".", ",").replace("X", ".")

    st.text_area("Copy the text below:", value=report_text, height=350)
    st.caption("Note: Values filtered per the sidebar selection (Year/Month).")


# --- Função Principal do Aplicativo Streamlit ---
def main():
    st.set_page_config(layout="wide", page_title="Churn Brazil")

    # APLICAÇÃO DOS ESTILOS VISUAIS
    st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
        :root{
            --navy:#0A2A66; --blue:#1E5FCC; --blue2:#3B82F6; --blue-soft:#EAF1FB;
            --ink:#16233F; --muted:#647393; --line:#E4EBF6; --bg:#F4F7FD; --card:#FFFFFF;
        }
        /* Escala geral menor para caber a 100% de zoom */
        html{ font-size:13px; }
        html, body, [class*="css"], .stApp, [data-testid="stAppViewContainer"] * { font-family:'Inter', sans-serif; }
        [data-testid="stAppViewContainer"]{ background:var(--bg); }
        [data-testid="stHeader"]{ background:rgba(0,0,0,0); height:0; }
        .block-container{ padding-top:1.6rem; padding-bottom:2rem; max-width:1600px; }
        h1,h2,h3,h4{ color:var(--navy); font-weight:700; letter-spacing:-0.01em; }
        h2{ font-size:1.22rem !important; }
        h3{ font-size:1.0rem !important; }
        [data-testid="stMarkdownContainer"] p{ font-size:0.9rem; }

        /* Sidebar */
        [data-testid="stSidebar"]{ background:var(--card); border-right:1px solid var(--line); }
        [data-testid="stSidebar"] .block-container{ padding-top:1rem; }
        [data-testid="stSidebar"] h1{ font-size:1.3rem; color:var(--navy); font-weight:800; text-align:center; margin:.3rem 0 .2rem; }
        [data-testid="stSidebar"] h2{ font-size:.9rem !important; color:var(--muted); text-transform:uppercase; letter-spacing:.08em; }
        [data-testid="stSidebar"] label, [data-testid="stSidebar"] p, [data-testid="stSidebar"] span{ color:var(--ink); }
        [data-testid="stSidebar"] img{ border-radius:8px; }

        /* KPI cards (compactos) */
        .kpi-container{
            background:var(--card); border:1px solid var(--line); border-radius:12px;
            padding:11px 8px 9px; text-align:center;
            box-shadow:0 1px 2px rgba(16,40,90,.04), 0 8px 20px rgba(16,40,90,.05);
            transition:transform .2s ease, box-shadow .2s ease;
            height:102px; display:flex; flex-direction:column; justify-content:center;
            position:relative; overflow:hidden;
        }
        .kpi-container::before{ content:""; position:absolute; top:0; left:0; right:0; height:3px; background:linear-gradient(90deg,var(--navy),var(--blue2)); }
        .kpi-container:hover{ transform:translateY(-3px); box-shadow:0 2px 4px rgba(16,40,90,.06), 0 14px 28px rgba(16,40,90,.10); }
        .kpi-title{ font-size:.62rem; color:var(--muted); font-weight:600; text-transform:uppercase; letter-spacing:.04em; margin-bottom:5px; line-height:1.15; }
        .kpi-value{ font-size:1.25rem; color:var(--navy); font-weight:800; line-height:1.05; }
        .kpi-value-small{ font-size:1.05rem; color:var(--navy); font-weight:800; line-height:1.05; }
        .kpi-value-revenue{ font-size:.9rem; color:var(--navy); font-weight:800; line-height:1.05; }
        .kpi-sub-value{ font-size:.62rem; color:var(--muted); margin-top:4px; }
        .kpi-delta{ font-size:.64rem; font-weight:700; }
        .kpi-value.positive, .kpi-delta.positive{ color:#0E9F6E; }
        .kpi-value.negative, .kpi-delta.negative{ color:#E02424; }

        /* Budget (highlight) box */
        .otl-box{
            background:linear-gradient(180deg,#0A2A66 0%, #163D8C 100%);
            border:1px solid #0A2A66; border-radius:12px; padding:9px 8px; text-align:center;
            box-shadow:0 10px 22px rgba(10,42,102,.20); height:102px;
            display:flex; flex-direction:column; justify-content:center;
        }
        .otl-title{ font-size:.60rem; color:#BBD0F2; font-weight:700; text-transform:uppercase; letter-spacing:.04em; margin-bottom:5px; }
        .otl-item{ font-size:.70rem; color:#EAF1FB; line-height:1.45; display:flex; justify-content:space-between; padding:0 4px; }
        .otl-value{ font-weight:800; color:#FFFFFF; }

        /* Tabs */
        [data-baseweb="tab-list"]{ gap:3px; border-bottom:1px solid var(--line); }
        [data-baseweb="tab-list"] button{ border-radius:9px 9px 0 0; padding:7px 13px; color:var(--muted); font-weight:600; font-size:.85rem; }
        [data-baseweb="tab-list"] button:hover{ background:var(--blue-soft); color:var(--navy); }
        [data-baseweb="tab-list"] button[aria-selected="true"]{ background:var(--blue-soft); color:var(--navy); border-bottom:3px solid var(--blue); font-weight:700; }

        /* Dataframe */
        [data-testid="stDataFrame"]{ border:1px solid var(--line); border-radius:10px; overflow:hidden; box-shadow:0 5px 14px rgba(16,40,90,.05); }

        /* Metric widget */
        [data-testid="stMetric"]{ background:var(--card); border:1px solid var(--line); border-radius:10px; padding:10px 14px; box-shadow:0 5px 14px rgba(16,40,90,.05); }
        [data-testid="stMetricLabel"] p{ color:var(--muted); font-weight:600; font-size:.8rem; }
        [data-testid="stMetricValue"]{ color:var(--navy); font-weight:800; font-size:1.4rem; }

        /* Buttons & inputs */
        .stButton>button, .stDownloadButton>button{ background:var(--navy); color:#fff; border:none; border-radius:9px; padding:.45rem .9rem; font-weight:600; }
        .stButton>button:hover{ background:var(--blue); color:#fff; }
        hr{ border-color:var(--line); }
        /* Projection tab */
        .proj-card{ box-sizing:border-box; background:var(--card); border:1px solid var(--line); border-radius:12px; padding:12px 14px 11px; box-shadow:0 1px 2px rgba(16,40,90,.04), 0 8px 20px rgba(16,40,90,.05); position:relative; overflow:hidden; min-height:96px; }
        .proj-card::before{ content:""; position:absolute; top:0; left:0; right:0; height:3px; background:linear-gradient(90deg,var(--navy),var(--blue2)); }
        .proj-title{ font-size:.64rem; color:var(--muted); font-weight:700; text-transform:uppercase; letter-spacing:.04em; margin-bottom:8px; }
        .proj-row{ display:flex; justify-content:space-between; align-items:baseline; padding:3px 0; font-size:.85rem; color:var(--ink); }
        .proj-row .lbl{ color:var(--muted); font-weight:600; font-size:.72rem; text-transform:uppercase; letter-spacing:.03em; }
        .proj-row .val{ font-weight:700; color:var(--ink); }
        .proj-row.budget{ border-top:1px dashed var(--line); margin-top:3px; padding-top:6px; }
        .proj-row.budget .val{ color:var(--navy); font-weight:800; font-size:1.12rem; }
        .proj-pct{ font-weight:700; color:var(--muted); font-size:.78rem; }
        .proj-pct.budget{ color:var(--blue); }
        .proj-table{ width:100%; border-collapse:collapse; background:var(--card); border:1px solid var(--line); border-radius:10px; overflow:hidden; box-shadow:0 5px 14px rgba(16,40,90,.05); font-size:.84rem; margin-top:6px; }
        .proj-table thead th{ background:var(--navy); color:#EAF1FB; font-weight:700; text-transform:uppercase; letter-spacing:.03em; font-size:.66rem; padding:9px 12px; text-align:right; }
        .proj-table thead th:first-child{ text-align:left; }
        .proj-table tbody td{ padding:8px 12px; text-align:right; color:var(--ink); border-top:1px solid var(--line); }
        .proj-table tbody td:first-child{ text-align:left; font-weight:600; color:var(--navy); }
        .proj-table tbody tr:nth-child(even){ background:#FAFCFF; }
        .proj-table tbody tr.atual{ background:var(--blue-soft); }
        .proj-table tbody tr.atual td{ font-weight:800; color:var(--navy); border-top:1px solid #C9DBF7; }
        </style>
    """, unsafe_allow_html=True)
    
    def reset_criado_message_state():
        if 'criado_message_shown' in st.session_state:
            del st.session_state.criado_message_shown

    with st.sidebar:
        _logo_path = os.path.join(data_dir, "logo.png")
        if os.path.exists(_logo_path):
            _lc1, _lc2, _lc3 = st.columns([1, 2, 1])
            with _lc2:
                st.image(_logo_path, use_container_width=True)
        st.markdown("<h1 style='text-align:center; margin:0.3rem 0 0.2rem; font-weight:800;'>Churn Brazil</h1>", unsafe_allow_html=True)
        try:
            with open(update_date_path, 'r', encoding='utf-8') as f:
                formatted_date = f.read().strip().split(' (')[0].strip()
                if not formatted_date: raise FileNotFoundError
        except (FileNotFoundError, IOError):
            current_time = datetime.now()
            portuguese_month_names = {m:n for m, n in zip(range(1, 13), ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"])}
            formatted_date = f"{portuguese_month_names[current_time.month]} {current_time.day}, {current_time.year}"
        st.markdown(f"<p style='text-align:center;color:#647393;font-size:0.78rem;margin-top:-0.2rem;'>Last update: {formatted_date}</p>", unsafe_allow_html=True)

        st.header("Filters")
        
        churn_view = st.radio(
            "Select Churn View",
            ('Executed', 'Created'),
            on_change=reset_criado_message_state
        )

    # Carregar dados
    df_churn_criado_kpi = load_created_churn_data(data_dir, file_creation_previous, file_creation_current)

    if churn_view == 'Executed':
        df_churn, df_active_raw, df_backlog_raw = load_and_transform_data(data_dir, file_previous_year, file_current_year, file_active_base, file_backlog_churn)
    else:
        df_churn = df_churn_criado_kpi.copy()
        df_active_raw, df_backlog_raw = pd.DataFrame(), pd.DataFrame()
        
    if churn_view == 'Created' and not st.session_state.get('criado_message_shown', False):
        st.warning(
            """
            **Important note about the Created Churn view:**
            After 90 days, RPI work orders are automatically cancelled and the contract goes through the sbaff process...
            """, icon="ℹ️"
        )
        if st.button("Got it"):
            st.session_state.criado_message_shown = True
            st.rerun()
        st.stop() 

    otl_projections = load_otl_projections_from_excel(otl_projections_file)
    projecao = load_projecao(projecao_file)

    if df_churn.empty:
        st.error(f"ERROR: No CHURN data found ({churn_view}). Please check the source files."); st.stop()

    with st.sidebar:
        all_years = sorted(df_churn['Ano Churn'].unique())
        default_year_selection = [2026] if 2026 in all_years else ["All"]
        selected_years_option = st.multiselect("Select Year(s)", ["All"] + all_years, default=default_year_selection)
        selected_years = all_years if "All" in selected_years_option or not selected_years_option else selected_years_option

        month_order_num_pt = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
        month_abbr_order_pt = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
        month_to_abbr_map = dict(zip(month_order_num_pt, month_abbr_order_pt))
        all_months = sorted(df_churn['Nome Mes Churn'].unique(), key=lambda x: month_order_num_pt.index(x) if x in month_order_num_pt else -1)
        _cur_month_name = month_order_num_pt[datetime.now().month - 1]
        if _cur_month_name in all_months: _default_months = [_cur_month_name]
        elif all_months: _default_months = [all_months[-1]]
        else: _default_months = ["All"]
        selected_months_option = st.multiselect("Select Month(s)", ["All"] + all_months, default=_default_months)
        selected_months = all_months if "All" in selected_months_option or not selected_months_option else selected_months_option

        all_churn_types_churn = sorted([str(x) for x in df_churn['Tipo de Churn'].dropna().unique()]) if 'Tipo de Churn' in df_churn.columns and not df_churn['Tipo de Churn'].isnull().all() else []
        all_churn_types_backlog = sorted([str(x) for x in df_backlog_raw['Tipo de Churn Backlog'].dropna().unique()]) if 'Tipo de Churn Backlog' in df_backlog_raw.columns and not df_backlog_raw['Tipo de Churn Backlog'].isnull().all() else []
        all_churn_types = sorted(list(set(all_churn_types_churn + all_churn_types_backlog)))
        all_churn_types = [t for t in all_churn_types if t not in ['nan', 'None', '', '<NA>']]
        selected_churn_type_option = st.radio("Select Churn Type", ["All"] + all_churn_types)
        selected_churn_types = all_churn_types if selected_churn_type_option == "All" else [selected_churn_type_option]

        all_client_types_churn = sorted([str(x) for x in df_churn['Tipo de Cliente'].dropna().unique()]) if 'Tipo de Cliente' in df_churn.columns and not df_churn['Tipo de Cliente'].isnull().all() else []
        all_client_types_backlog = sorted([str(x) for x in df_backlog_raw['Tipo de Cliente Backlog'].dropna().unique()]) if 'Tipo de Cliente Backlog' in df_backlog_raw.columns and not df_backlog_raw['Tipo de Cliente Backlog'].isnull().all() else []
        all_client_types = sorted(list(set(all_client_types_churn + all_client_types_backlog)))
        all_client_types = [c for c in all_client_types if c not in ['nan', 'None', '', '<NA>']]
        selected_client_type_option_radio = st.radio("Select Client Type(s)", ["All"] + all_client_types)
        selected_client_types = all_client_types if selected_client_type_option_radio == "All" else [selected_client_type_option_radio]

    # Filtrando os dados
    df_kpi1_filtered = df_churn[
        (df_churn['Ano Churn'].isin(selected_years)) &
        (df_churn['Nome Mes Churn'].isin(selected_months)) &
        (df_churn['Tipo de Cliente'].isin(selected_client_types))
    ]
    if selected_churn_types and 'Tipo de Churn' in df_kpi1_filtered.columns:
        df_kpi1_filtered = df_kpi1_filtered[df_kpi1_filtered['Tipo de Churn'].isin(selected_churn_types)]

    df_filtered = df_churn[
        (df_churn['Nome Mes Churn'].isin(selected_months)) &
        (df_churn['Tipo de Cliente'].isin(selected_client_types))
    ]
    if selected_churn_types and 'Tipo de Churn' in df_filtered.columns:
        df_filtered = df_filtered[df_filtered['Tipo de Churn'].isin(selected_churn_types)]

    if df_filtered.empty: st.warning("No CHURN data found for the selected filters."); st.stop()

    st.markdown(f"### Churn - Performance Indicators")
    def _latest_month(_df, _ycol, _mcol, _years):
        _sub = _df[_df[_ycol].isin(_years)] if _years else _df
        _months = sorted(pd.to_numeric(_sub[_mcol], errors='coerce').dropna().astype(int).unique()) if not _sub.empty else []
        if not _months: return None
        _cur = datetime.now().month
        return _cur if _cur in _months else max(_months)

    col1, col2, col3, col4, col5, col6, col7, col8, col9 = st.columns(9)

    with col1:
        total_churn_selected_years = df_kpi1_filtered['Volume'].sum()
        years_str = ', '.join(map(str, selected_years)) if selected_years else "None"
        st.markdown(f"""<div class="kpi-container"><div class="kpi-title">Total Churn {churn_view} ({years_str})</div><div class="kpi-value">{total_churn_selected_years:,.0f}</div></div>""".replace(",","."), unsafe_allow_html=True)
    with col3:
        df_kpi_revenue = df_kpi1_filtered.copy()
        df_kpi_revenue['Mes Churn'] = pd.to_numeric(df_kpi_revenue['Mes Churn'], errors='coerce').fillna(0).astype(int)
        df_kpi_revenue['Receita_Impacto'] = df_kpi_revenue['Preçovalidado'] * (13 - df_kpi_revenue['Mes Churn'])
        total_revenue_impact = df_kpi_revenue['Receita_Impacto'].sum()
        total_nominal_revenue = df_kpi_revenue['Preçovalidado'].sum()
        total_churn_volume = df_kpi_revenue['Volume'].sum()
        avg_ticket = (total_nominal_revenue / total_churn_volume) if total_churn_volume > 0 else 0
        full_revenue_display = format_BRL(total_revenue_impact)
        abbreviated_revenue_display = format_BRL_abbreviated(total_revenue_impact)
        avg_ticket_display = format_BRL(avg_ticket)
        st.markdown(f"""<div class="kpi-container"><div class="kpi-title">Fin. Impact (Year)</div><div class="kpi-value-revenue" title="Total impact: {full_revenue_display}">{abbreviated_revenue_display}</div><div class="kpi-sub-value" style="margin-top: 5px;"><b>Avg. Ticket:</b> {avg_ticket_display}</div></div>""", unsafe_allow_html=True)
    with col2:
        df_kpi_criado_filtered = df_churn_criado_kpi[
            (df_churn_criado_kpi['Ano Churn'].isin(selected_years)) &
            (df_churn_criado_kpi['Nome Mes Churn'].isin(selected_months)) &
            (df_churn_criado_kpi['Tipo de Cliente'].isin(selected_client_types))
        ]
        if selected_churn_types and 'Tipo de Churn' in df_kpi_criado_filtered.columns:
            df_kpi_criado_filtered = df_kpi_criado_filtered[df_kpi_criado_filtered['Tipo de Churn'].isin(selected_churn_types)]
        total_churn_criado_filtrado = df_kpi_criado_filtered['Volume'].sum()
        years_str = ', '.join(map(str, selected_years)) if selected_years else "None"
        st.markdown(f"""<div class="kpi-container"><div class="kpi-title">Total Created Churn ({years_str})</div><div class="kpi-value">{total_churn_criado_filtrado:,.0f}</div></div>""".replace(",","."), unsafe_allow_html=True)
    with col4:
        total_backlog_selected_months = 0
        if churn_view == 'Executed' and not df_backlog_raw.empty and selected_years:
            _ref_m_bl = _latest_month(df_backlog_raw, 'Ano Backlog', 'Mes Backlog', selected_years)
            if _ref_m_bl is not None:
                df_backlog_filtered_for_kpi = df_backlog_raw[(df_backlog_raw['Ano Backlog'].isin(selected_years)) & (df_backlog_raw['Mes Backlog'] == _ref_m_bl) & (df_backlog_raw['Tipo de Cliente Backlog'].isin(selected_client_types)) & (df_backlog_raw['Tipo de Churn Backlog'].isin(selected_churn_types))]
                total_backlog_selected_months = df_backlog_filtered_for_kpi['Volume Backlog'].sum()
        display_value_backlog = f"{int(total_backlog_selected_months):,.0f}".replace(",",".") if isinstance(total_backlog_selected_months, (int,float)) else str(total_backlog_selected_months)
        st.markdown(f"""<div class="kpi-container"><div class="kpi-title">Total Backlog (Month)</div><div class="kpi-value">{display_value_backlog}</div></div>""",unsafe_allow_html=True)
    with col5:
        projected_annual_churn = 0
        if all_years:
            current_year_churn_proj = max(all_years)
            df_current_year_churn = df_filtered[df_filtered['Ano Churn'] == current_year_churn_proj]
            if not df_current_year_churn.empty:
                num_months_data_churn = df_current_year_churn['Mes Churn'].nunique()
                if num_months_data_churn > 0: projected_annual_churn = (df_current_year_churn['Volume'].sum()/num_months_data_churn)*12
        display_value_proj = f"{int(projected_annual_churn):,.0f}".replace(",","." ) if projected_annual_churn > 0 else "-"
        st.markdown(f"""<div class="kpi-container"><div class="kpi-title">{churn_view} Annual Projection</div><div class="kpi-value">{display_value_proj}</div></div>""",unsafe_allow_html=True)
    with col6:
        churn_rate_value, avg_monthly_active_calc = "-", 0
        if churn_view == 'Executed' and not df_active_raw.empty and all_years:
            current_year_churn_proj_calc = max(all_years)
            df_current_year_churn_calc = df_churn[(df_churn['Ano Churn']==current_year_churn_proj_calc)&(df_churn['Nome Mes Churn'].isin(selected_months))&(df_churn['Tipo de Cliente'].isin(selected_client_types))]
            if selected_churn_types: df_current_year_churn_calc = df_current_year_churn_calc[df_current_year_churn_calc['Tipo de Churn'].isin(selected_churn_types)]
            projected_annual_churn_calc = 0
            if not df_current_year_churn_calc.empty:
                num_months_data_churn_calc = df_current_year_churn_calc['Mes Churn'].nunique()
                if num_months_data_churn_calc > 0: projected_annual_churn_calc = (df_current_year_churn_calc['Volume'].sum()/num_months_data_churn_calc)*12
            df_current_year_active_filtered_calc = df_active_raw[(df_active_raw['Mes Base Ativa'].isin([month_order_num_pt.index(m)+1 for m in selected_months]))&(df_active_raw['Tipo de Cliente Base Ativa'].isin(selected_client_types))]
            if not df_current_year_active_filtered_calc.empty:
                num_months_active_data_calc = df_current_year_active_filtered_calc['Mes Base Ativa'].nunique()
                if num_months_active_data_calc > 0: avg_monthly_active_calc = df_current_year_active_filtered_calc['Volume Base Ativa'].sum()/num_months_active_data_calc
            if avg_monthly_active_calc > 0: churn_rate_value = (projected_annual_churn_calc/avg_monthly_active_calc)*100
        display_value_cr = f"{churn_rate_value:.2f}%".replace(".",",") if isinstance(churn_rate_value,(int,float)) else "-"
        st.markdown(f"""<div class="kpi-container"><div class="kpi-title">Proj. Annual Churn Rate</div><div class="kpi-value-small">{display_value_cr}</div></div>""",unsafe_allow_html=True)
    with col7:
        current_month_active = 0
        if not df_active_raw.empty and selected_years:
            _ref_m_ab = _latest_month(df_active_raw, 'Ano Base Ativa', 'Mes Base Ativa', selected_years)
            if _ref_m_ab is not None:
                df_active_month = df_active_raw[(df_active_raw['Ano Base Ativa'].isin(selected_years)) & (df_active_raw['Mes Base Ativa'] == _ref_m_ab) & (df_active_raw['Tipo de Cliente Base Ativa'].isin(selected_client_types))]
                current_month_active = df_active_month['Volume Base Ativa'].sum()
        display_value_b = f"{int(current_month_active):,.0f}".replace(",",".") if current_month_active > 0 else "-"
        st.markdown(f"""<div class="kpi-container"><div class="kpi-title">Active Base (Month)</div><div class="kpi-value-small">{display_value_b}</div></div>""",unsafe_allow_html=True)
    with col8:
        df_monthly_volumes_kpi = df_filtered[df_filtered['Ano Churn'].isin([2025,2026])].groupby(['Ano Churn','Mes Churn']).agg(Volume_Churn=('Volume','sum')).reset_index()
        df_comparison = df_monthly_volumes_kpi.pivot_table(index='Mes Churn',columns='Ano Churn',values='Volume_Churn').reset_index()
        if 2025 not in df_comparison.columns: df_comparison[2025] = 0
        if 2026 not in df_comparison.columns: df_comparison[2026] = 0
        df_comparison_filtered_months = df_comparison[df_comparison['Mes Churn'].isin([month_order_num_pt.index(m)+1 for m in selected_months])]
        absolute_diff_yoy = df_comparison_filtered_months[2026].sum() - df_comparison_filtered_months[2025].sum()
        valid_comparison_rows = df_comparison_filtered_months[df_comparison_filtered_months[2025]>0]
        average_monthly_percentage_variation = ((valid_comparison_rows[2026]-valid_comparison_rows[2025])/valid_comparison_rows[2025]).mean() if not valid_comparison_rows.empty else pd.NA
        display_value_yoy = f"{int(absolute_diff_yoy):+,.0f}".replace(",","." )
        percentage_text_yoy = f'<div class="kpi-delta">({average_monthly_percentage_variation:.2%})</div>'.replace(".",",") if pd.notna(average_monthly_percentage_variation) else '<div class="kpi-delta">(-)</div>'
        delta_color_class_yoy = "positive" if absolute_diff_yoy < 0 else "negative"
        st.markdown(f"""<div class="kpi-container"><div class="kpi-title">Variation vs 2025</div><div class="kpi-value {delta_color_class_yoy}">{display_value_yoy}</div>{percentage_text_yoy}</div>""",unsafe_allow_html=True)
    with col9:
        if projecao and projecao.get("cards"):
            otl_churn_str = str(projecao["cards"].get("executado", {}).get("budget", 0))
            otl_backlog_str = str(projecao["cards"].get("backlog", {}).get("budget", 0))
        else:
            otl_churn_str = f'{otl_projections.get("OTL Churn", 0):,.0f}'.replace(",", ".")
            otl_backlog_str = f'{otl_projections.get("OTL Backlog", 0):,.0f}'.replace(",", ".")

        otl_churn_op = otl_projections.get("OTL Churn Operacional", 0)
        html_content = f"""<div class="otl-box"><div class="otl-title">Budget Current Month</div><div class="otl-item"><span>Churn Ex.:</span><span class="otl-value">{otl_churn_str}</span></div>"""
        if otl_churn_op > 0: html_content += f'<div class="otl-item"><span>Churn Op.:</span><span class="otl-value">{otl_churn_op:,.0f}</span></div>'.replace(",", ".")
        html_content += f'<div class="otl-item"><span>Backlog:</span><span class="otl-value">{otl_backlog_str}</span></div></div>'
        st.markdown(html_content, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    
    tabs = st.tabs([
        "Monthly Churn", "Projection", "Operational Churn", "By Segment",
        "By Type", "Churn Reasons", "By Franchise",
        "By Product", "Revenue", "Analytics"
    ])
    
    with tabs[0]:
        st.header(f"Monthly Churn by Year and Variation")
        df_plot_monthly_volume = df_filtered.groupby(['Ano Churn', 'Mes Churn', 'Nome Mes Churn']).agg(Volume_Churn=('Volume', 'sum')).reset_index().sort_values(by=['Ano Churn', 'Mes Churn'])
        df_plot_monthly_volume['Churn_Rate'] = pd.NA
        if churn_view == 'Executed' and not df_active_raw.empty:
            df_active_monthly_volumes = df_active_raw[df_active_raw['Tipo de Cliente Base Ativa'].isin(selected_client_types)].groupby(['Ano Base Ativa', 'Mes Base Ativa', 'Nome Mes Ativa']).agg(Volume_Base_Ativa=('Volume Base Ativa', 'sum')).reset_index()
            df_active_monthly_volumes.rename(columns={'Ano Base Ativa': 'Ano Churn', 'Mes Base Ativa': 'Mes Churn', 'Nome Mes Ativa': 'Nome Mes Churn'}, inplace=True)
            df_plot_monthly_volume = pd.merge(df_plot_monthly_volume, df_active_monthly_volumes, on=['Ano Churn', 'Mes Churn', 'Nome Mes Churn'], how='left')
            df_plot_monthly_volume['Churn_Rate'] = df_plot_monthly_volume.apply(lambda row: (row['Volume_Churn'] / row['Volume_Base_Ativa'] * 100) if row['Volume_Base_Ativa'] > 0 else float('nan'), axis=1)
        df_plot_monthly_volume['Bar_Text_Label'] = df_plot_monthly_volume.apply(lambda row: (f"{row['Volume_Churn']:,.0f}".replace(",", ".") + (f"<br>{row['Churn_Rate']:.2f}%".replace(".", ",") if pd.notna(row['Churn_Rate']) and row['Ano Churn'] == 2026 else "")), axis=1)
        df_yoy_comparison = df_plot_monthly_volume.pivot_table(index=['Mes Churn', 'Nome Mes Churn'], columns='Ano Churn', values='Volume_Churn').reset_index()
        if 2025 in df_yoy_comparison.columns and 2026 in df_yoy_comparison.columns: df_yoy_comparison['YoY_Variation'] = ((df_yoy_comparison[2026] - df_yoy_comparison[2025]) / df_yoy_comparison[2025].replace(0, pd.NA)).fillna(pd.NA)
        else: df_yoy_comparison['YoY_Variation'] = pd.NA
        df_yoy_pivot_for_label = df_yoy_comparison.copy()
        df_yoy_pivot_for_label['X_Axis_Month_Label'] = df_yoy_pivot_for_label.apply(lambda row: (f"{row['Nome Mes Churn']}" + (f"<br>({row['YoY_Variation']:.1%})".replace(".", ",") if pd.notna(row['YoY_Variation']) else "")), axis=1)
        df_plot_monthly = pd.merge(df_plot_monthly_volume, df_yoy_pivot_for_label[['Mes Churn', 'X_Axis_Month_Label', 'YoY_Variation']], on='Mes Churn', how='left')
        fig_monthly_bar_with_variation = px.bar(df_plot_monthly, x="X_Axis_Month_Label", y="Volume_Churn", color=df_plot_monthly['Ano Churn'].astype(str), barmode="group", labels={"X_Axis_Month_Label": "Month (YoY Variation)", "color": "Year"}, category_orders={"X_Axis_Month_Label": sorted(df_yoy_pivot_for_label['X_Axis_Month_Label'].unique(), key=lambda x: month_order_num_pt.index(x.split('<br>')[0]))}, text='Bar_Text_Label')
        fig_monthly_bar_with_variation.update_traces(textposition='outside', textfont=dict(color='black', weight='bold', size=10), textangle=0, hovertemplate="<b>Month:</b> %{customdata[1]}<br><b>Year:</b> %{fullData.name}<br><b>Volume:</b> %{y:,.0f}".replace(",", ".") + "<br><b>Variation (26 vs 25):</b> %{customdata[0]:.1%}<extra></extra>".replace(".", ","), customdata=df_plot_monthly[['YoY_Variation', 'Nome Mes Churn']])
        fig_monthly_bar_with_variation.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
        st.plotly_chart(aplicar_tema_moderno(fig_monthly_bar_with_variation), use_container_width=True)

    with tabs[1]:
        if not projecao:
            st.header("Closing Projection")
            st.info("No projection found. Run **projecao_churn.py** and copy the generated **projecao_churn.json** into the dashboard folder.")
        else:
            pj = projecao
            cards = pj.get("cards", {})
            st.header(f"Closing Projection — {_mes_en(pj.get('nome_mes',''))}/{pj.get('ano','')}")
            st.markdown(
                "<p style='color:#647393;font-size:.85rem;margin-top:-.4rem;'>"
                "Actual to date vs. <b>Budget</b> (projected close)"
                f" &nbsp;•&nbsp; Generated on {pj.get('gerado_em','-')}"
                f" &nbsp;•&nbsp; Active base: {pj.get('base_ativa','-')}"
                f" &nbsp;•&nbsp; Scheduling failure rate: {pj.get('pct_insucesso','-')}</p>",
                unsafe_allow_html=True)

            row1 = st.columns(3)
            with row1[0]:
                st.markdown(_proj_card_dual("Executed Churn", cards.get("executado", {}), pct=True), unsafe_allow_html=True)
            with row1[1]:
                st.markdown(_proj_card_dual("Created Churn", cards.get("criado", {}), pct=True), unsafe_allow_html=True)
            with row1[2]:
                st.markdown(_proj_card_dual("Total Backlog", cards.get("backlog", {})), unsafe_allow_html=True)

            st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)
            row2 = st.columns(3)
            with row2[0]:
                st.markdown(_proj_card_dual("Churn Revenue", cards.get("receita", {})), unsafe_allow_html=True)
            with row2[1]:
                st.markdown(_proj_card_single(f"{_mes_en(pj.get('nome_mes',''))} Annual Revenue", cards.get("anual", "-")), unsafe_allow_html=True)
            with row2[2]:
                st.markdown(_proj_card_dual("YTD Churn Revenue", cards.get("ytd", {})), unsafe_allow_html=True)

            st.markdown("<div style='height:18px;'></div>", unsafe_allow_html=True)
            st.markdown(f"##### Monthly Revenue ({pj.get('ano','')})")
            linhas = pj.get("tabela", [])
            if linhas:
                linhas_html = ""
                for r in linhas:
                    cls = " class='atual'" if r.get("atual") else ""
                    linhas_html += (f"<tr{cls}><td>{_mes_en(r.get('nome',''))}</td>"
                                    f"<td>{r.get('receita_mes','')}</td>"
                                    f"<td>{r.get('tkm','')}</td>"
                                    f"<td>{r.get('receita_ano','')}</td></tr>")
                st.markdown(
                    "<table class='proj-table'><thead><tr>"
                    "<th>Month</th><th>Monthly Revenue</th><th>Avg. Ticket</th><th>Annual Revenue</th>"
                    "</tr></thead><tbody>" + linhas_html + "</tbody></table>",
                    unsafe_allow_html=True)

    with tabs[2]:
        st.header("Operational Churn Analysis")
        churn_operacional_value, churn_operacional_percentage = "-", "-"
        if churn_view == 'Executed' and not df_backlog_raw.empty and selected_years and selected_months and not df_active_raw.empty:
            current_year_for_backlog = max(selected_years) if selected_years else datetime.now().year
            df_backlog_filtered_op = df_backlog_raw[(df_backlog_raw['Ano Backlog'] == current_year_for_backlog) & (df_backlog_raw['Tipo de Cliente Backlog'].isin(selected_client_types)) & (df_backlog_raw['Tipo de Churn Backlog'].isin(selected_churn_types))]
            if len(selected_months) == 1:
                current_month_num = month_order_num_pt.index(selected_months[0]) + 1
                backlog_current_month = df_backlog_filtered_op[df_backlog_filtered_op['Mes Backlog'] == current_month_num]['Volume Backlog'].sum()
                backlog_previous_month = 0
                if current_month_num == 1:
                    prev_year_for_backlog = current_year_for_backlog - 1
                    backlog_previous_month = df_backlog_raw[(df_backlog_raw['Ano Backlog'] == prev_year_for_backlog) & (df_backlog_raw['Mes Backlog'] == 12) & (df_backlog_raw['Tipo de Cliente Backlog'].isin(selected_client_types)) & (df_backlog_raw['Tipo de Churn Backlog'].isin(selected_churn_types))]['Volume Backlog'].sum()
                else:
                    backlog_previous_month = df_backlog_filtered_op[df_backlog_filtered_op['Mes Backlog'] == current_month_num - 1]['Volume Backlog'].sum()
                delta_backlog = backlog_current_month - backlog_previous_month
                churn_volume_current_month = df_filtered[(df_filtered['Ano Churn'] == current_year_for_backlog) & (df_filtered['Mes Churn'] == current_month_num)]['Volume'].sum()
                churn_operacional_value = churn_volume_current_month + delta_backlog
                active_base_current_month = df_active_raw[(df_active_raw['Ano Base Ativa'] == current_year_for_backlog) & (df_active_raw['Mes Base Ativa'] == current_month_num) & (df_active_raw['Tipo de Cliente Base Ativa'].isin(selected_client_types))]['Volume Base Ativa'].sum()
                if active_base_current_month > 0: churn_operacional_percentage = (churn_operacional_value/active_base_current_month)*100
            elif len(selected_months) > 1:
                total_churn_operacional_period, total_active_base_period = 0, 0
                sorted_selected_month_nums = sorted([month_order_num_pt.index(m)+1 for m in selected_months])
                for idx, current_month_num in enumerate(sorted_selected_month_nums):
                    current_month_backlog_data = df_backlog_raw[(df_backlog_raw['Ano Backlog'] == current_year_for_backlog) & (df_backlog_raw['Mes Backlog'] == current_month_num) & (df_backlog_raw['Tipo de Cliente Backlog'].isin(selected_client_types)) & (df_backlog_raw['Tipo de Churn Backlog'].isin(selected_churn_types))]['Volume Backlog'].sum()
                    prev_month_backlog_data = 0
                    if idx == 0 and current_month_num == 1:
                        prev_year_for_backlog = current_year_for_backlog - 1
                        prev_month_backlog_data = df_backlog_raw[(df_backlog_raw['Ano Backlog'] == prev_year_for_backlog) & (df_backlog_raw['Mes Backlog'] == 12) & (df_backlog_raw['Tipo de Cliente Backlog'].isin(selected_client_types)) & (df_backlog_raw['Tipo de Churn Backlog'].isin(selected_churn_types))]['Volume Backlog'].sum()
                    elif idx > 0:
                        prev_month_num = sorted_selected_month_nums[idx-1]
                        prev_month_backlog_data = df_backlog_raw[(df_backlog_raw['Ano Backlog'] == current_year_for_backlog) & (df_backlog_raw['Mes Backlog'] == prev_month_num) & (df_backlog_raw['Tipo de Cliente Backlog'].isin(selected_client_types)) & (df_backlog_raw['Tipo de Churn Backlog'].isin(selected_churn_types))]['Volume Backlog'].sum()
                    delta_backlog_month = current_month_backlog_data - prev_month_backlog_data
                    churn_volume_month = df_filtered[(df_filtered['Ano Churn'] == current_year_for_backlog) & (df_filtered['Mes Churn'] == current_month_num)]['Volume'].sum()
                    total_churn_operacional_period += (churn_volume_month + delta_backlog_month)
                    total_active_base_period += df_active_raw[(df_active_raw['Ano Base Ativa'] == current_year_for_backlog) & (df_active_raw['Mes Base Ativa'] == current_month_num) & (df_active_raw['Tipo de Cliente Base Ativa'].isin(selected_client_types))]['Volume Base Ativa'].sum()
                churn_operacional_value = total_churn_operacional_period
                if total_active_base_period > 0: churn_operacional_percentage = (churn_operacional_value/total_active_base_period)*100
        else:
            st.warning("Operational Churn analysis is only available for the 'Executed' view and requires Backlog data, Active Base data, and selected year/month filters.")
        
        if isinstance(churn_operacional_percentage, (int, float)):
            display_perc_co = f"{churn_operacional_percentage:.2f}%".replace(".", ",")
            st.metric(label=f"Operational Churn Rate ({', '.join(selected_months)} {', '.join(map(str, selected_years))})", value=display_perc_co)
        else:
            st.metric(label=f"Operational Churn Rate ({', '.join(selected_months)} {', '.join(map(str, selected_years))})", value="-")

    with tabs[3]:
        st.header(f"Distribution by Client Type")
        pie_color_map = {'Large Corporate': '#0A2A66', 'Medium Corporate': '#1E5FCC', 'House Hold': '#60A5FA', 'Other': '#A9C2EB'}
        df_churn_prev = df_filtered[df_filtered['Ano Churn'] == 2025]
        df_churn_curr = df_filtered[df_filtered['Ano Churn'] == 2026]
        df_plot_client_type_prev = df_churn_prev.groupby('Tipo de Cliente').agg(Volume_Churn=('Volume', 'sum')).reset_index()
        df_plot_client_type_curr = df_churn_curr.groupby('Tipo de Cliente').agg(Volume_Churn=('Volume', 'sum')).reset_index()
        col_prev, col_curr, col_comparison = st.columns([1, 1, 1])
        with col_prev:
            st.markdown("<h4 style='text-align: center; font-weight: bold; color: #5a6a79;'>Consolidated 2025</h4>", unsafe_allow_html=True)
            if not df_plot_client_type_prev.empty:
                fig_client_type_prev = px.pie(df_plot_client_type_prev, values="Volume_Churn", names="Tipo de Cliente", hole=0.5, color="Tipo de Cliente", color_discrete_map=pie_color_map)
                fig_client_type_prev.update_layout(showlegend=False)
                fig_client_type_prev.update_traces(textinfo="percent+label", pull=[0.05] * len(df_plot_client_type_prev))
                st.plotly_chart(aplicar_tema_moderno(fig_client_type_prev), use_container_width=True)
            else: st.info("No churn data for 2025 with the selected filters.")
        with col_curr:
            st.markdown("<h4 style='text-align: center; font-weight: bold; color: #5a6a79;'>Consolidated 2026</h4>", unsafe_allow_html=True)
            if not df_plot_client_type_curr.empty:
                fig_client_type_curr = px.pie(df_plot_client_type_curr, values="Volume_Churn", names="Tipo de Cliente", hole=0.5, color="Tipo de Cliente", color_discrete_map=pie_color_map)
                fig_client_type_curr.update_layout(showlegend=False)
                fig_client_type_curr.update_traces(textinfo="percent+label", pull=[0.05] * len(df_plot_client_type_curr))
                st.plotly_chart(aplicar_tema_moderno(fig_client_type_curr), use_container_width=True)
            else: st.info("No churn data for 2026 with the selected filters.")
        with col_comparison:
            st.markdown("<h4 style='text-align: center; font-weight: bold; color: #5a6a79;'>Annual Variation (26 vs 25)</h4>", unsafe_allow_html=True)
            df_comparison_client_type = pd.merge(df_plot_client_type_prev.rename(columns={'Volume_Churn': 'Volume_2025'}), df_plot_client_type_curr.rename(columns={'Volume_Churn': 'Volume_2026'}), on='Tipo de Cliente', how='outer').fillna(0)
            df_comparison_client_type['Diferenca_Absoluta'] = df_comparison_client_type['Volume_2026'] - df_comparison_client_type['Volume_2025']
            df_comparison_client_type['Diferenca_Percentual'] = df_comparison_client_type.apply(lambda row: ((row['Volume_2026'] - row['Volume_2025']) / row['Volume_2025']) * 100 if row['Volume_2025'] != 0 else (100 if row['Volume_2026'] > 0 else 0), axis=1)
            if not df_comparison_client_type.empty:
                for index, row in df_comparison_client_type.iterrows():
                    diff_abs, diff_perc, vol_prev_val, vol_curr_val = row['Diferenca_Absoluta'], row['Diferenca_Percentual'], row['Volume_2025'], row['Volume_2026']
                    delta_color_class = "positive" if diff_abs < 0 else "negative"
                    delta_symbol = "▼" if diff_abs < 0 else "▲"
                    st.markdown(f"""<div class="kpi-container"><div class="kpi-title">{row['Tipo de Cliente']}</div><div class="kpi-value {delta_color_class}">{int(diff_abs):+,.0f}</div><div class="kpi-delta {delta_color_class}">{delta_symbol} {diff_perc:.1f}%</div><div class="kpi-sub-value"><b>2026:</b> {int(vol_curr_val):,.0f} | <b>2025:</b> {int(vol_prev_val):,.0f}</div></div><div style="margin-top: 10px;"></div>""".replace(",", "."), unsafe_allow_html=True)
            else: st.info("No variation to display.")
            
        st.header("Monthly Churn Evolution by Client Type (2026)")
        if churn_view == 'Executed' and not df_active_raw.empty and 2026 in df_filtered['Ano Churn'].unique():
            df_churn_rate_segmento = df_filtered[df_filtered['Ano Churn'] == 2026].groupby(['Mes Churn', 'Nome Mes Churn', 'Tipo de Cliente']).agg(Volume_Churn=('Volume', 'sum')).reset_index()
            df_total_active = df_active_raw[(df_active_raw['Ano Base Ativa'] == 2026) & (df_active_raw['Tipo de Cliente Base Ativa'].isin(selected_client_types))].groupby(['Mes Base Ativa']).agg(Volume_Base=('Volume Base Ativa', 'sum')).reset_index()
            df_segmento_merged = pd.merge(df_churn_rate_segmento, df_total_active, left_on='Mes Churn', right_on='Mes Base Ativa', how='left')
            df_segmento_merged['Churn_Rate'] = (df_segmento_merged['Volume_Churn'] / df_segmento_merged['Volume_Base']) * 100
            
            # --- REMOVIDO INPLACE=TRUE para segurança ---
            df_segmento_merged = df_segmento_merged.rename(columns={'Tipo de Cliente': 'Categoria'})
            
            df_segmento_merged = df_segmento_merged[['Mes Churn', 'Nome Mes Churn', 'Categoria', 'Volume_Churn', 'Churn_Rate']].dropna(subset=['Churn_Rate'])
            
            # --- USO DE .MAP EM VEZ DE .APPLY ---
            df_segmento_merged['text_label'] = df_segmento_merged['Churn_Rate'].map(lambda x: f'{x:.2f}%'.replace('.', ','))
            
            df_total_churn = df_filtered[df_filtered['Ano Churn'] == 2026].groupby(['Mes Churn', 'Nome Mes Churn']).agg(Volume_Churn=('Volume', 'sum')).reset_index()
            df_total_merged = pd.merge(df_total_churn, df_total_active, left_on='Mes Churn', right_on='Mes Base Ativa', how='left')
            df_total_merged['Churn_Rate'] = (df_total_merged['Volume_Churn'] / df_total_merged['Volume_Base']) * 100
            df_total_merged['Categoria'] = 'EXECUTED CHURN'
            df_total_merged = df_total_merged[['Mes Churn', 'Nome Mes Churn', 'Categoria', 'Volume_Churn', 'Churn_Rate']].dropna(subset=['Churn_Rate'])
            
            if not df_segmento_merged.empty:
                line_color_map = {'House Hold': '#60A5FA', 'Medium Corporate': '#1E5FCC', 'Large Corporate': '#0A2A66', 'Other': '#A9C2EB'}
                
                # --- CORREÇÃO DO DICIONÁRIO DE LABELS PARA EVITAR DUPLICIDADE EM BRANCO NO PLOTLY ---
                fig_line_rate = px.line(
                    df_segmento_merged, 
                    x='Nome Mes Churn', 
                    y='Churn_Rate', 
                    color='Categoria', 
                    text='text_label', 
                    markers=True, 
                    labels={'Nome Mes Churn': 'Month', 'Churn_Rate': '% of Total Active Base', 'Categoria': 'Segment'}, 
                    category_orders={"Nome Mes Churn": month_order_num_pt}, 
                    color_discrete_map=line_color_map
                )
                
                # --- REMOVENDO TITULOS DO EIXO/LEGENDA PARA MANTER O MESMO VISUAL ANTERIOR ---
                fig_line_rate.update_layout(xaxis_title="", legend_title="")
                
                fig_line_rate.update_traces(textposition='top center', textfont=dict(size=10), hovertemplate="<b>%{x}</b><br>Percentage: %{y:.2f}%<extra></extra>")
                max_rate = df_segmento_merged['Churn_Rate'].max() if not df_segmento_merged.empty else 1
                line_y_position = max_rate * 1.3
                text_y_position = max_rate * 1.5 
                fig_line_rate.add_shape(type="line", x0=df_segmento_merged['Nome Mes Churn'].unique()[0], x1=df_segmento_merged['Nome Mes Churn'].unique()[-1], y0=line_y_position, y1=line_y_position, line=dict(color="#d62728", width=2))
                for index, row in df_total_merged.iterrows():
                    rate_str = f"{row['Churn_Rate']:.2f}%".replace('.', ',')
                    annotation_text = f"<b>{row['Volume_Churn']}</b><br>({rate_str})"
                    fig_line_rate.add_annotation(x=row['Nome Mes Churn'], y=text_y_position, text=annotation_text, showarrow=False, font=dict(color="#d62728", size=12))
                fig_line_rate.update_layout(yaxis_range=[0, max_rate * 1.8], legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
                st.plotly_chart(aplicar_tema_moderno(fig_line_rate), use_container_width=True)
                
                df_table_data = pd.concat([df_segmento_merged, df_total_merged], ignore_index=True)
                df_pivot = df_table_data.pivot_table(index='Categoria', columns='Nome Mes Churn', values='Churn_Rate')
                ordered_months_in_data = [m for m in month_order_num_pt if m in df_pivot.columns]
                df_pivot = df_pivot[ordered_months_in_data]
                order_index = ['EXECUTED CHURN', 'Individual', 'SME', 'Corporate']
                valid_order_index = [idx for idx in order_index if idx in df_pivot.index]
                df_pivot = df_pivot.reindex(valid_order_index)
                
                # --- CORREÇÃO APLICADA AQUI (.map ao invés de .applymap) ---
                df_formatted = df_pivot.map(lambda x: f'{x:.2f}%'.replace('.', ',') if pd.notna(x) else '-')
                
                st.dataframe(df_formatted, use_container_width=True)
            else: st.info("Could not compute the percentages for the selected filters.")
        else: st.info("This analysis is only available for the 'Executed' view with year 2026 selected and requires active customer base data.")

    with tabs[4]:
        st.header(f"Monthly Churn Volume by Type")
        if 'Tipo de Churn' in df_filtered.columns and not df_filtered['Tipo de Churn'].isnull().all() and df_filtered['Tipo de Churn'].nunique() > 0:
            df_plot_churn_type_monthly = df_filtered.groupby(['Ano Churn', 'Mes Churn', 'Nome Mes Churn', 'Tipo de Churn']).agg(Volume_Churn=('Volume', 'sum')).reset_index()
            df_plot_churn_type_monthly['Nome Mes Abreviado'] = df_plot_churn_type_monthly['Nome Mes Churn'].map(month_to_abbr_map)
            fig_churn_type_monthly_stacked = px.bar(df_plot_churn_type_monthly, x="Nome Mes Abreviado", y="Volume_Churn", color="Tipo de Churn", facet_col="Ano Churn", barmode="stack", labels={"Nome Mes Abreviado": "Month", "Volume_Churn": "Churn Volume", "Ano Churn": "Year", "Tipo de Churn": "Churn Type"}, category_orders={"Nome Mes Abreviado": month_abbr_order_pt}, text_auto=True)
            fig_churn_type_monthly_stacked.update_traces(textposition='inside', textfont=dict(color='white', weight='bold', size=12), textangle=0)
            fig_churn_type_monthly_stacked.for_each_annotation(lambda a: a.update(text=a.text.replace("Ano Churn=", "")))
            fig_churn_type_monthly_stacked.update_layout(legend=dict(orientation="h", yanchor="bottom", y=-0.3, xanchor="center", x=0.5))
            st.plotly_chart(aplicar_tema_moderno(fig_churn_type_monthly_stacked), use_container_width=True)
        else: st.warning("No 'Churn Type' data to display the stacked monthly chart.")

    with tabs[5]:
        st.header("Cancellation Reasons by Year (26 vs 25)")
        sub_col = 'Churn Sub-Reason (Global)' if 'Churn Sub-Reason (Global)' in df_filtered.columns else 'Categoria4_Motivo'
        rea_col = 'Churn Reason (Global)'
        has_reason = rea_col in df_filtered.columns
        group_cols = ([rea_col] if has_reason else []) + [sub_col]
        if sub_col in df_filtered.columns and not df_filtered[sub_col].isnull().all():
            df_curr_clean = df_filtered[df_filtered['Ano Churn'] == 2026].copy()
            df_curr_clean = df_curr_clean[~df_curr_clean[sub_col].astype(str).str.strip().str.lower().isin(['', 'nan', 'desconsiderar', 'none'])]
            summary_curr = df_curr_clean.groupby(group_cols).agg(Volume_Current=('Volume','sum')).reset_index()
            df_prev_clean = df_filtered[df_filtered['Ano Churn'] == 2025].copy()
            df_prev_clean = df_prev_clean[~df_prev_clean[sub_col].astype(str).str.strip().str.lower().isin(['', 'nan', 'desconsiderar', 'none'])]
            summary_prev = df_prev_clean.groupby(group_cols).agg(Volume_Previous=('Volume','sum')).reset_index()
            if not summary_curr.empty or not summary_prev.empty:
                df_combined_reasons = pd.merge(summary_curr, summary_prev, on=group_cols, how='outer').fillna(0)
                tot_c = df_combined_reasons['Volume_Current'].sum(); tot_p = df_combined_reasons['Volume_Previous'].sum()
                df_combined_reasons['Percentual_Current'] = (df_combined_reasons['Volume_Current'] / tot_c) * 100 if tot_c > 0 else 0
                df_combined_reasons['Percentual_Previous'] = (df_combined_reasons['Volume_Previous'] / tot_p) * 100 if tot_p > 0 else 0
                df_combined_reasons['Variation 26 vs 25'] = df_combined_reasons.apply(lambda row: ((row['Volume_Current'] / row['Volume_Previous']) - 1) if row['Volume_Previous'] > 0 else (float('inf') if row['Volume_Current'] > 0 else 0), axis=1)
                df_combined_reasons = df_combined_reasons.sort_values('Volume_Current', ascending=False)
                rename_map = {sub_col: 'Cancellation Sub-Reason', 'Volume_Current': 'Volume 2026', 'Percentual_Current': '% 2026', 'Volume_Previous': 'Volume 2025', 'Percentual_Previous': '% 2025'}
                if has_reason: rename_map[rea_col] = 'Cancellation Reason'
                df_display = df_combined_reasons.rename(columns=rename_map)
                for col in ['Volume 2026', 'Volume 2025']: df_display[col] = df_display[col].astype(int)
                for col in ['% 2026', '% 2025']: df_display[col] = df_display[col].map(lambda x: f"{x:.2f}%".replace(".", ","))
                df_display['Variation 26 vs 25'] = df_display['Variation 26 vs 25'].apply(lambda x: f"{x:.2%}".replace('.', ',') if pd.notna(x) and x != float('inf') else ("New Reason" if x == float('inf') else "0,00%"))
                cols_show = (['Cancellation Reason'] if has_reason else []) + ['Cancellation Sub-Reason', 'Volume 2026', '% 2026', 'Volume 2025', '% 2025', 'Variation 26 vs 25']
                st.dataframe(df_display[cols_show], use_container_width=True, hide_index=True)
            else: st.info("No cancellation reason to display with the selected filters.")
        else: st.info("No cancellation reason data found.")

    with tabs[6]:
        st.header(f"Churn by Franchise (26 vs 25)")
        if 'Filial' in df_filtered.columns and not df_filtered['Filial'].isnull().all():
            df_curr_clean = df_filtered[df_filtered['Ano Churn'] == 2026].copy()
            df_curr_clean = df_curr_clean[~df_curr_clean['Filial'].astype(str).str.strip().str.lower().isin(['', 'nan'])]
            summary_f_curr = df_curr_clean.groupby('Filial').agg(Volume_Current=('Volume','sum')).reset_index()
            df_prev_clean = df_filtered[df_filtered['Ano Churn'] == 2025].copy()
            df_prev_clean = df_prev_clean[~df_prev_clean['Filial'].astype(str).str.strip().str.lower().isin(['', 'nan'])]
            summary_f_prev = df_prev_clean.groupby('Filial').agg(Volume_Previous=('Volume','sum')).reset_index()
            if not summary_f_curr.empty or not summary_f_prev.empty:
                df_combined_franchises = pd.merge(summary_f_curr, summary_f_prev, on='Filial', how='outer').fillna(0)
                df_combined_franchises['Percentual_Current'] = (df_combined_franchises['Volume_Current'] / df_combined_franchises['Volume_Current'].sum()) * 100 if df_combined_franchises['Volume_Current'].sum() > 0 else 0
                df_combined_franchises['Percentual_Previous'] = (df_combined_franchises['Volume_Previous'] / df_combined_franchises['Volume_Previous'].sum()) * 100 if df_combined_franchises['Volume_Previous'].sum() > 0 else 0
                df_combined_franchises['Variation 26 vs 25'] = df_combined_franchises.apply(lambda row: ((row['Volume_Current'] / row['Volume_Previous']) - 1) if row['Volume_Previous'] > 0 else (float('inf') if row['Volume_Current'] > 0 else 0), axis=1)
                df_display = df_combined_franchises.rename(columns={'Filial': 'Franchise', 'Volume_Current': 'Volume 2026', 'Percentual_Current': '% 2026', 'Volume_Previous': 'Volume 2025', 'Percentual_Previous': '% 2025'})
                for col in ['Volume 2026', 'Volume 2025']: df_display[col] = df_display[col].astype(int)
                for col in ['% 2026', '% 2025']: df_display[col] = df_display[col].map(lambda x: f"{x:.2f}%".replace(".", ","))
                df_display['Variation 26 vs 25'] = df_display['Variation 26 vs 25'].apply(lambda x: f"{x:.2%}".replace('.', ',') if pd.notna(x) and x != float('inf') else ("New Franchise" if x == float('inf') else "0,00%"))
                st.dataframe(df_display[['Franchise', 'Volume 2026', '% 2026', 'Volume 2025', '% 2025', 'Variation 26 vs 25']], use_container_width=True, hide_index=True)
            else: st.info("No franchise to display with the selected filters.")
        else: st.info("No franchise data found.")

    with tabs[7]:
        st.header("Churn by Product (26 vs 25)")
        if 'Família' in df_filtered.columns and not df_filtered['Família'].isnull().all():
            df_curr_clean = df_filtered[df_filtered['Ano Churn'] == 2026].copy()
            df_curr_clean = df_curr_clean[~df_curr_clean['Família'].astype(str).str.strip().str.lower().isin(['', 'nan'])]
            summary_familia_curr = df_curr_clean.groupby('Família').agg(Volume_Current=('Volume', 'sum')).reset_index()
            df_prev_clean = df_filtered[df_filtered['Ano Churn'] == 2025].copy()
            df_prev_clean = df_prev_clean[~df_prev_clean['Família'].astype(str).str.strip().str.lower().isin(['', 'nan'])]
            summary_familia_prev = df_prev_clean.groupby('Família').agg(Volume_Previous=('Volume', 'sum')).reset_index()
            if not summary_familia_curr.empty or not summary_familia_prev.empty:
                df_combined_familia = pd.merge(summary_familia_curr, summary_familia_prev, on='Família', how='outer').fillna(0)
                total_vol_curr = df_combined_familia['Volume_Current'].sum()
                total_vol_prev = df_combined_familia['Volume_Previous'].sum()
                df_combined_familia['Percentual_Current'] = (df_combined_familia['Volume_Current'] / total_vol_curr) * 100 if total_vol_curr > 0 else 0
                df_combined_familia['Percentual_Previous'] = (df_combined_familia['Volume_Previous'] / total_vol_prev) * 100 if total_vol_prev > 0 else 0
                df_combined_familia['Variation 26 vs 25'] = df_combined_familia.apply(lambda row: ((row['Volume_Current'] / row['Volume_Previous']) - 1) if row['Volume_Previous'] > 0 else (float('inf') if row['Volume_Current'] > 0 else 0), axis=1)
                df_display = df_combined_familia.rename(columns={'Família': 'Product Family', 'Volume_Current': 'Volume 2026', 'Percentual_Current': '% 2026', 'Volume_Previous': 'Volume 2025', 'Percentual_Previous': '% 2025'})
                for col in ['Volume 2026', 'Volume 2025']: df_display[col] = df_display[col].astype(int)
                for col in ['% 2026', '% 2025']: df_display[col] = df_display[col].map(lambda x: f"{x:.2f}%".replace(".", ","))
                df_display['Variation 26 vs 25'] = df_display['Variation 26 vs 25'].apply(lambda x: f"{x:.2%}".replace('.', ',') if pd.notna(x) and x != float('inf') else ("New Family" if x == float('inf') else "0,00%"))
                st.dataframe(df_display[['Product Family', 'Volume 2026', '% 2026', 'Volume 2025', '% 2025', 'Variation 26 vs 25']], use_container_width=True, hide_index=True)
            else: st.info("No product family to display with the selected filters.")
        else: st.info("No 'Family' data found in the churn files for this analysis.")

    with tabs[8]:
        st.header("Financial Impact Analysis (Churn)")
        st.info("ℹ️ **Impact calculation:** The displayed value considers revenue lost cumulatively until year-end (e.g., churn in Jan = Value x 12 | churn in Feb = Value x 11).")
        def calculate_revenue_metrics(df_in, group_col):
            df_calc = df_in.copy()
            df_calc['Mes Churn'] = pd.to_numeric(df_calc['Mes Churn'], errors='coerce').fillna(0).astype(int)
            df_calc['Receita_Impacto'] = df_calc['Preçovalidado'] * (13 - df_calc['Mes Churn'])
            df_grouped = df_calc.groupby(group_col).agg(Volume=('Volume', 'sum'), Receita_Nominal_Total=('Preçovalidado', 'sum'), Receita_Impacto_Total=('Receita_Impacto', 'sum')).reset_index()
            df_grouped['TKM'] = df_grouped['Receita_Nominal_Total'] / df_grouped['Volume']
            return df_grouped
        st.subheader("Financial Impact by Client Type")
        df_curr_clean = df_filtered[df_filtered['Ano Churn'] == 2026].copy()
        df_prev_clean = df_filtered[df_filtered['Ano Churn'] == 2025].copy()
        if not df_curr_clean.empty or not df_prev_clean.empty:
            metrics_curr = calculate_revenue_metrics(df_curr_clean, 'Tipo de Cliente')
            metrics_prev = calculate_revenue_metrics(df_prev_clean, 'Tipo de Cliente')
            df_combined = pd.merge(metrics_curr, metrics_prev, on='Tipo de Cliente', how='outer', suffixes=('_2026', '_2025')).fillna(0)
            total_impact_2026 = df_combined['Receita_Impacto_Total_2026'].sum()
            df_combined['% Impact 2026'] = (df_combined['Receita_Impacto_Total_2026'] / total_impact_2026) * 100 if total_impact_2026 > 0 else 0
            df_display = df_combined[['Tipo de Cliente', 'Receita_Impacto_Total_2026', '% Impact 2026', 'TKM_2026', 'Receita_Impacto_Total_2025', 'TKM_2025']].copy()
            df_display.rename(columns={'Receita_Impacto_Total_2026': 'Total Impact 2026', 'TKM_2026': 'TKM 2026', 'Receita_Impacto_Total_2025': 'Total Impact 2025', 'TKM_2025': 'TKM 2025'}, inplace=True)
            for col in ['Total Impact 2026', 'Total Impact 2025', 'TKM 2026', 'TKM 2025']: df_display[col] = df_display[col].map(format_BRL)
            df_display['% Impact 2026'] = df_display['% Impact 2026'].map(lambda x: f"{x:.2f}%".replace(".", ","))
            st.dataframe(df_display, use_container_width=True, hide_index=True)
        else: st.info("No data to display in this view.")
        st.markdown("---")
        st.subheader("Financial Impact by Cancellation Reason")
        if 'Categoria4_Motivo' in df_filtered.columns:
            df_curr_motivo = df_curr_clean[~df_curr_clean['Categoria4_Motivo'].astype(str).str.strip().str.lower().isin(['', 'nan', 'desconsiderar', 'none'])]
            df_prev_motivo = df_prev_clean[~df_prev_clean['Categoria4_Motivo'].astype(str).str.strip().str.lower().isin(['', 'nan', 'desconsiderar', 'none'])]
            if not df_curr_motivo.empty or not df_prev_motivo.empty:
                metrics_curr_m = calculate_revenue_metrics(df_curr_motivo, 'Categoria4_Motivo')
                metrics_prev_m = calculate_revenue_metrics(df_prev_motivo, 'Categoria4_Motivo')
                df_combined_m = pd.merge(metrics_curr_m, metrics_prev_m, on='Categoria4_Motivo', how='outer', suffixes=('_2026', '_2025')).fillna(0)
                df_combined_m = df_combined_m.sort_values(by='Receita_Impacto_Total_2026', ascending=False)
                total_impact_m_2026 = df_combined_m['Receita_Impacto_Total_2026'].sum()
                df_combined_m['% Impact 2026'] = (df_combined_m['Receita_Impacto_Total_2026'] / total_impact_m_2026) * 100 if total_impact_m_2026 > 0 else 0
                df_display_m = df_combined_m[['Categoria4_Motivo', 'Receita_Impacto_Total_2026', '% Impact 2026', 'TKM_2026', 'Receita_Impacto_Total_2025']].copy()
                df_display_m.rename(columns={'Categoria4_Motivo': 'Reason', 'Receita_Impacto_Total_2026': 'Total Impact 2026', 'TKM_2026': 'TKM 2026', 'Receita_Impacto_Total_2025': 'Total Impact 2025'}, inplace=True)
                for col in ['Total Impact 2026', 'Total Impact 2025', 'TKM 2026']: df_display_m[col] = df_display_m[col].map(format_BRL)
                df_display_m['% Impact 2026'] = df_display_m['% Impact 2026'].map(lambda x: f"{x:.2f}%".replace(".", ","))
                st.dataframe(df_display_m, use_container_width=True, hide_index=True)
            else: st.info("No cancellation reason data found.")
        else: st.info("Reason column not found.")

    with tabs[9]:
        st.header("Advanced Analytics")
        if check_password():
            st.caption("The analyses below use the filters selected in the sidebar (Client Type and Churn Type).")
            st.write("---")
            with st.expander("Strategic Diagnosis (2026 vs 2025)"):
                st.markdown("""**What it does?** This analysis examines churn variation between 2025 and 2026 to surface the most critical insights.""")
                if st.button("Generate Prompt for Strategic Analysis", key="analysis_button"):
                    with st.spinner("Preparing the prompt with the data..."):
                        if not df_filtered.empty:
                            if 2025 in df_filtered['Ano Churn'].unique() and 2026 in df_filtered['Ano Churn'].unique():
                                generated_prompt = get_ai_analysis(df_filtered, churn_view, selected_years, selected_months, selected_client_types, all_client_types, selected_churn_types, all_churn_types)
                                st.info("Copy the prompt below and paste it into your preferred AI tool.")
                                st.text_area("Generated prompt:", value=generated_prompt, height=400)
                            else: st.warning("For the strategic diagnosis, please select years 2025 and 2026 in the sidebar filters.")
                        else: st.warning("No data to analyze with the current filters.")
            
            with st.expander("Detailed Period Comparison"):
                st.markdown("""**What it does?** Lets you compare churn between two specific months/years of your choice.""")
                analysis_years = sorted(df_churn['Ano Churn'].unique(), reverse=True)
                analysis_months = month_order_num_pt
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown("**Period 1 (Base)**")
                    year1_select = st.selectbox("Year", options=analysis_years, key="year1_select", index=1 if len(analysis_years) > 1 else 0)
                    month1_select = st.selectbox("Month", options=analysis_months, key="month1_select", index=0)
                with c2:
                    st.markdown("**Period 2 (Comparison)**")
                    year2_select = st.selectbox("Year", options=analysis_years, key="year2_select", index=0)
                    month2_select = st.selectbox("Month", options=analysis_months, key="month2_select", index=0)
                
                if st.button("Generate Prompt for Comparison", key="custom_compare_button"):
                    if year1_select == year2_select and month1_select == month2_select: 
                        st.warning("Please select two different periods for the comparison.")
                    else:
                        with st.spinner("Preparing the prompt with the data..."):
                            generated_prompt = get_ai_custom_comparison(
                                df_churn, 
                                df_backlog_raw, 
                                year1_select, 
                                month1_select, 
                                year2_select, 
                                month2_select, 
                                selected_client_types, 
                                selected_churn_types, 
                                churn_view
                            )
                            st.info("Copy the prompt below and paste it into your preferred AI tool.")
                            st.text_area("Generated prompt:", value=generated_prompt, height=450)

            with st.expander("Root Cause Analysis"):
                st.markdown("""**What it does?** Go beyond the numbers and dig deep into a specific problem.""")
                if 'Categoria4_Motivo' in df_filtered.columns and not df_filtered['Categoria4_Motivo'].isnull().all():
                    available_reasons = sorted([r for r in df_filtered['Categoria4_Motivo'].unique() if pd.notna(r) and str(r).strip() not in ['', 'nan']])
                    if available_reasons:
                        col_raiz1, col_raiz2 = st.columns(2)
                        with col_raiz1: reason_to_investigate = st.selectbox("Select the reason to investigate:", options=available_reasons, key="raiz_reason")
                        with col_raiz2: segment_to_investigate = st.selectbox("Select the segment to focus on:", options=["All"] + all_client_types, key="raiz_segment")
                        if st.button("Generate Root Cause Prompt", key="raiz_button"):
                            df_slice = df_filtered[df_filtered['Categoria4_Motivo'] == reason_to_investigate]
                            if segment_to_investigate != "All": df_slice = df_slice[df_slice['Tipo de Cliente'] == segment_to_investigate]
                            with st.spinner("Preparing the prompt with the data..."):
                                generated_prompt = get_ai_root_cause_deep_dive(df_slice, reason_to_investigate, segment_to_investigate, churn_view)
                                st.info("Copy the prompt below and paste it into your preferred AI tool.")
                                st.text_area("Generated prompt:", value=generated_prompt, height=400)
                    else: st.info("No cancellation reasons available for investigation with the current filters.")
                else: st.info("A cancellation reason column is required for this analysis.")
            with st.expander("Churn Lead Time Analysis"):
                st.markdown("""**What it does?** Measures the time (in days) between the creation date of a cancellation work order and the date the uninstall was completed.""")
                if churn_view == 'Executed':
                    df_lead_time = df_churn.copy()
                    df_lead_time['Data de Criacao da OS'] = pd.to_datetime(df_lead_time['Data de Criacao da OS'], errors='coerce')
                    df_lead_time['Data de Desinstalacao'] = pd.to_datetime(df_lead_time['Data de Desinstalacao'], errors='coerce')
                    df_lead_time.dropna(subset=['Data de Criacao da OS', 'Data de Desinstalacao'], inplace=True)
                    df_lead_time['Lead_Time_Days'] = (df_lead_time['Data de Desinstalacao'] - df_lead_time['Data de Criacao da OS']).dt.days
                    df_lead_time = df_lead_time[(df_lead_time['Lead_Time_Days'] >= 0) & (df_lead_time['Lead_Time_Days'] < 365)]
                    if not df_lead_time.empty:
                        avg_lead_time = df_lead_time['Lead_Time_Days'].mean()
                        median_lead_time = df_lead_time['Lead_Time_Days'].median()
                        col_lt1, col_lt2 = st.columns(2)
                        with col_lt1: st.metric("Average Lead Time", f"{avg_lead_time:.1f} days")
                        with col_lt2: st.metric("Median Lead Time", f"{median_lead_time:.1f} days")
                        st.subheader("Lead Time Distribution (days)")
                        fig_lead_time_hist = px.histogram(df_lead_time, x="Lead_Time_Days", nbins=50, labels={"Lead_Time_Days": "Days between WO creation and uninstall"}, title="Churn Lead Time Frequency")
                        st.plotly_chart(aplicar_tema_moderno(fig_lead_time_hist), use_container_width=True)
                        st.subheader("Prompt Generator for Qualitative Analysis")
                        if st.button("Generate Prompt for Lead Time Analysis", key="lead_time_ai_button"):
                            with st.spinner("Preparing the prompt with the data..."):
                                generated_prompt = get_ai_lead_time_analysis(df_lead_time)
                                st.info("Copy the prompt below and paste it into your preferred AI tool.")
                                st.text_area("Generated prompt:", value=generated_prompt, height=400)
                    else: st.warning("Not enough data to compute Lead Time with the current filters.")
                else: st.info("The Lead Time analysis is only available for the 'Executed' churn view.")
            with st.expander("Retention Impact Simulator (What-If)"):
                st.markdown("""**What it does?** Simulates the impact of a churn-reduction target on a specific segment.""")
                st.info("The calculation simulates the impact of the reduction target on the **total churn volume projected for the year**.")
                df_sim = df_filtered.copy()
                st.subheader("1. Define the Simulation Scenario")
                available_client_types = sorted(df_sim['Tipo de Cliente'].unique())
                options_sim = ["All"] + available_client_types
                target_client_type = st.selectbox("Select the client type to focus on", options=options_sim, key="sim_client_type")
                reduction_percentage = st.slider("Set the Churn Reduction target (%)", 0, 100, 10, key="sim_slider_cliente")
                base_churn_df = df_sim
                if target_client_type != "All": base_churn_df = df_sim[df_sim['Tipo de Cliente'] == target_client_type]
                projected_annual_churn_base = 0
                if all_years:
                    current_year_for_sim = max(all_years)
                    df_current_year_sim = base_churn_df[base_churn_df['Ano Churn'] == current_year_for_sim]
                    if not df_current_year_sim.empty:
                        num_months_data_sim = df_current_year_sim['Mes Churn'].nunique()
                        if num_months_data_sim > 0: projected_annual_churn_base = (df_current_year_sim['Volume'].sum() / num_months_data_sim) * 12
                if projected_annual_churn_base > 0:
                    absolute_reduction = projected_annual_churn_base * (reduction_percentage / 100)
                    projected_churn_new = projected_annual_churn_base - absolute_reduction
                    st.subheader("2. See the Impact on the Annual Projection")
                    sim_col1, sim_col2 = st.columns(2)
                    with sim_col1: st.metric(label=f"Current Projection ({target_client_type})", value=f"{int(projected_annual_churn_base)}")
                    with sim_col2: st.metric(label="New Projection (With Reduction)", value=f"{int(projected_churn_new)}", delta=f"{-int(absolute_reduction)} clients", delta_color="inverse")
                else: st.info("No churn data for the current year with the selected filters. Cannot compute the projection for the simulation.")
            st.markdown("---")
            st.caption("Insight quality depends on the amount and quality of available data.")


if __name__ == "__main__":
    main()
