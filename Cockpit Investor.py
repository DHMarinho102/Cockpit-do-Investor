# =============================================================================
# COCKPIT DO INVESTOR - ACELERADOR COMERCIAL COM PAINEL LATERAL & COPY WA
# =============================================================================

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from dash import Dash, dcc, html, Input, Output, State, dash_table, ctx
import os
import shutil
import datetime
import urllib.parse

import crm_engine as crm
import importlib
importlib.reload(crm)

# =============================================================================
# PALETA DE CORES INSTITUCIONAL
# =============================================================================
PALETTE = {
    'bg_dark': '#1A1A1A',
    'gold_main': '#AD8A4B',
    'gold_bright': '#C9A84C',
    'gold_sub': '#B7A46C',
    'gold_muted': '#8C8272',
    'bg_card': '#F7F6F3',
    'hairline': '#E7E4DC',
    'text_body': '#5C5C5C',
    'text_label': '#8C8C8C',
    'text_dark': '#1A1A1A',
    'green_soft': '#2E7D32',
    'red_soft': '#C53030',
    'red_bg_light': '#FFF5F5',
    'red_border_light': '#FEB2B2',
    'gold_scale': ['#8C8272', '#B7A46C', '#AD8A4B', '#C9A84C']
}

FONT_SERIF = "Georgia, 'Times New Roman', Cambria, serif"
FONT_SANS = "Calibri, Arial, sans-serif"

MESES_PT = {1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril', 5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto', 9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'}
hoje_dt = datetime.datetime.now()
NOME_MES_ATUAL = f"Mês Atual ({MESES_PT[hoje_dt.month]}/{hoje_dt.year})"

diretorio_atual = os.path.dirname(os.path.abspath(__file__))
caminho_excel = os.path.join(diretorio_atual, 'Base_Dashboard_Pronta.xlsx')
caminho_produtos_foco = os.path.join(diretorio_atual, 'Produtos_Foco.xlsx')

# Garantia de criação da planilha de produtos foco
if not os.path.exists(caminho_produtos_foco):
    df_foco_padrao = pd.DataFrame([
        {
            'Produto Foco': 'Paramis Arpeggio',
            'Fator de Risco': 'Pós-Fixado',
            'Tipo Veículo': 'Fundo Fechado',
            'Cap Compliance (% AUC)': 0.10,
            'Ticket Mínimo (R$)': 10000.0,
            'ROA (%)': 1.20,
            'Tese Comercial / Destaques': 'CDI+2%a.a. para 4 anos, com rendimento mensal isento de IR'
        },
        {
            'Produto Foco': 'NTN-B 2035',
            'Fator de Risco': 'Inflação',
            'Tipo Veículo': 'Renda Fixa Pública',
            'Cap Compliance (% AUC)': 1.00,
            'Ticket Mínimo (R$)': 5000.0,
            'ROA (%)': 0.25,
            'Tese Comercial / Destaques': 'Proteção contra inflação IPCA + taxa pré atrativa'
        },
        {
            'Produto Foco': 'Top Dividendos FIA',
            'Fator de Risco': 'Renda Variável',
            'Tipo Veículo': 'Ações / FIA',
            'Cap Compliance (% AUC)': 1.00,
            'Ticket Mínimo (R$)': 1000.0,
            'ROA (%)': 0.80,
            'Tese Comercial / Destaques': 'Foco em geração de renda passiva via proventos'
        }
    ])
    df_foco_padrao.to_excel(caminho_produtos_foco, index=False)

def ler_base_pronta(caminho):
    try:
        return pd.read_excel(caminho)
    except PermissionError:
        arquivo_temp = os.path.join(os.path.dirname(caminho), "temp_base.xlsx")
        shutil.copy2(caminho, arquivo_temp)
        df = pd.read_excel(arquivo_temp)
        if os.path.exists(arquivo_temp): os.remove(arquivo_temp)
        return df

def normalizar_conta(conta):
    if pd.isna(conta): return ''
    c_str = str(conta).strip()
    if c_str.endswith('.0'): c_str = c_str[:-2]
    return c_str

def converter_valor_br(val):
    if val is None or val == '': return 0.0
    if isinstance(val, (int, float)): return float(val)
    val_str = str(val).strip().replace('R$', '').replace('%', '').replace(' ', '')
    if '.' in val_str and ',' in val_str: val_str = val_str.replace('.', '').replace(',', '.')
    elif '.' in val_str and ',' not in val_str:
        parts = val_str.split('.')
        if len(parts[-1]) == 3: val_str = val_str.replace('.', '')
    elif ',' in val_str: val_str = val_str.replace(',', '.')
    try: return float(val_str)
    except ValueError: return 0.0

df_clean = ler_base_pronta(caminho_excel)

for col in ['Custódia', 'Gap Over', 'Gap Under', 'Aderência', 'LIQUIDEZ_TOTAL_360', 'Liquidez_Saldo', 'Liquidez_Fundos', 'Liquidez_RF']:
    if col in df_clean.columns: df_clean[col] = pd.to_numeric(df_clean[col], errors='coerce').fillna(0)

if 'Idade' not in df_clean.columns: df_clean['Idade'] = 0
if 'Modalidade' not in df_clean.columns: df_clean['Modalidade'] = 'Comission Based'
if 'Tipo_Pessoa' not in df_clean.columns: df_clean['Tipo_Pessoa'] = 'PESSOA FÍSICA'

df_clean['Cód. Conta'] = df_clean['Cód. Conta'].apply(normalizar_conta)
df_clean['Assessor'] = df_clean['Assessor'].fillna('Sem Código').astype(str).str.strip()
df_clean['Investor'] = df_clean['Investor'].fillna('Sem Investor').astype(str).str.strip()
df_clean['Tipo_Pessoa'] = df_clean['Tipo_Pessoa'].fillna('PESSOA FÍSICA').astype(str).str.upper()

df_liquidez_contas = df_clean[['Cód. Conta', 'Investor', 'Assessor', 'Tipo_Pessoa', 'Modalidade', 'LIQUIDEZ_TOTAL_360']].drop_duplicates(subset=['Cód. Conta'])

lista_classes_disponiveis = sorted([str(c) for c in df_clean['Fator de Risco'].dropna().unique() if str(c).strip() != ''])
lista_assessores = sorted([str(a) for a in df_clean['Assessor'].unique() if str(a) != 'Sem Código'])
lista_investors = sorted([str(inv) for inv in df_clean['Investor'].unique() if str(inv) != 'Sem Investor'])

money_format = {'locale': {'symbol': ['R$ ', ''], 'decimal': ',', 'group': '.'}, 'specifier': '$,.2f'}
pct_format = {'locale': {'symbol': ['', '%'], 'decimal': ',', 'group': '.'}, 'specifier': ',.2%'}
num_format = {'locale': {'symbol': ['', ''], 'decimal': ',', 'group': '.'}, 'specifier': ',.0f'}
formato_local_br = {'symbol': ['R$ ', ''], 'decimal': ',', 'group': '.'}

def formatar_moeda_br(valor):
    if pd.isna(valor) or valor == 0: return "R$ 0,00"
    return f"R$ {valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

def formatar_numero_br(valor):
    if pd.isna(valor) or valor == 0: return "0"
    return f"{valor:,.0f}".replace(',', 'X').replace('.', ',').replace('X', '.')

def formatar_percentual_br(valor):
    if pd.isna(valor): return "0,00%"
    return f"{valor * 100:,.2f}%".replace(',', 'X').replace('.', ',').replace('X', '.')

def calcular_aderencia_ponderada_df(df_target):
    if df_target.empty: return 0
    custodia_total = df_target['Custódia'].sum()
    if custodia_total == 0: return 0
    return (df_target['Custódia'] * df_target['Aderência']).sum() / custodia_total

estilo_btn_dourado = {
    'backgroundColor': PALETTE['gold_main'], 'color': 'white', 'border': 'none',
    'padding': '9px 18px', 'borderRadius': '6px', 'fontWeight': 'bold',
    'cursor': 'pointer', 'fontSize': '12px', 'fontFamily': FONT_SANS
}

def criar_card_kpi(titulo, valor, cor=PALETTE['text_dark']):
    return html.Div(style={
        'backgroundColor': PALETTE['bg_card'], 'padding': '18px 22px', 'borderRadius': '8px',
        'boxShadow': '0 4px 12px rgba(0,0,0,0.03)', 'border': f"1px solid {PALETTE['hairline']}",
        'textAlign': 'center', 'flex': '1', 'minWidth': '180px'
    }, children=[
        html.Span(titulo, style={
            'fontSize': '11px', 'fontWeight': 'bold', 'color': PALETTE['text_label'],
            'display': 'block', 'marginBottom': '8px', 'textTransform': 'uppercase',
            'letterSpacing': '0.8px', 'fontFamily': FONT_SANS
        }),
        html.Span(valor, style={
            'fontSize': '24px', 'fontWeight': 'bold', 'color': cor,
            'display': 'block', 'fontFamily': FONT_SERIF
        })
    ])

# =============================================================================
# INICIALIZAÇÃO DO APP E LAYOUT ESTRUTURAL
# =============================================================================
app = Dash(__name__, suppress_callback_exceptions=True)
app.title = "Cockpit do Investor"

estilo_tabela_comum = {'fontFamily': FONT_SANS, 'fontSize': '13px'}
tabs_styles = {'height': '46px', 'fontFamily': FONT_SANS}

tab_style = {
    'borderBottom': f"1px solid {PALETTE['hairline']}", 'padding': '10px 16px',
    'fontWeight': 'bold', 'backgroundColor': PALETTE['bg_card'], 'color': PALETTE['text_body'],
    'borderTop': '3px solid transparent'
}
tab_selected_style = {
    'borderTop': f"3px solid {PALETTE['gold_main']}", 'borderBottom': 'none',
    'backgroundColor': PALETTE['bg_dark'], 'color': PALETTE['gold_bright'],
    'padding': '10px 16px', 'fontWeight': 'bold'
}

app.layout = html.Div(id='topo-marcador-ancora', style={'backgroundColor': '#FFFFFF', 'padding': '25px', 'fontFamily': FONT_SANS}, children=[
    dcc.Store(id='store-crm-updated', data=0),
    dcc.Download(id="download-excel-geral-mesa"), dcc.Download(id="download-excel-geral-assessor"),
    dcc.Download(id="download-excel-classes"), dcc.Download(id="download-excel-clientes"),

    html.H1("Cockpit do Investor", style={
        'textAlign': 'center', 'color': PALETTE['bg_dark'], 'fontWeight': 'bold',
        'marginBottom': '20px', 'fontFamily': FONT_SERIF, 'letterSpacing': '-0.5px'
    }),
    html.Hr(style={'borderColor': PALETTE['hairline'], 'marginBottom': '25px'}),

    dcc.Tabs(id="abas-navegacao", value='aba-mesa', style=tabs_styles, children=[
        dcc.Tab(label='🏢 Visão Gerencial (Mesa / Investors)', value='aba-mesa', style=tab_style, selected_style=tab_selected_style),
        dcc.Tab(label='💼 Visão Tática (Assessor / Cliente)', value='aba-assessor', style=tab_style, selected_style=tab_selected_style),
        dcc.Tab(label='🎯 CRM Tático & Pipeline de Receita', value='aba-crm', style=tab_style, selected_style=tab_selected_style),
        dcc.Tab(label='🚀 Acelerador Comercial & Campanhas', value='aba-campanhas', style=tab_style, selected_style=tab_selected_style),
    ]),
    
    # CONTAINER ABA 1: MESA
    html.Div(id='container-aba-mesa', style={'marginTop': '25px', 'display': 'block'}, children=[
        html.Div(style={'display': 'flex', 'justify': 'space-between', 'alignItems': 'center', 'marginBottom': '20px', 'flexWrap': 'wrap', 'gap': '12px'}, children=[
            html.Div(style={'display': 'flex', 'gap': '15px', 'flexGrow': '1', 'flexWrap': 'wrap'}, children=[
                html.Div([
                    html.Label("Filtro Investor:", style={'fontWeight': 'bold', 'marginRight': '8px', 'color': PALETTE['text_body']}),
                    dcc.Dropdown(id='mesa-investor-dropdown', options=[{'label': 'Todos', 'value': 'Todos'}] + [{'label': str(inv), 'value': str(inv)} for inv in lista_investors], value='Todos', style={'width': '180px'}, clearable=False)
                ], style={'display': 'flex', 'alignItems': 'center'}),
                html.Div([
                    html.Label("Classe:", style={'fontWeight': 'bold', 'marginRight': '8px', 'color': PALETTE['text_body']}),
                    dcc.Dropdown(id='mesa-classe-dropdown', options=[{'label': 'Todas as Classes', 'value': 'Todas'}] + [{'label': str(c), 'value': str(c)} for c in lista_classes_disponiveis], value='Todas', style={'width': '180px'}, clearable=False)
                ], style={'display': 'flex', 'alignItems': 'center'}),
                html.Div([
                    html.Label("Pessoa:", style={'fontWeight': 'bold', 'marginRight': '8px', 'color': PALETTE['text_body']}),
                    dcc.Dropdown(id='mesa-pessoa-dropdown', options=[{'label': 'Todas', 'value': 'Todas'}, {'label': 'Física', 'value': 'PESSOA FÍSICA'}, {'label': 'Jurídica', 'value': 'PESSOA JURÍDICA'}], value='PESSOA FÍSICA', style={'width': '120px'}, clearable=False)
                ], style={'display': 'flex', 'alignItems': 'center'}),
                html.Div([
                    html.Label("Modalidade:", style={'fontWeight': 'bold', 'marginRight': '8px', 'color': PALETTE['text_body']}),
                    dcc.Dropdown(id='mesa-modalidade-dropdown', options=[{'label': 'Todas', 'value': 'Todas'}, {'label': 'Fee Based', 'value': 'Fee Based'}, {'label': 'Comission Based', 'value': 'Comission Based'}], value='Todas', style={'width': '150px'}, clearable=False)
                ], style={'display': 'flex', 'alignItems': 'center'})
            ]),
            html.Button("📊 Exportar Resumo (.xlsx)", id='btn-export-geral-mesa', n_clicks=0, style=estilo_btn_dourado)
        ]),
        html.Div(id='bloco-conteudo-mesa')
    ]),

    # CONTAINER ABA 2: ASSESSOR / CLIENTE
    html.Div(id='container-aba-assessor', style={'marginTop': '25px', 'display': 'none'}, children=[
        html.Div(style={'display': 'flex', 'justify': 'space-between', 'alignItems': 'center', 'marginBottom': '20px', 'flexWrap': 'wrap', 'gap': '12px'}, children=[
            html.Div(style={'display': 'flex', 'gap': '15px', 'flexGrow': '1', 'flexWrap': 'wrap'}, children=[
                html.Div([
                    html.Label("Assessor:", style={'fontWeight': 'bold', 'marginRight': '8px', 'color': PALETTE['text_body']}),
                    dcc.Dropdown(id='cockpit-assessor-dropdown', options=[{'label': 'Todos', 'value': 'Todos'}] + [{'label': str(ass), 'value': str(ass)} for ass in lista_assessores], value='Todos', style={'width': '140px'}, clearable=False)
                ], style={'display': 'flex', 'alignItems': 'center'}),
                html.Div([
                    html.Label("Classe:", style={'fontWeight': 'bold', 'marginRight': '8px', 'color': PALETTE['text_body']}),
                    dcc.Dropdown(id='cockpit-classe-dropdown', options=[{'label': 'Todas as Classes', 'value': 'Todas'}] + [{'label': str(c), 'value': str(c)} for c in lista_classes_disponiveis], value='Todas', style={'width': '140px'}, clearable=False)
                ], style={'display': 'flex', 'alignItems': 'center'}),
                html.Div([
                    html.Label("Conta:", style={'fontWeight': 'bold', 'marginRight': '8px', 'color': PALETTE['text_body']}),
                    dcc.Dropdown(id='cockpit-cliente-dropdown', options=[{'label': 'Todas', 'value': 'Todas'}], value='Todas', style={'width': '140px'}, clearable=True, placeholder="Buscar")
                ], style={'display': 'flex', 'alignItems': 'center'}),
                html.Div([
                    html.Label("Pessoa:", style={'fontWeight': 'bold', 'marginRight': '8px', 'color': PALETTE['text_body']}),
                    dcc.Dropdown(id='cockpit-pessoa-dropdown', options=[{'label': 'Todas', 'value': 'Todas'}, {'label': 'Física', 'value': 'PESSOA FÍSICA'}, {'label': 'Jurídica', 'value': 'PESSOA JURÍDICA'}], value='PESSOA FÍSICA', style={'width': '110px'}, clearable=False)
                ], style={'display': 'flex', 'alignItems': 'center'}),
                html.Div([
                    html.Label("Modalidade:", style={'fontWeight': 'bold', 'marginRight': '8px', 'color': PALETTE['text_body']}),
                    dcc.Dropdown(id='cockpit-modalidade-dropdown', options=[{'label': 'Todas', 'value': 'Todas'}, {'label': 'Fee Based', 'value': 'Fee Based'}, {'label': 'Comission Based', 'value': 'Comission Based'}], value='Todas', style={'width': '140px'}, clearable=False)
                ], style={'display': 'flex', 'alignItems': 'center'})
            ]),
            html.Button("📊 Exportar Assessor (.xlsx)", id='btn-export-geral-assessor-aba2', n_clicks=0, style=estilo_btn_dourado)
        ]),
        
        html.Div(id='bloco-kpis-graficos-assessor'),

        html.Div(style={'marginTop': '30px', 'backgroundColor': PALETTE['bg_card'], 'borderRadius': '8px', 'padding': '22px', 'boxShadow': '0 4px 12px rgba(0,0,0,0.03)', 'border': f"1px solid {PALETTE['hairline']}"}, children=[
            html.Div(style={'display': 'flex', 'justify': 'space-between', 'alignItems': 'center', 'marginBottom': '15px'}, children=[
                html.H3("📋 Resumo de Alocação por Classe do Assessor", style={'fontWeight':'bold', 'margin': '0', 'color': PALETTE['bg_dark'], 'fontFamily': FONT_SERIF}),
                html.Div(style={'display': 'flex', 'gap': '10px'}, children=[
                    html.Button("📥 Exportar Classes (.xlsx)", id='btn-export-classes', n_clicks=0, style=estilo_btn_dourado),
                    html.Button("🔄 Limpar Filtro", id='btn-limpar-classe', n_clicks=0, style={'backgroundColor': PALETTE['red_soft'], 'color': 'white', 'border': 'none', 'padding': '8px 15px', 'borderRadius': '6px', 'fontWeight': 'bold', 'cursor': 'pointer', 'fontSize': '12px', 'fontFamily': FONT_SANS})
                ])
            ]),
            dash_table.DataTable(
                id='tabela-classes-assessor', data=[], sort_action='native', locale_format=formato_local_br,
                columns=[
                    {"name": "Classe de Ativo", "id": "Classe"},
                    {"name": "Custódia (R$)", "id": "Custódia", "type": "numeric", "format": money_format},
                    {"name": "Aderência", "id": "Aderência", "type": "numeric", "format": pct_format},
                    {"name": "Gap Under (Entrada)", "id": "Gap Under", "type": "numeric", "format": money_format},
                    {"name": "Gap Over (Saída)", "id": "Gap Over", "type": "numeric", "format": money_format},
                    {"name": "Nº Clientes", "id": "Contas", "type": "numeric", "format": num_format}
                ],
                style_cell={'textAlign': 'left', 'padding': '12px', **estilo_tabela_comum},
                style_header={'backgroundColor': PALETTE['bg_dark'], 'color': PALETTE['gold_bright'], 'fontWeight': 'bold'},
                style_data_conditional=[{'if': {'row_index': 'odd'}, 'backgroundColor': '#FFFFFF'}]
            )
        ]),

        html.Div(style={'marginTop': '30px', 'backgroundColor': PALETTE['bg_card'], 'borderRadius': '8px', 'padding': '22px', 'boxShadow': '0 4px 12px rgba(0,0,0,0.03)', 'border': f"1px solid {PALETTE['hairline']}"}, children=[
            html.Div(style={'display': 'flex', 'justify': 'space-between', 'alignItems': 'center', 'marginBottom': '10px'}, children=[
                html.H3(id='titulo-detalhe-clientes-assessor', children="📋 Detalhe Geral de Clientes (Visão 360º)", style={'fontWeight':'bold', 'margin': '0', 'color': PALETTE['bg_dark'], 'fontFamily': FONT_SERIF}),
                html.Div(style={'display': 'flex', 'gap': '10px'}, children=[
                    html.Button("📥 Exportar Clientes (.xlsx)", id='btn-export-clientes', n_clicks=0, style=estilo_btn_dourado),
                    html.Button("❌ Fechar Raio-X", id='btn-limpar-cliente', n_clicks=0, style={'backgroundColor': PALETTE['gold_muted'], 'color': 'white', 'border': 'none', 'padding': '8px 15px', 'borderRadius': '6px', 'fontWeight': 'bold', 'cursor': 'pointer', 'fontFamily': FONT_SANS})
                ])
            ]),
            dash_table.DataTable(
                id='tabela-clientes-assessor', data=[], sort_action='native', locale_format=formato_local_br,
                columns=[
                    {"name": "Cód. Cliente", "id": "Cliente"},
                    {"name": "Idade", "id": "Idade"},
                    {"name": "Custódia Total (AUC)", "id": "Custódia", "type": "numeric", "format": money_format},
                    {"name": "Liquidez Até D+5", "id": "LIQUIDEZ_TOTAL_360", "type": "numeric", "format": money_format},
                    {"name": "Aderência", "id": "Aderência", "type": "numeric", "format": pct_format},
                    {"name": "Gap Under (Entrada)", "id": "Gap Under", "type": "numeric", "format": money_format},
                    {"name": "Gap Over (Saída)", "id": "Gap Over", "type": "numeric", "format": money_format}
                ],
                style_cell={'textAlign': 'left', 'padding': '12px', 'whiteSpace': 'normal', **estilo_tabela_comum},
                style_header={'backgroundColor': PALETTE['bg_dark'], 'color': PALETTE['gold_bright'], 'fontWeight': 'bold'},
                style_data_conditional=[{'if': {'row_index': 'odd'}, 'backgroundColor': '#FFFFFF'}]
            )
        ]),

        # RAIO-X BLINDADO DO CLIENTE
        html.Div(id='secao-ancora-raio-x', style={'marginTop': '30px', 'display': 'none'}, children=[
            html.Div(style={'backgroundColor': PALETTE['bg_card'], 'borderRadius': '8px', 'padding': '25px', 'boxShadow': '0 4px 12px rgba(0,0,0,0.04)', 'border': f"1px solid {PALETTE['hairline']}"}, children=[
                
                html.Div(style={'display': 'flex', 'alignItems': 'center', 'justify': 'center', 'gap': '10px'}, children=[
                    html.H2(id='rx-titulo-cliente', style={'margin': 0, 'color': PALETTE['bg_dark'], 'fontFamily': FONT_SERIF}),
                    dcc.Clipboard(id="clip-conta", style={'fontSize': '22px', 'color': PALETTE['gold_main'], 'cursor': 'pointer'}),
                    html.H2(id='rx-titulo-complemento', style={'margin': 0, 'color': PALETTE['bg_dark'], 'fontFamily': FONT_SERIF})
                ]),
                
                html.Div(style={'marginTop': '15px', 'textAlign': 'center'}, children=[
                    html.Span("🏷️ Tags Qualitativas de Perfil: ", style={'fontWeight': 'bold', 'color': PALETTE['text_label'], 'fontSize': '13px'}),
                    html.Div(id='rx-tags-badges-container', style={'display': 'inline-flex', 'gap': '8px', 'flexWrap': 'wrap', 'justify': 'center', 'marginTop': '6px'})
                ]),

                # LIQUIDIDADE
                html.Div(style={'display': 'grid', 'gridTemplateColumns': 'repeat(auto-fit, minmax(200px, 1fr))', 'gap': '15px', 'marginTop': '20px'}, children=[
                    html.Div(style={'backgroundColor': '#FFFFFF', 'padding': '16px 20px', 'borderRadius': '8px', 'border': f"1px solid {PALETTE['hairline']}", 'textAlign': 'center'}, children=[
                        html.Span("Saldo Conta", style={'fontSize': '11px', 'fontWeight': 'bold', 'color': PALETTE['text_label'], 'display': 'block', 'marginBottom': '6px', 'textTransform': 'uppercase'}),
                        html.Span(id='rx-saldo', style={'fontSize': '20px', 'fontWeight': 'bold', 'color': PALETTE['bg_dark'], 'fontFamily': FONT_SERIF})
                    ]),
                    html.Div(style={'backgroundColor': '#FFFFFF', 'padding': '16px 20px', 'borderRadius': '8px', 'border': f"1px solid {PALETTE['hairline']}", 'textAlign': 'center'}, children=[
                        html.Span("Fundos Até D+5", style={'fontSize': '11px', 'fontWeight': 'bold', 'color': PALETTE['text_label'], 'display': 'block', 'marginBottom': '6px', 'textTransform': 'uppercase'}),
                        html.Span(id='rx-fundos', style={'fontSize': '20px', 'fontWeight': 'bold', 'color': PALETTE['bg_dark'], 'fontFamily': FONT_SERIF})
                    ]),
                    html.Div(style={'backgroundColor': '#FFFFFF', 'padding': '16px 20px', 'borderRadius': '8px', 'border': f"1px solid {PALETTE['hairline']}", 'textAlign': 'center'}, children=[
                        html.Span("Renda Fixa Liquidez", style={'fontSize': '11px', 'fontWeight': 'bold', 'color': PALETTE['text_label'], 'display': 'block', 'marginBottom': '6px', 'textTransform': 'uppercase'}),
                        html.Span(id='rx-rf', style={'fontSize': '20px', 'fontWeight': 'bold', 'color': PALETTE['bg_dark'], 'fontFamily': FONT_SERIF})
                    ]),
                    html.Div(style={'backgroundColor': '#FFFFFF', 'padding': '16px 20px', 'borderRadius': '8px', 'border': f"1px solid {PALETTE['hairline']}", 'textAlign': 'center'}, children=[
                        html.Span("Liquidez Total D+5", style={'fontSize': '11px', 'fontWeight': 'bold', 'color': PALETTE['text_label'], 'display': 'block', 'marginBottom': '6px', 'textTransform': 'uppercase'}),
                        html.Span(id='rx-total', style={'fontSize': '22px', 'fontWeight': 'bold', 'color': PALETTE['gold_main'], 'fontFamily': FONT_SERIF})
                    ])
                ]),

                # FORMULÁRIO DE REGISTRO DO CRM
                html.Div(style={'backgroundColor': '#FFFFFF', 'borderLeft': f"4px solid {PALETTE['gold_main']}", 'padding': '22px', 'margin': '25px 0', 'borderRadius': '0 8px 8px 0', 'border': f"1px solid {PALETTE['hairline']}"}, children=[
                    html.H3("🚀 Cadastrar Oferta Comercial (Pipeline CRM)", style={'margin': '0 0 18px 0', 'color': PALETTE['bg_dark'], 'fontFamily': FONT_SERIF}),
                    
                    html.Div(style={'display': 'flex', 'gap': '15px', 'marginBottom': '15px', 'flexWrap': 'wrap'}, children=[
                        html.Div([
                            html.Label("Classe / Gap:", style={'fontWeight': 'bold', 'fontSize': '12px', 'display': 'block', 'marginBottom': '5px', 'color': PALETTE['text_body']}),
                            dcc.Dropdown(id='crm-fator-dropdown', options=[{'label': c, 'value': c} for c in lista_classes_disponiveis], placeholder="Selecione...", style={'width': '180px'})
                        ]),
                        html.Div([
                            html.Label("Produto Sugerido:", style={'fontWeight': 'bold', 'fontSize': '12px', 'display': 'block', 'marginBottom': '5px', 'color': PALETTE['text_body']}),
                            dcc.Input(id='crm-produto-input', type='text', placeholder="Ex: CYHF11 / CDB 110%", style={'padding': '8px', 'borderRadius': '6px', 'border': f"1px solid {PALETTE['hairline']}", 'width': '200px'})
                        ]),
                        html.Div([
                            html.Label("Volume Alvo (R$):", style={'fontWeight': 'bold', 'fontSize': '12px', 'display': 'block', 'marginBottom': '5px', 'color': PALETTE['text_body']}),
                            dcc.Input(id='crm-valor-input', type='text', placeholder="Ex: 50000", style={'padding': '8px', 'borderRadius': '6px', 'border': f"1px solid {PALETTE['hairline']}", 'width': '140px'})
                        ]),
                        html.Div([
                            html.Label("ROA Est. (% a.a.):", style={'fontWeight': 'bold', 'fontSize': '12px', 'display': 'block', 'marginBottom': '5px', 'color': PALETTE['gold_main']}),
                            dcc.Input(id='crm-roa-input', type='text', placeholder="Ex: 0,80 ou 1,20", style={'padding': '8px', 'borderRadius': '6px', 'border': f"1px solid {PALETTE['gold_main']}", 'width': '120px'})
                        ]),
                        html.Div([
                            html.Label("Prazo Follow-up:", style={'fontWeight': 'bold', 'fontSize': '12px', 'display': 'block', 'marginBottom': '5px', 'color': PALETTE['text_body']}),
                            dcc.Dropdown(id='crm-dias-dropdown', options=[{'label': '1 dia útil', 'value': 1}, {'label': '2 dias úteis', 'value': 2}, {'label': '3 dias úteis', 'value': 3}, {'label': '5 dias úteis', 'value': 5}], value=3, style={'width': '130px'}, clearable=False)
                        ])
                    ]),

                    html.Div(style={'display': 'flex', 'gap': '20px', 'alignItems': 'flex-end', 'flexWrap': 'wrap'}, children=[
                        html.Div(style={'flexGrow': '1', 'minWidth': '320px'}, children=[
                            html.Label("Observação Comercial / Objeções:", style={'fontWeight': 'bold', 'fontSize': '12px', 'display': 'block', 'marginBottom': '5px', 'color': PALETTE['text_body']}),
                            dcc.Input(id='crm-obs-input', type='text', placeholder="Digite os detalhes da tese comercial...", style={'padding': '10px', 'borderRadius': '6px', 'border': f"1px solid {PALETTE['hairline']}", 'width': '100%'})
                        ]),
                        html.Button("💾 Gravar Tese no Pipeline", id='btn-salvar-crm', n_clicks=0, style=estilo_btn_dourado)
                    ]),
                    html.Div(id='crm-feedback-registro', style={'marginTop': '12px', 'fontWeight': 'bold'})
                ]),

                html.Hr(style={'borderColor': PALETTE['hairline']}),
                html.Div(style={'backgroundColor': '#FFFFFF', 'borderLeft': f"4px solid {PALETTE['red_soft']}", 'padding': '22px', 'marginBottom': '20px', 'borderRadius': '0 8px 8px 0', 'border': f"1px solid {PALETTE['hairline']}"}, children=[
                    html.H3("🛡️ Travas de Compliance Institucional XP", style={'margin': '0 0 15px 0', 'color': PALETTE['red_soft'], 'fontFamily': FONT_SERIF}),
                    html.Div(id='rx-compliance-regras', style={'display': 'grid', 'gridTemplateColumns': 'repeat(auto-fit, minmax(280px, 1fr))', 'gap': '15px'})
                ]),

                dcc.Graph(id='grafico-cockpit-cliente-clicado')
            ])
        ])
    ]),

    # CONTAINER ABA 3: CRM TÁTICO & PIPELINE DE RECEITA
    html.Div(id='container-aba-crm', style={'marginTop': '25px', 'display': 'none'}, children=[
        html.Div(id='bloco-kpis-crm-pipeline', style={'marginBottom': '20px'}),

        html.Div(style={'backgroundColor': PALETTE['bg_card'], 'padding': '18px 22px', 'borderRadius': '8px', 'marginBottom': '20px', 'display': 'flex', 'gap': '20px', 'alignItems': 'center', 'flexWrap': 'wrap', 'border': f"1px solid {PALETTE['hairline']}"}, children=[
            html.Div([
                html.Label("Período / Safra:", style={'fontWeight': 'bold', 'marginRight': '8px', 'color': PALETTE['text_body']}),
                dcc.Dropdown(id='crm-filtro-periodo', options=[
                    {'label': NOME_MES_ATUAL, 'value': 'mes_atual'},
                    {'label': 'Histórico Completo (Todas as Safras)', 'value': 'todos'}
                ], value='mes_atual', style={'width': '250px'}, clearable=False)
            ]),
            html.Div([
                html.Label("Status na Tabela:", style={'fontWeight': 'bold', 'marginRight': '8px', 'color': PALETTE['text_body']}),
                dcc.Dropdown(id='crm-filtro-status', options=[
                    {'label': 'Pendentes de Retorno', 'value': 'Pendente'},
                    {'label': 'Convertidos (Sucesso)', 'value': 'Convertido'},
                    {'label': 'Recusados', 'value': 'Recusado'},
                    {'label': 'Todos os Registros', 'value': 'Todos'}
                ], value='Pendente', style={'width': '220px'}, clearable=False)
            ]),
            html.Div([
                html.Label("Investor:", style={'fontWeight': 'bold', 'marginRight': '8px', 'color': PALETTE['text_body']}),
                dcc.Dropdown(id='crm-filtro-investor', options=[{'label': 'Todos', 'value': 'Todos'}] + [{'label': inv, 'value': inv} for inv in lista_investors], value='Todos', style={'width': '160px'}, clearable=False)
            ]),
            html.Div([
                html.Label("Assessor:", style={'fontWeight': 'bold', 'marginRight': '8px', 'color': PALETTE['text_body']}),
                dcc.Dropdown(id='crm-filtro-assessor', options=[{'label': 'Todos', 'value': 'Todos'}] + [{'label': ass, 'value': ass} for ass in lista_assessores], value='Todos', style={'width': '150px'}, clearable=False)
            ]),
            html.Button("🔄 Atualizar Fila / KPIs", id='btn-atualizar-crm', n_clicks=0, style={'backgroundColor': PALETTE['bg_dark'], 'color': PALETTE['gold_bright'], 'border': 'none', 'padding': '9px 16px', 'borderRadius': '6px', 'fontWeight': 'bold', 'cursor': 'pointer', 'fontSize': '12px', 'fontFamily': FONT_SANS})
        ]),

        html.Div(style={'backgroundColor': PALETTE['bg_card'], 'borderRadius': '8px', 'padding': '22px', 'marginBottom': '25px', 'border': f"1px solid {PALETTE['hairline']}"}, children=[
            html.H3("📋 Fila de Ofertas Comerciais & Projeção de Receita", style={'marginTop': 0, 'color': PALETTE['bg_dark'], 'fontFamily': FONT_SERIF}),
            dash_table.DataTable(
                id='tabela-pipeline-crm',
                columns=[
                    {"name": "ID", "id": "id"},
                    {"name": "Data Oferta", "id": "data_criacao"},
                    {"name": "Assessor", "id": "assessor"},
                    {"name": "Conta Cliente", "id": "conta"},
                    {"name": "Classe / Gap", "id": "fator_risco"},
                    {"name": "Produto Sugerido", "id": "produto_sugerido"},
                    {"name": "Vol. Proposto (R$)", "id": "valor_alvo", "type": "numeric", "format": money_format},
                    {"name": "Vol. Convertido (R$)", "id": "valor_convertido", "type": "numeric", "format": money_format},
                    {"name": "ROA (%)", "id": "roa", "type": "numeric", "format": pct_format},
                    {"name": "Receita Est. (R$)", "id": "receita_estimada", "type": "numeric", "format": money_format},
                    {"name": "Data Follow-up", "id": "data_followup"},
                    {"name": "Status Atual", "id": "status"}
                ],
                data=[], sort_action='native',
                style_cell={'textAlign': 'left', 'padding': '12px', **estilo_tabela_comum},
                style_header={'backgroundColor': PALETTE['bg_dark'], 'color': PALETTE['gold_bright'], 'fontWeight': 'bold'},
                style_data_conditional=[
                    {'if': {'filter_query': '{status} = "Pendente"'}, 'backgroundColor': '#FFFFFF', 'color': PALETTE['red_soft']},
                    {'if': {'filter_query': '{status} = "Convertido"'}, 'backgroundColor': '#FFFFFF', 'color': PALETTE['green_soft']},
                    {'if': {'filter_query': '{status} = "Recusado"'}, 'backgroundColor': '#FFFFFF', 'color': PALETTE['gold_muted']}
                ]
            )
        ]),

        html.Div(id='card-baixa-crm-container', style={'backgroundColor': PALETTE['bg_card'], 'borderRadius': '8px', 'padding': '22px', 'border': f"1px solid {PALETTE['hairline']}"}, children=[
            html.H3("✏️ Dar Baixa / Registrar Devolutiva da Oferta Selecionada", style={'marginTop': 0, 'color': PALETTE['bg_dark'], 'fontFamily': FONT_SERIF}),
            html.Div(id='crm-detalhe-tarefa-selecionada', children="💡 Clique em uma linha da tabela acima para selecionar a oferta.", style={'marginBottom': '15px', 'fontWeight': 'bold', 'color': PALETTE['text_body']}),
            
            html.Div(style={'display': 'flex', 'gap': '15px', 'flexWrap': 'wrap', 'alignItems': 'flex-end'}, children=[
                html.Div([
                    html.Label("Novo Status:", style={'fontWeight': 'bold', 'fontSize': '12px', 'color': PALETTE['text_body']}),
                    dcc.Dropdown(id='crm-baixa-status-dropdown', options=[
                        {'label': '✅ Convertido', 'value': 'Convertido'},
                        {'label': '❌ Recusado', 'value': 'Recusado'},
                        {'label': '⏳ Manter Pendente', 'value': 'Pendente'}
                    ], value='Convertido', style={'width': '200px'}, clearable=False)
                ]),
                html.Div([
                    html.Label("Vol. Convertido (R$):", style={'fontWeight': 'bold', 'fontSize': '12px', 'color': PALETTE['text_body']}),
                    dcc.Input(id='crm-baixa-valor-input', type='text', placeholder="R$ real fechado...", style={'padding': '8px', 'borderRadius': '6px', 'border': f"1px solid {PALETTE['hairline']}", 'width': '160px'})
                ]),
                html.Div([
                    html.Label("ROA Final (%):", style={'fontWeight': 'bold', 'fontSize': '12px', 'color': PALETTE['gold_main']}),
                    dcc.Input(id='crm-baixa-roa-input', type='text', placeholder="ROA %...", style={'padding': '8px', 'borderRadius': '6px', 'border': f"1px solid {PALETTE['gold_main']}", 'width': '100px'})
                ]),
                html.Div([
                    html.Label("Motivo da Recusa:", style={'fontWeight': 'bold', 'fontSize': '12px', 'color': PALETTE['text_body']}),
                    dcc.Dropdown(id='crm-baixa-motivo-dropdown', options=[
                        {'label': 'Nenhum / Venda Convertida', 'value': ''},
                        {'label': 'Prefere Curto Prazo / Liquidez', 'value': 'Cliente prefere curto prazo e resgate rápido'},
                        {'label': 'Averso a Crédito Privado', 'value': 'Não gosta de crédito privado ou risco empresa'},
                        {'label': 'Não quer Volatilidade / MTM', 'value': 'Cliente não quer oscilação de marcação a mercado'},
                        {'label': 'Acha a Taxa / Fee Alta', 'value': 'Acha taxa de administração ou fee salgado'}
                    ], value='', style={'width': '220px'})
                ]),
                html.Div(style={'flexGrow': '1', 'minWidth': '240px'}, children=[
                    html.Label("Devolutiva / Comentário:", style={'fontWeight': 'bold', 'fontSize': '12px', 'color': PALETTE['text_body']}),
                    dcc.Input(id='crm-baixa-obs-input', type='text', placeholder="Ex: Cliente fechou R$ 30k hoje...", style={'padding': '8px', 'borderRadius': '6px', 'border': f"1px solid {PALETTE['hairline']}", 'width': '100%'})
                ]),
                html.Button("✅ Confirmar Baixa", id='btn-confirmar-baixa-crm', n_clicks=0, style=estilo_btn_dourado)
            ]),
            html.Div(id='crm-feedback-baixa', style={'marginTop': '12px', 'fontWeight': 'bold'})
        ])
    ]),

    # CONTAINER ABA 4: ACELERADOR COMERCIAL & CAMPANHAS (SPLIT LAYOUT LATERAL)
    html.Div(id='container-aba-campanhas', style={'marginTop': '25px', 'display': 'none'}, children=[
        html.Div(style={'backgroundColor': PALETTE['bg_card'], 'borderRadius': '8px', 'padding': '22px', 'marginBottom': '20px', 'border': f"1px solid {PALETTE['hairline']}"}, children=[
            html.H3("🚀 Match Automático de Produtos Foco & Disparo Comercial", style={'marginTop': 0, 'color': PALETTE['bg_dark'], 'fontFamily': FONT_SERIF}),
            html.P("💡 Este motor cruza automaticamente o Gap da Classe, a Liquidez D+5 e o Limite de Compliance do cliente para sugerir o volume exato por Produto Foco.", style={'fontSize': '12px', 'color': PALETTE['text_label']}),
            
            html.Div(style={'display': 'flex', 'gap': '20px', 'alignItems': 'center', 'flexWrap': 'wrap', 'marginTop': '15px'}, children=[
                html.Div([
                    html.Label("Selecione o Produto Foco / Campanha:", style={'fontWeight': 'bold', 'marginRight': '8px', 'color': PALETTE['text_body']}),
                    dcc.Dropdown(id='campanha-produto-dropdown', options=[], placeholder="Todos os Produtos Foco", style={'width': '260px'}, clearable=True)
                ]),
                html.Div([
                    html.Label("Filtrar Investor:", style={'fontWeight': 'bold', 'marginRight': '8px', 'color': PALETTE['text_body']}),
                    dcc.Dropdown(id='campanha-investor-dropdown', options=[{'label': 'Todos', 'value': 'Todos'}] + [{'label': inv, 'value': inv} for inv in lista_investors], value='Todos', style={'width': '180px'}, clearable=False)
                ]),
                html.Div([
                    html.Label("Filtrar Assessor:", style={'fontWeight': 'bold', 'marginRight': '8px', 'color': PALETTE['text_body']}),
                    dcc.Dropdown(id='campanha-assessor-dropdown', options=[{'label': 'Todos', 'value': 'Todos'}] + [{'label': a, 'value': a} for a in lista_assessores], value='Todos', style={'width': '180px'}, clearable=False)
                ]),
                html.Button("⚡ Executar Match de Oportunidades", id='btn-rodar-campanha', n_clicks=0, style=estilo_btn_dourado)
            ])
        ]),

        html.Div(id='bloco-kpis-campanha', style={'marginBottom': '20px'}),

        # GRID SPLIT: 68% TABELA DE OPORTUNIDADES | 32% CENTRAL FIXA LATERAL
        html.Div(style={'display': 'flex', 'gap': '20px', 'alignItems': 'flex-start', 'flexWrap': 'wrap'}, children=[
            
            # COLUNA ESQUERDA: TABELA DE LEADS DE CAMPANHA
            html.Div(style={'flex': '1 1 65%', 'minWidth': '600px', 'backgroundColor': PALETTE['bg_card'], 'borderRadius': '8px', 'padding': '22px', 'border': f"1px solid {PALETTE['hairline']}"}, children=[
                html.H3("🎯 Oportunidades Mapeadas na Campanha", style={'marginTop': 0, 'color': PALETTE['bg_dark'], 'fontFamily': FONT_SERIF}),
                html.P("💡 Clique em qualquer linha para carregar a mensagem formatada no painel lateral.", style={'fontSize': '12px', 'color': PALETTE['text_label']}),
                dash_table.DataTable(
                    id='tabela-resultado-campanha',
                    columns=[
                        {"name": "Conta", "id": "Cód. Conta"},
                        {"name": "Assessor", "id": "Assessor"},
                        {"name": "Produto Foco", "id": "Produto Foco"},
                        {"name": "Gap Classe", "id": "Gap Classe", "type": "numeric", "format": money_format},
                        {"name": "Liquidez D+5", "id": "Liquidez D+5", "type": "numeric", "format": money_format},
                        {"name": "Vol. Sugerido", "id": "Volume Sugerido", "type": "numeric", "format": money_format},
                        {"name": "ROA (%)", "id": "ROA (%)", "type": "numeric", "format": pct_format},
                        {"name": "Receita Est. (R$)", "id": "Receita Estimada", "type": "numeric", "format": money_format}
                    ],
                    data=[], sort_action='native', page_size=10,
                    style_cell={'textAlign': 'left', 'padding': '12px', **estilo_tabela_comum},
                    style_header={'backgroundColor': PALETTE['bg_dark'], 'color': PALETTE['gold_bright'], 'fontWeight': 'bold'}
                )
            ]),

            # COLUNA DIREITA: PAINEL LATERAL DE DISPARO FIXO (STICKY)
            html.Div(style={'flex': '0 0 32%', 'minWidth': '320px', 'position': 'sticky', 'top': '20px', 'backgroundColor': '#FFFFFF', 'borderRadius': '8px', 'padding': '20px', 'border': f"1px solid {PALETTE['hairline']}", 'boxShadow': '0 4px 16px rgba(0,0,0,0.06)'}, children=[
                html.Div(style={'display': 'flex', 'justify': 'space-between', 'alignItems': 'center', 'marginBottom': '12px'}, children=[
                    html.H4("💬 Central de Disparo WhatsApp", style={'margin': 0, 'color': PALETTE['bg_dark'], 'fontFamily': FONT_SERIF}),
                    dcc.Clipboard(
                        id='btn-copiar-wa',
                        target_id='campanha-preview-texto',
                        title="Copiar Mensagem",
                        style={'backgroundColor': PALETTE['gold_main'], 'color': 'white', 'border': 'none', 'padding': '6px 12px', 'borderRadius': '6px', 'fontWeight': 'bold', 'cursor': 'pointer', 'fontSize': '12px', 'fontFamily': FONT_SANS}
                    )
                ]),
                
                html.Div(id='campanha-preview-texto', children="Clique em uma linha da tabela ao lado para carregar o texto formatado para o assessor.", style={'backgroundColor': PALETTE['bg_card'], 'padding': '14px', 'borderRadius': '6px', 'border': f"1px solid {PALETTE['hairline']}", 'whiteSpace': 'pre-wrap', 'fontFamily': FONT_SANS, 'fontSize': '13px', 'lineHeight': '1.5', 'color': PALETTE['text_dark'], 'minHeight': '180px', 'maxHeight': '320px', 'overflowY': 'auto'}),
                
                html.Div(style={'display': 'flex', 'flexDirection': 'column', 'gap': '10px', 'marginTop': '15px'}, children=[
                    html.A("📱 Abrir no WhatsApp Web", id='link-abrir-wa', href='#', target='_blank', style={'backgroundColor': '#25D366', 'color': 'white', 'padding': '10px', 'borderRadius': '6px', 'fontWeight': 'bold', 'textDecoration': 'none', 'fontSize': '13px', 'textAlign': 'center'}),
                    html.Button("💾 Registrar Oportunidade no CRM", id='btn-push-crm-campanha', n_clicks=0, style={'backgroundColor': PALETTE['bg_dark'], 'color': PALETTE['gold_bright'], 'border': 'none', 'padding': '10px', 'borderRadius': '6px', 'fontWeight': 'bold', 'cursor': 'pointer', 'fontSize': '13px'})
                ]),
                html.Div(id='feedback-push-crm', style={'marginTop': '12px', 'fontWeight': 'bold', 'fontSize': '12px'})
            ])
        ])
    ]),
    
    html.A("▲ Voltar ao Topo", href="#topo-marcador-ancora", style={'position': 'fixed', 'bottom': '25px', 'right': '25px', 'backgroundColor': PALETTE['bg_dark'], 'color': PALETTE['gold_bright'], 'padding': '12px 18px', 'borderRadius': '30px', 'textDecoration': 'none', 'fontWeight': 'bold', 'fontSize': '12px', 'zIndex': '9999'})
])

# =============================================================================
# CALLBACKS DO DASHBOARD
# =============================================================================
@app.callback(
    [Output('container-aba-mesa', 'style'), Output('container-aba-assessor', 'style'), 
     Output('container-aba-crm', 'style'), Output('container-aba-campanhas', 'style')],
    Input('abas-navegacao', 'value')
)
def alternar_visibilidade_abas(aba_selecionada):
    styles = [{'marginTop': '25px', 'display': 'none'}] * 4
    if aba_selecionada == 'aba-mesa': styles[0] = {'marginTop': '25px', 'display': 'block'}
    elif aba_selecionada == 'aba-assessor': styles[1] = {'marginTop': '25px', 'display': 'block'}
    elif aba_selecionada == 'aba-crm': styles[2] = {'marginTop': '25px', 'display': 'block'}
    elif aba_selecionada == 'aba-campanhas': styles[3] = {'marginTop': '25px', 'display': 'block'}
    return tuple(styles)

@app.callback(
    Output('bloco-conteudo-mesa', 'children'),
    [Input('mesa-investor-dropdown', 'value'), Input('mesa-classe-dropdown', 'value'),
     Input('mesa-pessoa-dropdown', 'value'), Input('mesa-modalidade-dropdown', 'value')]
)
def carregar_dados_mesa_isolados(investor, classe, pessoa, modalidade):
    df_filtrado = df_clean.copy()
    if investor != 'Todos': df_filtrado = df_filtrado[df_filtrado['Investor'] == investor]
    if classe != 'Todas': df_filtrado = df_filtrado[df_filtrado['Fator de Risco'] == classe]
    if pessoa != 'Todas': df_filtrado = df_filtrado[df_filtrado['Tipo_Pessoa'] == pessoa]
    if modalidade != 'Todas': df_filtrado = df_filtrado[df_filtrado['Modalidade'] == modalidade]

    df_liq_filtrado = df_liquidez_contas.copy()
    if investor != 'Todos': df_liq_filtrado = df_liq_filtrado[df_liq_filtrado['Investor'] == investor]
    if pessoa != 'Todas': df_liq_filtrado = df_liq_filtrado[df_liq_filtrado['Tipo_Pessoa'] == pessoa]
    if modalidade != 'Todas': df_liq_filtrado = df_liq_filtrado[df_liq_filtrado['Modalidade'] == modalidade]

    if df_filtrado.empty: return html.Div("Sem dados para os filtros selecionados.")

    kpi_under = formatar_moeda_br(df_filtrado['Gap Under'].sum())
    kpi_over = formatar_moeda_br(df_filtrado['Gap Over'].sum())
    kpi_liquidez = formatar_moeda_br(df_liq_filtrado['LIQUIDEZ_TOTAL_360'].sum())
    
    val_aderencia = calcular_aderencia_ponderada_df(df_filtrado)
    cor_aderencia = PALETTE['green_soft'] if val_aderencia >= 0.70 else PALETTE['red_soft']
    kpi_aderencia = formatar_percentual_br(val_aderencia)
    
    kpi_contas = formatar_numero_br(df_filtrado['Cód. Conta'].nunique())

    custodia_por_classe = df_filtrado.groupby('Fator de Risco', as_index=False).agg({'Custódia': 'sum', 'Gap Under': 'sum', 'Gap Over': 'sum'})
    custodia_por_classe['Alvo Teórico'] = custodia_por_classe['Custódia'] + custodia_por_classe['Gap Under'] - custodia_por_classe['Gap Over']
    
    fig_comparativo = go.Figure()
    fig_comparativo.add_trace(go.Bar(
        x=custodia_por_classe['Fator de Risco'], y=custodia_por_classe['Custódia'],
        name='Custódia Atual', marker_color=PALETTE['bg_dark'],
        customdata=[formatar_moeda_br(v) for v in custodia_por_classe['Custódia']], hovertemplate="%{customdata}<extra></extra>"
    ))
    fig_comparativo.add_trace(go.Bar(
        x=custodia_por_classe['Fator de Risco'], y=custodia_por_classe['Alvo Teórico'],
        name='Alocação Alvo', marker_color=PALETTE['gold_main'],
        customdata=[formatar_moeda_br(v) for v in custodia_por_classe['Alvo Teórico']], hovertemplate="%{customdata}<extra></extra>"
    ))
    fig_comparativo.update_layout(
        title=dict(text='<b>Comparativo: Alocação Atual vs. Alvo por Classe</b>', font=dict(family=FONT_SERIF, size=16, color=PALETTE['bg_dark'])),
        barmode='group', xaxis_tickangle=-45, margin=dict(l=20, r=20, t=40, b=100),
        paper_bgcolor='#FFFFFF', plot_bgcolor='#FFFFFF', font=dict(family=FONT_SANS, color=PALETTE['text_body'])
    )

    tabela_investors = df_filtrado.groupby('Investor').agg({'Custódia': 'sum', 'Gap Under': 'sum', 'Gap Over': 'sum', 'Cód. Conta': 'nunique'}).reset_index().rename(columns={'Cód. Conta': 'Contas'})
    tabela_investors['Aderência'] = [calcular_aderencia_ponderada_df(df_filtrado[df_filtrado['Investor'] == inv]) for inv in tabela_investors['Investor']]
    tabela_investors = tabela_investors.merge(df_liq_filtrado.groupby('Investor')['LIQUIDEZ_TOTAL_360'].sum().reset_index(), on='Investor', how='left').fillna({'LIQUIDEZ_TOTAL_360': 0})
    colunas_mesa_investors = [{"name": "Investor", "id": "Investor"}, {"name": "Custódia", "id": "Custódia", "type": "numeric", "format": money_format}, {"name": "Liquidez Até D+5", "id": "LIQUIDEZ_TOTAL_360", "type": "numeric", "format": money_format}, {"name": "Gap Under", "id": "Gap Under", "type": "numeric", "format": money_format}, {"name": "Gap Over", "id": "Gap Over", "type": "numeric", "format": money_format}, {"name": "Contas", "id": "Contas", "type": "numeric", "format": num_format}, {"name": "Aderência", "id": "Aderência", "type": "numeric", "format": pct_format}]

    return html.Div([
        html.Div(style={'display': 'flex', 'gap': '15px', 'marginBottom': '20px', 'flexWrap': 'wrap'}, children=[
            criar_card_kpi("Gap Under Carteira", kpi_under, PALETTE['green_soft']),
            criar_card_kpi("Liquidez Até D+5", kpi_liquidez, PALETTE['bg_dark']),
            criar_card_kpi("Gap Over (Saída)", kpi_over, PALETTE['red_soft']),
            criar_card_kpi("Aderência Média", kpi_aderencia, cor_aderencia),
            criar_card_kpi("Nº de Contas", kpi_contas, PALETTE['bg_dark'])
        ]),
        html.Div(style={'width': '100%', 'backgroundColor': '#FFFFFF', 'borderRadius': '8px', 'padding': '18px', 'boxShadow': '0 4px 12px rgba(0,0,0,0.03)', 'border': f"1px solid {PALETTE['hairline']}"}, children=[dcc.Graph(figure=fig_comparativo)]),
        html.Div(style={'marginTop': '30px', 'backgroundColor': PALETTE['bg_card'], 'borderRadius': '8px', 'padding': '22px', 'border': f"1px solid {PALETTE['hairline']}"}, children=[
            html.H3("📋 Resumo por Investor", style={'fontWeight':'bold', 'margin': '0 0 15px 0', 'color': PALETTE['bg_dark'], 'fontFamily': FONT_SERIF}),
            dash_table.DataTable(columns=colunas_mesa_investors, data=tabela_investors.to_dict('records'), sort_action='native', locale_format=formato_local_br, style_cell={'textAlign': 'left', 'padding': '12px', **estilo_tabela_comum}, style_header={'backgroundColor': PALETTE['bg_dark'], 'color': PALETTE['gold_bright'], 'fontWeight': 'bold'})
        ])
    ])

# =============================================================================
# COCKPIT DO ASSESSOR (TÁTICA) & RAIO-X DO CLIENTE
# =============================================================================
@app.callback(
    [Output('cockpit-cliente-dropdown', 'options'), Output('cockpit-cliente-dropdown', 'value')],
    [Input('cockpit-assessor-dropdown', 'value'), Input('cockpit-pessoa-dropdown', 'value'), Input('cockpit-modalidade-dropdown', 'value')]
)
def atualizar_lista_clientes_dropdown(assessor, pessoa, modalidade):
    df_temp = df_clean.copy()
    if assessor and assessor != 'Todos': df_temp = df_temp[df_temp['Assessor'] == str(assessor)]
    if pessoa and pessoa != 'Todas': df_temp = df_temp[df_temp['Tipo_Pessoa'] == pessoa]
    if modalidade and modalidade != 'Todas': df_temp = df_temp[df_temp['Modalidade'] == modalidade]

    clientes = sorted(df_temp['Cód. Conta'].unique())
    return [{'label': 'Todas', 'value': 'Todas'}] + [{'label': str(c), 'value': str(c)} for c in clientes], 'Todas'

@app.callback(
    [Output('bloco-kpis-graficos-assessor', 'children'), Output('tabela-classes-assessor', 'data'),
     Output('tabela-clientes-assessor', 'data'), Output('titulo-detalhe-clientes-assessor', 'children')],
    [Input('cockpit-assessor-dropdown', 'value'), Input('cockpit-classe-dropdown', 'value'), 
     Input('cockpit-cliente-dropdown', 'value'), Input('cockpit-pessoa-dropdown', 'value'), Input('cockpit-modalidade-dropdown', 'value')]
)
def atualizar_cockpit_assessor_geral(assessor, classe, cliente, pessoa, modalidade):
    df_f = df_clean.copy()
    df_l = df_liquidez_contas.copy()

    if assessor and assessor != 'Todos':
        df_f = df_f[df_f['Assessor'] == str(assessor)]
        df_l = df_l[df_l['Assessor'] == str(assessor)]
    if pessoa and pessoa != 'Todas':
        df_f = df_f[df_f['Tipo_Pessoa'] == pessoa]
        df_l = df_l[df_l['Tipo_Pessoa'] == pessoa]
    if modalidade and modalidade != 'Todas':
        df_f = df_f[df_f['Modalidade'] == modalidade]
        df_l = df_l[df_l['Modalidade'] == modalidade]
    
    if cliente and str(cliente).strip() not in ['Todas', 'None', '', 'null']:
        df_f = df_f[df_f['Cód. Conta'] == str(cliente)]
        df_l = df_l[df_l['Cód. Conta'] == str(cliente)]

    if df_f.empty: return html.Div("Sem dados."), [], [], "📋 Detalhe Geral de Clientes"

    kpi_under = formatar_moeda_br(df_f['Gap Under'].sum())
    kpi_liquidez = formatar_moeda_br(df_l['LIQUIDEZ_TOTAL_360'].sum())
    val_aderencia = calcular_aderencia_ponderada_df(df_f)
    cor_aderencia = PALETTE['green_soft'] if val_aderencia >= 0.70 else PALETTE['red_soft']
    kpi_aderencia = formatar_percentual_br(val_aderencia)

    df_contas_base = df_f.groupby(['Cód. Conta', 'Idade']).agg({'Custódia': 'sum', 'Gap Under': 'sum', 'Gap Over': 'sum'}).reset_index()
    df_contas_base = df_contas_base.merge(df_l[['Cód. Conta', 'LIQUIDEZ_TOTAL_360']], on='Cód. Conta', how='left').fillna({'LIQUIDEZ_TOTAL_360': 0})
    df_contas_base['Aderência'] = [calcular_aderencia_ponderada_df(df_f[df_f['Cód. Conta'] == c]) for c in df_contas_base['Cód. Conta']]

    fig_ranking = px.bar(
        df_contas_base.sort_values('Aderência').head(15), y='Cód. Conta', x='Aderência', orientation='h',
        color='Aderência', color_continuous_scale=PALETTE['gold_scale']
    )
    fig_ranking.update_layout(
        title=dict(text='<b>Top 15 Contas com Menor Aderência</b>', font=dict(family=FONT_SERIF, size=16, color=PALETTE['bg_dark'])),
        margin=dict(l=20, r=20, t=40, b=20), paper_bgcolor='#FFFFFF', plot_bgcolor='#FFFFFF', font=dict(family=FONT_SANS)
    )

    bloco_kpis = html.Div([
        html.Div(style={'display': 'flex', 'gap': '15px', 'marginBottom': '20px', 'flexWrap': 'wrap'}, children=[
            criar_card_kpi("Gap Under Carteira", kpi_under, PALETTE['green_soft']),
            criar_card_kpi("Liquidez Até D+5", kpi_liquidez, PALETTE['bg_dark']),
            criar_card_kpi("Aderência Ponderada", kpi_aderencia, cor_aderencia)
        ]),
        html.Div(style={'width': '100%', 'backgroundColor': '#FFFFFF', 'borderRadius': '8px', 'padding': '18px', 'boxShadow': '0 4px 12px rgba(0,0,0,0.03)', 'border': f"1px solid {PALETTE['hairline']}"}, children=[dcc.Graph(figure=fig_ranking)])
    ])

    tabela_classes = df_f.groupby('Fator de Risco').agg({'Custódia': 'sum', 'Gap Under': 'sum', 'Gap Over': 'sum', 'Cód. Conta': 'nunique'}).reset_index().rename(columns={'Fator de Risco': 'Classe', 'Cód. Conta': 'Contas'})
    tabela_classes['Aderência'] = [calcular_aderencia_ponderada_df(df_f[df_f['Fator de Risco'] == c]) for c in tabela_classes['Classe']]

    tabela_clientes = df_contas_base.rename(columns={'Cód. Conta': 'Cliente'}).sort_values('Gap Under', ascending=False)
    return bloco_kpis, tabela_classes.to_dict('records'), tabela_clientes.to_dict('records'), "📋 Detalhe Geral de Clientes (Visão 360º)"

@app.callback(
    [Output('secao-ancora-raio-x', 'style'), Output('rx-titulo-cliente', 'children'), Output('rx-titulo-complemento', 'children'), Output('clip-conta', 'content'),
     Output('rx-tags-badges-container', 'children'), Output('rx-saldo', 'children'), Output('rx-fundos', 'children'), 
     Output('rx-rf', 'children'), Output('rx-total', 'children'), Output('rx-compliance-regras', 'children'), 
     Output('grafico-cockpit-cliente-clicado', 'figure'), Output('tabela-clientes-assessor', 'active_cell')],
    [Input('tabela-clientes-assessor', 'active_cell'), Input('tabela-clientes-assessor', 'derived_virtual_data'), Input('btn-limpar-cliente', 'n_clicks')],
    prevent_initial_call=True
)
def renderizar_raio_x_cliente(active_cell, table_data, n_clicks_limpar):
    if ctx.triggered_id == 'btn-limpar-cliente' or not active_cell or not table_data:
        return {'display': 'none'}, "", "", "", [], "", "", "", "", "", go.Figure(), None
        
    linha_index = active_cell['row']
    if linha_index >= len(table_data): return {'display': 'none'}, "", "", "", [], "", "", "", "", "", go.Figure(), None

    id_cliente = str(table_data[linha_index]['Cliente'])
    df_cli = df_clean[df_clean['Cód. Conta'] == id_cliente].copy()
    if df_cli.empty: return {'display': 'none'}, "", "", "", [], "", "", "", "", "", go.Figure(), None

    custodia_total = df_cli['Custódia'].sum()
    idade_raw = df_cli['Idade'].max()
    idade = int(idade_raw) if pd.notna(idade_raw) else 0
    modalidade = str(df_cli['Modalidade'].iloc[0]) if 'Modalidade' in df_cli.columns else 'Comission Based'
    
    liq_saldo, liq_fundos, liq_rf = df_cli['Liquidez_Saldo'].max(), df_cli['Liquidez_Fundos'].max(), df_cli['Liquidez_RF'].max()
    liq_total = df_cli['LIQUIDEZ_TOTAL_360'].max()

    tags_bd = crm.obter_tags_cliente(id_cliente)
    badges_tags = [html.Span(t[0], style={'backgroundColor': PALETTE['bg_card'], 'color': PALETTE['bg_dark'], 'padding': '4px 12px', 'borderRadius': '12px', 'fontSize': '12px', 'fontWeight': 'bold'}) for t in tags_bd] if tags_bd else [html.Span("Nenhuma tag registrada", style={'color': PALETTE['text_label'], 'fontSize': '12px'})]

    def card_regra(titulo, valor):
        return html.Div(style={'backgroundColor': PALETTE['red_bg_light'], 'padding': '14px', 'borderRadius': '8px', 'border': f"1px solid {PALETTE['red_border_light']}"}, children=[
            html.Span(titulo, style={'fontSize': '11px', 'fontWeight': 'bold', 'color': '#9B2C2C', 'display': 'block'}),
            html.Span(valor, style={'fontSize': '16px', 'fontWeight': 'bold', 'color': '#742A2A', 'fontFamily': FONT_SERIF})
        ])

    cards = [
        card_regra("CP - Único Emissor (Máx 5%):", formatar_moeda_br(custodia_total * 0.05)),
        card_regra("Fundos Fechados (Máx 10%):", formatar_moeda_br(custodia_total * 0.10)),
        card_regra("COE - Único Emissor (Máx 10%):", formatar_moeda_br(custodia_total * 0.10)),
        card_regra("NTN-B >= 2040 (Máx 15%):", formatar_moeda_br(custodia_total * 0.15))
    ]

    fig = go.Figure()
    fig.add_trace(go.Bar(x=df_cli['Fator de Risco'], y=df_cli['Gap Under'], name='Gap Under', marker_color=PALETTE['green_soft']))
    fig.add_trace(go.Bar(x=df_cli['Fator de Risco'], y=df_cli['Gap Over'], name='Gap Over', marker_color=PALETTE['red_soft']))
    fig.update_layout(title='Gaps da Carteira', barmode='group', paper_bgcolor='#FFFFFF', plot_bgcolor='#FFFFFF', font=dict(family=FONT_SANS))

    return (
        {'display': 'block', 'marginTop': '30px'}, f"🎯 Cliente: {id_cliente}",
        f"| Modalidade: {modalidade} | Idade: {idade} | AUC: {formatar_moeda_br(custodia_total)}", id_cliente,
        badges_tags, formatar_moeda_br(liq_saldo), formatar_moeda_br(liq_fundos), formatar_moeda_br(liq_rf), formatar_moeda_br(liq_total),
        cards, fig, active_cell
    )

# REGISTRO DE TESE DIRETO NO RAIO-X
@app.callback(
    [Output('crm-feedback-registro', 'children'), Output('store-crm-updated', 'data', allow_duplicate=True)],
    Input('btn-salvar-crm', 'n_clicks'),
    [State('tabela-clientes-assessor', 'active_cell'), State('tabela-clientes-assessor', 'derived_virtual_data'),
     State('crm-fator-dropdown', 'value'), State('crm-produto-input', 'value'), State('crm-valor-input', 'value'),
     State('crm-roa-input', 'value'), State('crm-dias-dropdown', 'value'), State('crm-obs-input', 'value'),
     State('store-crm-updated', 'data')],
    prevent_initial_call='initial_duplicate'
)
def gravar_tese_crm(n_clicks, active_cell, table_data, fator_risco, produto, valor, roa_str, dias_followup, obs, store_val):
    if not n_clicks or not active_cell or not table_data:
        return html.Span("⚠️ Selecione um cliente na tabela primeiro.", style={'color': PALETTE['red_soft']}), store_val
        
    id_cliente = str(table_data[active_cell['row']]['Cliente'])
    df_cli = df_clean[df_clean['Cód. Conta'] == id_cliente]
    if df_cli.empty: return html.Span("Cliente não localizado.", style={'color': PALETTE['red_soft']}), store_val
    
    assessor, investor = df_cli['Assessor'].iloc[0], df_cli['Investor'].iloc[0]
    
    if not fator_risco or not produto or not valor:
        return html.Span("⚠️ Preencha Classe, Produto e Volume Alvo.", style={'color': PALETTE['red_soft']}), store_val
        
    valor_num = converter_valor_br(valor)
    roa_num = converter_valor_br(roa_str)
    
    if valor_num <= 0: return html.Span("⚠️ Informe um volume numérico válido.", style={'color': PALETTE['red_soft']}), store_val

    msg_retorno = crm.registrar_oportunidade(investor, assessor, id_cliente, fator_risco, produto, valor_num, roa=roa_num, dias_followup=dias_followup, obs=obs)
    return html.Span(msg_retorno, style={'color': PALETTE['gold_main']}), (store_val or 0) + 1

# =============================================================================
# CALLBACKS DA ABA 3: CRM TÁTICO & RECEITA PROJETADA
# =============================================================================
@app.callback(
    [Output('bloco-kpis-crm-pipeline', 'children'), Output('tabela-pipeline-crm', 'data')],
    [Input('abas-navegacao', 'value'), Input('crm-filtro-periodo', 'value'),
     Input('crm-filtro-status', 'value'), Input('crm-filtro-investor', 'value'), 
     Input('crm-filtro-assessor', 'value'), Input('store-crm-updated', 'data'),
     Input('btn-atualizar-crm', 'n_clicks')]
)
def carregar_dados_aba_crm(aba, periodo, status, investor, assessor, store_trigger, n_clicks_refresh):
    df_global = crm.obter_todas_oportunidades(status_filtro='Todos', investor_filtro=investor, assessor_filtro=assessor, periodo_filtro=periodo)
    
    if df_global.empty:
        vol_pipeline, rec_pipeline = 0.0, 0.0
        vol_convertido, rec_convertido = 0.0, 0.0
        tx_conversao = 0.0
    else:
        df_pen = df_global[df_global['status'] == 'Pendente']
        vol_pipeline = df_pen['valor_alvo'].sum()
        rec_pipeline = df_pen['receita_estimada'].sum()
        
        df_conv = df_global[df_global['status'] == 'Convertido']
        vol_convertido = df_conv['valor_convertido'].fillna(df_conv['valor_alvo']).sum()
        rec_convertido = df_conv['receita_estimada'].sum()
        
        tot_oferecido = vol_pipeline + vol_convertido + df_global[df_global['status'] == 'Recusado']['valor_alvo'].sum()
        tx_conversao = (vol_convertido / tot_oferecido) if tot_oferecido > 0 else 0.0

    cards_funil = html.Div(style={'display': 'flex', 'gap': '15px', 'flexWrap': 'wrap'}, children=[
        criar_card_kpi("Pipeline Vol. Pendente", formatar_moeda_br(vol_pipeline), PALETTE['bg_dark']),
        criar_card_kpi("Receita Est. Pipeline", formatar_moeda_br(rec_pipeline), PALETTE['gold_main']),
        criar_card_kpi("Vendas Convertidas R$", formatar_moeda_br(vol_convertido), PALETTE['green_soft']),
        criar_card_kpi("Receita Realizada (ROA)", formatar_moeda_br(rec_convertido), PALETTE['green_soft']),
        criar_card_kpi("Taxa Conversão Vol.", formatar_percentual_br(tx_conversao), PALETTE['gold_main'])
    ])

    df_tabela = df_global[df_global['status'] == status].copy() if status != 'Todos' else df_global.copy()
    return cards_funil, df_tabela.to_dict('records')

@app.callback(
    [Output('crm-detalhe-tarefa-selecionada', 'children'), Output('crm-baixa-valor-input', 'value'), Output('crm-baixa-roa-input', 'value')],
    Input('tabela-pipeline-crm', 'active_cell'),
    State('tabela-pipeline-crm', 'derived_virtual_data')
)
def preencher_formulario_baixa_crm(active_cell, table_data):
    if not active_cell or not table_data:
        return "💡 Clique em uma linha da tabela para selecionar.", "", ""
    
    linha = table_data[active_cell['row']]
    texto_detalhe = f"🎯 Oferta #{linha['id']} | Cliente: {linha['conta']} | Produto: {linha['produto_sugerido']} | Vol. Proposto: R$ {linha['valor_alvo']:,.2f} | ROA: {linha.get('roa', 0):,.2f}%"
    
    val_sugerido = linha.get('valor_convertido') if pd.notna(linha.get('valor_convertido')) else linha.get('valor_alvo')
    roa_sugerido = linha.get('roa', 0.0)
    
    return texto_detalhe, formatar_moeda_br(val_sugerido), f"{roa_sugerido:,.2f}".replace('.', ',')

@app.callback(
    [Output('crm-feedback-baixa', 'children'), Output('store-crm-updated', 'data', allow_duplicate=True)],
    Input('btn-confirmar-baixa-crm', 'n_clicks'),
    [State('tabela-pipeline-crm', 'active_cell'), State('tabela-pipeline-crm', 'derived_virtual_data'),
     State('crm-baixa-status-dropdown', 'value'), State('crm-baixa-valor-input', 'value'),
     State('crm-baixa-roa-input', 'value'), State('crm-baixa-motivo-dropdown', 'value'),
     State('crm-baixa-obs-input', 'value'), State('store-crm-updated', 'data')],
    prevent_initial_call='initial_duplicate'
)
def executar_baixa_crm(n_clicks, active_cell, table_data, novo_status, valor_convertido_str, roa_str, motivo_recusa, obs_devolutiva, store_val):
    if not n_clicks or not active_cell or not table_data:
        return html.Span("⚠️ Selecione uma oferta primeiro.", style={'color': PALETTE['red_soft']}), store_val
        
    id_opp = table_data[active_cell['row']]['id']
    valor_real_num = converter_valor_br(valor_convertido_str)
    roa_num = converter_valor_br(roa_str)
    
    crm.atualizar_status_oportunidade(id_opp, novo_status, motivo_recusa, obs_devolutiva, valor_convertido=valor_real_num, roa=roa_num)
    return html.Span(f"✅ Oferta #{id_opp} baixada com sucesso!", style={'color': PALETTE['gold_main']}), (store_val or 0) + 1

# =============================================================================
# CALLBACKS DA ABA 4: ACELERADOR COMERCIAL & CAMPANHAS
# =============================================================================
@app.callback(
    Output('campanha-produto-dropdown', 'options'),
    Input('abas-navegacao', 'value')
)
def carregar_produtos_foco_dropdown(aba):
    if os.path.exists(caminho_produtos_foco):
        df_pf = pd.read_excel(caminho_produtos_foco)
        if 'Produto Foco' in df_pf.columns:
            return [{'label': str(p), 'value': str(p)} for p in df_pf['Produto Foco'].dropna().unique()]
    return []

# FILTRO EM CASCATA: Limita os Assessores ao Investor selecionado
@app.callback(
    [Output('campanha-assessor-dropdown', 'options'), Output('campanha-assessor-dropdown', 'value')],
    Input('campanha-investor-dropdown', 'value')
)
def atualizar_assessores_campanha(investor_selecionado):
    if not investor_selecionado or investor_selecionado == 'Todos':
        assessores = lista_assessores
    else:
        assessores = sorted([str(a) for a in df_clean[df_clean['Investor'] == investor_selecionado]['Assessor'].unique() if str(a) != 'Sem Código'])
    
    options = [{'label': 'Todos', 'value': 'Todos'}] + [{'label': a, 'value': a} for a in assessores]
    return options, 'Todos'

@app.callback(
    [Output('bloco-kpis-campanha', 'children'), Output('tabela-resultado-campanha', 'data')],
    [Input('btn-rodar-campanha', 'n_clicks')],
    [State('campanha-produto-dropdown', 'value'), 
     State('campanha-investor-dropdown', 'value'), 
     State('campanha-assessor-dropdown', 'value')]
)
def executar_match_campanha(n_clicks, produto_selecionado, investor_filtro, assessor_filtro):
    if not os.path.exists(caminho_produtos_foco):
        return html.Div("Planilha Produtos_Foco.xlsx não encontrada."), []
        
    df_pf = pd.read_excel(caminho_produtos_foco)
    if df_pf.empty:
        return html.Div("Planilha de Produtos Foco está vazia."), []

    # Mapeamento de Produtos Foco por Fator de Risco
    # Seleciona o produto de maior ROA por classe caso haja múltiplos
    dict_produtos_foco = {}
    for fator, grp in df_pf.groupby('Fator de Risco'):
        dict_produtos_foco[fator] = grp.sort_values('ROA (%)', ascending=False).iloc[0].to_dict()

    # Filtragem inicial da base de clientes
    df_f = df_clean.copy()
    if investor_filtro and investor_filtro != 'Todos':
        df_f = df_f[df_f['Investor'] == investor_filtro]
    if assessor_filtro and assessor_filtro != 'Todos':
        df_f = df_f[df_f['Assessor'] == assessor_filtro]

    custodia_por_conta = df_clean.groupby('Cód. Conta')['Custódia'].sum().to_dict()
    dict_liquidez = df_liquidez_contas.set_index('Cód. Conta')['LIQUIDEZ_TOTAL_360'].to_dict()

    oportunidades_cestas = []
    PERCENTUAL_MAX_LIQUIDEZ_USO = 0.80  # Trava: utiliza no máximo 80% da liquidez D+5 total do cliente

    # Processamento Conta a Conta (Análise Global dos Gaps do Cliente)
    for conta, grp_cliente in df_f.groupby('Cód. Conta'):
        liq_d5 = dict_liquidez.get(conta, 0.0)
        auc_total = custodia_por_conta.get(conta, 0.0)
        
        if liq_d5 <= 0 or auc_total <= 0:
            continue

        # Teto de liquidez alocável para a cesta (preserva reserva de segurança)
        liquidez_disponivel_cesta = liq_d5 * PERCENTUAL_MAX_LIQUIDEZ_USO

        # Identifica todos os Gaps Under > 0 do cliente
        gaps_under = grp_cliente[grp_cliente['Gap Under'] > 0].sort_values('Gap Under', ascending=False)
        if gaps_under.empty:
            continue

        itens_cesta = []
        volume_total_cesta = 0.0
        receita_total_cesta = 0.0

        for _, row_gap in gaps_under.iterrows():
            fator_risco = row_gap['Fator de Risco']
            gap_val = row_gap['Gap Under']

            # Se houver filtro de produto específico na tela, ignora outras classes
            if produto_selecionado and produto_selecionado != 'Todos':
                prod_foco_classe = df_pf[df_pf['Produto Foco'] == produto_selecionado]
                if prod_foco_classe.empty or prod_foco_classe.iloc[0]['Fator de Risco'] != fator_risco:
                    continue
                prod_info = prod_foco_classe.iloc[0].to_dict()
            else:
                if fator_risco not in dict_produtos_foco:
                    continue
                prod_info = dict_produtos_foco[fator_risco]

            nome_prod = prod_info['Produto Foco']
            tipo_veiculo = prod_info.get('Tipo Veículo', 'Investimentos')
            cap_pct = float(prod_info.get('Cap Compliance (% AUC)', 1.0))
            ticket_min = float(prod_info.get('Ticket Mínimo (R$)', 0.0))
            tese_prod = prod_info.get('Tese Comercial / Destaques', '')

            # Tratamento de ROA %
            roa_raw = prod_info.get('ROA (%)', 0.0)
            try:
                roa_val = float(roa_raw)
            except (ValueError, TypeError):
                roa_val = 0.0
            roa_decimal = (roa_val / 100.0) if roa_val > 1.0 else roa_val

            # Limite de Compliance Institucional para o produto
            limite_compliance = auc_total * cap_pct if cap_pct < 1.0 else auc_total

            # Cálculo do Volume Sugerido considerando o saldo disponível da cesta
            vol_sugerido = min(gap_val, limite_compliance, liquidez_disponivel_cesta)

            if vol_sugerido >= ticket_min and vol_sugerido > 0:
                receita_item = vol_sugerido * roa_decimal
                
                itens_cesta.append({
                    'fator': fator_risco,
                    'produto': nome_prod,
                    'tipo': tipo_veiculo,
                    'gap': gap_val,
                    'vol_sugerido': vol_sugerido,
                    'roa_decimal': roa_decimal,
                    'receita': receita_item,
                    'tese': tese_prod
                })

                volume_total_cesta += vol_sugerido
                receita_total_cesta += receita_item
                liquidez_disponivel_cesta -= vol_sugerido  # Deduz da liquidez alocável restante

                if liquidez_disponivel_cesta <= 0:
                    break  # Encerra a inclusão na cesta se a liquidez atingir o teto de segurança

        if not itens_cesta:
            continue

        assessor = grp_cliente['Assessor'].iloc[0]
        liq_remanescente = liq_d5 - volume_total_cesta
        roa_medio_cesta = (receita_total_cesta / volume_total_cesta) if volume_total_cesta > 0 else 0.0

        # FORMATAÇÃO DA MENSAGEM ÚNICA PARA WHATSAPP
        qtd_gaps = len(itens_cesta)
        s_gaps = "Gap" if qtd_gaps == 1 else "Gaps"
        
        msg_lines = [
            f"🎯 *Oportunidade Comercial - Readequação de Carteira*",
            f"📌 *Conta:* {conta}",
            f"",
            f"Identificamos a oportunidade de cobrir *{qtd_gaps} {s_gaps} de Alocação* utilizando parte da liquidez disponível (*R$ {liq_d5:,.0f}*):\n".replace(',', '.')
        ]

        for idx, item in enumerate(itens_cesta, 1):
            vol_str = f"R$ {item['vol_sugerido']:,.0f}".replace(',', '.')
            gap_str = f"R$ {item['gap']:,.0f}".replace(',', '.')
            
            msg_lines.append(f"*{idx}️⃣ {item['fator']}* (GAP: {gap_str})")
            msg_lines.append(f"👉 Sugestão: Alocar até *{vol_str}* no *{item['produto']}* ({item['tese']}).\n")

        msg_lines.append(f"📊 *Resumo da Proposta Combinada:*")
        msg_lines.append(f"• *Alocação Total Sugerida:* R$ {volume_total_cesta:,.0f}".replace(',', '.'))
        msg_lines.append(f"• *Liquidez Remanescente D+5:* R$ {liq_remanescente:,.0f} (Reserva mantida)".replace(',', '.'))
        
        msg_whatsapp = "\n".join(msg_lines)

        # Resumo dos produtos para exibição sintética na tabela
        produtos_nomes = " + ".join([i['produto'] for i in itens_cesta])
        gaps_nomes = " + ".join([i['fator'] for i in itens_cesta])

        oportunidades_cestas.append({
            'Cód. Conta': conta,
            'Assessor': assessor,
            'Produto Foco': produtos_nomes,
            'Fator de Risco': gaps_nomes,
            'Gap Classe': sum([i['gap'] for i in itens_cesta]),
            'Liquidez D+5': liq_d5,
            'Volume Sugerido': volume_total_cesta,
            'ROA (%)': roa_medio_cesta,
            'Receita Estimada': receita_total_cesta,
            'Mensagem WhatsApp': msg_whatsapp
        })

    df_res = pd.DataFrame(oportunidades_cestas)
    if df_res.empty:
        return html.Div("Nenhuma cesta de oportunidades atendeu aos critérios de ticket mínimo/liquidez."), []

    tot_vol = df_res['Volume Sugerido'].sum()
    tot_rec = df_res['Receita Estimada'].sum()
    total_leads = len(df_res)

    kpis_campanha = html.Div(style={'display': 'flex', 'gap': '15px', 'flexWrap': 'wrap'}, children=[
        criar_card_kpi("Cestas Mapeadas", f"{total_leads} clientes", PALETTE['bg_dark']),
        criar_card_kpi("Volume Total Mapeado", formatar_moeda_br(tot_vol), PALETTE['green_soft']),
        criar_card_kpi("Receita Potencial (ROA)", formatar_moeda_br(tot_rec), PALETTE['gold_main'])
    ])

    return kpis_campanha, df_res.to_dict('records')

@app.callback(
    [Output('campanha-preview-texto', 'children'), Output('link-abrir-wa', 'href')],
    Input('tabela-resultado-campanha', 'active_cell'),
    State('tabela-resultado-campanha', 'derived_virtual_data')
)
def selecionar_mensagem_whatsapp(active_cell, table_data):
    if not active_cell or not table_data:
        return "Clique em uma linha da tabela ao lado para carregar o texto formatado do WhatsApp.", "#"
        
    msg = table_data[active_cell['row']]['Mensagem WhatsApp']
    msg_encoded = urllib.parse.quote(msg)
    link_wa = f"https://api.whatsapp.com/send?text={msg_encoded}"
    
    return msg, link_wa

@app.callback(
    [Output('feedback-push-crm', 'children'), Output('store-crm-updated', 'data', allow_duplicate=True)],
    Input('btn-push-crm-campanha', 'n_clicks'),
    [State('tabela-resultado-campanha', 'active_cell'), State('tabela-resultado-campanha', 'derived_virtual_data'), State('store-crm-updated', 'data')],
    prevent_initial_call='initial_duplicate'
)
def disparar_oportunidade_para_crm(n_clicks, active_cell, table_data, store_val):
    if not n_clicks or not active_cell or not table_data:
        return html.Span("⚠️ Selecione uma linha da tabela ao lado primeiro.", style={'color': PALETTE['red_soft']}), store_val
        
    linha = table_data[active_cell['row']]
    conta = linha['Cód. Conta']
    assessor = linha['Assessor']
    df_cli = df_clean[df_clean['Cód. Conta'] == conta]
    investor = df_cli['Investor'].iloc[0] if not df_cli.empty else 'Mesa'
    
    produto = linha['Produto Foco']
    fator = linha['Fator de Risco']
    valor = linha['Volume Sugerido']
    
    roa_val_dash = linha.get('ROA (%)', 0.0)
    roa_final = roa_val_dash * 100.0 if roa_val_dash <= 1.0 else roa_val_dash
    
    obs = f"Campanha Automática: {produto}"
    
    msg_retorno = crm.registrar_oportunidade(investor, assessor, conta, fator, produto, valor, roa=roa_final, dias_followup=3, obs=obs)
    return html.Span(f"🚀 Oportunidade enviada ao CRM! {msg_retorno}", style={'color': PALETTE['gold_main']}), (store_val or 0) + 1

# EXPORTAÇÕES EXCEL
@app.callback(
    Output("download-excel-geral-mesa", "data"),
    Input("btn-export-geral-mesa", "n_clicks"),
    [State('mesa-investor-dropdown', 'value'), State('mesa-classe-dropdown', 'value'),
     State('mesa-pessoa-dropdown', 'value'), State('mesa-modalidade-dropdown', 'value')],
    prevent_initial_call=True
)
def exportar_geral_mesa(n_clicks, investor, classe, pessoa, modalidade):
    if not n_clicks: return None
    df_filtrado = df_clean.copy()
    if investor != 'Todos': df_filtrado = df_filtrado[df_filtrado['Investor'] == investor]
    if classe != 'Todas': df_filtrado = df_filtrado[df_filtrado['Fator de Risco'] == classe]
    if pessoa != 'Todas': df_filtrado = df_filtrado[df_filtrado['Tipo_Pessoa'] == pessoa]
    if modalidade != 'Todas': df_filtrado = df_filtrado[df_filtrado['Modalidade'] == modalidade]

    return dcc.send_data_frame(df_filtrado.to_excel, filename=f"Resumo_Mesa_{investor}.xlsx", index=False)

@app.callback(
    Output("download-excel-geral-assessor", "data"),
    Input("btn-export-geral-assessor-aba2", "n_clicks"),
    [State('cockpit-assessor-dropdown', 'value'), State('cockpit-classe-dropdown', 'value'),
     State('cockpit-cliente-dropdown', 'value'), State('cockpit-pessoa-dropdown', 'value'), State('cockpit-modalidade-dropdown', 'value')],
    prevent_initial_call=True
)
def exportar_geral_assessor(n_clicks, assessor, classe, cliente, pessoa, modalidade):
    if not n_clicks: return None
    df_f = df_clean.copy()
    if assessor and assessor != 'Todos': df_f = df_f[df_f['Assessor'] == str(assessor)]
    if classe and classe != 'Todas': df_f = df_f[df_f['Fator de Risco'] == classe]
    if pessoa and pessoa != 'Todas': df_f = df_f[df_f['Tipo_Pessoa'] == pessoa]
    if modalidade and modalidade != 'Todas': df_f = df_f[df_f['Modalidade'] == modalidade]
    if cliente and str(cliente).strip() not in ['Todas', 'None', '', 'null']: df_f = df_f[df_f['Cód. Conta'] == str(cliente)]

    return dcc.send_data_frame(df_f.to_excel, filename=f"Base_Assessor_{assessor}.xlsx", index=False)

if __name__ == '__main__':
    app.run(debug=True, dev_tools_ui=False)