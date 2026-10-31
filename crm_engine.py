# =============================================================================
# MÓDULO BACKEND - CRM TÁTICO & PIPELINE DE VENDAS (UNIFICADO & COMPATÍVEL)
# =============================================================================

import sqlite3
import pandas as pd
from datetime import datetime
import os
import re
import pythoncom
import win32com.client as win32

# =============================================================================
# DIRETÓRIO E RESTAURAÇÃO DE BANCO (crm_pipeline.db + MIGRAÇÃO crm_database.db)
# =============================================================================
PASTA_USUARIO = os.path.expanduser("~")
CAMINHO_PASTA_ONEDRIVE = os.path.join(
    PASTA_USUARIO, 
    "OneDrive - NUCLEOPAR AGENTE AUTONOMO DE INVESTIMENTO LTDA", 
    "Automação", 
    "Cockpit do Investor"
)

try:
    os.makedirs(CAMINHO_PASTA_ONEDRIVE, exist_ok=True)
    CAMINHO_PASTA = CAMINHO_PASTA_ONEDRIVE
except Exception:
    CAMINHO_PASTA = os.path.dirname(os.path.abspath(__file__))

# Nomes de banco de dados
DB_PATH = os.path.join(CAMINHO_PASTA, "crm_pipeline.db")
DB_ANTIGO_TEMP = os.path.join(CAMINHO_PASTA, "crm_database.db")
DB_LOCAL_PASTA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "crm_pipeline.db")


def calcular_data_followup_dias_uteis(dias_uteis=3):
    """Calcula dias úteis utilizando o motor de Business Days do Pandas."""
    hoje = pd.Timestamp.now()
    hoje_inicio_dia = hoje.floor('D')
    dt_target = (hoje_inicio_dia + pd.offsets.BDay(int(dias_uteis))).to_pydatetime()
    return hoje, dt_target


def inicializar_banco():
    """Garante o schema correto e migra dados de bancos alternativos/antigos se existirem."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Tabela 1: Pipeline de Oportunidades Comercial
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS oportunidades (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data_criacao TEXT,
            investor TEXT,
            assessor TEXT,
            conta TEXT,
            fator_risco TEXT,
            produto_sugerido TEXT,
            valor_alvo REAL,
            valor_convertido REAL,
            roa REAL DEFAULT 0.0,
            receita_estimada REAL DEFAULT 0.0,
            data_followup TEXT,
            status TEXT DEFAULT 'Pendente',
            motivo_recusa TEXT,
            observacao TEXT
        )
    ''')
    
    # Garantia de colunas novas em schemas antigos
    for col, tipo in [('valor_convertido', 'REAL'), ('roa', 'REAL DEFAULT 0.0'), ('receita_estimada', 'REAL DEFAULT 0.0')]:
        try:
            cursor.execute(f"ALTER TABLE oportunidades ADD COLUMN {col} {tipo}")
        except Exception:
            pass
            
    # Tabela 2: Tags Qualitativas de Perfil
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS tags_cliente (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            conta TEXT,
            tag TEXT,
            categoria TEXT,
            origem_observacao TEXT,
            data_identificacao TEXT,
            UNIQUE(conta, tag)
        )
    ''')
    
    conn.commit()
    
    # MIGRAÇÃO AUTOMÁTICA: Se houver dados no crm_database.db, migra para o crm_pipeline.db
    if os.path.exists(DB_ANTIGO_TEMP) and DB_ANTIGO_TEMP != DB_PATH:
        try:
            conn_temp = sqlite3.connect(DB_ANTIGO_TEMP)
            df_temp = pd.read_sql_query("SELECT * FROM oportunidades", conn_temp)
            conn_temp.close()
            
            if not df_temp.empty:
                for _, r in df_temp.iterrows():
                    val_alvo = r.get('valor_alvo', 0.0)
                    val_conv = r.get('valor_convertido', val_alvo)
                    roa_val = r.get('roa', 0.0)
                    rec_val = r.get('receita_estimada', (val_conv or 0.0) * ((roa_val or 0.0) / 100.0))
                    
                    cursor.execute('''
                        INSERT OR IGNORE INTO oportunidades 
                        (data_criacao, investor, assessor, conta, fator_risco, produto_sugerido, valor_alvo, valor_convertido, roa, receita_estimada, data_followup, status, motivo_recusa, observacao)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (r.get('data_criacao'), r.get('investor'), r.get('assessor'), r.get('conta'), r.get('fator_risco'), 
                          r.get('produto_sugerido'), val_alvo, val_conv, roa_val, rec_val, r.get('data_followup'), 
                          r.get('status'), r.get('motivo_recusa'), r.get('observacao')))
                conn.commit()
        except Exception:
            pass

    conn.close()

# =============================================================================
# DICIONÁRIO E MOTOR NLP LOCAL
# =============================================================================
DICIONARIO_NLP = {
    r'curto prazo|liquidez|resgate r[aá]pido|d\+0|d\+1|pouco tempo|n[aã]o quer travar|1 ano|6 meses|dinheiro na m[aã]o': 
        ('Perfil: Curto Prazo / Alta Liquidez', 'Horizonte'),
    r'longo prazo|aposentadoria|futuro|n[aã]o precisa de liquidez|pode travar|2030|2035|2040|car[êe]ncia longa': 
        ('Perfil: Longo Prazo', 'Horizonte'),
    r'n[aã]o quer travar|fundo fechado n[aã]o|medo de car[êe]ncia|odeia iliquidez|il[íi]quido': 
        ('Restrição: Iliquidez / Trava', 'Liquidez'),
    r'n[aã]o gosta de cr[eé]dito privado|foge de deb[êe]nture|sem cr[eé]dito privado|s[oó] t[íi]tulo p[úu]blico|s[oó] banc[aá]rio|medo de calote': 
        ('Restrição: Crédito Privado', 'Risco Crédito'),
    r'high yield|cr[eé]dito arriscado|rating baixo|sem rating|high-yield|empresa quebrada': 
        ('Restrição: High Yield / Rating Baixo', 'Risco Crédito'),
    r's[oó] cdb|quer fgc|bancos grandes|emiss[ãao] banc[aá]ria|garantia do fgc': 
        ('Foco: Emissão Bancária / FGC', 'Emissor'),
    r'n[aã]o quer risco|medo de perder|oscila[çc][aã]o|marca[çc][aã]o a mercado|volatilidade|agressivo mas conservador|n[aã]o aguenta cair': 
        ('Restrição: Volatilidade / MTM', 'Perfil Risco'),
    r'n[aã]o gosta de a[çc][oõ]es|foge de bolsa|n[aã]o quer rv|renda vari[aá]vel n[aã]o|decep[çc][aã]o com bolsa': 
        ('Aversão: Renda Variável', 'Classe Ativos'),
    r'n[aã]o quer rendimento mensal|sem proventos|sem cupom|sem cupons|prefere ac[úu]mulo|ganho de capital|reinvestir': 
        ('Restrição: Proventos Mensais (Prefere Acúmulo)', 'Fluxo de Caixa'),
    r'quer renda mensal|precisa de cupom|viver de renda|proventos mensais|pingar todo m[êe]s': 
        ('Foco: Renda Mensal / Proventos', 'Fluxo de Caixa'),
    r'isento|isen[çc][aã]o|sem ir|deb[êe]nture incentivada|lca|lci|fundo incentivado|efici[êe]ncia fiscal': 
        ('Foco: Isenção Fiscal', 'Tributação'),
    r'taxa alta|fee salgado|muito caro|performance alta|fee excessivo|comiss[ãao] alta': 
        ('Sensibilidade: Taxas / Fee', 'Custos')
}

def extrair_e_aplicar_tags(conta, texto_observacao):
    if not texto_observacao or not str(texto_observacao).strip():
        return
    
    texto = str(texto_observacao).lower()
    tags_detectadas = []
    
    for padrao, (tag, categoria) in DICIONARIO_NLP.items():
        if re.search(padrao, texto):
            tags_detectadas.append((tag, categoria))
            
    if tags_detectadas:
        inicializar_banco()
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        hoje = datetime.now().strftime('%Y-%m-%d')
        
        for tag, categoria in tags_detectadas:
            cursor.execute('''
                INSERT OR IGNORE INTO tags_cliente (conta, tag, categoria, origem_observacao, data_identificacao)
                VALUES (?, ?, ?, ?, ?)
            ''', (str(conta), tag, categoria, str(texto_observacao), hoje))
            
        conn.commit()
        conn.close()

# =============================================================================
# AGENDAMENTO OUTLOOK (THREAD-SAFE)
# =============================================================================
def agendar_compromisso_outlook(assessor, conta, produto, valor, roa=0.0, data_followup_dt=None, obs=""):
    pythoncom.CoInitialize()
    try:
        try:
            outlook = win32.GetActiveObject("Outlook.Application")
        except Exception:
            try:
                outlook = win32.Dispatch("Outlook.Application")
            except Exception as e_dispatch:
                return False, f"Outlook não detectado ({e_dispatch})."

        appointment = outlook.CreateItem(1)
        if data_followup_dt is None:
            _, data_followup_dt = calcular_data_followup_dias_uteis(3)
            
        dt_inicio = data_followup_dt.replace(hour=9, minute=0, second=0, microsecond=0)
        rec_est = float(valor or 0.0) * (float(roa or 0.0) / 100.0)
        
        appointment.Start = dt_inicio.strftime('%Y-%m-%d %H:%M')
        appointment.Duration = 30
        appointment.Subject = f"🎯 [CRM Follow-up] Oferta {produto} - Cliente {conta} (Assessor {assessor})"
        appointment.Body = f"""
==================================================
🎯 COCKPIT DO INVESTOR - LEMBRETE DE FOLLOW-UP CRM
==================================================

Cliente (Conta): {conta}
Assessor Responsável: {assessor}
Produto Recomendado: {produto}
Volume Estimado: R$ {float(valor or 0.0):,.2f}
ROA Estimado: {float(roa or 0.0):,.2f}%
Receita Projetada: R$ {rec_est:,.2f}

Observações / Tese Comercial:
"{obs}"

--------------------------------------------------
Instrução: Acesse a aba 'CRM Tático & Pipeline' no Dashboard 
para registrar o retorno do assessor e dar baixa nesta tarefa.
"""
        appointment.ReminderSet = True
        appointment.ReminderMinutesBeforeStart = 15
        appointment.Save()
        return True, "Lembrete cadastrado na Agenda do Outlook com sucesso!"
    except Exception as e:
        return False, f"Falha ao agendar no Outlook: {e}"
    finally:
        pythoncom.CoUninitialize()

# =============================================================================
# OPERAÇÕES DO PIPELINE
# =============================================================================
def registrar_oportunidade(investor, assessor, conta, fator_risco, produto, valor, roa=0.0, dias_followup=3, obs=""):
    inicializar_banco()
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    hoje, dt_followup = calcular_data_followup_dias_uteis(dias_followup)
    
    data_criacao_iso = hoje.strftime('%Y-%m-%d %H:%M')
    data_followup_iso = dt_followup.strftime('%Y-%m-%d')
    data_followup_br = dt_followup.strftime('%d/%m/%Y')
    
    val_num = float(valor) if valor else 0.0
    roa_num = float(roa) if roa else 0.0
    rec_num = val_num * (roa_num / 100.0)
    
    cursor.execute('''
        INSERT INTO oportunidades (data_criacao, investor, assessor, conta, fator_risco, produto_sugerido, valor_alvo, valor_convertido, roa, receita_estimada, data_followup, status, observacao)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Pendente', ?)
    ''', (data_criacao_iso, str(investor), str(assessor), str(conta), str(fator_risco), str(produto), val_num, val_num, roa_num, rec_num, data_followup_iso, str(obs)))
    
    conn.commit()
    conn.close()
    
    if obs:
        extrair_e_aplicar_tags(conta, obs)
        
    sucesso_out, msg_out = agendar_compromisso_outlook(assessor, conta, produto, val_num, roa_num, dt_followup, obs)
    return f"✅ Oportunidade salva! Follow-up agendado para {data_followup_br}. {msg_out}"

def atualizar_status_oportunidade(id_opp, novo_status, motivo_recusa="", observacao_extra="", valor_convertido=None, roa=None):
    inicializar_banco()
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute("SELECT conta, valor_alvo, roa FROM oportunidades WHERE id = ?", (id_opp,))
    res = cursor.fetchone()
    if not res:
        conn.close()
        return False
        
    conta, valor_alvo, roa_bd = res[0], res[1], res[2]
    
    roa_final = float(roa) if (roa is not None and float(roa) >= 0) else (float(roa_bd) if roa_bd else 0.0)
    
    if valor_convertido is not None and float(valor_convertido) > 0:
        val_conv = float(valor_convertido)
    else:
        val_conv = float(valor_alvo) if novo_status == 'Convertido' else None

    rec_final = (val_conv * (roa_final / 100.0)) if val_conv else 0.0
    obs_final = str(observacao_extra).strip() if observacao_extra else ""
    
    cursor.execute('''
        UPDATE oportunidades 
        SET status = ?, motivo_recusa = ?, observacao = COALESCE(observacao, '') || ' | Devolutiva: ' || ?, 
            valor_convertido = ?, roa = ?, receita_estimada = ?
        WHERE id = ?
    ''', (str(novo_status), str(motivo_recusa), obs_final, val_conv, roa_final, rec_final, id_opp))
    
    conn.commit()
    conn.close()
    
    texto_para_nlp = f"{motivo_recusa} {observacao_extra}"
    if conta and texto_para_nlp.strip():
        extrair_e_aplicar_tags(conta, texto_para_nlp)
        
    return True

def obter_todas_oportunidades(status_filtro='Todos', investor_filtro='Todos', assessor_filtro='Todos', periodo_filtro='mes_atual'):
    inicializar_banco()
    conn = sqlite3.connect(DB_PATH)
    
    query = "SELECT * FROM oportunidades WHERE 1=1"
    params = []
    
    if status_filtro and status_filtro != 'Todos':
        query += " AND status = ?"
        params.append(status_filtro)
        
    if investor_filtro and investor_filtro != 'Todos':
        query += " AND investor = ?"
        params.append(investor_filtro)
        
    if assessor_filtro and assessor_filtro != 'Todos':
        query += " AND assessor = ?"
        params.append(assessor_filtro)

    if periodo_filtro == 'mes_atual':
        mes_atual = datetime.now().strftime('%Y-%m')
        query += " AND strftime('%Y-%m', data_criacao) = ?"
        params.append(mes_atual)
        
    query += " ORDER BY id DESC"
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()

    if df.empty:
        # Retorna DataFrame com a estrutura de colunas exata esperada pela tabela Dash
        return pd.DataFrame(columns=[
            'id', 'data_criacao', 'investor', 'assessor', 'conta', 'fator_risco', 
            'produto_sugerido', 'valor_alvo', 'valor_convertido', 'roa', 
            'receita_estimada', 'data_followup', 'status', 'motivo_recusa', 'observacao'
        ])

    # Preenchimento de colunas para dados legados
    if 'valor_convertido' not in df.columns: df['valor_convertido'] = df['valor_alvo']
    if 'roa' not in df.columns: df['roa'] = 0.0
    if 'receita_estimada' not in df.columns: df['receita_estimada'] = 0.0
    
    df['valor_convertido'] = df['valor_convertido'].fillna(df['valor_alvo'])
    df['roa'] = df['roa'].fillna(0.0)
    df['receita_estimada'] = df['receita_estimada'].fillna(df['valor_convertido'] * (df['roa'] / 100.0))

    if 'data_criacao' in df.columns:
        df['data_criacao'] = pd.to_datetime(df['data_criacao'], errors='coerce').dt.strftime('%d/%m/%Y %H:%M').fillna(df['data_criacao'])
    if 'data_followup' in df.columns:
        df['data_followup'] = pd.to_datetime(df['data_followup'], errors='coerce').dt.strftime('%d/%m/%Y').fillna(df['data_followup'])

    return df

def obter_tags_cliente(conta):
    inicializar_banco()
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT tag, categoria FROM tags_cliente WHERE conta = ?", (str(conta),))
    rows = cursor.fetchall()
    conn.close()
    return rows

if __name__ == '__main__':
    inicializar_banco()
    print(f"✅ Conectado com sucesso ao banco histórico: {DB_PATH}")