import streamlit as st
import pandas as pd
import plotly.express as px
import datetime
from dateutil.relativedelta import relativedelta
from io import BytesIO
import pyodbc
import json
import os
import uuid
from dotenv import load_dotenv
import numpy as np

# --- Configuração Inicial ---
load_dotenv()

st.set_page_config(layout="wide", page_title="Dashboard de Vendas e Agenda")

ARQUIVO_NOTAS = "notas_supervisor.json"
PASTA_IMAGENS = "notas_imagens"  # Pasta para salvar as fotos

# Garante que a pasta de imagens existe
if not os.path.exists(PASTA_IMAGENS):
    os.makedirs(PASTA_IMAGENS)

# --- Funções Auxiliares (Persistência e Excel) ---

def carregar_notas():
    if not os.path.exists(ARQUIVO_NOTAS):
        return {}
    try:
        with open(ARQUIVO_NOTAS, "r", encoding="utf-8") as f:
            dados = json.load(f)
        
        # Migração automática (str -> list) se necessário
        novo_formato = {}
        alterou = False
        for k, v in dados.items():
            if isinstance(v, str): 
                novo_formato[k] = [{
                    "id": str(uuid.uuid4()),
                    "data": datetime.datetime.now().strftime("%d/%m/%Y %H:%M"),
                    "texto": v,
                    "imagem": None # Garante compatibilidade
                }]
                alterou = True
            else:
                novo_formato[k] = v
        
        if alterou:
            with open(ARQUIVO_NOTAS, "w", encoding="utf-8") as f:
                json.dump(novo_formato, f, ensure_ascii=False, indent=4)
                
        return novo_formato
    except Exception as e:
        st.error(f"Erro ao carregar notas: {e}")
        return {}

def adicionar_nota(cod_cliente, texto, arquivo_imagem=None):
    notas_dict = carregar_notas()
    cod = str(cod_cliente)
    
    caminho_imagem = None
    
    # Lógica de Salvamento da Imagem
    if arquivo_imagem is not None:
        try:
            extensao = arquivo_imagem.name.split('.')[-1]
            nome_arquivo = f"{uuid.uuid4()}.{extensao}"
            caminho_completo = os.path.join(PASTA_IMAGENS, nome_arquivo)
            
            with open(caminho_completo, "wb") as f:
                f.write(arquivo_imagem.getbuffer())
            
            caminho_imagem = caminho_completo
        except Exception as e:
            st.error(f"Erro ao salvar imagem: {e}")

    nova_nota = {
        "id": str(uuid.uuid4()),
        "data": datetime.datetime.now().strftime("%d/%m/%Y %H:%M"),
        "texto": texto,
        "imagem": caminho_imagem # Salva o caminho
    }
    
    if cod not in notas_dict:
        notas_dict[cod] = []
    notas_dict[cod].insert(0, nova_nota)
    
    with open(ARQUIVO_NOTAS, "w", encoding="utf-8") as f:
        json.dump(notas_dict, f, ensure_ascii=False, indent=4)

def excluir_nota(cod_cliente, note_id):
    notas_dict = carregar_notas()
    cod = str(cod_cliente)
    if cod in notas_dict:
        # Tenta remover imagem do disco se existir
        nota_a_remover = next((n for n in notas_dict[cod] if n['id'] == note_id), None)
        if nota_a_remover and nota_a_remover.get('imagem'):
            if os.path.exists(nota_a_remover['imagem']):
                try:
                    os.remove(nota_a_remover['imagem'])
                except:
                    pass

        notas_dict[cod] = [n for n in notas_dict[cod] if n['id'] != note_id]
        if not notas_dict[cod]: 
            del notas_dict[cod]
        with open(ARQUIVO_NOTAS, "w", encoding="utf-8") as f:
            json.dump(notas_dict, f, ensure_ascii=False, indent=4)

@st.cache_data
def to_excel(df, sheet_name='Resumo'):
    output = BytesIO()
    df_export = df.copy()
    for col in df_export.select_dtypes(include=['datetime64[ns]']).columns:
        df_export[col] = df_export[col].dt.strftime('%d/%m/%Y %H:%M').replace('NaT', '')
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name=sheet_name)
    processed_data = output.getvalue()
    return processed_data

def to_excel_com_imagens(df_notas):
    """Gera um Excel onde a coluna 'Imagem' exibe a foto visualmente."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export = df_notas.drop(columns=['CaminhoImagem'], errors='ignore')
        df_export.to_excel(writer, sheet_name='Notas', index=False)
        
        workbook = writer.book
        worksheet = writer.sheets['Notas']
        
        col_imagem_idx = len(df_export.columns) 
        worksheet.write(0, col_imagem_idx, "Foto Anexada")
        
        text_wrap = workbook.add_format({'text_wrap': True, 'valign': 'top'})
        worksheet.set_column(0, col_imagem_idx - 1, 20, text_wrap)
        worksheet.set_column(col_imagem_idx, col_imagem_idx, 30)
        
        for i, row in df_notas.iterrows():
            caminho_img = row.get('CaminhoImagem')
            worksheet.set_row(i + 1, 100)
            
            if caminho_img and pd.notna(caminho_img) and os.path.exists(caminho_img):
                try:
                    worksheet.insert_image(
                        i + 1, col_imagem_idx, 
                        caminho_img, 
                        {'x_scale': 0.2, 'y_scale': 0.2, 'object_position': 1}
                    )
                except:
                    worksheet.write(i + 1, col_imagem_idx, "Erro img")
            else:
                worksheet.write(i + 1, col_imagem_idx, "-")
                
    return output.getvalue()

# --- Lógica de Resumo (Power Query Style) ---
def calcular_resumo_power_query(df_input_rows, df_base_full, df_info_clientes):
    df_structure = df_input_rows[['UsuarioEncer', 'CodClien', 'Cliente']].drop_duplicates()
    
    if df_structure.empty:
        return pd.DataFrame(), []

    df_hist_pivot = df_input_rows[df_input_rows['Sit'] == 'EN'].copy()
    
    replacements_m = {'COMPROU CONCORRENT': 'COMPROU CONCORRENTE', 'COMPROU CONCORRENTEE': 'COMPROU CONCORRENTE', 'COTACAO': 'COTAÇÃO'}
    df_hist_pivot['Motivo2'] = df_hist_pivot['Motivo_Final'].replace(replacements_m)
    
    pivot_table = pd.pivot_table(
        df_hist_pivot,
        index=['UsuarioEncer', 'CodClien'],
        columns='Motivo2',
        aggfunc='size',
        fill_value=0
    ).reset_index()
    
    hoje = datetime.date.today()
    inicio_mes_atual = hoje.replace(day=1)
    ts_inicio_mes_atual = pd.Timestamp(inicio_mes_atual)
    
    df_contexto = df_base_full[
        (df_base_full['CodClien'].isin(df_structure['CodClien'])) & 
        (df_base_full['UsuarioEncer'].isin(df_structure['UsuarioEncer']))
    ]

    df_mes = df_contexto[df_contexto['DataEncer'] >= ts_inicio_mes_atual].groupby(['CodClien', 'UsuarioEncer']).size().reset_index(name='Total Mês')
    
    inicio_3m = (inicio_mes_atual - relativedelta(months=3))
    ts_inicio_3m = pd.Timestamp(inicio_3m)
    df_3m_full = df_contexto[(df_contexto['DataEncer'] >= ts_inicio_3m) & (df_contexto['DataEncer'] < ts_inicio_mes_atual)]
    df_3m = df_3m_full.groupby(['CodClien', 'UsuarioEncer']).size().reset_index(name='Ult 3 Meses')
    
    df_total = df_input_rows.groupby(['CodClien', 'UsuarioEncer']).size().reset_index(name='Total de Contatos')
    
    df_dt_ped = df_base_full.groupby('CodClien')['DtUltPed'].max().reset_index()
    
    df_last = df_input_rows.sort_values('DataEncer', ascending=False).drop_duplicates(['CodClien', 'UsuarioEncer'])[['CodClien', 'UsuarioEncer', 'DataEncer', 'Motivo_Final', 'Obs']]
    df_last.columns = ['CodClien', 'UsuarioEncer', 'Data Ult Contato', 'Ult Resultado', 'Ult Obs']
    
    df_prox = df_base_full[df_base_full['Sit']=='AB'].copy()
    df_prox = df_prox.sort_values('DataAgenda')
    df_prox = df_prox.drop_duplicates(subset=['CodClien', 'UsuarioAgenda'])
    df_prox = df_prox[['CodClien', 'UsuarioAgenda', 'DataAgenda']]
    df_prox.columns = ['CodClien', 'UsuarioEncer', 'Próx Agendamento']
    
    df_base_vend = df_info_clientes[['Cod Clien', 'Cod Vend']].drop_duplicates('Cod Clien')
    df_base_vend.columns = ['CodClien', 'Base']
    
    df_final = df_structure.merge(pivot_table, on=['UsuarioEncer', 'CodClien'], how='left')
    df_final = df_final.merge(df_dt_ped, on='CodClien', how='left')
    df_final = df_final.merge(df_mes, on=['CodClien', 'UsuarioEncer'], how='left')
    df_final = df_final.merge(df_3m, on=['CodClien', 'UsuarioEncer'], how='left')
    df_final = df_final.merge(df_total, on=['CodClien', 'UsuarioEncer'], how='left')
    df_final = df_final.merge(df_last, on=['CodClien', 'UsuarioEncer'], how='left')
    df_final = df_final.merge(df_prox, on=['CodClien', 'UsuarioEncer'], how='left')
    df_final = df_final.merge(df_base_vend, on='CodClien', how='left')
    
    return df_final, pivot_table.columns

# --- Lógica de Automação de Agenda ---
def calcular_dias_uteis(data_inicial, dias_a_frente):
    datas = []
    dia_atual = data_inicial
    while len(datas) < dias_a_frente:
        if dia_atual.weekday() < 5: 
            datas.append(dia_atual)
        dia_atual += datetime.timedelta(days=1)
    return datas

def gerar_sugestao_agenda(
    df_base_agenda, 
    df_info_clientes_raw, 
    vendedor_agenda, 
    capacidade_diaria, 
    dias_projecao,
    filtros_avancados
):
    hoje = datetime.date.today()
    inicio_mes_atual = hoje.replace(day=1)
    ts_hoje = pd.Timestamp(hoje)
    ts_inicio_mes = pd.Timestamp(inicio_mes_atual)
    
    diagnostico = {
        "Total Carteira (Bruto)": 0,
        "Removidos: Agendamento (Mesmo Vendedor)": 0,
        "Removidos: Atrasados (Já Inclusos)": 0,
        "Removidos: Filtros (Mun/Seg/Etc)": 0,
        "Removidos: Já Atendidos (Filtro)": 0,
        "Total Final Backlog": 0,
        "Total Atrasados": 0
    }

    df_info_clientes = df_info_clientes_raw.drop_duplicates(subset=['Cod Clien']).copy()
    df_info_clientes['Cod Vend Limpo'] = df_info_clientes['Cod Vend'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    
    cols_str = ['Segmento', 'AreaVenda', 'SitCred', 'CondPgto', 'Municipio', 'Grupo', 'Colig']
    for c in cols_str:
        if c in df_info_clientes.columns:
            df_info_clientes[c] = df_info_clientes[c].fillna('')
            
    df_info_clientes['CondPgto_Visual'] = df_info_clientes['CondPgto'].apply(lambda x: 'Depósito Antecipado' if x == '' or pd.isna(x) else x)

    df_hist_vendedor = df_base_agenda[
        (df_base_agenda['UsuarioEncer'] == vendedor_agenda) &
        (df_base_agenda['Sit'] == 'EN')
    ].copy()
    
    if not df_hist_vendedor.empty:
        df_hist_vendedor = df_hist_vendedor.sort_values('DataEncer', ascending=False).drop_duplicates('CodClien')
        df_hist_vendedor = df_hist_vendedor[['CodClien', 'Motivo_Final', 'DataEncer', 'Obs']]
        df_hist_vendedor.columns = ['Cod Clien', 'Ult_Motivo_Vend', 'Dt_Ult_Motivo_Vend', 'Ult_Obs_Vend']
    else:
        df_hist_vendedor = pd.DataFrame(columns=['Cod Clien', 'Ult_Motivo_Vend', 'Dt_Ult_Motivo_Vend', 'Ult_Obs_Vend'])

    df_atrasados_raw = df_base_agenda[
        (df_base_agenda['UsuarioAgenda'] == vendedor_agenda) &
        (df_base_agenda['Sit'] == 'AB') &
        (df_base_agenda['DataAgenda'] < ts_hoje) 
    ].copy()
    
    df_atrasados_raw = df_atrasados_raw.sort_values('DataAgenda').drop_duplicates('CodClien')
    
    df_atrasados_full = df_atrasados_raw.merge(
        df_info_clientes, 
        left_on='CodClien', 
        right_on='Cod Clien', 
        how='left',
        suffixes=('', '_base')
    )
    df_atrasados_full = df_atrasados_full.merge(df_hist_vendedor, on='Cod Clien', how='left')
    lista_ids_atrasados = df_atrasados_full['CodClien'].unique()
    diagnostico["Total Atrasados"] = len(lista_ids_atrasados)
    
    vendedor_base_str = str(filtros_avancados.get('cod_vend_base', '')).strip().replace('.0', '')
    
    if vendedor_base_str:
        df_carteira = df_info_clientes[df_info_clientes['Cod Vend Limpo'] == vendedor_base_str].copy()
        diagnostico["Total Carteira (Bruto)"] = len(df_carteira)
    else:
        return [], pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), "Código do Vendedor na Base não informado."
    
    len_antes = len(df_carteira)

    filtro_nota_opcao = filtros_avancados.get('filtro_notas', 'Todos')
    if filtro_nota_opcao != 'Todos':
        notas_dict = carregar_notas()
        ids_com_nota = set(notas_dict.keys())
        df_carteira['Cod_Str'] = df_carteira['Cod Clien'].astype(str)
        if filtro_nota_opcao == 'Com Notas':
            df_carteira = df_carteira[df_carteira['Cod_Str'].isin(ids_com_nota)]
        elif filtro_nota_opcao == 'Sem Notas':
            df_carteira = df_carteira[~df_carteira['Cod_Str'].isin(ids_com_nota)]

    if filtros_avancados.get('apenas_com_prazo'):
        df_carteira = df_carteira[df_carteira['CondPgto_Visual'] != 'Depósito Antecipado']

    if not filtros_avancados.get('incluir_grupo_colig'):
        df_carteira = df_carteira[(df_carteira['Grupo'] == '') & (df_carteira['Colig'] == '')]

    if filtros_avancados.get('municipios'):
        df_carteira = df_carteira[df_carteira['Municipio'].isin(filtros_avancados['municipios'])]
        
    if filtros_avancados.get('segmentos'):
        df_carteira = df_carteira[df_carteira['Segmento'].isin(filtros_avancados['segmentos'])]

    if filtros_avancados.get('areas'):
        df_carteira = df_carteira[df_carteira['AreaVenda'].isin(filtros_avancados['areas'])]

    if filtros_avancados.get('sit_cred'):
        df_carteira = df_carteira[df_carteira['SitCred'].isin(filtros_avancados['sit_cred'])]

    if filtros_avancados.get('cond_pgto'):
        df_carteira = df_carteira[df_carteira['CondPgto_Visual'].isin(filtros_avancados['cond_pgto'])]

    val_limite = filtros_avancados.get('limite_min', 0.0)
    if val_limite > 0:
        df_carteira = df_carteira[df_carteira['LimiteTotal'] >= val_limite]

    val_dias_min = filtros_avancados.get('dias_sem_compra_min', 0)
    if val_dias_min > 0:
        df_carteira = df_carteira[df_carteira['DiasSemCompra'] >= val_dias_min]

    if filtros_avancados.get('apenas_adimplentes'):
        df_carteira = df_carteira[(df_carteira['Inad-3dd'] <= 0) | (df_carteira['Inad-3dd'].isna())]
    
    len_depois_filtros = len(df_carteira)
    diagnostico["Removidos: Filtros (Mun/Seg/Etc)"] = len_antes - len_depois_filtros

    clientes_com_futuro = df_base_agenda[
        (df_base_agenda['Sit'] == 'AB') &
        (df_base_agenda['DataAgenda'] >= ts_hoje) &
        (df_base_agenda['UsuarioAgenda'] == vendedor_agenda)
    ]['CodClien'].unique()
    
    df_backlog = df_carteira[~df_carteira['Cod Clien'].isin(clientes_com_futuro)].copy()
    
    len_sem_futuro = len(df_backlog)
    diagnostico["Removidos: Agendamento (Mesmo Vendedor)"] = len_depois_filtros - len_sem_futuro
    
    df_backlog = df_backlog[~df_backlog['Cod Clien'].isin(lista_ids_atrasados)].copy()
    
    len_final_backlog = len(df_backlog)
    diagnostico["Removidos: Atrasados (Já Inclusos)"] = len_sem_futuro - len_final_backlog
    diagnostico["Total Final Backlog"] = len_final_backlog
    
    df_backlog['ComprouEsteMes'] = df_backlog['DtUltPed'].apply(
        lambda x: x.year == hoje.year and x.month == hoje.month if pd.notnull(x) else False
    )
    
    clientes_ja_atendidos = df_base_agenda[
        (df_base_agenda['UsuarioEncer'] == vendedor_agenda) &
        (df_base_agenda['Sit'] == 'EN')
    ]['CodClien'].unique()
    df_backlog['JaAtendido'] = df_backlog['Cod Clien'].isin(clientes_ja_atendidos)
    
    # --- FILTRO: Já Atendidos ---
    len_antes_ja_atendidos = len(df_backlog)
    if not filtros_avancados.get('incluir_ja_atendidos', True):
        df_backlog = df_backlog[~df_backlog['JaAtendido']].copy()
    
    diagnostico["Removidos: Já Atendidos (Filtro)"] = len_antes_ja_atendidos - len(df_backlog)
    diagnostico["Total Final Backlog"] = len(df_backlog)
    # -----------------------------
    
    contatos_mes = df_base_agenda[
        (df_base_agenda['UsuarioEncer'] == vendedor_agenda) &
        (df_base_agenda['DataEncer'] >= ts_inicio_mes)
    ].groupby('CodClien').size().reset_index(name='QtdContatosMes')
    
    df_backlog = df_backlog.merge(contatos_mes, left_on='Cod Clien', right_on='CodClien', how='left')
    df_backlog['QtdContatosMes'] = df_backlog['QtdContatosMes'].fillna(0)
    df_backlog = df_backlog.merge(df_hist_vendedor, on='Cod Clien', how='left')
    
    df_backlog = df_backlog.sort_values(
        by=['ComprouEsteMes', 'JaAtendido', 'QtdContatosMes', 'DiasSemCompra'],
        ascending=[True, True, True, False]
    )
    
    fila_unificada = []
    
    def extrair_dados(row, motivo_prioridade):
        def get_val(col, default='N/D'):
            val = row.get(col)
            return val if pd.notnull(val) else default
        
        cod_final = row.get('Cod Clien')
        if pd.isna(cod_final):
            cod_final = row.get('CodClien')
        
        return {
            'Cod Clien': cod_final,
            'Cliente': get_val('Cliente', 'Cliente S/ Cadastro'),
            'Cod Vend': get_val('Cod Vend', ''),
            'Segmento': get_val('Segmento', ''),
            'Municipio': get_val('Municipio', ''),
            'AreaVenda': get_val('AreaVenda', ''),
            'SitCred': get_val('SitCred', ''),
            'LimiteTotal': row.get('LimiteTotal', 0),
            'DtUltPed': row.get('DtUltPed', pd.NaT),
            'DiasSemCompra': row.get('DiasSemCompra', 0),
            'Inad-3dd': row.get('Inad-3dd', 0),
            'CondPgto_Visual': 'Depósito Antecipado' if pd.isna(row.get('CondPgto')) or row.get('CondPgto') == '' else row.get('CondPgto'),
            'MotivoPrioridade': motivo_prioridade,
            'Ult_Motivo_Vend': get_val('Ult_Motivo_Vend', ''),
            'Dt_Ult_Motivo_Vend': row.get('Dt_Ult_Motivo_Vend', pd.NaT),
            'Ult_Obs_Vend': get_val('Ult_Obs_Vend', '')
        }

    for _, row in df_atrasados_full.iterrows():
        fila_unificada.append(extrair_dados(row, '⚠️ 1. Atrasado'))
        
    for _, row in df_backlog.iterrows():
        if not row['ComprouEsteMes']:
            label = "2. Sem compra no mês"
        elif not row['JaAtendido']:
            label = "3. Nunca atendidos"
        elif row['QtdContatosMes'] < 1:
            label = "4. Frequência baixa"
        else:
            label = "5. Recorrência (Já comprou)"
        fila_unificada.append(extrair_dados(row, label))
    
    datas_projecao_uteis = calcular_dias_uteis(hoje, dias_projecao)
    
    resultado_final = []
    metricas_distribuicao = [] 
    
    idx_fila = 0
    total_fila = len(fila_unificada)
    
    for data in datas_projecao_uteis:
        ts_data = pd.Timestamp(data)
        ocupados_dia = df_base_agenda[
            (df_base_agenda['UsuarioAgenda'] == vendedor_agenda) &
            (df_base_agenda['Sit'] == 'AB') &
            (df_base_agenda['DataAgenda'].dt.date == data)
        ].shape[0]
        
        vagas = capacidade_diaria - ocupados_dia
        if vagas < 0: vagas = 0
        
        sugeridos_neste_dia = 0
        for _ in range(vagas):
            if idx_fila < total_fila:
                item = fila_unificada[idx_fila].copy()
                item['Data Sugerida'] = data
                if item['DiasSemCompra'] == 9999:
                    item['DiasSemCompra'] = None
                resultado_final.append(item)
                idx_fila += 1
                sugeridos_neste_dia += 1
            else:
                break
        
        metricas_distribuicao.append({
            'Data': data,
            'Meta': capacidade_diaria,
            'Ocupado': ocupados_dia,
            'Sugerido': sugeridos_neste_dia
        })
            
    df_resultado = pd.DataFrame(resultado_final)
    df_diagnostico = pd.DataFrame([diagnostico])
    df_metricas = pd.DataFrame(metricas_distribuicao)
    
    if df_resultado.empty:
        return datas_projecao_uteis, df_resultado, df_metricas, df_diagnostico, "Nenhum cliente disponível para agendar."
    return datas_projecao_uteis, df_resultado, df_metricas, df_diagnostico, "Sucesso"


# --- CARREGAMENTO DE DADOS (ETL) ---

@st.cache_data
def carregar_dados_agenda():
    hoje = datetime.date.today()
    # Aumentado para 6 meses para garantir histórico no gráfico
    data_inicio = (hoje - relativedelta(months=6)).replace(day=1)
    data_fim = (hoje + relativedelta(months=3)).replace(day=1) - relativedelta(days=1)
    
    data_inicio_str = data_inicio.strftime('%Y-%m-%d')
    data_fim_str = data_fim.strftime('%Y-%m-%d')
    
    driver = os.getenv("DB_DRIVER", "ODBC Driver 17 for SQL Server")
    server = os.getenv("DB_SERVER")
    database = os.getenv("DB_DATABASE")
    uid = os.getenv("DB_UID")
    pwd = os.getenv("DB_PWD")

    connection_string = (
        f"DRIVER={{{driver}}};SERVER={server};DATABASE={database};"
        f"UID={uid};PWD={{{pwd}}};TrustServerCertificate=yes;"
    )
    
    sql_call = "{CALL sp_createdbyMGR_ResultadoContatoVDI(?, ?)}"
    parametros = (data_inicio_str, data_fim_str)
    
    try:
        with pyodbc.connect(connection_string) as conexao:
            df = pd.read_sql_query(sql_call, conexao, params=parametros)
    except Exception as e:
        st.error(f"ERRO ao conectar na AGENDA: {e}")
        return pd.DataFrame()

    try:
        df['DataEncer'] = pd.to_datetime(df['DataEncer'], errors='coerce')
        df['DataAgenda'] = pd.to_datetime(df['DataAgenda'], errors='coerce')
        df_contatos = df.dropna(subset=['DataEncer', 'DataAgenda'], how='all').copy()
        df_contatos['Motivo'] = df_contatos['Motivo'].fillna('')
        df_contatos['Motivo_Final'] = df_contatos.apply(
            lambda row: row['Resultado'] if row['Motivo'] == '' else row['Motivo'], axis=1
        )
        replacements = {'COMPROU CONCORRENT': 'COMPROU CONCORRENTE', 'COTACAO': 'COTAÇÃO'}
        df_contatos['Motivo_Final'] = df_contatos['Motivo_Final'].replace(replacements)
        df_contatos['CodClien'] = pd.to_numeric(df_contatos['CodClien'], errors='coerce').fillna(0).astype(int)
        df_contatos['CodClien_str'] = df_contatos['CodClien'].astype(str)
        df_contatos['UsuarioAgenda'] = df_contatos['UsuarioAgenda'].astype(str).str.strip()
        df_contatos['UsuarioEncer'] = df_contatos['UsuarioEncer'].astype(str).str.strip()
        
        return df_contatos
    except Exception as e_etl:
        st.error(f"ERRO na limpeza Agenda: {e_etl}")
        return pd.DataFrame()

@st.cache_data
def carregar_dados_base_cliente():
    driver = os.getenv("DB_DRIVER", "ODBC Driver 17 for SQL Server")
    server = os.getenv("DB_SERVER")
    database = os.getenv("DB_DATABASE")
    uid = "excel_log" 
    pwd = os.getenv("DB_PWD") 

    connection_string = (
        f"DRIVER={{{driver}}};SERVER={server};DATABASE={database};"
        f"UID={uid};PWD={{{pwd}}};TrustServerCertificate=yes;"
    )
    
    sql_script = """
    SELECT      ge.descricao            AS [Gerencia]       ,
                eq.descricao            AS [Equipe]         ,
                ve.cd_vend              AS [Cod Vend]       ,
                ve.nome                 AS [Vendedor]       ,
                cl.cd_clien             AS [Cod Clien]      ,
                cl.nome                 AS [Cliente]        ,
                cl.e_mail               AS [E-mail]         ,
                ra.descricao            AS [Segmento]       ,
                ec.endereco             AS [Endereco]       ,
                ec.bairro               AS [Bairro]         ,
                ec.municipio            AS [Municipio]      ,
                ec.estado               AS [UF]             ,
                (SELECT descricao FROM coligacao WHERE cd_coligacao = cl.cd_coligacao) AS [Colig],
                (SELECT descricao FROM grupocli WHERE cd_grupocli = cl.cd_grupocli) AS [Grupo],
                cl.tp_pes               AS [TipoPessoa]     ,
                cl.cgc_cpf              AS [CNPJ/CPF]       ,
                a.descricao             AS [AreaVenda]      ,
                s.descricao             AS [SitCred]        ,
                cl.vl_lim_cred          AS [LimiteTotal]    ,
                (SELECT nu_ped FROM ped_vda WHERE pedvdaid = ped.pedvdaid) AS [UltPed],
                (SELECT cd_vend FROM ped_vda WHERE pedvdaid = ped.pedvdaid) AS VendUltPed,
                (SELECT dt_cad FROM ped_vda WHERE pedvdaid = ped.pedvdaid) AS DtUltPed,
                (SELECT sum(valor-vl_pago) FROM titrec t WHERE t.cd_clien = cl.cd_clien AND t.situacao in ('AB','PL') AND dt_venc <= getdate()-3) AS [Inad-3dd],
                (SELECT top 1 (P.descricao) FROM Promocao P JOIN ClienPrm CP ON CP.seq_prom = P.seq_prom WHERE CP.cd_clien = cl.cd_clien AND P.descricao LIKE '%Dias%') AS [CondPgto],
                UltContato.dt_encer     AS [UltContato]         ,
                UltContato.cd_usr_enc   AS [RespUltContato]     ,
                UltContato.resultado    AS [Resultado]          ,
                UltContato.motivo       AS [Motivo]             ,
                UltContato.obs          AS [OBS]                ,
                cl.dt_cad               AS [DataCadastro]
    FROM        cliente cl
    JOIN        end_cli ec on cl.cd_clien = ec.cd_clien
    JOIN        ram_ativ ra on cl.ram_ativ = ra.ram_ativ
    JOIN        vend_cli vc ON cl.cd_clien = vc.cd_clien
    JOIN        vendedor ve ON vc.cd_vend = ve.cd_vend
    JOIN        equipe eq on ve.cd_emp = eq.cd_emp AND ve.cd_equipe = eq.cd_equipe
    JOIN        gerencia ge on eq.cd_emp = ge.cd_emp AND eq.cd_gerencia = ge.cd_gerencia
    JOIN        st_cred s on cl.st_cred = s.st_cred
    LEFT JOIN   area a on cl.cd_area = a.cd_area
    LEFT JOIN   (SELECT cd_clien, max(pedvdaid) pedvdaid FROM ped_vda pv WHERE dt_cad >= '20230101' AND pv.situacao = 'AB' AND pv.tp_ped in (SELECT tp_ped FROM tp_ped WHERE estat_com=1 AND pv.tp_ped not in ('VA','BD') AND pv.cd_vend not in ('FELIPEB','RODRIGO')) GROUP BY cd_clien) AS ped ON cl.cd_clien = ped.cd_clien
    LEFT JOIN   (SELECT e.cd_clien, e.dt_encer, e.cd_usr_enc, r.descricao AS resultado, m.descricao AS motivo, l.texto AS obs FROM evento_tmkt e JOIN (SELECT cd_clien, max(seq_evento) seq_evento FROM evento_tmkt WHERE situacao = 'EN' GROUP BY cd_clien) AS ev2 ON e.cd_clien = ev2.cd_clien AND e.seq_evento = ev2.seq_evento JOIN res_tmkt r ON e.cd_resultado = r.cd_resultado LEFT JOIN mot_res m ON e.cd_resultado = m.cd_resultado AND e.cd_motivo = m.cd_motivo LEFT JOIN lin_txt l ON e.comentario = l.cd_texto AND l.num_lin = 1) UltContato ON cl.cd_clien = UltContato.cd_clien
    WHERE       cl.ativo=1 AND ec.tp_end='EN'
    """

    try:
        with pyodbc.connect(connection_string) as conexao:
            df = pd.read_sql_query(sql_script, conexao)
    except Exception as e:
        st.error(f"ERRO ao conectar na BASE: {e}")
        return pd.DataFrame()

    df['Cod Clien'] = pd.to_numeric(df['Cod Clien'], errors='coerce').fillna(0).astype(int)
    df['DtUltPed'] = pd.to_datetime(df['DtUltPed'], errors='coerce')
    hoje_dt = pd.to_datetime(datetime.date.today())
    df['DiasSemCompra'] = (hoje_dt - df['DtUltPed']).dt.days
    df['DiasSemCompra'] = df['DiasSemCompra'].fillna(9999).astype(int) 
    df['Inad-3dd'] = pd.to_numeric(df['Inad-3dd'], errors='coerce').fillna(0)
    df['LimiteTotal'] = pd.to_numeric(df['LimiteTotal'], errors='coerce').fillna(0)
    df['AreaVenda'] = df['AreaVenda'].fillna('Não Definida')
    df['Segmento'] = df['Segmento'].fillna('Não Definido')
    
    df['Cod Vend'] = pd.to_numeric(df['Cod Vend'], errors='coerce').fillna(0).astype(int).astype(str)
    df['Cod Vend'] = df['Cod Vend'].replace('0', '')
    
    df['VendUltPed'] = pd.to_numeric(df['VendUltPed'], errors='coerce').fillna(0).astype(int).astype(str)
    df['VendUltPed'] = df['VendUltPed'].replace('0', '')

    return df

# --- Inicialização ---
if 'cliente_index' not in st.session_state:
    st.session_state.cliente_index = 0

with st.spinner("Conectando ao banco e carregando dados..."):
    df_base = carregar_dados_agenda()
    df_info_cliente = carregar_dados_base_cliente()

if df_base.empty or df_info_cliente.empty:
    st.error("Aplicação parada. Verifique os erros de conexão ou query.")
    st.stop()

# --- MERGE GLOBAL ---
df_enrich = df_info_cliente[['Cod Clien', 'DiasSemCompra', 'Cod Vend']].drop_duplicates(subset=['Cod Clien'])
df_base = df_base.merge(
    df_enrich,
    left_on='CodClien',
    right_on='Cod Clien',
    how='left'
)
df_base['DiasSemCompra'] = df_base['DiasSemCompra'].fillna(0).astype(int)
df_base['Cod Vend'] = df_base['Cod Vend'].fillna('') 

st.toast("Dados carregados com sucesso!", icon="✅")

# --- Barra Lateral (Sidebar) com Filtros ---
st.sidebar.header("Filtros do Dashboard")

if st.sidebar.button("Atualizar Dados"):
    carregar_dados_agenda.clear()
    carregar_dados_base_cliente.clear()
    st.session_state.cliente_index = 0 
    st.rerun()

st.sidebar.markdown("---")

lista_vendedores = sorted(df_base['UsuarioAgenda'].dropna().unique())
vendedores_selecionados = st.sidebar.multiselect(
    "Vendedor (UsuarioEncer/Agenda):", 
    options=lista_vendedores,
    default=[]
)
if not vendedores_selecionados:
    vendedores_selecionados = lista_vendedores

st.sidebar.markdown("---")
lista_vendedores_base_input = st.sidebar.text_area(
    "Lista Cód. Vendedores Base (Enter ou vírgula):",
    help="Filtra pelo código do vendedor de cadastro (Base)."
)

st.sidebar.markdown("---")
lista_motivos = sorted(df_base['Motivo_Final'].dropna().unique())
motivos_selecionados = st.sidebar.multiselect(
    "Filtrar por Motivo/Resultado:",
    options=lista_motivos,
    default=[],
    help="Selecione os motivos. Se vazio, mostra todos."
)

st.sidebar.markdown("---")
st.sidebar.subheader("Filtro de Cliente")

lista_clientes_input = st.sidebar.text_area(
    "Lista Cód. Clientes (Enter ou vírgula):",
    help="Cole uma lista de códigos de clientes."
)

filtro_nome_cliente = st.sidebar.text_input("Buscar por Nome do Cliente:")

filtro_obs_texto = st.sidebar.text_input("Filtrar Obs (Contém):", help="Busca texto dentro das observações")

filtro_notas = st.sidebar.radio(
    "Filtro de Notas:", 
    ["Todos", "Com Notas", "Sem Notas"], 
    horizontal=True
)

filtro_dias_sem_compra = st.sidebar.number_input(
    "Dias sem Compra (Mínimo):", 
    min_value=0, 
    value=0, 
    step=10,
    help="Filtra TODO o dashboard para clientes que não compram há X dias."
)

st.sidebar.markdown("---")
st.sidebar.subheader("Filtro de Agendamento")
filtro_tipo_agendamento = st.sidebar.radio(
    "Status Agendamento:",
    ("Todos", "Com Agendamento", "Sem Agendamento", "Data do Agendamento")
)

filtro_data_agend_inicio = None
filtro_data_agend_fim = None

if filtro_tipo_agendamento == "Data do Agendamento":
    col_ag1, col_ag2 = st.sidebar.columns(2)
    filtro_data_agend_inicio = col_ag1.date_input("De:", value=datetime.date.today())
    filtro_data_agend_fim = col_ag2.date_input("Até:", value=datetime.date.today() + datetime.timedelta(days=7))

st.sidebar.markdown("---")
st.sidebar.subheader("Filtro de Período (Último Contato)")
st.sidebar.info("Filtra clientes cujo último encerramento ocorreu neste intervalo.")

data_min_agenda = df_base['DataAgenda'].dropna().min().date()
data_min_encer = df_base['DataEncer'].dropna().min().date()
data_min_original = min(data_min_agenda, data_min_encer) if not pd.isna(data_min_agenda) else datetime.date.today()
data_max_agenda = df_base['DataAgenda'].dropna().max().date()
data_max_encer = df_base['DataEncer'].dropna().max().date()
data_max_original = max(data_max_agenda, data_max_encer) if not pd.isna(data_max_agenda) else datetime.date.today()

filtro_data_inicio = st.sidebar.date_input("Data Inicial:", value=data_min_original)
filtro_data_fim = st.sidebar.date_input("Data Final:", value=data_max_original)

# --- APLICAÇÃO DOS FILTROS ---

# 1. Filtros Comuns
df_common_filters = df_base.copy()

if filtro_obs_texto:
    df_common_filters = df_common_filters[df_common_filters['Obs'].astype(str).str.contains(filtro_obs_texto, case=False, na=False)]

if filtro_notas != 'Todos':
    notas_carregadas_filtro = carregar_notas()
    ids_com_nota = set(notas_carregadas_filtro.keys())
    
    if filtro_notas == 'Com Notas':
        df_common_filters = df_common_filters[df_common_filters['CodClien_str'].isin(ids_com_nota)]
    elif filtro_notas == 'Sem Notas':
        df_common_filters = df_common_filters[~df_common_filters['CodClien_str'].isin(ids_com_nota)]

if vendedores_selecionados:
    mask_vend = (df_common_filters['UsuarioEncer'].isin(vendedores_selecionados)) | \
                (df_common_filters['UsuarioAgenda'].isin(vendedores_selecionados))
    df_common_filters = df_common_filters[mask_vend]

if motivos_selecionados:
    df_common_filters = df_common_filters[df_common_filters['Motivo_Final'].isin(motivos_selecionados)]
if filtro_dias_sem_compra > 0:
    df_common_filters = df_common_filters[df_common_filters['DiasSemCompra'] >= filtro_dias_sem_compra]
if filtro_nome_cliente:
    df_common_filters = df_common_filters[df_common_filters['Cliente'].str.contains(filtro_nome_cliente, case=False, na=False)]
if lista_clientes_input:
    codigos_raw = lista_clientes_input.replace(',', '\n').split('\n')
    codigos_limpos = [c.strip() for c in codigos_raw if c.strip()]
    if codigos_limpos:
        df_common_filters = df_common_filters[df_common_filters['CodClien_str'].isin(codigos_limpos)]
if lista_vendedores_base_input:
    vends_raw = lista_vendedores_base_input.replace(',', '\n').split('\n')
    vends_limpos = [v.strip() for v in vends_raw if v.strip()]
    if vends_limpos:
        df_common_filters = df_common_filters[df_common_filters['Cod Vend'].isin(vends_limpos)]

# 2. FILTRO DE AGENDAMENTO
df_schedules_full = df_base[df_base['Sit']=='AB'].groupby(['UsuarioAgenda', 'CodClien'])['DataAgenda'].min().reset_index()
df_schedules_full.columns = ['UsuarioEncer', 'CodClien', 'DataProxAgend']

df_pairs = df_common_filters[['UsuarioEncer', 'CodClien']].drop_duplicates()
df_pairs = df_pairs.merge(df_schedules_full, on=['UsuarioEncer', 'CodClien'], how='left')

df_sched_filtered = df_common_filters[df_common_filters['Sit'] == 'AB'].copy()

if filtro_tipo_agendamento == "Sem Agendamento":
    df_pairs = df_pairs[df_pairs['DataProxAgend'].isna()]
    df_sched_filtered = df_sched_filtered[0:0] 

elif filtro_tipo_agendamento == "Com Agendamento":
    df_pairs = df_pairs[df_pairs['DataProxAgend'].notna()]

elif filtro_tipo_agendamento == "Data do Agendamento" and filtro_data_agend_inicio and filtro_data_agend_fim:
    ts_inicio_ag = pd.Timestamp(filtro_data_agend_inicio)
    ts_fim_ag = pd.Timestamp(filtro_data_agend_fim) + pd.Timedelta(hours=23, minutes=59, seconds=59)
    
    df_sched_filtered = df_sched_filtered[
        (df_sched_filtered['DataAgenda'] >= ts_inicio_ag) &
        (df_sched_filtered['DataAgenda'] <= ts_fim_ag)
    ]
    
    df_pairs = df_pairs[
        (df_pairs['DataProxAgend'] >= ts_inicio_ag) &
        (df_pairs['DataProxAgend'] <= ts_fim_ag)
    ]

df_common_agend_filtered = df_common_filters.merge(df_pairs[['UsuarioEncer', 'CodClien']], on=['UsuarioEncer', 'CodClien'], how='inner')

# 3. DataFrames Visuais (Histórico)
ts_inicio_geral = pd.Timestamp(filtro_data_inicio)
ts_fim_geral = pd.Timestamp(filtro_data_fim) + pd.Timedelta(hours=23, minutes=59, seconds=59)

# BLINDAGEM DE DATA
df_filtrado_com_data = df_common_agend_filtered[
    (df_common_agend_filtered['Sit'] == 'EN') &
    (df_common_agend_filtered['DataEncer'] >= ts_inicio_geral) &
    (df_common_agend_filtered['DataEncer'] <= ts_fim_geral)
].copy()

df_encerrados_global = df_base[df_base['Sit'] == 'EN'].copy()
df_encerrados_global_sorted = df_encerrados_global.sort_values(by='DataEncer', ascending=False)
df_ultimos_contatos_usuario = df_encerrados_global_sorted.drop_duplicates(subset=['CodClien', 'UsuarioEncer'])
df_ultimos_contatos_usuario = df_ultimos_contatos_usuario[[
    'CodClien', 'UsuarioEncer', 'DataEncer', 'Motivo_Final', 'Obs'
]].copy() 
df_ultimos_contatos_usuario = df_ultimos_contatos_usuario.rename(columns={
    'DataEncer': 'Data Ult Contato',
    'Motivo_Final': 'Ultimo Motivo',
    'Obs': 'Obs Ult Contato',
    'UsuarioEncer': 'UsuarioEncer_JoinKey'
})
df_ultimos_contatos_usuario['UsuarioEncer_JoinKey'] = df_ultimos_contatos_usuario['UsuarioEncer_JoinKey'].astype(str)


# --- Página Principal ---
st.title("Dashboard de Análise de Contatos")

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Dashboard Principal", 
    "👨‍💼 Detalhe por Cliente", 
    "🗓️ Próximos Agendamentos",
    "📑 Tabela Resumo",
    "🤖 Automação de Agenda"
])

# --- ABA 1 (DASHBOARD) ---
with tab1:
    df_analise = df_filtrado_com_data
    
    hoje = datetime.date.today()
    ts_inicio_mes_vigente = pd.Timestamp(hoje.replace(day=1))
    
    df_abertos_kpi = df_base[
        (df_base['Sit'] == 'AB') &
        (df_base['UsuarioAgenda'].isin(vendedores_selecionados))
    ]
    if filtro_dias_sem_compra > 0:
        df_abertos_kpi = df_abertos_kpi[df_abertos_kpi['DiasSemCompra'] >= filtro_dias_sem_compra]
        
    agendamentos_aberto_total = df_abertos_kpi.shape[0]
    
    if not df_analise.empty:
        contatos_total_filtrado = df_analise.shape[0]
        contatos_mes_vigente_filtrado = df_analise[df_analise['DataEncer'] >= ts_inicio_mes_vigente].shape[0]
        clientes_unicos_filtrados = df_analise['CodClien'].nunique()
        
        ts_inicio_3_meses = pd.Timestamp((hoje.replace(day=1) - pd.DateOffset(months=2)).date())
        contatos_3_meses_filtrado = df_analise[df_analise['DataEncer'] >= ts_inicio_3_meses].shape[0]
    else:
        contatos_total_filtrado = 0
        contatos_mes_vigente_filtrado = 0
        contatos_3_meses_filtrado = 0
        clientes_unicos_filtrados = 0
        
    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Contatos (Filtrado)", f"{contatos_total_filtrado:,}".replace(",", "."))
    col2.metric("Contatos Mês Vigente", f"{contatos_mes_vigente_filtrado:,}".replace(",", "."))
    col3.metric("Contatos 3 Meses", f"{contatos_3_meses_filtrado:,}".replace(",", "."))
    col4.metric("Clientes Únicos", f"{clientes_unicos_filtrados:,}".replace(",", "."))
    col5.metric("Agend. em Aberto (Total)", f"{agendamentos_aberto_total:,}".replace(",", "."))
    
    st.markdown("---")
    
    df_base_resumo = df_base.copy()
    if vendedores_selecionados:
        df_base_resumo = df_base_resumo[
            (df_base_resumo['UsuarioEncer'].isin(vendedores_selecionados)) | 
            (df_base_resumo['UsuarioAgenda'].isin(vendedores_selecionados))
        ]
        
    df_encerrados = df_base_resumo[df_base_resumo['Sit'] == 'EN']
    df_abertos = df_base_resumo[df_base_resumo['Sit'] == 'AB']
    
    if filtro_dias_sem_compra > 0:
        df_encerrados = df_encerrados[df_encerrados['DiasSemCompra'] >= filtro_dias_sem_compra]
        df_abertos = df_abertos[df_abertos['DiasSemCompra'] >= filtro_dias_sem_compra]
    
    df_enc_mes = df_encerrados[df_encerrados['DataEncer'] >= ts_inicio_mes_vigente]
    contatos_mes = df_enc_mes.groupby('UsuarioEncer').size().rename('Contatos no Mês')
    
    ts_hoje = pd.Timestamp(hoje)
    df_enc_hoje = df_encerrados[df_encerrados['DataEncer'].dt.date == hoje]
    encerrados_hoje = df_enc_hoje.groupby('UsuarioEncer').size().rename('Encerrados Hoje')
    
    df_resumo_enc = pd.concat([contatos_mes, encerrados_hoje], axis=1)
    
    mes_atual_periodo = pd.Period(hoje, 'M')
    df_aber_mes = df_abertos[(df_abertos['DataAgenda'].dt.to_period('M') == mes_atual_periodo)]
    abertos_mes = df_aber_mes.groupby('UsuarioAgenda').size().rename('Agend. em Aberto (Mês)')
    
    df_aber_hoje = df_abertos[df_abertos['DataAgenda'].dt.date == hoje]
    aberto_hoje = df_aber_hoje.groupby('UsuarioAgenda').size().rename('Agend. para Hoje (Aberto)')
    
    df_aber_atrasados = df_abertos[df_abertos['DataAgenda'].dt.date < hoje]
    atrasados_total = df_aber_atrasados.groupby('UsuarioAgenda').size().rename('Agend. Atrasados (Total)')
    
    df_resumo_ab = pd.concat([abertos_mes, aberto_hoje, atrasados_total], axis=1)
    
    df_resumo_vendedores = pd.concat([df_resumo_enc, df_resumo_ab], axis=1).fillna(0).astype(int)
    df_resumo_vendedores.index.name = "Vendedor"
    df_resumo_vendedores['Total Workload (Mês)'] = df_resumo_vendedores['Contatos no Mês'] + df_resumo_vendedores['Agend. em Aberto (Mês)']
    df_resumo_vendedores['% Eficácia (Mês)'] = (
        df_resumo_vendedores['Contatos no Mês'].div(df_resumo_vendedores['Total Workload (Mês)'])
    ).fillna(0).map('{:.1%}'.format)
    
    colunas_ordenadas = [
        'Total Workload (Mês)', 'Contatos no Mês', 'Agend. em Aberto (Mês)',
        '% Eficácia (Mês)', 'Encerrados Hoje', 'Agend. para Hoje (Aberto)',
        'Agend. Atrasados (Total)'
    ]
    df_resumo_vendedores = df_resumo_vendedores[colunas_ordenadas].sort_values(by='Total Workload (Mês)', ascending=False)
    
    st.subheader("Resumo de Vendedores (Mês Vigente)")
    st.dataframe(df_resumo_vendedores, use_container_width=True)
    
    st.markdown("---")
    
    if df_analise.empty:
        st.warning("Nenhum dado para exibir nos gráficos com os filtros atuais.")
    else:
        col_graf1, col_graf2 = st.columns(2)
        template_graficos = "plotly_dark"
        with col_graf1:
            st.subheader("Top 15 Motivos de Contato")
            df_motivos = df_analise['Motivo_Final'].value_counts().head(15).reset_index()
            df_motivos.columns = ['Motivo', 'Contagem']
            fig_motivos = px.bar(
                df_motivos, x='Contagem', y='Motivo', orientation='h', text='Contagem',
                template=template_graficos
            )
            fig_motivos.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig_motivos, use_container_width=True)
        with col_graf2:
            st.subheader("Resultados por Vendedor")
            df_vendedor_res = df_analise.groupby('UsuarioEncer').size().reset_index(name='Contagem')
            fig_vendedores = px.bar(
                df_vendedor_res, x='UsuarioEncer', y='Contagem', text='Contagem',
                template=template_graficos
            )
            fig_vendedores.update_xaxes(tickangle=-45)
            st.plotly_chart(fig_vendedores, use_container_width=True)

        # --- NOVO GRÁFICO: HEATMAP (MAPA DE CALOR) - COM SEPARADOR DE SEMANAS ---
        st.markdown("---")
        st.subheader("Mapa de Calor: Intensidade de Atendimentos (Diário)")
        st.info("Visualização diária. Linhas verticais indicam o fim da semana (Sexta-feira).")

        df_timeline = df_analise.copy()
        df_timeline['Dia'] = df_timeline['DataEncer'].dt.date
        df_heatmap_data = df_timeline.groupby(['Dia', 'UsuarioEncer']).size().reset_index(name='Quantidade')
        df_pivot_heatmap = df_heatmap_data.pivot(index='UsuarioEncer', columns='Dia', values='Quantidade').fillna(0)

        fig_heatmap = px.imshow(
            df_pivot_heatmap,
            labels=dict(x="Data", y="Vendedor", color="Qtd."),
            text_auto=True,
            aspect="auto",
            color_continuous_scale='Blues',
            title="Distribuição de Volume Diário por Vendedor",
            template=template_graficos
        )

        linhas_verticais = []
        for i, data_col in enumerate(df_pivot_heatmap.columns):
            if pd.to_datetime(data_col).dayofweek == 4:
                linhas_verticais.append(dict(
                    type="line",
                    x0=i + 0.5,  
                    x1=i + 0.5,
                    y0=0,        
                    y1=1,        
                    xref="x",    
                    yref="paper",
                    line=dict(color="#666666", width=1) 
                ))

        fig_heatmap.update_layout(shapes=linhas_verticais)

        fig_heatmap.update_layout(
            coloraxis_colorbar=dict(title="Qtd."),
            xaxis=dict(type='category', title="Data"),
            margin=dict(r=20, l=20, b=20, t=40)
        )
        fig_heatmap.update_xaxes(tickangle=-45)

        st.plotly_chart(fig_heatmap, use_container_width=True)

# --- ABA 2 ---
with tab2:
    st.subheader("Análise Detalhada de Cliente")
    
    clientes_filtrados_agenda = df_filtrado_com_data.drop_duplicates(subset=['CodClien'])
    
    clientes_com_info = clientes_filtrados_agenda.merge(
        df_info_cliente, 
        left_on='CodClien',  
        right_on='Cod Clien', 
        how='left',
        suffixes=('', '_info')
    )
    
    clientes_navegacao = clientes_com_info.drop_duplicates(subset=['CodClien']).sort_values(by='Cliente')
    
    lista_codigos_cliente = clientes_navegacao['CodClien'].tolist() 

    if not lista_codigos_cliente:
        st.warning("Nenhum cliente encontrado para a combinação de filtros selecionada (neste período).")
    else:
        total_clientes = len(lista_codigos_cliente)
        
        if st.session_state.cliente_index >= total_clientes:
            st.session_state.cliente_index = 0
            
        col_nav1, col_nav2, col_nav3 = st.columns([2, 8, 2])
        
        if col_nav1.button("⬅️ Cliente Anterior", use_container_width=True):
            if st.session_state.cliente_index > 0:
                st.session_state.cliente_index -= 1
            else:
                st.session_state.cliente_index = total_clientes - 1
            st.rerun()

        if col_nav3.button("Cliente Próximo ➡️", use_container_width=True):
            if st.session_state.cliente_index < total_clientes - 1:
                st.session_state.cliente_index += 1
            else:
                st.session_state.cliente_index = 0
            st.rerun()

        current_cod_clien = lista_codigos_cliente[st.session_state.cliente_index]
        
        info_cliente_atual_df = df_info_cliente[df_info_cliente['Cod Clien'] == current_cod_clien]
        
        if not info_cliente_atual_df.empty:
            cliente_detalhe = info_cliente_atual_df.iloc[0]
            vendedores_assoc = info_cliente_atual_df['Vendedor'].unique()
        else:
            cliente_detalhe = df_base[df_base['CodClien'] == current_cod_clien].iloc[0]
            vendedores_assoc = ["N/D"]

        col_nav2.subheader(f"({current_cod_clien}) - {cliente_detalhe['Cliente']} ({st.session_state.cliente_index + 1}/{total_clientes})")
        
        st.subheader("Histórico Agenda")
        
        historico_cliente = df_base[df_base['CodClien'] == current_cod_clien].copy()
        
        col_hist1, col_hist2 = st.columns(2)
        with col_hist1:
            st.caption("Encerrados")
            st.dataframe(historico_cliente[historico_cliente['Sit'] == 'EN'][['DataEncer', 'UsuarioEncer', 'Resultado', 'Motivo_Final', 'Obs']].sort_values('DataEncer', ascending=False), use_container_width=True)
        with col_hist2:
            st.caption("Em Aberto")
            st.dataframe(historico_cliente[historico_cliente['Sit'] == 'AB'][['DataAgenda', 'UsuarioAgenda']].sort_values('DataAgenda'), use_container_width=True)

        st.markdown("---")

        with st.container(border=True):
            col_form1, col_form2 = st.columns(2)
            with col_form1:
                st.caption("Segmento")
                st.markdown(f"`{cliente_detalhe.get('Segmento', 'N/D')}`")
                st.caption("Área de Venda")
                st.markdown(f"`{cliente_detalhe.get('AreaVenda', 'N/D')}`")
                st.caption("Endereço")
                st.markdown(f"`{cliente_detalhe.get('Endereco', '')}, {cliente_detalhe.get('Bairro', '')}`")
                st.caption("Local")
                st.markdown(f"`{cliente_detalhe.get('Municipio', '')} - {cliente_detalhe.get('UF', '')}`")
                st.caption("Vendedores Associados (Base)")
                st.code("\n".join(vendedores_assoc), language=None) 
            with col_form2:
                cond_pgto = cliente_detalhe.get('CondPgto', pd.NA)
                if pd.isna(cond_pgto):
                    cond_pgto = "Depósito Antecipado - à vista"
                st.caption("Cond. Pagamento")
                st.markdown(f"`{cond_pgto}`")
                st.caption("Situação Crédito")
                st.markdown(f"`{cliente_detalhe.get('SitCred', 'N/D')}`")
                st.caption("Limite Total")
                lim = cliente_detalhe.get('LimiteTotal', 0)
                st.markdown(f"`R$ {lim:,.2f}`")
                dt_ult_ped = cliente_detalhe.get('DtUltPed', pd.NaT)
                dias_s_compra = cliente_detalhe.get('DiasSemCompra', 0)
                
                if pd.notna(dt_ult_ped):
                    dt_ult_ped_str = dt_ult_ped.strftime('%d/%m/%Y')
                    if dias_s_compra == 9999:
                         dias_sem_compra_str = " (Nunca comprou)"
                    else:
                         dias_sem_compra_str = f" ({dias_s_compra} dias)"
                else:
                    dt_ult_ped_str = "Nenhum pedido encontrado"
                    dias_sem_compra_str = ""
                
                st.caption("Data Últ. Pedido")
                st.markdown(f"`{dt_ult_ped_str}{dias_sem_compra_str}`")
                inad = cliente_detalhe.get('Inad-3dd', 0)
                st.caption("Inadimplência (-3d)")
                st.markdown(f"`R$ {inad:,.2f}`")
                
        st.markdown("---")
        
        st.subheader("📝 Notas do Supervisor")
        
        notas_carregadas = carregar_notas()
        lista_notas_cliente = notas_carregadas.get(str(current_cod_clien), [])
        
        if isinstance(lista_notas_cliente, str):
            lista_notas_cliente = [{"id": "old", "data": "Antigo", "texto": lista_notas_cliente, "imagem": None}]

        if lista_notas_cliente:
            st.write("Histórico de Notas:")
            for i, nota in enumerate(lista_notas_cliente):
                with st.expander(f"{nota.get('data', 'S/D')} - {nota.get('texto')[:50]}...", expanded=(i==0)):
                    col_txt, col_del = st.columns([0.9, 0.1])
                    with col_txt:
                        st.write(nota.get('texto'))
                        if nota.get('imagem') and os.path.exists(nota.get('imagem')):
                            st.image(nota.get('imagem'), caption="Anexo", use_column_width=True)
                            
                    with col_del:
                        if st.button("🗑️", key=f"del_{nota.get('id')}_{i}"):
                            excluir_nota(current_cod_clien, nota.get('id'))
                            st.rerun()
        else:
            st.info("Nenhuma nota registrada para este cliente.")

        with st.form(key='form_nova_nota', clear_on_submit=True):
            texto_nova_nota = st.text_area("Adicionar nova observação:", height=100)
            
            uploaded_file = st.file_uploader("Anexar Imagem (Opcional)", type=['png', 'jpg', 'jpeg'])
            
            if st.form_submit_button("Salvar Nota 💾"):
                if texto_nova_nota.strip():
                    adicionar_nota(current_cod_clien, texto_nova_nota, uploaded_file)
                    st.success("Nota adicionada!")
                    st.rerun()
                else:
                    st.warning("Escreva algo para salvar.")

        st.markdown("#### Exportação de Notas (Filtrada)")
        if st.button("Baixar planilha de notas (Conforme Filtros) 📥"):
            df_export_base, pivot_cols_names = calcular_resumo_power_query(df_filtrado_com_data, df_base, df_info_cliente)
            notas_carregadas = carregar_notas()
            
            linhas_export = []
            
            cols_export_req = [
                'DtUltPed', 'UsuarioEncer', 'CodClien', 'Cliente', 
                'Base', 'Total Mês', 'Ult 3 Meses', 
                'Total de Contatos', 'Data Ult Contato', 'Ult Resultado', 
                'Ult Obs', 'Próx Agendamento'
            ]
            
            pivot_cols_dynamic = [c for c in pivot_cols_names if c not in ['UsuarioEncer', 'CodClien', 'Cliente']]
            cols_base = cols_export_req + pivot_cols_dynamic
            cols_existentes = [c for c in cols_base if c in df_export_base.columns]
            
            for idx, row in df_export_base.iterrows():
                cid = str(row['CodClien'])
                notas_cli = notas_carregadas.get(cid, [])
                
                if isinstance(notas_cli, str): notas_cli = [{"data": "N/D", "texto": notas_cli, "imagem": None}]
                
                if notas_cli:
                    for nota in notas_cli:
                        linha_nova = row[cols_existentes].to_dict()
                        linha_nova['Data Nota'] = nota.get('data', '')
                        linha_nova['Nota'] = nota.get('texto', '')
                        linha_nova['CaminhoImagem'] = nota.get('imagem') 
                        linhas_export.append(linha_nova)
            
            if linhas_export:
                df_final_export = pd.DataFrame(linhas_export)
                
                colunas_finais_ordem = [
                    'DtUltPed', 'UsuarioEncer', 'CodClien', 'Cliente',
                    'Nota', 'Data Nota',
                    'Base', 'Total Mês', 'Ult 3 Meses',
                    'Total de Contatos', 'Data Ult Contato', 'Ult Resultado',
                    'Ult Obs', 'Próx Agendamento'
                ]
                colunas_dinamicas = [c for c in df_final_export.columns if c not in colunas_finais_ordem and c != 'Notas do Supervisor' and c != 'CaminhoImagem']
                colunas_finais_ordem += colunas_dinamicas
                
                cols_validas = [c for c in colunas_finais_ordem if c in df_final_export.columns]
                if 'CaminhoImagem' in df_final_export.columns:
                     cols_validas.append('CaminhoImagem')
                
                df_final_export = df_final_export[cols_validas]
                
                xlsx_notas = to_excel_com_imagens(df_final_export)
                
                st.download_button("Clique para baixar Excel", data=xlsx_notas, file_name="notas_detalhadas_com_imagens.xlsx")
            else:
                st.warning("Nenhum registro com nota encontrada para os filtros atuais.")


# --- ABA 3 (SEM DATA, FILTRO AGENDAMENTO DIRETO) ---
with tab3:
    st.subheader("Agendamentos em Aberto (Geral)")
    
    if filtro_tipo_agendamento == "Sem Agendamento":
        st.warning("O filtro 'Sem Agendamento' está ativo. Esta aba exibe apenas quem POSSUI agendamentos.")
        df_abertos_tab3 = pd.DataFrame()
    else:
        df_abertos_tab3 = df_sched_filtered.copy()
    
    if not df_abertos_tab3.empty:
        df_view = pd.merge(
            df_abertos_tab3,
            df_ultimos_contatos_usuario[['CodClien', 'UsuarioEncer_JoinKey', 'Ultimo Motivo', 'Data Ult Contato']], 
            left_on=['CodClien', 'UsuarioAgenda'],
            right_on=['CodClien', 'UsuarioEncer_JoinKey'],
            how='left'
        )
        df_view['Ultimo Motivo'] = df_view['Ultimo Motivo'].fillna('')
        
        df_info_resumo = df_info_cliente[['Cod Clien', 'Cod Vend', 'DtUltPed']].copy()
        df_info_resumo.columns = ['CodClien', 'Cod Vend Base', 'Data Ult Pedido']
        
        df_view_final = pd.merge(
            df_view,
            df_info_resumo,
            on='CodClien',
            how='left'
        )
        
        if 'UsuarioEncer_JoinKey' in df_view_final.columns:
            df_view_final = df_view_final.drop(columns=['UsuarioEncer_JoinKey'])

        df_view_final = df_view_final.sort_values(by=['DataAgenda', 'CodClien'])
        df_view_final = df_view_final.drop_duplicates(subset=['DataAgenda', 'UsuarioAgenda', 'CodClien'])

        df_view_final['CodClien'] = df_view_final['CodClien'].astype(str)

        st.info(f"Exibindo {len(df_view_final)} agendamentos.")
        
        st.dataframe(
            df_view_final[[
                'DataAgenda', 'UsuarioAgenda', 'Cliente', 'CodClien', 
                'Cod Vend Base', 'Data Ult Pedido', 
                'Ultimo Motivo', 'Data Ult Contato', 'Obs' 
            ]],
            use_container_width=True,
            column_config={
                "DataAgenda": st.column_config.DatetimeColumn("Data Agendada", format="DD/MM/YYYY HH:mm"),
                "Data Ult Pedido": st.column_config.DatetimeColumn("Data Últ. Pedido", format="DD/MM/YYYY"),
                "Data Ult Contato": st.column_config.DatetimeColumn("Data Últ. Contato", format="DD/MM/YYYY HH:mm"),
                "Cod Vend Base": "Cód Vend (Base)",
                "Ultimo Motivo": "Último Motivo (Histórico)"
            }
        )

# --- ABA 4 (COM DATA, USANDO DF_FILTRADO_COM_DATA) ---
with tab4:
    st.subheader("Resumo por Cliente")
    
    df_rows_input = df_filtrado_com_data
    
    if df_rows_input.empty and not df_common_agend_filtered.empty:
         df_rows_input = df_common_agend_filtered.copy()
    
    df_final, pivot_cols_names = calcular_resumo_power_query(df_rows_input, df_base, df_info_cliente)
    
    cols_order = [
        'DtUltPed', 'UsuarioEncer', 'CodClien', 'Cliente', 'Base', 
        'Total Mês', 'Ult 3 Meses', 'Total de Contatos', 
        'Data Ult Contato', 'Ult Resultado', 'Ult Obs', 'Próx Agendamento'
    ]
    pivot_cols_list = [c for c in pivot_cols_names if c not in ['UsuarioEncer', 'CodClien', 'Cliente']]
    final_cols = cols_order + pivot_cols_list
    
    final_cols = [c for c in final_cols if c in df_final.columns]
    
    cols_numeric_fillna = [c for c in final_cols if c not in ['DtUltPed', 'Data Ult Contato', 'Próx Agendamento', 'Cliente', 'Ult Resultado', 'Ult Obs']]
    df_final[cols_numeric_fillna] = df_final[cols_numeric_fillna].fillna(0)
    
    df_final = df_final[final_cols]
    
    if not df_final.empty:
        df_final['CodClien'] = df_final['CodClien'].astype(str)
        
        if 'DiasSemCompra' in df_final.columns:
             df_final['DiasSemCompra'] = df_final['DiasSemCompra'].replace(9999, None)
             
        st.dataframe(
            df_final,
            use_container_width=True,
            column_config={
                "DtUltPed": st.column_config.DateColumn("Data Últ Ped"),
                "Data Ult Contato": st.column_config.DatetimeColumn("Últ Contato", format="DD/MM/YYYY HH:mm"),
                "Próx Agendamento": st.column_config.DatetimeColumn("Próx Agend", format="DD/MM/YYYY HH:mm"),
            }
        )
        
        excel_pq = to_excel(df_final, sheet_name="Resumo")
        st.download_button("📥 Baixar Tabela Resumo (Excel)", data=excel_pq, file_name="resumo.xlsx")
    else:
        st.warning("Nenhum dado encontrado para os filtros selecionados.")


# --- ABA 5 (AUTOMAÇÃO DE AGENDA) ---
with tab5:
    st.subheader("🤖 Gerador de Agendas Automático (Importação)")
    st.markdown("""
    Gera listas de códigos de clientes para serem importados no sistema de agendamento.
    **Prioridade:** 1. Atrasados; 2. Sem compra no mês; 3. Nunca atendidos; 4. Frequência baixa; 5. Dias sem compra.
    """)
    
    col_sel_vendedor, col_sel_cod_base = st.columns(2)
    
    with col_sel_vendedor:
        vendedor_alvo = st.selectbox(
            "Selecione o Vendedor (UsuarioAgenda):", 
            options=lista_vendedores
        )
        
    with col_sel_cod_base:
        cod_vend_sugerido = ""
        try:
            cod_cli_exemplo = df_base[df_base['UsuarioAgenda'] == vendedor_alvo]['CodClien'].iloc[0]
            cod_vend_sugerido = df_info_cliente[df_info_cliente['Cod Clien'] == cod_cli_exemplo]['Cod Vend'].iloc[0]
        except:
            cod_vend_sugerido = ""
            
        cod_vend_base_input = st.text_input(
            "Confirme o Código do Vendedor na BASE (Obrigatório):", 
            value=cod_vend_sugerido,
            help="Esse código define quem são os clientes da carteira deste vendedor."
        )

    with st.expander("🛠️ Configurações e Filtros Avançados (Clique para abrir)"):
        st.info("Estes filtros aplicam-se apenas à CARTEIRA (Backlog). Clientes ATRASADOS sempre serão priorizados independentemente dos filtros.")
        
        col_adv1, col_adv2, col_adv3 = st.columns(3)
        
        with col_adv1:
            st.markdown("**⚙️ Parâmetros**")
            meta_diaria = st.number_input("Capacidade Diária (Clientes):", value=35, min_value=1)
            dias_projecao = st.number_input("Projetar para quantos dias úteis?", value=5, min_value=1, max_value=20)
            
            filtro_incluir_ja_atendidos = st.checkbox("Incluir clientes com contato anterior", value=True, help="Se desmarcado, sugere APENAS clientes que NUNCA foram contatados por este vendedor.")
            
            st.markdown("---")
            st.markdown("**💰 Financeiro**")
            filtro_limite_min = st.number_input("Limite de Crédito Mínimo (R$):", min_value=0.0, value=0.0, step=500.0)
            
            filtro_apenas_prazo = st.checkbox("Apenas Clientes com Prazo (Sem Depósito)", value=False)
            filtro_apenas_adimplentes = st.checkbox("Apenas Adimplentes (Sem títulos > 3 dias)", value=True)

        with col_adv2:
            st.markdown("**📍 Segmentação**")
            
            filtro_incluir_grupo_colig = st.checkbox("Incluir Clientes de Grupo/Coligação", value=True)
            
            opts_seg = sorted(df_info_cliente['Segmento'].fillna('').unique())
            filtro_segmentos = st.multiselect("Segmento:", options=opts_seg)
            
            opts_mun = sorted(df_info_cliente['Municipio'].fillna('').unique())
            filtro_municipios = st.multiselect("Município:", options=opts_mun)
            
            opts_area = sorted(df_info_cliente['AreaVenda'].fillna('').unique())
            filtro_areas = st.multiselect("Área de Venda:", options=opts_area)

        with col_adv3:
            st.markdown("**📊 Perfil**")
            
            opts_cred = sorted(df_info_cliente['SitCred'].fillna('').unique())
            filtro_sit_cred = st.multiselect("Situação Crédito:", options=opts_cred)
            
            df_info_cliente['CondPgto_Temp'] = df_info_cliente['CondPgto'].apply(lambda x: 'Depósito Antecipado' if x == '' or pd.isna(x) else x)
            opts_cond = sorted(df_info_cliente['CondPgto_Temp'].unique())
            filtro_cond_pgto = st.multiselect("Cond. Pagamento:", options=opts_cond)
            
            filtro_dias_sem_compra_min = st.number_input("Dias Sem Compra (Mínimo):", min_value=0, value=0)

    
    if st.button("Gerar Sugestão de Agenda 🚀"):
        if not vendedor_alvo or not cod_vend_base_input:
            st.error("Selecione o vendedor e verifique o Código Base.")
        else:
            filtros_dict = {
                "cod_vend_base": cod_vend_base_input,
                "municipios": filtro_municipios,
                "segmentos": filtro_segmentos,
                "areas": filtro_areas,
                "sit_cred": filtro_sit_cred,
                "cond_pgto": filtro_cond_pgto,
                "limite_min": filtro_limite_min,
                "dias_sem_compra_min": filtro_dias_sem_compra_min,
                "apenas_adimplentes": filtro_apenas_adimplentes,
                "filtro_notas": filtro_notas,
                "apenas_com_prazo": filtro_apenas_prazo, 
                "incluir_grupo_colig": filtro_incluir_grupo_colig, 
                "incluir_ja_atendidos": filtro_incluir_ja_atendidos
            }
            
            with st.spinner("Processando lógica de prioridade e atrasos..."):
                datas_uteis, df_sugestao, df_metricas, df_diagnostico, msg_status = gerar_sugestao_agenda(
                    df_base, df_info_cliente, 
                    vendedor_alvo, meta_diaria, dias_projecao, 
                    filtros_dict
                )
            
            if df_sugestao.empty:
                st.warning(msg_status)
            else:
                st.success(f"Agenda gerada com sucesso! {len(df_sugestao)} clientes distribuídos.")
                
                with st.expander("🔍 Diagnóstico: Por que alguns clientes não apareceram? (Clique para ver)"):
                    st.dataframe(df_diagnostico, use_container_width=True)
                    st.caption("*Nota: 'Agendamento (Mesmo Vendedor)' remove clientes que JÁ estão na agenda futura DESTE usuário.*")

                if not df_metricas.empty:
                    with st.expander("📊 Detalhes da Capacidade Diária (Por que sugeriu essa quantidade?)"):
                        st.dataframe(df_metricas, use_container_width=True)

                st.markdown("### 📅 Listas para Importação")
                
                grupos_por_data = df_sugestao.groupby('Data Sugerida')
                
                for data_util in datas_uteis:
                    if data_util in grupos_por_data.groups:
                        grupo = grupos_por_data.get_group(data_util)
                        dia_semana = data_util.strftime('%A')
                        dias_pt = {'Monday': 'Segunda', 'Tuesday': 'Terça', 'Wednesday': 'Quarta', 'Thursday': 'Quinta', 'Friday': 'Sexta'}
                        dia_nome = dias_pt.get(dia_semana, dia_semana)
                        
                        st.markdown(f"**{data_util.strftime('%d/%m/%Y')} ({dia_nome}) - {len(grupo)} Clientes**")
                        
                        s_cods = pd.to_numeric(grupo['Cod Clien'], errors='coerce').fillna(0).astype(int)
                        s_cods = s_cods[s_cods != 0]
                        codigos_lista = s_cods.astype(str).tolist()
                        string_copia = ", ".join(codigos_lista)
                        
                        st.code(string_copia, language="text")
                        
                        with st.expander(f"Ver detalhes dos clientes de {data_util.strftime('%d/%m')}"):
                            
                            cols_visual = [
                                'Cod Vend', 'Cod Clien', 'Cliente', 'MotivoPrioridade',
                                'Ult_Motivo_Vend', 'Dt_Ult_Motivo_Vend', 'Ult_Obs_Vend',
                                'Segmento', 'Municipio', 'AreaVenda', 
                                'SitCred', 'LimiteTotal', 'DtUltPed', 
                                'DiasSemCompra', 'Inad-3dd', 'CondPgto_Visual'
                            ]
                            
                            cols_final = [c for c in cols_visual if c in grupo.columns]
                            
                            st.dataframe(
                                grupo[cols_final], 
                                use_container_width=True,
                                hide_index=True,
                                column_config={
                                    "Cod Clien": st.column_config.TextColumn("Cód. Cliente"), 
                                    "Cod Vend": st.column_config.TextColumn("Cód. Vend"),
                                    "LimiteTotal": st.column_config.NumberColumn("Limite Total", format="R$ %.2f"),
                                    "Inad-3dd": st.column_config.NumberColumn("Inadimplência", format="R$ %.2f"),
                                    "DtUltPed": st.column_config.DateColumn("Últ. Pedido", format="DD/MM/YYYY"),
                                    "MotivoPrioridade": st.column_config.TextColumn("Status/Prioridade"),
                                    "CondPgto_Visual": st.column_config.TextColumn("Cond. Pagamento"),
                                    "DiasSemCompra": st.column_config.NumberColumn("Dias s/ Compra"),
                                    "Ult_Motivo_Vend": "Últ. Motivo (Deste Vendedor)",
                                    "Dt_Ult_Motivo_Vend": st.column_config.DatetimeColumn("Data Últ. Contato", format="DD/MM/YYYY HH:mm"),
                                    "Ult_Obs_Vend": "Últ. Obs"
                                }
                            )
                        st.markdown("---")
                
                col_down1, col_down2 = st.columns(2)
                excel_sugestao = to_excel(df_sugestao, sheet_name="Importacao")
                col_down1.download_button(
                    "📥 Baixar Lista Completa (Excel)",
                    data=excel_sugestao,
                    file_name=f"agenda_sugerida_{vendedor_alvo}.xlsx"
                )
