import streamlit as st
import pandas as pd
import plotly.express as px
import datetime
from dateutil.relativedelta import relativedelta
from io import BytesIO
import json
import os
import uuid
import numpy as np
import random

# --- Configuração Inicial ---
st.set_page_config(layout="wide", page_title="Dashboard de Vendas e Agenda (DEMO)")

ARQUIVO_NOTAS = "notas_supervisor.json"
PASTA_IMAGENS = "notas_imagens"  # Pasta para salvar as fotos

if not os.path.exists(PASTA_IMAGENS):
    os.makedirs(PASTA_IMAGENS)

# ==========================================
# GERADOR DE DADOS FICTÍCIOS (MOCK DATA)
# ==========================================
def gerar_dados_ficticios():
    np.random.seed(42) # Mantém os dados consistentes a cada recarregamento
    random.seed(42)
    
    qtd_clientes = 200
    vendedores = ['8701', '8702', '8705', '8708', '8724', '8725', '8727', '8728', '8740', '8741']
    
    # 1. Base de Clientes
    clientes_data = []
    for i in range(1, qtd_clientes + 1):
        cod_clien = 1000 + i
        dt_ult_ped = pd.Timestamp.today() - pd.Timedelta(days=np.random.randint(0, 400))
        if np.random.rand() > 0.8: # 20% nunca comprou
            dt_ult_ped = pd.NaT
            
        cod_vend_sorteado = np.random.choice(vendedores)
            
        clientes_data.append({
            'Cod Clien': cod_clien,
            'Cliente': f'CLIENTE FICTICIO {cod_clien} LTDA',
            'Cod Vend': cod_vend_sorteado,
            'Vendedor': f'VENDEDOR {cod_vend_sorteado}', # <-- CORREÇÃO: Coluna adicionada
            'Segmento': np.random.choice(['VAREJO', 'ATACADO', 'DISTRIBUIDOR', 'FARMACIA', 'MERCADO']),
            'AreaVenda': np.random.choice(['ZONA NORTE', 'ZONA SUL', 'CENTRO', 'BAIXADA', 'INTERIOR']),
            'Municipio': np.random.choice(['RIO DE JANEIRO', 'NOVA IGUACU', 'DUQUE DE CAXIAS', 'NITEROI', 'SAO GONCALO']),
            'UF': 'RJ',
            'Endereco': 'RUA FICTICIA, 123',
            'Bairro': 'BAIRRO TESTE',
            'Colig': '', 'Grupo': '',
            'SitCred': np.random.choice(['NORMAL', 'BLOQUEADO', 'ANALISE'], p=[0.8, 0.1, 0.1]),
            'CondPgto': np.random.choice(['21 Dias', '14/28 Dias', 'Depósito Antecipado', 'À Vista'], p=[0.4, 0.3, 0.2, 0.1]),
            'LimiteTotal': np.random.uniform(1000, 50000),
            'DtUltPed': dt_ult_ped,
            'Inad-3dd': 0 if np.random.rand() > 0.15 else np.random.uniform(100, 5000) # 15% inadimplentes
        })
    df_info = pd.DataFrame(clientes_data)
    
    # 2. Base de Agenda/Contatos
    contatos_data = []
    qtd_contatos = 1500
    motivos = ['COTAÇÃO', 'SEM INTERESSE', 'CLIENTE NAO ATENDE', 'COMPROU CONCORRENTE', 'REAGENDAMENTO', 'VENDA']
    
    hoje = pd.Timestamp.today()
    
    for _ in range(qtd_contatos):
        cod_clien = np.random.choice(df_info['Cod Clien'])
        vend = np.random.choice(vendedores)
        sit = np.random.choice(['EN', 'AB'], p=[0.8, 0.2]) # 80% encerrados, 20% em aberto
        
        data_encer = pd.NaT
        data_agenda = pd.NaT
        resultado = ''
        motivo = ''
        
        if sit == 'EN':
            # Data aleatória nos últimos 6 meses
            dias_atras = np.random.randint(0, 180)
            data_encer = hoje - pd.Timedelta(days=dias_atras)
            motivo = np.random.choice(motivos, p=[0.2, 0.2, 0.3, 0.1, 0.1, 0.1])
            resultado = 'NAO VENDA' if motivo != 'VENDA' else 'VENDA'
        else:
            # Agendamento entre 15 dias atrás (atrasado) e 30 dias pra frente
            dias_diff = np.random.randint(-15, 30)
            data_agenda = hoje + pd.Timedelta(days=dias_diff)
            
        contatos_data.append({
            'CodClien': cod_clien,
            'Cliente': f'CLIENTE FICTICIO {cod_clien} LTDA', # <-- CORREÇÃO: Coluna adicionada
            'UsuarioAgenda': vend,
            'UsuarioEncer': vend if sit == 'EN' else '',
            'DataEncer': data_encer,
            'DataAgenda': data_agenda,
            'Sit': sit,
            'Resultado': resultado,
            'Motivo': motivo,
            'Obs': 'Contato realizado via sistema demo.' if sit == 'EN' else 'Ligar pela manhã.'
        })
        
    df_agenda = pd.DataFrame(contatos_data)
    return df_agenda, df_info
# ==========================================

# --- Funções Auxiliares (Persistência e Excel) ---

def carregar_notas():
    if not os.path.exists(ARQUIVO_NOTAS):
        return {}
    try:
        with open(ARQUIVO_NOTAS, "r", encoding="utf-8") as f:
            dados = json.load(f)
        
        novo_formato = {}
        alterou = False
        for k, v in dados.items():
            if isinstance(v, str): 
                novo_formato[k] = [{"id": str(uuid.uuid4()), "data": datetime.datetime.now().strftime("%d/%m/%Y %H:%M"), "texto": v, "imagem": None}]
                alterou = True
            else:
                novo_formato[k] = v
        if alterou:
            with open(ARQUIVO_NOTAS, "w", encoding="utf-8") as f:
                json.dump(novo_formato, f, ensure_ascii=False, indent=4)
        return novo_formato
    except:
        return {}

def adicionar_nota(cod_cliente, texto, arquivo_imagem=None):
    notas_dict = carregar_notas()
    cod = str(cod_cliente)
    caminho_imagem = None
    
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

    nova_nota = {"id": str(uuid.uuid4()), "data": datetime.datetime.now().strftime("%d/%m/%Y %H:%M"), "texto": texto, "imagem": caminho_imagem}
    
    if cod not in notas_dict: notas_dict[cod] = []
    notas_dict[cod].insert(0, nova_nota)
    
    with open(ARQUIVO_NOTAS, "w", encoding="utf-8") as f:
        json.dump(notas_dict, f, ensure_ascii=False, indent=4)

def excluir_nota(cod_cliente, note_id):
    notas_dict = carregar_notas()
    cod = str(cod_cliente)
    if cod in notas_dict:
        nota_a_remover = next((n for n in notas_dict[cod] if n['id'] == note_id), None)
        if nota_a_remover and nota_a_remover.get('imagem'):
            if os.path.exists(nota_a_remover['imagem']):
                try: os.remove(nota_a_remover['imagem'])
                except: pass
        notas_dict[cod] = [n for n in notas_dict[cod] if n['id'] != note_id]
        if not notas_dict[cod]: del notas_dict[cod]
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
    return output.getvalue()

def to_excel_com_imagens(df_notas):
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
                    worksheet.insert_image(i + 1, col_imagem_idx, caminho_img, {'x_scale': 0.2, 'y_scale': 0.2, 'object_position': 1})
                except:
                    worksheet.write(i + 1, col_imagem_idx, "Erro img")
            else:
                worksheet.write(i + 1, col_imagem_idx, "-")
    return output.getvalue()

def calcular_resumo_power_query(df_input_rows, df_base_full, df_info_clientes):
    # Blindagem 1: Se a tabela chegar vazia, já encerra sem erro
    if df_input_rows.empty: return pd.DataFrame(), []

    # Blindagem 2: Verifica se a coluna Cliente existe antes de fatiar
    colunas_base = ['UsuarioEncer', 'CodClien']
    if 'Cliente' in df_input_rows.columns:
        df_structure = df_input_rows[colunas_base + ['Cliente']].drop_duplicates()
    else:
        # Se não existir, pega só o que tem e busca o nome na base de informações
        df_structure = df_input_rows[colunas_base].drop_duplicates()
        temp_clientes = df_info_clientes[['Cod Clien', 'Cliente']].drop_duplicates('Cod Clien')
        df_structure = df_structure.merge(temp_clientes, left_on='CodClien', right_on='Cod Clien', how='left')
        df_structure = df_structure.drop(columns=['Cod Clien'], errors='ignore')
        df_structure['Cliente'] = df_structure['Cliente'].fillna('Cliente S/ Cadastro')

    # Daqui para baixo o código segue normal...
    df_hist_pivot = df_input_rows[df_input_rows['Sit'] == 'EN'].copy()
    df_hist_pivot['Motivo2'] = df_hist_pivot['Motivo_Final']
    
    pivot_table = pd.pivot_table(df_hist_pivot, index=['UsuarioEncer', 'CodClien'], columns='Motivo2', aggfunc='size', fill_value=0).reset_index()
    
    hoje = datetime.date.today()
    inicio_mes_atual = hoje.replace(day=1)
    ts_inicio_mes_atual = pd.Timestamp(inicio_mes_atual)
    
    df_contexto = df_base_full[(df_base_full['CodClien'].isin(df_structure['CodClien'])) & (df_base_full['UsuarioEncer'].isin(df_structure['UsuarioEncer']))]
    df_mes = df_contexto[df_contexto['DataEncer'] >= ts_inicio_mes_atual].groupby(['CodClien', 'UsuarioEncer']).size().reset_index(name='Total Mês')
    
    inicio_3m = (inicio_mes_atual - relativedelta(months=3))
    ts_inicio_3m = pd.Timestamp(inicio_3m)
    df_3m_full = df_contexto[(df_contexto['DataEncer'] >= ts_inicio_3m) & (df_contexto['DataEncer'] < ts_inicio_mes_atual)]
    df_3m = df_3m_full.groupby(['CodClien', 'UsuarioEncer']).size().reset_index(name='Ult 3 Meses')
    
    df_total = df_input_rows.groupby(['CodClien', 'UsuarioEncer']).size().reset_index(name='Total de Contatos')
    df_dt_ped = df_base_full.groupby('CodClien')['DtUltPed'].max().reset_index()
    
    df_last = df_input_rows.sort_values('DataEncer', ascending=False).drop_duplicates(['CodClien', 'UsuarioEncer'])[['CodClien', 'UsuarioEncer', 'DataEncer', 'Motivo_Final', 'Obs']]
    df_last.columns = ['CodClien', 'UsuarioEncer', 'Data Ult Contato', 'Ult Resultado', 'Ult Obs']
    
    df_prox = df_base_full[df_base_full['Sit']=='AB'].copy().sort_values('DataAgenda').drop_duplicates(subset=['CodClien', 'UsuarioAgenda'])
    df_prox = df_prox[['CodClien', 'UsuarioAgenda', 'DataAgenda']]
    df_prox.columns = ['CodClien', 'UsuarioEncer', 'Próx Agendamento']
    
    df_base_vend = df_info_clientes[['Cod Clien', 'Cod Vend']].drop_duplicates('Cod Clien')
    df_base_vend.columns = ['CodClien', 'Base']
    
    df_final = df_structure.merge(pivot_table, on=['UsuarioEncer', 'CodClien'], how='left')
    df_final = df_final.merge(df_dt_ped, on='CodClien', how='left').merge(df_mes, on=['CodClien', 'UsuarioEncer'], how='left').merge(df_3m, on=['CodClien', 'UsuarioEncer'], how='left').merge(df_total, on=['CodClien', 'UsuarioEncer'], how='left').merge(df_last, on=['CodClien', 'UsuarioEncer'], how='left').merge(df_prox, on=['CodClien', 'UsuarioEncer'], how='left').merge(df_base_vend, on='CodClien', how='left')
    
    return df_final, pivot_table.columns

def calcular_dias_uteis(data_inicial, dias_a_frente):
    datas = []
    dia_atual = data_inicial
    while len(datas) < dias_a_frente:
        if dia_atual.weekday() < 5: datas.append(dia_atual)
        dia_atual += datetime.timedelta(days=1)
    return datas

def gerar_sugestao_agenda(df_base_agenda, df_info_clientes_raw, vendedor_agenda, capacidade_diaria, dias_projecao, filtros_avancados):
    hoje = datetime.date.today()
    ts_hoje = pd.Timestamp(hoje)
    ts_inicio_mes = pd.Timestamp(hoje.replace(day=1))
    
    diagnostico = {"Total Carteira (Bruto)": 0, "Removidos: Agendamento (Mesmo Vendedor)": 0, "Removidos: Atrasados (Já Inclusos)": 0, "Removidos: Filtros (Mun/Seg/Etc)": 0, "Removidos: Já Atendidos (Filtro)": 0, "Total Final Backlog": 0, "Total Atrasados": 0}

    df_info_clientes = df_info_clientes_raw.drop_duplicates(subset=['Cod Clien']).copy()
    df_info_clientes['Cod Vend Limpo'] = df_info_clientes['Cod Vend'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    
    for c in ['Segmento', 'AreaVenda', 'SitCred', 'CondPgto', 'Municipio', 'Grupo', 'Colig']:
        if c in df_info_clientes.columns: df_info_clientes[c] = df_info_clientes[c].fillna('')
            
    df_info_clientes['CondPgto_Visual'] = df_info_clientes['CondPgto'].apply(lambda x: 'Depósito Antecipado' if x == '' or pd.isna(x) else x)

    df_hist_vendedor = df_base_agenda[(df_base_agenda['UsuarioEncer'] == vendedor_agenda) & (df_base_agenda['Sit'] == 'EN')].copy()
    if not df_hist_vendedor.empty:
        df_hist_vendedor = df_hist_vendedor.sort_values('DataEncer', ascending=False).drop_duplicates('CodClien')[['CodClien', 'Motivo_Final', 'DataEncer', 'Obs']]
        df_hist_vendedor.columns = ['Cod Clien', 'Ult_Motivo_Vend', 'Dt_Ult_Motivo_Vend', 'Ult_Obs_Vend']
    else:
        df_hist_vendedor = pd.DataFrame(columns=['Cod Clien', 'Ult_Motivo_Vend', 'Dt_Ult_Motivo_Vend', 'Ult_Obs_Vend'])

    df_atrasados_raw = df_base_agenda[(df_base_agenda['UsuarioAgenda'] == vendedor_agenda) & (df_base_agenda['Sit'] == 'AB') & (df_base_agenda['DataAgenda'] < ts_hoje)].copy().sort_values('DataAgenda').drop_duplicates('CodClien')
    
    # --- BLINDAGEM CONTRA KEYERROR DE SUFIXOS (_x, _y) ---
    if 'Cod Clien' in df_atrasados_raw.columns:
        df_atrasados_raw = df_atrasados_raw.drop(columns=['Cod Clien'])
        
    df_atrasados_full = df_atrasados_raw.merge(
        df_info_clientes, left_on='CodClien', right_on='Cod Clien', how='left'
    ).merge(
        df_hist_vendedor, on='Cod Clien', how='left'
    )
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
        if filtro_nota_opcao == 'Com Notas': df_carteira = df_carteira[df_carteira['Cod_Str'].isin(ids_com_nota)]
        elif filtro_nota_opcao == 'Sem Notas': df_carteira = df_carteira[~df_carteira['Cod_Str'].isin(ids_com_nota)]

    if filtros_avancados.get('apenas_com_prazo'): df_carteira = df_carteira[df_carteira['CondPgto_Visual'] != 'Depósito Antecipado']
    if not filtros_avancados.get('incluir_grupo_colig'): df_carteira = df_carteira[(df_carteira['Grupo'] == '') & (df_carteira['Colig'] == '')]
    if filtros_avancados.get('municipios'): df_carteira = df_carteira[df_carteira['Municipio'].isin(filtros_avancados['municipios'])]
    if filtros_avancados.get('segmentos'): df_carteira = df_carteira[df_carteira['Segmento'].isin(filtros_avancados['segmentos'])]
    if filtros_avancados.get('areas'): df_carteira = df_carteira[df_carteira['AreaVenda'].isin(filtros_avancados['areas'])]
    if filtros_avancados.get('sit_cred'): df_carteira = df_carteira[df_carteira['SitCred'].isin(filtros_avancados['sit_cred'])]
    if filtros_avancados.get('cond_pgto'): df_carteira = df_carteira[df_carteira['CondPgto_Visual'].isin(filtros_avancados['cond_pgto'])]
    
    val_limite = filtros_avancados.get('limite_min', 0.0)
    if val_limite > 0: df_carteira = df_carteira[df_carteira['LimiteTotal'] >= val_limite]
    val_dias_min = filtros_avancados.get('dias_sem_compra_min', 0)
    if val_dias_min > 0: df_carteira = df_carteira[df_carteira['DiasSemCompra'] >= val_dias_min]
    if filtros_avancados.get('apenas_adimplentes'): df_carteira = df_carteira[(df_carteira['Inad-3dd'] <= 0) | (df_carteira['Inad-3dd'].isna())]
    
    len_depois_filtros = len(df_carteira)
    diagnostico["Removidos: Filtros (Mun/Seg/Etc)"] = len_antes - len_depois_filtros

    clientes_com_futuro = df_base_agenda[(df_base_agenda['Sit'] == 'AB') & (df_base_agenda['DataAgenda'] >= ts_hoje) & (df_base_agenda['UsuarioAgenda'] == vendedor_agenda)]['CodClien'].unique()
    df_backlog = df_carteira[~df_carteira['Cod Clien'].isin(clientes_com_futuro)].copy()
    
    len_sem_futuro = len(df_backlog)
    diagnostico["Removidos: Agendamento (Mesmo Vendedor)"] = len_depois_filtros - len_sem_futuro
    df_backlog = df_backlog[~df_backlog['Cod Clien'].isin(lista_ids_atrasados)].copy()
    
    len_final_backlog = len(df_backlog)
    diagnostico["Removidos: Atrasados (Já Inclusos)"] = len_sem_futuro - len_final_backlog
    
    df_backlog['ComprouEsteMes'] = df_backlog['DtUltPed'].apply(lambda x: x.year == hoje.year and x.month == hoje.month if pd.notnull(x) else False)
    clientes_ja_atendidos = df_base_agenda[(df_base_agenda['UsuarioEncer'] == vendedor_agenda) & (df_base_agenda['Sit'] == 'EN')]['CodClien'].unique()
    df_backlog['JaAtendido'] = df_backlog['Cod Clien'].isin(clientes_ja_atendidos)
    
    len_antes_ja_atendidos = len(df_backlog)
    if not filtros_avancados.get('incluir_ja_atendidos', True):
        df_backlog = df_backlog[~df_backlog['JaAtendido']].copy()
    diagnostico["Removidos: Já Atendidos (Filtro)"] = len_antes_ja_atendidos - len(df_backlog)
    diagnostico["Total Final Backlog"] = len(df_backlog)
    
    contatos_mes = df_base_agenda[(df_base_agenda['UsuarioEncer'] == vendedor_agenda) & (df_base_agenda['DataEncer'] >= ts_inicio_mes)].groupby('CodClien').size().reset_index(name='QtdContatosMes')
    df_backlog = df_backlog.merge(contatos_mes, left_on='Cod Clien', right_on='CodClien', how='left')
    df_backlog['QtdContatosMes'] = df_backlog['QtdContatosMes'].fillna(0)
    df_backlog = df_backlog.merge(df_hist_vendedor, on='Cod Clien', how='left')
    df_backlog = df_backlog.sort_values(by=['ComprouEsteMes', 'JaAtendido', 'QtdContatosMes', 'DiasSemCompra'], ascending=[True, True, True, False])
    
    fila_unificada = []
    def extrair_dados(row, motivo_prioridade):
        def get_val(col, default='N/D'): return row.get(col) if pd.notnull(row.get(col)) else default
        cod_final = row.get('Cod Clien') if pd.notnull(row.get('Cod Clien')) else row.get('CodClien')
        return {
            'Cod Clien': cod_final, 'Cliente': get_val('Cliente', 'Cliente S/ Cadastro'), 'Cod Vend': get_val('Cod Vend', ''),
            'Segmento': get_val('Segmento', ''), 'Municipio': get_val('Municipio', ''), 'AreaVenda': get_val('AreaVenda', ''),
            'SitCred': get_val('SitCred', ''), 'LimiteTotal': row.get('LimiteTotal', 0), 'DtUltPed': row.get('DtUltPed', pd.NaT),
            'DiasSemCompra': row.get('DiasSemCompra', 0), 'Inad-3dd': row.get('Inad-3dd', 0),
            'CondPgto_Visual': 'Depósito Antecipado' if pd.isna(row.get('CondPgto')) or row.get('CondPgto') == '' else row.get('CondPgto'),
            'MotivoPrioridade': motivo_prioridade, 'Ult_Motivo_Vend': get_val('Ult_Motivo_Vend', ''), 'Dt_Ult_Motivo_Vend': row.get('Dt_Ult_Motivo_Vend', pd.NaT), 'Ult_Obs_Vend': get_val('Ult_Obs_Vend', '')
        }

    for _, row in df_atrasados_full.iterrows(): fila_unificada.append(extrair_dados(row, '⚠️ 1. Atrasado'))
    for _, row in df_backlog.iterrows():
        if not row['ComprouEsteMes']: label = "2. Sem compra no mês"
        elif not row['JaAtendido']: label = "3. Nunca atendidos"
        elif row['QtdContatosMes'] < 1: label = "4. Frequência baixa"
        else: label = "5. Recorrência (Já comprou)"
        fila_unificada.append(extrair_dados(row, label))
    
    datas_projecao_uteis = calcular_dias_uteis(hoje, dias_projecao)
    resultado_final = []
    metricas_distribuicao = [] 
    idx_fila = 0
    total_fila = len(fila_unificada)
    
    for data in datas_projecao_uteis:
        ocupados_dia = df_base_agenda[(df_base_agenda['UsuarioAgenda'] == vendedor_agenda) & (df_base_agenda['Sit'] == 'AB') & (df_base_agenda['DataAgenda'].dt.date == data)].shape[0]
        vagas = max(0, capacidade_diaria - ocupados_dia)
        sugeridos_neste_dia = 0
        for _ in range(vagas):
            if idx_fila < total_fila:
                item = fila_unificada[idx_fila].copy()
                item['Data Sugerida'] = data
                if item['DiasSemCompra'] == 9999: item['DiasSemCompra'] = None
                resultado_final.append(item)
                idx_fila += 1
                sugeridos_neste_dia += 1
            else: break
        metricas_distribuicao.append({'Data': data, 'Meta': capacidade_diaria, 'Ocupado': ocupados_dia, 'Sugerido': sugeridos_neste_dia})
            
    df_resultado = pd.DataFrame(resultado_final)
    if df_resultado.empty: return datas_projecao_uteis, df_resultado, pd.DataFrame(metricas_distribuicao), pd.DataFrame([diagnostico]), "Nenhum cliente disponível para agendar."
    return datas_projecao_uteis, df_resultado, pd.DataFrame(metricas_distribuicao), pd.DataFrame([diagnostico]), "Sucesso"


# --- MOCK DE CARREGAMENTO DE DADOS ---

@st.cache_data
def carregar_dados_agenda():
    df_agenda, _ = gerar_dados_ficticios()
    df_agenda['Motivo_Final'] = df_agenda.apply(lambda row: row['Resultado'] if row['Motivo'] == '' else row['Motivo'], axis=1)
    df_agenda['CodClien_str'] = df_agenda['CodClien'].astype(str)
    return df_agenda

@st.cache_data
def carregar_dados_base_cliente():
    _, df_info = gerar_dados_ficticios()
    hoje_dt = pd.to_datetime(datetime.date.today())
    df_info['DiasSemCompra'] = (hoje_dt - df_info['DtUltPed']).dt.days.fillna(9999).astype(int)
    df_info['Cod Vend'] = df_info['Cod Vend'].astype(str)
    return df_info

# --- Inicialização ---
if 'cliente_index' not in st.session_state:
    st.session_state.cliente_index = 0

with st.spinner("Gerando dados fictícios para DEMO..."):
    df_base = carregar_dados_agenda()
    df_info_cliente = carregar_dados_base_cliente()

df_enrich = df_info_cliente[['Cod Clien', 'DiasSemCompra', 'Cod Vend', 'DtUltPed']].drop_duplicates(subset=['Cod Clien'])
df_base = df_base.merge(df_enrich, left_on='CodClien', right_on='Cod Clien', how='left')
df_base['DiasSemCompra'] = df_base['DiasSemCompra'].fillna(0).astype(int)
df_base['Cod Vend'] = df_base['Cod Vend'].fillna('') 

st.toast("Modo Demo Ativado (Dados Fictícios)", icon="🧪")

# --- Barra Lateral (Sidebar) com Filtros ---
st.sidebar.header("Filtros do Dashboard")
st.sidebar.info("Modo Demonstração. Dados gerados aleatoriamente.")

if st.sidebar.button("Gerar Novos Dados"):
    carregar_dados_agenda.clear()
    carregar_dados_base_cliente.clear()
    st.session_state.cliente_index = 0 
    st.rerun()

st.sidebar.markdown("---")
lista_vendedores = sorted(df_base['UsuarioAgenda'].dropna().unique())
vendedores_selecionados = st.sidebar.multiselect("Vendedor (UsuarioEncer/Agenda):", options=lista_vendedores, default=[])
if not vendedores_selecionados: vendedores_selecionados = lista_vendedores

st.sidebar.markdown("---")
lista_vendedores_base_input = st.sidebar.text_area("Lista Cód. Vendedores Base (Enter ou vírgula):")
st.sidebar.markdown("---")
lista_motivos = sorted(df_base['Motivo_Final'].dropna().unique())
motivos_selecionados = st.sidebar.multiselect("Filtrar por Motivo/Resultado:", options=lista_motivos, default=[])

st.sidebar.markdown("---")
st.sidebar.subheader("Filtro de Cliente")
lista_clientes_input = st.sidebar.text_area("Lista Cód. Clientes (Enter ou vírgula):")
filtro_nome_cliente = st.sidebar.text_input("Buscar por Nome do Cliente:")
filtro_obs_texto = st.sidebar.text_input("Filtrar Obs (Contém):")
filtro_notas = st.sidebar.radio("Filtro de Notas:", ["Todos", "Com Notas", "Sem Notas"], horizontal=True)
filtro_dias_sem_compra = st.sidebar.number_input("Dias sem Compra (Mínimo):", min_value=0, value=0, step=10)

st.sidebar.markdown("---")
st.sidebar.subheader("Filtro de Agendamento")
filtro_tipo_agendamento = st.sidebar.radio("Status Agendamento:", ("Todos", "Com Agendamento", "Sem Agendamento", "Data do Agendamento"))
filtro_data_agend_inicio = None
filtro_data_agend_fim = None
if filtro_tipo_agendamento == "Data do Agendamento":
    col_ag1, col_ag2 = st.sidebar.columns(2)
    filtro_data_agend_inicio = col_ag1.date_input("De:", value=datetime.date.today())
    filtro_data_agend_fim = col_ag2.date_input("Até:", value=datetime.date.today() + datetime.timedelta(days=7))

st.sidebar.markdown("---")
st.sidebar.subheader("Filtro de Período (Último Contato)")
data_min_agenda = df_base['DataAgenda'].dropna().min().date()
data_min_encer = df_base['DataEncer'].dropna().min().date()
data_min_original = min(data_min_agenda, data_min_encer) if not pd.isna(data_min_agenda) else datetime.date.today()
data_max_agenda = df_base['DataAgenda'].dropna().max().date()
data_max_encer = df_base['DataEncer'].dropna().max().date()
data_max_original = max(data_max_agenda, data_max_encer) if not pd.isna(data_max_agenda) else datetime.date.today()

filtro_data_inicio = st.sidebar.date_input("Data Inicial:", value=data_min_original)
filtro_data_fim = st.sidebar.date_input("Data Final:", value=data_max_original)

# --- APLICAÇÃO DOS FILTROS ---
df_common_filters = df_base.copy()

if filtro_obs_texto: df_common_filters = df_common_filters[df_common_filters['Obs'].astype(str).str.contains(filtro_obs_texto, case=False, na=False)]
if filtro_notas != 'Todos':
    notas_carregadas_filtro = carregar_notas()
    ids_com_nota = set(notas_carregadas_filtro.keys())
    if filtro_notas == 'Com Notas': df_common_filters = df_common_filters[df_common_filters['CodClien_str'].isin(ids_com_nota)]
    elif filtro_notas == 'Sem Notas': df_common_filters = df_common_filters[~df_common_filters['CodClien_str'].isin(ids_com_nota)]

if vendedores_selecionados: df_common_filters = df_common_filters[(df_common_filters['UsuarioEncer'].isin(vendedores_selecionados)) | (df_common_filters['UsuarioAgenda'].isin(vendedores_selecionados))]
if motivos_selecionados: df_common_filters = df_common_filters[df_common_filters['Motivo_Final'].isin(motivos_selecionados)]
if filtro_dias_sem_compra > 0: df_common_filters = df_common_filters[df_common_filters['DiasSemCompra'] >= filtro_dias_sem_compra]
if filtro_nome_cliente: df_common_filters = df_common_filters[df_common_filters['Cliente'].str.contains(filtro_nome_cliente, case=False, na=False)]
if lista_clientes_input:
    codigos_limpos = [c.strip() for c in lista_clientes_input.replace(',', '\n').split('\n') if c.strip()]
    if codigos_limpos: df_common_filters = df_common_filters[df_common_filters['CodClien_str'].isin(codigos_limpos)]
if lista_vendedores_base_input:
    vends_limpos = [v.strip() for v in lista_vendedores_base_input.replace(',', '\n').split('\n') if v.strip()]
    if vends_limpos: df_common_filters = df_common_filters[df_common_filters['Cod Vend'].isin(vends_limpos)]

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
    df_sched_filtered = df_sched_filtered[(df_sched_filtered['DataAgenda'] >= ts_inicio_ag) & (df_sched_filtered['DataAgenda'] <= ts_fim_ag)]
    df_pairs = df_pairs[(df_pairs['DataProxAgend'] >= ts_inicio_ag) & (df_pairs['DataProxAgend'] <= ts_fim_ag)]

df_common_agend_filtered = df_common_filters.merge(df_pairs[['UsuarioEncer', 'CodClien']], on=['UsuarioEncer', 'CodClien'], how='inner')

ts_inicio_geral = pd.Timestamp(filtro_data_inicio)
ts_fim_geral = pd.Timestamp(filtro_data_fim) + pd.Timedelta(hours=23, minutes=59, seconds=59)

df_filtrado_com_data = df_common_agend_filtered[(df_common_agend_filtered['Sit'] == 'EN') & (df_common_agend_filtered['DataEncer'] >= ts_inicio_geral) & (df_common_agend_filtered['DataEncer'] <= ts_fim_geral)].copy()

df_encerrados_global_sorted = df_base[df_base['Sit'] == 'EN'].sort_values(by='DataEncer', ascending=False)
df_ultimos_contatos_usuario = df_encerrados_global_sorted.drop_duplicates(subset=['CodClien', 'UsuarioEncer'])[['CodClien', 'UsuarioEncer', 'DataEncer', 'Motivo_Final', 'Obs']].rename(columns={'DataEncer': 'Data Ult Contato', 'Motivo_Final': 'Ultimo Motivo', 'Obs': 'Obs Ult Contato', 'UsuarioEncer': 'UsuarioEncer_JoinKey'})
df_ultimos_contatos_usuario['UsuarioEncer_JoinKey'] = df_ultimos_contatos_usuario['UsuarioEncer_JoinKey'].astype(str)

# --- Página Principal ---
st.title("Dashboard de Análise de Contatos (MODO DEMO)")
st.caption("Esta é uma versão de demonstração. Todos os dados são gerados aleatoriamente.")

tab1, tab2, tab3, tab4, tab5 = st.tabs(["📊 Dashboard Principal", "👨‍💼 Detalhe por Cliente", "🗓️ Próximos Agendamentos", "📑 Tabela Resumo", "🤖 Automação de Agenda"])

with tab1:
    df_analise = df_filtrado_com_data
    hoje = datetime.date.today()
    ts_inicio_mes_vigente = pd.Timestamp(hoje.replace(day=1))
    
    df_abertos_kpi = df_base[(df_base['Sit'] == 'AB') & (df_base['UsuarioAgenda'].isin(vendedores_selecionados))]
    if filtro_dias_sem_compra > 0: df_abertos_kpi = df_abertos_kpi[df_abertos_kpi['DiasSemCompra'] >= filtro_dias_sem_compra]
    agendamentos_aberto_total = df_abertos_kpi.shape[0]
    
    if not df_analise.empty:
        contatos_total_filtrado = df_analise.shape[0]
        contatos_mes_vigente_filtrado = df_analise[df_analise['DataEncer'] >= ts_inicio_mes_vigente].shape[0]
        clientes_unicos_filtrados = df_analise['CodClien'].nunique()
        contatos_3_meses_filtrado = df_analise[df_analise['DataEncer'] >= pd.Timestamp((hoje.replace(day=1) - pd.DateOffset(months=2)).date())].shape[0]
    else:
        contatos_total_filtrado = contatos_mes_vigente_filtrado = contatos_3_meses_filtrado = clientes_unicos_filtrados = 0
        
    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Contatos (Filtrado)", f"{contatos_total_filtrado:,}".replace(",", "."))
    col2.metric("Contatos Mês Vigente", f"{contatos_mes_vigente_filtrado:,}".replace(",", "."))
    col3.metric("Contatos 3 Meses", f"{contatos_3_meses_filtrado:,}".replace(",", "."))
    col4.metric("Clientes Únicos", f"{clientes_unicos_filtrados:,}".replace(",", "."))
    col5.metric("Agend. em Aberto (Total)", f"{agendamentos_aberto_total:,}".replace(",", "."))
    
    st.markdown("---")
    
    df_base_resumo = df_base.copy()
    if vendedores_selecionados: df_base_resumo = df_base_resumo[(df_base_resumo['UsuarioEncer'].isin(vendedores_selecionados)) | (df_base_resumo['UsuarioAgenda'].isin(vendedores_selecionados))]
        
    df_encerrados = df_base_resumo[df_base_resumo['Sit'] == 'EN']
    df_abertos = df_base_resumo[df_base_resumo['Sit'] == 'AB']
    
    if filtro_dias_sem_compra > 0:
        df_encerrados = df_encerrados[df_encerrados['DiasSemCompra'] >= filtro_dias_sem_compra]
        df_abertos = df_abertos[df_abertos['DiasSemCompra'] >= filtro_dias_sem_compra]
    
    df_resumo_enc = pd.concat([df_encerrados[df_encerrados['DataEncer'] >= ts_inicio_mes_vigente].groupby('UsuarioEncer').size().rename('Contatos no Mês'), df_encerrados[df_encerrados['DataEncer'].dt.date == hoje].groupby('UsuarioEncer').size().rename('Encerrados Hoje')], axis=1)
    df_resumo_ab = pd.concat([df_abertos[(df_abertos['DataAgenda'].dt.to_period('M') == pd.Period(hoje, 'M'))].groupby('UsuarioAgenda').size().rename('Agend. em Aberto (Mês)'), df_abertos[df_abertos['DataAgenda'].dt.date == hoje].groupby('UsuarioAgenda').size().rename('Agend. para Hoje (Aberto)'), df_abertos[df_abertos['DataAgenda'].dt.date < hoje].groupby('UsuarioAgenda').size().rename('Agend. Atrasados (Total)')], axis=1)
    
    df_resumo_vendedores = pd.concat([df_resumo_enc, df_resumo_ab], axis=1).fillna(0).astype(int)
    df_resumo_vendedores.index.name = "Vendedor"
    df_resumo_vendedores['Total Workload (Mês)'] = df_resumo_vendedores['Contatos no Mês'] + df_resumo_vendedores['Agend. em Aberto (Mês)']
    df_resumo_vendedores['% Eficácia (Mês)'] = (df_resumo_vendedores['Contatos no Mês'].div(df_resumo_vendedores['Total Workload (Mês)'])).fillna(0).map('{:.1%}'.format)
    
    colunas_ordenadas = ['Total Workload (Mês)', 'Contatos no Mês', 'Agend. em Aberto (Mês)', '% Eficácia (Mês)', 'Encerrados Hoje', 'Agend. para Hoje (Aberto)', 'Agend. Atrasados (Total)']
    st.subheader("Resumo de Vendedores (Mês Vigente)")
    st.dataframe(df_resumo_vendedores[colunas_ordenadas].sort_values(by='Total Workload (Mês)', ascending=False), use_container_width=True)
    
    st.markdown("---")
    if df_analise.empty:
        st.warning("Nenhum dado para exibir nos gráficos com os filtros atuais.")
    else:
        col_graf1, col_graf2 = st.columns(2)
        template_graficos = "plotly_dark"
        with col_graf1:
            st.subheader("Top 15 Motivos de Contato")
            fig_motivos = px.bar(df_analise['Motivo_Final'].value_counts().head(15).reset_index().rename(columns={'Motivo_Final':'Motivo', 'count':'Contagem'}), x='Contagem', y='Motivo', orientation='h', text='Contagem', template=template_graficos)
            fig_motivos.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig_motivos, use_container_width=True)
        with col_graf2:
            st.subheader("Resultados por Vendedor")
            fig_vendedores = px.bar(df_analise.groupby('UsuarioEncer').size().reset_index(name='Contagem'), x='UsuarioEncer', y='Contagem', text='Contagem', template=template_graficos)
            fig_vendedores.update_xaxes(tickangle=-45)
            st.plotly_chart(fig_vendedores, use_container_width=True)

        st.markdown("---")
        st.subheader("Mapa de Calor: Intensidade de Atendimentos (Diário)")
        
        df_timeline = df_analise.copy()
        df_timeline['Dia'] = df_timeline['DataEncer'].dt.date
        df_pivot_heatmap = df_timeline.groupby(['Dia', 'UsuarioEncer']).size().reset_index(name='Quantidade').pivot(index='UsuarioEncer', columns='Dia', values='Quantidade').fillna(0)

        fig_heatmap = px.imshow(df_pivot_heatmap, labels=dict(x="Data", y="Vendedor", color="Qtd."), text_auto=True, aspect="auto", color_continuous_scale='Blues', template=template_graficos)

        linhas_verticais = [dict(type="line", x0=i+0.5, x1=i+0.5, y0=0, y1=1, xref="x", yref="paper", line=dict(color="#666666", width=1)) for i, data_col in enumerate(df_pivot_heatmap.columns) if pd.to_datetime(data_col).dayofweek == 4]
        fig_heatmap.update_layout(shapes=linhas_verticais, coloraxis_colorbar=dict(title="Qtd."), xaxis=dict(type='category', title="Data"), margin=dict(r=20, l=20, b=20, t=40))
        fig_heatmap.update_xaxes(tickangle=-45)
        st.plotly_chart(fig_heatmap, use_container_width=True)

with tab2:
    st.subheader("Análise Detalhada de Cliente")
    clientes_com_info = df_filtrado_com_data.drop_duplicates(subset=['CodClien']).merge(df_info_cliente, left_on='CodClien', right_on='Cod Clien', how='left', suffixes=('', '_info'))
    lista_codigos_cliente = clientes_com_info.drop_duplicates(subset=['CodClien']).sort_values(by='Cliente')['CodClien'].tolist() 

    if not lista_codigos_cliente: st.warning("Nenhum cliente encontrado.")
    else:
        total_clientes = len(lista_codigos_cliente)
        if st.session_state.cliente_index >= total_clientes: st.session_state.cliente_index = 0
            
        col_nav1, col_nav2, col_nav3 = st.columns([2, 8, 2])
        if col_nav1.button("⬅️ Cliente Anterior", use_container_width=True):
            st.session_state.cliente_index = st.session_state.cliente_index - 1 if st.session_state.cliente_index > 0 else total_clientes - 1
            st.rerun()
        if col_nav3.button("Cliente Próximo ➡️", use_container_width=True):
            st.session_state.cliente_index = st.session_state.cliente_index + 1 if st.session_state.cliente_index < total_clientes - 1 else 0
            st.rerun()

        current_cod_clien = lista_codigos_cliente[st.session_state.cliente_index]
        info_cliente_atual_df = df_info_cliente[df_info_cliente['Cod Clien'] == current_cod_clien]
        
        if not info_cliente_atual_df.empty:
            cliente_detalhe = info_cliente_atual_df.iloc[0]
            # Verifica se a coluna existe antes de tentar puxar, evitando o KeyError
            if 'Vendedor' in info_cliente_atual_df.columns:
                vendedores_assoc = info_cliente_atual_df['Vendedor'].unique()
            else:
                vendedores_assoc = [f"Vend. Fictício {cliente_detalhe.get('Cod Vend', '')}"]
        else:
            cliente_detalhe = df_base[df_base['CodClien'] == current_cod_clien].iloc[0]
            vendedores_assoc = ["N/D"]

        col_nav2.subheader(f"({current_cod_clien}) - {cliente_detalhe['Cliente']} ({st.session_state.cliente_index + 1}/{total_clientes})")
        
        st.subheader("Histórico Agenda")
        historico_cliente = df_base[df_base['CodClien'] == current_cod_clien].copy()
        
        col_hist1, col_hist2 = st.columns(2)
        with col_hist1: st.caption("Encerrados"); st.dataframe(historico_cliente[historico_cliente['Sit'] == 'EN'][['DataEncer', 'UsuarioEncer', 'Resultado', 'Motivo_Final', 'Obs']].sort_values('DataEncer', ascending=False), use_container_width=True)
        with col_hist2: st.caption("Em Aberto"); st.dataframe(historico_cliente[historico_cliente['Sit'] == 'AB'][['DataAgenda', 'UsuarioAgenda']].sort_values('DataAgenda'), use_container_width=True)

        st.markdown("---")
        with st.container(border=True):
            col_form1, col_form2 = st.columns(2)
            with col_form1:
                st.caption("Segmento"); st.markdown(f"`{cliente_detalhe.get('Segmento', 'N/D')}`")
                st.caption("Área de Venda"); st.markdown(f"`{cliente_detalhe.get('AreaVenda', 'N/D')}`")
                st.caption("Local"); st.markdown(f"`{cliente_detalhe.get('Municipio', '')} - {cliente_detalhe.get('UF', '')}`")
                st.caption("Vendedores Associados"); st.code("\n".join(vendedores_assoc), language=None) 
            with col_form2:
                cond_pgto = cliente_detalhe.get('CondPgto', 'Depósito Antecipado')
                st.caption("Cond. Pagamento"); st.markdown(f"`{cond_pgto if pd.notna(cond_pgto) else 'Depósito Antecipado'}`")
                st.caption("Limite Total"); st.markdown(f"`R$ {cliente_detalhe.get('LimiteTotal', 0):,.2f}`")
                dt_ult_ped = cliente_detalhe.get('DtUltPed', pd.NaT)
                dias_s_compra = cliente_detalhe.get('DiasSemCompra', 0)
                st.caption("Data Últ. Pedido"); st.markdown(f"`{dt_ult_ped.strftime('%d/%m/%Y') if pd.notna(dt_ult_ped) else 'Nenhum'}{' (Nunca comprou)' if dias_s_compra == 9999 else f' ({dias_s_compra} dias)' if pd.notna(dt_ult_ped) else ''}`")
                st.caption("Inadimplência (-3d)"); st.markdown(f"`R$ {cliente_detalhe.get('Inad-3dd', 0):,.2f}`")

        st.markdown("---")
        st.subheader("📝 Notas do Supervisor")
        notas_carregadas = carregar_notas()
        lista_notas_cliente = notas_carregadas.get(str(current_cod_clien), [])
        if isinstance(lista_notas_cliente, str): lista_notas_cliente = [{"id": "old", "data": "Antigo", "texto": lista_notas_cliente, "imagem": None}]

        if lista_notas_cliente:
            for i, nota in enumerate(lista_notas_cliente):
                with st.expander(f"{nota.get('data', 'S/D')} - {nota.get('texto')[:50]}...", expanded=(i==0)):
                    col_txt, col_del = st.columns([0.9, 0.1])
                    with col_txt:
                        st.write(nota.get('texto'))
                        if nota.get('imagem') and os.path.exists(nota.get('imagem')): st.image(nota.get('imagem'), use_column_width=True)
                    with col_del:
                        if st.button("🗑️", key=f"del_{nota.get('id')}_{i}"): excluir_nota(current_cod_clien, nota.get('id')); st.rerun()
        else: st.info("Nenhuma nota registrada para este cliente.")

        with st.form(key='form_nova_nota', clear_on_submit=True):
            texto_nova_nota = st.text_area("Adicionar observação:", height=100)
            uploaded_file = st.file_uploader("Anexar Imagem", type=['png', 'jpg', 'jpeg'])
            if st.form_submit_button("Salvar Nota"):
                if texto_nova_nota.strip(): adicionar_nota(current_cod_clien, texto_nova_nota, uploaded_file); st.rerun()

with tab3:
    st.subheader("Agendamentos em Aberto")
    df_abertos_tab3 = pd.DataFrame() if filtro_tipo_agendamento == "Sem Agendamento" else df_sched_filtered.copy()
    
    if not df_abertos_tab3.empty:
        df_view = pd.merge(df_abertos_tab3, df_ultimos_contatos_usuario[['CodClien', 'UsuarioEncer_JoinKey', 'Ultimo Motivo', 'Data Ult Contato']], left_on=['CodClien', 'UsuarioAgenda'], right_on=['CodClien', 'UsuarioEncer_JoinKey'], how='left')
        df_view_final = pd.merge(df_view, df_info_cliente[['Cod Clien', 'Cod Vend', 'DtUltPed']].rename(columns={'Cod Clien': 'CodClien', 'Cod Vend': 'Cod Vend Base', 'DtUltPed': 'Data Ult Pedido'}), on='CodClien', how='left').sort_values(by=['DataAgenda', 'CodClien']).drop_duplicates(subset=['DataAgenda', 'UsuarioAgenda', 'CodClien'])
        df_view_final['CodClien'] = df_view_final['CodClien'].astype(str)
        st.info(f"Exibindo {len(df_view_final)} agendamentos.")
        
        # Lista de colunas que queremos mostrar
        colunas_desejadas = ['DataAgenda', 'UsuarioAgenda', 'Cliente', 'CodClien', 'Cod Vend Base', 'Data Ult Pedido', 'Ultimo Motivo', 'Data Ult Contato', 'Obs']
        
        # Filtra apenas as colunas que realmente existem no DataFrame no momento
        colunas_disponiveis = [c for c in colunas_desejadas if c in df_view_final.columns]
        
        st.dataframe(df_view_final[colunas_disponiveis], use_container_width=True)

with tab4:
    st.subheader("Resumo por Cliente")
    df_rows_input = df_common_agend_filtered.copy() if df_filtrado_com_data.empty and not df_common_agend_filtered.empty else df_filtrado_com_data
    df_final, _ = calcular_resumo_power_query(df_rows_input, df_base, df_info_cliente)
    
    if not df_final.empty:
        st.dataframe(df_final, use_container_width=True)
        st.download_button("📥 Baixar Resumo (Excel)", data=to_excel(df_final, sheet_name="Resumo"), file_name="resumo.xlsx")
    else: st.warning("Nenhum dado encontrado.")

with tab5:
    st.subheader("🤖 Gerador de Agendas Automático (DEMO)")
    col_sel_vendedor, col_sel_cod_base = st.columns(2)
    with col_sel_vendedor: vendedor_alvo = st.selectbox("Selecione o Vendedor:", options=lista_vendedores)
    with col_sel_cod_base: cod_vend_base_input = st.text_input("Código do Vendedor na BASE:", value=vendedor_alvo)

    with st.expander("🛠️ Configurações Avançadas"):
        col_adv1, col_adv2, col_adv3 = st.columns(3)
        with col_adv1:
            meta_diaria = st.number_input("Capacidade Diária:", value=15, min_value=1)
            dias_projecao = st.number_input("Dias úteis:", value=5, min_value=1)
            filtro_incluir_ja_atendidos = st.checkbox("Incluir contatos anteriores", value=True)
            filtro_limite_min = st.number_input("Limite Mínimo (R$):", value=0.0)
            filtro_apenas_adimplentes = st.checkbox("Apenas Adimplentes", value=True)
        with col_adv2:
            filtro_incluir_grupo_colig = st.checkbox("Incluir Grupo/Coligação", value=True)
            filtro_segmentos = st.multiselect("Segmento:", options=sorted(df_info_cliente['Segmento'].unique()))
            filtro_municipios = st.multiselect("Município:", options=sorted(df_info_cliente['Municipio'].unique()))
        with col_adv3:
            filtro_sit_cred = st.multiselect("Sit. Crédito:", options=sorted(df_info_cliente['SitCred'].unique()))
            filtro_dias_sem_compra_min = st.number_input("Dias Sem Compra (Mín):", value=0)

    if st.button("Gerar Sugestão de Agenda 🚀"):
        filtros_dict = {
            "cod_vend_base": cod_vend_base_input, "municipios": filtro_municipios, "segmentos": filtro_segmentos,
            "areas": [], "sit_cred": filtro_sit_cred, "cond_pgto": [], "limite_min": filtro_limite_min,
            "dias_sem_compra_min": filtro_dias_sem_compra_min, "apenas_adimplentes": filtro_apenas_adimplentes,
            "filtro_notas": filtro_notas, "apenas_com_prazo": False, "incluir_grupo_colig": filtro_incluir_grupo_colig, 
            "incluir_ja_atendidos": filtro_incluir_ja_atendidos
        }
        
        with st.spinner("Processando..."):
            datas_uteis, df_sugestao, df_metricas, df_diagnostico, msg = gerar_sugestao_agenda(df_base, df_info_cliente, vendedor_alvo, meta_diaria, dias_projecao, filtros_dict)
        
        if df_sugestao.empty: st.warning(msg)
        else:
            st.success(f"Agenda gerada! {len(df_sugestao)} clientes distribuídos.")
            st.dataframe(df_diagnostico, use_container_width=True)
            
            for data_util in datas_uteis:
                if data_util in df_sugestao.groupby('Data Sugerida').groups:
                    grupo = df_sugestao.groupby('Data Sugerida').get_group(data_util)
                    st.markdown(f"**{data_util.strftime('%d/%m/%Y')} - {len(grupo)} Clientes**")
                    st.dataframe(grupo, use_container_width=True, hide_index=True)
