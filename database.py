# database.py — Versão FINAL completa com conexão dinâmica e campo de retificação
import fdb
import hashlib
import os
import sys
from datetime import datetime
from PyQt5.QtCore import QSettings # Importa a classe para ler as configurações

#==============================================================================
# FUNÇÕES DE CONEXÃO E INICIALIZAÇÃO
#==============================================================================

def conectar():
    """
    Conecta ao banco de dados Firebird lendo as configurações salvas pelo usuário
    (local ou remoto). Inclui tratamento de erro aprimorado.
    """
    settings = QSettings()
    
    # --- Carrega as configurações salvas ---
    modo = settings.value("database/modo", "local")
    user = settings.value("database/usuario", "SYSDBA")
    password = settings.value("database/senha", "masterkey")
    
    host = ""
    port = 0
    database_path = ""

    if modo == "local":
        host = "localhost"
        port = 3050 # Porta padrão para Firebird local

        # --- INÍCIO DA CORREÇÃO ---
        
        # 1. Tenta carregar o caminho salvo nas configurações pelo usuário
        database_path = settings.value("database/caminho_local", "")

        # 2. Se nenhum caminho foi salvo (string vazia), usa a lógica antiga como PADRÃO.
        #    Isso mantém o comportamento de criar um banco automático na primeira vez.
        if not database_path:
            print(" Nenhum caminho de banco de dados local configurado. Usando caminho padrão.")
            # Determina o caminho base (onde o executável ou script está)
            if getattr(sys, 'frozen', False):  # Se for executável compilado (PyInstaller)
                base_path = os.path.dirname(sys.executable)
            else:  # Se estiver rodando como script .py
                base_path = os.path.dirname(os.path.abspath(__file__))

            # Constrói o caminho para a pasta 'Data'
            data_folder = os.path.join(base_path, 'Data')

            # Cria a pasta 'Data' se ela não existir
            os.makedirs(data_folder, exist_ok=True)

            # Define o caminho completo e final para o arquivo do banco de dados
            database_path = os.path.join(data_folder, 'CALENDARIO.FDB')
        
        # 3. A lógica de criação do banco de dados agora usará o caminho correto
        #    (seja o que veio das configurações ou o padrão)
        if not os.path.exists(database_path):
            print(f"🆕 Criando banco de dados local em: {database_path}")
            fdb.create_database(dsn=f"{host}/{port}:{database_path}", user=user, password=password)
        
        # --- FIM DA CORREÇÃO ---

    else: # modo == "remoto"
        host = settings.value("database/host_remoto", "localhost")
        port = settings.value("database/porta_remota", 3050, type=int)
        database_path = settings.value("database/caminho_remoto", "")
        if not host or not database_path:
            raise ConnectionError("Configuração remota incompleta: Host ou Caminho do banco não definido.")

    # --- Tenta conectar com as configurações carregadas ---
    print(f"Conectando ao banco ({modo}): {host}:{port}/{database_path}")
    try:
        return fdb.connect(
            host=host,
            port=port,
            database=database_path,
            user=user,
            password=password,
            charset='UTF8'
        )
    except fdb.Error as e:
        print(f"❌ Erro crítico ao conectar ao banco de dados: {e}")
        raise

def tabela_existe(cur, nome_tabela):
    cur.execute("SELECT RDB$RELATION_NAME FROM RDB$RELATIONS WHERE RDB$RELATION_NAME = ?", (nome_tabela.upper(),))
    return cur.fetchone() is not None

def criar_generator_e_trigger(cur, tabela):
    gen_name = f"GEN_{tabela}_ID"
    trg_name = f"TRG_{tabela}_BI"
    cur.execute(f"SELECT RDB$GENERATOR_NAME FROM RDB$GENERATORS WHERE RDB$GENERATOR_NAME = '{gen_name}'")
    if cur.fetchone() is None:
        cur.execute(f'CREATE GENERATOR {gen_name}')
    
    cur.execute(f"SELECT RDB$TRIGGER_NAME FROM RDB$TRIGGERS WHERE RDB$TRIGGER_NAME = '{trg_name}'")
    if cur.fetchone() is None:
        cur.execute(f"""
        CREATE TRIGGER {trg_name} FOR {tabela}
        ACTIVE BEFORE INSERT POSITION 0
        AS
        BEGIN
            IF (NEW.ID IS NULL) THEN NEW.ID = GEN_ID({gen_name}, 1);
        END
        """)

def iniciar_db():
    conn = conectar()
    cur = conn.cursor()
    
    # --- Tabelas ---
    if not tabela_existe(cur, 'USUARIOS'):
        cur.execute("CREATE TABLE USUARIOS (ID INTEGER NOT NULL PRIMARY KEY, USERNAME VARCHAR(50) UNIQUE NOT NULL, PASSWORD_HASH VARCHAR(64) NOT NULL)")
        criar_generator_e_trigger(cur, 'USUARIOS')
    if not tabela_existe(cur, 'STATUS'):
        cur.execute("CREATE TABLE STATUS (ID INTEGER NOT NULL PRIMARY KEY, NOME VARCHAR(50) UNIQUE NOT NULL, COR_HEX VARCHAR(10) NOT NULL)")
        criar_generator_e_trigger(cur, 'STATUS')
    if not tabela_existe(cur, 'CLIENTES'):
        cur.execute("CREATE TABLE CLIENTES (ID INTEGER NOT NULL PRIMARY KEY, NOME VARCHAR(150) NOT NULL, TIPO_ENVIO VARCHAR(100) NOT NULL, CONTATO VARCHAR(100) NOT NULL, GERA_RECIBO SMALLINT DEFAULT 0, CONTA_XMLS SMALLINT DEFAULT 0, NIVEL VARCHAR(20), OUTROS_DETALHES BLOB SUB_TYPE TEXT, NUMERO_COMPUTADORES INTEGER DEFAULT 0)")
        criar_generator_e_trigger(cur, 'CLIENTES')
    if not tabela_existe(cur, 'ENTREGAS'):
        # ADICIONADO O CAMPO IS_RETIFICACAO
        cur.execute("CREATE TABLE ENTREGAS (ID INTEGER NOT NULL PRIMARY KEY, DATA_VENCIMENTO DATE NOT NULL, HORARIO VARCHAR(10) NOT NULL, STATUS_ID INTEGER, CLIENTE_ID INTEGER NOT NULL, RESPONSAVEL VARCHAR(50), OBSERVACOES BLOB SUB_TYPE TEXT, IS_RETIFICACAO SMALLINT DEFAULT 0, FOREIGN KEY (STATUS_ID) REFERENCES STATUS (ID) ON DELETE SET NULL, FOREIGN KEY (CLIENTE_ID) REFERENCES CLIENTES (ID) ON DELETE CASCADE)")
        criar_generator_e_trigger(cur, 'ENTREGAS')

    else:
        try:
            cur.execute("""
                SELECT RDB$FIELD_NAME FROM RDB$RELATION_FIELDS
                WHERE RDB$RELATION_NAME = 'ENTREGAS' AND RDB$FIELD_NAME = 'IS_RETIFICACAO'
            """)
            if cur.fetchone() is None:
                print("🩵 Adicionando campo IS_RETIFICACAO na tabela ENTREGAS...")
                cur.execute("ALTER TABLE ENTREGAS ADD IS_RETIFICACAO SMALLINT DEFAULT 0;")
                conn.commit()
        except Exception as e:
            print(f"⚠️ Erro ao verificar/adicionar campo IS_RETIFICACAO: {e}")    
    if not tabela_existe(cur, 'LOGS'):
        cur.execute("CREATE TABLE LOGS (ID INTEGER NOT NULL PRIMARY KEY, DATAHORA TIMESTAMP DEFAULT CURRENT_TIMESTAMP, USUARIO_NOME VARCHAR(50), ACAO VARCHAR(50), DETALHES BLOB SUB_TYPE TEXT)")
        criar_generator_e_trigger(cur, 'LOGS')
    conn.commit()

    # --- Dados Padrão ---
    cur.execute("SELECT COUNT(*) FROM USUARIOS")
    if cur.fetchone()[0] == 0:
        senha_hash = hashlib.sha256('admin'.encode('utf-8')).hexdigest()
        cur.execute("INSERT INTO USUARIOS (USERNAME, PASSWORD_HASH) VALUES (?, ?)", ('admin', senha_hash))
    cur.execute("SELECT COUNT(*) FROM STATUS")
    if cur.fetchone()[0] == 0:
        status_padrao = [('Pendente', '#ffc107'), ('Feito e enviado', '#28a745'), ('Feito', '#007bff'), ('Retificado', '#17a2b8'), ('Houve Algum Erro', '#dc3545'), ('Chamado', '#6f42c1'), ('Remarcado', '#fd7e14'), ('Realocado', '#6c757d')]
        cur.executemany("INSERT INTO STATUS (NOME, COR_HEX) VALUES (?, ?)", status_padrao)
    conn.commit()
    conn.close()
    print("✅ Banco Firebird 2.5 inicializado com sucesso!")

#==============================================================================
# FUNÇÃO AUXILIAR
#==============================================================================
def dict_factory(cursor, row):
    """
    Converte uma tupla de resultado do Firebird em um dicionário.
    Ajustado para maior robustez.
    """
    if not row:
        return None
    
    d = {}
    try:
        for idx, col in enumerate(cursor.description):
            # Acessa o nome da coluna de forma segura
            col_name = col[0]
            if isinstance(col_name, bytes):
                col_name = col_name.decode('utf-8', errors='ignore')
            
            d[col_name.upper()] = row[idx]
    except Exception as e:
        print(f"Erro ao processar linha do banco de dados na dict_factory: {e}")
        return None # Retorna None se houver erro na conversão da linha

    return d

#==============================================================================
# LOGS
#==============================================================================
def registrar_log(usuario_nome, acao, detalhes):
    sql = "INSERT INTO LOGS (USUARIO_NOME, ACAO, DETALHES) VALUES (?, ?, ?)"
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (usuario_nome, acao, detalhes))
        conn.commit()
    except fdb.Error as e:
        print(f"Erro ao registrar log: {e}")
    finally:
        if conn: conn.close()

#==============================================================================
# USUÁRIOS
#==============================================================================
def verificar_usuario(username, password):
    senha_hash = hashlib.sha256(password.encode('utf-8')).hexdigest()
    sql = "SELECT ID, USERNAME FROM USUARIOS WHERE USERNAME = ? AND PASSWORD_HASH = ?"
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (username.strip(), senha_hash))
        return dict_factory(cur, cur.fetchone())
    finally:
        if conn: conn.close()

def listar_usuarios():
    sql = "SELECT USERNAME FROM USUARIOS ORDER BY USERNAME"
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql)
        return [row[0] for row in cur.fetchall()]
    finally:
        if conn: conn.close()
        
def get_usuario_por_nome(username):
    sql = "SELECT ID, USERNAME FROM USUARIOS WHERE USERNAME = ?"
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (username,))
        return dict_factory(cur, cur.fetchone())
    finally:
        if conn: conn.close()

def verificar_senha_usuario_atual(username, password):
    usuario = verificar_usuario(username, password)
    return usuario is not None

def criar_usuario(username, password, usuario_logado):
    if get_usuario_por_nome(username): return False
    senha_hash = hashlib.sha256(password.encode('utf-8')).hexdigest()
    sql = "INSERT INTO USUARIOS (USERNAME, PASSWORD_HASH) VALUES (?, ?)"
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (username, senha_hash))
        conn.commit()
        registrar_log(usuario_logado, "CRIAR_USUARIO", f"Usuário '{username}' criado.")
        return True
    finally:
        if conn: conn.close()

def atualizar_usuario(user_id, novo_username, nova_senha, usuario_logado):
    senha_hash = hashlib.sha256(nova_senha.encode('utf-8')).hexdigest()
    sql = "UPDATE USUARIOS SET USERNAME = ?, PASSWORD_HASH = ? WHERE ID = ?"
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (novo_username, senha_hash, user_id))
        conn.commit()
        registrar_log(usuario_logado, "ATUALIZAR_USUARIO", f"Usuário ID {user_id} atualizado para '{novo_username}'.")
    finally:
        if conn: conn.close()

def deletar_usuario(user_id, usuario_logado):
    sql = "DELETE FROM USUARIOS WHERE ID = ?"
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (user_id,))
        conn.commit()
        registrar_log(usuario_logado, "DELETAR_USUARIO", f"Usuário ID {user_id} excluído.")
        return True
    except fdb.Error:
        return False
    finally:
        if conn: conn.close()

#==============================================================================
# CLIENTES
#==============================================================================
def get_total_clientes():
    sql = "SELECT COUNT(*) FROM CLIENTES"
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql)
        return cur.fetchone()[0]
    finally:
        if conn: conn.close()

def listar_clientes():
    sql = "SELECT * FROM CLIENTES ORDER BY NOME"
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql)
        return [dict_factory(cur, row) for row in cur.fetchall()]
    finally:
        if conn: conn.close()

def adicionar_cliente(nome, tipo_envio, contato, gera_recibo, conta_xmls, nivel, detalhes, numero_computadores, usuario_logado):
    sql = "INSERT INTO CLIENTES (NOME, TIPO_ENVIO, CONTATO, GERA_RECIBO, CONTA_XMLS, NIVEL, OUTROS_DETALHES, NUMERO_COMPUTADORES) VALUES (?, ?, ?, ?, ?, ?, ?, ?)"
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (nome, tipo_envio, contato, 1 if gera_recibo else 0, 1 if conta_xmls else 0, nivel, detalhes, numero_computadores))
        conn.commit()
        registrar_log(usuario_logado, "CRIAR_CLIENTE", f"Cliente '{nome}' adicionado.")
    finally:
        if conn: conn.close()

def atualizar_cliente(cliente_id, nome, tipo_envio, contato, gera_recibo, conta_xmls, nivel, detalhes, numero_computadores, usuario_logado):
    sql = "UPDATE CLIENTES SET NOME=?, TIPO_ENVIO=?, CONTATO=?, GERA_RECIBO=?, CONTA_XMLS=?, NIVEL=?, OUTROS_DETALHES=?, NUMERO_COMPUTADORES=? WHERE ID=?"
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (nome, tipo_envio, contato, 1 if gera_recibo else 0, 1 if conta_xmls else 0, nivel, detalhes, numero_computadores, cliente_id))
        conn.commit()
        registrar_log(usuario_logado, "ATUALIZAR_CLIENTE", f"Cliente '{nome}' (ID: {cliente_id}) atualizado.")
    finally:
        if conn: conn.close()

def deletar_cliente(cliente_id, usuario_logado):
    sql = "DELETE FROM CLIENTES WHERE ID = ?"
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (cliente_id,))
        conn.commit()
        registrar_log(usuario_logado, "DELETAR_CLIENTE", f"Cliente ID {cliente_id} excluído.")
    finally:
        if conn: conn.close()

#==============================================================================
# STATUS
#==============================================================================
def listar_status():
    sql = "SELECT ID, NOME, COR_HEX FROM STATUS ORDER BY ID"
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql)
        return [dict_factory(cur, row) for row in cur.fetchall()]
    finally:
        if conn: conn.close()

def adicionar_status(nome, cor_hex, usuario_logado):
    sql = "INSERT INTO STATUS (NOME, COR_HEX) VALUES (?, ?)"
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (nome, cor_hex))
        conn.commit()
        registrar_log(usuario_logado, "CRIAR_STATUS", f"Status '{nome}' criado.")
    finally:
        if conn: conn.close()

def atualizar_status(status_id, nome, cor_hex, usuario_logado):
    sql = "UPDATE STATUS SET NOME = ?, COR_HEX = ? WHERE ID = ?"
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (nome, cor_hex, status_id))
        conn.commit()
        registrar_log(usuario_logado, "ATUALIZAR_STATUS", f"Status '{nome}' (ID: {status_id}) atualizado.")
    finally:
        if conn: conn.close()

def deletar_status(status_id, usuario_logado):
    sql = "DELETE FROM STATUS WHERE ID = ?"
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (status_id,))
        conn.commit()
        registrar_log(usuario_logado, "DELETAR_STATUS", f"Status ID {status_id} excluído.")
    finally:
        if conn: conn.close()

#==============================================================================
# ENTREGAS / AGENDAMENTOS
#==============================================================================
def adicionar_entrega(data, horario, status_id, cliente_id, responsavel, observacoes, is_retificacao, usuario_logado):
    sql = "INSERT INTO ENTREGAS (DATA_VENCIMENTO, HORARIO, STATUS_ID, CLIENTE_ID, RESPONSAVEL, OBSERVACOES, IS_RETIFICACAO) VALUES (?, ?, ?, ?, ?, ?, ?)"
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (data, horario, status_id, cliente_id, responsavel, observacoes, 1 if is_retificacao else 0))
        conn.commit()
        registrar_log(usuario_logado, "CRIAR_AGENDAMENTO", f"Agendamento para cliente ID {cliente_id} em {data} às {horario}.")
    finally:
        if conn: conn.close()

def atualizar_entrega(entrega_id, horario, status_id, cliente_id, responsavel, observacoes, is_retificacao, usuario_logado):
    sql = "UPDATE ENTREGAS SET HORARIO=?, STATUS_ID=?, CLIENTE_ID=?, RESPONSAVEL=?, OBSERVACOES=?, IS_RETIFICACAO=? WHERE ID=?"
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (horario, status_id, cliente_id, responsavel, observacoes, 1 if is_retificacao else 0, entrega_id))
        conn.commit()
        registrar_log(usuario_logado, "ATUALIZAR_AGENDAMENTO", f"Agendamento ID {entrega_id} atualizado.")
    finally:
        if conn: conn.close()

def deletar_entrega(entrega_id, usuario_logado):
    sql = "DELETE FROM ENTREGAS WHERE ID = ?"
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (entrega_id,))
        conn.commit()
        registrar_log(usuario_logado, "DELETAR_AGENDAMENTO", f"Agendamento ID {entrega_id} excluído.")
    finally:
        if conn: conn.close()

def get_entregas_por_dia(data_str):
    # ADICIONADO e.IS_RETIFICACAO AO SELECT
    sql = """
        SELECT
            e.ID, e.DATA_VENCIMENTO, e.HORARIO, e.STATUS_ID, e.CLIENTE_ID, e.RESPONSAVEL, e.OBSERVACOES, e.IS_RETIFICACAO,
            c.NOME as NOME_CLIENTE, c.CONTATO, c.TIPO_ENVIO, c.NUMERO_COMPUTADORES,
            s.NOME as NOME_STATUS, s.COR_HEX
        FROM ENTREGAS e
        JOIN CLIENTES c ON e.CLIENTE_ID = c.ID
        LEFT JOIN STATUS s ON e.STATUS_ID = s.ID
        WHERE e.DATA_VENCIMENTO = ?
    """
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (data_str,))
        entregas = {}
        for row in cur.fetchall():
            entrega_dict = dict_factory(cur, row)
            if entrega_dict:
                entregas[entrega_dict['HORARIO']] = entrega_dict
        return entregas
    finally:
        if conn: conn.close()
        
#==============================================================================
# FUNÇÕES PARA O DASHBOARD E RELATÓRIOS
#==============================================================================
def get_estatisticas_mensais(ano, mes):
    sql_concluidos = """
        SELECT COUNT(*) FROM ENTREGAS e
        JOIN STATUS s ON e.STATUS_ID = s.ID
        WHERE UPPER(s.NOME) LIKE '%FEITO%' AND EXTRACT(YEAR FROM e.DATA_VENCIMENTO) = ? AND EXTRACT(MONTH FROM e.DATA_VENCIMENTO) = ?
    """
    sql_retificados = """
        SELECT COUNT(*) FROM ENTREGAS e
        JOIN STATUS s ON e.STATUS_ID = s.ID
        WHERE UPPER(s.NOME) LIKE '%RETIFICADO%' AND EXTRACT(YEAR FROM e.DATA_VENCIMENTO) = ? AND EXTRACT(MONTH FROM e.DATA_VENCIMENTO) = ?
    """
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql_concluidos, (ano, mes))
        concluidos = cur.fetchone()[0]
        cur.execute(sql_retificados, (ano, mes))
        retificados = cur.fetchone()[0]
        return {'CONCLUIDOS': concluidos, 'RETIFICADOS': retificados}
    finally:
        if conn: conn.close()

def get_clientes_com_agendamento_no_mes(ano, mes):
    sql = "SELECT DISTINCT CLIENTE_ID FROM ENTREGAS WHERE EXTRACT(YEAR FROM DATA_VENCIMENTO) = ? AND EXTRACT(MONTH FROM DATA_VENCIMENTO) = ?"
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (ano, mes))
        return {row[0] for row in cur.fetchall()}
    finally:
        if conn: conn.close()

def get_clientes_com_agendamento_concluido_no_mes(ano, mes):
    sql = """
        SELECT DISTINCT e.CLIENTE_ID FROM ENTREGAS e
        JOIN STATUS s ON e.STATUS_ID = s.ID
        WHERE UPPER(s.NOME) LIKE '%FEITO%' AND EXTRACT(YEAR FROM e.DATA_VENCimento) = ? AND EXTRACT(MONTH FROM e.DATA_VENCIMENTO) = ?
    """
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (ano, mes))
        return {row[0] for row in cur.fetchall()}
    finally:
        if conn: conn.close()

def get_status_dias_para_mes(ano, mes):
    sql = """
        SELECT
            EXTRACT(DAY FROM e.DATA_VENCIMENTO),
            COUNT(e.ID),
            (SELECT FIRST 1 s.COR_HEX FROM ENTREGAS e2 JOIN STATUS s ON e2.STATUS_ID = s.ID WHERE e2.DATA_VENCIMENTO = e.DATA_VENCIMENTO ORDER BY s.ID)
        FROM ENTREGAS e
        WHERE EXTRACT(YEAR FROM e.DATA_VENCIMENTO) = ? AND EXTRACT(MONTH FROM e.DATA_VENCIMENTO) = ?
        GROUP BY e.DATA_VENCIMENTO
    """
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (ano, mes))
        status_dias = {}
        for row in cur.fetchall():
            dia, contagem, cor = row
            status_dias[int(dia)] = {'CONTAGEM': contagem, 'COR': cor}
        return status_dias
    finally:
        if conn: conn.close()

def get_entregas_no_intervalo(data, hora_inicio, hora_fim):
    sql = """
        SELECT e.ID, e.HORARIO, c.NOME AS NOME_CLIENTE, s.NOME AS NOME_STATUS
        FROM ENTREGAS e
        JOIN CLIENTES c ON e.CLIENTE_ID = c.ID
        LEFT JOIN STATUS s ON e.STATUS_ID = s.ID
        WHERE e.DATA_VENCIMENTO = ? AND e.HORARIO >= ? AND e.HORARIO < ?
    """
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (data, hora_inicio, hora_fim))
        return [dict_factory(cur, row) for row in cur.fetchall()]
    finally:
        if conn: conn.close()

def get_entregas_filtradas(data_inicio, data_fim, status_ids):
    base_sql = """
        SELECT e.DATA_VENCIMENTO, e.HORARIO, c.NOME AS NOME_CLIENTE, s.NOME AS NOME_STATUS, e.RESPONSAVEL, c.CONTATO, c.TIPO_ENVIO, e.OBSERVACOES
        FROM ENTREGAS e
        JOIN CLIENTES c ON e.CLIENTE_ID = c.ID
        LEFT JOIN STATUS s ON e.STATUS_ID = s.ID
        WHERE e.DATA_VENCIMENTO BETWEEN ? AND ?
    """
    params = [data_inicio, data_fim]
    if status_ids:
        placeholders = ', '.join(['?' for _ in status_ids])
        base_sql += f" AND s.ID IN ({placeholders})"
        params.extend(status_ids)
    
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(base_sql, params)
        return [dict_factory(cur, row) for row in cur.fetchall()]
    finally:
        if conn: conn.close()

def get_logs_filtrados(data_inicio, data_fim, usuario):
    base_sql = "SELECT DATAHORA as timestamp, USUARIO_NOME as usuario_nome, ACAO as acao, DETALHES as detalhes FROM LOGS WHERE DATAHORA BETWEEN ? AND ?"
    params = [data_inicio, data_fim]
    if usuario != "Todos":
        base_sql += " AND USUARIO_NOME = ?"
        params.append(usuario)

    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(base_sql, params)
        return [dict_factory(cur, row) for row in cur.fetchall()]
    finally:
        if conn: conn.close()

def get_estatisticas_por_usuario_e_status():
    sql = """
        SELECT
            EXTRACT(MONTH FROM e.DATA_VENCIMENTO) || '/' || EXTRACT(YEAR FROM e.DATA_VENCIMENTO) as MES,
            e.RESPONSAVEL,
            s.NOME as NOME_STATUS,
            COUNT(*) as CONTAGEM
        FROM ENTREGAS e
        JOIN STATUS s ON e.STATUS_ID = s.ID
        WHERE e.RESPONSAVEL IS NOT NULL AND e.RESPONSAVEL <> ''
        GROUP BY MES, e.RESPONSAVEL, s.NOME
    """
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql)
        return [dict_factory(cur, row) for row in cur.fetchall()]
    finally:
        if conn: conn.close()