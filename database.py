# database.py — Versão FINAL completa com conexão dinâmica e campo de retificação
import fdb
import hashlib
import os
import sys
from datetime import datetime
from PyQt5.QtCore import QSettings,QDate # Importa a classe para ler as configurações

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
    user = settings.value("database/usuario", "sysdba")
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

def coluna_existe(cur, nome_tabela, nome_coluna):
    """Verifica se uma coluna existe em uma tabela."""
    cur.execute("""
        SELECT 1
        FROM RDB$RELATION_FIELDS
        WHERE RDB$RELATION_NAME = ? AND RDB$FIELD_NAME = ?
    """, (nome_tabela.upper(), nome_coluna.upper()))
    return cur.fetchone() is not None

def iniciar_db():
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        
        # --- ETAPA 1: CRIAÇÃO E ATUALIZAÇÃO DE TABELAS ---

        # Tabela USUARIOS
        if not tabela_existe(cur, 'USUARIOS'):
            print("🆕 Criando tabela USUARIOS...")
            cur.execute("CREATE TABLE USUARIOS (ID INTEGER NOT NULL PRIMARY KEY, USERNAME VARCHAR(50) UNIQUE NOT NULL, PASSWORD_HASH VARCHAR(64) NOT NULL, IS_ADMIN SMALLINT DEFAULT 0 NOT NULL)")
            criar_generator_e_trigger(cur, 'USUARIOS')
        
        if not coluna_existe(cur, 'USUARIOS', 'IS_ADMIN'):
            print("🔧 Atualizando tabela USUARIOS: Adicionando coluna IS_ADMIN...")
            cur.execute("ALTER TABLE USUARIOS ADD IS_ADMIN SMALLINT DEFAULT 0 NOT NULL")
            conn.commit()

        # Tabela STATUS (sem alterações)
        if not tabela_existe(cur, 'STATUS'):
            print("🆕 Criando tabela STATUS...")
            cur.execute("CREATE TABLE STATUS (ID INTEGER NOT NULL PRIMARY KEY, NOME VARCHAR(50) UNIQUE NOT NULL, COR_HEX VARCHAR(10) NOT NULL)")
            criar_generator_e_trigger(cur, 'STATUS')

        # Tabela CLIENTES (sem alterações)
        if not tabela_existe(cur, 'CLIENTES'):
            cur.execute("CREATE TABLE CLIENTES (ID INTEGER NOT NULL PRIMARY KEY, NOME VARCHAR(150) NOT NULL, TIPO_ENVIO VARCHAR(100) NOT NULL, CONTATO VARCHAR(100) NOT NULL, GERA_RECIBO SMALLINT DEFAULT 0, CONTA_XMLS SMALLINT DEFAULT 0, NIVEL VARCHAR(20), OUTROS_DETALHES BLOB SUB_TYPE TEXT, NUMERO_COMPUTADORES INTEGER DEFAULT 0)")
            criar_generator_e_trigger(cur, 'CLIENTES')

        # Tabela ENTREGAS
        if not tabela_existe(cur, 'ENTREGAS'):
            cur.execute("""CREATE TABLE ENTREGAS (ID INTEGER NOT NULL PRIMARY KEY,DATA_VENCIMENTO DATE NOT NULL,HORARIO VARCHAR(10) NOT NULL,STATUS_ID INTEGER,CLIENTE_ID INTEGER NOT NULL,RESPONSAVEL VARCHAR(150),OBSERVACOES BLOB SUB_TYPE TEXT,IS_RETIFICACAO SMALLINT DEFAULT 0,TIPO_ATENDIMENTO VARCHAR(20) DEFAULT 'AGENDADO' NOT NULL,FOREIGN KEY (STATUS_ID) REFERENCES STATUS (ID) ON DELETE SET NULL,FOREIGN KEY (CLIENTE_ID) REFERENCES CLIENTES (ID) ON DELETE CASCADE)""")
            criar_generator_e_trigger(cur, 'ENTREGAS')

        # --- Verificação para atualizar a tabela ENTREGAS de bancos antigos ---
        if not coluna_existe(cur, 'ENTREGAS', 'TIPO_ATENDIMENTO'):
            print("🔧 Atualizando tabela ENTREGAS: Adicionando coluna TIPO_ATENDIMENTO...")
            cur.execute("ALTER TABLE ENTREGAS ADD TIPO_ATENDIMENTO VARCHAR(20) DEFAULT 'AGENDADO' NOT NULL")
            conn.commit() # Aplica a alteração da estrutura

        if not coluna_existe(cur, 'ENTREGAS', 'DATA_CONCLUSAO'):
            print("🔧 Atualizando tabela ENTREGAS: Adicionando coluna DATA_CONCLUSAO...")
            cur.execute("ALTER TABLE ENTREGAS ADD DATA_CONCLUSAO TIMESTAMP")
            conn.commit()

        # Tabela LOGS 
        if not tabela_existe(cur, 'LOGS'):
            cur.execute("CREATE TABLE LOGS (ID INTEGER NOT NULL PRIMARY KEY, DATAHORA TIMESTAMP DEFAULT CURRENT_TIMESTAMP, USUARIO_NOME VARCHAR(50), ACAO VARCHAR(50), DETALHES BLOB SUB_TYPE TEXT)")
            criar_generator_e_trigger(cur, 'LOGS')

        # Tabela FERIADOS 
        if not tabela_existe(cur, 'FERIADOS'):
            print("📅 Criando tabela de Feriados...")
            cur.execute("CREATE TABLE FERIADOS (ID INTEGER NOT NULL PRIMARY KEY, DATA DATE NOT NULL UNIQUE, TIPO VARCHAR(20) NOT NULL)")
            criar_generator_e_trigger(cur, 'FERIADOS')

        conn.commit() # Salva todas as alterações de estrutura

        # --- ETAPA 2: INSERÇÃO DE DADOS PADRÃO ---
        
        # Insere o usuário 'admin' se a tabela estiver vazia
        cur.execute("SELECT COUNT(*) FROM USUARIOS")
        if cur.fetchone()[0] == 0:
            print("👤 Inserindo usuário 'admin' padrão...")
            senha_hash = hashlib.sha256('admin'.encode('utf-8')).hexdigest()
            cur.execute("INSERT INTO USUARIOS (USERNAME, PASSWORD_HASH, IS_ADMIN) VALUES (?, ?, ?)", ('admin', senha_hash, 1))
            
        # Insere os status padrão se a tabela estiver vazia
        cur.execute("SELECT COUNT(*) FROM STATUS")
        if cur.fetchone()[0] == 0:
            print("🎨 Inserindo status padrão...")
            status_padrao = [('Pendente', '#ffc107'), ('Feito e enviado', '#28a745'), ('Feito', '#007bff'), ('Retificado', '#17a2b8'), ('Houve Algum Erro', '#dc3545'), ('Chamado', '#6f42c1'), ('Remarcado', '#fd7e14'), ('Realocado', '#6c757d')]
            cur.executemany("INSERT INTO STATUS (NOME, COR_HEX) VALUES (?, ?)", status_padrao)
        
        conn.commit()
        print("✅ Banco de dados verificado e inicializado com sucesso!")

    except fdb.Error as e:
        print(f"❌ Erro CRÍTICO durante a inicialização do banco de dados: {e}")
        if conn:
            conn.rollback()
        raise
    
    finally:
        if conn:
            conn.close()

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

def adicionar_feriado(data_str, tipo):
    """Adiciona ou atualiza um feriado no banco de dados."""
    # Tenta remover primeiro para evitar duplicatas e permitir a troca de tipo (ex: de municipal para nacional)
    remover_feriado(data_str)
    sql = "INSERT INTO FERIADOS (DATA, TIPO) VALUES (?, ?)"
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (data_str, tipo))
        conn.commit()
    except fdb.Error as e:
        print(f"Erro ao adicionar feriado: {e}")
    finally:
        if conn: conn.close()

def remover_feriado(data_str):
    """Remove um feriado do banco de dados."""
    sql = "DELETE FROM FERIADOS WHERE DATA = ?"
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (data_str,))
        conn.commit()
    finally:
        if conn: conn.close()

def get_feriados_do_mes(ano, mes):
    """Busca todos os feriados de um determinado mês e ano."""
    sql = "SELECT DATA, TIPO FROM FERIADOS WHERE EXTRACT(YEAR FROM DATA) = ? AND EXTRACT(MONTH FROM DATA) = ?"
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (ano, mes))
        # Retorna um dicionário no formato {QDate: 'tipo'}
        return {QDate(row[0].year, row[0].month, row[0].day): row[1] for row in cur.fetchall()}
    finally:
        if conn: conn.close()

def is_dia_invalido(data_qdate):
    """
    Função central que verifica se uma data é fim de semana OU feriado nacional.
    Retorna True se for um dia inválido para agendamento, False caso contrário.
    """
    # 1. Verifica se é Sábado (6) ou Domingo (7)
    if data_qdate.dayOfWeek() >= 6:
        return True
    
    # 2. Verifica no banco de dados se é um feriado do tipo 'nacional'
    sql = "SELECT ID FROM FERIADOS WHERE DATA = ? AND TIPO = 'nacional'"
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (data_qdate.toString("yyyy-MM-dd"),))
        if cur.fetchone():
            return True # Encontrou um feriado nacional
    finally:
        if conn: conn.close()
        
    return False # Se não for fim de semana nem feriado, é um dia útil

#==============================================================================
# USUÁRIOS
#==============================================================================
def verificar_usuario(username, password):
    senha_hash = hashlib.sha256(password.encode('utf-8')).hexdigest()
    # Adicionamos IS_ADMIN ao SELECT
    sql = "SELECT ID, USERNAME, IS_ADMIN FROM USUARIOS WHERE USERNAME = ? AND PASSWORD_HASH = ?" # <-- ALTERADO
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
    # Adicionamos IS_ADMIN ao SELECT
    sql = "SELECT ID, USERNAME, IS_ADMIN FROM USUARIOS WHERE USERNAME = ?" # <-- ALTERADO
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

def criar_usuario(username, password, is_admin, usuario_logado): # <-- NOVO PARÂMETRO
    if get_usuario_por_nome(username): return False
    senha_hash = hashlib.sha256(password.encode('utf-8')).hexdigest()
    # Adicionamos IS_ADMIN ao INSERT
    sql = "INSERT INTO USUARIOS (USERNAME, PASSWORD_HASH, IS_ADMIN) VALUES (?, ?, ?)" # <-- ALTERADO
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (username, senha_hash, 1 if is_admin else 0)) # <-- ALTERADO
        conn.commit()
        detalhes = f"Usuário '{username}' criado."
        if is_admin:
            detalhes += " (Como Administrador)"
        registrar_log(usuario_logado, "CRIAR_USUARIO", detalhes)
        return True
    finally:
        if conn: conn.close()

def atualizar_usuario(user_id, novo_username, nova_senha, is_admin, usuario_logado): # <-- NOVO PARÂMETRO
    senha_hash = hashlib.sha256(nova_senha.encode('utf-8')).hexdigest()
    # Adicionamos IS_ADMIN ao UPDATE
    sql = "UPDATE USUARIOS SET USERNAME = ?, PASSWORD_HASH = ?, IS_ADMIN = ? WHERE ID = ?" # <-- ALTERADO
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (novo_username, senha_hash, 1 if is_admin else 0, user_id)) # <-- ALTERADO
        conn.commit()
        detalhes = f"Usuário ID {user_id} atualizado para '{novo_username}'."
        if is_admin:
            detalhes += " (Status de Administrador Concedido)"
        registrar_log(usuario_logado, "ATUALIZAR_USUARIO", detalhes)
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

def get_admin_count():
    """
    Conta e retorna o número total de usuários que são administradores.
    """
    sql = "SELECT COUNT(ID) FROM USUARIOS WHERE IS_ADMIN = 1"
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql)
        # Retorna o primeiro (e único) resultado da contagem
        return cur.fetchone()[0]
    except fdb.Error as e:
        print(f"Erro ao contar administradores: {e}")
        return 0 # Em caso de erro, retorna 0 para segurança
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

def adicionar_cliente(nome, tipo_envio, contato, gera_recibo, conta_xmls, nivel, detalhes, numero_computadores, telefone1, telefone2, usuario_logado):
    sql = "INSERT INTO CLIENTES (NOME, TIPO_ENVIO, CONTATO, GERA_RECIBO, CONTA_XMLS, NIVEL, OUTROS_DETALHES, NUMERO_COMPUTADORES, TELEFONE1, TELEFONE2) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (nome, tipo_envio, contato, 1 if gera_recibo else 0, 1 if conta_xmls else 0, nivel, detalhes, numero_computadores, telefone1, telefone2))
        conn.commit()
        registrar_log(usuario_logado, "CRIAR_CLIENTE", f"Cliente '{nome}' adicionado.")
    finally:
        if conn: conn.close()

def atualizar_cliente(cliente_id, nome, tipo_envio, contato, gera_recibo, conta_xmls, nivel, detalhes, numero_computadores, telefone1, telefone2, usuario_logado):
    sql = "UPDATE CLIENTES SET NOME=?, TIPO_ENVIO=?, CONTATO=?, GERA_RECIBO=?, CONTA_XMLS=?, NIVEL=?, OUTROS_DETALHES=?, NUMERO_COMPUTADORES=?, TELEFONE1=?, TELEFONE2=? WHERE ID=?"
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (nome, tipo_envio, contato, 1 if gera_recibo else 0, 1 if conta_xmls else 0, nivel, detalhes, numero_computadores, telefone1, telefone2, cliente_id))
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
def adicionar_entrega(data, horario, status_id, cliente_id, responsavel, observacoes, is_retificacao, usuario_logado, tipo_atendimento='AGENDADO'):
    sql = "INSERT INTO ENTREGAS (DATA_VENCIMENTO, HORARIO, STATUS_ID, CLIENTE_ID, RESPONSAVEL, OBSERVACOES, IS_RETIFICACAO, TIPO_ATENDIMENTO) VALUES (?, ?, ?, ?, ?, ?, ?, ?)"
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (data, horario, status_id, cliente_id, responsavel, observacoes, 1 if is_retificacao else 0, tipo_atendimento))
        conn.commit()
        registrar_log(usuario_logado, "CRIAR_AGENDAMENTO", f"Agendamento para cliente ID {cliente_id} em {data} às {horario}.")
    finally:
        if conn: conn.close()

def atualizar_entrega(entrega_id, horario, status_id, cliente_id, responsavel, observacoes, is_retificacao, usuario_logado, tipo_atendimento):
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()

        # Descobre o nome do novo status para a lógica de conclusão
        cur.execute("SELECT NOME FROM STATUS WHERE ID = ?", (status_id,))
        status_result = cur.fetchone()
        status_nome = status_result[0].lower() if status_result else ""

        # Verifica se o status é um dos que marcam a tarefa como concluída
        is_status_concluido = 'feito' in status_nome or 'retificado' in status_nome

        # Monta a query SQL dinamicamente para atualizar a data de conclusão
        if is_status_concluido:
            # Se o status for de conclusão, define a data/hora atuais
            sql = "UPDATE ENTREGAS SET HORARIO=?, STATUS_ID=?, CLIENTE_ID=?, RESPONSAVEL=?, OBSERVACOES=?, IS_RETIFICACAO=?, TIPO_ATENDIMENTO=?, DATA_CONCLUSAO = CURRENT_TIMESTAMP WHERE ID=?"
        else:
            # Se for outro status, limpa a data de conclusão (define como nulo)
            sql = "UPDATE ENTREGAS SET HORARIO=?, STATUS_ID=?, CLIENTE_ID=?, RESPONSAVEL=?, OBSERVACOES=?, IS_RETIFICACAO=?, TIPO_ATENDIMENTO=?, DATA_CONCLUSAO = NULL WHERE ID=?"

        # Executa a atualização
        cur.execute(sql, (horario, status_id, cliente_id, responsavel, observacoes, 1 if is_retificacao else 0, tipo_atendimento, entrega_id))
        conn.commit()

        # Registra o log da operação
        registrar_log(usuario_logado, "ATUALIZAR_AGENDAMENTO", f"Agendamento ID {entrega_id} atualizado.")
    finally:
        if conn:
            conn.close()

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
    sql = """
        SELECT
            e.ID, e.DATA_VENCIMENTO, e.HORARIO, e.STATUS_ID, e.CLIENTE_ID,
            e.RESPONSAVEL, e.OBSERVACOES, e.IS_RETIFICACAO, e.DATA_CONCLUSAO, -- <--- COLUNA ADICIONADA AQUI
            c.NOME as NOME_CLIENTE, c.CONTATO, c.TIPO_ENVIO, c.NUMERO_COMPUTADORES,
            s.NOME as NOME_STATUS, s.COR_HEX
        FROM ENTREGAS e
        JOIN CLIENTES c ON e.CLIENTE_ID = c.ID
        LEFT JOIN STATUS s ON e.STATUS_ID = s.ID
        WHERE e.DATA_VENCIMENTO = ? AND e.TIPO_ATENDIMENTO = 'AGENDADO'
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

def get_solicitados_do_mes(ano, mes):
    sql = """
        SELECT
            e.ID, e.DATA_VENCIMENTO, e.HORARIO, e.STATUS_ID, e.CLIENTE_ID, e.RESPONSAVEL, e.OBSERVACOES, e.IS_RETIFICACAO, e.TIPO_ATENDIMENTO,
            c.NOME as NOME_CLIENTE,
            s.NOME as NOME_STATUS, s.COR_HEX
        FROM ENTREGAS e
        JOIN CLIENTES c ON e.CLIENTE_ID = c.ID
        LEFT JOIN STATUS s ON e.STATUS_ID = s.ID
        WHERE EXTRACT(YEAR FROM e.DATA_VENCIMENTO) = ? AND EXTRACT(MONTH FROM e.DATA_VENCIMENTO) = ?
        AND e.TIPO_ATENDIMENTO = 'SOLICITADO'
        ORDER BY e.DATA_VENCIMENTO DESC, e.HORARIO DESC
    """
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (ano, mes))
        return [dict_factory(cur, row) for row in cur.fetchall()]
    finally:
        if conn: conn.close()

def limpar_agendamentos_futuros_pendentes(cliente_id, usuario_logado):
    """
    Exclui todos os agendamentos FUTUROS com status 'Pendente' para um cliente específico.
    Isso evita a duplicação ao atualizar uma regra de recorrência.
    """
    sql_status = "SELECT ID FROM STATUS WHERE UPPER(NOME) = 'PENDENTE'"
    sql_delete = "DELETE FROM ENTREGAS WHERE CLIENTE_ID = ? AND STATUS_ID = ? AND DATA_VENCIMENTO > CURRENT_DATE"
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        
        cur.execute(sql_status)
        resultado_status = cur.fetchone()
        if not resultado_status:
            print("⚠️ Status 'Pendente' não encontrado. Não foi possível limpar agendamentos.")
            return

        status_pendente_id = resultado_status[0]
        
        cur.execute(sql_delete, (cliente_id, status_pendente_id))
        conn.commit()
        
        if cur.rowcount > 0:
            registrar_log(usuario_logado, "LIMPEZA_RECORRENCIA", f"{cur.rowcount} agendamentos futuros pendentes do cliente ID {cliente_id} foram removidos.")
            
    except fdb.Error as e:
        print(f"Erro ao limpar agendamentos futuros: {e}")
    finally:
        if conn: conn.close()

def criar_agendamentos_recorrentes(agendamentos_para_criar, usuario_logado):
    """
    Recebe uma lista de agendamentos já validados (com data, hora, etc.) e os insere no banco.
    """
    if not agendamentos_para_criar:
        return
        
    cliente_id = agendamentos_para_criar[0]['cliente_id']
    
    # Primeiro, limpa os agendamentos pendentes futuros para este cliente
    limpar_agendamentos_futuros_pendentes(cliente_id, usuario_logado)

    conn = conectar()
    cur = conn.cursor()
    
    cur.execute("SELECT ID FROM STATUS WHERE UPPER(NOME) = 'PENDENTE'")
    status_pendente_id = cur.fetchone()[0]

    for agendamento in agendamentos_para_criar:
        adicionar_entrega(
            data=agendamento['data'],
            horario=agendamento['hora'],
            status_id=status_pendente_id,
            cliente_id=cliente_id,
            responsavel=usuario_logado,
            observacoes=agendamento['obs'],
            is_retificacao=False,
            usuario_logado=usuario_logado
        )

    conn.close()
    registrar_log(usuario_logado, "CRIAR_RECORRENCIA", f"Criados {len(agendamentos_para_criar)} agendamentos recorrentes para o cliente ID {cliente_id}.")

def limpar_agendamentos_futuros_cliente(cliente_id, usuario_logado):
    """
    Exclui todos os agendamentos de um cliente a partir da data atual (inclusive).
    É usado para limpar a agenda de um cliente sem afetar o histórico.
    """
    # A condição DATA_VENCIMENTO >= CURRENT_DATE garante que apenas agendamentos
    # de hoje em diante sejam removidos.
    sql = "DELETE FROM ENTREGAS WHERE CLIENTE_ID = ? AND DATA_VENCIMENTO >= CURRENT_DATE"
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (cliente_id,))
        conn.commit()
        
        # Registra no log quantos agendamentos foram removidos para auditoria
        if cur.rowcount > 0:
            registrar_log(usuario_logado, "LIMPEZA_AGENDAMENTOS", f"{cur.rowcount} agendamentos futuros do cliente ID {cliente_id} foram removidos.")
        return cur.rowcount # Retorna o número de linhas afetadas
            
    except fdb.Error as e:
        print(f"Erro ao limpar agendamentos futuros do cliente: {e}")
        if conn:
            conn.rollback()
        return 0
    finally:
        if conn: conn.close()

def get_contagem_solicitados_do_mes(ano, mes):
    """
    Busca apenas a CONTAGEM de atendimentos do tipo 'SOLICITADO' para um mês/ano.
    É mais rápido do que buscar todos os dados.
    """
    sql = """
        SELECT COUNT(ID)
        FROM ENTREGAS
        WHERE EXTRACT(YEAR FROM DATA_VENCIMENTO) = ?
        AND EXTRACT(MONTH FROM DATA_VENCIMENTO) = ?
        AND TIPO_ATENDIMENTO = 'SOLICITADO'
    """
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (ano, mes))
        # cur.fetchone()[0] pega o primeiro (e único) resultado da contagem.
        return cur.fetchone()[0]
    except fdb.Error as e:
        print(f"Erro ao contar atendimentos solicitados: {e}")
        return 0
    finally:
        if conn: conn.close()

def get_contagem_solicitados_pendentes_do_mes(ano, mes):
    """
    Busca a CONTAGEM de atendimentos 'SOLICITADO' que estão com o status 'Pendente'.
    """
    sql = """
        SELECT COUNT(e.ID)
        FROM ENTREGAS e
        JOIN STATUS s ON e.STATUS_ID = s.ID
        WHERE EXTRACT(YEAR FROM e.DATA_VENCIMENTO) = ?
        AND EXTRACT(MONTH FROM e.DATA_VENCIMENTO) = ?
        AND e.TIPO_ATENDIMENTO = 'SOLICITADO'
        AND UPPER(s.NOME) = 'PENDENTE'
    """
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (ano, mes))
        # cur.fetchone()[0] pega o primeiro (e único) resultado da contagem.
        return cur.fetchone()[0]
    except fdb.Error as e:
        print(f"Erro ao contar atendimentos solicitados pendentes: {e}")
        return 0
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
    # A MUDANÇA ESTÁ AQUI: Usamos CAST(DATAHORA AS DATE) para comparar apenas as datas
    base_sql = "SELECT DATAHORA as data_hora, USUARIO_NOME as usuario_nome, ACAO as acao, DETALHES as detalhes FROM LOGS WHERE CAST(DATAHORA AS DATE) BETWEEN ? AND ?"
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

def buscar_agendamentos_globais(termo_busca):
    """
    Busca agendamentos em todo o banco de dados com base em um termo.
    A busca é case-insensitive e procura no nome do cliente, responsável e observações.
    """
    sql = """
        SELECT
            e.ID, e.DATA_VENCIMENTO, e.HORARIO, e.RESPONSAVEL,
            c.NOME as NOME_CLIENTE,
            s.NOME as NOME_STATUS
        FROM ENTREGAS e
        JOIN CLIENTES c ON e.CLIENTE_ID = c.ID
        LEFT JOIN STATUS s ON e.STATUS_ID = s.ID
        WHERE
            UPPER(c.NOME) LIKE ? OR
            UPPER(e.RESPONSAVEL) LIKE ? OR
            UPPER(e.OBSERVACOES) LIKE ?
        ORDER BY e.DATA_VENCIMENTO DESC
    """
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        # Adiciona os wildcards '%' para buscar o termo em qualquer parte do texto
        termo_like = f"%{termo_busca.upper()}%"
        cur.execute(sql, (termo_like, termo_like, termo_like))
        return [dict_factory(cur, row) for row in cur.fetchall()]
    finally:
        if conn: conn.close()

def get_status_de_atividade_clientes():
    """
    Busca todos os clientes e a data do último agendamento de cada um.
    Retorna uma lista de dicionários contendo o nome do cliente, contato,
    e a data do último agendamento (ou None se nunca agendou).
    """
    sql = """
        SELECT
            c.ID,
            c.NOME,
            c.CONTATO,
            MAX(e.DATA_VENCIMENTO) as ULTIMO_AGENDAMENTO
        FROM
            CLIENTES c
        LEFT JOIN
            ENTREGAS e ON c.ID = e.CLIENTE_ID
        GROUP BY
            c.ID, c.NOME, c.CONTATO
        ORDER BY
            ULTIMO_AGENDAMENTO ASC, c.NOME ASC
    """
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql)
        return [dict_factory(cur, row) for row in cur.fetchall()]
    finally:
        if conn: conn.close()


def get_estatisticas_cliente_periodo(cliente_id, data_inicio, data_fim):
    """
    Retorna a contagem de cada status para um cliente específico em um período.
    """
    sql = """
        SELECT
            s.NOME,
            COUNT(e.ID) as CONTAGEM
        FROM ENTREGAS e
        JOIN STATUS s ON e.STATUS_ID = s.ID
        WHERE e.CLIENTE_ID = ? AND e.DATA_VENCIMENTO BETWEEN ? AND ?
        GROUP BY s.NOME
        ORDER BY CONTAGEM DESC
    """
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (cliente_id, data_inicio, data_fim))
        # Transforma o resultado em um dicionário: {'Status': contagem}
        return {row[0]: row[1] for row in cur.fetchall()}
    finally:
        if conn: conn.close()


def get_dados_ranking_clientes_periodo(data_inicio, data_fim):
    """
    Calcula as contagens de status chave para TODOS os clientes em um período,
    para ser usado na montagem dos rankings.
    """
    sql = """
        SELECT
            c.NOME,
            SUM(CASE WHEN s.NOME IN ('Feito', 'Feito e enviado') THEN 1 ELSE 0 END) as CONCLUIDOS,
            SUM(CASE WHEN s.NOME = 'Retificado' THEN 1 ELSE 0 END) as RETIFICADOS,
            SUM(CASE WHEN s.NOME = 'Remarcado' THEN 1 ELSE 0 END) as REMARCADOS, -- <-- MUDANÇA AQUI
            SUM(CASE WHEN s.NOME = 'Houve Algum Erro' THEN 1 ELSE 0 END) as ERROS
        FROM
            CLIENTES c
        JOIN
            ENTREGAS e ON c.ID = e.CLIENTE_ID
        JOIN
            STATUS s ON e.STATUS_ID = s.ID
        WHERE
            e.DATA_VENCIMENTO BETWEEN ? AND ?
        GROUP BY
            c.NOME
    """
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (data_inicio, data_fim))
        return [dict_factory(cur, row) for row in cur.fetchall()]
    finally:
        if conn: conn.close()

def get_cliente_por_id(cliente_id):
    sql = "SELECT * FROM CLIENTES WHERE ID = ?"
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (cliente_id,))
        return dict_factory(cur, cur.fetchone())
    finally:
        if conn: conn.close()

def verificar_agendamento_pendente_existente(cliente_id):
    """
    Verifica se um cliente já possui um agendamento com status 'Pendente'.
    Retorna os dados do agendamento se encontrado, caso contrário, None.
    """
    sql = """
        SELECT FIRST 1 e.ID, e.DATA_VENCIMENTO, e.HORARIO
        FROM ENTREGAS e
        JOIN STATUS s ON e.STATUS_ID = s.ID
        WHERE e.CLIENTE_ID = ? AND UPPER(s.NOME) = 'PENDENTE'
    """
    conn = None
    try:
        conn = conectar()
        cur = conn.cursor()
        cur.execute(sql, (cliente_id,))
        return dict_factory(cur, cur.fetchone())
    finally:
        if conn: conn.close()