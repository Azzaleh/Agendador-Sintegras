# database.py
import sqlite3
import hashlib
from datetime import datetime, timezone, timedelta

def conectar():
    """Conecta ao banco de dados e habilita o acesso por nome de coluna."""
    conn = sqlite3.connect('calendario.db')
    conn.row_factory = sqlite3.Row
    return conn

def iniciar_db():
    """Cria/atualiza as tabelas do banco de dados."""
    conn = conectar()
    cursor = conn.cursor()
    
    cursor.execute('CREATE TABLE IF NOT EXISTS usuarios (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT NOT NULL UNIQUE, password_hash TEXT NOT NULL)')
    cursor.execute('CREATE TABLE IF NOT EXISTS status (id INTEGER PRIMARY KEY AUTOINCREMENT, nome TEXT NOT NULL UNIQUE, cor_hex TEXT NOT NULL)')
    cursor.execute('''CREATE TABLE IF NOT EXISTS clientes (id INTEGER PRIMARY KEY AUTOINCREMENT, nome TEXT NOT NULL, tipo_envio TEXT NOT NULL, contato TEXT NOT NULL, gera_recibo BOOLEAN NOT NULL DEFAULT 0, conta_xmls BOOLEAN NOT NULL DEFAULT 0, nivel TEXT, outros_detalhes TEXT, numero_computadores INTEGER DEFAULT 0)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS entregas (id INTEGER PRIMARY KEY AUTOINCREMENT, data_vencimento TEXT NOT NULL, horario TEXT NOT NULL, status_id INTEGER, cliente_id INTEGER NOT NULL, responsavel TEXT, observacoes TEXT, FOREIGN KEY (status_id) REFERENCES status (id) ON DELETE SET NULL, FOREIGN KEY (cliente_id) REFERENCES clientes (id) ON DELETE CASCADE)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS logs (id INTEGER PRIMARY KEY AUTOINCREMENT, timestamp DATETIME DEFAULT CURRENT_TIMESTAMP, usuario_nome TEXT NOT NULL, acao TEXT NOT NULL, detalhes TEXT)''')
    
    cursor.execute("SELECT COUNT(id) FROM usuarios")
    if cursor.fetchone()[0] == 0:
        senha_hash = hashlib.sha256('admin'.encode('utf-8')).hexdigest()
        cursor.execute("INSERT INTO usuarios (username, password_hash) VALUES (?, ?)", ('admin', senha_hash))

    cursor.execute("SELECT COUNT(id) FROM status")
    if cursor.fetchone()[0] == 0:
        status_padrao = [
            ('PENDENTE', '#ffc107'), ('Feito e enviado', '#28a745'), ('Feito', '#007bff'), 
            ('Retificado', '#17a2b8'), ('Houve Algum Erro', '#dc3545'), ('Chamado', '#6f42c1'), 
            ('Remarcado', '#fd7e14'), ('Realocado', '#6c757d')
        ]
        cursor.executemany("INSERT INTO status (nome, cor_hex) VALUES (?, ?)", status_padrao)

    conn.commit()
    conn.close()

def registrar_log(usuario_nome, acao, detalhes=""):
    conn = conectar()
    conn.execute("INSERT INTO logs (usuario_nome, acao, detalhes) VALUES (?, ?, ?)", (usuario_nome, acao, detalhes))
    conn.commit()
    conn.close()

def criar_usuario(username, password):
    conn = conectar()
    senha_hash = hashlib.sha256(password.encode('utf-8')).hexdigest()
    try:
        conn.execute("INSERT INTO usuarios (username, password_hash) VALUES (?, ?)", (username, senha_hash))
        conn.commit()
        registrar_log('sistema', 'USUARIO_CRIADO', f'Usuário: {username}')
        return True
    except sqlite3.IntegrityError:
        return False
    finally:
        conn.close()

def verificar_usuario(username, password):
    conn = conectar()
    senha_hash = hashlib.sha256(password.encode('utf-8')).hexdigest()
    cursor = conn.execute("SELECT * FROM usuarios WHERE username = ? AND password_hash = ?", (username, senha_hash))
    usuario = cursor.fetchone()
    conn.close()
    if usuario:
        return dict(usuario)
    return None

def listar_usuarios():
    conn = conectar()
    usuarios = conn.execute("SELECT username FROM usuarios ORDER BY username").fetchall()
    conn.close()
    return [row['username'] for row in usuarios]

def deletar_usuario(id, usuario_logado):
    """ Deleta um usuário do sistema, com proteção para o ID 1. """
    # --- NOVA CAMADA DE SEGURANÇA ---
    # Impede que o usuário com ID 1 (admin principal) seja deletado.
    if id == 1:
        registrar_log(usuario_logado, 'EXCLUSAO_NEGADA', f'Tentativa de excluir usuário admin principal (ID: {id})')
        return False # Retorna False para indicar que a operação falhou

    conn = conectar()
    # Pega o nome para o log antes de deletar
    cursor = conn.execute("SELECT username FROM usuarios WHERE id = ?", (id,))
    usuario = cursor.fetchone()
    nome_usuario = usuario['username'] if usuario else f"ID {id}"
    
    conn.execute("DELETE FROM usuarios WHERE id = ?", (id,))
    conn.commit()
    conn.close()
    
    registrar_log(usuario_logado, 'USUARIO_DELETADO', f'Usuário: {nome_usuario}')
    return True # Retorna True para indicar sucesso

# database.py

def adicionar_cliente(nome, tipo_envio, contato, gera_recibo, conta_xmls, nivel, detalhes, numero_computadores, usuario_logado):
    conn = conectar()
    conn.execute(
        "INSERT INTO clientes (nome, tipo_envio, contato, gera_recibo, conta_xmls, nivel, outros_detalhes, numero_computadores) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
        (nome, tipo_envio, contato, gera_recibo, conta_xmls, nivel, detalhes, numero_computadores)
    )
    conn.commit()
    conn.close()
    registrar_log(usuario_logado, "CLIENTE_CRIADO", f"Cliente: {nome}")

def atualizar_cliente(id, nome, tipo_envio, contato, gera_recibo, conta_xmls, nivel, detalhes, numero_computadores, usuario_logado):
    conn = conectar()
    conn.execute(
        "UPDATE clientes SET nome=?, tipo_envio=?, contato=?, gera_recibo=?, conta_xmls=?, nivel=?, outros_detalhes=?, numero_computadores=? WHERE id=?",
        (nome, tipo_envio, contato, gera_recibo, conta_xmls, nivel, detalhes, numero_computadores, id)
    )
    conn.commit()
    conn.close()
    registrar_log(usuario_logado, "CLIENTE_ATUALIZADO", f"Cliente ID: {id}, Nome: {nome}")

def listar_clientes():
    conn = conectar()
    clientes = conn.execute("SELECT * FROM clientes ORDER BY nome").fetchall()
    conn.close()
    return clientes

def deletar_cliente(id, usuario_logado):
    conn = conectar()
    cursor = conn.execute("SELECT nome FROM clientes WHERE id = ?", (id,))
    cliente = cursor.fetchone()
    nome_cliente = cliente['nome'] if cliente else f"ID {id}"
    conn.execute("DELETE FROM clientes WHERE id = ?", (id,))
    conn.commit()
    conn.close()
    registrar_log(usuario_logado, "CLIENTE_DELETADO", f"Cliente: {nome_cliente}")

def listar_status():
    conn = conectar()
    status_list = conn.execute("SELECT * FROM status ORDER BY nome").fetchall()
    conn.close()
    return status_list

def adicionar_status(nome, cor_hex, usuario_logado):
    conn = conectar()
    conn.execute("INSERT INTO status (nome, cor_hex) VALUES (?, ?)", (nome, cor_hex))
    conn.commit()
    conn.close()
    registrar_log(usuario_logado, "STATUS_CRIADO", f"Status: {nome}")

def atualizar_status(id, nome, cor_hex, usuario_logado):
    conn = conectar()
    conn.execute("UPDATE status SET nome=?, cor_hex=? WHERE id=?", (nome, cor_hex, id))
    conn.commit()
    conn.close()
    registrar_log(usuario_logado, "STATUS_ATUALIZADO", f"Status ID: {id}, Nome: {nome}")

def deletar_status(id, usuario_logado):
    conn = conectar()
    cursor = conn.execute("SELECT nome FROM status WHERE id = ?", (id,))
    status = cursor.fetchone()
    nome_status = status['nome'] if status else f"ID {id}"
    conn.execute("DELETE FROM status WHERE id = ?", (id,))
    conn.commit()
    conn.close()
    registrar_log(usuario_logado, "STATUS_DELETADO", f"Status: {nome_status}")

def adicionar_entrega(data_vencimento, horario, status_id, cliente_id, responsavel, observacoes, usuario_logado):
    conn = conectar()
    conn.execute("INSERT INTO entregas (data_vencimento, horario, status_id, cliente_id, responsavel, observacoes) VALUES (?, ?, ?, ?, ?, ?)", (data_vencimento, horario, status_id, cliente_id, responsavel, observacoes))
    conn.commit()
    conn.close()
    registrar_log(usuario_logado, "AGENDAMENTO_CRIADO", f"Data: {data_vencimento}, Hora: {horario}, Cliente ID: {cliente_id}")

# database.py

def atualizar_entrega(id, horario, status_id, cliente_id, responsavel, observacoes, usuario_logado):
    conn = conectar()
    # Usaremos um cursor para poder buscar dados
    cursor = conn.cursor()

    # 1. ANTES de atualizar, buscamos os dados antigos no banco
    #    Usamos JOINs para já pegar os nomes do cliente e do status
    cursor.execute("""
        SELECT e.horario, e.status_id, e.cliente_id, e.responsavel, e.observacoes,
               c.nome as nome_cliente, s.nome as nome_status
        FROM entregas e
        LEFT JOIN clientes c ON e.cliente_id = c.id
        LEFT JOIN status s ON e.status_id = s.id
        WHERE e.id = ?
    """, (id,))
    dados_antigos = cursor.fetchone()

    # Se por algum motivo o agendamento não for encontrado, encerramos aqui
    if not dados_antigos:
        conn.close()
        return

    # 2. Agora, realizamos o UPDATE no banco de dados com os novos dados
    cursor.execute("UPDATE entregas SET horario=?, status_id=?, cliente_id=?, responsavel=?, observacoes=? WHERE id=?",
                   (horario, status_id, cliente_id, responsavel, observacoes, id))

    # 3. Comparamos os dados antigos com os novos para montar os detalhes do log
    detalhes_log = []

    # Compara o Status (a sua solicitação principal)
    if dados_antigos['status_id'] != status_id:
        nome_status_antigo = dados_antigos['nome_status'] if dados_antigos['nome_status'] else "Nenhum"
        # Busca o nome do novo status
        cursor.execute("SELECT nome FROM status WHERE id = ?", (status_id,))
        res_status = cursor.fetchone()
        nome_status_novo = res_status['nome'] if res_status else "Nenhum"
        detalhes_log.append(f"Status: de '{nome_status_antigo}' para '{nome_status_novo}'")

    # BÔNUS: Vamos fazer o mesmo para outros campos importantes
    if dados_antigos['cliente_id'] != cliente_id:
        nome_cliente_antigo = dados_antigos['nome_cliente']
        cursor.execute("SELECT nome FROM clientes WHERE id = ?", (cliente_id,))
        res_cliente = cursor.fetchone()
        nome_cliente_novo = res_cliente['nome'] if res_cliente else "N/A"
        detalhes_log.append(f"Cliente: de '{nome_cliente_antigo}' para '{nome_cliente_novo}'")

    if dados_antigos['horario'] != horario:
        detalhes_log.append(f"Horário: de '{dados_antigos['horario']}' para '{horario}'")

    if dados_antigos['responsavel'] != responsavel:
        detalhes_log.append(f"Responsável: de '{dados_antigos['responsavel']}' para '{responsavel}'")

    if dados_antigos['observacoes'] != observacoes:
        detalhes_log.append("Observações foram alteradas.")

    # 4. Monta a string final para o log
    if not detalhes_log:
        detalhes_finais = f"Agendamento ID {id} salvo sem alterações."
    else:
        # Junta todas as alterações encontradas, separadas por "; "
        detalhes_finais = f"Agendamento ID {id}: " + "; ".join(detalhes_log)

    # 5. Salva as alterações no banco e fecha a conexão
    conn.commit()
    conn.close()
    
    # 6. Registra o log com a nossa nova mensagem detalhada
    registrar_log(usuario_logado, "AGENDAMENTO_ATUALIZADO", detalhes_finais)

def deletar_entrega(id, usuario_logado):
    conn = conectar()
    conn.execute("DELETE FROM entregas WHERE id = ?", (id,))
    conn.commit()
    conn.close()
    registrar_log(usuario_logado, "AGENDAMENTO_DELETADO", f"Agendamento ID: {id}")

def get_entregas_por_dia(data):
    conn = conectar()
    query = "SELECT e.id, e.data_vencimento, e.horario, e.responsavel, e.observacoes, e.cliente_id, e.status_id, c.nome as nome_cliente, c.tipo_envio, c.contato, c.numero_computadores, s.nome as nome_status, s.cor_hex FROM entregas e JOIN clientes c ON e.cliente_id = c.id LEFT JOIN status s ON e.status_id = s.id WHERE e.data_vencimento = ? "
    entregas_dia = conn.execute(query, (data,)).fetchall()
    conn.close()
    return {entrega['horario']: dict(entrega) for entrega in entregas_dia}

def get_status_dias_para_mes(ano, mes):
    data_inicio = f"{ano}-{mes:02d}-01"; data_fim = f"{ano}-{mes:02d}-31"; conn = conectar()
    query = "SELECT e.data_vencimento, s.nome as nome_status, s.cor_hex FROM entregas e LEFT JOIN status s ON e.status_id = s.id WHERE e.data_vencimento BETWEEN ? AND ?"
    entregas_mes = conn.execute(query, (data_inicio, data_fim)).fetchall(); conn.close()
    ordem_prioridade = ['Houve Algum Erro', 'Chamado', 'Remarcado', 'PENDENTE', 'Realocado', 'Retificado', 'Feito', 'Feito e enviado']
    status_por_dia = {}
    for entrega in entregas_mes:
        dia = int(entrega['data_vencimento'].split('-')[2])
        if dia not in status_por_dia: status_por_dia[dia] = []
        status_por_dia[dia].append(entrega['nome_status'])
    resultado_final = {}
    for dia, status_lista in status_por_dia.items():
        cor_final = '#28a745'; status_final = 'Feito e enviado'
        todos_concluidos = all(s in ['Feito', 'Feito e enviado'] for s in status_lista)
        if todos_concluidos: cor_final = '#28a745'
        else:
            for status_prioritario in ordem_prioridade:
                if status_prioritario in status_lista: status_final = status_prioritario; break
            for entrega in entregas_mes:
                if entrega['nome_status'] == status_final: cor_final = entrega['cor_hex']; break
        resultado_final[dia] = {'cor': cor_final, 'contagem': len(status_lista)}
    return resultado_final

def get_entregas_no_intervalo(data, hora_inicio, hora_fim):
    conn = conectar(); query = "SELECT e.id, e.horario, c.nome as nome_cliente, s.nome as nome_status FROM entregas e JOIN clientes c ON e.cliente_id = c.id LEFT JOIN status s ON e.status_id = s.id WHERE e.data_vencimento = ? AND e.horario BETWEEN ? AND ?"
    entregas = conn.execute(query, (data, hora_inicio, hora_fim)).fetchall(); conn.close()
    return [dict(row) for row in entregas]

def get_estatisticas_mensais(ano, mes):
    like_pattern = f"{ano}-{mes:02d}-%"; conn = conectar()
    query_concluidos = "SELECT COUNT(e.id) as contagem FROM entregas e JOIN status s ON e.status_id = s.id WHERE e.data_vencimento LIKE ? AND (s.nome = 'Feito' OR s.nome = 'Feito e enviado')"
    cursor_concluidos = conn.execute(query_concluidos, (like_pattern,)); contagem_concluidos = cursor_concluidos.fetchone()['contagem']
    query_retificados = "SELECT COUNT(e.id) as contagem FROM entregas e JOIN status s ON e.status_id = s.id WHERE e.data_vencimento LIKE ? AND s.nome = 'Retificado'"
    cursor_retificados = conn.execute(query_retificados, (like_pattern,)); contagem_retificados = cursor_retificados.fetchone()['contagem']
    conn.close(); return {'concluidos': contagem_concluidos, 'retificados': contagem_retificados}

# --- FUNÇÕES ADICIONADAS PARA RELATÓRIOS AVANÇADOS ---
def get_entregas_filtradas(data_inicio, data_fim, status_ids=None):
    """ Busca entregas com base em um período e uma lista de status. """
    conn = conectar()
    params = [data_inicio, data_fim]
    
    query = """
        SELECT e.data_vencimento, e.horario, e.responsavel, e.observacoes,
               c.nome as nome_cliente, c.tipo_envio, c.contato,
               s.nome as nome_status
        FROM entregas e
        JOIN clientes c ON e.cliente_id = c.id
        LEFT JOIN status s ON e.status_id = s.id
        WHERE e.data_vencimento BETWEEN ? AND ?
    """
    
    if status_ids:
        placeholders = ','.join('?' for _ in status_ids)
        query += f" AND e.status_id IN ({placeholders})"
        params.extend(status_ids)
        
    query += " ORDER BY e.data_vencimento, e.horario"
    
    entregas = conn.execute(query, params).fetchall()
    conn.close()
    return [dict(row) for row in entregas]

def get_logs_filtrados(data_inicio, data_fim, usuario_nome=None):
    """ Busca logs com base em um período e, opcionalmente, um usuário. """
    conn = conectar()
    params = [f"{data_inicio} 00:00:00", f"{data_fim} 23:59:59"]
    
    query = """
        SELECT timestamp, usuario_nome, acao, detalhes
        FROM logs
        WHERE timestamp BETWEEN ? AND ?
    """
    
    if usuario_nome and usuario_nome != "Todos":
        query += " AND usuario_nome = ?"
        params.append(usuario_nome)
        
    query += " ORDER BY timestamp DESC"
    
    logs_raw = conn.execute(query, params).fetchall()
    conn.close()
    
    # --- NOVA LÓGICA DE CONVERSÃO DE FUSO HORÁRIO ---
    logs_convertidos = []
    fuso_local = timezone(timedelta(hours=-3)) # Define o fuso como UTC-3 (Horário de Brasília)

    for log in logs_raw:
        log_dict = dict(log)
        
        # 1. Converte a string de data/hora do banco para um objeto datetime
        timestamp_utc_str = log_dict['timestamp']
        # O SQLite pode ter formatos diferentes, então tentamos os mais comuns
        try:
            timestamp_utc = datetime.strptime(timestamp_utc_str, '%Y-%m-%d %H:%M:%S')
        except ValueError:
            # Tenta com milissegundos, caso o formato do banco inclua
            timestamp_utc = datetime.strptime(timestamp_utc_str, '%Y-%m-%d %H:%M:%S.%f')

        # 2. Informa ao Python que este objeto está em UTC
        timestamp_utc = timestamp_utc.replace(tzinfo=timezone.utc)
        
        # 3. Converte o objeto para o nosso fuso horário local
        timestamp_local = timestamp_utc.astimezone(fuso_local)
        
        # 4. Formata de volta para uma string legível e atualiza o dicionário
        log_dict['timestamp'] = timestamp_local.strftime('%Y-%m-%d %H:%M:%S')
        
        logs_convertidos.append(log_dict)
        
    return logs_convertidos

def atualizar_usuario(id, username, password, usuario_logado):
    conn = conectar()
    senha_hash = hashlib.sha256(password.encode('utf-8')).hexdigest()
    conn.execute("UPDATE usuarios SET username=?, password_hash=? WHERE id=?", (username, senha_hash, id))
    conn.commit()
    conn.close()
    registrar_log(usuario_logado, "USUARIO_ATUALIZADO", f"Usuário ID: {id}, Nome: {username}")

def verificar_senha_usuario_atual(username, password):
    """Verifica se a senha fornecida corresponde à do usuário informado."""
    conn = conectar()
    senha_hash_fornecida = hashlib.sha256(password.encode('utf-8')).hexdigest()
    
    cursor = conn.execute("SELECT password_hash FROM usuarios WHERE username = ?", (username,))
    resultado = cursor.fetchone()
    conn.close()
    
    if resultado:
        senha_hash_salva = resultado['password_hash']
        return senha_hash_fornecida == senha_hash_salva
    return False