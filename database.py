# database.py
import sqlite3
import hashlib
from datetime import datetime

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
    cursor.execute('''CREATE TABLE IF NOT EXISTS clientes (id INTEGER PRIMARY KEY AUTOINCREMENT, nome TEXT NOT NULL, tipo_envio TEXT NOT NULL, contato TEXT NOT NULL, gera_recibo BOOLEAN NOT NULL DEFAULT 0, conta_xmls BOOLEAN NOT NULL DEFAULT 0, nivel TEXT, outros_detalhes TEXT)''')
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

def adicionar_cliente(nome, tipo_envio, contato, gera_recibo, conta_xmls, nivel, detalhes, usuario_logado):
    conn = conectar()
    conn.execute("INSERT INTO clientes (nome, tipo_envio, contato, gera_recibo, conta_xmls, nivel, outros_detalhes) VALUES (?, ?, ?, ?, ?, ?, ?)", (nome, tipo_envio, contato, gera_recibo, conta_xmls, nivel, detalhes))
    conn.commit()
    conn.close()
    registrar_log(usuario_logado, "CLIENTE_CRIADO", f"Cliente: {nome}")

def atualizar_cliente(id, nome, tipo_envio, contato, gera_recibo, conta_xmls, nivel, detalhes, usuario_logado):
    conn = conectar()
    conn.execute("UPDATE clientes SET nome=?, tipo_envio=?, contato=?, gera_recibo=?, conta_xmls=?, nivel=?, outros_detalhes=? WHERE id=?", (nome, tipo_envio, contato, gera_recibo, conta_xmls, nivel, detalhes, id))
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

def atualizar_entrega(id, horario, status_id, cliente_id, responsavel, observacoes, usuario_logado):
    conn = conectar()
    conn.execute("UPDATE entregas SET horario=?, status_id=?, cliente_id=?, responsavel=?, observacoes=? WHERE id=?", (horario, status_id, cliente_id, responsavel, observacoes, id))
    conn.commit()
    conn.close()
    registrar_log(usuario_logado, "AGENDAMENTO_ATUALIZADO", f"Agendamento ID: {id}")

def deletar_entrega(id, usuario_logado):
    conn = conectar()
    conn.execute("DELETE FROM entregas WHERE id = ?", (id,))
    conn.commit()
    conn.close()
    registrar_log(usuario_logado, "AGENDAMENTO_DELETADO", f"Agendamento ID: {id}")

def get_entregas_por_dia(data):
    conn = conectar()
    query = "SELECT e.id, e.data_vencimento, e.horario, e.responsavel, e.observacoes, e.cliente_id, e.status_id, c.nome as nome_cliente, c.tipo_envio, c.contato, s.nome as nome_status, s.cor_hex FROM entregas e JOIN clientes c ON e.cliente_id = c.id LEFT JOIN status s ON e.status_id = s.id WHERE e.data_vencimento = ? "
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

def get_entregas_para_relatorio(ano, mes):
    data_inicio = f"{ano}-{mes:02d}-01"; data_fim = f"{ano}-{mes:02d}-31"; conn = conectar()
    query = "SELECT e.data_vencimento, e.horario, e.responsavel, e.observacoes, c.nome as nome_cliente, c.tipo_envio, c.contato, s.nome as nome_status FROM entregas e JOIN clientes c ON e.cliente_id = c.id LEFT JOIN status s ON e.status_id = s.id WHERE e.data_vencimento BETWEEN ? AND ? ORDER BY e.data_vencimento, e.horario"
    entregas = conn.execute(query, (data_inicio, data_fim)).fetchall(); conn.close()
    return [dict(row) for row in entregas]

def get_entregas_no_intervalo(data, hora_inicio, hora_fim):
    conn = conectar(); query = "SELECT e.id, e.horario, c.nome as nome_cliente, s.nome as nome_status FROM entregas e JOIN clientes c ON e.cliente_id = c.id LEFT JOIN status s ON e.status_id = s.id WHERE e.data_vencimento = ? AND e.horario BETWEEN ? AND ?"
    entregas = conn.execute(query, (data, hora_inicio, hora_fim)).fetchall(); conn.close()
    return [dict(row) for row in entregas]

def get_estatisticas_mensais(ano, mes):
    """ Calcula as estatísticas de status para um determinado mês. """
    like_pattern = f"{ano}-{mes:02d}-%"
    conn = conectar()
    query_concluidos = "SELECT COUNT(e.id) as contagem FROM entregas e JOIN status s ON e.status_id = s.id WHERE e.data_vencimento LIKE ? AND (s.nome = 'Feito' OR s.nome = 'Feito e enviado')"
    cursor_concluidos = conn.execute(query_concluidos, (like_pattern,))
    contagem_concluidos = cursor_concluidos.fetchone()['contagem']
    query_retificados = "SELECT COUNT(e.id) as contagem FROM entregas e JOIN status s ON e.status_id = s.id WHERE e.data_vencimento LIKE ? AND s.nome = 'Retificado'"
    cursor_retificados = conn.execute(query_retificados, (like_pattern,))
    contagem_retificados = cursor_retificados.fetchone()['contagem']
    conn.close()
    return {'concluidos': contagem_concluidos, 'retificados': contagem_retificados}