# database.py
import sqlite3

def conectar():
    """Conecta ao banco de dados e habilita o acesso por nome de coluna."""
    conn = sqlite3.connect('calendario.db')
    conn.row_factory = sqlite3.Row
    return conn

def iniciar_db():
    """Cria/atualiza as tabelas do banco de dados."""
    conn = conectar()
    cursor = conn.cursor()
    
    # Tabela de Status
    cursor.execute('CREATE TABLE IF NOT EXISTS status (id INTEGER PRIMARY KEY AUTOINCREMENT, nome TEXT NOT NULL UNIQUE, cor_hex TEXT NOT NULL)')
    
    # Tabela de Clientes
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS clientes (
            id INTEGER PRIMARY KEY AUTOINCREMENT, nome TEXT NOT NULL, tipo_envio TEXT NOT NULL, contato TEXT NOT NULL,
            gera_recibo BOOLEAN NOT NULL DEFAULT 0, conta_xmls BOOLEAN NOT NULL DEFAULT 0,
            nivel TEXT, outros_detalhes TEXT
        )''')

    # Tabela de Entregas (Agendamentos)
    # Ajustado com ON DELETE SET NULL para o status_id
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS entregas (
            id INTEGER PRIMARY KEY AUTOINCREMENT, data_vencimento TEXT NOT NULL, horario TEXT NOT NULL, 
            status_id INTEGER, cliente_id INTEGER NOT NULL, responsavel TEXT, observacoes TEXT,
            FOREIGN KEY (status_id) REFERENCES status (id) ON DELETE SET NULL, 
            FOREIGN KEY (cliente_id) REFERENCES clientes (id) ON DELETE CASCADE
        )''')
    
    # Insere status padrão se a tabela estiver vazia
    cursor.execute("SELECT COUNT(id) FROM status")
    if cursor.fetchone()[0] == 0:
        status_padrao = [
            ('PENDENTE', '#ffc107'), 
            ('Feito e enviado', '#28a745'), 
            ('Feito', '#007bff'), 
            ('Retificado', '#17a2b8'), 
            ('Houve Algum Erro', '#dc3545'), 
            ('Chamado', '#6f42c1'), 
            ('Remarcado', '#fd7e14'), 
            ('Realocado', '#6c757d')
        ]
        cursor.executemany("INSERT INTO status (nome, cor_hex) VALUES (?, ?)", status_padrao)

    conn.commit()
    conn.close()

# --- Funções de Clientes ---
def adicionar_cliente(nome, tipo_envio, contato, gera_recibo, conta_xmls, nivel, detalhes):
    conn = conectar()
    conn.execute("INSERT INTO clientes (nome, tipo_envio, contato, gera_recibo, conta_xmls, nivel, outros_detalhes) VALUES (?, ?, ?, ?, ?, ?, ?)", (nome, tipo_envio, contato, gera_recibo, conta_xmls, nivel, detalhes))
    conn.commit()
    conn.close()

def atualizar_cliente(id, nome, tipo_envio, contato, gera_recibo, conta_xmls, nivel, detalhes):
    conn = conectar()
    conn.execute("UPDATE clientes SET nome=?, tipo_envio=?, contato=?, gera_recibo=?, conta_xmls=?, nivel=?, outros_detalhes=? WHERE id=?", (nome, tipo_envio, contato, gera_recibo, conta_xmls, nivel, detalhes, id))
    conn.commit()
    conn.close()

def listar_clientes():
    conn = conectar()
    clientes = conn.execute("SELECT * FROM clientes ORDER BY nome").fetchall()
    conn.close()
    return clientes

def deletar_cliente(id):
    conn = conectar()
    conn.execute("DELETE FROM clientes WHERE id = ?", (id,))
    conn.commit()
    conn.close()

# --- Funções de Status ---
def listar_status():
    conn = conectar()
    status_list = conn.execute("SELECT * FROM status ORDER BY nome").fetchall()
    conn.close()
    return status_list

def adicionar_status(nome, cor_hex):
    conn = conectar()
    conn.execute("INSERT INTO status (nome, cor_hex) VALUES (?, ?)", (nome, cor_hex))
    conn.commit()
    conn.close()

def atualizar_status(id, nome, cor_hex):
    conn = conectar()
    conn.execute("UPDATE status SET nome=?, cor_hex=? WHERE id=?", (nome, cor_hex, id))
    conn.commit()
    conn.close()

def deletar_status(id):
    conn = conectar()
    conn.execute("DELETE FROM status WHERE id = ?", (id,))
    conn.commit()
    conn.close()

# --- Funções de Entregas ---
def adicionar_entrega(data_vencimento, horario, status_id, cliente_id, responsavel, observacoes):
    conn = conectar()
    conn.execute("INSERT INTO entregas (data_vencimento, horario, status_id, cliente_id, responsavel, observacoes) VALUES (?, ?, ?, ?, ?, ?)", (data_vencimento, horario, status_id, cliente_id, responsavel, observacoes))
    conn.commit()
    conn.close()

def atualizar_entrega(id, horario, status_id, cliente_id, responsavel, observacoes):
    conn = conectar()
    conn.execute("UPDATE entregas SET horario=?, status_id=?, cliente_id=?, responsavel=?, observacoes=? WHERE id=?", (horario, status_id, cliente_id, responsavel, observacoes, id))
    conn.commit()
    conn.close()

def deletar_entrega(id):
    conn = conectar()
    conn.execute("DELETE FROM entregas WHERE id = ?", (id,))
    conn.commit()
    conn.close()

def get_entregas_por_dia(data):
    conn = conectar()
    query = """
        SELECT e.id, e.data_vencimento, e.horario, e.responsavel, e.observacoes, e.cliente_id, e.status_id,
               c.nome as nome_cliente, c.tipo_envio, c.contato, s.nome as nome_status, s.cor_hex
        FROM entregas e JOIN clientes c ON e.cliente_id = c.id
        LEFT JOIN status s ON e.status_id = s.id
        WHERE e.data_vencimento = ? """
    entregas_dia = conn.execute(query, (data,)).fetchall()
    conn.close()
    return {entrega['horario']: dict(entrega) for entrega in entregas_dia}

def get_status_dias_para_mes(ano, mes):
    data_inicio = f"{ano}-{mes:02d}-01"
    data_fim = f"{ano}-{mes:02d}-31"
    conn = conectar()
    query = """
        SELECT e.data_vencimento, s.nome as nome_status, s.cor_hex
        FROM entregas e
        LEFT JOIN status s ON e.status_id = s.id
        WHERE e.data_vencimento BETWEEN ? AND ?
    """
    entregas_mes = conn.execute(query, (data_inicio, data_fim)).fetchall()
    conn.close()

    ordem_prioridade = [
        'Houve Algum Erro', 'Chamado', 'Remarcado', 'PENDENTE', 
        'Realocado', 'Retificado', 'Feito', 'Feito e enviado'
    ]
    
    status_por_dia = {}
    for entrega in entregas_mes:
        dia = int(entrega['data_vencimento'].split('-')[2])
        if dia not in status_por_dia:
            status_por_dia[dia] = []
        status_por_dia[dia].append(entrega['nome_status'])

    resultado_final = {}
    for dia, status_lista in status_por_dia.items():
        cor_final = '#28a745'
        status_final = 'Feito e enviado'

        todos_concluidos = all(s in ['Feito', 'Feito e enviado'] for s in status_lista)
        if todos_concluidos:
            cor_final = '#28a745'
        else:
            for status_prioritario in ordem_prioridade:
                if status_prioritario in status_lista:
                    status_final = status_prioritario
                    break
            
            for entrega in entregas_mes:
                if entrega['nome_status'] == status_final:
                    cor_final = entrega['cor_hex']
                    break
        
        resultado_final[dia] = {'cor': cor_final, 'contagem': len(status_lista)}

    return resultado_final

def get_entregas_para_relatorio(ano, mes):
    """ Busca todas as entregas de um mês e retorna uma lista ordenada para relatórios. """
    data_inicio = f"{ano}-{mes:02d}-01"
    data_fim = f"{ano}-{mes:02d}-31"
    conn = conectar()
    query = """
        SELECT e.data_vencimento, e.horario, e.responsavel, e.observacoes,
               c.nome as nome_cliente, c.tipo_envio, c.contato,
               s.nome as nome_status
        FROM entregas e
        JOIN clientes c ON e.cliente_id = c.id
        LEFT JOIN status s ON e.status_id = s.id
        WHERE e.data_vencimento BETWEEN ? AND ?
        ORDER BY e.data_vencimento, e.horario
    """
    entregas = conn.execute(query, (data_inicio, data_fim)).fetchall()
    conn.close()
    return [dict(row) for row in entregas]

# database.py
import sqlite3
# ... (todo o código anterior permanece o mesmo até o final)

def get_status_dias_para_mes(ano, mes):
    # ... (código inalterado)
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
    # ... (código inalterado)
    data_inicio = f"{ano}-{mes:02d}-01"; data_fim = f"{ano}-{mes:02d}-31"; conn = conectar()
    query = "SELECT e.data_vencimento, e.horario, e.responsavel, e.observacoes, c.nome as nome_cliente, c.tipo_envio, c.contato, s.nome as nome_status FROM entregas e JOIN clientes c ON e.cliente_id = c.id LEFT JOIN status s ON e.status_id = s.id WHERE e.data_vencimento BETWEEN ? AND ? ORDER BY e.data_vencimento, e.horario"
    entregas = conn.execute(query, (data_inicio, data_fim)).fetchall(); conn.close()
    return [dict(row) for row in entregas]

# --- NOVA FUNÇÃO PARA NOTIFICAÇÕES ---
def get_entregas_no_intervalo(data, hora_inicio, hora_fim):
    """ Busca entregas em uma data e intervalo de horários específicos. """
    conn = conectar()
    query = """
        SELECT e.id, e.horario, c.nome as nome_cliente, s.nome as nome_status
        FROM entregas e
        JOIN clientes c ON e.cliente_id = c.id
        LEFT JOIN status s ON e.status_id = s.id
        WHERE e.data_vencimento = ? AND e.horario BETWEEN ? AND ?
    """
    entregas = conn.execute(query, (data, hora_inicio, hora_fim)).fetchall()
    conn.close()
    return [dict(row) for row in entregas]