# database.py
import sqlite3

def conectar():
    conn = sqlite3.connect('calendario.db')
    conn.row_factory = sqlite3.Row
    return conn

def iniciar_db():
    # ... (código inalterado)
    conn = conectar(); cursor = conn.cursor()
    cursor.execute('CREATE TABLE IF NOT EXISTS status (id INTEGER PRIMARY KEY AUTOINCREMENT, nome TEXT NOT NULL UNIQUE, cor_hex TEXT NOT NULL)')
    cursor.execute('''CREATE TABLE IF NOT EXISTS clientes (id INTEGER PRIMARY KEY AUTOINCREMENT, nome TEXT NOT NULL, tipo_envio TEXT NOT NULL, contato TEXT NOT NULL, gera_recibo BOOLEAN NOT NULL DEFAULT 0, conta_xmls BOOLEAN NOT NULL DEFAULT 0, nivel TEXT, outros_detalhes TEXT)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS entregas (id INTEGER PRIMARY KEY AUTOINCREMENT, data_vencimento TEXT NOT NULL, horario TEXT NOT NULL, status_id INTEGER, cliente_id INTEGER NOT NULL, responsavel TEXT, observacoes TEXT, FOREIGN KEY (status_id) REFERENCES status (id), FOREIGN KEY (cliente_id) REFERENCES clientes (id) ON DELETE CASCADE)''')
    cursor.execute("SELECT COUNT(id) FROM status")
    if cursor.fetchone()[0] == 0:
        status_padrao = [('PENDENTE', '#ffc107'), ('Feito e enviado', '#28a745'), ('Feito', '#007bff'), ('Retificado', '#17a2b8'), ('Houve Algum Erro', '#dc3545'), ('Chamado', '#6f42c1'), ('Remarcado', '#fd7e14'), ('Realocado', '#6c757d')]
        cursor.executemany("INSERT INTO status (nome, cor_hex) VALUES (?, ?)", status_padrao)
    conn.commit(); conn.close()

# --- Funções de Clientes e Status (Inalteradas) ---
def adicionar_cliente(nome, tipo_envio, contato, gera_recibo, conta_xmls, nivel, detalhes):
    conn = conectar(); conn.execute("INSERT INTO clientes (nome, tipo_envio, contato, gera_recibo, conta_xmls, nivel, outros_detalhes) VALUES (?, ?, ?, ?, ?, ?, ?)", (nome, tipo_envio, contato, gera_recibo, conta_xmls, nivel, detalhes)); conn.commit(); conn.close()
def atualizar_cliente(id, nome, tipo_envio, contato, gera_recibo, conta_xmls, nivel, detalhes):
    conn = conectar(); conn.execute("UPDATE clientes SET nome=?, tipo_envio=?, contato=?, gera_recibo=?, conta_xmls=?, nivel=?, outros_detalhes=? WHERE id=?", (nome, tipo_envio, contato, gera_recibo, conta_xmls, nivel, detalhes, id)); conn.commit(); conn.close()
def listar_clientes():
    conn = conectar(); clientes = conn.execute("SELECT * FROM clientes ORDER BY nome").fetchall(); conn.close(); return clientes
def deletar_cliente(id):
    conn = conectar(); conn.execute("DELETE FROM clientes WHERE id = ?", (id,)); conn.commit(); conn.close()
def listar_status():
    conn = conectar(); status_list = conn.execute("SELECT * FROM status ORDER BY nome").fetchall(); conn.close(); return status_list

# --- Funções de Entregas (Inalteradas) ---
def adicionar_entrega(data_vencimento, horario, status_id, cliente_id, responsavel, observacoes):
    conn = conectar(); conn.execute("INSERT INTO entregas (data_vencimento, horario, status_id, cliente_id, responsavel, observacoes) VALUES (?, ?, ?, ?, ?, ?)", (data_vencimento, horario, status_id, cliente_id, responsavel, observacoes)); conn.commit(); conn.close()
def atualizar_entrega(id, horario, status_id, cliente_id, responsavel, observacoes):
    conn = conectar(); conn.execute("UPDATE entregas SET horario=?, status_id=?, cliente_id=?, responsavel=?, observacoes=? WHERE id=?", (horario, status_id, cliente_id, responsavel, observacoes, id)); conn.commit(); conn.close()
def deletar_entrega(id):
    conn = conectar(); conn.execute("DELETE FROM entregas WHERE id = ?", (id,)); conn.commit(); conn.close()
def get_entregas_por_dia(data):
    conn = conectar()
    query = """
        SELECT e.id, e.data_vencimento, e.horario, e.responsavel, e.observacoes, e.cliente_id, e.status_id,
               c.nome as nome_cliente, c.tipo_envio, c.contato, s.nome as nome_status, s.cor_hex
        FROM entregas e JOIN clientes c ON e.cliente_id = c.id
        LEFT JOIN status s ON e.status_id = s.id
        WHERE e.data_vencimento = ? """
    entregas_dia = conn.execute(query, (data,)).fetchall()
    conn.close(); return {entrega['horario']: dict(entrega) for entrega in entregas_dia}

# --- NOVA FUNÇÃO DE LÓGICA DE CORES ---
def get_status_dias_para_mes(ano, mes):
    """
    Busca todas as entregas do mês e retorna para cada dia:
    a cor do status de maior prioridade e a contagem total de entregas.
    """
    data_inicio = f"{ano}-{mes:02d}-01"; data_fim = f"{ano}-{mes:02d}-31"
    conn = conectar()
    query = """
        SELECT e.data_vencimento, s.nome as nome_status, s.cor_hex
        FROM entregas e
        LEFT JOIN status s ON e.status_id = s.id
        WHERE e.data_vencimento BETWEEN ? AND ?
    """
    entregas_mes = conn.execute(query, (data_inicio, data_fim)).fetchall()
    conn.close()

    # Define a ordem de prioridade dos status (do mais para o menos urgente)
    ordem_prioridade = [
        'Houve Algum Erro', 'Chamado', 'Remarcado', 'PENDENTE', 
        'Realocado', 'Retificado', 'Feito', 'Feito e enviado'
    ]
    
    # Agrupa os status por dia
    status_por_dia = {}
    for entrega in entregas_mes:
        dia = int(entrega['data_vencimento'].split('-')[2])
        if dia not in status_por_dia:
            status_por_dia[dia] = []
        status_por_dia[dia].append(entrega['nome_status'])

    # Decide a cor final para cada dia
    resultado_final = {}
    for dia, status_lista in status_por_dia.items():
        cor_final = '#28a745' # Cor padrão "Verde"
        status_final = 'Feito e enviado'

        # Lógica da Regra de Ouro
        todos_concluidos = all(s in ['Feito', 'Feito e enviado'] for s in status_lista)
        if todos_concluidos:
            cor_final = '#28a745' # Verde
        else:
            # Lógica de Prioridade Normal
            for status_prioritario in ordem_prioridade:
                if status_prioritario in status_lista:
                    status_final = status_prioritario
                    break # Encontrou o mais prioritário, pode parar
            
            # Pega a cor do status de maior prioridade encontrado
            for entrega in entregas_mes:
                if entrega['nome_status'] == status_final:
                    cor_final = entrega['cor_hex']
                    break
        
        resultado_final[dia] = {'cor': cor_final, 'contagem': len(status_lista)}

    return resultado_final