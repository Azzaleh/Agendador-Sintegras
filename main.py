# main.py
import sys
from datetime import datetime, timedelta
import calendar
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QGridLayout, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QLabel, QDialog, QLineEdit, QComboBox, QMessageBox, 
                             QFileDialog, QFormLayout, QCheckBox, QTextEdit, QTableWidget, 
                             QTableWidgetItem, QHeaderView)
from PyQt5.QtGui import QPainter, QColor, QBrush, QFont
from PyQt5.QtCore import Qt, QDate, pyqtSignal

import database
import export

# --- Widgets Customizados ---
class DayCellWidget(QWidget):
    """ Widget ATUALIZADO para usar a cor de prioridade. """
    clicked = pyqtSignal(QDate)
    def __init__(self, date, dia_info, parent=None):
        super().__init__(parent)
        self.date = date
        # dia_info é um dicionário: {'cor': '#hex', 'contagem': N}
        self.dia_info = dia_info if dia_info else {'cor': '#ffffff', 'contagem': 0}
        self.setMinimumSize(100, 80)
        
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        
        # --- LÓGICA DE COR DE FUNDO BASEADA NA PRIORIDADE ---
        cor_hex = self.dia_info['cor']
        cor_fundo = QColor(cor_hex) if self.dia_info['contagem'] > 0 else QColor("#ffffff")
        
        if self.date.month() != datetime.now().month:
            cor_fundo = QColor("#f0f0f0")
        
        painter.setBrush(cor_fundo)
        
        # Borda
        if self.date == QDate.currentDate():
            painter.setPen(QColor("#007bff"))
        else:
            painter.setPen(cor_fundo.darker(110))
            
        painter.drawRect(self.rect())
        
        # Número do dia
        # Muda a cor do texto para branco se o fundo for muito escuro
        painter.setPen(Qt.black if cor_fundo.lightness() > 127 else Qt.white)
        painter.setFont(QFont('Arial', 10, QFont.Bold))
        painter.drawText(5, 20, str(self.date.day()))
        
        # Legenda da contagem
        contagem = self.dia_info['contagem']
        if contagem > 0:
            painter.setFont(QFont('Arial', 8))
            painter.drawText(5, self.height() - 5, f"{contagem} agend.")

    def mousePressEvent(self, event):
        self.clicked.emit(self.date)

# --- Janelas de Cliente ---
class DialogoCliente(QDialog):
    def __init__(self, cliente_id=None, parent=None):
        super().__init__(parent)
        self.cliente_id = cliente_id
        self.setWindowTitle("Adicionar Novo Cliente" if not cliente_id else "Editar Cliente")
        self.setMinimumWidth(400)
        layout = QVBoxLayout(self)
        form_layout = QFormLayout()
        self.nome_edit = QLineEdit()
        self.tipo_envio_combo = QComboBox()
        self.tipo_envio_combo.addItems(["Nosso", "Deles"])
        self.contato_edit = QLineEdit()
        self.gera_recibo_check = QCheckBox("Gera Recibo?")
        self.conta_xmls_check = QCheckBox("Contar XMLs?")
        self.nivel_edit = QLineEdit()
        self.detalhes_edit = QTextEdit()
        self.detalhes_edit.setPlaceholderText("Detalhes adicionais sobre o cliente...")
        form_layout.addRow("Nome*:", self.nome_edit)
        form_layout.addRow("Tipo de Envio*:", self.tipo_envio_combo)
        form_layout.addRow("Email/Local*:", self.contato_edit)
        form_layout.addRow("", self.gera_recibo_check)
        form_layout.addRow("", self.conta_xmls_check)
        form_layout.addRow("Nível:", self.nivel_edit)
        form_layout.addRow("Outros Detalhes:", self.detalhes_edit)
        layout.addLayout(form_layout)
        botoes_layout = QHBoxLayout()
        self.salvar_btn = QPushButton("Salvar")
        self.cancelar_btn = QPushButton("Cancelar")
        botoes_layout.addStretch()
        botoes_layout.addWidget(self.salvar_btn)
        botoes_layout.addWidget(self.cancelar_btn)
        layout.addLayout(botoes_layout)
        self.salvar_btn.clicked.connect(self.salvar)
        self.cancelar_btn.clicked.connect(self.reject)
        if self.cliente_id:
            self.carregar_dados()
            
    def carregar_dados(self):
        pass

    def salvar(self):
        nome = self.nome_edit.text().strip()
        tipo_envio = self.tipo_envio_combo.currentText()
        contato = self.contato_edit.text().strip()
        if not nome or not tipo_envio or not contato:
            QMessageBox.warning(self, "Campos Obrigatórios", "Por favor, preencha os campos Nome, Tipo de Envio e Email/Local.")
            return
        dados_cliente = { "nome": nome, "tipo_envio": tipo_envio, "contato": contato, "gera_recibo": self.gera_recibo_check.isChecked(), "conta_xmls": self.conta_xmls_check.isChecked(), "nivel": self.nivel_edit.text().strip(), "detalhes": self.detalhes_edit.toPlainText().strip() }
        if self.cliente_id:
            database.atualizar_cliente(self.cliente_id, **dados_cliente)
        else:
            database.adicionar_cliente(**dados_cliente)
        self.accept()

class JanelaClientes(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Gerenciamento de Clientes")
        self.setMinimumSize(800, 500)
        layout = QVBoxLayout(self)
        self.tabela_clientes = QTableWidget()
        self.tabela_clientes.setColumnCount(4)
        self.tabela_clientes.setHorizontalHeaderLabels(["Nome", "Tipo de Envio", "Email/Local", "Nível"])
        self.tabela_clientes.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tabela_clientes.setSelectionBehavior(QTableWidget.SelectRows)
        self.tabela_clientes.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.tabela_clientes)
        botoes_layout = QHBoxLayout()
        import_btn = QPushButton("Importar de XLSX")
        add_btn = QPushButton("Adicionar Novo")
        edit_btn = QPushButton("Editar Selecionado")
        del_btn = QPushButton("Excluir Selecionado")
        botoes_layout.addWidget(import_btn)
        botoes_layout.addStretch()
        botoes_layout.addWidget(add_btn)
        botoes_layout.addWidget(edit_btn)
        botoes_layout.addWidget(del_btn)
        layout.addLayout(botoes_layout)
        import_btn.clicked.connect(self.importar_clientes)
        add_btn.clicked.connect(self.adicionar_cliente)
        edit_btn.clicked.connect(self.editar_cliente)
        del_btn.clicked.connect(self.excluir_cliente)
        self.carregar_clientes()

    def importar_clientes(self):
        caminho_arquivo, _ = QFileDialog.getOpenFileName(self, "Selecionar arquivo XLSX", "", "Arquivos Excel (*.xlsx)")
        if not caminho_arquivo:
            return
        try:
            df = pd.read_excel(caminho_arquivo)
            mapa_colunas = { 'Clientes': 'nome', 'Envia do nosso ou deles?': 'tipo_envio', 'Email da contabilidade (Ou local a ser deixado)': 'contato', 'Gera Recibo?': 'gera_recibo', 'Contar XMLs?': 'conta_xmls', 'Nível': 'nivel', 'Outros detalhes': 'detalhes' }
            colunas_obrigatorias = ['Clientes', 'Envia do nosso ou deles?', 'Email da contabilidade (Ou local a ser deixado)']
            if not all(col in df.columns for col in colunas_obrigatorias):
                QMessageBox.critical(self, "Erro de Importação", f"O arquivo deve conter as colunas obrigatórias: {', '.join(colunas_obrigatorias)}")
                return
            clientes_importados = 0
            for _, row in df.iterrows():
                dados_cliente = {}
                for col_excel, col_db in mapa_colunas.items():
                    if col_excel in row:
                        valor = row[col_excel]
                        if col_excel in ['Gera Recibo?', 'Contar XMLs?']:
                            dados_cliente[col_db] = str(valor).lower() in ['sim', 'true', '1']
                        elif pd.isna(valor):
                            dados_cliente[col_db] = None
                        else:
                            dados_cliente[col_db] = str(valor)
                    else:
                        dados_cliente[col_db] = None
                if dados_cliente.get('nome') and dados_cliente.get('tipo_envio') and dados_cliente.get('contato'):
                    database.adicionar_cliente(nome=dados_cliente['nome'], tipo_envio=dados_cliente['tipo_envio'], contato=dados_cliente['contato'], gera_recibo=dados_cliente.get('gera_recibo', False), conta_xmls=dados_cliente.get('conta_xmls', False), nivel=dados_cliente.get('nivel'), detalhes=dados_cliente.get('detalhes'))
                    clientes_importados += 1
            QMessageBox.information(self, "Sucesso", f"{clientes_importados} clientes importados com sucesso!")
            self.carregar_clientes()
        except Exception as e:
            QMessageBox.critical(self, "Erro de Importação", f"Ocorreu um erro ao ler o arquivo:\n{e}")

    def carregar_clientes(self):
        self.tabela_clientes.setRowCount(0)
        for cliente in database.listar_clientes():
            row = self.tabela_clientes.rowCount()
            self.tabela_clientes.insertRow(row)
            self.tabela_clientes.setItem(row, 0, QTableWidgetItem(cliente['nome']))
            self.tabela_clientes.setItem(row, 1, QTableWidgetItem(cliente['tipo_envio']))
            self.tabela_clientes.setItem(row, 2, QTableWidgetItem(cliente['contato']))
            self.tabela_clientes.setItem(row, 3, QTableWidgetItem(cliente['nivel']))
            self.tabela_clientes.item(row, 0).setData(Qt.UserRole, dict(cliente))

    def adicionar_cliente(self):
        dialog = DialogoCliente(parent=self)
        if dialog.exec_() == QDialog.Accepted:
            self.carregar_clientes()

    def editar_cliente(self):
        linha_selecionada = self.tabela_clientes.currentRow()
        if linha_selecionada < 0:
            QMessageBox.information(self, "Aviso", "Por favor, selecione um cliente na tabela para editar.")
            return
        cliente_data = self.tabela_clientes.item(linha_selecionada, 0).data(Qt.UserRole)
        dialog = DialogoCliente(cliente_id=cliente_data['id'], parent=self)
        dialog.nome_edit.setText(cliente_data['nome'])
        dialog.tipo_envio_combo.setCurrentText(cliente_data['tipo_envio'])
        dialog.contato_edit.setText(cliente_data['contato'])
        dialog.gera_recibo_check.setChecked(bool(cliente_data['gera_recibo']))
        dialog.conta_xmls_check.setChecked(bool(cliente_data['conta_xmls']))
        dialog.nivel_edit.setText(cliente_data['nivel'])
        dialog.detalhes_edit.setText(cliente_data['outros_detalhes'])
        if dialog.exec_() == QDialog.Accepted:
            self.carregar_clientes()

    def excluir_cliente(self):
        linha_selecionada = self.tabela_clientes.currentRow()
        if linha_selecionada < 0:
            QMessageBox.information(self, "Aviso", "Por favor, selecione um cliente para excluir.")
            return
        cliente_data = self.tabela_clientes.item(linha_selecionada, 0).data(Qt.UserRole)
        reply = QMessageBox.question(self, "Confirmar Exclusão", f"Tem certeza que deseja excluir o cliente '{cliente_data['nome']}'?\nTODAS as entregas associadas a ele também serão excluídas.", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            database.deletar_cliente(cliente_data['id'])
            self.carregar_clientes()

# --- Janelas de Agendamento ---
class EntregaDialog(QDialog):
    def __init__(self, entrega_data=None, parent=None):
        super().__init__(parent)
        self.entrega_data = entrega_data
        self.setWindowTitle("Agendar Nova Entrega" if not entrega_data else "Editar Entrega")
        layout = QFormLayout(self)
        self.cliente_combo = QComboBox()
        self.status_combo = QComboBox()
        self.responsavel_edit = QLineEdit()
        self.observacoes_edit = QTextEdit()
        self.carregar_combos()
        layout.addRow("Cliente:", self.cliente_combo)
        layout.addRow("Status:", self.status_combo)
        layout.addRow("Responsável:", self.responsavel_edit)
        layout.addRow("Observações:", self.observacoes_edit)
        
        if entrega_data:
            index_cliente = self.cliente_combo.findData(entrega_data['cliente_id'])
            if index_cliente > -1: self.cliente_combo.setCurrentIndex(index_cliente)
            index_status = self.status_combo.findData(entrega_data['status_id'])
            if index_status > -1: self.status_combo.setCurrentIndex(index_status)
            self.responsavel_edit.setText(entrega_data['responsavel'])
            self.observacoes_edit.setText(entrega_data['observacoes'])
        else:
            index_pendente = self.status_combo.findText("PENDENTE")
            if index_pendente > -1:
                self.status_combo.setCurrentIndex(index_pendente)
        
        botoes_layout = QHBoxLayout()
        self.salvar_btn = QPushButton("Salvar")
        self.cancelar_btn = QPushButton("Cancelar")
        botoes_layout.addStretch()
        botoes_layout.addWidget(self.salvar_btn)
        botoes_layout.addWidget(self.cancelar_btn)
        layout.addRow(botoes_layout)
        self.salvar_btn.clicked.connect(self.accept)
        self.cancelar_btn.clicked.connect(self.reject)

    def carregar_combos(self):
        for cliente in database.listar_clientes():
            self.cliente_combo.addItem(cliente['nome'], cliente['id'])
        for status in database.listar_status():
            self.status_combo.addItem(status['nome'], status['id'])

    def get_data(self):
        return { "cliente_id": self.cliente_combo.currentData(), "status_id": self.status_combo.currentData(), "responsavel": self.responsavel_edit.text().strip(), "observacoes": self.observacoes_edit.toPlainText().strip() }

class DayViewDialog(QDialog):
    def __init__(self, date, parent=None):
        super().__init__(parent)
        self.date = date
        self.setWindowTitle(f"Agenda para {date.toString('dd/MM/yyyy')}")
        self.setMinimumSize(800, 600)
        self.horarios = [f"{h:02d}:{m:02d}" for h in range(8, 18) for m in (0, 30)][1:-1]
        layout = QVBoxLayout(self)
        self.tabela_agenda = QTableWidget()
        self.tabela_agenda.setColumnCount(5)
        self.tabela_agenda.setHorizontalHeaderLabels(["Horário", "Cliente", "Contato", "Tipo Envio", "Status"])
        self.tabela_agenda.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tabela_agenda.setSelectionBehavior(QTableWidget.SelectRows)
        self.tabela_agenda.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tabela_agenda.setColumnWidth(0, 80)
        layout.addWidget(self.tabela_agenda)
        self.tabela_agenda.cellDoubleClicked.connect(self.gerenciar_agendamento)
        self.carregar_agenda_dia()

    def carregar_agenda_dia(self):
        self.tabela_agenda.setRowCount(len(self.horarios))
        agendamentos_dia = database.get_entregas_por_dia(self.date.toString("yyyy-MM-dd"))
        for i, horario in enumerate(self.horarios):
            item_horario = QTableWidgetItem(horario)
            item_horario.setTextAlignment(Qt.AlignCenter)
            self.tabela_agenda.setItem(i, 0, item_horario)
            if horario in agendamentos_dia:
                agendamento = agendamentos_dia[horario]
                self.tabela_agenda.setItem(i, 1, QTableWidgetItem(agendamento['nome_cliente']))
                self.tabela_agenda.setItem(i, 2, QTableWidgetItem(agendamento['contato']))
                self.tabela_agenda.setItem(i, 3, QTableWidgetItem(agendamento['tipo_envio']))
                item_status = QTableWidgetItem(agendamento['nome_status'])
                item_status.setBackground(QColor(agendamento['cor_hex']))
                self.tabela_agenda.setItem(i, 4, item_status)
                item_horario.setData(Qt.UserRole, agendamento)

    def gerenciar_agendamento(self, row, column):
        item_horario = self.tabela_agenda.item(row, 0)
        horario = item_horario.text()
        agendamento_existente = item_horario.data(Qt.UserRole)
        if agendamento_existente:
            dialog = EntregaDialog(entrega_data=agendamento_existente, parent=self)
            if dialog.exec_() == QDialog.Accepted:
                data = dialog.get_data()
                database.atualizar_entrega(agendamento_existente['id'], horario, data['status_id'], data['cliente_id'], data['responsavel'], data['observacoes'])
        else:
            dialog = EntregaDialog(parent=self)
            if dialog.exec_() == QDialog.Accepted:
                data = dialog.get_data()
                database.adicionar_entrega(self.date.toString("yyyy-MM-dd"), horario, data['status_id'], data['cliente_id'], data['responsavel'], data['observacoes'])
        self.carregar_agenda_dia()
        self.parent().populate_calendar()

# --- Janela Principal ---
class CalendarWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Agendador Mensal de Entregas")
        self.setGeometry(100, 100, 900, 700)
        self.current_date = datetime.now()
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)
        self.setup_ui()
        self.populate_calendar()
        
    def setup_ui(self):
        nav_layout = QHBoxLayout()
        prev_btn = QPushButton("< Mês Anterior")
        prev_btn.clicked.connect(self.prev_month)
        self.month_label = QLabel()
        self.month_label.setAlignment(Qt.AlignCenter)
        next_btn = QPushButton("Próximo Mês >")
        next_btn.clicked.connect(self.next_month)
        nav_layout.addWidget(prev_btn)
        nav_layout.addWidget(self.month_label)
        nav_layout.addWidget(next_btn)
        self.main_layout.addLayout(nav_layout)
        header_layout = QGridLayout()
        dias = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]
        for i, dia in enumerate(dias):
            header_layout.addWidget(QLabel(f"<b>{dia}</b>", alignment=Qt.AlignCenter), 0, i)
        self.main_layout.addLayout(header_layout)
        self.calendar_grid = QGridLayout()
        self.calendar_grid.setSpacing(0)
        self.main_layout.addLayout(self.calendar_grid)
        action_layout = QHBoxLayout()
        clientes_btn = QPushButton("Gerenciar Clientes")
        clientes_btn.clicked.connect(self.gerenciar_clientes)
        status_btn = QPushButton("Gerenciar Status")
        status_btn.clicked.connect(self.manage_status)
        export_btn = QPushButton("Exportar Mês")
        export_btn.clicked.connect(self.export_month)
        action_layout.addWidget(clientes_btn)
        action_layout.addWidget(status_btn)
        action_layout.addWidget(export_btn)
        self.main_layout.addLayout(action_layout)
        
    def populate_calendar(self):
        for i in reversed(range(self.calendar_grid.count())):
            self.calendar_grid.itemAt(i).widget().setParent(None)
        year = self.current_date.year
        month = self.current_date.month
        self.month_label.setText(f"<b>{self.current_date.strftime('%B %Y')}</b>")
        
        status_dias = database.get_status_dias_para_mes(year, month)
        
        month_calendar = calendar.monthcalendar(year, month)
        for week_num, week in enumerate(month_calendar):
            for day_num, day in enumerate(week):
                if day != 0:
                    date = QDate(year, month, day)
                    info_do_dia = status_dias.get(day)
                    cell = DayCellWidget(date, info_do_dia, self)
                    cell.clicked.connect(self.open_day_view)
                    self.calendar_grid.addWidget(cell, week_num, day_num)

    def prev_month(self):
        self.current_date -= timedelta(days=self.current_date.day + 1)
        self.populate_calendar()

    def next_month(self):
        _, days_in_month = calendar.monthrange(self.current_date.year, self.current_date.month)
        self.current_date += timedelta(days=days_in_month - self.current_date.day + 1)
        self.populate_calendar()
        
    def open_day_view(self, date):
        dialog = DayViewDialog(date, self)
        dialog.exec_()

    def gerenciar_clientes(self):
        dialog = JanelaClientes(self)
        dialog.exec_()

    def manage_status(self):
        QMessageBox.information(self, "Aviso", "A janela de gerenciamento de status ainda não foi implementada.")

    def export_month(self):
        QMessageBox.information(self, "Aviso", "A função de exportação será atualizada para o novo formato de dados.")

if __name__ == '__main__':
    database.iniciar_db()
    app = QApplication(sys.argv)
    window = CalendarWindow()
    window.show()
    sys.exit(app.exec_())