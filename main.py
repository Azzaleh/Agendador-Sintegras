# main.py
import sys
import os
from datetime import datetime, timedelta
import calendar
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QGridLayout, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QLabel, QDialog, QLineEdit, QComboBox, QMessageBox, 
                             QFileDialog, QFormLayout, QCheckBox, QTextEdit, QTableWidget, 
                             QTableWidgetItem, QHeaderView, QColorDialog, QSystemTrayIcon, QStyle,
                             QTimeEdit, QSpinBox, QRadioButton, QGroupBox, QDateEdit, QListWidget, QListWidgetItem,QMenu)
from PyQt5.QtGui import QPainter, QColor, QBrush, QFont, QIcon
from PyQt5.QtCore import Qt, QDate, pyqtSignal, QTimer, QTime, QSettings

import database
import export

class LoginDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Login - Agendador")
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        layout = QVBoxLayout(self)
        form_layout = QFormLayout()
        self.usuario_edit = QLineEdit(self)
        self.usuario_edit.setPlaceholderText("Digite seu usuário")
        self.senha_edit = QLineEdit(self)
        self.senha_edit.setPlaceholderText("Digite sua senha")
        self.senha_edit.setEchoMode(QLineEdit.Password)
        form_layout.addRow(QLabel("Usuário:"), self.usuario_edit)
        form_layout.addRow(QLabel("Senha:"), self.senha_edit)
        layout.addLayout(form_layout)
        botoes_layout = QHBoxLayout()
        login_btn = QPushButton("Entrar")
        sair_btn = QPushButton("Sair")
        botoes_layout.addStretch()
        botoes_layout.addWidget(login_btn)
        botoes_layout.addWidget(sair_btn)
        layout.addLayout(botoes_layout)
        login_btn.clicked.connect(self.tentar_login)
        sair_btn.clicked.connect(self.reject)
        self.usuario_logado = None
    def tentar_login(self):
        username = self.usuario_edit.text()
        password = self.senha_edit.text()
        usuario = database.verificar_usuario(username, password)
        if usuario:
            self.usuario_logado = usuario
            self.accept()
        else:
            QMessageBox.warning(self, "Erro de Login", "Usuário ou senha inválidos.")
            self.senha_edit.clear()

class DayCellWidget(QWidget):
    clicked = pyqtSignal(QDate)

    def __init__(self, date, dia_info, day_of_week, parent=None):
        super().__init__(parent)
        self.date = date
        self.dia_info = dia_info if dia_info else {'cor': '#ffffff', 'contagem': 0}
        self.day_of_week = day_of_week
        self.calendar_window = parent
        self.setMinimumSize(100, 80)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        contagem = self.dia_info['contagem']
        is_weekend = self.day_of_week in [5, 6]

        # ---- Verifica se é feriado ----
        feriados = self.calendar_window.feriados
        feriado_tipo = feriados.get(self.date)

        if feriado_tipo == "nacional":
            cor_fundo = QColor("#d8bfd8")  # Roxo pastel
        elif feriado_tipo == "municipal":
            cor_fundo = QColor("#e6e6fa")  # Lilás pastel
        elif contagem > 0:
            cor_fundo = QColor(self.dia_info['cor'])
        elif is_weekend:
            cor_fundo = QColor("#ffe1c3")
        else:
            cor_fundo = QColor("#b9fa9f")

        if self.date.month() != self.calendar_window.current_date.month:
            cor_fundo = QColor("#f0f0f0")

        painter.setBrush(cor_fundo)

        if self.date == QDate.currentDate():
            painter.setPen(QColor("#007bff"))
        else:
            painter.setPen(cor_fundo.darker(110))

        painter.drawRect(self.rect())

        painter.setPen(Qt.black if cor_fundo.lightness() > 127 else Qt.white)
        painter.setFont(QFont('Arial', 10, QFont.Bold))
        painter.drawText(5, 20, str(self.date.day()))

        if contagem > 0:
            painter.setFont(QFont('Arial', 8))
            painter.drawText(5, self.height() - 5, f"{contagem} agend.")

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.clicked.emit(self.date)

    def contextMenuEvent(self, event):
        menu = QMenu(self)

        marcar_nacional = menu.addAction("Marcar como Feriado Nacional")
        marcar_municipal = menu.addAction("Marcar como Feriado Municipal")
        desmarcar = menu.addAction("Desmarcar Feriado")

        action = menu.exec_(event.globalPos())

        if action == marcar_nacional:
            self.calendar_window.feriados[self.date] = "nacional"
        elif action == marcar_municipal:
            self.calendar_window.feriados[self.date] = "municipal"
        elif action == desmarcar:
            if self.date in self.calendar_window.feriados:
                del self.calendar_window.feriados[self.date]

        # Redesenhar o calendário
        self.calendar_window.populate_calendar()



class DialogoCliente(QDialog):
    def __init__(self, usuario_logado, cliente_id=None, parent=None):
        super().__init__(parent)
        self.usuario_logado = usuario_logado
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
    def carregar_dados(self): pass
    def salvar(self):
        nome = self.nome_edit.text().strip()
        tipo_envio = self.tipo_envio_combo.currentText()
        contato = self.contato_edit.text().strip()
        if not nome or not tipo_envio or not contato:
            QMessageBox.warning(self, "Campos Obrigatórios", "Por favor, preencha os campos Nome, Tipo de Envio e Email/Local.")
            return
        dados_cliente = { "nome": nome, "tipo_envio": tipo_envio, "contato": contato, "gera_recibo": self.gera_recibo_check.isChecked(), "conta_xmls": self.conta_xmls_check.isChecked(), "nivel": self.nivel_edit.text().strip(), "detalhes": self.detalhes_edit.toPlainText().strip() }
        usuario_nome = self.usuario_logado['username']
        if self.cliente_id:
            database.atualizar_cliente(self.cliente_id, **dados_cliente, usuario_logado=usuario_nome)
        else:
            database.adicionar_cliente(**dados_cliente, usuario_logado=usuario_nome)
        self.accept()

class JanelaClientes(QDialog):
    def __init__(self, usuario_logado, parent=None):
        super().__init__(parent)
        self.usuario_logado = usuario_logado
        self.setWindowTitle("Gerenciamento de Clientes")
        self.setMinimumSize(800, 600)
        layout = QVBoxLayout(self)
        busca_layout = QHBoxLayout()
        busca_label = QLabel("Buscar Cliente:")
        self.busca_edit = QLineEdit()
        self.busca_edit.setPlaceholderText("Digite o nome do cliente para filtrar...")
        busca_layout.addWidget(busca_label)
        busca_layout.addWidget(self.busca_edit)
        layout.addLayout(busca_layout)
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
        self.busca_edit.textChanged.connect(self.filtrar_tabela)
        self.carregar_clientes()
    def filtrar_tabela(self):
        texto_busca = self.busca_edit.text().lower().strip()
        for i in range(self.tabela_clientes.rowCount()):
            item_nome = self.tabela_clientes.item(i, 0)
            if item_nome and texto_busca in item_nome.text().lower():
                self.tabela_clientes.setRowHidden(i, False)
            else:
                self.tabela_clientes.setRowHidden(i, True)
    def importar_clientes(self):
        caminho_arquivo, _ = QFileDialog.getOpenFileName(self, "Selecionar arquivo XLSX", "", "Arquivos Excel (*.xlsx)")
        if not caminho_arquivo: return
        try:
            df = pd.read_excel(caminho_arquivo)
            mapa_colunas = { 'Clientes': 'nome', 'Envia do nosso ou deles?': 'tipo_envio', 'Email da contabilidade (Ou local a ser deixado)': 'contato', 'Gera Recibo?': 'gera_recibo', 'Contar XMLs?': 'conta_xmls', 'Nível': 'nivel', 'Outros detalhes': 'detalhes' }
            colunas_obrigatorias = ['Clientes', 'Envia do nosso ou deles?', 'Email da contabilidade (Ou local a ser deixado)']
            if not all(col in df.columns for col in colunas_obrigatorias):
                QMessageBox.critical(self, "Erro de Importação", f"O arquivo deve conter as colunas obrigatórias: {', '.join(colunas_obrigatorias)}"); return
            clientes_importados = 0
            usuario_nome = self.usuario_logado['username']
            for _, row in df.iterrows():
                dados_cliente = {}
                for col_excel, col_db in mapa_colunas.items():
                    if col_excel in row:
                        valor = row[col_excel]
                        if col_excel in ['Gera Recibo?', 'Contar XMLs?']: dados_cliente[col_db] = str(valor).lower() in ['sim', 'true', '1']
                        elif pd.isna(valor): dados_cliente[col_db] = None
                        else: dados_cliente[col_db] = str(valor)
                    else: dados_cliente[col_db] = None
                if dados_cliente.get('nome') and dados_cliente.get('tipo_envio') and dados_cliente.get('contato'):
                    database.adicionar_cliente(nome=dados_cliente['nome'], tipo_envio=dados_cliente['tipo_envio'], contato=dados_cliente['contato'], gera_recibo=dados_cliente.get('gera_recibo', False), conta_xmls=dados_cliente.get('conta_xmls', False), nivel=dados_cliente.get('nivel'), detalhes=dados_cliente.get('detalhes'), usuario_logado=usuario_nome); clientes_importados += 1
            QMessageBox.information(self, "Sucesso", f"{clientes_importados} clientes importados com sucesso!"); self.carregar_clientes()
        except Exception as e: QMessageBox.critical(self, "Erro de Importação", f"Ocorreu um erro ao ler o arquivo:\n{e}")
    def carregar_clientes(self):
        self.tabela_clientes.setRowCount(0)
        for cliente in database.listar_clientes():
            row = self.tabela_clientes.rowCount(); self.tabela_clientes.insertRow(row)
            self.tabela_clientes.setItem(row, 0, QTableWidgetItem(cliente['nome'])); self.tabela_clientes.setItem(row, 1, QTableWidgetItem(cliente['tipo_envio']))
            self.tabela_clientes.setItem(row, 2, QTableWidgetItem(cliente['contato'])); self.tabela_clientes.setItem(row, 3, QTableWidgetItem(cliente['nivel']))
            self.tabela_clientes.item(row, 0).setData(Qt.UserRole, dict(cliente))
        self.filtrar_tabela()
    def adicionar_cliente(self):
        dialog = DialogoCliente(self.usuario_logado, parent=self)
        if dialog.exec_() == QDialog.Accepted: self.carregar_clientes()
    def editar_cliente(self):
        itens_selecionados = self.tabela_clientes.selectedItems()
        if not itens_selecionados: QMessageBox.information(self, "Aviso", "Por favor, selecione um cliente na tabela para editar."); return
        linha_selecionada = itens_selecionados[0].row()
        cliente_data = self.tabela_clientes.item(linha_selecionada, 0).data(Qt.UserRole)
        dialog = DialogoCliente(self.usuario_logado, cliente_id=cliente_data['id'], parent=self)
        dialog.nome_edit.setText(cliente_data['nome']); dialog.tipo_envio_combo.setCurrentText(cliente_data['tipo_envio']); dialog.contato_edit.setText(cliente_data['contato'])
        dialog.gera_recibo_check.setChecked(bool(cliente_data['gera_recibo'])); dialog.conta_xmls_check.setChecked(bool(cliente_data['conta_xmls']))
        dialog.nivel_edit.setText(cliente_data['nivel']); dialog.detalhes_edit.setText(cliente_data['outros_detalhes'])
        if dialog.exec_() == QDialog.Accepted: self.carregar_clientes()
    def excluir_cliente(self):
        itens_selecionados = self.tabela_clientes.selectedItems()
        if not itens_selecionados: QMessageBox.information(self, "Aviso", "Por favor, selecione um cliente para excluir."); return
        linha_selecionada = itens_selecionados[0].row()
        cliente_data = self.tabela_clientes.item(linha_selecionada, 0).data(Qt.UserRole)
        reply = QMessageBox.question(self, "Confirmar Exclusão", f"Tem certeza que deseja excluir o cliente '{cliente_data['nome']}'?\nTODAS as entregas associadas a ele também serão excluídas.", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes: database.deletar_cliente(cliente_data['id'], self.usuario_logado['username']); self.carregar_clientes()
class FormularioStatusDialog(QDialog):
    def __init__(self, status=None, parent=None):
        super().__init__(parent); self.status = status; self.setWindowTitle("Novo Status" if not status else "Editar Status"); layout = QFormLayout(self); self.nome_edit = QLineEdit(); self.cor_btn = QPushButton("Escolher Cor"); self.cor_label = QLabel(); self.cor_label.setFixedSize(20, 20); cor_layout = QHBoxLayout(); cor_layout.addWidget(self.cor_btn); cor_layout.addWidget(self.cor_label); layout.addRow("Nome:", self.nome_edit); layout.addRow("Cor:", cor_layout)
        if status: self.nome_edit.setText(status['nome']); self.cor_hex = status['cor_hex']
        else: self.cor_hex = "#ffffff"
        self.atualizar_label_cor(); self.cor_btn.clicked.connect(self.escolher_cor); botoes = QHBoxLayout(); salvar_btn = QPushButton("Salvar"); cancelar_btn = QPushButton("Cancelar"); botoes.addStretch(); botoes.addWidget(salvar_btn); botoes.addWidget(cancelar_btn); layout.addRow(botoes); salvar_btn.clicked.connect(self.accept); cancelar_btn.clicked.connect(self.reject)
    def escolher_cor(self):
        cor = QColorDialog.getColor(QColor(self.cor_hex), self)
        if cor.isValid(): self.cor_hex = cor.name(); self.atualizar_label_cor()
    def atualizar_label_cor(self): self.cor_label.setStyleSheet(f"background-color: {self.cor_hex}; border: 1px solid black;")
    def get_data(self): return self.nome_edit.text().strip(), self.cor_hex
class StatusDialog(QDialog):
    def __init__(self, usuario_logado, parent=None):
        super().__init__(parent); self.usuario_logado = usuario_logado; self.setWindowTitle("Gerenciar Status"); self.setMinimumSize(500, 400); layout = QVBoxLayout(self); self.tabela_status = QTableWidget(); self.tabela_status.setColumnCount(2); self.tabela_status.setHorizontalHeaderLabels(["Nome do Status", "Cor"]); self.tabela_status.setEditTriggers(QTableWidget.NoEditTriggers); self.tabela_status.setSelectionBehavior(QTableWidget.SelectRows); self.tabela_status.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch); layout.addWidget(self.tabela_status); botoes = QHBoxLayout(); add_btn = QPushButton("Adicionar"); edit_btn = QPushButton("Editar"); del_btn = QPushButton("Excluir"); botoes.addStretch(); botoes.addWidget(add_btn); botoes.addWidget(edit_btn); botoes.addWidget(del_btn); layout.addLayout(botoes); add_btn.clicked.connect(self.adicionar); edit_btn.clicked.connect(self.editar); del_btn.clicked.connect(self.excluir); self.carregar_status()
    def carregar_status(self):
        self.tabela_status.setRowCount(0)
        for status in database.listar_status():
            row = self.tabela_status.rowCount(); self.tabela_status.insertRow(row); item_nome = QTableWidgetItem(status['nome']); item_nome.setData(Qt.UserRole, dict(status)); item_cor = QTableWidgetItem(); item_cor.setBackground(QColor(status['cor_hex'])); self.tabela_status.setItem(row, 0, item_nome); self.tabela_status.setItem(row, 1, item_cor)
    def adicionar(self):
        dialog = FormularioStatusDialog(parent=self)
        if dialog.exec_() == QDialog.Accepted: nome, cor_hex = dialog.get_data();
        if nome: database.adicionar_status(nome, cor_hex, self.usuario_logado['username']); self.carregar_status(); self.parent().populate_calendar()
    def editar(self):
        linha = self.tabela_status.currentRow();
        if linha < 0: return;
        status_data = self.tabela_status.item(linha, 0).data(Qt.UserRole); dialog = FormularioStatusDialog(status=status_data, parent=self)
        if dialog.exec_() == QDialog.Accepted: nome, cor_hex = dialog.get_data();
        if nome: database.atualizar_status(status_data['id'], nome, cor_hex, self.usuario_logado['username']); self.carregar_status(); self.parent().populate_calendar()
    def excluir(self):
        linha = self.tabela_status.currentRow();
        if linha < 0: return;
        status_data = self.tabela_status.item(linha, 0).data(Qt.UserRole); reply = QMessageBox.question(self, "Confirmar Exclusão", f"Tem certeza que deseja excluir o status '{status_data['nome']}'?\nOs agendamentos que usam este status ficarão sem categoria.", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes: database.deletar_status(status_data['id'], self.usuario_logado['username']); self.carregar_status(); self.parent().populate_calendar()
# main.py

class EntregaDialog(QDialog):
    # SUBSTITUA TODA A SUA CLASSE POR ESTA VERSÃO ATUALIZADA
    def __init__(self, usuario_logado, date, entrega_data=None, parent=None):
        super().__init__(parent)
        self.entrega_data = entrega_data
        self.usuario_logado = usuario_logado
        self.date = date  # Armazena a data para usar no rascunho

        self.setWindowTitle("Agendar Nova Entrega" if not entrega_data else "Editar Entrega")
        layout = QFormLayout(self)

        # --- Definição dos Componentes ---
        self.cliente_combo = QComboBox()
        self.status_combo = QComboBox()
        self.responsavel_edit = QLineEdit()
        self.observacoes_edit = QTextEdit()
        self.rascunho_edit = QLineEdit()
        self.rascunho_edit.setReadOnly(True)
        self.rascunho_edit.setPlaceholderText("Selecione um cliente e status...")
        
        copiar_btn = QPushButton("Copiar")
        rascunho_layout = QHBoxLayout()
        rascunho_layout.addWidget(self.rascunho_edit)
        rascunho_layout.addWidget(copiar_btn)

        self.salvar_btn = QPushButton("Salvar")
        self.cancelar_btn = QPushButton("Cancelar")

        # --- Lógica de Preenchimento e Configuração ---
        self.responsavel_edit.setText(self.usuario_logado['username'])
        self.carregar_combos() # Carrega os dados e verifica se há clientes

        # --- Adiciona os componentes ao layout da janela ---
        layout.addRow("Cliente:", self.cliente_combo)
        layout.addRow("Status:", self.status_combo)
        layout.addRow("Responsável:", self.responsavel_edit)
        layout.addRow("Rascunho:", rascunho_layout)
        layout.addRow("Observações:", self.observacoes_edit)

        if entrega_data:
            # Carrega dados existentes se estiver no modo de edição
            index_cliente = self.cliente_combo.findData(entrega_data['cliente_id'])
            if index_cliente > -1: self.cliente_combo.setCurrentIndex(index_cliente)
            
            index_status = self.status_combo.findData(entrega_data['status_id'])
            if index_status > -1: self.status_combo.setCurrentIndex(index_status)
            
            if entrega_data.get('responsavel'):
                 self.responsavel_edit.setText(entrega_data['responsavel'])
            self.observacoes_edit.setText(entrega_data['observacoes'])
        else:
            # Define o status "PENDENTE" como padrão para novos agendamentos
            index_pendente = self.status_combo.findText("PENDENTE")
            if index_pendente > -1: self.status_combo.setCurrentIndex(index_pendente)
            
        botoes_layout = QHBoxLayout()
        botoes_layout.addStretch()
        botoes_layout.addWidget(self.salvar_btn)
        botoes_layout.addWidget(self.cancelar_btn)
        layout.addRow(botoes_layout)
        
        # --- Conexão dos Sinais (Eventos) ---
        self.salvar_btn.clicked.connect(self.accept)
        self.cancelar_btn.clicked.connect(self.reject)
        copiar_btn.clicked.connect(self.copiar_rascunho)
        self.cliente_combo.currentIndexChanged.connect(self.atualizar_rascunho)
        self.status_combo.currentIndexChanged.connect(self.atualizar_rascunho)

        self.atualizar_rascunho() # Gera o texto do rascunho ao iniciar

    def carregar_combos(self):
        # Refatorado para legibilidade e robustez
        self.cliente_combo.clear()
        self.status_combo.clear()

        clientes = database.listar_clientes()
        status_list = database.listar_status()

        if not clientes:
            self.cliente_combo.addItem("Nenhum cliente cadastrado")
            self.cliente_combo.setEnabled(False)
            self.salvar_btn.setEnabled(False)
            self.rascunho_edit.setEnabled(False)
            self.observacoes_edit.setPlaceholderText("É necessário cadastrar um cliente antes de criar um agendamento.")
        else:
            for cliente in clientes:
                self.cliente_combo.addItem(cliente['nome'], cliente['id'])

        for status in status_list:
            self.status_combo.addItem(status['nome'], status['id'])

    def get_data(self):
        # Refatorado para melhor legibilidade
        return {
            "cliente_id": self.cliente_combo.currentData(),
            "status_id": self.status_combo.currentData(),
            "responsavel": self.responsavel_edit.text().strip(),
            "observacoes": self.observacoes_edit.toPlainText().strip()
        }
    
    def atualizar_rascunho(self):
        """Gera a string do rascunho com base nos campos atuais."""
        if not self.cliente_combo.isEnabled():
            return

        nome_cliente = self.cliente_combo.currentText()
        nome_status = self.status_combo.currentText()
        data_str = self.date.toString("MM/yyyy")
        
        if nome_status.lower() == 'retificado':
            texto_rascunho = f"Sintegra Retificado-{data_str}-{nome_cliente}"
        else:
            texto_rascunho = f"Sintegra -{data_str}-{nome_cliente}"
            
        self.rascunho_edit.setText(texto_rascunho)

    def copiar_rascunho(self):
        """Copia o texto do rascunho para a área de transferência."""
        texto_para_copiar = self.rascunho_edit.text()
        if texto_para_copiar:
            clipboard = QApplication.clipboard()
            clipboard.setText(texto_para_copiar)
            
            botao = self.sender()
            botao.setText("Copiado!")
            QTimer.singleShot(1500, lambda: botao.setText("Copiar"))
class DayViewDialog(QDialog):
    def __init__(self, date, usuario_logado, parent=None):
        super().__init__(parent); self.date = date; self.usuario_logado = usuario_logado; self.setWindowTitle(f"Agenda para {date.toString('dd/MM/yyyy')}"); self.setMinimumSize(800, 600)
        self.horarios = self.gerar_horarios_dinamicos()
        layout = QVBoxLayout(self); self.tabela_agenda = QTableWidget(); self.tabela_agenda.setColumnCount(5); self.tabela_agenda.setHorizontalHeaderLabels(["Horário", "Cliente", "Contato", "Tipo Envio", "Status"]); self.tabela_agenda.setEditTriggers(QTableWidget.NoEditTriggers); self.tabela_agenda.setSelectionBehavior(QTableWidget.SelectRows); self.tabela_agenda.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch); self.tabela_agenda.setColumnWidth(0, 80); layout.addWidget(self.tabela_agenda); botoes_layout = QHBoxLayout(); edit_btn = QPushButton("Editar Agendamento"); del_btn = QPushButton("Excluir Agendamento"); botoes_layout.addStretch(); botoes_layout.addWidget(edit_btn); botoes_layout.addWidget(del_btn); layout.addLayout(botoes_layout); edit_btn.clicked.connect(self.editar_agendamento); del_btn.clicked.connect(self.excluir_agendamento); self.tabela_agenda.cellDoubleClicked.connect(self.gerenciar_agendamento_duplo_clique); self.carregar_agenda_dia()
    def gerar_horarios_dinamicos(self):
        settings = QSettings()
        modo = settings.value("horarios/modo", "automatico")
        if modo == "manual":
            lista_manual = settings.value("horarios/lista_manual", "09:00,10:00,11:00")
            horarios_validos = []
            for horario_str in lista_manual.split(','):
                horario_str = horario_str.strip()
                try:
                    time_obj = QTime.fromString(horario_str, 'HH:mm')
                    if time_obj.isValid(): horarios_validos.append(horario_str)
                except: pass
            return sorted(horarios_validos)
        else:
            hora_inicio = QTime.fromString(settings.value("horarios/hora_inicio", "08:30"), 'HH:mm')
            hora_fim = QTime.fromString(settings.value("horarios/hora_fim", "17:30"), 'HH:mm')
            intervalo = settings.value("horarios/intervalo_minutos", 30, type=int)
            horarios = []; hora_atual = hora_inicio
            while hora_atual <= hora_fim:
                horarios.append(hora_atual.toString('HH:mm'))
                hora_atual = hora_atual.addSecs(intervalo * 60)
            return horarios
    def carregar_agenda_dia(self):
        self.tabela_agenda.setRowCount(0); self.tabela_agenda.setRowCount(len(self.horarios)); agendamentos_dia = database.get_entregas_por_dia(self.date.toString("yyyy-MM-dd"))
        for i, horario in enumerate(self.horarios):
            item_horario = QTableWidgetItem(horario); item_horario.setTextAlignment(Qt.AlignCenter); self.tabela_agenda.setItem(i, 0, item_horario)
            if horario in agendamentos_dia:
                agendamento = agendamentos_dia[horario]; self.tabela_agenda.setItem(i, 1, QTableWidgetItem(agendamento['nome_cliente'])); self.tabela_agenda.setItem(i, 2, QTableWidgetItem(agendamento['contato'])); self.tabela_agenda.setItem(i, 3, QTableWidgetItem(agendamento['tipo_envio'])); item_status = QTableWidgetItem(agendamento['nome_status']); item_status.setBackground(QColor(agendamento['cor_hex']))
                self.tabela_agenda.setItem(i, 4, item_status); item_horario.setData(Qt.UserRole, agendamento)
            else:
                for j in range(1, 5): self.tabela_agenda.setItem(i, j, QTableWidgetItem(""))
    def gerenciar_agendamento_duplo_clique(self, row, column):
        item_horario = self.tabela_agenda.item(row, 0)
        agendamento_existente = item_horario.data(Qt.UserRole)
        if agendamento_existente:
            self.editar_agendamento()
        else:
            horario = item_horario.text()
            dialog = EntregaDialog(self.usuario_logado, self.date, parent=self)
            if dialog.exec_() == QDialog.Accepted:
                data = dialog.get_data()
                database.adicionar_entrega(self.date.toString("yyyy-MM-dd"), horario, data['status_id'], data['cliente_id'], data['responsavel'], data['observacoes'], self.usuario_logado['username'])
        self.carregar_agenda_dia(); self.parent().populate_calendar()
    def editar_agendamento(self):
        itens_selecionados = self.tabela_agenda.selectedItems()
        if not itens_selecionados: QMessageBox.information(self, "Aviso", "Por favor, selecione um agendamento na tabela para editar."); return
        linha_selecionada = itens_selecionados[0].row()
        item_horario = self.tabela_agenda.item(linha_selecionada, 0)
        agendamento_existente = item_horario.data(Qt.UserRole)
        if not agendamento_existente: QMessageBox.information(self, "Aviso", "Este horário está vago. Dê um clique duplo para criar um novo agendamento."); return
        dialog = EntregaDialog(self.usuario_logado, self.date, entrega_data=agendamento_existente, parent=self)
        if dialog.exec_() == QDialog.Accepted:
            data = dialog.get_data()
            database.atualizar_entrega(agendamento_existente['id'], agendamento_existente['horario'], data['status_id'], data['cliente_id'], data['responsavel'], data['observacoes'], self.usuario_logado['username'])
            self.carregar_agenda_dia(); self.parent().populate_calendar()
    def excluir_agendamento(self):
        itens_selecionados = self.tabela_agenda.selectedItems()
        if not itens_selecionados: QMessageBox.information(self, "Aviso", "Por favor, selecione um agendamento na tabela para excluir."); return
        linha_selecionada = itens_selecionados[0].row()
        item_horario = self.tabela_agenda.item(linha_selecionada, 0)
        agendamento_existente = item_horario.data(Qt.UserRole)
        if not agendamento_existente: QMessageBox.information(self, "Aviso", "Não há agendamento para excluir neste horário."); return
        reply = QMessageBox.question(self, "Confirmar Exclusão", f"Tem certeza que deseja excluir o agendamento para '{agendamento_existente['nome_cliente']}' às {agendamento_existente['horario']}?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            database.deletar_entrega(agendamento_existente['id'], self.usuario_logado['username'])
            self.carregar_agenda_dia(); self.parent().populate_calendar()
class ConfigDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent); self.setWindowTitle("Configurações"); self.settings = QSettings(); main_layout = QVBoxLayout(self); modo_groupbox = QGroupBox("Modo de Definição de Horários"); modo_layout = QVBoxLayout(); self.radio_auto = QRadioButton("Gerar horários automaticamente"); self.radio_manual = QRadioButton("Inserir horários manualmente"); modo_layout.addWidget(self.radio_auto); modo_layout.addWidget(self.radio_manual); modo_groupbox.setLayout(modo_layout); self.auto_groupbox = QGroupBox("Configuração Automática"); auto_layout = QFormLayout(); self.inicio_edit = QTimeEdit(); self.fim_edit = QTimeEdit(); self.intervalo_spin = QSpinBox(); self.intervalo_spin.setRange(5, 120); self.intervalo_spin.setSuffix(" minutos"); auto_layout.addRow("Hora de Início:", self.inicio_edit); auto_layout.addRow("Hora de Fim:", self.fim_edit); auto_layout.addRow("Intervalo:", self.intervalo_spin); self.auto_groupbox.setLayout(auto_layout); self.manual_groupbox = QGroupBox("Configuração Manual"); manual_layout = QFormLayout(); self.lista_manual_edit = QLineEdit(); self.lista_manual_edit.setPlaceholderText("Ex: 08:00, 09:15, 10:30, 14:00"); manual_layout.addRow("Lista de horários (separados por vírgula):", self.lista_manual_edit); self.manual_groupbox.setLayout(manual_layout); geral_groupbox = QGroupBox("Geral"); geral_layout = QFormLayout(); self.lembrete_spin = QSpinBox(); self.lembrete_spin.setRange(1, 60); self.lembrete_spin.setSuffix(" minutos"); geral_layout.addRow("Avisar com antecedência de:", self.lembrete_spin); geral_groupbox.setLayout(geral_layout); main_layout.addWidget(modo_groupbox); main_layout.addWidget(self.auto_groupbox); main_layout.addWidget(self.manual_groupbox); main_layout.addWidget(geral_groupbox); botoes = QHBoxLayout(); salvar_btn = QPushButton("Salvar"); cancelar_btn = QPushButton("Cancelar"); botoes.addStretch(); botoes.addWidget(salvar_btn); botoes.addWidget(cancelar_btn); main_layout.addLayout(botoes); salvar_btn.clicked.connect(self.salvar); cancelar_btn.clicked.connect(self.reject); self.radio_auto.toggled.connect(self.atualizar_modo_visivel); self.carregar_configs()
    def carregar_configs(self):
        modo = self.settings.value("horarios/modo", "automatico")
        if modo == "manual": self.radio_manual.setChecked(True)
        else: self.radio_auto.setChecked(True)
        self.inicio_edit.setTime(QTime.fromString(self.settings.value("horarios/hora_inicio", "08:30"), 'HH:mm')); self.fim_edit.setTime(QTime.fromString(self.settings.value("horarios/hora_fim", "17:30"), 'HH:mm')); self.intervalo_spin.setValue(self.settings.value("horarios/intervalo_minutos", 30, type=int)); self.lista_manual_edit.setText(self.settings.value("horarios/lista_manual", "09:00,10:00,11:00")); self.lembrete_spin.setValue(self.settings.value("geral/minutos_lembrete", 15, type=int)); self.atualizar_modo_visivel()
    def atualizar_modo_visivel(self):
        if self.radio_auto.isChecked(): self.auto_groupbox.setEnabled(True); self.manual_groupbox.setEnabled(False)
        else: self.auto_groupbox.setEnabled(False); self.manual_groupbox.setEnabled(True)
    def salvar(self):
        if self.radio_auto.isChecked():
            self.settings.setValue("horarios/modo", "automatico"); self.settings.setValue("horarios/hora_inicio", self.inicio_edit.time().toString('HH:mm')); self.settings.setValue("horarios/hora_fim", self.fim_edit.time().toString('HH:mm')); self.settings.setValue("horarios/intervalo_minutos", self.intervalo_spin.value())
        else:
            self.settings.setValue("horarios/modo", "manual"); self.settings.setValue("horarios/lista_manual", self.lista_manual_edit.text())
        self.settings.setValue("geral/minutos_lembrete", self.lembrete_spin.value()); QMessageBox.information(self, "Salvo", "Configurações salvas. Por favor, reinicie o programa para que todas as alterações tenham efeito."); self.accept()

class RelatorioDialog(QDialog):
    def __init__(self, usuario_logado, parent=None):
        super().__init__(parent); self.usuario_logado = usuario_logado; self.setWindowTitle("Gerar Relatórios"); self.setMinimumWidth(400)
        layout = QVBoxLayout(self); form_layout = QFormLayout(); self.tipo_relatorio_combo = QComboBox(); self.tipo_relatorio_combo.addItems(["Relatório de Agendamentos", "Relatório de Logs de Atividade"]); self.data_inicio_edit = QDateEdit(QDate.currentDate().addDays(-30)); self.data_inicio_edit.setCalendarPopup(True); self.data_fim_edit = QDateEdit(QDate.currentDate()); self.data_fim_edit.setCalendarPopup(True); self.filtros_agendamento_group = QGroupBox("Filtros de Agendamento"); filtros_agendamento_layout = QVBoxLayout(); self.status_list_widget = QListWidget(); self.status_list_widget.setSelectionMode(QListWidget.MultiSelection); filtros_agendamento_layout.addWidget(QLabel("Filtrar por Status (deixe sem selecionar para incluir todos):")); filtros_agendamento_layout.addWidget(self.status_list_widget); self.filtros_agendamento_group.setLayout(filtros_agendamento_layout); self.filtros_logs_group = QGroupBox("Filtros de Logs"); filtros_logs_layout = QFormLayout(); self.usuario_combo = QComboBox(); filtros_logs_layout.addRow("Filtrar por Usuário:", self.usuario_combo); self.filtros_logs_group.setLayout(filtros_logs_layout)
        self.carregar_filtros(); form_layout.addRow("Tipo de Relatório:", self.tipo_relatorio_combo); form_layout.addRow("Data de Início:", self.data_inicio_edit); form_layout.addRow("Data de Fim:", self.data_fim_edit); layout.addLayout(form_layout); layout.addWidget(self.filtros_agendamento_group); layout.addWidget(self.filtros_logs_group); gerar_btn = QPushButton("Gerar Relatório"); layout.addWidget(gerar_btn, alignment=Qt.AlignCenter); gerar_btn.clicked.connect(self.gerar_relatorio); self.tipo_relatorio_combo.currentIndexChanged.connect(self.atualizar_filtros_visiveis); self.atualizar_filtros_visiveis()
    def carregar_filtros(self):
        for status in database.listar_status():
            item = QListWidgetItem(status['nome']); item.setData(Qt.UserRole, status['id']); self.status_list_widget.addItem(item)
        self.usuario_combo.addItem("Todos")
        for username in database.listar_usuarios(): self.usuario_combo.addItem(username)
    def atualizar_filtros_visiveis(self):
        if "Agendamentos" in self.tipo_relatorio_combo.currentText():
            self.filtros_agendamento_group.setVisible(True); self.filtros_logs_group.setVisible(False)
        else:
            self.filtros_agendamento_group.setVisible(False); self.filtros_logs_group.setVisible(True)
    def gerar_relatorio(self):
        data_inicio = self.data_inicio_edit.date().toString("yyyy-MM-dd"); data_fim = self.data_fim_edit.date().toString("yyyy-MM-dd"); tipo_relatorio = self.tipo_relatorio_combo.currentText()
        dialog = QFileDialog(self); dialog.setAcceptMode(QFileDialog.AcceptSave); dialog.setNameFilter("Arquivos PDF (*.pdf);;Arquivos CSV (*.csv)"); dialog.setDefaultSuffix("pdf")
        if "Agendamentos" in tipo_relatorio:
            dialog.selectFile(f"Relatorio_Agendamentos_{data_inicio}_a_{data_fim}"); status_selecionados_ids = [item.data(Qt.UserRole) for item in self.status_list_widget.selectedItems()]; dados = database.get_entregas_filtradas(data_inicio, data_fim, status_selecionados_ids)
            titulo = f"Relatório de Agendamentos de {self.data_inicio_edit.date().toString('dd/MM/yyyy')} a {self.data_fim_edit.date().toString('dd/MM/yyyy')}"
            if not dados: QMessageBox.information(self, "Aviso", "Nenhum agendamento encontrado para os filtros selecionados."); return
            if dialog.exec_():
                nome_arquivo = dialog.selectedFiles()[0]
                if nome_arquivo.endswith(".pdf"): export.exportar_para_pdf(dados, nome_arquivo, titulo)
                elif nome_arquivo.endswith(".csv"): export.exportar_para_csv(dados, nome_arquivo)
        else:
            dialog.selectFile(f"Relatorio_Logs_{data_inicio}_a_{data_fim}"); usuario_selecionado = self.usuario_combo.currentText(); dados = database.get_logs_filtrados(data_inicio, data_fim, usuario_selecionado)
            titulo = f"Relatório de Logs de {self.data_inicio_edit.date().toString('dd/MM/yyyy')} a {self.data_fim_edit.date().toString('dd/MM/yyyy')}"
            if not dados: QMessageBox.information(self, "Aviso", "Nenhum log encontrado para os filtros selecionados."); return
            if dialog.exec_():
                nome_arquivo = dialog.selectedFiles()[0]
                if nome_arquivo.endswith(".pdf"): export.exportar_logs_pdf(dados, nome_arquivo, titulo)
                else: export.exportar_logs_csv(dados, nome_arquivo)
        self.accept()

# --- Janela Principal ---

class DialogoUsuario(QDialog):
    def __init__(self, parent=None, usuario=None):
        super().__init__(parent)
        self.setWindowTitle("Usuário")
        self.usuario = usuario
        layout = QFormLayout(self)

        self.username_edit = QLineEdit()
        self.password_edit = QLineEdit()
        self.password_edit.setEchoMode(QLineEdit.Password)

        if usuario:
            self.username_edit.setText(usuario['username'])

        layout.addRow("Usuário:", self.username_edit)
        layout.addRow("Senha:", self.password_edit)

        botoes = QHBoxLayout()
        salvar_btn = QPushButton("Salvar")
        salvar_btn.clicked.connect(self.accept)
        cancelar_btn = QPushButton("Cancelar")
        cancelar_btn.clicked.connect(self.reject)
        botoes.addWidget(salvar_btn)
        botoes.addWidget(cancelar_btn)

        layout.addRow(botoes)

    def get_dados(self):
        return self.username_edit.text(), self.password_edit.text()

class ConfirmacaoSenhaDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Confirmação de Administrador")
        self.setModal(True) # Trava a janela principal enquanto esta estiver aberta
        layout = QVBoxLayout(self)

        label = QLabel("Para prosseguir, por favor, digite sua senha de administrador:")
        self.senha_edit = QLineEdit()
        self.senha_edit.setEchoMode(QLineEdit.Password) # Esconde a senha

        layout.addWidget(label)
        layout.addWidget(self.senha_edit)

        botoes_layout = QHBoxLayout()
        ok_btn = QPushButton("Confirmar")
        cancelar_btn = QPushButton("Cancelar")

        botoes_layout.addStretch()
        botoes_layout.addWidget(ok_btn)
        botoes_layout.addWidget(cancelar_btn)
        
        layout.addLayout(botoes_layout)

        ok_btn.clicked.connect(self.accept)
        cancelar_btn.clicked.connect(self.reject)

    def get_senha(self):
        """Retorna a senha digitada pelo usuário."""
        return self.senha_edit.text()


class JanelaUsuarios(QDialog):
    def __init__(self, usuario_logado, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Gerenciar Usuários")
        self.usuario_logado = usuario_logado
        self.resize(400, 300)

        layout = QVBoxLayout(self)

        self.tabela = QTableWidget()
        self.tabela.setColumnCount(1)
        self.tabela.setHorizontalHeaderLabels(["Usuário"])
        self.tabela.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.tabela)

        botoes = QHBoxLayout()
        adicionar_btn = QPushButton("Adicionar")
        editar_btn = QPushButton("Editar")
        excluir_btn = QPushButton("Excluir")

        adicionar_btn.clicked.connect(self.adicionar_usuario)
        editar_btn.clicked.connect(self.editar_usuario)
        excluir_btn.clicked.connect(self.excluir_usuario)

        botoes.addWidget(adicionar_btn)
        botoes.addWidget(editar_btn)
        botoes.addWidget(excluir_btn)

        layout.addLayout(botoes)

        self.carregar_usuarios()

    def carregar_usuarios(self):
        usuarios = database.listar_usuarios()
        self.tabela.setRowCount(len(usuarios))
        for i, username in enumerate(usuarios):
            self.tabela.setItem(i, 0, QTableWidgetItem(username))

    def adicionar_usuario(self):
        dialog = DialogoUsuario(self)
        if dialog.exec_() == QDialog.Accepted:
            username, password = dialog.get_dados()
            if username and password:
                if database.criar_usuario(username, password):
                    QMessageBox.information(self, "Sucesso", "Usuário criado!")
                    self.carregar_usuarios()
                else:
                    QMessageBox.warning(self, "Erro", "Nome de usuário já existe!")

    def editar_usuario(self):
        linha = self.tabela.currentRow()
        if linha < 0:
            QMessageBox.warning(self, "Ação Necessária", "Por favor, selecione um usuário para editar.")
            return
            
        username_selecionado = self.tabela.item(linha, 0).text()

        # --- INÍCIO DA NOVA LÓGICA DE VERIFICAÇÃO DE SENHA ---
        
        # 1. Pede a senha do administrador logado
        dialogo_confirmacao = ConfirmacaoSenhaDialog(self)
        if dialogo_confirmacao.exec_() == QDialog.Accepted:
            senha_digitada = dialogo_confirmacao.get_senha()
            administrador_logado = self.usuario_logado['username']

            # 2. Verifica se a senha está correta usando a função do database
            if database.verificar_senha_usuario_atual(administrador_logado, senha_digitada):
                # 3. Se a senha estiver correta, continua com a edição
                dialog = DialogoUsuario(self, usuario={"username": username_selecionado})
                if dialog.exec_() == QDialog.Accepted:
                    novo_username, nova_senha = dialog.get_dados()
                    
                    if not nova_senha:
                        QMessageBox.warning(self, "Senha Obrigatória", "O campo de senha não pode estar vazio ao editar.")
                        return

                    if novo_username and nova_senha:
                        conn = database.conectar()
                        usuario = conn.execute("SELECT * FROM usuarios WHERE username=?", (username_selecionado,)).fetchone()
                        conn.close()
                        if usuario:
                            database.atualizar_usuario(usuario["id"], novo_username, nova_senha, self.usuario_logado["username"])
                            QMessageBox.information(self, "Sucesso", "Usuário atualizado!")
                            self.carregar_usuarios()
            else:
                # Se a senha estiver incorreta, exibe um erro
                QMessageBox.critical(self, "Falha na Autenticação", "Senha de administrador incorreta. Ação cancelada.")

    def excluir_usuario(self):
        linha = self.tabela.currentRow()
        if linha < 0:
            QMessageBox.warning(self, "Ação Necessária", "Por favor, selecione um usuário para excluir.")
            return

        # 1. Obter dados do usuário selecionado para pegar o ID
        username_selecionado = self.tabela.item(linha, 0).text()
        conn = database.conectar()
        usuario_para_excluir = conn.execute("SELECT * FROM usuarios WHERE username=?", (username_selecionado,)).fetchone()
        conn.close()

        if not usuario_para_excluir:
            QMessageBox.critical(self, "Erro", "Não foi possível encontrar o usuário no banco de dados.")
            return

        # 2. VERIFICAÇÃO DE SEGURANÇA: Proteger o admin principal (ID 1)
        if usuario_para_excluir['id'] == 1:
            QMessageBox.warning(self, "Ação Proibida", "O usuário administrador principal não pode ser excluído.")
            return

        # 3. VERIFICAÇÃO DE SEGURANÇA: Impedir autoexclusão
        if usuario_para_excluir['id'] == self.usuario_logado['id']:
            QMessageBox.warning(self, "Ação Proibida", "Você não pode excluir seu próprio usuário.")
            return

        # 4. Confirmação de senha do administrador logado
        dialogo_confirmacao = ConfirmacaoSenhaDialog(self)
        if dialogo_confirmacao.exec_() == QDialog.Accepted:
            senha_digitada = dialogo_confirmacao.get_senha()
            administrador_logado = self.usuario_logado['username']

            if database.verificar_senha_usuario_atual(administrador_logado, senha_digitada):
                # 5. Se tudo estiver correto, exibir confirmação final e proceder com a exclusão
                reply = QMessageBox.question(self, "Confirmar Exclusão", 
                                             f"Tem certeza que deseja excluir permanentemente o usuário '{username_selecionado}'?",
                                             QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                
                if reply == QMessageBox.Yes:
                    if database.deletar_usuario(usuario_para_excluir["id"], self.usuario_logado["username"]):
                        QMessageBox.information(self, "Sucesso", "Usuário excluído!")
                        self.carregar_usuarios()
                    else:
                        # Esta mensagem apareceria se o database.py falhasse (por ex, na trava do ID 1)
                        QMessageBox.critical(self, "Erro", "Ocorreu um erro ao excluir o usuário.")
            else:
                QMessageBox.critical(self, "Falha na Autenticação", "Senha de administrador incorreta. Ação cancelada.")


class CalendarWindow(QMainWindow):
    def __init__(self, usuario):
        super().__init__(); self.usuario_atual = usuario; titulo = f"Agendador Mensal - Bem-vindo, {self.usuario_atual['username']}!"; self.setWindowTitle(titulo); self.setGeometry(100, 100, 1100, 800); self.current_date = datetime.now(); self.central_widget = QWidget(); self.setCentralWidget(self.central_widget); self.main_layout = QVBoxLayout(self.central_widget); self.notificados_nesta_sessao = set(); self.feriados = {}; self.setup_ui(); self.setup_tray_icon(); self.setup_timer_notificacoes(); self.populate_calendar()
    # main.py
    def setup_ui(self):
        # Layout de navegação (mês anterior/próximo)
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

        # Label do dashboard
        self.dashboard_label = QLabel()
        self.dashboard_label.setAlignment(Qt.AlignCenter)
        self.main_layout.addWidget(self.dashboard_label)

        # Cabeçalho dos dias da semana
        header_layout = QGridLayout()
        dias = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]
        for i, dia in enumerate(dias):
            header_layout.addWidget(QLabel(f"<b>{dia}</b>", alignment=Qt.AlignCenter), 0, i)
        self.main_layout.addLayout(header_layout)
        
        # Grid do calendário
        self.calendar_grid = QGridLayout()
        self.calendar_grid.setSpacing(0)
        self.main_layout.addLayout(self.calendar_grid)

        # --- CÓDIGO AJUSTADO ---
        # Layout para os botões de ação na parte inferior
        action_layout = QHBoxLayout()

        # Criação de todos os botões de ação
        config_btn = QPushButton("Configurações")
        config_btn.clicked.connect(self.abrir_configuracoes)
        
        clientes_btn = QPushButton("Gerenciar Clientes")
        clientes_btn.clicked.connect(self.gerenciar_clientes)
        
        status_btn = QPushButton("Gerenciar Status")
        status_btn.clicked.connect(self.manage_status)
        
        usuarios_btn = QPushButton("Gerenciar Usuários")
        usuarios_btn.clicked.connect(self.gerenciar_usuarios)
        
        relatorio_btn = QPushButton("Gerar Relatório")
        relatorio_btn.clicked.connect(self.abrir_dialogo_relatorio)

        # Adicionando os botões ao layout na ordem desejada
        action_layout.addWidget(config_btn)
        action_layout.addStretch() # Adiciona um espaço flexível no meio
        action_layout.addWidget(clientes_btn)
        action_layout.addWidget(status_btn)
        action_layout.addWidget(usuarios_btn) # Botão agora é adicionado aqui
        action_layout.addWidget(relatorio_btn)
        
        # Adiciona o layout dos botões ao layout principal da janela
        self.main_layout.addLayout(action_layout)
    def setup_tray_icon(self):
        self.tray_icon = QSystemTrayIcon(self)
        caminho_icone = os.path.join('imagens', 'icon.png')
        icon = QIcon(caminho_icone)
        if icon.isNull(): self.tray_icon.setIcon(self.style().standardIcon(QStyle.SP_MessageBoxInformation))
        else: self.tray_icon.setIcon(icon)
        self.tray_icon.setToolTip("Agendador de Entregas"); self.tray_icon.show()
    def setup_timer_notificacoes(self):
        settings = QSettings(); minutos_lembrete = settings.value("geral/minutos_lembrete", 15, type=int)
        self.timer = QTimer(self); self.timer.timeout.connect(lambda: self.verificar_agendamentos_proximos(minutos_lembrete)); self.timer.start(60000)
    def verificar_agendamentos_proximos(self, minutos_antecedencia):
        agora = datetime.now(); limite = agora + timedelta(minutes=minutos_antecedencia); data_hoje = agora.strftime("%Y-%m-%d"); hora_inicio = agora.strftime("%H:%M"); hora_fim = limite.strftime("%H:%M")
        agendamentos = database.get_entregas_no_intervalo(data_hoje, hora_inicio, hora_fim)
        for ag in agendamentos:
            if ag['id'] not in self.notificados_nesta_sessao:
                titulo = f"Lembrete de Agendamento ({ag['horario']})"; mensagem = f"Cliente: {ag['nome_cliente']}\nStatus: {ag.get('nome_status', 'N/A')}"
                self.tray_icon.showMessage(titulo, mensagem, QSystemTrayIcon.Information, 15000); self.notificados_nesta_sessao.add(ag['id'])
    def closeEvent(self, event): self.tray_icon.hide(); event.accept()
    def populate_calendar(self):
        for i in reversed(range(self.calendar_grid.count())): self.calendar_grid.itemAt(i).widget().setParent(None)
        year = self.current_date.year; month = self.current_date.month
        self.month_label.setText(f"<b>{self.current_date.strftime('%B de %Y')}</b>")
        stats = database.get_estatisticas_mensais(year, month)
        texto_dashboard = (f"Concluídos no Mês: <b style='color:green;'>{stats['concluidos']}</b>"
                           f"&nbsp;&nbsp;|&nbsp;&nbsp;"
                           f"Retificados no Mês: <b style='color:#17a2b8;'>{stats['retificados']}</b>")
        self.dashboard_label.setText(texto_dashboard)
        status_dias = database.get_status_dias_para_mes(year, month); month_calendar = calendar.monthcalendar(year, month)
        for week_num, week in enumerate(month_calendar):
            for day_num, day in enumerate(week):
                if day != 0:
                    date = QDate(year, month, day); info_do_dia = status_dias.get(day); cell = DayCellWidget(date, info_do_dia, day_num, self)
                    cell.clicked.connect(self.open_day_view); self.calendar_grid.addWidget(cell, week_num, day_num)
    def prev_month(self): self.current_date -= timedelta(days=self.current_date.day + 1); self.populate_calendar()
    def next_month(self): _, days_in_month = calendar.monthrange(self.current_date.year, self.current_date.month); self.current_date += timedelta(days=days_in_month - self.current_date.day + 1); self.populate_calendar()
    def open_day_view(self, date):
        if self.feriados.get(date) == "nacional":
            QMessageBox.warning(self, "Feriado Nacional", "Não é permitido agendar em um feriado nacional.")
            return
        dialog = DayViewDialog(date, self.usuario_atual, self)
        dialog.exec_()
    def gerenciar_clientes(self): dialog = JanelaClientes(self.usuario_atual, self); dialog.exec_()
    def manage_status(self): dialog = StatusDialog(self.usuario_atual, self); dialog.exec_()
    def abrir_configuracoes(self):
        dialog = ConfigDialog(self); dialog.exec_()
    def abrir_dialogo_relatorio(self):
        dialog = RelatorioDialog(self.usuario_atual, self)
        dialog.exec_()
    def gerenciar_usuarios(self):
        dialog = JanelaUsuarios(self.usuario_atual, self)
        dialog.exec_()

if __name__ == '__main__':
    QApplication.setOrganizationName("SuaOrganizacao"); QApplication.setApplicationName("Agendador"); database.iniciar_db()
    app = QApplication(sys.argv); app.setQuitOnLastWindowClosed(False)
    login = LoginDialog()
    if login.exec_() == QDialog.Accepted:
        usuario_logado = login.usuario_logado
        window = CalendarWindow(usuario_logado); window.show()
        sys.exit(app.exec_())
    else:
        sys.exit(0)