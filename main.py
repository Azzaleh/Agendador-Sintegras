import sys
import os
import collections
import subprocess
from pathlib import Path
from urllib.request import urlopen
from datetime import datetime, timedelta
import calendar
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QGridLayout, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QLabel, QDialog, QLineEdit, QComboBox, QMessageBox, 
                             QFileDialog, QFormLayout, QCheckBox, QTextEdit, QTableWidget, 
                             QTableWidgetItem, QHeaderView, QColorDialog, QSystemTrayIcon, QStyle,
                             QTimeEdit, QSpinBox, QRadioButton, QGroupBox, QDateEdit, QListWidget, 
                             QListWidgetItem, QMenu, QTreeWidget, QTreeWidgetItem, QToolTip,QCompleter,QTextBrowser)
from PyQt5.QtGui import QPainter, QColor, QBrush, QFont, QIcon
from PyQt5.QtCore import Qt, QDate, pyqtSignal, QTimer, QTime, QSettings,QThread,QStringListModel

import database 
import export
#from theme_manager import ThemeManager, load_stylesheet

VERSAO_ATUAL = "1.5"

class SuggestionLabel(QLabel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setToolTip("Clique com o botão direito para copiar as sugestões.")

    def contextMenuEvent(self, event):
        full_text = self.text()
        cleaned_text = full_text.replace('<br>', '\n').replace('<b>', '').replace('</b>', '')
        clipboard = QApplication.clipboard()
        clipboard.setText(cleaned_text)
        QToolTip.showText(event.globalPos(), "Copiado!", self)


class UpdateCheckerThread(QThread):
    update_found = pyqtSignal(str)
    check_finished = pyqtSignal(bool) 
    def run(self):
        url_versao = "https://raw.githubusercontent.com/Azzaleh/Agendador-Sintegras/main/version.txt"
        update_encontrado = False
        try:
            with urlopen(url_versao, timeout=10) as response:
                versao_github_str = response.read().decode('utf-8').strip()
            versao_atual_numerica = tuple(map(int, (VERSAO_ATUAL.split("."))))
            versao_github_numerica = tuple(map(int, (versao_github_str.split("."))))
            if versao_github_numerica > versao_atual_numerica:
                self.update_found.emit(versao_github_str)
                update_encontrado = True
        except Exception as e:
            print(f"Erro ao verificar atualização: {e}")
        finally:            
            self.check_finished.emit(update_encontrado)


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
        reset_btn = QPushButton("Limpar Configurações")
        login_btn = QPushButton("Entrar")
        sair_btn = QPushButton("Sair")
        
        login_btn.setDefault(True)
        
        botoes_layout.addStretch()
        botoes_layout.addWidget(login_btn)
        botoes_layout.addWidget(sair_btn)
        botoes_layout.addWidget(reset_btn)
        layout.addLayout(botoes_layout)
        
        login_btn.clicked.connect(self.tentar_login)
        sair_btn.clicked.connect(self.reject)
        reset_btn.clicked.connect(self.resetar_configuracoes)
        
        self.usuario_logado = None

    def tentar_login(self):
        username = self.usuario_edit.text()
        password = self.senha_edit.text()

        try:
            # Esta função agora deve funcionar, pois o DB e as tabelas já foram criados
            usuario = database.verificar_usuario(username, password)
            if usuario:
                self.usuario_logado = usuario
                self.accept()
            else:
                QMessageBox.warning(self, "Erro de Login", "Usuário ou senha inválidos.")
                self.senha_edit.clear()
        except Exception as e:
            # Este erro agora só deve aparecer se o serviço do Firebird parar DEPOIS do programa abrir
            QMessageBox.critical(self, "Erro de Conexão com o Banco de Dados",
                                 f"Não foi possível conectar ao banco de dados.\n"
                                 f"Verifique as suas configurações ou se o serviço do Firebird está a ser executado.\n\n"
                                 f"Erro técnico: {e}")
    def resetar_configuracoes(self):
        reply = QMessageBox.question(self, "Confirmar Limpeza de Configurações",
                                   "Isso irá apagar todas as configurações salvas (incluindo dados de conexão com o banco e horários).\n\n"
                                   "O aplicativo será fechado após a limpeza. Deseja continuar?",
                                   QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            try:
                settings = QSettings()
                settings.clear()
                QMessageBox.information(self, "Sucesso", 
                                        "As configurações foram resetadas com sucesso. O aplicativo agora será fechado.")
                QApplication.instance().quit()
            except Exception as e:
                QMessageBox.critical(self, "Erro", f"Ocorreu um erro ao tentar limpar as configurações:\n{e}")


class DayCellWidget(QWidget):
    clicked = pyqtSignal(QDate)
    def __init__(self, date, dia_info, day_of_week, parent=None):
        super().__init__(parent)
        self.date = date
        self.dia_info = dia_info if dia_info else {'COR': '#ffffff', 'CONTAGEM': 0}
        self.day_of_week = day_of_week
        self.calendar_window = parent
        self.setMinimumSize(100, 80)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        contagem = self.dia_info['CONTAGEM']
        is_weekend = self.day_of_week in [5, 6]
        feriados = self.calendar_window.feriados
        feriado_tipo = feriados.get(self.date)
        if feriado_tipo == "nacional":
            cor_fundo = QColor("#d8bfd8")
        elif feriado_tipo == "municipal":
            cor_fundo = QColor("#e6e6fa")
        elif contagem > 0:
            cor_fundo = QColor(self.dia_info['COR'])
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
        self.calendar_window.populate_calendar()



class DialogoCliente(QDialog):
    def __init__(self, usuario_logado, cliente_id=None, parent=None):
        super().__init__(parent)
        self.main_window = parent
        self.usuario_logado = usuario_logado
        self.cliente_id = cliente_id
        self.setWindowTitle("Adicionar Novo Cliente" if not cliente_id else "Editar Cliente")
        self.setMinimumWidth(400)
        layout = QVBoxLayout(self)
        form_layout = QFormLayout()
        
        # --- INÍCIO DA ALTERAÇÃO ---
        self.nome_edit = QLineEdit()
        self.tipo_envio_combo = QComboBox()
        self.tipo_envio_combo.addItems(["Nosso", "Deles"])
        self.contato_edit = QLineEdit() # Email/Local
        
        # Novos campos de telefone com máscara de entrada
        self.telefone1_edit = QLineEdit()
        self.telefone1_edit.setInputMask("(00) 00000-0000")
        self.telefone2_edit = QLineEdit()
        self.telefone2_edit.setInputMask("(00) 00000-0000")

        self.gera_recibo_check = QCheckBox("Gera Recibo?")
        self.conta_xmls_check = QCheckBox("Contar XMLs?")
        self.nivel_edit = QLineEdit()
        self.detalhes_edit = QTextEdit()
        self.detalhes_edit.setPlaceholderText("Detalhes adicionais sobre o cliente...")
        self.detalhes_edit.setFixedHeight(80) # Diminui a altura da caixa de observações

        self.num_computadores_spin = QSpinBox()
        self.num_computadores_spin.setRange(1, 9999)
        self.num_computadores_spin.setValue(1)
        
        form_layout.addRow("Nome*:", self.nome_edit)
        form_layout.addRow("Tipo de Envio*:", self.tipo_envio_combo)
        form_layout.addRow("Email/Local*:", self.contato_edit)
        form_layout.addRow("Telefone 1:", self.telefone1_edit)
        form_layout.addRow("Telefone 2:", self.telefone2_edit)
        form_layout.addRow("Nº de Computadores:", self.num_computadores_spin)
        form_layout.addRow("", self.gera_recibo_check)
        form_layout.addRow("", self.conta_xmls_check)
        form_layout.addRow("Nível:", self.nivel_edit)
        form_layout.addRow("Outros Detalhes:", self.detalhes_edit)
        self.recorrencia_group = QGroupBox("Agendamento Recorrente (Opcional)")
        self.recorrencia_group.setCheckable(True)
        self.recorrencia_group.setChecked(False)

        recorrencia_layout = QFormLayout(self.recorrencia_group)
        
        self.dia_recorrencia_spin = QSpinBox()
        self.dia_recorrencia_spin.setRange(1, 31)
        self.hora_recorrencia_edit = QTimeEdit()
        self.hora_recorrencia_edit.setDisplayFormat("HH:mm")

        periodo_layout = QHBoxLayout()
        self.radio_4_meses = QRadioButton("4 meses")
        self.radio_8_meses = QRadioButton("8 meses")
        self.radio_12_meses = QRadioButton("12 meses")
        self.radio_4_meses.setChecked(True)
        periodo_layout.addWidget(self.radio_4_meses)
        periodo_layout.addWidget(self.radio_8_meses)
        periodo_layout.addWidget(self.radio_12_meses)

        recorrencia_layout.addRow("Agendar todo dia:", self.dia_recorrencia_spin)
        recorrencia_layout.addRow("Às:", self.hora_recorrencia_edit)
        recorrencia_layout.addRow("Para os próximos:", periodo_layout)
        
        layout.addLayout(form_layout)
        layout.addWidget(self.recorrencia_group)
        # ============== FIM DA NOVA ALTERAÇÃO (RECORRÊNCIA) ==============

        botoes_layout = QHBoxLayout()

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
            self.carregar_dados() # carregar_dados será modificado em JanelaClientes

    def carregar_dados(self): pass # A lógica de carga fica em JanelaClientes.editar_cliente

    def salvar(self):
        nome = self.nome_edit.text().strip()
        tipo_envio = self.tipo_envio_combo.currentText()
        contato = self.contato_edit.text().strip()
        
        telefone1 = self.telefone1_edit.text() if self.telefone1_edit.hasAcceptableInput() else ""
        telefone2 = self.telefone2_edit.text() if self.telefone2_edit.hasAcceptableInput() else ""

        if not nome or not tipo_envio or not contato:
            QMessageBox.warning(self, "Campos Obrigatórios", "Por favor, preencha os campos Nome, Tipo de Envio e Email/Local.")
            return

        dados_cliente = {
            "nome": nome, "tipo_envio": tipo_envio, "contato": contato,
            "gera_recibo": self.gera_recibo_check.isChecked(),
            "conta_xmls": self.conta_xmls_check.isChecked(),
            "nivel": self.nivel_edit.text().strip(),
            "detalhes": self.detalhes_edit.toPlainText().strip(),
            "numero_computadores": self.num_computadores_spin.value(),
            "telefone1": telefone1,
            "telefone2": telefone2
        }
        usuario_nome = self.usuario_logado['USERNAME']
        
        try:
            if self.cliente_id:
                database.atualizar_cliente(self.cliente_id, **dados_cliente, usuario_logado=usuario_nome)
                cliente_id_salvo = self.cliente_id
            else:
                database.adicionar_cliente(**dados_cliente, usuario_logado=usuario_nome)
                conn = database.conectar()
                cur = conn.cursor()
                cur.execute("SELECT MAX(ID) FROM CLIENTES")
                cliente_id_salvo = cur.fetchone()[0]
                conn.close()

            # ============== INÍCIO DA LÓGICA CORRIGIDA E NO LUGAR CERTO ==============
            if self.recorrencia_group.isChecked():
                dia_desejado = self.dia_recorrencia_spin.value()
                hora_desejada = self.hora_recorrencia_edit.time().toString("HH:mm")
                meses = 4
                if self.radio_8_meses.isChecked(): meses = 8
                elif self.radio_12_meses.isChecked(): meses = 12

                reply = QMessageBox.question(self, "Confirmar Agendamento Recorrente",
                                        f"Serão criados {meses} agendamentos mensais.\n"
                                        "Se a data cair em um Sábado ou Domingo, o sistema buscará o próximo dia útil com horários livres.\n\n"
                                        "Deseja continuar?",
                                        QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

                if reply == QMessageBox.Yes:
                    lista_agendamentos_validados = []
                    hoje = QDate.currentDate()

                    for i in range(meses):
                        ano_alvo = hoje.year()
                        mes_alvo = hoje.month() + i + 1
                        
                        if mes_alvo > 12:
                            ano_alvo += (mes_alvo - 1) // 12
                            mes_alvo = (mes_alvo - 1) % 12 + 1
                        
                        dias_no_mes = QDate(ano_alvo, mes_alvo, 1).daysInMonth()
                        dia_real = min(dia_desejado, dias_no_mes)
                        
                        data_alvo = QDate(ano_alvo, mes_alvo, dia_real)
                        
                        data_final_agendamento = data_alvo
                        while True:
                            if data_final_agendamento <= hoje:
                                data_final_agendamento = hoje.addDays(1)

                            # 1. Verifica se é fim de semana (Sábado=6, Domingo=7)
                            if data_final_agendamento.dayOfWeek() >= 6:
                                data_final_agendamento = data_final_agendamento.addDays(1)
                                continue

                            # 2. Verifica se o dia está lotado
                            horarios_possiveis = self.main_window.gerar_horarios_dinamicos(data_final_agendamento)
                            agendamentos_no_dia = database.get_entregas_por_dia(data_final_agendamento.toString("yyyy-MM-dd"))
                            
                            if len(agendamentos_no_dia) < len(horarios_possiveis):
                                obs = "Agendamento criado automaticamente por recorrência."
                                if data_final_agendamento != data_alvo:
                                    obs += f" (Data original: {data_alvo.toString('dd/MM/yyyy')})"
                                
                                lista_agendamentos_validados.append({
                                    "cliente_id": cliente_id_salvo,
                                    "data": data_final_agendamento.toString("yyyy-MM-dd"),
                                    "hora": hora_desejada,
                                    "obs": obs
                                })
                                break 
                            else:
                                data_final_agendamento = data_final_agendamento.addDays(1)
                    
                    # Passa a lista já validada para a função do banco de dados
                    database.criar_agendamentos_recorrentes(lista_agendamentos_validados, usuario_nome)
                    QMessageBox.information(self, "Sucesso", f"{len(lista_agendamentos_validados)} agendamentos recorrentes foram criados com sucesso!")

            # ============== FIM DA LÓGICA CORRIGIDA ==============
        finally:
            self.accept()

class DialogoClientesPendentes(QDialog):
    def __init__(self, lista_clientes, mes, ano, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Relatório de Clientes Pendentes")
        self.setMinimumSize(400, 500)
        layout = QVBoxLayout(self)
        titulo_label = QLabel(f"<b>Clientes com agendamentos pendentes para {mes}/{ano}:</b>")
        titulo_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(titulo_label)

        total_pendentes = len(lista_clientes)
        total_label = QLabel(f"<b>Total de clientes pendentes: {total_pendentes}</b>")
        total_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(total_label)

        if not lista_clientes:
            aviso_label = QLabel("Nenhum cliente pendente encontrado.\nTodos foram agendados este mês!")
            aviso_label.setAlignment(Qt.AlignCenter)
            layout.addWidget(aviso_label)
        else:
            busca_layout = QHBoxLayout()
            busca_layout.addWidget(QLabel("Buscar:"))
            self.busca_edit = QLineEdit()
            self.busca_edit.setPlaceholderText("Digite para filtrar a lista...")
            busca_layout.addWidget(self.busca_edit)
            layout.addLayout(busca_layout)
            self.lista_widget = QListWidget()
            for nome_cliente in sorted(lista_clientes):
                self.lista_widget.addItem(QListWidgetItem(nome_cliente))
            layout.addWidget(self.lista_widget)
            self.busca_edit.textChanged.connect(self.filtrar_lista)
        fechar_btn = QPushButton("Fechar")
        fechar_btn.clicked.connect(self.accept)
        layout.addWidget(fechar_btn, alignment=Qt.AlignRight)

    def filtrar_lista(self):
        texto_busca = self.busca_edit.text().lower().strip()
        for i in range(self.lista_widget.count()):
            item = self.lista_widget.item(i)
            nome_cliente = item.text().lower()
            if texto_busca in nome_cliente:
                item.setHidden(False)
            else:
                item.setHidden(True)

class DialogoClientesInativos(QDialog):
    def __init__(self, lista_clientes, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Relatório de Atividade de Clientes")
        self.setMinimumSize(700, 500)

        layout = QVBoxLayout(self)
        titulo_label = QLabel("<b>Clientes com 3 meses ou mais sem agendamentos:</b>")
        titulo_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(titulo_label)

        # Filtro de busca
        busca_layout = QHBoxLayout()
        busca_layout.addWidget(QLabel("Buscar:"))
        self.busca_edit = QLineEdit()
        self.busca_edit.setPlaceholderText("Digite para filtrar a lista...")
        busca_layout.addWidget(self.busca_edit)
        layout.addLayout(busca_layout)

        # Tabela de resultados
        self.tabela_inativos = QTableWidget()
        self.tabela_inativos.setColumnCount(3)
        self.tabela_inativos.setHorizontalHeaderLabels(["Cliente", "Contato", "Status (Tempo Inativo)"])
        self.tabela_inativos.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tabela_inativos.setSelectionBehavior(QTableWidget.SelectRows)
        self.tabela_inativos.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tabela_inativos.setSortingEnabled(True)
        layout.addWidget(self.tabela_inativos)
        
        self.popular_tabela(lista_clientes)
        
        self.busca_edit.textChanged.connect(self.filtrar_tabela)

        fechar_btn = QPushButton("Fechar")
        fechar_btn.clicked.connect(self.accept)
        layout.addWidget(fechar_btn, alignment=Qt.AlignRight)

    def popular_tabela(self, lista_clientes):
        hoje = datetime.now()
        clientes_inativos = []

        # Filtra a lista para incluir apenas clientes inativos há 3 meses ou mais
        for cliente in lista_clientes:
            ultimo_agendamento = cliente.get('ULTIMO_AGENDAMENTO')
            status = ""
            meses_inativo = 999 # Um número alto para quem nunca agendou

            if ultimo_agendamento is None:
                status = "Nunca agendou"
            else:
                # Converte a data do banco para objeto datetime
                data_ultimo = datetime.strptime(str(ultimo_agendamento), '%Y-%m-%d')
                diferenca = hoje - data_ultimo
                meses_inativo = diferenca.days // 30
                
                if meses_inativo == 0:
                    status = f"Agendou há {diferenca.days} dias"
                elif meses_inativo == 1:
                    status = "Agendou há 1 mês"
                else:
                    status = f"Não agenda há {meses_inativo} meses"
            
            # Adiciona à lista apenas se for inativo há 3 meses ou mais (ou nunca agendou)
            if meses_inativo >= 3:
                cliente['status_formatado'] = status
                clientes_inativos.append(cliente)

        self.tabela_inativos.setRowCount(len(clientes_inativos))
        for i, cliente in enumerate(clientes_inativos):
            self.tabela_inativos.setItem(i, 0, QTableWidgetItem(cliente['NOME']))
            self.tabela_inativos.setItem(i, 1, QTableWidgetItem(cliente.get('CONTATO', '')))
            self.tabela_inativos.setItem(i, 2, QTableWidgetItem(cliente['status_formatado']))
    
    def filtrar_tabela(self):
        texto_busca = self.busca_edit.text().lower().strip()
        for i in range(self.tabela_inativos.rowCount()):
            item_nome = self.tabela_inativos.item(i, 0)
            if item_nome and texto_busca in item_nome.text().lower():
                self.tabela_inativos.setRowHidden(i, False)
            else:
                self.tabela_inativos.setRowHidden(i, True)

class DialogoEstatisticasCliente(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Análise de Performance de Clientes")
        self.setMinimumSize(750, 600)

        # Layout Principal
        main_layout = QVBoxLayout(self)

        # --- Filtros ---
        filtros_group = QGroupBox("Filtros")
        filtros_layout = QFormLayout(filtros_group)
        self.cliente_combo = QComboBox()
        self.cliente_combo.setEditable(True)
        self.cliente_combo.setInsertPolicy(QComboBox.NoInsert)
        self.data_inicio_edit = QDateEdit(QDate.currentDate().addMonths(-3))
        self.data_inicio_edit.setCalendarPopup(True)
        self.data_fim_edit = QDateEdit(QDate.currentDate())
        self.data_fim_edit.setCalendarPopup(True)
        self.analisar_btn = QPushButton("Analisar")

        filtros_layout.addRow("Cliente:", self.cliente_combo)
        filtros_layout.addRow("Período de:", self.data_inicio_edit)
        filtros_layout.addRow("Até:", self.data_fim_edit)
        filtros_layout.addRow(self.analisar_btn)
        main_layout.addWidget(filtros_group)

        # --- Resultados ---
        resultados_layout = QHBoxLayout()
        # Resultado Individual
        individual_group = QGroupBox("Análise do Cliente Selecionado")
        individual_layout = QVBoxLayout(individual_group)
        self.resultado_individual_label = QLabel("<i>Selecione um cliente e um período e clique em 'Analisar'.</i>")
        self.resultado_individual_label.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        individual_layout.addWidget(self.resultado_individual_label)
        # Rankings Gerais
        ranking_group = QGroupBox("Destaques do Período (Todos os Clientes)")
        ranking_layout = QVBoxLayout(ranking_group)
        self.resultado_ranking_label = QLabel("<i>...</i>")
        self.resultado_ranking_label.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        ranking_layout.addWidget(self.resultado_ranking_label)

        resultados_layout.addWidget(individual_group)
        resultados_layout.addWidget(ranking_group)
        main_layout.addLayout(resultados_layout)
        
        self.carregar_clientes_combo()
        self.analisar_btn.clicked.connect(self.realizar_analise)

    def carregar_clientes_combo(self):
        clientes = database.listar_clientes()
        for cliente in sorted(clientes, key=lambda x: x['NOME']):
            self.cliente_combo.addItem(cliente['NOME'], cliente['ID'])
        
        completer = QCompleter([self.cliente_combo.itemText(i) for i in range(self.cliente_combo.count())], self)
        completer.setCaseSensitivity(Qt.CaseInsensitive)
        completer.setFilterMode(Qt.MatchContains)
        self.cliente_combo.setCompleter(completer)

    def realizar_analise(self):
        cliente_id = self.cliente_combo.currentData()
        cliente_nome = self.cliente_combo.currentText()
        data_inicio = self.data_inicio_edit.date().toString("yyyy-MM-dd")
        data_fim = self.data_fim_edit.date().toString("yyyy-MM-dd")

        if not cliente_id:
            QMessageBox.warning(self, "Aviso", "Por favor, selecione um cliente válido.")
            return

        # 1. Análise Individual (sem alteração)
        stats_cliente = database.get_estatisticas_cliente_periodo(cliente_id, data_inicio, data_fim)
        texto_individual = f"<b>{cliente_nome}</b><br><hr>"
        if not stats_cliente:
            texto_individual += "Nenhum agendamento encontrado no período."
        else:
            for status, contagem in stats_cliente.items():
                texto_individual += f"<b>• {status}:</b> {contagem}<br>"
        self.resultado_individual_label.setText(texto_individual)

        # --- INÍCIO DA CORREÇÃO ---
        # 2. Análise de Ranking (lógica atualizada)
        dados_ranking = database.get_dados_ranking_clientes_periodo(data_inicio, data_fim)
        texto_ranking = ""
        if not dados_ranking:
            texto_ranking = "Nenhum dado para gerar destaques."
        else:
            # --- Encontrar os "MAIS" ---
            mais_concluidos = max(dados_ranking, key=lambda x: x['CONCLUIDOS'])
            mais_retificados = max(dados_ranking, key=lambda x: x['RETIFICADOS'])
            mais_remarcados = max(dados_ranking, key=lambda x: x['REMARCADOS'])
            mais_erros = max(dados_ranking, key=lambda x: x['ERROS'])

            # --- Encontrar os "MENOS" (ignorando os que têm 0) ---
            clientes_com_retificados = [c for c in dados_ranking if c['RETIFICADOS'] > 0]
            menos_retificados = min(clientes_com_retificados, key=lambda x: x['RETIFICADOS']) if clientes_com_retificados else None
            
            clientes_com_remarcados = [c for c in dados_ranking if c['REMARCADOS'] > 0]
            menos_remarcados = min(clientes_com_remarcados, key=lambda x: x['REMARCADOS']) if clientes_com_remarcados else None

            clientes_com_erros = [c for c in dados_ranking if c['ERROS'] > 0]
            menos_erros = min(clientes_com_erros, key=lambda x: x['ERROS']) if clientes_com_erros else None

            # --- Monta o texto de exibição ---
            texto_ranking += "<b><u>Mais Produtivo</u> (Concluídos):</b><br>"
            texto_ranking += f"{mais_concluidos['NOME']} ({mais_concluidos['CONCLUIDOS']})<br><br>"
            
            texto_ranking += "<b><u>Mais Retificações:</u></b><br>"
            texto_ranking += f"{mais_retificados['NOME']} ({mais_retificados['RETIFICADOS']})<br><br>"
            
            texto_ranking += "<b><u>Mais Remarcações:</u></b><br>"
            texto_ranking += f"{mais_remarcados['NOME']} ({mais_remarcados['REMARCADOS']})<br><br>"

            texto_ranking += "<b><u>Mais Erros:</u></b><br>"
            texto_ranking += f"{mais_erros['NOME']} ({mais_erros['ERROS']})<br><hr>"

            texto_ranking += "<b><u>Menos Retificações</u> (acima de 0):</b><br>"
            texto_ranking += f"{menos_retificados['NOME']} ({menos_retificados['RETIFICADOS']})" if menos_retificados else "N/A"
            texto_ranking += "<br><br>"
            
            texto_ranking += "<b><u>Menos Remarcações</u> (acima de 0):</b><br>"
            texto_ranking += f"{menos_remarcados['NOME']} ({menos_remarcados['REMARCADOS']})" if menos_remarcados else "N/A"
            texto_ranking += "<br><br>"

            texto_ranking += "<b><u>Menos Erros</u> (acima de 0):</b><br>"
            texto_ranking += f"{menos_erros['NOME']} ({menos_erros['ERROS']})" if menos_erros else "N/A"
            texto_ranking += "<br>"
            
        self.resultado_ranking_label.setText(texto_ranking)

class JanelaClientes(QDialog):
    def __init__(self, usuario_logado, parent=None):
        super().__init__(parent)
        self.usuario_logado = usuario_logado
        self.setWindowTitle("Gerenciamento de Clientes")
        self.setMinimumSize(800, 600)
        layout = QVBoxLayout(self)
        
        # --- Layout de Busca (sem alterações) ---
        busca_layout = QHBoxLayout()
        busca_label = QLabel("Buscar Cliente:")
        self.busca_edit = QLineEdit()
        self.busca_edit.setPlaceholderText("Digite o nome do cliente para filtrar...")
        busca_layout.addWidget(busca_label)
        busca_layout.addWidget(self.busca_edit)
        layout.addLayout(busca_layout)
        
        # --- Tabela de Clientes (sem alterações) ---
        self.tabela_clientes = QTableWidget()
        self.tabela_clientes.setColumnCount(4)
        self.tabela_clientes.setHorizontalHeaderLabels(["Nome", "Tipo de Envio", "Email/Local", "Nível"])
        self.tabela_clientes.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tabela_clientes.setSelectionBehavior(QTableWidget.SelectRows)
        self.tabela_clientes.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.tabela_clientes)
        
        # --- INÍCIO DA CORREÇÃO ---
        # Bloco único e corrigido para os botões
        botoes_layout = QHBoxLayout()
        import_btn = QPushButton("Importar de XLSX")
        pendentes_btn = QPushButton("Verificar Pendentes no Mês")
        inativos_btn = QPushButton("Verificar Inativos") # Novo botão adicionado aqui
        add_btn = QPushButton("Adicionar Novo")
        edit_btn = QPushButton("Editar Selecionado")
        del_btn = QPushButton("Excluir Selecionado")
        
        botoes_layout.addWidget(import_btn)
        botoes_layout.addWidget(pendentes_btn)
        botoes_layout.addWidget(inativos_btn) # Adicionado ao layout
        botoes_layout.addStretch()
        botoes_layout.addWidget(add_btn)
        botoes_layout.addWidget(edit_btn)
        botoes_layout.addWidget(del_btn)
        
        layout.addLayout(botoes_layout)
        
        # Conexões dos sinais
        import_btn.clicked.connect(self.importar_clientes)
        add_btn.clicked.connect(self.adicionar_cliente)
        edit_btn.clicked.connect(self.editar_cliente)
        del_btn.clicked.connect(self.excluir_cliente)
        self.busca_edit.textChanged.connect(self.filtrar_tabela)
        pendentes_btn.clicked.connect(self.verificar_pendentes)
        inativos_btn.clicked.connect(self.verificar_inativos) # Conexão para o novo botão
        # --- FIM DA CORREÇÃO ---
        
        self.carregar_clientes()

    def verificar_pendentes(self):
        hoje = QDate.currentDate()
        ano_atual = hoje.year()
        mes_atual = hoje.month()
        ids_clientes_agendados = database.get_clientes_com_agendamento_no_mes(ano_atual, mes_atual)
        todos_clientes = database.listar_clientes()
        clientes_pendentes = []
        for cliente in todos_clientes:
            if cliente['ID'] not in ids_clientes_agendados:
                clientes_pendentes.append(cliente['NOME'])
        dialog = DialogoClientesPendentes(clientes_pendentes, mes_atual, ano_atual, self)
        dialog.exec_()

    def verificar_inativos(self):
        # Chama a nova função do banco de dados
        lista_completa_clientes = database.get_status_de_atividade_clientes()

        if not lista_completa_clientes:
            QMessageBox.information(self, "Aviso", "Nenhum cliente encontrado no sistema.")
            return
            
        # Cria e exibe o diálogo com a lista de clientes
        dialog = DialogoClientesInativos(lista_completa_clientes, self)
        dialog.exec_()

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
            
            mapa_colunas = { 
                'Clientes': 'nome', 'Envia do nosso ou deles?': 'tipo_envio', 
                'Email da contabilidade (Ou local a ser deixado)': 'contato', 
                'Gera Recibo?': 'gera_recibo', 'Contar XMLs?': 'conta_xmls', 
                'Nível': 'nivel', 'Outros detalhes': 'detalhes',
                'Nº de Computadores': 'numero_computadores',
                'Telefone 1': 'telefone1',
                'Telefone 2': 'telefone2'
            }

            colunas_obrigatorias = ['Clientes', 'Envia do nosso ou deles?', 'Email da contabilidade (Ou local a ser deixado)']
            if not all(col in df.columns for col in colunas_obrigatorias):
                QMessageBox.critical(self, "Erro de Importação", f"O arquivo deve conter as colunas obrigatórias: {', '.join(colunas_obrigatorias)}"); return
            
            clientes_importados = 0
            usuario_nome = self.usuario_logado['USERNAME']
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
                    # --- INÍCIO DA CORREÇÃO ---
                    # Pega o valor lido da planilha
                    num_pc = dados_cliente.get('numero_computadores')
                    
                    # Se o valor for None (célula vazia), define um padrão.
                    if num_pc is None:
                        num_pc = 1
                    
                    database.adicionar_cliente(
                        nome=dados_cliente['nome'], tipo_envio=dados_cliente['tipo_envio'], 
                        contato=dados_cliente['contato'], gera_recibo=dados_cliente.get('gera_recibo', False), 
                        conta_xmls=dados_cliente.get('conta_xmls', False), nivel=dados_cliente.get('nivel'), 
                        detalhes=dados_cliente.get('detalhes'),
                        numero_computadores=int(float(num_pc)), # Agora é seguro converter para int
                        telefone1=dados_cliente.get('telefone1'),
                        telefone2=dados_cliente.get('telefone2'),
                        usuario_logado=usuario_nome
                    )
                    # --- FIM DA CORREÇÃO ---
                    clientes_importados += 1
            QMessageBox.information(self, "Sucesso", f"{clientes_importados} clientes importados com sucesso!"); self.carregar_clientes()
        except Exception as e: 
            QMessageBox.critical(self, "Erro de Importação", f"Ocorreu um erro ao ler o arquivo:\n{e}")
    def carregar_clientes(self):
        self.tabela_clientes.setRowCount(0)
        for cliente in database.listar_clientes():
            row = self.tabela_clientes.rowCount()
            self.tabela_clientes.insertRow(row)
            self.tabela_clientes.setItem(row, 0, QTableWidgetItem(cliente['NOME']))
            self.tabela_clientes.setItem(row, 1, QTableWidgetItem(cliente['TIPO_ENVIO']))
            self.tabela_clientes.setItem(row, 2, QTableWidgetItem(cliente['CONTATO']))
            self.tabela_clientes.setItem(row, 3, QTableWidgetItem(str(cliente.get('NIVEL', ''))))
            self.tabela_clientes.item(row, 0).setData(Qt.UserRole, dict(cliente))
        self.filtrar_tabela()

    def adicionar_cliente(self):
        dialog = DialogoCliente(self.usuario_logado, parent=self.parent())
        if dialog.exec_() == QDialog.Accepted: 
            self.carregar_clientes()

    def editar_cliente(self):
        itens_selecionados = self.tabela_clientes.selectedItems()
        if not itens_selecionados: 
            QMessageBox.information(self, "Aviso", "Por favor, selecione um cliente na tabela para editar.")
            return
        linha_selecionada = itens_selecionados[0].row()
        cliente_data = self.tabela_clientes.item(linha_selecionada, 0).data(Qt.UserRole)
        
        # Apenas uma linha, usando self.parent() para passar a janela principal
        dialog = DialogoCliente(self.usuario_logado, cliente_id=cliente_data['ID'], parent=self.parent()) 
        
        dialog.nome_edit.setText(cliente_data['NOME'])
        dialog.tipo_envio_combo.setCurrentText(cliente_data['TIPO_ENVIO'])
        dialog.contato_edit.setText(cliente_data['CONTATO'])
        dialog.telefone1_edit.setText(cliente_data.get('TELEFONE1', ''))
        dialog.telefone2_edit.setText(cliente_data.get('TELEFONE2', ''))
        dialog.gera_recibo_check.setChecked(bool(cliente_data['GERA_RECIBO']))
        dialog.conta_xmls_check.setChecked(bool(cliente_data['CONTA_XMLS']))
        dialog.nivel_edit.setText(str(cliente_data.get('NIVEL', '')))
        dialog.detalhes_edit.setText(cliente_data.get('OUTROS_DETALHES', ''))
        dialog.num_computadores_spin.setValue(cliente_data.get('NUMERO_COMPUTADORES', 0))
            
        if dialog.exec_() == QDialog.Accepted: 
            self.carregar_clientes()

    def excluir_cliente(self):
        itens_selecionados = self.tabela_clientes.selectedItems()
        if not itens_selecionados: 
            QMessageBox.information(self, "Aviso", "Por favor, selecione um cliente para excluir.")
            return
        linha_selecionada = itens_selecionados[0].row()
        cliente_data = self.tabela_clientes.item(linha_selecionada, 0).data(Qt.UserRole)
        reply = QMessageBox.question(self, "Confirmar Exclusão", f"Tem certeza que deseja excluir o cliente '{cliente_data['NOME']}'?\nTODAS as entregas associadas a ele também serão excluídas.", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes: 
            database.deletar_cliente(cliente_data['ID'], self.usuario_logado['USERNAME'])
            self.carregar_clientes()

    def keyPressEvent(self, event):
        # Verifica se a tecla pressionada foi F3
        if event.key() == Qt.Key_F3:
            print("Tecla F3 pressionada, abrindo estatísticas...")
            dialog = DialogoEstatisticasCliente(self)
            dialog.exec_()
        else:
            # Passa o evento para o tratamento padrão para outras teclas
            super().keyPressEvent(event)



class FormularioStatusDialog(QDialog):
    def __init__(self, status=None, parent=None):
        super().__init__(parent)
        self.status = status
        self.setWindowTitle("Novo Status" if not status else "Editar Status")
        layout = QFormLayout(self)
        self.nome_edit = QLineEdit()
        self.cor_btn = QPushButton("Escolher Cor")
        self.cor_label = QLabel()
        self.cor_label.setFixedSize(20, 20)
        cor_layout = QHBoxLayout()
        cor_layout.addWidget(self.cor_btn)
        cor_layout.addWidget(self.cor_label)
        layout.addRow("Nome:", self.nome_edit)
        layout.addRow("Cor:", cor_layout)
        if status:
            self.nome_edit.setText(status['NOME'])
            self.cor_hex = status['COR_HEX']
        else:
            self.cor_hex = "#ffffff"
        self.atualizar_label_cor()
        self.cor_btn.clicked.connect(self.escolher_cor)
        botoes = QHBoxLayout()
        salvar_btn = QPushButton("Salvar")
        cancelar_btn = QPushButton("Cancelar")
        botoes.addStretch()
        botoes.addWidget(salvar_btn)
        botoes.addWidget(cancelar_btn)
        layout.addRow(botoes)
        salvar_btn.clicked.connect(self.accept)
        cancelar_btn.clicked.connect(self.reject)

    def escolher_cor(self):
        cor = QColorDialog.getColor(QColor(self.cor_hex), self)
        if cor.isValid():
            self.cor_hex = cor.name()
            self.atualizar_label_cor()

    def atualizar_label_cor(self):
        self.cor_label.setStyleSheet(f"background-color: {self.cor_hex}; border: 1px solid black;")

    def get_data(self):
        return self.nome_edit.text().strip(), self.cor_hex

class StatusDialog(QDialog):
    def __init__(self, usuario_logado, parent=None):
        super().__init__(parent)
        self.usuario_logado = usuario_logado
        self.setWindowTitle("Gerenciar Status")
        self.setMinimumSize(500, 400)
        layout = QVBoxLayout(self)
        self.tabela_status = QTableWidget()
        self.tabela_status.setColumnCount(2)
        self.tabela_status.setHorizontalHeaderLabels(["Nome do Status", "Cor"])
        self.tabela_status.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tabela_status.setSelectionBehavior(QTableWidget.SelectRows)
        self.tabela_status.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.tabela_status)
        botoes = QHBoxLayout()
        add_btn = QPushButton("Adicionar")
        edit_btn = QPushButton("Editar")
        del_btn = QPushButton("Excluir")
        botoes.addStretch()
        botoes.addWidget(add_btn)
        botoes.addWidget(edit_btn)
        botoes.addWidget(del_btn)
        layout.addLayout(botoes)
        add_btn.clicked.connect(self.adicionar)
        edit_btn.clicked.connect(self.editar)
        del_btn.clicked.connect(self.excluir)
        self.carregar_status()

    def carregar_status(self):
        self.tabela_status.setRowCount(0)
        for status in database.listar_status():
            row = self.tabela_status.rowCount()
            self.tabela_status.insertRow(row)
            item_nome = QTableWidgetItem(status['NOME'])
            item_nome.setData(Qt.UserRole, dict(status))
            item_cor = QTableWidgetItem()
            item_cor.setBackground(QColor(status['COR_HEX']))
            self.tabela_status.setItem(row, 0, item_nome)
            self.tabela_status.setItem(row, 1, item_cor)

    def adicionar(self):
        dialog = FormularioStatusDialog(parent=self)
        if dialog.exec_() == QDialog.Accepted:
            nome, cor_hex = dialog.get_data()
            if nome:
                database.adicionar_status(nome, cor_hex, self.usuario_logado['USERNAME'])
                self.carregar_status()
                self.parent().populate_calendar()

    def editar(self):
        linha = self.tabela_status.currentRow()
        if linha < 0:
            return
        status_data = self.tabela_status.item(linha, 0).data(Qt.UserRole)
        dialog = FormularioStatusDialog(status=status_data, parent=self)
        
        if dialog.exec_() == QDialog.Accepted:
            nome, cor_hex = dialog.get_data()
            if nome:
                database.atualizar_status(status_data['ID'], nome, cor_hex, self.usuario_logado['USERNAME'])
                self.carregar_status()
                self.parent().populate_calendar()

    def excluir(self):
        linha = self.tabela_status.currentRow()
        if linha < 0:
            return
        status_data = self.tabela_status.item(linha, 0).data(Qt.UserRole)
        reply = QMessageBox.question(self, "Confirmar Exclusão", f"Tem certeza que deseja excluir o status '{status_data['NOME']}'?\nOs agendamentos que usam este status ficarão sem categoria.", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            database.deletar_status(status_data['ID'], self.usuario_logado['USERNAME'])
            self.carregar_status()
            self.parent().populate_calendar()


class EntregaDialog(QDialog):
    def __init__(self, usuario_logado, date, entrega_data=None, parent=None):
        super().__init__(parent)
        self.entrega_data = entrega_data
        self.usuario_logado = usuario_logado
        self.date = date
        self.cliente_selecionado_data = None # Armazena dados do cliente atual

        self.setWindowTitle("Agendar Novo Sintegra" if not entrega_data else "Editar Agendamento")
        layout = QFormLayout(self)

        self.cliente_combo = QComboBox()
        self.cliente_combo.setEditable(True)
        self.cliente_combo.setInsertPolicy(QComboBox.NoInsert)
        self.cliente_combo.completer().setCompletionMode(QCompleter.PopupCompletion)

        self.status_combo = QComboBox()
        self.responsavel_edit = QLineEdit()
        self.responsavel_edit.setReadOnly(True) # Tarefa 1 da sua lista (bloquear campo)
        self.observacoes_edit = QTextEdit()
        
        self.retificacao_check = QCheckBox("É Retificação?")
        
        self.rascunho_edit = QLineEdit()
        self.rascunho_edit.setReadOnly(True)
        self.rascunho_edit.setPlaceholderText("Selecione um cliente...")

        # --- INÍCIO DA ALTERAÇÃO ---
        botoes_copia_layout = QHBoxLayout()
        copiar_email_btn = QPushButton("Copiar Email/Local")
        copiar_rascunho_btn = QPushButton("Copiar Rascunho")
        botoes_copia_layout.addWidget(copiar_email_btn)
        botoes_copia_layout.addWidget(copiar_rascunho_btn)
        # --- FIM DA ALTERAÇÃO ---

        self.salvar_btn = QPushButton("Salvar")
        self.cancelar_btn = QPushButton("Cancelar")

        self.responsavel_edit.setText(self.usuario_logado['USERNAME'])
        self.carregar_combos()

        layout.addRow("Cliente:", self.cliente_combo)
        layout.addRow("Status:", self.status_combo)
        layout.addRow("Responsável:", self.responsavel_edit)
        layout.addRow("", self.retificacao_check)
        layout.addRow("Rascunho:", self.rascunho_edit)
        layout.addRow("Ações:", botoes_copia_layout) # Adiciona a linha de botões
        layout.addRow("Observações:", self.observacoes_edit)

        if entrega_data:
            index_cliente = self.cliente_combo.findData(entrega_data['CLIENTE_ID'])
            if index_cliente > -1: self.cliente_combo.setCurrentIndex(index_cliente)
            index_status = self.status_combo.findData(entrega_data['STATUS_ID'])
            if index_status > -1: self.status_combo.setCurrentIndex(index_status)
            if entrega_data.get('RESPONSAVEL'):
                 self.responsavel_edit.setText(entrega_data['RESPONSAVEL'])
            self.observacoes_edit.setText(entrega_data.get('OBSERVACOES', ''))
            self.retificacao_check.setChecked(bool(entrega_data.get('IS_RETIFICACAO', 0)))
        else:
            self.cliente_combo.setCurrentIndex(-1)
            index_pendente = self.status_combo.findText("Pendente") 
            if index_pendente > -1: self.status_combo.setCurrentIndex(index_pendente)
            
        botoes_layout = QHBoxLayout()
        botoes_layout.addStretch()
        botoes_layout.addWidget(self.salvar_btn)
        botoes_layout.addWidget(self.cancelar_btn)
        layout.addRow(botoes_layout)
        
        self.salvar_btn.clicked.connect(self.accept)
        self.cancelar_btn.clicked.connect(self.reject)
        copiar_rascunho_btn.clicked.connect(self.copiar_rascunho)
        copiar_email_btn.clicked.connect(self.copiar_contato_cliente) # Conecta novo botão
        self.cliente_combo.currentIndexChanged.connect(self.atualizar_dados_cliente)
        self.status_combo.currentIndexChanged.connect(self.atualizar_rascunho)
        self.atualizar_dados_cliente() # Chamada inicial

    def carregar_combos(self):
        #... (sem alterações nesta função)
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
                self.cliente_combo.addItem(cliente['NOME'], cliente['ID'])
        completer = QCompleter([self.cliente_combo.itemText(i) for i in range(self.cliente_combo.count())], self)
        completer.setCaseSensitivity(Qt.CaseInsensitive)
        completer.setFilterMode(Qt.MatchContains)
        self.cliente_combo.setCompleter(completer)
        for status in status_list:
            self.status_combo.addItem(status['NOME'], status['ID'])

    def atualizar_dados_cliente(self):
        cliente_id = self.cliente_combo.currentData()
        if cliente_id:
            self.cliente_selecionado_data = database.get_cliente_por_id(cliente_id)
        else:
            self.cliente_selecionado_data = None
        self.atualizar_rascunho()

    def atualizar_rascunho(self):
        #... (sem alterações nesta função)
        if not self.cliente_combo.isEnabled(): return
        nome_cliente = self.cliente_combo.currentText()
        nome_status = self.status_combo.currentText()
        data_str = self.date.toString("MM/yyyy")
        if nome_status and nome_status.lower() == 'retificado':
            texto_rascunho = f"Sintegra Retificado-{data_str}-{nome_cliente}"
        else:
            texto_rascunho = f"Sintegra -{data_str}-{nome_cliente}"
        self.rascunho_edit.setText(texto_rascunho)
    
    def copiar_contato_cliente(self):
        if self.cliente_selecionado_data and self.cliente_selecionado_data.get('CONTATO'):
            texto_para_copiar = self.cliente_selecionado_data['CONTATO']
            clipboard = QApplication.clipboard()
            clipboard.setText(texto_para_copiar)
            botao = self.sender()
            botao.setEnabled(False) 
            botao.setText("Copiado!")
            QTimer.singleShot(1500, lambda: (botao.setText("Copiar Email/Local"), botao.setEnabled(True)))
        else:
            QMessageBox.information(self, "Aviso", "Nenhum contato encontrado para este cliente.")

    def copiar_rascunho(self):
        texto_para_copiar = self.rascunho_edit.text()
        if texto_para_copiar:
            clipboard = QApplication.clipboard()
            clipboard.setText(texto_para_copiar)
            botao = self.sender()
            botao.setEnabled(False) 
            botao.setText("Copiado!")
            def restaurar_botao():
                try:
                    botao.setText("Copiar Rascunho")
                    botao.setEnabled(True)
                except RuntimeError: pass
            QTimer.singleShot(1500, restaurar_botao)

    def get_data(self):
        #... (sem alterações nesta função)
        return {
            "cliente_id": self.cliente_combo.currentData(),
            "status_id": self.status_combo.currentData(),
            "responsavel": self.responsavel_edit.text().strip(),
            "observacoes": self.observacoes_edit.toPlainText().strip(),
            "is_retificacao": self.retificacao_check.isChecked()
        }

class DayViewDialog(QDialog):
    def __init__(self, date, usuario_logado, parent=None):
        super().__init__(parent)
        self.date = date
        self.usuario_logado = usuario_logado
        self.setWindowTitle(f"Agenda para {date.toString('dd/MM/yyyy')}")
        self.setMinimumSize(800, 600)
        self.horarios = self.parent().gerar_horarios_dinamicos(self.date)
        layout = QVBoxLayout(self)
        self.tabela_agenda = QTableWidget()
        self.tabela_agenda.setColumnCount(5)
        self.tabela_agenda.setHorizontalHeaderLabels(["Horário", "Cliente", "Contato", "Tipo Envio", "Status"])
        self.tabela_agenda.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tabela_agenda.setSelectionBehavior(QTableWidget.SelectRows)
        self.tabela_agenda.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tabela_agenda.setColumnWidth(0, 80)
        layout.addWidget(self.tabela_agenda)
        botoes_layout = QHBoxLayout()
        edit_btn = QPushButton("Editar Agendamento")
        del_btn = QPushButton("Excluir Agendamento")
        botoes_layout.addStretch()
        botoes_layout.addWidget(edit_btn)
        botoes_layout.addWidget(del_btn)
        layout.addLayout(botoes_layout)
        edit_btn.clicked.connect(self.editar_agendamento)
        del_btn.clicked.connect(self.excluir_agendamento)
        self.tabela_agenda.cellDoubleClicked.connect(self.gerenciar_agendamento_duplo_clique)
        self.carregar_agenda_dia()

    def carregar_agenda_dia(self):
        self.tabela_agenda.setRowCount(0)
        self.tabela_agenda.setRowCount(len(self.horarios))
        agendamentos_dia = database.get_entregas_por_dia(self.date.toString("yyyy-MM-dd"))
        
        for i, horario in enumerate(self.horarios):
            item_horario = QTableWidgetItem(horario)
            item_horario.setTextAlignment(Qt.AlignCenter)
            self.tabela_agenda.setItem(i, 0, item_horario)
            
            if horario in agendamentos_dia:
                agendamento = agendamentos_dia[horario]
                
                is_retificacao = agendamento.get('IS_RETIFICACAO', 0)
                cor_fundo_linha = QColor("#17a2b8") if is_retificacao else None

                def create_item_with_bg(text, background_color):
                    item = QTableWidgetItem(str(text))
                    if background_color:
                        item.setBackground(background_color)
                    return item

                self.tabela_agenda.setItem(i, 1, create_item_with_bg(agendamento['NOME_CLIENTE'], cor_fundo_linha))
                self.tabela_agenda.setItem(i, 2, create_item_with_bg(agendamento['CONTATO'], cor_fundo_linha))
                self.tabela_agenda.setItem(i, 3, create_item_with_bg(agendamento['TIPO_ENVIO'], cor_fundo_linha))
                
                item_status = create_item_with_bg(agendamento['NOME_STATUS'], cor_fundo_linha)
                if not is_retificacao:
                     item_status.setBackground(QColor(agendamento['COR_HEX']))
                self.tabela_agenda.setItem(i, 4, item_status)
                
                if cor_fundo_linha:
                    item_horario.setBackground(cor_fundo_linha)

                item_horario.setData(Qt.UserRole, agendamento)
                observacoes = agendamento.get('OBSERVACOES', 'Nenhuma observação.')
                num_computadores = agendamento.get('NUMERO_COMPUTADORES', 0)
                
                texto_tooltip = (f"<b>Nº de Computadores:</b> {num_computadores}<br><hr><b>Observações:</b><br>{observacoes}")
                for col in range(self.tabela_agenda.columnCount()):
                    item = self.tabela_agenda.item(i, col)
                    if item: 
                        item.setToolTip(texto_tooltip)
            else:
                for j in range(1, 5):
                    self.tabela_agenda.setItem(i, j, QTableWidgetItem(""))

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
                cliente_id_selecionado = data['cliente_id']

                # ============== INÍCIO DA ALTERAÇÃO ==============
                pendente_existente = database.verificar_agendamento_pendente_existente(cliente_id_selecionado)

                if pendente_existente:
                    data_obj = pendente_existente['DATA_VENCIMENTO']
                    data_pendente = data_obj.strftime('%d/%m/%Y')
                    hora_pendente = pendente_existente['HORARIO']
                    
                    reply = QMessageBox.question(self, "Cliente já possui agendamento pendente",
                                               f"Este cliente já tem um agendamento pendente para {data_pendente} às {hora_pendente}.\n\n"
                                               "Deseja desmarcar o agendamento antigo e criar este novo?",
                                               QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

                    if reply == QMessageBox.No:
                        QMessageBox.information(self, "Operação Cancelada", "O novo agendamento não foi criado.")
                        return 
                    
                    else:
                        database.deletar_entrega(pendente_existente['ID'], self.usuario_logado['USERNAME'])
                # ============== FIM DA ALTERAÇÃO ==============

                database.adicionar_entrega(
                    self.date.toString("yyyy-MM-dd"), 
                    horario, 
                    data['status_id'], 
                    data['cliente_id'], 
                    data['responsavel'], 
                    data['observacoes'], 
                    data['is_retificacao'], 
                    self.usuario_logado['USERNAME']
                )

        self.carregar_agenda_dia()
        self.parent().populate_calendar()

    def editar_agendamento(self):
        itens_selecionados = self.tabela_agenda.selectedItems()
        if not itens_selecionados:
            QMessageBox.information(self, "Aviso", "Por favor, selecione um agendamento na tabela para editar.")
            return
        linha_selecionada = itens_selecionados[0].row()
        item_horario = self.tabela_agenda.item(linha_selecionada, 0)
        agendamento_existente = item_horario.data(Qt.UserRole)
        if not agendamento_existente:
            QMessageBox.information(self, "Aviso", "Este horário está vago. Dê um clique duplo para criar um novo agendamento.")
            return
        dialog = EntregaDialog(self.usuario_logado, self.date, entrega_data=agendamento_existente, parent=self)
        if dialog.exec_() == QDialog.Accepted:
            data = dialog.get_data()
            database.atualizar_entrega(agendamento_existente['ID'], agendamento_existente['HORARIO'], data['status_id'], data['cliente_id'], data['responsavel'], data['observacoes'], data['is_retificacao'], self.usuario_logado['USERNAME'])
            self.carregar_agenda_dia()
            self.parent().populate_calendar()

    def excluir_agendamento(self):
        itens_selecionados = self.tabela_agenda.selectedItems()
        if not itens_selecionados:
            QMessageBox.information(self, "Aviso", "Por favor, selecione um agendamento na tabela para excluir.")
            return
        linha_selecionada = itens_selecionados[0].row()
        item_horario = self.tabela_agenda.item(linha_selecionada, 0)
        agendamento_existente = item_horario.data(Qt.UserRole)
        if not agendamento_existente:
            QMessageBox.information(self, "Aviso", "Não há agendamento para excluir neste horário.")
            return
        reply = QMessageBox.question(self, "Confirmar Exclusão", f"Tem certeza que deseja excluir o agendamento para '{agendamento_existente['NOME_CLIENTE']}' às {agendamento_existente['HORARIO']}?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            database.deletar_entrega(agendamento_existente['ID'], self.usuario_logado['USERNAME'])
            self.carregar_agenda_dia()
            self.parent().populate_calendar()


class DialogoResultadosBusca(QDialog):
    def __init__(self, resultados, usuario_logado, parent=None):
        super().__init__(parent)
        self.usuario_logado = usuario_logado
        self.setWindowTitle("Resultados da Busca")
        self.setMinimumSize(800, 500)

        layout = QVBoxLayout(self)
        self.tabela_resultados = QTableWidget()
        self.tabela_resultados.setColumnCount(5)
        self.tabela_resultados.setHorizontalHeaderLabels(["Data", "Horário", "Cliente", "Status", "Responsável"])
        self.tabela_resultados.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tabela_resultados.setSelectionBehavior(QTableWidget.SelectRows)
        self.tabela_resultados.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tabela_resultados.setSortingEnabled(True) # Habilita a ordenação por coluna

        layout.addWidget(QLabel("Clique duplo em uma linha para abrir o dia do agendamento."))
        layout.addWidget(self.tabela_resultados)

        self.popular_tabela(resultados)

        self.tabela_resultados.cellDoubleClicked.connect(self.abrir_agendamento)

    def popular_tabela(self, resultados):
        self.tabela_resultados.setRowCount(len(resultados))
        for i, res in enumerate(resultados):
            # Converte a data do DB para o formato dd/MM/yyyy para exibição
            data_obj = QDate.fromString(str(res['DATA_VENCIMENTO']), 'yyyy-MM-dd')
            item_data = QTableWidgetItem(data_obj.toString('dd/MM/yyyy'))
            
            # Armazena a data original para a funcionalidade de clique duplo
            item_data.setData(Qt.UserRole, res)

            self.tabela_resultados.setItem(i, 0, item_data)
            self.tabela_resultados.setItem(i, 1, QTableWidgetItem(res.get('HORARIO', '')))
            self.tabela_resultados.setItem(i, 2, QTableWidgetItem(res.get('NOME_CLIENTE', '')))
            self.tabela_resultados.setItem(i, 3, QTableWidgetItem(res.get('NOME_STATUS', 'N/A')))
            self.tabela_resultados.setItem(i, 4, QTableWidgetItem(res.get('RESPONSAVEL', '')))

    def abrir_agendamento(self, row, column):
        item_selecionado = self.tabela_resultados.item(row, 0)
        dados_agendamento = item_selecionado.data(Qt.UserRole)
        
        data_do_agendamento = QDate.fromString(str(dados_agendamento['DATA_VENCIMENTO']), 'yyyy-MM-dd')

        # Abre a janela do dia correspondente
        dialog = DayViewDialog(data_do_agendamento, self.usuario_logado, self.parent())
        dialog.exec_()

class ConfigDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Configurações")
        self.settings = QSettings()
        main_layout = QVBoxLayout(self)

        db_groupbox = QGroupBox("Configurações do Banco de Dados")
        db_main_layout = QVBoxLayout()
        self.radio_local = QRadioButton("Banco de Dados Local (arquivo no computador)")
        self.radio_remoto = QRadioButton("Banco de Dados Remoto (em um servidor)")
        db_main_layout.addWidget(self.radio_local)
        db_main_layout.addWidget(self.radio_remoto)
        self.local_db_group = QGroupBox("Conexão Local")
        local_db_layout = QFormLayout()
        caminho_layout = QHBoxLayout()
        self.caminho_local_edit = QLineEdit()
        procurar_btn = QPushButton("Procurar...")
        procurar_btn.clicked.connect(self.procurar_arquivo_db)
        caminho_layout.addWidget(self.caminho_local_edit)
        caminho_layout.addWidget(procurar_btn)
        local_db_layout.addRow("Caminho do Arquivo (.FDB):", caminho_layout)
        self.local_db_group.setLayout(local_db_layout)
        db_main_layout.addWidget(self.local_db_group)
        self.remoto_db_group = QGroupBox("Conexão Remota")
        remoto_db_layout = QFormLayout()
        self.host_remoto_edit = QLineEdit()
        self.porta_remota_spin = QSpinBox()
        self.porta_remota_spin.setRange(1, 65535)
        self.caminho_remoto_edit = QLineEdit()
        self.caminho_remoto_edit.setPlaceholderText("Ex: C:\\Bancos\\CALENDARIO.FDB")
        self.usuario_remoto_edit = QLineEdit()
        self.senha_remota_edit = QLineEdit()
        self.senha_remota_edit.setEchoMode(QLineEdit.Password)
        remoto_db_layout.addRow("Host (IP do servidor):", self.host_remoto_edit)
        remoto_db_layout.addRow("Porta:", self.porta_remota_spin)
        remoto_db_layout.addRow("Caminho do Arquivo no Servidor:", self.caminho_remoto_edit)
        remoto_db_layout.addRow("Usuário:", self.usuario_remoto_edit)
        remoto_db_layout.addRow("Senha:", self.senha_remota_edit)
        self.remoto_db_group.setLayout(remoto_db_layout)
        db_main_layout.addWidget(self.remoto_db_group)
        db_groupbox.setLayout(db_main_layout)

        modo_groupbox = QGroupBox("Modo de Definição de Horários")
        modo_layout = QVBoxLayout()
        self.radio_auto = QRadioButton("Gerar horários automaticamente")
        self.radio_manual = QRadioButton("Inserir horários manualmente")
        modo_layout.addWidget(self.radio_auto)
        modo_layout.addWidget(self.radio_manual)
        modo_groupbox.setLayout(modo_layout)
        self.auto_groupbox = QGroupBox("Configuração Automática")
        auto_layout = QFormLayout()
        self.inicio_edit = QTimeEdit()
        self.fim_edit = QTimeEdit()
        self.intervalo_spin = QSpinBox()
        self.intervalo_spin.setRange(5, 120)
        self.intervalo_spin.setSuffix(" minutos")
        auto_layout.addRow("Hora de Início:", self.inicio_edit)
        auto_layout.addRow("Hora de Fim:", self.fim_edit)
        auto_layout.addRow("Intervalo:", self.intervalo_spin)
        self.auto_groupbox.setLayout(auto_layout)
        self.manual_groupbox = QGroupBox("Configuração Manual")
        manual_layout = QFormLayout()
        self.lista_manual_edit = QLineEdit()
        self.lista_manual_edit.setPlaceholderText("Ex: 08:00, 09:15, 10:30, 14:00")
        manual_layout.addRow("Lista de horários (separados por vírgula):", self.lista_manual_edit)
        self.manual_groupbox.setLayout(manual_layout)
        
        geral_groupbox = QGroupBox("Geral")
        geral_layout = QFormLayout()
        self.lembrete_spin = QSpinBox()
        self.lembrete_spin.setRange(1, 60)
        self.lembrete_spin.setSuffix(" minutos")
        geral_layout.addRow("Avisar de agendamentos com antecedência de:", self.lembrete_spin)
        
        self.refresh_interval_spin = QSpinBox()
        self.refresh_interval_spin.setRange(5, 300)
        self.refresh_interval_spin.setSuffix(" segundos")
        geral_layout.addRow("Intervalo de atualização automática:", self.refresh_interval_spin)
        
        geral_groupbox.setLayout(geral_layout)

        main_layout.addWidget(db_groupbox)
        main_layout.addWidget(modo_groupbox)
        main_layout.addWidget(self.auto_groupbox)
        main_layout.addWidget(self.manual_groupbox)
        main_layout.addWidget(geral_groupbox)
        botoes = QHBoxLayout()
        salvar_btn = QPushButton("Salvar")
        cancelar_btn = QPushButton("Cancelar")
        botoes.addStretch()
        botoes.addWidget(salvar_btn)
        botoes.addWidget(cancelar_btn)
        main_layout.addLayout(botoes)
        salvar_btn.clicked.connect(self.salvar)
        cancelar_btn.clicked.connect(self.reject)
        self.radio_auto.toggled.connect(self.atualizar_modo_horario_visivel)
        self.radio_local.toggled.connect(self.atualizar_modo_db_visivel)
        self.carregar_configs()

    def procurar_arquivo_db(self):
        caminho, _ = QFileDialog.getSaveFileName(self, "Selecionar ou Criar Arquivo de Banco de Dados", "", "Firebird Database (*.fdb)")
        if caminho: self.caminho_local_edit.setText(caminho)

    def carregar_configs(self):
        modo_horario = self.settings.value("horarios/modo", "automatico")
        if modo_horario == "manual": self.radio_manual.setChecked(True)
        else: self.radio_auto.setChecked(True)
        self.inicio_edit.setTime(QTime.fromString(self.settings.value("horarios/hora_inicio", "08:30"), 'HH:mm'))
        self.fim_edit.setTime(QTime.fromString(self.settings.value("horarios/hora_fim", "17:30"), 'HH:mm'))
        self.intervalo_spin.setValue(self.settings.value("horarios/intervalo_minutos", 30, type=int))
        self.lista_manual_edit.setText(self.settings.value("horarios/lista_manual", "09:00,10:00,11:00"))
        self.lembrete_spin.setValue(self.settings.value("geral/minutos_lembrete", 15, type=int))
        self.refresh_interval_spin.setValue(self.settings.value("geral/refresh_intervalo_segundos", 30, type=int))
        modo_db = self.settings.value("database/modo", "local")
        if modo_db == "remoto": self.radio_remoto.setChecked(True)
        else: self.radio_local.setChecked(True)
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
        default_path = os.path.join(base_path, 'Data', 'CALENDARIO.FDB')
        self.caminho_local_edit.setText(self.settings.value("database/caminho_local", default_path))
        self.host_remoto_edit.setText(self.settings.value("database/host_remoto", "localhost"))
        self.porta_remota_spin.setValue(self.settings.value("database/porta_remota", 3050, type=int))
        self.caminho_remoto_edit.setText(self.settings.value("database/caminho_remoto", ""))
        self.usuario_remoto_edit.setText(self.settings.value("database/usuario", "SYSDBA"))
        self.senha_remota_edit.setText(self.settings.value("database/senha", "masterkey"))
        self.atualizar_modo_horario_visivel()
        self.atualizar_modo_db_visivel()

    def atualizar_modo_horario_visivel(self):
        self.auto_groupbox.setEnabled(self.radio_auto.isChecked())
        self.manual_groupbox.setEnabled(not self.radio_auto.isChecked())
            
    def atualizar_modo_db_visivel(self):
        self.local_db_group.setEnabled(self.radio_local.isChecked())
        self.remoto_db_group.setEnabled(not self.radio_local.isChecked())

    def salvar(self):
        # Salva as configurações de Horários
        if self.radio_auto.isChecked():
            self.settings.setValue("horarios/modo", "automatico")
        else:
            self.settings.setValue("horarios/modo", "manual")
        self.settings.setValue("horarios/hora_inicio", self.inicio_edit.time().toString('HH:mm'))
        self.settings.setValue("horarios/hora_fim", self.fim_edit.time().toString('HH:mm'))
        self.settings.setValue("horarios/intervalo_minutos", self.intervalo_spin.value())
        self.settings.setValue("horarios/lista_manual", self.lista_manual_edit.text())
        
        # Salva as configurações Gerais
        self.settings.setValue("geral/minutos_lembrete", self.lembrete_spin.value())
        self.settings.setValue("geral/refresh_intervalo_segundos", self.refresh_interval_spin.value())

        # Salva as configurações de Banco de Dados
        if self.radio_local.isChecked():
            self.settings.setValue("database/modo", "local")
        else:
            self.settings.setValue("database/modo", "remoto")
        self.settings.setValue("database/caminho_local", self.caminho_local_edit.text())
        self.settings.setValue("database/host_remoto", self.host_remoto_edit.text())
        self.settings.setValue("database/porta_remota", self.porta_remota_spin.value())
        self.settings.setValue("database/caminho_remoto", self.caminho_remoto_edit.text())
        self.settings.setValue("database/usuario", self.usuario_remoto_edit.text())
        self.settings.setValue("database/senha", self.senha_remota_edit.text())

        QMessageBox.information(self, "Sucesso", "Configurações salvas com sucesso.\nAlgumas alterações podem exigir que o programa seja reiniciado.")
        self.accept()

class SecretReportDialog(QDialog):
    def __init__(self, raw_data, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Relatório de Atividade por Usuário (Secreto)")
        self.setMinimumSize(600, 700)
        layout = QVBoxLayout(self)
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Agrupamento", "Contagem"])
        self.tree.setColumnWidth(0, 350)
        layout.addWidget(self.tree)
        processed_data = self.processar_dados(raw_data)
        self.popular_arvore(processed_data)
        self.tree.expandAll()

    def processar_dados(self, raw_data):
        data = collections.defaultdict(lambda: collections.defaultdict(lambda: collections.defaultdict(int)))
        for row in raw_data:
            mes = row['MES']
            responsavel = row['RESPONSAVEL']
            status = row['NOME_STATUS']
            contagem = row['CONTAGEM']
            if status.lower() in ['feito', 'feito e enviado']:
                status_agrupado = 'Concluído (Total)'
            else:
                status_agrupado = status
            data[mes][responsavel][status_agrupado] += contagem
            data[mes][responsavel]['Total'] += contagem
        return data

    def popular_arvore(self, data):
        self.tree.clear()
        for mes, usuarios in sorted(data.items(), reverse=True):
            mes_item = QTreeWidgetItem(self.tree, [f"Mês: {mes}"])
            mes_item.setFont(0, QFont('Arial', 10, QFont.Bold))
            for usuario, status_counts in sorted(usuarios.items()):
                total_usuario = status_counts.pop('Total', 0)
                usuario_item = QTreeWidgetItem(mes_item, [f"Usuário: {usuario}", str(total_usuario)])
                usuario_item.setFont(0, QFont('Arial', 9, QFont.Bold))
                for status, contagem in sorted(status_counts.items()):
                    status_item = QTreeWidgetItem(usuario_item, [f"   - {status}", str(contagem)])

class RelatorioDialog(QDialog):
    def __init__(self, usuario_logado, parent=None):
        super().__init__(parent)
        self.usuario_logado = usuario_logado
        self.setWindowTitle("Gerar Relatórios")
        self.setMinimumWidth(400)
        layout = QVBoxLayout(self)
        form_layout = QFormLayout()
        self.tipo_relatorio_combo = QComboBox()
        self.tipo_relatorio_combo.addItems(["Relatório de Agendamentos", "Relatório de Logs de Atividade"])
        self.data_inicio_edit = QDateEdit(QDate.currentDate().addDays(-30))
        self.data_inicio_edit.setCalendarPopup(True)
        self.data_fim_edit = QDateEdit(QDate.currentDate())
        self.data_fim_edit.setCalendarPopup(True)
        self.filtros_agendamento_group = QGroupBox("Filtros de Agendamento")
        filtros_agendamento_layout = QVBoxLayout()
        self.status_list_widget = QListWidget()
        self.status_list_widget.setSelectionMode(QListWidget.MultiSelection)
        filtros_agendamento_layout.addWidget(QLabel("Filtrar por Status (deixe sem selecionar para incluir todos):"))
        filtros_agendamento_layout.addWidget(self.status_list_widget)
        self.filtros_agendamento_group.setLayout(filtros_agendamento_layout)
        self.filtros_logs_group = QGroupBox("Filtros de Logs")
        filtros_logs_layout = QFormLayout()
        self.usuario_combo = QComboBox()
        filtros_logs_layout.addRow("Filtrar por Usuário:", self.usuario_combo)
        self.filtros_logs_group.setLayout(filtros_logs_layout)
        self.carregar_filtros()
        form_layout.addRow("Tipo de Relatório:", self.tipo_relatorio_combo)
        form_layout.addRow("Data de Início:", self.data_inicio_edit)
        form_layout.addRow("Data de Fim:", self.data_fim_edit)
        layout.addLayout(form_layout)
        layout.addWidget(self.filtros_agendamento_group)
        layout.addWidget(self.filtros_logs_group)
        gerar_btn = QPushButton("Gerar Relatório")
        layout.addWidget(gerar_btn, alignment=Qt.AlignCenter)
        gerar_btn.clicked.connect(self.gerar_relatorio)
        self.tipo_relatorio_combo.currentIndexChanged.connect(self.atualizar_filtros_visiveis)
        self.atualizar_filtros_visiveis()
    
    def keyPressEvent(self, event):
        if event.key() == Qt.Key_F12 and self.usuario_logado['USERNAME'].lower() == 'admin':
            self.abrir_relatorio_secreto()
        else:
            super().keyPressEvent(event)

    def abrir_relatorio_secreto(self):
        dados_brutos = database.get_estatisticas_por_usuario_e_status()
        dialog = SecretReportDialog(dados_brutos, self)
        dialog.exec_()
    
    def carregar_filtros(self):
        for status in database.listar_status():
            item = QListWidgetItem(status['NOME'])
            item.setData(Qt.UserRole, status['ID'])
            self.status_list_widget.addItem(item)
        self.usuario_combo.addItem("Todos")
        for username in database.listar_usuarios():
            self.usuario_combo.addItem(username)

    def atualizar_filtros_visiveis(self):
        if "Agendamentos" in self.tipo_relatorio_combo.currentText():
            self.filtros_agendamento_group.setVisible(True)
            self.filtros_logs_group.setVisible(False)
        else:
            self.filtros_agendamento_group.setVisible(False)
            self.filtros_logs_group.setVisible(True)

    def gerar_relatorio(self):
        data_inicio = self.data_inicio_edit.date().toString("yyyy-MM-dd")
        data_fim = self.data_fim_edit.date().toString("yyyy-MM-dd")
        tipo_relatorio = self.tipo_relatorio_combo.currentText()
        dialog = QFileDialog(self)
        dialog.setAcceptMode(QFileDialog.AcceptSave)
        dialog.setNameFilter("Arquivos PDF (*.pdf);;Arquivos CSV (*.csv)")
        dialog.setDefaultSuffix("pdf")
        if "Agendamentos" in tipo_relatorio:
            dialog.selectFile(f"Relatorio_Agendamentos_{data_inicio}_a_{data_fim}")
            status_selecionados_ids = [item.data(Qt.UserRole) for item in self.status_list_widget.selectedItems()]
            dados = database.get_entregas_filtradas(data_inicio, data_fim, status_selecionados_ids)
            titulo = f"Relatório de Agendamentos de {self.data_inicio_edit.date().toString('dd/MM/yyyy')} a {self.data_fim_edit.date().toString('dd/MM/yyyy')}"
            if not dados:
                QMessageBox.information(self, "Aviso", "Nenhum agendamento encontrado para os filtros selecionados.")
                return
            if dialog.exec_():
                nome_arquivo = dialog.selectedFiles()[0]
                if nome_arquivo.endswith(".pdf"):
                    export.exportar_para_pdf(dados, nome_arquivo, titulo)
                elif nome_arquivo.endswith(".csv"):
                    export.exportar_para_csv(dados, nome_arquivo)
        else:
            dialog.selectFile(f"Relatorio_Logs_{data_inicio}_a_{data_fim}")
            usuario_selecionado = self.usuario_combo.currentText()
            dados = database.get_logs_filtrados(data_inicio, data_fim, usuario_selecionado)
            titulo = f"Relatório de Logs de {self.data_inicio_edit.date().toString('dd/MM/yyyy')} a {self.data_fim_edit.date().toString('dd/MM/yyyy')}"
            if not dados:
                QMessageBox.information(self, "Aviso", "Nenhum log encontrado para os filtros selecionados.")
                return
            if dialog.exec_():
                nome_arquivo = dialog.selectedFiles()[0]
                if nome_arquivo.endswith(".pdf"):
                    export.exportar_logs_pdf(dados, nome_arquivo, titulo)
                else:
                    export.exportar_logs_csv(dados, nome_arquivo)
        self.accept()

# main.py - Adicione esta classe inteira no seu arquivo

class DialogoSobre(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Sobre o Agendador de Sintegras")
        self.setMinimumSize(800, 650)

        layout = QVBoxLayout(self)

        self.browser = QTextBrowser()
        self.browser.setOpenExternalLinks(True) # Para futuros links
        layout.addWidget(self.browser)

        # --- CONTEÚDO DA DOCUMENTAÇÃO ---
        # Usamos HTML para formatar o texto
        documentacao_html = f"""
            <h1>Agendador de Sintegras - Versão {VERSAO_ATUAL}</h1>
            <p>Este é um guia completo sobre como utilizar todas as funcionalidades do sistema.</p>
            <hr>

            <h2>Tela Principal (Calendário)</h2>
            <ul>
                <li><b>Visualização:</b> A tela principal mostra o calendário do mês. Cada dia exibe a contagem de agendamentos e uma cor baseada no status do primeiro agendamento do dia.</li>
                <li><b>Cores dos Dias:</b>
                    <ul>
                        <li><b>Verde:</b> Dia normal sem agendamentos.</li>
                        <li><b>Laranja:</b> Fim de semana.</li>
                        <li><b>Roxo:</b> Feriado.</li>
                        <li><b>Outras cores:</b> Indicam que há agendamentos, com a cor do status predominante.</li>
                    </ul>
                </li>
                <li><b>Interação:</b>
                    <ul>
                        <li><b>Clique Duplo:</b> Abre a visão detalhada da agenda daquele dia.</li>
                        <li><b>Clique Direito:</b> Permite marcar ou desmarcar um dia como feriado.</li>
                    </ul>
                </li>
                <li><b>Dashboard:</b> O painel no topo exibe estatísticas rápidas do mês atual, como a porcentagem de clientes atendidos e o total de retificações.</li>
            </ul>
            <hr>

            <h2>Agendamentos</h2>
            <p>Na janela de um dia específico (aberta com clique duplo), você pode gerenciar os agendamentos hora a hora.</p>
            <ul>
                <li><b>Criar:</b> Dê um clique duplo em um horário vago para abrir a janela de criação de agendamento.</li>
                <li><b>Editar:</b> Selecione um agendamento existente e clique em "Editar" (ou dê um clique duplo nele).</li>
                <li><b>Campos:</b>
                    <ul>
                        <li><b>Cliente:</b> Comece a digitar o nome para filtrar e selecionar. O campo inicia vazio por padrão.</li>
                        <li><b>Status:</b> Define o estado atual do agendamento (Pendente, Feito, etc.).</li>
                        <li><b>Responsável:</b> Preenchido automaticamente com o seu usuário e não pode ser alterado.</li>
                        <li><b>Ações:</b> Use os botões "Copiar Email/Local" e "Copiar Rascunho" para agilizar seu trabalho.</li>
                    </ul>
                </li>
            </ul>
            <hr>

            <h2>Gerenciamento de Clientes</h2>
            <p>Acessível pelo botão "Gerenciar Clientes" na tela principal.</p>
            <ul>
                <li><b>Adicionar/Editar/Excluir:</b> Funções padrão para gerenciar sua base de clientes.</li>
                <li><b>Telefones:</b> É possível cadastrar até 2 telefones por cliente no formato (xx) xxxxx-xxxx.</li>
                <li><b>Verificar Pendentes:</b> Mostra uma lista de todos os clientes que ainda não tiveram um agendamento no mês atual.</li>
                <li><b>Verificar Inativos:</b> Gera um relatório de clientes que não agendam há 3 meses ou mais.</li>
                <li><b>Importar de XLSX:</b> Permite importar uma lista de clientes a partir de uma planilha Excel.</li>
            </ul>
            <hr>
            
            <h2>Outras Funcionalidades</h2>
            <ul>
                <li><b>Busca Rápida:</b> A barra de busca na parte inferior da tela principal permite encontrar agendamentos por nome de cliente, responsável ou palavras nas observações.</li>
                <li><b>Gerenciar Status:</b> Permite criar, editar ou excluir os status de agendamento e suas cores correspondentes.</li>
                <li><b>Gerar Relatório:</b> Permite exportar relatórios detalhados de agendamentos ou logs de atividade (quem fez o quê) para PDF ou CSV.</li>
                <li><b>Configurações:</b> Permite ajustar as configurações de conexão do banco de dados (local ou remoto) e os horários de trabalho. O acesso é protegido pela senha do usuário 'admin'.</li>
            </ul>
            <hr>
            
            <h2>Atalhos de Teclado</h2>
            <ul>
                <li><b>F3 (na janela de Clientes):</b> Abre a tela de análise de performance de clientes, com estatísticas e rankings.</li>
            </ul>
        """
        self.browser.setHtml(documentacao_html)

        # --- Botão de Fechar ---
        botoes_layout = QHBoxLayout()
        fechar_btn = QPushButton("Fechar")
        fechar_btn.clicked.connect(self.accept)
        botoes_layout.addStretch()
        botoes_layout.addWidget(fechar_btn)
        botoes_layout.addStretch()
        layout.addLayout(botoes_layout)

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
            self.username_edit.setText(usuario['USERNAME'])
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
        self.setModal(True)
        layout = QVBoxLayout(self)
        label = QLabel("Para prosseguir, por favor, digite sua senha de administrador:")
        self.senha_edit = QLineEdit()
        self.senha_edit.setEchoMode(QLineEdit.Password)
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
            item = QTableWidgetItem(username)
            item.setData(Qt.UserRole, username) 
            self.tabela.setItem(i, 0, item)

    def adicionar_usuario(self):
        dialog = DialogoUsuario(self)
        if dialog.exec_() == QDialog.Accepted:
            username, password = dialog.get_dados()
            if username and password:
                if database.criar_usuario(username, password, self.usuario_logado['USERNAME']):
                    QMessageBox.information(self, "Sucesso", "Usuário criado!")
                    self.carregar_usuarios()
                else:
                    QMessageBox.warning(self, "Erro", "Nome de usuário já existe!")

    def editar_usuario(self):
        linha = self.tabela.currentRow()
        if linha < 0:
            QMessageBox.warning(self, "Ação Necessária", "Por favor, selecione um usuário para editar.")
            return
        
        username_selecionado = self.tabela.item(linha, 0).data(Qt.UserRole)
        usuario_obj = database.get_usuario_por_nome(username_selecionado)
        if not usuario_obj:
            QMessageBox.critical(self, "Erro", "Usuário não encontrado no banco de dados."); return

        dialogo_confirmacao = ConfirmacaoSenhaDialog(self)
        if dialogo_confirmacao.exec_() == QDialog.Accepted:
            senha_digitada = dialogo_confirmacao.get_senha()
            administrador_logado = self.usuario_logado['USERNAME']

            if database.verificar_senha_usuario_atual(administrador_logado, senha_digitada):
                dialog = DialogoUsuario(self, usuario=usuario_obj)
                if dialog.exec_() == QDialog.Accepted:
                    novo_username, nova_senha = dialog.get_dados()
                    if not nova_senha:
                        QMessageBox.warning(self, "Senha Obrigatória", "O campo de senha não pode estar vazio ao editar.")
                        return
                    if novo_username and nova_senha:
                        database.atualizar_usuario(usuario_obj["ID"], novo_username, nova_senha, self.usuario_logado["USERNAME"])
                        QMessageBox.information(self, "Sucesso", "Usuário atualizado!")
                        self.carregar_usuarios()
            else:
                QMessageBox.critical(self, "Falha na Autenticação", "Senha de administrador incorreta. Ação cancelada.")

    def excluir_usuario(self):
        linha = self.tabela.currentRow()
        if linha < 0:
            QMessageBox.warning(self, "Ação Necessária", "Por favor, selecione um usuário para excluir.")
            return
        
        username_selecionado = self.tabela.item(linha, 0).data(Qt.UserRole)
        usuario_para_excluir = database.get_usuario_por_nome(username_selecionado)

        if not usuario_para_excluir:
            QMessageBox.critical(self, "Erro", "Não foi possível encontrar o usuário no banco de dados."); return

        if usuario_para_excluir['ID'] == 1:
            QMessageBox.warning(self, "Ação Proibida", "O usuário administrador principal não pode ser excluído."); return

        if usuario_para_excluir['ID'] == self.usuario_logado['ID']:
            QMessageBox.warning(self, "Ação Proibida", "Você não pode excluir seu próprio usuário."); return

        dialogo_confirmacao = ConfirmacaoSenhaDialog(self)
        if dialogo_confirmacao.exec_() == QDialog.Accepted:
            senha_digitada = dialogo_confirmacao.get_senha()
            administrador_logado = self.usuario_logado['USERNAME']

            if database.verificar_senha_usuario_atual(administrador_logado, senha_digitada):
                reply = QMessageBox.question(self, "Confirmar Exclusão", f"Tem certeza que deseja excluir permanentemente o usuário '{username_selecionado}'?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                if reply == QMessageBox.Yes:
                    if database.deletar_usuario(usuario_para_excluir["ID"], self.usuario_logado["USERNAME"]):
                        QMessageBox.information(self, "Sucesso", "Usuário excluído!")
                        self.carregar_usuarios()
                    else:
                        QMessageBox.critical(self, "Erro", "Ocorreu um erro ao excluir o usuário.")
            else:
                QMessageBox.critical(self, "Falha na Autenticação", "Senha de administrador incorreta. Ação cancelada.")

class CalendarWindow(QMainWindow):
    def __init__(self, usuario):
        super().__init__()
        self.usuario_atual = usuario
        self.original_titulo = f"Agendador Mensal - Bem-vindo, {self.usuario_atual['USERNAME']}!"
        self.setWindowTitle(self.original_titulo)
        self.setGeometry(100, 100, 1100, 800)
        self.current_date = datetime.now()
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)
        self.notificados_nesta_sessao = set()
        self.feriados = {}
        self.setup_ui()
        self.setup_tray_icon()
        self.setup_timer_notificacoes()
        self.populate_calendar()
        self.verificar_atualizacao()

        settings = QSettings()
        intervalo_segundos = settings.value("geral/refresh_intervalo_segundos", 30, type=int)
        intervalo_ms = intervalo_segundos * 1000
        
        self.refresh_timer = QTimer(self)
        self.refresh_timer.timeout.connect(self.populate_calendar)
        self.refresh_timer.start(intervalo_ms)
        print(f"🔄 Atualização automática configurada para cada {intervalo_segundos} segundos.")

        self.center()

    def center(self):
        """Centraliza a janela na tela principal."""
        screen_geometry = QApplication.primaryScreen().availableGeometry()
        window_geometry = self.frameGeometry()
        window_geometry.moveCenter(screen_geometry.center())
        self.move(window_geometry.topLeft())

    def _get_cor_porcentagem(self, porcentagem):
        if porcentagem < 40: return "#d9534f"
        elif porcentagem < 80: return "#f0ad4e"
        else: return "#5cb85c"

    def _get_cor_pendentes(self, pendentes, total):
        if total == 0: return "#5bc0de"
        ratio = pendentes / total if total > 0 else 0
        if ratio > 0.6: return "#d9534f"
        elif ratio > 0.2: return "#f0ad4e"
        else: return "#5cb85c"
        
    def gerar_horarios_dinamicos(self, para_data):
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
    
    def _atualizar_sugestoes(self):
        data_busca = QDate.currentDate()
        for _ in range(30):
            dia_da_semana = data_busca.dayOfWeek()
            tipo_feriado = self.feriados.get(data_busca)
            if dia_da_semana in [6, 7] or tipo_feriado == "nacional":
                data_busca = data_busca.addDays(1)
                continue
            todos_horarios = self.gerar_horarios_dinamicos(data_busca)
            agendamentos_dia = database.get_entregas_por_dia(data_busca.toString("yyyy-MM-dd"))
            horarios_ocupados = set(agendamentos_dia.keys())
            horarios_livres = [h for h in todos_horarios if h not in horarios_ocupados]
            if horarios_livres:
                hoje = QDate.currentDate()
                if data_busca == hoje: texto_dia = "Hoje"
                elif data_busca == hoje.addDays(1): texto_dia = "Amanhã"
                else: texto_dia = data_busca.toString("dd/MM")
                sugestoes_str = ", ".join(horarios_livres[:5])
                if len(horarios_livres) > 5: sugestoes_str += "..."
                self.sugestoes_label.setText(f"<b>{texto_dia}:</b><br>{sugestoes_str}")
                return
            data_busca = data_busca.addDays(1)
        self.sugestoes_label.setText("Nenhum horário livre<br>encontrado.")

    def populate_calendar(self):
        for i in reversed(range(self.calendar_grid.count())):
            widget = self.calendar_grid.itemAt(i).widget()
            if widget is not None:
                widget.setParent(None)
        year = self.current_date.year
        month = self.current_date.month
        self.month_label.setText(f"<b>{self.current_date.strftime('%B de %Y')}</b>")
        stats = database.get_estatisticas_mensais(year, month)
        total_clientes = database.get_total_clientes()
        ids_clientes_atendidos = database.get_clientes_com_agendamento_concluido_no_mes(year, month)
        concluidos_mes = stats.get('CONCLUIDOS', 0)
        retificados_mes = stats.get('RETIFICADOS', 0)
        clientes_pendentes = total_clientes - len(ids_clientes_atendidos)
        porcentagem_conclusao = (concluidos_mes / total_clientes) * 100 if total_clientes > 0 else 0
        cor_porcentagem = self._get_cor_porcentagem(porcentagem_conclusao)
        cor_pendentes = self._get_cor_pendentes(clientes_pendentes, total_clientes)
        texto_dashboard = f"""
        <table style='margin-left: auto; margin-right: auto;'>
            <tr>
                <td align='center' style='padding-right: 20px;'>
                    Concluídos no Mês: <b style='color:green;'>{concluidos_mes}</b><br/>
                    <span style='font-size:14px; color:{cor_porcentagem};'>
                        <b>{porcentagem_conclusao:.1f}%</b>
                    </span>
                </td>
                <td align='center' style='padding-left: 20px;'>
                    Retificados no Mês: <b style='color:#17a2b8;'>{retificados_mes}</b><br/>
                    <span style='font-size:14px; color:{cor_pendentes};'>
                        <b>{clientes_pendentes} pendentes</b>
                    </span>
                </td>
            </tr>
        </table>
        """
        self.dashboard_label.setText(texto_dashboard)
        status_dias = database.get_status_dias_para_mes(year, month)
        month_calendar = calendar.monthcalendar(year, month)
        for week_num, week in enumerate(month_calendar):
            for day_num, day in enumerate(week):
                if day != 0:
                    date = QDate(year, month, day)
                    info_do_dia = status_dias.get(day)
                    cell = DayCellWidget(date, info_do_dia, day_num, self)
                    cell.clicked.connect(self.open_day_view)
                    self.calendar_grid.addWidget(cell, week_num, day_num)
        self._atualizar_sugestoes()

    def setup_ui(self):
            nav_layout = QHBoxLayout()
            prev_btn = QPushButton("< Mês Anterior")
            self.month_label = QLabel()
            self.month_label.setAlignment(Qt.AlignCenter)
            next_btn = QPushButton("Próximo Mês >")
            prev_btn.clicked.connect(self.prev_month)
            next_btn.clicked.connect(self.next_month)
            nav_layout.addWidget(prev_btn)
            nav_layout.addWidget(self.month_label)
            nav_layout.addWidget(next_btn)
            self.main_layout.addLayout(nav_layout)
            info_panel_layout = QHBoxLayout()
            sugestoes_group = QGroupBox("Sugestões de Agendamento")
            sugestoes_layout = QVBoxLayout(sugestoes_group)
            
            self.sugestoes_label = SuggestionLabel("Buscando horários...")
            
            self.sugestoes_label.setAlignment(Qt.AlignLeft | Qt.AlignTop)
            sugestoes_layout.addWidget(self.sugestoes_label)
            sugestoes_layout.addStretch(1)
            sugestoes_group.setFixedWidth(300)
            self.dashboard_label = QLabel()
            self.dashboard_label.setAlignment(Qt.AlignCenter)
            self.dashboard_label.setStyleSheet("padding-bottom: 15px;")
            spacer_widget = QWidget()
            spacer_widget.setFixedWidth(300)
            info_panel_layout.addWidget(sugestoes_group, 0, Qt.AlignBottom)
            info_panel_layout.addStretch(1)
            info_panel_layout.addWidget(self.dashboard_label, 0, Qt.AlignVCenter)
            info_panel_layout.addStretch(1)
            info_panel_layout.addWidget(spacer_widget, 0, Qt.AlignBottom)
            self.main_layout.addLayout(info_panel_layout)
            header_layout = QGridLayout()
            dias = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]
            for i, dia in enumerate(dias):
                header_layout.addWidget(QLabel(f"<b>{dia}</b>", alignment=Qt.AlignCenter), 0, i)
            self.main_layout.addLayout(header_layout)
            self.calendar_grid = QGridLayout()
            self.calendar_grid.setSpacing(0)
            self.main_layout.addLayout(self.calendar_grid)
            
            busca_layout = QHBoxLayout()
            busca_layout.addWidget(QLabel("<b>Busca Rápida:</b>"))
            self.busca_global_edit = QLineEdit()
            self.busca_global_edit.setPlaceholderText("Digite o nome do cliente, responsável ou observação...")
            
            self.completer = QCompleter(self)
            self.completer.setCaseSensitivity(Qt.CaseInsensitive)
            self.completer.setFilterMode(Qt.MatchContains)
            self.busca_global_edit.setCompleter(self.completer)
            self._atualizar_completer_busca()

            self.busca_global_btn = QPushButton("Buscar")
            busca_layout.addWidget(self.busca_global_edit)
            busca_layout.addWidget(self.busca_global_btn)
            self.main_layout.addLayout(busca_layout)

            self.busca_global_btn.clicked.connect(self.realizar_busca_global)
            self.busca_global_edit.returnPressed.connect(self.realizar_busca_global)
            
            # --- INÍCIO DA CORREÇÃO ---
            action_layout = QHBoxLayout()
            config_btn = QPushButton("Configurações")
            config_btn.clicked.connect(self.abrir_configuracoes)
            
            sobre_btn = QPushButton("Sobre")
            sobre_btn.clicked.connect(self.abrir_dialogo_sobre)
            
            clientes_btn = QPushButton("Gerenciar Clientes")
            clientes_btn.clicked.connect(self.gerenciar_clientes)
            
            status_btn = QPushButton("Gerenciar Status")
            status_btn.clicked.connect(self.manage_status)
            
            usuarios_btn = QPushButton("Gerenciar Usuários")
            usuarios_btn.clicked.connect(self.gerenciar_usuarios)
            
            relatorio_btn = QPushButton("Gerar Relatório")
            relatorio_btn.clicked.connect(self.abrir_dialogo_relatorio)
            
            # Adicionando os botões na ordem correta
            action_layout.addWidget(config_btn)
            action_layout.addWidget(sobre_btn) # Botão "Sobre" adicionado aqui
            action_layout.addStretch() 
            action_layout.addWidget(clientes_btn)
            action_layout.addWidget(status_btn)
            action_layout.addWidget(usuarios_btn)
            action_layout.addWidget(relatorio_btn)
            
            self.main_layout.addLayout(action_layout)

    def _atualizar_completer_busca(self):
    
        try:
            clientes = database.listar_clientes()
            nomes_clientes = [cliente['NOME'] for cliente in clientes if cliente.get('NOME')]
        
            modelo = QStringListModel()
            modelo.setStringList(nomes_clientes)
            self.completer.setModel(modelo)
            print("✅ Modelo do auto-complete atualizado com sucesso.")
        except Exception as e:
            print(f"⚠️ Erro ao atualizar o auto-complete: {e}")

    def realizar_busca_global(self):
        termo = self.busca_global_edit.text().strip()
        if not termo:
            QMessageBox.warning(self, "Busca Inválida", "Por favor, digite algo para buscar.")
            return

        resultados = database.buscar_agendamentos_globais(termo)

        if not resultados:
            QMessageBox.information(self, "Nenhum Resultado", f"Nenhum agendamento encontrado para o termo '{termo}'.")
        else:
            # Passa a janela principal (self) como pai para o diálogo
            dialog = DialogoResultadosBusca(resultados, self.usuario_atual, self)
            dialog.exec_()    

    def verificar_atualizacao(self):
        self.update_thread = UpdateCheckerThread()
        self.update_thread.update_found.connect(self.mostrar_dialogo_atualizacao)
        self.update_thread.check_finished.connect(self.finalizar_verificacao)
        self.update_thread.start()
        self.setWindowTitle(f"{self.original_titulo} (Verificando atualizações...)")

    def mostrar_dialogo_atualizacao(self, nova_versao):
        titulo = "Atualização Disponível!"
        mensagem = (f"Uma nova versão ({nova_versao}) do programa está disponível!\n\nDeseja atualizar agora?")
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle(titulo); msg_box.setText(mensagem); msg_box.setIcon(QMessageBox.Information)
        sim_btn = msg_box.addButton("Sim, Atualizar Agora", QMessageBox.YesRole)
        tarde_btn = msg_box.addButton("Atualizar Mais Tarde", QMessageBox.RejectRole)
        msg_box.setDefaultButton(tarde_btn); msg_box.exec_()
        if msg_box.clickedButton() == sim_btn: self.baixar_e_instalar_atualizacao(nova_versao)
        self.setWindowTitle(self.original_titulo)

    def finalizar_verificacao(self, update_encontrado):
        if not update_encontrado: self.setWindowTitle(self.original_titulo)

    def baixar_e_instalar_atualizacao(self, nova_versao):
            self.setWindowTitle(f"{self.original_titulo} (Baixando atualização...)")
            try:
                nome_arquivo = f"Calendario-v{nova_versao}.exe"
                url_download = f"https://github.com/Azzaleh/Agendador-Sintegras/releases/download/v{nova_versao}/{nome_arquivo}"
                
                # --- INÍCIO DA CORREÇÃO ---
                # 1. Descobre o caminho do executável que está rodando
                pasta_do_executavel = os.path.dirname(sys.executable)
                
                # 2. Define o caminho completo para salvar o arquivo na mesma pasta
                caminho_salvar = os.path.join(pasta_do_executavel, nome_arquivo)
                # --- FIM DA CORREÇÃO ---
                
                from urllib.request import urlretrieve
                urlretrieve(url_download, caminho_salvar)
                
                # Abre o novo instalador/executável e fecha o programa atual
                subprocess.Popen([caminho_salvar])
                QApplication.instance().quit()
            except Exception as e:
                QMessageBox.critical(self, "Erro na Atualização", f"Não foi possível baixar ou executar a atualização.\n\nErro: {e}")
                self.setWindowTitle(self.original_titulo)

    def setup_tray_icon(self):
        self.tray_icon = QSystemTrayIcon(self)
        caminho_icone = os.path.join('imagens', 'icon.ico')
        icon = QIcon(caminho_icone)
        if icon.isNull(): self.tray_icon.setIcon(self.style().standardIcon(QStyle.SP_MessageBoxInformation))
        else: self.tray_icon.setIcon(icon)
        self.tray_icon.setToolTip("Agendador de Entregas"); self.tray_icon.show()

    def setup_timer_notificacoes(self):
        settings = QSettings()
        minutos_lembrete = settings.value("geral/minutos_lembrete", 15, type=int)
        self.timer = QTimer(self)
        self.timer.timeout.connect(lambda: self.verificar_agendamentos_proximos(minutos_lembrete))
        self.timer.start(60000)

    def verificar_agendamentos_proximos(self, minutos_antecedencia):
        agora = datetime.now()
        limite = agora + timedelta(minutes=minutos_antecedencia)
        data_hoje = agora.strftime("%Y-%m-%d"); hora_inicio = agora.strftime("%H:%M"); hora_fim = limite.strftime("%H:%M")
        agendamentos = database.get_entregas_no_intervalo(data_hoje, hora_inicio, hora_fim)
        for ag in agendamentos:
            if ag['ID'] not in self.notificados_nesta_sessao and ag.get('NOME_STATUS', '').lower() == 'pendente':
                titulo = f"Lembrete de Agendamento ({ag['HORARIO']})"
                mensagem = f"Cliente: {ag['NOME_CLIENTE']}\nStatus: {ag.get('NOME_STATUS', 'N/A')}"
                self.tray_icon.showMessage(titulo, mensagem, QSystemTrayIcon.Information, 15000); self.notificados_nesta_sessao.add(ag['ID'])

    def closeEvent(self, event):
        self.tray_icon.hide(); event.accept()

    def prev_month(self):
        self.current_date = self.current_date.replace(day=1) - timedelta(days=1)
        self.populate_calendar()

    def next_month(self):
        _, days_in_month = calendar.monthrange(self.current_date.year, self.current_date.month)
        self.current_date = self.current_date.replace(day=1) + timedelta(days=days_in_month)
        self.populate_calendar()

    def open_day_view(self, date):
        if self.feriados.get(date) == "nacional":
            QMessageBox.warning(self, "Feriado Nacional", "Não é permitido agendar em um feriado nacional."); return
        dialog = DayViewDialog(date, self.usuario_atual, self); dialog.exec_()

    def gerenciar_clientes(self):
        dialog = JanelaClientes(self.usuario_atual, self); dialog.exec_()
        self.populate_calendar()
        self._atualizar_completer_busca()

    def manage_status(self):
        dialog = StatusDialog(self.usuario_atual, self); dialog.exec_()
        self.populate_calendar()

    def abrir_configuracoes(self):
        dialogo_confirmacao = ConfirmacaoSenhaDialog(self)
        if dialogo_confirmacao.exec_() == QDialog.Accepted:
            senha_digitada = dialogo_confirmacao.get_senha()
            
            if database.verificar_usuario('admin', senha_digitada):
                dialog_config = ConfigDialog(self)
                dialog_config.exec_()
            else:
                QMessageBox.warning(self, "Acesso Negado", "A senha do administrador está incorreta.")

    def abrir_dialogo_relatorio(self):
        dialog = RelatorioDialog(self.usuario_atual, self); dialog.exec_()

    def gerenciar_usuarios(self):
        dialog = JanelaUsuarios(self.usuario_atual, self); dialog.exec_()

    def abrir_dialogo_sobre(self):
            """Abre a janela de documentação do sistema."""
            dialog = DialogoSobre(self)
            dialog.exec_()

if __name__ == '__main__':
    QApplication.setOrganizationName("Data Servis")
    QApplication.setApplicationName("Agendador")
    
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(True)

    try:
        database.iniciar_db()
    except Exception as e:
        QMessageBox.critical(None, "Erro Crítico de Banco de Dados",
                             f"Não foi possível conectar ou inicializar o banco de dados.\n"
                             f"Verifique se o serviço do Firebird está sendo executado e se as configurações estão corretas.\n\n"
                             f"Erro técnico: {e}")
        sys.exit(1) 

    login = LoginDialog()
    
    if login.exec_() == QDialog.Accepted:
        usuario_logado = login.usuario_logado

        window = CalendarWindow(usuario_logado)
        window.show()
        sys.exit(app.exec_())
    else:

        sys.exit(0)
