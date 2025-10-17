import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel

class App(QWidget):
    def __init__(self):
        super().__init__()
        self.title = 'Filtrar Clientes Ativos'
        self.left = 100
        self.top = 100
        self.width = 400
        self.height = 200
        self.clientes_file = ''
        self.status_file = ''
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        layout = QVBoxLayout()

        self.label_clientes = QLabel('Arquivo de Clientes: Nenhum selecionado')
        layout.addWidget(self.label_clientes)
        self.btn_clientes = QPushButton('Selecionar Arquivo de Clientes', self)
        self.btn_clientes.clicked.connect(self.select_clientes_file)
        layout.addWidget(self.btn_clientes)

        self.label_status = QLabel('Arquivo de Status: Nenhum selecionado')
        layout.addWidget(self.label_status)
        self.btn_status = QPushButton('Selecionar Arquivo de Status', self)
        self.btn_status.clicked.connect(self.select_status_file)
        layout.addWidget(self.btn_status)

        self.btn_process = QPushButton('Processar e Salvar', self)
        self.btn_process.clicked.connect(self.process_files)
        layout.addWidget(self.btn_process)

        self.setLayout(layout)
        self.show()

    def select_clientes_file(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self, "Selecionar Arquivo de Clientes", "", "Todos os Arquivos (*);;Arquivos Excel (*.xlsx);;Arquivos CSV (*.csv)", options=options)
        if fileName:
            self.clientes_file = fileName
            self.label_clientes.setText(f'Arquivo de Clientes: {fileName}')

    def select_status_file(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self, "Selecionar Arquivo de Status", "", "Todos os Arquivos (*);;Arquivos Excel (*.xlsx);;Arquivos CSV (*.csv)", options=options)
        if fileName:
            self.status_file = fileName
            self.label_status.setText(f'Arquivo de Status: {fileName}')

    def process_files(self):
        if not self.clientes_file or not self.status_file:
            print("Por favor, selecione os dois arquivos.")
            return

        try:
            # Tenta ler como Excel, se der erro, tenta como CSV
            try:
                df_clientes = pd.read_excel(self.clientes_file)
            except:
                df_clientes = pd.read_csv(self.clientes_file)

            try:
                df_status = pd.read_excel(self.status_file)
            except:
                df_status = pd.read_csv(self.status_file)
                
            # IMPORTANTE: Ajuste os nomes das colunas aqui se necessário
            # Coloque o nome exato da coluna que contém os nomes dos clientes
            coluna_clientes_principal = 'Clientes'  # Do arquivo com todos os clientes
            coluna_clientes_status = 'CLIENTE'     # Do arquivo de status

            # Renomeia a coluna do arquivo de status para o mesmo nome da do arquivo principal
            df_status.rename(columns={coluna_clientes_status: coluna_clientes_principal}, inplace=True)
            
            # Filtra o dataframe principal para manter apenas os clientes que estão no arquivo de status
            clientes_ativos = pd.merge(df_clientes, df_status[[coluna_clientes_principal]], on=coluna_clientes_principal, how='inner')

            # Salva o novo arquivo
            options = QFileDialog.Options()
            saveFileName, _ = QFileDialog.getSaveFileName(self, "Salvar Arquivo de Clientes Ativos", "", "Arquivos Excel (*.xlsx);;Arquivos CSV (*.csv)", options=options)
            if saveFileName:
                if saveFileName.endswith('.xlsx'):
                    clientes_ativos.to_excel(saveFileName, index=False)
                else:
                    clientes_ativos.to_csv(saveFileName, index=False)
                print(f"Arquivo salvo com sucesso em: {saveFileName}")

        except Exception as e:
            print(f"Ocorreu um erro: {e}")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())