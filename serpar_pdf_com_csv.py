import sys
import os
import csv
import shutil

from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton,
    QFileDialog, QTextEdit, QVBoxLayout,
    QHBoxLayout, QMessageBox
)


class SeparadorXML(QWidget):

    def __init__(self):
        super().__init__()

        self.pasta_xmls = ""
        self.arquivo_csv = ""

        self.setWindowTitle("Separador de XMLs por CSV")
        self.setGeometry(200, 200, 800, 500)

        self.init_ui()

    def init_ui(self):

        layout = QVBoxLayout()

        # Pasta XMLs
        linha_pasta = QHBoxLayout()

        self.lbl_pasta = QLabel("Nenhuma pasta de XMLs selecionada")

        btn_pasta = QPushButton("Selecionar Pasta XMLs")
        btn_pasta.clicked.connect(self.selecionar_pasta)

        linha_pasta.addWidget(btn_pasta)
        linha_pasta.addWidget(self.lbl_pasta)

        # CSV
        linha_csv = QHBoxLayout()

        self.lbl_csv = QLabel("Nenhum CSV selecionado")

        btn_csv = QPushButton("Selecionar CSV")
        btn_csv.clicked.connect(self.selecionar_csv)

        linha_csv.addWidget(btn_csv)
        linha_csv.addWidget(self.lbl_csv)

        # Botão processar
        btn_processar = QPushButton("Separar XMLs")
        btn_processar.clicked.connect(self.processar)

        # Resultado
        self.resultado = QTextEdit()
        self.resultado.setReadOnly(True)

        layout.addLayout(linha_pasta)
        layout.addLayout(linha_csv)
        layout.addWidget(btn_processar)
        layout.addWidget(self.resultado)

        self.setLayout(layout)

    def selecionar_pasta(self):

        pasta = QFileDialog.getExistingDirectory(
            self,
            "Selecionar Pasta com XMLs"
        )

        if pasta:
            self.pasta_xmls = pasta
            self.lbl_pasta.setText(pasta)

    def selecionar_csv(self):

        arquivo, _ = QFileDialog.getOpenFileName(
            self,
            "Selecionar CSV",
            "",
            "CSV (*.csv);;Todos Arquivos (*)"
        )

        if arquivo:
            self.arquivo_csv = arquivo
            self.lbl_csv.setText(arquivo)

    def ler_chaves_csv(self):

        chaves = set()

        with open(self.arquivo_csv, "r", encoding="utf-8") as arq:

            leitor = csv.reader(arq)

            for linha in leitor:

                for coluna in linha:

                    texto = coluna.strip()

                    texto = texto.replace("NFe", "").strip()

                    numeros = ''.join(filter(str.isdigit, texto))

                    if len(numeros) == 44:
                        chaves.add(numeros)

        return chaves

    def processar(self):

        if not self.pasta_xmls:
            QMessageBox.warning(
                self,
                "Aviso",
                "Selecione a pasta dos XMLs."
            )
            return

        if not self.arquivo_csv:
            QMessageBox.warning(
                self,
                "Aviso",
                "Selecione o CSV."
            )
            return

        try:

            self.resultado.clear()

            chaves_csv = self.ler_chaves_csv()

            pasta_saida = os.path.join(
                self.pasta_xmls,
                "XMLS_SEPARADOS"
            )

            os.makedirs(pasta_saida, exist_ok=True)

            encontrados = 0
            nao_encontrados = []

            arquivos = os.listdir(self.pasta_xmls)

            for arquivo in arquivos:

                if arquivo.lower().endswith(".xml"):

                    nome = os.path.splitext(arquivo)[0]

                    nome = nome.replace("-nfe", "")

                    chave = ''.join(filter(str.isdigit, nome))

                    if chave in chaves_csv:

                        origem = os.path.join(
                            self.pasta_xmls,
                            arquivo
                        )

                        destino = os.path.join(
                            pasta_saida,
                            arquivo
                        )

                        shutil.copy2(origem, destino)

                        encontrados += 1

            # Verifica quais não foram encontrados
            xmls_encontrados = set()

            for arquivo in os.listdir(pasta_saida):

                nome = os.path.splitext(arquivo)[0]

                nome = nome.replace("-nfe", "")

                chave = ''.join(filter(str.isdigit, nome))

                xmls_encontrados.add(chave)

            for chave in chaves_csv:

                if chave not in xmls_encontrados:
                    nao_encontrados.append(chave)

            # Resultado
            self.resultado.append(
                "=== PROCESSAMENTO FINALIZADO ===\n"
            )

            self.resultado.append(
                f"Total de chaves no CSV: {len(chaves_csv)}"
            )

            self.resultado.append(
                f"XMLs copiados: {encontrados}"
            )

            self.resultado.append(
                f"Não encontrados: {len(nao_encontrados)}\n"
            )

            if nao_encontrados:

                self.resultado.append(
                    "CHAVES NÃO ENCONTRADAS:\n"
                )

                for chave in nao_encontrados:
                    self.resultado.append(chave)

            QMessageBox.information(
                self,
                "Concluído",
                f"XMLs separados com sucesso!\n\nPasta criada:\n{pasta_saida}"
            )

        except Exception as e:

            QMessageBox.critical(
                self,
                "Erro",
                str(e)
            )


if __name__ == "__main__":

    app = QApplication(sys.argv)

    janela = SeparadorXML()
    janela.show()

    sys.exit(app.exec_())