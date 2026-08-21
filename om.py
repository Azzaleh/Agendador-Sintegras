import os
import sys
import xml.etree.ElementTree as ET
import fdb
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QFileDialog, QVBoxLayout,
    QTextEdit, QMessageBox, QTabWidget, QTableWidget, QTableWidgetItem,
    QLabel, QHeaderView
)
from PyQt5.QtCore import Qt

class VerificadorXML(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Verificador XML x Firebird")
        self.resize(1200, 750)
        layout_principal = QVBoxLayout()

        self.lbl_info = QLabel("Selecione uma pasta contendo XMLs")
        self.btn_pasta = QPushButton("Selecionar Pasta XML")
        self.btn_pasta.clicked.connect(self.selecionar_pasta)

        self.tabs = QTabWidget()
        
        # Aba Log
        self.aba_log = QWidget()
        layout_log = QVBoxLayout()
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        layout_log.addWidget(self.log)
        self.aba_log.setLayout(layout_log)

        # Aba Tags
        self.aba_tags = QWidget()
        layout_tags = QVBoxLayout()
        self.tabela_tags = QTableWidget()
        self.tabela_tags.setColumnCount(20)
        self.tabela_tags.setHorizontalHeaderLabels([
            "NOTA", "vBC", "vICMS", "vICMSDeson", "vFCP", "vBCST", "vST",
            "vFCPST", "vFCPSTRet", "vProd", "vFrete", "vSeg", "vDesc",
            "vII", "vIPI", "vIPIDevol", "vPIS", "vCOFINS", "vOutro", "vNF"
        ])
        self.tabela_tags.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.tabela_tags.setAlternatingRowColors(True)
        layout_tags.addWidget(self.tabela_tags)
        self.aba_tags.setLayout(layout_tags)

        self.tabs.addTab(self.aba_log, "Verificação")
        self.tabs.addTab(self.aba_tags, "Tags XML")

        layout_principal.addWidget(self.lbl_info)
        layout_principal.addWidget(self.btn_pasta)
        layout_principal.addWidget(self.tabs)
        self.setLayout(layout_principal)

    def conectar_banco(self):
        return fdb.connect(
            dsn=r'localhost:C:\Users\D SERVIS\Desktop\BD.FDB',
            user='SYSDBA',
            password='masterkey',
            charset='WIN1252'
        )

    def selecionar_pasta(self):
        pasta = QFileDialog.getExistingDirectory(self, "Selecionar Pasta XML")
        if pasta:
            self.processar_xmls(pasta)

    def adicionar_tags_fiscais(self, numero_nf, impostos):
        linha = self.tabela_tags.rowCount()
        self.tabela_tags.insertRow(linha)
        colunas = [
            numero_nf, impostos.get('vBC', ''), impostos.get('vICMS', ''),
            impostos.get('vICMSDeson', ''), impostos.get('vFCP', ''),
            impostos.get('vBCST', ''), impostos.get('vST', ''),
            impostos.get('vFCPST', ''), impostos.get('vFCPSTRet', ''),
            impostos.get('vProd', ''), impostos.get('vFrete', ''),
            impostos.get('vSeg', ''), impostos.get('vDesc', ''),
            impostos.get('vII', ''), impostos.get('vIPI', ''),
            impostos.get('vIPIDevol', ''), impostos.get('vPIS', ''),
            impostos.get('vCOFINS', ''), impostos.get('vOutro', ''),
            impostos.get('vNF', '')
        ]
        for coluna, valor in enumerate(colunas):
            item = QTableWidgetItem(str(valor))
            item.setTextAlignment(Qt.AlignCenter)
            self.tabela_tags.setItem(linha, coluna, item)

    def processar_xmls(self, pasta):
        self.log.clear()
        self.tabela_tags.setRowCount(0)

        try:
            con = self.conectar_banco()
            cur = con.cursor()
        except Exception as e:
            QMessageBox.critical(self, "Erro ao conectar", str(e))
            return

        arquivos = [arq for arq in os.listdir(pasta) if arq.lower().endswith(".xml")]
        ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
        total_divergencias = 0

        for arquivo in arquivos:
            caminho = os.path.join(pasta, arquivo)
            try:
                tree = ET.parse(caminho)
                root = tree.getroot()
                numero_nf = root.find('.//nfe:nNF', ns).text
                valor_total_xml = float(root.find('.//nfe:vNF', ns).text)

                # Coleta de Tags para a Tabela
                tags_desejadas = ['vBC', 'vICMS', 'vICMSDeson', 'vFCP', 'vBCST', 'vST', 'vFCPST', 'vFCPSTRet', 'vProd', 'vFrete', 'vSeg', 'vDesc', 'vII', 'vIPI', 'vIPIDevol', 'vPIS', 'vCOFINS', 'vOutro', 'vNF']
                impostos = {tag: (root.find(f'.//nfe:{tag}', ns).text if root.find(f'.//nfe:{tag}', ns) is not None else '') for tag in tags_desejadas}
                self.adicionar_tags_fiscais(numero_nf, impostos)

                self.log.append("-" * 50)
                self.log.append(f"NOTA: {numero_nf} | ARQUIVO: {arquivo}")

                # 1. CONSULTA TOTAL DA NOTA (Cabeçalho)
                cur.execute("SELECT VLRTOTNF FROM NOTAFISCAL WHERE NUMNF = ?", [numero_nf])
                res_nota = cur.fetchone()
                
                if not res_nota:
                    self.log.append("❌ NOTA NÃO ENCONTRADA NO BANCO")
                    continue

                # 2. CONSULTA SOMA DOS PRODUTOS (Itens) - RESOLVE O PROBLEMA DA NOTA 44189
                cur.execute("SELECT SUM(VLRTOTAL) FROM NOTAFISCAL_PRODUTO WHERE NUMNF = ?", [numero_nf])
                res_soma_produtos = cur.fetchone()
                valor_soma_itens_banco = float(res_soma_produtos[0]) if res_soma_produtos[0] else 0.0

                # 3. COMPARAÇÃO DOS TOTAIS
                if round(valor_total_xml, 2) != round(valor_soma_itens_banco, 2):
                    self.log.append(f"⚠️ DIVERGÊNCIA ENCONTRADA:")
                    self.log.append(f"   XML Total: {valor_total_xml:.2f}")
                    self.log.append(f"   Banco (Soma de todos os Itens): {valor_soma_itens_banco:.2f}")
                    self.log.append(f"   Diferença: {valor_total_xml - valor_soma_itens_banco:.2f}")
                    total_divergencias += 1
                else:
                    self.log.append(f"✅ VALOR TOTAL OK (Soma de todos os itens bate com XML)")

                # Opcional: Detalhar se algum item específico não existe
                # (Lógica para listar itens individuais se necessário)

            except Exception as e:
                self.log.append(f"❌ ERRO NO ARQUIVO {arquivo}: {str(e)}")

        con.close()
        self.log.append("=" * 50)
        self.log.append(f"VERIFICAÇÃO FINALIZADA | DIVERGÊNCIAS: {total_divergencias}")
        QMessageBox.information(self, "Concluído", "Processamento finalizado.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    janela = VerificadorXML()
    janela.show()
    sys.exit(app.exec_())