# export.py
import csv
from PyQt5.QtWidgets import QMessageBox

try:
    from reportlab.lib.pagesizes import letter, landscape
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import inch
    from reportlab.lib import colors
    REPORTLAB_DISPONIVEL = True
except ImportError:
    REPORTLAB_DISPONIVEL = False

def exportar_para_csv(entregas, nome_arquivo):
    """Exporta uma lista de entregas para um arquivo CSV."""
    try:
        with open(nome_arquivo, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            # Cabeçalho
            writer.writerow([
                'Data', 'Horário', 'Cliente', 'Status', 'Responsável',
                'Contato', 'Tipo de Envio', 'Observações'
            ])
            
            # Dados
            for entrega in entregas:
                writer.writerow([
                    entrega.get('data_vencimento', ''),
                    entrega.get('horario', ''),
                    entrega.get('nome_cliente', ''),
                    entrega.get('nome_status', 'N/A'),
                    entrega.get('responsavel', ''),
                    entrega.get('contato', ''),
                    entrega.get('tipo_envio', ''),
                    entrega.get('observacoes', '')
                ])
        QMessageBox.information(None, "Sucesso", f"Dados exportados com sucesso para:\n{nome_arquivo}")
    except Exception as e:
        QMessageBox.critical(None, "Erro", f"Não foi possível exportar para CSV: {e}")


def exportar_para_pdf(entregas, nome_arquivo, titulo_relatorio):
    """Exporta uma lista de entregas para um arquivo PDF."""
    if not REPORTLAB_DISPONIVEL:
        QMessageBox.warning(None, "Atenção", 
            "A biblioteca 'reportlab' não está instalada.\n"
            "Não é possível exportar para PDF.\n\n"
            "Instale com: pip install reportlab")
        return

    try:
        doc = SimpleDocTemplate(nome_arquivo, pagesize=landscape(letter))
        styles = getSampleStyleSheet()
        story = []

        # Título
        story.append(Paragraph(titulo_relatorio, styles['h1']))
        story.append(Spacer(1, 0.2 * inch))

        # Conteúdo em formato de tabela
        dados_tabela = [
            # Cabeçalho da tabela
            ["Data", "Horário", "Cliente", "Status", "Responsável", "Contato"]
        ]
        
        for entrega in entregas:
            dados_tabela.append([
                entrega.get('data_vencimento', ''),
                entrega.get('horario', ''),
                entrega.get('nome_cliente', ''),
                entrega.get('nome_status', 'N/A'),
                entrega.get('responsavel', ''),
                entrega.get('contato', '')
            ])
        
        tabela = Table(dados_tabela)
        
        # Estilo da tabela
        estilo = TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.grey),
            ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0,0), (-1,0), 12),
            ('BACKGROUND', (0,1), (-1,-1), colors.beige),
            ('GRID', (0,0), (-1,-1), 1, colors.black)
        ])
        tabela.setStyle(estilo)
        
        story.append(tabela)
        doc.build(story)
        QMessageBox.information(None, "Sucesso", f"Dados exportados com sucesso para:\n{nome_arquivo}")

    except Exception as e:
        QMessageBox.critical(None, "Erro", f"Não foi possível exportar para PDF: {e}")