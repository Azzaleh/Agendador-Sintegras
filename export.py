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
            writer.writerow(['Data', 'Horário', 'Cliente', 'Status', 'Responsável', 'Contato', 'Tipo de Envio', 'Observações'])
            for entrega in entregas:
                writer.writerow([
                    entrega.get('data_vencimento', ''), entrega.get('horario', ''),
                    entrega.get('nome_cliente', ''), entrega.get('nome_status', 'N/A'),
                    entrega.get('responsavel', ''), entrega.get('contato', ''),
                    entrega.get('tipo_envio', ''), entrega.get('observacoes', '')
                ])
        QMessageBox.information(None, "Sucesso", f"Dados exportados com sucesso para:\n{nome_arquivo}")
    except Exception as e:
        QMessageBox.critical(None, "Erro", f"Não foi possível exportar para CSV: {e}")


def exportar_para_pdf(entregas, nome_arquivo, titulo_relatorio):
    """Exporta uma lista de entregas para um arquivo PDF."""
    if not REPORTLAB_DISPONIVEL:
        QMessageBox.warning(None, "Atenção", "A biblioteca 'reportlab' não está instalada.\nInstale com: pip install reportlab")
        return
    try:
        doc = SimpleDocTemplate(nome_arquivo, pagesize=landscape(letter))
        styles = getSampleStyleSheet()
        story = []
        story.append(Paragraph(titulo_relatorio, styles['h1']))
        story.append(Spacer(1, 0.2 * inch))
        dados_tabela = [["Data", "Horário", "Cliente", "Status", "Responsável", "Contato"]]
        for entrega in entregas:
            dados_tabela.append([
                entrega.get('data_vencimento', ''), entrega.get('horario', ''),
                entrega.get('nome_cliente', ''), entrega.get('nome_status', 'N/A'),
                entrega.get('responsavel', ''), entrega.get('contato', '')
            ])
        tabela = Table(dados_tabela)
        estilo = TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.grey), ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0,0), (-1,0), 12), ('BACKGROUND', (0,1), (-1,-1), colors.beige),
            ('GRID', (0,0), (-1,-1), 1, colors.black)
        ])
        tabela.setStyle(estilo)
        story.append(tabela)
        doc.build(story)
        QMessageBox.information(None, "Sucesso", f"Dados exportados com sucesso para:\n{nome_arquivo}")
    except Exception as e:
        QMessageBox.critical(None, "Erro", f"Não foi possível exportar para PDF: {e}")

# --- NOVAS FUNÇÕES PARA EXPORTAR LOGS ---
def exportar_logs_csv(logs, nome_arquivo):
    """Exporta uma lista de logs para um arquivo CSV."""
    try:
        with open(nome_arquivo, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(['Timestamp', 'Usuário', 'Ação', 'Detalhes'])
            for log in logs:
                writer.writerow([
                    log.get('timestamp', ''), log.get('usuario_nome', ''),
                    log.get('acao', ''), log.get('detalhes', '')
                ])
        QMessageBox.information(None, "Sucesso", f"Logs exportados com sucesso para:\n{nome_arquivo}")
    except Exception as e:
        QMessageBox.critical(None, "Erro", f"Não foi possível exportar para CSV: {e}")

def exportar_logs_pdf(logs, nome_arquivo, titulo_relatorio):
    """Exporta uma lista de logs para um arquivo PDF."""
    if not REPORTLAB_DISPONIVEL:
        QMessageBox.warning(None, "Atenção", "A biblioteca 'reportlab' não está instalada.")
        return
    try:
        doc = SimpleDocTemplate(nome_arquivo, pagesize=letter)
        styles = getSampleStyleSheet()
        story = []
        story.append(Paragraph(titulo_relatorio, styles['h1']))
        story.append(Spacer(1, 0.2 * inch))
        dados_tabela = [["Timestamp", "Usuário", "Ação", "Detalhes"]]
        for log in logs:
            dados_tabela.append([
                Paragraph(log.get('timestamp', ''), styles['BodyText']),
                Paragraph(log.get('usuario_nome', ''), styles['BodyText']),
                Paragraph(log.get('acao', ''), styles['BodyText']),
                Paragraph(log.get('detalhes', ''), styles['BodyText'])
            ])
        tabela = Table(dados_tabela, colWidths=[2*inch, 1*inch, 1.5*inch, 3*inch])
        estilo = TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.darkblue), ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('ALIGN', (0,0), (-1,-1), 'LEFT'), ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'), ('BOTTOMPADDING', (0,0), (-1,0), 12),
            ('BACKGROUND', (0,1), (-1,-1), colors.antiquewhite), ('GRID', (0,0), (-1,-1), 1, colors.black)
        ])
        tabela.setStyle(estilo)
        story.append(tabela)
        doc.build(story)
        QMessageBox.information(None, "Sucesso", f"Logs exportados com sucesso para:\n{nome_arquivo}")
    except Exception as e:
        QMessageBox.critical(None, "Erro", f"Não foi possível exportar para PDF: {e}")