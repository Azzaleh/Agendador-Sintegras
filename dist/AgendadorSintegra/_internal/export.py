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
        
        # O cabeçalho da tabela continua igual
        dados_tabela = [["Data", "Horário", "Cliente", "Status", "Responsável", "Contato"]]
        
        # --- INÍCIO DA CORREÇÃO ---
        for entrega in entregas:
            # CORREÇÃO: Usar chaves em MAIÚSCULAS para corresponder ao que vem do banco
            data_obj = entrega.get('DATA_VENCIMENTO') # <--- CORREÇÃO
            data_str = data_obj.strftime('%d/%m/%Y') if data_obj else ''

            dados_tabela.append([
                data_str, 
                entrega.get('HORARIO', ''),          # <--- CORREÇÃO
                entrega.get('NOME_CLIENTE', ''),     # <--- CORREÇÃO
                entrega.get('NOME_STATUS', 'N/A'),   # <--- CORREÇÃO
                entrega.get('RESPONSAVEL', ''),      # <--- CORREÇÃO
                entrega.get('CONTATO', '')           # <--- CORREÇÃO
            ])
        # --- FIM DA CORREÇÃO ---
            
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

def exportar_logs_csv(logs, nome_arquivo):
    """Exporta uma lista de logs para um arquivo CSV."""
    try:
        with open(nome_arquivo, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(['Timestamp', 'Usuário', 'Ação', 'Detalhes'])
            for log in logs:
                # CORREÇÃO: Usar chaves em MAIÚSCULAS
                data_obj = log.get('DATA_HORA')
                data_str = data_obj.strftime('%d/%m/%Y %H:%M:%S') if data_obj else ''
                writer.writerow([
                    data_str,
                    log.get('USUARIO_NOME', ''),
                    log.get('ACAO', ''),
                    log.get('DETALHES', '')
                ])
        QMessageBox.information(None, "Sucesso", f"Logs exportados com sucesso para:\n{nome_arquivo}")
    except Exception as e:
        QMessageBox.critical(None, "Erro", f"Não foi possível exportar para CSV: {e}")

# export.py - Substitua também esta função

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
            # CORREÇÃO: Usar chaves em MAIÚSCULAS
            data_objeto = log.get('DATA_HORA')
            if data_objeto:
                data_str = data_objeto.strftime('%d/%m/%Y %H:%M:%S')
            else:
                data_str = ''
            
            dados_tabela.append([
                Paragraph(data_str, styles['BodyText']),
                Paragraph(log.get('USUARIO_NOME', ''), styles['BodyText']),
                Paragraph(log.get('ACAO', ''), styles['BodyText']),
                Paragraph(log.get('DETALHES', ''), styles['BodyText'])
            ])

        tabela = Table(dados_tabela, colWidths=[1.8*inch, 1*inch, 1.5*inch, 3.2*inch])
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
        
        # --- INÍCIO DA CORREÇÃO ---
        for log in logs:
            # CORREÇÃO: Usar as chaves em MAIÚSCULAS que vêm do banco de dados
            data_objeto = log.get('DATA_HORA') # <--- CORREÇÃO
            
            if data_objeto:
                data_str = data_objeto.strftime('%d/%m/%Y %H:%M:%S')
            else:
                data_str = ''
            
            dados_tabela.append([
                Paragraph(data_str, styles['BodyText']),
                Paragraph(log.get('USUARIO_NOME', ''), styles['BodyText']), # <--- CORREÇÃO
                Paragraph(log.get('ACAO', ''), styles['BodyText']),         # <--- CORREÇÃO
                Paragraph(log.get('DETALHES', ''), styles['BodyText'])     # <--- CORREÇÃO
            ])
        # --- FIM DA CORREÇÃO ---

        tabela = Table(dados_tabela, colWidths=[1.8*inch, 1*inch, 1.5*inch, 3.2*inch])
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