# export.py
import csv
from PyQt5.QtWidgets import QMessageBox

try:
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import inch
    REPORTLAB_DISPONIVEL = True
except ImportError:
    REPORTLAB_DISPONIVEL = False

def exportar_para_csv(eventos, nome_arquivo):
    """Exporta uma lista de eventos para um arquivo CSV."""
    try:
        with open(nome_arquivo, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            # Cabeçalho
            writer.writerow(['Data', 'Horario', 'Descrição', 'Categoria'])
            
            # Dados
            for dia in sorted(eventos.keys()):
                for evento in eventos[dia]:
                    writer.writerow([
                        evento['data'],
                        evento['horario'],
                        evento['descricao'],
                        evento.get('categoria_nome', 'Sem Categoria')
                    ])
        QMessageBox.information(None, "Sucesso", f"Dados exportados para {nome_arquivo}")
    except Exception as e:
        QMessageBox.critical(None, "Erro", f"Não foi possível exportar para CSV: {e}")


def exportar_para_pdf(eventos, nome_arquivo, titulo):
    """Exporta uma lista de eventos para um arquivo PDF."""
    if not REPORTLAB_DISPONIVEL:
        QMessageBox.warning(None, "Atenção", 
            "A biblioteca 'reportlab' não está instalada.\n"
            "Não é possível exportar para PDF.\n\n"
            "Instale com: pip install reportlab")
        return

    try:
        doc = SimpleDocTemplate(nome_arquivo, pagesize=letter)
        styles = getSampleStyleSheet()
        story = []

        # Título
        story.append(Paragraph(titulo, styles['h1']))
        story.append(Spacer(1, 0.2 * inch))

        # Conteúdo
        for dia in sorted(eventos.keys()):
            # Adiciona um subtítulo para o dia
            data_formatada = eventos[dia][0]['data']
            story.append(Paragraph(f"<b>Dia: {data_formatada}</b>", styles['h2']))
            
            for evento in eventos[dia]:
                cor = evento.get('cor_hex', '#000000')
                categoria = evento.get('categoria_nome', 'N/A')
                texto_evento = (
                    f"<font color='{cor}'>●</font> "
                    f"<b>{evento['horario']}</b> - {evento['descricao']} "
                    f"<i>({categoria})</i>"
                )
                story.append(Paragraph(texto_evento, styles['BodyText']))
            
            story.append(Spacer(1, 0.2 * inch))
        
        doc.build(story)
        QMessageBox.information(None, "Sucesso", f"Dados exportados para {nome_arquivo}")

    except Exception as e:
        QMessageBox.critical(None, "Erro", f"Não foi possível exportar para PDF: {e}")