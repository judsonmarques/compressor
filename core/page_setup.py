"""
Módulo para configuração e ajuste de páginas no documento Word (.docx).
Padroniza tamanho da página (A4/Letter), margens e distâncias de cabeçalho/rodapé.
"""

from docx import Document
from docx.shared import Cm, Inches
from core.config import FormattingConfig
from typing import Dict, Any


def setup_pages(doc: Document, config: FormattingConfig) -> Dict[str, Any]:
    """
    Aplica as configurações de margem, tamanho de papel e cabeçalhos em todas as seções.
    
    Retorna métricas da operação.
    """
    sections_count = len(doc.sections)
    
    for section in doc.sections:
        # Configuração do tamanho de página
        if config.page_size.upper() == "A4":
            # Padrão A4: 21.0 x 29.7 cm
            section.page_width = Cm(21.0)
            section.page_height = Cm(29.7)
        elif config.page_size.upper() == "LETTER":
            # Padrão Carta: 8.5 x 11.0 polegadas
            section.page_width = Inches(8.5)
            section.page_height = Inches(11.0)

        # Configuração das margens
        section.top_margin = Cm(config.margin_top_cm)
        section.bottom_margin = Cm(config.margin_bottom_cm)
        section.left_margin = Cm(config.margin_left_cm)
        section.right_margin = Cm(config.margin_right_cm)

        # Distância de cabeçalho e rodapé
        section.header_distance = Cm(config.header_distance_cm)
        section.footer_distance = Cm(config.footer_distance_cm)

    return {
        "sections_adjusted": sections_count,
        "page_size": config.page_size,
        "margins_cm": {
            "top": config.margin_top_cm,
            "bottom": config.margin_bottom_cm,
            "left": config.margin_left_cm,
            "right": config.margin_right_cm,
        }
    }
