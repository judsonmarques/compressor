"""
Módulo de Configuração de Formatação e Perfis para Documentos Word (.docx).
"""

from dataclasses import dataclass, field, asdict
import json
from typing import Dict, Any


@dataclass
class FormattingConfig:
    """Configurações completas de alinhamento, formatação e limpeza."""

    # Nome do Perfil
    name: str = "Corporativo Moderno"

    # Layout de Página
    page_size: str = "A4"  # "A4" ou "Letter"
    margin_top_cm: float = 2.5
    margin_bottom_cm: float = 2.5
    margin_left_cm: float = 2.5
    margin_right_cm: float = 2.5
    header_distance_cm: float = 1.25
    footer_distance_cm: float = 1.25

    # Tipografia e Parágrafos
    change_font: bool = True
    font_name: str = "Calibri"  # Ex: "Calibri", "Arial", "Times New Roman"
    font_size_pt: float = 11.0  # Tamanho do corpo de texto
    heading_font_name: str = "Calibri"
    heading1_size_pt: float = 16.0
    heading2_size_pt: float = 13.0
    heading3_size_pt: float = 12.0

    body_alignment: str = "JUSTIFY"  # "JUSTIFY", "LEFT", "CENTER", "RIGHT"
    heading_alignment: str = "LEFT"  # "LEFT", "CENTER", "JUSTIFY"
    line_spacing: float = 1.15  # Multiplicador (ex: 1.15, 1.5)
    space_before_pt: float = 0.0
    space_after_pt: float = 6.0
    first_line_indent_cm: float = 0.0  # Recuo da 1ª linha (ex: 1.25 cm na ABNT)

    # Controle de Quebras e Paginação
    keep_headings_with_next: bool = True  # Impede títulos órfãos no final da página
    widow_control: bool = True  # Impede linhas isoladas no início/fim de página

    # Imagens
    center_images: bool = True
    fit_images_to_page: bool = True  # Reduz imagens que ultrapassam as margens
    image_space_before_pt: float = 8.0
    image_space_after_pt: float = 8.0

    # Tabelas
    center_tables: bool = True
    table_cant_split_rows: bool = True  # Impede que linhas de tabela se dividam entre páginas
    table_repeat_header: bool = True  # Repete cabeçalho da tabela se ela ocupar várias páginas
    table_cell_padding_pt: float = 4.0

    # Limpeza e Higienização
    remove_empty_paragraphs: bool = True
    max_consecutive_empty_lines: int = 0  # 0 = remove todas as linhas vazias redundantes
    remove_redundant_spaces: bool = True  # Substitui espaços duplos consecutivos por um único
    remove_trailing_empty_paragraphs: bool = True  # Remove parágrafos vazios ao final do arquivo

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "FormattingConfig":
        valid_keys = {f.name for f in cls.__dataclass_fields__.values()}
        filtered = {k: v for k, v in data.items() if k in valid_keys}
        return cls(**filtered)

    def save_to_json(self, file_path: str) -> None:
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, indent=4, ensure_ascii=False)

    @classmethod
    def load_from_json(cls, file_path: str) -> "FormattingConfig":
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return cls.from_dict(data)


# Perfis pré-configurados
PRESETS: Dict[str, FormattingConfig] = {
    "Corporativo Moderno": FormattingConfig(
        name="Corporativo Moderno",
        page_size="A4",
        margin_top_cm=2.5,
        margin_bottom_cm=2.5,
        margin_left_cm=2.5,
        margin_right_cm=2.5,
        change_font=True,
        font_name="Calibri",
        font_size_pt=11.0,
        heading_font_name="Calibri",
        heading1_size_pt=16.0,
        heading2_size_pt=13.0,
        heading3_size_pt=12.0,
        body_alignment="JUSTIFY",
        heading_alignment="LEFT",
        line_spacing=1.15,
        space_before_pt=0.0,
        space_after_pt=6.0,
        first_line_indent_cm=0.0,
        keep_headings_with_next=True,
        widow_control=True,
        center_images=True,
        fit_images_to_page=True,
        center_tables=True,
        table_cant_split_rows=True,
        table_repeat_header=True,
        remove_empty_paragraphs=True,
        max_consecutive_empty_lines=0,
        remove_redundant_spaces=True,
    ),
    "Padrão ABNT": FormattingConfig(
        name="Padrão ABNT",
        page_size="A4",
        margin_top_cm=3.0,
        margin_bottom_cm=2.0,
        margin_left_cm=3.0,
        margin_right_cm=2.0,
        change_font=True,
        font_name="Arial",
        font_size_pt=12.0,
        heading_font_name="Arial",
        heading1_size_pt=14.0,
        heading2_size_pt=12.0,
        heading3_size_pt=12.0,
        body_alignment="JUSTIFY",
        heading_alignment="LEFT",
        line_spacing=1.5,
        space_before_pt=0.0,
        space_after_pt=0.0,
        first_line_indent_cm=1.25,
        keep_headings_with_next=True,
        widow_control=True,
        center_images=True,
        fit_images_to_page=True,
        center_tables=True,
        table_cant_split_rows=True,
        table_repeat_header=True,
        remove_empty_paragraphs=True,
        max_consecutive_empty_lines=0,
        remove_redundant_spaces=True,
    ),
    "Executivo Compacto": FormattingConfig(
        name="Executivo Compacto",
        page_size="A4",
        margin_top_cm=2.0,
        margin_bottom_cm=2.0,
        margin_left_cm=2.0,
        margin_right_cm=2.0,
        change_font=True,
        font_name="Calibri",
        font_size_pt=10.5,
        heading_font_name="Calibri",
        heading1_size_pt=14.0,
        heading2_size_pt=12.0,
        heading3_size_pt=11.0,
        body_alignment="JUSTIFY",
        heading_alignment="LEFT",
        line_spacing=1.1,
        space_before_pt=0.0,
        space_after_pt=4.0,
        first_line_indent_cm=0.0,
        keep_headings_with_next=True,
        widow_control=True,
        center_images=True,
        fit_images_to_page=True,
        center_tables=True,
        table_cant_split_rows=True,
        table_repeat_header=True,
        remove_empty_paragraphs=True,
        max_consecutive_empty_lines=0,
        remove_redundant_spaces=True,
    ),
    "Preservar Fontes (Apenas Layout e Limpeza)": FormattingConfig(
        name="Preservar Fontes (Apenas Layout e Limpeza)",
        page_size="A4",
        margin_top_cm=2.5,
        margin_bottom_cm=2.5,
        margin_left_cm=2.5,
        margin_right_cm=2.5,
        change_font=False,
        body_alignment="JUSTIFY",
        heading_alignment="LEFT",
        line_spacing=1.15,
        space_before_pt=0.0,
        space_after_pt=6.0,
        first_line_indent_cm=0.0,
        keep_headings_with_next=True,
        widow_control=True,
        center_images=True,
        fit_images_to_page=True,
        center_tables=True,
        table_cant_split_rows=True,
        table_repeat_header=True,
        remove_empty_paragraphs=True,
        max_consecutive_empty_lines=0,
        remove_redundant_spaces=True,
    ),
}
