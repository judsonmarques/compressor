"""
Módulo para alinhamento, formatação e ajuste de paginação de tabelas no Word (.docx).
Impede corte indesejado de linhas entre páginas e centraliza tabelas no documento.
"""

from typing import Dict, Any
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from core.config import FormattingConfig


def adjust_tables(doc: Document, config: FormattingConfig) -> Dict[str, Any]:
    """
    Padroniza tabelas:
    - Centraliza a tabela no documento.
    - Aplica w:cantSplit em todas as linhas para evitar que linhas se partam no meio entre duas páginas.
    - Adiciona w:tblHeader na primeira linha para repetir o cabeçalho se a tabela continuar na página seguinte.
    """
    tables_count = len(doc.tables)
    rows_adjusted = 0
    headers_repeated = 0

    cant_split_xml = f'<w:cantSplit {nsdecls("w")}/>'
    tbl_header_xml = f'<w:tblHeader {nsdecls("w")}/>'

    for table in doc.tables:
        # Centralização da tabela
        if config.center_tables:
            table.alignment = WD_TABLE_ALIGNMENT.CENTER

        rows = table.rows
        num_rows = len(rows)

        for i, row in enumerate(rows):
            trPr = row._tr.get_or_add_trPr()

            # 1. Impedir corte da linha entre páginas
            if config.table_cant_split_rows:
                # Verificar se já possui cantSplit para não duplicar
                if not trPr.xpath('./w:cantSplit'):
                    trPr.append(parse_xml(cant_split_xml))
                    rows_adjusted += 1

            # 2. Repetir linha de cabeçalho na página seguinte
            if config.table_repeat_header and i == 0 and num_rows > 1:
                if not trPr.xpath('./w:tblHeader'):
                    trPr.append(parse_xml(tbl_header_xml))
                    headers_repeated += 1

    return {
        "tables_adjusted": tables_count,
        "rows_prevented_from_splitting": rows_adjusted,
        "headers_repeated": headers_repeated,
    }
