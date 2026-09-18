"""
Módulo Principal de Processamento e Otimização de Documentos Word (.docx).
Coordena todas as etapas de limpeza, alinhamento, formatação de imagens e layout.
"""

import os
import shutil
import time
from dataclasses import dataclass, field
from typing import Dict, Any, List, Optional, Callable
from docx import Document

from core.config import FormattingConfig, PRESETS
from core.page_setup import setup_pages
from core.cleaner import clean_document_whitespace
from core.image_handler import adjust_and_align_images
from core.table_handler import adjust_tables
from core.aligner import format_text_and_headings


@dataclass
class ProcessResult:
    """Resultado detalhado do processamento de um documento."""
    success: bool
    input_path: str
    output_path: str
    elapsed_time_seconds: float
    error_message: Optional[str] = None
    stats: Dict[str, Any] = field(default_factory=dict)


class DocxOptimizer:
    """Otimizador e padronizador de documentos Word."""

    def __init__(self, config: Optional[FormattingConfig] = None):
        self.config = config or PRESETS["Corporativo Moderno"]

    def set_config(self, config: FormattingConfig) -> None:
        self.config = config

    def process_file(
        self,
        input_path: str,
        output_path: Optional[str] = None,
        progress_callback: Optional[Callable[[str, float], None]] = None,
    ) -> ProcessResult:
        """
        Executa a otimização completa em um único arquivo .docx.
        """
        start_time = time.time()
        
        if not os.path.exists(input_path):
            return ProcessResult(
                success=False,
                input_path=input_path,
                output_path="",
                elapsed_time_seconds=0,
                error_message=f"Arquivo não encontrado: {input_path}"
            )

        if not input_path.lower().endswith(".docx"):
            return ProcessResult(
                success=False,
                input_path=input_path,
                output_path="",
                elapsed_time_seconds=0,
                error_message="O arquivo deve ser um documento no formato Word (.docx)"
            )

        # Definir caminho de saída padrão se não fornecido
        if not output_path:
            dir_name, file_name = os.path.split(input_path)
            base_name, ext = os.path.splitext(file_name)
            output_path = os.path.join(dir_name, f"{base_name}_corrigido{ext}")

        try:
            if progress_callback:
                progress_callback("Carregando documento...", 0.1)

            doc = Document(input_path)
            stats: Dict[str, Any] = {}

            # Etapa 1: Configuração de Página e Margens
            if progress_callback:
                progress_callback("Ajustando tamanho de página e margens...", 0.25)
            stats["page"] = setup_pages(doc, self.config)

            # Etapa 2: Alinhamento e Dimensionamento de Imagens
            if progress_callback:
                progress_callback("Centralizando e ajustando imagens às margens...", 0.45)
            stats["images"] = adjust_and_align_images(doc, self.config)

            # Etapa 3: Formatação e Proteção de Linhas de Tabelas
            if progress_callback:
                progress_callback("Ajustando tabelas e prevenindo corte entre páginas...", 0.60)
            stats["tables"] = adjust_tables(doc, self.config)

            # Etapa 4: Eliminação de Linhas em Branco e Espaços Redundantes
            if progress_callback:
                progress_callback("Eliminando parágrafos vazios e espaços excessivos...", 0.75)
            stats["cleaner"] = clean_document_whitespace(doc, self.config)

            # Etapa 5: Alinhamento de Texto, Títulos e Paginação
            if progress_callback:
                progress_callback("Alinhando textos, títulos e entrelinhas...", 0.90)
            stats["text"] = format_text_and_headings(doc, self.config)

            # Salvar documento final
            if progress_callback:
                progress_callback("Salvando documento padronizado...", 0.95)
            
            output_dir = os.path.dirname(output_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)

            doc.save(output_path)

            elapsed = round(time.time() - start_time, 2)

            if progress_callback:
                progress_callback("Concluído com sucesso!", 1.0)

            return ProcessResult(
                success=True,
                input_path=input_path,
                output_path=output_path,
                elapsed_time_seconds=elapsed,
                stats=stats
            )

        except Exception as e:
            elapsed = round(time.time() - start_time, 2)
            return ProcessResult(
                success=False,
                input_path=input_path,
                output_path=output_path or "",
                elapsed_time_seconds=elapsed,
                error_message=str(e)
            )

    def process_directory(
        self,
        directory_path: str,
        output_dir: Optional[str] = None,
        file_callback: Optional[Callable[[ProcessResult, int, int], None]] = None,
    ) -> List[ProcessResult]:
        """
        Processa todos os arquivos .docx de uma pasta em lote.
        """
        if not os.path.isdir(directory_path):
            return []

        docx_files = [
            os.path.join(directory_path, f)
            for f in os.listdir(directory_path)
            if f.lower().endswith(".docx") and not f.startswith("~$") and not "_corrigido" in f
        ]

        total = len(docx_files)
        results = []

        for idx, file_path in enumerate(docx_files, start=1):
            target_out = None
            if output_dir:
                os.makedirs(output_dir, exist_ok=True)
                file_name = os.path.basename(file_path)
                base_name, ext = os.path.splitext(file_name)
                target_out = os.path.join(output_dir, f"{base_name}_corrigido{ext}")

            res = self.process_file(file_path, output_path=target_out)
            results.append(res)

            if file_callback:
                file_callback(res, idx, total)

        return results
