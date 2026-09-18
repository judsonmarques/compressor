"""
Pacote core para otimização de documentos Word (.docx) e redução inteligente de tamanho de arquivos.
"""

from core.config import FormattingConfig, PRESETS
from core.processor import DocxOptimizer, ProcessResult
from core.compressor import FileCompressor, CompressionConfig, CompressionResult, format_file_size

__all__ = [
    "FormattingConfig",
    "PRESETS",
    "DocxOptimizer",
    "ProcessResult",
    "FileCompressor",
    "CompressionConfig",
    "CompressionResult",
    "format_file_size"
]
