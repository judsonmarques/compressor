"""
Módulo de Redução e Compressão Inteligente de Arquivos (PDF, DOCX, PPTX, XLSX e Imagens).
Focado na máxima preservação da qualidade visual, integridade absoluta dos dados
e garantia total de que nenhuma imagem ou conteúdo seja excluído ou corrompido.
"""

import os
import io
import time
import shutil
import tempfile
import zipfile
from dataclasses import dataclass, field
from typing import Dict, Any, List, Optional, Callable, Tuple
from PIL import Image

try:
    import pymupdf
    HAS_PYMUPDF = True
except ImportError:
    HAS_PYMUPDF = False

try:
    from docx import Document
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False


def format_file_size(size_bytes: int) -> str:
    """Retorna tamanho de arquivo formatado em B, KB, MB ou GB."""
    if size_bytes < 0:
        return "0 B"
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    elif size_bytes < 1024 * 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.2f} MB"
    else:
        return f"{size_bytes / (1024 * 1024 * 1024):.2f} GB"


@dataclass
class CompressionConfig:
    """Configurações para redução de tamanho de arquivos."""
    name: str = "Alta Fidelidade (Mínima Perda)"
    level_id: str = "high_quality"  # "high_quality", "balanced", "aggressive"
    
    # Parâmetros de compressão de imagens
    jpeg_quality: int = 85
    max_image_dimension: int = 2400  # Resolução máxima preservando 300 DPI A4
    png_optimize: bool = True
    png_compress_level: int = 9
    
    # Parâmetros OpenXML (DOCX, PPTX, XLSX)
    zip_compress_level: int = 9
    
    # Parâmetros PDF seguros
    pdf_garbage: int = 3          # Nível 3 é seguro (nível 4 pode remover objetos complexos)
    pdf_deflate: bool = True
    pdf_clean: bool = False       # IMPORTANTE: clean=False para NÃO modificar ou perder streams de imagem
    pdf_deflate_fonts: bool = True
    pdf_deflate_images: bool = True
    optimize_pdf_images: bool = True
    
    # Regra de segurança: nunca deixar o arquivo maior que o original
    only_if_smaller: bool = True

    @classmethod
    def get_presets(cls) -> Dict[str, "CompressionConfig"]:
        return {
            "Alta Fidelidade (Mínima Perda)": cls(
                name="Alta Fidelidade (Mínima Perda)",
                level_id="high_quality",
                jpeg_quality=85,
                max_image_dimension=2400,
                png_optimize=True,
                png_compress_level=9,
                zip_compress_level=9,
                pdf_garbage=3,
                pdf_deflate=True,
                pdf_clean=False,
                pdf_deflate_fonts=True,
                pdf_deflate_images=True,
                optimize_pdf_images=True,
                only_if_smaller=True
            ),
            "Equilibrado (Recomendado)": cls(
                name="Equilibrado (Recomendado)",
                level_id="balanced",
                jpeg_quality=78,
                max_image_dimension=1920,
                png_optimize=True,
                png_compress_level=9,
                zip_compress_level=9,
                pdf_garbage=3,
                pdf_deflate=True,
                pdf_clean=False,
                pdf_deflate_fonts=True,
                pdf_deflate_images=True,
                optimize_pdf_images=True,
                only_if_smaller=True
            ),
            "Máxima Redução": cls(
                name="Máxima Redução",
                level_id="aggressive",
                jpeg_quality=68,
                max_image_dimension=1400,
                png_optimize=True,
                png_compress_level=9,
                zip_compress_level=9,
                pdf_garbage=3,
                pdf_deflate=True,
                pdf_clean=False,
                pdf_deflate_fonts=True,
                pdf_deflate_images=True,
                optimize_pdf_images=True,
                only_if_smaller=True
            )
        }


@dataclass
class CompressionResult:
    """Resultado detalhado da redução de tamanho de um arquivo."""
    success: bool
    input_path: str
    output_path: str
    original_size: int
    compressed_size: int
    reduction_bytes: int
    reduction_percentage: float
    elapsed_time_seconds: float
    file_type: str
    images_optimized: int = 0
    details: Dict[str, Any] = field(default_factory=dict)
    error_message: Optional[str] = None

    @property
    def original_size_str(self) -> str:
        return format_file_size(self.original_size)

    @property
    def compressed_size_str(self) -> str:
        return format_file_size(self.compressed_size)

    @property
    def saved_size_str(self) -> str:
        return format_file_size(max(0, self.reduction_bytes))


class FileCompressor:
    """Compressor de arquivos com foco em preservação absoluta de imagens e dados."""

    SUPPORTED_EXTENSIONS = {
        ".pdf": "PDF",
        ".docx": "Word DOCX",
        ".pptx": "PowerPoint PPTX",
        ".xlsx": "Excel XLSX",
        ".png": "Imagem PNG",
        ".jpg": "Imagem JPEG",
        ".jpeg": "Imagem JPEG",
        ".webp": "Imagem WebP",
        ".bmp": "Imagem BMP",
        ".tiff": "Imagem TIFF",
        ".tif": "Imagem TIFF"
    }

    def __init__(self, config: Optional[CompressionConfig] = None):
        self.config = config or CompressionConfig.get_presets()["Alta Fidelidade (Mínima Perda)"]

    def set_config(self, config: CompressionConfig) -> None:
        self.config = config

    @classmethod
    def is_supported(cls, file_path: str) -> bool:
        ext = os.path.splitext(file_path)[1].lower()
        return ext in cls.SUPPORTED_EXTENSIONS

    @classmethod
    def get_supported_extensions(cls) -> List[str]:
        return list(cls.SUPPORTED_EXTENSIONS.keys())

    def compress_file(
        self,
        input_path: str,
        output_path: Optional[str] = None,
        progress_callback: Optional[Callable[[str, float], None]] = None
    ) -> CompressionResult:
        """
        Executa a compressão inteligente preservando 100% das imagens e dados.
        """
        start_time = time.time()

        if not os.path.exists(input_path):
            return CompressionResult(
                success=False,
                input_path=input_path,
                output_path="",
                original_size=0,
                compressed_size=0,
                reduction_bytes=0,
                reduction_percentage=0.0,
                elapsed_time_seconds=0.0,
                file_type="Desconhecido",
                error_message=f"Arquivo não encontrado: {input_path}"
            )

        original_size = os.path.getsize(input_path)
        ext = os.path.splitext(input_path)[1].lower()
        file_type = self.SUPPORTED_EXTENSIONS.get(ext, ext.upper())

        if not self.is_supported(input_path):
            return CompressionResult(
                success=False,
                input_path=input_path,
                output_path="",
                original_size=original_size,
                compressed_size=original_size,
                reduction_bytes=0,
                reduction_percentage=0.0,
                elapsed_time_seconds=0.0,
                file_type=file_type,
                error_message=f"Formato '{ext}' não suportado para compressão."
            )

        # Definir caminho de saída padrão se não fornecido
        if not output_path:
            dir_name, file_name = os.path.split(input_path)
            base_name, f_ext = os.path.splitext(file_name)
            output_path = os.path.join(dir_name, f"{base_name}_reduzido{f_ext}")

        out_dir = os.path.dirname(output_path)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

        temp_output = None
        try:
            # Criar arquivo temporário para processamento com validação
            with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tf:
                temp_output = tf.name

            images_opt = 0
            details: Dict[str, Any] = {}

            if ext == ".pdf":
                images_opt, details = self._compress_pdf(input_path, temp_output, progress_callback)
            elif ext in [".docx", ".pptx", ".xlsx"]:
                images_opt, details = self._compress_openxml(input_path, temp_output, progress_callback)
            elif ext in [".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tiff", ".tif"]:
                images_opt, details = self._compress_image(input_path, temp_output, progress_callback)
            else:
                raise ValueError(f"Extensão não suportada: {ext}")

            temp_size = os.path.getsize(temp_output)

            # Regra de Segurança: Só substituir se o arquivo comprimido for MENOR que o original
            if self.config.only_if_smaller and temp_size >= original_size:
                # O arquivo original já está compactado idealmente
                if os.path.abspath(input_path) != os.path.abspath(output_path):
                    shutil.copy2(input_path, output_path)
                final_size = original_size
                details["already_optimized"] = True
                details["note"] = "Arquivo já se encontrava na compressão máxima ideal. Preservado original sem alterações."
            else:
                shutil.move(temp_output, output_path)
                temp_output = None
                final_size = os.path.getsize(output_path)

            elapsed = round(time.time() - start_time, 2)
            reduction_bytes = max(0, original_size - final_size)
            reduction_pct = (reduction_bytes / original_size * 100) if original_size > 0 else 0.0

            if progress_callback:
                progress_callback("Concluído com sucesso!", 1.0)

            return CompressionResult(
                success=True,
                input_path=input_path,
                output_path=output_path,
                original_size=original_size,
                compressed_size=final_size,
                reduction_bytes=reduction_bytes,
                reduction_percentage=max(0.0, round(reduction_pct, 1)),
                elapsed_time_seconds=elapsed,
                file_type=file_type,
                images_optimized=images_opt,
                details=details
            )

        except Exception as e:
            elapsed = round(time.time() - start_time, 2)
            # Em caso de qualquer falha durante a compressão, preserva o arquivo original
            if output_path and os.path.abspath(input_path) != os.path.abspath(output_path) and not os.path.exists(output_path):
                try:
                    shutil.copy2(input_path, output_path)
                except Exception:
                    pass

            return CompressionResult(
                success=False,
                input_path=input_path,
                output_path=output_path or "",
                original_size=original_size,
                compressed_size=original_size,
                reduction_bytes=0,
                reduction_percentage=0.0,
                elapsed_time_seconds=elapsed,
                file_type=file_type,
                error_message=str(e)
            )
        finally:
            if temp_output and os.path.exists(temp_output):
                try:
                    os.remove(temp_output)
                except Exception:
                    pass

    # =========================================================================
    # COMPRESSÃO SEGURA DE PDF (GARANTIA DE 100% DAS IMAGENS PRESERVADAS)
    # =========================================================================
    def _compress_pdf(
        self,
        input_path: str,
        temp_output: str,
        progress_callback: Optional[Callable[[str, float], None]]
    ) -> Tuple[int, Dict[str, Any]]:
        """
        Otimiza PDF com PyMuPDF garantindo que NENHUMA imagem seja excluída ou corrompida.
        """
        if not HAS_PYMUPDF:
            raise RuntimeError("Biblioteca 'pymupdf' necessária para compactar PDF.")

        if progress_callback:
            progress_callback("Carregando e validando PDF de origem...", 0.1)

        # 1. Contagem rigorosa do original para validação pós-processamento
        doc_check_orig = pymupdf.open(input_path)
        orig_page_count = len(doc_check_orig)
        orig_images_total = sum(len(p.get_images()) for p in doc_check_orig)
        doc_check_orig.close()

        doc = pymupdf.open(input_path)
        images_optimized = 0
        details: Dict[str, Any] = {
            "pages": orig_page_count,
            "original_images": orig_images_total
        }

        total_pages = len(doc)
        seen_xrefs = set()

        if self.config.optimize_pdf_images:
            for page_idx, page in enumerate(doc):
                if progress_callback:
                    pct = 0.1 + (0.65 * (page_idx / max(1, total_pages)))
                    progress_callback(f"Otimizando imagens da página {page_idx + 1} de {total_pages}...", pct)

                try:
                    page_images = page.get_images()
                except Exception:
                    continue

                for img_info in page_images:
                    xref = img_info[0]
                    smask_xref = img_info[1] if len(img_info) > 1 else 0

                    if xref in seen_xrefs or xref <= 0:
                        continue
                    seen_xrefs.add(xref)

                    try:
                        base_image = doc.extract_image(xref)
                        if not base_image:
                            continue

                        image_bytes = base_image.get("image")
                        ext = base_image.get("ext", "").lower()
                        orig_w = base_image.get("width", 0)
                        orig_h = base_image.get("height", 0)

                        # Ignora imagens muito pequenas (< 25 KB) para evitar qualquer perda desnecessária
                        if not image_bytes or len(image_bytes) < 25000:
                            continue

                        # REGRA DE OURO DO PDF:
                        # 1. SÓ re-comprimimos imagens que JÁ ERAM JPEG originalmente no PDF!
                        # Nunca convertemos PNG/FlateDecode para JPEG dentro do PDF, pois o dicionário
                        # do XObject continuaria com /Filter /FlateDecode, corrompendo a imagem!
                        # 2. Se a imagem possuir máscara de transparência (smask_xref > 0),
                        # NUNCA redimensionamos as dimensões (Width/Height), pois a especificação
                        # PDF ISO 32000 exige que a imagem base e sua smask tenham a mesma dimensão!
                        if ext in ("jpeg", "jpg"):
                            has_smask = (smask_xref > 0)
                            with Image.open(io.BytesIO(image_bytes)) as pil_img:
                                needs_resize = (max(orig_w, orig_h) > self.config.max_image_dimension) and not has_smask
                                working_img = pil_img.copy()

                                if needs_resize:
                                    working_img.thumbnail(
                                        (self.config.max_image_dimension, self.config.max_image_dimension),
                                        Image.Resampling.LANCZOS
                                    )

                                out_buf = io.BytesIO()
                                if working_img.mode not in ("RGB", "L"):
                                    working_img = working_img.convert("RGB")

                                working_img.save(
                                    out_buf,
                                    format="JPEG",
                                    quality=self.config.jpeg_quality,
                                    optimize=True
                                )
                                new_bytes = out_buf.getvalue()

                                # Validação: testar se a nova imagem abre perfeitamente
                                if len(new_bytes) < len(image_bytes):
                                    with Image.open(io.BytesIO(new_bytes)) as test_v:
                                        test_v.verify()

                                    doc.update_stream(xref, new_bytes)
                                    if needs_resize:
                                        doc.xref_set_key(xref, "Width", str(working_img.width))
                                        doc.xref_set_key(xref, "Height", str(working_img.height))
                                    images_optimized += 1

                    except Exception:
                        # Em qualquer dúvida ou exceção, mantém a imagem original intacta
                        continue

        if progress_callback:
            progress_callback("Compactando fluxos de dados e fontes...", 0.85)

        # SALVAMENTO SEGURO:
        # - clean=False: essencial para NUNCA recriar fluxos de conteúdo que possam desvincular imagens
        # - garbage=3: limpeza segura sem expurgar referências indiretas complexas
        # - deflate_images=True: o MuPDF comprime perfeitamente imagens PNG/FlateDecode sem perdas!
        doc.save(
            temp_output,
            garbage=self.config.pdf_garbage,
            deflate=self.config.pdf_deflate,
            clean=False,
            deflate_images=self.config.pdf_deflate_images,
            deflate_fonts=self.config.pdf_deflate_fonts
        )
        doc.close()

        # VALIDAÇÃO PÓS-COMPRESSÃO OBRIGATÓRIA:
        # Garante que 100% das páginas e 100% das imagens continuam existindo e renderizando
        val_doc = pymupdf.open(temp_output)
        val_pages = len(val_doc)
        val_images_total = sum(len(p.get_images()) for p in val_doc)

        if val_pages != orig_page_count:
            val_doc.close()
            raise RuntimeError(f"Validação de segurança: contagem de páginas divergiu ({val_pages} vs {orig_page_count}). Revertendo.")

        if val_images_total < orig_images_total:
            val_doc.close()
            raise RuntimeError(
                f"Validação de segurança: detectada exclusão de imagem ({val_images_total} de {orig_images_total} presentes). Revertendo para o arquivo original."
            )

        # Teste de renderização rápida das páginas para garantir zero corrupção visual
        for p_idx, page in enumerate(val_doc):
            try:
                pix = page.get_pixmap(dpi=72)
                if pix.width <= 0 or pix.height <= 0:
                    val_doc.close()
                    raise RuntimeError(f"Falha de renderização na página {p_idx + 1}.")
            except Exception as e:
                val_doc.close()
                raise RuntimeError(f"Falha na renderização pós-compressão na página {p_idx + 1}: {e}")

        val_doc.close()

        details["images_processed"] = len(seen_xrefs)
        details["images_optimized"] = images_optimized
        details["final_images_count"] = val_images_total
        return images_optimized, details

    # =========================================================================
    # COMPRESSÃO SEGURA DE OPENXML (DOCX, PPTX, XLSX)
    # =========================================================================
    def _compress_openxml(
        self,
        input_path: str,
        temp_output: str,
        progress_callback: Optional[Callable[[str, float], None]]
    ) -> Tuple[int, Dict[str, Any]]:
        """
        Otimiza documentos OpenXML (.docx, .pptx, .xlsx).
        Garante que NENHUM arquivo de mídia seja excluído e preserva extensões e tipos estritamente.
        """
        if progress_callback:
            progress_callback("Lendo pacote do documento...", 0.1)

        images_optimized = 0
        total_media_images = 0

        # Formatos raster permitidos para recompressão segura
        safe_raster_exts = {".jpg", ".jpeg", ".png"}

        with zipfile.ZipFile(input_path, "r") as zin:
            infolist = zin.infolist()
            total_items = len(infolist)

            # Mapear todas as imagens originais para validação estrita
            orig_media_entries = {
                item.filename: len(zin.read(item.filename))
                for item in infolist
                if "media/" in item.filename.lower()
            }
            total_media_images = len(orig_media_entries)

            with zipfile.ZipFile(
                temp_output,
                "w",
                compression=zipfile.ZIP_DEFLATED,
                compresslevel=self.config.zip_compress_level
            ) as zout:
                for idx, item in enumerate(infolist):
                    if progress_callback and total_items > 0:
                        pct = 0.1 + (0.75 * (idx / total_items))
                        progress_callback(f"Otimizando elementos ({idx + 1}/{total_items})...", pct)

                    data = zin.read(item.filename)
                    f_ext = os.path.splitext(item.filename)[1].lower()

                    # Verificar se é mídia raster elegível
                    is_media = "media/" in item.filename.lower()
                    is_safe_raster = is_media and (f_ext in safe_raster_exts)

                    # Se for vetor (.emf, .wmf, .svg) ou formato especial (.bin, .gif),
                    # NUNCA ALTERAR. Mantém os bytes originais 100% intocados!
                    if is_safe_raster and len(data) > 25000:
                        optimized_bytes = self._optimize_openxml_image(data, f_ext)
                        if optimized_bytes is not None and len(optimized_bytes) < len(data):
                            data = optimized_bytes
                            images_optimized += 1

                    # Escreve a entrada de forma limpa preservando exatamente o nome e data
                    new_info = zipfile.ZipInfo(filename=item.filename, date_time=item.date_time)
                    new_info.compress_type = zipfile.ZIP_DEFLATED
                    new_info.external_attr = item.external_attr
                    zout.writestr(new_info, data)

        # VALIDAÇÃO CRÍTICA DO OPENXML:
        # 1. Checar se todas as entradas de mídia continuam presentes no ZIP de saída
        with zipfile.ZipFile(temp_output, "r") as zcheck:
            comp_media_entries = [f for f in zcheck.namelist() if "media/" in f.lower()]
            if set(comp_media_entries) != set(orig_media_entries.keys()):
                raise RuntimeError(
                    f"Falha de integridade: o documento resultante perdeu arquivos de imagem ({len(comp_media_entries)} vs {total_media_images}). Revertendo."
                )

            # 2. Validar que cada arquivo de imagem no novo ZIP abre sem erros
            for f in comp_media_entries:
                ext = os.path.splitext(f)[1].lower()
                if ext in safe_raster_exts:
                    img_data = zcheck.read(f)
                    with Image.open(io.BytesIO(img_data)) as v_img:
                        v_img.verify()

            # 3. Teste do arquivo ZIP
            corrupt = zcheck.testzip()
            if corrupt is not None:
                raise RuntimeError(f"Pacote ZIP corrompido no arquivo interno: {corrupt}. Revertendo.")

        # 4. Se for DOCX, validar abertura estrutural no python-docx
        if HAS_DOCX and input_path.lower().endswith(".docx"):
            try:
                doc_test = Document(temp_output)
                # Verifica se existem parágrafos
                _ = len(doc_test.paragraphs)
            except Exception as e:
                raise RuntimeError(f"O documento DOCX gerado falhou na validação estrutural do Word: {e}")

        details = {
            "total_media_images": total_media_images,
            "images_optimized": images_optimized,
            "all_images_preserved": True
        }
        return images_optimized, details

    def _optimize_openxml_image(self, data: bytes, f_ext: str) -> Optional[bytes]:
        """
        Otimiza dados de uma imagem mantendo estritamente seu formato original.
        Se for .png, salva SEMPRE como .png. Se for .jpg/.jpeg, salva SEMPRE como .jpeg.
        """
        try:
            with Image.open(io.BytesIO(data)) as img:
                orig_w, orig_h = img.size
                if orig_w <= 0 or orig_h <= 0:
                    return None

                needs_resize = max(orig_w, orig_h) > self.config.max_image_dimension
                working_img = img.copy()

                if needs_resize:
                    working_img.thumbnail(
                        (self.config.max_image_dimension, self.config.max_image_dimension),
                        Image.Resampling.LANCZOS
                    )

                out_buf = io.BytesIO()

                if f_ext in (".jpg", ".jpeg"):
                    if working_img.mode not in ("RGB", "L"):
                        working_img = working_img.convert("RGB")
                    working_img.save(
                        out_buf,
                        format="JPEG",
                        quality=self.config.jpeg_quality,
                        optimize=True,
                        progressive=True
                    )
                elif f_ext == ".png":
                    # IMPORTANTE: Preservar formato PNG estritamente para não quebrar o MIME type do Word
                    working_img.save(
                        out_buf,
                        format="PNG",
                        optimize=self.config.png_optimize,
                        compress_level=self.config.png_compress_level
                    )
                else:
                    return None

                new_bytes = out_buf.getvalue()

                # Só retorna se for menor E abrir perfeitamente no Pillow
                if len(new_bytes) < len(data):
                    with Image.open(io.BytesIO(new_bytes)) as test_v:
                        test_v.verify()
                    return new_bytes

        except Exception:
            return None

        return None

    # =========================================================================
    # COMPRESSÃO SEGURA DE IMAGENS AVULSAS
    # =========================================================================
    def _compress_image(
        self,
        input_path: str,
        temp_output: str,
        progress_callback: Optional[Callable[[str, float], None]]
    ) -> Tuple[int, Dict[str, Any]]:
        """Otimiza imagem avulsa garantindo integridade visual."""
        if progress_callback:
            progress_callback("Otimizando imagem...", 0.3)

        with Image.open(input_path) as img:
            orig_w, orig_h = img.size
            ext = os.path.splitext(input_path)[1].lower()

            needs_resize = max(orig_w, orig_h) > self.config.max_image_dimension
            working_img = img.copy()

            if needs_resize:
                working_img.thumbnail(
                    (self.config.max_image_dimension, self.config.max_image_dimension),
                    Image.Resampling.LANCZOS
                )

            if ext in [".jpg", ".jpeg"]:
                if working_img.mode not in ("RGB", "L"):
                    working_img = working_img.convert("RGB")
                working_img.save(
                    temp_output,
                    format="JPEG",
                    quality=self.config.jpeg_quality,
                    optimize=True,
                    progressive=True
                )
            elif ext == ".png":
                working_img.save(
                    temp_output,
                    format="PNG",
                    optimize=self.config.png_optimize,
                    compress_level=self.config.png_compress_level
                )
            elif ext == ".webp":
                working_img.save(
                    temp_output,
                    format="WEBP",
                    quality=self.config.jpeg_quality,
                    method=6
                )
            else:
                working_img.save(temp_output, optimize=True)

        # Validar integridade do arquivo gerado
        with Image.open(temp_output) as check_img:
            check_img.verify()

        details = {
            "original_dimensions": f"{orig_w}x{orig_h}",
            "final_dimensions": f"{working_img.size[0]}x{working_img.size[1]}"
        }
        return 1, details

    def compress_directory(
        self,
        dir_path: str,
        output_dir: Optional[str] = None,
        extensions: Optional[List[str]] = None,
        recursive: bool = False,
        file_callback: Optional[Callable[[CompressionResult, int, int], None]] = None
    ) -> List[CompressionResult]:
        """Processa todos os arquivos suportados em um diretório com segurança total."""
        allowed_exts = extensions or self.get_supported_extensions()
        allowed_exts = [e.lower() for e in allowed_exts]

        files_to_process = []
        if recursive:
            for root, _, files in os.walk(dir_path):
                for f in files:
                    ext = os.path.splitext(f)[1].lower()
                    if ext in allowed_exts and not f.startswith("~$") and "_reduzido" not in f:
                        files_to_process.append(os.path.join(root, f))
        else:
            for f in os.listdir(dir_path):
                full_p = os.path.join(dir_path, f)
                if os.path.isfile(full_p):
                    ext = os.path.splitext(f)[1].lower()
                    if ext in allowed_exts and not f.startswith("~$") and "_reduzido" not in f:
                        files_to_process.append(full_p)

        results = []
        total = len(files_to_process)

        for idx, file_path in enumerate(files_to_process, start=1):
            out_file = None
            if output_dir:
                rel_path = os.path.relpath(file_path, dir_path)
                out_file = os.path.join(output_dir, rel_path)

            res = self.compress_file(file_path, output_path=out_file)
            results.append(res)

            if file_callback:
                file_callback(res, idx, total)

        return results
