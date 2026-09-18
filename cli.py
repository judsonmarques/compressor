"""
Interface de Linha de Comando (CLI) para DocxFormatter Pro.
Permite automação de correções de formatação e redução inteligente de tamanho de arquivos
(PDF, DOCX, PPTX, XLSX, Imagens) com saídas visuais ricas via terminal.
"""

import sys
import os
import argparse

# Assegurar suporte a caracteres no console do Windows
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except AttributeError:
        pass

from rich.console import Console
from rich.panel import Panel
from rich.table import Table
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn, TimeElapsedColumn

from core.config import FormattingConfig, PRESETS
from core.processor import DocxOptimizer, ProcessResult
from core.compressor import FileCompressor, CompressionConfig, CompressionResult, format_file_size

console = Console(highlight=False)


def print_banner():
    banner_text = (
        "[bold cyan]DocxFormatter Pro & FileCompressor[/bold cyan] - [dim]Otimização de Documentos e Redução de Tamanho[/dim]\n"
        "[green]✓ Padronização e Limpeza DOCX[/green] | "
        "[green]✓ Redução de Arquivos (PDF, DOCX, Imagens)[/green] | "
        "[green]✓ Mínima Perda e Alta Fidelidade[/green]"
    )
    console.print(Panel(banner_text, border_style="cyan"))


def display_result_table(result: ProcessResult):
    """Exibe tabela detalhada com o resumo das alterações de formatação efetuadas."""
    if not result.success:
        console.print(f"[bold red]❌ Erro ao processar '{os.path.basename(result.input_path)}':[/bold red] {result.error_message}")
        return

    table = Table(title=f"Resultados de Formatação: {os.path.basename(result.input_path)}", border_style="green")
    table.add_column("Elemento / Ajuste", style="cyan", no_wrap=True)
    table.add_column("Modificação Efetuada", style="white")
    table.add_column("Métrica", style="bold yellow")

    stats = result.stats
    
    # Linhas vazias e espaços
    cleaner = stats.get("cleaner", {})
    empty_removed = cleaner.get("empty_paragraphs_removed", 0)
    spaces_cleaned = cleaner.get("runs_spaces_cleaned", 0)
    table.add_row("Parágrafos Vazios Eliminados", "Removidos para evitar saltos de página", f"{empty_removed} parágrafos")
    table.add_row("Espaços Duplos Normalizados", "Substituídos por espaçamento simples", f"{spaces_cleaned} trechos")

    # Imagens
    images = stats.get("images", {})
    img_found = images.get("images_found", 0)
    img_centered = images.get("images_centered", 0)
    img_resized = images.get("images_resized", 0)
    table.add_row("Imagens Centralizadas", "Alinhadas ao centro da página", f"{img_centered} de {img_found}")
    table.add_row("Imagens Redimensionadas", "Ajustadas à largura útil da margem", f"{img_resized} imagens")

    # Tabelas
    tables_stat = stats.get("tables", {})
    tbl_adj = tables_stat.get("tables_adjusted", 0)
    cant_split = tables_stat.get("rows_prevented_from_splitting", 0)
    headers_rep = tables_stat.get("headers_repeated", 0)
    table.add_row("Tabelas Centralizadas", "Centralizadas na folha A4", f"{tbl_adj} tabelas")
    table.add_row("Linhas contra corte de página", "Ativado cantSplit em linhas", f"{cant_split} linhas")
    if headers_rep > 0:
        table.add_row("Cabeçalhos de Tabela Repetidos", "Repetição em multi-páginas", f"{headers_rep} tabelas")

    # Textos
    text_stat = stats.get("text", {})
    body_aligned = text_stat.get("body_paragraphs_aligned", 0)
    headings_adj = text_stat.get("headings_adjusted", 0)
    font_app = text_stat.get("font_applied", "Original")
    align_app = text_stat.get("alignment_applied", "JUSTIFY")
    table.add_row("Parágrafos de Texto Formatados", f"Alinhamento {align_app}, fonte {font_app}", f"{body_aligned} parágrafos")
    table.add_row("Títulos e Cabeçalhos", "Configurado keep_with_next contra orfandade", f"{headings_adj} títulos")

    # Página e Margens
    page_stat = stats.get("page", {})
    margins = page_stat.get("margins_cm", {})
    margins_str = f"Sup: {margins.get('top')}cm, Inf: {margins.get('bottom')}cm, Esq: {margins.get('left')}cm, Dir: {margins.get('right')}cm"
    table.add_row("Página e Margens", f"Padrão {page_stat.get('page_size', 'A4')}", margins_str)

    console.print(table)
    console.print(f"[bold green]✓ Salvo em:[/bold green] [underline]{result.output_path}[/underline] [dim]({result.elapsed_time_seconds}s)[/dim]\n")


def display_compression_table(result: CompressionResult):
    """Exibe tabela rica com o resumo da redução de tamanho do arquivo."""
    if not result.success:
        console.print(f"[bold red]❌ Erro ao comprimir '{os.path.basename(result.input_path)}':[/bold red] {result.error_message}")
        return

    table = Table(title=f"Resultado de Compressão: {os.path.basename(result.input_path)}", border_style="magenta")
    table.add_column("Métrica", style="cyan", no_wrap=True)
    table.add_column("Valor / Descrição", style="white")

    table.add_row("Tipo de Arquivo", result.file_type)
    table.add_row("Tamanho Original", f"[yellow]{result.original_size_str}[/yellow]")
    table.add_row("Tamanho Otimizado", f"[bold green]{result.compressed_size_str}[/bold green]")
    
    reduction_color = "bold green" if result.reduction_percentage > 0 else "dim yellow"
    table.add_row(
        "Redução de Tamanho",
        f"[{reduction_color}]{result.reduction_percentage}% ({result.saved_size_str} economizados)[/{reduction_color}]"
    )

    if result.images_optimized > 0:
        table.add_row("Imagens Otimizadas", f"{result.images_optimized} imagem(ns)")

    note = result.details.get("note")
    if note:
        table.add_row("Status", f"[cyan]{note}[/cyan]")
    else:
        table.add_row("Status", "[green]Otimizado com sucesso preservando fidelidade visual[/green]")

    table.add_row("Tempo de Processamento", f"{result.elapsed_time_seconds} segundos")

    console.print(table)
    console.print(f"[bold green]✓ Salvo em:[/bold green] [underline]{result.output_path}[/underline]\n")


def run_cli():
    parser = argparse.ArgumentParser(
        description="DocxFormatter Pro - Otimização de documentos Word e redução inteligente de arquivos (PDF, DOCX, PPTX, XLSX, Imagens)."
    )
    
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("-f", "--file", type=str, help="Caminho para um arquivo específico (.docx, .pdf, .pptx, .xlsx, imagens)")
    group.add_argument("-d", "--folder", type=str, help="Caminho para uma pasta para processar em lote")

    parser.add_argument("-o", "--output", type=str, default=None, help="Caminho do arquivo ou pasta de destino")
    
    # Modo de Operação: Redução de Tamanho
    parser.add_argument(
        "-c", "--compress",
        action="store_true",
        help="Ativa o modo de redução de tamanho de arquivo (suporta PDF, DOCX, PPTX, XLSX e Imagens)"
    )
    parser.add_argument(
        "--level",
        type=str,
        choices=["alta", "equilibrado", "max"],
        default="alta",
        help="Nível de compressão: 'alta' (mínima perda / alta fidelidade), 'equilibrado' (recomendado) ou 'max' (máxima redução)"
    )

    # Opções de Formatação DOCX
    parser.add_argument(
        "-p", "--preset",
        type=str,
        choices=["corporativo", "abnt", "compacto", "preservar"],
        default="corporativo",
        help="Perfil de formatação pré-definido para DOCX (padrão: corporativo)"
    )
    parser.add_argument("--font", type=str, default=None, help="Sobrescrever família da fonte (ex: Arial, Calibri, Times New Roman)")
    parser.add_argument("--font-size", type=float, default=None, help="Sobrescrever tamanho da fonte do corpo em pontos (ex: 11, 12)")
    parser.add_argument("--spacing", type=float, default=None, help="Sobrescrever entrelinhas (ex: 1.15, 1.5)")
    parser.add_argument("--keep-empty", action="store_true", help="NÃO remover linhas em branco vazias")

    args = parser.parse_args()
    print_banner()

    # Identificar se a intenção é Redução de Tamanho (Compressão)
    is_compress_mode = args.compress
    if not is_compress_mode and args.file:
        ext = os.path.splitext(args.file)[1].lower()
        if ext in [".pdf", ".pptx", ".xlsx", ".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tiff", ".tif"]:
            is_compress_mode = True

    # ==========================================
    # FLUXO 1: REDUÇÃO DE TAMANHO (COMPRESSÃO)
    # ==========================================
    if is_compress_mode:
        level_map = {
            "alta": "Alta Fidelidade (Mínima Perda)",
            "equilibrado": "Equilibrado (Recomendado)",
            "max": "Máxima Redução"
        }
        chosen_level_name = level_map.get(args.level.lower(), "Alta Fidelidade (Mínima Perda)")
        compress_config = CompressionConfig.get_presets()[chosen_level_name]
        compressor = FileCompressor(compress_config)

        console.print(f"[bold magenta]Modo:[/bold magenta] Redução de Tamanho de Arquivo")
        console.print(f"[bold cyan]Nível Selecionado:[/bold cyan] {compress_config.name}")
        console.print(f"[dim]Qualidade JPEG: {compress_config.jpeg_quality}% | Dimensão Máx: {compress_config.max_image_dimension}px | Deflate Nível 9[/dim]\n")

        if args.file:
            file_path = os.path.abspath(args.file)
            if not os.path.exists(file_path):
                console.print(f"[bold red]Erro:[/bold red] Arquivo não encontrado: {file_path}")
                sys.exit(1)

            with Progress(
                SpinnerColumn(),
                TextColumn("[progress.description]{task.description}"),
                BarColumn(),
                TimeElapsedColumn(),
                console=console
            ) as progress:
                task = progress.add_task(f"Comprimindo {os.path.basename(file_path)}...", total=1.0)

                def prog_cb(msg, val):
                    progress.update(task, description=f"{msg}", completed=val)

                res = compressor.compress_file(file_path, output_path=args.output, progress_callback=prog_cb)

            display_compression_table(res)

        elif args.folder:
            folder_path = os.path.abspath(args.folder)
            if not os.path.isdir(folder_path):
                console.print(f"[bold red]Erro:[/bold red] Diretório não encontrado: {folder_path}")
                sys.exit(1)

            supported_exts = compressor.get_supported_extensions()
            files = [
                os.path.join(folder_path, f)
                for f in os.listdir(folder_path)
                if os.path.isfile(os.path.join(folder_path, f))
                and os.path.splitext(f)[1].lower() in supported_exts
                and not f.startswith("~$")
                and "_reduzido" not in f
            ]

            if not files:
                console.print(f"[yellow]Nenhum arquivo elegível para compressão encontrado em:[/yellow] {folder_path}")
                sys.exit(0)

            console.print(f"[bold green]Encontrados {len(files)} arquivo(s) elegíveis para redução de tamanho.[/bold green]\n")

            with Progress(
                SpinnerColumn(),
                TextColumn("[progress.description]{task.description}"),
                BarColumn(),
                TimeElapsedColumn(),
                console=console
            ) as progress:
                batch_task = progress.add_task("Processando lote de compressão...", total=len(files))

                def file_cb(res: CompressionResult, current: int, total: int):
                    progress.update(batch_task, description=f"[{current}/{total}] {os.path.basename(res.input_path)}", completed=current)

                results = compressor.compress_directory(folder_path, output_dir=args.output, file_callback=file_cb)

            for res in results:
                display_compression_table(res)

        return

    # ==========================================
    # FLUXO 2: FORMATAÇÃO E PADRONIZAÇÃO DOCX
    # ==========================================
    preset_map = {
        "corporativo": "Corporativo Moderno",
        "abnt": "Padrão ABNT",
        "compacto": "Executivo Compacto",
        "preservar": "Preservar Fontes (Apenas Layout e Limpeza)"
    }
    selected_preset_name = preset_map.get(args.preset.lower(), "Corporativo Moderno")
    config = PRESETS[selected_preset_name]

    if args.font:
        config.change_font = True
        config.font_name = args.font
    if args.font_size:
        config.font_size_pt = args.font_size
    if args.spacing:
        config.line_spacing = args.spacing
    if args.keep_empty:
        config.remove_empty_paragraphs = False

    console.print(f"[bold cyan]Perfil de Formatação:[/bold cyan] {config.name}")
    console.print(f"[dim]Fonte: {config.font_name if config.change_font else 'Original'} | Entrelinhas: {config.line_spacing} | Margens: {config.margin_top_cm}cm[/dim]\n")

    optimizer = DocxOptimizer(config)

    if args.file:
        file_path = os.path.abspath(args.file)
        if not os.path.exists(file_path):
            console.print(f"[bold red]Erro:[/bold red] Arquivo não encontrado: {file_path}")
            sys.exit(1)

        with Progress(
            SpinnerColumn(),
            TextColumn("[progress.description]{task.description}"),
            BarColumn(),
            TimeElapsedColumn(),
            console=console
        ) as progress:
            task = progress.add_task(f"Formatando {os.path.basename(file_path)}...", total=1.0)

            def progress_cb(msg, val):
                progress.update(task, description=f"{msg}", completed=val)

            result = optimizer.process_file(file_path, output_path=args.output, progress_callback=progress_cb)

        display_result_table(result)

    elif args.folder:
        folder_path = os.path.abspath(args.folder)
        if not os.path.isdir(folder_path):
            console.print(f"[bold red]Erro:[/bold red] Diretório não encontrado: {folder_path}")
            sys.exit(1)

        files = [
            f for f in os.listdir(folder_path)
            if f.lower().endswith(".docx") and not f.startswith("~$") and not "_corrigido" in f
        ]

        if not files:
            console.print(f"[yellow]Nenhum arquivo .docx encontrado em:[/yellow] {folder_path}")
            sys.exit(0)

        console.print(f"[bold green]Encontrados {len(files)} arquivo(s) para processar.[/bold green]\n")

        with Progress(
            SpinnerColumn(),
            TextColumn("[progress.description]{task.description}"),
            BarColumn(),
            TimeElapsedColumn(),
            console=console
        ) as progress:
            batch_task = progress.add_task("Processando lote...", total=len(files))

            def file_cb(res: ProcessResult, current: int, total: int):
                progress.update(batch_task, description=f"[{current}/{total}] {os.path.basename(res.input_path)}", completed=current)

            results = optimizer.process_directory(folder_path, output_dir=args.output, file_callback=file_cb)

        for res in results:
            display_result_table(res)


if __name__ == "__main__":
    run_cli()
