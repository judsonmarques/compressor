"""
Interface Gráfica Desktop Moderna para DocxFormatter Pro & FileCompressor.
Desenvolvida com CustomTkinter, oferecendo suporte a temas Claro/Escuro,
processamento assíncrono (sem travamentos), seleção de arquivos/pastas,
formatação e alinhamento de documentos DOCX e redução inteligente de tamanho
de arquivos (PDF, DOCX, PPTX, XLSX e Imagens) com mínima perda de conteúdo.

Design orientado a controle explícito: tarefas só são executadas sob clique direto
do usuário nos botões de execução da tela correspondente, sem execuções automáticas
ao navegar entre abas ou módulos, com botões dedicados de limpeza de tela e campos.
"""

import os
import sys
import threading
from typing import Optional, List

import customtkinter as ctk
from tkinter import filedialog, messagebox

from core.config import FormattingConfig, PRESETS
from core.processor import DocxOptimizer, ProcessResult
from core.compressor import FileCompressor, CompressionConfig, CompressionResult, format_file_size


# Definir padrão visual do CustomTkinter
ctk.set_appearance_mode("System")  # "System", "Dark", "Light"
ctk.set_default_color_theme("blue")  # "blue", "green", "dark-blue"


class DocxFormatterApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("DocxFormatter Pro & Compressor - Otimizador e Redutor de Arquivos")
        self.geometry("1060x780")
        self.minsize(920, 680)

        # Estado interno - Formatação
        self.current_config = PRESETS["Corporativo Moderno"]
        self.is_processing = False
        self.last_output_path: Optional[str] = None
        self.last_output_dir: Optional[str] = None

        # Estado interno - Compressão
        self.compression_presets = CompressionConfig.get_presets()
        self.current_compress_config = self.compression_presets["Alta Fidelidade (Mínima Perda)"]

        # Construir layout
        self._build_sidebar()
        self._build_main_area()
        self._load_config_into_ui(self.current_config)

    # =========================================================================
    # BARRA LATERAL ESQUERDA (SIDEBAR - APENAS NAVEGAÇÃO E AÇÕES GERAIS)
    # =========================================================================
    def _build_sidebar(self):
        """Constrói a barra lateral esquerda com navegação entre telas e controles gerais."""
        self.sidebar_frame = ctk.CTkFrame(self, width=240, corner_radius=0)
        self.sidebar_frame.pack(side="left", fill="y", padx=0, pady=0)
        self.sidebar_frame.pack_propagate(False)

        # Título / Logo
        title_label = ctk.CTkLabel(
            self.sidebar_frame,
            text="DocxFormatter",
            font=ctk.CTkFont(size=22, weight="bold")
        )
        title_label.pack(padx=20, pady=(20, 0), anchor="w")

        subtitle_label = ctk.CTkLabel(
            self.sidebar_frame,
            text="Pro & FileCompressor",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        )
        subtitle_label.pack(padx=20, pady=(0, 15), anchor="w")

        # Separador 1
        sep1 = ctk.CTkFrame(self.sidebar_frame, height=2, fg_color="gray30")
        sep1.pack(fill="x", padx=20, pady=(5, 10))

        # Seção de Navegação
        nav_label = ctk.CTkLabel(
            self.sidebar_frame,
            text="Navegação / Telas:",
            font=ctk.CTkFont(size=13, weight="bold")
        )
        nav_label.pack(padx=20, pady=(5, 5), anchor="w")

        # Botão Navegar: Formatar DOCX (apenas muda de aba, NUNCA executa automaticamente)
        self.btn_nav_format = ctk.CTkButton(
            self.sidebar_frame,
            text="📄 Formatar DOCX",
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#1f6aa5",
            hover_color="#144870",
            height=36,
            command=lambda: self.tabview.set("1. Formatar DOCX")
        )
        self.btn_nav_format.pack(padx=20, pady=4, fill="x")

        # Botão Navegar: Reduzir Tamanho (apenas muda de aba, NUNCA executa automaticamente)
        self.btn_nav_compress = ctk.CTkButton(
            self.sidebar_frame,
            text="🗜 Reduzir Tamanho",
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#7b2cbf",
            hover_color="#5a189a",
            height=36,
            command=lambda: self.tabview.set("3. Reduzir Tamanho (PDF/DOCX)")
        )
        self.btn_nav_compress.pack(padx=20, pady=4, fill="x")

        # Botão Navegar: Regras de Estilo
        self.btn_nav_rules = ctk.CTkButton(
            self.sidebar_frame,
            text="⚙️ Regras de Estilo",
            font=ctk.CTkFont(size=12),
            fg_color="transparent",
            border_width=1,
            text_color=("gray10", "gray90"),
            height=32,
            command=lambda: self.tabview.set("2. Regras de Estilo")
        )
        self.btn_nav_rules.pack(padx=20, pady=4, fill="x")

        # Botão Navegar: Relatório & Log
        self.btn_nav_logs = ctk.CTkButton(
            self.sidebar_frame,
            text="📊 Relatório & Log",
            font=ctk.CTkFont(size=12),
            fg_color="transparent",
            border_width=1,
            text_color=("gray10", "gray90"),
            height=32,
            command=lambda: self.tabview.set("4. Execução & Relatório")
        )
        self.btn_nav_logs.pack(padx=20, pady=4, fill="x")

        # Separador 2
        sep2 = ctk.CTkFrame(self.sidebar_frame, height=2, fg_color="gray30")
        sep2.pack(fill="x", padx=20, pady=(12, 10))

        # Perfil Rápido DOCX
        preset_label = ctk.CTkLabel(
            self.sidebar_frame,
            text="Perfil Formatação DOCX:",
            font=ctk.CTkFont(size=12, weight="bold")
        )
        preset_label.pack(padx=20, pady=(2, 4), anchor="w")

        self.preset_combobox = ctk.CTkComboBox(
            self.sidebar_frame,
            values=list(PRESETS.keys()) + ["Personalizado"],
            command=self._on_preset_change,
            width=200
        )
        self.preset_combobox.set("Corporativo Moderno")
        self.preset_combobox.pack(padx=20, pady=(0, 12), anchor="w")

        # Botão Limpar Toda a Tela (Zera todos os campos e reset)
        self.btn_clear_screen = ctk.CTkButton(
            self.sidebar_frame,
            text="🧹 Limpar Toda a Tela",
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color="#c1121f",
            hover_color="#780000",
            text_color="white",
            height=34,
            command=self._clear_all_screen
        )
        self.btn_clear_screen.pack(padx=20, pady=(5, 6), fill="x")

        # Botão para abrir pasta de saída
        self.open_folder_btn = ctk.CTkButton(
            self.sidebar_frame,
            text="📁 Abrir Pasta Destino",
            fg_color="transparent",
            border_width=1,
            text_color=("gray10", "gray90"),
            height=32,
            command=self._open_output_directory
        )
        self.open_folder_btn.pack(padx=20, pady=4, fill="x")

        # Espaçador vertical
        spacer = ctk.CTkLabel(self.sidebar_frame, text="")
        spacer.pack(fill="y", expand=True)

        # Seletor de Tema
        theme_label = ctk.CTkLabel(self.sidebar_frame, text="Tema Visual:", font=ctk.CTkFont(size=12))
        theme_label.pack(padx=20, pady=(10, 2), anchor="w")

        self.theme_menu = ctk.CTkOptionMenu(
            self.sidebar_frame,
            values=["System", "Dark", "Light"],
            command=ctk.set_appearance_mode,
            width=200
        )
        self.theme_menu.pack(padx=20, pady=(0, 20), anchor="w")

    def _build_main_area(self):
        """Constrói a área de abas com as funcionalidades principais."""
        self.tabview = ctk.CTkTabview(self)
        self.tabview.pack(side="right", fill="both", expand=True, padx=15, pady=15)

        self.tab_format_files = self.tabview.add("1. Formatar DOCX")
        self.tab_formatting_rules = self.tabview.add("2. Regras de Estilo")
        self.tab_compress = self.tabview.add("3. Reduzir Tamanho (PDF/DOCX)")
        self.tab_logs = self.tabview.add("4. Execução & Relatório")

        self._build_tab_format_files()
        self._build_tab_formatting_rules()
        self._build_tab_compress()
        self._build_tab_logs()

    # =========================================================================
    # ABA 1: FORMATAR DOCX (COM BOTÃO DE EXECUÇÃO MANUAL E BOTÃO DE LIMPEZA)
    # =========================================================================
    def _build_tab_format_files(self):
        tab = self.tab_format_files

        # Modo de operação
        mode_label = ctk.CTkLabel(tab, text="Modo de Operação (Formatação DOCX):", font=ctk.CTkFont(size=14, weight="bold"))
        mode_label.pack(padx=20, pady=(12, 4), anchor="w")

        self.mode_var = ctk.StringVar(value="file")
        mode_frame = ctk.CTkFrame(tab, fg_color="transparent")
        mode_frame.pack(padx=20, pady=(0, 10), anchor="w")

        rb_file = ctk.CTkRadioButton(
            mode_frame, text="Documento Único (.docx)",
            variable=self.mode_var, value="file",
            command=self._on_mode_change
        )
        rb_file.pack(side="left", padx=(0, 25))

        rb_folder = ctk.CTkRadioButton(
            mode_frame, text="Pasta Completa em Lote (Vários .docx)",
            variable=self.mode_var, value="folder",
            command=self._on_mode_change
        )
        rb_folder.pack(side="left")

        # Caixa Origem
        source_frame = ctk.CTkFrame(tab)
        source_frame.pack(fill="x", padx=20, pady=8)

        self.source_title = ctk.CTkLabel(
            source_frame,
            text="Documento Word de Origem (.docx):",
            font=ctk.CTkFont(size=13, weight="bold")
        )
        self.source_title.pack(padx=15, pady=(8, 4), anchor="w")

        source_input_frame = ctk.CTkFrame(source_frame, fg_color="transparent")
        source_input_frame.pack(fill="x", padx=15, pady=(0, 12))

        self.source_path_entry = ctk.CTkEntry(
            source_input_frame,
            placeholder_text="Clique em Procurar para selecionar o arquivo .docx..."
        )
        self.source_path_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self.btn_browse_source = ctk.CTkButton(
            source_input_frame,
            text="Procurar...",
            width=110,
            command=self._browse_source
        )
        self.btn_browse_source.pack(side="right")

        # Caixa Destino
        dest_frame = ctk.CTkFrame(tab)
        dest_frame.pack(fill="x", padx=20, pady=8)

        dest_title = ctk.CTkLabel(
            dest_frame,
            text="Local de Saída:",
            font=ctk.CTkFont(size=13, weight="bold")
        )
        dest_title.pack(padx=15, pady=(8, 4), anchor="w")

        self.dest_mode_var = ctk.BooleanVar(value=True)
        self.cb_same_dir = ctk.CTkCheckBox(
            dest_frame,
            text="Salvar no mesmo local com sufixo '_corrigido' (Recomendado)",
            variable=self.dest_mode_var,
            command=self._on_dest_mode_toggle
        )
        self.cb_same_dir.pack(padx=15, pady=(0, 8), anchor="w")

        self.dest_input_frame = ctk.CTkFrame(dest_frame, fg_color="transparent")
        self.dest_input_frame.pack(fill="x", padx=15, pady=(0, 12))

        self.dest_path_entry = ctk.CTkEntry(
            self.dest_input_frame,
            placeholder_text="Selecione a pasta de destino personalizada..."
        )
        self.dest_path_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.dest_path_entry.configure(state="disabled")

        self.btn_browse_dest = ctk.CTkButton(
            self.dest_input_frame,
            text="Procurar...",
            width=110,
            command=self._browse_dest,
            state="disabled"
        )
        self.btn_browse_dest.pack(side="right")

        # Barra de Ações Dedicada (EXECUTAR e LIMPAR CAMPOS)
        action_bar_format = ctk.CTkFrame(tab, fg_color="transparent")
        action_bar_format.pack(fill="x", padx=20, pady=(10, 8))

        self.btn_run_format = ctk.CTkButton(
            action_bar_format,
            text="▶ Executar Formatação DOCX",
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#1f6aa5",
            hover_color="#144870",
            height=44,
            command=self._start_processing
        )
        self.btn_run_format.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self.btn_clear_format = ctk.CTkButton(
            action_bar_format,
            text="🧹 Limpar Campos",
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="transparent",
            border_width=1,
            text_color=("gray10", "gray90"),
            height=44,
            width=140,
            command=self._clear_format_fields
        )
        self.btn_clear_format.pack(side="right")

        # Cartão de Informações Rápidas
        info_card = ctk.CTkFrame(tab, fg_color=("gray90", "gray18"))
        info_card.pack(fill="both", expand=True, padx=20, pady=(8, 12))

        info_title = ctk.CTkLabel(
            info_card,
            text="ℹ O que a padronização DOCX fará automaticamente:",
            font=ctk.CTkFont(size=13, weight="bold")
        )
        info_title.pack(padx=15, pady=(8, 4), anchor="w")

        bullets = [
            "• Alinhamento do texto e títulos com entrelinhas e espaçamentos uniformes",
            "• Centralização de todas as imagens e remoção de recuos indesejados nelas",
            "• Redimensionamento proporcional de imagens grandes para caber nas margens da folha",
            "• Eliminação de linhas em branco vazias repetidas que quebram o layout",
            "• Prevenção de quebra feia de linhas de tabelas e títulos isolados no fim da folha",
            "• Padronização de margens e tamanho de página (A4)"
        ]
        for b in bullets:
            b_label = ctk.CTkLabel(info_card, text=b, font=ctk.CTkFont(size=12), anchor="w")
            b_label.pack(padx=20, pady=1, anchor="w")

    # =========================================================================
    # ABA 2: REGRAS DE FORMATAÇÃO
    # =========================================================================
    def _build_tab_formatting_rules(self):
        tab = self.tab_formatting_rules

        scroll = ctk.CTkScrollableFrame(tab)
        scroll.pack(fill="both", expand=True, padx=10, pady=10)

        # 1. Grupo Texto & Tipografia
        grp_text = ctk.CTkFrame(scroll)
        grp_text.pack(fill="x", padx=10, pady=8)

        lbl_text = ctk.CTkLabel(grp_text, text="Tipografia & Parágrafos", font=ctk.CTkFont(size=14, weight="bold"))
        lbl_text.grid(row=0, column=0, columnspan=4, sticky="w", padx=15, pady=(10, 8))

        self.cb_change_font = ctk.CTkCheckBox(grp_text, text="Padronizar Família de Fonte")
        self.cb_change_font.grid(row=1, column=0, columnspan=2, sticky="w", padx=15, pady=5)

        self.font_name_combo = ctk.CTkComboBox(grp_text, values=["Calibri", "Arial", "Times New Roman", "Segoe UI"])
        self.font_name_combo.grid(row=1, column=2, columnspan=2, sticky="w", padx=15, pady=5)

        ctk.CTkLabel(grp_text, text="Tamanho Corpo (pt):").grid(row=2, column=0, sticky="w", padx=15, pady=5)
        self.entry_font_size = ctk.CTkEntry(grp_text, width=80)
        self.entry_font_size.grid(row=2, column=1, sticky="w", padx=15, pady=5)

        ctk.CTkLabel(grp_text, text="Alinhamento do Texto:").grid(row=2, column=2, sticky="w", padx=15, pady=5)
        self.combo_body_align = ctk.CTkComboBox(grp_text, values=["JUSTIFY", "LEFT", "CENTER"])
        self.combo_body_align.grid(row=2, column=3, sticky="w", padx=15, pady=5)

        ctk.CTkLabel(grp_text, text="Espaçamento Entrelinhas:").grid(row=3, column=0, sticky="w", padx=15, pady=5)
        self.entry_line_spacing = ctk.CTkEntry(grp_text, width=80)
        self.entry_line_spacing.grid(row=3, column=1, sticky="w", padx=15, pady=5)

        ctk.CTkLabel(grp_text, text="Espaço Depois (pt):").grid(row=3, column=2, sticky="w", padx=15, pady=5)
        self.entry_space_after = ctk.CTkEntry(grp_text, width=80)
        self.entry_space_after.grid(row=3, column=3, sticky="w", padx=15, pady=5)

        ctk.CTkLabel(grp_text, text="Recuo 1ª Linha (cm):").grid(row=4, column=0, sticky="w", padx=15, pady=5)
        self.entry_indent = ctk.CTkEntry(grp_text, width=80)
        self.entry_indent.grid(row=4, column=1, sticky="w", padx=15, pady=5)

        self.cb_keep_headings = ctk.CTkCheckBox(grp_text, text="Evitar títulos órfãos no fim da folha (keep_with_next)")
        self.cb_keep_headings.grid(row=5, column=0, columnspan=4, sticky="w", padx=15, pady=(8, 12))

        # 2. Grupo Imagens
        grp_img = ctk.CTkFrame(scroll)
        grp_img.pack(fill="x", padx=10, pady=8)

        lbl_img = ctk.CTkLabel(grp_img, text="Alinhamento & Ajuste de Imagens", font=ctk.CTkFont(size=14, weight="bold"))
        lbl_img.pack(padx=15, pady=(10, 8), anchor="w")

        self.cb_center_images = ctk.CTkCheckBox(grp_img, text="Centralizar todas as imagens automaticamente")
        self.cb_center_images.pack(padx=15, pady=4, anchor="w")

        self.cb_fit_images = ctk.CTkCheckBox(
            grp_img,
            text="Redimensionar imagens que ultrapassam as margens da folha (impede estouro de layout)"
        )
        self.cb_fit_images.pack(padx=15, pady=4, anchor="w")

        # 3. Grupo Tabelas
        grp_tbl = ctk.CTkFrame(scroll)
        grp_tbl.pack(fill="x", padx=10, pady=8)

        lbl_tbl = ctk.CTkLabel(grp_tbl, text="Tabelas & Paginação", font=ctk.CTkFont(size=14, weight="bold"))
        lbl_tbl.pack(padx=15, pady=(10, 8), anchor="w")

        self.cb_center_tables = ctk.CTkCheckBox(grp_tbl, text="Centralizar tabelas na página")
        self.cb_center_tables.pack(padx=15, pady=4, anchor="w")

        self.cb_cant_split = ctk.CTkCheckBox(
            grp_tbl,
            text="Impedir corte de linhas de tabelas entre páginas (cantSplit)"
        )
        self.cb_cant_split.pack(padx=15, pady=4, anchor="w")

        self.cb_repeat_header = ctk.CTkCheckBox(
            grp_tbl,
            text="Repetir linha de cabeçalho da tabela se ela ocupar mais de uma página"
        )
        self.cb_repeat_header.pack(padx=15, pady=4, anchor="w")

        # 4. Grupo Limpeza
        grp_clean = ctk.CTkFrame(scroll)
        grp_clean.pack(fill="x", padx=10, pady=8)

        lbl_clean = ctk.CTkLabel(grp_clean, text="Limpeza e Higienização", font=ctk.CTkFont(size=14, weight="bold"))
        lbl_clean.pack(padx=15, pady=(10, 8), anchor="w")

        self.cb_remove_empty = ctk.CTkCheckBox(grp_clean, text="Eliminar parágrafos e linhas em branco vazias")
        self.cb_remove_empty.pack(padx=15, pady=4, anchor="w")

        self.cb_remove_double_spaces = ctk.CTkCheckBox(grp_clean, text="Corrigir múltiplos espaços em branco consecutivos (ex: '  ' -> ' ')")
        self.cb_remove_double_spaces.pack(padx=15, pady=4, anchor="w")

        # 5. Grupo Margens & Página
        grp_page = ctk.CTkFrame(scroll)
        grp_page.pack(fill="x", padx=10, pady=8)

        lbl_page = ctk.CTkLabel(grp_page, text="Página & Margens (A4)", font=ctk.CTkFont(size=14, weight="bold"))
        lbl_page.grid(row=0, column=0, columnspan=4, sticky="w", padx=15, pady=(10, 8))

        ctk.CTkLabel(grp_page, text="Margem Superior (cm):").grid(row=1, column=0, sticky="w", padx=15, pady=4)
        self.entry_m_top = ctk.CTkEntry(grp_page, width=70)
        self.entry_m_top.grid(row=1, column=1, sticky="w", padx=15, pady=4)

        ctk.CTkLabel(grp_page, text="Margem Inferior (cm):").grid(row=1, column=2, sticky="w", padx=15, pady=4)
        self.entry_m_bottom = ctk.CTkEntry(grp_page, width=70)
        self.entry_m_bottom.grid(row=1, column=3, sticky="w", padx=15, pady=4)

        ctk.CTkLabel(grp_page, text="Margem Esquerda (cm):").grid(row=2, column=0, sticky="w", padx=15, pady=4)
        self.entry_m_left = ctk.CTkEntry(grp_page, width=70)
        self.entry_m_left.grid(row=2, column=1, sticky="w", padx=15, pady=4)

        ctk.CTkLabel(grp_page, text="Margem Direita (cm):").grid(row=2, column=2, sticky="w", padx=15, pady=4)
        self.entry_m_right = ctk.CTkEntry(grp_page, width=70)
        self.entry_m_right.grid(row=2, column=3, sticky="w", padx=15, pady=4)

    # =========================================================================
    # ABA 3: REDUZIR TAMANHO (COM BOTÃO DE EXECUÇÃO MANUAL E BOTÃO DE LIMPEZA)
    # =========================================================================
    def _build_tab_compress(self):
        tab = self.tab_compress

        scroll_c = ctk.CTkScrollableFrame(tab)
        scroll_c.pack(fill="both", expand=True, padx=10, pady=10)

        # Cabeçalho da funcionalidade
        header_frame = ctk.CTkFrame(scroll_c, fg_color="transparent")
        header_frame.pack(fill="x", padx=10, pady=(5, 8))

        c_title = ctk.CTkLabel(
            header_frame,
            text="🗜️ Redutor Inteligente de Tamanho de Arquivo",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        c_title.pack(anchor="w")

        c_desc = ctk.CTkLabel(
            header_frame,
            text="Comprime arquivos PDF, DOCX, PPTX, XLSX e Imagens com o mínimo de perda de qualidade visual.\n"
                 "Garante preservação de 100% dos dados, textos, tabelas e imagens originais.",
            font=ctk.CTkFont(size=12),
            text_color="gray",
            justify="left"
        )
        c_desc.pack(anchor="w", pady=(2, 0))

        # 1. Modo de Seleção (Único vs Lote)
        mode_box = ctk.CTkFrame(scroll_c)
        mode_box.pack(fill="x", padx=10, pady=6)

        lbl_c_mode = ctk.CTkLabel(mode_box, text="Modo de Redução:", font=ctk.CTkFont(size=13, weight="bold"))
        lbl_c_mode.pack(padx=15, pady=(8, 4), anchor="w")

        self.compress_mode_var = ctk.StringVar(value="file")
        cm_frame = ctk.CTkFrame(mode_box, fg_color="transparent")
        cm_frame.pack(padx=15, pady=(0, 8), anchor="w")

        self.rb_c_file = ctk.CTkRadioButton(
            cm_frame, text="Arquivo Único (PDF, DOCX, Imagem...)",
            variable=self.compress_mode_var, value="file",
            command=self._on_compress_mode_change
        )
        self.rb_c_file.pack(side="left", padx=(0, 25))

        self.rb_c_folder = ctk.CTkRadioButton(
            cm_frame, text="Pasta Completa em Lote",
            variable=self.compress_mode_var, value="folder",
            command=self._on_compress_mode_change
        )
        self.rb_c_folder.pack(side="left")

        # 2. Seleção de Origem
        source_c_box = ctk.CTkFrame(scroll_c)
        source_c_box.pack(fill="x", padx=10, pady=6)

        self.lbl_c_source_title = ctk.CTkLabel(
            source_c_box,
            text="Arquivo de Origem:",
            font=ctk.CTkFont(size=13, weight="bold")
        )
        self.lbl_c_source_title.pack(padx=15, pady=(8, 4), anchor="w")

        s_input_frame = ctk.CTkFrame(source_c_box, fg_color="transparent")
        s_input_frame.pack(fill="x", padx=15, pady=(0, 10))

        self.entry_compress_source = ctk.CTkEntry(
            s_input_frame,
            placeholder_text="Clique em Procurar para selecionar arquivo (PDF, DOCX, Imagem)..."
        )
        self.entry_compress_source.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self.btn_browse_compress_source = ctk.CTkButton(
            s_input_frame,
            text="Procurar...",
            width=110,
            command=self._browse_compress_source
        )
        self.btn_browse_compress_source.pack(side="right")

        # 3. Nível de Compressão (Presets)
        level_box = ctk.CTkFrame(scroll_c)
        level_box.pack(fill="x", padx=10, pady=6)

        lbl_level_title = ctk.CTkLabel(
            level_box,
            text="Nível de Redução & Qualidade Visual:",
            font=ctk.CTkFont(size=13, weight="bold")
        )
        lbl_level_title.pack(padx=15, pady=(8, 4), anchor="w")

        self.compress_level_combo = ctk.CTkComboBox(
            level_box,
            values=list(self.compression_presets.keys()),
            command=self._on_compress_level_change,
            width=320
        )
        self.compress_level_combo.set("Alta Fidelidade (Mínima Perda)")
        self.compress_level_combo.pack(padx=15, pady=(0, 6), anchor="w")

        self.lbl_level_desc = ctk.CTkLabel(
            level_box,
            text="🟢 Qualidade 85% | Máx 2400px | Deflate 9: Visivelmente idêntico ao original, ideal para relatórios, contratos e impressões.",
            font=ctk.CTkFont(size=11),
            text_color="#2ec4b6",
            justify="left",
            wraplength=700
        )
        self.lbl_level_desc.pack(padx=15, pady=(0, 10), anchor="w")

        # 4. Destino
        dest_c_box = ctk.CTkFrame(scroll_c)
        dest_c_box.pack(fill="x", padx=10, pady=6)

        dest_c_title = ctk.CTkLabel(
            dest_c_box,
            text="Local de Gravação:",
            font=ctk.CTkFont(size=13, weight="bold")
        )
        dest_c_title.pack(padx=15, pady=(8, 4), anchor="w")

        self.compress_dest_mode_var = ctk.BooleanVar(value=True)
        self.cb_compress_same_dir = ctk.CTkCheckBox(
            dest_c_box,
            text="Salvar no mesmo local com sufixo '_reduzido' (Recomendado)",
            variable=self.compress_dest_mode_var,
            command=self._on_compress_dest_toggle
        )
        self.cb_compress_same_dir.pack(padx=15, pady=(0, 6), anchor="w")

        self.compress_dest_frame = ctk.CTkFrame(dest_c_box, fg_color="transparent")
        self.compress_dest_frame.pack(fill="x", padx=15, pady=(0, 10))

        self.entry_compress_dest = ctk.CTkEntry(
            self.compress_dest_frame,
            placeholder_text="Selecione pasta de destino personalizada..."
        )
        self.entry_compress_dest.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.entry_compress_dest.configure(state="disabled")

        self.btn_browse_compress_dest = ctk.CTkButton(
            self.compress_dest_frame,
            text="Procurar...",
            width=110,
            command=self._browse_compress_dest,
            state="disabled"
        )
        self.btn_browse_compress_dest.pack(side="right")

        # 5. Barra de Ações Dedicada (EXECUTAR REDUÇÃO e LIMPAR CAMPOS)
        action_bar_compress = ctk.CTkFrame(scroll_c, fg_color="transparent")
        action_bar_compress.pack(fill="x", padx=10, pady=(10, 8))

        self.btn_run_compress = ctk.CTkButton(
            action_bar_compress,
            text="▶ Executar Redução de Tamanho",
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color="#7b2cbf",
            hover_color="#5a189a",
            height=46,
            command=self._start_compression
        )
        self.btn_run_compress.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self.btn_clear_compress = ctk.CTkButton(
            action_bar_compress,
            text="🧹 Limpar Campos",
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="transparent",
            border_width=1,
            text_color=("gray10", "gray90"),
            height=46,
            width=140,
            command=self._clear_compress_fields
        )
        self.btn_clear_compress.pack(side="right")

        # 6. Painel de Métricas / Resultados Visuais Rápidos
        self.metrics_card = ctk.CTkFrame(scroll_c, fg_color=("gray85", "gray15"))
        self.metrics_card.pack(fill="x", padx=10, pady=10)

        card_title = ctk.CTkLabel(
            self.metrics_card,
            text="📊 Resumo da Última Otimização:",
            font=ctk.CTkFont(size=13, weight="bold")
        )
        card_title.pack(padx=15, pady=(10, 8), anchor="w")

        metrics_grid = ctk.CTkFrame(self.metrics_card, fg_color="transparent")
        metrics_grid.pack(fill="x", padx=15, pady=(0, 12))

        # Coluna 1: Tamanho Original
        c1 = ctk.CTkFrame(metrics_grid)
        c1.pack(side="left", fill="both", expand=True, padx=4)
        ctk.CTkLabel(c1, text="Tamanho Original", font=ctk.CTkFont(size=11), text_color="gray").pack(pady=(6, 0))
        self.lbl_metric_orig = ctk.CTkLabel(c1, text="--", font=ctk.CTkFont(size=16, weight="bold"))
        self.lbl_metric_orig.pack(pady=(0, 6))

        # Coluna 2: Tamanho Reduzido
        c2 = ctk.CTkFrame(metrics_grid)
        c2.pack(side="left", fill="both", expand=True, padx=4)
        ctk.CTkLabel(c2, text="Tamanho Otimizado", font=ctk.CTkFont(size=11), text_color="gray").pack(pady=(6, 0))
        self.lbl_metric_comp = ctk.CTkLabel(c2, text="--", font=ctk.CTkFont(size=16, weight="bold"), text_color="#2ec4b6")
        self.lbl_metric_comp.pack(pady=(0, 6))

        # Coluna 3: Redução %
        c3 = ctk.CTkFrame(metrics_grid)
        c3.pack(side="left", fill="both", expand=True, padx=4)
        ctk.CTkLabel(c3, text="Taxa de Redução", font=ctk.CTkFont(size=11), text_color="gray").pack(pady=(6, 0))
        self.lbl_metric_pct = ctk.CTkLabel(c3, text="-- %", font=ctk.CTkFont(size=16, weight="bold"), text_color="#ffb703")
        self.lbl_metric_pct.pack(pady=(0, 6))

        # Coluna 4: Espaço Economizado
        c4 = ctk.CTkFrame(metrics_grid)
        c4.pack(side="left", fill="both", expand=True, padx=4)
        ctk.CTkLabel(c4, text="Espaço Poupado", font=ctk.CTkFont(size=11), text_color="gray").pack(pady=(6, 0))
        self.lbl_metric_saved = ctk.CTkLabel(c4, text="--", font=ctk.CTkFont(size=16, weight="bold"), text_color="#38b000")
        self.lbl_metric_saved.pack(pady=(0, 6))

    # =========================================================================
    # ABA 4: EXECUÇÃO & RELATÓRIO
    # =========================================================================
    def _build_tab_logs(self):
        tab = self.tab_logs

        # Barra de Progresso
        progress_frame = ctk.CTkFrame(tab, fg_color="transparent")
        progress_frame.pack(fill="x", padx=20, pady=(15, 10))

        self.status_label = ctk.CTkLabel(
            progress_frame,
            text="Pronto para iniciar.",
            font=ctk.CTkFont(size=13, weight="bold")
        )
        self.status_label.pack(anchor="w", pady=(0, 5))

        self.progress_bar = ctk.CTkProgressBar(progress_frame)
        self.progress_bar.pack(fill="x", pady=(0, 5))
        self.progress_bar.set(0)

        # Caixa de Log
        self.log_textbox = ctk.CTkTextbox(tab, font=ctk.CTkFont(family="Consolas", size=12))
        self.log_textbox.pack(fill="both", expand=True, padx=20, pady=(5, 15))

    # =========================================================================
    # FUNÇÕES DE LIMPEZA DE TELA E CAMPOS
    # =========================================================================
    def _clear_format_fields(self):
        """Limpa os campos de seleção da aba Formatar DOCX."""
        self.source_path_entry.delete(0, "end")
        self.dest_path_entry.configure(state="normal")
        self.dest_path_entry.delete(0, "end")
        self.dest_mode_var.set(True)
        self.dest_path_entry.configure(state="disabled")
        self.btn_browse_dest.configure(state="disabled")
        self.status_label.configure(text="Campos de formatação limpos.")

    def _clear_compress_fields(self):
        """Limpa os campos de seleção da aba Reduzir Tamanho."""
        self.entry_compress_source.delete(0, "end")
        self.entry_compress_dest.configure(state="normal")
        self.entry_compress_dest.delete(0, "end")
        self.compress_dest_mode_var.set(True)
        self.entry_compress_dest.configure(state="disabled")
        self.btn_browse_compress_dest.configure(state="disabled")

        # Resetar cards de métricas
        self.lbl_metric_orig.configure(text="--")
        self.lbl_metric_comp.configure(text="--")
        self.lbl_metric_pct.configure(text="-- %")
        self.lbl_metric_saved.configure(text="--")
        self.status_label.configure(text="Campos de redução limpos.")

    def _clear_all_screen(self):
        """Limpa completamente todos os campos de todas as abas, logs e progresso."""
        self._clear_format_fields()
        self._clear_compress_fields()
        self._clear_logs()
        self.progress_bar.set(0)
        self.last_output_path = None
        self.last_output_dir = None
        self.status_label.configure(text="Tela e campos limpos com sucesso. Pronto para iniciar.")

    # =========================================================================
    # EVENTOS E CONTROLE DE FORMATAÇÃO DOCX
    # =========================================================================
    def _on_preset_change(self, choice: str):
        if choice in PRESETS:
            config = PRESETS[choice]
            self._load_config_into_ui(config)
            self._log(f"Perfil de Formatação carregado: {choice}")

    def _load_config_into_ui(self, cfg: FormattingConfig):
        if cfg.change_font:
            self.cb_change_font.select()
        else:
            self.cb_change_font.deselect()

        self.font_name_combo.set(cfg.font_name)
        self.entry_font_size.delete(0, "end")
        self.entry_font_size.insert(0, str(cfg.font_size_pt))

        self.combo_body_align.set(cfg.body_alignment)
        self.entry_line_spacing.delete(0, "end")
        self.entry_line_spacing.insert(0, str(cfg.line_spacing))

        self.entry_space_after.delete(0, "end")
        self.entry_space_after.insert(0, str(cfg.space_after_pt))

        self.entry_indent.delete(0, "end")
        self.entry_indent.insert(0, str(cfg.first_line_indent_cm))

        if cfg.keep_headings_with_next:
            self.cb_keep_headings.select()
        else:
            self.cb_keep_headings.deselect()

        if cfg.center_images:
            self.cb_center_images.select()
        else:
            self.cb_center_images.deselect()

        if cfg.fit_images_to_page:
            self.cb_fit_images.select()
        else:
            self.cb_fit_images.deselect()

        if cfg.center_tables:
            self.cb_center_tables.select()
        else:
            self.cb_center_tables.deselect()

        if cfg.table_cant_split_rows:
            self.cb_cant_split.select()
        else:
            self.cb_cant_split.deselect()

        if cfg.table_repeat_header:
            self.cb_repeat_header.select()
        else:
            self.cb_repeat_header.deselect()

        if cfg.remove_empty_paragraphs:
            self.cb_remove_empty.select()
        else:
            self.cb_remove_empty.deselect()

        if cfg.remove_redundant_spaces:
            self.cb_remove_double_spaces.select()
        else:
            self.cb_remove_double_spaces.deselect()

        self.entry_m_top.delete(0, "end")
        self.entry_m_top.insert(0, str(cfg.margin_top_cm))

        self.entry_m_bottom.delete(0, "end")
        self.entry_m_bottom.insert(0, str(cfg.margin_bottom_cm))

        self.entry_m_left.delete(0, "end")
        self.entry_m_left.insert(0, str(cfg.margin_left_cm))

        self.entry_m_right.delete(0, "end")
        self.entry_m_right.insert(0, str(cfg.margin_right_cm))

    def _collect_config_from_ui(self) -> FormattingConfig:
        cfg = FormattingConfig()
        cfg.name = self.preset_combobox.get()
        cfg.change_font = bool(self.cb_change_font.get())
        cfg.font_name = self.font_name_combo.get()

        try:
            cfg.font_size_pt = float(self.entry_font_size.get() or 11.0)
        except ValueError:
            cfg.font_size_pt = 11.0

        cfg.body_alignment = self.combo_body_align.get()

        try:
            cfg.line_spacing = float(self.entry_line_spacing.get() or 1.15)
        except ValueError:
            cfg.line_spacing = 1.15

        try:
            cfg.space_after_pt = float(self.entry_space_after.get() or 6.0)
        except ValueError:
            cfg.space_after_pt = 6.0

        try:
            cfg.first_line_indent_cm = float(self.entry_indent.get() or 0.0)
        except ValueError:
            cfg.first_line_indent_cm = 0.0

        cfg.keep_headings_with_next = bool(self.cb_keep_headings.get())
        cfg.center_images = bool(self.cb_center_images.get())
        cfg.fit_images_to_page = bool(self.cb_fit_images.get())
        cfg.center_tables = bool(self.cb_center_tables.get())
        cfg.table_cant_split_rows = bool(self.cb_cant_split.get())
        cfg.table_repeat_header = bool(self.cb_repeat_header.get())
        cfg.remove_empty_paragraphs = bool(self.cb_remove_empty.get())
        cfg.remove_redundant_spaces = bool(self.cb_remove_double_spaces.get())

        try:
            cfg.margin_top_cm = float(self.entry_m_top.get() or 2.5)
            cfg.margin_bottom_cm = float(self.entry_m_bottom.get() or 2.5)
            cfg.margin_left_cm = float(self.entry_m_left.get() or 2.5)
            cfg.margin_right_cm = float(self.entry_m_right.get() or 2.5)
        except ValueError:
            pass

        return cfg

    def _on_mode_change(self):
        mode = self.mode_var.get()
        if mode == "file":
            self.source_title.configure(text="Documento Word de Origem (.docx):")
            self.source_path_entry.configure(placeholder_text="Clique em Procurar para selecionar o arquivo .docx...")
        else:
            self.source_title.configure(text="Pasta com Arquivos Word (.docx):")
            self.source_path_entry.configure(placeholder_text="Clique em Procurar para selecionar a pasta com os arquivos...")

    def _on_dest_mode_toggle(self):
        custom = not self.dest_mode_var.get()
        state = "normal" if custom else "disabled"
        self.dest_path_entry.configure(state=state)
        self.btn_browse_dest.configure(state=state)

    def _browse_source(self):
        mode = self.mode_var.get()
        if mode == "file":
            path = filedialog.askopenfilename(
                title="Selecione o documento Word",
                filetypes=[("Documentos Word", "*.docx")]
            )
        else:
            path = filedialog.askdirectory(title="Selecione a pasta com documentos Word")

        if path:
            self.source_path_entry.delete(0, "end")
            self.source_path_entry.insert(0, os.path.normpath(path))

    def _browse_dest(self):
        path = filedialog.askdirectory(title="Selecione a pasta de destino")
        if path:
            self.dest_path_entry.delete(0, "end")
            self.dest_path_entry.insert(0, os.path.normpath(path))

    # =========================================================================
    # EVENTOS E CONTROLE DE REDUÇÃO DE TAMANHO (COMPRESSÃO)
    # =========================================================================
    def _on_compress_mode_change(self):
        mode = self.compress_mode_var.get()
        if mode == "file":
            self.lbl_c_source_title.configure(text="Arquivo de Origem:")
            self.entry_compress_source.configure(placeholder_text="Clique em Procurar para selecionar arquivo (PDF, DOCX, Imagem)...")
        else:
            self.lbl_c_source_title.configure(text="Pasta com Arquivos para Reduzir:")
            self.entry_compress_source.configure(placeholder_text="Clique em Procurar para selecionar pasta com arquivos...")

    def _on_compress_level_change(self, choice: str):
        if choice in self.compression_presets:
            self.current_compress_config = self.compression_presets[choice]
            if choice == "Alta Fidelidade (Mínima Perda)":
                self.lbl_level_desc.configure(
                    text="🟢 Qualidade 85% | Máx 2400px | Deflate 9: Visivelmente idêntico ao original, ideal para relatórios, contratos e impressões.",
                    text_color="#2ec4b6"
                )
            elif choice == "Equilibrado (Recomendado)":
                self.lbl_level_desc.configure(
                    text="🟡 Qualidade 78% | Máx 1920px Full HD: Excelente economia de espaço, ideal para envio por e-mail e web com ótima nitidez.",
                    text_color="#ffb703"
                )
            else:
                self.lbl_level_desc.configure(
                    text="🔴 Qualidade 68% | Máx 1400px HD: Redução máxima para caber em limites rígidos de anexo (ex: 5MB ou 10MB no SEI/Tribunais).",
                    text_color="#ff6b6b"
                )

    def _on_compress_dest_toggle(self):
        custom = not self.compress_dest_mode_var.get()
        state = "normal" if custom else "disabled"
        self.entry_compress_dest.configure(state=state)
        self.btn_browse_compress_dest.configure(state=state)

    def _browse_compress_source(self):
        mode = self.compress_mode_var.get()
        if mode == "file":
            path = filedialog.askopenfilename(
                title="Selecione o arquivo para reduzir o tamanho",
                filetypes=[
                    ("Todos os Suportados", "*.pdf;*.docx;*.pptx;*.xlsx;*.png;*.jpg;*.jpeg;*.webp;*.bmp;*.tiff"),
                    ("Documentos PDF (*.pdf)", "*.pdf"),
                    ("Documentos Word (*.docx)", "*.docx"),
                    ("Apresentações PowerPoint (*.pptx)", "*.pptx"),
                    ("Planilhas Excel (*.xlsx)", "*.xlsx"),
                    ("Imagens (*.png;*.jpg;*.jpeg;*.webp;*.bmp;*.tiff)", "*.png;*.jpg;*.jpeg;*.webp;*.bmp;*.tiff"),
                    ("Todos os Arquivos", "*.*")
                ]
            )
        else:
            path = filedialog.askdirectory(title="Selecione a pasta com arquivos para reduzir")

        if path:
            self.entry_compress_source.delete(0, "end")
            self.entry_compress_source.insert(0, os.path.normpath(path))

    def _browse_compress_dest(self):
        path = filedialog.askdirectory(title="Selecione a pasta de destino para os arquivos reduzidos")
        if path:
            self.entry_compress_dest.delete(0, "end")
            self.entry_compress_dest.insert(0, os.path.normpath(path))

    # =========================================================================
    # EXECUÇÃO DE REDUÇÃO DE TAMANHO (SOMENTE SOB CLIQUE EXPLÍCITO NO BOTÃO)
    # =========================================================================
    def _start_compression(self):
        """Inicia a rotina de compressão SOMENTE sob clique explícito no botão Executar."""
        if self.is_processing:
            return

        source = self.entry_compress_source.get().strip()
        if not source or not os.path.exists(source):
            messagebox.showwarning("Atenção", "Por favor, selecione um arquivo ou pasta válido na aba Reduzir Tamanho!")
            return

        dest_custom = not self.compress_dest_mode_var.get()
        custom_dest = self.entry_compress_dest.get().strip() if dest_custom else None

        self.tabview.set("4. Execução & Relatório")
        self.is_processing = True
        self.btn_run_compress.configure(state="disabled", text="⏳ Reduzindo Arquivo...")
        self.btn_run_format.configure(state="disabled")
        self.progress_bar.set(0)
        self.status_label.configure(text="Iniciando redução de tamanho...")

        thread = threading.Thread(
            target=self._run_compression_thread,
            args=(source, self.current_compress_config, custom_dest),
            daemon=True
        )
        thread.start()

    def _run_compression_thread(self, source: str, config: CompressionConfig, custom_dest: Optional[str]):
        """Thread trabalhadora da redução de tamanho."""
        compressor = FileCompressor(config)
        mode = self.compress_mode_var.get()

        self._log("=" * 60)
        self._log(f"Iniciando Redução de Tamanho | Nível: {config.name}")
        self._log(f"Alvo: {source}")
        self._log("=" * 60)

        if mode == "file":
            def progress_cb(msg: str, val: float):
                self.after(0, lambda: self._update_progress(msg, val))

            result = compressor.compress_file(
                source,
                output_path=custom_dest,
                progress_callback=progress_cb
            )

            self.after(0, lambda: self._on_single_compression_complete(result))

        else:
            supported_exts = compressor.get_supported_extensions()
            files = [
                os.path.join(source, f)
                for f in os.listdir(source)
                if os.path.isfile(os.path.join(source, f))
                and os.path.splitext(f)[1].lower() in supported_exts
                and not f.startswith("~$")
                and "_reduzido" not in f
            ]
            total_files = len(files)

            if total_files == 0:
                self.after(0, lambda: self._on_compress_batch_empty(source))
                return

            def file_cb(res: CompressionResult, current: int, total: int):
                self.after(0, lambda: self._on_compress_batch_step(res, current, total))

            results = compressor.compress_directory(source, output_dir=custom_dest, file_callback=file_cb)
            self.after(0, lambda: self._on_compress_batch_complete(results, custom_dest or source))

    def _on_single_compression_complete(self, res: CompressionResult):
        self.is_processing = False
        self.btn_run_compress.configure(state="normal", text="▶ Executar Redução de Tamanho")
        self.btn_run_format.configure(state="normal", text="▶ Executar Formatação DOCX")

        if not res.success:
            self._log(f"❌ ERRO na compressão: {res.error_message}")
            self.status_label.configure(text="Erro durante a redução do arquivo.")
            messagebox.showerror("Erro", f"Não foi possível reduzir o arquivo:\n{res.error_message}")
            return

        self.last_output_path = res.output_path
        self.last_output_dir = os.path.dirname(res.output_path)
        self.progress_bar.set(1.0)
        self.status_label.configure(text=f"Redução concluída com sucesso em {res.elapsed_time_seconds}s!")

        # Atualizar cards de métricas
        self.lbl_metric_orig.configure(text=res.original_size_str)
        self.lbl_metric_comp.configure(text=res.compressed_size_str)
        self.lbl_metric_pct.configure(text=f"-{res.reduction_percentage}%")
        self.lbl_metric_saved.configure(text=res.saved_size_str)

        self._log(f"✓ Arquivo otimizado com sucesso!")
        self._log(f"  • Formato: {res.file_type}")
        self._log(f"  • Tamanho Original: {res.original_size_str}")
        self._log(f"  • Tamanho Final: {res.compressed_size_str}")
        self._log(f"  • Redução: {res.reduction_percentage}% (Economizou {res.saved_size_str})")
        if res.images_optimized > 0:
            self._log(f"  • Imagens internas otimizadas: {res.images_optimized}")
        if "note" in res.details:
            self._log(f"  • Nota: {res.details['note']}")
        self._log(f"Salvo em: {res.output_path}\n")

        # Voltar visualmente para a aba de compressão para exibir as métricas
        self.tabview.set("3. Reduzir Tamanho (PDF/DOCX)")

        messagebox.showinfo(
            "Sucesso!",
            f"Arquivo reduzido com sucesso!\n\n"
            f"• Original: {res.original_size_str}\n"
            f"• Otimizado: {res.compressed_size_str}\n"
            f"• Redução: {res.reduction_percentage}% ({res.saved_size_str} economizados)\n\n"
            f"Salvo em:\n{res.output_path}"
        )

    def _on_compress_batch_step(self, res: CompressionResult, current: int, total: int):
        val = current / total
        self.progress_bar.set(val)
        fname = os.path.basename(res.input_path)
        if res.success:
            self.status_label.configure(text=f"Reduzido ({current}/{total}): {fname}")
            self._log(f"[{current}/{total}] ✓ {fname} ({res.original_size_str} -> {res.compressed_size_str}, -{res.reduction_percentage}%)")
        else:
            self.status_label.configure(text=f"Falha em ({current}/{total}): {fname}")
            self._log(f"[{current}/{total}] ❌ Falha em {fname}: {res.error_message}")

    def _on_compress_batch_empty(self, folder: str):
        self.is_processing = False
        self.btn_run_compress.configure(state="normal", text="▶ Executar Redução de Tamanho")
        self.btn_run_format.configure(state="normal", text="▶ Executar Formatação DOCX")
        self.status_label.configure(text="Nenhum arquivo elegível encontrado.")
        self._log(f"Nenhum arquivo elegível (PDF, DOCX, Imagens) encontrado na pasta: {folder}")
        messagebox.showwarning("Aviso", "Nenhum arquivo suportado para redução encontrado na pasta selecionada.")

    def _on_compress_batch_complete(self, results: List[CompressionResult], output_dir: str):
        self.is_processing = False
        self.btn_run_compress.configure(state="normal", text="▶ Executar Redução de Tamanho")
        self.btn_run_format.configure(state="normal", text="▶ Executar Formatação DOCX")
        self.last_output_dir = output_dir

        success_count = sum(1 for r in results if r.success)
        total_orig = sum(r.original_size for r in results if r.success)
        total_comp = sum(r.compressed_size for r in results if r.success)
        total_saved = max(0, total_orig - total_comp)
        overall_pct = round((total_saved / total_orig * 100), 1) if total_orig > 0 else 0.0

        self.progress_bar.set(1.0)
        self.status_label.configure(text=f"Lote finalizado! {success_count}/{len(results)} arquivos reduzidos.")

        self.lbl_metric_orig.configure(text=format_file_size(total_orig))
        self.lbl_metric_comp.configure(text=format_file_size(total_comp))
        self.lbl_metric_pct.configure(text=f"-{overall_pct}%")
        self.lbl_metric_saved.configure(text=format_file_size(total_saved))

        self._log("-" * 60)
        self._log(f"Redução em lote finalizada: {success_count} de {len(results)} processados com sucesso.")
        self._log(f"Economia total em disco: {format_file_size(total_saved)} (-{overall_pct}%)")
        self._log(f"Pasta de destino: {output_dir}\n")

        self.tabview.set("3. Reduzir Tamanho (PDF/DOCX)")

        messagebox.showinfo(
            "Lote Finalizado",
            f"Processamento de redução em lote concluído!\n\n"
            f"• Arquivos processados: {success_count} de {len(results)}\n"
            f"• Espaço total economizado: {format_file_size(total_saved)} (-{overall_pct}%)\n\n"
            f"Arquivos salvos em:\n{output_dir}"
        )

    # =========================================================================
    # EXECUÇÃO DE FORMATAÇÃO DOCX (SOMENTE SOB CLIQUE EXPLÍCITO NO BOTÃO)
    # =========================================================================
    def _start_processing(self):
        """Inicia a rotina de formatação SOMENTE sob clique explícito no botão Executar."""
        if self.is_processing:
            return

        source = self.source_path_entry.get().strip()
        if not source or not os.path.exists(source):
            messagebox.showwarning("Atenção", "Por favor, selecione um arquivo ou pasta válido na aba Formatar DOCX!")
            return

        config = self._collect_config_from_ui()
        dest_custom = not self.dest_mode_var.get()
        custom_dest = self.dest_path_entry.get().strip() if dest_custom else None

        self.tabview.set("4. Execução & Relatório")
        self.is_processing = True
        self.btn_run_format.configure(state="disabled", text="⏳ Formatando...")
        self.btn_run_compress.configure(state="disabled")
        self.progress_bar.set(0)
        self.status_label.configure(text="Iniciando formatação...")

        thread = threading.Thread(
            target=self._run_optimization_thread,
            args=(source, config, custom_dest),
            daemon=True
        )
        thread.start()

    def _run_optimization_thread(self, source: str, config: FormattingConfig, custom_dest: Optional[str]):
        optimizer = DocxOptimizer(config)
        mode = self.mode_var.get()

        self._log("=" * 60)
        self._log(f"Iniciando Formatação DOCX | Perfil: {config.name}")
        self._log(f"Alvo: {source}")
        self._log("=" * 60)

        if mode == "file":
            def progress_cb(msg: str, val: float):
                self.after(0, lambda: self._update_progress(msg, val))

            result = optimizer.process_file(
                source,
                output_path=custom_dest,
                progress_callback=progress_cb
            )

            self.after(0, lambda: self._on_single_complete(result))

        else:
            docx_files = [
                os.path.join(source, f)
                for f in os.listdir(source)
                if f.lower().endswith(".docx") and not f.startswith("~$") and not "_corrigido" in f
            ]
            total_files = len(docx_files)

            if total_files == 0:
                self.after(0, lambda: self._on_batch_empty(source))
                return

            def file_cb(res: ProcessResult, current: int, total: int):
                self.after(0, lambda: self._on_batch_file_step(res, current, total))

            results = optimizer.process_directory(source, output_dir=custom_dest, file_callback=file_cb)
            self.after(0, lambda: self._on_batch_complete(results, custom_dest or source))

    def _update_progress(self, msg: str, val: float):
        self.status_label.configure(text=msg)
        self.progress_bar.set(val)

    def _on_single_complete(self, res: ProcessResult):
        self.is_processing = False
        self.btn_run_format.configure(state="normal", text="▶ Executar Formatação DOCX")
        self.btn_run_compress.configure(state="normal", text="▶ Executar Redução de Tamanho")

        if not res.success:
            self._log(f"❌ ERRO: {res.error_message}")
            self.status_label.configure(text="Erro durante a formatação.")
            messagebox.showerror("Erro", f"Não foi possível processar o documento:\n{res.error_message}")
            return

        self.last_output_path = res.output_path
        self.last_output_dir = os.path.dirname(res.output_path)
        self.progress_bar.set(1.0)
        self.status_label.configure(text=f"Concluído com sucesso em {res.elapsed_time_seconds}s!")

        stats = res.stats
        c_stat = stats.get("cleaner", {})
        img_stat = stats.get("images", {})
        tbl_stat = stats.get("tables", {})
        txt_stat = stats.get("text", {})

        self._log(f"✓ Documento padronizado com sucesso!")
        self._log(f"  • Parágrafos vazios eliminados: {c_stat.get('empty_paragraphs_removed', 0)}")
        self._log(f"  • Espaços duplos corrigidos: {c_stat.get('runs_spaces_cleaned', 0)}")
        self._log(f"  • Imagens centralizadas: {img_stat.get('images_centered', 0)} de {img_stat.get('images_found', 0)}")
        self._log(f"  • Imagens redimensionadas para caber na página: {img_stat.get('images_resized', 0)}")
        self._log(f"  • Tabelas centralizadas: {tbl_stat.get('tables_adjusted', 0)}")
        self._log(f"  • Linhas de tabela protegidas contra corte: {tbl_stat.get('rows_prevented_from_splitting', 0)}")
        self._log(f"  • Parágrafos de texto justificados/formatados: {txt_stat.get('body_paragraphs_aligned', 0)}")
        self._log(f"  • Títulos protegidos contra orfandade: {txt_stat.get('headings_adjusted', 0)}")
        self._log(f"Salvo em: {res.output_path}\n")

        messagebox.showinfo(
            "Sucesso!",
            f"O documento foi formatado com sucesso!\n\nArquivo salvo em:\n{res.output_path}"
        )

    def _on_batch_empty(self, folder: str):
        self.is_processing = False
        self.btn_run_format.configure(state="normal", text="▶ Executar Formatação DOCX")
        self.btn_run_compress.configure(state="normal", text="▶ Executar Redução de Tamanho")
        self.status_label.configure(text="Nenhum arquivo encontrado.")
        self._log(f"Nenhum arquivo .docx elegível encontrado na pasta: {folder}")
        messagebox.showwarning("Aviso", "Nenhum arquivo .docx elegível encontrado na pasta selecionada.")

    def _on_batch_file_step(self, res: ProcessResult, current: int, total: int):
        val = current / total
        self.progress_bar.set(val)
        fname = os.path.basename(res.input_path)
        if res.success:
            self.status_label.configure(text=f"Formatado ({current}/{total}): {fname}")
            self._log(f"[{current}/{total}] ✓ {fname} -> {os.path.basename(res.output_path)}")
        else:
            self.status_label.configure(text=f"Falha em ({current}/{total}): {fname}")
            self._log(f"[{current}/{total}] ❌ Falha em {fname}: {res.error_message}")

    def _on_batch_complete(self, results: list, output_dir: str):
        self.is_processing = False
        self.btn_run_format.configure(state="normal", text="▶ Executar Formatação DOCX")
        self.btn_run_compress.configure(state="normal", text="▶ Executar Redução de Tamanho")
        self.last_output_dir = output_dir

        success_count = sum(1 for r in results if r.success)
        self.progress_bar.set(1.0)
        self.status_label.configure(text=f"Lote finalizado! {success_count}/{len(results)} com sucesso.")
        self._log("-" * 60)
        self._log(f"Processamento em lote finalizado: {success_count} de {len(results)} arquivos formatados.")
        self._log(f"Pasta de destino: {output_dir}\n")

        messagebox.showinfo(
            "Lote Finalizado",
            f"O processamento em lote foi concluído!\n\n{success_count} de {len(results)} arquivos foram formatados com sucesso."
        )

    def _log(self, text: str):
        self.log_textbox.insert("end", text + "\n")
        self.log_textbox.see("end")

    def _clear_logs(self):
        self.log_textbox.delete("1.0", "end")

    def _open_output_directory(self):
        target = self.last_output_dir
        if not target or not os.path.exists(target):
            src = self.source_path_entry.get().strip() or self.entry_compress_source.get().strip()
            if src and os.path.exists(src):
                target = src if os.path.isdir(src) else os.path.dirname(src)

        if target and os.path.exists(target):
            os.startfile(target)
        else:
            messagebox.showinfo("Informação", "Nenhum documento processado ainda.")


def launch_gui():
    app = DocxFormatterApp()
    app.mainloop()


if __name__ == "__main__":
    launch_gui()
