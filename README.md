# DocxFormatter Pro & FileCompressor 📄🗜️✨

**Otimizador, Padronizador de Documentos Word (.docx) e Compressor Inteligente de Arquivos (PDF, DOCX, PPTX, XLSX e Imagens)** para ambiente local (Windows).

O **DocxFormatter Pro & FileCompressor** resolve os dois maiores problemas no fluxo de documentos corporativos e acadêmicos:
1. **Padronização e Estética de Layout:** textos tortos, imagens desalinhadas, linhas em branco em excesso, títulos isolados no final da página (órfãos) e tabelas cortadas.
2. **Inchaço e Peso Excessivo de Arquivos:** redução drástica do tamanho de PDFs, DOCXs, apresentações e imagens com **garantia de 100% dos dados e imagens preservados**.

---

## 🚀 Funcionalidades Principais

### 1. 🗜️ Redução Inteligente de Tamanho de Arquivos (Com Preservação Absoluta)
- **Garantia de Não-Exclusão:** Nenhuma imagem, texto, formulário ou tabela é excluído do arquivo. O sistema conta com validação automática pré e pós-processamento: se houver qualquer divergência na quantidade ou integridade de elementos visuais, a operação reverte imediatamente para o original.
- **Formatos Suportados:**
  - **PDF** (`.pdf`): Compactação de fluxos de dados e fontes (`deflate`), otimização segura de streams JPEG sem alterar filtros de XObjects, preservação de máscaras de transparência (`SMask`) e compressão de imagens via motor PyMuPDF com `clean=False`. Textos vetoriais continuam 100% nítidos em qualquer zoom.
  - **DOCX, PPTX e XLSX** (`.docx`, `.pptx`, `.xlsx`): Abre o pacote OpenXML, otimiza imagens raster mantendo estritamente a mesma extensão e MIME Type (PNG continua PNG, JPEG continua JPEG), preserva gráficos vetoriais (.emf, .wmf, .svg) e recompacta todos os XMLs com compressão zlib `Deflate nível 9`.
  - **Imagens Avulsas** (`.png`, `.jpg`, `.jpeg`, `.webp`, `.bmp`, `.tiff`): Compressão inteligente com Pillow, remoção de metadados inúteis e reamostragem `LANCZOS` de alta nitidez.
- **3 Níveis de Qualidade & Compressão:**
  - 🟢 **Alta Fidelidade (Mínima Perda) [Padrão]**: Qualidade 85%, resolução máx 2400px (300 DPI A4). Visivelmente idêntico ao original, ideal para relatórios oficiais, contratos e impressões.
  - 🟡 **Equilibrado (Recomendado)**: Qualidade 78%, resolução máx 1920px (Full HD). Excelente economia (reduções de 40% a 80%), ideal para e-mails e anexos.
  - 🔴 **Máxima Redução**: Qualidade 68%, resolução máx 1400px. Focado em caber em limites estritos de envio (ex: 5MB ou 10MB no SEI, Tribunais e órgãos públicos).
- **Regra "Never Worse":** Se o arquivo de entrada já estiver na compressão máxima ideal e a otimização não reduzir o tamanho, o sistema preserva o arquivo original intacto.

### 2. 📄 Alinhamento e Formatação de Texto (DOCX):
- Justificação automática do corpo do texto (ou alinhamento à esquerda conforme preferência).
- Padronização de espaçamento entrelinhas (ex: 1.15 corporativo ou 1.5 ABNT) e espaçamento após parágrafos.
- Aplicação consistente de fontes (Calibri, Arial, Times New Roman, Segoe UI) preservando negritos, itálicos e links.
- Ajuste inteligente de itens de lista (bullets e numerações) e títulos (Heading 1, 2, 3).

### 3. 🖼️ Alinhamento e Dimensionamento de Imagens (DOCX):
- Centralização automática de todas as imagens do documento.
- Remoção de recuos indesejados de parágrafos nas imagens (impedindo que fiquem tortas).
- **Auto-Ajuste Proporcional à Folha:** Redimensiona proporcionalmente imagens que excedem a largura útil da página.
- **Proteção Total:** Identifica imagens em qualquer formato (DrawingML, VML, OLE, AlternateContent, shapes) para nunca apagar parágrafos com elementos visuais.

### 4. 🧹 Eliminação de Linhas em Branco e Espaços Redundantes:
- Remoção de parágrafos vazios repetidos causados por múltiplos "Enters".
- Normalização de espaços múltiplos consecutivos dentro das frases.

### 5. 📑 Ajuste Fino por Páginas (Layout A4):
- **Prevenção de Títulos Órfãos (`keep_with_next`):** O título nunca fica isolado no rodapé da folha.
- **Proteção de Tabelas (`cantSplit` e `tblHeader`):** Impede linhas de tabelas cortadas e repete cabeçalhos em páginas subsequentes.

---

## 🛠️ Como Executar

### 1. Pela Interface Gráfica Desktop
- Dê um duplo clique no arquivo:
  `iniciar_programa.bat`
- Ou execute pelo terminal:
  ```bash
  python main.py
  ```
- **Na tela principal:**
  - Utilize a aba **"1. Formatar DOCX"** para padronizar e corrigir o layout de documentos Word.
  - Utilize a aba **"3. Reduzir Tamanho (PDF/DOCX)"** para comprimir arquivos PDF, DOCX ou imagens de forma rápida com métricas e validação em tempo real.

---

### 2. Pela Linha de Comando (CLI)

#### Reduzir Tamanho de Arquivo (PDF, DOCX, PPTX, XLSX, Imagens)

**Reduzir um arquivo específico:**
```bash
python main.py -f "meu_documento.pdf"
```
*(Arquivos `.pdf`, `.pptx`, `.xlsx` e imagens são automaticamente direcionados para o compressor!)*

**Reduzir com perfil equilibrado ou máxima redução:**
```bash
python main.py -f "relatorio.docx" -c --level equilibrado
python main.py -f "processo.pdf" -c --level max
```

**Reduzir todos os arquivos de uma pasta em lote:**
```bash
python main.py -d "C:/caminho/minha_pasta" -c --level alta -o "C:/caminho/pasta_saida"
```

#### Padronizar Formatação DOCX

**Formatar um documento Word:**
```bash
python main.py -f "caminho/do/documento.docx" -p corporativo
```

---

## 📁 Estrutura do Projeto

```text
Correcao texto/
│
├── core/                       # Núcleo de processamento e lógica
│   ├── __init__.py
│   ├── compressor.py           # Redutor inteligente com validação estrita de integridade
│   ├── config.py               # Definições de perfis e regras de estilo DOCX
│   ├── page_setup.py           # Tamanho de papel (A4), margens e cabeçalhos
│   ├── cleaner.py              # Remoção de linhas vazias com proteção contra perda de mídias
│   ├── image_handler.py        # Centralização e auto-ajuste de imagens
│   ├── table_handler.py        # Centralização e proteção de quebra de tabelas
│   ├── aligner.py              # Justificação, fontes, títulos e orfandade
│   └── processor.py            # Orquestrador da esteira de otimização DOCX
│
├── gui/                        # Interface Gráfica Desktop
│   ├── __init__.py
│   └── app.py                  # Janela CustomTkinter com abas de formatação e compressão
│
├── tests/                      # Scripts e arquivos de teste
│   ├── test_image_preservation.py # Testes de validação de preservação de 100% das imagens
│   ├── test_compressor.py         # Testes de compressão
│   └── generate_sample.py         # Gerador de documentos de teste
│
├── cli.py                      # Interface de linha de comando com Rich
├── main.py                     # Ponto de entrada universal
├── iniciar_programa.bat        # Inicializador rápido com dois cliques no Windows
└── requirements.txt            # Dependências Python (PyMuPDF, python-docx, CustomTkinter, Pillow, Rich)
```
