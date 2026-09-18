"""
DocxFormatter Pro - Ponto de Entrada Principal.

Execução:
- Sem argumentos: Inicia a Interface Gráfica Desktop Moderna (CustomTkinter).
- Com argumentos: Executa a Interface de Linha de Comando (CLI) para automação.
"""

import sys

def main():
    # Se foram passados argumentos via linha de comando (exceto o nome do script), executa o CLI
    if len(sys.argv) > 1:
        from cli import run_cli
        run_cli()
    else:
        # Modo padrão: Interface Gráfica Desktop
        from gui.app import launch_gui
        launch_gui()

if __name__ == "__main__":
    main()
