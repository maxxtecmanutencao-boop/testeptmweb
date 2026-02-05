"""
PTM JSL - Sistema de Consultas 2.0
Lançador da Aplicação
"""
import os
import sys
import subprocess
from pathlib import Path

def main():
    print("=" * 50)
    print(" PTM JSL - SISTEMA DE CONSULTAS 2.0")
    print("=" * 50)
    print()
    
    # Obtém o caminho do script
    if getattr(sys, 'frozen', False):
        application_path = Path(sys.executable).parent
    else:
        application_path = Path(__file__).parent
    
    # Define o arquivo principal
    app_file = application_path / "app_melhorado.py"
    
    if not app_file.exists():
        print("❌ Erro: Arquivo app_melhorado.py não encontrado!")
        input("Pressione ENTER para sair...")
        sys.exit(1)
    
    print("✅ Iniciando o sistema...")
    print("📂 Pasta:", application_path)
    print()
    print("🌐 O navegador abrirá automaticamente em: http://localhost:8501")
    print()
    print("⚠️  Para encerrar o sistema, pressione CTRL+C nesta janela")
    print()
    print("=" * 50)
    print()
    
    # Inicia o Streamlit
    try:
        subprocess.run([
            sys.executable, "-m", "streamlit", "run",
            str(app_file),
            "--server.port", "8501",
            "--server.headless", "false"
        ], cwd=str(application_path))
    except KeyboardInterrupt:
        print("\n\n✅ Sistema encerrado pelo usuário.")
    except Exception as e:
        print(f"\n❌ Erro ao iniciar: {e}")
        input("\nPressione ENTER para sair...")

if __name__ == "__main__":
    main()