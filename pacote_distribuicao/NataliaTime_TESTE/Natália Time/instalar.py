"""
instalar.py — Instala as dependências do Natalia Time
Uso: python instalar.py
"""
import subprocess, sys

pkgs = ["numpy", "pandas", "matplotlib", "scipy", "openpyxl", "pyinstaller"]

print("Instalando dependências do Natalia Time...\n")
subprocess.run([sys.executable, "-m", "pip", "install", "--upgrade"] + pkgs, check=True)
print("\nConcluído! Você já pode:\n  Rodar:    python NT_app_tk.py\n  Compilar: python build.py\n")
