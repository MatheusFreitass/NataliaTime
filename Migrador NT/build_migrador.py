"""
build_migrador.py — Gera Migrador.exe com PyInstaller
Uso: python build_migrador.py
"""
import subprocess, sys, os, shutil

def run(cmd, descricao):
    print(f"\n[...] {descricao}")
    r = subprocess.run(cmd, shell=True)
    if r.returncode != 0:
        print(f"\n[ERRO] Falhou: {descricao}")
        sys.exit(1)
    print(f"[ OK] {descricao}")

import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

print("""
╔══════════════════════════════════════╗
║     Natalia Time — Build Migrador    ║
╚══════════════════════════════════════╝
""")

run("pip install --upgrade pyinstaller", "Verificando PyInstaller")

for pasta in ["build", "dist"]:
    if os.path.exists(pasta):
        shutil.rmtree(pasta)
        print(f"[ OK] Pasta '{pasta}' removida")

run("pyinstaller Migrador.spec", "Gerando Migrador.exe (aguarde)...")

exe = os.path.join("dist", "Migrador.exe")
if os.path.isfile(exe):
    tamanho = os.path.getsize(exe) / 1_048_576
    print(f"""
════════════════════════════════════════
 Build concluído!
 Arquivo: dist\\Migrador.exe  ({tamanho:.0f} MB)

 Para distribuir, copie apenas:
   Migrador.exe
════════════════════════════════════════
""")
    os.startfile(os.path.abspath("dist"))
else:
    print("[ERRO] Migrador.exe não encontrado após build.")
    sys.exit(1)
