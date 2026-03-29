"""
build.py — Gera NataliaTime.exe com PyInstaller
Uso: python build.py
"""
import subprocess, sys, os, shutil

def run(cmd, descricao):
    print(f"\n[...] {descricao}")
    r = subprocess.run(cmd, shell=True)
    if r.returncode != 0:
        print(f"\n[ERRO] Falhou: {descricao}")
        sys.exit(1)
    print(f"[ OK] {descricao}")

print("""
╔══════════════════════════════════════╗
║       Natalia Time — Build EXE       ║
╚══════════════════════════════════════╝
""")

run("pip install --upgrade pyinstaller",          "Verificando PyInstaller")
run("pip install numpy pandas matplotlib scipy openpyxl", "Verificando dependências")

for pasta in ["build", "dist"]:
    if os.path.exists(pasta):
        shutil.rmtree(pasta)
        print(f"[ OK] Pasta '{pasta}' removida")

run("pyinstaller NataliaTime.spec", "Gerando NataliaTime.exe (aguarde 1-3 min...)")

exe = os.path.join("dist", "NataliaTime.exe")
if os.path.isfile(exe):
    tamanho = os.path.getsize(exe) / 1_048_576
    print(f"""
════════════════════════════════════════
 Build concluído!
 Arquivo: dist\\NataliaTime.exe  ({tamanho:.0f} MB)

 Para distribuir, copie apenas:
   NataliaTime.exe
════════════════════════════════════════
""")
    os.startfile(os.path.abspath("dist"))
else:
    print("[ERRO] NataliaTime.exe não encontrado após build.")
    sys.exit(1)
