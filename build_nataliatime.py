"""
build_nataliatime.py — Gera NataliaTime.exe com PyInstaller
Uso: python build_nataliatime.py
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
║    Natalia Time — Build NataliaTime  ║
╚══════════════════════════════════════╝
""")

run("pip install --upgrade pyinstaller", "Verificando PyInstaller")

for pasta in ["build", "dist"]:
    if os.path.exists(pasta):
        shutil.rmtree(pasta)
        print(f"[ OK] Pasta '{pasta}' removida")

run("pyinstaller NataliaTime.spec", "Gerando NataliaTime.exe (aguarde)...")

exe = os.path.join("dist", "NataliaTime.exe")
if os.path.isfile(exe):
    tamanho = os.path.getsize(exe) / 1_048_576
    print(f"""
════════════════════════════════════════
 Build concluído!
 Arquivo: dist\\NataliaTime.exe  ({tamanho:.0f} MB)

 Para distribuir, copie para a mesma pasta:
   NataliaTime.exe
   analise_natalia_time.py   ← obrigatório
   NT_app_ctk.py             ← obrigatório

 O .exe carrega esses arquivos do disco
 em tempo de execução.
════════════════════════════════════════
""")
    os.startfile(os.path.abspath("dist"))
else:
    print("[ERRO] NataliaTime.exe não encontrado após build.")
    sys.exit(1)
