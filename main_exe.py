"""
main_exe.py — Entry point para o executável PyInstaller do Natalia Time.

Modos de operação:
  (sem args)         → carrega NT_app_ctk.py do disco e abre a interface
  --analyze <file>   → executa o script de análise gerado
  (args internos mp) → freeze_support() intercepta e roda worker do multiprocessing

Sobre multiprocessing no Windows (spawn):
  Ao criar workers, o Windows relança NataliaTime.exe do zero.
  Para que os workers encontrem _worker_init e demais funções via pickle,
  carregamos o script de análise em __main__ ANTES do freeze_support(),
  usando a variável de ambiente _NT_ANALYZE_SCRIPT.
"""

import os
import sys
import io

# ── Passo 0: em modo windowed (PyInstaller sem console), stdout/stderr são None.
# Workers do multiprocessing tentam escrever neles e crasham com AttributeError.
# Redirecionar para null device antes de qualquer outra coisa.
if sys.stdout is None:
    sys.stdout = io.TextIOWrapper(open(os.devnull, "wb"), encoding="utf-8", errors="replace")
if sys.stderr is None:
    sys.stderr = io.TextIOWrapper(open(os.devnull, "wb"), encoding="utf-8", errors="replace")

# ── Passo 1: se somos um worker, carregar funções de análise em __main__ ──────
# O processo pai (--analyze) define _NT_ANALYZE_SCRIPT antes de criar o Pool.
# Workers herdam essa variável e precisam ter as funções em __main__ para pickle.
_NT_SCRIPT = os.environ.get("_NT_ANALYZE_SCRIPT", "")
if _NT_SCRIPT and os.path.isfile(_NT_SCRIPT):
    import __main__ as _main_mod
    _ns = {"__name__": "_nt_worker_setup", "__file__": _NT_SCRIPT}
    with open(_NT_SCRIPT, encoding="utf-8") as _f:
        _code = _f.read()
    # Exec com __name__ != "__main__" → define funções mas não roda o bloco main
    exec(compile(_code, _NT_SCRIPT, "exec"), _ns)
    # Copia tudo para __main__ para o pickle encontrar (ex: _worker_init)
    for _k, _v in _ns.items():
        if not _k.startswith("__"):
            try:
                setattr(_main_mod, _k, _v)
            except Exception:
                pass

# ── Passo 2: freeze_support DEVE vir antes de qualquer código de lógica ───────
import multiprocessing
multiprocessing.freeze_support()

import importlib.util


def _base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _carregar_do_disco(nome, caminho):
    """Carrega um .py do disco em tempo de execução."""
    if not os.path.isfile(caminho):
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Natalia Time — Arquivo não encontrado",
            f"O arquivo '{os.path.basename(caminho)}' não foi encontrado.\n\n"
            f"Ele deve estar na mesma pasta que o executável:\n{_base_dir()}"
        )
        root.destroy()
        sys.exit(1)
    spec = importlib.util.spec_from_file_location(nome, caminho)
    mod  = importlib.util.module_from_spec(spec)
    sys.modules[nome] = mod
    spec.loader.exec_module(mod)
    return mod


if __name__ == "__main__":

    base = _base_dir()

    # ── Modo análise ──────────────────────────────────────────────────────────
    if "--analyze" in sys.argv:
        idx = sys.argv.index("--analyze")
        if idx + 1 >= len(sys.argv):
            print("Uso: NataliaTime.exe --analyze <script_temporario.py>")
            sys.exit(1)

        script_file = sys.argv[idx + 1]
        if not os.path.isfile(script_file):
            print(f"Arquivo não encontrado: {script_file}")
            sys.exit(1)

        # Informa workers qual script carregar (herdado via os.environ)
        os.environ["_NT_ANALYZE_SCRIPT"] = script_file

        # Força UTF-8 no stdout/stderr — Windows usa cp1252 por padrão
        if hasattr(sys.stdout, "buffer"):
            import io
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8",
                                          errors="replace", line_buffering=True)
            sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8",
                                          errors="replace", line_buffering=True)

        # Executa no namespace do __main__ real para pickle funcionar
        import __main__
        __main__.__file__ = script_file
        with open(script_file, encoding="utf-8") as f:
            codigo = f.read()
        exec(compile(codigo, script_file, "exec"), __main__.__dict__)
        sys.exit(0)

    # ── Modo GUI ──────────────────────────────────────────────────────────────
    mod = _carregar_do_disco("NT_app_ctk", os.path.join(base, "NT_app_ctk.py"))
    app = mod.NTApp()
    app.mainloop()
