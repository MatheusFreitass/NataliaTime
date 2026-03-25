"""
NT_app.py — Interface Natalia Time (Tkinter)
Uso: python NT_app.py
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import subprocess
import sys
import os
import glob
import re
import tempfile
import importlib.util

# ── Detectar base dir ─────────────────────────────────────────────────────────
if getattr(sys, 'frozen', False):
    # Executável PyInstaller: arquivos de dados estão em _MEIPASS (onefile)
    # e o .exe fica ao lado dos arquivos de trabalho
    _base     = os.path.dirname(sys.executable)
    _meipass  = getattr(sys, '_MEIPASS', _base)
else:
    _base    = os.path.dirname(os.path.abspath(__file__))
    _meipass = _base

# analise_natalia_time.py: disco tem prioridade sobre bundled
# (permite atualizar o script sem rebuildar o .exe)
SCRIPT_PATH = os.path.join(_base, "analise_natalia_time.py")
if not os.path.isfile(SCRIPT_PATH):
    SCRIPT_PATH = os.path.join(_meipass, "analise_natalia_time.py")
RESULTADOS_DIR = os.path.join(_base, "Resultados")

# ── Cores ─────────────────────────────────────────────────────────────────────
BG      = "#0e1117"
BG2     = "#161b22"
BG3     = "#1c2128"
BORDA   = "#30363d"
AZUL    = "#388bfd"
TEXTO   = "#e0e0e0"
TEXTO2  = "#8b949e"
VERDE   = "#3fb950"
VERM    = "#f85149"
AMAR    = "#d29922"
FONTE   = ("Segoe UI", 9)
FONTE_M = ("Courier New", 9)
FONTE_T = ("Courier New", 10, "bold")

# ── Helpers de widgets ────────────────────────────────────────────────────────
def lbl(parent, texto, fg=TEXTO2, font=FONTE, **kw):
    return tk.Label(parent, text=texto, bg=parent["bg"] if hasattr(parent,"bg") else BG2,
                    fg=fg, font=font, **kw)

def entry(parent, width=12, default=""):
    e = tk.Entry(parent, bg=BG3, fg=TEXTO, insertbackground=TEXTO,
                 relief="flat", font=FONTE, width=width,
                 highlightthickness=1, highlightbackground=BORDA,
                 highlightcolor=AZUL)
    e.insert(0, default)
    return e

def btn(parent, texto, cmd, cor=BG2, fg=TEXTO, **kw):
    return tk.Button(parent, text=texto, command=cmd,
                     bg=cor, fg=fg, activebackground=AZUL,
                     activeforeground="white", relief="flat",
                     font=FONTE_M, cursor="hand2", padx=10, pady=5, **kw)

def check(parent, texto, valor=True, state="normal"):
    v = tk.BooleanVar(value=valor)
    cb = tk.Checkbutton(parent, text=texto, variable=v,
                        bg=parent["bg"] if hasattr(parent,"bg") else BG2,
                        fg=TEXTO, selectcolor=BG3,
                        activebackground=BG2, activeforeground=AZUL,
                        font=FONTE, state=state)
    return cb, v

def secao(parent, titulo):
    f = tk.LabelFrame(parent, text=f"  {titulo}  ",
                      bg=BG2, fg=AZUL, bd=1, relief="groove",
                      font=("Courier New", 8, "bold"), labelanchor="nw")
    f.pack(fill="x", padx=10, pady=4)
    return f

class Tooltip:
    """Tooltip que aparece ao passar o mouse sobre um widget."""
    def __init__(self, widget, texto, delay=500):
        self._widget = widget
        self._texto  = texto
        self._delay  = delay
        self._id     = None
        self._win    = None
        widget.bind("<Enter>",  self._agendar)
        widget.bind("<Leave>",  self._cancelar)
        widget.bind("<Button>", self._cancelar)

    def _agendar(self, _=None):
        self._cancelar()
        self._id = self._widget.after(self._delay, self._mostrar)

    def _cancelar(self, _=None):
        if self._id:
            self._widget.after_cancel(self._id)
            self._id = None
        if self._win:
            self._win.destroy()
            self._win = None

    def _mostrar(self):
        x = self._widget.winfo_rootx() + 20
        y = self._widget.winfo_rooty() + self._widget.winfo_height() + 4
        self._win = tw = tk.Toplevel(self._widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tk.Label(tw, text=self._texto, justify="left",
                 bg="#fffde7", fg="#333333",
                 font=("Segoe UI", 9), relief="solid", bd=1,
                 wraplength=340, padx=8, pady=6).pack()


def tip(widget, texto):
    """Atalho: adiciona tooltip a um widget e retorna o widget."""
    Tooltip(widget, texto)
    return widget


def help_btn(parent, titulo, texto):
    """Botão ? que abre popup com explicação."""
    def _abrir():
        win = tk.Toplevel()
        win.title(titulo)
        win.configure(bg=BG2)
        win.resizable(False, False)
        tk.Label(win, text=titulo, bg=BG2, fg=AZUL,
                 font=("Courier New", 10, "bold")).pack(padx=16, pady=(12, 4))
        tk.Label(win, text=texto, bg=BG2, fg=TEXTO,
                 font=("Segoe UI", 9), justify="left",
                 wraplength=420, padx=16, pady=8).pack()
        tk.Button(win, text="Fechar", command=win.destroy,
                  bg=AZUL, fg="white", relief="flat",
                  font=("Segoe UI", 9), padx=12, pady=4,
                  cursor="hand2").pack(pady=(0, 12))
        win.grab_set()
    b = tk.Button(parent, text=" ? ", command=_abrir,
                  bg=BG3, fg=AZUL, relief="flat",
                  font=("Courier New", 8, "bold"),
                  cursor="hand2", padx=2, pady=0,
                  highlightthickness=0)
    return b


def scroll_frame(parent):
    """Frame com scrollbar vertical."""
    outer = tk.Frame(parent, bg=BG)
    canvas = tk.Canvas(outer, bg=BG, highlightthickness=0)
    sb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
    inner = tk.Frame(canvas, bg=BG)
    inner.bind("<Configure>", lambda e: canvas.configure(
        scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=inner, anchor="nw")
    canvas.configure(yscrollcommand=sb.set)
    canvas.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")
    outer.pack(fill="both", expand=True)
    # Mousewheel
    def _scroll(e):
        canvas.yview_scroll(int(-1*(e.delta/120)), "units")
    canvas.bind_all("<MouseWheel>", _scroll)
    return inner

# ── Gerador de config (mesma lógica do NT_app.py Streamlit) ───────────────────
def gerar_config_python(cfg: dict) -> str:
    def py(v):
        if v is None:           return "None"
        if isinstance(v, bool): return str(v)
        if isinstance(v, str):  return repr(v)
        return repr(v)

    seq = cfg.get("sequencia", [])
    if seq:
        seq_lines = "SEQUENCIA = [\n"
        for item in seq:
            ativas = item.get("curvas_ativas", {})
            seq_lines += "    {\n"
            seq_lines += f"        'planilha': {repr(item['planilha'])},\n"
            seq_lines += f"        'curvas_ativas': {ativas},\n"
            seq_lines += "    },\n"
        seq_lines += "]"
    else:
        seq_lines = "SEQUENCIA = []"

    return f"""
SCORE_MINIMO = {py(cfg['score_minimo'])}
DESVIO_ALVO  = {py(cfg['desvio_alvo'])}
ITERACOES_BETA  = {py(cfg['iteracoes_beta'])}
ITERACOES_PESOS = {py(cfg['iteracoes_pesos'])}
SEMENTE = {py(cfg['semente'])}
N_PROCESSOS = {py(cfg['n_processos'])}
ARQUIVO_PARAR = "PARAR.txt"

INTERVALOS_BETA = {{
    "beta_exp": ({py(cfg['beta_exp_lo'])}, {py(cfg['beta_exp_hi'])}),
    "beta_g3":  ({py(cfg['beta_g1_lo'])}, {py(cfg['beta_g1_hi'])}),
    "beta_g4":  ({py(cfg['beta_g2_lo'])}, {py(cfg['beta_g2_hi'])}),
    "beta_g5":  ({py(cfg['beta_g3_lo'])}, {py(cfg['beta_g3_hi'])}),
}}

CURVAS_PESO_ATIVAS = {{
    "p_mb":  True,
    "p_exp": {py(cfg['ativa_exp'])},
    "p_g3":  {py(cfg['ativa_g1'])},
    "p_g4":  {py(cfg['ativa_g2'])},
    "p_g5":  {py(cfg['ativa_g3'])},
}}

ENERGIAS_VALOR = {{
    "e_g3": {py(cfg['e_g1_valor'])},
    "e_g4": {py(cfg['e_g2_valor'])},
    "e_g5": {py(cfg['e_g3_valor'])},
}}

ENERGIA_NM_LIVRE = {{
    "e_g3": {py(cfg['e_g1_nm_livre'])},
    "e_g4": {py(cfg['e_g2_nm_livre'])},
    "e_g5": {py(cfg['e_g3_nm_livre'])},
}}

INTERVALOS_ENERGIA = {{
    "e_g3": (0.1, 6.0), "e_g4": (0.1, 6.0), "e_g5": (0.1, 6.0),
}}

PESO_MIN = {{
    "p_mb":  {py(cfg['peso_min_mb'])},
    "p_exp": None, "p_g3": None, "p_g4": None, "p_g5": None,
}}

SALVAR_GRAFICOS_TOP = {py(cfg['salvar_graficos_top'])}
VIEWER_TOP          = 50

USAR_NELDER_MEAD   = {py(cfg['usar_nm'])}
NM_TOP_CANDIDATOS  = {py(cfg['nm_top_candidatos'])}
NM_MAX_ITER        = {py(cfg['nm_max_iter'])}
NM_TOLERANCIA      = {py(cfg['nm_tolerancia'])}
NM_MAX_REINICIAR   = {py(cfg['nm_max_reiniciar'])}
NM_PERTURB_ESCALA  = {py(cfg['nm_perturb_escala'])}

NM_USAR_CLUSTERS   = {py(cfg['nm_usar_clusters'])}
NM_CLUSTER_LIMIAR  = {py(cfg['nm_cluster_limiar'])}
NM_CLUSTERS_MAX    = {py(cfg['nm_clusters_max'])}

BUSCA_DUAS_FASES   = False
ADAPTAR_INTERVALOS = False
MODO_VALIDACAO     = False
VALIDACAO_REFINAR_NM = False

{seq_lines}

EXCEL_PATH          = {repr(cfg['excel_path'])}
RESULTADOS_DIR      = {repr(cfg['resultados_dir'])}
"""

def criar_script_temporario(script_path, cfg):
    with open(script_path, "r", encoding="utf-8") as f:
        codigo = f.read()
    m2s = re.search(r'#\s*={10,}\s*\n#\s*SEÇÃO 2', codigo)
    m3s = re.search(r'#\s*={10,}\s*\n#\s*SEÇÃO 3', codigo)
    if not m2s or not m3s:
        raise ValueError("Marcadores SEÇÃO 2 / SEÇÃO 3 não encontrados no script.")
    novo_cfg = gerar_config_python(cfg)
    novo_codigo = codigo[:m2s.start()] + novo_cfg + "\n" + codigo[m3s.start():]
    fd, tmp = tempfile.mkstemp(prefix="tof_run_", suffix=".py")
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        f.write(novo_codigo)
    return tmp


# ══════════════════════════════════════════════════════════════════════════════
# APLICAÇÃO PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════
class NTApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Natalia Time")
        self.geometry("960x700")
        self.configure(bg=BG)
        self.minsize(800, 600)

        self.rodando     = False
        self.proc_ref    = None
        self.ultima_pasta = None
        self.sequencia   = []

        self._build_style()
        self._build_ui()

    def _build_style(self):
        s = ttk.Style(self)
        s.theme_use("default")
        s.configure("TNotebook",      background=BG,  borderwidth=0)
        s.configure("TNotebook.Tab",  background=BG2, foreground=TEXTO2,
                    padding=[14, 6],  font=FONTE_M)
        s.map("TNotebook.Tab",
              background=[("selected", BG)],
              foreground=[("selected", AZUL)])
        s.configure("TFrame",         background=BG)
        s.configure("green.Horizontal.TProgressbar",
                    troughcolor=BG2, background=AZUL, thickness=14)

    def _build_ui(self):
        # Cabeçalho
        hdr = tk.Frame(self, bg=BG2, pady=8)
        hdr.pack(fill="x")
        tk.Label(hdr, text="⚗  Natalia Time", font=("Courier New", 14, "bold"),
                 bg=BG2, fg=AZUL).pack(side="left", padx=16)
        tk.Label(hdr, text="Análise de Espectroscopia DETOF",
                 font=FONTE, bg=BG2, fg=TEXTO2).pack(side="left")

        # Sidebar (planilha + aba + linha)
        self._build_sidebar()

        # Abas
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=0, pady=0)
        self._aba_inicio(nb)
        self._aba_config(nb)
        self._aba_sequencia(nb)
        self._aba_executar(nb)
        self._aba_resultados(nb)
        self._aba_historico(nb)

    # ── Sidebar ───────────────────────────────────────────────────────────────
    def _build_sidebar(self):
        sb = tk.Frame(self, bg=BG2, bd=0)
        sb.pack(fill="x", padx=10, pady=(4,0))

        tk.Label(sb, text="Experimento:", bg=BG2, fg=TEXTO2, font=FONTE).grid(
            row=0, column=0, sticky="w", padx=(4,4))
        self.entry_planilha = entry(sb, width=55)
        self.entry_planilha.grid(row=0, column=1, padx=(0,4), pady=4)
        btn(sb, "Procurar", self._procurar_planilha).grid(row=0, column=2, padx=(0,8))

        # Formato detectado automaticamente — exibido como informação
        self.lbl_formato = tk.Label(sb, text="", bg=BG2, fg=TEXTO2, font=FONTE)
        self.lbl_formato.grid(row=0, column=3, padx=(8,4))

        # Aviso pasta sincronizada
        _pastas_sync = ["dropbox", "onedrive", "google drive", "googledrive", "icloud"]
        if any(p in _base.lower().replace("\\","/") for p in _pastas_sync):
            tk.Label(sb, text="⚠ Pasta sincronizada detectada — pause a sincronização antes de rodar",
                     bg=BG2, fg=AMAR, font=FONTE).grid(
                row=1, column=0, columnspan=7, sticky="w", padx=4, pady=2)

    # ── Aba Início ────────────────────────────────────────────────────────────
    def _aba_inicio(self, nb):
        outer = tk.Frame(nb, bg=BG)
        nb.add(outer, text="🏠  Início")

        canvas = tk.Canvas(outer, bg=BG, highlightthickness=0)
        sb2 = tk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg=BG)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=sb2.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb2.pack(side="right", fill="y")
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(-1*(e.delta//120), "units"))

        # Título
        tk.Label(inner, text="Natalia Time",
                 bg=BG, fg=AZUL, font=("Segoe UI", 18, "bold")).pack(pady=(30, 2))
        tk.Label(inner, text="Ajuste automático de espectros DETOF de fotofragmentação",
                 bg=BG, fg=TEXTO2, font=("Segoe UI", 10)).pack(pady=(0, 24))

        # ── Como começar ──────────────────────────────────────────────────────
        tk.Label(inner, text="Como começar",
                 bg=BG, fg=TEXTO, font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=60, pady=(0, 6))

        frame_passos = tk.Frame(inner, bg=BG)
        frame_passos.pack(padx=60, fill="x")

        passos = [
            ("📂", "Selecione o experimento",
                    "Clique em Procurar na barra superior e escolha sua planilha (.xlsm / .xlsx) "
                    "ou o arquivo dados.csv da pasta do experimento (novo formato)."),
            ("⚙",  "Configure as curvas",
                    "Na aba Configuração, verifique quais curvas estão ativas (MB, Exponencial, "
                    "Gaussianas 1/2/3) e ajuste os intervalos de beta. "
                    "No formato CSV, as energias precisam estar preenchidas ou gravadas."),
            ("▶",  "Execute a análise",
                    "Na aba Executar, clique em Iniciar. O log mostrará o progresso em tempo real. "
                    "Você pode parar a qualquer momento — tudo que foi encontrado é salvo."),
            ("📊", "Veja os resultados",
                    "Ao terminar, os melhores ajustes aparecem na aba Resultados. "
                    "Gráficos e o relatório Excel são salvos automaticamente na pasta Resultados/."),
        ]

        for icon, titulo, desc in passos:
            card = tk.Frame(frame_passos, bg=BG2, padx=14, pady=10)
            card.pack(fill="x", pady=3)
            tk.Label(card, text=icon, bg=BG2, fg=AZUL,
                     font=("Segoe UI", 14)).pack(side="left", padx=(0, 12), anchor="n")
            col = tk.Frame(card, bg=BG2)
            col.pack(side="left", fill="x", expand=True)
            tk.Label(col, text=titulo, bg=BG2, fg=TEXTO,
                     font=("Segoe UI", 9, "bold"), anchor="w").pack(anchor="w")
            tk.Label(col, text=desc, bg=BG2, fg=TEXTO2,
                     font=FONTE, anchor="w", justify="left", wraplength=580).pack(anchor="w", pady=(2, 0))

        # ── Formatos aceitos ──────────────────────────────────────────────────
        tk.Label(inner, text="Formatos de entrada aceitos",
                 bg=BG, fg=TEXTO, font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=60, pady=(20, 6))

        frame_fmt = tk.Frame(inner, bg=BG)
        frame_fmt.pack(padx=60, fill="x")

        for icon, fmt, desc in [
            ("📊", "Excel (.xlsm / .xlsx)",
                    "Formato original. A aba e a linha de início dos dados são detectadas "
                    "automaticamente — basta selecionar o arquivo."),
            ("📄", "CSV + TOML (novo formato)",
                    "Pasta com dados.csv e parametros.toml. Criada com o programa Migrador. "
                    "Não requer o Excel instalado na máquina."),
        ]:
            r = tk.Frame(frame_fmt, bg=BG2, padx=14, pady=8)
            r.pack(fill="x", pady=3)
            tk.Label(r, text=f"{icon}  {fmt}", bg=BG2, fg=AZUL,
                     font=("Segoe UI", 9, "bold")).pack(anchor="w")
            tk.Label(r, text=desc, bg=BG2, fg=TEXTO2,
                     font=FONTE, wraplength=580, justify="left").pack(anchor="w", pady=(2, 0))

        tk.Label(inner, text="", bg=BG).pack(pady=16)

    # ── Aba Configuração ──────────────────────────────────────────────────────
    def _aba_config(self, nb):
        outer = ttk.Frame(nb)
        nb.add(outer, text="⚙  Configuração")
        inner = scroll_frame(outer)

        left  = tk.Frame(inner, bg=BG); left.pack(side="left",  fill="both", expand=True)
        right = tk.Frame(inner, bg=BG); right.pack(side="right", fill="both", expand=True)

        # ── Busca aleatória
        s = secao(left, "🔍 Busca aleatória")
        s["bg"] = BG2
        help_btn(s, "Busca aleatória",
            "O programa sorteia combinações aleatórias de parâmetros e avalia o ajuste.\n\n"
            "• Iterações beta: quantas combinações de beta (largura das curvas) sortear.\n"
            "• Iterações pesos: para cada beta, quantas combinações de pesos testar.\n"
            "  Total de cálculos = Iterações beta × Iterações pesos.\n\n"
            "• R² mínimo: limiar de qualidade — só resultados acima são considerados válidos.\n"
            "  (R² = 1.0 seria ajuste perfeito; tente ≥ 0.99)\n\n"
            "• RMSE alvo: erro quadrático médio desejado. A busca para ao atingi-lo.\n\n"
            "• Semente: número fixo = resultados reproduzíveis. Vazio = aleatório a cada run.\n"
            "• Núcleos CPU: quantos núcleos usar em paralelo. Mais núcleos = mais rápido."
        ).pack(anchor="e", padx=4, pady=2)
        r = tk.Frame(s, bg=BG2); r.pack(fill="x", padx=6, pady=4)
        tip(lbl(r, "Iterações beta"), "Quantas combinações de β (largura das curvas) sortear.\nTotal de cálculos = beta × pesos.").grid(row=0, column=0, sticky="w")
        self.e_iter_beta = entry(r, 8, "1500"); tip(self.e_iter_beta, "Padrão: 1500. Aumentar melhora a busca mas demora mais."); self.e_iter_beta.grid(row=1, column=0, padx=(0,10))
        tip(lbl(r, "Iterações pesos"), "Para cada β, quantas combinações de pesos testar.\nTotal de cálculos = beta × pesos.").grid(row=0, column=1, sticky="w")
        self.e_iter_pesos = entry(r, 8, "1500"); tip(self.e_iter_pesos, "Padrão: 1500. Total = 1500 × 1500 = 2.25 M cálculos."); self.e_iter_pesos.grid(row=1, column=1, padx=(0,10))
        tip(lbl(r, "R² mínimo"), "Limiar de qualidade. Resultados abaixo são descartados.\nR²=1 seria ajuste perfeito. Recomendado: ≥ 0.99").grid(row=0, column=2, sticky="w")
        self.e_score = entry(r, 7, "0.99"); tip(self.e_score, "Padrão: 0.99\nAbaixe para aceitar mais resultados (menos exigente)."); self.e_score.grid(row=1, column=2, padx=(0,10))
        tip(lbl(r, "RMSE alvo"), "Erro quadrático médio desejado.\nA busca para automaticamente ao atingir este valor.").grid(row=0, column=3, sticky="w")
        self.e_rmse = entry(r, 7, "0.04"); tip(self.e_rmse, "Padrão: 0.04\nAbaixe para exigir ajuste mais preciso."); self.e_rmse.grid(row=1, column=3)

        r2 = tk.Frame(s, bg=BG2); r2.pack(fill="x", padx=6, pady=(0,6))
        tip(lbl(r2, "Semente (vazio = aleatório)"),
            "Número inteiro para reproduzir exatamente os mesmos resultados.\n"
            "Deixe vazio para sortear diferente a cada execução.").grid(row=0, column=0, sticky="w")
        self.e_semente = entry(r2, 10, ""); tip(self.e_semente, "Ex: 42  →  sempre os mesmos resultados."); self.e_semente.grid(row=1, column=0, padx=(0,10))

        cpu = os.cpu_count() or 1
        tip(lbl(r2, f"Núcleos CPU (total: {cpu})"),
            "Quantos núcleos usar em paralelo.\n"
            "Posição máxima = usa todos os núcleos disponíveis.\n"
            "Reduza se quiser deixar o computador responsivo.").grid(row=0, column=1, sticky="w")
        self.slider_cpu = tk.Scale(r2, from_=1, to=cpu, orient="horizontal",
                                   bg=BG2, fg=TEXTO, troughcolor=BORDA,
                                   highlightthickness=0, activebackground=AZUL,
                                   font=("Segoe UI", 8), length=160)
        self.slider_cpu.set(cpu); self.slider_cpu.grid(row=1, column=1)

        # ── Nelder-Mead
        s2 = secao(left, "🎯 Nelder-Mead")
        s2["bg"] = BG2
        help_btn(s2, "Refinamento Nelder-Mead",
            "Após a busca aleatória encontrar bons candidatos, o Nelder-Mead refina cada um "
            "procurando o mínimo local mais preciso — sem sortear, apenas movendo os parâmetros "
            "em direção ao melhor ajuste.\n\n"
            "• Max iter/candidato: limite de passos de refinamento por candidato.\n"
            "• Máx reinícios: se o NM travar, ele perturba os parâmetros e tenta de novo.\n"
            "• Perturbação: intensidade da perturbação no reinício (0.10 = ±10%).\n"
            "• Tolerância: critério de convergência (10⁻⁹ é muito preciso; valores maiores = mais rápido).\n\n"
            "• Clusters: agrupa os candidatos por similaridade antes de refinar, garantindo "
            "que o NM explore regiões diversas do espaço de parâmetros.\n"
            "• Limiar de cluster: quanto menor, mais grupos distintos (mais diversidade).\n"
            "• Máx clusters: limite de grupos a explorar.\n"
            "• Top candidatos: usado quando clusters estão desativados — refina os N melhores."
        ).pack(anchor="e", padx=4, pady=2)
        cb_nm, self.v_usar_nm = check(s2, "Ativar refinamento NM", True)
        tip(cb_nm, "Recomendado: ativado.\nMelhora significativamente a qualidade do resultado final.")
        cb_nm.pack(anchor="w", padx=6, pady=(4,2))

        r3 = tk.Frame(s2, bg=BG2); r3.pack(fill="x", padx=6, pady=4)
        tip(lbl(r3, "Max iter/candidato"), "Limite de passos de refinamento por candidato.\nPadrão: 2000.").grid(row=0, column=0, sticky="w")
        self.e_nm_iter = entry(r3, 7, "2000"); tip(self.e_nm_iter, "Padrão: 2000. Aumentar pode melhorar mas demora mais."); self.e_nm_iter.grid(row=1, column=0, padx=(0,10))
        tip(lbl(r3, "Máx reinícios"), "Se o NM convergir para um mínimo ruim,\nele perturba e tenta novamente até este limite.").grid(row=0, column=1, sticky="w")
        self.e_nm_reinic = entry(r3, 5, "3"); tip(self.e_nm_reinic, "Padrão: 3."); self.e_nm_reinic.grid(row=1, column=1, padx=(0,10))
        tip(lbl(r3, "Perturbação reinício"), "Intensidade da perturbação ao reiniciar.\n0.10 = ±10% nos parâmetros atuais.").grid(row=0, column=2, sticky="w")
        self.e_nm_perturb = entry(r3, 6, "0.10"); tip(self.e_nm_perturb, "Padrão: 0.10\nAumente se o NM ficar preso no mesmo mínimo."); self.e_nm_perturb.grid(row=1, column=2)

        r4 = tk.Frame(s2, bg=BG2); r4.pack(fill="x", padx=6, pady=(0,4))
        tip(lbl(r4, "Tolerância (10^x)  [-15 a -3]"),
            "Critério de convergência do NM.\n"
            "−9 = muito preciso (padrão). −3 = mais rápido, menos preciso.").grid(row=0, column=0, sticky="w")
        self.slider_nm_tol = tk.Scale(r4, from_=-15, to=-3, orient="horizontal",
                                      bg=BG2, fg=TEXTO, troughcolor=BORDA,
                                      highlightthickness=0, activebackground=AZUL,
                                      font=("Segoe UI", 8), length=160)
        self.slider_nm_tol.set(-9); tip(self.slider_nm_tol, "Padrão: −9  (tolerância = 10⁻⁹)"); self.slider_nm_tol.grid(row=1, column=0, padx=(0,20))

        cb_cl, self.v_nm_clusters = check(r4, "Usar clusters", True)
        tip(cb_cl, "Agrupa candidatos similares antes de refinar.\nGarante que o NM explore regiões diversas do espaço de parâmetros.")
        cb_cl.grid(row=0, column=1, sticky="w")
        tip(lbl(r4, "Limiar cluster"), "Quanto menor, mais grupos distintos.\n0.25 = boa diversidade (padrão).").grid(row=0, column=2, sticky="w", padx=(10,0))
        self.e_nm_limiar = entry(r4, 5, "0.25"); tip(self.e_nm_limiar, "Padrão: 0.25\nReduzir cria mais clusters (mais diversidade)."); self.e_nm_limiar.grid(row=1, column=2, padx=(10,10))
        tip(lbl(r4, "Máx clusters"), "Limite de grupos a explorar com NM.").grid(row=0, column=3, sticky="w")
        self.e_nm_max_cl = entry(r4, 4, "10"); tip(self.e_nm_max_cl, "Padrão: 10."); self.e_nm_max_cl.grid(row=1, column=3, padx=(0,10))
        tip(lbl(r4, "Top candidatos"), "Usado apenas quando clusters estão desativados.\nQuantos dos melhores candidatos refinar.").grid(row=0, column=4, sticky="w")
        self.e_nm_top = entry(r4, 4, "5"); tip(self.e_nm_top, "Padrão: 5. Só usado se 'Usar clusters' estiver desmarcado."); self.e_nm_top.grid(row=1, column=4)

        # ── Curvas ativas
        s3 = secao(right, "📈 Curvas ativas")
        s3["bg"] = BG2
        help_btn(s3, "Curvas ativas",
            "O espectro DETOF é decomposto numa soma de curvas, cada uma representando "
            "uma contribuição física diferente:\n\n"
            "• Maxwell-Boltzmann (MB): distribuição de moléculas do gás de fundo. Sempre ativa.\n"
            "• Exponencial: contribuição de fragmentos com distribuição exponencial de energia.\n"
            "• Gaussiana 1, 2, 3: fragmentos com distribuições gaussianas de energia "
            "(picos em energias específicas, configuradas na seção Energias).\n\n"
            "• Peso mínimo MB: garante que a curva MB contribua com pelo menos este valor. "
            "Evita que o ajuste descarte fisicamente a contribuição do gás de fundo."
        ).pack(anchor="e", padx=4, pady=2)
        cf = tk.Frame(s3, bg=BG2); cf.pack(fill="x", padx=6, pady=4)
        _, _ = check(cf, "Maxwell-Boltzmann", True, "disabled"); _[0].grid if False else None
        cb_mb, _ = check(cf, "Maxwell-Boltzmann", True, "disabled")
        tip(cb_mb, "Sempre ativa — representa o gás de fundo (distribuição Maxwell-Boltzmann).")
        cb_mb.grid(row=0, column=0, padx=6)
        cb_exp, self.v_ativa_exp = check(cf, "Exponencial", True)
        tip(cb_exp, "Contribuição exponencial — fragmentos com distribuição de energia decaindo exponencialmente.")
        cb_exp.grid(row=0, column=1, padx=6)
        cb_g1, self.v_ativa_g1 = check(cf, "Gaussiana 1", True)
        tip(cb_g1, "Primeira gaussiana — pico de fragmentos em torno de uma energia específica (E_g1).")
        cb_g1.grid(row=0, column=2, padx=6)
        cb_g2, self.v_ativa_g2 = check(cf, "Gaussiana 2", True)
        tip(cb_g2, "Segunda gaussiana — segundo pico de fragmentos (E_g2).")
        cb_g2.grid(row=0, column=3, padx=6)
        cb_g3, self.v_ativa_g3 = check(cf, "Gaussiana 3", False)
        tip(cb_g3, "Terceira gaussiana — opcional. Ative se o espectro mostrar um terceiro pico (E_g3).")
        cb_g3.grid(row=0, column=4, padx=6)

        r5 = tk.Frame(s3, bg=BG2); r5.pack(fill="x", padx=6, pady=(0,6))
        tip(lbl(r5, "Peso mínimo MB"), "Garante que a MB contribua com pelo menos este peso no ajuste.\nEvita soluções fisicamente incorretas que ignoram o gás de fundo.").pack(side="left")
        self.e_peso_mb = entry(r5, 7, "0.01"); tip(self.e_peso_mb, "Padrão: 0.01  (1% mínimo para MB)."); self.e_peso_mb.pack(side="left", padx=8)

        # ── Energias
        s4 = secao(right, "⚡ Energias das gaussianas")
        s4["bg"] = BG2
        help_btn(s4, "Energias das gaussianas",
            "Cada gaussiana é centrada numa energia de translação específica (em eV), "
            "que define a posição do pico no espectro DETOF.\n\n"
            "• Vazio: usa o valor lido diretamente da planilha Excel (células fixas).\n"
            "• Valor numérico: sobrescreve o valor da planilha para esta execução.\n\n"
            "• NM livre: permite que o Nelder-Mead ajuste a energia durante o refinamento, "
            "dentro do intervalo [0.1, 6.0 eV]. Útil quando a energia não é bem conhecida a priori.\n\n"
            "Exemplo: se E_g1 = 0.7 eV, a Gaussiana 1 representa fragmentos com ~0.7 eV de energia "
            "de translação no centro de massa."
        ).pack(anchor="e", padx=4, pady=2)
        self.e_energias  = {}
        self.v_nm_livre  = {}
        self.v_gravar_e  = {}   # checkboxes Gravar (só CSV)
        self.cb_gravar_e = {}   # referência aos widgets para mostrar/ocultar
        for i, (g, nome) in enumerate([("g1","Gaussiana 1"), ("g2","Gaussiana 2"), ("g3","Gaussiana 3")]):
            r = tk.Frame(s4, bg=BG2); r.pack(fill="x", padx=6, pady=2)
            tip(lbl(r, f"E_{g} (eV)"),
                f"Energia de translação da {nome} em eV.\n"
                "Excel: vazio = lê da planilha. Número = sobrescreve.\n"
                "CSV: vazio = usa valor gravado no parametros.toml.").pack(side="left")
            e = entry(r, 8, ""); tip(e, f"Ex: 0.7  →  {nome} centrada em 0.7 eV."); e.pack(side="left", padx=(6,10))
            self.e_energias[g] = e
            cb_nm, v_nm = check(r, "NM livre", False)
            tip(cb_nm, f"Permite ao Nelder-Mead ajustar E_{g} durante o refinamento\n(intervalo: 0.1 – 6.0 eV).")
            cb_nm.pack(side="left"); self.v_nm_livre[g] = v_nm
            # Checkbox Gravar — só visível no formato CSV
            v_gr = tk.BooleanVar(value=False)
            cb_gr = tk.Checkbutton(r, text="Gravar", variable=v_gr,
                                   bg=BG2, fg=TEXTO2, selectcolor=BG3,
                                   activebackground=BG2, activeforeground=TEXTO,
                                   font=FONTE)
            tip(cb_gr, f"Grava o valor de E_{g} no parametros.toml da pasta do experimento.\n"
                       "Na próxima execução, campo vazio usará este valor gravado.")
            cb_gr.pack(side="left", padx=(8,0))
            cb_gr.pack_forget()   # oculto por padrão
            self.v_gravar_e[g] = v_gr
            self.cb_gravar_e[g] = cb_gr

        # ── Intervalos de beta
        s5 = secao(right, "📐 Intervalos de beta")
        s5["bg"] = BG2
        help_btn(s5, "Intervalos de beta",
            "Beta (β) controla a largura de cada curva no espectro DETOF — quanto maior o β, "
            "mais estreita a distribuição de velocidades.\n\n"
            "Estes intervalos definem os limites do sorteio aleatório na busca:\n"
            "• Min: valor mínimo de β a sortear\n"
            "• Max: valor máximo de β a sortear\n\n"
            "Os valores padrão cobrem bem a maioria dos experimentos, mas se você tiver "
            "estimativas prévias pode restringir o intervalo para acelerar a busca."
        ).pack(anchor="e", padx=4, pady=2)
        self.e_beta_lo = {}; self.e_beta_hi = {}
        labels = {"exp": "Exponencial", "g1": "Gaussiana 1", "g2": "Gaussiana 2", "g3": "Gaussiana 3"}
        defaults = {"exp": (50,600), "g1": (100,1200), "g2": (100,1200), "g3": (100,1200)}
        for k, (lo, hi) in defaults.items():
            r = tk.Frame(s5, bg=BG2); r.pack(fill="x", padx=6, pady=2)
            tip(lbl(r, f"beta_{k}  min"), f"Valor mínimo de β para {labels[k]}.\nβ controla a largura da curva.").pack(side="left")
            e_lo = entry(r, 6, str(lo)); tip(e_lo, f"Padrão: {lo}"); e_lo.pack(side="left", padx=(4,10))
            tip(lbl(r, "max"), f"Valor máximo de β para {labels[k]}.").pack(side="left")
            e_hi = entry(r, 6, str(hi)); tip(e_hi, f"Padrão: {hi}"); e_hi.pack(side="left", padx=4)
            self.e_beta_lo[k] = e_lo; self.e_beta_hi[k] = e_hi

        # ── Saídas
        s6 = secao(right, "💾 Saídas")
        s6["bg"] = BG2
        r6 = tk.Frame(s6, bg=BG2); r6.pack(fill="x", padx=6, pady=4)
        tip(lbl(r6, "Gráficos PNG a salvar (vazio = todos)"),
            "Quantos gráficos PNG salvar, ordenados por qualidade.\n"
            "20 = salva os 20 melhores ajustes. Vazio = salva todos os válidos.").pack(side="left")
        self.e_salvar_top = entry(r6, 5, "20"); tip(self.e_salvar_top, "Padrão: 20."); self.e_salvar_top.pack(side="left", padx=8)

    # ── Aba Sequência ─────────────────────────────────────────────────────────
    def _aba_sequencia(self, nb):
        outer = ttk.Frame(nb)
        nb.add(outer, text="📋  Sequência")
        inner = scroll_frame(outer)

        s = secao(inner, "Modo sequência")
        s["bg"] = BG2
        help_btn(s, "Modo sequência",
            "Processa múltiplas planilhas automaticamente, uma após a outra, sem intervenção.\n\n"
            "Útil quando você tem várias medições (diferentes moléculas, condições ou repetições) "
            "e quer analisar todas com as mesmas configurações.\n\n"
            "• Adicione as planilhas na ordem desejada.\n"
            "• Você pode escolher quais curvas ativar para cada planilha individualmente "
            "(ex: uma medição pode precisar de 3 gaussianas, outra só de 2).\n"
            "• Quando o modo sequência está ativo, a planilha da barra superior é ignorada.\n"
            "• Os resultados de cada planilha são salvos em pastas separadas dentro de Resultados/."
        ).pack(anchor="e", padx=4, pady=2)
        cb_seq, self.v_usar_seq = check(s, "Ativar modo sequência", False)
        tip(cb_seq, "Quando ativado, processa todas as planilhas da fila em ordem automática.")
        cb_seq.pack(anchor="w", padx=6, pady=4)

        # Adicionar planilha
        s2 = secao(inner, "➕ Adicionar planilha")
        s2["bg"] = BG2
        r = tk.Frame(s2, bg=BG2); r.pack(fill="x", padx=6, pady=4)
        lbl(r, "Caminho:").pack(side="left")
        self.entry_sq_path = entry(r, 45); self.entry_sq_path.pack(side="left", padx=6)
        btn(r, "Procurar", self._procurar_seq).pack(side="left")

        r2 = tk.Frame(s2, bg=BG2); r2.pack(fill="x", padx=6, pady=(0,4))
        lbl(r2, "Curvas:").pack(side="left", padx=(0,6))
        _, self.v_sq_exp = check(r2, "EXP", True); _.pack(side="left", padx=4)
        _, self.v_sq_g1  = check(r2, "G1",  True); _.pack(side="left", padx=4)
        _, self.v_sq_g2  = check(r2, "G2",  True); _.pack(side="left", padx=4)
        _, self.v_sq_g3  = check(r2, "G3",  False); _.pack(side="left", padx=4)
        btn(r2, "Adicionar", self._adicionar_seq, cor=AZUL, fg="white").pack(side="left", padx=12)

        # Lista
        self.frame_seq_lista = secao(inner, "Fila")
        self.frame_seq_lista["bg"] = BG2
        self._atualizar_lista_seq()

    # ── Aba Executar ──────────────────────────────────────────────────────────
    def _aba_executar(self, nb):
        outer = ttk.Frame(nb)
        nb.add(outer, text="▶  Executar")

        # Botões
        bf = tk.Frame(outer, bg=BG, pady=10)
        bf.pack(fill="x", padx=16)
        self.btn_iniciar = btn(bf, "▶  Iniciar análise", self._iniciar,
                               cor=AZUL, fg="white")
        self.btn_iniciar.pack(side="left", padx=(0,8))
        self.btn_parar = btn(bf, "⏹  Parar", self._parar, cor=VERM, fg="white")
        self.btn_parar.pack(side="left")
        self.btn_parar.config(state="disabled")
        self.lbl_status = tk.Label(bf, text="", bg=BG, fg=TEXTO2, font=FONTE)
        self.lbl_status.pack(side="left", padx=12)

        # Barra de progresso
        pf = tk.Frame(outer, bg=BG)
        pf.pack(fill="x", padx=16, pady=(0,6))
        self.progressbar = ttk.Progressbar(pf, style="green.Horizontal.TProgressbar",
                                           mode="determinate", maximum=100)
        self.progressbar.pack(fill="x")
        self.lbl_pct = tk.Label(pf, text="", bg=BG, fg=TEXTO2, font=FONTE)
        self.lbl_pct.pack(anchor="e")

        # Log (terminal compacto)
        lf = tk.LabelFrame(outer, text="  Output  ", bg=BG, fg=AZUL,
                           font=("Courier New", 8, "bold"), bd=1, relief="groove")
        lf.pack(fill="both", expand=True, padx=16, pady=(0,12))
        self.terminal = tk.Text(lf, bg="#0d1117", fg="#c9d1d9",
                                font=("Courier New", 8), relief="flat",
                                state="disabled", wrap="none")
        sb = ttk.Scrollbar(lf, command=self.terminal.yview)
        self.terminal.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.terminal.pack(fill="both", expand=True, padx=4, pady=4)

    # ── Aba Resultados ────────────────────────────────────────────────────────
    def _aba_resultados(self, nb):
        outer = ttk.Frame(nb)
        nb.add(outer, text="📊  Resultados")

        bf = tk.Frame(outer, bg=BG, pady=8)
        bf.pack(fill="x", padx=16)
        btn(bf, "🔄 Atualizar", self._atualizar_resultados).pack(side="left")

        self.frame_res = tk.Frame(outer, bg=BG)
        self.frame_res.pack(fill="both", expand=True, padx=16)
        self.lbl_res_info = tk.Label(self.frame_res,
                                     text="Rode uma análise para ver os resultados.",
                                     bg=BG, fg=TEXTO2, font=FONTE)
        self.lbl_res_info.pack(pady=30)

    # ── Aba Histórico ─────────────────────────────────────────────────────────
    def _aba_historico(self, nb):
        outer = ttk.Frame(nb)
        nb.add(outer, text="🕒  Histórico")

        bf = tk.Frame(outer, bg=BG, pady=8)
        bf.pack(fill="x", padx=16)
        btn(bf, "🔄 Atualizar histórico", self._atualizar_historico).pack(side="left")

        self.frame_hist = scroll_frame(outer)
        self._atualizar_historico()

    # ── Ações ─────────────────────────────────────────────────────────────────
    def _procurar_planilha(self):
        p = filedialog.askopenfilename(
            filetypes=[
                ("Experimento", "*.xlsm *.xlsx *.csv"),
                ("Excel", "*.xlsm *.xlsx"),
                ("CSV (novo formato)", "*.csv"),
                ("Todos", "*.*"),
            ])
        if p:
            self.entry_planilha.delete(0, "end")
            self.entry_planilha.insert(0, p)
            self._atualizar_lbl_formato(p)

    def _atualizar_lbl_formato(self, caminho):
        ext = os.path.splitext(caminho)[1].lower()
        eh_csv = (ext == ".csv")
        if eh_csv:
            self.lbl_formato.config(text="📄 CSV + TOML", fg=AZUL)
        elif ext in (".xlsm", ".xlsx"):
            self.lbl_formato.config(text="📊 Excel", fg=TEXTO2)
        else:
            self.lbl_formato.config(text="", fg=TEXTO2)
        # Mostrar checkboxes Gravar só no formato CSV
        for g in ("g1", "g2", "g3"):
            cb = self.cb_gravar_e.get(g)
            if cb:
                if eh_csv:
                    cb.pack(side="left", padx=(8,0))
                else:
                    cb.pack_forget()
                    self.v_gravar_e[g].set(False)

    def _gravar_energias_toml(self, caminho_csv):
        """Grava energias marcadas com Gravar no parametros.toml.
        Edita o arquivo como texto para preservar comentários e formatação."""
        pasta = os.path.dirname(os.path.abspath(caminho_csv))
        toml_path = os.path.join(pasta, "parametros.toml")
        mapa_g = {"g1": "e_g3", "g2": "e_g4", "g3": "e_g5"}
        nomes  = {"g1": "Gaussiana 1", "g2": "Gaussiana 2", "g3": "Gaussiana 3"}

        gravou = []
        erros  = []
        # Coletar valores a gravar
        para_gravar = {}
        for g, chave in mapa_g.items():
            if not self.v_gravar_e.get(g, tk.BooleanVar()).get():
                continue
            val_str = self.e_energias[g].get().strip()
            if not val_str:
                erros.append(f"{nomes[g]}: campo vazio")
                continue
            try:
                para_gravar[chave] = float(val_str)
                gravou.append(f"{nomes[g]} = {float(val_str)} eV")
            except ValueError:
                erros.append(f"{nomes[g]}: valor inválido '{val_str}'")

        if erros:
            messagebox.showwarning("Gravar energias",
                "Não foi possível gravar:\n" + "\n".join(erros))
        if not gravou:
            return

        try:
            try:
                with open(toml_path, "r", encoding="utf-8") as f:
                    linhas = f.readlines()
            except FileNotFoundError:
                linhas = []

            pendentes = set(para_gravar.keys())
            novas_linhas = []
            for linha in linhas:
                stripped = linha.strip().lstrip("# ").split("=")[0].strip()
                if stripped in para_gravar:
                    # Substituir linha (mesmo que estivesse comentada)
                    novas_linhas.append(f"{stripped} = {para_gravar[stripped]}\n")
                    pendentes.discard(stripped)
                else:
                    novas_linhas.append(linha)
            # Chaves que não existiam no arquivo: adicionar no final
            for chave in pendentes:
                novas_linhas.append(f"{chave} = {para_gravar[chave]}\n")

            with open(toml_path, "w", encoding="utf-8") as f:
                f.writelines(novas_linhas)

            print("Energias gravadas no parametros.toml:")
            for g in gravou:
                print(f"  {g}")
        except Exception as e:
            messagebox.showerror("Erro ao gravar",
                f"Não foi possível salvar parametros.toml:\n{e}")

    def _procurar_seq(self):
        p = filedialog.askopenfilename(
            filetypes=[("Excel", "*.xlsm *.xlsx"), ("Todos", "*.*")])
        if p:
            self.entry_sq_path.delete(0, "end")
            self.entry_sq_path.insert(0, p)

    def _adicionar_seq(self):
        p = self.entry_sq_path.get().strip()
        if not p:
            messagebox.showwarning("Aviso", "Informe o caminho da planilha.")
            return
        self.sequencia.append({
            "planilha": p,
            "curvas_ativas": {
                "p_mb": True,
                "p_exp": self.v_sq_exp.get(),
                "p_g3":  self.v_sq_g1.get(),
                "p_g4":  self.v_sq_g2.get(),
                "p_g5":  self.v_sq_g3.get(),
            }
        })
        self.entry_sq_path.delete(0, "end")
        self._atualizar_lista_seq()

    def _atualizar_lista_seq(self):
        for w in self.frame_seq_lista.winfo_children():
            w.destroy()
        if not self.sequencia:
            lbl(self.frame_seq_lista, "Nenhuma planilha na fila.").pack(padx=8, pady=4)
            return
        for i, item in enumerate(self.sequencia):
            r = tk.Frame(self.frame_seq_lista, bg=BG2)
            r.pack(fill="x", padx=6, pady=2)
            lbl(r, f"{i+1}. {os.path.basename(item['planilha'])}",
                fg=TEXTO).pack(side="left", padx=4)
            btn(r, "✕", lambda i=i: self._remover_seq(i),
                cor=VERM, fg="white").pack(side="right", padx=4)

    def _remover_seq(self, i):
        self.sequencia.pop(i)
        self._atualizar_lista_seq()

    def _coletar_cfg(self):
        def f(e): return float(e.get()) if e.get().strip() else None
        def i_v(e): return int(e.get()) if e.get().strip().isdigit() else None

        cpu_val = self.slider_cpu.get()
        n_proc  = None if cpu_val == (os.cpu_count() or 1) else cpu_val

        return dict(
            excel_path         = self.entry_planilha.get().strip(),
            resultados_dir     = RESULTADOS_DIR,
            score_minimo       = float(self.e_score.get() or 0.99),
            desvio_alvo        = float(self.e_rmse.get() or 0.04),
            iteracoes_beta     = int(self.e_iter_beta.get() or 1500),
            iteracoes_pesos    = int(self.e_iter_pesos.get() or 1500),
            semente            = i_v(self.e_semente),
            n_processos        = n_proc,
            beta_exp_lo        = float(self.e_beta_lo["exp"].get() or 50),
            beta_exp_hi        = float(self.e_beta_hi["exp"].get() or 600),
            beta_g1_lo         = float(self.e_beta_lo["g1"].get() or 100),
            beta_g1_hi         = float(self.e_beta_hi["g1"].get() or 1200),
            beta_g2_lo         = float(self.e_beta_lo["g2"].get() or 100),
            beta_g2_hi         = float(self.e_beta_hi["g2"].get() or 1200),
            beta_g3_lo         = float(self.e_beta_lo["g3"].get() or 100),
            beta_g3_hi         = float(self.e_beta_hi["g3"].get() or 1200),
            ativa_exp          = self.v_ativa_exp.get(),
            ativa_g1           = self.v_ativa_g1.get(),
            ativa_g2           = self.v_ativa_g2.get(),
            ativa_g3           = self.v_ativa_g3.get(),
            e_g1_valor         = f(self.e_energias["g1"]),
            e_g2_valor         = f(self.e_energias["g2"]),
            e_g3_valor         = f(self.e_energias["g3"]),
            e_g1_nm_livre      = self.v_nm_livre["g1"].get(),
            e_g2_nm_livre      = self.v_nm_livre["g2"].get(),
            e_g3_nm_livre      = self.v_nm_livre["g3"].get(),
            peso_min_mb        = float(self.e_peso_mb.get() or 0.01),
            salvar_graficos_top= int(self.e_salvar_top.get() or 20),
            usar_nm            = self.v_usar_nm.get(),
            nm_top_candidatos  = int(self.e_nm_top.get() or 5),
            nm_max_iter        = int(self.e_nm_iter.get() or 2000),
            nm_tolerancia      = 10 ** self.slider_nm_tol.get(),
            nm_max_reiniciar   = int(self.e_nm_reinic.get() or 3),
            nm_perturb_escala  = float(self.e_nm_perturb.get() or 0.10),
            nm_usar_clusters   = self.v_nm_clusters.get(),
            nm_cluster_limiar  = float(self.e_nm_limiar.get() or 0.25),
            nm_clusters_max    = int(self.e_nm_max_cl.get() or 10),
            sequencia          = self.sequencia if self.v_usar_seq.get() else [],
        )

    def _iniciar(self):
        if not os.path.isfile(SCRIPT_PATH):
            messagebox.showerror("Erro",
                f"Script não encontrado:\n{SCRIPT_PATH}")
            return
        planilha = self.entry_planilha.get().strip()
        if not self.v_usar_seq.get() and not planilha:
            messagebox.showwarning("Aviso", "Selecione um experimento antes de rodar.")
            return

        # Formato CSV — verificar energias e gravar se solicitado
        _eh_csv = planilha.lower().endswith(".csv")
        if _eh_csv and not self.v_usar_seq.get():
            self._gravar_energias_toml(planilha)
            # Verificar se todas as energias necessárias estão disponíveis
            import tomllib as _tl
            _pasta = os.path.dirname(os.path.abspath(planilha))
            _toml  = os.path.join(_pasta, "parametros.toml")
            _toml_dados = {}
            try:
                with open(_toml, "rb") as _f:
                    _toml_dados = _tl.load(_f)
            except Exception:
                pass
            _mapa = {"g1": "e_g3", "g2": "e_g4", "g3": "e_g5"}
            _nomes = {"g1": "Gaussiana 1", "g2": "Gaussiana 2", "g3": "Gaussiana 3"}
            _ativa = {"g1": self.v_ativa_g1.get(), "g2": self.v_ativa_g2.get(), "g3": self.v_ativa_g3.get()}
            _faltando = []
            for g, chave in _mapa.items():
                if not _ativa[g]:
                    continue
                _val_ui = self.e_energias[g].get().strip()
                _val_toml = _toml_dados.get(chave)
                if not _val_ui and not _val_toml:
                    _faltando.append(_nomes[g])
            if _faltando:
                faltando_str = ", ".join(_faltando)
                messagebox.showwarning(
                    "Energias não definidas",
                    f"As seguintes energias não estão definidas:\n{faltando_str}\n\n"
                    "Preencha os campos de energia ou grave valores no parametros.toml ""usando o checkbox Gravar.")
                return

        try:
            cfg = self._coletar_cfg()
            tmp = criar_script_temporario(SCRIPT_PATH, cfg)
        except Exception as ex:
            messagebox.showerror("Erro", f"Erro ao preparar script:\n{ex}")
            return

        self.rodando = True
        self.btn_iniciar.config(state="disabled")
        self.btn_parar.config(state="normal")
        self.lbl_status.config(text="Rodando...", fg=AMAR)
        self.progressbar["value"] = 0
        self.lbl_pct.config(text="")
        self._log_clear()

        import queue as _queue
        self._log_q = _queue.Queue()
        self._res_q = _queue.Queue()

        def _runner():
            env = os.environ.copy()
            env["PYTHONIOENCODING"] = "utf-8"
            env["PYTHONUTF8"]       = "1"
            env["PYTHONUNBUFFERED"] = "1"
            try:
                # Quando frozen (PyInstaller), sys.executable é o próprio .exe.
                # Chamamos com --analyze para entrar no modo análise sem abrir a GUI.
                if getattr(sys, "frozen", False):
                    cmd = [sys.executable, "--analyze", tmp]
                else:
                    cmd = [sys.executable, "-u", tmp]
                proc = subprocess.Popen(
                    cmd,
                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    text=True, encoding="utf-8", errors="replace",
                    bufsize=1, env=env,
                )
                self.proc_ref = proc
                ultima_pasta = None
                for line in proc.stdout:
                    self._log_q.put(line.rstrip())
                    for token in line.split():
                        token = token.strip().rstrip(")")
                        if os.path.isdir(token) and "Resultados" in token:
                            ultima_pasta = token
                proc.wait()
                self._res_q.put((proc.returncode, ultima_pasta))
            except Exception as ex:
                self._log_q.put(f"ERRO: {ex}")
                self._res_q.put((1, None))
            finally:
                self.proc_ref = None
                try: os.unlink(tmp)
                except: pass

        threading.Thread(target=_runner, daemon=True).start()
        self.after(300, self._poll)

    def _poll(self):
        """Lê queue e atualiza UI a cada 300ms."""
        while not self._log_q.empty():
            linha = self._log_q.get_nowait()
            self._log(linha + "\n")
            # Detectar progresso: "  50/1500  (  3.3%)"
            m = re.search(r'\(\s*([\d.]+)%\)', linha)
            if m:
                pct = float(m.group(1))
                self.progressbar["value"] = pct
                self.lbl_pct.config(text=f"{pct:.1f}%")
            # Detectar pasta de resultados
            for token in linha.split():
                token = token.strip().rstrip(")")
                if os.path.isdir(token) and "Resultados" in token:
                    self.ultima_pasta = token

        if not self._res_q.empty():
            rc, pasta = self._res_q.get()
            if pasta:
                self.ultima_pasta = pasta
            self._finalizar(rc)
            return

        if self.rodando:
            self.after(300, self._poll)

    def _parar(self):
        self.btn_parar.config(state="disabled")
        self.lbl_status.config(text="Encerrando...", fg=AMAR)

        # Cria PARAR.txt para parada limpa
        pasta = self.ultima_pasta
        if not pasta:
            candidatas = sorted([p for p in glob.glob(
                os.path.join(RESULTADOS_DIR, "*")) if os.path.isdir(p)], reverse=True)
            if candidatas:
                pasta = candidatas[0]
        if pasta and os.path.isdir(pasta):
            open(os.path.join(pasta, "PARAR.txt"), "w").close()

        # Aguarda até 10s para o script terminar sozinho, depois mata
        def _aguardar_e_matar():
            import time
            for _ in range(20):  # 20 x 0.5s = 10s
                if self.proc_ref is None or self.proc_ref.poll() is not None:
                    return  # terminou sozinho
                time.sleep(0.5)
            # Ainda rodando — mata forçado
            if self.proc_ref:
                try: self.proc_ref.terminate()
                except: pass

        threading.Thread(target=_aguardar_e_matar, daemon=True).start()

    def _finalizar(self, rc):
        self.rodando = False
        self.btn_iniciar.config(state="normal")
        self.btn_parar.config(state="disabled")
        if rc == 0:
            self.progressbar["value"] = 100
            self.lbl_pct.config(text="100%")
            self.lbl_status.config(text="✓ Concluído", fg=VERDE)
            self._atualizar_resultados()
        elif rc == -1:
            self.lbl_status.config(text="⏹ Interrompido", fg=AMAR)
        else:
            self.lbl_status.config(text=f"✗ Erro (código {rc})", fg=VERM)

    def _log(self, texto):
        self.terminal.config(state="normal")
        self.terminal.insert("end", texto)
        self.terminal.see("end")
        self.terminal.config(state="disabled")

    def _log_clear(self):
        self.terminal.config(state="normal")
        self.terminal.delete("1.0", "end")
        self.terminal.config(state="disabled")

    def _atualizar_resultados(self):
        for w in self.frame_res.winfo_children():
            w.destroy()

        pasta = self.ultima_pasta
        if not pasta:
            candidatas = sorted([p for p in glob.glob(
                os.path.join(RESULTADOS_DIR, "*")) if os.path.isdir(p)], reverse=True)
            if candidatas:
                pasta = candidatas[0]

        if not pasta or not os.path.isdir(pasta):
            tk.Label(self.frame_res, text="Nenhum resultado encontrado.",
                     bg=BG, fg=TEXTO2, font=FONTE).pack(pady=30)
            return

        tk.Label(self.frame_res, text=os.path.basename(pasta),
                 bg=BG, fg=AZUL, font=("Courier New", 11, "bold")).pack(anchor="w", pady=(8,4))

        # Métricas do TXT
        txts = glob.glob(os.path.join(pasta, "*.txt"))
        if txts:
            with open(txts[0], encoding="utf-8", errors="replace") as f:
                conteudo = f.read()
            r2 = re.search(r'R²\s*=\s*([\d.]+)', conteudo)
            rmse = re.search(r'RMSE\s*=\s*([\d.]+)', conteudo)
            mf = tk.Frame(self.frame_res, bg=BG); mf.pack(anchor="w", pady=4)
            if r2:
                tk.Label(mf, text=f"R² = {r2.group(1)}", bg=BG,
                         fg=VERDE, font=("Courier New", 12, "bold")).pack(side="left", padx=(0,20))
            if rmse:
                tk.Label(mf, text=f"RMSE = {rmse.group(1)}", bg=BG,
                         fg=TEXTO, font=("Courier New", 12)).pack(side="left")

        # Botões para abrir pasta e Excel
        bf = tk.Frame(self.frame_res, bg=BG); bf.pack(anchor="w", pady=6)
        btn(bf, "📁 Abrir pasta",
            lambda: os.startfile(pasta)).pack(side="left", padx=(0,8))
        xlsx = glob.glob(os.path.join(pasta, "*.xlsx"))
        if xlsx:
            btn(bf, "📊 Abrir Excel",
                lambda: os.startfile(xlsx[0])).pack(side="left")

    def _atualizar_historico(self):
        for w in self.frame_hist.winfo_children():
            w.destroy()

        if not os.path.isdir(RESULTADOS_DIR):
            tk.Label(self.frame_hist, text="Nenhum resultado ainda.",
                     bg=BG, fg=TEXTO2, font=FONTE).pack(pady=20)
            return

        pastas = sorted([p for p in glob.glob(os.path.join(RESULTADOS_DIR, "*"))
                         if os.path.isdir(p)], reverse=True)
        if not pastas:
            tk.Label(self.frame_hist, text="Nenhuma análise anterior encontrada.",
                     bg=BG, fg=TEXTO2, font=FONTE).pack(pady=20)
            return

        for pasta in pastas:
            f = tk.Frame(self.frame_hist, bg=BG2, bd=1, relief="groove")
            f.pack(fill="x", padx=10, pady=3)
            tk.Label(f, text=os.path.basename(pasta), bg=BG2,
                     fg=AZUL, font=FONTE_M).pack(side="left", padx=8, pady=4)
            bf = tk.Frame(f, bg=BG2); bf.pack(side="right", padx=8)
            btn(bf, "📁 Abrir", lambda p=pasta: os.startfile(p)).pack(side="left", padx=2)
            xlsx = glob.glob(os.path.join(pasta, "*.xlsx"))
            if xlsx:
                btn(bf, "📊 Excel",
                    lambda x=xlsx[0]: os.startfile(x)).pack(side="left", padx=2)


if __name__ == "__main__":
    if len(sys.argv) >= 3 and sys.argv[1] == "--analyze":
        import runpy
        script = sys.argv[2]
        sys.argv = [script]
        runpy.run_path(script, run_name="__main__")
    else:
        app = NTApp()
        app.mainloop()
