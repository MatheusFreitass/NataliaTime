"""
migrador.py — Criador de experimentos CSV + TOML para o Natalia Time

Permite ao usuário criar a pasta de um experimento no novo formato,
preenchendo os parâmetros físicos e colando os dados copiados do Excel.

Uso: python migrador.py  (ou duplo clique no .exe gerado)
"""

import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox

# ── Cores (mesmo tema do NT_app_tk.py) ────────────────────────────────────────
BG    = "#0e1117"
BG2   = "#161b22"
BG3   = "#1c2128"
BORDA = "#30363d"
AZUL  = "#388bfd"
TEXTO = "#e0e0e0"
TEXTO2= "#8b949e"
VERDE = "#3fb950"
VERM  = "#f85149"
AMAR  = "#d29922"
FONTE = ("Segoe UI", 10)
FONTE_TITULO = ("Segoe UI", 11, "bold")
FONTE_SEC    = ("Segoe UI", 10, "bold")


def _entry(parent, width=20, default=""):
    e = tk.Entry(parent, width=width, bg=BG3, fg=TEXTO, insertbackground=TEXTO,
                 relief="flat", font=FONTE,
                 highlightbackground=BORDA, highlightcolor=AZUL, highlightthickness=1)
    if default:
        e.insert(0, default)
    return e


def _lbl(parent, text, fg=TEXTO2):
    return tk.Label(parent, text=text, bg=BG2, fg=fg, font=FONTE)


def _btn(parent, text, cmd, cor=AZUL):
    return tk.Button(parent, text=text, command=cmd,
                     bg=cor, fg="#ffffff", font=FONTE,
                     relief="flat", padx=10, pady=4,
                     activebackground=cor, activeforeground="#ffffff",
                     cursor="hand2")


def _secao(parent, titulo):
    f = tk.LabelFrame(parent, text=f"  {titulo}  ", bg=BG2, fg=AZUL,
                      font=FONTE_SEC, relief="flat",
                      highlightbackground=BORDA, highlightthickness=1,
                      padx=8, pady=6)
    f.pack(fill="x", padx=10, pady=(6, 2))
    return f


class MigradorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Natalia Time — Criador de Experimento")
        self.configure(bg=BG)
        self.resizable(True, True)
        self.minsize(600, 700)

        self._build_ui()
        self._centralizar()

    def _centralizar(self):
        self.update_idletasks()
        w, h = 680, 780
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    def _build_ui(self):
        # Título
        tk.Label(self, text="Criador de Experimento",
                 bg=BG, fg=AZUL, font=("Segoe UI", 14, "bold")).pack(pady=(14, 2))
        tk.Label(self, text="Gera a pasta com dados.csv e parametros.toml para uso no Natalia Time",
                 bg=BG, fg=TEXTO2, font=FONTE).pack(pady=(0, 10))

        # ── Destino ───────────────────────────────────────────────────────────
        s0 = _secao(self, "📁 Destino")
        r0 = tk.Frame(s0, bg=BG2); r0.pack(fill="x", pady=2)
        _lbl(r0, "Pasta raiz:").pack(side="left")
        self.e_destino = _entry(r0, width=45,
                                default=os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                                     "Dados e Parâmetros"))
        self.e_destino.pack(side="left", padx=(6, 4))
        _btn(r0, "…", self._procurar_destino, cor=BG3).pack(side="left")

        r0b = tk.Frame(s0, bg=BG2); r0b.pack(fill="x", pady=2)
        _lbl(r0b, "Nome do experimento:").pack(side="left")
        self.e_nome = _entry(r0b, width=30, default="")
        self.e_nome.pack(side="left", padx=(6, 0))
        _lbl(r0b, "  (ex: O2_300eV)", fg=TEXTO2).pack(side="left")

        # ── Parâmetros ────────────────────────────────────────────────────────
        s1 = _secao(self, "⚙ Parâmetros físicos")

        campos = [
            ("T_gas",   "T gás (K)",          "Temperatura do gás de expansão",                 "300"),
            ("m_frag",  "m fragmento",         "Massa do fragmento iônico (ex: 16 para O⁺)",     "16.0"),
            ("m_mol",   "m molécula mãe",      "Massa da molécula mãe (ex: 32 para O₂)",  "32.0"),
            ("DE",      "DE — Extração (eV)",  "Campo Elétrico de extração",           "350.0"),
            ("D",       "D — distância (m)",   "Diâmetro do colimador",                               "6.8"),
            ("LL",      "LL — comprimento efetivo (m)","Comprimento efetivo do cilindro de criação dos ions",                            "6.8"),
            ("I7_NORM", "I7 — normaliz. (ns)", "Tempo de normalização das amplitudes",           "500.0"),
        ]
        self.e_params = {}
        for chave, rotulo, tooltip, padrao in campos:
            r = tk.Frame(s1, bg=BG2); r.pack(fill="x", pady=3)
            lbl = _lbl(r, f"{rotulo}:")
            lbl.config(width=16, anchor="w")
            lbl.pack(side="left")
            e = _entry(r, width=14, default=padrao)
            e.pack(side="left", padx=(4, 10))
            _lbl(r, tooltip, fg=TEXTO2).pack(side="left")
            self.e_params[chave] = e

        tk.Label(s1,
                 text="β_MB será calculado automaticamente a partir de T e m molécula mãe.",
                 bg=BG2, fg=TEXTO2, font=FONTE).pack(anchor="w", pady=(4, 0))
        tk.Label(s1,
                 text="As energias (E_G1, E_G2, E_G3) são definidas na interface do Natalia Time.",
                 bg=BG2, fg=AMAR, font=FONTE).pack(anchor="w", pady=(2, 0))

        # ── Dados ─────────────────────────────────────────────────────────────
        s2 = _secao(self, "📋 Dados experimentais")
        tk.Label(s2,
                 text="Cole aqui as células copiadas da planilha (colunas Tempo e Sinal, separadas por Tab):",
                 bg=BG2, fg=TEXTO2, font=FONTE).pack(anchor="w", pady=(0, 4))

        frame_txt = tk.Frame(s2, bg=BG2)
        frame_txt.pack(fill="both", expand=True)

        scroll_y = tk.Scrollbar(frame_txt, orient="vertical")
        scroll_y.pack(side="right", fill="y")
        scroll_x = tk.Scrollbar(frame_txt, orient="horizontal")
        scroll_x.pack(side="bottom", fill="x")

        self.txt_dados = tk.Text(frame_txt, height=14, bg=BG3, fg=TEXTO,
                                 insertbackground=TEXTO, font=("Consolas", 9),
                                 relief="flat", wrap="none",
                                 yscrollcommand=scroll_y.set,
                                 xscrollcommand=scroll_x.set,
                                 highlightbackground=BORDA, highlightthickness=1)
        self.txt_dados.pack(fill="both", expand=True)
        scroll_y.config(command=self.txt_dados.yview)
        scroll_x.config(command=self.txt_dados.xview)

        # Placeholder
        self._placeholder_ativo = True
        _placeholder = "510\t1.000000\n560\t0.985348\n610\t0.950200\n..."
        self.txt_dados.insert("1.0", _placeholder)
        self.txt_dados.config(fg=TEXTO2)
        self.txt_dados.bind("<FocusIn>",  self._limpar_placeholder)
        self.txt_dados.bind("<FocusOut>", self._repor_placeholder)
        self._placeholder_txt = _placeholder

        # Contagem de linhas
        self.lbl_linhas = tk.Label(s2, text="", bg=BG2, fg=TEXTO2, font=FONTE)
        self.lbl_linhas.pack(anchor="e", pady=(2, 0))
        self.txt_dados.bind("<KeyRelease>", self._atualizar_contagem)
        self.txt_dados.bind("<<Paste>>",    lambda e: self.after(50, self._atualizar_contagem))

        # ── Botões ────────────────────────────────────────────────────────────
        bf = tk.Frame(self, bg=BG); bf.pack(pady=12)
        _btn(bf, "✔ Criar experimento", self._criar, cor=VERDE).pack(side="left", padx=8)
        _btn(bf, "✖ Limpar", self._limpar, cor=BG3).pack(side="left", padx=4)

        # Status
        self.lbl_status = tk.Label(self, text="", bg=BG, fg=TEXTO2, font=FONTE)
        self.lbl_status.pack(pady=(0, 8))

    # ── Ações ─────────────────────────────────────────────────────────────────

    def _procurar_destino(self):
        p = filedialog.askdirectory(title="Escolha a pasta raiz para os experimentos")
        if p:
            self.e_destino.delete(0, "end")
            self.e_destino.insert(0, p)

    def _limpar_placeholder(self, event=None):
        if self._placeholder_ativo:
            self.txt_dados.delete("1.0", "end")
            self.txt_dados.config(fg=TEXTO)
            self._placeholder_ativo = False

    def _repor_placeholder(self, event=None):
        conteudo = self.txt_dados.get("1.0", "end").strip()
        if not conteudo:
            self.txt_dados.insert("1.0", self._placeholder_txt)
            self.txt_dados.config(fg=TEXTO2)
            self._placeholder_ativo = True
            self.lbl_linhas.config(text="")

    def _atualizar_contagem(self, event=None):
        if self._placeholder_ativo:
            return
        linhas = self._parsear_dados()
        n = len(linhas)
        self.lbl_linhas.config(
            text=f"{n} pontos válidos" if n else "Nenhum ponto válido detectado",
            fg=VERDE if n >= 5 else VERM
        )

    def _parsear_dados(self):
        """Lê o texto colado e retorna lista de (tempo, sinal)."""
        texto = self.txt_dados.get("1.0", "end")
        pontos = []
        for linha in texto.splitlines():
            linha = linha.strip()
            if not linha:
                continue
            # Suporta tab (padrão Excel) e vírgula
            partes = linha.replace(",", ".").split("\t") if "\t" in linha else linha.replace(";", ",").split(",")
            if len(partes) < 2:
                continue
            try:
                t = float(partes[0].replace(",", "."))
                s = float(partes[1].replace(",", "."))
                pontos.append((t, s))
            except ValueError:
                continue  # cabeçalho ou linha inválida
        return pontos

    def _limpar(self):
        for e in self.e_params.values():
            e.delete(0, "end")
        self.e_nome.delete(0, "end")
        self.txt_dados.delete("1.0", "end")
        self._placeholder_ativo = False
        self._repor_placeholder()
        self.lbl_status.config(text="")

    def _criar(self):
        self.lbl_status.config(text="")

        # ── Validar nome do experimento ───────────────────────────────────────
        nome = self.e_nome.get().strip()
        if not nome:
            messagebox.showwarning("Nome vazio", "Informe o nome do experimento.")
            self.e_nome.focus_set()
            return

        # Sanitizar nome para ser válido como nome de pasta
        chars_invalidos = set(r'\/:*?"<>|')
        if any(c in chars_invalidos for c in nome):
            messagebox.showwarning("Nome inválido",
                f"O nome '{nome}' contém caracteres inválidos para nome de pasta.\n"
                r"Evite: \ / : * ? \" < > |")
            return

        # ── Validar parâmetros ────────────────────────────────────────────────
        valores = {}
        for chave, e in self.e_params.items():
            val_str = e.get().strip()
            if not val_str:
                messagebox.showwarning("Parâmetro vazio",
                    f"O campo '{chave}' está vazio.")
                e.focus_set()
                return
            try:
                valores[chave] = float(val_str.replace(",", "."))
            except ValueError:
                messagebox.showwarning("Valor inválido",
                    f"O valor '{val_str}' para '{chave}' não é um número válido.")
                e.focus_set()
                return

        # ── Validar dados colados ─────────────────────────────────────────────
        if self._placeholder_ativo:
            messagebox.showwarning("Dados vazios",
                "Cole os dados experimentais no campo de texto.")
            return

        pontos = self._parsear_dados()
        if len(pontos) < 5:
            messagebox.showwarning("Dados insuficientes",
                f"Apenas {len(pontos)} pontos válidos detectados.\n"
                "Verifique se os dados estão no formato correto (Tab entre colunas).")
            return

        # ── Criar pasta e arquivos ────────────────────────────────────────────
        destino_raiz = self.e_destino.get().strip()
        if not destino_raiz:
            messagebox.showwarning("Destino vazio",
                "Informe a pasta raiz de destino.")
            return

        pasta_exp = os.path.join(destino_raiz, nome)
        if os.path.exists(pasta_exp):
            resp = messagebox.askyesno("Pasta existente",
                f"A pasta '{nome}' já existe em:\n{destino_raiz}\n\n"
                "Deseja sobrescrever os arquivos?")
            if not resp:
                return

        try:
            os.makedirs(pasta_exp, exist_ok=True)

            # dados.csv
            csv_path = os.path.join(pasta_exp, "dados.csv")
            with open(csv_path, "w", encoding="utf-8") as f:
                f.write("tempo,sinal\n")
                for t, s in pontos:
                    f.write(f"{t},{s}\n")

            # Calcular beta_mb a partir de T e m_mol
            import math as _math
            T_gas  = valores["T_gas"]
            m_mol  = valores["m_mol"]
            beta_mb = _math.sqrt((4.04e5 * (300.0 / T_gas) * m_mol) / 2.0)

            # parametros.toml
            toml_path = os.path.join(pasta_exp, "parametros.toml")
            with open(toml_path, "w", encoding="utf-8") as f:
                f.write("# Parâmetros físicos do experimento\n")
                f.write(f"# Gerado pelo Natalia Time — Criador de Experimento\n")
                f.write(f"# T_gas = {T_gas} K  →  beta_mb calculado = {beta_mb:.6f}\n\n")
                f.write("# Parâmetros da molécula\n")
                f.write(f"beta_mb = {beta_mb:.6f}\n")
                f.write(f"m_frag  = {valores['m_frag']}\n")
                f.write(f"m_mol   = {m_mol}\n")
                f.write("\n# Parâmetros geométricos do espectrômetro DETOF\n")
                f.write(f"DE      = {valores['DE']}\n")
                f.write(f"D       = {valores['D']}\n")
                f.write(f"LL      = {valores['LL']}\n")
                f.write(f"I7_NORM = {valores['I7_NORM']}\n")
                f.write("\n# Energias das gaussianas (eV)\n")
                f.write("# Deixe comentado ou remova para definir pelo Natalia Time\n")
                f.write("# e_g3 = 0.7\n")
                f.write("# e_g4 = 3.0\n")
                f.write("# e_g5 = 2.0\n")

        except Exception as ex:
            messagebox.showerror("Erro ao criar",
                f"Não foi possível criar os arquivos:\n{ex}")
            return

        self.lbl_status.config(
            text=f"✔ Experimento criado com {len(pontos)} pontos em: {pasta_exp}",
            fg=VERDE)

        resp = messagebox.askyesno("Pronto!",
            f"Experimento '{nome}' criado com sucesso!\n\n"
            f"  • {len(pontos)} pontos em dados.csv\n"
            f"  • Parâmetros em parametros.toml\n\n"
            f"Pasta: {pasta_exp}\n\n"
            "Abrir a pasta agora?")
        if resp:
            try:
                os.startfile(pasta_exp)
            except Exception:
                pass


if __name__ == "__main__":
    app = MigradorApp()
    app.mainloop()
