import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import glob
import json
import time
import pyautogui
import openpyxl
from datetime import datetime

pyautogui.FAILSAFE = True  # mover mouse pro canto superior esquerdo = parar tudo

# ─── CONFIG ────────────────────────────────────────────────────────────────────
CONFIG_FILE = "config_bio.json"

DEFAULT_STEPS_BIOVIDA = [
    {"nome": "Campo de busca (ID)",        "x": 0, "y": 0, "acao": "click_type",     "texto": "{ID}",        "ativo": True},
    {"nome": "Botão Pesquisar",            "x": 0, "y": 0, "acao": "click",          "texto": "",            "ativo": True},
    {"nome": "Primeiro resultado",         "x": 0, "y": 0, "acao": "click",          "texto": "",            "ativo": True},
    {"nome": "Botão Anexar",               "x": 0, "y": 0, "acao": "click",          "texto": "",            "ativo": True},
    {"nome": "Barra de endereço (pasta)",  "x": 0, "y": 0, "acao": "click_type",     "texto": "{PASTA_PDF}", "ativo": True},
    {"nome": "Confirmar pasta (Enter)",    "x": 0, "y": 0, "acao": "enter",          "texto": "",            "ativo": True},
    {"nome": "Campo nome arquivo",         "x": 0, "y": 0, "acao": "click_type",     "texto": "{NOME_PDF}",  "ativo": True},
    {"nome": "Botão Abrir/Confirmar",      "x": 0, "y": 0, "acao": "click",          "texto": "",            "ativo": True},
]

DEFAULT_STEPS_BIOCLOMA = [
    {"nome": "Campo de busca (ID)",        "x": 0, "y": 0, "acao": "click_type",     "texto": "{ID}",        "ativo": True},
    {"nome": "Botão Pesquisar",            "x": 0, "y": 0, "acao": "click",          "texto": "",            "ativo": True},
    {"nome": "Primeiro resultado",         "x": 0, "y": 0, "acao": "click",          "texto": "",            "ativo": True},
    {"nome": "Botão Anexar",               "x": 0, "y": 0, "acao": "click",          "texto": "",            "ativo": True},
    {"nome": "Barra de endereço (pasta)",  "x": 0, "y": 0, "acao": "click_type",     "texto": "{PASTA_PDF}", "ativo": True},
    {"nome": "Confirmar pasta (Enter)",    "x": 0, "y": 0, "acao": "enter",          "texto": "",            "ativo": True},
    {"nome": "Campo nome arquivo",         "x": 0, "y": 0, "acao": "click_type",     "texto": "{NOME_PDF}",  "ativo": True},
    {"nome": "Botão Abrir/Confirmar",      "x": 0, "y": 0, "acao": "click",          "texto": "",            "ativo": True},
]

ACOES = ["click", "click_type", "double_click", "right_click", "enter", "hotkey_ctrl_a"]
VARIAVEIS = ["{ID}", "{PASTA_PDF}", "{NOME_PDF}"]

# Paleta — base neutra escura, cores da marca como acento
BG     = "#141414"   # quase preto neutro — fundo principal
PANEL  = "#1e1e1e"   # cinza escuro — painéis
CARD   = "#252525"   # cinza médio — cards alternados
BORDER = "#3a3535"   # bordô acinzentado suave — bordas
ACCENT = "#9E192B"   # vermelho bordô da marca — botões e destaques
BLUE   = "#2a5248"   # verde escuro da marca clareado — títulos de seção
PURPLE = "#c28a50"   # terracota quente — variáveis {ID}
TEXT   = "#f0ebe8"   # branco quente — texto principal
MUTED  = "#7a7170"   # cinza muted — texto secundário
DANGER = "#9E192B"   # bordô — erros
WARN   = "#c4883a"   # âmbar — avisos
GREEN  = "#4caf82"   # verde menta — sucesso
FNT    = "Courier New"


def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {
        "pasta_biovida": "",
        "pasta_biocloma": "",
        "planilha_biocloma": "",
        "delay_entre_passos": 0.8,
        "delay_entre_registros": 2.0,
        "delay_pos_click": 0.3,
        "steps_biovida": DEFAULT_STEPS_BIOVIDA,
        "steps_biocloma": DEFAULT_STEPS_BIOCLOMA,
    }


def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)


# ─── LÓGICA ────────────────────────────────────────────────────────────────────

def extrair_id_biovida(nome):
    base = os.path.splitext(nome)[0]
    parts = base.split("-")
    return parts[-1].strip() if len(parts) >= 2 else None


def listar_pdfs_biovida(pasta):
    items = []
    for arq in glob.glob(os.path.join(pasta, "*.pdf")):
        nome = os.path.basename(arq)
        id_ = extrair_id_biovida(nome)
        if id_:
            items.append({"arquivo": arq, "nome": nome, "id": id_})
    return items


def listar_pdfs_biocloma(pasta, planilha):
    wb = openpyxl.load_workbook(planilha, read_only=True)
    ws = wb.active
    ids = [str(row[0]).strip() for row in ws.iter_rows(values_only=True)
           if row and row[0] is not None]
    wb.close()
    items = []
    for id_ in ids:
        found = glob.glob(os.path.join(pasta, f"*{id_}*.pdf"))
        if found:
            items.append({"arquivo": found[0], "nome": os.path.basename(found[0]), "id": id_})
    return items


def resolver_texto(texto, id_, pasta_pdf, nome_pdf):
    return (texto
            .replace("{ID}", id_)
            .replace("{PASTA_PDF}", pasta_pdf)
            .replace("{NOME_PDF}", nome_pdf))


def executar_passo(step, id_, pasta_pdf, nome_pdf, delay_click, log_fn):
    if not step.get("ativo", True):
        log_fn(f"  ⏭  Ignorado: {step['nome']}", "muted")
        return

    x    = int(step.get("x", 0))
    y    = int(step.get("y", 0))
    acao = step.get("acao", "click")
    texto = resolver_texto(step.get("texto", ""), id_, pasta_pdf, nome_pdf)
    nome  = step.get("nome", "")

    coords = f"({x},{y})" if (x or y) else "(sem coord)"
    log_fn(f"  › {nome}  [{acao}]  {coords}", "action")

    if acao == "click":
        pyautogui.click(x, y)

    elif acao == "click_type":
        pyautogui.click(x, y)
        time.sleep(delay_click)
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.15)
        pyautogui.typewrite(str(texto), interval=0.04)

    elif acao == "double_click":
        pyautogui.doubleClick(x, y)

    elif acao == "right_click":
        pyautogui.rightClick(x, y)

    elif acao == "enter":
        if x > 0 or y > 0:
            pyautogui.click(x, y)
            time.sleep(delay_click)
        pyautogui.press("enter")

    elif acao == "hotkey_ctrl_a":
        if x > 0 or y > 0:
            pyautogui.click(x, y)
            time.sleep(delay_click)
        pyautogui.hotkey("ctrl", "a")

    time.sleep(delay_click)


# ─── APP ───────────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("BioAutomação  ⬡  BioVida & Biocloma")
        self.geometry("1060x840")
        self.minsize(860, 680)
        self.configure(bg=BG)

        self.cfg = load_config()
        self._parar   = False
        self._rodando = False

        self._sv_biovida  = []
        self._sv_biocloma = []

        self._build()

    # ══════════════════════════════════════════════════════════════════════════
    # HELPERS DE LAYOUT
    # ══════════════════════════════════════════════════════════════════════════

    def _card(self, parent, titulo=None, pady=6, fill="x"):
        outer = tk.Frame(parent, bg=BORDER)
        outer.pack(fill=fill, pady=pady, expand=(fill == "both"))
        inner = tk.Frame(outer, bg=PANEL)
        inner.pack(fill=fill, padx=1, pady=1, expand=(fill == "both"))
        if titulo:
            tk.Label(inner, text=titulo, font=(FNT, 9, "bold"),
                     bg=PANEL, fg=BLUE, pady=7, padx=14).pack(anchor="w")
            tk.Frame(inner, bg=BORDER, height=1).pack(fill="x")
        return inner

    def _lbl(self, parent, text, width=0, fg=MUTED, bold=False, bg=None):
        bg = bg or PANEL
        f = (FNT, 9, "bold") if bold else (FNT, 9)
        kw = dict(text=text, font=f, bg=bg, fg=fg)
        if width:
            kw["width"] = width
        return tk.Label(parent, **kw)

    def _entry(self, parent, width=None, fg=TEXT):
        kw = dict(font=(FNT, 9), bg=CARD, fg=fg,
                  insertbackground=ACCENT, relief="flat",
                  highlightbackground=BORDER, highlightthickness=1,
                  highlightcolor=ACCENT)
        if width:
            kw["width"] = width
        return tk.Entry(parent, **kw)

    def _btn(self, parent, text, cmd, fg=BG, bg=ACCENT, padx=10, pady=4):
        return tk.Button(parent, text=text, font=(FNT, 9, "bold"),
                         bg=bg, fg=fg, relief="flat", cursor="hand2",
                         activebackground=bg, padx=padx, pady=pady,
                         command=cmd)

    # ══════════════════════════════════════════════════════════════════════════
    # BUILD PRINCIPAL
    # ══════════════════════════════════════════════════════════════════════════

    def _build(self):
        # Header
        hdr = tk.Frame(self, bg=BG, pady=12)
        hdr.pack(fill="x", padx=22)
        tk.Label(hdr, text="⬡", font=(FNT, 20, "bold"), bg=BG, fg=ACCENT).pack(side="left")
        tk.Label(hdr, text=" BIO", font=(FNT, 18, "bold"), bg=BG, fg=ACCENT).pack(side="left")
        tk.Label(hdr, text="AUTOMAÇÃO", font=(FNT, 18, "bold"), bg=BG, fg=TEXT).pack(side="left")
        tk.Label(hdr, text="  v2.0 — Coordenadas XY", font=(FNT, 9),
                 bg=BG, fg=MUTED).pack(side="left", pady=(6, 0))

        tk.Frame(self, bg=ACCENT, height=2).pack(fill="x", padx=22)

        # Notebook
        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("Bio.TNotebook", background=BG, borderwidth=0, tabmargins=0)
        style.configure("Bio.TNotebook.Tab", background=CARD, foreground=MUTED,
                        font=(FNT, 10, "bold"), padding=[18, 7], borderwidth=0)
        style.map("Bio.TNotebook.Tab",
                  background=[("selected", PANEL)],
                  foreground=[("selected", ACCENT)])

        nb = ttk.Notebook(self, style="Bio.TNotebook")
        nb.pack(fill="both", expand=True, padx=22, pady=10)

        self.t_cfg = tk.Frame(nb, bg=BG)
        self.t_bio = tk.Frame(nb, bg=BG)
        self.t_clo = tk.Frame(nb, bg=BG)
        self.t_run = tk.Frame(nb, bg=BG)

        nb.add(self.t_cfg, text="  ⚙  Configurações  ")
        nb.add(self.t_bio, text="  ⬡  BioVida — Passos  ")
        nb.add(self.t_clo, text="  ⬡  Biocloma — Passos  ")
        nb.add(self.t_run, text="  ▶  Execução  ")

        self._build_cfg()
        self._build_steps_tab(self.t_bio, "biovida")
        self._build_steps_tab(self.t_clo, "biocloma")
        self._build_run()

    # ══════════════════════════════════════════════════════════════════════════
    # ABA CONFIGURAÇÕES
    # ══════════════════════════════════════════════════════════════════════════

    def _build_cfg(self):
        wrap = tk.Frame(self.t_cfg, bg=BG)
        wrap.pack(fill="both", expand=True, padx=16, pady=8)

        c1 = self._card(wrap, "◈  BioVida")
        self._e_pasta_bio = self._make_dir_row(c1,  "Pasta dos PDFs:", "pasta_biovida")
        tk.Label(c1, text="  Formato: 999999999-12345.pdf  →  ID usado = 12345",
                 font=(FNT, 8), bg=PANEL, fg=MUTED, padx=14, pady=4).pack(anchor="w")

        c2 = self._card(wrap, "◈  Biocloma")
        self._e_plan_clo  = self._make_file_row(c2, "Planilha (.xlsx):", "planilha_biocloma")
        self._e_pasta_clo = self._make_dir_row(c2,  "Pasta dos PDFs:", "pasta_biocloma")
        tk.Label(c2, text="  Coluna A da planilha = IDs para busca.",
                 font=(FNT, 8), bg=PANEL, fg=MUTED, padx=14, pady=4).pack(anchor="w")

        c3 = self._card(wrap, "◈  Timing")
        delays = [
            ("Delay entre passos (s):",    "delay_entre_passos",    0.1, 15.0, 0.1),
            ("Delay entre registros (s):", "delay_entre_registros", 0.5, 30.0, 0.5),
            ("Delay pós-click (s):",       "delay_pos_click",       0.1,  5.0, 0.1),
        ]
        self._spins = {}
        for label, key, lo, hi, inc in delays:
            row = tk.Frame(c3, bg=PANEL); row.pack(fill="x", padx=14, pady=5)
            self._lbl(row, label, width=28).pack(side="left")
            sp = tk.Spinbox(row, from_=lo, to=hi, increment=inc, width=7,
                            font=(FNT, 9), bg=CARD, fg=TEXT, relief="flat",
                            buttonbackground=BORDER,
                            highlightbackground=BORDER, highlightthickness=1)
            sp.pack(side="left", ipady=3)
            sp.delete(0, "end"); sp.insert(0, str(self.cfg.get(key, lo)))
            self._spins[key] = sp

        tk.Label(c3,
                 text="  ⚠  Mover mouse pro canto superior esquerdo ativa FAILSAFE e para tudo.",
                 font=(FNT, 8), bg=PANEL, fg=WARN, padx=14, pady=6).pack(anchor="w")

        self._btn(wrap, "💾  SALVAR TUDO", self._save_all,
                  bg=BLUE, fg=BG, padx=22, pady=8).pack(pady=14, anchor="w")

    def _make_dir_row(self, parent, label, key):
        row = tk.Frame(parent, bg=PANEL); row.pack(fill="x", padx=14, pady=6)
        self._lbl(row, label, width=20).pack(side="left")
        e = self._entry(row); e.pack(side="left", fill="x", expand=True, ipady=4, padx=(0, 6))
        e.insert(0, self.cfg.get(key, ""))
        self._btn(row, "📁", lambda k=key, en=e: self._browse_dir(k, en),
                  bg=BORDER, fg=TEXT, padx=8, pady=3).pack(side="left")
        return e

    def _make_file_row(self, parent, label, key):
        row = tk.Frame(parent, bg=PANEL); row.pack(fill="x", padx=14, pady=6)
        self._lbl(row, label, width=20).pack(side="left")
        e = self._entry(row); e.pack(side="left", fill="x", expand=True, ipady=4, padx=(0, 6))
        e.insert(0, self.cfg.get(key, ""))
        self._btn(row, "📄", lambda k=key, en=e: self._browse_file(k, en),
                  bg=BORDER, fg=TEXT, padx=8, pady=3).pack(side="left")
        return e

    def _browse_dir(self, key, entry):
        p = filedialog.askdirectory()
        if p:
            entry.delete(0, "end"); entry.insert(0, p)
            self.cfg[key] = p

    def _browse_file(self, key, entry):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if p:
            entry.delete(0, "end"); entry.insert(0, p)
            self.cfg[key] = p

    # ══════════════════════════════════════════════════════════════════════════
    # ABA PASSOS
    # ══════════════════════════════════════════════════════════════════════════

    def _build_steps_tab(self, tab, modelo):
        sv_list = self._sv_biovida if modelo == "biovida" else self._sv_biocloma
        default = DEFAULT_STEPS_BIOVIDA if modelo == "biovida" else DEFAULT_STEPS_BIOCLOMA
        steps   = self.cfg.get(f"steps_{modelo}", default)

        # Canvas scrollável
        canvas = tk.Canvas(tab, bg=BG, highlightthickness=0)
        vsb = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        inner = tk.Frame(canvas, bg=BG)
        wid = canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(wid, width=e.width))
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

        wrap = tk.Frame(inner, bg=BG, padx=16, pady=10)
        wrap.pack(fill="both", expand=True)

        # Info de variáveis
        info = tk.Frame(wrap, bg=CARD, pady=8, padx=14)
        info.pack(fill="x", pady=(0, 8))
        tk.Label(info, text="Variáveis disponíveis nos campos 'Texto':  " + "    ".join(VARIAVEIS),
                 font=(FNT, 8), bg=CARD, fg=PURPLE).pack(anchor="w")
        tk.Label(info, text="Ações 'enter' / 'hotkey_ctrl_a': X=0,Y=0 = sem clique prévio  |  Ação 'click_type': clica + Ctrl+A + digita",
                 font=(FNT, 8), bg=CARD, fg=MUTED).pack(anchor="w")

        # Cabeçalho das colunas
        hdr = tk.Frame(wrap, bg=BORDER, pady=1)
        hdr.pack(fill="x", pady=(0, 2))
        hdr_inner = tk.Frame(hdr, bg=PANEL, padx=10, pady=5)
        hdr_inner.pack(fill="x", padx=1)
        for txt, w in [("#",3),("✓",2),("Nome do Passo",22),("Ação",16),("X",6),("Y",6),("Texto / Variável",18),("",4)]:
            tk.Label(hdr_inner, text=txt, font=(FNT, 8, "bold"),
                     bg=PANEL, fg=BLUE, width=w, anchor="w").pack(side="left", padx=2)

        # Frame dos passos
        steps_frame = tk.Frame(wrap, bg=BG)
        steps_frame.pack(fill="x")

        sv_list.clear()
        for i, step in enumerate(steps):
            self._add_step_row(steps_frame, sv_list, i, step)

        # Controles
        ctrl = tk.Frame(wrap, bg=BG, pady=8)
        ctrl.pack(fill="x")
        self._btn(ctrl, "+  ADICIONAR PASSO",
                  lambda sf=steps_frame, sv=sv_list: self._add_step_row(
                      sf, sv, len(sv),
                      {"nome": "Novo passo", "x": 0, "y": 0, "acao": "click", "texto": "", "ativo": True}
                  ), bg=ACCENT, fg=BG, padx=16, pady=6).pack(side="left", padx=(0, 8))

        self._btn(ctrl, "💾  SALVAR PASSOS",
                  lambda mo=modelo: self._save_steps(mo),
                  bg=BLUE, fg=BG, padx=16, pady=6).pack(side="left")

        # Ferramenta de captura
        cap = self._card(wrap, "🎯  Capturar Coordenada do Mouse", pady=8)
        cap_row = tk.Frame(cap, bg=PANEL); cap_row.pack(fill="x", padx=14, pady=8)
        self._lbl(cap_row, "Abra o app, posicione o mouse e clique Capturar →").pack(side="left")

        self._lbl(cap_row, "   X:", fg=TEXT).pack(side="left")
        ex = self._entry(cap_row, width=7, fg=ACCENT)
        ex.pack(side="left", padx=(2, 8), ipady=3)

        self._lbl(cap_row, "Y:", fg=TEXT).pack(side="left")
        ey = self._entry(cap_row, width=7, fg=ACCENT)
        ey.pack(side="left", padx=(2, 8), ipady=3)

        self._btn(cap_row, "📍 CAPTURAR",
                  lambda ex=ex, ey=ey: self._capturar(ex, ey),
                  bg=PURPLE, fg=BG, padx=12, pady=4).pack(side="left")

        tk.Label(cap,
                 text="  Cole o X e Y capturados no campo do passo correspondente acima.",
                 font=(FNT, 8), bg=PANEL, fg=MUTED, padx=14, pady=4).pack(anchor="w")

    def _add_step_row(self, frame, sv_list, idx, step):
        bg_row = CARD if idx % 2 == 0 else PANEL

        container = tk.Frame(frame, bg=BORDER, pady=1)
        container.pack(fill="x", pady=2)
        row = tk.Frame(container, bg=bg_row, pady=6, padx=10)
        row.pack(fill="x", padx=1)

        # Número
        tk.Label(row, text=f"{len(sv_list)+1:02d}", font=(FNT, 9, "bold"),
                 bg=bg_row, fg=PURPLE, width=3).pack(side="left")

        # Ativo
        v_ativo = tk.BooleanVar(value=step.get("ativo", True))
        ck = tk.Checkbutton(row, variable=v_ativo, bg=bg_row,
                             activebackground=bg_row, cursor="hand2",
                             selectcolor=BG)
        ck.pack(side="left", padx=(0, 4))

        # Nome
        v_nome = tk.StringVar(value=step.get("nome", ""))
        tk.Entry(row, textvariable=v_nome, font=(FNT, 9), bg=BG, fg=TEXT,
                 width=22, relief="flat",
                 highlightbackground=BORDER, highlightthickness=1,
                 highlightcolor=ACCENT, insertbackground=ACCENT
                 ).pack(side="left", ipady=3, padx=(0, 6))

        # Ação
        v_acao = tk.StringVar(value=step.get("acao", "click"))
        cb = ttk.Combobox(row, textvariable=v_acao, values=ACOES,
                          font=(FNT, 9), width=16, state="readonly")
        cb.pack(side="left", padx=(0, 6))

        # X
        tk.Label(row, text="X:", font=(FNT, 9), bg=bg_row, fg=MUTED).pack(side="left")
        v_x = tk.StringVar(value=str(step.get("x", 0)))
        tk.Entry(row, textvariable=v_x, font=(FNT, 9, "bold"), width=7,
                 bg=BG, fg=