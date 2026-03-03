import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading, os, glob, json, time, ctypes, sys
import pyautogui, openpyxl
from datetime import datetime

# ── DPI Awareness ──────────────────────────────────────────────────────────────
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try: ctypes.windll.user32.SetProcessDPIAware()
    except Exception: pass

pyautogui.FAILSAFE = True

# ── CONFIG ─────────────────────────────────────────────────────────────────────
CONFIG_FILE = "config_bio.json"

DEFAULT_STEPS = [
    {"nome": "Campo de busca (ID)",       "x":0,"y":0,"acao":"click_type", "texto":"{ID}",        "ativo":True},
    {"nome": "Botão Pesquisar",           "x":0,"y":0,"acao":"click",      "texto":"",            "ativo":True},
    {"nome": "Primeiro resultado",        "x":0,"y":0,"acao":"click",      "texto":"",            "ativo":True},
    {"nome": "Botão Anexar",              "x":0,"y":0,"acao":"click",      "texto":"",            "ativo":True},
    {"nome": "Barra de endereço (pasta)", "x":0,"y":0,"acao":"click_type", "texto":"{PASTA_PDF}", "ativo":True},
    {"nome": "Confirmar pasta (Enter)",   "x":0,"y":0,"acao":"enter",      "texto":"",            "ativo":True},
    {"nome": "Campo nome arquivo",        "x":0,"y":0,"acao":"click_type", "texto":"{NOME_PDF}",  "ativo":True},
    {"nome": "Botão Abrir/Confirmar",     "x":0,"y":0,"acao":"click",      "texto":"",            "ativo":True},
]
ACOES    = ["click","click_type","double_click","right_click","enter","hotkey_ctrl_a"]
VARIAVEIS= ["{ID}","{PASTA_PDF}","{NOME_PDF}"]

# ── PALETA ─────────────────────────────────────────────────────────────────────
BG      = "#0f0f0f"
SURFACE = "#181818"
CARD    = "#202020"
CARD2   = "#242424"
BORDER  = "#2c2c2c"
ACCENT  = "#9E192B"
ACCENTL = "#b81e31"
TEAL    = "#1B302E"
TEAL2   = "#2a5248"
GOLD    = "#c28a50"
TEXT    = "#f2ede9"
MUTED   = "#6a6360"
DIM     = "#333030"
DANGER  = "#9E192B"
WARN    = "#c47a30"
GREEN   = "#3dba7a"
FNT     = "Segoe UI"
MONO    = "Consolas"

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE,"r",encoding="utf-8") as f: return json.load(f)
        except: pass
    return {"pasta_biovida":"","pasta_biocloma":"","planilha_biocloma":"",
            "delay_entre_passos":0.8,"delay_entre_registros":2.0,"delay_pos_click":0.3,
            "steps_biovida":DEFAULT_STEPS,"steps_biocloma":DEFAULT_STEPS}

def save_config(cfg):
    with open(CONFIG_FILE,"w",encoding="utf-8") as f: json.dump(cfg,f,indent=2,ensure_ascii=False)

# ── LÓGICA ─────────────────────────────────────────────────────────────────────
def extrair_id(nome):
    base = os.path.splitext(nome)[0]; p = base.split("-")
    return p[-1].strip() if len(p)>=2 else None

def listar_biovida(pasta):
    r=[]
    for a in glob.glob(os.path.join(pasta,"*.pdf")):
        n=os.path.basename(a); i=extrair_id(n)
        if i: r.append({"arquivo":a,"nome":n,"id":i})
    return r

def listar_biocloma(pasta,planilha):
    wb=openpyxl.load_workbook(planilha,read_only=True); ws=wb.active
    ids=[str(row[0]).strip() for row in ws.iter_rows(values_only=True) if row and row[0] is not None]
    wb.close()
    r=[]
    for i in ids:
        f=glob.glob(os.path.join(pasta,f"*{i}*.pdf"))
        if f: r.append({"arquivo":f[0],"nome":os.path.basename(f[0]),"id":i})
    return r

def resolver(txt,id_,pasta,nome):
    return txt.replace("{ID}",id_).replace("{PASTA_PDF}",pasta).replace("{NOME_PDF}",nome)

def executar_passo(step,id_,pasta,nome,dc,log):
    if not step.get("ativo",True): log(f"  ⏭  Pulado: {step['nome']}","muted"); return
    x,y=int(step.get("x",0)),int(step.get("y",0))
    acao=step.get("acao","click"); txt=resolver(step.get("texto",""),id_,pasta,nome)
    coord=f"({x},{y})" if (x or y) else ""
    log(f"  › {step['nome']}  [{acao}]  {coord}","action")
    if   acao=="click":         pyautogui.click(x,y)
    elif acao=="click_type":    pyautogui.click(x,y); time.sleep(dc); pyautogui.hotkey("ctrl","a"); time.sleep(0.15); pyautogui.typewrite(str(txt),interval=0.04)
    elif acao=="double_click":  pyautogui.doubleClick(x,y)
    elif acao=="right_click":   pyautogui.rightClick(x,y)
    elif acao=="enter":
        if x or y: pyautogui.click(x,y); time.sleep(dc)
        pyautogui.press("enter")
    elif acao=="hotkey_ctrl_a":
        if x or y: pyautogui.click(x,y); time.sleep(dc)
        pyautogui.hotkey("ctrl","a")
    time.sleep(dc)

# ── APP ────────────────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("BioAutomação — BioVida & Biocloma")
        self.configure(bg=BG)
        self.minsize(900,640)

        sw,sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w,h   = min(1100, int(sw*0.80)), min(780, int(sh*0.85))
        x,y   = (sw-w)//2, (sh-h)//2
        self.geometry(f"{w}x{h}+{x}+{y}")

        self.cfg = load_config()
        self._parar = self._rodando = False
        self._sv_bio, self._sv_clo = [], []
        self._setup_styles()
        self._build()

    # ── STYLES ────────────────────────────────────────────────────────────────
    def _setup_styles(self):
        s = ttk.Style(self); s.theme_use("default")
        s.configure("TNotebook", background=BG, borderwidth=0, tabmargins=[0,0,0,0])
        s.configure("TNotebook.Tab", background=SURFACE, foreground=MUTED,
                    font=(FNT,10,"bold"), padding=[20,8], borderwidth=0)
        s.map("TNotebook.Tab",
              background=[("selected",CARD)],
              foreground=[("selected",TEXT)])
        s.configure("Accent.TProgressbar", troughcolor=BORDER,
                    background=ACCENT, borderwidth=0, thickness=6)
        s.configure("TScrollbar", background=SURFACE, troughcolor=BG,
                    borderwidth=0, arrowsize=12)
        s.map("TScrollbar", background=[("active",DIM)])
        s.configure("TCombobox", fieldbackground=CARD, background=CARD,
                    foreground=TEXT, selectbackground=ACCENT, borderwidth=0)
        s.map("TCombobox", fieldbackground=[("readonly",CARD)])

    # ── BUILD ROOT ────────────────────────────────────────────────────────────
    def _build(self):
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._build_header()

        nb = ttk.Notebook(self)
        nb.grid(row=1, column=0, sticky="nsew", padx=0, pady=0)

        self.t_cfg = tk.Frame(nb, bg=BG)
        self.t_bio = tk.Frame(nb, bg=BG)
        self.t_clo = tk.Frame(nb, bg=BG)
        self.t_run = tk.Frame(nb, bg=BG)
        nb.add(self.t_cfg, text="  Configurações  ")
        nb.add(self.t_bio, text="  BioVida — Passos  ")
        nb.add(self.t_clo, text="  Biocloma — Passos  ")
        nb.add(self.t_run, text="  Execução  ")

        self._build_cfg()
        self._build_steps(self.t_bio, "biovida", self._sv_bio)
        self._build_steps(self.t_clo, "biocloma", self._sv_clo)
        self._build_run()

    def _build_header(self):
        hdr = tk.Frame(self, bg=SURFACE, pady=0)
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.grid_columnconfigure(1, weight=1)

        # Barra vermelha lateral
        tk.Frame(hdr, bg=ACCENT, width=5).grid(row=0, column=0, sticky="ns", rowspan=2)

        title_row = tk.Frame(hdr, bg=SURFACE, padx=20, pady=12)
        title_row.grid(row=0, column=1, sticky="ew")
        title_row.grid_columnconfigure(0, weight=1)

        left = tk.Frame(title_row, bg=SURFACE)
        left.grid(row=0, column=0, sticky="w")
        tk.Label(left, text="Bio", font=(FNT,18,"bold"), bg=SURFACE, fg=ACCENT).pack(side="left")
        tk.Label(left, text="Automação", font=(FNT,18,"bold"), bg=SURFACE, fg=TEXT).pack(side="left")
        tk.Label(left, text="  BioVida & Biocloma", font=(FNT,10), bg=SURFACE, fg=MUTED).pack(side="left", pady=(5,0))

        right = tk.Frame(title_row, bg=SURFACE)
        right.grid(row=0, column=1, sticky="e")
        badge = tk.Frame(right, bg=TEAL, padx=10, pady=4)
        badge.pack(side="right")
        tk.Label(badge, text="v2.0  XY Coords", font=(FNT,8,"bold"), bg=TEAL, fg=TEXT).pack()

        tk.Frame(hdr, bg=BORDER, height=1).grid(row=1, column=0, columnspan=2, sticky="ew")

    # ── HELPERS ───────────────────────────────────────────────────────────────
    def _section(self, parent, title, pady=(0,12)):
        f = tk.Frame(parent, bg=BG)
        f.pack(fill="x", pady=pady)
        tk.Label(f, text=title, font=(FNT,9,"bold"),
                 bg=BG, fg=MUTED).pack(anchor="w", padx=2, pady=(0,4))
        card = tk.Frame(f, bg=SURFACE)
        card.pack(fill="x")
        card.config(highlightbackground=BORDER, highlightthickness=1)
        return card

    def _field_row(self, parent, label, padx=16, pady=10):
        row = tk.Frame(parent, bg=SURFACE, padx=padx, pady=pady)
        row.pack(fill="x")
        tk.Label(row, text=label, font=(FNT,9), bg=SURFACE,
                 fg=MUTED, width=22, anchor="w").pack(side="left")
        return row

    def _entry(self, parent, width=None, fg=TEXT, bg=CARD):
        kw = dict(font=(MONO,9), bg=bg, fg=fg, insertbackground=ACCENT,
                  relief="flat", highlightbackground=BORDER,
                  highlightthickness=1, highlightcolor=ACCENT)
        if width: kw["width"]=width
        return tk.Entry(parent, **kw)

    def _btn(self, parent, text, cmd, bg=ACCENT, fg=TEXT, bold=True, padx=18, pady=7):
        f = (FNT,9,"bold") if bold else (FNT,9)
        b = tk.Button(parent, text=text, font=f, bg=bg, fg=fg,
                      relief="flat", cursor="hand2", padx=padx, pady=pady,
                      activebackground=ACCENTL if bg==ACCENT else DIM,
                      activeforeground=TEXT, command=cmd)
        b.bind("<Enter>", lambda e: b.config(bg=ACCENTL if bg==ACCENT else DIM))
        b.bind("<Leave>", lambda e: b.config(bg=bg))
        return b

    def _divider(self, parent):
        tk.Frame(parent, bg=BORDER, height=1).pack(fill="x")

    # ── ABA CONFIGURAÇÕES ─────────────────────────────────────────────────────
    def _build_cfg(self):
        outer = tk.Frame(self.t_cfg, bg=BG)
        outer.pack(fill="both", expand=True)

        # Scroll
        canvas = tk.Canvas(outer, bg=BG, highlightthickness=0)
        vsb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        inner = tk.Frame(canvas, bg=BG)
        wid = canvas.create_window((0,0), window=inner, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(wid, width=e.width))
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)),"units"))

        wrap = tk.Frame(inner, bg=BG, padx=24, pady=20)
        wrap.pack(fill="both", expand=True)

        # ── BioVida ──
        c1 = self._section(wrap, "BIOVIDA")
        r1 = self._field_row(c1, "Pasta dos PDFs")
        self._e_pasta_bio = self._entry(r1)
        self._e_pasta_bio.pack(side="left", fill="x", expand=True, ipady=5, padx=(0,8))
        self._e_pasta_bio.insert(0, self.cfg.get("pasta_biovida",""))
        self._btn(r1,"Procurar", lambda: self._browse_dir("pasta_biovida",self._e_pasta_bio),
                  bg=DIM, fg=TEXT, padx=12, pady=5).pack(side="left")
        self._divider(c1)
        tk.Label(c1, text="  Formato esperado: 999999999-12345.pdf  →  ID de busca = 12345",
                 font=(FNT,8), bg=SURFACE, fg=MUTED, pady=8, padx=16).pack(anchor="w")

        # ── Biocloma ──
        c2 = self._section(wrap, "BIOCLOMA")
        r2a = self._field_row(c2, "Planilha .xlsx")
        self._e_plan = self._entry(r2a)
        self._e_plan.pack(side="left", fill="x", expand=True, ipady=5, padx=(0,8))
        self._e_plan.insert(0, self.cfg.get("planilha_biocloma",""))
        self._btn(r2a,"Procurar", lambda: self._browse_file("planilha_biocloma",self._e_plan),
                  bg=DIM, fg=TEXT, padx=12, pady=5).pack(side="left")
        self._divider(c2)
        r2b = self._field_row(c2, "Pasta dos PDFs")
        self._e_pasta_clo = self._entry(r2b)
        self._e_pasta_clo.pack(side="left", fill="x", expand=True, ipady=5, padx=(0,8))
        self._e_pasta_clo.insert(0, self.cfg.get("pasta_biocloma",""))
        self._btn(r2b,"Procurar", lambda: self._browse_dir("pasta_biocloma",self._e_pasta_clo),
                  bg=DIM, fg=TEXT, padx=12, pady=5).pack(side="left")
        self._divider(c2)
        tk.Label(c2, text="  IDs lidos da coluna A da planilha.",
                 font=(FNT,8), bg=SURFACE, fg=MUTED, pady=8, padx=16).pack(anchor="w")

        # ── Timing ──
        c3 = self._section(wrap, "TIMING")
        self._spins = {}
        for label, key, lo, hi, inc in [
            ("Delay entre passos (s):",    "delay_entre_passos",    0.1,15.0,0.1),
            ("Delay entre registros (s):", "delay_entre_registros", 0.5,30.0,0.5),
            ("Delay pós-click (s):",       "delay_pos_click",       0.1, 5.0,0.1),
        ]:
            self._divider(c3) if self._spins else None
            r = self._field_row(c3, label, pady=10)
            sp = tk.Spinbox(r, from_=lo, to=hi, increment=inc, width=8,
                            font=(MONO,9), bg=CARD, fg=TEXT, relief="flat",
                            buttonbackground=DIM, highlightbackground=BORDER,
                            highlightthickness=1, insertbackground=ACCENT)
            sp.pack(side="left", ipady=4)
            sp.delete(0,"end"); sp.insert(0, str(self.cfg.get(key,lo)))
            self._spins[key] = sp

        self._divider(c3)
        tk.Label(c3, text="  ⚠  Mover mouse para o canto superior esquerdo ativa FAILSAFE e para tudo.",
                 font=(FNT,8), bg=SURFACE, fg=WARN, pady=8, padx=16).pack(anchor="w")

        # ── Save ──
        save_row = tk.Frame(wrap, bg=BG, pady=16)
        save_row.pack(fill="x")
        self._btn(save_row, "  Salvar todas as configurações  ",
                  self._save_all, padx=24, pady=10).pack(side="left")

    def _browse_dir(self, key, entry):
        p = filedialog.askdirectory()
        if p: entry.delete(0,"end"); entry.insert(0,p); self.cfg[key]=p

    def _browse_file(self, key, entry):
        p = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx")])
        if p: entry.delete(0,"end"); entry.insert(0,p); self.cfg[key]=p

    # ── ABA PASSOS ────────────────────────────────────────────────────────────
    def _build_steps(self, tab, modelo, sv_list):
        default = DEFAULT_STEPS
        steps   = self.cfg.get(f"steps_{modelo}", default)

        tab.grid_rowconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=0)
        tab.grid_columnconfigure(0, weight=1)

        # Área scrollável
        top = tk.Frame(tab, bg=BG)
        top.grid(row=0, column=0, sticky="nsew")
        top.grid_rowconfigure(0, weight=1)
        top.grid_columnconfigure(0, weight=1)

        canvas = tk.Canvas(top, bg=BG, highlightthickness=0)
        vsb = ttk.Scrollbar(top, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.grid(row=0, column=1, sticky="ns")
        canvas.grid(row=0, column=0, sticky="nsew")
        inner = tk.Frame(canvas, bg=BG)
        wid = canvas.create_window((0,0), window=inner, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(wid, width=e.width))
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)),"units"))

        wrap = tk.Frame(inner, bg=BG, padx=24, pady=16)
        wrap.pack(fill="both", expand=True)

        # Info variáveis
        info = tk.Frame(wrap, bg=TEAL)
        info.pack(fill="x", pady=(0,12))
        info.config(highlightbackground=TEAL2, highlightthickness=1)
        tk.Label(info, text=f"  Variáveis: {'   '.join(VARIAVEIS)}",
                 font=(FNT,9,"bold"), bg=TEAL, fg=TEXT, pady=7, padx=4).pack(side="left")
        tk.Label(info, text="  |  click_type: clica + Ctrl+A + digita  |  enter/hotkey com X=Y=0: sem clique",
                 font=(FNT,8), bg=TEAL, fg=TEXT, pady=7).pack(side="left")

        # Header colunas
        hdr = tk.Frame(wrap, bg=CARD2, pady=6, padx=10)
        hdr.pack(fill="x")
        for txt,w in [("#",3),("✓",2),("Nome do Passo",24),("Ação",17),("X",7),("Y",7),("Texto/Variável",0)]:
            tk.Label(hdr, text=txt, font=(FNT,8,"bold"), bg=CARD2,
                     fg=MUTED, width=w, anchor="w").pack(side="left", padx=3)

        # Linhas de steps
        steps_frame = tk.Frame(wrap, bg=BG)
        steps_frame.pack(fill="x", pady=(2,0))

        sv_list.clear()
        for i,s in enumerate(steps):
            self._step_row(steps_frame, sv_list, i, s)

        # Barra inferior fixa (sempre visível)
        bottom = tk.Frame(tab, bg=SURFACE)
        bottom.grid(row=1, column=0, sticky="ew")
        bottom.config(highlightbackground=BORDER, highlightthickness=1)
        bot_row = tk.Frame(bottom, bg=SURFACE, padx=20, pady=10)
        bot_row.pack(fill="x")

        self._btn(bot_row, "+ Adicionar Passo",
                  lambda sf=steps_frame, sv=sv_list: self._step_row(
                      sf, sv, len(sv),
                      {"nome":"Novo passo","x":0,"y":0,"acao":"click","texto":"","ativo":True}
                  ), bg=DIM, fg=TEXT, padx=16, pady=7).pack(side="left", padx=(0,8))

        self._btn(bot_row, "Salvar Passos",
                  lambda mo=modelo: self._save_steps(mo),
                  padx=16, pady=7).pack(side="left")

        # Capturador inline
        sep = tk.Frame(bot_row, bg=BORDER, width=1)
        sep.pack(side="left", fill="y", padx=16)

        tk.Label(bot_row, text="Capturar mouse →", font=(FNT,9),
                 bg=SURFACE, fg=MUTED).pack(side="left")
        tk.Label(bot_row, text=" X:", font=(FNT,9,"bold"), bg=SURFACE, fg=TEXT).pack(side="left")
        ex = self._entry(bot_row, width=6, bg=BG); ex.pack(side="left", ipady=4, padx=(2,6))
        tk.Label(bot_row, text="Y:", font=(FNT,9,"bold"), bg=SURFACE, fg=TEXT).pack(side="left")
        ey = self._entry(bot_row, width=6, bg=BG); ey.pack(side="left", ipady=4, padx=(2,10))
        self._btn(bot_row, "📍 Capturar",
                  lambda ex=ex,ey=ey: self._capturar(ex,ey),
                  bg=TEAL, fg=TEXT, padx=12, pady=6).pack(side="left")

    def _step_row(self, frame, sv_list, idx, step):
        bg = CARD if idx%2==0 else CARD2
        cont = tk.Frame(frame, bg=BG)
        cont.pack(fill="x", pady=1)
        row = tk.Frame(cont, bg=bg, padx=10, pady=7)
        row.pack(fill="x")

        # Número
        tk.Label(row, text=f"{len(sv_list)+1:02d}", font=(MONO,9,"bold"),
                 bg=bg, fg=GOLD, width=3).pack(side="left")

        # Ativo
        v_ativo = tk.BooleanVar(value=step.get("ativo",True))
        tk.Checkbutton(row, variable=v_ativo, bg=bg, activebackground=bg,
                       selectcolor=BG, cursor="hand2").pack(side="left", padx=(0,4))

        # Nome
        v_nome = tk.StringVar(value=step.get("nome",""))
        tk.Entry(row, textvariable=v_nome, font=(FNT,9), bg=BG, fg=TEXT,
                 width=24, relief="flat", highlightbackground=BORDER,
                 highlightthickness=1, highlightcolor=ACCENT,
                 insertbackground=ACCENT).pack(side="left", ipady=4, padx=(0,6))

        # Ação
        v_acao = tk.StringVar(value=step.get("acao","click"))
        cb = ttk.Combobox(row, textvariable=v_acao, values=ACOES,
                          font=(FNT,9), width=16, state="readonly")
        cb.pack(side="left", ipady=2, padx=(0,6))

        # X
        tk.Label(row, text="X", font=(FNT,8), bg=bg, fg=