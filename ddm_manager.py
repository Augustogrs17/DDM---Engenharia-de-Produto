"""
DDM Manager - Metalfrio Solutions
v3.0 — Semana completa + checkboxes + PDF via Word
"""

# ─── DPI Awareness ───────────────────────────────────────────────────────────
import ctypes, sys
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

import os, copy, json, shutil, hashlib, datetime, subprocess, tkinter as tk
from tkinter import messagebox, filedialog, simpledialog
from pathlib import Path

# ─── Paleta ───────────────────────────────────────────────────────────────────
AZUL   = "#0B1D2F"
VERDE  = "#1CBABE"
CINZA1 = "#1A2E42"
CINZA2 = "#243A52"
CINZA3 = "#8EAABF"
TEXTO  = "#E8F0F8"
WARN   = "#F5A623"
ERR    = "#E74C3C"
OK_C   = "#2ECC71"
GOLD   = "#F0C040"
SEL_BG = "#0F3050"

F_TITLE  = ("Segoe UI", 17, "bold")
F_LBL    = ("Segoe UI",  9, "bold")
F_FIEL   = ("Segoe UI", 10)
F_BTN    = ("Segoe UI", 11, "bold")
F_BTN_SM = ("Segoe UI",  9, "bold")
F_LOG    = ("Consolas",  9)
F_SMALL  = ("Segoe UI",  8)

# ─── Mapeamento ───────────────────────────────────────────────────────────────
# Número da pasta → (weekday python, nome PT)
PASTA_DIA = {
    "2": (0, "Segunda-feira"),
    "3": (1, "Terça-feira"),
    "4": (2, "Quarta-feira"),
    "5": (3, "Quinta-feira"),
    "6": (4, "Sexta-feira"),
    "7": (5, "Sábado"),
}
DIA_MAP  = {0:"2", 1:"3", 2:"4", 3:"5", 4:"6", 5:"7", 6:"7"}
DIAS_PT  = ["Segunda","Terça","Quarta","Quinta","Sexta","Sábado","Domingo"]
MES_MAP  = {
    1:"1 - JANEIRO",  2:"2 - FEVEREIRO", 3:"3 - MARÇO",
    4:"4 - ABRIL",    5:"5 - MAIO",      6:"6 - JUNHO",
    7:"7 - JULHO",    8:"8 - AGOSTO",    9:"9 - SETEMBRO",
    10:"10 - OUTUBRO",11:"11 - NOVEMBRO",12:"12 - DEZEMBRO",
}

PARTICIPANTES = [
    ("10347", "ADRIANO GARCIA"),
    ("63363", "AGATHA LAÍS SALGUEIRO MAROTTI"),
    ("63051", "AMANDA FLORÊNCIO"),
    ("63175", "ANA EDUARDA SANO SILVA"),
    ("10934", "ANDRÉ LUIZ HIGA FREITAS"),
    ("63538", "AUGUSTO GRAÇA SILVESTRIN"),
    ("63231", "BARBARA RAMIRES IANHES"),
    ("62334", "CLEYTON FARIA ASSIS"),
    ("62397", "JOÃO VITOR RODRIGUES SOUZA"),
    ("184",   "JOEL BORGES DE CAMPOS JUNIOR"),
    ("61510", "KAROLINE PEDROSO SANTOS"),
    ("10437", "LETICIA FERNANDES GONÇALVES"),
    ("63186", "LUDIMILA SANTOS SOUZA"),
    ("63516", "LUIS GUILHERME DUENHAS SILVA"),
    ("63391", "RENAN DA SILVA MANTOVANI"),
    ("63999", "RODRIGO ARIAS SILVA"),
    ("11461", "VANESSA DIAS DA SILVA"),
    ("9018",  "WAGNER JUNIOR ALMEIDA"),
]

DEFAULTS = {
    "raiz":        r"Q:\Transferencia\DDM 2026",
    "setor":       "Engenharia Industrial",
    "subarea":     "Engenharia de Produto",
    "turno":       "ADM",
    "facilitador": "Leitura Individual",
}

# ─── Auth ─────────────────────────────────────────────────────────────────────
ADMIN_RE   = "184"
USERS_FILE = Path(sys.executable if getattr(sys, "frozen", False)
                  else __file__).parent / "ddm_users.json"

def _hash(pw): return hashlib.sha256(pw.encode()).hexdigest()

def _load_users():
    if USERS_FILE.exists():
        try: return json.loads(USERS_FILE.read_text(encoding="utf-8"))
        except: pass
    users = {"184": {"nome": "JOEL BORGES DE CAMPOS JUNIOR",
                     "pw": _hash("metalfrio"), "admin": True}}
    _save_users(users)
    return users

def _save_users(u):
    USERS_FILE.write_text(json.dumps(u, ensure_ascii=False, indent=2),
                          encoding="utf-8")

# ════════════════════════════════════════════════════════════════════════════
# Detecção de DDMs — semana completa
# ════════════════════════════════════════════════════════════════════════════

def semana_atual(data: datetime.date):
    """Retorna (segunda, sexta) da semana corrente."""
    seg = data - datetime.timedelta(days=data.weekday())
    sex = seg + datetime.timedelta(days=4)
    return seg, sex

def encontrar_ddm_pasta(pasta_dia_path: Path, data: datetime.date):
    """
    Dentro de pasta_dia/MES/, encontra o .docx cujo prefixo DD-MM
    pertence à semana de `data`.
    Retorna (caminho_docx | None, erro | "")
    """
    nome_mes = MES_MAP[data.month]
    pasta_mes = None
    for item in pasta_dia_path.iterdir():
        if item.is_dir() and (item.name.upper() == nome_mes.upper()
                              or item.name.startswith(str(data.month) + " -")):
            pasta_mes = item
            break
    if not pasta_mes:
        return None, f"Pasta '{nome_mes}' não encontrada"

    seg, sex = semana_atual(data)
    docxs = sorted(pasta_mes.glob("*.docx"))
    alvo, melhor = None, None
    for f in docxs:
        try:
            dd, mm = int(f.name[:2]), int(f.name[3:5])
            d_arq = datetime.date(data.year, mm, dd)
            # pertence à semana?
            if seg <= d_arq <= sex:
                delta = abs((d_arq - data).days)
                if melhor is None or delta < melhor:
                    alvo, melhor = f, delta
        except: continue
    if alvo:
        return alvo, ""
    return None, f"Sem DDM na semana {seg:%d/%m}–{sex:%d/%m}"

def varrer_semana(raiz: str, data: datetime.date) -> list[dict]:
    """
    Varre todas as pastas de dia dentro de raiz e retorna lista de dicts:
    { num, dia_nome, pasta_tema, docx_path, tema, data_ddm, erro }
    ordenados por dia da semana.
    """
    raiz_p = Path(raiz)
    resultado = []

    if not raiz_p.exists():
        return resultado

    for pasta in sorted(raiz_p.iterdir()):
        if not pasta.is_dir():
            continue
        # Extrai número da pasta (ex: "2" de "2 - Segunda-feira...")
        partes = pasta.name.split(" - ", 1)
        if len(partes) < 2:
            continue
        num = partes[0].strip()
        if num not in PASTA_DIA:
            continue

        weekday_pasta, dia_nome = PASTA_DIA[num]
        # Data representativa da pasta nesta semana
        seg, _ = semana_atual(data)
        data_pasta = seg + datetime.timedelta(days=weekday_pasta
                          if weekday_pasta <= 4 else 5)

        docx, erro = encontrar_ddm_pasta(pasta, data)
        tema = extrair_tema(docx.name) if docx else "—"
        data_ddm = extrair_data_ddm(docx.name if docx else "", data.year)

        resultado.append({
            "num":        int(num),
            "dia_nome":   dia_nome,
            "pasta_tema": partes[1].strip() if len(partes) > 1 else pasta.name,
            "docx_path":  docx,
            "tema":       tema,
            "data_ddm":   data_ddm,
            "erro":       erro,
        })

    resultado.sort(key=lambda x: x["num"])
    return resultado

def extrair_tema(nome_arq):
    stem = Path(nome_arq).stem
    p = stem.split(" ", 1)
    return p[1].strip() if len(p) == 2 else stem

def extrair_data_ddm(nome_arq, ano):
    try:
        dd, mm = int(nome_arq[:2]), int(nome_arq[3:5])
        return f"{dd:02d}/{mm:02d}/{ano}"
    except: return ""

# ════════════════════════════════════════════════════════════════════════════
# Processamento DOCX
# ════════════════════════════════════════════════════════════════════════════

def _cells_unicas(linha):
    seen, res = set(), []
    for c in linha.cells:
        cid = id(c._tc)
        if cid not in seen:
            seen.add(cid); res.append(c)
    return res

def _run_apos_label(para, valor):
    from docx.oxml.ns import qn
    from lxml import etree
    runs = para.runs
    if not runs: return
    for r in runs[1:]:
        r._element.getparent().remove(r._element)
    novo = copy.deepcopy(runs[0]._element)
    t = novo.find(qn("w:t"))
    if t is None: t = etree.SubElement(novo, qn("w:t"))
    t.text = " " + valor
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    para._element.append(novo)

def _escrever_celula(cell, valor, alinhar_h="center"):
    from docx.oxml.ns import qn
    from lxml import etree
    tcp = cell._tc.find(qn("w:tcPr"))
    if tcp is None: tcp = etree.SubElement(cell._tc, qn("w:tcPr"))
    val_el = tcp.find(qn("w:vAlign"))
    if val_el is None: val_el = etree.SubElement(tcp, qn("w:vAlign"))
    val_el.set(qn("w:val"), "center")

    para = cell.paragraphs[0] if cell.paragraphs else None
    if not para: return
    for r in para.runs:
        r._element.getparent().remove(r._element)
    ppr = para._element.find(qn("w:pPr"))
    if ppr is None:
        ppr = etree.SubElement(para._element, qn("w:pPr"))
        para._element.insert(0, ppr)
    jc = ppr.find(qn("w:jc"))
    if jc is None: jc = etree.SubElement(ppr, qn("w:jc"))
    jc.set(qn("w:val"), "center" if alinhar_h == "center" else "left")

    r_el = etree.SubElement(para._element, qn("w:r"))
    rpr  = etree.SubElement(r_el, qn("w:rPr"))
    fnts = etree.SubElement(rpr, qn("w:rFonts"))
    for a, v in [("w:ascii","Arial"),("w:hAnsi","Arial"),
                 ("w:eastAsia","Batang"),("w:cs","Arial")]:
        fnts.set(qn(a), v)
    for tag in ("w:sz","w:szCs"):
        e = etree.SubElement(rpr, qn(tag)); e.set(qn("w:val"), "18")
    t = etree.SubElement(r_el, qn("w:t"))
    t.text = valor
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

def _data_preenchida(doc) -> bool:
    for linha in doc.tables[0].rows:
        for cell in _cells_unicas(linha):
            for para in cell.paragraphs:
                txt = para.text.strip()
                if txt.startswith("DATA:") and txt[5:].strip():
                    return True
    return False

def processar_ddm(src: Path, dst: Path, campos: dict) -> dict:
    from docx import Document
    LABELS = {
        "SETOR/LINHA:": campos["setor"],
        "TURNO:":       campos["turno"],
        "SUBÁREA:":     campos["subarea"],
        "FACILITADOR:": campos["facilitador"],
    }
    doc = Document(str(src))
    data_no_doc = _data_preenchida(doc)
    data_extra = {} if data_no_doc else {"DATA:": campos["data"]}
    tabela = doc.tables[0]
    part_idx = 0
    info = {"cabecalho": [], "nomes": 0, "data_preenchida": not data_no_doc}

    for linha in tabela.rows:
        cells_u = _cells_unicas(linha)
        for cell in cells_u:
            for para in cell.paragraphs:
                txt = para.text.strip()
                for label, valor in {**LABELS, **data_extra}.items():
                    if txt.startswith(label) and txt[len(label):].strip() == "":
                        _run_apos_label(para, valor)
                        info["cabecalho"].append(label)
        if not cells_u: continue
        try: int(cells_u[0].text.strip())
        except ValueError: continue
        if part_idx < len(PARTICIPANTES) and len(cells_u) >= 3:
            re_v, nome_v = PARTICIPANTES[part_idx]
            _escrever_celula(cells_u[1], re_v,   alinhar_h="center")
            _escrever_celula(cells_u[2], nome_v, alinhar_h="left")
            info["nomes"] += 1
        part_idx += 1

    doc.save(str(dst))
    return info

# ─── Conversão PDF ────────────────────────────────────────────────────────────

def converter_pdf_word(caminho_docx: Path) -> Path:
    """Converte DOCX → PDF usando Microsoft Word via win32com (sem LibreOffice)."""
    try:
        import win32com.client
    except ImportError:
        raise RuntimeError(
            "pywin32 não encontrado.\n"
            "Adicione 'pywin32' nas dependências e recompile."
        )
    caminho_docx = Path(caminho_docx).resolve()
    caminho_pdf  = caminho_docx.with_suffix(".pdf")
    word = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(str(caminho_docx))
        doc.SaveAs(str(caminho_pdf), FileFormat=17)  # 17 = wdFormatPDF
        doc.Close(False)
    finally:
        if word:
            word.Quit()
    return caminho_pdf

# ════════════════════════════════════════════════════════════════════════════
# Login
# ════════════════════════════════════════════════════════════════════════════

class LoginWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("DDM Manager — Login")
        self.configure(bg=AZUL)
        self.resizable(False, False)
        self._users = _load_users()
        self._logged_user = None
        self._build()
        self._center(400, 340)

    def _center(self, w, h):
        self.update_idletasks()
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _build(self):
        hdr = tk.Frame(self, bg=VERDE, height=70)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        tk.Label(hdr, text="❄  DDM Manager", font=F_TITLE,
                 bg=VERDE, fg=AZUL).pack(pady=18)

        frm = tk.Frame(self, bg=AZUL, padx=36)
        frm.pack(fill="both", expand=True, pady=16)

        tk.Label(frm, text="RE (Registro de Empregado)", font=F_LBL,
                 bg=AZUL, fg=VERDE, anchor="w").pack(fill="x")
        self._v_re = tk.StringVar()
        e_re = tk.Entry(frm, textvariable=self._v_re, font=F_FIEL,
                        bg=CINZA2, fg=TEXTO, insertbackground=VERDE,
                        relief="flat", bd=0)
        e_re.pack(fill="x", ipady=7, ipadx=6, pady=(2,14))
        e_re.focus()

        tk.Label(frm, text="Senha", font=F_LBL,
                 bg=AZUL, fg=VERDE, anchor="w").pack(fill="x")
        self._v_pw = tk.StringVar()
        e_pw = tk.Entry(frm, textvariable=self._v_pw, font=F_FIEL,
                        bg=CINZA2, fg=TEXTO, insertbackground=VERDE,
                        relief="flat", bd=0, show="●")
        e_pw.pack(fill="x", ipady=7, ipadx=6, pady=(2,20))
        e_pw.bind("<Return>", lambda _: self._login())

        self._lbl_err = tk.Label(frm, text="", font=F_LBL, bg=AZUL, fg=ERR)
        self._lbl_err.pack()
        tk.Button(frm, text="Entrar →", font=F_BTN,
                  bg=VERDE, fg=AZUL, relief="flat", cursor="hand2",
                  padx=20, command=self._login).pack(pady=(4,0))

    def _login(self):
        re = self._v_re.get().strip()
        pw = self._v_pw.get()
        user = self._users.get(re)
        if not user or user["pw"] != _hash(pw):
            self._lbl_err.config(text="RE ou senha incorretos.")
            self._v_pw.set(""); return
        self._logged_user = {"re": re, **user}
        self.destroy()

    def get_user(self): return self._logged_user

# ════════════════════════════════════════════════════════════════════════════
# Painel Admin
# ════════════════════════════════════════════════════════════════════════════

class AdminPanel(tk.Toplevel):
    def __init__(self, parent, users):
        super().__init__(parent)
        self.title("Administração de Usuários")
        self.configure(bg=AZUL)
        self.resizable(False, False)
        self._users = users
        self._build()
        self._center(580, 500)
        self.grab_set()

    def _center(self, w, h):
        self.update_idletasks()
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _build(self):
        hdr = tk.Frame(self, bg=GOLD, height=46)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        tk.Label(hdr, text="⚙  Administração de Usuários", font=F_BTN,
                 bg=GOLD, fg=AZUL).pack(side="left", padx=16, pady=10)

        lista_frm = tk.Frame(self, bg=AZUL, padx=16, pady=10)
        lista_frm.pack(fill="both", expand=True)
        tk.Label(lista_frm, text="Usuários cadastrados", font=F_LBL,
                 bg=AZUL, fg=VERDE).pack(anchor="w")

        hdr_row = tk.Frame(lista_frm, bg=CINZA1)
        hdr_row.pack(fill="x", pady=(4,0))
        for txt, w in [("RE",8),("Nome",36),("Admin",6),("Ações",8)]:
            tk.Label(hdr_row, text=txt, font=F_LBL, bg=CINZA1,
                     fg=CINZA3, width=w, anchor="w").pack(side="left", padx=4, pady=3)

        self._lista_frm = tk.Frame(lista_frm, bg=AZUL)
        self._lista_frm.pack(fill="both", expand=True, pady=4)
        self._refresh_lista()

        sep = tk.Frame(self, bg=CINZA1, height=2); sep.pack(fill="x")

        novo_frm = tk.Frame(self, bg=AZUL, padx=16, pady=10)
        novo_frm.pack(fill="x")
        tk.Label(novo_frm, text="Adicionar usuário", font=F_LBL,
                 bg=AZUL, fg=VERDE).grid(row=0,column=0,columnspan=4,sticky="w",pady=(0,6))

        for col,(lbl,w) in enumerate([("RE",8),("Nome completo",26),("Senha",12)]):
            tk.Label(novo_frm, text=lbl, font=F_LBL, bg=AZUL,
                     fg=CINZA3).grid(row=1,column=col,sticky="w",padx=4)

        self._v_new_re   = tk.StringVar()
        self._v_new_nome = tk.StringVar()
        self._v_new_pw   = tk.StringVar()

        tk.Entry(novo_frm, textvariable=self._v_new_re,   font=F_FIEL,
                 bg=CINZA2, fg=TEXTO, insertbackground=VERDE,
                 relief="flat", bd=0, width=9
                 ).grid(row=2,column=0,padx=4,ipady=5)
        tk.Entry(novo_frm, textvariable=self._v_new_nome, font=F_FIEL,
                 bg=CINZA2, fg=TEXTO, insertbackground=VERDE,
                 relief="flat", bd=0, width=28
                 ).grid(row=2,column=1,padx=4,ipady=5)
        tk.Entry(novo_frm, textvariable=self._v_new_pw,   font=F_FIEL,
                 bg=CINZA2, fg=TEXTO, insertbackground=VERDE,
                 relief="flat", bd=0, width=14, show="●"
                 ).grid(row=2,column=2,padx=4,ipady=5)
        tk.Button(novo_frm, text="＋ Adicionar", font=F_BTN_SM,
                  bg=VERDE, fg=AZUL, relief="flat", cursor="hand2",
                  command=self._adicionar).grid(row=2,column=3,padx=8)

        self._lbl_msg = tk.Label(novo_frm, text="", font=F_LBL, bg=AZUL, fg=OK_C)
        self._lbl_msg.grid(row=3,column=0,columnspan=4,sticky="w",pady=4)

        rod = tk.Frame(self, bg=CINZA1, height=46)
        rod.pack(fill="x", side="bottom"); rod.pack_propagate(False)
        tk.Button(rod, text="Fechar", font=F_BTN_SM,
                  bg=CINZA1, fg=CINZA3, relief="flat", cursor="hand2",
                  command=self.destroy).pack(side="right", padx=12, pady=10)

    def _refresh_lista(self):
        for w in self._lista_frm.winfo_children(): w.destroy()
        for re, info in self._users.items():
            row = tk.Frame(self._lista_frm, bg=CINZA2)
            row.pack(fill="x", pady=1)
            tk.Label(row, text=re, font=F_FIEL, bg=CINZA2,
                     fg=TEXTO, width=8, anchor="w").pack(side="left", padx=4)
            tk.Label(row, text=info["nome"][:38], font=F_FIEL, bg=CINZA2,
                     fg=TEXTO, width=36, anchor="w").pack(side="left")
            tk.Label(row, text="★" if info.get("admin") else "", font=F_FIEL,
                     bg=CINZA2, fg=GOLD, width=5).pack(side="left")
            if not info.get("admin"):
                tk.Button(row, text="✕", font=F_BTN_SM,
                          bg=CINZA2, fg=ERR, relief="flat", cursor="hand2",
                          command=lambda r=re: self._remover(r)).pack(side="left", padx=4)
            tk.Button(row, text="🔑", font=F_BTN_SM,
                      bg=CINZA2, fg=WARN, relief="flat", cursor="hand2",
                      command=lambda r=re: self._redefinir_senha(r)).pack(side="left", padx=2)

    def _adicionar(self):
        re   = self._v_new_re.get().strip()
        nome = self._v_new_nome.get().strip().upper()
        pw   = self._v_new_pw.get()
        if not re or not nome or not pw:
            self._msg("Preencha RE, nome e senha.", ERR); return
        if re in self._users:
            self._msg(f"RE {re} já cadastrado.", ERR); return
        self._users[re] = {"nome": nome, "pw": _hash(pw), "admin": False}
        _save_users(self._users)
        self._v_new_re.set(""); self._v_new_nome.set(""); self._v_new_pw.set("")
        self._refresh_lista()
        self._msg(f"{re} — {nome} adicionado.", OK_C)

    def _remover(self, re):
        if not messagebox.askyesno("Confirmar",
                f"Remover {re} — {self._users[re]['nome']}?", parent=self): return
        del self._users[re]; _save_users(self._users)
        self._refresh_lista(); self._msg("Usuário removido.", WARN)

    def _redefinir_senha(self, re):
        nova = simpledialog.askstring("Redefinir senha",
               f"Nova senha para {self._users[re]['nome']}:", show="●", parent=self)
        if not nova: return
        self._users[re]["pw"] = _hash(nova)
        _save_users(self._users)
        self._msg("Senha redefinida.", OK_C)

    def _msg(self, txt, cor=OK_C):
        self._lbl_msg.config(text=txt, fg=cor)

# ════════════════════════════════════════════════════════════════════════════
# Janela Principal
# ════════════════════════════════════════════════════════════════════════════

class MainWindow(tk.Tk):
    def __init__(self, user, users):
        super().__init__()
        self._user  = user
        self._users = users
        self.title("DDM Manager — Metalfrio")
        self.configure(bg=AZUL)
        self.resizable(True, True)
        self.minsize(900, 620)
        # Lista de dicts da semana + BooleanVar para checkbox
        self._semana: list[dict] = []
        self._checks: list[tk.BooleanVar] = []
        self._build()
        self.after(100, self._varrer)

    def _center(self, w, h):
        self.geometry(f"{w}x{h}")
        self.update_idletasks()
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _build(self):
        self._center(1060, 720)

        # ══ 1. HEADER ════════════════════════════════════════════════════════
        hdr = tk.Frame(self, bg=VERDE, height=58)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        tk.Label(hdr, text="❄  DDM Manager", font=F_TITLE,
                 bg=VERDE, fg=AZUL).pack(side="left", padx=18, pady=10)

        user_frm = tk.Frame(hdr, bg=VERDE)
        user_frm.pack(side="right", padx=12)
        nome_curto = self._user["nome"].split()[0].capitalize()
        admin_tag  = "  ★ Admin" if self._user.get("admin") else ""
        tk.Label(user_frm, text=f"👤 {nome_curto}{admin_tag}",
                 font=("Segoe UI",9,"bold"), bg=VERDE, fg=AZUL).pack(anchor="e")
        self._lbl_clock = tk.Label(user_frm, text="",
                                   font=("Segoe UI",9), bg=VERDE, fg=AZUL)
        self._lbl_clock.pack(anchor="e")
        self._tick()

        if self._user.get("admin"):
            tk.Button(hdr, text="⚙ Usuários", font=F_BTN_SM,
                      bg=GOLD, fg=AZUL, relief="flat", cursor="hand2",
                      command=self._abrir_admin).pack(side="right", padx=8, pady=14)

        # ══ 2. RODAPÉ (pack ANTES do corpo → sempre visível) ════════════════
        rod = tk.Frame(self, bg=CINZA1, height=72)
        rod.pack(fill="x", side="bottom")
        rod.pack_propagate(False)

        tk.Button(rod, text="📄   Gerar DOCX selecionados",
                  font=("Segoe UI", 12, "bold"),
                  bg=CINZA2, fg=VERDE, relief="flat", cursor="hand2",
                  padx=18, command=self._gerar_docx_sel
                  ).pack(side="left", padx=10, pady=14, ipady=5)

        tk.Button(rod, text="📑   Gerar PDF + Abrir selecionados",
                  font=("Segoe UI", 12, "bold"),
                  bg=VERDE, fg=AZUL, relief="flat", cursor="hand2",
                  padx=18, command=self._gerar_pdf_sel
                  ).pack(side="left", padx=4, pady=14, ipady=5)

        tk.Button(rod, text="✖   Sair",
                  font=("Segoe UI", 12, "bold"),
                  bg=CINZA1, fg=ERR, relief="flat", cursor="hand2",
                  padx=18, command=self.destroy
                  ).pack(side="right", padx=10, pady=14, ipady=5)

        # ══ 3. BARRA PASTA RAIZ ══════════════════════════════════════════════
        raiz_bar = tk.Frame(self, bg=CINZA1, pady=0)
        raiz_bar.pack(fill="x")

        tk.Label(raiz_bar, text="📁  Pasta raiz DDM 2026", font=F_LBL,
                 bg=CINZA1, fg=VERDE).pack(side="left", padx=(12,6), pady=8)

        self._vars: dict[str, tk.StringVar] = {}
        v_raiz = tk.StringVar(value=DEFAULTS.get("raiz", ""))
        self._vars["raiz"] = v_raiz

        tk.Entry(raiz_bar, textvariable=v_raiz, font=F_FIEL,
                 bg=CINZA2, fg=TEXTO, insertbackground=VERDE,
                 relief="flat", bd=0
                 ).pack(side="left", fill="x", expand=True, ipady=6, ipadx=6, pady=6)

        tk.Button(raiz_bar, text="  …  ", font=("Segoe UI",10,"bold"),
                  bg=VERDE, fg=AZUL, relief="flat", cursor="hand2",
                  command=self._browse
                  ).pack(side="left", padx=(4,4), ipady=5, pady=6)

        tk.Button(raiz_bar, text="🔄  Atualizar", font=F_BTN_SM,
                  bg=CINZA2, fg=VERDE, relief="flat", cursor="hand2",
                  command=self._varrer
                  ).pack(side="left", padx=(0,12), ipady=5, ipadx=8, pady=6)

        # ══ 4. CORPO (fill+expand → ocupa o espaço restante) ════════════════
        body = tk.Frame(self, bg=AZUL)
        body.pack(fill="both", expand=True, padx=14, pady=8)

        # ── Coluna esquerda: campos ──
        esq = tk.Frame(body, bg=AZUL, width=300)
        esq.pack(side="left", fill="y", padx=(0,10))
        esq.pack_propagate(False)

        def campo(key, label):
            f = tk.Frame(esq, bg=AZUL); f.pack(fill="x", pady=3)
            tk.Label(f, text=label, font=F_LBL, bg=AZUL,
                     fg=VERDE, anchor="w").pack(fill="x")
            v = tk.StringVar(value=DEFAULTS.get(key, ""))
            self._vars[key] = v
            tk.Entry(f, textvariable=v, font=F_FIEL,
                     bg=CINZA2, fg=TEXTO, insertbackground=VERDE,
                     relief="flat", bd=0
                     ).pack(fill="x", ipady=6, ipadx=6)

        campo("setor",       "🏭  Setor / Linha")
        campo("subarea",     "🔹  Subárea")
        campo("turno",       "⏰  Turno")
        campo("facilitador", "👤  Facilitador")

        tk.Label(esq, text=f"👥  {len(PARTICIPANTES)} participantes carregados",
                 font=F_LBL, bg=AZUL, fg=OK_C).pack(anchor="w", pady=(12,0))

        # ── Coluna central: lista DDMs ──
        meio = tk.Frame(body, bg=AZUL, width=370)
        meio.pack(side="left", fill="y", padx=(0,10))
        meio.pack_propagate(False)

        seg, sex = semana_atual(datetime.date.today())
        tk.Label(meio, text=f"DDMs da semana  {seg:%d/%m} – {sex:%d/%m}",
                 font=F_LBL, bg=AZUL, fg=VERDE).pack(anchor="w", pady=(0,4))

        hdr_list = tk.Frame(meio, bg=CINZA1)
        hdr_list.pack(fill="x")
        tk.Label(hdr_list, text="✔", font=F_SMALL, bg=CINZA1,
                 fg=CINZA3, width=3).pack(side="left", padx=4, pady=3)
        tk.Label(hdr_list, text="Dia", font=F_SMALL, bg=CINZA1,
                 fg=CINZA3, width=12, anchor="w").pack(side="left")
        tk.Label(hdr_list, text="Tema / arquivo", font=F_SMALL,
                 bg=CINZA1, fg=CINZA3, anchor="w").pack(side="left", padx=4)

        self._lista_frm = tk.Frame(meio, bg=AZUL)
        self._lista_frm.pack(fill="both", expand=True, pady=2)

        sel_frm = tk.Frame(meio, bg=AZUL)
        sel_frm.pack(fill="x", pady=(6,0))
        tk.Button(sel_frm, text="Selecionar tudo", font=F_SMALL,
                  bg=CINZA2, fg=VERDE, relief="flat", cursor="hand2",
                  command=lambda: [v.set(True) for v in self._checks]
                  ).pack(side="left", padx=(0,4), ipady=4, ipadx=8)
        tk.Button(sel_frm, text="Limpar seleção", font=F_SMALL,
                  bg=CINZA2, fg=CINZA3, relief="flat", cursor="hand2",
                  command=lambda: [v.set(False) for v in self._checks]
                  ).pack(side="left", ipady=4, ipadx=8)

        # ── Coluna direita: log ──
        dir_ = tk.Frame(body, bg=AZUL)
        dir_.pack(side="right", fill="both", expand=True)
        tk.Label(dir_, text="Log de execução", font=F_LBL,
                 bg=AZUL, fg=VERDE).pack(anchor="w")
        self._log_w = tk.Text(dir_, font=F_LOG, bg=CINZA1, fg=TEXTO,
                              relief="flat", state="disabled", wrap="word")
        self._log_w.pack(fill="both", expand=True)

    def _tick(self):
        agora = datetime.datetime.now()
        self._lbl_clock.config(
            text=f"{DIAS_PT[agora.weekday()]}, {agora.strftime('%d/%m/%Y  %H:%M:%S')}"
        )
        self.after(1000, self._tick)

    def _browse(self):
        p = filedialog.askdirectory(title="Pasta raiz DDM 2026")
        if p:
            self._vars["raiz"].set(p); self._varrer()

    def _abrir_admin(self):
        AdminPanel(self, self._users)

    # ── Varredura da semana ───────────────────────────────────────────────
    def _varrer(self):
        self._log_clear()
        raiz = self._vars["raiz"].get().strip()
        hoje = datetime.date.today()
        seg, sex = semana_atual(hoje)
        self._log(f"Semana: {seg:%d/%m/%Y} – {sex:%d/%m/%Y}")
        self._log(f"Raiz: {raiz}\n")

        self._semana = varrer_semana(raiz, hoje)
        self._rebuild_lista()

        ok  = [d for d in self._semana if d["docx_path"]]
        err = [d for d in self._semana if not d["docx_path"]]
        self._log(f"✔ {len(ok)} DDM(s) encontrado(s)")
        for d in ok:
            self._log(f"  {d['dia_nome']}: {d['docx_path'].name}", OK_C)
        if err:
            self._log(f"\n⚠ {len(err)} pasta(s) sem DDM desta semana:", WARN)
            for d in err:
                self._log(f"  {d['dia_nome']}: {d['erro']}", WARN)

    def _rebuild_lista(self):
        for w in self._lista_frm.winfo_children():
            w.destroy()
        self._checks = []

        hoje = datetime.date.today()

        for idx, d in enumerate(self._semana):
            tem_ddm = d["docx_path"] is not None
            v = tk.BooleanVar(value=tem_ddm)
            self._checks.append(v)

            eh_hoje = (DIA_MAP.get(hoje.weekday()) == str(d["num"]))
            bg     = SEL_BG if eh_hoje else CINZA2
            fg_dia = VERDE  if eh_hoje else TEXTO

            row = tk.Frame(self._lista_frm, bg=bg, pady=2, cursor="hand2" if tem_ddm else "")
            row.pack(fill="x", pady=1)

            # ── Checkbox desenhado manualmente (evita bug de cor no Windows) ──
            cv = tk.Canvas(row, width=18, height=18, bg=bg,
                           highlightthickness=0, cursor="hand2" if tem_ddm else "")
            cv.pack(side="left", padx=8, pady=4)

            def _draw_cb(canvas, var, enabled):
                canvas.delete("all")
                fill = AZUL if var.get() and enabled else bg
                outline = VERDE if enabled else CINZA3
                canvas.create_rectangle(2, 2, 16, 16,
                                        outline=outline, fill=fill, width=2)
                if var.get() and enabled:
                    canvas.create_line(4, 9, 8, 13, fill=VERDE, width=2)
                    canvas.create_line(8, 13, 14, 5, fill=VERDE, width=2)

            _draw_cb(cv, v, tem_ddm)

            def _toggle(event, var=v, canvas=cv, enabled=tem_ddm):
                if not enabled:
                    return
                var.set(not var.get())
                _draw_cb(canvas, var, enabled)

            # Clique em qualquer parte da linha togla o checkbox
            for widget in (row, cv):
                widget.bind("<Button-1>", _toggle)

            tk.Label(row, text=d["dia_nome"], font=F_FIEL,
                     bg=bg, fg=fg_dia, width=12, anchor="w",
                     cursor="hand2" if tem_ddm else ""
                     ).pack(side="left")

            if tem_ddm:
                tema_txt = d["tema"][:30] if len(d["tema"]) > 30 else d["tema"]
                lbl_tema = tk.Label(row, text=tema_txt, font=F_SMALL,
                                    bg=bg, fg=TEXTO, anchor="w",
                                    cursor="hand2")
                lbl_tema.pack(side="left", padx=4)
                lbl_tema.bind("<Button-1>", _toggle)

                lbl_data = tk.Label(row, text=d["data_ddm"], font=F_SMALL,
                                    bg=bg, fg=CINZA3, cursor="hand2")
                lbl_data.pack(side="right", padx=8)
                lbl_data.bind("<Button-1>", _toggle)
            else:
                tk.Label(row, text="— sem arquivo —", font=F_SMALL,
                         bg=bg, fg=ERR, anchor="w").pack(side="left", padx=4)

            row.bind("<Button-1>", _toggle)

    # ── Campos cabeçalho ─────────────────────────────────────────────────
    def _campos_para(self, item: dict) -> dict:
        data_ddm = item["data_ddm"] or datetime.date.today().strftime("%d/%m/%Y")
        return {
            "setor":       self._vars["setor"].get(),
            "subarea":     self._vars["subarea"].get(),
            "turno":       self._vars["turno"].get(),
            "facilitador": self._vars["facilitador"].get(),
            "data":        data_ddm,
        }

    def _selecionados(self) -> list[dict]:
        sel = [d for d, v in zip(self._semana, self._checks)
               if v.get() and d["docx_path"]]
        if not sel:
            messagebox.showwarning("Nenhum selecionado",
                "Selecione ao menos um DDM da lista.")
        return sel

    # ── Gerar DOCX ───────────────────────────────────────────────────────
    def _gerar_docx_sel(self):
        sel = self._selecionados()
        if not sel: return
        self._log(f"\n─ Gerando {len(sel)} DOCX(s) ─")
        gerados = []
        for d in sel:
            src = d["docx_path"]
            dst = src.parent / (src.stem + "_PREENCHIDO.docx")
            try:
                info = processar_ddm(src, dst, self._campos_para(d))
                self._log(f"✔ {d['dia_nome']}: {dst.name} ({info['nomes']} nomes)", OK_C)
                gerados.append(dst)
            except Exception as e:
                self._log(f"✖ {d['dia_nome']}: {e}", ERR)
        if gerados and messagebox.askyesno("Abrir?",
                f"{len(gerados)} DOCX(s) gerado(s).\nAbrir pasta?"):
            os.startfile(str(gerados[0].parent))

    # ── Gerar PDF ────────────────────────────────────────────────────────
    def _gerar_pdf_sel(self):
        sel = self._selecionados()
        if not sel: return
        self._log(f"\n─ Gerando {len(sel)} PDF(s) via Word ─")

        pdfs_gerados = []

        for d in sel:
            src  = d["docx_path"]
            docx = src.parent / (src.stem + "_PREENCHIDO.docx")
            try:
                info = processar_ddm(src, docx, self._campos_para(d))
                self._log(f"✔ {d['dia_nome']}: {info['nomes']} nomes")
                self._log(f"  Convertendo para PDF...")
                pdf = converter_pdf_word(docx)
                self._log(f"✔ {pdf.name}", OK_C)
                pdfs_gerados.append(pdf)
            except Exception as e:
                self._log(f"✖ {d['dia_nome']}: {e}", ERR)

        if not pdfs_gerados:
            return

        if len(pdfs_gerados) == 1:
            # Só um PDF → abre direto
            os.startfile(str(pdfs_gerados[0]))
        else:
            # Múltiplos → mescla em um único PDF e abre
            self._log(f"\n  Mesclando {len(pdfs_gerados)} PDFs...")
            try:
                merged = self._mesclar_pdfs(pdfs_gerados)
                self._log(f"✔ PDF combinado: {merged.name}", OK_C)
                os.startfile(str(merged))
            except Exception as e:
                self._log(f"⚠ Não foi possível mesclar ({e}).", WARN)
                self._log("  Abrindo PDFs individualmente...")
                for p in pdfs_gerados:
                    os.startfile(str(p))

    def _mesclar_pdfs(self, pdfs: list) -> Path:
        """Mescla lista de PDFs em um único arquivo usando pypdf."""
        try:
            from pypdf import PdfWriter
        except ImportError:
            try:
                from PyPDF2 import PdfWriter  # fallback
            except ImportError:
                raise RuntimeError("pypdf não instalado")

        writer = PdfWriter()
        for pdf_path in pdfs:
            writer.append(str(pdf_path))

        # Salva na pasta do primeiro PDF com nome composto
        pasta = pdfs[0].parent
        nome  = "DDM_SEMANA_COMBINADO.pdf"
        saida = pasta / nome
        with open(saida, "wb") as f:
            writer.write(f)
        return saida

    # ── Log ──────────────────────────────────────────────────────────────
    def _log(self, msg, cor=None):
        cor = cor or TEXTO
        w = self._log_w
        w.configure(state="normal")
        tag = "c" + cor.replace("#","")
        w.tag_configure(tag, foreground=cor)
        w.insert("end", msg + "\n", tag)
        w.see("end")
        w.configure(state="disabled")

    def _log_clear(self):
        self._log_w.configure(state="normal")
        self._log_w.delete("1.0","end")
        self._log_w.configure(state="disabled")

# ════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    login = LoginWindow()
    login.mainloop()
    user = login.get_user()
    if not user:
        sys.exit(0)
    users = _load_users()
    app = MainWindow(user, users)
    app.mainloop()
