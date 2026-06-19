"""
DDM Manager - Metalfrio Solutions
v2.0 — Login + Admin + Participantes + PDF
"""

import os, copy, json, shutil, hashlib, datetime, subprocess, tkinter as tk
from tkinter import messagebox, filedialog, simpledialog
from pathlib import Path

# ─── Paleta ──────────────────────────────────────────────────────────────────
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

F_TITLE  = ("Segoe UI", 17, "bold")
F_SUB    = ("Segoe UI", 11)
F_LBL    = ("Segoe UI",  9, "bold")
F_FIEL   = ("Segoe UI", 10)
F_BTN    = ("Segoe UI", 11, "bold")
F_BTN_SM = ("Segoe UI",  9, "bold")
F_LOG    = ("Consolas",  9)

# ─── Mapeamento ───────────────────────────────────────────────────────────────
DIA_MAP = {0:"2", 1:"3", 2:"4", 3:"5", 4:"6", 5:"7", 6:"7"}
DIAS_PT = ["Segunda","Terça","Quarta","Quinta","Sexta","Sábado","Domingo"]
MES_MAP = {
    1:"1 - JANEIRO", 2:"2 - FEVEREIRO", 3:"3 - MARÇO",
    4:"4 - ABRIL",   5:"5 - MAIO",      6:"6 - JUNHO",
    7:"7 - JULHO",   8:"8 - AGOSTO",    9:"9 - SETEMBRO",
    10:"10 - OUTUBRO", 11:"11 - NOVEMBRO", 12:"12 - DEZEMBRO",
}

# ─── Participantes fixos ──────────────────────────────────────────────────────
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

# ─── Autenticação ─────────────────────────────────────────────────────────────
ADMIN_RE   = "184"   # Joel Borges de Campos Junior
USERS_FILE = Path(__file__).parent / "ddm_users.json"

def _hash(pw: str) -> str:
    return hashlib.sha256(pw.encode()).hexdigest()

def _load_users() -> dict:
    if USERS_FILE.exists():
        try:
            return json.loads(USERS_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    # Seed inicial: apenas Joel (admin), senha padrão "metalfrio"
    users = {
        "184": {
            "nome": "JOEL BORGES DE CAMPOS JUNIOR",
            "pw":   _hash("metalfrio"),
            "admin": True,
        }
    }
    _save_users(users)
    return users

def _save_users(users: dict) -> None:
    USERS_FILE.write_text(
        json.dumps(users, ensure_ascii=False, indent=2), encoding="utf-8"
    )

# ════════════════════════════════════════════════════════════════════════════
# Lógica de negócio — detecção DDM
# ════════════════════════════════════════════════════════════════════════════

def encontrar_pasta_dia(raiz, data):
    num = DIA_MAP[data.weekday()]
    p = Path(raiz)
    if not p.exists():
        return None, f"Pasta raiz não encontrada:\n{raiz}"
    for item in p.iterdir():
        if item.is_dir() and item.name.startswith(num + " -"):
            return str(item), ""
    return None, f"Pasta '{num} - {DIAS_PT[data.weekday()]}...' não encontrada."

def encontrar_pasta_mes(pasta_dia, data):
    nome = MES_MAP[data.month]
    for item in Path(pasta_dia).iterdir():
        if item.is_dir() and item.name.upper() == nome.upper():
            return str(item), ""
        if item.is_dir() and item.name.startswith(str(data.month) + " -"):
            return str(item), ""
    return None, f"Pasta '{nome}' não encontrada."

def encontrar_ddm_semana(pasta_mes, data):
    docxs = sorted(Path(pasta_mes).glob("*.docx"))
    if not docxs:
        return None, f"Nenhum .docx em:\n{pasta_mes}"
    alvo, melhor = None, None
    for f in docxs:
        try:
            dd, mm = int(f.name[:2]), int(f.name[3:5])
            delta = (data - datetime.date(data.year, mm, dd)).days
            if 0 <= delta <= 6 and (melhor is None or delta < melhor):
                alvo, melhor = f, delta
        except Exception:
            continue
    if alvo:
        return str(alvo), ""
    return None, f"Sem DDM para {data:%d/%m/%Y}.\nArquivos: {[f.name for f in docxs]}"

def extrair_tema(nome_arq):
    stem = Path(nome_arq).stem
    p = stem.split(" ", 1)
    return p[1].strip() if len(p) == 2 else stem

def extrair_data_ddm(nome_arq, ano):
    try:
        dd, mm = int(nome_arq[:2]), int(nome_arq[3:5])
        return f"{dd:02d}/{mm:02d}/{ano}"
    except Exception:
        return ""

# ════════════════════════════════════════════════════════════════════════════
# Processamento DOCX
# ════════════════════════════════════════════════════════════════════════════

def _cells_unicas(linha):
    seen, res = set(), []
    for c in linha.cells:
        cid = id(c._tc)
        if cid not in seen:
            seen.add(cid)
            res.append(c)
    return res

def _run_apos_label(para, valor):
    from docx.oxml.ns import qn
    from lxml import etree
    runs = para.runs
    if not runs:
        return
    for r in runs[1:]:
        r._element.getparent().remove(r._element)
    novo = copy.deepcopy(runs[0]._element)
    t = novo.find(qn("w:t"))
    if t is None:
        t = etree.SubElement(novo, qn("w:t"))
    t.text = " " + valor
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    para._element.append(novo)

def _escrever_celula(cell, valor, centralizar_v=True, alinhar_h="center"):
    """
    Escreve valor na célula com controle de alinhamento.
    alinhar_h: "center" | "left"
    centralizar_v: True → vAlign center
    """
    from docx.oxml.ns import qn
    from lxml import etree

    # Alinhamento vertical na célula (tcPr)
    tcp = cell._tc.find(qn("w:tcPr"))
    if tcp is None:
        tcp = etree.SubElement(cell._tc, qn("w:tcPr"))
    val_el = tcp.find(qn("w:vAlign"))
    if val_el is None:
        val_el = etree.SubElement(tcp, qn("w:vAlign"))
    val_el.set(qn("w:val"), "center" if centralizar_v else "top")

    para = cell.paragraphs[0] if cell.paragraphs else None
    if not para:
        return

    # Remove runs existentes
    for r in para.runs:
        r._element.getparent().remove(r._element)

    # Alinhamento horizontal no parágrafo (pPr)
    ppr = para._element.find(qn("w:pPr"))
    if ppr is None:
        ppr = etree.SubElement(para._element, qn("w:pPr"))
        para._element.insert(0, ppr)
    jc = ppr.find(qn("w:jc"))
    if jc is None:
        jc = etree.SubElement(ppr, qn("w:jc"))
    jc.set(qn("w:val"), "center" if alinhar_h == "center" else "left")

    # Run com texto
    r_el = etree.SubElement(para._element, qn("w:r"))
    rpr  = etree.SubElement(r_el, qn("w:rPr"))
    fnts = etree.SubElement(rpr, qn("w:rFonts"))
    for a, v in [("w:ascii","Arial"),("w:hAnsi","Arial"),
                 ("w:eastAsia","Batang"),("w:cs","Arial")]:
        fnts.set(qn(a), v)
    for tag in ("w:sz", "w:szCs"):
        e = etree.SubElement(rpr, qn(tag))
        e.set(qn("w:val"), "18")  # 9 pt
    t = etree.SubElement(r_el, qn("w:t"))
    t.text = valor
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

def _data_esta_preenchida(doc) -> bool:
    """Verifica se o campo DATA: já tem valor no documento."""
    tabela = doc.tables[0]
    for linha in tabela.rows:
        for cell in _cells_unicas(linha):
            for para in cell.paragraphs:
                txt = para.text.strip()
                if txt.startswith("DATA:"):
                    conteudo = txt[5:].strip()
                    return conteudo != ""
    return False

def processar_ddm(src, dst, campos):
    from docx import Document

    LABELS_CAB = {
        "SETOR/LINHA:": campos["setor"],
        "TURNO:":       campos["turno"],
        "SUBÁREA:":     campos["subarea"],
        "FACILITADOR:": campos["facilitador"],
    }

    doc = Document(src)

    # Decide se preenche DATA
    data_no_doc = _data_esta_preenchida(doc)
    data_label_valor = {} if data_no_doc else {"DATA:": campos["data"]}

    tabela = doc.tables[0]
    part_idx = 0
    info = {"cabecalho": [], "nomes": 0, "data_preenchida": not data_no_doc}

    for linha in tabela.rows:
        cells_u = _cells_unicas(linha)

        # ── Cabeçalho ──
        for cell in cells_u:
            for para in cell.paragraphs:
                txt = para.text.strip()
                # Campos normais (só se vazios)
                for label, valor in LABELS_CAB.items():
                    if txt.startswith(label) and txt[len(label):].strip() == "":
                        _run_apos_label(para, valor)
                        info["cabecalho"].append(label)
                # DATA: preenche se ausente no documento
                for label, valor in data_label_valor.items():
                    if txt.startswith(label) and txt[len(label):].strip() == "":
                        _run_apos_label(para, valor)
                        info["cabecalho"].append(label)

        # ── Participante ──
        if not cells_u:
            continue
        try:
            int(cells_u[0].text.strip())
        except ValueError:
            continue

        if part_idx < len(PARTICIPANTES) and len(cells_u) >= 3:
            re_v, nome_v = PARTICIPANTES[part_idx]
            # RE: centralizado horizontal + vertical
            _escrever_celula(cells_u[1], re_v,
                             centralizar_v=True, alinhar_h="center")
            # NOME: alinhado à esquerda + centralizado vertical
            _escrever_celula(cells_u[2], nome_v,
                             centralizar_v=True, alinhar_h="left")
            info["nomes"] += 1

        part_idx += 1

    doc.save(dst)
    return info

def converter_pdf(caminho_docx, pasta_saida):
    candidatos = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "soffice",
    ]
    soffice = next(
        (c for c in candidatos if shutil.which(c) or Path(c).exists()), None
    )
    if soffice is None:
        raise RuntimeError(
            "LibreOffice não encontrado.\n"
            "Baixe em: https://www.libreoffice.org/download/\n\n"
            "Alternativa: abra o DOCX e imprima pelo Word."
        )
    subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf",
         "--outdir", pasta_saida, caminho_docx],
        check=True, capture_output=True
    )
    return str(Path(pasta_saida) / (Path(caminho_docx).stem + ".pdf"))


# ════════════════════════════════════════════════════════════════════════════
# Tela de Login
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
        # Header
        hdr = tk.Frame(self, bg=VERDE, height=70)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="❄  DDM Manager", font=F_TITLE,
                 bg=VERDE, fg=AZUL).pack(pady=18)

        # Formulário
        frm = tk.Frame(self, bg=AZUL, padx=36)
        frm.pack(fill="both", expand=True, pady=16)

        tk.Label(frm, text="RE (Registro de Empregado)", font=F_LBL,
                 bg=AZUL, fg=VERDE, anchor="w").pack(fill="x")
        self._v_re = tk.StringVar()
        e_re = tk.Entry(frm, textvariable=self._v_re, font=F_FIEL,
                        bg=CINZA2, fg=TEXTO, insertbackground=VERDE,
                        relief="flat", bd=0)
        e_re.pack(fill="x", ipady=7, ipadx=6, pady=(2, 14))
        e_re.focus()

        tk.Label(frm, text="Senha", font=F_LBL,
                 bg=AZUL, fg=VERDE, anchor="w").pack(fill="x")
        self._v_pw = tk.StringVar()
        e_pw = tk.Entry(frm, textvariable=self._v_pw, font=F_FIEL,
                        bg=CINZA2, fg=TEXTO, insertbackground=VERDE,
                        relief="flat", bd=0, show="●")
        e_pw.pack(fill="x", ipady=7, ipadx=6, pady=(2, 20))
        e_pw.bind("<Return>", lambda _: self._login())

        self._lbl_err = tk.Label(frm, text="", font=F_LBL,
                                 bg=AZUL, fg=ERR)
        self._lbl_err.pack()

        tk.Button(frm, text="Entrar →", font=F_BTN,
                  bg=VERDE, fg=AZUL, relief="flat", cursor="hand2",
                  padx=20, command=self._login).pack(pady=(4, 0))

    def _login(self):
        re   = self._v_re.get().strip()
        pw   = self._v_pw.get()
        user = self._users.get(re)
        if not user or user["pw"] != _hash(pw):
            self._lbl_err.config(text="RE ou senha incorretos.")
            self._v_pw.set("")
            return
        self._logged_user = {"re": re, **user}
        self.destroy()

    def get_user(self):
        return self._logged_user


# ════════════════════════════════════════════════════════════════════════════
# Painel Admin (janela modal)
# ════════════════════════════════════════════════════════════════════════════

class AdminPanel(tk.Toplevel):
    def __init__(self, parent, users: dict):
        super().__init__(parent)
        self.title("Administração de Usuários")
        self.configure(bg=AZUL)
        self.resizable(False, False)
        self._users = users
        self._build()
        self._center(560, 480)
        self.grab_set()

    def _center(self, w, h):
        self.update_idletasks()
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _build(self):
        # Header
        hdr = tk.Frame(self, bg=GOLD, height=46)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="⚙  Administração de Usuários", font=F_BTN,
                 bg=GOLD, fg=AZUL).pack(side="left", padx=16, pady=10)

        # Lista
        lista_frm = tk.Frame(self, bg=AZUL, padx=16, pady=10)
        lista_frm.pack(fill="both", expand=True)

        tk.Label(lista_frm, text="Usuários cadastrados", font=F_LBL,
                 bg=AZUL, fg=VERDE).pack(anchor="w")

        cols_frm = tk.Frame(lista_frm, bg=CINZA1)
        cols_frm.pack(fill="x", pady=(4, 0))
        for txt, w in [("RE", 8), ("Nome", 34), ("Admin", 6), ("Ação", 6)]:
            tk.Label(cols_frm, text=txt, font=F_LBL, bg=CINZA1,
                     fg=CINZA3, width=w, anchor="w").pack(side="left", padx=4, pady=3)

        self._lista_frm = tk.Frame(lista_frm, bg=AZUL)
        self._lista_frm.pack(fill="both", expand=True, pady=4)
        self._refresh_lista()

        # Formulário novo usuário
        sep = tk.Frame(self, bg=CINZA1, height=2)
        sep.pack(fill="x")

        novo_frm = tk.Frame(self, bg=AZUL, padx=16, pady=10)
        novo_frm.pack(fill="x")
        tk.Label(novo_frm, text="Adicionar usuário", font=F_LBL,
                 bg=AZUL, fg=VERDE).grid(row=0, column=0, columnspan=4,
                                          sticky="w", pady=(0, 6))

        for col, (lbl, w) in enumerate([("RE", 8), ("Nome completo", 24),
                                         ("Senha", 12)]):
            tk.Label(novo_frm, text=lbl, font=F_LBL, bg=AZUL,
                     fg=CINZA3).grid(row=1, column=col, sticky="w", padx=4)

        self._v_new_re   = tk.StringVar()
        self._v_new_nome = tk.StringVar()
        self._v_new_pw   = tk.StringVar()

        tk.Entry(novo_frm, textvariable=self._v_new_re,   font=F_FIEL,
                 bg=CINZA2, fg=TEXTO, insertbackground=VERDE,
                 relief="flat", bd=0, width=9
                 ).grid(row=2, column=0, padx=4, ipady=5)
        tk.Entry(novo_frm, textvariable=self._v_new_nome, font=F_FIEL,
                 bg=CINZA2, fg=TEXTO, insertbackground=VERDE,
                 relief="flat", bd=0, width=26
                 ).grid(row=2, column=1, padx=4, ipady=5)
        tk.Entry(novo_frm, textvariable=self._v_new_pw,   font=F_FIEL,
                 bg=CINZA2, fg=TEXTO, insertbackground=VERDE,
                 relief="flat", bd=0, width=14, show="●"
                 ).grid(row=2, column=2, padx=4, ipady=5)
        tk.Button(novo_frm, text="＋ Adicionar", font=F_BTN_SM,
                  bg=VERDE, fg=AZUL, relief="flat", cursor="hand2",
                  command=self._adicionar
                  ).grid(row=2, column=3, padx=8)

        self._lbl_msg = tk.Label(novo_frm, text="", font=F_LBL,
                                 bg=AZUL, fg=OK_C)
        self._lbl_msg.grid(row=3, column=0, columnspan=4, sticky="w", pady=4)

        # Botão fechar
        rod = tk.Frame(self, bg=CINZA1, height=46)
        rod.pack(fill="x", side="bottom")
        rod.pack_propagate(False)
        tk.Button(rod, text="Fechar", font=F_BTN_SM,
                  bg=CINZA1, fg=CINZA3, relief="flat", cursor="hand2",
                  command=self.destroy).pack(side="right", padx=12, pady=10)

    def _refresh_lista(self):
        for w in self._lista_frm.winfo_children():
            w.destroy()
        for re, info in self._users.items():
            row = tk.Frame(self._lista_frm, bg=CINZA2)
            row.pack(fill="x", pady=1)
            admin_txt = "★" if info.get("admin") else ""
            tk.Label(row, text=re, font=F_FIEL, bg=CINZA2,
                     fg=TEXTO, width=8, anchor="w").pack(side="left", padx=4)
            tk.Label(row, text=info["nome"][:36], font=F_FIEL, bg=CINZA2,
                     fg=TEXTO, width=34, anchor="w").pack(side="left")
            tk.Label(row, text=admin_txt, font=F_FIEL, bg=CINZA2,
                     fg=GOLD, width=5).pack(side="left")
            if not info.get("admin"):
                tk.Button(row, text="✕", font=F_BTN_SM,
                          bg=CINZA2, fg=ERR, relief="flat", cursor="hand2",
                          command=lambda r=re: self._remover(r)
                          ).pack(side="left", padx=6)
            # Botão redefinir senha (todos exceto admin não pode se remover)
            tk.Button(row, text="🔑", font=F_BTN_SM,
                      bg=CINZA2, fg=WARN, relief="flat", cursor="hand2",
                      command=lambda r=re: self._redefinir_senha(r)
                      ).pack(side="left", padx=2)

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
        self._msg(f"Usuário {re} — {nome} adicionado.", OK_C)

    def _remover(self, re):
        if not messagebox.askyesno("Confirmar",
                f"Remover usuário {re} — {self._users[re]['nome']}?",
                parent=self):
            return
        del self._users[re]
        _save_users(self._users)
        self._refresh_lista()
        self._msg("Usuário removido.", WARN)

    def _redefinir_senha(self, re):
        nova = simpledialog.askstring(
            "Redefinir senha", f"Nova senha para {self._users[re]['nome']}:",
            show="●", parent=self
        )
        if not nova:
            return
        self._users[re]["pw"] = _hash(nova)
        _save_users(self._users)
        self._msg("Senha redefinida.", OK_C)

    def _msg(self, txt, cor=OK_C):
        self._lbl_msg.config(text=txt, fg=cor)


# ════════════════════════════════════════════════════════════════════════════
# Janela principal
# ════════════════════════════════════════════════════════════════════════════

class MainWindow(tk.Tk):
    def __init__(self, user: dict, users: dict):
        super().__init__()
        self._user  = user
        self._users = users
        self.title("DDM Manager — Metalfrio")
        self.configure(bg=AZUL)
        self.resizable(False, False)
        self._ddm_path = None
        self._ddm_tema = ""
        self._build()
        self.after(100, self._detectar)

    def _center(self, w, h):
        self.geometry(f"{w}x{h}")
        self.update_idletasks()
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _build(self):
        self._center(860, 640)

        # ── Header ──
        hdr = tk.Frame(self, bg=VERDE, height=58)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)

        tk.Label(hdr, text="❄  DDM Manager", font=F_TITLE,
                 bg=VERDE, fg=AZUL).pack(side="left", padx=18, pady=10)

        # Info do usuário logado (direita do header)
        user_frm = tk.Frame(hdr, bg=VERDE)
        user_frm.pack(side="right", padx=12)

        nome_curto = self._user["nome"].split()[0].capitalize()
        admin_tag  = "  ★ Admin" if self._user.get("admin") else ""
        tk.Label(user_frm,
                 text=f"👤 {nome_curto}{admin_tag}",
                 font=("Segoe UI", 9, "bold"),
                 bg=VERDE, fg=AZUL).pack(anchor="e")

        self._lbl_clock = tk.Label(user_frm, text="", font=("Segoe UI", 9),
                                   bg=VERDE, fg=AZUL)
        self._lbl_clock.pack(anchor="e")
        self._tick()

        # Botão admin (só para Joel)
        if self._user.get("admin"):
            tk.Button(hdr, text="⚙ Usuários", font=F_BTN_SM,
                      bg=GOLD, fg=AZUL, relief="flat", cursor="hand2",
                      command=self._abrir_admin
                      ).pack(side="right", padx=8, pady=14)

        # ── Corpo ──
        body = tk.Frame(self, bg=AZUL)
        body.pack(fill="both", expand=True, padx=18, pady=10)

        # Esquerda
        esq = tk.Frame(body, bg=AZUL, width=440)
        esq.pack(side="left", fill="y", padx=(0, 12))
        esq.pack_propagate(False)

        self._vars: dict[str, tk.StringVar] = {}

        def campo(key, label, tem_browse=False):
            f = tk.Frame(esq, bg=AZUL)
            f.pack(fill="x", pady=3)
            tk.Label(f, text=label, font=F_LBL, bg=AZUL,
                     fg=VERDE, anchor="w").pack(fill="x")
            row = tk.Frame(f, bg=AZUL)
            row.pack(fill="x")
            v = tk.StringVar(value=DEFAULTS.get(key, ""))
            self._vars[key] = v
            e = tk.Entry(row, textvariable=v, font=F_FIEL, width=38,
                         bg=CINZA2, fg=TEXTO, insertbackground=VERDE,
                         relief="flat", bd=0)
            e.pack(side="left", ipady=5, ipadx=4, fill="x", expand=True)
            if tem_browse:
                tk.Button(row, text="…", font=("Segoe UI", 10),
                          bg=VERDE, fg=AZUL, relief="flat", cursor="hand2",
                          command=self._browse).pack(side="left", padx=(4, 0))

        campo("raiz",        "📁  Pasta raiz  Q:\\Transferencia\\DDM 2026",
              tem_browse=True)
        campo("setor",       "🏭  Setor / Linha")
        campo("subarea",     "🔹  Subárea")
        campo("turno",       "⏰  Turno")
        campo("facilitador", "👤  Facilitador")

        # Painel DDM
        det = tk.LabelFrame(esq, text="  DDM da semana  ", font=F_LBL,
                            bg=AZUL, fg=VERDE, bd=1, relief="groove")
        det.pack(fill="x", pady=(10, 0))

        self._v_arq  = tk.StringVar(value="—")
        self._v_tema = tk.StringVar(value="—")
        self._v_data = tk.StringVar(value="—")

        for lbl, var in [("Arquivo:",  self._v_arq),
                         ("Tema:",     self._v_tema),
                         ("Data DDM:", self._v_data)]:
            r = tk.Frame(det, bg=AZUL)
            r.pack(fill="x", padx=8, pady=2)
            tk.Label(r, text=lbl, font=F_LBL, bg=AZUL, fg="#8EAABF",
                     width=9, anchor="w").pack(side="left")
            tk.Label(r, textvariable=var, font=F_FIEL, bg=AZUL, fg=TEXTO,
                     anchor="w", wraplength=310, justify="left").pack(side="left")

        tk.Button(det, text="🔄  Detectar novamente", font=F_FIEL,
                  bg=CINZA2, fg=VERDE, relief="flat", cursor="hand2",
                  command=self._detectar).pack(pady=5)

        tk.Label(esq,
                 text=f"👥  {len(PARTICIPANTES)} participantes carregados",
                 font=F_LBL, bg=AZUL, fg=OK_C).pack(anchor="w", pady=(8, 0))

        # Direita: log
        dir_ = tk.Frame(body, bg=AZUL)
        dir_.pack(side="right", fill="both", expand=True)
        tk.Label(dir_, text="Log de execução", font=F_LBL,
                 bg=AZUL, fg=VERDE).pack(anchor="w")
        self._log_w = tk.Text(dir_, font=F_LOG, bg=CINZA1, fg=TEXTO,
                              relief="flat", state="disabled",
                              wrap="word", width=32)
        self._log_w.pack(fill="both", expand=True)

        # Rodapé
        rod = tk.Frame(self, bg=CINZA1, height=58)
        rod.pack(fill="x", side="bottom")
        rod.pack_propagate(False)

        tk.Button(rod, text="📄  Gerar DOCX", font=F_BTN,
                  bg=CINZA2, fg=VERDE, relief="flat", cursor="hand2",
                  padx=14, command=self._gerar_docx).pack(side="left", padx=8, pady=10)

        tk.Button(rod, text="🖨️  Gerar PDF + Abrir", font=F_BTN,
                  bg=VERDE, fg=AZUL, relief="flat", cursor="hand2",
                  padx=14, command=self._gerar_pdf).pack(side="left", padx=4, pady=10)

        tk.Button(rod, text="✖  Sair", font=F_BTN,
                  bg=CINZA1, fg=ERR, relief="flat", cursor="hand2",
                  command=self.destroy).pack(side="right", padx=12, pady=10)

    def _tick(self):
        agora = datetime.datetime.now()
        self._lbl_clock.config(
            text=f"{DIAS_PT[agora.weekday()]}, {agora.strftime('%d/%m/%Y  %H:%M:%S')}"
        )
        self.after(1000, self._tick)

    def _browse(self):
        p = filedialog.askdirectory(title="Pasta raiz DDM 2026")
        if p:
            self._vars["raiz"].set(p)
            self._detectar()

    def _abrir_admin(self):
        AdminPanel(self, self._users)

    def _detectar(self):
        self._log_clear()
        raiz = self._vars["raiz"].get().strip()
        hoje = datetime.date.today()
        self._log(f"Data: {hoje:%d/%m/%Y} ({DIAS_PT[hoje.weekday()]})")
        self._log(f"Raiz: {raiz}\n")

        p_dia, err = encontrar_pasta_dia(raiz, hoje)
        if err:
            self._log(f"⚠  {err}", WARN); self._reset_info(); return
        self._log(f"✔ Dia  → {Path(p_dia).name}")

        p_mes, err = encontrar_pasta_mes(p_dia, hoje)
        if err:
            self._log(f"⚠  {err}", WARN); self._reset_info(); return
        self._log(f"✔ Mês  → {Path(p_mes).name}")

        ddm, err = encontrar_ddm_semana(p_mes, hoje)
        if err:
            self._log(f"⚠  {err}", WARN); self._reset_info(); return

        self._ddm_path = ddm
        self._ddm_tema = extrair_tema(Path(ddm).name)
        data_ddm       = extrair_data_ddm(Path(ddm).name, hoje.year)

        self._v_arq.set(Path(ddm).name)
        self._v_tema.set(self._ddm_tema)
        self._v_data.set(data_ddm or "—")
        self._log(f"✔ DDM  → {Path(ddm).name}")
        self._log(f"✔ Tema → {self._ddm_tema}", OK_C)
        self._log(f"\n✔ {len(PARTICIPANTES)} participantes prontos.", OK_C)

    def _reset_info(self):
        self._ddm_path = None
        self._v_arq.set("—"); self._v_tema.set("—"); self._v_data.set("—")

    def _campos(self):
        if not self._ddm_path:
            messagebox.showerror("Sem DDM", "DDM não detectado.\nVerifique a pasta raiz.")
            return None
        hoje = datetime.date.today()
        data_ddm = extrair_data_ddm(Path(self._ddm_path).name, hoje.year)
        return {
            "setor":       self._vars["setor"].get(),
            "subarea":     self._vars["subarea"].get(),
            "turno":       self._vars["turno"].get(),
            "facilitador": self._vars["facilitador"].get(),
            "data":        data_ddm or hoje.strftime("%d/%m/%Y"),
        }

    def _gerar_docx(self):
        c = self._campos()
        if not c: return
        src = self._ddm_path
        dst = str(Path(src).parent / (Path(src).stem + "_PREENCHIDO.docx"))
        try:
            self._log("\n─ Gerando DOCX ─")
            info = processar_ddm(src, dst, c)
            data_info = " (DATA preenchida)" if info["data_preenchida"] else ""
            self._log(f"✔ Cabeçalho: {len(info['cabecalho'])} campos{data_info}")
            self._log(f"✔ Nomes: {info['nomes']} participantes")
            self._log(f"✔ {Path(dst).name}", OK_C)
            if messagebox.askyesno("Abrir?", f"DOCX gerado:\n{Path(dst).name}\n\nAbrir?"):
                os.startfile(dst)
        except Exception as e:
            self._log(f"✖ {e}", ERR)
            messagebox.showerror("Erro", str(e))

    def _gerar_pdf(self):
        c = self._campos()
        if not c: return
        src  = self._ddm_path
        docx = str(Path(src).parent / (Path(src).stem + "_PREENCHIDO.docx"))
        try:
            self._log("\n─ Gerando PDF ─")
            info = processar_ddm(src, docx, c)
            self._log(f"✔ {info['nomes']} nomes injetados")
            self._log("Convertendo (LibreOffice)...")
            pdf = converter_pdf(docx, str(Path(docx).parent))
            self._log(f"✔ {Path(pdf).name}", OK_C)
            os.startfile(pdf)
        except Exception as e:
            self._log(f"✖ {e}", ERR)
            messagebox.showerror("Erro", str(e))

    def _log(self, msg, cor=None):
        cor = cor or TEXTO
        w = self._log_w
        w.configure(state="normal")
        tag = "c" + cor.replace("#", "")
        w.tag_configure(tag, foreground=cor)
        w.insert("end", msg + "\n", tag)
        w.see("end")
        w.configure(state="disabled")

    def _log_clear(self):
        self._log_w.configure(state="normal")
        self._log_w.delete("1.0", "end")
        self._log_w.configure(state="disabled")


# ════════════════════════════════════════════════════════════════════════════
# Entry point
# ════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    # 1. Login
    login = LoginWindow()
    login.mainloop()
    user = login.get_user()

    if not user:
        # Usuário fechou sem logar
        import sys; sys.exit(0)

    # 2. Janela principal
    users = _load_users()
    app = MainWindow(user, users)
    app.mainloop()
