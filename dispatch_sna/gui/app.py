"""Fenêtre principale du Dispatcher SNA (Tkinter).

Le travail est fait par dispatch_sna.services dans un thread ; la fenêtre ne
lit que la file de logs et l'état partagé (self._logq / self._state) via _poll.
"""
import os
import threading
import time as _time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from ..config import load_config, save_config
from ..services.chemins import get_dep, lister_communes_prod, detecter_colonnes, supprimer_si_vide
from ..services.constantes import COL_ID_ERR
from ..services.dispatcher import ranger_commune, ranger_ids
from ..services.error_ids import extraire_ids
from .theme import (C_ACCENT, C_AMBER, C_BG, C_BORDER, C_BORDER2, C_CARD, C_CYAN, C_GREEN, C_MUTED,
                    C_PANEL, C_PANEL2, C_PURPLE, C_RED, C_TEXT, C_TEXT2, F_BODY, F_MONO, F_TITLE)

PHASES = [("verif","Vérif dest."),("analyse","Analyse"),("excel","Excel"),
          ("ppt","PPT"),("nettoyage","Nettoyage")]


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Dispatcher SNA — Organisation")
        self.geometry("1120x820")
        self.minsize(960, 640)
        self.configure(bg=C_BG)
        self.cfg = load_config()
        self._debug_visible = False
        self._t_start = None
        self._logq = []
        self._state = {}
        self._running = False
        self._build_styles()
        self._build_ui()

    def _build_styles(self):
        s = ttk.Style(); s.theme_use("clam")
        s.configure("Treeview", background=C_CARD, foreground=C_TEXT2, fieldbackground=C_CARD,
                    rowheight=30, font=(F_BODY,10), borderwidth=0)
        s.configure("Treeview.Heading", background=C_PANEL2, foreground=C_ACCENT,
                    font=(F_BODY,9,"bold"), relief="flat", padding=6)
        s.map("Treeview", background=[("selected","#243150")], foreground=[("selected","white")])
        for n,col in [("blue",C_ACCENT),("green",C_GREEN),("cyan",C_CYAN)]:
            s.configure(f"{n}.Horizontal.TProgressbar", troughcolor=C_PANEL2, background=col,
                        darkcolor=col, lightcolor=col, bordercolor=C_PANEL2, thickness=10)

    def _build_ui(self):
        # HEADER
        hdr = tk.Frame(self, bg=C_PANEL, height=66); hdr.pack(fill="x"); hdr.pack_propagate(False)
        logo = tk.Frame(hdr, bg=C_PANEL); logo.pack(side="left", padx=20)
        tk.Label(logo, text="📦", font=(F_BODY,24), bg=C_PANEL, fg=C_ACCENT).pack(side="left", pady=12)
        tb = tk.Frame(logo, bg=C_PANEL); tb.pack(side="left", padx=10)
        tk.Label(tb, text="Dispatcher SNA", font=(F_TITLE,16,"bold"), bg=C_PANEL, fg=C_TEXT).pack(anchor="w", pady=(11,0))
        tk.Label(tb, text="PC SFR · Prod IA → Dep/commune", font=(F_BODY,8,"bold"), bg=C_PANEL, fg=C_MUTED).pack(anchor="w")
        self.hdr_state = tk.Label(hdr, text="●  Prêt", font=(F_BODY,10,"bold"), bg=C_PANEL, fg=C_GREEN)
        self.hdr_state.pack(side="right", padx=(0,18))
        self.btn_debug = tk.Button(hdr, text="🐛 Debug", command=self._toggle_debug, bg="#2c1424",
                    fg="#ff9ecb", font=(F_BODY,10,"bold"), relief="flat", cursor="hand2", padx=16)
        self.btn_debug.pack(side="right", padx=10, pady=17)

        # PAGE DEBUG
        self.page_debug = tk.Frame(self, bg="#04060a")
        dh = tk.Frame(self.page_debug, bg="#16040e", height=46); dh.pack(fill="x"); dh.pack_propagate(False)
        tk.Label(dh, text="   🐛 Journal", font=(F_TITLE,13,"bold"), bg="#16040e", fg="#ff6b9d").pack(side="left")
        tk.Button(dh, text="🗑 Effacer", font=(F_BODY,9), bg="#3a1020", fg="#ff9ecb", relief="flat",
                  cursor="hand2", command=self._clear_log).pack(side="right", padx=(0,10), pady=9)
        tk.Button(dh, text="◀ Retour", font=(F_BODY,9,"bold"), bg=C_ACCENT, fg="white", relief="flat",
                  cursor="hand2", command=self._toggle_debug, padx=14).pack(side="right", padx=6, pady=9)
        lw = tk.Frame(self.page_debug, bg="#04060a"); lw.pack(fill="both", expand=True, padx=12, pady=12)
        self.txt_log = tk.Text(lw, bg="#02040a", fg=C_GREEN, font=(F_MONO,10), relief="flat",
                               state="disabled", wrap="word", padx=10, pady=8)
        sbl = ttk.Scrollbar(lw, orient="vertical", command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=sbl.set)
        self.txt_log.pack(side="left", fill="both", expand=True); sbl.pack(side="left", fill="y")

        # PAGE PRINCIPALE (scroll)
        self.page_main = tk.Frame(self, bg=C_BG); self.page_main.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(self.page_main, bg=C_BG, highlightthickness=0)
        vsb = ttk.Scrollbar(self.page_main, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y"); self.canvas.pack(side="left", fill="both", expand=True)
        self.sf = tk.Frame(self.canvas, bg=C_BG)
        self._win = self.canvas.create_window((0,0), window=self.sf, anchor="nw")
        self.sf.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfig(self._win, width=e.width))
        self.canvas.bind_all("<MouseWheel>", lambda e: self.canvas.yview_scroll(int(-1*(e.delta/120)),"units"))
        self.canvas.bind_all("<Button-4>", lambda e: self.canvas.yview_scroll(-1,"units"))
        self.canvas.bind_all("<Button-5>", lambda e: self.canvas.yview_scroll(1,"units"))
        self._build_content(self.sf)

    def _card(self, parent, title, icon, accent, sub=""):
        wrap = tk.Frame(parent, bg=C_BORDER); wrap.pack(fill="x", padx=20, pady=(0,14))
        inner = tk.Frame(wrap, bg=C_CARD); inner.pack(fill="both", expand=True, padx=1, pady=1)
        head = tk.Frame(inner, bg=C_CARD); head.pack(fill="x", padx=18, pady=(14,4))
        tk.Frame(head, bg=accent, width=4, height=20).pack(side="left", padx=(0,10))
        tt = tk.Frame(head, bg=C_CARD); tt.pack(side="left")
        tk.Label(tt, text=f"{icon}  {title}", font=(F_TITLE,12,"bold"), bg=C_CARD, fg=C_TEXT).pack(anchor="w")
        if sub: tk.Label(tt, text=sub, font=(F_BODY,8), bg=C_CARD, fg=C_MUTED).pack(anchor="w")
        body = tk.Frame(inner, bg=C_CARD); body.pack(fill="both", expand=True, padx=18, pady=(6,16))
        return body

    def _build_content(self, root):
        tk.Frame(root, bg=C_BG, height=14).pack(fill="x")

        # CONFIG
        cfg = self._card(root, "Configuration", "⚙️", C_ACCENT, "Chemins — sauvegardés automatiquement")
        self._path_row(cfg, "Racine SharePoint", "root_sp", True, "Dossier audit_SNA (contient DepXX)")
        self._path_row(cfg, "Dossier À déposer", "prod_ia", True, "Prod IA/À déposer — dossiers commune de Babayaga")
        opt = tk.Frame(cfg, bg=C_CARD); opt.pack(fill="x", pady=(8,0))
        self.var_supp = tk.BooleanVar(value=bool(self.cfg.get("supprimer", True)))
        self._check(opt, "🗑️ Supprimer le dossier de Prod IA après rangement", self.var_supp, self._save_cfg)

        # STATUT
        st_wrap = tk.Frame(root, bg=C_BORDER); st_wrap.pack(fill="x", padx=20, pady=(0,14))
        self.frm_status = tk.Frame(st_wrap, bg="#0a0f1a"); self.frm_status.pack(fill="both", padx=1, pady=1)
        top = tk.Frame(self.frm_status, bg="#0a0f1a"); top.pack(fill="x", padx=18, pady=(14,4))
        self.lbl_status = tk.Label(top, text="⏸  Prêt — cliquez sur Ranger",
                    bg="#0a0f1a", fg=C_MUTED, font=(F_TITLE,13,"bold")); self.lbl_status.pack(side="left")
        self.lbl_timer = tk.Label(top, text="", bg="#0a0f1a", fg=C_CYAN, font=(F_MONO,11,"bold")); self.lbl_timer.pack(side="right")
        self.frm_timeline = tk.Frame(self.frm_status, bg="#0a0f1a"); self.frm_timeline.pack(fill="x", padx=18, pady=(2,4))
        self.progress = ttk.Progressbar(self.frm_status, style="blue.Horizontal.TProgressbar", mode="determinate")
        self.progress.pack(fill="x", padx=18, pady=(4,2))
        self.lbl_prog = tk.Label(self.frm_status, text="", bg="#0a0f1a", fg=C_MUTED, font=(F_BODY,9))
        self.lbl_prog.pack(pady=(0,12))

        # ACTION
        act = self._card(root, "Rangement", "📦", C_GREEN,
                         "Commune complète : Prod IA → Dep/commune · ID erreur : seules les lignes / PPT des ID collés")
        mrow = tk.Frame(act, bg=C_CARD); mrow.pack(fill="x", pady=(0,4))
        tk.Label(mrow, text="Ranger par :", bg=C_CARD, fg=C_MUTED, font=(F_BODY,9,"bold")).pack(side="left", padx=(0,8))
        self.var_mode = tk.StringVar(value=self.cfg.get("mode", "commune"))
        for txt,val in [("Commune complète (tout Prod IA)","commune"),("ID erreur (lignes ciblées)","ids")]:
            tk.Radiobutton(mrow, text=txt, value=val, variable=self.var_mode, command=self._on_mode,
                           bg=C_CARD, fg=C_TEXT2, selectcolor=C_PANEL2, activebackground=C_CARD,
                           activeforeground=C_TEXT, font=(F_BODY,10), relief="flat",
                           cursor="hand2").pack(side="left", padx=(0,16))

        # Saisie des ID erreur (mode ids uniquement)
        self.frm_ids = tk.Frame(act, bg=C_CARD)
        tk.Label(self.frm_ids, text="Collez vos ID erreur, un par ligne (ex : 74143_1, 74143_2) :",
                 bg=C_CARD, fg=C_TEXT2, font=(F_BODY,9)).pack(anchor="w")
        ids_box = tk.Frame(self.frm_ids, bg="#060a12", highlightthickness=1,
                           highlightbackground=C_BORDER2, highlightcolor=C_ACCENT)
        ids_box.pack(fill="x", pady=(4,2))
        self.txt_ids = tk.Text(ids_box, height=6, bg="#060a12", fg=C_TEXT, insertbackground="white",
                               font=(F_MONO,10), relief="flat", bd=0, padx=8, pady=6, wrap="none", undo=True)
        self.txt_ids.pack(fill="both", expand=True)
        self.txt_ids.bind("<<Modified>>", self._on_ids_modified)
        self.txt_ids.bind("<FocusIn>", lambda e: ids_box.config(highlightbackground=C_ACCENT))
        self.txt_ids.bind("<FocusOut>", lambda e: ids_box.config(highlightbackground=C_BORDER2))
        # Tk n'a pas de placeholder natif : label superposé, masqué dès que du texte est saisi
        self.lbl_ids_ph = tk.Label(ids_box, text="74143_1\n74143_2\n38068_207", bg="#060a12", fg=C_MUTED,
                                   font=(F_MONO,10), justify="left", cursor="xterm")
        self.lbl_ids_ph.bind("<Button-1>", lambda e: self.txt_ids.focus_set())
        self.lbl_ids_ph.place(x=9, y=7)
        self.lbl_ids = tk.Label(self.frm_ids, text="", bg=C_CARD, fg=C_MUTED, font=(F_BODY,8))
        self.lbl_ids.pack(anchor="w")

        # Colonnes Excel à mettre à jour
        colh = tk.Frame(self.frm_ids, bg=C_CARD); colh.pack(fill="x", pady=(10,0))
        tk.Label(colh, text="Colonnes Excel à mettre à jour :", bg=C_CARD, fg=C_MUTED,
                 font=(F_BODY,9,"bold")).pack(side="left")
        for txt, cmd in [("🔍 Détecter", self._detect_cols), ("☑ Toutes", lambda: self._cocher_cols(True)),
                         ("☐ Aucune", lambda: self._cocher_cols(False))]:
            tk.Button(colh, text=txt, command=cmd, bg=C_PANEL2, fg=C_TEXT2, font=(F_BODY,8), relief="flat",
                      cursor="hand2", padx=10, pady=2).pack(side="right", padx=(6,0))
        self.frm_cols = tk.Frame(self.frm_ids, bg=C_CARD); self.frm_cols.pack(fill="x", pady=(4,0))
        self.lbl_cols = tk.Label(self.frm_ids, text="", bg=C_CARD, fg=C_MUTED, font=(F_BODY,8), justify="left",
                                 wraplength=900)
        self.lbl_cols.pack(anchor="w")
        self._col_vars = {}

        prow = tk.Frame(self.frm_ids, bg=C_CARD); prow.pack(fill="x", pady=(10,0))
        tk.Label(prow, text="PPT :", bg=C_CARD, fg=C_MUTED, font=(F_BODY,9,"bold")).pack(side="left", padx=(0,4))
        self.var_ppt_rempl = tk.BooleanVar(value=bool(self.cfg.get("ppt_remplacer", False)))
        self.var_ppt_modif = tk.BooleanVar(value=bool(self.cfg.get("ppt_modifier", True))
                                           and not self.var_ppt_rempl.get())
        self._check(prow, "🔁 Remplacer les PPT", self.var_ppt_rempl, lambda: self._on_ppt("rempl"))
        self._check(prow, "✏️ Modifier les PPT (Analyse et Lien Street View)", self.var_ppt_modif,
                    lambda: self._on_ppt("modif"))
        tk.Label(self.frm_ids, text="  Remplacer : le PPT SharePoint est écrasé par celui de Prod IA · "
                 "Modifier : Analyse ← colonne « Remarque », Lien Street View ← colonne « Lien Street View » "
                 "· Aucune case : Excel seul",
                 bg=C_CARD, fg=C_MUTED, font=(F_BODY,8), justify="left", wraplength=900).pack(anchor="w", pady=(2,0))

        self.btn_run = tk.Button(act, text="📦  Ranger tout Prod IA", command=self._run,
                    bg=C_GREEN, fg="#062017", font=(F_TITLE,12,"bold"), relief="flat", cursor="hand2", pady=11)
        self.btn_run.pack(fill="x", pady=(10,0))
        self._on_mode(preview=False)

        # STATS
        stt = self._card(root, "Statistiques", "📊", C_AMBER)
        sg = tk.Frame(stt, bg=C_CARD); sg.pack(fill="x")
        self._stat_vars = {}
        defs = [("communes","Communes",C_ACCENT),("lignes","Lignes",C_CYAN),("excel","Excel",C_GREEN),
                ("ppt","PPT",C_PURPLE),("err","Erreurs",C_RED)]
        for i,(k,lbl,col) in enumerate(defs):
            sg.columnconfigure(i, weight=1)
            c = tk.Frame(sg, bg=C_PANEL2); c.grid(row=0,column=i,sticky="nsew",padx=4,pady=2)
            tk.Label(c, text=lbl.upper(), bg=C_PANEL2, fg=C_MUTED, font=(F_BODY,8,"bold")).pack(pady=(10,0))
            v = tk.StringVar(value="—"); self._stat_vars[k]=v
            tk.Label(c, textvariable=v, bg=C_PANEL2, fg=col, font=(F_TITLE,22,"bold")).pack(pady=(0,10))

        # TABLEAU
        tb = self._card(root, "Détail par commune", "📋", C_ACCENT)
        tf = tk.Frame(tb, bg=C_CARD); tf.pack(fill="both", expand=True)
        cols = ("Commune","Dep","IDs","Excel","Lignes","PPT","Erreurs","Statut")
        self.tree = ttk.Treeview(tf, columns=cols, show="headings", height=9)
        for c,w in zip(cols,[100,60,50,60,60,60,60,200]):
            self.tree.heading(c, text=c); self.tree.column(c, width=w, anchor="center")
        sbt = ttk.Scrollbar(tf, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sbt.set)
        self.tree.pack(side="left", fill="both", expand=True); sbt.pack(side="left", fill="y")
        actb = tk.Frame(tb, bg=C_CARD); actb.pack(fill="x", pady=(10,0))
        tk.Button(actb, text="🔄 Rafraîchir aperçu", command=self._preview, bg=C_PANEL2, fg=C_TEXT2,
                  font=(F_BODY,9), relief="flat", cursor="hand2", padx=14, pady=5).pack(side="left")
        tk.Button(actb, text="📂 Ouvrir Prod IA", command=self._open_prod, bg=C_PANEL2, fg=C_TEXT2,
                  font=(F_BODY,9), relief="flat", cursor="hand2", padx=14, pady=5).pack(side="left", padx=8)

        tk.Frame(root, bg=C_BG, height=16).pack(fill="x")
        self.after(300, self._preview)
        self.after(400, self._detect_cols)

    def _path_row(self, parent, label, key, is_dir, hint=""):
        row = tk.Frame(parent, bg=C_CARD); row.pack(fill="x", pady=5)
        top = tk.Frame(row, bg=C_CARD); top.pack(fill="x")
        tk.Label(top, text=label, bg=C_CARD, fg=C_TEXT2, font=(F_BODY,9,"bold"),
                 width=20, anchor="w").pack(side="left")
        ent = tk.Entry(top, font=(F_BODY,9), bg="#060a12", fg=C_TEXT, insertbackground="white",
                       relief="flat", bd=6); ent.pack(side="left", fill="x", expand=True, padx=(0,5))
        if self.cfg.get(key): ent.insert(0, self.cfg[key])
        def browse():
            p = filedialog.askdirectory() if is_dir else filedialog.askopenfilename()
            if p:
                ent.delete(0,tk.END); ent.insert(0,p); self.cfg[key]=p; save_config(self.cfg)
                self._preview(); self._detect_cols()
        tk.Button(top, text="📁", width=3, command=browse, bg=C_PANEL2, fg=C_TEXT2, relief="flat",
                  cursor="hand2").pack(side="left")
        if hint: tk.Label(row, text="  "+hint, bg=C_CARD, fg=C_MUTED, font=(F_BODY,8)).pack(anchor="w")
        setattr(self, f"ent_{key}", ent)

    def _check(self, parent, text, var, command):
        tk.Checkbutton(parent, text=f"  {text}", variable=var, bg=C_CARD, fg=C_TEXT2, selectcolor=C_PANEL2,
                       activebackground=C_CARD, activeforeground=C_TEXT, font=(F_BODY,10),
                       relief="flat", cursor="hand2", command=command).pack(side="left", padx=(0,12))

    # ── Mode ID erreur ──
    def _by_ids(self):
        return self.var_mode.get() == "ids"
    def _ids(self):
        return extraire_ids(self.txt_ids.get("1.0", "end"))
    def _ppt_mode(self):
        if self.var_ppt_rempl.get(): return "remplacer"
        if self.var_ppt_modif.get(): return "modifier"
        return None
    def _on_mode(self, preview=True):
        if self._by_ids():
            self.frm_ids.pack(fill="x", pady=(4,0), before=self.btn_run)
            self.btn_run.config(text="🎯  Mettre à jour les ID erreur")
        else:
            self.frm_ids.pack_forget()
            self.btn_run.config(text="📦  Ranger tout Prod IA")
        if preview: self._preview()
    def _on_ppt(self, which):
        # Les deux options sont exclusives (aucune cochée = Excel seul)
        if which == "rempl" and self.var_ppt_rempl.get(): self.var_ppt_modif.set(False)
        if which == "modif" and self.var_ppt_modif.get(): self.var_ppt_rempl.set(False)
        self._save_cfg()
    # ── Colonnes Excel ──
    def _detect_cols(self):
        """Lit les en-têtes en tâche de fond (OneDrive peut devoir télécharger le fichier)."""
        self._save_cfg()
        prod, root = self.ent_prod_ia.get().strip(), self.ent_root_sp.get().strip()
        self.lbl_cols.config(text="  🔍 Détection des colonnes…", fg=C_AMBER)
        res = {}
        threading.Thread(target=lambda: res.update(r=detecter_colonnes(prod, root)), daemon=True).start()
        def attendre():
            if "r" not in res: self.after(200, attendre); return
            self._afficher_cols(*res["r"])
        self.after(200, attendre)
    def _afficher_cols(self, cols, fichier):
        for w in self.frm_cols.winfo_children(): w.destroy()
        self._col_vars = {}
        cols = [c for c in cols if c != COL_ID_ERR]   # clé de correspondance, jamais modifiée
        if not cols:
            self.lbl_cols.config(fg=C_AMBER, text="  ⚠️ Aucun Excel trouvé dans Prod IA ni SharePoint — vérifiez "
                                 "les chemins puis 🔍 Détecter. Sans détection, toutes les colonnes sont mises à jour.")
            return
        self._cols_src = fichier
        saved = self.cfg.get("colonnes_maj")
        n = 4
        for c in range(n): self.frm_cols.columnconfigure(c, weight=1, uniform="cols")
        for i, col in enumerate(cols):
            v = tk.BooleanVar(value=(saved is None or col in saved)); self._col_vars[col] = v
            tk.Checkbutton(self.frm_cols, text=col, variable=v, command=self._on_col, anchor="w",
                           bg=C_CARD, fg=C_TEXT2, selectcolor=C_PANEL2, activebackground=C_CARD,
                           activeforeground=C_TEXT, font=(F_BODY,9), relief="flat",
                           cursor="hand2").grid(row=i//n, column=i%n, sticky="w", padx=(0,8))
        self._on_col()
    def _cocher_cols(self, etat):
        for v in self._col_vars.values(): v.set(etat)
        self._on_col()
    def _on_col(self):
        if not self._col_vars: return
        k = sum(v.get() for v in self._col_vars.values())
        self.lbl_cols.config(fg=C_ACCENT if k else C_AMBER,
            text=f"  {k}/{len(self._col_vars)} colonne(s) cochée(s) · détectées depuis "
                 f"{os.path.basename(self._cols_src)}" + ("" if k else " — l'Excel ne sera pas modifié")
                 + " · (« ID erreur » sert à retrouver les lignes. Les PPT lisent toujours "
                   "« Remarque » et « Lien Street View » de Prod IA)")
        self._save_cfg()
    def _colonnes_maj(self):
        """Colonnes cochées, ou None si la détection n'a rien trouvé (= toutes)."""
        if not self._col_vars: return None
        return [c for c, v in self._col_vars.items() if v.get()]

    def _on_ids_modified(self, _e=None):
        if not self.txt_ids.edit_modified(): return
        self.txt_ids.edit_modified(False)
        if self.txt_ids.get("1.0", "end-1c"): self.lbl_ids_ph.place_forget()
        else: self.lbl_ids_ph.place(x=9, y=7)
        self._schedule_preview()
    def _schedule_preview(self):
        if getattr(self, "_prev_job", None): self.after_cancel(self._prev_job)
        self._prev_job = self.after(400, self._preview)

    # ── Helpers ──
    def _clear_log(self):
        self.txt_log.config(state="normal"); self.txt_log.delete("1.0","end"); self.txt_log.config(state="disabled")
    def _toggle_debug(self):
        if self._debug_visible:
            self.page_debug.pack_forget(); self.page_main.pack(fill="both", expand=True)
            self.btn_debug.config(text="🐛 Debug"); self._debug_visible=False
        else:
            self.page_main.pack_forget(); self.page_debug.pack(fill="both", expand=True)
            self.btn_debug.config(text="◀ Retour"); self._debug_visible=True
    def _show_debug(self):
        if not self._debug_visible: self._toggle_debug()
    def _log(self, msg, color=C_GREEN):
        try: self._logq.append((msg, color))
        except: pass
    def _log_ui(self, msg, color=C_GREEN):
        self.txt_log.config(state="normal")
        tag=f"c{color.replace('#','')}"; self.txt_log.tag_config(tag, foreground=color)
        self.txt_log.insert(tk.END, msg+"\n", tag); self.txt_log.see(tk.END)
        self.txt_log.config(state="disabled")
    def _set_status(self, txt, color=C_MUTED, bg="#0a0f1a"):
        self.lbl_status.config(text=txt, fg=color, bg=bg); self.frm_status.config(bg=bg)
        self.lbl_prog.config(bg=bg); self.lbl_timer.config(bg=bg); self.frm_timeline.config(bg=bg)
        for w in self.frm_timeline.winfo_children():
            try: w.config(bg=bg)
            except: pass
        self.update_idletasks()
    def _set_hdr(self, txt, col):
        self.hdr_state.config(text=f"●  {txt}", fg=col); self.update_idletasks()
    def _save_cfg(self):
        for key in ["root_sp","prod_ia"]:
            ent = getattr(self, f"ent_{key}", None)
            if ent:
                v=ent.get().strip()
                if v: self.cfg[key]=v
        self.cfg["supprimer"]=bool(self.var_supp.get())
        if hasattr(self, "var_mode"):
            self.cfg["mode"]=self.var_mode.get()
            self.cfg["ppt_remplacer"]=bool(self.var_ppt_rempl.get())
            self.cfg["ppt_modifier"]=bool(self.var_ppt_modif.get())
        if getattr(self, "_col_vars", None):
            # garde aussi les choix des colonnes absentes du fichier détecté
            anciens = self.cfg.get("colonnes_maj") or []
            self.cfg["colonnes_maj"] = ([c for c, v in self._col_vars.items() if v.get()]
                                        + [c for c in anciens if c not in self._col_vars])
        save_config(self.cfg)
    def _open_prod(self):
        d=self.ent_prod_ia.get().strip()
        if d and os.path.isdir(d):
            try: os.startfile(d)
            except: messagebox.showinfo("Dossier", d)
        else: messagebox.showinfo("Dossier","Dossier Prod IA introuvable.")
    def _fmt(self, s):
        s=int(s); return f"{s}s" if s<60 else f"{s//60}m{s%60:02d}s"
    def _tick(self):
        if self._t_start is None: return
        self.lbl_timer.config(text=f"⏱ {self._fmt(_time.time()-self._t_start)}")
        self.after(500, self._tick)

    # ── Callbacks worker → state ──
    def _cb_phase(self, key, etat):
        self._state["phase"] = (key, etat)

    # ── Timeline ──
    def _build_timeline(self, phases):
        for w in self.frm_timeline.winfo_children(): w.destroy()
        self._phase_w = {}
        for i,(key,lbl) in enumerate(phases):
            chip = tk.Frame(self.frm_timeline, bg="#0a0f1a"); chip.pack(side="left", padx=(0,3))
            dot = tk.Label(chip, text="○", bg="#0a0f1a", fg=C_MUTED, font=(F_BODY,11)); dot.pack(side="left")
            txt = tk.Label(chip, text=lbl, bg="#0a0f1a", fg=C_MUTED, font=(F_BODY,8)); txt.pack(side="left", padx=(2,0))
            self._phase_w[key]=(dot,txt)
            if i < len(phases)-1:
                tk.Label(self.frm_timeline, text="→", bg="#0a0f1a", fg=C_BORDER2, font=(F_BODY,8)).pack(side="left", padx=2)
    def _reset_timeline_ui(self):
        if not hasattr(self, "_phase_w"): return
        for dot,txt in self._phase_w.values():
            dot.config(text="○", fg=C_MUTED); txt.config(fg=C_MUTED, font=(F_BODY,8))
    def _set_phase_ui(self, key, etat):
        w = getattr(self, "_phase_w", {}).get(key)
        if not w: return
        dot,txt = w
        if etat=="run":  dot.config(text="◉", fg=C_AMBER); txt.config(fg=C_AMBER, font=(F_BODY,8,"bold"))
        elif etat=="ok": dot.config(text="●", fg=C_GREEN); txt.config(fg=C_GREEN, font=(F_BODY,8))
        elif etat=="err":dot.config(text="✕", fg=C_RED);   txt.config(fg=C_RED, font=(F_BODY,8,"bold"))

    # ── Poller (main thread) ──
    def _poll(self):
        try:
            while self._logq:
                msg, color = self._logq.pop(0)
                self._log_ui(msg, color)
        except: pass
        st = self._state
        if "prog" in st: self.lbl_prog.config(text=st["prog"])
        if "phase" in st:
            key, etat = st["phase"]; self._set_phase_ui(key, etat); st.pop("phase", None)
        if "gpos" in st and "gmax" in st:
            self.progress.config(maximum=st["gmax"], value=st["gpos"])
        if st.get("row_insert"):
            commune = st.pop("row_insert")
            self._reset_timeline_ui()
            iid = f"row_{commune}"
            if not self.tree.exists(iid):
                self.tree.insert("","end", iid=iid, values=(commune,"","","","","","","🔄 En cours"), tags=("run",))
                self.tree.tag_configure("run", foreground=C_AMBER)
        if st.get("row_update"):
            vals, ok = st.pop("row_update")
            iid = f"row_{vals[0]}"
            if self.tree.exists(iid):
                tag = "ok" if ok else "err"
                self.tree.item(iid, values=vals, tags=(tag,))
                self.tree.tag_configure("ok", foreground=C_GREEN)
                self.tree.tag_configure("err", foreground=C_RED)
        if "stats" in st:
            for k,v in st["stats"].items():
                if k in self._stat_vars: self._stat_vars[k].set(str(v))
        if st.get("done"):
            self._finish(st.pop("done")); self._running=False; return
        if self._running:
            self.after(250, self._poll)

    # ── Aperçu ──
    def _scan(self):
        """Liste les dossiers commune dans Prod IA."""
        return lister_communes_prod(self.ent_prod_ia.get().strip())

    def _preview(self):
        if self._running: return
        self._save_cfg()
        dispo = self._scan()
        root_sp = self.ent_root_sp.get().strip()
        for r in self.tree.get_children(): self.tree.delete(r)
        self.tree.tag_configure("todo", foreground=C_AMBER)
        self.tree.tag_configure("no", foreground=C_RED)
        if self._by_ids():
            return self._preview_ids(dispo, root_sp)
        n_ok=0; n_no=0
        for com in sorted(dispo.keys()):
            dep = get_dep(com)
            dest_ok = os.path.isdir(os.path.join(root_sp, dep, com))
            if dest_ok:
                self.tree.insert("","end", values=(com,dep,"—","","—","","","⏳ À ranger"), tags=("todo",)); n_ok+=1
            else:
                self.tree.insert("","end", values=(com,dep,"—","","—","","","❌ Dest. absente"), tags=("no",)); n_no+=1
        self._stat_vars["communes"].set(str(len(dispo)))
        if len(dispo)==0:
            self._set_status("📭 Prod IA est vide — rien à ranger", C_MUTED)
        else:
            self._set_status(f"📁 {n_ok} commune(s) à ranger" + (f" · {n_no} sans destination" if n_no else ""),
                             C_ACCENT if n_ok else C_AMBER)

    def _preview_ids(self, dispo, root_sp):
        groupes = self._ids()
        nb = sum(len(v) for v in groupes.values())
        self.lbl_ids.config(text=f"  {nb} ID erreur détecté(s) sur {len(groupes)} commune(s)" if nb
                            else "  Aucun ID erreur détecté (format attendu : 74143_1)",
                            fg=C_ACCENT if nb else C_MUTED)
        n_ok=0; n_no=0
        for com, ids in groupes.items():
            dep = get_dep(com)
            if com not in dispo: statut = "❌ Absent de Prod IA"
            elif not os.path.isdir(os.path.join(root_sp, dep, com)): statut = "❌ Dest. absente"
            else: statut = "⏳ À mettre à jour"
            ok = statut.startswith("⏳"); n_ok += ok; n_no += not ok
            self.tree.insert("","end", values=(com,dep,len(ids),"","","","",statut), tags=("todo" if ok else "no",))
        self._stat_vars["communes"].set(str(len(groupes)))
        if not groupes:
            self._set_status("✍️ Collez des ID erreur pour préparer la mise à jour", C_MUTED)
        else:
            self._set_status(f"🎯 {nb} ID sur {n_ok} commune(s) prête(s)" + (f" · {n_no} non traitable(s)" if n_no else ""),
                             C_ACCENT if n_ok else C_AMBER)

    # ── Lancer le rangement ──
    def _run(self):
        self._save_cfg()
        root_sp = self.ent_root_sp.get().strip()
        prod_ia = self.ent_prod_ia.get().strip()
        supprimer = bool(self.var_supp.get())
        if not os.path.isdir(root_sp):
            messagebox.showwarning("Erreur","Racine SharePoint introuvable."); return
        if not os.path.isdir(prod_ia):
            messagebox.showwarning("Erreur","Dossier Prod IA introuvable."); return
        if self._by_ids():
            return self._run_ids(root_sp, prod_ia, supprimer)

        dispo = self._scan()
        if not dispo:
            messagebox.showinfo("Rien à faire","Prod IA est vide — aucun dossier commune à ranger."); return

        codes = sorted(dispo.keys())
        supp_txt = "supprimés de Prod IA" if supprimer else "conservés dans Prod IA"
        if not messagebox.askyesno("Rangement",
                f"Ranger {len(codes)} commune(s) : {', '.join(codes)} ?\n\n"
                "• Excel → Dep/commune/ (remplace le vierge)\n"
                "• PPT → Dep/commune/analyse/\n"
                f"• Dossiers {supp_txt} après rangement\n\n"
                "⚠️ Cette action écrase les fichiers SharePoint."): return

        self._start_ui(len(codes), "Rangement", "📦 RANGEMENT EN COURS")

        def do():
            tot={"mode":"commune","communes":0,"lignes":"—","excel":0,"ppt":0,"err":0}
            self._log(f"\n📦 RANGEMENT de {len(codes)} commune(s)", C_ACCENT)
            self._log(f"📂 Source : {prod_ia}", C_MUTED)
            self._log(f"📂 Cible  : {root_sp}\\DepXX\\commune", C_MUTED)
            self._log("─"*50, C_MUTED)
            for i,commune in enumerate(codes,1):
                self._state["gpos"]=i-1
                self._state["prog"]=f"Commune {i}/{len(codes)} — {commune}"
                self._state["row_insert"]=commune
                self._log(f"\n📍 Commune {commune}  ({i}/{len(codes)})", C_AMBER)
                stats = ranger_commune(prod_ia, root_sp, commune, dispo[commune],
                                       supprimer, self._log, phase_fn=self._cb_phase)
                self._state["gpos"]=i
                if stats is not None:
                    tot["communes"]+=1; tot["excel"]+=stats["excel"]
                    tot["ppt"]+=stats["ppt"]; tot["err"]+=stats["err"]
                    ok = stats["err"]==0
                    statut = "✅ Rangé" if ok else f"⚠️ {stats['err']} erreur(s)"
                    self._state["row_update"] = ((commune,get_dep(commune),"—",stats["excel"],"—",
                                                  stats["ppt"],stats["err"],statut), ok)
                else:
                    tot["err"]+=1
                    self._state["row_update"] = ((commune,get_dep(commune),"—","","—","","","❌ Dest. absente"), False)
                self._state["stats"]=dict(tot)
            el=_time.time()-self._t_start
            self._log(f"\n🎉 TERMINÉ — {tot['communes']} commune(s) en {self._fmt(el)}", C_GREEN)
            self._log(f"  Excel:{tot['excel']} · PPT:{tot['ppt']} · Erreurs:{tot['err']}", C_MUTED)
            self._supprimer_adeposer(prod_ia)
            tot["_el"]=el; self._state["done"]=tot

        threading.Thread(target=do, daemon=True).start()

    def _run_ids(self, root_sp, prod_ia, supprimer):
        groupes = self._ids()
        if not groupes:
            messagebox.showinfo("Rien à faire","Aucun ID erreur valide détecté (format attendu : 74143_1)."); return
        ppt_mode = self._ppt_mode()
        dispo = self._scan()
        codes = list(groupes)
        nb = sum(len(v) for v in groupes.values())
        ppt_txt = {"remplacer": "PPT → remplacés par ceux de Prod IA",
                   "modifier":  "PPT → Analyse et Lien Street View modifiés (colonnes Remarque / Lien Street View)",
                   None:        "PPT → non traités"}[ppt_mode]
        colonnes = self._colonnes_maj()
        if colonnes is None:
            excel_txt = "Excel → lignes de ces ID, toutes les colonnes (colonnes non détectées)"
        elif not colonnes:
            excel_txt = "Excel → non modifié (aucune colonne cochée)"
        else:
            apercu = ", ".join(colonnes[:6]) + (f" … (+{len(colonnes)-6})" if len(colonnes) > 6 else "")
            excel_txt = f"Excel → lignes de ces ID, {len(colonnes)} colonne(s) : {apercu}"
        if colonnes == [] and ppt_mode is None:
            messagebox.showinfo("Rien à faire", "Aucune colonne Excel cochée et aucune option PPT : rien à mettre à jour."); return
        supp_txt = ("supprimés de Prod IA (s'ils ne contiennent pas d'autres ID)" if supprimer
                    else "conservés dans Prod IA")
        if not messagebox.askyesno("Mise à jour par ID erreur",
                f"Mettre à jour {nb} ID erreur sur {len(codes)} commune(s) : {', '.join(codes)} ?\n\n"
                f"• {excel_txt}\n"
                f"• {ppt_txt}\n"
                f"• Dossiers {supp_txt}\n\n"
                "⚠️ Cette action écrase des données SharePoint."): return

        self._start_ui(len(codes), "Mise à jour ID", "🎯 MISE À JOUR PAR ID EN COURS")

        def do():
            tot={"mode":"ids","communes":0,"lignes":0,"excel":0,"ppt":0,"err":0}
            self._log(f"\n🎯 MISE À JOUR de {nb} ID erreur sur {len(codes)} commune(s)", C_ACCENT)
            self._log(f"📂 Source : {prod_ia}", C_MUTED)
            self._log(f"📂 Cible  : {root_sp}\\DepXX\\commune", C_MUTED)
            self._log(f"📊 {excel_txt}", C_MUTED)
            self._log(f"🖼️  {ppt_txt}", C_MUTED)
            self._log("─"*50, C_MUTED)
            for i,commune in enumerate(codes,1):
                ids = groupes[commune]; dep = get_dep(commune)
                self._state["gpos"]=i-1
                self._state["prog"]=f"Commune {i}/{len(codes)} — {commune} ({len(ids)} ID)"
                self._state["row_insert"]=commune
                self._log(f"\n📍 Commune {commune}  ({i}/{len(codes)}) — {', '.join(ids)}", C_AMBER)
                stats = ranger_ids(root_sp, commune, dispo.get(commune), ids, ppt_mode,
                                   supprimer, self._log, phase_fn=self._cb_phase, colonnes=colonnes)
                self._state["gpos"]=i
                if stats is not None:
                    tot["communes"]+=1
                    for k in ("lignes","excel","ppt","err"): tot[k]+=stats[k]
                    ok = stats["err"]==0
                    statut = "✅ Mis à jour" if ok else f"⚠️ {stats['err']} erreur(s)"
                    self._state["row_update"] = ((commune,dep,len(ids),stats["excel"],stats["lignes"],
                                                  stats["ppt"],stats["err"],statut), ok)
                else:
                    tot["err"]+=1
                    self._state["row_update"] = ((commune,dep,len(ids),"","","","","❌ Non traité"), False)
                self._state["stats"]=dict(tot)
            el=_time.time()-self._t_start
            self._log(f"\n🎉 TERMINÉ — {tot['communes']} commune(s) en {self._fmt(el)}", C_GREEN)
            self._log(f"  Lignes:{tot['lignes']} · PPT:{tot['ppt']} · Erreurs:{tot['err']}", C_MUTED)
            self._supprimer_adeposer(prod_ia)
            tot["_el"]=el; self._state["done"]=tot

        threading.Thread(target=do, daemon=True).start()

    def _start_ui(self, n, hdr, status):
        for r in self.tree.get_children(): self.tree.delete(r)
        for k in self._stat_vars: self._stat_vars[k].set("...")
        self.btn_run.config(state="disabled")
        self._show_debug(); self._set_hdr(hdr, C_GREEN)
        self._build_timeline(PHASES)
        self._set_status(status, C_GREEN, "#0a2018")
        self.progress.config(style="green.Horizontal.TProgressbar")
        self.progress["value"]=0; self.progress["maximum"]=n
        self._t_start=_time.time(); self._tick()
        self._logq=[]; self._state={"gmax":n,"gpos":0}; self._running=True
        self.after(250, self._poll)

    def _supprimer_adeposer(self, prod_ia):
        """Supprime le dossier "À déposer" entier s'il est vide (tout rangé)."""
        if not (bool(self.var_supp.get()) and self.cfg.get("supp_adeposer", True)): return
        supprimer_si_vide(prod_ia, self._log)

    def _finish(self, tot):
        el=tot.get("_el",0); self._t_start=None
        self.btn_run.config(state="normal"); self._set_hdr("Prêt", C_GREEN)
        par_ids = tot.get("mode") == "ids"
        detail = f"{tot['lignes']} ligne(s)" if par_ids else f"{tot['ppt']} PPT"
        self._set_status(f"🎉 TERMINÉ — {tot['communes']} commune(s) · {detail} en {self._fmt(el)}", C_GREEN, "#0a2018")
        try: self.bell()
        except: pass
        if par_ids:
            corps = (f"✅ {tot['communes']} commune(s) mise(s) à jour en {self._fmt(el)}\n\n"
                     f"Lignes Excel remplacées : {tot['lignes']}\nPPT traités : {tot['ppt']}\n")
        else:
            corps = (f"✅ {tot['communes']} commune(s) rangée(s) en {self._fmt(el)}\n\n"
                     f"Excel remplacés : {tot['excel']}\nPPT déposés : {tot['ppt']}\n")
        messagebox.showinfo("Terminé", corps + f"Erreurs : {tot['err']}\n\n"
                            "OneDrive synchronise vers SharePoint.")
