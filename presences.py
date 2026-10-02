#!/usr/bin/env python3
"""Attendance tracking app — generates a styled Excel report."""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from collections import OrderedDict

import base64
import io
import os
import queue
import re
import smtplib
import ssl
import subprocess
import sys
import tempfile
import threading
from email.message import EmailMessage

import openpyxl
import qrcode
from PIL import Image, ImageDraw, ImageFont
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


def _resource_path(filename: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, filename)


class AttendanceApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Présences")
        self.root.geometry("700x860")
        self.root.resizable(True, True)

        # Date frozen at startup
        self.event_date = datetime.now().strftime("%d/%m/%Y")

        # Main records: appt -> {debut, fin, comment}
        self.records: OrderedDict[str, dict] = OrderedDict()

        # Pending queues: list of (appt, time, comment)
        self.debut_pending: list[tuple[str, str, str]] = []
        self.fin_pending:   list[tuple[str, str, str]] = []

        # Time-capture guards
        self._debut_time_set = False
        self._fin_time_set   = False

        self._build_ui()

    # ── UI construction ───────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        # Logo header + window/taskbar icon
        try:
            self._logo_img = tk.PhotoImage(file=_resource_path("user.png"))
            self.root.iconphoto(True, self._logo_img)
            logo_lbl = ttk.Label(self.root, image=self._logo_img)
            logo_lbl.pack(pady=(10, 2))
        except Exception:
            pass

        # Main tabs: Présences / Absences
        main_nb = ttk.Notebook(self.root)
        main_nb.pack(fill="both", expand=True)
        pres = ttk.Frame(main_nb)
        main_nb.add(pres, text="Présences")
        abs_tab = ttk.Frame(main_nb)
        main_nb.add(abs_tab, text="Absences")
        self.absences = AbsencesTab(abs_tab, self)
        qr_tab = ttk.Frame(main_nb)
        main_nb.add(qr_tab, text="Générer Codes QR")
        self.qr = QrTab(qr_tab, self)

        # Event header
        info = ttk.LabelFrame(pres, text="Informations", padding=10)
        info.pack(fill="x", padx=12, pady=(10, 4))
        info.columnconfigure(1, weight=1)

        ttk.Label(info, text="Évènement :").grid(row=0, column=0, sticky="w")
        self.event_var = tk.StringVar()
        ttk.Entry(info, textvariable=self.event_var, width=34).grid(
            row=0, column=1, sticky="ew", padx=(8, 20)
        )
        ttk.Label(info, text="Date :").grid(row=0, column=2, sticky="e")
        ttk.Label(
            info,
            text=self.event_date,
            relief="sunken",
            width=12,
            anchor="center",
            padding=(4, 2),
        ).grid(row=0, column=3, padx=(6, 0))

        # Tabs
        nb = ttk.Notebook(pres)
        nb.pack(fill="x", padx=12, pady=6)

        tab1 = ttk.Frame(nb, padding=14)
        tab2 = ttk.Frame(nb, padding=14)
        nb.add(tab1, text="Début")
        nb.add(tab2, text="Fin")

        self._build_tab(
            tab1,
            appt_var_name="debut_appt_var",
            time_var_name="debut_time_var",
            comment_var_name="debut_comment_var",
            time_col="Début",
            on_change=self._on_debut_change,
            on_add=self._add_to_debut,
            on_submit=self._record_debut,
            pending_tree_attr="debut_tree",
        )
        self._build_tab(
            tab2,
            appt_var_name="fin_appt_var",
            time_var_name="fin_time_var",
            comment_var_name="fin_comment_var",
            time_col="Fin",
            on_change=self._on_fin_change,
            on_add=self._add_to_fin,
            on_submit=self._record_fin,
            pending_tree_attr="fin_tree",
        )

        # Action bar
        action_bar = ttk.Frame(pres)
        action_bar.pack(fill="x", padx=12, pady=(4, 2))

        self.count_label = ttk.Label(
            action_bar, text="Nombre de présents : 0", font=("TkDefaultFont", 10, "bold")
        )
        self.count_label.pack(side="left")

        ttk.Button(
            action_bar,
            text="Générer",
            command=self._generate_excel,
            padding=(18, 6),
        ).pack(side="right")

        # Search bar
        search_frame = ttk.Frame(pres)
        search_frame.pack(fill="x", padx=12, pady=(2, 0))

        ttk.Label(search_frame, text="Rechercher :").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._refresh_table())
        ttk.Entry(search_frame, textvariable=self.search_var, width=28).pack(
            side="left", padx=6
        )
        ttk.Button(
            search_frame,
            text="Effacer",
            command=lambda: self.search_var.set(""),
        ).pack(side="left")

        # Main records table
        tbl = ttk.LabelFrame(pres, text="Présences enregistrées", padding=6)
        tbl.pack(fill="both", expand=True, padx=12, pady=(4, 10))

        cols = ("Appartement", "Début", "Fin", "Commentaire")
        self.tree = ttk.Treeview(tbl, columns=cols, show="headings", height=7)
        self.tree.heading("Appartement", text="Appartement")
        self.tree.heading("Début",       text="Début")
        self.tree.heading("Fin",         text="Fin")
        self.tree.heading("Commentaire", text="Commentaire")
        self.tree.column("Appartement", width=140, anchor="center")
        self.tree.column("Début",       width=80,  anchor="center")
        self.tree.column("Fin",         width=80,  anchor="center")
        self.tree.column("Commentaire", width=280, anchor="w")

        sb = ttk.Scrollbar(tbl, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # Copyright footer
        ttk.Label(
            self.root,
            text="© 2026 Djaouida Kharchi",
            foreground="gray",
            font=("TkDefaultFont", 9),
        ).pack(pady=(0, 6))

    def _build_tab(
        self,
        parent: ttk.Frame,
        *,
        appt_var_name: str,
        time_var_name: str,
        comment_var_name: str,
        time_col: str,
        on_change,
        on_add,
        on_submit,
        pending_tree_attr: str,
    ) -> None:
        parent.columnconfigure(1, weight=1)

        # Row 0: Appartement + time
        ttk.Label(parent, text="Appartement :").grid(row=0, column=0, sticky="w", pady=3)
        appt_var = tk.StringVar()
        appt_var.trace_add("write", on_change)
        setattr(self, appt_var_name, appt_var)

        entry = ttk.Entry(parent, textvariable=appt_var, width=24)
        entry.grid(row=0, column=1, padx=(8, 20), pady=3, sticky="w")
        entry.bind("<Return>", on_add)

        ttk.Label(parent, text=f"{time_col} :").grid(row=0, column=2, sticky="w")
        time_var = tk.StringVar()
        setattr(self, time_var_name, time_var)
        ttk.Entry(parent, textvariable=time_var, width=10, state="readonly").grid(
            row=0, column=3, padx=(4, 0), pady=3, sticky="w"
        )

        # Row 1: Comment
        ttk.Label(parent, text="Commentaire :").grid(row=1, column=0, sticky="w", pady=3)
        comment_var = tk.StringVar()
        setattr(self, comment_var_name, comment_var)
        comment_entry = ttk.Entry(parent, textvariable=comment_var, width=48)
        comment_entry.grid(row=1, column=1, columnspan=3, padx=(8, 0), pady=3, sticky="ew")
        comment_entry.bind("<Return>", on_add)

        # Row 2: Hint
        ttk.Label(
            parent,
            text="Entrée pour ajouter · Suppr pour retirer",
            foreground="gray",
        ).grid(row=2, column=0, columnspan=4, sticky="w", pady=(2, 4))

        # Row 3: Pending treeview
        pending_tree = ttk.Treeview(
            parent,
            columns=("Appartement", time_col, "Commentaire"),
            show="headings",
            height=4,
        )
        pending_tree.heading("Appartement", text="Appartement")
        pending_tree.heading(time_col,      text=time_col)
        pending_tree.heading("Commentaire", text="Commentaire")
        pending_tree.column("Appartement", width=140, anchor="center")
        pending_tree.column(time_col,      width=80,  anchor="center")
        pending_tree.column("Commentaire", width=280, anchor="w")
        pending_tree.grid(row=3, column=0, columnspan=4, sticky="ew", pady=4)
        setattr(self, pending_tree_attr, pending_tree)

        pending_list = self.debut_pending if "debut" in appt_var_name else self.fin_pending
        for seq in ("<Delete>", "<BackSpace>"):
            pending_tree.bind(
                seq,
                lambda e, t=pending_tree, p=pending_list: self._delete_selected(t, p),
            )

        # Row 4: Buttons
        btn_frame = ttk.Frame(parent)
        btn_frame.grid(row=4, column=0, columnspan=4, sticky="e", pady=(4, 0))

        ttk.Button(
            btn_frame,
            text="Supprimer la sélection",
            command=lambda t=pending_tree, p=pending_list: self._delete_selected(t, p),
        ).pack(side="left", padx=(0, 8))

        ttk.Button(btn_frame, text="Enregistrer", command=on_submit).pack(side="left")

    # ── Auto-fill time on first character ─────────────────────────────────────

    def _on_debut_change(self, *_) -> None:
        val = self.debut_appt_var.get()
        if val and not self._debut_time_set:
            self.debut_time_var.set(datetime.now().strftime("%H:%M"))
            self._debut_time_set = True
        elif not val:
            self._debut_time_set = False
            self.debut_time_var.set("")

    def _on_fin_change(self, *_) -> None:
        val = self.fin_appt_var.get()
        if val and not self._fin_time_set:
            self.fin_time_var.set(datetime.now().strftime("%H:%M"))
            self._fin_time_set = True
        elif not val:
            self._fin_time_set = False
            self.fin_time_var.set("")

    # ── Add to pending queue (Enter key) ──────────────────────────────────────

    def _add_pending(self, pending: list, tree: ttk.Treeview, appt: str, t: str, comment: str) -> None:
        # Same appartement (case-insensitive) in the same tab: one row, comments merged
        for i, (a, t0, c0) in enumerate(pending):
            if a.lower() == appt.lower():
                c = self._join_comments(c0, comment)
                pending[i] = (a, t0, c)
                tree.item(tree.get_children()[i], values=(a, t0, c))
                return
        pending.append((appt, t, comment))
        tree.insert("", "end", values=(appt, t, comment))

    def _add_to_debut(self, *_) -> None:
        appt = self.debut_appt_var.get().strip()
        if not appt:
            return
        t       = self.debut_time_var.get() or datetime.now().strftime("%H:%M")
        comment = self.debut_comment_var.get().strip()
        self._add_pending(self.debut_pending, self.debut_tree, appt, t, comment)
        self.debut_appt_var.set("")
        self.debut_time_var.set("")
        self.debut_comment_var.set("")
        self._debut_time_set = False

    def _add_to_fin(self, *_) -> None:
        appt = self.fin_appt_var.get().strip()
        if not appt:
            return
        t       = self.fin_time_var.get() or datetime.now().strftime("%H:%M")
        comment = self.fin_comment_var.get().strip()
        self._add_pending(self.fin_pending, self.fin_tree, appt, t, comment)
        self.fin_appt_var.set("")
        self.fin_time_var.set("")
        self.fin_comment_var.set("")
        self._fin_time_set = False

    # ── Delete selected row from pending treeview ─────────────────────────────

    def _delete_selected(self, tree: ttk.Treeview, pending: list) -> None:
        for item in tree.selection():
            idx = tree.index(item)
            tree.delete(item)
            if 0 <= idx < len(pending):
                pending.pop(idx)

    # ── Enregistrer — flush pending queue to main records ─────────────────────

    @staticmethod
    def _join_comments(old: str, new: str) -> str:
        if not new or new in [c.strip() for c in old.split(",")]:
            return old
        return f"{old}, {new}" if old else new

    def _merge(self, appt: str, key: str, t: str, comment: str) -> None:
        # Appartement is case-insensitive; keep the first spelling entered
        appt = next((k for k in self.records if k.lower() == appt.lower()), appt)
        rec = self.records.setdefault(
            appt, {"debut": "", "fin": "", "debut_comment": "", "fin_comment": ""}
        )
        rec[key] = t
        rec[f"{key}_comment"] = self._join_comments(rec[f"{key}_comment"], comment)

    @staticmethod
    def _comment(rec: dict) -> str:
        return ", ".join(c for c in (rec["debut_comment"], rec["fin_comment"]) if c)

    def _flush_debut(self) -> None:
        self._add_to_debut()
        for appt, t, comment in self.debut_pending:
            self._merge(appt, "debut", t, comment)
        self.debut_pending.clear()
        self.debut_tree.delete(*self.debut_tree.get_children())

    def _flush_fin(self) -> None:
        self._add_to_fin()
        for appt, t, comment in self.fin_pending:
            self._merge(appt, "fin", t, comment)
        self.fin_pending.clear()
        self.fin_tree.delete(*self.fin_tree.get_children())

    def _record_debut(self, *_) -> None:
        self._add_to_debut()
        if not self.debut_pending:
            messagebox.showwarning("Attention", "Aucun appartement à enregistrer.")
            return
        self._flush_debut()
        self._refresh_table()

    def _record_fin(self, *_) -> None:
        self._add_to_fin()
        if not self.fin_pending:
            messagebox.showwarning("Attention", "Aucun appartement à enregistrer.")
            return
        self._flush_fin()
        self._refresh_table()

    # ── Refresh main table (applies search filter) ────────────────────────────

    def _refresh_table(self) -> None:
        self.tree.delete(*self.tree.get_children())
        query = self.search_var.get().strip().lower()
        for appt, times in self.records.items():
            if query and query not in appt.lower():
                continue
            self.tree.insert(
                "", "end",
                values=(appt, times["debut"], times["fin"], self._comment(times)),
            )
        self.count_label.config(text=f"Nombre de présents : {len(self.records)}")

    # ── Excel export ──────────────────────────────────────────────────────────

    def _generate_excel(self) -> None:
        # Include anything not yet saved with "Enregistrer" in either tab
        self._flush_debut()
        self._flush_fin()
        self._refresh_table()
        if not self.records:
            messagebox.showwarning("Attention", "Aucune présence enregistrée.")
            return

        event_name = self.event_var.get().strip()
        safe    = event_name.replace(" ", "_") or "evenement"
        default = f"presences_{safe}_{self.event_date.replace('/', '-')}.xlsx"

        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Fichier Excel", "*.xlsx")],
            initialfile=default,
            title="Enregistrer le fichier Excel",
        )
        self.root.lift()
        self.root.focus_force()
        if not path:
            return

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Présences"

        bold     = Font(bold=True)
        hdr_font = Font(bold=True, color="FFFFFF", size=11)
        hdr_fill = PatternFill("solid", fgColor="2E5FA3")
        center   = Alignment(horizontal="center", vertical="center")
        left     = Alignment(horizontal="left",   vertical="center")
        thin     = Side(style="thin")
        border   = Border(left=thin, right=thin, top=thin, bottom=thin)

        def cell(row, col, value, *, font=None, fill=None, align=None, brd=None):
            c = ws.cell(row=row, column=col, value=value)
            if font:  c.font      = font
            if fill:  c.fill      = fill
            if align: c.alignment = align
            if brd:   c.border    = brd
            return c

        cell(1, 1, "Évènement :", font=bold)
        cell(1, 2, event_name)
        cell(1, 4, "Date :", font=bold)
        cell(1, 5, self.event_date)

        cell(3, 1, "Nombre de présents :", font=bold)
        cell(3, 2, len(self.records))

        for col, header in enumerate(["Appartement", "Début", "Fin", "Commentaire"], start=1):
            cell(5, col, header, font=hdr_font, fill=hdr_fill, align=center, brd=border)

        for row_idx, (appt, times) in enumerate(self.records.items(), start=6):
            cell(row_idx, 1, appt,                       align=center, brd=border)
            cell(row_idx, 2, times["debut"],              align=center, brd=border)
            cell(row_idx, 3, times["fin"],                align=center, brd=border)
            cell(row_idx, 4, self._comment(times),   align=left,   brd=border)

        for col, width in zip("ABCDE", [22, 12, 12, 35, 14]):
            ws.column_dimensions[col].width = width

        wb.save(path)
        messagebox.showinfo("Succès", f"Fichier généré :\n{path}")
        self.root.lift()
        self.root.focus_force()


def _natural_key(text: str):
    return [int(p) if p.isdigit() else p for p in re.split(r"(\d+)", text.lower())]


_EVENT_PATTERNS = (
    ("AGA", r"aga"),
    ("AGE", r"age"),
    ("AG", r"ag"),
    ("Corvée", r"corv[ée]es?"),
)


def _event_type(*texts: str) -> str:
    """Canonical event type (AGA, AGE, AG, Corvée) found in the first matching text."""
    for text in texts:
        for canon, pat in _EVENT_PATTERNS:
            if re.search(rf"(?<![A-Za-z]){pat}(?![A-Za-z])", text or "", re.I):
                return canon
    return ""


def _read_members(wb) -> list[tuple[str, str, str, str]]:
    """(appartment, prénom, nom, courriel) from the members workbook.

    Each sheet has a header row with IMM/NUM; the number is in NUM, except on the
    Maisonnettes sheet where it is in IMM. Vacant appartments are skipped.
    Falls back to column A for plain lists.
    """
    out: list[tuple[str, str, str, str]] = []
    found = False
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        for h, row in enumerate(rows[:15]):
            heads = [str(c).strip().upper() if c is not None else "" for c in row]
            if "IMM" not in heads and "NUM" not in heads:
                continue
            want = "IMM" if "maison" in ws.title.lower() else "NUM"
            col = heads.index(want) if want in heads else heads.index("IMM" if want == "NUM" else "NUM")
            found = True

            def cell(r, name):
                i = heads.index(name) if name in heads else -1
                v = r[i] if 0 <= i < len(r) else None
                return str(v).strip() if v is not None else ""

            for r in rows[h + 1:]:
                v = r[col] if col < len(r) else None
                if isinstance(v, float) and v.is_integer():
                    v = int(v)
                if v is not None and str(v).strip():
                    prenom, nom = cell(r, "PRENOM"), cell(r, "NOM")
                    if "vacant" in prenom.lower() or "vacant" in nom.lower():
                        continue  # empty appartment
                    out.append((str(v).strip(), prenom, nom, cell(r, "COURRIEL")))
            break
    if found:
        return out
    for (v,) in wb.active.iter_rows(max_col=1, values_only=True):
        if v is not None and str(v).strip():
            out.append((str(v).strip(), "", "", ""))
    if out and out[0][0].lower().startswith("appartement"):
        out = out[1:]
    return out


def _apt_key(text: str) -> str:
    """Appartments are compared on their first 4 characters, ignoring case."""
    return text.strip()[:4].lower()


class AbsencesTab:
    """Lists appartments missing from the presences list and exports them."""

    def __init__(self, parent: ttk.Frame, app: "AttendanceApp") -> None:
        self.app = app
        self.all_appts: list[tuple[str, str, str]] = []  # (appt, prénom, nom)
        self.pres_file: dict | None = None  # {"event", "date", "appts"} from an uploaded file
        self.absents: list[tuple[str, str, str]] = []
        self.event = ""
        self.date = ""

        src = ttk.LabelFrame(parent, text="Sources", padding=10)
        src.pack(fill="x", padx=12, pady=(10, 4))
        src.columnconfigure(1, weight=1)

        ttk.Button(src, text="Charger la liste des membres…",
                   command=self._load_all).grid(row=0, column=0, sticky="w", pady=3)
        self.all_lbl = ttk.Label(src, text="Aucun fichier", foreground="gray")
        self.all_lbl.grid(row=0, column=1, sticky="w", padx=8)

        ttk.Button(src, text="Charger la liste des présences…",
                   command=self._load_pres).grid(row=1, column=0, sticky="w", pady=3)
        self.pres_lbl = ttk.Label(src, text="Présences de l'onglet Présences", foreground="gray")
        self.pres_lbl.grid(row=1, column=1, sticky="w", padx=8)
        ttk.Button(src, text="Utiliser l'onglet Présences",
                   command=self._use_app).grid(row=1, column=2)

        bar = ttk.Frame(parent)
        bar.pack(fill="x", padx=12, pady=(4, 2))
        ttk.Button(bar, text="Absences", command=self._compute, padding=(18, 6)).pack(side="left")
        self.count_lbl = ttk.Label(bar, text="Nombre d'absents : 0",
                                   font=("TkDefaultFont", 10, "bold"))
        self.count_lbl.pack(side="left", padx=14)
        ttk.Button(bar, text="Générer", command=self._generate, padding=(18, 6)).pack(side="right")

        self.header_lbl = ttk.Label(parent, text="", font=("TkDefaultFont", 11, "bold"))
        self.header_lbl.pack(fill="x", padx=12, pady=(8, 2))

        tbl = ttk.LabelFrame(parent, text="Absences", padding=6)
        tbl.pack(fill="both", expand=True, padx=12, pady=(4, 10))
        self.tree = ttk.Treeview(tbl, columns=("Appartement", "Prénom", "Nom"),
                             show="headings", height=12)
        for col, w in (("Appartement", 120), ("Prénom", 200), ("Nom", 200)):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, anchor="center" if col == "Appartement" else "w")
        sb = ttk.Scrollbar(tbl, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

    # ── Sources ───────────────────────────────────────────────────────────────

    @staticmethod
    def _ask_xlsx(title: str) -> str:
        return filedialog.askopenfilename(
            title=title, filetypes=[("Fichier Excel", "*.xlsx *.xlsm")]
        )

    def _load_all(self) -> None:
        path = self._ask_xlsx("Liste des membres")
        if not path:
            return
        try:
            vals = _read_members(openpyxl.load_workbook(path, data_only=True))
        except Exception as exc:
            messagebox.showerror("Erreur", f"Lecture impossible :\n{exc}")
            return
        seen: dict[str, str] = {}
        for v in vals:
            seen.setdefault(_apt_key(v[0]), v[:3])
        self.all_appts = list(seen.values())
        self.all_lbl.config(text=f"{os.path.basename(path)} ({len(self.all_appts)} appartements)",
                            foreground="")

    @staticmethod
    def _find_date(text: str) -> str:
        """First date in text (yyyy-mm-dd or dd-mm-yyyy), as dd/mm/yyyy."""
        m = re.search(r"(?<!\d)(\d{4})[-_./](\d{1,2})[-_./](\d{1,2})(?!\d)", text)
        if m:
            y, mo, d = m.groups()
        else:
            m = re.search(r"(?<!\d)(\d{1,2})[-_./](\d{1,2})[-_./](\d{4})(?!\d)", text)
            if not m:
                return ""
            d, mo, y = m.groups()
        return f"{int(d):02d}/{int(mo):02d}/{y}"

    def _load_pres(self) -> None:
        path = self._ask_xlsx("Liste des présences")
        if not path:
            return
        try:
            ws = openpyxl.load_workbook(path, data_only=True).active
            rows = list(ws.iter_rows(values_only=True))
            hdr = col = type_col = None
            for h, row in enumerate(rows[:15]):
                heads = [str(c).strip().lower() if c is not None else "" for c in row]
                for name in ("appartement", "num", "imm"):
                    if name in heads:
                        hdr, col = h, heads.index(name)
                        type_col = heads.index("type") if "type" in heads else None
                        break
                if hdr is not None:
                    break
            if hdr is None:
                messagebox.showerror("Erreur", "Format de présences non reconnu "
                                     "(colonne « Appartement » introuvable).")
                return
            appts, event = [], ""
            for r in rows[hdr + 1:]:
                v = r[col] if col < len(r) else None
                if v is not None and str(v).strip():
                    appts.append(str(v).strip())
                    if type_col is not None and not event and type_col < len(r) and r[type_col]:
                        event = str(r[type_col]).strip()
            # Event: "Type" column if present, else the Évènement cell (B1);
            # normalised to AGA / AGE / AG / Corvée, also looking in the file name
            b1 = ws["B1"].value
            b1 = "" if b1 is None or hasattr(b1, "strftime") else str(b1).strip()
            if type_col is None:
                event = b1
            event = _event_type(event, b1, os.path.basename(path)) or event
            # Date: from the B1 cell if it has one, else from the file name
            bv = ws["B1"].value
            date = bv.strftime("%d/%m/%Y") if hasattr(bv, "strftime") else self._find_date(str(bv or ""))
            if not date:
                date = self._find_date(os.path.splitext(os.path.basename(path))[0])
        except Exception as exc:
            messagebox.showerror("Erreur", f"Lecture impossible :\n{exc}")
            return
        self.pres_file = {"event": event, "date": date, "appts": appts}
        self.pres_lbl.config(text=f"{os.path.basename(path)} ({len(appts)} présences)",
                             foreground="")

    def _use_app(self) -> None:
        self.pres_file = None
        self.pres_lbl.config(text="Présences de l'onglet Présences", foreground="gray")

    # ── Compute / display ─────────────────────────────────────────────────────

    def _compute(self) -> None:
        if not self.all_appts:
            messagebox.showwarning("Attention", "Chargez d'abord la liste des appartements.")
            return
        if self.pres_file:
            present = self.pres_file["appts"]
            self.event, self.date = self.pres_file["event"], self.pres_file["date"]
        else:
            self.app._flush_debut()
            self.app._flush_fin()
            self.app._refresh_table()
            present = list(self.app.records)
            ev = self.app.event_var.get().strip()
            self.event, self.date = _event_type(ev) or ev, self.app.event_date
            if not present:
                messagebox.showwarning("Attention", "Aucune présence enregistrée.")
                return
        present_set = {_apt_key(a) for a in present}
        self.absents = sorted(
            (a for a in self.all_appts if _apt_key(a[0]) not in present_set), key=lambda a: _natural_key(a[0])
        )
        self.header_lbl.config(text=self._header())
        self.tree.delete(*self.tree.get_children())
        for a in self.absents:
            self.tree.insert("", "end", values=a)
        self.count_lbl.config(text=f"Nombre d'absents : {len(self.absents)}")

    def _header(self) -> str:
        return f"Absences à {self.event} du: {self.date}"

    # ── Excel export ──────────────────────────────────────────────────────────

    def _generate(self) -> None:
        self._compute()
        if not self.header_lbl.cget("text"):
            return
        safe = self.event.replace(" ", "_") or "evenement"
        default = f"absences_{safe}_{self.date.replace('/', '-')}.xlsx"
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Fichier Excel", "*.xlsx")],
            initialfile=default, title="Enregistrer le fichier Excel",
        )
        self.app.root.lift()
        self.app.root.focus_force()
        if not path:
            return

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Absences"
        thin = Side(style="thin")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        center = Alignment(horizontal="center", vertical="center")

        ws.cell(row=1, column=1, value=self._header()).font = Font(bold=True, size=12)
        ws.cell(row=3, column=1, value="Nombre d'absents :").font = Font(bold=True)
        ws.cell(row=3, column=2, value=len(self.absents))
        for col, title in enumerate(("Appartement", "Prénom", "Nom"), start=1):
            h = ws.cell(row=5, column=col, value=title)
            h.font = Font(bold=True, color="FFFFFF", size=11)
            h.fill = PatternFill("solid", fgColor="2E5FA3")
            h.alignment, h.border = center, border
        for i, rec in enumerate(self.absents, start=6):
            for col, v in enumerate(rec, start=1):
                c = ws.cell(row=i, column=col, value=v)
                c.alignment = center if col == 1 else Alignment(horizontal="left", vertical="center")
                c.border = border
        for col, w in zip("ABC", (22, 26, 26)):
            ws.column_dimensions[col].width = w

        wb.save(path)
        messagebox.showinfo("Succès", f"Fichier généré :\n{path}")
        self.app.root.lift()
        self.app.root.focus_force()


# ── QR codes ──────────────────────────────────────────────────────────────────

def _member_stem(m: tuple[str, str, str, str]) -> str:
    """A101_Prenom_Nom — used as file name and as the QR code content."""
    parts = [p.strip().replace(" ", "_") for p in m[:3] if p.strip()]
    return re.sub(r'[\\/:*?"<>|]', "", "_".join(parts))


def _font(size: int):
    for name in ("Helvetica.ttc", "Arial.ttf", "arial.ttf", "DejaVuSans.ttf"):
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            continue
    return ImageFont.load_default(size)


def _make_qr(text: str) -> Image.Image:
    qr = qrcode.QRCode(
        error_correction=qrcode.constants.ERROR_CORRECT_H,  # leaves room for the logo
        box_size=12,
        border=4,
    )
    qr.add_data(text)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white").convert("RGBA")
    try:
        logo = Image.open(_resource_path("logo.png")).convert("RGBA")
    except OSError:
        return img.convert("RGB")
    side = img.width * 22 // 100
    logo.thumbnail((side, side), Image.LANCZOS)
    pad = img.width // 70
    plate = Image.new("RGBA", (logo.width + 2 * pad, logo.height + 2 * pad), "white")
    plate.paste(logo, (pad, pad), logo)
    img.paste(plate, ((img.width - plate.width) // 2, (img.height - plate.height) // 2), plate)
    return img.convert("RGB")


def _qr_png_bytes(img: Image.Image) -> bytes:
    buf = io.BytesIO()
    img.save(buf, "PNG")
    return buf.getvalue()


def _label_pages(members: list[tuple[str, str, str, str]]) -> list[Image.Image]:
    """Letter-size pages (300 dpi), 3 x 4 labelled QR codes each."""
    W, H, cols, rows = 2550, 3300, 3, 4
    cw, ch = W // cols, H // rows
    pages = []
    for start in range(0, len(members), cols * rows):
        page = Image.new("RGB", (W, H), "white")
        draw = ImageDraw.Draw(page)
        for i, m in enumerate(members[start:start + cols * rows]):
            x, y = (i % cols) * cw, (i // cols) * ch
            qr = _make_qr(_member_stem(m)).resize((cw - 140, cw - 140), Image.LANCZOS)
            page.paste(qr, (x + 70, y + 30))
            caption = f"{m[0]}  {m[1]} {m[2]}".strip()
            font = _font(44)
            while draw.textlength(caption, font=font) > cw - 40 and font.size > 20:
                font = _font(font.size - 4)
            tw = draw.textlength(caption, font=font)
            draw.text((x + (cw - tw) / 2, y + cw - 100), caption, fill="black", font=font)
        pages.append(page)
    return pages


class QrTab:
    """Generate, e-mail and print one QR code per member."""

    ALL = "(Tous les membres)"

    def __init__(self, parent: ttk.Frame, app: "AttendanceApp") -> None:
        self.app = app
        self.members: list[tuple[str, str, str, str]] = []
        self._preview = None
        self._queue: queue.Queue = queue.Queue()

        src = ttk.LabelFrame(parent, text="Membres", padding=10)
        src.pack(fill="x", padx=12, pady=(10, 4))
        src.columnconfigure(1, weight=1)
        ttk.Button(src, text="Charger la liste des membres…",
                   command=self._load).grid(row=0, column=0, sticky="w", pady=3)
        self.file_lbl = ttk.Label(src, text="Aucun fichier", foreground="gray")
        self.file_lbl.grid(row=0, column=1, sticky="w", padx=8)

        ttk.Label(src, text="Membre :").grid(row=1, column=0, sticky="w", pady=3)
        self.member_var = tk.StringVar(value=self.ALL)
        self.combo = ttk.Combobox(src, textvariable=self.member_var, state="readonly",
                                  values=[self.ALL], width=48)
        self.combo.grid(row=1, column=1, sticky="w", padx=8)
        self.combo.bind("<<ComboboxSelected>>", lambda e: self._show_preview())

        bar = ttk.Frame(parent)
        bar.pack(fill="x", padx=12, pady=(4, 2))
        ttk.Button(bar, text="Générer tous", command=self._generate_all,
                   padding=(14, 6)).pack(side="left", padx=(0, 8))
        ttk.Button(bar, text="Générer pour membre", command=self._generate_one,
                   padding=(14, 6)).pack(side="left", padx=(0, 8))
        ttk.Button(bar, text="Imprimer", command=self._print,
                   padding=(14, 6)).pack(side="left", padx=(0, 8))
        ttk.Button(bar, text="Envoyer", command=self._send,
                   padding=(14, 6)).pack(side="left")

        self.preview_lbl = ttk.Label(parent, anchor="center")
        self.preview_lbl.pack(pady=6)
        self.preview_txt = ttk.Label(parent, text="", foreground="gray")
        self.preview_txt.pack()

        # E-mail settings (password is kept in memory only)
        mail = ttk.LabelFrame(parent, text="Envoi par courriel (SMTP)", padding=10)
        mail.pack(fill="x", padx=12, pady=(8, 4))
        mail.columnconfigure(1, weight=1)
        mail.columnconfigure(3, weight=1)
        self.smtp_host = tk.StringVar(value="smtp.gmail.com")
        self.smtp_port = tk.StringVar(value="587")
        self.smtp_user = tk.StringVar()
        self.smtp_pass = tk.StringVar()
        self.mail_subject = tk.StringVar(value="Votre code QR")
        ttk.Label(mail, text="Serveur :").grid(row=0, column=0, sticky="w", pady=2)
        ttk.Entry(mail, textvariable=self.smtp_host).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Label(mail, text="Port :").grid(row=0, column=2, sticky="e")
        ttk.Entry(mail, textvariable=self.smtp_port, width=8).grid(row=0, column=3, sticky="w", padx=6)
        ttk.Label(mail, text="Courriel :").grid(row=1, column=0, sticky="w", pady=2)
        ttk.Entry(mail, textvariable=self.smtp_user).grid(row=1, column=1, sticky="ew", padx=6)
        ttk.Label(mail, text="Mot de passe :").grid(row=1, column=2, sticky="e")
        ttk.Entry(mail, textvariable=self.smtp_pass, show="•").grid(row=1, column=3, sticky="ew", padx=6)
        ttk.Label(mail, text="Objet :").grid(row=2, column=0, sticky="w", pady=2)
        ttk.Entry(mail, textvariable=self.mail_subject).grid(row=2, column=1, columnspan=3, sticky="ew", padx=6)
        ttk.Label(mail, text="Message :").grid(row=3, column=0, sticky="nw", pady=2)
        self.mail_body = tk.Text(mail, height=4, width=40, wrap="word")
        self.mail_body.insert(
            "1.0",
            "Bonjour {prenom},\n\nVeuillez trouver ci-joint votre code QR "
            "pour l'appartement {appartement}.\n\nCordialement,\nLes Voisins de Viau-Robert",
        )
        self.mail_body.grid(row=3, column=1, columnspan=3, sticky="ew", padx=6)
        self.status = ttk.Label(parent, text="", foreground="gray")
        self.status.pack(fill="x", padx=12, pady=(2, 8))

    # ── Members ───────────────────────────────────────────────────────────────

    def _load(self) -> None:
        path = filedialog.askopenfilename(
            title="Liste des membres", filetypes=[("Fichier Excel", "*.xlsx *.xlsm")]
        )
        if not path:
            return
        try:
            vals = _read_members(openpyxl.load_workbook(path, data_only=True))
        except Exception as exc:
            messagebox.showerror("Erreur", f"Lecture impossible :\n{exc}")
            return
        seen: dict[str, tuple] = {}
        for v in vals:
            seen.setdefault(v[0].lower(), v)
        self.members = sorted(seen.values(), key=lambda m: _natural_key(m[0]))
        self.combo.config(values=[self.ALL] + [self._label(m) for m in self.members])
        self.member_var.set(self.ALL)
        self.file_lbl.config(text=f"{os.path.basename(path)} ({len(self.members)} membres)",
                             foreground="")
        self._show_preview()

    @staticmethod
    def _label(m: tuple) -> str:
        return f"{m[0]} — {m[1]} {m[2]}".strip(" —")

    def _selected(self) -> list[tuple[str, str, str, str]] | None:
        """Chosen member(s); None (after a warning) if there is nothing to act on."""
        if not self.members:
            messagebox.showwarning("Attention", "Chargez d'abord la liste des membres.")
            return None
        idx = self.combo.current()
        return self.members if idx <= 0 else [self.members[idx - 1]]

    def _show_preview(self) -> None:
        idx = self.combo.current()
        if idx <= 0:
            self._preview = None
            self.preview_lbl.config(image="")
            self.preview_txt.config(text="")
            return
        m = self.members[idx - 1]
        img = _make_qr(_member_stem(m)).resize((220, 220), Image.LANCZOS)
        self._preview = tk.PhotoImage(data=base64.b64encode(_qr_png_bytes(img)))
        self.preview_lbl.config(image=self._preview)
        self.preview_txt.config(text=_member_stem(m))

    # ── Generate ──────────────────────────────────────────────────────────────

    def _save(self, members: list[tuple]) -> None:
        folder = filedialog.askdirectory(title="Dossier de destination des codes QR")
        self.app.root.lift()
        self.app.root.focus_force()
        if not folder:
            return
        try:
            for m in members:
                _make_qr(_member_stem(m)).save(os.path.join(folder, f"{_member_stem(m)}.png"))
        except OSError as exc:
            messagebox.showerror("Erreur", f"Écriture impossible :\n{exc}")
            return
        messagebox.showinfo("Succès", f"{len(members)} code(s) QR généré(s) dans :\n{folder}")

    def _generate_all(self) -> None:
        if self._selected() is not None:
            self._save(self.members)

    def _generate_one(self) -> None:
        if self._selected() is None:
            return
        idx = self.combo.current()
        if idx <= 0:
            messagebox.showwarning("Attention", "Choisissez un membre dans la liste.")
            return
        self._show_preview()
        self._save([self.members[idx - 1]])

    # ── Print ─────────────────────────────────────────────────────────────────

    def _print(self) -> None:
        sel = self._selected()
        if sel is None:
            return
        pages = _label_pages(sel)
        if len(sel) > 1 and not messagebox.askyesno(
            "Imprimer", f"Imprimer {len(sel)} codes QR ({len(pages)} page(s)) ?"
        ):
            return
        path = os.path.join(tempfile.mkdtemp(prefix="qr_"), "codes_qr.pdf")
        pages[0].save(path, "PDF", resolution=300, save_all=True, append_images=pages[1:])
        try:
            if sys.platform.startswith("win"):
                os.startfile(path, "print")  # type: ignore[attr-defined]
            else:
                subprocess.run(["lpr", path], check=True)
        except Exception as exc:
            messagebox.showerror("Erreur", f"Impression impossible :\n{exc}")
            return
        self.status.config(text=f"Envoyé à l'imprimante : {len(sel)} code(s) QR.")

    # ── E-mail ────────────────────────────────────────────────────────────────

    def _send(self) -> None:
        sel = self._selected()
        if sel is None:
            return
        host, user, pw = self.smtp_host.get().strip(), self.smtp_user.get().strip(), self.smtp_pass.get()
        try:
            port = int(self.smtp_port.get())
        except ValueError:
            messagebox.showerror("Erreur", "Port invalide.")
            return
        if not (host and user and pw):
            messagebox.showwarning("Attention", "Renseignez le serveur, le courriel et le mot de passe.")
            return
        targets = [m for m in sel if "@" in m[3]]
        skipped = len(sel) - len(targets)
        if not targets:
            messagebox.showwarning("Attention", "Aucune adresse courriel pour la sélection.")
            return
        msg = f"Envoyer {len(targets)} courriel(s) depuis {user} ?"
        if skipped:
            msg += f"\n({skipped} membre(s) sans adresse seront ignorés.)"
        if not messagebox.askyesno("Envoyer", msg):
            return
        subject, body = self.mail_subject.get(), self.mail_body.get("1.0", "end").strip()
        threading.Thread(
            target=self._send_worker, args=(host, port, user, pw, subject, body, targets, skipped),
            daemon=True,
        ).start()
        self.status.config(text="Envoi en cours…")
        self.app.root.after(100, self._poll)

    def _send_worker(self, host, port, user, pw, subject, body, targets, skipped) -> None:
        sent, failed = 0, []
        try:
            smtp = (smtplib.SMTP_SSL(host, port, timeout=30) if port == 465
                    else smtplib.SMTP(host, port, timeout=30))
            with smtp:
                if port != 465 and smtp.has_extn("starttls"):
                    smtp.starttls(context=ssl.create_default_context())
                if smtp.has_extn("auth"):
                    smtp.login(user, pw)
                for i, m in enumerate(targets, 1):
                    mail = EmailMessage()
                    mail["From"], mail["To"], mail["Subject"] = user, m[3], subject
                    mail.set_content(body.replace("{prenom}", m[1]).replace("{nom}", m[2])
                                     .replace("{appartement}", m[0]))
                    mail.add_attachment(_qr_png_bytes(_make_qr(_member_stem(m))),
                                        maintype="image", subtype="png",
                                        filename=f"{_member_stem(m)}.png")
                    try:
                        smtp.send_message(mail)
                        sent += 1
                    except smtplib.SMTPException as exc:
                        failed.append(f"{m[0]} ({m[3]}): {exc}")
                    self._queue.put(("progress", f"Envoi… {i}/{len(targets)}"))
        except Exception as exc:
            self._queue.put(("done", False, f"Échec de l'envoi : {exc}"))
            return
        text = f"{sent} courriel(s) envoyé(s)."
        if skipped:
            text += f" {skipped} sans adresse."
        if failed:
            text += f"\n{len(failed)} échec(s) :\n" + "\n".join(failed[:10])
        self._queue.put(("done", not failed, text))

    def _poll(self) -> None:
        try:
            while True:
                item = self._queue.get_nowait()
                if item[0] == "progress":
                    self.status.config(text=item[1])
                else:
                    self.status.config(text=item[2].splitlines()[0])
                    (messagebox.showinfo if item[1] else messagebox.showwarning)(
                        "Envoi" if item[1] else "Attention", item[2])
                    self.app.root.lift()
                    return
        except queue.Empty:
            pass
        self.app.root.after(100, self._poll)


def main() -> None:
    root = tk.Tk()
    AttendanceApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
