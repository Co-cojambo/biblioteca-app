#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Căutare în BIBLIOTECA – aplicație cu interfață grafică (tkinter).

Scrii numele unui autor sau titlul unei cărți și vezi imediat toate informațiile
din fișierul Excel: autor, titlu, volume, exemplare, editură, colecție.

Cerințe:   pip install openpyxl
Pornire:   python biblioteca.py
(Fișierul BIBLIOTECA_1_.xlsx trebuie să fie în același folder cu scriptul;
 dacă nu e găsit, ți se cere să-l alegi.)
"""

import os
import sys
import unicodedata

from openpyxl import load_workbook

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
except ImportError:  # permite testarea logicii fără interfață
    tk = None

FISIER_IMPLICIT = "BIBLIOTECA_1_.xlsx"
ANONIM = "(autor necunoscut / anonim)"


# ----------------------------------------------------------------------------
# Logica: citirea datelor și căutarea
# ----------------------------------------------------------------------------
def normalizeaza(text):
    """Litere mici, fără diacritice, spații curățate: 'Mioriţa' -> 'miorita'."""
    text = unicodedata.normalize("NFKD", str(text or ""))
    text = "".join(c for c in text if not unicodedata.combining(c))
    return " ".join(text.lower().split())


def _curat(valoare):
    if valoare is None:
        return ""
    return str(valoare).strip()


def _numar(valoare):
    if valoare is None or _curat(valoare) == "":
        return ""
    try:
        return str(int(float(valoare)))
    except (TypeError, ValueError):
        return _curat(valoare)


def incarca_carti(cale):
    """Citește fișierul Excel și întoarce lista de cărți (dicționare)."""
    wb = load_workbook(cale, read_only=True, data_only=True)
    ws = wb.active
    carti = []
    for nr_rand, rand in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        rand = list(rand) + [None] * 7
        _nr, autor, titlu, volume, exemplare, editura, colectia = rand[:7]

        titlu = _curat(titlu)
        if not titlu:          # sare peste rândurile goale / de total
            continue

        autor = _curat(autor)
        if autor in ("", "//"):
            autor = ANONIM

        carte = {
            "autor": autor,
            "titlu": titlu,
            "volume": _numar(volume),
            "exemplare": _numar(exemplare),
            "editura": _curat(editura),
            "colectia": _curat(colectia),
            "rand": nr_rand,
        }
        carte["_autor_n"] = normalizeaza(autor)
        carte["_titlu_n"] = normalizeaza(titlu)
        carti.append(carte)
    wb.close()
    return carti


def cauta(carti, interogare, mod="tot"):
    """
    Întoarce cărțile care conțin TOATE cuvintele scrise (în orice ordine),
    ignorând diacriticele și majusculele.
    mod: 'tot' (autor + titlu), 'autor' sau 'titlu'.
    """
    cuvinte = normalizeaza(interogare).split()
    if not cuvinte:
        rezultate = list(carti)
    else:
        rezultate = []
        for c in carti:
            if mod == "autor":
                text = c["_autor_n"]
            elif mod == "titlu":
                text = c["_titlu_n"]
            else:
                text = c["_autor_n"] + " " + c["_titlu_n"]
            if all(cuv in text for cuv in cuvinte):
                rezultate.append(c)
    rezultate.sort(key=lambda c: (c["_autor_n"], c["_titlu_n"]))
    return rezultate


# ----------------------------------------------------------------------------
# Interfața grafică
# ----------------------------------------------------------------------------
if tk is not None:

    class Aplicatie(ttk.Frame):
        COLOANE = (
            ("autor", "Autor", 220),
            ("titlu", "Titlu", 330),
            ("editura", "Editura", 110),
            ("colectia", "Colecția", 170),
        )

        def __init__(self, master):
            super().__init__(master, padding=10)
            self.pack(fill="both", expand=True)
            self.carti = []
            self.afisate = []
            self.cale = None
            self._ordine_inversata = {}
            self._construieste()
            self._incarca_initial()

        # ---------- construirea ferestrei ----------
        def _construieste(self):
            # Rândul de căutare
            sus = ttk.Frame(self)
            sus.pack(fill="x")

            ttk.Label(sus, text="Caută autor sau carte:",
                      font=("TkDefaultFont", 11, "bold")).pack(side="left")

            self.var_cautare = tk.StringVar()
            self.var_cautare.trace_add("write", lambda *_: self.actualizeaza())
            self.camp = ttk.Entry(sus, textvariable=self.var_cautare,
                                  font=("TkDefaultFont", 12))
            self.camp.pack(side="left", fill="x", expand=True, padx=8)
            self.camp.focus_set()

            ttk.Button(sus, text="Șterge", command=self.sterge).pack(side="left")

            # Opțiuni
            optiuni = ttk.Frame(self)
            optiuni.pack(fill="x", pady=(6, 6))
            ttk.Label(optiuni, text="Caută în:").pack(side="left")
            self.var_mod = tk.StringVar(value="tot")
            for text, val in (("Autor și titlu", "tot"), ("Doar autor", "autor"),
                              ("Doar titlu", "titlu")):
                ttk.Radiobutton(optiuni, text=text, value=val, variable=self.var_mod,
                                command=self.actualizeaza).pack(side="left", padx=6)
            ttk.Button(optiuni, text="Alege alt fișier Excel…",
                       command=self.alege_fisier).pack(side="right")

            # Panou împărțit: tabel sus, detalii jos
            panou = ttk.PanedWindow(self, orient="vertical")
            panou.pack(fill="both", expand=True)

            cadru_tabel = ttk.Frame(panou)
            panou.add(cadru_tabel, weight=3)

            self.tabel = ttk.Treeview(cadru_tabel, columns=[c[0] for c in self.COLOANE],
                                      show="headings", selectmode="browse")
            for cheie, titlu, latime in self.COLOANE:
                self.tabel.heading(cheie, text=titlu,
                                   command=lambda k=cheie: self.sorteaza(k))
                self.tabel.column(cheie, width=latime, anchor="w")
            bara = ttk.Scrollbar(cadru_tabel, orient="vertical", command=self.tabel.yview)
            self.tabel.configure(yscrollcommand=bara.set)
            self.tabel.pack(side="left", fill="both", expand=True)
            bara.pack(side="right", fill="y")
            self.tabel.bind("<<TreeviewSelect>>", self.arata_detalii)

            cadru_det = ttk.LabelFrame(panou, text="Detalii", padding=6)
            panou.add(cadru_det, weight=2)

            self.detalii = tk.Text(cadru_det, height=11, wrap="word", state="disabled",
                                   font=("TkDefaultFont", 11), relief="flat",
                                   background=self.winfo_toplevel().cget("background"))
            bara2 = ttk.Scrollbar(cadru_det, orient="vertical", command=self.detalii.yview)
            self.detalii.configure(yscrollcommand=bara2.set)
            self.detalii.pack(side="left", fill="both", expand=True)
            bara2.pack(side="right", fill="y")
            self.detalii.tag_configure("eticheta", font=("TkDefaultFont", 11, "bold"))
            self.detalii.tag_configure("titlu", font=("TkDefaultFont", 14, "bold"))
            self.detalii.tag_configure("sectiune", font=("TkDefaultFont", 11, "bold"),
                                       spacing1=8)

            # Bara de stare
            self.var_stare = tk.StringVar()
            ttk.Label(self, textvariable=self.var_stare, anchor="w").pack(fill="x", pady=(6, 0))

            self.winfo_toplevel().bind("<Escape>", lambda e: self.sterge())

        # ---------- încărcarea fișierului ----------
        def _incarca_initial(self):
            foldere = [os.path.dirname(os.path.abspath(__file__)), os.getcwd()]
            for folder in foldere:
                cale = os.path.join(folder, FISIER_IMPLICIT)
                if os.path.exists(cale):
                    if self.incarca(cale):
                        return
            self.alege_fisier()

        def incarca(self, cale):
            try:
                carti = incarca_carti(cale)
            except Exception as e:  # fișier corupt, deschis în Excel etc.
                messagebox.showerror("Eroare", f"Nu am putut citi fișierul:\n{cale}\n\n{e}")
                return False
            if not carti:
                messagebox.showwarning("Fișier gol", "Nu am găsit nicio carte în fișier.")
                return False
            self.carti = carti
            self.cale = cale
            self.winfo_toplevel().title(f"Biblioteca – {os.path.basename(cale)}")
            self.actualizeaza()
            return True

        def alege_fisier(self):
            cale = filedialog.askopenfilename(
                title="Alege fișierul Excel al bibliotecii",
                filetypes=[("Fișiere Excel", "*.xlsx *.xlsm"), ("Toate fișierele", "*.*")])
            if cale:
                self.incarca(cale)
            elif not self.carti:
                self.var_stare.set("Niciun fișier încărcat. Apasă „Alege alt fișier Excel…”.")

        # ---------- căutare și afișare ----------
        def sterge(self):
            self.var_cautare.set("")
            self.camp.focus_set()

        def actualizeaza(self):
            self.afisate = cauta(self.carti, self.var_cautare.get(), self.var_mod.get())
            self.tabel.delete(*self.tabel.get_children())
            for i, c in enumerate(self.afisate):
                self.tabel.insert("", "end", iid=str(i),
                                  values=(c["autor"], c["titlu"], c["editura"], c["colectia"]))
            total_ex = sum(int(c["exemplare"]) for c in self.afisate if c["exemplare"].isdigit())
            self.var_stare.set(
                f"{len(self.afisate)} rezultate din {len(self.carti)} cărți "
                f"({total_ex} exemplare în rezultate)")
            self._scrie_detalii([])
            # dacă există un singur rezultat, îl selectăm automat
            if len(self.afisate) == 1:
                self.tabel.selection_set("0")

        def sorteaza(self, cheie):
            inversat = self._ordine_inversata.get(cheie, False)
            self.afisate.sort(key=lambda c: normalizeaza(c[cheie]), reverse=inversat)
            self._ordine_inversata[cheie] = not inversat
            self.tabel.delete(*self.tabel.get_children())
            for i, c in enumerate(self.afisate):
                self.tabel.insert("", "end", iid=str(i),
                                  values=(c["autor"], c["titlu"], c["editura"], c["colectia"]))

        def arata_detalii(self, _event=None):
            sel = self.tabel.selection()
            if not sel:
                return
            c = self.afisate[int(sel[0])]
            linii = [
                ("titlu", c["titlu"] + "\n"),
                ("eticheta", "Autor: "), (None, c["autor"] + "\n"),
                ("eticheta", "Titlu: "), (None, c["titlu"] + "\n"),
                ("eticheta", "Volume: "), (None, (c["volume"] or "—") + "\n"),
                ("eticheta", "Exemplare: "), (None, (c["exemplare"] or "—") + "\n"),
                ("eticheta", "Editura: "), (None, (c["editura"] or "—") + "\n"),
                ("eticheta", "Colecția: "), (None, (c["colectia"] or "—") + "\n"),
                ("eticheta", "Rând în Excel: "), (None, str(c["rand"]) + "\n"),
            ]
            if c["autor"] != ANONIM:
                altele = [x for x in self.carti
                          if x["_autor_n"] == c["_autor_n"] and x is not c]
                if altele:
                    linii.append(("sectiune", f"\nAlte cărți de {c['autor']} ({len(altele)}):\n"))
                    for x in sorted(altele, key=lambda x: x["_titlu_n"]):
                        extra = f"  [{x['editura']}]" if x["editura"] else ""
                        linii.append((None, f"  • {x['titlu']}{extra}\n"))
            self._scrie_detalii(linii)

        def _scrie_detalii(self, linii):
            self.detalii.configure(state="normal")
            self.detalii.delete("1.0", "end")
            for tag, text in linii:
                self.detalii.insert("end", text, tag or ())
            self.detalii.configure(state="disabled")


    def main():
        radacina = tk.Tk()
        radacina.title("Biblioteca")
        radacina.geometry("920x700")
        radacina.minsize(700, 500)
        Aplicatie(radacina)
        radacina.mainloop()


else:
    def main():
        print("Eroare: tkinter nu este instalat.\n"
              "Pe Linux: sudo apt install python3-tk\n"
              "Pe Windows/macOS vine cu instalarea standard de Python.")
        sys.exit(1)


if __name__ == "__main__":
    main()
