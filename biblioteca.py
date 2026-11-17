# -*- coding: utf-8 -*-
"""
Căutare în BIBLIOTECA – Aplicație Web Streamlit
"""

import os
import unicodedata
import pandas as pd
from openpyxl import load_workbook
import streamlit as st

# Numele fișierului Excel urcat pe GitHub
FISIER_IMPLICIT = "BIBLIOTECA_CO.xlsx"
ANONIM = "(autor necunoscut / anonim)"

st.set_page_config(page_title="Biblioteca", page_icon="📚", layout="wide")

def normalizeaza(text):
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

@st.cache_data
def incarca_carti(cale):
    if not os.path.exists(cale):
        return []
    wb = load_workbook(cale, read_only=True, data_only=True)
    ws = wb.active
    carti = []
    for nr_rand, rand in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        rand = list(rand) + [None] * 7
        _nr, autor, titlu, volume, exemplare, editura, colectia = rand[:7]

        titlu = _curat(titlu)
        if not titlu:
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

def main():
    st.title("📚 Căutare în Biblioteca")
    
    # Încărcare date
    if not os.path.exists(FISIER_IMPLICIT):
        st.error(f"Nu am găsit fișierul Excel '{FISIER_IMPLICIT}' în repository-ul tău!")
        return

    carti = incarca_carti(FISIER_IMPLICIT)
    
    if not carti:
        st.warning("Fișierul Excel este gol sau nu conține date valide.")
        return

    # Opțiuni de căutare
    col1, col2 = st.columns([3, 1])
    with col1:
        interogare = st.text_input("Caută autor sau carte:", placeholder="Scrie aici...")
    with col2:
        mod = st.selectbox("Caută în:", ["Tot (autor și titlu)", "Doar autor", "Doar titlu"])

    # Mapare mod căutare
    mod_map = {"Tot (autor și titlu)": "tot", "Doar autor": "autor", "Doar titlu": "titlu"}
    mod_ales = mod_map[mod]

    # Logica de căutare
    cuvinte = normalizeaza(interogare).split()
    if not cuvinte:
        rezultate = list(carti)
    else:
        rezultate = []
        for c in carti:
            if mod_ales == "autor":
                text = c["_autor_n"]
            elif mod_ales == "titlu":
                text = c["_titlu_n"]
            else:
                text = c["_autor_n"] + " " + c["_titlu_n"]
            if all(cuv in text for cuv in cuvinte):
                rezultate.append(c)
    rezultate.sort(key=lambda c: (c["_autor_n"], c["_titlu_n"]))

    # Statistici
    total_ex = sum(int(c["exemplare"]) for c in rezultate if c["exemplare"].isdigit())
    st.info(f"📊 S-au găsit **{len(rezultate)}** rezultate din totalul de **{len(carti)}** cărți ({total_ex} exemplare în rezultate).")

    # Pregătire tabel pentru afișare
    if rezultate:
        df_afisaj = pd.DataFrame([{
            "Autor": c["autor"],
            "Titlu": c["titlu"],
            "Volume": c["volume"],
            "Exemplare": c["exemplare"],
            "Editura": c["editura"],
            "Colecția": c["colectia"]
        } for c in rezultate])
        
        st.dataframe(df_afisaj, use_container_width=True, hide_index=True)
    else:
        st.warning("Nicio carte găsită conform căutării.")

if __name__ == "__main__":
    main()