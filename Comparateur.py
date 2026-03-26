import streamlit as st
import pandas as pd
import pdfplumber
import unicodedata
import re
from rapidfuzz import fuzz

st.title("Comparateur Excel / PDF")

# ---------------------------
# 1️⃣ Charger le fichier Excel
# ---------------------------
excel_file = st.file_uploader("Choisir un fichier Excel", type=["xlsx", "xls"])

if excel_file:
    # ⚠️ on ignore la première ligne
    df_excel = pd.read_excel(excel_file, header=1)

    st.write("Aperçu Excel :")
    st.dataframe(df_excel)

    colonnes_excel = st.multiselect(
        "Choisir les colonnes Excel à comparer",
        df_excel.columns
    )

# ---------------------------
# 2️⃣ Charger le fichier PDF
# ---------------------------
pdf_file = st.file_uploader("Choisir un fichier PDF", type=["pdf"])

if pdf_file:
    all_rows = []
    header = None

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            table = page.extract_table()

            if table:
                current_header = [str(col).strip() for col in table[0]]

                # garder seulement le premier header
                if header is None:
                    header = current_header

                # ajouter les lignes
                for row in table[1:]:
                    if row and len(row) == len(header):
                        all_rows.append(row)

    if all_rows and header:
        df_pdf = pd.DataFrame(all_rows, columns=header)

        # nettoyage
        df_pdf = df_pdf.dropna(how="all")
        df_pdf = df_pdf.fillna("")

        # supprimer faux headers
        df_pdf = df_pdf[
            ~(df_pdf.apply(lambda row: list(row) == header, axis=1))
        ]

        st.write("Aperçu PDF (fusion propre) :")
        st.dataframe(df_pdf)

        colonnes_pdf = st.multiselect(
            "Choisir les colonnes PDF à comparer",
            df_pdf.columns
        )
    else:
        st.error("Aucun tableau détecté dans le PDF")

# ---------------------------
# Nettoyage texte (robuste)
# ---------------------------
def nettoyer_texte(texte):
    if pd.isna(texte):
        return ""

    texte = str(texte).lower()

    # enlever accents
    texte = unicodedata.normalize('NFD', texte)
    texte = ''.join(c for c in texte if unicodedata.category(c) != 'Mn')

    # enlever caractères spéciaux
    texte = re.sub(r'[^a-z0-9\s]', ' ', texte)

    # enlever espaces multiples
    texte = re.sub(r'\s+', ' ', texte).strip()

    return texte


# ---------------------------
# Création clé + filtrage intelligent
# ---------------------------
def creer_cle(df, colonnes):

    # Combiner + nettoyer
    series = df[colonnes].astype(str).apply(
        lambda row: nettoyer_texte(" ".join(row)),
        axis=1
    )

    # mots parasites (version robuste)
    mots_a_ignorer = ["agence", "total", "sous total", "total general"]

    def est_valide(val):
        for mot in mots_a_ignorer:
            if mot in val:   # 🔥 plus robuste que regex
                return False
        return True

    series_filtre = series[series.apply(est_valide)]

    return series_filtre


# ---------------------------
# Fuzzy matching intelligent
# ---------------------------
def comparer_listes(liste1, liste2, seuil=85):

    correspondances = []
    non_trouves = []

    for val1 in liste1:
        match_trouve = False

        for val2 in liste2:
            score = fuzz.token_sort_ratio(val1, val2)

            if score >= seuil:
                correspondances.append({
                    "excel": val1,
                    "pdf": val2,
                    "score": score
                })
                match_trouve = True
                break

        if not match_trouve:
            non_trouves.append(val1)

    return correspondances, non_trouves

if st.button("Comparer"):

    if excel_file and pdf_file and colonnes_excel and colonnes_pdf:

        liste_excel = creer_cle(df_excel, colonnes_excel).tolist()
        liste_pdf = creer_cle(df_pdf, colonnes_pdf).tolist()

        seuil = 85

        correspondances = []
        only_excel = []
        only_pdf = []

        pdf_utilises = set()  # 🔥 évite les doublons

        # Excel → PDF (avec meilleur match)
        for val_excel in liste_excel:

            meilleur_score = 0
            meilleur_match = None

            for val_pdf in liste_pdf:

                if val_pdf in pdf_utilises:
                    continue  # déjà utilisé

                score = fuzz.token_sort_ratio(val_excel, val_pdf)

                if score > meilleur_score:
                    meilleur_score = score
                    meilleur_match = val_pdf

            if meilleur_score >= seuil:
                correspondances.append((val_excel, meilleur_match, meilleur_score))
                pdf_utilises.add(meilleur_match)
            else:
                only_excel.append(val_excel)

        # PDF non matchés
        for val_pdf in liste_pdf:
            if val_pdf not in pdf_utilises:
                only_pdf.append(val_pdf)

        # Affichage propre
        st.subheader("Correspondances trouvées")
        st.dataframe(pd.DataFrame(correspondances, columns=["Excel", "PDF", "Score"]))

        st.subheader("Présents dans Excel mais absents dans PDF")
        st.write(only_excel)

        st.subheader("Présents dans PDF mais absents dans Excel")
        st.write(only_pdf)

    else:
        st.warning("Charge les fichiers ET sélectionne les colonnes")