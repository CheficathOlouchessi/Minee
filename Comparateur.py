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
    # Lire toutes les feuilles
    all_sheets = pd.read_excel(excel_file, sheet_name=None, header=1)

    # Fusionner toutes les feuilles
    df_excel = pd.concat(all_sheets.values(), ignore_index=True)
    df_excel.columns = df_excel.columns.astype(str).str.strip()

    # 🔥 Supprimer les lignes totalement vides
    df_excel = df_excel.dropna(how="all")

    # 🔥 Supprimer les doublons sur TOUTES les colonnes
    df_excel = df_excel.drop_duplicates()

    st.subheader("Aperçu Excel")
    st.dataframe(df_excel.head())

    # 🔥 Nombre de lignes final
    st.write(f"Nombre de lignes Excel (après nettoyage) : {len(df_excel)}")

    colonnes_excel = st.multiselect(
        "Colonnes Excel à comparer",
        df_excel.columns
    )

# ---------------------------
# 2️⃣ Charger le fichier PDF
# ---------------------------
pdf_file = st.file_uploader("Choisir un fichier PDF", type=["pdf"])

if pdf_file:
    all_rows = []
    header = None

    def corriger_ligne(row):
        row = [str(x) if x is not None else "" for x in row]

        # 🔥 détecter ligne décalée (colonne vide + chiffre après)
        if len(row) >= 4:
            if row[1] == "" and row[2].strip().isdigit():
                # décalage vers la gauche
                row[1] = row[2]
                row[2] = row[3]
                row[3] = row[4] if len(row) > 4 else ""
                if len(row) > 4:
                    row[4] = ""

        return row

    def est_ligne_total(row):
        texte = " ".join([str(x).lower() for x in row])
        return any(m in texte for m in ["total", "sous total", "effectif"])

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()

            if not tables:
                continue

            for table in tables:

                if not table or len(table) < 2:
                    continue

                current_header = [str(col).strip() for col in table[0]]

                # garder premier header
                if header is None:
                    header = current_header

                for row in table[1:]:

                    if not row:
                        continue

                    # 🔥 harmoniser taille
                    if len(row) < len(header):
                        row = row + [""] * (len(header) - len(row))
                    elif len(row) > len(header):
                        row = row[:len(header)]

                    # 🔥 corriger décalage
                    row = corriger_ligne(row)

                    all_rows.append(row)

    if all_rows and header:

        df_pdf = pd.DataFrame(all_rows, columns=header)

        # nettoyage
        df_pdf = df_pdf.dropna(how="all")
        df_pdf = df_pdf.fillna("")

        # 🔥 supprimer lignes TOTAL / EFFECTIF
        df_pdf = df_pdf[
            ~df_pdf.apply(est_ligne_total, axis=1)
        ]

        # supprimer faux headers
        df_pdf = df_pdf[
            ~(df_pdf.apply(lambda row: list(row) == header, axis=1))
        ]

        # 🔥 supprimer doublons
        df_pdf = df_pdf.drop_duplicates()

        st.write(f"Nombre de lignes PDF (après nettoyage) : {len(df_pdf)}")
        st.write("Aperçu PDF (corrigé) :")
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
    series = df[colonnes].fillna("").astype(str).apply(
        lambda row: nettoyer_texte(" ".join(row.values)),
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
def dedoublonner_fuzzy(liste, seuil=90):
    resultat = []

    for val in liste:
        deja_existe = False

        for existant in resultat:
            score = fuzz.token_set_ratio(val, existant)

            if score >= seuil:
                deja_existe = True
                break

        if not deja_existe:
            resultat.append(val)

    return resultat
def comparer_listes(liste1, liste2, seuil=75):

    correspondances = []
    non_trouves = []

    for val1 in liste1:
        match_trouve = False

        for val2 in liste2:
            score = fuzz.token_set_ratio(val1, val2)

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
        
        # Mapping clean → original
        # 🔒 sécuriser colonnes Excel
        colonnes_valides_excel = [c for c in colonnes_excel if c in df_excel.columns]

        if not colonnes_valides_excel:
            st.error("Aucune colonne Excel valide")
            st.stop()

        df_excel_temp = df_excel.copy()
        df_excel_temp["original"] = (
            df_excel[colonnes_valides_excel]
            .fillna("")
            .astype(str)
            .apply(lambda row: " ".join(row.values), axis=1)
        )
        df_excel_temp["clean"] = df_excel_temp["original"].apply(nettoyer_texte)
       
        df_pdf_temp = df_pdf.copy()
        df_pdf_temp["original"] = df_pdf[colonnes_pdf].astype(str).agg(" ".join, axis=1)
        df_pdf_temp["clean"] = df_pdf_temp["original"].apply(nettoyer_texte)

        dict_excel = dict(zip(df_excel_temp["clean"], df_excel_temp["original"]))
        dict_pdf = dict(zip(df_pdf_temp["clean"], df_pdf_temp["original"]))

        liste_excel = creer_cle(df_excel, colonnes_excel).tolist()
        liste_pdf = creer_cle(df_pdf, colonnes_pdf).tolist()

        liste_excel = dedoublonner_fuzzy(liste_excel)
        liste_pdf = dedoublonner_fuzzy(liste_pdf)

        seuil = 75

        correspondances = []
        only_excel = []
        only_pdf = []

        pdf_utilises = set()

        # Excel → PDF
        for val_excel in liste_excel:

            meilleur_score = 0
            meilleur_match = None

            for val_pdf in liste_pdf:

                if val_pdf in pdf_utilises:
                    continue

                score = fuzz.token_set_ratio(val_excel, val_pdf)

                if score > meilleur_score:
                    meilleur_score = score
                    meilleur_match = val_pdf

            if meilleur_score >= seuil:
                correspondances.append((
    dict_excel.get(val_excel, val_excel),
    dict_pdf.get(meilleur_match, meilleur_match),
    meilleur_score))
                pdf_utilises.add(meilleur_match)
            else:
                only_excel.append(dict_excel.get(val_excel, val_excel))
               

        # PDF non matchés
        for val_pdf in liste_pdf:
            if val_pdf not in pdf_utilises:
                only_pdf.append(dict_pdf.get(val_pdf, val_pdf))
        # ---------------------------
        # 🔥 Affichage (BIEN INDENTÉ)
        # ---------------------------

        nom_col_excel = " + ".join(colonnes_excel)
        nom_col_pdf = " + ".join(colonnes_pdf)

        # Correspondances
        st.subheader("Correspondances trouvées")
        df_corr = pd.DataFrame(correspondances,
        columns=[f"Excel ({nom_col_excel})", f"PDF ({nom_col_pdf})", "Score"])
        st.dataframe(df_corr, use_container_width=True)

        # Excel absent
        only_excel = [x for x in only_excel if str(x).strip().lower() not in ["", "nan"]]
        st.subheader("Présents dans Excel mais absents dans PDF")
        df_only_excel = pd.DataFrame({
        "N°": range(1, len(only_excel) + 1),
        f"Excel ({nom_col_excel})": only_excel})
        st.dataframe(df_only_excel, use_container_width=True, hide_index=True)

        # PDF absent
        st.subheader("Présents dans PDF mais absents dans Excel")
        df_only_pdf = pd.DataFrame({"N°": range(1, len(only_pdf) + 1),f"PDF ({nom_col_pdf})": only_pdf
         })
        st.dataframe(df_only_pdf, use_container_width=True, hide_index=True)

    else:
        st.warning("Charge les fichiers ET sélectionne les colonnes")