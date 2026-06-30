"""
Genere un releve de notes (.xlsx) par etudiant a partir d'un ou deux PV de
deliberation (Semestre 1 et/ou Semestre 2), pour n'importe quelle filiere
(le nombre d'UE/ECUE par semestre est detecte dynamiquement depuis le PV,
le bloc UE du gabarit est redimensionne en consequence).

Usage CLI:
    python generate_releves.py <pv_s1.xlsx> [pv_s2.xlsx] [--template gabarit.xlsx] [--out dossier]

Ce module est aussi importe par app.py (interface web).

Hypotheses actuelles (a ajuster si la situation change) :
- "Decision du jury" / mention laissees vides (remplissage manuel par le jury).
- Nom & Prenom recopie tel quel depuis le PV (ordre "Nom Prenom", pas inverse).
- Une UE est consideree "Validee/Capitalisee" (V/C) si ses Credits Capitalises
  (colonne du PV) sont > 0 ; a corriger si la regle exacte de l'etablissement
  differe.
- Si un etudiant n'apparait que dans un des deux PV fournis, le releve est
  quand meme genere avec les donnees disponibles (l'autre semestre reste vide)
  et un avertissement est renvoye par generate_all().
- Date/lieu de naissance et N CIN/Passeport viennent du fichier de
  coordonnees des etudiants (optionnel) si fourni, en faisant correspondre
  les etudiants par nom ; sinon laisses vides (remplissage manuel).
"""
import copy
import re
import sys
import unicodedata
from dataclasses import dataclass
from pathlib import Path

import openpyxl

BASE_DIR = Path(__file__).resolve().parent
PV_SHEET_NAME = "PV_Deliberation"
NAME_COLUMN = 2  # colonne B : "Nom & Prenom"

# Bloc UE/ECUE du gabarit : ligne de depart, et ligne ou commence le pied de
# page (TOTAL/Moyenne Annuelle/Decision...) dans le gabarit d'origine. Le
# bloc est redimensionne dynamiquement (insert_rows/delete_rows) pour
# s'adapter au nombre reel d'UE/ECUE trouve dans le PV, donc ces constantes
# ne sont que le point de depart du calcul, pas une contrainte sur le PV.
UE_BLOCK_FIRST_ROW = 19
FOOTER_FIRST_ROW = 40  # ligne "TOTAL" dans le gabarit d'origine

# Filieres proposees dans l'interface -> code utilise pour retrouver l'onglet
# correspondant dans le fichier de coordonnees des etudiants (ex. onglet
# "L1-BD" pour le code "BD" au niveau L1).
FILIERES = {
    "Licence Nationale en Génie Logiciel et Système d'Information": "GLSI",
    "Licence Nationale en Big Data et Analyse des Données": "BD",
    "Licence en Ingénierie des Systèmes et Réseaux": "RSYS",
    "Mastère en Cloud Computing et Virtualisation": "CLOUD",
    "Mastère en Cybersécurité": "CYBER",
}

# Niveau -> (chiffre, suffixe ordinal) pour l'Attestation de Reussite
# (ex. L1 -> "1" + "ère" = "1ère année").
NIVEAU_TEXTE = {
    "L1": ("1", "ère"), "L2": ("2", "ème"), "L3": ("3", "ème"),
    "M1": ("1", "ère"), "M2": ("2", "ème"),
}

# Niveaux proposes dans l'interface selon le cycle (3 ans de Licence, 2 ans
# de Mastere).
NIVEAUX_PAR_CYCLE = {
    "Licence": [("L1", "1ère année"), ("L2", "2ème année"), ("L3", "3ème année")],
    "Mastère": [("M1", "1ère année"), ("M2", "2ème année")],
}


def cycle_de_filiere(label):
    return "Licence" if label.startswith("Licence") else "Mastère" if label.startswith("Mastère") else ""


@dataclass
class Ecue:
    code: str
    name: str
    coef: float
    credit: float
    cc_col: int
    ex_col: int
    moy_col: int
    cr_col: int


@dataclass
class UniteEnseignement:
    label: str
    semestre: str
    coef: float
    credit: float
    ecues: list
    moyenne_ue_col: int
    credits_cap_col: int


@dataclass
class StudentRecord:
    num: object
    nom_prenom: str
    row: int
    grades: dict          # ecue_code -> (cc, ex, moy, cr)
    ue_results: dict       # ue_label -> (moyenne_ue, credits_capitalises)
    moyenne_semestre: float
    total_credits_semestre: int


def _clean(text):
    if text is None:
        return ""
    return re.sub(r"\s+", " ", str(text)).strip()


def normalize_name(text):
    """Normalise un nom pour comparaison (insensible aux accents, a la
    casse et aux espaces multiples)."""
    text = unicodedata.normalize("NFKD", _clean(text))
    text = "".join(c for c in text if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", text).strip().upper()


def parse_label_value(ws, rows, label):
    """Cherche 'label :' dans les lignes donnees, meme quand label et valeur
    sont dans la meme cellule ou dans deux cellules separees de la meme ligne."""
    for row in rows:
        cells = [ws.cell(row=row, column=c) for c in range(1, ws.max_column + 1)]
        for idx, cell in enumerate(cells):
            val = cell.value
            if not isinstance(val, str) or label not in val:
                continue
            after = val.split(label, 1)[1].strip(" : ")
            if after:
                return _clean(after)
            for nxt in cells[idx + 1:]:
                if nxt.value not in (None, ""):
                    return _clean(nxt.value)
    return ""


def merges_starting_at(ws, row, min_col=1, max_col=None):
    max_col = max_col or ws.max_column
    spans = [mc for mc in ws.merged_cells.ranges
             if mc.min_row == row and min_col <= mc.min_col <= max_col]
    return sorted(spans, key=lambda r: r.min_col)


def parse_pv(path, sheet_name=PV_SHEET_NAME):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.worksheets[0]

    domaine = parse_label_value(ws, range(1, 10), "Domaine")
    filiere = parse_label_value(ws, range(1, 10), "Filière")
    specialite = parse_label_value(ws, range(1, 10), "Spécialité")

    header_row = None
    for row in range(1, 30):
        val = ws.cell(row=row, column=NAME_COLUMN).value
        if val and "Nom" in str(val):
            header_row = row
            break
    if header_row is None:
        raise ValueError("Ligne d'en-tete introuvable (colonne B contenant 'Nom').")

    ue_name_row = header_row - 6
    ecue_name_row = header_row - 5
    coef_ecue_row = header_row - 4
    credit_ecue_row = header_row - 3
    coef_ue_row = header_row - 2
    credit_ue_row = header_row - 1
    semester_row = ue_name_row - 1
    summary_max_row = header_row

    ues = []
    for ue_mc in merges_starting_at(ws, ue_name_row, min_col=NAME_COLUMN + 1):
        if ue_mc.max_row != ue_name_row:
            continue  # exclut les merges verticaux de synthese (Moyenne Semestre, Total Credits)
        label = _clean(ws.cell(row=ue_name_row, column=ue_mc.min_col).value)
        if not label:
            continue

        # Le merge "Semestre X" ne couvre pas toujours toutes les colonnes
        # des UE de ce semestre (vu sur certains PV) -> on rattache l'UE au
        # dernier "Semestre X" dont le merge demarre avant ou a sa colonne,
        # plutot que d'exiger une inclusion stricte dans le merge.
        semestre = ""
        for sem_mc in sorted(merges_starting_at(ws, semester_row), key=lambda mc: mc.min_col):
            if sem_mc.min_col > ue_mc.min_col:
                break
            semestre = _clean(ws.cell(row=semester_row, column=sem_mc.min_col).value)

        ecue_spans = [
            mc for mc in merges_starting_at(ws, ecue_name_row, ue_mc.min_col, ue_mc.max_col)
            if mc.max_row == ecue_name_row
        ]
        ecues = []
        for span in ecue_spans:
            raw = _clean(ws.cell(row=ecue_name_row, column=span.min_col).value).replace("\n", " ")
            m = re.search(r"\(([^)]+)\)\s*$", raw)
            code = m.group(1) if m else raw
            name = raw.split("(")[0].strip()
            ecues.append(Ecue(
                code=code,
                name=name,
                coef=ws.cell(row=coef_ecue_row, column=span.min_col).value,
                credit=ws.cell(row=credit_ecue_row, column=span.min_col).value,
                cc_col=span.min_col,
                ex_col=span.min_col + 1,
                moy_col=span.min_col + 2,
                cr_col=span.min_col + 3,
            ))

        # Les colonnes "Moyenne UE" / "Credits Capitalises" suivent toujours
        # directement la derniere ECUE. On ne se fie pas a la largeur du merge
        # du nom d'UE (row11) car elle est incoherente d'un bloc a l'autre
        # dans ce fichier (parfois elle inclut ces 2 colonnes, parfois non).
        last_ecue_max_col = max(span.max_col for span in ecue_spans)
        moyenne_ue_col, credits_cap_col = last_ecue_max_col + 1, last_ecue_max_col + 2
        found = {
            mc.min_col for mc in ws.merged_cells.ranges
            if mc.min_row in (ue_name_row, ecue_name_row) and mc.max_row == summary_max_row
            and mc.min_col in (moyenne_ue_col, credits_cap_col)
        }
        if found != {moyenne_ue_col, credits_cap_col}:
            raise ValueError(
                f"UE '{label}': colonnes de synthese attendues en {moyenne_ue_col}/{credits_cap_col} "
                f"non confirmees par un merge (Moyenne UE / Credits Capitalises)."
            )

        ues.append(UniteEnseignement(
            label=label,
            semestre=semestre,
            coef=ws.cell(row=coef_ue_row, column=ue_mc.min_col).value,
            credit=ws.cell(row=credit_ue_row, column=ue_mc.min_col).value,
            ecues=ecues,
            moyenne_ue_col=moyenne_ue_col,
            credits_cap_col=credits_cap_col,
        ))

    semestre_summary = []
    for mc in ws.merged_cells.ranges:
        if mc.max_row == summary_max_row and mc.min_row in (ue_name_row, ecue_name_row):
            txt = _clean(ws.cell(row=mc.min_row, column=mc.min_col).value)
            if "Moyenne Semestre" in txt or "Total des Cr" in txt:
                semestre_summary.append((txt, mc.min_col))
    moyenne_sem_col = next(c for t, c in semestre_summary if "Moyenne" in t)
    total_credits_col = next(c for t, c in semestre_summary if "Total" in t)

    students = []
    row = header_row + 1
    while True:
        nom = ws.cell(row=row, column=NAME_COLUMN).value
        if nom is None or _clean(nom) == "":
            break
        grades, ue_results = {}, {}
        for ue in ues:
            for ec in ue.ecues:
                grades[ec.code] = (
                    ws.cell(row=row, column=ec.cc_col).value or 0,
                    ws.cell(row=row, column=ec.ex_col).value or 0,
                    ws.cell(row=row, column=ec.moy_col).value or 0,
                    ws.cell(row=row, column=ec.cr_col).value or 0,
                )
            ue_results[ue.label] = (
                ws.cell(row=row, column=ue.moyenne_ue_col).value or 0,
                ws.cell(row=row, column=ue.credits_cap_col).value or 0,
            )
        students.append(StudentRecord(
            num=ws.cell(row=row, column=1).value,
            nom_prenom=_clean(nom),
            row=row,
            grades=grades,
            ue_results=ue_results,
            moyenne_semestre=ws.cell(row=row, column=moyenne_sem_col).value or 0,
            total_credits_semestre=ws.cell(row=row, column=total_credits_col).value or 0,
        ))
        row += 1

    meta = {"domaine": domaine, "filiere": filiere, "specialite": specialite}
    return ues, students, meta


def parse_filename_meta(path):
    name = path.stem
    annee = ""
    m = re.search(r"AU(\d{2})-(\d{2})", name)
    if m:
        annee = f"20{m.group(1)}-20{m.group(2)}"
    niveau = ""
    m = re.search(r"\b([LM][1-3])\b", name.upper())
    if m:
        niveau = m.group(1)
    return annee, niveau


def sanitize_filename(text):
    text = unicodedata.normalize("NFKD", text)
    text = "".join(c for c in text if not unicodedata.combining(c))
    return re.sub(r"[^A-Za-z0-9]+", "_", text).strip("_")


def sanitize_sheet_name(text):
    return re.sub(r"[\\/*?:\[\]]", "", text)[:31] or "Etudiant"


def to_number(value):
    """Convertit une valeur de cellule en float, en tolerant les nombres
    saisis comme texte avec virgule decimale (ex. '16,5') vus sur certains PV."""
    if value in (None, ""):
        return 0
    if isinstance(value, str):
        value = value.strip().replace(",", ".")
    return float(value)


def format_note(value):
    return round(to_number(value), 2) if value not in (None, "") else 0


def semester_digit(label):
    m = re.search(r"(\d)", label or "")
    return m.group(1) if m else None


def merge_students(students_per_pv):
    """Fusionne les etudiants de plusieurs PV (par Nom & Prenom). Renvoie un
    dict nom_prenom -> {"grades", "ue_results", "present_in"}."""
    merged = {}
    for idx, students in enumerate(students_per_pv):
        for s in students:
            rec = merged.setdefault(s.nom_prenom, {"grades": {}, "ue_results": {}, "present_in": set()})
            rec["grades"].update(s.grades)
            rec["ue_results"].update(s.ue_results)
            rec["present_in"].add(idx)
    return merged


def annual_aggregates(ue_results, ue_by_label):
    """Moyenne ponderee (par coefficient d'UE) et total des credits
    capitalises, sur toutes les UE presentes dans ue_results."""
    total_weight = 0.0
    weighted_sum = 0.0
    total_credits = 0
    for label, (moyenne_ue, credits_cap) in ue_results.items():
        ue = ue_by_label[label]
        coef = to_number(ue.coef)
        total_weight += coef
        weighted_sum += to_number(moyenne_ue) * coef
        total_credits += int(to_number(credits_cap))
    moyenne = round(weighted_sum / total_weight, 2) if total_weight else 0
    return moyenne, total_credits


# ---------------------------------------------------------------------------
# Bloc UE/ECUE dynamique : le gabarit fournit deux lignes-modeles (1ere ligne
# d'une UE, et ligne de continuation pour une UE a plusieurs ECUE) que l'on
# duplique/style autant de fois que necessaire pour coller exactement a la
# structure UE/ECUE trouvee dans le PV, quelle que soit la filiere.
# ---------------------------------------------------------------------------

_STYLE_COLUMNS = range(1, 21)  # A..T


def _row_style_snapshot(ws, row):
    snapshot = {}
    for col in _STYLE_COLUMNS:
        cell = ws.cell(row=row, column=col)
        snapshot[col] = {
            "font": copy.copy(cell.font),
            "border": copy.copy(cell.border),
            "fill": copy.copy(cell.fill),
            "alignment": copy.copy(cell.alignment),
            "number_format": cell.number_format,
        }
    return snapshot


def _apply_row_style(ws, row, snapshot):
    for col, style in snapshot.items():
        cell = ws.cell(row=row, column=col)
        cell.font = style["font"]
        cell.border = style["border"]
        cell.fill = style["fill"]
        cell.alignment = style["alignment"]
        cell.number_format = style["number_format"]


def build_row_layout(ues_ordered):
    """Calcule, pour une liste d'UE (deja dans l'ordre du/des PV, donc
    regroupees par semestre), la ligne de depart de chaque UE dans le bloc
    dynamique, le nombre total de lignes ECUE, et les plages de lignes par
    semestre (pour la colonne "Sem")."""
    row_of = {}
    row = UE_BLOCK_FIRST_ROW
    spans = []
    if ues_ordered:
        seg_label = ues_ordered[0].semestre
        seg_start = row
        for ue in ues_ordered:
            if ue.semestre != seg_label:
                spans.append((seg_label, seg_start, row - 1))
                seg_label, seg_start = ue.semestre, row
            row_of[ue.label] = row
            row += len(ue.ecues)
        spans.append((seg_label, seg_start, row - 1))
    n_rows = row - UE_BLOCK_FIRST_ROW
    return row_of, n_rows, spans


def _write_dynamic_ue_block(ws, ues_ordered, row_of, n_rows, semester_spans):
    """Redimensionne le bloc UE/ECUE du gabarit (deja charge dans ws) pour
    contenir exactement n_rows lignes ECUE, puis ecrit le contenu structurel
    (libelles UE/ECUE, coefficients, credits) et les fusions correspondantes.
    Renvoie le decalage (diff) applique au pied de page."""
    style_first = _row_style_snapshot(ws, UE_BLOCK_FIRST_ROW)
    style_continuation = _row_style_snapshot(ws, UE_BLOCK_FIRST_ROW + 1)

    # openpyxl ne decale PAS les fusions existantes lors de insert_rows /
    # delete_rows (seules les valeurs/styles des cellules le sont) : on
    # retire nous-memes toutes les fusions a partir de la ligne 19 (bloc UE
    # d'origine + pied de page + la bordure decorative colonne O qui
    # traverse les deux), et on recree celles du pied de page / colonne O au
    # bon endroit une fois le bloc redimensionne. Celles du bloc UE
    # d'origine (19 a 39) sont abandonnees : on les reconstruit nous-memes
    # plus bas, adaptees a la structure UE/ECUE reelle.
    carried_merges = []  # (min_row, max_row, min_col, max_col, shift_min_row)
    for mc in list(ws.merged_cells.ranges):
        if mc.max_row < UE_BLOCK_FIRST_ROW:
            continue  # zone d'en-tete, jamais touchee
        ws.unmerge_cells(str(mc))
        if mc.min_row < UE_BLOCK_FIRST_ROW:
            # fusion qui traverse le bloc depuis l'en-tete (bordure colonne O) :
            # le haut ne bouge pas, seul le bas suit le bloc UE.
            carried_merges.append((mc.min_row, mc.max_row, mc.min_col, mc.max_col, False))
        elif mc.min_row >= FOOTER_FIRST_ROW:
            # fusion du pied de page : suit integralement le decalage.
            carried_merges.append((mc.min_row, mc.max_row, mc.min_col, mc.max_col, True))
        # sinon (19 <= min_row < 40) : ancienne fusion du bloc UE, abandonnee.

    diff = n_rows - (FOOTER_FIRST_ROW - UE_BLOCK_FIRST_ROW)
    if diff > 0:
        ws.insert_rows(FOOTER_FIRST_ROW, diff)
    elif diff < 0:
        ws.delete_rows(FOOTER_FIRST_ROW + diff, -diff)

    for orig_min_row, orig_max_row, min_col, max_col, shift_min_row in carried_merges:
        new_min_row = orig_min_row + diff if shift_min_row else orig_min_row
        ws.merge_cells(start_row=new_min_row, start_column=min_col,
                        end_row=orig_max_row + diff, end_column=max_col)

    for ue in ues_ordered:
        start = row_of[ue.label]
        n = len(ue.ecues)
        for offset, ec in enumerate(ue.ecues):
            r = start + offset
            _apply_row_style(ws, r, style_first if offset == 0 else style_continuation)
            ws.cell(row=r, column=3, value=ec.name)   # C : nom ECUE
            ws.cell(row=r, column=6, value=ec.coef)    # F : coef ECUE
            ws.cell(row=r, column=7, value=ec.credit)  # G : credit ECUE
            ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)  # C:D
        ws.cell(row=start, column=2, value=ue.label)   # B : nom UE
        ws.cell(row=start, column=8, value=ue.credit)  # H : credit UE
        if n > 1:
            ws.merge_cells(start_row=start, start_column=2, end_row=start + n - 1, end_column=2)
            ws.merge_cells(start_row=start, start_column=8, end_row=start + n - 1, end_column=8)
            for col in (17, 18, 19, 20):  # Q,R,S,T : synthese UE
                ws.merge_cells(start_row=start, start_column=col, end_row=start + n - 1, end_column=col)

    for label, start, end in semester_spans:
        digit = semester_digit(label)
        ws.cell(row=start, column=5, value=f"S{digit}" if digit else label)  # E : Sem
        if end > start:
            ws.merge_cells(start_row=start, start_column=5, end_row=end, end_column=5)

    return diff


# ---------------------------------------------------------------------------
# Fichier de coordonnees des etudiants (date/lieu de naissance, CIN, etc.)
# ---------------------------------------------------------------------------

def _normalize_sheet_key(name):
    return re.sub(r"[^A-Z0-9]", "", (name or "").upper())


def find_coordinates_sheet(wb, niveau, code):
    """Cherche l'onglet correspondant au niveau+filiere (ex. niveau='L1',
    code='BD' -> onglet 'L1-BD'). Renvoie None si aucun onglet ne correspond."""
    target = _normalize_sheet_key(f"{niveau}{code}")
    if not target:
        return None
    for name in wb.sheetnames:
        if target in _normalize_sheet_key(name):
            return wb[name]
    return None


def parse_student_coordinates(path, niveau, code):
    """Lit l'onglet niveau+filiere du fichier de coordonnees des etudiants.
    Renvoie un dict {nom_normalise: {naissance_date, naissance_lieu,
    nationalite, cin}}, ou {} si l'onglet est introuvable."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = find_coordinates_sheet(wb, niveau, code)
    if ws is None:
        return {}

    header_row = None
    for row in range(1, 20):
        val = ws.cell(row=row, column=2).value
        if val and "اسم" in str(val):
            header_row = row
            break
    if header_row is None:
        return {}

    coords = {}
    row = header_row + 2  # +1 = sous-entetes (Baccalaureat), +1 = 1ere ligne de donnees
    while True:
        prenom = ws.cell(row=row, column=2).value
        nom = ws.cell(row=row, column=3).value
        if not _clean(prenom) and not _clean(nom):
            break
        key = normalize_name(f"{_clean(nom)} {_clean(prenom)}")
        cin = ws.cell(row=row, column=8).value
        coords[key] = {
            "naissance_date": ws.cell(row=row, column=5).value,
            "naissance_lieu": _clean(ws.cell(row=row, column=6).value),
            "nationalite": _clean(ws.cell(row=row, column=7).value),
            "cin": _clean(cin) if cin is not None else "",
        }
        row += 1
    return coords


def format_naissance(info):
    date_val = info.get("naissance_date")
    lieu = info.get("naissance_lieu") or ""
    date_str = date_val.strftime("%d/%m/%Y") if hasattr(date_val, "strftime") else _clean(date_val)
    if date_str and lieu:
        return f"{date_str} à {lieu}"
    return date_str or lieu


def fill_releve(template_path, output_path, nom_prenom, grades, ue_results,
                 ues_ordered, row_of, n_rows, semester_spans,
                 moyenne_annuelle, total_credits, meta, coordonnees=None):
    wb = openpyxl.load_workbook(template_path)
    ws = wb.worksheets[0]

    coordonnees = coordonnees or {}
    ws["D9"] = nom_prenom
    ws["D10"] = format_naissance(coordonnees)
    ws["D11"] = coordonnees.get("cin", "")
    ws["L8"] = f"Année universitaire : {meta['annee_universitaire']}"
    ws["C13"] = meta["diplome"]
    ws["C14"] = meta["domaine"]
    ws["L14"] = f"Niveau d'études : {meta['niveau']}"
    ws["C15"] = meta["specialite"]

    diff = _write_dynamic_ue_block(ws, ues_ordered, row_of, n_rows, semester_spans)

    for ue in ues_ordered:
        ue_row = row_of[ue.label]
        if ue.label in ue_results:
            moyenne_ue, credits_cap = ue_results[ue.label]
            credits_cap = int(to_number(credits_cap))
            ws[f"Q{ue_row}"] = credits_cap
            ws[f"R{ue_row}"] = "V" if credits_cap > 0 else None
            ws[f"S{ue_row}"] = "C" if credits_cap > 0 else None
            ws[f"T{ue_row}"] = meta["annee_universitaire"]
            for offset, ec in enumerate(ue.ecues):
                r = ue_row + offset
                cc, ex, moy, cr = grades.get(ec.code, (None, None, None, None))
                ws[f"J{r}"] = format_note(cc) if cc is not None else None
                ws[f"K{r}"] = format_note(ex) if ex is not None else None
                ws[f"L{r}"] = format_note(moy) if moy is not None else None
        # sinon : UE non couverte par les PV fournis -> deja vide (lignes neuves)

    total_row = FOOTER_FIRST_ROW + diff
    moyenne_row, controle_row = 42 + diff, 43 + diff
    credits_row = 45 + diff
    mention_row = 48 + diff
    print_end_row = 50 + diff

    total_credits_possible = ws[f"G{total_row}"].value
    ws[f"G{moyenne_row}"] = moyenne_annuelle
    ws[f"G{controle_row}"] = "-"
    ws[f"F{credits_row}"] = f"{total_credits}/{total_credits_possible}"
    ws[f"C{mention_row}"] = ""  # decision/mention : remplissage manuel par le jury

    # Mise en page prete a l'impression (1 page A4).
    ws.print_area = f"A1:T{print_end_row}"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    ws.title = sanitize_sheet_name(nom_prenom)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def generate_all(pv_paths, template_path, output_dir, *,
                  coords_path=None, filiere_code=None, ar_template_path=None, niveau=None):
    """Pipeline complet : parse 1..N PV, fusionne par etudiant, ecrit un
    releve (et une attestation de reussite si ar_template_path est fourni)
    par etudiant dans output_dir. Renvoie (fichiers_generes, avertissements).

    niveau (ex. "L1", "M2") est fourni explicitement par l'appelant (menu
    deroulant de l'interface) ; a defaut, on retombe sur une detection
    depuis le nom du 1er fichier PV (utile pour la CLI)."""
    parsed = [parse_pv(p) for p in pv_paths]
    ues_per_pv = [p[0] for p in parsed]
    students_per_pv = [p[1] for p in parsed]
    pv_meta = parsed[0][2]

    annee_universitaire, niveau_detecte = parse_filename_meta(pv_paths[0])
    niveau = niveau or niveau_detecte
    cycle = "Licence" if niveau.startswith("L") else "Mastère" if niveau.startswith("M") else ""
    meta = {
        "annee_universitaire": annee_universitaire,
        "niveau": niveau,
        "domaine": pv_meta["domaine"],
        "specialite": pv_meta["specialite"],
        "diplome": f"{cycle} en {pv_meta['specialite']}".strip(),
    }

    ues_ordered = [ue for ues in ues_per_pv for ue in ues]
    ue_by_label = {ue.label: ue for ue in ues_ordered}
    row_of, n_rows, semester_spans = build_row_layout(ues_ordered)
    merged = merge_students(students_per_pv)

    coords = {}
    warnings = []
    if coords_path and filiere_code:
        coords = parse_student_coordinates(coords_path, niveau, filiere_code)
        if not coords:
            warnings.append(
                f"Fichier de coordonnées : aucun onglet trouvé pour {niveau} {filiere_code}."
            )

    ar = None
    if ar_template_path:
        import generate_attestation as ar  # import tardif : optionnel (depend de python-docx)

    output_dir.mkdir(parents=True, exist_ok=True)
    generated = []
    for nom_prenom, rec in merged.items():
        moyenne, total_credits = annual_aggregates(rec["ue_results"], ue_by_label)
        etu_coords = coords.get(normalize_name(nom_prenom)) if coords else None
        if coords and not etu_coords:
            warnings.append(f"{nom_prenom} : coordonnées non trouvées dans le fichier fourni")

        out_path = output_dir / f"Releve_{sanitize_filename(nom_prenom)}.xlsx"
        fill_releve(template_path, out_path, nom_prenom, rec["grades"], rec["ue_results"],
                    ues_ordered, row_of, n_rows, semester_spans,
                    moyenne, total_credits, meta, coordonnees=etu_coords)
        generated.append(out_path)

        if ar_template_path:
            ar_path = output_dir / f"AR_{sanitize_filename(nom_prenom)}.docx"
            niveau_num, niveau_suffixe = NIVEAU_TEXTE.get(niveau, ("", ""))
            ar.fill_attestation(
                ar_template_path, ar_path, nom_prenom,
                naissance_date=(etu_coords or {}).get("naissance_date"),
                cin=(etu_coords or {}).get("cin", ""),
                nationalite=(etu_coords or {}).get("nationalite") or "",
                cycle=cycle,
                niveau_num=niveau_num,
                niveau_suffixe=niveau_suffixe,
                specialite=meta["specialite"],
            )

        if len(pv_paths) > 1 and len(rec["present_in"]) < len(pv_paths):
            missing = [i + 1 for i in range(len(pv_paths)) if i not in rec["present_in"]]
            warnings.append(f"{nom_prenom} : absent du PV Semestre {', '.join(map(str, missing))}")
        if rec["grades"] and all(g[2] == 0 for g in rec["grades"].values()):
            warnings.append(f"{nom_prenom} : toutes les notes sont a 0 (probablement absent(e))")

    return generated, warnings


def main():
    import argparse
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pv", nargs="*", help="PV de deliberation (S1, et optionnellement S2)")
    parser.add_argument("--template", type=Path, default=None)
    parser.add_argument("--out", type=Path, default=BASE_DIR / "releves_generes")
    args = parser.parse_args()

    template_path = args.template or next(BASE_DIR.glob("Exemple*.xlsx"))
    pv_paths = [Path(p) for p in args.pv] if args.pv else [next(BASE_DIR.glob("PV*.xlsx"))]

    generated, warnings = generate_all(pv_paths, template_path, args.out)

    print(f"{len(generated)} releve(s) genere(s) dans {args.out}")
    for path in generated:
        print(f"  - {path.name}")
    if warnings:
        print("\nA verifier :")
        for w in warnings:
            print(f"  - {w}")


if __name__ == "__main__":
    main()
