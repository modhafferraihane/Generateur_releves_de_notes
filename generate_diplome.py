"""
Genere un diplome national (.docx) par etudiant a partir d'un modele a
placeholders (modeles/Diplome Licence.docx ou Diplome Mastere.docx).

Le modele contient des jetons [[NOM]], [[CIN]], [[SPECIALITE]], etc. que l'on
remplace ici par les donnees de l'etudiant. Le texte d'un diplome vit dans des
zones de texte (souvent dupliquees pour la compatibilite avec les anciennes
versions de Word), et python-docx ne les expose pas via doc.paragraphs : on
remplace donc directement dans TOUS les runs <w:r> du document.

Champs remplis automatiquement : nom, date/lieu de naissance, CIN, annee
universitaire, domaine, mention (LMD), specialite, mention d'honneur (calculee
depuis la moyenne annuelle) et date (Fait a Tunis, le ...). Pour le mastere,
les moyennes par semestre / credits restent a completer a la main (le PV de la
2e annee ne couvre pas la 1re).
"""
from pathlib import Path

import docx
from docx.oxml.ns import qn


def _fill_runs(doc, values):
    """Remplace chaque jeton [[...]] par sa valeur dans tous les runs du
    document (zones de texte comprises)."""
    for r in doc.element.iter(qn("w:r")):
        for t in r.findall(qn("w:t")):
            txt = t.text or ""
            if "[[" not in txt:
                continue
            for token, val in values.items():
                if token in txt:
                    txt = txt.replace(token, val)
            t.text = txt


def fill_diplome(template_path, output_path, *, nom_prenom, civilite="M.", ne="Né",
                 naissance_date=None, naissance_lieu="", cin="", annee_universitaire="",
                 domaine="", mention="", specialite="", honneur="", date_diplome=""):
    doc = docx.Document(template_path)

    naissance_str = (naissance_date.strftime("%d/%m/%Y")
                     if hasattr(naissance_date, "strftime") else (naissance_date or ""))
    annee_slash = (annee_universitaire or "").replace("-", "/")  # '2024-2025' -> '2024/2025'

    values = {
        "[[CIVILITE]]": civilite or "M.",   # 'M.' ou 'Mme' selon le sexe
        "[[NE]]": ne or "Né",               # 'Né' ou 'Née'
        "[[NOM]]": nom_prenom or "",
        "[[NAISSANCE_DATE]]": naissance_str,
        "[[NAISSANCE_LIEU]]": naissance_lieu or "",
        "[[CIN]]": cin or "",
        "[[ANNEE]]": annee_slash,
        "[[DOMAINE]]": domaine or "",
        "[[MENTION]]": mention or "",
        "[[SPECIALITE]]": specialite or "",
        "[[HONNEUR]]": honneur or "",
        "[[DATE]]": date_diplome or "",
    }
    _fill_runs(doc, values)

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)
