# Mémoire du projet — Générateur de relevés de notes

Ce fichier sert de mémoire du projet : il retrace le besoin tel qu'il a évolué
au fil de nos échanges, les décisions prises (et pourquoi), et l'état actuel
des règles métier. À relire avant de reprendre le projet après une pause, et
à tenir à jour à chaque nouvelle étape importante (nouveau besoin, décision,
changement de règle).

## Vue d'ensemble

Outil qui génère automatiquement un relevé de notes (.xlsx, et en option
.pdf prêt à imprimer) par étudiant, à partir des PV de délibération Excel
(Semestre 1 et Semestre 2). Deux interfaces : application web (`app.py`) et
ligne de commande (`generate_releves.py`).

Contexte : programme "Licence Big Data et Analyse de Données" L1, en
Tunisie (jury signé "Skander CHRIGUI", lieu "Tunis").

---

## Étape 1 — Génération de base (PV → relevés .xlsx)

**Besoin initial :** à partir d'un PV de délibération Excel (une ligne par
étudiant, toutes les notes), générer automatiquement un relevé de notes par
étudiant, sur la base de deux formats fournis en capture d'écran (le PV
source et le relevé cible).

**Décisions prises :**
- Stack technique : Python + openpyxl, plutôt que PowerShell/Excel COM —
  plus robuste, plus facile à étendre, export PDF possible plus tard.
  → Python installé via winget (l'alias `python` du Windows Store ne
  fonctionnait pas).
- Sortie : un fichier `.xlsx` par étudiant (pas un classeur multi-onglets).
- Le PV fourni ne couvrait que le Semestre 1 → générer avec S1 seulement
  pour l'instant, lignes S2 laissées vides dans le relevé.
- "Décision du jury" / mention : laissées vides — c'est le jury qui les
  remplit à la main, pas un calcul automatique.
- Date/lieu de naissance et N° CIN/Passeport : absents du PV (et de ses
  feuilles annexes `FN_*`) → laissés vides, pas de fichier registre fourni.
- Nom & Prénom : recopié tel quel depuis le PV (ordre "Nom Prénom"), sans
  inversion — pour ne pas mal découper les noms composés (ex. "Louhichi
  Mohamed Mortadha").

**Réalisé :** `generate_releves.py`, qui détecte dynamiquement la structure
du PV (UE/ECUE, coefficients, crédits) via les cellules fusionnées plutôt
que des colonnes codées en dur, et remplit une copie du gabarit
`Exemple Relevé de Notes.xlsx` par étudiant.

**Règle déduite (non confirmée à 100%, à un seul exemple) :** une UE est
"Validée/Capitalisée" (colonnes V/C du relevé) si ses crédits capitalisés
(donnée du PV) sont > 0.

---

## Étape 2 — Interface web, fusion S1+S2, export PDF

**Besoin :** une interface web avec un bouton qui demande le PV, et qui ne
génère les relevés qu'à condition d'avoir le PV du Semestre 1 **et** celui
du Semestre 2 ; un README avec toutes les étapes d'installation des
dépendances ; une option pour générer aussi un PDF prêt à imprimer.

**Réalisé :**
- `app.py` (Flask) : formulaire avec les deux PV obligatoires (bouton
  désactivé tant que les deux ne sont pas choisis), case à cocher PDF, page
  de résultat avec téléchargement individuel ou ZIP global.
- `generate_releves.py` réécrit pour fusionner les deux PV par étudiant
  (rapprochement par "Nom & Prénom"), et recalculer la moyenne annuelle et
  le total des crédits en pondérant par le coefficient de chaque UE.
- `export_pdf.py` : conversion `.xlsx` → `.pdf` via Excel en COM
  (`pywin32`) — pas de LibreOffice sur la machine, donc Excel est la seule
  voie disponible ; mise à l'échelle 1 page A4.
- `requirements.txt` et `README.md` (installation Python, dépendances,
  lancement, dépannage).
- Validé avec un PV de Semestre 2 **synthétique** (fabriqué pour le test,
  puis supprimé) car le vrai PV S2 n'a jamais été fourni — voir la section
  "Limites" plus bas.

---

## Étape 3 — Amélioration du design (1ère passe)

**Besoin :** "améliore un peu le design".

**Réalisé :** zones de dépôt drag & drop avec retour visuel (bordure verte +
nom du fichier), case PDF présentée clairement, bouton avec spinner de
chargement, page de résultat avec icônes (succès, avertissement) et tableau
stylé. Vérifié par capture d'écran (Edge en mode headless).

---

## Étape 4 — Amélioration du design (2e passe, plus poussée)

**Besoin :** "ajoute css le design est trop null" — le rendu de l'étape 3
jugé insuffisant.

**Réalisé :** refonte visuelle complète : police Google Fonts ("Plus
Jakarta Sans"), fond animé à base de formes dégradées floutées ("blobs"),
titre en texte dégradé, carte en verre dépoli (glassmorphism), zones de
dépôt avec badges numérotés (1, 2) et icône dans un cercle, interrupteur
animé (toggle) pour l'option PDF, bouton en dégradé avec icône éclair,
anneau de validation animé sur la page de résultat. Revérifié par capture
d'écran.

---

## Étape 5 — Passage à un seul fichier PV en entrée (web)

**Besoin :** changement de contexte — l'interface web ne doit plus exiger
deux PV (S1 + S2) séparés, mais accepter **un seul fichier** Excel qui
couvre déjà les deux semestres. Exemple fourni : `L1 BIG DATA.xlsx`.

**Constat sur le fichier d'exemple :** une seule feuille (nommée `S1 `, pas
`PV_Deliberation` comme avant), mais elle contient bien les 10 UE des deux
semestres (5 UE "Semestre 1" + 5 UE "Semestre 2"), conformes à
`TEMPLATE_BLOCKS`. Deux écarts de structure ont dû être corrigés :
- Le merge d'en-tête "Semestre 2" (ligne semestre) ne couvrait pas les 2
  dernières UE de ce semestre (Bases de données, Langues et Culture
  Numérique) → la détection du semestre par UE se base maintenant sur le
  dernier en-tête "Semestre X" dont le merge démarre avant (ou à) la
  colonne de l'UE, plutôt que sur une inclusion stricte dans le merge.
- Quelques notes d'examen étaient saisies en texte avec virgule décimale
  (ex. `'16,5'` au lieu de `16.5`) au lieu d'un nombre — ajout d'un helper
  `to_number()` qui tolère ce format avant toute conversion en float/int.

**Réalisé :**
- `generate_releves.py` : `parse_pv()` détecte la feuille automatiquement
  (essaie `PV_Deliberation`, sinon prend la première feuille du classeur) au
  lieu d'exiger un nom de feuille fixe ; détection de semestre par UE
  rendue tolérante aux merges d'en-tête incomplets ; nouveau `to_number()`
  utilisé pour les notes, coefficients et crédits capitalisés.
- `app.py` : route `/generate` n'attend plus qu'un seul champ fichier
  (`pv_file`, obligatoire) au lieu de `pv_s1` + `pv_s2` obligatoires tous
  les deux ; appelle `generate_all([pv_path], ...)`.
- `templates/index.html` + `static/style.css` : une seule zone de dépôt
  (au lieu de deux), texte d'accroche mis à jour, bouton activé dès qu'un
  fichier est choisi.
- `README.md` : sections usage web/CLI et "Comment ça marche" mises à jour
  en conséquence. La CLI garde la possibilité de fournir 2 PV séparés
  (fusion par étudiant) — seule l'interface web est limitée à un fichier.
- Validé de bout en bout avec `L1 BIG DATA.xlsx` (14 étudiants, aucun
  avertissement, deux semestres bien remplis dans le relevé généré), via
  `generate_all()` directement puis via une requête `POST /generate` sur
  l'app Flask (test client).

**Notes :**
- `L1 BIG DATA.xlsx` ne suit pas la convention de nom `AU24-25-...` utilisée
  pour déduire l'année universitaire depuis le nom de fichier → "Année
  universitaire" reste vide dans le relevé généré pour ce fichier (aucune
  info d'année dans le contenu du PV non plus). À surveiller si ça doit
  être saisi autrement (champ manuel dans l'interface, ou renommage du
  fichier avant import).
- La prise en charge CLI de 2 PV séparés (logique de fusion `merge_students`)
  est conservée intacte ; seule l'UI web a changé.

---

## Étape 6 — Multi-filières, relevé dynamique, coordonnées, Attestation de Réussite

**Besoin :** l'établissement gère 5 filières (Licence GLSI, Licence Big
Data, Licence Réseaux & Systèmes, Mastère Cloud, Mastère Cybersécurité),
chacune avec un nombre d'UE/ECUE différent et plusieurs niveaux (L1/L2/L3,
M1/M2) — le gabarit à structure fixe (`TEMPLATE_BLOCKS`, calé sur L1 Big
Data) ne pouvait gérer que ce seul cas. Demandé : (1) page d'accueil avec
menu déroulant pour choisir la filière avant de déposer le PV, (2) un
relevé qui s'adapte à n'importe quelle structure UE/ECUE (y compris des
noms de matière vides), (3) compléter automatiquement date/lieu de
naissance et CIN à partir d'un nouveau fichier de coordonnées des étudiants
(`قائمـة الطلبـة المسجليـن 25-26.xlsx`, un onglet par filière+niveau), (4)
générer en plus une Attestation de Réussite Word par étudiant (gabarit
`AR.docx`, sans champs de fusion).

**Décisions :**
- Parcours en 2 écrans : `/` (choix filière) → `/upload` (dépôt PV +
  coordonnées) → `/generate`.
- Fichier de coordonnées déposé à chaque génération (pas un fichier fixe du
  projet), car mis à jour chaque année.
- L'Attestation de Réussite est générée pour **tous** les étudiants du lot
  (mention et date de signature laissées en remplissage manuel, comme la
  "Décision du jury" dans le relevé) — pas de moyen fiable de déterminer
  automatiquement qui a réussi.

**Réalisé :**
- `generate_releves.py` : remplacement de `TEMPLATE_BLOCKS` /
  `build_template_row_map` par un moteur dynamique
  (`build_row_layout` + `_write_dynamic_ue_block`) qui redimensionne le
  bloc UE/ECUE du gabarit (`insert_rows`/`delete_rows` autour de la ligne
  40 = ligne "TOTAL" d'origine) pour coller exactement au nombre d'UE/ECUE
  du PV, quelle que soit la filière. Plus de validation stricte du nombre
  d'UE/ECUE attendu — un PV avec une UE ou un nom de matière vide ne fait
  plus planter la génération.
  - **Piège rencontré :** openpyxl ne décale **pas** les plages fusionnées
    (`merged_cells.ranges`) lors de `insert_rows`/`delete_rows` (seules les
    valeurs/styles de cellules le sont) — il a fallu démonter toutes les
    fusions du pied de page et de la bordure décorative colonne O
    (`O17:O39`, qui traverse la frontière du bloc) *avant* le
    redimensionnement, et les recréer à la main aux nouvelles coordonnées
    après. Validé avec un bloc réduit (6 lignes) et un bloc agrandi (25
    lignes) : fusions et pied de page atterrissent au bon endroit dans les
    deux cas, et le cas L1 Big Data (21 lignes, `diff=0`) reste identique à
    avant (0 régression).
  - Ajout de `FILIERES` (libellé → code GLSI/BD/RSYS/CLOUD/CYBER, sert à
    retrouver l'onglet du fichier de coordonnées) et `NIVEAU_TEXTE`
    (L1→("1","ère"), L2→("2","ème"), etc., pour l'attestation).
  - Ajout de `parse_student_coordinates()` / `find_coordinates_sheet()` /
    `normalize_name()` : croise les étudiants du PV avec l'onglet
    niveau+filière du fichier de coordonnées (recherche par nom normalisé,
    insensible accents/casse/espaces). Étudiant non trouvé → avertissement,
    champs laissés vides (même mécanisme que les avertissements existants).
  - `fill_releve()` remplit désormais D10 (naissance) / D11 (CIN) quand
    disponibles (avant : toujours vides).
- `generate_attestation.py` (nouveau, dépend de `python-docx`) : remplit
  `AR.docx` en repérant les paragraphes par leur libellé et en complétant
  les runs vides, ou en remplaçant le texte d'un run existant (ex.
  "Tunisienne" → nationalité réelle, "1" + "ère" → niveau réel en
  préservant la mise en forme superscript du suffixe ordinal).
- `app.py` : nouvelle route `/` (choix filière), `/upload` (dépôt PV +
  coordonnées, filière en champ caché), `/generate` étendu (coordonnées +
  attestations, colonne supplémentaire dans le résultat).
- `templates/index.html` devient la page d'accueil (menu déroulant) ;
  nouveau `templates/upload.html` (ex-contenu de `index.html`, + 2e zone de
  dépôt optionnelle pour les coordonnées) ; `result.html` : colonne
  "Attestation".
- `requirements.txt` : ajout de `python-docx`.
- Validé de bout en bout : CLI sans coordonnées/filière (rétrocompatible,
  identique à l'étape 5), et parcours web complet
  (`/` → `/upload` → `/generate`) avec `L1 BIG DATA.xlsx` +
  `قائمـة الطلبـة المسجليـن 25-26.xlsx` via le client de test Flask — 14
  relevés + 14 attestations générés, coordonnées correctement croisées
  (ex. Abdessalem Fady : naissance "10/06/2006 à Tunis", CIN "14553283"),
  aucun avertissement.

**Notes / limites :**
- La date de signature de l'attestation ("TUNIS, le ...") n'est pas
  modifiée par le programme — reprise telle quelle du gabarit `AR.docx`,
  remplissage manuel comme la mention.
- `generate_attestation` est importé tardivement dans `generate_all()`
  (seulement si `ar_template_path` est fourni) pour que `python-docx` reste
  une dépendance optionnelle côté CLI.

## Étape 7 — Niveau choisi dans l'interface (au lieu d'être déduit du nom du PV)

**Besoin :** le croisement avec le fichier de coordonnées ne marchait pas
de façon fiable car le niveau (L1/L2/L3/M1/M2) était déduit du **nom du
fichier PV** (`parse_filename_meta`) — si le PV n'était pas nommé selon la
convention attendue, aucune correspondance n'était trouvée, donc naissance
et CIN restaient vides. Demandé : ajouter un 2e menu déroulant "Niveau" sur
la page d'accueil, avec 3 choix pour une Licence (1ère/2ème/3ème année) et
2 pour un Mastère (1ère/2ème année), peuplé dynamiquement selon la filière
choisie.

**Vérifié au passage :** le mapping filière → code (`GLSI`/`BD`/`RSYS`/
`CLOUD`/`CYBER`) utilisé pour retrouver l'onglet est correct sur les 11
onglets du fichier de coordonnées (testé niveau × code pour les 25
combinaisons possibles) ; les combinaisons absentes (ex. L3-BD, L3-RSYS,
toute Licence en Cloud/Cyber) renvoient bien "aucun onglet trouvé" au lieu
de planter.

**Réalisé :**
- `generate_releves.py` : `generate_all()` accepte maintenant un paramètre
  `niveau` explicite (prioritaire sur la détection via nom de fichier, qui
  reste le comportement par défaut pour la CLI). Ajout de
  `NIVEAUX_PAR_CYCLE` (Licence → L1/L2/L3, Mastère → M1/M2) et
  `cycle_de_filiere()`.
- `app.py` : `/` transmet `niveaux_par_cycle` et le cycle de chaque
  filière au template ; `/upload` et `/generate` valident `niveau` contre
  le cycle de la filière choisie (`_niveaux_valides()`) ; `niveau` propagé
  jusqu'à `generate_all()`.
- `templates/index.html` : 2e `<select>` "Niveau", désactivé tant qu'aucune
  filière n'est choisie, repeuplé en JS (`niveauxParCycle`, injecté via
  `tojson`) selon l'attribut `data-cycle` de la filière sélectionnée.
- `templates/upload.html` : niveau ajouté au fil d'Ariane et en champ
  caché du formulaire.
- Validé : génération avec un PV **renommé sans aucun indice de niveau**
  dans son nom, niveau "L1" choisi explicitement dans le formulaire →
  naissance/CIN correctement remplis (preuve que ça ne dépend plus du nom
  du fichier).

---

## État actuel des règles métier (à respecter si on reprend le projet)

- Sortie : un `.xlsx` (+ `.pdf` en option) + une attestation `.docx` par
  étudiant, zip téléchargeable par lot de génération.
- "Décision du jury" / mention : toujours vide, remplissage manuel (relevé
  et attestation).
- Date/lieu de naissance, CIN/Passeport : complétés automatiquement si un
  fichier de coordonnées est fourni et qu'une correspondance de nom est
  trouvée ; sinon vides (remplissage manuel).
- Nom & Prénom : jamais réordonné, recopié tel quel du PV.
- Validation d'UE (V/C) : crédits capitalisés > 0 → "V" et "C", sinon vide.
- Bloc UE/ECUE du relevé entièrement dynamique : s'adapte au nombre réel
  d'UE/ECUE du PV, donc fonctionne pour les 5 filières de l'établissement
  sans gabarit dédié par filière.
- Interface web : un seul PV en entrée (couvrant idéalement les deux
  semestres) + filière choisie au préalable. La fusion de 2 PV séparés par
  étudiant n'existe plus que côté CLI ; si un·e étudiant·e n'apparaît que
  dans un des PV fournis en CLI, son relevé est quand même généré avec les
  données disponibles, et un avertissement est affiché.
- Notes/coefficients/crédits saisis en texte avec virgule décimale (ex.
  `"16,5"`) sont tolérés et convertis automatiquement (`to_number()`).
- Moyenne annuelle et total des crédits : recalculés à partir des UE
  réellement présentes (pondération par coefficient d'UE), pas une simple
  recopie des colonnes du PV.

## Limites connues / dette technique

- Pas de tests automatisés (pytest non installé) — validation faite
  manuellement à chaque étape.
- Export PDF dépendant de Microsoft Excel installé localement (Windows
  uniquement, pas de fallback LibreOffice).
- Seule la filière Big Data a été testée de bout en bout avec de vraies
  données (PV + fichier de coordonnées). Les 4 autres filières (GLSI,
  RSYS, Cloud, Cybersécurité) n'ont pas encore été essayées avec un vrai
  PV — le moteur dynamique a été validé avec des structures UE/ECUE
  synthétiques plus petites/plus grandes que l'exemple Big Data, mais pas
  avec leurs PV réels.
- Le niveau (L1/L2/L3/M1/M2) est choisi explicitement dans l'interface web
  (étape 7) ; côté CLI sans argument explicite, il reste déduit du nom du
  fichier PV (`parse_filename_meta`).
