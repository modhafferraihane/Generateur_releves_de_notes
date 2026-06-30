# Générateur de relevés de notes

Ce programme crée automatiquement les relevés de notes et les attestations
de réussite des étudiants, à partir du PV de délibération (le fichier Excel
avec toutes les notes). Plus besoin de les remplir un par un à la main.

## Ce qu'il vous faut avant de commencer

- Un ordinateur Windows.
- Le PV de délibération (fichier Excel).
- (Optionnel) Microsoft Excel installé, si vous voulez aussi des PDF prêts
  à imprimer.

## Installation (à faire une seule fois)

### Étape 1 — Installer Python

Ouvrez le menu Démarrer, tapez **PowerShell**, et ouvrez-le.

Copiez-collez cette ligne, puis appuyez sur Entrée :

```powershell
winget install -e --id Python.Python.3.12
```

Une fois l'installation terminée, **fermez puis rouvrez PowerShell**.

### Étape 2 — Installer le programme

Dans PowerShell, déplacez-vous dans le dossier du programme (remplacez le
chemin par celui où se trouve le dossier sur votre ordinateur) :

```powershell
cd "D:\relevenote\Generateur_releves_de_notes"
```

Puis installez tout ce dont le programme a besoin avec cette seule
commande :

```powershell
python -m pip install -r requirements.txt
```

Cette installation est à faire une seule fois.

## Utilisation

1. Toujours dans PowerShell (dans le dossier du programme), tapez :

   ```powershell
   python app.py
   ```

2. Ouvrez votre navigateur internet et allez sur l'adresse :
   **http://127.0.0.1:5000**

3. Choisissez la filière, puis le niveau (1ère, 2ème, 3ème année...).

4. Déposez le PV de délibération. Si vous avez aussi le fichier des
   coordonnées des étudiants, déposez-le également : il sert à compléter
   automatiquement la date de naissance et le numéro de CIN.

5. Cochez la case "Générer aussi les PDF" si vous voulez des versions
   prêtes à imprimer.

6. Cliquez sur "Générer les relevés".

7. Téléchargez les fichiers un par un, ou tous d'un coup avec le bouton
   "Télécharger tout (ZIP)".

Pour arrêter le programme : retournez dans la fenêtre PowerShell et
appuyez sur **Ctrl + C**.

## À savoir

- La mention du jury (Admis, Assez Bien...) n'est pas remplie
  automatiquement : à ajouter vous-même après génération.
- La date de naissance et le CIN ne sont remplis que si vous fournissez le
  fichier des coordonnées des étudiants.

## En cas de problème

- **Le programme ne se lance pas / "Python was not found"** : refaites
  l'étape 1 (installation de Python), puis fermez et rouvrez PowerShell.
- **L'export PDF échoue** : vérifiez que Microsoft Excel n'a pas de fenêtre
  ouverte avec un message en attente, fermez-la et réessayez.
