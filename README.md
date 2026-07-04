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

Ouvrez le menu Démarrer, tapez **PowerShell**, et ouvrez-le.

Copiez-collez cette ligne, puis appuyez sur Entrée :

```powershell
irm https://raw.githubusercontent.com/modhafferraihane/Generateur_releves_de_notes/main/install.ps1 | iex
```

C'est tout : le programme s'installe tout seul (y compris Python si besoin)
et se lance automatiquement dans votre navigateur à la fin. Une icône
**"Generateur de releves de notes"** est créée sur le Bureau pour le
relancer facilement la prochaine fois.

> Si une fenêtre Windows demande une autorisation, acceptez-la.
> L'installation peut prendre quelques minutes la première fois.

## Utilisation

1. Double-cliquez sur l'icône **"Generateur de releves de notes"** sur le
   Bureau (ou relancez la commande d'installation ci-dessus).

2. Le site s'ouvre automatiquement dans votre navigateur sur
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

Pour arrêter le programme : fermez la fenêtre noire qui s'est ouverte avec
lui.

## À savoir

- La mention du jury (Admis, Assez Bien...) n'est pas remplie
  automatiquement : à ajouter vous-même après génération.
- La date de naissance et le CIN ne sont remplis que si vous fournissez le
  fichier des coordonnées des étudiants.

## En cas de problème

- **L'installation affiche une erreur la première fois** : fermez
  PowerShell, rouvrez-le, et relancez la commande d'installation une
  deuxième fois (Python a parfois besoin d'un redémarrage de la fenêtre
  pour être détecté juste après son installation).
- **L'export PDF échoue** : vérifiez que Microsoft Excel n'a pas de fenêtre
  ouverte avec un message en attente, fermez-la et réessayez.

## Autre façon d'installer : avec Docker (optionnel)

Cette méthode est une alternative à l'installation ci-dessus. Elle est
utile si vous êtes sur Mac ou Linux, ou si vous préférez ne pas installer
Python directement sur votre ordinateur.

> ⚠️ Avec Docker, la case "Générer aussi les PDF" ne fonctionnera pas : cette
> option a besoin de Microsoft Excel installé sur Windows. Si vous avez
> besoin des PDF, utilisez l'installation classique ci-dessus.

### Ce qu'il vous faut

- [Docker Desktop](https://www.docker.com/products/docker-desktop/) installé
  et lancé.
- Le fichier modèle fourni par votre établissement (son nom commence par
  `Exemple` et finit par `.xlsx`), et éventuellement le fichier
  d'attestation (nom commençant par `AR` et finissant par `.docx`).

### Installation et démarrage

1. Téléchargez ce projet : bouton vert **Code** puis **Download ZIP** sur la
   page GitHub du projet, et dézippez-le où vous voulez.
2. Dans le dossier obtenu, créez un nouveau dossier nommé **modeles**.
3. Mettez-y votre fichier `Exemple ... .xlsx` (et `AR ... .docx` si vous en
   avez un).
4. Ouvrez un terminal dans ce dossier (sous Windows : clic droit dans le
   dossier > **Ouvrir dans le terminal**) et lancez :

   ```
   docker compose up -d --build
   ```

   La toute première fois, cela peut prendre quelques minutes (Docker
   télécharge et prépare tout).
5. Ouvrez votre navigateur sur **http://127.0.0.1:5000** : le site
   fonctionne exactement comme dans le reste de ce README.

   Pour être sûr·e que c'est bien la version Docker qui s'affiche (et pas
   une autre copie du programme déjà lancée sur votre ordinateur), regardez
   en haut de la page : un badge **🐳 Docker** apparaît à côté de "100%
   automatique" uniquement dans la version conteneurisée.

### Arrêter / relancer

- Pour arrêter : `docker compose down`
- Pour relancer plus tard : `docker compose up -d` (sans `--build`, c'est
  plus rapide ; utilisez `--build` uniquement si vous avez retéléchargé une
  nouvelle version du projet)
- Si le site ne démarre pas ou que le badge **🐳 Docker** n'apparaît pas :
  une autre copie du programme (installation classique) tourne peut-être
  déjà sur le port 5000. Fermez-la (voir section "Utilisation" ci-dessus)
  avant de relancer `docker compose up -d`.

### À savoir

- Les fichiers générés restent disponibles tant que vous ne supprimez pas
  le conteneur : utilisez toujours les boutons de téléchargement du site
  pour les récupérer sur votre ordinateur.
- Si vous mettez à jour votre fichier modèle dans le dossier `modeles`,
  relancez avec `docker compose up -d` pour que le changement soit pris en
  compte.
