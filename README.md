# Générateur de relevés de notes

Ce programme crée automatiquement les relevés de notes des étudiants (et,
en dernière année, leur diplôme national), à partir du PV de délibération
(le fichier Excel avec toutes les notes). Plus besoin de les remplir un par
un à la main.

## Ce qu'il vous faut avant de commencer

- Un ordinateur Windows.
- Le PV de délibération (fichier Excel).
- (Optionnel) Microsoft Excel installé, si vous voulez aussi des PDF prêts
  à imprimer.

## Installation (à faire une seule fois)

Ouvrez le menu Démarrer, tapez **PowerShell**, ouvrez-le, puis
copiez-collez cette ligne et appuyez sur Entrée :

```powershell
irm https://raw.githubusercontent.com/modhafferraihane/Generateur_releves_de_notes/main/install.ps1 | iex
```

C'est tout : le programme s'installe tout seul (y compris Python si besoin)
et se lance dans votre navigateur à la fin. Une icône **"Generateur de
releves de notes"** est créée sur le Bureau pour le relancer plus tard.

> Acceptez toute demande d'autorisation Windows. L'installation peut
> prendre quelques minutes la première fois.

## Utilisation

1. Double-cliquez sur l'icône du Bureau (ou relancez la commande
   d'installation ci-dessus).
2. Le site s'ouvre sur **http://127.0.0.1:5000**.
3. Choisissez la filière, puis le niveau.
4. Déposez le PV de délibération (et le fichier des coordonnées des
   étudiants si vous l'avez : il complète la date de naissance, le CIN, et
   le "Mme / M." des diplômes).
5. Vérifiez l'année universitaire et la date de remplissage (déjà
   pré-remplies), puis cochez "Générer aussi les PDF" si besoin.
6. Cliquez sur "Générer les relevés", puis téléchargez les fichiers (un par
   un ou tous d'un coup via "Télécharger tout (ZIP)").

Pour arrêter : fermez la fenêtre noire ouverte avec le programme.

## À savoir

- La mention du jury (Admis, Assez Bien...) n'est pas remplie
  automatiquement sur le relevé : à ajouter vous-même après génération.
- Date de naissance, CIN et civilité (Mme / M., Né / Née) ne sont remplis
  que si le fichier des coordonnées des étudiants est fourni. Cet onglet doit
  correspondre au niveau et à la filière choisis (ex. onglet "L3GLSI").
- **Diplôme (dernière année seulement).** Pour une 3ème année de Licence
  (L3) ou une 2ème année de Mastère (M2), un diplôme national (.docx) est
  aussi généré pour chaque étudiant, en plus du relevé. La mention d'honneur
  est calculée depuis la moyenne annuelle. Le champ "Mention" reprend la
  filière indiquée dans le PV. Pour le Mastère, les moyennes par semestre et
  le nombre de crédits restent à compléter à la main (le PV de 2ème année ne
  couvre pas la 1ère).

## En cas de problème

- **Erreur à la première installation** : fermez PowerShell, rouvrez-le, et
  relancez la commande une deuxième fois (Python a parfois besoin d'un
  redémarrage pour être détecté juste après son installation).
- **L'export PDF échoue** : fermez toute fenêtre Excel avec un message en
  attente, puis réessayez.

## Autre façon d'installer : avec Docker

Alternative utile sur Mac/Linux, ou pour ne pas installer Python
directement sur votre ordinateur. Le conteneur consomme très peu de
ressources, comme on le voit dans Docker Desktop :

![Conteneur tournant avec peu de ressources dans Docker Desktop](pic/dockerDesktop.png)

> ⚠️ Avec Docker, l'export PDF ne fonctionne pas (il nécessite Excel sur
> Windows) : utilisez l'installation classique ci-dessus si vous en avez
> besoin.

Il vous faut seulement [Docker Desktop](https://docs.docker.com/desktop/setup/install/windows-install/)
installé et lancé. Le fichier modèle est déjà inclus dans l'image : rien
d'autre à préparer.

1. Ouvrez un terminal (PowerShell recommandé) et lancez :

   ```
   docker run -d --name generateur-releves -p 127.0.0.1:5000:5000 modovar/generateur-releves:2.0
   ```

2. Ouvrez **http://127.0.0.1:5000** — un badge **🐳 Docker** confirme que
   c'est bien cette version qui tourne.

Pour arrêter/relancer : `docker stop generateur-releves` /
`docker start generateur-releves`.

> **Modèle personnalisé (optionnel).** Pour utiliser *votre* fichier modèle
> (logo/établissement différent) au lieu de celui inclus, mettez-le dans un
> sous-dossier **modeles** et montez-le au lancement :
>
> ```
> docker run -d --name generateur-releves -p 127.0.0.1:5000:5000 -v ./modeles://app/modeles:ro modovar/generateur-releves:2.0
> ```
>
> Le double `//` avant `app/modeles` est volontaire (bug connu de Git Bash
> sinon). Après avoir changé le fichier, faites `docker restart
> generateur-releves`.

Vos fichiers générés ne sont pas conservés ailleurs : téléchargez-les tout
de suite. Supprimer le conteneur (`docker rm`) efface tout, volontairement,
pour ne garder aucune donnée d'étudiant.
