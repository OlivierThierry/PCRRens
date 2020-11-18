# Description
Permet de contrôler automatiquement (et de manière périodique) si des sources de données ont été mises à jour sur internet.
Si une source est mise à jour, un message sera affiché à l'écran et un signal sonore sera émis.

# Installation et exécution
1. Télécharger le ZIP du code en cliquant sur "*Code*" puis "*Télécharger le ZIP*"
1. Extraire le contenu dans un dossier
1. Aller dans le dossier avec l'explorateur Windows.
1. Appuyer sur "shift" et cliquer avec le bouton droit de la souris dans le dossier. Choisir "*Ouvrir une console PowerShell ici...*"
1. Lancer l'exécution du script
   > ./check-sources.ps1

**Note**

Si c'est la première exécution du script, il va bien évidemment trouver que toutes les sources ont changé.

# Fonctionnement 
Le script va contrôler périodiquement les différentes sources configurées pour trouver leur dernière date de modification. 

A chaque fois qu'une nouvelle version d'une source est identifiée, l'ordinateur va biper quelques fois pour notifier de la chose. En plus de ça, un lien sur la source qui a changé va être affiché à l'écran.

Il est possible de copier le lien en sélectionnant celui-ci et en cliquant sur le bouton droit de la souris. Il suffit ensuite d'aller dans un navigateur Web et de copier le lien pour arriver sur la page web concernée