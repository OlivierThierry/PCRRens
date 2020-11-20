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

# Sources prises en charge
Deux types de sources sont actuellement prises en charge par le script, il s'agit des suivantes:
- **Web** pour pouvoir contrôler si une page web a changé, en se basant sur le contenu de celle-ci
- **File** pour surveiller quand un fichier (sur le réseau en général) a été modifié

# Fonctionnement 
Le script va contrôler périodiquement les différentes sources configurées pour trouver leur dernière date de modification. 

A chaque fois qu'une nouvelle version d'une source est identifiée, l'ordinateur va informer de manière sonore notifier de la chose. En plus de ça, un lien sur la source qui a changé va être affiché à l'écran.

Il est possible de copier le lien en sélectionnant celui-ci et en cliquant sur le bouton droit de la souris. Il suffit ensuite d'aller dans un navigateur Web et de copier le lien pour arriver sur la page web concernée.

La possibilité de mettre une liste d'actions à effectuer lorsqu'une source change est disponible via les données présentes dans le fichier JSON du type de la source. Il s'agit d'un tableau avec la liste (descriptions textuelles) des choses à faire.

En plus des opérations à effectuer manuellement, il est possible d'ajouter un certain degré d'automatisation, plus particulièrement pour les pages web, afin d'extraire les données nécessaires dans celles-ci et de les mettre dans un fichier CSV.


# Ajouter une nouvelle source de données

## Web
Le descriptif des sources de données **web** se trouvent dans le fichier `data/web-sources.json`. Le format du fichier est le suivant:
```
[
    {
        "name": "...",
        "location": "...",
        "searchFilters": [
            {
                "attribute": "...",
                "operator": "...",
                "value": "..."
            },
            {
                "attribute": "...",
                "operator": "...",
                "value": "..."
            }
        ],
        "dateRegex": "...",
        "actions": [
            "...",
            "..."
        ],
        "textToSpeech": "...",
        "callbackFunc": "..."
    },
    ...
```
Voici un descriptif des différents champs intervenant dans la description d'une source web:
**name** > Nom de la source, doit être unique pour toutes les sources Web.
**location** > URL de la page web à surveiller.
**searchFilter** > Liste des filtres de recherche qui sont appliquer en mode *ET* pour trouver un noeud HTML. 
   *attribute* > Nom de l'attribut du noeud que l'on cherche (ça peut être le type du noeud `tagName`, une classe appliquée `class` ou le texte contenu `innerText`)
   *operator* > Opérateur de comparaison PowerShell à utiliser (ex: `-eq`, `-like`, ...)
   *value* > Valeur à contrôler via *operator* pour valider que la condition est remplie (ça peut être le nom du noeud cherché ou la classe appliquée)
**dateRegex** > Expression régulière à utiliser pour extraire la date dans le texte du noeud qui a été recherché via les valeurs de **searchFilter**. :warning: à faire en sorte de mettre entre parenthèses l'expression régulière qui fera ressortir la date.
**actions** > Tableau avec la liste des actions à effectuer, dans l'ordre. Pas besoin de se préoccuper d'une éventuelle numérotation, celle-ci sera ajoutée automatiquement
**textToSpeech** > Texte à lire dans les haut-parleurs lorsque la source est mise à jour. :warning: le texte doit être en anglais, car pas d'autre langue installée pour la fonctionnalité "Text-to-Speech"... 
**callbackFunc** > Fonction à appeler lorsque la page web surveillée change. Cela permet par exemple d'extraire des informations de la page et de mettre celles-ci dans un fichier CSV. La fonction dont le nom est mis ici devra être présente dans la classe `WebCallbackFunc` ([fichier de la classe](https://github.com/LuluTchab/PCRRens/blob/main/include/CallbackFunc.inc.ps1))

## File
Le descriptif des sources de données **file** se trouvent dans le fichier `data/file-sources.json`. Le format du fichier est le suivant:
```
[
    {
        "name": "...",
        "location": "...",
        "textToSpeech": "..."
    },
    ...
]
```
Voici un descriptif des différents champs intervenant dans la description d'une source file:
**name** > Nom de la source, doit être unique parmi toutes les sources **file**
**location** > Chemin jusqu'au fichier à surveiller. :warning tous les backslash `\` qui se trouvent dans le chemin d'accès doivent être doublés (`\\`), ceci est une "contrainte" de l'utilisation de fichiers JSON
**textToSpeech** > Texte à lire dans les haut-parleurs lorsque la source est mise à jour. :warning: le texte doit être en anglais, car pas d'autre langue installée pour la fonctionnalité "Text-to-Speech"... 
