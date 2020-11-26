<#
    BUT : Classe contenant les fonctions pour traiter les données d'une page Web (extraction)
            et générer un fichier CSV de données (via la fonction 'writeCSVFile')
#>
class WebCallbackFunc
{
    hidden [string]$outputFolder

    <#
        -------------------------------------------------------------------------------------
        BUT : Constructeur de classe

        IN  : $outputFolder -> Chemin jusqu'au dossier où créer les fichiers CSV avec les données
    #>
    WebCallbackFunc([string]$outputFolder)
    {
        $this.outputFolder = $outputFolder
    }


    <#
        -------------------------------------------------------------------------------------
        BUT : Génère un fichier CSV avec les infos passées et renvoie le chemin jusqu'à celui-ci'
        
        IN  : $fileId     	-> Id du fichier, sera utilisé pour générer le nom.
        IN  : $csvHeader    -> Tableau avec les colonnes à utiliser pour le Header du CSV
        IN  : $csvData      -> Tableau dont chaque élément est en fait une ligne du fichier CSV. 
                                Chaque ligne est représentée par un tableau avec les données, 
                                une case du tableau par colonne du fichier CSV
        
        RET : Chemin jusqu'au fichier de données
    #>
    hidden [string] writeCSVFile([string]$fileId, [Array]$csvHeader, [Array]$csvData)
    {
         # Génération du nom du fichier de données
        $date =(Get-Date -format "yyyy-MM-dd")
        $csvFile = ([IO.Path]::Combine($this.outputFolder, "$($fileId)_$($date).csv"))

        # Création du fichier de données, on commence par l'entête
        $csvHeader -join ";" | Out-File $csvFile -Encoding:utf8
        # Ajout des lignes de données dans le fichier
        $csvData | ForEach-Object {
            $_ -join ";" | Out-File $csvFile -Append -Encoding:utf8
        }

        return $csvFile
    }

    
    <#
        -------------------------------------------------------------------------------------
        BUT : Permet de retrouver une information dans une page.
        
        IN  : $webPage     	-> Objet représentant la page web téléchargée
        IN  : $contentPath  -> Tableau avec les chaînes de caractères à rechercher successivement
                                dans le contenu de la page (string) afin d'arriver au plus près de
                                l'information, juste avant en fait.
        IN  : $strAfter     -> Chaîne de caractère qui délimite la fin de la valeur que l'on cherche
        
        RET : Information recherchée
    #>
    hidden [string]findInPage([PSObject]$webPage, [Array]$contentPath, [string]$strAfter)
    {
        $startPos = 0

        # Avance successive dans la page pour trouver la position de la dernière chaîne
        # définie dans $contentPath et se positionner après
        $contentPath | ForEach-Object {
            $startPos = $webPage.Content.IndexOf($_, $startPos) + $_.length
            # Si pas trouvé
            if($startPos -eq -1)
            {
                return ("$($_) pas trouvé dans la page en cherchant le chemin suivant: {0}" -f ($contentPath -join " >> "))
            }
        }
        # Recherche de la position de la chaîne de caractères qui délimite la fin de la valeur à chercher
        $endPos = $webPage.Content.IndexOf($strAfter, $startPos)
        if($endPos -eq -1)
        {
            return ("$($_) pas trouvé dans la page après avoir cherché le chemin suivant: {0}" -f ($contentPath -join " >> "))
        }
        # Extraction de la chaîne de caractères et retour
        return $webPage.Content.substring($startPos, $endPos-$startPos)
    }


    <#
        -------------------------------------------------------------------------------------
        BUT : Extrait les données qui sont présente sur le site suivant pour créer un fichier 
                de données exploitables directement pour la mise à jour sur les panneaux
            https://www.covid19.admin.ch/en/overview
        
        IN  : $source     	-> Objet représentant la source que l'on traite
        
        RET : Chemin jusqu'au fichier de données créé
    #>
    [string] extractActualSituationCH([PSObject]$source)
    {
        # La page "normale" ne contient que les statistiques pour les 14 derniers jours. 
        $webPage14Days = Invoke-WebRequest -uri $source.location
        # La page "total" contient juste une seule stat dont on a besoin...
        $webPageAll = Invoke-WebRequest -uri "$($source.location)?ovTime=total"

        # Recherche des informations dans la page
        $fields = @(
            @(
                "OFSP - NB nouveaux cas"
                $this.findInPage($webPage14Days, @("Laboratory-confirmed cases", `
                                                    "Difference to previous day", `
                                                    'bag-key-value-list__entry-value">'), "<")
            ),
            @(
                "OFSP - Nb de cas pour 100'000 habitants sur 14 derniers jours"
                $this.findInPage($webPage14Days, @("Laboratory-confirmed cases", `
                                                    "Per 100 000 inhabitants", `
                                                    'bag-key-value-list__entry-value">'), "<")
            ),
            @(
                "OFSP - Nb de décès sup."
                $this.findInPage($webPage14Days, @("Laboratory-confirmed deaths", `
                                                    "Difference to previous day", `
                                                    'bag-key-value-list__entry-value">'), "<")
            ),
            @(
                "OFSP - Nb. hospitalisations"
                $this.findInPage($webPageAll, @("Laboratory-confirmed hospitalisations", `
                                                            "Total since ", `
                                                            'bag-key-value-list__entry-value">'), "<")
            ),
            @(
                "OFSP - Nb. tests PCR positif sup."
                $this.findInPage($webPage14Days, @("Tests and share of positive tests", `
                                                            "Difference to previous day", `
                                                            'bag-key-value-list__entry-value">'), "<")
            ),
            @(
                "OFSP - Nb. personnes en quarantaine"
                $this.findInPage($webPage14Days, @("Contact tracing", `
                                                            "In quarantine", `
                                                            'bag-key-value-list__entry-value">'), "<")
            ),
            @(
                "OFSP - Nb. personnes en isolement"
                $this.findInPage($webPage14Days, @("Contact tracing", `
                                                            "In isolation", `
                                                            'bag-key-value-list__entry-value">'), "<")
            ),
            @(
                "OFSP - Nb. personnes en quarantaine (retour)"
                $this.findInPage($webPage14Days, @("Contact tracing", `
                                                            "Additionally in quarantine", `
                                                            'bag-key-value-list__entry-value">'), "<")
            )
        )
        
        # Génération du nom du fichier de données
        return $this.writeCSVFile("Covid19-Switzerland", @("Champ", "Valeur"), $fields)
        
    }

}

