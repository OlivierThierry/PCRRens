<#
    BUT : Classe contenant les fonctions pour traiter les données d'une page Web (extraction)
            et générer un fichier CSV de données (via la fonction 'writeCSVFile')
#>
class CallbackFunc
{
    hidden [string]$outputFolder

    <#
        -------------------------------------------------------------------------------------
        BUT : Constructeur de classe

        IN  : $outputFolder -> Chemin jusqu'au dossier où créer les fichiers CSV avec les données
    #>
    callbackFunc([string]$outputFolder)
    {
        $this.outputFolder = $outputFolder
    }


    <#
        -------------------------------------------------------------------------------------
        BUT : Génère un fichier CSV avec les infos passées et renvoie le chemin jusqu'à celui-ci'
        
        IN  : $fileId     	-> Id du fichier, sera utilisé pour générer le nom.
        IN  : $csvHeader    -> Tableau avec les 2 colonnes à utiliser pour le Header du CSV
        IN  : $csvData      -> Tableau dont chaque élément est un dictionnaire avec:
                                .name -> colonne 1
                                .value -> colonne 2
        
        RET : Chemin jusqu'au fichier de données
    #>
    hidden [string] writeCSVFile([string]$fileId, [Array]$csvHeader, [Array]$csvData)
    {
         # Génération du nom du fichier de données
        $date =(Get-Date -format "yyyy-MM-dd")
        $csvFile = ([IO.Path]::Combine($this.outputFolder, "$($fileId)_$($date).csv"))

        # Création du fichier de données
        $csvHeader -join ";" | Out-File $csvFile -Encoding:utf8
        $csvData | ForEach-Object {
            "$($_.name);$($_.value)" | Out-File $csvFile -Append -Encoding:utf8
        }

        return $csvFile
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
        # Récupération des <span> qui contiennent les chiffres
        $span14DaysList = $webPage14Days.AllElements | Where-Object { $_.tagName -eq 'span' -and  $_.class -eq "bag-key-value-list__entry-value" }

        $webPageAll = Invoke-WebRequest -uri "$($source.location)?ovTime=total"
        # Récupération des <span> qui contiennent les chiffres
        $spanAllList = $webPageAll.AllElements | Where-Object { $_.tagName -eq 'span' -and  $_.class -eq "bag-key-value-list__entry-value" }

        <# Ce n'est pas particulièrement "propre" de passer par des index pour récupérer des valeurs mais pour le moment aucun
        autre moyen de faire la chose n'a été trouvée...
        FIXME: voir pour améliorer la chose #>
        $fields = @(
            @{
                name = "OFSP - NB nouveaux cas"
                value = $span14DaysList[0].innerText
            },
            @{
                name = "OFSP - Nb de cas pour 100'000 habitants sur 14 derniers jours"
                value = $span14DaysList[2].innerText
            },
            @{
                name = "OFSP - Nb de décès sup."
                value = $span14DaysList[6].innerText
            },
            @{
                name= "OFSP - Nb. hospitalisations"
                value = $spanAllList[4].innerText
            },
            @{
                name = "OFSP - Nb. tests PCR positif sup."
                value = $span14DaysList[9].innerText
            },
            @{
                name = "OFSP - Nb. personnes en quarantaine"
                value = $span14DaysList[14].innerText
            },
            @{
                name = "OFSP - Nb. personnes en isolement"
                value = $span14DaysList[13].innerText
            },
            @{
                name = "OFSP - Nb. personnes en quarantaine (retour)"
                value = $span14DaysList[15].innerText
            }
        )
        
        # Génération du nom du fichier de données
        return $this.writeCSVFile("Covid19-Switzerland", @("Champ", "Valeur"), $fields)
        
    }

}

