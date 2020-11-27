# Documentation: http://woshub.com/parsing-html-webpages-with-powershell/
# https://www.pipehow.tech/invoke-webscrape/

. ([IO.Path]::Combine("$PSScriptRoot", "include", "func.inc.ps1"))
. ([IO.Path]::Combine("$PSScriptRoot", "include", "WebCallbackFunc.inc.ps1"))
. ([IO.Path]::Combine("$PSScriptRoot", "include", "LogHistory.inc.ps1"))


Enum sourceType {
    web
    file
}

# Dossier où se trouveront les données générées
$global:OUTPUT_FOLDER = ([IO.Path]::Combine("$PSScriptRoot", "extracted-data"))


<#
 ------------------------------------------------------------------------------------------------
 -------------------------------------- FONCTIONS -----------------------------------------------
#>

<#
    ---------------------------------------------------------------------------------------------
    BUT : Contrôle si une source a changé au vu de la date de sa dernière modification

    IN  : $sourceStatusList         -> tableau avec la liste des dernières dates de modification pour les sources.
                                        On doit pouvoir y stocker (et trouver si elle existe) la source $source
    IN  : $source                   -> Objet représentant la source que l'on est en train de contrôler. Provient d'un 
                                        fichier JSON
    IN  : $sourceDate               -> Date trouvée pour la source à l'endroit où on doit la chercher
    IN  : $sourceType               -> Type de la source (web, fichier, etc...)

    RET : Tableau avec:
        [0] $true|$false pour dire si source mise à jour
        [1] $sourceStatusList mis à jour
#>
function checkIfChanged([Array]$sourceStatusList, [PSObject]$source, [string]$sourceDate, [sourceType]$sourceType)
{
    $sourceHasChanged = $false
    $logHistory.addLine("> Dernière mise à jour de la source: $($sourceDate)")
    # Si date présente dans le fichier de log et différente de la courante
    # OU
    # date pas présente dans le fichier log
    $sourceStatus = $sourceStatusList | Where-Object { 
        ($_.name -eq $source.name) -and ($_.sourceType -eq $sourceType.toString())}
    
    if((($null -ne $sourceStatus) -and ($sourceStatus.date -ne $sourceDate) ) `
        -or `
        ($null -eq $sourceStatus))
    {
        $logHistory.addLine("> La source a été mise à jour depuis la dernière vérification")
        
        $sourceHasChanged = $true
        Write-Host ("{0} - " -f (Get-Date -format "yyyy-MM-dd HH:mm:ss")) -NoNewline -BackgroundColor:DarkGreen
        Write-host "'$($source.name)' mis à jour! => $($sourceDate)`n$($source.location)"

        # s'il y a des actions à effectuer
        if($source.actions.Count -gt 0)
        {
            # Affichage des actions à entreprendre
            Write-Host ">> Actions à entreprendre <<"
            $stepNo = 1
            $source.actions | ForEach-Object {
                # Gestion des retours à la ligne pouvant être présents dans l'action
                $actionDesc = ($_ -replace "\\n", "`n")
                Write-Host "$($stepNo): $($actionDesc)"
                $stepNo++
            }

        } # FIN S'IL y a des actions à effectuer

    } # FIN SI la source a changé

    # Pour mettre à jour les infos dans le fichier log
    $newSourceStatus = [PSCustomObject] @{
        name = $source.name
        date = $sourceDate
        lastCheck = (Get-Date).ToString()
        sourceType = $sourceType.ToString()
    }

    # Si on a des informations pour la source dans le fichier log
    if($null -ne $sourceStatus)
    {
        # Suppression de la source courante du tableau si elle s'y trouve
        $updatedSourceStatusList = @()
        $sourceStatusList | Foreach-Object { 
            # Si le nom de la source correspond,
            if($_.name -eq $sourceStatus.name)
            {
                # Si ce n'est pas le même type de source
                if($_.sourceType -ne $sourceType.ToString())
                {
                    # Ce n'est pas la source courante
                    $updatedSourceStatusList += $_
                }
            }
            else # Le nom de la source ne correspond pas,
            {
                # On peut donc l'y ajouter dans tous les cas.
                $updatedSourceStatusList += $_
            }
        }
        $sourceStatusList = $updatedSourceStatusList

    }# FIN Si on a des informations sur la source dans le fichier Log

    # Ajout de l'élément qu'on a supprimé
    $sourceStatusList += $newSourceStatus

    return @($sourceHasChanged, $sourceStatusList)
}


<#
    ---------------------------------------------------------------------------------------------
    BUT : Gère une source web

    IN  : $source                   -> Objet représentant la source que l'on est en train de contrôler. Provient d'un 
                                        fichier JSON

    RET : La date de mise à jour de la source
            $null si pas trouvé ou erreur
#>
function handleWebSource([PSObject]$source)
{
    # Récupération de la page web
    $response = Invoke-WebRequest -uri $source.location

    $checkIn = "AllElements"
    $conditions = @()
    # Création des filtres de recherche pour trouver l'élément HTML contenant l'information sur 
    # la date de mise à jour des données
    $source.searchFilters | Foreach-Object { 
        $conditions += ('$_.{0} {1} "{2}"' -f $_.attribute, $_.operator, $_.value)

        # Si l'élément qu'on cherche est un lien, on fait en sorte de ne chercher que parmis
        # les liens par la suite.
        if(($_.attribute -eq "tagName") -and ($_.value -eq "a"))
        {
            $checkIn = "Links"
        }
    }

    $logHistory.addLine("> Recherche de la date de mise à jour de la page $($source.location)")
    $cmd = '$result = $response.{0} | Where-Object {{ {1} }}' -f $checkIn, ($conditions -join " -and ")
    Invoke-Expression $cmd

    # Si pas de résultat trouvé
    if($null -eq $result)
    {
        $logHistory.addWarningAndDisplay("Pas possible de trouver l'élément contenant l'information dans la page. Veuillez contrôler la valeur de 'searchFilters' dans le fichier de configuration ($($webSourcesFile))")
        continue
    }

    # Si les filtres définis ne sont pas assez précis et que plusieurs résultats sont renvoyés,
    if($result -is [System.Array])
    {
        $logHistory.addWarningAndDisplay("Plus d'un élément pouvant contenir l'information de trouvé. C'est le premier qui a va etre pris.")
        $result = $result[0]
    }

    # Tentative d'extraire la date de mise à jour à l'aide de l'expression régulière donnée dans le fichier de configuration
    $regexSearch = [Regex]::Match($result.innerText, $source.dateRegex)

    # Si on a pu trouver la date de mise à jour
    if($regexSearch.Success)
    {
        # Extraction de la date
        $sourceDate = $regexSearch.Groups[$regexSearch.Groups.count-1].Value.Trim()

        if($sourceDate -eq "")
        {
            $logHistory.addWarningAndDisplay("La recherche de la date sur la page web n'a rien donné, veuillez contrôler la valeur de 'dateRegex' pour la source courante dans le fichier JSON")
            return $null
        }
        
        return $sourceDate
        
    }
    else # Pas pu trouver de date
    {
        $logHistory.addWarningAndDisplay("Pas possible de trouver la date de mise à jour, veuillez contrôler la valeur de 'dateRegex' pour la source courante dans le fichier JSON de configuration")
    }

    return $null
}


<#
    ---------------------------------------------------------------------------------------------
    BUT : Gère une source fichier

    IN  : $source       -> Objet représentant la source que l'on est en train de contrôler. Provient d'un 
                            fichier JSON

    RET : Date de mise à jour du fichier
#>
function handleFileSource([PSObject]$source)
{
    $logHistory.addLine("Contrôle de la date de modification de la source $($source.name) ('$($source.location)')")
    # Si le fichier existe
    if(Test-Path -Path $source.location)
    {
        # Récupération de la date de modification du fichier et renvoi
        return (Get-Item -path $source.location).LastWriteTime.ToString()
    }
    else 
    {
        $logHistory.addWarningAndDisplay("Source $($source.name): Le fichier '$($source.location)' n'existe pas")
    }

    return $null
}


<#
 ------------------------------------------------------------------------------------------------
 --------------------------------- PROGRAMME PRINCIPAL ------------------------------------------
#>

Add-Type –AssemblyName System.Speech
$speechSynthesizer = New-Object –TypeName System.Speech.Synthesis.SpeechSynthesizer

# Objet pour gérer le logging. Des fichiers seront créés au fur et à mesure dans le dossier "logs"
$logName = "sources"
$logHistory = [LogHistory]::new($logName, (Join-Path $PSScriptRoot "logs"), 30)
# Objet contenant les fonctions à appeler lorsqu'une source de donnée change
$webCallbackFunc = [WebCallbackFunc]::new($global:OUTPUT_FOLDER)

# Pour contenir toutes les sources
$allSources = @{}

# -------------------------------
# Chargement des sources de données
$webSourcesFile = ([IO.Path]::Combine("$PSScriptRoot", "data", "web-sources.json"))
#$webSourcesFile = ([IO.Path]::Combine("$PSScriptRoot", "data", "test-sources.json"))
$logHistory.addLine("Chargement des sources de données 'web' depuis $($webSourcesFile)")
$webSources = Get-Content -Raw -Path $webSourcesFile | ConvertFrom-Json

$fileSourcesFile = ([IO.Path]::Combine("$PSScriptRoot", "data", "file-sources.json"))
$logHistory.addLine("Chargement des sources de données 'file' depuis $($fileSourcesFile)")
$fileSources = Get-Content -Raw -Path $fileSourcesFile | ConvertFrom-Json

# Si aucune source
if(($webSources.count -eq 0) -and ($fileSources.count -eq 0))
{
    $logHistory.addErrorAndDisplay("Aucune source de données trouvée dans $($webSourcesFile) ou $($fileSourcesFile), veuillez en ajouter au moins une")
    exit
}

# Ajout des différentes sources dans une autre structure, afin qu'elle puisse être traitée après
$allSources.add(([sourceType]::web).ToString(), $webSources )
$allSources.add(([sourceType]::file).toString(), $fileSources)


# -------------------------------
# Chargement du statut des différentes sources
$sourcesStatusFile = ([IO.Path]::Combine("$PSScriptRoot", "data", "sources-status.json"))
if(!(Test-Path -Path $sourcesStatusFile))
{
    $logHistory.addLine("Fichier avec les dernières dates des sources non trouvé, un nouveau va être créé automatiquement")
    $sourceStatusList = @()
}
else # Un fichier de suivi des modifications a été trouvé, on le charge.
{
    $logHistory.addLine("Fichier avec les dernières dates des sources trouvé, chargement...")
    $sourceStatusList = [Array](Get-Content -Raw -Path $sourcesStatusFile | ConvertFrom-JSON )

    # Affichage des infos trouvées dans le fichier
    Write-Host "Etat des sources" -BackgroundColor:DarkGray
    $sourceStatusList | ForEach-Object {
        Write-Host "$($_.name) ($($_.sourceType)) => $($_.date)"
    }
    Write-Host ""
}



# ----------------------------------------------------------
# Et on démarre
$logHistory.addLineAndDisplay("Début de la surveillance...")
# Boucle infinie
While ($true)
{
    # Parcours des types de sources qu'on a 
    ForEach($sourceType in $allSources.Keys)
    {
        $sourceTypeEnum = [sourceType]$sourceType
        # Parcours des sources pour le type courant
        Foreach ($source in $allSources.$sourceType)
        {
            $logHistory.addLine("Contrôle de $($source.name)...")

            # En fonction du type de source (voir la définition du type enuméré 'sourceType')
            switch($sourceTypeEnum)
            {
                # Source de type "page web"
                web { $sourceDate = handleWebSource -source $source }

                # Source de types "fichier"
                file { $sourceDate = handleFileSource -source $source }

            }# FIN EN FONCTION du type de source

            # Si on a pu trouver une date de modification pour la source
            if($null -ne $sourceDate)
            {
                # Contrôle si la source a changé
                $sourceHasChanged, $sourceStatusList = checkIfChanged -sourceStatusList $sourceStatusList -source $source `
                                                                        -sourceDate $sourceDate -sourceType $sourceTypeEnum
                # Si la source a changé
                if($sourceHasChanged)
                {
                    # Si on est en train de traiter une page web et qu'on a une fonction de callback à appeler
                    if($sourceTypeEnum -eq [sourceType]::web -and $source.callbackFunc -ne "")
                    {
                        # Création de la commande 
                        $cmd = '$outFile = $webCallbackFunc.{0}($source)' -f $source.callbackFunc
                        Invoke-Expression $cmd
                        Write-Host "Un fichier de données a été généré depuis la page web, il peut être trouvé ici:`n$($outFile)"
                        
                    }
                    # On fait une petite alerte sonore pour notifier de la mise à jour
                    soundAlert -speechSynthesizer $speechSynthesizer -message $source.textToSpeech

                    Write-Host ""
                }
                # Mise à jour du fichier
                $sourceStatusList | ConvertTo-Json | Out-file $sourcesStatusFile -Encoding:utf8

            }# FIN SI on a pu trouver une date de modification pour la source

        }# FIN BOUCLE de parcours des sources du type courant

    } # FIN BOUCLE de parcours des types de sources

    $sleepMin = 10
    $logHistory.addLine("Toutes les sources ont été contrôlées... attente de $($sleepMin) minutes jusqu'au prochain contrôle...")
    Start-Sleep -Seconds (60 * $sleepMin)
}
