# Documentation: http://woshub.com/parsing-html-webpages-with-powershell/
# https://www.pipehow.tech/invoke-webscrape/

. ([IO.Path]::Combine("$PSScriptRoot", "include", "func.inc.ps1"))
. ([IO.Path]::Combine("$PSScriptRoot", "include", "LogHistory.inc.ps1"))


Enum sourceType {
    web
    file
}
<#
 ------------------------------------------------------------------------------------------------
 --------------------------------- FONCTIONS ------------------------------------------
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

    RET : Objet $sourceStatusList mis à jour
#>
function checkIfChanged([Array]$sourceStatusList, [PSObject]$source, [string]$sourceDate, [sourceType]$sourceType)
{
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
        
        Write-Host ("{0} - " -f (Get-Date -format "yyyy-MM-dd HH:mm:ss")) -NoNewline -ForegroundColor:Green
        Write-host "'$($source.name)' mis à jour! => $($sourceDate)`n$($source.location)"

        if($source.actions.Count -gt 0)
        {
            # Affichage des actions à entreprendre
            Write-Host "Actions à entreprendre:" -BackgroundColor:DarkGray 
            $stepNo = 1
            $source.actions | ForEach-Object {
                Write-Host "$($stepNo): $($_)"
                $stepNo++
            }
        }
        
        Write-Host ""

        # On fait une petite alerte sonore pour notifier de la mise à jour
        soundAlert
    }

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

    return $sourceStatusList
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
    $response = Invoke-WebRequest -uri $source.location

    $checkIn = "AllElements"
    $conditions = @()
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
        $logHistory.addWarningAndDisplay("Plus d'un élément pouvant contenir l'information de trouvé. Veuillez contrôler la valeur de 'searchFilters'  dans le fichier ($($webSourcesFile)) pour ajouter plus de précision à la recherche")
        continue
    }

    # Tentative d'extraire la date de mise à jour
    $regexSearch = [Regex]::Match($result.innerText, $source.dateRegex)

    # Si on a pu trouver la date de mise à jour
    if($regexSearch.Success)
    {
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


$logName = "sources"
$logHistory = [LogHistory]::new($logName, (Join-Path $PSScriptRoot "logs"), 30)

# Pour contenir toutes les sources
$allSources = @{}

# -------------------------------
# Chargement des sources de données
$webSourcesFile = ([IO.Path]::Combine("$PSScriptRoot", "data", "web-sources.json"))
#$webSourcesFile = ([IO.Path]::Combine("$PSScriptRoot", "data", "test-sources.json"))
$logHistory.addLine("Chargement des sources de données 'web' depuis $($webSourcesFile)")
$webSources = Get-Content -Raw -Path $webSourcesFile | ConvertFrom-Json

$fileSourcesFile = ([IO.Path]::Combine("$PSScriptRoot", "data", "file-sources.json"))
#$webSourcesFile = ([IO.Path]::Combine("$PSScriptRoot", "data", "test-sources.json"))
$logHistory.addLine("Chargement des sources de données 'file' depuis $($fileSourcesFile)")
$fileSources = Get-Content -Raw -Path $fileSourcesFile | ConvertFrom-Json

# Si aucune source
if(($webSources.count -eq 0) -and ($fileSources.count -eq 0))
{
    $logHistory.addErrorAndDisplay("Aucune source de données trouvée dans $($webSourcesFile) ou $($fileSourcesFile), veuillez en ajouter au moins une")
    exit
}

# Ajout de la source
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
else
{
    $logHistory.addLine("Fichier avec les dernières dates des sources trouvé, chargement...")
    $sourceStatusList = [Array](Get-Content -Raw -Path $sourcesStatusFile | ConvertFrom-JSON )

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
                # Sources de type "page web"
                web 
                {
                    $sourceDate = handleWebSource -source $source
                }

                # Sources de types "fichier"
                file
                {
                    $sourceDate = handleFileSource -source $source
                }
            }

            # Si on a pu trouver une date pour la source
            if($null -ne $sourceDate)
            {
                # Contrôle si la source a changé
                $sourceStatusList = checkIfChanged -sourceStatusList $sourceStatusList -source $source -sourceDate $sourceDate -sourceType $sourceTypeEnum
                # Mise à jour du fichier
                $sourceStatusList | ConvertTo-Json | Out-file $sourcesStatusFile -Encoding:utf8
            }


        }# FIN BOUCLE de parcours des sources du type courant

    } # FIN BOUCLE de parcours des types de sources

    $sleepMin = 10
    $logHistory.addLine("Toutes les sources ont été contrôlées... attente de $($sleepMin) minutes jusqu'au prochain contrôle...")
    Start-Sleep -Seconds (60 * $sleepMin)
}
