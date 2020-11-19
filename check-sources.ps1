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
        # On supprime juste l'élément du tableau et on ajoute le nouveau (on met à jour en fait)
        $sourceStatusList = [Array]($sourceStatusList | Where-Object { 
            ($_.name -ne $sourceStatus.name) -and ($_.sourceType -eq $sourceType.ToString())})

        # Si c'est égal à $null, c'est qu'il n'y a plus aucun élément donc on recréé un tableau vide
        if($null -eq $sourceStatusList)
        {
            $sourceStatusList = @()
        }
    }

    # Ajout de l'élément qu'on a supprimé
    $sourceStatusList += $newSourceStatus

    return $sourceStatusList
}



<#
 ------------------------------------------------------------------------------------------------
 --------------------------------- PROGRAMME PRINCIPAL ------------------------------------------
#>


$logName = "sources"
$logHistory = [LogHistory]::new($logName, (Join-Path $PSScriptRoot "logs"), 30)


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


# -------------------------------
# Et on démarre
$logHistory.addLineAndDisplay("Début de la surveillance...")
# Boucle infinie
While ($true)
{
    Foreach ($source in $webSources)
    {
        $logHistory.addLine("Contrôle de $($source.name)...")

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

        # Si on a pu trouver un résultat
        if($regexSearch.Success)
        {
            $sourceDate = $regexSearch.Groups[$regexSearch.Groups.count-1].Value.Trim()

            if($sourceDate -eq "")
            {
                $logHistory.addWarningAndDisplay("La recherche de la date sur la page web n'a rien donné, veuillez contrôler la valeur de 'dateRegex' pour la source courante dans le fichier JSON")
                continue
            }
            
            # Contrôle si la source a changé
            $sourceStatusList = checkIfChanged -sourceStatusList $sourceStatusList -source $source -sourceDate $sourceDate -sourceType ([sourceType]::web)
            # Mise à jour du fichier
            $sourceStatusList | ConvertTo-Json | Out-file $sourcesStatusFile -Encoding:utf8
        }
        else # Pas pu trouver de date
        {
            $logHistory.addWarningAndDisplay("Pas possible de trouver la date de mise à jour, veuillez contrôler la valeur de 'dateRegex' pour la source courante dans le fichier JSON de configuration")
        }

    }# FIN BOUCLE de parcours des sources

    $sleepMin = 10
    $logHistory.addLine("Toutes les sources ont été contrôlées... attente de $($sleepMin) minutes jusqu'au prochain contrôle...")
    Start-Sleep -Seconds (60 * $sleepMin)
}
