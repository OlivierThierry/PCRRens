# Documentation: http://woshub.com/parsing-html-webpages-with-powershell/
# https://www.pipehow.tech/invoke-webscrape/

. ([IO.Path]::Combine("$PSScriptRoot", "include", "func.inc.ps1"))
. ([IO.Path]::Combine("$PSScriptRoot", "include", "LogHistory.inc.ps1"))


$logName = "sources"
$logHistory = [LogHistory]::new($logName, (Join-Path $PSScriptRoot "logs"), 30)

$sourcesFile = ([IO.Path]::Combine("$PSScriptRoot", "data", "sources.json"))
#$sourcesFile = ([IO.Path]::Combine("$PSScriptRoot", "data", "test-sources.json"))
$logHistory.addLine("Loading sources from $($sourcesFile)")
$sources = Get-Content -Raw -Path $sourcesFile | ConvertFrom-Json

if($sources.count -eq 0)
{
    $logHistory.addErrorAndDisplay("No source found in $($sourcesFile), please add at least one source")
    exit
}

$lastSourceDateListFile = ([IO.Path]::Combine("$PSScriptRoot", "lastSourceDateList.json"))

if(!(Test-Path -Path $lastSourceDateListFile))
{
    $logHistory.addLine("Last sources date file doesn't exists, a new one will automatically be created")
    $lastSourceDateList = @()
}
else
{
    $logHistory.addLine("Last sources date file found, loading content...")
    $lastSourceDateList = [Array](Get-Content -Raw -Path $lastSourceDateListFile | ConvertFrom-JSON)
}



# Boucle infinie
While ($true)
{

    Foreach ($source in $sources)
    {
        $logHistory.addLine("Contrôle de $($source.name)...")

        $response = Invoke-WebRequest -uri $source.url

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

        $logHistory.addLine("> Recherche de la date de mise à jour de la page $($source.url)")
        $cmd = '$result = $response.{0} | Where-Object {{ {1} }}' -f $checkIn, ($conditions -join " -and ")
        Invoke-Expression $cmd

        # Si pas de résultat trouvé
        if($null -eq $result)
        {
            $logHistory.addWarningAndDisplay("Pas possible de trouver l'élément contenant l'information dans la page. Veuillez contrôler la valeur de 'searchFilters' dans le fichier de configuration ($($sourcesFile))")
            continue
        }

        # Si les filtres définis ne sont pas assez précis et que plusieurs résultats sont renvoyés,
        if($result -is [System.Array])
        {
            $logHistory.addWarningAndDisplay("Plus d'un élément pouvant contenir l'information de trouvé. Veuillez contrôler la valeur de 'searchFilters'  dans le fichier ($($sourcesFile)) pour ajouter plus de précision à la recherche")
            continue
        }

        # Tentative d'extraire la date de mise à jour
        $regexSearch = [Regex]::Match($result.innerText, $source.dateRegex)

        # Si on a pu trouver un résultat
        if($regexSearch.Success)
        {
            $webSourceDate = $regexSearch.Groups[$regexSearch.Groups.count-1].Value.Trim()

            if($webSourceDate -eq "")
            {
                $logHistory.addWarningAndDisplay("La recherche de la date sur la page web n'a rien donné, veuillez contrôler la valeur de 'dateRegex' pour la source courante dans le fichier JSON")
                continue
            }

            $logHistory.addLine("> Dernière mise à jour de la page: $($webSourceDate)")
            # Si date présente dans le fichier de log et différente de la courante
            # OU
            # date pas présente dans le fichier log
            $sourceInfos = $lastSourceDateList | Where-Object { $_.name -eq $source.name}
            if((($null -ne $sourceInfos) -and ($sourceInfos.date -ne $webSourceDate) ) `
                -or `
                ($null -eq $sourceInfos))
            {
                $logHistory.addLine("> La page a été mise à jour depuis la dernière vérification")
                
                
                Write-Host ("{0} - " -f (Get-Date -format "yyyy-MM-dd HH:mm:ss")) -NoNewline -ForegroundColor:Green
                Write-host "'$($source.name)' mis à jour! => $($webSourceDate)`n$($source.url)"

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
            $newSourceInfos = @{
                name = $source.name
                date = $webSourceDate
                lastCheck = (Get-Date).ToString()
            }
            # Si on a des informations pour la source dans le fichier log
            if($null -ne $sourceInfos)
            {
                # On supprime juste l'élément du tableau
                $lastSourceDateList = ($lastSourceDateList | Where-Object { $_.name -ne $source.name})
                # Si plus aucun élément,
                if($null -eq $lastSourceDateList)
                {
                    # On est obligé de refaire en sorte que ça soit un tableau... les joies de PowerShell
                    $lastSourceDateList = @()
                }
                # S'il n'y a plus qu'un élément, ça va nous renvoyer un seul élément et pas un tableau... encore une joie de PowerShell
                elseif($lastSourceDateList -isnot [Array])
                {
                    # On retransforme donc en tableau
                    $lastSourceDateList = @($lastSourceDateList)
                }
            }

            $lastSourceDateList += $newSourceInfos

            # Mise à jour du fichier Log
            $lastSourceDateList | ConvertTo-Json | Out-file $lastSourceDateListFile -Encoding:utf8
        }
        else # Pas pu trouver de date
        {
            $logHistory.addWarningAndDisplay("Pas possible de trouver la date de mise à jour, veuillez contrôler la valeur de 'dateRegex' pour la source courante dans le fichier JSON de configuration")
        }

    }

    $sleepMin = 10
    $logHistory.addLine("Toutes les sources ont été contrôlées... attente de $($sleepMin) minutes jusqu'au prochain contrôle...")
    Start-Sleep -Seconds (60 * $sleepMin)
}
