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
        $logHistory.addLine("Checking $($source.name)...")

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

        $logHistory.addLine("> Looking for update date in webpage ($($source.url))")
        $cmd = '$result = $response.{0} | Where-Object {{ {1} }}' -f $checkIn, ($conditions -join " -and ")
        Invoke-Expression $cmd

        # Si pas de résultat trouvé
        if($null -eq $result)
        {
            $logHistory.addWarningAndDisplay("Cannot find node in HTML DOM, please check 'searchFilters' values in sources file ($($sourcesFile))")
            continue
        }

        # Si les filtres définis ne sont pas assez précis et que plusieurs résultats sont renvoyés,
        if($result -is [System.Array])
        {
            $logHistory.addWarningAndDisplay("More than one node found in HTML DOM, please check 'searchFilters' values in sources file ($($sourcesFile)) to add more precision in search")
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
                $logHistory.addWarningAndDisplay("Web source date is empty, please check regex in JSON file")
                continue
            }

            $logHistory.addLine("> Last web page update is ($($webSourceDate))")
            # Si date présente dans le fichier de log et différente de la courante
            # OU
            # date pas présente dans le fichier log
            $sourceInfos = $lastSourceDateList | Where-Object { $_.name -eq $source.name}
            if((($null -ne $sourceInfos) -and ($sourceInfos.date -ne $webSourceDate) ) `
                -or `
                ($null -eq $sourceInfos))
            {
                $logHistory.addLine("> Web page has been updated since last check")
                
                
                Write-Host ">> " -NoNewline -ForegroundColor:Green
                Write-host "'$($source.name)' source updated! ($($webSourceDate))`n$($source.url)"
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
            $logHistory.addWarningAndDisplay("Cannot find update date, please check regex in JSON file")
        }

    }

    $sleepMin = 10
    $logHistory.addLine("All sources have been checked, sleeping $($sleepMin) until next check...")
    Start-Sleep -Seconds (60 * $sleepMin)
}
