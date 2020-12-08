<#
    BUT : Classe contenant les fonctions pour traiter les données d'une page Web (extraction)
            et générer un fichier CSV de données (via la fonction 'writeCSVFile')
#>
class WebCallbackFunc : WebSearchFunc
{
    hidden [string]$outputFolder

    <#
        -------------------------------------------------------------------------------------
        BUT : Constructeur de classe

        IN  : $outputFolder -> Chemin jusqu'au dossier où créer les fichiers CSV avec les données
    #>
    WebCallbackFunc([string]$outputFolder): base()
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


    <#
        -------------------------------------------------------------------------------------
        BUT : Extrait les données qui sont présente sur le site suivant pour créer un fichier 
                de données exploitables directement pour la mise à jour sur les panneaux
            https://www.vd.ch/toutes-les-actualites/hotline-et-informations-sur-le-coronavirus/point-de-situation-statistique-dans-le-canton-de-vaud/
        
        IN  : $source     	-> Objet représentant la source que l'on traite
        
        RET : Chemin jusqu'au fichier de données créé
    #>
    [string] extractVaudSituation([PSObject]$source)
    {
        $response = Invoke-WebRequest -uri $source.location

        # Recherche du lien qui pointe sur le fichier Excel de données
        $excelLink = $response.Links | Where-Object { $_.innerText -like "*Donn*es compl*tes depuis*"}

        # Fichier temporaire pour télécharger le fichier Excel
        $tmpFile = New-TemporaryFile 
        
        # Téléchargement du fichier Excel
        Invoke-WebRequest -Uri $excelLink.href -OutFile $tmpFile.FullName

        $objExcel = New-Object -ComObject Excel.Application  
        $workBook = $objExcel.Workbooks.Open($tmpFile.FullName)  
        $workSheet = $workBook.Sheets.Item(1)

        # Recherche des informations dans la page
        $fields = @(
            @(
                "Vaud - NB nouveaux cas"
                ($workSheet.Cells.Item(4,5).text - $workSheet.Cells.Item(5,5).text)
            ),
            @(
                "Vaud - NB décès supp."
                ($workSheet.Cells.Item(4,4).text - $workSheet.Cells.Item(5,4).text)
            ),
            @(
                "Vaud - NB hospitalisations"
                $workSheet.Cells.Item(4,2).text
            ),
            @(
                "Vaud - NB hospitalisations soins intensifs"
                $workSheet.Cells.Item(4,3).text
            )
        )

        # Suppression du fichier temporaire
        Remove-Item $tmpFile.FullName

        # Génération du nom du fichier de données
        return $this.writeCSVFile("Vaud-Stats", @("Champ", "Valeur"), $fields)
    }


    [void] sezaneDone([PSObject]$source)
    {
        $mailConfigFile = ([IO.Path]::Combine("$PSScriptRoot", "..", "data", "mail-config.json"))
        if(! (Test-Path $mailConfigFile))
        {
            Throw "Mail config file not found!"
        }

        $mailParams = Get-Content $mailConfigFile | ConvertFrom-JSON
        $mailMessage = $mailParams.mailMessage -f $source.location

        $smtpServer = "smtp.gmail.com" 
         
         
        $msg = new-object Net.Mail.MailMessage 
        $smtp = new-object Net.Mail.SmtpClient($smtpServer) 
        $smtp.EnableSsl = $true 
        $msg.From = $mailParams.username
        $msg.To.Add($mailParams.mailTo) 
        $msg.BodyEncoding = [system.Text.Encoding]::Unicode 
        $msg.SubjectEncoding = [system.Text.Encoding]::Unicode 
        $msg.IsBodyHTML = $true  
        $msg.Subject = $mailParams.mailSubject
        $msg.Body = $mailMessage
        $SMTP.Credentials = New-Object System.Net.NetworkCredential($mailParams.username, $mailParams.password); 
        Write-Host ("Source changée! envoi d'un mail à {0}" -f $mailParams.mailTo)
        $smtp.Send($msg)
        Write-Host "Et on sort..."
        exit
        
    }

}

