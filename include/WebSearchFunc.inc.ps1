

class WebSearchFunc
{

    WebSearchFunc()
    {

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



    [string] findGiletAchille([PSObject]$source)
    {
        $webPage = Invoke-WebRequest -uri $source.location

        return $this.findInPage($webPage, @('<ul class="o-grid o-grid--gutter-sm u-flex-cross-center u-list-unstyled u-none@sm u-mt-n"', `
                                    'data-size="S"', `
                                    '<label', `
                                    'class="'), '"')

    }


    [string] findRobeBarbara([PSObject]$source)
    {
        $webPage = Invoke-WebRequest -uri $source.location

        return $this.findInPage($webPage, @('<ul class="o-grid o-grid--gutter-sm u-flex-cross-center u-list-unstyled u-none@sm u-mt-n"', `
                                    'data-size="38"', `
                                    '<label', `
                                    'class="'), '"')

    }

    

}