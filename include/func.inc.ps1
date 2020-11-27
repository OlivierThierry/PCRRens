<#
    -------------------------------------------------------------------------------------
    BUT : Permet de savoir si un objet contient une propriété d'un nom donné.
    
    IN  : $obj     		-> L'objet concerné
    IN  : $propertyName -> Nom de la propriété que l'on cherche

    RET : $true ou $false
#>
function objectPropertyExists([PSCustomObject]$obj, [string]$propertyName)
{
	return ((($obj).PSobject.Properties | Select-Object -ExpandProperty "Name") -contains $propertyName)
}


<#
    -------------------------------------------------------------------------------------
    BUT : Emet une alerte sonore

    IN  : $speechSynthesizer    -> Objet pour faire de la synthèse vocale
    IN  : $message              -> message à dire (il faut qu'il soit en anglais)
#>
function soundAlert([System.Speech.Synthesis.SpeechSynthesizer]$speechSynthesizer, [string]$message)
{
    [System.Console]::Beep(500, 100)
    # [System.Console]::Beep(1000, 100)
    # [System.Console]::Beep(2000, 100)
    [System.Console]::Beep(1000, 100)
    [System.Console]::Beep(500, 100)

    $speechSynthesizer.Speak($message)
}