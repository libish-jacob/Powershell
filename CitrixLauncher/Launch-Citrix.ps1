#Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
#$ErrorActionPreference = "silentlyContinue"

function GetLoginItemName([psobject] $ie, [string] $tagName) {
    $inputs = $ie.Document.getElementsByTagName("input")
    $loginItemName = ($inputs | Where-Object {$_.name -eq $tagName})
    return $loginItemName
}

function GetLoginItemId([psobject] $ie, [string] $tagName) {
    $inputs = $ie.Document.getElementsByTagName("input")
    $loginItemName = ($inputs | Where-Object {$_.id -eq $tagName})
    return $loginItemName
}

function GetVDIItemLink([psobject] $ie) {
    $inputs = $ie.Document.getElementsByTagName("div")
    $link = ($inputs | Where-Object {$_.id -Like "desktopSpinner_idCitrix.MPS.Desktop.XD75.XD_0020Windows_0020Dedicated_*"})   
    if ($link -is [array]) {
        return $link[0]
    }
    else {
        return $link
    }    
}

function GetItem([psobject] $ie, [String]$element, [string]$id) {
    $inputs = $ie.Document.getElementsByTagName($element)
    $link = ($inputs | Where-Object {$_.id -eq $id})   
    if ($link -is [array]) {
        return $link[0]
    }
    else {
        return $link
    }    
}


$ie = new-object -com internetexplorer.application
$ie.visible = $false
$ie.ParsedHtml
$ie.navigate2("citrix")

"Launching Please wait..."
while ($ie.busy) {Start-Sleep 2}

$loginItemName = GetLoginItemName $ie "login" 
$link = GetVDIItemLink($ie)
$iteration = 100
while (($iteration -gt 0) -and (!$loginItemName -and !$link)) {
    Start-Sleep 2
    $loginItemName = GetLoginItemName $ie "login"
    $link = GetVDIItemLink($ie)
    $iteration--
}

if ($loginItemName) {
    $username = $env:UserName
    $SecurePassword = Read-Host -AsSecureString -Prompt "Provide password for $username"
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
    $Password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    
    $loginItemName = GetLoginItemName $ie "login"
    if ($loginItemName) {
        "Logging in..."
        $loginItemName.value = $username
        $loginPasswordItem = GetLoginItemName $ie "passwd" 
        $loginPasswordItem.value = $Password

        $logonLink = GetLoginItemId $ie "Log_On"
        $logonLink.click()
    }
}

$iteration = 100
$errorContainer = GetItem $ie "span" "errorMessageLabel"
while (($iteration -gt 0) -and (!$link -and !$errorContainer)) {
    $link = GetVDIItemLink($ie)
    $errorContainer = GetItem $ie "span" "errorMessageLabel"
    Start-Sleep 3
    $iteration--
}

if ($link) {
    "Starting VDI... Enjoy the session."
    $link.click()
}
elseif ($errorContainer) {
    $errorContainer.textContent
}
else {
    "Couldnt find any VDI. Please try manually."
}

#wait for some time before exit so that user can see the activity log.
Start-Sleep 5

# close browser.
$ie.Quit();