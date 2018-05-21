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
    if($link -is [array])
    {
        return $link[0]
    }
    else{
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

while (!$loginItemName -and !$link) {
    Start-Sleep 2
    $loginItemName = GetLoginItemName $ie "login"
    $link = GetVDIItemLink($ie)
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

while (!$link) {
    $link = GetVDIItemLink($ie)
    Start-Sleep 3
}

"Starting VDI... Enjoy the session."

$link.click()

#wait for some time before exit so that user can see the activity log.
Start-Sleep -s 5

# close browser.
$ie.Quit();