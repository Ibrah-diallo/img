


# Maxime paquet Template updater script
#-------------------------------------------------
if (!$copyPath) {
    New-Item -ItemType Directory -Path $env:appdata\Microsoft\RS-Templates
}
$Username = "administrateur"
$Password = "XouLouWP23*-"

$pathToFile = "\\10.2.1.36\TemplateTemp$\Lettre Reference Systeme inc.dotx" 
$copyPath = "$env:appdata\Microsoft\RS-Templates\Lettre Reference Systeme inc.dotx"

$WebClient = New-Object System.Net.WebClient
$WebClient.Credentials = New-Object System.Net.NetworkCredential($Username, $Password)

$WebClient.DownloadFile($pathToFile, $copyPath)

#Copy-Item $pathToFile -Destination $copyPath -Recurse -Force # - dont use
#-------------------------------------------------
$KeyPath = "HKCU:\Software\Microsoft\Office\16.0\Common\General"

$name = "SharedTemplates"

Set-ItemProperty -Path $KeyPath -Name $name -Value "C:\users\$env:USERNAME\AppData\Roaming\Microsoft\RS-Templates" -Force