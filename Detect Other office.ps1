$Blacklist = @(
    "*Microsoft 365 - en-us*",
    "*Microsoft 365 Apps for business - en-us*",
    "*es-es*",
    "*fr-fr*",
    "*de Microsoft*",
    "*Microsoft 365 Apps for business - fr-fr*"
)

foreach ($App in $Blacklist) {
    if (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where {$_.DisplayName -like $App}) {
        exit 1
    }
}
Else{
exit 0
}