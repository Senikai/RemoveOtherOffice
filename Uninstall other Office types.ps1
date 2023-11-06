$ErrorActionPreference = “SilentlyContinue”

Stop-Process -Name EXCEL -Force
Stop-Process -Name MSAccess -Force
Stop-Process -Name OneNote -Force
Stop-Process -Name Outlook -Force
Stop-Process -Name POWERPNT -Force
Stop-Process -Name WINProj -Force
Stop-Process -Name Visio -Force
Stop-Process -Name MSPub -Force
Stop-Process -Name WINWORD -Force
Stop-Process -Name Acrobat -Force

$OfficeUninstallStrings = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where {$_.DisplayName -like "*Microsoft 365 - en-us*"} | Select UninstallString).UninstallString
    ForEach ($UninstallString in $OfficeUninstallStrings) {
        $UninstallEXE = ($UninstallString -split '"')[1]
        $UninstallArg = ($UninstallString -split '"')[2] + " DisplayLevel=False"
        Start-Process -FilePath $UninstallEXE -ArgumentList $UninstallArg -Wait
    }
    $OfficeUninstallStrings1 = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where {$_.DisplayName -like "*Microsoft 365 Apps for business - en-us*"} | Select UninstallString).UninstallString
    ForEach ($UninstallString in $OfficeUninstallStrings1) {
        $UninstallEXE = ($UninstallString -split '"')[1]
        $UninstallArg = ($UninstallString -split '"')[2] + " DisplayLevel=False"
        Start-Process -FilePath $UninstallEXE -ArgumentList $UninstallArg -Wait
    }
$OfficeUninstallStrings2 = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where {$_.DisplayName -like "*Microsoft 365 - es-es*"} | Select UninstallString).UninstallString
    ForEach ($UninstallString in $OfficeUninstallStrings2) {
        $UninstallEXE = ($UninstallString -split '"')[1]
        $UninstallArg = ($UninstallString -split '"')[2] + " DisplayLevel=False"
        Start-Process -FilePath $UninstallEXE -ArgumentList $UninstallArg -Wait
    }
$OfficeUninstallStrings3 = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where {$_.DisplayName -like "*Microsoft 365 - fr-fr*"} | Select UninstallString).UninstallString
    ForEach ($UninstallString in $OfficeUninstallStrings3) {
        $UninstallEXE = ($UninstallString -split '"')[1]
        $UninstallArg = ($UninstallString -split '"')[2] + " DisplayLevel=False"
        Start-Process -FilePath $UninstallEXE -ArgumentList $UninstallArg -Wait
    }
    $OfficeUninstallStrings4 = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where {$_.DisplayName -like "*de Microsoft*"} | Select UninstallString).UninstallString
    ForEach ($UninstallString in $OfficeUninstallStrings4) {
        $UninstallEXE = ($UninstallString -split '"')[1]
        $UninstallArg = ($UninstallString -split '"')[2] + " DisplayLevel=False"
        Start-Process -FilePath $UninstallEXE -ArgumentList $UninstallArg -Wait
        }
        $OfficeUninstallStrings5 = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where {$_.DisplayName -like "*Microsoft 365 Apps for business - fr-fr*"} | Select UninstallString).UninstallString
    ForEach ($UninstallString in $OfficeUninstallStrings5) {
        $UninstallEXE = ($UninstallString -split '"')[1]
        $UninstallArg = ($UninstallString -split '"')[2] + " DisplayLevel=False"
        Start-Process -FilePath $UninstallEXE -ArgumentList $UninstallArg -Wait
        }
$OfficeUninstallStrings6 = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where {$_.DisplayName -like "*Microsoft 365 - pt-br*"} | Select UninstallString).UninstallString
    ForEach ($UninstallString in $OfficeUninstallStrings6) {
        $UninstallEXE = ($UninstallString -split '"')[1]
        $UninstallArg = ($UninstallString -split '"')[2] + " DisplayLevel=False"
        Start-Process -FilePath $UninstallEXE -ArgumentList $UninstallArg -Wait
    }
    $OfficeUninstallStrings8 = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where {$_.DisplayName -like "*OneNote - es-es*"} | Select UninstallString).UninstallString
    ForEach ($UninstallString in $OfficeUninstallStrings8) {
        $UninstallEXE = ($UninstallString -split '"')[1]
        $UninstallArg = ($UninstallString -split '"')[2] + " DisplayLevel=False"
        Start-Process -FilePath $UninstallEXE -ArgumentList $UninstallArg -Wait
    }
    $OfficeUninstallStrings9 = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where {$_.DisplayName -like "*OneNote - fr-fr*"} | Select UninstallString).UninstallString
    ForEach ($UninstallString in $OfficeUninstallStrings9) {
        $UninstallEXE = ($UninstallString -split '"')[1]
        $UninstallArg = ($UninstallString -split '"')[2] + " DisplayLevel=False"
        Start-Process -FilePath $UninstallEXE -ArgumentList $UninstallArg -Wait
    }
    $OfficeUninstallStrings10 = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where {$_.DisplayName -like "*OneNote - pt-br*"} | Select UninstallString).UninstallString
    ForEach ($UninstallString in $OfficeUninstallStrings10) {
        $UninstallEXE = ($UninstallString -split '"')[1]
        $UninstallArg = ($UninstallString -split '"')[2] + " DisplayLevel=False"
        Start-Process -FilePath $UninstallEXE -ArgumentList $UninstallArg -Wait
    }