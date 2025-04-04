#Requires -Version 5.1
#alarson@hbs.net - 2025-04-2
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'
Start-Transcript -Path "$($env:ProgramData)\Microsoft\IntuneManagementExtension\Logs\HPBloat.log" -Append

$orderofuninstall = @(
        "Poly Lens"
        "HP Client Security Manager"
        "HP Notifications"
        "HP Security Update Service"
        "HP System Default Settings"
        "HP Wolf Security"
        "HP Wolf Security Application Support for Sure Sense"
        "HP Wolf Security Application Support for Windows"
        "AD2F1837.HPPCHardwareDiagnosticsWindows"
        "AD2F1837.HPPowerManager"
        "AD2F1837.HPPrivacySettings"
        "AD2F1837.HPQuickDrop"
        "AD2F1837.HPSupportAssistant"
        "AD2F1837.HPSystemInformation"
        "AD2F1837.myHP"
        "RealtekSemiconductorCorp.HPAudioControl",
        "HP Sure Recover",
        "HP Sure Run Module"
        "RealtekSemiconductorCorp.HPAudioControl_2.39.280.0_x64__dt26b99r8h8gj"
        "HP Wolf Security - Console"
        "HP Wolf Security Application Support for Chrome 122.0.6261.139"
        "Windows Driver Package - HP Inc. sselam_4_4_2_453 AntiVirus  (11/01/2022 4.4.2.453)"
        "HP Insights"
        "HP Insights Analytics"
        "HP Insights Analytics - Dependencies"
    )

Function Get-HP-MSI {
    $registryPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )

    $msiApplications = foreach ($path in $registryPaths) {
        if (Test-Path $path) {
            Get-ChildItem -Path $path | ForEach-Object {
                $appName = $_.GetValue("DisplayName", $null)
                $uninstallString = $_.GetValue("UninstallString", $null)

                if ($appName -match "(HP|ICS)" -and $uninstallString -match "{[A-F0-9\-]+}") {
                    $guid = $matches[0]  

                    [PSCustomObject]@{
                        Name = $appName
                        GUID = $guid
                    }
                }
            }
        }
    }

    return $msiApplications  # Ensure the function returns results
}
$initalHPBloatCheck = Get-HP-MSI 

$sortedHPApps = $initalHPBloatCheck | Sort-Object {
    $index = $orderofuninstall.IndexOf($_.Name)
    if ($index -ge 0) { $index } else { [int]::MaxValue }
}


Try{ 
    if ($sortedHPApps){
        Write-Output "These HP app exist:"
            foreach ($msi in $sortedHPApps) {Write-Output $msi.name }
            Write-Output "Starting HP Bloat Removal for these apps.."
 

                foreach ($msi in $sortedHPApps) { 
                Write-Output "Uninstalling $($msi.Name)"
                Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $($msi.GUID) /qn" -Wait -Verbose
}
        #Remove HP Documentation if it exists
        Write-Output 'Checking for HP Documenation'
        if (test-path -Path "C:\Program Files\HP\Documentation\Doc_uninstall.cmd") {
        Write-Output 'HP Documentation found'
        Start-Process -FilePath "C:\Program Files\HP\Documentation\Doc_uninstall.cmd" -Wait -passthru -NoNewWindow
    }


        #Remove HP Connect Optimizer if setup.exe exists
        Write-Output 'Checking for HP Connection Optimizer'
        if (test-path -Path 'C:\Program Files (x86)\InstallShield Installation Information\{6468C4A5-E47E-405F-B675-A70A70983EA6}\setup.exe') {
        Write-Output 'Connection Optimizer found'
        invoke-webrequest -uri "https://raw.githubusercontent.com/alarson-cloud/HP-De-Bloat/refs/heads/main/HPConnectionOptimizer.iss" -outfile "C:\Windows\Temp\HPConnectionOptimizer.iss"

        &'C:\Program Files (x86)\InstallShield Installation Information\{6468C4A5-E47E-405F-B675-A70A70983EA6}\setup.exe' @('-s', '-f1C:\Windows\Temp\HPConnectionOptimizer.iss')
    }

        Write-Output 'Testing path Amazon'
        if (Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Amazon.com.lnk" -PathType Leaf) { Remove-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Amazon.com.lnk" -Force }
        Write-Output 'Testing path Angebote'
        if (Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Angebote.lnk" -PathType Leaf) { Remove-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Angebote.lnk" -Force }
        Write-Output 'Testing path TCO Certified'
        if (Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\TCO Certified.lnk" -PathType Leaf) { Remove-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\TCO Certified.lnk" -Force }
        Write-Output 'Testing path Booking.com'
        if (Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Booking.com.lnk" -PathType Leaf) { Remove-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Booking.com.lnk" -Force }
        Write-Output 'Testing path Adobe offers'
        if (Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Adobe offers.lnk" -PathType Leaf) { Remove-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Adobe offers.lnk" -Force }
        Write-Output 'Testing path Miro Offer'
        if (Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Miro Offer.lnk" -PathType Leaf) { Remove-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Miro offer.lnk" -Force }

        Get-CimInstance -ClassName Win32_Product | Where-Object { $_.Name -eq 'HP Wolf Security' } | Invoke-CimMethod -MethodName Uninstall
        Get-CimInstance -ClassName Win32_Product | Where-Object { $_.Name -eq 'HP Wolf Security - Console' } | Invoke-CimMethod -MethodName Uninstall
        Get-CimInstance -ClassName Win32_Product | Where-Object { $_.Name -eq 'HP Security Update Service' } | Invoke-CimMethod -MethodName Uninstall

Write-output "Starting Microsoft 365 Apps uninstallation..."

$OfficeUninstallStrings = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where {$_.DisplayName -like "*Microsoft 365 - en-us*"} | Select UninstallString).UninstallString
    ForEach ($UninstallString in $OfficeUninstallStrings) {
        Write-output "Uninstalling: $UninstallString"
        $UninstallEXE = ($UninstallString -split '"')[1]
        $UninstallArg = ($UninstallString -split '"')[2] + " DisplayLevel=False"
        Start-Process -FilePath $UninstallEXE -ArgumentList $UninstallArg -Wait
    }    


$OfficeUninstallStrings = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where {$_.DisplayName -like "*Aplicaciones de Microsoft 365*"} | Select UninstallString).UninstallString
    ForEach ($UninstallString in $OfficeUninstallStrings) {
        Write-output "Uninstalling: $UninstallString"
        $UninstallEXE = ($UninstallString -split '"')[1]
        $UninstallArg = ($UninstallString -split '"')[2] + " DisplayLevel=False"
        Start-Process -FilePath $UninstallEXE -ArgumentList $UninstallArg -Wait
    }    

$OfficeUninstallStrings = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where {$_.DisplayName -like "*Microsoft 365 Apps for Enterprise - fr-ca*"} | Select UninstallString).UninstallString
    ForEach ($UninstallString in $OfficeUninstallStrings) {
        Write-output "Uninstalling: $UninstallString"
        $UninstallEXE = ($UninstallString -split '"')[1]
        $UninstallArg = ($UninstallString -split '"')[2] + " DisplayLevel=False"
        Start-Process -FilePath $UninstallEXE -ArgumentList $UninstallArg -Wait
    }    


$OfficeUninstallStrings = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where {$_.DisplayName -like "*Microsoft OneNote - es-mx*"} | Select UninstallString).UninstallString
    ForEach ($UninstallString in $OfficeUninstallStrings) {
        Write-output "Uninstalling: $UninstallString"
        $UninstallEXE = ($UninstallString -split '"')[1]
        $UninstallArg = ($UninstallString -split '"')[2] + " DisplayLevel=False"
        Start-Process -FilePath $UninstallEXE -ArgumentList $UninstallArg -Wait
    }   
    
$OfficeUninstallStrings = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where {$_.DisplayName -like "*Microsoft OneNote - fr-ca*"} | Select UninstallString).UninstallString
    ForEach ($UninstallString in $OfficeUninstallStrings) {
        Write-output "Uninstalling: $UninstallString"
        $UninstallEXE = ($UninstallString -split '"')[1]
        $UninstallArg = ($UninstallString -split '"')[2] + " DisplayLevel=False"
        Start-Process -FilePath $UninstallEXE -ArgumentList $UninstallArg -Wait
    }  

Write-output "Microsoft 365 Apps uninstallation script completed."

    $finalHPBloatCheck = Get-HP-MSI 
        if ($finalHPBloatCheck){
        Write-Output "These HP app still exist:"
            foreach ($msi in $finalHPBloatCheck) {Write-Output $msi.name }
            $exitCode = 1 
    } else { 
        Write-Output "All HP Bloat Removed"
        $exitCode = 0
    } 

}else{
    Write-Output "No HP Bloat Apps detected"
    $exitCode = 0

}}catch{
	$errMsg = $_.Exception.Message
	Write-Error $errMsg 
	$exitCode = 1
}finally{
	Stop-Transcript
	exit $exitCode
}
     
