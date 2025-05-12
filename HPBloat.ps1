#Requires -Version 5.1
#alarson@hbs.net - 2025-04-2
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'
Start-Transcript -Path "$($env:ProgramData)\Microsoft\IntuneManagementExtension\Logs\HPBloat.log" -Append
$exitCode = 0

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
    "RealtekSemiconductorCorp.HPAudioControl"
    "HP Sure Recover"
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

    return $msiApplications
}

Try {
    $initalHPBloatCheck = Get-HP-MSI
    $sortedHPApps = $initalHPBloatCheck | Sort-Object {
        $index = $orderofuninstall.IndexOf($_.Name)
        if ($index -ge 0) { $index } else { [int]::MaxValue }
    }

        if ($sortedHPApps) {
                Write-Output "These HP app exist:"
                foreach ($msi in $sortedHPApps) { Write-Output $msi.name }
                Write-Output "Starting HP Bloat Removal for these apps.."
        
                foreach ($msi in $sortedHPApps) {
                    Write-Output "Uninstalling $($msi.Name)"
                    Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $($msi.GUID) /qn" -Wait -Verbose
                }
        
                # Remove HP Documentation if it exists
                Write-Output 'Checking for HP Documentation'
                if (Test-Path -Path "C:\Program Files\HP\Documentation\Doc_uninstall.cmd") {
                    Write-Output 'HP Documentation found'
                    Start-Process -FilePath "C:\Program Files\HP\Documentation\Doc_uninstall.cmd" -Wait -passthru -NoNewWindow
                }
        
                # Remove HP Connection Optimizer
                Write-Output 'Checking for HP Connection Optimizer'
                if (Test-Path -Path 'C:\Program Files (x86)\InstallShield Installation Information\{6468C4A5-E47E-405F-B675-A70A70983EA6}\setup.exe') {
                    Write-Output 'Connection Optimizer found'
                    Invoke-WebRequest -Uri "https://raw.githubusercontent.com/alarson-cloud/HP-De-Bloat/refs/heads/main/HPConnectionOptimizer.iss" -OutFile "C:\Windows\Temp\HPConnOpt.iss"
                    & 'C:\Program Files (x86)\InstallShield Installation Information\{6468C4A5-E47E-405F-B675-A70A70983EA6}\setup.exe' @('-s', '-f1C:\Windows\Temp\HPConnOpt.iss')
                }
        
                # Remove shortcuts
                Write-Output 'Removing Start Menu shortcuts'
                $shortcuts = @(
                    "Amazon.com.lnk",
                    "Angebote.lnk",
                    "TCO Certified.lnk",
                    "Booking.com.lnk",
                    "Adobe offers.lnk",
                    "Miro Offer.lnk"
                )
                foreach ($shortcut in $shortcuts) {
                    $path = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\$shortcut"
                    if (Test-Path -Path $path -PathType Leaf) {
                        Remove-Item -Path $path -Force
                        Write-Output "Removed: $shortcut"
                    }
                }
        
                # Remove specific HP apps
                "HP Wolf Security", "HP Wolf Security - Console", "HP Security Update Service" | ForEach-Object {
                    Get-CimInstance -ClassName Win32_Product | Where-Object { $_.Name -eq $_ } | Invoke-CimMethod -MethodName Uninstall
                }
        
                $finalHPBloatCheck = Get-HP-MSI
                if ($finalHPBloatCheck) {
                    Write-Output "These HP apps still exist:"
                    foreach ($msi in $finalHPBloatCheck) { Write-Output $msi.name }
                    $exitCode = 1
                } else {
                    Write-Output "All HP Bloat Removed"
                    $exitCode = 0
                }
            } else {
                Write-Output "No Bloat Apps detected"
                $exitCode = 0
            }
        
            # ---------- MICROSOFT 365 UNINSTALL SECTION ----------
            Start-Sleep -Seconds 20
            Write-Output "Checking for Microsoft 365 Apps..."

            $patterns = @(
                "*Microsoft 365 - en-us*",
                "*Aplicaciones de Microsoft 365*",
                "*Microsoft 365 Apps for Enterprise - fr-ca*",
                "*Microsoft OneNote - es-mx*",
                "*Microsoft OneNote - fr-ca*"
            )

            # Collect all uninstall strings matching the patterns
            $OfficeUninstallStrings = foreach ($pattern in $patterns) {
                Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction SilentlyContinue |
                        Where-Object { $_.PSObject.Properties.Name -contains "DisplayName" -and $_.DisplayName -like $pattern } |
                        Select-Object -ExpandProperty UninstallString -ErrorAction SilentlyContinue
            }

            # If anything was found, proceed
            if ($OfficeUninstallStrings) {
                Write-Output "Starting Microsoft 365 Apps uninstallation..."

                    foreach ($UninstallString in $OfficeUninstallStrings) {
                        if (-not [string]::IsNullOrWhiteSpace($UninstallString)) {
                        Write-Output "Uninstalling: $UninstallString"

                        try {
                            $UninstallEXE = ($UninstallString -split '"')[1]
                            $UninstallArg = ($UninstallString -split '"')[2] + " DisplayLevel=False"
                            Start-Process -FilePath $UninstallEXE -ArgumentList $UninstallArg -Wait
                      } catch {
                Write-Warning "Failed to uninstall with string: $UninstallString. Error: $_"
            }
        }
    }

    Write-Output "Microsoft 365 Apps uninstallation script completed."
} else {
    Write-Output "No Microsoft 365 Apps install detected."
}

                   # ---------- AppX UNINSTALL SECTION ----------

$Appx = @(
    "*CandyCrush*"
    "*DevHome*"
    "*Disney*"
    "*Dolby*"
    "*EclipseManager*"
    "*Facebook*"
    "*Flipboard*"
    "*gaming*"
    "*Minecraft*"
    "*Office*"
    "*PandoraMediaInc*"
    "*Speed Test*"
    "*Spotify*"
    "*Sway*"
    "*Twitter*"
    "Microsoft.MicrosoftSolitaireCollection"
    "Microsoft.Xbox.TCUI"
    "Microsoft.XboxApp"
    "Microsoft.XboxGameOverlay"
    "Microsoft.XboxGameCallableUI"
    "Microsoft.XboxGamingOverlay"
    "Microsoft.XboxGamingOverlay_5.721.10202.0_neutral_~_8wekyb3d8bbwe"
    "Microsoft.XboxIdentityProvider"
    "Microsoft.XboxSpeechToTextOverlay"
        )

        $appXInstalled = Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -in $Appx}
        if($appXInstalled){ 
            foreach($app in $appXInstalled){
                Remove-AppxProvisionedPackage -PackageName $app.PackageName -Online -ErrorAction SilentlyContinue
                Write-Output "Removed $($app.DisplayName)" 
            }
            $AllappXInstalled = Get-AppxPackage -AllUsers | Where-Object {$_.Name -in $Appx}
            foreach($app in $AllappXInstalled){ 
                Remove-AppxPackage -package $app.PackageFullName -AllUsers -ErrorAction SilentlyContinue
                Write-Output "Removed $($app.Name)"

            }
            

        }else { 
            Write-Output "No Appx detected."
        } 
                
            
        
        } Catch {
            $errMsg = if ($_.Exception -and $_.Exception.Message) { $_.Exception.Message } else { "Unknown error occurred." }
            Write-Error "An error occurred: $errMsg`nFull error details: $_"
            $exitCode = 1
        } Finally {
            Stop-Transcript
            exit $exitCode
        }
        
