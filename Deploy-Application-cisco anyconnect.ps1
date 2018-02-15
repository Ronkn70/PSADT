<#
.SYNOPSIS
	This script performs the installation or uninstallation of an application(s).
.DESCRIPTION
	The script is provided as a template to perform an install or uninstall of an application(s).
	The script either performs an "Install" deployment type or an "Uninstall" deployment type.
	The install deployment type is broken down into 3 main sections/phases: Pre-Install, Install, and Post-Install.
	The script dot-sources the AppDeployToolkitMain.ps1 script which contains the logic and functions required to install or uninstall an application.
.PARAMETER DeploymentType
	The type of deployment to perform. Default is: Install.
.PARAMETER DeployMode
	Specifies whether the installation should be run in Interactive, Silent, or NonInteractive mode. Default is: Interactive. Options: Interactive = Shows dialogs, Silent = No dialogs, NonInteractive = Very silent, i.e. no blocking apps. NonInteractive mode is automatically set if it is detected that the process is not user interactive.
.PARAMETER AllowRebootPassThru
	Allows the 3010 return code (requires restart) to be passed back to the parent process (e.g. SCCM) if detected from an installation. If 3010 is passed back to SCCM, a reboot prompt will be triggered.
.PARAMETER TerminalServerMode
	Changes to "user install mode" and back to "user execute mode" for installing/uninstalling applications for Remote Destkop Session Hosts/Citrix servers.
.PARAMETER DisableLogging
	Disables logging to file for the script. Default is: $false.
.EXAMPLE
    powershell.exe -Command "& { & '.\Deploy-Application.ps1' -DeployMode 'Silent'; Exit $LastExitCode }"
.EXAMPLE
    powershell.exe -Command "& { & '.\Deploy-Application.ps1' -AllowRebootPassThru; Exit $LastExitCode }"
.EXAMPLE
    powershell.exe -Command "& { & '.\Deploy-Application.ps1' -DeploymentType 'Uninstall'; Exit $LastExitCode }"
.EXAMPLE
    Deploy-Application.exe -DeploymentType "Install" -DeployMode "Silent"
.NOTES
	Toolkit Exit Code Ranges:
	60000 - 68999: Reserved for built-in exit codes in Deploy-Application.ps1, Deploy-Application.exe, and AppDeployToolkitMain.ps1
	69000 - 69999: Recommended for user customized exit codes in Deploy-Application.ps1
	70000 - 79999: Recommended for user customized exit codes in AppDeployToolkitExtensions.ps1
.LINK 
	http://psappdeploytoolkit.com
#>
[CmdletBinding()]
Param (
	[Parameter(Mandatory=$false)]
	[ValidateSet('Install','Uninstall')]
	[string]$DeploymentType = 'Install',
	[Parameter(Mandatory=$false)]
	[ValidateSet('Interactive','Silent','NonInteractive')]
	[string]$DeployMode = 'Interactive',
	[Parameter(Mandatory=$false)]
	[switch]$AllowRebootPassThru = $false,
	[Parameter(Mandatory=$false)]
	[switch]$TerminalServerMode = $false,
	[Parameter(Mandatory=$false)]
	[switch]$DisableLogging = $false
)

Try {
	## Set the script execution policy for this process
	Try { Set-ExecutionPolicy -ExecutionPolicy 'ByPass' -Scope 'Process' -Force -ErrorAction 'Stop' } Catch {}
	
	##*===============================================
	##* VARIABLE DECLARATION
	##*===============================================
	## Variables: Application
	[string]$appVendor = 'Cisco'
	[string]$appName = 'AnyConnect VPN Client'
	[string]$appVersion = '3.1.13015'
	[string]$appArch = ''
	[string]$appLang = 'EN'
	[string]$appRevision = '01'
	[string]$appScriptVersion = '1.0.0'
	[string]$appScriptDate = '05/03/2016'
	[string]$appScriptAuthor = ''
	##*===============================================
	## Variables: Install Titles (Only set here to override defaults set by the toolkit)
	[string]$installName = ''
	[string]$installTitle = ''
	
	##* Do not modify section below
	#region DoNotModify
	
	## Variables: Exit Code
	[int32]$mainExitCode = 0
	
	## Variables: Script
	[string]$deployAppScriptFriendlyName = 'Deploy Application'
	[version]$deployAppScriptVersion = [version]'3.6.8'
	[string]$deployAppScriptDate = '02/06/2016'
	[hashtable]$deployAppScriptParameters = $psBoundParameters
	
	## Variables: Environment
	If (Test-Path -LiteralPath 'variable:HostInvocation') { $InvocationInfo = $HostInvocation } Else { $InvocationInfo = $MyInvocation }
	[string]$scriptDirectory = Split-Path -Path $InvocationInfo.MyCommand.Definition -Parent
	
	## Dot source the required App Deploy Toolkit Functions
	Try {
		[string]$moduleAppDeployToolkitMain = "$scriptDirectory\AppDeployToolkit\AppDeployToolkitMain.ps1"
		If (-not (Test-Path -LiteralPath $moduleAppDeployToolkitMain -PathType 'Leaf')) { Throw "Module does not exist at the specified location [$moduleAppDeployToolkitMain]." }
		If ($DisableLogging) { . $moduleAppDeployToolkitMain -DisableLogging } Else { . $moduleAppDeployToolkitMain }
	}
	Catch {
		If ($mainExitCode -eq 0){ [int32]$mainExitCode = 60008 }
		Write-Error -Message "Module [$moduleAppDeployToolkitMain] failed to load: `n$($_.Exception.Message)`n `n$($_.InvocationInfo.PositionMessage)" -ErrorAction 'Continue'
		## Exit the script, returning the exit code to SCCM
		If (Test-Path -LiteralPath 'variable:HostInvocation') { $script:ExitCode = $mainExitCode; Exit } Else { Exit $mainExitCode }
	}
	
	#endregion
	##* Do not modify section above
	##*===============================================
	##* END VARIABLE DECLARATION
	##*===============================================
		
	If ($deploymentType -ine 'Uninstall') {
		##*===============================================
		##* PRE-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Pre-Installation'
			
		## <Perform Pre-Installation tasks here>
		
		# Start install script
        Write-Log "Cisco AnyConnect 3.1.13015 INSTALL: START SCRIPT"
        $isLaptop = (Test-Battery -PassThru).IsLaptop
        Write-Log "Is Laptop results: $isLaptop"
        Write-Log "Computer name: $envComputerName"
        Write-Log "Script running as: $envUserName"
        $systemAccount = "$envComputerName$"
        Write-Log "System account name: $systemAccount"
        
   		# Check for logged on users
        $getUser = (Get-LoggedOnUser)
        Write-Log "LOGGED ON USER results: $getUser"
        $userLoggedOn = "UNKNOWN"
        If ($getUser.ConnectState -eq 'Active') {
            Write-Log "YES, USER(s) is logged on to the computer, ACTIVE."
            $userLoggedOn = "YES"
        } ElseIf ($getUser.ConnectState -eq 'Disconnected') {
            Write-Log "YES, USER(s) is logged on to the computer, DISCONNECTED."
            $userLoggedOn = "YES"
        } Else {
            Write-Log "NO, USER is not logged on to the computer."
            $userLoggedOn = "NO"
        }
        Write-Log "LOGGED ON USER variable results: $userLoggedOn"
        
		# If powerpoint is running in presentation mode abort
	    $presentationPowerPoint = (Test-PowerPoint)
        Write-Log "PowerPoint presentation mode results: $presentationPowerPoint"
        If ($presentationPowerPoint -eq $true) {
            Write-Log "YES, detected PowerPoint in presentation mode, aborting script with exit code 69000."
            Start-Sleep -s 30
            Exit-Script -ExitCode "69000" 
        } Else {
            Write-Log "NO, PowerPoint is NOT in presentation mode."
        }

        If ($isLaptop -eq $true) {
            Write-Log "YES, chassis type is a laptop, continue with install."
        } Else {
            Write-Log "NO, chassis type is NOT a laptop, aborting script with exit code 69001."
            Show-InstallationPrompt -Message "Chassis type is NOT a laptop. Cisco AnyConnect VPN Client install ABORTED. NO changes made. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
            Start-Sleep -s 30
            Exit-Script -ExitCode "69001"
        }
        
        # Perform Cisco AnyConnect 3.1.13015 check task here
        $installAnyConnectCount = (Get-InstalledApplication -ProductCode "{EDEB4A62-FE20-4F95-8B90-26BB74CEB6A9}")
        Write-Log "Cisco AnyConnect 3.1.13015 check results: $installAnyConnectCount"
        If (($installAnyConnectCount | Measure-Object).Count -gt 0) {
            Write-Log "YES, Cisco AnyConnect 3.1.13015 is INSTALLED, aborting script with exit code 0."
            Show-InstallationPrompt -Message "Cisco AnyConnect VPN Client 3.1.13015 is previously INSTALLED. NO changes made. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
            Start-Sleep -s 30
            Exit-Script -ExitCode "0"
        } Else {
            Write-Log "NO, Cisco AnyConnect 3.1.13015 is NEEDED."
        }

	    $explorerRunning1 = (Get-Process explorer -ea 0) 
        Write-Log "RUNNING explorer BEGIN results: $explorerRunning1"
        
        # Stop the AnyConnect service
        Stop-Service -name "vpnagnet" -Force -ErrorAction SilentlyContinue
        Stop-Service -name "Cisco AnyConnect VPN Agent" -Force -ErrorAction SilentlyContinue
        
        Write-Log "START SHOWING INSTALLATION PROGRESS MESSAGES."
        Start-Sleep -s 60
        
        Write-Log "Show welcome message, close Cisco AnyConnect processes, allow up to 7 day deferral, and persist the prompt"
        # Show Welcome Message, close Cisco AnyConnect processes, allow up to 7 day deferral, and persist the prompt
        Show-InstallationWelcome -CloseApps "vpnagent,vpnui,vpndownloader,vacon" -AllowDeferCloseApps -AllowDefer -DeferTimes 25 -DeferDays 7 -PersistPrompt -MinimizeWindows $false
        Start-Sleep -s 30
        
        # Perform OLD Cisco AnyConnect check task here
        $installOldAnyConnectCount1 = (Get-InstalledApplication -Name "Cisco AnyConnect VPN Client")
        $installOldAnyConnectCount2 = (Get-InstalledApplication -Name "Cisco AnyConnect Secure Mobility Client")
        Write-Log "OLD Cisco AnyConnect #1 check results: $installOldAnyConnectCount1"
        Write-Log "OLD Cisco AnyConnect #2 check results: $installOldAnyConnectCount2"
        If ((($installOldAnyConnectCount1 | Measure-Object).Count -gt 0) -or (($installOldAnyConnectCount2 | Measure-Object).Count -gt 0)) {
            Write-Log "YES, OLD Cisco AnyConnect is INSTALLED."
            Write-Log "UNINSTALLING OLD Cisco AnyConnect."

            # Stop the AnyConnect service
            Stop-Service -name "vpnagnet" -Force -ErrorAction SilentlyContinue
            Stop-Service -name "Cisco AnyConnect VPN Agent" -Force -ErrorAction SilentlyContinue
            Stop-Service -name "Cisco AnyConnect Secure Mobility Agent" -Force -ErrorAction SilentlyContinue

            # Show Progress Message (with the default message) if any Cisco AnyConnect processes were running and Show-InstallationWelcome triggered
            Show-InstallationProgress -StatusMessage "UNINSTALLING previous Cisco AnyConnect VPN Client. `r`n`r`nThis may take some time. Please wait..." -WindowLocation 'Default' -TopMost $true
            Start-Sleep -s 30
                
            # Perform uninstallation tasks here, uninstalling Cisco AnyConnect
            Remove-MSIApplications "Cisco AnyConnect VPN Client" -ErrorAction SilentlyContinue
            Remove-MSIApplications "Cisco AnyConnect Secure Mobility Client" -ErrorAction SilentlyContinue
            Execute-MSI -Action Uninstall -Path "{7240A69A-AC53-46A1-9039-1281DDBBE452}"
            Execute-MSI -Action Uninstall -Path "{C1EC4E2D-6F63-4806-B88E-7685B6EC186E}"
            Execute-MSI -Action Uninstall -Path "{92083A9A-549D-4057-88E8-223EA08563FA}"
            
            Write-Log "OLD Cisco AnyConnect UNINSTALL COMPLETED." 
            # Display a message at the end of the install
            Show-InstallationProgress -StatusMessage "Previous Cisco AnyConnect VPN Client uninstall COMPLETE. `r`n`r`nPlease wait..." -WindowLocation 'Default' -TopMost $true
            Start-Sleep -s 30
        } Else {
            Write-Log "NO, OLD Cisco AnyConnect is NOT installed."
            Show-InstallationProgress -StatusMessage "NO previous Cisco AnyConnct VPN Client installed, nothing to uninstall. `r`n`r`nPlease wait..." -WindowLocation 'Default' -TopMost $true
            Start-Sleep -s 30
        }
				
		##*===============================================
		##* INSTALLATION 
		##*===============================================
		[string]$installPhase = 'Installation'
		
		## <Perform Installation tasks here>
		
        Write-Log "START installing Cisco AnyConnect 3.1.13015."
        # Show Progress Message (with the default message) if any Cisco AnyConnect processes were running and Show-InstallationWelcome triggered
        Show-InstallationProgress -StatusMessage "INSTALLING Cisco AnyConnect VPN Clinet 3.1.13015. `r`n`r`nThis may take some time. Please wait..." -WindowLocation 'Default' -TopMost $true        
        Start-Sleep -s 30
        
        # Perform installation tasks here
        Write-Log "INSTALLING Cisco AnyConnect 3.1.13015."
        Execute-MSI -Action 'Install' -Path "$dirFiles\anyconnect-win-3.1.13015-pre-deploy-k9.msi"-ContinueOnError $true

        # Create Cisco Profile variables / folders
        $strDefPath1="C:\ProgramData\Cisco\Cisco AnyConnect Secure Mobility Client\Profile"
        $strDefPath2="C:\Documents and Settings\All Users\Application Data\Cisco\Cisco AnyConnect Secure Mobility Client\Profile"
        
        If ($envOSName -eq "Microsoft Windows XP Professional") {
            Write-Log "OS Name: $envOSName"
            Write-Log "Windows XP, continue with the script."
            # Check to see if the folder exists, if not then create folder
            If ( -Not (Test-Path $strDefPath2 -PathType Any)) {
                New-Item -Path $strDefPath2 -ItemType Directory -Force -ErrorAction SilentlyContinue
            }
            # Copy Cisco AnyConnect Profile file
            Write-Log "COPYING Cisco AnyConnect Profile file."
            Copy-Item "$dirFiles\Freedom.xml" $strDefPath2 -Force -ErrorAction SilentlyContinue
        } ElseIf ($envOSName -eq "Microsoft(R) Windows(R) XP Professional x64 Edition") {
            Write-Log "OS Name: $envOSName"
            Write-Log "Windows XP, continue with the script."
            # Check to see if the folder exists, if not then create folder
            If ( -Not (Test-Path $strDefPath2 -PathType Any)) {
                New-Item -Path $strDefPath2 -ItemType Directory -Force -ErrorAction SilentlyContinue
            }
            # Copy Cisco AnyConnect Profile file
            Write-Log "COPYING Cisco AnyConnect Profile file."
            Copy-Item "$dirFiles\Freedom.xml" $strDefPath2 -Force -ErrorAction SilentlyContinue
        } Else {
            Write-Log "OS Name: $envOSName"
            Write-Log "Windows 7, continue with the script."
            # Check to see if the folder exists, if not then create folder
            If ( -Not (Test-Path $strDefPath1 -PathType Any)) {
                New-Item -Path $strDefPath1 -ItemType Directory -Force -ErrorAction SilentlyContinue
            }
            # Copy Cisco AnyConnect Profile file
            Write-Log "COPYING Cisco AnyConnect Profile file."
            Copy-Item "$dirFiles\Freedom.xml" $strDefPath1 -Force -ErrorAction SilentlyContinue
        }
        
        Write-Log "REMOVE registry start up of Cisco AnyConnect."
        CMD.EXE /C "C:\Windows\System32\reg.exe delete ""HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"" /v ""Cisco AnyConnect Secure Mobility Agent for Windows"" /f /reg:32"
        CMD.EXE /C "C:\Windows\SysWOW64\reg.exe delete ""HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"" /v ""Cisco AnyConnect Secure Mobility Agent for Windows"" /f /reg:32"
        CMD.EXE /C "C:\Windows\System32\reg.exe delete ""HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"" /v ""Cisco AnyConnect Secure Mobility Agent for Windows"" /f /reg:64"
        CMD.EXE /C "C:\Windows\SysWOW64\reg.exe delete ""HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"" /v ""Cisco AnyConnect Secure Mobility Agent for Windows"" /f /reg:64"
        
        # Fix desktop shortcut if exists
        Write-Log "FIX Cisco AnyConnect VPN Client desktop shorcut if exists."
        If (Test-Path "C:\Users" -PathType Container -ErrorAction SilentlyContinue) {
            Get-ChildItem -Path "C:\Users" -Include "*" -Force -ErrorAction SilentlyContinue | ForEach-Object ($_) {
                $path1 = $_.FullName + "\Desktop\Cisco AnyConnect VPN Client.lnk"
                $path2 = $_.FullName + "\Desktop\Cisco AnyConnect Secure Mobility Client.lnk"
                $path3 = $_.FullName + "\Desktop"
                If (Test-Path $path1 -PathType Any -ErrorAction SilentlyContinue) {
                    Remove-Item -Force -Path $path1 -ErrorAction SilentlyContinue
                    
                    Write-Log "COPYING new desktop shortcut for Cisco AnyConnect VPN Client."
                    If (Test-Path "C:\Program Files (x86)\Cisco\Cisco AnyConnect Secure Mobility Client\vpnui.exe" -PathType Any) {
                        Copy-Item "$dirFiles\x64\Cisco AnyConnect Secure Mobility Client.lnk" "$path3" -Force -ErrorAction SilentlyContinue
                    } ElseIf (Test-Path "C:\Program Files\Cisco\Cisco AnyConnect Secure Mobility Client\vpnui.exe" -PathType Any) {
                        Copy-Item "$dirFiles\x86\Cisco AnyConnect Secure Mobility Client.lnk" "$path3" -Force -ErrorAction SilentlyContinue
                    } Else {
                        Write-Log "COPYING ERROR cannot find Cisco AnyConnect folder path."
                    }
                }
                If (Test-Path $path2 -PathType Any -ErrorAction SilentlyContinue) {
                    Remove-Item -Force -Path $path2 -ErrorAction SilentlyContinue
                    
                    Write-Log "COPYING new desktop shortcut for Cisco AnyConnect VPN Client."
                    If (Test-Path "C:\Program Files (x86)\Cisco\Cisco AnyConnect Secure Mobility Client\vpnui.exe" -PathType Any) {
                        Copy-Item "$dirFiles\x64\Cisco AnyConnect Secure Mobility Client.lnk" "$path3" -Force -ErrorAction SilentlyContinue
                    } ElseIf (Test-Path "C:\Program Files\Cisco\Cisco AnyConnect Secure Mobility Client\vpnui.exe" -PathType Any) {
                        Copy-Item "$dirFiles\x86\Cisco AnyConnect Secure Mobility Client.lnk" "$path3" -Force -ErrorAction SilentlyContinue
                    } Else {
                        Write-Log "COPYING ERROR cannot find Cisco AnyConnect folder path."
                    }
                }
            }
        }

        Write-Log "Cisco AnyConnect 3.1.13015 install COMPLETE." 
        # Display a message at the end of the install
        Show-InstallationProgress -StatusMessage "Cisco AnyConnect VPN Client 3.1.13015 install COMPLETE. `r`n`r`nPlease wait..." -WindowLocation 'Default' -TopMost $true
        Start-Sleep -s 30


		##*===============================================
		##* POST-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Installation'
		
		## <Perform Post-Installation tasks here>
		
        Write-Log "Cisco AnyConnect 3.1.13015 INSTALL COMPLETED."
        $installNewAnyConnectCount = (Get-InstalledApplication -ProductCode "{EDEB4A62-FE20-4F95-8B90-26BB74CEB6A9}")
        
        If (($installNewAnyConnectCount | Measure-Object).Count -gt 0) {
            # Display a message at the end of the install
            Unblock-AppExecution
            $explorerRunning2 = (Get-Process explorer -ea 0) 
            Write-Log "RUNNING explorer END results: $explorerRunning2"
            If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "Cisco AnyConnect VPN Client 3.1.13015 installation complete.  `r`n`r`nIn order to begin using the new client, please click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then reboot your machine in order to finish the install." -ButtonRightText "OK" -Icon Information -NoWait
                Start-Sleep -s 30
                Exit-Script -ExitCode "3010"
            } Else {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "Cisco AnyConnect VPN Client 3.1.13015 installation complete.  `r`n`r`nIn order to begin using the new client, please reboot your machine at your earliest convenience to complete the installation." -ButtonRightText "OK" -Icon Information -NoWait
                Start-Sleep -s 30
                Exit-Script -ExitCode "3010" 
            }
        } Else {
            # Display a message at the end of the install
            $explorerRunning2 = (Get-Process explorer -ea 0) 
            Write-Log "RUNNING explorer END results: $explorerRunning2"
            If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "Cisco AnyConnect VPN Client 3.1.13015 installation encounter an error.  `r`n`r`nPlease click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then reboot your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait
                Start-Sleep -s 30
                Exit-Script -ExitCode "1"
            } Else {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "Cisco AnyConnect VPN Client 3.1.13015 installation encounter an error.  `r`n`r`nPlease reboot your machine at your earliest convenience in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait
                Start-Sleep -s 30
                Exit-Script -ExitCode "1"
            }
        }

        # End install script
        Write-Log "Cisco AnyConnect 3.1.13015 INSTALL: END SCRIPT"

	}
	ElseIf ($deploymentType -ieq 'Uninstall')
	{
		##*===============================================
		##* PRE-UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Pre-Uninstallation'
		
		## <Perform Pre-Uninstallation tasks here>
		
		##*===============================================
		##* UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Uninstallation'
		
		## <Perform Uninstallation tasks here>
		
        # Start uninstall script
        Write-Log "Cisco AnyConnect 3.1.13015 UNINSTALL: START SCRIPT"
        Write-Log "Computer name: $envComputerName"
        Write-Log "Script running as: $envUserName"
        $systemAccount = "$envComputerName$"
        Write-Log "System account name: $systemAccount"
        
   		# Check for logged on users
        $getUser = (Get-LoggedOnUser)
        Write-Log "LOGGED ON USER results: $getUser"
        $userLoggedOn = "UNKNOWN"
        If ($getUser.ConnectState -eq 'Active') {
            Write-Log "YES, USER(s) is logged on to the computer, ACTIVE."
            $userLoggedOn = "YES"
        } ElseIf ($getUser.ConnectState -eq 'Disconnected') {
            Write-Log "YES, USER(s) is logged on to the computer, DISCONNECTED."
            $userLoggedOn = "YES"
        } Else {
            Write-Log "NO, USER is not logged on to the computer."
            $userLoggedOn = "NO"
        }
        Write-Log "LOGGED ON USER variable results: $userLoggedOn"
        
		# If powerpoint is running in presentation mode abort
	    $presentationPowerPoint = (Test-PowerPoint)
        Write-Log "PowerPoint presentation mode results: $presentationPowerPoint"
        If ($presentationPowerPoint -eq $true) {
            Write-Log "YES, detected PowerPoint in presentation mode, aborting script with exit code 69000."
            Start-Sleep -s 30
            Exit-Script -ExitCode "69000" 
        } Else {
            Write-Log "NO, PowerPoint is NOT in presentation mode."
        }

        # Perform Cisco AnyConnect 3.1.13015 check task here
        $installAnyConnectCount = (Get-InstalledApplication -ProductCode "{EDEB4A62-FE20-4F95-8B90-26BB74CEB6A9}")
        Write-Log "Cisco AnyConnect 3.1.13015 check results: $installAnyConnectCount"
        If (($installAnyConnectCount | Measure-Object).Count -gt 0) {
            Write-Log "YES, Cisco AnyConnect 3.1.13015 is INSTALLED."
            Write-Log "UNINSTALLING Cisco AnyConnect 3.1.13015."     
            
            $explorerRunning1 = (Get-Process explorer -ea 0)
            Write-Log "RUNNING explorer BEGIN results: $explorerRunning1"
            
            Write-Log "START SHOWING UNINSTALLATION PROGRESS MESSAGES."
            Start-Sleep -s 60

            # Stop the AnyConnect service
            Stop-Service -name "vpnagnet" -Force -ErrorAction SilentlyContinue
            Stop-Service -name "Cisco AnyConnect VPN Agent" -Force -ErrorAction SilentlyContinue
            Stop-Service -name "Cisco AnyConnect Secure Mobility Agent" -Force -ErrorAction SilentlyContinue

            # Show Welcome Message, close processes, allow up to 7 day deferral, and persist the prompt
            Show-InstallationWelcome -CloseApps "vpnagent,vpnui,vpndownloader,vacon" -AllowDeferCloseApps -AllowDefer -DeferTimes 25 -DeferDays 7 -PersistPrompt -MinimizeWindows $false
            Start-Sleep -s 30

            # Show Progress Message (with the default message) if any WCWGM processes were running and Show-InstallationWelcome triggered
            Show-InstallationProgress -StatusMessage "UNINSTALLING Cisco AnyConnect VPN Client 3.1.13015. `r`n`r`nThis may take some time. Please wait..." -WindowLocation 'Default' -TopMost $true
            Start-Sleep -s 30
                
            # Perform uninstallation tasks here, uninstalling Cisco AnyConnect
            Remove-MSIApplications "Cisco AnyConnect VPN Client" -ErrorAction SilentlyContinue
            Remove-MSIApplications "Cisco AnyConnect Secure Mobility Client" -ErrorAction SilentlyContinue
            Execute-MSI -Action Uninstall -Path "{7240A69A-AC53-46A1-9039-1281DDBBE452}"
            Execute-MSI -Action Uninstall -Path "{C1EC4E2D-6F63-4806-B88E-7685B6EC186E}"
            Execute-MSI -Action Uninstall -Path "{92083A9A-549D-4057-88E8-223EA08563FA}"
            Execute-MSI -Action Uninstall -Path "{EDEB4A62-FE20-4F95-8B90-26BB74CEB6A9}"
            
            # Create Cisco Profile variables / folders
            $strDefPath1="C:\ProgramData\Cisco\Cisco AnyConnect Secure Mobility Client\Profile"
            $strDefPath2="C:\Documents and Settings\All Users\Application Data\Cisco\Cisco AnyConnect Secure Mobility Client\Profile"
            
            If (Test-Path $strDefPath1 -PathType Any) {
                Remove-Folder -Path $strDefPath1 -ContinueOnError $true
            }

            If (Test-Path $strDefPath2 -PathType Any) {
                Remove-Folder -Path $strDefPath2 -ContinueOnError $true
            }

            Write-Log "REMOVE registry start up of Cisco AnyConnect."
            CMD.EXE /C "C:\Windows\System32\reg.exe delete ""HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"" /v ""Cisco AnyConnect Secure Mobility Agent for Windows"" /f /reg:32"
            CMD.EXE /C "C:\Windows\SysWOW64\reg.exe delete ""HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"" /v ""Cisco AnyConnect Secure Mobility Agent for Windows"" /f /reg:32"
            CMD.EXE /C "C:\Windows\System32\reg.exe delete ""HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"" /v ""Cisco AnyConnect Secure Mobility Agent for Windows"" /f /reg:64"
            CMD.EXE /C "C:\Windows\SysWOW64\reg.exe delete ""HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"" /v ""Cisco AnyConnect Secure Mobility Agent for Windows"" /f /reg:64"
        
            # REMOVE desktop shortcut if exists
            Write-Log "REMOVE Cisco AnyConnect VPN Client desktop shorcut if exists."
            If (Test-Path "C:\Users" -PathType Container -ErrorAction SilentlyContinue) {
                Get-ChildItem -Path "C:\Users" -Include "*" -Force -ErrorAction SilentlyContinue | ForEach-Object ($_) {
                    $path1 = $_.FullName + "\Desktop\Cisco AnyConnect VPN Client.lnk"
                    $path2 = $_.FullName + "\Desktop\Cisco AnyConnect Secure Mobility Client.lnk"
                    If (Test-Path $path1 -PathType Any -ErrorAction SilentlyContinue) {
                        Remove-Item -Force -Path $path1 -ErrorAction SilentlyContinue
                    }
                    If (Test-Path $path2 -PathType Any -ErrorAction SilentlyContinue) {
                        Remove-Item -Force -Path $path2 -ErrorAction SilentlyContinue
                    }
                }
            }
            
            Write-Log "Cisco AnyConnect 3.1.13015 UNINSTALL COMPLETED." 
            # Display a message at the end of the uninstall
            Unblock-AppExecution
            $explorerRunning2 = (Get-Process explorer -ea 0) 
            Write-Log "RUNNING explorer END results: $explorerRunning2"

            # Perform uninstall Cisco AnyConnect 3.1.13015 check #2
            $installAnyConnectCount2 = (Get-InstalledApplication -ProductCode "{EDEB4A62-FE20-4F95-8B90-26BB74CEB6A9}")
            Write-Log "Cisco AnyConnect 3.1.13015 check results #2: $installAnyConnectCount2"
            If (($installAnyConnectCount2 | Measure-Object).Count -gt 0) {
                If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
					Show-InstallationPrompt -Message "Cisco AnyConnect VPN Client 3.1.13015 uninstallation encounter an error.  `r`n`r`nPlease click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then reboot your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait
					Start-Sleep -s 30
                    Exit-Script -ExitCode "1"
                } Else {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
					Show-InstallationPrompt -Message "Cisco AnyConnect VPN Client 3.1.13015 uninstallation encounter an error.  `r`n`r`nPlease reboot your machine at your earliest convenience in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait
					Start-Sleep -s 30
                    Exit-Script -ExitCode "1"
                }
            } Else {
                If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
					Show-InstallationPrompt -Message "Cisco AnyConnect VPN Client 3.1.13015 uninstallation complete.  `r`n`r`nPlease click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then reboot your machine in order to finish the uninstall." -ButtonRightText "OK" -Icon Information -NoWait					
					Start-Sleep -s 30
                    Exit-Script -ExitCode "3010"
                } Else {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "Cisco AnyConnect VPN Client 3.1.13015 uninstallation complete.  `r`n`r`nPlease reboot your machine at your earliest convenience to complete the uninstall." -ButtonRightText "OK" -Icon Information -NoWait
					Start-Sleep -s 30
                    Exit-Script -ExitCode "3010" 
                }
            }
        } Else {
            Write-Log "NO, Cisco AnyConnect VPN Client 3.1.13015 is NOT installed."
            Show-InstallationPrompt -Message "Cisco AnyConnect VPN Client 3.1.13015 NOT installed, nothing to uninstall. NO changes made. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
            Start-Sleep -s 30
            Exit-Script -ExitCode "0"
        }

        # End uninstall script
        Write-Log "Cisco AnyConnect VPN Client 3.1.13015 UNINSTALL: END SCRIPT"
		
		##*===============================================
		##* POST-UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Uninstallation'
		
		## <Perform Post-Uninstallation tasks here>
		
	}
	
	##*===============================================
	##* END SCRIPT BODY
	##*===============================================
	
	## Call the Exit-Script function to perform final cleanup operations
	Exit-Script -ExitCode $mainExitCode
}
Catch {
	[int32]$mainExitCode = 60001
	[string]$mainErrorMessage = "$(Resolve-Error)"
	Write-Log -Message $mainErrorMessage -Severity 3 -Source $deployAppScriptFriendlyName
	Show-DialogBox -Text $mainErrorMessage -Icon 'Stop'
	Exit-Script -ExitCode $mainExitCode
}