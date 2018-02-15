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
	Deploy-Application.ps1
.EXAMPLE
	Deploy-Application.ps1 -DeployMode 'Silent'
.EXAMPLE
	Deploy-Application.ps1 -AllowRebootPassThru -AllowDefer
.EXAMPLE
	Deploy-Application.ps1 -DeploymentType Uninstall
.NOTES
	Toolkit Exit Code Ranges:
	60000 - 68999: Reserved for built-in exit codes in Deploy-Application.ps1, Deploy-Application.exe, and AppDeployToolkitMain.ps1
	69000 - 69999: Recommended for user customized exit codes in Deploy-Application.ps1
	70000 - 79999: Recommended for user customized exit codes in AppDeployToolkitExtensions.ps1
.LINK 
	http://psappdeploytoolkit.codeplex.com
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
	[string]$appVendor = ''
	[string]$appName = '7-Zip 9.20 64 Bit'
	[string]$appVersion = ''
	[string]$appArch = 'x64'
	[string]$appLang = 'EN'
	[string]$appRevision = '01'
	[string]$appScriptVersion = '1.0.0'
	[string]$appScriptDate = '09/14/2015'
	[string]$appScriptAuthor = ''
	##*===============================================
	
	##* Do not modify section below
	#region DoNotModify
	
	## Variables: Exit Code
	[int32]$mainExitCode = 0
	
	## Variables: Script
	[string]$deployAppScriptFriendlyName = 'Deploy Application'
	[version]$deployAppScriptVersion = [version]'3.6.1'
	[string]$deployAppScriptDate = '03/26/2015'
	[hashtable]$deployAppScriptParameters = $psBoundParameters
	
	## Variables: Environment
	[string]$scriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
	
	## Dot source the required App Deploy Toolkit Functions
	Try {
		[string]$moduleAppDeployToolkitMain = "$scriptDirectory\AppDeployToolkit\AppDeployToolkitMain.ps1"
		If (-not (Test-Path -Path $moduleAppDeployToolkitMain -PathType Leaf)) { Throw "Module does not exist at the specified location [$moduleAppDeployToolkitMain]." }
		If ($DisableLogging) { . $moduleAppDeployToolkitMain -DisableLogging } Else { . $moduleAppDeployToolkitMain }
	}
	Catch {
		[int32]$mainExitCode = 60008
		Write-Error -Message "Module [$moduleAppDeployToolkitMain] failed to load: `n$($_.Exception.Message)`n `n$($_.InvocationInfo.PositionMessage)" -ErrorAction 'Continue'
		Exit $mainExitCode
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
        Write-Log "7-Zip 9.20 64 Bit INSTALL: START SCRIPT"
        Write-Log "Computer name: $envComputerName"
        Write-Log "Script running as: $envUserName"
        $systemAccount = "$envComputerName$"
        Write-Log "System account name: $systemAccount"

        # Check for logged on users
        $getUser = (Get-LoggedOnUser)
        Write-Log "Logged on user results: $getUser"
        If ($getUser.ConnectState -eq 'Active') {
            Write-Log "YES, USER(s) is logged on to the computer, ACTIVE."
        } ElseIf ($getUser.ConnectState -eq 'Disconnected') {
            Write-Log "YES, USER(s) is logged on to the computer, DISCONNECTED."
        } Else {
            Write-Log "NO, USER is not logged on to the computer."
        }
        
		# If powerpoint is running in presentation mode abort
	    $presentationPowerPoint = (Test-PowerPoint)
        Write-Log "PowerPoint presentation mode results: $presentationPowerPoint"
        If ($presentationPowerPoint -eq $true) {
            Write-Log "YES, detected PowerPoint in presentation mode, aborting script with exit code 5001."
            Exit-Script -ExitCode "5001" 
        } Else {
            Write-Log "NO, PowerPoint is NOT in presentation mode."
        }

        # Perform 7-Zip 9.20 64 Bit check task here
        $install7zipcount = (Get-InstalledApplication -ProductCode "{23170F69-40C1-2702-0920-000001000000}")
        Write-Log "7-Zip 9.20 64 Bit check results: $install7zipcount"
        If (($install7zipcount | Measure-Object).Count -gt 0) {
            Write-Log "YES, 7-Zip 9.20 64 Bit is INSTALLED, aborting script with exit code 0."
            #Show-InstallationPrompt -Message "7-Zip 9.20 64 Bit is previously INSTALLED. NO changes made. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
            Start-Sleep -s 7
            Exit-Script -ExitCode "0" 
        } Else {
            Write-Log "NO, 7-Zip 9.20 64 Bit is NEEDED."             
        }		

	    $explorerRunning1 = (Get-Process explorer -ea 0) 
        Write-Log "RUNNING explorer BEGIN results: $explorerRunning1"
        
        # Show Welcome Message, close processes, allow up to 7 day deferral, and persist the prompt
        If ($getUser.ConnectState -eq 'Active') {
            Write-Log "YES, USER(s) is logged on to the computer, ACTIVE."
            Show-InstallationWelcome -CloseApps "7zFM" -AllowDeferCloseApps -AllowDefer -DeferTimes 25 -DeferDays 7 -PersistPrompt -MinimizeWindows $false
        } ElseIf ($getUser.ConnectState -eq 'Disconnected') {
            Write-Log "YES, USER(s) is logged on to the computer, DISCONNECTED."
            Show-InstallationWelcome -CloseApps "7zFM" -Silent -MinimizeWindows $false
        } Else {
            Write-Log "NO, USER is not logged on to the computer."
            Show-InstallationWelcome -CloseApps "7zFM" -Silent -MinimizeWindows $false
        }
                        		
		##*===============================================
		##* INSTALLATION 
		##*===============================================
		[string]$installPhase = 'Installation'

        Show-InstallationProgress -StatusMessage "Running PRE-INSTALLATION steps for 7-Zip 9.20 64 Bit. `r`n`r`nThis may take some time. Please wait..." -WindowLocation 'Default' -TopMost $true        
        Write-Log "START Removal of other versions of 7-Zip"

        #Remove other versions
        $strUninstallPath1 = "C:\Program Files (x86)\7-Zip\Uninstall.exe"
        $strUninstallPath2 = "D:\Program Files (x86)\7-Zip\Uninstall.exe"
        $strUninstallPath3 = "C:\Program Files\7-Zip\Uninstall.exe"

        Execute-MSI -Action Uninstall -Path "{23170F69-40C1-2702-0922-000001000000}"
        Execute-MSI -Action Uninstall -Path "{23170F69-40C1-2702-0938-000001000000}"

        If (Test-Path $strUninstallPath1 -PathType Any) {
            Execute-Process -FilePath $strUninstallPath1 -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010"
            Write-Log "RUN $strUninstallPath1"
            }
        ElseIf (Test-Path $strUninstallPath2 -PathType Any) {
            Execute-Process -FilePath $strUninstallPath2 -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010"
            Write-Log "RUN $strUninstallPath2"
            }
        ElseIf (Test-Path $strUninstallPath3 -PathType Any) {
            Execute-Process -FilePath $strUninstallPath3 -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010"
            Write-Log "RUN $strUninstallPath3"
            }
        Else    {
            Write-Log "ERROR: Either cannot find or cannot remove old version of 7-ZIP."
            }
		
		## <Perform Installation tasks here>

        Write-Log "START installing 7-Zip 9.20 64 Bit."
        # Show Progress Message (with the default message)
        Show-InstallationProgress -StatusMessage "INSTALLING 7-Zip 9.20 64 Bit. `r`n`r`nThis may take some time. Please wait..." -WindowLocation 'Default' -TopMost $true        
        Start-Sleep -s 7
        
        # Perform installation tasks here
        Write-Log "INSTALLING 7-Zip 9.20 64 Bit."

        Execute-MSI -Action Install -Path "$dirFiles\7z920-x64.msi" -Parameters "/quiet /passive /norestart /qn INSTALLDIR=""C:\Program Files\7-Zip"""
        
        Write-Log "FIX 7-zip default archive format to zip."
        [scriptblock]$HKCURegistrySettings = {
            Set-RegistryKey -Key 'HKCU\Software\7-Zip\Compression' -Name 'Archiver' -Value 'zip' -Type String -SID $UserProfile.SID -ContinueOnError $true
        }
        Invoke-HKCURegistrySettingsForAllUsers -RegistrySettings $HKCURegistrySettings
        
        Write-Log "7-Zip 9.20 64 Bit install COMPLETE." 
        # Display a message at the end of the install
        Show-InstallationProgress -StatusMessage "7-Zip 9.20 64 Bit install COMPLETE. `r`n`r`nPlease wait..." -WindowLocation 'Default' -TopMost $true
        Start-Sleep -s 7
                        		
		##*===============================================
		##* POST-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Installation'
		
		## <Perform Post-Installation tasks here>

        Write-Log "7-Zip 9.20 64 Bit INSTALL COMPLETED." 
        
        # Display a message at the end of the install
        Unblock-AppExecution
        $explorerRunning2 = (Get-Process explorer -ea 0) 
        Write-Log "RUNNING explorer END results: $explorerRunning2"
        If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
            If ($installSuccess = $true) {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                If ($getUser.ConnectState -eq 'Active') {
                    Write-Log "YES, USER(s) is logged on to the computer, ACTIVE."
                    Show-InstallationPrompt -Message "COMPLETE: 7-Zip 9.20 64 Bit install COMPLETE. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to finish the install." -ButtonRightText "OK" -Icon Information -NoWait
                } ElseIf ($getUser.ConnectState -eq 'Disconnected') {
                    Write-Log "YES, USER(s) is logged on to the computer, DISCONNECTED."
                    Show-InstallationPrompt -Message "COMPLETE: 7-Zip 9.20 64 Bit install COMPLETE. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to finish the install." -ButtonRightText "OK" -Icon Information -NoWait -Timeout 300
                } Else {
                    Write-Log "NO, USER is not logged on to the computer."
                    Show-InstallationPrompt -Message "COMPLETE: 7-Zip 9.20 64 Bit install COMPLETE. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to finish the install." -ButtonRightText "OK" -Icon Information -NoWait -Timeout 300
                }
                Start-Sleep -s 10
            } ElseIf ($installSuccess = $false) {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                If ($getUser.ConnectState -eq 'Active') {
                    Write-Log "YES, USER(s) is logged on to the computer, ACTIVE."
                    Show-InstallationPrompt -Message "FAILED: 7-Zip 9.20 64 Bit install FAILED. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait
                } ElseIf ($getUser.ConnectState -eq 'Disconnected') {
                    Write-Log "YES, USER(s) is logged on to the computer, DISCONNECTED."
                    Show-InstallationPrompt -Message "FAILED: 7-Zip 9.20 64 Bit install FAILED. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait -Timeout 300
                } Else {
                    Write-Log "NO, USER is not logged on to the computer."
                    Show-InstallationPrompt -Message "FAILED: 7-Zip 9.20 64 Bit install FAILED. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait -Timeout 300
                }
                Start-Sleep -s 10
            } Else {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                If ($getUser.ConnectState -eq 'Active') {
                    Write-Log "YES, USER(s) is logged on to the computer, ACTIVE."
                    Show-InstallationPrompt -Message "ERROR: 7-Zip 9.20 64 Bit install encounter an ERROR. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait
                } ElseIf ($getUser.ConnectState -eq 'Disconnected') {
                    Write-Log "YES, USER(s) is logged on to the computer, DISCONNECTED."
                    Show-InstallationPrompt -Message "ERROR: 7-Zip 9.20 64 Bit install encounter an ERROR. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait -Timeout 300
                } Else {
                    Write-Log "NO, USER is not logged on to the computer."
                    Show-InstallationPrompt -Message "ERROR: 7-Zip 9.20 64 Bit install encounter an ERROR. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait -Timeout 300
                }
                Start-Sleep -s 10
            }
        } Else {
            If ($installSuccess = $true) {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                If ($getUser.ConnectState -eq 'Active') {
                    Write-Log "YES, USER(s) is logged on to the computer, ACTIVE."
                    Show-InstallationPrompt -Message "COMPLETE: 7-Zip 9.20 64 Bit install COMPLETE. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
                } ElseIf ($getUser.ConnectState -eq 'Disconnected') {
                    Write-Log "YES, USER(s) is logged on to the computer, DISCONNECTED."
                    Show-InstallationPrompt -Message "COMPLETE: 7-Zip 9.20 64 Bit install COMPLETE. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait -Timeout 300
                } Else {
                    Write-Log "NO, USER is not logged on to the computer."
                    Show-InstallationPrompt -Message "COMPLETE: 7-Zip 9.20 64 Bit install COMPLETE. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait -Timeout 300
                }

                Start-Sleep -s 10
            } ElseIf ($installSuccess = $false) {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                If ($getUser.ConnectState -eq 'Active') {
                    Write-Log "YES, USER(s) is logged on to the computer, ACTIVE."
                    Show-InstallationPrompt -Message "FAILED: 7-Zip 9.20 64 Bit install FAILED. `r`n`r`nACTION REQUIRED: Click OK and then REBOOT / RESTART your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait
                } ElseIf ($getUser.ConnectState -eq 'Disconnected') {
                    Write-Log "YES, USER(s) is logged on to the computer, DISCONNECTED."
                    Show-InstallationPrompt -Message "FAILED: 7-Zip 9.20 64 Bit install FAILED. `r`n`r`nACTION REQUIRED: Click OK and then REBOOT / RESTART your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait -Timeout 300
                } Else {
                    Write-Log "NO, USER is not logged on to the computer."
                    Show-InstallationPrompt -Message "FAILED: 7-Zip 9.20 64 Bit install FAILED. `r`n`r`nACTION REQUIRED: Click OK and then REBOOT / RESTART your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait -Timeout 300
                }
                Start-Sleep -s 10
            } Else {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                If ($getUser.ConnectState -eq 'Active') {
                    Write-Log "YES, USER(s) is logged on to the computer, ACTIVE."
                    Show-InstallationPrompt -Message "ERROR: 7-Zip 9.20 64 Bit install encounter an ERROR. `r`n`r`nACTION REQUIRED: Click OK and then REBOOT / RESTART your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait
                } ElseIf ($getUser.ConnectState -eq 'Disconnected') {
                    Write-Log "YES, USER(s) is logged on to the computer, DISCONNECTED."
                    Show-InstallationPrompt -Message "ERROR: 7-Zip 9.20 64 Bit install encounter an ERROR. `r`n`r`nACTION REQUIRED: Click OK and then REBOOT / RESTART your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait -Timeout 300
                } Else {
                    Write-Log "NO, USER is not logged on to the computer."
                    Show-InstallationPrompt -Message "ERROR: 7-Zip 9.20 64 Bit install encounter an ERROR. `r`n`r`nACTION REQUIRED: Click OK and then REBOOT / RESTART your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait -Timeout 300
                }
                Start-Sleep -s 10
            }
        }
		
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
        Write-Log "7-Zip 9.20 64 Bit UNINSTALL: START SCRIPT"
        Write-Log "Computer name: $envComputerName"
        Write-Log "Script running as: $envUserName"
        $systemAccount = "$envComputerName$"
        Write-Log "System account name: $systemAccount"

   		# Check for logged on users
        $getUser = (Get-LoggedOnUser)
        Write-Log "Logged on user results: $getUser"
        If ($getUser.ConnectState -eq 'Active') {
            Write-Log "YES, USER(s) is logged on to the computer, ACTIVE."
        } ElseIf ($getUser.ConnectState -eq 'Disconnected') {
            Write-Log "YES, USER(s) is logged on to the computer, DISCONNECTED."
        } Else {
            Write-Log "NO, USER is not logged on to the computer."
        }
        
        # Perform uninstall 7-Zip 9.20 64 Bit check
        $installOld7Zipcount = (Get-InstalledApplication -ProductCode "{23170F69-40C1-2702-0920-000001000000}")
        Write-Log "7-Zip 9.20 64 Bit check results: $installOld7Zipcount"
        If (($installOld7Zipcount | Measure-Object).Count -gt 0) {
            Write-Log "YES, 7-Zip 9.20 64 Bit is INSTALLED."
            Write-Log "UNINSTALLING 7-Zip 9.20 64 Bit."     
            
            $explorerRunning1 = (Get-Process explorer -ea 0) 
            Write-Log "RUNNING explorer BEGIN results: $explorerRunning1"
            
            # Show Welcome Message, close processes, allow up to 7 day deferral, and persist the prompt
            If ($getUser.ConnectState -eq 'Active') {
                Write-Log "YES, USER(s) is logged on to the computer, ACTIVE."
                Show-InstallationWelcome -CloseApps "7zFM" -AllowDeferCloseApps -AllowDefer -DeferTimes 25 -DeferDays 7 -PersistPrompt -MinimizeWindows $false
            } ElseIf ($getUser.ConnectState -eq 'Disconnected') {
                Write-Log "YES, USER(s) is logged on to the computer, DISCONNECTED."
                Show-InstallationWelcome -CloseApps "7zFM" -Silent -MinimizeWindows $false
            } Else {
                Write-Log "NO, USER is not logged on to the computer."
                Show-InstallationWelcome -CloseApps "7zFM" -Silent -MinimizeWindows $false
            }

            # Show Progress Message (with the default message)
            Show-InstallationProgress -StatusMessage "UNINSTALLING 7-Zip 9.20 64 Bit. `r`n`r`nThis may take some time. Please wait..." -WindowLocation 'Default' -TopMost $true
            Start-Sleep -s 7
                
            # Perform uninstallation tasks here
            Execute-MSI -Action Uninstall -Path "{23170F69-40C1-2702-0920-000001000000}"
            
            Write-Log "REMOVE 7-Zip 9.20 64 Bit"
            Remove-RegistryKey -Key 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{23170F69-40C1-2702-0920-000001000000}' -Recurse -ContinueOnError $true

            Write-Log "REMOVE 7-Zip Registry for ALL USERS."
            [scriptblock]$HKCURegistrySettings = {
                Remove-RegistryKey -Key 'HKCU\Software\7-Zip' -Recurse -SID $UserProfile.SID -ContinueOnError $true
            }
            Invoke-HKCURegistrySettingsForAllUsers -RegistrySettings $HKCURegistrySettings

            # Check to see if the folder exists, if so delete folder
#            $str7ZipPath1="C:\Program Files\7-Zip"
#            $str7ZipPath2="C:\Program Files (x86)\7-Zip"
#            If (Test-Path $str7ZipPath1 -PathType Any) {
#                Remove-Folder -Path $str7ZipPath1 -ContinueOnError $true
#            }
#            If (Test-Path $str7ZipPath2 -PathType Any) {
#                Remove-Folder -Path $str7ZipPath2 -ContinueOnError $true
#            }

            Write-Log "7-Zip 9.20 64 Bit UNINSTALL COMPLETED." 
            # Display a message at the end of the install
            Unblock-AppExecution
            $explorerRunning2 = (Get-Process explorer -ea 0) 
            Write-Log "RUNNING explorer END results: $explorerRunning2"
            If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
                If ($installSuccess = $true) {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    If ($getUser.ConnectState -eq 'Active') {
                        Write-Log "YES, USER(s) is logged on to the computer, ACTIVE."
                        Show-InstallationPrompt -Message "COMPLETE: 7-Zip 9.20 64 Bit uninstall COMPLETE. ALL DONE. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to finish the uninstall." -ButtonRightText "OK" -Icon Information -NoWait
                    } ElseIf ($getUser.ConnectState -eq 'Disconnected') {
                        Write-Log "YES, USER(s) is logged on to the computer, DISCONNECTED."
                        Show-InstallationPrompt -Message "COMPLETE: 7-Zip 9.20 64 Bit uninstall COMPLETE. ALL DONE. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to finish the uninstall." -ButtonRightText "OK" -Icon Information -NoWait -Timeout 300
                    } Else {
                        Write-Log "NO, USER is not logged on to the computer."
                        Show-InstallationPrompt -Message "COMPLETE: 7-Zip 9.20 64 Bit uninstall COMPLETE. ALL DONE. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to finish the uninstall." -ButtonRightText "OK" -Icon Information -NoWait -Timeout 300
                    }
                    Start-Sleep -s 10
                } ElseIf ($installSuccess = $false) {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    If ($getUser.ConnectState -eq 'Active') {
                        Write-Log "YES, USER(s) is logged on to the computer, ACTIVE."
                        Show-InstallationPrompt -Message "FAILED: 7-Zip 9.20 64 Bit uninstall FAILED. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait
                    } ElseIf ($getUser.ConnectState -eq 'Disconnected') {
                        Write-Log "YES, USER(s) is logged on to the computer, DISCONNECTED."
                        Show-InstallationPrompt -Message "FAILED: 7-Zip 9.20 64 Bit uninstall FAILED. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait -Timeout 300
                    } Else {
                        Write-Log "NO, USER is not logged on to the computer."
                        Show-InstallationPrompt -Message "FAILED: 7-Zip 9.20 64 Bit uninstall FAILED. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait -Timeout 300
                    }
                    Start-Sleep -s 10
                } Else {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    If ($getUser.ConnectState -eq 'Active') {
                        Write-Log "YES, USER(s) is logged on to the computer, ACTIVE."
                        Show-InstallationPrompt -Message "ERROR: 7-Zip 9.20 64 Bit uninstall encounter an ERROR. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait
                    } ElseIf ($getUser.ConnectState -eq 'Disconnected') {
                        Write-Log "YES, USER(s) is logged on to the computer, DISCONNECTED."
                        Show-InstallationPrompt -Message "ERROR: 7-Zip 9.20 64 Bit uninstall encounter an ERROR. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait -Timeout 300
                    } Else {
                        Write-Log "NO, USER is not logged on to the computer."
                        Show-InstallationPrompt -Message "ERROR: 7-Zip 9.20 64 Bit uninstall encounter an ERROR. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait -Timeout 300
                    }
                    Start-Sleep -s 10
                }
            } Else {
                If ($installSuccess = $true) {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    If ($getUser.ConnectState -eq 'Active') {
                        Write-Log "YES, USER(s) is logged on to the computer, ACTIVE."
                        Show-InstallationPrompt -Message "COMPLETE: 7-Zip 9.20 64 Bit uninstall COMPLETE. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
                    } ElseIf ($getUser.ConnectState -eq 'Disconnected') {
                        Write-Log "YES, USER(s) is logged on to the computer, DISCONNECTED."
                        Show-InstallationPrompt -Message "COMPLETE: 7-Zip 9.20 64 Bit uninstall COMPLETE. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait -Timeout 300
                    } Else {
                        Write-Log "NO, USER is not logged on to the computer."
                        Show-InstallationPrompt -Message "COMPLETE: 7-Zip 9.20 64 Bit uninstall COMPLETE. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait -Timeout 300
                    }
                    Start-Sleep -s 10
                } ElseIf ($installSuccess = $false) {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    If ($getUser.ConnectState -eq 'Active') {
                        Write-Log "YES, USER(s) is logged on to the computer, ACTIVE."
                        Show-InstallationPrompt -Message "FAILED: 7-Zip 9.20 64 Bit uninstall FAILED. `r`n`r`nACTION REQUIRED: Click OK and then REBOOT / RESTART your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait
                    } ElseIf ($getUser.ConnectState -eq 'Disconnected') {
                        Write-Log "YES, USER(s) is logged on to the computer, DISCONNECTED."
                        Show-InstallationPrompt -Message "FAILED: 7-Zip 9.20 64 Bit uninstall FAILED. `r`n`r`nACTION REQUIRED: Click OK and then REBOOT / RESTART your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait -Timeout 300
                    } Else {
                        Write-Log "NO, USER is not logged on to the computer."
                        Show-InstallationPrompt -Message "FAILED: 7-Zip 9.20 64 Bit uninstall FAILED. `r`n`r`nACTION REQUIRED: Click OK and then REBOOT / RESTART your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait -Timeout 300
                    }
                    Start-Sleep -s 10
                } Else {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    If ($getUser.ConnectState -eq 'Active') {
                        Write-Log "YES, USER(s) is logged on to the computer, ACTIVE."
                        Show-InstallationPrompt -Message "ERROR: 7-Zip 9.20 64 Bit uninstall encounter an ERROR. `r`n`r`nACTION REQUIRED: Click OK and then REBOOT / RESTART your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait
                    } ElseIf ($getUser.ConnectState -eq 'Disconnected') {
                        Write-Log "YES, USER(s) is logged on to the computer, DISCONNECTED."
                        Show-InstallationPrompt -Message "ERROR: 7-Zip 9.20 64 Bit uninstall encounter an ERROR. `r`n`r`nACTION REQUIRED: Click OK and then REBOOT / RESTART your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait -Timeout 300
                    } Else {
                        Write-Log "NO, USER is not logged on to the computer."
                        Show-InstallationPrompt -Message "ERROR: 7-Zip 9.20 64 Bit uninstall encounter an ERROR. `r`n`r`nACTION REQUIRED: Click OK and then REBOOT / RESTART your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait -Timeout 300
                    }
                    Start-Sleep -s 10
                }
            }
        } Else {
            Write-Log "NO, 7-Zip 9.20 64 Bit is NOT installed."
            If ($getUser.ConnectState -eq 'Active') {
                Write-Log "YES, USER(s) is logged on to the computer, ACTIVE."
                Show-InstallationPrompt -Message "COMPLETE: 7-Zip 9.20 64 Bit NOT installed, nothing to uninstall. NO changes made. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
            } ElseIf ($getUser.ConnectState -eq 'Disconnected') {
                Write-Log "YES, USER(s) is logged on to the computer, DISCONNECTED."
                Show-InstallationPrompt -Message "COMPLETE: 7-Zip 9.20 64 Bit NOT installed, nothing to uninstall. NO changes made. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait -Timeout 300
            } Else {
                Write-Log "NO, USER is not logged on to the computer."
                Show-InstallationPrompt -Message "COMPLETE: 7-Zip 9.20 64 Bit NOT installed, nothing to uninstall. NO changes made. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait -Timeout 300
            }
            Start-Sleep -s 7
        }
				
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