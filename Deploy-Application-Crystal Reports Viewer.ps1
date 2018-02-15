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
	[string]$appVendor = 'SAP'
	[string]$appName = 'Crystal Reports Viewer'
	[string]$appVersion = '2013 SP3'
	[string]$appArch = 'x86'
	[string]$appLang = 'EN'
	[string]$appRevision = '01'
	[string]$appScriptVersion = '1.0.0'
	[string]$appScriptDate = '10/08/2015'
	[string]$appScriptAuthor = ''
	##*===============================================
	
	##* Do not modify section below
	#region DoNotModify
	
	## Variables: Exit Code
	[int32]$mainExitCode = 0
	
	## Variables: Script
	[string]$deployAppScriptFriendlyName = 'Deploy Application'
	[version]$deployAppScriptVersion = [version]'3.6.7'
	[string]$deployAppScriptDate = '09/22/2015'
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
        Write-Log "Crystal Reports Viewer 2013 SP3 INSTALL: START SCRIPT"
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

        # Perform Crystal Reports Viewer 2013 SP3 check
        $installCRViewerCount = (Get-InstalledApplication -Name "SAP Crystal Reports 2013 viewer SP3")
        Write-Log "Crystal Reports Viewer 2013 SP3 check results: $installCRViewerCount"
        If (($installCRViewerCount | Measure-Object).Count -gt 0) {
            Write-Log "YES, Crystals Report Viewer 2013 SP3 is INSTALLED."
            Show-InstallationPrompt -Message "Crystal Report Viewer 2013 SP3 is previously INSTALLED. NO changes made. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
            Write-Log "Crystal Reports Viewer 2013 SP3 is previously INSTALLED. NO changes made."
            Start-Sleep -s 10
            Exit-Script -ExitCode "0" 
        } Else {
            Write-Log "NO, Crystal Reports Viewer 2013 SP3 is NEEDED."
        }		

        $explorerRunning1 = (Get-Process explorer -ea 0) 
        Write-Log "RUNNING explorer BEGIN results: $explorerRunning1"
                
        # Show Welcome Message, close processes, allow up to 7 day deferral, and persist the prompt
        Show-InstallationWelcome -CloseApps "CrystalReportsViewer" -AllowDeferCloseApps -AllowDefer -DeferTimes 25 -DeferDays 7 -PersistPrompt -MinimizeWindows $false
        
        ##*===============================================
		##* INSTALLATION 
		##*===============================================
		[string]$installPhase = 'Installation'
		
		## <Perform Installation tasks here>

        Write-Log "START installing Crystal Reports Viewer 2013 SP3."
        # Show Progress Message (with the default message) and Show-InstallationWelcome triggered
        Show-InstallationProgress -StatusMessage "INSTALLING Crystal Reports Viewer 2013 SP3. `r`n`r`nThis may take some time. Please wait..."        
        Start-Sleep -s 10
        
        If (Test-Path "C:\Program Files\Crystal Reports Viewer" -PathType Any -ErrorAction SilentlyContinue) {
            Remove-Folder -Path "C:\Program Files\Crystal Reports Viewer" -ContinueOnError $true
        }
        If (Test-Path "C:\Program Files (x86)\Crystal Reports Viewer" -PathType Any -ErrorAction SilentlyContinue) {
            Remove-Folder -Path "C:\Program Files (x86)\Crystal Reports Viewer" -ContinueOnError $true
        }
        [scriptblock]$HKCURegistrySettings = {
            Remove-RegistryKey -Key 'HKCU\Software\SAP BusinessObjects\Suite XI 4.0' -Recurse -SID $UserProfile.SID -ContinueOnError $true
            Remove-RegistryKey -Key 'HKCU\Software\Wow6432Node\SAP BusinessObjects\Suite XI 4.0' -Recurse -SID $UserProfile.SID -ContinueOnError $true
        }
        Invoke-HKCURegistrySettingsForAllUsers -RegistrySettings $HKCURegistrySettings
        Remove-RegistryKey -Key 'HKLM\Software\SAP BusinessObjects\Suite XI 4.0' -Recurse -ContinueOnError $true
        Remove-RegistryKey -Key 'HKLM\Software\Wow6432Node\SAP BusinessObjects\Suite XI 4.0' -Recurse -ContinueOnError $true

        # Perform installation tasks here
        Write-Log "INSTALLING Crystal Reports Viewer 2013 SP3."
        Execute-Process -FilePath "$dirFiles\setup.exe" -Arguments "-q -r ""$dirFiles\response.ini""" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $false

        If (Test-Path "C:\Users\Public\Desktop" -PathType Container -ErrorAction SilentlyContinue) {
            Remove-Item -Force -Path "C:\Users\Public\Desktop\SAP Crystal Reports 2013 viewer.lnk" -ErrorAction SilentlyContinue
        }
        Get-ChildItem -Path "C:\Users" -Include "*" -Force -ErrorAction SilentlyContinue | ForEach-Object ($_) {
            $path0 = $_.FullName + "\Desktop"
            $path1 = $_.FullName + "\Desktop\SAP Crystal Reports 2013 viewer.lnk"
            If (Test-Path $path0 -PathType Container -ErrorAction SilentlyContinue) {
                Remove-Item -Force -Path $path1 -ErrorAction SilentlyContinue
            }
        }
        
        ##*===============================================
		##* POST-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Installation'
		
		## <Perform Post-Installation tasks here>
        
        Write-Log "Crystal Reports Viewer 2013 SP3 INSTALL COMPLETED." 
        
        # Display a message at the end of the install
        Unblock-AppExecution
        $explorerRunning2 = (Get-Process explorer -ea 0) 
        Write-Log "RUNNING explorer END results: $explorerRunning2"
        If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
            If ($installSuccess = $true) {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "COMPLETE: Crystal Reports Viewer 2013 SP3 install COMPLETE. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to finish the install." -ButtonRightText "OK" -Icon Information -NoWait
                Start-Sleep -s 10
            } ElseIf ($installSuccess = $false) {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "FAILED: Crystal Reports Viewer 2013 SP3 install FAILED. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait
                Start-Sleep -s 10
            } Else {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "ERROR: Crystal Reports Viewer 2013 SP3 install encounter an ERROR. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait
                Start-Sleep -s 10
            }
        } Else {
            If ($installSuccess = $true) {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "COMPLETE: Crystal Reports Viewer 2013 SP3 install COMPLETE. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
                Start-Sleep -s 10
            } ElseIf ($installSuccess = $false) {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "FAILED: Crystal Reports Viewer 2013 SP3 install FAILED. `r`n`r`nACTION REQUIRED: Click OK and then REBOOT / RESTART your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait
                Start-Sleep -s 10
            } Else {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "ERROR: Crystal Reports Viewer 2013 SP3 install encounter an ERROR. `r`n`r`nACTION REQUIRED: Click OK and then REBOOT / RESTART your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait
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
        
        # Start install script
        Write-Log "Crystal Reports Viewer 2013 SP3 UNINSTALL: START SCRIPT"
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
        
        # Perform Crystal Reports Viewer 2013 SP3 check
        $installCRViewerCount = (Get-InstalledApplication -Name "SAP Crystal Reports 2013 viewer SP3")
        Write-Log "Crystal Reports Viewer 2013 SP3 check results: $installCRViewerCount"
        If (($installCRViewerCount | Measure-Object).Count -gt 0) {
            Write-Log "YES, Crystal Reports Viewer 2013 SP3 is INSTALLED."
            
            $explorerRunning1 = (Get-Process explorer -ea 0) 
            Write-Log "RUNNING explorer BEGIN results: $explorerRunning1"
            
            # Show Welcome Message, close processes, allow up to 7 day deferral, and persist the prompt
            Show-InstallationWelcome -CloseApps "CrystalReportsViewer" -AllowDeferCloseApps -AllowDefer -DeferTimes 25 -DeferDays 7 -PersistPrompt -MinimizeWindows $false
            Show-InstallationProgress -StatusMessage "UNINSTALLING Crystal Reports Viewer 2013 SP3. `r`n`r`nThis may take some time. Please wait..." -WindowLocation 'Default' -TopMost $true
            Start-Sleep -s 10
                                   
            Write-Log "UNINSTALLING Crystal Reports Viewer 2013 SP3."
            Execute-Process -FilePath "$dirFiles\setup.exe" -Arguments "-q -r ""$dirFiles\uninstall.ini""" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $false
            Start-Sleep -s 10
            
            If (Test-Path "C:\Program Files\Crystal Reports Viewer" -PathType Any -ErrorAction SilentlyContinue) {
                Remove-Folder -Path "C:\Program Files\Crystal Reports Viewer" -ContinueOnError $true
            }
            If (Test-Path "C:\Program Files (x86)\Crystal Reports Viewer" -PathType Any -ErrorAction SilentlyContinue) {
                Remove-Folder -Path "C:\Program Files (x86)\Crystal Reports Viewer" -ContinueOnError $true
            }
            [scriptblock]$HKCURegistrySettings = {
                Remove-RegistryKey -Key 'HKCU\Software\SAP BusinessObjects\Suite XI 4.0' -Recurse -SID $UserProfile.SID -ContinueOnError $true
                Remove-RegistryKey -Key 'HKCU\Software\Wow6432Node\SAP BusinessObjects\Suite XI 4.0' -Recurse -SID $UserProfile.SID -ContinueOnError $true
            }
            Invoke-HKCURegistrySettingsForAllUsers -RegistrySettings $HKCURegistrySettings
            Remove-RegistryKey -Key 'HKLM\Software\SAP BusinessObjects\Suite XI 4.0' -Recurse -ContinueOnError $true
            Remove-RegistryKey -Key 'HKLM\Software\Wow6432Node\SAP BusinessObjects\Suite XI 4.0' -Recurse -ContinueOnError $true

            If (Test-Path "C:\Users\Public\Desktop" -PathType Container -ErrorAction SilentlyContinue) {
                Remove-Item -Force -Path "C:\Users\Public\Desktop\SAP Crystal Reports 2013 viewer.lnk" -ErrorAction SilentlyContinue
            }
            Get-ChildItem -Path "C:\Users" -Include "*" -Force -ErrorAction SilentlyContinue | ForEach-Object ($_) {
                $path0 = $_.FullName + "\Desktop"
                $path1 = $_.FullName + "\Desktop\SAP Crystal Reports 2013 viewer.lnk"
                If (Test-Path $path0 -PathType Container -ErrorAction SilentlyContinue) {
                    Remove-Item -Force -Path $path1 -ErrorAction SilentlyContinue
                }
            }
            
            Write-Log "Crystal Reports Viewer 2013 SP3 UNINSTALL COMPLETED." 
            # Display a message at the end of the uninstall
            Unblock-AppExecution
            $explorerRunning2 = (Get-Process explorer -ea 0) 
            Write-Log "RUNNING explorer END results: $explorerRunning2"
            If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
                If ($installSuccess = $true) {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "COMPLETE: Crystal Reports Viewer 2013 SP3 uninstall COMPLETE. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to finish the uninstall." -ButtonRightText "OK" -Icon Information -NoWait
                    Start-Sleep -s 10
                } ElseIf ($installSuccess = $false) {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "FAILED: Crystal Reports Viewer 2013 SP3 uninstall FAILED. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait
                    Start-Sleep -s 10
                } Else {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "ERROR: Crystal Reports Viewer 2013 SP3 uninstall encounter an ERROR. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait
                    Start-Sleep -s 10
                }
            } Else {
                If ($installSuccess = $true) {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "COMPLETE: Crystal Reports Viewer 2013 SP3 uninstall COMPLETE. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
                    Start-Sleep -s 10
                } ElseIf ($installSuccess = $false) {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "FAILED: Crystal Reports Viewer 2013 SP3 uninstall FAILED. `r`n`r`nACTION REQUIRED: Click OK and then REBOOT / RESTART your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait
                    Start-Sleep -s 10
                } Else {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "ERROR: Crystal Reports Viewer 2013 SP3 uninstall encounter an ERROR. `r`n`r`nACTION REQUIRED: Click OK and then REBOOT / RESTART your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait
                    Start-Sleep -s 10
                }
            }            
        } Else {
            Write-Log "NO, Crystal Reports Viewer 2013 SP3 is NOT installed."
            Unblock-AppExecution
            Show-InstallationPrompt -Message "COMPLETE: Crystal Reports Viewer 2013 SP3 NOT installed, nothing to uninstall. NO changes made. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
            Start-Sleep -s 10
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