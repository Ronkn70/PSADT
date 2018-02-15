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
	[string]$appVendor = 'Citrix'
	[string]$appName = 'Receiver'
	[string]$appVersion = '4.4.1000'
	[string]$appArch = ''
	[string]$appLang = 'EN'
	[string]$appRevision = '01'
	[string]$appScriptVersion = '1.0.0'
	[string]$appScriptDate = '07/1/2016'
	[string]$appScriptAuthor = 'Rees Bauer'
	##*===============================================
	
	##* Do not modify section below
	#region DoNotModify
	
	## Variables: Exit Code
	[int32]$mainExitCode = 0
	
	## Variables: Script
	[string]$deployAppScriptFriendlyName = 'Deploy Application'
	[version]$deployAppScriptVersion = [version]'3.6.7'
	[string]$deployAppScriptDate = '08/17/2015'
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
        Write-Log "Citrix Receiver 4.4.1000 INSTALL: START SCRIPT"
        $isLaptop = (Test-Battery -PassThru).IsLaptop
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

        # Perform Citrix 4.4 Receiver check task here
        $installCitrix4Count = (Get-InstalledApplication -ProductCode "{2B335385-EAB0-4272-BDF9-D475AE51297D}")
        Write-Log "Citrix Receiver 4.4.1000 check results: $installCitrix4Count"
        If (($installCitrix4Count | Measure-Object).Count -gt 0) {
            Write-Log "YES, Citrix Receiver 4.4.1000 is intalled, aborting script with exit code 0."
            Show-InstallationPrompt -Message "Citrix Receiver 4.4.1000 is previously INSTALLED. NO changes made. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
            Start-Sleep -s 30
            Exit-Script -ExitCode "0"
        } Else {
            Write-Log "NO, Citrix Receiver 4.4.1000 is NEEDED."
        }

	    $explorerRunning1 = (Get-Process explorer -ea 0) 
        Write-Log "RUNNING explorer BEGIN results: $explorerRunning1"
        
        Write-Log "START SHOWING INSTALLATION PROGRESS MESSAGES."
        
        #Check if Citrix is in use

        $CitrixUseCheck = Get-Process -Name wfica32 -ErrorAction Ignore

        If (($CitrixUseCheck | Measure-Object).Count -gt 0) {
            # Show Welcome Message, close Citrix processes, allow up to 7 day deferral, and persist the prompt
            Write-Log "Show welcome message, close Citrix processes, allow up to 7 day deferral, and persist the prompt"
            Show-InstallationWelcome -CloseApps "CDViewer,Receiver,concentr,cpviewer,pnagent,pnamain,wfcrun32,wfica32,AuthManSvr,PrimaryAuthModule,CtxCFRUI,redirector,ssonsvr,webhelper" -AllowDeferCloseApps -AllowDefer -DeferTimes 25 -DeferDays 7 -PersistPrompt -MinimizeWindows $false
            Start-Sleep -s 30
        } Else {
            $StopProcess = @(
                "CDViewer",
                "Receiver",
                "concentr",
                "cpviewer",
                "pnagent",
                "pnamain",
                "wfcrun32",
                "wfica32",
                "AuthManSvr",
                "PrimaryAuthModule",
                "CtxCFRUI",
                "redirector",
                "ssonsvr",
                "webhelper")
            Write-Log "Stopping Processes"
            Stop-Process -Processname $StopProcess -Force -ErrorAction Ignore -Confirm:$false
        }
        Start-Sleep -s 5
        
        # Perform OLD Citrix removal here
        Write-Log "Remove existing Citrix client"
        #Show-InstallationProgress -StatusMessage "INSTALLING Citrix Receiver 4.4.1000 `r`n`r`nThis may take some time. Please wait..." -WindowLocation 'Default' -TopMost $true
        Execute-Process -Path "$dirFiles\ReceiverCleanupUtility.exe" -Parameters '/silent' -ContinueOnError $true

        #Remove per user registry Dazzle settings.
        [scriptblock]$HKCURegistrySettings = {
            Remove-RegistryKey -Key 'HKCU\Software\Citrix\Dazzle\Sites\' -Recurse -SID $UserProfile.SID -ErrorAction Ignore
        }
        
        Invoke-HKCURegistrySettingsForAllUsers -RegistrySettings $HKCURegistrySettings -ErrorAction Ignore


		##*===============================================
		##* INSTALLATION 
		##*===============================================
		[string]$installPhase = 'Installation'
		
		## <Perform Installation tasks here>
		
        Write-Log "START installing Citrix Receiver 4.4.1000"
        # Show Progress Message (with the default message) if any Citrix Receiver processes were running and Show-InstallationWelcome triggered
        #Show-InstallationProgress -StatusMessage "INSTALLING Citrix Receiver 4.4 `r`n`r`nThis may take some time. Please wait..." -WindowLocation 'Default' -TopMost $true
        Start-Sleep -s 5

        #Stop Receiver from starting up and prompting user for server info during install. Requires Powershell 3+ for Register-CimIndicationEvent
        Get-Job | Remove-Job -Force
        Start-Job -Name ReceiverPromptJob -WarningAction Ignore -ScriptBlock {
        Unregister-Event ReceiverPromptEvent -ErrorAction Ignore
        $ReceiverPrompt = "Select * from win32_ProcessStartTrace where processname = 'Receiver.exe'"
        Register-CimIndicationEvent -Query $ReceiverPrompt -SourceIdentifier ReceiverPromptEvent -Action {Stop-Process $event.SourceEventArgs.newevent.processID -Force}
        Wait-Event
        }
        Start-Sleep -s 5

        # Perform installation tasks here
        Write-Log "INSTALLING Citrix Receiver 4.4.1000"

        Execute-Process -Path "$dirFiles\CitrixReceiver.exe" -Parameters '/silent /INCLUDESSON /ENABLETRACING=False /EnableCEIP=False /SELFSERVICEMODE=True /ALLOWADDSTORE=A /STARTMENUDIR=\CitrixAPPS /DESKTOPDIR=\CitrixAPPS /STORE0="PNAgent;https://citrix.sncorp.intranet.com/Citrix/PNAgent/config.xml;on;SNC Apps"' -ContinueOnError $true
        
        Write-Log "Citrix Receiver 4.4.1000 install COMPLETE." 
        # Display a message at the end of the install
        Show-InstallationProgress -StatusMessage "Citrix Receiver 4.4.1000 install COMPLETE. `r`n`r`nPlease wait..." -WindowLocation 'Default' -TopMost $true
        Start-Sleep -s 5

		##*===============================================
		##* POST-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Installation'
		
		## <Perform Post-Installation tasks here>
		
        Write-Log "Citrix Receiver 4.4.1000 INSTALL COMPLETED."
        $installedCitrix4Count = (Get-InstalledApplication -ProductCode "{2B335385-EAB0-4272-BDF9-D475AE51297D}")
        
        Unregister-Event ReceiverPrompt -ErrorAction Ignore
        Get-Job | Remove-Job -Force -ErrorAction Ignore
        Unblock-AppExecution

        If (($installedCitrix4Count | Measure-Object).Count -gt 0) {
            # Display a message at the end of the install
            $explorerRunning2 = (Get-Process explorer -ea 0) 
            Write-Log "RUNNING explorer END results: $explorerRunning2"
            If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "COMPLETE: Citrix Receiver 4.4.1000 install COMPLETE. `r`n`r`nClick OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to finish the install." -ButtonRightText "OK" -Icon Information -NoWait
                Start-Sleep -s 30
                Exit-Script -ExitCode "3010"
            } Else {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "COMPLETE: Citrix Receiver 4.4.1000 install COMPLETE. `r`n`r`nPlease reboot your machine at your earliest convenience to complete the installation." -ButtonRightText "OK" -Icon Information -NoWait
                Start-Sleep -s 30
                Exit-Script -ExitCode "3010" 
            }
        } Else {
            # Display a message at the end of the install
            $explorerRunning2 = (Get-Process explorer -ea 0) 
            Write-Log "RUNNING explorer END results: $explorerRunning2"
            If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "ERROR: Citrix Receiver 4.4.1000 install encounter an ERROR. `r`n`r`nClick OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait
                Start-Sleep -s 30
                Exit-Script -ExitCode "1"
            } Else {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "ERROR: Citrix Receiver 4.4.1000 install encounter an ERROR. `r`n`r`nPlease reboot your machine at your earliest convenience to complete the installation." -ButtonRightText "OK" -Icon Error -NoWait
                Start-Sleep -s 30
                Exit-Script -ExitCode "1"
            }
        }

        # End install script
        Write-Log "Citrix Receiver 4.4.1000 INSTALL: END SCRIPT"

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
        Write-Log "Citrix Receiver 4.4.1000 UNINSTALL: START SCRIPT"
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

        # Perform Citrix Receiver 4.4 check task here
        $installedCitrix4Count = (Get-InstalledApplication -ProductCode "{2B335385-EAB0-4272-BDF9-D475AE51297D}")
        Write-Log "Citrix Receiver check results: $installedCitrix4Count"
        If (($installedCitrix4Count | Measure-Object).Count -gt 0) {
            Write-Log "YES, Citrix Receiver 4.4.1000 is INSTALLED."
            Write-Log "UNINSTALLING Citrix Receiver 4.4.1000"     
            
            $explorerRunning1 = (Get-Process explorer -ea 0)
            Write-Log "RUNNING explorer BEGIN results: $explorerRunning1"
            
            Write-Log "START SHOWING UNINSTALLATION PROGRESS MESSAGES."
            Start-Sleep -s 5

            # Show Welcome Message, close processes, allow up to 7 day deferral, and persist the prompt
            Show-InstallationWelcome -CloseApps "CDViewer,concentr,cpviewer,pnagent,pnamain,wfcrun32,wfica32,AuthManSvr,PrimaryAuthModule,CtxCFRUI,redirector,ssonsvr,webhelper" -AllowDeferCloseApps -AllowDefer -DeferTimes 25 -DeferDays 7 -PersistPrompt -MinimizeWindows $false
            Start-Sleep -s 5

            # Show Progress Message (with the default message) if any Citrix processes were running and Show-InstallationWelcome triggered
            Show-InstallationProgress -StatusMessage "UNINSTALLING Citrix Receiver 4.4.1000 `r`n`r`nThis may take some time. Please wait..." -WindowLocation 'Default' -TopMost $true
            Start-Sleep -s 15
            
            # Perform uninstallation tasks here, uninstalling Citrix Receiver 4.4.1000
            Execute-Process -Path "$dirFiles\ReceiverCleanupUtility.exe" -Parameters '/silent' -ContinueOnError $true
            
            Write-Log "Citrix Receiver 4.4.1000 UNINSTALL COMPLETED." 
            # Display a message at the end of the uninstall
            Unblock-AppExecution
            $explorerRunning2 = (Get-Process explorer -ea 0) 
            Write-Log "RUNNING explorer END results: $explorerRunning2"

            # Perform uninstall Citrix Receiver check #2
            $installedCitrix4Count2 = (Get-InstalledApplication -ProductCode "{2B335385-EAB0-4272-BDF9-D475AE51297D}")
            Write-Log "Citrix Receiver check results #2: $installedCitirx4Count2"
            If (($installedCitrix4Count2 | Measure-Object).Count -gt 0) {
                If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "ERROR: Citrix Receiver 4.4.1000 uninstall encounter an ERROR. `r`n`r`nClick OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait
                    Start-Sleep -s 30
                    Exit-Script -ExitCode "1"
                } Else {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "ERROR: Citrix Receiver 4.4.1000 uninstall encounter an ERROR. `r`n`r`nPlease reboot your machine at your earliest convenience to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait
                    Start-Sleep -s 30
                    Exit-Script -ExitCode "1"
                }
            } Else {
                If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "COMPLETE: Citrix Receiver 4.4.1000 uninstall COMPLETE. `r`n`r`nClick OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to finish the uninstall." -ButtonRightText "OK" -Icon Information -NoWait
                    Start-Sleep -s 30
                    Exit-Script -ExitCode "3010"
                } Else {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "COMPLETE: Citrix Receiver 4.4.1000 uninstall COMPLETE. `r`n`r`nPlease reboot your machine at your earliest convenience to complete the uninstall." -ButtonRightText "OK" -Icon Information -NoWait
                    Start-Sleep -s 30
                    Exit-Script -ExitCode "3010" 
                }
            }
        } Else {
            Write-Log "NO, Citrix Receiver 4.4.1000 is NOT installed."
            Show-InstallationPrompt -Message "COMPLETE: Citrix Receiver 4.4.1000 NOT installed, nothing to uninstall. NO changes made. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
            Start-Sleep -s 30
            Exit-Script -ExitCode "0"
        }

        # End uninstall script
        Write-Log "Citrix Receiver 4.4.1000 UNINSTALL: END SCRIPT"
		
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