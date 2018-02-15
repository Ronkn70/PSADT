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
	[string]$appVendor = 'Adobe'
	[string]$appName = 'Acrobat'
	[string]$appVersion = 'XI Standard'
	[string]$appArch = '32bit'
	[string]$appLang = 'EN'
	[string]$appRevision = '01'
	[string]$appScriptVersion = '1.0.0'
	[string]$appScriptDate = '3/10/2015'
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
        Write-Log "Adobe Acrobat XI Standard 32bit INSTALL: START SCRIPT"

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
            Write-Log "YES, detected PowerPoint in presentation mode, aborting script with exit code 69000."
            Start-Sleep -s 30
            Exit-Script -ExitCode "69000" 
        } Else {
            Write-Log "NO, PowerPoint is NOT in presentation mode."
        }

        # Perform Adobe Acrobat XI Standard 32bit check task here
        $installNewAcroStdCount = (Get-InstalledApplication -Name "Adobe Acrobat XI Standard")
        Write-Log "Adobe Acrobat XI Standard 32bit check results: $installNewAcroStdCount"
        If (($installNewAcroStdCount | Measure-Object).Count -gt 0) {
            Write-Log "YES, Adobe Acrobat XI Standard 32bit is INSTALLED, aborting script with exit code 0."
            Show-InstallationPrompt -Message "Adobe Acrobat XI Standard 32bit is previously INSTALLED. NO changes made. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
            Start-Sleep -s 30
            Exit-Script -ExitCode "0" 
        } Else {
            Write-Log "NO, Adobe Acrobat XI Standard 32bit is NEEDED."
        }		

        Write-Log "START SHOWING INSTALLATION PROGRESS MESSAGES."
        Start-Sleep -s 60
        
        # Show Welcome Message, close processes, allow up to 7 day deferral, and persist the prompt
        Stop-Process -name acrotray -Force -ErrorAction SilentlyContinue
        Show-InstallationWelcome -CloseApps "excel,groove,onenote,infopath,outlook,mspub,powerpnt,winword,winproj,visio,msaccess,nlnotes,notes,notes2,notes3,acrobat,acrord32" -AllowDeferCloseApps -AllowDefer -DeferTimes 25 -DeferDays 7 -PersistPrompt -BlockExecution -MinimizeWindows $false
        Start-Sleep -s 30
        
		$installOldAcroProCount = (Get-InstalledApplication -Name "Adobe Acrobat XI Pro")
        Write-Log "Adobe Acrobat XI Pro check results: $installOldAcroProCount"
        If (($installOldAcroProCount | Measure-Object).Count -gt 0) {
            Write-Log "Uninstalling Adobe Acrobat XI Pro."
            Show-InstallationProgress -StatusMessage "UNINSTALLING Adobe Acrobat XI Pro. `r`n`r`nThis may take some time. Please wait..."        
            Start-Sleep -s 30
            Remove-MSIApplications "Adobe Acrobat XI Pro" -ErrorAction SilentlyContinue
        }

        $installOldReaderCount = (Get-InstalledApplication -Name "Adobe Reader XI (11.")
        Write-Log "Adobe Reader XI check results: $installOldReaderCount"
        If (($installOldReaderCount | Measure-Object).Count -gt 0) {
            Write-Log "Uninstalling Adobe Reader XI."
            Show-InstallationProgress -StatusMessage "UNINSTALLING Adobe Reader XI. `r`n`r`nThis may take some time. Please wait..."        
            Start-Sleep -s 30
            Remove-MSIApplications "Adobe Reader XI (11." -ErrorAction SilentlyContinue
        }

        ##*===============================================
		##* INSTALLATION 
		##*===============================================
		[string]$installPhase = 'Installation'
		
		## <Perform Installation tasks here>

        Write-Log "START installing Adobe Acrobat XI Standard 32bit."
        # Show Progress Message (with the default message) and Show-InstallationWelcome triggered
        Show-InstallationProgress -StatusMessage "INSTALLING Adobe Acrobat XI Standard 32bit. `r`n`r`nThis may take some time. Please wait..."        
        Start-Sleep -s 30
        
        # Perform installation tasks here
        Write-Log "INSTALLING Adobe Acrobat XI Standard 32bit."
        
        Execute-Process -FilePath "$dirFiles\Setup.exe" -Windowstyle Hidden -IgnoreExitCodes "3010"
        Stop-Process -name acrotray -Force -ErrorAction SilentlyContinue
        Execute-MSI -Action 'Patch' -Path "$dirFiles\AcrobatUpd11014.msp"
        Stop-Process -name acrotray -Force -ErrorAction SilentlyContinue
        Execute-MSI -Action 'Patch' -Path "$dirFiles\AcrobatSecUpd11015.msp"
        Stop-Process -name acrotray -Force -ErrorAction SilentlyContinue
        Set-RegistryKey -Key 'HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Adobe\Adobe Acrobat\11.0\FeatureLockDown' -Name 'iProtectedView' -Value 1 -Type DWord -ContinueOnError $true
        
        ##*===============================================
		##* POST-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Installation'

		## <Perform Post-Installation tasks here>

        Write-Log "Adobe Acrobat XI Standard 32bit INSTALL COMPLETED."
        Unblock-AppExecution 
        
        # Display a message at the end of the install
        $installNewAcroStdCount2 = (Get-InstalledApplication -Name "Adobe Acrobat XI Standard")
        If (($installNewAcroStdCount2 | Measure-Object).Count -gt 0) {
            # Display a message at the end of the install
            Unblock-AppExecution
            Set-RegistryKey -Key 'HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Adobe\Adobe Acrobat\11.0\FeatureLockDown' -Name 'iProtectedView' -Value 1 -Type DWord -ContinueOnError $true
            $explorerRunning2 = (Get-Process explorer -ea 0) 
            Write-Log "RUNNING explorer END results: $explorerRunning2"
            If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "COMPLETE: Adobe Acrobat XI Standard 32bit install COMPLETE. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to finish the install." -ButtonRightText "OK" -Icon Information -NoWait
                Start-Sleep -s 30
            } Else {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "COMPLETE: Adobe Acrobat XI Standard 32bit install COMPLETE. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
                Start-Sleep -s 30
            }
        } Else {
            # Display a message at the end of the install
            Unblock-AppExecution
            $explorerRunning2 = (Get-Process explorer -ea 0) 
            Write-Log "RUNNING explorer END results: $explorerRunning2"
            If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "ERROR: Adobe Acrobat XI Standard 32bit install encounter an ERROR. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait
                Start-Sleep -s 30
                Exit-Script -ExitCode "1"
            } Else {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "ERROR: Adobe Acrobat XI Standard 32bit install encounter an ERROR. `r`n`r`nACTION REQUIRED: Click OK and then REBOOT / RESTART your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait
                Start-Sleep -s 30
                Exit-Script -ExitCode "1"
            }
        }

        # End install script
        Write-Log "Adobe Acrobat XI Standard INSTALL: END SCRIPT"

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
		
		# <Perform Uninstallation tasks here>

        # Start uninstall script
        Write-Log "Adobe Acrobat XI Standard UNINSTALL: START SCRIPT"
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

        $explorerRunning1 = (Get-Process explorer -ea 0) 
        Write-Log "RUNNING explorer BEGIN results: $explorerRunning1"

        # Perform Adobe Acrobat XI Standard check task here
        $installNewAcroStdCount = (Get-InstalledApplication -Name "Adobe Acrobat XI Standard")
        Write-Log "Adobe Acrobat XI Standard check results: $installNewAcroStdCount"
        If (($installNewAcroStdCount | Measure-Object).Count -gt 0) {
            Write-Log "YES, Adobe Acrobat XI Standard is INSTALLED."
            Write-Log "UNINSTALLING Adobe Acrobat XI Standard."

            Write-Log "START SHOWING UNINSTALLATION PROGRESS MESSAGES."
            Start-Sleep -s 60
            
            # Show Welcome Message, close processes, allow up to 7 day deferral, and persist the prompt
            Stop-Process -name acrotray -Force -ErrorAction SilentlyContinue
            Show-InstallationWelcome -CloseApps "excel,groove,onenote,infopath,outlook,mspub,powerpnt,winword,winproj,visio,msaccess,nlnotes,notes,notes2,notes3,acrobat,acrord32" -AllowDeferCloseApps -AllowDefer -DeferTimes 25 -DeferDays 7 -PersistPrompt -BlockExecution -MinimizeWindows $false

            Start-Sleep -s 30
            
            # Show Progress Message (with the default message) and Show-InstallationWelcome triggered
            Show-InstallationProgress -StatusMessage "UNINSTALLING Adobe Acrobat XI Standard. `r`n`r`nThis may take some time. Please wait..."
            Start-Sleep -s 30
            
            # Perform uninstallation tasks here
            Remove-MSIApplications "Adobe Acrobat XI Standard" -ErrorAction SilentlyContinue

            Write-Log "Adobe Acrobat XI Standard UNINSTALL COMPLETED." 
            # Display a message at the end of the uninstall
            $installNewAcroStdCount2 = (Get-InstalledApplication -Name "Adobe Acrobat XI Standard")
            If (($installNewAcroStdCount2 | Measure-Object).Count -gt 0) {
                # Display a message at the end of the install
                Unblock-AppExecution
                $explorerRunning2 = (Get-Process explorer -ea 0) 
                Write-Log "RUNNING explorer END results: $explorerRunning2"
                If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "ERROR: Adobe Acrobat XI Standard uninstall encounter an ERROR. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait
                    Start-Sleep -s 30
                    Exit-Script -ExitCode "1"
                } Else {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "ERROR: Adobe Acrobat XI Standard uninstall encounter an ERROR. `r`n`r`nACTION REQUIRED: Click OK and then REBOOT / RESTART your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait
                    Start-Sleep -s 30
                    Exit-Script -ExitCode "1"
                }
            } Else {
                # Display a message at the end of the uninstall
                Unblock-AppExecution
                $explorerRunning2 = (Get-Process explorer -ea 0) 
                Write-Log "RUNNING explorer END results: $explorerRunning2"
                If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "COMPLETE: Adobe Acrobat XI Standard uninstall COMPLETE. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to finish the install." -ButtonRightText "OK" -Icon Information -NoWait
                    Start-Sleep -s 30
                } Else {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "COMPLETE: Adobe Acrobat XI Standard uninstall COMPLETE. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
                    Start-Sleep -s 30
                }
            }
        } Else {
            Unblock-AppExecution
            Write-Log "NO, Adobe Acrobat XI Standard is NOT installed."
            Show-InstallationPrompt -Message "COMPLETE: Adobe Acrobat XI Standard NOT installed, nothing to uninstall.  NO changes made. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
            Start-Sleep -s 30
        }
        
        # End uninstall script
        Write-Log "Adobe Acrobat XI Standard UNINSTALL: END SCRIPT"
	    
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