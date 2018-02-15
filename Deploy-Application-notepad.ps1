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
	[string]$appVendor = ''
	[string]$appName = 'Notepad++'
	[string]$appVersion = '6.9.2'
	[string]$appArch = 'X86'
	[string]$appLang = 'EN'
	[string]$appRevision = '01'
	[string]$appScriptVersion = '1.0.0'
	[string]$appScriptDate = '08/17/2016'
	[string]$appScriptAuthor = 'Ron Knight'
	##*===============================================
	## Variables: Install Titles (Only set here to override defaults set by the toolkit)
	[string]$installName = ''
	[string]$installTitle = ''
    [string]$ApplicationName = 'Notepad++ 6.9.2'
    [string]$InstalledApplicationName = 'Notepad++'
	[string]$InstalledApplicationVersion = '6.9.2'
	
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
        Write-Log "Notepad++ 6.9.2 32bit INSTALL: START SCRIPT"		

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
		
		# Perform $ApplicationName check task here
        $installNewSoftwareCount = (Get-InstalledApplication -Name "$InstalledApplicationName")
        Write-Log "$ApplicationName check results: $installNewSoftwareCount"
        If (($installNewSoftwareCount | Measure-Object).Count -gt 0) {
            Write-Log "YES, $ApplicationName or higher is INSTALLED, aborting script with exit code 0."
			$installNewSoftwareVersion = (Get-InstalledApplication -Name "$InstalledApplicationName").DisplayVersion
			Write-Log "$InstalledApplicationName check version: $installNewSoftwareVersion"
			If ($installNewSoftwareVersion -ge $InstalledApplicationVersion) {
				Show-InstallationPrompt -Message "$ApplicationName or a higher version is previously INSTALLED. NO changes made. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
				Start-Sleep -s 10
				Exit-Script -ExitCode "0"
			} Else {
				Write-Log "NO, $ApplicationName is NEEDED / UPGRADE."
				#Show-InstallationPrompt -Message "$ApplicationName needs to be INSTALLED/UPGRADED. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
				Write-Log "$ApplicationName needs to be INSTALLED/UPGRADED."
				Start-Sleep -s 20
				#Exit-Script -ExitCode "0"
			} 
        } Else {
            Write-Log "NO, $ApplicationName is NEEDED."
        }
        
        $explorerRunning1 = (Get-Process explorer -ea 0) 
        Write-Log "RUNNING explorer BEGIN results: $explorerRunning1"

		# Show Welcome Message, close processes, allow up to 7 day deferral, and persist the prompt
        Show-InstallationWelcome -CloseApps "Notepad++,dummynotepad++" -AllowDeferCloseApps -AllowDefer -DeferTimes 25 -DeferDays 7 -PersistPrompt -BlockExecution -MinimizeWindows $false
        Start-Sleep -s 30	

		# Uninstall Older versions of Notepad++
           If (($installNewSoftwareCount | Measure-Object).Count -gt 0)
           {    
           # Show Progress Message (with the default message) and Show-InstallationWelcome triggered
              Show-InstallationProgress -StatusMessage "UNINSTALLING previous Notepad++. `r`n`r`nThis may take some time. Please wait..."
              Start-Sleep -s 20
           # Perform uninstallation tasks here
               If (Test-Path "C:\Program Files (x86)\notepad++\uninstall.exe" -PathType Any -ErrorAction SilentlyContinue) {
                 Execute-Process -FilePath "C:\Program Files (x86)\Notepad++\uninstall.exe" -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
              }
              If (Test-Path "C:\Program Files\Notepad++\uninstall.exe" -PathType Any -ErrorAction SilentlyContinue) {
                  Execute-Process -FilePath "C:\Program Files\notepad++\uninstall.exe" -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
              } 
	    	}  

		##*===============================================
		##* INSTALLATION 
		##*===============================================
		[string]$installPhase = 'Installation'
		
        Write-Log "START installing Notepad++ 6.9.2 32bit."
        
		# Show Progress Message (with the default message) and Show-InstallationWelcome triggered
        Show-InstallationProgress -StatusMessage "INSTALLING Notepad++ 6.9.2 32bit. `r`n`r`nThis may take some time. Please wait..."        
		Start-Sleep -s 30
		
		# Perform installation tasks here
        Write-Log "INSTALLING Notepad++ 6.9.2 32bit."
        Execute-Process -FilePath "$dirFiles\npp.6.9.2.Installer.exe" -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010"
		
		# Copy configuration files
        Write-Log "COPYING CUSTOM configuration files for Notepad++."
        If (Test-Path "C:\Program Files (x86)\Notepad++" -PathType Any) {
            Copy-Item "$dirFiles\config.model.xml" "C:\Program Files (x86)\Notepad++" -Force -ErrorAction SilentlyContinue
		}
		If (Test-Path "C:\Program Files\Notepad++" -PathType Any) {
            Copy-Item "$dirFiles\config.model.xml" "C:\Program Files\Notepad++" -Force -ErrorAction SilentlyContinue
		}
		
		##*===============================================
		##* POST-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Installation'

		## <Perform Post-Installation tasks here>
        
        Write-Log "$ApplicationName INSTALL COMPLETED." 
        Unblock-AppExecution 
        
        # Display a message at the end of the install
 
        $installNewSoftwareCount2 = (Get-InstalledApplication -Name "$InstalledApplicationName")
        Write-Log "$ApplicationName check results: $installNewSoftwareCount2"
        If (($installNewSoftwareCount2 | Measure-Object).Count -gt 0) {
            # Display a message at the end of the install
            Unblock-AppExecution
            $explorerRunning2 = (Get-Process explorer -ea 0) 
            Write-Log "RUNNING explorer END results: $explorerRunning2"
            If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "$ApplicationName installation complete.  `r`n`r`nIn order to begin using the new software, please click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then reboot your machine in order to finish the install." -ButtonRightText "OK" -Icon Information -NoWait
				Start-Sleep -s 0
				Exit-Script -ExitCode "3010"
            } Else {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "$ApplicationName installation complete.  `r`n`r`nIn order to begin using the new software, please reboot your machine at your earliest convenience to complete the install." -ButtonRightText "OK" -Icon Information -NoWait
				Start-Sleep -s 30
				Exit-Script -ExitCode "3010"
            }
        } Else {
            # Display a message at the end of the install
            $explorerRunning2 = (Get-Process explorer -ea 0) 
            Write-Log "RUNNING explorer END results: $explorerRunning2"
            If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "$ApplicationName installation encounter an error.  `r`n`r`nPlease click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then reboot your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait
				Start-Sleep -s 30
                Exit-Script -ExitCode "1"
            } Else {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "$ApplicationName installation encounter an error.  `r`n`r`nPlease reboot your machine at your earliest convenience in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait
				Start-Sleep -s 30
                Exit-Script -ExitCode "1"
            }
        }

        # End install script
        Write-Log "Notepad++ INSTALL: END SCRIPT"

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
        Write-Log " UNINSTALL: START SCRIPT"
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

		# Perform uninstall $ApplicationName check
		$installNewSoftwareVersion = (Get-InstalledApplication -Name "$InstalledApplicationName")
		Write-Log "$ApplicationName check results: $installNewSoftwareVersion"
        If (($installNewSoftwareVersion | Measure-Object).Count -gt 0) {				
				
            Write-Log "YES, $ApplicationName is INSTALLED." 
            
            $explorerRunning1 = (Get-Process explorer -ea 0)
            Write-Log "UNINSTALLING Notepad++."

            Write-Log "START SHOWING UNINSTALLATION PROGRESS MESSAGES."
            Start-Sleep -s 20
            
            # Show Welcome Message, close processes, allow up to 7 day deferral, and persist the prompt
            Show-InstallationWelcome -CloseApps "notepad++,dummynotepad++" -AllowDeferCloseApps -AllowDefer -DeferTimes 25 -DeferDays 7 -PersistPrompt -BlockExecution -MinimizeWindows $false
            Start-Sleep -s 20
            
            # Show Progress Message (with the default message) and Show-InstallationWelcome triggered
            Show-InstallationProgress -StatusMessage "UNINSTALLING Notepad++. `r`n`r`nThis may take some time. Please wait..."
            Start-Sleep -s 20
            
            # Perform uninstallation tasks here
            
            # 64bit OS
            If (Test-Path "C:\Program Files (x86)\notepad++\uninstall.exe" -PathType Any -ErrorAction SilentlyContinue) {
                Write-Log "UNINSTALLING Notepad++ 6.9.2 32bit."
                Execute-Process -Path "C:\Program Files (x86)\Notepad++\Uninstall.exe" -Parameters "/S" -Windowstyle Hidden -IgnoreExitCodes "3010"
            }
            
            # 32bit OS
            If (Test-Path "C:\Program Files\notepad++\uninstall.exe" -PathType Any -ErrorAction SilentlyContinue) {
                Execute-Process -FilePath "C:\Program Files\Notepad++\uninstall.exe" -Parameters "/S" -Windowstyle Hidden -IgnoreExitCodes "3010"
            }
            
				Write-Log "$ApplicationName UNINSTALL COMPLETED." 

            # Display a message at the end of the uninstall
            Unblock-AppExecution
            $explorerRunning2 = (Get-Process explorer -ea 0) 
            Write-Log "RUNNING explorer END results: $explorerRunning2"
			$installNewSoftwareCount2 = (Get-InstalledApplication -Name "$InstalledApplicationName")
            Write-Log "$ApplicationName check results: $installNewSoftwareCount2"
            If (($installNewSoftwareCount2 | Measure-Object).Count -gt 0) {
                If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "$ApplicationName uninstallation encounter an error.  `r`n`r`nPlease click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then reboot your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait
					Start-Sleep -s 30
                    Exit-Script -ExitCode "1"
                } Else {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "$ApplicationName uninstallation encounter an error.  `r`n`r`nPlease reboot your machine at your earliest convenience in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait
					Start-Sleep -s 30
                    Exit-Script -ExitCode "1"
                }
            } Else {
                If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "$ApplicationName uninstallation complete.  `r`n`r`nPlease click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then reboot your machine in order to finish the uninstall." -ButtonRightText "OK" -Icon Information -NoWait
					Start-Sleep -s 30
                    Exit-Script -ExitCode "3010"
                } Else {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "$ApplicationName uninstallation complete.  `r`n`r`nPlease reboot your machine at your earliest convenience to complete the uninstall." -ButtonRightText "OK" -Icon Information -NoWait
					Start-Sleep -s 30
                    Exit-Script -ExitCode "3010"
                }
            }
        } Else {
            Write-Log "NO, $ApplicationName is NOT installed."
            Unblock-AppExecution
            Show-InstallationPrompt -Message "$ApplicationName NOT installed, nothing to uninstall. NO changes made. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
			Start-Sleep -s 30
			Exit-Script -ExitCode "0"
        }
        
        # End uninstall script
        Write-Log "$ApplicationName UNINSTALL: END SCRIPT"
	
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