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
	[string]$appName = 'Tableau Reader 9.0'
	[string]$appVersion = ''
	[string]$appArch = 'x64'
	[string]$appLang = 'EN'
	[string]$appRevision = '01'
	[string]$appScriptVersion = '1.0.0'
	[string]$appScriptDate = '08/17/2015'
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
        Write-Log "Tableau Reader 9.0 64bit INSTALL: START SCRIPT"
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

        # Perform Tableau Reader 9.0 64bit check task here
        $installReaderCount = (Get-InstalledApplication -ProductCode "{F3F06CA7-E47A-4B51-A5FD-B2379875C94D}")
        Write-Log "Tableau Reader 9.0 64bit check results: $installReaderCount"
        If (($installReaderCount | Measure-Object).Count -gt 0) {
            Write-Log "YES, Tableau Reader 9.0 64bit is INSTALLED, aborting script with exit code 0."
            #Show-InstallationPrompt -Message "Tableau Reader 9.0 64bit is previously INSTALLED. NO changes made. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
            Start-Sleep -s 7
            Exit-Script -ExitCode "0" 
        } Else {
            Write-Log "NO, Tableau Reader 9.0 64bit is NEEDED."             
        }		

	    $explorerRunning1 = (Get-Process explorer -ea 0) 
        Write-Log "RUNNING explorer BEGIN results: $explorerRunning1"
        
        # Show Welcome Message, close processes, allow up to 7 day deferral, and persist the prompt
        If ($getUser.ConnectState -eq 'Active') {
            Write-Log "YES, USER(s) is logged on to the computer, ACTIVE."
            Show-InstallationWelcome -CloseApps "tabreader" -AllowDeferCloseApps -AllowDefer -DeferTimes 25 -DeferDays 7 -PersistPrompt -MinimizeWindows $false
        } ElseIf ($getUser.ConnectState -eq 'Disconnected') {
            Write-Log "YES, USER(s) is logged on to the computer, DISCONNECTED."
            Show-InstallationWelcome -CloseApps "tabreader" -Silent -MinimizeWindows $false
        } Else {
            Write-Log "NO, USER is not logged on to the computer."
            Show-InstallationWelcome -CloseApps "tabreader" -Silent -MinimizeWindows $false
        }
                        		
		##*===============================================
		##* INSTALLATION 
		##*===============================================
		[string]$installPhase = 'Installation'
		
		## <Perform Installation tasks here>

        Write-Log "START installing Tableau Reader 9.0 64bit."
        # Show Progress Message (with the default message)
        Show-InstallationProgress -StatusMessage "INSTALLING Tableau Reader 9.0 64bit. `r`n`r`nThis may take some time. Please wait..." -WindowLocation 'Default' -TopMost $true        
        Start-Sleep -s 7
        
        # Perform installation tasks here
        Write-Log "INSTALLING Tableau Reader 9.0 64bit."
        Execute-MSI -Action Install -Path "$dirFiles\tableau-setup-rdr-tableau-9-0.15.0720.1008-x64.msi" -Parameters "DESKTOPSHORTCUT=0 /quiet /passive /norestart /qn"

        Write-Log "REGISTER TABLEAU for ALL USERS."
        [scriptblock]$HKCURegistrySettings = {
            Set-RegistryKey -Key 'HKCU\Software\Tableau\Registration' -Name 'b1924958' -Value '7eda3439' -Type String -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Tableau\Registration' -Name 'Pe69174c0' -Value 'bf242b9e' -Type String -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Tableau\Registration\Data' -Name 'company' -Value 'SNC' -Type String -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Tableau\Registration\Data' -Name 'country' -Value 'US' -Type String -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Tableau\Registration\Data' -Name 'email' -Value 'sncorp@sncorp.com' -Type String -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Tableau\Registration\Data' -Name 'first_name' -Value 'SNCUser' -Type String -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Tableau\Registration\Data' -Name 'last_name' -Value 'SNCUser' -Type String -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Tableau\Registration\Data' -Name 'state' -Value 'NV' -Type String -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Tableau\Registration\Data' -Name 'zip' -Value '89434' -Type String -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Tableau\Registration\License' -Name '182be600' -Value 'f48ee489' -Type String -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Tableau\Tableau Reader 9.0\Install' -Name 'SMSC' -Value 1 -Type DWord -SID $UserProfile.SID           
            Set-RegistryKey -Key 'HKCU\Software\Tableau\Tableau Reader 9.0\LicenseCache' -Name 'Desktop' -Value "ce787794|<license-cache>
  <map key='expiration-date' value='12/31/2028 12:00:00 AM' />
  <map key='feature' value='' />
  <map key='maintenance-date' value='12/31/2028 12:00:00 AM' />
  <map key='signature' value='' />
  <map key='vendor-string' value='CAP=NOLICUI,NOSERVER,REG:SHORTLOC,WARN:0;DC_CAP=;DC_STD=default;EDITION=Standard;MAP_CAP=;MAP_STD=default;OEMNAME=;OFFLINE=true;TRIALVER=' />
</license-cache>
" -Type String -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Tableau\Tableau Reader 9.0\Settings' -Name 'LanguageCode' -Value 'en_US-US' -Type String -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Tableau\Tableau Reader 9.0\Settings' -Name 'RepositoryLanguage' -Value 'en_US' -Type String -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Tableau\Tableau Reader 9.0\Settings' -Name 'SamplesLanguage' -Value 'en_US-US' -Type String -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Tableau\Tableau Reader 9.0\Settings' -Name 'CacheSplashID1' -Value 2001 -Type DWord -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Tableau\Tableau Reader 9.0\Settings' -Name 'CacheSplashID2' -Value 182 -Type DWord -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Tableau\Tableau Reader 9.0\Settings' -Name 'CacheVersion' -Value '' -Type String -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Tableau\Tableau Reader 9.0\Settings' -Name 'NextWorkbookNumber' -Value 2 -Type DWord -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Tableau\Tableau Reader 9.0\Settings' -Name 'Maximized' -Value '1' -Type String -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Tableau\Tableau Reader 9.0\Settings' -Name 'FirstCrashReport' -Value '0' -Type String -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Tableau\Tableau Reader 9.0\Settings' -Name 'LastInstWindowState' -Value 2 -Type DWord -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Tableau\Tableau Reader 9.0\Settings' -Name 'LastInstPlacement' -Value '0,27,1919,1031' -Type String -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Wow6432Node\Tableau\Registration' -Name 'b1924958' -Value '7eda3439' -Type String -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Wow6432Node\Tableau\Registration' -Name 'Pe69174c0' -Value 'bf242b9e' -Type String -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Wow6432Node\Tableau\Registration\Data' -Name 'company' -Value 'SNC' -Type String -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Wow6432Node\Tableau\Registration\Data' -Name 'country' -Value 'US' -Type String -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Wow6432Node\Tableau\Registration\Data' -Name 'email' -Value 'sncorp@sncorp.com' -Type String -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Wow6432Node\Tableau\Registration\Data' -Name 'first_name' -Value 'SNCUser' -Type String -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Wow6432Node\Tableau\Registration\Data' -Name 'last_name' -Value 'SNCUser' -Type String -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Wow6432Node\Tableau\Registration\Data' -Name 'state' -Value 'NV' -Type String -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Wow6432Node\Tableau\Registration\Data' -Name 'zip' -Value '89434' -Type String -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Wow6432Node\Tableau\Registration\License' -Name '182be600' -Value 'f48ee489' -Type String -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Wow6432Node\Tableau\Tableau Reader 9.0\Install' -Name 'SMSC' -Value 1 -Type DWord -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Wow6432Node\Tableau\Tableau Reader 9.0\LicenseCache' -Name 'Desktop' -Value "ce787794|<license-cache>
  <map key='expiration-date' value='12/31/2028 12:00:00 AM' />
  <map key='feature' value='' />
  <map key='maintenance-date' value='12/31/2028 12:00:00 AM' />
  <map key='signature' value='' />
  <map key='vendor-string' value='CAP=NOLICUI,NOSERVER,REG:SHORTLOC,WARN:0;DC_CAP=;DC_STD=default;EDITION=Standard;MAP_CAP=;MAP_STD=default;OEMNAME=;OFFLINE=true;TRIALVER=' />
</license-cache>
" -Type String -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Wow6432Node\Tableau\Tableau Reader 9.0\Settings' -Name 'LanguageCode' -Value 'en_US-US' -Type String -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Wow6432Node\Tableau\Tableau Reader 9.0\Settings' -Name 'RepositoryLanguage' -Value 'en_US' -Type String -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Wow6432Node\Tableau\Tableau Reader 9.0\Settings' -Name 'SamplesLanguage' -Value 'en_US-US' -Type String -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Wow6432Node\Tableau\Tableau Reader 9.0\Settings' -Name 'CacheSplashID1' -Value 2001 -Type DWord -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Wow6432Node\Tableau\Tableau Reader 9.0\Settings' -Name 'CacheSplashID2' -Value 182 -Type DWord -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Wow6432Node\Tableau\Tableau Reader 9.0\Settings' -Name 'CacheVersion' -Value '' -Type String -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Wow6432Node\Tableau\Tableau Reader 9.0\Settings' -Name 'NextWorkbookNumber' -Value 2 -Type DWord -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Wow6432Node\Tableau\Tableau Reader 9.0\Settings' -Name 'Maximized' -Value '1' -Type String -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Wow6432Node\Tableau\Tableau Reader 9.0\Settings' -Name 'FirstCrashReport' -Value '0' -Type String -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Wow6432Node\Tableau\Tableau Reader 9.0\Settings' -Name 'LastInstWindowState' -Value 2 -Type DWord -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Wow6432Node\Tableau\Tableau Reader 9.0\Settings' -Name 'LastInstPlacement' -Value '0,27,1919,1031' -Type String -SID $UserProfile.SID
        }
        Invoke-HKCURegistrySettingsForAllUsers -RegistrySettings $HKCURegistrySettings
        
        Write-Log "Tableau Reader 9.0 64bit install COMPLETE." 
        # Display a message at the end of the install
        Show-InstallationProgress -StatusMessage "Tableau Reader 9.0 64bit install COMPLETE. `r`n`r`nPlease wait..." -WindowLocation 'Default' -TopMost $true
        Start-Sleep -s 7
                        		
		##*===============================================
		##* POST-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Installation'
		
		## <Perform Post-Installation tasks here>

        Write-Log "Tableau Reader 9.0 64bit INSTALL COMPLETED." 
        
        # Display a message at the end of the install
        Unblock-AppExecution
        $explorerRunning2 = (Get-Process explorer -ea 0) 
        Write-Log "RUNNING explorer END results: $explorerRunning2"
        If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
            If ($installSuccess = $true) {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                If ($getUser.ConnectState -eq 'Active') {
                    Write-Log "YES, USER(s) is logged on to the computer, ACTIVE."
                    Show-InstallationPrompt -Message "COMPLETE: Tableau Reader 9.0 64bit install COMPLETE. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to finish the install." -ButtonRightText "OK" -Icon Information -NoWait
                } ElseIf ($getUser.ConnectState -eq 'Disconnected') {
                    Write-Log "YES, USER(s) is logged on to the computer, DISCONNECTED."
                    Show-InstallationPrompt -Message "COMPLETE: Tableau Reader 9.0 64bit install COMPLETE. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to finish the install." -ButtonRightText "OK" -Icon Information -NoWait -Timeout 300
                } Else {
                    Write-Log "NO, USER is not logged on to the computer."
                    Show-InstallationPrompt -Message "COMPLETE: Tableau Reader 9.0 64bit install COMPLETE. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to finish the install." -ButtonRightText "OK" -Icon Information -NoWait -Timeout 300
                }
                Start-Sleep -s 10
            } ElseIf ($installSuccess = $false) {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                If ($getUser.ConnectState -eq 'Active') {
                    Write-Log "YES, USER(s) is logged on to the computer, ACTIVE."
                    Show-InstallationPrompt -Message "FAILED: Tableau Reader 9.0 64bit install FAILED. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait
                } ElseIf ($getUser.ConnectState -eq 'Disconnected') {
                    Write-Log "YES, USER(s) is logged on to the computer, DISCONNECTED."
                    Show-InstallationPrompt -Message "FAILED: Tableau Reader 9.0 64bit install FAILED. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait -Timeout 300
                } Else {
                    Write-Log "NO, USER is not logged on to the computer."
                    Show-InstallationPrompt -Message "FAILED: Tableau Reader 9.0 64bit install FAILED. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait -Timeout 300
                }
                Start-Sleep -s 10
            } Else {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                If ($getUser.ConnectState -eq 'Active') {
                    Write-Log "YES, USER(s) is logged on to the computer, ACTIVE."
                    Show-InstallationPrompt -Message "ERROR: Tableau Reader 9.0 64bit install encounter an ERROR. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait
                } ElseIf ($getUser.ConnectState -eq 'Disconnected') {
                    Write-Log "YES, USER(s) is logged on to the computer, DISCONNECTED."
                    Show-InstallationPrompt -Message "ERROR: Tableau Reader 9.0 64bit install encounter an ERROR. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait -Timeout 300
                } Else {
                    Write-Log "NO, USER is not logged on to the computer."
                    Show-InstallationPrompt -Message "ERROR: Tableau Reader 9.0 64bit install encounter an ERROR. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait -Timeout 300
                }
                Start-Sleep -s 10
            }
        } Else {
            If ($installSuccess = $true) {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                If ($getUser.ConnectState -eq 'Active') {
                    Write-Log "YES, USER(s) is logged on to the computer, ACTIVE."
                    Show-InstallationPrompt -Message "COMPLETE: Tableau Reader 9.0 64bit install COMPLETE. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
                } ElseIf ($getUser.ConnectState -eq 'Disconnected') {
                    Write-Log "YES, USER(s) is logged on to the computer, DISCONNECTED."
                    Show-InstallationPrompt -Message "COMPLETE: Tableau Reader 9.0 64bit install COMPLETE. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait -Timeout 300
                } Else {
                    Write-Log "NO, USER is not logged on to the computer."
                    Show-InstallationPrompt -Message "COMPLETE: Tableau Reader 9.0 64bit install COMPLETE. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait -Timeout 300
                }

                Start-Sleep -s 10
            } ElseIf ($installSuccess = $false) {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                If ($getUser.ConnectState -eq 'Active') {
                    Write-Log "YES, USER(s) is logged on to the computer, ACTIVE."
                    Show-InstallationPrompt -Message "FAILED: Tableau Reader 9.0 64bit install FAILED. `r`n`r`nACTION REQUIRED: Click OK and then REBOOT / RESTART your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait
                } ElseIf ($getUser.ConnectState -eq 'Disconnected') {
                    Write-Log "YES, USER(s) is logged on to the computer, DISCONNECTED."
                    Show-InstallationPrompt -Message "FAILED: Tableau Reader 9.0 64bit install FAILED. `r`n`r`nACTION REQUIRED: Click OK and then REBOOT / RESTART your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait -Timeout 300
                } Else {
                    Write-Log "NO, USER is not logged on to the computer."
                    Show-InstallationPrompt -Message "FAILED: Tableau Reader 9.0 64bit install FAILED. `r`n`r`nACTION REQUIRED: Click OK and then REBOOT / RESTART your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait -Timeout 300
                }
                Start-Sleep -s 10
            } Else {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                If ($getUser.ConnectState -eq 'Active') {
                    Write-Log "YES, USER(s) is logged on to the computer, ACTIVE."
                    Show-InstallationPrompt -Message "ERROR: Tableau Reader 9.0 64bit install encounter an ERROR. `r`n`r`nACTION REQUIRED: Click OK and then REBOOT / RESTART your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait
                } ElseIf ($getUser.ConnectState -eq 'Disconnected') {
                    Write-Log "YES, USER(s) is logged on to the computer, DISCONNECTED."
                    Show-InstallationPrompt -Message "ERROR: Tableau Reader 9.0 64bit install encounter an ERROR. `r`n`r`nACTION REQUIRED: Click OK and then REBOOT / RESTART your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait -Timeout 300
                } Else {
                    Write-Log "NO, USER is not logged on to the computer."
                    Show-InstallationPrompt -Message "ERROR: Tableau Reader 9.0 64bit install encounter an ERROR. `r`n`r`nACTION REQUIRED: Click OK and then REBOOT / RESTART your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait -Timeout 300
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
        Write-Log "Tableau Reader 9.0 64bit UNINSTALL: START SCRIPT"
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
        
        # Perform uninstall Tableau Reader 9.0 64bit check
        $installOldReaderCount = (Get-InstalledApplication -ProductCode "{F3F06CA7-E47A-4B51-A5FD-B2379875C94D}")
        Write-Log "Tableau Reader 9.0 64bit check results: $installOldReaderCount"
        If (($installOldReaderCount | Measure-Object).Count -gt 0) {
            Write-Log "YES, Tableau Reader 9.0 64bit is INSTALLED."
            Write-Log "UNINSTALLING Tableau Reader 9.0 64bit."     
            
            $explorerRunning1 = (Get-Process explorer -ea 0) 
            Write-Log "RUNNING explorer BEGIN results: $explorerRunning1"
            
            # Show Welcome Message, close processes, allow up to 7 day deferral, and persist the prompt
            If ($getUser.ConnectState -eq 'Active') {
                Write-Log "YES, USER(s) is logged on to the computer, ACTIVE."
                Show-InstallationWelcome -CloseApps "tabreader" -AllowDeferCloseApps -AllowDefer -DeferTimes 25 -DeferDays 7 -PersistPrompt -MinimizeWindows $false
            } ElseIf ($getUser.ConnectState -eq 'Disconnected') {
                Write-Log "YES, USER(s) is logged on to the computer, DISCONNECTED."
                Show-InstallationWelcome -CloseApps "tabreader" -Silent -MinimizeWindows $false
            } Else {
                Write-Log "NO, USER is not logged on to the computer."
                Show-InstallationWelcome -CloseApps "tabreader" -Silent -MinimizeWindows $false
            }

            # Show Progress Message (with the default message)
            Show-InstallationProgress -StatusMessage "UNINSTALLING Tableau Reader 9.0 64bit. `r`n`r`nThis may take some time. Please wait..." -WindowLocation 'Default' -TopMost $true
            Start-Sleep -s 7
                
            # Perform uninstallation tasks here
            Execute-MSI -Action Uninstall -Path "{F3F06CA7-E47A-4B51-A5FD-B2379875C94D}"
            
            Write-Log "REMOVE TABLEAU for ALL USERS."
            [scriptblock]$HKCURegistrySettings = {
                Remove-RegistryKey -Key 'HKCU\Software\Tableau' -Recurse -SID $UserProfile.SID -ContinueOnError $true
                Remove-RegistryKey -Key 'HKCU\Software\Wow6432Node\Tableau' -Recurse -SID $UserProfile.SID -ContinueOnError $true
            }
            Invoke-HKCURegistrySettingsForAllUsers -RegistrySettings $HKCURegistrySettings
            
            CMD.EXE /C "C:\Windows\System32\reg.exe delete ""HKEY_LOCAL_MACHINE\SOFTWARE\Tableau"" /f /reg:32"
            CMD.EXE /C "C:\Windows\SysWOW64\reg.exe delete ""HKEY_LOCAL_MACHINE\SOFTWARE\Tableau"" /f /reg:64"
            CMD.EXE /C "C:\Windows\System32\reg.exe delete ""HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Tableau"" /f /reg:32"
            CMD.EXE /C "C:\Windows\SysWOW64\reg.exe delete ""HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Tableau"" /f /reg:64"
            
            # Check to see if the folder exists, if so delete folder
            $strReaderPath1="C:\Program Files\Tableau"
            $strReaderPath2="C:\Program Files (x86)\Tableau"
            If (Test-Path $strReaderPath1 -PathType Any) {
                Remove-Folder -Path $strReaderPath1 -ContinueOnError $true
            }
            If (Test-Path $strReaderPath2 -PathType Any) {
                Remove-Folder -Path $strReaderPath2 -ContinueOnError $true
            }

            Write-Log "Tableau Reader 9.0 64bit UNINSTALL COMPLETED." 
            # Display a message at the end of the install
            Unblock-AppExecution
            $explorerRunning2 = (Get-Process explorer -ea 0) 
            Write-Log "RUNNING explorer END results: $explorerRunning2"
            If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
                If ($installSuccess = $true) {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    If ($getUser.ConnectState -eq 'Active') {
                        Write-Log "YES, USER(s) is logged on to the computer, ACTIVE."
                        Show-InstallationPrompt -Message "COMPLETE: Tableau Reader 9.0 64bit uninstall COMPLETE. ALL DONE. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to finish the uninstall." -ButtonRightText "OK" -Icon Information -NoWait
                    } ElseIf ($getUser.ConnectState -eq 'Disconnected') {
                        Write-Log "YES, USER(s) is logged on to the computer, DISCONNECTED."
                        Show-InstallationPrompt -Message "COMPLETE: Tableau Reader 9.0 64bit uninstall COMPLETE. ALL DONE. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to finish the uninstall." -ButtonRightText "OK" -Icon Information -NoWait -Timeout 300
                    } Else {
                        Write-Log "NO, USER is not logged on to the computer."
                        Show-InstallationPrompt -Message "COMPLETE: Tableau Reader 9.0 64bit uninstall COMPLETE. ALL DONE. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to finish the uninstall." -ButtonRightText "OK" -Icon Information -NoWait -Timeout 300
                    }
                    Start-Sleep -s 10
                } ElseIf ($installSuccess = $false) {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    If ($getUser.ConnectState -eq 'Active') {
                        Write-Log "YES, USER(s) is logged on to the computer, ACTIVE."
                        Show-InstallationPrompt -Message "FAILED: Tableau Reader 9.0 64bit uninstall FAILED. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait
                    } ElseIf ($getUser.ConnectState -eq 'Disconnected') {
                        Write-Log "YES, USER(s) is logged on to the computer, DISCONNECTED."
                        Show-InstallationPrompt -Message "FAILED: Tableau Reader 9.0 64bit uninstall FAILED. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait -Timeout 300
                    } Else {
                        Write-Log "NO, USER is not logged on to the computer."
                        Show-InstallationPrompt -Message "FAILED: Tableau Reader 9.0 64bit uninstall FAILED. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait -Timeout 300
                    }
                    Start-Sleep -s 10
                } Else {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    If ($getUser.ConnectState -eq 'Active') {
                        Write-Log "YES, USER(s) is logged on to the computer, ACTIVE."
                        Show-InstallationPrompt -Message "ERROR: Tableau Reader 9.0 64bit uninstall encounter an ERROR. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait
                    } ElseIf ($getUser.ConnectState -eq 'Disconnected') {
                        Write-Log "YES, USER(s) is logged on to the computer, DISCONNECTED."
                        Show-InstallationPrompt -Message "ERROR: Tableau Reader 9.0 64bit uninstall encounter an ERROR. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait -Timeout 300
                    } Else {
                        Write-Log "NO, USER is not logged on to the computer."
                        Show-InstallationPrompt -Message "ERROR: Tableau Reader 9.0 64bit uninstall encounter an ERROR. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait -Timeout 300
                    }
                    Start-Sleep -s 10
                }
            } Else {
                If ($installSuccess = $true) {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    If ($getUser.ConnectState -eq 'Active') {
                        Write-Log "YES, USER(s) is logged on to the computer, ACTIVE."
                        Show-InstallationPrompt -Message "COMPLETE: Tableau Reader 9.0 64bit uninstall COMPLETE. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
                    } ElseIf ($getUser.ConnectState -eq 'Disconnected') {
                        Write-Log "YES, USER(s) is logged on to the computer, DISCONNECTED."
                        Show-InstallationPrompt -Message "COMPLETE: Tableau Reader 9.0 64bit uninstall COMPLETE. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait -Timeout 300
                    } Else {
                        Write-Log "NO, USER is not logged on to the computer."
                        Show-InstallationPrompt -Message "COMPLETE: Tableau Reader 9.0 64bit uninstall COMPLETE. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait -Timeout 300
                    }
                    Start-Sleep -s 10
                } ElseIf ($installSuccess = $false) {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    If ($getUser.ConnectState -eq 'Active') {
                        Write-Log "YES, USER(s) is logged on to the computer, ACTIVE."
                        Show-InstallationPrompt -Message "FAILED: Tableau Reader 9.0 64bit uninstall FAILED. `r`n`r`nACTION REQUIRED: Click OK and then REBOOT / RESTART your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait
                    } ElseIf ($getUser.ConnectState -eq 'Disconnected') {
                        Write-Log "YES, USER(s) is logged on to the computer, DISCONNECTED."
                        Show-InstallationPrompt -Message "FAILED: Tableau Reader 9.0 64bit uninstall FAILED. `r`n`r`nACTION REQUIRED: Click OK and then REBOOT / RESTART your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait -Timeout 300
                    } Else {
                        Write-Log "NO, USER is not logged on to the computer."
                        Show-InstallationPrompt -Message "FAILED: Tableau Reader 9.0 64bit uninstall FAILED. `r`n`r`nACTION REQUIRED: Click OK and then REBOOT / RESTART your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait -Timeout 300
                    }
                    Start-Sleep -s 10
                } Else {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    If ($getUser.ConnectState -eq 'Active') {
                        Write-Log "YES, USER(s) is logged on to the computer, ACTIVE."
                        Show-InstallationPrompt -Message "ERROR: Tableau Reader 9.0 64bit uninstall encounter an ERROR. `r`n`r`nACTION REQUIRED: Click OK and then REBOOT / RESTART your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait
                    } ElseIf ($getUser.ConnectState -eq 'Disconnected') {
                        Write-Log "YES, USER(s) is logged on to the computer, DISCONNECTED."
                        Show-InstallationPrompt -Message "ERROR: Tableau Reader 9.0 64bit uninstall encounter an ERROR. `r`n`r`nACTION REQUIRED: Click OK and then REBOOT / RESTART your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait -Timeout 300
                    } Else {
                        Write-Log "NO, USER is not logged on to the computer."
                        Show-InstallationPrompt -Message "ERROR: Tableau Reader 9.0 64bit uninstall encounter an ERROR. `r`n`r`nACTION REQUIRED: Click OK and then REBOOT / RESTART your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait -Timeout 300
                    }
                    Start-Sleep -s 10
                }
            }
        } Else {
            Write-Log "NO, Tableau Reader 9.0 64bit is NOT installed."
            If ($getUser.ConnectState -eq 'Active') {
                Write-Log "YES, USER(s) is logged on to the computer, ACTIVE."
                Show-InstallationPrompt -Message "COMPLETE: Tableau Reader 9.0 64bit NOT installed, nothing to uninstall. NO changes made. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
            } ElseIf ($getUser.ConnectState -eq 'Disconnected') {
                Write-Log "YES, USER(s) is logged on to the computer, DISCONNECTED."
                Show-InstallationPrompt -Message "COMPLETE: Tableau Reader 9.0 64bit NOT installed, nothing to uninstall. NO changes made. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait -Timeout 300
            } Else {
                Write-Log "NO, USER is not logged on to the computer."
                Show-InstallationPrompt -Message "COMPLETE: Tableau Reader 9.0 64bit NOT installed, nothing to uninstall. NO changes made. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait -Timeout 300
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