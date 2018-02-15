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
	[string]$appVendor = 'SEP'
	[string]$appName = 'Antivirus'
	[string]$appVersion = '12.1.7004.6500'
	[string]$appArch = 'x64'
	[string]$appLang = 'EN'
	[string]$appRevision = '01'
	[string]$appScriptVersion = '1.0.0'
	[string]$appScriptDate = '06/30/2016'
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
        Write-Log "SEP Antivirus 12.1.7004.6500 x64 INSTALL: START SCRIPT"
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

        # Perform OS architecture check
        If ($envOSArchitecture -eq "64-Bit") {
            Write-Log "Architecture: $envOSArchitecture" 
            Write-Log "YES, 64-bit architecture continue with the script."
        } Else {
            Write-Log "Architecture: $envOSArchitecture" 
            Write-Log "NO, 32-bit architecture exit the script with code 69001."
            Show-InstallationPrompt -Message "OS Architecture is 32-bit.  SEP Antivirus 12.1.7004.6500 x64 install CANCELED. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Error -NoWait
            Start-Sleep -s 30
            Exit-Script -ExitCode "69001" 
        }

        # Perform SEP Antivirus 12.1.7004.6500 check
        $installNewSEPCount = (Get-InstalledApplication -ProductCode "{F90EEB64-A4CB-484A-8666-812D9F92B37B}")
        Write-Log "SEP Antivirus check results: $installNewSEPCount"
        $registryVersion = (Get-RegistryKey -Key 'HKLM\SOFTWARE\Symantec\Symantec Endpoint Protection\CurrentVersion' -Value 'ProductVersion' -ContinueOnError $true)
        Write-Log "SEP Antivirus registry check results: $registryVersion"
        If ((($installNewSEPCount | Measure-Object).Count -gt 0) -and ($registryVersion -eq "12.1.7004.6500")) {
            Write-Log "YES, SEP Antivirus 12.1.7004.6500 is INSTALLED."
            Write-Log "SEP Antivirus 12.1.7004.6500 is previously INSTALLED. NO changes made."
            Show-InstallationPrompt -Message "SEP Antivirus 12.1.7004.6500 is previously INSTALLED. NO changes made. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
            Start-Sleep -s 30
            Exit-Script -ExitCode "0" 
        } Else {
            Write-Log "NO, SEP Antivirus 12.1.7004.6500 is NEEDED or PENDING."
        }
        
        $explorerRunning1 = (Get-Process explorer -ea 0) 
        Write-Log "RUNNING explorer BEGIN results: $explorerRunning1"

        Write-Log "START SHOWING INSTALLATION PROGRESS MESSAGES."
        Start-Sleep -s 60
        
        Show-InstallationProgress -StatusMessage "SEP Antivirus 12.1.7004.6500 installation STARTED. `r`n`r`nPlease wait..." -WindowLocation "Default" -TopMost $true
        Start-Sleep -s 30
        
		##*===============================================
		##* INSTALLATION 
		##*===============================================
		[string]$installPhase = 'Installation'
		
		## <Perform Installation tasks here>
		
        Write-Log "START installing & checking SEP Antivirus 12.1.7004.6500 x64."
        
        $strDefPath1="C:\Program Files (x86)\Symantec\Symantec Endpoint Protection\12.1.7004.6500.105"
        If ((Test-Path $strDefPath1 -PathType Any) -and ($registryVersion -ne "12.1.7004.6500")) {
            # Show Progress Message (with the default message) and Show-InstallationWelcome triggered
            Write-Log "SEP Antivirus 12.1.7004.6500 is installed, however a reboot is still required."
            Show-InstallationProgress -StatusMessage "SEP Antivirus 12.1.7004.6500 is INSTALLED, however a REBOOT is still required. `r`n`r`nPlease wait..." -WindowLocation "Default" -TopMost $true
            Start-Sleep -s 30
        } Else {
            # Show Progress Message (with the default message) and Show-InstallationWelcome triggered
            Show-InstallationProgress -StatusMessage "INSTALLING SEP Antivirus 12.1.7004.6500 x64. `r`n`r`nThis may take some time. Please wait..." -WindowLocation "Default" -TopMost $true
            Start-Sleep -s 30
            # Perform installation tasks here
            Write-Log "INSTALLING SEP Antivirus 12.1.7004.6500 x64."
            Execute-Process -FilePath "$dirFiles\setup.exe" -WaitForMsiExec -Windowstyle Hidden -IgnoreExitCodes "3010,1" -ContinueOnError $true
            Start-Sleep -s 30
        }
        
		##*===============================================
		##* POST-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Installation'
		
		## <Perform Post-Installation tasks here>
		
        Write-Log "SEP Antivirus 12.1.7004.6500 INSTALL COMPLETED & CHECK." 
        $installNewSEPCount2 = (Get-InstalledApplication -ProductCode "{F90EEB64-A4CB-484A-8666-812D9F92B37B}")

        If ((Test-Path $strDefPath1 -PathType Any) -and (($installNewSEPCount2 | Measure-Object).Count -gt 0)) {
            # Display a message at the end of the install
            $explorerRunning2 = (Get-Process explorer -ea 0) 
            Write-Log "RUNNING explorer END results: $explorerRunning2"
            If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "SEP Antivirus 12.1.7004.6500 x64 installation complete.  `r`n`r`nIn order to begin using the new software, please click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then reboot your machine in order to finish the install." -ButtonRightText "OK" -Icon Information -NoWait
				Start-Sleep -s 30
                Exit-Script -ExitCode "3010"
            } Else {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "SEP Antivirus 12.1.7004.6500 x64 installation complete.  `r`n`r`nIn order to begin using the new software, please reboot your machine at your earliest convenience to complete the install." -ButtonRightText "OK" -Icon Information -NoWait
				Start-Sleep -s 30
                Exit-Script -ExitCode "3010" 
            }
        } Else {
            # Display a message at the end of the install
            $explorerRunning2 = (Get-Process explorer -ea 0) 
            Write-Log "RUNNING explorer END results: $explorerRunning2"
            If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "SEP Antivirus 12.1.7004.6500 x64 installation encounter an error.  `r`n`r`nPlease click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then reboot your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait
				Start-Sleep -s 30
                Exit-Script -ExitCode "1"
            } Else {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "SEP Antivirus 12.1.7004.6500 x64 installation encounter an error.  `r`n`r`nPlease reboot your machine at your earliest convenience in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait
				Start-Sleep -s 30
                Exit-Script -ExitCode "1"
            }
        }

        # End install script
        Write-Log "SEP Antivirus 12.1.7004.6500 x64 INSTALL: END SCRIPT"
        
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