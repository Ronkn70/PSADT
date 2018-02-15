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

	## Set the script execution policy for this process
	Try { Set-ExecutionPolicy -ExecutionPolicy 'ByPass' -Scope 'Process' -Force -ErrorAction 'Stop' } Catch {}
	
	##*===============================================
	##* VARIABLE DECLARATION
	##*===============================================
	## Variables: Application
	[string]$appVendor = 'Adobe'
	[string]$appName = 'Acrobat Reader'
	[string]$appVersion = '11.0.17'
	[string]$appArch = ''
	[string]$appLang = 'EN'
	[string]$appRevision = '01'
	[string]$appScriptVersion = '1.0.0'
	[string]$appScriptDate = '7/14/2016'
	[string]$appScriptAuthor = ''
    [string]$ProductName = $appVendor + " " + $appName + " " + $appVersion
    [string]$ReaderGUID = '{AC76BA86-7AD7-1033-7B44-AB0000000001}' #Only changes on major versions
    [string]$PatchPath = "$dirFiles\AdbeRdrUpd11017.msp"
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
        Write-Log "$ProductName INSTALL: START SCRIPT"
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

        # $ProductName check
        $getAcrobatRdrProductCode = (Get-InstalledApplication -ProductCode $ReaderGUID)
        
        If ($envOSArchitecture -eq "64-Bit") {
            $getAcrobatRdrVersion = (Get-RegistryKey -Key "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$ReaderGUID" -Value 'DisplayVersion' -ContinueOnError $true)
        } Else {
            $getAcrobatRdrVersion = (Get-RegistryKey -Key "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$ReaderGUID" -Value 'DisplayVersion' -ContinueOnError $true)
        }

        Write-Log "Acrobat Rdr Registry Version Check Result: $getAcrobatStdVersion"

        If (($getAcrobatRdrProductCode | Measure-Object).Count -eq 0) {
            Write-Log "NO, $AppName is not installed. No patch needed. No changes made."
            Show-InstallationPrompt -Message "$AppName is not installed. No Patch needed. No changes made. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
            Exit-Script -Exitcode "0"
        
        } ElseIf ((($getAcrobatRdrProductCode | Measure-Object).Count -gt 0) -and ($getAcrobatRdrVersion -ge $appVersion)) {
            Write-Log "YES, $ProductName or greater is INSTALLED."
            Show-InstallationPrompt -Message "$ProductName or greater is previously INSTALLED. NO changes made. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
            Write-Log "$ProductName or greater is previously INSTALLED. NO changes made."
            Start-Sleep -s 10
            Exit-Script -ExitCode "0"

        } Else {
            Write-Log "$ProductName is NEEDED."
        }
        
        $explorerRunning1 = (Get-Process explorer -ea 0) 
        Write-Log "RUNNING explorer BEGIN results: $explorerRunning1"
        
		##*===============================================
		##* INSTALLATION 
		##*===============================================
		[string]$installPhase = 'Installation'
		
		## <Perform Installation tasks here>
		       
        Write-Log "START installing $ProductName"
        Show-InstallationProgress -StatusMessage "INSTALLING $ProductName `r`n`r`nThis may take some time. Please wait..."        
        Start-Sleep -s 10
        Write-Log "INSTALLING $ProductName"
        Execute-MSI -Action 'Patch' -Path $PatchPath
          	
		##*===============================================
		##* POST-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Installation'
		
		## <Perform Post-Installation tasks here>
		
        Set-RegistryKey -Key 'HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Adobe\Acrobat Reader\11.0\FeatureLockDown' -Name 'iProtectedView' -Value 1 -Type DWord -ContinueOnError $true

        Write-Log "$ProductName INSTALL COMPLETED." 

        # Display a message at the end of the install
        $explorerRunning2 = (Get-Process explorer -ea 0) 
        Write-Log "RUNNING explorer END results: $explorerRunning2"
        If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
            If ($installSuccess = $true) {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "COMPLETE: $ProductName install COMPLETE. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to finish the install." -ButtonRightText "OK" -Icon Information -NoWait
                Start-Sleep -s 10
            } ElseIf ($installSuccess = $false) {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "FAILED: $ProductName install FAILED. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait
                Start-Sleep -s 10
            } Else {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "ERROR: $ProductName install encounter an ERROR. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait
                Start-Sleep -s 10
            }
        } Else {
            If ($installSuccess = $true) {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "COMPLETE: $ProductName install COMPLETE. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
                Start-Sleep -s 10
            } ElseIf ($installSuccess = $false) {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "FAILED: $ProductName install FAILED. `r`n`r`nACTION REQUIRED: Click OK and then REBOOT / RESTART your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait
                Start-Sleep -s 10
            } Else {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "ERROR: $ProductName install encounter an ERROR. `r`n`r`nACTION REQUIRED: Click OK and then REBOOT / RESTART your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait
                Start-Sleep -s 10
            }
        }
        
        # End install script
        Write-Log "$ProductName INSTALL: END SCRIPT"
        		
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
