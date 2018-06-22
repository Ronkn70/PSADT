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
	[string]$appVendor = 'Microsoft'
	[string]$appName = 'Office_Update_CTR'
	[string]$appVersion = ''
	[string]$appArch = ''
	[string]$appLang = 'EN'
	[string]$appRevision = '01'
	[string]$appScriptVersion = '1.0.0'
	[string]$appScriptDate = '04/19/2018'
	[string]$appScriptAuthor = '<Ron Knight>'
	##*===============================================
	## Variables: Install Titles (Only set here to override defaults set by the toolkit)
	[string]$installName = 'Microsoft Office Click to Run Update to Current Version'
	[string]$installTitle = ''
	
	##* Do not modify section below
	#region DoNotModify
	
	## Variables: Exit Code
	[int32]$mainExitCode = 0
	
	## Variables: Script
	[string]$deployAppScriptFriendlyName = 'Deploy Application'
	[version]$deployAppScriptVersion = [version]'3.6.9'
	[string]$deployAppScriptDate = '02/12/2017'
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
		
		## Start Transcript and Verbose (Add -verbose to any cmdlet to log to "C:\Windows\logs\software\WSAScript_$appVendor_$AppName_$LogDate.log")
		$LogDate = get-date -format "MM-d-yy-HH"
		$WSAScriptName = "$appVendor $AppName $AppVersion $appArch"
		Function WSAScript
		{
			function global:Write-Verbose ([string]$Message)
			
			# check $VerbosePreference variable, and turns -Verbose on
			{
				if ($VerbosePreference -ne 'SilentlyContinue')
				{ Write-Host " $Message" -ForegroundColor 'Yellow' }
			}
			
			$VerbosePreference = "Continue"
			
			Start-Transcript -Path "C:\Windows\logs\software\WSAScript_$WSAScriptName_$LogDate.log"
			
			Write-Log "$appVendor $AppName $AppVersion $appArch INSTALL: START SCRIPT"
			$isLaptop = (Test-Battery -PassThru).IsLaptop
			Write-Log "Is Laptop results: $isLaptop"
			Write-Log "Computer name: $envComputerName"
			Write-Log "Script running as: $envUserName"
			$systemAccount = "$envComputerName$"
			Write-Log "System account name: $systemAccount"
			
			# Check for logged on users
			$getUser = (Get-LoggedOnUser)
			Write-Log "Logged on user results: $getUser"
			If ($getUser.ConnectState -eq 'Active')
			{
				Write-Log "YES, USER(s) is logged on to the computer, ACTIVE."
			}
			ElseIf ($getUser.ConnectState -eq 'Disconnected')
			{
				Write-Log "YES, USER(s) is logged on to the computer, DISCONNECTED."
			}
			Else
			{
				Write-Log "NO, USER is not logged on to the computer."
			}
			
			# If powerpoint is running in presentation mode abort
			$presentationPowerPoint = (Test-PowerPoint)
			Write-Log "PowerPoint presentation mode results: $presentationPowerPoint"
			If ($presentationPowerPoint -eq $true)
			{
				Write-Log "YES, detected PowerPoint in presentation mode, aborting script with exit code 69000."
				Exit-Script -ExitCode "69000"
			}
			Else
			{
				Write-Log "NO, PowerPoint is NOT in presentation mode."
			}
			
		# Shows Installation Welcome and prompts to close ALL MS Office apps.
			Show-InstallationWelcome -CloseApps "excel,groove,onenote,infopath,outlook,mspub,powerpnt,winword,winproj,visio,msaccess,skype for business,lync" -AllowDeferCloseApps -AllowDefer -DeferTimes 3 -DeferDays 3 -PersistPrompt -BlockExecution -MinimizeWindows $false
			Start-Sleep -s 15
			
		# Show Progress Message (with the default message)
			Show-InstallationProgress -StatusMessage "UPDATING your Microsoft Office Click to Run to the current version. `r`n`r`nThis may take some time. Please wait..." -WindowLocation 'Default' -TopMost $true
			Start-Sleep -s 15
			
			## <Perform Pre-Installation tasks here>
				
		##*===============================================
		##* INSTALLATION 
		##*===============================================
		[string]$installPhase = 'Installation'
			
		## <Perform Installation tasks here>
		Write-Log "BEGINNING UPGRADE PROCESS FOR: $appVendor $appName $appVersion $appArch."
			
			<# 
			# Force Channel Update
			
			Write-Log "STARTING Change Update Channel To Targeted."
			Show-InstallationProgress -StatusMessage "Verifying MS Office is on the correct update path..." -WindowLocation 'Default' -TopMost $true
			Execute-Process -FilePath "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe" -Arguments "/changesetting Channel=Targeted displaylevel=False" -WindowStyle Hidden -IgnoreExitCodes "3010" -PassThru
			Write-Log "COMPLETED Changing Update Channel To Targeted."
			#>
			
			# Force Office Update 
			Write-Log "STARTNG UPDATE PROCESS FOR: $appVendor $appName $appVersion $appArch."
			Show-InstallationProgress -StatusMessage "Updating your Microsoft Office Click to Run to the latest version" -WindowLocation 'Default' -TopMost $true
			Execute-Process -FilePath "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe" -Arguments "/update user displaylevel=False" -WindowStyle Hidden -IgnoreExitCodes "3010" -PassThru
			Start-Sleep -s 90
			Write-Log "FINISHED UPDATE PROCESS FOR: $appVendor $appName $appVersion $appArch."
		
		##*===============================================
		##* POST-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Installation'
		
		## <Perform Post-Installation tasks here>
			
			write-Log "Applying Skype/Lync Authentication Reg Fix"
			## Define E-mail Address 
			$adaccount = [adsisearcher]"(samaccountname=$env:USERNAME)"
			$email = $adaccount.FindOne().Properties.mail
			
			# SkypeAuthFix
			Set-ItemProperty -path "HKCU:\Software\Microsoft\Office\16.0\Lync\$email\" -name 'EnableRestoreOAuthUsedKeyWhenUsingCachedWebTicket' -Value 1 –Force
			Set-ItemProperty -path "HKCU:\Software\Microsoft\Office\16.0\Lync\$email\" -name 'OAuthUsed' -Value 1 –Force
			
		## End WSAScript and Stop Transcript
		Stop-Transcript
		} WSAScript
		
		## Display a message at the end of the install
		If (-not $useDefaultMsi) { Show-InstallationPrompt -Message "$installName has completed successfully. `r`n`r`n Please RESTART your computer at your earliest convenience." -ButtonRightText 'OK' -Icon Information -NoWait }
	}
	ElseIf ($deploymentType -ieq 'Uninstall')
	{
		##*===============================================
		##* PRE-UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Pre-Uninstallation'
		
		## Show Welcome Message, close Internet Explorer with a 60 second countdown before automatically closing
		Show-InstallationWelcome -CloseApps 'iexplore' -CloseAppsCountdown 60
		
		## Show Progress Message (with the default message)
		Show-InstallationProgress
		
		## <Perform Pre-Uninstallation tasks here>
		
		
		##*===============================================
		##* UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Uninstallation'
		
		## Handle Zero-Config MSI Uninstallations
		If ($useDefaultMsi) {
			[hashtable]$ExecuteDefaultMSISplat =  @{ Action = 'Uninstall'; Path = $defaultMsiFile }; If ($defaultMstFile) { $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile) }
			Execute-MSI @ExecuteDefaultMSISplat
		}
		
		# <Perform Uninstallation tasks here>
		Write-Log "UNSTALLING $appVendor $appName $appVersion $appArch."
		
		
		Write-Log "FINISHED UNSTALLING $appVendor $appName $appVersion $appArch."
		
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