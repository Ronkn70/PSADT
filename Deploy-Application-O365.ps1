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
	[string]$appName = 'Office 365 Pro Plus 2016'
	[string]$appVersion = ''
	[string]$appArch = 'x86'
	[string]$appLang = 'EN'
	[string]$appRevision = '01'
	[string]$appScriptVersion = '1.0.0'
	[string]$appScriptDate = '03/23/2017'
	[string]$appScriptAuthor = '<Ron Knight>'
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
        Write-Log "MS Office 365 Pro Plus 2016 x86 INSTALL: START SCRIPT"
		$isLaptop = (Test-Battery -PassThru).IsLaptop
        Write-Log "Is Laptop results: $isLaptop"
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
            Write-Log "YES, detected PowerPoint in presentation mode, aborting script with exit code 69000."
            Exit-Script -ExitCode "69000" 
        } Else {
            Write-Log "NO, PowerPoint is NOT in presentation mode."
        }
		
    	# Show Welcome Message, close processes, allow up to 5 days of deferrals and persist the prompt
        Show-InstallationWelcome -CloseApps "excel,groove,onenote,infopath,outlook,mspub,powerpnt,winword,winproj,visio,msaccess,Skype for Business,skype,acrobat,acrord32" -BlockExecution -AllowDeferCloseApps -AllowDefer -Deferdays 7 -PersistPrompt -MinimizeWindows $false
        Show-InstallationPrompt -Message 'TriNet IT is about to install MS Office 365 on your computer. Please save ALL your work and click Continue'-PersistPrompt -Icon Warning -ButtonMiddleText 'Continue'
        Show-InstallationProgress -StatusMessage "Checking installed Microsoft Office Products. `r`n`r`nThis may take some time. Please wait..." -WindowLocation 'Default' -TopMost $true
        Start-Sleep -s 10		
		## Show Progress Message (with the default message)
		Show-InstallationProgress
	Write-Log "Uninstalling RMS Packages"
    $RMSGUID = @(
        "{03B04A7D-47EB-48E1-A06D-6AC8F0014DF5}",
        "{3948914B-D208-4593-AC68-57A110E3D2B7}",
        "{6D183BF9-78FA-4024-A5D7-1CF6CEA35C01}"
    )
    	ForEach ($GUID in $RMSGUID) {
                 Execute-MSI -Action Uninstall -Path $GUID
                 }
        Write-Log "RMS Uninstallation Compelete"

        Write-Log "Checking for Microsoft Office Telemetry Agent (x64)"
    		$installMSTelem = (Get-InstalledApplication -Name "Microsoft Office Telemetry Agent (x64)")
        Write-Log "Microsoft Office Telemetry Agent (x64) check results: $installMSTelem"
        If (($installMSTelem | Measure-Object).Count -gt 0) {
            Write-Log "Uninstalling Microsoft Office Telemetry Agent (x64)."
            Remove-MSIApplications "{90150000-0132-0409-1000-0000000FF1CE}" -ErrorAction SilentlyContinue
            Write-Log "Uninstall Microsoft Office Telemetry Agent (x64) COMPLETE."
        }    
    
    Write-Log "Beginning uninstall of all MS Office installs"
    Write-Log "More Information can be found inside the Windows Temp Directory."
  
    <#
This script will remove all versions of office. It runs 2016 and C2R seperately due to bugs in each. 
C2R doesn't remove the extensibility component via the Remove-PreviousInstall command
2016 takes 4 times longer to uninstall via the cmdlet as well
#>


#Measures time it takes to run, this can be removed later.
$ElapsedTime = Measure-Command{

$Log = $env:windir+'\Temp\RemoveOfficeMaster.txt'

. $scriptDirectory\SupportFiles\Remove-AppVConnectionGroupsAll.ps1
Write-Log "App-V Removal Started"
$AppVRemoval = Remove-AppVConnectionGroupsAll
If($AppVRemoval){
if($AppVRemoval.Count -ne 1){
$AppVRemoval[(($AppVRemoval.count) - 1)]
}
else{
$AppVRemoval
}
}
Write-Log "App-V Removal Finished"
. $scriptDirectory\SupportFiles\Remove-PreviousOfficeInstalls.ps1

#This get's a baseline of whats installed, cmdlet is inside the Remove-PreviousOfficeInstalls.ps1
$Products = Get-OfficeVersion -ShowAllInstalledProducts
Write-Log "Office 2007-2013 Removal Started"
#Uses the cmdlet to remove all versions of office 2013 and prior.
Remove-PreviousOfficeInstalls -Confirm -ProductsToRemove AllOfficeProducts
#Remove-PreviousOfficeInstalls -Remove2016Installs $true -Confirm -ProductsToRemove AllOfficeProducts
Write-Log "Office 2007-2013 Removal Finished"

<#
Does a search to see if C2R office is likely installed via the Get-Service cmdlet. 
Can be modified/upgraded to use the $Products variale found earlier.
Not using skipsd command due to it leaving behind infopath 2013 shortcuts on W10
SpotCheck queries for installed versions of office.
#>
$SpotCheck = Get-OfficeVersion -ShowAllInstalledProducts
$ClickToRun = Get-Service -Name ClickToRunSvc -ErrorAction SilentlyContinue
Write-Log "C2R Removal Started"
if($ClickToRun -or $Products.DisplayName -like 'Microsoft Office 365 ProPlus'){
wscript $scriptDirectory'\SupportFiles\OffScrubC2R.vbs' All /Quiet | Out-Null
}
Write-Log "C2R Removal Finished"
#Checks to see what is left and removes MSI versions of 2016 if it still exists.
$SpotCheck = Get-OfficeVersion -ShowAllInstalledProducts
Write-Output $SpotCheck.DisplayName
Write-Log "Office 2016 Removal Started"
if($SpotCheck.Version -match '16.*'){
wscript $scriptDirectory'\SupportFiles\OffScrub_O16msi.vbs' ALL /Quiet | Out-Null
Write-Output $SpotCheck.version
}
Write-Log "Office 2016 Removal Finished"
#Runs final check to see if anything remains.
$SpotCheck = Get-OfficeVersion -ShowAllInstalledProducts
$AppVSpotCheck  = Get-AppvClientConnectionGroup -all
$AppVSpotCheck += Get-AppvClientPackage -all
}
Write-Log "Office Removal Finished, please check the log in the windows temp directory for more information,"

$ElapsedTime , "Initial Items:" , $Products.DisplayName , $AppVRemoval , "Remaining Items:" , $SpotCheck.DisplayName , $AppVSpotCheck | Out-File $Log -Append   
    
	
		##*===============================================
		##* INSTALLATION 
		##*===============================================
		[string]$installPhase = 'Installation'
		
		## Handle Zero-Config MSI Installations
		If ($useDefaultMsi) {
			[hashtable]$ExecuteDefaultMSISplat =  @{ Action = 'Install'; Path = $defaultMsiFile }; If ($defaultMstFile) { $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile) }
			Execute-MSI @ExecuteDefaultMSISplat; If ($defaultMspFiles) { $defaultMspFiles | ForEach-Object { Execute-MSI -Action 'Patch' -Path $_ } }
		}
		
		## <Perform Installation tasks here>
       Write-Log "Office 365 Pro Plus 2016 x86 Install Started."
        Execute-Process -FilePath "$dirFiles\setup.exe" -Arguments "/configure installFRD.xml" -WindowStyle Hidden -IgnoreExitCodes "3010" -PassThru
        Write-Log "Office 365 Pro Plus 2016 x86 Install Finished."		
		##*===============================================
		##* POST-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Installation'
		
		## <Perform Post-Installation tasks here>
        Write-Log "Applying post-installation reg entries."
        [scriptblock]$HKCURegistrySettings = {
            Set-RegistryKey -Key 'HKCU\SOFTWARE\Microsoft\Office\16.0\Outlook\AutoDiscover' -Name 'ZeroConfigExchange' -Value 1 -Type DWord -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\SOFTWARE\Policies\Microsoft\office\16.0\Outlook\AutoDiscover' -Name 'ZeroConfigExchange' -Value 1 -Type DWord -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Microsoft\Office\16.0\Common\General' -Name 'ShownFirstRunOptin' -Value 1 -Type DWord -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\Software\Microsoft\Office\16.0\Common' -Name 'QMEnable' -Value 0 -Type DWord -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\SOFTWARE\Microsoft\Office\16.0\Registration' -Name 'AcceptAllEulas' -Value 1 -Type DWord -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\SOFTWARE\Microsoft\Office\16.0\Lync' -Name 'IsBasicTutorialSeenByUser' -Value 1 -Type DWord -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\SOFTWARE\Microsoft\Office\16.0\Lync' -Name 'AutoSignInWhenUserSessionStarts' -Value 1 -Type DWord -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\SOFTWARE\Microsoft\Office\16.0\Lync' -Name 'FirstRun' -Value 1 -Type DWord -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\SOFTWARE\Microsoft\Office\16.0\Lync' -Name 'EnableEventLogging' -Value 1 -Type DWord -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\SOFTWARE\Microsoft\Office\16.0\Lync' -Name 'MinimizeWindowToNotificationArea' -Value 0 -Type DWord -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\SOFTWARE\Microsoft\Office\16.0\Lync' -Name 'AutoOpenMainWindowWhenStartup' -Value 1 -Type DWord -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\SOFTWARE\Microsoft\Office\16.0\Lync' -Name 'UserConsentedTelemetryUpload' -Value 1 -Type DWord -SID $UserProfile.SID
            Set-RegistryKey -Key 'HKCU\SOFTWARE\Microsoft\Office\16.0\Lync' -Name 'AutoSignInWhenUserSessionStarts' -Value 1 -Type DWord -SID $UserProfile.SID
        }
        Invoke-HKCURegistrySettingsForAllUsers -RegistrySettings $HKCURegistrySettings
        Write-Log "Finished applying post-installation reg entries."
        Write-Log "$AppVendor $AppName $appVersion $AppArch install COMPLETE."

        ## Display a message at the end of the install

		If (-not $useDefaultMsi) { Show-InstallationPrompt -Message 'Thanks for your patience. MS Office 365 Pro Plus 2016 x86 has successfully been installed. Please Reboot your computer at your earliest convenience.' -ButtonRightText 'OK' -Icon Information -NoWait }
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