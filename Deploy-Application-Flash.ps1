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
	[string]$appName = 'Flash'
	[string]$appVersion = '23.0.0.162'
	[string]$appArch = 'ALL'
	[string]$appLang = 'EN'
	[string]$appRevision = '01'
	[string]$appScriptVersion = '1.0.0'
	[string]$appScriptDate = '9/13/2016'
	[string]$appScriptAuthor = ''
    ## 7/14/16 rewrite additions below
    [string]$ProductName = $appVendor + " " + $appName + " " + $appVersion
    [string]$FlashActiveX = '{3898EF12-76AA-4116-BFCE-EA063420B9E2}'
    [string]$FlashPlugin = '{CE102F76-8858-4CA8-B500-030C9A735C8A}'
    ##*===============================================
	
       
    ## To update this script with new versions:
    ## - Update installation files and update app info in the VARs above.
    ## - Add previous GUID values from $FlashActiveX and $FlashPlugin to the $FlashGUID array below.
    ## - Put the new MSIs' GUIDs for the $FlashActiveX and $FlashPlugin VARs above.
    ## - Update install file MSI names, below in the script, between major version releases. To Do: change this so it uses regex to pull this number from $appVersion.
    ## - Add new GUIDs to previous deployment's detection method.
    ## - Deploy
    ## Rees Bauer

    $FlashGUID = @(
        "{392DC543-3818-4DCA-B95E-C587D83E1C20}",
        "{80C20E2F-B4EF-44E8-BF4A-6A625A9AF168}"
    )


    ## More Flash names to remove

    $FlashName = @(
        "Adobe Flash Player 9 ActiveX",
        "Adobe Flash Player 10 ActiveX",
        "Adobe Flash Player 11 ActiveX",
        "Adobe Flash Player 12 ActiveX",
        "Adobe Flash Player 13 ActiveX",
        "Adobe Flash Player 14 ActiveX",
        "Adobe Flash Player 15 ActiveX",
        "Adobe Flash Player 16 ActiveX",
        "Adobe Flash Player 17 ActiveX",
        "Adobe Flash Player 18 ActiveX",
        "Adobe Flash Player 19 ActiveX",
        "Adobe Flash Player 20 ActiveX",
        "Adobe Flash Player 21 ActiveX",
        "Adobe Flash Player 22 ActiveX",
        "Adobe Flash Player 9 Plugin",
        "Adobe Flash Player 10 Plugin",
        "Adobe Flash Player 11 Plugin",
        "Adobe Flash Player 12 Plugin",
        "Adobe Flash Player 13 Plugin",
        "Adobe Flash Player 14 Plugin",
        "Adobe Flash Player 15 Plugin",
        "Adobe Flash Player 16 Plugin",
        "Adobe Flash Player 17 Plugin",
        "Adobe Flash Player 18 Plugin",
        "Adobe Flash Player 19 Plugin",
        "Adobe Flash Player 20 Plugin",
        "Adobe Flash Player 21 Plugin",
        "Adobe Flash Player 22 Plugin",
        "Adobe Flash Player 9 NPAPI",
        "Adobe Flash Player 10 NPAPI",
        "Adobe Flash Player 11 NPAPI",
        "Adobe Flash Player 12 NPAPI",
        "Adobe Flash Player 13 NPAPI",
        "Adobe Flash Player 14 NPAPI",
        "Adobe Flash Player 15 NPAPI",
        "Adobe Flash Player 16 NPAPI",
        "Adobe Flash Player 17 NPAPI",
        "Adobe Flash Player 18 NPAPI",
        "Adobe Flash Player 19 NPAPI",
        "Adobe Flash Player 20 NPAPI",
        "Adobe Flash Player 21 NPAPI",
        "Adobe Flash Player 22 NPAPI"
   )


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

        # Perform Adobe Flash $ProductName check
        $installNewActivexCount = (Get-InstalledApplication -ProductCode $FlashActiveX)
        $installNewPluginCount = (Get-InstalledApplication -ProductCode $FlashPlugin)
        Write-Log "Adobe Flash $ProductName ActiveX results: $installNewActivexCount"
        Write-Log "Adobe Flash $ProductName Plugin / NPAPI results: $installNewPluginCount"
        If ((($installNewActivexCount | Measure-Object).Count -gt 0) -and (($installNewPluginCount | Measure-Object).Count -gt 0)) {
            Write-Log "YES, Adobe Flash $ProductName is INSTALLED."
            Show-InstallationPrompt -Message "Adobe Flash $ProductName is previously INSTALLED. NO changes made. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
            Write-Log "Adobe Flash $ProductName is previously INSTALLED. NO changes made."
            Start-Sleep -s 10

            Write-Log "START uninstalling old versions of Adobe Flash."
            


            ForEach ($Name in $FlashName) {
                Remove-MSIApplications $Name -ErrorAction SilentlyContinue
            }


            ForEach ($GUID in $FlashGUID) {
                Execute-MSI -Action Uninstall -Path $GUID
            }


            Exit-Script -ExitCode "0" 
        } Else {
            Write-Log "NO, $ProductName is NEEDED."
        }
        
        $explorerRunning1 = (Get-Process explorer -ea 0) 
        Write-Log "RUNNING explorer BEGIN results: $explorerRunning1"
        
		##*===============================================
		##* INSTALLATION 
		##*===============================================
		[string]$installPhase = 'Installation'
		
		## <Perform Installation tasks here>
		
        Write-Log "START uninstalling old versions of Adobe Flash by Name."


        ForEach ($Name in $FlashName) {
            Remove-MSIApplications $Name -ErrorAction SilentlyContinue
        }

        Write-Log "START uninstalling old versions of Adobe Flash by GUID."

        ForEach ($GUID in $FlashGUID) {
            Execute-MSI -Action Uninstall -Path $GUID
        }


        Write-Log "START installing $ProductName."
        
        Show-InstallationProgress -StatusMessage "INSTALLING $ProductName ActiveX. `r`n`r`nThis may take some time. Please wait..."        
        Start-Sleep -s 10
        
        # Perform installation tasks here
        Write-Log "RUNNING Adobe Flash uninstall CLEAN tool."
        Execute-Process -Path "$dirFiles\uninstall_flash_player.exe" -Parameters "-uninstall" -WindowStyle "Hidden" -ContinueOnError $true
        Start-Sleep -s 10
        
        # Perform installation tasks here
        Start-Sleep -s 10
        Write-Log "INSTALLING $ProductName ActiveX."
        Execute-MSI -Action Install -Path "$dirFiles\install_flash_player_23_active_x.msi" -Parameters "/quiet /passive /norestart /qn"
        
        Show-InstallationProgress -StatusMessage "INSTALLING $ProductName Plugin / NPAPI. `r`n`r`nThis may take some time. Please wait..."        
        Start-Sleep -s 10

        # Perform installation tasks here
        Write-Log "INSTALLING $ProductName Plugin / NPAPI."
        Execute-MSI -Action Install -Path "$dirFiles\install_flash_player_23_plugin.msi" -Parameters "/quiet /passive /norestart /qn"
        		
		##*===============================================
		##* POST-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Installation'
		
		## <Perform Post-Installation tasks here>
		
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
                Show-InstallationPrompt -Message "COMPLETE: $Productname install COMPLETE. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
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

        # Start uninstall script
        Write-Log "$ProductName UNINSTALL: START SCRIPT"
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
        
        # Perform uninstall $ProductName check
        $installNewActivexCount = (Get-InstalledApplication -ProductCode $FlashActiveX)
        $installNewPluginCount = (Get-InstalledApplication -ProductCode $FlashPlugin)
        Write-Log "$ProductName ActiveX results: $installNewActivexCount"
        Write-Log "$ProductName Plugin / NPAPI results: $installNewPluginCount"
        If ((($installNewActivexCount | Measure-Object).Count -gt 0) -or (($installNewPluginCount | Measure-Object).Count -gt 0)) {
            Write-Log "YES, $ProductName is INSTALLED."
            Write-Log "UNINSTALLING $ProductName."     
            
            $explorerRunning1 = (Get-Process explorer -ea 0) 
            Write-Log "RUNNING explorer BEGIN results: $explorerRunning1"
            
            # Show Progress Message (with the default message)
            Show-InstallationProgress -StatusMessage "UNINSTALLING $ProductName. `r`n`r`nThis may take some time. Please wait..." -WindowLocation 'Default' -TopMost $true
            Start-Sleep -s 7
            
            # Perform uninstallation tasks here
            Execute-MSI -Action Uninstall -Path $FlashActiveX
            Execute-MSI -Action Uninstall -Path $FlashPlugin

            Write-Log "$ProductName UNINSTALL COMPLETED." 
            # Display a message at the end of the uninstall
            Unblock-AppExecution
            $explorerRunning2 = (Get-Process explorer -ea 0) 
            Write-Log "RUNNING explorer END results: $explorerRunning2"
            If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
                If ($installSuccess = $true) {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "COMPLETE: $ProductName uninstall COMPLETE. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to finish the uninstall." -ButtonRightText "OK" -Icon Information -NoWait
                    Start-Sleep -s 10
                } ElseIf ($installSuccess = $false) {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "FAILED: $ProductName uninstall FAILED. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait
                    Start-Sleep -s 10
                } Else {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "ERROR: $ProductName uninstall encounter an ERROR. `r`n`r`nACTION REQUIRED: Click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then REBOOT / RESTART your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait
                    Start-Sleep -s 10
                }
            } Else {
                If ($installSuccess = $true) {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "COMPLETE: $ProductName uninstall COMPLETE. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
                    Start-Sleep -s 10
                } ElseIf ($installSuccess = $false) {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "FAILED: $ProductName uninstall FAILED. `r`n`r`nACTION REQUIRED: Click OK and then REBOOT / RESTART your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait
                    Start-Sleep -s 10
                } Else {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "ERROR: $ProductName uninstall encounter an ERROR. `r`n`r`nACTION REQUIRED: Click OK and then REBOOT / RESTART your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait
                    Start-Sleep -s 10
                }
            }            
        } Else {
            Write-Log "NO, $ProductName is NOT installed."
            Unblock-AppExecution
            Show-InstallationPrompt -Message "COMPLETE: $ProductName NOT installed, nothing to uninstall. NO changes made. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
            Start-Sleep -s 10
        }

        # End uninstall script
        Write-Log "$ProductName UNINSTALL: END SCRIPT"
				
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
