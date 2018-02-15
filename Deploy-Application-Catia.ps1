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
	[string]$appName = 'CATIA B25 SP3'
	[string]$appVersion = ''
	[string]$appArch = ''
	[string]$appLang = 'EN'
	[string]$appRevision = '01'
	[string]$appScriptVersion = '1.0.0'
	[string]$appScriptDate = '09/14/2016'
	[string]$appScriptAuthor = ''
	##*===============================================
	## Variables: Install Titles (Only set here to override defaults set by the toolkit)
	[string]$installName = ''
	[string]$installTitle = ''
	[string]$ApplicationName = 'CATIA B25 SP3'
    [string]$InstalledApplicationName = 'Dassault Systemes Software Version 5-6 Release 2015 (B25)'
	$CATIAGUID = @(
		"{C857169D-3F1A-4530-99A0-CAE966CE267E}",
		"{7C534131-6431-4ECB-9069-525CB5F75CC8}",
		"{F2F2DEA7-36AB-4E13-907C-D8BDE775EF97}",
		"{CF1EB598-B424-436A-B15F-B763846BA970}"
	)
	
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
			
		# <Perform Pre-Installation tasks here>

        # Start install script
        Write-Log "$ApplicationName INSTALL: START SCRIPT"
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

        # Perform $ApplicationName check task here
        ####Name = "$InstalledApplicationName"
        $installNewSoftwareCount = (Get-InstalledApplication -Name "$InstalledApplicationName")
        Write-Log "$ApplicationName check results: $installNewSoftwareCount"
        If (($installNewSoftwareCount | Measure-Object).Count -gt 0) {
            Write-Log "YES, $ApplicationName is INSTALLED, aborting script with exit code 0."
			Show-InstallationPrompt -Message "$ApplicationName is previously INSTALLED. NO changes made. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
			Start-Sleep -s 10
			Exit-Script -ExitCode "0"
        } Else {
            Write-Log "NO, $ApplicationName is NEEDED."
			#Exit-Script -ExitCode "0"
        }
        
        $explorerRunning1 = (Get-Process explorer -ea 0) 
        Write-Log "RUNNING explorer BEGIN results: $explorerRunning1"
        
        # Show Welcome Message, close processes
		Show-InstallationWelcome -CloseApps "smartworkspacemanager,viewwrap,smarteam,smartbox,catutil,catprintermanager,catiaenv,cnext,acad,acsignapply,adcadmn,addplwiz,adflashvideoplayer,adrefman,adsubaware,dwgcheckstandards,pc3exe,styexe,styshwiz" -AllowDeferCloseApps -AllowDefer -DeferTimes 25 -DeferDays 7 -PersistPrompt -MinimizeWindows $false
		Start-Sleep -s 10
        		
		##*===============================================
		##* INSTALLATION 
		##*===============================================
		[string]$installPhase = 'Installation'
		
		# <Perform Installation tasks here>

        Write-Log "START installing $ApplicationName."
        Show-InstallationProgress -StatusMessage "INSTALLING $ApplicationName. `r`n`r`nThis may take some time. Please wait..." -WindowLocation 'Default' -TopMost $true
        
		# Perform installation tasks here
		Write-Log "INSTALLING Prerequisites For $ApplicationName."
		Execute-MSI -Action "Install" -Path "$dirFiles\DOC_CATIA_PLM_Express.AllOS\1\INTEL\InstallDSSoftwarePrerequisites_x86_x64.msi" -Parameters "REBOOT=ReallySuppress /QN /passive /norestart" -ContinueOnError $true
		Start-Sleep -s 20
		Execute-MSI -Action "Install" -Path "$dirFiles\CATIA-PLM-Express.windows\1\CATIA_PLM_Express.win_b64\1\WIN64\InstallDSSoftwarePrerequisites_x86_x64.msi" -Parameters "REBOOT=ReallySuppress /QN /passive /norestart" -ContinueOnError $true
		Start-Sleep -s 20
		Execute-MSI -Action "Install" -Path "$dirFiles\SPK.win_b64\1\WIN64\InstallDSSoftwareVC9Prerequisites_x86_x64.msi" -Parameters "REBOOT=ReallySuppress /QN /passive /norestart" -ContinueOnError $true
		Start-Sleep -s 20
		Execute-MSI -Action "Install" -Path "$dirFiles\CATIA-PLM-Express.windows\1\CATIA_PLM_Express.win_b64\1\WIN64\InstallDSSoftwareVC10Prerequisites_x86_x64.msi" -Parameters "REBOOT=ReallySuppress /QN /passive /norestart" -ContinueOnError $true
		Start-Sleep -s 20
		Execute-MSI -Action "Install" -Path "$dirFiles\DOC_CATIA_PLM_Express.AllOS\1\INTEL\InstallDSSoftwareVC11Prerequisites_x86_x64.msi" -Parameters "REBOOT=ReallySuppress /QN /passive /norestart" -ContinueOnError $true
		Start-Sleep -s 20
		Execute-MSI -Action "Install" -Path "$dirFiles\CATIA-PLM-Express.windows\1\CATIA_PLM_Express.win_b64\1\WIN64\InstallDSSoftwareVC11Prerequisites_x86_x64.msi" -Parameters "REBOOT=ReallySuppress /QN /passive /norestart" -ContinueOnError $true
		Start-Sleep -s 20
		Execute-MSI -Action "Install" -Path "$dirFiles\CATIA-PLM-Express.windows\1\CATIA_PLM_Express.win_b64\1\VBA\Vba71_x64.msi" -Parameters "REBOOT=ReallySuppress /QN /passive /norestart" -ContinueOnError $true
		Start-Sleep -s 20
		Execute-MSI -Action "Install" -Path "$dirFiles\CATIA-PLM-Express.windows\1\CATIA_PLM_Express.win_b64\1\VBA\Vba71_x64_1033.MSI" -Parameters "REBOOT=ReallySuppress /QN /passive /norestart" -ContinueOnError $true
		Start-Sleep -s 20
		Execute-MSI -Action "Patch" -Path "$dirFiles\CATIA-PLM-Express.windows\1\CATIA_PLM_Express.win_b64\1\VBA\Vba71_x64_1033_KB2803801.msp" -Parameters "REBOOT=ReallySuppress /QN /passive /norestart" -ContinueOnError $true
		Start-Sleep -s 20
		Execute-MSI -Action "Patch" -Path "$dirFiles\CATIA-PLM-Express.windows\1\CATIA_PLM_Express.win_b64\1\VBA\Vba71_x64_KB2803801.msp" -Parameters "REBOOT=ReallySuppress /QN /passive /norestart" -ContinueOnError $true
		Start-Sleep -s 20
		Execute-MSI -Action "Install" -Path "$dirFiles\CATIA-PLM-Express.windows\1\CATIA_PLM_Express.win_b64\1\VBA\Vba71_x86.msi" -Parameters "REBOOT=ReallySuppress /QN /passive /norestart" -ContinueOnError $true
		Start-Sleep -s 20
		Execute-MSI -Action "Install" -Path "$dirFiles\CATIA-PLM-Express.windows\1\CATIA_PLM_Express.win_b64\1\VBA\Vba71_x86_1033.MSI" -Parameters "REBOOT=ReallySuppress /QN /passive /norestart" -ContinueOnError $true
		Start-Sleep -s 20
		Execute-MSI -Action "Patch" -Path "$dirFiles\CATIA-PLM-Express.windows\1\CATIA_PLM_Express.win_b64\1\VBA\Vba71_x86_1033_KB2803801.msp" -Parameters "REBOOT=ReallySuppress /QN /passive /norestart" -ContinueOnError $true
		Start-Sleep -s 20
		Execute-MSI -Action "Patch" -Path "$dirFiles\CATIA-PLM-Express.windows\1\CATIA_PLM_Express.win_b64\1\VBA\Vba71_x86_KB2803801.msp" -Parameters "REBOOT=ReallySuppress /QN /passive /norestart" -ContinueOnError $true
		Start-Sleep -s 20
		
		Write-Log "INSTALLING PLM Express GA For $ApplicationName."
		Start-Sleep -s 20
		Execute-Process -FilePath "$dirFiles\CATIA-PLM-Express.windows\1\CATIA_PLM_Express.win_b64\1\WIN64\StartB.exe" -Arguments "-u ""C:\Program Files\Dassault Systemes\B25"" -newdir -all -orbixport 1570 -orbixbase 1590 -orbixrange 200 -addUserPrivilegesForOrbix -backbonePorts 55555 55556 -VRPort 6668 -all -allextra_prd -v -noLang all -noFonts -noDesktopIcon -noStartMenuIcon -noStartMenuTools -noreboot" -WindowStyle Normal -IgnoreExitCodes "3010" -PassThru -ContinueOnError $true -NoWait
		#CMD.EXE /C """$dirFiles\CATIA-PLM-Express.windows\1\CATIA_PLM_Express.win_b64\1\WIN64\StartB.exe"" -u ""C:\Program Files\Dassault Systemes\B25"" -newdir -all -orbixport 1570 -orbixbase 1590 -orbixrange 200 -addUserPrivilegesForOrbix -backbonePorts 55555 55556 -VRPort 6668 -all -allextra_prd -v -noLang all -noFonts -noDesktopIcon -noStartMenuIcon -noStartMenuTools -noreboot"
		Start-Sleep -s 20
        Wait-Process -name StartB -ErrorAction SilentlyContinue
        Start-Sleep -s 60
		Wait-Process -name StartB -ErrorAction SilentlyContinue
        Start-Sleep -s 60
		
		Write-Log "INSTALLING Service Pack For $ApplicationName."
		Start-Sleep -s 20
		Execute-Process -FilePath "$dirFiles\SPK.win_b64\1\WIN64\StartSPKB.exe" -Arguments "-v -bC -killprocess" -WindowStyle Normal -IgnoreExitCodes "3010" -PassThru -ContinueOnError $true
		#CMD.EXE /C """$dirFiles\SPK.win_b64\1\WIN64\StartSPKB.exe"" -v -bC -killprocess"
		Start-Sleep -s 20
		
		Write-Log "COPYING Extra Files CATIA Desktop Links and DSLS file."
		Copy-Item "$dirFiles\ExtraFiles\Start\CATIA PLM Express" "$envProgramData\Microsoft\Windows\Start Menu\Programs" -Force -Recurse -ErrorAction SilentlyContinue
		Copy-Item "$dirFiles\ExtraFiles\SNC CATIA - R25.lnk" "$envCommonDesktop" -Force -ErrorAction SilentlyContinue
		Copy-Item "$dirFiles\ExtraFiles\Licenses\Licenses" "$envProgramData\DassaultSystemes" -Force -Recurse -ErrorAction SilentlyContinue
		Start-Sleep -s 20
		
		Write-Log "Registering CNEXT."
		If (Test-Path "C:\Program Files\Dassault Systemes\B25\win_b64\code\bin\cnext.exe" -PathType Any) {
            Execute-Process -FilePath "C:\Program Files\Dassault Systemes\B25\win_b64\code\bin\cnext.exe" -Arguments "-regserver" -WindowStyle Hidden -IgnoreExitCodes "3010" -PassThru -ContinueOnError $true
        }
		If (Test-Path "C:\Program Files (x86)\Dassault Systemes\B25\win_b64\code\bin\cnext.exe" -PathType Any) {
            Execute-Process -FilePath "C:\Program Files (x86)\Dassault Systemes\B25\win_b64\code\bin\cnext.exe" -Arguments "-regserver" -WindowStyle Hidden -IgnoreExitCodes "3010" -PassThru -ContinueOnError $true
        }
		Start-Sleep -s 20
				
		##*===============================================
		##* POST-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Installation'
		
		# <Perform Post-Installation tasks here>
        
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
				Start-Sleep -s 30
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
        Write-Log "$ApplicationName INSTALL: END SCRIPT"
		
	}
	ElseIf ($deploymentType -ieq 'Uninstall')
	{
		##*===============================================
		##* PRE-UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Pre-Uninstallation'
		
		# <Perform Pre-Uninstallation tasks here>
		
		##*===============================================
		##* UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Uninstallation'
		
		# <Perform Uninstallation tasks here>

        # Start uninstall script
        Write-Log "$ApplicationName UNINSTALL: START SCRIPT"
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

		# Perform uninstall $ApplicationName check
		$installNewSoftwareCount = (Get-InstalledApplication -Name "$InstalledApplicationName")
        Write-Log "$ApplicationName check results: $installNewSoftwareCount"
        If (($installNewSoftwareCount | Measure-Object).Count -gt 0) {
        
            Write-Log "YES, $ApplicationName is INSTALLED." 
            
            $explorerRunning1 = (Get-Process explorer -ea 0) 
            Write-Log "RUNNING explorer BEGIN results: $explorerRunning1"
            
            # Show Welcome Message, close processes
			Show-InstallationWelcome -CloseApps "smartworkspacemanager,viewwrap,smarteam,smartbox,catutil,catprintermanager,catiaenv,cnext,acad,acsignapply,adcadmn,addplwiz,adflashvideoplayer,adrefman,adsubaware,dwgcheckstandards,pc3exe,styexe,styshwiz" -AllowDeferCloseApps -AllowDefer -DeferTimes 25 -DeferDays 7 -PersistPrompt -MinimizeWindows $false
			Show-InstallationProgress -StatusMessage "UNINSTALLING $ApplicationName. `r`n`r`nThis may take some time. Please wait..." -WindowLocation 'Default' -TopMost $true
            Start-Sleep -s 10
                                   
            # Perform uninstallation tasks here
			Write-Log "UNINSTALLING $ApplicationName."          
            
			# Remove older versions
			If (Test-Path "C:\Program Files\Dassault Systemes\B25\win_b64\code\bin\KillV5Process.exe" -PathType Any -ErrorAction SilentlyContinue) {
				Execute-Process -FilePath "C:\Program Files\Dassault Systemes\B25\win_b64\code\bin\KillV5Process.exe" -WindowStyle Hidden -IgnoreExitCodes "3010" -PassThru -ContinueOnError $true
			}
			If (Test-Path "C:\Program Files\Dassault Systemes\B25\win_b64\code\bin\Uninstall.exe" -PathType Any -ErrorAction SilentlyContinue) {
				Execute-Process -FilePath "C:\Program Files\Dassault Systemes\B25\win_b64\code\bin\Uninstall.exe" -Arguments """C:\Program Files\Dassault Systemes\B25"" ""CODE"" ""GUI"" ""B25"" ""0""" -WindowStyle Normal -IgnoreExitCodes "3010" -PassThru -ContinueOnError $true
			}
			
			ForEach ($GUID in $CATIAGUID) {
				Execute-MSI -Action Uninstall -Path $GUID -ContinueOnError $true
			}
			Start-Sleep -s 10
			ForEach ($GUID in $CATIAGUID) {
				Execute-MSI -Action Uninstall -Path $GUID -ContinueOnError $true
			}
			Start-Sleep -s 10
			
			If (Test-Path "$envCommonDesktop\SNC CATIA - R25.lnk" -PathType Any -ErrorAction SilentlyContinue) {
				Remove-Item "$envCommonDesktop\SNC CATIA - R25.lnk" -ErrorAction SilentlyContinue
			}
			If (Test-Path "$envProgramData\Microsoft\Windows\Start Menu\Programs\CATIA PLM Express" -PathType Any -ErrorAction SilentlyContinue) {
                Remove-Folder -Path "$envProgramData\Microsoft\Windows\Start Menu\Programs\CATIA PLM Express" -ContinueOnError $true
            }
			If (Test-Path "$envProgramData\DassaultSystemes" -PathType Any -ErrorAction SilentlyContinue) {
                Remove-Folder -Path "$envProgramData\DassaultSystemes" -ContinueOnError $true
            }
			If (Test-Path "C:\Program Files\Dassault Systemes\B25" -PathType Any -ErrorAction SilentlyContinue) {
                Remove-Folder -Path "C:\Program Files\Dassault Systemes\B25" -ContinueOnError $true
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
			Start-Sleep -s 10
			Exit-Script -ExitCode "0"
        }
        
        # End uninstall script
        Write-Log "$ApplicationName UNINSTALL: END SCRIPT"
		
		##*===============================================
		##* POST-UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Uninstallation'
		
		# <Perform Post-Uninstallation tasks here>
		
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