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
	[string]$appVendor = 'Mozilla'
	[string]$appName = 'Firefox'
	[string]$appVersion = '48.0 32bit'
	[string]$appArch = 'X86'
	[string]$appLang = 'EN'
	[string]$appRevision = '01'
	[string]$appScriptVersion = '1.0.0'
	[string]$appScriptDate = '06/15/2016'
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
        Write-Log "Mozilla Firefox 48.0 32bit INSTALL: START SCRIPT"

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

        # Perform Mozilla Firefox 48.0 32bit check task here
        $installNewFirefoxCount = (Get-InstalledApplication -Name "Mozilla Firefox 48.0")
        Write-Log "Mozilla Firefox 48.0 32bit check results: $installNewFirefoxCount"
        If (($installNewFirefoxCount | Measure-Object).Count -gt 0) {
            Write-Log "YES, Mozilla Firefox 48.0 32bit is INSTALLED, aborting script with exit code 0."
            Show-InstallationPrompt -Message "Mozilla Firefox 48.0 32bit is previously INSTALLED. NO changes made. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait -Timeout "180"
            Start-Sleep -s 30
            Exit-Script -ExitCode "0" 
        } Else {
            Write-Log "NO, Mozilla Firefox 48.0 32bit is NEEDED."
        }		

        Write-Log "START SHOWING INSTALLATION PROGRESS MESSAGES."
        Start-Sleep -s 60
        
        # Show Welcome Message, close processes, allow up to 7 day deferral, and persist the prompt
        Show-InstallationWelcome -CloseApps "firefox,dummyfirefox" -AllowDeferCloseApps -AllowDefer -DeferTimes 25 -DeferDays 7 -PersistPrompt -BlockExecution -MinimizeWindows $false
        Start-Sleep -s 30

        # Perform OLD Mozilla Firefox check task here
        $installOldFirefox = (Get-InstalledApplication -Name "Mozilla Firefox")
        If ($envOSName -eq "Microsoft Windows XP Professional") {
            Write-Log "OS Name: $envOSName"
            Write-Log "Windows XP, continue with the script."
            $installOldFirefoxFile = (Get-ChildItem -Path "C:\Documents and Settings\" -Recurse -Include "firefox.exe" -Force -ErrorAction SilentlyContinue)
        } ElseIf ($envOSName -eq "Microsoft(R) Windows(R) XP Professional x64 Edition") {
            Write-Log "OS Name: $envOSName"
            Write-Log "Windows XP, continue with the script."
            $installOldFirefoxFile = (Get-ChildItem -Path "C:\Documents and Settings\" -Recurse -Include "firefox.exe" -Force -ErrorAction SilentlyContinue)
        } Else {
            Write-Log "OS Name: $envOSName"
            Write-Log "Windows 7, continue with the script."
            $installOldFirefoxFile = (Get-ChildItem -Path "C:\Users\" -Recurse -Include "firefox.exe" -Force -ErrorAction SilentlyContinue)
        }
        Write-Log "OLD Mozilla Firefox check results: $installOldFirefox"
        Write-Log "OLD Mozilla Firefox File check results: $installOldFirefoxFile"
        If ((($installOldFirefox | Measure-Object).Count -gt 0) -or (($installOldFirefoxFile | Measure-Object).Count -gt 0)) {
            Write-Log "YES, OLD Mozilla Firefox is INSTALLED."
            Write-Log "UNINSTALLING OLD Mozilla Firefox."

            # Show Progress Message (with the default message) and Show-InstallationWelcome triggered
            Show-InstallationProgress -StatusMessage "UNINSTALLING previous Mozilla Firefox. `r`n`r`nThis may take some time. Please wait..."
            Start-Sleep -s 30
                
            # Perform uninstallation tasks here
            # Check to see if the shorcuts exists, if so delete
            If (Test-Path "C:\Program Files (x86)\Mozilla Firefox\uninstall\uninst.exe" -PathType Any -ErrorAction SilentlyContinue) {
                Execute-Process -FilePath "C:\Program Files (x86)\Mozilla Firefox\uninstall\uninst.exe" -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
            }
            If (Test-Path "C:\Program Files\Mozilla Firefox\uninstall\uninst.exe" -PathType Any -ErrorAction SilentlyContinue) {
                Execute-Process -FilePath "C:\Program Files\Mozilla Firefox\uninstall\uninst.exe" -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
            }
            If (Test-Path "C:\Program Files (x86)\Mozilla Firefox\uninstall\helper.exe" -PathType Any -ErrorAction SilentlyContinue) {
                Execute-Process -FilePath "C:\Program Files (x86)\Mozilla Firefox\uninstall\helper.exe" -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
            }
            If (Test-Path "C:\Program Files\Mozilla Firefox\uninstall\helper.exe" -PathType Any -ErrorAction SilentlyContinue) {
                Execute-Process -FilePath "C:\Program Files\Mozilla Firefox\uninstall\helper.exe" -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
            }
            If (Test-Path "C:\Program Files (x86)\Mozilla Maintenance Service\uninstall.exe" -PathType Any -ErrorAction SilentlyContinue) {
                Execute-Process -FilePath "C:\Program Files (x86)\Mozilla Maintenance Service\uninstall.exe" -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
            }
            If (Test-Path "C:\Program Files\Mozilla Maintenance Service\uninstall.exe" -PathType Any -ErrorAction SilentlyContinue) {
                Execute-Process -FilePath "C:\Program Files\Mozilla Maintenance Service\uninstall.exe" -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
            }
            If (Test-Path "C:\Documents and Settings\All Users\Application Data\Mozilla Firefox\uninstall\helper.exe" -PathType Any -ErrorAction SilentlyContinue) {
                Execute-Process -FilePath "C:\Documents and Settings\All Users\Application Data\Mozilla Firefox\uninstall\helper.exe" -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
            }
            If (Test-Path "C:\Documents and Settings\All Users\Application Data\Mozilla Maintenance Service\uninstall.exe" -PathType Any -ErrorAction SilentlyContinue) {
                Execute-Process -FilePath "C:\Documents and Settings\All Users\Application Data\Mozilla Maintenance Service\uninstall.exe" -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
            }
            If (Test-Path "C:\Program Files (x86)\Firefox Developer Edition\uninstall\helper.exe" -PathType Any -ErrorAction SilentlyContinue) {
                Execute-Process -FilePath "C:\Program Files (x86)\Firefox Developer Edition\uninstall\helper.exe" -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
            }
            If (Test-Path "C:\Program Files\Firefox Developer Edition\uninstall\helper.exe" -PathType Any -ErrorAction SilentlyContinue) {
                Execute-Process -FilePath "C:\Program Files\Firefox Developer Edition\uninstall\helper.exe" -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
            }
            If (Test-Path "C:\Program Files (x86)\Mozilla Firefox 4.0 Beta 11\uninstall\helper.exe" -PathType Any -ErrorAction SilentlyContinue) {
                Execute-Process -FilePath "C:\Program Files (x86)\Mozilla Firefox 4.0 Beta 11\uninstall\helper.exe" -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
            }
            If (Test-Path "C:\Program Files\Mozilla Firefox 4.0 Beta 11\uninstall\helper.exe" -PathType Any -ErrorAction SilentlyContinue) {
                Execute-Process -FilePath "C:\Program Files\Mozilla Firefox 4.0 Beta 11\uninstall\helper.exe" -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
            }
            If (Test-Path "C:\Users\Public\Desktop\Mozilla Firefox.lnk" -PathType Any -ErrorAction SilentlyContinue) {
                Remove-Item -Force -Path "C:\Users\Public\Desktop\Mozilla Firefox.lnk" -ErrorAction SilentlyContinue
            }
            If (Test-Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Mozilla Firefox.lnk" -PathType Any -ErrorAction SilentlyContinue) {
                Remove-Item -Force -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Mozilla Firefox.lnk" -ErrorAction SilentlyContinue
            }
            If (Test-Path "C:\Users" -PathType Container -ErrorAction SilentlyContinue) {
                Get-ChildItem -Path "C:\Users" -Include "*" -Force -ErrorAction SilentlyContinue | ForEach-Object ($_) {
                    $path1 = $_.FullName + "\AppData\Local\Mozilla Firefox\uninstall\uninst.exe"
                    $path2 = $_.FullName + "\AppData\Local\Mozilla Firefox\uninstall\helper.exe"
                    $path3 = $_.FullName + "\AppData\Local\Mozilla Maintenance Service\uninstall.exe"
                    $path4 = $_.FullName + "\Desktop\Mozilla Firefox.lnk"
                    $path5 = $_.FullName + "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Mozilla Firefox.lnk"
                    $path6 = $_.FullName + "\AppData\Local\Mozilla Firefox"
                    If (Test-Path $path1 -PathType Any -ErrorAction SilentlyContinue) {
                        Execute-Process -FilePath $path1 -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
                    }
                    If (Test-Path $path2 -PathType Any -ErrorAction SilentlyContinue) {
                        Execute-Process -FilePath $path2 -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
                    }
                    If (Test-Path $path3 -PathType Any -ErrorAction SilentlyContinue) {
                        Execute-Process -FilePath $path3 -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
                    }
                    If (Test-Path $path4 -PathType Any -ErrorAction SilentlyContinue) {
                        Remove-Item -Force -Path $path4 -ErrorAction SilentlyContinue
                    }
                    If (Test-Path $path5 -PathType Any -ErrorAction SilentlyContinue) {
                        Remove-Item -Force -Path $path5 -ErrorAction SilentlyContinue
                    }
                    If (Test-Path $path6 -PathType Container -ErrorAction SilentlyContinue) {
                        Remove-Item -Force -Recurse -Path $path6 -ErrorAction SilentlyContinue
                        Get-ChildItem -Path $path6 -Include * -Recurse -ErrorAction SilentlyContinue | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
                        Remove-Folder -Path $path6 -ContinueOnError $true
                    }
                }
            }

            # Clean Up Registry(s)
            Write-Log "Clean up Mozilla Firefox registry." 
            Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ -Recurse -Include "Mozilla Firefox*" -ErrorAction SilentlyContinue | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
            $Temp = New-PSDrive HKU Registry HKEY_USERS -ErrorAction SilentlyContinue
            Get-ChildItem -Path HKU:\ -ErrorAction SilentlyContinue | ForEach-Object {
                [string]$pathRegistry1 = $_.Name
                $pathRegistry2 = $pathRegistry1 -replace "HKEY_USERS","HKU:"
                Get-ChildItem -Path $pathRegistry2\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ -Recurse -Include "Mozilla Firefox*" -ErrorAction SilentlyContinue | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
            }
            If ($envOSArchitecture -eq "64-Bit") {
                Get-ChildItem -Path HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ -Recurse -Include "Mozilla Firefox*" -ErrorAction SilentlyContinue | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
                $Temp = New-PSDrive HKU Registry HKEY_USERS -ErrorAction SilentlyContinue
                Get-ChildItem -Path HKU:\ -ErrorAction SilentlyContinue | ForEach-Object {
                    [string]$pathRegistry3 = $_.Name
                    $pathRegistry4 = $pathRegistry3 -replace "HKEY_USERS","HKU:"
                    Get-ChildItem -Path $pathRegistry4\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ -Recurse -Include "Mozilla Firefox*" -ErrorAction SilentlyContinue | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
                }
            }

            Write-Log "OLD Mozilla Firefox UNINSTALL COMPLETED."
            # Display a message at the end of the install
            Show-InstallationProgress -StatusMessage "Previous Mozilla Firefox uninstall COMPLETE. `r`n`r`nPlease wait..."
            Start-Sleep -s 30
        } Else {
            Unblock-AppExecution
            Write-Log "NO, OLD Mozilla Firefox is NOT installed."
            Show-InstallationProgress -StatusMessage "NO previous Mozilla Firefox installed. `r`n`r`nPlease wait..."
            Start-Sleep -s 30
        }
		
		##*===============================================
		##* INSTALLATION 
		##*===============================================
		[string]$installPhase = 'Installation'
		
		## <Perform Installation tasks here>

        Write-Log "START installing Mozilla Firefox 48.0 32bit."
        # Show Progress Message (with the default message) and Show-InstallationWelcome triggered
        Show-InstallationProgress -StatusMessage "INSTALLING Mozilla Firefox 48.0 32bit. `r`n`r`nThis may take some time. Please wait..."        
        Start-Sleep -s 30
        
        # Perform installation tasks here
        Write-Log "INSTALLING Mozilla Firefox 48.0 32bit."
        Execute-Process -FilePath "$dirFiles\Firefox Setup 48.0.exe" -Arguments "-ms /INI=$dirFiles\config.ini" -Windowstyle Hidden -IgnoreExitCodes "3010"
        
        # Copy configuration files
        Write-Log "COPYING CUSTOM configuration files for Mozilla Firefox."
        If (Test-Path "C:\Program Files (x86)\Mozilla Firefox" -PathType Any) {
            Copy-Item "$dirFiles\autoconfig.js" "C:\Program Files (x86)\Mozilla Firefox\defaults\pref" -Force -ErrorAction SilentlyContinue
            Copy-Item "$dirFiles\Firefox.cfg" "C:\Program Files (x86)\Mozilla Firefox" -Force -ErrorAction SilentlyContinue
            Copy-Item "$dirFiles\override.ini" "C:\Program Files (x86)\Mozilla Firefox\browser" -Force -ErrorAction SilentlyContinue
        }
        If (Test-Path "C:\Program Files\Mozilla Firefox" -PathType Any) {
            Copy-Item "$dirFiles\autoconfig.js" "C:\Program Files\Mozilla Firefox\defaults\pref" -Force -ErrorAction SilentlyContinue
            Copy-Item "$dirFiles\Firefox.cfg" "C:\Program Files\Mozilla Firefox" -Force -ErrorAction SilentlyContinue
            Copy-Item "$dirFiles\override.ini" "C:\Program Files\Mozilla Firefox\browser" -Force -ErrorAction SilentlyContinue
        }
        
        # Perform uninstallation Mozilla Maintenance Service tasks here
        # Check to see if the shorcuts exists, if so delete
        If (Test-Path "C:\Program Files (x86)\Mozilla Maintenance Service\uninstall.exe" -PathType Any -ErrorAction SilentlyContinue) {
            Execute-Process -FilePath "C:\Program Files (x86)\Mozilla Maintenance Service\uninstall.exe" -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
        }
        If (Test-Path "C:\Program Files\Mozilla Maintenance Service\uninstall.exe" -PathType Any -ErrorAction SilentlyContinue) {
            Execute-Process -FilePath "C:\Program Files\Mozilla Maintenance Service\uninstall.exe" -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
        }
        
## Import-Module "$dirFiles\PinnedApplications.psm1" -ErrorAction SilentlyContinue
## Set-PinnedApplication -Action UnPinFromTaskbar -FilePath "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"

        ##*===============================================
		##* POST-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Installation'

		## <Perform Post-Installation tasks here>

        Write-Log "Mozilla Firefox 48.0 32bit INSTALL COMPLETED."
        Unblock-AppExecution 
        
        # Display a message at the end of the install
        $installNewFirefoxCount2 = (Get-InstalledApplication -Name "Mozilla Firefox 48.0")
        If (($installNewFirefoxCount2 | Measure-Object).Count -gt 0) {
            # Display a message at the end of the install
            Unblock-AppExecution
            $explorerRunning2 = (Get-Process explorer -ea 0) 
            Write-Log "RUNNING explorer END results: $explorerRunning2"
            If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "Mozilla Firefox 48.0 32bit installation complete.  `r`n`r`nIn order to begin using the new client, please click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then reboot your machine in order to finish the install." -ButtonRightText "OK" -Icon Information -NoWait -Timeout "180"
				Start-Sleep -s 30
				Exit-Script -ExitCode "3010"
            } Else {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "Mozilla Firefox 48.0 32bit installation complete. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait -Timeout "180"
				Start-Sleep -s 30
				Exit-Script -ExitCode "3010"
            }
        } Else {
            # Display a message at the end of the install
            $explorerRunning2 = (Get-Process explorer -ea 0) 
            Write-Log "RUNNING explorer END results: $explorerRunning2"
            If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "Mozilla Firefox 48.0 32bit installation encounter an error.  `r`n`r`nPlease click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then reboot your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait -Timeout "180"
				Start-Sleep -s 30
                Exit-Script -ExitCode "1"
            } Else {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "Mozilla Firefox 48.0 32bit installation encounter an error.  `r`n`r`nPlease reboot your machine at your earliest convenience in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait -Timeout "180"
				Start-Sleep -s 30
                Exit-Script -ExitCode "1"
            }
        }

        # End install script
        Write-Log "Mozilla Firefox INSTALL: END SCRIPT"

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
        Write-Log "Mozilla Firefox UNINSTALL: START SCRIPT"
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

        # Perform Mozilla Firefox check task here
        $installOldFirefox = (Get-InstalledApplication -Name "Mozilla Firefox")
        # Perform OS name check
        If ($envOSName -eq "Microsoft Windows XP Professional") {
            Write-Log "OS Name: $envOSName"
            Write-Log "Windows XP, continue with the script."
            $installOldFirefoxFile = (Get-ChildItem -Path "C:\Documents and Settings\" -Recurse -Include "firefox.exe" -Force -ErrorAction SilentlyContinue)
        } ElseIf ($envOSName -eq "Microsoft(R) Windows(R) XP Professional x64 Edition") {
            Write-Log "OS Name: $envOSName"
            Write-Log "Windows XP, continue with the script."
            $installOldFirefoxFile = (Get-ChildItem -Path "C:\Documents and Settings\" -Recurse -Include "firefox.exe" -Force -ErrorAction SilentlyContinue)
        } Else {
            Write-Log "OS Name: $envOSName"
            Write-Log "Windows 7, continue with the script."
            $installOldFirefoxFile = (Get-ChildItem -Path "C:\Users\" -Recurse -Include "firefox.exe" -Force -ErrorAction SilentlyContinue)
        }
        Write-Log "Mozilla Firefox check results: $installOldFirefox"
        Write-Log "Mozilla Firefox File check results: $installOldFirefoxFile"
        If ((($installOldFirefox | Measure-Object).Count -gt 0) -or (($installOldFirefoxFile | Measure-Object).Count -gt 0)) {
            Write-Log "YES, Mozilla Firefox is INSTALLED."
            Write-Log "UNINSTALLING Mozilla Firefox."

            Write-Log "START SHOWING UNINSTALLATION PROGRESS MESSAGES."
            Start-Sleep -s 60
            
            # Show Welcome Message, close processes, allow up to 7 day deferral, and persist the prompt
            Show-InstallationWelcome -CloseApps "firefox,dummyfirefox" -AllowDeferCloseApps -AllowDefer -DeferTimes 25 -DeferDays 7 -PersistPrompt -BlockExecution -MinimizeWindows $false
            Start-Sleep -s 30
            
            # Show Progress Message (with the default message) and Show-InstallationWelcome triggered
            Show-InstallationProgress -StatusMessage "UNINSTALLING Mozilla Firefox. `r`n`r`nThis may take some time. Please wait..."
            Start-Sleep -s 30
            
            # Perform uninstallation tasks here
            # Check to see if the shorcuts exists, if so delete
            If (Test-Path "C:\Program Files (x86)\Mozilla Firefox\uninstall\uninst.exe" -PathType Any -ErrorAction SilentlyContinue) {
                Execute-Process -FilePath "C:\Program Files (x86)\Mozilla Firefox\uninstall\uninst.exe" -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
            }
            If (Test-Path "C:\Program Files\Mozilla Firefox\uninstall\uninst.exe" -PathType Any -ErrorAction SilentlyContinue) {
                Execute-Process -FilePath "C:\Program Files\Mozilla Firefox\uninstall\uninst.exe" -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
            }
            If (Test-Path "C:\Program Files (x86)\Mozilla Firefox\uninstall\helper.exe" -PathType Any -ErrorAction SilentlyContinue) {
                Execute-Process -FilePath "C:\Program Files (x86)\Mozilla Firefox\uninstall\helper.exe" -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
            }
            If (Test-Path "C:\Program Files\Mozilla Firefox\uninstall\helper.exe" -PathType Any -ErrorAction SilentlyContinue) {
                Execute-Process -FilePath "C:\Program Files\Mozilla Firefox\uninstall\helper.exe" -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
            }
            If (Test-Path "C:\Program Files (x86)\Mozilla Maintenance Service\uninstall.exe" -PathType Any -ErrorAction SilentlyContinue) {
                Execute-Process -FilePath "C:\Program Files (x86)\Mozilla Maintenance Service\uninstall.exe" -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
            }
            If (Test-Path "C:\Program Files\Mozilla Maintenance Service\uninstall.exe" -PathType Any -ErrorAction SilentlyContinue) {
                Execute-Process -FilePath "C:\Program Files\Mozilla Maintenance Service\uninstall.exe" -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
            }
            If (Test-Path "C:\Documents and Settings\All Users\Application Data\Mozilla Firefox\uninstall\helper.exe" -PathType Any -ErrorAction SilentlyContinue) {
                Execute-Process -FilePath "C:\Documents and Settings\All Users\Application Data\Mozilla Firefox\uninstall\helper.exe" -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
            }
            If (Test-Path "C:\Documents and Settings\All Users\Application Data\Mozilla Maintenance Service\uninstall.exe" -PathType Any -ErrorAction SilentlyContinue) {
                Execute-Process -FilePath "C:\Documents and Settings\All Users\Application Data\Mozilla Maintenance Service\uninstall.exe" -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
            }
            If (Test-Path "C:\Program Files (x86)\Firefox Developer Edition\uninstall\helper.exe" -PathType Any -ErrorAction SilentlyContinue) {
                Execute-Process -FilePath "C:\Program Files (x86)\Firefox Developer Edition\uninstall\helper.exe" -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
            }
            If (Test-Path "C:\Program Files\Firefox Developer Edition\uninstall\helper.exe" -PathType Any -ErrorAction SilentlyContinue) {
                Execute-Process -FilePath "C:\Program Files\Firefox Developer Edition\uninstall\helper.exe" -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
            }
            If (Test-Path "C:\Program Files (x86)\Mozilla Firefox 4.0 Beta 11\uninstall\helper.exe" -PathType Any -ErrorAction SilentlyContinue) {
                Execute-Process -FilePath "C:\Program Files (x86)\Mozilla Firefox 4.0 Beta 11\uninstall\helper.exe" -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
            }
            If (Test-Path "C:\Program Files\Mozilla Firefox 4.0 Beta 11\uninstall\helper.exe" -PathType Any -ErrorAction SilentlyContinue) {
                Execute-Process -FilePath "C:\Program Files\Mozilla Firefox 4.0 Beta 11\uninstall\helper.exe" -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
            }
            If (Test-Path "C:\Users\Public\Desktop\Mozilla Firefox.lnk" -PathType Any -ErrorAction SilentlyContinue) {
                Remove-Item -Force -Path "C:\Users\Public\Desktop\Mozilla Firefox.lnk" -ErrorAction SilentlyContinue
            }
            If (Test-Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Mozilla Firefox.lnk" -PathType Any -ErrorAction SilentlyContinue) {
                Remove-Item -Force -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Mozilla Firefox.lnk" -ErrorAction SilentlyContinue
            }
            If (Test-Path "C:\Users" -PathType Container -ErrorAction SilentlyContinue) {
                Get-ChildItem -Path "C:\Users" -Include "*" -Force -ErrorAction SilentlyContinue | ForEach-Object ($_) {
                    $path1 = $_.FullName + "\AppData\Local\Mozilla Firefox\uninstall\uninst.exe"
                    $path2 = $_.FullName + "\AppData\Local\Mozilla Firefox\uninstall\helper.exe"
                    $path3 = $_.FullName + "\AppData\Local\Mozilla Maintenance Service\uninstall.exe"
                    $path4 = $_.FullName + "\Desktop\Mozilla Firefox.lnk"
                    $path5 = $_.FullName + "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Mozilla Firefox.lnk"
                    $path6 = $_.FullName + "\AppData\Local\Mozilla Firefox"
                    If (Test-Path $path1 -PathType Any -ErrorAction SilentlyContinue) {
                        Execute-Process -FilePath $path1 -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
                    }
                    If (Test-Path $path2 -PathType Any -ErrorAction SilentlyContinue) {
                        Execute-Process -FilePath $path2 -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
                    }
                    If (Test-Path $path3 -PathType Any -ErrorAction SilentlyContinue) {
                        Execute-Process -FilePath $path3 -Arguments "/S" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
                    }
                    If (Test-Path $path4 -PathType Any -ErrorAction SilentlyContinue) {
                        Remove-Item -Force -Path $path4 -ErrorAction SilentlyContinue
                    }
                    If (Test-Path $path5 -PathType Any -ErrorAction SilentlyContinue) {
                        Remove-Item -Force -Path $path5 -ErrorAction SilentlyContinue
                    }
                    If (Test-Path $path6 -PathType Container -ErrorAction SilentlyContinue) {
                        Remove-Item -Force -Recurse -Path $path6 -ErrorAction SilentlyContinue
                        Get-ChildItem -Path $path6 -Include * -Recurse -ErrorAction SilentlyContinue | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
                        Remove-Folder -Path $path6 -ContinueOnError $true
                    }
                }
            }
            
            # Clean Up Registry(s)
            Write-Log "Clean up Mozilla Firefox registry." 
            Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ -Recurse -Include "Mozilla Firefox*" -ErrorAction SilentlyContinue | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
            $Temp = New-PSDrive HKU Registry HKEY_USERS -ErrorAction SilentlyContinue
            Get-ChildItem -Path HKU:\ -ErrorAction SilentlyContinue | ForEach-Object {
                [string]$pathRegistry1 = $_.Name
                $pathRegistry2 = $pathRegistry1 -replace "HKEY_USERS","HKU:"
                Get-ChildItem -Path $pathRegistry2\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ -Recurse -Include "Mozilla Firefox*" -ErrorAction SilentlyContinue | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
            }
            If ($envOSArchitecture -eq "64-Bit") {
                Get-ChildItem -Path HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ -Recurse -Include "Mozilla Firefox*" -ErrorAction SilentlyContinue | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
                $Temp = New-PSDrive HKU Registry HKEY_USERS -ErrorAction SilentlyContinue
                Get-ChildItem -Path HKU:\ -ErrorAction SilentlyContinue | ForEach-Object {
                    [string]$pathRegistry3 = $_.Name
                    $pathRegistry4 = $pathRegistry3 -replace "HKEY_USERS","HKU:"
                    Get-ChildItem -Path $pathRegistry4\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ -Recurse -Include "Mozilla Firefox*" -ErrorAction SilentlyContinue | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
                }
            }
            
            Write-Log "Mozilla Firefox UNINSTALL COMPLETED." 
            Unblock-AppExecution
            $explorerRunning2 = (Get-Process explorer -ea 0) 
            Write-Log "RUNNING explorer END results: $explorerRunning2"
            # Display a message at the end of the uninstall
            # Perform Mozilla Firefox check #2 task here
            $installOldFirefox2 = (Get-InstalledApplication -Name "Mozilla Firefox")
            # Perform OS name check
            If ($envOSName -eq "Microsoft Windows XP Professional") {
                Write-Log "OS Name: $envOSName"
                Write-Log "Windows XP, continue with the script."
                $installOldFirefoxFile2 = (Get-ChildItem -Path "C:\Documents and Settings\" -Recurse -Include "firefox.exe" -Force -ErrorAction SilentlyContinue)
            } ElseIf ($envOSName -eq "Microsoft(R) Windows(R) XP Professional x64 Edition") {
                Write-Log "OS Name: $envOSName"
                Write-Log "Windows XP, continue with the script."
                $installOldFirefoxFile2 = (Get-ChildItem -Path "C:\Documents and Settings\" -Recurse -Include "firefox.exe" -Force -ErrorAction SilentlyContinue)
            } Else {
                Write-Log "OS Name: $envOSName"
                Write-Log "Windows 7, continue with the script."
                $installOldFirefoxFile2 = (Get-ChildItem -Path "C:\Users\" -Recurse -Include "firefox.exe" -Force -ErrorAction SilentlyContinue)
            }
            Write-Log "Mozilla Firefox check results: $installOldFirefox2"
            Write-Log "Mozilla Firefox File check results: $installOldFirefoxFile2"
            If ((($installOldFirefox2 | Measure-Object).Count -gt 0) -or (($installOldFirefoxFile2 | Measure-Object).Count -gt 0)) {
                If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "Mozilla Firefox uninstallation encounter an error.  `r`n`r`nPlease click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then reboot your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait -Timeout "180"
					Start-Sleep -s 30
                    Exit-Script -ExitCode "1"
                } Else {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "Mozilla Firefox uninstallation encounter an error.  `r`n`r`nPlease reboot your machine at your earliest convenience in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait -Timeout "180"
					Start-Sleep -s 30
                    Exit-Script -ExitCode "1"
                }
            } Else {
                If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "Mozilla Firefox uninstallation complete.  `r`n`r`nPlease click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then reboot your machine in order to finish the uninstall." -ButtonRightText "OK" -Icon Information -NoWait -Timeout "180"
					Start-Sleep -s 30
                    Exit-Script -ExitCode "3010"
                } Else {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "Mozilla Firefox uninstallation complete.  `r`n`r`nPlease reboot your machine at your earliest convenience to complete the uninstall." -ButtonRightText "OK" -Icon Information -NoWait -Timeout "180"
					Start-Sleep -s 30
                    Exit-Script -ExitCode "3010"
                }
            }
        } Else {
            Unblock-AppExecution
            Write-Log "NO, Mozilla Firefox is NOT installed."
            Show-InstallationPrompt -Message "Mozilla Firefox NOT installed, nothing to uninstall. NO changes made. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait -Timeout "180"
			Start-Sleep -s 30
            Exit-Script -ExitCode "0"
        }
        
        # End uninstall script
        Write-Log "Mozilla Firefox UNINSTALL: END SCRIPT"
	
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