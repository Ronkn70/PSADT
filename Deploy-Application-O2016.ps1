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
	[string]$appName = 'Office'
	[string]$appVersion = 'Professional 2016 32bit'
	[string]$appArch = ''
	[string]$appLang = 'EN'
	[string]$appRevision = '01'
	[string]$appScriptVersion = '1.0.0'
	[string]$appScriptDate = '03/14/2016'
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
        Write-Log "MS Office Professional 2016 32bit INSTALL: START SCRIPT"
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

        # Perform MS Office Professional 2016 32bit check task here
        ####Name = "Microsoft Office Professional Plus 2016" /  Version = "16.0.4266.1001"
        ####$installNewOfficeCount = (Get-InstalledApplication -Name "Microsoft Office Professional Plus 2016")
        $installNewOfficeCount = (Get-InstalledApplication -ProductCode "{90160000-0011-0000-0000-0000000FF1CE}")
        Write-Log "MS Office Professional 2016 32bit check results: $installNewOfficeCount"
        If (($installNewOfficeCount | Measure-Object).Count -gt 0) {
            Write-Log "YES, MS Office Professional 2016 32bit is INSTALLED, aborting script with exit code 0."
            
			# Perform Access check task here
			$32bitAccessExe = (Get-ChildItem -Path "C:\Program Files (x86)\Microsoft Office\" -Recurse -Include "msaccess.exe" -Force -ErrorAction SilentlyContinue)
            $32bitAccessDll = (Get-ChildItem -Path "C:\Program Files (x86)\Microsoft Office\" -Recurse -Include "msain.dll" -Force -ErrorAction SilentlyContinue)
			$64bitAccessExe = (Get-ChildItem -Path "C:\Program Files\Microsoft Office\" -Recurse -Include "msaccess.exe" -Force -ErrorAction SilentlyContinue)           
			$64bitAccessDll = (Get-ChildItem -Path "C:\Program Files\Microsoft Office\" -Recurse -Include "msain.dll" -Force -ErrorAction SilentlyContinue)
			If ((($32bitAccessExe | Measure-Object).Count -gt 0) -and (($32bitAccessDll | Measure-Object).Count -gt 0)) {
				Write-Log "YES, Access IS previously INSTALLED."
				$AccessCount = "1"
   				Write-Log "Access value: $AccessCount"
            } ElseIf ((($64bitAccessExe | Measure-Object).Count -gt 0) -and (($64bitAccessDll | Measure-Object).Count -gt 0)) {
				Write-Log "YES, Access IS previously INSTALLED."
				$AccessCount = "1"
				Write-Log "Access value: $AccessCount"

			} Else {
				Write-Log "NO, Access NOT previously INSTALLED."
				$AccessCount = "0"
				Write-Log "Access value: $AccessCount"
			}
			
			# Perform OneNote check task here
			$32bitOneNote = (Get-ChildItem -Path "C:\Program Files (x86)\Microsoft Office\" -Recurse -Include "onenote.exe" -Force -ErrorAction SilentlyContinue)
			$64bitOneNote = (Get-ChildItem -Path "C:\Program Files\Microsoft Office\" -Recurse -Include "onenote.exe" -Force -ErrorAction SilentlyContinue)
			If ((($32bitOneNote | Measure-Object).Count -gt 0) -or (($64bitOneNote | Measure-Object).Count -gt 0)) {
				Write-Log "YES, OneNote IS previously INSTALLED."
				$OneNoteCount = "1"
				Write-Log "OneNote value: $OneNoteCount"
			} Else {
				Write-Log "NO, OneNote NOT previously INSTALLED."
				$OneNoteCount = "0"
				Write-Log "OneNote value: $OneNoteCount"
			}
			
			If (($AccessCount -eq 1) -or ($OneNoteCount -eq 1)) {
				Show-InstallationWelcome -CloseApps "excel,groove,onenote,infopath,outlook,mspub,powerpnt,winword,winproj,visio,msaccess,nlnotes,notes,notes2,notes3,acrobat,acrord32" -BlockExecution -AllowDeferCloseApps -AllowDefer -DeferTimes 25 -DeferDays 7 -PersistPrompt -MinimizeWindows $false
				Show-InstallationProgress -StatusMessage "UPDATING Microsoft Office Professional 2016 32bit. `r`n`r`nThis may take some time. Please wait..." -WindowLocation 'Default' -TopMost $true
				# If Access and/or OneNote is INSTALLED REMOVE
				Write-Log "YES, OneNote or Access is INSTALLED, REMOVE features."
				Execute-Process -FilePath "$dirFiles\setup.exe" -Arguments "/modify ProPlus /config $dirFiles\ProPlus.WW\none.xml" -WindowStyle Hidden -IgnoreExitCodes "3010" -PassThru
				CMD.EXE /C "C:\Windows\System32\msiexec.exe /p $dirFiles\2016_03_15_MS_Office_2016_32bit_PRO_Silent.msp"				
				
				If (Test-Path "C:\Windows\SNC\Office\OFFICE_ACCESS_*.snc" -PathType Any -ErrorAction SilentlyContinue) {
					Remove-Item "C:\Windows\SNC\Office\OFFICE_ACCESS_*.snc" -ErrorAction SilentlyContinue
				}
				$nowDate=Get-Date -format "yyyy_MM_dd"
				If ( -Not (Test-Path C:\Windows\SNC -PathType Any)) {
					New-Item -Path C:\Windows\SNC -ItemType Directory -Force -ErrorAction SilentlyContinue
				}
				If ( -Not (Test-Path C:\Windows\SNC\Office -PathType Any)) {
					New-Item -Path C:\Windows\SNC\Office -ItemType Directory -Force -ErrorAction SilentlyContinue
				}
				$accessFileName="OFFICE_ACCESS_NO_"+$env:computername+"_"+$nowDate+".SNC"
				New-Item C:\Windows\SNC\Office\$accessFileName -type file -force -ErrorAction SilentlyContinue

				Unblock-AppExecution
				Show-InstallationPrompt -Message "Microsoft Office Professional 2016 32bit UPDATE applied. `r`n`r`nIn order to begin using the software, please reboot your machine at your earliest convenience to complete the update." -ButtonRightText "OK" -Icon Information -NoWait
				Start-Sleep -s 10
				Exit-Script -ExitCode "3010" 
			} Else {				
				Show-InstallationPrompt -Message "Microsoft Office Professional 2016 32bit is previously INSTALLED. NO changes made. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
				Start-Sleep -s 10
				Exit-Script -ExitCode "0" 
			}
        } Else {
            Write-Log "NO, MS Office Professional 2016 32bit is NEEDED."
        }
        
        $explorerRunning1 = (Get-Process explorer -ea 0) 
        Write-Log "RUNNING explorer BEGIN results: $explorerRunning1"
        
        # Show Welcome Message, close processes, allow up to 7 day deferral, and persist the prompt
        Show-InstallationWelcome -CloseApps "iexplore,excel,groove,onenote,infopath,outlook,mspub,powerpnt,winword,winproj,visio,msaccess,nlnotes,notes,notes2,notes3,acrobat,acrord32" -BlockExecution -AllowDeferCloseApps -AllowDefer -DeferTimes 25 -DeferDays 7 -PersistPrompt -MinimizeWindows $false
        Show-InstallationProgress -StatusMessage "CHECKING installed Microsoft Office Products. `r`n`r`nThis may take some time. Please wait..." -WindowLocation 'Default' -TopMost $true
        Start-Sleep -s 10
        		
		##*===============================================
		##* INSTALLATION 
		##*===============================================
		[string]$installPhase = 'Installation'
		
		## <Perform Installation tasks here>

        Write-Log "START installing MS Office Professional 2016 32bit."
        
		# Perform MS Office 2010 Pro Registry check
        [string]$regKeySNCorp = 'Registry::HKEY_LOCAL_MACHINE\Software\SNCorp'
		[string]$regKeyMSOffice2010 = 'Registry::HKEY_LOCAL_MACHINE\Software\SNCorp\MSOffice2010'
		[string]$regKeyMSOffice2013 = 'Registry::HKEY_LOCAL_MACHINE\Software\SNCorp\MSOffice2013'
		[string]$regKeyMSOffice2016 = 'Registry::HKEY_LOCAL_MACHINE\Software\SNCorp\MSOffice2016'
        [string]$regKeyMSOffice2010Pro = 'Registry::HKEY_LOCAL_MACHINE\Software\SNCorp\MSOffice2010\OfficePro2010'
		[string]$regKeyMSOffice2013Pro = 'Registry::HKEY_LOCAL_MACHINE\Software\SNCorp\MSOffice2013\OfficePro2013'
		[string]$regKeyMSOffice2016Pro = 'Registry::HKEY_LOCAL_MACHINE\Software\SNCorp\MSOffice2016\OfficePro2016'

        #Write a log if MS Office Pro Key is found
		If (Test-Path "C:\Windows\SNC\Office\OFFICE_ACCESS_YES_*.snc" -PathType Any -ErrorAction SilentlyContinue) {
            Write-Log "YES, Access IS previously INSTALLED."
			$AccessCount = "1"
			Write-Log "Access value: $AccessCount"
        } ElseIf (Test-Path "C:\Windows\SNC\Office\OFFICE_ACCESS_NO_*.snc" -PathType Any -ErrorAction SilentlyContinue) {
            Write-Log "NO, Access NOT previously INSTALLED."
			$AccessCount = "0"
			Write-Log "Access value: $AccessCount"
        } ElseIf (Test-Path -Path $regKeyMSOffice2010Pro -ErrorAction 'SilentlyContinue') {
		    Write-Log -Message 'YES, MSOffice2010Pro Registry Value IS found.'
            $AccessCount = "0"
            Write-Log "Access value: $AccessCount"
		} ElseIf (Test-Path -Path $regKeyMSOffice2013Pro -ErrorAction 'SilentlyContinue') {
		    Write-Log -Message 'YES, MSOffice2013Pro Registry Value IS found.'
            $AccessCount = "0"
            Write-Log "Access value: $AccessCount"
		} ElseIf (Test-Path -Path $regKeyMSOffice2016Pro -ErrorAction 'SilentlyContinue') {
		    Write-Log -Message 'YES, MSOffice2016Pro Registry Value IS found.'
            $AccessCount = "0"
            Write-Log "Access value: $AccessCount"
		} Else {  
			# Perform Access check task here
			$32bitAccessExe = (Get-ChildItem -Path "C:\Program Files (x86)\Microsoft Office\" -Recurse -Include "msaccess.exe" -Force -ErrorAction SilentlyContinue)
            $32bitAccessDll = (Get-ChildItem -Path "C:\Program Files (x86)\Microsoft Office\" -Recurse -Include "msain.dll" -Force -ErrorAction SilentlyContinue)
			$64bitAccessExe = (Get-ChildItem -Path "C:\Program Files\Microsoft Office\" -Recurse -Include "msaccess.exe" -Force -ErrorAction SilentlyContinue)           
			$64bitAccessDll = (Get-ChildItem -Path "C:\Program Files\Microsoft Office\" -Recurse -Include "msain.dll" -Force -ErrorAction SilentlyContinue)
			If ((($32bitAccessExe | Measure-Object).Count -gt 0) -and (($32bitAccessDll | Measure-Object).Count -gt 0)) {
				Write-Log "YES, Access IS previously INSTALLED."
				$AccessCount = "1"
   				Write-Log "Access value: $AccessCount"
            } ElseIf ((($64bitAccessExe | Measure-Object).Count -gt 0) -and (($64bitAccessDll | Measure-Object).Count -gt 0)) {
				Write-Log "YES, Access IS previously INSTALLED."
				$AccessCount = "1"
				Write-Log "Access value: $AccessCount"

			} Else {
				Write-Log "NO, Access NOT previously INSTALLED."
				$AccessCount = "0"
				Write-Log "Access value: $AccessCount"
			}
        }
		        
        # Perform OneNote check task here
        $32bitOneNote = (Get-ChildItem -Path "C:\Program Files (x86)\Microsoft Office\" -Recurse -Include "onenote.exe" -Force -ErrorAction SilentlyContinue)
        $64bitOneNote = (Get-ChildItem -Path "C:\Program Files\Microsoft Office\" -Recurse -Include "onenote.exe" -Force -ErrorAction SilentlyContinue)
        If ((($32bitOneNote | Measure-Object).Count -gt 0) -or (($64bitOneNote | Measure-Object).Count -gt 0)) {
            Write-Log "YES, OneNote IS previously INSTALLED."
            $OneNoteCount = "1"
            Write-Log "OneNote value: $OneNoteCount"
        } Else {
            Write-Log "NO, OneNote NOT previously INSTALLED."
            $OneNoteCount = "0"
            Write-Log "OneNote value: $OneNoteCount"
        }

        # Perform Enterprise 2007 uninstallation tasks here
        [string]$2007EntLoc = "\\salpsccmpss01\cmsource$\core\applications\Microsoft\Office\2007\MSOffice2007_Enterprise"
		$installed2007EntCount = (Get-InstalledApplication -Name "Microsoft Office Enterprise 2007")
		If (($installed2007EntCount | Measure-Object).Count -gt 0) {
			Write-Log "UNINSTALLING MS Office 2007 Enterprise."
            Show-InstallationProgress -StatusMessage "MS Office 2007 Enterprise is detected. `r`n`r`nRemoval in progress. Please wait..." -WindowLocation 'Default' -TopMost $true
            Execute-Process -Filepath "$2007EntLoc\setup.exe" -Arguments "/uninstall enterprise /config $2007EntLoc\Enterprise.WW\silentuninstallconfig.xml" -Windowstyle Hidden -IgnoreExitCodes "3010" -PassThru
            Start-Sleep -s 10
		}
		
		# Perform Professional 2007 uninstallation tasks here
        [string]$2007ProPlusLoc = "\\salpsccmpss01\cmsource$\core\applications\Microsoft\Office\2007\MSOffice2007_ProPlus"
		$installed2007ProCount = (Get-InstalledApplication -Name "Microsoft Office Professional Plus 2007")
		If (($installed2007ProCount | Measure-Object).Count -gt 0) {
			Write-Log "UNINSTALLING MS Office 2007 Professional."
            Show-InstallationProgress -StatusMessage "MS Office 2007 Professional is detected. `r`n`r`nRemoval in progress. Please wait..." -WindowLocation 'Default' -TopMost $true
            Execute-Process -Filepath "$2007ProPlusLoc\setup.exe" -Arguments "/uninstall ProPlus /config $2007ProPlusLoc\ProPlus.WW\silentuninstallconfig.xml" -Windowstyle Hidden -IgnoreExitCodes "3010" -PassThru
            Start-Sleep -s 10
		}
		
		# Perform Standard 2007 uninstallation tasks here
        [string]$2007StdLoc = "\\salpsccmpss01\cmsource$\core\applications\Microsoft\Office\2007\MSOffice2007_Std"
		$installed2007StdCount = (Get-InstalledApplication -Name "Microsoft Office Standard 2007")
		If (($installed2007StdCount | Measure-Object).Count -gt 0) {
			Write-Log "UNINSTALLING MS Office 2007 Standard."
            Show-InstallationProgress -StatusMessage "MS Office 2007 Standard is detected. `r`n`r`nRemoval in progress. Please wait..." -WindowLocation 'Default' -TopMost $true
            Execute-Process -Filepath "$2007StdLoc\setup.exe" -Arguments "/uninstall Standard /config $2007StdLoc\Standard.WW\silentuninstallconfig.xml" -Windowstyle Hidden -IgnoreExitCodes "3010" -PassThru
            Start-Sleep -s 10
		}
        
        # Perform professional 2010 uninstallation tasks here
        $installpro2010Count = (Get-InstalledApplication -Name "Microsoft Office Professional Plus 2010")
        If (($installpro2010Count | Measure-Object).Count -gt 0) {
            Write-Log "UNINSTALLING MS Office 2010 Professional."
            Show-InstallationProgress -StatusMessage "MS Office 2010 Professional is detected. `r`n`r`nRemoval in progress. Please wait..." -WindowLocation 'Default' -TopMost $true
            Execute-Process -FilePath "$dirFiles\2010Professional\setup.exe" -Arguments "/uninstall proplus /config $dirFiles\2010Professional\proplus.ww\silentuninstallconfig.xml" -WindowStyle Hidden -IgnoreExitCodes "3010" -PassThru
            Start-Sleep -s 10
        }

        # Perform standard 2010 uninstallation tasks here
        $installstd2010Count = (Get-InstalledApplication -Name "Microsoft Office Standard 2010")
        If (($installstd2010Count | Measure-Object).Count -gt 0) {
            Write-Log "UNINSTALLING MS Office 2010 Standard."
            Show-InstallationProgress -StatusMessage "MS Office 2010 Standard is detected. `r`n`r`nRemoval in progress. Please wait..." -WindowLocation 'Default' -TopMost $true
            Execute-Process -FilePath "$dirFiles\2010Standard\setup.exe" -Arguments "/uninstall standard /config $dirFiles\2010Standard\standard.ww\silentuninstallconfig.xml" -WindowStyle Hidden -IgnoreExitCodes "3010" -PassThru
            Start-Sleep -s 10
        }

        # Perform professional 2013 uninstallation tasks here
        [string]$2013ProLoc = "\\salpsccmpss01\cmsource$\core\applications\Microsoft\Office\2013\Office2013_32bit_Professional_SP1\Files"
        $installpro2013Count = (Get-InstalledApplication -Name "Microsoft Office Professional Plus 2013")
        If (($installpro2013Count | Measure-Object).Count -gt 0) {
            Write-Log "UNINSTALLING MS Office 2013 Professional."
            Show-InstallationProgress -StatusMessage "MS Office 2013 Professional is detected. `r`n`r`nRemoval in progress. Please wait..." -WindowLocation 'Default' -TopMost $true
            Execute-Process -Filepath "$2013ProLoc\setup.exe" -Arguments "/uninstall ProPlus /config $2013ProLoc\proplus.ww\silentuninstallconfig.xml" -Windowstyle Hidden -IgnoreExitCodes "3010" -PassThru
            Start-Sleep -s 10
        }

        # Perform standard 2013 uninstallation tasks here
        [string]$2013StdLoc = "\\salpsccmpss01\cmsource$\core\applications\Microsoft\Office\2013\Office2013_32bit_Standard_SP1\Files"
        $installstd2013Count = (Get-InstalledApplication -Name "Microsoft Office Standard 2013")
        If (($installstd2013Count | Measure-Object).Count -gt 0) {
            Write-Log "UNINSTALLING MS Office 2013 Standard."
            Show-InstallationProgress -StatusMessage "MS Office 2013 Standard is detected. `r`n`r`nRemoval in progress. Please wait..." -WindowLocation 'Default' -TopMost $true
            Execute-Process -Filepath "$2013StdLoc\setup.exe" -Arguments "/uninstall Standard /config $2013StdLoc\standard.ww\silentuninstallconfig.xml" -Windowstyle Hidden -IgnoreExitCodes "3010" -PassThru
            Start-Sleep -s 10
        }

		# Perform visio viewer uninstallation tasks here
		$install32bitVisioViewer = (Get-InstalledApplication -Name "Microsoft Visio Viewer 2013")
        Write-Log "MS Visio Viewer 2013 check results: $install32bitVisioViewer"
        If (($install32bitVisioViewer | Measure-Object).Count -gt 0) {
            Write-Log "YES, MS Visio Viewer 2013 is INSTALLED, uninstall 64bit and install 32bit."
            #Execute-MSI -Action Uninstall -Path "{90140000-0057-0000-0000-0000000FF1CE}" -ContinueOnError $true
            Execute-MSI -Action Uninstall -Path "{95150000-0052-0409-1000-0000000FF1CE}" -ContinueOnError $true
			Remove-MSIApplications "Microsoft Visio Viewer 2013" -ErrorAction SilentlyContinue -ContinueOnError $true
            Start-Sleep -s 10
            Execute-Process -FilePath "$dirFiles\32bitvisioviewer\visioviewer32bit.exe" -Arguments "/quiet /norestart /passive" -WindowStyle Hidden -IgnoreExitCodes "3010" -PassThru -ContinueOnError $true
            Start-Sleep -s 10
        } Else {
            $installVisio = (Get-InstalledApplication -Name "Microsoft Visio")
            $installVisioOffice = (Get-InstalledApplication -Name "Microsoft Office Visio")
            Write-Log "MS Visio check results: $installVisio"
            Write-Log "MS Visio Office check results: $installVisioOffice"
            If ((($installVisio | Measure-Object).Count -gt 0) -or (($installVisioOffice | Measure-Object).Count -gt 0)) {
                Write-Log "NO, MS Visio Viewer 2013 is NOT installed."
            } Else {
                #Execute-MSI -Action Uninstall -Path "{90140000-0057-0000-0000-0000000FF1CE}" -ContinueOnError $true
                Execute-MSI -Action Uninstall -Path "{95150000-0052-0409-1000-0000000FF1CE}" -ContinueOnError $true
                Remove-MSIApplications "Microsoft Visio Viewer 2013" -ErrorAction SilentlyContinue -ContinueOnError $true
                Start-Sleep -s 10
                Execute-Process -FilePath "$dirFiles\32bitvisioviewer\visioviewer32bit.exe" -Arguments "/quiet /norestart /passive" -WindowStyle Hidden -IgnoreExitCodes "3010" -PassThru -ContinueOnError $true
                Start-Sleep -s 10
            }
        }

		# Perform standard 2016 uninstallation tasks here
        [string]$2016StdLoc = "\\salpsccmpss01\cmsource$\core\applications\Microsoft\Office\2016\Office2016_Standard_32bit\Files"
        $installstd2016Count = (Get-InstalledApplication -Name "Microsoft Office Standard 2016")
        If (($installstd2016Count | Measure-Object).Count -gt 0) {
            Write-Log "UNINSTALLING MS Office 2016 Standard."
            Show-InstallationProgress -StatusMessage "MS Office 2016 Standard is detected. `r`n`r`nRemoval in progress. Please wait..." -WindowLocation 'Default' -TopMost $true
            Execute-Process -Filepath "$2016StdLoc\setup.exe" -Arguments "/uninstall Standard /config $2016StdLoc\standard.ww\silentuninstallconfig.xml" -Windowstyle Hidden -IgnoreExitCodes "3010" -PassThru
            Start-Sleep -s 10
        }
		
		# Perform installation tasks here
        Write-Log "INSTALLING MS Office Professional 2016 32bit."
        Show-InstallationProgress -StatusMessage "INSTALLING Microsoft Office Professional 2016 32bit. `r`n`r`nThis may take some time. Please wait..." -WindowLocation 'Default' -TopMost $true
        If ($AccessCount -eq 1) {
			If ($OneNoteCount -eq 1) {
				$office2016=Execute-Process -FilePath "$dirFiles\setup.exe" -Arguments "/adminfile $dirFiles\2016_03_15_MS_Office_2016_32bit_PRO_Silent_With_OneNote_Access.msp" -WindowStyle Hidden -IgnoreExitCodes "3010" -PassThru
			} Else {
				$office2016=Execute-Process -FilePath "$dirFiles\setup.exe" -Arguments "/adminfile $dirFiles\2016_03_15_MS_Office_2016_32bit_PRO_Silent_With_Access.msp" -WindowStyle Hidden -IgnoreExitCodes "3010" -PassThru
			}			
            Start-Sleep -s 10
            If (Test-Path -Path $regKeyMSOffice2010Pro -ErrorAction 'SilentlyContinue') {
			    Write-Log -Message 'YES, MSOffice2010Pro Registry key IS found.  It will be removed.'
			    Remove-RegistryKey -Key $regKeyMSOffice2010Pro -Recurse -ErrorAction 'SilentlyContinue'
                Remove-RegistryKey -Key $regKeyMSOffice2010 -Recurse -ErrorAction 'SilentlyContinue'
			    Write-Log -Message 'MSOffice2010Pro Registry key is REMOVED.'
			} Else {
			    Write-Log -Message 'NO, MSOffice2010Pro Registry key is NOT found. No acation is required'
			}
            If (Test-Path -Path $regKeyMSOffice2013Pro -ErrorAction 'SilentlyContinue') {
			    Write-Log -Message 'YES, MSOffice2013Pro Registry key IS found.  It will be removed.'
			    Remove-RegistryKey -Key $regKeyMSOffice2013Pro -Recurse -ErrorAction 'SilentlyContinue'
                Remove-RegistryKey -Key $regKeyMSOffice2013 -Recurse -ErrorAction 'SilentlyContinue'
			    Write-Log -Message 'MSOffice2013Pro Registry key is REMOVED.'
			} Else {
			    Write-Log -Message 'NO, MSOffice2013Pro Registry key is NOT found. No acation is required'
			}
            If (Test-Path -Path $regKeyMSOffice2016Pro -ErrorAction 'SilentlyContinue') {
			    Write-Log -Message 'MSOffice2016Pro Registry Value found.'
			} Else {
                If (-not (Test-Path -Path $regKeyMSOffice2016 -ErrorAction 'SilentlyContinue')) { 
				    Write-Log -Message 'MSOffice2016 Registry Key NOT found. Create MSOffice2016 Registry Key'
				    Set-RegistryKey -Key $regKeyMSOffice2016 -ErrorAction 'SilentlyContinue'
				    Write-Log -Message 'MSOffice2016 Registry Key created'
			    }
			    Write-Log -Message 'MSOffice2016Pro registry Key NOT found. Create MSOffice2016Pro registry key'
			    Set-RegistryKey -Key $regKeyMSOffice2016Pro -ErrorAction 'SilentlyContinue'
			    Write-Log -Message 'MSOffice Pro 2016 Registry Key created'
            }

            If (Test-Path "C:\Windows\SNC\Office\OFFICE_ACCESS_*.snc" -PathType Any -ErrorAction SilentlyContinue) {
                Remove-Item "C:\Windows\SNC\Office\OFFICE_ACCESS_*.snc" -ErrorAction SilentlyContinue
            }
            $nowDate=Get-Date -format "yyyy_MM_dd"
            If (Test-Path "C:\Program Files (x86)\Microsoft Office" -PathType Any -ErrorAction SilentlyContinue) {
                $accessVersion=(Get-ChildItem -Path "C:\Program Files (x86)\Microsoft Office\" -Recurse -Include "msaccess.exe" -Force -ErrorAction SilentlyContinue) | foreach-object ($_) {[System.Diagnostics.FileVersionInfo]::GetVersionInfo($_).FileVersion}
            } ElseIf (Test-Path "C:\Program Files\Microsoft Office" -PathType Any -ErrorAction SilentlyContinue) {
                $accessVersion=(Get-ChildItem -Path "C:\Program Files\Microsoft Office\" -Recurse -Include "msaccess.exe" -Force -ErrorAction SilentlyContinue) | foreach-object ($_) {[System.Diagnostics.FileVersionInfo]::GetVersionInfo($_).FileVersion}
            }
            $accessFileName="OFFICE_ACCESS_YES_"+$accessVersion+"_"+$env:computername+"_"+$nowDate+".SNC"
            If ( -Not (Test-Path C:\Windows\SNC -PathType Any)) {
                New-Item -Path C:\Windows\SNC -ItemType Directory -Force -ErrorAction SilentlyContinue
            }
            If ( -Not (Test-Path C:\Windows\SNC\Office -PathType Any)) {
                New-Item -Path C:\Windows\SNC\Office -ItemType Directory -Force -ErrorAction SilentlyContinue
            }
            New-Item C:\Windows\SNC\Office\$accessFileName -type file -force -ErrorAction SilentlyContinue

        } ElseIf ($OneNoteCount -eq 1) {
			$office2016=Execute-Process -FilePath "$dirFiles\setup.exe" -Arguments "/adminfile $dirFiles\2016_03_15_MS_Office_2016_32bit_PRO_Silent_With_OneNote.msp" -WindowStyle Hidden -IgnoreExitCodes "3010" -PassThru

            If (Test-Path "C:\Windows\SNC\Office\OFFICE_ACCESS_*.snc" -PathType Any -ErrorAction SilentlyContinue) {
				Remove-Item "C:\Windows\SNC\Office\OFFICE_ACCESS_*.snc" -ErrorAction SilentlyContinue
			}
			$nowDate=Get-Date -format "yyyy_MM_dd"
            If ( -Not (Test-Path C:\Windows\SNC -PathType Any)) {
                New-Item -Path C:\Windows\SNC -ItemType Directory -Force -ErrorAction SilentlyContinue
            }
            If ( -Not (Test-Path C:\Windows\SNC\Office -PathType Any)) {
                New-Item -Path C:\Windows\SNC\Office -ItemType Directory -Force -ErrorAction SilentlyContinue
            }
            $accessFileName="OFFICE_ACCESS_NO_"+$env:computername+"_"+$nowDate+".SNC"
            New-Item C:\Windows\SNC\Office\$accessFileName -type file -force -ErrorAction SilentlyContinue

        } Else {
			$office2016=Execute-Process -FilePath "$dirFiles\setup.exe" -Arguments "/adminfile $dirFiles\2016_03_15_MS_Office_2016_32bit_PRO_Silent.msp" -WindowStyle Hidden -IgnoreExitCodes "3010" -PassThru
            
            If (Test-Path "C:\Windows\SNC\Office\OFFICE_ACCESS_*.snc" -PathType Any -ErrorAction SilentlyContinue) {
				Remove-Item "C:\Windows\SNC\Office\OFFICE_ACCESS_*.snc" -ErrorAction SilentlyContinue
			}
			$nowDate=Get-Date -format "yyyy_MM_dd"
            If ( -Not (Test-Path C:\Windows\SNC -PathType Any)) {
                New-Item -Path C:\Windows\SNC -ItemType Directory -Force -ErrorAction SilentlyContinue
            }
            If ( -Not (Test-Path C:\Windows\SNC\Office -PathType Any)) {
                New-Item -Path C:\Windows\SNC\Office -ItemType Directory -Force -ErrorAction SilentlyContinue
            }
            $accessFileName="OFFICE_ACCESS_NO_"+$env:computername+"_"+$nowDate+".SNC"
            New-Item C:\Windows\SNC\Office\$accessFileName -type file -force -ErrorAction SilentlyContinue

        }
        Start-Sleep -s 10
        $office2016ExitCode = $office2016.ExitCode
        Write-Log "OFFICE2016 EXIT CODE: $office2016ExitCode"

        If (($office2016ExitCode -eq 0) -or ($office2016ExitCode -eq 3010)) {
            Write-Log "OFFICE2016 install COMPLETE." 
            Show-InstallationProgress -StatusMessage "Microsoft Office Professional 2016 32bit install COMPLETE. `r`n`r`nPlease wait..."
            Start-Sleep -s 10
        } ElseIf ($office2016ExitCode -eq 1603) {
            Write-Log "OFFICE2016 INSTALL FAILED."
            Write-Log "OFFICE2016 install NOT COMPLETE." 
            Show-InstallationPrompt -Message "Microsoft Office Professional 2016 32bit installation encounter an error.  `r`n`r`nPlease reboot your machine at your earliest convenience in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait
			Start-Sleep -s 10
            Exit-Script -ExitCode "1603"    
        } ElseIf ($office2016ExitCode -eq 1618) {
            Write-Log "OFFICE2016 INSTALL FAILED."
            Write-Log "OFFICE2016 install NOT COMPLETE." 
            Show-InstallationPrompt -Message "Microsoft Office Professional 2016 32bit installation encounter an error.  `r`n`r`nPlease reboot your machine at your earliest convenience in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait
			Start-Sleep -s 10
            Exit-Script -ExitCode "1618" 
        } Else {
            Write-Log "OFFICE2016 INSTALL ERROR."
            Write-Log "OFFICE2016 install NOT COMPLETE." 
            Show-InstallationPrompt -Message "Microsoft Office Professional 2016 32bit installation encounter an error.  `r`n`r`nPlease reboot your machine at your earliest convenience in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait
			Start-Sleep -s 10
            Exit-Script -ExitCode $office2016ExitCode
        }

        Write-Log "FIX Office 2016 Registry for ALL USERS."
        [scriptblock]$HKCURegistrySettings = {
            Remove-RegistryKey -Key 'HKCU\Software\Microsoft\Office\16.0' -Recurse -SID $UserProfile.SID -ContinueOnError $true
            Remove-RegistryKey -Key 'HKCU\Software\Wow6432Node\Microsoft\Office\16.0' -Recurse -SID $UserProfile.SID -ContinueOnError $true
            Remove-RegistryKey -Key 'HKCU\Software\Microsoft\Office\16.0' -SID $UserProfile.SID -ContinueOnError $true
            Remove-RegistryKey -Key 'HKCU\Software\Wow6432Node\Microsoft\Office\16.0' -SID $UserProfile.SID -ContinueOnError $true
            #####YES, allow online content
            Set-RegistryKey -Key 'HKCU\Software\Microsoft\Office\16.0\Common\Internet' -Name 'UseOnlineContent' -Value 2 -Type DWord -SID $UserProfile.SID -ContinueOnError $true
            Set-RegistryKey -Key 'HKCU\Software\Wow6432Node\Microsoft\Office\16.0\Common\Internet' -Name 'UseOnlineContent' -Value 2 -Type DWord -SID $UserProfile.SID -ContinueOnError $true
            #####NO, do not allow online content
            #Set-RegistryKey -Key 'HKCU\Software\Microsoft\Office\16.0\Common\Internet' -Name 'UseOnlineContent' -Value 1 -Type DWord -SID $UserProfile.SID -ContinueOnError $true
            #Set-RegistryKey -Key 'HKCU\Software\Wow6432Node\Microsoft\Office\16.0\Common\Internet' -Name 'UseOnlineContent' -Value 1 -Type DWord -SID $UserProfile.SID -ContinueOnError $true
            #####REMOVE Office Upload Center start up
            Remove-RegistryKey -Key 'HKCU\Software\Microsoft\Windows\CurrentVersion\Run' -Name 'OfficeSyncProcess' -SID $UserProfile.SID -ContinueOnError $true
            Remove-RegistryKey -Key 'HKCU\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Run' -Name 'OfficeSyncProcess' -SID $UserProfile.SID -ContinueOnError $true
            Set-RegistryKey -Key 'HKCU\Software\Microsoft\Office\16.0\OneNote' -Name 'FirstBootStatus' -Value 16777473 -Type DWord -SID $UserProfile.SID -ContinueOnError $true
        }
        Invoke-HKCURegistrySettingsForAllUsers -RegistrySettings $HKCURegistrySettings
        #####REMOVE Office Upload Center start up
        Remove-RegistryKey -Key 'HKLM\Software\Microsoft\Windows\CurrentVersion\Run' -Name 'OfficeSyncProcess' -ContinueOnError $true
        Remove-RegistryKey -Key 'HKLM\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Run' -Name 'OfficeSyncProcess' -ContinueOnError $true

        ##*===============================================
		##* POST-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Installation'
		
		## <Perform Post-Installation tasks here>
        
        Write-Log "MS Office Professional 2016 32-bit INSTALL COMPLETED." 
        Unblock-AppExecution 
        
        # Display a message at the end of the install
		$installNewOfficeCount2 = (Get-InstalledApplication -ProductCode "{90160000-0011-0000-0000-0000000FF1CE}")
        Write-Log "MS Office Professional 2016 32-bit check results: $installNewOfficeCount2"
        If (($installNewOfficeCount2 | Measure-Object).Count -gt 0) {
            # Display a message at the end of the install
            Unblock-AppExecution
            $explorerRunning2 = (Get-Process explorer -ea 0) 
            Write-Log "RUNNING explorer END results: $explorerRunning2"
            If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "Microsoft Office Professional 2016 32-bit installation complete.  `r`n`r`nIn order to begin using the new software, please click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then reboot your machine in order to finish the install." -ButtonRightText "OK" -Icon Information -NoWait
				Start-Sleep -s 30
				Exit-Script -ExitCode "3010"
            } Else {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "Microsoft Office Professional 2016 32-bit installation complete.  `r`n`r`nIn order to begin using the new software, please reboot your machine at your earliest convenience to complete the install." -ButtonRightText "OK" -Icon Information -NoWait
				Start-Sleep -s 30
				Exit-Script -ExitCode "3010"
            }
        } Else {
            # Display a message at the end of the install
            $explorerRunning2 = (Get-Process explorer -ea 0) 
            Write-Log "RUNNING explorer END results: $explorerRunning2"
            If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "Microsoft Office Professional 2016 32-bit installation encounter an error.  `r`n`r`nPlease click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then reboot your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait
				Start-Sleep -s 30
                Exit-Script -ExitCode "1"
            } Else {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "Microsoft Office Professional 2016 32-bit installation encounter an error.  `r`n`r`nPlease reboot your machine at your earliest convenience in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait
				Start-Sleep -s 30
                Exit-Script -ExitCode "1"
            }
        }

        # End install script
        Write-Log "MS Office Professional 2016 32-bit INSTALL: END SCRIPT"

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
        Write-Log "MS Office Professional 2016 32bit UNINSTALL: START SCRIPT"
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

		# Perform uninstall MS Office Professional 2016 32-bit check
        $installNewOfficeCount = (Get-InstalledApplication -ProductCode "{90160000-0011-0000-0000-0000000FF1CE}")
        Write-Log "MS Office Professional 2016 32-bit check results: $installNewOfficeCount"
        If (($installNewOfficeCount | Measure-Object).Count -gt 0) {
        
            Write-Log "YES, MS Office Professional 2016 32-bit is INSTALLED." 
            
            $explorerRunning1 = (Get-Process explorer -ea 0) 
            Write-Log "RUNNING explorer BEGIN results: $explorerRunning1"
            
            # Show Welcome Message, close processes, allow up to 7 day deferral, and persist the prompt
            Show-InstallationWelcome -CloseApps "iexplore,excel,groove,onenote,infopath,outlook,mspub,powerpnt,winword,winproj,visio,msaccess,nlnotes,notes,notes2,notes3,acrobat,acrord32" -BlockExecution -AllowDeferCloseApps -AllowDefer -DeferTimes 25 -DeferDays 7 -PersistPrompt -MinimizeWindows $false
            Show-InstallationProgress -StatusMessage "UNINSTALLING Microsoft Office Professional 2016 32-bit. `r`n`r`nThis may take some time. Please wait..." -WindowLocation 'Default' -TopMost $true
            Start-Sleep -s 10
                                   
            # Perform visio viewer uninstallation tasks here
			$install32bitVisioViewer = (Get-InstalledApplication -Name "Microsoft Visio Viewer 2013")
            Write-Log "MS Visio Viewer 2013 check results: $install32bitVisioViewer"
            If (($install32bitVisioViewer | Measure-Object).Count -gt 0) {
                Write-Log "YES, MS Visio Viewer 2013 is INSTALLED, uninstall 64bit and install 32bit."
                #Execute-MSI -Action Uninstall -Path "{90140000-0057-0000-0000-0000000FF1CE}" -ContinueOnError $true
                Execute-MSI -Action Uninstall -Path "{95150000-0052-0409-1000-0000000FF1CE}" -ContinueOnError $true
                Remove-MSIApplications "Microsoft Visio Viewer 2013" -ErrorAction SilentlyContinue -ContinueOnError $true
                Start-Sleep -s 10
                Execute-Process -FilePath "$dirFiles\32bitvisioviewer\visioviewer32bit.exe" -Arguments "/quiet /norestart /passive" -WindowStyle Hidden -IgnoreExitCodes "3010" -PassThru -ContinueOnError $true
                Start-Sleep -s 10
            } Else {
                Write-Log "NO, MS Visio Viewer 2013 is NOT installed."
            }
			
			Write-Log "UNINSTALLING MS Office Professional 2016 32-bit."          
            Show-InstallationWelcome -CloseApps "OfficeClickToRun,OfficeC2RClient,AppVShNotify,setup" -Silent -MinimizeWindows $false
            Stop-Service -name "ClickToRunSvc" -Force -ErrorAction SilentlyContinue
            $office2016=Execute-Process -FilePath "$dirFiles\setup.exe" -Arguments "/uninstall ProPlus /config $dirFiles\proplus.ww\silentuninstallconfig.xml" -WindowStyle Hidden -IgnoreExitCodes "3010" -PassThru
            Start-Sleep -s 10
            $office2016ExitCode = $office2016.ExitCode
            Write-Log "OFFICE2016 EXIT CODE: $office2016ExitCode"

            If (($office2016ExitCode -eq 0) -or ($office2016ExitCode -eq 3010)) {
                Write-Log "OFFICE2016 uninstall COMPLETE." 
                Show-InstallationProgress -StatusMessage "Microsoft Office Professional 2016 32-bit uninstall COMPLETE. `r`n`r`nPlease wait..."
                Start-Sleep -s 10
            } ElseIf ($office2016ExitCode -eq 1603) {
                Write-Log "OFFICE2016 UNINSTALL FAILED."
                Write-Log "OFFICE2016 uninstall NOT COMPLETE." 
                Show-InstallationPrompt -Message "Microsoft Office Professional 2016 32-bit uninstallation encounter an error.  `r`n`r`nPlease reboot your machine at your earliest convenience in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait
				Start-Sleep -s 10
                Exit-Script -ExitCode "1603"    
            } ElseIf ($office2016ExitCode -eq 1618) {
                Write-Log "OFFICE2016 UNINSTALL FAILED."
                Write-Log "OFFICE2016 uninstall NOT COMPLETE." 
                Show-InstallationPrompt -Message "Microsoft Office Professional 2016 32-bit uninstallation encounter an error.  `r`n`r`nPlease reboot your machine at your earliest convenience in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait
				Start-Sleep -s 10
                Exit-Script -ExitCode "1618" 
            } Else {
                Write-Log "OFFICE2016 UNINSTALL ERROR."
                Write-Log "OFFICE2016 UNinstall NOT COMPLETE." 
                Show-InstallationPrompt -Message "Microsoft Office Professional 2016 32-bit uninstallation encounter an error.  `r`n`r`nPlease reboot your machine at your earliest convenience in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait
				Start-Sleep -s 10
                Exit-Script -ExitCode $office2016ExitCode
            }

            Write-Log "FIX Office 2016 Registry for ALL USERS."
            [scriptblock]$HKCURegistrySettings = {
                Remove-RegistryKey -Key 'HKCU\Software\Microsoft\Office\16.0\Common\Internet' -Name 'UseOnlineContent' -SID $UserProfile.SID -ContinueOnError $true
                Remove-RegistryKey -Key 'HKCU\Software\Wow6432Node\Microsoft\Office\16.0\Common\Internet' -Name 'UseOnlineContent' -SID $UserProfile.SID -ContinueOnError $true
                Remove-RegistryKey -Key 'HKCU\Software\Microsoft\Windows\CurrentVersion\Run' -Name 'OfficeSyncProcess' -SID $UserProfile.SID -ContinueOnError $true
                Remove-RegistryKey -Key 'HKCU\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Run' -Name 'OfficeSyncProcess' -SID $UserProfile.SID -ContinueOnError $true
            }
            Invoke-HKCURegistrySettingsForAllUsers -RegistrySettings $HKCURegistrySettings
            Remove-RegistryKey -Key 'HKLM\Software\Microsoft\Windows\CurrentVersion\Run' -Name 'OfficeSyncProcess' -ContinueOnError $true
            Remove-RegistryKey -Key 'HKLM\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Run' -Name 'OfficeSyncProcess' -ContinueOnError $true         

			# Perform MS Office 2010, 2013, & 2016 Pro Registry check
			[string]$regKeySNCorp = 'Registry::HKEY_LOCAL_MACHINE\Software\SNCorp'
			[string]$regKeyMSOffice2010 = 'Registry::HKEY_LOCAL_MACHINE\Software\SNCorp\MSOffice2010'
			[string]$regKeyMSOffice2013 = 'Registry::HKEY_LOCAL_MACHINE\Software\SNCorp\MSOffice2013'
			[string]$regKeyMSOffice2016 = 'Registry::HKEY_LOCAL_MACHINE\Software\SNCorp\MSOffice2016'
			[string]$regKeyMSOffice2010Pro = 'Registry::HKEY_LOCAL_MACHINE\Software\SNCorp\MSOffice2010\OfficePro2010'
			[string]$regKeyMSOffice2013Pro = 'Registry::HKEY_LOCAL_MACHINE\Software\SNCorp\MSOffice2013\OfficePro2013'
			[string]$regKeyMSOffice2016Pro = 'Registry::HKEY_LOCAL_MACHINE\Software\SNCorp\MSOffice2016\OfficePro2016'

            If (Test-Path -Path $regKeyMSOffice2010Pro -ErrorAction 'SilentlyContinue') {
			    Write-Log -Message 'YES, MSOffice2010Pro Registry key IS found.  It will be removed.'
			    Remove-RegistryKey -Key $regKeyMSOffice2010Pro -Recurse -ErrorAction 'SilentlyContinue'
			    Remove-RegistryKey -Key $regKeyMSOffice2010 -Recurse -ErrorAction 'SilentlyContinue'
			    Write-Log -Message 'MSOffice2010Pro Registry key is REMOVED.'
			} Else {
			    Write-Log -Message 'NO, MSOffice2010Pro Registry key is NOT found. No acation is required'
			}
            If (Test-Path -Path $regKeyMSOffice2013Pro -ErrorAction 'SilentlyContinue') {
			    Write-Log -Message 'YES, MSOffice2013Pro Registry key IS found.  It will be removed.'
			    Remove-RegistryKey -Key $regKeyMSOffice2013Pro -Recurse -ErrorAction 'SilentlyContinue'
			    Remove-RegistryKey -Key $regKeyMSOffice2013 -Recurse -ErrorAction 'SilentlyContinue'
			    Write-Log -Message 'MSOffice2013Pro Registry key is REMOVED.'
			} Else {
			    Write-Log -Message 'NO, MSOffice2013Pro Registry key is NOT found. No acation is required'
			}
			If (Test-Path -Path $regKeyMSOffice2016Pro -ErrorAction 'SilentlyContinue') {
			    Write-Log -Message 'YES, MSOffice2016Pro Registry key IS found.  It will be removed.'
			    Remove-RegistryKey -Key $regKeyMSOffice2016Pro -Recurse -ErrorAction 'SilentlyContinue'
			    Remove-RegistryKey -Key $regKeyMSOffice2016 -Recurse -ErrorAction 'SilentlyContinue'
			    Write-Log -Message 'MSOffice2016Pro Registry key is REMOVED.'
			} Else {
			    Write-Log -Message 'NO, MSOffice2016Pro Registry key is NOT found. No acation is required'
			}
			If (Test-Path "C:\Windows\SNC\Office\OFFICE_ACCESS_*.snc" -PathType Any -ErrorAction SilentlyContinue) {
                Remove-Item "C:\Windows\SNC\Office\OFFICE_ACCESS_*.snc" -ErrorAction SilentlyContinue
            }
            
			Write-Log "MS Office Professional 2016 32-bit UNINSTALL COMPLETED." 
            # Display a message at the end of the uninstall
            Unblock-AppExecution
            $explorerRunning2 = (Get-Process explorer -ea 0) 
            Write-Log "RUNNING explorer END results: $explorerRunning2"
            $installNewOfficeCount2 = (Get-InstalledApplication -ProductCode "{90160000-0011-0000-0000-0000000FF1CE}")
            Write-Log "MS Office Professional 2016 32-bit check results: $installNewOfficeCount2"
            If (($installNewOfficeCount2 | Measure-Object).Count -gt 0) {
                If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "Microsoft Office Professional 2016 32-bit uninstallation encounter an error.  `r`n`r`nPlease click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then reboot your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait
					Start-Sleep -s 30
                    Exit-Script -ExitCode "1"
                } Else {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "Microsoft Office Professional 2016 32-bit uninstallation encounter an error.  `r`n`r`nPlease reboot your machine at your earliest convenience in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait
					Start-Sleep -s 30
                    Exit-Script -ExitCode "1"
                }
            } Else {
                If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "Microsoft Office Professional 2016 32-bit uninstallation complete.  `r`n`r`nPlease click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then reboot your machine in order to finish the uninstall." -ButtonRightText "OK" -Icon Information -NoWait
					Start-Sleep -s 30
                    Exit-Script -ExitCode "3010"
                } Else {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "Microsoft Office Professional 2016 32-bit uninstallation complete.  `r`n`r`nPlease reboot your machine at your earliest convenience to complete the uninstall." -ButtonRightText "OK" -Icon Information -NoWait
					Start-Sleep -s 30
                    Exit-Script -ExitCode "3010"
                }
            }
        } Else {
            Write-Log "NO, MS Office Professional 2016 32-bit is NOT installed."
            Unblock-AppExecution
            Show-InstallationPrompt -Message "Microsoft Office Professional 2016 32-bit NOT installed, nothing to uninstall. NO changes made. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
			Start-Sleep -s 7
			Exit-Script -ExitCode "0"
        }
        
        # End uninstall script
        Write-Log "MS Office Professional 2016 32-bit UNINSTALL: END SCRIPT"
				
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