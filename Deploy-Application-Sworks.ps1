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
	[string]$appVendor = 'Solidworks'
	[string]$appName = '2015 SP5.0'
	[string]$appVersion = ''
	[string]$appArch = 'x64'
	[string]$appLang = 'EN'
	[string]$appRevision = '01'
	[string]$appScriptVersion = '1.0.0'
	[string]$appScriptDate = '06/22/2016'
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
		Write-Log "Solidworks 2015 SP5.0 INSTALL: START SCRIPT"
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

        # Perform Solidworks 2015 SP5.0 check task here
        ####Name = "SOLIDWORKS 2015 x64 Edition SP05" /  Version = "23.5.0.81"
        $installNewSWCount = (Get-InstalledApplication -Name "SOLIDWORKS 2015 x64 Edition SP05")
        Write-Log "Solidworks 2015 SP5.0 check results: $installNewSWCount"
        If (($installNewSWCount | Measure-Object).Count -gt 0) {
            Write-Log "YES, Solidworks 2015 SP5.0 is INSTALLED, aborting script with exit code 0."
            Show-InstallationPrompt -Message "Solidworks 2015 SP5.0 is previously INSTALLED. NO changes made. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
            Start-Sleep -s 10
            Exit-Script -ExitCode "0" 
        } Else {
            Write-Log "NO, Solidworks 2015 SP5.0 is NEEDED."
        }
        
        # Get IPV4Address
        Write-Log "Getting IPV4Address."
        If(Test-Connection -ComputerName $envComputerName -Count 1 -ea 0 -quiet) {
            Write-Log "IPV4Address EXISTS."
            $getip = Test-Connection -ComputerName $envComputerName -Count 1
            $ip = $getip.IPV4Address.IPAddressToString
        } Else {
            Write-Log "IPV4Address DOES NOT EXISTS."
            $ip='1.1.1.1'
        }
        Write-Log "IPV4Address is: $ip"
        $ip2 = $ip.split('.') 
        $ip2[-1] = "" 
        $range = $ip2 -join '.'
        Write-Log "IPV4Address range is: $range"

        # Check explorer to see if running
        $explorerRunning1 = (Get-Process explorer -ea 0) 
        Write-Log "RUNNING explorer BEGIN results: $explorerRunning1"
        
        # Stop the services
        Stop-Service -name "WSearch" -Force -ErrorAction SilentlyContinue
        Stop-Service -name "ConisioWebServer" -Force -ErrorAction SilentlyContinue
        Stop-Service -name "ConisioDbServer" -Force -ErrorAction SilentlyContinue
        Stop-Service -name "ArchiveServerService" -Force -ErrorAction SilentlyContinue
        Stop-Service -name "WSearch" -Force -ErrorAction SilentlyContinue
        Start-Sleep -s 10

        # Show Welcome Message, close processes, allow up to 7 day deferral, and persist the prompt
        Show-InstallationWelcome -CloseApps "Search,dummySearch" -Silent -MinimizeWindows $false
        Show-InstallationWelcome -CloseApps "EdmServer,AddInRegSrv64,AddInSrv,CardEdit,ConisioAdmin,ConisioUrl,ConisioWebServer,DbUpdate,EdmServer,FileViewer,Inbox,InventorServer,NetRegSrv,Report,Search,SettingsDialog,TaskExecutor,ViewServer,ViewSetup,VLink,DsgnChkRptView,gabiswengine,LocalSldService,propertyManagerUpload,RTLibraryManager,setcatenv,sldbgproc,sldCostingTemplateEditorAppU,sldexitapp,sldphotoshopcon,sldProcMon,sldShellExtServer,sldu3d,SLDWORKS,sldworks_fs,swShellFileLauncher,swspmanager,UtlReportViewer,eDrawingOfficeAutomator,EModelViewer,AddInRegSrv32,sldtoolboxupdater,sldIM" -AllowDeferCloseApps -AllowDefer -DeferTimes 25 -DeferDays 7 -PersistPrompt -MinimizeWindows $false

        # Show Progress Message (with the default message) and Show-InstallationWelcome triggered
        Show-InstallationProgress -StatusMessage "INSTALLING Solidworks 2015 SP5.0. `r`n`r`nThis may take some time. Please wait..."
        Start-Sleep -s 10
        
		# Perform installation tasks here Office Web Components
        Write-Log "INSTALLING Office Web Components."
        Execute-Process -FilePath "$dirFiles\64bit\OfficeWeb_11\owc11.exe" -Arguments "/quiet" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
        
        # Perform installation tasks here Visual C++ 2005 Redistributable Package
        Write-Log "INSTALLING Visual C++ 2005 Redistributable Package."
        Execute-Process -FilePath "$dirFiles\64bit\Microsoft_C++_2005_Redistributable\vcredist_x86.exe" -Arguments "/q:a /c:""msiexec /i vcredist.msi /quiet /passive /norestart /qn""" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
        Execute-Process -FilePath "$dirFiles\64bit\Microsoft_C++_2005_Redistributable_(x64)\vcredist_x64.exe" -Arguments "/q:a /c:""msiexec /i vcredist.msi /quiet /passive /norestart /qn""" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
        
        # Perform installation tasks here Visual C++ 2008 Redistributable Package
        Write-Log "INSTALLING Visual C++ 2008 Redistributable Package."
        Execute-Process -FilePath "$dirFiles\64bit\Microsoft_C++_2008_Redistributable\vcredist_x86.exe" -Arguments "/q /norestart" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
        Execute-Process -FilePath "$dirFiles\64bit\Microsoft_C++_2008_Redistributable_(x64)\vcredist_x64.exe" -Arguments "/q /norestart" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
        
        # Perform installation tasks here Visual C++ 2010 Redistributable Package
        Write-Log "INSTALLING Visual C++ 2010 Redistributable Package."
        Execute-Process -FilePath "$dirFiles\64bit\Microsoft_C++_2010_Redistributable\vcredist_x86.exe" -Arguments "/q /norestart" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
        Execute-Process -FilePath "$dirFiles\64bit\Microsoft_C++_2010_Redistributable_(x64)\vcredist_x64.exe" -Arguments "/q /norestart" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
        
        # Perform installation tasks here Visual C++ 2012 Redistributable Package
        Write-Log "INSTALLING Visual C++ 2012 Redistributable Package."
        Execute-Process -FilePath "$dirFiles\64bit\Microsoft_C++_2012_Redistributable\vcredist_x86.exe" -Arguments "/install /quiet /norestart" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
        Execute-Process -FilePath "$dirFiles\64bit\Microsoft_C++_2012_Redistributable_(x64)\vcredist_x64.exe" -Arguments "/install /quiet /norestart" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
        
        # Perform installation tasks here Visual Studio Tools For Application
        Write-Log "INSTALLING Visual Studio Tools For Application."
        Execute-MSI -Action Install -Path "$dirFiles\64bit\Microsoft_VSTA\vsta_aide.msi" -Parameters "/quiet /passive /norestart /qn" -ContinueOnError $true

        # Perform installation tasks here Visual Studio Remote Debugger
        Write-Log "INSTALLING Visual Studio Remote Debugger."
        Execute-Process -FilePath "$dirFiles\64bit\VSRemoteDebugger\install.exe" -Arguments "/q /norestart" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true      
        
        # Perform installation tasks here Visual Basic For Applications 7.1
        Write-Log "INSTALLING Visual Basic For Applications 7.1."
        Execute-MSI -Action Install -Path "$dirFiles\64bit\Microsoft_VBA\vba71.msi" -Parameters "/quiet /passive /norestart /qn" -ContinueOnError $true
        Execute-MSI -Action Install -Path "$dirFiles\64bit\Microsoft_VBA_1033\vba71_1033.msi" -Parameters "/quiet /passive /norestart /qn" -ContinueOnError $true
        Execute-MSI -Action Patch -Path "$dirFiles\64bit\Microsoft_VBA_KB2783832\vba71-kb2783832-x64.msp" -Parameters "/quiet /passive /norestart /qn" -ContinueOnError $true

        # Perform installation tasks here .NET Framework 4.5
        Write-Log "INSTALLING .NET Framework 4.5."
        Execute-Process -FilePath "$dirFiles\64bit\.Net_Framework_4.5\dotnetfx45_full_x86_x64.exe" -Arguments "/q /NoSetupVersionCheck /norestart" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true

        # Perform installation tasks here Bonjour Service For Windows
        Write-Log "INSTALLING Bonjour Service For Windows."
        Execute-MSI -Action Install -Path "$dirFiles\64bit\Bonjour\Bonjour64.msi" -Parameters "/quiet /passive /norestart /qn" -ContinueOnError $true

		##*===============================================
		##* INSTALLATION 
		##*===============================================
		[string]$installPhase = 'Installation'
		
		## <Perform Installation tasks here>
        
        Write-Log "START installing Solidworks 2015 SP5.0."
        
        # Perform installation tasks here
        Write-Log "INSTALLING Solidworks 2015 SP5.0."
        Execute-Process -FilePath "$dirFiles\64bit\sldim\sldIM.exe" -Arguments "/adminclient /new /source ""$dirFiles\64bit\AdminDirector.xml"" /norunsw /pushdeployment" -Windowstyle Hidden -IgnoreExitCodes "3010"
        Start-Sleep -s 60
        Wait-Process -name sldIM -ErrorAction SilentlyContinue
        Start-Sleep -s 300
        Wait-Process -name sldIM -ErrorAction SilentlyContinue
        Start-Sleep -s 60
        Wait-Process -name sldIM -ErrorAction SilentlyContinue
		Start-Sleep -s 10
                     
        # # Check to see if the folder exists, if not create it
        # If (Test-Path "C:\MCAD\SWDATA" -PathType Container -ErrorAction SilentlyContinue) {
            # Write-Log "SWDATA folder exists, nothing to create."
        # } Else {
            # Write-Log "SWDATA folder does NOT exists, create it."
            # New-Folder -Path "C:\MCAD\SWDATA" -ContinueOnError $true
        # }

        # # Modifying permissions
        # Write-Log "MODIFYING folder permissions."
        # CMD.EXE /C "ICACLS ""C:\MCAD\SW2015"" /grant Users:(OI)(CI)F /t /c /q"
        # CMD.EXE /C "ICACLS ""C:\MCAD\SWDATA"" /grant Users:(OI)(CI)F /t /c /q"

        # # Check to see if the folder exists, if so delete folder
        # If (Test-Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\SolidWorks 2013" -PathType Container -ErrorAction SilentlyContinue) {
            # Remove-Folder -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\SolidWorks 2013" -ContinueOnError $true
        # }
        # If (Test-Path "C:\MCAD\SW2013" -PathType Container -ErrorAction SilentlyContinue) {
            # Remove-Folder -Path "C:\MCAD\SW2013" -ContinueOnError $true
        # }
        
        If (Test-Path "C:\Users\Public\Desktop" -PathType Container -ErrorAction SilentlyContinue) {
            # Remove-Item -Force -Path "C:\Users\Public\Desktop\SolidWorks Explorer 2011.lnk" -ErrorAction SilentlyContinue
            # Remove-Item -Force -Path "C:\Users\Public\Desktop\SolidWorks eDrawings 2011.lnk" -ErrorAction SilentlyContinue
            # Remove-Item -Force -Path "C:\Users\Public\Desktop\SolidWorks eDrawings 2011 x64 Edition.lnk" -ErrorAction SilentlyContinue
            # Remove-Item -Force -Path "C:\Users\Public\Desktop\SolidWorks eDrawings 2012.lnk" -ErrorAction SilentlyContinue
            # Remove-Item -Force -Path "C:\Users\Public\Desktop\SolidWorks eDrawings 2012 x64 Edition.lnk" -ErrorAction SilentlyContinue
            # Remove-Item -Force -Path "C:\Users\Public\Desktop\SolidWorks Explorer 2013.lnk" -ErrorAction SilentlyContinue
            # Remove-Item -Force -Path "C:\Users\Public\Desktop\SolidWorks eDrawings 2013.lnk" -ErrorAction SilentlyContinue
            # Remove-Item -Force -Path "C:\Users\Public\Desktop\SolidWorks eDrawings 2013 x64 Edition.lnk" -ErrorAction SilentlyContinue
            # Remove-Item -Force -Path "C:\Users\Public\Desktop\eDrawings 2014.lnk" -ErrorAction SilentlyContinue
            # Remove-Item -Force -Path "C:\Users\Public\Desktop\SolidWorks Explorer 2014.lnk" -ErrorAction SilentlyContinue
            # Remove-Item -Force -Path "C:\Users\Public\Desktop\SolidWorks eDrawings 2014.lnk" -ErrorAction SilentlyContinue
            # Remove-Item -Force -Path "C:\Users\Public\Desktop\SolidWorks eDrawings 2014 x64 Edition.lnk" -ErrorAction SilentlyContinue
            # Remove-Item -Force -Path "C:\Users\Public\Desktop\eDrawings 2015.lnk" -ErrorAction SilentlyContinue
            Remove-Item -Force -Path "C:\Users\Public\Desktop\SolidWorks Explorer 2015.lnk" -ErrorAction SilentlyContinue
            # Remove-Item -Force -Path "C:\Users\Public\Desktop\SolidWorks eDrawings 2015.lnk" -ErrorAction SilentlyContinue
            # Remove-Item -Force -Path "C:\Users\Public\Desktop\SolidWorks eDrawings 2015 x64 Edition.lnk" -ErrorAction SilentlyContinue
            # Remove-Item -Force -Path "C:\Users\Public\Desktop\eDrawings 2015.lnk" -ErrorAction SilentlyContinue
            Remove-Item -Force -Path "C:\Users\Public\Desktop\eDrawings 2015 x64 Edition.lnk" -ErrorAction SilentlyContinue
            # Remove-Item -Force -Path "C:\Users\Public\Desktop\SOLIDWORKS Composer 2015 x64 Edition.lnk" -ErrorAction SilentlyContinue
            Remove-Item -Force -Path "C:\Users\Public\Desktop\SOLIDWORKS Composer Player 2015 - x64 Edition.lnk" -ErrorAction SilentlyContinue
            # Remove-Item -Force -Path "C:\Users\Public\Desktop\SOLIDWORKS 2015 x64 Edition.lnk" -ErrorAction SilentlyContinue
        }
		If (Test-Path "C:\Users\Public\Desktop\SOLIDWORKS 2015 x64 Edition.lnk" -PathType Any -ErrorAction SilentlyContinue) {
			Write-Log "Solidworks 2015 SP5.0 desktop shortcut INSTALLED for all users." 
		} Else {
			Write-Log "Solidworks 2015 SP5.0 desktop shortcut NOT INSTALLED for all users."
			Write-Log "Try to copy Solidworks 2015 SP5.0 desktop shortcut for all users."
			If (Test-Path "C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\sldworks.exe" -PathType Any -ErrorAction SilentlyContinue) {			
				Copy-Item "$dirFiles\shortcut\SOLIDWORKS 2015 x64 Edition.lnk" "C:\Users\Public\Desktop" -Force -ErrorAction SilentlyContinue
				Write-Log "SUCCESS copy Solidworks 2015 SP5.0 desktop shortcut for all users."
			} Else {
				Write-Log "FAILED copy Solidworks 2015 SP5.0 desktop shortcut for all users."
			}
		}
		If (Test-Path "C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\lang\english" -PathType Any -ErrorAction SilentlyContinue) {
			Write-Log "Solidworks 2015 SP5.0 INSTALLED, copying gtol.sys file." 
			Copy-Item "$dirFiles\update\Gtol.sym" "C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\lang\english" -Force -ErrorAction SilentlyContinue
		} Else {
			Write-Log "Solidworks 2015 SP5.0 NOT INSTALLED."
		}
        Get-ChildItem -Path "C:\Users" -Include "*" -Force -ErrorAction SilentlyContinue | ForEach-Object ($_) {
            $path0 = $_.FullName + "\Desktop"
            $path1 = $_.FullName + "\Desktop\SolidWorks Explorer 2011.lnk"
            $path2 = $_.FullName + "\Desktop\SolidWorks eDrawings 2011.lnk"
            $path3 = $_.FullName + "\Desktop\SolidWorks eDrawings 2011 x64 Edition.lnk"
            $path4 = $_.FullName + "\Desktop\SolidWorks eDrawings 2012.lnk"
            $path5 = $_.FullName + "\Desktop\SolidWorks eDrawings 2012 x64 Edition.lnk"
            $path6 = $_.FullName + "\Desktop\SolidWorks Explorer 2013.lnk"
            $path7 = $_.FullName + "\Desktop\SolidWorks eDrawings 2013.lnk"
            $path8 = $_.FullName + "\Desktop\SolidWorks eDrawings 2013 x64 Edition.lnk"
            $path9 = $_.FullName + "\Desktop\eDrawings 2014.lnk"
            $path10 = $_.FullName + "\Desktop\SolidWorks Explorer 2014.lnk"
            $path11 = $_.FullName + "\Desktop\SolidWorks eDrawings 2014.lnk"
            $path12 = $_.FullName + "\Desktop\SolidWorks eDrawings 2014 x64 Edition.lnk"
            $path13 = $_.FullName + "\Desktop\eDrawings 2015.lnk"
            $path14 = $_.FullName + "\Desktop\SolidWorks Explorer 2015.lnk"
            $path15 = $_.FullName + "\Desktop\SolidWorks eDrawings 2015.lnk"
            $path16 = $_.FullName + "\Desktop\SolidWorks eDrawings 2015 x64 Edition.lnk"
            $path17 = $_.FullName + "\Desktop\eDrawings 2015.lnk"
            $path18 = $_.FullName + "\Desktop\eDrawings 2015 x64 Edition.lnk"
            $path19 = $_.FullName + "\Desktop\SOLIDWORKS Composer 2015 x64 Edition.lnk"
            $path20 = $_.FullName + "\Desktop\SOLIDWORKS Composer Player 2015 - x64 Edition.lnk"
            $path21 = $_.FullName + "\Desktop\SOLIDWORKS 2015 x64 Edition.lnk"
            If (Test-Path $path0 -PathType Container -ErrorAction SilentlyContinue) {
                # Remove-Item -Force -Path $path1 -ErrorAction SilentlyContinue
                # Remove-Item -Force -Path $path2 -ErrorAction SilentlyContinue
                # Remove-Item -Force -Path $path3 -ErrorAction SilentlyContinue
                # Remove-Item -Force -Path $path4 -ErrorAction SilentlyContinue
                # Remove-Item -Force -Path $path5 -ErrorAction SilentlyContinue
                # Remove-Item -Force -Path $path6 -ErrorAction SilentlyContinue
                # Remove-Item -Force -Path $path7 -ErrorAction SilentlyContinue
                # Remove-Item -Force -Path $path8 -ErrorAction SilentlyContinue
                # Remove-Item -Force -Path $path9 -ErrorAction SilentlyContinue
                # Remove-Item -Force -Path $path10 -ErrorAction SilentlyContinue
                # Remove-Item -Force -Path $path11 -ErrorAction SilentlyContinue
                # Remove-Item -Force -Path $path12 -ErrorAction SilentlyContinue
                # Remove-Item -Force -Path $path13 -ErrorAction SilentlyContinue
                Remove-Item -Force -Path $path14 -ErrorAction SilentlyContinue
                # Remove-Item -Force -Path $path15 -ErrorAction SilentlyContinue
                # Remove-Item -Force -Path $path16 -ErrorAction SilentlyContinue
                # Remove-Item -Force -Path $path17 -ErrorAction SilentlyContinue
                Remove-Item -Force -Path $path18 -ErrorAction SilentlyContinue
                # Remove-Item -Force -Path $path19 -ErrorAction SilentlyContinue
                Remove-Item -Force -Path $path20 -ErrorAction SilentlyContinue
                #Remove-Item -Force -Path $path21 -ErrorAction SilentlyContinue
            }
        }
        
        ##*===============================================
		##* POST-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Installation'

		## <Perform Post-Installation tasks here>
        
        Write-Log "Solidworks 2015 SP5.0 INSTALL COMPLETED." 
        Unblock-AppExecution 
        
        # Display a message at the end of the install
		$installNewSWCount2 = (Get-InstalledApplication -Name "SOLIDWORKS 2015 x64 Edition SP05")        
        Write-Log "Solidworks 2015 SP5.0 check results: $installNewSWCount2"
        If (($installNewSWCount2 | Measure-Object).Count -gt 0) {
            # Display a message at the end of the install
            Unblock-AppExecution
            $explorerRunning2 = (Get-Process explorer -ea 0) 
            Write-Log "RUNNING explorer END results: $explorerRunning2"
            If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "Solidworks 2015 SP5.0 installation complete.  `r`n`r`nIn order to begin using the new software, please click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then reboot your machine in order to finish the install." -ButtonRightText "OK" -Icon Information -NoWait
				Start-Sleep -s 30
				Exit-Script -ExitCode "3010"
            } Else {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "Solidworks 2015 SP5.0 installation complete.  `r`n`r`nIn order to begin using the new software, please reboot your machine at your earliest convenience to complete the install." -ButtonRightText "OK" -Icon Information -NoWait
				Start-Sleep -s 30
				Exit-Script -ExitCode "3010"
            }
        } Else {
            # Display a message at the end of the install
            $explorerRunning2 = (Get-Process explorer -ea 0) 
            Write-Log "RUNNING explorer END results: $explorerRunning2"
            If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "Solidworks 2015 SP5.0 installation encounter an error.  `r`n`r`nPlease click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then reboot your machine in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait
				Start-Sleep -s 30
                Exit-Script -ExitCode "1"
            } Else {
                Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                Show-InstallationPrompt -Message "Solidworks 2015 SP5.0 installation encounter an error.  `r`n`r`nPlease reboot your machine at your earliest convenience in order to try the install again." -ButtonRightText "OK" -Icon Error -NoWait
				Start-Sleep -s 30
                Exit-Script -ExitCode "1"
            }
        }

        # End install script
        Write-Log "Solidworks 2015 SP5.0 INSTALL: END SCRIPT"

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
		Write-Log "Solidworks 2015 SP5.0 UNINSTALL: START SCRIPT"
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
        
   		# Perform uninstall SolidWorks 2015 SP5.0 check
        $installNewSWCount = (Get-InstalledApplication -Name "SOLIDWORKS 2015 x64 Edition SP05")
        Write-Log "SolidWorks 2015 SP5.0 check results: $installNewSWCount"
        If (($installNewSWCount | Measure-Object).Count -gt 0) {
            Write-Log "YES, SolidWorks 2015 SP5.0 is INSTALLED."
                       
            # Check explorer to see if running
            $explorerRunning1 = (Get-Process explorer -ea 0) 
            Write-Log "RUNNING explorer BEGIN results: $explorerRunning1"
            
            # Stop the services
            Stop-Service -name "WSearch" -Force -ErrorAction SilentlyContinue
            Stop-Service -name "ConisioWebServer" -Force -ErrorAction SilentlyContinue
            Stop-Service -name "ConisioDbServer" -Force -ErrorAction SilentlyContinue
            Stop-Service -name "ArchiveServerService" -Force -ErrorAction SilentlyContinue
            Stop-Service -name "WSearch" -Force -ErrorAction SilentlyContinue
            Start-Sleep -s 7

            # Show Welcome Message, close processes, allow up to 7 day deferral, and persist the prompt
            Show-InstallationWelcome -CloseApps "Search,dummySearch" -Silent -MinimizeWindows $false
            Show-InstallationWelcome -CloseApps "EdmServer,AddInRegSrv64,AddInSrv,CardEdit,ConisioAdmin,ConisioUrl,ConisioWebServer,DbUpdate,EdmServer,FileViewer,Inbox,InventorServer,NetRegSrv,Report,Search,SettingsDialog,TaskExecutor,ViewServer,ViewSetup,VLink,DsgnChkRptView,gabiswengine,LocalSldService,propertyManagerUpload,RTLibraryManager,setcatenv,sldbgproc,sldCostingTemplateEditorAppU,sldexitapp,sldphotoshopcon,sldProcMon,sldShellExtServer,sldu3d,SLDWORKS,sldworks_fs,swShellFileLauncher,swspmanager,UtlReportViewer,eDrawingOfficeAutomator,EModelViewer,AddInRegSrv32,sldtoolboxupdater,sldIM" -AllowDeferCloseApps -AllowDefer -DeferTimes 25 -DeferDays 7 -PersistPrompt -MinimizeWindows $false

            # Show Progress Message (with the default message) and Show-InstallationWelcome triggered
            Show-InstallationProgress -StatusMessage "UNINSTALLING SolidWorks 2015 SP5.0. `r`n`r`nThis may take some time. Please wait..."
            Start-Sleep -s 7
                            
            # Perform uninstallation tasks here, uninstalling SolidWorks 2015 SP5.0 23.5.0.81
            If (Test-Path "C:\WINDOWS\SolidWorks\IM_20150-40500-1100-100\sldim\sldIM.exe" -PathType Any -ErrorAction SilentlyContinue) {
                Execute-Process -FilePath "C:\WINDOWS\SolidWorks\IM_20150-40500-1100-100\sldim\sldIM.exe" -Arguments "/remove ""C:\WINDOWS\SolidWorks\IM_20150-40500-1100-100\sldim\sldIM_installed.xml""" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
                Start-Sleep -s 30
                Wait-Process -name sldIM -ErrorAction SilentlyContinue
                Start-Sleep -s 30
                Wait-Process -name sldIM -ErrorAction SilentlyContinue
            }
            If (Test-Path "C:\WINDOWS\SolidWorks\IM_20150-40500-1100-100 (1)\sldim\sldIM.exe" -PathType Any -ErrorAction SilentlyContinue) {
                Execute-Process -FilePath "C:\WINDOWS\SolidWorks\IM_20150-40500-1100-100 (1)\sldim\sldIM.exe" -Arguments "/remove ""C:\WINDOWS\SolidWorks\IM_20150-40500-1100-100 (1)\sldim\sldIM_installed.xml""" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
                Start-Sleep -s 30
                Wait-Process -name sldIM -ErrorAction SilentlyContinue
                Start-Sleep -s 30
                Wait-Process -name sldIM -ErrorAction SilentlyContinue
            }
            If (Test-Path "C:\WINDOWS\SolidWorks\IM_20150-40500-1100-100 (2)\sldim\sldIM.exe" -PathType Any -ErrorAction SilentlyContinue) {
                Execute-Process -FilePath "C:\WINDOWS\SolidWorks\IM_20150-40500-1100-100 (2)\sldim\sldIM.exe" -Arguments "/remove ""C:\WINDOWS\SolidWorks\IM_20150-40500-1100-100 (2)\sldim\sldIM_installed.xml""" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
                Start-Sleep -s 30
                Wait-Process -name sldIM -ErrorAction SilentlyContinue
                Start-Sleep -s 30
                Wait-Process -name sldIM -ErrorAction SilentlyContinue
            }
            If (Test-Path "C:\WINDOWS\SolidWorks\IM_20150-40500-1100-100 (3)\sldim\sldIM.exe" -PathType Any -ErrorAction SilentlyContinue) {
                Execute-Process -FilePath "C:\WINDOWS\SolidWorks\IM_20150-40500-1100-100 (3)\sldim\sldIM.exe" -Arguments "/remove ""C:\WINDOWS\SolidWorks\IM_20150-40500-1100-100 (3)\sldim\sldIM_installed.xml""" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
                Start-Sleep -s 30
                Wait-Process -name sldIM -ErrorAction SilentlyContinue
                Start-Sleep -s 30
                Wait-Process -name sldIM -ErrorAction SilentlyContinue
            }
            If (Test-Path "C:\WINDOWS\SolidWorks\IM_20150-40500-1100-100 (4)\sldim\sldIM.exe" -PathType Any -ErrorAction SilentlyContinue) {
                Execute-Process -FilePath "C:\WINDOWS\SolidWorks\IM_20150-40500-1100-100 (4)\sldim\sldIM.exe" -Arguments "/remove ""C:\WINDOWS\SolidWorks\IM_20150-40500-1100-100 (4)\sldim\sldIM_installed.xml""" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
                Start-Sleep -s 30
                Wait-Process -name sldIM -ErrorAction SilentlyContinue
                Start-Sleep -s 30
                Wait-Process -name sldIM -ErrorAction SilentlyContinue
            }
            If (Test-Path "C:\WINDOWS\SolidWorks\IM_20150-40500-1100-100 (5)\sldim\sldIM.exe" -PathType Any -ErrorAction SilentlyContinue) {
                Execute-Process -FilePath "C:\WINDOWS\SolidWorks\IM_20150-40500-1100-100 (5)\sldim\sldIM.exe" -Arguments "/remove ""C:\WINDOWS\SolidWorks\IM_20150-40500-1100-100 (5)\sldim\sldIM_installed.xml""" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
                Start-Sleep -s 30
                Wait-Process -name sldIM -ErrorAction SilentlyContinue
                Start-Sleep -s 30
                Wait-Process -name sldIM -ErrorAction SilentlyContinue
            }
            If (Test-Path "C:\WINDOWS\SolidWorks\IM_20150-40500-1100-100 (6)\sldim\sldIM.exe" -PathType Any -ErrorAction SilentlyContinue) {
                Execute-Process -FilePath "C:\WINDOWS\SolidWorks\IM_20150-40500-1100-100 (6)\sldim\sldIM.exe" -Arguments "/remove ""C:\WINDOWS\SolidWorks\IM_20150-40500-1100-100 (6)\sldim\sldIM_installed.xml""" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
                Start-Sleep -s 30
                Wait-Process -name sldIM -ErrorAction SilentlyContinue
                Start-Sleep -s 30
                Wait-Process -name sldIM -ErrorAction SilentlyContinue
            }
            If (Test-Path "C:\WINDOWS\SolidWorks\IM_20150-40500-1100-100 (7)\sldim\sldIM.exe" -PathType Any -ErrorAction SilentlyContinue) {
                Execute-Process -FilePath "C:\WINDOWS\SolidWorks\IM_20150-40500-1100-100 (7)\sldim\sldIM.exe" -Arguments "/remove ""C:\WINDOWS\SolidWorks\IM_20150-40500-1100-100 (7)\sldim\sldIM_installed.xml""" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
                Start-Sleep -s 30
                Wait-Process -name sldIM -ErrorAction SilentlyContinue
                Start-Sleep -s 30
                Wait-Process -name sldIM -ErrorAction SilentlyContinue
            }
            If (Test-Path "C:\WINDOWS\SolidWorks\IM_20150-40500-1100-100 (8)\sldim\sldIM.exe" -PathType Any -ErrorAction SilentlyContinue) {
                Execute-Process -FilePath "C:\WINDOWS\SolidWorks\IM_20150-40500-1100-100 (8)\sldim\sldIM.exe" -Arguments "/remove ""C:\WINDOWS\SolidWorks\IM_20150-40500-1100-100 (8)\sldim\sldIM_installed.xml""" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
                Start-Sleep -s 30
                Wait-Process -name sldIM -ErrorAction SilentlyContinue
                Start-Sleep -s 30
                Wait-Process -name sldIM -ErrorAction SilentlyContinue
            }
            If (Test-Path "C:\WINDOWS\SolidWorks\IM_20150-40500-1100-100 (9)\sldim\sldIM.exe" -PathType Any -ErrorAction SilentlyContinue) {
                Execute-Process -FilePath "C:\WINDOWS\SolidWorks\IM_20150-40500-1100-100 (9)\sldim\sldIM.exe" -Arguments "/remove ""C:\WINDOWS\SolidWorks\IM_20150-40500-1100-100 (9)\sldim\sldIM_installed.xml""" -Windowstyle Hidden -IgnoreExitCodes "3010" -ContinueOnError $true
                Start-Sleep -s 30
                Wait-Process -name sldIM -ErrorAction SilentlyContinue
                Start-Sleep -s 30
                Wait-Process -name sldIM -ErrorAction SilentlyContinue
            }         

            # Perform uninstallation tasks here, uninstalling SolidWorks
			Execute-MSI -Action Uninstall -Path "{F8093877-4F2C-40ED-9BA7-2F9F48F5176F}" -ContinueOnError $true

			# Perform uninstallation tasks here, uninstalling SolidWorks Flow Simulation
			Execute-MSI -Action Uninstall -Path "{4A7898B4-068C-45DE-9994-CFC347C87182}" -ContinueOnError $true

			# Perform uninstallation tasks here, uninstalling SolidWorks Composer Player
			Execute-MSI -Action Uninstall -Path "{35CFB7E6-939A-4A06-A4FC-AA9CBB9B1E0F}" -ContinueOnError $true

			# Perform uninstallation tasks here, uninstalling SolidWorks Plastics
			Execute-MSI -Action Uninstall -Path "{25AF0A62-A60A-4112-BD59-857D600B3B0F}" -ContinueOnError $true
			
			# # Check to see if the folder exists, if so delete folder
            # If (Test-Path "C:\MCAD\SWDATA" -PathType Container -ErrorAction SilentlyContinue) {
                # If((Get-ChildItem "C:\MCAD\SWDATA" -Force | Select-Object -First 1 | Measure-Object).Count -eq 0) {
                   # Remove-Folder -Path "C:\MCAD\SWDATA" -ContinueOnError $true
                # }
            # }
            # If (Test-Path "C:\MCAD\SW2015" -PathType Container -ErrorAction SilentlyContinue) {
                # If((Get-ChildItem "C:\MCAD\SW2015" -Force | Select-Object -First 1 | Measure-Object).Count -eq 0) {
                   # Remove-Folder -Path "C:\MCAD\SW2015" -ContinueOnError $true
                # }
            # }
            # If (Test-Path "C:\MCAD" -PathType Container -ErrorAction SilentlyContinue) {
                # If((Get-ChildItem "C:\MCAD\" -Force | Select-Object -First 1 | Measure-Object).Count -eq 0) {
                   # Remove-Folder -Path "C:\MCAD" -ContinueOnError $true
                # }
            # }
            
            If (Test-Path "C:\Users\Public\Desktop" -PathType Container -ErrorAction SilentlyContinue) {
                # Remove-Item -Force -Path "C:\Users\Public\Desktop\eDrawings 2015.lnk" -ErrorAction SilentlyContinue
                Remove-Item -Force -Path "C:\Users\Public\Desktop\SolidWorks Explorer 2015.lnk" -ErrorAction SilentlyContinue
                # Remove-Item -Force -Path "C:\Users\Public\Desktop\SolidWorks eDrawings 2015.lnk" -ErrorAction SilentlyContinue
                # Remove-Item -Force -Path "C:\Users\Public\Desktop\SolidWorks eDrawings 2015 x64 Edition.lnk" -ErrorAction SilentlyContinue
                # Remove-Item -Force -Path "C:\Users\Public\Desktop\eDrawings 2015.lnk" -ErrorAction SilentlyContinue
                Remove-Item -Force -Path "C:\Users\Public\Desktop\eDrawings 2015 x64 Edition.lnk" -ErrorAction SilentlyContinue
                # Remove-Item -Force -Path "C:\Users\Public\Desktop\Explorer 2015.lnk" -ErrorAction SilentlyContinue
                # Remove-Item -Force -Path "C:\Users\Public\Desktop\SOLIDWORKS Composer 2015 x64 Edition.lnk" -ErrorAction SilentlyContinue
                Remove-Item -Force -Path "C:\Users\Public\Desktop\SOLIDWORKS Composer Player 2015 - x64 Edition.lnk" -ErrorAction SilentlyContinue
                Remove-Item -Force -Path "C:\Users\Public\Desktop\SOLIDWORKS 2015 x64 Edition.lnk" -ErrorAction SilentlyContinue
            }
            Get-ChildItem -Path "C:\Users" -Include "*" -Force -ErrorAction SilentlyContinue | ForEach-Object ($_) {
                $path0 = $_.FullName + "\Desktop"
                $path13 = $_.FullName + "\Desktop\eDrawings 2015.lnk"
                $path14 = $_.FullName + "\Desktop\SolidWorks Explorer 2015.lnk"
                $path15 = $_.FullName + "\Desktop\SolidWorks eDrawings 2015.lnk"
                $path16 = $_.FullName + "\Desktop\SolidWorks eDrawings 2015 x64 Edition.lnk"
                $path17 = $_.FullName + "\Desktop\eDrawings 2015.lnk"
                $path18 = $_.FullName + "\Desktop\eDrawings 2015 x64 Edition.lnk"
                $path19 = $_.FullName + "\Desktop\Explorer 2015.lnk"
                $path20 = $_.FullName + "\Desktop\SOLIDWORKS Composer 2015 x64 Edition.lnk"
                $path21 = $_.FullName + "\Desktop\SOLIDWORKS Composer Player 2015 - x64 Edition.lnk"
                $path22 = $_.FullName + "\Desktop\SOLIDWORKS 2015 x64 Edition.lnk"
                If (Test-Path $path0 -PathType Container -ErrorAction SilentlyContinue) {
                    # Remove-Item -Force -Path $path13 -ErrorAction SilentlyContinue
                    Remove-Item -Force -Path $path14 -ErrorAction SilentlyContinue
                    # Remove-Item -Force -Path $path15 -ErrorAction SilentlyContinue
                    # Remove-Item -Force -Path $path16 -ErrorAction SilentlyContinue
                    # Remove-Item -Force -Path $path17 -ErrorAction SilentlyContinue
                    Remove-Item -Force -Path $path18 -ErrorAction SilentlyContinue
                    # Remove-Item -Force -Path $path19 -ErrorAction SilentlyContinue
                    # Remove-Item -Force -Path $path20 -ErrorAction SilentlyContinue
                    Remove-Item -Force -Path $path21 -ErrorAction SilentlyContinue
                    Remove-Item -Force -Path $path22 -ErrorAction SilentlyContinue
                }
            }

            Write-Log "Solidworks 2015 SP5.0 UNINSTALL COMPLETED." 
            # Display a message at the end of the uninstall
            Unblock-AppExecution
            $explorerRunning2 = (Get-Process explorer -ea 0) 
            Write-Log "RUNNING explorer END results: $explorerRunning2"
			$installNewSWCount2 = (Get-InstalledApplication -Name "SOLIDWORKS 2015 x64 Edition SP05")
            Write-Log "Solidworks 2015 SP5.0 check results: $installNewSWCount2"
            If (($installNewSWCount2 | Measure-Object).Count -gt 0) {
                If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "Solidworks 2015 SP5.0 uninstallation encounter an error.  `r`n`r`nPlease click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then reboot your machine in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait
					Start-Sleep -s 30
                    Exit-Script -ExitCode "1"
                } Else {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "Solidworks 2015 SP5.0 uninstallation encounter an error.  `r`n`r`nPlease reboot your machine at your earliest convenience in order to try the uninstall again." -ButtonRightText "OK" -Icon Error -NoWait
					Start-Sleep -s 30
                    Exit-Script -ExitCode "1"
                }
            } Else {
                If ((($explorerRunning1 | Measure-Object).Count -gt 0) -and (($explorerRunning2 | Measure-Object).Count -le 0)) {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "Solidworks 2015 SP5.0 uninstallation complete.  `r`n`r`nPlease click OK, then press CTRL+ATL+DELETE, then click LOG OFF, and then reboot your machine in order to finish the uninstall." -ButtonRightText "OK" -Icon Information -NoWait
					Start-Sleep -s 30
                    Exit-Script -ExitCode "3010"
                } Else {
                    Write-Log "$installName $deploymentTypeName completed with exit code [$mainExitCode]."
                    Show-InstallationPrompt -Message "Solidworks 2015 SP5.0 uninstallation complete.  `r`n`r`nPlease reboot your machine at your earliest convenience to complete the uninstall." -ButtonRightText "OK" -Icon Information -NoWait
					Start-Sleep -s 30
                    Exit-Script -ExitCode "3010"
                }
            }
        } Else {
            Write-Log "NO, Solidworks 2015 SP5.0 is NOT installed."
            Unblock-AppExecution
            Show-InstallationPrompt -Message "Solidworks 2015 SP5.0 NOT installed, nothing to uninstall. NO changes made. `r`n`r`nClick OK." -ButtonRightText "OK" -Icon Information -NoWait
			Start-Sleep -s 10
			Exit-Script -ExitCode "0"
        }
        
        # End uninstall script
        Write-Log "Solidworks 2015 SP5.0 UNINSTALL: END SCRIPT"
		
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