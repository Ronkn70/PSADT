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
	[string]$appVendor = 'TriNet HR Corp:'
	[string]$appName = 'Java Repair Tool'
	[string]$appVersion = ''
	[string]$appArch = ''
	[string]$appLang = 'EN'
	[string]$appRevision = '01'
	[string]$appScriptVersion = '1.0.0'
	[string]$appScriptDate = '02/03/2017'
	[string]$appScriptAuthor = 'Ron Knight'
	##*===============================================
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
        Write-Log "Java Repair Tool INSTALL: START SCRIPT"
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
        		
		## Show Welcome Message, close Internet Explorer if required, allow up to 3 deferrals, verify there is enough disk space to complete the install, and persist the prompt
				Show-InstallationWelcome -CloseApps 'iexplore,firefox,chrome,safari,opera' -CheckDiskSpace -PersistPrompt

		## Show Progress Message (with the default message)
		Show-InstallationProgress
		Show-InstallationPrompt -Message 'This Application will uninstall/reinstall Java. Please save all your work and click Continue'-PersistPrompt -Icon Warning -ButtonMiddleText 'Continue'

		## <Perform Pre-Installation tasks here>
        Write-Log "Beginning Uninstall of Java"
		    $UninstallJava6andBelow = $true
 
       #Declare version arrays
    $32bitJava = @()
    $64bitJava = @()
    $32bitVersions = @()
    $64bitVersions = @()
 
        #Perform WMI query to find installed Java Updates
        if ($UninstallJava6andBelow) {
             $32bitJava += Get-WmiObject -Class Win32_Product | Where-Object { 
              $_.Name -match "(?i)Java(\(TM\))*\s\d+(\sUpdate\s\d+)*$"
         }
       #Also find Java version 5, but handled slightly different as CPU bit is only distinguishable by the GUID
        $32bitJava += Get-WmiObject -Class Win32_Product | Where-Object { 
         ($_.Name -match "(?i)J2SE\sRuntime\sEnvironment\s\d[.]\d(\sUpdate\s\d+)*$") -and ($_.IdentifyingNumber -match "^\{32")
       }
    } else {
        $32bitJava += Get-WmiObject -Class Win32_Product | Where-Object { 
             $_.Name -match "(?i)Java((\(TM\) 7)|(\s\d+))(\sUpdate\s\d+)*$"
        }
    }
 
    #Perform WMI query to find installed Java Updates (64-bit)
        if ($UninstallJava6andBelow) {
            $64bitJava += Get-WmiObject -Class Win32_Product | Where-Object { 
               $_.Name -match "(?i)Java(\(TM\))*\s\d+(\sUpdate\s\d+)*\s[(]64-bit[)]$"
            }
            #Also find Java version 5, but handled slightly different as CPU bit is only distinguishable by the GUID
            $64bitJava += Get-WmiObject -Class Win32_Product | Where-Object { 
                ($_.Name -match "(?i)J2SE\sRuntime\sEnvironment\s\d[.]\d(\sUpdate\s\d+)*$") -and ($_.IdentifyingNumber -match "^\{64")
            }
        } else {
           $64bitJava += Get-WmiObject -Class Win32_Product | Where-Object { 
                $_.Name -match "(?i)Java((\(TM\) 7)|(\s\d+))(\sUpdate\s\d+)*\s[(]64-bit[)]$"
            }
        }
 
        #Enumerate and populate array of versions
        Foreach ($app in $32bitJava) {
           if ($app -ne $null) { $32bitVersions += $app.Version }
        }
 
        #Enumerate and populate array of versions
        Foreach ($app in $64bitJava) {
           if ($app -ne $null) { $64bitVersions += $app.Version }
        }
 
        #Create an array that is sorted correctly by the actual Version (as a System.Version object) rather than by value.
        $sorted32bitVersions = $32bitVersions | %{ New-Object System.Version ($_) } | sort
        $sorted64bitVersions = $64bitVersions | %{ New-Object System.Version ($_) } | sort
        #If a single result is returned, convert the result into a single value array so we don't run in to trouble calling .GetUpperBound later
        if($sorted32bitVersions -isnot [system.array]) { $sorted32bitVersions = @($sorted32bitVersions)}
        if($sorted64bitVersions -isnot [system.array]) { $sorted64bitVersions = @($sorted64bitVersions)}
        #Grab the value of the newest version from the array, first converting 
        $newest32bitVersion = $sorted32bitVersions[$sorted32bitVersions.GetUpperBound(0)]
        $newest64bitVersion = $sorted64bitVersions[$sorted64bitVersions.GetUpperBound(0)]
 
        Foreach ($app in $32bitJava) {
            if ($app -ne $null)
            {
           # Remove all versions of Java, where the version does not match the newest version.
            #if (($app.Version -ne $newest32bitVersion) -and ($newest32bitVersion -ne $null)) {
            $appGUID = $app.Properties["IdentifyingNumber"].Value.ToString()
            Start-Process -FilePath "msiexec.exe" -ArgumentList "/qn /norestart /x $($appGUID)" -Wait -Passthru
            #write-host "Uninstalling 32-bit version: " $app
             #}
              }
        }
 
        Foreach ($app in $64bitJava) {
           if ($app -ne $null)
           {
            # Remove all versions of Java, where the version does not match the newest version.
            #if (($app.Version -ne $newest64bitVersion) -and ($newest64bitVersion -ne $null)) {
               $appGUID = $app.Properties["IdentifyingNumber"].Value.ToString()
                   Start-Process -FilePath "msiexec.exe" -ArgumentList "/qn /norestart /x $($appGUID)" -Wait -Passthru
                   #write-host "Uninstalling 64-bit version: " $app
                #}
            }
        }
		
		##*===============================================
		##* INSTALLATION 
		##*===============================================
		[string]$installPhase = 'Installation'
        Write-Log "Installing New Java version"
		
		## Handle Zero-Config MSI Installations

		
		## <Perform Installation tasks here>
        Write-Log "Installing Java"
        Show-InstallationProgress "All previous version of Java have been removed. Reinstalling current Version of Java. Please wait..."
	    #Execute-MSI -Action Install -Path "$dirFiles\jre.msi" -Parameter "/q /norestart AUTOUPDATECHECK=0 JAVAUPDATE=0 JU=0 STATIC=0"
		Execute-MSI -Action 'Install' -Path "$dirFiles\jre.msi" -Parameters "STATIC=0 AUTUPDATECHECK=0 ALLUSERS=1 DWUSINTERVAL=120 JAVAUPDATE=0 AUTO_UPDATE=0 WEB_JAVA=1 WEB_JAVA_SECURITY_LEVEL=H WEB_ANALYTICS=0 EULA=0 REBOOT=0 SPONSORS=0 /qn"
        		
		
		##*===============================================
		##* POST-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Installation'
		
		## <Perform Post-Installation tasks here>
		
		## Display a message at the end of the install
		If (-not $useDefaultMsi) { Show-InstallationPrompt -Message 'Java Has been Replaced' -ButtonRightText 'OK' -Icon Information -NoWait }
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