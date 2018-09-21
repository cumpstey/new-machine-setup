<#
.SYNOPSIS
Script to do a lot of the grunt work in setting up a new Windows developer machine.

.DESCRIPTION
The script:
- Removes pointless Windows 10 built-in apps
- Uses Chocolatey to install a lot of software
- Installs IIS
- Installs some front-end developer tools
- Sets up Pageant to run when the computer starts up

.PARAMETER removeWindowsApps
Remove pointless Windows 10 built-in apps.

.PARAMETER software
Install Chocolatey, and default set of useful apps.

.PARAMETER devSoftware
Install Chocolatey, and developer-specific apps.

.PARAMETER iis
Install IIS, including some optional components.

.PARAMETER startup
Set apps to run on startup.

.PARAMETER frontend
Install tools for front end development work.

.PARAMETER uninstallable
Print a list of useful software which should be installed by different means.

.NOTES
Author: Neil Cumpstey
Date: Sep 2018
#>
# Requires -RunAsAdministrator

[CmdletBinding()]
Param(
  [switch]$removeWindowsApps,
  [switch]$software,
  [switch]$devSoftware,
  [switch]$iis,
  [switch]$startup,
  [switch]$frontend,
  [switch]$uninstallable
)

$ErrorActionPreference = 'Stop'
$WarningPreference = 'Continue'
$InformationPreference = 'Continue'
$DebugPreference = if ($PSCmdlet.MyInvocation.BoundParameters["Debug"].IsPresent) { 'Continue' } else { 'SilentlyContinue' }

<#
.SYNOPSIS
Checks whether the supplied pathVar string contains the supplied item string,
when considered in semicolon-separated PATH format, and disregarding any trailing slash.
#>
function PathContains($pathVar, $item) {
  $pathVar -match ('(^|;)' + [Regex]::Escape($item.trimend('\')) + '\\?($|;)')
}

<#
.SYNOPSIS
Updates the system PATH variable with the supplied path, if it's not already specified.
#>
function UpdatePath($item)
{
  Push-Location
  Set-Location "HKLM:\System\CurrentControlSet\Control\Session Manager\Environment"

  try {
    $pathVar = (Get-ItemProperty . -Name Path).Path

    if (PathContains $pathVar $item)
    {
      Write-Debug "The path $pathToCheck is not included in the system path, adding..."
      $pathVar += ';' + $item
      Set-ItemProperty . Path $pathVar
      if ($?)
      {
        Write-Information "Added $item to PATH"
      }
      else
      {
        Write-Warning "Could not set PATH when adding $item. Run as administrator, or do this manually."
      }
    }
    else
    {
      Write-Information "$item already exists in PATH"
    }
  } finally {
    Pop-Location
  }
}

<#
.SYNOPSIS
Ensures Chocolatey is installed, by checking for the `choco` command
and installing Chocolatey if the command is not recognised.
#>
function EnsureChocolateyIsInstalled() {
  # Install chocolatey
  if (!(Get-Command choco -errorAction SilentlyContinue))
  {
    iex ((new-object net.webclient).DownloadString('https://chocolatey.org/install.ps1'))
    refreshenv
  }

  choco install chocolateygui -y
}

<#
.SYNOPSIS
Removes a whole lot of useless preinstalled Windows crap.
Windows may reinstall any or all of this during an update process.
#>
function RemoveWindowsApps() {
  Get-AppxPackage *3dbuilder* | Remove-AppxPackage
  Get-AppxPackage *windowsalarms* | Remove-AppxPackage
  Get-AppxPackage *windowscommunicationsapps* | Remove-AppxPackage
  Get-AppxPackage *windowscamera* | Remove-AppxPackage
  Get-AppxPackage *officehub* | Remove-AppxPackage
  Get-AppxPackage *skypeapp* | Remove-AppxPackage
  Get-AppxPackage *getstarted* | Remove-AppxPackage
  Get-AppxPackage *zunemusic* | Remove-AppxPackage
  Get-AppxPackage *windowsmaps* | Remove-AppxPackage
  Get-AppxPackage *solitairecollection* | Remove-AppxPackage
  Get-AppxPackage *bingfinance* | Remove-AppxPackage
  Get-AppxPackage *zunevideo* | Remove-AppxPackage
  Get-AppxPackage *bingnews* | Remove-AppxPackage
  Get-AppxPackage *windowsphone* | Remove-AppxPackage
  Get-AppxPackage *photos* | Remove-AppxPackage
  Get-AppxPackage *bingsports* | Remove-AppxPackage
  Get-AppxPackage *soundrecorder* | Remove-AppxPackage
  Get-AppxPackage *bingweather* | Remove-AppxPackage
  Get-AppxPackage *xboxapp* | Remove-AppxPackage
}

<#
.SYNOPSIS
Uses Chocolatey to install useful software:
- 7-Zip: 7-Zip is a file archiver with a high compression ratio.
- Carbon: Carbon is a PowerShell module for automating the installation and configuration of Windows applications, websites, and services.
- FileZilla: FileZilla is a cross-platform FTP, SFTP, and FTPS client with a vast list of features.
- Firefox: Mozilla Firefox is a free and open-source web browser.
- f.lux: f.lux makes the color of your computer's display adapt to the time of day, warm at night and like sunlight during the day.
- Google Chrome: Chrome is a fast, secure and free browser for all your devices.
- Greenshot: Greenshot is a free screenshot tool optimized for productivity.
- grepWin: grepWin is a simple search and replace tool which can use regular expressions to do its job.
- IrfanView: IrfanView is a very fast, small, compact and innovative graphic viewer.
- Java Runtime Environment: The JRE is the runtime environment for Java programs.
- Notepad++: Notepad++ is a free source code editor which supports several programming languages.
- Remote Desktop Connection Manager: RDCMan manages multiple remote desktop connections.
- WinDirStat: WinDirStat is a disk usage statistics viewer and cleanup tool.
- WinSCP: WinSCP is an open source free SFTP client, SCP client, FTPS client and FTP client.
#>
function InstallSoftware() {
  EnsureChocolateyIsInstalled

  # Use Chocolatey to install software
  choco install 7zip -y
  choco install carbon -y
  choco install filezilla -y
  choco install firefox -y
  choco install f.lux -y
  choco install googlechrome -y
  choco install greenshot -y
  choco install grepwin -y
  choco install irfanview -y
  choco install jre8 -y
  choco install notepadplusplus -y
  choco install rdcman -y
  choco install windirstat -y
  choco install winscp -y
}

<#
.SYNOPSIS
Uses Chocolatey to install useful software:
- AWS Command Line Interface: The AWS CLI is an open source tool built on top of the AWS SDK for Python (Boto) that provides commands for interacting with AWS services.
- DBeaver: DBeaver is free and open source universal database tool for developers and database administrators.
- Fiddler: Fiddler is a free web debugging tool which logs all http(s) traffic between your computer and the internet.
- Git Extensions: Git Extensions is a graphical user interface for Git that allows you to control Git without using the commandline.
- KDiff3: KDiff3 is a graphical text difference analyzer for up to 3 input files, provides character-by-character analysis and a text merge tool with integrated editor.
- Java Development Kit: The JDK is a software development environment used for developing Java applications and applets.
- NSSM: Non-Sucking Service Manager is a service helper which doesn't suck.
- Nuget: NuGet is the package manager for the Microsoft development platforms including .NET.
- Nuget Package Explorer: Create, update and deploy Nuget Packages with a GUI.
- Node Version Manager: Manage multiple installations of Node.js on a Windows computer.
- Papercut: Papercut is a simplified SMTP server designed to only receive messages (not to send them on) with a GUI on top of it allowing you to see the messages it receives.
- PuTTY: PuTTY is a free implementation of Telnet and SSH for Windows and Unix platforms, along with an xterm terminal emulator.
- Ruby: Ruby is a dynamic, open source programming language with a focus on simplicity and productivity.
- Slack: Slack is a platform for team communication.
- Visual Studio Code: Visual Studio Code is a code editor redefined and optimized for building and debugging modern web and cloud applications.
- Web PI: The Microsoft Web Platform Installer is a free tool that makes getting the latest components of the Microsoft Web Platform easy

Also installs:
- Node.js (latest version, using NVM): Node.js is a JavaScript runtime built on Chrome's V8 JavaScript engine.
#>
function InstallDevSoftware() {
  EnsureChocolateyIsInstalled

  # Use Chocolatey to install software
  choco install awscli -y
  choco install dbeaver -y
  choco install fiddler -y
  choco install gitextensions -y
  choco install kdiff3 -y
  choco install jdk8 -y
  choco install nssm -y
  choco install nuget.commandline -y
  choco install nugetpackageexplorer -y
  choco install nvm -y
  choco install papercut -y
  choco install putty -y
  choco install python2 -y --params "/InstallDir:C:\tools\python2"
  choco install python3 -y --params "/InstallDir:C:\tools\python3"
  choco install ruby -y
  choco install slack -y
  choco install vscode -y
  choco install webpi -y

  # Install any extensions, which should be installed after the main packages above
  choco install vscode-docker -y

  # Add git install directory to the Path, so git can be used from commandline
  UpdatePath "C:\Program Files\Git\cmd"

  # Load the new Path value
  refreshenv

  # Install any other software
  nvm install latest
  Write-Warning "The latest version of NodeJS has been installed using NVM, but you should run 'nvm use <version>' in order to enable it."
}

<#
.SYNOPSIS
Installs:
- IIS, including ASP.NET and management console
- IIS components (using Web PI):
  - ASP.NET MVC 3
  - Url Rewrite 2
  - Application Request Routing
#>
function InstallIis() {
  Set-ExecutionPolicy Bypass -Scope Process

  # Install IIS
  Enable-WindowsOptionalFeature -Online -FeatureName IIS-WebServerRole
  Enable-WindowsOptionalFeature -Online -FeatureName IIS-WebServer
  Enable-WindowsOptionalFeature -Online -FeatureName IIS-CommonHttpFeatures
  Enable-WindowsOptionalFeature -Online -FeatureName IIS-HttpErrors
  Enable-WindowsOptionalFeature -Online -FeatureName IIS-HttpRedirect
  Enable-WindowsOptionalFeature -Online -FeatureName IIS-ApplicationDevelopment
  Enable-WindowsOptionalFeature -Online -All -FeatureName IIS-NetFxExtensibility
  Enable-WindowsOptionalFeature -Online -FeatureName IIS-HealthAndDiagnostics
  Enable-WindowsOptionalFeature -Online -FeatureName IIS-HttpLogging
  Enable-WindowsOptionalFeature -Online -FeatureName IIS-LoggingLibraries
  Enable-WindowsOptionalFeature -Online -FeatureName IIS-RequestMonitor
  Enable-WindowsOptionalFeature -Online -FeatureName IIS-HttpTracing
  Enable-WindowsOptionalFeature -Online -FeatureName IIS-Security
  Enable-WindowsOptionalFeature -Online -FeatureName IIS-RequestFiltering
  Enable-WindowsOptionalFeature -Online -FeatureName IIS-Performance
  Enable-WindowsOptionalFeature -Online -FeatureName IIS-WebServerManagementTools
  Enable-WindowsOptionalFeature -Online -FeatureName IIS-IIS6ManagementCompatibility
  Enable-WindowsOptionalFeature -Online -FeatureName IIS-Metabase
  Enable-WindowsOptionalFeature -Online -FeatureName IIS-ManagementConsole
  Enable-WindowsOptionalFeature -Online -FeatureName IIS-BasicAuthentication
  Enable-WindowsOptionalFeature -Online -FeatureName IIS-WindowsAuthentication
  Enable-WindowsOptionalFeature -Online -FeatureName IIS-StaticContent
  Enable-WindowsOptionalFeature -Online -FeatureName IIS-DefaultDocument
  Enable-WindowsOptionalFeature -Online -FeatureName IIS-WebSockets
  Enable-WindowsOptionalFeature -Online -FeatureName IIS-ApplicationInit
  Enable-WindowsOptionalFeature -Online -FeatureName IIS-NetFxExtensibility45
  Enable-WindowsOptionalFeature -Online -All -FeatureName IIS-ASPNET45
  Enable-WindowsOptionalFeature -Online -FeatureName IIS-ISAPIExtensions
  Enable-WindowsOptionalFeature -Online -FeatureName IIS-ISAPIFilter
  Enable-WindowsOptionalFeature -Online -FeatureName IIS-HttpCompressionStatic

  # Install additional IIS components using Web PI
  $webpi = "C:\Program Files\Microsoft\Web Platform Installer\WebpiCmd.exe"
  if (Get-Command $webpi -ErrorAction SilentlyContinue)
  {
    & $webpi /Install /Products:'MVC3,UrlRewrite2,Arrv3_0' /AcceptEULA
  }
  else
  {
    Write-Warning "Web Platform Installer is not installed. Cannot install additional IIS components."
  }
}

<#
.SYNOPSIS
Creates shortcuts to items which should be run at startup:
- Pageant, with appropriate keys
#>
function CreateStartupSettings() {
  # Create a Windows PowerShell Com Object
  $wshShell = New-Object -ComObject WScript.Shell

  # Use the Com Object to create the Pageant shortcut
  $pageant = $wshShell.CreateShortcut("$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\Pageant.lnk")
  $pageant.TargetPath = "C:\ProgramData\chocolatey\bin\Pageant.exe"
  $pageant.IconLocation = "C:\ProgramData\chocolatey\bin\Pageant.exe"
  $pageant.Arguments = """$env:USERPROFILE\.ssh\id_rsa.ppk"" ""$env:USERPROFILE\.ssh\bitbucket.ppk"" ""$env:USERPROFILE\.ssh\zone_bitbucket.ppk"""
  $pageant.Save()
}

<#
.SYNOPSIS
Installs the following:
- Grunt
- Sass
- Compass
#>
function InstallFrontEndTools() {
  if (!(Get-Command node -errorAction SilentlyContinue) -or !(Get-Command npm -errorAction SilentlyContinue))
  {
    Write-Warning "NodeJS and NPM must be installed in order for the following to be installed:
    - Grunt"
    return
  } else {
    npm install -g grunt-cli
  }

  if (!(Get-Command ruby -errorAction SilentlyContinue) -or !(Get-Command gem -errorAction SilentlyContinue))
  {
    Write-Warning "Ruby and Gem must be installed in order for the following to be installed:
    - Sass
    - Compass"
    return
  } else {
    gem install sass
    gem install compass
  }
}

<#
.SYNOPSIS
Prints a list of software which cannot be installed automatically, and should be installed by other means.
#>
function ListManualSteps() {
  Write-Warning "The following software should be installed manually:
  - BluOS
  - Microsoft Office
  - Visual Studio Enterprise
  - SQL Server
  - SQL Server Management Studio
  - Visual Studio Code extensions:
    - .ejs
    - Laravel Blade Snippets
    - Rainbow CSV
    - XML Tools
  - AWS tools
  - Egnyte Connect
  - Starleaf
  - Watchguard VPN Client
  
  Also:
  - Install Zone fonts (Z:\Shared\Zone\2018 Brand Resources\01 Font - Install first!\ZoneEuclid\CFF)
  - Set up Zone, EPiServer and Sitecore Nuget feeds in Visual Studio"
}

if ($removeWindowsApps) {
  RemoveWindowsApps
}

if ($software) {
  InstallSoftware
}

if ($devSoftware) {
  InstallDevSoftware
}

if ($iis) {
  InstallIis
}

if ($startup) {
  CreateStartupSettings
}

if ($frontend) {
  InstallFrontEndTools
}

if ($uninstallable) {
  WarnUninstallableSoftware
}
