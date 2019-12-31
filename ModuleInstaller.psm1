<#
# Module Installer - Created by -Kai (https://social.technet.microsoft.com/profile/-kai/)
# Created 20th November, 2019
#
# Module/Functions is used for installing Skype & Azure PowerShell Modules for the Office 365 Signature Creator PowerShell Tool
# ~ First attempt at a "Multi-Function" script in PowerShell - IDK what this is
#>

<# ArrayList of installed Modules #>
$InstalledModules = [System.Collections.ArrayList]@()

<# Initialize the ModuleInstaller Method - Calling Functions in this Module ~ [Void] #>
Function Initialize-ModuleInstaller {

    [CmdletBinding()] param([string[]]$Modules)

    <# Display Install Modules ASCII Message #>
    MI_Message

    <# Initialize the Installer #>
    $AllowedModules = [System.Collections.ArrayList]@("Skype", "Azure")

    <# While loop if modules don't exist - Poop way of doing it [PowerShell Dummy]#>
    while(($AllowedModules | Where-Object {$InstalledModules -NotContains $_}).Count -ne 0) {
        <# Check if Modules are installed - Install if they aren't #>
        ForEach ($Module in $Modules) {
            if(($AllowedModules -contains $Module) -and 
                !(GetRequiredModuleData $Module "INSTALLED")) {
                MI_Install $Module
            }
        }
    }

    #Write-Output "Installed Modules: "$InstalledModules

    <# Import Installed Modules #>
    MI_Import
  
} 

<# Install Modules that are required & are not currently installed on the system ~ [Void] #>
Function MI_Install {

    [CmdletBinding()] param([string]$Module)

    <# Confirm if Module exists - return true, otherwise execute code below #>
    If(isModuleExist $Module) {
        if($InstalledModules -notcontains $Module) {
            ($InstalledModules.Add($Module) | Out-Null)
        }
        Return
    }

    <# Display Install Modules ASCII Message #>
    MI_Message

    <# Install Module from Microsoft's Download Centre - No generic PowerShell repository yet #>
    If($Module -eq "Skype") {
        <# Initialize download through the Web Client class #>
        try {
            If(!(GetRequiredModuleData $Module "DOWNLOAD_FILE_EXISTS")) {
                $WebClient = New-Object System.Net.WebClient
                $WebClient.DownloadFile((GetRequiredModuleData $Module "DIRECT"), (GetRequiredModuleData $Module "DOWNLOAD_FILE_PATH"))
            } else {
                Read-Host "File downloaded - Press the [ENTER] key to open the installer"
                Start-Process -Filepath (GetRequiredModuleData $Module "DOWNLOAD_FILE_PATH")
                Read-Host "An installer window will now pop-up - Once installed, press the [ENTER] key to continue"
            }
        } catch {
            Write-Host "Unable to automatically download the SkypeForBusiness PowerShell Module" -ForegroundColor Red
            Read-Host "Press the [ENTER] key to open the Skype for Business PowerShell download page"
            <# Start the .exe Installer from the WebClient download #>
            Start-Process (GetRequiredModuleData $Module "WEB")
        }
    }

    <# Install Module from generic PowerShell repository #>
    If($Module -eq "Azure") {
        (GetRequiredModuleData $Module "DIRECT")
    }

}

<# Import Installed Modules ~ [Void] #>
Function MI_Import {

    [CmdletBinding()] param([string]$Module)

    <# If Skype Module exists - Import Module to current PowerShell session #>
    if($Module -eq "Skype") {
        <# Firstly, check if the psd1 file exists - Import Manually #>
        If(GetRequiredModuleData $Module "MODULE_INSTALLED_MANUAL") {
            GetRequiredModuleData $Module "IMPORT_MANUAL"
        <# Secondarly, check if the Module is installed - Import from installed repositories #>
        } ElseIf (GetRequiredModuleData $Module "MODULE_INSTALLED") {
            GetRequiredModuleData $Module "IMPORT"
        }
    }

    <# If Azure Module exists - Import Module to current PowerShell session #>
    If($Module -eq "Azure") {
        <# Check if the Module is installed - Import from installed repositories #>
         If(GetRequiredModuleData $Module "MODULE_INSTALLED") {
            GetRequiredModuleData $Module "IMPORT"
         }
    }

}

<# Confirms if the Module Exists ~ [True/False Boolean] #>
Function isModuleExist {

    [CmdletBinding()] param([string]$Module)

    <# Output generic (Checking is installed message) #>
    GetRequiredModuleData $Module "CHECK_INSTALL_MESSAGE"

    <# Confirm if Skype Module is installed #>
    If($Module -eq "Skype") {
        #Method 1 - Manual Psd1 Check
        If(GetRequiredModuleData $Module "MODULE_INSTALLED_MANUAL") {
            GetRequiredModuleData $Module "SUCCESS_INSTALL_MESSAGE"
            Return $True
        }
        #Method 2 - Module List Check
        If(GetRequiredModuleData $Module "MODULE_INSTALLED") {
            GetRequiredModuleData $Module "SUCCESS_INSTALL_MESSAGE"
            Return $True
        }
        <# Return Module is not installed boolean #>
        Return $False
    }

    <# Confirm if Azure Module is installed #>
    If($Module -eq "Azure") {
        #Method 1 - Module List Check
        If(GetRequiredModuleData $Module "MODULE_INSTALLED") {
            GetRequiredModuleData $Module "SUCCESS_INSTALL_MESSAGE"
            Return $True
        }
        <# Return Module is not installed boolean #>
        Return $False
    }

}

<# Module DATA ~ [Function Boolean] #>
Function GetRequiredModuleData {

    [CmdletBinding()]
    param(
        [string]$Module,
        [string]$Action
    )

    If($Module -eq "Skype") {
        switch($Action) {
            "DIRECT" { return "https://download.microsoft.com/download/2/0/5/2050B39B-4DA5-48E0-B768-583533B42C3B/SkypeOnlinePowerShell.exe" }
            "WEB" { return "https://www.microsoft.com/en-gb/download/details.aspx?id=39366" }
            "IMPORT" { return Import-Module SkypeOnlineConnector }
            "IMPORT_MANUAL" { return Import-Module "C:\\Program Files\\Common Files\\Skype for Business Online\\Modules\\SkypeOnlineConnector\\SkypeOnlineConnector.psd1" }
            "MODULE_INSTALLED" { return Get-Module -ListAvailable -Name SkypeOnlineConnector }
            "MODULE_INSTALLED_MANUAL" { return Test-Path -Path "C:\Program Files\Common Files\Skype for Business Online\Modules\SkypeOnlineConnector\SkypeOnlineConnector.psd1" }
            "DOWNLOAD_FILE_PATH" { return "$env:temp/SkypeOnlineConnector.exe" }
            "DOWNLOAD_FILE_EXISTS" { return Test-Path -Path "$env:temp/SkypeOnlineConnector.exe" }
            "CHECK_INSTALL_MESSAGE" { return Write-Host "Checking if $Module PowerShell Module exists...`n" -ForegroundColor Yellow }
            "SUCCESS_INSTALL_MESSAGE" { return Write-Host "$Module PowerShell Module is installed`n" -ForegroundColor Green }
            "INSTALLED" { return $InstalledModules.Contains($Module) }
        }
    } ElseIf ($Module -eq "Azure") {
        switch($Action) {
            "DIRECT" { return Install-Module AzureAD }
            "IMPORT" { return Import-Module AzureAD }
            "MODULE_INSTALLED" { return Get-Module -ListAvailable -Name AzureAD }
            "CHECK_INSTALL_MESSAGE" { return Write-Host "Checking if $Module PowerShell Module exists...`n" -ForegroundColor Yellow }
            "SUCCESS_INSTALL_MESSAGE" { return Write-Host "$Module PowerShell Module is installed`n" -ForegroundColor Green }
            "INSTALLED" { return $InstalledModules.Contains($Module) }
        }
    } ElseIf($Module -eq "Exchange") {
        #TODO - Exchange is [Currently] using Basic Authentication
    }

    return "Invalid Module /or Action."

}

<# Output ASCII message for ModuleInstaller #>
Function MI_Message {
Clear-Host
@"
    ___ _____       _         ___         _        _ _   __  __         _      _        
   | _ \_   _|     / |  ___  |_ _|_ _  __| |_ __ _| | | |  \/  |___  __| |_  _| |___ ___
   |  _/ | |    _  | | |___|  | || ' \(_-<  _/ _`` | | | | |\/| / _ \/ _`` | || | / -_|_-<
   |_|   |_|   (_) |_|       |___|_||_/__/\__\__,_|_|_| |_|  |_\___/\__,_|\_,_|_\___/__/
                                                                                        
"@
}  