<#
# Connect Office - Created by -Kai (https://social.technet.microsoft.com/profile/-kai/)
# Created 21st November, 2019
#
# Module/Functions is used for installing Skype & Azure PowerShell Modules for the Office 365 Signature Creator PowerShell Tool
# ~ First attempt at a "Multi-Function" script in PowerShell - IDK what this is
#>

<# ArrayList of installed Modules #>
$ImportedSessions = [System.Collections.ArrayList]@()

<# Initialize the ConnectOffice Method - Calling Functions in this Module ~ [Void] #>
Function Initialize-ConnectOffice {

    [CmdletBinding()] param([string[]]$Sessions)

    <# Display Connect 365 ASCII Message #>
    CO_Message

    <# Request user-input to enter in their Office 365 credentials #>

    <# Initialize the Installer #>
    $AllowedSessions = [System.Collections.ArrayList]@("Skype", "Azure", "Exchange")

    <# While loop if modules don't exist - bad way of doing it [PowerShell Dummy]#>
    while(($AllowedSessions | Where-Object {$ImportedSessions -NotContains $_}).Count -ne 0) {
        <# Check if Modules are installed - Install if they aren't #>
        ForEach ($Session in $Sessions) {
            if(($AllowedSessions -contains $Session)) {
                CO_Login # CO_Login [Void] will be skipped if UserCred
                CO_CreateSession $Session 
            }
        }
        Start-Sleep -Seconds 5
    }

    #GetServiceData "Skype" "NEW_SESSION"

    #Write-Output "TestOutput: "$SkypeSession
}

<# Create session needed to connect to the required services #>
Function CO_CreateSession {

    [CmdletBinding()] param([string]$Session)

    $script:CO_CreateSessionNewException = $null
    $script:CO_CreateSessionImportException = $null

    <# Check if session already exists before creating a new session #>
    $ActiveSession = GetRequiredSessionData $Session "ACTIVE_SESSION"
    If(($null -eq $ActiveSession) -or ($ActiveSession -eq $false)) {
        <# Try to create a new session with provided UserCredentials #>
        Write-Host "Attempting to create & import a new $Session session..."
        try {
            <# Initiate a new session #>
            GetRequiredSessionData $Session "NEW_SESSION"
        <# TryCatch errors - used for letting the user know what is wrong #>
        } catch {
            Write-Output "Catch Error1 for: " $Session
            <# Error Outputs #>
            $InvalidCredentials = @("The remote server returned an error: (463)", #Skype for Business/Lync
                                    "Access is denied.", #Exchange Basic Auth
                                    "Invalid username or password.", #Azure AD - Email exists but password is in-correct
                                    "Authentication Failure") #Azure AD - Email does not exist

            $LoginLimitExceeded = @("maximum number of concurrent shells for this user has been exceeded") #Skype for Business/Lync

            $script:CO_CreateSessionNewException = $_.Exception.Message
            $script:CO_CreateSessionNewException

            <# Warn user that the credentials are incorrect or MFA is enabled & reset the UserCredential to re-prompt username + password #>
            If(($InvalidCredentials | Where-Object {$script:CO_CreateSessionNewException -like "*$_*"}).Count -gt 0) {
                Write-Host "Invalid username or password / OR Multi-Factor Authentication HAS to be DISABLED." -ForegroundColor Red
                $script:UserCredential = $null
            }
            <# Alert user if they have reached the maximum login limit - caused by connecting to office 365 too many times #>
            If(($LoginLimitExceeded | Where-Object {$script:CO_CreateSessionNewException -like "*$_*"}).Count -gt 0) {
                Write-Host "You have reached the maximum login limit. Please wait 15-30 minutes and try again." -ForegroundColor Red
                $script:UserCredential = $null
            }
        }

        <# If there are no errors with creating the session - Import the session into current shell #>
        If($null -eq $script:CO_CreateSessionImportException) {
            try {
                Write-Host "Importing $Session session!" -ForegroundColor Green
                GetRequiredSessionData $Session "IMPORT_SESSION"
            } catch {
                Write-Output "Catch Error2 for: " $Session
                $script:CO_CreateSessionImportException = $_.Exception.Message
            }
        }

        <# If there are no errors on both CreateNewSession & ImportSession - Add to ImportedSessions ArrayList #>
        If(($null -eq $script:CO_CreateSessionNewException) -and ($null -eq $script:CO_CreateSessionImportException)) {
            If($ImportedSessions -notcontains $Session) {
                ($ImportedSessions.Add($Session) | Out-Null)
            } 
        }
    <# Add session to ImportedSessions if connection is already active #>
    } else {
        If($ImportedSessions -notcontains $Session) {
            ($ImportedSessions.Add($Session) | Out-Null)
        } 
    }
}

<# Simple command-line email + password login interface #>
Function CO_Login {

    <# Check if UserCredential is null before proceeding to while-loop #>
    If($null -ne $script:UserCredential) {
        return
    }

    <# Loop credential request - user will need to input email/password #>
    while($null -eq $script:UserCredential) {

        Write-Host "Please enter your Office 365 credentials below:`n" -ForegroundColor Yellow

        $Email = Read-Host "Email Address"

        <# Check if the email address is valid before proceeding #>
        $EmailRegex = "^[a-zA-Z0-9.!£#$%&'^_`{}~-]+@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)*$"
        If(!$Email -cmatch $EmailRegex) {
            CO_Message
            Write-Host "[ERROR]: Invalid email address - You must include an @ symbol." -ForegroundColor Red
            continue
        }

        $Password = Read-Host "Password" -AsSecureString

        <# Confirm the password is at least 1 character long (un-sure of Azure's min char limit) #>
        if($Password.Length -lt 1) {
            CO_Message
            Write-Host "[ERROR]: Invalid password - Must be at least 1 character long." -ForegroundColor Red
            continue
        }

        <# Set script-local UserCredential to use for creating sessions #>
        $script:UserCredential = New-Object System.Management.Automation.PSCredential -ArgumentList $Email, $Password
    }

}

<# Session DATA ~ [Function Boolean] #>
Function GetRequiredSessionData {

    [CmdletBinding()]
    param(
        [string]$Session,
        [string]$Action
    )

    If($Session -eq "Skype") {
        switch($Action) {
            "NEW_SESSION" { return $global:SkypeSession = New-CsOnlineSession -Credential $script:UserCredential  }
            "IMPORT_SESSION" { return Import-Module (Import-PSSession $global:SkypeSession -AllowClobber) -Global }
            "ACTIVE_SESSION" { return (Get-PSSession | Where-Object -Property "ComputerName" -like "*online.lync*" | Select-Object -ExpandProperty Availability) -eq "Available" }
            "ACTIVE_CONNECTION" { return (Get-PSSession | Where-Object -Property "ComputerName" -like "*online.lync*" |  Select-Object -ExpandProperty State) -eq "Opened" }
        }
    } ElseIf($Session -eq "Azure") {
        switch($Action) {
            "NEW_SESSION" { return $global:AzureSession = Connect-AzureAD -Credential $script:UserCredential  }
            "IMPORT_SESSION" { }
            "ACTIVE_SESSION" { return $null -ne $global:AzureSession.TenantId }
        }
    } ElseIf($Session -eq "Exchange") {
        switch($Action) {
            "NEW_SESSION" { return $global:ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $script:UserCredential -Authentication Basic -AllowRedirection }
            "IMPORT_SESSION" { return Import-Module (Import-PSSession $global:ExchangeSession -DisableNameChecking -AllowClobber) -Global  }
            "ACTIVE_SESSION" { return (Get-PSSession | Where-Object -Property "ComputerName" -like "*office365*" | Select-Object -ExpandProperty Availability) -eq "Available" }
            "ACTIVE_CONNECTION" { return (Get-PSSession | Where-Object -Property "ComputerName" -like "*office365*" |  Select-Object -ExpandProperty State) -eq "Opened" }
        }
    }
}

Function CO_Message {
Clear-Host
@"
    ____ _____       ____             ____                            _    ____    _____  __  ____  
   |  _ \_   _|     |___ \           / ___|___  _ __  _ __   ___  ___| |_  \ \ \  |___ / / /_| ___| 
   | |_) || |         __) |  _____  | |   / _ \| '_ \| '_ \ / _ \/ __| __|  \ \ \   |_ \| '_ \___ \ 
   |  __/ | |    _   / __/  |_____| | |__| (_) | | | | | | |  __/ (__| |_   / / /  ___) | (_) |__) |
   |_|    |_|   (_) |_____|          \____\___/|_| |_|_| |_|\___|\___|\__| /_/_/  |____/ \___/____/ 
                                                                                                    
"@
}