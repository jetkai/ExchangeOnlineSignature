<#
# Create Signature - Created by -Kai (https://social.technet.microsoft.com/profile/-kai/)
# Created 21st November, 2019
#
# Module/Functions is used for installing Skype & Azure PowerShell Modules for the Office 365 Signature Creator PowerShell Tool
# ~ First attempt at a "Multi-Function" script in PowerShell - IDK what this is
#>

Function Initialize-CreateSignature {
    <# Display ASCII Message #>
    CS_Message

    <# Create TEMP directory - if needed (Stores signature backups)#>
    CS_CreateFolderDirectory

    <# Dump data from Office 365 ~ Used for signatures #>
    CS_UpdateData

    <# Builds useful data into PSObjects & ArrayList #>
    CS_BuildUsersData
    CS_BuildCompanyData

    <# Finally create the signature #>
    CS_CreateSignature
}

<# Office 365 Data #>
Function CS_UpdateData {
    #ArrayList ~ Stores all user information for every user on the tenant with a mailbox
    $script:GlobalUserList = [System.Collections.ArrayList]@()
    #ArrayList ~ Only is used when executing the final stage
    $script:FinalUserList = [System.Collections.ArrayList]@()
    #All Azure/Office365 Users
    $script:AzureAdUserList = Get-AzureADUser
    #Get All Mailboxes
    $script:ExchangeUserList = Get-Mailbox
    #All SFB Numbers/ObjectId
    $script:SkypeLineUriList = Get-CsOnlineUser | Select-Object -Property "ObjectId", "LineUri"
    #Temp Location Path
    $script:RootTempDirectory = "$env:temp\365Signatures"
}

<#Create TEMP Folder Structure#>
Function CS_CreateFolderDirectory {
    $Folders = @("", "\Backup", "\Export", "\Profile pictures")

    ForEach($Folder in $Folders) {
        $TempFolderPath = $RootTempDirectory+$Folder
        If(!(Test-Path $TempFolderPath)) { New-Item -ItemType Directory -Force -Path $TempFolderPath }
    }
    #If(!(Test-Path "$RootTempDirectory")) { New-Item -ItemType Directory -Force -Path "$RootTempDirectory" }
    #If(!(Test-Path "$RootTempDirectory\Backup")) { New-Item -ItemType Directory -Force -Path "$RootTempDirectory\Backup" }
    #If(!(Test-Path "$RootTempDirectory\Export")) { New-Item -ItemType Directory -Force -Path "$RootTempDirectory\Export" }
}

<# Build Users_MAIN_LIST with all User Information; DisplayName, Email, SFB PhoneNumber, Company & Address Info #>
Function CS_BuildUsersData {

    $script:TotalUsersIntegerAmount = 0

    ForEach($LocalUser in $script:AzureAdUserList) {

        #Check if Mailbox exists for user before adding them to the MAIN_LIST
        $isMailbox = $script:ExchangeUserList | Where-Object "ExternalDirectoryObjectId" -eq $LocalUser.ObjectId
        
        if(!$isMailbox) {
            continue
        }

        $User_Azure_TelephoneNumber = $LocalUser.TelephoneNumber
        $User_Azure_Mobile = $LocalUser.Mobile
        $User_Skype_PhoneNumber = ($script:SkypeLineUriList | Where-Object "ObjectId" -eq $LocalUser.ObjectId).LineUri
        
        if($User_Skype_PhoneNumber -like '*tel:*') {
            $User_Skype_PhoneNumber = $User_Skype_PhoneNumber.Replace("tel:", "")
        }

        $User_PhoneNumberList = @($User_Skype_PhoneNumber, $User_Azure_TelephoneNumber, $User_Azure_Mobile)
        $PhoneNumber = $null
        #Autopick phone number if SFB is not available
        ForEach($LocalPhoneNumber in $User_PhoneNumberList) {
            if($LocalPhoneNumber.Length -gt 1 -and $null -eq $PhoneNumber) {
                $PhoneNumber = $LocalPhoneNumber
                 continue
            }
        }

        #$LocalUser.Mail Attribute can be broken sometimes, due to AzureAD - You can change this if needed
        $User_PSObject = [PsCustomObject]@{
        ID=$script:TotalUsersIntegerAmount;
        DisplayName=$LocalUser.DisplayName;
        FirstName=$LocalUser.GivenName;
        LastName=$LocalUser.Surname;
        Email=$LocalUser.UserPrincipalName;
        PhoneNumber=$PhoneNumber;
        Fax=$LocalUser.FacsimileTelephoneNumber;
            
        StreetAddress=$LocalUser.StreetAddress;
        City=$LocalUser.City;
        Country=$LocalUser.Country;   
        State=$LocalUser.State;
        PostCode=$LocalUser.PostalCode;
        JobTitle=$LocalUser.JobTitle;
        ObjectId=$LocalUser.ObjectId}

        $script:GlobalUserList.Add($User_PSObject) | Out-Null
        $script:TotalUsersIntegerAmount++
    }
}

<# Update Company Information & Address #>
Function CS_BuildCompanyData {

    $script:Company_Azure_Info = Get-AzureADTenantDetail
    $script:Company_Name = $Company_Azure_Info.DisplayName
    $script:Company_WebSite = ($Company_Azure_Info.VerifiedDomains | Where-Object -Property "_Default" -EQ "True").Name

    $script:Company_Address = [PsCustomObject]@{
    StreetAddress=$Company_Azure_Info.Street;
    PhoneNumber=$Company_Azure_Info.TelephoneNumber;
    City=$Company_Azure_Info.City;
    Country=$Company_Azure_Info.Country;   
    State=$Company_Azure_Info.State;
    PostCode=$Company_Azure_Info.PostalCode }

}

 <# Show everyone who has a mailbox in formatted table form #>
Function CS_DisplayUsersAsTable {
    return $script:GlobalUserList | Format-Table -Property ID,DisplayName,Email,PhoneNumber | Out-String | ForEach-Object { Write-Host $_ -ForegroundColor Green }
}


Function CS_CreateSignature {

    <# Show everyone who has a mailbox in formatted table form #>
    CS_DisplayUsersAsTable

    while($true) {

        Write-Host ">> [CTRL+C] TO EXIT <<`n"
        $WhoSignature = Read-Host "Who's signature would you like to update? ALL, ID or Alias ~ TotalUsers:"($script:TotalUsersIntegerAmount-1)

        if($WhoSignature -eq "all") {
            $script:FinalUserList = $script:GlobalUserList

        } elseif($WhoSignature -match "^\d+$") {
            #Check if the value of the user exists in the list
            if($null -ne $script:GlobalUserList[$WhoSignature]) {
                $script:FinalUserList.Add($script:GlobalUserList[$WhoSignature]) | Out-Null
            }

        } elseif($WhoSignature -like '*@*') {
            $LocalUser = $script:GlobalUserList | Where-Object Email -EQ $WhoSignature
            $script:FinalUserList.Add($LocalUser) | Out-Null
        } else {
            Write-Host "Invalid syntax." -ForegroundColor "Red"
            continue
        }

        $Preview = Read-Host "Would you like to preview the signatures before automatically applying them? [Y/N]"

        Write-Host "`n~ Colours ~`n"
        Write-Host ">> Red" -ForegroundColor Red 
        Write-Host ">> Green" -ForegroundColor Green
        Write-Host ">> Gray" -ForegroundColor Gray
        Write-Host ">> Magenta" -ForegroundColor Magenta
        Write-Host ">> Black" -ForegroundColor White
        Write-Host ">> Blue" -ForegroundColor Blue
        Write-Host ">> Yellow" -ForegroundColor Yellow
        Write-Host ">> Custom [#69adf6]`n"
        
        $AllowedColours = @("Red", "Green", "Gray", "Magenta", "Black", "Blue", "Yellow", "Custom")
        $ColourHex = [PsCustomObject]@{
            Red="#FF0000";
            Green="#20C162";
            Gray="#3D3D3D";
            Magenta="#5209D8";
            Blue="1A38E3";
            Black="#000000";
            Yellow="#F7BC03"
        }
        $ColourStyle = Read-Host "What colour theme would you like to use? [Press Enter to SKIP]"
        $FinalColourStyle = "#69adf6"

        if($AllowedColours -contains $ColourStyle) {
            if($ColourStyle -like "*custom*") {
                $FinalColourStyle = Read-Host "Enter your HTML Colour Hex"
                if($FinalColourStyle -notlike "*#*") {
                    $FinalColourStyle = "#"+$FinalColourStyle
                }
                Write-Host "Signature theme updated to $FinalColourStyle"
            } else {
                $FinalColourStyle = $ColourHex.$ColourStyle
                Write-Host "Signature theme updated to $ColourStyle" -ForegroundColor $ColourStyle
            }
        } else {
            Write-Host "Using default theme, as the colour theme you chose is invalid."
        }

        ForEach($LocalUser in $script:FinalUserList) {

        if($null -eq $LocalUser) {
            continue;
        }

        if($Preview -notlike "*y*" -and $Preview -notlike "*n*") {
            Write-Host "Invalid syntax." -ForegroundColor "Red"
            continue;
        }

        #Get Signature File & Store in memory
        #You can replace your own signature with Signature_Default @CompressSignature is a backup incase the .ps1 is ran out of the root file directory

        $CompressedSignature = @'
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"><HTML lang="en"><HEAD><TITLE>Email Signature</TITLE> <META content="text/html; charset=utf-8" http-equiv="Content-Type"></HEAD><BODY style="font-size:10pt; font-family: Verdana, sans-serif;"><table style="width:300px; font-size:10pt; font-family: Verdana, sans-serif; color:#69adf6;" width="300" cellpadding="0" cellspacing="0"> <tbody> <tr> <td style="padding:0; line-height:27px; vertical-align:top; font-family: Verdana, sans-serif; font-size:14pt; color:#69adf6;" valign="top"> <strong><span style="font-size:14pt; color:#69adf6; font-family:Verdana, sans-serif;">_FIRSTNAME_&nbsp;_LASTNAME_</span></strong> </td> </tr> <tr> <td style="padding:0; line-height:18px; vertical-align:top; font-family:Verdana, sans-serif; color:#000000; font-size:10pt;" valign="top"> <span style="font-size:10pt; color:#000000; font-family:Verdana, sans-serif;">_JOBTITLE_<br><br></span> </td> </tr> <tr> <td style="padding: 5px 0;line-height:18px; border-bottom:1px solid; border-bottom-color:#69adf6; font-family:Verdana, sans-serif; color:#000000; font-size:10pt; vertical-align:top;" valign="top"> <span style="font-family:Verdana, sans-serif; font-size:10pt; color:#000000;">_ADDRESS1_</span><span style="color:#69adf6;"> | </span> <span style="font-family:Verdana, sans-serif; font-size:10pt; color:#000000;">_ADDRESS2_</span> <span style="font-family:Verdana, sans-serif; font-size:10pt; color:#000000;"><span><br></span><b>_COMPANYNAME_</b></span> </td> </tr> <tr> <td style="padding: 5px 0 0;line-height:18px; font-family:Verdana, sans-serif; color:#000000; font-size:10pt; vertical-align:top;" valign="top"> <span style="font-family:Verdana, sans-serif; font-size:10pt;"><a href="tel:_MOBILENUMBER_" style="text-decoration:none;"><span style="color:#69adf6">📞&nbsp;&nbsp;_MOBILENUMBER_</span></a><span><br></span> </span> <span style="font-family:Verdana, sans-serif; font-size:10pt;"><a href="tel:_TELEPHONENUMBER_" style="text-decoration:none;"><span style="color:#69adf6">☎️&nbsp;&nbsp;_TELEPHONENUMBER_</span></a><span><br></span> </span> <span style="font-family:Verdana, sans-serif; font-size:10pt;"><a href="mailto:_EMAILADDRESS_" style="text-decoration:none;"><span style="color:#69adf6">📧&nbsp;&nbsp;_EMAILADDRESS_</span></a><span><br></span> </span> <span style="font-family:Verdana, sans-serif; font-size:10pt;">🌐&nbsp;&nbsp;<a href="https://_WEBSITE_" target="_blank" rel="noopener" style="text-decoration:none;"><span style="font-size:10pt;font-family:Verdana, sans-serif;color:#69adf6"><span style="font-size:10pt;font-family:Verdana, sans-serif;color:#69adf6">_WEBSITE_</span></span></a> </span> </td> </tr> </tbody></table></BODY></HTML>
'@

        if(Test-Path -Path ".\templates\Signature_Default.htm") {
            $Content = Get-Content ".\templates\Signature_Default.htm" -Encoding UTF8
        } else {
            $Content = $CompressedSignature
        }

        #Set Colour Styling Theme for Signature
        if($null -ne $FinalColourStyle) {
            $Content = $Content.Replace("#69adf6", $FinalColourStyle)
        }

        #You can modify all these attributes to anything you like
        #Local Attributes

        #If First Name & Last Name is Empty, Failback to DisplayName
        if($null -eq $LocalUser.FirstName -and $null -eq $LocalUser.LastName) {
            $Content = $Content.Replace("_FIRSTNAME_", (Get-Culture).TextInfo.ToTitleCase($LocalUser.DisplayName))
            $Content = $Content.Replace("_LASTNAME_", "")
        } else {
            $Content = $Content.Replace("_FIRSTNAME_", (Get-Culture).TextInfo.ToTitleCase($LocalUser.FirstName))
            $Content = $Content.Replace("_LASTNAME_", (Get-Culture).TextInfo.ToTitleCase($LocalUser.LastName))
        }
        $Content = $Content.Replace("_JOBTITLE_", (Get-Culture).TextInfo.ToTitleCase($LocalUser.JobTitle))
        $Content = $Content.Replace("_EMAILADDRESS_", $LocalUser.Email)
        $Content = $Content.Replace("_MOBILENUMBER_", $LocalUser.PhoneNumber)

        $LocalAttributes = $false
        
        if($LocalAttributes) {
            $Content = $Content.Replace("_ADDRESS1_", (Get-Culture).TextInfo.ToTitleCase($LocalUser.StreetAddress))
            $Address2 = @($LocalUser.City, $LocalUser.PostalCode)
            $Content = $Content.Replace("_ADDRESS2_", ($Address2 -join ", "))
        } else {
            $Content = $Content.Replace("_ADDRESS1_", (Get-Culture).TextInfo.ToTitleCase($Company_Address.StreetAddress))
            $Address2 = @((Get-Culture).TextInfo.ToTitleCase($Company_Address.City), $Company_Address.PostCode.ToUpper())
            $Content = $Content.Replace("_ADDRESS2_", ($Address2 -join ", "))
        }

        #Global Attributes
        $Content = $Content.Replace("_COMPANYNAME_", (Get-Culture).TextInfo.ToTitleCase($Company_Name))
        $Content = $Content.Replace("_TELEPHONENUMBER_", $Company_Azure_Info.TelephoneNumber)
        $Content = $Content.Replace("_WEBSITE_", $Company_Website)

        if($Preview -like "*y*") {
            $Content | Out-File "$RootTempDirectory\temp.htm"
            Start-Process "$RootTempDirectory\temp.htm"
            $PreviewFinal = "No"
            $PreviewFinal = Read-Host "Press [Y] to update signature, [N] to cancel - ["$LocalUser.Email"]"
        }

        #Upload Signature Data to Outlook Web App for User
        if($null -eq $PreviewFinal -or $PreviewFinal -like "*y*") {
            $CurrentTime = (Get-Date).ToString('MM-dd-yyyy_hh-mm-ss_tt')
            $UserEmail = $LocalUser.Email
            #Backup Signature - Before Updating Signature
            Get-MailboxMessageConfiguration -Identity $LocalUser.ObjectId | Select-Object -ExpandProperty SignatureHtml | Out-File -FilePath "$RootTempDirectory\Backup\$CurrentTime-$UserEmail.htm" -Encoding UTF8
            #Update New Signature
            Set-MailboxMessageConfiguration -Identity $LocalUser.ObjectId -SignatureHtml $Content -AutoAddSignature $true -AutoAddSignatureOnMobile $true -AutoAddSignatureOnReply $true
            #Export HTML Signature - After Updating Signature
            $Content | Out-File -FilePath "$RootTempDirectory\Export\$CurrentTime-$UserEmail.htm" -Encoding UTF8

            Write-Host "["$LocalUser.Email"] - Updated signature..." -ForegroundColor Green
        } else {
            Write-Host "["$LocalUser.Email"] - Unable to update signature, as you requested not to update the signature..." -ForegroundColor Red
        }
        $PreviewFinal = $null
    }

    #Reset the Email List
    $script:FinalUserList = [System.Collections.ArrayList]@()
    $Preview = $null
    }
}

Function CS_SaveProfilePicture {

    [CmdletBinding()] param([string]$Email)

    $UserPhoto = Get-UserPhoto $Email

    $SaveLocation = $script:RootTempDirectory+"\Profile pictures"

    $FileName = $UserPhoto.Identity+".jpg"

    $UserPhoto.PictureData | Set-Content ($SaveLocation+$FileName) -Encoding byte

}

Function CS_Message {
Clear-Host
@"
        ____ _____       _____            ____                _         ____  _                   _                  
       |  _ \_   _|     |___ /           / ___|_ __ ___  __ _| |_ ___  / ___|(_) __ _ _ __   __ _| |_ _   _ _ __ ___ 
       | |_) || |         |_ \   _____  | |   | '__/ _ \/ _`` | __/ _ \ \___ \| |/ _`` | '_ \ / _`` | __| | | | '__/ _ \
       |  __/ | |    _   ___) | |_____| | |___| | |  __/ (_| | ||  __/  ___) | | (_| | | | | (_| | |_| |_| | | |  __/
       |_|    |_|   (_) |____/           \____|_|  \___|\__,_|\__\___| |____/|_|\__, |_| |_|\__,_|\__|\__,_|_|  \___|
                                                                                |___/                                
"@
}