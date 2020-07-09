<#
New-RecreatedSharedMailbox
Author: SuperDOS
Last Edit: SuperDOS
Version: 0.2
Date: 13:37 2020-03-30
Description:
Make onprem migrated shared mailbox online only

Make sure you connect to Msonline and MSExchangeOnline before running
You must use the mailbox alias attribute when adding which mailboxes that need to be recreated.
Also Alias/MailNickname need to be same as samaccountname
Don't add back send on behalf look it up before!
Get-exoMailbox -RecipientTypeDetails sharedmailbox -Properties GrantSendOnBehalfTo | select alias,GrantSendOnBehalfTo
Make sure these values are on all the shared mailboxes that you want to recreate!
This happens if first migrated as user and then converted to shared mailbox
msExchRecipientDisplayType	-2147483642
msExchRecipientTypeDetails	34359738368
msExchRemoteRecipientType	100
get-aduser -Filter {msExchRemoteRecipientType -eq 4} -SearchBase "OU=Shared Mailboxes,OU=Exchange Resources,OU=Company,DC=Company,DC=com" `
| set-aduser -Replace @{'msExchRemoteRecipientType'=100;'msExchRecipientDisplayType'=-2147483642;msExchRecipientTypeDetails=34359738368}


NEEDED MODULES:
ExchangeOnlineManagement -- https://www.powershellgallery.com/packages/ExchangeOnlineManagement
ActiveDirectory
MSOnline -- https://www.powershellgallery.com/packages/MSOnline


CHANGELOG:
0.1: First version!
0.2: Added parameter and functions to accept multiple mailboxes
#>

#Disable Positional Parameters
[CmdletBinding(PositionalBinding = $false)] 
# Define Parameters
Param (
    [parameter(mandatory = $false)][alias('ms')] [array]$Mailboxes,	#alias of mailboxes
    [parameter(mandatory = $false)][alias('d')] [switch]$RemoveUserAccount	#Switch for deleting the AD User 
)

#Set Global ErrorActionPreference
$global:ErrorActionPreference = 'Stop'

$ExportDirectory = ".\ExportedAddresses\"
$domain = "company.com" # change this to your mail domain
$OnlineDomain = "company.onmicrosoft.com" #change this to your MSOdomain
$OldMailboxes = @()
$OldMBMembers = @()

#Updated Start-Sleep Function to show progress
function Start-Sleep($seconds) {
    $doneDT = (Get-Date).AddSeconds($seconds)
    while ($doneDT -gt (Get-Date)) {
        $secondsLeft = $doneDT.Subtract((Get-Date)).TotalSeconds
        $percent = ($seconds - $secondsLeft) / $seconds * 100
        Write-Progress -Activity "Sleeping" -Status "Sleeping..." -SecondsRemaining $secondsLeft -PercentComplete $percent
        [System.Threading.Thread]::Sleep(500)
    }
    Write-Progress -Activity "Sleeping" -Status "Sleeping..." -SecondsRemaining 0 -Completed
}

#Get Last AD Sync
$strCurrentTimeZone = (Get-WmiObject win32_timezone).StandardName
$TZ = [System.TimeZoneInfo]::FindSystemTimeZoneById($strCurrentTimeZone)
$LastDirSyncTime = [System.TimeZoneInfo]::ConvertTimeFromUtc((Get-MsolCompanyInformation).LastDirSyncTime, $TZ)
$LastDirSyncAge = (Get-Date) - $LastDirSyncTime

#If AD sync is triggered soon please wait since we need to export data from mailboxes
if ($LastDirSyncAge.Minutes -gt 28) {
    Start-Sleep -Seconds 300
}
#Remove Mailboxes from Sync and store mailboxes-object
foreach ($User in $Mailboxes) {
    #Set-aduser $User -Add @{adminDescription = "User_DoNotSync" }
    $OldMailboxes += Get-EXOMailbox $User
    $OldMembers = Get-EXOMailboxPermission $User | Where-Object { $_.user -like '*' + $domain } | Select-Object -Property Identity, user, @{ Name = 'Accessrights'; Expression = { $_.accessrights } }
    $OldMBMembers += $OldMembers

    #Check if exportdir exists
    If (!(Test-Path -Path $ExportDirectory )) {
        Write-Host "  Creating Directory: $ExportDirectory"
        New-Item -ItemType directory -Path $ExportDirectory | Out-Null
    }

    #Access rights for mailbox, export to csv for backup
    $OldMembers | Export-Csv -Path "$ExportDirectory\$User.csv" -Delimiter ';' -Encoding utf8 -NoTypeInformation 

}

"Last sync: " + $LastDirSyncAge.Minutes + " minutes ago"
#Wait until AD Sync is complete so mailboxes is soft deleted
Start-Sleep -Seconds $(New-TimeSpan -Minutes $(30 - $LastDirSyncAge.Minutes)).TotalSeconds

foreach ($Mailbox in $OldMailboxes) {
    
    $ExchGUID = $null
    $OldMB = $null
    $OldPrimarySmtpAddress = $null
    $OldMBAlias = $null
    $OldMBIdentity = $null

    #Get old mailbox settings
    $OldMB = $Mailbox
    $OldMBDisplayname = [string]$OldMB.Displayname
    $OldPrimarySmtpAddress = [string]$OldMB.PrimarySmtpAddress
    $OldMBAlias = [string]$OldMB.Alias
    $OldMBIdentity = [string]$OldMB.Identity
    Do {
        Try {    
            $ExchGUID = ((get-exomailbox -SoftDeletedMailbox $Mailbox.alias -Properties ExchangeGuid).ExchangeGuid).Guid     
        }
        catch { Start-Sleep -seconds 5 }
    
    } While ($null -eq $ExchGUID)

    Start-Sleep -seconds 60
    #Recreate the inactivemailbox
    $NewMB = New-Mailbox -InactiveMailbox $ExchGUID -name $OldMBAlias -alias $OldMBAlias -displayname $OldMBDisplayname -MicrosoftOnlineServicesID $OldMBAlias$OnlineDomain 

    Start-Sleep -seconds 60

    #CONVERT TO SHARED MAILBOX
    Set-Mailbox $NewMB.alias -Type Shared

    #Add permissions
    foreach ($Mem in $OldMBMembers) {
        if ($Mem.Identity -eq $OldMBIdentity) {       
            Add-MailboxPermission -Identity $NewMB.alias -AccessRights $Mem.AccessRights -User $Mem.user | Out-Null
            Add-RecipientPermission -Identity $NewMB.alias -AccessRights SendAs -Trustee $Mem.user -Confirm:$false | Out-Null
        }
    }

    #Add back primary SMTP-address
    Set-Mailbox $NewMB.alias -WindowsEmailAddress $OldPrimarySmtpAddress

    if ($RemoveUserAccount) {
        #Remove AD User
        Remove-ADUser $Mailbox    
    }
    
}
