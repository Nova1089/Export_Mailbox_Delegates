# Version 1.0

# functions
function Show-Introduction
{
    Write-Host "This script exports the delegates of a mailbox." -ForegroundColor "DarkCyan"
    Read-Host "Press Enter to continue"
}

function TryConnect-ExchangeOnline
{
    $connectionStatus = Get-ConnectionInformation -ErrorAction SilentlyContinue

    while ($null -eq $connectionStatus)
    {
        Write-Host "Connecting to Exchange Online..." -ForegroundColor DarkCyan
        Connect-ExchangeOnline -ErrorAction SilentlyContinue
        $connectionStatus = Get-ConnectionInformation

        if ($null -eq $connectionStatus)
        {
            Read-Host -Prompt "Failed to connect to Exchange Online. Press Enter to try again"
        }
    }
}

function PromptFor-Mailbox
{
    $mailboxEmail = Read-Host "Enter the email address of the mailbox"
    $mailboxEmail = $mailboxEmail.Trim()
    $mailbox = Get-ExoMailbox -Identity $mailboxEmail -ErrorAction "Stop"
    return $mailbox
}

function Export-MailboxDelegates($mailbox)
{
    Write-Host "Exporting mailbox delegates..." -ForegroundColor "DarkCyan"
    $timeStamp = Get-Date -Format 'yyyy-MM-dd-hh-mmtt'
    $readDelegates = Get-EXOMailboxPermission -Identity $mailbox.UserPrincipalName -ResultSize "Unlimited" | 
        Where-Object { $_.User -inotlike "*NT AUTHORITY*" } |
        Select-Object -Property @( 
        @{ Label = "Mailbox"; Expression = { $mailbox.UserPrincipalName } }, 
        @{ Label = "Delegate"; Expression = { $_.User } }, 
        @{ Label = "AccessRights"; Expression = { Convert-ListToText $_.AccessRights } }, 
        "IsInherited", 
        "InheritanceType" ) |
        Export-Csv -Path "$PSScriptRoot\$($mailbox.DisplayName)_Mailbox_Delegates_$timeStamp.csv" -Append -Force -NoTypeInformation
    $sendDelegates = Get-EXORecipientPermission -Identity $mailbox.UserPrincipalName -ResultSize "Unlimited" | 
        Where-Object { $_.Trustee -inotlike "*NT AUTHORITY*" } |
        Select-Object -Property @( 
            @{ Label = "Mailbox"; Expression = { $mailbox.UserPrincipalName } }, 
            @{ Label = "Delegate"; Expression = { $_.Trustee } }, 
            @{ Label = "AccessRights"; Expression = { Convert-ListToText $_.AccessRights } }, 
            "IsInherited", 
            "InheritanceType" ) |
        Export-Csv -Path "$PSScriptRoot\$($mailbox.DisplayName)_Mailbox_Delegates_$timeStamp.csv" -Append -Force -NoTypeInformation
}

function Convert-ListToText($list)
{
    if ( ($null -eq $list) -or ($list.Count -eq 0) ) { return }

    if ($list.Count -eq 1)
    {
        return $list[0].ToString()
    }

    $text = ""
    for ([int]$i = 0; $i -lt $list.Count; $i++)
    {
        $text += $list[$i].ToString()
        if ($i -lt ($list.Count - 1) ) { $text += ", " }
    }
    return $text
}

# main
Show-Introduction
TryConnect-ExchangeOnline
$mailbox = PromptFor-Mailbox
Export-MailboxDelegates $mailbox
Write-Host "All done!" -ForegroundColor "Green"
Read-Host -Prompt "Press Enter to exit"

<#
Option A
One-by-one > Export > Map > Import

Option B
Export all > Map all > Import all

Ensure that when importing delegates it works with both UPN and primarySMTP (even if they're different)
#>