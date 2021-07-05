################################################################################
# file: MailboxAddFullAccessRights.ps1
# path: C:\Windows\System32\WindowsPowerShell\v1.0\
# auth: https://github.com/andreoidb64
# date: 01.12.2016
#
# Copyright(C) 2016 Bobst Italia S.p.A.
# S.P. Casale-Asti n.70
# 15020 San Giorgio Monferrato (AL), ITALY
# Tel  +39 0142 4071
# Fax  +39 0142 460954
################################################################################
# Description:
#
# Add mailbox FullAccess permission
#
################################################################################
#
# ... Set main environment
#
$ConnectionUri = "https://exchange.mydomain.local/powershell/"

Add-Type -AssemblyName Microsoft.VisualBasic

################################################################################
#
# ... Functions
#

function Read-InputBoxDialog([string]$Message, [string]$WindowTitle, [string]$DefaultText)
{
    return [Microsoft.VisualBasic.Interaction]::InputBox($Message, $WindowTitle, $DefaultText)
}

function Read-MsgBoxDialog([string]$Message, [string]$WindowTitle, [int]$MessageType)
{
    ### Message type codes:
    # OKOnly                  0 Displays OK button only.
    # OKCancel                1 Displays OK and Cancel buttons.
    # AbortRetryIgnore        2 Displays Abort, Retry, and Ignore buttons.
    # YesNoCancel             3 Displays Yes, No, and Cancel buttons.
    # YesNo                   4 Displays Yes and No buttons.
    # RetryCancel             5 Displays Retry and Cancel buttons.
    # Critical               16 Displays Critical Message icon.
    # Question               32 Displays Warning Query icon.
    # Exclamation            48 Displays Warning Message icon.
    # Information            64 Displays Information Message icon.
    # DefaultButton1          0 First button is default.
    # DefaultButton2        256 Second button is default.
    # DefaultButton3        512 Third button is default.
    # ApplicationModal        0 Application is modal. The user must respond to the message box before continuing work in the current application.
    # SystemModal          4096 System is modal. All applications are suspended until the user responds to the message box.
    # MsgBoxSetForeground 65536 Specifies the message box window as the foreground window.
    # MsgBoxRight        524288 Text is right-aligned.
    # MsgBoxRtlReading  1048576 Specifies text should appear as right-to-left reading on Hebrew and Arabic systems.
    return [Microsoft.VisualBasic.Interaction]::MsgBox("$Message", $MessageType, "$WindowTitle")
}

function CreateExchangeRemoteSession()
{
    $a_myaccount = Get-Credential
    $SessionOptions = New-PSSessionOption -SkipCNCheck
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ConnectionUri -Authentication basic -Credential $a_myaccount -SessionOption $SessionOptions
    Import-PSSession $session
}

function CloseExchangeRemoteSession($session)
{
    Get-PSSession | Remove-PSSession
}

################################################################################
#
# ... Main
#

# ... Create Exchange remote session
CreateExchangeRemoteSession

# ... Get user mailbox
$UserLogonName = ""
while ("$UserLogonName" -eq "")
{
    $UserLogonName = Read-InputBoxDialog -Message "Please enter the User Logon Name`n(Leave 'quit' to abort)" -WindowTitle "Add mailbox FullAccess permission" -DefaultText "quit"
    if ($UserLogonName -eq "quit") {
        CloseExchangeRemoteSession
        exit
    }
    Get-Mailbox "$UserLogonName"
    if ($? -eq 0) {$UserLogonName = ""}
}    

# ... Get allowed Group
$GroupNameAllowed = ""
while ("$GroupNameAllowed" -eq "")
{
    $GroupNameAllowed = Read-InputBoxDialog -Message "Please enter the Group Name to be allowed`n(Write 'quit' to abort)" -WindowTitle "Add mailbox FullAccess permission" -DefaultText "GRP_$UserLogonName"
    if ($GroupNameAllowed -eq "quit") {
        CloseExchangeRemoteSession
        exit
    }
}    

#Get-MailboxPermission "$UserLogonName"
Add-MailboxPermission -Identity "$UserLogonName" -User "$GroupNameAllowed" -AccessRights FullAccess -InheritanceType All -AutoMapping $false

Read-MsgBoxDialog -Message "Command completed.`nPlease check console window for success/errors logs." -WindowTitle "Set mailbox FullAccess" -MessageType 4096

CloseExchangeRemoteSession
exit
