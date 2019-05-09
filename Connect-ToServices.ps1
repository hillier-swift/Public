<#
.SYNOPSIS
Module to connnect Microsoft services or onPrems infrastructure

.DESCRIPTION

.PARAMETER ExchangeOnline
Switch to enable connections to Exchange Online

.INPUTS
Inputs are the script parameters.

.OUTPUTS
No outputs if successful.
Outputs will be if there are errors, which are logged appropriately.

.NOTES
Version:           1.0
Author:            Tony Hillier-Swift
Creation Date:     19/04/2019
Purpose:           To provide a easy way to connect to a range of services.
Changes:           None - first release.

.EXAMPLE
Connect-ToServices -ExchangeOnline -OnPremExchange
This will connect to both onPremise Exchange and Exchanage Online

Connect to;
    1. Office 365
    2. On Prem Exchange
    3. Exchange Online
    4. On Prem Lync

.Prerequirements

Exchange Online - A guide to setting this up with MFA can be found here https://hillier-swift.co.uk/connecting-to-exchange-online-with-mfa/
Connect-MSolService Install https://www.microsoft.com/en-us/download/details.aspx?id=41950 then Install-Module MSOnline https://docs.microsoft.com/en-us/office365/enterprise/powershell/connect-to-office-365-powershell#connect-with-the-microsoft-azure-active-directory-module-for-windows-powershell 
Azure-AD - "Install-Module -Name AzureAD"
Skype for Business Onlune Module  requires Microsoft Visual C++ 2017 x64 Minimum runtime 14.10.25008 https://aka.ms/vs/15/release/vc_redist.x64.exe
    Skype for Business Online Powershell Module - https://www.microsoft.com/en-us/download/details.aspx?id=39366
Teams - Install-Module -Name MicrosoftTeams

################
##    BUGS    ##
################
Azure AD Doesn't throw a termeating error

#>

Param (

[switch]$ExchangeOnline,
[switch]$MSOL,
[switch]$AzureAD,
[switch]$OnPremExchange,
[switch]$OnPremLync,
[switch]$SkypeOnline,
[switch]$TeamsOnline
)

# Var range for customation
function Get-Customs {

    $script:Emaildomain = "@YourDomain.com"
    $script:EXOLModule = "C:\PathtoModule\EXOL\"
    # Prefix for On Prem Exchange to allow for simultaneous Exchange online and onPremise connections, Exchange online Prefix is handed by the Exchange Online Module
    $script:ExchangeOnPremPrefix = "onPrem"
    # URI to be used for On Premise Exchange connections
    $script:ExchangeonPremURI = "http://ServerName.Domain.local/PowerShell/"
    # URI to be used for On Premise Lync connections
    $script:LynconPremURI = "https://ServerName.Domain.local/ocspowershell"
    # Prefix for On Prem Lync to allow for simultaneous Lync on Premise and Lync online.
    $script:LyncPrefix = "Lync"
    # URI for Skype for Business Admin Address
    $script:SfBOnlineAdminOverride = "Tennent.onmicrosoft.com"
    # Prefix for Skype for Business Online to allow for simultaneous connections.
    $script:SfBOnlinePrefix = "cloud"

    [switch]$script:CustomsLoaded = $true

}

function Get-Creds {
    param (
    [switch]$UPN,
    [switch]$onPrem,
    [string]$Username = $env:USERNAME
    )
        #Build Vars needed for connection
        if ($upn) 
        {
            $script:UserPrincipalName = $Username + $script:Emaildomain
            $answer = Read-Host -Prompt "Is $UserPrincipalName your email to be used for this connection [y/n]"
                if ($answer -notmatch "[yY]")
                { 
                    $script:UserPrincipalName = Read-Host "Please provide the email address used for this connection"
                    Write-Host $UserPrincipalName
                    
                }  
            Write-Host "$UserPrincipalName is being used as the email address for these connections" -ForegroundColor Green
            [switch]$script:UPNSet = $true
        }
       if($onPrem)
        {
            $script:OnPremCreds = Get-Credential -Message "Please enter your on Premise credentials"
            [switch]$script:OnPremCredsSet = $true
        }

}
# Exchange Online Connection
function Connect-ExchangeOnline 
{
    if ($UPNSet) 
    {
        try 
            {
                Write-Host "Creating Exchange Online Connection - Have your MFA device ready" -ForegroundColor green
                Import-Module $EXOLModule\CreateExoPSSession.ps1 -Force
                Connect-EXOPSSession -UserPrincipalName $UserPrincipalName    
            }
        catch
            {
                $EventMessage = "$(Get-Date) - $(split-path $PScommandPath -leaf) - Could not Connect to EXOL Pausing connection" + ". " + $_.Exception.Message
                Write-Host $EventMessage
                pause
            }  
    }
        else 
            {
                Get-Creds -UPN
                Connect-ExchangeOnline
            }
}

function Connect-MSOL {
    try 
        {
            Write-Host -ForegroundColor Yellow "$(Get-Date) - $(split-path $PScommandPath -leaf) - MFA passthrough not supported sorry you have to type your email in."
            Connect-MSolService 
        }
    catch
        {
            $EventMessage = "$(Get-Date) - $(split-path $PScommandPath -leaf) - Could not Connect to Mircosoft Online Pausing connection" + ". " + $_.Exception.Message
            Write-Host $EventMessage
            pause
        }  
}

function Connect-AzureActiveDirectory {
    if ($UPNSet) 
    {
        try
            {
                Write-Host -ForegroundColor Yellow "Currently Password passthrough not supported for MFA"
                Connect-AzureAD -AccountId $UserPrincipalName 
            }
        catch
            {
                $EventMessage = "$(Get-Date) -$(split-path $PScommandPath -leaf) - Could not Connect to Azure AD Pausing connection" + ". " + $_.Exception.Message
                Write-Host $EventMessage
                Pause
            }  
    }
        else 
            {
                Get-Creds -UPN
                Connect-AzureActiveDirectory
            }
}

# On Prem Exchange
function Connect-OnPremExchange
{
    if ($OnPremCredsSet)
    {
        try 
            {
                Write-Host "Creating On-Prem Exchange Connection" -ForegroundColor green
                $OnPremExchange = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $script:ExchangeonPremURI -Credential $script:OnPremCreds
                Import-PSSession $OnPremExchange -Prefix $script:ExchangeOnPremPrefix   
            }
        catch
            {
                $EventMessage = "$(Get-Date) - $(split-path $PScommandPath -leaf) - Could not Connect to On-Prem Exchange Pausing connection" + ". " + $_.Exception.Message
                Write-Host $EventMessage
                Pause
            }
        }

    else {
            Get-Creds -onPrem
            Connect-OnPremExchange 
        }
}

#Lync Server
function Connect-OnPremLync
{
    if ($OnPremCredsSet)
    {
        try 
            {
                Write-Host "Creating Lync on Prem Connection" -ForegroundColor green
                $LyncServer = New-PSSession -ConnectionUri $script:LynconPremURI -Credential $OnPremCreds -Authentication Negotiate
                Import-PSSession $LyncServer -Prefix $script:LyncPrefix   
            }
        catch
            {
                $EventMessage = "$(Get-Date) - $(split-path $PScommandPath -leaf) - Could not Connect to $script:LynconPremURI Pausing connection" + ". " + $_.Exception.Message
                Write-Host $EventMessage
                Pause
            }
    }
    else {
        Get-Creds -onPrem
        Connect-OnPremLync
    }
}

function Connect-SkypeOnline
{
    if($UPNSet)
    {
        try 
            {
                Write-Host "Creating Skype for Business Online Connection" -ForegroundColor green
                Import-Module SkypeOnlineConnector
                $CSSession = New-CsOnlineSession -OverrideAdminDomain $script:SfBOnlineAdminOverride -UserName $UserPrincipalName -verbose
                Import-PSSession $CSSession -AllowClobber -prefix $script:SfBOnlinePrefix 
            }
        catch
            {
                $EventMessage = "$(Get-Date) - $(split-path $PScommandPath -leaf) - Could not Connect to Skype Online Pausing connection" + ". " + $_.Exception.Message
                Write-Host $EventMessage
                Pause
            }
    }
    else {
        Get-Creds -UPN
        Connect-SkypeOnline
    }
}

function Connect-Teams
{
    if($UPNSet)
    {
        try 
            {
                Write-Host "Creating Microsoft Teams Connection" -ForegroundColor green
                Connect-MicrosoftTeams -AccountId $upn
            }
        catch
            {
                $EventMessage = "$(Get-Date) - $(split-path $PScommandPath -leaf) - Could not Connect to Teams Online Pausing connection" + ". " + $_.Exception.Message
                Write-Host $EventMessage
                Pause
            }
        }
    else {
        Get-Creds -UPN
        Connect-Teams
    }
}

#region Connections

if ($ExchangeOnline)
    {
        $Service = "Exchange Online"
        If(!$script:CustomsLoaded){Get-Customs}
        try {
            Connect-ExchangeOnline
            Write-Verbose "$(Get-Date) - $(split-path $PScommandPath -leaf) Connected to $Service with $UserPrincipalName"
            $EventMessage = "$(Get-Date) - $(split-path $PScommandPath -leaf) - Connected to $Service"
            $script:ConnectedServices += $EventMessage + "`r`n"
        }
        catch {
            $EventMessage ="$(Get-Date) - $(split-path $PScommandPath -leaf) - something went wrong connecting to $service. ". " + $_.Exception.Message"
            Write-Host $EventMessage
        }
        
    }

if ($MSOL)
    {
        $Service = "MSOL"
        try {
            Connect-MSOL
            Write-Verbose "$(Get-Date) - $(split-path $PScommandPath -leaf) Connected to $Service"
            $EventMessage = "$(Get-Date) - $(split-path $PScommandPath -leaf) - Connected to $Service"
            $script:ConnectedServices += $EventMessage + "`r`n"
        }
        catch {
            $EventMessage ="$(Get-Date) - $(split-path $PScommandPath -leaf) - something went wrong connecting to $Service. ". " + $_.Exception.Message"
            Write-Host $EventMessage
        }
        
    }

if ($AzureAD)
    {
        If(!$script:CustomsLoaded){Get-Customs}
        $Service = "Azure Active Directory"
        try {
            Connect-AzureActiveDirectory
            Write-Verbose "$(Get-Date) - $(split-path $PScommandPath -leaf) Connected to $Service"
            $EventMessage = "$(Get-Date) - $(split-path $PScommandPath -leaf) - Connected to $Service"
            $script:ConnectedServices += $EventMessage + "`r`n"
        }
        catch {
            $EventMessage ="$(Get-Date) - $(split-path $PScommandPath -leaf) - something went wrong connecting to $Service. ". " + $_.Exception.Message"
            Write-Host $EventMessage
        }
        
    }

If($OnPremExchange)
    {
        If(!$script:CustomsLoaded){Get-Customs}
        $Service = "On Premise Exchange"
        try {
            Connect-OnPremExchange
            Write-Verbose "$(Get-Date) - $(split-path $PScommandPath -leaf) Connected to $Service"
            $EventMessage = "$(Get-Date) - $(split-path $PScommandPath -leaf) - Connected to $Service"
            $script:ConnectedServices += $EventMessage + "`r`n"
        }
        catch {
            $EventMessage ="$(Get-Date) - $(split-path $PScommandPath -leaf) - something went wrong connecting to $Service. ". " + $_.Exception.Message"
            Write-Host $EventMessage
        }
    }

If($OnPremLync)
    {
        If(!$script:CustomsLoaded){Get-Customs}
        $Service = "On Premise Lync"
        try {
            Connect-OnPremLync
            Write-Verbose "$(Get-Date) - $(split-path $PScommandPath -leaf) Connected to $Service"
            $EventMessage = "$(Get-Date) - $(split-path $PScommandPath -leaf) - Connected to $Service"
            $script:ConnectedServices += $EventMessage + "`r`n"
        }
        catch {
            $EventMessage ="$(Get-Date) - $(split-path $PScommandPath -leaf) - something went wrong connecting to $Service. ". " + $_.Exception.Message"
            Write-Host $EventMessage
        }
    }

If($SkypeOnline)
{
    If(!$script:CustomsLoaded){Get-Customs}
    $Service = "Skype for Buiness Online"
    try {
        Connect-SkypeOnline
        Write-Verbose "$(Get-Date) - $(split-path $PScommandPath -leaf) Connected to $Service"
        $EventMessage = "$(Get-Date) - $(split-path $PScommandPath -leaf) - Connected to $Service"
        $script:ConnectedServices += $EventMessage + "`r`n"
    }
    catch {
        $EventMessage ="$(Get-Date) - $(split-path $PScommandPath -leaf) - something went wrong connecting to $Service. ". " + $_.Exception.Message"
        Write-Host $EventMessage
    }
}

if($TeamsOnline)
{
    $Service = "Microsoft Teams"
    try {
        Connect-Teams
        Write-Verbose "$(Get-Date) - $(split-path $PScommandPath -leaf) Connected to $Service"
        $EventMessage = "$(Get-Date) - $(split-path $PScommandPath -leaf) - Connected to $Service"
        $script:ConnectedServices += $EventMessage + "`r`n"
    }
    catch {
        $EventMessage ="$(Get-Date) - $(split-path $PScommandPath -leaf) - something went wrong connecting to $Service. ". " + $_.Exception.Message"
        Write-Host $EventMessage
    }   
}

#endregion
Write-Host "Services Connected" -ForegroundColor Green
Write-Host $script:ConnectedServices
