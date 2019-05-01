#TODO: make .PARAMETERS more consistent and add lines for function used in and default value

<#
    .SYNOPSIS
    configure_exchange

    Jacob Brown
    jmb7438@rit.edu

    .DESCRIPTION
    This script is meant to set up the base functionality of an exchange server. It will do the following:
    1) Connect to your exchange server
    2) Add a send connector
    3) Add an email address policy
    4) Create firewall rules
    5) Create mailboxes from AD users
    6) Create mailboxes from a CSV file

    .NOTES
    Script tested on Exchange Server 2016. 

    .PARAMETER ADOrganizationalUnit
    Used when adding mailboxes from Active Directory in Add-MailboxFromAD. Mailboxes for every user in the organizational unit will be made

    .PARAMETER ConnectorDomain
    Specifies the domain names to which the send connector routes mail. Used in function Add-SendConnector as the -AddressSpaces argument

    .PARAMETER ConnectorName
    Specifies the name of the send connector. Used in function Add-SendConnector as the -Name argument

    .PARAMETER ConnectorTransportServer
    Specifies the names of the Mailbox servers that can use this send connector. Used in function Add-SendConnector as the -SourceTransportServers argument

    .PARAMETER EmailPolicyIdentity
    Specifies the email address policy that you want to modify. Used in function Add-EmailAddressPolicy as the -Identity argument

    # TODO: add link to template options here
    .PARAMETER EmailPolicyTemplate
    Specifies the rules in the email address policy that are used to generate email addresses for recipients. Used in function Add-EmailAddressPolicy as the -EnabledEmailAddressTemplates argument

    .PARAMETER Password
    Specifies password to be used when connecting to the exchange server. Used in function Connect-ExchangeServer. NOTE: password must be of System.Security.SecureString type

    .PARAMETER PathToCSV
    Instructs the script to add mailboxes based on information in a CSV file and gives a path to the CSV file. Used in function Add-MailboxFromCsv. NOTE: The CSV file must of the following format:
    ###FORMAT###
    Name, LName, FName, Alias, Password, UPN
    ###END FORMAT###

    ###EXAMPLE FILE###
    Name, LName, FName, Alias, Password, UPN
    Joe Johnson, Johnson, Joe, Joe, Password1, joe.Johnson@capstone
    ###END EXAMPLE FILE###

    .PARAMETER ServerURI
    Specifies the URI of the exchange server. Used in function Connect-ExchangeServer as the -ConnectionUri argument. NOTE: format must be either one of the following
    ###FORMAT 1###
    http://<FQDN>/PowerShell
    ###END FORMAT 1###

    ###FORMAT 2###
    https://<FQDN>/PowerShell
    ###END FORMAT 2###

    .PARAMETER UseAD
    Specifies whether to add users from Active Directory or not. If set, ADOrganizationalUnit must be set as well. Used in Add-MailboxFromAD

    .PARAMETER Username
    Specifies username used to connect to the exchange server. Used in Connect-ExchangeServer

    .EXAMPLE
    $password = ConvertTo-SecureString 'Password1' -AsPlainText -Force
    ./Configure-Exchange $password

    .EXAMPLE
    $password = ConvertTo-SecureString 'Password1' -AsPlainText -Force
    ./Configure-Exchange $password -UseAD -ADOrganizationalUnit "Exchange Users"

    .EXAMPLE
    $password = ConvertTo-SecureString 'Password1' -AsPlainText -Force
    ./Configure-Exchange $password -UseAD -ADOrganizationalUnit "Exchange Users" -PathToCSV ./ADUSers.csv

#>

[CmdletBinding(DefaultParametersetName='None')]
    param (
        [Parameter(ParameterSetName="UseAD",Mandatory=$true)]
        [string]
        $ADOrganizationalUnit,

        # TODO: Update this regex to allow for more than one entry
        [ValidatePattern('^[\S+]+\.[\S]+$')]
        [string[]]
        $ConnectorDomain = $env:USERDNSDOMAIN,

        [string]
        $ConnectorName = "SMTP Mail Send",

        # TODO: Update this regex to allow for more than one entry
        [ValidatePattern('^\S+$')]
        [string[]]
        $ConnectorTransportServer = $env:COMPUTERNAME,

        [string]
        $EmailPolicyIdentity = "Default Policy",

        [string]
        $EmailPolicyTemplate = "SMTP:%g.%s@" + $env:USERDNSDOMAIN,

        [Parameter(Mandatory=$true,Position=1)]
        [securestring]
        $Password,

        [ValidateScript({Test-Path $_})]
        [string]
        $PathToCSV,

        [ValidatePattern('^https?://[\S+]+\.[\S+]+\.[\S]+/PowerShell$', Options='None')]
        [string]
        $ServerURI = "http://" + $env:COMPUTERNAME + "." + $env:USERDNSDOMAIN + "/PowerShell",

        [Parameter(ParameterSetName='UseAD')]
        [switch]
        $UseAD,

        # TODO: update regex and default value to accomodate capstone\Administrator form
        [ValidatePattern('^\S+$')]
        [string]
        $Username = $env:USERNAME
    )

process {
    function Connect-ExchangeServer {
        param (
        [Parameter(Mandatory=$true,Position=1)]
        [securestring]
        $Password,

        [ValidatePattern('^https?://[\S+]+\.[\S+]+\.[\S]+/PowerShell$', Options='None')]
        [string]
        $ServerURI = "http://" + $env:COMPUTERNAME + "." + $env:USERDNSDOMAIN + "/PowerShell",

        # TODO: copy updates to above script params
        [ValidatePattern('^\S+$')]
        [string]
        $Username = $env:USERNAME
    )
        $username = "capstone\administrator"
        $password = ConvertTo-SecureString 'Password1' -AsPlainText -Force
        $credentials = New-Object System.Management.Automation.PSCredential -ArgumentList $username, $password
        $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://WIN-90T98I75S6C.capstone.net/PowerShell -Credential $credentials
        Import-PSSession $session -AllowClobber
    }

    function Add-SendConnector {
        # Configure a Send Connector that connects to the internet
        New-SendConnector -Name "SMTP Mail Send" -Internet -AddressSpaces capstone.net -SourceTransportServers WIN-90T98I75S6C
    }

    function Add-EmailAddressPolicy {
        #Set Default Email Policy
        Set-EmailAddressPolicy -Identity "Default Policy" -EnabledEmailAddressTemplates "SMTP:%g.%s@capstone.net"
        Update-EmailAddressPolicy -Identity "Default Policy"   
    }

    function Update-FirewallRules {
        New-NetFirewallRule -DisplayName "SMTP UDP In" -Direction Inbound -Action Allow -Protocol UDP -LocalPort 25
        New-NetFirewallRule -DisplayName "SMTP TCP In" -Direction Inbound -Action Allow -Protocol TCP -LocalPort 25
        New-NetFirewallRule -DisplayName "SMTP UDP Out" -Direction Outbound -Action Allow -Protocol UDP -LocalPort 25
        New-NetFirewallRule -DisplayName "SMTP TCP Out" -Direction Outbound -Action Allow -Protocol TCP -LocalPort 25
    }

    function Add-MailboxFromAD {
        $existing_users = Get-Mailbox | Select-Object Name
        $target_users = Get-ADUser -filter * -SearchBase "ou=Exchange Users,dc=capstone,dc=net" | Select-Object Name, ObjectGUID
        foreach ($user in $target_users){
            If (-Not($existing_users.Name -contains $user.Name)) {
                $ObjectGUID = $user.ObjectGUID.ToString()
                Enable-Mailbox -Identity $ObjectGUID
            }
        }
    }

    function Add-MailboxFromCsv {
        $existing_users = Get-Mailbox | Select-Object Name
        $target_users = Import-Csv ./users.csv
        foreach($user in $target_users) {
            If (-Not($existing_users.Name -contains $user.Name)){
                $secure_password = ConvertTo-SecureString $user.password -asplaintext -force
                New-Mailbox -Name $user.Name -LastName $user.LName -FirstName $user.FName -Alias $user.Alias -Password $secure_password -UserPrincipalName $user.UPN
            }
        }
    }

    function Remove-AllExchangeUsersAD {
        $ad_users = Get-ADUser -filter * -SearchBase "ou=Exchange Users,dc=capstone,dc=net" | Select-Object Name, ObjectGUID
        foreach ($ad_user in $ad_users) {
            Remove-ADUser -Identity $ad_user.ObjectGUID -Confirm:$false
        }

    }

    function Remove-AllMailboxes {
        $mailboxes = Get-Mailbox | Select-Object Name
        foreach ($mailbox in $mailboxes) {
            If (-Not($mailbox.Name -like "Administrator" -or $mailbox.Name -like "DiscoverySearchMailbox {D919BA05-46A6-415f-80AD-7E09334BB852}")){
                Remove-Mailbox -Identity $mailbox.Name -Confirm:$false
            }
            
        }
    }

    ########################################
    # MAIN
    ########################################

    Write-Output "Username: $Username"
    Write-Output "Password: $Password"
    Write-Output "ServerURI: $ServerURI"
    Write-Output "ConnectorName: $ConnectorName"
    Write-Output "ConnectorDomain: $ConnectorDomain"
    Write-Output "ConnectorTransportServer: $ConnectorTransportServer"
    Write-Output "EmailPolicyIdentity: $EmailPolicyIdentity"
    Write-Output "EmailPolicyTemplate: $EmailPolicyTemplate"
    Write-Output "PathToCSV: $PathToCSV"
    Write-Output "UseAD: $UseAD"
    Write-Output "ADOrganizationalUnit: $ADOrganizationalUnit"


    # Connect-ExchangeServer

    # Add-SendConnector

    # Add-EmailAddressPolicy

    # Update-FirewallRules

    # Add-MailboxFromAD

    # Add-MailboxFromCsv
} # End Process