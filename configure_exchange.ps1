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
     - Script tested on Exchange Server 2016. 
     - If you wish to just use functions within this script without running the whole script, comment out the main
       function code section at the bottom and dot source the file. The you will be able to use these functions as powershell
       functions from the commandline.

    .PARAMETER ADOrganizationalUnit
    Description: Mailboxes for every user in this organizational unit will be made.
    Mandatory: False
    Default Value: None
    Function Found In: Add-MailboxFromAD

    .PARAMETER ConnectorDomain
    Description: Specifies the domain name to which the send connector routes mail.
    Mandatory: False
    Default Value: $env:USERDNSDOMAIN
    Function Found In: Add-SendConnector

    .PARAMETER ConnectorName
    Description: Specifies the name of the send connector.
    Mandatory: False
    Default Value: "SMTP Mail Send"
    Function Found In: Add-SendConnector

    .PARAMETER ConnectorTransportServer
    Description: Specifies the name of the Mailbox servers that can use this send connector.
    Mandatory: False
    Default Value: $env:COMPUTERNAME
    Function Found In: Add-SendConnector

    .PARAMETER DCDomain
    Description: Used to specify which domain to pull users from. NOTE: should be in format of "something.something". Example: domain.com
    Mandatory: False
    Default Value: $env:USERDNSDOMAIN 
    Function Found In: Add-MailboxFromAD

    .PARAMETER EmailPolicyIdentity
    Description: Specifies the email address policy that you want to modify.
    Mandatory: False
    Default Value: Defalt Policy
    Function Found In: Add-EmailAddressPolicy

    .PARAMETER EmailPolicyTemplate
    Description: Specifies the rules in the email address policy that are used to generate email addresses for recipients.
    Mandatory: False
    Defualt Value: "SMTP:%g.%s@" + $env:USERDNSDOMAIN
    Function Found In: Add-EmailAddressPolicy

    .PARAMETER Password
    Description: Specifies password to be used when connecting to the exchange server. NOTE: password must be of type System.Security.SecureString
    Mandatory: True
    Default Value: None
    Fuction Used In: Connect-ExchangeServer

    .PARAMETER PathToCSV
    Instructs the script to add mailboxes based on information in a CSV file and gives a path to the CSV file.
    Mandatory: False
    Default Value: None
    Function Found In: Add-MailboxFromCsv

    .PARAMETER ServerURI
    Description: Specifies the URI of the exchange server. NOTE: format must be either one of the following:
    ###FORMAT 1###
    http://<FQDN>/PowerShell
    ###END FORMAT 1###

    ###FORMAT 2###
    https://<FQDN>/PowerShell
    ###END FORMAT 2###

    Mandatory: False
    Default Value: http://$env:COMPUTERNAME.$env:USERDNSDOMAIN/PowerShell
        NOTE: This default value is designed for running the script locally on the exchange server
    Function Found In: Connect-ExchangeServer

    .PARAMETER UseAD
    Specifies whether to add users from Active Directory or not. If set, ADOrganizationalUnit must be set as well. 
    Mandatory: False
    Default Value: N/A
    Function Used In: Add-MailboxFromAD

    .PARAMETER Username
    Description: Specifies username used to connect to the exchange server.
    Mandatory: False
    Default Value: $env:USERNAME
    Function Found In: Connect-ExchangeServer

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

        [ValidatePattern('^[\S+]+\.[\S]+$')]
        [string[]]
        $ConnectorDomain = $env:USERDNSDOMAIN,

        [string]
        $ConnectorName = "SMTP Mail Send",

        [ValidatePattern('^\S+$')]
        [string[]]
        $ConnectorTransportServer = $env:COMPUTERNAME,

        [Parameter(ParameterSetName="UseAD",Mandatory=$false)]
        [ValidatePattern('^[\S+]+\.[\S]+$')]
        [string]
        $DCDomain = $env:USERDNSDOMAIN,

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

        [ValidatePattern('^https?://[\S+]+/PowerShell$', Options='None')]
        [string]
        $ServerURI = "http://" + $env:COMPUTERNAME + "." + $env:USERDNSDOMAIN + "/PowerShell",

        [Parameter(ParameterSetName='UseAD')]
        [switch]
        $UseAD,

        [ValidatePattern('^\S+$')]
        [string]
        $Username = $env:USERNAME
    )

process {
    <#
        .DESCRIPTION
        This function connects to the exchange server. This is necessary in order to run any exchange cmdlets later in the script.

        .PARAMETER Password
        Description: Specifies password to be used when connecting to the exchange server. NOTE: password must be of type System.Security.SecureString
        Mandatory: True
        Default Value: None

        .PARAMETER ServerURI
        Description: Specifies the URI of the exchange server. NOTE: format must be either one of the following:
        ###FORMAT 1###
        http://<FQDN>/PowerShell
        ###END FORMAT 1###

        ###FORMAT 2###
        https://<FQDN>/PowerShell
        ###END FORMAT 2###

        Mandatory: False
        Default Value: http://$env:COMPUTERNAME.$env:USERDNSDOMAIN/PowerShell
        NOTE: This default value is designed for running the script locally on the exchange server

        .PARAMETER Username
        Description: Specifies username used to connect to the exchange server.
        Mandatory: False
        Default Value: $env:USERNAME
    #>
    function Connect-ExchangeServer {
        param (
            [Parameter(Mandatory=$true,Position=1)]
            [securestring]
            $Password,

            [ValidatePattern('^https?://[\S+]+/PowerShell$', Options='None')]
            [string]
            $ServerURI = "http://" + $env:COMPUTERNAME + "." + $env:USERDNSDOMAIN + "/PowerShell",

            # TODO: copy updates to above script params
            [ValidatePattern('^\S+$')]
            [string]
            $Username = $env:USERNAME
        )

        $credentials = New-Object System.Management.Automation.PSCredential -ArgumentList $Username, $Password
        $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ServerURI -Credential $credentials
        Import-PSSession $session -AllowClobber
    }

    <#
        .DESCRIPTION
        This function adds a send connector which is needed to send mail from the exchange server to the internet. By default it adds a connector named
        "SMTP Mail Send" and connects it to your current domain and uses the local server as a transport server.

        .PARAMETER ConnectorDomain
        Description: Specifies the domain name to which the send connector routes mail.
        Mantatory: False
        Default Value: $env:USERDNSDOMAIN

        .PARAMETER ConnectorName
        Description: Specifies the name of the send connector.
        Mandatory: False
        Default Value: "SMTP Mail Send"

        .PARAMETER ConnectorTransportServer
        Description: Specifies the name of the Mailbox servers that can use this send connector.
        Mandatory: False
        Default Value: $env:COMPUTERNAME
    #>
    function Add-SendConnector {
        param (
            [ValidatePattern('^[\S+]+\.[\S]+$')]
            [string[]]
            $ConnectorDomain = $env:USERDNSDOMAIN,

            [string]
            $ConnectorName = "SMTP Mail Send",

            [ValidatePattern('^\S+$')]
            [string[]]
            $ConnectorTransportServer = $env:COMPUTERNAME
        )
        $existing_connectors = Get-SendConnector | Select-Object Name

        If (-Not($existing_connectors.Name -contains $ConnectorName)) {
            # Configure a Send Connector that connects to the internet
            New-SendConnector -Name $ConnectorName -Internet -AddressSpaces $ConnectorDomain -SourceTransportServers $ConnectorTransportServer
            Write-Output "Created new send connector called $ConnectorName"
        }
    }

    <#
        .DESCRIPTION
        This function adds an email address policy. By default it will create a policy where emails follow the following pattern: firstname.lastname@$env:USERDNSDOMAIN

        .PARAMETER EmailPolicyIdentity
        Description: Specifies the email address policy that you want to modify.
        Mandatory: False
        Default Value: Defalt Policy

        .PARAMETER EmailPolicyTemplate
        Description: Specifies the rules in the email address policy that are used to generate email addresses for recipients.
        Mandatory: False
        Defualt Value: "SMTP:%g.%s@" + $env:USERDNSDOMAIN
    #>
    function Add-EmailAddressPolicy {
        param (
            [string]
            $EmailPolicyIdentity = "Default Policy",

            [string]
            $EmailPolicyTemplate = "SMTP:%g.%s@" + $env:USERDNSDOMAIN
        )

        $existing_policies = Get-EmailAddressPolicy | Select-Object Name

        If ($EmailPolicyIdentity -eq "Default Policy") {
            Set-EmailAddressPolicy -Identity "Default Policy" -EnabledEmailAddressTemplates $EmailPolicyTemplate
            Update-EmailAddressPolicy -Identity "Default Policy"
            Write-Output "Updated the default policy" 
        } else {
            If ($existing_policies.Name -contains $EmailPolicyIdentity) {
                Write-Output "Policy already exists with the following name: $EmailPolicyIdentity"
                Write-Output "Updating the policy with the values given"
                Set-EmailAddressPolicy -Identity $EmailPolicyIdentity -EnabledEmailAddressTemplates $EmailPolicyTemplate
                Update-EmailAddressPolicy -Identity $EmailPolicyIdentity
            } else {
                New-EmailAddressPolicy -Name $EmailPolicyIdentity -IncludedRecipients "MailboxUsers" -EnabledEmailAddressTemplates $EmailPolicyTemplate
                Update-EmailAddressPolicy -Identity $EmailPolicyIdentity
            }
        }
    }

    <#
        .DESCRIPTION
        This function updates the firewall rules to accomodate exchange.
    #>
    function Update-FirewallRules {
        New-NetFirewallRule -DisplayName "SMTP UDP In" -Direction Inbound -Action Allow -Protocol UDP -LocalPort 25
        New-NetFirewallRule -DisplayName "SMTP TCP In" -Direction Inbound -Action Allow -Protocol TCP -LocalPort 25
        New-NetFirewallRule -DisplayName "SMTP UDP Out" -Direction Outbound -Action Allow -Protocol UDP -LocalPort 25
        New-NetFirewallRule -DisplayName "SMTP TCP Out" -Direction Outbound -Action Allow -Protocol TCP -LocalPort 25
    }

    <#
        .DESCRIPTION
        This function adds mailboxes from a specific organizational unit in active directory if the -UseAD flag is set.

        .PARAMETER ADOrganizationalUnit
        Description: Mailboxes for every user in this organizational unit will be made.
        Mandatory: True
        Default Value: None

        .PARAMETER DCDomain
        Description: Used to specify which domain to pull users from. NOTE: should be in format of "something.something". Example: domain.com
        Mandatory: False
        Default Value: $env:USERDNSDOMAIN
    #>
    function Add-MailboxFromAD {
        param (
            [Parameter(Mandatory=$true)]
            [string]
            $ADOrganizationalUnit,
            
            [ValidatePattern('^[\S+]+\.[\S]+$')]
            [string]
            $DCDomain = $env:USERDNSDOMAIN
        )

        $dc_company_component = $DCDomain.Split(".")[0]
        $dc_tld_component = $DCDomain.Split(".")[1]

        $existing_users = Get-Mailbox | Select-Object Name
        $target_users = Get-ADUser -filter * -SearchBase "ou=$ADOrganizationalUnit,dc=$dc_company_component,dc=$dc_tld_component" | Select-Object Name, ObjectGUID
        foreach ($user in $target_users){
            If (-Not($existing_users.Name -contains $user.Name)) {
                $ObjectGUID = $user.ObjectGUID.ToString()
                Enable-Mailbox -Identity $ObjectGUID
            }
        }
    }

    <#
        .DESCRIPTION
        This function adds mailboxes from a CSV file. NOTE: The CSV file must of the following format:
        ###FORMAT###
        Name, LName, FName, Alias, Password, UPN
        ###END FORMAT###

        Here is an example: 
        ###EXAMPLE FILE###
        Name, LName, FName, Alias, Password, UPN
        Joe Johnson, Johnson, Joe, Joe, Password1, joe.Johnson@capstone
        ###END EXAMPLE FILE###

        .PARAMETER PathToCSV
        Instructs the script to add mailboxes based on information in a CSV file and gives a path to the CSV file.
        Mandatory: True
        Default Value: None
    #>
    function Add-MailboxFromCsv {
        param (
            [Parameter(Mandatory=$true)]
            [ValidateScript({Test-Path $_})]
            [string]
            $PathToCSV
        )
        $existing_users = Get-Mailbox | Select-Object Name
        $target_users = Import-Csv ./users.csv
        foreach($user in $target_users) {
            If (-Not($existing_users.Name -contains $user.Name)){
                $secure_password = ConvertTo-SecureString $user.password -AsPlainText -Force
                New-Mailbox -Name $user.Name -LastName $user.LName -FirstName $user.FName -Alias $user.Alias -Password $secure_password -UserPrincipalName $user.UPN
            }
        }
    }

    <#
        .DESCRIPTION
        This function deletes users from a specified organizational unit. This was mainly used when testing.

        .PARAMETER ADOrganizationalUnit
        Description: Every user in this organizational unit will be deleted
        Mandatory: True
        Default Value: None

        .PARAMETER DCDomain
        Description: Domain to delete users in. NOTE: should be in format of "something.something". Example: domain.com
        Mandatory: False
        Default Value: $env:USERDNSDOMAIN
    #>
    function Remove-AllExchangeUsersAD {
        param (
            [Parameter(Mandatory=$true)]
            [string]
            $ADOrganizationalUnit,
            
            [ValidatePattern('^[\S+]+\.[\S]+$')]
            [string]
            $DCDomain = $env:USERDNSDOMAIN
        )

        $dc_company_component = $DCDomain.Split(".")[0]
        $dc_tld_component = $DCDomain.Split(".")[1]

        $ad_users = Get-ADUser -filter * -SearchBase "ou=$ADOrganizationalUnit,dc=$dc_company_component,dc=$dc_tld_component" | Select-Object Name, ObjectGUID
        foreach ($ad_user in $ad_users) {
            Remove-ADUser -Identity $ad_user.ObjectGUID -Confirm:$false
        }

    }

    <#
        .DESCRIPTION
        This function deletes all mailboxes except the ones belonging to the Administrator and DiscoverySearchMailbox. This was mainly used when testing.
    #>
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

    Connect-ExchangeServer -Password $Password -Username $Username -ServerURI $ServerURI

    Add-SendConnector -ConnectorName $ConnectorName -ConnectorDomain $ConnectorDomain -ConnectorTransportServer $ConnectorTransportServer 

    Add-EmailAddressPolicy -EmailPolicyIdentity $EmailPolicyIdentity -EmailPolicyTemplate $EmailPolicyTemplate

    Update-FirewallRules

    If ($UseAD) {
        Add-MailboxFromAD -ADOrganizationalUnit $ADOrganizationalUnit -DCDomain $DCDomain
    }

    If ($PathToCSV) {
        Add-MailboxFromCsv -PathToCSV $PathToCSV
    }

} # End Process