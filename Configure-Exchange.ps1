<#
    .DESCRIPTION
    This script can install Exchange 2013/2016/2019 Preview prerequisites, optionally create the Exchange
    organization (prepares Active Directory) and installs Exchange Server. When the AutoPilot switch is
    specified, it will do all the required rebooting and automatic logging on using provided credentials.
    To keep track of provided parameters and state, it uses an XML file; if this file is
    present, this information will be used to resume the process. Note that you can use a central
    location for Install (UNC path with proper permissions) to re-use additional downloads.

    .NOTES
    Requirements:
    - Operating Systems
        - Windows Server 2008 R2 SP1
        - Windows Server 2012
        - Windows Server 2012 R2
        - Windows Server 2016 (Exchange 2016 CU3+ only)
        - Windows Server 2019 (Desktop or Core, for Exchange 2019)
    - Domain-joined system (Except for Edge)
    - "AutoPilot" mode requires account with elevated administrator privileges
    - When you let the script prepare AD, the account needs proper permissions.

    .PARAMETER Phase
    Internal Use Only :)

    .EXAMPLE
    $Cred=Get-Credential
    .\Install-Exchange15.ps1 -Organization Fabrikam -InstallMailbox -MDBDBPath C:\MailboxData\MDB1\DB -MDBLogPath C:\MailboxData\MDB1\Log -MDBName MDB1 -InstallPath C:\Install -AutoPilot -Credentials $Cred -SourcePath '\\server\share\Exchange 2013\mu_exchange_server_2013_x64_dvd_1112105' -SCP https://autodiscover.fabrikam.com/autodiscover/autodiscover.xml -Verbose

#>

[CmdletBinding(DefaultParametersetName='None')]
    param (
        [ValidatePattern('^\S+$')]
        [string]
        $Username = $env:UserName,

        [Parameter(Mandatory=$true,Position=1)]
        [securestring]
        $Password,

        [ValidatePattern('^http://[\S+]+\.[\S]+/PowerShell$')]
        [string]
        $ServerURI = "http://" + $env:COMPUTERNAME + "." + $env:USERDNSDOMAIN + "/PowerShell",

        [string]
        $ConnectorName = "SMTP Mail Send",

        [ValidatePattern('^[\S+]+\.[\S]+$')]
        [string]
        $ConnectorDomain = $env:USERDNSDOMAIN,

        [ValidatePattern('^\S+$')]
        $ConnectorTransportServer = $env:COMPUTERNAME,

        [string]
        $EmailPolicyIdentity = "Default Policy",

        [string]
        $EmailPolicyTemplate = "SMTP:%g.%s@" + $env:USERDNSDOMAIN,

        $PathToCSV,

        [Parameter(ParameterSetName='UseAD')][switch]
        $UseAD,

        [string]
        [Parameter(ParameterSetName="UseAD",Mandatory=$true)]
        $ADOU
    )

process {
    function Connect-ExchangeServer {
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

    Write-Output "You specified $Username"

    # Connect-ExchangeServer

    # Add-SendConnector

    # Add-EmailAddressPolicy

    # Update-FirewallRules

    # Add-MailboxFromAD

    # Add-MailboxFromCsv
} # End Process