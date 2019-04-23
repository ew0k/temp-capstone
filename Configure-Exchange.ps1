# TODO: 
#       - parameterize the following: username and password for $credentials, ConnectionUri for $session (make it use server name by default)
#       - make it so if no username or password is specified, it jsut uses the current credentials which will work if the script is being run on the exchange server
#       - make it so there is an option to either make mailboxes from AD, a CSV file, or both

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
