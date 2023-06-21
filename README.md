# ⛔️ DEPRECATED: Exchange Connection Manager 2.0

The new release finally adds support for Powershell Core‼️ You can now use ECM on basically any platform like macOS, Linux and even iOS using Azure Cloud Shell.

## About
Are you an Exchange Admin? Are you managing multiple Exchange OnPrem and Online environments? Tired of switching between multiple shells and sessions, updating and entering credentials? Check out the PowerShell Exchange Connection Manager module.

## Note
Due to „limitations“ of the secure string CMDlets in PS Core, ECM on PS Core generates a random encryption key after installation to handle your saved credentials. If you‘re planning to save your credentials, make sure that nobody can access your key and config files inside your PS profile directory. However, if you have security concerns you can still use ECM by providing your credentials each time a connection is selected.

## Known Issues 
Currently there are no known issues. 🤗

# Installation
## PowerShell
For introductions on how to install PowerShell on your platform visit [Microsoft Docs](https://docs.microsoft.com/en-us/powershell/scripting/install/installing-powershell).

## Exchange Connection Manager
Install ECM from the PowerShell gallery.
```Powershell
Install-Module Exchangea.Connection.Manager
```
If you’re already using an older version of ECM use the update CMDlet.
```Powershell
Update-Module Exchange.Connection.Manager
```
Thanks to the module auto load feature starting with PowerShell 3.0 you can use the ECM commands right away. However, if you’re having issues try to import ECM manually.
```Powershell
Import-Module Exchange.Connection.Manager
```

# Configuration & Usage Examples 

Create an Exchange Online Connection for your contoso tenant (contoso.onmicrosoft.com).

```Powershell
New-ECMConnection -TenantName contosocom -Name EXOContoso -Credential (get-credential)
```
Create an Exchange OnPrem Connection with `OnPrem` as command suffix.
```Powershell
New-ECMOnPremConnection -Name Exchange2016 -Authentication Kerberos -URI http://mail.contoso.com/PowerShell -SessionPrefix OnPrem
```
Remove all connections and do not ask for confirmation.
```Powershell
Get-ECMConnection | Remove-ECMConnection -Confirm:$false
```
Get all configured connections.
```Powershell
Get-ECMConnection
```
Connect to Exchange Online and Exchange OnPrem in a single shell and get all mailboxes.
```Powershell
Connect-Exchange -Identity EXOContoso
Connect-Exchange -Identity Exchange2016

$allMBX = @()
$allMBX = Get-Mailbox -Resultsize unlimited
$allMBX += Get-OnPremMailbox -Resultsize unlimited 
$allMBX

```
You can even use `Connect-Exchange` in a loop since ECM won’t reconnect your session when the session is already active but will reconnect your session when the session is broken or disconnected.

```Powershell
do{
  Connect-Exchange EXOContoso
  $mbx = Get-Mailbox jdoe
  Start-Sleep -seconds 10
}until($mbx)

```

# Command Reference

Command | Parameter | Description 
------------ | ------------- | -------------
Connect-Exchange | |Start the connection menu.
||`[string]` Identity|Skip the menu and connect directly. Good option for use in scripts.
||`[switch]` CoreMenu|Start the core menu when having issues with the full menu on your platform 
|Get-ECMConnection||Get all connections
||`[string]`  Name |Get a single connection configuration by name   
|New-ECMConnection||Create a new persistent connection
||**Mandatory** `[string]` TenantName |**Important** Use the real name of your tenant. The string before your .onmicrosoft.com address. Actually that is not really required for the connection to work, but makes sure you can use your connection database in future ECM versions.
||`[string]` Name| The name of the connection. The name will be displayed in the connection menu.
||`[PSCredential]` Credential| Pass a credential object to the command e.g. with (get-credential). However, you can skip this parameter and ECM will ask you for your credential.
||`[string]` SessionPrefix|Use a prefix with your connection. e.g. If you specify `Online` as prefix the Get-Mailbox command will map to Get-`Online`Mailbox.
|New-ECMOnPremConnection||Create a connection to your local Exchange environment 
||**Mandatory** `[string]` Name| The name of the connection. The name will be displayed in the connection menu.
||**Mandatory** `[string]` URI|Specifies a URI that defines the connection endpoint for the session. The URI must be fully qualified.
||`[string]` Authentication|Specifies the mechanism that is used to authenticate the user's credentials. The acceptable values for this parameter are: Kerberos (default) & Basic.
||`[string]` SessionPrefix|Use a prefix with your connection. e.g. If you specify `OnPrem` as prefix the Get-Mailbox command will map to Get-`OnPrem`Mailbox.
|Remove-ECMConnection|| Remove a connection from the database.
||`[string]` Name| The name of the connection you want to remove.
||`[bool]` Confirm| Prompts you for confirmation before running the command. Default value is `$true`.
|Set-ECMConnection||Change connection config details.
||**Mandatory** `[string]`$Identity|The name of the connection you want to modify.
||`[string]`Name|Set the name of the connection provided by the Identity parameter.
||`[string]`TenantName|Set the TenantName.
||`[PSCredential]`Credential|Provide credentials for the connection.
||`[string]`SessionPrefix|Set a prefix for the connection. e.g. If you specify `Online` as prefix the Get-Mailbox command will map to Get-`Online`Mailbox.
||`[string]`URI|Specifies a URI that defines the connection endpoint for the session. The URI must be fully qualified.
||`[string]`Authentication|Specifies the mechanism that is used to authenticate the user's credentials. The acceptable values for this parameter are: Kerberos (default) & Basic.
|Get-ECMConfig||Show all global ECM configurations.
|Set-ECMConfig||Change global ECM configurations.
||`[string]`Menu|Specifies the default ECM menu. The acceptable values for this parameter are: Auto (Default) & Core
||`[bool]`CheckForUpdates|Check for ECM module updates once a day.
||`[bool]`EditEnumerationLimit|Prevents truncation of long output after a session is established.
||`[bool]`Logging|Log ECM activities.
||`[string]`Logfile|ECM logfile path.

# FAQ
**Can I use ECM with MFA enabled accounts?**

Unfortunately ECM has no MFA support in general. However, you can create an app password and use it with ECM.

**How to provide feedback and get support?**

Feedback is always appreciated. You can drop a line using the [contact form](https://andreasbode.net/contact/) or use the comment section below.

**Can I use my old connection file?**

No. To support PS Core I have made some changes in how ECM handles your connections. Recreate your connections with New-ECMConnection.
