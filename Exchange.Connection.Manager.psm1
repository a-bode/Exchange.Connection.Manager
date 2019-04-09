# Global Handling
function Initialize-ECM {

    #PowerShell Profile Directory
    $PSProfilePath = Split-Path $profile
    if (!(Test-path $PSProfilePath)) {
        $null = New-Item -Path $PSProfilePath -Force -ItemType  Directory
    }

    #ECM Directory
    $PSProfilePathObject = Get-Item $PSProfilePath -force
    $PSProfilePathObject = join-path -Path $PSProfilePathObject -ChildPath '.ECM'
    if (!(Test-path $PSProfilePathObject)) {
        $null = New-Item -Path $PSProfilePathObject -Force -ItemType  Directory
    }
    $PSProfilePathObject = Get-Item $PSProfilePathObject -force
    $ECMFolderPath = $PSProfilePathObject.FullName


    #Generate ECM key and file
    $EncryptionKeyPath = join-path -Path $PSProfilePathObject -ChildPath '.ECMkey'
    if (!(Test-Path $EncryptionKeyPath)) {
        $key = @()
        for ($i = 1; $i -le 24; $i++) {$key += (get-random -Maximum 255)}
        $key |ConvertTo-Json |Out-File $EncryptionKeyPath -Force
    }

    
    #ECM connection file
    $ConnectionFilePath = join-path -Path $PSProfilePathObject -ChildPath '.ECMconnect'
    if (!(Test-Path $ConnectionFilePath)) {
        $null = New-Item -Path $ConnectionFilePath -Force -ItemType File
    }

    #ECM log file
    $ECMLogFileObject = Get-Item $PSProfilePath -force
    $ECMLogFile = join-path -Path $ECMLogFileObject -ChildPath 'ECM.log'

    #ECM config file
    $ECMConfigFile = join-path -Path $PSProfilePathObject -ChildPath '.ECMconfig'
    if (!(Test-Path $ECMConfigFile)) {
        $ECMConfig = @{
            'Logging'              = $true
            'Logfile'              = $ECMLogFile
            'Menu'                 = 'Auto'
            'EditEnumerationLimit' = $true
            'CheckForUpdates'      = $true
            'LastUpdateCheck'      = (get-date).AddDays(-1)
        }
        $ECMConfig |ConvertTo-Json |Out-File -FilePath $ECMConfigFile -Force
    }


    $files = @{
        'PSProfilePath'      = $PSProfilePath;
        'EncryptionKeyPath'  = $EncryptionKeyPath;
        'ConnectionFilePath' = $ConnectionFilePath;
        'ECMConfigFile'      = $ECMConfigFile;
        'ECMFolderPath'      = $ECMFolderPath
    }
    
    Lock-ECMFiles
    return $files

}

function Lock-ECMFiles {

    foreach ($key in $global:ECMFiles.Keys) {
        if ($key -ne 'PSProfilePath') {
            $path = $global:ECMFiles.$key
            $item = get-Item -Path $path -Force
            $item.Attributes = 'Hidden'        
        }  
    }
}

function Write-ScriptLog {
    [CmdletBinding()]
    param
    (
        [string]
        $String,

        [ValidateSet('INFO', 'WARNING', 'ERROR')]
        [string]
        $Type = 'INFO',

        [switch]
        $Silent

    )
    
    function New-ScriptLog {
        
        if (!(Test-Path $ScriptLogFile )) {
            $null = New-Item -Path $ScriptLogFile -ItemType File -Force
        }
        
        [int]$CurrentLogFileSIze = (get-item $ScriptLogFile).Length / 1mb

        if ($CurrentLogFileSIze -ge 10 ) {
            $null = New-Item -Path $ScriptLogFile -ItemType File -Force
        }
    }

    #Create new log file
    $ScriptLogFile = (Get-ECMConfig).Logfile
    New-ScriptLog
    
    if ((Get-ECMConfig).Logging) {
        #Create and append log string
        (Get-Date -Format s) + '|' + $Env:USERNAME + '|' + $Type.ToUpper() + '|' + $String |Out-File $ScriptLogFile -Append
    }

    #Write to console
    if ($Silent -ne $true) {
        #Set host output color
        switch ($Type) {
            'INFO' {$OutputColor = 'green'}
            'WARNING' {$OutputColor = 'yellow'}
            'ERROR' {$OutputColor = 'red'}
            default {$OutputColor = 'green'}
        }
        Write-Host -Object ('[' + ($Type.ToUpper()) + '] ' + $String) -ForegroundColor $OutputColor
    }
}


function Unlock-ECMFiles {
    foreach ($key in $global:ECMFiles.Keys) {
        if ($key -ne 'PSProfilePath') { 
            $path = $global:ECMFiles.$key
            $item = get-Item -Path $path -Force
            $item.Attributes = 'Archive'
        }
    }    
}


function Update-ECM {
    Write-ScriptLog -String 'Checking for ECM Updates...' -Type WARNING
    $CurrentECMVersion = (Get-Module Exchange.Connection.Manager -ListAvailable).version.ToString()
    $AvailableECMVersion = (Find-Module Exchange.Connection.Manager).version
    Write-ScriptLog -String ('ECM version running: ' + $CurrentECMVersion) -Type INFO -Silent
    if ($CurrentECMVersion -lt $AvailableECMVersion) {
        Write-ScriptLog "There is an ECM Update available. Use: Update-Module Exchange.Connection.Manager" -Type WARNING
        Write-ScriptLog -String ('ECM version available: ' + $AvailableECMVersion) -Type INFO -Silent
        $ECMConfig = Get-ECMConfig
        $ECMConfig.LastUpdateCheck = get-date -Format s
        Unlock-ECMFiles
        $ECMConfig  |ConvertTo-Json |Out-File -FilePath $global:ECMFiles.ECMConfigFile -Force

    }

}


# Credential Handling
function Get-EncryptedCredential {
    param(
        [pscredential]$Credential
    )

    $UserName = $Credential.UserName
    if ($PSVersionTable.Platform -eq 'Unix') {
        $key = Get-EncryptionKey
        $Password = $Credential.Password |ConvertFrom-SecureString -Key $key
    }
    else {
        $Password = $Credential.Password |ConvertFrom-SecureString
    }

    return $cred = @{
        'Username' = $UserName;
        'Password' = $Password
    }

}

function Get-DecryptedCredential {
    param(
        [object]$Credential
    )

    $UserName = $Credential.UserName
    if ($PSVersionTable.Platform -eq 'Unix') {
        $key = Get-EncryptionKey
        $Password = $Credential.Password |ConvertTo-SecureString -Key $key
    }
    else {
        $Password = $Credential.Password |ConvertTo-SecureString
    }

    $cred = New-Object System.Management.Automation.PSCredential ($UserName, $Password)
    return $cred

}

function Get-EncryptionKey {


    $EncryptionKey = Get-Content $global:ECMFiles.EncryptionKeyPath |ConvertFrom-Json
    return $EncryptionKey

}

# Connection Handling
function Get-ConnectionFile {

    $ConnectionFile = Get-Content -Path $global:ECMFiles.ConnectionFilePath |ConvertFrom-Json
    return $ConnectionFile

}

function Show-CoreMenu {
    param(
        $Connections
    )

    Clear-Host 
    Write-Host "=======================================" 
    Write-Host "==    Exchange Connection Manager    =="
    Write-Host "======================================="  
    write-host ''

    if ($Connections) {
        foreach ($element in $Connections) {
            $i++
            Write-Host "      $i : $($element.Name)"
        }
    }
    else {
        Write-Host "      NO CONNECTION CONFIGURED" -foregroundcolor red
    }
    Write-Host ''
    Write-Host "      Q: Press 'Q' to quit."
    Write-Host '' 

}


function Show-FullMenu {
    param(
        $Connections
    )
    Function Show-Menu {
        param ($menuItems, $menuPosition, $menuTitel)
        $fcolor = $host.UI.RawUI.ForegroundColor
        $bcolor = $host.UI.RawUI.BackgroundColor
  
  
        $l = $menuItems.length + 1
        $menuwidth = $menuTitel.length + 4
        Clear-Host
        Write-Host -Object "`t" -NoNewline
        Write-Host -Object ('*' * $menuwidth) -ForegroundColor $fcolor -BackgroundColor $bcolor
        Write-Host -Object "`t" -NoNewline
        Write-Host -Object "* $menuTitel *" -ForegroundColor $fcolor -BackgroundColor $bcolor
        Write-Host -Object "`t" -NoNewline
        Write-Host -Object ('*' * $menuwidth) -ForegroundColor $fcolor -BackgroundColor $bcolor
        Write-Host -Object ''
        Write-Debug -Message "L: $l MenuItems: $menuItems MenuPosition: $menuPosition"
        for ($i = 0; $i -le $l; $i++) {
            Write-Host -Object "`t" -NoNewline
            if ($i -eq $menuPosition) {
                Write-Host -Object "$($menuItems[$i])" -ForegroundColor $bcolor -BackgroundColor $fcolor
            }
            else {
                Write-Host -Object "$($menuItems[$i])" -ForegroundColor $fcolor -BackgroundColor $bcolor
            }
        }
    }
  
    Function Get-Menu {
        param ([array]$menuItems, $menuTitel = 'MENU')
        $vkeycode = 0
        $pos = 0
        Show-Menu $menuItems $pos $menuTitel
        While ($vkeycode -ne 13) {
            $press = $host.ui.rawui.readkey('NoEcho,IncludeKeyDown')
            $vkeycode = $press.virtualkeycode
            Write-Host -Object "$($press.character)" -NoNewline
            If ($vkeycode -eq 38) {
                $pos--
            }
            If ($vkeycode -eq 40) {
                $pos++
            }
            if ($pos -lt 0) {
                $pos = 0
            }
            if ($pos -ge $menuItems.length) {
                $pos = $menuItems.length - 1
            }
            Show-Menu $menuItems $pos $menuTitel
        }
        Write-Output -InputObject $($menuItems[$pos])
    }
  
    Function Start-Menue {
        param
        (
            [Object]
            $MenueOptions
        )
  
        $MenueSelection = Get-Menu $MenueOptions 'Exchange Connection Manager'
        return $MenueSelection
    }

    $ScriptMenueItems = @()
    foreach ($connection in $connections) {
        $ScriptMenueItems += $connection.name
    }
    if ($ScriptMenueItems -ne $Null) {
        $choice = Start-Menue $ScriptMenueItems
        $connection = $connections |? {$_.Name -eq $choice}
        Start-ExchangeConnection -Connection $connection
    }
    else {
        $choice = Start-Menue 'NO CONNECTION CONFIGURED'
    }


}


function Start-ExchangeConnection {
    param($Connection)

    function Disconnect-Connection {
        param(
            [string]
            $ConnectionName,
            [string]
            $Prefix
        )
        $ConnectionDB = Get-ECMConnection
        $PSSessions = Get-PSSession |Select-Object -Property Name, State
  
        foreach ($PSSession in $PSSessions) {
            if ($PSSession.State -ne 'Opened' -and $ConnectionName -eq $PSSession.Name) {
                Write-ScriptLog -String ('Removing existing Connection: ' + $PSSession.Name) -Type WARNING
                Get-PSSession -Name $PSSession.Name |Remove-PSSession
                #$host.ui.RawUI.WindowTitle = $null
            }
            if ($PSSession.State -eq 'Opened' -and $ConnectionName -eq $PSSession.Name) {
                #Write-Warning -Message ('Session already active: ' + $PSSession.Name)
                Write-ScriptLog -String ('Session already active: ' + $PSSession.Name) -Type WARNING
                if ($Connection.SessionPrefix -ne $null -and $Connection.SessionPrefix -notlike '' ) {
                    Write-Host ("Command Prefix: " + $Prefix) -ForegroundColor DarkGray
                }
                return $false
            }
            if ($ConnectionName -ne $PSSession.Name) {
                #Get Session Prefix
                $SessionPrefix = $ConnectionDB |Where-Object -FilterScript {
                    $_.Name -eq $PSSession.Name
                }
          
                if ($SessionPrefix.SessionPrefix -eq $Prefix -or ($SessionPrefix.SessionPrefix -eq $null -and $Prefix -eq $null) -or ($SessionPrefix.SessionPrefix -like '' -and $Prefix -like '')  ) {
                    Write-ScriptLog -String ('Connection ' + $ConnectionName + ' uses the same Prefix as Connection ' + $PSSession.Name) -Type WARNING
                    Write-ScriptLog -String  ('Removing existing Connection: ' + $PSSession.Name) -Type WARNING
                    Get-PSSession -Name $PSSession.Name |Remove-PSSession
                }
            }
        }
        return $true
    }

    $RequiresNewSession = Disconnect-Connection -ConnectionName $Connection.Name -Prefix $Connection.SessionPrefix
    
    if ($RequiresNewSession) {
        Write-ScriptLog -String "Connecting to $($Connection.Name)" -Type WARNING

        #Create credential object
        if (($Connection.UserName -eq $null -or $Connection.Password -eq $null) -and ($Connection.Authentication -ne 'Kerberos') ) {
            $cred = Get-Credential
        }
        else {
            if ($Connection.Authentication -ne 'Kerberos') {
                $cred = Get-DecryptedCredential $Connection
            }
        }

        if ($Connection.Type -eq 'Cloud') {
            $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $Connection.URI -Credential $cred -Authentication Basic -AllowRedirection -Name $Connection.Name
            $cred = $null
        }
        else {

            if ($Connection.Authentication -eq 'Kerberos') {
                $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $Connection.URI -Authentication Kerberos -AllowRedirection -Name $Connection.Name
            }
            if ($Connection.Authentication -eq 'Basic') {
                $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $Connection.URI -Credential $cred -Authentication Basic -AllowRedirection -Name $Connection.Name
            }
        }

        $cred = $null
        
        if ($ExchangeSession -eq $null) {
            Write-ScriptLog ('Unable to connect to Session ' + $Connection.Name) -Type ERROR
            return 
        }

        if ($Connection.SessionPrefix -ne $null -and $Connection.SessionPrefix -notlike '') {
            Import-Module (Import-PSSession -Session $ExchangeSession -AllowClobber -DisableNameChecking) -Global -DisableNameChecking -Prefix $Connection.SessionPrefix
            
        }
        else {
            Import-Module (Import-PSSession -Session $ExchangeSession -AllowClobber -DisableNameChecking) -Global -DisableNameChecking
        }

        Clear-Host
        Write-ScriptLog -String "Session connected: $($Connection.Name)"
        if ($Connection.SessionPrefix -ne $null -and $Connection.SessionPrefix -notlike '' ) {
            Write-ScriptLog -String ("Command Prefix: " + $Connection.SessionPrefix)
        }

        if ((Get-ECMConfig).EditEnumerationLimit) {
            $global:FormatEnumerationLimit = -1
        }
        
    }
}

###########################
# User functions
###########################

function New-ECMConnection {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantName,
        [PSCredential]$Credential,
        [string]$Name,
        [string]$SessionPrefix
    )
    try {
        if (!$Name) {
            $Name = $TenantName
        }

        $connection = Get-ECMConnection -Name $Name
        if ($connection) {
            Write-Error "A connection with the name $Name already exists. Please specify a differnet name."
            return
        }

        if (!$Credential) {
            Write-host 'WARNING: Credential missing. Are you prepared to enter your credential each time the connection is selected ? [Y/n]' -NoNewline -ForegroundColor Yellow
            $input = Read-Host
            if ($input -eq 'n') {
                $Credential = get-credential -Message ('Credential for connection: ' + $Name )
            }
        }

        $ConnectionObject = @{
            'TenantName'     = $TenantName;
            'Name'           = $Name;
            'SessionPrefix'  = $SessionPrefix;
            'Credential'     = if ($Credential) {(Get-EncryptedCredential $Credential)};
            'Type'           = 'Cloud'
            'URI'            = 'https://outlook.office365.com/powershell-liveid/'
            'Authentication' = 'Basic'
        }

        $db = Get-ConnectionFile
        $Newdb = @()
        $db | % {
            $Newdb += $_
        }
        $Newdb += $ConnectionObject

        Unlock-ECMFiles
        $Newdb |ConvertTo-Json |Out-File $global:ECMFiles.ConnectionFilePath -Force
        Lock-ECMFiles

    }
    catch {
        Write-Error 'Cannot create connection'
        $_ |select -expandproperty invocationinfo
    }
}

function New-ECMOnPremConnection {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [Parameter(Mandatory = $true)]
        [string]$URI,
        [ValidateSet('Kerberos', 'Basic')]
        [string]$Authentication = 'Kerberos',
        [string]$SessionPrefix

    )

    DynamicParam {

        if ($Authentication -eq 'Basic') {
            #create a new ParameterAttribute Object
            $ageAttribute = New-Object System.Management.Automation.ParameterAttribute
            $ageAttribute.Position = 4
            $ageAttribute.Mandatory = $true

            #create an attributecollection object for the attribute we just created.
            $attributeCollection = new-object System.Collections.ObjectModel.Collection[System.Attribute]

            #add our custom attribute
            $attributeCollection.Add($ageAttribute)

            #add our paramater specifying the attribute collection
            $ageParam = New-Object System.Management.Automation.RuntimeDefinedParameter('Credential', [pscredential], $attributeCollection)

            #expose the name of our parameter
            $paramDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
            $paramDictionary.Add('Credential', $ageParam)
            return $paramDictionary

        }
    }

    process {

        try {

            $connection = Get-ECMConnection -Name $Name
            if ($connection) {
                Write-Error "A connection with the name $Name already exists. Please specify a differnet name."
                return
            }

            $ConnectionObject = @{
                'Name'           = $Name;
                'SessionPrefix'  = $SessionPrefix;
                'Credential'     = if ($PSBoundParameters.credential) {(Get-EncryptedCredential $PSBoundParameters.credential )};
                'Type'           = 'OnPrem'
                'URI'            = $URI
                'Authentication' = $Authentication
            }

            $db = Get-ConnectionFile
            $Newdb = @()
            $db | % {
                $Newdb += $_
            }
            $Newdb += $ConnectionObject

            Unlock-ECMFiles
            $Newdb |ConvertTo-Json |Out-File $global:ECMFiles.ConnectionFilePath -Force
            Lock-ECMFiles

        }
        catch {
            Write-Error 'Cannot create connection'
            $_ |select -expandproperty invocationinfo
        }

    }



}

function Get-ECMConfig {
  
    $ECMConfig = Get-Content -Path $global:ECMFiles.ECMConfigFile |ConvertFrom-Json
    return $ECMConfig

}

function Set-ECMConfig {
    param(
        [ValidateSet('Auto', 'Core')]
        [string]$Menu,
        [bool]$CheckForUpdates,
        [bool]$EditEnumerationLimit,
        [bool]$Logging,
        [string]$Logfile
    )

    $ECMConfig = Get-ECMConfig

    if ($PSBoundParameters.ContainsKey('Menu')) {
        $ECMConfig.Menu = $PSBoundParameters.Menu
    }
    if ($PSBoundParameters.ContainsKey('CheckForUpdates')) {
        $ECMConfig.CheckForUpdates = $PSBoundParameters.CheckForUpdates
    }
    if ($PSBoundParameters.ContainsKey('EditEnumerationLimit')) {
        $ECMConfig.EditEnumerationLimit = $PSBoundParameters.EditEnumerationLimit
    }
    if ($PSBoundParameters.ContainsKey('Logging')) {
        $ECMConfig.Logging = $PSBoundParameters.Logging
    }
    if ($PSBoundParameters.ContainsKey('Logfile')) {
        $ECMConfig.Logfile = $PSBoundParameters.Logfile
    }

    if ($PSBoundParameters -ne $null) {
        Unlock-ECMFiles
        $ECMConfig |ConvertTo-Json |Out-File -FilePath $global:ECMFiles.ECMConfigFile -Force
        Lock-ECMFiles
    }
   
}


function Get-ECMConnection {
    param(
        [string]$Name
    )
    begin {

    }process {

 
        $defaultDisplaySet = 'Name', 'UserName', 'SessionPrefix', 'Type'
        $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet', [string[]]$defaultDisplaySet)
        $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
        $db = Get-ConnectionFile 
        Write-ScriptLog -String ($db |Out-String) -Type INFO -Silent
        $db = $db |select * -ExpandProperty credential |select Name, TenantName, UserName, Password, Type, SessionPrefix, URI, Authentication
        $db.PSObject.TypeNames.Insert(0, 'ECM.ConnectionObject')
        $db | Add-Member MemberSet PSStandardMembers $PSStandardMembers
        if ($Name) {
            $db |? {$_.Name -eq $Name}
        }
        else {
            return $db
        }
        
    }
    end {

    }
}
function Remove-ECMConnection {
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$Name,
        [bool]$Confirm = $true
    )
  
    begin {}
    process {
        $db = Get-ConnectionFile 

        #Verfify Connection Name
        $check = $db |? {$_.Name -eq $Name}

        if (!$check) {
            Write-Error "Cannot remove connection. Connection not found: $Name "
            return
        }
    
        #Create New DB
        $Newdb = @()
        $db | % {
            if ($_.Name -ne $Name) {
                $Newdb += $_
            }  
        }

        if ($Confirm -ne $false) {
            Write-Host "Do you really want to permanently remove the connection: $Name " -ForegroundColor Yellow
            Write-Host -Object '[Y] Yes [N] No (default is "N"): ' -NoNewline -ForegroundColor Yellow
            $cv = Read-Host 
            if ($cv -ne 'Y') {
                return 
            }
        }

        Unlock-ECMFiles
        $Newdb |ConvertTo-Json |Out-File $global:ECMFiles.ConnectionFilePath -Force
        Lock-ECMFiles
 
    }end {}
}
function Connect-Exchange {

    param(
        [string]$Identity,
        [switch]$CoreMenu
        
    )

    #load data
    $Connections = Get-ECMConnection
    
    #get env.
    if (($PSVersionTable.PSEdition -like '*core*') -or ($host.name -ne 'ConsoleHost') -or ((Get-ECMConfig).Menu -eq 'Core')) {
        $CoreMenu = $true
    }


    if (!$Identity) {
        if ($CoreMenu) {
            
            do { 
                Show-CoreMenu -Connections $Connections
                $input = Read-Host "Please make a selection" 
            } 
            until ($input -eq 'q' -or ($input -le @($Connections).count))
    
            Clear-Host
            if ($input -ne 'q') {
                $connection = $connections[$input - 1]
                Start-ExchangeConnection -Connection $connection
            }

        }
        else {
            Show-FullMenu -Connections $Connections
        }
        
    }
    else {
        $connection = $connections |? {$_.Name -eq $Identity}
        if (!$Connection) {
            Write-Error "Connection not found: "$Identity
            return
        }
        Start-ExchangeConnection -Connection $connection
    }


}

function Set-ECMConnection {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity,
        [string]$Name,
        [string]$TenantName,
        [pscredential]$Credential,
        [string]$SessionPrefix,
        [string]$URI,
        [string]$Authentication
    )

    $db = Get-ConnectionFile
    
    #Verfify Connection Name
    $connection = $db |? {$_.Name -eq $Identity}

    if (!$connection) {
        Write-Error "Cannot change connection. Connection not found: $Identity "
        return
    }

    if ($Name) {
        $connection.Name = $Name
    }
    if ($TenantName) {
        $connection.TenantName = $TenantName
    }
    if ($TenantName) {
        $connection.SessionPrefix = $SessionPrefix
    }
    if ($Credential) {
        $connection.Credential = Get-EncryptedCredential $Credential
    }
    if ($URI) {
        $connection.URI = $URI
    }
    if ($Credential) {
        $connection.Authentication = $Authentication
    }

    #Update connection file
    $UpdatedConnectionFile = @()
    $UpdatedConnectionFile += $connection
    $UpdatedConnectionFile += $db |? {$_.Name -ne $Identity}
    Unlock-ECMFiles
    $UpdatedConnectionFile |ConvertTo-Json |Out-File $global:ECMFiles.ConnectionFilePath -Force
    Lock-ECMFiles

}

function Start-ECM {
    Connect-Exchange
}

#Process
$global:ECMFiles = Initialize-ECM

if ((Get-ECMConfig).CheckForUpdates -and (((Get-ECMConfig).LastUpdateCheck|get-date).AddDays(1) -lt (get-date))) {
    Update-ECM
}

Export-ModuleMember -Function Connect-Exchange, Remove-ECMConnection, Get-ECMConnection, Get-ECMConfig, New-ECMConnection, New-ECMOnPremConnection, Set-ECMConnection, Start-ECM, Set-ECMConfig
