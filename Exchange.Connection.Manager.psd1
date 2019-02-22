
@{

    # Script module or binary module file associated with this manifest.
    RootModule = 'Exchange.Connection.Manager.psm1'
    
    # Version number of this module.
    ModuleVersion = '2.0.0'
    
    # ID used to uniquely identify this module
    GUID = 'b92d5fd8-a9cf-4a0e-9439-c85b12715e32'
    
    # Author of this module
    Author = 'Andreas Bode'
    
    # Company or vendor of this module
    CompanyName = 'atwork deutschland GmbH'
    
    # Copyright statement for this module
    Copyright = '(c) 2016 Andreas Bode. All rights reserved.'
    
    # Description of the functionality provided by this module
    Description = 'Are you an Exchange Admin? Are you managing multiple Exchange OnPrem and Online environments? Tired of switching between multiple shells and sessions, updating and entering credentials? Check out the PowerShell Exchange Connection Manager module.'
    
    # Minimum version of the Windows PowerShell engine required by this module
    # PowerShellVersion = ''
    
    # Name of the Windows PowerShell host required by this module
    # PowerShellHostName = ''
    
    # Minimum version of the Windows PowerShell host required by this module
    # PowerShellHostVersion = ''
    
    # Minimum version of Microsoft .NET Framework required by this module
    # DotNetFrameworkVersion = ''
    
    # Minimum version of the common language runtime (CLR) required by this module
    # CLRVersion = ''
    
    # Processor architecture (None, X86, Amd64) required by this module
    # ProcessorArchitecture = ''
    
    # Modules that must be imported into the global environment prior to importing this module
    # dModules = @()
    
    # Assemblies that must be loaded prior to importing this module
    # RequiredAssemblies = @()
    
    # Script files (.ps1) that are run in the caller's environment prior to importing this module.
    # ScriptsToProcess = @()
    
    # Type files (.ps1xml) to be loaded when importing this module
    # TypesToProcess = @()
    
    # Format files (.ps1xml) to be loaded when importing this module
    # FormatsToProcess = @()
    
    # Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
    # NestedModules = @()
    
    # Functions to export from this module
    FunctionsToExport = @('Connect-Exchange', 'Remove-ECMConnection', 'Get-ECMConnection', 'Get-ECMConfig', 'New-ECMConnection', 'New-ECMOnPremConnection', 'Set-ECMConnection', 'Start-ECM', 'Set-ECMConfig')
    
    # Cmdlets to export from this module
    #CmdletsToExport = '*'
    
    # Variables to export from this module
    VariablesToExport = '*'
    
    # Aliases to export from this module
    #AliasesToExport = '*'
    
    # DSC resources to export from this module
    # DscResourcesToExport = @()
    
    # List of all modules packaged with this module
    # ModuleList = @()
    
    # List of all files packaged with this module
    # FileList = @()
    
    # Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
    PrivateData = @{
    
        PSData = @{
    
            #Prerelease = ''
            # Tags applied to this module. These help with module discovery in online galleries.
            # Tags = @()
    
            # A URL to the license for this module.
            LicenseUri = 'https://andreasbode.net/Exchange.Connection.Manager'
    
            # A URL to the main website for this project.
            ProjectUri = 'https://andreasbode.net/Exchange.Connection.Manager'
    
            # A URL to an icon representing this module.
            IconUri = 'https://en.gravatar.com/userimage/150089748/d366f3c6469387b15bb6a110cc0fb83b.jpeg'
    
            # ReleaseNotes of this module
            ReleaseNotes = 'The new release finally adds support for Powershell Core‼️ You can now use ECM on basically any platform like macOS, Linux and even iOS using Azure Cloud Shell.'                
             
        } # End of PSData hashtable
    
    } # End of PrivateData hashtable
    
    # HelpInfo URI of this module
    # HelpInfoURI = ''
    
    # Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
    # DefaultCommandPrefix = ''
    
    }
    
    