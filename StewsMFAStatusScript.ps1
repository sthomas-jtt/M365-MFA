<# 

This script was developed with the assistance of Microsoft Copilot, an AI companion that provided 
technical insights, optimization suggestions, and structural enhancements to improve functionality 
and efficiency.

It was designed to be user friendly, assist with installing the required modules, and offer methods 
for filtering the results to retrieve specific information. It will offer options to export the 
results to a CSV file (saved to the directory where the script was executed from), and an option
to see a summary with statistics based on your filter criteria. 

For MFA, we're concerned about having MFA configured, specifically with Strong authentication methods. 
We consider Email Authentication and Phone Authentication (SMS) to be legacy methods (Weak MFA). We
especially need to verify MFA for Global Administrators, as these accounts have Global access to all
aspects of Azure/M365. This is the key to the Kingdom!

Now that legacy (per-user MFA) authentication is being retired and MFA management moving to Security 
Defaults or Conditional Access, we should be setting all per-user MFA settings to disabled!
https://learn.microsoft.com/en-us/microsoft-365/admin/security-and-compliance/multi-factor-authentication-microsoft-365 

Script Information
------------------
Author: Stewart Thomas | Jackson Thornton Technologies
Contact: sthomas@jttconnect.com
Description: Extract and export results of Microsoft 365 Multi-factor Authentication
Last Updated: 6/2/2025
Modules: Microsoft.Graph.Authentication, Microsoft.Graph.Identity.DirectoryManagement
Modules: Microsoft.Graph.Users, Microsoft.Graph.Beta.Users, Microsoft.Graph.Beta.Identity.SignIns
Scopes: User.Read.All, UserAuthenticationMethod.Read.All, Policy.ReadWrite.AuthenticationMethod, Domain.Read.All 

#>


# Declare some variables and arrays
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm"
$Location=Get-Location
$results = @()  # Initialize an empty array
$ProcessedUserCount=0
$globalAdmins = @()
$counter = 0
$MFAMethodsCount = 0
$MFAPhoneDetail = $null
$MicrosoftAuthenticatorDevices = $null
$hasAuthenticator = $false

# Use [Environment]::NewLine to create a newline variable for clarity
$newline = [Environment]::NewLine




# Setup all the functions


function Maximize-PowerShellWindow {
    # Clear the screen
    Clear-Host

    # Check if the WinAPI type already exists
    if (-not ("WinAPI" -as [type])) {
        Add-Type -TypeDefinition @"
        using System;
        using System.Runtime.InteropServices;
        public class WinAPI {
            [DllImport("user32.dll")]
            public static extern int ShowWindow(IntPtr hWnd, int nCmdShow);
            [DllImport("user32.dll")]
            public static extern IntPtr GetForegroundWindow();
        }
"@
    }

    # Get the current window handle and maximize the window
    $WinHandle = [WinAPI]::GetForegroundWindow()
    [WinAPI]::ShowWindow($WinHandle, 3) | Out-Null # 3 = Maximize window
}

function Clear-MgContextCache {
<# 
In modern versions of MSAL (used by Microsoft.Graph), the token cache is stored in a folder called .IdentityService
Check for Get-MgContext which indicates the script has been run already and we have context
If not we will delete the tokens cache folder so that we don't run the script and automatically use a cached token when there's no context
#>
    if (-not (Get-MgContext)) {
        $cacheFolder = "$env:LOCALAPPDATA\.IdentityService"
        if (Test-Path $cacheFolder) {
            Remove-Item $cacheFolder -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

Function Check_Modules {

$UserChoice = Get-YesNoInput "First, do you need to check for missing Powershell modules?"

if ($UserChoice -eq $true) {
    Write-Host "`nChecking for required Microsoft Graph modules..." -ForegroundColor Cyan

    $RequiredModules = @(
        "Microsoft.Graph.Authentication",
        "Microsoft.Graph.Beta.Identity.SignIns",
        "Microsoft.Graph.Beta.Users",
        "Microsoft.Graph.Identity.DirectoryManagement",
        "Microsoft.Graph.Users"
    )

    # Get installed modules once (caching for performance)
    $AvailableModules = Get-Module -ListAvailable

    $MissingModules = $RequiredModules | Where-Object { -not ($AvailableModules | Where-Object Name -eq $_) }

    if ($MissingModules.Count -gt 0) {
        Write-Host "`nInstalling missing modules..." -ForegroundColor Yellow
        Write-Host ""
        Install-Module -Name $MissingModules -Scope CurrentUser -AllowClobber -Force
        Write-Host "`nInstallation completed." -ForegroundColor Magenta
        Write-Host ""
    } else {
        Write-Host "`nAll required modules are already installed." -ForegroundColor Green
        Write-Host ""
    }

    # Import only necessary modules
    $RequiredModules | ForEach-Object {
        if (-not (Get-Module -Name $_)) {
            Import-Module $_ -ErrorAction SilentlyContinue
        }
    }

    Write-Host "`nModules verification complete." -ForegroundColor Cyan
    Write-Host ""
} else {
    Write-Host "`nSkipping module verification..." -ForegroundColor Green
    Write-Host ""
}

}

function Show-SummaryInNotepad {
    param(
        [string]$Title = "MFA Summary Information",
        [string]$SummaryText
    )

    # Define a temporary file path
    $TempFile = "$env:TEMP\Summary.txt"

    # Write the summary text to the file
    $SummaryText | Out-File -Encoding UTF8 $TempFile

    # Open the file in Notepad
    Start-Process "notepad.exe" -ArgumentList $TempFile
}

function Get-YesNoInput {
    param(
        [string]$Prompt
    )
    # Display the prompt without advancing to a new line
    #Write-Host "$Prompt (Y/N): " -NoNewline
    Write-Host $Prompt -NoNewline -ForegroundColor White
    Write-Host " (Y/N): " -NoNewline -ForegroundColor Yellow


    do {
        # Read a single key; NoEcho prevents it from displaying automatically, IncludeKeyDown captures the key when pressed
        $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        $response = $key.Character
    } until ($response -match "^[YyNn]$")
    
    # Optionally, echo the user's choice on the same line
    Write-Host $response
    
    return $response -match "^[Yy]$"
}

function Get-FilterModeInput {
   
    
    Write-Host "(1) " -ForegroundColor Yellow -NoNewline
    Write-Host "None - Get all results, " -ForegroundColor White -NoNewline
    Write-Host "(2) " -ForegroundColor Yellow -NoNewline
    Write-Host "Flexible - Match any, " -ForegroundColor White -NoNewline
    Write-Host "(3) " -ForegroundColor Yellow -NoNewline
    Write-Host "Strict - Must match all, " -ForegroundColor White -NoNewline
    Write-Host "(4) " -ForegroundColor Yellow -NoNewline
    Write-Host "Default filter" -ForegroundColor White
    Write-Host "Choose filter mode " -ForegroundColor White -NoNewline
    Write-Host "(1-4): " -ForegroundColor Yellow -NoNewline

    do {
        # Read a single key immediately 
        $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        $choice = $key.Character
    } until ($choice -match "^[1-4]$")  # Ensure the input is 1, 2, 3, or 4

    # Optionally, display the captured input
    Write-Host $choice

    # Convert the input from a character to its numeric value
    return [int]::Parse($choice)
}

function Get-FilterConfiguration {
    # Get the filter mode selection.
    $FilterMode = Get-FilterModeInput 

    # Create a custom object to hold the configuration.
    $config = [PSCustomObject]@{
        FilterMode           = $FilterMode
        UseDefaultFilter     = $false
        IncludeGlobalAdmins  = $false
        IncludeLicensed      = $false
        IncludeMFADisabled   = $false
        IncludePeruserMFA    = $false
        IncludeSigninAllowed = $false
    }

    # If a filter mode is one of the modes that might require additional inputs
    if ($FilterMode -eq 2 -or $FilterMode -eq 3 -or $FilterMode -eq 4) {
        if ($FilterMode -eq 4) {
            $config.UseDefaultFilter    = $true
            # Change filter mode to 2 internally; default filter uses flexible matching.
            $config.FilterMode          = 2
            $config.IncludeGlobalAdmins = $true
            $config.IncludeLicensed     = $true
        }
        else {
            $config.IncludeGlobalAdmins  = Get-YesNoInput "Include Global Admins?"
            $config.IncludeLicensed      = Get-YesNoInput "Include Licensed Users?"
            $config.IncludeMFADisabled   = Get-YesNoInput "Include Users with no MFA?"
            $config.IncludePeruserMFA    = Get-YesNoInput "Include Users with Per-user MFA NOT disabled?"
            $config.IncludeSigninAllowed = Get-YesNoInput "Include Users allowed to sign-in?"
            Write-Host ""
        }
    }

    # Return a configuration object containing all the settings.
    return $config
}

Function Connect_MgGraph {
    $Scopes = @(
        "User.Read.All",
        "UserAuthenticationMethod.Read.All",
        #"Policy.ReadWrite.AuthenticationMethod",
        "Policy.Read.All"
        "Directory.Read.All",
        "Domain.Read.All"
    )

    # Check if already authenticated
    $MgContext = Get-MgContext
    if ($MgContext) {
        #Write-Host "Currently authenticated as: $($MgContext.Account)" -ForegroundColor Cyan
        
        # Get Tenant Info if authenticated
        $TenantInfo = Get-MgOrganization
        if ($TenantInfo) {
            $script:TenantDomain = ($TenantInfo.VerifiedDomains | Where-Object { $_.IsInitial -eq $true } | Select-Object -ExpandProperty Name)

            Write-Host "Connected successfully as: $($MgContext.Account) to $TenantDomain" -ForegroundColor Cyan
            Write-Host ""
        } else {
            Write-Host "Unable to retrieve tenant information." -ForegroundColor Red
        }

        #Write-Host ""
        $disconnectConfirm = Get-YesNoInput "Do you want to disconnect and sign in with a different account?"
        
        
        if ($disconnectConfirm -eq $true) {
            Write-Host ""
            Write-Host "Disconnecting Microsoft Graph session..." -ForegroundColor Yellow
            Disconnect-MgGraph | Out-Null
            
            # Clear authentication context after disconnect
            $MgContext = $null
            Write-Host ""

            # Proceed to authentication after disconnect
            Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
            Connect-MgGraph -Scopes $Scopes #-NoWelcome

            # Refresh authentication context after logging in
            $MgContext = Get-MgContext
        } else {
            Write-Host "Continuing with the current authentication session." -ForegroundColor Green
            Write-Host ""
            return
        }
    } else {
        # Authenticate with Microsoft Graph if no active session
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
        Connect-MgGraph -Scopes $Scopes #-NoWelcome

        # Refresh authentication context after logging in
        $MgContext = Get-MgContext
    }

    # Confirm successful connection
    if ($MgContext) {
        # Get Tenant Info if authenticated
        $TenantInfo = Get-MgOrganization
        if ($TenantInfo) {
            $TenantDomain = ($TenantInfo.VerifiedDomains | Where-Object { $_.IsInitial -eq $true } | Select-Object -ExpandProperty Name)
       } else {
            Write-Host "Unable to retrieve tenant information." -ForegroundColor Red
        }
        Write-Host "Connected successfully as: $($MgContext.Account) to $TenantDomain" -ForegroundColor Green
        Write-Host ""
    } else {
        Write-Host "Microsoft Graph connection failed." -ForegroundColor Red
    }
}

   

######################
#
# Starting the script
#
######################

# Call the function to maximize the window
Maximize-PowerShellWindow

# Call the function to check for MgContext
Clear-MgContextCache

# Write some instructions to the screen
$menutext = @"
---------------------------------------

Welcome to the Stew's fancy MFA script!

The results will always be displayed in a separate Grid View Powershell window. (Like a spreadsheet)

You can choose to filter the results. Filter mode options are:
  1) Including all users in the report
  2) Filter where any condition can be true. Flexible = More Results
  3) Filter where ALL conditions must be true. Strict = Fewer Results
  4) Choose Default Filter = Include any Global Admins or any Licensed accounts

If you choose to filter, you will be asked which filter mode, and if you want to include:
  - Global Admins
  - Licensed users
  - Users without MFA methods
  - Users NOT set to disabled in Per-user MFA
  - Users with Sign-in allowed

You can choose to export the results to a CSV file that will be saved to the same location as the script.

You can choose to display a summary of the results in a separate window.

----------------------------------------

"@
Write-Output $menutext


# Prompt for check_modules, filters, export results, show summary, and then run the Connect_MgGraph
Check_Modules # Runs the Check_Modules function
$filterConfig = Get-FilterConfiguration # Runs the Get-FilterConfiguration function
$ExportResults = (Get-YesNoInput "Export Results to CSV?") # Prompts to export results
if($ExportResults -eq $true) { 
    Write-Host "File path will be: " -NoNewline
    Write-Host $Location -ForegroundColor Cyan $newline
}
$ShowSummaryWindow = (Get-YesNoInput "Would you like to see a summary?") # Prompts to output summary window
Write-Host ""
# Ask user if they want to disconnect when the script is finished
$DisconnectLater = Get-YesNoInput "Do you want to disconnect from Microsoft Graph after the script completes?" 
Write-Host ""
Connect_MgGraph # Now we'll run the Connect_MgGraph function to connect to MgGraph


# Get list of users
$users = Get-MgUser -All -Property "Id,DisplayName,UserPrincipalName,UserType,AccountEnabled,AssignedLicenses" |
    Where-Object { $_.UserType -eq "Member" }

# Determine if user is a Global Admin
$roles = Get-MgDirectoryRole
$globalAdminRole = $roles | Where-Object { $_.DisplayName -eq "Global Administrator" }
if ($globalAdminRole) {
    $globalAdmins = Get-MgDirectoryRoleMember -DirectoryRoleId $globalAdminRole.Id | ForEach-Object { $_.Id }
}

#Check if Security Defaults are enabled
$SecurityDefaultsEnabled = Get-MgPolicyIdentitySecurityDefaultEnforcementPolicy | Select-Object -ExpandProperty IsEnabled
if ($SecurityDefaultsEnabled -eq $true) {
    Write-Host "Security Defaults are ENABLED$newline" -ForegroundColor Green
    } else {
    Write-Host "Security Defaults are DISABLED$newline" -ForegroundColor Red
    }

# Get total users count
$total = $users.Count
Write-Host "There are $total Users in this tenant$newline"

# Get all users from MgBetaUser. This is the main loop for checking users and MFA details
Get-MgBetaUser -Filter "userType eq 'Member'" | foreach {

    $ProcessedUserCount++
    $Name= $_.DisplayName
    $UPN=$_.UserPrincipalName
    $UserId=$_.Id

    $PercentComplete = [math]::Floor(($ProcessedUserCount / $total) * 100)
    Write-Progress -Activity "Processing user: $ProcessedUserCount - Processing $Name" -PercentComplete ([Math]::Min(100, $PercentComplete))

    $isGlobalAdmin = $_.Id -in $globalAdmins  # Check if user ID is in global admin list

    if($_.AccountEnabled -eq $true) { $SigninStatus="Allowed" } else { $SigninStatus="Blocked" }
    if(($_.AssignedLicenses).Count -ne 0) { $LicenseStatus="Licensed" } else { $LicenseStatus="Unlicensed" }   
 
    [array]$MFAData=Get-MgBetaUserAuthenticationMethod -UserId $UserId # changed to $UserId instead of $UPN
    $AuthenticationMethod=@()
    $AdditionalDetails=@()
    
    #Filter out password authentication as an "authentication method" because it isn't what we're looking for
    $FilteredMFAData = $MFAData | Where-Object { $_.AdditionalProperties["@odata.type"] -ne "#microsoft.graph.passwordAuthenticationMethod" }
 
    foreach($MFA in $FilteredMFAData)
    { 
        Switch ($MFA.AdditionalProperties["@odata.type"]) 
        { 
        "#microsoft.graph.passwordAuthenticationMethod"
        {
            Continue  # Skip processing this method
        } 
        "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod"  
        { # Microsoft Authenticator App
            $AuthMethod     = 'AuthenticatorApp'
            $AuthMethodDetails = $MFA.AdditionalProperties["displayName"] 
        }
        "#microsoft.graph.phoneAuthenticationMethod"                  
        { # Phone authentication
            $AuthMethod     = 'PhoneAuthentication'
            $AuthMethodDetails = $MFA.AdditionalProperties["phoneType", "phoneNumber"] 
        } 
        "#microsoft.graph.fido2AuthenticationMethod"                   
        { # FIDO2 key
            $AuthMethod     = 'Passkeys(FIDO2)'
            $AuthMethodDetails = $MFA.AdditionalProperties["model"] 
        }  
        "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod" 
        { # Windows Hello
            $AuthMethod     = 'WindowsHelloForBusiness'
            $AuthMethodDetails = $MFA.AdditionalProperties["displayName"]
        }                        
        "#microsoft.graph.emailAuthenticationMethod"        
        { # Email Authentication
            $AuthMethod     = 'EmailAuthentication'
            $AuthMethodDetails = $MFA.AdditionalProperties["emailAddress"] 
        }               
        "#microsoft.graph.temporaryAccessPassAuthenticationMethod"   
        { # Temporary Access pass
            $AuthMethod     = 'TemporaryAccessPass'
            $AuthMethodDetails = 'Access pass lifetime (minutes): ' + $MFA.AdditionalProperties["lifetimeInMinutes"] 
        }
        "#microsoft.graph.passwordlessMicrosoftAuthenticatorAuthenticationMethod" 
        { # Passwordless
            $AuthMethod     = 'PasswordlessMSAuthenticator'
            $AuthMethodDetails = $MFA.AdditionalProperties["displayName"] 
        }      
        "#microsoft.graph.softwareOathAuthenticationMethod"
        { # SoftwareOath
            $AuthMethod     = 'SoftwareOath'
            $AuthMethodDetails = $MFA.id           
        }
   }

  
   $AuthenticationMethod +=$AuthMethod
   if($AuthMethodDetails -ne $null) 
   {
       $MFAMethodsCount ++
   }
 }
    
    # Set some variables
    $AuthenticatorAppMethods = $MFAData | Where-Object { $_.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" } | ForEach-Object { $_.AdditionalProperties["displayName"] }
    $HelloMethods = $MFAData | Where-Object { $_.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod" } | ForEach-Object { $_.AdditionalProperties["displayName"] }
    $MFAPhoneDetail = $MFAData | Where-Object { $_.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.phoneAuthenticationMethod" } | ForEach-Object { $_.AdditionalProperties["phoneNumber"] }
    $MFAOathDetail = $MFAData | Where-Object { $_.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.softwareOathAuthenticationMethod" } | ForEach-Object { $MFA.Id }
    $EmailMethodDetails = $MFAData | Where-Object { $_.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.emailAuthenticationMethod" } | ForEach-Object { $_.AdditionalProperties["emailAddress"] }
    $PasswordlessMethods = $MFAData | Where-Object { $_.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.passwordlessMicrosoftAuthenticatorAuthenticationMethod" } | ForEach-Object { $_.AdditionalProperties["displayName"] }
    $FidoMethods = $MFAData | Where-Object { $_.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.fido2AuthenticationMethod" } | ForEach-Object { $_.AdditionalProperties["model"] }
    $TempAccessMethods = $MFAData | Where-Object { $_.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.temporaryAccessPassAuthenticationMethod" } | ForEach-Object { $_.AdditionalProperties["lifetimeInMinutes"] }

    # To remove duplicate authentication methods
    $AuthenticationMethod =$AuthenticationMethod | Sort-Object | Get-Unique
    $AuthenticationMethods= $AuthenticationMethod  -join ","
    $AdditionalDetail=$AdditionalDetails -join ", "
        
    # Get the default MFA method
    $DefaultMFAUri = "https://graph.microsoft.com/beta/users/$UserId/authentication/signInPreferences"
    $GetDefaultMFAMethod = Invoke-MgGraphRequest -Uri $DefaultMFAUri -Method GET
    if ($GetDefaultMFAMethod.userPreferredMethodForSecondaryAuthentication) {
        $MFAMethodisDefault = $GetDefaultMFAMethod.userPreferredMethodForSecondaryAuthentication
        Switch ($MFAMethodisDefault) {
            "push" { $MFAMethodisDefault = "Microsoft authenticator app" }
            "oath" { $MFAMethodisDefault = "Authenticator app or hardware token" }
            "voiceMobile" { $MFAMethodisDefault = "Mobile phone" }
            "voiceAlternateMobile" { $MFAMethodisDefault = "Alternate mobile phone" }
            "voiceOffice" { $MFAMethodisDefault = "Office phone" }
            "sms" { $MFAMethodisDefault = "SMS" }
            Default { $MFAMethodisDefault = "Unknown method" }
        }
    }
    else {
        $MFAMethodisDefault = "Not Enabled"
    }    
       
     
    # Per-user MFA status
    $PerUserMFAStatus=@(Invoke-MgGraphRequest -Method GET -Uri "/beta/users/$UserId/authentication/requirements").perUserMfaState


    # Define strong MFA methods
    $StrongMFAMethods = @("Fido2", "SoftwareOath", "PasswordlessMSAuthenticator", "AuthenticatorApp", "WindowsHelloForBusiness", "TemporaryAccessPass")

    # Default MFA strength
    $MFAStrength = "Disabled"

    # Check if AuthenticationMethod contains a strong MFA method. 
    if ($AuthenticationMethod | ForEach-Object { $StrongMFAMethods -contains $_ }) {
        $MFAStrength = "Strong"
    }
    
    # Ensure PhoneAuthentication, or EmailAuthentication are marked as Weak. Put this last in case they have a weak and a strong, we consider them Weak
    if ($MFAStatus -ne "Strong" -and ($AuthenticationMethod -match "PhoneAuthentication|EmailAuthentication")) {
        $MFAStrength = "Weak"
    }   


    # Lets write some output to the screen here
    Write-Host "###########################"
    Write-Host "Processing User $ProcessedUserCount of ($total) - $Name"
    if ($isGlobalAdmin) { Write-Host "Role: Global Admin" -ForegroundColor Yellow } else { "Role: User" }
    Write-Host "Sign-in Satus: $SigninStatus"
    Write-Host "License Status: $LicenseStatus"
    Write-Host "Per-user MFA Status: $PerUserMFAStatus"
    if($MFAStrength -eq "Disabled"){ Write-Host "MFA Strength: $MFAStrength" -ForegroundColor Red } 
    if($MFAStrength -eq "Strong") { Write-Host "MFA Strength: $MFAStrength" -ForegroundColor Green } 
    if($MFAStrength -eq "Weak") { Write-Host "MFA Strength: $MFAStrength" -ForegroundColor Cyan }
    Write-Host "Default MFA Method: $MFAMethodisDefault"
    if($AuthenticationMethods){ Write-Host "$MFAMethodsCount MFA Methods: " $AuthenticationMethods }
    if($AuthenticatorAppMethods){ Write-Host "Authenticator Apps: $($AuthenticatorAppMethods -join ', ')" }
    if($PasswordlessMethods){ Write-Host "Passwordless Authenticator methods: $($PasswordlessMethods -join ', ')" }
    if($EmailMethodDetails){ Write-Host "Email methods: $($EmailMethodDetails -join ', ')" }
    if($MFAPhoneDetail){ Write-Host "Phone methods: $($MFAPhoneDetail -join ', ')" }
    if($HelloMethods){ Write-Host "Hello for Business methods: $($HelloMethods -join ', ')" }
    if($MFAOathDetail){ Write-Host "Software Oath methods: $($MFAOathDetail -join ', ')" }
    if($FidoMethods){ Write-Host "Fido methods: $($FidoMethods -join ', ')" }
    if($TempAccessMethods){ Write-Host "Temp access methods: $($TempAccessMethods -join ', ')" }
      
    # Start building a PSCustomObject to display results
    $results += [PSCustomObject]@{
        DisplayName               = $Name
        UserPrincipalName         = $UPN
        Role                      = if ($isGlobalAdmin) { "Global Admin" } else { "User" }
        'License Status'          = $LicenseStatus
        'Sign-in Status'          = $signInStatus
        'Per-user MFA Status'     = $PerUserMFAStatus
        'MFA Strength'            = $MFAStrength
        'MFA Method Count'        = $MFAMethodsCount
        'Default MFA Method'      = $MFAMethodisDefault
        'MFA Methods'             = $AuthenticationMethods
        'MS Authenticator App'    = $($AuthenticatorAppMethods -join ', ')
        'Authentication Phone'    = $($MFAPhoneDetail -join ', ')
        'Email Methods'           = $($EmailMethodDetails -join ', ')
        'Software Methods'        = $($MFAOathDetail -join ', ')
        'Hello for Business'      = $($HelloMethods -join ', ')
        'Passwordless Methods'    = $($PasswordlessMethods -join ', ')
        'Fido_ Methods'           = $($FidoMethods -join ', ')
        'Temp Access Methods'     = $($TempAccessMethods -join ', ')
    }

    # Reset $MFAMethodsCount for the next run
    $MFAMethodsCount = 0

    Write-Output "" #new line in between users
}


# Set a counter for before results are filtered
$BeforeFilterCount = @($results).Count # shouldn't need this part $($results.count)

# Apply chosen filters
$ApplyFiltering = $filterConfig.IncludeGlobalAdmins -or $filterConfig.IncludeLicensed -or $filterConfig.IncludeMFADisabled -or $filterConfig.IncludePeruserMFA -or $filterConfig.IncludeSigninAllowed
if (!$ApplyFiltering) { <#Not filtering, so no filters#> } else {
        if ($filterConfig.FilterMode -eq 2) {
            # Flexible mode: record is included if it matches any enabled criteria.
            $results = $results | Where-Object {
                ($filterConfig.IncludeGlobalAdmins -and $_.Role -eq "Global Admin") -or
                ($filterConfig.IncludeLicensed -and $_.'License Status' -eq "Licensed") -or
                ($filterConfig.IncludeMFADisabled -and ($_.'MFA Strength' -eq "Disabled" -or $_.MFA_Method_Count -eq 0)) -or
                ($filterConfig.IncludePeruserMFA -and ($_.'Per-user MFA Status' -ne "disabled")) -or
                ($filterConfig.IncludeSigninAllowed -and $_.'Sign-in Status' -eq "Allowed")
            }
        }
        elseif ($filterConfig.FilterMode -eq 3) {
            # Strict mode: record is included only if it meets all conditions that are enabled.
            $results = $results | Where-Object {
                (!$filterConfig.IncludeGlobalAdmins -or $_.Role -eq "Global Admin") -and
                (!$filterConfig.IncludeLicensed -or $_.'License Status' -eq "Licensed") -and
                (!$filterConfig.IncludeMFADisabled -or ($_.'MFA Strength' -eq "Disabled" -or $_.MFA_Method_Count -eq 0)) -and
                (!$filterConfig.IncludePeruserMFA -or ($_.'Per-user MFA Status' -ne "disabled")) -and
                (!$filterConfig.IncludeSigninAllowed -or $_.'Sign-in Status' -eq "Allowed")
        }
    }
} 

# Set some counters after the results were filtered
$AfterFilterCount = @($results).Count 
#SkippedUserCount = ($BeforeFilterCount - $AfterFilterCount)



################
# Start building the Summary 
################
$TenantDomain = (Get-MgDomain | Where-Object {$_.isInitial}).Id
$TenantName = $TenantDomain -replace "\.onmicrosoft\.com",""
$summary += "MFA Summary for: $TenantName$newline"
$summary += "$timestamp $newline$newline"
if ($SecurityDefaultsEnabled -eq $true) {
    $summary += "Security Defaults are ENABLED$newline$newline"
    } else {
    Write-Host "Security Defaults are DISABLED$newline$newline"
    }


##########
# Filter configuration summary
##########
$summary += @"
##################################################
Filtered Configuration Summary
##################################################
$newline
"@
if($filterConfig.FilterMode -eq 1) { 
    $summary += "No filtering applied.$newline"
} else {
    if($filterConfig.FilterMode -eq 2){ $summary += "Filter Mode: Flexible$newline$newline" }
    if($filterConfig.FilterMode -eq 3){ $summary += "Filter Mode: Strict$newline$newline" }
    $summary += "IncludeGlobal: $($filterConfig.IncludeGlobalAdmins)$newline"
    $summary += "IncludeLicensed: $($filterConfig.IncludeLicensed)$newline"
    $summary += "IncludeMFADisabled: $($filterConfig.IncludeMFADisabled)$newline"
    $summary += "IncludePeruserMFA: $($filterConfig.IncludePeruserMFA)$newline"
    $summary += "IncludeSigninAllowed: $($filterConfig.IncludeSigninAllowed)$newline$newline"
}
$summary += "Export to CSV: $ExportResults$newline"
$summary += "Show Summary: $ShowSummaryWindow$newline$newline"





##########
# Filtered users summary
##########
$summary += @"
##################################################
Filtered Users Summary
##################################################
$newline
"@
$summary += "Total users processed: $total $newline"
$summary += "Users included in report: $AfterFilterCount $newline"
$summary += "Users skipped due to filters: $($total - $AfterFilterCount) $newline$newline"



##########
# Users Summary
##########

# Get counts for Global Admins, Licensed Users, Sign-in Allowed users
$GlobalAdminsCount  = ($results | Where-Object { $_.Role -eq "Global Admin" } | Measure-Object).Count
$LicensedUserCount  = ($results | Where-Object { $_.'License Status' -eq "Licensed" } | Measure-Object).Count
$SigninAllowedCount = ($results | Where-Object { $_.'Sign-in Status' -eq "Allowed" } | Measure-Object).Count

$summary += @"
##################################################
Users Summary
##################################################
$newline
"@
$summary += "Global Admins: $GlobalAdminsCount$newline"
$summary += "Licensed Users: $LicensedUserCount$newline"
$summary += "Sign-in Allowed Users: $SigninAllowedCount$newline$newline"



##########
# Users MFA Summary
##########
$summary += @"
##################################################
Users MFA Summary
##################################################
$newline
"@

# Calculate your values
$noMfaCount      = ($results | Where-Object { $_.'MFA Method Count' -eq 0 } | Measure-Object).Count
$totalUsers      = ($results | Select-Object -ExpandProperty UserPrincipalName -Unique).Count
$WeakMFACount    = ($results | Where-Object { $_.'MFA Strength' -eq "Weak" } | Measure-Object).Count
$StrongMFACount  = ($results | Where-Object { $_.'MFA Strength' -eq "Strong" } | Measure-Object).Count
$perUserMFACount = ($results | Where-Object { $_.'Per-user MFA Status' -ne "disabled" } | Measure-Object).Count
$summary += "Users With MFA Methods: $(($totalUsers - $noMfaCount))$newline"
$summary += "Users With NO MFA Methods: $noMfaCount$newline"
$summary += "Users with Weak MFA Strength: $WeakMFACount$newline"
$summary += "Users with Strong MFA Strength: $StrongMFACount$newline"
$summary += "Per-User MFA not disabled: $perUserMFACount$newline$newline"


##########
# Uses with No MFA Summary
##########
# Filter out users with no MFA methods (i.e. MFA Method Count equals 0)
$noMfaUsers = $results | Where-Object { $_.'MFA Method Count' -eq 0 }
# Create table output as a string
$noMfaTable = $noMfaUsers | Select-Object DisplayName, UserPrincipalName | Format-Table -AutoSize | Out-String
if ($noMfaUsers.Count -gt 0) {
# Build the summary string
$summary += @"
##################################################
Users with no MFA Summary
##################################################
"@
$summary += $noMfaTable + "`r`n"
} else {
$summary += @"
##################################################
No users with no MFA Methods found
##################################################
$newline
"@
}


##########
# MFA Method Breakdown Summary
##########
# Filter users who have at least one MFA method
if ($usersWithMfa = $results | Where-Object { $_.'MFA Method Count' -gt 0 }) {

    # For each user, split the MFA methods string into separate values
    $mfaMethodsArray = $usersWithMfa | ForEach-Object {
        $_.'MFA Methods' -split ',\s*'
    }
    
    # Group identical MFA methods and sort the results by count (most common first)
    $mfaBreakdown = $mfaMethodsArray | Group-Object | Sort-Object Count -Descending
    
    # Convert the formatted table output into a string
    $mfaBreakdownString = $mfaBreakdown | Format-Table Name, Count -AutoSize | Out-String

    # Append header text to the summary using CRLF ($newline) for new lines
    $summary += @"
##################################################
MFA Method Breakdown
##################################################
"@
    # Append the string version of the breakdown
    $summary += $mfaBreakdownString + "`r`n"
}


# Always display the report in a grid view
$results | Sort-Object Role, DisplayName | Out-GridView -Title "Microsoft 365 MFA Report"

# Export Results to CSV if the user chose to export
if($ExportResults -eq $true) {
    $outputPath = "$Location\MFA-Report-$TenantName-$timestamp.csv"
    $results | Sort-Object Role, DisplayName | Export-Csv -NoTypeInformation -Path $outputPath
    Write-Host "-------------------------------------"
    Write-Host "MFA report saved to: $outputPath"
}

# Show Summary Window if selected
if($ShowSummaryWindow -eq $true){ Show-SummaryInNotepad -SummaryText $summary }

# Check if user selected to disconnect session
if ($DisconnectLater -eq $true) {
    Write-Host "Disconnecting Microsoft Graph session..." -ForegroundColor Yellow
    Disconnect-MgGraph | Out-Null
    # clear tokens
    Clear-MgContextCache
    Write-Host "Session disconnected successfully." -ForegroundColor Green
} else {
    Write-Warning "You are still connected to Microsoft Graph."
}


