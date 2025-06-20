<#
Author: Stewart Thomas | Jackson Thornton Technologies
Description: Retrieve and export results of Microsoft 365 Multi-factor Authentication
Last Updated: 6/5/2025
Modules: Microsoft.Graph.Authentication
         Microsoft.Graph.Identity.DirectoryManagement
         Microsoft.Graph.Beta.Identity.SignIns
         Microsoft.Graph.Users
         Microsoft.Graph.Beta.Users
         Microsoft.Graph.Beta.Identity.SignIns
         Microsoft.Graph.Mail
Scopes:  User.Read.All
         UserAuthenticationMethod.Read.All
         Policy.Read.All
         Directory.Read.All
         Domain.Read.All
         Organization.Read.All
         MailboxSettings.Read
#>

#---------------------------------------------------------------------------------------------#
# Variable and function declarations
#---------------------------------------------------------------------------------------------#
param(
    [string]$ExportPath = (Get-Location)
)

$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm"
$newline = [Environment]::NewLine
$results = @()
$ProcessedUserCount = 0
$globalAdmins = @()

function Maximize-PowerShellWindow {
    # Maximizes the PowerShell window for better visibility.
    if ($psISE) {
        Write-Host "Detected PowerShell ISE. Launching regular PowerShell..."
        Start-Process -FilePath "powershell.exe" -ArgumentList "-NoExit", "-File `"$PSCommandPath`"" -WindowStyle Normal
        exit
    }
    Clear-Host
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
    $WinHandle = [WinAPI]::GetForegroundWindow()
    [WinAPI]::ShowWindow($WinHandle, 3) | Out-Null
}

function Clear-MgContextCache {
    # Removes cached Microsoft Graph tokens if not currently connected.
    if (-not (Get-MgContext)) {
        $cacheFolder = "$env:LOCALAPPDATA\.IdentityService"
        if (Test-Path $cacheFolder) {
            Remove-Item $cacheFolder -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Check_Modules {
    # Checks for and installs required PowerShell modules and providers.
    $newline
    $UserChoice = Get-YesNoInput "First, do you need to check for missing Powershell modules?"
    if ($UserChoice -eq $true) {
        Write-Host "`nChecking for required prerequisites and Microsoft Graph modules..." -ForegroundColor Cyan
        
        # Prerequisite checks (PowerShellGet, NuGet, PSGallery)
        if (-not (Get-Module -ListAvailable -Name PowerShellGet)) {
            Write-Host "PowerShellGet module is required. Installing..." -ForegroundColor Yellow
            Install-Module -Name PowerShellGet -Force -Scope CurrentUser
        } else {
            Write-Host "PowerShellGet module is already installed." -ForegroundColor Green
        }
        if (-not (Get-PackageProvider -ListAvailable | Where-Object Name -eq "NuGet")) {
            Write-Host "NuGet provider is required. Installing..." -ForegroundColor Yellow
            Install-PackageProvider -Name NuGet -Force -Scope CurrentUser
        } else {
            Write-Host "NuGet provider is already installed." -ForegroundColor Green
        }
        $psGallery = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
        if (-not $psGallery) {
            Write-Host "Registering PSGallery repository..." -ForegroundColor Yellow
            Register-PSRepository -Default
        } elseif ($psGallery.InstallationPolicy -ne "Trusted") {
            Write-Host "Setting PSGallery repository as trusted..." -ForegroundColor Yellow
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
        } else {
            Write-Host "PSGallery repository is already registered and trusted." -ForegroundColor Green
        }

        $RequiredModules = @(
            "Microsoft.Graph.Authentication",
            "Microsoft.Graph.Beta.Users",
            "Microsoft.Graph.Identity.SignIns",
            "Microsoft.Graph.Identity.DirectoryManagement",
            "Microsoft.Graph.Users",
            "Microsoft.Graph.Mail",
            "ExchangeOnlineManagement"
        )
        $AvailableModules = Get-Module -ListAvailable
        $MissingModules = $RequiredModules | Where-Object { -not ($AvailableModules | Where-Object Name -eq $_) }
        if ($MissingModules.Count -gt 0) {
            Write-Host "`nInstalling missing modules..." -ForegroundColor Yellow
            Install-Module -Name $MissingModules -Scope CurrentUser -AllowClobber -Force
            Write-Host "`nInstallation completed. Starting modules verification ..." -ForegroundColor Green
        } else {
            Write-Host "`nAll required modules are already installed." -ForegroundColor Green
        }
        $RequiredModules | ForEach-Object {
            if (-not (Get-Module -Name $_)) {
                Import-Module $_ -ErrorAction SilentlyContinue
            }
        }
        Write-Host "`nModules verification complete.$newline" -ForegroundColor Cyan
        Start-Sleep -Seconds 2
        Clear-Host
    } else {
        Write-Host "`nSkipping module verification...$newline" -ForegroundColor Cyan
        Start-Sleep -Seconds 1
        Clear-Host      
    }
}

function ExportResults {
    $outputPath = "$ExportPath\MFA-Report-$TenantName-$timestamp.csv"
    $results | Sort-Object Role, DisplayName | Export-Csv -NoTypeInformation -Path $outputPath
    Write-Host "MFA report saved to: $outputPath$newline"
    Start-Process $outputPath  
}

function Show-SummaryInNotepad {
    param(
        [string]$Title = "MFA Summary Information",
        [string]$SummaryText
    )
    $SummaryFile = "$ExportPath\MFA-Summary-$TenantName-$timestamp.txt"
    $SummaryText | Out-File -Encoding UTF8 $SummaryFile
    Start-Process "notepad.exe" -ArgumentList $SummaryFile
}

function Get-YesNoInput {
    param([string]$Prompt)
    Write-Host $Prompt -NoNewline -ForegroundColor White
    Write-Host " (Y/N): " -NoNewline -ForegroundColor Yellow
    do {
        $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        $response = $key.Character
    } until ($response -match "^[YyNn]$")
    Write-Host $response
    return $response -match "^[Yy]$"
}

function Get-YesNoInputTimeout {
    param(
        [string]$Prompt,
        [int]$Timeout = 10
    )
    Write-Host $Prompt -NoNewline -ForegroundColor White
    Write-Host " (Y/N): " -NoNewline -ForegroundColor Yellow
    $startTime = Get-Date
    $response = $null
    while (((Get-Date) - $startTime).TotalSeconds -lt $Timeout) {
        if ([console]::KeyAvailable) {
            $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            $response = $key.Character
            if ($response -match "^[YyNn]$") {
                Write-Host $response
                return $response -match "^[Yy]$"
            }
        }
        Start-Sleep -Milliseconds 100
    }
    Write-Host "`nTime's up! No response received." -ForegroundColor Yellow
    return $true
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
        $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        $choice = $key.Character
    } until ($choice -match "^[1-4]$")
    Write-Host $choice
    return [int]::Parse($choice)
}

function Get-FilterConfiguration {
    $FilterMode = Get-FilterModeInput
    $config = [PSCustomObject]@{
        FilterMode           = $FilterMode
        UseDefaultFilter     = $false
        IncludeGlobalAdmins  = $false
        IncludeLicensed      = $false
        IncludeMFADisabled   = $false
        IncludePeruserMFA    = $false
        IncludeSigninAllowed = $false
        IncludeHasMailbox    = $false
    }
    if ($FilterMode -eq 2 -or $FilterMode -eq 3 -or $FilterMode -eq 4) {
        if ($FilterMode -eq 4) {
            $config.UseDefaultFilter     = $true
            $config.IncludeGlobalAdmins  = $true
            $config.IncludeLicensed      = $true
            $config.IncludeSigninAllowed = $true
            Write-Host ""
        } else {
            Write-Host ""
            $config.IncludeGlobalAdmins  = Get-YesNoInput "Include Global Admins?"
            $config.IncludeLicensed      = Get-YesNoInput "Include Licensed Users?"
            $config.IncludeHasMailbox    = Get-YesNoInput "Include Users with Exchange Mailbox?"
            $config.IncludeMFADisabled   = Get-YesNoInput "Include Users with no MFA?"
            $config.IncludePeruserMFA    = Get-YesNoInput "Include Users with Per-user MFA NOT disabled?"
            $config.IncludeSigninAllowed = Get-YesNoInput "Include Users allowed to sign-in?"
            Write-Host ""
        }
    }
    return $config
}

function Connect_MgGraph {
    $Scopes = @(
        "User.Read.All",
        "UserAuthenticationMethod.Read.All",
        "Policy.Read.All",
        "Directory.Read.All",
        "Domain.Read.All",
        "Organization.Read.All",
        "MailboxSettings.Read"
    )
    $MgContext = Get-MgContext
    if ($MgContext) {
        # Already connected, ask if user wants to disconnect and switch accounts
        $TenantInfo = Get-MgOrganization
        if ($TenantInfo) {
            $script:TenantDomain = ($TenantInfo.VerifiedDomains | Where-Object { $_.IsInitial -eq $true } | Select-Object -ExpandProperty Name)
            Write-Host "Connected as: $($MgContext.Account) to $TenantDomain$newline" -ForegroundColor Cyan
        }
        $disconnectConfirm = Get-YesNoInput "Do you want to disconnect and sign in with a different account?"
        if ($disconnectConfirm) {
            Write-Host "Disconnecting Microsoft Graph session..." -ForegroundColor Yellow
            Disconnect-MgGraph | Out-Null
            $MgContext = $null
        } else {
            Write-Host ""
            Write-Host "Continuing with the current authentication session." -ForegroundColor Cyan
            Write-Host ""
            return
        }
    }
    # If not connected (either first run or after disconnect), prompt for authentication
    if (-not $MgContext) {
        Write-Host "Connecting to Microsoft Graph...$newline" -ForegroundColor Cyan
        Connect-MgGraph -Scopes $Scopes -NoWelcome -ContextScope Process
        $MgContext = Get-MgContext
        if (-not $MgContext) {
            Write-Host "Microsoft Graph connection failed." -ForegroundColor Red
            $retry = Get-YesNoInput "Try to connect again?"
            if ($retry) {
                Write-Host "Retrying connection..." -ForegroundColor Yellow
                Connect-MgGraph -Scopes $Scopes -NoWelcome -ContextScope Process
                $MgContext = Get-MgContext
                if (-not $MgContext) {
                    Write-Host "Microsoft Graph connection failed again. Exiting." -ForegroundColor Red
                    exit 1
                }
            } else {
                exit 1
            }
        }
    }
}

function Get-UserMfaDetails {
    param (
        $User,
        $skuTable,
        $globalAdmins
    )
    $Name = $User.DisplayName
    $UPN = $User.UserPrincipalName
    $UserId = $User.Id
    $isGlobalAdmin = $User.Id -in $globalAdmins
    $SigninStatus = if ($User.AccountEnabled) { "Allowed" } else { "Blocked" }
    $LicenseStatus = if ($User.AssignedLicenses.Count -ne 0) { "Licensed" } else { "Unlicensed" }
    $skuNames = @()
    if ($LicenseStatus -eq "Licensed") {
        foreach ($lic in $User.AssignedLicenses) {
            if ($lic.SkuId) {
                $skuName = ($skuTable | Where-Object { $_.GUID -eq $lic.SkuId }).Product_Display_Name
                if ($skuName) { $skuNames += $skuName }
            }
        }
    }
    $uniqueLicenseNames = $skuNames | Select-Object -Unique

    try {
        [array]$MFAData = Get-MgBetaUserAuthenticationMethod -UserId $UserId -ErrorAction Stop
    } catch {
        #Write-Verbose "Failed to get MFA methods for $Name: $_"
        Write-Verbose ("Failed to get MFA methods for {0}: {1}" -f $Name, $_)
        $MFAData = @()
    }
    $AuthenticationMethod = @()
    $MFAMethodsCount = 0
    $FilteredMFAData = $MFAData | Where-Object { $_.AdditionalProperties["@odata.type"] -ne "#microsoft.graph.passwordAuthenticationMethod" }
    foreach ($MFA in $FilteredMFAData) {
        switch ($MFA.AdditionalProperties["@odata.type"]) {
            "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" {
                $AuthMethod = 'AuthenticatorApp'
                $AuthMethodDetails = $MFA.AdditionalProperties["displayName"]
            }
            "#microsoft.graph.phoneAuthenticationMethod" {
                $AuthMethod = 'PhoneAuthentication'
                $AuthMethodDetails = $MFA.AdditionalProperties["phoneType", "phoneNumber"]
            }
            "#microsoft.graph.fido2AuthenticationMethod" {
                $AuthMethod = 'Passkeys(FIDO2)'
                $AuthMethodDetails = $MFA.AdditionalProperties["model"]
            }
            "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod" {
                $AuthMethod = 'WindowsHelloForBusiness'
                $AuthMethodDetails = $MFA.AdditionalProperties["displayName"]
            }
            "#microsoft.graph.emailAuthenticationMethod" {
                $AuthMethod = 'EmailAuthentication'
                $AuthMethodDetails = $MFA.AdditionalProperties["emailAddress"]
            }
            "#microsoft.graph.temporaryAccessPassAuthenticationMethod" {
                $AuthMethod = 'TemporaryAccessPass'
                $AuthMethodDetails = 'Access pass lifetime (minutes): ' + $MFA.AdditionalProperties["lifetimeInMinutes"]
            }
            "#microsoft.graph.passwordlessMicrosoftAuthenticatorAuthenticationMethod" {
                $AuthMethod = 'PasswordlessMSAuthenticator'
                $AuthMethodDetails = $MFA.AdditionalProperties["displayName"]
            }
            "#microsoft.graph.softwareOathAuthenticationMethod" {
                $AuthMethod = 'SoftwareOath'
                $AuthMethodDetails = $MFA.id
            }
            default { continue }
        }
        $AuthenticationMethod += $AuthMethod
        if ($AuthMethodDetails) { $MFAMethodsCount++ }
    }
    $AuthenticatorAppMethods = $MFAData | Where-Object { $_.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" } | ForEach-Object { $_.AdditionalProperties["displayName"] }
    $HelloMethods = $MFAData | Where-Object { $_.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod" } | ForEach-Object { $_.AdditionalProperties["displayName"] }
    $MFAPhoneDetail = $MFAData | Where-Object { $_.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.phoneAuthenticationMethod" } | ForEach-Object { $_.AdditionalProperties["phoneNumber"] }
    $MFAOathDetail = $MFAData | Where-Object { $_.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.softwareOathAuthenticationMethod" } | ForEach-Object { $_.Id }
    $EmailMethodDetails = $MFAData | Where-Object { $_.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.emailAuthenticationMethod" } | ForEach-Object { $_.AdditionalProperties["emailAddress"] }
    $PasswordlessMethods = $MFAData | Where-Object { $_.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.passwordlessMicrosoftAuthenticatorAuthenticationMethod" } | ForEach-Object { $_.AdditionalProperties["displayName"] }
    $FidoMethods = $MFAData | Where-Object { $_.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.fido2AuthenticationMethod" } | ForEach-Object { $_.AdditionalProperties["model"] }
    $TempAccessMethods = $MFAData | Where-Object { $_.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.temporaryAccessPassAuthenticationMethod" } | ForEach-Object { $_.AdditionalProperties["lifetimeInMinutes"] }
    $AuthenticationMethod = $AuthenticationMethod | Sort-Object | Get-Unique
    $AuthenticationMethods = $AuthenticationMethod -join ","

    try {
        $DefaultMFAUri = "https://graph.microsoft.com/beta/users/$UserId/authentication/signInPreferences"
        $GetDefaultMFAMethod = Invoke-MgGraphRequest -Uri $DefaultMFAUri -Method GET -ErrorAction Stop
        if ($GetDefaultMFAMethod.userPreferredMethodForSecondaryAuthentication) {
            $MFAMethodisDefault = $GetDefaultMFAMethod.userPreferredMethodForSecondaryAuthentication
            switch ($MFAMethodisDefault) {
                "push" { $MFAMethodisDefault = "Microsoft authenticator app" }
                "oath" { $MFAMethodisDefault = "Authenticator app or hardware token" }
                "voiceMobile" { $MFAMethodisDefault = "Mobile phone" }
                "voiceAlternateMobile" { $MFAMethodisDefault = "Alternate mobile phone" }
                "voiceOffice" { $MFAMethodisDefault = "Office phone" }
                "sms" { $MFAMethodisDefault = "SMS" }
                "email" { $MFAMethodisDefault = "Email" }
                "passwordless" { $MFAMethodisDefault = "Passwordless Microsoft Authenticator app" }
                "fido2" { $MFAMethodisDefault = "FIDO2 security key" }
                default { $MFAMethodisDefault = "Unknown method" }
            }
        } else {
            $MFAMethodisDefault = "Not Enabled"
        }
    } catch {
        $MFAMethodisDefault = "Not Enabled"
    }

    try {
        $PerUserMFAStatus = @(Invoke-MgGraphRequest -Method GET -Uri "/beta/users/$UserId/authentication/requirements" -ErrorAction Stop).perUserMfaState
    } catch {
        $PerUserMFAStatus = "unknown"
    }

    $StrongMFAMethods = @("Fido2", "SoftwareOath", "PasswordlessMSAuthenticator", "AuthenticatorApp", "WindowsHelloForBusiness", "TemporaryAccessPass")
    $MFAStrength = "Disabled"
    if ($AuthenticationMethod | ForEach-Object { $StrongMFAMethods -contains $_ }) { $MFAStrength = "Strong" }
    if ($MFAStrength -ne "Strong" -and ($AuthenticationMethod -match "PhoneAuthentication|EmailAuthentication")) { $MFAStrength = "Weak" }

   
    $mailboxInfo = Has-ExchangeMailbox -UserPrincipalName $UPN
    return [PSCustomObject]@{
        DisplayName               = $Name
        UserPrincipalName         = $UPN
        'Sign-in Status'          = $SigninStatus
        Role                      = if ($isGlobalAdmin) { "Global Admin" } else { "User" }
        'License Status'          = $LicenseStatus
        'License Names'           = if ($uniqueLicenseNames) { $uniqueLicenseNames -join ', ' } else { "None" }
        'Has Mailbox'             = $mailboxInfo.HasMailbox
        'Mailbox Type'            = $mailboxInfo.MailboxType
        'Per-user MFA Status'     = $PerUserMFAStatus
        'MFA Strength'            = $MFAStrength
        'MFA Method Count'        = $MFAMethodsCount
        'Default MFA Method'      = $MFAMethodisDefault
        'Enabled MFA Methods'     = $AuthenticationMethods
        'MS Authenticator App'    = $($AuthenticatorAppMethods -join ', ')
        'Authentication Phone'    = $($MFAPhoneDetail -join ', ')
        'Email Methods'           = $($EmailMethodDetails -join ', ')
        'Software Methods'        = $($MFAOathDetail -join ', ')
        'Hello for Business'      = $($HelloMethods -join ', ')
        'Passwordless Methods'    = $($PasswordlessMethods -join ', ')
        'Fido_ Methods'           = $($FidoMethods -join ', ')
        'Temp Access Methods'     = $($TempAccessMethods -join ', ')
    }


}

#---------------------------------------------------------------------------------------------#
# Script execution
#---------------------------------------------------------------------------------------------#

Maximize-PowerShellWindow
Check_Modules
Clear-MgContextCache

$menutext = @"
---------------------------------------

Welcome to the Stew's fancy MFA script!

The results will be displayed automatically in your default app for CSV files, 
and a summary will be shown in Notepad.

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

----------------------------------------

"@
Write-Output $menutext

# Prompt user for filter configuration (which users to include in the report)
$filterConfig = Get-FilterConfiguration
Write-Host $newline"File path will be: $ExportPath" -ForegroundColor Cyan $newline

# Connect to Microsoft Graph and prompt for account if needed
Connect_MgGraph 

# Prompt user for Exchange mailbox status retrieval method
$RetrieveMailboxStatus = Get-YesNoInput "Do you want to retrieve Exchange mailbox status using Exchange Online (most accurate, requires separate login)"
if($RetrieveMailboxStatus -eq $false) {
    Write-Host "Will retrieve mailbox status using MgGraph, but might not be as accurate as Exchange Online module.$newline" -ForegroundColor Yellow
}

if ($RetrieveMailboxStatus) {
    # Connect to Exchange Online
    try {
        Connect-ExchangeOnline -ShowBanner:$false | Out-Null
    } catch {
        Write-Host "Failed to connect to Exchange Online: $_" -ForegroundColor Red
        exit 1
    }
    function Has-ExchangeMailbox {
        param([string]$UserPrincipalName)
        try {
            $mailbox = Get-Mailbox -Identity $UserPrincipalName -ErrorAction Stop
            if ($mailbox) {
                return [PSCustomObject]@{
                    HasMailbox = $true
                    MailboxType = $mailbox.RecipientTypeDetails
                }
            } else {
                return [PSCustomObject]@{
                    HasMailbox = $false
                    MailboxType = $null
                }
            }
        } catch {
            return [PSCustomObject]@{
                HasMailbox = $false
                MailboxType = $null
            }
        }
    }
} else {
    function Has-ExchangeMailbox {
        param([string]$UserPrincipalName)
        try {
            $user = Get-MgUser -UserId $UserPrincipalName -Property mail -ErrorAction Stop
            if ($user.mail) {
                return $true
            } else {
                return $false
            }
        } catch {
            return $false
        }
    }
}


# Download the latest Microsoft SKU reference file
$csvUrl = "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv"
$csvPath = "$env:TEMP\LicenseNames.csv"
try {
    Invoke-WebRequest -Uri $csvUrl -OutFile $csvPath -ErrorAction Stop
    $skuTable = Import-Csv $csvPath
} catch {
    Write-Host "Failed to download or import the SKU reference file: $_" -ForegroundColor Red
    exit 1
}

# Retrieve all users from Microsoft Graph to begin processing
try {
    $users = Get-MgUser -All -Property "Id,DisplayName,UserPrincipalName,UserType,AccountEnabled,AssignedLicenses" -ErrorAction Stop |
        Where-Object { $_.UserType -eq "Member" } | Sort-Object DisplayName
} catch {
    Write-Host "Failed to retrieve users from Microsoft Graph: $_" -ForegroundColor Red
    exit 1
}

# Determine if user is a Global Admin
try {
    $roles = Get-MgDirectoryRole -ErrorAction Stop
    $globalAdminRole = $roles | Where-Object { $_.DisplayName -eq "Global Administrator" }
    if ($globalAdminRole) {
        $globalAdmins = @(Get-MgDirectoryRoleMember -DirectoryRoleId $globalAdminRole.Id -ErrorAction Stop | ForEach-Object { $_.Id })
    } else {
        $globalAdmins = @()
    }
} catch {
    Write-Host "Failed to retrieve directory roles or global admins: $_" -ForegroundColor Red
    $globalAdmins = @()
}

# Check if Security Defaults are enabled
try {
    $SecurityDefaultsEnabled = Get-MgPolicyIdentitySecurityDefaultEnforcementPolicy -ErrorAction Stop | Select-Object -ExpandProperty IsEnabled
} catch {
    $SecurityDefaultsEnabled = $false
    Write-Host "Could not determine Security Defaults status." -ForegroundColor Yellow
}
if ($SecurityDefaultsEnabled -eq $true) {
    Write-Host ""
    Write-Host "Security Defaults are ENABLED$newline" -ForegroundColor Green
} else {
    Write-Host ""
    Write-Host "Security Defaults are DISABLED$newline" -ForegroundColor Red
}

# Check for Conditional Access
try {
    $conditionalAccessPolicies = Get-MgIdentityConditionalAccessPolicy -ErrorAction Stop
} catch {
    $conditionalAccessPolicies = $null
    Write-Host "Could not retrieve Conditional Access policies." -ForegroundColor Yellow
}
if ($conditionalAccessPolicies -and $conditionalAccessPolicies.Count -gt 0) {
    Write-Host "Conditional Access policies found:" -ForegroundColor Green
    $conditionalAccessPolicies | Format-Table DisplayName, State
} else {
    Write-Host "No Conditional Access policies found.$newline" -ForegroundColor Yellow
}

$total = $users.Count
Write-Host "There are $total Users in this tenant$newline"


#----------------------------------------------------------------------------------------------#
# Start processing user account and MFA details
#----------------------------------------------------------------------------------------------#

$ProcessedUserCount = 0
$results = @()
foreach ($user in $users) {
    $ProcessedUserCount++
    $PercentComplete = [math]::Floor(($ProcessedUserCount / $total) * 100)
    Write-Progress -Activity "Processing user: $ProcessedUserCount - Processing $($user.DisplayName)" -PercentComplete ([Math]::Min(100, $PercentComplete))
    $userResult = Get-UserMfaDetails -User $user -skuTable $skuTable -globalAdmins $globalAdmins
    $results += $userResult

    # Check if the user has an Exchange mailbox
    #$hasMailbox = Has-ExchangeMailbox -UserPrincipalName $user.UserPrincipalName

    # Detailed per-user output 
    Write-Host "###########################"
    Write-Host "[$ProcessedUserCount/$total]: $($userResult.DisplayName)"
    if ($userResult.Role -eq "Global Admin") { Write-Host "Role: Global Admin" -ForegroundColor Yellow } else { Write-Host "Role: User" }
    Write-Host "Sign-in Status: $($userResult.'Sign-in Status')"
    Write-Host "License Status: $($userResult.'License Status')"
    if ($userResult.'License Names' -ne "None") { Write-Host "License Names: $($userResult.'License Names')" }
    if ($userResult.'Has Mailbox') { Write-Host "Has Mailbox: True ($($userResult.'Mailbox Type'))" -ForegroundColor Cyan }
    Write-Host "Per-user MFA Status: $($userResult.'Per-user MFA Status')"
    if ($userResult.'MFA Strength' -eq "Disabled") { Write-Host "MFA Strength: $($userResult.'MFA Strength')" -ForegroundColor Red }
    if ($userResult.'MFA Strength' -eq "Strong")   { Write-Host "MFA Strength: $($userResult.'MFA Strength')" -ForegroundColor Green }
    if ($userResult.'MFA Strength' -eq "Weak")     { Write-Host "MFA Strength: $($userResult.'MFA Strength')" -ForegroundColor Yellow }
    Write-Host "Default MFA Method: $($userResult.'Default MFA Method')"
    if ($userResult.'MFA Methods')           { Write-Host "$($userResult.'MFA Method Count') MFA Methods: $($userResult.'MFA Methods')" }
    if ($userResult.'MS Authenticator App')  { Write-Host "Authenticator Apps: $($userResult.'MS Authenticator App')" }
    if ($userResult.'Passwordless Methods')  { Write-Host "Passwordless Authenticator methods: $($userResult.'Passwordless Methods')" }
    if ($userResult.'Email Methods')         { Write-Host "Email methods: $($userResult.'Email Methods')" }
    if ($userResult.'Authentication Phone')  { Write-Host "Phone methods: $($userResult.'Authentication Phone')" }
    if ($userResult.'Hello for Business')    { Write-Host "Hello for Business methods: $($userResult.'Hello for Business')" }
    if ($userResult.'Software Methods')      { Write-Host "Software Oath methods: $($userResult.'Software Methods')" }
    if ($userResult.'Fido_ Methods')         { Write-Host "Fido methods: $($userResult.'Fido_ Methods')" }
    if ($userResult.'Temp Access Methods')   { Write-Host "Temp access methods: $($userResult.'Temp Access Methods')" }
    Write-Host ""

}
Write-Progress -Activity "Processing user: $ProcessedUserCount" -Completed

#----------------------------------------------------------------------------------------------#
# Filtering
#----------------------------------------------------------------------------------------------#
$ApplyFiltering = $filterConfig.IncludeGlobalAdmins -or $filterConfig.IncludeLicensed -or $filterConfig.IncludeMFADisabled -or $filterConfig.IncludePeruserMFA -or $filterConfig.IncludeSigninAllowed
if ($ApplyFiltering) {
    if ($filterConfig.FilterMode -eq 2) {
        $results = $results | Where-Object {
            ($filterConfig.IncludeGlobalAdmins -and $_.Role -eq "Global Admin") -or
            ($filterConfig.IncludeLicensed -and $_.'License Status' -eq "Licensed") -or
            ($filterConfig.IncludeMFADisabled -and ($_. 'MFA Strength' -eq "Disabled" -or $_.'MFA Method Count' -eq 0)) -or
            ($filterConfig.IncludePeruserMFA -and ($_. 'Per-user MFA Status' -ne "disabled")) -or
            ($filterConfig.IncludeSigninAllowed -and $_.'Sign-in Status' -eq "Allowed") -or
            ($filterConfig.IncludeHasMailbox -and $_.'Has Mailbox' -eq $true)
        }
    } elseif ($filterConfig.FilterMode -eq 3) {
        $results = $results | Where-Object {
            (!$filterConfig.IncludeGlobalAdmins -or $_.Role -eq "Global Admin") -and
            (!$filterConfig.IncludeLicensed -or $_.'License Status' -eq "Licensed") -and
            (!$filterConfig.IncludeMFADisabled -or ($_. 'MFA Strength' -eq "Disabled" -or $_.'MFA Method Count' -eq 0)) -and
            (!$filterConfig.IncludePeruserMFA -or ($_. 'Per-user MFA Status' -ne "disabled")) -and
            (!$filterConfig.IncludeSigninAllowed -or $_.'Sign-in Status' -eq "Allowed") -and
            (!$filterConfig.IncludeHasMailbox -or $_.'Has Mailbox' -eq $true) 
        }
    } elseif ($filterConfig.FilterMode -eq 4) {
            $results = $results | Where-Object {
            (!$filterConfig.IncludeGlobalAdmins -or $_.Role -eq "Global Admin") -or
            (!$filterConfig.IncludeLicensed -or $_.'License Status' -eq "Licensed") -and
            (!$filterConfig.IncludeSigninAllowed -or $_.'Sign-in Status' -eq "Allowed")
        }

    }

}
$AfterFilterCount = @($results).Count

#----------------------------------------------------------------------------------------------#
# Build the summary information for the report
#----------------------------------------------------------------------------------------------#
$TenantDomain = (Get-MgDomain | Where-Object {$_.isInitial}).Id
$TenantName = $TenantDomain -replace "\.onmicrosoft\.com",""
$summary = ""
$summary += "MFA Summary for: $TenantName$newline"
$summary += "$timestamp $newline$newline"
if ($SecurityDefaultsEnabled -eq $true) { 
    $summary += "Security Defaults are ENABLED$newline" 
} else { 
    $summary += "Security Defaults are DISABLED$newline" 
}
if ($conditionalAccessPolicies) {
    $summary += "Conditional Access is enabled. Policies found:`n"
    $summary += $conditionalAccessPolicies | Format-Table DisplayName, State | Out-String
} else {
    $summary += "No Conditional Access policies found.`n"
    $summary += $newline
}

$summary += @"
##################################################
Filtered Configuration Summary
##################################################
$newline
"@
if($filterConfig.FilterMode -eq 1) { 
    $summary += "No filtering applied.$newline$newline"
} else {
    if($filterConfig.FilterMode -eq 2){ $summary += "Filter Mode: Flexible$newline$newline" }
    if($filterConfig.FilterMode -eq 3){ $summary += "Filter Mode: Strict$newline$newline" }
    $summary += "Default Filter: $($filterConfig.UseDefaultFilter)$newline"
    $summary += "IncludeGlobal: $($filterConfig.IncludeGlobalAdmins)$newline"
    $summary += "IncludeLicensed: $($filterConfig.IncludeLicensed)$newline"
    $summary += "IncludeHasMailbox: $($filterConfig.IncludeHasMailbox)$newline"
    $summary += "IncludeMFADisabled: $($filterConfig.IncludeMFADisabled)$newline"
    $summary += "IncludePeruserMFA: $($filterConfig.IncludePeruserMFA)$newline"
    $summary += "IncludeSigninAllowed: $($filterConfig.IncludeSigninAllowed)$newline$newline"
}

$summary += @"
##################################################
Filtered Users Summary
##################################################
$newline
"@
$summary += "Total users processed: $total $newline"
$summary += "Users included in report: $AfterFilterCount $newline"
$summary += "Users skipped due to filters: $($total - $AfterFilterCount) $newline$newline"

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


# Count unique license types
$allLicenseNames = $results | Where-Object { $_.'License Names' -ne "None" } | ForEach-Object { $_.'License Names' -split ',\s*' }
$licenseTypeCounts = $allLicenseNames | Group-Object | Sort-Object Count -Descending

if ($licenseTypeCounts.Count -gt 0) {
    $licenseSummary += ($licenseTypeCounts | Format-Table Name, Count -AutoSize | Out-String)
    $summary += @"
##################################################
License Type Breakdown
##################################################

"@
    $summary += $licenseSummary + "`r`n"
}


$MailboxCount      = ($results | Where-Object { $_.'Has Mailbox' -eq $true } | Measure-Object).Count
$NoMailboxCount    = ($results | Where-Object { $_.'Has Mailbox' -eq $false } | Measure-Object).Count
$UserMailboxCount    = ($results | Where-Object { $_.'Mailbox Type' -eq 'UserMailbox' } | Measure-Object).Count
$SharedMailboxCount  = ($results | Where-Object { $_.'Mailbox Type' -eq 'SharedMailbox' } | Measure-Object).Count

$summary += @"
##################################################
Exchange Mailbox Status Summary
##################################################

Users with Mailbox: $MailboxCount
  Total UserMailbox: $UserMailboxCount
  Total SharedMailbox: $SharedMailboxCount
Users without Mailbox: $NoMailboxCount
$newline
"@



$summary += @"
##################################################
Users MFA Summary
##################################################
$newline
"@
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

$noMfaUsers = @($results | Where-Object { $_.'MFA Method Count' -eq 0 })
$noMfaTable = $noMfaUsers | Select-Object DisplayName, UserPrincipalName | Format-Table -AutoSize | Out-String
if ($noMfaUsers.Count -gt 0) {
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

if ($usersWithMfa = @($results | Where-Object { $_.'MFA Method Count' -gt 0 })) {
    $mfaMethodsArray = $usersWithMfa | ForEach-Object {
        $_.'Enabled MFA Methods' -split ',\s*'
    }
    $mfaBreakdown = $mfaMethodsArray | Group-Object | Sort-Object Count -Descending
    $mfaBreakdownString = $mfaBreakdown | Format-Table Name, Count -AutoSize | Out-String
    $summary += @"
##################################################
MFA Method Breakdown
##################################################
"@
    $summary += $mfaBreakdownString + "`r`n"
}

@"
---------------------------------------------------------------------
Script has completed analyzing users
---------------------------------------------------------------------$newline
"@

#----------------------------------------------------------------------------------------------#
# Wrap up
#----------------------------------------------------------------------------------------------#




# Ask if the user wants to disconnect from Microsoft Graph
$result = Get-YesNoInputTimeout -Prompt "Disconnect from MgGraph? I'll wait 10 seconds, then disconnect automatically."
if ($result) {
    Write-Host "Disconnecting from Microsoft Graph and clearing tokens ...$newline" -ForegroundColor Cyan
    Disconnect-MgGraph | Out-Null
    Start-Sleep -Seconds 1  # Allow time for disconnection
    Clear-MgContextCache # Run function to remove any cached tokens
    try {
        Disconnect-ExchangeOnline -Confirm:$false
    } catch {}
} else {
    Write-Host "Maintaining connection to MgGraph.$newline" -ForegroundColor Cyan
}

# Show a summary in Notepad
Show-SummaryInNotepad -SummaryText $summary

# Only export and display results if there are results
$results = @($results)  # <-- This ensures $results is always an array

if ($results.Count -gt 0) {
    ExportResults # Show results in Excel or the default app
} else {
    Write-Host "No results found. Not exporting an empty CSV file." -ForegroundColor Yellow
}
