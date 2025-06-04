<#
Author: Stewart Thomas | Jackson Thornton Technologies
Description: Extract and export results of Microsoft 365 Multi-factor Authentication
Last Updated: 6/3/2025
Modules: Microsoft.Graph.Authentication, Microsoft.Graph.Identity.DirectoryManagement
Modules: Microsoft.Graph.Users, Microsoft.Graph.Beta.Users, Microsoft.Graph.Beta.Identity.SignIns
Scopes: User.Read.All, UserAuthenticationMethod.Read.All, Policy.Read.All, Domain.Read.All 
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
    if (-not (Get-MgContext)) {
        $cacheFolder = "$env:LOCALAPPDATA\.IdentityService"
        if (Test-Path $cacheFolder) {
            Remove-Item $cacheFolder -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Check_Modules {
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
        $AvailableModules = Get-Module -ListAvailable
        $MissingModules = $RequiredModules | Where-Object { -not ($AvailableModules | Where-Object Name -eq $_) }
        if ($MissingModules.Count -gt 0) {
            Write-Host "`nInstalling missing modules..." -ForegroundColor Yellow
            Install-Module -Name $MissingModules -Scope CurrentUser -AllowClobber -Force
            Write-Host "`nInstallation completed." -ForegroundColor Magenta
        } else {
            Write-Host "`nAll required modules are already installed." -ForegroundColor Green
        }
        $RequiredModules | ForEach-Object {
            if (-not (Get-Module -Name $_)) {
                Import-Module $_ -ErrorAction SilentlyContinue
            }
        }
        Write-Host "`nModules verification complete.$newline" -ForegroundColor Cyan
    } else {
        Write-Host "`nSkipping module verification...$newline" -ForegroundColor Cyan
    }
}

function Show-SummaryInNotepad {
    param(
        [string]$Title = "MFA Summary Information",
        [string]$SummaryText
    )
    $TempFile = "$env:TEMP\Summary.txt"
    $SummaryText | Out-File -Encoding UTF8 $TempFile
    Start-Process "notepad.exe" -ArgumentList $TempFile
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
    }
    if ($FilterMode -eq 2 -or $FilterMode -eq 3 -or $FilterMode -eq 4) {
        if ($FilterMode -eq 4) {
            $config.UseDefaultFilter    = $true
            $config.FilterMode          = 2
            $config.IncludeGlobalAdmins = $true
            $config.IncludeLicensed     = $true
        } else {
            $config.IncludeGlobalAdmins  = Get-YesNoInput "Include Global Admins?"
            $config.IncludeLicensed      = Get-YesNoInput "Include Licensed Users?"
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
        "Organization.Read.All"
    )
    $MgContext = Get-MgContext
    if ($MgContext) {
        $TenantInfo = Get-MgOrganization
        if ($TenantInfo) {
            $script:TenantDomain = ($TenantInfo.VerifiedDomains | Where-Object { $_.IsInitial -eq $true } | Select-Object -ExpandProperty Name)
            Write-Host "Connected successfully as: $($MgContext.Account) to $TenantDomain$newline" -ForegroundColor Cyan
        } else {
            Write-Host "Unable to retrieve tenant information." -ForegroundColor Red
        }
        $disconnectConfirm = Get-YesNoInput "Do you want to disconnect and sign in with a different account?"
        if ($disconnectConfirm -eq $true) {
            Write-Host "Disconnecting Microsoft Graph session..." -ForegroundColor Yellow
            Disconnect-MgGraph | Out-Null
            $MgContext = $null
            Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
            Connect-MgGraph -Scopes $Scopes
            $MgContext = Get-MgContext
        } else {
            Write-Host "$newline Continuing with the current authentication session." -ForegroundColor Cyan
            Write-Host ""
            return
        }
    } else {
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
        Connect-MgGraph -Scopes $Scopes
        $MgContext = Get-MgContext
    }
    if ($MgContext) {
        $TenantInfo = Get-MgOrganization
        if ($TenantInfo) {
            $TenantDomain = ($TenantInfo.VerifiedDomains | Where-Object { $_.IsInitial -eq $true } | Select-Object -ExpandProperty Name)
        } else {
            Write-Host "Unable to retrieve tenant information." -ForegroundColor Red
        }
        Write-Host "Connected successfully as: $($MgContext.Account) to $TenantDomain$newline" -ForegroundColor Cyan
    } else {
        Write-Host "Microsoft Graph connection failed." -ForegroundColor Red
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
        #if ($AuthMethodDetails -ne $null) { $MFAMethodsCount++ }
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

   
    return [PSCustomObject]@{
        DisplayName               = $Name
        UserPrincipalName         = $UPN
        Role                      = if ($isGlobalAdmin) { "Global Admin" } else { "User" }
        'License Status'          = $LicenseStatus
        'License Names'           = if ($uniqueLicenseNames) { $uniqueLicenseNames -join ', ' } else { "None" }
        'Sign-in Status'          = $SigninStatus
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
Clear-MgContextCache

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

Check_Modules
$filterConfig = Get-FilterConfiguration
$ExportResults = (Get-YesNoInput "Export Results to CSV?")
if($ExportResults -eq $true) { Write-Host "File path will be: $ExportPath" -ForegroundColor Cyan $newline }
$ShowSummaryWindow = (Get-YesNoInput "Would you like to see a summary?")
Write-Host ""
Connect_MgGraph

# Download the latest Microsoft SKU reference file
$csvUrl = "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv"
$csvPath = "$env:TEMP\LicenseNames.csv"
Invoke-WebRequest -Uri $csvUrl -OutFile $csvPath
$skuTable = Import-Csv $csvPath

# Get list of users to begin processing
$users = Get-MgUser -All -Property "Id,DisplayName,UserPrincipalName,UserType,AccountEnabled,AssignedLicenses" | Where-Object { $_.UserType -eq "Member" }

# Determine if user is a Global Admin
$roles = Get-MgDirectoryRole
$globalAdminRole = $roles | Where-Object { $_.DisplayName -eq "Global Administrator" }
if ($globalAdminRole) {
    $globalAdmins = Get-MgDirectoryRoleMember -DirectoryRoleId $globalAdminRole.Id | ForEach-Object { $_.Id }
}

# Check if Security Defaults are enabled
$SecurityDefaultsEnabled = Get-MgPolicyIdentitySecurityDefaultEnforcementPolicy | Select-Object -ExpandProperty IsEnabled
if ($SecurityDefaultsEnabled -eq $true) {
    Write-Host "Security Defaults are ENABLED$newline" -ForegroundColor Green
} else {
    Write-Host "Security Defaults are DISABLED$newline" -ForegroundColor Red
}

# Check for Conditional Access
$conditionalAccessPolicies = Get-MgIdentityConditionalAccessPolicy
if ($conditionalAccessPolicies) {
    Write-Host "Conditional Access is enabled." -ForegroundColor Green "Policies found:"
    $conditionalAccessPolicies | Format-Table DisplayName, State
} else {
    Write-Host "No Conditional Access policies found.$newline"
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

    # Output per-user summary (customize as needed)
    #Write-Host ("[{0}/{1}] {2} ({3}) - MFA: {4}, Methods: {5}" -f $ProcessedUserCount, $total, $userResult.DisplayName, $userResult.UserPrincipalName, $userResult.'MFA Strength', $userResult.'MFA Methods')

    # Detailed per-user output
    Write-Host "###########################"
    Write-Host "[$ProcessedUserCount/$total] Processing: $($userResult.DisplayName)"
    if ($userResult.Role -eq "Global Admin") { Write-Host "Role: Global Admin" -ForegroundColor Yellow } else { Write-Host "Role: User" }
    Write-Host "Sign-in Status: $($userResult.'Sign-in Status')"
    Write-Host "License Status: $($userResult.'License Status')"
    if ($userResult.'License Names') { Write-Host "License Names: $($userResult.'License Names')" }
    Write-Host "Per-user MFA Status: $($userResult.'Per-user MFA Status')"
    if ($userResult.'MFA Strength' -eq "Disabled") { Write-Host "MFA Strength: $($userResult.'MFA Strength')" -ForegroundColor Red }
    if ($userResult.'MFA Strength' -eq "Strong")   { Write-Host "MFA Strength: $($userResult.'MFA Strength')" -ForegroundColor Green }
    if ($userResult.'MFA Strength' -eq "Weak")     { Write-Host "MFA Strength: $($userResult.'MFA Strength')" -ForegroundColor Cyan }
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
            ($filterConfig.IncludeSigninAllowed -and $_.'Sign-in Status' -eq "Allowed")
        }
    } elseif ($filterConfig.FilterMode -eq 3) {
        $results = $results | Where-Object {
            (!$filterConfig.IncludeGlobalAdmins -or $_.Role -eq "Global Admin") -and
            (!$filterConfig.IncludeLicensed -or $_.'License Status' -eq "Licensed") -and
            (!$filterConfig.IncludeMFADisabled -or ($_. 'MFA Strength' -eq "Disabled" -or $_.'MFA Method Count' -eq 0)) -and
            (!$filterConfig.IncludePeruserMFA -or ($_. 'Per-user MFA Status' -ne "disabled")) -and
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

$noMfaUsers = $results | Where-Object { $_.'MFA Method Count' -eq 0 }
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

if ($usersWithMfa = $results | Where-Object { $_.'MFA Method Count' -gt 0 }) {
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

# If the user chose to export results, save to CSV
if($ExportResults -eq $true) {
    $outputPath = "$ExportPath\MFA-Report-$TenantName-$timestamp.csv"
    $results | Sort-Object Role, DisplayName | Export-Csv -NoTypeInformation -Path $outputPath
    Write-Host "MFA report saved to: $outputPath$newline"
}

# Ask if the user wants to disconnect from Microsoft Graph
$result = Get-YesNoInputTimeout -Prompt "Disconnect from MgGraph? I'll wait 10 seconds, then disconnect automatically."
if ($result) {
    Write-Host "Disconnecting from Microsoft Graph and clearing tokens ...$newline" -ForegroundColor Cyan
    Disconnect-MgGraph | Out-Null
    Clear-MgContextCache
} else {
    Write-Host "Maintaining connection to MgGraph.$newline" -ForegroundColor Cyan
}

# If the user chose to show a summary, display it in Notepad
if($ShowSummaryWindow -eq $true){ Show-SummaryInNotepad -SummaryText $summary }

# Display the results in a GridView window
$results | Sort-Object Role, DisplayName | Out-GridView -Title "Microsoft 365 MFA Report"