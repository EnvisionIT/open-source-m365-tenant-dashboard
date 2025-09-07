<##################################################################
 Name: EIT-Open-Source-M365-Tenant-Dashboard

 .SYNOPSIS
 Get inventory of Sharepoint Site Collections, Sites and SubSites

 .DESCRIPTION
 The scripts catalogs all the site collections and SubSites the account has access
 and generate output csv file

 SHAREPOINT ONLINE
    The account used in script needs to be tenant admin

  The following prerequistises
    1. Powershell version 7.4 or above
    2. SharePointPnPPowerShellOnline module needs to be installed
        Install-Module -Name PnP.PowerShell -RequiredVersion 3.1.0

    This PowerShell script is developed and maintained by Envision IT.
    It is published under the MIT License and is intended for use in [brief use case or environment].

.AUTHOR
    Envision IT (https://envisionit.com)
    Contact: info@envisionit.com
.LICENSE
    MIT License

    Copyright (c) 2025 Envision IT
    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.

.VERSION
    1.0.0

.LASTUPDATED
    2025-09-07

.NOTES
    This script is part of the Envision IT open-source initiative.
    Contributions and feedback are welcome via GitHub or email.
##################################################################>

#requires -version 7.4
#Requires -Modules @{ ModuleName="PnP.PowerShell"; ModuleVersion="3.1.0" }

# Get the current timestamp
function Get-TimeStamp {
    return "[{0:MM/dd/yy} {0:HH:mm:ss.fff}] `t" -f (Get-Date)
}

function IIf($If, $Then, $Else) {
    If ($If -IsNot "Boolean") {$_ = $If}
    If ($If) {If ($Then -is "ScriptBlock") {&$Then} Else {$Then}}
    Else {If ($Else -is "ScriptBlock") {&$Else} Else {$Else}}
}

function Read-Host-With-Default {
    Param (
        [Parameter(Mandatory=$true)][string] $Prompt,
        [Parameter()][string] $Default
    )

    $input = Read-Host "$Prompt [$Default]"
    if ([string]::IsNullOrWhiteSpace($input)) {
        return $Default
    } else {
        return $input
    }
}

function Connect-PnPOnlineHelper {
    Param
    (
        [Parameter(Mandatory = $true)][string] $URL
    )

    $newConn = Connect-PnPOnline -ReturnConnection -Url $URL -ClientId $ClientId -Tenant $tenantName -Thumbprint $Thumbprint -InformationAction Ignore

    return $newConn
}

function Write-Config {
    # Create a hashtable with the variables
    $data = @{
        Tenant                  = $tenantName
        SPRootSiteCollectionURL = $SPRootSiteCollectionURL
        ClientId                = $ClientId
        Thumbprint              = $Thumbprint
    }

    # Convert to JSON
    $json = $data | ConvertTo-Json -Depth 3

    # Write to file
    $fileName = "$tenantName.json"
    $json | Out-File -FilePath $fileName -Encoding UTF8
}

function Read-Config {
    $Global:tenantName = Read-Host-With-Default "Tenant (tenantName.onmicrosoft.com)" $tenantName
    $fileName = "$tenantName.json"

    if (Test-Path $fileName) {
        $jsonContent = Get-Content -Path $fileName -Raw | ConvertFrom-Json

        $Global:SPRootSiteCollectionURL = $jsonContent.SPRootSiteCollectionURL
        $Global:ClientId                = $jsonContent.ClientId
        $Global:Thumbprint              = $jsonContent.Thumbprint

        Write-Host "Configuration loaded successfully from '$fileName'."
    }
}

function Process-SitesOnline {
    Param
    (
        [Parameter(Mandatory = $true)][string] $SiteURL,
        [Parameter(Mandatory = $true)][string] $ReportFile
    )

    Try {

        $ConnSite = Connect-PnPOnlineHelper -Url $SiteURL 

        $CurrentWeb = Get-PnPWeb -Includes Webs -Connection $ConnSite -ErrorAction Stop

        [int] $Total = $CurrentWeb.Webs.Count
        [int] $i = 0

        if ($Total -gt 0) {
            $SubWebs = Get-PnPSubWeb -Identity $CurrentWeb -Recurse -Includes "WebTemplate" -Connection $ConnSite
            $Total = $SubWebs.Count

            #Loop throuh all Sub Sites
            foreach($SubWeb in $SubWebs) {
                $i++
                Write-Progress -PercentComplete ($i / ($Total) * 100) -Activity "Processing Subsites $i of $($Total)" -Status "Processing Subsite $($SubWeb.URL)'" -Id 2 -ParentId 1

                # Do Not Root Site again
                if ($SubWeb.ServerRelativeUrl -ne $CurrentWeb.ServerRelativeUrl) {
                    Write-Host "$(Get-TimeStamp) `tProcessing Subsite: $($SubWeb.Url)"

                    $SiteTemplate = ($SiteTemplates | Where-Object { $_.Name -eq $SubWeb.WebTemplate } | Select-Object -First 1).Title
                    If (-not $SiteTemplate) {$SiteTemplate = $SubWeb.WebTemplate}

                    $SiteDetails = New-Object PSObject

                    $SiteDetails | Add-Member NoteProperty 'Site name'($SubWeb.Title)
                    $SiteDetails | Add-Member NoteProperty 'URL'($SubWeb.URL)
                    $SiteDetails | Add-Member NoteProperty 'Site Type'("Subsite")
                    $SiteDetails | Add-Member NoteProperty 'Teams'("")
                    $SiteDetails | Add-Member NoteProperty 'Storage used (GB)'("")
                    $SiteDetails | Add-Member NoteProperty 'Primary admin'($SubWeb.Author.LoginName)
                    $SiteDetails | Add-Member NoteProperty 'Hub'("")
                    $SiteDetails | Add-Member NoteProperty 'Template'($SiteTemplate)
                    $SiteDetails | Add-Member NoteProperty 'Last activity (UTC)'("")
                    $SiteDetails | Add-Member NoteProperty 'Date created'("")
                    $SiteDetails | Add-Member NoteProperty 'Created by'($SubWeb.Author.LoginName)
                    $SiteDetails | Add-Member NoteProperty 'Storage limit (GB)'("")
                    $SiteDetails | Add-Member NoteProperty 'Storage used (%)'("")
                    $SiteDetails | Add-Member NoteProperty 'Microsoft 365 group'("")
                    $SiteDetails | Add-Member NoteProperty 'Files viewed or edited'("")
                    $SiteDetails | Add-Member NoteProperty 'Page views'("")
                    $SiteDetails | Add-Member NoteProperty 'Page visits'("")
                    $SiteDetails | Add-Member NoteProperty 'Files'("")
                    $SiteDetails | Add-Member NoteProperty 'Sensitivity'("")
                    $SiteDetails | Add-Member NoteProperty 'External sharing'("")
                    $SiteDetails | Add-Member NoteProperty 'Sharing Domain Restriction Mode'("")
                    $SiteDetails | Add-Member NoteProperty 'Sharing Allowed Domain List'("")
                    $SiteDetails | Add-Member NoteProperty 'Sharing Blocked Domain List'("")
                    $SiteDetails | Add-Member NoteProperty 'Default Link Permission'("")
                    $SiteDetails | Add-Member NoteProperty 'Default Sharing Link Type'("")
                    $SiteDetails | Add-Member NoteProperty 'External User Expiration In Days'("")
                    $SiteDetails | Add-Member NoteProperty 'Override Tenant Anonymous Link Expiration Policy'("")
                    $SiteDetails | Add-Member NoteProperty 'Override Tenant External User Expiration Policy'("")
                    $SiteDetails | Add-Member NoteProperty 'Allow Downloading Non Web Viewable Files'("")
                    $SiteDetails | Add-Member NoteProperty 'Allow Editing'("")
                    $SiteDetails | Add-Member NoteProperty 'Allow Self Service Upgrade'("")
                    $SiteDetails | Add-Member NoteProperty 'Anonymous Link Expiration In Days'("")
                    $SiteDetails | Add-Member NoteProperty 'Block Download Links File Type'("")
                    $SiteDetails | Add-Member NoteProperty 'Comments On Site Pages Disabled'("")
                    $SiteDetails | Add-Member NoteProperty 'Compatibility Level'("")
                    $SiteDetails | Add-Member NoteProperty 'Conditional Access Policy'("")
                    $SiteDetails | Add-Member NoteProperty 'Default Link To Existing Access'("")
                    $SiteDetails | Add-Member NoteProperty 'Deny Add And Customize Pages'("")
                    $SiteDetails | Add-Member NoteProperty 'Description'("")
                    $SiteDetails | Add-Member NoteProperty 'Disable App Views'("")
                    $SiteDetails | Add-Member NoteProperty 'Disable Company Wide Sharing Links'("")
                    $SiteDetails | Add-Member NoteProperty 'Disable Flows'("")
                    $SiteDetails | Add-Member NoteProperty 'Disable Sharing For Non Owners Status'("")
                    $SiteDetails | Add-Member NoteProperty 'Group Id'("")
                    $SiteDetails | Add-Member NoteProperty 'Hub Site Id'("")
                    $SiteDetails | Add-Member NoteProperty 'Information Segment'("")
                    $SiteDetails | Add-Member NoteProperty 'Limited Access File Type'("")
                    $SiteDetails | Add-Member NoteProperty 'Locale Id'("")
                    $SiteDetails | Add-Member NoteProperty 'Lock Issue'("")
                    $SiteDetails | Add-Member NoteProperty 'Lock State'("")
                    $SiteDetails | Add-Member NoteProperty 'Owner'("")
                    $SiteDetails | Add-Member NoteProperty 'Owner Login Name'("")
                    $SiteDetails | Add-Member NoteProperty 'Owner Name'("")
                    $SiteDetails | Add-Member NoteProperty 'Protection Level Name'("")
                    $SiteDetails | Add-Member NoteProperty 'PWA Enabled'("")
                    $SiteDetails | Add-Member NoteProperty 'Related Group Id'("")
                    $SiteDetails | Add-Member NoteProperty 'Resource Quota'("")
                    $SiteDetails | Add-Member NoteProperty 'Resource Quota Warning Level'("")
                    $SiteDetails | Add-Member NoteProperty 'Resource Usage Average'("")
                    $SiteDetails | Add-Member NoteProperty 'Resource Usage Current'("")
                    $SiteDetails | Add-Member NoteProperty 'Restricted To Geo'("")
                    $SiteDetails | Add-Member NoteProperty 'Sandboxed Code Activation Capability'("")
                    $SiteDetails | Add-Member NoteProperty 'Show People Picker Suggestions For Guest Users'("")
                    $SiteDetails | Add-Member NoteProperty 'Site Defined Sharing Capability'("")
                    $SiteDetails | Add-Member NoteProperty 'Social Bar On Site Pages Disabled'("")
                    $SiteDetails | Add-Member NoteProperty 'Status'("")
                    $SiteDetails | Add-Member NoteProperty 'Storage Quota Type'("")
                    $SiteDetails | Add-Member NoteProperty 'Storage Quota Warning Level'("")
                    $SiteDetails | Add-Member NoteProperty 'Storage Usage Current'("")
                    $SiteDetails | Add-Member NoteProperty 'Webs Count'("")
                    #Export details to CSV File
                    $SiteDetails | Export-CSV $ReportFile -Encoding UTF8 -NoTypeInformation -Append
                }
            }
        }
    }
    Catch {
        write-host -f Red "$(Get-TimeStamp) Error Processing Site: $SiteURL"
        Write-Host -f Red $_.Exception.Message 
   }
}

function Process-SiteCollectionOnline {

    Read-Config
    
    $Global:SPRootSiteCollectionURL = Read-Host-With-Default "Root Sharepoint site collection" $SPRootSiteCollectionURL
    $Global:ClientId = Read-Host-With-Default "Application Client ID" $ClientId
    $Global:Thumbprint = Read-Host-With-Default "Certificate Thumbprint" $Thumbprint

    Write-Config

    # Define the output CSV file path
    $ReportFile = "$PSScriptRoot\Sites.csv"

    # check if Sites.csv already exist
    if (test-path "$ReportFile") {
        remove-item $ReportFile
    }
   
    # connect to sharepoint
    $conn = Connect-PnPOnlineHelper -Url $SPRootSiteCollectionURL

    # Get Tenant Defaults
    $TenantInfo = Get-PnPTenant -Connection $conn
    $DefaultSharingCapability = $TenantInfo.SharingCapability
    $DefaultRequireAnonymousLinksExpireInDays = $TenantInfo.RequireAnonymousLinksExpireInDays
    $DefaultSharingDomainRestrictionMode = $TenantInfo.SharingDomainRestrictionMode
    $DefaultSharingAllowedDomainList = $TenantInfo.SharingAllowedDomainList
    $DefaultSharingBlockedDomainList = $TenantInfo.SharingBlockedDomainList
    $DefaultSharingLinkType = $TenantInfo.DefaultSharingLinkType
    $DefaultLinkPermission = $TenantInfo.DefaultLinkPermission

    # Get all the site collection
    $SiteCollections = Get-PnPTenantSite -Connection $conn

    # Get Templates on the Farm
    $SiteTemplates = Get-PnPWebTemplates -Connection $conn | SELECT ID, Name, Title -Unique

    [int] $Total = $SiteCollections.Count
    [int] $i = 0
    Write-Host "$(Get-TimeStamp) `tTotal Site Collections found: $($Total)"

    foreach ($SiteCollection in $SiteCollections) {
        $i++
        Write-Progress -PercentComplete ($i / ($Total) * 100) -Activity "Processing Site Collections $i of $($Total)" -Status "Processing Site $($SiteCollection.URL)'" -Id 1
        
        if ($SiteCollection.Url.StartsWith($SPRootSiteCollectionURL)) {
            Write-Host "$(Get-TimeStamp) `tProcessing Site Collection: $($SiteCollection.Url)"
            
            $conn = Connect-PnPOnlineHelper -Url $SiteCollection.Url

            $CurrentWeb = Get-PnPWeb -Includes Created,LastItemUserModifiedDate -Connection $conn
            $CurrentSiteCollection = Get-PnPTenantSite -Identity $SiteCollection.Url -Detailed -DisableSharingForNonOwnersStatus -Connection $conn

            if ($CurrentSiteCollection) {
                $SiteTemplate = ($SiteTemplates | Where-Object { $_.Name -eq $SiteCollection.Template } | Select-Object -First 1).Title
                If (-not $SiteTemplate) {$SiteTemplate = $CurrentSiteCollection.Template}

                $SiteDetails = New-Object PSObject

                $SiteDetails | Add-Member NoteProperty 'Site name'($CurrentSiteCollection.Title)
                $SiteDetails | Add-Member NoteProperty 'URL'($CurrentSiteCollection.Url)
                $SiteDetails | Add-Member NoteProperty 'Site Type'("Site Collection")
                $SiteDetails | Add-Member NoteProperty 'Teams'("")
                $SiteDetails | Add-Member NoteProperty 'Storage used (GB)'($CurrentSiteCollection.StorageUsageCurrent -f {0:N2})
                $SiteDetails | Add-Member NoteProperty 'Primary admin'($CurrentSiteCollection.OwnerEmail)
                $SiteDetails | Add-Member NoteProperty 'Hub'($CurrentSiteCollection.IsHubSite)
                $SiteDetails | Add-Member NoteProperty 'Template'($SiteTemplate)
                $SiteDetails | Add-Member NoteProperty 'Last activity (UTC)'($CurrentWeb.LastItemUserModifiedDate.ToShortDateString())
                $SiteDetails | Add-Member NoteProperty 'Date created'($CurrentWeb.Created.ToShortDateString())
                $SiteDetails | Add-Member NoteProperty 'Created by'($CurrentWeb.Author.LoginName)
                $SiteDetails | Add-Member NoteProperty 'Storage limit (GB)'($CurrentSiteCollection.StorageQuota -f {0:N2})
                $SiteDetails | Add-Member NoteProperty 'Sensitivity'($CurrentSiteCollection.SensitivityLabel)
            
                $SiteDetails | Add-Member NoteProperty 'External sharing'((IIf (-not $CurrentSiteCollection.SharingCapability) $CurrentSiteCollection.SharingCapability $DefaultSharingCapability))
                $SiteDetails | Add-Member NoteProperty 'Sharing Domain Restriction Mode'((IIf (-not $CurrentSiteCollection.SharingDomainRestrictionMode) $CurrentSiteCollection.SharingDomainRestrictionMode $DefaultSharingDomainRestrictionMode))
                $SiteDetails | Add-Member NoteProperty 'Sharing Allowed Domain List'((IIf (-not $CurrentSiteCollection.SharingAllowedDomainList) $CurrentSiteCollection.SharingAllowedDomainList $DefaultSharingAllowedDomainList))
                $SiteDetails | Add-Member NoteProperty 'Sharing Blocked Domain List'((IIf (-not $CurrentSiteCollection.SharingBlockedDomainList) $CurrentSiteCollection.SharingBlockedDomainList $DefaultSharingBlockedDomainList))
                $SiteDetails | Add-Member NoteProperty 'Default Link Permission'((IIf (-not $CurrentSiteCollection.DefaultLinkPermission) $CurrentSiteCollection.DefaultLinkPermission $DefaultLinkPermission))
                $SiteDetails | Add-Member NoteProperty 'Default Sharing Link Type'((IIf (-not $CurrentSiteCollection.DefaultSharingLinkType) $CurrentSiteCollection.DefaultSharingLinkType $DefaultSharingLinkType))
                $SiteDetails | Add-Member NoteProperty 'External User Expiration In Days'((IIf ($CurrentSiteCollection.ExternalUserExpirationInDays -gt 0) $CurrentSiteCollection.ExternalUserExpirationInDays $DefaultRequireAnonymousLinksExpireInDays))

                $SiteDetails | Add-Member NoteProperty 'Override Tenant Anonymous Link Expiration Policy'($CurrentSiteCollection.OverrideTenantAnonymousLinkExpirationPolicy)
                $SiteDetails | Add-Member NoteProperty 'Override Tenant External User Expiration Policy'($CurrentSiteCollection.OverrideTenantExternalUserExpirationPolicy)
                $SiteDetails | Add-Member NoteProperty 'Allow Downloading Non Web Viewable Files'($CurrentSiteCollection.AllowDownloadingNonWebViewableFiles)
                $SiteDetails | Add-Member NoteProperty 'Allow Editing'($CurrentSiteCollection.AllowEditing)
                $SiteDetails | Add-Member NoteProperty 'Allow Self Service Upgrade'($CurrentSiteCollection.AllowSelfServiceUpgrade)
                $SiteDetails | Add-Member NoteProperty 'Anonymous Link Expiration In Days'($CurrentSiteCollection.AnonymousLinkExpirationInDays)
                $SiteDetails | Add-Member NoteProperty 'Block Download Links File Type'($CurrentSiteCollection.BlockDownloadLinksFileType)
                $SiteDetails | Add-Member NoteProperty 'Comments On Site Pages Disabled'($CurrentSiteCollection.CommentsOnSitePagesDisabled)
                $SiteDetails | Add-Member NoteProperty 'Compatibility Level'($CurrentSiteCollection.CompatibilityLevel)
                $SiteDetails | Add-Member NoteProperty 'Conditional Access Policy'($CurrentSiteCollection.ConditionalAccessPolicy)
                $SiteDetails | Add-Member NoteProperty 'Default Link To Existing Access'($CurrentSiteCollection.DefaultLinkToExistingAccess)
                $SiteDetails | Add-Member NoteProperty 'Deny Add And Customize Pages'($CurrentSiteCollection.DenyAddAndCustomizePages)
                $SiteDetails | Add-Member NoteProperty 'Description'($CurrentSiteCollection.Description)
                $SiteDetails | Add-Member NoteProperty 'Disable App Views'($CurrentSiteCollection.DisableAppViews)
                $SiteDetails | Add-Member NoteProperty 'Disable Company Wide Sharing Links'($CurrentSiteCollection.DisableCompanyWideSharingLinks)
                $SiteDetails | Add-Member NoteProperty 'Disable Flows'($CurrentSiteCollection.DisableFlows)
                $SiteDetails | Add-Member NoteProperty 'Disable Sharing For Non Owners Status'($CurrentSiteCollection.DisableSharingForNonOwnersStatus)
                $SiteDetails | Add-Member NoteProperty 'Group Id'($CurrentSiteCollection.GroupId)
                $SiteDetails | Add-Member NoteProperty 'Hub Site Id'($CurrentSiteCollection.HubSiteId)
                $SiteDetails | Add-Member NoteProperty 'Information Segment'($CurrentSiteCollection.InformationSegment)
                $SiteDetails | Add-Member NoteProperty 'Limited Access File Type'($CurrentSiteCollection.LimitedAccessFileType)
                $SiteDetails | Add-Member NoteProperty 'Locale Id'($CurrentSiteCollection.LocaleId)
                $SiteDetails | Add-Member NoteProperty 'Lock Issue'($CurrentSiteCollection.LockIssue)
                $SiteDetails | Add-Member NoteProperty 'Lock State'($CurrentSiteCollection.LockState)
                $SiteDetails | Add-Member NoteProperty 'Owner'($CurrentSiteCollection.Owner)
                $SiteDetails | Add-Member NoteProperty 'Owner Login Name'($CurrentSiteCollection.OwnerLoginName)
                $SiteDetails | Add-Member NoteProperty 'Owner Name'($CurrentSiteCollection.OwnerName)
                $SiteDetails | Add-Member NoteProperty 'Protection Level Name'($CurrentSiteCollection.ProtectionLevelName)
                $SiteDetails | Add-Member NoteProperty 'PWA Enabled'($CurrentSiteCollection.PWAEnabled)
                $SiteDetails | Add-Member NoteProperty 'Related Group Id'($CurrentSiteCollection.RelatedGroupId)
                $SiteDetails | Add-Member NoteProperty 'Resource Quota'($CurrentSiteCollection.ResourceQuota)
                $SiteDetails | Add-Member NoteProperty 'Resource Quota Warning Level'($CurrentSiteCollection.ResourceQuotaWarningLevel)
                $SiteDetails | Add-Member NoteProperty 'Resource Usage Average'($CurrentSiteCollection.ResourceUsageAverage)
                $SiteDetails | Add-Member NoteProperty 'Resource Usage Current'($CurrentSiteCollection.ResourceUsageCurrent)
                $SiteDetails | Add-Member NoteProperty 'Restricted To Geo'($CurrentSiteCollection.RestrictedToGeo)
                $SiteDetails | Add-Member NoteProperty 'Sandboxed Code Activation Capability'($CurrentSiteCollection.SandboxedCodeActivationCapability)
                $SiteDetails | Add-Member NoteProperty 'Show People Picker Suggestions For Guest Users'($CurrentSiteCollection.ShowPeoplePickerSuggestionsForGuestUsers)
                $SiteDetails | Add-Member NoteProperty 'Site Defined Sharing Capability'($CurrentSiteCollection.SiteDefinedSharingCapability)
                $SiteDetails | Add-Member NoteProperty 'Social Bar On Site Pages Disabled'($CurrentSiteCollection.SocialBarOnSitePagesDisabled)
                $SiteDetails | Add-Member NoteProperty 'Status'($CurrentSiteCollection.Status)
                $SiteDetails | Add-Member NoteProperty 'Storage Quota Type'($CurrentSiteCollection.StorageQuotaType)
                $SiteDetails | Add-Member NoteProperty 'Storage Quota Warning Level'($CurrentSiteCollection.StorageQuotaWarningLevel)
                $SiteDetails | Add-Member NoteProperty 'Storage Usage Current'($CurrentSiteCollection.StorageUsageCurrent)
                $SiteDetails | Add-Member NoteProperty 'Webs Count'($CurrentSiteCollection.WebsCount)

                #Export details to CSV File
                $SiteDetails | Export-CSV $ReportFile -Encoding UTF8 -NoTypeInformation -Append
        
                #Loop throuh all Sub Sites
                Process-SitesOnline -SiteURL $SiteCollection.Url -ReportFile $ReportFile
            }
        }
    }
}

function Process-SensitivityLabels {
    $SPAdminURL = $SPRootSiteCollectionURL.Replace(".sharepoint.com", "-admin.sharepoint.com")
    $ConnSite = Connect-PnPOnlineHelper -Url $SPAdminURL

    $CsvPath = "$PSScriptRoot\SensitivityLabels.csv"

    # Get sensitivity labels and export to CSV
    Get-PnPAvailableSensitivityLabel -Connection $ConnSite |
        Select-Object Name, Id |
        Export-Csv -Path $CsvPath -NoTypeInformation
}

#Password Generation Function used for creating Certs with a Password
function GeneratePassword {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [int]$PasswordLength
    )

    # Character sets
    $upper = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    $lower = 'abcdefghijklmnopqrstuvwxyz'
    $digits = '0123456789'
    $special = '!@#$%^&*()-_=+[]{}|;:,.<>?'

    # Combined set for general selection
    $allChars = $upper + $lower + $digits + $special

    $PassComplexCheck = $false

    do {
        $passwordChars = @()

        # Fill password with random characters from the full set
        for ($i = 0; $i -lt $PasswordLength; $i++) {
            $passwordChars += $allChars.ToCharArray() | Get-Random
        }

        # Join into a string
        $newPassword = -join $passwordChars

        # Check complexity: must contain at least one of each category
        if (
            ($newPassword -cmatch "[A-Z]") -and
            ($newPassword -cmatch "[a-z]") -and
            ($newPassword -match "\d") -and
            ($newPassword -match "[^\w]")
        ) {
            $PassComplexCheck = $true
        }
    } while (-not $PassComplexCheck)

    return $newPassword
}

#Register Azure AD App
function RegisterEntraIdApp(){
    
    Read-Config
    
    $Global:SPRootSiteCollectionURL = Read-Host-With-Default "Root Sharepoint site collection (https://tenantname.sharepoint.com)" $SPRootSiteCollectionURL

    #Prompt for the App Name
    if($Global:appRegistrationName -eq $null -or $appRegistrationName -eq ""){
        $appRegistrationName = "Envision IT Open Source M365 Tenant Dashboard"
    }
    $appRegistrationName = Read-Host-With-Default "Enter a name for the App Registration" $appRegistrationName

    # Generate a password
    $UnsecurePassword = GeneratePassword (25)
    $certPassword = ConvertTo-SecureString $UnsecurePassword -AsPlainText -Force

    # Creates App Registration, generates certificate, uploads the cert to the app registration and cert store on local machine, prompts user to grant admin consent for permissions - all in 1 command
    $app = Register-PnPAzureADApp -ApplicationName $appRegistrationName -Tenant $tenantName `
                -CertificatePassword $certPassword `
                -SharePointApplicationPermissions "Sites.FullControl.All" `
                -GraphApplicationPermissions @(
                    "Application.Read.All",
                    "InformationProtectionPolicy.Read.All"
                ) `
                -Store CurrentUser
    
    $ClientId = $app.'AzureAppId/ClientId'
    $Thumbprint = $app.'Certificate Thumbprint'
    Write-Config

    Write-Host "`n=== Azure AD App Registration Summary ===" -ForegroundColor Cyan
    Write-Host "App Name:               $($appRegistrationName)"
    Write-Host "Client Id:              $($ClientId)"
    Write-Host "Certificate Name:       CN=$($appRegistrationName)"
    Write-Host "Certificate Thumbprint: $($Thumbprint)"
    Write-Host "Certificate Password:   $($UnsecurePassword)"

    Write-Host ""
    Write-Host "Note: If you are planning on importing the generated certificate into another machine, you should record the Certificate Password in your password manager. It cannot be retrieved."
}

##########################################
#  Main Scripts
##########################################

Clear-Host

Set-Location -Path $PSScriptRoot

Write-Host "What would you like to do?"
Write-Host "1. Run Inventory"
Write-Host "2. Register Entra ID App"

Do {
    [int]$ChoiceId = Read-Host-With-Default "Enter the ID from the above list" $ChoiceId
}
Until (($ChoiceId -gt 0) -and ($ChoiceId -le 2))

if ($ChoiceId -eq 1) {
    Process-SiteCollectionOnline
    Process-SensitivityLabels
}
elseif ($ChoiceId -eq 2){
    RegisterEntraIdApp
}

Write-Host -ForegroundColor Green "$(Get-TimeStamp) `tAll Done."