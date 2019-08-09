# $ProgressPreference = "SilentlyContinue"
# $WarningPreference = "SilentlyContinue"
# $DebugPreference = "SilentlyContinue"
# $ErrorActionPreference = "Continue"
Import-Module $PSScriptRoot\SharePointPnPPowerShell\SharePointPnPPowerShellOnline.psd1
Write-Progress -Activity "Starting process" -PercentComplete 0

$templateHubComponentId = "997fe955-efac-4f36-8383-e51c3c8eed40"

$global:__AppId = ""
$global:__AppSecret = ""

$global:allSitesCount = 0
$global:doneSites = 0
$global:currentSite = ""
$global:currentSiteSubWebs = 0
$global:currentSiteSubWebsDone = 0

Function Start-WebConnection ($url) {
  Connect-PnPOnline -Url $url -AppId $global:__AppId -AppSecret $global:__AppSecret  #-Credentials $global:credentials
  $web = Get-PnPWeb

  return $web
}

Function Start-Connection ($url) {
  Connect-PnPOnline -Url $url -AppId $global:__AppId -AppSecret $global:__AppSecret  #-Credentials $global:credentials
}

Function Update-Progress($webUrl)
{
  $totalCompleted = $global:doneSites;
  $totalJobs = $global:allSitesCount;
  if ($totalJobs -eq 0  -or $totalJobs -eq $null) { $totalJobs = 1 }
  $percent = ($totalCompleted / $totalJobs*100);

  Write-Progress -Activity "Updating $global:currentSite" -status "SubWeb $global:currentSiteSubWebsDone of $global:currentSiteSubWebs ($webUrl)" -percentComplete $percent
}

Function Add-ModernTemplateHub($webUrl)
{
  Update-Progress -webUrl $webUrl

  #Connect-PnPOnline -Url $webUrl -AppId $global:__AppId -AppSecret $global:__AppSecret  #-Credentials $global:credentials
  $web = Start-WebConnection -url $webUrl

  # Check if Template.SPFx exists
  $ca = Get-PnPCustomAction -Web $web | Where-Object {$_.ClientSideComponentId -eq $templateHubComponentId}

  # Add if TemplateHub.SFPx does not exist in site
  if ($ca -eq $null) {
    Write-Host "    [Adding TemplateHub] $webUrl" -ForegroundColor Green
    $ca = $web.UserCustomActions.Add()
  }
  else {
    Write-Host "    [Re-Add TemplateHub] $webUrl" -ForegroundColor Green
    $ca | ForEach-Object {
      try {
        Remove-PnPCustomAction -Web $web -Identity $_.Id -Force
      }
      catch {
        Write-Host "    retry..." -ForegroundColor DarkGreen
        #Connect-PnPOnline -Url $webUrl -AppId $global:__AppId -AppSecret $global:__AppSecret  #-Credentials $global:credentials
        $web = Start-WebConnection -url $webUrl
        Remove-PnPCustomAction -Web $web -Identity $_.Id -Force
      }

    }

    $ca = $web.UserCustomActions.Add()
  }

  # TemplateHub.SPFx properites
  $ca.ClientSideComponentId = "997fe955-efac-4f36-8383-e51c3c8eed40"
  $ca.ClientSideComponentProperties = "{}"
  $ca.Location = "ClientSideExtension.ListViewCommandSet"
  $ca.Name = "templatehub.spfx"
  $ca.Title = "New from TemplateHub"
  $ca.Description = "Create new documents using TemplateHub"

  $ca.RegistrationId = "101"
  $ca.RegistrationType = 1
  $ca.Sequence = 10001
  $ca.Update()
  $web.Context.ExecuteQuery()

  return $web;
}

Function Add-TemplateToSite($spSite, $isSiteCollection = $true)
{
  $spSite |
    ForEach-Object  {
      $url = [uri]::EscapeUriString($_.Url).ToString()
      # add to root web
      $web = Add-ModernTemplateHub($url)
      # add to subwebs
      #Connect-PnPOnline -Url $url -AppId $global:__AppId -AppSecret $global:__AppSecret  #-Credentials $global:credentials
      $subWebs = $null

      try {
        $subWebs = Get-PnPSubWebs -Web $web
      }
      catch {
        #Connect-PnPOnline -Url $url -AppId $global:__AppId -AppSecret $global:__AppSecret  #-Credentials $global:credentials
        $web = Start-WebConnection -url $url
        $subWebs = Get-PnPSubWebs -Web $web
      }

      # work on subWebs. Initialize
      if ($isSiteCollection)
      {
        $global:currentSiteSubWebsDone = 0
        $global:currentSiteSubWebs = $subWebs.Count
      }

      $subWebs | ForEach-Object {
        if ($isSiteCollection)
        {
          $global:currentSiteSubWebsDone = $global:currentSiteSubWebsDone + 1
        }
        Add-TemplateToSite -spSite $_ -isSiteCollection $false

      }
    }
}

# Get templateHub installed to AppCatalog
# Get-PnPAppInstance |
#   Select-Object * |
#   Where-Object -Property ProductId -eq -Value $templateHubPropertyId


# Connect to AppCatalog
#Connect-PnPOnline -Url (Get-PnPTenantAppCatalogUrl) -Credentials inmetademo-admin

<#
.SYNOPSIS
    .
.DESCRIPTION
    .
.PARAMETER TenantAdminUrl
    .
.PARAMETER AppId
    .
.EXAMPLE
    C:\PS>
    Start-TemplateProvisioning -TenantAdminUrl "https://contoso-admin.sharepoint.com" -AppId "..." -AppSecret "..."
.NOTES
    Author: Inmeta Consulting AS
    Date:   April, 2018
#>
Function Start-TemplateProvisioning
{
  Param(
    [Parameter(Mandatory=$true, HelpMessage="Url of Tenant Admin site, eg. https://contoso-admin.sharepoint.com")]
    $TenantAdminUrl,

    [Parameter(Mandatory=$true, HelpMessage="AppId created with AppInv.aspx")]
    [string] $AppId = $global:__AppId,

    [Parameter(Mandatory=$true, HelpMessage="AppSecret created with AppInv.aspx")]
    $AppSecret = $global:__AppSecret,

    [Parameter(Mandatory=$false, HelpMessage="Filter query, eg. [Url -like https://contoso.sharepoint.com/sites/customers]")]
    $FilterQuery = ""
    )

  $global:__AppId = $appId
  $global:__AppSecret = $appSecret

  $global:allSitesCount = 0
  $global:doneSites = 0
  $global:currentSite = ""
  $global:currentSiteSubWebs = 0
  $global:currentSiteSubWebsDone = 0

  # exclude following WebTemplates
  $excludeTemplates = "EHS#1", "SRCHCEN#0", "SPSPERS#9", "SPSPERS#6", "SPSPERS#2", "SPSMSITEHOST#0", "POINTPUBLISHINGTOPIC#0", "POINTPUBLISHINGPERSONAL#0", "OFFILE#1", "EDISC#0", "BICenterSite#0", "APPCATALOG#0"

  #Connect-PnPOnline -Url $TenantAdminUrl -AppId $global:__AppId -AppSecret $global:__AppSecret  #-Credentials $global:credentials
  Start-Connection -url $TenantAdminUrl

  Write-Progress -Activity "Loading Sites" -status "loading from $TenantAdminUrl" -percentComplete 0

  $allSites = Get-PnPTenantSite -IncludeOneDriveSites -Detailed -Filter $FilterQuery
  $filteredSites = $allSites |
    Where-Object -Property Template -NotIn -Value $excludeTemplates |
    Select-Object Url

  Write-Progress -Activity "Loading Sites" -status "..." -percentComplete 100 -Completed

  $global:allSitesCount = $filteredSites.Count

  $filteredSites |
    ForEach-Object {
      $global:currentSite = $_.Url
      $global:doneSites = $global:doneSites + 1

      Add-TemplateToSite -spSite $_
    }
}
