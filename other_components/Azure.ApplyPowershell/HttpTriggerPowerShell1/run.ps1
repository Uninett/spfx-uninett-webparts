Import-Module 'D:\Home\site\wwwroot\HttpTriggerPowerShell1\Modules\Microsoft.Online.SharePoint.PowerShell'

# POST method: $req
$requestBody = Get-Content $req -Raw | ConvertFrom-Json
 
try {
	#URL of the target site collection
	$targetSiteCollection = $requestBody.siteUrl

	# Site type of site, which determines which theme to apply
	$siteType = $requestBody.siteType

	$themeName = $siteType

	#Autenticate on SharePoint Online site collection, credentials might be requested
	$adminUPN = $env:TenantAdminUPN
	$orgName = $env:OrgName
	$password = $env:Password

	$userCredential = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $adminUPN, $(convertto-securestring $password -asplaintext -force)

	Connect-SPOService -Url https://$orgName-admin.sharepoint.com -Credential $userCredential

	Set-SPOWebTheme -Theme $themeName -Web $targetSiteCollection

	Out-File -Encoding Ascii -FilePath $res -inputObject "Theme $siteType applied to $targetSiteCollection"

} catch {
	Write-Host -ForegroundColor Red "Exception occurred!" 
    Write-Host -ForegroundColor Red "Exception Type: $($_.Exception.GetType().FullName)"
    Write-Host -ForegroundColor Red "Exception Message: $($_.Exception.Message)"

	Out-File -Encoding Ascii -FilePath $res -inputObject "Exception Message: $($_.Exception.Message)"
}
