
$tenant = Read-Host -Prompt "Enter tenant name"
$url = "https://" + $tenant + "-admin.sharepoint.com"

$cred = Get-Credential -Message "Please enter credentials for " + $url + ":"

Write-Host "Connecting to "  $url
Connect-SPOService -Url $url -Credential $cred


$DepartmentTheme = @{
    "themePrimary" = "#003057";
    "themeLighterAlt" = "#adffd9";
    "themeLighter" = "#47d7ac";
    "themeLight" = "#45b4e2";
    "themeTertiary" = "#007dba";
    "themeSecondary" = "#00558c";
    "themeDarkAlt" = "#007dba";
    "themeDark" = "#00558c";
    "themeDarker" = "#012639";
    "neutralLighterAlt" = "#f8f8f8";
    "neutralLighter" = "#f4f4f4";
    "neutralLight" = "#eaeaea";
    "neutralQuaternaryAlt" = "#dadada";
    "neutralQuaternary" = "#d0d0d0";
    "neutralTertiaryAlt" = "#c8c8c8";
    "neutralTertiary" = "#a6a6a6";
    "neutralSecondary" = "#007dba";
    "neutralPrimaryAlt" = "#00558c";
    "neutralPrimary" = "#012639";
    "neutralDark" = "#003057";
    "black" = "#023469";
    "white" = "#ffffff";
    "primaryBackground" = "#ffffff";
    "primaryText" = "#012639";
    "bodyBackground" = "#ffffff";
    "bodyText" = "#012639";
    "disabledBackground" = "#f4f4f4";
    "disabledText" = "#c8c8c8";
    "accent" = "#47d7ac";
}

$ActivityTheme = @{
    "themePrimary" = "#007dba";
    "themeLighterAlt" = "#e8f7ff";
    "themeLighter" = "#adffd9";
    "themeLight" = "#47d7ac";
    "themeTertiary" = "#45b4e2";
    "themeSecondary" = "#00558c";
    "themeDarkAlt" = "#00558c";
    "themeDark" = "#003057";
    "themeDarker" = "#012639";
    "neutralLighterAlt" = "#f8f8f8";
    "neutralLighter" = "#f4f4f4";
    "neutralLight" = "#eaeaea";
    "neutralQuaternaryAlt" = "#dadada";
    "neutralQuaternary" = "#d0d0d0";
    "neutralTertiaryAlt" = "#c8c8c8";
    "neutralTertiary" = "#a6a6a6";
    "neutralSecondary" = "#007dba";
    "neutralPrimaryAlt" = "#00558c";
    "neutralPrimary" = "#012639";
    "neutralDark" = "#003057";
    "black" = "#024769";
    "white" = "#ffffff";
    "primaryBackground" = "#ffffff";
    "primaryText" = "#012639";
    "bodyBackground" = "#ffffff";
    "bodyText" = "#012639";
    "disabledBackground" = "#f4f4f4";
    "disabledText" = "#c8c8c8";
    "accent" = "#adffd9";
    }

$projectTheme = @{
    "themePrimary" = "#47d7ac";
    "themeLighterAlt" = "#f2fcf9";
    "themeLighter" = "#def8f0";
    "themeLight" = "#afedda";
    "themeTertiary" = "#75e0c0";
    "themeSecondary" = "#4fd8af";
    "themeDarkAlt" = "#2fd1a1";
    "themeDark" = "#219572";
    "themeDarker" = "#1c8062";
    "neutralLighterAlt" = "#f8f8f8";
    "neutralLighter" = "#f4f4f4";
    "neutralLight" = "#eaeaea";
    "neutralQuaternaryAlt" = "#dadada";
    "neutralQuaternary" = "#d0d0d0";
    "neutralTertiaryAlt" = "#c8c8c8";
    "neutralTertiary" = "#a6a6a6";
    "neutralSecondary" = "#007dba";
    "neutralPrimaryAlt" = "#00558c";
    "neutralPrimary" = "#012639";
    "neutralDark" = "#003057";
    "black" = "#024769";
    "white" = "#ffffff";
    "primaryBackground" = "#ffffff";
    "primaryText" = "#012639";
    "bodyBackground" = "#ffffff";
    "bodyText" = "#012639";
    "disabledBackground" = "#f4f4f4";
    "disabledText" = "#c8c8c8";
    "accent" = "#007dba";
    }

Write-Host "Adding globally available themes"    
Add-SPOTheme -Identity "Avdeling" -Palette $DepartmentTheme -IsInverted $false -Overwrite
Add-SPOTheme -Identity "Aktivitet" -Palette $ActivityTheme -IsInverted $false -Overwrite
Add-SPOTheme -Identity "Prosjekt" -Palette $ProjectTheme -IsInverted $false -Overwrite

Write-Host "Disabling SharePoint default themes"    
Set-SPOHideDefaultThemes $true

Write-Host "Done!"