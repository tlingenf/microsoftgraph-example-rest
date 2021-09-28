Import-Module MSAL.PS

Write-Host "Reading configuration from config.json" -ForegroundColor Yellow
$config = Get-Content -Path (Join-Path $PSScriptRoot "config.json") -Raw | ConvertFrom-Json

if ($auth -and $auth.ExpiresOn.LocalDateTime -lt (Get-Date)) {
    $auth = Get-MsalToken -ClientId "205589cc-1f7f-4c77-9ad7-5e0fc82c1eb8" -TenantId "5b52477b-6502-4b3c-8c70-3e3ff25efc07" -RedirectUri "https://login.microsoftonline.com/common/oauth2/nativeclient" -Interactive
}

$httpHeader = @{
    "Authorization"="Bearer $($auth.AccessToken)";
    "Content-Type"="application/json";
    "Accept"="application/json";
}

$getSiteResponse = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/sites/$($config.SharePoint.hostname):$($config.SharePoint.sitePath)" -Headers $httpHeader
Write-Host "Found site at: $($getSiteResponse.webUrl) with id: $($getSiteResponse.id)" -ForegroundColor Green

$getLibrariesReponse = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/sites/$($getSiteResponse.id)/drives" -Headers $httpHeader

$docLib = $getLibrariesReponse.value |? { $_.name -eq $config.SharePoint.libName }

if ($docLib) {
    $getDriveItemsResponse = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/sites/$($getSiteResponse.id)/drives/$($docLib.id)/root/children" -Headers $httpHeader

    Write-Host "Files in this library" -ForegroundColor Yellow
    $getDriveItemsResponse.value | ft name, size, lastModifiedDateTime
}