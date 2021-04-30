Import-Module MSAL.PS

Write-Host "Reading configuration from config.json" -ForegroundColor Yellow
$config = Get-Content -Path (Join-Path $PSScriptRoot "config.json") -Raw | ConvertFrom-Json

$authToken = Get-MsalToken -ClientId $config.auth.appId -TenantId $config.auth.tenant -RedirectUri $config.auth.redirectUri -Interactive

$httpHeader = @{
    "Authorization"="Bearer $($authToken.AccessToken)";
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