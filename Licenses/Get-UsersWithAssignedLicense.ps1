Import-Module MSAL.PS -ErrorAction Stop

$SKU_NAME = "POWERAPPS_VIRAL"
$CLIENT_ID = "f2694bc6-5dc8-48bc-880b-14eb1034f65d"
$TENANT_ID = "4e2c0a7e-91a7-4c7f-9ba3-5eb81c88cb8f"
$SCOPES = "Directory.ReadWrite.All"
$LICENSE_NAME = "POWERAPPS_VIRAL"

$auth = Get-MsalToken -ClientId $CLIENT_ID -TenantId $TENANT_ID -Scopes $SCOPES -RedirectUri "https://login.microsoftonline.com/common/oauth2/nativeclient" -Interactive
$graphHeader = @{
    "Authorization" = "Bearer $($auth.AccessToken)";
    "Accept" = "application/json";
    "Content-Type" = "application/json";
}

$getLicenseInfo = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/subscribedSkus" -Headers $graphHeader

$licenseInfo = $getLicenseInfo.value |? { $_.skuPartNumber -eq $LICENSE_NAME }

$foundUsers = @()
$allUsersResponse = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/users?`$select=id,userPrincipalName&`$filter=assignedLicenses/any(s:s/skuId eq $($licenseInfo.skuId))" -Headers $graphHeader
if ($allUsersResponse.value) {
    $foundUsers += $allUsersResponse.value
}

while ($allUsersResponse.'@odata.nextLink') {
    $allUsersResponse = Invoke-RestMethod -Method Get -Uri $allUsersResponse.'@odata.nextLink' -Headers $graphHeader
    $foundUsers += $allUsersResponse.value
}

$foundUsers | Select id, userPrincipalName | Export-Csv -Path "./$($LICENSE_NAME)_$((Get-Date).ToString("MM-dd-yyyyThh-mm-ss")).csv" -NoTypeInformation