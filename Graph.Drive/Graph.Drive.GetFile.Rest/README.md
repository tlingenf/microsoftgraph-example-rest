## About This Sample

This sample downloads a drive file in 5 MB chunks using the Graph REST API. If the file is updated during the download an InvalidOperationException will be thrown with the message "eTag missmatch.".

## To Configure

Register an app in Azure AD and grant at least Files.Read.All permissions.
Update settings in the appsetting.json file.

```
{
  "appId": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
  "appSecret": "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx",
  "tenantId": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
  "downloadsFolder": "C:\\Drive-Get-File",
  "driveItemUri": "https://graph.microsoft.com/v1.0/users/user@domain.com/drive/items/xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
}
```