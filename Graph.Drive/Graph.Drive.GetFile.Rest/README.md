```
##########################################################################################################
# Disclaimer
# This sample code, scripts, and other resources are not supported under any Microsoft standard support 
# program or service and are meant for illustrative purposes only.
#
# The sample code, scripts, and resources are provided AS IS without warranty of any kind. Microsoft 
# further disclaims all implied warranties including, without limitation, any implied warranties of 
# merchantability or of fitness for a particular purpose. The entire risk arising out of the use or 
# performance of this material and documentation remains with you. In no event shall Microsoft, its 
# authors, or anyone else involved in the creation, production, or delivery of the sample be liable 
# for any damages whatsoever (including, without limitation, damages for loss of business profits, 
# business interruption, loss of business information, or other pecuniary loss) arising out of the 
# use of or inability to use the samples or documentation, even if Microsoft has been advised of 
# the possibility of such damages.
##########################################################################################################
```
 
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
