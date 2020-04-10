[string]$TenantId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
[string]$AppId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
[string]$AppSecret = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
[string]$AccountName = "licenseduser@domain.com"
[string]$AccountPwd = "xxxxxxxx"

##### MAIN #####

# authenticate using delegated credentials
$authResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
    -Headers @{
        "Content-Type" = "application/x-www-form-urlencoded";
    } `
    -Body @{
        "grant_type" = "password";
        "client_id" = $AppId;
        "client_secret" = $AppSecret;
        "scope" = "Tasks.Read.Shared Group.ReadWrite.All";
        "username" = $AccountName;
        "password" = $AccountPwd;
    }

# Get plans for the user account
$availablePlansResponse = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/me/planner/plans" `
    -Headers @{
        "Content-Type" = "application/json";
        "Authorization" = "Bearer $($authResponse.access_token)";
    }

foreach($plan in $availablePlansResponse.value) {
    # Get all the tasks of the current plan
    $getAllTasksResponse = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/planner/plans/$($plan.id)/tasks" `
        -Headers @{
            "Content-Type" = "application/json";
            "Authorization" = "Bearer $($authResponse.access_token)";
        }

    # Delete each task
    foreach($task in $getAllTasksResponse.value) {
        $deleteTaskResponse = Invoke-RestMethod -Method Delete -Uri "https://graph.microsoft.com/v1.0/planner/tasks/$($task.id)" `
        -Headers @{
            "Content-Type" = "application/json";
            "Authorization" = "Bearer $($authResponse.access_token)";
            "If-Match" = $task.'@odata.etag';
        }
    }
}