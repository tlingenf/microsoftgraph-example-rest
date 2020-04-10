[string]$TenantId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
[string]$AppId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
[string]$AppSecret = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
[string]$AccountName = "licenseduser@domain.com"
[string]$AccountPwd = "xxxxxxxx"
[int]$numTasks = 5 # number of tasks to create across all plans for the user

###### Sample values used to randomly generate task values #####
$sampleTaskTitles = @("Touchdown!","Golf Classic","Soccer Fever","Alpine Skiing World","NFL Live","Premier League","Olympic Sports")
$percentCompletion = @(0,0,0,0,50,50,100,100,100)
$checklistTitles = @("Lorem ipsum dolor sit amet, consectetur adipiscing elit.","Sed eleifend neque vitae lacus vestibulum, sit amet condimentum lorem varius.","Aliquam ac eros a elit blandit tristique.","Sed non velit et ligula faucibus facilisis quis ut quam.","Etiam sodales lorem ac dictum faucibus.","Vestibulum dignissim mi elementum ligula dapibus, et placerat tortor elementum.","Maecenas sollicitudin nibh at tortor tincidunt rutrum.","Sed dapibus dolor non elit iaculis iaculis.","Vestibulum pretium risus ornare ipsum fermentum, quis lacinia lorem maximus.","Nullam sagittis arcu sit amet pellentesque feugiat.")
$taskPreviewTypes = @("checklist","description")

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

# Create tasks randomly across plans
for ([int]$i = 0; $i -lt $numTasks; $i++) {

    # select a plan at random
    $selectedPlan = Get-Random -InputObject $availablePlansResponse.value

    # get buckets for the plan so we can create a task in a random bucket
    $planBucketsResponse = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/planner/plans/$($selectedPlan.id)/buckets" `
        -Headers @{
            "Content-Type" = "application/json";
            "Authorization" = "Bearer $($authResponse.access_token)";
        }

    # Create a new task with random values
    $createTaskResponse = Invoke-RestMethod -Method Post -Uri "https://graph.microsoft.com/v1.0/planner/tasks" `
        -Headers @{
            "Content-Type" = "application/json";
            "Authorization" = "Bearer $($authResponse.access_token)";
        } `
        -Body (ConvertTo-Json @{            
            "planId" = $selectedPlan.id;
            "bucketId" = (Get-Random -InputObject $planBucketsResponse.value).id;
            "title" = ("{0} {1}" -f (Get-Random -InputObject $sampleTaskTitles), (-join ((65..90) + (97..122) | Get-Random -Count 5 | % {[char]$_})));
            "appliedCategories" = @{
                "category$(Get-Random -Minimum 1 -Maximum 6)" = $true
            }
            "percentComplete" = Get-Random -InputObject $percentCompletion;
            "dueDateTime" = [DateTime]::Now.AddDays((Get-Random -Minimum -10 -Maximum 10)).ToString("yyyy-MM-ddThh:mm:ssZ")
        })

    Write-Host ("{0} created on plan {1}" -f $createTaskResponse.title, $selectedPlan.title)

    # get the task details so we get use the Etag in the task update command
    $getTaskResponse = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/planner/tasks/$($createTaskResponse.id)/details" `
        -Headers @{
            "Content-Type" = "application/json";
            "Authorization" = "Bearer $($authResponse.access_token)";
        }

    # Update additional task details
    $updateTaskResponse = Invoke-RestMethod -Method Patch -Uri "https://graph.microsoft.com/v1.0/planner/tasks/$($createTaskResponse.id)/details" `
        -Headers @{
            "Content-Type" = "application/json";
            "Authorization" = "Bearer $($authResponse.access_token)";
            "If-Match" = $getTaskResponse.'@odata.etag';
        } `
        -Body (ConvertTo-Json @{
            "previewType" = Get-Random -InputObject $taskPreviewTypes;
            "description" = Get-Random -InputObject $checklistTitles;
            "checklist" = @{
                [System.Guid]::NewGuid().ToString() = @{
                    "@odata.type" = "microsoft.graph.plannerChecklistItem";
                    "title" = Get-Random -InputObject $checklistTitles;  
                    "isChecked" = if ($createTaskResponse.percentComplete -eq 100) { $true } else { (($createTaskResponse.percentComplete -gt 0) -and ((Get-Random -Minimum 1 -Maximum 10) -gt 7)) }
                };
                [System.Guid]::NewGuid().ToString() = @{
                    "@odata.type" = "microsoft.graph.plannerChecklistItem";
                    "title" = Get-Random -InputObject $checklistTitles;  
                    "isChecked" = if ($createTaskResponse.percentComplete -eq 100) { $true } else { (($createTaskResponse.percentComplete -gt 0) -and ((Get-Random -Minimum 1 -Maximum 10) -gt 7)) } #(($getTaskResponse.percentComplete -gt 0) -and ((Get-Random -Minimum 1 -Maximum 10) -gt 7))
                };
                [System.Guid]::NewGuid().ToString() = @{
                    "@odata.type" = "microsoft.graph.plannerChecklistItem";
                    "title" = Get-Random -InputObject $checklistTitles;  
                    "isChecked" = if ($createTaskResponse.percentComplete -eq 100) { $true } else { (($createTaskResponse.percentComplete -gt 0) -and ((Get-Random -Minimum 1 -Maximum 10) -gt 7)) }
                };
            }
        })
}