[string]$TenantId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
[string]$AppId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
[string]$AppSecret = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
[string]$AccountName = "licenseduser@domain.com"
[string]$AccountPwd = "xxxxxxxx"

$bucketNames = @("Programming","Events","Production","Schedule")  # Values for bucket names to create
$categoryLabels = @{
    "category1" = "Scheduled To Air";
    "category2" = "Ready For Scheduling";
    "category3" = "Category 3";
    "category4" = "Category 4";
    "category5" = "Category 5";
    "category6" = "Category 6";
} # Hashtable of all category label values to set

# Function to check if buckets exist, if not create them
function CreateBuckets([string]$PlanId) {
    Write-Host ("Creating buckets for plan {0}" -f $PlanId)

    # Get the current collection of buckets in the specified plan
    $planBucketsResponse = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/planner/plans/$PlanId/buckets" `
        -Headers @{
            "Content-Type" = "application/json";
            "Authorization" = "Bearer $($authResponse.access_token)";
        }

    # For each bucket that is to be created, check if it exists, if not create it
    foreach ($bucketName in $bucketNames) {
        if (($planBucketsResponse.value |? {$_.name -eq $bucketName }) -eq $null) {
            Write-Host ("Plan {0} does not have bucket {1}. Creating bucket {1}" -f $PlanId, $bucketName)
            $newBucketResponse = Invoke-RestMethod -Method Post -Uri "https://graph.microsoft.com/v1.0/planner/buckets" `
            -Headers @{
                "Content-Type" = "application/json";
                "Authorization" = "Bearer $($authResponse.access_token)";
            } `
            -Body (ConvertTo-Json @{
                "name" = $bucketName;
                "planId" = $PlanId;
            })             
        }
    }

    # Get the buckets for the plan again after any modification have been made
    $planBucketsResponse = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/planner/plans/$PlanId/buckets" `
        -Headers @{
            "Content-Type" = "application/json";
            "Authorization" = "Bearer $($authResponse.access_token)";
        }

    return $planBucketsResponse.value
}

# Sets the category labels for the plan
function UpdateCategoryLabels([string]$PlanId, $LabelCollection) {
    Write-Host ("Updating Categories for plan {0}" -f $PlanId)

    # get the plan details, required so we can obtain the current ETag value
    $getPlanResponse = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/planner/plans/$PlanId/details" `
        -Headers @{
            "Content-Type" = "application/json";
            "Authorization" = "Bearer $($authResponse.access_token)";
        }

    # Set the label collection using the Etag in the If-Match header for concurrency validation
    $UpdatePlanResponse = Invoke-RestMethod -Method Patch -Uri "https://graph.microsoft.com/v1.0/planner/plans/$PlanId/details" `
        -Headers @{
            "Content-Type" = "application/json";
            "Authorization" = "Bearer $($authResponse.access_token)";
            "If-Match" =  $getPlanResponse.'@odata.etag';
        } `
        -Body (ConvertTo-Json @{
            "categoryDescriptions" = $LabelCollection;
        })
}

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

foreach ($plan in $availablePlansResponse.value) {
    UpdateCategoryLabels -PlanId $plan.id -LabelCollection $categoryLabels
    $buckets = CreateBuckets -PlanId $plan.id
}