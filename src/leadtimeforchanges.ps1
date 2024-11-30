#Parameters for the top level  leadtimeforchanges.ps1 PowerShell script
Param(
    [string] $ownerRepo,
    [string] $workflows,
    [string] $branch,
    [Int32] $numberOfDays,
    [string] $commitCountingMethod = "last",
    [string] $patToken = "",
    [string] $actionsToken = "",
    [string] $appId = "",
    [string] $appInstallationId = "",
    [string] $appPrivateKey = "",
    [string] $apiUrl = "https://api.github.com",
    [string] $ignoreList = ""
)

#The main function
function Main ([string] $ownerRepo,
    [string] $workflows,
    [string] $branch,
    [Int32] $numberOfDays,
    [string] $commitCountingMethod,
    [string] $patToken = "",
    [string] $actionsToken = "",
    [string] $appId = "",
    [string] $appInstallationId = "",
    [string] $appPrivateKey = "",
    [string] $apiUrl = "https://api.github.com",
    [string] $ignoreList = ""
)
{

    #==========================================
    #Input processing
    $ownerRepoArray = $ownerRepo -split '/'
    $owner = $ownerRepoArray[0]
    $repo = $ownerRepoArray[1]
    $workflowsArray = $workflows -split ','
    $numberOfDays = $numberOfDays        
    if ($commitCountingMethod -eq "")
    {
        $commitCountingMethod = "last"
    }
    Write-Host "Owner/Repo: $owner/$repo"
    Write-Host "Number of days: $numberOfDays"
    Write-Host "Workflows: $($workflowsArray[0])"
    Write-Host "Branch: $branch"
    Write-Host "Commit counting method '$commitCountingMethod' being used"

    $ignorePatterns = @()
    if ($ignoreList -ne "") {
        $ignorePatterns = $ignoreList -split ',' | ForEach-Object { $_.Trim() }
    }

    #==========================================
    # Get authorization headers
    $authHeader = GetAuthHeader $patToken $actionsToken $appId $appInstallationId $appPrivateKey

    #Get pull requests from the repo 
    #https://developer.GitHub.com/v3/pulls/#list-pull-requests
    $uri = "$apiUrl/repos/$owner/$repo/pulls?state=all&head=$branch&per_page=100&state=closed";
    if (!$authHeader)
    {
        #No authentication
        $prsResponse = Invoke-RestMethod -Uri $uri -ContentType application/json -Method Get -SkipHttpErrorCheck -StatusCodeVariable "HTTPStatus"
    }
    else
    {
        $prsResponse = Invoke-RestMethod -Uri $uri -ContentType application/json -Method Get -Headers @{Authorization=($authHeader["Authorization"])} -SkipHttpErrorCheck -StatusCodeVariable "HTTPStatus" 
    }
    if ($HTTPStatus -eq "404")
    {
        Write-Output "Repo is not found or you do not have access"
        break
    }  

    $prCounter = 0
    $totalPRHours = 0
    $collectedPRs = @()
    Write-Host "`nCollected Pull Requests:"
    Write-Host "======================="
    Foreach ($pr in $prsResponse){

        $mergedAt = $pr.merged_at
        if ($mergedAt -ne $null -and $pr.merged_at -gt (Get-Date).AddDays(-$numberOfDays))
        {
            # Check if branch should be ignored
            $shouldIgnore = $false
            foreach ($pattern in $ignorePatterns) {
                if (Test-WildcardMatch $pr.head.ref $pattern) {
                    $shouldIgnore = $true
                    Write-Host "Ignoring PR #$($pr.number) from branch '$($pr.head.ref)' (matched pattern: $pattern)"
                    break
                }
            }
            
            if (-not $shouldIgnore) {
                $prCounter++
                # Get the PR timeline to find when it was marked ready for review
                $timelineUrl = "$apiUrl/repos/$owner/$repo/issues/$($pr.number)/timeline"
                if (!$authHeader)
                {
                    $timelineResponse = Invoke-RestMethod -Uri $timelineUrl -Headers @{Accept="application/vnd.github.mockingbird-preview+json"} -Method Get -SkipHttpErrorCheck -StatusCodeVariable "HTTPStatus"
                }
                else
                {
                    $timelineResponse = Invoke-RestMethod -Uri $timelineUrl -Headers @{Authorization=($authHeader["Authorization"]); Accept="application/vnd.github.mockingbird-preview+json"} -Method Get -SkipHttpErrorCheck -StatusCodeVariable "HTTPStatus"
                }
                
                # Find when the PR was marked ready for review (if it was ever a draft)
                $readyForReviewDate = $null
                if ($pr.draft) {
                    foreach ($event in $timelineResponse) {
                        if ($event.event -eq "ready_for_review") {
                            $readyForReviewDate = $event.created_at
                            break
                        }
                    }
                }

                $url2 = "$apiUrl/repos/$owner/$repo/pulls/$($pr.number)/commits?per_page=100";
                if (!$authHeader)
                {
                    #No authentication
                    $prCommitsresponse = Invoke-RestMethod -Uri $url2 -ContentType application/json -Method Get -SkipHttpErrorCheck -StatusCodeVariable "HTTPStatus"
                }
                else
                {
                    $prCommitsresponse = Invoke-RestMethod -Uri $url2 -ContentType application/json -Method Get -Headers @{Authorization=($authHeader["Authorization"])} -SkipHttpErrorCheck -StatusCodeVariable "HTTPStatus" 
                }
                if ($prCommitsresponse.Length -ge 1)
                {
                    if ($commitCountingMethod -eq "last")
                    {
                        $startDate = $prCommitsresponse[$prCommitsresponse.Length-1].commit.committer.date
                    }
                    elseif ($commitCountingMethod -eq "first")
                    {
                        $startDate = $prCommitsresponse[0].commit.committer.date
                    }
                    else
                    {
                        Write-Output "Commit counting method '$commitCountingMethod' is unknown. Expecting 'first' or 'last'"
                    }
                }
            
                if ($startDate -ne $null)
                {
                    # Use ready_for_review date if it exists, otherwise use the original start date
                    $effectiveStartDate = if ($readyForReviewDate -and ($readyForReviewDate -gt $startDate)) { $readyForReviewDate } else { $startDate }
                    $totalDuration = New-TimeSpan –Start $effectiveStartDate –End $mergedAt
                    $businessHours = Get-BusinessHours -StartDate $effectiveStartDate -EndDate $mergedAt
                    $totalPRHours += $businessHours  # Use business hours for the total
                    $pr | Add-Member -NotePropertyName business_hours -NotePropertyValue $businessHours
                    $collectedPRs += $pr
                    Write-Host "PR #$($pr.number): '$($pr.title)'"
                    Write-Host "  Branch: $($pr.head.ref)"
                    Write-Host "  Created: $($pr.created_at)$(if ($pr.draft) { " (as draft)" })"
                    if ($readyForReviewDate) {
                        Write-Host "  Ready for Review: $readyForReviewDate"
                    }
                    Write-Host "  Merged: $($mergedAt)"
                    Write-Host "  Total Calendar Duration: $($totalDuration.TotalHours) hours"
                    Write-Host "  Business Hours Duration: $($businessHours) hours"
                    Write-Host "  URL: $($pr.html_url)`n"
                }
            }
        }
    }

    Write-Host "`nSummary:"
    Write-Host "========"
    Write-Host "Total PRs processed: $prCounter"
    Write-Host "Total hours: $totalPRHours"
    Write-Host "Average hours per PR: $($totalPRHours / [math]::Max(1, $prCounter))`n"

    # Calculate and display daily averages
    Get-DailyAverages -PullRequests $collectedPRs -NumberOfDays $numberOfDays

    #==========================================
    #Get workflow definitions from github
    $uri3 = "$apiUrl/repos/$owner/$repo/actions/workflows"
    if (!$authHeader) #No authentication
    {
        $workflowsResponse = Invoke-RestMethod -Uri $uri3 -ContentType application/json -Method Get -SkipHttpErrorCheck -StatusCodeVariable "HTTPStatus"
    }
    else  #there is authentication
    {
        $workflowsResponse = Invoke-RestMethod -Uri $uri3 -ContentType application/json -Method Get -Headers @{Authorization=($authHeader["Authorization"])} -SkipHttpErrorCheck -StatusCodeVariable "HTTPStatus" 
    }
    if ($HTTPStatus -eq "404")
    {
        Write-Output "Repo is not found or you do not have access"
        break
    }  

    #Extract workflow ids from the definitions, using the array of names. Number of Ids should == number of workflow names
    $workflowIds = [System.Collections.ArrayList]@()
    $workflowNames = [System.Collections.ArrayList]@()
    Foreach ($workflow in $workflowsResponse.workflows){

        Foreach ($arrayItem in $workflowsArray){
            if ($workflow.name -eq $arrayItem)
            {
                #This looks odd: but assigning to a (throwaway) variable stops the index of the arraylist being output to the console. Using an arraylist over an array has advantages making this worth it for here
                if (!$workflowIds.Contains($workflow.id))
                {
                    $result = $workflowIds.Add($workflow.id)
                }
                if (!$workflowNames.Contains($workflow.name))
                {
                    $result = $workflowNames.Add($workflow.name)
                }
            }
        }
    }

    #==========================================
    #Filter out workflows that were successful. Measure the number by date/day. Aggegate workflows together
    $workflowList = @()
    
    #For each workflow id, get the last 100 workflows from github
    Foreach ($workflowId in $workflowIds){
        #set workflow counters    
        $workflowCounter = 0
        $totalWorkflowHours = 0
        
        #Get workflow definitions from github
        $uri4 = "$apiUrl/repos/$owner/$repo/actions/workflows/$workflowId/runs?per_page=100&status=completed"
        if (!$authHeader)
        {
            $workflowRunsResponse = Invoke-RestMethod -Uri $uri4 -ContentType application/json -Method Get -SkipHttpErrorCheck -StatusCodeVariable "HTTPStatus"
        }
        else
        {
            $workflowRunsResponse = Invoke-RestMethod -Uri $uri4 -ContentType application/json -Method Get -Headers @{Authorization=($authHeader["Authorization"])} -SkipHttpErrorCheck -StatusCodeVariable "HTTPStatus"      
        }

        Foreach ($run in $workflowRunsResponse.workflow_runs){
            #Count workflows that are completed, on the target branch, and were created within the day range we are looking at
            if ($run.head_branch -eq $branch -and $run.created_at -gt (Get-Date).AddDays(-$numberOfDays))
            {
                #Write-Host "Adding item with status $($run.status), branch $($run.head_branch), created at $($run.created_at), compared to $((Get-Date).AddDays(-$numberOfDays))"
                $workflowCounter++       
                #calculate the workflow duration            
                $workflowDuration = New-TimeSpan –Start $run.created_at –End $run.updated_at
                $totalworkflowHours += $workflowDuration.TotalHours    
            }
        }
        
        #Save the workflow duration working per workflow
        if ($workflowCounter -gt 0)
        {             
            $workflowList += New-Object PSObject -Property @{totalworkflowHours=$totalworkflowHours;workflowCounter=$workflowCounter}                
        }
    }

    #==========================================
    #Prevent divide by zero errors
    if ($prCounter -eq 0)
    {   
        $prCounter = 1
    }
    $totalAverageworkflowHours = 0
    Foreach ($workflowItem in $workflowList){
        if ($workflowItem.workflowCounter -eq 0)
        {
            $workflowItem.workflowCounter = 1
        }
        $totalAverageworkflowHours += $workflowItem.totalworkflowHours / $workflowItem.workflowCounter
    }
    
    #Aggregate the PR and workflow processing times to calculate the average number of hours 
    Write-Host "PR average time duration $($totalPRHours / $prCounter)"
    Write-Host "Workflow average time duration $($totalAverageworkflowHours)"
    $leadTimeForChangesInHours = ($totalPRHours / $prCounter) + ($totalAverageworkflowHours)
    Write-Host "Lead time for changes in hours: $leadTimeForChangesInHours"

    #==========================================
    #Show current rate limit
    $uri5 = "$apiUrl/rate_limit"
    if (!$authHeader)
    {
        $rateLimitResponse = Invoke-RestMethod -Uri $uri5 -ContentType application/json -Method Get -SkipHttpErrorCheck -StatusCodeVariable "HTTPStatus"
    }
    else
    {
        $rateLimitResponse = Invoke-RestMethod -Uri $uri5 -ContentType application/json -Method Get -Headers @{Authorization=($authHeader["Authorization"])} -SkipHttpErrorCheck -StatusCodeVariable "HTTPStatus"
    }    
    Write-Host "Rate limit consumption: $($rateLimitResponse.rate.used) / $($rateLimitResponse.rate.limit)"

    #==========================================
    #output result
    $dailyDeployment = 24
    $weeklyDeployment = 24 * 7
    $monthlyDeployment = 24 * 30
    $everySixMonthsDeployment = 24 * 30 * 6 #Every 6 months

    #Calculate rating, metric and unit  
    if ($leadTimeForChangesInHours -le 0)
    {
        $rating = "None"
        $color = "lightgrey"
        $displayMetric = 0
        $displayUnit = "hours"
    }
    elseif ($leadTimeForChangesInHours -lt 1) 
    {
        $rating = "Elite"
        $color = "brightgreen"
        $displayMetric = [math]::Round($leadTimeForChangesInHours * 60, 2)
        $displayUnit = "minutes"
    }
    elseif ($leadTimeForChangesInHours -le $dailyDeployment) 
    {
        $rating = "Elite"
        $color = "brightgreen"
        $displayMetric = [math]::Round($leadTimeForChangesInHours, 2)
        $displayUnit = "hours"
    }
    elseif ($leadTimeForChangesInHours -gt $dailyDeployment -and $leadTimeForChangesInHours -le $weeklyDeployment)
    {
        $rating = "High"
        $color = "green"
        $displayMetric = [math]::Round($leadTimeForChangesInHours / 24, 2)
        $displayUnit = "days"
    }
    elseif ($leadTimeForChangesInHours -gt $weeklyDeployment -and $leadTimeForChangesInHours -le $monthlyDeployment)
    {
        $rating = "High"
        $color = "green"
        $displayMetric = [math]::Round($leadTimeForChangesInHours / 24, 2)
        $displayUnit = "days"
    }
    elseif ($leadTimeForChangesInHours -gt $monthlyDeployment -and $leadTimeForChangesInHours -le $everySixMonthsDeployment)
    {
        $rating = "Medium"
        $color = "yellow"
        $displayMetric = [math]::Round($leadTimeForChangesInHours / 24 / 30, 2)
        $displayUnit = "months"
    }
    elseif ($leadTimeForChangesInHours -gt $everySixMonthsDeployment)
    {
        $rating = "Low"
        $color = "red"
        $displayMetric = [math]::Round($leadTimeForChangesInHours / 24 / 30, 2)
        $displayUnit = "months"
    }
    if ($leadTimeForChangesInHours -gt 0 -and $numberOfDays -gt 0)
    {
        Write-Host "Lead time for changes average over last $numberOfDays days, is $displayMetric $displayUnit, with a DORA rating of '$rating'"
        return GetFormattedMarkdown -workflowNames $workflowNames -displayMetric $displayMetric -displayUnit $displayUnit -repo $ownerRepo -branch $branch -numberOfDays $numberOfDays -color $color -rating $rating
    }
    else
    {
        Write-Host "No lead time for changes to display for this workflow and time period"
        return GetFormattedMarkdownForNoResult -workflows $workflows -numberOfDays $numberOfDays
    }
}


#Generate the authorization header for the PowerShell call to the GitHub API
#warning: PowerShell has really wacky return semantics - all output is captured, and returned
#reference: https://stackoverflow.com/questions/10286164/function-return-value-in-powershell
function GetAuthHeader ([string] $patToken, [string] $actionsToken, [string] $appId, [string] $appInstallationId, [string] $appPrivateKey) 
{
    #Clean the string - without this the PAT TOKEN doesn't process
    $patToken = $patToken.Trim()
    #Write-Host  $appId
    #Write-Host "pattoken: $patToken"
    #Write-Host "app id is something: $(![string]::IsNullOrEmpty($appId))"
    #Write-Host "patToken is something: $(![string]::IsNullOrEmpty($patToken))"
    if (![string]::IsNullOrEmpty($patToken))
    {
        Write-Host "Authentication detected: PAT TOKEN"
        $base64AuthInfo = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes(":$patToken"))
        $authHeader = @{Authorization=("Basic {0}" -f $base64AuthInfo)}
    }
    elseif (![string]::IsNullOrEmpty($actionsToken))
    {
        Write-Host "Authentication detected: GITHUB TOKEN"  
        $authHeader = @{Authorization=("Bearer {0}" -f $actionsToken)}
    }
    elseif (![string]::IsNullOrEmpty($appId)) # GitHup App auth
    {
        Write-Host "Authentication detected: GITHUB APP TOKEN"  
        $token = Get-JwtToken $appId $appInstallationId $appPrivateKey        
        $authHeader = @{Authorization=("token {0}" -f $token)}
    }    
    else
    {
        Write-Host "No authentication detected" 
        $base64AuthInfo = $null
        $authHeader = $null
    }

    return $authHeader
}

function ConvertTo-Base64UrlString(
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]$in) 
{
    if ($in -is [string]) {
        return [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($in)) -replace '\+','-' -replace '/','_' -replace '='
    }
    elseif ($in -is [byte[]]) {
        return [Convert]::ToBase64String($in) -replace '\+','-' -replace '/','_' -replace '='
    }
    else {
        throw "GitHub App authenication error: ConvertTo-Base64UrlString requires string or byte array input, received $($in.GetType())"
    }
}

function Get-JwtToken([string] $appId, [string] $appInstallationId, [string] $appPrivateKey, [string] $apiUrl )
{
    # Write-Host "appId: $appId"
    $now = (Get-Date).ToUniversalTime()
    $createDate = [Math]::Floor([decimal](Get-Date($now) -UFormat "%s"))
    $expiryDate = [Math]::Floor([decimal](Get-Date($now.AddMinutes(4)) -UFormat "%s"))
    $rawclaims = [Ordered]@{
        iat = [int]$createDate
        exp = [int]$expiryDate
        iss = $appId
    } | ConvertTo-Json
    # Write-Host "expiryDate: $expiryDate"
    # Write-Host "rawclaims: $rawclaims"

    $Header = [Ordered]@{
        alg = "RS256"
        typ = "JWT"
    } | ConvertTo-Json
    # Write-Host "Header: $Header"
    $base64Header = ConvertTo-Base64UrlString $Header
    # Write-Host "base64Header: $base64Header"
    $base64Payload = ConvertTo-Base64UrlString $rawclaims
    # Write-Host "base64Payload: $base64Payload"

    $jwt = $base64Header + '.' + $base64Payload
    $toSign = [System.Text.Encoding]::UTF8.GetBytes($jwt)

    $rsa = [System.Security.Cryptography.RSA]::Create();    
    # https://stackoverflow.com/a/70132607 lead to the right import
    $rsa.ImportRSAPrivateKey([System.Convert]::FromBase64String($appPrivateKey), [ref] $null);

    try { $sig = ConvertTo-Base64UrlString $rsa.SignData($toSign,[Security.Cryptography.HashAlgorithmName]::SHA256,[Security.Cryptography.RSASignaturePadding]::Pkcs1) }
    catch { throw New-Object System.Exception -ArgumentList ("GitHub App authenication error: Signing with SHA256 and Pkcs1 padding failed using private key $($rsa): $_", $_.Exception) }
    $jwt = $jwt + '.' + $sig
    # send headers
    $uri = "$apiUrl/app/installations/$appInstallationId/access_tokens"
    $jwtHeader = @{
        Accept = "application/vnd.github+json"
        Authorization = "Bearer $jwt"
    }
    $tokenResponse = Invoke-RestMethod -Uri $uri -Headers $jwtHeader -Method Post -ErrorAction Stop
    # Write-Host $tokenResponse.token
    return $tokenResponse.token
}

# Format output for deployment frequency in markdown
function GetFormattedMarkdown([array] $workflowNames, [string] $rating, [string] $displayMetric, [string] $displayUnit, [string] $repo, [string] $branch, [string] $numberOfDays, [string] $numberOfUniqueDates, [string] $color)
{
    $encodedString = [uri]::EscapeUriString($displayMetric + " " + $displayUnit)
    #double newline to start the line helps with formatting in GitHub logs
    $markdown = "`n`n![Lead time for changes](https://img.shields.io/badge/frequency-" + $encodedString + "-" + $color + "?logo=github&label=Lead%20time%20for%20changes)`n" +
        "**Definition:** For the primary application or service, how long does it take to go from code committed to code successfully running in production.`n" +
        "**Results:** Lead time for changes is **$displayMetric $displayUnit** with a **$rating** rating, over the last **$numberOfDays days**.`n" + 
        "**Details**:`n" + 
        "- Repository: $repo using $branch branch`n" + 
        "- Workflow(s) used: $($workflowNames -join ", ")`n" +
        "---"
    return $markdown
}

function GetFormattedMarkdownForNoResult([string] $workflows, [string] $numberOfDays)
{
    #double newline to start the line helps with formatting in GitHub logs
    $markdown = "`n`n![Lead time for changes](https://img.shields.io/badge/frequency-none-lightgrey?logo=github&label=Lead%20time%20for%20changes)`n`n" +
        "No data to display for $ownerRepo over the last $numberOfDays days`n`n" + 
        "---"
    return $markdown
}

# Add new helper function to test wildcard matches
function Test-WildcardMatch {
    param (
        [string]$BranchName,
        [string]$Pattern
    )
    
    # Convert the glob pattern to regex
    $regexPattern = $Pattern.Replace(".", "\.")
    $regexPattern = $regexPattern.Replace("**", "###")  # Temp placeholder for **
    $regexPattern = $regexPattern.Replace("*", "[^/]*")
    $regexPattern = $regexPattern.Replace("###", ".*")  # Replace ** with .*
    $regexPattern = "^" + $regexPattern + "$"
    
    return $BranchName -match $regexPattern
}

# Add this new helper function to calculate business hours
function Get-BusinessHours {
    param (
        [DateTime]$StartDate,
        [DateTime]$EndDate
    )
    
    # Convert to local time to make date comparisons easier
    $currentDate = $StartDate.ToLocalTime()
    $endDate = $EndDate.ToLocalTime()
    $totalHours = 0
    
    while ($currentDate -lt $endDate) {
        # Check if current day is weekday (Monday = 1, Sunday = 0)
        if ($currentDate.DayOfWeek -ne 'Saturday' -and $currentDate.DayOfWeek -ne 'Sunday') {
            # For the first day, only count remaining hours in the day
            if ($currentDate.Date -eq $StartDate.Date) {
                $hoursInDay = [Math]::Min(
                    (New-TimeSpan -Start $currentDate -End $currentDate.Date.AddDays(1)).TotalHours,
                    (New-TimeSpan -Start $currentDate -End $endDate).TotalHours
                )
                $totalHours += $hoursInDay
            }
            # For the last day, only count hours until the end time
            elseif ($currentDate.Date -eq $endDate.Date) {
                $hoursInDay = (New-TimeSpan -Start $currentDate.Date -End $endDate).TotalHours
                $totalHours += $hoursInDay
            }
            # For full days in between, add 24 hours
            else {
                $totalHours += 24
            }
        }
        $currentDate = $currentDate.AddDays(1).Date
    }
    
    return $totalHours
}

# Add this helper function to calculate DORA rating
function Get-DoraRating {
    param (
        [double]$Hours
    )
    
    $dailyDeployment = 24
    $weeklyDeployment = 24 * 7
    $monthlyDeployment = 24 * 30
    $everySixMonthsDeployment = 24 * 30 * 6

    if ($Hours -le 0) {
        return @{ Rating = "None"; Color = "lightgrey" }
    }
    elseif ($Hours -lt 1) {
        return @{ Rating = "Elite"; Color = "brightgreen" }
    }
    elseif ($Hours -le $dailyDeployment) {
        return @{ Rating = "Elite"; Color = "brightgreen" }
    }
    elseif ($Hours -gt $dailyDeployment -and $Hours -le $weeklyDeployment) {
        return @{ Rating = "High"; Color = "green" }
    }
    elseif ($Hours -gt $weeklyDeployment -and $Hours -le $monthlyDeployment) {
        return @{ Rating = "High"; Color = "green" }
    }
    elseif ($Hours -gt $monthlyDeployment -and $Hours -le $everySixMonthsDeployment) {
        return @{ Rating = "Medium"; Color = "yellow" }
    }
    else {
        return @{ Rating = "Low"; Color = "red" }
    }
}

# Modify the Get-DailyAverages function
function Get-DailyAverages {
    param (
        [array]$PullRequests,
        [int]$NumberOfDays,
        [int]$WindowDays = 14  # Default to show last 14 days
    )
    
    # Create a hashtable to store PRs by date
    $dailyStats = @{}
    
    # Initialize all dates in the range with empty arrays
    $endDate = Get-Date
    $startDate = $endDate.AddDays(-[Math]::Min($WindowDays, $NumberOfDays))
    
    for ($date = $startDate; $date -le $endDate; $date = $date.AddDays(1)) {
        $dailyStats[$date.Date] = @{
            PRCount = 0
            TotalHours = 0
            PRs = @()
        }
    }
    
    # Group PRs by merge date
    foreach ($pr in $PullRequests) {
        if ($pr.merged_at -ne $null) {
            $mergeDate = [DateTime]$pr.merged_at
            if ($mergeDate -ge $startDate -and $mergeDate -le $endDate) {
                $dailyStats[$mergeDate.Date].PRCount++
                $dailyStats[$mergeDate.Date].TotalHours += $pr.business_hours
                $dailyStats[$mergeDate.Date].PRs += $pr
            }
        }
    }
    
    # Calculate and display the daily averages
    Write-Host "`nDaily DORA Metrics (Last $WindowDays days):"
    Write-Host "================================="
    
    $totalPRs = 0
    $totalHours = 0
    $daysWithPRs = 0
    $ratings = @{
        Elite = 0
        High = 0
        Medium = 0
        Low = 0
        None = 0
    }
    
    for ($date = $startDate; $date -le $endDate; $date = $date.AddDays(1)) {
        $stats = $dailyStats[$date.Date]
        $avgHours = if ($stats.PRCount -gt 0) { $stats.TotalHours / $stats.PRCount } else { 0 }
        $doraRating = Get-DoraRating -Hours $avgHours
        
        $totalPRs += $stats.PRCount
        $totalHours += $stats.TotalHours
        if ($stats.PRCount -gt 0) { 
            $daysWithPRs++
            $ratings[$doraRating.Rating]++
        }
        
        $dateStr = $date.ToString("yyyy-MM-dd")
        $dayName = $date.DayOfWeek.ToString()
        
        # Format the output with DORA rating
        $ratingColor = switch ($doraRating.Rating) {
            "Elite" { "Green" }
            "High" { "DarkGreen" }
            "Medium" { "Yellow" }
            "Low" { "Red" }
            default { "Gray" }
        }
        
        Write-Host ("{0} ({1}): " -f $dateStr, $dayName) -NoNewline
        Write-Host $doraRating.Rating -ForegroundColor $ratingColor -NoNewline
        Write-Host (", {0} PRs, Avg {1:F2} hours" -f $stats.PRCount, [math]::Round($avgHours, 2))
        
        if ($stats.PRCount -gt 0) {
            foreach ($pr in $stats.PRs) {
                Write-Host ("  - #$($pr.number): $($pr.title) ({0:F2} hours)" -f $pr.business_hours)
            }
        }
    }
    
    # Calculate overall statistics
    $overallAvgHours = if ($totalPRs -gt 0) { $totalHours / $totalPRs } else { 0 }
    $dailyAvgPRs = if ($daysWithPRs -gt 0) { $totalPRs / $daysWithPRs } else { 0 }
    $overallRating = Get-DoraRating -Hours $overallAvgHours
    
    Write-Host "`nPeriod Statistics:"
    Write-Host "=================="
    Write-Host "Total PRs: $totalPRs"
    Write-Host "Days with PRs: $daysWithPRs"
    Write-Host ("Average PRs per active day: {0:F2}" -f $dailyAvgPRs)
    Write-Host ("Overall average hours per PR: {0:F2}" -f $overallAvgHours)
    Write-Host "Overall DORA Rating: $($overallRating.Rating)"
    Write-Host "`nDaily DORA Ratings Distribution:"
    Write-Host "Elite: $($ratings.Elite) days"
    Write-Host "High: $($ratings.High) days"
    Write-Host "Medium: $($ratings.Medium) days"
    Write-Host "Low: $($ratings.Low) days"
    Write-Host "No PRs: $($ratings.None) days"
}

main -ownerRepo $ownerRepo -workflows $workflows -branch $branch -numberOfDays $numberOfDays -commitCountingMethod $commitCountingMethod  -patToken $patToken -actionsToken $actionsToken -appId $appId -appInstallationId $appInstallationId -appPrivateKey $appPrivateKey -apiUrl $apiUrl -ignoreList $ignoreList
