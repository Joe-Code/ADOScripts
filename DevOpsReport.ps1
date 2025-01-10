# Install the ImportExcel module if not already installed
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}
# Import the ImportExcel module
Import-Module -Name ImportExcel

# Set your organization and PAT here
$organization = ""
$personalAccessToken = ""

# Base URLs for Azure DevOps APIs
$baseUrl = "https://dev.azure.com/$organization/_apis/"
# $graphApiUrl = "https://vssps.dev.azure.com/$organization/_apis/"

# Function to create an Authorization Header using the PAT
function Get-AuthorizationHeader {
    $authHeader = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$personalAccessToken"))
    return @{ Authorization = "Basic $authHeader" }
}

# Function to get all projects
function Get-Projects {
    $projects = @()
    $url = "$baseUrl/projects?api-version=6.0&`$top=100"
    $headers = Get-AuthorizationHeader

    do {
        # Write-Host "Request URL: $url"

        try {
            $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get -ResponseHeadersVariable responseHeaders
            $projects += $response.value

            if ($responseHeaders.'x-ms-continuationtoken') {
                $url = "$baseUrl/projects?continuationToken=$($responseHeaders.'x-ms-continuationtoken')&api-version=6.0&`$top=100"
            } else {
                $url = $null
            }
        } catch {
            Write-Host "Error: $_"
            return $null
        }
    } while ($url)

    return $projects
}

# Function to get teams in a project
function Get-Teams {
    param ($projectId)

    $url = $baseUrl + "projects/" + $projectId + "/teams?api-version=7.1"
    $headers = Get-AuthorizationHeader

    # Write-Host "Request URL: $url"

    try {
        $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get
        return $response.value
    } catch {
        Write-Host "Error: $_"
        return $null
    }
}

# Function to get team members
function Get-TeamMembers {
    param ($projectId, $teamId)

    $url = $baseUrl + "projects/" + $projectId + "/teams/" + $teamId + "/members?api-version=7.1"
    $headers = Get-AuthorizationHeader

    # Write-Host "Request URL: $url"

    try {
        $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get
        return $response.value
    } catch {
        Write-Host "Error: $_"
        return $null
    }
}

# Main Script Execution
$projects = Get-Projects
if ($null -ne $projects) {
    $results = @()

    foreach ($project in $projects) {
        $projectId = $project.id
        $projectName = $project.name
        $lastUpdated = if ($project.lastUpdateTime -ne "0001-01-01T00:00:00") { $project.lastUpdateTime } else { "No recent update available" }

        Write-Host "Project: $projectName, Last Updated: $lastUpdated"

        # Get the teams in the project
        $teams = Get-Teams -projectId $projectId
        if ($null -ne $teams) {
            foreach ($team in $teams) {
                $teamName = $team.name
                Write-Host "    - Team: $teamName"

                # Get the team members
                $members = Get-TeamMembers -projectId $projectId -teamId $team.id
                if ($null -ne $members) {
                    foreach ($member in $members) {
                        $displayName = $member.identity.displayName
                        $uniqueName = $member.identity.uniqueName
                        $isTeamAdmin = $member.isTeamAdmin
                        Write-Host "        - Member: $displayName, Unique Name: $uniqueName, Team Admin: $isTeamAdmin"

                        # Add the result to the array
                        $results += [PSCustomObject]@{
                            Project      = $projectName
                            LastUpdated  = $lastUpdated
                            Team         = $teamName
                            DisplayName  = $displayName
                            UniqueName   = $uniqueName
                            TeamAdmin    = $isTeamAdmin
                        }
                    }
                }
                else {
                    Write-Host "    - No members found for team $teamName"
                }
            }
        }
        else {
            Write-Host "    - No teams found for project $projectName"
        }
    }

    # Export results to Excel
    $results | Export-Csv -Path "output.csv" -NoTypeInformation
    
    
    # $results | Export-Excel -Path "output.xlsx" -AutoSize
} else {
    Write-Host "No projects found."
}
