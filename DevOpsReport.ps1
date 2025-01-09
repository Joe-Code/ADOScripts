# Set your organization and PAT here
$organization = "jfitzgerald0964"
$personalAccessToken = "DAoyApOjBVz0SicixYZdq5JlkvpXISFKRHltXMoxjhifOHAHR0UAJQQJ99BAACAAAAA3WEulAAASAZDOoDGu"

# Base URLs for Azure DevOps APIs
$baseUrl = "https://dev.azure.com/$organization/_apis/"
$graphApiUrl = "https://vssps.dev.azure.com/$organization/_apis/"

# Function to create an Authorization Header using the PAT
function Get-AuthorizationHeader {
    $authHeader = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$personalAccessToken"))
    return @{ Authorization = "Basic $authHeader" }
}

# Function to get all projects
function Get-Projects {
    $projects = @()
    $url = "$baseUrl/projects?api-version=6.0&`$top=4"
    $headers = Get-AuthorizationHeader

    do {
        Write-Host "Request URL: $url"

        try {
            $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get -ResponseHeadersVariable responseHeaders
            $projects += $response.value

            if ($responseHeaders.'x-ms-continuationtoken') {
                $url = "$baseUrl/projects?continuationToken=$($responseHeaders.'x-ms-continuationtoken')&api-version=6.0&`$top=4"
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

# Function to get project descriptor
function Get-ProjectDescriptor {
    param ($projectId)

    # $url = "$graphApiUrl/graph/descriptors/`$projectId?api-version=6.0-preview.1"
    $url = $graphApiUrl + "graph/descriptors/" + $projectId + "?api-version=6.0-preview.1"
    $headers = Get-AuthorizationHeader

    Write-Host "Request URL: $url"

    try {
        $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get
        return $response.value
    } catch {
        Write-Host "Error: $_"
        return $null
    }
}

# Function to get project memberships using descriptor
function Get-ProjectMemberships {
    param ($projectDescriptor)

    $url = $graphApiUrl + "graph/memberships/" + $projectDescriptor + "?api-version=6.0-preview.1"
    $headers = Get-AuthorizationHeader

    Write-Host "Request URL: $url"

    try {
        $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get
        return $response.value
    } catch {
        Write-Host "Error: $_"
        return $null
    }
}

# Function to get user details by user descriptor
function Get-UserDetails {
    param ($userDescriptor)

    $url = "$graphApiUrl/graph/users/$userDescriptor?api-version=7.1-preview.1"
    $headers = Get-AuthorizationHeader

    Write-Host "Request URL: $url"
    Write-Host "Headers: $headers"

    try {
        $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get
        return $response
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

        # Get the project descriptor
        $projectDescriptor = Get-ProjectDescriptor -projectId $projectId
        Write-Host "Project Descriptor: $projectDescriptor"

        if ($null -ne $projectDescriptor) {
            # Get the project memberships using the descriptor
            $memberships = Get-ProjectMemberships -projectDescriptor $projectDescriptor

            foreach ($membership in $memberships) {
                $userDescriptor = $membership.descriptor

                # Get user details using the user descriptor
                $user = Get-UserDetails -userDescriptor $userDescriptor

                if ($user) {
                    $displayName = $user.displayName
                    $email = $user.principalName
                    Write-Host "- User: $displayName, Email: $email"

                    # Add the result to the array
                    $results += [PSCustomObject]@{
                        Project      = $projectName
                        LastUpdated  = $lastUpdated
                        User         = $displayName
                        Email        = $email
                    }
                } else {
                    Write-Host "- No user information found for descriptor $userDescriptor"
                }
            }
        } else {
            Write-Host "- No project descriptor found for project $projectName"
        }
    }

    # Export results to Excel
    # $results | Export-Excel -Path "output.xlsx" -AutoSize
} else {
    Write-Host "No projects found."
}
