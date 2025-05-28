# Define Variables

#Location of the imported file
$downloadPath = "$env:USERPROFILE\Downloads"
#Location of the resulting exported file
$exportPath = #"$env:USERPROFILE\YOUR DESIRED PATH\Employees.$(Get-Date -Format 'M-dd-yyyy_hhmmss').csv"
#Location of old exported files for deletion of all but the newest one (Janitorial code)
$oldfilePath = #"$env:USERPROFILE\YOUR DESIRED PATH\"

# Janitorial code - removes all but the newest exported file from ServiceNow into Windows

$removeOld = Get-ChildItem -Path $oldfilePath | Sort-Object LastWriteTime -Descending | Select -First 1
Get-ChildItem -Path $oldfilePath | Where-Object { $_.FullName -ne $removeOld.FullName } | Remove-Item -Force

#Identifies and uses the newest exported file
$latestFile = Get-ChildItem -Path $downloadPath | Where-Object { $_.Name -like "*sc_task*" } | Sort-Object LastWriteTime -Descending | Select -First 1

#Error check for newest exported file
if (-not $latestFile) {
    throw "Error: No 'sc_task.csv' file found in the Downloads folder."
}

#Imports the exported file and checks for tickets in the "Open" status
$tasks = Import-Csv -Path $latestFile.FullName -Encoding UTF8
$hasOpenTasks = $tasks | Where-Object { $_.state -eq 'Open' }

#Error if no tickets in the file have the status of "Open"
if (-not $hasOpenTasks) {
    throw "Error: No tasks with state 'Open' found in the imported CSV."
}

#List of names from the file to plug into Active Directory
$names = $hasOpenTasks | Select-Object -ExpandProperty request.requested_for
#Array used later to construct the CSV file headers - Line 83
$idList = @()
#Flag warning for 'name' cells containing null values in the exported file - default is false until true
$warnedAboutNulls = $false

#Array of OUs to search your list of names within Active Directory
$searchOUs = @(
   # 'Insert your fully qualified OU path to search',
   # 'Insert your fully qualified OU path to search'
)

#Array of OUs to exclude during your AD query
$excludeManagerOU = @(
   # 'Insert your fully qualified OU path to exclude from the search',
   # 'Insert your fully qualified OU path to exclude from the search'
)

#Main Loop - Nested for loop to construct the AD query and resulting CSV file
ForEach ($Name in $names) {
    if (-not [string]::IsNullOrWhiteSpace($Name)) {
        ForEach ($ou in $searchOUs) {
            $search = Get-ADUser -LDAPFilter "(ObjectClass=User)(anr=$($Name))" -SearchScope OneLevel -SearchBase $ou -Properties cn, employeeID, title, manager

            if ($search) {
                foreach ($ID in $search) {
                    $managerDisplayName = if ($ID.manager) {
                        (Get-ADUser -Identity $ID.manager -Properties DisplayName, DistinguishedName).DisplayName
                    } else {
                        $null
                    }

                    if ($ID.manager) {
                        $managerDN = (Get-ADUser -Identity $ID.manager -Properties DistinguishedName).DistinguishedName
                        if ($managerDN -like "*$excludeManagerOU*") {
                            continue
                        }
                    }
                    #Generates randomized temporary password for end user
                    $randint = Get-Random -Maximum 99999 -Minimum 10000
                    #Array that formats the file table headers
                    $customObject = [PSCustomObject]@{
                        Name           = $ID.cn
                        EmployeeID     = $ID.employeeID
                        Password       = $randint
                        Title          = $ID.title
                        Manager        = $managerDisplayName
                        #EmployeeType   = if ($ou -eq 'Insert contractors OU') { "Contractor" } else { "Employee" }
                    }

                    $idList += $customObject
                }
            }
        }
        #
        if (-not $idList) {
            if (-not $warnedAboutNulls) {
                Write-Host "Warning: This file contains 'null' name values."
                $warnedAboutNulls = $true
            }
        }
    }
}

#Removes any duplicate names within the file
$uniqueIDList = $idList | Sort-Object EmployeeID -Unique

#Creates the CSV file from the AD query results
$uniqueIDList | Export-Csv -Path $exportPath -NoTypeInformation

#Opens the newly created CSV file
Start-Process -FilePath $exportPath

## NOTE: Make sure to add inclusions at Lines 6, 8, 41-42, 47-48, and 80 ##