## NOTE: Make sure to edit Lines 38 & 43 ##

# Define Variables

#Location of the imported file
$importPath = "$env:USERPROFILE\Downloads"
#Location of the resulting exported file
$exportPath = "$env:USERPROFILE\Downloads\Employees.$(Get-Date -Format 'M-dd-yyyy_hhmmss').csv"
#Location of old exported files for deletion of all but the newest one (Janitorial code)

if (-not (Test-Path $exportPath)) {
        New-Item -ItemType Directory -Path $exportPath -Force | Out-Null
}

# Janitorial code - removes all but the newest exported file from ServiceNow into Windows

$removeOld = Get-ChildItem -Path $importPath | Sort-Object LastWriteTime -Descending | Select -First 1
Get-ChildItem -Path $importPath | Where-Object { $_.FullName -ne $removeOld.FullName } | Remove-Item -Force

#Identifies and uses the newest exported file
$latestFile = Get-ChildItem -Path $importPath | Where-Object { $_.Name -like "*sc_task*" } | Sort-Object LastWriteTime -Descending | Select -First 1

if (-not $latestFile) {
    throw "Error: No 'sc_task.csv' file found in the Downloads folder."
}

$tasks = Import-Csv -Path $latestFile.FullName -Encoding UTF8
$hasOpenTasks = $tasks | Where-Object { $_.state -eq 'Open' }

if (-not $hasOpenTasks) {
    throw "Error: No tasks with state 'Open' found in the imported CSV."
}

$names = $hasOpenTasks | Select-Object -ExpandProperty request.requested_for
#Array used later to construct the CSV file headers - Line 83
$idList = @()
#Flag warning for 'name' cells containing null values in the exported file - default is false until true
$warnedAboutNulls = $false

$searchOUs = @(
   # 'Insert your fully qualified OU path to search',
   # 'Insert your fully qualified OU path to search'
)

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

Start-Process -FilePath $exportPath