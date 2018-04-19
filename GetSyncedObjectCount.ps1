#Copyright © 2012 Microsoft Corporation.  All rights reserved.
#
# THIS CODE AND ANY ASSOCIATED INFORMATION ARE PROVIDED “AS IS” WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR 
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR
# PURPOSE. THE ENTIRE RISK OF USE, INABILITY TO USE, OR RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.


# Global Functions
function Check-MsolServiceConnection
{

# Verifies if we are connected to the service
# Check if the module is loaded
$bModuleLoaded=$false
$bConnectedToService=$false

Get-Module|%{if($_.Name -eq "MsOnline"){$bModuleLoaded = $true}}
if($bModuleLoaded -eq $true)
	{
		#Module is loaded proceed checking if we are logged in.
		Write-Host -ForeGroundColor yellow "MSOnline PowerShell module is loaded."
	}
else
	{
		# Module is not loaded. Load the module and connect.
		Import-Module MSOnline
		Connect-MsolService
	}

try
	{
		$tenantInfo=Get-MsolCompanyInformation -ErrorAction Stop
		Write-Host -ForeGroundColor Yellow "Connected to" $tenantInfo.DisplayName
	}
catch
	{
		Write-Host -ForeGroundColor Red "Could not find a connected session. Reconnecting."
		Connect-MsolService
		Get-MsolCompanyInformation |fl
	}
}

# Start the Timer
$timerStartTime=Get-Date

Check-MsolServiceConnection

Write-Host -ForegroundColor Yellow "Retrieving Objects from the Online Directory. Depending on the number of objects this may take several minutes to complete."
# Get all objects
$users=Get-MsolUser -all -Synchronized
$groups=Get-MsolGroup -all
$contacts=Get-MsolContact -All

"Filtering Objects. Please wait."
# Filter for lastDirSync
$users | %{if($_.LastDirSyncTime -ne ""){$ucount=$ucount+1}}
$groups | %{if($_.LastDirSyncTime -ne ""){$gcount=$gcount+1}}
$contacts | %{if($_.LastDirSyncTime -ne ""){$ccount=$ccount+1}}

# Caluclate Total Synced Object Count
$totalSyncedObjects=$uCount+$gCount+$cCount

# Output Synced Object Count
"Found $uCount synchronized Users."
"Found $gCount synchronized Groups."
"Found $cCount synchronized Contacts."
"======================================="
"TOTAL SYNCHRONIZED OBJECTS: $totalSyncedObjects"

# End the Timer
$timerEndTime=Get-Date
$executionTime=($timerEndTime - $timerStartTime).TotalSeconds

# Write Execution Time
Write-Host -ForeGroundColor Cyan "Object count query completed in $executionTime seconds."

