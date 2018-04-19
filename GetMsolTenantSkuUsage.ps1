#------------------------------------------------------------------------------
#
# Copyright © 2012 Microsoft Corporation.  All rights reserved.
#
# THIS CODE AND ANY ASSOCIATED INFORMATION ARE PROVIDED “AS IS” WITHOUT
# WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
# LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS
# FOR A PARTICULAR PURPOSE. THE ENTIRE RISK OF USE, INABILITY TO USE, OR 
# RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
#
#------------------------------------------------------------------------------
#
# PowerShell Source Code
#
# NAME:
#    GetMsolTenantSkuUsage.ps1
#
#------------------------------------------------------------------------------

<#
	.SYNOPSIS
		Generates a CSV usage report of licenses owned, licenses consumed, and available licenses.

	.DESCRIPTION
		This script will establish a connection with the Office 365 provision web service API and
		collect information about number of licenses owned, licenses consumed, and available 
		licenses.  The results will be displayed to the PowerShell console and saved to a CSV file.
		
		If a credential is specified, it will be used to establish a connection with the provisioning
		web service API.
		
		If a credential is not specified, an attempt is made to identify an existing connection to
		the provisioning web service API.  If an existing connection is identified, the existing
		connection is used.  If an existing connection is not identified, the user is prompted for
		credentials so that a new connection can be established.

	.PARAMETER Credential
		Specifies the credential to use when connecting to the Office 365 PowerShell web service.

	.PARAMETER OutputFile
		Specifies the name of the output file.  The arguement can be the full path including the file
		name, or only the path to the folder in which to save the file (uses default name).
		
		Default filename is in the format of "YYYYMMDDhhmmss_MsolTenantSkuUsage.csv"

	.EXAMPLE
		PS> .\GetMsolTenantSkuUsage.ps1

	.EXAMPLE
		PS> .\GetMsolTenantSkuUsage.ps1 -Credential (Get-Credential)

	.EXAMPLE
		PS> .\GetMsolTenantSkuUsage.ps1 -OutputFile "C:\Folder\Sub Folder"

	.EXAMPLE
		PS> .\GetMsolTenantSkuUsage.ps1 -OutputFile "C:\Folder\Sub Folder\File Name.csv"

	.EXAMPLE
		PS> .\GetMsolTenantSkuUsage.ps1 -Credential (Get-Credential) -OutputFile "C:\Folder\Sub Folder"

	.EXAMPLE
		PS> .\GetMsolTenantSkuUsage.ps1 -Credential (Get-Credential) -OutputFile "C:\Folder\Sub Folder\File Name.csv"

	.INPUTS
		System.Management.Automation.PsCredential
		System.String

	.OUTPUTS
		A CSV file.

	.NOTES

#>

[CmdletBinding()]
param
(
	[Parameter(Mandatory = $False)]
	[System.Management.Automation.PsCredential]$Credential,
	
	[Parameter(Mandatory = $False)]
	[ValidateNotNullOrEmpty()]
	[String]$OutputFile = "$((Get-Date -uformat %Y%m%d%H%M%S).ToString())_MsolTenantSkuUsage.csv"
)


Function WriteConsoleMessage
{
	<#
		.SYNOPSIS
			Writes the specified message of the specified message type to
			the PowerShell console.

		.DESCRIPTION
			Writes the specified message of the specified message type to
			the PowerShell console.

		.PARAMETER Message
			Specifies the actual message to be written to the console.

		.PARAMETER MessageType
			Specifies the type of message to be written of either "error", "warning",
			"verbose", or "information".  The message type simply changes the 
			background and foreground colors so that the type of message is known
			at a glance.

		.EXAMPLE
			PS> WriteConsoleMessage -Message "This is an error message" -MessageType "Error"

		.EXAMPLE
			PS> WriteConsoleMessage -Message "This is a warning message" -MessageType "Warning"

		.EXAMPLE
			PS> WriteConsoleMessage -Message "This is a verbose message" -MessageType "Verbose"

		.EXAMPLE
			PS> WriteConsoleMessage -Message "This is an information message" -MessageType "Information"

		.INPUTS
			System.String

		.OUTPUTS
			A message is written to the PowerShell console.

		.NOTES

	#>

	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $True, Position = 0)]
		[ValidateNotNullOrEmpty()]
		[string]$Message,
		
		[Parameter(Mandatory = $True, Position = 1)]
		[ValidateSet("Error", "Warning", "Verbose", "Information")]
		[string]$MessageType
	)
	
	Switch ($MessageType)
	{
		"Error"
		{
			$Message = "ERROR: SCRIPT: {0}" -f $Message
			Write-Host $Message -ForegroundColor Black -BackgroundColor Red
		}
		"Warning"
		{
			$Message = "WARNING: SCRIPT: {0}" -f $Message
			Write-Host $Message -ForegroundColor Black -BackgroundColor Yellow
		}
		"Verbose"
		{
			$Message = "VERBOSE: SCRIPT: {0}" -f $Message
			If ($VerbosePreference -eq "Continue") {Write-Host $Message -ForegroundColor Gray -BackgroundColor Black}
		}
		"Information"
		{
			$Message = "INFORMATION: SCRIPT: {0}" -f $Message
			Write-Host $Message -ForegroundColor Cyan -BackgroundColor Black
		}
	}
}


Function TestFolderExists
{
	<#
		.SYNOPSIS
			Verifies that the specified folder/path exists.

		.DESCRIPTION
			Verifies that the specified folder/path exists.

		.PARAMETER Folder
			Specifies the absolute or relative path to the file.

		.EXAMPLE
			PS> TestFolderExists -Folder "C:\Folder\Sub Folder\File name.csv"

		.EXAMPLE
			PS> TestFolderExists -Folder "File name.csv"

		.EXAMPLE
			PS> TestFolderExists -Folder "C:\Folder\Sub Folder"

		.EXAMPLE
			PS> TestFolderExists -Folder ".\Folder\Sub Folder"

		.INPUTS
			System.String

		.OUTPUTS
			System.Boolean

		.NOTES

	#>

	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $True)]
		[ValidateNotNullOrEmpty()]
		[string]$Folder
	)

	If ([System.IO.Path]::HasExtension($Folder)) {$PathToFile = ([System.IO.Directory]::GetParent($Folder)).FullName}
	Else {$PathToFile = [System.IO.Path]::GetFullPath($Folder)}
	If ([System.IO.Directory]::Exists($PathToFile)) {Return $True}
	Return $False
}


Function GetElapsedTime
{
	<#
		.SYNOPSIS
			Calculates a time interval between two DateTime objects.

		.DESCRIPTION
			Calculates a time interval between two DateTime objects.

		.PARAMETER Start
			Specifies the start time.

		.PARAMETER End
			Specifies the end time.

		.EXAMPLE
			PS> GetElapsedTime -Start "1/1/2011 12:00:00 AM" -End "1/2/2011 2:00:00 PM"

		.EXAMPLE
			PS> GetElapsedTime -Start ([datetime]"1/1/2011 12:00:00 AM") -End ([datetime]"1/2/2011 2:00:00 PM")

		.INPUTS
			System.String

		.OUTPUTS
			System.Management.Automation.PSObject

		.NOTES

	#>

	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $True, Position = 0)]
		[ValidateNotNullOrEmpty()]
		[DateTime]$Start,
		
		[Parameter(Mandatory = $True, Position = 1)]
		[ValidateNotNullOrEmpty()]
		[DateTime]$End
	)
	
	$TotalSeconds = ($End).Subtract($Start).TotalSeconds
	$objElapsedTime = New-Object PSObject
	
	# less than 1 minute
	If ($TotalSeconds -lt 60)
	{
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Days -Value 0
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Hours -Value 0
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Minutes -Value 0
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Seconds -Value $($TotalSeconds)
	}

	# more than 1 minute, less than 1 hour
	If (($TotalSeconds -ge 60) -and ($TotalSeconds -lt 3600))
	{
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Days -Value 0
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Hours -Value 0
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Minutes -Value $([Math]::Truncate($TotalSeconds / 60))
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Seconds -Value $([Math]::Truncate($TotalSeconds % 60))
	}

	# more than 1 hour, less than 1 day
	If (($TotalSeconds -ge 3600) -and ($TotalSeconds -lt 86400))
	{
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Days -Value 0
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Hours -Value $([Math]::Truncate($TotalSeconds / 3600))
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Minutes -Value $([Math]::Truncate(($TotalSeconds % 3600) / 60))
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Seconds -Value $([Math]::Truncate($TotalSeconds % 60))
	}

	# more than 1 day, less than 1 year
	If (($TotalSeconds -ge 86400) -and ($TotalSeconds -lt 31536000))
	{
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Days -Value $([Math]::Truncate($TotalSeconds / 86400))
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Hours -Value $([Math]::Truncate(($TotalSeconds % 86400) / 3600))
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Minutes -Value $([Math]::Truncate((($TotalSeconds - 86400) % 3600) / 60))
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Seconds -Value $([Math]::Truncate($TotalSeconds % 60))
	}
	
	Return $objElapsedTime
}


Function ConnectProvisioningWebServiceAPI
{
	<#
		.SYNOPSIS
			Connects to the Office 365 provisioning web service API.

		.DESCRIPTION
			Connects to the Office 365 provisioning web service API.
			
			If a credential is specified, it will be used to establish a connection with the provisioning
			web service API.
			
			If a credential is not specified, an attempt is made to identify an existing connection to
			the provisioning web service API.  If an existing connection is identified, the existing
			connection is used.  If an existing connection is not identified, the user is prompted for
			credentials so that a new connection can be established.

		.PARAMETER Credential
			Specifies the credential to use when connecting to the provisioning web service API
			using Connect-MsolService.

		.EXAMPLE
			PS> ConnectProvisioningWebServiceAPI

		.EXAMPLE
			PS> ConnectProvisioningWebServiceAPI -Credential
			
		.INPUTS
			[System.Management.Automation.PsCredential]

		.OUTPUTS

		.NOTES

	#>
	
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $False)]
		[System.Management.Automation.PsCredential]$Credential
	)
	
	# if a credential was supplied, assume a new connection is intended and create a new
	# connection using specified credential
	If ($Credential)
	{
		If ((!$Credential) -or (!$Credential.Username) -or ($Credential.Password.Length -eq 0))
		{
			WriteConsoleMessage -Message ("Invalid credential.  Please verify the credential and try again.") -MessageType "Error"
			Exit
		}
		
		# connect to provisioning web service api
		WriteConsoleMessage -Message "Connecting to the Office 365 provisioning web service API.  Please wait..." -MessageType "Information"
		Connect-MsolService -Credential $Credential
		If($? -eq $False){WriteConsoleMessage -Message "Error while connecting to the Office 365 provisioning web service API.  Quiting..." -MessageType "Error";Exit}
	}
	Else
	{
		WriteConsoleMessage -Message "Attempting to identify an open connection to the Office 365 provisioning web service API.  Please wait..." -MessageType "Information"
		$getMsolCompanyInformationResults = Get-MsolCompanyInformation -ErrorAction SilentlyContinue
		If (!$getMsolCompanyInformationResults)
		{
			WriteConsoleMessage -Message "Could not identify an open connection to the Office 365 provisioning web service API." -MessageType "Information"
			If (!$Credential)
			{
				$Credential = $Host.UI.PromptForCredential("Enter Credential",
					"Enter the username and password of an Office 365 administrator account.",
					"",
					"userCreds")
			}
			If ((!$Credential) -or (!$Credential.Username) -or ($Credential.Password.Length -eq 0))
			{
				WriteConsoleMessage -Message ("Invalid credential.  Please verify the credential and try again.") -MessageType "Error"
				Exit
			}
			
			# connect to provisioning web service api
			WriteConsoleMessage -Message "Connecting to the Office 365 provisioning web service API.  Please wait..." -MessageType "Information"
			Connect-MsolService -Credential $Credential
			If($? -eq $False){WriteConsoleMessage -Message "Error while connecting to the Office 365 provisioning web service API.  Quiting..." -MessageType "Error";Exit}
			$getMsolCompanyInformationResults = Get-MsolCompanyInformation -ErrorAction SilentlyContinue
			WriteConsoleMessage -Message ("Connected to Office 365 tenant named: `"{0}`"." -f $getMsolCompanyInformationResults.DisplayName) -MessageType "Information"
		}
		Else
		{
			WriteConsoleMessage -Message ("Connected to Office 365 tenant named: `"{0}`"." -f $getMsolCompanyInformationResults.DisplayName) -MessageType "Warning"
		}
	}
	If (!$Script:Credential) {$Script:Credential = $Credential}
}


# -----------------------------------------------------------------------------
#
# Main Script Execution
#
# -----------------------------------------------------------------------------

$Error.Clear()
$ScriptStartTime = Get-Date

# verify that the MSOnline module is installed and import into current powershell session
If (!([System.IO.File]::Exists(("{0}\modules\msonline\Microsoft.Online.Administration.Automation.PSModule.dll" -f $pshome))))
{
	WriteConsoleMessage -Message ("Please download and install the Microsoft Online Services Module.") -MessageType "Error"
	Exit
}
$getModuleResults = Get-Module
If (!$getModuleResults) {Import-Module MSOnline -ErrorAction SilentlyContinue}
Else {$getModuleResults | ForEach-Object {If (!($_.Name -eq "MSOnline")){Import-Module MSOnline -ErrorAction SilentlyContinue}}}

# verify output directory exists for results file
WriteConsoleMessage -Message ("Verifying folder:  {0}" -f $OutputFile) -MessageType "Verbose"
If (!(TestFolderExists $OutputFile))
{
	WriteConsoleMessage -Message ("Directory not found:  {0}" -f $OutputFile) -MessageType "Error"
	Exit
}

# if a filename was not specified as part of $OutputFile, auto generate a name
# in the format of YYYYMMDDhhmmss.csv and append to the directory path
If (!([System.IO.Path]::HasExtension($OutputFile)))
{
	If ($OutputFile.substring($OutputFile.length - 1) -eq "\")
	{
		$OutputFile += "{0}.csv" -f (Get-Date -uformat %Y%m%d%H%M%S).ToString()
	}
	Else
	{
		$OutputFile += "\{0}.csv" -f (Get-Date -uformat %Y%m%d%H%M%S).ToString()
	}
}

ConnectProvisioningWebServiceAPI -Credential $Credential

# get Office 365 SKU info
WriteConsoleMessage -Message "Getting SKU information.  Please wait..." -MessageType "Information"
$getMsolAccountSkuResults = Get-MsolAccountSku

# iterate through the sku results
WriteConsoleMessage -Message "Processing SKU results.  Please wait..." -MessageType "Information"
$arrSkuData = @()
foreach($sku in $getMsolAccountSkuResults)
{
	$objSkuData = New-Object PSObject
	Add-Member -InputObject $objSkuData -MemberType NoteProperty -Name "AccountSkuId" -Value $sku.accountskuid
	Add-Member -InputObject $objSkuData -MemberType NoteProperty -Name "ActiveUnits" -Value $sku.activeunits
	Add-Member -InputObject $objSkuData -MemberType NoteProperty -Name "ConsumedUnits" -Value $sku.consumedunits
	Add-Member -InputObject $objSkuData -MemberType NoteProperty -Name "AvailableUnits" -Value $($sku.activeunits - $sku.consumedunits)
	Add-Member -InputObject $objSkuData -MemberType NoteProperty -Name "WarningUnits" -Value $sku.warningunits
	Add-Member -InputObject $objSkuData -MemberType NoteProperty -Name "SuspendedUnits" -Value $sku.suspendedunits
	$arrSkuData += $objSkuData
}

If ($OutputFile) {
	WriteConsoleMessage -Message "Saving results to outputfile.  Please wait..." -MessageType "Information"
	$arrSkuData | Export-Csv -Path $OutputFile -NoTypeInformation
}

# script is complete
$ScriptStopTime = Get-Date
$elapsedTime = GetElapsedTime -Start $ScriptStartTime -End $ScriptStopTime
WriteConsoleMessage -Message ("Script Start Time  :  {0}" -f ($ScriptStartTime)) -MessageType "Information"
WriteConsoleMessage -Message ("Script Stop Time   :  {0}" -f ($ScriptStopTime)) -MessageType "Information"
WriteConsoleMessage -Message ("Elapsed Time       :  {0:N0}.{1:N0}:{2:N0}:{3:N1}  (Days.Hours:Minutes:Seconds)" -f $elapsedTime.Days, $elapsedTime.Hours, $elapsedTime.Minutes, $elapsedTime.Seconds) -MessageType "Information"
WriteConsoleMessage -Message ("Output File        :  {0}" -f $OutputFile) -MessageType "Information"

Format-Table -InputObject $arrSkuData -AutoSize

# -----------------------------------------------------------------------------
#
# End of Script.
#
# -----------------------------------------------------------------------------