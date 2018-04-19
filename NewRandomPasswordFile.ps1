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
# PowerShell source code
#
# NAME:
#    NewRandomPasswordFile.ps1
#
#------------------------------------------------------------------------------


<#
	.SYNOPSIS
		Generates a random password by referencing a CSV file which contains the
		User Principal Name of a list of users.  The script will create a new CSV
		file containing the original UPN and new password.

	.DESCRIPTION
		Generates a random password by referencing a CSV file which contains the
		User Principal Name of a list of users.  The script will create a new CSV
		file containing the original UPN and new password.  

	.PARAMETER CsvFile
		Specifies the path and file name of the CSV file.

	.PARAMETER Outputfile
		Specifies the name of the output file.  The arguement can be the full path including the file
		name, or only the path to the folder in which to save the file (uses default name).
		
		Default filename is in the format of "YYYYMMDDhhmmss_NewRandomPasswordFile.csv"

	.EXAMPLE
		PS> NewRandomPasswordFile -CsvFile "C:\Folder\Sub Folder\File name.csv"

	.INPUTS
		System.String

	.OUTPUTS
		CSV File

	.NOTES
		A CSV file of UPN's can	be created from Office 365 user objects by using the following command:
		
		PS> Get-MsolUser -All | Select-Object userprincipalname | export-csv results.csv -NoTypeInformation

#>


[CmdletBinding()]
param
(
	[Parameter(Mandatory = $True, Position = 0)]
	[string]$CsvFileName,
	
	[Parameter(Mandatory = $False)]
	[string]$OutputFile = "$((Get-Date -uformat %Y%m%d%H%M%S).ToString())_NewRandomPasswordFile.csv"
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
			None

		.NOTES

	#>

	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $True)]
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


Function New-RandomPassword
{
	[Void][Reflection.Assembly]::LoadWithPartialName(”System.Security.Cryptography”)
	[Void][Reflection.Assembly]::LoadWithPartialName(”System.Byte”)

	[int32]$PasswordLength = 8
	[string]$Numbers = "0123456789"
	[string]$Consonants = "bcdfghjklmnpqrstvwxyz"
	[Int32]$i = 0
	[Int32]$indexOffSet = 0

	$RNGCryptoServiceProvider = New-Object System.Security.Cryptography.RNGCryptoServiceProvider
	$arrBytes = New-Object System.Byte[] $PasswordLength
	$arrCharacters = New-Object System.Char[] $PasswordLength

	# get random byte values and save to byte array
	$RNGCryptoServiceProvider.GetBytes($arrBytes)

	# first character is an upper case consonant
	$indexOffSet = [Int32]$arrBytes[$i] % $Consonants.Length
	$arrCharacters[$i] = ([string]$Consonants[$indexOffSet]).ToUpper()
	$i++

	# second through fourth characters are a lower case consonant
	for($i; $i -le 3 ; $i++)
	{
		$indexOffSet = [int32]$arrBytes[$i] % $Consonants.Length
		$arrCharacters[$i] = $Consonants[$indexOffSet]
	}

	# remaining characters are numbers
	for ($i; $i -le ($PasswordLength - 1); $i++)
	{
	    $indexOffSet = [int32]$arrBytes[$i] % $Numbers.Length
	    $arrCharacters[$i] = $Numbers[$indexOffSet]
	}

	# convert character array to a string
	[string]$Password = [string]::Join("",$arrCharacters)

	Return $Password
}


# -----------------------------------------------------------------------------
#
# Main Script Execution
#
# -----------------------------------------------------------------------------

$Error.Clear()
$ScriptStartTime = Get-Date

$CsvFile = Import-Csv -Path $CsvFileName

# verify output directory exists for results file
WriteConsoleMessage -Message ("Verifying folder:  {0}" -f $OutputFile) -MessageType "Verbose"
If (!(TestFolderExists $OutputFile))
{
	WriteConsoleMessage -Message ("Directory not found:  {0}" -f $OutputFile) -MessageType "Error"
	Exit
}

# iterate through collection of users and generate a random password
WriteConsoleMessage -Message "Processing collection of users.  Please wait..." -MessageType "Information"
$arrUsers = @()
$CsvFile | ForEach-Object -Begin {
	$count = 1
} -Process {
	$ActivityMessage = "Creating a new password file, Please wait..."
	$StatusMessage = ("Processing: {0}" -f ($_.UserPrincipalName))
	$PercentComplete = ($count / @($CsvFile).count * 100)
	Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
	
	$objUser = New-Object PSObject
	
	$NewPassword = New-RandomPassword
	
	Add-Member -InputObject $objUser -MemberType NoteProperty -Name UserPrincipalName -Value $($_.userprincipalname)
	Add-Member -InputObject $objUser -MemberType NoteProperty -Name NewPassword -Value $NewPassword
	
	$arrUsers += $objUser
	$count++
}

WriteConsoleMessage -Message "Saving results to outputfile.  Please wait..." -MessageType "Information"
If ($OutputFile) {$arrUsers | Export-Csv -Path $OutputFile -NoTypeInformation}

# script is complete
$ScriptStopTime = Get-Date
$elapsedTime = GetElapsedTime -Start $ScriptStartTime -End $ScriptStopTime
WriteConsoleMessage -Message ("Script Start Time  :  {0}" -f ($ScriptStartTime)) -MessageType "Information"
WriteConsoleMessage -Message ("Script Stop Time   :  {0}" -f ($ScriptStopTime)) -MessageType "Information"
WriteConsoleMessage -Message ("Elapsed Time       :  {0:N0}.{1:N0}:{2:N0}:{3:N1}  (Days.Hours:Minutes:Seconds)" -f $elapsedTime.Days, $elapsedTime.Hours, $elapsedTime.Minutes, $elapsedTime.Seconds) -MessageType "Information"
WriteConsoleMessage -Message ("Output File        :  {0}" -f $OutputFile) -MessageType "Information"

# -----------------------------------------------------------------------------
#
# End of Script.
#
# -----------------------------------------------------------------------------