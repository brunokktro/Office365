#Copyright © 2012 Microsoft Corporation.  All rights reserved.
#
# THIS CODE AND ANY ASSOCIATED INFORMATION ARE PROVIDED “AS IS” WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR 
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR
# PURPOSE. THE ENTIRE RISK OF USE, INABILITY TO USE, OR RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
#

<#
	.SYNOPSIS
		Generates a CSV report containing both general and mailbox related information about
		users in Office 365.

	.DESCRIPTION
		AssignLicenseByDG is an example script with can be used to automatically assign licenses based on a membership within 
		a specified distribution group which aligns to an Office 365 license type. 

	.PARAMETER Credential
		Specifies the credential to use when connecting to the Office 365 PowerShell web service
		using Connect-MsolService, and when connecting to Exchange Online (https://ps.outlook.com/powershell).

	.PARAMETER License
		Specifies the license name associated with a distribution group maintained within the on-premises directory.
		Accepts one of the following values:		
			ENTERPRISEPACK
			DESKLESSWOFFPACK
			DESKLESSPACK
			LITEPACK
			
			
	.EXAMPLE
		PS> .\AssignLicenseByDG.ps1 -LicenseType "ENTERPRISEPACK"

	.EXAMPLE
		PS> .\AssignLicenseByDG.ps1 -LicenseType "ENTERPRISEPACK" -Credential $cred


	.NOTES

#>



[CmdletBinding()]
param
(
	[Parameter(Mandatory = $False)]
	[System.Management.Automation.PsCredential]$Credential = $Host.UI.PromptForCredential("Enter Credential",
		"Enter the username and password of an MSOnline administrator account.",
		"",
		"userCreds"),
	
	[Parameter(Mandatory = $true, helpmessage="Enter one of the following valid license types: ENTERPRISEPACK, DESKLESSWOFFPACK, DESKLESSPACK or LITEPACK")]
	[ValidateSet("ENTERPRISEPACK", "DESKLESSWOFFPACK", "DESKLESSPACK","LITEPACK")]
	[String]$DGName
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


# -----------------------------------------------------------------------------
#
# Main Script Execution
# main()
# -----------------------------------------------------------------------------

$Error.Clear()
# connect to MSOnline PowerShell Web Service
WriteConsoleMessage -Message "Connecting to MSOnline web service.  Please wait..." -MessageType "Information"
Connect-MsolService -Credential $Credential
If($? -eq $False){Exit}

# connect to Exchange Online (Outlook.com) remote PowerShell Web Service
If (!(Get-PSSession | Where-Object {If (($_.configurationname -eq "microsoft.exchange") -and ($_.computername -like "*.outlook.com") -and ($_.runspace.runspacestateinfo.state -eq "opened")) {Return $True}}))
{
	WriteConsoleMessage -Message "Connecting to Exchange Online (http://ps.outlook.com/powershell).  Please wait..." -MessageType "Information"
	$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://ps.outlook.com/powershell/" -Credential $Credential -Authentication "Basic" -AllowRedirection
	Import-PSSession $session -AllowClobber 
}
Else
{
	WriteConsoleMessage -Message "Existing connection to Exchange Online (http://ps.outlook.com/powershell) detected.  Using existing connection and corresponding credential." -MessageType "Warning"
}

$CompanyInfo=Get-MSOLCompanyInformation
$CompanyName=$CompanyInfo.DisplayName
$LicenseName=$CompanyName+":"+$DGName.ToUpper()
$DGList=Get-DistributionGroup
$DGMembers = $null

foreach ($DG in $DGList)
{
	if ($DG.Name.ToUpper()-eq $DGName.ToUpper())
	{
	$DGMembers = Get-DistributionGroupMember $DGName
	break
	}
}

if ($DGmembers -eq $null)
{
	WriteConsoleMessage -Message "Distribution group named " $DGName " does not exist" -MessageType "Error" 
	Exit
}


foreach($member in $DGMembers)
{
    $msoUserLicense=(get-msoluser -UserPrincipalName (get-user $member.name).userprincipalname).Licenses
    
    foreach($License in $msoUserLicense)
    {
        if($license.AccountSkuId -eq $LicenseName)
        {
            "User $member.Name already is licensed with $LicenseName"
        }
        else
        {
            If($msoUserLicense.Count -ne 0)
            {
                # user has a license we need to switch the license
                Set-MsolUserLicense -UserPrincipalName (get-user $member.name).userprincipalname -RemoveLicenses $License.AccountSkuId -AddLicenses $LicenseName
            }
            else
            {
                # user ha no license
                Set-MsolUserLicense -UserPrincipalName (get-user $member.name).userprincipalname -AddLicenses $LicenseName
            
            }
        }     
    }
}