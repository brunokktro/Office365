<# 
    .Synopsis 
     Manages the storage limit of inidvidual mailboxes. 
    .Example 
     .\Set-MailboxStorageLimit.ps1 -Identity smithb -Level 0 
     This example sets the mailbox storage limit for smithb to use the database default. 
    .Example 
     "smithb","john.doe" | .\Set-MailboxStorageLimit.ps1 -IncreaseLevel 
     This example increases the mailbox storage limit for smithb and john.doe by one level. 
    .Description 
     The Set-MailboxStorageLimit.ps1 script is used to manage the mailbox storage limits of individual mailboxes in a standardized and uniform way. 
    .Parameter Identity 
     The Identity parameter specifies the identity of the mailbox. You can use one of the following values: 
        * GUID 
        * Distinguished name (DN) 
        * Display name 
        * Domain\Account 
        * User principal name (UPN) 
        * LegacyExchangeDN 
        * SmtpAddress 
        * Alias 
        * Mailbox object 
    .Parameter Level 
     The Level parameter specifies what predefined mailbox storage limit you wish to assign to a particular mailbox. 
     Listed below are the defined available mailbox storage limit levels. 
 
     Level        IssueWarning    ProhibitSend 
     -----      ------------    ------------ 
     0            800MB            850MB  (Default/Database Limits) 
     1            1.0GB            1.2GB 
     2            2.2GB            2.4GB 
     3            4.6GB            4.8GB 
     4            9.4GB            9.6GB 
     5            Unlimited        Unlimited 
    .Parameter IncreaseLevel 
     The IncreaseLevel parameter is used to increase the mailbox storage limit for a given mailbox up one level. 
    .Parameter DecreaseLevel 
     The DecreaseLevel parameter is used to decrease the mailbox storage limit for a given mailbox down one level. 
    .Outputs 
     None 
    .Notes 
     Name:         Set-MailboxStorageLimit.ps1 
     Author:       Jeremy Engel 
     CreatedDate:  02.02.2012 
     ModifiedDate: 03.01.2012 
     Version:      2.0.0 
  #> 
[CmdletBinding(DefaultParameterSetName="Manual")] 
Param([Parameter(Mandatory=$true,ValuefromPipeline=$true)][PSObject]$Identity, 
      [Parameter(Mandatory=$true,ParameterSetName="Manual")][ValidateRange(0,5)][int]$Level, 
      [Parameter(Mandatory=$true,ParameterSetName="DynamicUp")][switch]$IncreaseLevel, 
      [Parameter(Mandatory=$true,ParameterSetName="DynamicDown")][switch]$DecreaseLevel 
      ) 
 
begin { 
  Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue 
  $limits = @{ 0 = @{Warning=380MB;Error=390MB;Stop=400MB} 
               1 = @{Warning=700MB;Error=750MB;Stop=800MB} 
               2 = @{Warning=1.8GB;Error=1.9GB;Stop=2.0GB}  
               3 = @{Warning="Unlimited";Error="Unlimited";Stop="Unlimited"} 
               } 
  } 
process { 
  $mailbox = if($Identity -is [Microsoft.Exchange.Data.Directory.Management.Mailbox]){$Identity}else{Get-Mailbox -Identity $Identity} 
  if(!$mailbox) { return } 
  if($PSCmdlet.ParameterSetName -match "Dynamic") { 
    if($mailbox.UseDatabaseQuotaDefaults) { 
      if($DecreaseLevel) { Write-Host "ERROR: The mailbox for $mailbox is already at the minimum storage limit." -ForegroundColor Red; return } 
      else { $Level = 0 } 
      } 
    else { 
      $limit = $mailbox.ProhibitSendQuota.Value 
      if(!$limit) { 
        if($IncreaseLevel) { Write-Host "ERROR: The mailbox for $mailbox is already at the maximum storage limit level." -ForegroundColor Red; return } 
        else { $Level = $limits.Count-1 } 
        } 
      else { for($i=0;$i-lt$limits.Count-1;$i++) { if($limit -le $limits[$i].Error+1MB) { $Level = $i; break } } } 
      } 
    $Level = if($IncreaseLevel){$Level+1}elseif($DecreaseLevel){$Level-1} 
    } 
  $warning = if($Level -eq 0){"DatabaseDefault"}elseif($Level -eq $limits.Count-1){$limits[$Level].Warning}else{New-Object Microsoft.Exchange.Data.ByteQuantifiedSize($limits[$Level].Warning)} 
  $limit = if($Level -eq 0){"DatabaseDefault"}elseif($Level -eq $limits.Count-1){$limits[$Level].Error}else{New-Object Microsoft.Exchange.Data.ByteQuantifiedSize($limits[$Level].Error)} 
  $stop = if($Level -eq 0){"DatabaseDefault"}elseif($Level -eq $limits.Count-1){$limits[$Level].Stop}else{New-Object Microsoft.Exchange.Data.ByteQuantifiedSize($limits[$Level].Stop)}
  if($Level -eq 0) { $mailbox | Set-Mailbox -UseDatabaseQuotaDefaults $true -IssueWarningQuota "Unlimited" -ProhibitSendQuota "Unlimited" } 
  else { $mailbox | Set-Mailbox -UseDatabaseQuotaDefaults $false -IssueWarningQuota $limits[$Level].Warning -ProhibitSendQuota $limits[$Level].Error -ProhibitSendReceiveQuota $limits[$Level].Stop} 
  $item = New-Object PSObject 
  $item | Add-Member -MemberType NoteProperty -Name DisplayName -Value $mailbox.DisplayName 
  $item | Add-Member -MemberType NoteProperty -Name StorageLevel -Value $Level 
  $item | Add-Member -MemberType NoteProperty -Name IssueWarningQuota -Value $warning 
  $item | Add-Member -MemberType NoteProperty -Name ProhibitSendQuota -Value $limit 
  $item | Add-Member -MemberType NoteProperty -Name ProhibitSendReceiveQuota -Value $stop 
  $item | Add-Member -MemberType ScriptMethod -Name ToString -Value { $this.Alias } -Force 
  $item 
  } 
end { }