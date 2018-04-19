
Get-QuarantineMessage -RecipientAddress user@domain.com | Release-QuarantineMessage

Get-QuarantineMessage -SenderAddress sender@domain.com | Release-QuarantineMessage -ReleaseToAll

Get-QuarantineMessage -SenderAddress sender@domain.com | Release-QuarantineMessage -ReleaseToAll -ReportFalsePositive