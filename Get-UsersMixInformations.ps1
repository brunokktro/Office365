#### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### ####
#### 
#### Information to be extracted is:
#### From Get-User – Title, FirstName, LastName, Department, Phone
#### From Get-Mailbox - PrimarySMTPAddress, RecipientTypeDetails, EmailAddresses, HiddenFromAddressLists, Database
#### From Get-MailboxStatistics - LastLogonTime, TotalItemSize, ItemCount
#### 
#### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### #### 

Get-User -ResultSize Unlimited -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Select-Object DisplayName,Company,FirstName,LastName,@{label="PrimarySMTPAddress";expression={(Get-Mailbox $_).PrimarySMTPAddress}},Title,Department,Phone,@{label="RecipientTypeDetails";expression={(Get-Mailbox $_).RecipientTypeDetails}},@{Name="EmailAddresses";Expression={(Get-Mailbox $_).EmailAddresses -Join "`n"}},@{label="HiddenFromAddressLists";expression={(Get-Mailbox $_).HiddenFromAddressLists}},@{label="LastLogonTime";expression={(Get-MailboxStatistics $_).LastLogonTime}},@{label="TotalItemSize(MB)";expression={(Get-MailboxStatistics $_).TotalItemSize.Value.ToMB()}}, @{label="ItemCount";expression={(Get-MailboxStatistics $_).ItemCount}},@{label="Database";expression={(Get-Mailbox $_).Database}}| Export-Csv C:\Scripts\UserData.Csv -NoTypeInformation