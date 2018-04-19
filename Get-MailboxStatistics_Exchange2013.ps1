########################################################### 
# AUTOR  : Bruno Lopes - Axter Soluções em TI 
# DATA    : 04/06/2014
# 
# SOBRE : Este script realiza a obtenção da listagem de usuários
#         e mailboxes, com os respectivos tamanhos das mesmas no
#         Exchange 2013.
########################################################### 

$Mailboxes = Get-Mailbox -ResultSize Unlimited | Select UserPrincipalName, identity, ArchiveStatus, DisplayName, PrimarySmtpAddress
 
$MailboxSizes = @()
 
foreach ($Mailbox in $Mailboxes) {
 
                $ObjProperties = New-Object PSObject
               
                $MailboxStats = Get-MailboxStatistics $Mailbox.UserPrincipalname | Select LastLogonTime, TotalItemSize, ItemCount
               
                Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "UserPrincipalName" -Value $Mailbox.UserPrincipalName
                Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "DisplayName" -Value $Mailbox.DisplayName
                Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "PrimarySmtpAddress" -Value $Mailbox.PrimarySmtpAddress
                Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "Last Logged In" -Value $MailboxStats.LastLogonTime
                Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "Mailbox Size" -Value $MailboxStats.TotalItemSize
                Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "Mailbox Item Count" -Value $MailboxStats.ItemCount
               
                if ($Mailbox.ArchiveStatus -eq "Active") {
               
                                $ArchiveStats = Get-MailboxStatistics $Mailbox.UserPrincipalname -Archive | Select TotalItemSize, ItemCount
                               
                                Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "Archive Size" -Value $ArchiveStats.TotalItemSize
                                Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "Archive Item Count" -Value $ArchiveStats.ItemCount
 
                }
                else {
               
                                Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "Archive Size" -Value "No Archive"
                                Add-Member -InputObject $ObjProperties -MemberType NoteProperty -Name "Archive Item Count" -Value "No Archive"
                               
                }
               
                $MailboxSizes += $ObjProperties
 
}             
               
$MailboxSizes | Export-Csv C:\Scripts\MailboxesStatistics.csv -Encoding Unicode -NoTypeInformation