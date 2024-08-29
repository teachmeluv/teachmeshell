﻿$boxaccess =Get-Mailbox -ResultSize Unlimited | ForEach-Object {Get-MailboxPermission -Identity $_.guid} | Get-MailboxPermission -User "username"
$boxaccess | Export-Csv -Path C:\Script\Dateien\User_boxaccess.csv -NoTypeInformation -Delimiter ";" -Encoding UTF8