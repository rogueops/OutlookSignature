<#
Run the following command via EAC powershell terminal to add the signature generated to OWA
Generation of the signature file can be done using the first script
#>
$mailboxes = Get-Mailbox
$mailboxes| foreach {$file= "C:\signatures" + ($_.alias) + ".htm"; Set-MailboxMessageConfiguration -identity $_.alias -SignatureHtml "$(Get-Content -Path $file -ReadCount 0)" -autoaddsignature $true 
