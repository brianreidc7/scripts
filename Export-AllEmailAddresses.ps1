Get-Recipient * -ResultSize Unlimited | Select-Object Name,DisplayName,recipienttype,RecipientTypeDetails,ExternalEmailAddress,PrimarySmtpAddress,@{Name="EmailAddresses";Expression={ ($_.EmailAddresses | Where-Object {$_ -cmatch "smtp:*"} | ForEach-Object {$_ -replace 'smtp:' }) -join ';' }} | 
Export-Csv "m365-email-addresses.csv" -NoTypeInformation -Encoding UTF8
