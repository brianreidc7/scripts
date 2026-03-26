get-aduser -Filter {admincount -gt 0} -Properties adminCount -ResultSetSize $null | FT DistinguishedName,Enabled,SamAccountName

