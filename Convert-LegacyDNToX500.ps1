$Addr= Read-Host "Enter full IMCEAEX/EX address in the email NDR"
$Repl= @(@("_","/"), @("\+20"," "), @("\+28","("), @("\+29",")"), @("\+2C",","), @("\+5F", "_" ), @("\+40", "@" ), @("\+2E", "." ))
$Repl | ForEach-Object { $Addr= $Addr -replace $_[0], $_[1] }
$Addr= "$Addr" -replace "EX:","" -replace "@.*$", ""
$Addr= "X500:$Addr" -replace "IMCEAEX-","" -replace "@.*$", ""
Write-Host $Addr