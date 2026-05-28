$Addr= Read-Host "Enter full IMCEAEX address"
$Repl= @(@("_","/"), @("\+20"," "), @("\+28","("), @("\+29",")"), @("\+2C",","), @("\+5F", "_" ), @("\+40", "@" ), @("\+2E", "." ))
$Repl | ForEach { $Addr= $Addr -replace $_[0], $_[1] }
$Addr= "X500:$Addr" -replace "IMCEAEX-","" -replace "@.*$", ""
Write-Host $Addr