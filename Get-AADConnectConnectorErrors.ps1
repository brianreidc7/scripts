# Cmd prompt stuff to export errors from AADConnect

del %temp%\Errors-Export.csv
del %temp%\Errors-Export.xml

cd "C:\Program Files\Microsoft Azure AD Sync\Bin"
C:

# Export Synchronization Errors from named connector
csexport.exe sys.element.com %temp%\Errors-Export.xml /f:i
CSExportAnalyzer.exe %temp%\Errors-Export.xml > %temp%\Errors-Export.csv

#or
csexport.exe exova-eu.local %temp%\Errors-Export.xml /f:i
CSExportAnalyzer.exe %temp%\Errors-Export.xml > %temp%\Errors-Export.csv

#or
csexport.exe "elementmaterialstechno.onmicrosoft.com - AAD" %temp%\Errors-Export.xml /f:e
CSExportAnalyzer.exe %temp%\Errors-Export.xml > %temp%\Errors-Export.csv


# show results

notepad %temp%\Errors-Export.csv