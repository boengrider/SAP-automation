# HADES Drag and Drop

## Uploads files to SAP via OAWD transaction

HADES.au3 -s|--system -c|--client -src|--source -cc|--companycode -oawd [-nosub|--nosubdir] [-as|--archivesharepoint] [-nal|--noarchlocal]  
```
-s | --system <- Mandatory parameter  
-c | --client <- Mandatory parameter  
-src | --source <- Mandatory parameter  
-cc | --companycode <- Mandatory parameter  
-oawd <- Mandatory parameter  
-nosub | --nosubdir <- Optional parameter  
-as | --archivesharepoint <- Optional parameter  
-nal | --noarchlocal <- Optional parameter  
```  
  
  
### Example Call #1  
-s fq2  -c 105  -oawd "SI01 FI Scan"  -src C:\!AUTO\SI01_HADES_DND_NET  -cc si01  -nal	-nosub	-as  https://TENANT.sharepoint.com/sites/unit-hades/SI01_HADES_ARCHIVE  
```
-s SAP System nane
-c SAP Client name
-oawd OAWD folder name
-src Source folder where to look for files to be uploaded
-cc Company code
-nal Do not archive uploaded files locally
-nosub Do not process subdirectories in the source directory. Only process source dir 
-as Archive uploaded file to the sharepoint library
```

### Example Call #2
-s fq2 -c 105 -src https://TENANT.sharepoint.com/sites/unit-hades/SI01_HADES_SOURCE -cc si01 -nal -nosub -as https://TENANT.sharepoint.com/sites/unit-hades/SI01_HADES_ARCHIVE -oawd "SI01 FI Scan"
### Example Call #3
-s fq2 -c 105 -oawd "SI01 FI Scan" -src \\\Czpragn006\hades_qa -cc si01 -as https://TENANT.sharepoint.com/sites/unit-hades/SI01_HADES_ARCHIVE -oawd "SI01 FI Scan"  
### Example Call #4
-s fq2 -c 105 -oawd "SI01 FI Scan" -src \\\10.229.128.6\mk01_hades_dnd -cc si01 -as https://TENANT.sharepoint.com/sites/unit-hades/SI01_HADES_ARCHIVE

### Exit codes
```
1-99 - Something other than errors e.g no parameters passed
100 - General script errors
200 - SAP related errors
300 - Sharepoint related errors
```
