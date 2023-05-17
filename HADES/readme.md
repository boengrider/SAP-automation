# HADES Drag and Drop
Example Call #1  
-s fq2  -c 105  -oawd "SI01 FI Scan"  -src C:\!AUTO\SI01_HADES_DND_NET  -cc si01  -nal	-nosub	-as  https://TENANT.sharepoint.com/sites/unit-hades/SI01_HADES_ARCHIVE  


Example Call #2
-s fq2 -c 105 -src https://TENANT.sharepoint.com/sites/unit-hades/SI01_HADES_SOURCE -cc si01 -nal -nosub -as https://TENANT.sharepoint.com/sites/unit-hades/SI01_HADES_ARCHIVE -oawd "SI01 FI Scan"
Example Call #3
-s fq2 -c 105 -oawd "SI01 FI Scan" -src \\Czpragn006\hades_qa -cc si01 -as https://TENANT.sharepoint.com/sites/unit-hades/SI01_HADES_ARCHIVE -oawd "SI01 FI Scan"  
Example Call #4
-s fq2 -c 105 -oawd "SI01 FI Scan" -src \\10.229.128.6\mk01_hades_dnd -cc si01 -as https://TENANT.sharepoint.com/sites/unit-hades/SI01_HADES_ARCHIVE

Exit codes
1-99 - Something other than errors e.g no parameters passed
100 - General script errors
200 - SAP related errors
300 - Sharepoint related errors
