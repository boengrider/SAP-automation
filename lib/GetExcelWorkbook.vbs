Function GetExcelWorkbook(ByVal reportName, ByRef outWorkbook)
          
	  On Error Resume Next
      Dim waitTime : waitTime = 500
      Dim waitTurns : waitTurns = 5
      Dim turn : turn = 0

      err.Clear
      Dim excel : set Excel = GetObject(,"Excel.Application")
    
       Do While err.Number <> 0
         If turn > waitTurns Then
            Exit Do
         End if  

         wscript.sleep waitTime
         turn = turn + 1

         err.Clear
         set excel = GetObject(,"Excel.Application")
       Loop

      If Not isobject(excel) Then
          GetExcelWorkbook = 0
          Exit Function
      End If 

      'Excel instance exists
      Dim workbook__
     
      For each workbook__ in excel.Workbooks
        if workbook__.Name = reportName Then
          set outWorkbook = workbook__
          GetExcelWorkbook = 1
          Exit Function
        end if 
      Next

      GetExcelWorkbook = 0
      

End Function

Dim wb
If GetExcelWorkbook("report.xlsx",wb) Then 
   Wscript.echo "Workbook found"
   Wscript.echo wb.Name & " " & wb.fullname
else 
   wscript.echo "Workbook not found"
end if 





