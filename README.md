<div align="center">

## Create Excel file using ADOX


</div>

### Description

This sample shows how create Excel file using ADOX. In database apps when ADO and ADOX is used it's simple way to create 'Excel reports'. Using ADOX is about 3 times faster than Excel Automation. If you find this code useful, please vote...
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Grzegorz P\.](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/grzegorz-p.md)
**Level**          |Intermediate
**User Rating**    |4.8 (38 globes from 8 users)
**Compatibility**  |VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/grzegorz-p-create-excel-file-using-adox__1-52806/archive/master.zip)





### Source Code

```
Public Function SaveRecordsetAsExcelFile(ByRef SourceRecordset As ADODB.Recordset, _
                     ByVal ExcelFileName As String, _
                     ByVal WorksheetName As String) As Boolean
 'Don't forget to add reference to Microsoft ADO 2.8 and ADOX 2.8 Libraries
 Dim cnnExcel As ADODB.Connection
 Dim catExcel As ADOX.Catalog
 Dim tblWorksheet As ADOX.Table
 Dim rstExcelData As ADODB.Recordset
 Dim fldColumnHeader As ADODB.Field
 Dim strWkshtName As String
  On Error GoTo EH_SaveRecordsetAsExcelFile
  'Create Excel file and worksheet
  Set cnnExcel = New ADODB.Connection
  Set catExcel = New ADOX.Catalog
  Set tblWorksheet = New ADOX.Table
  cnnExcel.CursorLocation = adUseClient
  cnnExcel.Provider = "Microsoft.Jet.OLEDB.4.0"
  cnnExcel.Properties("Extended Properties") = "Excel 8.0"
  cnnExcel.Open "Data Source = " & ExcelFileName
  Set catExcel.ActiveConnection = cnnExcel
  tblWorksheet.Name = WorksheetName
  For Each fldColumnHeader In SourceRecordset.Fields
    tblWorksheet.Columns.Append fldColumnHeader.Name, fldColumnHeader.Type
  Next 'fldColumnHeader
  catExcel.Tables.Append tblWorksheet
  Set tblWorksheet = Nothing
  Set catExcel = Nothing
  Set cnnExcel = Nothing
  'Fill worksheet with data
  Set cnnExcel = New ADODB.Connection
  Set rstExcelData = New ADODB.Recordset
  With cnnExcel
    .CursorLocation = adUseClient
    .Provider = "Microsoft.Jet.OLEDB.4.0"
    .Properties("Extended Properties") = "Excel 8.0"
    .Open ExcelFileName
    strWkshtName = "[" & WorksheetName & "$]"
    With rstExcelData
      Set .ActiveConnection = cnnExcel
      .CursorLocation = adUseClient
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Source = strWkshtName
      .Open
    End With 'rstExcelData
    With SourceRecordset
      .MoveFirst
      Do While Not .EOF
        rstExcelData.AddNew
          For Each fldColumnHeader In .Fields
            rstExcelData.Fields(fldColumnHeader.Name) = fldColumnHeader 'insert value
          Next 'fldColumnHeader
        rstExcelData.Update
        .MoveNext
      Loop
    End With 'SourceRecordset
    .Close 'cnnExcel
  End With 'cnnExcel
  Set cnnExcel = Nothing
  Set rstExcelData = Nothing
  Set fldColumnHeader = Nothing
  SaveRecordsetAsExcelFile = True
Exit Function
EH_SaveRecordsetAsExcelFile:
  SaveRecordsetAsExcelFile = False
  Set tblWorksheet = Nothing
  Set catExcel = Nothing
  Set cnnExcel = Nothing
  Set rstExcelData = Nothing
  Set fldColumnHeader = Nothing
End Function
```

