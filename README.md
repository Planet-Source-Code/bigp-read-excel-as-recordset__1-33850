<div align="center">

## Read Excel as Recordset


</div>

### Description

An alternative method of reading an MS Excel Spreadsheet.
 
### More Info
 
You must supply the full path and file name of the Excel Sheet you wish to read.

It returns the Specified Sheets' data in an ADO recordset.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[BigP](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bigp.md)
**Level**          |Advanced
**User Rating**    |4.6 (51 globes from 11 users)
**Compatibility**  |VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bigp-read-excel-as-recordset__1-33850/archive/master.zip)





### Source Code

```
Dim cn As ADODB.Connection
Dim rsADO As New ADODB.Recordset
Dim strSQL As String
Dim strPath as string
Set cn = New ADODB.Connection
strPath = '[ADD FULL PATH AND FILE NAME]
With cn
  .Provider = "MSDASQL"
  .ConnectionString = "Driver={Microsoft Excel Driver (*.xls)};" & _
  "DBQ=" & strPath & " ; ReadOnly=false;MaxScanRows= 0;"
  .Open
End With
  ' Specify Sheet Name and Cell Range
  strSQL = "SELECT * FROM [Sheet1$A1:Z10]"
  rsADO.Open strSQL, cn
  Do while not rs.EOF
  	' Add code here to work with recordset
  rsADO.MoveNext
  Loop
Set cn = Nothing
Set rsADO = Nothing
```

