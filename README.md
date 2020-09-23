<div align="center">

## Very useful Oracle/VB ADO samples


</div>

### Description

Need to do any oracle/ADO work? I wrote these to help me along in my projects. I hope you find them useful too.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Qucami](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/qucami.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/qucami-very-useful-oracle-vb-ado-samples__1-36245/archive/master.zip)





### Source Code

```
Function ConnectToOracle(ByVal sWorld As String, ByVal sUID As String, ByVal sPWD As String) As String
'******************************************************************************************
'***** Connection to Oracle using Oracle OLE driver
'*****
On Error GoTo Ouch
  P11D_DB.Open "Provider=OraOLEDB.Oracle;data source=" & _
	sWorld & ".World;User id=" & sUID & ";password=" & sPWD & ";"
  ConnectToOracle = ""
  Exit Function
Ouch:
ConnectToOracle = Err.Description & " (" & Err.Number & ")"
End Function
'-------------------------------------------------------------------------------------------------------
Sub CloseConnectionToOracle()
'******************************************************************************************
'***** Close Connection to Oracle
'*****
On Error Resume Next
  If P11D_DB.State <> 0 Then
    P11D_DB.Close
  End If
End Sub
'-------------------------------------------------------------------------------------------------------
Function OracleDate(dIn As Date) As String
'******************************************************************************************
'***** Insert/Update/Retrieve an oracle date in it's proper format
'***** sSQl=".... where DATE_COL = " & oracledate(VBDateField) & "....."
  OracleDate = "to_date('" & Format(dIn, "dd/mm/yyyy") & "','dd/mm/yyyy')"
End Function
'-------------------------------------------------------------------------------------------------------
Public Function GetColumnData() As String()
'******************************************************************************************
'***** Return a column of data via an array
'*****
Dim sColRetr() As String
Dim rsColRetr As New ADODB.Recordset
Dim sSQL As String
Dim x As Integer
  sSQL = "select COLUMN_NAME from TABLE"
  rsColRetr.Open sSQL, ADO_Connection, adOpenStatic, adLockReadOnly
  ReDim sColRetr(rsColRetr.RecordCount)
  x = 0
  While Not rsColRetr.EOF
    sColRetr(x) = rsColRetr!band_description
    rsColRetr.MoveNext
    x = x + 1
  Wend
  rsColRetr.Close
  Set rsColRetr = Nothing
  ReDim preserve sColRetr(ubound(sColRetr)-1)
  GetColumnData = sColRetr
End Function
'-------------------------------------------------------------------------------------------------------
Sub OracleCommit()
'******************************************************************************************
'***** Commit inserts and updates
'*****
On Error Resume Next
Dim rsCMD As New ADODB.Command
  With rsCMD
    .ActiveConnection = P11D_DB
    .CommandText = "commit"
    .Execute
  End With
  Set rsCMD = Nothing
End Sub
'-------------------------------------------------------------------------------------------------------
Function GetDescForTable(ByVal sTable As String, ByVal sOwner As String) As String()
'******************************************************************************************
'***** Get the Column names for a table
'*****
Dim TD() As String
Dim rsD As New ADODB.Recordset
Dim sSQL As String
  sSQL = "select column_name " & _
      "from dba_tab_columns where owner = '" & sOwner & "' " & _
      "and table_name = '" & sTable & "'"
  rsD.Open sSQL, ADO_Connection, adOpenStatic, adLockReadOnly
  ReDim TD(0)
  While Not rsD.EOF
    ReDim Preserve TD(UBound(TD) + 1)
    TD(UBound(TD) - 1) = rsD!column_name
    rsD.MoveNext
  Wend
  rsD.Close
  ReDim Preserve TD(UBound(TD) - 1)
  GetDescForTable = TD
End Function
'-------------------------------------------------------------------------------------------------------
Function GetTables(ByVal sOwner as string) As String()
'******************************************************************************************
'***** Get the Table names for an owner
'*****
Dim TL() As String
Dim rs As New ADODB.Recordset
Dim sSQL As String
  sSQL = "select table_name from sys.all_tables where owner = '" & sOwner & "'"
  rs.Open sSQL, ADO_Connection, adOpenStatic, adLockReadOnly
  ReDim TL(0)
  While Not rs.EOF
    ReDim Preserve TL(UBound(TL) + 1)
    TL(UBound(TL) - 1) = rs!table_name
    rs.MoveNext
  Wend
  rs.Close
  Set rs = Nothing
  ReDim Preserve TL(UBound(TL) - 1)
  GetTables = TL
End Function
'-------------------------------------------------------------------------------------------------------
Function HandleQuotes(ByVal sIn As String) As String
'******************************************************************************************
'***** take care of single quotes on record update/retrieval to handle data like Mike O'Sullivan
'*****
HandleQuotes = Replace(sIn, "'", "''")
End Function
'-------------------------------------------------------------------------------------------------------
Function ScrNull(sIn As Variant) As String
'******************************************************************************************
'***** when referencing a recordset field wrap it with this function to return a ""
'***** to a string where the column data held a null, eg; sString=ScrNull(rsCol!Column_Data
'*****
  If IsNull(sIn) Then
    ScrNull = ""
  Else
    ScrNull = sIn
  End If
End Function
'-------------------------------------------------------------------------------------------------------
Function GetTotal() As Double
'******************************************************************************************
'***** Get the total of a column of data
'*****
Dim rsFT As New ADODB.Recordset
Dim sSQL As String
  sSQL = "select sum(COLUMN_DATA) as FC_Total from TABLE where ...condition..."
  rsFT.Open sSQL, P11D_DB, adOpenStatic, adLockReadOnly
  If rsFT.EOF Then
    GetTotal = 0
  ElseIf IsNull(rsFT!fc_total) Then
    GetTotal = 0
  Else
    GetTotal = CDbl(rsFT!fc_total)
  End If
  rsFT.Close
  Set rsFT = Nothing
End Function
```

