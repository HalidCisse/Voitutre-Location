Attribute VB_Name = "DB"
Global Connect As ADODB.Connection, Rcs As New ADODB.Recordset

Sub OpenData(SQL As String)

On Error Resume Next
    Connect.Close
    Connect.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "/db.mdb"
    Set Rcs = Connect.Execute(SQL)

End Sub

Sub CloseData()
    
    On Error GoTo x
    Rcs.Close
    Connect.Close
    Set Rcs = Nothing
    Set Connect = Nothing
x:
     
End Sub

Function DoSQL(SQL As String) As Integer

    On Error Resume Next
      Connect.Close
    On Error GoTo x
      Connect.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "/db.mdb"
      Connect.Execute (SQL), DoSQL
x:

End Function

