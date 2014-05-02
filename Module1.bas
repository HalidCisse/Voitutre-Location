Attribute VB_Name = "Module1"
Public cn As New ADODB.Connection
Function opencn(fdb As String) As Boolean
Set cn = New ADODB.Connection
cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & fdb
cn.Open
If cn.State = adStateOpen Then
opencn = True
Else
opencn = False
End If
End Function

