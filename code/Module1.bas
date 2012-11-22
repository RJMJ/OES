Attribute VB_Name = "Module1"
Public roll As Integer
Public sname As String
Public qsubject As String
Public sclass As Integer
Public stream As String
Public tim As Integer
Public timstatus As Integer     ' to calculate time
Public subchoice As Integer     ' for checking subchoice
Public resulttype As Integer
Public cn As ADODB.Connection
Public rec As ADODB.Recordset
Public rptroll As Integer
Public rptflag As Integer

Public Sub Connection1()
Set cn = New ADODB.Connection
cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source =" & "data.mdb"
cn.CursorLocation = adUseClient
'cn.ConnectionString = "Provider=OraOLEDB.Oracle;dbq=localhost:1521/XE;Database=XE;User Id=dheeraj;Password=joshi;"
cn.Open
End Sub
