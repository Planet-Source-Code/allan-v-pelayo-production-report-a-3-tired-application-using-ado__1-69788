VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReadFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Function PROC_READ_TXT_FILE_SERVER() As ADODB.Recordset
On Error GoTo PROC_ERR
      Dim rs As ADODB.Recordset
      Set rs = New ADODB.Recordset
      Dim conn As ADODB.Connection
      Set conn = New ADODB.Connection
      conn.Open "DRIVER={Microsoft Text Driver (*.txt; *.csv)};" & _
                  "DBQ=" & App.Path & ";", "", ""

      rs.Open "select * from [server.txt]", conn, adOpenStatic, _
                  adLockReadOnly, adCmdText
      Set PROC_READ_TXT_FILE_SERVER = rs
      Set rs = Nothing
      Set conn = Nothing
Exit Function
PROC_ERR:
    MsgBox err.Description
End Function
Public Function PROC_READ_TXT_FILE_DATABASE() As ADODB.Recordset
On Error GoTo PROC_ERR
      Dim rs As ADODB.Recordset
      Set rs = New ADODB.Recordset
      Dim conn As ADODB.Connection
      Set conn = New ADODB.Connection
      conn.Open "DRIVER={Microsoft Text Driver (*.txt; *.csv)};" & _
                  "DBQ=" & App.Path & ";", "", ""

      rs.Open "select * from [database.txt]", conn, adOpenStatic, _
                  adLockReadOnly, adCmdText
      Set PROC_READ_TXT_FILE_DATABASE = rs
      Set rs = Nothing
      Set conn = Nothing
Exit Function
PROC_ERR:
    MsgBox err.Description
End Function

