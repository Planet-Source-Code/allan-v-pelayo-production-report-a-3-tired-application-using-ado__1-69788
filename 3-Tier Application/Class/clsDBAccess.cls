VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDBAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim Cn As ADODB.Connection
Dim objCmd As ADODB.Command
Dim mDataSource As String
Dim mProvider As String
Dim strSQL As String
Private mmc01_id_new As Long
Private mcc01_id_new As Long
Private mso01_id_dtl_new As Long
Private mot01_id_new As Long
Public Property Get DataSource() As String
    DataSource = mDataSource
End Property
Public Property Let DataSource(ByVal NewI As String)
    mDataSource = NewI
End Property
Public Function OpenConnection() As Boolean
On Error GoTo err
    If Cn.State = adStateOpen Then
        Cn.Close
    End If
    Cn.CursorLocation = adUseClient
    Cn.ConnectionString = Me.ConnectionString
    Cn.Mode = adModeReadWrite
    Cn.Open
On Error GoTo 0
    OpenConnection = True
sair:
Exit Function
err:
    OpenConnection = False
Resume sair
    err.Raise err.Number, "Database Access Class", err.Description
    Exit Function
End Function
Public Function ConnectionString()
Dim strConnection As String
On Error GoTo err
    strConnection = "Provider=" & "MSDASQL.1" & " ;Persist Security Info=False;Data Source=" & mDataSource
    ConnectionString = strConnection
Exit Function
err:
    MsgBox err.Description, vbCritical
End Function
Public Sub Class_Initialize()
    Set Cn = New ADODB.Connection
End Sub
Public Sub Class_Terminate()
    Set Cn = Nothing
End Sub
Public Function proc_usp_sel_us01_user_01(ByVal strus01_user_name As String, ByVal strus01_password As String, ByRef objRs As Recordset) As Long
'On Error GoTo PROC_ERR
    Set objCmd = New ADODB.Command
    strSQL = "usp_sel_us01_user_01"
    With objCmd
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, 0)
        .Parameters.Append .CreateParameter("@us01_user_name", adVarChar, adParamInput, 100, IIf(strus01_user_name = vbNullString, Null, strus01_user_name))
        .Parameters.Append .CreateParameter("@us01_password", adVarChar, adParamInput, 100, IIf(strus01_password = vbNullString, Null, strus01_password))
        .ActiveConnection = Cn
    End With
    With objRs
        .CursorLocation = adUseClient
        .Open objCmd, , adOpenStatic, adLockBatchOptimistic
        .ActiveConnection = Nothing
    End With
    proc_usp_sel_us01_user_01 = objCmd.Parameters("RETURN_VALUE").Value
    Set objCmd = Nothing
'Exit Function
'PROC_ERR:
'    Set objCmd = Nothing
'    proc_usp_sel_us01_user_01 = err.Number
End Function
Public Function proc_usp_insert_update_us01_user(ByVal strhr01_emplyee_id As String, ByVal strhr01_lastname As String, ByVal strhr01_firstname As String, _
                                            ByVal strhr01_middlename As String, ByVal strus01_user_name As String, ByVal strus01_password As String, _
                                            ByVal strus01_status As String) As Long
On Error GoTo PROC_ERR
    Set objCmd = New ADODB.Command
    strSQL = "usp_insert_update_us01_user"
    With objCmd
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        .ActiveConnection = Cn
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, 0)
        .Parameters.Append .CreateParameter("@hr01_emplyee_id", adVarChar, adParamInput, 100, IIf(strhr01_emplyee_id = vbNullString, Null, strhr01_emplyee_id))
        .Parameters.Append .CreateParameter("@hr01_lastname", adVarChar, adParamInput, 100, IIf(strhr01_lastname = vbNullString, Null, strhr01_lastname))
        .Parameters.Append .CreateParameter("@hr01_firstname", adVarChar, adParamInput, 100, IIf(strhr01_firstname = vbNullString, Null, strhr01_firstname))
        .Parameters.Append .CreateParameter("@hr01_middlename", adVarChar, adParamInput, 100, IIf(strhr01_middlename = vbNullString, Null, strhr01_middlename))
        .Parameters.Append .CreateParameter("@us01_user_name", adVarChar, adParamInput, 100, IIf(strus01_user_name = vbNullString, Null, strus01_user_name))
        .Parameters.Append .CreateParameter("@us01_password", adVarChar, adParamInput, 100, IIf(strus01_password = vbNullString, Null, strus01_password))
        .Parameters.Append .CreateParameter("@us01_status", adSmallInt, adParamInput, 0, IIf(strus01_status = vbNullString, Null, strus01_status))
        .ActiveConnection = Cn
        .Execute Options:=adExecuteNoRecords
    End With
    proc_usp_insert_update_us01_user = objCmd.Parameters("RETURN_VALUE")
    Set objCmd = Nothing
Exit Function
PROC_ERR:
    Set objCmd = Nothing
    proc_usp_insert_update_us01_user = err.Number
End Function
Public Function proc_usp_sel_us01_user_02(ByRef objRs As Recordset) As Long
On Error GoTo PROC_ERR
    Set objCmd = New ADODB.Command
    strSQL = "usp_sel_us01_user_02"
    With objCmd
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, 0)
        .ActiveConnection = Cn
    End With
    With objRs
        .CursorLocation = adUseClient
        .Open objCmd, , adOpenStatic, adLockBatchOptimistic
        .ActiveConnection = Nothing
    End With
    proc_usp_sel_us01_user_02 = objCmd.Parameters("RETURN_VALUE").Value
    Set objCmd = Nothing
Exit Function
PROC_ERR:
    Set objCmd = Nothing
    proc_usp_sel_us01_user_02 = err.Number
End Function
Public Function proc_usp_sel_us01_user_03(ByVal strhr01_emplyee_id As String, ByRef objRs As Recordset) As Long
On Error GoTo PROC_ERR
    Set objCmd = New ADODB.Command
    strSQL = "usp_sel_us01_user_03"
    With objCmd
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, 0)
        .Parameters.Append .CreateParameter("@hr01_emplyee_id", adVarChar, adParamInput, 100, IIf(strhr01_emplyee_id = vbNullString, Null, strhr01_emplyee_id))
        .ActiveConnection = Cn
    End With
    With objRs
        .CursorLocation = adUseClient
        .Open objCmd, , adOpenStatic, adLockBatchOptimistic
        .ActiveConnection = Nothing
    End With
    proc_usp_sel_us01_user_03 = objCmd.Parameters("RETURN_VALUE").Value
    Set objCmd = Nothing
Exit Function
PROC_ERR:
    Set objCmd = Nothing
    proc_usp_sel_us01_user_03 = err.Number
End Function
Public Function proc_usp_insert_update_op01_operator(ByVal strhr01_emplyee_id As String, ByVal strop01_lastname As String, _
                                ByVal strop01_firstname As String, ByVal strop01_middlename As String, _
                                ByVal strop01_status As String, ByVal strop01_who_created As String, _
                                ByVal strop01_who_modified As String) As Long
On Error GoTo PROC_ERR
    Set objCmd = New ADODB.Command
    strSQL = "usp_insert_update_op01_operator"
    With objCmd
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        .ActiveConnection = Cn
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, 0)
        .Parameters.Append .CreateParameter("@hr01_emplyee_id", adVarChar, adParamInput, 100, IIf(strhr01_emplyee_id = vbNullString, Null, strhr01_emplyee_id))
        .Parameters.Append .CreateParameter("@op01_lastname", adVarChar, adParamInput, 100, IIf(strop01_lastname = vbNullString, Null, strop01_lastname))
        .Parameters.Append .CreateParameter("@op01_firstname", adVarChar, adParamInput, 100, IIf(strop01_firstname = vbNullString, Null, strop01_firstname))
        .Parameters.Append .CreateParameter("@op01_middlename", adVarChar, adParamInput, 100, IIf(strop01_middlename = vbNullString, Null, strop01_middlename))
        .Parameters.Append .CreateParameter("@op01_status", adVarChar, adParamInput, 100, IIf(strop01_status = vbNullString, Null, strop01_status))
        .Parameters.Append .CreateParameter("@op01_who_created", adVarChar, adParamInput, 100, IIf(strop01_who_created = vbNullString, Null, strop01_who_created))
        .Parameters.Append .CreateParameter("@op01_who_modified", adVarChar, adParamInput, 100, IIf(strop01_who_modified = vbNullString, Null, strop01_who_modified))
        .ActiveConnection = Cn
        .Execute Options:=adExecuteNoRecords
    End With
    proc_usp_insert_update_op01_operator = objCmd.Parameters("RETURN_VALUE")
    Set objCmd = Nothing
Exit Function
PROC_ERR:
    Set objCmd = Nothing
    proc_usp_insert_update_op01_operator = err.Number
End Function
Public Function proc_usp_sel_op01_operator_01(ByVal strhr01_emplyee_id As String, ByRef objRs As Recordset) As Long
On Error GoTo PROC_ERR
    Set objCmd = New ADODB.Command
    strSQL = "usp_sel_op01_operator_01"
    With objCmd
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, 0)
        .Parameters.Append .CreateParameter("@hr01_emplyee_id", adVarChar, adParamInput, 100, IIf(strhr01_emplyee_id = vbNullString, Null, strhr01_emplyee_id))
        .ActiveConnection = Cn
    End With
    With objRs
        .CursorLocation = adUseClient
        .Open objCmd, , adOpenStatic, adLockBatchOptimistic
        .ActiveConnection = Nothing
    End With
    proc_usp_sel_op01_operator_01 = objCmd.Parameters("RETURN_VALUE").Value
    Set objCmd = Nothing
Exit Function
PROC_ERR:
    Set objCmd = Nothing
    proc_usp_sel_op01_operator_01 = err.Number
End Function
Public Function proc_usp_sel_op01_operator_02(ByRef objRs As Recordset) As Long
On Error GoTo PROC_ERR
    Set objCmd = New ADODB.Command
    strSQL = "usp_sel_op01_operator_02"
    With objCmd
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, 0)
        .ActiveConnection = Cn
    End With
    With objRs
        .CursorLocation = adUseClient
        .Open objCmd, , adOpenStatic, adLockBatchOptimistic
        .ActiveConnection = Nothing
    End With
    proc_usp_sel_op01_operator_02 = objCmd.Parameters("RETURN_VALUE").Value
    Set objCmd = Nothing
Exit Function
PROC_ERR:
    Set objCmd = Nothing
    proc_usp_sel_op01_operator_02 = err.Number
End Function
Public Function proc_usp_insert_update_mc01_machinery(ByVal strmc01_id As String, ByVal strmc01_machine As String, _
                    ByVal strmc01_machine_desc As String, ByVal strmc01_serial_no As String, _
                    ByVal strmc01_machine_model As String, ByVal strmc01_date_acquired As String, _
                    ByVal strmc01_acquisition_cost As String, ByVal strmc01_who_created As String, _
                    ByVal strmc01_who_modified As String) As Long
On Error GoTo PROC_ERR
    Set objCmd = New ADODB.Command
    strSQL = "usp_insert_update_mc01_machinery"
    With objCmd
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        .ActiveConnection = Cn
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, 0)
        .Parameters.Append .CreateParameter("@mc01_id", adSmallInt, adParamInput, 0, IIf(strmc01_id = vbNullString, Null, strmc01_id))
        .Parameters.Append .CreateParameter("@mc01_machine", adVarChar, adParamInput, 100, IIf(strmc01_machine = vbNullString, Null, strmc01_machine))
        .Parameters.Append .CreateParameter("@mc01_machine_desc", adVarChar, adParamInput, 255, IIf(strmc01_machine_desc = vbNullString, Null, strmc01_machine_desc))
        .Parameters.Append .CreateParameter("@mc01_serial_no", adVarChar, adParamInput, 50, IIf(strmc01_serial_no = vbNullString, Null, strmc01_serial_no))
        .Parameters.Append .CreateParameter("@mc01_machine_model", adVarChar, adParamInput, 50, IIf(strmc01_machine_model = vbNullString, Null, strmc01_machine_model))
        .Parameters.Append .CreateParameter("@mc01_date_acquired", adDBTimeStamp, adParamInput, 0, IIf(strmc01_date_acquired = vbNullString, Null, strmc01_date_acquired))
        .Parameters.Append .CreateParameter("@mc01_acquisition_cost", adCurrency, adParamInput, 8, IIf(strmc01_acquisition_cost = vbNullString, Null, strmc01_acquisition_cost))
        .Parameters.Append .CreateParameter("@mc01_who_created", adVarChar, adParamInput, 100, IIf(strmc01_who_created = vbNullString, Null, strmc01_who_created))
        .Parameters.Append .CreateParameter("@mc01_who_modified", adVarChar, adParamInput, 100, IIf(strmc01_who_modified = vbNullString, Null, strmc01_who_modified))
        .Parameters.Append .CreateParameter("@mc01_id_new", adInteger, adParamOutput)
        .ActiveConnection = Cn
        .Execute Options:=adExecuteNoRecords
        mmc01_id_new = .Parameters("@mc01_id_new")
    End With
    proc_usp_insert_update_mc01_machinery = objCmd.Parameters("RETURN_VALUE")
    Set objCmd = Nothing
Exit Function
PROC_ERR:
    Set objCmd = Nothing
    proc_usp_insert_update_mc01_machinery = err.Number
End Function
Public Property Let lmc01_id_new(ByVal vNewValue As Long)
    mmc01_id_new = vNewValue
End Property
Public Property Get lmc01_id_new() As Long
   lmc01_id_new = mmc01_id_new
End Property
Public Function proc_usp_sel_mc01_machinery_02(ByRef objRs As Recordset) As Long
On Error GoTo PROC_ERR
    Set objCmd = New ADODB.Command
    strSQL = "usp_sel_mc01_machinery_02"
    With objCmd
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, 0)
        .ActiveConnection = Cn
    End With
    With objRs
        .CursorLocation = adUseClient
        .Open objCmd, , adOpenStatic, adLockBatchOptimistic
        .ActiveConnection = Nothing
    End With
    proc_usp_sel_mc01_machinery_02 = objCmd.Parameters("RETURN_VALUE").Value
    Set objCmd = Nothing
Exit Function
PROC_ERR:
    Set objCmd = Nothing
    proc_usp_sel_mc01_machinery_02 = err.Number
End Function
Public Function proc_usp_sel_mc01_machinery_01(ByVal strmc01_id As String, ByRef objRs As Recordset) As Long
On Error GoTo PROC_ERR
    Set objCmd = New ADODB.Command
    strSQL = "usp_sel_mc01_machinery_01"
    With objCmd
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, 0)
        .Parameters.Append .CreateParameter("@mc01_id", adVarChar, adParamInput, 100, IIf(strmc01_id = vbNullString, Null, strmc01_id))
        .ActiveConnection = Cn
    End With
    With objRs
        .CursorLocation = adUseClient
        .Open objCmd, , adOpenStatic, adLockBatchOptimistic
        .ActiveConnection = Nothing
    End With
    proc_usp_sel_mc01_machinery_01 = objCmd.Parameters("RETURN_VALUE").Value
    Set objCmd = Nothing
Exit Function
PROC_ERR:
    Set objCmd = Nothing
    proc_usp_sel_mc01_machinery_01 = err.Number
End Function
Public Function proc_usp_sel_cc01_capacity_01(ByVal strcc01_id As String, ByRef objRs As Recordset) As Long
On Error GoTo PROC_ERR
    Set objCmd = New ADODB.Command
    strSQL = "usp_sel_cc01_capacity_01"
    With objCmd
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, 0)
        .Parameters.Append .CreateParameter("@cc01_id", adInteger, adParamInput, 0, IIf(strcc01_id = vbNullString, Null, strcc01_id))
        .ActiveConnection = Cn
    End With
    With objRs
        .CursorLocation = adUseClient
        .Open objCmd, , adOpenStatic, adLockBatchOptimistic
        .ActiveConnection = Nothing
    End With
    proc_usp_sel_cc01_capacity_01 = objCmd.Parameters("RETURN_VALUE").Value
    Set objCmd = Nothing
Exit Function
PROC_ERR:
    Set objCmd = Nothing
    proc_usp_sel_cc01_capacity_01 = err.Number
End Function
Public Function proc_usp_sel_cc01_capacity_02(ByRef objRs As Recordset) As Long
On Error GoTo PROC_ERR
    Set objCmd = New ADODB.Command
    strSQL = "usp_sel_cc01_capacity_02"
    With objCmd
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, 0)
        .ActiveConnection = Cn
    End With
    With objRs
        .CursorLocation = adUseClient
        .Open objCmd, , adOpenStatic, adLockBatchOptimistic
        .ActiveConnection = Nothing
    End With
    proc_usp_sel_cc01_capacity_02 = objCmd.Parameters("RETURN_VALUE").Value
    Set objCmd = Nothing
Exit Function
PROC_ERR:
    Set objCmd = Nothing
    proc_usp_sel_cc01_capacity_02 = err.Number
End Function
Public Function proc_usp_insert_update_cc01_capacity(ByVal strcc01_id As String, ByVal strcc01_capacity As String, ByVal strcc01_desc As String, _
                                ByVal strcc01_wastage As String, ByVal strcc01_who_created As String, _
                                ByVal strcc01_who_modified As String) As Long
On Error GoTo PROC_ERR
    Set objCmd = New ADODB.Command
    strSQL = "usp_insert_update_cc01_capacity"
    With objCmd
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        .ActiveConnection = Cn
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, 0)
        .Parameters.Append .CreateParameter("@cc01_id", adSmallInt, adParamInput, 0, IIf(strcc01_id = vbNullString, Null, strcc01_id))
        .Parameters.Append .CreateParameter("@cc01_capacity", adVarChar, adParamInput, 100, IIf(strcc01_capacity = vbNullString, Null, strcc01_capacity))
        .Parameters.Append .CreateParameter("@cc01_desc", adVarChar, adParamInput, 100, IIf(strcc01_desc = vbNullString, Null, strcc01_desc))
        .Parameters.Append .CreateParameter("@cc01_wastage", adVarChar, adParamInput, 50, IIf(strcc01_wastage = vbNullString, Null, strcc01_wastage))
        .Parameters.Append .CreateParameter("@cc01_who_created", adVarChar, adParamInput, 100, IIf(strcc01_who_created = vbNullString, Null, strcc01_who_created))
        .Parameters.Append .CreateParameter("@cc01_who_modified", adVarChar, adParamInput, 100, IIf(strcc01_who_modified = vbNullString, Null, strcc01_who_modified))
        .Parameters.Append .CreateParameter("@cc01_id_new", adInteger, adParamOutput)
        .ActiveConnection = Cn
        .Execute Options:=adExecuteNoRecords
        mcc01_id_new = .Parameters("@cc01_id_new")
    End With
    proc_usp_insert_update_cc01_capacity = objCmd.Parameters("RETURN_VALUE")
    Set objCmd = Nothing
Exit Function
PROC_ERR:
    Set objCmd = Nothing
    proc_usp_insert_update_cc01_capacity = err.Number
End Function
Public Property Let lcc01_id_new(ByVal vNewValue As Long)
    mcc01_id_new = vNewValue
End Property
Public Property Get lcc01_id_new() As Long
   lcc01_id_new = mcc01_id_new
End Property

Public Function proc_usp_sel_cc01_capacity_03(ByVal strcc01_capacity As String, ByRef objRs As Recordset) As Long
On Error GoTo PROC_ERR
    Set objCmd = New ADODB.Command
    strSQL = "usp_sel_cc01_capacity_03"
    With objCmd
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, 0)
        .Parameters.Append .CreateParameter("@cc01_capacity", adVarChar, adParamInput, 100, IIf(strcc01_capacity = vbNullString, Null, strcc01_capacity))
        .ActiveConnection = Cn
    End With
    With objRs
        .CursorLocation = adUseClient
        .Open objCmd, , adOpenStatic, adLockBatchOptimistic
        .ActiveConnection = Nothing
    End With
    proc_usp_sel_cc01_capacity_03 = objCmd.Parameters("RETURN_VALUE").Value
    Set objCmd = Nothing
Exit Function
PROC_ERR:
     Set objCmd = Nothing
    proc_usp_sel_cc01_capacity_03 = err.Number
End Function
Public Function proc_usp_insert_update_mc01_machine_capacity_hdr(ByVal strmc01_machine As String, ByVal strmc01_who_created As String, ByVal strmc01_who_modified As String, ByVal strmc01_machine_capacity_dtl) As Long
On Error GoTo PROC_ERR
    Set objCmd = New ADODB.Command
    strSQL = "usp_insert_update_mc01_machine_capacity_hdr"
    With objCmd
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        .ActiveConnection = Cn
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, 0)
        .Parameters.Append .CreateParameter("@mc01_machine", adVarChar, adParamInput, 100, IIf(strmc01_machine = vbNullString, Null, strmc01_machine))
        .Parameters.Append .CreateParameter("@mc01_who_created", adVarChar, adParamInput, 100, IIf(strmc01_who_created = vbNullString, Null, strmc01_who_created))
        .Parameters.Append .CreateParameter("@mc01_who_modified", adVarChar, adParamInput, 100, IIf(strmc01_who_modified = vbNullString, Null, strmc01_who_modified))
        .Parameters.Append .CreateParameter("@mc01_machine_capacity_dtl", adVarChar, adParamInput, Len(strmc01_machine_capacity_dtl), IIf(strmc01_machine_capacity_dtl = vbNullString, Null, strmc01_machine_capacity_dtl))
        .ActiveConnection = Cn
        .Execute Options:=adExecuteNoRecords
    End With
    proc_usp_insert_update_mc01_machine_capacity_hdr = objCmd.Parameters("RETURN_VALUE")
    Set objCmd = Nothing
Exit Function
PROC_ERR:
    Set objCmd = Nothing
    proc_usp_insert_update_mc01_machine_capacity_hdr = err.Number
End Function

Public Function proc_usp_sel_mc01_machine_capacity_hdr_01(ByVal strmc01_machine As String, ByRef objRs As Recordset) As Long
On Error GoTo PROC_ERR
    Set objCmd = New ADODB.Command
    strSQL = "usp_sel_mc01_machine_capacity_hdr_01"
    With objCmd
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, 0)
        .Parameters.Append .CreateParameter("@mc01_machine", adVarChar, adParamInput, 100, IIf(strmc01_machine = vbNullString, Null, strmc01_machine))
        .ActiveConnection = Cn
    End With
    With objRs
        .CursorLocation = adUseClient
        .Open objCmd, , adOpenStatic, adLockBatchOptimistic
        .ActiveConnection = Nothing
    End With
    proc_usp_sel_mc01_machine_capacity_hdr_01 = objCmd.Parameters("RETURN_VALUE").Value
    Set objCmd = Nothing
Exit Function
PROC_ERR:
     Set objCmd = Nothing
    proc_usp_sel_mc01_machine_capacity_hdr_01 = err.Number
End Function
Public Function proc_usp_insert_update_ot01_output_operator(ByVal strot01_id As String, ByVal strhr01_emplyee_id As String, _
                        ByVal strmc01_machine As String, ByVal strcc01_capacity As String, _
                        ByVal strot01_date_product As String, ByVal strot01_set_time As String, _
                        ByVal strot01_quantity_produce As String, ByVal strot01_running_time As String, _
                        ByVal strot01_who_created As String, ByVal strot01_who_modified As String) As Long
On Error GoTo PROC_ERR
    Set objCmd = New ADODB.Command
    strSQL = "usp_insert_update_ot01_output_operator"
    With objCmd
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        .ActiveConnection = Cn
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, 0)
        .Parameters.Append .CreateParameter("@ot01_id", adInteger, adParamInput, 0, IIf(strot01_id = vbNullString, Null, strot01_id))
        .Parameters.Append .CreateParameter("@hr01_emplyee_id", adVarChar, adParamInput, 100, IIf(strhr01_emplyee_id = vbNullString, Null, strhr01_emplyee_id))
        .Parameters.Append .CreateParameter("@mc01_machine", adVarChar, adParamInput, 100, IIf(strmc01_machine = vbNullString, Null, strmc01_machine))
        .Parameters.Append .CreateParameter("@cc01_capacity", adVarChar, adParamInput, 100, IIf(strcc01_capacity = vbNullString, Null, strcc01_capacity))
        .Parameters.Append .CreateParameter("@ot01_date_product", adDBTimeStamp, adParamInput, 0, IIf(strot01_date_product = vbNullString, Null, strot01_date_product))
        .Parameters.Append .CreateParameter("@ot01_set_time", adDouble, adParamInput, 8, IIf(strot01_set_time = vbNullString, Null, strot01_set_time))
        .Parameters.Append .CreateParameter("@ot01_quantity_produce", adDouble, adParamInput, 8, IIf(strot01_quantity_produce = vbNullString, Null, strot01_quantity_produce))
        .Parameters.Append .CreateParameter("@ot01_running_time", adDouble, adParamInput, 8, IIf(strot01_running_time = vbNullString, Null, strot01_running_time))
        .Parameters.Append .CreateParameter("@ot01_who_created", adVarChar, adParamInput, 100, IIf(strot01_who_created = vbNullString, Null, strot01_who_created))
        .Parameters.Append .CreateParameter("@ot01_who_modified", adVarChar, adParamInput, 100, IIf(strot01_who_modified = vbNullString, Null, strot01_who_modified))
        .Parameters.Append .CreateParameter("@ot01_id_new", adInteger, adParamOutput)
        .ActiveConnection = Cn
        .Execute Options:=adExecuteNoRecords
         mot01_id_new = .Parameters("@ot01_id_new")
    End With
    proc_usp_insert_update_ot01_output_operator = objCmd.Parameters("RETURN_VALUE")
    Set objCmd = Nothing
Exit Function
PROC_ERR:
    Set objCmd = Nothing
    proc_usp_insert_update_ot01_output_operator = err.Number
End Function
Public Function proc_usp_sel_op01_operator_03(ByVal strhr01_emplyee_id As String, ByRef objRs As Recordset) As Long
On Error GoTo PROC_ERR
    Set objCmd = New ADODB.Command
    strSQL = "usp_sel_op01_operator_03"
    With objCmd
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, 0)
        .Parameters.Append .CreateParameter("@hr01_emplyee_id", adVarChar, adParamInput, 100, IIf(strhr01_emplyee_id = vbNullString, Null, strhr01_emplyee_id))
        .ActiveConnection = Cn
    End With
    With objRs
        .CursorLocation = adUseClient
        .Open objCmd, , adOpenStatic, adLockBatchOptimistic
        .ActiveConnection = Nothing
    End With
    proc_usp_sel_op01_operator_03 = objCmd.Parameters("RETURN_VALUE").Value
    Set objCmd = Nothing
Exit Function
PROC_ERR:
    Set objCmd = Nothing
    proc_usp_sel_op01_operator_03 = err.Number
End Function
Public Function proc_usp_sel_ot01_output_operator_01(ByVal strhr01_emplyee_id As String, ByVal strmc01_machine As String, _
            ByVal strot01_date_product_low As String, ByVal strot01_date_product_high As String, ByRef objRs As Recordset) As Long
On Error GoTo PROC_ERR
    Set objCmd = New ADODB.Command
    strSQL = "usp_sel_ot01_output_operator_01"
    With objCmd
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, 0)
        .Parameters.Append .CreateParameter("@hr01_emplyee_id", adVarChar, adParamInput, 100, IIf(strhr01_emplyee_id = vbNullString, Null, strhr01_emplyee_id))
        .Parameters.Append .CreateParameter("@mc01_machine", adVarChar, adParamInput, 100, IIf(strmc01_machine = vbNullString, Null, strmc01_machine))
        .Parameters.Append .CreateParameter("@ot01_date_product_low", adDBTimeStamp, adParamInput, 0, IIf(strot01_date_product_low = vbNullString, Null, strot01_date_product_low))
        .Parameters.Append .CreateParameter("@ot01_date_product_high", adDBTimeStamp, adParamInput, 0, IIf(strot01_date_product_high = vbNullString, Null, strot01_date_product_high))
        .ActiveConnection = Cn
    End With
    With objRs
        .CursorLocation = adUseClient
        .Open objCmd, , adOpenStatic, adLockBatchOptimistic
        .ActiveConnection = Nothing
    End With
    proc_usp_sel_ot01_output_operator_01 = objCmd.Parameters("RETURN_VALUE").Value
    Set objCmd = Nothing
Exit Function
PROC_ERR:
    Set objCmd = Nothing
    proc_usp_sel_ot01_output_operator_01 = err.Number
End Function
Public Function proc_usp_sel_ot01_output_operator_02(ByVal strhr01_emplyee_id As String, ByVal strmc01_machine As String, _
            ByVal strot01_id As String, ByRef objRs As Recordset) As Long
On Error GoTo PROC_ERR
    Set objCmd = New ADODB.Command
    strSQL = "usp_sel_ot01_output_operator_02"
    With objCmd
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, 0)
        .Parameters.Append .CreateParameter("@hr01_emplyee_id", adVarChar, adParamInput, 100, IIf(strhr01_emplyee_id = vbNullString, Null, strhr01_emplyee_id))
        .Parameters.Append .CreateParameter("@mc01_machine", adVarChar, adParamInput, 100, IIf(strmc01_machine = vbNullString, Null, strmc01_machine))
        .Parameters.Append .CreateParameter("@ot01_id", adInteger, adParamInput, 0, IIf(strot01_id = vbNullString, Null, strot01_id))
        .ActiveConnection = Cn
    End With
    With objRs
        .CursorLocation = adUseClient
        .Open objCmd, , adOpenStatic, adLockBatchOptimistic
        .ActiveConnection = Nothing
    End With
    proc_usp_sel_ot01_output_operator_02 = objCmd.Parameters("RETURN_VALUE").Value
    Set objCmd = Nothing
Exit Function
PROC_ERR:
    Set objCmd = Nothing
    proc_usp_sel_ot01_output_operator_02 = err.Number
End Function
Public Function proc_usp_sel_op01_operator_04(ByVal strsearchcriteria As String, ByVal strsearchvalue As String, _
            ByRef objRs As Recordset) As Long
On Error GoTo PROC_ERR
    Set objCmd = New ADODB.Command
    strSQL = "usp_sel_op01_operator_04"
    With objCmd
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, 0)
        .Parameters.Append .CreateParameter("@searchcriteria", adVarChar, adParamInput, 100, IIf(strsearchcriteria = vbNullString, Null, strsearchcriteria))
        .Parameters.Append .CreateParameter("@searchvalue", adVarChar, adParamInput, 100, IIf(strsearchvalue = vbNullString, Null, strsearchvalue))
        .ActiveConnection = Cn
    End With
    With objRs
        .CursorLocation = adUseClient
        .Open objCmd, , adOpenStatic, adLockBatchOptimistic
        .ActiveConnection = Nothing
    End With
    proc_usp_sel_op01_operator_04 = objCmd.Parameters("RETURN_VALUE").Value
    Set objCmd = Nothing
Exit Function
PROC_ERR:
    Set objCmd = Nothing
    proc_usp_sel_op01_operator_04 = err.Number
End Function
Public Function proc_usp_sel_ot01_output_operator_03(ByVal strot01_date_product_low As String, ByVal strot01_date_product_high As String, ByRef objRs As Recordset) As Long
On Error GoTo PROC_ERR
    Set objCmd = New ADODB.Command
    strSQL = "usp_sel_ot01_output_operator_03"
    With objCmd
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, 0)
        .Parameters.Append .CreateParameter("@ot01_date_product_low", adDBTimeStamp, adParamInput, 0, IIf(strot01_date_product_low = vbNullString, Null, strot01_date_product_low))
        .Parameters.Append .CreateParameter("@ot01_date_product_high", adDBTimeStamp, adParamInput, 0, IIf(strot01_date_product_high = vbNullString, Null, strot01_date_product_high))
        .ActiveConnection = Cn
    End With
    With objRs
        .CursorLocation = adUseClient
        .Open objCmd, , adOpenStatic, adLockBatchOptimistic
        .ActiveConnection = Nothing
    End With
    proc_usp_sel_ot01_output_operator_03 = objCmd.Parameters("RETURN_VALUE").Value
    Set objCmd = Nothing
Exit Function
PROC_ERR:
    Set objCmd = Nothing
    proc_usp_sel_ot01_output_operator_03 = err.Number
End Function
Public Function proc_usp_insert_update_so01_so_dtl(ByVal strso01_id_dtl As String, ByVal strso01_ref As String, _
                                                    ByVal strjo01_ref As String, ByVal strso01_shift As String, _
                                                    ByVal strhr01_emplyee_id As String, ByVal strmc01_machine As String, _
                                                    ByVal strcc01_capacity As String, ByVal strso01_date_product As String, _
                                                    ByVal strso01_set_time As String, ByVal strso01_quantity_produce As String, _
                                                    ByVal strso01_running_time As String, ByVal strso01_downtime As String, _
                                                    ByVal strmc01_reason_downtime As String, ByVal strso01_who_created As String, ByVal strso01_who_modified As String) As Long
On Error GoTo PROC_ERR
    Set objCmd = New ADODB.Command
    strSQL = "usp_insert_update_so01_so_dtl"
    With objCmd
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, 0)
        .Parameters.Append .CreateParameter("@so01_id_dtl", adInteger, adParamInput, 0, IIf(strso01_id_dtl = vbNullString, Null, strso01_id_dtl))
        .Parameters.Append .CreateParameter("@so01_ref", adVarChar, adParamInput, 50, IIf(strso01_ref = vbNullString, Null, strso01_ref))
        .Parameters.Append .CreateParameter("@jo01_ref", adVarChar, adParamInput, 50, IIf(strjo01_ref = vbNullString, Null, strjo01_ref))
        .Parameters.Append .CreateParameter("@so01_shift", adVarChar, adParamInput, 50, IIf(strso01_shift = vbNullString, Null, strso01_shift))
        .Parameters.Append .CreateParameter("@hr01_emplyee_id", adVarChar, adParamInput, 100, IIf(strhr01_emplyee_id = vbNullString, Null, strhr01_emplyee_id))
        .Parameters.Append .CreateParameter("@mc01_machine", adVarChar, adParamInput, 100, IIf(strmc01_machine = vbNullString, Null, strmc01_machine))
        .Parameters.Append .CreateParameter("@cc01_capacity", adVarChar, adParamInput, 100, IIf(strcc01_capacity = vbNullString, Null, strcc01_capacity))
        .Parameters.Append .CreateParameter("@so01_date_product", adDBTimeStamp, adParamInput, 0, IIf(strso01_date_product = vbNullString, Null, strso01_date_product))
        .Parameters.Append .CreateParameter("@so01_set_time", adDouble, adParamInput, 8, IIf(strso01_set_time = vbNullString, Null, strso01_set_time))
        .Parameters.Append .CreateParameter("@so01_quantity_produce", adDouble, adParamInput, 8, IIf(strso01_quantity_produce = vbNullString, Null, strso01_quantity_produce))
        .Parameters.Append .CreateParameter("@so01_running_time", adDouble, adParamInput, 8, IIf(strso01_running_time = vbNullString, Null, strso01_running_time))
        .Parameters.Append .CreateParameter("@so01_downtime", adDouble, adParamInput, 8, IIf(strso01_downtime = vbNullString, Null, strso01_downtime))
        .Parameters.Append .CreateParameter("@mc01_reason_downtime", adVarChar, adParamInput, 300, IIf(strmc01_reason_downtime = vbNullString, Null, strmc01_reason_downtime))
        .Parameters.Append .CreateParameter("@so01_who_created", adVarChar, adParamInput, 100, IIf(strso01_who_created = vbNullString, Null, strso01_who_created))
        .Parameters.Append .CreateParameter("@so01_who_modified", adVarChar, adParamInput, 100, IIf(strso01_who_modified = vbNullString, Null, strso01_who_modified))
        .Parameters.Append .CreateParameter("@so01_id_dtl_new", adInteger, adParamOutput, 0)
        .ActiveConnection = Cn
        .Execute Options:=adExecuteNoRecords
         mso01_id_dtl_new = .Parameters("@so01_id_dtl_new")
    End With
    proc_usp_insert_update_so01_so_dtl = objCmd.Parameters("RETURN_VALUE").Value
    Set objCmd = Nothing
Exit Function
PROC_ERR:
    Set objCmd = Nothing
    proc_usp_insert_update_so01_so_dtl = err.Number
End Function
Public Property Let lso01_id_dtl_new(ByVal vNewValue As Long)
    mso01_id_dtl_new = vNewValue
End Property
Public Property Get lso01_id_dtl_new() As Long
   lso01_id_dtl_new = mso01_id_dtl_new
End Property
Public Function proc_usp_sel_so01_so_dtl_01(ByRef objRs As Recordset) As Long
On Error GoTo PROC_ERR
    Set objCmd = New ADODB.Command
    strSQL = "usp_sel_so01_so_dtl_01"
    With objCmd
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, 0)
        .ActiveConnection = Cn
    End With
    With objRs
        .CursorLocation = adUseClient
        .Open objCmd, , adOpenStatic, adLockBatchOptimistic
        .ActiveConnection = Nothing
    End With
    proc_usp_sel_so01_so_dtl_01 = objCmd.Parameters("RETURN_VALUE").Value
    Set objCmd = Nothing
Exit Function
PROC_ERR:
    Set objCmd = Nothing
    proc_usp_sel_so01_so_dtl_01 = err.Number
End Function
Public Function proc_usp_sel_so01_so_dtl_02(ByVal strso01_id_dtl As String, ByVal strso01_ref As String, ByRef objRs As Recordset) As Long
On Error GoTo PROC_ERR
    Set objCmd = New ADODB.Command
    strSQL = "usp_sel_so01_so_dtl_02"
    With objCmd
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, 0)
        .Parameters.Append .CreateParameter("@so01_id_dtl", adVarChar, adParamInput, 100, IIf(strso01_id_dtl = vbNullString, Null, strso01_id_dtl))
        .Parameters.Append .CreateParameter("@so01_ref", adVarChar, adParamInput, 100, IIf(strso01_ref = vbNullString, Null, strso01_ref))
        .ActiveConnection = Cn
    End With
    With objRs
        .CursorLocation = adUseClient
        .Open objCmd, , adOpenStatic, adLockBatchOptimistic
        .ActiveConnection = Nothing
    End With
    proc_usp_sel_so01_so_dtl_02 = objCmd.Parameters("RETURN_VALUE").Value
    Set objCmd = Nothing
Exit Function
PROC_ERR:
    Set objCmd = Nothing
    proc_usp_sel_so01_so_dtl_02 = err.Number
End Function
Public Function proc_usp_sel_so01_so_dtl_03(ByVal strfilter As String, ByVal strso01_date_product_low As String, _
                                            ByVal strso01_date_product_high As String, ByVal strmc01_machine As String, _
                                            ByRef objRs As Recordset) As Long
On Error GoTo PROC_ERR
    Set objCmd = New ADODB.Command
    strSQL = "usp_sel_so01_so_dtl_03"
    With objCmd
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, 0)
        .Parameters.Append .CreateParameter("@filter", adInteger, adParamInput, 0, IIf(strfilter = vbNullString, Null, strfilter))
        .Parameters.Append .CreateParameter("@so01_date_product_low", adDBTimeStamp, adParamInput, 0, IIf(strso01_date_product_low = vbNullString, Null, strso01_date_product_low))
        .Parameters.Append .CreateParameter("@so01_date_product_high", adDBTimeStamp, adParamInput, 0, IIf(strso01_date_product_high = vbNullString, Null, strso01_date_product_high))
        .Parameters.Append .CreateParameter("@mc01_machine", adVarChar, adParamInput, 100, IIf(strmc01_machine = vbNullString, Null, strmc01_machine))
        .ActiveConnection = Cn
    End With
    With objRs
        .CursorLocation = adUseClient
        .Open objCmd, , adOpenStatic, adLockBatchOptimistic
        .ActiveConnection = Nothing
    End With
    proc_usp_sel_so01_so_dtl_03 = objCmd.Parameters("RETURN_VALUE").Value
    Set objCmd = Nothing
Exit Function
PROC_ERR:
    Set objCmd = Nothing
    proc_usp_sel_so01_so_dtl_03 = err.Number
End Function
Public Function proc_usp_sel_so01_so_dtl_04(ByVal strso01_date_product_low As String, _
                                            ByVal strso01_date_product_high As String, ByVal strmc01_machine As String, _
                                            ByVal strhr01_emplyee_id As String, ByRef objRs As Recordset) As Long
On Error GoTo PROC_ERR
    Set objCmd = New ADODB.Command
    strSQL = "usp_sel_so01_so_dtl_04"
    With objCmd
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, 0)
        .Parameters.Append .CreateParameter("@so01_date_product_low", adDBTimeStamp, adParamInput, 0, IIf(strso01_date_product_low = vbNullString, Null, strso01_date_product_low))
        .Parameters.Append .CreateParameter("@so01_date_product_high", adDBTimeStamp, adParamInput, 0, IIf(strso01_date_product_high = vbNullString, Null, strso01_date_product_high))
        .Parameters.Append .CreateParameter("@mc01_machine", adVarChar, adParamInput, 100, IIf(strmc01_machine = vbNullString, Null, strmc01_machine))
        .Parameters.Append .CreateParameter("@hr01_emplyee_id", adVarChar, adParamInput, 100, IIf(strhr01_emplyee_id = vbNullString, Null, strhr01_emplyee_id))
        .ActiveConnection = Cn
    End With
    With objRs
        .CursorLocation = adUseClient
        .Open objCmd, , adOpenStatic, adLockBatchOptimistic
        .ActiveConnection = Nothing
    End With
    proc_usp_sel_so01_so_dtl_04 = objCmd.Parameters("RETURN_VALUE").Value
    Set objCmd = Nothing
Exit Function
PROC_ERR:
    Set objCmd = Nothing
    proc_usp_sel_so01_so_dtl_04 = err.Number
End Function
Public Function proc_usp_sel_sh01_shift_01(ByRef objRs As Recordset) As Long
On Error GoTo PROC_ERR
    Set objCmd = New ADODB.Command
    strSQL = "usp_sel_sh01_shift_01"
    With objCmd
        .CommandText = strSQL
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, 0)
        .ActiveConnection = Cn
    End With
    With objRs
        .CursorLocation = adUseClient
        .Open objCmd, , adOpenStatic, adLockBatchOptimistic
        .ActiveConnection = Nothing
    End With
    proc_usp_sel_sh01_shift_01 = objCmd.Parameters("RETURN_VALUE").Value
    Set objCmd = Nothing
Exit Function
PROC_ERR:
    Set objCmd = Nothing
    proc_usp_sel_sh01_shift_01 = err.Number
End Function

