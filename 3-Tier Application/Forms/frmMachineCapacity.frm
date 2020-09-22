VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Object = "{3F666E72-7F79-447A-BCFF-C42C44E3DBE7}#1.0#0"; "iGrid251_75B4A91C.ocx"
Begin VB.Form frmMachineCapacity 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Machine Capacity"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8280
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMachineCapacity.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   8280
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Capacity"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   8055
      Begin iGrid251_75B4A91C.iGrid grdMachineCapacity 
         Height          =   4095
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   7223
      End
      Begin LVbuttons.LaVolpeButton cmdAddRow 
         Height          =   410
         Left            =   120
         TabIndex        =   1
         Top             =   4920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   714
         BTYPE           =   6
         TX              =   "&Add Row"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   14215660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmMachineCapacity.frx":127A
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton cmdDeleteRow 
         Height          =   410
         Left            =   1800
         TabIndex        =   2
         Top             =   4920
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   714
         BTYPE           =   6
         TX              =   "&Delete Row"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   14215660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmMachineCapacity.frx":1296
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin VB.Line Line6 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         X1              =   0
         X2              =   8040
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Line Line5 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         X1              =   8040
         X2              =   8040
         Y1              =   120
         Y2              =   5400
      End
      Begin VB.Label lblCapacity 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   4480
         Width           =   7815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Machine"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   8055
      Begin VB.ComboBox cboMcCapacity 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   2655
      End
      Begin VB.Line Line4 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         X1              =   0
         X2              =   8040
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         X1              =   8040
         X2              =   8040
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Machine"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   7
         Top             =   405
         Width           =   855
      End
   End
   Begin LVbuttons.LaVolpeButton cmdSave 
      Height          =   410
      Left            =   120
      TabIndex        =   3
      Top             =   6600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   714
      BTYPE           =   6
      TX              =   "&Save"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmMachineCapacity.frx":12B2
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton cmdClose 
      Height          =   410
      Left            =   1800
      TabIndex        =   4
      Top             =   6600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   714
      BTYPE           =   6
      TX              =   "&Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmMachineCapacity.frx":12CE
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000E&
      X1              =   120
      X2              =   8160
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      X1              =   120
      X2              =   8160
      Y1              =   6480
      Y2              =   6480
   End
End
Attribute VB_Name = "frmMachineCapacity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cdbcn As clsDBAccess
'Dim cdbcn As cDataAccess.clsDBAccess
Private m_sDecSep As String ' the current system decimal separator char
Private Sub exec_proc_usp_sel_mc01_machinery_02()
Dim Rsmc01_machinery_02 As ADODB.Recordset
On Error GoTo ErrHandler
    Set cdbcn = New clsDBAccess
    'Set cdbcn = New cDataAccess.clsDBAccess
    cdbcn.DataSource = strDatabase
    If Not cdbcn.OpenConnection = True Then GoTo ErrHandler
    Set Rsmc01_machinery_02 = New ADODB.Recordset
        If cdbcn.proc_usp_sel_mc01_machinery_02(Rsmc01_machinery_02) = 0 Then
            If Rsmc01_machinery_02.RecordCount > 0 Then Rsmc01_machinery_02.MoveFirst
                Do Until Rsmc01_machinery_02.EOF
                If Not IsNull(Rsmc01_machinery_02!mc01_machine) Then cboMcCapacity.AddItem Rsmc01_machinery_02!mc01_machine
                    Rsmc01_machinery_02.MoveNext
                DoEvents
                Loop
        Else
            Set Rsmc01_machinery_02 = Nothing
            Set cdbcn = Nothing
            MsgBox "There is an error executing the command.", vbCritical, App.Title
            Exit Sub
        End If
    Set Rsmc01_machinery_02 = Nothing
    Set cdbcn = Nothing
Exit Sub
ErrHandler:
    Set Rsmc01_machinery_02 = Nothing
    Set cdbcn = Nothing
    If err.Number <> 0 Then prompt_errlog err, frmMachineCapacity.Name, "proc_usp_sel_mc01_machinery_02"
End Sub
Private Sub cboMcCapacity_Change()
    exec_proc_usp_sel_mc01_machine_capacity_hdr_01
End Sub

Private Sub cboMcCapacity_Click()
    exec_proc_usp_sel_mc01_machine_capacity_hdr_01
End Sub

Private Sub cboMcCapacity_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub cmdAddRow_Click()
Dim irow As Long
    With grdMachineCapacity
        For irow = 1 To .RowCount
            If .CellValue(irow, 1) = "" Or .CellValue(irow, 2) = 0 Then
                .SetCurCell .RowCount, "cc01_capacity"
                .SetFocus
                Exit Sub
            End If
        Next irow
        .Redraw = False
        .AddRow
        .SetCurCell .RowCount, "cc01_capacity"
        .SetFocus
        .Redraw = True
    End With
End Sub

Private Sub cmdAddRow_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub cmdDeleteRow_Click()
Dim reply As String
On Error Resume Next
    With grdMachineCapacity
        If .RowCount = 0 Then Exit Sub
        reply = MsgBox("Record " & .CellValue(.CurRow, "cc01_capacity") & " will be removed. Proceed?", vbYesNo, App.Title)
        If reply = vbYes Then
            .Redraw = False
            .RemoveRow (.CurRow)
            .Redraw = True
            .SetCurCell 1, "cc01_capacity"
        Else
            Exit Sub
        End If
    End With
End Sub

Private Sub cmdDeleteRow_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub
Private Sub cmdSave_Click()
Dim lRow As Long
Dim DetLine As String
Dim item As String * 50
    If bEmpty(cboMcCapacity) = True Then Exit Sub
    With grdMachineCapacity
        For lRow = 1 To .RowCount
            LSet item = Trim$(.CellValue(lRow, "cc01_capacity"))
            DetLine = DetLine & item
            LSet item = Trim$(.CellValue(lRow, "cc01_wastage"))
            DetLine = DetLine & item
            LSet item = Trim$(.CellValue(lRow, "mc01_output"))
            DetLine = DetLine & item
            LSet item = Trim$(.CellValue(lRow, "mc01_setting_time"))
            DetLine = DetLine & item
        Next lRow
    End With
    Set cdbcn = New clsDBAccess
    'Set cdbcn = New cDataAccess.clsDBAccess
    cdbcn.DataSource = strDatabase
    If Not cdbcn.OpenConnection = True Then GoTo ErrHandler
    If cdbcn.proc_usp_insert_update_mc01_machine_capacity_hdr(cboMcCapacity.Text, strUser, strUser, DetLine) = 0 Then
        Set cdbcn = Nothing
        MsgBox "Record successfully saved.", vbInformation, App.Title
        Exit Sub
    Else
        Set cdbcn = Nothing
        MsgBox "There is an error executing the command.", vbCritical, App.Title
        Exit Sub
    End If
    Set cdbcn = Nothing
Exit Sub
ErrHandler:
    Set cdbcn = Nothing
    If err.Number <> 0 Then prompt_errlog err, frmMachineCapacity.Name, "cmdSave_Click"
End Sub

Private Sub cmdSave_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
Dim lCol  As Long
On Error GoTo ErrHandler
    m_sDecSep = Mid$(CStr(0.5), 2, 1)
    With grdMachineCapacity
        .Redraw = True
        .DefaultRowHeight = 1 + 18 * 15 / Screen.TwipsPerPixelY
        .HighlightSelIcons = False
        .FocusRect = False
        .DrawRowText = False
        .MemMngWantFreeRows = 75
        'Set .BackgroundPicture = MDIForm1.picBackground.Picture
        .Font.Name = "Tahoma"
        .Font.Size = 10
        With .Header.Font
            .Name = "Tahoma"
            .Bold = True
            .Size = 10
        End With
        exec_proc_usp_sel_cc01_capacity_02
        
        With .AddCol(sKey:="cc01_capacity", sHeader:="Capacity", lWidth:=100)
            .eType = igCellCombo
            .sCtrlKey = "cc01_capacity"
        End With
        With .AddCol(sKey:="cc01_wastage", sHeader:="Wastage", lWidth:=100)
            .eTextFlags = igTextRight
        End With
        With .AddCol(sKey:="mc01_output", sHeader:="Output", lWidth:=100)
            .eTextFlags = igTextRight
        End With
        With .AddCol(sKey:="mc01_setting_time", sHeader:="Setting Time(Hr)", lWidth:=100)
            .eTextFlags = igTextRight
        End With
        .RowCount = 1
        .Editable = True
        .Redraw = True
    End With
    exec_proc_usp_sel_mc01_machinery_02
Exit Sub
ErrHandler:
    If err.Number <> 0 Then prompt_errlog err, frmMachineCapacity.Name, "Form_Load Event"
End Sub
Private Sub exec_proc_usp_sel_cc01_capacity_02()
Dim Rscc01_capacity_02 As ADODB.Recordset
Dim scc01_capacity As String
On Error GoTo ErrHandler
    Set cdbcn = New clsDBAccess
    'Set cdbcn = New cDataAccess.clsDBAccess
    cdbcn.DataSource = strDatabase
    If Not cdbcn.OpenConnection = True Then GoTo ErrHandler
    Set Rscc01_capacity_02 = New ADODB.Recordset
        If cdbcn.proc_usp_sel_cc01_capacity_02(Rscc01_capacity_02) = 0 Then
            On Error GoTo 0
            With grdMachineCapacity.Combos.Add("cc01_capacity")
                Do While Not Rscc01_capacity_02.EOF
                    scc01_capacity = Rscc01_capacity_02![cc01_capacity]
                    .AddItem sItemText:=scc01_capacity, vItemValue:=scc01_capacity
                    Rscc01_capacity_02.MoveNext
                Loop
                With .Font
                    .Name = "Arial"
                    .Size = 8
                    .Bold = False
                End With
                .AutoAdjustWidth
            End With
        End If
    Set cdbcn = Nothing
    Set Rscc01_capacity_02 = Nothing
Exit Sub
ErrHandler:
    If err.Number <> 0 Then prompt_errlog err, frmMachineCapacity.Name, "exec_proc_usp_sel_cc01_capacity_02"
    Set cdbcn = Nothing
End Sub
Private Sub grdMachineCapacity_AfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
Dim Rscc01_capacity_03 As ADODB.Recordset
On Error GoTo ErrHandler
    Set cdbcn = New clsDBAccess
    'Set cdbcn = New cDataAccess.clsDBAccess
    cdbcn.DataSource = strDatabase
    If Not cdbcn.OpenConnection = True Then GoTo ErrHandler
        Set Rscc01_capacity_03 = New ADODB.Recordset
        With grdMachineCapacity
            If cdbcn.proc_usp_sel_cc01_capacity_03(.CellValue(.CurRow, "cc01_capacity"), Rscc01_capacity_03) = 0 Then
                If Not (Rscc01_capacity_03.BOF Or Rscc01_capacity_03.EOF) Then
                    .CellValue(.CurRow, "cc01_wastage") = Trim$(Rscc01_capacity_03!cc01_wastage)
                Else
                    Set cdbcn = Nothing
                    Set Rscc01_capacity_03 = Nothing
                    MsgBox "Record not found.", vbCritical, App.Title
                End If
            Else
                Set Rscc01_capacity_03 = Nothing
                Set cdbcn = Nothing
                MsgBox "There is an error executing the command.", vbCritical, App.Title
                Exit Sub
            End If
        End With
    Set Rscc01_capacity_03 = Nothing
    Set cdbcn = Nothing
Exit Sub
ErrHandler:
    Set cdbcn = Nothing
    If err.Number <> 0 Then prompt_errlog err, frmMachineCapacity.Name, "proc_usp_sel_cc01_capacity_03"
End Sub

Private Sub grdMachineCapacity_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
    bDoDefault = False
End Sub

Private Sub grdMachineCapacity_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)
    lblCapacity.Caption = "Record Number " & lRow & " of " & grdMachineCapacity.RowCount
End Sub
Private Sub exec_proc_usp_sel_mc01_machine_capacity_hdr_01()
Dim lCol As Long
Dim Rsmc01_machine_capacity_hdr_01 As ADODB.Recordset
On Error GoTo ErrHandler
    Set cdbcn = New clsDBAccess
    'Set cdbcn = New cDataAccess.clsDBAccess
    cdbcn.DataSource = strDatabase
    If Not cdbcn.OpenConnection = True Then GoTo ErrHandler
        Set Rsmc01_machine_capacity_hdr_01 = New ADODB.Recordset
        If cdbcn.proc_usp_sel_mc01_machine_capacity_hdr_01(cboMcCapacity.Text, Rsmc01_machine_capacity_hdr_01) = 0 Then
            If Not (Rsmc01_machine_capacity_hdr_01.BOF Or Rsmc01_machine_capacity_hdr_01.EOF) Then
                With grdMachineCapacity
                    .Clear
                    .Redraw = False
                    .FillFromRS Rsmc01_machine_capacity_hdr_01
                    .Redraw = True
                    .SetCurCell 1, 1
                    For lCol = 1 To .ColCount
                        .AutoWidthCol lCol
                    Next lCol
                End With
                Set Rsmc01_machine_capacity_hdr_01 = Nothing
                Set cdbcn = Nothing
            Else
                grdMachineCapacity.Clear
                grdMachineCapacity.RowCount = 1
                grdMachineCapacity.SetCurCell 1, 1
            End If
        End If
    Set Rsmc01_machine_capacity_hdr_01 = Nothing
    Set cdbcn = Nothing
Exit Sub
ErrHandler:
    Set Rsmc01_machine_capacity_hdr_01 = Nothing
    Set cdbcn = Nothing
    If err.Number <> 0 Then prompt_errlog err, frmMachineCapacity.Name, "proc_usp_sel_mc01_machine_capacity_hdr_01"
End Sub
Private Sub pKeyFilter(ByRef piKeyAscii As Integer)
   If piKeyAscii > 27 Then ' do not process control keys
      If Not ((piKeyAscii >= 48 And piKeyAscii <= 57) Or (piKeyAscii = Asc(m_sDecSep))) Then
         piKeyAscii = 0
      End If
   End If
End Sub

Private Sub grdMachineCapacity_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub grdMachineCapacity_KeyPress(KeyAscii As Integer)
    With grdMachineCapacity
        If .CurCol = 3 Then pKeyFilter KeyAscii
        If .CurCol = 4 Then pKeyFilter KeyAscii
    End With
End Sub

Private Sub grdMachineCapacity_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
    Select Case grdMachineCapacity.ColKey(lCol)
        Case "cc01_wastage"
            bCancel = True
    End Select
End Sub
Private Sub grdMachineCapacity_Validate(Cancel As Boolean)
    grdMachineCapacity.CommitEdit
End Sub
