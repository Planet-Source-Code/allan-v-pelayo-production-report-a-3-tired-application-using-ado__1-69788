VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Object = "{23F895D7-45A6-4886-931B-89D88C2857ED}#1.2#0"; "iGrid250_75B4A91C.ocx"
Begin VB.Form frmSearchOperator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Operator"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7410
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSearchOperator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   7410
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboSrcEmployee 
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
      ItemData        =   "frmSearchOperator.frx":127A
      Left            =   1560
      List            =   "frmSearchOperator.frx":128A
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   100
      Width           =   2655
   End
   Begin VB.TextBox txtSrcEmployee 
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
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   5655
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   7215
      Begin iGrid250_75B4A91C.iGrid grdSrchEmployee 
         Height          =   4425
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   7805
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line Line6 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         X1              =   0
         X2              =   7200
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Line Line5 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         X1              =   7200
         X2              =   7200
         Y1              =   120
         Y2              =   5040
      End
      Begin VB.Label lblSrchEmployee 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   4680
         Width           =   6975
      End
   End
   Begin LVbuttons.LaVolpeButton cmdOK 
      Default         =   -1  'True
      Height          =   405
      Left            =   3240
      TabIndex        =   2
      Top             =   6360
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   714
      BTYPE           =   6
      TX              =   "&OK"
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
      MICON           =   "frmSearchOperator.frx":12DA
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
      Height          =   405
      Left            =   5280
      TabIndex        =   3
      Top             =   6360
      Width           =   1920
      _ExtentX        =   3387
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
      MICON           =   "frmSearchOperator.frx":12F6
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
   Begin VB.Line Line4 
      BorderColor     =   &H8000000E&
      X1              =   120
      X2              =   7320
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      X1              =   120
      X2              =   7320
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000E&
      X1              =   120
      X2              =   7320
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      X1              =   120
      X2              =   7320
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label lblSrcEmployee 
      Alignment       =   1  'Right Justify
      Caption         =   "Search By"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   195
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Enter Text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   540
      Width           =   1455
   End
End
Attribute VB_Name = "frmSearchOperator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cdbcn As clsDBAccess
'Dim cdbcn As cDataAccess.clsDBAccess
Dim lCol As Long
Private Sub exec_proc_usp_sel_op01_operator_04()
Dim Rsop01_operator_04 As ADODB.Recordset
On Error GoTo ErrHandler
    Set cdbcn = New clsDBAccess
    'Set cdbcn = New cDataAccess.clsDBAccess
    cdbcn.DataSource = strDatabase
    If Not cdbcn.OpenConnection = True Then GoTo ErrHandler
    Set Rsop01_operator_04 = New ADODB.Recordset
        If Trim$(txtSrcEmployee.Text) = vbNullString Then grdSrchEmployee.Clear: Exit Sub
            If cdbcn.proc_usp_sel_op01_operator_04(cboSrcEmployee.Text, Trim$(txtSrcEmployee.Text), Rsop01_operator_04) = 0 Then
                If Not (Rsop01_operator_04.BOF Or Rsop01_operator_04.EOF) Then
                    With grdSrchEmployee
                        .Redraw = False
                        .FillFromRS Rsop01_operator_04
                        For lCol = 1 To .ColCount
                            .AutoWidthCol lCol
                        Next lCol
                        .Redraw = True
                        .SetCurCell 1, 1
                    End With
                Else
                    Set Rsop01_operator_04 = Nothing
                    Set cdbcn = Nothing
                    grdSrchEmployee.Clear
                    Exit Sub
                End If
            Else
                Set Rsop01_operator_04 = Nothing
                Set cdbcn = Nothing
                MsgBox "An error occur while trying to execute the command.", vbCritical, App.Title
                Exit Sub
            End If
        Set Rsop01_operator_04 = Nothing
        Set cdbcn = Nothing
Exit Sub
ErrHandler:
    Set Rsop01_operator_04 = Nothing
    Set cdbcn = Nothing
    If err.Number <> 0 Then prompt_errlog err, frmSearchOperator.Name, "exec_proc_usp_sel_op01_operator_04"
End Sub
Private Sub cboSrcEmployee_Change()
    txtSrcEmployee.Text = vbNullString
End Sub
Private Sub cboSrcEmployee_Click()
    txtSrcEmployee.Text = vbNullString
End Sub

Private Sub cboSrcEmployee_KeyDown(KeyCode As Integer, Shift As Integer)
    If bOutputOperator = True Then
        If KeyCode = vbKeyEscape Then Unload Me: frmOutput.txtOutput(2).SetFocus
    Else
        If bUserNew = True Then
            If KeyCode = vbKeyEscape Then Unload Me: frmUser.txtUser(0).SetFocus
        End If
    End If
End Sub
Private Sub cmdClose_Click()
    If bOutputOperator = True Then
        frmOutput.txtOutput(2).SetFocus
    Else
        If bUserNew = True Then
           frmUser.txtUser(0).SetFocus
        End If
    End If
End Sub
Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)
If bOutputOperator = True Then
        If KeyCode = vbKeyEscape Then Unload Me: frmOutput.txtOutput(2).SetFocus
    Else
        If bUserNew = True Then
            If KeyCode = vbKeyEscape Then Unload Me: frmUser.txtUser(0).SetFocus
        End If
    End If
End Sub
Private Sub cmdOK_Click()
    GetEmployee
End Sub
Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)
If bOutputOperator = True Then
        If KeyCode = vbKeyEscape Then Unload Me: frmOutput.txtOutput(2).SetFocus
    Else
        If bUserNew = True Then
            If KeyCode = vbKeyEscape Then Unload Me: frmUser.txtUser(0).SetFocus
        End If
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If bOutputOperator = True Then
        If KeyCode = vbKeyEscape Then Unload Me: frmOutput.txtOutput(2).SetFocus
    Else
        If bUserNew = True Then
            If KeyCode = vbKeyEscape Then Unload Me: frmUser.txtUser(0).SetFocus
        End If
    End If
End Sub
Private Sub Form_Load()
On Error Resume Next
    cboSrcEmployee.ListIndex = 1
    With grdSrchEmployee
        .Redraw = True
        .DefaultRowHeight = 1 + 18 * 15 / Screen.TwipsPerPixelY
        .HighlightSelIcons = False
        .FocusRect = True
        .DrawRowText = False
        .MemMngWantFreeRows = 75
        'Set .BackgroundPicture = MDIForm1.picBackground.Picture
        .RowMode = True
        .Font.Name = "Tahoma"
        .Font.Size = 10
        .Font.Bold = False
        With .Header.Font
            .Name = "Tahoma"
            .Bold = False
            .Size = 10
        End With
        With .AddCol(sKey:="hr01_emplyee_id", sHeader:="ID", lWidth:=50, bvisible:=True)
        End With
        With .AddCol(sKey:="op01_lastname", sHeader:="Last Name", lWidth:=50, bvisible:=True)
        End With
        With .AddCol(sKey:="op01_firstname", sHeader:="Fist Name", lWidth:=50, bvisible:=True)
        End With
        With .AddCol(sKey:="op01_middlename", sHeader:="Middle Name", lWidth:=50, bvisible:=True)
        End With
        For lCol = 1 To .ColCount
            .AutoWidthCol lCol
        Next lCol
        .Editable = False
        .Redraw = True
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    bUserNew = False
    bOutputOperator = False
End Sub

Private Sub grdSrchEmployee_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)
    lblSrchEmployee.Caption = "Record Number " & lRow & " of " & grdSrchEmployee.RowCount
End Sub
Private Sub grdSrchEmployee_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)
    GetEmployee
End Sub
Private Sub grdSrchEmployee_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
    If bOutputOperator = True Then
        If KeyCode = vbKeyEscape Then Unload Me: frmOutput.txtOutput(2).SetFocus
    Else
        If bUserNew = True Then
            If KeyCode = vbKeyEscape Then Unload Me: frmUser.txtUser(0).SetFocus
        End If
    End If
End Sub
Private Sub grdSrchEmployee_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then
        GetEmployee
    End If
End Sub
Private Sub txtSrcEmployee_Change()
    exec_proc_usp_sel_op01_operator_04
End Sub
Private Sub txtSrcEmployee_KeyDown(KeyCode As Integer, Shift As Integer)
    If bOutputOperator = True Then
        If KeyCode = vbKeyEscape Then Unload Me: frmOutput.txtOutput(2).SetFocus
    Else
        If bUserNew = True Then
            If KeyCode = vbKeyEscape Then Unload Me: frmUser.txtUser(0).SetFocus
        End If
    End If
End Sub
Private Sub GetEmployee()
On Error Resume Next
    With grdSrchEmployee
        If .CurCol > 0 And .CurRow > 0 Then
            If bOutputOperator = True Then
                frmOutput.txtOutput(2).Text = .CellValue(.CurRow, "hr01_emplyee_id")
                Unload Me
                frmOutput.txtOutput(2).SetFocus
            Else
                If bUserNew = True Then
                    frmUser.txtUser(0).Text = .CellValue(.CurRow, "hr01_emplyee_id")
                    Unload Me
                    frmUser.txtUser(0).SetFocus
                End If
            End If
        Else
            MsgBox "No record has been selected. Grid is empty.", vbCritical, App.Title
            Exit Sub
        End If
    End With
End Sub

