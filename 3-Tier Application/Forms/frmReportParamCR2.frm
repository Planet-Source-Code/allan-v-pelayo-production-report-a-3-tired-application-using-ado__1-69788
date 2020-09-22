VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReportParamCR2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Production Output Parameter"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReportParamCR2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   4455
      Begin VB.ComboBox cboOuput 
         Height          =   360
         Index           =   1
         ItemData        =   "frmReportParamCR2.frx":127A
         Left            =   1080
         List            =   "frmReportParamCR2.frx":1281
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1080
         Width           =   3255
      End
      Begin VB.ComboBox cboOuput 
         Height          =   360
         Index           =   0
         ItemData        =   "frmReportParamCR2.frx":128A
         Left            =   1080
         List            =   "frmReportParamCR2.frx":1291
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   720
         Width           =   3255
      End
      Begin MSComCtl2.DTPicker dtpOutput 
         Height          =   330
         Index           =   0
         Left            =   1080
         TabIndex        =   0
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   16384001
         CurrentDate     =   39392
      End
      Begin MSComCtl2.DTPicker dtpOutput 
         Height          =   330
         Index           =   1
         Left            =   2880
         TabIndex        =   1
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   16384001
         CurrentDate     =   39392
      End
      Begin VB.Line Line4 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         X1              =   4440
         X2              =   0
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         X1              =   4440
         X2              =   4440
         Y1              =   120
         Y2              =   1680
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Operator"
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   1160
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Machine"
         Height          =   270
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   800
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "To"
         Height          =   255
         Left            =   2520
         TabIndex        =   6
         Top             =   405
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "From"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   405
         Width           =   615
      End
   End
   Begin LVbuttons.LaVolpeButton cmdClose 
      Height          =   405
      Left            =   1920
      TabIndex        =   3
      Top             =   1920
      WhatsThisHelpID =   1920
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
      COLTYPE         =   2
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmReportParamCR2.frx":129A
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
   Begin LVbuttons.LaVolpeButton cmdPreview 
      Height          =   405
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      WhatsThisHelpID =   1920
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   714
      BTYPE           =   6
      TX              =   "&Preview"
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
      COLTYPE         =   2
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmReportParamCR2.frx":12B6
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
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000E&
      X1              =   120
      X2              =   4560
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      X1              =   120
      X2              =   4560
      Y1              =   1800
      Y2              =   1800
   End
End
Attribute VB_Name = "frmReportParamCR2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cdbcn As clsDBAccess
'Dim cdbcn As cDataAccess.clsDBAccess
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    If cboOuput(0).ListIndex = -1 Then cboOuput(0).Text = "ALL"
    If cboOuput(1).ListIndex = -1 Then cboOuput(1).Text = "ALL"
    If dtpOutput(0).Value > dtpOutput(1).Value Then MsgBox "Beginning Date is greater than Ending Date.", vbCritical, App.Title: Exit Sub
    frmReportCR2.Show
    Unload Me
End Sub
Private Sub SetDates()
Dim sDayFirst As String, sDayLast As String, sMonth As String, sYear As String
sDayLast = GetLastDayOfMonth(dtpOutput(0).Value)
sDayFirst = "01"
sYear = Year(dtpOutput(0).Value)
sMonth = Month(dtpOutput(0).Value)
Label2.Caption = sMonth & "/" & sDayLast & "/" & sYear
Label1.Caption = sMonth & "/" & sDayFirst & "/" & sYear
dtpOutput(0).Value = Label1.Caption
dtpOutput(1).Value = Label2.Caption
End Sub
Private Function GetLastDayOfMonth(pDate As Date)
    GetLastDayOfMonth = Day(DateSerial(Year(pDate), Month(pDate) + 1, 0))
End Function

Private Sub dtpOutput_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
    dtpOutput(0).Value = Date
    dtpOutput(1).Value = Date
    SetDates
    exec_proc_usp_sel_mc01_machinery_02
    exec_proc_usp_sel_op01_operator_02
End Sub
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
                If Not IsNull(Rsmc01_machinery_02!mc01_machine) Then cboOuput(0).AddItem Rsmc01_machinery_02!mc01_machine
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
    If err.Number <> 0 Then prompt_errlog err, frmOutputView.Name, "proc_usp_sel_mc01_machinery_02"
End Sub
Private Sub exec_proc_usp_sel_op01_operator_02()
Dim Rsop01_operator_02 As ADODB.Recordset
On Error GoTo ErrHandler
    Set cdbcn = New clsDBAccess
    'Set cdbcn = New cDataAccess.clsDBAccess
    cdbcn.DataSource = strDatabase
    If Not cdbcn.OpenConnection = True Then GoTo ErrHandler
    Set Rsop01_operator_02 = New ADODB.Recordset
        If cdbcn.proc_usp_sel_op01_operator_02(Rsop01_operator_02) = 0 Then
            If Rsop01_operator_02.RecordCount > 0 Then Rsop01_operator_02.MoveFirst
                Do Until Rsop01_operator_02.EOF
                If Not IsNull(Rsop01_operator_02!hr01_emplyee_id) Then cboOuput(1).AddItem Rsop01_operator_02!hr01_emplyee_id
                    Rsop01_operator_02.MoveNext
                DoEvents
                Loop
        Else
            Set Rsop01_operator_02 = Nothing
            Set cdbcn = Nothing
            MsgBox "There is an error executing the command.", vbCritical, App.Title
            Exit Sub
        End If
    Set Rsop01_operator_02 = Nothing
    Set cdbcn = Nothing
Exit Sub
ErrHandler:
    Set Rsop01_operator_02 = Nothing
    Set cdbcn = Nothing
    If err.Number <> 0 Then prompt_errlog err, frmOutputView.Name, "proc_usp_sel_op01_operator_02"
End Sub

