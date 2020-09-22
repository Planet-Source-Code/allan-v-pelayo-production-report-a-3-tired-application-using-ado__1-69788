VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7530
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   7530
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   7335
      Begin VB.CheckBox chkUser 
         Caption         =   "Enabled"
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
         Left            =   1440
         TabIndex        =   4
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox txtUser 
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
         IMEMode         =   3  'DISABLE
         Index           =   6
         Left            =   5400
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtUser 
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
         IMEMode         =   3  'DISABLE
         Index           =   5
         Left            =   5400
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtUser 
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
         Index           =   4
         Left            =   5400
         TabIndex        =   5
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtUser 
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
         Index           =   3
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txtUser 
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
         Index           =   2
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtUser 
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
         Index           =   1
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtUser 
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
         Index           =   0
         Left            =   1440
         TabIndex        =   0
         Top             =   330
         Width           =   1815
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   10
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUser.frx":127A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUser.frx":1814
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUser.frx":1DAE
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUser.frx":2688
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUser.frx":2F62
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUser.frx":34FC
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUser.frx":3A96
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUser.frx":4030
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUser.frx":45CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUser.frx":4B64
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin LVbuttons.LaVolpeButton cmdSearchEmployee 
         Height          =   315
         Left            =   3270
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   330
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   13160660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmUser.frx":50FE
         ALIGN           =   1
         IMGLST          =   "ImageList1"
         IMGICON         =   "5"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin VB.Line Line4 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         X1              =   7320
         X2              =   0
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         X1              =   7320
         X2              =   7320
         Y1              =   120
         Y2              =   2280
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Confirm Password"
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
         Index           =   6
         Left            =   3720
         TabIndex        =   18
         Top             =   1470
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Password"
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
         Index           =   5
         Left            =   4080
         TabIndex        =   17
         Top             =   1110
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "User Name"
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
         Index           =   4
         Left            =   4080
         TabIndex        =   16
         Top             =   750
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Middle Name"
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
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   1470
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "First Name"
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
         Index           =   2
         Left            =   240
         TabIndex        =   14
         Top             =   1110
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Name"
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
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   750
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Employee ID"
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
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
   End
   Begin LVbuttons.LaVolpeButton cmdSave 
      Height          =   410
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   1920
      _ExtentX        =   3387
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
      MICON           =   "frmUser.frx":511A
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
   Begin LVbuttons.LaVolpeButton cmdNew 
      Height          =   405
      Left            =   2160
      TabIndex        =   9
      Top             =   2520
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   714
      BTYPE           =   6
      TX              =   "&New"
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
      MICON           =   "frmUser.frx":5136
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
      Left            =   4200
      TabIndex        =   10
      Top             =   2520
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
      MICON           =   "frmUser.frx":5152
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
      X1              =   0
      X2              =   7320
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      X1              =   0
      X2              =   7320
      Y1              =   2400
      Y2              =   2400
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cdbcn As clsDBAccess
'Dim cdbcn As cDataAccess.clsDBAccess
Dim lCol As Long
Private Sub move_menu_form()
    Me.Left = frmMenu.Width
    Me.Top = 0
End Sub

Private Sub chkUser_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub chkUser_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub cmdNew_Click()
    ClearText
End Sub
Private Sub ClearText()
Dim i As Integer
    For i = 0 To 6
        txtUser(i).Text = vbNullString
    Next i
    chkUser.Value = 1
    txtUser(0).Enabled = True
    txtUser(0).SetFocus
End Sub

Private Sub cmdNew_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub cmdSave_Click()
Dim i As Integer
On Error GoTo ErrHandler
    For i = 0 To 6
        If bEmpty(txtUser(i)) = True Then Exit Sub
    Next i
    For i = 4 To 6
        If Len(txtUser(i).Text) < 5 Then MsgBox "User and Password should not less than 5 characters long.", vbInformation, App.Title: Exit Sub
    Next i
    If Trim$(txtUser(5).Text) <> Trim$(txtUser(6).Text) Then
        MsgBox "Password confirmation failed.", vbCritical, App.Title
        txtUser(5).SetFocus
        Exit Sub
    Else
        Set cdbcn = New clsDBAccess
        cdbcn.DataSource = strDatabase
        If Not cdbcn.OpenConnection = True Then GoTo ErrHandler
            If cdbcn.proc_usp_insert_update_us01_user(Trim$(txtUser(0).Text), Trim$(txtUser(1).Text), Trim$(txtUser(2).Text), _
                Trim$(txtUser(3).Text), Trim$(txtUser(4).Text), Decode_Pass(Trim$(txtUser(5).Text)), chkUser.Value) = 0 Then
                frmUserView.exec_proc_usp_sel_us01_user_02
                Set cdbcn = Nothing
                MsgBox "User has been successfully created.", vbInformation, App.Title
                Exit Sub
            Else
                Set cdbcn = Nothing
                MsgBox "An error occured while trying to save records.", vbCritical, App.Title
                Exit Sub
            End If
        Set cdbcn = Nothing
    End If
Exit Sub
ErrHandler:
    Set cdbcn = Nothing
    If err.Number <> 0 Then prompt_errlog err, frmUser.Name, "cmdSave_Click"
End Sub

Private Sub cmdSave_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub cmdSearchEmployee_Click()
    bUserNew = True
    bOutputOperator = False
    frmSearchOperator.Show 1
End Sub

Private Sub cmdSearchEmployee_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtUser_Change(Index As Integer)
    Select Case Index
        Case 0
            exec_proc_usp_sel_op01_operator_03
    End Select
End Sub

Private Sub txtUser_GotFocus(Index As Integer)
Dim i As Integer
    For i = 0 To 6
        SelText txtUser(i)
    Next i
End Sub

Private Sub txtUser_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        bUserNew = True
        bOutputOperator = False
        frmSearchOperator.Show 1
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtUser_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{Tab}"
    Select Case Index
        Case 1, 2, 3, 4
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub
Private Sub exec_proc_usp_sel_op01_operator_03()
Dim Rsop01_operator_03 As ADODB.Recordset
On Error GoTo ErrHandler
    Set cdbcn = New clsDBAccess
    'Set cdbcn = New cDataAccess.clsDBAccess
    cdbcn.DataSource = strDatabase
    If Not cdbcn.OpenConnection = True Then GoTo ErrHandler
    Set Rsop01_operator_03 = New ADODB.Recordset
        If cdbcn.proc_usp_sel_op01_operator_03(Trim$(txtUser(0).Text), Rsop01_operator_03) = 0 Then
            If Not (Rsop01_operator_03.BOF Or Rsop01_operator_03.EOF) Then
                txtUser(1).Text = Trim$(Rsop01_operator_03!op01_lastname)
                txtUser(2).Text = Trim$(Rsop01_operator_03!op01_firstname)
                txtUser(3).Text = Trim$(Rsop01_operator_03!op01_middlename)
            Else
                txtUser(1).Text = vbNullString
                txtUser(2).Text = vbNullString
                txtUser(3).Text = vbNullString
                Set Rsop01_operator_03 = Nothing
                Set cdbcn = Nothing
            End If
        Else
            Set Rsop01_operator_03 = Nothing
            Set cdbcn = Nothing
            MsgBox "There is an error executing the command.", vbCritical, App.Title
            Exit Sub
        End If
    Set cdbcn = Nothing
    Set Rsop01_operator_03 = Nothing
Exit Sub
ErrHandler:
    Set Rsop01_operator_03 = Nothing
    Set cdbcn = Nothing
    If err.Number <> 0 Then prompt_errlog err, frmUser.Name, "proc_usp_sel_op01_operator_03"
End Sub
Function Decode_Pass(p_str As String) As String
Dim i As Integer
Dim strs As String
    For i = 1 To Len(p_str) Step 1
        strs = strs + Chr(Asc(Mid(p_str, i, 1)) * 2)
    Next i
        Decode_Pass = strs
End Function
Function UnCode_Pass(p_str As String) As String
Dim i As Integer
Dim strs As String
    For i = 1 To Len(p_str) Step 1
        strs = strs + Chr(Asc(Mid(p_str, i, 1)) / 2)
    Next i
        UnCode_Pass = strs
End Function

