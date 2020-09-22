VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Application Sign-on Security"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4545
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   4545
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      ScaleHeight     =   1065
      ScaleWidth      =   705
      TabIndex        =   9
      Top             =   480
      Width           =   735
      Begin VB.Image Image1 
         Height          =   720
         Left            =   0
         Picture         =   "frmLogin.frx":0000
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.TextBox txtLogin 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox txtLogin 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Top             =   840
      Width           =   2415
   End
   Begin VB.ComboBox cboLogin 
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
      ItemData        =   "frmLogin.frx":0ECA
      Left            =   2040
      List            =   "frmLogin.frx":0ECC
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   460
      Width           =   2415
   End
   Begin LVbuttons.LaVolpeButton cmdLogin 
      Default         =   -1  'True
      Height          =   405
      Left            =   1440
      TabIndex        =   2
      Top             =   1800
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   714
      BTYPE           =   6
      TX              =   "&Login"
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
      MICON           =   "frmLogin.frx":0ECE
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "1"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton cmdClose 
      CausesValidation=   0   'False
      Height          =   405
      Left            =   3000
      TabIndex        =   3
      Top             =   1800
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   714
      BTYPE           =   6
      TX              =   "&Cancel"
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
      MICON           =   "frmLogin.frx":0EEA
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
      BorderColor     =   &H80000009&
      X1              =   120
      X2              =   4440
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      X1              =   120
      X2              =   4440
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter your User Name and Password to logon...."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   4455
   End
   Begin VB.Image Image3 
      Height          =   345
      Left            =   0
      Picture         =   "frmLogin.frx":0F06
      Stretch         =   -1  'True
      Top             =   0
      Width           =   30000
   End
   Begin VB.Label lblLogin 
      Alignment       =   1  'Right Justify
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   7
      Top             =   1245
      Width           =   1095
   End
   Begin VB.Label lblLogin 
      Alignment       =   1  'Right Justify
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   6
      Top             =   880
      Width           =   1215
   End
   Begin VB.Label lblLogin 
      Alignment       =   1  'Right Justify
      Caption         =   "Database"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   5
      Top             =   520
      Width           =   1095
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cdbcn As clsDBAccess
'Dim cdbcn As cDataAccess.clsDBAccess
Dim iAttempt As Integer
Dim i As Integer
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Sub cboLogin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        MDIForm1.bCloseMe = True
        Unload Me
    End If
End Sub
Private Sub cboLogin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub
Private Sub cmdClose_Click()
    MDIForm1.bCloseMe = True
    Unload Me
End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        MDIForm1.bCloseMe = True
        Unload Me
    End If
End Sub
Private Sub cmdLogin_Click()
Dim Rsus01_user_01 As ADODB.Recordset
Dim sRemainAttempt As String
On Error GoTo ErrHandler
    For i = 0 To 1
        If bEmpty(txtLogin(i)) = True Then Exit Sub
    Next i
    Set Rsus01_user_01 = New ADODB.Recordset
    Set cdbcn = New clsDBAccess
    'Set cdbcn = New cDataAccess.clsDBAccess
    strDatabase = cboLogin.Text
    cdbcn.DataSource = strDatabase
    If Not cdbcn.OpenConnection = True Then GoTo ErrHandler
        Screen.MousePointer = vbHourglass
        If cdbcn.proc_usp_sel_us01_user_01(Trim$(txtLogin(0).Text), Decode_Pass(Trim$(txtLogin(1).Text)), Rsus01_user_01) = 0 Then
            If (Rsus01_user_01.BOF Or Rsus01_user_01.EOF) Then
                iAttempt = iAttempt + 1
                If iAttempt = 4 Then
                    Check_Attempt
                Else
                    Set Rsus01_user_01 = Nothing
                    Set cdbcn = Nothing
                    sRemainAttempt = "Application sign-on security validation failed." & vbCrLf & "You have" & Str(4 - iAttempt) & " tries left."
                    If iAttempt = 3 Then
                        sRemainAttempt = "This is your last chance."
                    End If
                    MsgBox "Access Denied." & vbCrLf & sRemainAttempt, vbOKOnly + vbCritical, "Sign-on Security"
                    txtLogin(0).SetFocus
                End If
            Else
                strUser = Trim$(Rsus01_user_01!us01_user_name)
                strpassword = Trim$(Rsus01_user_01!us01_password)
                MDIForm1.StatusBar1.Panels(2).Text = "USER NAME : " & UCase(strUser)
                MDIForm1.StatusBar1.Panels(4).Text = "DATABASE : " & UCase(strDatabase)
                FillMenuExplorer
                Set Rsus01_user_01 = Nothing
                Set cdbcn = Nothing
                Unload Me
            End If
        End If
        Screen.MousePointer = vbDefault
    Set Rsus01_user_01 = Nothing
    Set cdbcn = Nothing
Exit Sub
ErrHandler:
    Set Rsus01_user_01 = Nothing
    Set cdbcn = Nothing
    If err.Number <> 0 Then prompt_errlog err, frmLogin.Name, "cmdLogin_Click Event"
    Screen.MousePointer = vbDefault
End Sub
Private Sub cmdLogin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        MDIForm1.bCloseMe = True
        Unload Me
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        MDIForm1.bCloseMe = True
        Unload Me
    End If
End Sub

Private Sub Form_Load()
Dim strWint As String
Dim lretval As Long
    EXEC_PROC_READ_TXT_FILE_DATABASE
    strWint = Space$(25)
    lretval = GetUserName(strWint, Len(strWint))
    strWinNT = LCase$(strWint)
    cboLogin.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cdbcn = Nothing
End Sub

Private Sub txtLogin_GotFocus(Index As Integer)
On Error Resume Next
    For i = 0 To 1
        SelText txtLogin(i)
    Next i
End Sub

Private Sub txtLogin_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        MDIForm1.bCloseMe = True
        Unload Me
    End If
End Sub

Private Sub Check_Attempt()
    If iAttempt >= 4 Then MDIForm1.bCloseMe = True: Unload Me
End Sub
Private Sub FillMenuExplorer()
On Error Resume Next
    frmMenu.trvMenu.Nodes.Clear
    Set nd = frmMenu.trvMenu.Nodes.Add(, , "Setup", "Setup", 3, 3)
    frmMenu.trvMenu.Nodes(1).Bold = True
    frmMenu.trvMenu.Nodes(1).Expanded = False
    Set nd = frmMenu.trvMenu.Nodes.Add("Setup", tvwChild, "", "User", 8, 9)
    Set nd = frmMenu.trvMenu.Nodes.Add("Setup", tvwChild, "", "Operator", 8, 9)
    Set nd = frmMenu.trvMenu.Nodes.Add("Setup", tvwChild, "", "Machinery", 8, 9)
    Set nd = frmMenu.trvMenu.Nodes.Add("Setup", tvwChild, "", "Capacity", 8, 9)
    Set nd = frmMenu.trvMenu.Nodes.Add("Setup", tvwChild, "", "Machine Capacity", 8, 9)
    frmMenu.trvMenu.Nodes(frmMenu.trvMenu.Nodes.Count).ForeColor = &H8000&
    
    Set nd = frmMenu.trvMenu.Nodes.Add(, , "Production", "Production", 3, 3)
    frmMenu.trvMenu.Nodes(1).Bold = True
    frmMenu.trvMenu.Nodes(1).Expanded = False
    Set nd = frmMenu.trvMenu.Nodes.Add("Production", tvwChild, "", "Output", 8, 9)
    frmMenu.trvMenu.Nodes(frmMenu.trvMenu.Nodes.Count).ForeColor = &H8000&

    
    Set nd = frmMenu.trvMenu.Nodes.Add(, , "Report", "Report", 3, 3)
    frmMenu.trvMenu.Nodes(1).Bold = True
    frmMenu.trvMenu.Nodes(1).Expanded = False
    Set nd = frmMenu.trvMenu.Nodes.Add("Report", tvwChild, "Production1", "Production", 8, 9)
    Set nd = frmMenu.trvMenu.Nodes.Add("Production1", tvwChild, , "Production Output", 8, 9)
    frmMenu.trvMenu.Nodes(frmMenu.trvMenu.Nodes.Count).ForeColor = &H8000&
End Sub

Private Sub txtLogin_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub
Private Sub EXEC_PROC_READ_TXT_FILE_DATABASE()
On Error Resume Next
Dim obj As cDataAccess.clsReadFile
Set obj = New clsReadFile
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim conn As ADODB.Connection
Set conn = New ADODB.Connection
      conn.Open "DRIVER={Microsoft Text Driver (*.txt; *.csv)};" & _
                   "DBQ=" & App.Path & ";", "", ""
      rs.Open "select * from [database.txt]", conn, adOpenStatic, _
                  adLockReadOnly, adCmdText
      Set cboLogin.DataSource = obj.PROC_READ_TXT_FILE_DATABASE
      
If rs.RecordCount > 0 Then rs.MoveFirst
    Do Until rs.EOF
        If Not IsNull(rs!Database) Then cboLogin.AddItem rs!Database
            rs.MoveNext
        DoEvents
    Loop
     Set obj = Nothing
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
