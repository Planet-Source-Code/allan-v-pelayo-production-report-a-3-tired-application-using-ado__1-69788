VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Object = "{3F666E72-7F79-447A-BCFF-C42C44E3DBE7}#1.0#0"; "iGrid251_75B4A91C.ocx"
Begin VB.Form frmUserView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   Icon            =   "frmUserView.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8775
   ScaleWidth      =   11910
   Begin iGrid251_75B4A91C.iGrid grdUser 
      Height          =   7815
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   13785
   End
   Begin VB.PictureBox picUser 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11835
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   8160
      Width           =   11895
      Begin LVbuttons.LaVolpeButton cmdNew 
         Height          =   410
         Left            =   120
         TabIndex        =   0
         Top             =   80
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
         COLTYPE         =   2
         BCOL            =   14215660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmUserView.frx":127A
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
      Begin LVbuttons.LaVolpeButton cmdView 
         Height          =   405
         Left            =   2160
         TabIndex        =   1
         Top             =   75
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   714
         BTYPE           =   6
         TX              =   "&Edit"
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
         MICON           =   "frmUserView.frx":1296
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
         TabIndex        =   2
         Top             =   75
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
         COLTYPE         =   2
         BCOL            =   14215660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmUserView.frx":12B2
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
   End
   Begin VB.Label lblUser 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   4
      Top             =   7820
      Width           =   11895
   End
End
Attribute VB_Name = "frmUserView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cdbcn As clsDBAccess
'Dim cdbcn As cDataAccess.clsDBAccess
Dim lCol As Long
Public Sub move_menu_form()
    Me.Left = frmMenu.Width
    Me.Top = 0
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdNew_Click()
    frmUser.Caption = "New"
    frmUser.Show 1
End Sub
Private Sub cmdView_Click()
    exec_proc_usp_sel_us01_user_03
End Sub
Private Sub Form_Load()
On Error GoTo ErrHandler
    move_menu_form
    With grdUser
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
            .Bold = False
            .Size = 10
        End With
        With .AddCol(sKey:="hr01_emplyee_id", sHeader:="ID", lWidth:=50, bvisible:=True)
            .eTextFlags = igTextRight
        End With
        With .AddCol(sKey:="hr01_lastname", sHeader:="Last Name", lWidth:=50)
            .eTextFlags = igTextLeft
        End With
        With .AddCol(sKey:="hr01_firstname", sHeader:="First Name", lWidth:=50)
            .eTextFlags = igTextLeft
        End With
        With .AddCol(sKey:="hr01_middlename", sHeader:="Middle Name", lWidth:=50)
            .eTextFlags = igTextLeft
        End With
        With .AddCol(sKey:="us01_user_name", sHeader:="User Name", lWidth:=50)
            .eTextFlags = igTextLeft
        End With
        With .AddCol(sKey:="us01_status", sHeader:="Enabled", lWidth:=50)
            .eType = igCellCheck
            .eTypeFlags = igCheckBox3State
        End With
        For lCol = 1 To .ColCount
            .AutoWidthCol lCol
        Next lCol
        .Editable = False
        .Redraw = True
    End With
    exec_proc_usp_sel_us01_user_02
Exit Sub
ErrHandler:
    If err.Number <> 0 Then prompt_errlog err, frmUserView.Name, "Form_Load Event"
End Sub
Private Sub Form_Resize()
On Error Resume Next
   With picUser
    .Move 0, Me.ScaleHeight - .Height - 120, Me.ScaleWidth
   End With
   With lblUser
      .Move 0, Me.ScaleHeight - .Height - picUser.Height - 120, Me.ScaleWidth
   End With
   With grdUser
      .Move 0, .Top, Me.ScaleWidth, lblUser.Top - .Top - 60
   End With
End Sub

Private Sub grdUser_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
    bDoDefault = False
End Sub
Public Sub exec_proc_usp_sel_us01_user_02()
Dim Rsus01_user_02 As ADODB.Recordset
On Error GoTo ErrHandler
    Set cdbcn = New clsDBAccess
    'Set cdbcn = New cDataAccess.clsDBAccess
    cdbcn.DataSource = strDatabase
    If Not cdbcn.OpenConnection = True Then GoTo ErrHandler
    Set Rsus01_user_02 = New ADODB.Recordset
    If cdbcn.proc_usp_sel_us01_user_02(Rsus01_user_02) = 0 Then
            If Not (Rsus01_user_02.BOF Or Rsus01_user_02.EOF) Then
                With grdUser
                    .Redraw = False
                    .FillFromRS Rsus01_user_02
                    For lCol = 1 To .ColCount
                        .AutoWidthCol lCol
                    Next lCol
                    .Redraw = True
                    .SetCurCell 1, 2
                End With
            Else
                Set Rsus01_user_02 = Nothing
                Set cdbcn = Nothing
                MsgBox "No record found.", vbInformation, App.Title
                Exit Sub
            End If
        Else
            Set Rsus01_user_02 = Nothing
            Set cdbcn = Nothing
            MsgBox "There is an error executing the command.", vbCritical, App.Title
            Exit Sub
        End If
    Set Rsus01_user_02 = Nothing
    Set cdbcn = Nothing
Exit Sub
ErrHandler:
    Set cdbcn = Nothing
    If err.Number <> 0 Then prompt_errlog err, frmUserView.Name, "proc_usp_sel_us01_user_02"
End Sub
Private Sub grdUser_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)
    lblUser.Caption = "Record Number " & lRow & " of " & grdUser.RowCount
End Sub
Private Sub grdUser_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)
    exec_proc_usp_sel_us01_user_03
End Sub
Private Sub exec_proc_usp_sel_us01_user_03()
Dim Rsus01_user_03 As ADODB.Recordset
On Error GoTo ErrHandler
    Set cdbcn = New clsDBAccess
    'Set cdbcn = New cDataAccess.clsDBAccess
    cdbcn.DataSource = strDatabase
    If Not cdbcn.OpenConnection = True Then GoTo ErrHandler
    Set Rsus01_user_03 = New ADODB.Recordset
            With grdUser
                If .CurRow <> 0 And .ColCount <> 0 Then
                    If cdbcn.proc_usp_sel_us01_user_03(.CellValue(.CurRow, "hr01_emplyee_id"), Rsus01_user_03) = 0 Then
                        If Not (Rsus01_user_03.BOF Or Rsus01_user_03.EOF) Then
                            frmUser.txtUser(0).Text = Trim$(Rsus01_user_03!hr01_emplyee_id)
                            frmUser.txtUser(1).Text = Trim$(Rsus01_user_03!hr01_lastname)
                            frmUser.txtUser(2).Text = Trim$(Rsus01_user_03!hr01_firstname)
                            frmUser.txtUser(3).Text = Trim$(Rsus01_user_03!hr01_middlename)
                            frmUser.txtUser(4).Text = Trim$(Rsus01_user_03!us01_user_name)
                            frmUser.txtUser(5).Text = UnCode_Pass(Trim$(Rsus01_user_03!us01_password))
                            frmUser.chkUser.Value = CInt(Abs(Rsus01_user_03!us01_status))
                            frmUser.txtUser(0).Enabled = False
                            frmUser.Caption = "User:" & Trim$(Rsus01_user_03!hr01_emplyee_id)
                            Set Rsus01_user_03 = Nothing
                            Set cdbcn = Nothing
                            frmUser.Show 1
                        Else
                            Set Rsus01_user_03 = Nothing
                            Set cdbcn = Nothing
                            MsgBox "No records available.", vbInformation, App.Title
                            Exit Sub
                        End If
                    Else
                        Set Rsus01_user_03 = Nothing
                        Set cdbcn = Nothing
                        MsgBox "There is an error executing the command.", vbCritical, App.Title
                        Exit Sub
                    End If
                Else
                    Set Rsus01_user_03 = Nothing
                    Set cdbcn = Nothing
                    MsgBox "Select record to edit.", vbInformation, App.Title
                    Exit Sub
                End If
            End With
        Set Rsus01_user_03 = Nothing
        Set cdbcn = Nothing
        Exit Sub
ErrHandler:
    Set Rsus01_user_03 = Nothing
    Set cdbcn = Nothing
    If err.Number <> 0 Then prompt_errlog err, frmUserView.Name, "exec_proc_usp_sel_us01_user_03"
End Sub
Private Sub grdUser_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then exec_proc_usp_sel_us01_user_03
End Sub
Function UnCode_Pass(p_str As String) As String
Dim i As Integer
Dim strs As String
    For i = 1 To Len(p_str) Step 1
        strs = strs + Chr(Asc(Mid(p_str, i, 1)) / 2)
    Next i
        UnCode_Pass = strs
End Function



