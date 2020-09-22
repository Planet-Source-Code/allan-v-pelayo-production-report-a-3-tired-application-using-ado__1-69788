VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Object = "{3F666E72-7F79-447A-BCFF-C42C44E3DBE7}#1.0#0"; "iGrid251_75B4A91C.ocx"
Begin VB.Form frmOperatorView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Operator"
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
   Icon            =   "frmOperatorView.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8775
   ScaleWidth      =   11910
   Begin iGrid251_75B4A91C.iGrid grdOperator 
      Height          =   7815
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   13785
   End
   Begin VB.PictureBox picOperator 
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
         MICON           =   "frmOperatorView.frx":127A
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
         Height          =   410
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
         MICON           =   "frmOperatorView.frx":1296
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
         MICON           =   "frmOperatorView.frx":12B2
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
   Begin VB.Label lblOperator 
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
      Top             =   7830
      Width           =   11895
   End
End
Attribute VB_Name = "frmOperatorView"
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
    frmOperator.Caption = "New"
    frmOperator.Show 1
End Sub
Private Sub cmdView_Click()
    exec_proc_usp_sel_op01_operator_01
End Sub
Private Sub Form_Load()
On Error GoTo ErrHandler
    move_menu_form
    With grdOperator
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
            .eTextFlags = igTextLeft
        End With
        With .AddCol(sKey:="op01_lastname", sHeader:="Last Name", lWidth:=50)
            .eTextFlags = igTextLeft
        End With
        With .AddCol(sKey:="op01_firstname", sHeader:="First Name", lWidth:=50)
           .eTextFlags = igTextLeft
        End With
        With .AddCol(sKey:="op01_middlename", sHeader:="Middle Name", lWidth:=50)
            .eTextFlags = igTextLeft
        End With
        With .AddCol(sKey:="op01_status", sHeader:="Enabled", lWidth:=50)
            .eType = igCellCheck
            .eTypeFlags = igCheckBox3State
        End With
        For lCol = 1 To .ColCount
            .AutoWidthCol lCol
        Next lCol
        .Editable = False
        .Redraw = True
    End With
    exec_proc_usp_sel_op01_operator_02
Exit Sub
ErrHandler:
    If err.Number <> 0 Then prompt_errlog err, frmOperatorView.Name, "Form_Load Event"
End Sub
Private Sub Form_Resize()
On Error Resume Next
   With picOperator
    .Move 0, Me.ScaleHeight - .Height - 120, Me.ScaleWidth
   End With
   With lblOperator
      .Move 0, Me.ScaleHeight - .Height - picOperator.Height - 120, Me.ScaleWidth
   End With
   With grdOperator
      .Move 0, .Top, Me.ScaleWidth, lblOperator.Top - .Top - 60
   End With
End Sub

Private Sub grdOperator_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
    bDoDefault = False
End Sub
Public Sub exec_proc_usp_sel_op01_operator_02()
Dim Rsop01_operator_02 As ADODB.Recordset
On Error GoTo ErrHandler
    Set cdbcn = New clsDBAccess
    'Set cdbcn = New cDataAccess.clsDBAccess
    cdbcn.DataSource = strDatabase
    If Not cdbcn.OpenConnection = True Then GoTo ErrHandler
    Set Rsop01_operator_02 = New ADODB.Recordset
        If cdbcn.proc_usp_sel_op01_operator_02(Rsop01_operator_02) = 0 Then
            If Not (Rsop01_operator_02.BOF Or Rsop01_operator_02.EOF) Then
                With grdOperator
                    .Redraw = False
                    .FillFromRS Rsop01_operator_02
                    For lCol = 1 To .ColCount
                        .AutoWidthCol lCol
                    Next lCol
                    .Redraw = True
                    .SetCurCell 1, 2
                End With
            Else
                Set Rsop01_operator_02 = Nothing
                Set cdbcn = Nothing
                MsgBox "No record found.", vbInformation, App.Title
                Exit Sub
            End If
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
    If err.Number <> 0 Then prompt_errlog err, frmOperatorView.Name, "proc_usp_sel_op01_operator_02"
End Sub
Private Sub grdOperator_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)
    lblOperator.Caption = "Record Number " & lRow & " of " & grdOperator.RowCount
End Sub
Private Sub grdOperator_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)
    exec_proc_usp_sel_op01_operator_01
End Sub
Private Sub exec_proc_usp_sel_op01_operator_01()
Dim Rsop01_operator_01 As ADODB.Recordset
On Error GoTo ErrHandler
    Set cdbcn = New clsDBAccess
    'Set cdbcn = New cDataAccess.clsDBAccess
    cdbcn.DataSource = strDatabase
    If Not cdbcn.OpenConnection = True Then GoTo ErrHandler
        Set Rsop01_operator_01 = New ADODB.Recordset
            With grdOperator
                If .CurRow <> 0 And .ColCount <> 0 Then
                    If cdbcn.proc_usp_sel_op01_operator_01(.CellValue(.CurRow, "hr01_emplyee_id"), Rsop01_operator_01) = 0 Then
                        If Not (Rsop01_operator_01.BOF Or Rsop01_operator_01.EOF) Then
                            frmOperator.txtOperator(0).Text = Trim$(Rsop01_operator_01!hr01_emplyee_id)
                            frmOperator.txtOperator(1).Text = Trim$(Rsop01_operator_01!op01_lastname)
                            frmOperator.txtOperator(2).Text = Trim$(Rsop01_operator_01!op01_firstname)
                            frmOperator.txtOperator(3).Text = Trim$(Rsop01_operator_01!op01_middlename)
                            frmOperator.chkOperator.Value = CInt(Abs(Rsop01_operator_01!op01_status))
                            frmOperator.txtOperator(0).Enabled = False
                            frmOperator.Caption = "Operator:" & Trim$(Rsop01_operator_01!hr01_emplyee_id)
                            frmOperator.Show 1
                        Else
                            Set Rsop01_operator_01 = Nothing
                            Set cdbcn = Nothing
                            MsgBox "No records available.", vbInformation, App.Title
                            Exit Sub
                        End If
                    Else
                        Set Rsop01_operator_01 = Nothing
                        Set cdbcn = Nothing
                        MsgBox "There is an error executing the command.", vbCritical, App.Title
                        Exit Sub
                    End If
                Else
                    Set Rsop01_operator_01 = Nothing
                    Set cdbcn = Nothing
                    MsgBox "Select record to edit.", vbInformation, App.Title
                    Exit Sub
                End If
            End With
        Set Rsop01_operator_01 = Nothing
        Set cdbcn = Nothing
        Exit Sub
ErrHandler:
    Set Rsop01_operator_01 = Nothing
    Set cdbcn = Nothing
    If err.Number <> 0 Then prompt_errlog err, frmOperatorView.Name, "exec_proc_usp_sel_op01_operator_01"
End Sub
Private Sub grdOperator_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then exec_proc_usp_sel_op01_operator_01
End Sub
