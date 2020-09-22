VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Object = "{3F666E72-7F79-447A-BCFF-C42C44E3DBE7}#1.0#0"; "iGrid251_75B4A91C.ocx"
Begin VB.Form frmMachineryView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Machinery"
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
   Icon            =   "frmMachineryView.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8775
   ScaleWidth      =   11910
   Begin iGrid251_75B4A91C.iGrid grdMachinery 
      Height          =   7815
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   13785
   End
   Begin VB.PictureBox picMachinery 
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
         MICON           =   "frmMachineryView.frx":127A
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
         Top             =   80
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
         MICON           =   "frmMachineryView.frx":1296
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
         Top             =   80
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
         MICON           =   "frmMachineryView.frx":12B2
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
   Begin VB.Label lblMachinery 
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
Attribute VB_Name = "frmMachineryView"
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
    frmMachinery.Caption = "New"
    frmMachinery.Show 1
End Sub
Private Sub cmdView_Click()
    exec_proc_usp_sel_mc01_machinery_01
End Sub
Private Sub Form_Load()
On Error GoTo ErrHandler
    move_menu_form
    With grdMachinery
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
        With .AddCol(sKey:="mc01_id", sHeader:="ID", lWidth:=50, bvisible:=True)
            .eTextFlags = igTextCenter
        End With
        With .AddCol(sKey:="mc01_machine", sHeader:="Machine", lWidth:=50)
            .eTextFlags = igTextLeft
        End With
        With .AddCol(sKey:="mc01_machine_desc", sHeader:="Description", lWidth:=50)
            .iRightIndent = 2
        End With
        With .AddCol(sKey:="mc01_serial_no", sHeader:="Serial Number", lWidth:=50)
            .iRightIndent = 2
        End With
        With .AddCol(sKey:="mc01_machine_model", sHeader:="Model", lWidth:=50)
            .iRightIndent = 2
        End With
        With .AddCol(sKey:="mc01_date_acquired", sHeader:="Date Acquired", lWidth:=50)
            .iRightIndent = 2
        End With
        With .AddCol(sKey:="mc01_acquisition_cost", sHeader:="Acquisition Cost", lWidth:=50)
            .iRightIndent = 2
        End With
        For lCol = 1 To .ColCount
            .AutoWidthCol lCol
        Next lCol
        .Editable = False
        .Redraw = True
    End With
    exec_proc_usp_sel_mc01_machinery_02
Exit Sub
ErrHandler:
    If err.Number <> 0 Then prompt_errlog err, frmMachineryView.Name, "Form_Load Event"
End Sub
Private Sub Form_Resize()
On Error Resume Next
   With picMachinery
    .Move 0, Me.ScaleHeight - .Height - 120, Me.ScaleWidth
   End With
   With lblMachinery
      .Move 0, Me.ScaleHeight - .Height - picMachinery.Height - 120, Me.ScaleWidth
   End With
   With grdMachinery
      .Move 0, .Top, Me.ScaleWidth, lblMachinery.Top - .Top - 60
   End With
End Sub

Private Sub grdMachinery_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
    bDoDefault = False
End Sub
Public Sub exec_proc_usp_sel_mc01_machinery_02()
Dim Rsmc01_machinery_02 As ADODB.Recordset
On Error GoTo ErrHandler
    Set cdbcn = New clsDBAccess
    'Set cdbcn = New cDataAccess.clsDBAccess
    cdbcn.DataSource = strDatabase
    If Not cdbcn.OpenConnection = True Then GoTo ErrHandler
    Set Rsmc01_machinery_02 = New ADODB.Recordset
        If cdbcn.proc_usp_sel_mc01_machinery_02(Rsmc01_machinery_02) = 0 Then
            If Not (Rsmc01_machinery_02.BOF Or Rsmc01_machinery_02.EOF) Then
                With grdMachinery
                    .Redraw = False
                    .FillFromRS Rsmc01_machinery_02
                    For lCol = 1 To .ColCount
                        .AutoWidthCol lCol
                    Next lCol
                    .Redraw = True
                    .SetCurCell 1, 2
                End With
            Else
                Set Rsmc01_machinery_02 = Nothing
                Set cdbcn = Nothing
                MsgBox "No record found.", vbInformation, App.Title
                Exit Sub
            End If
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
    If err.Number <> 0 Then prompt_errlog err, frmMachineryView.Name, "proc_usp_sel_mc01_machinery_02"
End Sub
Private Sub grdMachinery_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)
    lblMachinery.Caption = "Record Number " & lRow & " of " & grdMachinery.RowCount
End Sub
Private Sub grdMachinery_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)
    exec_proc_usp_sel_mc01_machinery_01
End Sub
Private Sub exec_proc_usp_sel_mc01_machinery_01()
Dim Rsmc01_machinery_01 As ADODB.Recordset
On Error GoTo ErrHandler
    Set cdbcn = New clsDBAccess
    'Set cdbcn = New cDataAccess.clsDBAccess
    cdbcn.DataSource = strDatabase
    If Not cdbcn.OpenConnection = True Then GoTo ErrHandler
    Set Rsmc01_machinery_01 = New ADODB.Recordset
            With grdMachinery
                If .CurRow <> 0 And .ColCount <> 0 Then
                    If cdbcn.proc_usp_sel_mc01_machinery_01(.CellValue(.CurRow, "mc01_id"), Rsmc01_machinery_01) = 0 Then
                        If Not (Rsmc01_machinery_01.BOF Or Rsmc01_machinery_01.EOF) Then
                            frmMachinery.lblMachinery.Caption = Trim$(Rsmc01_machinery_01!mc01_id)
                            frmMachinery.txtMachinery(0).Text = Trim$(Rsmc01_machinery_01!mc01_machine)
                            frmMachinery.txtMachinery(1).Text = Trim$(Rsmc01_machinery_01!mc01_machine_desc)
                            frmMachinery.txtMachinery(2).Text = Trim$(Rsmc01_machinery_01!mc01_serial_no)
                            frmMachinery.txtMachinery(3).Text = Trim$(Rsmc01_machinery_01!mc01_machine_model)
                            frmMachinery.dtpMachinery.Value = Trim$(Rsmc01_machinery_01!mc01_date_acquired)
                            frmMachinery.txtMachinery(4).Text = Trim$(Rsmc01_machinery_01!mc01_acquisition_cost)
                            frmMachinery.txtMachinery(0).Enabled = False
                            frmMachinery.Caption = "Machinery:" & Trim$(Rsmc01_machinery_01!mc01_id)
                            frmMachinery.Show 1
                        Else
                            Set Rsmc01_machinery_01 = Nothing
                            Set cdbcn = Nothing
                            MsgBox "No records available.", vbInformation, App.Title
                            Exit Sub
                        End If
                    Else
                        Set Rsmc01_machinery_01 = Nothing
                        Set cdbcn = Nothing
                        MsgBox "There is an error executing the command.", vbCritical, App.Title
                        Exit Sub
                    End If
                Else
                    Set Rsmc01_machinery_01 = Nothing
                    Set cdbcn = Nothing
                    MsgBox "Select record to edit.", vbInformation, App.Title
                    Exit Sub
                End If
            End With
        Set Rsmc01_machinery_01 = Nothing
        Set cdbcn = Nothing
        Exit Sub
ErrHandler:
    Set cdbcn = Nothing
    If err.Number <> 0 Then prompt_errlog err, frmMachineryView.Name, "exec_proc_usp_sel_mc01_machinery_01"
End Sub
Private Sub grdMachinery_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then exec_proc_usp_sel_mc01_machinery_01
End Sub


