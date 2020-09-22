VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3F666E72-7F79-447A-BCFF-C42C44E3DBE7}#1.0#0"; "iGrid251_75B4A91C.ocx"
Begin VB.Form frmOutputView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Output Per Operator"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOutputView.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8775
   ScaleWidth      =   11910
   Begin iGrid251_75B4A91C.iGrid grdOutput 
      Height          =   7095
      Left            =   0
      TabIndex        =   18
      Top             =   720
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   12515
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   11895
      Begin VB.ComboBox cboOuput 
         Height          =   360
         ItemData        =   "frmOutputView.frx":127A
         Left            =   7440
         List            =   "frmOutputView.frx":1281
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   180
         Width           =   3015
      End
      Begin VB.OptionButton optOutput 
         Caption         =   "Filter"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optOutput 
         Caption         =   "View All"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpOutput 
         Height          =   360
         Index           =   0
         Left            =   3120
         TabIndex        =   2
         Top             =   180
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   635
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
         Format          =   16449537
         CurrentDate     =   39392
      End
      Begin MSComCtl2.DTPicker dtpOutput 
         Height          =   360
         Index           =   1
         Left            =   5040
         TabIndex        =   3
         Top             =   180
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   635
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
         Format          =   16449537
         CurrentDate     =   39392
      End
      Begin LVbuttons.LaVolpeButton cmdRun 
         Height          =   375
         Index           =   1
         Left            =   10560
         TabIndex        =   5
         Top             =   180
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   6
         TX              =   "&View"
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
         MICON           =   "frmOutputView.frx":128A
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
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Machine"
         Height          =   270
         Index           =   1
         Left            =   6600
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   14
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
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
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   8160
      Width           =   11895
      Begin LVbuttons.LaVolpeButton cmdNew 
         Height          =   405
         Left            =   120
         TabIndex        =   6
         Top             =   75
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
         MICON           =   "frmOutputView.frx":12A6
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
         TabIndex        =   7
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
         MICON           =   "frmOutputView.frx":12C2
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
         Left            =   6240
         TabIndex        =   9
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
         MICON           =   "frmOutputView.frx":12DE
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
         Left            =   4200
         TabIndex        =   8
         Top             =   75
         Width           =   1920
         _ExtentX        =   3387
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
         MICON           =   "frmOutputView.frx":12FA
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
      Begin VB.Label Label7 
         Caption         =   "Label1"
         Height          =   375
         Left            =   7320
         TabIndex        =   17
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Label2"
         Height          =   375
         Left            =   9000
         TabIndex        =   16
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.Label lblOutput 
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
      TabIndex        =   11
      Top             =   7830
      Width           =   11895
   End
End
Attribute VB_Name = "frmOutputView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cdbcn As clsDBAccess
'Dim cdbcn As cDataAccess.clsDBAccess
Public Sub move_menu_form()
    Me.Left = frmMenu.Width
    Me.Top = 0
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdNew_Click()
    frmOutput.Show 1
End Sub
Private Sub cmdPreview_Click()
    frmReportCR1.Show
End Sub

Private Sub cmdSearchEmployee_Click()
    bUserNew = False
    bOutputOperator = True
    frmSearchOperator.Show 1
End Sub
Private Sub cmdRun_Click(Index As Integer)
    If optOutput(1).Value Then
        exec_proc_usp_sel_so01_so_dtl_03
    Else
        Exit Sub
    End If
End Sub

Private Sub cmdView_Click()
    exec_proc_usp_sel_so01_so_dtl_02
End Sub
Private Sub dtpOutput_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub
Private Sub Form_Load()
Dim lCol As Long
On Error GoTo ErrHandler
    move_menu_form
    dtpOutput(0).Value = Date
    SetDates
    dtpOutput(0).Value = Label7.Caption
    With grdOutput
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
            .Size = 8
        End With
        With .AddCol(sKey:="so01_id_dtl", sHeader:="ID", lWidth:=50, bvisible:=False)
            .eTextFlags = igTextCenter
        End With
        With .AddCol(sKey:="so01_ref", sHeader:="SO", lWidth:=50, bvisible:=True)
            .eTextFlags = igTextLeft
        End With
        With .AddCol(sKey:="jo01_ref", sHeader:="JO", lWidth:=50, bvisible:=True)
            .eTextFlags = igTextLeft
        End With
        With .AddCol(sKey:="operator", sHeader:="Operator", lWidth:=50)
            .eTextFlags = igTextLeft
        End With
        With .AddCol(sKey:="so01_shift", sHeader:="Shift", lWidth:=50)
            .eTextFlags = igTextLeft
        End With
        With .AddCol(sKey:="mc01_machine", sHeader:="Machine", lWidth:=50)
            .eTextFlags = igTextRight
            .sFmtString = "###,##0.00"
        End With
        With .AddCol(sKey:="cc01_capacity", sHeader:="Machine Capacity", lWidth:=50)
            .eTextFlags = igTextRight
            .sFmtString = "###,##0.00"
        End With
        With .AddCol(sKey:="so01_date_product", sHeader:="Production Date", lWidth:=50)
            .eTextFlags = igTextRight
            .sFmtString = "MM/DD/YYYY"
        End With
        With .AddCol(sKey:="mc01_setting_time", sHeader:="Pre-Setting Time(Hr)", lWidth:=50)
            .eTextFlags = igTextRight
            .sFmtString = "###,##0.00"
        End With
        With .AddCol(sKey:="so01_set_time", sHeader:="Actual Setting Time(Hr)", lWidth:=50)
            .eTextFlags = igTextRight
            .sFmtString = "###,##0.00"
        End With
        With .AddCol(sKey:="so01_net_setting_time", sHeader:="Net Setting Time(Hr)", lWidth:=50)
            .eTextFlags = igTextRight
            .sFmtString = "###,##0.00"
        End With
        With .AddCol(sKey:="so01_running_time", sHeader:="Running Time(Hr)", lWidth:=50)
            .eTextFlags = igTextRight
            .sFmtString = "###,##0.00"
        End With
        With .AddCol(sKey:="so01_downtime", sHeader:="Downtime(Hr)", lWidth:=50)
            .eTextFlags = igTextRight
            .sFmtString = "###,##0.00"
        End With
        With .AddCol(sKey:="so01_quantity_produce", sHeader:="Actual Output('000)", lWidth:=50)
            .eTextFlags = igTextRight
            .sFmtString = "###,##0.00"
        End With
        With .AddCol(sKey:="so01_preset_quantity", sHeader:="Preset Output('000)", lWidth:=50)
            .eTextFlags = igTextRight
            .sFmtString = "###,##0.00"
        End With
        With .AddCol(sKey:="variance", sHeader:="Variance", lWidth:=50)
            .eTextFlags = igTextRight
            .sFmtString = "###,##0.00"
        End With
        With .AddCol(sKey:="percentage", sHeader:="Percentage", lWidth:=50)
            .eTextFlags = igTextRight
        End With
        
        For lCol = 1 To .ColCount
            .AutoWidthCol lCol
        Next lCol
        '.RowCount = 1
        .Editable = False
        .Redraw = True
        .AddCol bRowTextCol:=True
    End With
    exec_proc_usp_sel_so01_so_dtl_01
    exec_proc_usp_sel_mc01_machinery_02
Exit Sub
ErrHandler:
    If err.Number <> 0 Then prompt_errlog err, frmOutputView.Name, "Form_Load Event"
End Sub
Private Sub grdOutput_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
     bDoDefault = False
End Sub
Private Sub grdOutput_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)
    lblOutput.Caption = "Record Number " & lRow & " of " & grdOutput.RowCount

End Sub
Private Sub grdOutput_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)
    exec_proc_usp_sel_so01_so_dtl_02
End Sub
Private Sub grdOutput_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then exec_proc_usp_sel_so01_so_dtl_02
End Sub
Private Sub SetDates()
Dim sDayFirst As String, sDayLast As String, sMonth As String, sYear As String
    sDayLast = GetLastDayOfMonth(dtpOutput(0).Value)
    sDayFirst = "01"
    sYear = Year(dtpOutput(0).Value)
    sMonth = Month(dtpOutput(0).Value)
    Label8.Caption = sMonth & "/" & sDayLast & "/" & sYear
    Label7.Caption = sMonth & "/" & sDayFirst & "/" & sYear
    dtpOutput(0).Value = Label7.Caption
    dtpOutput(1).Value = Label8.Caption
End Sub
Private Function GetLastDayOfMonth(pDate As Date)
    GetLastDayOfMonth = Day(DateSerial(Year(pDate), Month(pDate) + 1, 0))
End Function
Public Sub exec_proc_usp_sel_so01_so_dtl_01()
Dim lCol As Long
Dim Rsso01_so_dtl_01 As ADODB.Recordset
On Error GoTo ErrHandler
    Set cdbcn = New clsDBAccess
    'Set cdbcn = New cDataAccess.clsDBAccess
    cdbcn.DataSource = strDatabase
    If Not cdbcn.OpenConnection = True Then GoTo ErrHandler
        Set Rsso01_so_dtl_01 = New ADODB.Recordset
        If cdbcn.proc_usp_sel_so01_so_dtl_01(Rsso01_so_dtl_01) = 0 Then
            If Not (Rsso01_so_dtl_01.BOF Or Rsso01_so_dtl_01.EOF) Then
                With grdOutput
                    .Redraw = False
                    .FillFromRS Rsso01_so_dtl_01
                    .Redraw = True
                    .SetCurCell 1, 4
                    For lCol = 1 To .ColCount
                        .AutoWidthCol lCol
                    Next lCol
                End With
            Else
                With grdOutput
                    .Redraw = False
                    .Clear
                    'pGrandTotals
                    .RowCount = 1
                    .Redraw = True
                    .SetCurCell 1, 4
                    For lCol = 1 To .ColCount
                        .AutoWidthCol lCol
                    Next lCol
                End With
            End If
        Else
            Set Rsso01_so_dtl_01 = Nothing
            Set cdbcn = Nothing
            MsgBox "There is an error executing the command.", vbCritical, App.Title
            Exit Sub
        End If
    Set Rsso01_so_dtl_01 = Nothing
    Set cdbcn = Nothing
Exit Sub
ErrHandler:
    Set cdbcn = Nothing
    If err.Number <> 0 Then prompt_errlog err, frmOutputView.Name, "proc_usp_sel_so01_so_dtl_01"
End Sub
Private Sub exec_proc_usp_sel_so01_so_dtl_02()
Dim Rsso01_so_dtl_02 As ADODB.Recordset
On Error GoTo ErrHandler
    Set cdbcn = New clsDBAccess
    'Set cdbcn = New cDataAccess.clsDBAccess
    cdbcn.DataSource = strDatabase
    If Not cdbcn.OpenConnection = True Then GoTo ErrHandler
    Set Rsso01_so_dtl_02 = New ADODB.Recordset
        With grdOutput
            If .CurRow <> 0 And .ColCount <> 0 Then
                If cdbcn.proc_usp_sel_so01_so_dtl_02(.CellValue(.CurRow, "so01_id_dtl"), .CellValue(.CurRow, "so01_ref"), Rsso01_so_dtl_02) = 0 Then
                    If Not (Rsso01_so_dtl_02.BOF Or Rsso01_so_dtl_02.EOF) Then
                        frmOutput.lblOutput(0).Caption = Trim$(Rsso01_so_dtl_02!so01_id_dtl)
                        frmOutput.txtOutput(0).Text = Trim$(Rsso01_so_dtl_02!so01_ref)
                        frmOutput.txtOutput(1).Text = Trim$(Rsso01_so_dtl_02!jo01_ref)
                        frmOutput.txtOutput(2).Text = Trim$(Rsso01_so_dtl_02!hr01_emplyee_id)
                        frmOutput.lblOutput(1).Caption = Trim$(Rsso01_so_dtl_02!operator)
                        frmOutput.cboOuput(0).Text = Trim$(Rsso01_so_dtl_02!mc01_machine)
                        frmOutput.cboOuput(1).Text = Trim$(Rsso01_so_dtl_02!cc01_capacity)
                        frmOutput.cboOuput(2).Text = Trim$(Rsso01_so_dtl_02!so01_shift)
                        frmOutput.txtOutput(3).Text = Format$(Rsso01_so_dtl_02!so01_set_time, "###,##0.00")
                        frmOutput.dtpOutput.Value = Rsso01_so_dtl_02!so01_date_product
                        frmOutput.txtOutput(4).Text = Format$(Rsso01_so_dtl_02!so01_running_time, "###,##0.00")
                        frmOutput.txtOutput(5).Text = Format$(Rsso01_so_dtl_02!so01_quantity_produce, "###,##0.00")
                        frmOutput.txtOutput(6).Text = Format$(Rsso01_so_dtl_02!so01_downtime, "###,##0.00")
                        frmOutput.txtOutput(7).Text = Trim$(Rsso01_so_dtl_02!mc01_reason_downtime)
                        frmOutput.Show 1
                    Else
                        Set Rsso01_so_dtl_02 = Nothing
                        Set cdbcn = Nothing
                        MsgBox "Record not available for editing.", vbCritical, App.Title
                        Exit Sub
                    End If
                Else
                    Set Rsso01_so_dtl_02 = Nothing
                    Set cdbcn = Nothing
                    MsgBox "There is an error executing the command.", vbCritical, App.Title
                    Exit Sub
                End If
            Else
                Set Rsso01_so_dtl_02 = Nothing
                Set cdbcn = Nothing
                MsgBox "Select an item in a grid.", vbCritical, App.Title
                Exit Sub
            End If
        End With
    Set Rsso01_so_dtl_02 = Nothing
    Set cdbcn = Nothing
Exit Sub
ErrHandler:
    Set cdbcn = Nothing
    If err.Number <> 0 Then prompt_errlog err, frmOutputView.Name, "proc_usp_sel_so01_so_dtl_02"
End Sub
Private Sub Form_Resize()
On Error Resume Next
    Frame1.Move 0, Frame1.Top, Me.ScaleWidth
    With lblOutput
       .Move 0, Me.ScaleHeight - 1000, Me.ScaleWidth
    End With
    With picMachinery
          .Move 0, Me.ScaleHeight - .Height - 120, Me.ScaleWidth
    End With
    
    With grdOutput
          .Move 0, .Top, Me.ScaleWidth, picMachinery.Top - .Top - 320
    End With

End Sub

Private Sub optOutput_Click(Index As Integer)
    If optOutput(0) Then
        exec_proc_usp_sel_so01_so_dtl_01
    Else
        grdOutput.Clear
    End If
End Sub
Public Sub exec_proc_usp_sel_so01_so_dtl_03()
Dim lCol As Long
Dim Rsso01_so_dtl_03 As ADODB.Recordset
On Error GoTo ErrHandler
    Set cdbcn = New clsDBAccess
    'Set cdbcn = New cDataAccess.clsDBAccess
    cdbcn.DataSource = strDatabase
    If Not cdbcn.OpenConnection = True Then GoTo ErrHandler
    Set Rsso01_so_dtl_03 = New ADODB.Recordset
        If cdbcn.proc_usp_sel_so01_so_dtl_03(CInt(Abs(optOutput(1).Value)), dtpOutput(0).Value, _
            dtpOutput(1).Value, cboOuput.Text, Rsso01_so_dtl_03) = 0 Then
            If Not (Rsso01_so_dtl_03.BOF Or Rsso01_so_dtl_03.EOF) Then
                With grdOutput
                    .Redraw = False
                    .Clear
                    .FillFromRS Rsso01_so_dtl_03
                    'pGrandTotals
                    .Redraw = True
                    .SetCurCell 1, 4
                    For lCol = 1 To .ColCount
                        .AutoWidthCol lCol
                    Next lCol
                End With
            Else
                With grdOutput
                    .Redraw = False
                    .Clear
                    'pGrandTotals
                    .RowCount = 1
                    .Redraw = True
                    .SetCurCell 1, 4
                    For lCol = 1 To .ColCount
                        .AutoWidthCol lCol
                    Next lCol
                End With
            End If
        Else
            Set Rsso01_so_dtl_03 = Nothing
            Set cdbcn = Nothing
            MsgBox "There is an error executing the command.", vbCritical, App.Title
            Exit Sub
        End If
    Set Rsso01_so_dtl_03 = Nothing
    Set cdbcn = Nothing
Exit Sub
ErrHandler:
    Set cdbcn = Nothing
    If err.Number <> 0 Then prompt_errlog err, frmOutputView.Name, "proc_usp_sel_so01_so_dtl_03"
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
                If Not IsNull(Rsmc01_machinery_02!mc01_machine) Then cboOuput.AddItem Rsmc01_machinery_02!mc01_machine
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
