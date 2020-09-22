VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmReportCR2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report"
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
   Icon            =   "frmReportCR2.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8775
   ScaleWidth      =   11910
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   8160
      Width           =   11895
      Begin LVbuttons.LaVolpeButton cmdPrinSetting 
         Height          =   410
         Left            =   120
         TabIndex        =   1
         Top             =   80
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   714
         BTYPE           =   6
         TX              =   "Prin&ter Setting"
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
         MICON           =   "frmReportCR2.frx":127A
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
      Begin LVbuttons.LaVolpeButton cmdPrint 
         Height          =   410
         Left            =   2040
         TabIndex        =   2
         Top             =   80
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   714
         BTYPE           =   6
         TX              =   "&Print"
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
         MICON           =   "frmReportCR2.frx":1296
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
      Begin LVbuttons.LaVolpeButton cmdExport 
         Height          =   410
         Left            =   3960
         TabIndex        =   3
         Top             =   80
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   714
         BTYPE           =   6
         TX              =   "&Export"
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
         MICON           =   "frmReportCR2.frx":12B2
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
         Left            =   5760
         TabIndex        =   4
         Top             =   80
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
         MICON           =   "frmReportCR2.frx":12CE
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
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      Height          =   8085
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11925
      lastProp        =   500
      _cx             =   21034
      _cy             =   14261
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   0   'False
      EnablePrintButton=   0   'False
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   0   'False
      EnableSearchControl=   0   'False
      EnableRefreshButton=   0   'False
      EnableDrillDown =   0   'False
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "frmReportCR2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CR2 As New CrystalReport2
'Dim cdbcn As cDataAccess.clsDBAccess
Dim cdbcn As clsDBAccess
Option Explicit
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Screen.MousePointer = vbHourglass
    move_menu_form
    exec_proc_usp_sel_so01_so_dtl_04
    proc_usp_sel_so01_so_dtl_04_heading
Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Resize()
On Error Resume Next
   With Frame1
    .Move 0, Me.ScaleHeight - .Height - 120, Me.ScaleWidth
   End With
   With CRViewer91
      .Move 0, .Top, Me.ScaleWidth, Frame1.Top - .Top - 60
   End With
End Sub
Private Sub exec_proc_usp_sel_so01_so_dtl_04()
Dim Rsso01_so_dtl_04 As ADODB.Recordset
On Error GoTo ErrHandler
    Set cdbcn = New clsDBAccess
    'Set cdbcn = New cDataAccess.clsDBAccess
    cdbcn.DataSource = strDatabase
    If Not cdbcn.OpenConnection = True Then GoTo ErrHandler
        Set Rsso01_so_dtl_04 = New ADODB.Recordset
        If cdbcn.proc_usp_sel_so01_so_dtl_04(frmReportParamCR2.dtpOutput(0).Value, frmReportParamCR2.dtpOutput(1).Value, _
            frmReportParamCR2.cboOuput(0).Text, frmReportParamCR2.cboOuput(1).Text, Rsso01_so_dtl_04) = 0 Then
            If Not (Rsso01_so_dtl_04.BOF Or Rsso01_so_dtl_04.EOF) Then
            CreateFieldDefFile Rsso01_so_dtl_04, App.Path & "\Report\so01_so_dtl_04.ttx", 1
            CR2.DiscardSavedData
            CR2.Database.SetDataSource Rsso01_so_dtl_04, 3, 1
            CRViewer91.ReportSource = CR2
            CRViewer91.ViewReport
            CRViewer91.Zoom (130)
            Set Rsso01_so_dtl_04 = Nothing
            Set cdbcn = Nothing
        Else
            Set Rsso01_so_dtl_04 = Nothing
            Set cdbcn = Nothing
            MsgBox "There is no data in this report. Supply different selection criteria"
        End If
    End If
    Set Rsso01_so_dtl_04 = Nothing
    Set cdbcn = Nothing
Exit Sub
ErrHandler:
    Set cdbcn = Nothing
    If err.Number <> 0 Then prompt_errlog err, frmReportCR2.Name, "exec_proc_usp_sel_so01_so_dtl_04"
End Sub
Private Sub move_menu_form()
    Me.Left = frmMenu.Width
    Me.Top = 0
End Sub
Private Sub cmdExport_Click()
On Error Resume Next
    CR2.Export True
End Sub
Private Sub cmdPrinSetting_Click()
On Error Resume Next
    CR2.PrinterSetup Me.hWnd
End Sub
Private Sub cmdPrint_Click()
On Error Resume Next
    CR2.PrintOut
End Sub
Private Sub proc_usp_sel_so01_so_dtl_04_heading()
        CR2.Text15.SetText "FOR THE PERIOD : " + CStr(frmReportParamCR2.dtpOutput(0).Value) + "-" + CStr(frmReportParamCR2.dtpOutput(1).Value)
        If frmReportParamCR2.cboOuput(0).ListIndex = -1 Then
            CR2.Text14.SetText "MACHINE : ALL"
        Else
            CR2.Text14.SetText "MACHINE : " + frmReportParamCR2.cboOuput(0).Text
        End If
        
        If frmReportParamCR2.cboOuput(1).ListIndex = -1 Then
            CR2.Text25.SetText "OPERATOR : ALL"
        Else
            CR2.Text25.SetText "OPERATOR : " + frmReportParamCR2.cboOuput(1).Text
        End If

End Sub
