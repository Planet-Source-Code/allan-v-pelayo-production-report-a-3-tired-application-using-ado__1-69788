VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmOutput 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Output Per Operator"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8400
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOutput.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   8400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   8175
      Begin VB.TextBox txtOutput 
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
         Index           =   7
         Left            =   2160
         TabIndex        =   13
         Top             =   2760
         Width           =   5895
      End
      Begin VB.TextBox txtOutput 
         Alignment       =   1  'Right Justify
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
         Index           =   6
         Left            =   6120
         TabIndex        =   12
         Top             =   2400
         Width           =   1935
      End
      Begin VB.ComboBox cboOuput 
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
         Index           =   2
         ItemData        =   "frmOutput.frx":127A
         Left            =   2160
         List            =   "frmOutput.frx":127C
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1680
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
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOutput.frx":127E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtOutput 
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
         Index           =   1
         Left            =   6120
         TabIndex        =   2
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtOutput 
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
         Index           =   0
         Left            =   2160
         TabIndex        =   1
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtOutput 
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
         Index           =   2
         Left            =   2160
         TabIndex        =   3
         Top             =   960
         Width           =   1815
      End
      Begin VB.ComboBox cboOuput 
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
         Index           =   0
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtOutput 
         Alignment       =   1  'Right Justify
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
         Index           =   4
         Left            =   6120
         TabIndex        =   10
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox txtOutput 
         Alignment       =   1  'Right Justify
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
         Index           =   5
         Left            =   2160
         TabIndex        =   11
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox txtOutput 
         Alignment       =   1  'Right Justify
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
         Index           =   3
         Left            =   6120
         TabIndex        =   8
         Top             =   1680
         Width           =   1935
      End
      Begin VB.ComboBox cboOuput 
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
         Index           =   1
         Left            =   6120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1320
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker dtpOutput 
         Height          =   360
         Left            =   2160
         TabIndex        =   9
         Top             =   2040
         Width           =   1815
         _ExtentX        =   3201
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
         Format          =   16384001
         CurrentDate     =   39392
      End
      Begin LVbuttons.LaVolpeButton cmdSearchEmployee 
         Height          =   360
         Left            =   4080
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   960
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
         BTYPE           =   7
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
         MICON           =   "frmOutput.frx":1818
         ALIGN           =   1
         IMGLST          =   "ImageList1"
         IMGICON         =   "1"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Reason for Downtime"
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
         Index           =   5
         Left            =   120
         TabIndex        =   32
         Top             =   2805
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Downtime"
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
         Index           =   4
         Left            =   4680
         TabIndex        =   31
         Top             =   2445
         Width           =   1335
      End
      Begin VB.Line Line4 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         X1              =   0
         X2              =   8160
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         X1              =   8160
         X2              =   8160
         Y1              =   120
         Y2              =   3240
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Shift"
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
         Index           =   2
         Left            =   120
         TabIndex        =   30
         Top             =   1725
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "JO Reference"
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
         Left            =   3840
         TabIndex        =   29
         Top             =   660
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "SO Reference"
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
         Left            =   120
         TabIndex        =   28
         Top             =   660
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Operator"
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
         TabIndex        =   27
         Top             =   1005
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Name"
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
         Left            =   4440
         TabIndex        =   26
         Top             =   1005
         Width           =   615
      End
      Begin VB.Label lblOutput 
         BackColor       =   &H8000000E&
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
         Height          =   360
         Index           =   1
         Left            =   5160
         TabIndex        =   4
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Label4 
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
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   1365
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Running Time"
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
         Index           =   2
         Left            =   3840
         TabIndex        =   23
         Top             =   2085
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Qty Produced"
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
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   2445
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Setting Time"
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
         Index           =   0
         Left            =   3840
         TabIndex        =   21
         Top             =   1740
         Width           =   2175
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Production Date"
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
         Left            =   120
         TabIndex        =   20
         Top             =   2100
         Width           =   1935
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Machine Capacity"
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
         Index           =   0
         Left            =   3600
         TabIndex        =   19
         Top             =   1365
         Width           =   2415
      End
      Begin VB.Label lblOutput 
         BackColor       =   &H8000000E&
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
         Height          =   360
         Index           =   0
         Left            =   2160
         TabIndex        =   0
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ID"
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
         Left            =   360
         TabIndex        =   18
         Top             =   300
         Width           =   1695
      End
   End
   Begin LVbuttons.LaVolpeButton cmdSave 
      Height          =   410
      Left            =   120
      TabIndex        =   14
      Top             =   3480
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
      MICON           =   "frmOutput.frx":1834
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
      TabIndex        =   15
      Top             =   3480
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
      MICON           =   "frmOutput.frx":1850
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
      TabIndex        =   16
      Top             =   3480
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
      MICON           =   "frmOutput.frx":186C
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
      X2              =   8280
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      X1              =   120
      X2              =   8280
      Y1              =   3360
      Y2              =   3360
   End
End
Attribute VB_Name = "frmOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cdbcn As clsDBAccess
'Dim cdbcn As cDataAccess.clsDBAccess
Private Sub ClearControl()
Dim i As Integer
    For i = 0 To 7
        txtOutput(i).Text = vbNullString
    Next i
    cboOuput(0).ListIndex = -1
    cboOuput(1).ListIndex = -1
    cboOuput(2).ListIndex = -1
    lblOutput(0).Caption = vbNullString
    lblOutput(1).Caption = vbNullString
    dtpOutput.Value = Date
    txtOutput(0).SetFocus
End Sub
Private Sub cboOuput_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyEscape Then Unload Me
End Sub
Private Sub cmdNew_Click()
    ClearControl
End Sub
Private Sub exec_proc_usp_sel_cc01_capacity_02()
Dim Rscc01_capacity_02 As ADODB.Recordset
On Error GoTo ErrHandler
    Set cdbcn = New clsDBAccess
    'Set cdbcn = New cDataAccess.clsDBAccess
    cdbcn.DataSource = strDatabase
    If Not cdbcn.OpenConnection = True Then GoTo ErrHandler
    Set Rscc01_capacity_02 = New ADODB.Recordset
        If cdbcn.proc_usp_sel_cc01_capacity_02(Rscc01_capacity_02) = 0 Then
            If Rscc01_capacity_02.RecordCount > 0 Then Rscc01_capacity_02.MoveFirst
                Do Until Rscc01_capacity_02.EOF
                If Not IsNull(Rscc01_capacity_02!cc01_capacity) Then cboOuput(1).AddItem Rscc01_capacity_02!cc01_capacity
                    Rscc01_capacity_02.MoveNext
                DoEvents
                Loop
        Else
            Set Rscc01_capacity_02 = Nothing
            Set cdbcn = Nothing
            MsgBox "There is an error executing the command.", vbCritical, App.Title
            Exit Sub
        End If
    Set cdbcn = Nothing
    Set Rscc01_capacity_02 = Nothing
Exit Sub
ErrHandler:
    Set cdbcn = Nothing
    If err.Number <> 0 Then prompt_errlog err, frmOutput.Name, "exec_proc_usp_sel_cc01_capacity_02"
End Sub

Private Sub cmdNew_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub cmdSave_Click()
Dim i As Integer
On Error GoTo ErrHandler
    For i = 3 To 7
        If txtOutput(i).Text = vbNullString Then
            txtOutput(i).Text = 0
        End If
    Next i
    For i = 0 To 7
        If bEmpty(txtOutput(i)) = True Then Exit Sub
    Next i
    Set cdbcn = New clsDBAccess
    'Set cdbcn = New cDataAccess.clsDBAccess
    cdbcn.DataSource = strDatabase

    If Not cdbcn.OpenConnection = True Then GoTo ErrHandler
        If cdbcn.proc_usp_insert_update_so01_so_dtl(lblOutput(0).Caption, Trim$(txtOutput(0).Text), Trim$(txtOutput(1).Text), _
            cboOuput(2).Text, Trim$(txtOutput(2).Text), cboOuput(0).Text, cboOuput(1).Text, dtpOutput.Value, Trim$(txtOutput(3).Text), _
            Trim$(txtOutput(5).Text), Trim$(txtOutput(4).Text), Trim$(txtOutput(6).Text), Trim$(txtOutput(7).Text), strUser, strUser) = 0 Then
            lblOutput(0).Caption = cdbcn.lso01_id_dtl_new
            frmOutputView.exec_proc_usp_sel_so01_so_dtl_01
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
    If err.Number <> 0 Then prompt_errlog err, frmOutput.Name, "cmdSave_Click"
End Sub

Private Sub cmdSave_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub cmdSearchEmployee_Click()
    bUserNew = False
    bOutputOperator = True
    frmSearchOperator.Show 1
End Sub

Private Sub dtpOutput_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyEscape Then Unload Me
     If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    dtpOutput.Value = Date
    exec_proc_usp_sel_cc01_capacity_02
    exec_proc_usp_sel_mc01_machinery_02
    exec_proc_usp_sel_sh01_shift_01
    FormatTextNumber
End Sub

Private Sub txtOutput_Change(Index As Integer)
    Select Case Index
        Case 2
            exec_proc_usp_sel_op01_operator_03
    End Select
End Sub

Private Sub txtOutput_GotFocus(Index As Integer)
Dim i As Integer
    For i = 0 To 7
        SelText txtOutput(i)
    Next i
End Sub

Private Sub txtOutput_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyEscape Then Unload Me
     Select Case Index
        Case 2
            If KeyCode = vbKeyF2 Then
                 bUserNew = False
                bOutputOperator = True
                frmSearchOperator.Show 1
            End If
    End Select
End Sub

Private Sub txtOutput_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{Tab}"
    Select Case Index
        Case 3, 4, 5, 6
            If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 13 And KeyAscii <> 46 Then
                KeyAscii = 0
            End If
    End Select
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
    Set cdbcn = Nothing
    If err.Number <> 0 Then prompt_errlog err, frmOutputView.Name, "proc_usp_sel_mc01_machinery_02"
End Sub
Private Sub exec_proc_usp_sel_op01_operator_03()
Dim Rsop01_operator_03 As ADODB.Recordset
On Error GoTo ErrHandler
    Set cdbcn = New clsDBAccess
    'Set cdbcn = New cDataAccess.clsDBAccess
    cdbcn.DataSource = strDatabase
    If Not cdbcn.OpenConnection = True Then GoTo ErrHandler
    Set Rsop01_operator_03 = New ADODB.Recordset
        If cdbcn.proc_usp_sel_op01_operator_03(Trim$(txtOutput(2).Text), Rsop01_operator_03) = 0 Then
            If Not (Rsop01_operator_03.BOF Or Rsop01_operator_03.EOF) Then
                    lblOutput(1).Caption = Trim$(Rsop01_operator_03!fullname)
            Else
                lblOutput(1).Caption = vbNullString
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
    Set cdbcn = Nothing
    If err.Number <> 0 Then prompt_errlog err, frmOutputView.Name, "proc_usp_sel_op01_operator_03"
End Sub
Private Sub exec_proc_usp_sel_sh01_shift_01()
Dim Rssh01_shift_01 As ADODB.Recordset
On Error GoTo ErrHandler
    Set cdbcn = New clsDBAccess
    'Set cdbcn = New cDataAccess.clsDBAccess
    cdbcn.DataSource = strDatabase
    If Not cdbcn.OpenConnection = True Then GoTo ErrHandler
    Set Rssh01_shift_01 = New ADODB.Recordset
        If cdbcn.proc_usp_sel_sh01_shift_01(Rssh01_shift_01) = 0 Then
            If Rssh01_shift_01.RecordCount > 0 Then Rssh01_shift_01.MoveFirst
                Do Until Rssh01_shift_01.EOF
                If Not IsNull(Rssh01_shift_01!sh01_shit) Then cboOuput(2).AddItem Rssh01_shift_01!sh01_shit
                    Rssh01_shift_01.MoveNext
                DoEvents
                Loop
        Else
            Set Rssh01_shift_01 = Nothing
            Set cdbcn = Nothing
            MsgBox "There is an error executing the command.", vbCritical, App.Title
            Exit Sub
        End If
    Set Rssh01_shift_01 = Nothing
    Set cdbcn = Nothing
Exit Sub
ErrHandler:
    Set Rssh01_shift_01 = Nothing
    Set cdbcn = Nothing
    If err.Number <> 0 Then prompt_errlog err, frmOutputView.Name, "proc_usp_sel_sh01_shift_01"
End Sub

Private Sub FormatTextNumber()
Dim i As Integer
    For i = 3 To 6
        TextFormat txtOutput(i)
    Next i
End Sub

Private Sub txtOutput_LostFocus(Index As Integer)
    Select Case Index
        Case 3, 4, 5, 6
            FormatTextNumber
    End Select
End Sub
