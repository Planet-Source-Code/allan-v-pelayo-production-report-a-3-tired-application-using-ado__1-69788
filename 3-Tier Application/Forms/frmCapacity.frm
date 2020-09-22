VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmCapacity 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Capacity"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5505
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCapacity.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   5505
   StartUpPosition =   2  'CenterScreen
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
      Height          =   1695
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   5295
      Begin VB.TextBox txtCapacity 
         Height          =   315
         Index           =   2
         Left            =   1320
         TabIndex        =   2
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtCapacity 
         Height          =   315
         Index           =   1
         Left            =   1320
         TabIndex        =   1
         Top             =   960
         Width           =   3855
      End
      Begin VB.TextBox txtCapacity 
         Height          =   315
         Index           =   0
         Left            =   1320
         TabIndex        =   0
         Top             =   600
         Width           =   3855
      End
      Begin VB.Line Line4 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         X1              =   0
         X2              =   5280
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         X1              =   5280
         X2              =   5280
         Y1              =   120
         Y2              =   1680
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Wastage"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   1380
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Description"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   1020
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Capacity"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label lblCapacity 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1320
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Capacity ID"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Width           =   1095
      End
   End
   Begin LVbuttons.LaVolpeButton cmdSave 
      Height          =   410
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   1575
      _ExtentX        =   2778
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
      MICON           =   "frmCapacity.frx":127A
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
      Height          =   410
      Left            =   1800
      TabIndex        =   4
      Top             =   1920
      Width           =   1575
      _ExtentX        =   2778
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
      MICON           =   "frmCapacity.frx":1296
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
      Left            =   3480
      TabIndex        =   5
      Top             =   1920
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
      MICON           =   "frmCapacity.frx":12B2
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
      X2              =   5400
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      X1              =   120
      X2              =   5400
      Y1              =   1800
      Y2              =   1800
   End
End
Attribute VB_Name = "frmCapacity"
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

Private Sub chkCapacity_KeyPress(KeyAscii As Integer)
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
    For i = 0 To 2
        txtCapacity(i).Text = vbNullString
    Next i
    lblCapacity.Caption = vbNullString
    txtCapacity(0).Enabled = True
    txtCapacity(0).SetFocus
End Sub

Private Sub cmdNew_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub cmdSave_Click()
Dim i As Integer
On Error GoTo ErrHandler
    For i = 0 To 2
        If bEmpty(txtCapacity(i)) = True Then Exit Sub
    Next i
        Set cdbcn = New clsDBAccess
        'Set cdbcn = New cDataAccess.clsDBAccess
        cdbcn.DataSource = strDatabase
        If Not cdbcn.OpenConnection = True Then GoTo ErrHandler
            If cdbcn.proc_usp_insert_update_cc01_capacity(lblCapacity.Caption, Trim$(txtCapacity(0).Text), Trim$(txtCapacity(1).Text), Trim$(txtCapacity(2).Text), _
                 strUser, strUser) = 0 Then
                frmCapacityView.exec_proc_usp_cc01_capacity_02
                lblCapacity.Caption = cdbcn.lcc01_id_new
                Set cdbcn = Nothing
                MsgBox "Record successfully saved.", vbInformation, App.Title
                Exit Sub
            Else
                Set cdbcn = Nothing
                MsgBox "An error occured while trying to save records.", vbCritical, App.Title
                Exit Sub
            End If
            Set cdbcn = Nothing
Exit Sub
ErrHandler:
    Set cdbcn = Nothing
    If err.Number <> 0 Then prompt_errlog err, frmCapacity.Name, "cmdSave_Click"
End Sub

Private Sub cmdSave_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtCapacity_GotFocus(Index As Integer)
Dim i As Integer
    For i = 0 To 2
        SelText txtCapacity(i)
    Next i
End Sub




Private Sub txtCapacity_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtCapacity_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub
