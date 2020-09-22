VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMenu 
   BorderStyle     =   0  'None
   Caption         =   " Menu"
   ClientHeight    =   9255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3135
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   3135
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   120
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":0B4A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3360
      Top             =   500
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
            Picture         =   "frmMenu.frx":1694
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1A2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1DD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":2176
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":251A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":28BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":2C62
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":2FFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":339A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":3736
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView trvMenu 
      Height          =   8715
      Left            =   75
      TabIndex        =   1
      Top             =   360
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   15372
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Menu Explorer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      BorderWidth     =   5
      Height          =   9255
      Left            =   0
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cdbcn As clsDBAccess
'Dim cdbcn As cDataAccess.clsDBAccess
Public Sub move_menu_form()
    Me.Left = 0
    Me.Top = 0
End Sub
Private Sub Form_Load()
Dim Rsus01_user_02 As ADODB.Recordset
On Error GoTo ErrHandler
    move_menu_form
Exit Sub
ErrHandler:
    If err.Number <> 0 Then prompt_errlog err, frmLogin.Name, "cmdLogin_Click Event"
    Screen.MousePointer = vbDefault
End Sub

Private Sub trvMenu_DblClick()
On Error Resume Next
    If trvMenu.SelectedItem.Text = "User" Then
        frmUserView.Show
    End If
    If trvMenu.SelectedItem.Text = "Operator" Then
        frmOperatorView.Show
    End If
    If trvMenu.SelectedItem.Text = "Machinery" Then
        frmMachineryView.Show
    End If
    If trvMenu.SelectedItem.Text = "Capacity" Then
        frmCapacityView.Show
    End If
    If trvMenu.SelectedItem.Text = "Machine Capacity" Then
        frmMachineCapacity.Show 1
    End If
    If trvMenu.SelectedItem.Text = "Output" Then
        frmOutputView.Show
    End If
    If trvMenu.SelectedItem.Text = "Production Output" Then
        frmReportParamCR2.Show
    End If
End Sub

Private Sub trvMenu_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then
        If trvMenu.SelectedItem.Text = "User" Then
            frmUserView.Show
        End If
        If trvMenu.SelectedItem.Text = "Operator" Then
            frmOperatorView.Show
        End If
        If trvMenu.SelectedItem.Text = "Machinery" Then
            frmMachineryView.Show
        End If
        If trvMenu.SelectedItem.Text = "Capacity" Then
            frmCapacityView.Show
        End If
        If trvMenu.SelectedItem.Text = "Machine Capacity" Then
            frmMachineCapacity.Show 1
        End If
        If trvMenu.SelectedItem.Text = "Output" Then
            frmOutputView.Show
        End If
        If trvMenu.SelectedItem.Text = "Production Output" Then
            frmReportParamCR2.Show
        End If
    End If
End Sub
