Attribute VB_Name = "modFunctions"
Option Explicit
Dim i As Integer
Dim strs As String
Public Sub SelText(ByRef strControlName As Control)
On Error Resume Next
    strControlName.SelStart = 0
    strControlName.SelLength = Len(strControlName.Text)
End Sub
Public Sub prompt_errlog(ByVal sError As ErrObject, ByVal ModuleName As String, ByVal OccurIn As String)
On Error Resume Next
    MsgBox "Error From: " & ModuleName & vbNewLine & _
           "Occur In: " & OccurIn & vbNewLine & _
           "Error Number: " & sError.Number & vbNewLine & _
           "Description: " & sError.Description, vbCritical, "Application Error"
    Open App.Path & "\Error.log" For Append As #1
        Print #1, "Error Date Occured:"; Format(Date, "MMM-dd-yyyy") & "  Time: " & Time & "  " & " in Module: " & ModuleName & " " & "Error Number:  " & " " & sError.Number & " " & "Error Description: " & " " & sError.Description & ":" & OccurIn
    Close #1
End Sub
Public Function bEmpty(ByRef sText As Variant) As Boolean
On Error Resume Next
    If sText = vbNullString Then
        bEmpty = True
        MsgBox "Required field is left empty.", vbExclamation, App.Title
        sText.SetFocus: SendKeys "{Home}+{End}"
    Else
        bEmpty = False
    End If
End Function
Public Function CenterForm(ByRef frm As Form)
    frm.Move (Screen.Width - frm.Width) / 2, (Screen.Height - frm.Height) / 2
End Function
Public Function msgWho(ByVal sCreate As String, ByVal sCreateDate As Date, ByVal sModified As String, ByVal sDateModified As String, ByRef iRecordCount As Integer) As Boolean
On Error Resume Next
    If iRecordCount > 0 Then
        msgWho = True
        MsgBox "Created By        : " & sCreate & vbCrLf & "Dated Created  : " & Format(sCreateDate, "dddd MMMM dd,yyyy h:mm:ss AMPM") & vbCrLf & "Modified By        : " & sModified & vbCrLf & "Date Modified    : " & Format(sDateModified, "dddd MMMM dd,yyyy h:mm:ss AMPM"), vbInformation, "Created and Modified"
    End If
End Function
Public Sub move_menu_form(frm As Form)
    frm.Left = frmMenu.Width
    frm.Top = 0
End Sub
Public Sub TextFormat(ByRef strControlName As Control)
    On Error Resume Next
    strControlName.Text = Format(strControlName.Text, "###,##0.00")
End Sub
