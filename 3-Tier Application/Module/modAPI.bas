Attribute VB_Name = "modAPI"
Option Explicit
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function CreateFieldDefFile Lib "p2smon.dll" (lpUnk As Object, ByVal filename As String, ByVal bOverWriteExistingFile As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long

