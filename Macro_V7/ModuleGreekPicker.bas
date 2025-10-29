Attribute VB_Name = "ModuleGreekPicker"
Option Explicit

' Entry point you can bind to a keyboard shortcut (Alt+F8 ? Macros ? Options)
Public Sub InsertGreek_ShowPicker()
    On Error Resume Next
    frmGreekPicker.Show
    If Err.Number <> 0 Then
        MsgBox "Form 'frmGreekPicker' not found. Create the UserForm as described.", vbExclamation
    End If
End Sub

