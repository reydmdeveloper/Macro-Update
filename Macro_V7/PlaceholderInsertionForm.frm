VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PlaceholderInsertionForm 
   Caption         =   "Place Holder"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   4560
   OleObjectBlob   =   "PlaceholderInsertionForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PlaceholderInsertionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub InsertButton_Click()
    Dim placeholderText As String
    placeholderText = PlaceholderDropdown.Value
    If placeholderText <> "" Then
        InsertPlaceholder placeholderText
        Unload Me
    Else
        MsgBox "Please select a placeholder.", vbExclamation
    End If
End Sub

Private Sub PlaceholderDropdown_Change()

End Sub


Private Sub UserForm_Initialize()
    With PlaceholderDropdown
        .AddItem "[LOGO: ThermoFisher" & vbCrLf & "SCIENTIFIC" & vbCrLf & "The world leader in serving science]"
        .AddItem "[SIGNATURE]"
        .AddItem "[SIGNATURE: [ILLEGIBLE]]"
        .AddItem "[SIGNATURE: XXX]"
        .AddItem "[LOGO]"
        .AddItem "[IMAGES]"
        .AddItem "[HW]"
        .AddItem "[EMBLEM]"
        .AddItem "[ILLEGIBLE]"
        .AddItem "[DIGITAL SIGNATURE]"
        .AddItem "[REDACTION]"
        .AddItem "[STAMP]"
        .AddItem "[STAMP: [SIGNATURE: [ILLEGIBLE]]]"
        .AddItem "[ICON]"
        .AddItem "[handwritten text to be added]"
        ' Add more placeholder options as needed
    End With
End Sub

