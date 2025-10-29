Attribute VB_Name = "Module1"
Public Sub ShowPlaceholderPicker()
    frmPlaceholders.Show
End Sub

Public Sub InsertPlaceholder(ByVal placeholderText As String)
    Dim toInsert As String
    toInsert = Replace(placeholderText, " ? ", vbCrLf)
    Selection.TypeText Text:=toInsert
End Sub


