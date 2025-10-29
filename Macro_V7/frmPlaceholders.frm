VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPlaceholders 
   Caption         =   "Insert Placebolder Tags"
   ClientHeight    =   2940
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   5076
   OleObjectBlob   =   "frmPlaceholders.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPlaceholders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' Master list for all placeholders (unfiltered)
Private masterList As Variant

Private Sub UserForm_Initialize()
    ' Full list
    masterList = Array( _
        "[LOGO: ThermoFisher ? SCIENTIFIC ? The world leader in serving science]", _
        "[SIGNATURE]", _
        "[SIGNATURE: [ILLEGIBLE]]", _
        "[SIGNATURE: XXX]", _
        "[LOGO]", _
        "[IMAGES]", _
        "[HW]", _
        "[EMBLEM]", _
        "[ILLEGIBLE]", _
        "[DIGITAL SIGNATURE]", _
        "[REDACTION]", _
        "[STAMP]", _
        "[STAMP: [SIGNATURE: [ILLEGIBLE]]]", _
        "[ICON]", _
        "[handwritten text to be added]" _
    )

    KeepOpen.Value = False
    SearchBox.Text = ""

    PopulateList ""
    UpdateButtonsState

    SearchBox.SetFocus
End Sub

Private Sub PopulateList(ByVal filterText As String)
    Dim i As Long, itemText As String
    PlaceholderDropdown.Clear

    For i = LBound(masterList) To UBound(masterList)
        itemText = CStr(masterList(i))
        If filterText = "" Then
            PlaceholderDropdown.AddItem itemText
        Else
            If InStr(1, itemText, filterText, vbTextCompare) > 0 Then
                PlaceholderDropdown.AddItem itemText
            End If
        End If
    Next i

    If PlaceholderDropdown.ListCount > 0 Then
        PlaceholderDropdown.ListIndex = 0
    Else
        PlaceholderDropdown.ListIndex = -1
    End If

    UpdateButtonsState
End Sub

Private Sub UpdateButtonsState()
    InsertButton.Enabled = (PlaceholderDropdown.ListCount > 0)
    ClearButton.Enabled = (Len(Trim$(SearchBox.Text)) > 0)
End Sub

Private Sub SearchBox_Change()
    PopulateList Trim$(SearchBox.Text)
End Sub

' Arrow keys + Enter/Esc while typing in SearchBox
Private Sub SearchBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            If PlaceholderDropdown.ListCount > 0 Then
                If PlaceholderDropdown.ListIndex < PlaceholderDropdown.ListCount - 1 Then
                    PlaceholderDropdown.ListIndex = PlaceholderDropdown.ListIndex + 1
                End If
            End If
            KeyCode = 0
        Case vbKeyUp
            If PlaceholderDropdown.ListCount > 0 Then
                If PlaceholderDropdown.ListIndex > 0 Then
                    PlaceholderDropdown.ListIndex = PlaceholderDropdown.ListIndex - 1
                End If
            End If
            KeyCode = 0
        Case vbKeyReturn   ' Enter
            If InsertButton.Enabled Then InsertButton_Click
            KeyCode = 0
        Case vbKeyEscape   ' Esc
            CancelButton_Click
            KeyCode = 0
    End Select
End Sub

' Also support Enter/Esc when the ComboBox has focus
Private Sub PlaceholderDropdown_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If InsertButton.Enabled Then InsertButton_Click
            KeyCode = 0
        Case vbKeyEscape
            CancelButton_Click
            KeyCode = 0
    End Select
End Sub

Private Sub InsertButton_Click()
    Dim placeholderText As String

    If PlaceholderDropdown.ListCount = 0 Then
        MsgBox "No matching items to insert.", vbExclamation
        Exit Sub
    End If

    If PlaceholderDropdown.ListIndex >= 0 Then
        placeholderText = PlaceholderDropdown.list(PlaceholderDropdown.ListIndex)
    Else
        ' Fallback if Style isn’t DropDown List
        placeholderText = PlaceholderDropdown.Value
    End If

    If Len(placeholderText) > 0 Then
        InsertPlaceholder placeholderText
        If Not KeepOpen.Value Then
            Unload Me
        Else
            SearchBox.SetFocus
            SearchBox.SelStart = 0
            SearchBox.SelLength = Len(SearchBox.Text)
        End If
    Else
        MsgBox "Please select a placeholder.", vbExclamation
    End If
End Sub

Private Sub PlaceholderDropdown_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If InsertButton.Enabled Then InsertButton_Click
End Sub

Private Sub ClearButton_Click()
    SearchBox.Text = ""
    SearchBox.SetFocus
    UpdateButtonsState
End Sub

Private Sub CancelButton_Click()
    Unload Me
End Sub


