VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmShortcutManager 
   Caption         =   "Macro Manager"
   ClientHeight    =   5280
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   8436.001
   OleObjectBlob   =   "frmShortcutManager.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmShortcutManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Private Sub btnAssign_Click()
'    If lstShortcuts.ListIndex < 2 Then MsgBox "Select a macro", vbExclamation: Exit Sub
'    Dim macroName As String: macroName = Trim(Split(lstShortcuts.List(lstShortcuts.ListIndex), " ")(0))
'    Dim KeyCode As Long: KeyCode = GetKeyCode()
'    If KeyCode = 0 Then MsgBox "Invalid key", vbExclamation: Exit Sub
'    AssignShortcut macroName, ShortcutText(), KeyCode
'    PopulateShortcutList lstShortcuts, txtFilter.Text
'    MsgBox "Shortcut assigned!", vbInformation
'End Sub
'
'
'Private Sub btnClear_Click()
'    If lstShortcuts.ListIndex < 2 Then MsgBox "Select a macro", vbExclamation: Exit Sub
'    Dim macroName As String: macroName = Trim(Split(lstShortcuts.List(lstShortcuts.ListIndex), " ")(0))
'    Dim kb As KeyBinding
'    CustomizationContext = NormalTemplate
'    For Each kb In KeyBindings
'        If InStr(1, kb.Command, macroName, vbTextCompare) > 0 Then kb.Clear
'    Next
'    PopulateShortcutList lstShortcuts, txtFilter.Text
'    MsgBox "Shortcut cleared", vbInformation
'End Sub



Private Sub btnExport_Click()
    Dim dlg As FileDialog
    Dim filePath As String
    Dim f As Integer
    Dim kb As KeyBinding

    Set dlg = Application.FileDialog(msoFileDialogSaveAs)
    With dlg
        .Title = "Export Macro Shortcuts to HTML"
        .InitialFileName = "Shortcuts.html"
        If .Show <> -1 Then Exit Sub
        filePath = .SelectedItems(1)
    End With

    ' Ensure .html extension
    If LCase(Right(filePath, 5)) <> ".html" Then filePath = filePath & ".html"

    ' Write the HTML content
    On Error GoTo ExportError
    f = FreeFile
    Open filePath For Output As #f

    Print #f, "<html><head><title>Macro Shortcuts</title></head><body>"
    Print #f, "<h2>Macro Shortcuts (Normal.dotm)</h2>"
    Print #f, "<table border='1' cellpadding='5' cellspacing='0'>"
    Print #f, "<tr><th>Macro Name</th><th>Shortcut Key</th></tr>"

    For Each kb In Application.KeyBindings
        If kb.KeyCategory = wdKeyCategoryMacro Then
            If Not kb.context Is Nothing And kb.context.Name = "Normal" Then
                If Len(kb.keyString) > 0 Then
                    Print #f, "<tr><td>" & kb.Command & "</td><td>" & kb.keyString & "</td></tr>"
                End If
            End If
        End If
    Next

    Print #f, "</table></body></html>"
    Close #f

    MsgBox "? Exported to: " & filePath, vbInformation
    Exit Sub

ExportError:
    MsgBox "? Export failed: " & Err.Description, vbCritical
    On Error Resume Next: Close #f
End Sub


Private Sub btnImport_Click()
    Dim dlg As FileDialog
    Dim filePath As String
    Dim f As Integer
    Dim line As String
    Dim macroName As String, shortcutKey As String
    Dim parts() As String
    Dim kb As KeyBinding
    Dim logMsg As String

    ' Ask user to pick the HTML file
    Set dlg = Application.FileDialog(msoFileDialogFilePicker)
    With dlg
        .Title = "Import Macro Shortcuts from HTML"
        If .Show <> -1 Then Exit Sub
        filePath = .SelectedItems(1)
    End With

    On Error GoTo ImportError
    f = FreeFile
    Open filePath For Input As #f

    Application.CustomizationContext = NormalTemplate
    logMsg = "Imported:" & vbCrLf

    Do Until EOF(f)
        Line Input #f, line

        If InStr(line, "<tr><td>") > 0 Then
            line = Replace(line, "<tr><td>", "")
            line = Replace(line, "</td><td>", "|")
            line = Replace(line, "</td></tr>", "")
            parts = Split(line, "|")

            If UBound(parts) = 1 Then
                macroName = Trim(parts(0))
                shortcutKey = Trim(parts(1))

                logMsg = logMsg & macroName & " ? " & shortcutKey

                ' Remove existing binding
                For Each kb In Application.KeyBindings
                    If LCase(kb.Command) = LCase(macroName) And Not kb.context Is Nothing Then
                        If kb.context.Name = "Normal" Then kb.Clear
                    End If
                Next

                ' Assign new shortcut
                On Error Resume Next
                KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
                                Command:=macroName, _
                                KeyCode:=Application.BuildKeyCodeFromString(shortcutKey)
                On Error GoTo 0

                logMsg = logMsg & vbCrLf
            End If
        End If
    Loop

    Close #f
    MsgBox "? Import complete:" & vbCrLf & logMsg, vbInformation
    Exit Sub

ImportError:
    MsgBox "? Import error: " & Err.Description, vbCritical
    On Error Resume Next: Close #f
End Sub


Function MacroExists(ByVal macroName As String) As Boolean
    On Error Resume Next
    Dim test As Object
    test = Application.Run(macroName)
    MacroExists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function




Private Sub AssignParsedShortcut(macroName As String, shortcutKey As String)
    Dim parts() As String
    Dim KeyCode As Long
    Dim mainKey As String
    Dim ctrl As Boolean, alt As Boolean, Shift As Boolean
    Dim i As Integer

    shortcutKey = Trim(UCase(shortcutKey))
    parts = Split(shortcutKey, "+")

    For i = 0 To UBound(parts)
        Select Case Trim(parts(i))
            Case "CTRL": ctrl = True
            Case "ALT": alt = True
            Case "SHIFT": Shift = True
            Case Else: mainKey = Trim(parts(i))
        End Select
    Next i

    On Error GoTo ErrHandler

    Select Case mainKey
        Case "A" To "Z"
            KeyCode = BuildKeyCode( _
                IIf(ctrl, wdKeyControl, 0) + _
                IIf(alt, wdKeyAlt, 0) + _
                IIf(Shift, wdKeyShift, 0), _
                Asc(mainKey))
        Case "F1" To "F12"
            KeyCode = BuildKeyCode( _
                IIf(ctrl, wdKeyControl, 0) + _
                IIf(alt, wdKeyAlt, 0) + _
                IIf(Shift, wdKeyShift, 0), _
                wdKeyF1 + CInt(Mid(mainKey, 2)) - 1)
        Case Else
            MsgBox "? Unsupported key: " & mainKey, vbExclamation
            Exit Sub
    End Select

    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
                    Command:=macroName, _
                    KeyCode:=KeyCode
    Exit Sub

ErrHandler:
    MsgBox "? Error assigning shortcut to: " & macroName & vbCrLf & Err.Description, vbCritical
End Sub


Private Sub txtFilter_Change()
    PopulateShortcutList lstShortcuts, txtFilter.Text
End Sub

Private Sub UserForm_Initialize()
    ' Populate macros
    PopulateShortcutList Me.lstShortcuts

'    ' Populate key dropdown
'    FillKeyDropdown
End Sub

'Private Sub FillKeyDropdown()
'    Dim i As Integer
'    cmbKey.Clear
'
'    ' A-Z
'    For i = 65 To 90
'        cmbKey.AddItem Chr(i)
'    Next i
'
'    ' F1–F12
'    For i = 1 To 12
'        cmbKey.AddItem "F" & i
'    Next i
'
'    ' Special keys
'    cmbKey.AddItem "Up"
'    cmbKey.AddItem "Down"
'    cmbKey.AddItem "Left"
'    cmbKey.AddItem "Right"
'    cmbKey.AddItem "Tab"
'    cmbKey.AddItem "Esc"
'End Sub







