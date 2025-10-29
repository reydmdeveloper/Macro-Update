Attribute VB_Name = "modShortcutViewer"
#If VBA7 Then
    Declare PtrSafe Function BuildKeyCode Lib "word" ( _
        ByVal KeyCode As Long, _
        Optional ByVal KeyCode2 As Long = 0, _
        Optional ByVal KeyCode3 As Long = 0, _
        Optional ByVal KeyCode4 As Long = 0 _
    ) As Long
#Else
    Declare Function BuildKeyCode Lib "word" ( _
        ByVal KeyCode As Long, _
        Optional ByVal KeyCode2 As Long = 0, _
        Optional ByVal KeyCode3 As Long = 0, _
        Optional ByVal KeyCode4 As Long = 0 _
    ) As Long
#End If



' Modifier Keys
Public Const wdKeyControl = 512
Public Const wdKeyShift = 256
Public Const wdKeyAlt = 1024

' Arrow Keys
Public Const wdKeyUp = 128
Public Const wdKeyDown = 129
Public Const wdKeyLeft = 130
Public Const wdKeyRight = 131

' F1–F12
Public Const wdKeyF1 = 112
Public Const wdKeyF2 = 113
Public Const wdKeyF3 = 114
Public Const wdKeyF4 = 115
Public Const wdKeyF5 = 116
Public Const wdKeyF6 = 117
Public Const wdKeyF7 = 118
Public Const wdKeyF8 = 119
Public Const wdKeyF9 = 120
Public Const wdKeyF10 = 121
Public Const wdKeyF11 = 122
Public Const wdKeyF12 = 123



' Module: modShortcutViewer
Option Explicit

' Populates the ListBox with all macros and assigned shortcuts
Sub PopulateShortcutList(lst As ListBox, Optional filter As String = "")
    Dim macroName As Variant
    Dim kb As KeyBinding
    Dim shortcut As String
    Dim allMacros As Collection: Set allMacros = GetAllMacros()

    lst.Clear
    lst.AddItem PadRight("Macro Name", 40) & "Shortcut"
    lst.AddItem String(60, "-")

    For Each macroName In allMacros
        If filter = "" Or InStr(1, macroName, filter, vbTextCompare) > 0 Then
            shortcut = "(none)"
            For Each kb In KeyBindings
                If kb.KeyCategory = wdKeyCategoryMacro Then
                    If LCase(kb.Command) = LCase(macroName) Or _
                       Right(LCase(kb.Command), Len(macroName)) = LCase(macroName) Then
                        shortcut = kb.keyString
                        Exit For
                    End If
                End If
            Next kb

            lst.AddItem PadRight(macroName, 40) & shortcut
        End If
    Next
End Sub




' Pads a string with spaces for table formatting
Function PadRight(s As Variant, n As Integer) As String
    Dim padding As Integer
    padding = n - Len(CStr(s))
    If padding < 0 Then padding = 0
    PadRight = CStr(s) & Space(padding)
End Function

' Returns a list of all macros in all standard modules
Function GetAllMacros() As Collection
    Dim vbComp As Object, codeMod As Object
    Dim line As Long, totalLines As Long, codeLine As String, macroName As String
    Dim result As New Collection

    For Each vbComp In NormalTemplate.VBProject.VBComponents
        If vbComp.Type = 1 Then ' Standard module
            Set codeMod = vbComp.CodeModule
            totalLines = codeMod.CountOfLines
            line = 1
            Do While line <= totalLines
                codeLine = Trim(codeMod.Lines(line, 1))
                If Left(codeLine, 4) = "Sub " Then
                    macroName = Split(Split(codeLine, "Sub ")(1), "(")(0)
                    result.Add macroName
                End If
                line = line + 1
            Loop
        End If
    Next vbComp

    Set GetAllMacros = result
End Function


' Assign a keyboard shortcut to a macro
Sub AssignShortcut(macroName As String, ShortcutText As String, Optional KeyCode As Long = 0)
    On Error Resume Next
    CustomizationContext = NormalTemplate

    Dim kb As KeyBinding
    ' Clear existing shortcut for the macro
    For Each kb In KeyBindings
        If InStr(1, kb.Command, macroName, vbTextCompare) > 0 Then
            kb.Clear
        End If
    Next kb

    ' Assign new shortcut
    If KeyCode = 0 Then KeyCode = BuildKeyCodeFromString(ShortcutText)
    If KeyCode > 0 Then
        KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
            Command:=macroName, KeyCode:=KeyCode
    End If
End Sub

' Converts text-based shortcut into a usable KeyCode
Function BuildKeyCodeFromString(s As String) As Long
    Dim ctrl As Boolean, Shift As Boolean, alt As Boolean
    ctrl = InStr(s, "Ctrl") > 0
    Shift = InStr(s, "Shift") > 0
    alt = InStr(s, "Alt") > 0

    Dim key As String, baseKey As Long
    key = Split(s, "+")(UBound(Split(s, "+")))
    Select Case Trim(UCase(key))
        Case "UP": baseKey = wdKeyUp
        Case "DOWN": baseKey = wdKeyDown
        Case "LEFT": baseKey = wdKeyLeft
        Case "RIGHT": baseKey = wdKeyRight
        Case Else
            If Len(key) = 1 Then
                baseKey = Asc(UCase(key))
            ElseIf Left(key, 1) = "F" Then
                baseKey = GetFKeyCode(key)
            End If
    End Select

    Dim modif As Long
    If ctrl Then modif = modif + wdKeyControl
    If alt Then modif = modif + wdKeyAlt
    If Shift Then modif = modif + wdKeyShift

    BuildKeyCodeFromString = BuildKeyCode(baseKey, modif)
End Function

' Returns KeyCode from current form selections
Function GetKeyCode() As Long
    On Error GoTo FailSafe

    Dim modifier As Long: modifier = 0
    If frmShortcutManager.chkCtrl.Value Then modifier = modifier + wdKeyControl
    If frmShortcutManager.chkAlt.Value Then modifier = modifier + wdKeyAlt
    If frmShortcutManager.chkShift.Value Then modifier = modifier + wdKeyShift

    Dim keyText As String: keyText = Trim(frmShortcutManager.cmbKey.Value)
    If keyText = "" Then GoTo FailSafe

    Dim baseKey As Long
    Select Case UCase(keyText)
        Case "UP": baseKey = wdKeyUp
        Case "DOWN": baseKey = wdKeyDown
        Case "LEFT": baseKey = wdKeyLeft
        Case "RIGHT": baseKey = wdKeyRight
        Case "TAB": baseKey = wdKeyTab
        Case "ESC": baseKey = wdKeyEscape
        Case Else
            If Left(UCase(keyText), 1) = "F" Then
                baseKey = GetFKeyCode(keyText)
                If baseKey = 0 Then
                    MsgBox "Invalid function key: " & keyText, vbExclamation
                    GoTo FailSafe
                End If
            ElseIf Len(keyText) = 1 And Asc(UCase(keyText)) >= 65 And Asc(UCase(keyText)) <= 90 Then
                baseKey = Asc(UCase(keyText))
            Else
                MsgBox "Invalid key: " & keyText, vbExclamation
                GoTo FailSafe
            End If
    End Select

    If baseKey > 0 Then
        GetKeyCode = BuildKeyCode(baseKey, modifier)
    Else
        GoTo FailSafe
    End If
    Exit Function

FailSafe:
    GetKeyCode = 0
End Function





' Converts combo box & checkbox values to shortcut string
Function ShortcutText() As String
    Dim s As String
    If frmShortcutManager.chkCtrl Then s = s & "Ctrl+"
    If frmShortcutManager.chkAlt Then s = s & "Alt+"
    If frmShortcutManager.chkShift Then s = s & "Shift+"
    ShortcutText = s & frmShortcutManager.cmbKey.Value
End Function

' Maps F-key strings to wdKey constants
Function GetFKeyCode(fKey As String) As Long
    Select Case UCase(fKey)
        Case "F1": GetFKeyCode = wdKeyF1
        Case "F2": GetFKeyCode = wdKeyF2
        Case "F3": GetFKeyCode = wdKeyF3
        Case "F4": GetFKeyCode = wdKeyF4
        Case "F5": GetFKeyCode = wdKeyF5
        Case "F6": GetFKeyCode = wdKeyF6
        Case "F7": GetFKeyCode = wdKeyF7
        Case "F8": GetFKeyCode = wdKeyF8
        Case "F9": GetFKeyCode = wdKeyF9
        Case "F10": GetFKeyCode = wdKeyF10
        Case "F11": GetFKeyCode = wdKeyF11
        Case "F12": GetFKeyCode = wdKeyF12
        Case Else: GetFKeyCode = 0
    End Select
End Function


