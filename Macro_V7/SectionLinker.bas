Attribute VB_Name = "SectionLinker"
Option Explicit

' --- Public type to pass options from the form ---
Public Type LinkOps
    requireSection As Boolean
    anyDepth As Boolean          ' True: n.n.n..., False: exactly n.n
    AskEach As Boolean           ' True: confirm each, False: change all
    ScopeSelection As Boolean    ' True: only selection, False: whole doc
End Type

' --- Entry points ---
Public Sub OpenHyperlinkManager()
    frmHyperlinkManager.Show
End Sub

Public Sub RunSectionLinker(ByVal ops As LinkOps)
    Dim targetRng As Range
    If ops.ScopeSelection And Selection.Range.Characters.count > 0 Then
        Set targetRng = Selection.Range.Duplicate
    Else
        Set targetRng = ActiveDocument.Range(0, 0)
    End If
    
    Dim pattern As String
    pattern = BuildPattern(ops.requireSection, ops.anyDepth)
    
    Dim matches As Collection
    Set matches = RegexFindAll(targetRng.Text, pattern, True) ' case-insensitive
    
    If matches Is Nothing Or matches.count = 0 Then
        MsgBox "No matches found.", vbInformation, "Hyperlink Manager"
        Exit Sub
    End If
    
    ' Prepare an array of resolved match info with absolute positions
    Dim list() As MatchInfo
    ReDim list(1 To matches.count)
    
    Dim i As Long, m As Object
    For i = 1 To matches.count
        Set m = matches(i) ' VBScript_RegExp_55.Match
        list(i).absStart = targetRng.Start + m.FirstIndex
        list(i).absEnd = list(i).absStart + m.Length
        list(i).FullText = m.Value
        ' SubMatches(0) is the number part (e.g., 1.1 or 1.2.3)
        If m.SubMatches.count > 0 Then
            list(i).NumberPart = CStr(m.SubMatches(0))
        Else
            list(i).NumberPart = ExtractNumberPart(m.Value) ' fallback
        End If
    Next i
    
    ' Process from end to start
    Dim linkedCount As Long, skippedCount As Long, missingBmkCount As Long, cancelled As Boolean
    For i = UBound(list) To LBound(list) Step -1
        Dim mi As MatchInfo
        mi = list(i)
        
        ' Verify bookmark exists
        If Not ActiveDocument.Bookmarks.Exists(mi.NumberPart) Then
            missingBmkCount = missingBmkCount + 1
            GoTo ContinueLoop
        End If
        
        ' Confirm if required
        If ops.AskEach Then
            Dim resp As VbMsgBoxResult
            resp = MsgBox( _
                "Link this text:" & vbCrLf & "  " & mi.FullText & vbCrLf & _
                "to bookmark: " & mi.NumberPart & " ?", _
                vbQuestion + vbYesNoCancel, "Create Hyperlink")
            If resp = vbCancel Then
                cancelled = True
                Exit For
            ElseIf resp = vbNo Then
                skippedCount = skippedCount + 1
                GoTo ContinueLoop
            End If
            ' vbYes ? proceed
        End If
        
        ' Make hyperlink
        If MakeHyperlink(mi.absStart, mi.absEnd, mi.NumberPart) Then
            linkedCount = linkedCount + 1
        Else
            skippedCount = skippedCount + 1
        End If
        
ContinueLoop:
    Next i
    
    Dim msg As String
    msg = "Done." & vbCrLf & _
          "Linked: " & linkedCount & vbCrLf & _
          "Skipped: " & skippedCount & vbCrLf & _
          "Missing bookmarks: " & missingBmkCount
    If cancelled Then msg = msg & vbCrLf & "(Operation cancelled by user.)"
    MsgBox msg, vbInformation, "Hyperlink Manager"
End Sub

' --- Build regex pattern ---
' Returns a VBScript.RegExp pattern with one capturing group for the number
Private Function BuildPattern(ByVal requireSection As Boolean, ByVal anyDepth As Boolean) As String
    Dim numPart As String
    If anyDepth Then
        ' e.g., 1.1 or 2.3.4 etc ? (\d+(\.\d+)+)
        numPart = "(\d+(\.\d+)+)"
    Else
        ' exactly n.n ? (\d+\.\d+)
        numPart = "(\d+\.\d+)"
    End If
    
    If requireSection Then
        ' \bSection\s+(\d+(\.\d+)+)
        BuildPattern = "\bSection\s+" & numPart
    Else
        ' just the number pattern
        BuildPattern = "\b" & numPart
    End If
End Function

' --- Minimal structure to hold a match ---
Private Type MatchInfo
    absStart As Long
    absEnd As Long
    FullText As String
    NumberPart As String
End Type

' --- Execute a VBScript.RegExp on a string; returns Collection of Match objects ---
Private Function RegexFindAll(ByVal Text As String, ByVal pattern As String, Optional ByVal ignoreCase As Boolean = True) As Collection
    On Error GoTo EH
    Dim re As Object, ms As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.pattern = pattern
    re.Global = True
    re.MultiLine = True
    re.ignoreCase = ignoreCase
    
    Set ms = re.Execute(Text)
    Dim col As New Collection
    For Each m In ms
        col.Add m
    Next
    Set RegexFindAll = col
    Exit Function
EH:
    Set RegexFindAll = Nothing
End Function

' --- Fallback to pull number part from a full match ---
Private Function ExtractNumberPart(ByVal s As String) As String
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.pattern = "(\d+(\.\d+)+)"
    re.Global = False
    re.ignoreCase = True
    If re.test(s) Then
        Set m = re.Execute(s)(0)
        ExtractNumberPart = m.SubMatches(0) ' first group
    Else
        ExtractNumberPart = ""
    End If
End Function

' --- Create a hyperlink for a given absolute range to SubAddress = bookmark name ---
Private Function MakeHyperlink(ByVal absStart As Long, ByVal absEnd As Long, ByVal bmkName As String) As Boolean
    On Error GoTo EH
    Dim r As Range
    Set r = ActiveDocument.Range(Start:=absStart, End:=absEnd)
    
    ' Remove existing hyperlink on the same range (if any)
    If r.Hyperlinks.count > 0 Then
        Dim i As Long
        For i = r.Hyperlinks.count To 1 Step -1
            r.Hyperlinks(i).Delete
        Next i
    End If
    
    ActiveDocument.Hyperlinks.Add Anchor:=r, Address:="", SubAddress:=bmkName, TextToDisplay:=r.Text
    MakeHyperlink = True
    Exit Function
EH:
    MakeHyperlink = False
End Function


