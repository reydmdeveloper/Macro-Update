Attribute VB_Name = "modCommon"
Option Explicit

' ===== Common utilities used across managers =====

Public Function SanitizeName(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    s = Trim$(s)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Za-z0-9_]" Then
            out = out & ch
        ElseIf ch = " " Or ch = "-" Or ch = "." Then
            out = out & "_"
        End If
    Next i
    If Len(out) = 0 Then out = "B"
    If Not (Left$(out, 1) Like "[A-Za-z]") Then out = "B" & out
SanizeAgain:
    Do While InStr(out, "__") > 0: out = Replace(out, "__", "_"): Loop
    SanitizeName = out
End Function

Public Function AlphaSuffix(ByVal idx As Long) As String
    Dim s As String, n As Long
    n = idx
    Do
        s = Chr$(Asc("a") + (n Mod 26)) & s
        n = n \ 26 - 1
    Loop While n >= 0
    AlphaSuffix = s
End Function

Public Function UniqueBookmarkName(ByVal baseName As String) As String
    Dim i As Long
    If Not ActiveDocument.Bookmarks.Exists(baseName) Then
        UniqueBookmarkName = baseName
        Exit Function
    End If
    For i = 0 To 675
        If Not ActiveDocument.Bookmarks.Exists(baseName & AlphaSuffix(i)) Then
            UniqueBookmarkName = baseName & AlphaSuffix(i)
            Exit Function
        End If
    Next i
    i = 2
    Do While ActiveDocument.Bookmarks.Exists(baseName & "_" & CStr(i))
        i = i + 1
    Loop
    UniqueBookmarkName = baseName & "_" & CStr(i)
End Function

Public Function StripTrailingCR(ByVal r As Range) As Range
    Dim rr As Range: Set rr = r.Duplicate
    If Right$(rr.Text, 1) = vbCr Or Right$(rr.Text, 2) = vbCrLf Then rr.End = rr.End - 1
    Set StripTrailingCR = rr
End Function

Public Function StripLeadingNumberingText(ByVal s As String) As String
    Dim re As Object, T As String
    T = LTrim$(s)
    Set re = CreateObject("VBScript.Regexp")
    re.ignoreCase = True: re.Global = False
    re.pattern = "^\s*([0-9]+(\.[0-9]+)*|[A-Za-z]|[IVXLC]+)(\)|\.)?\s+"
    If re.test(T) Then StripLeadingNumberingText = re.Replace(T, "") Else StripLeadingNumberingText = T
End Function

Private Function MinLong(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then MinLong = a Else MinLong = b
End Function

Public Function FirstNWords(ByVal txt As String, ByVal n As Long) As String
    Dim clean As String, a() As String, i As Long, out As String, lastIdx As Long
    clean = Replace(Replace(Trim$(txt), vbCr, " "), vbLf, " ")
    clean = Application.Trim(clean)
    If Len(clean) = 0 Then FirstNWords = "Heading": Exit Function
    a = Split(clean, " ")
    If n < 1 Then n = 1
    If n > 4 Then n = 4
    lastIdx = MinLong(UBound(a), LBound(a) + n - 1)
    For i = LBound(a) To lastIdx
        If out <> "" Then out = out & "_" & a(i) Else out = a(i)
    Next i
    FirstNWords = out
End Function

Public Function GetNumberChain(ByVal p As Paragraph) As String
    On Error GoTo noList
    Dim s As String
    s = p.Range.listFormat.ListString
    If Len(s) = 0 Then GoTo noList
    s = Trim$(s)
    Do While Len(s) > 0 And (Right$(s, 1) = "." Or Right$(s, 1) = ")" Or Right$(s, 1) = ":")
        s = Left$(s, Len(s) - 1)
    Loop
    s = Replace(s, ".", "_")
    s = Replace(s, " ", "")
    GetNumberChain = s
    Exit Function
noList:
    GetNumberChain = ""
End Function

Public Function ProperCase(ByVal s As String) As String
    If Len(s) = 0 Then ProperCase = s Else ProperCase = UCase$(Left$(s, 1)) & Mid$(s, 2)
End Function

