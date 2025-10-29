Attribute VB_Name = "modREY_Bookmarks"
Sub InsertBookmark_CurrentParagraph()
    Dim para As Paragraph
    Dim paraText As String
    Dim listNum As String
    Dim re As Object
    Dim match As Object
    Dim bookmarkName As String
    Dim safeName As String
    Dim rng As Range
    Dim bm As Bookmark
    Dim bmExists As Boolean
    Dim f As Field
    Dim captionLabel As String
    Dim captionNum As String
    Dim userInput As String
    Dim authorName As String
    Dim pubYear As String
    Dim yearMatches As Object
    Dim yearPattern As Object
    Dim lastYear As String
    Dim i As Long, ch As String
    Dim candidateYear As Long
    Dim chosenYear As String

    On Error GoTo ErrHandler

    Set para = Selection.Paragraphs(1)
    paraText = Trim(para.Range.Text)

    ' === Capture autonumber value if paragraph has one ===
    listNum = ""
    If para.Range.listFormat.ListType <> wdListNoNumbering Then
        listNum = Trim(para.Range.listFormat.ListString)
    End If

    ' === Create RegExp ===
    Set re = CreateObject("VBScript.RegExp")
    re.ignoreCase = True
    re.Global = False

    ' === 1?? FIGURE / TABLE fields ===
    For Each f In para.Range.Fields
        If InStr(1, f.Code.Text, "SEQ", vbTextCompare) > 0 Then
            If InStr(1, f.Code.Text, "Figura", vbTextCompare) > 0 Or _
               InStr(1, f.Code.Text, "Figure", vbTextCompare) > 0 Then
                bookmarkName = "Fig_" & Replace(f.result.Text, ".", "_")
                GoTo CleanAndAdd
            End If
            If InStr(1, f.Code.Text, "Tabela", vbTextCompare) > 0 Or _
               InStr(1, f.Code.Text, "Table", vbTextCompare) > 0 Then
                bookmarkName = "Tab_" & Replace(f.result.Text, ".", "_")
                GoTo CleanAndAdd
            End If
        End If
    Next f

    ' === 2?? SECTION detection (autonumber or manual) ===
    re.pattern = "^\s*(\d+([\.]\d+)*)"
    If listNum <> "" Then
        If re.test(listNum) Then
            bookmarkName = "Sec_" & Replace(listNum, ".", "_")
            GoTo CleanAndAdd
        End If
    End If
    If re.test(paraText) Then
        Set match = re.Execute(paraText)(0)
        bookmarkName = "Sec_" & Replace(match.SubMatches(0), ".", "_")
        GoTo CleanAndAdd
    End If

    ' === 3?? Manual FIGURE / TABLE text ===
    re.pattern = "^\s*(Figure|Figura)\s*(\d+([\.]\d+)*)"
    If re.test(paraText) Then
        Set match = re.Execute(paraText)(0)
        bookmarkName = "Fig_" & Replace(match.SubMatches(1), ".", "_")
        GoTo CleanAndAdd
    End If

    re.pattern = "^\s*(Table|Tabela)\s*(\d+([\.]\d+)*)"
    If re.test(paraText) Then
        Set match = re.Execute(paraText)(0)
        bookmarkName = "Tab_" & Replace(match.SubMatches(1), ".", "_")
        GoTo CleanAndAdd
    End If

    ' === 4?? APPENDIX / APÊNDICE ===
    re.pattern = "^\s*(Appendix|Apêndice)\s*([A-Za-z0-9]+)"
    If re.test(paraText) Then
        Set match = re.Execute(paraText)(0)
        bookmarkName = "App_" & match.SubMatches(1)
        GoTo CleanAndAdd
    End If

    ' === 5?? REFERENCES section header ===
    If LCase(paraText) Like "*referenc*" Or LCase(paraText) Like "*bibliogr*" Then
        bookmarkName = "Ref_Main"
        GoTo CleanAndAdd
    End If

    ' === 6?? Reference entries (APA, Vancouver, Abstract styles) ===
    ' --- APA/Harvard (Author (Year)) ---
    re.pattern = "^\s*([A-ZÁÉÍÓÚÜÑÇ][A-Za-zÁÉÍÓÚÜÑÇ\-\']+).*?\((\d{4}[a-z]?)\)"
    If re.test(paraText) Then
        Set match = re.Execute(paraText)(0)
        authorName = match.SubMatches(0)
        pubYear = match.SubMatches(1)
        authorName = Split(authorName, " ")(0)
        bookmarkName = "Ref_" & authorName & "_" & pubYear
        GoTo CleanAndAdd
    End If

    ' --- Vancouver (Author ... 2010;28:...) ---
    re.pattern = "^\s*([A-ZÁÉÍÓÚÜÑÇ][A-Za-zÁÉÍÓÚÜÑÇ\-\']+).*?(\d{4}[a-z]?)[;:\.]"
    If re.test(paraText) Then
        Set match = re.Execute(paraText)(0)
        authorName = match.SubMatches(0)
        pubYear = match.SubMatches(1)
        If InStr(authorName, " ") > 0 Then authorName = Split(authorName, " ")(0)
        If InStr(authorName, ",") > 0 Then authorName = Split(authorName, ",")(0)
        authorName = Replace(authorName, ".", "")
        bookmarkName = "Ref_" & authorName & "_" & pubYear
        GoTo CleanAndAdd
    End If

    ' --- Flexible: first author + best valid year in text (prefer 1900-2099) ---
    Set re = CreateObject("VBScript.RegExp")
    re.ignoreCase = True
    re.Global = False
    re.pattern = "^\s*([A-ZÁÉÍÓÚÜÑÇ][A-Za-zÁÉÍÓÚÜÑÇ\-\']+)"
    If re.test(paraText) Then
        Set match = re.Execute(paraText)(0)
        authorName = match.SubMatches(0)
    Else
        authorName = "Ref"
    End If

    Set yearPattern = CreateObject("VBScript.RegExp")
    yearPattern.Global = True
    yearPattern.ignoreCase = True
    yearPattern.pattern = "(\d{4}[a-z]?)"

    chosenYear = ""
    If yearPattern.test(paraText) Then
        Set yearMatches = yearPattern.Execute(paraText)
        ' Prefer first year in reasonable publication range 1900-2099
        For i = 0 To yearMatches.count - 1
            Dim ytext As String
            ytext = yearMatches(i).SubMatches(0)
            ' numeric part:
            Dim numericPart As String
            numericPart = ""
            numericPart = ytext
            ' remove trailing alpha (a/b) if present for numeric check
            If Len(numericPart) > 0 Then
                If Not IsNumeric(Right(numericPart, 1)) Then
                    numericPart = Left(numericPart, Len(numericPart) - 1)
                End If
            End If
            If numericPart <> "" Then
                If IsNumeric(numericPart) Then
                    candidateYear = CLng(numericPart)
                    If candidateYear >= 1900 And candidateYear <= 2099 Then
                        chosenYear = yearMatches(i).SubMatches(0)
                        Exit For
                    End If
                End If
            End If
        Next i
        ' if no year in 1900-2099 found, fallback to last match
        If chosenYear = "" Then
            lastYear = yearMatches(yearMatches.count - 1).SubMatches(0)
            chosenYear = lastYear
        End If
        bookmarkName = "Ref_" & Replace(authorName, ".", "") & "_" & chosenYear
        GoTo CleanAndAdd
    End If

    ' === 7?? Manual fallback ===
    userInput = InputBox("No pattern found. Enter bookmark name (e.g., Smith_2021a):", _
                         "Manual Bookmark Entry")
    If Trim(userInput) <> "" Then
        bookmarkName = Replace(userInput, " ", "_")
        GoTo CleanAndAdd
    Else
        Exit Sub
    End If

CleanAndAdd:
    ' === Sanitize bookmark name (remove invalid chars) ===
    safeName = ""
    For i = 1 To Len(bookmarkName)
        ch = Mid(bookmarkName, i, 1)
        If ch Like "[A-Za-z0-9_]" Then
            safeName = safeName & ch
        Else
            safeName = safeName & "_"
        End If
    Next i

    ' Replace multiple underscores and trim trailing
    Do While InStr(safeName, "__") > 0
        safeName = Replace(safeName, "__", "_")
    Loop
    If Right(safeName, 1) = "_" Then
        safeName = Left(safeName, Len(safeName) - 1)
    End If

    bookmarkName = safeName

AddBookmark:
    ' === Prevent duplicates ===
    bmExists = False
    For Each bm In ActiveDocument.Bookmarks
        If LCase(bm.Name) = LCase(bookmarkName) Then
            bmExists = True
            Exit For
        End If
    Next bm

    If bmExists Then
        MsgBox "Bookmark '" & bookmarkName & "' already exists.", vbInformation
    Else
        Set rng = para.Range
        rng.Collapse Direction:=wdCollapseStart
        ActiveDocument.Bookmarks.Add Name:=bookmarkName, Range:=rng
        MsgBox "? Bookmark inserted: " & bookmarkName, vbInformation
    End If
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub



Sub InsertCrossRef_ToSelectedTextBookmark()
    Dim selText As String
    Dim re As Object
    Dim match As Object
    Dim bookmarkName As String
    Dim bm As Bookmark
    Dim bmFound As Boolean
    Dim linkRange As Range
    Dim safeName As String
    Dim ch As String
    Dim i As Long

    On Error GoTo ErrHandler

    selText = Trim(Selection.Text)
    If selText = "" Then
        MsgBox "Please select the text to convert into a cross-reference.", vbExclamation
        Exit Sub
    End If

    ' --- Prepare regex ---
    Set re = CreateObject("VBScript.RegExp")
    re.ignoreCase = True
    re.Global = False

    ' === Detect SECTION variants: Section / Seção / Sec. / Sec ===
    re.pattern = "^(Sec(tion|ção)?\.?)\s+(\d+([\.]\d+)*)"
    If re.test(selText) Then
        Set match = re.Execute(selText)(0)
        bookmarkName = "Sec_" & Replace(match.SubMatches(2), ".", "_")
        GoTo TryInsert
    End If

    ' === Detect FIGURE ===
    re.pattern = "(Figure|Figura)\s+(\d+([\.]\d+)*)"
    If re.test(selText) Then
        Set match = re.Execute(selText)(0)
        bookmarkName = "Fig_" & Replace(match.SubMatches(1), ".", "_")
        GoTo TryInsert
    End If

    ' === Detect TABLE ===
    re.pattern = "(Table|Tabela)\s+(\d+([\.]\d+)*)"
    If re.test(selText) Then
        Set match = re.Execute(selText)(0)
        bookmarkName = "Tab_" & Replace(match.SubMatches(1), ".", "_")
        GoTo TryInsert
    End If

    ' === Detect APPENDIX ===
    re.pattern = "(Appendix|Apêndice)\s*([A-Za-z0-9]+)"
    If re.test(selText) Then
        Set match = re.Execute(selText)(0)
        bookmarkName = "App_" & match.SubMatches(1)
        GoTo TryInsert
    End If

    ' === Detect full Ref_ pattern ===
    re.pattern = "(Ref[_\-A-Za-z0-9_]+)"
    If re.test(selText) Then
        Set match = re.Execute(selText)(0)
        bookmarkName = match.Value
        GoTo TryInsert
    End If

    ' === Detect simple author-year pattern (e.g., Pritchard 2013) ===
    re.pattern = "^\s*([A-ZÁÉÍÓÚÜÑÇ][A-Za-zÁÉÍÓÚÜÑÇ\-']+)\s+(\d{4}[a-z]?)\s*$"
    If re.test(selText) Then
        Set match = re.Execute(selText)(0)
        bookmarkName = "Ref_" & match.SubMatches(0) & "_" & match.SubMatches(1)
        GoTo TryInsert
    End If

    ' === No valid pattern ===
    MsgBox "No valid pattern found in selected text." & vbCrLf & _
           "Expected formats: Section 1.1 / Sec. 18.1 / Figure 2.3 / Table 4 / Appendix A / Pritchard 2013 / Ref_Smith_2021.", vbInformation
    Exit Sub

TryInsert:
    ' === Clean bookmark name ===
    safeName = ""
    For i = 1 To Len(bookmarkName)
        ch = Mid(bookmarkName, i, 1)
        If ch Like "[A-Za-z0-9_]" Then
            safeName = safeName & ch
        Else
            safeName = safeName & "_"
        End If
    Next i

    Do While InStr(safeName, "__") > 0
        safeName = Replace(safeName, "__", "_")
    Loop
    If Right(safeName, 1) = "_" Then
        safeName = Left(safeName, Len(safeName) - 1)
    End If

    bookmarkName = safeName

    ' === Check if bookmark exists ===
    bmFound = False
    For Each bm In ActiveDocument.Bookmarks
        If StrComp(bm.Name, bookmarkName, vbTextCompare) = 0 Then
            bmFound = True
            Exit For
        End If
    Next bm

    If Not bmFound Then
        MsgBox "Bookmark '" & bookmarkName & "' not found in this document.", vbExclamation
        Exit Sub
    End If

    ' === Add hyperlink without changing formatting ===
    Set linkRange = Selection.Range.Duplicate
    ActiveDocument.Hyperlinks.Add Anchor:=linkRange, _
        Address:="", SubAddress:=bookmarkName, TextToDisplay:=selText

    ' --- Preserve formatting ---
    With linkRange.Font
        .Underline = Selection.Font.Underline
        .Color = Selection.Font.Color
        .Bold = Selection.Font.Bold
        .Italic = Selection.Font.Italic
        .Size = Selection.Font.Size
        .Name = Selection.Font.Name
    End With

    MsgBox "? Cross-reference inserted to bookmark '" & bookmarkName & "'.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub


