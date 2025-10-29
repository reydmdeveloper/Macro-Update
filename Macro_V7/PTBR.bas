Attribute VB_Name = "PTBR"
Sub CreateCenteredTitleWithLines()
    Dim doc As Document
    Dim tbl As Table
    Dim userText As String
    Dim r As Integer, c As Integer

    Set doc = ActiveDocument
    userText = InputBox("Enter the title text for the center cell:", "Center Title")
    If userText = "" Then Exit Sub

    ' Insert 2-row, 3-column table at selection
    Set tbl = doc.Tables.Add(Selection.Range, 2, 3)

    ' Set column widths
    tbl.Columns(1).Width = InchesToPoints(0.31)
    tbl.Columns(2).Width = InchesToPoints(1.62)
    tbl.Columns(3).Width = InchesToPoints(5.53)

    ' Merge middle column vertically
    tbl.cell(1, 2).Merge tbl.cell(2, 2)

    ' Set center cell text and formatting (Arial, 8pt, bold, centered, no spacing)
    With tbl.cell(1, 2).Range
        .Text = userText
        .Font.Name = "Arial"
        .Font.Size = 8
        .Font.Bold = True
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .ParagraphFormat.SpaceBefore = 0
        .ParagraphFormat.SpaceAfter = 0
    End With

    ' Set font for other cells (Arial, 4pt, no spacing)
    For r = 1 To 2
        For c = 1 To 3 Step 2 ' Only columns 1 and 3
            With tbl.cell(r, c).Range
                .Text = ""
                .Font.Name = "Arial"
                .Font.Size = 4
                .Font.Bold = False
                .ParagraphFormat.SpaceBefore = 0
                .ParagraphFormat.SpaceAfter = 0
            End With
        Next c
    Next r

    ' Add bottom borders to Row 1, Col 1 and Col 3
    With tbl.cell(1, 1).Borders(wdBorderBottom)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth025pt
        .Color = wdColorAutomatic
    End With
    With tbl.cell(1, 3).Borders(wdBorderBottom)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth025pt
        .Color = wdColorAutomatic
    End With

    ' Remove all cell margins (padding)
    With tbl
        .TopPadding = 0
        .BottomPadding = 0
        .LeftPadding = 0
        .RightPadding = 0
    End With

    ' Set row alignment and apply -0.04" left indent
    Selection.Tables(1).Rows.LeftIndent = InchesToPoints(-0.04)
End Sub
Sub FormatSelectedTableLayout_PTBR()
    Dim tbl As Table
    Dim r As row, c As cell
    Dim para As Paragraph
    Dim rowIndex As Integer, colIndex As Integer
    Dim paraText As String
    Dim charLimit As Integer

    ' Ensure selection has a table
    If Selection.Tables.count = 0 Then
        MsgBox "Please place the cursor inside a table.", vbExclamation
        Exit Sub
    End If

    Set tbl = Selection.Tables(1)

    ' Set left indent
    tbl.Rows.LeftIndent = InchesToPoints(0.06)

    ' Set cell padding: all 0 except left = 0.02"
    With tbl
        .TopPadding = 0
        .BottomPadding = 0
        .RightPadding = 0
        .LeftPadding = InchesToPoints(0.02)
    End With

    ' Set row height = 3pt with "At least"
    For Each r In tbl.Rows
        With r
            .Height = 3
            .HeightRule = wdRowHeightAtLeast
        End With
    Next r

    ' Loop through cells
    For rowIndex = 1 To tbl.Rows.count
        For colIndex = 1 To tbl.Columns.count
            Set c = tbl.cell(rowIndex, colIndex)

            ' Remove paragraph returns if more than one
            With c.Range
                If .Paragraphs.count > 1 Then
                    .Text = Replace(.Text, vbCr, " ")
                    .Text = Trim(.Text)
                End If
            End With

            ' Special cleanup for column 2: remove breaks and spaces
            If colIndex = 2 Then
                With c.Range.Paragraphs(1).Range
                    Dim cleanedText As String
                    cleanedText = .Text
                    cleanedText = Replace(cleanedText, vbCr, "")
                    cleanedText = Replace(cleanedText, Chr(11), "")
                    cleanedText = Replace(cleanedText, " ", "")
                    cleanedText = Trim(cleanedText)
                    .Text = cleanedText
                    .MoveEnd Unit:=wdCharacter, count:=-1
                End With
            End If

            ' Set character limit by column
            Select Case colIndex
                Case 1: charLimit = 53    ' Column 1 (updated from 5 to 53)
                Case 2: charLimit = 80    ' Column 2
                Case Else: charLimit = 80
            End Select

            ' Apply paragraph formatting
            For Each para In c.Range.Paragraphs
                paraText = Replace(para.Range.Text, vbCr, "")

                With para.Format
                    If rowIndex = 1 Then
                        .Alignment = wdAlignParagraphCenter
                        .SpaceBefore = 0
                        .SpaceAfter = 1
                    Else
                        .Alignment = wdAlignParagraphLeft
                        .SpaceBefore = 1
                        If Len(paraText) > charLimit Then
                            .SpaceAfter = 1.5
                        Else
                            .SpaceAfter = 3.5
                        End If
                    End If
                End With
            Next para

            ' Header row background color
            If rowIndex = 1 Then
                c.Shading.BackgroundPatternColor = RGB(204, 204, 204)
            End If
        Next colIndex
    Next rowIndex

    MsgBox "PT-BR table formatting complete with updated spacing logic.", vbInformation
End Sub

Sub ApplyRecentTextFill_PTBR()
    Dim userInput As String
    Dim r As Long, g As Long, b As Long
    Dim pickedColor As Long
    Dim lastUsedColor As Long

    ' Try to get the last stored color
    On Error Resume Next
    lastUsedColor = CLng(ActiveDocument.Variables("PTBR_TextFillColor").Value)
    On Error GoTo 0

    userInput = InputBox( _
        "Enter a color in Hex (#RRGGBB) or RGB (e.g., 204,204,255)." & vbCrLf & _
        "Leave empty to use the most recent color.", _
        "Apply Fill Color")

    If Trim(userInput) = "" Then
        If lastUsedColor = 0 Then
            MsgBox "No previous color found.", vbExclamation
            Exit Sub
        End If
        pickedColor = lastUsedColor
    ElseIf InStr(userInput, "#") = 1 Then
        Dim hexVal As String
        hexVal = Replace(userInput, "#", "")
        If Len(hexVal) <> 6 Or Not hexVal Like "[0-9A-Fa-f]*" Then
            MsgBox "Invalid hex format. Use #RRGGBB.", vbExclamation
            Exit Sub
        End If
        r = CLng("&H" & Mid(hexVal, 1, 2))
        g = CLng("&H" & Mid(hexVal, 3, 2))
        b = CLng("&H" & Mid(hexVal, 5, 2))
        pickedColor = RGB(r, g, b)
    ElseIf InStr(userInput, ",") > 0 Then
        Dim rgbParts() As String
        rgbParts = Split(userInput, ",")
        If UBound(rgbParts) <> 2 Then
            MsgBox "Invalid RGB format. Use R,G,B", vbExclamation
            Exit Sub
        End If
        r = CLng(Trim(rgbParts(0)))
        g = CLng(Trim(rgbParts(1)))
        b = CLng(Trim(rgbParts(2)))
        pickedColor = RGB(r, g, b)
    Else
        MsgBox "Invalid input. Enter hex or RGB.", vbExclamation
        Exit Sub
    End If

    ' Save for next time
    ActiveDocument.Variables("PTBR_TextFillColor").Value = CStr(pickedColor)

    ' Apply to selection only
    If Selection.Type = wdSelectionNormal Then
        Selection.Font.Shading.BackgroundPatternColor = pickedColor
    Else
        MsgBox "Please select text first.", vbExclamation
    End If
End Sub


Sub PTBR_format1()
    Dim doc As Document: Set doc = ActiveDocument
    Dim i As Long, j As Integer
    Dim secRange As Range, rng As Range
    Dim placeholderText As String, parecerNum As String
    Dim tbl As Table, hdrRange As Range, otherHdrRange As Range
    Dim cellRange As Range, c As Long

    ' ===== 1. Remove Section Breaks and Page Breaks =====
    For i = doc.Sections.count - 1 To 1 Step -1
        Set secRange = doc.Sections(i).Range
        secRange.Collapse Direction:=wdCollapseEnd
        If secRange.Characters.Last.Previous = Chr(12) Then
            secRange.MoveStart wdCharacter, -1
            secRange.Delete
        End If
    Next i

    With Selection.Find
        .ClearFormatting: .replacement.ClearFormatting
        .Text = "^m": .replacement.Text = "^p"
        .Forward = True: .Wrap = wdFindContinue
        .Format = False: .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With

    With Selection.Find
        .Text = "^n": .replacement.Text = "^p"
        .Execute Replace:=wdReplaceAll
    End With

    ' ===== 2. Set Page Layout =====
    With doc.PageSetup
        .TopMargin = InchesToPoints(1.69)
        .BottomMargin = InchesToPoints(1.71)
        .LeftMargin = InchesToPoints(0.78)
        .RightMargin = InchesToPoints(0.77)
        .HeaderDistance = InchesToPoints(0.76)
        .FooterDistance = InchesToPoints(0.57)
        .DifferentFirstPageHeaderFooter = True
        .OddAndEvenPagesHeaderFooter = False
    End With

    ' ===== 3. Font and Paragraph Cleanup =====
    Selection.WholeStory
    With Selection.Font
        .Name = "Arial"
        .Spacing = 0
        .Scaling = 100
        .Position = 0
    End With

    With Selection.ParagraphFormat
        .LeftIndent = 0
        .RightIndent = 0
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .FirstLineIndent = 0
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = 1.415 * 12
    End With

    ' ===== 4. Convert Numbering to Text =====
    doc.ConvertNumbersToText

    ' ===== 5. Remove Extra Paragraphs =====
    For j = 1 To 15
        Set rng = doc.Content
        With rng.Find
            .ClearFormatting: .replacement.ClearFormatting
            .Text = "^p^p": .replacement.Text = "^p"
            .Forward = True: .Wrap = wdFindStop
            .Format = False: .MatchWildcards = False
        End With
        rng.Find.Execute Replace:=wdReplaceAll
    Next j

    ' ===== 6. User Inputs =====
    placeholderText = InputBox("Enter placeholder title for header (center column):", "Header Title", "[Title]")
    If placeholderText = "" Then placeholderText = "[Title]"

    parecerNum = InputBox("Enter number for 'Continuação do Parecer:'", "Parecer Number", "[insert number]")
    If parecerNum = "" Then parecerNum = "[insert number]"

    ' ===== 7. First Page Header =====
    Set hdrRange = doc.Sections(1).Headers(wdHeaderFooterFirstPage).Range
    hdrRange.Text = ""
    Set tbl = hdrRange.Tables.Add(Range:=hdrRange, numRows:=1, NumColumns:=3)

    With tbl
        .AllowAutoFit = False
        .PreferredWidthType = wdPreferredWidthPoints
        .Spacing = 0
        .TopPadding = 0: .BottomPadding = 0
        .LeftPadding = 0: .RightPadding = 0
        .Rows.Alignment = wdAlignRowLeft
        .Rows(1).HeightRule = wdRowHeightAuto

        .Columns(1).PreferredWidth = InchesToPoints(1.63)
        .Columns(2).PreferredWidth = InchesToPoints(3.26)
        .Columns(3).PreferredWidth = InchesToPoints(1.81)

        On Error Resume Next
        .Borders(wdBorderInsideVertical).LineStyle = wdLineStyleNone
        .Borders(wdBorderInsideHorizontal).LineStyle = wdLineStyleNone
        On Error GoTo 0

        For c = wdBorderTop To wdBorderRight
            With .Borders(c)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth075pt
                .Color = RGB(166, 166, 166)
            End With
        Next c
    End With

    For c = 1 To 3
        tbl.cell(1, c).VerticalAlignment = wdCellAlignVerticalCenter
    Next c

    With tbl.cell(1, 1).Range
        .Text = ""
        .Font.Name = "Arial": .Font.Size = 11
        With .ParagraphFormat
            .Alignment = wdAlignParagraphCenter
            .LineSpacingRule = wdLineSpaceMultiple
            .LineSpacing = 1.01 * 12
        End With
    End With

    With tbl.cell(1, 2).Range
        .Text = placeholderText
        .Font.Name = "Arial": .Font.Size = 15
        With .ParagraphFormat
            .Alignment = wdAlignParagraphCenter
            .SpaceBefore = 14: .SpaceAfter = 16
            .LineSpacingRule = wdLineSpaceMultiple
            .LineSpacing = 1.01 * 12
        End With
    End With

    Set cellRange = tbl.cell(1, 3).Range
    cellRange.Text = "[LOGO: Plataforma Brasil]"
    With cellRange
        .Font.Name = "Arial": .Font.Size = 10
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .ParagraphFormat.LineSpacingRule = wdLineSpaceMultiple
        .ParagraphFormat.LineSpacing = 1.01 * 12
        .ParagraphFormat.LeftIndent = InchesToPoints(0.1)
        .ParagraphFormat.RightIndent = InchesToPoints(0.1)
        .End = .End - 1
        With .Find
            .ClearFormatting
            .Text = "Plataforma Brasil"
            .replacement.ClearFormatting
            .replacement.Font.Bold = True
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .Execute Replace:=wdReplaceOne
        End With
    End With

    hdrRange.Collapse Direction:=wdCollapseEnd
    hdrRange.InsertAfter vbCr
    hdrRange.Collapse Direction:=wdCollapseEnd
    With hdrRange.Paragraphs(1)
        .Range.Font.Name = "Arial": .Range.Font.Size = 11
        .Format.LineSpacingRule = wdLineSpaceSingle
        .Format.SpaceBefore = 0
        .Format.SpaceAfter = 0
        .Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
        .Borders(wdBorderBottom).LineWidth = wdLineWidth025pt
        .Borders(wdBorderBottom).Color = RGB(166, 166, 166)
        .Shading.BackgroundPatternColor = RGB(217, 217, 217)
    End With

    ' ===== 8. Other Pages Header =====
    Set otherHdrRange = doc.Sections(1).Headers(wdHeaderFooterPrimary).Range
    otherHdrRange.Text = ""
    Set tbl = otherHdrRange.Tables.Add(Range:=otherHdrRange, numRows:=1, NumColumns:=3)

    With tbl
        .AllowAutoFit = False
        .PreferredWidthType = wdPreferredWidthPoints
        .Spacing = 0
        .TopPadding = 0: .BottomPadding = 0
        .LeftPadding = 0: .RightPadding = 0
        .Rows.Alignment = wdAlignRowLeft
        .Rows(1).HeightRule = wdRowHeightAuto

        .Columns(1).PreferredWidth = InchesToPoints(1.63)
        .Columns(2).PreferredWidth = InchesToPoints(3.26)
        .Columns(3).PreferredWidth = InchesToPoints(1.81)

        On Error Resume Next
        .Borders(wdBorderInsideVertical).LineStyle = wdLineStyleNone
        .Borders(wdBorderInsideHorizontal).LineStyle = wdLineStyleNone
        On Error GoTo 0

        For c = wdBorderTop To wdBorderRight
            With .Borders(c)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth075pt
                .Color = RGB(166, 166, 166)
            End With
        Next c
    End With

    For c = 1 To 3
        tbl.cell(1, c).VerticalAlignment = wdCellAlignVerticalCenter
    Next c

    With tbl.cell(1, 1).Range
        .Text = ""
        .Font.Name = "Arial": .Font.Size = 11
        With .ParagraphFormat
            .Alignment = wdAlignParagraphCenter
            .LineSpacingRule = wdLineSpaceMultiple
            .LineSpacing = 1.01 * 12
        End With
    End With

    With tbl.cell(1, 2).Range
        .Text = placeholderText
        .Font.Name = "Arial": .Font.Size = 15
        With .ParagraphFormat
            .Alignment = wdAlignParagraphCenter
            .SpaceBefore = 14: .SpaceAfter = 16
            .LineSpacingRule = wdLineSpaceMultiple
            .LineSpacing = 1.01 * 12
        End With
    End With

    Set cellRange = tbl.cell(1, 3).Range
    cellRange.Text = "[LOGO: Plataforma Brasil]"
    With cellRange
        .Font.Name = "Arial": .Font.Size = 10
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .ParagraphFormat.LineSpacingRule = wdLineSpaceMultiple
        .ParagraphFormat.LineSpacing = 1.01 * 12
        .ParagraphFormat.LeftIndent = InchesToPoints(0.1)
        .ParagraphFormat.RightIndent = InchesToPoints(0.1)
        .End = .End - 1
        With .Find
            .ClearFormatting
            .Text = "Plataforma Brasil"
            .replacement.ClearFormatting
            .replacement.Font.Bold = True
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .Execute Replace:=wdReplaceOne
        End With
    End With

    otherHdrRange.Collapse Direction:=wdCollapseEnd
    otherHdrRange.InsertAfter vbCrLf & "Continuação do Parecer: " & parecerNum
    With otherHdrRange.Paragraphs.Last
        .Range.Font.Name = "Arial": .Range.Font.Size = 7
        .Format.SpaceBefore = 8: .Format.SpaceAfter = 18
        .Format.LeftIndent = 0: .Format.RightIndent = 0
        .Alignment = wdAlignParagraphLeft
    End With

    MsgBox "PTBR format applied (header finalized, footer pending).", vbInformation
End Sub


Sub Insert_PTBR_Footer_FromTemplate()
    Dim currentDoc As Document
    Dim tplDoc As Document
    Dim tplPath As String
    Dim footerSource As Range
    Dim sec As Section
    Dim para As Paragraph
    Dim inputDict As Object
    Dim match As Object
    Dim regex As Object
    Dim placeholder As Variant
    Dim userInput As String
    Dim r As Range
    Dim ftrRange As Range

    Set currentDoc = ActiveDocument
    tplPath = ActiveDocument.path & "\Table_footer_PTBR_format1.docx"

    If Dir(tplPath) = "" Then
        MsgBox "Template file not found: " & tplPath, vbCritical
        Exit Sub
    End If

    ' Open template invisibly
    Set tplDoc = Documents.Open(FileName:=tplPath, Visible:=False)
    Set footerSource = tplDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range
    footerSource.Copy
    tplDoc.Close SaveChanges:=False

    ' Apply footer to all sections (first and other pages)
    For Each sec In currentDoc.Sections
        ' --- Primary footer ---
        Set ftrRange = sec.Footers(wdHeaderFooterPrimary).Range
        ftrRange.Text = ""
        ftrRange.Paste
        ftrRange.Characters.Last.Delete ' Remove extra paragraph after table

        ' --- First page footer (if exists) ---
        If sec.Footers(wdHeaderFooterFirstPage).Exists Then
            Set ftrRange = sec.Footers(wdHeaderFooterFirstPage).Range
            ftrRange.Text = ""
            ftrRange.Paste
            ftrRange.Characters.Last.Delete
        End If
    Next sec

    ' === Prepare placeholder replacement ===
    Set inputDict = CreateObject("Scripting.Dictionary")
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.pattern = "<<(.*?)>>"

    ' Collect placeholder values from user
    For Each sec In currentDoc.Sections
        ' Check primary footer
        For Each para In sec.Footers(wdHeaderFooterPrimary).Range.Paragraphs
            Set r = para.Range
            If regex.test(r.Text) Then
                For Each match In regex.Execute(r.Text)
                    placeholder = match.SubMatches(0)
                    If Not inputDict.Exists(placeholder) Then
                        userInput = InputBox("Enter value for: " & placeholder, "Replace Placeholder", "")
                        inputDict.Add placeholder, userInput
                    End If
                Next
            End If
        Next para

        ' Check first page footer
        If sec.Footers(wdHeaderFooterFirstPage).Exists Then
            For Each para In sec.Footers(wdHeaderFooterFirstPage).Range.Paragraphs
                Set r = para.Range
                If regex.test(r.Text) Then
                    For Each match In regex.Execute(r.Text)
                        placeholder = match.SubMatches(0)
                        If Not inputDict.Exists(placeholder) Then
                            userInput = InputBox("Enter value for: " & placeholder, "Replace Placeholder", "")
                            inputDict.Add placeholder, userInput
                        End If
                    Next
                End If
            Next para
        End If
    Next sec

    ' Replace placeholders in both footers
    For Each placeholder In inputDict.Keys
        For Each sec In currentDoc.Sections
            ' Primary
            With sec.Footers(wdHeaderFooterPrimary).Range.Find
                .ClearFormatting
                .replacement.ClearFormatting
                .Text = "<<" & placeholder & ">>"
                .replacement.Text = inputDict(placeholder)
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceAll
            End With
            ' First Page
            If sec.Footers(wdHeaderFooterFirstPage).Exists Then
                With sec.Footers(wdHeaderFooterFirstPage).Range.Find
                    .ClearFormatting
                    .replacement.ClearFormatting
                    .Text = "<<" & placeholder & ">>"
                    .replacement.Text = inputDict(placeholder)
                    .Wrap = wdFindContinue
                    .Execute Replace:=wdReplaceAll
                End With
            End If
        Next sec
    Next

    MsgBox "? Footer inserted and updated across all sections.", vbInformation
End Sub

