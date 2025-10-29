Attribute VB_Name = "NewMacros"
Sub Table_Clean()
Attribute Table_Clean.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Table_Clean"
'
' Table_Clean Macro
'
'
    With Selection.Tables(1)
        .TopPadding = InchesToPoints(0)
        .BottomPadding = InchesToPoints(0)
        .LeftPadding = InchesToPoints(0)
        .RightPadding = InchesToPoints(0)
        .Spacing = 0
        .AllowPageBreaks = True
        .AllowAutoFit = True
    End With
    With Selection.Tables(1)
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth050pt
        .DefaultBorderColor = wdColorAutomatic
    End With
    Selection.Rows.HeightRule = wdRowHeightAtLeast
    Selection.Rows.Height = InchesToPoints(0.04)
End Sub
Sub Remove_Section_Break()
Attribute Remove_Section_Break.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Remove_Section_Break"
'
' Remove_Section_Break Macro
    Selection.Find.ClearFormatting
    Selection.Find.replacement.ClearFormatting
    With Selection.Find
        .Text = "[^12^m^n]"
        .replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchByte = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub Remove_Text_BG()
Attribute Remove_Text_BG.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Remove_Text_BG"
'
' Remove_Text_BG Macro
'
'
    With Selection.Font
        With .Shading
            .Texture = wdTextureNone
            .ForegroundPatternColor = wdColorAutomatic
            .BackgroundPatternColor = wdColorAutomatic
        End With
        .Borders(1).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth050pt
        .DefaultBorderColor = wdColorAutomatic
    End With
    With Selection.ParagraphFormat.Shading
        .Texture = wdTextureNone
        .ForegroundPatternColor = wdColorAutomatic
        .BackgroundPatternColor = wdColorAutomatic
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth050pt
        .DefaultBorderColor = wdColorAutomatic
    End With
End Sub
Sub Highlight_Hidden()
Attribute Highlight_Hidden.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Highlight_Hidden"
'
' Highlight_Hidden Macro
'
'
    With Selection.Font
        .Hidden = True
    End With
    Options.DefaultHighlightColorIndex = wdYellow
    Selection.Range.HighlightColorIndex = wdYellow
End Sub
Sub Highlight_Hidden_Text()
Attribute Highlight_Hidden_Text.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Highlight_Hidden_Text"
'
' Highlight_Hidden_Text Macro
'
'
    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Font.Hidden = True
    Selection.Find.replacement.ClearFormatting
    Selection.Find.replacement.Highlight = True
    With Selection.Find
        .Text = ""
        .replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchByte = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub Remove_Highlight_Hidden_Text()
'

'
'
    Selection.WholeStory
    'Selection.Find.ClearFormatting
    'Selection.Find.Font.Hidden = True
    Selection.Find.replacement.ClearFormatting
    Selection.Find.replacement.Highlight = False
    With Selection.Find
        .Text = ""
        .replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchByte = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub Format_Tab()
Attribute Format_Tab.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Format_Tab"
'
' Format_Tab Macro
'
'
    Selection.Rows.HeightRule = wdRowHeightAtLeast
    Selection.Rows.Height = InchesToPoints(0.04)
    Selection.Shading.Texture = wdTextureNone
    Selection.Shading.ForegroundPatternColor = wdColorAutomatic
    Selection.Shading.BackgroundPatternColor = wdColorAutomatic
End Sub
Sub CheckParagraphMarks()
    Dim doc As Document
    Set doc = ActiveDocument
    
    Dim i As Long
    Dim totalParagraphMarks As Long
    
    totalParagraphMarks = 0
    
    ' Loop through each character in the document
    For i = 1 To doc.Characters.count
        ' Check if the character is a paragraph mark
        If doc.Characters(i).Text = vbCrLf Then ' vbCrLf represents a paragraph mark in Word
            totalParagraphMarks = totalParagraphMarks + 1
        End If
    Next i
    
    ' Display the total count of paragraph marks found
    MsgBox "Total paragraph marks in the document: " & totalParagraphMarks, vbInformation
End Sub
Sub temp()
Attribute temp.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.temp"
'
' temp Macro
'
'
    Selection.Tables(1).Rows.LeftIndent = InchesToPoints(-0.08)
    With Selection.Tables(1)
        .TopPadding = InchesToPoints(0)
        .BottomPadding = InchesToPoints(0)
        .LeftPadding = InchesToPoints(0.07)
        .RightPadding = InchesToPoints(0.07)
        .Spacing = 0
        .AllowPageBreaks = True
        .AllowAutoFit = False
    End With
    ActiveDocument.Save
End Sub

Sub ReplaceNumberedListText()
    Dim para As Paragraph
    Dim listFormat As listFormat
    
    ' Loop through all paragraphs in the document
    For Each para In ActiveDocument.Paragraphs
        Set listFormat = para.Range.listFormat
        
        ' Check if the paragraph is part of a numbered list
        If listFormat.ListType = wdListSimpleNumbering Then
            ' Replace "Note:" with "Oponba:" in the numbering format
            listFormat.ListTemplate.ListLevels(1).NumberFormat = Replace(listFormat.ListTemplate.ListLevels(1).NumberFormat, "Note:", "Oponba:")
        End If
    Next para
End Sub
Sub Portrait()
Attribute Portrait.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Portrait"
'
' Portrait Macro
'
'
    If Selection.PageSetup.Orientation = wdOrientPortrait Then
        Selection.PageSetup.Orientation = wdOrientLandscape
    Else
        Selection.PageSetup.Orientation = wdOrientPortrait
    End If
End Sub
Sub Margin()
Attribute Margin.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Margin"
'
' Margin Macro
'
'
    Selection.WholeStory
    With ActiveDocument.Styles(wdStyleNormal).Font
        If .NameFarEast = .NameAscii Then
            .NameAscii = ""
        End If
        .NameFarEast = ""
    End With
    With ActiveDocument.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientPortrait
        .TopMargin = InchesToPoints(0.79)
        .BottomMargin = InchesToPoints(0.79)
        .LeftMargin = InchesToPoints(0.98)
        .RightMargin = InchesToPoints(0.98)
        .Gutter = InchesToPoints(0)
        .HeaderDistance = InchesToPoints(0.51)
        .FooterDistance = InchesToPoints(0.51)
        .PageWidth = InchesToPoints(8.27)
        .PageHeight = InchesToPoints(11.69)
        .FirstPageTray = wdPrinterDefaultBin
        .OtherPagesTray = wdPrinterDefaultBin
        .SectionStart = wdSectionNewPage
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = True
        .VerticalAlignment = wdAlignVerticalTop
        .SuppressEndnotes = False
        .MirrorMargins = False
        .TwoPagesOnOne = False
        .BookFoldPrinting = False
        .BookFoldRevPrinting = False
        .BookFoldPrintingSheets = 1
        .GutterPos = wdGutterPosLeft
        .SectionDirection = wdSectionDirectionLtr
    End With
End Sub
Sub Font()
Attribute Font.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Font"
'
' Font Macro
'
'
    Selection.WholeStory
    Selection.Font.Name = "Times New Roman"
    Selection.Font.Size = 11
    With Selection.Font
        .Name = ""
        .Bold = False
        .Italic = False
        .Spacing = 0
        .Scaling = 100
        .Position = 0
        .NameBi = ""
        .BoldBi = False
        .ItalicBi = False
    End With
    With Selection.Font
        .Name = ""
        .Color = -587137025
        .NameBi = ""
    End With
End Sub
Sub Line_Spacing()
Attribute Line_Spacing.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Line_Spacing"
'
' Line_Spacing Macro
'
'
    With Selection.ParagraphFormat
        .LeftIndent = InchesToPoints(0)
        .RightIndent = InchesToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .FirstLineIndent = InchesToPoints(0)
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
End Sub
Sub reset_space()
Attribute reset_space.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.reset_space"
'
' reset_space Macro
'
'
    With Selection.ParagraphFormat
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
    End With
End Sub
Sub before_space_up()
Attribute before_space_up.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.before_space_up"
'
' before_space_up Macro
'
'
    With Selection.ParagraphFormat
        .SpaceBefore = .SpaceBefore + 1
        .SpaceBeforeAuto = False
    End With
End Sub
Sub before_space_DOWN()
'
' before_space_DOWN Macro
'
'
    With Selection.ParagraphFormat
        .SpaceBefore = .SpaceBefore - 1
        .SpaceBeforeAuto = False
    End With
End Sub
Sub after_space_DOWN()
'
' after_space_DOWN Macro
'
'
    With Selection.ParagraphFormat
        .SpaceAfter = .SpaceAfter - 1
        .SpaceAfterAuto = False
    End With
End Sub
Sub after_space_UP()
'
' after_space_UP Macro
'
'
    With Selection.ParagraphFormat
        .SpaceAfter = .SpaceAfter + 1
        .SpaceAfterAuto = False
    End With
End Sub
Sub left_Indent()
'
' left_Indent Macro
'
'
    With Selection.Layo
        
    End With
End Sub
Sub ConfigureImageLayoutSettings()
'
' ConfigureImageLayoutSettings Macro
'
'
    Dim shp As Shape
    Dim inShp As inlineShape
    
    ' Loop through all inline shapes (pictures in line with text)
    For Each inShp In ActiveDocument.InlineShapes
        ' Convert inline shape to a floating shape to allow position and wrapping adjustments
        Set shp = inShp.ConvertToShape
        
        ' Set position
        shp.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
        shp.Left = wdShapeCenter ' Centered relative to the page
        shp.RelativeVerticalPosition = wdRelativeVerticalPositionPage
        shp.Top = 0 ' Top alignment relative to the page
        
        ' Enable lock anchor
        shp.LockAnchor = True
        
        ' Set text wrapping style
        shp.WrapFormat.Type = wdWrapBehind
        
        ' Reset size to original
        shp.LockAspectRatio = msoTrue ' Maintain aspect ratio
        shp.ScaleHeight = 100  ' Reset height to 100%
        shp.ScaleWidth = 100 ' Reset width to 100%
    Next inShp
    
    ' Loop through all floating shapes (already set as shapes)
    For Each shp In ActiveDocument.Shapes
        ' Ensure it's a picture
        If shp.Type = msoPicture Then
            ' Set position
            shp.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
            shp.Left = wdShapeCenter ' Centered relative to the page
            shp.RelativeVerticalPosition = wdRelativeVerticalPositionPage
            shp.Top = 0 ' Top alignment relative to the page
            
            ' Enable lock anchor
            shp.LockAnchor = True
            
            ' Set text wrapping style
            shp.WrapFormat.Type = wdWrapBehind
            
            ' Reset size to original
            shp.LockAspectRatio = msoTrue ' Maintain aspect ratio
            shp.ScaleHeight = 100 ' Reset height to 100%
            shp.ScaleWidth = 100 ' Reset width to 100%
        End If
    Next shp
    
    MsgBox "Image layout settings applied successfully!", vbInformation
End Sub
Sub Enter()
'
' Enter Macro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.replacement.ClearFormatting
    With Selection.Find
        .Text = "^p^p"
        .replacement.Text = " ^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub


Sub IncreaseIndent()
    With Selection.ParagraphFormat
        .LeftIndent = .LeftIndent + InchesToPoints(0.01) ' Increase by 0.25 inch
    End With
End Sub


Sub DecreaseIndent()
    With Selection.ParagraphFormat
        If .LeftIndent >= InchesToPoints(0.01) Then
            .LeftIndent = .LeftIndent - InchesToPoints(0.01)
        Else
            .LeftIndent = 0
        End If
    End With
End Sub
Sub Tab_008_0()
Attribute Tab_008_0.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Tab_008_0"
'
' Tab_008_0 Macro
'
'
    Selection.Tables(1).Rows.LeftIndent = InchesToPoints(0)
    With Selection.Tables(1)
        .TopPadding = InchesToPoints(0)
        .BottomPadding = InchesToPoints(0)
        .LeftPadding = InchesToPoints(0.08)
        .RightPadding = InchesToPoints(0.08)
        .Spacing = 0
        .AllowPageBreaks = True
        .AllowAutoFit = False
    End With
    With Selection.ParagraphFormat
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .LineUnitBefore = 0
    End With
    Selection.Rows.HeightRule = wdRowHeightAtLeast
    Selection.Rows.Height = InchesToPoints(0.04)
End Sub
Sub SetColumnWidthsWithDefaults()
    Dim tbl As Table
    Dim colWidths(1 To 10) As Single
    Dim i As Integer
    Dim inputVal As String
    Dim defaultVal As String
    Dim numCols As Integer

    ' Check if selection includes a table
    If Selection.Information(wdWithInTable) = False Then
        MsgBox "Please select at least one table.", vbExclamation
        Exit Sub
    End If

    ' Use the first selected table to get column count
    Set tbl = Selection.Tables(1)
    numCols = tbl.Columns.count
    If numCols > 10 Then
        MsgBox "This macro supports up to 10 columns only.", vbExclamation
        Exit Sub
    End If

    ' Get stored defaults (if any), prompt for input
    For i = 1 To numCols
        On Error Resume Next
        defaultVal = ActiveDocument.Variables("ColWidth" & i).Value
        On Error GoTo 0
        If defaultVal = "" Then defaultVal = "1" ' Default to 1 inch if nothing stored

        inputVal = InputBox("Enter width (in inches) for Column " & i, "Set Column Width", defaultVal)
        If IsNumeric(inputVal) Then
            colWidths(i) = CSng(inputVal) * 72 ' Convert to points
            ActiveDocument.Variables("ColWidth" & i).Value = inputVal ' Save for next time
        Else
            MsgBox "Invalid input for column " & i & ". Please enter a number.", vbCritical
            Exit Sub
        End If
    Next i

    ' Apply widths to all selected tables
    For Each tbl In Selection.Tables
        numCols = tbl.Columns.count
        For i = 1 To numCols
            tbl.Columns(i).Width = colWidths(i)
        Next i
    Next tbl

    MsgBox "Column widths applied and saved as defaults.", vbInformation
End Sub

Sub PTBR_Table()
    ' Show the userform
    frmPTBR_Table.Show

    ' Exit if form is unloaded (Cancel pressed)
    If Not frmPTBR_Table.Visible Then Exit Sub

    With frmPTBR_Table
        If .chkMerge.Value Then Call MergeBasedOnAceito
        If .chkClean.Value Then Call CleanBreaksInSelectedTableOnly
        If .chkFormat.Value Then Call FormatSelectedTable
    End With

    Unload frmPTBR_Table
End Sub


Sub MergeBasedOnAceito()
    Dim tbl As Table
    Dim rowIdx As Long, colIdx As Long
    Dim maxRows As Long, maxCols As Long
    Dim startRow As Long, keywordCol As Long
    Dim keyword As String
    Dim nextKeyRow As Long
    Dim cellText As String
    Dim i As Long
    Dim changesMade As Boolean
    Dim inputKeyword As String, inputCol As String

    ' ==== INPUT PROMPTS ====
    inputKeyword = InputBox("Enter keyword to trigger merging (default: Aceito):", "Keyword", "Aceito")
    If Trim(inputKeyword) = "" Then
        MsgBox "Cancelled.", vbExclamation
        Exit Sub
    End If
    keyword = LCase(Trim(inputKeyword))

    inputCol = InputBox("Enter column number to check for keyword (default: last column):", "Column Number", "")
    
    ' ==== VALIDATE TABLE ====
    If ActiveDocument.Tables.count = 0 Then
        MsgBox "No tables found in the document.", vbExclamation
        Exit Sub
    End If

    Set tbl = ActiveDocument.Tables(1)
    maxCols = tbl.Columns.count
    maxRows = tbl.Rows.count

    ' ==== SET COLUMN ====
    If IsNumeric(inputCol) And Val(inputCol) >= 1 And Val(inputCol) <= maxCols Then
        keywordCol = CLng(inputCol)
    Else
        keywordCol = maxCols ' default to last column
    End If

    ' ==== LOOP UNTIL FULL MERGE ====
    Do
        changesMade = False
        maxRows = tbl.Rows.count ' refresh if merged
        rowIdx = 1

        Do While rowIdx <= maxRows
            cellText = CleanCellText(tbl.cell(rowIdx, keywordCol).Range.Text)

            If LCase(cellText) = keyword Then
                startRow = rowIdx
                nextKeyRow = maxRows + 1

                ' Find next keyword row
                For i = rowIdx + 1 To maxRows
                    If LCase(CleanCellText(tbl.cell(i, keywordCol).Range.Text)) = keyword Then
                        nextKeyRow = i
                        Exit For
                    End If
                Next i

                ' Merge from startRow to row before next keyword
                If nextKeyRow - 1 > startRow Then
                    For colIdx = 1 To maxCols
                        On Error Resume Next
                        tbl.cell(startRow, colIdx).Merge tbl.cell(nextKeyRow - 1, colIdx)
                        On Error GoTo 0
                        changesMade = True
                    Next colIdx
                End If

                rowIdx = nextKeyRow
            Else
                rowIdx = rowIdx + 1
            End If
        Loop

    Loop While changesMade

    MsgBox "Merging complete using keyword '" & inputKeyword & "' in column " & keywordCol & ".", vbInformation
    tbl.Range.Select
End Sub

Function CleanCellText(txt As String) As String
    txt = Replace(txt, Chr(13), "")
    txt = Replace(txt, Chr(7), "")
    CleanCellText = Trim(txt)
End Function

Sub CleanBreaksInSelectedTableOnly()
    Dim tbl As Table
    Dim rng As Range
    Dim i As Long

    ' Ensure cursor is inside a table
    If Not Selection.Information(wdWithInTable) Then
        MsgBox "Please place the cursor inside a table.", vbExclamation
        Exit Sub
    End If

    Set tbl = Selection.Tables(1)

    ' Step 1: Clean the whole table first (^l, ^p, double spaces)
    Set rng = tbl.Range
    rng.End = rng.End - 1 ' remove end-of-table marker

    ' Replace ^l with space
    With rng.Find
        .ClearFormatting
        .replacement.ClearFormatting
        .Text = "^l"
        .replacement.Text = " "
        .Forward = True
        .Wrap = wdFindStop
        .Execute Replace:=wdReplaceAll
    End With

    ' Replace ^p with space
    With rng.Find
        .ClearFormatting
        .replacement.ClearFormatting
        .Text = "^p"
        .replacement.Text = " "
        .Forward = True
        .Wrap = wdFindStop
        .Execute Replace:=wdReplaceAll
    End With

    ' Replace multiple spaces with single space
    With rng.Find
        .ClearFormatting
        .replacement.ClearFormatting
        .Text = "[ ]{2,}"
        .replacement.Text = " "
        .MatchWildcards = True
        .Wrap = wdFindStop
        .Execute Replace:=wdReplaceAll
    End With

    ' Step 2: Remove all spaces from 2nd column (excluding header row)
    For i = 2 To tbl.Rows.count ' start from row 2 (skip header)
        On Error Resume Next ' skip if cell missing
        Set rng = tbl.cell(i, 2).Range
        rng.End = rng.End - 1 ' exclude end-of-cell marker

        With rng.Find
            .ClearFormatting
            .replacement.ClearFormatting
            .Text = " "
            .replacement.Text = ""
            .Wrap = wdFindStop
            .Execute Replace:=wdReplaceAll
        End With
    Next i

    MsgBox "Table cleaned. All spaces removed from column 2 (excluding header).", vbInformation
    tbl.Range.Select
End Sub

Sub FormatSelectedTable()
    Dim tbl As Table
    Dim rowIdx As Long, colIdx As Long
    Dim cel As cell
    Dim rng As Range

    ' Ensure cursor is in a table
    If Not Selection.Information(wdWithInTable) Then
        MsgBox "Please place the cursor inside a table.", vbExclamation
        Exit Sub
    End If

    Set tbl = Selection.Tables(1)

    ' Loop through each cell
    For rowIdx = 1 To tbl.Rows.count
        For colIdx = 1 To tbl.Columns.count
            Set cel = tbl.cell(rowIdx, colIdx)
            Set rng = cel.Range
            rng.End = rng.End - 1 ' Exclude end-of-cell marker

            ' 1. Font formatting
            With rng.Font
                .Name = "Arial"
                .Size = 10
            End With

            ' 2. Paragraph formatting
            With rng.ParagraphFormat
                .LeftIndent = 0
                .RightIndent = 0
                .SpaceBefore = 0
                .SpaceAfter = 0
                .LineSpacingRule = wdLineSpaceMultiple
                .LineSpacing = 0.97 * 12 ' 0.97 lines (Word expects points)
                .Alignment = wdAlignParagraphLeft ' Set default as Left
            End With

            ' 3. Header formatting (row 1)
            If rowIdx = 1 Then
                rng.ParagraphFormat.Alignment = wdAlignParagraphCenter
                rng.ParagraphFormat.RightIndent = InchesToPoints(0.04)
            Else
                ' 4. Align columns for other rows
                Select Case colIdx
                    Case 1, 2, 4
                        rng.ParagraphFormat.Alignment = wdAlignParagraphLeft
                    Case 3, 5
                        rng.ParagraphFormat.Alignment = wdAlignParagraphCenter
                End Select
            End If
        Next colIdx
    Next rowIdx

    ' 5. Set row height and cell margins
    With tbl
        For rowIdx = 1 To .Rows.count
            With .Rows(rowIdx)
                .HeightRule = wdRowHeightAtLeast
                .Height = InchesToPoints(0.04)
            End With
        Next rowIdx

        ' Set uniform cell margins
        .TopPadding = 0
        .BottomPadding = 0
        .RightPadding = 0
        .LeftPadding = InchesToPoints(0.04)
    End With

    MsgBox "Formatting applied successfully to selected table.", vbInformation
End Sub




Sub MergeSelectedTablesOnly()
    Dim selRange As Range
    Dim tbl As Table
    Dim tblList As Collection
    Dim i As Long
    Dim rngBetween As Range

    Set selRange = Selection.Range
    Set tblList = New Collection

    ' Collect tables that fall within the selection
    For Each tbl In ActiveDocument.Tables
        If tbl.Range.Start >= selRange.Start And tbl.Range.End <= selRange.End Then
            tblList.Add tbl
        End If
    Next tbl

    ' Work backwards to merge and delete in-between content
    For i = tblList.count To 2 Step -1
        Dim tblCurrent As Table
        Dim tblPrev As Table
        Set tblCurrent = tblList(i)
        Set tblPrev = tblList(i - 1)

        ' Delete any text between tables
        Set rngBetween = ActiveDocument.Range(tblPrev.Range.End, tblCurrent.Range.Start)
        rngBetween.Delete

        ' Merge tables: move rows from current to previous, then delete current
        Do While tblCurrent.Rows.count > 0
            tblCurrent.Rows(1).Range.Cut
            tblPrev.Rows.Last.Range.InsertAfter vbCr
            tblPrev.Rows.Last.Next.Range.Paste
        Loop
        tblCurrent.Delete
    Next i
End Sub

Sub FlipRTLTableWithFormatting()
    Dim tbl As Table, newTbl As Table
    Dim tempRange As Range
    Dim rowCount As Long, colCount As Long
    Dim i As Long, j As Long
    Dim colWidths() As Single
    Dim srcRange As Range, dstRange As Range

    ' Step 1: Ensure the cursor is inside a table
    If Not Selection.Information(wdWithInTable) Then
        MsgBox "Place the cursor inside a table.", vbExclamation
        Exit Sub
    End If

    Set tbl = Selection.Tables(1)
    rowCount = tbl.Rows.count
    colCount = tbl.Columns.count

    ' Step 2: Store column widths
    ReDim colWidths(1 To colCount)
    For j = 1 To colCount
        colWidths(j) = tbl.Columns(j).Width
    Next j

    ' Step 3: Safely insert a temp range after the table
    Set tempRange = tbl.Range.Duplicate
    tempRange.Collapse Direction:=wdCollapseEnd
    tempRange.InsertAfter vbCrLf & vbCrLf
    tempRange.Collapse Direction:=wdCollapseEnd

    ' Step 4: Insert a new table in the temp range
    Set newTbl = ActiveDocument.Tables.Add(Range:=tempRange, numRows:=rowCount, NumColumns:=colCount)
    newTbl.TableDirection = wdTableDirectionRtl

    ' Step 5: Apply reversed column widths
    For j = 1 To colCount
        On Error Resume Next
        newTbl.Columns(j).SetWidth colWidths(colCount - j + 1), wdAdjustNone
        On Error GoTo 0
    Next j

    ' Step 6: Copy formatted content cell-by-cell (including merged header row)
    For i = 1 To rowCount
        For j = 1 To colCount
            On Error Resume Next
            If tbl.cell(i, colCount - j + 1).Range.Cells.count = 1 Then
                Set srcRange = tbl.cell(i, colCount - j + 1).Range
                Set dstRange = newTbl.cell(i, j).Range
                srcRange.End = srcRange.End - 1
                dstRange.End = dstRange.End - 1
                dstRange.FormattedText = srcRange.FormattedText
            End If
            On Error GoTo 0
        Next j
    Next i

    ' Step 7: Recreate merged top row if applicable (safe logic)
    If tbl.Rows(1).Cells.count = 1 And colCount > 1 Then
        newTbl.cell(1, 1).Merge newTbl.cell(1, colCount)
    End If

    ' Step 8: Optionally delete original table
    If MsgBox("Delete original table?", vbYesNo + vbQuestion, "Confirm") = vbYes Then
        tbl.Delete
    End If

    MsgBox "? Table flipped with RTL direction, column order reversed, formatting and merged headers preserved.", vbInformation
End Sub

Sub FlipFullTableIntoOne()
    Dim origDoc As Document, tempDoc As Document
    Dim origTbl As Table, rowTbl As Table, finalTbl As Table
    Dim i As Long, j As Long, rowCount As Long, colCount As Long
    Dim colWidths() As Single
    Dim srcRange As Range, dstRange As Range
    Dim tempRange As Range
    Dim firstRow As Boolean

    If Not Selection.Information(wdWithInTable) Then
        MsgBox "Please place the cursor inside a table.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone

    ' Copy table to temp document
    Set origDoc = ActiveDocument
    Set origTbl = Selection.Tables(1)
    origTbl.Range.Copy

    Set tempDoc = Documents.Add(Visible:=False)
    tempDoc.Range.Paste
    Set origTbl = tempDoc.Tables(1)
    rowCount = origTbl.Rows.count

    firstRow = True

    ' Loop through each row of the original table
    For i = 1 To rowCount
        On Error Resume Next
        origTbl.Rows(i).Range.Copy
        tempDoc.Range.Collapse wdCollapseEnd
        tempDoc.Range.Paste
        Set rowTbl = tempDoc.Tables(tempDoc.Tables.count)
        If rowTbl.Rows.count = 0 Then GoTo SkipRow

        colCount = rowTbl.Columns.count
        If colCount = 0 Then GoTo SkipRow

        ReDim colWidths(1 To colCount)
        For j = 1 To colCount
            colWidths(j) = rowTbl.cell(1, j).Width
        Next j

        ' First flipped row creates the master final table
        If firstRow Then
            Set tempRange = tempDoc.Range
            tempRange.Collapse wdCollapseEnd
            Set finalTbl = tempDoc.Tables.Add(Range:=tempRange, numRows:=1, NumColumns:=colCount)
            finalTbl.TableDirection = wdTableDirectionRtl
            finalTbl.AutoFitBehavior (wdAutoFitFixed)
            firstRow = False
        Else
            finalTbl.Rows.Add
        End If

        ' Fill the last row in finalTbl with reversed content
        For j = 1 To colCount
            Set srcRange = rowTbl.cell(1, colCount - j + 1).Range
            srcRange.End = srcRange.End - 1
            Set dstRange = finalTbl.cell(finalTbl.Rows.count, j).Range
            dstRange.End = dstRange.End - 1
            dstRange.FormattedText = srcRange.FormattedText
            dstRange.ParagraphFormat.ReadingOrder = wdReadingOrderRtl
            finalTbl.cell(finalTbl.Rows.count, j).Width = colWidths(colCount - j + 1)
        Next j

        ' Delete temporary row table
        rowTbl.Delete

SkipRow:
        On Error GoTo 0
    Next i

    ' Replace original table with final flipped table
    tempDoc.Content.Copy
    origDoc.Activate
    origTbl.Range.Delete
    Selection.Range.Paste

    tempDoc.Close SaveChanges:=wdDoNotSaveChanges
    Application.ScreenUpdating = True
    Application.DisplayAlerts = wdAlertsAll

    MsgBox "? All rows flipped and merged into one RTL table.", vbInformation
End Sub

Sub InsertInlinePictureWithTransparency()
    Dim dlgOpen As FileDialog
    Dim selectedFile As String
    Dim inlineShape As inlineShape
    Dim shp As Shape

    ' Open file dialog to select the picture
    Set dlgOpen = Application.FileDialog(msoFileDialogFilePicker)
    With dlgOpen
        .Title = "Select a Picture"
        .Filters.Clear
        .Filters.Add "Images", "*.jpg;*.jpeg;*.png;*.bmp;*.gif"
        If .Show <> -1 Then Exit Sub ' Cancelled
        selectedFile = .SelectedItems(1)
    End With

    ' Insert image inline at cursor position
    Set inlineShape = Selection.InlineShapes.AddPicture(FileName:=selectedFile, LinkToFile:=False, SaveWithDocument:=True)

    ' Convert InlineShape to Shape to set transparency
    Set shp = inlineShape.ConvertToShape
    shp.WrapFormat.Type = wdWrapInline ' keep it inline
    'shp.PictureFormat.TransparentBackground = msoFalse
    'shp.Fill.Transparency = 0.3 ' 30% transparency
End Sub

Sub ListMacrosAndAssignShortcut()
    Dim vbComp As Object
    Dim codeMod As Object
    Dim macroName As String
    Dim macros As Collection
    Dim i As Long
    Dim line As Long, totalLines As Long
    Dim codeLine As String
    Dim userMacro As String, shortcutKey As String

    Set macros = New Collection

    ' Only scan Normal.dotm
    Set vbComp = Nothing
    For Each vbComp In NormalTemplate.VBProject.VBComponents
        If vbComp.Type = 1 Then ' Standard Module
            Set codeMod = vbComp.CodeModule
            totalLines = codeMod.CountOfLines
            line = 1
            Do While line <= totalLines
                codeLine = Trim(codeMod.Lines(line, 1))
                If Left(codeLine, 4) = "Sub " Then
                    macroName = Split(Split(codeLine, "Sub ")(1), "(")(0)
                    macros.Add macroName
                End If
                line = line + 1
            Loop
        End If
    Next vbComp

    ' Show list
    Dim macroList As String
    If macros.count = 0 Then
        MsgBox "No macros found in Normal.dotm.", vbExclamation
        Exit Sub
    End If

    macroList = "Available macros in Normal.dotm:" & vbCrLf & vbCrLf
    For i = 1 To macros.count
        macroList = macroList & i & ". " & macros(i) & vbCrLf
    Next i

    userMacro = InputBox(macroList & vbCrLf & "Enter the exact macro name to assign shortcut:", "Select Macro")
    If userMacro = "" Then Exit Sub

    shortcutKey = InputBox("Enter the keyboard shortcut (e.g., ^g for Ctrl+G, ^+g for Ctrl+Shift+G):", "Assign Shortcut")
    If shortcutKey = "" Then Exit Sub

    ' Assign shortcut
    On Error Resume Next
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
        Command:=userMacro, KeyCode:=BuildKeyCodeFromString(shortcutKey)
    On Error GoTo 0

    MsgBox "Shortcut '" & shortcutKey & "' assigned to macro '" & userMacro & "'.", vbInformation
End Sub

Function BuildKeyCodeFromString(shortcut As String) As Long
    Dim ctrl As Boolean, Shift As Boolean, alt As Boolean
    Dim ch As String, KeyCode As Long
    ch = Right(shortcut, 1)
    ctrl = InStr(shortcut, "^") > 0
    Shift = InStr(shortcut, "+") > 0
    alt = InStr(shortcut, "%") > 0

    KeyCode = Asc(UCase(ch))
    BuildKeyCodeFromString = BuildKeyCode(KeyCode, _
        IIf(Shift, wdKeyShift, 0) + _
        IIf(ctrl, wdKeyControl, 0) + _
        IIf(alt, wdKeyAlt, 0))
End Function


Sub FlipSelectedTableColumns()
    Dim tbl As Table
    Dim keyColIndex As Long
    Dim colCount As Long
    Dim colWidths() As Single
    Dim i As Long, r As Long, c As Long

    ' Ensure selection is inside a table
    If Not Selection.Information(wdWithInTable) Then
        MsgBox "Please place the cursor inside a table.", vbExclamation
        Exit Sub
    End If

    Set tbl = Selection.Tables(1)
    colCount = tbl.Columns.count

    ' Store column widths
    ReDim colWidths(1 To colCount)
    For i = 1 To colCount
        colWidths(i) = tbl.Columns(i).Width
    Next i

    Application.ScreenUpdating = False

    ' Flip logic: fix key column, move others before it
    keyColIndex = 1
    Do While keyColIndex < tbl.Columns.count
        tbl.Columns(tbl.Columns.count).Select
        Selection.Cut

        tbl.Columns(keyColIndex).Select
        Selection.Paste

        keyColIndex = keyColIndex + 1
    Loop

    ' Re-apply stored widths in reversed order
    For i = 1 To colCount
        tbl.Columns(i).Width = colWidths(colCount - i + 1)
    Next i

    ' ? Set table direction to Right-to-Left
    Selection.Tables(1).TableDirection = wdTableDirectionRtl

    ' ? Apply cell formatting
    For r = 1 To tbl.Rows.count
        With tbl.Rows(r)
            .HeightRule = wdRowHeightAtLeast
            .Height = 3
        End With

        For c = 1 To colCount
            With tbl.cell(r, c)
                ' Set cell padding
                .LeftPadding = InchesToPoints(0.08)
                .RightPadding = InchesToPoints(0.08)

                ' ? Set RTL text direction first
                .Range.ParagraphFormat.ReadingOrder = wdReadingOrderRtl

                ' ? Then set alignment
                If c = 1 Then
                    .Range.ParagraphFormat.Alignment = wdAlignParagraphRight
                Else
                    .Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                End If

                ' Set space before
                .Range.ParagraphFormat.SpaceBefore = 1
            End With
        Next c
    Next r

    Application.ScreenUpdating = True
End Sub

Sub ShowShortcutForm()
    frmShortcutManager.Show
End Sub


Sub AssignMultipleShortcuts()
    Dim successCount As Integer: successCount = 0
    CustomizationContext = NormalTemplate
    On Error GoTo Failed

    ' === Macro 1 ===
    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
                    Command:="before_space_up", _
                    KeyCode:=Application.BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyUp)
    successCount = successCount + 1

    ' === Macro 2 ===
    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
                    Command:="before_space_DOWN", _
                    KeyCode:=Application.BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyDown)
    successCount = successCount + 1

    ' === Macro 3 ===
    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
                    Command:="after_space_DOWN", _
                    KeyCode:=Application.BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyLeft)
    successCount = successCount + 1

    ' === Macro 4 ===
    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
                    Command:="after_space_UP", _
                    KeyCode:=Application.BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyRight)
    successCount = successCount + 1

    ' === Macro 5 ===
    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
                    Command:="Tab_008_0", _
                    KeyCode:=Application.BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyShift, wdKeyT)
    successCount = successCount + 1

    ' === Macro 6 ===
    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
                    Command:="InsertInlinePictureWithTransparency", _
                    KeyCode:=Application.BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyShift, wdKeyD)
    successCount = successCount + 1

    ' === Macro 7 ===
    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
                    Command:="IncrementLeading", _
                    KeyCode:=Application.BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyShift, wdKeyUp)
    successCount = successCount + 1

    ' === Macro 8 ===
    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
                    Command:="DecrementLeading", _
                    KeyCode:=Application.BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyShift, wdKeyDown)
    successCount = successCount + 1

    ' === Macro 9 ===
    ' KeyBindings.Add ...

    ' === Macro 10 ===
    ' KeyBindings.Add ...

    ' === Macro 11 ===
    ' KeyBindings.Add ...

    ' === Macro 12 ===
    ' KeyBindings.Add ...

    ' === Macro 13 ===
    ' KeyBindings.Add ...

    ' === Macro 14 ===
    ' KeyBindings.Add ...

    ' === Macro 15 ===
    ' KeyBindings.Add ...

    MsgBox "" & successCount & " shortcuts assigned successfully!", vbInformation
    Exit Sub

Failed:
    MsgBox "Failed to assign shortcut: " & Err.Description, vbCritical
End Sub

Sub MergeLinesFromBottomToTop()
    Dim sel As Range
    Set sel = Selection.Range

    If sel.Paragraphs.count < 2 Then
        MsgBox "Please select at least two lines.", vbExclamation
        Exit Sub
    End If

    Dim i As Long
    For i = sel.Paragraphs.count To 2 Step -1
        Dim lastPara As Range, prevPara As Range
        Set lastPara = sel.Paragraphs(i).Range
        Set prevPara = sel.Paragraphs(i - 1).Range

        ' Trim paragraph marks
        lastPara.End = lastPara.End - 1
        prevPara.End = prevPara.End - 1

        ' Insert last line content + space at start of previous line
        prevPara.InsertBefore lastPara.Text & ChrW(32)

        ' Delete the original last line
        lastPara.Delete
    Next i
End Sub

Sub CheckSpaceFontMismatch()
    Dim doc As Document: Set doc = ActiveDocument
    Dim scopeRng As Range, para As Paragraph, rngPara As Range
    Dim i As Long, fontBefore As String, fontAfter As String, fontSpace As String
    Dim mismatchedCount As Long, totalChecked As Long, paraChecked As Long
    Dim startTime As Single, endTime As Single, duration As Single
    Dim ch As Range

    ' Ask user scope
    Dim userChoice As VbMsgBoxResult
    userChoice = MsgBox("Check space font mismatch for:" & vbCrLf & vbCrLf & _
                        "Yes = Selected Paragraph(s)" & vbCrLf & "No = Whole Document", _
                        vbYesNoCancel + vbQuestion, "Check Scope")
    If userChoice = vbCancel Then Exit Sub
    If userChoice = vbYes Then
        Set scopeRng = Selection.Range
    Else
        Set scopeRng = doc.StoryRanges(wdMainTextStory)
    End If

    startTime = Timer
    Application.ScreenUpdating = False
    Application.StatusBar = "Checking paragraph by paragraph..."

    For Each para In scopeRng.Paragraphs
        Set rngPara = para.Range
        rngPara.End = rngPara.End - 1 ' exclude paragraph mark

        If rngPara.Characters.count < 3 Then GoTo SkipPara

        ' Directly loop over space characters (no skipping for font checks)
        For i = 2 To rngPara.Characters.count - 1
            If rngPara.Characters(i).Text = " " Then
                Set ch = rngPara.Characters(i)

                On Error Resume Next
                fontBefore = Trim(rngPara.Characters(i - 1).Font.Name)
                fontAfter = Trim(rngPara.Characters(i + 1).Font.Name)
                fontSpace = Trim(ch.Font.Name)
                On Error GoTo 0

                If fontSpace <> fontBefore Or fontSpace <> fontAfter Then
                    ch.HighlightColorIndex = wdYellow
                    mismatchedCount = mismatchedCount + 1
                End If
                totalChecked = totalChecked + 1
            End If
        Next i

SkipPara:
        paraChecked = paraChecked + 1
        If paraChecked Mod 10 = 0 Then
            Application.StatusBar = "Checked " & paraChecked & " paragraphs..."
            DoEvents
        End If
    Next para

    endTime = Timer
    duration = Round(endTime - startTime, 2)
    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "Check complete." & vbCrLf & _
           "Paragraphs checked: " & paraChecked & vbCrLf & _
           "Spaces checked: " & totalChecked & vbCrLf & _
           "Mismatched spaces highlighted: " & mismatchedCount & vbCrLf & _
           "Time taken: " & duration & " seconds", vbInformation
End Sub


Sub FixAllBlackFontToAutomatic_WithThemeSupport()
    Dim rng As Range
    Dim ch As Range
    Dim highlightChoice As VbMsgBoxResult
    Dim changedCount As Long
    Dim startTime As Single, endTime As Single
    Dim rColor As Long
    Dim rr As Long, rG As Long, rB As Long
    Dim i As Long

    highlightChoice = MsgBox("Highlight characters changed from Black to Automatic?", _
                             vbYesNoCancel + vbQuestion, "Highlight?")
    If highlightChoice = vbCancel Then Exit Sub

    If Selection.Type <> wdNoSelection Then
        Set rng = Selection.Range
    Else
        Set rng = ActiveDocument.Content
    End If

    startTime = Timer
    Application.ScreenUpdating = False

    For i = rng.Characters.count To 1 Step -1
        Set ch = rng.Characters(i)
        With ch.Font

            rColor = .TextColor.RGB
            rr = rColor Mod 256
            rG = (rColor \ 256) Mod 256
            rB = (rColor \ 65536) Mod 256

            If _
              .Color = wdColorBlack Or _
              rColor = RGB(0, 0, 0) Or _
              (.TextColor.Type = wdColorTypeRGB And rr <= 10 And rG <= 10 And rB <= 10) Or _
              (.Color = wdColorAutomatic And .TextColor.Type = wdColorTypeRGB And rColor = RGB(0, 0, 0)) _
            Then
                .Color = wdColorAutomatic
                If highlightChoice = vbYes Then ch.HighlightColorIndex = wdYellow
                changedCount = changedCount + 1
            End If
        End With
    Next i

    endTime = Timer
    Application.ScreenUpdating = True

    MsgBox "Theme-safe font cleanup complete." & vbCrLf & _
           "Characters changed: " & changedCount & vbCrLf & _
           "? Time taken: " & Round(endTime - startTime, 2) & " seconds", vbInformation
End Sub


Sub HidePlaceholderTags_AllSafe()
    Dim tagList As Variant, tag As Variant
    Dim replaceText As String
    Dim doc As Document: Set doc = ActiveDocument
    Dim rng As Range
    Dim storyRange As Range

    ' Placeholder tags
    tagList = Array("SIGNATURE", "LOGO", "EMBLEM", _
                    "DIGITAL SIGNATURE", "REDACTION", "STAMP")

    ' Replacement input
    replaceText = InputBox("Enter replacement text (leave blank to keep original):", "Replace Text")
    If StrPtr(replaceText) = 0 Then Exit Sub

    ' Loop through all available story ranges
    For Each storyRange In doc.StoryRanges
        Do While Not storyRange Is Nothing
            HideTagsInRange storyRange, tagList, replaceText
            Set storyRange = storyRange.NextStoryRange
        Loop
    Next storyRange

    MsgBox "All placeholder tags hidden (main body, headers, footers, text boxes, etc.).", vbInformation
End Sub

Private Sub HideTagsInRange(ByVal rng As Range, ByVal tags As Variant, ByVal replaceText As String)
    Dim tag As Variant
    For Each tag In tags
        With rng.Find
            .ClearFormatting
            .replacement.ClearFormatting
            .Text = tag
            .replacement.Text = IIf(replaceText = "", tag, replaceText)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .replacement.Font.Hidden = True
            .Execute Replace:=wdReplaceAll
        End With
    Next tag
End Sub

Sub ApplyTemplateLayoutAndHeaderFooter()
    Dim dlgOpen As FileDialog
    Dim templateDoc As Document
    Dim currentDoc As Document
    Dim sec As Section, tplSec As Section
    Dim hdr As HeaderFooter, ftr As HeaderFooter
    Dim applyPageSize As VbMsgBoxResult

    Set currentDoc = ActiveDocument

    ' Prompt for template document
    Set dlgOpen = Application.FileDialog(msoFileDialogFilePicker)
    With dlgOpen
        .Title = "Select the template Word document"
        .Filters.Clear
        .Filters.Add "Word Documents", "*.docx;*.doc"
        If .Show <> -1 Then Exit Sub
    End With

    ' Open template document (hidden)
    Set templateDoc = Documents.Open(FileName:=dlgOpen.SelectedItems(1), Visible:=False)

    ' Ask whether to apply page size from template
    applyPageSize = MsgBox("Do you want to apply the template's page size (width, height, orientation)?", _
                           vbYesNoCancel + vbQuestion, "Apply Page Size?")
    If applyPageSize = vbCancel Then
        templateDoc.Close SaveChanges:=False
        Exit Sub
    End If

    ' Apply page setup (margins, spacing, optional page size)
    With currentDoc.PageSetup
        .TopMargin = templateDoc.PageSetup.TopMargin
        .BottomMargin = templateDoc.PageSetup.BottomMargin
        .LeftMargin = templateDoc.PageSetup.LeftMargin
        .RightMargin = templateDoc.PageSetup.RightMargin
        .HeaderDistance = templateDoc.PageSetup.HeaderDistance
        .FooterDistance = templateDoc.PageSetup.FooterDistance

        If applyPageSize = vbYes Then
            .PageWidth = templateDoc.PageSetup.PageWidth
            .PageHeight = templateDoc.PageSetup.PageHeight
            .Orientation = templateDoc.PageSetup.Orientation
        End If
    End With

    ' Copy styles to preserve header/footer formatting
    templateDoc.CopyStylesFromTemplate templateDoc.FullName

    ' Apply header/footer settings and content with formatting
    For Each sec In currentDoc.Sections
        Set tplSec = templateDoc.Sections(1)

        ' Apply header/footer layout settings
        sec.PageSetup.DifferentFirstPageHeaderFooter = tplSec.PageSetup.DifferentFirstPageHeaderFooter
        sec.PageSetup.OddAndEvenPagesHeaderFooter = tplSec.PageSetup.OddAndEvenPagesHeaderFooter

        ' Copy headers
        For Each hdr In tplSec.Headers
            If hdr.Exists Then
                With sec.Headers(hdr.Index).Range
                    .FormattedText = hdr.Range.FormattedText
                    If Right(.Text, 1) = vbCr Then .Characters.Last.Delete

                    ' Fix font and paragraph spacing
                    With .Font
                        .Name = hdr.Range.Font.Name
                        .Size = hdr.Range.Font.Size
                    End With
                    With .ParagraphFormat
                        .SpaceBefore = hdr.Range.ParagraphFormat.SpaceBefore
                        .SpaceAfter = hdr.Range.ParagraphFormat.SpaceAfter
                        .LineSpacingRule = hdr.Range.ParagraphFormat.LineSpacingRule
                        .LineSpacing = hdr.Range.ParagraphFormat.LineSpacing
                    End With
                End With
            End If
        Next hdr

        ' Copy footers
        For Each ftr In tplSec.Footers
            If ftr.Exists Then
                With sec.Footers(ftr.Index).Range
                    .FormattedText = ftr.Range.FormattedText
                    If Right(.Text, 1) = vbCr Then .Characters.Last.Delete

                    ' Fix font and paragraph spacing
                    With .Font
                        .Name = ftr.Range.Font.Name
                        .Size = ftr.Range.Font.Size
                    End With
                    With .ParagraphFormat
                        .SpaceBefore = ftr.Range.ParagraphFormat.SpaceBefore
                        .SpaceAfter = ftr.Range.ParagraphFormat.SpaceAfter
                        .LineSpacingRule = ftr.Range.ParagraphFormat.LineSpacingRule
                        .LineSpacing = ftr.Range.ParagraphFormat.LineSpacing
                    End With
                End With
            End If
        Next ftr
    Next sec

    templateDoc.Close SaveChanges:=False

    MsgBox "Page layout, header/footer, and formatting copied successfully.", vbInformation
End Sub



Sub InsertBosnianCharacter()
    Dim choice As String
    Dim charMap As Object
    Set charMap = CreateObject("Scripting.Dictionary")

    ' Unicode-safe Bosnian characters using ChrW
    charMap.Add "1", ChrW(&H10D)   ' c
    charMap.Add "2", ChrW(&H107)   ' c
    charMap.Add "3", ChrW(&H111)   ' d
    charMap.Add "4", ChrW(&H161)   ' š
    charMap.Add "5", ChrW(&H17E)   ' ž
    charMap.Add "6", ChrW(&H10C)   ' C
    charMap.Add "7", ChrW(&H106)   ' C
    charMap.Add "8", ChrW(&H110)   ' Ð
    charMap.Add "9", ChrW(&H160)   ' Š
    charMap.Add "10", ChrW(&H17D)  ' Ž

    ' Build the prompt with both character and its name
    Dim promptText As String
    promptText = "Insert Bosnian character (1–10):" & vbCrLf & _
        "1  = c (c caron)" & vbCrLf & _
        "2  = c (c acute)" & vbCrLf & _
        "3  = d (d stroke)" & vbCrLf & _
        "4  = š (s caron)" & vbCrLf & _
        "5  = ž (z caron)" & vbCrLf & _
        "6  = C (C caron)" & vbCrLf & _
        "7  = C (C acute)" & vbCrLf & _
        "8  = Ð (D stroke)" & vbCrLf & _
        "9  = Š (S caron)" & vbCrLf & _
        "10 = Ž (Z caron)"

    ' Prompt the user
    choice = InputBox(promptText, "Insert Bosnian Character")

    If charMap.Exists(Trim(choice)) Then
        Selection.TypeText Text:=charMap(Trim(choice))
    ElseIf choice <> "" Then
        MsgBox "Invalid selection. Choose 1–10.", vbExclamation
    End If
End Sub



Sub ExportAllKeyboardShortcuts()
    Dim kb As KeyBinding
    Dim exportPath As String
    Dim f As Integer
    Dim contextName As String

    exportPath = Environ$("USERPROFILE") & "\Desktop\Word_Shortcuts_List.txt"
    f = FreeFile

    Open exportPath For Output As #f
    Print #f, "Shortcut Keys Assigned in Word:" & vbCrLf
    Print #f, "-----------------------------------"

    For Each kb In KeyBindings
        On Error Resume Next
        contextName = kb.context.NameLocal
        If Err.Number <> 0 Then contextName = "(Unavailable)"
        On Error GoTo 0

        Print #f, "Command: " & kb.Command
        Print #f, "Key: " & kb.keyString
        Print #f, "Context: " & contextName
        Print #f, "-----------------------------------"
    Next kb

    Close #f

    MsgBox "Shortcut keys exported to:" & vbCrLf & exportPath, vbInformation, "Export Complete"
End Sub

Sub ProcessTags_HideUnhide_Highlight()
    Dim frm As New frmHideTags
    Dim selectedTags As Collection
    Dim storyRange As Range
    Dim tag As Variant
    Dim customText As String

    Dim countHide As Long, countUnhide As Long
    Dim countHighlight As Long, countUnhighlight As Long

    frm.Show
    If frm.tag <> "OK" Then Exit Sub

    Set selectedTags = New Collection
    If frm.chkDIGITALSIGNATURE.Value Then
    selectedTags.Add "[DIGITAL SIGNATURE]"
    selectedTags.Add "DIGITAL SIGNATURE"
    End If
    If frm.chkSIGNATURE.Value Then
    selectedTags.Add "[SIGNATURE]"
    selectedTags.Add "SIGNATURE"
    End If
    If frm.chkLOGO.Value Then
    selectedTags.Add "[LOGO]"
    selectedTags.Add "LOGO"
    End If
    If frm.chkEMBLEM.Value Then
    selectedTags.Add "[EMBLEM]"
    selectedTags.Add "EMBLEM"
    End If
    If frm.chkREDACTION.Value Then
    selectedTags.Add "[REDACTION]"
    selectedTags.Add "REDACTION"
    End If
    If frm.chkSTAMP.Value Then
    selectedTags.Add "[STAMP]"
    selectedTags.Add "STAMP"
    End If

    customText = Trim(frm.txtCustomText.Text)
    If customText <> "" Then
    selectedTags.Add "[" & customText & "]"
    selectedTags.Add customText
    End If

    If selectedTags.count = 0 Then
        MsgBox "No tags or custom text selected.", vbExclamation
        Exit Sub
    End If

    Dim tagIndex As Long, totalSteps As Long
    totalSteps = selectedTags.count * CountStoryRanges(ActiveDocument)

    tagIndex = 0
    For Each storyRange In ActiveDocument.StoryRanges
        Do While Not storyRange Is Nothing
            For Each tag In selectedTags
                tagIndex = tagIndex + 1
                Application.StatusBar = "Processing: " & tag & " (" & tagIndex & " of " & totalSteps & ")"
                DoEvents
                ProcessSingleTag storyRange, tag, _
                    frm.optHide.Value, frm.optUnhide.Value, _
                    frm.optHighlight.Value, frm.optRemoveHighlight.Value, _
                    countHide, countUnhide, countHighlight, countUnhighlight
            Next tag
            Set storyRange = storyRange.NextStoryRange
        Loop
    Next storyRange

    Application.StatusBar = False ' Reset

    ' Build result message
    Dim msg As String: msg = "Processing complete." & vbCrLf & vbCrLf
    If countHide > 0 Then msg = msg & "Hidden: " & countHide & vbCrLf
    If countUnhide > 0 Then msg = msg & "Unhidden: " & countUnhide & vbCrLf
    If countHighlight > 0 Then msg = msg & "Highlighted: " & countHighlight & vbCrLf
    If countUnhighlight > 0 Then msg = msg & "Highlight Removed: " & countUnhighlight & vbCrLf
    If msg = "Processing complete." & vbCrLf & vbCrLf Then msg = msg & "No matches found."

    MsgBox msg, vbInformation, "Summary"
End Sub


Private Sub ProcessSingleTag(ByVal rng As Range, ByVal tag As String, _
                             ByVal doHide As Boolean, ByVal doUnhide As Boolean, _
                             ByVal doHighlight As Boolean, ByVal doRemoveHighlight As Boolean, _
                             ByRef countHide As Long, ByRef countUnhide As Long, _
                             ByRef countHighlight As Long, ByRef countUnhighlight As Long)

    Dim foundRange As Range
    Dim searchRange As Range

    Set searchRange = rng.Duplicate

    With searchRange.Find
        .ClearFormatting
        .Text = tag
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = True
        .MatchWholeWord = True
        .MatchWildcards = False

        ' Restrict search to formatted text if needed
        If doUnhide Then .Font.Hidden = True
        If doRemoveHighlight Then .Highlight = True
    End With

    Do While searchRange.Find.Execute
        Set foundRange = searchRange.Duplicate
        foundRange.End = foundRange.Start + Len(tag)

        ' Only apply actions to exact match
        If doHide Then
    If Not foundRange.Font.Hidden Then
        foundRange.Font.Hidden = True
        countHide = countHide + 1
    End If
End If

If doUnhide Then
    If foundRange.Font.Hidden Then
        foundRange.Font.Hidden = False
        countUnhide = countUnhide + 1
    End If
End If

If doHighlight Then
    If foundRange.HighlightColorIndex <> wdYellow Then
        foundRange.HighlightColorIndex = wdYellow
        countHighlight = countHighlight + 1
    End If
End If

If doRemoveHighlight Then
    If foundRange.HighlightColorIndex <> wdNoHighlight Then
        foundRange.HighlightColorIndex = wdNoHighlight
        countUnhighlight = countUnhighlight + 1
    End If
End If


        ' Move searchRange to after current match
        searchRange.Start = foundRange.End
        searchRange.End = rng.End
        Set searchRange = searchRange.Duplicate
    Loop
End Sub




Private Function CountStoryRanges(doc As Document) As Long
    Dim sr As Range, count As Long
    count = 0
    For Each sr In doc.StoryRanges
        Do While Not sr Is Nothing
            count = count + 1
            Set sr = sr.NextStoryRange
        Loop
    Next sr
    CountStoryRanges = count
End Function


Sub LowerTextBy1pt()
    Dim rng As Range
    Set rng = Selection.Range

    With rng.Font
        ' If no previous position set, assume 0
        If .Position = wdUndefined Then
            .Position = -1
        Else
            .Position = .Position - 1
        End If
    End With
End Sub

Sub RaiseTextBy1pt()
    Dim rng As Range
    Set rng = Selection.Range

    With rng.Font
        ' If no previous position set, assume 0
        If .Position = wdUndefined Then
            .Position = 1
        Else
            .Position = .Position + 1
        End If
    End With
End Sub

Option Explicit


Sub HideTextUntilSlashInTable()
    Dim cel As cell
    Dim rng As Range
    Dim slashPos As Long
    
    ' Ensure selection is inside a table
    If Not Selection.Information(wdWithInTable) Then
        MsgBox "Please select a table or table cells first.", vbExclamation
        Exit Sub
    End If
    
    ' Work only on the selected cells
    For Each cel In Selection.Cells
        Set rng = cel.Range
        rng.End = rng.End - 1 ' Exclude end-of-cell marker
        
        slashPos = InStr(rng.Text, "/")
        
        If slashPos > 0 Then
            Dim hideRange As Range
            Set hideRange = rng.Duplicate
            ' Include the slash in the hidden part
            hideRange.End = hideRange.Start + slashPos
            hideRange.Font.Hidden = True
        End If
    Next cel
End Sub

Sub FormatSelectedText()
    Dim a As String
    
    ' Check if the user has selected some text
    If Selection.Type = wdSelectionNormal Then
    
        ' Store the selected text in variable 'a'
        a = Trim(Selection.Text)
        
        ' Make sure it's not empty
        If Len(a) > 0 Then
            ' Replace with "a / a"
            Selection.Text = a & " / " & a
        End If
        
    Else
        MsgBox "Please select some text first.", vbInformation, "No Selection"
    End If
End Sub

Sub Table_0_04()
Attribute Table_0_04.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Table_0_04"
'
' Table_0_04 Macro
'
'
    With Selection.Tables(1)
        .TopPadding = InchesToPoints(0)
        .BottomPadding = InchesToPoints(0)
        .LeftPadding = InchesToPoints(0.04)
        .RightPadding = InchesToPoints(0.04)
        .Spacing = 0
        .AllowPageBreaks = True
        .AllowAutoFit = False
    End With
    Selection.Tables(1).Rows.LeftIndent = InchesToPoints(0.18)
    With Selection.ParagraphFormat
        .LeftIndent = InchesToPoints(0)
        .RightIndent = InchesToPoints(0)
        .SpaceBefore = 2
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(1.05)
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = InchesToPoints(0)
        .outlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
        .ReadingOrder = wdReadingOrderLtr
    End With
    With Selection.ParagraphFormat
        .LeftIndent = InchesToPoints(0)
        .RightIndent = InchesToPoints(0)
        .SpaceBefore = 2
        .SpaceBeforeAuto = False
        .SpaceAfter = 2
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(1.05)
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = InchesToPoints(0)
        .outlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
        .ReadingOrder = wdReadingOrderLtr
    End With
    Selection.Tables(1).Select
    With Selection.ParagraphFormat
        .SpaceBefore = 2
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .LineUnitBefore = 0
    End With
    With Selection.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfter = 2
        .SpaceAfterAuto = False
        .LineUnitAfter = 0
    End With
    Selection.Rows.HeightRule = wdRowHeightAtLeast
    Selection.Rows.Height = InchesToPoints(0.04)
End Sub
