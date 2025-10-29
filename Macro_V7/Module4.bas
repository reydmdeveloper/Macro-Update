Attribute VB_Name = "Module4"
Sub FormatWholeDocument()

    Dim para As Paragraph
    Dim tbl As Table
    Dim cell As cell

    ' 1. Format all normal paragraphs (outside tables)
    For Each para In ActiveDocument.Paragraphs
        With para.Format
            .Alignment = wdAlignParagraphJustify
            .SpaceBefore = 6
            .SpaceAfter = 6
            .LeftIndent = InchesToPoints(0.2)
            .RightIndent = InchesToPoints(0)
            .FirstLineIndent = InchesToPoints(0.2)
        End With
    Next para

    ' 2. Format all paragraphs inside all tables
    For Each tbl In ActiveDocument.Tables
        For Each cell In tbl.Range.Cells
            For Each para In cell.Range.Paragraphs
                With para.Format
                    .Alignment = wdAlignParagraphJustify
                    .SpaceBefore = 6
                    .SpaceAfter = 6
                    .LeftIndent = InchesToPoints(0.2)
                    .RightIndent = InchesToPoints(0)
                    .FirstLineIndent = InchesToPoints(0.2)
                End With
            Next para
        Next cell
    Next tbl

    MsgBox "Formatting completed successfully.", vbInformation

End Sub


