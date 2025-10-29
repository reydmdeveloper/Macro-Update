Attribute VB_Name = "Module3"
Sub ShowCleanupTool()
    CleanupToolForm.Show
End Sub

Sub Remove_Section_Breaks_And_PageBreaks()
    Dim i As Long
    Dim doc As Document
    Dim secRange As Range
    Set doc = ActiveDocument

    ' === Safely remove all section breaks ===
    ' Loop from second-to-last section to the first
    For i = doc.Sections.count - 1 To 1 Step -1
        Set secRange = doc.Sections(i).Range
        ' Move range to end of section to target the section break only
        secRange.Collapse Direction:=wdCollapseEnd
        If secRange.Characters.Last.Previous = Chr(12) Then ' Section break char
            secRange.MoveStart wdCharacter, -1
            secRange.Delete
        End If
    Next i

    ' === Replace manual page breaks (^m) ===
    With Selection.Find
        .ClearFormatting
        .replacement.ClearFormatting
        .Text = "^m"
        .replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With

    ' === Replace column breaks (^n) ===
    With Selection.Find
        .Text = "^n"
        .replacement.Text = "^p"
        .Execute Replace:=wdReplaceAll
    End With
End Sub

Sub Char_Clean()
    Selection.WholeStory
    With Selection.Font
        .NameFarEast = ""
        .NameAscii = ""
        .NameOther = ""
        .Name = ""
        .Spacing = 0
        .Scaling = 100
        .Position = 0
        .NameBi = ""
    End With
End Sub

Sub Para_Clean()
    Selection.WholeStory
    With Selection.ParagraphFormat
        .LeftIndent = InchesToPoints(0)
        .RightIndent = InchesToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .FirstLineIndent = InchesToPoints(0)
    End With
End Sub

Sub NumeringToText()
    ActiveDocument.ConvertNumbersToText
End Sub

Sub Remove_Extra_Paragraphs()
    Dim rng As Range
    Dim i As Integer

    ' Run multiple times to collapse large gaps
    For i = 1 To 15
        Set rng = ActiveDocument.Content
        With rng.Find
            .ClearFormatting
            .replacement.ClearFormatting
            .Text = "^p^p"
            .replacement.Text = "^p"
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchWildcards = False
        End With
        rng.Find.Execute Replace:=wdReplaceAll
    Next i
End Sub

Sub InsertAndFormatPicture()
    Dim dlgOpen As FileDialog
    Dim selectedImage As String
    Dim shp As Shape
    Dim rng As Range

    ' Open File Dialog to select the image
    Set dlgOpen = Application.FileDialog(msoFileDialogFilePicker)
    dlgOpen.Title = "Select an Image"
    dlgOpen.Filters.Clear
    dlgOpen.Filters.Add "Image Files", "*.jpg; *.jpeg; *.png; *.bmp; *.gif"
    
    If dlgOpen.Show <> -1 Then Exit Sub ' User cancelled

    selectedImage = dlgOpen.SelectedItems(1)

    ' Get range at cursor (anchor point)
    Set rng = Selection.Range

    ' Insert picture as shape anchored at current cursor location
    Set shp = ActiveDocument.Shapes.AddPicture( _
        FileName:=selectedImage, _
        LinkToFile:=False, _
        SaveWithDocument:=True, _
        Anchor:=rng)

    ' Set layout to Behind Text
    shp.WrapFormat.Type = wdWrapBehind

    ' Position settings
    With shp
        .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
        .Left = wdShapeCenter

        .RelativeVerticalPosition = wdRelativeVerticalPositionPage
        .Top = 0

        .LockAnchor = True
    End With

    ' Reset size
    shp.LockAspectRatio = msoFalse
    shp.ScaleHeight 1, msoTrue
    shp.ScaleWidth 1, msoTrue
End Sub


Sub IncrementLeading()
    Dim para As Paragraph
    Dim currentSpacing As Single

    For Each para In Selection.Paragraphs
        With para.Format
            ' Ensure spacing rule is set to Multiple
            .LineSpacingRule = wdLineSpaceMultiple
            
            ' If spacing is already set to multiple, use that value
            currentSpacing = .LineSpacing

            ' If it’s unusually low (or defaulted), assume it's 1.0
            If currentSpacing < 1 Then currentSpacing = 1

            ' Increase spacing by 0.1
            .LineSpacing = currentSpacing + 0.1
        End With
    Next para
End Sub

Sub DecrementLeading()
    Dim para As Paragraph
    Dim currentSpacing As Single

    For Each para In Selection.Paragraphs
        With para.Format
            ' Ensure spacing rule is set to Multiple
            .LineSpacingRule = wdLineSpaceMultiple
            
            ' If spacing is already set to multiple, use that value
            currentSpacing = .LineSpacing

            ' If it’s unusually low (or defaulted), assume it's 1.0
            If currentSpacing < 1 Then currentSpacing = 1

            ' Decrease spacing by 0.1, but not below 1.0
            If currentSpacing > 1 Then
                .LineSpacing = currentSpacing - 0.1
            Else
                .LineSpacing = 1
            End If
        End With
    Next para
End Sub
