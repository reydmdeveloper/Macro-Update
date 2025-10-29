Attribute VB_Name = "Module6"
Sub SetCustomStyleOutlineLevels()
    Dim styleName As String
    Dim outlineLevel As String
    Dim s As Style
    Dim msg As String
    Dim iLevel As Integer

    MsgBox "This macro lets you assign Outline Levels (1–9) to your custom styles." & vbCrLf & _
           "Those styles will then appear in the Navigation Pane.", vbInformation, "Link Styles to Navigation Pane"

    Do
        ' Ask for style name
        styleName = InputBox("Enter the style name to set (e.g. H1, MyHeading, etc.)" & vbCrLf & _
                             "Leave blank and click OK to finish.", "Style to Configure")

        If styleName = "" Then Exit Do ' User pressed OK with empty input

        On Error Resume Next
        Set s = ActiveDocument.Styles(styleName)
        On Error GoTo 0

        If s Is Nothing Then
            MsgBox "Style '" & styleName & "' was not found in this document.", vbExclamation
        Else
            ' Ask for outline level
            outlineLevel = InputBox("Enter Outline Level for '" & styleName & "' (1–9)" & vbCrLf & _
                                    "Level 1 = Main Heading, Level 2 = Subheading, etc.", _
                                    "Outline Level", "1")
            If outlineLevel = "" Then Exit Do
            
            iLevel = Val(outlineLevel)
            If iLevel < 1 Or iLevel > 9 Then
                MsgBox "Invalid level. Please enter a number between 1 and 9.", vbCritical
            Else
                s.ParagraphFormat.outlineLevel = iLevel
                MsgBox "Style '" & styleName & "' linked to Outline Level " & iLevel & ".", vbInformation
            End If
        End If
    Loop

    MsgBox "? All selected styles have been updated. Apply them to headings to show in Navigation Pane.", vbInformation
End Sub

Sub HideStyleFromNavigationPane()
    Dim styleName As String
    Dim s As Style
    Dim para As Paragraph
    Dim countChanged As Long

    ' Ask for style name
    styleName = InputBox("Enter the Style Name you want to hide from the Navigation Pane:", _
                         "Hide Style from Navigation")

    If Trim(styleName) = "" Then
        MsgBox "No style entered. Operation cancelled.", vbInformation
        Exit Sub
    End If

    On Error Resume Next
    Set s = ActiveDocument.Styles(styleName)
    On Error GoTo 0

    If s Is Nothing Then
        MsgBox "Style '" & styleName & "' was not found in this document.", vbExclamation
        Exit Sub
    End If

    ' Set the style’s outline level to Body Text
    s.ParagraphFormat.outlineLevel = wdOutlineLevelBodyText

    ' Clear any direct outline formatting applied to paragraphs of this style
    countChanged = 0
    For Each para In ActiveDocument.Paragraphs
        If para.Style = s.NameLocal Then
            If para.outlineLevel <> wdOutlineLevelBodyText Then
                para.outlineLevel = wdOutlineLevelBodyText
                countChanged = countChanged + 1
            End If
        End If
    Next para

    MsgBox "Style '" & styleName & "' has been hidden from the Navigation Pane." & vbCrLf & _
           "Paragraphs updated: " & countChanged, vbInformation
End Sub

