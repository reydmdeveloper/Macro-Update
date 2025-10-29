Attribute VB_Name = "Module5"
Option Explicit

Sub ReplaceSelectedPicture_FromLastFolder_Single()
    On Error GoTo Fail

    Const VAR_NAME As String = "REY_LastImageFolder"
    Dim folderPath As String, pickedPath As String
    Dim dlg As FileDialog
    Dim sNum As String, n As Long
    Dim pattern As String, foundPath As String
    Dim p As Long
    Dim f As String

    ' 1) Get last-used folder (store in doc variable). If missing, ask once.
    On Error Resume Next
    folderPath = ActiveDocument.Variables(VAR_NAME).Value
    On Error GoTo 0

    If Len(folderPath) = 0 Then
        Set dlg = Application.FileDialog(msoFileDialogFilePicker)
        With dlg
            .Title = "Pick ANY image inside the folder you want to use"
            .AllowMultiSelect = False
            .Filters.Clear
            .Filters.Add "Image Files", "*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.tif;*.tiff"
            If .Show <> -1 Then Exit Sub
            pickedPath = .SelectedItems(1)
        End With

        p = InStrRev(pickedPath, Application.PathSeparator)
        If p = 0 Then
            MsgBox "Could not determine folder from selection.", vbExclamation
            Exit Sub
        End If

        folderPath = Left$(pickedPath, p)
        ActiveDocument.Variables(VAR_NAME).Value = folderPath
        ActiveDocument.Fields.Update
    Else
        ' ensure trailing separator
        If Right$(folderPath, 1) <> "\" And Right$(folderPath, 1) <> "/" Then
            folderPath = folderPath & Application.PathSeparator
        End If
    End If

    ' 2) Ask for image number
    sNum = InputBox("Enter image number (e.g., 1 ? matches *_Page_1.png):", "Replace Picture")
    If Len(Trim$(sNum)) = 0 Then Exit Sub
    If Not IsNumeric(sNum) Then
        MsgBox "Please enter a valid number.", vbExclamation: Exit Sub
    End If

    n = CLng(Val(sNum))
    If n <= 0 Then
        MsgBox "Please enter a number greater than zero.", vbExclamation: Exit Sub
    End If

    ' 3) Find file *_Page_<n>.png in that folder (strict .png per your spec)
    pattern = "*_Page_" & CStr(n) & ".png"
    f = Dir(folderPath & pattern, vbNormal)
    If Len(f) = 0 Then
        MsgBox "No file found in:" & vbCrLf & folderPath & vbCrLf & _
               "matching: " & pattern, vbExclamation, "Not found"
        Exit Sub
    End If
    foundPath = folderPath & f

    ' 4) Ensure a picture is selected
    If Selection.Type = wdSelectionInlineShape Then
        If Selection.InlineShapes.count = 0 Then
            MsgBox "Select the picture to replace and run again.", vbExclamation: Exit Sub
        End If
    ElseIf Selection.Type = wdSelectionShape Then
        If Selection.ShapeRange.count = 0 Then
            MsgBox "Select the picture to replace and run again.", vbExclamation: Exit Sub
        End If
    Else
        MsgBox "Select the picture to replace and run again.", vbExclamation: Exit Sub
    End If

    ' 5) Replace while preserving layout/size
    If Selection.Type = wdSelectionInlineShape And Selection.InlineShapes.count > 0 Then
        Dim ils As inlineShape, rng As Range
        Dim inlineW As Single, inlineH As Single

        Set ils = Selection.InlineShapes(1)
        Set rng = ils.Range.Duplicate
        inlineW = ils.Width: inlineH = ils.Height

        ils.Delete

        Dim ilsNew As inlineShape
        Set ilsNew = rng.InlineShapes.AddPicture(FileName:=foundPath, LinkToFile:=False, SaveWithDocument:=True)
        On Error Resume Next
        ilsNew.LockAspectRatio = msoFalse
        ilsNew.Width = inlineW: ilsNew.Height = inlineH
        On Error GoTo 0

    ElseIf Selection.Type = wdSelectionShape And Selection.ShapeRange.count > 0 Then
        Dim shpOld As Shape
        Dim anc As Range
        Dim relH As WdRelativeHorizontalPosition
        Dim relV As WdRelativeVerticalPosition
        Dim shpLeft As Single, shpTop As Single
        Dim shapeW As Single, shapeH As Single
        Dim wrapT As WdWrapType
        Dim lockAnc As Boolean, layoutInCell As Boolean
        Dim zPos As Long

        Set shpOld = Selection.ShapeRange(1)

        Set anc = shpOld.Anchor.Duplicate
        relH = shpOld.RelativeHorizontalPosition
        relV = shpOld.RelativeVerticalPosition
        shpLeft = shpOld.Left
        shpTop = shpOld.Top
        shapeW = shpOld.Width
        shapeH = shpOld.Height
        wrapT = shpOld.WrapFormat.Type
        lockAnc = shpOld.LockAnchor
        layoutInCell = shpOld.layoutInCell
        zPos = shpOld.ZOrderPosition

        shpOld.Delete

        Dim shpNew As Shape
        Set shpNew = ActiveDocument.Shapes.AddPicture( _
            FileName:=foundPath, LinkToFile:=False, SaveWithDocument:=True, Anchor:=anc)

        With shpNew
            .WrapFormat.Type = wrapT
            .RelativeHorizontalPosition = relH
            .RelativeVerticalPosition = relV
            .Left = shpLeft: .Top = shpTop
            .LockAnchor = lockAnc
            .layoutInCell = layoutInCell
            On Error Resume Next
            .LockAspectRatio = msoFalse
            .Width = shapeW: .Height = shapeH
            If zPos > 1 Then .ZOrder msoBringForward
            On Error GoTo 0
        End With
    End If

    Exit Sub
Fail:
    MsgBox "Replace failed: " & Err.Description, vbExclamation, "Replace Picture"
End Sub


