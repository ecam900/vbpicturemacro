Option Explicit

'********************************************************************************
'Picture
'
' Purpose:  Looks for Image names posted in column B in the file folder and
'           then resizes the images and pastes them in Column A
'
'
'********************************************************************************

Sub Picture()

    Const EXIT_TEXT         As String = "Please Check Data Sheet"
    Const NO_PICTURE_FOUND  As String = "No picture found"

    Dim picName             As String
    Dim picFullName         As String
    Dim rowIndex            As Long
    Dim lastRow             As Long
    Dim selectedFolder      As String
    Dim data()              As Variant
    Dim wks                 As Worksheet
    Dim cell                As Range
    Dim pic                 As Picture

    On Error GoTo ErrorHandler

    selectedFolder = GetFolder
    If Len(selectedFolder) = 0 Then GoTo ExitRoutine

    Application.ScreenUpdating = False

    Set wks = ActiveSheet
    ' this is not bulletproof but for now should work fine
    lastRow = wks.Cells(14, "B").End(xlDown).Row
    data = wks.Range(wks.Cells(1, "B"), wks.Cells(lastRow, "B")).Value2
    
    ' This is the entire one-dimensional array of names
    
    For rowIndex = 14 To UBound(data, 1)
        If StrComp(data(rowIndex, 1), EXIT_TEXT, vbTextCompare) = 0 Then GoTo ExitRoutine

        picName = data(rowIndex, 1)
        picFullName = selectedFolder & picName & ".jpg"

        If Len(Dir(picFullName)) > 0 Then
            Set cell = wks.Cells(rowIndex, "A")
'            Set pic = wks.Pictures.Insert(picFullName)
            wks.Shapes.AddPicture Filename:=(picFullName), _
            linktofile:=msoFalse, _
            savewithdocument:=msoCTrue, _
            Left:=cell.Left, Top:=cell.Top + 2, Width:=cell.Width, Height:=cell.Height - 2
            
'   -----------------------------------------------------------
'        PREVIOUS WAY OF DOING IT:

'            With pic
'                .ShapeRange.LockAspectRatio = msoFalse
'                .Height = cell.Height
'                .Width = cell.Width
'                .Top = cell.Top
'                .Left = cell.Left
'                .Placement = xlMoveAndSize
'            End With
'   -----------------------------------------------------------

        Else
            wks.Cells(rowIndex, "A").Value = NO_PICTURE_FOUND
        End If

    Next rowIndex

    Range("A10").Select

ExitRoutine:
    Set wks = Nothing
    Set pic = Nothing
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Range("B20").Select
    MsgBox Prompt:="Unable to find photo", _
           Title:="An error occured", _
           Buttons:=vbExclamation
    Resume ExitRoutine

End Sub

Private Function GetFolder() As String

    Dim selectedFolder  As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = Application.DefaultFilePath & "\"
        .Title = "Select the folder containing the Image/PDF files."
        .Show

        If .SelectedItems.Count > 0 Then
            selectedFolder = .SelectedItems(1)
            If Right$(selectedFolder, 1) <> Application.PathSeparator Then _
                selectedFolder = selectedFolder & Application.PathSeparator
        End If

    End With
    GetFolder = selectedFolder

End Function


