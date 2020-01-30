
Public Sub GetFolder(TargetCell As Range)
    Dim fldr As FileDialog
    Dim sItem As String
    Dim strPath As String
    
    If TargetCell.Value = "" Then
        strPath = Application.ActiveWorkbook.Path
    Else
        strPath = TargetCell.Value
    End If
    On Error GoTo handler:
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select Target Folder"
        .AllowMultiSelect = False
        .InitialFileName = strPath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    TargetCell.Value = sItem
    Set fldr = Nothing
    
    Exit Sub

handler:
    TargetCell.Value = strPath

End Sub