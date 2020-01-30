

Public Sub CopyRowAndColumnSizes(TargetRange As Range, SourceRange As Range)
    CopyRowHeights TargetRange, SourceRange
    CopyColumnWidths TargetRange, SourceRange
End Sub


Public Sub CopyRowHeights(TargetRange As Range, SourceRange As Range)
    Dim row_ As Long
    
    For row_ = 1 To SourceRange.Rows.Count
        TargetRange.Rows(row_).RowHeight = SourceRange.Rows(row_).RowHeight
    Next row_
End Sub


Public Sub CopyColumnWidths(TargetRange As Range, SourceRange As Range)
    Dim column_ As Long
    
    For column_ = 1 To SourceRange.Columns.Count
        TargetRange.Columns(column_).ColumnWidth = SourceRange.Columns(column_).ColumnWidth
    Next column_
End Sub