
'Intended to be used to search a column range
'Iterates over the cells in the range and returns the value of the first cell with bold text
Public Function ReturnBoldValue(BoldRange As Range)
    Dim cell As Range
    For Each cell In BoldRange
        If cell.Font.Bold Then
            ReturnBoldValue = cell.Value
            Exit Function
        End If
    Next cell
End Function