

Function AssignFilePath(toPath As String)
    
    If FileOrDirExists(toPath) Then
        Dim originalDotIndex As Integer
        'Use InsStrRev to check the string from the end in case there are "."s in folder names
        originalDotIndex = InStrRev(toPath, ".")
        Dim copies_ As Integer
        copies_ = 0
        
        Dim extLen As Integer
        ' The length of the string minus one to get the last index and then subtract that by the index of the dot before the extension to get the length of the extension
        extLen = Len(toPath) - originalDotIndex
        
        Do While FileOrDirExists(toPath)
            copies_ = copies_ + 1
            ' This should add a number in parenthesis equal to the number of copies of the file
            toPath = Left(toPath, originalDotIndex - 1) & "(" & copies_ & ")." & Right(toPath, extLen)
        Loop
    End If
    
    AssignFilePath = toPath

End Function