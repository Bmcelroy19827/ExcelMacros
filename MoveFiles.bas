
Private Sub MovePreviousFile(workbookSheet_ as string, fromRange As String, toRange As String, fileNameExpression_ As String)
    Dim oldFile As String
    Dim currentWorkbook As Workbook
    Dim pathForOldFile As String 'This is the folder the old csv file will be moved to when complete
    Dim pathForNewFile As String
    Dim fromPath As String 'This is the current directory of the file which is where new Csv Files are saved
    Dim toPath As String 'This is where the old csv will be moved.
    Dim FSO As Object
    
    Set currentWorkbook = ThisWorkbook
    With currentWorkbook.Worksheets(workbookSheet_)
        pathForOldFile = .Range(toRange)
        pathForNewFile = .Range(fromRange)
    End With
    
    If pathForOldFile = "" Then
        Exit Sub
    End If
    
    If pathForNewFile = "" Then
        pathForNewFile = currentWorkbook.Path
    End If
    oldFile = Dir(pathForNewFile & fileNameExpression_) 'The first one found
    'If a file is found matching the above expression then we will assign the path variables
    If oldFile <> "" Then
        Set FSO = CreateObject("scripting.filesystemobject")
        Do While oldFile <> ""
            fromPath = pathForNewFile & "\" & oldFile
            toPath = pathForOldFile & "\" & oldFile
            If FileOrDirExists(toPath) Then
                Dim originalDotIndex As Integer
                'Use InsStrRev to check the string from the end in case there are "."s in folder names
                originalDotIndex = InStrRev(toPath, ".")
                
                Dim copies_ As Integer
                copies_ = 0
                Do While FileOrDirExists(toPath)
                    Dim extLen As Integer
                    Dim newDotIndex As Integer
                    copies_ = copies_ + 1
                    newDotIndex = InStrRev(toPath, ".")
                    ' The length of the string minus one to get the last index and then subtract that by the index of the dot before the extension to get the length of the extension
                    extLen = Len(toPath) - newDotIndex
                    ' This should add a number in parenthesis equal to the number of copies of the file
                    toPath = Left(toPath, originalDotIndex - 1) & "(" & copies_ & ")." & Right(toPath, extLen)
                Loop
            End If
                              
                
            FSO.moveFile Source:=fromPath, Destination:=toPath
            oldFile = Dir(pathForNewFile & fileNameExpression_)
        Loop
    End If
End Sub