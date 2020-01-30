Function FileOrDirExists(PathName As String) As Boolean
     'Macro Purpose: Function returns TRUE if the specified file
     '               or folder exists, false if not.
     'PathName     : Supports Windows mapped drives or UNC
     '             : Supports Macintosh paths
     'File usage   : Provide full file path and extension
     'Folder usage : Provide full folder path
     '               Accepts with/without trailing "\" (Windows)
     '               Accepts with/without trailing ":" (Macintosh)
     
    Dim iTemp As Integer
     
     'if error then file does not exist - go to handler
    On Error GoTo handler:
    iTemp = GetAttr(PathName)
    
    'If the value was assigned with no errors then the file exists.
    FileOrDirExists = True
    
    Exit Function
    
handler:
    FileOrDirExists = False
    
End Function