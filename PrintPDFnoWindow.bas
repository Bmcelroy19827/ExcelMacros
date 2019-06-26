Attribute VB_Name = "PrintPDFnoWindow"

Dim daName
Dim daPDFfolder
Dim daDateFormatted As String
Dim pdfPath As String
Dim characters As String
Dim charIndex As Integer
Dim baseName As String







Sub exportAPDFwindowless()
Attribute exportAPDFwindowless.VB_ProcData.VB_Invoke_Func = "d\n14"
'********************************************************
'
'@author Bryan McElroy 9/18
'
' This macro will create a pdf from selected cells
'
'********************************************************
       
    characters = "bcdefghijklmnopqrstuvwxyz"
    daName = Application.ActiveWorkbook.Name
    daDateFormatted = Format(Date, "mmddyyyy")
    charIndex = 1
    ' this gets rid of the extension on file ( everything from the "." to end)
    If InStr(daName, ".") > 0 Then
       daName = Left(daName, InStr(daName, ".") - 1)
    End If
    
    ' this gets rid of the part number in parenthesis by looking for a "(" and then reassigning the string name to everything to the left of that "("
    If InStr(daName, "(") > 0 Then
        daName = Left(daName, InStr(daName, "(") - 1)
    End If
    
    ' this adds the formatted date to the end of filename
    daName = daName & "(" & daDateFormatted & ")"
    
    ' Assign the value of daName to baseName so we can alter daName later
    baseName = daName
    
    daPDFfolder = "Test Data PDF"
    
    ' will only export if a file name exists
    If daName <> "" Then
    
        pdfPath = Application.ActiveWorkbook.Path & "\" & Fname & daPDFfolder & "\" & daName

        'Test if directory or file exists. if so, then add letters to end alphabetically starting with 'b'
        Do
            If FileOrDirExists(pdfPath & ".pdf") Then
                ' if it exists then we add a letter to the end from the "characters" string
                daName = baseName & Mid(characters, charIndex, 1)
                charIndex = charIndex + 1
                pdfPath = Application.ActiveWorkbook.Path & "\" & Fname & daPDFfolder & "\" & daName
            End If
        Loop While FileOrDirExists(pdfPath & ".pdf")
        
        ' if pdf folder not found this will save in current directory
        On Error GoTo handler:
        
        Selection.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfPath, _
        Quality:=xlQualityStandar, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
    End If
       
    Exit Sub
    
handler:

    pdfPath = Application.ActiveWorkbook.Path & "\" & Fname & baseName
    charIndex = 1
    Do
        If FileOrDirExists(pdfPath & ".pdf") Then
            daName = baseName & Mid(characters, charIndex, 1)
            charIndex = charIndex + 1
            pdfPath = Application.ActiveWorkbook.Path & "\" & Fname & "\" & daName
        End If
    Loop While FileOrDirExists(pdfPath & ".pdf")
    
    MsgBox ("PDF folder (" & daPDFfolder & ") not found so pdf exported to current folder")
    Selection.ExportAsFixedFormat _
    Type:=xlTypePDF, _
    Filename:=pdfPath, _
    Quality:=xlQualityStandar, _
    IncludeDocProperties:=True, _
    IgnorePrintAreas:=False, _
    OpenAfterPublish:=True
    
    
End Sub
' altered so it's easier to understand
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
     
     'if error then file does not exist so go to handler
    On Error GoTo handler:
    iTemp = GetAttr(PathName)
    
    'this only happens if no error occurs in assignment
    FileOrDirExists = True
    
    Exit Function
    
handler:
    FileOrDirExists = False
    
End Function
