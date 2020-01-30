

Public Sub OutlookMail( sheetName_ as string ,attachment As String, subject_ As String, body_ As String, toRange_ As String, Optional secondAttachment As String = " ")
    Dim OutApp As Object
    Dim OutMail As Object
    Dim recipients_ As String
    
    recipients_ = ThisWorkbook.Sheets(sheetName_).Range(toRange_).Value
    
    'The below loop should replace all of the commas "," with semicolons ";" in case the user uses them to separate email addresses
    Do While InStr(recipients_, ",") > 0
        Dim comIndex As Integer
        Dim lastIndex As Integer
        comIndex = InStr(recipients_, ",")
        lastIndex = Len(recipients_) - 1
        recipients_ = Left(recipients_, comIndex - 1) & ";" & Right(recipients_, lastIndex - comIndex)
    Loop
    
    Application.EnableEvents = False

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    'On Error Resume Next
    With OutMail
        .to = recipients_
        .CC = ""
        .BCC = ""
        .Subject = subject_
        .Body = body_
        .attachments.Add attachment
        If secondAttachment <> " " Then
            .attachments.Add secondAttachment
        End If
        .Display
    End With
           
    Set OutMail = Nothing
    Set OutApp = Nothing

    Application.EnableEvents = True

End Sub