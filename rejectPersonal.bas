Attribute VB_Name = "rejectPersonal"
Option Explicit  ' require variable declarations

' Global variables
Dim technician
Dim shiftMixed
Dim keepGoing
Dim customerName
Dim customerCode
Dim lmiCode
Dim batches
Dim mixerNumber
Dim rejLbs
Dim reason
Dim currentRow
Dim columnCheck As Integer
Dim rejectCheck


Sub RejectTagLogBryan()
Attribute RejectTagLogBryan.VB_ProcData.VB_Invoke_Func = "r\n14"
    '  Bryan McElroy : 9-2018
    '  This code is to facilitate reject tag log completion
    '
    currentRow = ActiveCell.Row
    ' **************************************
    ' your initials below in quotes
    technician = "BM"
    ' your shift below
    shiftMixed = 3
    ' **************************************
    
    keepGoing = 6 ' yes in other words
    customerName = "LMI"
    customerCode = "customer code"
    lmiCode = "LMI Compound Number"
    batches = "#"
    mixerNumber = 2
    rejLbs = "#"
    reason = "cuz"
    
            
    
    Do While keepGoing = 6
            Application.ScreenUpdating = False
            
            ' This is here to make sure that the reject-hold tag log file is actually open
            ' If it is not then the code jumps to the Handler at bottom of code
            On Error GoTo Handler:
            
                rejectCheck = Workbooks("Reject-Hold Tag Log.xls").Worksheets("RMA's").Cells(currentRow, 1).Value
                
            On Error GoTo 0
            
            columnCheck = 1
            Do While columnCheck < 20
                If Cells(currentRow, columnCheck) <> "" Then
                    keepGoing = 7
                    columnCheck = 20
                    MsgBox ("There is already data in row #" & currentRow & ", Please clear or choose another row and try again")
                End If
                columnCheck = columnCheck + 1
            Loop
            
            If keepGoing = 7 Then
                Exit Do
            End If
            
            customerName = InputBox("Customer", "Customer", customerName)
            ' customer check IF
            If customerName = "quit" Or customerName = "" Then
                keepGoing = 7
                Exit Do
            End If ' customer IF
                   
            customerCode = InputBox("Customer Compound Number", "Customer Compound", customerCode)
            ' customer code IF
            If customerCode = "quit" Or customerCode = "" Then
                keepGoing = 7
                Exit Do
            End If
            
            lmiCode = InputBox("LMI Compound Number", "Our Code", lmiCode)
            ' lmiCode IF
            If lmiCode = "quit" Or lmiCode = "" Then
                keepGoing = 7
                Exit Do
            End If
            
            batches = InputBox("Batches", "Batch Numbers", batches)
            ' batches IF
            If batches = "quit" Or batches = "" Then
                keepGoing = 7
                Exit Do
            End If
            
            mixerNumber = InputBox("Mixer #", "Mixer", mixerNumber)
            ' mixerNumber IF
            If mixerNumber = "quit" Or mixerNumber = "" Then
                keepGoing = 7
                Exit Do
            End If
            
            rejLbs = InputBox("Rejected lbs", "Bad Rubber", rejLbs)
            ' rejLbs IF
            If rejLbs = "quit" Or rejLbs = "" Then
                keepGoing = 7
                Exit Do
            End If
            
            reason = InputBox("Reason for rejection?", "Just cuz", reason)
            ' reason IF
            If reason = "quit" Or reason = "" Then
                keepGoing = 7
                Exit Do
            End If
            
            
            ' enter all data into the cells
            With Application.Workbooks("Reject-Hold Tag Log.xls").Worksheets("RMA's")
                Cells(currentRow, 2).Value = Date
                Cells(currentRow, 10).Value = Date
                Cells(currentRow, 19).Value = Date
                Cells(currentRow, 3).Value = technician
                Cells(currentRow, 4).Value = shiftMixed
                Cells(currentRow, 5).Value = customerName
                Cells(currentRow, 6).Value = customerCode
                Cells(currentRow, 7).Value = lmiCode
                Cells(currentRow, 8).Value = batches
                Cells(currentRow, 9).Value = mixerNumber
                Cells(currentRow, 14).Value = rejLbs
                Cells(currentRow, 16).Value = reason
            End With
            Application.ScreenUpdating = True
            currentRow = currentRow + 1
            ' creates msg box with "yes" and "No" options - yes = 6, no = 7 in vba
            keepGoing = MsgBox("More to add?", 4)
                                                        
        Loop
    '
    ' Exit sub is needed so handler is not ran with code
        Exit Sub
        
Handler:
        MsgBox ("Make sure Reject hold tag log is open")
                
    

End Sub
