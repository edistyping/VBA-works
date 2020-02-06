Sub compareWithMaster()

' Fetch the data from ITEMin.xlsx
Dim send_to As String
Dim file_date As String
Dim file_left As String
Dim email_to As String
Dim email_cc As String
Dim email_bcc As String
Dim file_fullname As String

' Fetch the data from MASTER.xlsx
Dim send_to2 As String
Dim file_date2 As String
Dim file_left2 As String
Dim email_to2 As String
Dim email_cc2 As String
Dim email_bcc2 As String
Dim file_fullname2 As String

Dim intRow As Integer
Dim intSuccess As Integer
Dim logFlags(5) As String

Dim keepSearching As Boolean

' create system objects for file handling
Set objExcel = CreateObject("Excel.Application")
Set objExcel2 = CreateObject("Excel.Application")

' For testing
'objExcel.Visible = True
'objExcel2.Visible = True

' path to the excel file
Set objWorkbook = objExcel.Workbooks.Open("C:\Users\ekim\Desktop\Projects\hello\ITEMinTest.xlsx")
Set objWorkbook2 = objExcel2.Workbooks.Open("C:\Users\ekim\Desktop\Projects\hello\MASTERTEST2.xlsx")

' Log file stuffs - Initially create/clear it
Set objOutput = CreateObject("Scripting.FileSystemObject") ' changed from objFTPOutput to objOutput
logFileName = "C:\Users\ekim\Desktop\Projects\hello\fulltest1.log"
Set logFile = objOutput.CreateTextFile(logFileName, True)
logFile.Write "Process 1   -   Datetime of Log Creation: " & Now() & vbCrLf & _
              "-------------------------------------------------------------" & vbCrLf

' start at the second row, ie, not the column header
intRow = 2

' Below for analysis purpose
intSucess = 0

' loop through each row of the excel file
Do Until objExcel.Cells(intRow, 1).Value = ""

   
' Fetch the data to compare it with Master later
file_date = Trim(objExcel.Cells(intRow, 10).Value)
file_left = Trim(objExcel.Cells(intRow, 14).Value)
send_to = Trim(objExcel.Cells(intRow, 1).Value)
email_to = Trim(objExcel.Cells(intRow, 2).Value)
email_cc = Trim(objExcel.Cells(intRow, 3).Value)
email_bcc = Trim(objExcel.Cells(intRow, 4).Value)
file_fullname = Trim(objExcel.Cells(intRow, 9).Value)

Dim dataFound As Boolean
dataFound = False

' Reset for the next line
dataFound = False

''''''''''''MsgBox "Now doing... " & send_to & vbCrLf & _
       "    " & file_left

' loop against the MASTER file
'''''''''''''' ISSUEEEEEEEEEE: Only gets call for first outer loop
For Each sh In objWorkbook2.Worksheets
    If sh.Index = 1 Then GoTo NextSheet ' Skip the first 'Uniq Customer Projects' in the MASTER file
     '   MsgBox "Searching " & send_to & vbCrLf & _
               "Sh Index: " & sh.Index
        
    ' In each worksheet in MASTER, check each row for the matching record
    For Each rw In sh.Rows
    
        logFlags(0) = "0"
        logFlags(1) = "0"
        logFlags(2) = "0"
        logFlags(3) = "0"
        logFlags(4) = "0"
        logFlags(5) = "0"
        
        keepSearching = False
        
        'If is in Red Highlight, then Ignore this Inactive one
        If rw.Cells(1).Interior.Color = RGB(255, 0, 0) And rw.Cells(2).Interior.Color = RGB(255, 0, 0) Then
            GoTo NextRow
        End If
                
        ' Check further comparison if send_to data matches
        send_to2 = Trim(rw.Cells(1).Value)
        
        ' MsgBox "    wtfwtf:" & send_to2
        
        If send_to <> send_to2 Or send_to2 = "" Then
            logFlags(0) = "1"
            logFlags(1) = "1"
            GoTo NextRow
        End If
 
        
        If send_to = send_to2 Then
            
            ' Fetch the data to compare it with Master later
            email_to2 = Trim(rw.Cells(2).Value)
            email_cc2 = Trim(rw.Cells(3).Value)
            email_bcc2 = Trim(rw.Cells(4).Value)
            file_date2 = Trim(rw.Cells(10).Value)
            file_left2 = Trim(rw.Cells(14).Value)

            
            ' Below checks for any failure/unmatched. If so, log this bad boys
            If file_left <> file_left2 Then
                keepSearching = True
                logFlags(1) = "1"
                GoTo NextRow ' Since it's bad just skip it (Error: Wrong Filename!)
            End If
            'If file_date <> file_date2 Then keepSearching = True: logFlags(2) = "1"
            If email_to <> email_to2 Then keepSearching = True: logFlags(3) = "1"
            If email_cc <> email_cc2 Then keepSearching = True: logFlags(4) = "1"
            If email_bcc <> email_bcc2 Then keepSearching = True: logFlags(5) = "1"
            
            ' Below checks if a correct record is found in MASTER. If so, log this good too
            If keepSearching = False Then
                dataFound = True
                Exit For
            End If
                        
        End If
        
        
NextRow:
        ' Increase if you must, Luke Skypwalker
        If rw.Row >= 150 Then
            ''''''''MsgBox "wtf... " & rw.Row
            Exit For
        End If
    Next rw
    

NextSheet:
    ' If the matching record is found in MASTER, then we just move to next item in ITEMin file.
    If dataFound = True Then
        intSuccess = intSuccess + 1
        Exit For
    End If
    
   ''''''' MsgBox "Going to Next sheet"
Next sh


' Logging for Success/Fail
If logFlags(0) = "1" And logFlags(1) = "1" Then
    ' Output to Log
    logFile.Write "(" & (intRow - 1) & ") " & send_to & " | " & file_left & vbCrLf & _
                  "     Status: Not Found!" & vbCrLf & _
                  "         Error: Looks like there wasn't a matching record in Master Excel file... Please check again!" & vbCrLf & _
                  "     ---> FAIL" & vbCrLf

    ' Additional details to logs if applicable
    'If logFlags(2) = "1" Then logFile.Write "      File Date is not matching!" & vbCrLf
    If logFlags(3) = "1" Then logFile.Write "      CC Email is not matching!" & vbCrLf
    If logFlags(4) = "1" Then logFile.Write "      BCC Email is not matching!" & vbCrLf
    If logFlags(5) = "1" Then logFile.Write "      File Fullname is not matching!" & vbCrLf

Else
    ' Output to Log
    logFile.Write "(" & (intRow - 1) & ") " & send_to & " | " & file_left & vbCrLf & _
                  "     Status: Found! " & vbCrLf & _
                  "     Sheet: " & sh.Name & "  Row: " & rw.Row & vbCrLf & _
                  "        Detail below: " & vbCrLf & _
                  "        send_to: " & send_to & vbCrLf & _
                  "        file_left: " & file_left & vbCrLf & _
                  "        email_cc: " & email_cc & vbCrLf & _
                  "        email_bcc: " & email_bcc & vbCrLf & _
                  "        file_date: " & file_date & vbCrLf & _
                  "     ---> SUCCESS!" & vbCrLf
End If


' This coloring is needed for Phase 2 to update Wo0rds
' If the data is NOT found in MASTER file, the row will be fonted in red.
If dataFound = True Then
    objExcel.Range(objExcel.Cells(intRow, 1), objExcel.Cells(intRow, 24)).Font.Color = RGB(0, 0, 0)
    
    ' Since its Word File is updated, we update the MASTER file as well
    objExcel2.Worksheets(sh.Name).Cells(rw.Row, 10).Value = file_date
    objExcel2.Worksheets(sh.Name).Cells(rw.Row, 10).Font.Color = RGB(0, 0, 255) ' Updated is in Blue
    
Else ' Not Found in MASTER -> Change the Font in ITEMin
    objExcel.Range(objExcel.Cells(intRow, 1), objExcel.Cells(intRow, 24)).Font.Color = RGB(255, 0, 0)
End If


' Onto the Next Loop! (Next record in ITEMin)
logFile.Write vbCrLf
intRow = intRow + 1
Loop

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Output the Success/Failure Rate to Log File
logFile.Write "---------------------------------------------------------" & vbCrLf
logFile.Write "Total Record: " & CStr(intRow - 2) & vbCrLf & _
       "Success: " & CStr(intSuccess) & vbCrLf & _
       "Failure: " & CStr(intRow - 2 - intSuccess) & vbCrLf & _
       "*Note: For Successful " & CStr(intSuccess) & " records, Dates will be updated in MASTER File!"

' Output the Success/Failure Rate as a Pop-up message
MsgBox "Total Record: " & CStr(intRow - 2) & vbCrLf & _
       "Success: " & CStr(intSuccess) & vbCrLf & _
       "Failure: " & CStr(intRow - 2 - intSuccess)

objWorkbook.Close saveChanges:=True
objExcel.Quit
objWorkbook2.Close saveChanges:=True
objExcel2.Quit

End Sub



