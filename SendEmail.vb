Sub SendEmail()
' Process 4, Send Email
    
'
Set objExcel = CreateObject("Excel.Application")

' path to the excel file
Set objWorkbook = objExcel.Workbooks.Open("C:\Users\ekim\Desktop\Projects\hello\ITEMinTest.xlsx") ' Changed this

' Creating and Preparing a Log File
Set objOutput = CreateObject("Scripting.FileSystemObject") ' changed from objFTPOutput to objOutput
logFileName = "C:\Users\ekim\Desktop\Projects\hello\logs\Edfull4.log"
Set logFile = objOutput.CreateTextFile(logFileName, True)
logFile.Write "Process 4: Sending a File via Email - Datetime babyyy: " & Now() & vbCrLf & _
              "-------------------------------------------------------------" & vbCrLf

      
' Declaration of variables to use for this process
Dim file_left As String ' Includes Directory
Dim em_to As String
Dim em_cc As String
Dim em_bcc As String
Dim em_subj As String
Dim em_attach1 As String
Dim em_attach2 As String

Dim MAIN_PATH As String ' Directory only using file_left

Dim intSuccess As Integer
Dim intRow As Integer
intSuccess = 0

Dim xOutApp As Object
Dim xOutMail As Object
Dim xMailBody As String



' start at the second row, ie, not the column header
intRow = 2
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Check each record from the ITEMin.xlsx until it finds a blank row
Do Until objExcel.Cells(intRow, 1).Value = ""

' Test
Set xOutApp = CreateObject("Outlook.Application")
Set xOutMail = xOutApp.CreateItem(0)

' Fetch Email data from ITEMin.xlsx
send_to = Trim(objExcel.Cells(intRow, 1).Value)
em_to = Trim(objExcel.Cells(intRow, 2).Value)
em_cc = Trim(objExcel.Cells(intRow, 3).Value)
em_bcc = Trim(objExcel.Cells(intRow, 4).Value)
em_subj = Trim(objExcel.Cells(intRow, 5).Value)
xMailBody = Trim(objExcel.Cells(intRow, 6).Value) ' Email Body
em_attach1 = Trim(objExcel.Cells(intRow, 7).Value) ' C:\Users\ekim\Desktop\Projects\hello\ftpTestFiles\TestWord_2019-02-37.docx
em_attach2 = Trim(objExcel.Cells(intRow, 8).Value)
              
              
'MsgBox "Attach1: " & em_attach1 & vbCrLf & _
       "intRow: " & intRow & vbCrLf & _
       "send_to: " & send_to
       
Dim hasAttachment As Boolean
hasAttachment = True
              
' If Red Font, then just ignore and move to the next file (Red Font is Invalid ones from Process 1)
If objExcel.Cells(intRow, 1).Font.Color = RGB(255, 0, 0) Then
    logFile.Write "(" & (intRow - 1) & ") " & send_to & vbCrLf & _
                  "     Error: This item will be skipped as it's determined Invalid... (Red Font)" & vbCrLf & _
                  "     Error: This is most likely due to not being found in the MASTER Excel file." & vbCrLf
    GoTo NextIteration
ElseIf em_attach1 = "" Then
    logFile.Write "(" & (intRow - 1) & ") " & send_to & vbCrLf & _
                  "     Note: No File Needed for this Record" & vbCrLf
    hasAttachment = False
Else
    logFile.Write "(" & (intRow - 1) & ") " & send_to & vbCrLf & _
                  "     Note: File is Needed for this Record. " & vbCrLf

End If

' Check how Weekly Reports send data and copy the logic

If hasAttachment = True Then
    With xOutMail
            .To = em_to
            .CC = em_cc
            .BCC = em_bcc
            .Subject = em_subj
            .htmlBody = xMailBody
            .attachments.Add (em_attach1)
    '        .attachments.Add ("\\viper\data\Sales\Public\CPFR\Alteryx\POSReportsSQL\Outputs\FinalOutputs\" & SatDay & "\POSWeeklyUnitsPerCustomer_Thru" & SatDay & "_w_AZCC_Dummy_Thru_" & AZCCDay & ".xlsx")
            .Send   'or use .Send
    End With
ElseIf hasAttachment = False Then
    With xOutMail
            .To = em_to
            .CC = em_cc
            .BCC = em_bcc
            .Subject = em_subj
            .htmlBody = xMailBody
            '.attachments.Add (em_attach1)
    '        .attachments.Add ("\\viper\data\Sales\Public\CPFR\Alteryx\POSReportsSQL\Outputs\FinalOutputs\" & SatDay & "\POSWeeklyUnitsPerCustomer_Thru" & SatDay & "_w_AZCC_Dummy_Thru_" & AZCCDay & ".xlsx")
            .Send   'or use .Send
                        
    End With
End If


NextIteration:
logFile.Write vbCrLf
intRow = intRow + 1

Loop
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Create Another loop to check emails were sent in your Outlook 'Sent' Folder
'' start at the second row, ie, not the column header
'' Check each record from the ITEMin.xlsx until it finds a blank row
'Dim intRow2 As Integer
'intRow2 = 2
'Do Until objExcel.Cells(intRow2, 1).Value = ""
'
'' Test
'Set xOutApp = CreateObject("Outlook.Application")
'Set xOutMail = xOutApp.CreateItem(0)
'
'' Fetch Email data from ITEMin.xlsx
'send_to = Trim(objExcel.Cells(intRow2, 1).Value)
'em_to = Trim(objExcel.Cells(intRow2, 2).Value)
'em_cc = Trim(objExcel.Cells(intRow2, 3).Value)
'em_bcc = Trim(objExcel.Cells(intRow2, 4).Value)
'em_subj = Trim(objExcel.Cells(intRow2, 5).Value)
'xMailBody = Trim(objExcel.Cells(intRow2, 6).Value) ' Email Body
'em_attach1 = Trim(objExcel.Cells(intRow2, 7).Value) ' C:\Users\ekim\Desktop\Projects\hello\ftpTestFiles\TestWord_2019-02-37.docx
'em_attach2 = Trim(objExcel.Cells(intRow2, 8).Value)
'
'intRow2 = intRow2 + 1
'Loop
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Log Footer
logFile.Write "-------------------------------------------------------------" & vbCrLf
logFile.Write "Total Items Checked: " & (intRow - 2) & vbCrLf
    
MsgBox "Process 4 is Finished!"

logFile.Close
objWorkbook.Close saveChanges:=False
objExcel.Quit

End Sub




