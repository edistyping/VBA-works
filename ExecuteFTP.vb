Sub ExecuteFTP()

Set Wshell = CreateObject("WScript.Shell")
Set objExcel = CreateObject("Excel.Application")
Set objFTPOutput = CreateObject("Scripting.FileSystemObject")
Set objFTPFSO = CreateObject("Scripting.FileSystemObject")

' path to the excel file
Set objWorkbook = objExcel.Workbooks.Open("C:\Users\ekim\Desktop\Projects\hello\ITEMinTest.xlsx")

' path to full output log file
logFile = "C:\Users\ekim\Desktop\Projects\hello\ftpTestFiles\full3.log"

' clear the ftp output file
Set objFTPOutputFile = objFTPOutput.CreateTextFile(logFile, True)
objFTPOutputFile.Write Now() & vbCrLf
objFTPOutputFile.Close

' path to simple output log file
simplelogFile = "C:\Users\ekim\Desktop\Projects\hello\ftpTestFiles\simple3.log"


'''''''''''''''''''''''''''''''''''''''
' Ed's log
Set objOutputTest = CreateObject("Scripting.FileSystemObject") ' changed from objFTPOutput to objOutput
testlogFileName = "C:\Users\ekim\Desktop\Projects\hello\ftpTestFiles\fulltest3.log"
Set objTestLog = objOutputTest.CreateTextFile(testlogFileName, True)
objTestLog.Write "Process 3   -   Datetime of Log Creation: " & Now() & vbCrLf & _
              "-------------------------------------------------------------" & vbCrLf
''''''''''''''''''''''''''''''''''


' path to dat file
datFile = "C:\Users\ekim\Desktop\Projects\hello\ftpTestFiles\ftpcmd.dat"
      
' start at the second row, ie, not the column header
intRow = 2

' Check each record from the ITEMin.xlsx until it finds a blank row
Do Until objExcel.Cells(intRow, 1).Value = ""

' open the ftp dat file for writing
Set objFile = objFTPFSO.CreateTextFile(datFile, True)

' If Red Font, then just ignore and move to the next file (Red Font is Invalid ones from Process 1)
If objExcel.Cells(intRow, 1).Font.Color = RGB(255, 0, 0) Then
    objTestLog.Write "Error Occured for " & objExcel.Cells(intRow, 1).Value & " (Row " & intRow & ") - " & Trim(objExcel.Cells(intRow, 21).Value) & vbCrLf
    objFile.Close
    GoTo NextIteration
End If


' Extract all data from ITEMin
file_name = Trim(objExcel.Cells(intRow, 9).Value)
ftp_server = Trim(objExcel.Cells(intRow, 15).Value)
If ftp_server = "" Then
    Skip = "Y"
Else
    Skip = "N"
End If
ftp_user = Trim(objExcel.Cells(intRow, 16).Value)
ftp_pass = Trim(objExcel.Cells(intRow, 17).Value)
ftpPort = Trim(objExcel.Cells(intRow, 18).Value)
upload_path = Trim(objExcel.Cells(intRow, 19).Value & objExcel.Cells(intRow, 9).Value)
server_path = Trim(objExcel.Cells(intRow, 20).Value)
protocol = Trim(objExcel.Cells(intRow, 21).Value)


' if the clients protocol is ftps, sftp, or standard we have to
' build a different ftpcmd.dat file
If protocol = "FTPS" Then
    ftp_mode = "FTPS"
    
    ' if the port is blank, use port 21, if it is not blank, use the port specified
    If objExcel.Cells(intRow, 18).Value = "" Then
        port = 21
    Else
        port = objExcel.Cells(intRow, 18).Value
    End If
    
    objFile.Write "option echo on" & vbCrLf ''' test
    
    objFile.Write "option batch on" & vbCrLf
    objFile.Write "option confirm off" & vbCrLf
    objFile.Write "open ftps://" & Replace(ftp_user, " ", "%20") & ":" & ftp_pass & "@" & ftp_server & ":" & port & vbCrLf
    ' if the server path is blank, then just put the file, else cd to the directory
    If server_path <> "" Then
        objFile.Write "cd " & """" & server_path & """" & vbCrLf
    End If
    objFile.Write "put " & """" & upload_path & """" & vbCrLf
    objFile.Write "close" & vbCrLf
    objFile.Write "exit" & vbCrLf
ElseIf protocol = "SFTP" Then
    ftp_mode = "SFTP"
    
    ' if the port is blank, use port 22, if it is not blank, use the port specified
    If objExcel.Cells(intRow, 18).Value = "" Then
        port = 22
    Else
        port = objExcel.Cells(intRow, 18).Value
    End If

    ' if the server requires hostkey, grab it from the excel document
    If objExcel.Cells(intRow, 22).Value <> "" Then
        hostkey_command = "-hostkey=""" & Trim(objExcel.Cells(intRow, 22).Value) & """"
    Else
        hostkey_command = ""
    End If
    
    objFile.Write "option echo on" & vbCrLf ' Test''fdgpgpigpi

    objFile.Write "option batch on" & vbCrLf
    objFile.Write "option confirm off" & vbCrLf
    objFile.Write "open sftp://" & Replace(ftp_user, " ", "%20") & ":" & Replace(Replace(ftp_pass, "+", "%2B"), "@", "%40") & "@" & ftp_server & ":" & port & " " & hostkey_command & vbCrLf
    ' if the server path is blank, then just put the file, else cd to the directory
    If server_path <> "" Then
        objFile.Write "cd " & """" & server_path & """" & vbCrLf
    End If
    'objFile.Write "put " & """" & upload_path & """ -nopreservetime" & vbCrL
    objFile.Write "put " & """" & upload_path & """ -nopreservetime" & vbCrLf
    objFile.Write "close" & vbCrLf
    objFile.Write "exit" & vbCrLf
Else
    ftp_mode = "FTP"
    
    ' if the port is blank, use port 21, if it is not blank, use the port specified
    If objExcel.Cells(intRow, 18).Value = "" Then
        port = 21
    Else
        port = objExcel.Cells(intRow, 18).Value
    End If

    ' write the information to the ftp dat file
    objFile.Write "open " & ftp_server & " " & port & vbCrLf
    objFile.Write ftp_user & vbCrLf
    objFile.Write ftp_pass & vbCrLf
    objFile.Write "binary " & vbCrLf

    ' if the server path is blank, then just put the file, else cd to the directory
    If server_path <> "" Then
        objFile.Write "cd " & """" & server_path & """" & vbCrLf
    End If
    objFile.Write "put " & """" & upload_path & """" & vbCrLf
    objFile.Write "disconnect" & vbCrLf
    objFile.Write "quit" & vbCrLf
End If

'''''''''''''''''''''''''''
' !!! TESTING for FTP!
'

''''' Since objFTPOutputFile was closed earlier, we need to open the logFile again
''''''''''''''' Log Test
Set objFTPOutputFile = objFTPOutput.OpenTextFile(logFile, 8, -2)
objFTPOutputFile.Write "***Now processing: " & objExcel.Cells(intRow, 1).Value & " (Row " & intRow & ") - " & Trim(objExcel.Cells(intRow, 21).Value) & vbCrLf
objFTPOutputFile.Close


' !!! TESTING section ending...
'''''''''''''''''''''''''''''''''''''''''''''
' Call Shell("C:\Users\ekim\Desktop\Projects\hello\WinSCP.com /ini=nul /script=C:\Users\ekim\Desktop\Projects\hello\ftpTestFiles\ftpcmd.dat")
  'Call Shell("C:\Users\ekim\Desktop\Projects\hello\WinSCP.com /ini=nul /command ""open wefjwf"" ")
 
 'SFTP = Wshell.Run("C:Users/ekim/Desktop/Projects/hello/WinSCP.com")

'C:\Users\ekim\AppData\Local\WinSCP.com
' if the server name is blank, do not ftp the file
If Skip = "N" Then
    
' Just initialized with temp value for us to check response/return value from FTP
ftp = 999
FTPS = 999
SFTP = 999

' Close the ftp dat file (Needed before we use its associated datFile in Run() below
objFile.Close
   
Dim gotFTPRespond As Boolean
gotFTPRespond = False

    If ftp_mode = "FTP" Then
        ftp = Wshell.Run("%comspec% /c ftp -d -i -s:""" & datFile & """>>""" & logFile & """ ", 0, True)
        gotFTPRespond = True
    ElseIf ftp_mode = "FTPS" Then
        FTPS = Wshell.Run("C:/Users/ekim/Desktop/Projects/hello/WinSCP.com /script=" & datFile & " /log=" & logFile, 0, True)  ' Good
        gotFTPRespond = True
    ElseIf ftp_mode = "SFTP" Then
        'SFTP = Wshell.Run("C:Users/ekim/Desktop/Projects/hello/WinSCP.com /script=""" & datFile & """ /log=""" & logFile & """", 1, True)
        SFTP = Wshell.Run("C:Users/ekim/Desktop/Projects/hello/WinSCP.com /script=""" & datFile & """ & /log=""" & logFile & """ ", 0, True)
        gotFTPRespond = True
    End If
       
    
End If

''''''''''''''' Log Test

' For Logging...
Set objFTPOutputFile = objFTPOutput.OpenTextFile(logFile, 8, -2)
If ftp <> 999 And gotFTPRespond = True Then
    objFTPOutputFile.Write "***Return Code: " & ftp & vbCrLf
ElseIf FTPS <> 999 And gotFTPRespond = True Then
    objFTPOutputFile.Write "***Return Code: " & FTPS & vbCrLf
ElseIf SFTP <> 999 And gotFTPRespond = True Then
    objFTPOutputFile.Write "***Return Code: " & SFTP & vbCrLf
End If
objFTPOutputFile.Write "***Finished Processing..." & vbCrLf & vbCrLf
objFTPOutputFile.Close

NextIteration:

' Onto the next record in ITEMin
intRow = intRow + 1
Loop

' Close the excel file
objWorkbook.Close
objExcel.Quit


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Below are purely for Logging purpose
' take a look at the log file and check for errors
' open the simple log file for writing
Set objSimpleFile = objFTPFSO.CreateTextFile(simplelogFile, True)

' read the full log file for error and add the line from this log to the simple log
Set objLogFile = objFTPOutput.OpenTextFile(logFile)

' These are for Log
Dim fileName As String
Dim ftpType As String
Dim prevLine As String
Dim logRow As Integer
Dim logSkip As Boolean
logRow = 1
prevLine = "zero"
logSkip = False

Do Until objLogFile.AtEndOfStream
    
    strLine = LCase(objLogFile.ReadLine)
    
    ' Check the FTP Type for accurate logging purpose
    If InStr(strLine, "now processing") >= 1 Then
        objTestLog.Write vbCrLf
        objTestLog.Write "***Reading Line strLine(~5): " & Mid(strLine, InStr(strLine, ":") + 2) & vbCrLf ' Test File
        ftpType = Mid(strLine, InStrRev(strLine, "-") + 2)
        logSkip = False
    'ElseIf InStr(strLine, "finished processing") >= 1 Then ' I don't think this is needed
        'ftpType = ""
    End If
        
    ' For corresponding FTP type, process logging
    If ftpType = "ftp" And logSkip = False Then ' For FTP
        If InStr(strLine, "unknown host") >= 1 Then
            objTestLog.Write "  Error: Unknown or Invalid Host Address. Please Check Again" & vbCrLf
            logSkip = True
            ' Perhaps just skip out of this one? (as in go to ***Finished Processing line?)
        ElseIf InStr(strLine, "530") >= 1 Then
            objTestLog.Write "  Error: Wrong ID or Password! Please Check Again" & vbCrLf
            logSkip = True
        ElseIf InStr(strLine, "226") >= 1 Then
            objTestLog.Write "  Success: File has been successfully transferred" & vbCrLf
        End If
    ElseIf ftpType = "ftps" And logSkip = False Then ' For FTPS
        If InStr(strLine, "TLS connection established") >= 1 Then
            objTestLog.Write "  Error: Unknown or Invalid Host Address. Please Check Again" & vbCrLf
            logSkip = True
        ElseIf InStr(strLine, "connection failed") >= 1 Then
            objTestLog.Write "  Error: Connection Failed. Please Check the FTP Address or Port" & vbCrLf
            logSkip = True
        ElseIf InStr(strLine, "530 login or password incorrect!") >= 1 Then
            objTestLog.Write "  Error: Login or Password is Incorrect! Please Check Again1" & vbCrLf
            logSkip = True
        ElseIf InStr(strLine, "transfer done") >= 1 Then
            objTestLog.Write "  Success: File is Successfully Transferred" & vbCrLf
        End If
    ElseIf ftpType = "sftp" And logSkip = False Then ' For SFTP
        If InStr(strLine, "access granted") >= 1 Then ' Eventually change this to File Transffered instead of access granted
            objTestLog.Write "  Note: Connection is Successfully Made! Now, Attempting to Transfer the File..." & vbCrLf
        ElseIf InStr(strLine, "the system cannot find the file specified") >= 1 Then
            objTestLog.Write "  Error: File to Send Was Not Found" & vbCrLf
            logSkip = True
        ElseIf InStr(strLine, "password authentication failed") >= 1 Then
            objTestLog.Write "  Error: Login or Password is Incorrect! Please Check Again1" & vbCrLf
            logSkip = True
        ElseIf InStr(prevLine, "looking up host") >= 1 And InStr(strLine, "finished processing") >= 1 Then '(Looking up host ...)'s next line is (***finished processing...) then no connection is made
            objTestLog.Write "  Error: Connection wasn't Made" & vbCrLf
            logSkip = True
        ElseIf InStr(prevLine, "transfer done") >= 1 And InStr(strLine, "transfer successfully") >= 1 Then ' In Worst scenario, use "Return Code" to find Success/Fail -> Actually maybe not. FTP one returned 0 on Fail... wat
            objTestLog.Write "  Success: File is Successfully Transferred" & vbCrLf
        End If
    End If
    
    prevLine = strLine
    logRow = logRow + 1
    
Loop


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

ToEnd:
' Close the excel file

MsgBox "Process 3 is Ended"

End Sub




