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

''''''''''''''''''''''''''''''''''
Set objOutputTest = CreateObject("Scripting.FileSystemObject") ' changed from objFTPOutput to objOutput
testlogFileName = "C:\Users\ekim\Desktop\Projects\hello\ftpTestFiles\fulltest3.log"
Set testlog = objOutputTest.CreateTextFile(testlogFileName, True)
testlog.Write "Process 3   -   Datetime of Log Creation: " & Now() & vbCrLf & _
              "-------------------------------------------------------------" & vbCrLf


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
    testlog.Write "Error Occured for " & objExcel.Cells(intRow, 1).Value & " (Row " & intRow & ")"
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
objFTPOutputFile.Write "***Now processing: " & objExcel.Cells(intRow, 1).Value & " (Row " & intRow & ")" & vbCrLf
objFTPOutputFile.Close

' !!! TESTING section ending...
'''''''''''''''''''''''''''''''''''''''''''''
' Call Shell("C:\Users\ekim\Desktop\Projects\hello\WinSCP.com /ini=nul /script=C:\Users\ekim\Desktop\Projects\hello\ftpTestFiles\ftpcmd.dat")
  'Call Shell("C:\Users\ekim\Desktop\Projects\hello\WinSCP.com /ini=nul /command ""open wefjwf"" ")
 
 'SFTP = Wshell.Run("C:Users/ekim/Desktop/Projects/hello/WinSCP.com")

'C:\Users\ekim\AppData\Local\WinSCP.com
' if the server name is blank, do not ftp the file
If Skip = "N" Then
    
    ftp = 333
    FTPS = 333
    SFTP = 333

' Close the ftp dat file   !!! This was hte issue of FTPS stuff. I guess since the datFile was still opened and connetcted _
objFile variable, it's causing problem
objFile.Close
   

    If ftp_mode = "FTP" Then
        ftp = Wshell.Run("%comspec% /c ftp -d -i -s:""" & datFile & """>>""" & logFile & """ ", 1, True)
    ElseIf ftp_mode = "FTPS" Then
        'FTPS = Wshell.Run("C:/Users/ekim/Desktop/Projects/hello/WinSCP.com /int=nul /script=C:/Users/ekim/Desktop/Projects/hello/ftpTestFiles/ftpcmd.dat", 1, True)  ' Good
        
        ' it stopped working all of sudden wtf. Investigate
' (OG)  FTPS = WShell.Run("WinSCP.com /script=""" & datFile & """ /log=""" & logFile & """", 0, true)
         FTPS = Wshell.Run("C:/Users/ekim/Desktop/Projects/hello/WinSCP.com /script=" & datFile & " /log=" & logFile, 1, True)  ' Good
                
    ElseIf ftp_mode = "SFTP" Then
        'SFTP = Wshell.Run("C:Users/ekim/Desktop/Projects/hello/WinSCP.com /script=""" & datFile & """ /log=""" & logFile & """", 1, True)
        SFTP = Wshell.Run("C:Users/ekim/Desktop/Projects/hello/WinSCP.com /script=""" & datFile & """ & /log=""" & logFile & """ ", 1, True)
    End If
'
'    MsgBox "ftp: " & ftp & vbCrLf & _
'           "FTPS: " & FTPS & vbCrLf & _
'           "SFTP: " & SFTP & vbCrLf
'
'
'    If ftp = 1 Or FTPS = 1 Or SFTP = 1 Then
'        MsgBox objExcel.Cells(intRow, 1).Value & "---- 1"
'    ElseIf ftp = 0 Or FTPS = 0 Or SFTP = 0 Then
'        MsgBox objExcel.Cells(intRow, 1).Value & "---- 0"
'    End If
    
    
End If

''''''''''''''' Log Test
Set objFTPOutputFile = objFTPOutput.OpenTextFile(logFile, 8, -2)
objFTPOutputFile.Write "***Finished Processing..." & vbCrLf & vbCrLf
objFTPOutputFile.Close
''''''''''!@#@!$#$@$!#$@#@!#@!$@!$@!$@!$

NextIteration:

' Close the ftp dat file
'objFile.Close




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

Do Until objLogFile.AtEndOfStream
    strLine = LCase(objLogFile.ReadLine)
    
    If InStr(strLine, "now processing") >= 1 Then
        objSimpleFile.Write "***We are now processing a file" & vbCrLf
    End If
    
    ' Provide (send_to, Server, Username/Password, Binary type, Put statement)
    
    
    
    ' depending on what error is found, try to display the correct error message
    If InStr(strLine, "530 login incorrect") >= 1 Or InStr(strLine, "530 you aren't logged in") >= 1 Or InStr(strLine, "service not available") >= 1 Or InStr(strLine, "unable to authenticate") >= 1 Or InStr(strLine, "login or password incorrect") >= 1 Then
       objSimpleFile.Write "One or more of the client's has username/password problems." & vbCrLf
    ElseIf InStr(strLine, "file not found") >= 1 Or InStr(strLine, "the system cannot find the file specified") >= 1 Or InStr(strLine, "the system cannot find the path specified") >= 1 Then
       objSimpleFile.Write "One or more of the FTP files could not be found." & vbCrLf
    ElseIf InStr(strLine, "unknown host") >= 1 Or InStr(strLine, "connection failed") >= 1 Then
        objSimpleFile.Write "One or more of the client's has something wrong with the server name." & vbCrLf
    ElseIf InStr(strLine, "can't change directory to") >= 1 Then
        objSimpleFile.Write "One or more of the client's has something wrong with the remote server directory." & vbCrLf
    ElseIf InStr(strLine, "network error") >= 1 Or InStr(strLine, "ftp port did not open") >= 1 Or InStr(strLine, "system error") >= 1 Then
        objSimpleFile.Write "One or more client's had an unknown error." & vbCrLf
    ElseIf InStr(strLine, "exception") >= 1 Then
        objSimpleFile.Write "Potential host key problem." & vbCrLf
    End If
    
    If InStr(strLine, "finished processing") >= 1 Then
        objSimpleFile.Write "***Processing finished for the file..." & vbCrLf & vbCrLf
    End If
    


Loop
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

ToEnd:
' Close the excel file

MsgBox "Process 3 is Ended"

End Sub

