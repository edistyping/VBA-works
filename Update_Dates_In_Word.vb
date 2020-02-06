Sub updateDateWord()

Application.ScreenUpdating = False

' This one simply update dates in the MS Word title and also the file
    ' Create a new file with updated Dates
    ' Store oldFiles to a subfolder named oldFiles (Might switch to Archives)
    ' Include logs for both Success and Failure
    ' Note: This process will skip INVALID records that were determined from the previous Process.
    
' Declaration of variables to use for this process
Dim new_fileDate As String
Dim file_left As String ' Includes Directory
Dim new_fileName As String ' file_left + new Date
Dim MAIN_PATH As String ' Directory only using file_left
Dim intSuccess As Integer
Dim intRow As Integer


' Open the ITEMin Excel file to fetch the updated Date
Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open("C:\Users\ekim\Desktop\Projects\hello\ITEMinTest.xlsx")

' Log file stuffs - Initially create/clear it
Set objOutput = CreateObject("Scripting.FileSystemObject") ' changed from objFTPOutput to objOutput
logFileName = "C:\Users\ekim\Desktop\Projects\hello\fulltest2.log"
Set logFile = objOutput.CreateTextFile(logFileName, True)
logFile.Write "Process 2    -   Datetime babyyy: " & Now() & vbCrLf & _
              "-------------------------------------------------------------" & vbCrLf

' loop through each row of the excel file
intSuccess = 0
intRow = 2
Do Until objExcel.Cells(intRow, 1).Value = ""

send_to = Trim(objExcel.Cells(intRow, 1).Value) '
new_fileDate = Trim(objExcel.Cells(intRow, 10).Value)
file_left = Trim(objExcel.Cells(intRow, 24).Value) ' Includes its directory
new_fileName = file_left + new_fileDate + ".docx" ' Also includes its directory
MAIN_PATH = Left(file_left, InStrRev(file_left, "\"))
              
' Using substring of the title (file_left) to correctly locate the file from its directory
Dim file As String
Dim oldFile As String

' If font color is Red, then skip it. (Red indicates that the file was not found in Master file
If objExcel.Cells(intRow, 1).Font.Color = RGB(255, 0, 0) Then
    logFile.Write "(" & (intRow - 1) & ") " & send_to & vbCrLf & _
                  "     Filename: " & file_left & vbCrLf & _
                  "     Error: This item will be skipped as it's determined Invalid... (Red Font)" & vbCrLf & _
                  "     Error: This is most likely due to not being found in the MASTER Excel file." & vbCrLf

    'MsgBox "Skipping this one..."
    GoTo NextIteration
Else
    'MsgBox "Good stuff, processing..."
End If


' Check if a file with the new name already exists; Due to duplicate files that might exist
If Len(Dir(new_fileName)) > 0 Then
    logFile.Write "(" & (intRow - 1) & ") " & send_to & vbCrLf & _
                  "     New Filename: " & new_fileName & vbCrLf & _
                  "         ---> Already EXISTS in the directory! " & vbCrL

    GoTo NextIteration
End If

' Check for the existing/old file to prepare for file-read and date-update.
If file_left <> "" Then
    file = Dir$(file_left & "*" & ".*")
    If (Len(file) > 0) Then
        oldFile = file ' 'file' includes only the file name excluding its Directory
    Else
        logFile.Write "(" & (intRow - 1) & ") " & send_to & vbCrLf & _
                      "     Filename (Prefix): " & file_left & vbCrLf & _
                      "     Error: The Word file wasn't found in the directory!" & vbCrLf
        GoTo NextIteration
    End If
Else
    ' Do nothing since no file exists
    logFile.Write "(" & (intRow - 1) & ") " & send_to & vbCrLf & _
                  "     Filename (Prefix): " & file_left & vbCrLf & _
                  "     ---> This record doesn't require a file" & vbCrLf
    
    GoTo NextIteration
End If


' Note: Below creates a new file and move the old file to a subdirectory named 'OldFiles'
' Open the Word document
Dim wdApp As Object, wdDoc As Object
Set wdApp = CreateObject("Word.Application")
Set wdDoc = wdApp.Documents.Open(MAIN_PATH + oldFile) '("C:\Users\ekim\Desktop\Projects\hello\testbaby.docx")
wdApp.Visible = True

With wdApp.ActiveDocument.Content.Find ' Note: Saw some people using Content instead of Range
    .ClearFormatting ' Clear any existing value for Find dialog box
    .Replacement.ClearFormatting ' Clear any existing value
    '.Text = "([0-9]{4})[-]([0-9]{1,2})[-]([0-9]{1,2})"
    .Text = "testbaby"
    .MatchWildcards = True
    .Replacement.Text = "testbaby" ' new_fileDate ' Change .Text to what's provided here
    '.Execute Replace:=2 ' Apply for Replace All
End With
 

'1) Store the full path of the document in a string: oldfile = ActiveDocument.FullName
    oldFullFile = MAIN_PATH + oldFile '
    
    'MsgBox "Old Name: " & oldFile & vbCrLf & _
           "Full File: " & oldFullFile & vbCrLf & _
           "New Name: " & new_fileName

'2) Create a new file with the new Filename using SaveAs
    wdDoc.SaveAs (new_fileName)

    ' Save changes and close MS Words
    wdApp.ActiveDocument.Close ' SaveChanges:=True ' Since we are making a copy, i don't think we should save the active one
    wdApp.Quit

'3) Move old file to OldFiles direcotry (instead of copy+move+delete_original)
    
    'Check if OldFiles directory exists where we can move to! If not, create it for old files storage
    Call makeDirectory(MAIN_PATH)
    
    Dim OLD_FILES_PATH As String
    OLD_FILES_PATH = MAIN_PATH + "OldFiles\"
    
    ' Move vs Copy And Delete
    '!!! There might be an existing file in the OLD_FILES_PATH so we need a filter
    Dim send_to0 As String
    send_to0 = Trim(objExcel.Cells(intRow, 1).Value)
        
    If Dir(OLD_FILES_PATH + oldFile) = "" Then
        ' Old File doesn't exist in OldFiles directory so just move it
        Set fso = CreateObject("scripting.filesystemobject")
        fso.MoveFile Source:=oldFullFile, Destination:=OLD_FILES_PATH
        
        logFile.Write "(" & (intRow - 1) & ") " & send_to & vbCrLf & _
                      "     New Filename: " & new_fileName & vbCrLf & _
                      "     Old Filename: " & oldFile & vbCrLf & _
                      "         ---> Moved to the following: " & OLD_FILES_PATH & vbCrLf & _
                      "     Note: Old File was Successfully moved to the subfolder OldFiles" & vbCrLf

    Else
        ' Old File already exist (mainly due to same client for different report or etc.)
        logFile.Write "(" & (intRow - 1) & ") " & send_to & vbCrLf & _
                      "     New Filename: " & new_fileName & vbCrLf & _
                      "     Old Filename: " & oldFile & vbCrLf & _
                      "         ---> Already EXISTS in the oldFile folder." & vbCrLf & _
                      "     Note: Old File is already archived in OldFiles folder." & vbCrLf
    End If
    
    intSuccess = intSuccess + 1
       
NextIteration:

' Go to next loop!
logFile.Write vbCrLf & vbCrLf
intRow = intRow + 1
Loop

logFile.Write "---------------------------------------------------------" & vbCrLf
logFile.Write "Total Items Checked: " & (intRow - 2) & vbCrLf & _
              "New Files Created: " & intSuccess & vbCrLf

objWorkbook.Close
objExcel.Quit

Application.ScreenUpdating = True

MsgBox "Total Items Checked: " & (intRow - 2) & vbCrLf & _
              "New Files Created: " & intSuccess & vbCrLf

End Sub


Function makeDirectory(strDir As String)
    Set fso = CreateObject("scripting.filesystemobject")
    Dim path As String
    path = strDir + "OldFiles\"
    
    If Not fso.FolderExists(path) Then
        fso.CreateFolder path
    End If
    
End Function

