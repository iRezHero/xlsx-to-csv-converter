Dim oExcel, newFolderPath, newFolder, logFile, alreadyConvertedNum
'Sets the excel object
Set oExcel = CreateObject("Excel.Application") 

Set objFSO = CreateObject("Scripting.FileSystemObject")

'Sets the variable to zero and will be incremented if there are already converted files
alreadyConvertedNum = 0

'If there is an argument passed by command line we use it as working directory
If WScript.Arguments.Count > 0 Then
  CurrentDirectory = objFSO.GetAbsolutePathName(WScript.Arguments.Item(0))
else
'Otherwise we will be targeting on current directory
  CurrentDirectory = objFSO.GetAbsolutePathName(".")
End If

'Declaring the path of the logfile
logFile = CurrentDirectory & "\log.txt"

'It opens up the log to be ready for writing into it
Set myLog = objFSO.OpenTextFile(logFile, 8, True)

'Verifies in both cases that folder exists
If Not objFSO.FolderExists(CurrentDirectory) Then
  myLog.WriteLine "#[" & Now & "] - [ERROR] - The directory you specified does not exist or cannot be reached."
  WScript.Quit
End If

'Declares the new folder to be created if doesn't exist in the directory
newFolderPath = CurrentDirectory & "\XLSX_OLD Exports"

'Even in this case it checks if the new folder exists, if not it creates that
If Not objFSO.FolderExists(newFolderPath) Then
  Set newFolder = objFSO.CreateFolder(newFolderPath) 
End If

'If the log doesn't exist it creates a new one
If Not objFSO.FileExists(logFile) Then
 objFSO.CreateTextFile(logFile)
End if

'Cycle for all the files with 'xlsx' extension
For Each oFile In objFSO.GetFolder(CurrentDirectory).Files
  'Check if into the folder we have got XLSX files to be converted
  'If an already converted file exists we jump to the next
  If UCase(objFSO.GetExtensionName(oFile.Name)) = "XLSX" Then 
    If Not objFSO.FileExists(Replace(oFile.Name, "xlsx", "csv")) Then
      'Calls the ProcessFiles method passing the actual cycling item
      ProcessFiles objFSO, oFile
      'If success it writes down a new line into the log
        myLog.WriteLine "#[" & Now & "] - [SUCCESS] " & oFile.Name & " successfully converted;"
      If Not objFSO.FileExists(newFolderPath & "\" & oFile.Name) Then
        'Moves the freshly processed file into '\XLSX_OLD Exports'
        oFile.Move newFolderPath & "\"
      Else
        'Throws the message box error if some files have already been converted
        myLog.WriteLine "#[" & Now & "] - [WARNING] - " & oFile.Name & " already exists in " & newFolderPath & ". Not being moved;"
      End If
    Else
      alreadyConvertedNum = alreadyConvertedNum + 1
      myLog.WriteLine "#[" & Now & "] - [WARNING] - " & oFile.Name & " already converted!"
    End If
  End If
Next

If alreadyConvertedNum <> 0 Then
  myLog.WriteLine "#[" & Now & "] - [WARNING] - " & alreadyConvertedNum & " files have been already converted."
  WScript.Quit
End If

'This method converts into .csv file and separates each cell by ';'
Private Sub ProcessFiles(objFSO, oFile)
  'Sets up the destination file and replaces its extension
  dest_file = Replace(Replace(oFile,".xlsx",".csv"),".xls",".csv")

  oExcel.DisplayAlerts = FALSE 'To avoid prompts
  Dim oBook, local
  'Opens up the file by excel
  Set oBook = oExcel.Workbooks.Open(oFile)
  local = true 
  'And then it saves into the 'CurrentDirectory'
  call oBook.SaveAs(dest_file, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, local) 
  oExcel.Quit 
End Sub