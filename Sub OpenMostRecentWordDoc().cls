Sub OpenMostRecentWordDoc()
    Dim folderPath As String
    Dim fileSystem As Object
    Dim folder As Object
    Dim file As Object
    Dim latestFile As String
    Dim latestDate As Date
    Dim ws As Worksheet
    
    ' Reference the Parameters sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Parameters")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Sheet 'Parameters' not found. Please create it and specify the folder path in cell D11.", vbExclamation, "Missing Sheet"
        Exit Sub
    End If
    
    ' Retrieve folder path from cell D11 in the Parameters sheet
    folderPath = Trim(ws.Range("D11").Value)
    
    ' Validate folder path
    If folderPath = "" Then
        MsgBox "Folder path is empty. Please enter a valid path in cell D11 of the 'Parameters' sheet.", vbExclamation, "Invalid Path"
        Exit Sub
    End If
    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If
    
    ' Create a FileSystemObject
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FolderExists(folderPath) Then
        MsgBox "The specified folder path does not exist: " & folderPath, vbExclamation, "Folder Not Found"
        Exit Sub
    End If
    
    Set folder = fileSystem.GetFolder(folderPath)
    
    ' Initialize variables
    latestDate = 0
    latestFile = ""
    
    ' Loop through files in the folder
    For Each file In folder.Files
        If UCase(fileSystem.GetExtensionName(file)) = "DOCX" Or UCase(fileSystem.GetExtensionName(file)) = "DOC" Then
            If file.DateLastModified > latestDate Then
                latestDate = file.DateLastModified
                latestFile = file.Path
            End If
        End If
    Next file
    
    ' Open the most recent Word document
    If latestFile <> "" Then
        Shell "explorer.exe """ & latestFile & """", vbNormalFocus
    Else
        MsgBox "No Word documents found in the specified folder.", vbExclamation, "No Files"
    End If
End Sub

