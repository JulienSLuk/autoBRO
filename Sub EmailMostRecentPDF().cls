Sub EmailMostRecentPDF()
    Dim folderPath As String
    Dim fileSystem As Object
    Dim folder As Object
    Dim file As Object
    Dim latestFile As String
    Dim latestDate As Date
    Dim ws As Worksheet
    Dim emailRecipients As String
    Dim OutApp As Object
    Dim OutMail As Object
    Dim emailArray() As String
    Dim i As Integer
    Dim emailValid As Boolean
    
    ' Reference the Parameters sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Parameters")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Sheet 'Parameters' not found. Please create it and specify the folder path in cell D11 and email recipients in cell D14.", vbExclamation, "Missing Sheet"
        Exit Sub
    End If
    
    ' Retrieve folder path from cell D11
    folderPath = Trim(ws.Range("D11").Value)
    
    ' Retrieve email recipients from cell D14
    emailRecipients = Trim(ws.Range("D14").Value)
    
    ' Validate folder path
    If folderPath = "" Then
        MsgBox "Folder path is empty. Please enter a valid path in cell D11 of the 'Parameters' sheet.", vbExclamation, "Invalid Path"
        Exit Sub
    End If
    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If
    
    ' Validate email format
    emailValid = True
    If emailRecipients = "" Then
        MsgBox "Email address field (D14) is empty. Please enter a valid email address.", vbExclamation, "Invalid Email"
        Exit Sub
    End If
    
    ' Split email recipients by semicolon and check each one
    emailArray = Split(emailRecipients, ";")
    
    For i = LBound(emailArray) To UBound(emailArray)
        If Not IsValidEmail(Trim(emailArray(i))) Then
            emailValid = False
            MsgBox "Invalid email address: " & emailArray(i), vbExclamation, "Invalid Email Format"
            Exit Sub
        End If
    Next i
    
    If Not emailValid Then Exit Sub
    
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
        If UCase(fileSystem.GetExtensionName(file)) = "PDF" Then
            If file.DateLastModified > latestDate Then
                latestDate = file.DateLastModified
                latestFile = file.Path
            End If
        End If
    Next file
    
    ' Check if a PDF was found
    If latestFile = "" Then
        MsgBox "No PDF files found in the specified folder.", vbExclamation, "No Files"
        Exit Sub
    End If
    
    ' Create Outlook email
    On Error Resume Next
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    On Error GoTo 0
    
    If OutApp Is Nothing Then
        MsgBox "Outlook is not installed or configured on this system.", vbExclamation, "Outlook Not Found"
        Exit Sub
    End If
    
    ' Configure email
    With OutMail
        .To = emailRecipients
        .Subject = "Latest PDF File"
        .Body = "Please find the latest PDF file attached."
        .Attachments.Add latestFile
        .Display ' Use .Send to send the email directly
    End With
    
    ' Clean up
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub

' Function to validate email format
Function IsValidEmail(email As String) As Boolean
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Define the regex pattern for a valid email address
    regex.IgnoreCase = True
    regex.Global = True
    regex.Pattern = "^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$"
    
    IsValidEmail = regex.Test(email)
End Function
