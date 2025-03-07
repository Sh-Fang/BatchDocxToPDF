Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objWord = CreateObject("Word.Application")
objWord.Visible = False

' Ask for user confirmation
response = MsgBox("Do you want to start the conversion?", vbYesNo + vbQuestion, "Confirm Conversion")
If response <> vbYes Then
    WScript.Quit
End If

' Get the current directory
currentFolder = objFSO.GetAbsolutePathName(".")
outputFolder = currentFolder & "\output"

' Ensure the output directory exists
If Not objFSO.FolderExists(outputFolder) Then
    objFSO.CreateFolder(outputFolder)
End If

' Count total Word files
Set objFolder = objFSO.GetFolder(currentFolder)
totalFiles = 0
For Each objFile In objFolder.Files
    ext = LCase(objFSO.GetExtensionName(objFile.Name))
    If ext = "docx" Or ext = "doc" Then
        totalFiles = totalFiles + 1
    End If
Next

' Convert files with progress
convertedFiles = 0
Set objFolder = objFSO.GetFolder(currentFolder)
For Each objFile In objFolder.Files
    ext = LCase(objFSO.GetExtensionName(objFile.Name))
    If ext = "docx" Or ext = "doc" Then
        docPath = objFile.Path
        pdfPath = outputFolder & "\" & objFSO.GetBaseName(objFile.Name) & ".pdf"
        
        ' Open the Word document and save it as PDF
        Set objDoc = objWord.Documents.Open(docPath, False, True)
        objDoc.SaveAs2 pdfPath, 17 ' 17 means PDF format
        objDoc.Close False

    End If
Next

' Quit Word application
objWord.Quit

WScript.Echo "Conversion completed! PDF files are saved in: " & outputFolder
