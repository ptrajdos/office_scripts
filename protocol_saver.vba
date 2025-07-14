Sub ConvertPDFsToDocx()
    Dim folderPath As String
    Dim shellCmd As String
    Dim fileList As String
    Dim files() As String
    Dim i As Long
    Dim pdfFile As String
    Dim pdfFullPath As String
    Dim docxFullPath As String
    Dim doc As Document

    ' Select the folder
    folderPath = MacSelectFolder()
    If folderPath = "" Then Exit Sub

    ' Ensure folder path ends with slash
    If Right(folderPath, 1) <> "/" Then folderPath = folderPath & "/"

    ' Use simple ls command to get all files in folder
    shellCmd = "ls " & folderPath

    On Error Resume Next
    fileList = MacScript("do shell script """ & shellCmd & """")
    On Error GoTo 0

    If Trim(fileList) = "" Then
        MsgBox "No files found in folder.", vbExclamation
        Exit Sub
    End If

    fileList = Replace(fileList, vbCrLf, vbLf)
    fileList = Replace(fileList, vbCr, vbLf)
    files = Split(fileList, vbLf)

    For i = LBound(files) To UBound(files)
        pdfFile = files(i)
        ' Check if file ends with .pdf or .PDF (case insensitive)
        If LCase(Right(pdfFile, 4)) = ".pdf" Then
            pdfFullPath = folderPath & pdfFile
            docxFullPath = folderPath & Left(pdfFile, Len(pdfFile) - 4) & ".docx"

            On Error Resume Next
            Set doc = Documents.Open(fileName:=pdfFullPath, ConfirmConversions:=False, ReadOnly:=True)
            If Err.Number <> 0 Or doc Is Nothing Then
                Debug.Print "Failed to open: " & pdfFullPath
                Err.Clear
                GoTo NextFile
            End If
            On Error GoTo 0

            ' Save as DOCX
            doc.SaveAs2 fileName:=docxFullPath, FileFormat:=wdFormatXMLDocument
            doc.Close SaveChanges:=False
        End If
NextFile:
    Next i

    MsgBox "PDF to DOCX conversion complete."
End Sub

Function MacSelectFolder() As String
    Dim script As String
    script = "set theFolder to POSIX path of (choose folder with prompt ""Select folder:"")"
    On Error Resume Next
    MacSelectFolder = MacScript(script)
    On Error GoTo 0
End Function

