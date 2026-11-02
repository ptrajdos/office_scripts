Sub MakeDocsSummary()
    Dim folderPath As String
    Dim shellCmd As String
    Dim fileList As String
    Dim files() As String
    Dim i As Long, j As Long
    Dim doc As Document
    Dim docPath As String


    ' Select folder
    folderPath = MacSelectFolder()
    If folderPath = "" Then Exit Sub
    If Right(folderPath, 1) <> "/" Then folderPath = folderPath & "/"

    ' Get all .docx files in folder
    shellCmd = "ls " & folderPath & "*.docx"

    On Error Resume Next
    fileList = MacScript("do shell script """ & shellCmd & """")
    On Error GoTo 0

    If Trim(fileList) = "" Then
        MsgBox "No .docx files found in folder.", vbExclamation
        Exit Sub
    End If

    ' Normalize line breaks and split
    fileList = Replace(fileList, vbCrLf, vbLf)
    fileList = Replace(fileList, vbCr, vbLf)
    files = Split(fileList, vbLf)

    Application.ScreenUpdating = False

    For i = LBound(files) To UBound(files)
        docPath = files(i)

        On Error Resume Next
        Set doc = Documents.Open(fileName:=docPath, ReadOnly:=False)
        If Err.Number <> 0 Or doc Is Nothing Then
            Debug.Print "Failed to open: " & docPath
            Err.Clear
            GoTo NextFile
        End If
        On Error GoTo 0
        Debug.Print "Processing: " & docPath
        
    
        ' doc.Save
        ' doc.Close SaveChanges:=False
NextFile:
    Next i

    Application.ScreenUpdating = True
    MsgBox "Done striking through phrases in documents."
End Sub
