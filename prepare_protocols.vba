Sub StrikeThroughPhrasesInDocs()
    Dim folderPath As String
    Dim shellCmd As String
    Dim fileList As String
    Dim files() As String
    Dim i As Long, j As Long
    Dim doc As Document
    Dim docPath As String

    ' Define your phrases here (case insensitive search)
    Dim phrasesToStrike As Variant
    phrasesToStrike = Array("większością głosów", "nie uzyskał tytułu zawodowego", "nie uzyskała tytułu zawodowego")

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
        
        Call MoveLastLineUp(doc)
        Call DeleteLastPage(doc)

        ' Loop through all phrases and strike through each occurrence
        For j = LBound(phrasesToStrike) To UBound(phrasesToStrike)
            With doc.Content.Find
                .ClearFormatting
                .Text = phrasesToStrike(j)
                .Replacement.ClearFormatting
                .Replacement.Font.StrikeThrough = True
                .Forward = True
                .Wrap = wdFindContinue
                .Format = True
                .MatchCase = False
                .MatchWholeWord = False
                .Execute Replace:=wdReplaceAll
            End With
        Next j

        doc.Save
        doc.Close SaveChanges:=False
NextFile:
    Next i

    Application.ScreenUpdating = True
    MsgBox "Done striking through phrases in documents."
End Sub

Function MacSelectFolder() As String
    Dim script As String
    script = "set theFolder to POSIX path of (choose folder with prompt ""Select folder:"")"
    On Error Resume Next
    MacSelectFolder = MacScript(script)
    On Error GoTo 0
End Function

Sub MoveLastLineUp(doc As Document)
Dim paras As Paragraphs
    Dim lastText As String, secondLastText As String
    Dim lastPara As Range, secondLastPara As Range
    Dim paraCount As Long

    Set paras = doc.Paragraphs
    paraCount = paras.Count

    If paraCount < 2 Then Exit Sub

    ' Find last non-empty paragraph
    Dim i As Long
    For i = paraCount To 1 Step -1
        If Trim(paras(i).Range.Text) <> "" Then
            Set lastPara = paras(i).Range
            Exit For
        End If
    Next i

    If lastPara Is Nothing Or i = 1 Then Exit Sub

    ' Find second-last non-empty paragraph
    Dim j As Long
    For j = i - 1 To 1 Step -1
        If Trim(paras(j).Range.Text) <> "" Then
            Set secondLastPara = paras(j).Range
            Exit For
        End If
    Next j

    If secondLastPara Is Nothing Then Exit Sub

    ' Combine second-last and last paragraph with Shift+Enter (manual line break)
    secondLastPara.End = secondLastPara.End - 1 ' remove paragraph mark
    ' No paragraph formatting is saved. 
    secondLastPara.InsertAfter Chr(11) & lastPara
    lastPara.Delete
    End Sub


Sub DeleteLastPage(doc As Document)
    Dim totalPages As Long
    Dim startRange As Range
    Dim endRange As Range
    Dim lastPageRange As Range

    totalPages = doc.ComputeStatistics(wdStatisticPages)
    If totalPages < 2 Then Exit Sub

    ' Start of last page
    Set startRange = doc.GoTo(What:=wdGoToPage, Name:=totalPages)
    ' End of document
    Set endRange = doc.Range(doc.Content.End - 1, doc.Content.End)

    ' Delete last page
    Set lastPageRange = doc.Range(Start:=startRange.Start, End:=endRange.End)
    lastPageRange.Delete
End Sub
