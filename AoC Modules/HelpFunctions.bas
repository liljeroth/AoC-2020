Attribute VB_Name = "HelpFunctions"
Function ReadFile(file)

    Dim text As String
    Open ActiveWorkbook.Path & "\Input\" & file For Input As #1
    Do Until EOF(1)
        Line Input #1, textline
        text = text & Replace(Replace(textline, vbCr, ""), vbLf, "") & vbNewLine
    Loop
    Close #1
    
    ReadFile = Left(text, Len(text) - 1)

End Function
