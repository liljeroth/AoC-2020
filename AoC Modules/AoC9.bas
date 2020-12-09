Attribute VB_Name = "AoC9"
Sub Day09_a()

    ' Initiate routine
    Worksheets("AoC 9").Activate
    
    Range("I6") = 0 ' Clean the result cell, just in case...
    
    ' Read input data
    d = Split(ReadFile("AoC9Data.txt"), vbNewLine)
    preamble = 25
    
    i = preamble
    ' Loop through to find all iBag containing a "shiny gold bag"
    Do While i < UBound(d)
        
        sumFound = False
        For j = i - preamble To i - 1
        
            For k = i - preamble To i - 1
                
                If CLng(d(i)) = CLng(d(j)) + CLng(d(k)) Then
                    
                    sumFound = True
                    Exit For
                    
                End If
            
            Next k
            
        If sumFound Then Exit For
        
        Next j
        
        If Not sumFound Then Exit Do
        
        i = i + 1
    Loop
    
    MsgBox d(i) & ", i=" & i
    Range("I6") = d(i)
    
    '393911906

End Sub

Sub Day09_b()

    ' Initiate routine
    Dim minval As Long
    Dim maxval As Long
    
    Worksheets("AoC 9").Activate
    
    Range("I8") = 0 ' Clean the result cell, just in case...
    
    ' Read input data
    d = Split(ReadFile("AoC9Data.txt"), vbNewLine)
    
    'irow = 14 ' Example
    irow = 632 ' Puzzle
        
    For j = 0 To irow - 1
        
        vSum = CLng(d(j))
        
        minval = CLng(d(j))
        maxval = CLng(d(j))
            
        For k = j + 1 To irow - 1
            
            vSum = vSum + CLng(d(k))
            
            If CLng(d(k)) < minval Then minval = CLng(d(k))
            If CLng(d(k)) > maxval Then maxval = CLng(d(k))
            
            If CLng(d(irow)) = vSum Then
            
                sumFound = True
                Exit For
                
            End If
        
        Next k
        
        If sumFound Then Exit For
    
    Next j
    
    MsgBox minval & " + " & maxval & " = " & minval + maxval
    Range("I8") = minval + maxval
    
    '59341885
    
End Sub

