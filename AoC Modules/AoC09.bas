Attribute VB_Name = "AoC09"
Sub Day09A()
    
    ' Load input for puzzle
    d = Split(ReadFile("AoC09.txt"), vbNewLine)
    
    ' Initiate variables
    preamble = 25
    i = preamble
    
    '
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
    
    'Answer: 393911906
    Range("D09A") = d(i)

End Sub

Sub Day09B()

    ' Initiate routine
    'Dim minval As Long
    'Dim maxval As Long
    
    ' Load input for puzzle
    d = Split(ReadFile("AoC09.txt"), vbNewLine)
    
    ' Initiate variables
    irow = 632
        
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
    
    'Answer: 59341885
    Range("D09B") = minval + maxval
    
End Sub

