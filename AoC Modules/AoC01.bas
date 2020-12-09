Attribute VB_Name = "AoC01"
Sub Day01A()
    
    ' Load input for puzzle
    d = Split(ReadFile("AoC01.txt"), vbNewLine)

    For i = 1 To UBound(d)
        
        For j = i + 1 To UBound(d)
            
            If CInt(d(i)) + CInt(d(j)) = 2020 Then
                
                Range("D01A") = CLng(d(i)) * CLng(d(j))
            
            End If
    
        Next j
    
    Next i

    ' Answer: 482811
    
End Sub

Sub Day01B()

    ' Load input for puzzle
    d = Split(ReadFile("AoC01.txt"), vbNewLine)

    For i = 1 To UBound(d)
        
        For j = i + 1 To UBound(d)
            
            For k = j + 1 To UBound(d)
            
                If CInt(d(i)) + CInt(d(j)) + CInt(d(k)) = 2020 Then
                    
                    Range("D01B") = CLng(d(i)) * CLng(d(j)) * CLng(d(k))
                
                End If
                
            Next k
            
        Next j
        
    Next i

    ' Answer: 193171814
    
End Sub
