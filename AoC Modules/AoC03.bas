Attribute VB_Name = "AoC03"
Sub Day03A()
    
    ' Load input for puzzle
    d = Split(ReadFile("AoC03.txt"), vbNewLine)
    
    ' Initiate variables
    pos = 1
    res = 0
    
    For i = 1 To UBound(d)
        
        pos = pos + 3
        
        If pos > Len(d(i)) Then pos = pos - Len(d(i))
        
        If Mid(d(i), pos, 1) = "#" Then res = res + 1
        
    Next i
    
    ' Answer: 203
    Range("D03A") = res
    
End Sub

Sub Day03B()
    
    ' Load input for puzzle
    d = Split(ReadFile("AoC03.txt"), vbNewLine)
    
    ' Initiate variables
    r = Array(0, 0, 0, 0, 0)
    sX = Array(1, 3, 5, 7, 1)
    sY = Array(1, 1, 1, 1, 2)
    
    For i = 0 To UBound(sX)
        
        pos = 1
        res = 0
        
        For j = 1 To UBound(d)
        
            j = j + sY(i) - 1
            pos = pos + sX(i)
            
            If pos > Len(d(j)) Then pos = pos - Len(d(j))
            
            If Mid(d(j), pos, 1) = "#" Then r(i) = r(i) + 1
            
        Next j
        
    Next i
    
    res = r(0)
    For i = 1 To UBound(r)
        res = res * r(i)
    Next i
    
    ' Answer: 3316272960
    Range("D03B") = res

End Sub
