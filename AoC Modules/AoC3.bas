Attribute VB_Name = "AoC3"
Sub Day03_a()

    d = Split(Worksheets("AoC 3").Range("D4").Value, vbLf)
    Worksheets("AoC 3").Range("I6") = 0
    
    pos = 1
    res = 0
    
    init = 1
    For Each c In d
        
        If init = 1 Then
            init = 0
        Else
            pos = pos + 3
            If pos > Len(c) Then
                pos = pos - Len(c)
            End If
            
            If Mid(c, pos, 1) = "#" Then
                res = res + 1
            End If
        End If
    Next c
    
    Worksheets("AoC 3").Range("I6") = res
    
End Sub

Sub Day03_b()

    d = Split(Worksheets("AoC 3").Range("D4").Value, vbLf)
    Worksheets("AoC 3").Range("I8") = 0
    
    Worksheets("AoC 3").Range("D8") = 0
    Worksheets("AoC 3").Range("E8") = 0
    Worksheets("AoC 3").Range("F8") = 0
    Worksheets("AoC 3").Range("G8") = 0
    Worksheets("AoC 3").Range("H8") = 0
    
    Dim r(0 To 4) As Integer
    r(0) = 1
    r(1) = 3
    r(2) = 5
    r(3) = 7
    
    For Each p In r
        
        pos = 1
        res = 0
        
        init = 1
        For Each c In d
            
            If init = 1 Then
            
                init = 0
                
            Else
            
                pos = pos + p
                
                If pos > Len(c) Then
                    pos = pos - Len(c)
                End If
                
                If Mid(c, pos, 1) = "#" Then
                    res = res + 1
                End If
                
            End If
            
        Next c
        
        If Worksheets("AoC 3").Range("D8") = 0 Then
            Worksheets("AoC 3").Range("D8") = res
        ElseIf Worksheets("AoC 3").Range("E8") = 0 Then
            Worksheets("AoC 3").Range("E8") = res
        ElseIf Worksheets("AoC 3").Range("F8") = 0 Then
            Worksheets("AoC 3").Range("F8") = res
        ElseIf Worksheets("AoC 3").Range("G8") = 0 Then
            Worksheets("AoC 3").Range("G8") = res
        End If
        
    Next p
    
    pos = 1
    res = 0
    
    init = 2
    For Each c In d
        
        If init >= 1 Then
            
            init = init - 1
            
        Else
            
            init = 1
            pos = pos + 1
            
            If pos > Len(c) Then
                pos = pos - Len(c)
            End If
            
            If Mid(c, pos, 1) = "#" Then
                res = res + 1
            End If
            
        End If
        
    Next c
    
    Worksheets("AoC 3").Range("H8") = res
    
    Worksheets("AoC 3").Range("I8") = "=D8*E8*F8*G8*H8"

End Sub
