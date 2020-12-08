Attribute VB_Name = "AOC1"
Sub Day01_a()

    Worksheets("AoC 1").Activate
    d = Split(Range("D4").Value, vbLf)

    For Each c1 In d
    
        For Each c2 In d
            
            If CInt(c1) + CInt(c2) = 2020 Then
            
                Range("E6") = CInt(c1)
                Range("F6") = CInt(c2)
                
                Range("I6") = "=E6*F6"
            
            End If
    
        Next c2
    
    Next c1

End Sub

Sub Day01_b()

    Worksheets("AoC 1").Activate
    d = Split(Range("D4").Value, vbLf)

    For Each c1 In d
        For Each c2 In d
            For Each c3 In d
                If CInt(c1) + CInt(c2) + CInt(c3) = 2020 Then
            
                    Range("E8") = CInt(c1)
                    Range("F8") = CInt(c2)
                    Range("G8") = CInt(c3)
                    
                    Range("I8") = "=E8*F8*G8"
                
                End If
            Next c3
        Next c2
    Next c1
    
End Sub
