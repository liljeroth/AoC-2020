Attribute VB_Name = "AoC2"
Sub Day02_a()

    d = Split(Worksheets("AoC 2").Range("D4").Value, vbLf)
    Worksheets("AoC 2").Range("I6") = 0

    For Each c In d
        
        pol = Split(c, ":")(0)
        pas = Trim(Split(c, ":")(1))
        
        Min = CInt(Split(pol, "-")(0))
        Max = CInt(Split(Split(pol, "-")(1), " ")(0))
        Cha = Split(pol, " ")(1)
        
        Count = Len(pas) - Len(Replace(pas, Cha, ""))
        
        If Count >= Min And Count <= Max Then
            
            Worksheets("AoC 2").Range("I6") = Worksheets("AoC 2").Range("I6") + 1
        
        End If
    
    Next c

End Sub

Sub Day02_b()

    d = Split(Worksheets("AoC 2").Range("D4").Value, vbLf)
    Worksheets("AoC 2").Range("I8") = 0

    For Each c In d
        
        pol = Split(c, ":")(0)
        pas = Trim(Split(c, ":")(1))
        
        Min = CInt(Split(pol, "-")(0))
        Max = CInt(Split(Split(pol, "-")(1), " ")(0))
        Cha = Split(pol, " ")(1)
        
        If Mid(pas, Min, 1) = Cha And Mid(pas, Max, 1) = Cha Then
        
            ' Do nothing
        
        ElseIf Mid(pas, Min, 1) = Cha Or Mid(pas, Max, 1) = Cha Then
            
            Worksheets("AoC 2").Range("I8") = Worksheets("AoC 2").Range("I8") + 1
        
        Else
        
            ' Do nothing
        
        End If
    
    Next c

End Sub

