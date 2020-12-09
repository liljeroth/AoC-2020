Attribute VB_Name = "AoC02"
Sub Day02A()
    
    ' Load input for puzzle
    d = Split(ReadFile("AoC02.txt"), vbNewLine)
    
    res = 0
    For Each c In d
        
        pol = Split(c, ":")(0)
        pas = Trim(Split(c, ":")(1))
        
        Min = CInt(Split(pol, "-")(0))
        Max = CInt(Split(Split(pol, "-")(1), " ")(0))
        Cha = Split(pol, " ")(1)
        
        Count = Len(pas) - Len(Replace(pas, Cha, ""))
        
        If Count >= Min And Count <= Max Then
            
            res = res + 1
        
        End If
    
    Next c
    
    ' Answer: 640
    Range("D02A") = res
    
End Sub

Sub Day02B()
    
    ' Load input for puzzle
    d = Split(ReadFile("AoC02.txt"), vbNewLine)

    res = 0
    For Each c In d
        
        pol = Split(c, ":")(0)
        pas = Trim(Split(c, ":")(1))
        
        Min = CInt(Split(pol, "-")(0))
        Max = CInt(Split(Split(pol, "-")(1), " ")(0))
        Cha = Split(pol, " ")(1)
        
        If Mid(pas, Min, 1) = Cha Xor Mid(pas, Max, 1) = Cha Then
            
            res = res + 1
        
        End If
    
    Next c
    
    ' Answer: 472
    Range("D02B") = res

End Sub

