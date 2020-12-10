Attribute VB_Name = "AoC10"
Sub Day10A()

    ' Load input for puzzle
    d = Split(ReadFile("AoC10.txt"), vbNewLine)
    
    ' Initiate variables
    res = 0
    
    '
    curr = 1000
    
    For i = 0 To UBound(d)
    
        If CLng(d(i)) < curr Then curr = CLng(d(i))
            
    Next i
    
    used = "," & curr & ","
    
    poss = 0
    found1 = False
    found3 = False
    
    res1 = 1
    res3 = 1
    i = 0
    Do While i < UBound(d) Or found1 Or found3
        For i = 0 To UBound(d)
        
            If Not InStr(used, "," & i & ",") Then
        
                If CLng(d(i)) = curr + 1 Then
                
                    found1 = True
                    curr = CLng(d(i))
                    used = used & "," & i & ","
                    res1 = res1 + 1
                    i = 0
                    found1 = False
                    found3 = False
                    
                    Exit For
                    
                ElseIf CLng(d(i)) = curr + 3 Then
                
                    found3 = True
                    poss = i
                    
                End If
                
            End If
                
        Next i
        
        If found3 And Not found1 Then
            curr = CLng(d(poss))
            used = used & "," & poss & ","
            res3 = res3 + 1
            i = 0
            found1 = False
            found3 = False
        End If
        
    Loop
    
    'Answer:
    Range("D10A") = res1 * res3

End Sub


Sub Day10B()
    
    ' Load input for puzzle
    d = Split(ReadFile("AoC10.txt"), vbNewLine)
    
    ' Initiate variables
    res = 0
    curr = 1000
    
    For i = 0 To UBound(d)
    
        If CLng(d(i)) < curr Then
            curr = CLng(d(i))
            past = i
        End If
    
        If CLng(d(i)) > maxval Then
            maxval = CLng(d(i))
        End If
            
    Next i
    
    Sort = curr & vbNewLine
    For i = 1 To UBound(d)
        n = maxval
        For j = 0 To UBound(d)
            If CLng(d(j)) > curr And CLng(d(j)) < n Then
                n = CLng(d(j))
            End If
        Next j
        curr = n
        Sort = Sort & n & vbNewLine
    Next i
    Sort = Split(Left(Sort, Len(Sort) - 1), vbNewLine)
    
    Range("D10B") = getTree2(Sort)

End Sub

Function getTree2(d)
    
    a = Array(0, 0, 0, 0, 0)
    l = 0
    r = 0
    
    For Each c In d
        
        If CLng(c) = CLng(l) + 1 Then
        
            r = r + 1
            
            If r = 4 Then
                a(r) = a(r) + 1
                r = 0
            End If
        
        Else
            a(r) = a(r) + 1
            r = 0
        
        End If
        
        l = c
        
    Next c
    
    getTree2 = (2 ^ a(2)) * (4 ^ a(3)) * (7 ^ a(4))
    
    ' 12089663946752

End Function
