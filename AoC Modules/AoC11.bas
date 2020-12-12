Attribute VB_Name = "AoC11"
Sub Day11A()

    ' Load input for puzzle
    d = Split(ReadFile("AoC11.txt"), vbNewLine)
    nd = d
    
    ' Initiate variables
    res = 0
    
    '
    changing = True
    Do While changing
        changing = False
        For i = 0 To UBound(d)
            
            For j = 1 To Len(d(i))
                
                n = 0
                If Mid(d(i), j, 1) = "L" Or Mid(d(i), j, 1) = "#" Then
                    n = getN(d, i, j)
                End If
                
                If Mid(d(i), j, 1) = "L" And n = 0 Then
                    
                    changing = True
                    c = Left(nd(i), j - 1) & "#" & Mid(d(i), j + 1)
                    nd(i) = c
                    
                End If
                
                If Mid(d(i), j, 1) = "#" And n >= 4 Then
                    
                    changing = True
                    c = Left(nd(i), j - 1) & "L" & Mid(d(i), j + 1)
                    nd(i) = c
                    
                End If
                    
            Next j
            
        Next i
        d = nd
    Loop
    
    
    For i = 0 To UBound(d)
        
        For j = 1 To Len(d(i))
            If Mid(d(i), j, 1) = "#" Then res = res + 1
        Next j
        
    Next i
    
    'Answer: 2166
    Range("D11A") = res

End Sub

Function getN(d, i, j)
    
    n = 0
    
    If i - 1 >= 0 Then
    
        If j - 1 > 0 Then If Mid(d(i - 1), j - 1, 1) = "#" Then n = n + 1
        If j + 1 <= Len(d(i)) Then If Mid(d(i - 1), j + 1, 1) = "#" Then n = n + 1
        
        If Mid(d(i - 1), j - 0, 1) = "#" Then n = n + 1
    End If
    
    If i + 1 <= UBound(d) Then
        If j - 1 > 0 Then If Mid(d(i + 1), j - 1, 1) = "#" Then n = n + 1
        If j + 1 <= Len(d(i)) Then If Mid(d(i + 1), j + 1, 1) = "#" Then n = n + 1
        
        If Mid(d(i + 1), j - 0, 1) = "#" Then n = n + 1
    End If
    
    If j - 1 > 0 Then If Mid(d(i), j - 1, 1) = "#" Then n = n + 1
    If j + 1 <= Len(d(i)) Then If Mid(d(i), j + 1, 1) = "#" Then n = n + 1
    
    getN = n

End Function

Function getN2(d, i, j)
    
    n = 0
    
    found = Array(False, False, False, False, False, False, False, False)
    For k = 1 To UBound(d)
    
        If i - k >= 0 Then
       
            If Not found(0) And j - k > 0 Then
                If Mid(d(i - k), j - k, 1) = "#" Then
                    n = n + 1
                    found(0) = True
                End If
                If Mid(d(i - k), j - k, 1) = "L" Then
                    found(0) = True
                End If
            End If
       
            If Not found(1) And j + k <= Len(d(i)) Then
                If Mid(d(i - k), j + k, 1) = "#" Then
                    n = n + 1
                    found(1) = True
                End If
                If Mid(d(i - k), j + k, 1) = "L" Then
                    found(1) = True
                End If
            End If
       
            If Not found(2) And Mid(d(i - k), j, 1) = "#" Then
                n = n + 1
                found(2) = True
            End If
            If Not found(2) And Mid(d(i - k), j, 1) = "L" Then
                found(2) = True
            End If
       
        End If
   
        If i + k <= UBound(d) Then
   
            If Not found(3) And j - k > 0 Then
                If Mid(d(i + k), j - k, 1) = "#" Then
                    n = n + 1
                    found(3) = True
                End If
                If Mid(d(i + k), j - k, 1) = "L" Then
                    found(3) = True
                End If
            End If
   
            If Not found(4) And j + k <= Len(d(i)) Then
                If Mid(d(i + k), j + k, 1) = "#" Then
                    n = n + 1
                    found(4) = True
                End If
                If Mid(d(i + k), j + k, 1) = "L" Then
                    found(4) = True
                End If
            End If
   
            If Not found(5) And Mid(d(i + k), j - 0, 1) = "#" Then
                n = n + 1
                found(5) = True
            End If
            If Not found(5) And Mid(d(i + k), j - 0, 1) = "L" Then
                found(5) = True
            End If
   
        End If
    
        If j - k > 0 Then
            If Not found(6) And Mid(d(i), j - k, 1) = "#" Then
                n = n + 1
                found(6) = True
            End If
            If Not found(6) And Mid(d(i), j - k, 1) = "L" Then
                found(6) = True
            End If
        End If
       
        If j + k <= Len(d(i)) Then
            If Not found(7) And Mid(d(i), j + k, 1) = "#" Then
                n = n + 1
                found(7) = True
            End If
            If Not found(7) And Mid(d(i), j + k, 1) = "L" Then
                found(7) = True
            End If
        End If
       
        If found(0) And found(1) And found(2) And found(3) And found(4) And found(5) And found(6) And found(7) Then Exit For
       
    Next k
    
    getN2 = n

End Function

Sub Day11B()

    ' Load input for puzzle
    d = Split(ReadFile("AoC11.txt"), vbNewLine)
    nd = d
    
    ' Initiate variables
    res = 0
    
    '
    changing = True
    num = 0
    Do While changing
        num = num + 1
        changing = False
        
        For i = 0 To UBound(d)
            
            For j = 1 To Len(d(i))
                
                n = 0
                
                If Mid(d(i), j, 1) = "L" Or Mid(d(i), j, 1) = "#" Then
                    
                    n = getN2(d, i, j)
                    
                End If
                
                If Mid(d(i), j, 1) = "L" And n = 0 Then
                    
                    changing = True
                    nd(i) = Left(nd(i), j - 1) & "#" & Mid(d(i), j + 1)
                    
                End If
                
                If Mid(d(i), j, 1) = "#" And n >= 5 Then
                    
                    changing = True
                    nd(i) = Left(nd(i), j - 1) & "L" & Mid(d(i), j + 1)
                    
                End If
                
            Next j
            
        Next i
        d = nd
        
    Loop
    
    For i = 0 To UBound(d)
        
        For j = 1 To Len(d(i))
            If Mid(d(i), j, 1) = "#" Then res = res + 1
        Next j
        
    Next i
    
    'Answer: 1955
    Range("D11B") = res

End Sub

