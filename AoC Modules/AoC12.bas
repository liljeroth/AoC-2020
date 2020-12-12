Attribute VB_Name = "AoC12"
Sub Day12A()

    ' Load input for puzzle
    d = Split(ReadFile("AoC12.txt"), vbNewLine)
    
    ' Initiate variables
    res = 0
    e = 0
    n = 0
    
    
    '
    Direction = "E"
    For Each c In d
        If Mid(c, 1, 1) = "N" Then n = n + CLng(Mid(c, 2))
        If Mid(c, 1, 1) = "S" Then n = n - CLng(Mid(c, 2))
        
        If Mid(c, 1, 1) = "E" Then e = e + CLng(Mid(c, 2))
        If Mid(c, 1, 1) = "W" Then e = e - CLng(Mid(c, 2))
        
        If Mid(c, 1, 1) = "F" Then
        
            If Direction = "N" Then n = n + CLng(Mid(c, 2))
            If Direction = "S" Then n = n - CLng(Mid(c, 2))
            
            If Direction = "E" Then e = e + CLng(Mid(c, 2))
            If Direction = "W" Then e = e - CLng(Mid(c, 2))
        
        End If
        
        If Mid(c, 1, 1) = "R" Then
            
            'MsgBox c & " " & Direction & " = " & CInt(Mid(c, 2, 3)) / 90
            
            For i = 1 To CInt(Mid(c, 2, 3)) / 90
        
                If Direction = "N" Then
                    Direction = "E"
                ElseIf Direction = "S" Then
                    Direction = "W"
                
                ElseIf Direction = "E" Then
                    Direction = "S"
                ElseIf Direction = "W" Then
                    Direction = "N"
                End If
            
            Next i
            
            'MsgBox c & " -> " & Direction
        
        End If
        
        If Mid(c, 1, 1) = "L" Then
            
            'MsgBox c & " " & Direction & " = " & CInt(Mid(c, 2, 3)) / 90
            
            For i = 1 To CInt(Mid(c, 2, 3)) / 90
        
                If Direction = "E" Then
                    Direction = "N"
                ElseIf Direction = "W" Then
                    Direction = "S"
                
                ElseIf Direction = "S" Then
                    Direction = "E"
                ElseIf Direction = "N" Then
                    Direction = "W"
                End If
            
            Next i
            
            'MsgBox c & " -> " & Direction
        
        End If
        
    
    Next c
    
    'Answer: 1441
    res = Abs(e) + Abs(n)
    
    MsgBox "E:" & e & " + N:" & n & " = " & res
    Range("D12A") = res

End Sub

Sub Day12B()

    ' Load input for puzzle
    d = Split(ReadFile("AoC12.txt"), vbNewLine)
    
    ' Initiate variables
    shp = Array(0, 0)
    wpt = Array(10, 1)
    
    
    '
    For Each c In d
        If Mid(c, 1, 1) = "N" Then wpt(1) = wpt(1) + CLng(Mid(c, 2))
        If Mid(c, 1, 1) = "S" Then wpt(1) = wpt(1) - CLng(Mid(c, 2))
        
        If Mid(c, 1, 1) = "E" Then wpt(0) = wpt(0) + CLng(Mid(c, 2))
        If Mid(c, 1, 1) = "W" Then wpt(0) = wpt(0) - CLng(Mid(c, 2))
        
        If Mid(c, 1, 1) = "F" Then
            
            shp(0) = shp(0) + CLng(Mid(c, 2)) * wpt(0)
            shp(1) = shp(1) + CLng(Mid(c, 2)) * wpt(1)
            
        End If
        
        If Mid(c, 1, 1) = "R" Then
            
            
            For i = 1 To CInt(Mid(c, 2, 3)) / 90
                
                tmp = wpt(1)
                wpt(1) = -wpt(0)
                wpt(0) = tmp
                    
            Next i
        End If
        
        If Mid(c, 1, 1) = "L" Then
            
            For i = 1 To CInt(Mid(c, 2, 3)) / 90
                
                tmp = wpt(1)
                wpt(1) = wpt(0)
                wpt(0) = -tmp
            
            Next i
        
        End If
    
    Next c
    
    'Answer: 61616
    res = Abs(shp(0)) + Abs(shp(1))
    
    MsgBox "E:" & shp(0) & " + N:" & shp(1) & " = " & res
    Range("D12B") = res

End Sub

