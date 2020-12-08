Attribute VB_Name = "AoC5"
Sub Day05_a()

    d = Split(Worksheets("AoC 5").Range("D4").Value, vbLf)
    Worksheets("AoC 5").Range("I6") = 0
    
    res = 0
    For Each c In d
    
        vMin = 0
        vMax = 127
        For i = 1 To 7
            
            If Mid(c, i, 1) = "F" Then
                vMax = vMax - (vMax - vMin + 1) / 2
            Else
                vMin = vMin + (vMax - vMin + 1) / 2
            End If
            
        Next i
        
        seatRow = vMax
        
        vMin = 0
        vMax = 7
        For i = 8 To 10
            
            If Mid(c, i, 1) = "L" Then
                vMax = vMax - (vMax - vMin + 1) / 2
            Else
                vMin = vMin + (vMax - vMin + 1) / 2
            End If
            
        Next i
        
        seatCol = vMax
        
        If seatRow * 8 + seatCol > res Then
            res = seatRow * 8 + seatCol
        End If
    Next c
    
    Worksheets("AoC 5").Range("I6") = res
    
End Sub

Sub Day05_b()

    d = Split(Worksheets("AoC 5").Range("D4").Value, vbLf)
    Worksheets("AoC 5").Range("I8") = 0
    
    res = 0
    
    For Each c1 In d
        For Each c2 In d
            If c1 <> c2 Then
                
                v1 = getID(c1)
                v2 = getID(c2)
                
                If v1 = v2 + 2 Then
                    
                    foundSeat = False
                    
                    For Each c3 In d
                        If c1 <> c3 And c2 <> c3 Then
                                
                            v3 = getID(c3)
                                
                            If v1 = v3 + 1 Then
                                foundSeat = True
                                Exit For
                            End If
                            
                        End If
                    Next c3
                    
                    If Not foundSeat Then
                        res = v2 + 1
                    End If
                    
                End If
                
            End If
        Next c2
    Next c1
    
    Worksheets("AoC 5").Range("I8") = res
    
End Sub

Function getID(ticket)

    vMin = 0
    vMax = 127
    For i = 1 To 7
        
        If Mid(ticket, i, 1) = "F" Then
            vMax = vMax - (vMax - vMin + 1) / 2
        Else
            vMin = vMin + (vMax - vMin + 1) / 2
        End If
        
    Next i
    seatRow = vMax
    
    vMin = 0
    vMax = 7
    For i = 8 To 10
        
        If Mid(ticket, i, 1) = "L" Then
            vMax = vMax - (vMax - vMin + 1) / 2
        Else
            vMin = vMin + (vMax - vMin + 1) / 2
        End If
        
    Next i
    seatCol = vMax
    
    getID = seatRow * 8 + seatCol
    
End Function
