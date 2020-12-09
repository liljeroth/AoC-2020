Attribute VB_Name = "AoC05"
Sub Day05A()
    
    ' Load input for puzzle
    d = Split(ReadFile("AoC05.txt"), vbNewLine)
    
    ' Initiate variables
    res = 0
    
    For Each c In d
            
        If getID(c) > res Then res = getID(c)
        
    Next c
    
    ' Answer: 866
    Range("D05A") = res
    
End Sub

Sub Day05B()
    
    ' Load input for puzzle
    d = Split(ReadFile("AoC05.txt"), vbNewLine)
    
    ' Initiate variables
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
    
    ' Answer: 583
    Range("D05B") = res
    
End Sub

Function getID(ticket)
    
    vMin = 0
    vMax = 127
    
    For i = 1 To Len(ticket)
        
        If Mid(ticket, i, 1) = "F" Or Mid(ticket, i, 1) = "L" Then
        
            vMax = vMax - (vMax - vMin + 1) / 2
            
        Else
        
            vMin = vMin + (vMax - vMin + 1) / 2
            
        End If
        
        If i = 7 Then
            
            tmp = vMax * 8
        
            vMin = 0
            vMax = 7
            
        End If
        
    Next i
    
    getID = tmp + vMax
    
End Function
