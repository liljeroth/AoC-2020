Attribute VB_Name = "AoC15"
Sub Day15A()
    
    ' Examples
    'Start = Array(0, 3, 6) ' 2020 => 436
    'Start = Array(1, 3, 2) ' 2020 => 1
    'Start = Array(2, 1, 3) ' 2020 => 10
    'Start = Array(1, 2, 3) ' 2020 => 27
    
    ' Puzzle data
    Start = Array(1, 17, 0, 10, 18, 11, 6) ' 2020 => 595
    
    ' Initiate variables
    Dim turns(2200) As Integer
    
    '
    For i = 1 To 2019
    
        If i <= UBound(Start) + 1 Then turn = Start(i - 1)
        
        'text = text & vbNewLine & i & ": " & turn
        
        If turns(turn) = 0 Then
        
            turns(turn) = i
            turn = 0
            
        Else
        
            tmp = i - turns(turn)
            turns(turn) = i
            turn = tmp
        End If
        
    Next i
    
    'Answer: 595
    Range("D15A") = turn

End Sub

Sub Day15B()

    ' Examples
    'Start = Array(0, 3, 6) ' 30000000 => 175594
    'Start = Array(1, 3, 2) ' 30000000 => 2578
    'Start = Array(2, 1, 3) ' 30000000 => 3544142
    'Start = Array(1, 2, 3) ' 30000000 => 261214
    
    ' Puzzle data
    Start = Array(1, 17, 0, 10, 18, 11, 6)
    
    ' Initiate variables
    Dim turns(30000000) As Long
    
    '
    For i = 1 To 30000000 - 1
    
        If i <= UBound(Start) + 1 Then turn = Start(i - 1)
        
        If turns(turn) = 0 Then
        
            turns(turn) = CDec(i)
            turn = 0
            
        Else
        
            tmp = CDec(i) - CDec(turns(turn))
            turns(CDec(turn)) = CDec(i)
            turn = CDec(tmp)
        End If
        
    Next i
    
    'Answer: 1708310
    Range("D15B") = turn

End Sub

