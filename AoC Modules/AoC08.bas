Attribute VB_Name = "AoC08"
Sub Day08A()
    
    ' Load input for puzzle
    d = Split(ReadFile("AoC08.txt"), vbNewLine)
    
    ' Initiate variables
    acc = 0         ' Accumulator
    i = 0           ' Index, line in code
    iUsed = ""      ' "Array" of previous lines executed
    
    '
    Do While True
        
        iUsed = iUsed & "," & i & ","
    
        cmd = Split(d(i), " ")(0)
        arg = CInt(Split(d(i), " ")(1))
        
        If cmd = "jmp" Then i = i + arg - 1
        If cmd = "acc" Then acc = acc + arg
        i = i + 1
        
        If InStr(iUsed, "," & i & ",") > 0 Then Exit Do
        
    Loop
    
    'Answer: 1766
    Range("D08A") = acc

End Sub

Sub Day08B()
    
    ' Load input for puzzle
    d = Split(ReadFile("AoC08.txt"), vbNewLine)
    
    ' Initiate variables
    acc = 0         ' Accumulator
    i = 0           ' Index, line in code
    iUsed = ""      ' "Array" of previous lines executed
    iChanged = ""   ' "Array" of previous lines that has been corrected
    
    '
    changed = False
    Do While i <= UBound(d)
        
        iUsed = iUsed & "," & i & ","
    
        cmd = Split(d(i), " ")(0)
        arg = CInt(Split(d(i), " ")(1))
        
        If InStr(iChanged, "," & i & ",") = 0 And (cmd = "jmp" Or cmd = "nop") And Not changed Then
            
            changed = True
            iChanged = iChanged & "," & i & ","
            
            If cmd = "nop" Then i = i + arg - 1
            
        Else
        
            If cmd = "jmp" Then i = i + arg - 1
            If cmd = "acc" Then acc = acc + arg
            
        End If
        i = i + 1
        
        If InStr(iUsed, "," & i & ",") > 0 Then
        
            i = 0
            acc = 0
            changed = False
            iUsed = ""
            
        End If
        
    Loop
    
    'Answer: 1639
    Range("D08B") = acc

End Sub

