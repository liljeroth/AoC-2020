Attribute VB_Name = "AoC8"
Sub Day08_a()

    ' Initiate routine
    Worksheets("AoC 8").Activate
    
    Range("I6") = 0 ' Clean the result cell, just in case...
    acc = 0         ' Accumulator
    i = 0           ' Index, line in code
    iUsed = ""      ' "Array" of previous lines executed
    
    ' Read input data
    d = Split(ReadFile("AoC8Data.txt"), vbNewLine)
    
    ' Loop through to find all iBag containing a "shiny gold bag"
    Do While True
        
        iUsed = iUsed & "," & i & ","
    
        cmd = Split(d(i), " ")(0)
        arg = CInt(Split(d(i), " ")(1))
        
        If cmd = "jmp" Then i = i + arg - 1
        If cmd = "acc" Then acc = acc + arg
        i = i + 1
        
        If InStr(iUsed, "," & i & ",") > 0 Then Exit Do
        
    Loop
    
    MsgBox acc
    Range("I6") = acc

End Sub

Sub Day08_b()

    ' Initiate routine
    Worksheets("AoC 8").Activate
    
    Range("I6") = 0 ' Clean the result cell, just in case...
    acc = 0         ' Accumulator
    i = 0           ' Index, line in code
    iUsed = ""      ' "Array" of previous lines executed
    iChanged = ""   ' "Array" of previous lines that has been corrected
    
    ' Read input data
    d = Split(ReadFile("AoC8Data.txt"), vbNewLine)
    
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
    
    MsgBox acc
    Range("I8") = acc

End Sub

