Attribute VB_Name = "AoC06"
Sub Day06A()
    
    ' Load input for puzzle
    d = Split(ReadFile("AoC06.txt"), vbNewLine & vbNewLine)
    
    ' Initiate variables
    res = 0
    
    For Each c In d
        
        res = res + Len(Trim(Replace(UNIQUECHARS("" & c), vbNewLine, "")))
    
    Next c
    
    ' Answer: 6457
    Range("D06A") = res
    
End Sub

Sub Day06B()
    
    ' Load input for puzzle
    d = Split(ReadFile("AoC06.txt"), vbNewLine & vbNewLine)
    
    ' Initiate variables
    res = 0
    
    For Each c In d
    
        f = Split(c, vbNewLine)(0)
        For i = 1 To Len(f)
        
            If Len(c) - Len(Replace(c, Mid(f, i, 1), "")) = (Len(c) - Len(Replace(c, vbNewLine, ""))) / 2 + 1 Then
                
                res = res + 1
                
            End If
            
        Next i
    
    Next c
    
    ' Answer: 3260
    Range("D06B") = res
    
End Sub

Public Function UNIQUECHARS(chtxt As String)
    Dim x, i As Long
    
    With CreateObject("Scripting.Dictionary")
    
        x = Split(StrConv(Replace(chtxt, " ", ""), 64), Chr(0))
        
        For i = 0 To UBound(x) - 1
            .Item(x(i)) = Empty
        Next
        
        UNIQUECHARS = Join(.keys, "")
        
    End With
    
End Function
