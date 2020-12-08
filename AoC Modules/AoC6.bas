Attribute VB_Name = "AoC6"
Sub Day06_a()

    d = Split(Worksheets("AoC 6").Range("D4").Value, vbLf & vbLf)
    Worksheets("AoC 6").Range("I6") = 0
    
    res = 0
    For Each c In d
        
        res = res + Len(Trim(Replace(UNIQUECHARS("" & c), vbLf, "")))
    
    Next c
    
    Worksheets("AoC 6").Range("I6") = res
    
End Sub

Sub Day06_b()

    d = Split(Worksheets("AoC 6").Range("D4").Value, vbLf & vbLf)
    Worksheets("AoC 6").Range("I8") = 0
    
    res = 0
    For Each c In d
    
        f = Split(c, vbLf)(0)
        For i = 1 To Len(f)
        
            If Len(c) - Len(Replace(c, Mid(f, i, 1), "")) = Len(c) - Len(Replace(c, vbLf, "")) + 1 Then
                res = res + 1
            End If
            
        Next i
    
    Next c
    
    Worksheets("AoC 6").Range("I8") = res
    
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
