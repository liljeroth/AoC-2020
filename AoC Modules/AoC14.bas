Attribute VB_Name = "AoC14"
Sub Day14A()

    ' Load input for puzzle
    d = Split(ReadFile("AoC14.txt"), vbNewLine)
    
    ' Initiate variables
    Dim Memory(99999) As String
    Mask = ""
    
    '
    For i = 0 To UBound(d)
    
        If Left(d(i), 4) = "mask" Then
        
            Mask = Trim(Split(d(i), "=")(1))
            
        Else
        
            Value = Dec2Bin(CLng(Trim(Split(d(i), "=")(1))))
            Address = CLng(Trim(Split(Split(d(i), "[")(1), "]")(0)))
            
            Memory(Address) = Value
            
            For j = 1 To Len(Mask)
                
                If Len(Mask) > Len(Memory(Address)) Then Memory(Address) = Mid(Memory(Address), 1, j - 1) & "0" & Mid(Memory(Address), j)
                
                If Mid(Mask, j, 1) <> "X" Then Memory(Address) = Mid(Memory(Address), 1, j - 1) & Mid(Mask, j, 1) & Mid(Memory(Address), j + 1)
                
            Next j
            
        End If
        
    Next i
    
    For Each m In Memory
        
        If m <> "" Then
        
            m1 = CDec(Mid(m, 1, 18))
            m2 = CDec(Mid(m, 19))
            
            r1 = r1 + CLng(Bin2Dec(CDec(m1)))
            r2 = r2 + CLng(Bin2Dec(CDec(m2)))
        
        End If
        
    Next m
    
    res = r2 + 2 ^ 18 * r1
    
    'Answer: 6386593869035
    MsgBox res
    Range("D14A") = res

End Sub

Sub Day14B()

    ' WARNING! WARNING! WARNING! WARNING! WARNING!
    ' WARNING!  16 MINUTES OF EXECUTION!  WARNING!
    ' WARNING! WARNING! WARNING! WARNING! WARNING!

    ' Load input for puzzle
    d = Split(ReadFile("AoC14.txt"), vbNewLine)
    '295115141728
    
    ' Initiate variables
    res = 0
    addresslist = ""
    Dim Memory(99999) As String
    
    '
    Mask = ""
    For i = 0 To UBound(d)
    
        If Left(d(i), 4) = "mask" Then
        
            Mask = Trim(Split(d(i), "=")(1))
            
        Else
            
            Address = Dec2Bin(CDec(Trim(Split(Split(d(i), "[")(1), "]")(0))))
            
            For j = 1 To Len(Mask)
                
                If Len(Mask) > Len(Address) Then Address = Mid(Address, 1, j - 1) & "0" & Mid(Address, j)
                If Mid(Mask, j, 1) <> "0" Then Address = Mid(Address, 1, j - 1) & Mid(Mask, j, 1) & Mid(Address, j + 1)
                
            Next j
            
            Address = getAddress2(Address)
            
            For Each ma In Split(Address, ",")
            
                If ma = "" Then Exit For
                
                ma = CLng(Bin2Dec(CDec(Mid(ma, 1, Len(ma) - 18)))) * 2 ^ 18 + CLng(Bin2Dec(CDec(Mid(ma, Len(ma) - 18 + 1))))
                
                If InStr(addresslist, ma) = 0 Then addresslist = addresslist & ma & ","
                
                na = Mid(addresslist, 1, InStr(addresslist, ma))
                Memory(Len(na) - Len(Replace(na, ",", ""))) = Split(d(i), "=")(1)
                
            Next ma
            
        End If
        
    Next i
    
    For Each m In Memory
    
        If m = "" Then Exit For
        
        res = res + CDec(m)
        
    Next m
    
    'Answer: 4288986482164
    MsgBox res
    Range("D14B") = res

End Sub

Function getAddress2(a)

    getAddress2 = a
    For i = 1 To Len(a)
    
        If Mid(a, i, 1) = "X" Then
            
            newAddr = ""
            For Each addr In Split(getAddress2, ",")
            
                If addr = "" Then Exit For
                
                newAddr = newAddr & Mid(addr, 1, i - 1) & "0" & Mid(addr, i + 1) & ","
                newAddr = newAddr & Mid(addr, 1, i - 1) & "1" & Mid(addr, i + 1) & ","
            
            Next addr
            
            getAddress2 = newAddr
            
        End If
        
    Next i

End Function

Function Dec2Bin(ByVal DecimalIn As Variant, _
              Optional NumberOfBits As Variant) As String
    Dec2Bin = ""
    DecimalIn = Int(CDec(DecimalIn))
    Do While DecimalIn <> 0
        Dec2Bin = Format$(DecimalIn - 2 * Int(DecimalIn / 2)) & Dec2Bin
        DecimalIn = Int(DecimalIn / 2)
    Loop
    If Not IsMissing(NumberOfBits) Then
       If Len(Dec2Bin) > NumberOfBits Then
          Dec2Bin = "Error - Number exceeds specified bit size"
       Else
          Dec2Bin = Right$(String$(NumberOfBits, _
                    "0") & Dec2Bin, NumberOfBits)
       End If
    End If
End Function

Function Bin2Dec(BinaryString As String) As Variant
    Dim X As Integer
    For X = 0 To Len(BinaryString) - 1
        Bin2Dec = CDec(Bin2Dec) + Val(Mid(BinaryString, _
                  Len(BinaryString) - X, 1)) * 2 ^ X
    Next
End Function
