Attribute VB_Name = "AoC14"
Sub Day14A()

    ' Load input for puzzle
    d = Split(ReadFile("AoC14.txt"), vbNewLine)
    
    ' Initiate variables
    res = 0
    Dim Memory(9999) As String
    For i = 0 To UBound(Memory)
        Memory(i) = "000000000000000000000000000000000000"
    Next i
    
    '
    Mask = Trim(Split(d(0), "=")(1))
    For i = 1 To UBound(d)
        If Left(d(i), 4) = "mask" Then
        
            Mask = Trim(Split(d(i), "=")(1))
            
        Else
        
            Value = Dec2Bin(CLng(Trim(Split(d(i), "=")(1))))
            Address = CLng(Trim(Split(Split(d(i), "[")(1), "]")(0)))
            
            Memory(Address) = Value
            For j = Len(Memory(Address)) + 1 To Len(Mask)
                Memory(Address) = "0" & Memory(Address)
            Next j
            
            For j = 0 To Len(Mask) - 1
            
                k = Len(Mask) - j
                
                If Mid(Mask, k, 1) = "0" Then Memory(Address) = Mid(Memory(Address), 1, k - 1) & "0" & Mid(Memory(Address), k + 1)
                If Mid(Mask, k, 1) = "1" Then Memory(Address) = Mid(Memory(Address), 1, k - 1) & "1" & Mid(Memory(Address), k + 1)
                
            Next j
            
        End If
    Next i
    
    For Each m In Memory
        
        m1 = CDec(Mid(m, 1, 18))
        m2 = CDec(Mid(m, 19))
        
        r1 = r1 + CLng(Bin2Dec(CDec(m1)))
        r2 = r2 + CLng(Bin2Dec(CDec(m2)))
        
    Next m
    
    res = r2 + 2 ^ 18 * r1
    
    'Answer:
    MsgBox res
    Range("D14A") = res

End Sub

Sub Day14B()

    ' WARNING! WARNING! WARNING! WARNING! WARNING!
    ' WARNING!  20 MINUTES OF EXECUTION!  WARNING!
    ' WARNING! WARNING! WARNING! WARNING! WARNING!

    ' Load input for puzzle
    d = Split(ReadFile("AoC14.txt"), vbNewLine)
    
    ' Initiate variables
    res = 0
    addresslist = ""
    Dim Memory(99999) As String
    For i = 0 To UBound(Memory)
        Memory(i) = "000000000000000000000000000000000000"
    Next i
    
    '
    Mask = Trim(Split(d(0), "=")(1))
    For i = 1 To UBound(d)
        If Left(d(i), 4) = "mask" Then
            Mask = Trim(Split(d(i), "=")(1))
        Else
            
            Value = Dec2Bin(CDec(Trim(Split(d(i), "=")(1))))
            Address = Dec2Bin(CDec(Trim(Split(Split(d(i), "[")(1), "]")(0))))
            
            For j = Len(Address) + 1 To Len(Mask)
                Address = "0" & Address
            Next j
            
            For j = 1 To Len(Mask)
                If Mid(Mask, j, 1) <> "0" Then Address = Mid(Address, 1, j - 1) & Mid(Mask, j, 1) & Mid(Address, j + 1)
            Next j
            
            Address = Replace(getAddress(Address, 1), ",,", ",")
            Address = Replace(Address, ",,", ",")
            Address = Replace(Address, ",,", ",")
            Address = Replace(Address, ",,", ",")
            Address = Replace(Trim(Replace(Address, ",", " ")), " ", ",")
            
            For Each ma In Split(Address, ",")
        
                ma1 = CDec(Mid(ma, 1, 18))
                ma2 = CDec(Mid(ma, 19))
                
                ma = CLng(Bin2Dec(CDec(ma1))) * 2 ^ 18 + CLng(Bin2Dec(CDec(ma2)))
                
                If InStr(addresslist, ma) = 0 Then
                    addresslist = addresslist & ma & ","
                End If
                
                na = Mid(addresslist, 1, InStr(addresslist, ma))
                myAddress = Len(na) - Len(Replace(na, ",", ""))
                
                Memory(myAddress) = Value
                For j = Len(Memory(myAddress)) + 1 To Len(Mask)
                    Memory(myAddress) = "0" & Memory(myAddress)
                Next j
                
            Next ma
            
        End If
        
    Next i
    
    For Each m In Memory
        
        m1 = CDec(Mid(m, 1, 18))
        m2 = CDec(Mid(m, 19))
        
        r1 = r1 + CLng(Bin2Dec(CDec(m1)))
        r2 = r2 + CLng(Bin2Dec(CDec(m2)))
        
        'MsgBox CLng(m)
    Next m
    
    res = r2 + 2 ^ 18 * r1
    
    'Answer:
    MsgBox res
    Range("D14B") = res

End Sub

Function getAddress(a, init)

    getAddress = ""
    freeX = True
    For i = init To Len(a)
    
        If Mid(a, i, 1) = "X" Then
            freeX = False
            
            getAddress = getAddress & "," & getAddress(Mid(a, 1, i - 1) & "0" & Mid(a, i + 1), i)
            getAddress = getAddress & "," & getAddress(Mid(a, 1, i - 1) & "1" & Mid(a, i + 1), i)
            
            Exit For
        End If
        
    Next i
    
    If freeX Then getAddress = a

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
