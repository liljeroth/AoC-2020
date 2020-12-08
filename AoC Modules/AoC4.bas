Attribute VB_Name = "AoC4"
Sub Day04_a()

    d = Split(Worksheets("AoC 4").Range("D4").Value, vbLf & vbLf)
    Worksheets("AoC 4").Range("I6") = 0
    
    res = 0
    For Each c In d
    
        'MsgBox c
        If (InStr(c, "byr") > 0) And (InStr(c, "iyr") > 0) And (InStr(c, "eyr") > 0) And (InStr(c, "hgt") > 0) And (InStr(c, "hcl") > 0) And (InStr(c, "ecl") > 0) And (InStr(c, "pid") > 0) Then
            res = res + 1
        End If
        
    Next c
    
    Worksheets("AoC 4").Range("I6") = res
    
End Sub

Sub Day04_b()
    
    On Error Resume Next

    d = Split(Worksheets("AoC 4").Range("D4").Value, vbLf & vbLf)
    Worksheets("AoC 4").Range("I8") = 0
    
    res = 0
    For Each c In d
        tmp = 0
        byr = 0
        iyr = 0
        eyr = 0
        hcl = False
        ecl = False
        pid = False
        
        If (InStr(c, "byr") > 0) And (InStr(c, "iyr") > 0) And (InStr(c, "eyr") > 0) And (InStr(c, "hgt") > 0) And (InStr(c, "hcl") > 0) And (InStr(c, "ecl") > 0) And (InStr(c, "pid") > 0) Then
        
            byr = CInt(Trim(Replace(Mid(c, InStr(c, "byr") + 4, 5), vbLf, "")))
            iyr = CInt(Trim(Replace(Mid(c, InStr(c, "iyr") + 4, 5), vbLf, "")))
            eyr = CInt(Trim(Replace(Mid(c, InStr(c, "eyr") + 4, 5), vbLf, "")))
            
            hgt = False
            If InStr(Split(Mid(c, InStr(c, "hgt") + 4, 5), " ")(0), "in") > 0 Then
                tmp1 = Mid(c, InStr(c, "hgt") + 4, 5)
                tmp1 = Trim(Replace(Replace(tmp1, "in", ""), vbLf, ""))
                
                hgt = CInt(tmp1) >= 59 And CInt(tmp1) <= 76
            ElseIf InStr(Split(Mid(c, InStr(c, "hgt") + 4, 5), " ")(0), "cm") > 0 Then
                tmp1 = Mid(c, InStr(c, "hgt") + 4, 5)
                tmp1 = Trim(Replace(Replace(tmp1, "cm", ""), vbLf, ""))
                
                hgt = CInt(tmp1) >= 150 And CInt(tmp1) <= 193
            End If
            
            hcl = Mid(c, InStr(c, "hcl") + 4, 1) = "#" And Len(Trim(Replace(Mid(c, InStr(c, "hcl") + 5, 7), vbLf, ""))) = 6
            
            tmp2 = Trim(Replace(Mid(c, InStr(c, "ecl") + 4, 4), vbLf, ""))
            ecl = (tmp2 = "amb") Or (tmp2 = "blu") Or (tmp2 = "brn") Or (tmp2 = "gry") Or (tmp2 = "grn") Or (tmp2 = "hzl") Or (tmp2 = "oth")
            
            tmp = CLng(Trim(Replace(Mid(c, InStr(c, "pid") + 4, 10), vbLf, "")))
            pid = (tmp > 0) And (Len(Replace(Replace(Trim(Mid(c, InStr(c, "pid") + 4, 10)), vbLf, ""), " ", "")) = 9)
            
            If byr >= 1920 And byr <= 2002 Then
                If iyr >= 2010 And iyr <= 2020 Then
                    If eyr >= 2020 And eyr <= 2030 Then
                        If hgt Then
                            If hcl Then
                                If ecl Then
                                    If pid Then
                                        res = res + 1
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            
        End If
        
    Next c
    
    Worksheets("AoC 4").Range("I8") = res
    
End Sub
