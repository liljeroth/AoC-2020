Attribute VB_Name = "AoC04"
Sub Day04A()
    
    ' Load input for puzzle
    d = Split(ReadFile("AoC04.txt"), vbNewLine & vbNewLine)
    
    ' Initiate variables
    res = 0
    
    For Each c In d
    
        If MeetRequirements(c) Then res = res + 1
        
    Next c
    
    ' Answer: 239
    Range("D04A") = res
    
End Sub

Function MeetRequirements(s)

    MeetRequirements = InStr(s, "byr") > 0 And InStr(s, "iyr") > 0 And InStr(s, "eyr") > 0 And InStr(s, "hgt") > 0 And InStr(s, "hcl") > 0 And InStr(s, "ecl") > 0 And InStr(s, "pid") > 0
    
End Function

Function GetValue(s, p, l)

    GetValue = Trim(Replace(Replace(Mid(s, InStr(s, p) + 4, l), vbNewLine, ""), vbCr, ""))
    
End Function

Sub Day04B()
    
    ' Load input for puzzle
    d = Split(ReadFile("AoC04.txt"), vbNewLine & vbNewLine)
    
    For Each c In d
        
        If MeetRequirements(c) Then
            
            byr = CInt(GetValue(c, "byr", 5))
            If Not (byr >= 1920 And byr <= 2002) Then GoTo continue
            
            iyr = CInt(GetValue(c, "iyr", 5))
            If Not (iyr >= 2010 And iyr <= 2020) Then GoTo continue
            
            eyr = CInt(GetValue(c, "eyr", 5))
            If Not (eyr >= 2020 And eyr <= 2030) Then GoTo continue
            
            If InStr(GetValue(c, "hgt", 5), "in") > 0 Then
            
                hgt = CLng(Trim(Replace(GetValue(c, "hgt", 5), "in", "")))
                If Not (hgt >= 59 And hgt <= 76) Then GoTo continue
                
            ElseIf InStr(GetValue(c, "hgt", 6), "cm") > 0 Then
            
                hgt = CInt(Trim(Replace(GetValue(c, "hgt", 6), "cm", "")))
                If Not (hgt >= 150 And hgt <= 193) Then GoTo continue
            
            Else
                
                GoTo continue
            
            End If
            
            hcl = GetValue(c, "hcl", 1) = "#" And Len(GetValue(c, "hcl", 8)) = 7
            If Not hcl Then GoTo continue
            
            tmp = GetValue(c, "ecl", 4)
            ecl = tmp = "amb" Or tmp = "blu" Or tmp = "brn" Or tmp = "gry" Or tmp = "grn" Or tmp = "hzl" Or tmp = "oth"
            If Not ecl Then GoTo continue
            
            pid = IsNumeric(GetValue(c, "pid", 10)) And (Len(GetValue(c, "pid", 10)) = 9)
            If Not pid Then GoTo continue
            
            ' The survivor!
            res = res + 1
            
        End If
        
continue:
        ' Do nothing
    Next c
    
    ' Answer: 188
    Range("D04B") = res
    
End Sub
