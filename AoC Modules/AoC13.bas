Attribute VB_Name = "AoC13"
Sub Day13A()

    ' Load input for puzzle
    d = Split(ReadFile("AoC13.txt"), vbNewLine)
    
    ' Initiate variables
    dTime = CLng(d(0)) ' Initial time from input data
    busDt = 100        ' Simply a high enough number to be overwritten
    
    ' Find shortest time
    For Each c In Split(d(1), ",")
    
        If IsNumeric(c) Then
            
            ' Get departure time
            depTime = 0
            If dTime Mod CLng(c) <> 0 Then depTime = CLng(c) - dTime Mod CLng(c)
            
            ' Test if this buss next departure time is closest
            If depTime < busDt Then
                busNo = CLng(c)
                busDt = depTime
            End If
            
        End If
        
    Next c
    
    'Answer: 3035
    Range("D13A") = busNo * busDt

End Sub

Sub Day13B()

    ' Load input for puzzle
    d = Split(ReadFile("AoC13.txt"), vbNewLine)
    
    ' Calculate relative buss schedule
    t = 0
    For Each c In Split(d(1), ",")
        If c <> "x" Then
            buss = buss & "," & c
            rels = rels & "," & t
        End If
        t = t + 1
    Next c
    
    ' Setup arrays
    buss = Split(Mid(buss, 2), ",")
    rels = Split(Mid(rels, 2), ",")
    
    ' Initiate first set of parameters
    c = UBound(buss) ' c = Current Buss
    i = CDec(0)      ' i = Current Time
    
    '
    Do While c > 0
    
        Do While True
        
            ' Get relative position to i
            nc = CDec(rels(c)) + i
            no = CDec(rels(0)) + i
            
            ' Mod operator overflows.. Plan B!
            nce = CDec(nc) / CDec(buss(c)) = Round(CDec(nc) / CDec(buss(c)), 0)
            noe = CDec(no) / CDec(buss(0)) = Round(CDec(no) / CDec(buss(0)), 0)
            
            ' If both are divisable (i.e. relative schedule matches) - > Exit
            If nce And noe Then Exit Do
            
            ' Otherwise, calculate next step skipping as much as possible
            ncv = CDec(buss(c))
            nov = CDec(buss(0))
            
            If Not nce Then ncv = ncv - CDec("0." & Split("" & nc / CDec(buss(c)), ".")(1)) * CDec(buss(c))
            If Not noe Then nov = nov - CDec("0." & Split("" & no / CDec(buss(0)), ".")(1)) * CDec(buss(0))
            
            ' Apply largest leap
            If ncv > nov Then
                i = i + ncv
            Else
                i = i + nov
            End If
            
        Loop
        
        ' Combine Buss C and Buss 0
        buss(0) = CDec(buss(c)) * CDec(buss(0))
        rels(0) = buss(0) - i
        
        ' Next Buss
        c = c - 1
        
    Loop
    
    'Answer: 725 169 163 285 238
    Range("D13B") = buss(0) - rels(0)

End Sub

