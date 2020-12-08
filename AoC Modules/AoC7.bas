Attribute VB_Name = "AoC7"
Sub Day07_a()

    ' Initiate routine
    Worksheets("AoC 7").Activate
    
    Range("I6") = 0 ' Clean the result cell, just in case...
    iBag = ""       ' List of identified bags
    sBag = 0        ' Sum of bags
    
    ' Read input data
    d = Split(ReadFile("AoC7Data.txt"), vbNewLine)
    bags = ""
    nbag = ","
    
    ' Loop through to find all bags containing a "shiny gold bag"
    For Each c In d
        
        If InStr(Split(c, "contain")(1), "shiny gold bag") > 0 Then
        
            bag = cleanBagName(Trim(Split(c, "contain")(0)))
            bags = bags & "," & cleanBagName(findParentBags(d, bag))
        
        End If
            
    Next c
    
    For Each b In Split(bags, ",")
    
        If b <> "" And InStr(nbag, "," & b & ",") = 0 Then
        
            nbag = nbag & b & ","
        
        End If
    
    Next b
    
    MsgBox Len(nbag) - Len(Replace(nbag, ",", "")) - 1
    Range("I6") = Len(nbag) - Len(Replace(nbag, ",", "")) - 1

End Sub

Sub Day07_b()

    ' Initiate routine
    Worksheets("AoC 7").Activate
    
    Range("I8") = 0 ' Clean the result cell, just in case...
    iBag = ""       ' List of identified bags
    sBag = 0        ' Sum of bags
    
    ' Read input data
    d = Split(ReadFile("AoC7Data.txt"), vbNewLine)
    
    ' Loop through all bags to find the one we are looking for
    For Each c In d
        
        ' If current is the bag
        If InStr(Split(c, "contain")(0), "shiny gold") > 0 Then
            
            ' Find children bags
            sBag = findNofChildrenBags(d, Split(c, "contain")(1))
                
            ' As bag was found, exit loop
            Exit For
            
        End If
    
    Next c
    
    Worksheets("AoC 7").Range("I8") = sBag

End Sub

Function cleanBagName(n)
    
    cleanBagName = n
    
    cleanBagName = Replace(cleanBagName, "bags", "")
    cleanBagName = Replace(cleanBagName, "bag", "")
    cleanBagName = Replace(cleanBagName, ".", "")
    
    cleanBagName = Trim(cleanBagName)
    
End Function

Function findParentBags(fullList, child)

    ' Loop though the list to fint parent bag
    For Each c In fullList
        
        ' If current bag is parent
        If InStr(Split(c, "contain")(1), child) > 0 Then
        
            ' Clean the string of found bag
            fBag = cleanBagName(Split(c, "contain")(0))
            iBag = iBag & "," & bag & "," & findParentBags(fullList, fBag)
        
        End If
        
    Next c
    
    findParentBags = iBag & "," & child

End Function

Function findNofChildrenBags(fullList, children)
    findNofChildrenBags = 0
    
    ' Loop each child bag
    For Each b In Split(children, ",")
        b = Trim(b)
        
        ' Extract name and count
        bagName = cleanBagName(Mid(b, InStr(b, " ")))
        nofBags = CInt(Trim(Mid(b, 1, InStr(b, " "))))
        
        ' Find bag in the full list
        For Each c In fullList
            
            ' See if current item is the one we are looking for
            If InStr(Split(c, "contain")(0), bagName) > 0 Then
                
                ' If found, extract children names
                newChildren = cleanBagName(Split(c, "contain")(1))
                
                ' If children, perform nested function call
                If InStr(newChildren, "no other") = 0 Then
                    findNofChildrenBags = findNofChildrenBags + nofBags * (findNofChildrenBags(fullList, newChildren) + 1)
                Else
                    findNofChildrenBags = findNofChildrenBags + nofBags
                End If
                
                ' As bag was found, exit loop
                Exit For
                
            End If
            
        Next c
        
    Next b
        
End Function

