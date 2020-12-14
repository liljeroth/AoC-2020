Attribute VB_Name = "Export"
Sub X()
    
    Dim objMyProj As VBProject
    Dim objVBComp As VBComponent
    
    Set objMyProj = Application.VBE.ActiveVBProject
    
    For Each objVBComp In objMyProj.VBComponents
        If objVBComp.Type = vbext_ct_StdModule Then
            objVBComp.Export ActiveWorkbook.Path & "\AoC Modules\" & objVBComp.Name & ".bas"
        End If
    Next

End Sub
