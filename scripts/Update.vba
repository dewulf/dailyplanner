Sub update_workbook()
    ' set version name in configuration
    Dim current_version As String
    current_version = Sheets("Configuration").Cells(1, 11).value
    
    ' Remove all missing references
    ' do late binding only
    Dim theRef As Variant, i As Long

    ' loop through all References in VB Project
    For i = ThisWorkbook.VBProject.References.Count To 1 Step -1
        Set theRef = ThisWorkbook.VBProject.References.item(i)
    
        ' if reference is "Missing" >> remove it to avoid error message
        If theRef.isbroken = True Then
            ThisWorkbook.VBProject.References.remove theRef
        End If
    
        ' just for Debug
        ' Debug.Print theRef.Description & ";" & theRef.FullPath & ";" & theRef.isbroken & vbCr
    Next i
    
    If current_version = "0.1" Then
        Dim new_version As String
        new_version = "0.11"
        Sheets("Configuration").Cells(1, 11).value = new_version
        
        
    End If
    
End Sub