Function CONCATENATE_MULTIPLE(ref As Range, Separator As String) As String
    Dim cell As Range
    Dim result As String
    For Each cell In ref
        result = result & cell.value & Separator
    Next cell
    CONCATENATE_MULTIPLE = Left(result, Len(result) - 1)
End Function

Sub Move_Row(ws_from As Worksheet, row_from As Integer, ws_to As Worksheet, row_to As Integer)
    ws_to.rows(row_to).EntireRow.insert
    ws_from.rows(row_from).EntireRow.Copy ws_to.rows(row_to).EntireRow
    ws_from.rows(row_from).EntireRow.Delete
End Sub

Sub Copy_Row(ws_from As Worksheet, row_from As Integer, ws_to As Worksheet, row_to As Integer)
    ws_to.rows(row_to).EntireRow.insert
    ws_from.rows(row_from).EntireRow.Copy ws_to.rows(row_to).EntireRow
End Sub

Sub Copy_Rowrange(ws_from As Worksheet, row_range_from As Range, ws_to As Worksheet, row_to As Integer)
    Dim i As Integer
    ws_to.rows(CStr(row_to) + ":" + CStr(row_to + row_range_from.rows.Count - 1)).EntireRow.insert
    row_range_from.rows("1:" + CStr(row_range_from.rows.Count)).Copy _
        ws_to.rows(CStr(row_to) + ":" + CStr(row_to + row_range_from.rows.Count - 1))
End Sub

Sub Update_Task_Desc(ws_name As String, proj_rng As String, task_rng As String, project_name As String, old_task_desc As String, new_task_desc As String)
    i = 4
    Dim rng As String
    Do Until Worksheets(ws_name).Range(proj_rng).Cells(i, 1) = "XXXXXXXXXXXXXXX"
        If Worksheets(ws_name).Range(proj_rng).Cells(i, 1) = project_name And Worksheets(ws_name).Range(task_rng).Cells(i, 1) = old_task_desc Then
            Worksheets(ws_name).Range(task_rng).Cells(i, 1) = new_task_desc
        End If
        i = i + 1
    Loop
End Sub

Function Get_Final_Row(search_col As Range) As Integer
    Dim r As Integer
    r = 1
    Do Until search_col.Cells(r, 1) = "XXXXXXXXXXXXXXX"
        r = r + 1
    Loop
    Get_Final_Row = r
End Function

Function IsArrayAllocated(arr As Variant) As Boolean
        On Error Resume Next
        IsArrayAllocated = IsArray(arr) And _
                           Not IsError(LBound(arr, 1)) And _
                           LBound(arr, 1) <= UBound(arr, 1)
End Function