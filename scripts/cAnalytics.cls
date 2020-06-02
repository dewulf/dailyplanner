Option Explicit

Private pWs As Worksheet

Private pCol_owner As Range
Private pCol_project As Range
Private pCol_task As Range
Private pCol_hours As Range

Public Property Get Ws()
    Set Ws = pWs
End Property

Public Sub Do_Analytics(a_proj As Scripting.Dictionary, week As Integer, year As Integer)
    
    Dim proj As Variant, task As Variant
    Dim r As Integer
    
    ' find the row where to start inserting
    Dim ins_row As Integer, first_ins_row As Integer, last_ins_row As Integer
    ins_row = Get_Final_Row(pCol_project)
       
    ' Remember first and last row to insert
    ' format last row at the bottom doubledline
    ' insert after this rows with the project overall sums
    ' loop over all projects and insert sumif formula
    r = ins_row
    first_ins_row = r
    For Each proj In a_proj.Keys
        For Each task In a_proj(proj).Keys
            Debug.Print proj + " - " + task + " - " + CStr(a_proj(proj)(task))
            pWs.rows(r).EntireRow.insert
            With pCol_project
                .Cells(r, 1) = proj
                .Cells(r, 1).Font.Italic = False
                .Cells(r, 1).Font.Bold = False
            End With
            With pCol_task
                .Cells(r, 1) = task
                .Cells(r, 1).Font.Italic = False
                .Cells(r, 1).Font.Bold = False
            End With
            With pCol_hours
                .Cells(r, 1) = a_proj(proj)(task)
                .Cells(r, 1).Font.Italic = False
                .Cells(r, 1).Font.Bold = False
            End With
            r = r + 1
        Next task
    Next proj
    last_ins_row = r - 1
    
    With pWs.Range(pCol_project.Cells(r - 1, 1), pCol_hours.Cells(r - 1, 1)).Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    
    ' insert the overall calculations
    For Each proj In a_proj.Keys
        pWs.rows(r).EntireRow.insert
        With pCol_project
            .Cells(r, 1) = proj
            .Cells(r, 1).Font.Italic = True
            .Cells(r, 1).Font.Bold = True
        End With
        With pCol_task
            .Cells(r, 1) = "Overall"
            .Cells(r, 1).Font.Italic = True
            .Cells(r, 1).Font.Bold = True
        End With
        With pCol_hours
            .Cells(r, 1).formula = "=SUMIF($B$" + CStr(first_ins_row) + _
                                   ":$B$" + CStr(last_ins_row) + ",B" + CStr(r) + _
                                   ",$D$" + CStr(first_ins_row) + ":$D$" + CStr(last_ins_row) + ")"
            .Cells(r, 1).Font.Italic = True
            .Cells(r, 1).Font.Bold = True
        End With
        r = r + 1
    Next proj
    
    ' copy the second row
    Copy_Row pWs, 2, pWs, ins_row
        
    pCol_project.Cells(ins_row, 1) = week
    pCol_task.Cells(ins_row, 1) = year
    
End Sub

Private Sub Class_Initialize()
    Set pWs = Worksheets("Analytics")
    
    Set pCol_owner = pWs.Range("wsr_analytics_owner")
    Set pCol_project = pWs.Range("wsr_analytics_project")
    Set pCol_task = pWs.Range("wsr_analytics_task")
    Set pCol_hours = pWs.Range("wsr_analytics_hours")
End Sub
