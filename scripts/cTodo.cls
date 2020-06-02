Option Explicit

Private pWs As Worksheet

Private pCol_date As Range ' the Date column
Private pCol_week As Range ' the Week column

Private todo_section_indicator As String
Private todo_section_name_nextdays As String
Private todo_section_name_followups As String
Private end_of_table As String

Public Property Get Ws() As Worksheet
    Set Ws = pWs
End Property

Public Sub Move_From_Dailyplan_To_Nextdays(dailyPlan As cDailyPlan, row As Integer)
    Dim r As Integer
    r = 1
    Do Until pCol_week.Cells(r, 1) = end_of_table
        If pCol_week.Cells(r, 1) = todo_section_indicator And _
            pCol_date.Cells(r, 1) = todo_section_name_nextdays Then
            Call Move_Row(dailyPlan.Ws, row, pWs, r + 1)
            pCol_date.Cells(r + 1, 1) = ""
            Exit Do
        End If
        r = r + 1
    Loop
End Sub

Public Sub Copy_From_Dailyplan_To_Followups(dailyPlan As cDailyPlan, row As Integer)
    Dim r As Integer
    r = 1
    Do Until pCol_week.Cells(r, 1) = end_of_table
        If pCol_week.Cells(r, 1) = todo_section_indicator And _
            pCol_date.Cells(r, 1) = todo_section_name_followups Then
            Call Copy_Row(dailyPlan.Ws, row, pWs, r + 1)
            dailyPlan.set_followup (row)
            Exit Do
        End If
        r = r + 1
    Loop
End Sub

Public Sub Add_To_Dailyplan_Someday(someday As Date, row As Integer)
    Dim dailyPlan As cDailyPlan
    Set dailyPlan = New cDailyPlan
    Call dailyPlan.Move_From_Todo_To_Someday(Me, someday, row)
End Sub

Public Sub Insert_New_Entry(row As Integer)
    Dim dayTemplate As cDayTemplate
    Set dayTemplate = New cDayTemplate
    Call Copy_Rowrange(dayTemplate.Ws, dayTemplate.Row_Dayplan_Item, pWs, row + 1)
End Sub

Private Sub Class_Initialize()
    Set pWs = Sheets("ToDo")
    
    Set pCol_date = pWs.Range("wsr_todo_date")
    Set pCol_week = pWs.Range("wsr_todo_week")
    
    todo_section_indicator = "TODO_START"
    todo_section_name_nextdays = "Next Days"
    end_of_table = "XXXXXXXXXXXXXXX"
    todo_section_name_followups = "Follow ups"
End Sub