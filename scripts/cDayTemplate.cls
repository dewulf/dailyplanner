Option Explicit

Private pWs As Worksheet

Private pCol_week As Range
Private pCol_date As Range
    
Private pRow_dayplan_sep As Range
Private pRow_dayplan_item As Range

Private template_start As String
Private end_of_table As String
Private system_templates As String

Public Property Get Row_Dayplan_Item() As Range
    Set Row_Dayplan_Item = pRow_dayplan_item
End Property

Public Property Get Ws() As Worksheet
    Set Ws = pWs
End Property

'----------------------------------------------------
' this function adds a new day after the current day
'----------------------------------------------------
Sub Insert_Day_Template(day_before_row As Integer)
       
    ' show a dialog to select the dayTemplate
    ' dialog returns the range from where the template to take

    Dim Insert_Date As Date, date_day_before As Date
    Dim dailyPlan As cDailyPlan
    Set dailyPlan = New cDailyPlan
    
    date_day_before = DateValue(dailyPlan.Col_Date.Cells(day_before_row, 1))
    Insert_Date = DateValue(date_day_before + 1)

    Dim form As New DayTemplateSelect
    form.Day_Templates_List = Get_Day_Templates
    form.Insert_Date = Insert_Date
    form.Show
    
    If form.Ok Then
        Dim day_template_range As String
        day_template_range = Get_Day_Template_Range(form.SelectedOption)

        Insert_Date = form.Insert_Date
               
        ' goto end of the before_day
        Dim r As Integer
        r = day_before_row
        Do While dailyPlan.Col_Date.Cells(r, 1).value = date_day_before
            r = r + 1
        Loop
        ' insert a day separator
        Call Copy_Rowrange(pWs, pRow_dayplan_sep, dailyPlan.Ws, r)
        ' copy the day template
        Call Copy_Rowrange(pWs, pWs.Range(day_template_range), dailyPlan.Ws, r + 1)
                
        dailyPlan.Col_Date.rows(CStr(r + 1) + ":" + CStr(r + pWs.Range(day_template_range).rows.Count)).value = Insert_Date
          
        dailyPlan.Conditional_Formating
    End If
    
End Sub

Private Function Get_Day_Template_Range(day_template_name As String) As String
    Dim r As Integer, tmp As String
    Dim rng_start As Integer, rng_end As Integer

    r = 1
    tmp = ""
    
    Do Until pCol_week.Cells(r, 1) = end_of_table
        If pCol_week.Cells(r, 1) = template_start And tmp = "" Then
            If pCol_date.Cells(r, 1) = day_template_name Then
                tmp = day_template_name
                rng_start = r + 1
            End If
        ElseIf (pCol_week.Cells(r, 1) = template_start And tmp <> "") _
            Or (pCol_week.Cells(r, 1) = system_templates And tmp <> "") Then
            rng_end = r - 1
            Exit Do
        End If
        r = r + 1
    Loop
    
    Get_Day_Template_Range = "$" + CStr(rng_start) + ":$" + CStr(rng_end)

End Function

Private Function Get_Day_Templates() As String()
    Dim day_template_names() As String
    Dim i As Integer, r As Integer
    
    i = 0
    r = 1
    Do Until pCol_week.Cells(r, 1) = end_of_table
        If pCol_week.Cells(r, 1) = template_start Then
            ReDim Preserve day_template_names(i)
            day_template_names(i) = pCol_date.Cells(r, 1)
            i = i + 1
        End If
        r = r + 1
    Loop
    Get_Day_Templates = day_template_names
End Function

Private Sub Class_Initialize()
    Set pWs = Worksheets("DayTemplates")
    
    template_start = "TEMPLATE_START"
    end_of_table = "XXXXXXXXXXXXXXX"
    system_templates = "SYSTEM_TEMPLATES"
    
    Set pCol_week = pWs.Range("wsr_daytemplates_week")
    Set pCol_date = pWs.Range("wsr_daytemplates_date")
    
    Set pRow_dayplan_sep = pWs.Range("default_dayplan_separator")
    Set pRow_dayplan_item = pWs.Range("default_dayplan_item")
End Sub
