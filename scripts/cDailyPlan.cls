Option Explicit

Private pWs As Worksheet

Private pCol_week As Range
Private pCol_date As Range
Private pCol_project As Range
Private pCol_task As Range
Private pCol_fixed_start As Range
Private pCol_start As Range
Private pCol_end As Range
Private pCol_followup As Range
Private pCol_done As Range
Private pCol_duration As Range
Private pCol_activity_type As Range
Private pCol_activity As Range
Private pCol_booked_ref As Range
Private pCol_rm_spenttime_id As Range
Private pCol_force_update As Range
Private pCol_booked_time As Range
Private pCol_cal_indicator As Range
Private pCol_cal_id As Range
Private pCol_cal_link As Range

Private end_of_table As String

Public Property Get Ws() As Worksheet
    Set Ws = pWs
End Property

Public Property Get Col_Date()
    Set Col_Date = pCol_date
End Property

Public Property Get Col_Project()
    Set Col_Project = pCol_project
End Property

Public Property Get Col_Task()
    Set Col_Task = pCol_task
End Property

Public Property Get Col_Done()
    Set Col_Done = pCol_done
End Property

Public Sub set_followup(row As Integer)
    pCol_followup.Cells(row, 1) = "x"
End Sub

Public Sub Do_Analytics(row As Integer)
    Dim a_week As Integer, a_year As Integer
    
    a_week = pCol_week.Cells(row, 1)
    a_year = year(DateValue(pCol_date.Cells(row, 1)))
    
    Dim a_proj As Scripting.Dictionary
    Set a_proj = New Scripting.Dictionary
    
    Dim a_task As Scripting.Dictionary
    Dim proj As String, task As String
    Dim hours As Double, datum As Date, activity As String
    
    Dim conf As New cConfiguration
    
    Dim item As Variant
    Dim enabled_proj As New Scripting.Dictionary
    For Each item In conf.Projects.rows
        If conf.Ws.Cells(item.row, item.column + 1) = "x" Then
            enabled_proj.add item.value, 1
        End If
    Next
    
    Dim r As Integer
    r = 3
    Do Until pCol_date.Cells(r, 1) = end_of_table
        If pCol_week.Cells(r, 1) <> "" Then
            If pCol_week.Cells(r, 1).value = a_week And _
               year(DateValue(pCol_date.Cells(r, 1))) = a_year And _
               pCol_project.Cells(r, 1) <> "" And _
               enabled_proj.exists(pCol_project.Cells(r, 1).value) And _
               pCol_done.Cells(r, 1) = "x" Then
                         
                proj = pCol_project.Cells(r, 1).value
                task = pCol_task.Cells(r, 1).value
                hours = pCol_duration.Cells(r, 1) * 24
                
                datum = pCol_date.Cells(r, 1)
                activity = pCol_activity.Cells(r, 1)
                
                If task = "" Then
                    task = "*"
                End If
                
                If a_proj.exists(proj) Then
                    If a_proj(proj).exists(task) Then
                        ' sumup hours
                        a_proj(proj)(task) = a_proj(proj)(task) + hours
                    ElseIf Not a_proj(proj).exists(task) Then
                        a_proj(proj).add task, hours
                    End If
                Else
                    Set a_task = New Scripting.Dictionary
                    a_task.add task, hours
                    
                    a_proj.add proj, a_task
                End If
            End If
        End If
        r = r + 1
    Loop
    
    Dim analytics As New cAnalytics
    analytics.Do_Analytics a_proj, a_week, a_year
    
End Sub

Public Sub Insert_New_Entry(row As Integer)
    Dim dayTemplate As cDayTemplate
    Set dayTemplate = New cDayTemplate
    Call Copy_Rowrange(dayTemplate.Ws, dayTemplate.Row_Dayplan_Item, pWs, row + 1)
    pCol_date.Cells(row + 1, 1) = pCol_date.Cells(row, 1)
    Me.Conditional_Formating
End Sub

Public Sub Move_To_Tomorrow(row As Integer)
    Dim r As Integer
    r = 3
    Do Until pCol_date.Cells(r, 1) = end_of_table
        If pCol_date.Cells(r, 1).value = DateValue(Now() + 1) Then
            Call Move_Row(pWs, row, pWs, r)
            pCol_date.Cells(r - 1, 1) = pCol_date.Cells(r, 1)
            Exit Do
        End If
        r = r + 1
    Loop
    Me.Conditional_Formating
End Sub

Public Sub Move_To_Todo_Nextdays(row As Integer)
    Dim todo As cTodo
    Set todo = New cTodo
    todo.Move_From_Dailyplan_To_Nextdays Me, row
    Me.Conditional_Formating
End Sub

Public Sub Copy_To_Todo_Followups(row As Integer)
    Dim todo As cTodo
    Set todo = New cTodo
    todo.Copy_From_Dailyplan_To_Followups Me, row
    Me.Conditional_Formating
End Sub

Public Sub Move_From_Todo_To_Someday(todo As cTodo, someday As Date, row As Integer)
    Dim r As Integer
    r = 3
    Do Until pCol_date.Cells(r, 1) = end_of_table
        If pCol_date.Cells(r, 1).value = someday Then
            Call Move_Row(todo.Ws, row, pWs, r)
            pCol_date.Cells(r, 1) = pCol_date.Cells(r + 1, 1)
            Exit Do
        End If
        r = r + 1
    Loop
    Me.Conditional_Formating
End Sub

Public Sub Conditional_Formating()
    Dim rng_all As Range
    Dim rng_end_part As Range
    Dim r As Integer

    r = 4
    Do Until pCol_date.Cells(r, 1) = end_of_table
        r = r + 1
    Loop
    Set rng_all = Range(pWs.Cells(4, 1), pWs.Cells(r - 1, 22))
    Set rng_end_part = Range(pCol_end.Cells(4, 1), pCol_end.Cells(r - 1, 1))

    pWs.Cells.FormatConditions.Delete
              
    Dim formula As String
    
    If Application.LanguageSettings.LanguageID(msoLanguageIDInstall) = msoLanguageIDGerman Then
        formula = "=BEREICH.VERSCHIEBEN(INDIREKT(ADRESSE(ZEILE(); SPALTE()));1;-1)"
    ElseIf Application.LanguageSettings.LanguageID(msoLanguageIDInstall) = msoLanguageIDRussian Then
        formula = "=����(������(�����(������(); �������()));1;-1)"
    Else
        formula = "=OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),1,-1)"
    End If
    
    With rng_end_part.FormatConditions.add _
        (Type:=xlCellValue, Operator:=xlGreater, Formula1:=formula)
        With .Interior
            .PatternColorIndex = xlAutomatic
            .Color = 49407
            .TintAndShade = 0
        End With
        .StopIfTrue = False
    End With
    
    With rng_all.FormatConditions.add _
        (Type:=xlExpression, Formula1:="=$M4=""x""")
        With .Font
            .Strikethrough = True
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.319946592608417
        End With
        .StopIfTrue = False
    End With
        
    If Application.LanguageSettings.LanguageID(msoLanguageIDInstall) = msoLanguageIDGerman Then
        formula = "=$B4=HEUTE()"
    ElseIf Application.LanguageSettings.LanguageID(msoLanguageIDInstall) = msoLanguageIDRussian Then
        formula = "=$B4=�������()"
    Else
        formula = "=$B4=TODAY()"
    End If
    
        
    With rng_all.FormatConditions.add _
        (Type:=xlExpression, Formula1:=formula)
        With .Font
            .Bold = True
            .Italic = True
            .TintAndShade = 0
        End With
        With .Interior
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent3
            .TintAndShade = 0.399945066682943
        End With
        .StopIfTrue = False
    End With
End Sub

Public Sub Book_Spent_Time_To_Redmine(book_date As String)
    Dim z As Long
    Dim rm_task_id As Integer, rm_task As String, rm_project As String
    
    Dim tsk As cTasks
    Set tsk = New cTasks
    
    Dim redmine As cRedmineTask
    Set redmine = New cRedmineTask
    
    'Find all tasks that have redmine task ids and put them on a list
    z = 3
    Do Until tsk.Col_Project.Cells(z, 1) = ""
        If tsk.Col_Rm_Taskid.Cells(z, 1) <> "" Then 'if the task has an assigned redmine issue_id
            rm_task_id = tsk.Col_Rm_Taskid.Cells(z, 1)
            rm_task = tsk.Col_Taskdescription.Cells(z, 1)
            rm_project = tsk.Col_Project.Cells(z, 1)
            
            Debug.Print (rm_task_id)
            Dim r As Integer, dp_project As String, dp_task As String
            Dim dp_hours As Double, dp_comments As String, dp_spent_on As String, dp_hash As String
            Dim time_entry_id As Integer
            r = 3
            Do Until pCol_date.Cells(r, 1) = end_of_table
                If pCol_date.Cells(r, 1).value = DateValue(book_date) And pCol_task.Cells(r, 1) <> "" Then
                    dp_project = pCol_project.Cells(r, 1)
                    dp_task = pCol_task.Cells(r, 1)
                    Debug.Print ("Checking: " + dp_project + " - " + dp_task)
                    
                    If (dp_project = rm_project) And (dp_task = rm_task) And (pCol_done.Cells(r, 1) = "x") Then
                        
                        dp_hours = pCol_duration.Cells(r, 1) * 24
                        'dp_comments = Format(.Range("wsr_dailyplan_start").Cells(r, 1), "hh:mm") + " - "
                        'dp_comments = dp_comments + Format(.Range("wsr_dailyplan_end").Cells(r, 1), "hh:mm") + " : "
                        dp_comments = pCol_activity_type.Cells(r, 1) + " : " + pCol_activity.Cells(r, 1)
                        
                        dp_spent_on = Format(book_date, "yyyy-mm-dd")
                        dp_hash = SHA1TRUNC(CONCATENATE_MULTIPLE(pWs.Range(pWs.Cells(r, 1), pWs.Cells(r, 11)), ""))
                        
                        If (IsEmpty(pCol_booked_ref.Cells(r, 1))) Then
                            ' new spent time entry in Redmine
                            time_entry_id = redmine.Redmine_Time_Entry("POST", rm_task_id, dp_hours, dp_comments, dp_spent_on, Null)
                            If time_entry_id <> -1 Then
                                pCol_rm_spenttime_id.Cells(r, 1) = time_entry_id
                                pCol_force_update.Cells(r, 1) = ""
                                pCol_booked_ref.Cells(r, 1).NumberFormat = "@"
                                pCol_booked_ref.Cells(r, 1) = dp_hash
                                pCol_booked_time.Cells(r, 1) = Now()
                                Debug.Print ("Activity booked: " + dp_project + " - " + dp_task + " - " + str(time_entry_id))
                            Else
                                MsgBox ("Creation of Redmine Time Entry failed")
                            End If
                        End If
                        
                        ' wenn forced update, dann check sha difference und wenn gegeben, dann, wenn spent_time_id da ist
                        ' mach Update auf dem eintrag
                        If (Not IsEmpty(pCol_booked_ref.Cells(r, 1))) _
                            And (pCol_booked_ref.Cells(r, 1) <> dp_hash) _
                            And (Not IsEmpty(pCol_rm_spenttime_id.Cells(r, 1))) _
                            And (pCol_force_update.Cells(r, 1) = "x") Then
                            
                            ' Update the entry
                            time_entry_id = pCol_rm_spenttime_id.Cells(r, 1)
                            time_entry_id = redmine.Redmine_Time_Entry("PUT", rm_task_id, dp_hours, dp_comments, dp_spent_on, time_entry_id)
                            If time_entry_id <> -1 Then
                                pCol_rm_spenttime_id.Cells(r, 1) = time_entry_id
                                pCol_force_update.Cells(r, 1) = ""
                                pCol_booked_ref.Cells(r, 1) = dp_hash
                                pCol_booked_time.Cells(r, 1) = Now()
                                Debug.Print ("Updating " + dp_hash)
                            Else
                                MsgBox ("Update of Redmine Time Entry failed")
                            End If
                        End If
                        pCol_force_update.Cells(r, 1) = ""
                    End If
                End If
                r = r + 1
            Loop
        End If
        z = z + 1
    Loop
End Sub

Public Sub Insert_Cal_Entry(ByVal cal_event As Object, cur_date_rows() As Integer, cal_engine_id As String)
    Dim inserted As Boolean
    Dim Summary As String, iCalUID As String, htmllink As String
    Dim start_time As Date, end_time As Date, cal_date As Date
    
    inserted = False
    Summary = cal_event("summary")
    iCalUID = cal_event("iCalUID")
    htmllink = cal_event("htmlLink")
    start_time = cal_event("start_time")
    end_time = cal_event("end_time")
    cal_date = cal_event("cal_date")
    
    Debug.Print (CStr(start_time) + " - " + CStr(end_time))
    
    Dim i As Integer, r As Integer
    Dim dayTemplate As cDayTemplate
    Set dayTemplate = New cDayTemplate
    
    For i = 0 To UBound(cur_date_rows)
        r = cur_date_rows(i)
        If pCol_start.Cells(r, 1).value > TimeValue(start_time) And inserted = False Then
            pCol_start.rows(r).EntireRow.insert
            dayTemplate.Row_Dayplan_Item.EntireRow.Copy pCol_start.Cells(r, 1).EntireRow
            
            pCol_date.Cells(r, 1) = CDate(cal_date)
            pCol_activity.Cells(r, 1) = Summary
            pCol_fixed_start.Cells(r, 1) = start_time
            pCol_duration.Cells(r, 1) = end_time - start_time
            pCol_cal_indicator.Cells(r, 1) = cal_engine_id
            pCol_cal_id.Cells(r, 1) = iCalUID
            pWs.Hyperlinks.add Anchor:=pCol_cal_link.Cells(r, 1), Address:=htmllink
                                                           
            inserted = True
            Exit For
        End If
        
        If i = UBound(cur_date_rows) And inserted = False Then
            pCol_start.Cells(r + 1, 1).EntireRow.insert Shift:=xlShiftUp
            dayTemplate.Row_Dayplan_Item.EntireRow.Copy pCol_start.Cells(r + 1, 1).EntireRow
            
            pCol_date.Cells(r + 1, 1) = CDate(cal_date)
            pCol_activity.Cells(r + 1, 1) = Summary
            pCol_fixed_start.Cells(r + 1, 1) = start_time
            pCol_duration.Cells(r + 1, 1) = end_time - start_time
            pCol_cal_indicator.Cells(r + 1, 1) = cal_engine_id
            pCol_cal_id.Cells(r + 1, 1) = iCalUID
            pWs.Hyperlinks.add Anchor:=pCol_cal_link.Cells(r + 1, 1), Address:=htmllink
        End If
    Next i
    
End Sub

Public Function Update_Cal_Entry(ByVal cal_event As Object, cur_date_rows() As Integer, cal_engine_id As String) As Boolean
    ' check if the entry exists
    Dim iCalUID As String, i As Integer, r As Integer
    Dim inserted As Boolean, k As Integer, m As Integer
    Dim Summary As String, start_time As Date, end_time As Date
    
    iCalUID = cal_event("iCalUID")
    
    For i = 0 To UBound(cur_date_rows)
        r = cur_date_rows(i)
        If pCol_cal_id.Cells(r, 1) = iCalUID Then
            Update_Cal_Entry = True ' even if we do not need to update we indicate that the entry is already existing
                        
            Summary = cal_event("summary")
            start_time = cal_event("start_time")
            end_time = cal_event("end_time")
            
            ' actually we do not update the Summary
            If pCol_duration.Cells(r, 1) <> end_time - start_time Then
                pCol_duration.Cells(r, 1) = end_time - start_time
            End If
                        
            If pCol_fixed_start.Cells(r, 1) <> start_time Then
                inserted = False
                For k = 0 To UBound(cur_date_rows)
                    m = cur_date_rows(k)

                    If pCol_start.Cells(m, 1).value > TimeValue(start_time) And inserted = False Then
                        pWs.rows(m).EntireRow.insert
                        If m > r Then ' moved to a later start_time
                            pCol_fixed_start.Cells(r, 1) = start_time
                            pWs.rows(r).EntireRow.Copy pWs.rows(m).EntireRow
                            pWs.rows(r).EntireRow.Delete
                        Else ' moved to an earlier start_time
                            pCol_fixed_start.Cells(r + 1, 1) = start_time
                            pWs.rows(r + 1).EntireRow.Copy pWs.rows(m).EntireRow
                            pWs.rows(r + 1).EntireRow.Delete
                        End If
                        inserted = True
                    End If
                    If k = UBound(cur_date_rows) And inserted = False Then
                        pCol_fixed_start.Cells(r, 1) = start_time
                        pWs.rows(m + 1).EntireRow.insert Shift:=xlShiftUp
                        pWs.rows(r).EntireRow.Copy pWs.rows(m + 1).EntireRow
                        pWs.rows(r).EntireRow.Delete
                        inserted = True
                    End If
                Next k
            End If
            Exit Function
        End If
    Next i
    Update_Cal_Entry = False
End Function

Public Sub Remove_Cal_Entries(iCalUIDs() As String, cur_date_rows() As Integer, cal_engine_id As String)
    
    Dim del_list() As Integer
    Dim dl_idx As Integer, k As Integer, r As Integer, i As Integer
    Dim icaluid_found As Boolean
    
    dl_idx = 0
    For k = 0 To UBound(cur_date_rows)
        r = cur_date_rows(k)
        If pCol_cal_id.Cells(r, 1) <> "" Then
            icaluid_found = False
            If Not Not iCalUIDs Then
                For i = 0 To UBound(iCalUIDs)
                    If pCol_cal_id.Cells(r, 1) = iCalUIDs(i) Then
                        icaluid_found = True
                    End If
                    
                Next i
            End If
            ' if the entry comes from anouther calendar we ignore it
            If pCol_cal_indicator.Cells(r, 1) <> "" And pCol_cal_indicator.Cells(r, 1) <> cal_engine_id Then
                icaluid_found = True
            End If
            If icaluid_found = False Then
                ReDim Preserve del_list(dl_idx)
                del_list(dl_idx) = r
                dl_idx = dl_idx + 1
            End If
        End If
    Next k
    
    Dim x As Integer
    If Not Not del_list Then
        For x = UBound(del_list) To 0 Step -1
            pWs.rows(del_list(x)).EntireRow.Delete
        Next x
    End If

End Sub
'--------------------------------------------------------------------
' function returns all rows that belong to a certain date in an array
'--------------------------------------------------------------------
Public Function Get_Rows_For_Date(cal_date As String) As Integer()
    Dim cur_date_rows() As Integer, r As Integer, i As Integer
    r = 3
    i = 0
    Do Until pCol_date.Cells(r, 1) = end_of_table
        If pCol_date.Cells(r, 1).value = DateValue(cal_date) Then
            ReDim Preserve cur_date_rows(i)
            cur_date_rows(i) = r
            i = i + 1
        End If
        r = r + 1
    Loop
    Get_Rows_For_Date = cur_date_rows
End Function

Private Sub Class_Initialize()
    Set pWs = Worksheets("DailyPlan")
    
    Set pCol_week = pWs.Range("wsr_dailyplan_week")
    Set pCol_date = pWs.Range("wsr_dailyplan_date")
    Set pCol_project = pWs.Range("wsr_dailyplan_project")
    Set pCol_task = pWs.Range("wsr_dailyplan_task")
    Set pCol_fixed_start = pWs.Range("wsr_dailyplan_fixed_start")
    Set pCol_start = pWs.Range("wsr_dailyplan_start")
    Set pCol_end = pWs.Range("wsr_dailyplan_end")
    Set pCol_followup = pWs.Range("wsr_dailyplan_followup")
    Set pCol_done = pWs.Range("wsr_dailyplan_done")
    Set pCol_duration = pWs.Range("wsr_dailyplan_duration")
    Set pCol_activity_type = pWs.Range("wsr_dailyplan_activity_type")
    Set pCol_activity = pWs.Range("wsr_dailyplan_activity")
    Set pCol_rm_spenttime_id = pWs.Range("wsr_dailyplan_rm_spenttime_id")
    Set pCol_force_update = pWs.Range("wsr_dailyplan_force_update")
    Set pCol_booked_ref = pWs.Range("wsr_dailyplan_booked_ref")
    Set pCol_booked_time = pWs.Range("wsr_dailyplan_booked_time")
    Set pCol_cal_indicator = pWs.Range("wsr_dailyplan_cal_indicator")
    Set pCol_cal_id = pWs.Range("wsr_dailyplan_cal_id")
    Set pCol_cal_link = pWs.Range("wsr_dailyplan_cal_link")
    
    end_of_table = "XXXXXXXXXXXXXXX"
End Sub