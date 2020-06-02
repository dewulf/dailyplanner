Option Explicit

Sub Conditional_Formating()
    Dim dailyPlan As cDailyPlan
    Set dailyPlan = New cDailyPlan
    dailyPlan.Conditional_Formating
End Sub

Sub Insert_New_Dailyplan_Entry()
    Dim dailyPlan As cDailyPlan
    Set dailyPlan = New cDailyPlan
    Dim row As Integer
    row = Application.CommandBars.ActionControl.tag
    dailyPlan.Insert_New_Entry row
End Sub

Sub Move_Activity_To_Tomorrow()
    Dim dailyPlan As cDailyPlan
    Set dailyPlan = New cDailyPlan
    Dim row As Integer
    row = Application.CommandBars.ActionControl.tag
    dailyPlan.Move_To_Tomorrow (row)
End Sub

Sub Move_Activity_To_Todo_Nextdays()
    Dim dailyPlan As cDailyPlan
    Set dailyPlan = New cDailyPlan
    Dim row As Integer
    row = Application.CommandBars.ActionControl.tag
    dailyPlan.Move_To_Todo_Nextdays row
End Sub

Sub Copy_Activity_To_Todo_Followups()
    Dim dailyPlan As cDailyPlan
    Set dailyPlan = New cDailyPlan
    Dim row As Integer
    row = Application.CommandBars.ActionControl.tag
    dailyPlan.Copy_To_Todo_Followups row
End Sub

Sub Todo_Add_To_Tomorrow()
    Dim todo As cTodo
    Set todo = New cTodo
    Dim row As Integer
    row = Application.CommandBars.ActionControl.tag
    Call todo.Add_To_Dailyplan_Someday(DateValue(Now() + 1), row)
End Sub

Sub Todo_Add_To_Today()
    Dim todo As cTodo
    Set todo = New cTodo
    Dim row As Integer
    row = Application.CommandBars.ActionControl.tag
    Call todo.Add_To_Dailyplan_Someday(DateValue(Now()), row)
End Sub

Sub Configuration_Update_Lists()
    Dim configuration As cConfiguration
    Set configuration = New cConfiguration
    configuration.Update_Lists
End Sub

Sub Update_Task_Ranges()
    Dim task As cTasks
    Set task = New cTasks
    task.Update_Task_Ranges
End Sub

Sub Read_Redmine_Projects()
    Dim redmineTask As cRedmineTask
    Set redmineTask = New cRedmineTask
    redmineTask.Read_Redmine_Projects
End Sub

Sub Book_Spent_Time_To_Redmine()
    Dim dailyPlan As cDailyPlan
    Set dailyPlan = New cDailyPlan
    Dim book_date As String
    book_date = Application.CommandBars.ActionControl.tag
    Call dailyPlan.Book_Spent_Time_To_Redmine(book_date)
End Sub

Sub Redmine_Add_To_Task()
    Dim redmineTask As cRedmineTask
    Set redmineTask = New cRedmineTask
    Dim rm_issue_row As Integer
    rm_issue_row = Application.CommandBars.ActionControl.tag
    Call redmineTask.Redmine_Add_To_Task(rm_issue_row)
End Sub

Sub Insert_Day_Template()
    Dim dayTemplate As cDayTemplate
    Set dayTemplate = New cDayTemplate
    Dim day_before_row As Integer
    day_before_row = Application.CommandBars.ActionControl.tag
    Call dayTemplate.Insert_Day_Template(day_before_row)
End Sub

Sub Get_Google_Calendar_Events()
    Dim gCal As cGoogleCalendar
    Set gCal = New cGoogleCalendar
    Dim cur_date As String
    cur_date = Application.CommandBars.ActionControl.tag
    Call gCal.Get_Google_Calendar_Events(cur_date)
End Sub

Sub Get_Outlook_Calendar_Events()
    Dim Outlook As cOutlook
    Set Outlook = New cOutlook
    Dim cur_date As String
    cur_date = Application.CommandBars.ActionControl.tag
    Call Outlook.Get_Events(cur_date)
End Sub

Sub Receive_Google_Calendars_List()
    Dim conf As cConfiguration
    Set conf = New cConfiguration
    conf.Receive_Google_Calendars_List
End Sub

Sub Insert_New_Todo_Entry()
    Dim todo As cTodo
    Set todo = New cTodo
    Dim row As Integer
    row = Application.CommandBars.ActionControl.tag
    Call todo.Insert_New_Entry(row)
End Sub

Sub Do_Analytics_For_Week()
    Dim dp As cDailyPlan
    Set dp = New cDailyPlan
    Dim row As Integer
    row = Application.CommandBars.ActionControl.tag
    Call dp.Do_Analytics(row)
End Sub