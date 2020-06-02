Option Explicit

Private pCalEngineId As String

Public Sub Get_Events(cal_date As String)
On Error GoTo ErrHand:
    
    Application.ScreenUpdating = False
    
    Dim olApp As Object
    Dim olNS As Object
    Dim olFolder As Object
    Dim olApt As Object
    
    'Dim cal_date As String: cal_date = "30/04/2020"
    Dim cal_event As Dictionary: Set cal_event = New Dictionary
    Dim dp As cDailyPlan: Set dp = New cDailyPlan

    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If Err.Number > 0 Then Set olApp = CreateObject("Outlook.Application")
    If olApp Is Nothing Then
        GoTo ErrHand
    End If

    
    Dim myAppointments As Object 'Outlook.Items
    Dim currentAppointment As Object 'Outlook.AppointmentItem

    Set olNS = olApp.GetNamespace("MAPI")
    
    Set myAppointments = olNS.GetDefaultFolder(9).Items
    myAppointments.Sort "[Start]"
    
    myAppointments.IncludeRecurrences = True

    'get list with all rows of current date
    Dim cur_date_rows() As Integer
    cur_date_rows = dp.Get_Rows_For_Date(cal_date)
    
    Dim iCalUIDs() As String
    Dim start_date As Date, end_date As Date
    Dim i As Integer
    i = 0
    
    Dim tdyend As Date, tdystart As Date
    Dim cur_date As Date
        
    cur_date = DateValue(cal_date)
    tdystart = Format(cur_date, "Short Date")
    tdyend = Format(cur_date + 1, "Short Date")
            
    Set currentAppointment = myAppointments.find("[Start] >= """ & tdystart & """ and [Start] <= """ & tdyend & """")
    While TypeName(currentAppointment) <> "Nothing"
        start_date = DateValue(currentAppointment.Start)
        end_date = DateValue(currentAppointment.End)
        
        If currentAppointment.AllDayEvent Then ' necessary because outlook gives the end time of 00:00:00 from the next day
            end_date = DateAdd("d", -1, end_date)
        End If
        
        cal_event("summary") = currentAppointment.Subject
        cal_event("iCalUID") = currentAppointment.GlobalAppointmentID
        'cal_event ("htmlLink")
        
        If start_date < cur_date Then
            cal_event("start_time") = TimeValue("00:00:00")
        Else
            cal_event("start_time") = TimeValue(currentAppointment.Start)
        End If
        
        If end_date > cur_date Then
            cal_event("end_time") = TimeValue("23:59:59")
        Else
            cal_event("end_time") = TimeValue(currentAppointment.End)
        End If
                    
        cal_event("cal_date") = cal_date
        ' check if cal_entry already exists
        If Not dp.Update_Cal_Entry(cal_event, cur_date_rows, pCalEngineId) Then
            Call dp.Insert_Cal_Entry(cal_event, cur_date_rows, pCalEngineId)
            ' after inserting a new entry the list of rows with entries of the current date needs to be expanded
            ReDim Preserve cur_date_rows(UBound(cur_date_rows) + 1)
            cur_date_rows(UBound(cur_date_rows)) = cur_date_rows(UBound(cur_date_rows) - 1) + 1
        End If
        
        ' remember all iCalUID in array for removal procedure
        ReDim Preserve iCalUIDs(i)
        iCalUIDs(i) = cal_event("iCalUID")
        i = i + 1
        
        Set currentAppointment = myAppointments.FindNext
    Wend
 
    Call dp.Remove_Cal_Entries(iCalUIDs, cur_date_rows, pCalEngineId)
    dp.Conditional_Formating

    Set olApt = Nothing
    Set olFolder = Nothing
    Set olNS = Nothing
    Set olApp = Nothing
    
cleanExit:
    Application.ScreenUpdating = True
    Exit Sub

ErrHand:
    'Add error handler
    Resume cleanExit
End Sub

Private Sub Class_Initialize()
    pCalEngineId = "Ol"
End Sub