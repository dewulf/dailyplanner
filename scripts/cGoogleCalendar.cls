Option Explicit

Private pClientId As String
Private pClientSecret As String
Private pSelectedGoogleCalendars() As String
Private pCalEngineId As String

Private Function Get_Auth_Header()
    Dim authHeader As String
    With getGoogled("calendar")
        If (.hasToken) Then
            authHeader = .authHeader
        Else
            With getGoogled("calendar", , pClientId, pClientSecret)
                If (.hasToken) Then
                    authHeader = .authHeader
                End If
            End With
        End If
        .tearDown
    End With
    Get_Auth_Header = authHeader
End Function

Sub Get_Google_Calendar_Events(cal_date As String)
On Error GoTo ErrHand:
    
    Application.ScreenUpdating = False
    
    Dim client_id As String, client_secret As String
    Dim authHeader As String, url As String, accessToken As String
    Dim cal_start As Date, cal_end As Date
    Dim parsed As Object, cal_id_events As Object
    Dim timeMin As String, timeMax As String
    
    cal_start = cal_date + " 00:00:00"
    cal_end = DateAdd("d", 1, CDate(cal_date))
    timeMin = Format(ConvertLocalToGMT(cal_start, True), "yyyy-mm-ddThh:mm:ssZ")
    timeMax = Format(ConvertLocalToGMT(cal_end, True), "yyyy-mm-ddThh:mm:ssZ")
    
    Dim cal_events() As Object, cal_id_ev As Object
        
    Dim cb As cBrowser
    Set cb = New cBrowser
    Dim cal_id As Variant
    Dim i As Integer: i = 0
    For Each cal_id In pSelectedGoogleCalendars
        Debug.Print cal_id
        url = "https://www.googleapis.com/calendar/v3/calendars/" + cal_id + "/events?singleEvents=true&timeMin=" + timeMin + "&timeMax=" + timeMax
        With cb
            Dim jsonData As String
            jsonData = .httpGET(url, , , , , Get_Auth_Header)
            Debug.Print jsonData
            Set parsed = JsonConverter.ParseJson(jsonData)
            Set cal_id_events = parsed("items")
        End With
        
        If cal_id_events.Count > 0 Then
            If IsArrayAllocated(cal_events) Then
                ReDim Preserve cal_events(UBound(cal_events) + cal_id_events.Count)
            Else
                ReDim cal_events(cal_id_events.Count - 1)
            End If
        End If
        
        For Each cal_id_ev In cal_id_events
            Set cal_events(i) = cal_id_ev
            i = i + 1
        Next cal_id_ev
    Next cal_id
    
    Dim dp As cDailyPlan
    Set dp = New cDailyPlan
    
    Dim cur_date_rows() As Integer
    cur_date_rows = dp.Get_Rows_For_Date(cal_date)
    
    Dim iCalUIDs() As String
    Dim start_date As Date, end_date As Date

    i = 0
    Dim cal_event As Variant
    If Not Not cal_events Then
        For Each cal_event In cal_events
            
            If cal_event("start").exists("dateTime") Then
                start_date = DateValue(get_google_time(CStr(cal_event("start")("dateTime"))))
                If start_date < DateValue(cal_date) Then
                    cal_event("start_time") = TimeValue("00:00:00")
                Else
                    cal_event("start_time") = TimeValue(get_google_time(CStr(cal_event("start")("dateTime"))))
                End If
            ElseIf cal_event("start").exists("date") Then
                cal_event("start_time") = TimeValue("00:00:00")
            End If
            
            If cal_event("end").exists("dateTime") Then
                end_date = DateValue(get_google_time(CStr(cal_event("end")("dateTime"))))
                If end_date > DateValue(cal_date) Then
                    cal_event("end_time") = TimeValue("23:59:59")
                Else
                    cal_event("end_time") = TimeValue(get_google_time(CStr(cal_event("end")("dateTime"))))
                End If
            ElseIf cal_event("end").exists("date") Then
                cal_event("end_time") = TimeValue("23:59:59")
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
        Next cal_event
    End If
    
    Call dp.Remove_Cal_Entries(iCalUIDs, cur_date_rows, pCalEngineId)
    dp.Conditional_Formating

cleanExit:
    Application.ScreenUpdating = True
    Exit Sub

ErrHand:
    MsgBox "The following error occurred: " & Err.Description
    Resume cleanExit
End Sub

Public Function Get_Calendars() As String()
    Dim cb As cBrowser: Set cb = New cBrowser
    Dim url As String: url = "https://www.googleapis.com/calendar/v3/users/me/calendarList"
    Dim calendars As Object
    With cb
        Dim jsonData As String
        jsonData = .httpGET(url, , , , , Get_Auth_Header)
        Debug.Print jsonData
        Set calendars = JsonConverter.ParseJson(jsonData)("items")
    End With
    
    Dim calArr() As String: ReDim calArr(0 To 1, 0 To calendars.Count - 1)
    Dim cal As Object, i As Integer: i = 0
    For Each cal In calendars
        calArr(1, i) = cal("id")
        calArr(0, i) = cal("summary")
        i = i + 1
    Next cal
   
    Get_Calendars = calArr
End Function

Private Function get_google_time(timestring As String) As Date
    timestring = Replace(timestring, "T", " ")
    timestring = Left(timestring, Len(timestring) - 6)
    get_google_time = CDate(timestring)
End Function

Private Function getGoogled(scope As String, _
                                Optional replacementpackage As cJobject = Nothing, _
                                Optional clientID As String = vbNullString, _
                                Optional clientSecret As String = vbNullString, _
                                Optional complain As Boolean = True, _
                                Optional cloneFromeScope As String = vbNullString) As cOauth2
    Dim o2 As cOauth2
    Set o2 = New cOauth2
    With o2.googleAuth(scope, replacementpackage, clientID, clientSecret, complain, cloneFromeScope)
        If Not .hasToken And complain Then
            MsgBox ("Failed to authorize to google for scope " & scope & ":denied code " & o2.denied)
        End If
    End With
    
    Set getGoogled = o2
End Function

Private Sub Class_Initialize()
    Dim conf As cConfiguration
    Set conf = New cConfiguration
    
    pClientId = "718388455410-m7qnab6j07dfgdss5gkldlr3b9trd1sp.apps.googleusercontent.com" 'conf.Google_Client_Id
    pClientSecret = "JjZMLIvqbq2Z8QOV6yS_eIkR" 'conf.Google_Client_Secret
    pSelectedGoogleCalendars = conf.Selected_Google_Calendars
    pCalEngineId = "Go"
End Sub