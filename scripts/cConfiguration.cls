Option Explicit

Private pWs As Worksheet
Public test As String

Public Property Get Ws() As Worksheet
    Set Ws = pWs
End Property

Public Property Get Projects() As Range
    Set Projects = pWs.Range("projects")
End Property

Public Property Get Redmine_Base_Url()
    Redmine_Base_Url = pWs.Cells(2, 8)
End Property

Public Property Get Redmine_Api_Key()
    Redmine_Api_Key = pWs.Cells(3, 8)
End Property

Public Property Get Module_Google_Cal_Enabled()
    Module_Google_Cal_Enabled = Module_Check_Enabled("GoogleCal")
End Property

Public Property Get Module_Outlook_Enabled()
    Module_Outlook_Enabled = Module_Check_Enabled("Outlook")
End Property

Public Property Get Module_Day_Templates_Enabled()
    Module_Day_Templates_Enabled = Module_Check_Enabled("DayTemplates")
End Property

Public Property Get Module_Redmine_Enabled()
    Module_Redmine_Enabled = Module_Check_Enabled("Redmine")
End Property

Public Property Get Module_Todo_Enabled()
    Module_Todo_Enabled = Module_Check_Enabled("ToDo")
End Property

Public Property Get Module_Base_Enabled()
    Module_Base_Enabled = Module_Check_Enabled("Base")
End Property

Public Property Get Module_Analytics_Enabled()
    Module_Analytics_Enabled = Module_Check_Enabled("Analytics")
End Property


Private Function Module_Check_Enabled(mod_name As String)
    Dim i As Integer, k As Integer
    i = 7
    Do Until pWs.Cells(i, 7) = ""
        If pWs.Cells(i, 7) = mod_name Then
            If pWs.Cells(i, 8) = "x" Then
                Module_Check_Enabled = True
            Else
                Module_Check_Enabled = False
            End If
            Exit Function
        End If
        i = i + 1
    Loop
End Function

Public Property Get Selected_Google_Calendars() As String()
    Dim cal_ids() As String
    Dim i As Integer, k As Integer
    i = 22
    k = 0
    Do Until pWs.Cells(i, 8) = ""
        If pWs.Cells(i, 7) = "x" Then
            ReDim Preserve cal_ids(k)
            cal_ids(k) = pWs.Cells(i, 9)
            k = k + 1
        End If
        i = i + 1
    Loop
    
    Selected_Google_Calendars = cal_ids
End Property

Public Sub Receive_Google_Calendars_List()
    Dim gCal As cGoogleCalendar
    Set gCal = New cGoogleCalendar
    Dim calendars() As String
    
    Dim i As Integer, j As Integer
    ' remove old entries
    i = 22
    Do Until pWs.Cells(i, 8) = ""
        i = i + 1
    Loop
    pWs.Range(pWs.Cells(22, 8), pWs.Cells(i, 9)).value = ""
    
    calendars = gCal.Get_Calendars
    For i = 0 To UBound(calendars, 2)
        For j = 0 To UBound(calendars, 1)
            pWs.Cells(22 + i, 8 + j).value = calendars(j, i)
        Next j
    Next i
End Sub

Public Sub Update_Lists()
    Dim z As Long
    Dim range_name As String
    Dim i As Integer
          
    With pWs
        i = 1
        For i = 1 To 5
            If i <> 3 Then
                range_name = Cells(2, i)
                z = 4
                Do Until .Cells(z, i) = ""
                    z = z + 1
                Loop
                .Range(.Cells(4, i), .Cells(z - 1, i)).name = range_name
            End If
        Next
    End With
End Sub

Public Sub Module_Selection_Changed(ByVal Target As Range)
    If Target.column = 8 And Target.row >= 7 And Target.row <= 17 Then
        Dim enabled As Boolean
        
        If Target.value = "x" Then
            enabled = True
        Else
            enabled = False
        End If
                
        Dim mod_name As String
        mod_name = ActiveSheet.Cells(Target.row, 7)
                        
        Dim i As Integer
        i = 0
        
        If mod_name = "Analytics" _
           Or mod_name = "Reminder" _
           Or mod_name = "WeekPlan" _
           Or mod_name = "MonthPlan" _
           Or mod_name = "WeekWorkPlan" Then
            Worksheets(mod_name).visible = enabled
        End If
        
        If mod_name = "Redmine" Then
            Worksheets("RedmineTasks").visible = enabled
            Worksheets("RedmineConfig").visible = enabled
            ' switch the contextmenu entries
        End If
        
        If mod_name = "DayTemplates" Then
            Worksheets(mod_name).visible = enabled
            ' switch the contextmenu entries
        End If
        
        If mod_name = "ToDo" Then
            Worksheets(mod_name).visible = enabled
            ' switch the contextmenu entries
        End If
        
        If mod_name = "Outlook" Then
            ' switch the contextmenu entries
        End If
        
        If mod_name = "GoogleCal" Then
            ' switch the contextmenu entries
        End If
                
    End If
End Sub

Private Sub Class_Initialize()
    Set pWs = Sheets("Configuration")
End Sub
