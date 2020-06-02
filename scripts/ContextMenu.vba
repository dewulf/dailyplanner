Option Explicit
 
Dim ctrl_caption As String
 
' Nice resource about context menu
' https://www.rondebruin.nl/win/s6/win002.htm
' FacedIDs
' https://bettersolutions.com/vba/ribbon/face-ids-2003.htm

 
Private Sub BuildCustomMenu()
     
    Dim ctrl As CommandBarControl
    Dim btn As CommandBarControl, btn2 As CommandBarControl, btn3 As CommandBarControl
    Dim i As Integer
    
    Dim conf As cConfiguration
    Set conf = New cConfiguration
     
    If ActiveSheet.name = "DailyPlan" Then
        Set ctrl = Application.CommandBars("Cell").Controls.add _
            (Type:=msoControlPopup, Before:=1)
        ctrl.Caption = "DailyPlan..."
                                
        If conf.Module_Base_Enabled Then
            With ctrl.Controls.add
                .Caption = "Add New Activity below"
                .FaceId = 137
                .tag = ActiveCell.row
                .OnAction = "Insert_New_Dailyplan_Entry"
            End With
            
            With ctrl.Controls.add
                .Caption = "Move Activity to Tomorrow"
                .FaceId = 156
                .tag = ActiveCell.row
                .OnAction = "Move_Activity_To_Tomorrow"
            End With
        End If
                        
        If conf.Module_Day_Templates_Enabled Then
            With ctrl.Controls.add
                .BeginGroup = True
                .FaceId = 237
                .Caption = "Insert New Day"
                .tag = ActiveCell.row
                .OnAction = "Insert_Day_Template"
            End With
        End If
                                
        If conf.Module_Todo_Enabled Then
            With ctrl.Controls.add
                .BeginGroup = True
                .FaceId = 136
                .Caption = "Move Activity to ToDo - Next Days"
                .tag = ActiveCell.row
                .OnAction = "Move_Activity_To_Todo_Nextdays"
            End With
    
            With ctrl.Controls.add
                .Caption = "Copy Activity to ToDo - Follow Ups"
                .FaceId = 37
                .tag = ActiveCell.row
                .OnAction = "Copy_Activity_To_Todo_Followups"
            End With
        End If
        
        If conf.Module_Google_Cal_Enabled Then
            With ctrl.Controls.add
                .BeginGroup = True
                .FaceId = 1106
                .Caption = "Get Google Calendar Events"
                .tag = ActiveSheet.Range("wsr_dailyplan_date").Cells(ActiveCell.row, 1) 'we'll use the tag property to hold a value
                .OnAction = "Get_Google_Calendar_Events"
            End With
        End If
                
        If conf.Module_Outlook_Enabled Then
            With ctrl.Controls.add
                .BeginGroup = True
                .FaceId = 6225
                .Caption = "Get Outlook Calendar Events"
                .tag = ActiveSheet.Range("wsr_dailyplan_date").Cells(ActiveCell.row, 1) 'we'll use the tag property to hold a value
                .OnAction = "Get_Outlook_Calendar_Events"
            End With
        End If
        
        If conf.Module_Redmine_Enabled Then
            With ctrl.Controls.add
                .BeginGroup = True
                .FaceId = 9935
                .Caption = "Book Spent Time to Redmine"
                .tag = ActiveSheet.Range("wsr_dailyplan_date").Cells(ActiveCell.row, 1) 'we'll use the tag property to hold a value
                .OnAction = "Book_Spent_Time_To_Redmine" 'the routine called by the control
            End With
        End If

        If conf.Module_Analytics_Enabled Then
            With ctrl.Controls.add
                .BeginGroup = True
                .FaceId = 427
                .Caption = "Do Analytics for the selected week"
                .tag = ActiveCell.row
                .OnAction = "Do_Analytics_For_Week" 'the routine called by the control
            End With
        End If
    End If
    
    If ActiveSheet.name = "RedmineTasks" Then
        Set ctrl = Application.CommandBars("Cell").Controls.add _
            (Type:=msoControlPopup, Before:=1)
        ctrl.Caption = "DailyPlan..."
        With ctrl.Controls.add
            .Caption = "Create new Task in Tasks"
            .FaceId = 9960
            .tag = ActiveCell.row 'ActiveSheet.Range("wsr_redminetasks_issue_id").Cells(ActiveCell.row, 1) 'we'll use the tag property to hand over the issue_id
            .OnAction = "Redmine_Add_To_Task" 'the routine called by the control
        End With
    End If
    
    If ActiveSheet.name = "ToDo" Then
        Set ctrl = Application.CommandBars("Cell").Controls.add _
            (Type:=msoControlPopup, Before:=1)
        ctrl.Caption = "DailyPlan..."
        
        With ctrl.Controls.add
            .FaceId = 137
            .Caption = "Add New Todo Entry below"
            .tag = ActiveCell.row
            .OnAction = "Insert_New_Todo_Entry"
        End With
        
        With ctrl.Controls.add
            .BeginGroup = True
            .Caption = "Add to today"
            .FaceId = 5473
            .tag = ActiveCell.row
            .OnAction = "Todo_Add_To_Today"
        End With
        
        With ctrl.Controls.add
            .Caption = "Add to tomorrow"
            .FaceId = 350
            .tag = ActiveCell.row
            .OnAction = "Todo_Add_To_Tomorrow"
        End With
        
    End If
    
    If ActiveSheet.name = "Configuration" Then
        Set ctrl = Application.CommandBars("Cell").Controls.add _
            (Type:=msoControlPopup, Before:=1)
        ctrl.Caption = "DailyPlan..."
        
        With ctrl.Controls.add
            .Caption = "Receive Google Calendars"
            .tag = ActiveCell.row
            .OnAction = "Receive_Google_Calendars_List"
        End With
        
    End If
End Sub
 
Private Sub DeleteCustomMenu()
    Dim ctrl As CommandBarControl
     'go thru all the cell commandbar controls and delete our menu item
    For Each ctrl In Application.CommandBars("Cell").Controls
        If ctrl.Caption = "DailyPlan..." Then
            ctrl.Delete
        ElseIf ctrl.Caption = "Insert Shape..." Then
            ctrl.Delete
        End If
    Next
End Sub




