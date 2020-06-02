Option Explicit

Private pWs As Worksheet

Private pBase_Url As String
Private pApi_Key As String

Private pCol_issue_id As Range
Private pCol_project_id As Range
Private pCol_project_name As Range
Private pCol_version_id As Range
Private pCol_version_name As Range
Private pCol_assigned_to_id As Range
Private pCol_assigned_to_name As Range
Private pCol_subject As Range
Private pCol_due_date As Range
Private pCol_estimated_hours As Range
Private pCol_priority_id As Range
Private pCol_priority_name As Range
Private pCol_hash_key As Range

Sub Redmine_Add_To_Task(from_row As Integer)
    Dim to_row As Integer
    Dim task As cTasks
    Set task = New cTasks
    
    Application.EnableEvents = False
    
    ' getting the redmine task
    ' issue_id and project_name
    Dim issue_id As String, rm_project_name As String, _
        rm_task_description As String
    
    issue_id = pCol_issue_id.Cells(from_row, 1)
    rm_project_name = pCol_project_name.Cells(from_row, 1)
    rm_task_description = pCol_subject.Cells(from_row, 1)
    
    Dim project_owner As String, project_name As String
    
    ' find rm_project_name in RedmineConfig
    Dim i As Integer
    i = 3
    Do Until Worksheets("RedmineConfig").Cells(i, 3) = ""
        If Worksheets("RedmineConfig").Cells(i, 3) = rm_project_name Then
            project_owner = Worksheets("RedmineConfig").Cells(i, 4)
            project_name = Worksheets("RedmineConfig").Cells(i, 5)
        End If
        i = i + 1
    Loop
    
    ' determine to_row on base of the rm_project_name
    to_row = task.Get_Insert_Row_For_Project(project_owner, project_name)
    
    ' insert as last element in the group
    task.Ws.rows(to_row).EntireRow.insert
    
    task.Col_Rm_Taskid.Cells(to_row, 1) = issue_id
    task.Col_Owner.Cells(to_row, 1) = project_owner
    task.Col_Project.Cells(to_row, 1) = project_name
    task.Col_Taskdescription.Cells(to_row, 1) = rm_task_description
        
    ' finally update the range names
    task.Update_Task_Ranges
    
    Application.EnableEvents = True
    
End Sub

Public Function Redmine_Time_Entry(method, rm_task_id, rm_hours, rm_comments, rm_spent_on, rm_time_entry_id) As Integer
    Dim xmlhttp As New MSXML2.XMLHTTP60, myurl As String
    Dim json_string As String

    If method = "POST" Then
        myurl = pBase_Url + "/time_entries.json"
    ElseIf method = "PUT" Then
        myurl = pBase_Url + "/time_entries/" + CStr(rm_time_entry_id) + ".json"
    End If
     
    xmlhttp.Open method, myurl, False
    xmlhttp.SetRequestHeader "X-Redmine-API-Key", pApi_Key
    xmlhttp.SetRequestHeader "Content-Type", "application/json"
    
    json_string = "{ ""time_entry"" : { ""issue_id"": " + str(rm_task_id) + ", ""hours"": " + Replace(FormatNumber(rm_hours, 2), ",", ".") + ", ""comments"": """ + rm_comments + """, ""spent_on"" : """ + rm_spent_on + """, ""activity"" : 9 } }"
    
    Debug.Print (json_string)
    xmlhttp.Send json_string
        
    Debug.Print (xmlhttp.ResponseText)
    Debug.Print (xmlhttp.status)
    
    If method = "POST" And xmlhttp.status = 201 Then
        Dim parsed As Object
        Set parsed = JsonConverter.ParseJson(xmlhttp.ResponseText)
        Redmine_Time_Entry = parsed("time_entry")("id")
    ElseIf method = "PUT" And xmlhttp.status = 200 Then
        Redmine_Time_Entry = rm_time_entry_id
    Else
        Redmine_Time_Entry = -1
    End If

End Function

Private Sub Update_RM_Projects_List()
    Dim i As Integer
    
    Dim d As Scripting.Dictionary
    Set d = New Scripting.Dictionary
    For i = 3 To Get_End_Of_Table() - 1
        d.item(pCol_project_name.Cells(i, 1).value) = 1
    Next i
    
    Dim end_row As Integer
    end_row = d.Count + 3 - 1

    Worksheets("RedmineConfig").Range("A3:A" + CStr(end_row)).value = WorksheetFunction.Transpose(d.Keys)
    Worksheets("RedmineConfig").Range("A3:A" + CStr(end_row)).name = "redmine_project_list"
        
End Sub

Public Sub Read_Redmine_Projects()
On Error GoTo ErrHand:
           
    Dim xmlhttp As New MSXML2.XMLHTTP60, myurl As String
    Dim response As String
    
    myurl = pBase_Url + "/issues.json?assigned_to_id=me&status_id=open&limit=100&sort=created_on:desc"
    ' TODO: make second request with watcher_id=me instead of assigned_to_id=me
     
    xmlhttp.Open "GET", myurl, False
    xmlhttp.SetRequestHeader "X-Redmine-API-Key", pApi_Key
    xmlhttp.SetRequestHeader "Content-Type", "application/json"
    xmlhttp.Send
    response = xmlhttp.ResponseText
    Debug.Print response
    
    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(response)
    Dim issues As Object
    Set issues = parsed("issues")
       
    'now getting also the issues where I'm watcher - these are issues where I'm involved
    'and where I want to also book time on
    myurl = pBase_Url + "/issues.json?watcher_id=me&status_id=open&limit=100&sort=created_on:asc"
    xmlhttp.Open "GET", myurl, False
    xmlhttp.SetRequestHeader "X-Redmine-API-Key", pApi_Key
    xmlhttp.SetRequestHeader "Content-Type", "application/json"
    xmlhttp.Send
    response = xmlhttp.ResponseText
    Set parsed = JsonConverter.ParseJson(response)
    Dim issues_watched As Object, issue_watched As Object
    Set issues_watched = parsed("issues")
    
    For Each issue_watched In issues_watched
        issues.add issue_watched
    Next issue_watched
    
    ' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    ' TODO Write all records to a worksheet from an array, this is much faster
    ' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    'Dim ws          As Worksheet: Set ws = ThisWorkbook.Sheets("Sheet1")
    'Ws.Range("A2:D" & NextRow + 1).value = WorksheetFunction.Transpose(myArr)
    
    pWs.Range(pWs.rows(3), pWs.rows(Get_End_Of_Table())) = ""
    
    Application.ScreenUpdating = False

    pCol_issue_id.NumberFormat = "General"
    pCol_project_id.NumberFormat = "General"
    pCol_version_id.NumberFormat = "General"
    pCol_assigned_to_id.NumberFormat = "General"
    pCol_estimated_hours.NumberFormat = "General"
    pCol_priority_id.NumberFormat = "General"
    pCol_hash_key.NumberFormat = "@"
    
    Dim myArr() As Variant: ReDim myArr(0 To 12, 0 To issues.Count - 1)
    
    Dim i As Integer
    i = 3
    Dim issue As Object
    For Each issue In issues
        myArr(pCol_issue_id.column - 1, i - 3) = issue("id")
        myArr(pCol_project_id.column - 1, i - 3) = issue("project")("id")
        myArr(pCol_project_name.column - 1, i - 3) = issue("project")("name")
        If issue.exists("fixed_version") Then
            myArr(pCol_version_id.column - 1, i - 3) = issue("fixed_version")("id")
            myArr(pCol_version_name.column - 1, i - 3) = issue("fixed_version")("name")
        End If
        myArr(pCol_assigned_to_id.column - 1, i - 3) = issue("assigned_to")("id")
        myArr(pCol_assigned_to_name.column - 1, i - 3) = issue("assigned_to")("name")
        myArr(pCol_subject.column - 1, i - 3) = issue("subject")
        myArr(pCol_due_date.column - 1, i - 3) = issue("due_date")
        myArr(pCol_estimated_hours.column - 1, i - 3) = issue("estimated_hours")
        myArr(pCol_priority_id.column - 1, i - 3) = issue("priority")("id")
        myArr(pCol_priority_name.column - 1, i - 3) = issue("priority")("name")
            
        Dim concat As String, k As Integer
        For k = 0 To UBound(myArr, 1)
            concat = concat & myArr(k, i - 3)
        Next k
        myArr(pCol_hash_key.column - 1, i - 3) = SHA1TRUNC(concat) 'TODO bug SHA1TRUNC for UTF8
        i = i + 1
    Next issue
    pWs.Range(pWs.Cells(3, 1), pCol_hash_key.Cells(i - 1, 1)).value = WorksheetFunction.Transpose(myArr)
    pWs.Range(pWs.Cells(3, 1), pCol_hash_key.Cells(i - 1, 1)).name = "redmine_issues"
    
    Update_RM_Projects_List

cleanExit:
    Application.ScreenUpdating = True
    Exit Sub

ErrHand:
    MsgBox "The following error occurred: " & Err.Description
    Resume cleanExit
End Sub

Private Function Get_End_Of_Table() As Integer
    Dim i As Integer
    i = 3
    Do Until pCol_issue_id.Cells(i, 1) = ""
        i = i + 1
    Loop
    Get_End_Of_Table = i
End Function

Private Sub Class_Initialize()
    Set pWs = Worksheets("RedmineTasks")
    
    Dim configuration As cConfiguration
    Set configuration = New cConfiguration
    pBase_Url = configuration.Redmine_Base_Url
    pApi_Key = configuration.Redmine_Api_Key
    
    Set pCol_issue_id = pWs.Range("wsr_redminetasks_issue_id")
    Set pCol_project_id = pWs.Range("wsr_redminetasks_project_id")
    Set pCol_project_name = pWs.Range("wsr_redminetasks_project_name")
    Set pCol_version_id = pWs.Range("wsr_redminetasks_version_id")
    Set pCol_version_name = pWs.Range("wsr_redminetasks_version_name")
    Set pCol_assigned_to_id = pWs.Range("wsr_redminetasks_assigned_to_id")
    Set pCol_assigned_to_name = pWs.Range("wsr_redminetasks_assigned_to_name")
    Set pCol_subject = pWs.Range("wsr_redminetasks_subject")
    Set pCol_due_date = pWs.Range("wsr_redminetasks_due_date")
    Set pCol_estimated_hours = pWs.Range("wsr_redminetasks_estimated_hours")
    Set pCol_priority_id = pWs.Range("wsr_redminetasks_priority_id")
    Set pCol_priority_name = pWs.Range("wsr_redminetasks_priority_name")
    Set pCol_hash_key = pWs.Range("wsr_redminetasks_hash_key")
    
End Sub
