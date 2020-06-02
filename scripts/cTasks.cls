Option Explicit

Private pWs As Worksheet
Private pCol_project As Range
Private pCol_taskdescription As Range
Private pCol_rm_taskid As Range
Private pCol_owner As Range

Public Property Get Ws()
    Set Ws = pWs
End Property

Public Property Get Col_Project()
    Set Col_Project = pCol_project
End Property

Public Property Get Col_Owner()
    Set Col_Owner = pCol_owner
End Property

Public Property Get Col_Rm_Taskid()
    Set Col_Rm_Taskid = pCol_rm_taskid
End Property

Public Property Get Col_Taskdescription()
    Set Col_Taskdescription = pCol_taskdescription
End Property

Public Function Get_Insert_Row_For_Project(project_owner As String, project_name As String)
    Dim i As Integer
    i = 3
    Do Until pCol_owner.Cells(i, 1) = project_owner _
        And pCol_project.Cells(i, 1) = project_name
        i = i + 1
    Loop
    Get_Insert_Row_For_Project = i
End Function

Public Sub Update_Task_Ranges()
    Dim z As Integer
    Dim range_name As String
    Dim First As Integer
          
    z = 3
    First = z
    range_name = pCol_project.Cells(z, 1)
    Do Until pCol_project.Cells(z, 1) = ""
        If pCol_project.Cells(z, 1) <> range_name Then
            pWs.Range(pCol_taskdescription.Cells(First, 1), pCol_taskdescription.Cells(z - 1, 1)).name = range_name
            range_name = pCol_project.Cells(z, 1)
            First = z
        End If
        z = z + 1
    Loop
    pWs.Range(pCol_taskdescription.Cells(First, 1), pCol_taskdescription.Cells(z - 1, 1)).name = range_name
End Sub

Public Sub Add_To_Todo()
    'TODO
End Sub

Public Sub Add_To_Tomorrow()
    'TODO
End Sub

Public Sub Add_To_Today()
    'TODO
End Sub

Public Sub Add_New()
    'TODO
End Sub

Public Sub Duplicate()
    'TODO
End Sub

Private Sub Class_Initialize()
    Set pWs = Worksheets("Tasks")
    
    Set pCol_project = pWs.Range("wsr_tasks_project")
    Set pCol_taskdescription = pWs.Range("wsr_tasks_taskdescription")
    Set pCol_rm_taskid = pWs.Range("wsr_tasks_rm_taskid")
    Set pCol_owner = pWs.Range("wsr_tasks_owner")
End Sub

