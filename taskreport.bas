Attribute VB_Name = "taskreport"
Sub Export_TaskItems()
    'Adapted from: https://www.youtube.com/watch?v=AUSftiL4GyU
    
    Dim NS As NameSpace
    Dim TaskFolder As Folder
    Dim Item As Object
    Dim TaskItem As TaskItem
    
    Dim xlApp As Excel.Application
    Dim xlwb As Excel.Workbook
    Dim xlws As Excel.Worksheet
    Dim iRow As Long
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlwb = xlApp.Workbooks.Add
    Set xlws = xlwb.Worksheets(1)
    
    Set NS = Session
    Set TaskFolder = NS.GetDefaultFolder(olFolderTasks)
    
    xlws.Range("A1").Resize(1, 7).Value = Array("Status", "Time (Min)", "Task Name", "Due Date", "Completed", "Importance", "Category")
    
    iRow = 2
    
    For Each Item In TaskFolder.Items
        If Item.Class = olTask Then
            
            Set TaskItem = Item
            
            'Skip old Tasks
            If TaskItem.Status = 2 Then If TaskItem.DateCompleted < Now - 7 Then GoTo nextitem
            
            xlws.Cells(iRow, 1).Value = Convert_Status(TaskItem.Status)
            xlws.Cells(iRow, 2).Value = TaskItem.ActualWork
            xlws.Cells(iRow, 3).Value = TaskItem.Subject
            xlws.Cells(iRow, 4).Value = TaskItem.DueDate
            xlws.Cells(iRow, 5).Value = Convert_DateComplete(TaskItem.DateCompleted)
            xlws.Cells(iRow, 6).Value = Convert_Importance(TaskItem.Importance)
            xlws.Cells(iRow, 7).Value = TaskItem.Categories
            
            iRow = iRow + 1
nextitem:
        End If
    Next Item
    
    xlws.Range("A1:G1").AutoFilter
    
    'TODO: Finish formating export
    xlws.Columns("A:H").AutoFit
    xlApp.Range("A1:H100").Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes
       
    xlApp.Visible = True
    
    Set xlApp = Nothing
    
End Sub

Public Function Convert_Status(ByVal Status_Value As Integer) As String

    On Error Resume Next
    Convert_Status = Array("Not Started", "In Progress", "Complete", "Waiting", "Deferred")(Status_Value)
    
End Function

Public Function Convert_Importance(ByVal Importance_Value As Integer) As String
    
    On Error Resume Next
    Convert_Importance = Array("Low", "Normal", "High")(Importance_Value)
    
End Function
Public Function Convert_DateComplete(DateComplete_Value As Date)
    If DateComplete_Value = 949998 Then
        Convert_DateComplete = ""
        Else
        Convert_DateComplete = DateComplete_Value
    End If
        
End Function
