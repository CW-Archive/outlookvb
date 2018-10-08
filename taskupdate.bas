Attribute VB_Name = "taskupdate"
Option Explicit
Public updateMinutes As Long
Public updateDescription As String
Public updateCanceled As Boolean

Public Sub TaskUpdate()
    Dim objItem As Outlook.TaskItem
    Dim objApp
    Set objApp = Outlook.Application
    Set objItem = objApp.ActiveInspector.CurrentItem
    
    updatetaskform.Show
    
    'Exit if the userform cancel button
    If updateCanceled = True Then Exit Sub
    
    With objItem
        .Body = updateDescription & " - " & Now & " (" & updateMinutes & " Minutes)" & vbNewLine & vbNewLine & .Body
        .ActualWork = .ActualWork + updateMinutes
        .TotalWork = .TotalWork + updateMinutes
        .Save

    End With
End Sub
Sub RunTaskUpdate()
    'For some reason when I add TaskUpdate to the ribbon of my task it doesn't run. _
    It does however if I add a different sub that calls TaskUpdate.
    
    TaskUpdate
End Sub


