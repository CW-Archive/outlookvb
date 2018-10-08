VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} updatetaskform 
   Caption         =   "Update Task"
   ClientHeight    =   3225
   ClientLeft      =   90
   ClientTop       =   405
   ClientWidth     =   6390
   OleObjectBlob   =   "updatetaskform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "updatetaskform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelButton_Click()
    updateCanceled = True
    Unload Me
    Exit Sub
    
End Sub

Private Sub ComboBox1_Change()
    
End Sub

Private Sub CreateButton_Click()
    '
    updateMinutes = MinutesBox.Value
    updateDescription = DescriptionBox.Value
    updateStatus = StatusBox.ListIndex
    updateDueDate = DueDateBox.Value
    
    Unload Me
    
End Sub


Private Sub DueDateBox_Change()

End Sub

Private Sub MinutesSpinButton_SpinDown()
    If MinutesBox.Value = 15 Then
        MinutesBox.Value = 5
    Else
        If MinutesBox.Value <> 5 Then MinutesBox.Value = MinutesBox.Value - 15
    End If
    
    
End Sub

Private Sub MinutesSpinButton_SpinUp()
    If MinutesBox.Value = 5 Then
        MinutesBox.Value = 15
    Else
        MinutesBox.Value = MinutesBox.Value + 15
    End If

End Sub


Private Sub SpinButton1_SpinDown()
    DueDateBox.Value = CDate(DueDateBox.Value) - 1
    Label4.Caption = "Due Date ( Days: " & CDate(DueDateBox.Value) - Date & " )"
    
End Sub

Private Sub SpinButton1_SpinUp()
    DueDateBox.Value = CDate(DueDateBox.Value) + 1
    Label4.Caption = "Due Date ( Days: " & CDate(DueDateBox.Value) - Date & " )"
    
End Sub

Private Sub StatusBox_Change()

End Sub

Private Sub UserForm_Initialize()
    
    MinutesBox.Value = 15
    DescriptionBox.Value = updateDescription
    DueDateBox.Value = updateDueDate
    
    StatusBox.List = Array("Not Started", "In Progress", "Complete", "Waiting", "Deferred")
    
    If updateStatus = 0 Then updateStatus = 1
    StatusBox.ListIndex = updateStatus
    Label4.Caption = "Due Date ( Days: " & CDate(DueDateBox.Value) - Date & " )"
    
    
End Sub

 Private Sub DescriptionBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
     If KeyCode = vbKeyReturn Then
          CreateButton_Click
     End If
     
End Sub
