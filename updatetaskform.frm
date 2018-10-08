VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} updatetaskform 
   Caption         =   "Update Task"
   ClientHeight    =   2387
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   6587
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

Private Sub CreateButton_Click()
    '
    updateMinutes = MinutesBox.Value
    updateDescription = DescriptionBox.Value
    
    Unload Me
    
End Sub

Private Sub MinutesSpinButton_SpinDown()
    MinutesBox.Value = MinutesBox.Value - 15
    
End Sub

Private Sub MinutesSpinButton_SpinUp()
    MinutesBox.Value = MinutesBox.Value + 15
    
End Sub

Private Sub UserForm_Initialize()
    MinutesBox.Value = 15
    
End Sub


