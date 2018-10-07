VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} newtaskform 
   Caption         =   "New Task"
   ClientHeight    =   5082
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   8512.001
   OleObjectBlob   =   "newtaskform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "newtaskform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DaysRemainingBox_AfterUpdate()
    ' 7
    DueDateBox.Value = DaysRemainingBox.Value + Now()
End Sub

Private Sub DueDateBox_AfterUpdate()
    DaysRemainingBox.Value = CDate(DueDateBox.Value) - Now()
    DaysRemainingBox.Value = Round(DaysRemainingBox.Value)
        
End Sub

Private Sub SpinButton1_SpinDown()
    DaysRemainingBox.Value = DaysRemainingBox.Value - 1
    DaysRemainingBox_AfterUpdate
    
End Sub

Private Sub SpinButton1_SpinUp()
    DaysRemainingBox.Value = DaysRemainingBox.Value + 1
    DaysRemainingBox_AfterUpdate
    
End Sub

Private Sub UserForm_Initialize()
    'Prefill from selected mail object

    If isMemo = True Then Category1.Value = True
    If isCloseout = True Then Category2.Value = True
    If isPersonal = True Then Category3.Value = True
    If isPlanUpdate = True Then Category4.Value = True
    If isPricing = True Then Category5.Value = True
    If isRFI = True Then Category6.Value = True
    If isSubmittal = True Then Category7.Value = True
    
    StartDateBox = Now()
    DueDateBox = Now() + 7
    DaysRemainingBox = 7

    SubjectBox.Value = objSubject
    BodyBox.Value = objBody
    
End Sub

Private Sub CreateButton_Click()
    'Create Button
    
    isMemo = Category1.Value
    isCloseout = Category2.Value
    isPersonal = Category3.Value
    isPlanUpdate = Category4.Value
    isPricing = Category5.Value
    isRFI = Category6.Value
    isSubmittal = Category7.Value
    
    objStartDate = StartDateBox.Value
    objDueDate = DueDateBox.Value
    objSubject = SubjectBox.Value
    objBody = BodyBox.Value
    
    Unload Me
    
End Sub

Private Sub CancelButton_Click()
    'Cancel Button
    canceled = True
    Unload Me
    Exit Sub
    
End Sub
