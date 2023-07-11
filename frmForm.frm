VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmForm 
   Caption         =   "Automated Data Entry form version 1.0"
   ClientHeight    =   5940
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   11340
   OleObjectBlob   =   "frmForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdReset_Click()

    Dim msgValue As VbMsgBoxResult
    msgValue = MsgBox("Do you want to reset the form?", vbYesNo + vbInformation, "Confirmation")
    
    If msgValue = vbNo Then Exit Sub
    
    Call Reset
        

End Sub

Private Sub cmdSave_Click()

    Dim msgValue As VbMsgBoxResult
    msgValue = MsgBox("Do you want to save the data? ", vbYesNo + vbInformation, "Confirmation")
    
    If msgValue = vbNo Then Exit Sub
    
    Call Submit
    Call Reset
        

End Sub

Private Sub UserForm_Initialize()
    
    Call Reset

End Sub


