VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "ProgressBar"
   ClientHeight    =   3495
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8865
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Activate()

ProgressBar.Width = 0

Do While ProgressBar.Width < 396
    
    Sleep (10)

    ProgressBar.Width = ProgressBar.Width + 2
    
    DoEvents
    
Loop

MsgBox "Seja Bem Vindo ao ProgressBar!!!", vbInformation, "ProgressBar"

Me.Hide

End Sub
