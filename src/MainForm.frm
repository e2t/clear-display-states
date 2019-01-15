VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Стереть состояния отображения"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4185
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub butClose_Click()
    ExitApp
End Sub

Private Sub butRun_Click()
    Run Me.radOneColor.value
    ExitApp
End Sub
