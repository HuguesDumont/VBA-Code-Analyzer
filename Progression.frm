VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Progression 
   Caption         =   "Please wait while code is analyzed. It may take several minutes."
   ClientHeight    =   1785
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9840
   OleObjectBlob   =   "Progression.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Progression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Me.BarProgress.Value = 0
End Sub
