VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uAbout 
   Caption         =   "Acerca de"
   ClientHeight    =   4260
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7650
   OleObjectBlob   =   "uAbout.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "uAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cExit_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    lVersion.Caption = SysVersion
End Sub
