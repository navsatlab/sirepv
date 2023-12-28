VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uSelect 
   Caption         =   "Selección de tabla de trabajo"
   ClientHeight    =   2040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7875
   OleObjectBlob   =   "uSelect.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "uSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ###################################################################
' # NAVSATLAB                                                       #
' # Application deployed over VBA                                   #
' # https://github.com/navsatlab/sirepv                             #
' # Some rights reserved                                            #
' ###################################################################

Option Explicit

Private Sub cChange_Click()
    Dim sData() As String
    sData = Split(cItems.text, "-")
    tSheet = sData(0)
    tTable = sData(1)
    Unload Me
End Sub

Private Sub cExit_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim ExcludeTabs() As String
    ExcludeTabs = Split(GetParam("0x11"), ";")

    Dim ws As Worksheet, tbl As ListObject, i As Integer, Exclude As Boolean
    
    For Each ws In ThisWorkbook.Worksheets
        For i = LBound(ExcludeTabs) To UBound(ExcludeTabs)
            If ws.Name = ExcludeTabs(i) Then _
                Exclude = True
        Next i
        If Exclude = False Then
            For Each tbl In ws.ListObjects
                cItems.AddItem ws.Name & "-" & tbl.Name
            Next tbl
        End If
        Exclude = False
    Next ws
    cItems.ListIndex = 0
    lblVersion.Caption = SysVersion
End Sub
