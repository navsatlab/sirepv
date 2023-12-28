VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uTheme 
   Caption         =   "Edición de temas en biblioteca"
   ClientHeight    =   4440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6585
   OleObjectBlob   =   "uTheme.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "uTheme"
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

Dim xSec As Integer, xFolio As Integer, xSection As Integer
Dim xdata As Object
Dim lFirst, lSecond As Long

Private Sub cCancel_Click()
    If cCancel.Caption = "Limpiar [ESC]" Then
        tMain.text = ""
        tSecondary.text = ""
        cCancel.Caption = "Cerrar [ESC]"
        tMain.Enabled = False
        tSecondary.Enabled = False
        tFirst.Enabled = True
        tLast.Enabled = True
        cValidate.Enabled = True
        cModify.Enabled = False
        cValidate.Default = True
        
        tFirst.SetFocus
        tFirst.SelStart = 0
        tFirst.SelLength = Len(tFirst.text)
    Else
        Unload Me
    End If
End Sub

Private Sub cModify_Click()
    ' Add or modify range tables for this
    If Len(tMain.text) = 0 Then
        tMain.SetFocus
        Exit Sub
    ElseIf Len(tSecondary.text) = 0 Then
        tSecondary.SetFocus
        Exit Sub
    End If
    
    Dim i As Long
    Dim content As ListObject

    Set content = ThisWorkbook.Sheets(tSheet).ListObjects(tTable)
    For i = lFirst To lSecond
        If Len(tSecondary.text) = 0 Then
            content.Range(i, xSection) = tMain.text
        Else
            content.Range(i, xSection) = tMain.text & Chr(10) & tSecondary.text
        End If
    Next i
    
    Unload Me
    'Set content = ThisWorkbook.Sheets("Settings").ListObjects("THEME")
End Sub

Private Sub cValidate_Click()
    ' Validate if first item is located low than last item
    If Len(tFirst.text) = 0 Then
        tFirst.SetFocus
        Exit Sub
    ElseIf Len(tLast.text) = 0 Then
        tLast.SetFocus
        Exit Sub
    End If
    
    lFirst = FindExcelData(tFirst.text, xFolio)
    lSecond = FindExcelData(tLast.text, xFolio)
    
    If lSecond > lFirst Then
        cCancel.Caption = "Limpiar [ESC]"
        tFirst.Enabled = False
        tLast.Enabled = False
        tMain.Enabled = True
        tSecondary.Enabled = True
        cValidate.Enabled = False
        cModify.Enabled = True
        cModify.Default = True
        tMain.SetFocus
        
    Else
        MsgBox "Los valores que ingresó entre el libro inicial y final no pueden procesarse. Esto puede deberse a que quizá el libro final esté más arriba de la tabla que la posición inicial, o viceversa.", vbCritical, "Error de validación"
        tFirst.SetFocus
        tFirst.SelStart = 0
        tFirst.SelLength = Len(tFirst.text)
    End If
End Sub

Private Sub UserForm_Initialize()
    Set xdata = CreateObject("Scripting.Dictionary")
    
    Dim Data As ListObject
    Dim buff As Range
    Dim lData As Variant
    
    xSec = 1
    xFolio = GetPos("N° de adquisición")
    xSection = GetPos("Área que pertenece")
    
    Set Data = ThisWorkbook.Sheets("Settings").ListObjects("THEME")
    For Each buff In Data.ListColumns(xSec).DataBodyRange
        If Not xdata.Exists(buff.Value) Then
            xdata.Add buff.Value, buff.Value
        End If
    Next buff
    
    For Each lData In xdata.Keys
        tMain.AddItem lData
    Next lData
    
    cValidate.Default = True
    tFirst.SetFocus
    lblVersion.Caption = SysVersion
    
    Application.Calculation = xlCalculationManual
End Sub

Private Sub UserForm_Terminate()
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
