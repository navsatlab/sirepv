Attribute VB_Name = "main"
' ###################################################################
' # NAVSATLAB                                                       #
' # Application deployed over VBA                                   #
' # https://github.com/navsatlab/sirepv                             #
' # Some rights reserved                                            #
' ###################################################################

Option Explicit

Global lUser As String
Global Const SysVersion As String = "SIREP version 0.92 pre-release | NAVSATLAB - AV" & vbNewLine & "®2024 All rights reserved"
Global Const LockPC As String = "USUARIO-MXL2503"

Public CompareItems As String

Public Type DataContainer
    Titulo As String
    Autor As String
    Lugar As String
    Editorial As String
    Año As String
    Donante As String
    Idioma As String
    Clasificacion As String
    IsValid As Boolean
    ID As String
End Type

Public Sub LoadMakeRefs()
    If Init Then _
        uMakeRefs.Show
End Sub

Public Sub LoadSearch()
Attribute LoadSearch.VB_ProcData.VB_Invoke_Func = "F\n14"
    If Init Then _
        uSearch.Show
End Sub

Public Sub LoadDefinePositions()
    If Init Then _
        uDefinePositions.Show
End Sub

Public Sub LoadSelect()
Attribute LoadSelect.VB_ProcData.VB_Invoke_Func = "S\n14"
    If Init Then _
        uSelect.Show
End Sub

Public Sub LoadListItems()
Attribute LoadListItems.VB_ProcData.VB_Invoke_Func = "W\n14"
    If Init Then _
        uListItems.Show
End Sub

Public Sub LoadCompare()
Attribute LoadCompare.VB_ProcData.VB_Invoke_Func = "C\n14"
    If Init Then _
        uCompare.Show
End Sub

Public Sub LoadPrint()
Attribute LoadPrint.VB_ProcData.VB_Invoke_Func = "P\n14"
    If Init Then _
        uPrint.Show
End Sub

Public Sub LoadRegBooks()
    If Init Then _
        uRegBooks.Show
End Sub

Public Sub LoadReportBookView()
    If Init Then _
        uReportBookView.Show
End Sub

Private Function Init() As Boolean
    If tSheet = "" Then tSheet = "Libros en sala"
    If tTable = "" Then tTable = "MAIN"

    Dim strPC As String
    strPC = Environ("COMPUTERNAME")
    If strPC = LockPC Then
        Init = True
    Else
        Init = False
    End If
End Function

Private Function LoadModify(ID As DataContainer) As DataContainer
    MsgBox "jsjsj"
    uModify.Show
    MsgBox "zout"
End Function

' Verifica que los datos existentes aquí y en MDB sean los correctos y que los datos no se repitan
Public Sub InitVerifyFails()
    If Init = False Then Exit Sub
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    Dim xTAGS As Integer, xFolio As Integer, xCol As Integer, xSeccion As Integer, Data As ListObject, i As Long, buff As Variant, lError As Boolean
    xTAGS = GetPos("TAGS") 'modify it to N° Ficha en Base
    xFolio = GetPos("N° de adquisición")
    xCol = GetPos("Columna")
    xSeccion = GetPos("Área que pertenece")
    
    Set Data = ThisWorkbook.Sheets(tSheet).ListObjects(tTable)
    Set cnData = New ADODB.Connection
    cnData.ConnectionString = ADOPathQuery
    cnData.Open

    i = 1
    For Each buff In Data.ListColumns(xFolio).DataBodyRange
        lError = False
        i = i + 1
        
        Dim lContent As tabID
        Dim lData() As String, xout As String
        Dim dataOut As ListObject
        Set dataOut = ThisWorkbook.Sheets(tSheet).ListObjects(tTable)
        
        If Not buff.Value = "[sin folio]" Or Not Len(buff.Value) = 0 Then
            lData = Split(Trim(ParseNumber(buff.Value)), "-")
            If UBound(lData) = 0 Then
                ' Posiblemente los datos estén erróneos en el folio
                lError = True
            Else
                If Left(lData(1), 1) = "9" Then
                    lData(1) = "19" & lData(1)
                Else
                    lData(1) = "20" & lData(1)
                End If
                xout = lData(1) & "-" & lData(0)
                
                On Error Resume Next
                lContent = FindData(xout)
                On Error GoTo 0
        
                If lContent.Valid = True And lContent.IsRepeated = False Then
                    ' Los datos son correctos
                ElseIf lContent.Valid = False And lContent.IsRepeated = False Then
                    ' No hay ningún registro existente
                    lError = True
                ElseIf lContent.IsRepeated = True Then
                    ' Los datos se encuentran repetidos (duplicación de folio)
                    lError = True
                End If
                
                lContent.Valid = False
            End If
        Else
            ' La celda está vacía o indica «[sin folio]»
            lError = True
        End If
        
        Dim j As Long
        If lError Then
            ' Marcamos el error en la celda
            Dim lTAGS As String
            lTAGS = dataOut.Range(i, xTAGS)
            If Len(lTAGS) = 0 Then
                dataOut.Range(i, xTAGS) = "0x1C"
            Else
                dataOut.Range(i, xTAGS) = lTAGS & ";0x1C"
            End If
            
            ' Re-coloreamos la celda
            Dim Value As Variant
            lData = Split(lTAGS & ";0x1C", ";")

            For j = xCol To xSeccion
                dataOut.Range(i, j).Interior.ColorIndex = 0
                dataOut.Range(i, j).Font.ColorIndex = 1
            Next j
            
            For Each Value In lData
                If Value = "0x10" Then      ' CI
                    For j = xCol To xSeccion
                        dataOut.Range(i, j).Font.ColorIndex = 3
                    Next j
                ElseIf Value = "0x12" Then  ' Para restaurar
                    For j = xCol To xSeccion
                        dataOut.Range(i, j).Interior.ColorIndex = 6
                    Next j
                ElseIf Value = "0x1C" Then  ' En restauración
                    For j = xCol To xSeccion
                        dataOut.Range(i, j).Interior.Color = rgbYellowGreen
                    Next j
                ElseIf Value = "0x1A" Then  ' Libro con errores en ficha
                    For j = xCol To xSeccion
                        dataOut.Range(i, j).Interior.Color = rgbPaleTurquoise
                    Next j
                ElseIf Value = "0x14" Then  ' En catalogación
                    For j = xCol To xSeccion
                        dataOut.Range(i, j).Interior.ColorIndex = 14
                        dataOut.Range(i, j).Font.ColorIndex = 2
                    Next j
                ElseIf Value = "0xFF" Then  ' Perdido
                    For j = xCol To xSeccion
                        dataOut.Range(i, j).Font.ColorIndex = 2
                        dataOut.Range(i, j).Interior.ColorIndex = 9
                    Next j
                End If
            Next
            
        End If
    Next buff
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
