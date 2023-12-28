Attribute VB_Name = "main"
' ###################################################################
' # NAVSATLAB                                                       #
' # Application deployed over VBA                                   #
' # https://github.com/navsatlab/sirepv                             #
' # Some rights reserved                                            #
' ###################################################################

Option Explicit

Global lUser As String
Global Const SysVersion As String = "SIREP version 0.9 pre-release | NAVSATLAB - AV" & vbNewLine & "®2023 All rights reserved"
Global Const LockPC As String = "IAGO-MXL2503VB9"

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

Private Sub getclas()
    Init
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    Dim xdata As Integer, xFolio As Integer, Data As ListObject, lp As Long, i As Long, buff As Variant
    xdata = GetPos("C-TEST") 'modify it to N° Ficha en Base
    xFolio = GetPos("N° de adquisición")
    
    Set Data = ThisWorkbook.Sheets(tSheet).ListObjects(tTable)
    i = 1
    
    Set cnData = New ADODB.Connection
    
    cnData.ConnectionString = ADOPathQuery
    cnData.Open

    For Each buff In Data.ListColumns(xFolio).DataBodyRange
        i = i + 1
        lp = lp + 1
        
        Dim lContent As tabID
        Dim lData() As String, xout As String
        Dim dataOut As ListObject
        Set dataOut = ThisWorkbook.Sheets(tSheet).ListObjects(tTable)
        
        If Not buff.Value = "[sin folio]" Then
            lData = Split(Trim(buff.Value), "-")
            If UBound(lData) = 0 Then
                dataOut.Range(i, xdata) = "?REV-" & buff.Value
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
                    dataOut.Range(i, xdata) = Trim(lContent.MARC082)
                ElseIf lContent.Valid = False And lContent.IsRepeated = False Then
                    dataOut.Range(i, xdata) = ""
                ElseIf lContent.IsRepeated = True Then
                    dataOut.Range(i, xdata) = "!REP-" & Trim(lContent.MARC082)
                End If
                
                lContent.Valid = False
            End If
        Else
            dataOut.Range(i, xdata) = ""
        End If
    Next buff
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
