VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uReportBookView 
   Caption         =   "Reporte de consultas y estadística mensuales"
   ClientHeight    =   7080
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15120
   OleObjectBlob   =   "uReportBookView.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "uReportBookView"
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

Private Sub cCancel_Click()
    Unload Me
End Sub

Private Sub cGenerate_Click()
    If cMonth.ListIndex = -1 Then
        MsgBox "Por favor selecciona un mes válido para hacer el reporte", vbCritical, "Reporte de consultas"
        Exit Sub
    End If
    
    ' Initialize lists
    Dim Data As ListObject, Item As ListItem, i As Integer, ListDict As Object, lData As Variant
    Set Data = ThisWorkbook.Sheets("Consultas").ListObjects("READS")
    Set ListDict = CreateObject("Scripting.Dictionary")
    ReDim lList(0)

    For i = 2 To Data.Range.Rows.Count
        Dim rDate As Date, rContent As String
        rDate = CDate(Data.Range.Cells(i, 1))
        rContent = Trim(Data.Range.Cells(i, 3))
        If (Month(rDate) = (cMonth.ListIndex + 1)) And (Year(rDate) = CInt(cYear.text)) Then
            If Not ListDict.Exists(rContent) Then _
                ListDict.Add rContent, rContent
        End If
    Next i
    
    If ListDict.Count = 0 Then
        MsgBox "No se encontraron datos de consulta disponibles para ese rango seleccionado. Vuelve a verificarlo.", vbExclamation, "Reporte de consultas"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Worksheets("temp").Cells.Clear
    
    ' Generates chart based with provided data
    Dim sTempFile As String
    Dim oChart As Object, oRange As Range
    
    With Worksheets("temp")
        .Range("A1") = "Tema"
        .Range("B1") = "Conteo"
        
        Dim RowsCount As Long
        .ListObjects.Add(xlSrcRange, .Range("A1:B1"), , xlYes).Name = "tmpList"
        With .ListObjects("tmpList")
            Dim row As ListRow
            For Each lData In ListDict.Keys
                Set row = .ListRows.Add
                row.Range(1) = lData
                
                
                row.Range(2).Formula = "=COUNTIFS(READS[Sección a la que pertenece], [@Tema], READS[Fecha], "">=" & _
                    CDate("1/" & CStr(cMonth.ListIndex + 1) & "/" & cYear.text) & """, READS[Fecha], ""<" & _
                    DateAdd("m", 1, CDate("1/" & CStr(cMonth.ListIndex + 1) & "/" & cYear.text)) & """)"
                
                RowsCount = RowsCount + 1
            Next lData
        End With
    End With

    sTempFile = Environ("temp") & "\temp.gif"
    Dim RangeSelected As String
    RangeSelected = "temp!$A1:$B" + CStr(RowsCount + 1)
    
    Set oRange = Worksheets("temp").Range(RangeSelected)
    Set oChart = Worksheets("temp").Shapes.AddChart2

    With oChart.Chart
        .Parent.width = 480
        .Parent.Height = 360
        
        .SetSourceData Source:=oRange
        .HasTitle = True
        .ChartTitle.text = "Reporte de lecturas del mes de " & cMonth.text
        .ChartType = xlBarStacked
        .SetElement msoElementPrimaryCategoryAxisShow
        .SetElement msoElementDataLabelShow
        .SetElement msoElementDataLabelInsideEnd

        .Export Filename:=sTempFile, FilterName:="GIF"
        .Parent.Delete
    End With
    Dim HTMLImage As String
    HTMLImage = "<!DOCTYPE html><html><head><style type='text/css'>img.big-img{display:block;width:100%;}</style></head><body><img src='" & sTempFile & "' class='big-img'></body></html>"
    
    wTest.Navigate "about:blank"
    Do While wTest.Busy Or wTest.ReadyState <> 4
        DoEvents
    Loop
    
    wTest.Document.Write HTMLImage
    Do While wTest.Busy Or wTest.ReadyState <> 4
        DoEvents
    Loop

    Repaint

    On Error Resume Next
    Kill sTempFile
    Application.ScreenUpdating = True
End Sub

Private Sub UserForm_Initialize()
    StartUpPosition = 0
    Left = Application.Left + (Application.width - width - 10)
    Top = Application.Top + (0.5 * Application.Height) - (0.5 * Height)
    
    cMonth.AddItem "Enero"
    cMonth.AddItem "Febrero"
    cMonth.AddItem "Marzo"
    cMonth.AddItem "Abril"
    cMonth.AddItem "Mayo"
    cMonth.AddItem "Junio"
    cMonth.AddItem "Julio"
    cMonth.AddItem "Agosto"
    cMonth.AddItem "Septiembre"
    cMonth.AddItem "Octubre"
    cMonth.AddItem "Noviembre"
    cMonth.AddItem "Diciembre"
    
    cMonth.text = cMonth.List(Month(Now) - 1)
    cYear.text = Year(Now)
    
    lblVersion.Caption = SysVersion
    ThisWorkbook.Save
End Sub
