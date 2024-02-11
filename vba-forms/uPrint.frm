VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uPrint 
   Caption         =   "Impresión de fichas catalográficas"
   ClientHeight    =   7305
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16935
   OleObjectBlob   =   "uPrint.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "uPrint"
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

Private lPosition As Long
Private lReset As Boolean, lLoad As Boolean

Private xClasificacion, xFolio, xNotas, xFicha As Long
Private HTMLFicha As String, HTMLPrint() As String

Private Const HTMLHeader As String = "<!DOCTYPE html><head><meta name='viewport' content='width=device-width, initial-scale=1.0'><style>body{justify-content: center;align-items: center;}.iframe-container {width: 5in;height: 3in;float: left;overflow: hidden;}iframe{width: 100%;height: 100%;}table, th, td {border: 1px solid black;border-collapse: collapse;}</style></head>"

Private Sub FillExcelData(ID As Long)
    If ID = 0 Then
        tNotas.text = ""
        tClasificacion.text = ""
        tFolio.text = ""
        tNotas.Enabled = False
        Exit Sub
    End If
    tNotas.Enabled = True
    Dim content As ListObject

    Set content = ThisWorkbook.Sheets(tSheet).ListObjects(tTable)
    tClasificacion.text = Trim(content.Range(ID, xClasificacion))
    tFolio.text = Trim(content.Range(ID, xFolio))
    tNotas.text = Trim(content.Range(ID, xNotas))
End Sub

Private Sub SaveExcelData(ID As Long)
    Dim content As ListObject

    Set content = ThisWorkbook.Sheets(tSheet).ListObjects(tTable)
    content.Range(ID, xNotas) = Trim(tNotas.text)
End Sub

Private Function LoadWebData(ID As String) As String
    Dim result As tabID
    
    result = GetBookData(wFicha, ID)
    HTMLFicha = result.HTMLContent
    
    xFicha = ID
    If Not result.ContentEmpty Then
        Dim lData() As String, Data As String
        lData = Split(result.CONTENTS(0).ID, "-")
        Data = lData(1) & "-" & Right(lData(0), 2)
    End If
    
    LoadWebData = Data
End Function

Private Sub ShowError(description As String)
    wFicha.Navigate "about:blank"
    Do While wFicha.Busy Or wFicha.ReadyState <> 4
        DoEvents
    Loop
    
    Dim buff As String
    buff = ""
    wFicha.Document.Write description
End Sub

Private Sub cAdd_Click()
    On Error Resume Next
    Application.ScreenUpdating = False
    pList.ListItems.Add , , xFicha
    
    SaveExcelData (lPosition)
    tNotas.text = ""
    tFolio.text = ""
    tClasificacion.text = ""
    
    cAdd.Enabled = False
    cAdd.Default = False
    cLocate.Default = True
    
    cAdd.Caption = "Agregar ficha"
    cLocate.Caption = "[Enter] Buscar"
    cCancel.Caption = "[Esc] Cancelar"
    lLoad = False
    
    ReDim Preserve HTMLPrint(UBound(HTMLPrint) + 1)
    HTMLPrint(UBound(HTMLPrint)) = HTMLFicha
    
    xFicha = 0
    HTMLFicha = ""
    
    lCount.Caption = pList.ListItems.Count
    wFicha.Navigate "about:blank"
    tID.SetFocus
    Application.ScreenUpdating = True
End Sub

Private Sub cCancel_Click()
    If lLoad = True Then
        cAdd.Enabled = False
        cAdd.Default = False
        cLocate.Default = True
        wFicha.Navigate "about:blank"
        
        cAdd.Caption = "Agregar ficha"
        cCancel.Caption = "[Esc] Cancelar"
        cLocate.Caption = "[Enter] Buscar"
        
        lLoad = False
        tID.SetFocus
    Else
        Unload Me
    End If
End Sub

Private Sub cExit_Click()
    Unload Me
End Sub

Private Sub cDeleteItem_Click()
    If pList.selectedItem.Index <= 0 Then _
        Exit Sub
    
    pList.ListItems.Remove pList.selectedItem.Index
    Dim i As Integer, htmltemp() As String
    ReDim htmltemp(UBound(HTMLPrint) - 1)
    For i = 1 To UBound(htmltemp)
        Dim lPos As Integer
        If i < pList.selectedItem.Index Then
            htmltemp(i) = HTMLPrint(i)
        ElseIf i >= pList.selectedItem.Index Then
            htmltemp(i) = HTMLPrint(i + 1)
        End If
    Next i
    
    HTMLPrint = htmltemp
End Sub

Private Sub cLocate_Click()
    If Len(tID.text) = 0 Then
        tID.SetFocus
        Exit Sub
    End If
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Dim out As String
    out = LoadWebData(tID.text)
    
    If Not out = "" Then
        lPosition = FindExcelData(out, xFolio)
        FillExcelData lPosition
    End If
    
    cLocate.Default = False
    cAdd.Enabled = True
    cAdd.Default = True
    
    cLocate.Caption = "Buscar"
    cCancel.Caption = "[Esc] Limpiar"
    cAdd.Caption = "[Enter] Agregar ficha"
    lLoad = True
    
    If lPosition > 0 Then tNotas.SetFocus
    If lPosition = 0 Then cAdd.SetFocus
    tID.text = ""
    Application.ScreenUpdating = True
    Exit Sub
Fail:
    ShowError (Err.description)
    tID.SetFocus
    tID.SelStart = 0
    tID.SelLength = Len(tID.text)
    Application.ScreenUpdating = True
End Sub

Private Sub cPrint_Click()
    ' Imprimimos
    Dim result As String, i As Integer
    result = HTMLHeader & "<body>"
    
    Dim zCol As Integer, zRow As Integer
    zCol = 1
    zRow = 1
    For i = 1 To UBound(HTMLPrint)
        If zCol = 1 Then
            If zRow = 1 Then
                result = result & "<br><br><br><br><br><br><br><table width='10in' height='6in'><tr><td><div class='iframe-container'>"
                result = result & "<iframe id='datax' srcdoc=" & """"
                result = result & HTMLPrint(i) & """"
                result = result & " frameborder='0' scrolling='no'></iframe></div></td>"
                
                zRow = zRow + 1
            Else
                result = result & "<td><div class='iframe-container'>"
                result = result & "<iframe id='datax' srcdoc=" & """"
                result = result & HTMLPrint(i) & """"
                result = result & " frameborder='0' scrolling='no'></iframe></div></td></tr>"
                
                zRow = 1
                zCol = zCol + 1
            End If
        Else
            If zRow = 1 Then
                result = result & "<tr><td><div class='iframe-container'>"
                result = result & "<iframe id='datax' srcdoc=" & """"
                result = result & HTMLPrint(i) & """"
                result = result & " frameborder='0' scrolling='no'></iframe></div></td>"
                
                zRow = zRow + 1
            Else
                result = result & "<td><div class='iframe-container'>"
                result = result & "<iframe id='datax' srcdoc=" & """"
                result = result & HTMLPrint(i) & """"
                result = result & " frameborder='0' scrolling='no'></iframe></div></td></tr></table>"
                
                zRow = 1
                zCol = 1
            End If
        End If
    Next i
    
    result = result & "<script>var iframe = document.getElementById('datax');iframe.onload = function () {var iframeDocument = iframe.contentDocument || iframe.contentWindow.document;var contenidoWidth = iframeDocument.documentElement.scrollWidth;var contenidoHeight = iframeDocument.documentElement.scrollHeight;var scaleX = iframe.clientWidth / contenidoWidth;var scaleY = iframe.clientHeight / contenidoHeight;iframeDocument.documentElement.style.transform = 'scale(' + scaleX + ',' + scaleY + ')';iframeDocument.documentElement.style.transformOrigin = '0 0';};</script>"
    
    result = result & "</body></html>"
    
    Dim SaveFileName As String
    Dim FileNumber As Integer
    
    ' Prompt the user to select a location and name for the file
    SaveFileName = Application.GetSaveAsFilename( _
        InitialFileName:="Fichas.html", _
        FileFilter:="Documento HTML (*.html), *.html")
    
    ' Check if the user canceled the Save As dialog
    If SaveFileName = "False" Then
        Exit Sub
    End If
    
    ' Open the file for writing
    FileNumber = FreeFile
    Open SaveFileName For Output As #FileNumber
    
    ' Write the content to the file
    Print #FileNumber, result
    
    ' Close the file
    Close #FileNumber
    
    ' Inform the user that the file has been saved
    MsgBox "Se ha guardado el archivo satisfactoriamente en " & SaveFileName, vbInformation
    Unload Me
End Sub

Private Sub pList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item.Selected = True Then
        cDeleteItem.Enabled = True
        wFicha.Navigate "about:blank"
        Do While wFicha.Busy Or wFicha.ReadyState <> 4
            DoEvents
        Loop
        
        wFicha.Document.Write HTMLPrint(Item.Index)
    ElseIf Item.Selected = False Then
        cDeleteItem.Enabled = False
        wFicha.Navigate "about:blank"
    End If
End Sub

Private Sub UserForm_Activate()
    tID.SetFocus
End Sub

Private Sub UserForm_Initialize()
    StartUpPosition = 0
    Left = Application.Left + (Application.width - width - 10)
    Top = Application.Top + (0.5 * Application.Height) - (0.5 * Height)
    
    Set cnData = New ADODB.Connection
    
    cnData.ConnectionString = ADOPathQuery
    cnData.Open
    
    xClasificacion = GetPos("Clasificación")
    xFolio = GetPos("N° de adquisición")
    xNotas = GetPos("Notas")
    
    Me.Caption = "Impresión de fichas catalográficas -- " & tSheet
    ReDim HTMLPrint(0)
    
    Application.Calculation = xlCalculationManual
    
    pList.View = lvwReport
    pList.Gridlines = True
    pList.LabelEdit = lvwManual
    pList.FullRowSelect = True
    pList.ColumnHeaders.Add , , "Fichas catalográficas", 100
    
    lblVersion.Caption = SysVersion
    
    ThisWorkbook.Save
    Me.Caption = "Impresión de fichas catalográficas -- " & tSheet & " (" & ThisWorkbook.Sheets(tSheet).ListObjects(tTable).Range.Rows.Count & " libros registrados)"
End Sub

Private Sub UserForm_Terminate()
    cnData.Close
    
    ThisWorkbook.Save
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

