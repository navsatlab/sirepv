VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uMakeRefs 
   Caption         =   "Generador de guías en sala"
   ClientHeight    =   8730.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8580.001
   OleObjectBlob   =   "uMakeRefs.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "uMakeRefs"
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

Private Type ArrayContent
    Columna As String
    Charola As String
    Folio1 As String
    Folio2 As String
End Type

Private lItems() As ArrayContent, lContent As tabID, lNameDefined() As String

Private Sub cAdd_Click()
    On Error GoTo Fail
    If Len(cColumna.text) = 0 Then
        MsgBox "Por favor escriba el número de columna para realizar la búsqueda", vbCritical, "Generar guías"
        cColumna.SetFocus
        Exit Sub
    End If
    If Len(cCharola.text) = 0 Then
        MsgBox "Por favor escriba el número de charola para realizar la búsqueda", vbCritical, "Generar guías"
        cCharola.SetFocus
        Exit Sub
    End If
    
    Dim i As Long, lFound As Boolean, xSeccion As Long, xFolio As Long, lItemID As Long
    
    ' Create script dictionary
    xSeccion = GetPos("Área que pertenece")
    xFolio = GetPos("N° de Adquisición")
    
    For i = 1 To UBound(lItems)
        If (lItems(i).Columna = Trim(cColumna.text)) And (lItems(i).Charola = Trim(cCharola.text)) Then
            lFound = True
            lItemID = i
        End If
    Next i
    
    If lFound Then
        ' Localizamos el primer ID en la tabla de Excel, para generar una lista de posibles guías a generar
        Dim FirstID As Long, SecondID As Long, lExcelItems As Object, lExcelData As ListObject, lExcelItem As Range, lItemsAdded As Integer, X As tabID
        Set lExcelItems = CreateObject("Scripting.Dictionary")
        
        For i = FindExcelData(lItems(lItemID).Folio1, xFolio) To FindExcelData(lItems(lItemID).Folio2, xFolio)
            ' Generamos un script dictionary para dimensionar los elementos a generar
            Dim content As ListObject, TextSection As String
            Set content = ThisWorkbook.Sheets(tSheet).ListObjects(tTable)
            TextSection = Replace(Trim(content.Range(i, xSeccion)), Chr(10), "|")
            
            Dim Item As ListItem
            If Not lExcelItems.Exists(TextSection) Then
                ' Si el elemento no existe, agregamos nueva guía
                lExcelItems.Add TextSection, TextSection
                
                Set Item = pList.ListItems.Add(, , lItems(lItemID).Columna)
                Item.SubItems(1) = lItems(lItemID).Charola
                
                ' Añadimos la información de guía (clasificación)
                X = FindData(InverseID(Trim(content.Range(i, xFolio))))
                Item.SubItems(2) = UCase(X.MARC082)
                
                ' Identificamos si ésta contiene alguna etiqueta de autor, para separarla y agregarle el autor a la guía
                Dim j As Integer, lAddAuthor As Boolean
                For j = 1 To UBound(lNameDefined)
                    If StrComp(lNameDefined(j), TextSection, vbTextCompare) = 0 Then
                        lAddAuthor = True
                    End If
                Next j
                
                If lAddAuthor Then
                    Dim lAuthor() As String, lTemp() As String
                    lTemp = Split(TextSection, "|")
                    If Len(X.MARC100) > 0 Then
                        lAuthor = Split(X.MARC100, ",")
                        Item.SubItems(3) = lTemp(0) & "|" & lAuthor(0) & " - "
                    Else
                        Item.SubItems(3) = lTemp(0) & "|[sin autor] - "
                    End If
                Else
                    Item.SubItems(3) = UCase(TextSection)
                End If
                lAddAuthor = False
                
                ' Identificamos si éste es el primer elemento añadido, si no se busca el elemento anterior para definir cuál es su último libro
                If lItemsAdded > 0 Then
                    X = FindData(InverseID(Trim(content.Range(i - 1, xFolio))))
                    Set Item = pList.ListItems(pList.ListItems.Count - 1)
                    TextSection = Replace(Trim(content.Range(i - 1, xSeccion)), Chr(10), "|")
                    
                    Item.SubItems(4) = UCase(X.MARC082)
                    
                    For j = 1 To UBound(lNameDefined)
                        If StrComp(lNameDefined(j), TextSection, vbTextCompare) = 0 Then
                            lAddAuthor = True
                        End If
                    Next j
                    
                    If lAddAuthor Then
                        If Len(X.MARC100) > 0 Then
                            lAuthor = Split(X.MARC100, ",")
                            Item.SubItems(3) = UCase(Item.SubItems(3) & lAuthor(0))
                        Else
                            Item.SubItems(3) = UCase(Item.SubItems(3) & "[sin autor]")
                        End If
                    End If
                    lAddAuthor = False
                End If
                
                Set Item = Nothing
                lAddAuthor = False
                lItemsAdded = lItemsAdded + 1
            End If
        Next i
        
        ' Agregamos el último elemento a la lista
        X = FindData(InverseID(Trim(content.Range(i - 1, xFolio))))
        Set Item = pList.ListItems(pList.ListItems.Count)
        TextSection = Replace(Trim(content.Range(i - 1, xSeccion)), Chr(10), "|")
        
        Item.SubItems(4) = UCase(X.MARC082)
        
        For j = 1 To UBound(lNameDefined)
            If StrComp(lNameDefined(j), TextSection, vbTextCompare) = 0 Then
                lAddAuthor = True
            End If
        Next j
        
        If lAddAuthor Then
            If Len(X.MARC100) > 0 Then
                lAuthor = Split(X.MARC100, ",")
                Item.SubItems(3) = UCase(Item.SubItems(3) & lAuthor(0))
            Else
                Item.SubItems(3) = UCase(Item.SubItems(3) & "[sin autor]")
            End If
        End If
        lAddAuthor = False
    End If
    
    If Not lFound Then
        MsgBox "No se pudo localizar la charola específica. Por favor reintenta la búsqueda nuevamente", vbCritical, "Generar guías"
        cColumna.SetFocus
        Exit Sub
    End If
    
    cGenerate.Enabled = True
    cColumna.SetFocus
    
    Exit Sub
Fail:
    MsgBox "Ha ocurrido un error, posiblemente la ficha catalográfica a la que se hace referencia no existe. Por favor verifica", vbCritical, "Generar guías"
    Exit Sub
End Sub

Private Sub cExit_Click()
    Unload Me
End Sub

Private Sub cGenerate_Click()
    ' Generamos guías y guardamos
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    With ThisWorkbook.Sheets("Guías")
        .Cells.Clear
        .Cells.VerticalAlignment = xlCenter
        .Cells.HorizontalAlignment = xlCenter
        .Cells.Font.Name = "Times News Roman"
        .Cells.Font.Size = 16
        
        ThisWorkbook.Sheets("Guías").Columns("A").ColumnWidth = 16.86
        ThisWorkbook.Sheets("Guías").Columns("B").ColumnWidth = 49.43
        ThisWorkbook.Sheets("Guías").Columns("C").ColumnWidth = 16.86
        ThisWorkbook.Sheets("Guías").Columns("D").ColumnWidth = 2
        ThisWorkbook.Sheets("Guías").Columns("E").ColumnWidth = 5
        
        ThisWorkbook.Sheets("Guías").Rows.RowHeight = 84.75
        Dim i As Long
        For i = 1 To pList.ListItems.Count
            .Cells(i, 1) = Replace(Replace(pList.ListItems(i).SubItems(2), "-", Chr(10)), " ", "")
            .Cells(i, 2) = UCase(Replace(pList.ListItems(i).SubItems(3), "|", Chr(10)))
            
            .Cells(i, 2).Characters(Start:=InStr(pList.ListItems(i).SubItems(3), "|"), _
                Length:=(Len(pList.ListItems(i).SubItems(3)) - InStr(pList.ListItems(i).SubItems(3), "|")) + 2).Font.Size = 18
                
            .Cells(i, 3) = Replace(Replace(pList.ListItems(i).SubItems(4), "-", Chr(10)), " ", "")
            .Cells(i, 5) = pList.ListItems(i).text & "," & pList.ListItems(i).SubItems(1)
            
            .Cells(i, 1).Borders.LineStyle = xlDouble
            .Cells(i, 2).Borders.LineStyle = xlDouble
            .Cells(i, 3).Borders.LineStyle = xlDouble
            .Cells(i, 5).Borders.LineStyle = xlContinuous
        Next i
        
        .Cells(i, 2) = "Guías generadas automáticamente el" & Chr(10) & Now
    End With
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ThisWorkbook.Save
    
    MsgBox "Por favor revisa las guías generadas, algunas pueden tener errores de autor o simplemente el autor sea otro y no el que indica la sección", vbInformation, "Generar guías"
    ThisWorkbook.Sheets("Guías").Activate
    Unload Me
End Sub

Private Sub pList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    lClasif1.Caption = Replace(Item.SubItems(2), "-", vbNewLine)
    lArea.Caption = Replace(Item.SubItems(3), "|", vbNewLine)
    lClasif2.Caption = Replace(Item.SubItems(4), "-", vbNewLine)
End Sub

Private Sub UserForm_Initialize()
    StartUpPosition = 0
    Left = Application.Left + (Application.width - width - 10)
    Top = Application.Top + (0.5 * Application.Height) - (0.5 * Height)
    
    Set cnData = New ADODB.Connection
    
    cnData.ConnectionString = ADOPathQuery
    cnData.Open
    
    pList.View = lvwReport
    pList.Gridlines = True
    pList.LabelEdit = lvwManual
    pList.FullRowSelect = True
    pList.ColumnHeaders.Add , , "Col", 25
    pList.ColumnHeaders.Add , , "Cha", 25
    pList.ColumnHeaders.Add , , "Clasif. Inicial", 90
    pList.ColumnHeaders.Add , , "Área que pertenece", 150
    pList.ColumnHeaders.Add , , "Clasif. Final", 90
    
    ' Localizamos los tags de qué guías requieren ingreso de nombres, y las añadimos a una lista
    Dim buff As Range, content As ListObject, i As Integer

    Set content = ThisWorkbook.Sheets("Settings").ListObjects("SUFFIX")
    ReDim lNameDefined(0)
    i = 2
    
    For Each buff In content.ListColumns(8).DataBodyRange
        If StrComp(buff.Value, "X", vbTextCompare) = 0 Then
            ReDim Preserve lNameDefined(UBound(lNameDefined) + 1)
            lNameDefined(UBound(lNameDefined)) = Replace(Trim(ThisWorkbook.Sheets("Settings").ListObjects("Suffix").Range(i, 6)), Chr(10), "|")
        End If
        i = i + 1
    Next buff
    
    Dim Data As ListObject
    Set Data = ThisWorkbook.Sheets("Settings").ListObjects("EXTERN_PREFIX")
    
    ReDim lItems(0)
    For i = 2 To Data.Range.Rows.Count
        ReDim Preserve lItems(UBound(lItems) + 1)
        lItems(i - 1).Columna = Data.Range.Cells(i, 1)
        lItems(i - 1).Charola = Data.Range.Cells(i, 2)
        lItems(i - 1).Folio1 = Data.Range.Cells(i, 3)
        lItems(i - 1).Folio2 = Data.Range.Cells(i, 4)
    Next i
    
    lblVersion.Caption = SysVersion
End Sub

Private Sub UserForm_Terminate()
    cnData.Close
End Sub
