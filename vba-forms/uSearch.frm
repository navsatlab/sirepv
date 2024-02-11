VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uSearch 
   Caption         =   "Búsqueda de libros"
   ClientHeight    =   9240.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19680
   OleObjectBlob   =   "uSearch.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "uSearch"
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
' Cuadro de búsqueda de información en el acervo bibliográfico

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
#End If

Private bookID() As tabID
Private xTitulo, XAutor, xCol, xCha, xTAGS, xSeccion, xClasificacion, xFolio As Long

Private Sub FillExcelData(ID As Long)
    On Error Resume Next
    Dim content As ListObject
    
    Set content = ThisWorkbook.Sheets(tSheet).ListObjects(tTable)
    tSection.Caption = Replace(Trim(content.Range(ID, xSeccion)), Chr(10), " -> ")
    
    If Not tSheet = "Dados de baja" Then
        If xCol > 0 Then tColumna.Caption = content.Range(ID, xCol)
        If xCha > 0 Then tCharola.Caption = content.Range(ID, xCha)
        
        Dim lContent() As String, i As Integer, lSkip As Boolean, Value As Variant
        i = 1
        Do While True
            lContent = Split(content.Range(ID - i, xTAGS), ";")
            If UBound(lContent) = -1 Then _
                lSkip = False
            For Each Value In lContent
                If Value = "0x14" Or Value = "0x1C" Or Value = "0xFF" Or Value = "0x1E" Then
                    lSkip = True
                Else
                    lSkip = False
                End If
            Next
            If Not lSkip Then
                Exit Do
            Else: i = i + 1
            End If
        Loop
        tBack.Caption = " " & content.Range(ID - i, xClasificacion) & " | " & content.Range(ID - i, xFolio) & vbNewLine & " " & content.Range(ID - i, xTitulo) & " / " & content.Range(ID - i, XAutor)
        
        i = 1
        lSkip = False
        Do While True
            lContent = Split(content.Range(ID + i, xTAGS), ";")
            If UBound(lContent) = -1 Then _
                lSkip = False
            For Each Value In lContent
                If Value = "0x14" Or Value = "0x1C" Or Value = "0xFF" Or Value = "0x1E" Then
                    lSkip = True
                Else
                    lSkip = False
                End If
            Next
            If Not lSkip Then
                Exit Do
            Else: i = i + 1
            End If
        Loop
        tNext.Caption = " " & content.Range(ID + i, xClasificacion) & " | " & content.Range(ID + i, xFolio) & vbNewLine & " " & content.Range(ID + i, xTitulo) & " / " & content.Range(ID + i, XAutor)
    End If
End Sub

Private Sub cClean_Click()
    tSearch.text = ""
    pList.ListItems.Clear
    ReDim bookID(0)
    wFicha.Navigate "about:blank"
    tColumna.Caption = ""
    tCharola.Caption = ""
    tSection.Caption = ""
    tBack.Caption = ""
    tNext.Caption = ""
    
    tSearch.SetFocus
End Sub

Private Sub cExit_Click()
    Unload Me
End Sub

Private Sub cSearch_Click()
    If cType.ListIndex = -1 Then
        MsgBox "Por favor selecciona un tipo válido para la búsqueda", vbCritical, "Búsqueda de libros"
        cType.SetFocus
        Exit Sub
    End If
    If Len(tSearch.text) < 3 Then
        MsgBox "La búsqueda no puede ser menor a 3 caracteres o encontrarse vacía. Por favor revisa de nuevo", vbCritical, "Búsqueda de libros"
        tSearch.SetFocus
        Exit Sub
    End If
    
    ' Start search inside DB
    Dim query As String, match As String, i As Long, rs As Recordset, Item As ListItem, MARC As String
    match = Trim(tSearch.text)
    match = Replace(match, " ", "%")
    query = "SELECT Ficha_No FROM FICHAS WHERE EtiquetasMARC LIKE '%" & match & "%';"
    
    Set rs = cnData.Execute(query)
    
    i = 0
    pList.ListItems.Clear
    ReDim bookID(0)
    wFicha.Navigate "about:blank"
    tColumna.Caption = ""
    tCharola.Caption = ""
    tSection.Caption = ""
    tBack.Caption = ""
    tNext.Caption = ""
    
    Do While Not rs.EOF
        ReDim Preserve bookID(UBound(bookID) + 1)
        bookID(UBound(bookID)) = FindData(rs.Fields(0).Value)
        
        Set Item = pList.ListItems.Add(, , bookID(UBound(bookID)).MARC245)
        Item.SubItems(1) = bookID(UBound(bookID)).MARC100
        Item.SubItems(2) = bookID(UBound(bookID)).MARC082
        Item.SubItems(3) = bookID(UBound(bookID)).ID
        
        i = i + 1
        rs.MoveNext
    Loop
    
    fItems.Caption = "Resultados de la búsqueda - " & i & " libros encontrados"
End Sub

Private Sub pList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo Fail
    wFicha.Navigate "about:blank"
    DoEvents
    
    Sleep 100
    tColumna.Caption = ""
    tCharola.Caption = ""
    tSection.Caption = ""
    tBack.Caption = ""
    tNext.Caption = ""
    
    Dim lParsed As String, lData() As String
    
    On Error Resume Next
    If Len(bookID(Item.Index).CONTENTS(0).ID) > 0 Then
        lData = Split(bookID(Item.Index).CONTENTS(0).ID, "-")
        lParsed = lData(1) & "-" & Right(lData(0), 2)
        
        FillExcelData FindExcelData(lParsed, xFolio)
    End If
    
    On Error GoTo 0
    wFicha.Document.Write bookID(Item.Index).HTMLContent
    Exit Sub
    
Fail:
    wFicha.Document.Write Err.description
    Exit Sub
End Sub

Private Sub UserForm_Activate()
    tSearch.SetFocus
End Sub

Private Sub UserForm_Initialize()
    StartUpPosition = 0
    Left = Application.Left + (Application.width - width - 10)
    Top = Application.Top + (0.5 * Application.Height) - (0.5 * Height)
    
    Set cnData = New ADODB.Connection
    
    cnData.ConnectionString = ADOPathQuery
    cnData.Open
    
    cType.AddItem "Búsqueda general"
    cType.AddItem "Título"
    cType.AddItem "Autor"
    cType.AddItem "Editorial"
    cType.AddItem "Lugar de publicación"
    cType.AddItem "Clasificación"
    cType.AddItem "Temas"
    cType.AddItem "Donante"
    cType.AddItem "ISBN"
    
    cType.ListIndex = 0
    
    pList.View = lvwReport
    pList.Gridlines = True
    pList.LabelEdit = lvwManual
    pList.FullRowSelect = True
    pList.MultiSelect = False
    pList.AllowColumnReorder = True
    
    pList.ColumnHeaders.Add , , "Título", 200
    pList.ColumnHeaders.Add , , "Autor", 150
    pList.ColumnHeaders.Add , , "Clasificación", 110
    pList.ColumnHeaders.Add , , "Ficha", 50
    
    xCol = GetPos("Columna")
    xCha = GetPos("Charola")
    xTitulo = GetPos("Título")
    xFolio = GetPos("N° de adquisición")
    xSeccion = GetPos("Área que pertenece")
    xClasificacion = GetPos("Clasificación")
    XAutor = GetPos("Autor")
    xTAGS = GetPos("TAGS")
    
    lblVersion.Caption = SysVersion
End Sub

Private Sub UserForm_Terminate()
    cnData.Close
End Sub
