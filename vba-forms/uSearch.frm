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
    Dim query As String, match As String, i As Long, rs As Recordset, Item As ListItem, MARC As String, lContent As tabID
    match = Trim(tSearch.text)
    match = Replace(match, " ", "%")
    query = "SELECT Ficha_No FROM FICHAS WHERE EtiquetasMARC LIKE '%" & match & "%';"
    
    Set rs = cnData.Execute(query)
    
    i = 0
    pList.ListItems.Clear
    wFicha.Navigate "about:blank"
    Do While Not rs.EOF
        lContent = FindData(rs.Fields(0).Value)
        
        Set Item = pList.ListItems.Add(, , lContent.MARC245)
        Item.SubItems(1) = lContent.MARC100
        Item.SubItems(2) = lContent.MARC082
        Item.SubItems(3) = lContent.ID
        i = i + 1
        rs.MoveNext
    Loop
    
    fItems.Caption = "Resultados de la búsqueda - " & i & " libros encontrados"
End Sub

Private Sub pList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo Fail
    LoadBookData wFicha, Item.SubItems(3)
    Exit Sub
    
Fail:
    
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
    
    lblVersion.Caption = SysVersion
End Sub

Private Sub UserForm_Terminate()
    cnData.Close
End Sub
