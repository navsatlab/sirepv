VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uCompare 
   Caption         =   "Cotejo y verificación"
   ClientHeight    =   10650
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18540
   OleObjectBlob   =   "uCompare.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "uCompare"
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
Private lReset As Boolean, IsOutside As Boolean, lDataModified As Boolean

Private xTitulo, XAutor, xEditorial, xDonante, xAño, xPais, xClasificacion, xFolio, xNotas, xCol, xCha, xSeccion, xIdioma, xTAGS As Integer
Private nFicha As String, nISBN As String

Private Sub FillExcelData(ID As Long)
    'On Error Resume Next
    Dim content As ListObject

    Set content = ThisWorkbook.Sheets(tSheet).ListObjects(tTable)
    tTitulo.text = Trim(content.Range(ID, xTitulo))
    tAutor.text = Trim(content.Range(ID, XAutor))
    tPais.text = Trim(content.Range(ID, xPais))
    tEditorial.text = Trim(content.Range(ID, xEditorial))
    tAño.text = Trim(content.Range(ID, xAño))
    tClasificacion.text = Trim(content.Range(ID, xClasificacion))
    tFolio.text = Trim(content.Range(ID, xFolio))
    tDonante.text = Trim(content.Range(ID, xDonante))
    tNotas.text = Trim(content.Range(ID, xNotas))
    tSection.Caption = Replace(Trim(content.Range(ID, xSeccion)), Chr(10), " -> ")
    
    If Not tIdioma.Tag = "_unabled" Then _
        tIdioma.text = Trim(content.Range(ID, xIdioma))
    
    If Not tSheet = "Dados de baja" Then
        If xCol > 0 Then tColumna.Caption = content.Range(ID, xCol)
        If xCha > 0 Then tCharola.Caption = content.Range(ID, xCha)
        
        Dim lData() As String, Value As Variant
        lData = Split(content.Range(ID, xTAGS), ";")
        ch10.Value = False
        ch12.Value = False
        ch14.Value = False
        chFF.Value = False
        ch1A.Value = False
        ch1C.Value = False
        ch1E.Value = False
        
        For Each Value In lData
            If Value = "0x10" Then      ' CI
                ch10.Value = True
            ElseIf Value = "0x12" Then  ' Para restaurar
                ch12.Value = True
            ElseIf Value = "0x1E" Then  ' Libro de Gran Formato
                ch1E.Value = True
            ElseIf Value = "0x1A" Then  ' Libro con errores en ficha
                ch1A.Value = True
            ElseIf Value = "0x1C" Then  ' En restauración
                ch1C.Value = True
            ElseIf Value = "0x14" Then  ' En catalogación
                ch14.Value = True
            ElseIf Value = "0xFF" Then  ' Perdido
                chFF.Value = True
            End If
        Next
        RedrawLabel
    
        Dim lContent() As String, i As Integer, lSkip As Boolean
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
    
    If cBack.Enabled = False Then cBack.Enabled = True
    If cNext.Enabled = False Then cNext.Enabled = True
    cSelect.Enabled = True
    
    If lPosition <= 2 Then
        cBack.Enabled = False
        tBack.Caption = " Libro al principio del inventario"
    End If
    If lPosition >= ThisWorkbook.Sheets(tSheet).ListObjects(tTable).Range.Rows.Count Then
        cNext.Enabled = False
        tNext.Caption = " Libro al final del inventario"
    End If
    'MsgBox ArrangeLC(content.Range(ID, xClasificacion), content.Range(ID - 1, xClasificacion))
End Sub

Private Sub RedrawLabel()
    lState.Visible = False
    lState.Caption = ""
    lState.ForeColor = RGB(0, 0, 0)
    lState.BackColor = RGB(255, 255, 255)
    
    If ch10.Value Then      ' CI
        lState.ForeColor = RGB(255, 0, 0)
        lState.Visible = True
        lState.Caption = " Libro de Consulta Interna" & vbNewLine & _
            " Libro que no sale para préstamo a domicilio"
    End If
    If ch12.Value Then  ' Para restaurar
        lState.BackColor = RGB(255, 255, 0)
        lState.Visible = True
        lState.Caption = " Libro que necesita restauración" & vbNewLine & _
            " Por favor especifica en las Notas el diagnóstico del libro"
    End If
    If ch10.Value And ch12.Value Then ' CI + Para restaurar
        lState.BackColor = RGB(255, 255, 0)
        lState.ForeColor = RGB(255, 0, 0)
        lState.Caption = " Libro de Consulta Interna que necesita restauración" & vbNewLine & _
            " Por favor especifica en las Notas el diagnóstico del libro"
        lState.Visible = True
    End If
    If ch1E.Value Then  ' Libro de Gran Formato
        lState.BackColor = RGB(230, 230, 250)
        lState.Visible = True
        lState.Caption = " Libro de Gran Formato" & vbNewLine & _
            " Ubicado en otra área designada por sus dimensiones"
    End If
    If ch1A.Value Then  ' Libro con errores en ficha
        lState.BackColor = RGB(204, 255, 255)
        lState.Visible = True
        lState.Caption = " Libro con posibles errores en ficha" & vbNewLine & _
            " Por favor verifica los datos de la ficha catalográfica"
    End If
    If ch1C.Value Then  ' En restauración
        lState.BackColor = RGB(146, 208, 80)
        lState.Visible = True
        lState.Caption = " Libro en restauración" & vbNewLine & _
            " Libro fuera de charola que se está restaurando"
    End If
    If ch14.Value Then   ' En catalogación
        lState.BackColor = RGB(51, 51, 0)
        lState.ForeColor = RGB(255, 255, 255)
        lState.Visible = True
        lState.Caption = " El libro se encuentra actualmente en catalogación"
    End If
    If chFF.Value Then   ' Perdido
        lState.BackColor = RGB(128, 0, 0)
        lState.ForeColor = RGB(255, 255, 255)
        lState.Visible = True
        lState.Caption = " El libro se encuentra perdido / no ha sido encontrado"
    End If
End Sub

Private Sub SaveExcelData(ID As Long)
    Dim content As ListObject

    Set content = ThisWorkbook.Sheets(tSheet).ListObjects(tTable)
    content.Range(ID, xTitulo) = Trim(tTitulo.text)
    content.Range(ID, XAutor) = Trim(tAutor.text)
    content.Range(ID, xPais) = Trim(tPais.text)
    content.Range(ID, xEditorial) = Trim(tEditorial.text)
    content.Range(ID, xAño) = Trim(tAño.text)
    content.Range(ID, xClasificacion) = Trim(tClasificacion.text)
    content.Range(ID, xFolio) = Trim(tFolio.text)
    content.Range(ID, xDonante) = Trim(tDonante.text)
    content.Range(ID, xNotas) = Trim(tNotas.text)
    If Not tIdioma.Tag = "_unabled" Then _
        content.Range(ID, xIdioma) = Trim(tIdioma.text)
    
    If tPais.ListIndex = -1 Then _
        tPais.AddItem tPais.text
    
    If tEditorial.ListIndex = -1 Then _
        tEditorial.AddItem tEditorial.text
    
    If tDonante.ListIndex = -1 Then _
        tDonante.AddItem tDonante.text
    
    If tIdioma.ListIndex = -1 Then _
        tIdioma.AddItem tIdioma.text
    
    If Not tSheet = "Dados de baja" Then
        Dim i As Integer
        
        On Error Resume Next
        Dim lData As String
        For i = xCol To xSeccion
            content.Range(ID, i).Interior.ColorIndex = 0
            content.Range(ID, i).Font.ColorIndex = 1
        Next i
        If ch10.Value Then
            lData = lData & "0x10;" ' CI
            For i = xCol To xSeccion
                content.Range(ID, i).Font.ColorIndex = 3
            Next i
        End If
        If ch12.Value Then
            lData = lData & "0x12;" ' Para restaurar
            For i = xCol To xSeccion
                content.Range(ID, i).Interior.ColorIndex = 6
            Next i
        End If
        If ch1E.Value Then
            lData = lData & "0x1E;" ' Libro de Gran Formato
            For i = xCol To xSeccion
                content.Range(ID, i).Interior.Color = rgbLavender
            Next i
        End If
        If ch1A.Value Then
            lData = lData & "0x1A;" ' Libro con errores en ficha
            For i = xCol To xSeccion
                content.Range(ID, i).Interior.Color = rgbPaleTurquoise
            Next i
        End If
        If ch1C.Value Then
            lData = lData & "0x1C;" ' En restauración
            For i = xCol To xSeccion
                content.Range(ID, i).Interior.Color = rgbYellowGreen
            Next i
        End If
        If ch14.Value Then
            lData = lData & "0x14;" ' En catalogación
            For i = xCol To xSeccion
                content.Range(ID, i).Interior.ColorIndex = 52
                content.Range(ID, i).Font.ColorIndex = 2
            Next i
        End If
        If chFF.Value Then
            lData = lData & "0xFF;" ' Perdido
            For i = xCol To xSeccion
                content.Range(ID, i).Font.ColorIndex = 2
                content.Range(ID, i).Interior.ColorIndex = 9
            Next i
        End If
        
        lData = Left(lData, Len(lData) - 1)
        content.Range(ID, xTAGS) = lData
    End If
End Sub

Sub LoadWebData(ID As String)
    On Error Resume Next
    ID = ParseNumber(ID)
    Dim lData() As String, buff As String
    lData = Split(Trim(ID), "-")
    If UBound(lData) > 0 Then
        If Left(lData(1), 1) = "9" Then
            lData(1) = "19" & lData(1)
        Else
            lData(1) = "20" & lData(1)
        End If
        buff = lData(1) & "-" & lData(0)
    Else
        buff = ID
    End If
    
    Dim X As tabID
    
    X = GetBookData(wFicha, buff)
    lCreated.Caption = X.DateInfo.Created
    lModified.Caption = X.DateInfo.Modified
    
    'Prefill var with loaded info, to start Google Forms with prefilled data
    nFicha = Trim(X.ID)
    nISBN = Trim(Replace(X.MARC020, "-", ""))
End Sub

Sub ShowError(description As String)
    wFicha.Navigate "about:blank"
    Do While wFicha.Busy Or wFicha.ReadyState <> 4
        DoEvents
    Loop
    
    Dim buff As String
    buff = ""
    wFicha.Document.Write description
End Sub

Private Sub ch10_Click()
    lDataModified = True
    RedrawLabel
End Sub

Private Sub ch12_Click()
    lDataModified = True
    RedrawLabel
End Sub

Private Sub ch1A_Click()
    lDataModified = True
    RedrawLabel
End Sub

Private Sub ch14_Click()
    lDataModified = True
    RedrawLabel
End Sub

Private Sub ch1C_Click()
    lDataModified = True
    RedrawLabel
End Sub

Private Sub ch1E_Click()
    lDataModified = True
    RedrawLabel
End Sub

Private Sub chFF_Click()
    lDataModified = True
    RedrawLabel
End Sub

Private Sub cmdSave_Click()
On Error GoTo Fail
    Application.ScreenUpdating = False
    SaveExcelData (lPosition)
    FillExcelData (lPosition)
    
    lDataModified = False
    
    Application.ScreenUpdating = True
    Exit Sub
Fail:
    ShowError (Err.description)
    Application.ScreenUpdating = True
End Sub

Private Sub cNext_Click()
On Error GoTo Fail
    Application.ScreenUpdating = False
    If cNext.Default = False Then
        cNext.Default = True
        cNext.Caption = "Siguiente [Enter] ->"
        cBack.Caption = "<- Atrás"
    End If
    Dim X As Object
    Set X = fData.ActiveControl
    If lPosition >= ThisWorkbook.Sheets(tSheet).ListObjects(tTable).Range.Rows.Count Then Exit Sub
    SaveExcelData lPosition

    lPosition = lPosition + 1
    FillExcelData lPosition
    LoadWebData (tFolio.text)
    
    On Error Resume Next
    X.SetFocus
    X.SelStart = 0
    X.SelLength = Len(X.text)
    
    lDataModified = False
    
    Application.ScreenUpdating = True
    Exit Sub
    
Fail:
    ShowError (Err.description)
    Application.ScreenUpdating = True
    Exit Sub
    
End Sub

Private Sub cBack_Click()
On Error GoTo Fail
    Application.ScreenUpdating = False
    If cBack.Default = False Then
        cBack.Default = True
        cNext.Caption = "Siguiente ->"
        cBack.Caption = "<- Atrás [Enter]"
    End If
    Dim X As Object
    Set X = fData.ActiveControl
    If lPosition <= 2 Then Exit Sub
    SaveExcelData lPosition
    
    lPosition = lPosition - 1
    FillExcelData lPosition
    LoadWebData (tFolio.text)
    
    On Error Resume Next
    X.SetFocus
    X.SelStart = 0
    X.SelLength = Len(X.text)
    
    lDataModified = False
    
    Application.ScreenUpdating = True
    Exit Sub
    
Fail:
    ShowError (Err.description)
    Application.ScreenUpdating = True
    Exit Sub
    
End Sub

Private Sub cCancel_Click()
    If lReset = False Then
        Unload Me
    Else
        ' Verificamos si hay cambios para guardar previos
        Dim result As VbMsgBoxResult
        'If lDataModified And IsOutside = False Then
            'If MsgBox("Al parecer se efectuaron algunos cambios realizados." & vbNewLine & "¿Deseas guardarlos?", vbQuestion + vbYesNo + vbDefaultButton1, "Cotejo de inventario") = vbYes Then
                'SaveExcelData (lPosition)
            'End If
        'End If
        IsOutside = False
        wFicha.Navigate "about:blank"
        cBack.Enabled = False
        cNext.Enabled = False
        cSelect.Enabled = False
        
        Dim ctl As Control
        fData.Enabled = False
        For Each ctl In fData.Controls
            ' Check if the control is a valid control type that can be enabled/disabled
            ctl.Enabled = False
            If TypeOf ctl Is MSForms.TextBox Or TypeOf ctl Is MSForms.ComboBox Then
                ctl.text = ""
            End If
            If TypeOf ctl Is MSForms.CheckBox Then
                ctl.Value = False
            End If
        Next ctl
        
        lCreated.Caption = ""
        lModified.Caption = ""
        
        tColumna.Caption = ""
        tCharola.Caption = ""
        tBack.Caption = ""
        tNext.Caption = ""
        tSection.Caption = ""
        
        cLocate.Default = True
        cCancel.Caption = "Cerrar [ESC]"
        cLocate.Caption = "Buscar [Enter]"
        cNext.Caption = "Siguiente ->"
        cBack.Caption = "<- Atrás"
        lReset = False
        lDataModified = False
        tID.SetFocus
        tID.SelStart = 0
        tID.SelLength = Len(tID.text)
    End If
End Sub

Private Sub cLocate_Click()
On Error GoTo Fail
    Application.ScreenUpdating = False
    
    ' First load db data
    LoadWebData tID.text
    
    lPosition = FindExcelData(tID.text, xFolio)
    If Len(CompareItems) = 0 Then
        CompareItems = tID.text
    Else
        CompareItems = CompareItems & ";" & tID.text
    End If
    tID.AddItem tID.text
    
    If lPosition = 0 Then _
        GoTo RenderOutside
        
    IsOutside = False
    ' And then, fills with excel variables
    FillExcelData lPosition
    
    cLocate.Default = False
    cNext.Default = True
    lReset = True
    cLocate.Caption = "Buscar"
    cCancel.Caption = "Limpiar [ESC]"
    cNext.Caption = "Siguiente [Enter] ->"
    
    Dim ctl As Control
    fData.Enabled = True
    For Each ctl In fData.Controls
        ' Check if the control is a valid control type that can be enabled/disabled
        If Not ctl.Tag = "_unabled" Then _
            ctl.Enabled = True
    Next ctl
    
    tTitulo.SetFocus
    lDataModified = False
    Application.ScreenUpdating = True
    Exit Sub

Fail:
    ShowError (Err.description)
    Application.ScreenUpdating = True
    Exit Sub

RenderOutside:
    lReset = True
    IsOutside = True
    cCancel.Caption = "Limpiar [ESC]"
    
    tID.SetFocus
    tID.SelStart = 0
    tID.SelLength = Len(tID.text)
    Application.ScreenUpdating = True
    Exit Sub
    
End Sub

Private Sub cSelect_Click()
    Sheets(tSheet).Activate
    Sheets(tSheet).Range("$D$" & lPosition).Select
End Sub

Private Sub cSugest_Click()
    tNotas.SetFocus
    
    Dim lData As DataContainer, lReturn As DataContainer
    lData.Titulo = tTitulo.text
    lData.Autor = tAutor.text
    lData.Año = tAño.text
    lData.Clasificacion = tClasificacion.text
    lData.Donante = tDonante.text
    lData.Editorial = tEditorial.text
    lData.Idioma = tIdioma.text
    lData.Lugar = tPais.text
    
    'lReturn = LoadModify(lData)
    
    ' Most of all GForms entry's can be founded in FB_PUBLIC_LOAD_DATA_ but if it's changes, try to retrieve it using REGEX
    ' 1531173234 bibliotecario
    ' 1250767358 date
    ' 1295936122 sala a cargo
    ' 1077648177 n° ficha
    ' 215912048 isbn
    ' 798850835 lugar de edición
    ' 1968072427 donante
    
    ' Base URL
    ' https://docs.google.com/forms/d/e/1FAIpQLSeMC0Ox2AVFiQ9RKPmIvTZRV9nf_ZcqXUtrOoKxj9vNrifxgg/viewform?
    
    Dim baseURL As String
    baseURL = "https://docs.google.com/forms/d/e/1FAIpQLSe11o_DMn2XyqSEYlNoNxC1h5HXjoH2hmCLE-8omLI9y-GSyw/viewform?"
    
    Dim entryUser, entrySala, entryNFicha, entryISBN, entryPais, entryDonante, entryDate As String
    entryUser = "&entry.1531173234="    ' Bibliotecario
    entrySala = "&entry.1295936122="    ' Sala a cargo
    entryNFicha = "&entry.1077648177="  ' Número de ficha
    entryISBN = "&entry.215912048="     ' ISBN
    entryPais = "&entry.798850835="     ' Lugar de edición
    entryDonante = "&entry.1968072427=" ' Donante
    'entryDate = "&entry.1250767358_"   ' Fecha
    
    If lUser = "" Then _
        lUser = InputBox("Por favor escribe tu nombre, se usará para rellenar los datos de Google Forms", "Cotejo / Verificación")
       
    If lUser = "" Then Exit Sub
    baseURL = baseURL & entryUser & lUser & entrySala & GetParam("Name") & entryNFicha & nFicha & _
        entryISBN & nISBN & entryPais & Trim(tPais.text) & entryDonante & "Donación : " & Trim(tDonante.text)
    
    'baseURL = baseURL & "entry.1531173234=" & lUser & "&entry.1295936122=" & GetParam("Name") & _
        "&entry.1077648177=" & nFicha & "&entry.215912048=" & nISBN & "&entry.798850835=" & _
        Trim(tPais.text) & "&entry.1968072427=Donación : " & Trim(tDonante.text) & _
        "&entry.1250767358_day=" & Day(Now) & "&entry.1250767358_month=" & Month(Now) & _
        "&entry.1250767358_year=" & Year(Now)
    
    ' Execute URL with prefilled data
    
    ThisWorkbook.FollowHyperlink (baseURL)
End Sub

Private Sub tAño_AfterUpdate()
    lDataModified = True
End Sub

Private Sub tAutor_AfterUpdate()
    lDataModified = True
End Sub

Private Sub tClasificacion_AfterUpdate()
    lDataModified = True
End Sub

Private Sub tDonante_AfterUpdate()
    lDataModified = True
End Sub

Private Sub tEditorial_AfterUpdate()
    lDataModified = True
End Sub

Private Sub tFolio_AfterUpdate()
    lDataModified = True
End Sub

Private Sub tIdioma_AfterUpdate()
    lDataModified = True
End Sub

Private Sub tNotas_AfterUpdate()
    lDataModified = True
End Sub

Private Sub tPais_AfterUpdate()
    lDataModified = True
End Sub

Private Sub tTitulo_AfterUpdate()
    lDataModified = True
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
    
    Dim obEditorial As Object, obDonante As Object, obPais As Object, obIdioma As Object
    Set obEditorial = CreateObject("Scripting.Dictionary")
    Set obDonante = CreateObject("Scripting.Dictionary")
    Set obPais = CreateObject("Scripting.Dictionary")
    Set obIdioma = CreateObject("Scripting.Dictionary")
    
    Dim Data As ListObject
    Dim buff As Range
    Dim lData As Variant
    
    If Not tSheet = "Dados de baja" Then
        If tSheet = "Libros en sala" Then
            xCol = GetPos("Columna")
            xCha = GetPos("Charola")
        End If
        xTAGS = GetPos("TAGS")
        xIdioma = GetPos("Idiomas")
    End If
    
    xTitulo = GetPos("Título")
    XAutor = GetPos("Autor")
    xPais = GetPos("Lugar de Edición")
    xEditorial = GetPos("Editorial")
    xAño = GetPos("Año de edición")
    xClasificacion = GetPos("Clasificación")
    xFolio = GetPos("N° de adquisición")
    xDonante = GetPos("Donante")
    xNotas = GetPos("Notas")
    xSeccion = GetPos("Área que pertenece")
    
    Set Data = ThisWorkbook.Sheets(tSheet).ListObjects(tTable)
    For Each buff In Data.ListColumns(xPais).DataBodyRange
        If Not obPais.Exists(buff.Value) Then
            obPais.Add buff.Value, buff.Value
        End If
    Next buff
    For Each buff In Data.ListColumns(xEditorial).DataBodyRange
        If Not obEditorial.Exists(buff.Value) Then
            obEditorial.Add buff.Value, buff.Value
        End If
    Next buff
    For Each buff In Data.ListColumns(xDonante).DataBodyRange
        If Not obDonante.Exists(buff.Value) Then
            obDonante.Add buff.Value, buff.Value
        End If
    Next buff
    
    If Not tSheet = "Dados de baja" Then
        For Each buff In Data.ListColumns(xIdioma).DataBodyRange
            If Not obIdioma.Exists(buff.Value) Then
                obIdioma.Add buff.Value, buff.Value
            End If
        Next buff
        
        For Each lData In obIdioma.Keys
            tIdioma.AddItem lData
        Next lData
    Else
        tIdioma.Tag = "_unabled"
        ch10.Tag = "_unabled"
        ch12.Tag = "_unabled"
        ch14.Tag = "_unabled"
        ch1C.Tag = "_unabled"
        ch1E.Tag = "_unabled"
        chFF.Tag = "_unabled"
    End If
    
    For Each lData In obEditorial.Keys
        tEditorial.AddItem lData
    Next lData
    For Each lData In obPais.Keys
        tPais.AddItem lData
    Next lData
    For Each lData In obDonante.Keys
        tDonante.AddItem lData
    Next lData
    
    Dim lSearchedItems() As String
    If Len(CompareItems) > 0 Then
        lSearchedItems = Split(CompareItems, ";")
        Dim i As Long
        For i = LBound(lSearchedItems) To UBound(lSearchedItems)
            tID.AddItem lSearchedItems(i)
        Next i
    End If

    Me.Caption = "Cotejo y verificación -- " & tSheet & " (" & ThisWorkbook.Sheets(tSheet).ListObjects(tTable).Range.Rows.Count & " libros registrados)"
    
    lblVersion.Caption = SysVersion
    
    ThisWorkbook.Save
    Application.Calculation = xlCalculationManual
End Sub

Private Sub UserForm_Terminate()
    cnData.Close
    
    ThisWorkbook.Save
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
