VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uRegBooks 
   Caption         =   "Registro de libros consultados en sala"
   ClientHeight    =   7710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15090
   OleObjectBlob   =   "uRegBooks.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "uRegBooks"
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

Dim BookCount As Integer, UserBooks() As ReadedBooks
Dim xTitulo, xCol, xCha, xSeccion, xFolio, xClasificacion As Integer, xAutor As Integer, xTAGS As Integer
Dim objSection As Object, objItems() As String

Private Type ReadedBooks
    Name As String
    section As String
End Type

Private Function LoadWebData(ID As String) As tabID
    On Error Resume Next
    ID = ParseNumber(ID)
    Dim lData() As String, buff As String
    lData = Split(Trim(ID), "-")
    If Left(lData(1), 1) = "9" Then
        lData(1) = "19" & lData(1)
    Else
        lData(1) = "20" & lData(1)
    End If
    buff = lData(1) & "-" & lData(0)
    
    LoadWebData = GetBookData(wFicha, buff)
End Function

Sub ShowError(description As String)
    wFicha.Navigate "about:blank"
    Do While wFicha.Busy Or wFicha.ReadyState <> 4
        DoEvents
    Loop
    
    Dim buff As String
    buff = ""
    wFicha.Document.Write description
End Sub

Private Sub FillExcelData(ID As Long)
    On Error Resume Next
    Dim content As ListObject

    Set content = ThisWorkbook.Sheets(tSheet).ListObjects(tTable)
    tTitulo.text = Trim(content.Range(ID, xTitulo))
    tSection.Caption = Replace(Trim(content.Range(ID, xSeccion)), Chr(10), " -> ")
    
    Dim pos As Integer
    pos = InStr(content.Range(ID, xSeccion), Chr(10))
    If pos = 0 Then _
        tSeccion.text = content.Range(ID, xSeccion)
    If pos > 0 Then _
        tSeccion.text = Left(content.Range(ID, xSeccion), pos - 1)
    
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
                If Value = "0x14" Or Value = "0xFF" Then
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
        tBack.Caption = " " & content.Range(ID - i, xClasificacion) & " | " & content.Range(ID - i, xFolio) & vbNewLine & " " & content.Range(ID - i, xTitulo) & " / " & content.Range(ID - i, xAutor)
        
        i = 1
        lSkip = False
        Do While True
            lContent = Split(content.Range(ID + i, xTAGS), ";")
            If UBound(lContent) = -1 Then _
                lSkip = False
            For Each Value In lContent
                If Value = "0x14" Or Value = "0xFF" Then
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
        tNext.Caption = " " & content.Range(ID + i, xClasificacion) & " | " & content.Range(ID + i, xFolio) & vbNewLine & " " & content.Range(ID + i, xTitulo) & " / " & content.Range(ID + i, xAutor)
    End If
End Sub

Private Sub cAdd_Click()
    If Len(tTitulo.text) = 0 Then
        MsgBox "Por favor ingresa un título para el libro consultado", vbCritical, "Registro de consultas"
        tTitulo.SetFocus
        Exit Sub
    End If
    If Len(tSeccion.text) = 0 Then
        MsgBox "Por favor ingresa la sección a la que pertenece el libro consultado", vbCritical, "Registro de consultas"
        tSeccion.SetFocus
        Exit Sub
    End If
    
    If BookCount > 0 Then
        ReDim Preserve UserBooks(UBound(UserBooks) + 1)
    End If
    UserBooks(UBound(UserBooks)).Name = Trim(tTitulo.text)
    UserBooks(UBound(UserBooks)).section = Trim(tSeccion.text)
    BookCount = BookCount + 1
    
    tColumna.Caption = ""
    tCharola.Caption = ""
    tSection.Caption = ""
    tBack.Caption = ""
    tNext.Caption = ""
    tTitulo.text = ""
    tSeccion.text = ""
    tSeccion.Clear
    wFicha.Navigate "about:blank"
    lLibros.Caption = BookCount & " libros consultados por usuario"
    
    Dim ctl As Control
    fAdd.Enabled = False
    For Each ctl In fAdd.Controls
        ' Check if the control is a valid control type that can be enabled/disabled
        ctl.Enabled = False
    Next ctl
    cAdd.Default = False
    cLocate.Default = True
    cClean.Enabled = False
    cRegister.Enabled = True
    
    tFolio.SetFocus
End Sub

Private Sub cClean_Click()
    tColumna.Caption = ""
    tCharola.Caption = ""
    tSection.Caption = ""
    tBack.Caption = ""
    tNext.Caption = ""
    tTitulo.text = ""
    tSeccion.text = ""
    tSeccion.Clear
    wFicha.Navigate "about:blank"
    
    Dim ctl As Control
    fAdd.Enabled = False
    For Each ctl In fAdd.Controls
        ' Check if the control is a valid control type that can be enabled/disabled
        ctl.Enabled = False
    Next ctl
    cAdd.Default = False
    cLocate.Default = True
    cRegister.Enabled = False
    
    tFolio.SetFocus
End Sub

Private Sub cCorregir_Click()
    MsgBox "TODO: c'est possible qu'il a beaucoup de rêve..."
End Sub

Private Sub cExit_Click()
    If BookCount > 0 Then
        If MsgBox("Al parecer tiene libros registrados sin guardar. ¿Deseas guardarlos o descartarlos?", vbCritical + vbYesNo, "Registro de consultas") = vbYes Then _
            cRegister_Click
    End If
    Unload Me
End Sub

Private Sub cLocate_Click()
    tSeccion.Clear
    
    Dim lPosition As Long, i As Integer
    If Len(tFolio.text) = 0 Then
        tFolio.SetFocus
        Exit Sub
    End If
    On Error GoTo Fail
    Dim out As tabID
    out = LoadWebData(tFolio.text)
    
    lPosition = FindExcelData(tFolio.text, xFolio)
    Dim lData As Variant, LocalData As Boolean
    If lPosition > 0 Then
        FillExcelData lPosition
        
        For Each lData In objSection.Keys
            tSeccion.AddItem lData
        Next lData
        LocalData = True
    ElseIf lPosition = 0 Then
        Dim pos As Integer
        pos = InStr(out.MARC245, "/")
        If pos = 0 Then _
            tTitulo.text = out.MARC245
        If pos > 0 Then _
            tTitulo.text = Trim(Left(out.MARC245, pos - 1))
        
        For Each lData In objItems
            tSeccion.AddItem lData
        Next lData
        LocalData = False
    End If
    
    tFolio.text = ""
    Dim ctl As Control
    fAdd.Enabled = True
    For Each ctl In fAdd.Controls
        ' Check if the control is a valid control type that can be enabled/disabled
        If Not ctl.Tag = "_oncurrent" Then _
            ctl.Enabled = True
        
        If ctl.Tag = "_oncurrent" And LocalData = True Then _
            ctl.Enabled = True
    Next ctl
    cLocate.Default = False
    cAdd.Default = True
    cClean.Enabled = True
    cRegister.Enabled = False
    
    If Len(tTitulo.text) > 0 And Len(tSeccion.text) > 0 Then _
        cAdd.SetFocus
    If Len(tTitulo.text) > 0 And Len(tSeccion.text) = 0 Then _
        tSeccion.SetFocus
    If Len(tTitulo.text) = 0 Then _
        tTitulo.SetFocus
    Exit Sub
Fail:
    ShowError (Err.description)
    tFolio.SetFocus
    tFolio.SelStart = 0
    tFolio.SelLength = Len(tFolio.text)
End Sub

Private Sub cRegister_Click()
    If BookCount = 0 Then
        MsgBox "Por favor agrege algunos libros para registrar consulta", vbCritical, "Registro de consultas"
        tFolio.SetFocus
        Exit Sub
    End If
    If tUsuarios.text = "" Then
        MsgBox "Por favor escribe cuántos usuarios realizaron estas consultas", vbCritical, "Registro de consultas"
        tUsuarios.SetFocus
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim content As ListObject

    Set content = ThisWorkbook.Sheets("Consultas").ListObjects("READS")
    
    Dim i As Integer
    For i = LBound(UserBooks) To UBound(UserBooks)
        Dim row As ListRow
        Set row = content.ListRows.Add
        row.Range(1) = Now
        row.Range(2) = UserBooks(i).Name
        row.Range(3) = UserBooks(i).section
        If i = 0 Then _
            row.Range(4) = Trim(tUsuarios.text)
        Set row = Nothing
    Next i
    
    BookCount = 0
    ReDim UserBooks(0)
    tUsuarios.text = "1"
    lLibros.Caption = "0 libros consultados por usuario"
    tColumna.Caption = ""
    tCharola.Caption = ""
    tSection.Caption = ""
    tBack.Caption = ""
    tNext.Caption = ""
    tTitulo.text = ""
    tSeccion.text = ""
    tSeccion.Clear
    wFicha.Navigate "about:blank"
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    Dim ctl As Control
    fAdd.Enabled = False
    For Each ctl In fAdd.Controls
        ' Check if the control is a valid control type that can be enabled/disabled
        ctl.Enabled = False
    Next ctl
    cAdd.Default = False
    cLocate.Default = True
    cRegister.Enabled = False
    
    tFolio.SetFocus
End Sub

Private Sub UserForm_Activate()
    tFolio.SetFocus
End Sub

Private Sub UserForm_Initialize()
    StartUpPosition = 0
    Left = Application.Left + (Application.width - width - 10)
    Top = Application.Top + (0.5 * Application.Height) - (0.5 * Height)
    
    Set cnData = New ADODB.Connection
    
    cnData.ConnectionString = ADOPathQuery
    cnData.Open
    
    Set objSection = CreateObject("Scripting.Dictionary")
    
    Dim Data As ListObject, buff As Range, content As String, pos As Integer
    
    xCol = GetPos("Columna")
    xCha = GetPos("Charola")
    xTitulo = GetPos("Título")
    xFolio = GetPos("N° de adquisición")
    xSeccion = GetPos("Área que pertenece")
    xClasificacion = GetPos("Clasificación")
    xAutor = GetPos("Autor")
    xTAGS = GetPos("TAGS")
    
    Set Data = ThisWorkbook.Sheets(tSheet).ListObjects(tTable)
    For Each buff In Data.ListColumns(xSeccion).DataBodyRange
        pos = InStr(buff.Value, Chr(10))
        If pos = 0 Then _
            content = buff.Value
        If pos > 0 Then _
            content = Left(buff.Value, InStr(buff.Value, Chr(10)) - 1)
            
        If Not objSection.Exists(content) Then
            objSection.Add content, content
        End If
    Next buff
    
    objItems = Split(GetParam("0x20"), ";")
    ReDim UserBooks(0)
    
    lblVersion.Caption = SysVersion
    
    ThisWorkbook.Save
    Me.Caption = "Registro de libros consultados en sala -- " & tSheet & " (" & ThisWorkbook.Sheets(tSheet).ListObjects(tTable).Range.Rows.Count & " libros registrados)"
End Sub

Private Sub UserForm_Terminate()
    cnData.Close
    ThisWorkbook.Save
End Sub
