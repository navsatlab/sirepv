Attribute VB_Name = "adoconnection"
' ###################################################################
' # NAVSATLAB                                                       #
' # Application deployed over VBA                                   #
' # https://github.com/navsatlab/sirepv                             #
' # Some rights reserved                                            #
' ###################################################################

Option Explicit
' Funciones específicas de la conexión ADO

Global tSheet As String ' = "Narrativa"
Global tTable As String ' = "NARRATIVA"
Global Const mdbPath As String = "C:\PROMETEO\PROMETEO.mdb"
Global Const ADOPathQuery As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & mdbPath & ";"
Global cnData As ADODB.Connection

Public Type contentsID
    number As String 'Ejemplar
    VOLUME As String 'Volumen, obviamente
    TOME As String ' Tomo, por si fuera obvio
    ID As String ' El número de folio
End Type

Public Type dataInfo
    Created As String   ' Fecha
    Modified As String  ' FechaMod
End Type

Public Type tabID
    MARC082 As String 'CLASIFICACION
    MARC100 As String 'AUTOR
    MARC245 As String 'TITULO
    MARC250 As String 'N° EDIDICÓN
    MARC260 As String 'PAIS Y EDIT
    MARC300 As String 'PAGINAS Y DIMENSIONES
    MARC008 As String 'AÑO, HAY OTRA VARIABLE QUE CONTIENE EL AÑO
    MARC440 As String 'COLECCION Y NUMERO
    MARC500 As String 'CONTENIDOS, DONANTE, ETC
    MARC020 As String 'ISBN
    MARC650 As String 'TAGS1
    MARC700 As String 'TAGS2
    IsRepeated As Boolean
    CONTENTTAGS() As String
    CONTENTS() As contentsID
    ContentEmpty As Boolean
    ID As String
    HTMLContent As String
    DateInfo As dataInfo
    Valid As Boolean
End Type

' Convierte del formato 1-23 a 2023-1
Public Function InverseID(ByVal ID As String) As String
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
    
    InverseID = buff
End Function

' Devuelve un HTML de una ficha determinada, cargándola además en un objeto WebBrowser
Public Function GetBookData(ByVal container As WebBrowser, ID As String) As tabID
    Dim result As tabID
    
    result = FindData(ID)
    If result.IsRepeated Then _
        MsgBox "Advertencia: Posiblemente el folio del libro está siendo utilizado en uno o más libros diferentes. Por favor revisa que sea correcto", vbExclamation, "Cotejo"
    
    container.Navigate "about:blank"
    DoEvents
    
    container.Document.Write result.HTMLContent
    GetBookData = result
End Function

' Carga únicamente una ficha determinada en un WebBrowser
Public Function LoadBookData(ByVal container As WebBrowser, ID As String) As dataInfo
    container.Navigate "about:blank"
    DoEvents
    
    Dim result As tabID
    result = FindData(ID)
    If result.IsRepeated Then _
        MsgBox "Advertencia: Posiblemente el folio del libro está siendo utilizado en uno o más libros diferentes. Por favor revisa que sea correcto", vbExclamation, "Cotejo"
    container.Document.Write result.HTMLContent
    LoadBookData = result.DateInfo
End Function

' Only set HTML readable code for rows
Private Function returnAsHTML(ID As tabID) As String
    Dim struct As String
    ' HTML style
    struct = "<!DOCTYPE html><html><head><style type='text/css'>"
    struct = struct & ".margen1 { font-family: arial; font-size: 10pt; color: #000000; text-indent: -10px; margin-left: 30px; text-align: justify; }"
    struct = struct & ".margen2 { font-family: arial; font-size: 10pt; color: #000055; text-indent: 20px; margin-left: 20px; text-align: justify; }"
    struct = struct & ".margen3 { font-family: arial; font-size: 10pt; color: #000055; text-indent: 20px; margin-left: 20px; text-align: justify; font-weight: bold; }"
    struct = struct & "table.numadq { empty-cells: show; font-size: 9pt; font-family: arial; text-align: center; }"
    struct = struct & "</style></head><body><table width='100%' border='0' cellpadding='4' cellspacing='0' bgcolor='#ffde00'>"
    
    ' Encabezado Tarjeta + FOLIO + Clasificación
    struct = struct & "<tr><td> Tarjeta <font size='3' face='Arial' color='red'><b>" & ID.ID & "</b></font></td><td>"
    struct = struct & "<b>" & ID.MARC082 & "</b></td></tr></table>"
    
    ' Autor
    struct = struct & "<div class='margen1'>" & ID.MARC100 & "</div>"
    
    ' Título
    struct = struct & "<div class='margen2'>" & ID.MARC245
    
    ' Edición
    If Len(ID.MARC250) Then
        struct = struct & " -- " & ID.MARC250
    End If
    
    ' Lugar de edición
    struct = struct & " -- " & ID.MARC260
    
    ' Año
    If Len(ID.MARC008) Then
        struct = struct & ", " & ID.MARC008
    End If
    
    ' Páginas y dimensiones
    struct = struct & "</div><div class='margen2'>" & Replace(ID.MARC300, ";", "")
    
    ' Contenido del libro
    If UBound(ID.CONTENTTAGS) > 0 Then
        Dim lLetter As Variant
        struct = struct & ";"
        For Each lLetter In ID.CONTENTTAGS
            Select Case lLetter
                Case "a": struct = struct & " Il."
                Case "b": struct = struct & " Map."
                Case "c": struct = struct & " Retrs."
                Case "d": struct = struct & " Fot."
                Case "e": struct = struct & " Plans."
                Case "f": struct = struct & " Lamns."
                Case "g": struct = struct & " Música"
                Case "h": struct = struct & " Facsímiles"
                Case "i": struct = struct & " Diagrs."
                Case "j": struct = struct & " Grabs."
                Case "k": struct = struct & " Litograf."
                Case "l": struct = struct & " Discos"
                Case "m": struct = struct & " Gráficas"
                Case "n": struct = struct & " Tablas"
                Case "p": struct = struct & " Iluminaciones"
                Case "q": struct = struct & " Diskettes"
                Case "r": struct = struct & " Tablas genealógicas"
                Case "s": struct = struct & " Diapos."
                Case "t": struct = struct & " Formas y formularios"
                Case "u": struct = struct & " Muestras"
            End Select
            If Not lLetter = " " Or Len(lLetter) = 0 Then _
                struct = struct & ","
        Next
        struct = Left(struct, Len(struct) - 1)
    End If
    
    ' Colección
    If Len(ID.MARC440) > 0 Then
        struct = struct & " -- (" & ID.MARC440 & ")"
    End If
    
    ' Detalles de libro
    struct = struct & "</div>"
    Dim lData() As String, i As Integer
    lData = Split(Replace(ID.MARC500, vbNewLine, "\"), "\")
    For i = LBound(lData) To UBound(lData)
        struct = struct & "<div class='margen2'>" & Trim(lData(i)) & "</div>"
    Next i
    
    ' ISBN
    If Len(ID.MARC020) > 0 Then
        Dim isbn() As String, U As Integer
        isbn = Split(ID.MARC020, "\")
        For U = LBound(isbn) To UBound(isbn)
            struct = struct & "<div class='margen2'>ISBN " & Trim(isbn(U)) & "</div>"
        Next U
    End If
    
    ' Temas
    struct = struct & "<div class='margen3'>" & ID.MARC650 & "</div>"
    struct = struct & "<div class='margen3'>" & ID.MARC700 & "</div>"
    'Struct = Struct & "<div class='margen3'><br></div>"
    
    ' Tabla de folios
    struct = struct & "<div align='right'><table class='numadq' border='1' cellpadding='1' cellspacing='0' bgcolor='#ddddff'>"
    struct = struct & "<tr><th>Núm. Adquisición</th><th>Biblioteca</th><th>Ejemplar</th><th>Volumen</th><th>Tomo</th></tr>"
    Dim lInitialized As Boolean
    lInitialized = False
    On Error Resume Next
    lInitialized = IsNumeric(UBound(ID.CONTENTS))
    On Error GoTo 0
    
    If lInitialized Then
        For i = LBound(ID.CONTENTS) To UBound(ID.CONTENTS)
            struct = struct & "<tr><td>" & ID.CONTENTS(i).ID & "</td>"
            struct = struct & "<td>1</td>"
            struct = struct & "<td>" & ID.CONTENTS(i).number & "</td>"
            struct = struct & "<td>" & ID.CONTENTS(i).VOLUME & "</td>"
            struct = struct & "<td>" & ID.CONTENTS(i).TOME & "</td></tr>"
        Next i
    End If
    struct = struct & "</table></div></body></html>"
    
    returnAsHTML = struct
End Function

' This only search and locate specified data into dataset
Public Function FindData(ID As String) As tabID
    Dim query As String, dataArray() As String, idData As Variant
    Dim i As Integer, j As Integer
    Dim lData As tabID, strMARC() As String, strISBN As String, strAnee As String, strData As dataInfo
    Dim buffer As String
    Dim rs As Recordset
    Dim X() As String
    
    ' Cargamos la ficha, si se especifica por folio o por número de ficha
    X = Split(ID, "-")
    If UBound(X) = 1 Then
        query = "SELECT Ficha_No FROM Ejemplares WHERE NumAdqui = " & "'" & ID & "'" & ";"
    ElseIf UBound(X) = 0 Then
        query = "SELECT Ficha_No FROM Ejemplares WHERE Ficha_No = " & ID & ";"
    End If
    Set rs = cnData.Execute(query)
    
    i = 0
    Do While Not rs.EOF
        idData = Trim(rs.Fields(0).Value) ' You need to replace 0 with the index of the desired column
        i = i + 1
        rs.MoveNext
    Loop
    If i > 1 Then lData.IsRepeated = True
    rs.Close
    
    If Len(idData) = 0 Then
        lData.ContentEmpty = True
        If UBound(X) = 0 Then _
            idData = ID
    End If
    
    ' idData contiene el número de ficha que hay que buscar, por lo que ejecutamos de nuevo la búsqueda de todos los que corresponden con esa ficha
    query = "SELECT NumAdqui , Ejemplar , Volumen , Tomo FROM Ejemplares WHERE Ficha_No = " & idData & ";"
    Set rs = cnData.Execute(query)
    
    Dim lContents() As contentsID
    i = 0
    On Error Resume Next
    Do While Not rs.EOF
        ReDim Preserve lContents(i)
        lContents(i).ID = Trim(rs.Fields(0).Value)
        lContents(i).number = Trim(rs.Fields(1).Value)
        lContents(i).VOLUME = Trim(rs.Fields(2).Value)
        lContents(i).TOME = Trim(rs.Fields(3).Value)
        i = i + 1
        rs.MoveNext
    Loop
    On Error GoTo 0
    
    ' Buscamos los valores para hacer split de los contenidos reales de la ficha
    query = "SELECT EtiquetasMARC , ISBN , DatosFijos, Fecha, FechaMod FROM FICHAS WHERE Ficha_No = " & idData & ";"
    Set rs = cnData.Execute(query)
    
    lData.CONTENTS = lContents
    i = 0
    Do While Not rs.EOF
        strMARC = Split(rs.Fields(0).Value, "¦")
        strISBN = Trim(rs.Fields(1).Value)
        strAnee = Trim(rs.Fields(2).Value)
        strData.Created = Trim(rs.Fields(3).Value)
        strData.Modified = Trim(rs.Fields(4).Value)
        i = i + 1
        rs.MoveNext
    Loop
    lData.DateInfo = strData
    If i = 0 Then Err.Raise 15, , "Il n'y a pas d'information pour voir. Désolée."
    
    ' Loads all MARC data
    On Error Resume Next
    For i = LBound(strMARC) + 1 To UBound(strMARC)
        Dim ptrLeft As String, ptrRight
        ptrLeft = Left(strMARC(i), 3)
        ptrRight = Trim(Right(strMARC(i), Len(strMARC(i)) - 3))
        Select Case ptrLeft
            Case "100"
                lData.MARC100 = ptrRight
            Case "020"
                lData.MARC020 = ptrRight
            Case "082"
                lData.MARC082 = ptrRight
            Case "245"
                lData.MARC245 = ptrRight
            Case "250"
                lData.MARC250 = ptrRight
            Case "260"
                lData.MARC260 = ptrRight
            Case "300"
                lData.MARC300 = ptrRight
            Case "440"
                lData.MARC440 = ptrRight
            Case "500"
                lData.MARC500 = ptrRight
            Case "650"
                buffer = ""
                X = Split(ptrRight, "\")
                
                If Not Left(Trim(X(j)), 2) = "1." Then
                    For j = LBound(X) To UBound(X)
                        buffer = buffer & (j + 1) & "." & Trim(X(j)) & " "
                    Next j
                    ReDim X(0)
                    lData.MARC650 = buffer
                Else
                    lData.MARC650 = Trim(X(0))
                End If
            Case "700"
                buffer = ""
                X = Split(ptrRight, "\")
                
                If Not Left(Trim(X(j)), 2) = "I." Then
                    For j = LBound(X) To UBound(X)
                        buffer = buffer & WorksheetFunction.roman(j + 1) & "." & Trim(X(j)) & " "
                    Next j
                    ReDim X(0)
                    lData.MARC700 = buffer
                Else
                    lData.MARC700 = Trim(X(0))
                End If
        End Select
    Next i
    On Error GoTo 0
    
    ' In some cases, year are not specified correctly
    If Len(strAnee) >= 25 Then
        Dim firstDate As String, secondDate As String
        buffer = Right(strAnee, Len(strAnee) - 6)
        firstDate = Trim(Left(buffer, 4))
        buffer = Right(buffer, Len(buffer) - 15)
        secondDate = Trim(Left(buffer, 4))
        If Len(firstDate) Then
            buffer = firstDate
            If Len(secondDate) Then
                buffer = buffer & ", c" & secondDate
            End If
            lData.MARC008 = buffer & "."
        End If
    End If
    
    ' From strAnee, retrieve only letter tags to identify if book has images, index, or something else
    Dim lOut() As String
    ReDim lOut(0)
    If Len(strAnee) >= 25 Then
        buffer = Right(Left(strAnee, 43), 4)
        For i = 1 To 4
            Dim letter As String
            letter = Right(Left(buffer, i), 1)
            ReDim Preserve lOut(UBound(lOut) + 1)
            lOut(i) = letter
        Next i
    End If
    lData.CONTENTTAGS = lOut
    
    If Len(lData.MARC020) = 0 Then
        lData.MARC020 = strISBN
    End If
    lData.ID = idData
    lData.Valid = True
    
    lData.HTMLContent = returnAsHTML(lData)
    
    FindData = lData
End Function


