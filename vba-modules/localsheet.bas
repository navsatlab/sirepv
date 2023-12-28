Attribute VB_Name = "localsheet"
' ###################################################################
' # NAVSATLAB                                                       #
' # Application deployed over VBA                                   #
' # https://github.com/navsatlab/sirepv                             #
' # Some rights reserved                                            #
' ###################################################################

Option Explicit
' Funciones específicas de búsqueda en inventario de Excel

' Get value parametter saved in SETTINGS table
Public Function GetParam(ByVal ParamName As String) As String
    Dim Data As ListObject, buff As ListRow
    Set Data = ThisWorkbook.Sheets("Settings").ListObjects("SETTINGS")
    For Each buff In Data.ListRows
        If Data.DataBodyRange.Cells(buff.Index, 1).Value = ParamName Then
            GetParam = Data.DataBodyRange.Cells(buff.Index, 2).Value
            Exit Function
        End If
    Next buff
End Function

' Set value parametter saved in SETTINGS table
Public Sub SetParam(ByVal ParamName As String, ByVal ParamValue As String)
    Dim Data As ListObject, buff As ListRow
    Set Data = ThisWorkbook.Sheets("Settings").ListObjects("SETTINGS")
    For Each buff In Data.ListRows
        If Data.DataBodyRange.Cells(buff.Index, 1).Value = ParamName Then
            Data.DataBodyRange.Cells(buff.Index, 2).Value = ParamValue
            Exit Sub
        End If
    Next buff
End Sub

' Get ID position of some comumn
Public Function GetPos(ByVal columnName As String) As Variant
    Dim content As ListObject
    Dim col As ListColumn
    
    Set content = ThisWorkbook.Sheets(tSheet).ListObjects(tTable)
    On Error Resume Next
    Set col = content.ListColumns(columnName)
    On Error GoTo 0
    If Not col Is Nothing Then
        GetPos = col.Index
    End If
End Function

' Find Excel row integer giving book-number
Public Function FindExcelData(ByVal ID As String, ByVal colFolio) As Integer
    Dim buff As Range
    Dim content As ListObject

    Set content = ThisWorkbook.Sheets(tSheet).ListObjects(tTable)
    ID = ParseNumber(ID)
    For Each buff In content.ListColumns(colFolio).DataBodyRange
        If StrComp(buff.Value, ID, vbTextCompare) = 0 Then
            FindExcelData = buff.row
            Exit For
        End If
    Next buff
End Function

' Elimina todo cero extra en el folio ingresado
Public Function ParseNumber(ByVal ID As String) As String
    Dim Data() As String, i As Integer
    Data = Split(ID, "-")
    If UBound(Data) = 0 Then
        ParseNumber = ID
        Exit Function
    End If
    For i = LBound(Data) To UBound(Data)
        If Not Len(Data(i)) = 2 And Not i = 1 Then _
            Data(i) = RegExpReplace(Data(i), "^0+", "")
        If i = 1 And Len(Data(i)) > 2 Then _
            Data(i) = RegExpReplace(Data(i), "^0+", "")
    Next i
    ParseNumber = Data(0) & "-" & Data(1)
End Function

' Ordena por LC los contenidos, especificando ítem actual y el anterior, para hacer comparación
' Devolviendo true si es correcto y false si no lo es
Public Function ArrangeLC(ByVal Item As String, ByVal LastItem As String) As Boolean
    Dim ParsedLC As String, ParsedLastLC As String
    ParsedLC = Replace(Replace(Item, "T-", ""), "C-", "")
    ParsedLastLC = Replace(Replace(LastItem, "T-", ""), "C-", "")
    
    ' Get items at least to find ' ' character (to avoid volume, and item count)
    On Error Resume Next
    ParsedLC = Trim(UCase(Left(ParsedLC, InStrRev(ParsedLC, " ") - 1)))
    ParsedLastLC = Trim(UCase(Left(ParsedLastLC, InStrRev(ParsedLastLC, " ") - 1)))
    On Error GoTo 0
    
    If ParsedLC = ParsedLastLC Then
        ArrangeLC = False
        Exit Function
    End If
    
    ' Split items
    Dim Split1() As String, Split2() As String, i As Integer
    Split1 = Split(ParsedLC, ".")
    Split2 = Split(ParsedLastLC, ".")
    
    ' View all items of first param and compare it and arrange
    For i = LBound(Split1) To UBound(Split1)
        Dim LCNumber1 As String, LCNumber2 As String, LCLetter1 As String, LCLetter2 As String
        
        ' If i < than other item, it means that next item is more greater than this
        If Not i > UBound(Split2) Then
            LCLetter1 = RegExpReplace(Split1(i), "[\d]", "")
            LCNumber1 = RegExpReplace(Split1(i), "[^\d]", "")
            
            LCLetter2 = RegExpReplace(Split2(i), "[\d]", "")
            LCNumber2 = RegExpReplace(Split2(i), "[^\d]", "")
            
            
        End If
    Next i
End Function


