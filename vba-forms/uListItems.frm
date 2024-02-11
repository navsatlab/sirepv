VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uListItems 
   Caption         =   "Listar elementos repetidos"
   ClientHeight    =   9090.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7110
   OleObjectBlob   =   "uListItems.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "uListItems"
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

Private cName As String

Private Sub InitArray(Optional columnName As String = "Editorial")
    Dim tbl As ListObject
    Dim colData As ListColumn
    Dim dict As Object
    Dim key As Variant
    Dim i As Long
    cName = columnName

    Set tbl = ThisWorkbook.Sheets(tSheet).ListObjects(tTable)
    Set colData = tbl.ListColumns(cName)
    Set dict = CreateObject("Scripting.Dictionary")
    
    For i = 1 To colData.DataBodyRange.Rows.Count
        Dim Value As Variant
        Value = colData.DataBodyRange(i, 1).Value
        If Not dict.Exists(Value) Then
            dict.Add Value, 1
        Else
            dict(Value) = dict(Value) + 1
        End If
    Next i
    
    tItems.ListItems.Clear
    tItems.View = lvwReport
    tItems.MultiSelect = True
    tItems.FullRowSelect = True
    tItems.Gridlines = True
    tItems.LabelEdit = lvwManual
    tItems.ColumnHeaders.Clear
    tItems.ColumnHeaders.Add , , "Nombre del Valor", 200
    tItems.ColumnHeaders.Add , , "Cantidad", 100
    
    For Each key In dict.Keys
        Dim Item As ListItem
        Set Item = tItems.ListItems.Add(, , key)
        Item.SubItems(1) = dict(key)
    Next key
    
    tItems.Sorted = True
End Sub

Private Sub SearchReplace()
    Dim tbl As ListObject
    Dim colData As ListColumn
    Dim searchData() As String
    Dim replacementText As String
    Dim cell As Range
    Dim i As Long
        
    searchData = Split(tValues.text, "|")
    replacementText = tNew.text
    Set tbl = ThisWorkbook.Sheets(tSheet).ListObjects(tTable)
    Set colData = tbl.ListColumns(cName)
    
    For Each cell In colData.DataBodyRange
        For i = LBound(searchData) To UBound(searchData)
            If StrComp(cell.Value, searchData(i), vbTextCompare) = 0 Then
                cell.Value = replacementText
                Exit For
            End If
        Next i
    Next cell
End Sub

Private Sub cModify_Click()
    Application.ScreenUpdating = False
    SearchReplace
    
    Dim listView As listView
    Dim selectedItem As ListItem
    Dim i As Long
    
    Set listView = tItems
    For i = 1 To listView.ListItems.Count
        If listView.ListItems(i).Selected Then
            listView.ListItems(i).Checked = True
            listView.ListItems(i).Bold = True
        End If
    Next i
   
    tNew.text = ""
    tNew.SetFocus
    Application.ScreenUpdating = True
End Sub

Private Function GetSelectedItems(ByVal listView As listView) As Collection
    Dim selectedItems As New Collection
    Dim i As Long
    
    For i = 1 To listView.ListItems.Count
        If listView.ListItems(i).Selected Then
            selectedItems.Add listView.ListItems(i)
        End If
    Next i
    
    Set GetSelectedItems = selectedItems
End Function

Private Sub cAnalize_Click()
    Application.ScreenUpdating = False
    InitArray cHead.text
    cAnalize.Default = False
    cModify.Default = True
    Me.Caption = "Listar elementos repetidos | Encontrados: " & tItems.ListItems.Count
    Application.ScreenUpdating = True
End Sub

Private Sub cExit_Click()
    Unload Me
End Sub

Private Sub tItems_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim selectedItem As ListItem
    Dim selectedText As String
    
    Dim selectedItems As Collection
    Set selectedItems = GetSelectedItems(tItems)
    
    If selectedItems.Count = 1 Then
        Set selectedItem = selectedItems(1)
        selectedText = selectedItem.text
        tValues.Value = selectedText
    ElseIf selectedItems.Count > 1 Then
        Dim i As Long
        For i = 1 To selectedItems.Count
            selectedText = selectedText & selectedItems(i).text & "|"
        Next i
        tValues.Value = Left(selectedText, Len(selectedText) - 1)
    Else
        tValues.Value = ""
    End If
End Sub

Private Sub tNew_Change()
    If Len(tValues.text) > 0 And Len(cName) > 0 Then
        If Len(tNew.text) > 0 Then
            cModify.Enabled = True
        Else
            cModify.Enabled = False
        End If
    Else
        cModify.Enabled = False
    End If
End Sub

Private Sub UserForm_Initialize()
    StartUpPosition = 0
    Left = Application.Left + (Application.width - width - 10)
    Top = Application.Top + (0.5 * Application.Height) - (0.5 * Height)
    
    Dim tbl As ListObject
    Dim col As ListColumn
    Dim columnName As String
    
    On Error Resume Next
    Set tbl = ThisWorkbook.Sheets(tSheet).ListObjects(tTable)
    On Error GoTo 0
    
    If tbl Is Nothing Then
        MsgBox "La tabla especificada no fue encontrada.", vbExclamation
        Exit Sub
    End If
    
    For Each col In tbl.ListColumns
        cHead.AddItem col.Name
    Next col
    cHead.ListIndex = 0
    
    
    lblVersion.Caption = SysVersion
    
    ThisWorkbook.Save
    Application.Calculation = xlCalculationManual
End Sub

Private Sub UserForm_Terminate()
    ThisWorkbook.Save
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
