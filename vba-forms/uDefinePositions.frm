VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uDefinePositions 
   Caption         =   "Definición de charolas y columnas para la sala"
   ClientHeight    =   5670
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8385.001
   OleObjectBlob   =   "uDefinePositions.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "uDefinePositions"
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

Private SetModify As Boolean

Private Sub cAddItem_Click()
    pVista.ListItems.Add
    pVista.ListItems(pVista.ListItems.Count).Selected = True
    
    fList.Visible = False
    fModify.Visible = True
    
    cEditItem.Default = False
    cCancel.Cancel = False
    cSave.Default = True
    cEscape.Cancel = True
    
    cAddItem.Enabled = False
    cDeleteItem.Enabled = False
    cEditItem.Enabled = False
    cMoveUp.Enabled = False
    cMoveDown.Enabled = False
    cCancel.Enabled = False
End Sub

Private Sub cCancel_Click()
    If SetModify Then
        If MsgBox("Se detectaron algunos cambios, ¿deseas guardarlos?", vbInformation + vbYesNo, "Columnas y charolas") = vbYes Then
            Application.ScreenUpdating = False
            Dim Data As ListObject, Item As ListItem, i As Integer
            Set Data = ThisWorkbook.Sheets("Settings").ListObjects("EXTERN_PREFIX")
            
            Dim ActData As Integer
            ActData = Data.Range.Rows.Count - 1
            If pVista.ListItems.Count > ActData Then
                For i = 1 To (pVista.ListItems.Count - ActData)
                    Data.ListRows.Add
                Next i
            ElseIf pVista.ListItems.Count < ActData Then
                For i = pVista.ListItems.Count To ActData
                    Data.ListRows(i - 2).Delete
                Next i
            End If
            
            For i = 1 To pVista.ListItems.Count
                Data.Range.Cells(i + 1, 1) = pVista.ListItems(i).text
                Data.Range.Cells(i + 1, 2) = pVista.ListItems(i).SubItems(1)
                Data.Range.Cells(i + 1, 3) = pVista.ListItems(i).SubItems(2)
                Data.Range.Cells(i + 1, 4) = pVista.ListItems(i).SubItems(3)
            Next i
            
            Application.ScreenUpdating = True
        End If
    End If
    Unload Me
End Sub

Private Sub cDeleteItem_Click()
    If pVista.selectedItem.Selected = True Then
        pVista.ListItems.Remove (pVista.selectedItem.Index)
        SetModify = True
    End If
End Sub

Private Sub cEditItem_Click()
    If pVista.selectedItem.Selected = True Then
        tColumna.text = pVista.ListItems(pVista.selectedItem.Index).text
        tCharola.text = pVista.ListItems(pVista.selectedItem.Index).SubItems(1)
        tFolio1.text = pVista.ListItems(pVista.selectedItem.Index).SubItems(2)
        tFolio2.text = pVista.ListItems(pVista.selectedItem.Index).SubItems(3)
        
        fList.Visible = False
        fModify.Visible = True
        
        cEditItem.Default = False
        cCancel.Cancel = False
        cSave.Default = True
        cEscape.Cancel = True
        
        cAddItem.Enabled = False
        cDeleteItem.Enabled = False
        cEditItem.Enabled = False
        cMoveUp.Enabled = False
        cMoveDown.Enabled = False
        cCancel.Enabled = False
    End If
End Sub

Private Sub cEscape_Click()
    tColumna.text = ""
    tCharola.text = ""
    tFolio1.text = ""
    tFolio2.text = ""
    
    fList.Visible = True
    fModify.Visible = False
    
    cSave.Default = False
    cEscape.Cancel = False
    cEditItem.Default = True
    cCancel.Cancel = True
    
    cAddItem.Enabled = True
    cDeleteItem.Enabled = True
    cEditItem.Enabled = True
    cMoveUp.Enabled = True
    cMoveDown.Enabled = True
    cCancel.Enabled = True
End Sub

Private Sub cMoveDown_Click()
    'Pending
End Sub

Private Sub cMoveUp_Click()
    'Pending
End Sub

Private Sub cSave_Click()
    fList.Visible = True
    pVista.ListItems(pVista.selectedItem.Index).text = tColumna.text
    pVista.ListItems(pVista.selectedItem.Index).SubItems(1) = tCharola.text
    pVista.ListItems(pVista.selectedItem.Index).SubItems(2) = tFolio1.text
    pVista.ListItems(pVista.selectedItem.Index).SubItems(3) = tFolio2.text
    fModify.Visible = False
    
    cSave.Default = False
    cEscape.Cancel = False
    cEditItem.Default = True
    cCancel.Cancel = True
    
    cAddItem.Enabled = True
    cDeleteItem.Enabled = True
    cEditItem.Enabled = True
    cMoveUp.Enabled = True
    cMoveDown.Enabled = True
    cCancel.Enabled = True
    
    tColumna.text = ""
    tCharola.text = ""
    tFolio1.text = ""
    tFolio2.text = ""
    
    SetModify = True
End Sub

Private Sub UserForm_Initialize()
    StartUpPosition = 0
    Left = Application.Left + (Application.width - width - 10)
    Top = Application.Top + (0.5 * Application.Height) - (0.5 * Height)
    
    pVista.View = lvwReport
    pVista.Gridlines = True
    pVista.LabelEdit = lvwManual
    pVista.FullRowSelect = True
    pVista.ColumnHeaders.Add , , "Columna", 50
    pVista.ColumnHeaders.Add , , "Charola", 50
    pVista.ColumnHeaders.Add , , "Primer folio", 65
    pVista.ColumnHeaders.Add , , "Último folio", 65
    
    Dim Data As ListObject, Item As ListItem, i As Integer
    Set Data = ThisWorkbook.Sheets("Settings").ListObjects("EXTERN_PREFIX")
    
    For i = 2 To Data.Range.Rows.Count
        Set Item = pVista.ListItems.Add(, , Data.Range.Cells(i, 1))
        Item.SubItems(1) = Data.Range.Cells(i, 2)
        Item.SubItems(2) = Data.Range.Cells(i, 3)
        Item.SubItems(3) = Data.Range.Cells(i, 4)
    Next i
    
    lblVersion.Caption = SysVersion
    
    ThisWorkbook.Save
    Application.Calculation = xlCalculationManual
End Sub

Private Sub UserForm_Terminate()
    ThisWorkbook.Save
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
