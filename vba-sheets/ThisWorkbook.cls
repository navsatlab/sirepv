VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ###################################################################
' # NAVSATLAB                                                       #
' # Application deployed over VBA                                   #
' # https://github.com/navsatlab/sirepv                             #
' # Some rights reserved                                            #
' ###################################################################

Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error GoTo Fail
    EncryptData
    ThisWorkbook.Save
    Exit Sub
    
Fail:
    MsgBox Err.description
End Sub

Private Sub Workbook_Open()
    On Error GoTo Fail
    DecryptData
    Exit Sub
    
Fail:
    MsgBox Err.description
End Sub
