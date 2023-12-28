VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "sMain"
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

Private Sub cAyuda_Click()
    uAbout.Show
End Sub

Private Sub cBuscar_Click()
    LoadSearch
End Sub

Private Sub cCambiar_Click()
    LoadSelect
End Sub

Private Sub cCharolas_Click()
    LoadDefinePositions
End Sub

Private Sub cCotejo_Click()
    LoadCompare
End Sub

Private Sub cImprimir_Click()
    LoadPrint
End Sub

Private Sub cRegistrar_Click()
    LoadRegBooks
End Sub

Private Sub cReporteMensual_Click()
    LoadReportBookView
End Sub

Private Sub Worksheet_Activate()
    cCotejo.Caption = "Cotejo / Inventario" & vbNewLine & "Ctrl + Shift + C"
    cImprimir.Caption = "Imprimir fichas" & vbNewLine & "Ctrl + Shift + P"
    cCambiar.Caption = "Cambiar tabla" & vbNewLine & "Ctrl + Shift + S"
    cBuscar.Caption = "Consulta" & vbNewLine & "Ctrl + Shift + F"
End Sub

