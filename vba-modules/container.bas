Attribute VB_Name = "container"
' ###################################################################
' # NAVSATLAB                                                       #
' # Application deployed over VBA                                   #
' # https://github.com/navsatlab/sirepv                             #
' # Some rights reserved                                            #
' ###################################################################

Option Explicit
' Módulo de habilitación y activación de funciones

Public Sub EncryptData()
    ' Encrypt before close
    If GetParam("0x25") = "load" Then Exit Sub
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    Dim Data As ListObject, i As Integer, key As String, str As String
    key = LoadKey
    
    Set Data = ThisWorkbook.Sheets("Settings").ListObjects("EXTERN_PREFIX")
    For i = 2 To Data.Range.Rows.Count
        str = Data.Range.Cells(i, 3)
        If Len(str) > 0 Then _
            Data.Range.Cells(i, 3) = StoreEncryptAES(str, key, 10)
            
        str = Data.Range.Cells(i, 4)
        If Len(str) > 0 Then _
            Data.Range.Cells(i, 4) = StoreEncryptAES(str, key, 10)
    Next i
    
    Set Data = ThisWorkbook.Sheets("Settings").ListObjects("SUFFIX")
    For i = 2 To Data.Range.Rows.Count
        str = Data.Range.Cells(i, 1)
        If Len(str) > 0 Then _
            Data.Range.Cells(i, 1) = StoreEncryptAES(str, key, 10)
            
        str = Data.Range.Cells(i, 2)
        If Len(str) > 0 Then _
            Data.Range.Cells(i, 2) = StoreEncryptAES(str, key, 10)
    Next i
    
    Set Data = ThisWorkbook.Sheets("Settings").ListObjects("PREFIX")
    For i = 2 To Data.Range.Rows.Count
        str = Data.Range.Cells(i, 1)
        If Len(str) > 0 Then _
            Data.Range.Cells(i, 1) = StoreEncryptAES(str, key, 10)
            
        str = Data.Range.Cells(i, 2)
        If Len(str) > 0 Then _
            Data.Range.Cells(i, 2) = StoreEncryptAES(str, key, 10)
    Next i
    
    SetParam "0x25", "load"
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Public Sub DecryptData()
    ' Decrypt after open
    If GetParam("0x25") = "unload" Then Exit Sub
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    Dim Data As ListObject, i As Integer, key As String, str As String
    key = LoadKey
    
    Set Data = ThisWorkbook.Sheets("Settings").ListObjects("EXTERN_PREFIX")
    For i = 2 To Data.Range.Rows.Count
        str = Data.Range.Cells(i, 3)
        If Len(str) > 0 Then _
            Data.Range.Cells(i, 3) = RetrieveDecryptAES(str, key, 10)
            
        str = Data.Range.Cells(i, 4)
        If Len(str) > 0 Then _
            Data.Range.Cells(i, 4) = RetrieveDecryptAES(str, key, 10)
    Next i
    
    Set Data = ThisWorkbook.Sheets("Settings").ListObjects("SUFFIX")
    For i = 2 To Data.Range.Rows.Count
        str = Data.Range.Cells(i, 1)
        If Len(str) > 0 Then _
            Data.Range.Cells(i, 1) = RetrieveDecryptAES(str, key, 10)
            
        str = Data.Range.Cells(i, 2)
        If Len(str) > 0 Then _
            Data.Range.Cells(i, 2) = RetrieveDecryptAES(str, key, 10)
    Next i
    
    Set Data = ThisWorkbook.Sheets("Settings").ListObjects("PREFIX")
    For i = 2 To Data.Range.Rows.Count
        str = Data.Range.Cells(i, 1)
        If Len(str) > 0 Then _
            Data.Range.Cells(i, 1) = RetrieveDecryptAES(str, key, 10)
            
        str = Data.Range.Cells(i, 2)
        If Len(str) > 0 Then _
            Data.Range.Cells(i, 2) = RetrieveDecryptAES(str, key, 10)
    Next i
    
    SetParam "0x25", "unload"
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Public Function LoadKey() As String
    ' Read saved key in machine
    Dim FSO As New FileSystemObject
    Dim JsonTS As TextStream, lKey As String
    
    Set JsonTS = FSO.OpenTextFile(Environ("AppData") & "\key.json", ForReading)
    lKey = JsonTS.ReadAll
    JsonTS.Close
    
    LoadKey = lKey
End Function

Public Sub GenerateKey()
    ' Generates some random key and store it
    Dim FSO As New FileSystemObject
    Dim JsonTS As TextStream, lKey As String
    
    Set JsonTS = FSO.CreateTextFile(Environ("AppData") & "\key.json", True)
    
    lKey = Random(12)
    JsonTS.Write lKey
    JsonTS.Close
End Sub

Public Function Random(RLength As Integer) As String
' This function creates a string of random characters, both numbers
' and alpha, with a length of RLength.  It uses Timer to seed the Rnd
' function.

' Random() Version 1.0.0
' Copyright © 2009 Extra Mile Data, www.extramiledata.com.
' For questions or issues, please contact support@extramiledata.com.
' Use (at your own risk) and modify freely as long as proper credit is given.

On Error GoTo Err_Random

    Dim strTemp As String
    Dim intLoop As Integer
    Dim strCharBase As String
    Dim intPos As Integer
    Dim intLen As Integer

    ' Build the base.
    strCharBase = "01234ABCDEFGHIJKLMNOPQRSTUVWXYZ" _
    & "abcdefghijklmnopqrstuvwxyz56789"
    ' Get it's length.
    intLen = Len(strCharBase)

    ' Initialize the results.
    strTemp = String(RLength, "A")

    ' Reset the random seed.
    Rnd -1
    ' Initialize the seed using Timer.
    Randomize (Timer)

    ' Loop until you hit the end of strTemp.  Replace each character
    ' with a character selected at random from strCharBase.
    For intLoop = 1 To Len(strTemp)
        ' Use the Rnd function to pick a position number in strCharBase.
        ' If the result exceeds the length of strCharBase, subtract one.
        intPos = CInt(Rnd() * intLen + 1)
        If intPos > intLen Then intPos = intPos - 1
        ' Now assign the character at that position in the base to the
        ' next strTemp position.
        Mid$(strTemp, intLoop, 1) = Mid$(strCharBase, intPos, 1)
    Next

    ' Return the results.
    Random = strTemp

Exit_Random:
    On Error Resume Next
    Exit Function

Err_Random:
    MsgBox Err.number & " " & Err.description, vbCritical, "Random"
    Random = ""
    Resume Exit_Random

End Function

