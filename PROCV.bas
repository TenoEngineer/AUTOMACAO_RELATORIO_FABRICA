Attribute VB_Name = "Módulo3"
Sub procv()

Dim wb As Workbook
Dim perfil, chapa As Worksheet

Dim r As Long
Dim i, procv, prancha As Variant
Dim cel As Variant

Set wb = ThisWorkbook
Set perfil = wb.Worksheets(1)
Set chapa = wb.Worksheets(2)

If perfil.Range("B13").Value = Empity And chapa.Range("B13").Value = Empity Then
    MsgBox "Primeiro deve inserir os dados na tabela", , "ERROR"
    Exit Sub
End If


perfil.Activate
If perfil.Range("B13").Value <> Empity Then
    r = perfil.Range("B13", Range("B13").End(xlDown)).Rows.Count
    cel = 13
    For i = 0 To r
        prancha = perfil.Cells(cel, 10).Value
        procv = Application.VLookup(prancha, Sheets("PRANCHA").Range("B:C"), 2, False)
        perfil.Cells(cel, 9).Value = procv
        cel = cel + 1
    Next
    perfil.Range("I13").End(xlDown).ClearContents
End If

chapa.Activate
If chapa.Range("B13").Value <> Empity Then
    r = chapa.Range("B13", Range("B13").End(xlDown)).Rows.Count
    cel = 13
    For i = 0 To r
        prancha = chapa.Cells(cel, 10).Value
        procv = Application.VLookup(prancha, Sheets("PRANCHA").Range("B:C"), 2, False)
        chapa.Cells(cel, 9).Value = procv
        cel = cel + 1
    Next
    chapa.Range("I13").End(xlDown).ClearContents
End If

perfil.Active


End Sub
