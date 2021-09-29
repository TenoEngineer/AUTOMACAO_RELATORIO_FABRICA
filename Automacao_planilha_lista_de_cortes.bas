Attribute VB_Name = "Módulo1"

Sub clean()
Attribute clean.VB_ProcData.VB_Invoke_Func = "X\n14"
'
' clean file
'
Dim importFileName, pathFile As Variant
Dim importWorkbook As Workbook
Dim importSheet As Worksheet

' Variáveis do tecnometal
Dim importPosition, importQuant, importDim As Range
Dim importComp, importName, importKg, importArea As Range

Dim wb As Workbook
Dim perfil, chapa As Worksheet
Dim find, replace As Variant

Dim rng As Range
Dim r As Long
Dim i, j, h As Variant
Dim cell, cel, select_row As Variant

Dim procv As Variant
Dim peca As Variant

'Dim fso As New FileSystemObject
Dim fileName As String

Set wb = ThisWorkbook
Set perfil = wb.Worksheets(1)
Set chapa = wb.Worksheets(2)

' Limpeza do arquivo
wb.Sheets("PRANCHA").Activate
wb.Sheets("PRANCHA").Range("A2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents

chapa.Activate
chapa.Range("A13:S4000").Select
Selection.Borders.LineStyle = xlNone
Selection.ClearContents

perfil.Activate
perfil.Range("A13:S4000").Select
Selection.Borders.LineStyle = xlNone
Selection.ClearContents

perfil.Cells.Interior.ColorIndex = 0
chapa.Cells.Interior.ColorIndex = 0
  

'Determina o arquivo onde sera colado os dados
 
    ' perfilow open file dialog
    importFileName = Application.GetOpenFilename(FileFilter:="Arquivo do Excel (*.xls; *.xlsx; *.R35), *.xls;*.xlsx; *.R35", Title:="Escolha um arquivo do Excel")
    
    
    ' if user pressed cancel buton: exit
    If importFileName = False Then Exit Sub

        parentName = CreateObject("scripting.filesystemobject").GetParentFolderName(importFileName)
        
        If importFileName Like "*.R35" Then
            With CreateObject("wscript.shell")
               .currentdirectory = parentName
               .Run "%comspec% /c ren *.R35 *.xls", 0, True
            End With
            pathFile = Split(importFileName, ".")
            pathFile = pathFile(0) & "." & pathFile(1) & ".xls"
        Else
            pathFile = importFileName
        End If
    
    
    Application.ScreenUpdating = False
    
         ' if user selected a excel file, open it
         Set importWorkbook = Application.Workbooks.Open(pathFile)
         Set importperfileet = importWorkbook.Worksheets(1)
         
        importperfileet.Cells.Select
        importperfileet.Sort.SortFields.Clear
        importperfileet.Sort.SortFields.Add Key:=Range( _
            "E2", Range("E2").End(xlDown)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
        With ActiveWorkbook.Worksheets(1).Sort
            .SetRange Range("A1:U2000")
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With

         ' copy from import perfileet
         Set importPosition = importperfileet.Range("D2", importperfileet.Range("D2").End(xlDown).Offset(-1, 0))
         Set importQuant = importperfileet.Range("I2", importperfileet.Range("I2").End(xlDown))
         Set importDim = importperfileet.Range("E2", importperfileet.Range("E2").End(xlDown))
         Set importComp = importperfileet.Range("K2", importperfileet.Range("K2").End(xlDown))
         Set importName = importperfileet.Range("B2", importperfileet.Range("B2").End(xlDown))
         Set importKg = importperfileet.Range("S2", importperfileet.Range("S2").End(xlDown).Offset(-1, 0))
         Set importArea = importperfileet.Range("U2").End(xlDown)
         
         'MENSAGEM DE ERRO CASO ABRA PLANILHA ERRADA
         If importperfileet.Range("D1").Value <> "POS_PEZ" Then
            MsgBox "Planilha não exportada por POSIÇÕES PARA MARCA no Tecnometal", , "Error"
            Exit Sub
         End If
         If importperfileet.Range("I1").Value <> "QTA_TOT" Then
            MsgBox "Planilha não exportada por POSIÇÕES PARA MARCA no Tecnometal", , "Error"
            Exit Sub
         End If
         If importperfileet.Range("E1").Value <> "NOM_PRO" Then
            MsgBox "Planilha não exportada por POSIÇÕES PARA MARCA no Tecnometal", , "Error"
            Exit Sub
         End If
         If importperfileet.Range("K1").Value <> "LUN_PRO" Then
            MsgBox "Planilha não exportada por POSIÇÕES PARA MARCA no Tecnometal", , "Error"
            Exit Sub
         End If
         If importperfileet.Range("B1").Value <> "MAR_PEZ" Then
            MsgBox "Planilha não exportada por POSIÇÕES PARA MARCA no Tecnometal", , "Error"
            Exit Sub
         End If
         If importperfileet.Range("S1").Value <> "PTO_LIS" Then
            MsgBox "Planilha não exportada por POSIÇÕES PARA MARCA no Tecnometal", , "Error"
            Exit Sub
         End If
         If importperfileet.Range("U1").Value <> "STO_LIS" Then
            MsgBox "Planilha não exportada por POSIÇÕES PARA MARCA no Tecnometal", , "Error"
            Exit Sub
         End If
                
         ' paste into Data perfileet
         importPosition.Copy
         wb.Activate
         perfil.Range("B13").PasteSpecial xlValues
         
         importQuant.Copy
         wb.Activate
         perfil.Range("C13").PasteSpecial xlValues
         
         importDim.Copy
         wb.Activate
         perfil.Range("E13").PasteSpecial xlValues
         
         importComp.Copy
         wb.Activate
         perfil.Range("F13").PasteSpecial xlValues
         
         importName.Copy
         wb.Activate
         perfil.Range("J13").PasteSpecial xlValues
                  
         importKg.Copy
         wb.Activate
         perfil.Range("K13").PasteSpecial xlValues
         
         importArea.Copy
         wb.Activate
         perfil.Range("L7").PasteSpecial xlValues
        
         Application.CutCopyMode = False
 
'Fecha o excel do tecnometal
importWorkbook.Close

'Substitui os caracteres bugados
find = "Ï"
replace = "Ø"
 
perfil.Cells.replace what:=find, Replacement:=replace, _
LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
SearchFormat:=False, ReplaceFormat:=False

'Copia todas as peças da prancha para fazer o UNIQUE
perfil.Range("J13", Range("J13").End(xlDown)).Copy
wb.Sheets("PRANCHA").Cells(2, 1).PasteSpecial xlValues

'Conta a quantidade de células utilizadas
r = perfil.Range("B13", Range("B13").End(xlDown)).Rows.Count
cell = 13
h = 1

'Ajusta as vírgulas nos comprimentos e pesos
For i = 0 To r
    perfil.Cells(cell, 11).NumberFormat = "0.0"
    perfil.Cells(cell, 6).Value = Round(perfil.Cells(cell, 6).Value, 0)
    cell = cell + 1
Next

cell = 13
h = 1
'Insere todas as chapas na aba da CHAPA
For i = 0 To r
    num = cell
    If perfil.Range("E" & num).Value Like "*CH*" Then
        select_row = 12 + h
        perfil.Cells(num, 1).EntireRow.Copy
        chapa.Cells(select_row, 1).PasteSpecial xlValues
        h = h + 1
    End If
    cell = cell + 1
Next

'Deleta todas as chapas na aba PERFIL
num = 13
h = 0
For i = 0 To r
    num = num + h
    If perfil.Range("E" & num).Value Like "*CH*" Then
        perfil.Cells(num, 1).EntireRow.Delete
        h = 0
    Else
        h = 1
    End If
Next


Worksheets(1).Activate

r = perfil.Range("B13", Range("B13").End(xlDown)).Rows.Count

'Gera a simbologia da peça
'peca = Cells(
'resultado_procv = Application.VLookup(

'Gera o comprimento total da peça
cel = 13
For i = 0 To r
    perfil.Cells(cel, 7).Value = perfil.Cells(cel, 6).Value * perfil.Cells(cel, 3).Value
    cel = cel + 1
Next
perfil.Range("G13").End(xlDown).ClearContents

'Numera os itens da aba PERFIL
cel = 13
j = 1
For i = 0 To r
    perfil.Cells(cel, 1).Value = j
    cel = cel + 1
    j = j + 1
Next
perfil.Range("A13").End(xlDown).ClearContents

'Numera os itens da Aba CHAPA
Worksheets(2).Activate
r2 = chapa.Range("B13", Range("B13").End(xlDown)).Rows.Count
cel = 13
j = 1
For i = 0 To r2
    chapa.Cells(cel, 1).Value = j
    cel = cel + 1
    j = j + 1
Next
chapa.Range("A13").End(xlDown).ClearContents
 
wb.Sheets("RESUMO_PERFIS").PivotTables("Tabela dinâmica16").PivotCache.Refresh
wb.Sheets("RESUMO_CHAPAS").PivotTables("Tabela dinâmica16").PivotCache.Refresh
wb.Sheets("PRANCHA").PivotTables("Tabela dinâmica1").PivotCache.Refresh

'Marca as peças por família de bitola
chapa.Activate
r = chapa.Range("B13", Range("B13").End(xlDown)).Rows.Count
r = r - 2
cel = 14
For i = 0 To r
    If chapa.Cells(cel, 5).Value <> chapa.Cells(cel - 1, 5).Value Then
        If chapa.Cells(cel - 1, 5).Interior.Color <> RGB(170, 170, 170) Then
            chapa.Cells(cel, 1).EntireRow.Interior.Color = RGB(170, 170, 170)
        End If
    Else
        If chapa.Cells(cel - 1, 5).Interior.Color = RGB(170, 170, 170) Then
            chapa.Cells(cel, 1).EntireRow.Interior.Color = RGB(170, 170, 170)
        End If
    End If
cel = cel + 1
Next

'Marca as peças por família de bitola
perfil.Activate
r = perfil.Range("B13", Range("B13").End(xlDown)).Rows.Count
r = r - 2
cel = 14
For i = 0 To r
    If perfil.Cells(cel, 5).Value <> perfil.Cells(cel - 1, 5).Value Then
        If perfil.Cells(cel - 1, 5).Interior.Color <> RGB(170, 170, 170) Then
            perfil.Cells(cel, 1).EntireRow.Interior.Color = RGB(170, 170, 170)
        End If
    Else
        If perfil.Cells(cel - 1, 5).Interior.Color = RGB(170, 170, 170) Then
            perfil.Cells(cel, 1).EntireRow.Interior.Color = RGB(170, 170, 170)
        End If
    End If
cel = cel + 1
Next

'Ajusta as linhas
perfil.Activate
perfil.Range("A13:S13", Range("A13:S13").End(xlDown)).Select
Selection.Font.Name = "Arial"
Selection.Font.Size = 14
Selection.RowHeight = 23.25
Selection.Borders.LineStyle = xlContinuous
Selection.HorizontalAlignment = xlCenter
perfil.Range("F13").End(xlDown).ClearContents

chapa.Activate
chapa.Range("A13:S13", Range("A13:S13").End(xlDown)).Select
Selection.Font.Name = "Arial"
Selection.Font.Size = 14
Selection.RowHeight = 23.25
Selection.Borders.LineStyle = xlContinuous
Selection.HorizontalAlignment = xlCenter

perfil.Activate

End Sub

