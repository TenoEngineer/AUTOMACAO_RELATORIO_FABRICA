Attribute VB_Name = "Módulo1"

Sub clean()
'
'
' ******************* AUTOMAÇÃO DO PREÇO DE LISTA DE CORTE    *********************
'#AUTHOR: HEITOR TENO MÜLLER
'#DATE: 01/10/2021
'#CONTACT: heitortmuller@gmail.com

' ----------------  DENIÇÃO DAS VARIÁVIES    ----------------------
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

Dim fileName As String

' ----------------  DEFININDO WORKSHEETS    ----------------------
Set wb = ThisWorkbook
Set perfil = wb.Worksheets(1)
Set chapa = wb.Worksheets(2)

' ----------------  LIMPEZA DA PLANILHA CASO TENHA SIDO UTILIZADA    ----------------------
Dim answer As Integer
 
answer = MsgBox("Deseja limpar os dados da planilha?", vbQuestion + vbYesNo + vbDefaultButton2, "AVISO")
If answer = vbYes Then
    wb.Sheets("PRANCHA").Activate
    If wb.Sheets("PRANCHA").Range("A2") <> Empity Then
    wb.Sheets("PRANCHA").Range("A2", Range("A2").End(xlDown)).ClearContents
    End If
    
    chapa.Activate
    If chapa.Range("A13") <> Empity Then
    chapa.Range("A13:S13", Range("A13:S13").End(xlDown)).Select
    Selection.Borders.LineStyle = xlNone
    Selection.ClearContents
    chapa.Cells.Interior.ColorIndex = 0
    End If
    
    perfil.Activate
    If perfil.Range("A13") <> Empity Then
    perfil.Range("A13:S13", Range("A13:S13").End(xlDown)).Select
    Selection.Borders.LineStyle = xlNone
    Selection.ClearContents
    perfil.Cells.Interior.ColorIndex = 0
    perfil.Range("K5:N5").ClearContents
    perfil.Range("K7").ClearContents
    perfil.Range("L7").ClearContents
    End If
Else: Exit Sub
End If
  
' ----------------  SELEÇÃO DO ARQUIVO PARA IMPORTAÇÃO DOS DADOS    ----------------------
importFileName = Application.GetOpenFilename(FileFilter:="Arquivo do Excel (*.xls; *.xlsx; *.R35), *.xls;*.xlsx; *.R35", Title:="Escolha um arquivo do Excel")
    
    
' CASO A AÇÃO TENHA SIDO CANCELADO
If importFileName = False Then
    chapa.Activate
    chapa.Range("A13").Select
    perfil.Activate
    perfil.Range("A13").Select
    wb.Sheets("RESUMO_PERFIS").PivotTables("Tabela dinâmica16").PivotCache.Refresh
    wb.Sheets("RESUMO_CHAPAS").PivotTables("Tabela dinâmica16").PivotCache.Refresh
    wb.Sheets("PRANCHA").PivotTables("Tabela dinâmica1").PivotCache.Refresh
    Exit Sub
End If
    
' ------------------  TRANSFORMANDO EXTENSÃO DE R35 PARA XLS    ----------------------
'SELECIONA O CAMINHO DO ARQUIVO
parentName = CreateObject("scripting.filesystemobject").GetParentFolderName(importFileName)
pathFile = Split(importFileName, ".")

Dim MyValue As Integer
MyValue = Int((100 * Rnd))

If Not Dir(pathFile(0) & "." & pathFile(1) & ".xls", vbDirectory) = vbNullString And Dir(importFileName, vbDirectory) = vbNullString Then
    Dim change, file As String
    file = pathFile(0) & "." & pathFile(1) & ".xls"
    change = pathFile(0) & "." & pathFile(1) & "_" & MyValue & ".xls"
    Name file As change
End If

'FAZ A TRANSFORMAÇÃO DA EXTENSÃO
If importFileName Like "*.R35" Then
    With CreateObject("wscript.shell")
       .currentdirectory = parentName
       .Run "%comspec% /c ren *.R35 *.xls", 0, True
    End With
    pathFile = pathFile(0) & "." & pathFile(1) & ".xls"
Else
    pathFile = importFileName
End If


Application.ScreenUpdating = False

'APÓS SELECIONAR O ARQUIVO
Set importWorkbook = Application.Workbooks.Open(pathFile)
Set importperfileet = importWorkbook.Worksheets(1)

' ------------------  ORGANIZA OS DADOS EM GRUPOS POR ORDEM ALFABÉTICA   ------------------
importperfileet.Cells.Select
importperfileet.Sort.SortFields.Clear
importperfileet.Sort.SortFields.Add Key:=Range( _
   "E2", Range("E2").End(xlDown)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
   xlSortNormal
With ActiveWorkbook.Worksheets(1).Sort
   .SetRange Range("A1:U1", Range("A1:U1").End(xlDown))
   .Header = xlYes
   .MatchCase = False
   .Orientation = xlTopToBottom
   .SortMethod = xlPinYin
   .Apply
End With

' ------------------  SELECIONA OS DADOS QUE DEVEM SER IMPORTADOS   ------------------
Set importPosition = importperfileet.Range("D2", importperfileet.Range("D2").End(xlDown).Offset(-1))
Set importQuant = importperfileet.Range("I2", importperfileet.Range("I2").End(xlDown))
Set importDim = importperfileet.Range("E2", importperfileet.Range("E2").End(xlDown))
Set importComp = importperfileet.Range("K2", importperfileet.Range("K2").End(xlDown))
Set importName = importperfileet.Range("B2", importperfileet.Range("B2").End(xlDown))
Set importKg = importperfileet.Range("S2", importperfileet.Range("S2").End(xlDown).Offset(-1))
Set importArea = importperfileet.Range("U2").End(xlDown)

'MENSAGEM DE ERRO CASO ABRA PLANILHA ERRADA
If importperfileet.Range("D1").Value <> "POS_PEZ" Or importperfileet.Range("I1").Value <> "QTA_TOT" Or importperfileet.Range("E1").Value <> "NOM_PRO" Or importperfileet.Range("K1").Value <> "LUN_PRO" Or importperfileet.Range("B1").Value <> "MAR_PEZ" Or importperfileet.Range("S1").Value <> "PTO_LIS" Or importperfileet.Range("U1").Value <> "STO_LIS" Then
   MsgBox "Planilha não exportada por POSIÇÕES PARA MARCA no Tecnometal", , "Error"
   importWorkbook.Close Savechanges:=False
   Exit Sub
End If
       
' ------------------  COLA OS DADOS NA LISTA DE CORTES   ------------------
wb.Activate
importPosition.Copy
perfil.Range("B13").PasteSpecial xlValues
importQuant.Copy
perfil.Range("C13").PasteSpecial xlValues
importDim.Copy
perfil.Range("E13").PasteSpecial xlValues
importComp.Copy
perfil.Range("F13").PasteSpecial xlValues
importName.Copy
perfil.Range("J13").PasteSpecial xlValues
importKg.Copy
perfil.Range("K13").PasteSpecial xlValues
importArea.Copy
perfil.Range("L7").PasteSpecial xlValues
Application.CutCopyMode = False
 
'FECHA A PLANILHA SEM SALVAR
importWorkbook.Close Savechanges:=False

' ------------------  AJUSTA OS DADOS IMPORTADOS   ------------------
'SUBSTITUI OS CARACTERES BUGADOS
find = "Ï"
replace = "Ø"
 
perfil.Cells.replace what:=find, Replacement:=replace, _
LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
SearchFormat:=False, ReplaceFormat:=False

'COPIA TODAS AS PEÇAS DA PRANCHA PARA FAZER O UNIQUE
perfil.Range("J13", Range("J13").End(xlDown)).Copy
wb.Sheets("PRANCHA").Cells(2, 1).PasteSpecial xlValues

'CONTA A QUANTIDADE DE CÉLULAR UTILIZADAS
r = perfil.Range("B13", Range("B13").End(xlDown)).Rows.Count
cell = 13
h = 1

'AJUSTA AS VÍRGULAS NOS COMPRIMENTOS E PESOS
If perfil.Range("B13").Value <> Empity Then
    For i = 1 To r
        perfil.Cells(cell, 11).NumberFormat = "0.0"
        perfil.Cells(cell, 6).Value = Round(perfil.Cells(cell, 6).Value, 0)
        cell = cell + 1
    Next
End If
cell = 13
h = 1

' ------------------  INSERE AS LINHAS COM CHAPAS PARA A ABA CHAPA   ------------------
For i = 1 To r
    num = cell
    If perfil.Range("E" & num).Value Like "CH*" Then
        select_row = 12 + h
        perfil.Cells(num, 1).EntireRow.Copy
        chapa.Cells(select_row, 1).PasteSpecial xlValues
        h = h + 1
    End If
    cell = cell + 1
Next

'DELETA TODAS AS CHAPAS NA ABA PERFIL
num = 13
h = 0
For i = 1 To r
    num = num + h
    If perfil.Range("E" & num).Value Like "CH*" Then
        perfil.Cells(num, 1).EntireRow.Delete
        h = 0
    Else
        h = 1
    End If
Next

' ------------------  COMPLEMENTA INFORMAÇÕES NA PLANILHA   ------------------
perfil.Activate
If perfil.Range("B13").Value <> Empity Then
    r = perfil.Range("B13", Range("B13").End(xlDown)).Rows.Count
    'Gera o comprimento total da peça
    cel = 13
    For i = 1 To r
        perfil.Cells(cel, 7).Value = perfil.Cells(cel, 6).Value * perfil.Cells(cel, 3).Value
        cel = cel + 1
    Next
    
    'Numera os itens da aba PERFIL
    cel = 13
    j = 1
    For i = 1 To r
        perfil.Cells(cel, 1).Value = j
        cel = cel + 1
        j = j + 1
    Next
End If

'Numera os itens da Aba CHAPA
chapa.Activate
If chapa.Range("B13").Value <> Empity Then
    r2 = chapa.Range("B13", Range("B13").End(xlDown)).Rows.Count
    cel = 13
    j = 1
    For i = 1 To r2
        chapa.Cells(cel, 1).Value = j
        cel = cel + 1
        j = j + 1
    Next
End If

' ------------------  ATUALIZA AS TABELAS DINÂMICAS   ------------------
Dim newRange As Range

perfil.Activate
Set newRange = perfil.Range("A12:K12", Range("A12:K12").End(xlDown))
wb.Sheets("RESUMO_PERFIS").PivotTables("Tabela dinâmica16").ChangePivotCache _
    ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=newRange.Address(, , xlR1C1, True), Version:=xlPivotTableVersion15)
wb.Sheets("RESUMO_PERFIS").PivotTables("Tabela dinâmica16").PivotCache.Refresh

chapa.Activate
Set newRange = chapa.Range("A12:K12", Range("A12:K12").End(xlDown))
wb.Sheets("RESUMO_CHAPAS").PivotTables("Tabela dinâmica16").ChangePivotCache _
    ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=newRange.Address(, , xlR1C1, True), _
                                        Version:=xlPivotTableVersion15)
wb.Sheets("RESUMO_CHAPAS").PivotTables("Tabela dinâmica16").PivotCache.Refresh

wb.Sheets("PRANCHA").Activate
Set newRange = wb.Sheets("PRANCHA").Range("A1", Range("A1").End(xlDown))
wb.Sheets("PRANCHA").PivotTables("Tabela dinâmica1").ChangePivotCache _
    ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=newRange.Address(, , xlR1C1, True), _
                                        Version:=xlPivotTableVersion15)
wb.Sheets("PRANCHA").PivotTables("Tabela dinâmica1").PivotCache.Refresh

' ------------------  FAZ A MARCAÇÃO DOS GRUPOS POR BITOLAS   ------------------
chapa.Activate
If chapa.Range("B13").Value <> Empity Then
    r = chapa.Range("B13", Range("B13").End(xlDown)).Rows.Count
    r = r - 1
    cel = 14
    For i = 1 To r
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
End If

perfil.Activate
If perfil.Range("B13").Value <> Empity Then
    r = perfil.Range("B13", Range("B13").End(xlDown)).Rows.Count
    r = r - 1
    cel = 14
    For i = 1 To r
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
End If

' ------------------  AJUSTA AS LINHAS PARA MANTER UM DEFAULT   ------------------
perfil.Activate
If perfil.Range("B13").Value <> Empity Then
    perfil.Range("A13:S13", Range("A13:S13").End(xlDown)).Select
    Selection.Font.Name = "Arial"
    Selection.Font.Size = 14
    Selection.RowHeight = 23.25
    Selection.Borders.LineStyle = xlContinuous
    Selection.HorizontalAlignment = xlCenter
    perfil.Range("F13").End(xlDown).ClearContents
    perfil.Range("A13").Select
End If

chapa.Activate
If chapa.Range("B13").Value <> Empity Then
    chapa.Range("A13:S13", Range("A13:S13").End(xlDown)).Select
    Selection.Font.Name = "Arial"
    Selection.Font.Size = 14
    Selection.RowHeight = 23.25
    Selection.Borders.LineStyle = xlContinuous
    Selection.HorizontalAlignment = xlCenter
    chapa.Range("A13").Select
End If

' ------------------  COMPLEMENTO FINAL   ------------------
perfil.Activate
perfil.Range("K5:N5").Value = Now()
perfil.Range("K7").Value = "A"

End Sub


