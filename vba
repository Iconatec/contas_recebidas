Sub Contas_recebidas()

Dim ws As Worksheet
Dim ultimalinha As Long
Dim final As Long
Dim i As Long
Dim txt As Long

Set ws = ActiveSheet

'Tirar a mescla
ws.Cells.UnMerge

'Tirar a quebra
ActiveSheet.Cells.WrapText = False

'Apagar as linhas e a imagem
For i = 1 To 4
    ws.Rows(1).Delete
    Next i
ws.Shapes.Range(Array("Picture 1")).Delete

'Deletar últimas 4 linhas
ultimalinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
final = ultimalinha - 4
ws.Rows(final & ":" & ultimalinha).Delete

'Tirar número entre parênteses
ultimalinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row - 2
For i = 7 To ultimalinha
    celula = ws.Cells(i, "A").Value
    txt = InStr(celula, "(")
    ws.Cells(i, "A").Value = Left(celula, txt - 2)
    Next i
    
'Mudar o cabeçalho de lugar
Range("N1:O4").Value = Range("A1:B4").Value
Range("A1:B4").ClearContents
    
'Deletar colunas
ws.Columns("B:E").Delete
ws.Columns("C").Delete
ws.Columns("E").Delete

'Formatar texto, colunas e  arrumar o cabeçalho
Range("A1:B4").Value = Range("H1:I4").Value
Range("H1:I4").ClearContents
ws.Columns("C:G").AutoFit
Columns("B").ColumnWidth = 11

'Congelar linha 6 e deletar a linha 5
Rows("7:7").Select
ActiveWindow.FreezePanes = True

End Sub
