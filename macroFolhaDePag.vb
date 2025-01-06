Sub PreencherDadosGenerico()
    Dim wsOrigem As Worksheet, wsDestino As Worksheet
    Dim ultimaLinhaOrigem As Long, ultimaColunaOrigem As Long
    Dim linhaDestino As Long
    Dim i As Long, j As Long
    Dim valorPago As String
    Dim cabecalhoColuna As String
    Dim matricula As String
    Dim quantidade As String
    Dim regex As Object

    ' Configurar regex para extrair os dois primeiros números
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "\d{1,2}" ' Capturar até os dois primeiros números consecutivos
    regex.Global = False

    ' Definir as planilhas
    Set wsOrigem = ThisWorkbook.Sheets(1) ' Primeira aba como origem
    Set wsDestino = ThisWorkbook.Sheets(2) ' Segunda aba como destino

    ' Encontrar a última linha e coluna da aba origem
    ultimaLinhaOrigem = wsOrigem.Cells(wsOrigem.Rows.Count, "A").End(xlUp).Row
    ultimaColunaOrigem = wsOrigem.Cells(1, wsOrigem.Columns.Count).End(xlToLeft).Column

    ' Localizar a primeira linha vazia na aba destino
    linhaDestino = wsDestino.Cells(wsDestino.Rows.Count, "B").End(xlUp).Row + 1

    ' Começar a varredura da aba origem
    For i = 2 To ultimaLinhaOrigem ' Começa da linha 2 (ignora o cabeçalho)
        matricula = Trim(wsOrigem.Cells(i, 1).Text) ' Coluna A como "MATRÍCULA"
        
        ' Obter os dois primeiros números da coluna "HORAS/DIAS" (Coluna F)
        quantidade = wsOrigem.Cells(i, 6).Text
        If regex.Test(quantidade) Then
            quantidade = regex.Execute(quantidade)(0) ' Extrair o valor capturado
        Else
            quantidade = "" ' Caso não haja números, deixar vazio
        End If

        ' Percorrer todas as colunas de dados (da G em diante)
        For j = 7 To ultimaColunaOrigem
            cabecalhoColuna = Trim(wsOrigem.Cells(1, j).Text) ' Nome da coluna
            valorPago = Trim(wsOrigem.Cells(i, j).Text) ' Valor como texto para manter o formato

            ' Verificar se "valorPago" é numérico e não está vazio
            If IsNumeric(valorPago) And valorPago <> "" Then
                ' Adicionar dados na aba destino
                wsDestino.Cells(linhaDestino, "B").Value = matricula           ' MATRÍCULA
                wsDestino.Cells(linhaDestino, "C").Value = "'" & cabecalhoColuna ' Código da verba como texto exato
                wsDestino.Cells(linhaDestino, "E").Value = quantidade         ' QUANTIDADE
                wsDestino.Cells(linhaDestino, "F").Value = valorPago          ' VALOR PAGO (como texto)

                ' Incrementar a linha na aba destino
                linhaDestino = linhaDestino + 1
            End If
        Next j
    Next i

    MsgBox "Processamento concluído com sucesso!", vbInformation
End Sub
