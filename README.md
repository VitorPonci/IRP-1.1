Sub PreencherTabela()

    Dim wsFonte As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim dataAtual As Date
    Dim anoAtual As Integer
    Dim anoMin As Integer: anoMin = 2010
    Dim anoMax As Integer: anoMax = 0

    Dim dictPetroleo As Object
    Dim dictGas As Object
    Dim dictTotal As Object

    Set wsFonte = ThisWorkbook.Sheets("Dados de Produção")
    Set wsDestino = ThisWorkbook.Sheets("Graficos")

    Set dictPetroleo = CreateObject("Scripting.Dictionary")
    Set dictGas = CreateObject("Scripting.Dictionary")
    Set dictTotal = CreateObject("Scripting.Dictionary")

    ultimaLinha = wsFonte.Cells(wsFonte.Rows.Count, 1).End(xlUp).Row

    ' Agrupar dados mensais por ano
    For i = 2 To ultimaLinha
        If IsDate(wsFonte.Cells(i, 1).Value) Then
            dataAtual = wsFonte.Cells(i, 1).Value
            anoAtual = Year(dataAtual)

            If anoAtual >= anoMin Then
                dictPetroleo(anoAtual) = dictPetroleo(anoAtual) + Val(wsFonte.Cells(i, 2).Value)
                dictGas(anoAtual) = dictGas(anoAtual) + Val(wsFonte.Cells(i, 3).Value)
                dictTotal(anoAtual) = dictTotal(anoAtual) + Val(wsFonte.Cells(i, 4).Value)

                If anoAtual > anoMax Then anoMax = anoAtual
            End If
        End If
    Next i

    ' Legendas fixas na coluna B
    wsDestino.Cells(1, 2).Value = "Produção / Ano"
    wsDestino.Cells(2, 2).Value = "Petróleo (Mboe/d)"
    wsDestino.Cells(3, 2).Value = "Gás Natural (Mboe/d)"
    wsDestino.Cells(4, 2).Value = "Produção Total (Mboe/d)"

    ' Cabeçalhos de anos dinâmicos: C1 em diante
    Dim col As Long, anoCol As Long
    For anoCol = anoMin To anoMax
        wsDestino.Cells(1, (anoCol - anoMin + 3)).Value = anoCol
    Next anoCol

    ' Preencher dados de forma dinâmica nas colunas
    For anoAtual = anoMin To anoMax
        col = 3 + (anoAtual - anoMin)
        wsDestino.Cells(2, col).Value = dictPetroleo(anoAtual)
        wsDestino.Cells(3, col).Value = dictGas(anoAtual)
        wsDestino.Cells(4, col).Value = dictTotal(anoAtual)
    Next anoAtual

    MsgBox "Tabela preenchida automaticamente até o ano " & anoMax & "!"

End Sub
