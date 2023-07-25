# controle-de-estoq
Planilha para controle de estoque


Public Function lfUltimaLinhaAtiva(ByVal lrng As Range) As Long
    Dim lWorksheet As String
    Dim lColuna    As Long
    
    lWorksheet = lrng.Worksheet.Name
    lColuna = lrng.Column
    
    lfUltimaLinhaAtiva = Worksheets(lWorksheet).Cells(Worksheets(lWorksheet).Rows.Count, lColuna).End(xlUp).Row
    
End Function

Public Sub lsEntradaSaida()
    Dim lLinha      As Long
    
    lLinha = lfUltimaLinhaAtiva(RegistrodeInventario.Range("F1")) + 1
    
    'Código do produto
    RegistrodeInventario.Cells(lLinha, 7).Value = EntradaeSaida.Range("G6").Value
    'Tipo de movimentação
    RegistrodeInventario.Cells(lLinha, 8).Value = EntradaeSaida.Range("G12").Value
    'Operação fiscal
    RegistrodeInventario.Cells(lLinha, 13).Value = EntradaeSaida.Range("G11").Value
    'Data
    RegistrodeInventario.Cells(lLinha, 9).Value = EntradaeSaida.Range("g13").Value
    'Quantidade
    RegistrodeInventario.Cells(lLinha, 10).Value = EntradaeSaida.Range("g14").Value
    'Valor unitário
    RegistrodeInventario.Cells(lLinha, 11).Value = EntradaeSaida.Range("g15").Value
    'Valor total
    RegistrodeInventario.Cells(lLinha, 12).Value = EntradaeSaida.Range("g16").Value
    'Série
    RegistrodeInventario.Cells(lLinha, 14).Value = EntradaeSaida.Range("g17").Value
    'Nota fiscal
    RegistrodeInventario.Cells(lLinha, 15).Value = EntradaeSaida.Range("g18").Value
    'Fornecedor
    RegistrodeInventario.Cells(lLinha, 16).Value = EntradaeSaida.Range("g19").Value
    'Complemento
    RegistrodeInventario.Cells(lLinha, 17).Value = EntradaeSaida.Range("g20").Value
    
    lsLimpar
    
    EntradaeSaida.Range("G6").Select
    
End Sub

Public Sub lsLimpar()
    'Código do produto
    EntradaeSaida.Range("g6").Value = ""
    'Operação Fiscal
    EntradaeSaida.Range("g11").Value = ""
    'vários
    EntradaeSaida.Range("g13:g15").Value = ""
    'Vários
    EntradaeSaida.Range("g17:g20").Value = ""
End Sub

Public Sub lsAtualizar()
    ConsultadeEstoque.ListObjects("tConsultadeEstoque").Range.ListObject.QueryTable.Refresh BackgroundQuery:=False
End Sub
