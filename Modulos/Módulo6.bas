Attribute VB_Name = "Módulo6"
Sub Soma()
    Dim total As Double
    Dim i As Long
    Dim Final As Long
    
    ' Inicializar o total como zero
    total = 0
    
    With UserForm1
        Final = .CaixadeDados.ListCount - 1
        ' Iterar sobre os itens na CaixadeDados (supondo que a CaixadeDados se chama CaixadeDados1 e a coluna desejada é a segunda coluna)
        For i = 0 To Final
            ' Adicionar o valor da coluna à soma total
            total = total + CDbl(.CaixadeDados.List(i, 4)) ' Substitua "1" pelo índice da coluna desejada, lembrando que a indexação em VBA começa em 0
        Next i
        
        ' Atribuir o valor da soma ao TextBox (supondo que o TextBox se chama txtTotal)
        .txtTotal = "R$ " & Format(Round(total, 2))
    End With
End Sub

