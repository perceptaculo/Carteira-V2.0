Attribute VB_Name = "M�dulo6"
Sub Soma()
    Dim total As Double
    Dim i As Long
    Dim Final As Long
    
    ' Inicializar o total como zero
    total = 0
    
    With UserForm1
        Final = .CaixadeDados.ListCount - 1
        ' Iterar sobre os itens na CaixadeDados (supondo que a CaixadeDados se chama CaixadeDados1 e a coluna desejada � a segunda coluna)
        For i = 0 To Final
            ' Adicionar o valor da coluna � soma total
            total = total + CDbl(.CaixadeDados.List(i, 4)) ' Substitua "1" pelo �ndice da coluna desejada, lembrando que a indexa��o em VBA come�a em 0
        Next i
        
        ' Atribuir o valor da soma ao TextBox (supondo que o TextBox se chama txtTotal)
        .txtTotal = "R$ " & Format(Round(total, 2))
    End With
End Sub

