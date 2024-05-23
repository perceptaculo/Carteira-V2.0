VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   ClientHeight    =   14415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   27765
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cButton As New clsBtn 'Importei a classe do botao'

Private Sub btnAllCartao_Click()
    Call filtrarDados
End Sub

Private Sub btnAllModalidade_Click()
    Call filtrarDados
End Sub

Private Sub btnAlltipos_Click()
    Call filtrarDados
End Sub

Private Sub btnAtivo_Click()
    Call filtrarDados
End Sub

Private Sub btnCredito_Click()
    Call filtrarDados
End Sub

Private Sub btnDebito_Click()
    Call filtrarDados
End Sub

Private Sub btnFilterClear_Click()
    UserForm1.btnAllCartao.Value = True
    UserForm1.btnAllModalidade.Value = True
    UserForm1.btnAllTipos.Value = True
    UserForm1.btnAllTodos.Value = True
    Call filtrarDados
End Sub

Private Sub btnLeonel_Click()
    Call filtrarDados
End Sub

Private Sub btnML_Click()
    Call filtrarDados
End Sub

Private Sub btnNU_Click()
    Call filtrarDados
End Sub

Private Sub btnPaola_Click()
    Call filtrarDados
End Sub

Private Sub btnPassivo_Click()
    Call filtrarDados
End Sub

Private Sub btnQuemTodos_Click()
    Call filtrarDados
End Sub

Private Sub TabelaDados_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    lLinha = TabelaDados.ListIndex + 2
    lsPreencher lLinha
End Sub

Private Sub UserForm_Initialize()
    ' Supondo que cButton é um objeto e classButton é um método que precisa ser chamado
    cButton.classButton Me 'Userform1 = me'
    
    ' Declarando lista_txts como uma variante para conter o array
    Dim lista_txts As Variant
    lista_txts = Array("txtItem", "txtData", "txtTipo", "txtCartao", "txtModalidade", "txtQuem")
    
    ' Chamando a função dropdown para cada item da lista
    Dim i As Integer 'declarando i
    Dim caixaDrop As String
    For i = LBound(lista_txts) To UBound(lista_txts) 'iterando sobre a lista do comeco (lbound) ao fim (ubound) de maneira que i retorne o numero da posicao de cada lista, poderia deixar isso fixo, de 1 a 7 por exemplo
         'chamando a funcao para cada elemento
         caixaDrop = lista_txts(i)
         Call dropdown(caixaDrop)
    Next i
    
    Call mostrarTabela
    Call filtrarDados
    
    
End Sub

Private Sub btnExcluir_Click()
    If lsExcluir Then
        UserForm1.Hide
    End If
End Sub

Private Sub btnLimpar_Click()
    lsLimpar
End Sub

Private Sub btnNovo_Click()
    lLinha = 0
    lsLimpar
End Sub
Private Sub btnSalvar_Click()
    'Verificar se algum campo está vazio'
    With UserForm1
        If .txtItem = "" Or .txtSubitem = "" Or .txtData = "" Or .txtValor = "" Or .txtCartao = "" Or .txtModalidade = "" Or .txtTipo = "" Or .txtQuem = "" Then
        ' Encontra a próxima linha vazia
            MsgBox "Existem campos vazios que devem obrigatoriamente serem preenchidos.", vbOK
            
            If .txtItem = "" Then
                ListaVazios.Add "Item"
            End If
            
            If .txtSubitem = "" Then
                ListaVazios.Add "Subitem"
            End If
            
            If .txtData = "" Then
                ListaVazios.Add "Data"
            End If
            
            If .txtValor = "" Then
                ListaVazios.Add "Valor"
            End If
            
            If .txtCartao = "" Then
                ListaVazios.Add "Cartao"
            End If
            
            If .txtModalidade = "" Then
                ListaVazios.Add "Modalidade"
            End If
            
            If .txtTipo = "" Then
                ListaVazios.Add "Tipo"
            End If
            
            If .txtQuem = "" Then
                ListaVazios.Add "Quem"
            End If
            
            
            Dim mensagem As String
            Dim i As Integer
            
            mensagem = "Os seguintes itens estão vazios: "
            
            For i = 1 To ListaVazios.Count
                mensagem = mensagem & ListaVazios(i)
                
                ' Adiciona uma vírgula e espaço após cada item, exceto o último
                If i < ListaVazios.Count Then
                    mensagem = mensagem & ", "
                End If
            Next i

            MsgBox mensagem

            
        Else
            If lLinha = 0 Then
                lLinha = Dados.Cells(Dados.Rows.Count, 2).End(xlUp).Row + 1
            End If
        
            ' Chama a sub-rotina para cadastrar os dados na linha lLinha
            lsCadastrar lLinha
            ' Limpa os campos do formulário
        End If
    End With
    
End Sub


Private Sub txtValor_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

If KeyAscii = 47 Or KeyAscii = 46 Then KeyAscii = 0
If KeyAscii < 44 Or KeyAscii > 57 Then KeyAscii = 0

End Sub

Private Sub UserForm_Terminate()
    Sheets("Dados").AutoFilterMode = False
End Sub
