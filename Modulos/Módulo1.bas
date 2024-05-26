Attribute VB_Name = "Módulo1"
'Informacao do numero da linha, variavel global chamada lLinha'
Global lLinha As Long
Global ListaVazios As New Collection
Global VaziosString
'Funcao para limpar os campos do formulario'
Public Sub lsLimpar()
'apenas para nao ter que ficar referenciando de qual objeto eu devo pegar o atributo, se nao teria q ficar passando UserForm.txtItem etc..'
    With UserForm1
        .txtItem = ""
        .txtSubitem = ""
        .txtData = ""
        .txtCartao = ""
        .txtTipo = ""
        .txtQuem = ""
        .txtModalidade = ""
        .txtValor = ""
        .txtStatus = ""
    End With
    lsLinha = 0
End Sub
'Já validado, está funcionando.'

Public Sub lsCadastrar(ByVal lLinha As Long)

    ' Jogar a linha nova em lLinha
    With Dados
        .Cells(lLinha, 3).Value = UserForm1.txtItem
        .Cells(lLinha, 4).Value = UserForm1.txtSubitem
        .Cells(lLinha, 5).Value = Format(UserForm1.txtData, "mmm/yy")

        ' Tratar txtValor
        If UserForm1.txtValor = "" Then
            .Cells(lLinha, 6).Value = 0
        Else
            .Cells(lLinha, 6).Value = CDbl(Format(UserForm1.txtValor, "###0.00"))
        End If

        .Cells(lLinha, 7).Value = UserForm1.txtCartao
        .Cells(lLinha, 8).Value = UserForm1.txtModalidade
        .Cells(lLinha, 9).Value = UserForm1.txtTipo
        .Cells(lLinha, 10).Value = UserForm1.txtQuem
        .Cells(lLinha, 11).Value = UserForm1.txtStatus
    End With
End Sub

Public Sub lsPreencher(ByVal lLinha As Long)
With Dados
  UserForm1.txtItem = .Cells(lLinha, 3).Value
  UserForm1.txtSubitem = .Cells(lLinha, 4).Value
  
  UserForm1.txtData = Format(.Cells(lLinha, 5).Value, "mmm/yy")
  
  UserForm1.txtValor = .Cells(lLinha, 6).Value
  UserForm1.txtCartao = .Cells(lLinha, 7).Value
  UserForm1.txtModalidade = .Cells(lLinha, 8).Value
  UserForm1.txtTipo = .Cells(lLinha, 9).Value
  UserForm1.txtQuem = .Cells(lLinha, 10).Value
  UserForm1.txtStatus = .Cells(lLinha, 11).Value
End With
End Sub

Public Function lsExcluir() As Boolean
    If MsgBox("Deseja excluir este registro?", vbYesNo) = vbYes Then
        Dados.Cells(lLinha, 2).EntireRow.Delete
        lsLimpar
        lsExcluir = True
    Else
        lsExcluir = False
    End If
End Function


