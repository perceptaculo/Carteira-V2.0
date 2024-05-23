VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
'O primeiro comando exclui a possibilidade de alcancarmos a parte de cima de uma tabela'
'E a segunda condicional explicita tambem a exclusao das colunas que geralmente esta contida na linha 1'
    If Dados.Cells(Target.Row, 2) <> "" And Target.Row > 1 Then
        lLinha = Target.Row 'pegar a linha q foi selecionada e armazenar em lLinha'
    End If

    If lLinha > 0 Then
    lsPreencher lLinha
    End If
    
    UserForm1.Show 'mostrar o formulario'
    Cancel = True 'tirar o duplo click da celula'
End Sub

