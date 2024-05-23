Attribute VB_Name = "Módulo5"
Sub filtrarDados()
    Dim wsDados As Worksheet
    Dim wsDadosSelecionados As Worksheet
    Set wsDados = Sheets("Dados")
    Set wsDadosSelecionados = Sheets("DadosSelecionados")
    
    ' Retirar os filtros existentes
    On Error Resume Next
    wsDados.ShowAllData
    On Error GoTo 0
    
    ' Verificar os filtros a serem aplicados
    Dim filtroAtivo As Boolean
    Dim filtroPassivo As Boolean
    Dim filtroLeonel As Boolean
    Dim filtroPaola As Boolean
    Dim filtroCredito As Boolean
    Dim filtroDebito As Boolean
    Dim filtroNU As Boolean
    Dim filtroML As Boolean
    
    filtroAtivo = UserForm1.btnAtivo.Value
    filtroPassivo = UserForm1.btnPassivo.Value
    filtroLeonel = UserForm1.btnLeonel.Value
    filtroPaola = UserForm1.btnPaola.Value
    filtroCredito = UserForm1.btnCredito.Value
    filtroDebito = UserForm1.btnDebito.Value
    filtroNU = UserForm1.btnNU.Value
    filtroML = UserForm1.btnML.Value
    
    ' Aplicar os filtros de acordo com os botões selecionados
    With wsDados.UsedRange
        If filtroAtivo Then
            .AutoFilter Field:=8, Criteria1:="Ativo"
        ElseIf filtroPassivo Then
            .AutoFilter Field:=8, Criteria1:="Passivo"
        End If
        
        If filtroLeonel Then
            .AutoFilter Field:=9, Criteria1:="Leonel"
        ElseIf filtroPaola Then
            .AutoFilter Field:=9, Criteria1:="Paola"
        End If
        
        If filtroCredito Then
            .AutoFilter Field:=7, Criteria1:="Crédito"
        ElseIf filtroDebito Then
            .AutoFilter Field:=7, Criteria1:="Débito"
        End If
        
        If filtroNU Then
            .AutoFilter Field:=6, Criteria1:="NU"
        ElseIf filtroML Then
            .AutoFilter Field:=6, Criteria1:="ML"
        End If
    End With
    
    ' Limpar a planilha de destino antes de copiar os dados filtrados
    wsDadosSelecionados.UsedRange.Clear
    
    ' Copiar apenas os dados visíveis após o filtro
    On Error Resume Next ' Caso não haja células visíveis
    wsDados.UsedRange.SpecialCells(xlCellTypeVisible).Copy
    On Error GoTo 0
    
    wsDadosSelecionados.Range("A1").PasteSpecial xlPasteAll
    
    ' Retirar os filtros
    On Error Resume Next
    wsDados.ShowAllData
    On Error GoTo 0
    
    ' Determinar a última linha da planilha de dados selecionados
    Dim ultimaRow As Long
    ultimaRow = wsDadosSelecionados.Range("A" & Rows.Count).End(xlUp).Row
    
    ' Configurar o UserForm com os dados filtrados
    With UserForm1.CaixadeDados
        .ColumnCount = 10
        .ColumnHeads = True
        .ColumnWidths = "60;120;120;80;95;70;95;70;70"
        .RowSource = "DadosSelecionados!A2:J" & ultimaRow
    End With
End Sub

