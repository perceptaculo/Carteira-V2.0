Attribute VB_Name = "Módulo4"
Sub mostrarTabela()
    lastRow = Dados.Cells(Dados.Rows.Count, 2).End(xlUp).Row + 1
    UserForm1.TabelaDados.ColumnCount = 10
    UserForm1.TabelaDados.ColumnHeads = True
    UserForm1.TabelaDados.ColumnWidths = "60;120;120;80;95;70;95;70;70"
    UserForm1.TabelaDados.RowSource = "Dados!B2:K" & lastRow
End Sub
