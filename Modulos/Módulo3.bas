Attribute VB_Name = "Módulo3"
Global param As String
Global lastRow As Long

Sub dropdown(txt As String)
    ' Declarar UserForm1 explicitamente
    Dim uf As Object
    Set uf = UserForm1
    
    ' Declarar a planilha explicitamente
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dropdown")
    
    With uf
        Select Case txt
            Case "txtItem"
                ' Encontra a última linha na coluna B
                lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
                ' Define o RowSource do controle txtItem
                .Controls(txt).RowSource = "Dropdown!B2:B" & lastRow
                
            Case "txtItem2"
                ' Encontra a última linha na coluna B
                lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
                ' Define o RowSource do controle txtItem
                .Controls(txt).RowSource = "Dropdown!B2:B" & lastRow
                
            Case "txtSubitem"
                ' Encontra a última linha na coluna C
                lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
                ' Define o RowSource do controle txtSubitem
                .Controls(txt).Value = "Dropdown!C2:C" & lastRow
                
            Case "txtData"
                ' Encontra a última linha na coluna D
                lastRow = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
                ' Define o RowSource do controle txtMes
                .Controls(txt).RowSource = "Dropdown!D14:D" & lastRow
                
            Case "txtData2"
                ' Encontra a última linha na coluna D
                lastRow = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
                ' Define o RowSource do controle txtMes
                .Controls(txt).RowSource = "Dropdown!D14:D" & lastRow

            Case "txtTipo"
                ' Encontra a última linha na coluna E
                lastRow = ws.Cells(ws.Rows.Count, 5).End(xlUp).Row
                ' Define o RowSource do controle txtTipo
                .Controls(txt).RowSource = "Dropdown!E2:E" & lastRow
                
            Case "txtCartao"
                ' Encontra a última linha na coluna F
                lastRow = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row
                ' Define o RowSource do controle txtCartao
                .Controls(txt).RowSource = "Dropdown!F2:F" & lastRow
                
            Case "txtModalidade"
                ' Encontra a última linha na coluna G
                lastRow = ws.Cells(ws.Rows.Count, 7).End(xlUp).Row
                ' Define o RowSource do controle txtModalidade
                .Controls(txt).RowSource = "Dropdown!G2:G" & lastRow
                
            Case "txtQuem"
                ' Encontra a última linha na coluna H
                lastRow = ws.Cells(ws.Rows.Count, 8).End(xlUp).Row
                ' Define o RowSource do controle txtQuem
                .Controls(txt).RowSource = "Dropdown!H2:H" & lastRow
        End Select
    End With
End Sub


