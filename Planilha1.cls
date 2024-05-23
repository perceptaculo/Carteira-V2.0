VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
    ' Ocultar a faixa de opções
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", False)"
    ThisWorkbook.Save
End Sub

Private Sub Worksheet_Deactivate()
    ' Mostrar a faixa de opções ao sair da planilha
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", True)"
End Sub

