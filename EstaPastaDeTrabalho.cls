VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstaPastaDeTrabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    ' Ir para Planilha1
    Worksheets("Dados").Activate
    
    ' Ir para Planilha2
    Worksheets("Início").Activate
    UserForm1.Show
    
End Sub
