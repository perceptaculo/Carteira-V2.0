Attribute VB_Name = "Módulo2"
Sub lsAbaixar()
Attribute lsAbaixar.VB_ProcData.VB_Invoke_Func = " \n14"
'
' lsAbaixar Macro
'

'
    ActiveWindow.SmallScroll Down:=-12
    Range("A1").Select
    Selection.End(xlDown).Select
    Range("B1048576").Select
    Selection.End(xlUp).Select
End Sub

Sub AbrirBI()
    Dim caminho As String
    caminho = "C:\Users\Code\OneDrive\Economy\Reports.pbix"
    Shell "cmd /c start """" """ & caminho & """", vbNormalFocus
End Sub

Sub irDados()
'
' irDados Macro
'

'
    Sheets("Dados").Select
    Range("N4").Select
    ActiveWindow.ScrollColumn = 1
    Range("A1").Select
    Selection.End(xlDown).Select
    Range("B1048576").Select
    Selection.End(xlUp).Select
End Sub

Sub irPlanejamentos()
'
' irDados Macro
'

'
    Sheets("Planejamento").Select
    
End Sub

Sub abrirTool()
    UserForm1.Show
End Sub


