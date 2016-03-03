Attribute VB_Name = "Módulo4"
Sub Clear()
Attribute Clear.VB_ProcData.VB_Invoke_Func = "C\n14"
'
' Clear Macro
'
' Atalho do teclado: Ctrl+Shift+C
'
    Rows("6:6").Select
    Selection.ClearContents
    
    Rows("10:14").Select
    Selection.ClearContents
End Sub
