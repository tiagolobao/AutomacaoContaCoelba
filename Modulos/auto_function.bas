Attribute VB_Name = "Módulo3"
Sub AUTO()
Attribute AUTO.VB_Description = "Função do botão para preencher a planilha."
Attribute AUTO.VB_ProcData.VB_Invoke_Func = "A\n14"
'
' AUTO Macro
' Função do botão para preencher a planilha.
'
' Atalho do teclado: Ctrl+Shift+A
'

Dim rng As Range




'PREENCHIMENTO UNIDADE------------------------------------------------
    Sheets("INPUT").Select
    ActiveWindow.SmallScroll DOWN:=-3
    Range("A1").Select
    Cells.Find(What:="UNIVERSIDADE FEDERAL DA BAHIA", After:=ActiveCell, _
        LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Activate
    ActiveCell.Offset(1, 0).Select 'Seleciona celula abaixo
    Selection.Copy
    Sheets("OUTPUT").Select
    Range("D6").Select
    ActiveSheet.Paste
'PREENCHIMENTO DEMANDA ATIVA (CONTRATADA)----------------------------------------
    Sheets("INPUT").Select
    Range("A1").Select
    Cells.Find(What:="Demanda:", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("OUTPUT").Select
    Range("G6").Select
    ActiveSheet.Paste
'------------------------------------------------------------------------------
'PREENCHIMENTO DEMANDA ATIVA ------------------------------------------------
    Sheets("INPUT").Select
    Range("A1").Select
    Cells.Find(What:="Demanda Ativa", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("OUTPUT").Select
    Range("J6").Select
    ActiveSheet.Paste
'-------------------------------------------------------------------------------------------
'PREENCHIMENTO CONSUMO ATIVO NA PONTA-----------------------------------------------------
    Sheets("INPUT").Select
    Range("A1").Select
    Cells.Find(What:="Consumo Ativo Na Ponta", After:=ActiveCell, LookIn:= _
        xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
        xlNext, MatchCase:=False, SearchFormat:=False).Activate
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("OUTPUT").Select
    Range("W6").Select
    ActiveSheet.Paste
'-------------------------------------------------------------------------------------------
    
'PREENCHIMENTO CONSUMO ATIVO FORA PONTA-----------------------------------------------------
    Sheets("INPUT").Select
    Range("A1").Select
    Cells.Find(What:="Consumo Ativo Fora Ponta", After:=ActiveCell, LookIn:= _
        xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
        xlNext, MatchCase:=False, SearchFormat:=False).Activate
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("OUTPUT").Select
    Range("Y6").Select
    ActiveSheet.Paste
'-------------------------------------------------------------------------------------------
'PREENCHIMENTO CONSUMO REATIVO EXC NA PONTA-----------------------------------------------------
    Sheets("INPUT").Select
    Range("A1").Select
    Cells.Find(What:="Consumo Reativo Exc. Na Ponta", After:=ActiveCell, _
        LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Activate
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("OUTPUT").Select
    Range("AA6").Select
    ActiveSheet.Paste
'--------------------------------------------------------------------------------------------
'PREENCHIMENTO CONSUMO REATIVO EXC FORA PONTA-----------------------------------------------------
    Sheets("INPUT").Select
    Range("A1").Select
    Cells.Find(What:="Consumo Reativo Exc. Fora Ponta", After:=ActiveCell, _
        LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Activate
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("OUTPUT").Select
    Range("AC6").Select
    ActiveSheet.Paste
'--------------------------------------------------------------------------------------------
'PREENCHIMENTO CONTRIBUIÇÃO ILUMINAÇÃO PÚBLICA--------------------------------------------------
    Sheets("INPUT").Select
    Application.CutCopyMode = False
    Range("A1").Select
    Cells.Find(What:="Contribuição Iluminação Pública", After:=ActiveCell, _
        LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Activate
    Selection.Copy
    Sheets("OUTPUT").Select
    Range("AE6").Select
    ActiveSheet.Paste
'--------------------------------------------------------------------------------------------------
    
'PREENCHIMENTO TRIBUTO FEDERAL--------------------------------------------------------------------------------------------------
    Sheets("INPUT").Select
    Range("A1").Select
    Cells.Find(What:="Tributo Federal", After:=ActiveCell, LookIn:=xlFormulas _
        , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("OUTPUT").Select
    Range("AG6").Select
'--------------------------------------------------------------------------------------------------
'PREENCHIMENTO DEMANDA MÁXIMA NA PONTA--------------------------------------------------------------------------------------------------
    ActiveSheet.Paste
    Sheets("INPUT").Select
    Range("A1").Select
    Cells.Find(What:="Demanda Máxima Na Ponta", After:=ActiveCell, LookIn:= _
        xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
        xlNext, MatchCase:=False, SearchFormat:=False).Activate
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("OUTPUT").Select
    Range("AI6").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
'--------------------------------------------------------------------------------------------------



'PREENCHIMENTO DEMANDA MÁXIMA FORA PONTA--------------------------------------------------------------------------------------------------
    Sheets("INPUT").Select
    Range("A1").Select
    Cells.Find(What:="Demanda Máxima Fora de Ponta", After:=ActiveCell, LookIn:= _
        xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
        xlNext, MatchCase:=False, SearchFormat:=False).Activate
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("OUTPUT").Select
    Range("AJ6").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
'--------------------------------------------------------------------------------------------------
'PREENCHIMENTO CONSUMO REATIVO NA PONTA--------------------------------------------------------------------------------------------------
    Sheets("INPUT").Select
    Range("A1").Select
    Cells.Find(What:="Consumo Reativo Na Ponta", After:=ActiveCell, LookIn:= _
        xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
        xlNext, MatchCase:=False, SearchFormat:=False).Activate
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("OUTPUT").Select
    Range("AK6").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
'--------------------------------------------------------------------------------------------------
'PREENCHIMENTO CONSUMO REATIVO FORA PONTA--------------------------------------------------------------------------------------------------
    Sheets("INPUT").Select
    Range("A1").Select
    Cells.Find(What:="Consumo Reativo Fora de Ponta", After:=ActiveCell, LookIn:= _
        xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
        xlNext, MatchCase:=False, SearchFormat:=False).Activate
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("OUTPUT").Select
    Range("AL6").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
'--------------------------------------------------------------------------------------------------
'PREENCHIMENTO Fator de carga na ponta e fora de ponta----------------------
    Sheets("INPUT").Select
    Range("A1").Select
    Cells.Find(What:="Fator de carga", After:=ActiveCell, LookIn:=xlFormulas _
        , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        Application.CutCopyMode = False
    ActiveCell.Offset(1, 0).Select
    Selection.Copy
    Sheets("OUTPUT").Select
    Range("AO6").Select
    ActiveSheet.Paste
    Range("AP6").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
'--------------------------------------------------------------
'PREENCHIMENTO N do medidor----------------------
    Sheets("INPUT").Select
    Range("A1").Select
    Cells.Find(What:="Medidor", After:=ActiveCell, LookIn:=xlFormulas _
        , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("OUTPUT").Select
    Range("AS6").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
'--------------------------------------------------------------
'PREENCHIMENTO DATA----------------------
    Sheets("INPUT").Select
    Range("A1").Select
    Cells.Find(What:="Medidor", After:=ActiveCell, LookIn:=xlFormulas _
        , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    Application.CutCopyMode = False
    ActiveCell.Offset(1, 0).Select
    Selection.Copy
    Sheets("OUTPUT").Select
    Range("AM6").Select
    ActiveSheet.Paste
    Range("AN6").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
'--------------------------------------------------------------

'PREENCHIMENTO VALOR DA FATURA----------------------
    Sheets("INPUT").Select
    Range("A1").Select
    Cells.Find(What:="TOTAL A PAGAR", After:=ActiveCell, LookIn:=xlFormulas _
        , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    Application.CutCopyMode = False
    ActiveCell.Offset(1, 0).Select
    Selection.Copy
    Sheets("OUTPUT").Select
    Range("I6").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
'--------------------------------------------------------------

'PREENCHIMENTO INTERRUPÇÃO DE ENERGIA---------------------------------
    Sheets("INPUT").Select
    Range("A1").Select
    Set rng = Cells.Find(What:="Interrupção de energia", After:=ActiveCell, LookIn:=xlFormulas, _
    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
    False, SearchFormat:=False)
    If Not rng Is Nothing Then 'when rng <> nothing means found something'
        rng.Activate
        Selection.Copy
        Sheets("OUTPUT").Select
        Range("AH6").Select
        ActiveSheet.Paste
    Else
        Sheets("OUTPUT").Select
        Range("AH6").Value = 0
    End If
'-------------------------------------------------------------------


   'IPCA (multa+Juros)------------------------------------------------------
    Sheets("INPUT").Select
    Range("A1").Select
    Set rng = Cells.Find(What:="IPCA", After:=ActiveCell, LookIn:=xlFormulas, _
    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
    False, SearchFormat:=False)
    If Not rng Is Nothing Then 'when rng <> nothing means found something'
        rng.Activate
        Selection.Copy
        Sheets("OUTPUT").Select
        Range("AF10").Select
        ActiveSheet.Paste
        Else
        Sheets("OUTPUT").Select
        Range("AF10").Value = 0
    End If
'--------------------------------------------------------------------------
   'Multa COSIP (multa+Juros)------------------------------------------------------
    Sheets("INPUT").Select
    Range("A1").Select
    Set rng = Cells.Find(What:="Multa COSIP", After:=ActiveCell, LookIn:=xlFormulas, _
    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
    False, SearchFormat:=False)
    If Not rng Is Nothing Then 'when rng <> nothing means found something'
        rng.Activate
        Selection.Copy
        Sheets("OUTPUT").Select
        Range("AF11").Select
        ActiveSheet.Paste
        Else
        Sheets("OUTPUT").Select
        Range("AF11").Value = 0
    End If
'--------------------------------------------------------------------------

   'Juros COSIP (multa+Juros)------------------------------------------------------
    Sheets("INPUT").Select
    Range("A1").Select
    Set rng = Cells.Find(What:="Juros COSIP", After:=ActiveCell, LookIn:=xlFormulas, _
    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
    False, SearchFormat:=False)
    If Not rng Is Nothing Then 'when rng <> nothing means found something'
        rng.Activate
        Selection.Copy
        Sheets("OUTPUT").Select
        Range("AF12").Select
        ActiveSheet.Paste
        Else
        Sheets("OUTPUT").Select
        Range("AF12").Value = 0
    End If
    '--------------------------------------------------------------------------

   'Juros por atraso (multa+Juros)------------------------------------------------------
    Sheets("INPUT").Select
    Range("A1").Select
    Set rng = Cells.Find(What:="Juros por atraso", After:=ActiveCell, LookIn:=xlFormulas, _
    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
    False, SearchFormat:=False)
    If Not rng Is Nothing Then 'when rng <> nothing means found something'
        rng.Activate
        Selection.Copy
        Sheets("OUTPUT").Select
        Range("AF13").Select
        ActiveSheet.Paste
        Else
        Sheets("OUTPUT").Select
        Range("AF13").Value = 0
    End If
    '--------------------------------------------------------------------------

   'Multa por atraso (multa+Juros)------------------------------------------------------
    Sheets("INPUT").Select
    Range("A1").Select
    Set rng = Cells.Find(What:="Multa por atraso", After:=ActiveCell, LookIn:=xlFormulas, _
    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
    False, SearchFormat:=False)
    If Not rng Is Nothing Then 'when rng <> nothing means found something'
        rng.Activate
        Selection.Copy
        Sheets("OUTPUT").Select
        Range("AF14").Select
        ActiveSheet.Paste
        Else
        Sheets("OUTPUT").Select
        Range("AF14").Value = 0
    End If
    '--------------------------------------------------------------------------


'PREENCHIMENTO DEMANDA DE ULTRAPASSAGEM NA PONTA---------------------------------
    Sheets("INPUT").Select
    Range("A1").Select
    Set rng = Cells.Find(What:="Demanda Ativa Ultrapassagem", After:=ActiveCell, LookIn:=xlFormulas, _
    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
    False, SearchFormat:=False)
    If Not rng Is Nothing Then 'when rng <> nothing means found something'
        rng.Activate
        Selection.Copy
        Sheets("OUTPUT").Select
        Range("N6").Select
        ActiveSheet.Paste
    Else
        Sheets("OUTPUT").Select
        Range("N6").Value = 0
    End If
'-------------------------------------------------------------------

'PREENCHIMENTO DEMANDA REATIVA EXCEDENTE---------------------------------
    Sheets("INPUT").Select
    Range("A1").Select
    Set rng = Cells.Find(What:="Demanda Reativa Excedente", After:=ActiveCell, LookIn:=xlFormulas, _
    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
    False, SearchFormat:=False)
    If Not rng Is Nothing Then 'when rng <> nothing means found something'
        rng.Activate
        Selection.Copy
        Sheets("OUTPUT").Select
        Range("S6").Select
        ActiveSheet.Paste
    Else
        Sheets("OUTPUT").Select
        Range("S6").Value = 0
    End If
'-------------------------------------------------------------------



'PREENCHIMENTO TIPO DE CONTA----------------------
    Sheets("INPUT").Select
    Range("A1").Select
    Cells.Find(What:="CLASSIFICAÇÃO", After:=ActiveCell, LookIn:=xlFormulas _
        , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    Application.CutCopyMode = False
    ActiveCell.Offset(1, 0).Select
    Selection.Copy
    Sheets("OUTPUT").Select
    Range("F6").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
'--------------------------------------------------------------

'Células que não precisam de informações adicionais
Range("AQ6").Value = 0
Range("AR6").Value = 0
Range("AF6").Value = 0
Range("AD6").Value = 0
Range("AB6").Value = 0
Range("Z6").Value = 0
Range("X6").Value = 0
Range("P6").Value = 0
Range("K6").Value = 0





'PREENCHIMENTO CONTA CONTRATO----------------------
    Sheets("INPUT").Select
    Range("A1").Select
    Cells.Find(What:="CONTA CONTRATO", After:=ActiveCell, LookIn:=xlFormulas _
        , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    Application.CutCopyMode = False
    ActiveCell.Offset(1, 0).Select
    Selection.Copy
    Sheets("OUTPUT").Select
    Range("E6").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
'--------------------------------------------------------------

'PREENCHIMENTO Transformadores--------------
    Dim conta As Variant
    conta = Range("E6").Value
    Sheets("CONTA CONTRATO e NOME").Select
    Range("A1").Select
    Cells.Find(What:=conta, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    Application.CutCopyMode = False
    ActiveCell.Offset(0, 2).Select
    Selection.Copy
    Sheets("OUTPUT").Select
    Range("AT6").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Sheets("CONTA CONTRATO e NOME").Select
    ActiveCell.Offset(0, 1).Select
    Selection.Copy
    Sheets("OUTPUT").Select
    Range("AU6").Select
    ActiveSheet.Paste
'------------------------------------------------------------------------------

'ALGORITIMO PARA VER SE EXISTE OU NÃO
'    Sheets("INPUT").Select
'    Range("A1").Select
'    Dim rng As Range
'    Set rng = Cells.Find(What:="Tiago", After:=ActiveCell, LookIn:=xlFormulas, _
'    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
'    False, SearchFormat:=False)
'    If Not rng Is Nothing Then 'when rng <> nothing means found something'
'        rng.Activate
'        MsgBox "existe"
'        Else
'        MsgBox "nao existe"
'    End If
'-------------------------------------------------------------



MsgBox "Preenchimento realizado com sucesso."
End Sub
