Attribute VB_Name = "Módulo1"
Sub Geral()

    Application.ScreenUpdating = False

    Call Base_tratada
    Call Base_de_resultados

    Sheets("MACROS").Select
    Range("B7").Select

    Application.ScreenUpdating = True

End Sub

Sub Base_tratada()
Attribute Base_tratada.VB_ProcData.VB_Invoke_Func = " \n14"

    Application.ScreenUpdating = False

    Sheets("BASE TRATADA").Select
    Range("R6").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range("R7").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("R6").Select
    Application.CutCopyMode = False
    Range("B6").Select
    
    Application.ScreenUpdating = True

End Sub

Sub Base_de_resultados()

    Application.ScreenUpdating = False

    'Tipo Var
    Dim atual As Double
    Dim final As Double
    Dim linhai As Double
    Dim linhaf As Double
    
    atual = Abs(Worksheets("BASE DE RESULTADOS").Range("C1").Value)
    final = Abs(Worksheets("BASE DE RESULTADOS").Range("B1").Value)
 
    Do While atual > final
        Sheets("BASE DE RESULTADOS").Select
        Range("B4").Select
        Selection.End(xlDown).Select
        linhai = ActiveCell.Row - 1
        linhaf = Range("B4").Row + 2
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
        atual = Abs(Worksheets("BASE DE RESULTADOS").Range("C1").Value)
        final = Abs(Worksheets("BASE DE RESULTADOS").Range("B1").Value)
    Loop

    Sheets("BASE DE RESULTADOS").Select
    Range("B4").Select
    Selection.End(xlDown).Select
    linhai = ActiveCell.Row - 1
    
    If Range("C1").Value > 0 Then
        linhaf = linhai - Range("C1").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
    ElseIf Range("C1").Value < 0 Then
        linhaf = linhai + Range("C1").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Delete Shift:=xlUp
    Else
    End If
  
    Application.CutCopyMode = False
    Range("B4").Select
    
    Sheets("BASE TRATADA").Select
    Range("U5").Select
    ActiveSheet.Range("$B$5:$AM$10000").AutoFilter Field:=20, Criteria1:="=1", _
        Operator:=xlAnd
    Range("Y5").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("BASE DE RESULTADOS").Select
    Range("B3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B4").Select
    Application.CutCopyMode = False
    Sheets("BASE TRATADA").Select
    Range("U5").Select
    ActiveSheet.Range("$B$5:$AM$10000").AutoFilter Field:=20
    Range("B6").Select
    Sheets("BASE DE RESULTADOS").Select
    Range("B4").Select
    ActiveWorkbook.RefreshAll
    
    Application.ScreenUpdating = True

End Sub


Sub Arquivo_de_envio()
Attribute Arquivo_de_envio.VB_ProcData.VB_Invoke_Func = " \n14"

    Application.ScreenUpdating = False
    
    ActiveWorkbook.Save
    
    ActiveWorkbook.SaveAs Filename:= _
        ActiveWorkbook.Path & "\" & Worksheets("MACROS").Range("C11").Value & " - Gestão de Manutenção de Base ADQ - Dados até dia " & Worksheets("MACROS").Range("C12").Value & ".xlsm" _
        , FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
        
    Sheets("QUADRO DE PERFORMANCE").Select
    Cells.Select
    Range("A2").Activate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B6").Select
    Application.CutCopyMode = False
    ActiveWindow.DisplayHeadings = False
    Sheets("BASE DE RESULTADOS").Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B4").Select
    Application.CutCopyMode = False
    Range("B1:C1").Select
    Selection.ClearContents
    Range("B4").Select
    ActiveWindow.DisplayHeadings = False
    Sheets(Array("MACROS", "FAT. ADQ", "DADOS GERAIS", "BASE TRATADA", "TD", "GRÁFICOS") _
        ).Select
    Sheets("GRÁFICOS").Activate
    ActiveWindow.SelectedSheets.Delete
    Sheets("QUADRO DE PERFORMANCE").Select
    ActiveWorkbook.Save
    
    Application.ScreenUpdating = True

End Sub
