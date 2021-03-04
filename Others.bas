Attribute VB_Name = "Others"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    ActiveCell.FormulaR1C1 = "=R[-1]C"
    Range("L23").Select
    Selection.ClearContents
    Range("J23").Select
    Selection.Copy
    Range("L21:L37").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    ActiveSheet.Paste

End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'

    Range("W7").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("X7").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Paste
    Range("X7").Select
End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    Rows("1:37").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$W$52").AutoFilter Field:=23, Criteria1:="0"
    Range("W7:W39").Select
    Selection.ClearContents
    Range("W16").Select
    ActiveSheet.Range("$A$1:$W$52").AutoFilter Field:=23
    Range("W24").Select
End Sub
