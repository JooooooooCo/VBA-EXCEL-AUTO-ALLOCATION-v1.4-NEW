Attribute VB_Name = "STARTALLOCATION"
Public planaloca As String
Public planmacro As String
Public trade As String


Sub STARTALLOCATION()
Attribute STARTALLOCATION.VB_ProcData.VB_Invoke_Func = " \n14"

If Range("D10") = "" Then
    MsgBox ("ERRO: NOME DA PLANILHA DE ALOCAÇÃO NÃO INFORMADA!")
    Exit Sub
End If

For Each wb In Application.Workbooks
If InStr(LCase(wb.Name), LCase(Range("D10"))) = 0 Then
planaberta = False
Else
planaberta = True
Exit For
End If
Next wb
If planaberta = False Then
    MsgBox ("ERRO: VERIFIQUE SE PLANILHA DE ALOCAÇÃO ESTÁ ABERTA E SE O NOME INFORMADO ESTÁ CORRETO!")
    Exit Sub
End If

If Range("D12") = "" Then
    MsgBox ("ERRO: PASTA COM BOOKING LIST NÃO INFORMADA!")
    Exit Sub
End If

If Range("D14") = "" Then
    MsgBox ("ERRO: TRADE NÃO INFORMADA!")
    Exit Sub
End If

'--------------------- INSERINDO BKG LIST -------------------
'--------------------- INSERINDO BKG LIST -------------------
'--------------------- INSERINDO BKG LIST -------------------

Dim sling As String
Dim ultimalinha As Long
Dim bookinglist As Variant
Dim txtnaoencontrado As VbMsgBoxResult
Dim caminho As String
Dim line As Long
Dim excluishipperrow As Long
Dim inseribkglistrow As Long
Dim inseribkglistcol As String

planmacro = ThisWorkbook.Name

Windows(planmacro).Activate

planaloca = Sheets("MACRO").Range("D10")

caminho = Sheets("MACRO").Range("D12") & "\"
trade = Sheets("MACRO").Range("D14")


Sheets("Allocation").Select
Cells.Select
Selection.ClearContents



If trade = "Asia&Amsul" Then
    inseribkglistcol = "B"
End If

If trade = "Euromed" Then
    inseribkglistcol = "C"
End If
    
If trade = "Americas" Then
    inseribkglistcol = "D"
End If

For inseribkglistrow = 18 To 40
    
    sling = Sheets("MACRO").Range(inseribkglistcol & inseribkglistrow)
    
    If sling <> "" Then
        
            bookinglist = caminho & sling & ".txt"
            If Dir(bookinglist) <> vbNullString Then
            
                    ultimalinha = Cells(Rows.Count, "A").End(xlUp).row
            
                    ultimalinha = ultimalinha + 5
            
                    Range("A" & ultimalinha).Select
            
                        With ActiveSheet.QueryTables.Add(Connection:= _
                        "TEXT;" & caminho & sling & ".txt", Destination:=Range("A" & ultimalinha))
                            .Name = "TEXTO"
                            .FieldNames = True
                            .RowNumbers = False
                            .FillAdjacentFormulas = False
                            .PreserveFormatting = True
                            .RefreshOnFileOpen = False
                            .RefreshStyle = xlInsertDeleteCells
                            .SavePassword = False
                            .SaveData = True
                            .AdjustColumnWidth = True
                            .RefreshPeriod = 0
                            .TextFilePromptOnRefresh = False
                            .TextFilePlatform = 850
                            .TextFileStartRow = 1
                            .TextFileParseType = xlFixedWidth
                            .TextFileTextQualifier = xlTextQualifierDoubleQuote
                            .TextFileConsecutiveDelimiter = False
                            .TextFileTabDelimiter = True
                            .TextFileSemicolonDelimiter = False
                            .TextFileCommaDelimiter = False
                            .TextFileSpaceDelimiter = False
                            .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
                            .TextFileFixedColumnWidths = Array(5, 6, 11, 23, 23, 21, 5, 4, 4, 4, 6, 4, 4, 3, 5, 4, 4, 3, _
                            3, 12)
                            .TextFileTrailingMinusNumbers = True
                            .Refresh BackgroundQuery:=False
                        End With
            
            Else
                    txtnaoencontrado = MsgBox("O arquivo: " & bookinglist & " não foi encontrado!" & Chr(13) & Chr(13) & "Deseja continuar ?", vbYesNo, "!! ERRO !!")
                    If txtnaoencontrado = vbNo Then
                         Exit Sub
                    End If
            
                
            End If
            
    End If
    
Next inseribkglistrow


    Sheets("Allocation").Select
    Range("A1").Select
    Columns("E:E").Select
    Selection.Delete Shift:=xlToLeft

    Rows("1:100000").Select
    Selection.AutoFilter Field:=1, Criteria1:="-----"
    Selection.Delete Shift:=xlUp



'' DESCOBRINDO POL

    Range("W1").Select
    ActiveCell.FormulaR1C1 = "POL"
    Range("W7").Select
    ActiveCell.FormulaR1C1 = _
        "=LEFT(IF(ISERR(SEARCH(""-"",R[-3]C[-22],1)),""0"",R[-3]C[-22]&R[-3]C[-21]),3)"
    Range("W7").Select
    Selection.Copy
    Range("W8:W100000").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Columns("W:W").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

    Rows("1:100000").Select
    Selection.AutoFilter Field:=23, Criteria1:="0"
    Range("W1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Rows("1:100000").Select
    Selection.AutoFilter


    Range("W7").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C"
    Range("W7").Select
    Selection.Copy
    Range("W8:W100000").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    ActiveSheet.Paste
    
    Range("W1").Select
    ActiveCell.FormulaR1C1 = "POL"
    Range("W1").Select
    Selection.Copy
    Range("W2:W6").Select
    ActiveSheet.Paste
   
    Columns("W:W").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False


'' DESCOBRINDO POD

    Range("Z1").Select
    ActiveCell.FormulaR1C1 = "POD"
    Range("Z7").Select
    ActiveCell.FormulaR1C1 = _
        "=RIGHT(LEFT(IF(ISERR(SEARCH(""-"",R[-3]C[-25],1)),0,R[-3]C[-25]&R[-3]C[-24]),8),3)"
    Range("Z7").Select
    Selection.Copy
    Range("Z8:Z100000").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Columns("Z:Z").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

    Rows("1:100000").Select
    Selection.AutoFilter Field:=26, Criteria1:="0"
    Range("Z1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Rows("1:100000").Select
    Selection.AutoFilter


    Range("Z7").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C"
    Range("Z7").Select
    Selection.Copy
    Range("Z8:Z100000").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    ActiveSheet.Paste
    
    Range("Z1").Select
    ActiveCell.FormulaR1C1 = "POD"
    Range("Z1").Select
    Selection.Copy
    Range("Z2:Z6").Select
    ActiveSheet.Paste
   
    Columns("Z:Z").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    
'' DESCOBRINDO origem POR
 
    Range("AA1").Select
    ActiveCell.FormulaR1C1 = "POR"
    Range("AA7").Select
    ActiveCell.FormulaR1C1 = _
        "=LEFT(IF(ISERR(SEARCH(""00"",RC[-26],1)),""0"",RC[-26]),3)"
    Range("AA7").Select
    Selection.Copy
    Range("AA8:AA100000").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Columns("AA:AA").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    Rows("1:100000").Select
    Selection.AutoFilter Field:=27, Criteria1:="0"
    Range("AA1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Rows("1:100000").Select
    Selection.AutoFilter

    Range("AA7").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C"
    Range("AA7").Select
    Selection.Copy
    Range("AA8:AA100000").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    ActiveSheet.Paste
    
    Range("AA1").Select
    ActiveCell.FormulaR1C1 = "POR"
    Range("AA1").Select
    Selection.Copy
    Range("AA2:AA6").Select
    ActiveSheet.Paste
   
    Columns("AA:AA").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False


'' DESCOBRINDO VESSEL

    Range("Y1").Select
    ActiveCell.FormulaR1C1 = "VESSEL"
    Range("Y7").Select
    ActiveCell.FormulaR1C1 = _
        "=RIGHT(LEFT(IF(ISERR(SEARCH(""Vesse"",R[-6]C[-24],1)),""NAADAA"",R[-6]C[-24]&R[-6]C[-23]&R[-6]C[-22]),18),10) & RIGHT(LEFT(IF(ISERR(SEARCH(""Vesse"",R[-6]C[-24],1)),,R[-6]C[-24]&R[-6]C[-23]&R[-6]C[-22]),21),1)"
    Range("Y7").Select
    Selection.Copy
    Range("Y8:Y100000").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Columns("Y:Y").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

    Rows("1:100000").Select
    Selection.AutoFilter Field:=25, Criteria1:="NAADAA"
    Range("Y1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Rows("1:100000").Select
    Selection.AutoFilter

    Range("Y7").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C"
    Range("Y7").Select
    Selection.Copy
    Range("Y8:Y100000").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    ActiveSheet.Paste
    
    Range("Y1").Select
    ActiveCell.FormulaR1C1 = "VESSEL"
    Range("Y1").Select
    Selection.Copy
    Range("Y2:Y6").Select
    ActiveSheet.Paste
   
    Columns("Y:Y").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

' END DESCOBRINDO VESSEL

    Range("A1").Select
    Rows("1:100000").Select
    Selection.AutoFilter Field:=1, Criteria1:="Danie"
    Selection.Delete Shift:=xlUp
    Selection.AutoFilter Field:=1, Criteria1:=""
    Selection.Delete Shift:=xlUp
    Selection.AutoFilter Field:=7, Criteria1:=""
    Selection.Delete Shift:=xlUp
    Selection.AutoFilter Field:=1, Criteria1:="Rec."
    Selection.Delete Shift:=xlUp

   
    Sheets("Allocation").Select
    Range("A1").Select
    Selection.EntireRow.Insert
    Selection.EntireRow.Insert
    Selection.EntireRow.Insert
    ActiveCell.FormulaR1C1 = "rel."
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Del."
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Booking No"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Shipper"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Commodity"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "EQU"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "B"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "I"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "R"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "TEUs"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "IMO"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Ree"
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "Ov"
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "Fum"
    Range("O1").Select
    ActiveCell.FormulaR1C1 = "IM"
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "APS"
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "FR"
    Range("R1").Select
    ActiveCell.FormulaR1C1 = "SR"
    Range("S1").Select
    ActiveCell.FormulaR1C1 = "Weight [KG]"
    Range("T1").Select
    ActiveCell.FormulaR1C1 = "PESO TONS"
    Range("W1").Select
    ActiveCell.FormulaR1C1 = "POL"
    Range("Y1").Select
    ActiveCell.FormulaR1C1 = "VESSEL"
    Range("Z1").Select
    ActiveCell.FormulaR1C1 = "POD"
    Range("AA1").Select
    ActiveCell.FormulaR1C1 = "POR"

    Columns("H:I").Select
    Selection.Delete Shift:=xlToLeft
    Columns("I:I").Select
    Selection.Delete Shift:=xlToLeft
    Columns("J:O").Select
    Selection.Delete Shift:=xlToLeft
    Columns("K:M").Select
    Selection.Delete Shift:=xlToLeft
    Columns("A:B").Select
    Selection.Delete Shift:=xlToLeft
    Columns("J:J").Select
    Selection.Delete Shift:=xlToLeft


' AJUSTANDO VIAGEM ESPACO NAVIO


    Columns("J:J").Select
    Selection.Replace What:="    ", Replacement:=" 00", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

    Columns("J:J").Select
    Selection.Replace What:="   ", Replacement:=" 0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

    Selection.Replace What:="  ", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

' FIM AJUSTANDO VIAGEM ESPACO NAVIO


Range("F:F").Select
    Selection.Replace What:="RH", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

Range("F:F").Select
    Selection.Replace What:="DC", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

Range("F:F").Select
    Selection.Replace What:="DV", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

Range("F:F").Select
    Selection.Replace What:="IH", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

Range("F:F").Select
    Selection.Replace What:="FR", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

Range("F:F").Select
    Selection.Replace What:="bk", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

Range("F:F").Select
    Selection.Replace What:="FH", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

Range("F:F").Select
    Selection.Replace What:="RF", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

Range("F:F").Select
    Selection.Replace What:="HC", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

Range("F:F").Select
    Selection.Replace What:="OT", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

Range("F:F").Select
    Selection.Replace What:="TK", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

Range("F:F").Select
    Selection.Replace What:="RA", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

Range("F:F").Select
    Selection.Replace What:="OH", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

Range("F:F").Select
    Selection.Replace What:="VT", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

Range("F:F").Select
    Selection.Replace What:="IN", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False


    Cells.Select
    Cells.EntireColumn.AutoFit

    Rows("2:100000").Select
    Selection.AutoFilter Field:=4, Criteria1:=""
    Selection.Delete Shift:=xlUp


' APAGAR BOOKING DE COBERTURA PELO SHIPPER

    For excluishipperrow = 17 To 40
        
        excluishipper = Sheets("MACRO").Range("F" & excluishipperrow)
        
        If excluishipper <> "" Then
            
            Range("A2").Select
            Selection.EntireRow.Insert
            Rows("2:100000").Select
            Selection.AutoFilter Field:=2, Criteria1:=excluishipper
            Selection.EntireRow.Delete
        
        End If
        
    Next excluishipperrow

'--------------------- CHAMANDO MACRO POR TRADE -------------
'--------------------- CHAMANDO MACRO POR TRADE -------------
'--------------------- CHAMANDO MACRO POR TRADE -------------
'--------------------- CHAMANDO MACRO POR TRADE -------------

Windows(planmacro).Activate
If trade = "Asia&Amsul" Then
    Call Módulo1.ALLOCATIONASIA
End If

Windows(planmacro).Activate
If trade = "Euromed" Then
    Call Módulo3.ALLOCATIONEUROMED
End If
    
Windows(planmacro).Activate
If trade = "Americas" Then
    Call Módulo2.ALLOCATIONAMERICAS
End If

Windows(planmacro).Activate
Sheets("MACRO").Select
Range("B7").Select

ThisWorkbook.Save

End Sub
