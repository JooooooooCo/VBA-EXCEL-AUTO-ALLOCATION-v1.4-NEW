Attribute VB_Name = "ALLOCATIONEUROMED"
Sub ALLOCATIONEUROMED()

'                                                       ALOCANDO

Windows(planaloca).Activate
Sheets(trade).Select


Dim rowaloca As Integer
Dim rowcompass As Long
Dim vessel As String
Dim pol As String
Dim kgs As Long
Dim tons As String
Dim teus As Long
Dim plugs As Long
Dim moves As Long
Dim unit20pes As Long
Dim fator40hc As Long
Dim fator40rh As Long
Dim fatorothers As Long
Dim naviobookinglist As String
Dim resultmovesaloc As Boolean


For rowaloca = 4 To 300

    If Range("D" & rowaloca) <> "" Then
    
            If Range("A" & rowaloca) <> "" Then
                service = Range("A" & rowaloca)
            End If

            If Range("B" & rowaloca) <> "" Then
                vessel = Left(Range("B" & rowaloca), 10)
            End If
                    
                If service = "MESA" Or service = "SAEC 3" Then
                
'ALOCANDO MESA
                    
                    If service = "MESA" Then
                    
                            pol = Range("D" & rowaloca)
        
                            kgs = 0
                            plugs = 0
                            unit20pes = 0
                            fator40hc = 0
                            fator40rh = 0
                            fatorothers = 0
                            naviobookinglist = ""
        
                            Windows(planmacro).Activate
                            Sheets("Allocation").Select
        
                                For rowcompass = 2 To 100000
        
                                        If Range("D" & rowcompass) <> "" Then
        
                                            If Range("J" & rowcompass) = vessel And Range("I" & rowcompass) = pol Then
        
                                                naviobookinglist = "ok"
                                                kgs = Range("H" & rowcompass) + kgs
        
                                                If Range("G" & rowcompass) = "Y" Then
                                                        plugs = Range("E" & rowcompass) + plugs
                                                End If
        
                                                If Left(Range("D" & rowcompass), 2) = "20" Then
                                                        unit20pes = Range("E" & rowcompass) + unit20pes
                                                End If
        
                                                If Range("D" & rowcompass) = "40HC" Or Range("D" & rowcompass) = "40OH" Then
                                                        fator40hc = Range("E" & rowcompass) + fator40hc
                                                End If
        
                                                If Range("D" & rowcompass) = "40RH" Then
                                                        fator40rh = Range("E" & rowcompass) + fator40rh
                                                End If
        
                                                If Range("D" & rowcompass) = "40HC" Or Range("D" & rowcompass) = "40OH" Or Range("D" & rowcompass) = "40RH" Then
                                                        ' N�O FAZ NADA PQ J� CONTABILIZOU
                                                Else
                                                        fatorothers = Range("E" & rowcompass) + fatorothers
                                                End If
        
        
                                            End If
        
                                        Else
        
                                            rowcompass = 100000
        
                                        End If
        
                                Next rowcompass
        
                            Windows(planaloca).Activate
                            Sheets(trade).Select
        
                                If naviobookinglist = "ok" Then
        
                                        tons = kgs
                                        If tons <> 0 Then
                                        tons = Left(tons, (Len(tons) - 3))
                                        End If
        
                                        Range("F" & rowaloca).Value = tons
                                        Range("L" & rowaloca).Value = plugs
                                        Range("N" & rowaloca).Value = unit20pes
                                       
                                        Range("I" & rowaloca).Value = (fator40hc + fator40rh) * 2.25 + fatorothers
        
                                Else
        
                                        erro = vessel & " - " & pol & Chr(13) & erro
        
        
                                End If
                        
                
                    End If
                
'ALOCANDO SAEC 3
                
                
                            pol = Range("D" & rowaloca)
        
        
                            kgs = 0
                            teus = 0
                            plugs = 0
                            unit20pes = 0
                            naviobookinglist = ""
        
                            Windows(planmacro).Activate
                            Sheets("Allocation").Select
        
                                For rowcompass = 2 To 100000
        
                                        If Range("D" & rowcompass) <> "" Then
        
                                            If Range("J" & rowcompass) = vessel And Range("I" & rowcompass) = pol Then
        
                                                naviobookinglist = "ok"
                                                kgs = Range("H" & rowcompass) + kgs
                                                teus = Range("F" & rowcompass) + teus
                                                moves = Range("E" & rowcompass) + moves
        
                                                If Range("G" & rowcompass) = "Y" Then
                                                        plugs = Range("E" & rowcompass) + plugs
                                                End If
                                                
                                                If Left(Range("D" & rowcompass), 2) = "20" Then
                                                        unit20pes = Range("E" & rowcompass) + unit20pes
                                                End If
        
        
                                            End If
        
                                        Else
        
                                            rowcompass = 100000
        
                                        End If
        
                                Next rowcompass
        
                            Windows(planaloca).Activate
                            Sheets(trade).Select
        
                                If naviobookinglist = "ok" Then
        
                                        tons = kgs
                                        If tons <> 0 Then
                                        tons = Left(tons, (Len(tons) - 3))
                                        End If
        
                                        Range("F" & rowaloca).Value = tons
                                        Range("I" & rowaloca).Value = teus
                                        Range("L" & rowaloca).Value = plugs
                                        Range("N" & rowaloca).Value = unit20pes
        
                                Else
        
                                        erro = vessel & " - " & pol & Chr(13) & erro
        
        
                                End If
                        
                        
                
                Else
                
'ALOCANDO SAEC 1 e 2
                
                
                            pol = Range("D" & rowaloca)
        
                            If pol = "PNG - NB" Then
                            pol = "PNG"
                            End If
        
                            If pol = "PNG - SB" Then
                            pol = "PNG"
                            vessel = Left(vessel, (Len(vessel) - 1)) & "W"
                            End If
        
        
                            kgs = 0
                            teus = 0
                            plugs = 0
                            moves = 0
                            naviobookinglist = ""
        
                            Windows(planmacro).Activate
                            Sheets("Allocation").Select
        
                                For rowcompass = 2 To 100000
        
                                        If Range("D" & rowcompass) <> "" Then
        
                                            If Range("J" & rowcompass) = vessel And Range("I" & rowcompass) = pol Then
        
                                                naviobookinglist = "ok"
                                                kgs = Range("H" & rowcompass) + kgs
                                                teus = Range("F" & rowcompass) + teus
                                                moves = Range("E" & rowcompass) + moves
        
                                                If Range("G" & rowcompass) = "Y" Then
                                                        plugs = Range("E" & rowcompass) + plugs
                                                End If
        
                                            End If
        
                                        Else
        
                                            rowcompass = 100000
        
                                        End If
        
                                Next rowcompass
        
                            Windows(planaloca).Activate
                            Sheets(trade).Select
        
                                If naviobookinglist = "ok" Then
        
                                        tons = kgs
                                        If tons <> 0 Then
                                        tons = Left(tons, (Len(tons) - 3))
                                        End If
        
                                        Range("F" & rowaloca).Value = tons
                                        Range("I" & rowaloca).Value = teus
                                        Range("L" & rowaloca).Value = plugs
        
                                        'inserindo moves
        
                                        resultmovesaloc = Range("M" & rowaloca) Like "*mbar*"
        
                                        If resultmovesaloc = False And Range("M" & rowaloca) <> "" Then
        
                                            Range("N" & rowaloca).Value = moves
        
                                        End If
        
                                Else
        
                                        erro = vessel & " - " & pol & Chr(13) & erro
        
        
                                End If
        
        
                            vessel = Left(vessel, (Len(vessel) - 1)) & "N"
                
                End If

    End If

Next rowaloca

ThisWorkbook.Save

If erro <> "" Then

MsgBox ("!!! ERRO !!! A(s) escala(s) a seguir n�o foi(foram) atualizada(s), pois n�o foram encontrados dados no booking list:" & Chr(13) & Chr(13) & erro)

Else

MsgBox ("Aloca��o Finalizada.")

End If

End Sub




