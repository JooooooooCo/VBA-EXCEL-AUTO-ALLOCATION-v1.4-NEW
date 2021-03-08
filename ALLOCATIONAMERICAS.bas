Attribute VB_Name = "ALLOCATIONAMERICAS"
Sub ALLOCATIONAMERICAS()

'                                                       ALOCANDO
    'by Joao Costa

Windows(planaloca).Activate
Sheets(trade).Select


Dim rowaloca As Integer
Dim rowcompass As Long
Dim vessel As String
Dim pol As String
Dim por As String
Dim kgs As Long
Dim tons As String
Dim teus As Long
Dim plugs As Long
Dim moves As Long
Dim naviobookinglist As String
Dim resultmovesaloc As Boolean

For rowaloca = 4 To 300

    If Range("D" & rowaloca) <> "" Then

            If Range("B" & rowaloca) <> "" Then
                vessel = Left(Range("B" & rowaloca), 10)
            End If
                    
                    
                    pol = Range("D" & rowaloca)
                    por = ""
                    polaloca = ""

                    If pol = "PNG via BUE" Then
                    pol = "BUE"
                    polaloca = "PNG via BUE"
                    End If

                    If pol = "RIG via SSZ" Then
                    pol = "SSZ"
                    por = "RIG"
                    polaloca = "RIG via SSZ"
                    End If

                    If pol = "IBB via SSZ" Then
                    pol = "SSZ"
                    por = "IBB"
                    polaloca = "IBB via SSZ"
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

                                        If por = "" Then

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

                                            If Range("J" & rowcompass) = vessel And Range("I" & rowcompass) = pol And Range("L" & rowcompass) = por Then

                                                naviobookinglist = "ok"
                                                kgs = Range("H" & rowcompass) + kgs
                                                teus = Range("F" & rowcompass) + teus
                                                moves = Range("E" & rowcompass) + moves

                                                If Range("G" & rowcompass) = "Y" Then
                                                        plugs = Range("E" & rowcompass) + plugs
                                                End If

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

                                If polaloca <> "" Then
                                    erro = vessel & " - " & polaloca & Chr(13) & erro
                                Else
                                    erro = vessel & " - " & pol & Chr(13) & erro
                                End If


                        End If

    End If

Next rowaloca

ThisWorkbook.Save

If erro <> "" Then

MsgBox ("!!! ERRO !!! A(s) escala(s) a seguir não foi(foram) atualizada(s), pois não foram encontrados dados no booking list:" & Chr(13) & Chr(13) & erro)

Else

MsgBox ("Alocação Finalizada.")

End If


End Sub






