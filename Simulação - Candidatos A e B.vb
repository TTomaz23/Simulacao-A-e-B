'Declaração de variáveis globais
Dim escolhido
Dim candidatoA
Dim candidatoB
Dim resposta


'Programa para gerar os votos na quantidade de eleitores e proporção dos candidatos
Sub gerarcandidatos()

'Gera uma caixa de mensagem para alertar que a operação pode demorar e dá a chance de cancelar a operação
resposta = MsgBox("Note que o tempo desta operação depende da quantidade de eleitores escolhidos", 1)

'Pula para o fim do programa caso o usuário deseje cancelar
If (resposta = 2) Then GoTo cancelar

'O valor do loop é iniciado
Range("A1") = 1

'A quantidade de votos para cada candidato é zerada
candidatoA = 0
candidatoB = 0

'Os valores referentes a simulação anterior são apagados
Range(Cells(14, 6), Cells(14 + Range("D6") + Range("D7"), 7)).ClearContents

'Repetição do número de vez dos eleitores
For i = Range("A1") To Range("C3")
  
    'Se o número gerado aleatório for maior que a chance de A e todos os candidatos B ainda n tiverem sido escolhidos
    'será escolhido um candidato B e contabilizado
    If (Rnd > Range("C10") And candidatoB < Range("C11") * Range("C3")) Then
        escolhido = 1
        candidatoB = candidatoB + 1
        GoTo B
    End If
    
    'Se ainda existirem candidatos A para serem escolhidos, então ele será escolhido e contabilizado
    If (candidatoA < Range("C10") * Range("C3")) Then
        escolhido = 0
        candidatoA = candidatoA + 1
        
    'Caso contrário será escolhido um B
    Else
        escolhido = 1
        candidatoB = candidatoB + 1
    End If
    
B:
    
    'A linha referente ao loop receberá o candidato escolhido
    Cells(13 + Range("A1"), 7) = escolhido
    
    'O voto será atualizado
    Range("D6") = candidatoA
    Range("D7") = candidatoB

    'A linha referente ao loop receberá o número do eleitor
    Cells(13 + Range("A1"), 6) = i
    
    'Número do loop incrementado
    Range("A1") = Range("A1") + 1
    
    'O loop é reiniciado
    Next i
    
cancelar:

End Sub


'Programa para gerar duas amostras com valores e em quantidades desejadas
Sub contaramostra()
    
'Gera uma caixa de mensagem para alertar que a operação pode demorar e dá a chance de cancelar a operação
resposta = MsgBox("Note que o tempo desta operação depende da quantidade de amostras escolhidas", 1)

'Pula para o fim do programa caso o usuário deseje cancelar
	If (resposta = 2) Then GoTo fim
    
    
    'O valor do loop é iniciado
    Range("A2") = 1
    
    'Os valores da simulação anteriores são apagados
    Range(Cells(14, 10), Cells(10000, 16)).ClearContents
    
    'O loop da quantidade de amostras 1 é iniciado
    For j = Range("A2") To Range("I3")
    
        'O valor do loop é iniciado
        Range("A1") = 1
        
        'O loop para quantidade de valores em cada amostra é iniciado
        For i = Range("A1") To Range("L12")
        
            'A linha referente ao loop do candidato B recebe seu próprio valor somado
            'ao valor de uma célula aleatória com um voto
            Cells(13 + Range("A2"), 11) = Cells(13 + Range("A2"), 11) + Cells(14 + Rnd * Range("C3"), 7)
            
            'O valor do loop é incrementado
            Range("A1") = Range("A1") + 1
        
        'O loop é reiniciado
        Next i
        
        'O valor de A será a quantidade de votos na amostra menos os votos em B
        Cells(13 + Range("A2"), 10) = Range("L12") - Cells(13 + Range("A2"), 11)
        
        'A proporção de A é calculada
        Cells(13 + Range("A2"), 12) = (Cells(13 + Range("A2"), 10) * 100) / Range("L12")
        
        'O valor do loop é reiniciado
        Range("A1") = 1

        'O loop da quantidade de amostras 2 é iniciado
        For i = Range("A1") To Range("P12")
        
            'A linha referente ao loop do candidato B recebe seu próprio valor somado
            'ao valor de uma célula aleatória com um voto
            Cells(13 + Range("A2"), 15) = Cells(13 + Range("A2"), 15) + Cells(14 + Rnd * Range("C3"), 7)
            
            'O valor do loop é incrementado
            Range("A1") = Range("A1") + 1
            
        'O loop é reiniciado
        Next i
        
        'O valor de A será a quantidade de votos na amostra menos os votos em B
        Cells(13 + Range("A2"), 14) = Range("P12") - Cells(13 + Range("A2"), 15)
        
        'A proporção de A é calculada
        Cells(13 + Range("A2"), 16) = (Cells(13 + Range("A2"), 14) * 100) / Range("P12")
    
    'O valor do loop é incrementado
    Range("A2") = Range("A2") + 1
    
    'O loop é reiniciado
    Next j
    
    'Chama a função que atualiza os valores máximos dos gráficos
    Call grafico
    
fim:
    
End Sub

'Função para atualizar os valores máximos dos gráficos
Sub grafico()
    
    'Atualiza o valor máximo do gráfico 1 da planilha 1
    Worksheets(1).ChartObjects(1).Activate
        ActiveChart.Axes(xlCategory).Select
            ActiveChart.Axes(xlCategory).MinimumScale = 0
            ActiveChart.Axes(xlCategory).MaximumScale = (Range("I3") * 1.1)
            
    'Atualiza o valor máximo do gráfico 2 da planilha 1
    Worksheets(1).ChartObjects(2).Activate
        ActiveChart.Axes(xlCategory).Select
            ActiveChart.Axes(xlCategory).MinimumScale = 0
            ActiveChart.Axes(xlCategory).MaximumScale = (Range("I3") * 1.1)
            
    'Atualiza o valor máximo do gráfico 1 da planilha 2
    Worksheets(2).ChartObjects(1).Activate
        ActiveChart.Axes(xlCategory).Select
            ActiveChart.Axes(xlCategory).MinimumScale = 0
            ActiveChart.Axes(xlCategory).MaximumScale = Worksheets(1).Range("I3") * 1.1
            
    'Atualiza o valor máximo do gráfico 2 da planilha 2
    Worksheets(2).ChartObjects(2).Activate
        ActiveChart.Axes(xlCategory).Select
            ActiveChart.Axes(xlCategory).MinimumScale = 0
            ActiveChart.Axes(xlCategory).MaximumScale = Worksheets(1).Range("I3") * 1.1
            
End Sub
