Dim aPermutacoes() As String
Dim aPalavrasValidas() As String
Dim aResultadosValidos() As Variant

Sub main()
    Dim tempoIni        As Date
    Dim texto           As String
    Dim sCel            As String
    Dim palavras()      As String
    Dim resultados()    As String
    Dim textoTemp       As String
    Dim i, iIniPalavras As Long
    Dim iIniTemp        As Long
    Dim j               As Integer
    Dim iRV, iPV        As Long
    
    tempoIni = Now
    
    texto = Worksheets("Resultados").Range("C2").Value
    'texto = "vermelho"
    
    'verificacoes
    
    
    'trata texto
    texto = Replace(texto, " ", "")
    texto = UCase(texto)
    
    Worksheets("Resultados").Range("C5:C250").Clear
    
    'lista de palavras
    i = 0
    Do
        i = i + 1
        sCel = Worksheets("Palavras").Cells(i, 1).Value
        
        If sCel = "" Then
            Exit Do
        Else
            ReDim Preserve palavras(i)
            palavras(i) = sCel
        End If
    Loop
    
    iRV = 0
    ReDim aResultadosValidos(iRV)
    
    ' procura por palavra com o texto
    textoTemp = texto
    Debug.Print texto
    iIniPalavras = 1
inicio:
    iPV = 0
    ReDim aPalavrasValidas(iPV)
    resultado = GetPalavra(palavras, textoTemp, iIniPalavras, iPV)
    If resultado <> "" Then
        aPalavrasValidas(iPV) = resultado
    End If
    
    iRV = iRV + 1
    ReDim Preserve aResultadosValidos(iRV)
    aResultadosValidos(iRV) = aPalavrasValidas
    
    ' busca por outros resultados validos
    For i = 0 To UBound(aPalavrasValidas)
        iIniTemp = 1
proximo:
        resultado = GetPalavra(palavras, aPalavrasValidas(i), iIniTemp, iPV)
        If resultado <> "" Then
            If resultado = aPalavrasValidas(i) Then
                iIniTemp = iIniTemp + 1
                GoTo proximo
            Else
                aPalavrasValidas(i) = resultado
                iRV = iRV + 1
                ReDim Preserve aResultadosValidos(iRV)
                aResultadosValidos(iRV) = aPalavrasValidas
            End If
        End If
    Next
    
    If iIniPalavras < UBound(palavras) Then
        iIniPalavras = iIniPalavras + 1
        GoTo inicio
    End If
    
    Call printaResultados
    
    'permutacoes
    'Call GetPermutation("", texto, 0)
    
    'For i = 0 To UBound(aPermutacoes)
        'Debug.Print aPermutacoes(i)
    'Next
    
    ' tempo decorrido em segundos
    Debug.Print Int(CSng((Now - tempoIni) * 24 * 3600)) & " Segundos"
End Sub

Sub printaResultados()
    Dim txt As String
    Dim i   As Integer
    For i = 1 To UBound(aResultadosValidos)
        For j = 0 To UBound(aResultadosValidos(i))
            txt = txt & aResultadosValidos(i)(j) & " "
        Next
        Worksheets("Resultados").Cells(i + 4, 3).Value = txt
        txt = ""
        'txt = txt & Chr(13)
    Next
    'Debug.Print txt
End Sub

Function GetPalavra(palavras As Variant, texto As String, iIniPalavras As Long, iPV As Long) As String
    Dim textoTemp   As String
    Dim letra       As String
    Dim posLetra    As String
    Dim resultado   As String
    Dim i           As Long
    Dim j           As Integer
    
    GetPalavra = ""
    
    'compara palavras com texto
    textoTemp = texto
    For i = iIniPalavras To UBound(palavras)
        ' loop letras palavra
        For j = 1 To Len(palavras(i))
            letra = Mid(palavras(i), j, 1)
            'verif se tem a letra no texto temporario
            posLetra = InStr(1, textoTemp, letra)
            If posLetra > 0 Then
                ' se achou a letra, retira do texto temporario
                textoTemp = Mid(textoTemp, 1, posLetra - 1) & Mid(textoTemp, posLetra + 1, 20)
            Else
                ' se nao achou, invalida palavra
                textoTemp = texto   ' retorna texto temporario
                Exit For
            End If
        Next
        ' verif se a palavra foi validada
        If textoTemp = texto Then
            'palavra invalida, ir para a proxima palavra
        Else
            ' verif se utilizada todas as letras do texto
            If textoTemp = "" Then
                GetPalavra = palavras(i)
                If iPV > UBound(aPalavrasValidas) Then
                    ReDim Preserve aPalavrasValidas(iPV)
                End If
                Exit For
            Else
                ' procura por palavra com o texto restante
                'Debug.Print palavras(i) & " - " & textoTemp
                resultado = GetPalavra(palavras, textoTemp, 1, iPV + 1)
                ' verif se retornou palavra valida
                If resultado <> "" Then
                    aPalavrasValidas(iPV + 1) = resultado
                    GetPalavra = palavras(i)
                    Exit For
                End If
            End If
        End If
    Next
    iIniPalavras = i
    'Debug.Print "Fim LOOP - " & texto
End Function

Sub GetPermutation(sElemento1 As String, sElemento2 As String, lArray As Long)
    Dim iQtdCaract  As Integer
    Dim sIsola      As String
    Dim sDemaisInf  As String
    Dim sDemaisSup  As String
    Dim sDemais     As String
    Dim i           As Integer
    
    ReDim Preserve aPermutacoes(lArray)
    
    iQtdCaract = Len(sElemento2)
    
    If iQtdCaract < 2 Then                                 ' verif se ultimo item unico
        aPermutacoes(lArray) = sElemento1 & sElemento2     ' - adiciona item
        lArray = lArray + 1
    Else
        For i = 1 To iQtdCaract                            ' loop entre demais elementos
            sIsola = sElemento1 + Mid(sElemento2, i, 1)    ' - isola parte dos elementos
            sDemaisInf = Left(sElemento2, i - 1)           ' - abaixo elemento isolado
            sDemaisSup = Right(sElemento2, iQtdCaract - i) ' - acima elemento isolado
            sDemais = sDemaisInf + sDemaisSup              ' - monta demais elementos
            Call GetPermutation(sIsola, sDemais, lArray)   ' - permuta demais itens
        Next
    End If
End Sub
