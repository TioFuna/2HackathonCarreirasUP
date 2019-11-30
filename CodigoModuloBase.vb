Dim aPermutacoes() As String
Dim aPalavrasValidas() As String
Dim aResultadosValidos() As Variant

Sub main()
    Dim tempoIni        As Date
    Dim texto           As String
    Dim iChar           As Integer
    Dim sCel            As String
    Dim palavras()      As String
    Dim resultados()    As String
    Dim textoTemp       As String
    Dim i, iIniPalavras As Double
    Dim iIniTemp        As Double
    Dim iRV, iPV        As Double
    
    tempoIni = Now
    
    texto = Worksheets("Resultados").Range("C2").Value
    'texto = "vermelho"
    
    'verificacoes
    'qtd caracteres 16
    If Len(texto) > 16 Then
        Call MsgBox("Muitos caracteres!" & Chr(13) & "Utilize no máximo 16 letras.", vbOKOnly + vbCritical, "ERRO!")
        Worksheets("Resultados").Range("C2").Select
        Exit Sub
    End If
    'somente caracteres de a-z ou A-Z
    For i = 1 To Len(texto)
        iChar = Asc(Mid(texto, i, 1))
        If iChar < 65 Or iChar > 127 Or (iChar > 90 And iChar < 97) Then
            Call MsgBox("Utilize somente letras de 'a'-'z' ou 'A'-'Z'.", vbOKOnly + vbCritical, "ERRO!")
            Worksheets("Resultados").Range("C2").Select
            Exit Sub
        End If
    Next
    
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
    Else
        If UBound(aResultadosValidos) < 1 Then
            Call MsgBox("Não foi encontrado nenhum anagrama com estas letras.", vbOKOnly + vbInformation, "Resultado")
            Worksheets("Resultados").Range("C2").Select
            Exit Sub
        End If
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
    
    ' tempo decorrido em segundos
    'Call MsgBox("Fim da execução!!!" & Chr(13) & "Tempo decorrido:" & Int(CSng((Now - tempoIni) * 24 * 3600)) & " Segundos", vbOKOnly + vbInformation, "FIM")
    Worksheets("Resultados").Range("C2").Select
    Debug.Print Int(CSng((Now - tempoIni) * 24 * 3600)) & " Segundos"
End Sub

Sub printaResultados()
    Dim txt     As String
    Dim i, j    As Integer
    
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

Function GetPalavra(palavras As Variant, texto As String, iIniPalavras As Double, iPV As Double) As String
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
