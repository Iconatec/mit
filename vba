Option Explicit

Public Sub GerarMIT_JSON()

    Dim wsCab As Worksheet, wsDeb As Worksheet
    Dim cnpjContribuinte As String, cnpjRaiz As String
    Dim mesApuracao As Long, anoApuracao As Long
    Dim caminhoPasta As String, nomeArquivo As String, caminhoArquivo As String
    Dim json As String
    Dim semMov As Boolean
    
    Set wsCab = ThisWorkbook.Worksheets("MIT_Cabecalho")
    Set wsDeb = ThisWorkbook.Worksheets("MIT_Debitos")
    
    mesApuracao = CLng(wsCab.Range("A2").Value)
    anoApuracao = CLng(wsCab.Range("B2").Value)
    semMov = ToBoolean(wsCab.Range("C2").Value)
    
    cnpjContribuinte = ApenasDigitos(CStr(wsCab.Range("O2").Value))
    If Len(cnpjContribuinte) <> 14 Then
        MsgBox "Preencha o CnpjContribuinte com 14 dígitos na aba MIT_Cabecalho.", vbExclamation
        Exit Sub
    End If
    
    cnpjRaiz = Left(cnpjContribuinte, 8)
    
    caminhoPasta = "C:\MIT\"
If Dir(caminhoPasta, vbDirectory) = "" Then MkDir caminhoPasta
    
    nomeArquivo = cnpjRaiz & "-MIT-" & Format(DateSerial(anoApuracao, mesApuracao, 1), "yyyymm") & ".json"
    caminhoArquivo = caminhoPasta & nomeArquivo
    
    If Trim(CStr(wsCab.Range("F2").Value)) = "" And Not semMov Then
        MsgBox "Preencha VariacoesMonetarias na aba MIT_Cabecalho.", vbExclamation
        Exit Sub
    End If
    
    json = "{"
    json = json & vbCrLf & MontarPeriodoApuracao(wsCab)
    json = json & "," & vbCrLf & MontarDadosIniciais(wsCab)
    
    If Not semMov Then
        json = json & "," & vbCrLf & MontarDebitos(wsCab, wsDeb)
    End If
    
    json = json & vbCrLf & "}"
    
    SalvarTextoUTF8 caminhoArquivo, json
    
    MsgBox "Arquivo gerado com sucesso:" & vbCrLf & caminhoArquivo, vbInformation

End Sub

Private Function MontarPeriodoApuracao(ws As Worksheet) As String
    Dim s As String
    
    s = Space(4) & """PeriodoApuracao"": {" & vbCrLf
    s = s & Space(8) & """MesApuracao"": " & CLng(ws.Range("A2").Value) & "," & vbCrLf
    s = s & Space(8) & """AnoApuracao"": " & CLng(ws.Range("B2").Value) & vbCrLf
    s = s & Space(4) & "}"
    
    MontarPeriodoApuracao = s
End Function

Private Function MontarDadosIniciais(ws As Worksheet) As String
    Dim s As String
    Dim semMov As Boolean
    Dim qualif As String, tribLucro As String, varMon As String, regPisCof As String
    Dim cpf As String, ddd As String, tel As String, email As String, uf As String, reg As String
    
    semMov = ToBoolean(ws.Range("C2").Value)
    qualif = Trim(CStr(ws.Range("D2").Value))
    tribLucro = Trim(CStr(ws.Range("E2").Value))
    varMon = Trim(CStr(ws.Range("F2").Value))
    regPisCof = Trim(CStr(ws.Range("G2").Value))
    cpf = ApenasDigitos(CStr(ws.Range("H2").Value))
    ddd = ApenasDigitos(CStr(ws.Range("I2").Value))
    tel = ApenasDigitos(CStr(ws.Range("J2").Value))
    email = Trim(CStr(ws.Range("K2").Value))
    uf = UCase(Trim(CStr(ws.Range("L2").Value)))
    reg = UCase(Trim(CStr(ws.Range("M2").Value)))
    
    s = Space(4) & """DadosIniciais"": {" & vbCrLf
    s = s & Space(8) & """SemMovimento"": " & LCase(CStr(semMov)) & "," & vbCrLf
    s = s & Space(8) & """QualificacaoPj"": " & qualif
    
    If Not semMov And qualif <> "11" And tribLucro <> "" Then
        s = s & "," & vbCrLf & Space(8) & """TributacaoLucro"": " & tribLucro
    End If
    
    If Not semMov And varMon <> "" Then
        s = s & "," & vbCrLf & Space(8) & """VariacoesMonetarias"": " & varMon
    End If
    
    If Not semMov And regPisCof <> "" Then
        s = s & "," & vbCrLf & Space(8) & """RegimePisCofins"": " & regPisCof
    End If
    
    s = s & "," & vbCrLf & Space(8) & """ResponsavelApuracao"": {" & vbCrLf
    s = s & Space(12) & """CpfResponsavel"": """ & cpf & """"
    
    If ddd <> "" And tel <> "" Then
        s = s & "," & vbCrLf & Space(12) & """TelResponsavel"": {" & vbCrLf
        s = s & Space(16) & """Ddd"": """ & ddd & """," & vbCrLf
        s = s & Space(16) & """NumTelefone"": """ & tel & """" & vbCrLf
        s = s & Space(12) & "}"
    End If
    
    If email <> "" Then
        s = s & "," & vbCrLf & Space(12) & """EmailResponsavel"": """ & EscapeJson(email) & """"
    End If
    
    If uf <> "" And reg <> "" Then
        s = s & "," & vbCrLf & Space(12) & """RegistroCrc"": {" & vbCrLf
        s = s & Space(16) & """UfRegistro"": """ & uf & """," & vbCrLf
        s = s & Space(16) & """NumRegistro"": """ & EscapeJson(reg) & """" & vbCrLf
        s = s & Space(12) & "}"
    End If
    
    s = s & vbCrLf & Space(8) & "}" & vbCrLf
    s = s & Space(4) & "}"
    
    MontarDadosIniciais = s
End Function

Private Function MontarDebitos(wsCab As Worksheet, wsDeb As Worksheet) As String
    Dim s As String
    Dim grupos As Variant
    Dim grupo As Variant
    Dim bloco As String
    Dim tributacaoLucro As String
    Dim balanco As Boolean
    Dim temAlgum As Boolean
    
    grupos = Array( _
        "Irpj", _
        "Csll", _
        "Irrf", _
        "Ipi", _
        "Iof", _
        "PisPasep", _
        "Cofins", _
        "ContribuicoesDiversas", _
        "Cpss", _
        "RetPagamentoUnificado" _
    )
    
    tributacaoLucro = Trim(CStr(wsCab.Range("E2").Value))
    balanco = ToBoolean(wsCab.Range("N2").Value)
    
    s = Space(4) & """Debitos"": {" & vbCrLf
    temAlgum = False
    
    If tributacaoLucro = "1" Then
        s = s & Space(8) & """BalancoLucroReal"": " & LCase(CStr(balanco))
        temAlgum = True
    End If
    
    For Each grupo In grupos
        bloco = MontarGrupoDebito(wsDeb, CStr(grupo))
        
        If bloco <> "" Then
            If temAlgum Then s = s & "," & vbCrLf
            s = s & bloco
            temAlgum = True
        End If
    Next grupo
    
    s = s & vbCrLf & Space(4) & "}"
    
    MontarDebitos = s
End Function

Private Function MontarGrupoDebito(wsDeb As Worksheet, nomeGrupo As String) As String
    Dim ult As Long, i As Long
    Dim listaNormal As String, listaPosEvento As String
    Dim tipoLista As String, grupo As String
    Dim temNormal As Boolean, temPosEvento As Boolean
    Dim item As String
    Dim s As String
    
    ult = wsDeb.Cells(wsDeb.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To ult
        grupo = Trim(CStr(wsDeb.Cells(i, 1).Value))
        tipoLista = Trim(CStr(wsDeb.Cells(i, 2).Value))
        
        If StrComp(grupo, nomeGrupo, vbTextCompare) = 0 Then
            item = MontarItemDebito(wsDeb, i, (UCase(tipoLista) = "LISTADEBITOSAPOSEVENTO"))
            
            If UCase(tipoLista) = "LISTADEBITOSAPOSEVENTO" Then
                If temPosEvento Then listaPosEvento = listaPosEvento & "," & vbCrLf
                listaPosEvento = listaPosEvento & item
                temPosEvento = True
            Else
                If temNormal Then listaNormal = listaNormal & "," & vbCrLf
                listaNormal = listaNormal & item
                temNormal = True
            End If
        End If
    Next i
    
    If Not temNormal And Not temPosEvento Then Exit Function
    
    s = Space(8) & """" & nomeGrupo & """: {" & vbCrLf
    
    If temNormal Then
        s = s & Space(12) & """ListaDebitos"": [" & vbCrLf
        s = s & listaNormal & vbCrLf
        s = s & Space(12) & "]"
    End If
    
    If temPosEvento Then
        If temNormal Then s = s & "," & vbCrLf
        s = s & Space(12) & """ListaDebitosAposEvento"": [" & vbCrLf
        s = s & listaPosEvento & vbCrLf
        s = s & Space(12) & "]"
    End If
    
    s = s & vbCrLf & Space(8) & "}"
    
    MontarGrupoDebito = s
End Function

Private Function MontarItemDebito(wsDeb As Worksheet, linha As Long, aposEvento As Boolean) As String
    Dim s As String
    Dim grupo As String
    Dim idDebito As String, idEvento As String, codigo As String
    Dim pa As String, anoPost As String, trimPost As String, anoDeb As String
    Dim cnpjEst As String, cnpjInc As String, cnpjScp As String, codMun As String
    Dim valor As Double
    
    grupo = Trim(CStr(wsDeb.Cells(linha, 1).Value))
    idDebito = Trim(CStr(wsDeb.Cells(linha, 3).Value))
    idEvento = Trim(CStr(wsDeb.Cells(linha, 4).Value))
    codigo = ApenasDigitos(CStr(wsDeb.Cells(linha, 5).Value))
    pa = Trim(CStr(wsDeb.Cells(linha, 6).Value))
    anoPost = Trim(CStr(wsDeb.Cells(linha, 7).Value))
    trimPost = Trim(CStr(wsDeb.Cells(linha, 8).Value))
    anoDeb = Trim(CStr(wsDeb.Cells(linha, 9).Value))
    cnpjEst = ApenasDigitos(CStr(wsDeb.Cells(linha, 10).Value))
    cnpjInc = ApenasDigitos(CStr(wsDeb.Cells(linha, 11).Value))
    cnpjScp = ApenasDigitos(CStr(wsDeb.Cells(linha, 12).Value))
    codMun = ApenasDigitos(CStr(wsDeb.Cells(linha, 13).Value))
    valor = CDbl(wsDeb.Cells(linha, 14).Value)
    
    s = Space(16) & "{" & vbCrLf
    s = s & Space(20) & """IdDebito"": " & idDebito
    
    If aposEvento And idEvento <> "" Then
        s = s & "," & vbCrLf & Space(20) & """IdEventoDebito"": " & idEvento
    End If
    
    s = s & "," & vbCrLf & Space(20) & """CodigoDebito"": """ & codigo & """"
    
    If UCase(grupo) <> "RETPAGAMENTOUNIFICADO" Then
        If pa <> "" Then s = s & "," & vbCrLf & Space(20) & """PaDebito"": " & pa
    End If
    
    If anoPost <> "" Then s = s & "," & vbCrLf & Space(20) & """AnoPostergado"": " & anoPost
    If trimPost <> "" Then s = s & "," & vbCrLf & Space(20) & """TrimPostergado"": " & trimPost
    If anoDeb <> "" Then s = s & "," & vbCrLf & Space(20) & """AnoDebito"": " & anoDeb
    If cnpjEst <> "" Then s = s & "," & vbCrLf & Space(20) & """CnpjEstabelecimento"": """ & cnpjEst & """"
    If cnpjInc <> "" Then s = s & "," & vbCrLf & Space(20) & """CnpjIncorporacao"": """ & cnpjInc & """"
    If cnpjScp <> "" Then s = s & "," & vbCrLf & Space(20) & """CnpjScp"": """ & cnpjScp & """"
    If codMun <> "" Then s = s & "," & vbCrLf & Space(20) & """CodigoMunicipioOuro"": """ & codMun & """"
    
    s = s & "," & vbCrLf & Space(20) & """ValorDebito"": " & Replace(Format(valor, "0.00"), ",", ".") & vbCrLf
    s = s & Space(16) & "}"
    
    MontarItemDebito = s
End Function

Private Function ToBoolean(v As Variant) As Boolean
    Dim t As String
    t = UCase(Trim(CStr(v)))
    ToBoolean = (t = "TRUE" Or t = "VERDADEIRO" Or t = "1" Or t = "SIM")
End Function

Private Function ApenasDigitos(texto As String) As String
    Dim i As Long, c As String, r As String
    For i = 1 To Len(texto)
        c = Mid$(texto, i, 1)
        If c >= "0" And c <= "9" Then r = r & c
    Next i
    ApenasDigitos = r
End Function

Private Function EscapeJson(ByVal texto As String) As String
    texto = Replace(texto, "\", "\\")
    texto = Replace(texto, """", "\""")
    texto = Replace(texto, vbCrLf, "\n")
    texto = Replace(texto, vbCr, "\n")
    texto = Replace(texto, vbLf, "\n")
    EscapeJson = texto
End Function

Private Sub SalvarTextoUTF8(ByVal caminhoArquivo As String, ByVal conteudo As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    
    With stream
        .Charset = "utf-8"
        .Open
        .WriteText conteudo
        .SaveToFile caminhoArquivo, 2
        .Close
    End With
    
    Set stream = Nothing
End Sub
