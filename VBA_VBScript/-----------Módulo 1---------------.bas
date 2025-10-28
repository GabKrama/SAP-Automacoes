-----------Módulo 1---------------
Option Explicit

Public Appl, SapGuiAuto, Connection, session, WScript

Function SomaOuTexto(rng As Range) As Variant
    Dim cell As Range
    Dim total As Double
    Dim textoUnico As String
    Dim temTexto As Boolean
    Dim temNumero As Boolean
    
    For Each cell In rng
        If IsNumeric(cell.Value) Then
            total = total + cell.Value
            temNumero = True
        ElseIf cell.Value <> "" Then
            If Not temTexto Then
                textoUnico = cell.Value
                temTexto = True
            ElseIf textoUnico <> cell.Value Then
                ' Se houver mais de um texto diferente, retorna o primeiro encontrado
                SomaOuTexto = textoUnico
                Exit Function
            End If
        End If
    Next cell
    
    If temNumero And Not temTexto Then
        SomaOuTexto = total
    ElseIf temTexto And Not temNumero Then
        SomaOuTexto = textoUnico
    ElseIf Not temTexto And Not temNumero Then
        SomaOuTexto = 0
    Else
        ' Se houver mistura de número e texto, retorna o texto
        SomaOuTexto = textoUnico
    End If
End Function

Function SanitizarNomeArquivo(nome As String) As String
    Dim caracteresInvalidos As Variant
    Dim i As Integer
    caracteresInvalidos = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    For i = LBound(caracteresInvalidos) To UBound(caracteresInvalidos)
        nome = Replace(nome, caracteresInvalidos(i), "_")
    Next i

    SanitizarNomeArquivo = nome
End Function

Function BuscarAliquotas(ncm As String, wsAliquota As Worksheet) As Variant
    Dim linhaAliquota As Range
    Set linhaAliquota = wsAliquota.Columns(1).Find(What:=Trim(ncm), LookIn:=xlValues, LookAt:=xlWhole)
    If Not linhaAliquota Is Nothing Then
        BuscarAliquotas = Array(linhaAliquota.Offset(0, 1).Value, linhaAliquota.Offset(0, 2).Value, _
                                linhaAliquota.Offset(0, 3).Value, linhaAliquota.Offset(0, 4).Value)
    Else
        BuscarAliquotas = Array("NCM não encontrado", "", "", "")
    End If
End Function

Function VerificarAtoConcessorio(material As String, ncm As String, wsAto As Worksheet) As Variant
    Dim linha As Range
    Dim melhorLinha As Range
    Dim maiorSaldo As Double
    Dim saldoAtual As Double
    Dim valorH As Variant, valorK As Variant, valorAtoConcessorio As Variant

    valorH = "Não tem"
    valorK = "Não tem"
    valorAtoConcessorio = "Não tem"
    maiorSaldo = -1

    For Each linha In wsAto.Range("C20:C719")
        If linha.Value = material And linha.Offset(0, 1).Value = ncm Then
            If IsNumeric(linha.Offset(0, 5).Value) Then
                saldoAtual = linha.Offset(0, 5).Value
                If saldoAtual > maiorSaldo Then
                    maiorSaldo = saldoAtual
                    Set melhorLinha = linha
                End If
            End If
        End If
    Next linha

    If Not melhorLinha Is Nothing Then
        valorH = melhorLinha.Offset(0, 5).Value
        valorK = melhorLinha.Offset(0, 8).Value
        valorAtoConcessorio = melhorLinha.Offset(0, -2).Value
    End If

    VerificarAtoConcessorio = Array(IIf(valorH <> "Não tem" Or valorK <> "Não tem", "TEM", "NÃO TEM"), valorH, valorK, valorAtoConcessorio)
End Function

Function BuscarCodigoPlanta(cnpj As String, wsCNPJ As Worksheet) As String
    Dim linhaCNPJ As Range
    Set linhaCNPJ = wsCNPJ.Columns(1).Find(What:=cnpj, LookIn:=xlValues, LookAt:=xlWhole)
    If Not linhaCNPJ Is Nothing Then
        BuscarCodigoPlanta = linhaCNPJ.Offset(0, 1).Value
    Else
        BuscarCodigoPlanta = "CNPJ não encontrado"
    End If
End Function

Function BuscarTaxaDolar(ws As Worksheet) As Double
    If ws.Range("A9").Value Like "*TAXA*" Then
        BuscarTaxaDolar = ws.Range("B9").Value
    ElseIf ws.Range("A10").Value Like "*TAXA*" Then
        BuscarTaxaDolar = ws.Range("B10").Value
    Else
        BuscarTaxaDolar = 0
    End If
End Function

Function BuscarDadosSAPPorPedido(pedido As String) As Object
    Dim dadosSAP As Object
    Set dadosSAP = CreateObject("Scripting.Dictionary")

    ' Inicializa SAP GUI
    Set SapGuiAuto = GetObject("SAPGUI")
    Set Appl = SapGuiAuto.GetScriptingEngine
    Set Connection = Appl.Children(0)
    Set session = Connection.Children(0)

    ' Acessa transação /NME2L
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/NME2L"
    session.findById("wnd[0]").sendVKey 0

    ' Preenche filtros
    session.findById("wnd[0]/usr/ctxtEL_LIFNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").Text = "*"
    session.findById("wnd[0]/usr/ctxtEL_EKORG-LOW").Text = "*"
    session.findById("wnd[0]/usr/ctxtS_EBELN-LOW").Text = pedido
    session.findById("wnd[0]/tbar[1]/btn[8]").press

    ' Extrai dados da grid
    Dim grid As Object
    Set grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")

    Dim totalLinhas As Long
    totalLinhas = grid.RowCount

    Dim linhaGrid As Long
    For linhaGrid = 0 To totalLinhas - 1
        
    Dim material As String, ncmSAP As String, utilizSAP As String
    material = grid.GetCellValue(linhaGrid, "EMATN")
    ncmSAP = grid.GetCellValue(linhaGrid, "J_1BNBM")       ' Código NCM
    utilizSAP = grid.GetCellValue(linhaGrid, "J_1BMATUSE") ' Utilização do material


        If Not dadosSAP.exists(material) Then
                dadosSAP.Add material, Array( _
                    grid.GetCellValue(linhaGrid, "VENDOR_NAME"), _
                    grid.GetCellValue(linhaGrid, "EBELN"), _
                    grid.GetCellValue(linhaGrid, "EBELP"), _
                    grid.GetCellValue(linhaGrid, "MGLIEF"), _
                    grid.GetCellValue(linhaGrid, "NETPR"), _
                    grid.GetCellValue(linhaGrid, "MWSKZ"), _
                    grid.GetCellValue(linhaGrid, "EKORG"), _
                    grid.GetCellValue(linhaGrid, "PEINH"), _
                    ncmSAP, utilizSAP _
                )
                        End If
    Next linhaGrid

    Set BuscarDadosSAPPorPedido = dadosSAP
End Function

Sub BuscarDadosDoDraft()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual


    ' [Declarações de variáveis]
    Dim draftFile As Variant
    Dim draftWB As Workbook, aliquotaWB As Workbook, atoWB As Workbook
    Dim emissaoWB As Workbook, novoWB As Workbook
    Dim wsDraft As Worksheet, wsEmissao As Worksheet
    Dim wsAliquota As Worksheet, wsAto As Worksheet, wsCNPJ As Worksheet
    Dim abaCISNova As Worksheet
    Dim taxaDolar As Double, taxaVMLE As Double, valorVMLE As Double, taxaFRETE As Double, taxaSEGURO As Double
    Dim freteMoeda As Double, seguroMoeda As Double
    Dim afrmm As Double, taxaSiscomex As Double, acrescimos As Double
    Dim i As Long, linhaDestino As Long
    Dim ncm As String, material As String, processo As String, cnpj As String
    Dim quantidade As Variant, vucv As Variant, invoice As Variant
    Dim pesoLiquido As Variant
    Dim resultadoAto As Variant, resultadoAliquota As Variant
    Dim caminhoSalvar As String
    Dim aba As Worksheet
    Dim incoterm As String
    Dim isDrawback As Boolean
    Dim celulaN As Range

    ' [Abrir arquivos]
    draftFile = Application.GetOpenFilename("Arquivos Excel (*.xlsx), *.xlsx", , "Selecione o Draft")
    If draftFile = False Then Exit Sub

    Set draftWB = Workbooks.Open(draftFile)
    Set wsDraft = draftWB.Sheets(1)
    Set emissaoWB = ThisWorkbook
    Set wsEmissao = emissaoWB.Sheets("CUSTO DE IMPORTAÇÃO - SAP")
    Set aliquotaWB = Workbooks.Open("\\gpa-fs\FCL$\Fiscal\GabrielHK\BANCO DE DADOS\Banco de dados\NCM e Aliquotas.xlsx")
    Set wsAliquota = aliquotaWB.Sheets("NCM x ALIQUOTA")
    Set wsCNPJ = aliquotaWB.Sheets("CNPJ COD")
    Set atoWB = Workbooks.Open("\\gpa-fs\FCL$\Fiscal\GabrielHK\BANCO DE DADOS\Banco de dados\AtoConcessorio.xlsx")
    Set wsAto = atoWB.Sheets("Planilha1")

    ' [Captura de dados fixos]
    taxaVMLE = wsDraft.Range("G2").Value
    taxaDolar = wsDraft.Range("N2").Value
    taxaFRETE = wsDraft.Range("N2").Value
    taxaSEGURO = wsDraft.Range("R2").Value
    valorVMLE = Application.WorksheetFunction.Sum(wsDraft.Range("AG2:AG1000"))
    freteMoeda = Application.WorksheetFunction.Sum(wsDraft.Range("P2:P1000")) / taxaFRETE
    seguroMoeda = Application.WorksheetFunction.Sum(wsDraft.Range("T2:T1000")) / taxaSEGURO
    afrmm = Application.WorksheetFunction.Sum(wsDraft.Range("AY2:AY1000"))
    taxaSiscomex = Application.WorksheetFunction.Sum(wsDraft.Range("AT2:AT1000"))
    acrescimos = Application.WorksheetFunction.Sum(wsDraft.Range("J2:J1000"))
    cnpj = wsDraft.Range("A2").Value
    processo = wsDraft.Range("B2").Value
    incoterm = wsDraft.Range("L2").Value

    ' [Preenchimento na CIS]
    wsEmissao.Range("D2").Value = taxaVMLE
    wsEmissao.Range("D3").Value = taxaDolar
    wsEmissao.Range("D4").Value = valorVMLE
    If incoterm = "DAP" Or incoterm = "CFR" Then
        wsEmissao.Range("D4").Value = valorVMLE + freteMoeda
        freteMoeda = 0
    End If
    wsEmissao.Range("D5").Value = freteMoeda
    wsEmissao.Range("D6").Value = seguroMoeda
    wsEmissao.Range("D7").Value = IIf(IsNumeric(acrescimos), acrescimos, 0)
    wsEmissao.Range("D9").Value = afrmm
    wsEmissao.Range("D10").Value = taxaSiscomex
    wsEmissao.Range("H2").Value = BuscarCodigoPlanta(cnpj, wsCNPJ)
    wsEmissao.Range("A2").Value = processo
    wsEmissao.Range("B14").Value = incoterm
    wsEmissao.Range("B12").Value = Application.WorksheetFunction.Sum(wsDraft.Range("AF2:AF1000"))
    wsEmissao.Range("B13").Value = Application.WorksheetFunction.Sum(wsDraft.Range("AH2:AH1000"))

    ' [Totais auxiliares]
    wsEmissao.Range("M2").Value = Application.WorksheetFunction.Sum(wsDraft.Range("I2:I1000"))
    wsEmissao.Range("M3").Value = Application.WorksheetFunction.Sum(wsDraft.Range("P2:P1000"))
    wsEmissao.Range("M4").Value = Application.WorksheetFunction.Sum(wsDraft.Range("T2:T1000"))
    wsEmissao.Range("M5").Value = Application.WorksheetFunction.Sum(wsDraft.Range("K2:K1000"))
    wsEmissao.Range("M6").Value = Application.WorksheetFunction.Sum(wsDraft.Range("AT2:AT1000"))
    wsEmissao.Range("M7").Value = Application.WorksheetFunction.Sum(wsDraft.Range("AY2:AY1000"))
    wsEmissao.Range("M8").Value = Application.WorksheetFunction.Sum(wsDraft.Range("U2:U1000"))

    wsEmissao.Range("M9").Value = SomaOuTexto(wsDraft.Range("AQ2:AQ1000"))
    
    ' M9: Se houver "INTEGRAL" em qualquer parte do texto
    If UCase(wsEmissao.Range("M9").Text) Like "*INTEGRAL*" Then
        wsEmissao.Range("M9").Formula = "=K9"
    End If

    wsEmissao.Range("M10").Value = SomaOuTexto(wsDraft.Range("AS2:AS1000"))
        
    ' M10: Se houver "INTEGRAL" em qualquer parte do texto
    If UCase(wsEmissao.Range("M10").Text) Like "*INTEGRAL*" Then
        wsEmissao.Range("M10").Formula = "=K10"
    End If
   
    wsEmissao.Range("M11").Value = Application.WorksheetFunction.Sum(wsDraft.Range("AV2:AV1000"))
    wsEmissao.Range("M12").Value = Application.WorksheetFunction.Sum(wsDraft.Range("AX2:AX1000"))
    wsEmissao.Range("M13").Value = Application.WorksheetFunction.Sum(wsDraft.Range("AM2:AM1000"))
    wsEmissao.Range("M16").Value = Application.WorksheetFunction.Sum(wsDraft.Range("AK2:AK1000"))
    
    ' [Loop de materiais]
    linhaDestino = 20
    i = 2
    
    'teste
    Dim pedidosSAP As Object
    Set pedidosSAP = CreateObject("Scripting.Dictionary")
    'teste
    
    Do While Trim(wsDraft.Cells(i, 24).Value) <> ""
    material = wsDraft.Cells(i, 24).Value
    ncm = wsDraft.Cells(i, 27).Value
    quantidade = wsDraft.Cells(i, 30).Value
    pesoLiquido = wsDraft.Cells(i, 36).Value
    vucv = wsDraft.Cells(i, 34).Value
    invoice = wsDraft.Cells(i, 29).Value
    
    Dim pedidoDraft As String, materialDraft As String
    pedidoDraft = wsDraft.Cells(i, 59).Value ' Coluna BG
    materialDraft = wsDraft.Cells(i, 24).Value ' Coluna X
    
    ' Busca dados do SAP uma vez por pedido
    If Not pedidosSAP.exists(pedidoDraft) Then
        Set pedidosSAP(pedidoDraft) = BuscarDadosSAPPorPedido(pedidoDraft)
    End If
    
    ' Preenche dados do SAP se o material existir
    If pedidosSAP(pedidoDraft).exists(materialDraft) Then
        
        Dim dadosItem As Variant
        dadosItem = pedidosSAP(pedidoDraft)(materialDraft)
    
        wsEmissao.Cells(linhaDestino, 1).Value = Left(dadosItem(0), 6)
        wsEmissao.Cells(linhaDestino, 2).Value = dadosItem(1) ' Pedido
        wsEmissao.Cells(linhaDestino, 3).Value = dadosItem(2) ' Linha
        wsEmissao.Cells(linhaDestino, 40).Value = dadosItem(8) ' NCM
        wsEmissao.Cells(linhaDestino, 41).Value = dadosItem(9) ' Utilização
       
        With wsEmissao.Cells(linhaDestino, 38)
            .NumberFormat = "#.##0,00" ' Exibe com separador de milhar e 2 casas decimais
        
            Dim valorSAP As String
            valorSAP = Trim(CStr(dadosItem(3)))
        
            ' Remove pontos de milhar e mantém vírgula como separador decimal
            If InStr(valorSAP, ",") > 0 Then
                valorSAP = Replace(valorSAP, ".", "") ' remove milhar
            Else
                valorSAP = Replace(valorSAP, ".", "") ' remove milhar
            End If
        
            If IsNumeric(valorSAP) Then
                .Value = CDbl(valorSAP)
            Else
                .Value = valorSAP ' fallback
            End If
        End With
                        
        ' === Cálculo do valor unitário SAP com base em CTD (PEINH) ===
        Dim valorPedidoSAP As Double
        Dim ctdSAP As Double
        Dim valorUnitarioSAP As Double
        
        If IsNumeric(dadosItem(4)) And IsNumeric(dadosItem(7)) Then
            valorPedidoSAP = CDbl(dadosItem(4))
            ctdSAP = CDbl(dadosItem(7))
            
            If ctdSAP <> 0 Then
                valorUnitarioSAP = valorPedidoSAP / ctdSAP
            Else
                valorUnitarioSAP = 0
            End If
        Else
            valorUnitarioSAP = 0
        End If
        
        With wsEmissao.Cells(linhaDestino, 39)
            .NumberFormat = "#,##0.000000" ' separador de milhar + 6 casas decimais
            .Value = Round(valorUnitarioSAP, 6)
        End With
    End If
    
    wsEmissao.Cells(linhaDestino, 4).Value = material
    wsEmissao.Cells(linhaDestino, 5).Value = ncm
    wsEmissao.Cells(linhaDestino, 11).Value = quantidade
    wsEmissao.Cells(linhaDestino, 13).Value = pesoLiquido
    wsEmissao.Cells(linhaDestino, 10).Value = vucv
    wsEmissao.Cells(linhaDestino, 12).Value = invoice
    wsEmissao.Cells(linhaDestino, 13).Value = wsDraft.Cells(i, 36).Value
    wsEmissao.Cells(linhaDestino, 14).Value = wsDraft.Cells(i, 35).Value

    ' === Validação de saldo SAP vs. quantidade ===
    With wsEmissao.Cells(linhaDestino, 11)
        If IsNumeric(.Value) And IsNumeric(wsEmissao.Cells(linhaDestino, 38).Value) Then
            If .Value > wsEmissao.Cells(linhaDestino, 38).Value Then
                .Interior.Color = RGB(255, 199, 206) ' vermelho pastel
                If Not .Comment Is Nothing Then .Comment.Delete
                .AddComment "Sem saldo suficiente no pedido"
            Else
                .Interior.ColorIndex = xlNone
                If Not .Comment Is Nothing Then .Comment.Delete
            End If
        End If
    End With
    
  ' === Validação de suspensão do IPI ===
Dim utilizacaoSAP As String, plantaSAP As String, plantaReferencia As String
utilizacaoSAP = Trim(wsEmissao.Cells(linhaDestino, 41).Value) ' Coluna com utilização
plantaSAP = Trim(wsEmissao.Cells(linhaDestino, 1).Value)       ' Coluna com código da planta
plantaReferencia = UCase(Trim(wsEmissao.Range("H2").Value))    ' Valor da célula H2 (em maiúsculas para padronizar)

If utilizacaoSAP = "1" Then
    ' Verifica se H2 contém SBR, GR2 ou SOR
    If plantaReferencia <> "SBR" And plantaReferencia <> "GR2" And plantaReferencia <> "SOR" Then
        wsEmissao.Cells(linhaDestino, 16).Interior.Color = RGB(255, 199, 206) ' vermelho pastel
        If Not wsEmissao.Cells(linhaDestino, 16).Comment Is Nothing Then
            wsEmissao.Cells(linhaDestino, 16).Comment.Delete
        End If
        wsEmissao.Cells(linhaDestino, 16).AddComment "IPI suspenso conforme utilização 1"
    Else
        wsEmissao.Cells(linhaDestino, 16).Interior.ColorIndex = xlNone
        If Not wsEmissao.Cells(linhaDestino, 16).Comment Is Nothing Then
            wsEmissao.Cells(linhaDestino, 16).Comment.Delete
        End If
    End If
End If
    
 ' === Validação de NCM SAP vs. NCM da planilha ===
Dim ncmPlanilha As String, ncmSAP As String
ncmPlanilha = wsEmissao.Cells(linhaDestino, 5).Value   ' NCM da planilha
ncmSAP = wsEmissao.Cells(linhaDestino, 40).Value       ' NCM do SAP

If Trim(ncmPlanilha) <> "" And Trim(ncmSAP) <> "" Then
    With wsEmissao.Cells(linhaDestino, 5) ' Comentário e cor na coluna 5
        If ncmPlanilha <> ncmSAP Then
            .Interior.Color = RGB(255, 199, 206) ' vermelho pastel
            If Not .Comment Is Nothing Then .Comment.Delete
            .AddComment "NCM divergente do SAP"
        Else
            .Interior.ColorIndex = xlNone
            If Not .Comment Is Nothing Then .Comment.Delete
        End If
    End With
End If
    
    ' === Validação de valor unitário SAP vs. planilha ===
    With wsEmissao.Cells(linhaDestino, 10)
        If IsNumeric(.Value) And IsNumeric(wsEmissao.Cells(linhaDestino, 39).Value) Then
            If Round(.Value, 2) <> Round(wsEmissao.Cells(linhaDestino, 39).Value, 2) Then
                .Interior.Color = RGB(255, 199, 206) ' vermelho pastel
                If Not .Comment Is Nothing Then .Comment.Delete
                .AddComment "Vlr unitário divergente no pedido"
            Else
                .Interior.ColorIndex = xlNone
                If Not .Comment Is Nothing Then .Comment.Delete
            End If
        End If
    End With

    resultadoAto = VerificarAtoConcessorio(material, ncm, wsAto)
    wsEmissao.Cells(linhaDestino, 6).Value = resultadoAto(0)

    ' [Validação de saldo progressivo]
    Dim chaveItem As String
    chaveItem = material & "|" & ncm

    Dim saldoDisponivel As Double
    If IsNumeric(resultadoAto(1)) Then
        saldoDisponivel = resultadoAto(1)
    Else
        saldoDisponivel = 0
    End If

    ' Inicializa o dicionário de controle de saldo se necessário
    Dim saldoUtilizado As Object
    If saldoUtilizado Is Nothing Then
        Set saldoUtilizado = CreateObject("Scripting.Dictionary")
    End If

    Dim saldoJaUtilizado As Double
    If saldoUtilizado.exists(chaveItem) Then
        saldoJaUtilizado = saldoUtilizado(chaveItem)
    Else
        saldoJaUtilizado = 0
    End If

    Dim saldoRestante As Double
    saldoRestante = saldoDisponivel - saldoJaUtilizado

    If saldoRestante >= pesoLiquido Then
        wsEmissao.Cells(linhaDestino, 6).Value = "TEM"
        saldoUtilizado(chaveItem) = saldoJaUtilizado + pesoLiquido

        wsEmissao.Cells(linhaDestino, 7).Value = resultadoAto(1)
        wsEmissao.Cells(linhaDestino, 8).Value = resultadoAto(2)
        wsEmissao.Cells(linhaDestino, 9).Value = resultadoAto(3)

    ElseIf saldoRestante > 0 Then
        wsEmissao.Cells(linhaDestino, 6).Value = "TEM (PARCIAL)"
        saldoUtilizado(chaveItem) = saldoJaUtilizado + pesoLiquido

        wsEmissao.Cells(linhaDestino, 7).Value = resultadoAto(1)
        wsEmissao.Cells(linhaDestino, 8).Value = resultadoAto(2)
        wsEmissao.Cells(linhaDestino, 9).Value = resultadoAto(3)

    Else
        If resultadoAto(0) = "NÃO TEM" Then
            wsEmissao.Cells(linhaDestino, 6).Value = "NÃO TEM"
            wsEmissao.Cells(linhaDestino, 7).Value = "Não tem"
            wsEmissao.Cells(linhaDestino, 8).Value = "Não tem"
            wsEmissao.Cells(linhaDestino, 9).Value = "Não tem"
        Else
            wsEmissao.Cells(linhaDestino, 6).Value = "TEM (S/SALDO)"
            wsEmissao.Cells(linhaDestino, 7).Value = "S/ saldo"
            wsEmissao.Cells(linhaDestino, 8).Value = "S/ saldo"
            wsEmissao.Cells(linhaDestino, 9).Value = "S/ saldo"
        End If

        resultadoAliquota = BuscarAliquotas(ncm, wsAliquota)
        wsEmissao.Cells(linhaDestino, 15).Value = resultadoAliquota(0) ' II
        wsEmissao.Cells(linhaDestino, 16).Value = resultadoAliquota(1) ' IPI
        wsEmissao.Cells(linhaDestino, 17).Value = resultadoAliquota(2) ' PIS
        wsEmissao.Cells(linhaDestino, 18).Value = resultadoAliquota(3) ' COFINS
    End If

    wsEmissao.Cells(linhaDestino, 24).Value = wsDraft.Cells(i, 38).Value ' ICMS (AL)
    wsEmissao.Cells(linhaDestino, 20).Value = wsDraft.Cells(i, 42).Value ' II (AP)
    wsEmissao.Cells(linhaDestino, 21).Value = wsDraft.Cells(i, 44).Value ' IPI (AR)
    wsEmissao.Cells(linhaDestino, 22).Value = wsDraft.Cells(i, 47).Value ' PIS (AU)
    wsEmissao.Cells(linhaDestino, 23).Value = wsDraft.Cells(i, 49).Value ' COFINS (AW)
    
    ' Comparação de alíquotas e destaque em vermelho se houver divergência
    Dim celDraft As Range, celOficial As Range
    Dim colDraftAliquotas As Variant, colOficialAliquotas As Variant
    Dim valorDraft As Double, valorOficial As Double
    Dim k As Integer

    colDraftAliquotas = Array(20, 21, 22, 23, 24)   ' II, IPI, PIS, COFINS, ICMS
    colOficialAliquotas = Array(15, 16, 17, 18, 19) ' II, IPI, PIS, COFINS, ICMS

    For k = 0 To 4
        Set celDraft = wsEmissao.Cells(linhaDestino, colDraftAliquotas(k))
        Set celOficial = wsEmissao.Cells(linhaDestino, colOficialAliquotas(k))

        If IsNumeric(celDraft.Value) And IsNumeric(celOficial.Value) Then
            valorDraft = CDbl(celDraft.Value)
            valorOficial = CDbl(celOficial.Value)

            ' Só divide se o valor for maior que 1 e menor que 100

            If Not celDraft.Column = 24 Then
                If valorDraft > 1 And valorDraft <= 100 Then
                    valorDraft = valorDraft / 100
                    celDraft.Value = Round(valorDraft, 4)
                End If
                    celDraft.NumberFormat = "0.00%"
            End If
            
            ' Compara os valores
            If Round(valorDraft, 2) <> Round(valorOficial, 2) Then
                celDraft.Interior.Color = RGB(255, 199, 206) ' vermelho pastel
            Else
                celDraft.Interior.ColorIndex = xlNone ' limpa cor se estiver igual
            End If
        Else
            celDraft.Interior.ColorIndex = xlNone ' limpa cor se não for número
        End If
    Next k

    linhaDestino = linhaDestino + 1
    i = i + 1
    Loop

    ' [Criação do novo arquivo]
    Set novoWB = Workbooks.Add
    emissaoWB.Sheets("CUSTO DE IMPORTAÇÃO - SAP").Visible = xlSheetVisible
    emissaoWB.Sheets("IMPOSTOS").Visible = xlSheetVisible
    emissaoWB.Sheets("CUSTO DE IMPORTAÇÃO - SAP").Copy After:=novoWB.Sheets(novoWB.Sheets.Count)
    emissaoWB.Sheets("IMPOSTOS").Copy After:=novoWB.Sheets(novoWB.Sheets.Count)

    ' Renomear a aba "CUSTO DE IMPORTAÇÃO - SAP" com o nome do processo
    Dim nomeAbaProcesso As String
    nomeAbaProcesso = SanitizarNomeArquivo(processo)

    Dim abaProcesso As Worksheet
    Set abaProcesso = novoWB.Sheets(novoWB.Sheets.Count - 1) ' A aba do processo foi copiada antes da aba IMPOSTOS
    abaProcesso.Name = nomeAbaProcesso

    ' Substituir fórmulas na aba do processo
    Dim cel As Range
    For Each cel In novoWB.Sheets(nomeAbaProcesso).UsedRange
        If cel.HasFormula Then
            cel.Formula = Replace(cel.Formula, "[" & ThisWorkbook.Name & "]", "")
            cel.Formula = Replace(cel.Formula, "'CUSTO DE IMPORTAÇÃO - SAP'", "'" & nomeAbaProcesso & "'")
        End If
    Next cel
    
    ' Substituir referências externas à aba "CUSTO DE IMPORTAÇÃO - SAP" na aba IMPOSTOS
    With novoWB.Sheets("IMPOSTOS")
        .Activate
        For Each cel In .UsedRange
            If cel.HasFormula Then
                On Error Resume Next
                Dim formulaOriginal As String
                formulaOriginal = cel.Formula
    
                ' Remove referência ao nome do arquivo da planilha mãe
                formulaOriginal = Replace(formulaOriginal, "[" & ThisWorkbook.Name & "]", "")
    
                ' Substitui qualquer ocorrência de 'CUSTO DE IMPORTAÇÃO - SAP' por nomeAbaProcesso
                formulaOriginal = Replace(formulaOriginal, "'CUSTO DE IMPORTAÇÃO - SAP'", "'" & nomeAbaProcesso & "'")
                formulaOriginal = Replace(formulaOriginal, "CUSTO DE IMPORTAÇÃO - SAP", nomeAbaProcesso)
    
                cel.Formula = formulaOriginal
                On Error GoTo 0
            End If
        Next cel
    End With

    nomeAbaProcesso = SanitizarNomeArquivo(processo)
    emissaoWB.Sheets("CUSTO DE IMPORTAÇÃO - SAP").Visible = xlSheetVeryHidden
    emissaoWB.Sheets("IMPOSTOS").Visible = xlSheetVeryHidden
    emissaoWB.Sheets("MACRO").Visible = xlSheetVisible
    
    ' [Inserir fórmula na célula D8 da nova aba]
    Set abaCISNova = novoWB.Sheets(nomeAbaProcesso)
    abaCISNova.Range("D8").FormulaLocal = "=SOMA(D4:D7)"
    Set abaCISNova = novoWB.Sheets(nomeAbaProcesso)
    abaCISNova.Range("B12").FormulaLocal = "=SOMA(N20:N719)"
    Set abaCISNova = novoWB.Sheets(nomeAbaProcesso)
    abaCISNova.Range("B13").FormulaLocal = "=SOMA(M20:M719)"
    Set abaCISNova = novoWB.Sheets(nomeAbaProcesso)
    abaCISNova.Range("M14").FormulaLocal = "=SOMA(M9:M13)"
    Set abaCISNova = novoWB.Sheets(nomeAbaProcesso)
    abaCISNova.Range("M15").FormulaLocal = "=M14+M6"
    
    Application.DisplayAlerts = False
    For Each aba In novoWB.Sheets
        If aba.Name <> nomeAbaProcesso And aba.Name <> "IMPOSTOS" Then
            aba.Delete
        End If
    Next aba
    Application.DisplayAlerts = True
    
    ' Mover a aba "CUSTO DE IMPORTAÇÃO - SAP" para a primeira posição
    novoWB.Sheets(nomeAbaProcesso).Move Before:=novoWB.Sheets(1)
    
    ' Após copiar a aba "CUSTO DE IMPORTAÇÃO - SAP" para novoWB
    Dim btn As Button
    Set abaCISNova = novoWB.Sheets(nomeAbaProcesso)
    
    ' Ativar o novo arquivo antes de salvar
    novoWB.Activate

    ' Remover nomes definidos que referenciam a planilha mãe
    Dim nm As Name
    For Each nm In novoWB.Names
        If InStr(nm.RefersTo, ThisWorkbook.Name) > 0 Then
            nm.Delete
        End If
    Next nm
    Dim moduloOrigem As VBComponent

    ' Copiar módulo da macro para o novo arquivo
    
    Dim caminhoModuloTemp As String
    Dim novoModulo As VBComponent
    Dim btnVerificarDI As Button
    Dim btnEmissaoNF As Button
    
    ' Exporta o módulo Módulo3 (MacroEmissaoNF)
    Dim caminhoModuloEmissaoNF As String
    
    caminhoModuloEmissaoNF = Environ("TEMP") & "\ModuloEmissaoNF.bas"
    
    Set moduloOrigem = ThisWorkbook.VBProject.VBComponents("Módulo3")
    moduloOrigem.Export caminhoModuloEmissaoNF
    novoWB.VBProject.VBComponents.Import caminhoModuloEmissaoNF
    
    Set novoModulo = novoWB.VBProject.VBComponents(novoWB.VBProject.VBComponents.Count)
    novoModulo.Name = "ModuloEmissaoNF"
    
    ' Exporta o módulo Módulo5 (VerificarDI)
    Dim caminhoModuloVerificarDI As String
    caminhoModuloVerificarDI = Environ("TEMP") & "\ModuloVerificarDI.bas"
    
    Set moduloOrigem = ThisWorkbook.VBProject.VBComponents("Módulo5")
    moduloOrigem.Export caminhoModuloVerificarDI
    novoWB.VBProject.VBComponents.Import caminhoModuloVerificarDI
    
    Set novoModulo = novoWB.VBProject.VBComponents(novoWB.VBProject.VBComponents.Count)
    novoModulo.Name = "ModuloVerificarDI"
    
    ' Adiciona botão "Emissão NF" entre A721 e B724
    Set btnEmissaoNF = abaCISNova.Buttons.Add(abaCISNova.Range("S5").Left, abaCISNova.Range("S5").Top, 150, 30)
    With btnEmissaoNF
        .Caption = "Emissão NF"
        .OnAction = "'" & novoWB.Name & "'!MacroEmissaoNF"
        .Name = "btnEmissaoNF"
    End With
    

    'Block button
    Dim codModulo As CodeModule
    Set codModulo = novoWB.VBProject.VBComponents(abaCISNova.CodeName).CodeModule
    Dim linhaInsercao As Long
    
    'Evento Worksheet_Change
    linhaInsercao = codModulo.CountOfLines + 1
    codModulo.InsertLines linhaInsercao, _
    "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf & _
    "    Call AtualizarBotaoEmissaoNF" & vbCrLf & _
    "End Sub"
    
    'Evento Worksheet_Calculate
    linhaInsercao = codModulo.CountOfLines + 1
    codModulo.InsertLines linhaInsercao, _
    "Private Sub Worksheet_Calculate()" & vbCrLf & _
    "    Call AtualizarBotaoEmissaoNF" & vbCrLf & _
    "End Sub"
    
    'Procedimento AtualizarBotaoEmissaoNF
    linhaInsercao = codModulo.CountOfLines + 1
    codModulo.InsertLines linhaInsercao, _
    "Private Sub AtualizarBotaoEmissaoNF()" & vbCrLf & _
    "    Dim celulaVerificacao As Range" & vbCrLf & _
    "    Dim somaValores As Double" & vbCrLf & _
    "    Dim podeMostrarBotao As Boolean" & vbCrLf & _
    "    somaValores = 0" & vbCrLf & _
    "    Set celulaVerificacao = Me.Range(""L2:L16"")" & vbCrLf & _
    "    For Each cel In celulaVerificacao" & vbCrLf & _
    "        If IsNumeric(cel.Value) Then" & vbCrLf & _
    "            somaValores = somaValores + Round(cel.Value, 2)" & vbCrLf & _
    "        Else" & vbCrLf & _
    "            podeMostrarBotao = False" & vbCrLf & _
    "            Me.Buttons(""btnEmissaoNF"").Visible = False" & vbCrLf & _
    "            Exit Sub" & vbCrLf & _
    "        End If" & vbCrLf & _
    "    Next cel" & vbCrLf & _
    "    If somaValores >= -0.05 And somaValores <= 0.05 Then" & vbCrLf & _
    "        podeMostrarBotao = True" & vbCrLf & _
    "    Else" & vbCrLf & _
    "        podeMostrarBotao = False" & vbCrLf & _
    "    End If" & vbCrLf & _
    "    Me.Buttons(""btnEmissaoNF"").Visible = podeMostrarBotao" & vbCrLf & _
    "End Sub"


    ' Adiciona botão "Verificar DI" na célula R1
    Set btnVerificarDI = abaCISNova.Buttons.Add(abaCISNova.Range("S1").Left, abaCISNova.Range("S1").Top, 150, 30)
    With btnVerificarDI
        .Caption = "Verificar DI"
        .OnAction = "'" & novoWB.Name & "'!VerificarDI"
        .Name = "btnVerificarDI"
    End With
    
    ' Salvar novo arquivo com a macro incluída
    caminhoSalvar = "\\gpa-fs\FCL$\Fiscal\2.RECEBIMENTO\09. IMPORTAÇÃO\IMPORTAÇÃO\TESTE PROJETO\DOC SALVO\" & SanitizarNomeArquivo(processo) & ".xlsm"
    novoWB.SaveAs Filename:=caminhoSalvar, FileFormat:=xlOpenXMLWorkbookMacroEnabled

    ' Fechar todos os arquivos auxiliares
    novoWB.Close SaveChanges:=False
    draftWB.Close SaveChanges:=False
    aliquotaWB.Close SaveChanges:=False
    atoWB.Close SaveChanges:=False

    ' [Limpeza da planilha mãe]
    With wsEmissao
        .Range("A2").ClearContents
        .Range("D2:D10").ClearContents
        .Range("H2").ClearContents
        .Range("B12:B14").ClearContents
        .Range("D20:R" & linhaDestino - 1).ClearContents
        .Range("T20:W" & linhaDestino - 1).ClearContents
        .Range("X20:X" & linhaDestino - 1).ClearContents
        .Range("M2:M16").ClearContents
        .Range("B12").ClearContents
        .Range("B13").ClearContents
        .Range("D8").ClearContents
        .Range("A20:A719").ClearContents
        .Range("B20:B719").ClearContents
        .Range("C20:C719").ClearContents
        .Range("AL20:AL719").ClearContents
        .Range("AM20:AM719").ClearContents
        .Range("AN20:AN719").ClearContents
        .Range("AO20:AO719").ClearContents
                
    Dim linhaLimpeza As Long
    For linhaLimpeza = 20 To linhaDestino - 1
        wsEmissao.Range("T" & linhaLimpeza & ":X" & linhaLimpeza).Interior.ColorIndex = xlNone
    Next linhaLimpeza

    For linhaLimpeza = 20 To linhaDestino - 1
        ' Limpa cores das colunas de validação (J a K)
        wsEmissao.Range("J" & linhaLimpeza & ":K" & linhaLimpeza).Interior.ColorIndex = xlNone
    
        ' Limpa cor e comentário da quantidade (coluna K = 11)
        With wsEmissao.Cells(linhaLimpeza, 11)
            .Interior.ColorIndex = xlNone
            If Not .Comment Is Nothing Then .Comment.Delete
        End With
    
        ' Limpa cor e comentário do valor unitário (coluna J = 10)
        With wsEmissao.Cells(linhaLimpeza, 10)
            .Interior.ColorIndex = xlNone
            If Not .Comment Is Nothing Then .Comment.Delete
        End With
    Next linhaLimpeza

    End With

    MsgBox "Dados importados, arquivo salvo e planilha limpa com sucesso!"
    
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True


End Sub

------------Módulo 2------------
Sub DesbloquearAbasOcultas()
    Dim senha As String
    Dim aba As Worksheet

    ' Solicita a senha ao usuário
    senha = InputBox("Digite a senha para desbloquear as abas:")

    ' Verifica a senha
    If senha <> "230919" Then
        MsgBox "Senha incorreta. Acesso negado.", vbCritical
        Exit Sub
    End If

    ' Torna todas as abas visíveis
    For Each aba In ThisWorkbook.Sheets
        aba.Visible = xlSheetVisible
    Next aba

    MsgBox "Todas as abas foram desbloqueadas com sucesso!", vbInformation

End Sub

---------Módulo 3-------------
Option Explicit

Public Appl, SapGuiAuto, Connection, session, WScript

Function Nz(valor As Variant, Optional padrao As Variant = 0) As Variant
    If IsEmpty(valor) Or IsNull(valor) Or Trim(valor) = "" Then
        Nz = padrao
    Else
        Nz = valor
    End If
End Function

Sub MacroEmissaoNF()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim linha As Long
    Dim itemIndex As Integer
    Dim nftype As String
    Dim bukrs As String
    
    Set wb = ThisWorkbook
    Set ws = wb.Sheets(1)

    ' Conexão com o SAP
    If Not IsObject(Appl) Then
       Set SapGuiAuto = GetObject("SAPGUI")
       Set Appl = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(Connection) Then
       Set Connection = Appl.Children(0)
    End If
    If Not IsObject(session) Then
       Set session = Connection.Children(0)
    End If
    If IsObject(WScript) Then
       WScript.ConnectObject session, "on"
       WScript.ConnectObject Application, "on"
    End If

    ' Solicita ao usuário os valores de NFTYPE e BUKRS
    Dim cfop As String
    Dim encontrado As Boolean

    encontrado = False

    ' Percorre as linhas da planilha para encontrar o CFOP
    For linha = 1432 To 2131
        cfop = Trim(ws.Cells(linha, "O").Value) ' Coluna O = CFOP

        Select Case cfop
            Case "3101/AA"
                nftype = "Z4"
                encontrado = True
                Exit For
            Case "3551/AA"
                nftype = "Z8"
                encontrado = True
                Exit For
            Case "3949/AA", "3556/AA"
                nftype = "Z9"
                encontrado = True
                Exit For
        End Select
    Next linha

    If Not encontrado Then
        MsgBox "CFOP não reconhecido. Tipo de nota não pode ser definido automaticamente.", vbExclamation
        Exit Sub
    Else
    End If
    If MsgBox("Tipo de nota identificado como '" & nftype & "'. Deseja continuar?", vbYesNo + vbQuestion, "Confirmação") = vbNo Then
        Exit Sub
    End If
   

    Dim empresaPlanilha As String

    empresaPlanilha = Trim(ws.Range("H2").Value)

    Select Case empresaPlanilha
        Case "GPA", "MBB", "MB2", "GRA", "GR2", "GSN", "BMG", "SOR", "GPI"
            bukrs = "GPA"
        Case "SBR"
            bukrs = "SBR"
        Case Else
            MsgBox "Empresa não reconhecida na célula H2. Operação cancelada.", vbExclamation
            Exit Sub
    End Select

    ' Confirma com o usuário
    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("Empresa identificada como '" & bukrs & "'. Deseja continuar?", vbYesNo + vbQuestion, "Confirmação")

    If resposta = vbNo Then
        MsgBox "Operação cancelada.", vbInformation
        Exit Sub
    End If

    ' Acessa a transação SAP /nzj1b1n
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nzj1b1n"
    session.findById("wnd[0]").sendVKey 0

    ' Preenche os campos iniciais da tela
    session.findById("wnd[0]/usr/chkJ_1BDYLIN-INCLTX").Selected = True
    session.findById("wnd[0]/usr/ctxtJ_1BDYDOC-NFTYPE").Text = nftype
    session.findById("wnd[0]/usr/ctxtJ_1BDYDOC-BUKRS").Text = bukrs
    session.findById("wnd[0]/usr/ctxtJ_1BDYDOC-BRANCH").Text = ws.Range("H2").Value
    session.findById("wnd[0]/usr/cmbJ_1BDYDOC-PARVW").Key = "LF"
    session.findById("wnd[0]/usr/chkJ_1BDYLIN-INCLTX").Selected = True
    session.findById("wnd[0]/usr/chkJ_1BDYLIN-INCLTX").SetFocus
    session.findById("wnd[0]/usr/ctxtJ_1BDYDOC-PARID").Text = ws.Range("H13").Value
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/txtZTMM_XML_IMP-NDI").Text = ws.Range("H5").Value
    session.findById("wnd[0]/usr/ctxtZTMM_XML_IMP-DDI").Text = ws.Range("H6").Value

     ' === Etapa 1: Preenchimento da primeira tela do SAP ===
    Dim i As Integer
    Dim linhaSAP1 As Integer
    Dim linhasValidas As Integer
    Dim linhasVisiveis As Integer
    Dim cliquesNecessarios As Integer
    
    ' Conta quantas linhas válidas existem no Excel
    linhasValidas = 0
    For linha = 728 To 1427
        If ws.Cells(linha, "J").Value <> 0 Then
            linhasValidas = linhasValidas + 1
        End If
    Next linha
    
    ' Define quantas linhas o SAP mostra por padrão (ajuste se necessário)
    linhasVisiveis = 10 ' Exemplo: SAP mostra 10 linhas inicialmente
    ' Calcula quantos cliques no botão INSE são necessários
    If linhasValidas > linhasVisiveis Then
        cliquesNecessarios = linhasValidas - linhasVisiveis
        For i = 1 To cliquesNecessarios
            session.findById("wnd[0]/usr/cntlV_CONT_9002/shellcont/shell").pressToolbarButton "INSE"
        Next i
    End If
    
    ' Agora preenche as linhas
    linhaSAP1 = 0
    For linha = 728 To 1427
        If ws.Cells(linha, "J").Value <> 0 Then
            session.findById("wnd[0]/usr/cntlV_CONT_9002/shellcont/shell").modifyCell linhaSAP1, "NADICAO", ws.Cells(linha, "A").Value
            session.findById("wnd[0]/usr/cntlV_CONT_9002/shellcont/shell").modifyCell linhaSAP1, "NSEQADIC", ws.Cells(linha, "B").Value
            session.findById("wnd[0]/usr/cntlV_CONT_9002/shellcont/shell").modifyCell linhaSAP1, "XLOCDESEMB", ws.Cells(linha, "C").Value
            session.findById("wnd[0]/usr/cntlV_CONT_9002/shellcont/shell").modifyCell linhaSAP1, "UFDESEMB", ws.Cells(linha, "D").Value
            session.findById("wnd[0]/usr/cntlV_CONT_9002/shellcont/shell").modifyCell linhaSAP1, "DDESEMB", ws.Cells(linha, "E").Value
            session.findById("wnd[0]/usr/cntlV_CONT_9002/shellcont/shell").modifyCell linhaSAP1, "CEXPORTADOR", ws.Cells(linha, "F").Value
            session.findById("wnd[0]/usr/cntlV_CONT_9002/shellcont/shell").modifyCell linhaSAP1, "CFABRICANTE", ws.Cells(linha, "G").Value
            session.findById("wnd[0]/usr/cntlV_CONT_9002/shellcont/shell").modifyCell linhaSAP1, "VDESCDI", ws.Cells(linha, "H").Value
            session.findById("wnd[0]/usr/cntlV_CONT_9002/shellcont/shell").modifyCell linhaSAP1, "CUSTVENDOR", ws.Cells(linha, "I").Value
            session.findById("wnd[0]/usr/cntlV_CONT_9002/shellcont/shell").modifyCell linhaSAP1, "O_VDESPADU", Format(ws.Cells(linha, "J").Value, "0.00")
            session.findById("wnd[0]/usr/cntlV_CONT_9002/shellcont/shell").modifyCell linhaSAP1, "O_VIOF", ws.Cells(linha, "K").Value
            session.findById("wnd[0]/usr/cntlV_CONT_9002/shellcont/shell").modifyCell linhaSAP1, "TP_VIA_TRANSP", ws.Cells(linha, "M").Value
            session.findById("wnd[0]/usr/cntlV_CONT_9002/shellcont/shell").modifyCell linhaSAP1, "TP_INTERMEDIO", ws.Cells(linha, "N").Value
            linhaSAP1 = linhaSAP1 + 1
        End If
    Next linha
    
    session.findById("wnd[0]/tbar[0]/btn[11]").press

    
    ' === Etapa 2: Preenchimento dos itens da NF e impostos ===
    itemIndex = 0
    For linha = 1432 To 2131
        If ws.Cells(linha, "B").Value <> "" And ws.Cells(linha, "B").Value <> "0" Then
            session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-ITMTYP[1," & itemIndex & "]").Text = ws.Cells(linha, "A").Value
            session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-MATNR[4," & itemIndex & "]").Text = ws.Cells(linha, "B").Value
            session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-WERKS[6," & itemIndex & "]").Text = ws.Cells(linha, "D").Value
            session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/txtJ_1BDYLIN-MENGE[10," & itemIndex & "]").Text = ws.Cells(linha, "H").Value
            session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/txtJ_1BDYLIN-NETPR[11," & itemIndex & "]").Text = Replace(Format(ws.Cells(linha, "I").Value, "0.000000"), ".", ",")
            session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/txtJ_1BDYLIN-NETOTH[15," & itemIndex & "]").Text = Replace(Format(ws.Cells(linha, "M").Value, "0.00"), ".", ",")
            session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-CFOP[17," & itemIndex & "]").Text = ws.Cells(linha, "O").Value
            session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-TAXLW1[18," & itemIndex & "]").Text = ws.Cells(linha, "P").Value
            session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-TAXLW2[19," & itemIndex & "]").Text = ws.Cells(linha, "Q").Value
            session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-TAXLW4[21," & itemIndex & "]").Text = ws.Cells(linha, "S").Value
            session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-TAXLW5[22," & itemIndex & "]").Text = ws.Cells(linha, "T").Value
            itemIndex = itemIndex + 1
        End If
    Next linha
    
    'Pressiona Enter até o aviso amarelo desaparecer
    Do
        session.findById("wnd[0]").sendVKey 0
        If session.findById("wnd[0]/sbar").MessageType <> "W" Then
            Exit Do
        End If
    Loop

    Dim wsImpostos As Worksheet
    Dim j As Integer
    Dim linhaBase As Long
    Dim impostoTipos As Variant
    Dim colBase As Integer, colAliquota As Integer, colOutrasBases As Integer

    ' Definir aba de impostos
    Set wsImpostos = ThisWorkbook.Sheets("IMPOSTOS")

    ' Conectar ao SAP
    Set session = GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(0)

    ' Tipos de impostos esperados
    impostoTipos = Array("ICON", "ICM1", "II01", "IPI1", "IPIS", "ICM2", "IPI2")

    ' Colunas da planilha
    colBase = 5         ' Coluna E = BASE
    colAliquota = 6     ' Coluna F = ALÍQUOTA
    colOutrasBases = 9  ' Coluna I = OUTRAS BASES


    ' Loop pelos blocos de impostos
    i = 0
    linhaBase = 1

    Do While wsImpostos.Cells(linhaBase, 1).Value <> ""
        If Left(wsImpostos.Cells(linhaBase, 1).Value, 5) = "ITEM " Then

        ' Tenta acessar o item no SAP
        On Error Resume Next
        session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-MATNR[4," & i & "]").SetFocus
        ' Continua o processo normalmente
        session.findById("wnd[0]").sendVKey 2
        session.findById("wnd[0]/usr/tabsITEM_TAB/tabpTAX").Select

        If Err.Number <> 0 Then
            ' Se não encontrar o item, sai do loop de impostos e continua a macro
            Err.Clear
            Exit Do ' ou: GoTo ProximaEtapa
        End If
        On Error GoTo 0

        ' Preenche os impostos
        For j = 0 To 4
            Dim linhaImposto As Long
            linhaImposto = linhaBase + j + 1
            
            ' Lê o tipo de imposto diretamente da planilha
            Dim tipoImposto As String
            Dim colImpostos As Long
            colImpostos = 2 ' Coluna B
            tipoImposto = Trim(wsImpostos.Cells(linhaImposto, colImpostos).Value)

            ' Captura os valores, mesmo que sejam zero ou vazios
            Dim baseImposto As String
            Dim aliquotaImposto As String
            Dim outrasBasesImposto As String

            baseImposto = Format(Nz(wsImpostos.Cells(linhaImposto, colBase).Value), "0.00")
            aliquotaImposto = Format(Nz(wsImpostos.Cells(linhaImposto, colAliquota).Value), "0.00")
            outrasBasesImposto = Format(Nz(wsImpostos.Cells(linhaImposto, colOutrasBases).Value), "0.00")

            ' Preenche no SAP
            session.findById("wnd[0]/usr/tabsITEM_TAB/tabpTAX/ssubITEM_TABS:SAPLJ1BB2:3200/tblSAPLJ1BB2TAX_CONTROL/ctxtJ_1BDYSTX-TAXTYP[0," & j & "]").Text = tipoImposto
            session.findById("wnd[0]/usr/tabsITEM_TAB/tabpTAX/ssubITEM_TABS:SAPLJ1BB2:3200/tblSAPLJ1BB2TAX_CONTROL/txtJ_1BDYSTX-BASE[3," & j & "]").Text = baseImposto
            session.findById("wnd[0]/usr/tabsITEM_TAB/tabpTAX/ssubITEM_TABS:SAPLJ1BB2:3200/tblSAPLJ1BB2TAX_CONTROL/txtJ_1BDYSTX-RATE[4," & j & "]").Text = aliquotaImposto
            session.findById("wnd[0]/usr/tabsITEM_TAB/tabpTAX/ssubITEM_TABS:SAPLJ1BB2:3200/tblSAPLJ1BB2TAX_CONTROL/txtJ_1BDYSTX-OTHBAS[7," & j & "]").Text = outrasBasesImposto
        Next j
        
        ' Calcula os impostos
        session.findById("wnd[0]/usr/tabsITEM_TAB/tabpTAX/ssubITEM_TABS:SAPLJ1BB2:3200/btnPB_SELECT_ALL").press
        session.findById("wnd[0]/usr/tabsITEM_TAB/tabpTAX/ssubITEM_TABS:SAPLJ1BB2:3200/btnPB_CALCULATOR").press
        session.findById("wnd[0]/tbar[0]/btn[3]").press

        ' Avança para o próximo bloco
        linhaBase = linhaBase + 7
        i = i + 1
    Else
        linhaBase = linhaBase + 1

    End If

    Loop
    
    ' Acessar aba de mensagens
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB4").Select

    ' Preencher mensagens do processo
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB4/ssubHEADER_TAB:SAPLJ1BB2:2400/tblSAPLJ1BB2MESSAGE_CONTROL/txtJ_1BDYFTX-MESSAGE[0,0]").Text = ws.Range("A2134").Value
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB4/ssubHEADER_TAB:SAPLJ1BB2:2400/tblSAPLJ1BB2MESSAGE_CONTROL/txtJ_1BDYFTX-MESSAGE[0,1]").Text = ws.Range("A2135").Value
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB4/ssubHEADER_TAB:SAPLJ1BB2:2400/tblSAPLJ1BB2MESSAGE_CONTROL/txtJ_1BDYFTX-MESSAGE[0,2]").Text = ws.Range("A2136").Value
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB4/ssubHEADER_TAB:SAPLJ1BB2:2400/tblSAPLJ1BB2MESSAGE_CONTROL/txtJ_1BDYFTX-MESSAGE[0,3]").Text = ws.Range("A2137").Value
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB4/ssubHEADER_TAB:SAPLJ1BB2:2400/tblSAPLJ1BB2MESSAGE_CONTROL/txtJ_1BDYFTX-MESSAGE[0,4]").Text = ws.Range("A2138").Value
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB4/ssubHEADER_TAB:SAPLJ1BB2:2400/tblSAPLJ1BB2MESSAGE_CONTROL/txtJ_1BDYFTX-MESSAGE[0,5]").Text = ws.Range("A2139").Value
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB4/ssubHEADER_TAB:SAPLJ1BB2:2400/tblSAPLJ1BB2MESSAGE_CONTROL/txtJ_1BDYFTX-MESSAGE[0,6]").Text = ws.Range("A2140").Value
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB4/ssubHEADER_TAB:SAPLJ1BB2:2400/tblSAPLJ1BB2MESSAGE_CONTROL/txtJ_1BDYFTX-MESSAGE[0,7]").Text = ws.Range("A2141").Value
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB4/ssubHEADER_TAB:SAPLJ1BB2:2400/tblSAPLJ1BB2MESSAGE_CONTROL/txtJ_1BDYFTX-MESSAGE[0,8]").Text = ws.Range("A2142").Value
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB4/ssubHEADER_TAB:SAPLJ1BB2:2400/tblSAPLJ1BB2MESSAGE_CONTROL/txtJ_1BDYFTX-MESSAGE[0,9]").Text = ws.Range("A2143").Value
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB4/ssubHEADER_TAB:SAPLJ1BB2:2400/tblSAPLJ1BB2MESSAGE_CONTROL/txtJ_1BDYFTX-MESSAGE[0,10]").Text = ws.Range("A2144").Value
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB4/ssubHEADER_TAB:SAPLJ1BB2:2400/tblSAPLJ1BB2MESSAGE_CONTROL/txtJ_1BDYFTX-MESSAGE[0,11]").Text = ws.Range("A2145").Value
       'Pressiona Enter até o aviso amarelo desaparecer
    Do
        session.findById("wnd[0]").sendVKey 0
        If session.findById("wnd[0]/sbar").MessageType <> "W" Then
            Exit Do
        End If
    Loop
    ' === TRANSPORTE ===
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB3").Select
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB3/ssubHEADER_TAB:SAPLJ1BB2:2300/tblSAPLJ1BB2PARTNER_CONTROL/cmbJ_1BDYNAD-PARVW[0,1]").Key = "SP"
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB3/ssubHEADER_TAB:SAPLJ1BB2:2300/tblSAPLJ1BB2PARTNER_CONTROL/ctxtJ_1BDYNAD-PARID[1,1]").Text = ws.Range("H7").Value
    session.findById("wnd[0]").sendVKey 0
    
     ' === PESO E INCOTERMS ===
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB5").Select

Dim incoterm1 As String
Dim incoterm2 As String

incoterm1 = ws.Range("B14").Value
incoterm2 = ws.Range("C14").Value

' Condição para alterar os valores
If incoterm1 = "FCA" And incoterm2 = "FREE CARRIER" Then
    incoterm1 = "EXW"
    incoterm2 = "EX WORKS"
End If

' Atribuindo os valores (com a condição aplicada)
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB5/ssubHEADER_TAB:SAPLJ1BB2:2500/ctxtJ_1BDYDOC-INCO1").Text = incoterm1
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB5/ssubHEADER_TAB:SAPLJ1BB2:2500/txtJ_1BDYDOC-INCO2").Text = incoterm2
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB5/ssubHEADER_TAB:SAPLJ1BB2:2500/txtJ_1BDYDOC-ANZPK").Text = ws.Range("H10").Value
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB5/ssubHEADER_TAB:SAPLJ1BB2:2500/txtJ_1BDYDOC-NTGEW").Text = ws.Range("B13").Value
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB5/ssubHEADER_TAB:SAPLJ1BB2:2500/txtJ_1BDYDOC-BRGEW").Text = ws.Range("B12").Value
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB5/ssubHEADER_TAB:SAPLJ1BB2:2500/ctxtJ_1BDYDOC-GEWEI").Text = "KG"

session.findById("wnd[0]").sendVKey 0
        
    ' === PAGAMENTO ===
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB6").Select
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB6/ssubHEADER_TAB:SAPLJ1BB2:2600/ctxtJ_1BDYDOC-ZTERM").Text = "0000"
    session.findById("wnd[0]").sendVKey 0
    
    ' === DOCS IMPORTAÇÃO ===
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB9").Select
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB9/ssubHEADER_TAB:SAPLJ1BB2:2900/subIMPORT_SUBDI:SAPLJ1BB2:2901/tblSAPLJ1BB2IMPORT_DI_CONTROL/txtJ_1BDYIMPORT_DI-NDI[0,0]").Text = ws.Range("H5").Value
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB9/ssubHEADER_TAB:SAPLJ1BB2:2900/subIMPORT_SUBDI:SAPLJ1BB2:2901/tblSAPLJ1BB2IMPORT_DI_CONTROL/ctxtJ_1BDYIMPORT_DI-DDI[1,0]").Text = ws.Range("H6").Value
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB9/ssubHEADER_TAB:SAPLJ1BB2:2900/subIMPORT_SUBDI:SAPLJ1BB2:2901/tblSAPLJ1BB2IMPORT_DI_CONTROL/txtJ_1BDYIMPORT_DI-XLOCDESEMB[2,0]").Text = ws.Range("H3").Value
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB9/ssubHEADER_TAB:SAPLJ1BB2:2900/subIMPORT_SUBDI:SAPLJ1BB2:2901/tblSAPLJ1BB2IMPORT_DI_CONTROL/ctxtJ_1BDYIMPORT_DI-UFDESEMB[3,0]").Text = ws.Range("H4").Value
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB9/ssubHEADER_TAB:SAPLJ1BB2:2900/subIMPORT_SUBDI:SAPLJ1BB2:2901/tblSAPLJ1BB2IMPORT_DI_CONTROL/ctxtJ_1BDYIMPORT_DI-DDESEMB[4,0]").Text = ws.Range("H6").Value
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB9/ssubHEADER_TAB:SAPLJ1BB2:2900/subIMPORT_SUBDI:SAPLJ1BB2:2901/tblSAPLJ1BB2IMPORT_DI_CONTROL/txtJ_1BDYIMPORT_DI-CEXPORTADOR[5,0]").Text = ws.Range("H13").Value
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB9/ssubHEADER_TAB:SAPLJ1BB2:2900/subIMPORT_SUBDI:SAPLJ1BB2:2901/tblSAPLJ1BB2IMPORT_DI_CONTROL/ctxtJ_1BDYIMPORT_DI-COD_DOC_IMP[6,0]").Text = "0"
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB9/ssubHEADER_TAB:SAPLJ1BB2:2900/subIMPORT_SUBDI:SAPLJ1BB2:2901/tblSAPLJ1BB2IMPORT_DI_CONTROL/ctxtJ_1BDYIMPORT_DI-TRANSPORT_MODE[8,0]").Text = ws.Range("M728").Value
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB9/ssubHEADER_TAB:SAPLJ1BB2:2900/subIMPORT_SUBDI:SAPLJ1BB2:2901/tblSAPLJ1BB2IMPORT_DI_CONTROL/ctxtJ_1BDYIMPORT_DI-INTERMEDIATE_MODE[11,0]").Text = "1"
    session.findById("wnd[0]").sendVKey 0
    Dim linhaADI As Integer
    Dim indiceADI As Integer
    linhaADI = 728
    indiceADI = 0
    
    For linha = 728 To 1427
        If ws.Cells(linha, "J").Value <> 0 Then
            session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB9/ssubHEADER_TAB:SAPLJ1BB2:2900/subIMPORT_SUBADI:SAPLJ1BB2:2902/tblSAPLJ1BB2IMPORT_ADI_CONTROL/ctxtJ_1BDYIMPORT_ADI-NDI[4," & indiceADI & "]").Text = ws.Range("H5").Value
            session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB9/ssubHEADER_TAB:SAPLJ1BB2:2900/subIMPORT_SUBADI:SAPLJ1BB2:2902/tblSAPLJ1BB2IMPORT_ADI_CONTROL/txtJ_1BDYIMPORT_ADI-NADICAO[5," & indiceADI & "]").Text = "1"
            session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB9/ssubHEADER_TAB:SAPLJ1BB2:2900/subIMPORT_SUBADI:SAPLJ1BB2:2902/tblSAPLJ1BB2IMPORT_ADI_CONTROL/txtJ_1BDYIMPORT_ADI-NSEQADIC[6," & indiceADI & "]").Text = CStr((indiceADI + 1) * 10)
            indiceADI = indiceADI + 1
        End If
    Next linha
    
    For i = 0 To indiceADI - 1
        session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB9/ssubHEADER_TAB:SAPLJ1BB2:2900/subIMPORT_SUBADI:SAPLJ1BB2:2902/tblSAPLJ1BB2IMPORT_ADI_CONTROL/txtJ_1BDYIMPORT_ADI-CFABRICANTE[7," & i & "]").Text = ws.Range("H13").Value
    Next i
    
    Dim valorTotalNota As Double
    Dim valorICM1 As Double
    Dim diferenca As Double
    Dim valorICM2 As Double
    
    ' Acessa as abas necessárias
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB4").Select
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB2").Select

    ' Captura o valor total da nota
    valorTotalNota = CDbl(Replace(session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB2/ssubHEADER_TAB:SAPLJ1BB2:2200/txtJ_1BDYDOC-NFTOT").Text, ".", ""))

    ' Captura o montante básico do ICM1
    valorICM1 = CDbl(Replace(session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB2/ssubHEADER_TAB:SAPLJ1BB2:2200/tblSAPLJ1BB2TOTAL_CONTROL/txtJ_1BDYTAX-BASE[3,1]").Text, ".", ""))
    
    ' Captura o moentante básico do ICM2
    valorICM2 = CDbl(Replace(session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB2/ssubHEADER_TAB:SAPLJ1BB2:2200/tblSAPLJ1BB2TOTAL_CONTROL/txtJ_1BDYTAX-OTHBAS[6,1]").Text, ".", ""))
    
    Dim cfopFinal As String
    Dim ultimaLinhaComItem As Long
    
    ' Encontrar a última linha com item válido
    For linha = 2131 To 1432 Step -1
        If Trim(ws.Cells(linha, "B").Value) <> "" And Trim(ws.Cells(linha, "B").Value) <> "0" Then
            ultimaLinhaComItem = linha
            Exit For
        End If
    Next linha
    
    cfopFinal = Trim(ws.Cells(ultimaLinhaComItem, "O").Value) ' Coluna O=CFOP
    
    If cfopFinal = "3556/AA" Or cfopFinal = "3949/AA" Then
        diferenca = Round(valorTotalNota - valorICM2, 2)
    Else
        diferenca = Round(valorTotalNota - valorICM1, 2)
    End If
    
    If diferenca <> 0 Then
        MsgBox "Gentileza verificar valores. Diferença de R$ " & Format(diferenca, "0.00"), vbExclamation
    End If
    
   ' Mensagem adicional sempre exibida
        MsgBox "DADOS IMPORTADOS COM SUCESSO, GENTILEZA VALIDAR TRANSPORTADORA, VENCIMENTO E CONTAINER.", vbInformation

End Sub

---------Módulo 4----------
Sub SubstituirReferenciasO()
Dim ws As Worksheet
    Dim cel As Range
    Dim i As Long
    Dim oldRef As String, newRef As String
    
    Set ws = ThisWorkbook.Sheets("IMPOSTOS")
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Loop de O2046 até O2131(86 linhas)
    For i = 0 To 85
        oldRef = "O" & (2046 + i)
        newRef = "O" & (1440 + i)

        ' Substitui em todas as células da planilha
        ws.Cells.Replace What:=oldRef, Replacement:=newRef, LookAt:=xlPart, _
                         SearchOrder:=xlByRows, MatchCase:=False

    Next i
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "Substituições concluídas com sucesso!", vbInformation
End Sub


----------Módulo 5------------
Sub VerificarDI()
    Dim xmlFile As String
    Dim xmlDoc As Object
    Dim cidadeCompleta As String
    Dim cidade As String
    Dim uf As String
    Dim dataRaw As String
    Dim dia As String, mes As String, ano As String
    Dim infoComplementar As String

    xmlFile = Application.GetOpenFilename("Arquivos XML (*.xml), *.xml", , "Selecione o XML da DI")
    If xmlFile = "Falso" Then Exit Sub

    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    xmlDoc.async = False
    xmlDoc.Load xmlFile
    If xmlDoc.ParseError.ErrorCode <> 0 Then
        MsgBox "Erro ao carregar o XML: " & xmlDoc.ParseError.Reason
        Exit Sub
    End If

    With ThisWorkbook.Sheets(1)
        On Error Resume Next

        .Range("H9").Value = GetTagValue(xmlDoc, "nomeEmbalagem")
        .Range("H10").Value = GetTagValue(xmlDoc, "quantidadeVolume")
        .Range("H5").Value = GetTagValue(xmlDoc, "numeroDI")

        dataRaw = GetTagValue(xmlDoc, "dataRegistro")
        If Len(dataRaw) = 8 Then
            ano = Left(dataRaw, 4)
            mes = Mid(dataRaw, 5, 2)
            dia = Right(dataRaw, 2)
            .Range("H6").Value = dia & "." & mes & "." & ano
        Else
            .Range("H6").Value = "Data inválida"
        End If

        cidadeCompleta = GetTagValue(xmlDoc, "cargaUrfEntradaNome")
        If Len(cidadeCompleta) > 0 Then
            .Range("H3").Value = ObterCidadeFormatada(cidadeCompleta)
            .Range("H4").Value = ObterUF(cidadeCompleta)
        Else
            .Range("H3").Value = "Cidade não encontrada"
            .Range("H4").Value = "??"
        End If
                
        infoComplementar = GetTagValue(xmlDoc, "informacaoComplementar")
        .Range("D2").Value = ExtrairTaxaVMLE(infoComplementar)
        .Range("D3").Value = ExtrairTaxaDolar(infoComplementar)
        .Range("Q6").Value = ConverterParaNumero(ExtrairValorPorTexto(infoComplementar, "VALOR DA TX SISCOMEX"))
        .Range("Q13").Value = ConverterParaNumero(ExtrairValorPorTexto(infoComplementar, "VALOR DO ICMS RECOLHIDO"))
        .Range("Q16").Value = ExtrairValorBaseICMS(infoComplementar)
        .Range("Q2:Q16").NumberFormat = "General"
        .Range("Q9").Value = ConverterParaNumero(ExtrairValorPorTexto(infoComplementar, "II RECOLHIDO"))
        .Range("Q10").Value = ConverterParaNumero(ExtrairValorPorTexto(infoComplementar, "IPI RECOLHIDO"))
        .Range("Q11").Value = ConverterParaNumero(ExtrairValorPorTexto(infoComplementar, "PIS RECOLHIDO"))
        .Range("Q12").Value = ConverterParaNumero(ExtrairValorPorTexto(infoComplementar, "COFINS RECOLHIDO"))
        .Range("Q2").Value = ConverterParaNumero(ExtrairValorPorTexto(infoComplementar, "VMLE"))
        .Range("Q3").Value = ConverterParaNumero(ExtrairValorPorTexto(infoComplementar, "FRETE"))
        .Range("Q4").Value = ConverterParaNumero(ExtrairValorPorTexto(infoComplementar, "SEGURO"))
        
        Dim afrmmValor As Double
        afrmmValor = ExtrairValorAFRMM(infoComplementar)
        If afrmmValor > 0 Then .Range("Q7").Value = afrmmValor
        
        Dim acrescimosValor As Double
        acrescimosValor = ExtrairValorAcrescimos(infoComplementar)
        If acrescimosValor > 0 Then .Range("Q5").Value = acrescimosValor
        
         Dim fobValor As Double
        fobValor = ExtrairValorFOB(infoComplementar)
        If fobValor > 0 Then .Range("Q2").Value = fobValor
        
        On Error GoTo 0
    End With

    MsgBox "Dados da DI importados com sucesso!"
End Sub

Function ConverterParaNumero(valorTexto As String) As Double
    Dim temp As String
    temp = Replace(temp, ",", ".") ' Troca vírgula por ponto
    temp = Replace(valorTexto, ".", "") ' Remove pontos
    
    If IsNumeric(temp) Then
        ConverterParaNumero = CDbl(temp)
    Else
        ConverterParaNumero = 0
    End If
End Function

Function GetTagValue(xmlDoc As Object, tagName As String) As String
    On Error Resume Next
    GetTagValue = xmlDoc.getElementsByTagName(tagName)(0).Text
End Function

Function ValorComCasasDecimais(xmlDoc As Object, tagName As String) As Double
    Dim rawValue As String
    rawValue = GetTagValue(xmlDoc, tagName)
    If IsNumeric(rawValue) Then
        ValorComCasasDecimais = CDbl(rawValue) / 100
    Else
        ValorComCasasDecimais = 0
    End If
End Function

Function ExtrairUltimaPalavra(texto As String) As String
    Dim partes() As String
    partes = Split(texto, "/")
    ExtrairUltimaPalavra = Trim(partes(UBound(partes)))
End Function

Function ObterUF(cidadeCompleta As String) As String
    Dim cidadeDetectada As String
    cidadeDetectada = UCase(cidadeCompleta)

    Select Case True
        Case InStr(cidadeDetectada, "VIRACOPOS") > 0
            ObterUF = "SP"
        Case InStr(cidadeDetectada, "GUARULHOS") > 0
            ObterUF = "SP"
        Case InStr(cidadeDetectada, "SANTOS") > 0
            ObterUF = "SP"
        Case InStr(cidadeDetectada, "CURITIBA") > 0
            ObterUF = "PR"
        Case InStr(cidadeDetectada, "SALVADOR") > 0
            ObterUF = "BA"
        Case InStr(cidadeDetectada, "SÃO PAULO") > 0
            ObterUF = "SP"
        Case InStr(cidadeDetectada, "RECIFE") > 0
            ObterUF = "PE"
        Case InStr(cidadeDetectada, "MANAUS") > 0
            ObterUF = "AM"
        Case InStr(cidadeDetectada, "PORTO ALEGRE") > 0
            ObterUF = "RS"
        Case Else
            ObterUF = "??"
    End Select
End Function

Function ExtrairTaxaVMLE(texto As String) As Double
    Dim linhas() As String
    Dim linha As String
    Dim i As Integer
    Dim taxaTexto As String
    Dim partes() As String

    linhas = Split(texto, vbLf)
    For i = LBound(linhas) To UBound(linhas)
        linha = Trim(linhas(i))
        If InStr(linha, "MOEDA VMLE") > 0 Then
            partes = Split(linha, " ")
            taxaTexto = partes(UBound(partes))
            If IsNumeric(taxaTexto) Then
                ExtrairTaxaVMLE = CDbl(taxaTexto)
                Exit Function
            End If
        End If
    Next i
    ExtrairTaxaVMLE = 0
End Function
Function ExtrairTaxaDolar(texto As String) As Double
    Dim linhas() As String
    Dim linha As String
    Dim i As Integer
    Dim taxaTexto As String
    Dim partes() As String

    linhas = Split(texto, vbLf)
    For i = LBound(linhas) To UBound(linhas)
        linha = Trim(linhas(i))
        If InStr(linha, "TAXA DOLAR") > 0 Then
            partes = Split(linha, " ")
            taxaTexto = partes(UBound(partes))
            If IsNumeric(taxaTexto) Then
                ExtrairTaxaDolar = CDbl(taxaTexto)
                Exit Function
            End If
        End If
    Next i
    ExtrairTaxaDolar = 0
End Function

Function ExtrairValorPorTexto(texto As String, chave As String) As String
    Dim linhas() As String
    Dim linha As String
    Dim i As Integer
    Dim pos As Long
    Dim valorTexto As String

    linhas = Split(texto, vbLf)
    For i = LBound(linhas) To UBound(linhas)
        linha = Trim(linhas(i))
        If InStr(1, linha, chave, vbTextCompare) > 0 Then
            ' Verifica se existe ": R$" ou "BRL"
            If InStr(linha, ": R$") > 0 Then
                pos = InStr(linha, ": R$")
                valorTexto = Trim(Mid(linha, pos + 4))
            ElseIf InStr(linha, "BRL") > 0 Then
                pos = InStr(linha, "BRL")
                valorTexto = Trim(Mid(linha, pos + 3))
            End If

            If Len(valorTexto) > 0 Then
                ExtrairValorPorTexto = valorTexto ' Retorna exatamente como está no XML
                Exit Function
            End If
        End If
    Next i
    ExtrairValorPorTexto = ""
End Function

Function ExtrairValorBaseICMS(texto As String) As Double
    ExtrairValorBaseICMS = ExtrairValorPorTexto(texto, "BASE DE CALCULO ICMS")
End Function

Function ExtrairValorAFRMM(texto As String) As Double
    ExtrairValorAFRMM = ExtrairValorPorTexto(texto, "CEMERCANTE")
End Function

Function ExtrairValorAcrescimos(texto As String) As Double
    ExtrairValorAcrescimos = ExtrairValorPorTexto(texto, "ACRESC")
End Function
Function ExtrairValorFOB(texto As String) As Double
    ExtrairValorFOB = ExtrairValorPorTexto(texto, "VMLE")
End Function
Function ExtrairValorFrete(texto As String) As Double
    ExtrairValorFrete = ExtrairValorPorTexto(texto, "FRETE")
End Function
Function ExtrairValorSeguro(texto As String) As Double
    ExtrairValorSeguro = ExtrairValorPorTexto(texto, "SEGURO")
End Function
Function ExtrairValorII(texto As String) As Double
    ExtrairValorII = ExtrairValorPorTexto(texto, "II RECOLHIDO")
End Function
Function ExtrairValorIPI(texto As String) As Double
    ExtrairValorIPI = ExtrairValorPorTexto(texto, "IPI RECOLHIDO")
End Function
Function ExtrairValorPIS(texto As String) As Double
    ExtrairValorPIS = ExtrairValorPorTexto(texto, "PIS RECOLHIDO")
End Function
Function ExtrairValorCOFINS(texto As String) As Double
    ExtrairValorCOFINS = ExtrairValorPorTexto(texto, "COFINS RECOLHIDO")
End Function
Function ExtrairValorICMS(texto As String) As Double
    ExtrairValorICMS = ExtrairValorPorTexto(texto, "ICMS RECOLHIDO")
End Function
Function ObterCidadeFormatada(cidadeCompleta As String) As String
    Dim cidadeDetectada As String
    cidadeDetectada = UCase(cidadeCompleta)

    Select Case True
        Case InStr(cidadeDetectada, "VIRACOPOS") > 0
            ObterCidadeFormatada = "CAMPINAS"
        Case InStr(cidadeDetectada, "GUARULHOS") > 0
            ObterCidadeFormatada = "GUARULHOS"
        Case InStr(cidadeDetectada, "SAO PAULO") > 0
            ObterCidadeFormatada = "SÃO PAULO"
    Case InStr(cidadeDetectada, "PORTO DE SANTOS") > 0
        ObterCidadeFormatada = "SANTOS"
        Case Else
            ObterCidadeFormatada = cidadeCompleta
    End Select
End Function




