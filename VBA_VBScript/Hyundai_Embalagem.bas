Sub LancarJ1B1N_SAP()
    Dim SapGuiAuto As Object
    Dim SAPApp As Object
    Dim SAPCon As Object
    Dim session As Object
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long

    ' Conectar ao SAP
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SAPApp = SapGuiAuto.GetScriptingEngine
    Set SAPCon = SAPApp.Children(0)
    Set session = SAPCon.Children(0)

    ' Definir a planilha
    Set ws = ThisWorkbook.Sheets("Planilha1") ' Altere para o nome correto da aba
    ultimaLinha = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    ' Loop pelas linhas da planilha
    For i = 2 To ultimaLinha
        Dim docSAP As String, nfNum As String, serie As String
        Dim dataDoc As String, chave9 As String, logNum As String, codAut As String

        docSAP = ws.Cells(i, 2).Value ' Coluna B
        nfNum = ws.Cells(i, 3).Value ' Coluna C
        serie = ws.Cells(i, 4).Value ' Coluna D
        dataDoc = ws.Cells(i, 7).Value ' Coluna G
        chave9 = ws.Cells(i, 10).Value ' Coluna J
        logNum = ws.Cells(i, 9).Value ' Coluna I

        ' Início do processo SAP
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/NJ1B1N"
        session.findById("wnd[0]").sendVKey 0

        session.findById("wnd[0]/usr/ctxtJ_1BDYDOC-NFTYPE").Text = "F1"
        session.findById("wnd[0]/usr/ctxtJ_1BDYDOC-BRANCH").Text = "centro"
        session.findById("wnd[0]/usr/cmbJ_1BDYDOC-PARVW").Key = "LF"
        session.findById("wnd[0]/usr/ctxtJ_1BDYDOC-PARID").Text = "Fornecedor"

        session.findById("wnd[0]/mbar/menu[0]/menu[5]").Select
        session.findById("wnd[1]/usr/ctxtJ_1BDYDOC-COP_DOCNUM").Text = docSAP
        session.findById("wnd[1]").sendVKey 0

        session.findById("wnd[0]/usr/subNF_NUMBER:SAPLJ1BB2:2002/txtJ_1BDYDOC-NFENUM").Text = nfNum
        session.findById("wnd[0]/usr/txtJ_1BDYDOC-SERIES").Text = serie
        session.findById("wnd[0]/usr/ctxtJ_1BDYDOC-DOCDAT").Text = dataDoc

        ' Detectar número de itens com material preenchido
        Dim linhaItem As Integer
        linhaItem = 0
        Do While True
            On Error Resume Next
            Dim matnr As String
            matnr = session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-MATNR[4," & linhaItem & "]").Text
            If Err.Number <> 0 Or matnr = "" Then Exit Do
            On Error GoTo 0

            ' Aplicar CFOP e Leis
            session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-CFOP[17," & linhaItem & "]").Text = "1920/AA"
            session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-TAXLW1[18," & linhaItem & "]").Text = "IC4"
            session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-TAXLW2[19," & linhaItem & "]").Text = "IP3"
            session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-TAXLW4[21," & linhaItem & "]").Text = "C98"
            session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-TAXLW5[22," & linhaItem & "]").Text = "P98"
            linhaItem = linhaItem + 1
        Loop

        session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB2").Select
        session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB8").Select
        session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB8/ssubHEADER_TAB:SAPLJ1BB2:2800/subRANDOM_NUMBER:SAPLJ1BB2:2801/txtJ_1BNFE_DOCNUM9_DIVIDED-DOCNUM8").Text = Left(chave9, 8)
        session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB8/ssubHEADER_TAB:SAPLJ1BB2:2800/subTIMESTAMP:SAPLJ1BB2:2803/ctxtJ_1BDYDOC-AUTHDATE").Text = dataDoc
        session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB8/ssubHEADER_TAB:SAPLJ1BB2:2800/subTIMESTAMP:SAPLJ1BB2:2803/subAUTHCODE_AREA:SAPLJ1BB2:2805/txtJ_1BDYDOC-AUTHCOD").Text = logNum
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/tbar[0]/btn[11]").press
        
    Next i

    MsgBox "Processo J1B1N concluído para todas as linhas!", vbInformation
End Sub
