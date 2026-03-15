#INCLUDE "TOTVS.ch"
#INCLUDE "topconn.ch"

/*/{Protheus.doc} U_RECALCEMP
Rotina para recalcular o campo C6_QTDEMP (Quantidade Empenhada) 
considerando:
- SC9: Quantidade liberada real
- SD2: Quantidade já faturada
- SD3: Quantidade entregue por Movimento Interno

@type User Function
@author Adaptado
@since 28/01/2026
@version 1.0
/*/

User Function RECALCEMP()
    Local aArea      := GetArea()
    Local dDtIni     := CToD("")
    Local dDtFim     := CToD("")
    Local cPedIni    := ""
    Local cPedFim    := ""
    Local cProdIni   := ""
    Local cProdFim   := ""
    Local nProcessar := 1  // 1=Somente Divergentes, 2=Todos
    
    If !MsgYesNo("Esta rotina irá recalcular o campo C6_QTDEMP de todos os pedidos." + CRLF + ;
                 "Deseja continuar?", "Atenção")
        RestArea(aArea)
        Return
    EndIf
    
    // Monta tela de parâmetros
    If !MontaPerguntas(@dDtIni, @dDtFim, @cPedIni, @cPedFim, @cProdIni, @cProdFim, @nProcessar)
        MsgInfo("Operação cancelada.", "Aviso")
        RestArea(aArea)
        Return
    EndIf
    
    Processa({|| ProcRecalc(dDtIni, dDtFim, cPedIni, cPedFim, cProdIni, cProdFim, nProcessar)}, ;
             "Processando", "Recalculando C6_QTDEMP...", .T.)
    
    RestArea(aArea)
    
Return

/*/{Protheus.doc} MontaPerguntas
Função para montar tela de parâmetros
@type function
/*/
Static Function MontaPerguntas(dDtIni, dDtFim, cPedIni, cPedFim, cProdIni, cProdFim, nProcessar)
    Local lRet          := .F.
    Local oDlg
    Local oGetDtIni, oGetDtFim
    Local oGetPedIni, oGetPedFim
    Local oGetProdIni, oGetProdFim
    Local oCmbProcessar
    Local aProcessar    := {"Somente Divergentes", "Todos"}
    Local cProcessar    := aProcessar[1]
    Local nOpcA         := 0
    
    // Valores padrão
    dDtIni     := FirstDate(Date())  // Primeiro dia do mês
    dDtFim     := Date()             // Hoje
    cPedIni    := Space(6)
    cPedFim    := Replicate("Z", 6)
    cProdIni   := Space(15)
    cProdFim   := Replicate("Z", 15)
    nProcessar := 1
    
    DEFINE MSDIALOG oDlg TITLE "Parâmetros - Recálculo C6_QTDEMP" FROM 000,000 TO 280,400 PIXEL
        
        @ 010,010 SAY "Data Emissão De:" SIZE 060,007 PIXEL OF oDlg
        @ 010,080 MSGET oGetDtIni VAR dDtIni SIZE 060,010 PIXEL OF oDlg VALID !Empty(dDtIni)
        
        @ 025,010 SAY "Data Emissão Até:" SIZE 060,007 PIXEL OF oDlg
        @ 025,080 MSGET oGetDtFim VAR dDtFim SIZE 060,010 PIXEL OF oDlg VALID !Empty(dDtFim) .And. dDtFim >= dDtIni
        
        @ 040,010 SAY "Pedido Venda De:" SIZE 060,007 PIXEL OF oDlg
        @ 040,080 MSGET oGetPedIni VAR cPedIni SIZE 060,010 PIXEL OF oDlg F3 "SC5"
        
        @ 055,010 SAY "Pedido Venda Até:" SIZE 060,007 PIXEL OF oDlg
        @ 055,080 MSGET oGetPedFim VAR cPedFim SIZE 060,010 PIXEL OF oDlg F3 "SC5"
        
        @ 070,010 SAY "Produto De:" SIZE 060,007 PIXEL OF oDlg
        @ 070,080 MSGET oGetProdIni VAR cProdIni SIZE 060,010 PIXEL OF oDlg F3 "SB1"
        
        @ 085,010 SAY "Produto Até:" SIZE 060,007 PIXEL OF oDlg
        @ 085,080 MSGET oGetProdFim VAR cProdFim SIZE 060,010 PIXEL OF oDlg F3 "SB1"
        
        @ 100,010 SAY "Processar:" SIZE 060,007 PIXEL OF oDlg
        @ 100,080 MSCOMBOBOX oCmbProcessar VAR cProcessar ITEMS aProcessar SIZE 060,010 PIXEL OF oDlg
        
        @ 125,080 BUTTON "Confirmar" SIZE 040,012 PIXEL OF oDlg ACTION (nOpcA := 1, oDlg:End())
        @ 125,130 BUTTON "Cancelar"  SIZE 040,012 PIXEL OF oDlg ACTION (nOpcA := 0, oDlg:End())
        
    ACTIVATE MSDIALOG oDlg CENTERED
    
    If nOpcA == 1
        // Ajusta campos em branco
        If Empty(cPedIni)
            cPedIni := Space(6)
        EndIf
        If Empty(cPedFim)
            cPedFim := Replicate("Z", 6)
        EndIf
        If Empty(cProdIni)
            cProdIni := Space(15)
        EndIf
        If Empty(cProdFim)
            cProdFim := Replicate("Z", 15)
        EndIf
        
        // Define tipo de processamento
        nProcessar := aScan(aProcessar, cProcessar)
        
        lRet := .T.
    EndIf
    
Return lRet

/*/{Protheus.doc} ProcRecalc
Função principal que processa o recálculo
/*/
Static Function ProcRecalc(dDtIni, dDtFim, cPedIni, cPedFim, cProdIni, cProdFim, nProcessar)
    Local cQuery := ""
    Local nRecalc := 0
    Local nErros := 0
    Local aLog := {}
    
    // Query que busca todos os itens de pedido
    cQuery := " SELECT " + CRLF
    cQuery += "     SC6.R_E_C_N_O_ AS RECNOSC6, " + CRLF
    cQuery += "     SC6.C6_FILIAL, " + CRLF
    cQuery += "     SC6.C6_NUM AS PEDIDO, " + CRLF
    cQuery += "     SC6.C6_ITEM AS ITEM, " + CRLF
    cQuery += "     SC6.C6_PRODUTO AS PRODUTO, " + CRLF
    cQuery += "     SC6.C6_QTDVEN AS QTD_VENDIDA, " + CRLF
    cQuery += "     SC6.C6_QTDEMP AS QTD_EMP_ATUAL, " + CRLF
    cQuery += "     SC6.C6_QTDENT AS QTD_ENTREGUE, " + CRLF
    cQuery += " " + CRLF
    cQuery += "     -- Quantidade liberada na SC9 (real) " + CRLF
    cQuery += "     ISNULL(SC9_TOTAL.TOTAL_SC9, 0) AS QTD_LIBERADA_SC9, " + CRLF
    cQuery += " " + CRLF
    cQuery += "     -- Quantidade faturada em NF (SD2) " + CRLF
    cQuery += "     ISNULL(SD2_TOTAL.TOTAL_SD2, 0) AS QTD_FATURADA_SD2, " + CRLF
    cQuery += " " + CRLF
    cQuery += "     -- Quantidade entregue por MI (SD3) " + CRLF
    cQuery += "     ISNULL(SD3_TOTAL.TOTAL_SD3, 0) AS QTD_MI_SD3 " + CRLF
    cQuery += " " + CRLF
    cQuery += " FROM " + RetSqlName("SC6") + " SC6 (NOLOCK) " + CRLF
    cQuery += " " + CRLF
    cQuery += " INNER JOIN " + RetSqlName("SC5") + " SC5 (NOLOCK) ON " + CRLF
    cQuery += "     SC5.C5_FILIAL = SC6.C6_FILIAL " + CRLF
    cQuery += "     AND SC5.C5_NUM = SC6.C6_NUM " + CRLF
    cQuery += "     AND SC5.D_E_L_E_T_ = ' ' " + CRLF
    cQuery += " " + CRLF
    cQuery += " -- LEFT JOIN com SC9 (Liberações) " + CRLF
    cQuery += " LEFT JOIN ( " + CRLF
    cQuery += "     SELECT " + CRLF
    cQuery += "         C9_FILIAL, C9_PEDIDO, C9_ITEM, C9_PRODUTO, " + CRLF
    cQuery += "         SUM(C9_QTDLIB) AS TOTAL_SC9 " + CRLF
    cQuery += "     FROM " + RetSqlName("SC9") + " (NOLOCK) " + CRLF
    cQuery += "     WHERE D_E_L_E_T_ = ' ' " + CRLF
    cQuery += "       AND C9_NFISCAL = ' '  -- Ainda não faturado " + CRLF
    cQuery += "       AND C9_BLEST = ' '    -- Sem bloqueio de estoque " + CRLF
    cQuery += "       AND C9_BLCRED = ' '   -- Sem bloqueio de crédito " + CRLF
    cQuery += "     GROUP BY C9_FILIAL, C9_PEDIDO, C9_ITEM, C9_PRODUTO " + CRLF
    cQuery += " ) SC9_TOTAL ON " + CRLF
    cQuery += "     SC9_TOTAL.C9_FILIAL = SC6.C6_FILIAL " + CRLF
    cQuery += "     AND SC9_TOTAL.C9_PEDIDO = SC6.C6_NUM " + CRLF
    cQuery += "     AND SC9_TOTAL.C9_ITEM = SC6.C6_ITEM " + CRLF
    cQuery += "     AND SC9_TOTAL.C9_PRODUTO = SC6.C6_PRODUTO " + CRLF
    cQuery += " " + CRLF
    cQuery += " -- LEFT JOIN com SD2 (Itens de NF Saída) " + CRLF
    cQuery += " LEFT JOIN ( " + CRLF
    cQuery += "     SELECT " + CRLF
    cQuery += "         D2_FILIAL, D2_PEDIDO, D2_ITEMPV, D2_COD, " + CRLF
    cQuery += "         SUM(D2_QUANT) AS TOTAL_SD2 " + CRLF
    cQuery += "     FROM " + RetSqlName("SD2") + " (NOLOCK) " + CRLF
    cQuery += "     WHERE D_E_L_E_T_ = ' ' " + CRLF
    cQuery += "     GROUP BY D2_FILIAL, D2_PEDIDO, D2_ITEMPV, D2_COD " + CRLF
    cQuery += " ) SD2_TOTAL ON " + CRLF
    cQuery += "     SD2_TOTAL.D2_FILIAL = SC6.C6_FILIAL " + CRLF
    cQuery += "     AND SD2_TOTAL.D2_PEDIDO = SC6.C6_NUM " + CRLF
    cQuery += "     AND SD2_TOTAL.D2_ITEMPV = SC6.C6_ITEM " + CRLF
    cQuery += "     AND SD2_TOTAL.D2_COD = SC6.C6_PRODUTO " + CRLF
    cQuery += " " + CRLF
    cQuery += " -- LEFT JOIN com SD3 (Movimento Interno) " + CRLF
    cQuery += " LEFT JOIN ( " + CRLF
    cQuery += "     SELECT " + CRLF
    cQuery += "         D3_FILIAL, " + CRLF
    cQuery += "         LEFT(D3_YPEDVEN, 6) AS PEDIDO, " + CRLF
    cQuery += "         RTRIM(LTRIM(RIGHT(D3_YPEDVEN, 2))) AS ITEM, " + CRLF
    cQuery += "         D3_COD, " + CRLF
    cQuery += "         SUM(D3_QUANT) AS TOTAL_SD3 " + CRLF
    cQuery += "     FROM " + RetSqlName("SD3") + " (NOLOCK) " + CRLF
    cQuery += "     WHERE D_E_L_E_T_ = ' ' " + CRLF
    cQuery += "       AND D3_ESTORNO <> 'S' " + CRLF
    cQuery += "       AND D3_YPEDVEN <> '' " + CRLF
    cQuery += "       AND D3_LOCAL <> '08' " + CRLF
    cQuery += "     GROUP BY D3_FILIAL, LEFT(D3_YPEDVEN, 6), RTRIM(LTRIM(RIGHT(D3_YPEDVEN, 2))), D3_COD " + CRLF
    cQuery += " ) SD3_TOTAL ON " + CRLF
    cQuery += "     SD3_TOTAL.D3_FILIAL = SC6.C6_FILIAL " + CRLF
    cQuery += "     AND SD3_TOTAL.PEDIDO = SC6.C6_NUM " + CRLF
    cQuery += "     AND SD3_TOTAL.ITEM = SC6.C6_ITEM " + CRLF
    cQuery += "     AND SD3_TOTAL.D3_COD = SC6.C6_PRODUTO " + CRLF
    cQuery += " " + CRLF
    cQuery += " WHERE SC6.D_E_L_E_T_ = ' ' " + CRLF
    cQuery += "   AND SC6.C6_FILIAL = '" + xFilial("SC6") + "' " + CRLF
    cQuery += "   AND SC5.C5_EMISSAO BETWEEN '" + DToS(dDtIni) + "' AND '" + DToS(dDtFim) + "' " + CRLF
    cQuery += "   AND SC6.C6_NUM BETWEEN '" + cPedIni + "' AND '" + cPedFim + "' " + CRLF
    cQuery += "   AND SC6.C6_PRODUTO BETWEEN '" + cProdIni + "' AND '" + cProdFim + "' " + CRLF
    
    // Filtro opcional: apenas divergentes
    If nProcessar == 1  // 1=Somente Divergentes
        cQuery += "   AND SC6.C6_QTDEMP > 0 " + CRLF
    EndIf
    
    cQuery += " ORDER BY SC6.C6_FILIAL, SC6.C6_NUM, SC6.C6_ITEM " + CRLF
    
    // Grava query para debug
    MemoWrite("C:\temp\RECALCEMP.sql", cQuery)
    
    TCQUERY cQuery NEW ALIAS "TMPSC6"
    
    TCSetField("TMPSC6", "QTD_VENDIDA", "N", 12, 2)
    TCSetField("TMPSC6", "QTD_EMP_ATUAL", "N", 12, 2)
    TCSetField("TMPSC6", "QTD_ENTREGUE", "N", 12, 2)
    TCSetField("TMPSC6", "QTD_LIBERADA_SC9", "N", 12, 2)
    TCSetField("TMPSC6", "QTD_FATURADA_SD2", "N", 12, 2)
    TCSetField("TMPSC6", "QTD_MI_SD3", "N", 12, 2)
    
    Count To nTotal
    TMPSC6->(dbGoTop())
    
    ProcRegua(nTotal)
    
    dbSelectArea("SC6")
    SC6->(dbSetOrder(1))  // C6_FILIAL+C6_NUM+C6_ITEM
    
    While TMPSC6->(!EOF())
        
        IncProc("Processando: " + TMPSC6->PEDIDO + "/" + TMPSC6->ITEM)
        
        // Calcula o C6_QTDEMP correto
        nQtdEmpNova := CalcQtdEmp(TMPSC6->QTD_LIBERADA_SC9, ;
                                   TMPSC6->QTD_FATURADA_SD2, ;
                                   TMPSC6->QTD_MI_SD3)
        
        // Se houver diferença, atualiza
        If TMPSC6->QTD_EMP_ATUAL != nQtdEmpNova
            
            SC6->(dbGoto(TMPSC6->RECNOSC6))
            
            If RecLock("SC6", .F.)
                SC6->C6_QTDEMP := nQtdEmpNova
                SC6->(MsUnlock())
                
                nRecalc++
                
                // Log da alteração
                AADD(aLog, {TMPSC6->C6_FILIAL, ;
                           TMPSC6->PEDIDO, ;
                           TMPSC6->ITEM, ;
                           TMPSC6->PRODUTO, ;
                           TMPSC6->QTD_EMP_ATUAL, ;
                           nQtdEmpNova, ;
                           (TMPSC6->QTD_EMP_ATUAL - nQtdEmpNova), ;
                           "OK"})
                
                ConOut("[RECALCEMP] Atualizado: " + TMPSC6->PEDIDO + "/" + TMPSC6->ITEM + ;
                       " | Anterior: " + cValToChar(TMPSC6->QTD_EMP_ATUAL) + ;
                       " | Novo: " + cValToChar(nQtdEmpNova))
            Else
                nErros++
                AADD(aLog, {TMPSC6->C6_FILIAL, ;
                           TMPSC6->PEDIDO, ;
                           TMPSC6->ITEM, ;
                           TMPSC6->PRODUTO, ;
                           TMPSC6->QTD_EMP_ATUAL, ;
                           nQtdEmpNova, ;
                           (TMPSC6->QTD_EMP_ATUAL - nQtdEmpNova), ;
                           "ERRO - Não conseguiu travar registro"})
            EndIf
        EndIf
        
        TMPSC6->(dbSkip())
    EndDo
    
    TMPSC6->(dbCloseArea())
    
    // Gera relatório em Excel
    If Len(aLog) > 0
        GeraExcel(aLog, nRecalc, nErros)
    Else
        MsgInfo("Nenhum registro necessitou de atualização!", "Recálculo Concluído")
    EndIf
    
Return

/*/{Protheus.doc} CalcQtdEmp
Calcula a quantidade empenhada real
@param nQtdSC9 - Quantidade liberada na SC9
@param nQtdSD2 - Quantidade faturada em NF
@param nQtdSD3 - Quantidade entregue por MI
@return nQtdEmp - Quantidade empenhada real
/*/
Static Function CalcQtdEmp(nQtdSC9, nQtdSD2, nQtdSD3)
    Local nQtdEmp := 0
    
    // Lógica: 
    // C6_QTDEMP = Qtd na SC9 - Qtd faturada (SD2) - Qtd MI (SD3)
    // Mas não pode ser negativo
    
    nQtdEmp := nQtdSC9 - nQtdSD2 - nQtdSD3
    
    If nQtdEmp < 0
        nQtdEmp := 0
    EndIf
    
Return nQtdEmp

/*/{Protheus.doc} GeraExcel
Gera planilha Excel com log das alterações
/*/
Static Function GeraExcel(aLog, nRecalc, nErros)
    Local oFWMsExcel
    Local oExcel
    Local cArquivo := GetTempPath() + 'RecalcEmp_' + DTOS(Date()) + '_' + StrTran(Time(), ':', '') + '.xml'
    Local i
    
    oFWMsExcel := FWMSExcel():New()
    
    // Aba de log
    oFWMsExcel:AddworkSheet("Log de Alterações")
    oFWMsExcel:AddTable("Log de Alterações", "LOG")
    
    oFWMsExcel:AddColumn("Log de Alterações", "LOG", "Filial", 1)
    oFWMsExcel:AddColumn("Log de Alterações", "LOG", "Pedido", 1)
    oFWMsExcel:AddColumn("Log de Alterações", "LOG", "Item", 1)
    oFWMsExcel:AddColumn("Log de Alterações", "LOG", "Produto", 1)
    oFWMsExcel:AddColumn("Log de Alterações", "LOG", "C6_QTDEMP Anterior", 1)
    oFWMsExcel:AddColumn("Log de Alterações", "LOG", "C6_QTDEMP Novo", 1)
    oFWMsExcel:AddColumn("Log de Alterações", "LOG", "Diferença", 1)
    oFWMsExcel:AddColumn("Log de Alterações", "LOG", "Status", 1)
    
    For i := 1 To Len(aLog)
        oFWMsExcel:AddRow("Log de Alterações", "LOG", aLog[i])
    Next i
    
    // Linha de resumo
    oFWMsExcel:AddRow("Log de Alterações", "LOG", {"", "", "", "TOTAL:", nRecalc, "", "", ""})
    
    oFWMsExcel:Activate()
    oFWMsExcel:GetXMLFile(cArquivo)
    
    oExcel := MsExcel():New()
    oExcel:WorkBooks:Open(cArquivo)
    oExcel:SetVisible(.T.)
    oExcel:Destroy()
    
    MsgInfo("Recálculo concluído!" + CRLF + ;
            "Registros atualizados: " + cValToChar(nRecalc) + CRLF + ;
            "Erros: " + cValToChar(nErros) + CRLF + CRLF + ;
            "Planilha gerada: " + cArquivo, "Processamento Finalizado")
    
Return
