#Include "PROTHEUS.ch"
#Include "TOPCONN.ch"

/*/{Protheus.doc} fnLimpaSC9
Rotina para remover reservas indevidas de itens de pedidos com entrega parcial
Calcula MI dinamicamente através da tabela SD3 direto na query principal
@type function
@author Adaptado
@since 05/02/2026
/*/
User Function fnLimpaSC9()
    Local aArea     := GetArea()
    Local cQuery    := ""
    Local nDeleted  := 0
    Local nAnalisados := 0
    Local cPedAnt   := ""
    Local aItens    := {}
    Local cMensagem := ""
    Local nI        := 0
    Local dDtIni    := CToD("")
    Local dDtFim    := CToD("")
    Local cPedIni   := ""
    Local cPedFim   := ""
    Local nMi       := 0
    Local nQtdAtendida := 0
    
    // Pergunta de confirmação
    If !MsgYesNo("Esta rotina irá remover reservas indevidas de itens parcialmente atendidos" + CRLF + ;
                 "em pedidos com entrega parcial." + CRLF + CRLF + ;
                 "Deseja continuar?", "ATENÇÃO")
        Return
    EndIf

    // Monta tela de parâmetros
    If !MontaPerguntas(@dDtIni, @dDtFim, @cPedIni, @cPedFim)
        MsgInfo("Operação cancelada.", "Aviso")
        RestArea(aArea)
        Return
    EndIf
    
    
    // Monta query CORRIGIDA com cálculo de MI direto na query principal
    cQuery := ""
    cQuery += " SELECT SC6.C6_NUM, SC6.C6_ITEM, SC6.C6_PRODUTO, " + CRLF
    cQuery += "        SC6.C6_QTDVEN, " + CRLF
    cQuery += "        SC6.C6_QTDEMP AS RES, " + CRLF  
    cQuery += "        SC6.C6_QTJADEV AS DEV, " + CRLF
    cQuery += "        SC6.C6_QTJACAN AS CAN, " + CRLF
    cQuery += "        SC6.C6_QTDENT AS NF, " + CRLF
    cQuery += "        ISNULL(TBMI.TOTAL_MI, 0) AS MI " + CRLF  // MI calculado
    cQuery += " FROM " + RetSqlName("SC6") + " SC6 (NOLOCK) " + CRLF
    cQuery += " INNER JOIN " + RetSqlName("SC5") + " SC5 (NOLOCK) " + CRLF
    cQuery += "      ON SC5.C5_NUM = SC6.C6_NUM " + CRLF
    cQuery += "     AND SC5.C5_FILIAL = SC6.C6_FILIAL " + CRLF
    cQuery += "     AND SC5.C5_EMISSAO BETWEEN '" + DToS(dDtIni) + "' AND '" + DToS(dDtFim) + "' " + CRLF   
    cQuery += "     AND SC5.D_E_L_E_T_ = ' ' " + CRLF
    cQuery += " INNER JOIN " + RetSqlName("SC9") + " SC9 (NOLOCK) " + CRLF
    cQuery += "      ON SC9.C9_FILIAL = SC6.C6_FILIAL " + CRLF
    cQuery += "     AND SC9.C9_PEDIDO = SC6.C6_NUM " + CRLF
    cQuery += "     AND SC9.C9_ITEM = SC6.C6_ITEM " + CRLF
    cQuery += "     AND SC9.C9_PRODUTO = SC6.C6_PRODUTO " + CRLF
    cQuery += "     AND SC9.C9_NFISCAL = ' ' " + CRLF  // Só reservas não faturadas
    cQuery += "     AND SC9.D_E_L_E_T_ = ' ' " + CRLF
    // LEFT JOIN com SD3 para calcular MI dinamicamente
    cQuery += " LEFT JOIN ( " + CRLF
    cQuery += "     SELECT LEFT(D3.D3_YPEDVEN, 6) AS PEDIDO, " + CRLF
    cQuery += "            RTRIM(LTRIM(RIGHT(D3.D3_YPEDVEN, 2))) AS ITEM, " + CRLF
    cQuery += "            D3.D3_COD AS PRODUTO, " + CRLF
    cQuery += "            SUM(D3.D3_QUANT) AS TOTAL_MI " + CRLF
    cQuery += "     FROM " + RetSqlName("SD3") + " D3 (NOLOCK) " + CRLF
    cQuery += "     WHERE D3.D_E_L_E_T_ = ' ' " + CRLF
    cQuery += "       AND D3.D3_FILIAL = '" + xFilial("SD3") + "' " + CRLF
    cQuery += "       AND D3.D3_ESTORNO <> 'S' " + CRLF
    cQuery += "       AND D3.D3_LOCAL <> '08' " + CRLF
    cQuery += "       AND D3.D3_YPEDVEN <> ' ' " + CRLF
    cQuery += "     GROUP BY LEFT(D3.D3_YPEDVEN, 6), " + CRLF
    cQuery += "              RTRIM(LTRIM(RIGHT(D3.D3_YPEDVEN, 2))), " + CRLF
    cQuery += "              D3.D3_COD " + CRLF
    cQuery += " ) TBMI ON TBMI.PEDIDO = SC6.C6_NUM " + CRLF
    cQuery += "       AND TBMI.ITEM = RTRIM(LTRIM(SC6.C6_ITEM)) " + CRLF
    cQuery += "       AND TBMI.PRODUTO = SC6.C6_PRODUTO " + CRLF
    cQuery += " WHERE SC6.C6_FILIAL = '" + xFilial("SC6") + "' " + CRLF 
    cQuery += "   AND SC6.C6_NUM BETWEEN '" + cPedIni + "' AND '" + cPedFim + "' " + CRLF   
    cQuery += "   AND SC6.C6_QTDEMP > 0 " + CRLF  // Tem reserva
    cQuery += "   AND SC6.D_E_L_E_T_ = ' ' " + CRLF
    // CRITÉRIO: Quantidade atendida (NF + MI - DEV - CAN) >= Quantidade vendida
    cQuery += "   AND (SC6.C6_QTDENT + ISNULL(TBMI.TOTAL_MI, 0) - SC6.C6_QTJADEV - SC6.C6_QTJACAN) >= SC6.C6_QTDVEN " + CRLF
    cQuery += " ORDER BY SC6.C6_NUM, SC6.C6_ITEM " + CRLF
    
    cQuery := ChangeQuery(cQuery)
    TcQuery cQuery Alias "TMPSC6" New
    
    DbSelectArea("TMPSC6")
    TMPSC6->(DbGoTop())
    
    // Verifica se encontrou registros
    If TMPSC6->(Eof())
        MsgInfo("Não foram encontrados itens com reservas indevidas.", "Aviso")
        TMPSC6->(DbCloseArea())
        RestArea(aArea)
        Return
    EndIf
    
    // Monta preview dos itens que serão processados
    cMensagem := "Itens que serão processados:" + CRLF + CRLF
    cMensagem += "Pedido Item  Produto          Qtd.Vend  Reserva  Qtd.Atend  MI" + CRLF
    cMensagem += Replicate("-", 75) + CRLF
    
    While !TMPSC6->(Eof())
        // MI já vem calculado da query
        nMi := TMPSC6->MI
        
        // Calcula quantidade atendida incluindo MI
        nQtdAtendida := TMPSC6->NF + nMi - TMPSC6->DEV - TMPSC6->CAN
        
        cMensagem += PadR(TMPSC6->C6_NUM, 7) + " "
        cMensagem += PadR(TMPSC6->C6_ITEM, 5) + " "
        cMensagem += PadR(TMPSC6->C6_PRODUTO, 17) + " "
        cMensagem += Transform(TMPSC6->C6_QTDVEN, "@E 999,999.99") + " "
        cMensagem += Transform(TMPSC6->RES, "@E 999,999.99") + " "
        cMensagem += Transform(nQtdAtendida, "@E 999,999.99") + " "
        cMensagem += Transform(nMi, "@E 999.99") + CRLF
        
        aAdd(aItens, {TMPSC6->C6_NUM, TMPSC6->C6_ITEM, TMPSC6->C6_PRODUTO})
        nAnalisados++
        
        TMPSC6->(DbSkip())
    EndDo
    
    TMPSC6->(DbCloseArea())
    
    If nAnalisados == 0
        MsgInfo("Não foram encontrados itens com reservas indevidas (considerando MI).", "Aviso")
        RestArea(aArea)
        Return
    EndIf
    
    cMensagem += CRLF + "Total: " + cValToChar(nAnalisados) + " item(ns)" + CRLF + CRLF
    cMensagem += "Confirma a remoção das reservas?"
    
    // Confirmação final
    If Aviso("Confirmação", cMensagem, {"Sim", "Não"}, 3) != 1
        MsgInfo("Operação cancelada.", "Cancelado")
        RestArea(aArea)
        Return
    EndIf
    
    // Processa a exclusão
    Processa({|| nDeleted := ProcessaExclusao(aItens) }, "Processando", "Removendo reservas...", .F.)
    
    MsgInfo("Processo finalizado!" + CRLF + CRLF + ;
            "Itens analisados: " + cValToChar(nAnalisados) + CRLF + ;
            "Itens removidos: " + cValToChar(nDeleted), "Resultado")
    
    RestArea(aArea)
Return

/*/{Protheus.doc} ProcessaExclusao
Processa a exclusão das reservas
@type function
/*/
Static Function ProcessaExclusao(aItens)
    Local nI        := 0
    Local nDeleted  := 0
    Local cUpdate   := ""
    
    ProcRegua(Len(aItens))
    
    For nI := 1 To Len(aItens)
        IncProc("Removendo: " + aItens[nI][1] + "-" + aItens[nI][2] + "-" + aItens[nI][3])
        
        // DELETE lógico do item específico no SC9
        cUpdate := " UPDATE " + RetSqlName("SC9") + CRLF
        cUpdate += " SET D_E_L_E_T_ = '*', " + CRLF
        cUpdate += "     R_E_C_D_E_L_ = R_E_C_N_O_ " + CRLF
        cUpdate += " WHERE C9_FILIAL = '" + xFilial("SC9") + "' " + CRLF
        cUpdate += "   AND C9_PEDIDO = '" + aItens[nI][1] + "' " + CRLF
        cUpdate += "   AND C9_ITEM = '" + aItens[nI][2] + "' " + CRLF
        cUpdate += "   AND C9_PRODUTO = '" + aItens[nI][3] + "' " + CRLF
        cUpdate += "   AND C9_NFISCAL = ' ' " + CRLF
        cUpdate += "   AND D_E_L_E_T_ = ' ' " + CRLF
        
        If TcSqlExec(cUpdate) < 0
            ConOut("ERRO ao deletar SC9: " + TcSqlError())
            ConOut("Pedido: " + aItens[nI][1] + " Item: " + aItens[nI][2] + " Produto: " + aItens[nI][3])
        Else
            nDeleted++
        EndIf
    Next nI
    
Return nDeleted

/*/{Protheus.doc} MontaPerguntas
Função para montar tela de parâmetros
@type function 
/*/
Static Function MontaPerguntas(dDtIni, dDtFim, cPedIni, cPedFim)
    Local lRet      := .F.
    Local oDlg
    Local oGetDtIni, oGetDtFim, oGetPedIni, oGetPedFim
    Local nOpcA     := 0
    
    // Valores padrão
    dDtIni  := FirstDate(Date())  // Primeiro dia do mês
    dDtFim  := Date()              // Hoje
    cPedIni := Space(6)
    cPedFim := Replicate("Z", 6)
    
    DEFINE MSDIALOG oDlg TITLE "Parâmetros - PV Entrega Parcial" FROM 000,000 TO 200,400 PIXEL
        
        @ 010,010 SAY "Data Emissão De:" SIZE 060,007 PIXEL OF oDlg
        @ 010,080 MSGET oGetDtIni VAR dDtIni SIZE 060,010 PIXEL OF oDlg VALID !Empty(dDtIni)
        
        @ 025,010 SAY "Data Emissão Até:" SIZE 060,007 PIXEL OF oDlg
        @ 025,080 MSGET oGetDtFim VAR dDtFim SIZE 060,010 PIXEL OF oDlg VALID !Empty(dDtFim) .And. dDtFim >= dDtIni
        
        @ 040,010 SAY "Pedido Venda De:" SIZE 060,007 PIXEL OF oDlg
        @ 040,080 MSGET oGetPedIni VAR cPedIni SIZE 060,010 PIXEL OF oDlg F3 "SC5"
        
        @ 055,010 SAY "Pedido Venda Até:" SIZE 060,007 PIXEL OF oDlg
        @ 055,080 MSGET oGetPedFim VAR cPedFim SIZE 060,010 PIXEL OF oDlg F3 "SC5"
        
        @ 080,080 BUTTON "Confirmar" SIZE 040,012 PIXEL OF oDlg ACTION (nOpcA := 1, oDlg:End())
        @ 080,130 BUTTON "Cancelar"  SIZE 040,012 PIXEL OF oDlg ACTION (nOpcA := 0, oDlg:End())
        
    ACTIVATE MSDIALOG oDlg CENTERED
    
    If nOpcA == 1
        If Empty(cPedIni)
            cPedIni := Space(6)
        EndIf
        If Empty(cPedFim)
            cPedFim := Replicate("Z", 6)
        EndIf
        lRet := .T.
    EndIf
    
Return lRet

/*/{Protheus.doc} TrMi
Calcula a Movimentação Interna (MI) dinamicamente através da tabela SD3
NOTA: Esta função foi mantida para compatibilidade, mas não é mais necessária
pois o MI agora é calculado direto na query principal
@type function
@param cPed - Número do Pedido
@param cItem - Item do Pedido  
@param cProd - Código do Produto
@return nQtdMi - Quantidade de Movimentação Interna
@author Adaptado de RPOSPEDVEN
@since 05/02/2026
/*/
Static Function TrMi(cPed, cItem, cProd)
    Local nQtdMi    := 0                      
    Local aAreaMi   := GetArea()
    Local cQueryMi  := ""
    
    // Query para buscar movimentações internas do pedido/item
    cQueryMi := " SELECT SUM(D3_QUANT) AS TOTAL_MI " + CRLF
    cQueryMi += " FROM " + RetSqlName("SD3") + " D3 (NOLOCK) " + CRLF
    cQueryMi += " WHERE D3.D_E_L_E_T_ <> '*' " + CRLF
    cQueryMi += "   AND D3.D3_FILIAL = '" + xFilial("SD3") + "' " + CRLF
    cQueryMi += "   AND LEFT(D3.D3_YPEDVEN, 6) = '" + cPed + "' " + CRLF
    cQueryMi += "   AND RTRIM(LTRIM(RIGHT(D3.D3_YPEDVEN, 2))) = RTRIM(LTRIM('" + cItem + "')) " + CRLF
    cQueryMi += "   AND D3.D3_COD = '" + cProd + "' " + CRLF
    cQueryMi += "   AND D3.D3_ESTORNO <> 'S' " + CRLF
    cQueryMi += "   AND D3.D3_LOCAL <> '08' " + CRLF
    
    cQueryMi := ChangeQuery(cQueryMi)
    TcQuery cQueryMi Alias "TMPMI" New
    
    DbSelectArea("TMPMI")
    TMPMI->(DbGoTop())
    
    If TMPMI->(!Eof()) .AND. !Empty(TMPMI->TOTAL_MI)
        nQtdMi := TMPMI->TOTAL_MI
    EndIf
    
    If Select("TMPMI") > 0
        TMPMI->(DbCloseArea())
    EndIf
    
    RestArea(aAreaMi)
    
Return nQtdMi
