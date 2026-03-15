#Include "PROTHEUS.ch"
#Include "TOPCONN.ch"

/*/{Protheus.doc} FNVLPCXPV
Rotina para validar e limpar números de pedidos de compra inválidos no SC6
@type function
@author Seu Nome
@since 06/02/2026
/*/
User Function FNVLPCXPV()
    Local aArea     := GetArea()
    Local cQuery    := ""
    Local aRegistros := {}
    Local cMensagem := ""
    Local nI        := 0
    Local nLimpados := 0
    Local dDtIni    := CToD("")
    Local dDtFim    := CToD("")
    Local cPedIni   := ""
    Local cPedFim   := ""

    	// Pergunta de confirmação
	If !MsgYesNo("Esta rotina irá desvincular pedidos de vendas com pedidos de compras ja excluidos."+CRLF+CRLF+;
			"Deseja continuar?", "ATENÇÃO")
		Return
	EndIf
    
    // Monta tela de parâmetros
    If !MontaPerguntas(@dDtIni, @dDtFim, @cPedIni, @cPedFim)
        MsgInfo("Operação cancelada.", "Aviso")
        RestArea(aArea)
        Return
    EndIf
    
    // Query para buscar itens do SC6 com pedido de compra que não existe no SC7
    cQuery := ""
    cQuery += " SELECT " + CRLF
    cQuery += "     C6.C6_NUM, " + CRLF
    cQuery += "     C5.C5_EMISSAO, " + CRLF
    cQuery += "     C6.C6_ITEM, " + CRLF
    cQuery += "     C5.C5_CLIENTE, " + CRLF
    cQuery += "     C5.C5_LOJACLI, " + CRLF
    cQuery += "     A1.A1_NOME, " + CRLF
    cQuery += "     C6.C6_QTDVEN, " + CRLF
    cQuery += "     C6.C6_PRODUTO, " + CRLF
    cQuery += "     B1.B1_DESC, " + CRLF
    cQuery += "     C6.C6_YNUMPC, " + CRLF
    cQuery += "     C6.C6_ITEMPC, " + CRLF
    cQuery += "     C6.R_E_C_N_O_ AS RECNOSC6 " + CRLF
    cQuery += " FROM " + RetSqlName("SC6") + " C6 (NOLOCK) " + CRLF
    cQuery += " INNER JOIN " + RetSqlName("SC5") + " C5 (NOLOCK) " + CRLF
    cQuery += "     ON C5.C5_FILIAL = C6.C6_FILIAL " + CRLF
    cQuery += "    AND C5.C5_NUM = C6.C6_NUM " + CRLF
    cQuery += "    AND C5.D_E_L_E_T_ = ' ' " + CRLF
    cQuery += " INNER JOIN " + RetSqlName("SA1") + " A1 (NOLOCK) " + CRLF
    cQuery += "     ON A1.A1_FILIAL = '" + xFilial("SA1") + "' " + CRLF
    cQuery += "    AND A1.A1_COD = C5.C5_CLIENTE " + CRLF
    cQuery += "    AND A1.A1_LOJA = C5.C5_LOJACLI " + CRLF
    cQuery += "    AND A1.D_E_L_E_T_ = ' ' " + CRLF
    cQuery += " INNER JOIN " + RetSqlName("SB1") + " B1 (NOLOCK) " + CRLF
    cQuery += "     ON B1.B1_FILIAL = '" + xFilial("SB1") + "' " + CRLF
    cQuery += "    AND B1.B1_COD = C6.C6_PRODUTO " + CRLF
    cQuery += "    AND B1.D_E_L_E_T_ = ' ' " + CRLF
    cQuery += " WHERE C6.C6_FILIAL = '" + xFilial("SC6") + "' " + CRLF
    cQuery += "   AND C6.C6_YNUMPC <> '' " + CRLF  // Tem pedido de compra preenchido
    cQuery += "   AND C5.C5_EMISSAO BETWEEN '" + DToS(dDtIni) + "' AND '" + DToS(dDtFim) + "' " + CRLF
    cQuery += "   AND C6.C6_NUM BETWEEN '" + cPedIni + "' AND '" + cPedFim + "' " + CRLF
    cQuery += "   AND C6.D_E_L_E_T_ = ' ' " + CRLF
    // Verifica se o pedido de compra NÃO existe no SC7
    cQuery += "   AND NOT EXISTS ( " + CRLF
    cQuery += "       SELECT 1 " + CRLF
    cQuery += "       FROM " + RetSqlName("SC7") + " C7 (NOLOCK) " + CRLF
    cQuery += "       WHERE C7.C7_FILIAL = '" + xFilial("SC7") + "' " + CRLF
    cQuery += "         AND C7.C7_NUM = C6.C6_YNUMPC " + CRLF
    cQuery += "         AND C7.D_E_L_E_T_ = ' ' " + CRLF
    cQuery += "   ) " + CRLF
    cQuery += " ORDER BY C6.C6_NUM, C6.C6_ITEM " + CRLF
    
    Processa({|| cQuery := ChangeQuery(cQuery), ;
                 DbUseArea(.T., "TOPCONN", TCGenQry(,,cQuery), "TMPSC6", .F., .T.) }, ;
                 "Aguarde", "Consultando registros...", .F.)
    
    DbSelectArea("TMPSC6")
    TMPSC6->(DbGoTop())
    
    // Verifica se encontrou registros
    If TMPSC6->(Eof())
        MsgInfo("Não foram encontrados pedidos de venda com pedido de compra inválido.", "Aviso")
        TMPSC6->(DbCloseArea())
        RestArea(aArea)
        Return
    EndIf
    
    // Monta array e mensagem com os registros
    cMensagem := "Registros com Pedido de Compra INEXISTENTE que serão limpos:" + CRLF + CRLF
    cMensagem += "Emissão   Pedido Item Cliente Loja Nome Cliente          Qtd    Produto          Descrição               PC      ItemPC" + CRLF
    cMensagem += Replicate("-", 150) + CRLF
    
    While !TMPSC6->(Eof())
        cMensagem += DToC(SToD(TMPSC6->C5_EMISSAO)) + " "
        cMensagem += PadR(TMPSC6->C6_NUM, 7) + " "
        cMensagem += PadR(TMPSC6->C6_ITEM, 5) + " "
        cMensagem += PadR(TMPSC6->C5_CLIENTE, 8) + " "
        cMensagem += PadR(TMPSC6->C5_LOJACLI, 5) + " "
        cMensagem += PadR(TMPSC6->A1_NOME, 23) + " "
        cMensagem += Transform(TMPSC6->C6_QTDVEN, "@E 999.99") + " "
        cMensagem += PadR(TMPSC6->C6_PRODUTO, 17) + " "
        cMensagem += PadR(TMPSC6->B1_DESC, 24) + " "
        cMensagem += PadR(TMPSC6->C6_YNUMPC, 8) + " "
        cMensagem += TMPSC6->C6_ITEMPC + CRLF
        
        aAdd(aRegistros, TMPSC6->RECNOSC6)
        TMPSC6->(DbSkip())
    EndDo
    
    TMPSC6->(DbCloseArea())
    
    cMensagem += CRLF + "Total de registros: " + cValToChar(Len(aRegistros)) + CRLF + CRLF
    cMensagem += "Deseja prosseguir com a limpeza dos campos C6_YNUMPC e C6_ITEMPC?"
    
    // Pergunta ao usuário
    If Aviso("Confirmação", cMensagem, {"Sim", "Não"}, 3) == 1
        Processa({|| nLimpados := LimpaNumPC(aRegistros) }, "Processando", "Limpando registros...", .F.)
        
        MsgInfo("Processo finalizado!" + CRLF + CRLF + ;
                "Registros limpos: " + cValToChar(nLimpados), "Sucesso")
    Else
        MsgInfo("Operação cancelada pelo usuário.", "Cancelado")
    EndIf
    
    RestArea(aArea)
Return

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
    
    DEFINE MSDIALOG oDlg TITLE "Parâmetros - Validação PC x PV" FROM 000,000 TO 200,400 PIXEL
        
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
        // Ajusta campos em branco
        If Empty(cPedIni)
            cPedIni := Space(6)
        EndIf
        If Empty(cPedFim)
            cPedFim := Replicate("Z", 6)
        EndIf
        lRet := .T.
    EndIf
    
Return lRet

/*/{Protheus.doc} LimpaNumPC
Função para limpar os campos C6_YNUMPC e C6_ITEMPC
@type function
/*/
Static Function LimpaNumPC(aRegistros)
    Local nI        := 0
    Local nLimpados := 0
    
    ProcRegua(Len(aRegistros))
    
    DbSelectArea("SC6")
    
    Begin Transaction
    
    For nI := 1 To Len(aRegistros)
        IncProc("Limpando registro " + cValToChar(nI) + "/" + cValToChar(Len(aRegistros)))
        
        SC6->(DbGoTo(aRegistros[nI]))
        
        If !SC6->(Eof()) .And. !SC6->(Deleted())
            RecLock("SC6", .F.)
            SC6->C6_YNUMPC := ""
            SC6->C6_ITEMPC := ""
            SC6->(MsUnlock())
            nLimpados++
        EndIf
    Next nI
    
    End Transaction
    
Return nLimpados
