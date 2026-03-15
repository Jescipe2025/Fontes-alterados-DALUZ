#Include "PROTHEUS.ch"
#Include "TOPCONN.ch"

/*/{Protheus.doc} FNDELSC9
Programa para exclusão lógica de liberações de pedidos totalmente entregues
@type function
@author Seu Nome
@since 03/02/2026
/*/
User Function FNDELSC9()
	Local aArea     := GetArea()
	Local cAlias    := GetNextAlias()
	Local cQuery    := ""
	Local nRecCount := 0
	Local aPedidos  := {}
	Local cMensagem := ""
	Local nOpc      := 0
	Local dDtIni    := CToD("")
    Local dDtFim    := CToD("")
    Local cPedIni   := ""
    Local cPedFim   := ""

    // Pergunta de confirmação
    If !MsgYesNo("Esta rotina irá remover reservas indevidas de itens totalmente atendidos" + CRLF + ;
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

	// Query para listar os pedidos que serão afetados
	cQuery := " SELECT DISTINCT " + CRLF
	cQuery += "     SC5.C5_EMISSAO, " + CRLF 
	cQuery += "     SC6.C6_NUM, " + CRLF
	cQuery += "     SC6.C6_ITEM, " + CRLF
	cQuery += "     SC6.C6_PRODUTO, " + CRLF
	cQuery += "     SC6.C6_QTDVEN, " + CRLF
	cQuery += "     SC6.C6_QTDENT, " + CRLF
	cQuery += "     SC6.C6_QTENTMI, " + CRLF
	cQuery += "     SC6.C6_QTJADEV, " + CRLF
	cQuery += "     SC6.C6_QTJACAN, " + CRLF
	cQuery += "     (SC6.C6_QTDENT + SC6.C6_QTENTMI - SC6.C6_QTJADEV - SC6.C6_QTJACAN) AS QTD_ENTREGUE " + CRLF
	cQuery += " FROM " + RetSqlName("SC6") + " AS SC6 " + CRLF
	cQuery += " INNER JOIN " + RetSqlName("SC5") + " SC5 " + CRLF
    cQuery += "      ON SC5.C5_NUM =  SC6.C6_NUM " + CRLF
    cQuery += "     AND SC5.C5_FILIAL =  SC6.C6_FILIAL " + CRLF
    cQuery += "     AND SC5.D_E_L_E_T_ = ' ' " + CRLF
	cQuery += "     AND SC5.C5_EMISSAO BETWEEN '" + DToS(dDtIni) + "' AND '" + DToS(dDtFim) + "' " + CRLF
	cQuery += " INNER JOIN " + RetSqlName("SC9") + " AS SC9 ON " + CRLF
	cQuery += "     SC9.C9_FILIAL = SC6.C6_FILIAL AND " + CRLF
	cQuery += "     SC9.C9_PEDIDO = SC6.C6_NUM AND " + CRLF
	cQuery += "     SC9.C9_ITEM = SC6.C6_ITEM AND " + CRLF
	cQuery += "     SC9.C9_PRODUTO = SC6.C6_PRODUTO AND " + CRLF
	cQuery += "     SC9.C9_NFISCAL = ' ' AND " + CRLF
	cQuery += "     SC9.D_E_L_E_T_ = ' ' " + CRLF
	cQuery += " WHERE SC6.D_E_L_E_T_ = ' ' " + CRLF	
    cQuery += "   AND SC6.C6_NUM BETWEEN '" + cPedIni + "' AND '" + cPedFim + "' " + CRLF   	
	cQuery += "   AND SC6.C6_FILIAL = '" + xFilial("SC6") + "' " + CRLF
	cQuery += "   AND SC6.C6_QTDVEN <= ((SC6.C6_QTDENT + SC6.C6_QTENTMI) - (SC6.C6_QTJADEV + SC6.C6_QTJACAN)) " + CRLF
	cQuery += " ORDER BY SC6.C6_NUM, SC6.C6_ITEM " + CRLF

	// Mostra processamento da consulta
	Processa({|| cQuery := ChangeQuery(cQuery), ;
		DbUseArea(.T., "TOPCONN", TCGenQry(,,cQuery), cAlias, .F., .T.) }, ;
		"Aguarde", "Consultando pedidos...", .F.)

	// Verifica se encontrou registros
	If (cAlias)->(Eof())
		MsgInfo("Não foram encontrados pedidos totalmente entregues para processar.", "Aviso")
		(cAlias)->(DbCloseArea())
		RestArea(aArea)
		Return
	EndIf

	// Monta array com os pedidos e mensagem
	cMensagem := "Pedidos que serão excluídos:" + CRLF + CRLF
	cMensagem += "Pedido    Item  Produto          Qtd.Vendida  Qtd.Entregue" + CRLF
	cMensagem += Replicate("-", 70) + CRLF

	While !(cAlias)->(Eof())
		nRecCount++

		cMensagem += PadR((cAlias)->C6_NUM, 10) + " "
		cMensagem += PadR((cAlias)->C6_ITEM, 6) + " "
		cMensagem += PadR((cAlias)->C6_PRODUTO, 17) + " "
		cMensagem += Transform((cAlias)->C6_QTDVEN, "@E 999,999.99") + "  "
		cMensagem += Transform((cAlias)->QTD_ENTREGUE, "@E 999,999.99") + CRLF

		aAdd(aPedidos, {(cAlias)->C6_NUM, (cAlias)->C6_ITEM, (cAlias)->C6_PRODUTO})

		(cAlias)->(DbSkip())
	EndDo

	(cAlias)->(DbCloseArea())

	cMensagem += CRLF + "Total de registros: " + cValToChar(nRecCount) + CRLF + CRLF
	cMensagem += "Deseja continuar com a exclusão?"

	// Pergunta ao usuário
	nOpc := Aviso("Confirmação", cMensagem, {"Sim", "Não"}, 3)

	If nOpc == 1  // Usuário confirmou

		Processa({|| ProcessaDelecao(nRecCount) }, "Processando", "Excluindo registros...", .F.)

	Else
		MsgInfo("Operação cancelada pelo usuário.", "Cancelado")
	EndIf

	RestArea(aArea)

Return

/*/{Protheus.doc} ProcessaDelecao
Função para processar a exclusão dos registros
@type function
@author Seu Nome
@since 03/02/2026
/*/
Static Function ProcessaDelecao(nRecCount)
	Local cQuery := ""
	Local lOk    := .T.

	ProcRegua(2)  // Define 2 etapas

	// Inicia transação
	Begin Transaction

		// Etapa 1: Marca filial como 'JJ' - NÃO MARCA MAIS EXLUIO DIRETO
		IncProc("Marcando registros (1/2)...")

		cQuery := " UPDATE " + RetSqlName("SC9") + CRLF
		//cQuery += " SET C9_FILIAL = 'JJ' " + CRLF
		cQuery += " SET D_E_L_E_T_ = '*' " + CRLF
		cQuery += " WHERE EXISTS ( " + CRLF
		cQuery += "     SELECT 1 " + CRLF
		cQuery += "     FROM " + RetSqlName("SC6") + " AS C6 " + CRLF
		cQuery += "     WHERE C6.C6_FILIAL = " + RetSqlName("SC9") + ".C9_FILIAL " + CRLF
		cQuery += "       AND C6.C6_NUM = " + RetSqlName("SC9") + ".C9_PEDIDO " + CRLF
		cQuery += "       AND C6.C6_ITEM = " + RetSqlName("SC9") + ".C9_ITEM " + CRLF
		cQuery += "       AND C6.C6_PRODUTO = " + RetSqlName("SC9") + ".C9_PRODUTO " + CRLF
		cQuery += "       AND C6.D_E_L_E_T_ = ' ' " + CRLF
		cQuery += "       AND C6.C6_FILIAL = '" + xFilial("SC6") + "' " + CRLF
		cQuery += "       AND C6.C6_QTDVEN <= ((C6.C6_QTDENT + C6.C6_QTENTMI) - (C6.C6_QTJADEV + C6.C6_QTJACAN)) " + CRLF
		cQuery += " ) " + CRLF
		cQuery += " AND C9_NFISCAL = ' '  " + CRLF
		cQuery += " AND " + RetSqlName("SC9") + ".D_E_L_E_T_ = ' ' " + CRLF

		If TCSqlExec(cQuery) < 0
			DisarmTransaction()
			MsgStop("Erro ao executar UPDATE 1: " + TCSQLError(), "Erro")
			lOk := .F.
			Return lOk
		EndIf
    /*
    // Etapa 2: Exclusão lógica
    IncProc("Excluindo registros (2/2)...")
    
    cQuery := " UPDATE " + RetSqlName("SC9") + CRLF
    cQuery += " SET D_E_L_E_T_ = '*', " + CRLF
    cQuery += "     R_E_C_D_E_L_ = R_E_C_N_O_ " + CRLF
    cQuery += " WHERE C9_FILIAL = 'JJ' " + CRLF
    
    If TCSqlExec(cQuery) < 0
        DisarmTransaction()
        MsgStop("Erro ao executar UPDATE 2: " + TCSQLError(), "Erro")
        lOk := .F.
        Return lOk
    EndIf
    */
	End Transaction

	If lOk
		MsgInfo("Processo concluído com sucesso!" + CRLF + ;
			"Total de registros processados: " + cValToChar(nRecCount), "Sucesso")
	EndIf

Return lOk


/*/{Protheus.doc} MontaPerguntas
Função para montar tela de parâmetros
@type function /*/

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
    
    DEFINE MSDIALOG oDlg TITLE "Parâmetros - PV Entrega Total" FROM 000,000 TO 200,400 PIXEL
        
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
