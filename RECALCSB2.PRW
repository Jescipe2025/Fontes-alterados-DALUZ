#INCLUDE "TOTVS.CH"

User Function RECALCSB2()

	Local oDlg
	Local cProduto := Space(15)
	Local lTodos   := .F.
	Local oChk

	DEFINE MSDIALOG oDlg TITLE "Recalculo SB2" FROM 000,000 TO 170,420 PIXEL

	@ 015,010 SAY "Produto:" SIZE 050,009 PIXEL OF oDlg
	@ 015,065 MSGET cProduto SIZE 070,010 OF oDlg PIXEL PICTURE "@!" F3 "SB1"

	@ 035,010 CHECKBOX oChk VAR lTodos PROMPT "Processar todos" SIZE 150,012 OF oDlg PIXEL

	@ 050,030 BUTTON "Executar" SIZE 070,014 PIXEL OF oDlg ;
		ACTION (oDlg:End(), StartProc(cProduto, lTodos))

	@ 050,110 BUTTON "Sair" SIZE 060,014 PIXEL OF oDlg ;
		ACTION oDlg:End()

	ACTIVATE MSDIALOG oDlg CENTERED

Return


// =============================
// START
// =============================
Static Function StartProc(cProduto, lTodos)

	Local _cFilial
	Local lOk := .F.

	_cFilial := xFilial("SB2")
	cProduto := AllTrim(cProduto)

	If !lTodos .And. Empty(cProduto)
		If !MsgYesNo("Processar todos?", "Atencao")
			Return
		EndIf
		lTodos := .T.
	EndIf

	Processa({|| lOk := MontaTemp(cProduto, lTodos, _cFilial)}, ;
		"Aguarde", "Montando dados...", .F.)

	If !lOk
		MsgStop("Erro ao montar dados")
		Return
	EndIf

	If MsgYesNo("Deseja aplicar atualizacao?", "Confirmacao")

		Processa({|| SalvaBackupSB2(cProduto, lTodos, _cFilial)}, ;
			"Salva de Saldos", "Salvando backup SB2...", .T.)

		Processa({|| ExecUpdate(cProduto, lTodos, _cFilial)}, ;
			"Atualizando", "Gravando SB2...", .T.)

	EndIf

Return


// =============================
// TEMPORARIAS
// =============================
Static Function MontaTemp(cProduto, lTodos, _cFilial)

	Local cSQL  := ""
	Local nErro := 0

	TcSqlExec("IF OBJECT_ID('tempdb..#TMP_VENDAS') IS NOT NULL DROP TABLE #TMP_VENDAS")
	TcSqlExec("IF OBJECT_ID('tempdb..#TMP_LIB')    IS NOT NULL DROP TABLE #TMP_LIB")
	TcSqlExec("IF OBJECT_ID('tempdb..#TMP_COMP')   IS NOT NULL DROP TABLE #TMP_COMP")

	// SC6 - Pedidos de venda
	cSQL := " SELECT SC6.C6_PRODUTO PRODUTO, SC6.C6_LOCAL LOCAL, " + ;
		" SUM(SC6.C6_QTDVEN - SC6.C6_QTDENT) QPEDVEN, " + ;
		" SUM(CASE WHEN SC5.C5_LIBEROK='S' THEN SC6.C6_QTDVEN-SC6.C6_QTDENT ELSE 0 END) QEMP, " + ;
		" 0 AS QEMP2 " + ;
		" INTO #TMP_VENDAS " + ;
		" FROM " + RetSqlName("SC6") + " SC6 (NOLOCK) " + ;
		" JOIN " + RetSqlName("SC5") + " SC5 (NOLOCK) ON SC5.C5_NUM=SC6.C6_NUM " + ;
		" AND SC5.C5_FILIAL=SC6.C6_FILIAL AND SC5.D_E_L_E_T_=' ' " + ;
		" WHERE SC6.D_E_L_E_T_=' ' AND SC6.C6_FILIAL='" + _cFilial + "' "

	If !lTodos .And. !Empty(cProduto)
		cSQL += " AND SC6.C6_PRODUTO='" + cProduto + "' "
	EndIf

	cSQL += " GROUP BY SC6.C6_PRODUTO, SC6.C6_LOCAL "

	nErro := TcSqlExec(cSQL)
	If nErro < 0
		MsgStop("Erro na TMP_VENDAS (SC6): " + TcSqlError(), "Erro SQL")
		Return .F.
	EndIf

	// SC9 - Liberados
	cSQL := " SELECT C9_PRODUTO PRODUTO, C9_LOCAL LOCAL, SUM(C9_QTDLIB) QTDLIB " + ;
		" INTO #TMP_LIB FROM " + RetSqlName("SC9") + ;
		" WHERE D_E_L_E_T_=' ' AND C9_FILIAL='" + _cFilial + "' AND C9_NFISCAL=' ' "

	If !lTodos .And. !Empty(cProduto)
		cSQL += " AND C9_PRODUTO='" + cProduto + "' "
	EndIf

	cSQL += " GROUP BY C9_PRODUTO, C9_LOCAL "

	nErro := TcSqlExec(cSQL)
	If nErro < 0
		MsgStop("Erro na TMP_LIB (SC9): " + TcSqlError(), "Erro SQL")
		Return .F.
	EndIf

	// SC7 - Pedidos de compra
	cSQL := " SELECT C7_PRODUTO PRODUTO, C7_LOCAL LOCAL, " + ;
		" SUM(C7_QUANT-C7_QUJE) SALPEDI, " + ;
		" 0 AS SALPED2 " + ;
		" INTO #TMP_COMP FROM " + RetSqlName("SC7") + ;
		" WHERE D_E_L_E_T_=' ' AND C7_FILIAL='" + _cFilial + "' "

	If !lTodos .And. !Empty(cProduto)
		cSQL += " AND C7_PRODUTO='" + cProduto + "' "
	EndIf

	cSQL += " GROUP BY C7_PRODUTO, C7_LOCAL "

	nErro := TcSqlExec(cSQL)
	If nErro < 0
		MsgStop("Erro na TMP_COMP (SC7): " + TcSqlError(), "Erro SQL")
		Return .F.
	EndIf

Return .T.


// =============================
// UPDATE COM TRANSACAO
// =============================
Static Function ExecUpdate(cProduto, lTodos, _cFilial)

	Local cSQL  := ""
	Local lErro := .F.

	 

	 

		// ZERA campos que serao recalculados
		cSQL := " UPDATE " + RetSqlName("SB2") + " SET " + ;
			" B2_QEMP=0, B2_QEMP2=0, B2_QPEDVEN=0, " + ;
			" B2_RESERVA=0, B2_SALPEDI=0, B2_SALPED2=0 " + ;
			" WHERE B2_FILIAL='" + _cFilial + "' AND D_E_L_E_T_=' ' "

		If !lTodos .And. !Empty(cProduto)
			cSQL += " AND B2_COD='" + cProduto + "' "
		EndIf

	     TcSqlExec(cSQL) < 0
	 

		// ATUALIZA com os novos valores
		cSQL := " UPDATE SB2 SET " + ;
			" B2_QEMP    = ISNULL(V.QEMP,0), " + ;
			" B2_QEMP2   = 0, " + ;
			" B2_QPEDVEN = ISNULL(V.QPEDVEN,0), " + ;
			" B2_RESERVA = ISNULL(L.QTDLIB,0), " + ;
			" B2_SALPEDI = ISNULL(C.SALPEDI,0), " + ;
			" B2_SALPED2 = 0 " + ;
			" FROM " + RetSqlName("SB2") + " SB2 " + ;
			" LEFT JOIN #TMP_VENDAS V ON V.PRODUTO=B2_COD AND V.LOCAL=B2_LOCAL " + ;
			" LEFT JOIN #TMP_LIB    L ON L.PRODUTO=B2_COD AND L.LOCAL=B2_LOCAL " + ;
			" LEFT JOIN #TMP_COMP   C ON C.PRODUTO=B2_COD AND C.LOCAL=B2_LOCAL " + ;
			" WHERE B2_FILIAL='" + _cFilial + "' AND SB2.D_E_L_E_T_=' ' "

		If !lTodos .And. !Empty(cProduto)
			cSQL += " AND B2_COD='" + cProduto + "' "
		EndIf

		 TcSqlExec(cSQL)  
		MsgInfo("Recalculo SB2 concluido com sucesso!")
	 

Return


// =============================
// BACKUP DOS VALORES ANTERIORES
// =============================
Static Function SalvaBackupSB2(cProduto, lTodos, _cFilial)

	Local cSQL := ""

	cSQL := " UPDATE " + RetSqlName("SB2") + " SET " + ;
		" B2_XQEMP   = B2_QEMP,   " + ;
		" B2_XQEMP2  = B2_QEMP2,  " + ;
		" B2_XRESERV = B2_RESERVA, " + ;
		" B2_XRESER2 = B2_RESERV2, " + ;
		" B2_XSALPE  = B2_SALPEDI, " + ;
		" B2_XSALPE2 = B2_SALPED2, " + ;
		" B2_XQPEDV  = B2_QPEDVEN, " + ;
		" B2_XSALVA  = '" + dToS(Date()) + "' " + ;
		" FROM " + RetSqlName("SB2") + " SB2 " + ;
		" WHERE B2_FILIAL='" + _cFilial + "' AND SB2.D_E_L_E_T_=' ' " + ;
		" AND EXISTS ( " + ;
		"   SELECT 1 FROM #TMP_VENDAS V " + ;
		"   WHERE V.PRODUTO=B2_COD AND V.LOCAL=B2_LOCAL " + ;
		"   UNION ALL " + ;
		"   SELECT 1 FROM #TMP_LIB L " + ;
		"   WHERE L.PRODUTO=B2_COD AND L.LOCAL=B2_LOCAL " + ;
		"   UNION ALL " + ;
		"   SELECT 1 FROM #TMP_COMP C " + ;
		"   WHERE C.PRODUTO=B2_COD AND C.LOCAL=B2_LOCAL " + ;
		" ) "

	If !lTodos .And. !Empty(cProduto)
		cSQL += " AND B2_COD='" + cProduto + "' "
	EndIf

	// TcSqlExec(cSQL) // Descomentar em producao

Return
