/*
???????????????????????????????????????????????????????????????????????????????
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±?????????????????????????????????????????????????????????????????????????±±
±±?Programa   ?MATR001    ? Autor ? SEU NOME AQUI            ?  30/01/2026 ?±±
±±?????????????????????????????????????????????????????????????????????????±±
±±?Descricao ? Relatorio de Inconsistencias SC6 x SC9                    ?±±
±±?          ? Identifica itens totalmente entregues com liberacao       ?±±
±±?          ? ativa no SC9 (o que nao deveria existir)                  ?±±
±±?????????????????????????????????????????????????????????????????????????±±
±±?Uso       ? Generico                                                  ?±±
±±?????????????????????????????????????????????????????????????????????????±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
???????????????????????????????????????????????????????????????????????????????
*/
#Include "protheus.ch"
#Include "rwmake.ch"
#Include "TopConn.ch"
#Include "Report.ch"

User Function fnsc9fant()
	Local cPerg := PadR("FNSC9FANT", Len(SX1->X1_GRUPO))
	Local oReport
	
	ValidPerg(cPerg)
	
	If Pergunte(cPerg, .T.)
		oReport := ReportDef()
		oReport:PrintDialog()
	EndIf
	
Return

Static Function ReportDef()
	Local oReport
	Local oSection1
	Local oSection2
	Local cTitulo := "Inconsistências SC6 x SC9"
	
	oReport := TReport():New("MATR001", cTitulo, "FNSC9FANT", {|oReport| PrintReport(oReport)}, "Relatório de Inconsistências SC6 x SC9")
	oReport:SetLandscape()
	oReport:SetTotalInLine(.F.)
	
	// Seção 1 - Detalhes (Analítico ou Sintético)
	oSection1 := TRSection():New(oReport, "Inconsistências", {"QRYPRO"})
	oSection1:SetTotalInLine(.F.)
	
	// Colunas do relatório ANALÍTICO (todas)
	TRCell():New(oSection1, "C6_FILIAL",   "QRYPRO", "Filial",        "@!",            04)
	TRCell():New(oSection1, "C6_NUM",      "QRYPRO", "Pedido",        "@!",            06)
	TRCell():New(oSection1, "C6_ITEM",     "QRYPRO", "Item",          "@!",            04)
	TRCell():New(oSection1, "C6_PRODUTO",  "QRYPRO", "Produto",       "@!",            15)
	TRCell():New(oSection1, "B1_DESC",     "QRYPRO", "Descricao",     "@!",            30)
	TRCell():New(oSection1, "C5_EMISSAO",  "QRYPRO", "Emissao",       "@!",            08)
	TRCell():New(oSection1, "C5_CLIENTE",  "QRYPRO", "Cliente",       "@!",            06)
	TRCell():New(oSection1, "C5_LOJACLI",  "QRYPRO", "Loja",          "@!",            02)
	TRCell():New(oSection1, "A1_NOME",     "QRYPRO", "Nome Cliente",  "@!",            30)
	TRCell():New(oSection1, "C6_QTDVEN",   "QRYPRO", "Qtd Vendida",   "@E 999,999.99", 12)
	TRCell():New(oSection1, "C6_QTDENT",   "QRYPRO", "Qtd Entregue",  "@E 999,999.99", 12)
	TRCell():New(oSection1, "C6_QTENTMI",  "QRYPRO", "Qtd Ent MI",    "@E 999,999.99", 12)
	TRCell():New(oSection1, "C6_QTJADEV",  "QRYPRO", "Qtd Ja Dev",    "@E 999,999.99", 12)
	TRCell():New(oSection1, "C6_QTJACAN",  "QRYPRO", "Qtd Ja Can",    "@E 999,999.99", 12)
	TRCell():New(oSection1, "C6_RESERVA",  "QRYPRO", "Reserva",       "@!",            10)
	TRCell():New(oSection1, "C9_DATALIB",  "QRYPRO", "Data Lib SC9",  "@!",            08)
	TRCell():New(oSection1, "C9_QTDLIB",   "QRYPRO", "Qtd Lib SC9",   "@E 999,999.99", 12)
	
	// Seção 2 - Totalizador
	oSection2 := TRSection():New(oReport, "Totalizador", {})
	oSection2:SetTotalInLine(.F.)
	oSection2:SetHeaderSection(.F.)
	
	TRCell():New(oSection2, "DESCRICAO",   "", "Descrição",      "@!",            80)
	TRCell():New(oSection2, "C6_QTDVEN",   "", "Qtd Vendida",    "@E 999,999.99", 12)
	TRCell():New(oSection2, "C6_QTDENT",   "", "Qtd Entregue",   "@E 999,999.99", 12)
	TRCell():New(oSection2, "C6_QTENTMI",  "", "Qtd Ent MI",     "@E 999,999.99", 12)
	TRCell():New(oSection2, "C6_QTJADEV",  "", "Qtd Ja Dev",     "@E 999,999.99", 12)
	TRCell():New(oSection2, "C6_QTJACAN",  "", "Qtd Ja Can",     "@E 999,999.99", 12)
	TRCell():New(oSection2, "C9_QTDLIB",   "", "Qtd Lib SC9",    "@E 999,999.99", 12)
	
Return oReport

Static Function PrintReport(oReport)
	Local oSection1 := oReport:Section(1)
	Local oSection2 := oReport:Section(2)
	Local cQuery    := ""
	Local aPedidos  := {}
	Local nTotQtdVen := 0
	Local nTotQtdEnt := 0
	Local nTotQtdEntMi := 0
	Local nTotQtdJaDev := 0
	Local nTotQtdJaCan := 0
	Local nTotQtdLib := 0
	Local lSintetico := (mv_par11 == 1)  // 1=Sintético, 2=Analítico
	
	// Monta a query baseada no tipo de relatório
	If lSintetico
		// Query para relatório SINTÉTICO (agrupado por pedido)
		cQuery := " SELECT " + CHR(13) + CHR(10)
		cQuery += "     C6.C6_FILIAL, " + CHR(13) + CHR(10)
		cQuery += "     C6.C6_NUM, " + CHR(13) + CHR(10)
		cQuery += "     C5.C5_EMISSAO, " + CHR(13) + CHR(10)
		cQuery += "     C5.C5_CLIENTE, " + CHR(13) + CHR(10)
		cQuery += "     C5.C5_LOJACLI, " + CHR(13) + CHR(10)
		cQuery += "     A1.A1_NOME, " + CHR(13) + CHR(10)
		cQuery += "     SUM(C6.C6_QTDVEN) AS C6_QTDVEN, " + CHR(13) + CHR(10)
		cQuery += "     SUM(C6.C6_QTDENT) AS C6_QTDENT " + CHR(13) + CHR(10)
		cQuery += " FROM " + RetSqlName("SC6") + " AS C6 " + CHR(13) + CHR(10)
	Else
		// Query para relatório ANALÍTICO (detalhado)
		cQuery := " SELECT DISTINCT " + CHR(13) + CHR(10)
		cQuery += "     C6.C6_FILIAL, " + CHR(13) + CHR(10)
		cQuery += "     C6.C6_NUM, " + CHR(13) + CHR(10)
		cQuery += "     C6.C6_ITEM, " + CHR(13) + CHR(10)
		cQuery += "     C6.C6_PRODUTO, " + CHR(13) + CHR(10)
		cQuery += "     C5.C5_EMISSAO, " + CHR(13) + CHR(10)
		cQuery += "     C6.C6_QTDVEN, " + CHR(13) + CHR(10)
		cQuery += "     C6.C6_QTDENT, " + CHR(13) + CHR(10)
		cQuery += "     C6.C6_QTENTMI, " + CHR(13) + CHR(10)
		cQuery += "     C6.C6_QTJADEV, " + CHR(13) + CHR(10)
		cQuery += "     C6.C6_QTJACAN, " + CHR(13) + CHR(10)
		cQuery += "     ISNULL(C6.C6_RESERVA, '') AS C6_RESERVA, " + CHR(13) + CHR(10)
		cQuery += "     C5.C5_CLIENTE, " + CHR(13) + CHR(10)
		cQuery += "     C5.C5_LOJACLI, " + CHR(13) + CHR(10)
		cQuery += "     A1.A1_NOME, " + CHR(13) + CHR(10)
		cQuery += "     B1.B1_DESC, " + CHR(13) + CHR(10)
		cQuery += "     C9.C9_DATALIB, " + CHR(13) + CHR(10)
		cQuery += "     C9.C9_QTDLIB " + CHR(13) + CHR(10)
		cQuery += " FROM " + RetSqlName("SC6") + " AS C6 " + CHR(13) + CHR(10)
	EndIf
	
	// JOINs comuns
	cQuery += " INNER JOIN " + RetSqlName("SC9") + " AS C9 ON " + CHR(13) + CHR(10)
	cQuery += "     C9.C9_FILIAL = C6.C6_FILIAL AND " + CHR(13) + CHR(10)
	cQuery += "     C9.C9_PEDIDO = C6.C6_NUM AND " + CHR(13) + CHR(10)
	cQuery += "     C9.C9_ITEM = C6.C6_ITEM AND " + CHR(13) + CHR(10)
	cQuery += "     C9.C9_PRODUTO = C6.C6_PRODUTO AND " + CHR(13) + CHR(10)
	cQuery += "     C9.D_E_L_E_T_ = ' ' " + CHR(13) + CHR(10)
	cQuery += " INNER JOIN " + RetSqlName("SC5") + " AS C5 ON " + CHR(13) + CHR(10)
	cQuery += "     C5.C5_FILIAL = C6.C6_FILIAL AND " + CHR(13) + CHR(10)
	cQuery += "     C5.C5_NUM = C6.C6_NUM AND " + CHR(13) + CHR(10)
	cQuery += "     C5.C5_EMISSAO BETWEEN '" + DTOS(mv_par01) + "' AND '" + DTOS(mv_par02) + "' AND " + CHR(13) + CHR(10)
	cQuery += "     C5.D_E_L_E_T_ = ' ' " + CHR(13) + CHR(10)
	cQuery += " LEFT JOIN " + RetSqlName("SA1") + " AS A1 ON " + CHR(13) + CHR(10)
	cQuery += "     A1.A1_FILIAL = '" + xFilial("SA1") + "' AND " + CHR(13) + CHR(10)
	cQuery += "     A1.A1_COD = C5.C5_CLIENTE AND " + CHR(13) + CHR(10)
	cQuery += "     A1.A1_LOJA = C5.C5_LOJACLI AND " + CHR(13) + CHR(10)
	cQuery += "     A1.D_E_L_E_T_ = ' ' " + CHR(13) + CHR(10)
	
	If !lSintetico
		cQuery += " LEFT JOIN " + RetSqlName("SB1") + " AS B1 ON " + CHR(13) + CHR(10)
		cQuery += "     B1.B1_FILIAL = '" + xFilial("SB1") + "' AND " + CHR(13) + CHR(10)
		cQuery += "     B1.B1_COD = C6.C6_PRODUTO AND " + CHR(13) + CHR(10)
		cQuery += "     B1.D_E_L_E_T_ = ' ' " + CHR(13) + CHR(10)
	EndIf
	
	cQuery += "   WHERE C6.D_E_L_E_T_ = ' ' " + CHR(13) + CHR(10)
	cQuery += "   AND  C6.C6_FILIAL = '" + xFilial("SC6") + "'  " + CHR(13) + CHR(10)
	
	// Filtros opcionais
	If !Empty(mv_par03)
		cQuery += "    AND C6.C6_NUM >= '" + mv_par03 + "' " + CHR(13) + CHR(10)
	EndIf
	If !Empty(mv_par04)
		cQuery += "    AND C6.C6_NUM <= '" + mv_par04 + "' " + CHR(13) + CHR(10)
	EndIf
	If !Empty(mv_par05)
		cQuery += "    AND C6.C6_PRODUTO >= '" + mv_par05 + "' " + CHR(13) + CHR(10)
	EndIf
	If !Empty(mv_par06)
		cQuery += "    AND C6.C6_PRODUTO <= '" + mv_par06 + "' " + CHR(13) + CHR(10)
	EndIf
	If !Empty(mv_par07)
		cQuery += "    AND C5.C5_CLIENTE >= '" + mv_par07 + "' " + CHR(13) + CHR(10)
	EndIf
	If !Empty(mv_par08)
		cQuery += "    AND C5.C5_LOJACLI >= '" + mv_par08 + "' " + CHR(13) + CHR(10)
	EndIf
	If !Empty(mv_par09)
		cQuery += "    AND C5.C5_CLIENTE <= '" + mv_par09 + "' " + CHR(13) + CHR(10)
	EndIf
	If !Empty(mv_par10)
		cQuery += "    AND C5.C5_LOJACLI <= '" + mv_par10 + "' " + CHR(13) + CHR(10)
	EndIf
	
	cQuery += "    AND C6.C6_QTDVEN <= ((C6.C6_QTDENT + C6.C6_QTENTMI) - (C6.C6_QTJADEV + C6.C6_QTJACAN)) " + CHR(13) + CHR(10)
	
	If lSintetico
		cQuery += "   GROUP BY C6.C6_FILIAL, C6.C6_NUM, C5.C5_EMISSAO, C5.C5_CLIENTE, C5.C5_LOJACLI, A1.A1_NOME " + CHR(13) + CHR(10)
	EndIf
	
	cQuery += "   ORDER BY C6.C6_FILIAL, C6.C6_NUM" + IIf(!lSintetico, ", C6.C6_ITEM", "") + CHR(13) + CHR(10)
	
	TCQuery cQuery New Alias "QRYPRO"
	
	If (QRYPRO->(EoF()))
		MsgAlert("Não há registros para gerar o relatório", "Aviso")
		QRYPRO->(DbCloseArea())
		Return
	EndIf
	
	oReport:SetMeter(QRYPRO->(RecCount()))
	
	// Oculta colunas conforme o tipo de relatório
	If lSintetico
		// Sintético - mostra apenas colunas marcadas em verde
		oSection1:Cell("C6_ITEM"):Hide()
		oSection1:Cell("C6_ITEM"):HideHeader()
		
		oSection1:Cell("C6_PRODUTO"):Hide()
		oSection1:Cell("C6_PRODUTO"):HideHeader()
		
		oSection1:Cell("B1_DESC"):Hide()
		oSection1:Cell("B1_DESC"):HideHeader()
		
		oSection1:Cell("C6_QTENTMI"):Hide()
		oSection1:Cell("C6_QTENTMI"):HideHeader()
		
		oSection1:Cell("C6_QTJADEV"):Hide()
		oSection1:Cell("C6_QTJADEV"):HideHeader()
		
		oSection1:Cell("C6_QTJACAN"):Hide()
		oSection1:Cell("C6_QTJACAN"):HideHeader()
		
		oSection1:Cell("C6_RESERVA"):Hide()
		oSection1:Cell("C6_RESERVA"):HideHeader()
		
		oSection1:Cell("C9_DATALIB"):Hide()
		oSection1:Cell("C9_DATALIB"):HideHeader()
		
		oSection1:Cell("C9_QTDLIB"):Hide()
		oSection1:Cell("C9_QTDLIB"):HideHeader()
		
		// Oculta colunas da seção 2 - Totalizador
		oSection2:Cell("C6_QTENTMI"):Hide()
		oSection2:Cell("C6_QTJADEV"):Hide()
		oSection2:Cell("C6_QTJACAN"):Hide()
		oSection2:Cell("C9_QTDLIB"):Hide()
	EndIf
	
	oSection1:Init()
	
	While !QRYPRO->(Eof()) .And. !oReport:Cancel()
		// Controla pedidos distintos para o contador
		If aScan(aPedidos, QRYPRO->C6_FILIAL + QRYPRO->C6_NUM) == 0
			aAdd(aPedidos, QRYPRO->C6_FILIAL + QRYPRO->C6_NUM)
		EndIf
		
		// Acumula totalizadores
		nTotQtdVen += QRYPRO->C6_QTDVEN
		nTotQtdEnt += QRYPRO->C6_QTDENT
		
		If !lSintetico
			nTotQtdEntMi += QRYPRO->C6_QTENTMI
			nTotQtdJaDev += QRYPRO->C6_QTJADEV
			nTotQtdJaCan += QRYPRO->C6_QTJACAN
			nTotQtdLib   += QRYPRO->C9_QTDLIB
			
			// Trata campo Reserva vazio - mostra 0,00
			If Empty(AllTrim(QRYPRO->C6_RESERVA))
				oSection1:Cell("C6_RESERVA"):SetValue("0,00")
			EndIf
		EndIf
		
		oSection1:PrintLine()
		QRYPRO->(DbSkip())
		oReport:IncMeter()
	EndDo
	
	oSection1:Finish()
	
	QRYPRO->(DbCloseArea())
	
	// Imprime totalizador
	oReport:SkipLine(2)
	oSection2:Init()
	
	// Linha de separação
	oSection2:Cell("DESCRICAO"):SetValue(Replicate("-", 80))
	oSection2:Cell("C6_QTDVEN"):SetValue("")
	oSection2:Cell("C6_QTDENT"):SetValue("")
	If !lSintetico
		oSection2:Cell("C6_QTENTMI"):SetValue("")
		oSection2:Cell("C6_QTJADEV"):SetValue("")
		oSection2:Cell("C6_QTJACAN"):SetValue("")
		oSection2:Cell("C9_QTDLIB"):SetValue("")
	EndIf
	oSection2:PrintLine()
	
	// Total de pedidos distintos
	oSection2:Cell("DESCRICAO"):SetValue("TOTAL DE PEDIDOS DISTINTOS: " + cValToChar(Len(aPedidos)))
	oSection2:Cell("C6_QTDVEN"):SetValue("")
	oSection2:Cell("C6_QTDENT"):SetValue("")
	If !lSintetico
		oSection2:Cell("C6_QTENTMI"):SetValue("")
		oSection2:Cell("C6_QTJADEV"):SetValue("")
		oSection2:Cell("C6_QTJACAN"):SetValue("")
		oSection2:Cell("C9_QTDLIB"):SetValue("")
	EndIf
	oSection2:PrintLine()
	
	// Totais das colunas
	oSection2:Cell("DESCRICAO"):SetValue("TOTAIS:")
	oSection2:Cell("C6_QTDVEN"):SetValue(nTotQtdVen)
	oSection2:Cell("C6_QTDENT"):SetValue(nTotQtdEnt)
	If !lSintetico
		oSection2:Cell("C6_QTENTMI"):SetValue(nTotQtdEntMi)
		oSection2:Cell("C6_QTJADEV"):SetValue(nTotQtdJaDev)
		oSection2:Cell("C6_QTJACAN"):SetValue(nTotQtdJaCan)
		oSection2:Cell("C9_QTDLIB"):SetValue(nTotQtdLib)
	EndIf
	oSection2:PrintLine()
	
	oSection2:Finish()
	
Return


/*/
???????????????????????????????????????????????????????????????????????????????
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±?????????????????????????????????????????????????????????????????????????±±
±±?Programa  ?VALIDPERG ? Autor ?                           ? Data ?      ?±±
±±?????????????????????????????????????????????????????????????????????????±±
±±?Descricao ? Criacao de perguntas para o programa                      ?±±
±±?????????????????????????????????????????????????????????????????????????±±
±±?Uso       ? MATR001                                                    ?±±
±±?????????????????????????????????????????????????????????????????????????±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
???????????????????????????????????????????????????????????????????????????????
/*/

Static Function ValidPerg(cPerg)
	Local _sAlias := Alias()
	Local aRegs   := {}
	Local nX      := 0
	Local nY      := 0
	
	DBSelectArea("SX1")
	DBSetOrder(1)
	cPerg := PADR(cPerg, 10)
	
	AADD(aRegs, {cPerg, "01", "Emissao De       ?", "", "", "mv_ch1", "D", 08, 0, 0, "G", "", "mv_par01", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""})
	AADD(aRegs, {cPerg, "02", "Emissao Ate      ?", "", "", "mv_ch2", "D", 08, 0, 0, "G", "", "mv_par02", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""})
	AADD(aRegs, {cPerg, "03", "Pedido De        ?", "", "", "mv_ch3", "C", 06, 0, 0, "G", "", "mv_par03", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "SC5", ""})
	AADD(aRegs, {cPerg, "04", "Pedido Ate       ?", "", "", "mv_ch4", "C", 06, 0, 0, "G", "", "mv_par04", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "SC5", ""})
	AADD(aRegs, {cPerg, "05", "Produto De       ?", "", "", "mv_ch5", "C", 15, 0, 0, "G", "", "mv_par05", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "SB1", ""})
	AADD(aRegs, {cPerg, "06", "Produto Ate      ?", "", "", "mv_ch6", "C", 15, 0, 0, "G", "", "mv_par06", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "SB1", ""})
	AADD(aRegs, {cPerg, "07", "Cliente De       ?", "", "", "mv_ch7", "C", 06, 0, 0, "G", "", "mv_par07", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "SA1", ""})
	AADD(aRegs, {cPerg, "08", "Loja De          ?", "", "", "mv_ch8", "C", 02, 0, 0, "G", "", "mv_par08", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""})
	AADD(aRegs, {cPerg, "09", "Cliente Ate      ?", "", "", "mv_ch9", "C", 06, 0, 0, "G", "", "mv_par09", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "SA1", ""})
	AADD(aRegs, {cPerg, "10", "Loja Ate         ?", "", "", "mv_cha", "C", 02, 0, 0, "G", "", "mv_par10", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""})
	AADD(aRegs, {cPerg, "11", "Tipo Relatorio   ?", "", "", "mv_chb", "N", 01, 0, 0, "C", "", "mv_par11", "Sintetico", "", "", "", "", "Analitico", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""})
	
	DBSkip()
	
	Do While x1_grupo == cPerg
		RecLock("SX1", .F.)
		DBDelete()
		DBSkip()
	EndDo
	
	For nX := 1 To Len(aRegs)
		If !DBSeek(cPerg + aRegs[nX][2])
			RecLock("SX1", .T.)
			For nY := 1 To FCount()
				If nY <= Len(aRegs[nX])
					FieldPut(nY, aRegs[nX][nY])
				EndIf
			Next
			MsUnlock()
		EndIf
	Next
	
	DBSelectArea(_sAlias)
	
Return
