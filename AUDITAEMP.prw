#INCLUDE "TOTVS.ch"
#INCLUDE "topconn.ch"

/*/{Protheus.doc} U_AUDITAEMP
Relatório de Auditoria do C6_QTDEMP
Mostra as divergências entre C6_QTDEMP e a realidade (SC9 - SD2 - SD3)
SEM FAZER ALTERAÇÕES - apenas para análise

@type User Function
@author Adaptado
@since 28/01/2026
@version 1.0
/*/

User Function AUDITAEMP()
    Local aArea := GetArea()
    Local cPerg := "AUDITAEMP"
    
    ValidPerg(cPerg)
    
    If !Pergunte(cPerg, .T.)
        Return
    EndIf
    
    Processa({|| GeraRelatorio()}, "Processando", "Gerando Relatório de Auditoria...", .T.)
    
    RestArea(aArea)
    
Return

/*/{Protheus.doc} GeraRelatorio
Gera planilha Excel com análise das divergências
/*/
Static Function GeraRelatorio()
    Local cQuery := ""
    Local oFWMsExcel
    Local oExcel
    Local cArquivo := GetTempPath() + 'AuditEmp_' + DTOS(Date()) + '_' + StrTran(Time(), ':', '') + '.xml'
    Local nDiverg := 0
    Local nOK := 0
    
    cQuery := " SELECT " + CRLF
    cQuery += "     SC6.C6_FILIAL AS FILIAL, " + CRLF
    cQuery += "     SC6.C6_NUM AS PEDIDO, " + CRLF
    cQuery += "     SC6.C6_ITEM AS ITEM, " + CRLF
    cQuery += "     SC6.C6_PRODUTO AS PRODUTO, " + CRLF
    cQuery += "     SB1.B1_DESC AS DESCRICAO, " + CRLF
    cQuery += "     SC5.C5_CLIENTE AS CLIENTE, " + CRLF
    cQuery += "     SA1.A1_NOME AS NOME_CLIENTE, " + CRLF
    cQuery += "     SC5.C5_EMISSAO AS EMISSAO, " + CRLF
    cQuery += " " + CRLF
    cQuery += "     SC6.C6_QTDVEN AS QTD_VENDIDA, " + CRLF
    cQuery += "     SC6.C6_QTDEMP AS C6_QTDEMP_ATUAL, " + CRLF
    cQuery += "     SC6.C6_QTDENT AS QTD_ENTREGUE, " + CRLF
    cQuery += " " + CRLF
    cQuery += "     -- Detalhamento das quantidades " + CRLF
    cQuery += "     ISNULL(SC9_TOTAL.TOTAL_SC9, 0) AS QTD_SC9_LIBERADA, " + CRLF
    cQuery += "     ISNULL(SD2_TOTAL.TOTAL_SD2, 0) AS QTD_SD2_FATURADA, " + CRLF
    cQuery += "     ISNULL(SD3_TOTAL.TOTAL_SD3, 0) AS QTD_SD3_MI, " + CRLF
    cQuery += " " + CRLF
    cQuery += "     -- Cálculo do que DEVERIA ser o C6_QTDEMP " + CRLF
    cQuery += "     ( " + CRLF
    cQuery += "         CASE  " + CRLF
    cQuery += "             WHEN (ISNULL(SC9_TOTAL.TOTAL_SC9, 0) - ISNULL(SD2_TOTAL.TOTAL_SD2, 0) - ISNULL(SD3_TOTAL.TOTAL_SD3, 0)) < 0  " + CRLF
    cQuery += "             THEN 0  " + CRLF
    cQuery += "             ELSE (ISNULL(SC9_TOTAL.TOTAL_SC9, 0) - ISNULL(SD2_TOTAL.TOTAL_SD2, 0) - ISNULL(SD3_TOTAL.TOTAL_SD3, 0))  " + CRLF
    cQuery += "         END " + CRLF
    cQuery += "     ) AS C6_QTDEMP_CORRETO, " + CRLF
    cQuery += " " + CRLF
    cQuery += "     -- Diferença (Divergência) " + CRLF
    cQuery += "     ( " + CRLF
    cQuery += "         SC6.C6_QTDEMP -  " + CRLF
    cQuery += "         CASE  " + CRLF
    cQuery += "             WHEN (ISNULL(SC9_TOTAL.TOTAL_SC9, 0) - ISNULL(SD2_TOTAL.TOTAL_SD2, 0) - ISNULL(SD3_TOTAL.TOTAL_SD3, 0)) < 0  " + CRLF
    cQuery += "             THEN 0  " + CRLF
    cQuery += "             ELSE (ISNULL(SC9_TOTAL.TOTAL_SC9, 0) - ISNULL(SD2_TOTAL.TOTAL_SD2, 0) - ISNULL(SD3_TOTAL.TOTAL_SD3, 0))  " + CRLF
    cQuery += "         END " + CRLF
    cQuery += "     ) AS DIFERENCA, " + CRLF
    cQuery += " " + CRLF
    cQuery += "     -- Tipo de Divergência " + CRLF
    cQuery += "     CASE " + CRLF
    cQuery += "         WHEN SC6.C6_QTDEMP > 0 AND ISNULL(SC9_TOTAL.TOTAL_SC9, 0) = 0 " + CRLF
    cQuery += "         THEN 'FANTASMA TOTAL' " + CRLF
    cQuery += "         WHEN SC6.C6_QTDEMP <> ( " + CRLF
    cQuery += "             CASE  " + CRLF
    cQuery += "                 WHEN (ISNULL(SC9_TOTAL.TOTAL_SC9, 0) - ISNULL(SD2_TOTAL.TOTAL_SD2, 0) - ISNULL(SD3_TOTAL.TOTAL_SD3, 0)) < 0  " + CRLF
    cQuery += "                 THEN 0  " + CRLF
    cQuery += "                 ELSE (ISNULL(SC9_TOTAL.TOTAL_SC9, 0) - ISNULL(SD2_TOTAL.TOTAL_SD2, 0) - ISNULL(SD3_TOTAL.TOTAL_SD3, 0))  " + CRLF
    cQuery += "             END " + CRLF
    cQuery += "         ) " + CRLF
    cQuery += "         THEN 'DIVERGENCIA' " + CRLF
    cQuery += "         ELSE 'OK' " + CRLF
    cQuery += "     END AS TIPO_PROBLEMA " + CRLF
    cQuery += " " + CRLF
    cQuery += " FROM " + RetSqlName("SC6") + " SC6 (NOLOCK) " + CRLF
    cQuery += " " + CRLF
    cQuery += " INNER JOIN " + RetSqlName("SC5") + " SC5 (NOLOCK) ON " + CRLF
    cQuery += "     SC5.C5_FILIAL = SC6.C6_FILIAL " + CRLF
    cQuery += "     AND SC5.C5_NUM = SC6.C6_NUM " + CRLF
    cQuery += "     AND SC5.D_E_L_E_T_ = ' ' " + CRLF
    cQuery += " " + CRLF
    cQuery += " INNER JOIN " + RetSqlName("SB1") + " SB1 (NOLOCK) ON " + CRLF
    cQuery += "     SB1.B1_COD = SC6.C6_PRODUTO " + CRLF
    cQuery += "     AND SB1.D_E_L_E_T_ = ' ' " + CRLF
    cQuery += " " + CRLF
    cQuery += " INNER JOIN " + RetSqlName("SA1") + " SA1 (NOLOCK) ON " + CRLF
    cQuery += "     SA1.A1_COD = SC5.C5_CLIENTE " + CRLF
    cQuery += "     AND SA1.A1_LOJA = SC5.C5_LOJACLI " + CRLF
    cQuery += "     AND SA1.D_E_L_E_T_ = ' ' " + CRLF
    cQuery += " " + CRLF
    cQuery += " LEFT JOIN ( " + CRLF
    cQuery += "     SELECT C9_FILIAL, C9_PEDIDO, C9_ITEM, C9_PRODUTO, SUM(C9_QTDLIB) AS TOTAL_SC9 " + CRLF
    cQuery += "     FROM " + RetSqlName("SC9") + " (NOLOCK) " + CRLF
    cQuery += "     WHERE D_E_L_E_T_ = ' ' AND C9_NFISCAL = ' ' AND C9_BLEST = ' ' AND C9_BLCRED = ' ' " + CRLF
    cQuery += "     GROUP BY C9_FILIAL, C9_PEDIDO, C9_ITEM, C9_PRODUTO " + CRLF
    cQuery += " ) SC9_TOTAL ON SC9_TOTAL.C9_FILIAL = SC6.C6_FILIAL AND SC9_TOTAL.C9_PEDIDO = SC6.C6_NUM " + CRLF
    cQuery += "     AND SC9_TOTAL.C9_ITEM = SC6.C6_ITEM AND SC9_TOTAL.C9_PRODUTO = SC6.C6_PRODUTO " + CRLF
    cQuery += " " + CRLF
    cQuery += " LEFT JOIN ( " + CRLF
    cQuery += "     SELECT D2_FILIAL, D2_PEDIDO, D2_ITEMPV, D2_COD, SUM(D2_QUANT) AS TOTAL_SD2 " + CRLF
    cQuery += "     FROM " + RetSqlName("SD2") + " (NOLOCK) " + CRLF
    cQuery += "     WHERE D_E_L_E_T_ = ' ' " + CRLF
    cQuery += "     GROUP BY D2_FILIAL, D2_PEDIDO, D2_ITEMPV, D2_COD " + CRLF
    cQuery += " ) SD2_TOTAL ON SD2_TOTAL.D2_FILIAL = SC6.C6_FILIAL AND SD2_TOTAL.D2_PEDIDO = SC6.C6_NUM " + CRLF
    cQuery += "     AND SD2_TOTAL.D2_ITEMPV = SC6.C6_ITEM AND SD2_TOTAL.D2_COD = SC6.C6_PRODUTO " + CRLF
    cQuery += " " + CRLF
    cQuery += " LEFT JOIN ( " + CRLF
    cQuery += "     SELECT D3_FILIAL, LEFT(D3_YPEDVEN, 6) AS PEDIDO, " + CRLF
    cQuery += "            RTRIM(LTRIM(RIGHT(D3_YPEDVEN, 2))) AS ITEM, D3_COD, SUM(D3_QUANT) AS TOTAL_SD3 " + CRLF
    cQuery += "     FROM " + RetSqlName("SD3") + " (NOLOCK) " + CRLF
    cQuery += "     WHERE D_E_L_E_T_ = ' ' AND D3_ESTORNO <> 'S' AND D3_YPEDVEN <> '' AND D3_LOCAL <> '08' " + CRLF
    cQuery += "     GROUP BY D3_FILIAL, LEFT(D3_YPEDVEN, 6), RTRIM(LTRIM(RIGHT(D3_YPEDVEN, 2))), D3_COD " + CRLF
    cQuery += " ) SD3_TOTAL ON SD3_TOTAL.D3_FILIAL = SC6.C6_FILIAL AND SD3_TOTAL.PEDIDO = SC6.C6_NUM " + CRLF
    cQuery += "     AND SD3_TOTAL.ITEM = SC6.C6_ITEM AND SD3_TOTAL.D3_COD = SC6.C6_PRODUTO " + CRLF
    cQuery += " " + CRLF
    cQuery += " WHERE SC6.D_E_L_E_T_ = ' ' " + CRLF
    cQuery += "   AND SC6.C6_FILIAL BETWEEN '" + MV_PAR01 + "' AND '" + MV_PAR02 + "' " + CRLF
    cQuery += "   AND SC6.C6_NUM BETWEEN '" + MV_PAR03 + "' AND '" + MV_PAR04 + "' " + CRLF
    
    If MV_PAR05 == 1  // Somente divergentes
        cQuery += " HAVING ( " + CRLF
        cQuery += "     SC6.C6_QTDEMP -  " + CRLF
        cQuery += "     CASE  " + CRLF
        cQuery += "         WHEN (ISNULL(SC9_TOTAL.TOTAL_SC9, 0) - ISNULL(SD2_TOTAL.TOTAL_SD2, 0) - ISNULL(SD3_TOTAL.TOTAL_SD3, 0)) < 0  " + CRLF
        cQuery += "         THEN 0  " + CRLF
        cQuery += "         ELSE (ISNULL(SC9_TOTAL.TOTAL_SC9, 0) - ISNULL(SD2_TOTAL.TOTAL_SD2, 0) - ISNULL(SD3_TOTAL.TOTAL_SD3, 0))  " + CRLF
        cQuery += "     END " + CRLF
        cQuery += " ) <> 0 " + CRLF
    EndIf
    
    cQuery += " ORDER BY SC6.C6_FILIAL, SC6.C6_NUM, SC6.C6_ITEM " + CRLF
    
    MemoWrite("C:\temp\AUDITAEMP.sql", cQuery)
    
    TCQUERY cQuery NEW ALIAS "TMPAUD"
    
    TCSetField("TMPAUD", "EMISSAO", "D")
    TCSetField("TMPAUD", "QTD_VENDIDA", "N", 12, 2)
    TCSetField("TMPAUD", "C6_QTDEMP_ATUAL", "N", 12, 2)
    TCSetField("TMPAUD", "QTD_ENTREGUE", "N", 12, 2)
    TCSetField("TMPAUD", "QTD_SC9_LIBERADA", "N", 12, 2)
    TCSetField("TMPAUD", "QTD_SD2_FATURADA", "N", 12, 2)
    TCSetField("TMPAUD", "QTD_SD3_MI", "N", 12, 2)
    TCSetField("TMPAUD", "C6_QTDEMP_CORRETO", "N", 12, 2)
    TCSetField("TMPAUD", "DIFERENCA", "N", 12, 2)
    
    oFWMsExcel := FWMSExcel():New()
    
    oFWMsExcel:AddworkSheet("Auditoria C6_QTDEMP")
    oFWMsExcel:AddTable("Auditoria C6_QTDEMP", "AUDIT")
    
    oFWMsExcel:AddColumn("Auditoria C6_QTDEMP", "AUDIT", "Filial", 1)
    oFWMsExcel:AddColumn("Auditoria C6_QTDEMP", "AUDIT", "Pedido", 1)
    oFWMsExcel:AddColumn("Auditoria C6_QTDEMP", "AUDIT", "Item", 1)
    oFWMsExcel:AddColumn("Auditoria C6_QTDEMP", "AUDIT", "Produto", 1)
    oFWMsExcel:AddColumn("Auditoria C6_QTDEMP", "AUDIT", "Descrição", 1)
    oFWMsExcel:AddColumn("Auditoria C6_QTDEMP", "AUDIT", "Cliente", 1)
    oFWMsExcel:AddColumn("Auditoria C6_QTDEMP", "AUDIT", "Nome Cliente", 1)
    oFWMsExcel:AddColumn("Auditoria C6_QTDEMP", "AUDIT", "Emissão", 1)
    oFWMsExcel:AddColumn("Auditoria C6_QTDEMP", "AUDIT", "Qtd Vendida", 1)
    oFWMsExcel:AddColumn("Auditoria C6_QTDEMP", "AUDIT", "C6_QTDEMP Atual", 1)
    oFWMsExcel:AddColumn("Auditoria C6_QTDEMP", "AUDIT", "Qtd Entregue", 1)
    oFWMsExcel:AddColumn("Auditoria C6_QTDEMP", "AUDIT", "SC9 Liberada", 1)
    oFWMsExcel:AddColumn("Auditoria C6_QTDEMP", "AUDIT", "SD2 Faturada", 1)
    oFWMsExcel:AddColumn("Auditoria C6_QTDEMP", "AUDIT", "SD3 MI", 1)
    oFWMsExcel:AddColumn("Auditoria C6_QTDEMP", "AUDIT", "C6_QTDEMP Correto", 1)
    oFWMsExcel:AddColumn("Auditoria C6_QTDEMP", "AUDIT", "DIFERENÇA", 1)
    oFWMsExcel:AddColumn("Auditoria C6_QTDEMP", "AUDIT", "Tipo Problema", 1)
    
    While TMPAUD->(!EOF())
        
        oFWMsExcel:AddRow("Auditoria C6_QTDEMP", "AUDIT", {;
            TMPAUD->FILIAL,;
            TMPAUD->PEDIDO,;
            TMPAUD->ITEM,;
            TMPAUD->PRODUTO,;
            TMPAUD->DESCRICAO,;
            TMPAUD->CLIENTE,;
            TMPAUD->NOME_CLIENTE,;
            TMPAUD->EMISSAO,;
            TMPAUD->QTD_VENDIDA,;
            TMPAUD->C6_QTDEMP_ATUAL,;
            TMPAUD->QTD_ENTREGUE,;
            TMPAUD->QTD_SC9_LIBERADA,;
            TMPAUD->QTD_SD2_FATURADA,;
            TMPAUD->QTD_SD3_MI,;
            TMPAUD->C6_QTDEMP_CORRETO,;
            TMPAUD->DIFERENCA,;
            TMPAUD->TIPO_PROBLEMA})
        
        If TMPAUD->TIPO_PROBLEMA != "OK"
            nDiverg++
        Else
            nOK++
        EndIf
        
        TMPAUD->(dbSkip())
    EndDo
    
    TMPAUD->(dbCloseArea())
    
    oFWMsExcel:Activate()
    oFWMsExcel:GetXMLFile(cArquivo)
    
    oExcel := MsExcel():New()
    oExcel:WorkBooks:Open(cArquivo)
    oExcel:SetVisible(.T.)
    oExcel:Destroy()
    
    MsgInfo("Auditoria concluída!" + CRLF + ;
            "Registros OK: " + cValToChar(nOK) + CRLF + ;
            "Divergências: " + cValToChar(nDiverg) + CRLF + CRLF + ;
            "Planilha: " + cArquivo, "Auditoria Finalizada")
    
Return

/*/{Protheus.doc} ValidPerg
Cria perguntas no SX1
/*/
Static Function ValidPerg(cPerg)
    Local aRegs := {}
    
    dbSelectArea("SX1")
    dbSetOrder(1)
    cPerg := PADR(cPerg, 10)
    
    AADD(aRegs, {cPerg, "01", "Filial De            ?", "", "", "mv_ch1", "C", 02, 0, 0, "G", "", "mv_par01", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""})
    AADD(aRegs, {cPerg, "02", "Filial Até           ?", "", "", "mv_ch2", "C", 02, 0, 0, "G", "", "mv_par02", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""})
    AADD(aRegs, {cPerg, "03", "Pedido De            ?", "", "", "mv_ch3", "C", 06, 0, 0, "G", "", "mv_par03", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "SC5", ""})
    AADD(aRegs, {cPerg, "04", "Pedido Até           ?", "", "", "mv_ch4", "C", 06, 0, 0, "G", "", "mv_par04", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "SC5", ""})
    AADD(aRegs, {cPerg, "05", "Mostrar              :", "", "", "mv_ch5", "N", 01, 0, 0, "C", "", "mv_par05", "Somente Divergentes", "", "", "", "", "Todos", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""})
    
    For i := 1 To Len(aRegs)
        If !DbSeek(cPerg + aRegs[i,2])
            RecLock("SX1", .T.)
            For j := 1 To FCount()
                If j <= Len(aRegs[i])
                    FieldPut(j, aRegs[i,j])
                EndIf
            Next
            MsUnlock()
        EndIf
    Next
    
Return
