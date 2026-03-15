#INCLUDE "TOTVS.ch"
#INCLUDE "topconn.ch"

/*/{Protheus.doc} U_FNRFDEV
Funcao para recalcular o campo C6_QTJADEV (Quantidade Devolvida)
Calcula pela soma de ZE4_QUANT da ZE4 onde ZE4_TPMOV = 'D'

@type User Function
@author Adaptado
@since 08/02/2026
@version 1.0
/*/

User Function FNRFDEV()
    Local aArea      := GetArea()
    Local dDtIni     := CToD("")
    Local dDtFim     := CToD("")
    Local cPedIni    := ""
    Local cPedFim    := ""
    Local cProdIni   := ""
    Local cProdFim   := ""
    Local nTipoVis   := 1  // 1=Apenas Diferencas, 2=Todos
    
    If !MsgYesNo("Esta rotina ira recalcular o campo C6_QTJADEV (Quantidade Devolvida)." + CRLF + ;
                 "Deseja continuar?", "Atencao")
        RestArea(aArea)
        Return
    EndIf
    
    // Monta tela de parametros
    If !MontaPerguntas(@dDtIni, @dDtFim, @cPedIni, @cPedFim, @cProdIni, @cProdFim, @nTipoVis)
        MsgInfo("Operacao cancelada.", "Aviso")
        RestArea(aArea)
        Return
    EndIf
    
    // Processa e mostra previa
    ProcPrevia(dDtIni, dDtFim, cPedIni, cPedFim, cProdIni, cProdFim, nTipoVis)
    
    RestArea(aArea)
    
Return

/*/{Protheus.doc} MontaPerguntas
Funcao para montar tela de parametros
@type function
/*/
Static Function MontaPerguntas(dDtIni, dDtFim, cPedIni, cPedFim, cProdIni, cProdFim, nTipoVis)
    Local lRet          := .F.
    Local oDlg
    Local oGetDtIni, oGetDtFim
    Local oGetPedIni, oGetPedFim
    Local oGetProdIni, oGetProdFim
    Local oCmbTipoVis
    Local aTipoVis      := {"Apenas Diferencas (DEVOLUCOES)", "Todos (DEVOLUCOES)"}
    Local cTipoVis      := aTipoVis[1]
    Local nOpcA         := 0
    
    // Valores padrao
    dDtIni     := FirstDate(Date())
    dDtFim     := Date()
    cPedIni    := Space(6)
    cPedFim    := Replicate("Z", 6)
    cProdIni   := Space(15)
    cProdFim   := Replicate("Z", 15)
    nTipoVis   := 1
    
    DEFINE MSDIALOG oDlg TITLE "Recalculo C6_QTJADEV - Quantidade Devolvida" FROM 000,000 TO 280,400 PIXEL
        
        @ 010,010 SAY "Data Emissao De:" SIZE 060,007 PIXEL OF oDlg
        @ 010,080 MSGET oGetDtIni VAR dDtIni SIZE 060,010 PIXEL OF oDlg VALID !Empty(dDtIni)
        
        @ 025,010 SAY "Data Emissao Ate:" SIZE 060,007 PIXEL OF oDlg
        @ 025,080 MSGET oGetDtFim VAR dDtFim SIZE 060,010 PIXEL OF oDlg VALID !Empty(dDtFim) .And. dDtFim >= dDtIni
        
        @ 040,010 SAY "Pedido Venda De:" SIZE 060,007 PIXEL OF oDlg
        @ 040,080 MSGET oGetPedIni VAR cPedIni SIZE 060,010 PIXEL OF oDlg F3 "SC5"
        
        @ 055,010 SAY "Pedido Venda Ate:" SIZE 060,007 PIXEL OF oDlg
        @ 055,080 MSGET oGetPedFim VAR cPedFim SIZE 060,010 PIXEL OF oDlg F3 "SC5"
        
        @ 070,010 SAY "Produto De:" SIZE 060,007 PIXEL OF oDlg
        @ 070,080 MSGET oGetProdIni VAR cProdIni SIZE 060,010 PIXEL OF oDlg F3 "SB1"
        
        @ 085,010 SAY "Produto Ate:" SIZE 060,007 PIXEL OF oDlg
        @ 085,080 MSGET oGetProdFim VAR cProdFim SIZE 060,010 PIXEL OF oDlg F3 "SB1"
        
        @ 100,010 SAY "Visualizar:" SIZE 060,007 PIXEL OF oDlg
        @ 100,080 MSCOMBOBOX oCmbTipoVis VAR cTipoVis ITEMS aTipoVis SIZE 080,010 PIXEL OF oDlg
        
        @ 125,080 BUTTON "Confirmar" SIZE 040,012 PIXEL OF oDlg ACTION (nOpcA := 1, oDlg:End())
        @ 125,130 BUTTON "Cancelar"  SIZE 040,012 PIXEL OF oDlg ACTION (nOpcA := 0, oDlg:End())
        
    ACTIVATE MSDIALOG oDlg CENTERED
    
    If nOpcA == 1
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
        
        nTipoVis := aScan(aTipoVis, cTipoVis)
        
        lRet := .T.
    EndIf
    
Return lRet

/*/{Protheus.doc} ProcPrevia
Processa e mostra previa dos dados
@type function
/*/
Static Function ProcPrevia(dDtIni, dDtFim, cPedIni, cPedFim, cProdIni, cProdFim, nTipoVis)
    Local cQuery := ""
    Local aDados := {}
    Local cTitulo := ""
    
    // Monta query
    cQuery := MontaQuery(dDtIni, dDtFim, cPedIni, cPedFim, cProdIni, cProdFim, nTipoVis)
    
    // Executa query
    Processa({|| aDados := ExecutaQuery(cQuery)}, "Aguarde", "Consultando dados...", .F.)
    
    If Len(aDados) == 0
        MsgInfo("Nenhum registro encontrado com os parametros informados.", "Aviso")
        Return
    EndIf
    
    // Define titulo
    cTitulo := IIF(nTipoVis == 1, "Previa - Apenas Diferencas", "Previa - Todos os Registros")
    
    // Mostra browse
    MostraBrowse(aDados, cTitulo, nTipoVis)
    
Return

/*/{Protheus.doc} MontaQuery
Monta a query de consulta
@type function
/*/
Static Function MontaQuery(dDtIni, dDtFim, cPedIni, cPedFim, cProdIni, cProdFim, nTipoVis)
    Local cQuery := ""
    
    cQuery := " SELECT " + CRLF
    cQuery += "     SC6.R_E_C_N_O_ AS RECNOSC6, " + CRLF
    cQuery += "     SC6.C6_FILIAL, " + CRLF
    cQuery += "     SC6.C6_NUM AS PEDIDO, " + CRLF
    cQuery += "     SC6.C6_ITEM AS ITEM, " + CRLF
    cQuery += "     SC6.C6_PRODUTO AS PRODUTO, " + CRLF
    cQuery += "     SC6.C6_QTDVEN AS QTD_VENDIDA, " + CRLF
    cQuery += "     SC6.C6_QTJADEV AS C6_QTJADEV_ATUAL, " + CRLF
    cQuery += "     ISNULL(ZE4_DEV.TOTAL_ZE4, 0) AS QTDEV_NOVO, " + CRLF
    cQuery += "     (SC6.C6_QTJADEV - ISNULL(ZE4_DEV.TOTAL_ZE4, 0)) AS DIFERENCA " + CRLF
    cQuery += " FROM " + RetSqlName("SC6") + " SC6 " + CRLF
    cQuery += " INNER JOIN " + RetSqlName("SC5") + " SC5 ON " + CRLF
    cQuery += "     SC5.C5_FILIAL = SC6.C6_FILIAL " + CRLF
    cQuery += "     AND SC5.C5_NUM = SC6.C6_NUM " + CRLF
    cQuery += "     AND SC5.D_E_L_E_T_ = ' ' " + CRLF
    cQuery += " LEFT JOIN ( " + CRLF
    cQuery += "     SELECT " + CRLF
    cQuery += "         ZE4_FILIAL, " + CRLF
    cQuery += "         ZE4_PEDIDO, " + CRLF
    cQuery += "         ZE4_ITEM, " + CRLF
    cQuery += "         ZE4_PRODUT, " + CRLF
    cQuery += "         SUM(ZE4_QUANT) AS TOTAL_ZE4 " + CRLF
    cQuery += "     FROM " + RetSqlName("ZE4") + " ZE4 " + CRLF
    cQuery += "     WHERE ZE4.D_E_L_E_T_ = ' ' " + CRLF
    cQuery += "       AND ZE4.ZE4_TPMOV = 'D' " + CRLF
    cQuery += "     GROUP BY ZE4_FILIAL, ZE4_PEDIDO, ZE4_ITEM, ZE4_PRODUT " + CRLF
    cQuery += " ) ZE4_DEV ON " + CRLF
    cQuery += "     ZE4_DEV.ZE4_FILIAL = SC6.C6_FILIAL " + CRLF
    cQuery += "     AND ZE4_DEV.ZE4_PEDIDO = SC6.C6_NUM " + CRLF
    cQuery += "     AND ZE4_DEV.ZE4_ITEM = SC6.C6_ITEM " + CRLF
    cQuery += "     AND ZE4_DEV.ZE4_PRODUT = SC6.C6_PRODUTO " + CRLF
    cQuery += " WHERE SC6.D_E_L_E_T_ = ' ' " + CRLF
    cQuery += "   AND SC6.C6_FILIAL = '" + xFilial("SC6") + "' " + CRLF
    cQuery += "   AND SC5.C5_EMISSAO BETWEEN '" + DToS(dDtIni) + "' AND '" + DToS(dDtFim) + "' " + CRLF
    cQuery += "   AND SC6.C6_NUM BETWEEN '" + cPedIni + "' AND '" + cPedFim + "' " + CRLF
    cQuery += "   AND SC6.C6_PRODUTO BETWEEN '" + cProdIni + "' AND '" + cProdFim + "' " + CRLF
    
    // Se for apenas diferencas
    If nTipoVis == 1
        cQuery += "   AND SC6.C6_QTJADEV <> ISNULL(ZE4_DEV.TOTAL_ZE4, 0) " + CRLF
    EndIf
    
    cQuery += " ORDER BY SC6.C6_NUM, SC6.C6_ITEM " + CRLF
    
Return cQuery

/*/{Protheus.doc} ExecutaQuery
Executa query e retorna array com dados
@type function
/*/
Static Function ExecutaQuery(cQuery)
    Local aDados := {}
    Local cAlias := GetNextAlias()
    
    TCQUERY cQuery NEW ALIAS (cAlias)
    
    TCSetField(cAlias, "QTD_VENDIDA", "N", 12, 2)
    TCSetField(cAlias, "C6_QTJADEV_ATUAL", "N", 12, 2)
    TCSetField(cAlias, "QTDEV_NOVO", "N", 12, 2)
    TCSetField(cAlias, "DIFERENCA", "N", 12, 2)
    
    While (cAlias)->(!EOF())
        AADD(aDados, {(cAlias)->RECNOSC6, ;
                      (cAlias)->C6_FILIAL, ;
                      (cAlias)->PEDIDO, ;
                      (cAlias)->ITEM, ;
                      (cAlias)->PRODUTO, ;
                      (cAlias)->QTD_VENDIDA, ;
                      (cAlias)->C6_QTJADEV_ATUAL, ;
                      (cAlias)->QTDEV_NOVO, ;
                      (cAlias)->DIFERENCA})
        (cAlias)->(dbSkip())
    EndDo
    
    (cAlias)->(dbCloseArea())
    
Return aDados

/*/{Protheus.doc} MostraBrowse
Mostra browse com os dados
@type function
/*/
Static Function MostraBrowse(aDados, cTitulo, nTipoVis)
    Local oDlg
    Local oBrowse
    Local nOpcao := 0
    
    DEFINE MSDIALOG oDlg TITLE cTitulo FROM 000,000 TO 500,900 PIXEL
        
        @ 005,005 SAY "Total de registros: " + cValToChar(Len(aDados)) SIZE 200,010 PIXEL OF oDlg
        
        oBrowse := TWBrowse():New(015,005,440,210,,{"Filial","Pedido","Item","Produto","Qtd Vendida","C6_QTJADEV Atual","Devolucao Localizada","Diferenca"},,oDlg,,,,,,,,,,,,.F.,,.T.,,.F.,,,)
        oBrowse:SetArray(aDados)
        oBrowse:bLine := {|| {aDados[oBrowse:nAt][2], ;
                              aDados[oBrowse:nAt][3], ;
                              aDados[oBrowse:nAt][4], ;
                              aDados[oBrowse:nAt][5], ;
                              Transform(aDados[oBrowse:nAt][6], "@E 999,999.99"), ;
                              Transform(aDados[oBrowse:nAt][7], "@E 999,999.99"), ;
                              Transform(aDados[oBrowse:nAt][8], "@E 999,999.99"), ;
                              Transform(aDados[oBrowse:nAt][9], "@E 999,999.99")}}
        
        // Se for diferencas, mostra botao GRAVAR
        If nTipoVis == 1
            @ 230,300 BUTTON "Gravar C6_QTJADEV" SIZE 060,012 PIXEL OF oDlg ;
                ACTION (nOpcao := 1, oDlg:End())
            @ 230,370 BUTTON "Fechar" SIZE 040,012 PIXEL OF oDlg ;
                ACTION (nOpcao := 0, oDlg:End())
        Else
            @ 230,350 BUTTON "Fechar" SIZE 040,012 PIXEL OF oDlg ;
                ACTION (nOpcao := 0, oDlg:End())
        EndIf
        
    ACTIVATE MSDIALOG oDlg CENTERED
    
    // Se confirmou gravacao
    If nOpcao == 1 .And. nTipoVis == 1
        If MsgYesNo("Confirma a gravacao do campo C6_QTJADEV para " + cValToChar(Len(aDados)) + " registro(s)?", "Confirmacao")
            GravarC6QTJADEV(aDados)
        EndIf
    EndIf
    
Return

/*/{Protheus.doc} GravarC6QTJADEV
Grava o campo C6_QTJADEV
@type function
/*/
Static Function GravarC6QTJADEV(aDados)
    Local nRecalc := 0
    Local nErros := 0
    
    dbSelectArea("SC6")
    SC6->(dbSetOrder(1))
    
    Processa({|| ProcGrava(@aDados, @nRecalc, @nErros)}, "Gravando", "Atualizando C6_QTJADEV...", .T.)
    
    MsgInfo("Processo concluido!" + CRLF + ;
            "Registros atualizados: " + cValToChar(nRecalc) + CRLF + ;
            "Erros: " + cValToChar(nErros), "Finalizado")
    
Return

/*/{Protheus.doc} ProcGrava
Processa gravacao
@type function
/*/
Static Function ProcGrava(aDados, nRecalc, nErros)
    Local i
    
    ProcRegua(Len(aDados))
    
    For i := 1 To Len(aDados)
        
        IncProc("Gravando: " + aDados[i][3] + "/" + aDados[i][4])
        
        SC6->(dbGoto(aDados[i][1]))
        
        If RecLock("SC6", .F.)
            SC6->C6_QTJADEV := aDados[i][8]  // QTDEV_NOVO
            SC6->(MsUnlock())
            nRecalc++
        Else
            nErros++
        EndIf
        
    Next i
    
Return
