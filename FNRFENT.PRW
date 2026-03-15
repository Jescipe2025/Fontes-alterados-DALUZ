#INCLUDE "TOTVS.ch"
#INCLUDE "topconn.ch"

/*/{Protheus.doc} U_FNRFENT
Função para recalcular o campo C6_QTDENT (Quantidade Entregue)
Calcula pela soma de D2_QUANT da SD2 (Itens de NF de Saída)

@type User Function
@author Adaptado
@since 08/02/2026
@version 1.0
/*/

User Function FNRFENT()
    Local aArea      := GetArea()
    Local dDtIni     := CToD("")
    Local dDtFim     := CToD("")
    Local cPedIni    := ""
    Local cPedFim    := ""
    Local cProdIni   := ""
    Local cProdFim   := ""
    Local nTipoVis   := 1  // 1=Apenas Diferenças, 2=Todos
    
    If !MsgYesNo("Esta rotina irá recalcular o campo C6_QTDENT (Quantidade Entregue)." + CRLF + ;
                 "Deseja continuar?", "Atenção")
        RestArea(aArea)
        Return
    EndIf
    
    // Monta tela de parâmetros
    If !MontaPerguntas(@dDtIni, @dDtFim, @cPedIni, @cPedFim, @cProdIni, @cProdFim, @nTipoVis)
        MsgInfo("Operação cancelada.", "Aviso")
        RestArea(aArea)
        Return
    EndIf
    
    // Processa e mostra prévia
    ProcPrevia(dDtIni, dDtFim, cPedIni, cPedFim, cProdIni, cProdFim, nTipoVis)
    
    RestArea(aArea)
    
Return

/*/{Protheus.doc} MontaPerguntas
Função para montar tela de parâmetros
@type function
/*/
Static Function MontaPerguntas(dDtIni, dDtFim, cPedIni, cPedFim, cProdIni, cProdFim, nTipoVis)
    Local lRet          := .F.
    Local oDlg
    Local oGetDtIni, oGetDtFim
    Local oGetPedIni, oGetPedFim
    Local oGetProdIni, oGetProdFim
    Local oCmbTipoVis
    Local aTipoVis      := {"Apenas Diferenças (ENTREGAS)", "Todos (ENTREGAS)"}
    Local cTipoVis      := aTipoVis[1]
    Local nOpcA         := 0
    
    // Valores padrão
    dDtIni     := FirstDate(Date())
    dDtFim     := Date()
    cPedIni    := Space(6)
    cPedFim    := Replicate("Z", 6)
    cProdIni   := Space(15)
    cProdFim   := Replicate("Z", 15)
    nTipoVis   := 1
    
    DEFINE MSDIALOG oDlg TITLE "Recálculo C6_QTDENT - Quantidade Entregue" FROM 000,000 TO 280,400 PIXEL
        
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
Processa e mostra prévia dos dados
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
        MsgInfo("Nenhum registro encontrado com os parâmetros informados.", "Aviso")
        Return
    EndIf
    
    // Define título
    cTitulo := IIF(nTipoVis == 1, "Prévia - Apenas Diferenças", "Prévia - Todos os Registros")
    
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
    cQuery += "     SC6.C6_QTDENT AS QTDENT_ATUAL, " + CRLF
    cQuery += "     ISNULL(SD2_TOTAL.TOTAL_SD2, 0) AS QTDENT_NOVO, " + CRLF
    cQuery += "     (SC6.C6_QTDENT - ISNULL(SD2_TOTAL.TOTAL_SD2, 0)) AS DIFERENCA " + CRLF
    cQuery += " FROM " + RetSqlName("SC6") + " SC6 " + CRLF
    cQuery += " INNER JOIN " + RetSqlName("SC5") + " SC5 ON " + CRLF
    cQuery += "     SC5.C5_FILIAL = SC6.C6_FILIAL " + CRLF
    cQuery += "     AND SC5.C5_NUM = SC6.C6_NUM " + CRLF
    cQuery += "     AND SC5.D_E_L_E_T_ = ' ' " + CRLF
    cQuery += " LEFT JOIN ( " + CRLF
    cQuery += "     SELECT " + CRLF
    cQuery += "         D2_FILIAL, D2_PEDIDO, D2_ITEMPV, D2_COD, " + CRLF
    cQuery += "         SUM(D2_QUANT) AS TOTAL_SD2 " + CRLF
    cQuery += "     FROM " + RetSqlName("SD2") + " " + CRLF
    cQuery += "     WHERE D_E_L_E_T_ <> '*' " + CRLF
    cQuery += "       AND D2_TIPO = 'N' " + CRLF
    cQuery += "     GROUP BY D2_FILIAL, D2_PEDIDO, D2_ITEMPV, D2_COD " + CRLF
    cQuery += " ) SD2_TOTAL ON " + CRLF
    cQuery += "     SD2_TOTAL.D2_FILIAL = SC6.C6_FILIAL " + CRLF
    cQuery += "     AND SD2_TOTAL.D2_PEDIDO = SC6.C6_NUM " + CRLF
    cQuery += "     AND SD2_TOTAL.D2_ITEMPV = SC6.C6_ITEM " + CRLF
    cQuery += "     AND SD2_TOTAL.D2_COD = SC6.C6_PRODUTO " + CRLF
    cQuery += " WHERE SC6.D_E_L_E_T_ = ' ' " + CRLF
    cQuery += "   AND SC6.C6_FILIAL = '" + xFilial("SC6") + "' " + CRLF
    cQuery += "   AND SC5.C5_EMISSAO BETWEEN '" + DToS(dDtIni) + "' AND '" + DToS(dDtFim) + "' " + CRLF
    cQuery += "   AND SC6.C6_NUM BETWEEN '" + cPedIni + "' AND '" + cPedFim + "' " + CRLF
    cQuery += "   AND SC6.C6_PRODUTO BETWEEN '" + cProdIni + "' AND '" + cProdFim + "' " + CRLF
    
    // Se for apenas diferenças
    If nTipoVis == 1
        cQuery += "   AND SC6.C6_QTDENT <> ISNULL(SD2_TOTAL.TOTAL_SD2, 0) " + CRLF
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
    TCSetField(cAlias, "QTDENT_ATUAL", "N", 12, 2)
    TCSetField(cAlias, "QTDENT_CALCULADA", "N", 12, 2)
    TCSetField(cAlias, "DIFERENCA", "N", 12, 2)
    
    While (cAlias)->(!EOF())
        AADD(aDados, {(cAlias)->RECNOSC6, ;
                      (cAlias)->C6_FILIAL, ;
                      (cAlias)->PEDIDO, ;
                      (cAlias)->ITEM, ;
                      (cAlias)->PRODUTO, ;
                      (cAlias)->QTD_VENDIDA, ;
                      (cAlias)->QTDENT_ATUAL, ;
                      (cAlias)->QTDENT_NOVO, ;
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
        
        oBrowse := TWBrowse():New(015,005,440,210,,{"Filial","Pedido","Item","Produto","Qtd Vendida","C6_QTDENT Atual","Entrega Localizada","Diferenca"},,oDlg,,,,,,,,,,,,.F.,,.T.,,.F.,,,)
        oBrowse:SetArray(aDados)
        oBrowse:bLine := {|| {aDados[oBrowse:nAt][2], ;
                              aDados[oBrowse:nAt][3], ;
                              aDados[oBrowse:nAt][4], ;
                              aDados[oBrowse:nAt][5], ;
                              Transform(aDados[oBrowse:nAt][6], "@E 999,999.99"), ;
                              Transform(aDados[oBrowse:nAt][7], "@E 999,999.99"), ;
                              Transform(aDados[oBrowse:nAt][8], "@E 999,999.99"), ;
                              Transform(aDados[oBrowse:nAt][9], "@E 999,999.99")}}
        
        // Se for diferenças, mostra botão GRAVAR
        If nTipoVis == 1
            @ 230,300 BUTTON "Gravar C6_QTDENT" SIZE 060,012 PIXEL OF oDlg ;
                ACTION (nOpcao := 1, oDlg:End())
            @ 230,370 BUTTON "Fechar" SIZE 040,012 PIXEL OF oDlg ;
                ACTION (nOpcao := 0, oDlg:End())
        Else
            @ 230,350 BUTTON "Fechar" SIZE 040,012 PIXEL OF oDlg ;
                ACTION (nOpcao := 0, oDlg:End())
        EndIf
        
    ACTIVATE MSDIALOG oDlg CENTERED
    
    // Se confirmou gravação
    If nOpcao == 1 .And. nTipoVis == 1
        If MsgYesNo("Confirma a gravacao do campo C6_QTDENT para " + cValToChar(Len(aDados)) + " registro(s)?", "Confirmacao")
            GravarC6QTDENT(aDados)
        EndIf
    EndIf
    
Return

/*/{Protheus.doc} GravarC6QTDENT
Grava o campo C6_QTDENT
@type function
/*/
Static Function GravarC6QTDENT(aDados)
    Local nRecalc := 0
    Local nErros := 0
    
    dbSelectArea("SC6")
    SC6->(dbSetOrder(1))
    
    Processa({|| ProcGrava(@aDados, @nRecalc, @nErros)}, "Gravando", "Atualizando C6_QTDENT...", .T.)
    
    MsgInfo("Processo concluído!" + CRLF + ;
            "Registros atualizados: " + cValToChar(nRecalc) + CRLF + ;
            "Erros: " + cValToChar(nErros), "Finalizado")
    
Return

/*/{Protheus.doc} ProcGrava
Processa gravação
@type function
/*/
Static Function ProcGrava(aDados, nRecalc, nErros)
    Local i
    
    ProcRegua(Len(aDados))
    
    For i := 1 To Len(aDados)
        
        IncProc("Gravando: " + aDados[i][3] + "/" + aDados[i][4])
        
        SC6->(dbGoto(aDados[i][1]))
        
        If RecLock("SC6", .F.)
            SC6->C6_QTDENT := aDados[i][8]  // QTDENT_NOVO
            SC6->(MsUnlock())
            nRecalc++
        Else
            nErros++
        EndIf
        
    Next i
    
Return
