#INCLUDE "TOTVS.ch"
#INCLUDE "topconn.ch"

/*/{Protheus.doc} U_FNRFSLD
Função para recalcular o campo B2_RESERVA (Quantidade Reservada)
Calcula pela soma de C9_QTDLIB da SC9 (Pedidos Liberados sem NF) por Produto/Armazém
Filial: 02

@type User Function
@author Adaptado
@since 09/02/2026
@version 1.0
/*/

User Function FNRFSLD()
    Local aArea      := GetArea()
    Local cProdIni   := ""
    Local cProdFim   := ""
    Local cArmIni    := ""
    Local cArmFim    := ""
    Local nTipoVis   := 1  // 1=Apenas Diferenças, 2=Todos
    
    If !MsgYesNo("Esta rotina irá recalcular o campo B2_RESERVA (Quantidade Reservada)." + CRLF + ;
                 "Filial: 02" + CRLF + ;
                 "Deseja continuar?", "Atenção")
        RestArea(aArea)
        Return
    EndIf
    
    // Monta tela de parâmetros
    If !MontaPerguntas(@cProdIni, @cProdFim, @cArmIni, @cArmFim, @nTipoVis)
        MsgInfo("Operação cancelada.", "Aviso")
        RestArea(aArea)
        Return
    EndIf
    
    // Processa e mostra prévia
    ProcPrevia(cProdIni, cProdFim, cArmIni, cArmFim, nTipoVis)
    
    RestArea(aArea)
    
Return

/*/{Protheus.doc} MontaPerguntas
Função para montar tela de parâmetros
@type function
/*/
Static Function MontaPerguntas(cProdIni, cProdFim, cArmIni, cArmFim, nTipoVis)
    Local lRet          := .F.
    Local oDlg
    Local oGetProdIni, oGetProdFim
    Local oGetArmIni, oGetArmFim
    Local oCmbTipoVis
    Local aTipoVis      := {"Apenas Diferenças (SALDOS)", "Todos (SALDO)"}
    Local cTipoVis      := aTipoVis[1]
    Local nOpcA         := 0
    
    // Valores padrão
    cProdIni   := Space(15)
    cProdFim   := Replicate("Z", 15)
    cArmIni    := Space(2)
    cArmFim    := Replicate("Z", 2)
    nTipoVis   := 1
    
    DEFINE MSDIALOG oDlg TITLE "Recálculo B2_RESERVA - Saldos Reservados (Filial 02)" FROM 000,000 TO 220,400 PIXEL
        
        @ 010,010 SAY "Produto De:" SIZE 060,007 PIXEL OF oDlg
        @ 010,080 MSGET oGetProdIni VAR cProdIni SIZE 060,010 PIXEL OF oDlg F3 "SB1"
        
        @ 025,010 SAY "Produto Até:" SIZE 060,007 PIXEL OF oDlg
        @ 025,080 MSGET oGetProdFim VAR cProdFim SIZE 060,010 PIXEL OF oDlg F3 "SB1"
        
        @ 040,010 SAY "Armazém De:" SIZE 060,007 PIXEL OF oDlg
        @ 040,080 MSGET oGetArmIni VAR cArmIni SIZE 030,010 PIXEL OF oDlg
        
        @ 055,010 SAY "Armazém Até:" SIZE 060,007 PIXEL OF oDlg
        @ 055,080 MSGET oGetArmFim VAR cArmFim SIZE 030,010 PIXEL OF oDlg
        
        @ 070,010 SAY "Visualizar:" SIZE 060,007 PIXEL OF oDlg
        @ 070,080 MSCOMBOBOX oCmbTipoVis VAR cTipoVis ITEMS aTipoVis SIZE 080,010 PIXEL OF oDlg
        
        @ 090,080 BUTTON "Confirmar" SIZE 040,012 PIXEL OF oDlg ACTION (nOpcA := 1, oDlg:End())
        @ 090,130 BUTTON "Cancelar"  SIZE 040,012 PIXEL OF oDlg ACTION (nOpcA := 0, oDlg:End())
        
    ACTIVATE MSDIALOG oDlg CENTERED
    
    If nOpcA == 1
        If Empty(cProdIni)
            cProdIni := Space(15)
        EndIf
        If Empty(cProdFim)
            cProdFim := Replicate("Z", 15)
        EndIf
        If Empty(cArmIni)
            cArmIni := Space(2)
        EndIf
        If Empty(cArmFim)
            cArmFim := Replicate("Z", 2)
        EndIf
        
        nTipoVis := aScan(aTipoVis, cTipoVis)
        
        lRet := .T.
    EndIf
    
Return lRet

/*/{Protheus.doc} ProcPrevia
Processa e mostra prévia dos dados
@type function
/*/
Static Function ProcPrevia(cProdIni, cProdFim, cArmIni, cArmFim, nTipoVis)
    Local cQuery := ""
    Local aDados := {}
    Local cTitulo := ""
    
    // Monta query
    cQuery := MontaQuery(cProdIni, cProdFim, cArmIni, cArmFim, nTipoVis)
    
    // Executa query
    Processa({|| aDados := ExecutaQuery(cQuery)}, "Aguarde", "Consultando dados...", .F.)
    
    If Len(aDados) == 0
        MsgInfo("Nenhum registro encontrado com os parâmetros informados.", "Aviso")
        Return
    EndIf
    
    // Define título
    cTitulo := IIF(nTipoVis == 1, "Prévia - Apenas Diferenças (SALDO RESERVA)", "Prévia - Todos os Registros" (SALDO RESERVA))
    
    // Mostra browse
    MostraBrowse(aDados, cTitulo, nTipoVis)
    
Return

/*/{Protheus.doc} MontaQuery
Monta a query de consulta
@type function
/*/
Static Function MontaQuery(cProdIni, cProdFim, cArmIni, cArmFim, nTipoVis)
    Local cQuery := ""
    
    cQuery := " SELECT " + CRLF
    cQuery += "     SB2.R_E_C_N_O_ AS RECNOSB2, " + CRLF
    cQuery += "     SB2.B2_FILIAL, " + CRLF
    cQuery += "     SB2.B2_COD AS PRODUTO, " + CRLF
    cQuery += "     SB2.B2_LOCAL AS ARMAZEM, " + CRLF
    cQuery += "     SB1.B1_DESC AS DESCRICAO, " + CRLF
    cQuery += "     SB2.B2_QATU AS SALDO_ATUAL, " + CRLF
    cQuery += "     SB2.B2_RESERVA AS RESERVA_ATUAL, " + CRLF
    cQuery += "     ISNULL(SC9_TOTAL.TOTAL_SC9, 0) AS RESERVA_NOVO, " + CRLF
    cQuery += "     (SB2.B2_RESERVA - ISNULL(SC9_TOTAL.TOTAL_SC9, 0)) AS DIFERENCA " + CRLF
    cQuery += " FROM " + RetSqlName("SB2") + " SB2 " + CRLF
    cQuery += " INNER JOIN " + RetSqlName("SB1") + " SB1 ON " + CRLF
    cQuery += "     SB1.B1_COD = SB2.B2_COD " + CRLF
    cQuery += "     AND SB1.D_E_L_E_T_ = ' ' " + CRLF
    cQuery += " LEFT JOIN ( " + CRLF
    cQuery += "     SELECT " + CRLF
    cQuery += "         C9_FILIAL, C9_LOCAL, C9_PRODUTO, " + CRLF
    cQuery += "         SUM(C9_QTDLIB) AS TOTAL_SC9 " + CRLF
    cQuery += "     FROM " + RetSqlName("SC9") + " " + CRLF
    cQuery += "     WHERE D_E_L_E_T_ = ' ' " + CRLF
    cQuery += "       AND C9_NFISCAL = ' ' " + CRLF
    cQuery += "     GROUP BY C9_FILIAL, C9_LOCAL, C9_PRODUTO " + CRLF
    cQuery += " ) SC9_TOTAL ON " + CRLF
    cQuery += "     SC9_TOTAL.C9_FILIAL = SB2.B2_FILIAL " + CRLF
    cQuery += "     AND SC9_TOTAL.C9_LOCAL = SB2.B2_LOCAL " + CRLF
    cQuery += "     AND SC9_TOTAL.C9_PRODUTO = SB2.B2_COD " + CRLF
    cQuery += " WHERE SB2.D_E_L_E_T_ = ' ' " + CRLF
    cQuery += "   AND SB2.B2_FILIAL = '02' " + CRLF
    cQuery += "   AND SB2.B2_COD BETWEEN '" + cProdIni + "' AND '" + cProdFim + "' " + CRLF
    cQuery += "   AND SB2.B2_LOCAL BETWEEN '" + cArmIni + "' AND '" + cArmFim + "' " + CRLF
    
    // Se for apenas diferenças
    If nTipoVis == 1
        cQuery += "   AND SB2.B2_RESERVA <> ISNULL(SC9_TOTAL.TOTAL_SC9, 0) " + CRLF
    EndIf
    
    cQuery += " ORDER BY SB2.B2_COD, SB2.B2_LOCAL " + CRLF
    
Return cQuery

/*/{Protheus.doc} ExecutaQuery
Executa query e retorna array com dados
@type function
/*/
Static Function ExecutaQuery(cQuery)
    Local aDados := {}
    Local cAlias := GetNextAlias()
    
    TCQUERY cQuery NEW ALIAS (cAlias)
    
    TCSetField(cAlias, "SALDO_ATUAL", "N", 12, 2)
    TCSetField(cAlias, "RESERVA_ATUAL", "N", 12, 2)
    TCSetField(cAlias, "RESERVA_NOVO", "N", 12, 2)
    TCSetField(cAlias, "DIFERENCA", "N", 12, 2)
    
    While (cAlias)->(!EOF())
        AADD(aDados, {(cAlias)->RECNOSB2, ;
                      (cAlias)->B2_FILIAL, ;
                      (cAlias)->PRODUTO, ;
                      (cAlias)->ARMAZEM, ;
                      (cAlias)->DESCRICAO, ;
                      (cAlias)->SALDO_ATUAL, ;
                      (cAlias)->RESERVA_ATUAL, ;
                      (cAlias)->RESERVA_NOVO, ;
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
    
    DEFINE MSDIALOG oDlg TITLE cTitulo FROM 000,000 TO 500,1000 PIXEL
        
        @ 005,005 SAY "Total de registros: " + cValToChar(Len(aDados)) SIZE 200,010 PIXEL OF oDlg
        
        oBrowse := TWBrowse():New(015,005,490,210,,{"Filial","Produto","Armazém","Descrição","Saldo Atual","B2_RESERVA Atual","Reserva Localizada","Diferença"},,oDlg,,,,,,,,,,,,.F.,,.T.,,.F.,,,)
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
            @ 230,350 BUTTON "Gravar B2_RESERVA" SIZE 060,012 PIXEL OF oDlg ;
                ACTION (nOpcao := 1, oDlg:End())
            @ 230,420 BUTTON "Fechar" SIZE 040,012 PIXEL OF oDlg ;
                ACTION (nOpcao := 0, oDlg:End())
        Else
            @ 230,410 BUTTON "Fechar" SIZE 040,012 PIXEL OF oDlg ;
                ACTION (nOpcao := 0, oDlg:End())
        EndIf
        
    ACTIVATE MSDIALOG oDlg CENTERED
    
    // Se confirmou gravação
    If nOpcao == 1 .And. nTipoVis == 1
        If MsgYesNo("Confirma a gravação do campo B2_RESERVA para " + cValToChar(Len(aDados)) + " registro(s)?", "Confirmação")
            GravarB2RESERVA(aDados)
        EndIf
    EndIf
    
Return

/*/{Protheus.doc} GravarB2RESERVA
Grava o campo B2_RESERVA
@type function
/*/
Static Function GravarB2RESERVA(aDados)
    Local nRecalc := 0
    Local nErros := 0
    
    dbSelectArea("SB2")
    SB2->(dbSetOrder(1))
    
    Processa({|| ProcGrava(@aDados, @nRecalc, @nErros)}, "Gravando", "Atualizando B2_RESERVA...", .T.)
    
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
        
        SB2->(dbGoto(aDados[i][1]))
        
        If RecLock("SB2", .F.)
            SB2->B2_RESERVA := aDados[i][8]  // RESERVA_NOVO
            SB2->(MsUnlock())
            nRecalc++
        Else
            nErros++
        EndIf
        
    Next i
    
Return
