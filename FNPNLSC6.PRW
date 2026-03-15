#INCLUDE "TOTVS.ch"
#INCLUDE "FWMVCDEF.CH"

/*/{Protheus.doc} U_PAINELSC6
Painel de Controle - Recálculo de Campos SC6
Centraliza acesso às rotinas de recálculo dos campos:
- C6_QTDENT  - Quantidade Entregue
- C6_QTENTMI - Quantidade Movimentada
- C6_QTJACAN - Quantidade Cancelada
- C6_QTJADEV - Quantidade Devolvida
- C6_QTDEMP  - Quantidade Empenhada
- C6_SLDEST  - Saldos em Estoque

@type User Function
@author Adaptado
@since 08/02/2026
@version 1.0
/*/

User Function fnpnlsc6()
    Local oDlg
    Local oPanel
    Local oFont14
    Local oFont12
    Local oFont10
    
    // Define fontes
    oFont14 := TFont():New("Arial", , 14, .T., .T.)
    oFont12 := TFont():New("Arial", , 12, .T.)
    oFont10 := TFont():New("Arial", , 10)
    
    DEFINE MSDIALOG oDlg TITLE "Painel de Recálculo SC6 - Itens de Pedido de Venda" FROM 000,000 TO 500,600 PIXEL
        
        // Painel principal
        oPanel := TPanel():New(000,000,,oDlg,,.F.,.F.,,,300,250,.T.,.F.)
        oPanel:Align := CONTROL_ALIGN_ALLCLIENT
        
        // Título
        @ 010,010 SAY "PAINEL DE RECALCULO - CAMPOS SC6" SIZE 280,015 PIXEL OF oPanel FONT oFont14 COLOR CLR_BLUE
        
        @ 030,010 SAY "_______________________________________________________________________________" SIZE 280,007 PIXEL OF oPanel
        
        // Descrição
        @ 040,010 SAY "Selecione a função desejada para recalcular os campos do SC6:" SIZE 280,010 PIXEL OF oPanel FONT oFont10
        
        // Botão 1 - FNRFENT (C6_QTDENT)
        @ 060,020 BUTTON "C6_QTDENT - Quantidade Entregue" SIZE 260,020 PIXEL OF oPanel ;
            ACTION (oDlg:End(), U_FNRFENT()) FONT oFont12
        @ 082,030 SAY "Recalcula pela soma Notas Venda (SD2 - Itens de NF Saída)" SIZE 250,007 PIXEL OF oPanel FONT oFont10 COLOR CLR_GRAY
        
        // Botão 2 - FNRFMI (C6_QTENTMI)
        @ 095,020 BUTTON "C6_QTENTMI - Quantidade Movimentada" SIZE 260,020 PIXEL OF oPanel ;
            ACTION (oDlg:End(), U_FNRFMI()) FONT oFont12
        @ 117,030 SAY "Recalcula pela soma Movimentos (SD3 - Movimentos Internos)" SIZE 250,007 PIXEL OF oPanel FONT oFont10 COLOR CLR_GRAY
        
        // Botão 3 - FNRFCAN (C6_QTJACAN)
        @ 130,020 BUTTON "C6_QTJACAN - Quantidade Cancelada" SIZE 260,020 PIXEL OF oPanel ;
            ACTION (oDlg:End(), U_FNRFCAN()) FONT oFont12
        @ 152,030 SAY "Recalcula quantidades canceladas de Movimentos (ZE4 - Cancelamentos e Devolucoes)" SIZE 250,007 PIXEL OF oPanel FONT oFont10 COLOR CLR_GRAY
        
        // Botão 4 - FNRFDEV (C6_QTJADEV)
        @ 165,020 BUTTON "C6_QTJADEV - Quantidade Devolvida" SIZE 260,020 PIXEL OF oPanel ;
            ACTION (oDlg:End(), U_FNRFDEV()) FONT oFont12
        @ 187,030 SAY "Recalcula devoluções (ZE4 - Cancelamentos e Devolucoes)" SIZE 250,007 PIXEL OF oPanel FONT oFont10 COLOR CLR_GRAY
        
        // Botão 5 - FNRFEMP (C6_QTDEMP)
        @ 200,020 BUTTON "C6_QTDEMP - Quantidade Empenhada" SIZE 260,020 PIXEL OF oPanel ;
            ACTION (oDlg:End(), U_FNRFEMP()) FONT oFont12
        @ 222,030 SAY "Recalcula empenho (SC9 - Liberações pendentes)" SIZE 250,007 PIXEL OF oPanel FONT oFont10 COLOR CLR_GRAY
        
        // Botão 6 - FNRFSLD (Saldos em Estoque)
        //@ 235,020 BUTTON "Saldos em Estoque - Consulta/Atualização" SIZE 260,020 PIXEL OF oPanel ;
        //    ACTION (oDlg:End(), U_FNRFSLD()) FONT oFont12
        //@ 257,030 SAY "Consulta e atualiza saldos em estoque (SB2 - Saldos por Armazém)" SIZE 250,007 PIXEL OF oPanel FONT oFont10 COLOR CLR_GRAY
        
        // Rodapé
        @ 275,010 SAY "_______________________________________________________________________________" SIZE 280,007 PIXEL OF oPanel
        @ 285,010 SAY "Versão 1.0 - " + DToC(Date()) SIZE 280,007 PIXEL OF oPanel FONT oFont10 COLOR CLR_GRAY
        
        // Botão Sair - Centralizado e maior
        @ 305,100 BUTTON "Sair" SIZE 100,020 PIXEL OF oPanel ACTION oDlg:End() FONT oFont12
        
    ACTIVATE MSDIALOG oDlg CENTERED
    
Return

/*/{Protheus.doc} MenuDef
Menu funcional
@type function
/*/
Static Function MenuDef()
    Local aRotina := {}
    
    ADD OPTION aRotina TITLE "C6_QTDENT - Qtd Entregue"     ACTION "U_FNRFENT()"  OPERATION 2 ACCESS 0
    ADD OPTION aRotina TITLE "C6_QTENTMI - Qtd Movimentada" ACTION "U_FNRFMI()"   OPERATION 2 ACCESS 0
    ADD OPTION aRotina TITLE "C6_QTJACAN - Qtd Cancelada"   ACTION "U_FNRFCAN()"  OPERATION 2 ACCESS 0
    ADD OPTION aRotina TITLE "C6_QTJADEV - Qtd Devolvida"   ACTION "U_FNRFDEV()"  OPERATION 2 ACCESS 0
    ADD OPTION aRotina TITLE "C6_QTDEMP - Qtd Empenhada"    ACTION "U_FNRFEMP()"  OPERATION 2 ACCESS 0
    //ADD OPTION aRotina TITLE "B2_QRESaldos em Estoque"            ACTION "U_FNRFSLD()"  OPERATION 2 ACCESS 0
    
Return aRotina
