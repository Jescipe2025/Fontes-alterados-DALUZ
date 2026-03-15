#INCLUDE 'PROTHEUS.CH'
#INCLUDE 'RWMAKE.CH'
#INCLUDE 'FONT.CH'
#INCLUDE 'COLORS.CH'
#INCLUDE "TOPCONN.ch"

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Programa  ³         ³ Autor ³ PEDRO QUEIROZ         ³ Data ³           ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Locacao   ³                  ³Contato ³                                ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descricao ³ TRANSFERENCIA DE RESERVA ENTRE PEDIDOS                     ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³                                                            ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Retorno   ³                                                            ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Aplicacao ³                                                            ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Uso       ³                                                            ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÁÄÄÄÂÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Analista Resp.³  Data  ³                                               ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³              ³  /  /  ³                                               ³±±
±±³              ³  /  /  ³                                               ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

User function DLTRANRES() 

/*ÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±± Declaração de cVariable dos componentes                                 ±±
Ù±±ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ*/

Private cDescProd  := Space(40)
Private cCliente   := Space(35)
Private cPedido    := Space(6)
Private cProduto   := Space(15)
Private nResAtu    := 0
Private nResTrn    := 0
Private nQtdVen    := 0
Private aPedidos   := {}

/*ÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±± Declaração de Variaveis Private dos Objetos                             ±±
Ù±±ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ*/
SetPrvt("oDlg1","oSay1","oSay2","oGet1","oGrp1","oLBx1","oGrp2","oSay3","oSay4","oSay5","oSay6","oSay7")
SetPrvt("oGet3","oGet4","oGet5","oBtn1","oBtn2","oBtn3")

/*ÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±± Definicao do Dialog e todos os seus componentes.                        ±±
Ù±±ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ*/
oDlg1      := MSDialog():New( 149,372,522,950,"Transferência de Reserva Entre Pedidos",,,.F.,,,,,,.T.,,,.T. )
oSay1      := TSay():New( 008,016,{||"Produto"},oDlg1,,,.F.,.F.,.F.,.T.,CLR_BLACK,CLR_WHITE,032,008)
oSay2      := TSay():New( 017,085,{||cDescProd},oDlg1,,,.F.,.F.,.F.,.T.,CLR_BLACK,CLR_WHITE,200,008)
oGet1      := TGet():New( 016,016,{|u| If(PCount()>0,cProduto:=u,cProduto)},oDlg1,060,008,'',{||VProduto()},CLR_BLACK,CLR_WHITE,,,,.T.,"",,,.F.,.F.,,.F.,.F.,"SB1","cProduto",,)
oGrp1      := TGroup():New( 032,016,108,276,"Origem",oDlg1,CLR_BLACK,CLR_WHITE,.T.,.F. )
oGrp2      := TGroup():New( 116,016,148,276,"Destino",oDlg1,CLR_BLACK,CLR_WHITE,.T.,.F. )
oSay3      := TSay():New( 124,020,{||"Pedido"},oGrp2,,,.F.,.F.,.F.,.T.,CLR_BLACK,CLR_WHITE,032,008)
oSay4      := TSay():New( 124,175,{||"Qtd.Venda"},oGrp2,,,.F.,.F.,.F.,.T.,CLR_BLACK,CLR_WHITE,028,008)
oSay5      := TSay():New( 124,208,{||"Res.Atual"},oGrp2,,,.F.,.F.,.F.,.T.,CLR_BLACK,CLR_WHITE,028,008)
oSay6      := TSay():New( 124,240,{||"Res.Trans"},oGrp2,,,.F.,.F.,.F.,.T.,CLR_BLACK,CLR_WHITE,028,008)
oSay7      := TSay():New( 132,060,{||cCliente},oGrp2,,,.F.,.F.,.F.,.T.,CLR_BLACK,CLR_WHITE,144,008)
oGet2      := TGet():New( 132,020,{|u| If(PCount()>0,cPedido:=u,cPedido)},oGrp2,032,008,'',{||VPedido()},CLR_BLACK,CLR_WHITE,,,,.T.,"",,,.F.,.F.,,.F.,.F.,"SC5","cPedido",,)
oGet3      := TGet():New( 132,175,{|u| If(PCount()>0,nQtdVen:=u,nQtdVen)},oGrp2,024,008,'@e 9999',,CLR_BLACK,CLR_WHITE,,,,.T.,"",,{||.f.},.F.,.F.,,.F.,.F.,"","nQtdVen",,)
oGet4      := TGet():New( 132,208,{|u| If(PCount()>0,nResAtu:=u,nResAtu)},oGrp2,024,008,'@e 9999',,CLR_BLACK,CLR_WHITE,,,,.T.,"",,{||.f.},.F.,.F.,,.F.,.F.,"","nResAtu",,)
oGet5      := TGet():New( 132,240,{|u| If(PCount()>0,nResTrn:=u,nResTrn)},oGrp2,024,008,'@e 9999',{||VQtdTrn()},CLR_BLACK,CLR_WHITE,,,,.T.,"",,,.F.,.F.,,.F.,.F.,"","nResTrn",,)
oBtn1      := TButton():New( 016,240,"Buscar",oDlg1,{||Iif(Empty(cProduto),Alert("Produto tem que ser informado!"), MsAguarde({|| Buscar()},"Aguarde...","Carregando pedidos com reserva para o produto!",.T.))},037,012,,,,.T.,,"",,,,.F. )
oBtn2      := TButton():New( 160,188,"Confirmar",oDlg1,{|| BConfirma()},037,012,,,,.T.,,"",,,,.F. )
oBtn3      := TButton():New( 160,240,"Sair",oDlg1,{||oDlg1:End()},037,012,,,,.T.,,"",,,,.F. )

@ 040,020 LISTBOX oLbx1 FIELDS HEADER "Pedido","Cliente","Qtd.Venda","Res.Atual" SIZE 252,064 PIXEL ON DBLCLICK( SelPedido() ) OF oGrp1

aadd( aPedidos, {Space(6),Space(35),Transform(0,"@e 9999"),Transform(0,"@e 9999")} )

oLbx1:SetArray(aPedidos)
oLbx1:bLine := {|| {aPedidos[oLbx1:nAt,1],;
                    aPedidos[oLbx1:nAt,2],;
                    aPedidos[oLbx1:nAt,3],;
                    aPedidos[oLbx1:nAt,4]}}
oBtn2:Disable()
oGet2:Disable()
oGet3:Disable()
oGet4:Disable()
oGet5:Disable()

oDlg1:Activate(,,,.T.)

Return
//
//
//----------------------------------------------------------
Static Function VProduto()
//----------------------------------------------------------
//  Validação de produto
//----------------------------------------------------------
Limpar()

If Empty(cProduto)
	Return .T.
Endif

DbSelectArea("SB1")
DbSetOrder(1)

If !SB1->(DbSeek(xFilial("SB1")+cProduto))	
	Alert("Produto não cadastrado!")
	Return .F.
Endif

cDescProd := Left(SB1->B1_DESC,40)
oSay2:Refresh()

Return .T.
//
//     
//----------------------------------------------------------
Static Function Limpar()
//----------------------------------------------------------
//  Limpa campos
//----------------------------------------------------------
aPedidos := {}
aadd( aPedidos, {Space(6),Space(35),Transform(0,"@e 9999"),Transform(0,"@e 9999")} )

oLbx1:SetArray(aPedidos)
oLbx1:bLine := {|| {aPedidos[oLbx1:nAt,1],;
                    aPedidos[oLbx1:nAt,2],;
                    aPedidos[oLbx1:nAt,3],;
                    aPedidos[oLbx1:nAt,4]}}
oLbx1:Refresh() 
oLbx1:Disable() 
oBtn2:Disable()
oGet2:Enable()
oGet3:Enable()
oGet4:Enable()
oGet5:Enable()
cDescProd := Space(40)   
cPedido   := Space(6)   
cCliente  := Space(35)
nQtdVen   := 0
nResAtu   := 0
nResTrn   := 0
oGet2:Refresh()
oGet3:Refresh()
oGet4:Refresh()
oGet5:Refresh()
oGet2:Disable()
oGet3:Disable()
oGet4:Disable()
oGet5:Disable()
oSay2:Refresh()
oSay7:Refresh()
Return
//
//
//----------------------------------------------------------
Static Function Buscar()
//----------------------------------------------------------
//  Buscar os pedidos com reserva no produto
//----------------------------------------------------------
Local	cQuery := ""
                         
//cQuery := " SELECT C5.C5_NUM, A1.A1_NOME, SUM(C6.C6_QTDVEN) QTDVEN, SUM(C6.C6_QTDEMP) QTDRES " + chr(10)
cQuery := " SELECT C5.C5_NUM, A1.A1_NOME, SUM(C6.C6_QTDVEN) QTDVEN, SUM(C9.C9_QTDLIB) QTDRES " + chr(10)
cQuery += " FROM "+RETSQLNAME("SC6")+" C6 " + chr(10)
cQuery += " JOIN "+RetSqlName("SC9")+" C9 ON C9_FILIAL = C6_FILIAL AND C9_PEDIDO = C6_NUM AND C9_PRODUTO = C6_PRODUTO AND C9_ITEM = C6_ITEM AND C9_BLEST = '' AND C9.D_E_L_E_T_='' "
cQuery += " JOIN "+RETSQLNAME("SC5")+" C5 ON C5.C5_FILIAL=C6.C6_FILIAL AND C5.C5_NUM=C6.C6_NUM AND C5.D_E_L_E_T_=' ' AND C5.C5_STATENT='001' " + chr(10)
cQuery += " JOIN "+RETSQLNAME("SA1")+" A1 ON A1.A1_COD = C5.C5_CLIENTE AND A1.A1_LOJA = C5.C5_LOJACLI AND A1.D_E_L_E_T_=' ' " + chr(10)
cQuery += " WHERE C6_QTDEMP > 0 " + chr(10)
cQuery += "   AND C6.D_E_L_E_T_=' ' " + chr(10)
cQuery += "   AND C6.C6_PRODUTO ='"+cProduto+"' " + chr(10)
cQuery += " GROUP BY C5.C5_NUM, A1_NOME " + chr(10)
cQuery += " ORDER BY C5.C5_NUM " 
MEMOWRITE("PEDRESER.TXT",cQuery)
Tcquery cQuery Alias TMPP NEW
dbselectarea("TMPP")

aPedidos  := {}
TMPP->(DbGotop())
While TMPP->(!Eof())           
	aadd( aPedidos, {TMPP->C5_NUM, Left(TMPP->A1_NOME,35), TMPP->QTDVEN, TMPP->QTDRES} )
    TMPP->(DbSkip())
Enddo    
TMPP->(dbclosearea())

If Len(aPedidos) = 0
	Alert("Nenhum Pedido foi selecionado!")
	aadd( aPedidos, {Space(6),Space(35),Transform(0,"@e 9999"),Transform(0,"@e 9999")} )
	oGet1:SetFocus()
	Return
Endif

oLbx1:SetArray(aPedidos)
oLbx1:bLine := {|| {aPedidos[oLbx1:nAt,1],;
   	              aPedidos[oLbx1:nAt,2],;
       	           aPedidos[oLbx1:nAt,3],;
                    aPedidos[oLbx1:nAt,4]}}


oLbx1:Enable()
oLbx1:Refresh()
oLbx1:SetFocus()

Return
//
//
//----------------------------------------------------------
Static Function SelPedido()
//----------------------------------------------------------
//  Seleciona o pedido para transferencia de reserva
//----------------------------------------------------------
Local aPedAux := aClone(aPedidos)
Local nSel	  := oLbx1:nAt

aPedidos := {}
Aadd(aPedidos, aPedAux[nSel]) 

oLbx1:SetArray(aPedidos)
oLbx1:bLine := {|| {aPedidos[oLbx1:nAt,1],;
   	              aPedidos[oLbx1:nAt,2],;
       	           aPedidos[oLbx1:nAt,3],;
                    aPedidos[oLbx1:nAt,4]}}

oLbx1:Refresh()

oGet2:Enable()
oGet5:Enable()
nResTrn := aPedidos[1,4]
oGet5:Refresh()
oGet5:Disable()
oGet2:SetFocus()
Return
//
//
//----------------------------------------------------------
Static Function VPedido()
//----------------------------------------------------------
//  Validação do pedido destino
//----------------------------------------------------------
Local	cQuery := ""

oBtn2:Disable()

If Empty(cPedido)
	nResTrn := 0
	oGet5:Refresh()
	oGet5:Disable()
	oGet1:SetFocus()
	Return .T.
Endif
oGet5:Enable()
nResTrn := aPedidos[1,4]
oGet5:Refresh()
oGet5:Disable()
	
cQuery := " SELECT C5.C5_NUM, A1.A1_NOME, SUM(C6.C6_QTDVEN) QTDVEN, SUM(C6.C6_QTDENT+C6.C6_QTENTMI+ISNULL(LIB.C9_QTDLIB,0)) QTDRES " + chr(10)
//cQuery := " SELECT C5.C5_NUM, A1.A1_NOME, SUM(C6.C6_QTDVEN) QTDVEN, SUM(C9.C9_QTDLIB) QTDLIB " + chr(10)
cQuery += " FROM "+RETSQLNAME("SC6")+" C6 " + chr(10)
cQuery += " JOIN "+RETSQLNAME("SC5")+" C5 ON C5.C5_FILIAL=C6.C6_FILIAL AND C5.C5_NUM=C6.C6_NUM AND C5.D_E_L_E_T_=' ' AND C5.C5_STATENT<>'005' " + chr(10)
cQuery += " JOIN "+RETSQLNAME("SA1")+" A1 ON A1.A1_COD = C5.C5_CLIENTE AND A1.A1_LOJA = C5.C5_LOJACLI AND A1.D_E_L_E_T_=' ' " + chr(10)
cQuery += " LEFT JOIN "+RetSqlName("SC9")+" SC9 ON C6_FILIAL = SC9.C9_FILIAL AND C6_NUM = SC9.C9_PEDIDO AND C6_PRODUTO = SC9.C9_PRODUTO AND C6_ITEM = SC9.C9_ITEM AND SC9.C9_BLEST='02' AND SC9.D_E_L_E_T_='' "
cQuery += " LEFT JOIN "+RetSqlName("SC9")+" LIB ON C6_FILIAL = LIB.C9_FILIAL AND C6_NUM = LIB.C9_PEDIDO AND C6_PRODUTO = LIB.C9_PRODUTO AND C6_ITEM = LIB.C9_ITEM AND LIB.C9_BLEST='' AND LIB.D_E_L_E_T_='' "
//cQuery += " WHERE C6_QTDVEN-C6_QTDEMP > 0 " + chr(10)
cQuery += " WHERE (ISNULL(SC9.C9_QTDLIB,0) - ISNULL(LIB.C9_QTDLIB,0) <> 0 OR (SC9.C9_QTDLIB IS NULL AND LIB.C9_QTDLIB IS NULL)) " + chr(10)
cQuery += "   AND C6.D_E_L_E_T_=' ' " + chr(10)
cQuery += "   AND C6.C6_PRODUTO ='"+cProduto+"' " + chr(10)
cQuery += "   AND C6.C6_NUM ='"+cPedido+"' " + chr(10)
cQuery += " GROUP BY C5.C5_NUM, A1_NOME " + chr(10)
cQuery += " ORDER BY C5.C5_NUM " 
MEMOWRITE("PEDDEST"+cPedido+".TXT",cQuery)
Tcquery cQuery Alias TMPP NEW
dbselectarea("TMPP")

TMPP->(DbGotop())     
If TMPP->(Eof())
	Alert("Pedido não cadastrado ou não contém o produto!")
	TMPP->(DbCloseArea())
	Return .F.
Endif

nQtdVen := TMPP->QTDVEN
nResAtu := TMPP->QTDRES

If nQtdVen <= nResAtu
	Alert("Pedido não requer reservas para o produto")
	TMPP->(DbCloseArea())
	Return .F.
Endif

oGet5:Enable()

cCliente := Left(TMPP->A1_NOME,35)
oSay7:Refresh()

If nQtdVen - nResAtu < nResTrn
	nResTrn := nQtdVen - nResAtu
Endif	

oGet5:Refresh()

TMPP->(dbclosearea())

Return .T.
//
//
//----------------------------------------------------------
Static Function VQtdTrn()
//----------------------------------------------------------
//  Validação da quantidade de transferencia de reserva
//----------------------------------------------------------

If nResTrn <=0 .or. nResTrn > aPedidos[1,4] .or. nResTrn > nQTdVen - nResAtu
	Alert("Quantidade para transferência de reserva inválida")
	Return .f.
Endif

oBtn2:Enable()	
oBtn2:SetFocus()
Return .T.
//
//
//----------------------------------------------------------
Static Function BConfirma()
//----------------------------------------------------------
//  Confirmação da transferência de reserva
//----------------------------------------------------------
Local nSldTrn6 := 0
Local nSldFal 	:= 0
Local nSldAnt	:= 0
Local cChvItem	:= ""

DbSelectArea("SC6")
DbSetOrder(1)

DbSelectArea("SC9")
DbSetOrder(1)

nSldTrn6 := nResTrn
SC6->(DbSeek(xFilial("SC6") + aPedidos[1,1]))
While SC6->(C6_FILIAL + C6_NUM) == xFilial("SC6") + aPedidos[1,1] .and. SC6->(!Eof())
	If SC6->C6_PRODUTO == cProduto .and. nSldTrn6 > 0
		nSldFal := SC6->C6_QTDEMP
	   If nSldFal > nSldTrn6
			nSldFal := nSldTrn6
			nSldTrn6 := 0
		Else
			nSldTrn6 -= nSldFal
		Endif

		cChvItem := SC6->(C6_FILIAL + C6_NUM + C6_ITEM)
		SC9->(DbSeek(cChvItem))
		While SC9->(C9_FILIAL + C9_PEDIDO + C9_ITEM) == cChvItem .and. SC9->(!Eof())
			nSldAnt += SC9->C9_QTDLIB 
         SC9->(a460Estorna(.T.))
         SC9->(DbSkip())
      EndDo
     
		If nSldAnt - nSldFal > 0
			MaLibDoFat(SC6->(RecNo()),nSldAnt - nSldFal,.T.,.T.,.F.,.F.,.T.,.F.)  
		Endif
		If SC6->C6_QTDEMP = 0 .AND. !Empty(SC6->C6_YNUMPC)
			RecLock("SC6",.F.)
			SC6->C6_YNUMPC := ''
		EndIf 

	Endif
	SC6->(DbSkip())
Enddo

nSldTrn6 := nResTrn
nSldLib := 0
SC6->(DbSeek(xFilial("SC6") + cPedido))
While SC6->(C6_FILIAL + C6_NUM) == xFilial("SC6") + cPedido .and. SC6->(!Eof())
	If SC6->C6_PRODUTO == cProduto .and. nSldTrn6 > 0
		nSldLib := VerSldLib()
		nSldFal :=	SC6->(C6_QTDVEN - (C6_QTDENT+C6_QTENTMI+nSldLib))
	    If nSldFal > nSldTrn6
			nSldFal := nSldTrn6
			nSldTrn6 := 0
		Else
			nSldTrn6 -= nSldFal
		Endif

		cChvItem := SC6->(C6_FILIAL + C6_NUM + C6_ITEM)
		SC9->(DbSeek(cChvItem))
		While SC9->(C9_FILIAL + C9_PEDIDO + C9_ITEM) == cChvItem .and. SC9->(!Eof()) 
			If !Empty(SC9->C9_BLEST)
	        	SC9->(a460Estorna(.T.))
			ENDIF 
			SC9->(DbSkip())
		EndDo


		MaLibDoFat(SC6->(RecNo()),nSldFal,.T.,.T.,.F.,.F.,.T.,.F.)  

		If SC6->C6_QTDEMP = 0 .AND. !Empty(SC6->C6_YNUMPC)
			RecLock("SC6",.F.)
			SC6->C6_YNUMPC := ''
		EndIf 
	Endif
	SC6->(DbSkip())
Enddo

Alert("Transferência de Reserva concluída!")

Limpar()
oGet1:SetFocus()

Return .T.
//
//

Static Function VerSldLib()

Local cSql := ""
Local cSC9 := GetNextAlias()


cSql := " SELECT "
cSql += " 	SUM(C9_QTDLIB) AS LIB "
cSql += " FROM "
cSql += RetSqlName("SC9")
cSql += " WHERE D_E_L_E_T_='' "
cSql += " AND C9_FILIAL  = '"+SC6->C6_FILIAL+"'
cSql += " AND C9_PEDIDO  = '"+SC6->C6_NUM+"'
cSql += " AND C9_PRODUTO = '"+SC6->C6_PRODUTO+"'"
cSql += " AND C9_ITEM = '"+SC6->C6_ITEM+"'
cSql += " AND C9_BLEST = ''"

MpSysOpenQuery(cSql,(cSC9))

nRet := (cSC9)->LIB 

(cSC9)->(dbCloseArea())
Return nRet 
