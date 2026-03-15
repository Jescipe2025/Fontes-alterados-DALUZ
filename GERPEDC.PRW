#INCLUDE "rwmake.ch"
#INCLUDE "topconn.ch"
#INCLUDE 'COLORS.CH'
/*/
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³NOVO2     º Autor ³ AP6 IDE            º Data ³  28/01/11   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDescricao ³ Codigo gerado pelo AP6 IDE.                                º±±
±±º          ³                                                            º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AP6 IDE                                                    º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
/*/

User Function GERPEDC

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Montagem da tela de processamento.                                  ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
Private cNumPed1 := Space(6)
Private cNumPed2 := Space(6)

@ 200,1 TO 380,380 DIALOG oGeraTxt TITLE OemToAnsi("Geracao de Pedido de Compras")
@ 02,10 TO 070,180
@ 10,018 Say " Este programa ira gerar pedido de compras baseado nas vendas "
@ 18,018 Say " do período ou para um Pedido específico"
@ 26,018 Say "                                                            "

SetPrvt("oSay1","oSay2","oGet1","oGet2")

/*ÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±± Definicao do Dialog e todos os seus componentes.                        ±±
Ù±±ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ*/
//oDlg1      := MSDialog():New( 086,232,252,927,"oDlg3",,,.F.,,,,,,.T.,,,.T. )
oSay1      := TSay():New( 040,030,{||"Faixa de Pedidos"},oGeraTxt,,,.F.,.F.,.F.,.T.,CLR_BLACK,CLR_WHITE,060,008)
oGet1      := TGet():New( 048,030,{|u| If(PCount()>0,cNumPed1:=u,cNumPed1)},oGeraTxt,048,008,'',,CLR_BLACK,CLR_WHITE,,,,.T.,"",,,.F.,.F.,,.F.,.F.,"SC5","",,)
oSay1      := TSay():New( 048,080,{||"à"},oGeraTxt,,,.F.,.F.,.F.,.T.,CLR_BLACK,CLR_WHITE,032,008)
oGet2      := TGet():New( 048,085,{|u| If(PCount()>0,cNumPed2:=u,cNumPed2)},oGeraTxt,048,008,'',,CLR_BLACK,CLR_WHITE,,,,.T.,"",,,.F.,.F.,,.F.,.F.,"SC5","",,)

@ 75,128 BMPBUTTON TYPE 01 ACTION OkGeraTxt()
@ 75,158 BMPBUTTON TYPE 02 ACTION Close(oGeraTxt)

Activate Dialog oGeraTxt Centered

Return

/*/
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºFun‡„o    ³ OKGERATXTº Autor ³ AP5 IDE            º Data ³  28/01/11   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDescri‡„o ³ Funcao chamada pelo botao OK na tela inicial de processamenº±±
±±º          ³ to. Executa a geracao do arquivo texto.                    º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ Programa principal                                         º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
/*/

Static Function OkGeraTxt

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Inicializa a regua de processamento                                 ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

Processa({|| RunCont() },"Processando...")
Return

/*/
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºFun‡„o    ³ RUNCONT  º Autor ³ AP5 IDE            º Data ³  28/01/11   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDescri‡„o ³ Funcao auxiliar chamada pela PROCESSA.  A funcao PROCESSA  º±±
±±º          ³ monta a janela com a regua de processamento.               º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ Programa principal                                         º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
/*/

Static Function RunCont
Local n_Valor, ni, i 

/*
c_Query := " SELECT C5_NUM,B1_PROC,B1_PE,B1_LOJPROC,B1_COD,B1_PROC,B1_DESC,B1_UM,B1_UPRC,A2_COD,A2_LOJA,C6_YNUMPC,C5_EMISSAO AS EMISSAO,C6_ITEM,C6_QTDVEN,C6_QTDEMP,CONVERT(VARCHAR(8000),CONVERT(VARBINARY(8000),C6_OBSITEM)) AS C6_OBSITEM, C6_QTDVEN-C6_QTDEMP-C6_QTDENT - C6_QTJACAN - ISNULL(D3_QUANT,0) A_PEDIR,C9_QTDLIB,ISNULL(B2_QATU,0) AS B2_QATU,ISNULL(B2_RESERVA,0) AS B2_RESERVA,ISNULL(B2_QATU,0)-ISNULL(B2_RESERVA,0) AS SALDO_DISPONIVEL,SC6.R_E_C_N_O_ AS RECSC6,SC5.R_E_C_N_O_ AS RECSC5 "
c_Query += " FROM "+RETSQLNAME("SC6")+" SC6 (NOLOCK) "
c_Query += " LEFT JOIN "+RETSQLNAME("SC9")+ " SC9 (NOLOCK) ON C6_NUM = C9_PEDIDO AND SC9.D_E_L_E_T_='' AND SC6.C6_ITEM = SC9.C9_ITEM AND SC9.C9_FILIAL ='"+XFILIAL("SC9")+"' "
//c_Query += " JOIN "+RETSQLNAME("SC5")+" SC5 (NOLOCK) ON C6_NUM = C5_NUM AND SC5.D_E_L_E_T_='' AND C5_EMISSAO>='20110115' AND SC5.C5_YLOTCOM =''   AND SC5.C5_FILIAL ='"+XFILIAL("SC5")+"' "
c_Query += " JOIN "+RETSQLNAME("SC5")+" SC5 (NOLOCK) ON C6_NUM = C5_NUM AND SC5.D_E_L_E_T_='' AND SC5.C5_MUNPRES = '' AND SC5.C5_FILIAL ='"+XFILIAL("SC5")+"' "
c_Query += " LEFT JOIN "+RETSQLNAME("SB2")+" SB2 (NOLOCK) ON B2_COD = C6_PRODUTO AND SB2.D_E_L_E_T_='' AND SB2.B2_FILIAL='02' AND SB2.B2_LOCAL='01' AND SB2.B2_FILIAL ='"+XFILIAL("SB2")+"' "
c_Query += " JOIN "+RETSQLNAME("SB1")+" SB1 (NOLOCK) ON SB1.B1_COD = SC6.C6_PRODUTO AND SB1.D_E_L_E_T_='' AND SB1.B1_FILIAL ='"+XFILIAL("SB1")+"' "
//c_Query += " JOIN "+RETSQLNAME("SB1")+" SB1 (NOLOCK) ON SB1.B1_COD = SC6.C6_PRODUTO AND SB1.D_E_L_E_T_='' AND SB1.B1_YLINHA IN (' S') AND SB1.B1_FILIAL ='"+XFILIAL("SB1")+"' "
c_Query += " LEFT JOIN "+RETSQLNAME("SA2")+ " SA2 (NOLOCK) ON SA2.A2_COD = SB1.B1_PROC AND SB1.B1_LOJPROC = SA2.A2_LOJA AND SA2.D_E_L_E_T_='' AND SA2.A2_FILIAL ='"+XFILIAL("SA2")+"' "
c_Query += " JOIN "+RETSQLNAME("SX5")+" SX5 ON X5_TABELA='13' AND X5_CHAVE=C6_CF AND SX5.D_E_L_E_T_=' ' AND X5_DESCRI LIKE '%VEND%' AND X5_FILIAL='"+XFILIAL("SX5")+"' "

c_Query += " LEFT JOIN "+RETSQLNAME("SD3")+" SD3 (NOLOCK) ON SUBSTRING(D3_YPEDVEN,1,6) = C6_NUM AND SUBSTRING(D3_YPEDVEN,7,2) = C6_ITEM AND SD3.D_E_L_E_T_=' ' AND D3_ESTORNO<>'S' AND D3_LOCAL <> '08' AND SD3.D3_FILIAL = '"+XFILIAL("SD3")+"' "

c_Query += " WHERE SC6.D_E_L_E_T_='' "
c_Query += " AND C6_QTDEMP<>C6_QTDVEN "
c_Query += " AND C6_YNUMPC = '' "
c_Query += " AND (C6_QTDEMP+C6_QTDENT)<> C6_QTDVEN "
c_Query += " AND SC6.C6_FILIAL='"+XFILIAL("SC6")+"' "
If !Empty(cNumPed1)
	c_Query += " AND SC6.C6_NUM >= '"+cNumPed1+"' "
Endif
If !Empty(cNumPed2)
	c_Query += " AND SC6.C6_NUM <= '"+cNumPed2+"' "
Endif

c_Query += " ORDER BY A2_COD,A2_LOJA,B1_PE,C5_NUM,B1_COD "
*/

/*
c_Query := " SELECT C5_NUM,B1_PROC,B1_PE,B1_LOJPROC,B1_COD,B1_PROC,B1_DESC,B1_UM,B1_UPRC,A2_COD,A2_LOJA,C6_YNUMPC,C5_EMISSAO AS EMISSAO,C6_ITEM,C6_QTDVEN,C6_QTDEMP,CONVERT(VARCHAR(8000),CONVERT(VARBINARY(8000),C6_OBSITEM)) AS C6_OBSITEM, C6_QTDVEN-C6_QTDEMP-C6_QTDENT-C6_QTENTMI-C6_QTJACAN-C6_QTJADEV A_PEDIR,C9_QTDLIB,ISNULL(B2_QATU,0) AS B2_QATU,ISNULL(B2_RESERVA,0) AS B2_RESERVA,ISNULL(B2_QATU,0)-ISNULL(B2_RESERVA,0) AS SALDO_DISPONIVEL,SC6.R_E_C_N_O_ AS RECSC6,SC5.R_E_C_N_O_ AS RECSC5 "
c_Query += " FROM "+RETSQLNAME("SC6")+" SC6 (NOLOCK) "
c_Query += " LEFT JOIN "+RETSQLNAME("SC9")+ " SC9 (NOLOCK) ON C6_NUM = C9_PEDIDO AND SC9.D_E_L_E_T_='' AND SC6.C6_ITEM = SC9.C9_ITEM AND SC9.C9_FILIAL ='"+XFILIAL("SC9")+"' "
c_Query += " JOIN "+RETSQLNAME("SC5")+" SC5 (NOLOCK) ON C6_NUM = C5_NUM AND SC5.D_E_L_E_T_='' AND SC5.C5_MUNPRES = '' AND SC5.C5_FILIAL ='"+XFILIAL("SC5")+"' AND C5_LIBEROK IN ('S','') And C5_BLQ='' "
c_Query += " LEFT JOIN "+RETSQLNAME("SB2")+" SB2 (NOLOCK) ON B2_COD = C6_PRODUTO AND SB2.D_E_L_E_T_='' AND SB2.B2_FILIAL='02' AND SB2.B2_LOCAL='01' AND SB2.B2_FILIAL ='"+XFILIAL("SB2")+"' "
c_Query += " JOIN "+RETSQLNAME("SB1")+" SB1 (NOLOCK) ON SB1.B1_COD = SC6.C6_PRODUTO AND SB1.D_E_L_E_T_='' AND SB1.B1_FILIAL ='"+XFILIAL("SB1")+"' "
c_Query += " LEFT JOIN "+RETSQLNAME("SA2")+ " SA2 (NOLOCK) ON SA2.A2_COD = SB1.B1_PROC AND SB1.B1_LOJPROC = SA2.A2_LOJA AND SA2.D_E_L_E_T_='' AND SA2.A2_FILIAL ='"+XFILIAL("SA2")+"' "
c_Query += " JOIN "+RETSQLNAME("SX5")+" SX5 ON X5_TABELA='13' AND X5_CHAVE=C6_CF AND SX5.D_E_L_E_T_=' ' AND X5_DESCRI LIKE '%VEND%' AND X5_FILIAL='"+XFILIAL("SX5")+"' "
c_Query += " WHERE SC6.D_E_L_E_T_='' "
c_Query += " AND C6_YNUMPC = '' AND SC6.C6_QTDEMP > 0 "
c_Query += " AND C6_QTDVEN-C6_QTJACAN-C6_QTJADEV-C6_QTENTMI  <> ISNULL(SC9.C9_QTDLIB, 0)" //AND (C6_QTDEMP+C6_QTDENT+C6_QTENTMI+C6_QTJACAN+C6_QTJADEV) < C6_QTDVEN "
c_Query += " AND SC6.C6_FILIAL='"+XFILIAL("SC6")+"' "
c_Query += " AND SC6.C6_LOCAL = '01' "
If !Empty(cNumPed1)
	c_Query += " AND SC6.C6_NUM >= '"+cNumPed1+"' "
Endif
If !Empty(cNumPed2)
	c_Query += " AND SC6.C6_NUM <= '"+cNumPed2+"' "
Endif
c_Query += " ORDER BY A2_COD,A2_LOJA,B1_PE,C5_NUM,B1_COD "
*/

// Alteração feita por Sérgio Arruda em 01/06/2022
c_Query := " SELECT C5_NUM,B1_PROC,B1_PE,B1_LOJPROC,B1_COD,B1_PROC,B1_DESC,B1_UM,B1_UPRC,A2_COD,A2_LOJA,C6_YNUMPC,C5_EMISSAO AS EMISSAO,C6_ITEM,C6_QTDVEN,C6_QTDEMP,CONVERT(VARCHAR(8000),CONVERT(VARBINARY(8000),C6_OBSITEM)) AS C6_OBSITEM, C6_QTDVEN-C6_QTDENT-C6_QTENTMI-C6_QTJACAN-C6_QTJADEV-C6_QTDEMP AS A_PEDIR,C9_QTDLIB,ISNULL(B2_QATU,0) AS B2_QATU,ISNULL(B2_RESERVA,0) AS B2_RESERVA,ISNULL(B2_QATU,0)-ISNULL(B2_RESERVA,0) AS SALDO_DISPONIVEL,SC6.R_E_C_N_O_ AS RECSC6,SC5.R_E_C_N_O_ AS RECSC5 "
c_Query += " FROM "+RETSQLNAME("SC6")+" SC6 (NOLOCK) "
c_Query += " LEFT JOIN "+RETSQLNAME("SC9")+ " SC9 (NOLOCK) ON C6_NUM = C9_PEDIDO AND SC9.D_E_L_E_T_='' AND SC6.C6_ITEM = SC9.C9_ITEM AND SC9.C9_FILIAL ='"+XFILIAL("SC9")+"' "
c_Query += " JOIN "+RETSQLNAME("SC5")+" SC5 (NOLOCK) ON C6_NUM = C5_NUM AND SC5.D_E_L_E_T_='' AND SC5.C5_MUNPRES = '' AND SC5.C5_FILIAL ='"+XFILIAL("SC5")+"' AND C5_LIBEROK IN ('S','') And C5_BLQ='' "
c_Query += " LEFT JOIN "+RETSQLNAME("SB2")+" SB2 (NOLOCK) ON B2_COD = C6_PRODUTO AND SB2.D_E_L_E_T_='' AND SB2.B2_FILIAL='02' AND SB2.B2_LOCAL='01' AND SB2.B2_FILIAL ='"+XFILIAL("SB2")+"' "
c_Query += " JOIN "+RETSQLNAME("SB1")+" SB1 (NOLOCK) ON SB1.B1_COD = SC6.C6_PRODUTO AND SB1.D_E_L_E_T_='' AND SB1.B1_FILIAL ='"+XFILIAL("SB1")+"' "
c_Query += " LEFT JOIN "+RETSQLNAME("SA2")+ " SA2 (NOLOCK) ON SA2.A2_COD = SB1.B1_PROC AND SB1.B1_LOJPROC = SA2.A2_LOJA AND SA2.D_E_L_E_T_='' AND SA2.A2_FILIAL ='"+XFILIAL("SA2")+"' "
c_Query += " JOIN "+RETSQLNAME("SX5")+" SX5 ON X5_TABELA='13' AND X5_CHAVE=C6_CF AND SX5.D_E_L_E_T_=' ' AND X5_DESCRI LIKE '%VEND%' AND X5_FILIAL='"+XFILIAL("SX5")+"' "
c_Query += " WHERE SC6.D_E_L_E_T_='' "
c_Query += " AND SC6.C6_YNUMPC = '' AND SC6.C6_QTDVEN - SC6.C6_QTDENT - SC6.C6_QTJACAN - SC6.C6_QTJADEV - SC6.C6_QTENTMI - ISNULL(SC9.C9_QTDLIB,0) > 0 "
//c_Query += " AND SC6.C6_QTDVEN - SC6.C6_QTDENT - SC6.C6_QTJACAN - SC6.C6_QTJADEV - SC6.C6_QTENTMI > ISNULL(SC9.C9_QTDLIB, 0) "
c_Query += " AND SC6.C6_QTDVEN - SC6.C6_QTDENT - SC6.C6_QTJACAN - SC6.C6_QTJADEV - SC6.C6_QTENTMI - ISNULL(SC9.C9_QTDLIB,0) >= (SB2.B2_QATU - SB2.B2_RESERVA) "
//AND (C6_QTDEMP+C6_QTDENT+C6_QTENTMI+C6_QTJACAN+C6_QTJADEV) < C6_QTDVEN "
c_Query += " AND SC6.C6_FILIAL='"+XFILIAL("SC6")+"' "
c_Query += " AND SC6.C6_LOCAL = '01' "
If !Empty(cNumPed1)
	c_Query += " AND SC6.C6_NUM >= '"+cNumPed1+"' "
Endif
If !Empty(cNumPed2)
	c_Query += " AND SC6.C6_NUM <= '"+cNumPed2+"' "
Endif
c_Query += " ORDER BY A2_COD,A2_LOJA,B1_PE,C5_NUM,B1_COD "

tcquery c_Query alias TRBA NEW

aStru := SC6->(dbStruct())
nT := len(aStru)
For nI := 1 to nT
	If ( aStru[nI][2] $ 'DNL')
		TCSetField('TRBA',aStru[ni,1], aStru[ni,2],aStru[ni,3],aStru[ni,4])
	Endif
Next

MEMOWRITE("GERPEDC_SC6.TXT",c_Query)

ProcRegua(TRBA->(RECCOUNT()))

c_Query := "SELECT MAX(C5_YLOTCOM) MAXLOT FROM "+retsqlname("SC5")+" SC5 WHERE SC5.D_E_L_E_T_='' AND SC5.C5_FILIAL = '"+XFILIAL("SC5")+"' "
tcquery c_Query alias TRBB NEW
c_NumLotPC	:= TRBB->MAXLOT
c_NumLotPc := Soma1(c_NumLotPc,6)
a_ItSemFor:={}

DBSELECTAREA("TRBB")
DBCLOSEAREA()
dbSelectArea("TRBA")
TRBA->(dbGoTop())

begin transaction

While !TRBA->(EOF())
	
	cFornece		:= TRBA->A2_COD
	cLoja			:= TRBA->A2_LOJA
	nPrzEnt         := TRBA->B1_PE
	
	If Empty(cFornece) .or. Empty(cLoja)  //se o campo de fornecedor do item estiver em branco não trata este item
		AADD(a_ItSemFor,{TRBA->B1_COD,TRBA->B1_DESC,TRBA->C5_NUM,TRBA->C6_ITEM})
		TRBA->(DBSKIP())
		Loop
	Endif
	
	cNumero := U_ProxSC7(6) //   GetSX8Num("SC7","C7_NUM")
	cCond   := Posicione("SA2",1,xFilial("SA2")+cFornece+cLoja,"A2_COND")
	cCond	:= IIf(Empty(cCond),"001",cCond)
	aCab 			:= {}
	aItem 			:= {}
	
	aCab:={{"C7_NUM"	,cNumero	    	,Nil},; // Numero do Pedido
	{"C7_EMISSAO"		,dDataBase 		,Nil},; // Data de Emissao
	{"C7_FORNECE" 		,cFornece	   ,Nil},; // Fornecedor
	{"C7_LOJA"    		,cLoja	      ,Nil},; // Loja do Fornecedor
	{"C7_COND"    		,cCond		   ,Nil},; // Condicao de pagamento
	{"C7_CONTATO" 		," "				,Nil},; // Contato
	{"C7_FILENT"  		," "				,Nil}}  // Filial Entrega
	
	c_Item    	:= "00"
	a_Auxiliar:={}
	While cFornece = TRBA->A2_COD .AND. cLoja = TRBA->A2_LOJA .AND. TRBA->B1_PE = nPrzEnt .AND. !TRBA->(EOF())
		
		INCPROC()
		
		If TRBA->A_PEDIR <=0
			TRBA->(DBSKIP())
			Loop
		Endif
		
		//contados dos itens do pedido
		c_Item := Soma1(c_Item,Len(SC7->C7_ITEM))
		
		//ponto de entrada para pegar um valor do produto conforme regra da empresa
		If ExistBlock("RFATM50E")
			n_Valor := ExecBlock("RFATM50E",.F.,.F.,TRBA->B1_COD)
		Else
			n_Valor := IIf(TRBA->B1_UPRC==0,1,TRBA->B1_UPRC)
		EndIf
		
		//Este Trecho Vai checar se existe pedido manual disponível para o produto
		c_String := " SELECT C7_NUM,C7_ITEM,C7_PRODUTO,C7_NUMPED,C7_QUANT,C7_QUJE,ISNULL(A2_COD,'') AS A2_COD,ISNULL(A2_LOJA,'') AS A2_LOJA,ISNULL(A2_NOME,'') AS A2_NOME,SC7.C7_QUANT - SC7.C7_QUJE AS QTDDISPO "
		c_String += " FROM SC7990 SC7 (NOLOCK) "
		c_String += " JOIN SB1990 SB1 (NOLOCK) ON SB1.B1_COD = SC7.C7_PRODUTO AND SB1.D_E_L_E_T_='' AND SB1.B1_FILIAL='' "
		c_String += " LEFT JOIN SA2990 SA2 (NOLOCK) ON SA2.A2_COD = SB1.B1_PROC AND SA2.A2_LOJA = SB1.B1_LOJPROC AND SA2.D_E_L_E_T_='' AND SA2.A2_FILIAL='' "
		c_String += " WHERE SC7.C7_NUMPED='' "
		c_String += " AND SC7.C7_QUANT>SC7.C7_QUJE "
		c_String += " AND SC7.D_E_L_E_T_='' "
		c_String += " AND SC7.C7_FILIAL='01' "
		//************************************************************************
		//Carrega os itens do pedido de compra                                               
		//************************************************************************
		aadd(aItem,{{"C7_ITEM"	,c_Item          	,Nil},; //Numero do Item
		{"C7_PRODUTO"			,TRBA->B1_COD		,Nil},; //Codigo do Produto
		{"C7_UM"     			,TRBA->B1_UM 		,Nil},; //Unidade de Medida
		{"C7_QUANT"  			,TRBA->A_PEDIR		,Nil},; //Quantidade
		{"C7_PRECO"  			,n_Valor			,Nil},; //Preco
		{"C7_DATPRF" 			,dDataBase			,Nil},; //Data De Entrega
		{"C7_OBSITEM" 			,TRBA->C6_OBSITEM	,Nil},; //Data De Entrega		
		{"C7_NUMPED"   			,TRBA->C5_NUM  		,Nil},; // Condicao de pagamento
		{"C7_ITEMPV"   			,TRBA->C6_ITEM 		,Nil},; // Condicao de pagamento
		{"C7_FLUXO"  			,"S"		 		,Nil},; //Fluxo de Caixa (S/N)
		{"C7_LOCAL"  			,"01"				,Nil},; //Localizacao
		{"C7_TIPOOC"  			,"C"				,Nil}}) //Tipo de OC
		
		aadd(a_Auxiliar,{TRBA->C5_NUM,TRBA->C6_ITEM,TRBA->RECSC6,TRBA->RECSC5,c_Item,TRBA->A_PEDIR})
		
		dbSelectArea("TRBA")
		TRBA->(DBSKIP())
		
	EndDo
	
	If Len(aItem) >0
		
		lMsErroAuto := .f.
		MSExecAuto({|v,x,y,z| MATA120(v,x,y,z)},1,aCab,aItem,3)

//		MATA120(1,aCab,aItem,3)
		If lMsErroAuto
			RollBackSX8()
			mostraerro()
		Else
			ConfirmSX8()
			For i:=1 to len (a_Auxiliar)
				
				SC6->(DBGOTO(a_Auxiliar[i,3]))
				SC5->(DBGOTO(a_Auxiliar[i,4]))
				AtuPedVen(a_Auxiliar[i,5],a_Auxiliar[i,6],c_NumLotPc,cNumero)
				
			Next i
		EndIf
		
	EndIf
	
End

End transaction

DBSELECTAREA("TRBA")
DBCLOSEAREA()

If Len(a_ItSemFor) > 0
	RelItSemFor()
Endif

MsgBox("Processo Finalizado!")

Return


/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³GERPEDC   ºAutor  ³Microsiga           º Data ³  05/15/11   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³                                                            º±±
±±º          ³                                                            º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AP                                                        º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

Static Function AtuPedVen(cItemPc,n_ResPc,cNumLotePc,cNumPc)


DBSELECTAREA("SC6")
//SC6->(DBGOTO(TRBA->RECSC6))
RecLock("SC6",.F.)
SC6->C6_YNUMPC	:= cNumPc
SC6->C6_ITEMPC 	:= cItemPc //c_Item
SC6->C6_YRESPC	:= n_ResPc
MsUnLock()

DBSELECTAREA("SC5")
//SC5->(DBGOTO(TRBA->RECSC5))
RecLock("SC5",.F.)
SC5->C5_YLOTCOM := cNumLotePc //c_NumLotPC
MsUnLock()

Return

/*/
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³NOVO4     º Autor ³ AP6 IDE            º Data ³  15/05/11   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDescricao ³ Codigo gerado pelo AP6 IDE.                                º±±
±±º          ³                                                            º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AP6 IDE                                                    º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
/*/

//User Function NOVO4
Static Function RelItSemFor


//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Declaracao de Variaveis                                             ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

Local cDesc1         := "Este programa tem como objetivo imprimir relatorio "
Local cDesc2         := "de acordo com os parametros informados pelo usuario."
Local cDesc3         := "Itens sem Fornecedor"
//Local cPict          := ""
Local titulo       := "Itens sem Fornecedor"
Local nLin         := 80
Local Cabec1       := ""
Local Cabec2       := ""
//Local imprime      := .T.
Local aOrd := {}
Private lEnd         := .F.
Private lAbortPrint  := .F.
Private CbTxt        := ""
Private limite           := 80
Private tamanho          := "P"
Private nomeprog         := "NOME" // Coloque aqui o nome do programa para impressao no cabecalho
Private nTipo            := 18
Private aReturn          := { "Zebrado", 1, "Administracao", 2, 2, 1, "", 1}
Private nLastKey        := 0
//Private cbtxt      := Space(10)
Private cbcont     := 00
Private CONTFL     := 01
Private m_pag      := 01
Private wnrel      := "NOME" // Coloque aqui o nome do arquivo usado para impressao em disco

Private cString := ""

//dbSelectArea("")
//dbSetOrder(1)

If !Empty(cNumPed1) .or. !Empty(cNumPed2)
	Cabec1 := "Pedidos de Venda: De "+cNumPed1+" a "+cNumPed2
Else
	Cabec1 := ""	
Endif

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Monta a interface padrao com o usuario...                           ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

wnrel := SetPrint(cString,NomeProg,"",@titulo,cDesc1,cDesc2,cDesc3,.T.,aOrd,.T.,Tamanho,,.T.)

If nLastKey == 27
	Return
Endif

SetDefault(aReturn,cString)

If nLastKey == 27
	Return
Endif

nTipo := If(aReturn[4]==1,15,18)

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Processamento. RPTSTATUS monta janela com a regua de processamento. ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

RptStatus({|| RunReport(Cabec1,Cabec2,Titulo,nLin) },Titulo)
Return

/*/
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºFun‡„o    ³RUNREPORT º Autor ³ AP6 IDE            º Data ³  15/05/11   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDescri‡„o ³ Funcao auxiliar chamada pela RPTSTATUS. A funcao RPTSTATUS º±±
±±º          ³ monta a janela com a regua de processamento.               º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ Programa principal                                         º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
/*/

Static Function RunReport(Cabec1,Cabec2,Titulo,nLin)

Local i 

//dbSelectArea(cString)
//dbSetOrder(1)

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ SETREGUA -> Indica quantos registros serao processados para a regua ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

SetRegua(Len(a_ItSemFor))

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Posicionamento do primeiro registro e loop principal. Pode-se criar ³
//³ a logica da seguinte maneira: Posiciona-se na filial corrente e pro ³
//³ cessa enquanto a filial do registro for a filial corrente. Por exem ³
//³ plo, substitua o dbGoTop() e o While !EOF() abaixo pela sintaxe:    ³
//³                                                                     ³
//³ dbSeek(xFilial())                                                   ³
//³ While !EOF() .And. xFilial() == A1_FILIAL                           ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

For i:=1 to Len(a_ItSemFor)
	//While !EOF()
	
	//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
	//³ Verifica o cancelamento pelo usuario...                             ³
	//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
	
	If lAbortPrint
		@nLin,00 PSAY "*** CANCELADO PELO OPERADOR ***"
		Exit
	Endif
	
	//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
	//³ Impressao do cabecalho do relatorio. . .                            ³
	//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
	
	If nLin > 55 // Salto de Página. Neste caso o formulario tem 55 linhas...
		Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
		nLin := 8
	Endif
	
	
	// Coloque aqui a logica da impressao do seu programa...
	// Utilize PSAY para saida na impressora. Por exemplo:
	@ nLin,00 PSAY a_ItSemFor[i,1]
	@ nLin,17 PSAY a_ItSemFor[i,2]
	
	nLin := nLin + 1 // Avanca a linha de impressao
	
	//   dbSkip() // Avanca o ponteiro do registro no arquivo
Next i
//EndDo

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Finaliza a execucao do relatorio...                                 ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

SET DEVICE TO SCREEN

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Se impressao em disco, chama o gerenciador de impressao...          ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

If aReturn[5]==1
	dbCommitAll()
	SET PRINTER TO
	OurSpool(wnrel)
Endif

MS_FLUSH()

Return



/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³GERPEDC   ºAutor  ³Microsiga           º Data ³  06/10/11   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³                                                            º±±
±±º          ³                                                            º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AP                                                        º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

User Function ProxSC7(nTamStr)

c_String := " SELECT MAX(C7_NUM)+1 AS PROX "
c_String += " FROM "+retsqlname("SC7") +" SC7 (NOLOCK) "
c_String += " WHERE SC7.D_E_L_E_T_='' "
c_String += " AND SC7.C7_FILIAL='"+xfilial("SC7")+"' "
c_String += " AND LEN(LTRIM(RTRIM(C7_NUM))) ="+Alltrim(Str(nTamStr))

TcQuery c_String Alias 	TRBX NEW
DBSELECTAREA("TRBX")
cProxNum := TRBX->PROX
cProxNum := StrZero(cProxNum,nTamStr)
dbclosearea()


Return (cProxNum)
