#INCLUDE "TOTVS.ch"
#INCLUDE "topconn.ch"

/*/
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³NOVO2     º Autor ³ AP6 IDE            º Data ³  13/11/13   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDescricao ³ Codigo gerado pelo AP6 IDE.                                º±±
±±º          ³                                                            º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AP6 IDE                                                    º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
/*/

User Function RPOSPEDVEN() //


//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Declaracao de Variaveis                                             ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

Local  cDesc1         := "Este programa tem como objetivo imprimir relatorio "
Local cDesc2         := "de acordo com os parametros informados pelo usuario."
Local cDesc3         := "Relatório Posição de Pedido"
Local cPict          := ""
Local titulo       := "Relatorio de Posição Pedido de Venda"
Local nLin         := 59

Local Cabec1       := ""
//               0         1         2         3         4         5         6         7         8         9         0         1         2         3
//               01234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789
Local Cabec2 := "It Produto         Descricao                       Vend  Res  Dev  Can   NF   MI"

Local imprime      	:= .T.
Local aOrd 			:= {}
Private lEnd        := .F.
Private lAbortPrint := .F.
Private CbTxt      	:= ""
Private limite     	:= 80
Private tamanho    	:= "P"
Private nomeprog	:= "RPOSPEDVEN" // Coloque aqui o nome do programa para impressao no cabecalho
Private nTipo       := 18
Private aReturn     := { "Zebrado", 1, "Administracao", 2, 2, 1, "", 1}
Private nLastKey    := 0
Private cPerg       := "RPOSPEDVEN"
Private cbtxt      	:= Space(10)
Private cbcont     	:= 00
Private CONTFL     	:= 01
Private m_pag      	:= 01
Private wnrel      	:= "RPOSPEDVEN" // Coloque aqui o nome do arquivo usado para impressao em disco

Private cString 	:= "SC6"

dbSelectArea("SC5")
dbSetOrder(1)

dbSelectArea("SC6")
dbSetOrder(1)

//ValidPerg(cPerg)
pergunte(cPerg,.F.)

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Monta a interface padrao com o usuario...                           ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

wnrel := SetPrint(cString,NomeProg,cPerg,@titulo,cDesc1,cDesc2,cDesc3,.T.,aOrd,.T.,Tamanho,,.T.)

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
±±ºFun‡„o    ³RUNREPORT º Autor ³ AP6 IDE            º Data ³  13/11/13   º±±
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

Local  nOrdem
Local cNomVen 
Local cNomArq
Local lAglutina := .F.
Local n_PosAtu
Local n_Mi := 0  // NOVO: Variável para armazenar MI calculado

   
 FwMsgRun(,{|oSay|U_R5RESERV()},"Depurando Reservas","Aguarde...")


c_String := ""
c_String += " SELECT C6_NUM,C6_ITEM,C6_PRODUTO,B1_DESC,C6_DESCRI,C6_QTDVEN,C6_QTDEMP,C6_QTDENT,C6_QTJADEV,C6_QTJACAN,C6_QTENTMI,C6_DESCAMB, "
c_String += " C5_CLIENTE,C5_LOJACLI,C5_EMISSAO,C5_VEND1,C5_VEND4,A1_COD,A1_LOJA,A1_NOME,isnull(CONVERT(VARCHAR(8000),CONVERT(VARBINARY(8000),C6_OBSITEM)),' ')  AS C6_OBSITEM , SUM(C9_QTDLIB) AS C9_QTDLIB "
c_String += " FROM "+RETSQLNAME("SC6")+" SC6 (NOLOCK)"
c_String += " JOIN "+RETSQLNAME("SC5")+" SC5 (NOLOCK) ON SC5.C5_NUM=SC6.C6_NUM AND SC5.D_E_L_E_T_=' ' AND SC5.C5_FILIAL='"+XFILIAL("SC5")+"' "
c_String += " JOIN "+RETSQLNAME("SA1")+" SA1 (NOLOCK) ON SA1.A1_COD = SC5.C5_CLIENTE AND SC5.C5_LOJACLI = A1_LOJA AND SA1.D_E_L_E_T_=' ' AND SA1.A1_FILIAL='"+XFILIAL("SA1")+"' "
c_String += " JOIN "+RETSQLNAME("SB1")+" SB1 (NOLOCK) ON SB1.B1_COD = SC6.C6_PRODUTO AND SB1.D_E_L_E_T_=' ' AND SB1.B1_FILIAL='"+XFILIAL("SB1")+"' "
c_String += " LEFT JOIN "+RetSqlName("SC9")+" SC9 (NOLOCK) ON C6_FILIAL = C9_FILIAL AND C6_NUM = C9_PEDIDO AND C6_PRODUTO = C9_PRODUTO AND C6_ITEM = C9_ITEM AND C9_BLEST='' AND SC9.D_E_L_E_T_='' AND SC9.C9_NFISCAL = ' '"
c_String += " WHERE SC6.C6_NUM 		BETWEEN '"+MV_PAR01 +"' AND '"+MV_PAR02+"' "
c_String += " AND 	SC5.C5_CLIENTE  BETWEEN '"+MV_PAR03 +"' AND '"+MV_PAR04+"' "
c_String += " AND 	SC5.C5_VEND1	BETWEEN '"+MV_PAR05 +"' AND '"+MV_PAR06+"' "
c_String += " AND 	SC6.D_E_L_E_T_=' ' "
c_String += " AND 	SC6.C6_FILIAL='"+XFILIAL("SC6")+"' "
c_String += " GROUP BY C6_NUM,C6_ITEM,C6_PRODUTO,B1_DESC,C6_DESCRI,C6_QTDVEN,C6_QTDEMP,C6_QTDENT,C6_QTJADEV,C6_QTJACAN,C6_QTENTMI,C6_DESCAMB, "
c_String += " C5_CLIENTE,C5_LOJACLI,C5_EMISSAO,C5_VEND1,C5_VEND4,A1_COD,A1_LOJA,A1_NOME,isnull(CONVERT(VARCHAR(8000),CONVERT(VARBINARY(8000),C6_OBSITEM)),' ') "

//AJP-20210413 - Inclusão da ordem do relatório - Chamado #880
If MV_PAR08 == 1 		// Totaliza por Produto? 1 - Sim  2 - Nao
	C_string += " ORDER BY C6_NUM, C6_PRODUTO, C6_ITEM "
Else 
   If MV_PAR07 = 1    	// 1 - Ordem natural    2 - Por Produto
	  C_string += " ORDER BY C6_NUM, C6_ITEM, C6_PRODUTO "
   Else 	
	  C_string += " ORDER BY C6_NUM, C6_PRODUTO, C6_ITEM "
   EndiF
Endif 



TCQUERY c_String Alias TRBA NEW
TCSetField("TRBA", "C5_EMISSAO","D")
TCSetField("TRBA", "C6_QTDVEN","N",12,2)
TCSetField("TRBA", "C6_QTDEMP","N",12,2)
TCSetField("TRBA", "C6_QTDENT","N",12,2)
TCSetField("TRBA", "C6_QTJACAN","N",12,2)
TCSetField("TRBA", "C6_QTJADEV","N",12,2)

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ SETREGUA -> Indica quantos registros serao processados para a regua ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
SetRegua(RecCount())

TRBA->(dbGoTop())
c_Pedido := ""
n_PagPed:= 0

If MV_PAR08 == 1   	// Aglutina Produtos por Pedido

	While !TRBA->(EOF())

		n_TPedVendido	:= 0            
		n_TPedLiber		:= 0
		n_TPedDevol  	:= 0 
		n_TPedPCance  	:= 0
		n_TPedPEntre	:= 0
		n_TPedMi		:= 0

		c_Pedido := TRBA->C6_NUM

		While c_Pedido == TRBA->C6_NUM .and. !TRBA->(EOF())

			n_TotVendido	:= 0 	
			n_TotLiber		:= 0 	
			n_TotDevol		:= 0 	
			n_TotCance		:= 0 	
			n_TotEntre		:= 0 	
			n_TotMi     	:= 0

			//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
			//³ Verifica o cancelamento pelo usuario...                             ³
			//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
			If lAbortPrint
				@nLin,00 PSAY "*** CANCELADO PELO OPERADOR ***"
				Exit
			Endif

			d_Emissao := TRBA->C5_EMISSAO
			c_Cliente := TRBA->A1_COD
			c_Loja    := TRBA->A1_LOJA
			c_Nome    := TRBA->A1_NOME
			c_Vend1   := TRBA->C5_VEND1
			c_Vend4   := TRBA->C5_VEND4

			//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
			//³ Impressao do cabecalho do relatorio. . .                            ³
			//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ	
			If nLin > 55 

				n_PagPed++

   				Cabec1:=Left("PED:"+TRBA->C6_NUM +" EMIS:"+dtoc(TRBA->C5_EMISSAO)+" CLI:"+TRBA->A1_COD+"-"+TRBA->A1_LOJA+"-"+TRBA->A1_NOME,80)
				Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)

				cNomVen := POSICIONE ("SA3",1,XFILIAL("SA3")+TRBA->C5_VEND1,"A3_NOME")
				cNomArq := POSICIONE ("SA3",1,XFILIAL("SA3")+TRBA->C5_VEND4,"A3_NOME")

				@ 009,000 PSAY "Vend: "+TRBA->C5_VEND1+"-"+Left(cNomVen,20)
				@ 009,045 PSAY "Arq: "+TRBA->C5_VEND4+"-"+Left(cNomArq,20)
				@ 010,00 Psay __PrtThinLine()

				nLin := 11
			Endif
	
			c_Produto 	:= TRBA->C6_PRODUTO
			n_Cont 		:= 0

			While c_Pedido == TRBA->C6_NUM .and. c_Produto == TRBA->C6_PRODUTO .and. !Eof() 

				c_Item 		:= TRBA->C6_ITEM
				c_DescPrd 	:= SUBSTR(UPPER(TRBA->B1_DESC),1,31)
				n_Cont++

				n_TotVendido+= TRBA->C6_QTDVEN 		//Total vendido
				//n_TotLiber	+= TRBA->C6_QTDEMP 		//Total Reservada
				n_TotLiber	+= TRBA->C9_QTDLIB 		//Total Reservada
				n_TotDevol  += TRBA->C6_QTJADEV		//Total Devolvida
				n_TotCance  += TRBA->C6_QTJACAN		//Total Cancelada
				n_TotEntre	+= TRBA->C6_QTDENT 		//Total Entregue
				//n_TotMi		+= TRBA->C6_QTENTMI		//Total MI - COMENTADO: usava campo da tabela
				n_TotMi		+= TrMi(TRBA->C6_NUM, TRBA->C6_ITEM, TRBA->C6_PRODUTO)  // CORRIGIDO: Chama a função TrMi

				TRBA->(dbSkip()) // Avanca o ponteiro do registro no arquivo
		
			EndDo 

			If n_Cont > 1
				@ nLin,003 PSAY c_Produto
				@ nLin,019 Psay c_DescPrd
				@ nLin,052 PSAY Transform(n_TotVendido,"@E 999")	//QUANTIDADE VENDIDA
				@ nLin,057 PSAY Transform(n_TotLiber,"@E 999")	 	//QUANTIDADE RESERVADA
				@ nLin,062 PSAY Transform(n_TotDevol,"@E 999") 		//QUANTIDADE DEVOLVIDA
				@ nLin,067 PSAY Transform(n_TotCance,"@E 999") 		//QUANTIDADE CANCELADA
				@ nLin,072 PSAY Transform(n_TotEntre,"@E 999")		//QUANTIDADE ENTREGUE - NF 
				@ nLin,077 PSAY Transform(n_TotMi,"@E 999")			//Movimentação Interna	Adicioinado por Bruno Santos em 05/11/2017
			
				// nLin := nLin + 1 // Avanca a linha de impressao

				// @ nLin,003 PSAY "Total de Itens Aglutinados: " + Alltrim(Str(n_Cont))

				n_Cont := 0 
			Else 
			  //@ nLin,000 PSAY c_Item
				@ nLin,003 PSAY c_Produto
				@ nLin,019 Psay c_DescPrd
				@ nLin,052 PSAY Transform(n_TotVendido,"@E 999")	//QUANTIDADE VENDIDA
				@ nLin,057 PSAY Transform(n_TotLiber,"@E 999")	 	//QUANTIDADE RESERVADA
				@ nLin,062 PSAY Transform(n_TotDevol,"@E 999") 		//QUANTIDADE DEVOLVIDA
				@ nLin,067 PSAY Transform(n_TotCance,"@E 999") 		//QUANTIDADE CANCELADA
				@ nLin,072 PSAY Transform(n_TotEntre,"@E 999")		//QUANTIDADE ENTREGUE - NF 
				@ nLin,077 PSAY Transform(n_TotMi,"@E 999")			//Movimentação Interna	Adicioinado por Bruno Santos em 05/11/2017

				n_Cont := 0

			Endif 

			n_TPedVendido	+= n_TotVendido
			n_TPedLiber		+= n_TotLiber
			n_TPedDevol  	+= n_TotDevol
			n_TPedPCance  	+= n_TotCance
			n_TPedPEntre	+= n_TotEntre
			n_TPedMi		+= n_TotMi
	
			n_TotVendido	:= 0 
			n_TotLiber		:= 0 	
			n_TotDevol		:= 0 	
			n_TotCance		:= 0 	
			n_TotEntre		:= 0 	
			n_TotMi     	:= 0

			nLin := nLin + 1 // Avanca a linha de impressao
		
		EndDo 

		//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
		//³ Impressao do cabecalho do relatorio. . .                            ³
		//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ	
		If nLin > 55 
			//	If c_Pedido <> TRBA->C6_NUM
			//		n_PagPed:= 0
			//		c_Pedido := TRBA->C6_NUM
			//	EndIf
			
			n_PagPed++

			Cabec1:=Left("PED:"+c_Pedido +" EMIS:"+dtoc(d_Emissao)+" CLI:"+c_Cliente+"-"+c_Loja+"-"+c_Nome,80)
			Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)

			cNomVen := POSICIONE ("SA3",1,XFILIAL("SA3")+c_Vend1,"A3_NOME")
			cNomArq := POSICIONE ("SA3",1,XFILIAL("SA3")+c_Vend4,"A3_NOME")

			@ 009,000 PSAY "Vend: "+c_Vend1+"-"+Left(cNomVen,20)
			@ 009,045 PSAY "Arq: "+c_Vend4+"-"+Left(cNomArq,20)
			@ 010,00 Psay __PrtThinLine()

			nLin := 11
		Endif

		nLin++
		@ nLin,000 PSAY "TOTAIS DO PEDIDO "+c_Pedido+" : ----------------------->"	
		@ nLin,051 PSAY Transform(n_TPedVendido,"@E 9999")	//QUANTIDADE VENDIDA
		@ nLin,056 PSAY Transform(n_TPedLiber,"@E 9999")	//QUANTIDADE RESERVADA
		@ nLin,061 PSAY Transform(n_TPedDevol,"@E 9999")	//QUANTIDADE DEVOLVIDA
		@ nLin,066 PSAY Transform(n_TPedPCance,"@E 9999")	//QUANTIDADE CANCELADA
		@ nLin,071 PSAY Transform(n_TPedPEntre,"@E 9999")	//QUANTIDADE ENTREGUE
		@ nLin,076 PSAY Transform(n_TPedMi,"@E 9999")		//Movimentação Interna  Adicioinado por Bruno Santos em 05/11/2017		
		
		n_TPedVendido	:= 0            
		n_TPedLiber		:= 0
		n_TPedDevol  	:= 0 
		n_TPedPCance  	:= 0
		n_TPedPEntre	:= 0
		n_TPedMi		:= 0
		n_PagPed     	:= 0

		If nLin > 55 .And. c_Pedido == TRBA->C6_NUM
			//	If c_Pedido <> TRBA->C6_NUM
			//		n_PagPed:= 0
			//		c_Pedido := TRBA->C6_NUM
			//	EndIf
			
			n_PagPed++

			Cabec1:=Left("PED:"+c_Pedido +" EMIS:"+dtoc(d_Emissao)+" CLI:"+c_Cliente+"-"+c_Loja+"-"+c_Nome,80)
			Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)

			cNomVen := POSICIONE ("SA3",1,XFILIAL("SA3")+c_Vend1,"A3_NOME")
			cNomArq := POSICIONE ("SA3",1,XFILIAL("SA3")+c_Vend4,"A3_NOME")

			@ 009,000 PSAY "Vend: "+c_Vend1+"-"+Left(cNomVen,20)
			@ 009,045 PSAY "Arq: "+c_Vend4+"-"+Left(cNomArq,20)
			@ 010,00 Psay __PrtThinLine()

			nLin := 11
		Endif

   		If SC5->(Dbseek(XFILIAL("SC5")+c_Pedido))
   			If !Empty(SC5->C5_YBLQFIN) 
           		nLin := nLin + 2         
           		@nLin, 000 Psay "********************************* A T E N Ç Ã O ********************************"
           		nLin := nLin + 1
           		@nLin, 000 Psay " PEDIDO COM BLOQUEIO FINANCEIRO - Data: "+DTOC(SC5->C5_YBLQFIN)+" Usuario: "+Left(SC5->C5_YUSUBLQ,19)
           		nLin := nLin + 1         
           		@nLin, 018 Psay "********************************************************************************"
       		Endif
    	Endif   
			
		nLin++

		If c_Pedido <> TRBA->C6_NUM
			n_PagPed:= 0
			nLin := 99
		Endif 	

	EndDo

Else 

	n_TotVendido:= 0 //Total vendido
	n_TotEntre	:= 0 //Total Entregue
	n_TotLiber	:= 0 //Total Liberado
	n_TotCance	:= 0 //Total Cancelado
	n_TotDevol	:= 0 //Total Devolvido

	n_TotMi     := 0
	While !TRBA->(EOF())
	
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
	
		If nLin > 55 .or. c_Pedido <> TRBA->C6_NUM    // Salto de Página. Neste caso o formulario tem 55 linhas...
			If c_Pedido <> TRBA->C6_NUM
				n_PagPed:= 0
				c_Pedido := TRBA->C6_NUM
			EndIf
			n_PagPed++
//			Cabec1:="PEDIDO: "+TRBA->C6_NUM +"  EMISSAO: "+dtoc(TRBA->C5_EMISSAO)+"  CLIENTE: "+TRBA->A1_COD+"  "+TRBA->A1_LOJA+"  "+TRBA->A1_NOME
      		Cabec1:=Left("PED:"+TRBA->C6_NUM +" EMIS:"+dtoc(TRBA->C5_EMISSAO)+" CLI:"+TRBA->A1_COD+"-"+TRBA->A1_LOJA+"-"+TRBA->A1_NOME,80)
			Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)

			cNomVen := POSICIONE ("SA3",1,XFILIAL("SA3")+TRBA->C5_VEND1,"A3_NOME")
			cNomArq := POSICIONE ("SA3",1,XFILIAL("SA3")+TRBA->C5_VEND4,"A3_NOME")

			@ 009,000 PSAY "Vend: "+TRBA->C5_VEND1+"-"+Left(cNomVen,20)
			@ 009,045 PSAY "Arq: "+TRBA->C5_VEND4+"-"+Left(cNomArq,20)
			@ 010,00 Psay __PrtThinLine()

			nLin := 11
		Endif
	

		//	n_Mi := TRBA->C6_QTENTMI  // COMENTADO: usava campo da tabela
		//	n_Mi := TrMi(TRBA->C6_NUM+TRBA->C6_ITEM,TRBA->C6_PRODUTO)  // COMENTADO: chamada antiga com concatenação errada
		
		// NOVO: Calcula MI dinamicamente antes de imprimir
		n_Mi := TrMi(TRBA->C6_NUM, TRBA->C6_ITEM, TRBA->C6_PRODUTO)
		
		If MV_PAR07 == 1
			@ nLin,000 PSAY TRBA->C6_ITEM
		Endif 

		@ nLin,003 PSAY TRBA->C6_PRODUTO
		@ nLin,019 Psay SUBSTR(UPPER(TRBA->B1_DESC),1,31)
		@ nLin,052 PSAY Transform(TRBA->C6_QTDVEN,"@E 999")		//QUANTIDADE VENDIDA
		@ nLin,057 PSAY Transform(TRBA->C6_QTDEMP,"@E 999")		//QUANTIDADE RESERVADA
		@ nLin,062 PSAY Transform(TRBA->C6_QTJADEV,"@E 999")  	//QUANTIDADE DEVOLVIDA
		@ nLin,067 PSAY Transform(TRBA->C6_QTJACAN,"@E 999")  	//QUANTIDADE CANCELADA
		@ nLin,072 PSAY Transform(TRBA->C6_QTDENT,"@E 999")		//QUANTIDADE ENTREGUE - NF 
		//@ nLin,077 PSAY Transform(TRBA->C6_QTENTMI,"@E 999")	//Movimentação Interna - COMENTADO: usava campo da tabela
		@ nLin,077 PSAY Transform(n_Mi,"@E 999")				// CORRIGIDO: Usa variável calculada dinamicamente

		n_TotVendido+= TRBA->C6_QTDVEN //Total vendido
		n_TotEntre	+= TRBA->C6_QTDENT //Total Entregue
		n_TotLiber	+= TRBA->C6_QTDEMP //Total Liberado
		//n_TotMi		+= TRBA->C6_QTENTMI  // COMENTADO: usava campo da tabela
		n_TotMi		+= n_Mi  // CORRIGIDO: Usa variável calculada
		n_TotCance  += TRBA->C6_QTJACAN
		n_TotDevol  += TRBA->C6_QTJADEV

/*  COMENTADO: Código antigo que tratava cancelamento separadamente
	n_QtdLaco := 1
	
	If TRBA->C6_QTJACAN > 0
		n_QtdLaco := 2
	EndIf

	n_Mi := 0
	
	For i:= 1 to n_QtdLaco
		
		If i = 1 //Imprime as quantidades normais
			
			If TRBA->(C6_QTDVEN-C6_QTJACAN) >0
				n_Mi := TrMi(TRBA->C6_NUM+TRBA->C6_ITEM,TRBA->C6_PRODUTO)  
				n_TotMi += n_Mi

				@ nLin,000 PSAY TRBA->C6_ITEM
				@ nLin,003 PSAY TRBA->C6_PRODUTO
				@ nLin,019 Psay SUBSTR(UPPER(TRBA->B1_DESC),1,36)
				@ nLin,057 PSAY Transform(TRBA->(C6_QTDVEN-C6_QTJACAN),"@E 999")	//QUANTIDADE VENDIDA
				@ nLin,062 PSAY Transform(TRBA->C6_QTDEMP,"@E 999")	//QUANTIDADE RESERVADA
				@ nLin,072 PSAY Transform(TRBA->C6_QTDENT,"@E 999")	//QUANTIDADE ENTREGUE - NF 
				@ nLin,077 PSAY Transform(n_Mi,"@E 999")	//Movimentação Interna	Adicioinado por Bruno Santos em 05/11/2017
				
				n_TotVendido+= TRBA->(C6_QTDVEN-C6_QTJACAN) //Total vendido
				n_TotEntre	+= TRBA->C6_QTDENT //Total Entregue
				n_TotLiber	+= TRBA->C6_QTDEMP //Total Liberado
				
				If n_QtdLaco = 2
					nLin++
				EndIf
			EndIf
			
		Else //Imprime o Cancelado
			n_Mi := TrMi(TRBA->C6_NUM+TRBA->C6_ITEM,TRBA->C6_PRODUTO)    

			@ nLin,000 PSAY TRBA->C6_ITEM
			@ nLin,003 PSAY TRBA->C6_PRODUTO
			@ nLin,019 Psay SUBSTR(UPPER(TRBA->B1_DESC),1,36)
			@ nLin,057 PSAY Transform(TRBA->(C6_QTJACAN),"@E 999")	//QUANTIDADE CANCELADA
			@ nLin,067 PSAY Transform(TRBA->(C6_QTJACAN),"@E 999") //Transform(TRBA->C6_QTJACAN,"@E 999.999")	//QUANTIDADE CANCELADA
			@ nLin,077 PSAY Transform(n_Mi,"@E 999")	//Movimentação Interna  Adicioinado por Bruno Santos em 05/11/2017

			n_TotVendido+= TRBA->(C6_QTJACAN) //Total vendido
			n_TotCance	+= TRBA->(C6_QTJACAN) //Total Cancelado
		EndiF
		
	Next i
*/	
	
		nLin := nLin + 1 // Avanca a linha de impressao
	
		//****************************************************************************************************************
		//INCLUÍDO EM 11/09/2014, PARA TRATAR A OBSERVAÇÃO DE ITENS A PARTIR DO CAMPO C6_OBSITEM (CAMPO MEMO CUSTOMIZADO)
		//****************************************************************************************************************
		IF !EMPTY(TRBA->C6_OBSITEM)
			c_ObsItem := TRBA->C6_OBSITEM
			c_Obs	:= RTRIM(c_ObsItem)
			c_Obs	:= strtran(c_Obs,chr(13)," ")
			c_Obs	:= strtran(c_Obs,chr(10),"")
			n_TamLin	:= 60
			FOR n_PosAtu := 1 TO LEN(c_Obs)
			
				c_TextLin 	:= ""
				n_ProxPos = n_PosAtu+n_TamLin
				For n_PosAtu:=n_PosAtu to (n_ProxPos)
				
					c_TextLin += SubStr(c_Obs,n_PosAtu,1)
				
				Next n_PosAtu
			
			
				//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
				//³ Impressao do cabecalho do relatorio. . .                            ³
				//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
			
				If nLin > 55 .or. c_Pedido <> TRBA->C6_NUM    // Salto de Página. Neste caso o formulario tem 55 linhas...
					If c_Pedido <> TRBA->C6_NUM
						n_PagPed:= 0
						c_Pedido := TRBA->C6_NUM
					EndIf
					n_PagPed++
					Cabec1:=Left("PED:"+TRBA->C6_NUM +" EMIS:"+dtoc(TRBA->C5_EMISSAO)+" CLI:"+TRBA->A1_COD+"-"+TRBA->A1_LOJA+"-"+TRBA->A1_NOME,80)
					Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
					nLin := 9
				Endif
			
				c_TextLin += SubStr(c_Obs,n_PosAtu,1)
				@ nLin, 012 Psay alltrim(c_TextLin)
				nLin++
			
			Next i
		
		EndIf
	
		TRBA->(dbSkip()) // Avanca o ponteiro do registro no arquivo
	
		If c_Pedido <> TRBA->C6_NUM  .or. TRBA->(EOF())
		    nLin++
			@ nLin,000 PSAY "TOTAIS DO PEDIDO "+c_Pedido+" : ----------------------->"	
			@ nLin,051 PSAY Transform(n_TotVendido,"@E 9999")	//QUANTIDADE VENDIDA
			@ nLin,056 PSAY Transform(n_TotLiber,"@E 9999")	//QUANTIDADE RESERVADA
			@ nLin,061 PSAY Transform(n_TotDevol,"@E 9999")	//QUANTIDADE DEVOLVIDA
			@ nLin,066 PSAY Transform(n_TotCance,"@E 9999")	//QUANTIDADE CANCELADA
			@ nLin,071 PSAY Transform(n_TotEntre,"@E 9999")	//QUANTIDADE ENTREGUE
			@ nLin,076 PSAY Transform(n_TotMi,"@E 9999")	//Movimentação Interna  Adicioinado por Bruno Santos em 05/11/2017		
		
			n_TotVendido:= 0 //Total vendido
			n_TotEntre	:= 0 //Total Entregue
			n_TotLiber	:= 0 //Total Liberado
			n_TotCance	:= 0 //Total Cancelado
			n_TotDevol	:= 0 //Total Devolvido
			n_TotMi     := 0 
		
      		If SC5->(Dbseek(XFILIAL("SC5")+c_Pedido))
         		If !Empty(SC5->C5_YBLQFIN) 
            		nLin := nLin + 2         
            		@nLin, 000 Psay "********************************* A T E N Ç Ã O ********************************"
            		nLin := nLin + 1
            		@nLin, 000 Psay " PEDIDO COM BLOQUEIO FINANCEIRO - Data: "+DTOC(SC5->C5_YBLQFIN)+" Usuario: "+Left(SC5->C5_YUSUBLQ,19)
            		nLin := nLin + 1         
            		@nLin, 018 Psay "********************************************************************************"
         		Endif
      		Endif   
			nLin++

		EndIf
	
	EndDo
Endif 	

dbselectarea("TRBA")
TRBA->(DBCLOSEAREA("TRBA"))

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
±±ºPrograma  ³IMPORCA   ºAutor  ³Microsiga           º Data ³  12/31/10   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³                                                            º±±
±±º          ³                                                            º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AP                                                        º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

Static Function ValidPerg(cPerg)
_sAlias     :=     Alias()
dbSelectArea("SX1")
dbSetOrder(1)
cPerg      :=     PADR(cPerg,10)
/*
putSx1(cPerg,"01","Pedido Inicial"		,".",".","mv_ch1","C",06,0,0,"G","","SC5","","","mv_par01","","","","","","","","","","","","","","","","")
putSx1(cPerg,"02","Pedido Final"		,".",".","mv_ch2","C",06,0,0,"G","","SC5","","","mv_par02","","","","","","","","","","","","","","","","")

putSx1(cPerg,"03","Cliente Inicial"		,".",".","mv_ch3","C",06,0,0,"G","","SA1","","","mv_par03","","","","","","","","","","","","","","","","")
putSx1(cPerg,"04","Cliente Final "		,".",".","mv_ch4","C",06,0,0,"G","","SA1","","","mv_par04","","","","","","","","","","","","","","","","")

putSx1(cPerg,"05","Vendedor Inicial"	,".",".","mv_ch5","C",06,0,0,"G","","SA3","","","mv_par05","","","","","","","","","","","","","","","","")
putSx1(cPerg,"06","Vendedor Final "		,".",".","mv_ch6","C",06,0,0,"G","","SA3","","","mv_par06","","","","","","","","","","","","","","","","")
*/
u_zIncSX1(cPerg, "01", "Pedido inicial?", "MV_PAR01", "MV_CH1", "C", 06, 0, "G", "","SC5","","","","","","", "Informe o produto inicial")
u_zIncSX1(cPerg, "02", "Pedido final?"  , "MV_PAR02", "MV_CH2", "C", 06, 0, "G", "","SC5","","","","","","", "Informe o produto final")
u_zIncSX1(cPerg, "03", "Do cliente?"    , "MV_PAR03", "MV_CH3", "C", 06, 0, "G", "","SA1","","","","","","", "Informe a data inicial a ser considerada")
u_zIncSX1(cPerg, "04", "Até o cliente?" , "MV_PAR04", "MV_CH4", "C", 06, 0, "G", "","SA1","","","","","","", "Informe a média de atraso que será considerada")
u_zIncSX1(cPerg, "05", "Do cliente?"    , "MV_PAR05", "MV_CH5", "C", 06, 0, "G", "","SA1","","","","","","", "Informe a data inicial a ser considerada")
u_zIncSX1(cPerg, "06", "Até o cliente?" , "MV_PAR06", "MV_CH6", "C", 06, 0, "G", "","SA1","","","","","","", "Informe a média de atraso que será considerada")
Return




*-----------------------------------------------------------------------------
* Autor    : Bruno Santos                                           05/11/2017
*-----------------------------------------------------------------------------
* Rotina   : TrMi  - CORRIGIDA EM 20/01/2026
*-----------------------------------------------------------------------------
* Cliente  : DaLuz
*-----------------------------------------------------------------------------
* Objetivo : Trazer a quantidade da movimentação interna de cada item do pedido
*-----------------------------------------------------------------------------

Static Function TrMi(cPed,cItem,cProdb)
Local nQtdB := 0                      
Local aArea1b := GetArea()
Local cQuery := ""

*------------------------------------------------------------------------------------------------------
* Pedido de compras - QUERY CORRIGIDA
*------------------------------------------------------------------------------------------------------
cQuery   := " SELECT SUM(D3_QUANT) AS TOTAL_MI							 		" + chr(13)+chr(10)
cQuery   += "   FROM 	  " + RETSQLNAME("SD3")  + " D3 (NOLOCK)					" + chr(13)+chr(10)
cQuery   += "  WHERE D3.D_E_L_E_T_ <> '*' 											" + chr(13)+chr(10)  // CORRIGIDO: era D_E_L_E_T_ = ''
cQuery   += "    AND LEFT(D3_YPEDVEN,6)   = '" + cPed   + "'						" + chr(13)+chr(10)
cQuery   += "    AND RTRIM(LTRIM(RIGHT(D3_YPEDVEN, 2))) = RTRIM(LTRIM('" + cItem + "')) " + chr(13)+chr(10)  // CORRIGIDO: cItem estava sem aspas
cQuery   += "    AND D3_COD        = '" + cProdb + "'								" + chr(13)+chr(10)
cQuery   += "    AND D3_ESTORNO    <> 'S'							                " + chr(13)+chr(10)
cQuery   += "    AND D3_LOCAL      <> '08'	 AND D3_TM = '501'    	                " + chr(13)+chr(10)

TCQUERY cQUERY Alias BSD3 New

DbSelectArea("BSD3")
BSD3->(DbGoTop())

// CORRIGIDO: Agora usa SUM diretamente na query (mais eficiente)
If BSD3->(!Eof()) .AND. !Empty(BSD3->TOTAL_MI)
	nQtdB := BSD3->TOTAL_MI
EndIf

/* COMENTADO - Rotina antiga corrigida em 23/11/2018 (Pedro Queiroz) que fazia loop
While BSD3->(!Eof())
   nQtdB += BSD3->D3_QUANT
   BSD3->(Dbskip())
Enddo
*/

If Select("BSD3") > 0
	BSD3->(DBCloseArea())
Endif

RestArea(aArea1b)
    
Return(nQtdB)
