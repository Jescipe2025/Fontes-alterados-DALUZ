#INCLUDE 'PROTHEUS.CH'
#INCLUDE 'RWMAKE.CH'
#INCLUDE 'FONT.CH'
#INCLUDE 'COLORS.CH'

/*���������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Programa  �         � Autor �                       � Data �           ���
�������������������������������������������������������������������������Ĵ��
���Locacao   �                  �Contato �                                ���
�������������������������������������������������������������������������Ĵ��
���Descricao �                                                            ���
�������������������������������������������������������������������������Ĵ��
���Parametros�                                                            ���
�������������������������������������������������������������������������Ĵ��
���Retorno   �                                                            ���
�������������������������������������������������������������������������Ĵ��
���Aplicacao �                                                            ���
�������������������������������������������������������������������������Ĵ��
���Uso       �                                                            ���
�������������������������������������������������������������������������Ĵ��
���Analista Resp.�  Data  �                                               ���
�������������������������������������������������������������������������Ĵ��
���              �  /  /  �                                               ���
���              �  /  /  �                                               ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
����������������������������������������������������������������������������*/

User Function DALSMI2
Local aCores   := {}


/*/
+----------------------------------------------------------------+
| Define Array contendo as Rotinas a executar do programa +
| ----------- Elementos contidos por dimensao ------------ +
| 1. Nome a aparecer no cabecalho +
| 2. Nome da Rotina associada +
| 3. Usado pela rotina +
| 4. Tipo de Transacao a ser efetuada +
| 1 - Pesquisa e Posiciona em um Banco de Dados +
| 2 - Simplesmente Mostra os Campos +
| 3 - Inclui registros no Bancos de Dados +
| 4 - Altera o registro corrente +
| 5 - Remove o registro corrente do Banco de Dados +
+----------------------------------------------------------------+
/*/
PRIVATE aRotina := { { OemToAnsi("Pesquisar") ,"AxPesqui", 0 , 1},; //"Pesquisar"
{ OemToAnsi("Movimentos") ,"U_DALSMI2A", 0 , 2},; //"Visualizar"
{ OemToAnsi("Legenda"),"U_LEGDALSMI"	,0,1,0 ,.F.},;
{ OemToAnsi("Ajusta Status"),"u_AJSTATPED(SC5->C5_NUM)", 0 , 2}}
//{ OemToAnsi("Incluir") ,"AxInclui", 0 , 3},; //"Incluir"
//{ OemToAnsi("Alterar") ,"AxAltera", 0 , 4},; //"Alterar"
//{ OemToAnsi("Excluir") ,"FA010Del", 0 , 5, 3} } //"Excluir"


//+----------------------------------------------------------------+
//| Define o cabecalho da tela de atualizacoes |
//+----------------------------------------------------------------+


PRIVATE cCadastro := OemToAnsi("Movimenta��o de Pedidos de Venda") //"Atualizacao de Naturezas"

aCores := {{"Empty(C5_LIBEROK).And.Empty(C5_NOTA) .And. Empty(C5_BLQ)",'ENABLE' },;		//Pedido em Aberto
{ "!Empty(C5_NOTA).Or.C5_LIBEROK=='E' .And. Empty(C5_BLQ)" ,'DISABLE'},;		   	//Pedido Encerrado
{ "!Empty(C5_LIBEROK).And.Empty(C5_NOTA).And. Empty(C5_BLQ)",'BR_AMARELO'},;
{ "C5_BLQ == '1'",'BR_AZUL'},;	//Pedido Bloquedo por regra
{ "C5_BLQ == '2'",'BR_LARANJA'}}	//Pedido Bloquedo por verba

If ExistBlock("MA410COR")
	aCores := ExecBlock("MA410COR",.F.,.F.,aCores)
Endif

//+----------------------------------------------------------------+
//| Endereca funcao Mbrowse |
//+----------------------------------------------------------------+
//mBrowse( 6, 1,22,75,"SC5")
mBrowse( 6,1,22,75,'SC5',,,,,,aCores,,,,,,,,)
Return



/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �DALSMI2   �Autor  �Microsiga           � Data �  08/03/16   ���
�������������������������������������������������������������������������͹��
���Desc.     �                                                            ���
���          �                                                            ���
�������������������������������������������������������������������������͹��
���Uso       � AP                                                        ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
User Function LEGDALSMI()
Local aCores := {}

aCores := {{"ENABLE","Pedido em Aberto"},;      //Pedido em aberto
{"DISABLE","Pedido Encerrado"},;    //Pedido encerrado
{"BR_AMARELO","Pedido Liberado"}}

If !__lPyme
	Aadd(aCores,{"BR_AZUL","Pedido"})
	Aadd(aCores,{"BR_LARANJA","Pedido Liberado"})  //Pedido Liberado
EndIf



If ExistBlock("MA410LEG")
	aCores := ExecBlock("MA410LEG",.F.,.F.,aCores)
Endif
BrwLegenda(cCadastro,"Legenda A",aCores)
Return(.T.)


User Function DALSMI2A


/*������������������������������������������������������������������������ٱ�
�� Declara��o de cVariable dos componentes                                 ��
ٱ�������������������������������������������������������������������������*/
//Local nOpc := GD_INSERT+GD_DELETE+GD_UPDATE
Local nOpc := 0 //GD_INSERT+GD_DELETE+GD_UPDATE
Local oTempXXA := FWTemporaryTable():New( "TMPXXA")
Private oTempXXB := FWTemporaryTable():New( "TMPXXB")
Private oTempXXC := FWTemporaryTable():New( "TMPXXC")
Private aCoBrPED := {}
Private aHoBrPED := {}
Private cGetNomCli := Posicione("SA1",1,XFILIAL("SA1")+SC5->C5_CLIENTE+SC5->C5_LOJACLI,"A1_NOME")
Private cGetNumPed := SC5->C5_NUM
Private noBrPED  := 0

/*������������������������������������������������������������������������ٱ�
�� Declara��o de Variaveis Private dos Objetos                             ��
ٱ�������������������������������������������������������������������������*/
SetPrvt("oFont1","oFont2","oFont3","oDlg1","oPanel1","oBtnSair","oPanel2","oBtn3","oPanel3","oBtnMov")
SetPrvt("oBtn4","oGrpSC6","oBrPED","oGrpMovItem","oBrwMovItem","oGrpMovPed","oGrpItemMov","oGrp4","oSay1")
SetPrvt("oGetNumPed","oGet2")

/*������������������������������������������������������������������������ٱ�
�� Definicao do Dialog e todos os seus componentes.                        ��
ٱ�������������������������������������������������������������������������*/
oFont1     := TFont():New( "Calibri",0,-21,,.T.,0,,700,.F.,.F.,,,,,, )
oFont2     := TFont():New( "Times New Roman",0,-24,,.T.,0,,700,.F.,.F.,,,,,, )
oFont3     := TFont():New( "Arial Rounded MT Bold",0,-13,,.F.,0,,400,.F.,.F.,,,,,, )
oDlg1      := MSDialog():New( 092,200,621,1358,"Movimento de Pedido",,,.F.,,,,,,.T.,,,.T. )
oPanel1    := TPanel():New( 216,432,"",oDlg1,,.F.,.F.,,,132,032,.T.,.F. )
oBtnSair   := TButton():New( 004,012,"Sair",oPanel1,{ || IIF(APMSGYESNO("Deseja realmente sair?"),Close(oDlg1),) },104,024,,oFont3,,.T.,,"",,,,.F. )
oPanel2    := TPanel():New( 216,012,"",oPanel1,,.F.,.F.,,,132,044,.T.,.F. )
oBtn3      := TButton():New( 012,012,"Sair",oPanel2,,104,024,,,,.T.,,"",,,,.F. )
oPanel3    := TPanel():New( 216,008,"",oDlg1,,.F.,.F.,,,132,032,.T.,.F. )
oBtnMov    := TButton():New( 004,012,"Incluir Movimento ",oPanel3,{||U_GeraMov()},104,024,,oFont3,,.T.,,"",,,,.F. )
oPanel4    := TPanel():New( 216,012,"",oPanel3,,.F.,.F.,,,132,044,.T.,.F. )

oGrpMovPed := TGroup():New( 132,008,208,300,"Movimentos do Pedido de Venda",oDlg1,CLR_BLACK,CLR_WHITE,.T.,.F. )


oGrpSC6    := TGroup():New( 036,008,128,300,"Itens do Pedido de Venda",oDlg1,CLR_BLACK,CLR_WHITE,.T.,.F. )
oGrpItemMo := TGroup():New( 132,312,208,564,"Itens do Movimento",oDlg1,CLR_BLACK,CLR_WHITE,.T.,.F. )
oGrpMovIte := TGroup():New( 036,312,128,564,"Movimentos do Item",oDlg1,CLR_BLACK,CLR_WHITE,.T.,.F. )

oGrp4      := TGroup():New( 004,008,032,564,"",oDlg1,CLR_BLACK,CLR_WHITE,.T.,.F. )
oSay1      := TSay():New( 012,020,{||"Pedido"},oGrp4,,oFont1,.F.,.F.,.F.,.T.,CLR_BLACK,CLR_WHITE,032,012)
oSay2      := TSay():New( 012,145,{||"Cliente"},oGrp4,,oFont1,.F.,.F.,.F.,.T.,CLR_BLACK,CLR_WHITE,032,012)
oGetNumPed := TGet():New( 011,056,{|u| If(PCount()>0,cGetNumPed:=u,cGetNumPed)},oGrp4,076,015,'',,CLR_RED,CLR_WHITE,oFont2,,,.T.,"",,,.F.,.F.,,.T.,.F.,"","cGetNumPed",,)
oGet2      := TGet():New( 011,181,{|u| If(PCount()>0,cGetNomCli:=u,cGetNomCli)},oGrp4,375,015,'',,CLR_RED,CLR_WHITE,oFont2,,,.T.,"",,,.F.,.F.,,.T.,.F.,"","cGetNomCli",,)
oSayDescIt := TSay():New( 060,320,{||AtuDsItem(1)},oGrpMovItem,,oFont1,.F.,.F.,.F.,.T.,CLR_RED,CLR_WHITE,236,012)
oSayCodIte := TSay():New( 048,320,{||AtuDsItem(2)},oGrpMovItem,,oFont1,.F.,.F.,.F.,.T.,CLR_RED,CLR_WHITE,236,012)


DbSelectArea("SC6")
dbSetOrder(1)
If !SC6->(DBSEEK(XFILIAL("SC6")+SC5->C5_NUM))
	ApMsgInfo("ITENS DO PEDIDO N�O LOCALIZADOS!")
	Return()
EndIf
cSeek  := xFilial("SC6")+SC5->C5_NUM
bWhile := {|| C6_FILIAL+C6_NUM }

FillGetDados(2,"SC6",1,cSeek,bWhile,/*{{bCond,bAction1,bAction2}}*/,/*aNoFields*/,/*aYesFields*/,/*lOnlyYes*/,/*cQuery*/,/*bMontCols*/,.F.,aHoBrPED,aCoBrPED,/*{|| AfterCols(cArqQry) }*/,/*bBeforeCols*/,/*bAfterHeader*/,"SC6")
oBrPED	:= MsNewGetDados():New(056,016,120,292,nOpc,'AllwaysTrue()','AllwaysTrue()','',,0,99,'AllwaysTrue()','','AllwaysTrue()',oGrpSC6,aHoBrPED,aCoBrPED )
oBrPED	:oBrowse:bChange    := {||CarrItPED(oBrPED:aCols[oBrPED:nAt,1] )}

//**************************************************
//cria e carrega mmovimentos dos itens dos pedidos
//**************************************************
a_CpsITPED := {}
CriaITPED(oTempXXA)
DbSelectArea("TMPXXA")
oBrwMovIte := MsSelect():New( "TMPXXA","","",a_CpsITPED,.F.,,{072,320,120,556},,, oGrpMovItem )
CarrItPED(oBrPED:aCols[oBrPED:nAt,1] )
//***************************************************************************************************


//**************************************************
//cria e carrega Movimentos do Pedido
//**************************************************
a_CpsMovPed := {}
CriaMvPed(oTempXXB)
DbSelectArea("TMPXXB")
oBwMov     := MsSelect():New( "TMPXXB","","",a_CpsMovPed,.F.,,{148,016,204,292},,, oGrpMovPed )
CarrMvPED(oTempXXB)
oBwMov:oBrowse:bChange := {||CarrItMov(oTempXXB,oTempXXC)}

MENU oMenu POPUP
MENUITEM "Estornar MI" 				ACTION (EstMI())
MENUITEM "Transformar MI em NF"		ACTION (GerNF())
ENDMENU
oBwMov:oBrowse:BRClicked :=  { |oObject,nX,nY| oMenu:Activate( nX, (nY-20), oObject ) }




//**************************************************
//cria e carrega Itens Movimentos Movimento do Pedido
//**************************************************
a_CpsItMov := {}
CriaItMov(oTempXXC)
DbSelectArea("TMPXXC")
oBwITMov   := MsSelect():New( "TMPXXC","","",a_CpsItMov,.F.,,{148,320,204,556},,, oGrpItemMov )
CarrItMov(oTempXXB,oTempXXC)

oDlg1:Activate(,,,.T.)

dbselectarea("TMPXXA")
oTempXXA:Delete()

dbselectarea("TMPXXB")
oTempXXB:Delete()

dbselectarea("TMPXXC")
oTempXXC:Delete()


Return

/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �DALSMI2   �Autor  �Microsiga           � Data �  08/02/16   ���
�������������������������������������������������������������������������͹��
���Desc.     �                                                            ���
���          �                                                            ���
�������������������������������������������������������������������������͹��
���Uso       � AP                                                        ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Static Function CriaITPED(oTempXXA)

aStruct := {}
AADD(aStruct,{"TPDOC","C",2,0})
AADD(aStruct,{"DTMOV","D",8,0})
AADD(aStruct,{"QUANTIDADE","N",12,0})
AADD(aStruct,{"ENTSAI","C",2,0})
AADD(aStruct,{"DOCUMEN","C",15,0})

AADD(a_CpsITPED,{"DTMOV",,"Dt Mov", })
AADD(a_CpsITPED,{"QUANTIDADE",,"Quant","@E 999,999" })
AADD(a_CpsITPED,{"TPDOC",,"Tp Doc","@!" })
AADD(a_CpsITPED,{"DOCUMEN",,"Doc","@!" })

// CORRIGIDO: era oTempXXA:oSetFields( aStruct ) � metodo nao existe
oTempXXA:SetFields( aStruct )
oTempXXA:AddIndex("DTMOV", {"DTMOV"} )

//------------------
//Cria��o da tabela
//------------------
// CORRIGIDO: era oTempTXXA:Create() � nome do objeto com typo
oTempXXA:Create()

dbselectarea("TMPXXA")

Return()


/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �DALSMI2   �Autor  �Microsiga           � Data �  08/02/16   ���
�������������������������������������������������������������������������͹��
���Desc.     �                                                            ���
���          �                                                            ���
�������������������������������������������������������������������������͹��
���Uso       � AP                                                        ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Static Function CarrItPED(c_Item)
Local cArea := "TTTXX"
Default c_Item :=  "01"


If Type("oBrwMovIte") == "U"
	Return
EndIf

DbSelectArea("TMPXXA")
DbGoTop()
While !EOF()
    DbDelete()
    DbSkip()
EndDo
DbGoTop()

//***************************************************************************
//IRA CARREGAROS MOVIMENTOS DO ITEM ATRAV�S DAS NOTAS FISCAIS DE SA�DA (SD2)
//***************************************************************************
c_String := " SELECT * FROM "+ RETSQLNAME("SD2")+" WHERE D_E_L_E_T_=' ' AND D2_PEDIDO='"+SC5->C5_NUM+"' AND D2_ITEMPV = '"+c_Item+"' AND  D2_FILIAL='"+xfilial("SD2")+"' "
DbUseArea(.T.,"TOPCONN",TcGenQry(,,c_String ),cArea,.T.,.T.)
dbselectarea(cArea)
While !Eof()
	dbselectarea("TMPXXA")
	RecLock("TMPXXA",.T.)
	TMPXXA->TPDOC 		:= "NF"
	TMPXXA->DTMOV 		:= STOD((cArea)->D2_EMISSAO)
	TMPXXA->QUANTIDADE 	:= (cArea)->D2_QUANT
	TMPXXA->ENTSAI		:= "S"
	TMPXXA->DOCUMEN 	:= (cArea)->(D2_DOC + "/"+D2_SERIE)
	MsUnLock()
	dbselectarea(cArea)
	dbskip()
EndDo
DBSELECTAREA(cArea)
(cArea)->(DbCloseArea())
//***************************************************************************


//***************************************************************************
//IRA CARREGAROS MOVIMENTOS DO ITEM ATRAV�S DAS MOVIMENTA��ES INTERNAS (SD3)
//***************************************************************************
c_String := " SELECT * FROM "+ RETSQLNAME("SD3")+" SD3 WHERE SD3.D_E_L_E_T_=' ' AND D3_ESTORNO <>'S' AND D3_LOCAL <> '08' AND D3_YPEDVEN ='"+SC5->C5_NUM+c_Item+"' AND  D3_FILIAL='"+xfilial("SD3")+"' "
DbUseArea(.T.,"TOPCONN",TcGenQry(,,c_String ),cArea,.T.,.T.)
dbselectarea(cArea)
While !Eof()
	dbselectarea("TMPXXA")
	RecLock("TMPXXA",.T.)
	TMPXXA->TPDOC 		:= "MI"
	TMPXXA->DTMOV 		:= STOD((cArea)->D3_EMISSAO)
	TMPXXA->QUANTIDADE 	:= (cArea)->D3_QUANT
	TMPXXA->ENTSAI		:= IIF((cArea)->D3_TM <="499","E", "S")
	TMPXXA->DOCUMEN 	:= (cArea)->(D3_DOC)
	MsUnLock()
	dbselectarea(cArea)
	dbskip()
EndDo
DBSELECTAREA(cArea)
(cArea)->(DbCloseArea())
//***************************************************************************


DBSELECTAREA("TMPXXA")
DBGOTOP()

AtuDsItem(1)
oSayDescIt:Refresh()
AtuDsItem(2)
oSayCodIte:Refresh()

oBrwMovIte:oBrowse:Refresh()

Return


/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �DALSMI2   �Autor  �Microsiga           � Data �  08/02/16   ���
�������������������������������������������������������������������������͹��
���Desc.     �                                                            ���
���          �                                                            ���
�������������������������������������������������������������������������͹��
���Uso       � AP                                                        ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Static Function AtuDsItem(nTipo)
Local c_Retorno := ""
Local nPosPrd

If Type("oBrPED") == "U"
	Return(c_Retorno)
EndIf

nPosPrd 	:=	aScan(oBrPED:aHeader,{|x| Alltrim(x[2]) == "C6_PRODUTO"})
nPosItem 	:=	aScan(oBrPED:aHeader,{|x| Alltrim(x[2]) == "C6_ITEM"})


c_Item := oBrPED:aCols[oBrPED:nAt,nPosItem]
c_Prod := RTRIM(oBrPED:aCols[oBrPED:nAt,nPosPrd])
c_Desc := RTRIM(Posicione("SB1",1,XFILIAL("SB1")+c_Prod,"B1_DESC"))

If nTipo == 1
	c_Retorno := "["+c_Desc+"]"
Else
	c_Retorno := "["+c_Item+" - "+c_Prod+"]"
EndIf

Return(c_Retorno)




/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �DALSMI2   �Autor  �Microsiga           � Data �  08/02/16   ���
�������������������������������������������������������������������������͹��
���Desc.     �                                                            ���
���          �                                                            ���
�������������������������������������������������������������������������͹��
���Uso       � AP                                                        ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Static Function CriaMvPed(oTempXXB)

aStruct := {}
AADD(aStruct,{"TPDOC","C",2,0})
AADD(aStruct,{"DTMOV","D",8,0})
AADD(aStruct,{"ENTSAI","C",2,0})
AADD(aStruct,{"DOCUMEN","C",15,0})
AADD(aStruct,{"TM","C",TAMSX3('D3_TM')[01],0})
AADD(aStruct,{"ESTORNO","C",TAMSX3('D3_ESTORNO')[01],0})


AADD(a_CpsMovPed,{"DTMOV",,"Dt Mov", })
AADD(a_CpsMovPed,{"TPDOC",,"Tp Doc","@!" })
AADD(a_CpsMovPed,{"DOCUMEN",,"Doc","@!" })

// CORRIGIDO: era oTempXXB:oSetFields( aStruct ) � metodo nao existe
oTempXXB:SetFields( aStruct )
oTempXXB:AddIndex("ENTSAI", {"ENTSAI"} )

//------------------
//Cria��o da tabela
//------------------
// CORRIGIDO: era oTempTXXB:Create() � nome do objeto com typo
oTempXXB:Create()
Return()



/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �DALSMI2   �Autor  �Microsiga           � Data �  08/02/16   ���
�������������������������������������������������������������������������͹��
���Desc.     �                                                            ���
���          �                                                            ���
�������������������������������������������������������������������������͹��
���Uso       � AP                                                        ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Static Function CarrMvPED(oTempXXB)
Local cArea := "TTTXX"

DbSelectArea("TMPXXB")
// CORRIGIDO: era oTempXXB:FWTemporaryTable():Zap() � instanciava nova classe ao inves de chamar no objeto
oTempXXB:Zap()

//**********************************************************************
//IR� CARREGAR AS NOTAS FISCAIS DE SA�DA

c_String := " SELECT DISTINCT D2_DOC,D2_SERIE,D2_EMISSAO FROM "+ RETSQLNAME("SD2")+" WHERE D_E_L_E_T_=' ' AND D2_PEDIDO='"+SC5->C5_NUM+"' AND  D2_FILIAL='"+xfilial("SD2")+"' "
DbUseArea(.T.,"TOPCONN",TcGenQry(,,c_String ),cArea,.T.,.T.)
dbselectarea(cArea)
While !Eof()
	dbselectarea("TMPXXB")
	RecLock("TMPXXB",.T.)
	TMPXXB->TPDOC 		:= "NF"
	TMPXXB->DTMOV 		:= STOD((cArea)->D2_EMISSAO)
	TMPXXB->ENTSAI		:= "S"
	TMPXXB->DOCUMEN 	:= (cArea)->(D2_DOC + "/"+D2_SERIE)
	MsUnLock()
	dbselectarea(cArea)
	dbskip()
EndDo

DBSELECTAREA(cArea)
(cArea)->(DbCloseArea())


//******************************************************
//IR� CARREGAR OS MOVIMENTOS INTERNOS SD3
c_String := " SELECT DISTINCT D3_DOC,D3_EMISSAO,D3_TM,D3_ESTORNO FROM "+ RETSQLNAME("SD3")+" SD3 WHERE SD3.D_E_L_E_T_=' ' AND D3_ESTORNO <>'S' AND D3_LOCAL <> '08' AND LEFT(D3_YPEDVEN,6)='"+SC5->C5_NUM+"' AND  D3_FILIAL='"+xfilial("SD3")+"' "
DbUseArea(.T.,"TOPCONN",TcGenQry(,,c_String ),cArea,.T.,.T.)
dbselectarea(cArea)
While !Eof()
	dbselectarea("TMPXXB")
	RecLock("TMPXXB",.T.)
	TMPXXB->TPDOC 		:= "MI"
	TMPXXB->TM	 		:= (cArea)->(D3_TM)
	TMPXXB->DTMOV 		:= STOD((cArea)->D3_EMISSAO)
	TMPXXB->ENTSAI		:= IIF((cArea)->(D3_TM) >= "500", "S","E")
	TMPXXB->DOCUMEN 	:= (cArea)->(D3_DOC)
	TMPXXB->ESTORNO 	:= (cArea)->(D3_ESTORNO)
	MsUnLock()
	dbselectarea(cArea)
	dbskip()
EndDo

DBSELECTAREA(cArea)
(cArea)->(DbCloseArea())


DBSELECTAREA("TMPXXB")
DBGOTOP()

oBwMov:oBrowse:Refresh()

Return


/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �DALSMI2   �Autor  �Microsiga           � Data �  08/02/16   ���
�������������������������������������������������������������������������͹��
���Desc.     �                                                            ���
���          �                                                            ���
�������������������������������������������������������������������������͹��
���Uso       � AP                                                        ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Static Function CriaItMov(oTempXXC)

aStruct := {}
AADD(aStruct,{"COD","C", TamSX3('B1_COD')[01],0})
AADD(aStruct,{"ITEMPV","C", TamSX3('C6_ITEM')[01],0})
AADD(aStruct,{"QUANTIDADE","N",12,0})
AADD(aStruct,{"DESCRI","C", TamSX3('B1_DESC')[01],0})
AADD(aStruct,{"PEDIDO","C",TamSX3('C6_NUM')[01],0})

AADD(a_CpsItMov,{"ITEMPV",,"Item", })
AADD(a_CpsItMov,{"COD",,"Cod", })
AADD(a_CpsItMov,{"QUANTIDADE",,"Quant","@E 999,999" })
AADD(a_CpsItMov,{"DESCRI",,"Descricao", })

// CORRIGIDO: era oTempXXB:oSetFields( aStruct ) � objeto errado (XXB vs XXC) E metodo errado
oTempXXC:SetFields( aStruct )
oTempXXC:AddIndex("ITEMPV", {"ITEMPV"} )

//------------------
//Cria��o da tabela
//------------------
// CORRIGIDO: era oTempTXXB:Create() � objeto errado (XXB vs XXC) E typo no nome
oTempXXC:Create()
dbselectarea("TMPXXC")

Return



/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �DALSMI2   �Autor  �Microsiga           � Data �  08/02/16   ���
�������������������������������������������������������������������������͹��
���Desc.     �                                                            ���
���          �                                                            ���
�������������������������������������������������������������������������͹��
���Uso       � AP                                                        ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Static Function CarrItMov(oTempXXB,oTempXXC)
Local cArea := "TTTXX"

DbSelectArea("TMPXXC")
// CORRIGIDO: era oTempXXC : FWTemporaryTable():Zap() � instanciava nova classe ao inves de chamar no objeto
oTempXXC:Zap()

If TMPXXB->TPDOC == "NF"
	
	c_Documen 	:= rtrim(Substr(TMPXXB->DOCUMEN,1,At("/",TMPXXB->DOCUMEN)-1))
	c_Serie		:= rtrim(Substr(TMPXXB->DOCUMEN,At("/",TMPXXB->DOCUMEN)+1,10))
	c_String := " SELECT * FROM "+ RETSQLNAME("SD2")+" WHERE D_E_L_E_T_=' 'AND D2_DOC = '"+c_Documen+"' and D2_SERIE = '"+c_Serie+"'  AND D2_PEDIDO='"+SC5->C5_NUM+"' AND  D2_FILIAL='"+xfilial("SD2")+"' "
	DbUseArea(.T.,"TOPCONN",TcGenQry(,,c_String ),cArea,.T.,.T.)
	dbselectarea(cArea)
	While !Eof()
		dbselectarea("TMPXXC")
		RecLock("TMPXXC",.T.)
		TMPXXC->ITEMPV  	:= (cArea)->(D2_ITEMPV)
		TMPXXC->PEDIDO  	:= (cArea)->(D2_PEDIDO)
		TMPXXC->QUANTIDADE 	:= (cArea)->D2_QUANT
		TMPXXC->COD		 	:= (cArea)->D2_COD
		TMPXXC->DESCRI	 	:= RTRIM(Posicione("SB1",1,XFILIAL("SB1")+(cArea)->D2_COD,"B1_DESC"))
		
		MsUnLock()
		dbselectarea(cArea)
		dbskip()
	EndDo
	
	DBSELECTAREA(cArea)
	(cArea)->(DbCloseArea())
	
	
ElseIf TMPXXB->TPDOC == "MI"
	
	c_Documen 	:= rtrim(TMPXXB->DOCUMEN)
	c_String := " SELECT * FROM "+ RETSQLNAME("SD3")+" SD3 WHERE SD3.D_E_L_E_T_=' ' AND D3_TM = '"+TMPXXB->TM+"'  AND D3_DOC = '"+c_Documen+"'  AND LEFT(D3_YPEDVEN,6)='"+SC5->C5_NUM+"' AND  D3_FILIAL='"+xfilial("SD3")+"' "
	DbUseArea(.T.,"TOPCONN",TcGenQry(,,c_String ),cArea,.T.,.T.)
	dbselectarea(cArea)
	While !Eof()
		dbselectarea("TMPXXC")
		RecLock("TMPXXC",.T.)
		TMPXXC->PEDIDO  	:= LEFT((cArea)->(D3_YPEDVEN),6)
		TMPXXC->ITEMPV  	:= RIGHT((cArea)->(D3_YPEDVEN),2)
		TMPXXC->QUANTIDADE 	:= (cArea)->D3_QUANT
		TMPXXC->COD		 	:= (cArea)->D3_COD
		TMPXXC->DESCRI	 	:= RTRIM(Posicione("SB1",1,XFILIAL("SB1")+(cArea)->D3_COD,"B1_DESC"))
		
		MsUnLock()
		dbselectarea(cArea)
		dbskip()
	EndDo
	
	DBSELECTAREA(cArea)
	(cArea)->(DbCloseArea())
	
EndIf

DBSELECTAREA("TMPXXC")
DBGOTOP()

oBwITMov:oBrowse:Refresh()

Return




/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �DALSMI2   �Autor  �Microsiga           � Data �  08/02/16   ���
�������������������������������������������������������������������������͹��
���Desc.     �                                                            ���
���          �                                                            ���
�������������������������������������������������������������������������͹��
���Uso       � AP                                                        ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
User Function GeraMov()
Local oTempXXT := FWTemporaryTable():New( "TMPXXT" )

If !(upper(alltrim(upper(PSWRET()[1,13]))) $ "RESPONSAVEL/GESTOR/DIRETOR")
	If !(ALLTRIM(UPPER(cUsername)) $ "/ADMINISTRADOR/ADMIN/")
		APMSGINFO("Usuario Sem Permiss�o para Efetuar esta Opera��o")
		Return()
	EndIf
Endif

Private cSayDescQt := Space(1)
Private nQtdItens := 0

SetPrvt("oFont1","oDlgGerSai","oSay1","oSay2","oBrw1","oBtnGerar","oBtnCanc")
Private lInverte := .F.
Private cMarca := GetMark()

oFont1     := TFont():New( "Arial Rounded MT Bold",0,-16,,.F.,0,,400,.F.,.F.,,,,,, )
oDlgGerSai := MSDialog():New( 139,288,655,1154,"Gerar Movimento de Sa�da",,,.F.,,,,,,.T.,,oFont1,.T. )
oSay1      := TSay():New( 008,016,{||"Pedido"},oDlgGerSai,,oFont1,.F.,.F.,.F.,.T.,CLR_BLACK,CLR_WHITE,032,008)
oSay2      := TSay():New( 006,052,{||SC5->C5_NUM},oDlgGerSai,,oFont1,.F.,.F.,.F.,.T.,CLR_HRED,CLR_WHITE,032,008)
oSayDescQt := TSay():New( 228,012,{||"Quantidade  Selecionada:"},oDlgGerSai,,oFont1,.F.,.F.,.F.,.T.,CLR_BLACK,CLR_WHITE,140,008)
oBtnGerar  := TButton():New( 024,012,"Gerar",oDlgGerSai,{||U_GeraMI()},037,012,,,,.T.,,"",,,,.F. )
oBtnCanc   := TButton():New( 024,068,"Cancelar",oDlgGerSai,{ || IIF(APMSGYESNO("Deseja realmente sair?"),Close(oDlgGerSai),) },037,012,,,,.T.,,"",,,,.F. )
oGet1      := TGet():New( 224,132,{|u| If(PCount()>0,nQtdItens:=u,nQtdItens)},oDlgGerSai,120,018,'',,CLR_RED,CLR_WHITE,oFont1,,,.T.,"",,,.F.,.F.,,.F.,.F.,"","nQtdItens",,)
oGet1:Disable()

a_CpsGerar := {}
CriaGerar(oTempXXT)
DbSelectArea("TMPXXT")
oBrwGerar := MsSelect():New( "TMPXXT","OK","",a_CpsGerar,@lInverte,@cMarca,{040,012,220,416},,, oDlgGerSai )

oBrwGerar:bMark := {||FMarcaP(cMarca,lInverte,oGet1)}
oBrwGerar:oBrowse:lhasMark    := .T.
oBrwGerar:oBrowse:lCanAllmark := .T.
oBrwGerar:oBrowse:bAllMark := {||FInverteT(cMarca,lInverte,oGet1)}

CarrGerar()

oDlgGerSai:Activate(,,,.T.)
DbSelectArea("TMPXXT")
TMPXXT->(dbCloseArea())

Return


/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �DALSMI2   �Autor  �Microsiga           � Data �  08/02/16   ���
�������������������������������������������������������������������������͹��
���Desc.     �                                                            ���
���          �                                                            ���
�������������������������������������������������������������������������͹��
���Uso       � AP                                                        ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Static Function CriaGerar(oTempXXT)

aStruct := {}

AADD(aStruct,{"OK","C",2,0})
AADD(aStruct,{"ITEM","C",TamSX3('C9_ITEM')[1],0})
AADD(aStruct,{"PRODUTO","C",TamSX3('C9_PRODUTO')[1],0})
AADD(aStruct,{"DESCRI","C",TamSX3('B1_DESC')[1],0})
AADD(aStruct,{"QUANTIDADE","N",12,0})
AADD(aStruct,{"CODAMB","C",3,0})
AADD(aStruct,{"DESCAMB","C",40,0})
AADD(aStruct,{"NREC","N",12,0})
AADD(aStruct,{"ALMOX","C",TamSX3('C9_LOCAL')[1],0})
AADD(aStruct,{"PEDIDO","C",TamSX3('C9_PEDIDO')[1],0})


AADD(a_CpsGerar,{"OK",,"Ok","@!" })
AADD(a_CpsGerar,{"PRODUTO",,"Codigo","@!" })
AADD(a_CpsGerar,{"QUANTIDADE",,"Quant","@E 999,999" })
AADD(a_CpsGerar,{"DESCAMB",,"Ambiente","@!"})
AADD(a_CpsGerar,{"ITEM",,"Item", })
AADD(a_CpsGerar,{"DESCRI",,"Descricao","@!" })
AADD(a_CpsGerar,{"CODAMB",,"Cod Ambiente", })

// SetFields (sem o 'o') � este ja estava correto
oTempXXT:SetFields( aStruct )
oTempXXT:AddIndex("TMPXXT", {"CODAMB","DESCAMB"} )

//------------------
//Cria��o da tabela
//------------------
oTempXXT:Create()
dbselectarea("TMPXXT")

// CORRIGIDO: era oTempXXT:FWTemporaryTable():Zap() � instanciava nova classe ao inves de chamar no objeto
oTempXXT:Zap()
Return



/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �DALSMI2   �Autor  �Microsiga           � Data �  08/02/16   ���
�������������������������������������������������������������������������͹��
���Desc.     �                                                            ���
���          �                                                            ���
�������������������������������������������������������������������������͹��
���Uso       � AP                                                        ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Static Function CarrGerar()
Local cArea := "TTTXX"

DbSelectArea("TMPXXT")
ZAP

c_String := " SELECT * FROM "+ RETSQLNAME("SC9")+" SC9 "
c_String += " JOIN "+RETSQLNAME("SC5")+" SC5 ON C5_NUM = C9_PEDIDO AND SC5.D_E_L_E_T_=' ' AND SC5.C5_FILIAL = '"+XFILIAL("SC5")+"' "
c_String += " JOIN "+RETSQLNAME("SC6")+" SC6 ON C6_NUM = C9_PEDIDO AND SC6.C6_ITEM = SC9.C9_ITEM  AND SC6.D_E_L_E_T_=' ' AND SC6.C6_FILIAL = '"+XFILIAL("SC6")+"' "
c_String += " WHERE SC9.D_E_L_E_T_=' ' AND C9_PEDIDO='"+SC5->C5_NUM+"' AND C9_NFISCAL = ' ' AND C9_BLEST =' ' AND C9_BLCRED=' ' AND  C9_FILIAL='"+xfilial("SC9")+"' "
c_String += " ORDER BY C5_FILIAL,C5_NUM,C6_CODAMB,C6_DESCAMB"
DbUseArea(.T.,"TOPCONN",TcGenQry(,,c_String ),cArea,.T.,.T.)
dbselectarea(cArea)
While !Eof()
	dbselectarea("TMPXXT")
	RecLock("TMPXXT",.T.)
	TMPXXT->QUANTIDADE	:= (cArea)->C9_QTDLIB
	TMPXXT->PRODUTO 	:= (cArea)->C9_PRODUTO
	TMPXXT->DESCAMB		:= (cArea)->C6_DESCAMB
	TMPXXT->CODAMB 		:= (cArea)->C6_CODAMB
	TMPXXT->ITEM 		:= (cArea)->C9_ITEM
	TMPXXT->DESCRI 		:= Posicione("SB1",1,XFILIAL("SB1")+(cArea)->C9_PRODUTO,"B1_DESC")
	TMPXXT->ALMOX		:= (cArea)->C9_LOCAL
	TMPXXT->PEDIDO		:= (cArea)->C9_PEDIDO
	TMPXXT->NREC 		:= (cArea)->R_E_C_N_O_
	MsUnLock()
	dbselectarea(cArea)
	dbskip()
EndDo

DBSELECTAREA(cArea)
(cArea)->(DbCloseArea())

DBSELECTAREA("TMPXXT")
DBGOTOP()

oBrwGerar:oBrowse:Refresh()

Return


/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �DALSMI2   �Autor  �Microsiga           � Data �  08/02/16   ���
�������������������������������������������������������������������������͹��
���Desc.     �                                                            ���
���          �                                                            ���
�������������������������������������������������������������������������͹��
���Uso       � AP                                                        ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
User Function GeraMI(cTipo)
Local _aCab1 := {}
Local _aItem := {}
Local _atotitem:={}
Local cCodigoTM:="501"
Local cCodProd:="00001 "
Local cUnid:="CX "
Local cStatEnt := "005"
Local lSucesso := .F.
Private lMsHelpAuto := .t.
Private lMsErroAuto := .f.
Private _acod:={"1","MP1"}

If !Empty(SC5->C5_YBLQFIN) 
   Alert("Pedido com bloqueio financeiro - Data: "+DTOC(SC5->C5_YBLQFIN)+" Usuario: "+SC5->C5_YUSUBLQ)
   Return()
Endif   

If !ApMsgYesNo("Confirma Gera��o de MI para os itens marcados?")
	Return()
EndIf


DbSelectArea("TMPXXT")
DbGoTop()
BEGIN TRANSACTION
While !TMPXXT->(eof())
	
	DbSelectArea("TMPXXT")
	If MARKED("OK")
		_nRec := TMPXXT->NREC
		DBSELECTAREA("SC9")
		SC9->(DBSETORDER(1))
		DBGOTO(_nRec)
		a460Estorna(,,,.T.)
		
		
		_aItem	:= { 	{"D3_COD" ,TMPXXT->PRODUTO ,NIL},;
		{"D3_QUANT" ,TMPXXT->QUANTIDADE ,NIL},;
		{"D3_YPEDVEN" ,SC5->C5_NUM+TMPXXT->ITEM ,NIL},;
		{"D3_LOCAL" ,TMPXXT->ALMOX ,NIL}}
		
		aadd(_atotitem,_aitem)
		
	EndIf
	TMPXXT->(DBSKIP())
EndDo

If Len(_atotitem) > 0
	
	_aCab1 	:= 	{ 	{"D3_TM" ,cCodigoTM , NIL},;
	{"D3_EMISSAO" ,ddatabase, NIL}}
	
	MSExecAuto({|x,y,z| MATA241(x,y,z)},_aCab1,_atotitem,3)
	If lMsErroAuto
		Mostraerro()
		DisarmTransaction()
		break
	Else
		DbSelectArea("SC6")
		SC6->(DBSETORDER(1))
		DbSelectArea("TMPXXT")
		DbGotop()
		While !TMPXXT->(eof())
			DbSelectArea("TMPXXT")
			If MARKED("OK")
				
				If SC6->(DBSEEK(XFILIAL("SC6")+TMPXXT->PEDIDO+TMPXXT->ITEM))
					DbSelectARea("SC6")
					RecLock("SC6",.F.)
					SC6->C6_QTENTMI	+= TMPXXT->QUANTIDADE
					MsUnLock()
				EndIf
				
			EndIf
			TMPXXT->(DBSKIP())
		EndDo

		u_AJSTATPED(SC5->C5_NUM)
		lSucesso := .T.
	EndIf

EndIf


END TRANSACTION

// ZAP (dentro de CarrGerar) nao pode ser executado dentro de transacao;
// os refreshes de tela ocorrem apenas apos o commit.
If lSucesso
	CarrItPED(oBrPED:aCols[oBrPED:nAt,1])
	CarrGerar()
	CarrMvPED(oTempXXB)
	CarrItMov(oTempXXB,oTempXXC)
	ApMsgInfo("Movimenta��o Gerada com Sucesso!!")
EndIf

Return


/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �DALSMI2   �Autor  �Microsiga           � Data �  08/03/16   ���
�������������������������������������������������������������������������͹��
���Desc.     �                                                            ���
���          �                                                            ���
�������������������������������������������������������������������������͹��
���Uso       � AP                                                        ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Static Function ESTMI(l_GerNF)
Local _aCab1 := {}
Local cCodigoTM := "501"
Local l_Ret :=.T.
Default l_GerNF := .F.
Private lMsHelpAuto := .t.
Private lMsErroAuto := .f.
Private lGeraNCC := .T.

If !(upper(alltrim(upper(PSWRET()[1,13]))) $ "GESTOR/DIRETOR/RESPONSAVEL/ADMIN")
	If !(cUsername $ "/ADMINISTRADOR/ADMIN/")
		APMSGINFO("Usuario Sem Permiss�o para Efetuar esta Opera��o")
		Return()
	EndIf
Endif

If TMPXXB->TPDOC <> "MI"
	ApMsgStop("Este lan�amento n�o � MI!")
	Return(.F.)
EndIf

If TMPXXB->ESTORNO == "S"
	ApMsgStop("Este lan�amento j� se encontra estornado!")
	Return(.F.)
EndIf

If !l_GerNF
   lGeraNCC := ApMsgYesNo("Gerar cr�dito para o cliente (NCC) ?")

   If !ApMsgYesNo("Deseja Realmente ESTORNAR este lan�amento ?")
      Return(.F.)
   EndIf
	
EndIf
Begin Transaction

DbSelectArea("SD3")
SD3->(DBSETORDER(2))
If SD3->(DBSeek(Xfilial("SD3")+RTRIM(TMPXXB->DOCUMEN)))
	
	_aCab1 	:= 	{ 	{"D3_TM" ,cCodigoTM , NIL},;
	{"D3_DOC" ,TMPXXB->DOCUMEN , NIL},;
	{"D3_EMISSAO" ,ddatabase, NIL}}
	
	MSExecAuto({|x,y,z| MATA241(x,y,z)},_aCab1,,6)
	If lMsErroAuto
		Mostraerro()
		DisarmTransaction()
		Return(.F.)
	Else
		DbSelectArea("SC6")
		SC6->(DbSetOrder(1))
		DbSelectArea("TMPXXC")
		DbGotop()
		While !TMPXXC->(EOF())
			If SC6->(DBSEEK(XFILIAL("SC6")+SC5->C5_NUM + TMPXXC->ITEMPV+TMPXXC->COD))
				DbSelectArea("SC6")
				RecLock("SC6",.F.)
				SC6->C6_QTENTMI	-= TMPXXC->QUANTIDADE
				MsUnLock()
			EndIf
			
			TMPXXC->(DBSKIP())
		EndDo
		If !l_GerNF
			CarrItPED(oBrPED:aCols[oBrPED:nAt,1])
			CarrMvPED(oTempXXB)
			CarrItMov(oTempXXB,oTempXXC)
			ApMsgInfo("Estorno realizado com Sucesso!")
		EndIf
	EndIf
Else
	ApMsgStop("Documento n�o localizado no SD3!")
	Return(.F.)
EndIf

End Transaction
Return(.T.)




/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �DALSMI2   �Autor  �Microsiga           � Data �  08/03/16   ���
�������������������������������������������������������������������������͹��
���Desc.     �                                                            ���
���          �                                                            ���
�������������������������������������������������������������������������͹��
���Uso       � AP                                                        ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Static Function GerNF()

Local aPvlNfs   := {}
Local nI        := 0
Local cSerie    := "20"
Local c_Texto	:= ""
Local l_Ok		:= .T.

If !(alltrim(upper(PSWRET()[1,13])) $ "GESTOR/DIRETOR/RESPONSAVEL/ADMIN")
	If !(cUsername $ "/ADMINISTRADOR/ADMIN/")
		APMSGINFO("Usuario Sem Permiss�o para Efetuar esta Opera��o")
		Return()
	EndIf
Endif


If TMPXXB->TPDOC <> "MI"
	ApMsgStop("Este lan�amento n�o � MI!")
	Return(.F.)
EndIf

If TMPXXB->ESTORNO == "S"
	ApMsgStop("Este lan�amento j� se encontra estornado!")
	Return(.F.)
EndIf

If !ApMsgYesNo("Deseja realmente gerar a nota ?")
	Return()
EndIf

If Sx5NumNota(@cSerie,SuperGetMV("MV_TPNRNFS"))
	
	Begin Transaction
	
	If !EstMI(.T.)
		DisarmTransaction()
		Return()
	EndIf
	TMPXXC->(DBGOTOP())
	
	SC5->(DBSETORDER(1))
	If SC5->(DBSEEK(XFILIAL("SC5")+TMPXXC->PEDIDO))

	   If !Empty(SC5->C5_YBLQFIN) 
          Alert("Pedido com bloqueio financeiro - Data: "+DTOC(SC5->C5_YBLQFIN)+" Usuario: "+SC5->C5_YUSUBLQ)
          l_Ok := .F.
       Endif   
		
		While !TMPXXC->(EOF()) .and. l_Ok
			SC6->(DBSETORDER(1))
			If SC6->(DBSEEK(XFILIAL("SC6")+TMPXXC->PEDIDO+TMPXXC->ITEMPV))
				
				RecLock("SC6")
				Begin Transaction
				MaLibDoFat(SC6->(RecNo()),TMPXXC->QUANTIDADE,.T.,.T.,.F.,.T.,.F.,.F.)
				End Transaction
				SC6->(MsUnLock())
				
				
				SC9->(DbSetOrder(1))
				If SC6->C6_NUM+SC6->C6_ITEM <> TMPXXC->PEDIDO+TMPXXC->ITEMPV
					c_Texto := "PEDIDO: "+SC5->C5_NUM+" [SC9 N�O LOCALIZADO] [PRODUTO: "+SC6->C6_PRODUTO+"] [ITEM : "+SC6->C6_ITEM+"]"+chr(13)+chr(10)
					l_OK := .F.
				Else
					
					If SC9->C9_BLEST <> ' ' .OR. SC9->C9_BLCRED <> ' '
						c_Texto := "ITEM BLOQUEADO -> "+SC9->C9_PRODUTO+" "+SC9->C9_ITEM+chr(13)+chr(10)
						l_OK := .F.
					EndIf
					
					If SC9->C9_QTDLIB <> TMPXXC->QUANTIDADE
						c_Texto := "ITEM FOI RESERVADO PARCIALMENTE -> "+SC9->C9_PRODUTO+" "+SC9->C9_ITEM+chr(13)+chr(10)
						l_OK := .F.
					EndIf
					
					
				EndIf
				
				If l_OK
					aAdd( aPvlNfs , {     SC5->C5_NUM,;
					SC6->C6_ITEM   	  ,;
					SC9->C9_SEQUEN     ,;
					SC9->C9_QTDLIB     ,;
					SC9->C9_PRCVEN     ,;
					SC9->C9_PRODUTO     ,;
					.F.                         ,;
					SC9->(RecNo())          ,;
					SC5->(RecNo())          ,;
					SC6->(RecNo())          ,;
					SE4->(RecNo())          ,;
					SB1->(RecNo())          ,;
					SB2->(RecNo())          ,;
					SF4->(RecNo())          })
				EndIf
			Else
				c_Texto := "ITEM N�O LOCALIZADO NO SC6 -> "+TMPXXC->PEDIDO+"  "+TMPXXC->ITEMPV+chr(13)+chr(10)
				l_OK := .F.
			EndIf
			
			TMPXXC->(DBSKIP())
		EndDo
		
		If l_Ok	
	       MaLiberOk({ SC5->C5_NUM },.T.)
		Endif
		
	Else
		c_Texto := "PEDIDO N�O LOCALIZADO NO SC5 -> "+TMPXXC->PEDIDO+chr(13)+chr(10)
		l_Ok := .F.
	EndIf
	
	If !Empty(aPvlNfs) .AND. l_OK
		
		MaPvlNfs(aPvlNfs,cSerie,.F.,.F.,.F.,.F.,.F.,0,0,.F.,.F.)
		
		If  RTRIM(cSerie)  == RTRIM(SF2->F2_SERIE) .AND. SF2->F2_FILIAL == XFILIAL("SF2") .AND. SD2->D2_PEDIDO == SC5->C5_NUM
			CarrItPED(oBrPED:aCols[oBrPED:nAt,1])
			CarrMvPED(oTempXXB)
			CarrItMov(oTempXXB,oTempXXC)
			apMsgInfo("Nota "+SF2->F2_DOC+" Gerada com SUCESSO!!!")
		ELse
			DisarmTransaction()
			ApMsgInfo("Problema na Gera��o da Nota. Nota n�o foi gerada!"+chr(13)+chr(10)+c_Texto)
		EndIf
		
	Else
		
		DisarmTransaction()
		ApMsgInfo("Nota n�o Gerada! "+c_Texto)
		
	EndIf
	
	
	
	End Transaction
	
EndIf

Return



Static Function FMarcaP(cMarca,lInverte,oGet1)


IF TMPXXT->OK == cMarca
	nQtdItens	+= TMPXXT->QUANTIDADE
Else
	nQtdItens -= TMPXXT->QUANTIDADE
Endif

oGet1:Refresh()
Return




Static Function FInverteT(cMarca,lInverte,oGet1)

Local lIsMarcado:= .F.     
nQtdItens := 0
dbSelectArea("TMPXXT")
dbGotop()
If TMPXXT->OK == cMarca
	lIsMarcado:= .T.
EndIf
While ! Eof()
	RecLock("TMPXXT")
	If lIsMarcado
		TMPXXT->OK := " "
	Else     
		TMPXXT->OK := cMarca   
		nQtdItens += TMPXXT->QUANTIDADE
	EndIf
	dbSkip()
End

dbSelectArea("TMPXXT")
dbGotop()
oGet1:Refresh()
oBrwGerar:oBrowse:Refresh(.t.)
Return Nil
