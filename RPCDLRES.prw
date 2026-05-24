#INCLUDE "TOTVS.CH"
#INCLUDE "TBICONN.CH"
#include "TOPCONN.ch"
User Function RELDL()//aqui eh onde escreve SIGAADV U_RFINR09
	MsApp():New("SIGAFAT")
	oApp:CreateEnv()
	PtSetTheme("OCEAN")
	oApp:cStartProg := "U_RPCDLRES"
	__lInternet     := .T.
	oApp:Activate()
Return Nil


/*/
	*=========================================================================*
	| Programa  | Autor ERIVALDO   | D a t a  |  Alterado Por   |    Data     |
	|-------------------------------------------------------------------------|
	| IMPMINUTA |                  | 10.11.25 |                 | 10.11.2025  |
	|-------------------------------------------------------------------------|
	| Descrição : Emissão o Pedido de Venda digitado ou alterado              |
	*=========================================================================*
/*/

User Function RPCDLRES()
	Local cHtml := ""          as character
	Local cFile                as character
	Local cPasta               as character
	Local aSize := MsAdvSize() as array
	Local nPort := 0           as numeric
	Local oModal               as object
	Local oWebEngine           as object
	Local oSay                 as object
	Local oEdit                as object
	Local oDlg                 as object
	Local cNumPed := Space(6)
	Local cArqHTML           as character
	Local cArqPDF            as character
	Local cperg := "RDALPCOMP"
	local c_Query := " "
	Local cUrl := " "
	Private cAliasTrab := "TRBPC"
	Private cNomeEmp :=" "
	Private oWebChannel                     as object
	Private nColFinal   := (aSize[5] / 1.4) as numeric
	Private nLinFinal   := aSize[6]         as numeric
	Private nLinInicial := aSize[7]         as numeric
	Private cpdfComp    := " "
	oWebChannel := TWebChannel():New()
	// Inicia a estrutura HTML
	cPasta   := GetTempPath()

	nImpressao := 1

	//Declarando as fontes
	If ! Pergunte(cPerg,.T.)
		Return
	EndIf
	cNumPed := mv_par01
	cArqHTML := "RDLCOMP_N_" + AllTrim(cNumPed) + ".html"
	cArqPDF  := "RDLCOMP_N_" + AllTrim(cNumPed) + ".pdf"
	cArqPDF  := STRTRAN( cArqPDF , "/","_")
	cArqPDF := STRTRAN( cArqPDF, ":","_")


	FWMsgRun(, {|oSay| cHTML := u_FPrtPcHTML(cArqHTML) }, "Processando", "Analisando o ambiente...")

	// Salva o HTML em um arquivo temporário
	//cPasta := '\data\'
	cFile  := cArqHTML //"vrgestao"+DtoS(dDataBase)+Alltrim(Time())+".html"
	cFile  := STRTRAN( cFile , "/","_")
	cFile  := STRTRAN( cFile , ":","_")

	ccompleto := cPasta+cFile
	cpdfComp  := cpasta+cArqPDF
// === SALVA E ABRE ===
	If !Empty(cHtml) //MemoWrite(cCompleto, cHtml)
		oEdit := tSimpleEditor():New(0, 0, oDlg, 260, 184)
		oEdit:SetMaxTextLength(-1 )
		oEdit:TextFormat(1)
		oEdit:Load(cHtml)
		oEdit:SaveToPDF(cpdfComp)// "C:\RELATO\PEDIDOS_DL.PDF"
		//	ShellExecute("open", cCompleto, "", "", 1)
		//ShellExecute("open", cPasta+cArqPDF, "", "", 1)
		//MsgInfo("Pedido gerado com sucesso! e Sera aberto no Navegador, utilize CTRL+P para impressão" + CRLF + ccompleto)
	Else
		MsgStop("Erro ao salvar o arquivo!")
	EndIf
	//oModal := MSDialog():New(nLinInicial,0,nLinFinal,nColFinal, "Página Local",,,,,,,,,.T./*lPixel*/)
	////Prepara o conector
	//nPort := oWebChannel:connect()
	////Cria o componente que irá carregar o arquivo local
	//oWebEngine := TWebEngine():New(oModal, 0, 0, 100, 100,/*cUrl*/, nPort)
	//cUrl := "file:///" + StrTran(cCompleto, "\", "/")
	//oWebEngine:SetHtml( MemoRead(ccompleto) )
	//oWebEngine:Align := CONTROL_ALIGN_ALLCLIENT
	//TButton():New( 010, 300, "Print", oModal, {|| oWebEngine:PrintPDF() },50,010,,,.F.,.T.,.F.,,.F.,,,.F. ) //original
	//oModal:Activate()

Return

User Function FPrtPcHTM(cArqHTML)
	Local cHTML := ""
	Local c_Comprador := Space(15)
	Local nTotalPc := 0
	Local nTotQtdpc := 0
	Local cCond := ""
	Local cDescCond:= ""
	Local dDtEmis  := ctod("  /  /  ")
	Local cPCant := " "
	Local nLine := 0
	Local nPage := 0
	//Local cFile := "pedido.html"

	// Posiciona nos arquivos
	c_Query := " SELECT C7_USER,E4_CODIGO,E4_COND, E4_DESCRI,A2_COD,A2_END,A2_LOJA,A2_NOME,A2_NREDUZ,A2_BAIRRO,A2_MUN,A2_EST,A2_CEP,A2_TEL,A2_FAX,A2_CGC,A2_INSCR,C7_PRODUTO,B1_COD,B1_DESC,C7_UM, C7_PRECO, C7_FORNECE, C7_LOJA,C7_EMISSAO, C7_NUM, SUM(C7_QUANT) AS C7_QUANT , SUM(C7_TOTAL) AS C7_TOTAL, SUM(C7_QUJE) AS C7_QUJA, "+CRLF
	c_Query += " ISNULL(CONVERT(VARCHAR(8000),CONVERT(VARBINARY(8000),C7_OBS)),'') AS C7_OBS, "+CRLF
	c_Query += " ISNULL(CONVERT(VARCHAR(8000),CONVERT(VARBINARY(8000),ZOB.ZOB_OBS)),'') AS ZOB_OBS, "+CRLF
	c_Query += " ISNULL(CONVERT(VARCHAR(8000),CONVERT(VARBINARY(8000),C7_OBSITEM)),'') AS C7_OBSITEM "+CRLF
	c_Query += " FROM "+RETSQLNAME("SC7")+" SC7 "+CRLF
	c_Query += " JOIN "+RETSQLNAME("SA2")+" SA2 ON SA2.A2_COD = SC7.C7_FORNECE AND SA2.A2_LOJA = SC7.C7_LOJA AND SA2.D_E_L_E_T_=' ' AND SA2.A2_FILIAL='"+XFILIAL("SA2")+"' "+CRLF
	c_Query += " JOIN "+RETSQLNAME("SB1")+" SB1 ON SB1.B1_COD = SC7.C7_PRODUTO AND SB1.D_E_L_E_T_=' ' AND SB1.B1_FILIAL='"+XFILIAL("SB1")+"' "+CRLF
	c_Query += " LEFT JOIN "+RETSQLNAME("SE4")+" SE4 ON SE4.E4_CODIGO = SC7.C7_COND AND SE4.D_E_L_E_T_=' 'AND SE4.E4_FILIAL='"+XFILIAL("SE4")+"' "+CRLF
	c_Query += " LEFT JOIN "+RETSQLNAME("ZOB")+" ZOB ON ZOB.ZOB_TABELA='SC7' AND ZOB.ZOB_CHAVE = SC7.C7_NUM AND ZOB.D_E_L_E_T_= ' '  AND ZOB.ZOB_FILIAL='"+XFILIAL("ZOB")+"' "+CRLF
	c_Query += " WHERE SC7.C7_FILIAL=' '  "+CRLF
	c_Query += " AND SC7.D_E_L_E_T_=' ' "+CRLF
	c_Query += " AND SC7.C7_NUM BETWEEN '"+MV_PAR02+"' AND '"+MV_PAR03+"' "+CRLF
	c_Query += " GROUP BY ISNULL(CONVERT(VARCHAR(8000),CONVERT(VARBINARY(8000),C7_OBSITEM)),''),ISNULL(CONVERT(VARCHAR(8000),CONVERT(VARBINARY(8000),ZOB.ZOB_OBS)),''),C7_USER,E4_CODIGO,E4_COND, E4_DESCRI,ISNULL(CONVERT(VARCHAR(8000),CONVERT(VARBINARY(8000),C7_OBS)),''),A2_COD,A2_END,A2_LOJA,A2_NOME,A2_NREDUZ,A2_BAIRRO,A2_MUN,A2_EST,A2_CEP,A2_TEL,A2_FAX,A2_CGC,A2_INSCR,C7_PRODUTO,B1_COD,B1_DESC,C7_UM, C7_PRECO, C7_OBS, C7_FORNECE, C7_LOJA,C7_EMISSAO, C7_NUM "+CRLF
	c_Query += " ORDER BY C7_NUM,B1_DESC"+CRLF
//	MEMOWRITE("SQL.TXT",c_Query)

	MPSysOpenQuery( c_Query, cAliasTrab )
	TCSetField(cAliasTrab, "C7_EMISSAO","D",8,0)
	DBSelectArea(cAliasTrab)
	DBGoTop()
	// Início da estrutura HTML única
// Início da estrutura HTML única
	cHTML += '<!DOCTYPE html>' + CRLF
	cHTML += '<html lang="pt-BR">' + CRLF
	cHTML += '<head>' + CRLF
	cHTML += '    <meta charset="UTF-8">' + CRLF
	cHTML += '    <meta name="viewport" content="width=device-width, initial-scale=1.0">' + CRLF
	cHTML += '    <title>Pedido de Compra</title>' + CRLF
	cHTML += '    <style>' + CRLF
	cHTML += '        body {' + CRLF
	cHTML += '            font-family: Arial, sans-serif;' + CRLF
	cHTML += '            margin: 0;' + CRLF
	cHTML += '            padding: 20px;' + CRLF
	cHTML += '            box-sizing: border-box;' + CRLF
	cHTML += '        }' + CRLF
	cHTML += '        @page {' + CRLF
	cHTML += '            size: A4;' + CRLF
	cHTML += '            margin: 2cm;' + CRLF
	cHTML += '        }' + CRLF
	cHTML += '        @media print {' + CRLF
	cHTML += '            body {' + CRLF
	cHTML += '                margin: 0;' + CRLF
	cHTML += '                padding: 0;' + CRLF
	cHTML += '            }' + CRLF
	cHTML += '            .top-header, .header-info, .footer {' + CRLF
	cHTML += '                page-break-inside: avoid;' + CRLF
	cHTML += '            }' + CRLF
	cHTML += '            table {' + CRLF
	cHTML += '                page-break-inside: auto;' + CRLF
	cHTML += '            }' + CRLF
	cHTML += '            tr {' + CRLF
	cHTML += '                page-break-inside: avoid;' + CRLF
	cHTML += '                page-break-after: auto;' + CRLF
	cHTML += '            }' + CRLF
	cHTML += '            .page {' + CRLF
	cHTML += '                page-break-after: always;' + CRLF
	cHTML += '            }' + CRLF
	cHTML += '        }' + CRLF
	cHTML += '        .top-header {' + CRLF
	cHTML += '            display: flex;' + CRLF
	cHTML += '            justify-content: space-between;' + CRLF
	cHTML += '            margin-bottom: 5px;' + CRLF
	cHTML += '            font-size: 12px;' + CRLF
	cHTML += '        }' + CRLF
	cHTML += '        .header-info {' + CRLF
	cHTML += '            display: flex;' + CRLF
	cHTML += '            justify-content: space-between;' + CRLF
	cHTML += '            gap: 10px;' + CRLF
	cHTML += '            margin-bottom: 10px;' + CRLF
	cHTML += '        }' + CRLF
	cHTML += '        .company-box, .supplier-box {' + CRLF
	cHTML += '            width: calc(50% - 5px);' + CRLF
	cHTML += '            padding: 5px;' + CRLF
	cHTML += '            border: 1px solid #ddd;' + CRLF
	cHTML += '            box-sizing: border-box;' + CRLF
	cHTML += '            font-size: 12px;' + CRLF
	cHTML += '        }' + CRLF
	cHTML += '        .company-box p, .supplier-box p {' + CRLF
	cHTML += '            margin: 2px 0;' + CRLF
	cHTML += '        }' + CRLF
	cHTML += '        table {' + CRLF
	cHTML += '            width: 100%;' + CRLF
	cHTML += '            border-collapse: collapse;' + CRLF
	cHTML += '            margin-bottom: 20px;' + CRLF
	cHTML += '            font-size: 12px;' + CRLF
	cHTML += '        }' + CRLF
	cHTML += '        th, td {' + CRLF
	cHTML += '            border: 1px solid #ddd;' + CRLF
	cHTML += '            padding: 4px;' + CRLF
	cHTML += '            text-align: left;' + CRLF
	cHTML += '        }' + CRLF
	cHTML += '        th {' + CRLF
	cHTML += '            background-color: #f2f2f2;' + CRLF
	cHTML += '        }' + CRLF
	cHTML += '        .total {' + CRLF
	cHTML += '            text-align: right;' + CRLF
	cHTML += '            font-weight: bold;' + CRLF
	cHTML += '        }' + CRLF
	cHTML += ' th:nth-child(4), td:nth-child(4),' + CRLF
	cHTML += ' th:nth-child(5), td:nth-child(5),' + CRLF
	cHTML += ' th:nth-child(6), td:nth-child(6) {' + CRLF
	cHTML += '     text-align: right;' + CRLF
	cHTML += '     }' + CRLF
	cHTML += '    .currency {' + CRLF
	cHTML += '     white-space: nowrap;' + CRLF
	cHTML += '     }' + CRLF
	cHTML += '        .footer {' + CRLF
	cHTML += '            margin-top: 10px;' + CRLF
	cHTML += '            font-size: 12px;' + CRLF
	cHTML += '        }' + CRLF
	cHTML += '        .footer-info {' + CRLF
	cHTML += '            display: flex;' + CRLF
	cHTML += '            justify-content: space-between;' + CRLF
	cHTML += '            gap: 10px;' + CRLF
	cHTML += '            margin-bottom: 10px;' + CRLF
	cHTML += '        }' + CRLF
	cHTML += '        .payment-box, .observations-box {' + CRLF
	cHTML += '            width: calc(50% - 5px);' + CRLF
	cHTML += '            padding: 5px;' + CRLF
	cHTML += '            border: 1px solid #ddd;' + CRLF
	cHTML += '            box-sizing: border-box;' + CRLF
	cHTML += '            font-size: 12px;' + CRLF
	cHTML += '        }' + CRLF
	cHTML += '        .payment-box p, .observations-box p {' + CRLF
	cHTML += '            margin: 2px 0;' + CRLF
	cHTML += '        }' + CRLF
	cHTML += '        .delivery-box {' + CRLF
	cHTML += '            width: 100%;' + CRLF
	cHTML += '            padding: 5px;' + CRLF
	cHTML += '            border: 1px solid #ddd;' + CRLF
	cHTML += '            box-sizing: border-box;' + CRLF
	cHTML += '            font-size: 12px;' + CRLF
	cHTML += '        }' + CRLF
	cHTML += '        .delivery-box p {' + CRLF
	cHTML += '            margin: 2px 0;' + CRLF
	cHTML += '        }' + CRLF
	cHTML += '        .logo {' + CRLF
	cHTML += '            text-align: center;' + CRLF
	cHTML += '            margin-bottom: 10px;' + CRLF
	cHTML += '        }' + CRLF
	cHTML += '        .logo img {' + CRLF
	cHTML += '            max-width: 200px;' + CRLF
	cHTML += '            height: auto;' + CRLF
	cHTML += '        }' + CRLF
	cHTML += '        .continued {' + CRLF
	cHTML += '            font-size: 14px;' + CRLF
	cHTML += '            font-weight: bold;' + CRLF
	cHTML += '            margin-bottom: 10px;' + CRLF
	cHTML += '        }' + CRLF
	cHTML += '    </style>' + CRLF
	cHTML += '</head>' + CRLF
	cHTML += '<body>' + CRLF

	nMaxLines := 18  // Defina o limite máximo de linhas por página aqui. Ajuste conforme necessário.

	While !Eof()
		cPCant := (cAliasTrab)->C7_NUM
		nTotalPc := 0
		nTotQtdpc := 0
		nPage := 1
		nLine := 0
		PswOrder(1)
		PswSeek((cAliasTrab)->C7_USER, .T. )
		c_Comprador := PswRet()[1,2]
		cCond := (cAliasTrab)->E4_COND
		cDescCond := (cAliasTrab)->E4_DESCRI
		cObspc := (cAliasTrab)->C7_OBS+CRLF+(cAliasTrab)->ZOB_OBS+CRLF
		dDtEmis := (cAliasTrab)->C7_EMISSAO

		// Início da primeira página
		cHTML := u_fCabPchtml(cHTML,cAliasTrab) // Cabeçalho do pedido  

		While !Eof()
			If Eof() .Or. (cAliasTrab)->C7_NUM != cPCant
				Exit
			EndIf

			If nLine >= nMaxLines
				// Fechar tabela sem total (pois ainda há itens)
				cHTML += '        </tbody>' + CRLF
				cHTML += '    </table>' + CRLF
				// Não adicionar footer ainda
				cHTML += '    </div>' + CRLF  // Fechar div da página atual

				// Iniciar nova página de continuação
				nPage++
				nLine := 0
				cHTML += '    <div class="page">' + CRLF
				cHTML += '    <div class="logo">' + CRLF
				//cHTML += '        <img src="http://www.daluziluminacao.com.br/images/logo.png" alt="DaLuz Iluminação Logo">' + CRLF
				cHTML += '    </div>' + CRLF
				cHTML += '    <p class="continued">Continuação do Pedido de Compra: ' + cPCant + ' - Página ' + cValToChar(nPage) + '</p>' + CRLF
				// Repetir cabeçalho simplificado ou completo, aqui simplificado para economia de espaço
				cHTML += '    <div class="header-info">' + CRLF
				cHTML += '        <div class="company-box">' + CRLF
				cHTML += '            <p><strong>Empresa: </strong> DALUZ ILUMINACAO</p>' + CRLF
				cHTML += '            <p><strong>Endereço: </strong> RUA PADRE BERNARDINO PESSOA, 238</p>' + CRLF
				cHTML += '            <p><strong>Cep: </strong> 51020-010 <strong>Cidade: </strong>  RECIFE <strong>Estado:  </strong> PE</p>' + CRLF
				cHTML += '            <p><strong>Telefone:</strong> 81 34659433 <strong>Fax:  </strong> 81 34659433</p>' + CRLF
				cHTML += '            <p><strong>CNPJ:</strong>  05.484.465/0001-89 <strong>IE:  </strong> 029898323</p>' + CRLF
				cHTML += '        </div>' + CRLF
				cHTML += '        <div class="supplier-box">' + CRLF
				cHTML += '            <p><strong>PEDIDO DE COMPRA:</strong>'+(cAliasTrab)->C7_NUM+'</p>' + CRLF
				cHTML += '            <p><strong>Fornecedor:</strong>'+(cAliasTrab)->A2_COD+"-"+(cAliasTrab)->A2_LOJA+'</p>' + CRLF
				cHTML += '            <p><strong>Razão Social:</strong>'+(cAliasTrab)->A2_NOME+'</p>' + CRLF
				cHTML += '            <p><strong>Nome Fantasia:</strong>'+(cAliasTrab)->A2_NREDUZ+'</p>' + CRLF
				cHTML += '            <p><strong>Endereço:</strong>'+(cAliasTrab)->A2_END+'<strong>Bairro:</strong>'+(cAliasTrab)->A2_BAIRRO+'</p>' + CRLF
				cHTML += '            <p><strong>Municipio:</strong>'+(cAliasTrab)->A2_MUN+'<strong>Estado:</strong>'+(cAliasTrab)->A2_EST+'<strong>CEP:</strong>'+TRANSFORM((cAliasTrab)->A2_CEP,"@R 99999-999")+'</p>' + CRLF
				cHTML += '            <p><strong>Telefone:</strong>'+(cAliasTrab)->A2_TEL+'<strong>Fax:</strong> </p>' + CRLF
				cHTML += '            <p><strong>CNPJ:</strong>'+TRANSFORM((cAliasTrab)->A2_CGC,"@R 99.999.999/9999-99")+'<strong>IE:</strong>'+(cAliasTrab)->A2_INSCR+'</p>' + CRLF
				cHTML += '        </div>' + CRLF
				cHTML += '    </div>' + CRLF
				cHTML += '' + CRLF
				cHTML += '    <table>' + CRLF
				cHTML += '        <thead>' + CRLF
				cHTML += '            <tr>' + CRLF
				cHTML += '                <th>Código</th>' + CRLF
				cHTML += '                <th>Descrição</th>' + CRLF
				cHTML += '                <th>UM</th>' + CRLF
				cHTML += '                <th>Quantidade</th>' + CRLF
				cHTML += '                <th>Preço</th>' + CRLF
				cHTML += '                <th>Total</th>' + CRLF
				cHTML += '            </tr>' + CRLF
				cHTML += '        </thead>' + CRLF
				cHTML += '        <tbody>' + CRLF
			EndIf

			// Adicionar linha do item
			cHTML += '            <tr>' + CRLF
			cHTML += '                <td>'+(cAliasTrab)->C7_PRODUTO+'</td>' + CRLF
			cHTML += '                <td>'+(cAliasTrab)->B1_DESC+'</td>' + CRLF
			cHTML += '                <td>'+(cAliasTrab)->C7_UM+'</td>' + CRLF
			cHTML += '                <td>'+TRANSFORM((cAliasTrab)->C7_QUANT,"@E 999,999.99")+'</td>' + CRLF
			cHTML += '                <td>'+TRANSFORM((cAliasTrab)->C7_PRECO,"@E 999,999.99")+'</td>' + CRLF
			cHTML += '                <td>'+TRANSFORM((cAliasTrab)->C7_TOTAL,"@E 999,999.99")+'</td>' + CRLF
			cHTML += '            </tr>' + CRLF

			nTotalPc += (cAliasTrab)->C7_TOTAL
			nTotQtdpc += (cAliasTrab)->C7_QUANT
			cObspc += (cAliasTrab)->C7_OBSITEM

			nLine++
			(cAliasTrab)->(DBSkip())
		End

		// Fechar tabela com total na última página
		cHTML += '        </tbody>' + CRLF
		cHTML += '        <tfoot>' + CRLF
		cHTML += '            <tr class="total">' + CRLF
		cHTML += '                <td colspan="3">Total:</td>' + CRLF
		cHTML += '                <td>'+TRANSFORM(nTotQtdpc,"@E 99,999,999.99")+'</td>' + CRLF
		cHTML += '                <td></td>' + CRLF
		cHTML += '                <td>'+TRANSFORM(nTotalPc,"@E 99,999,999.99")+'</td>' + CRLF
		cHTML += '            </tr>' + CRLF
		cHTML += '        </tfoot>' + CRLF
		cHTML += '    </table>' + CRLF

		// Adicionar footer apenas na última página
		cHTML += '    <div class="footer">' + CRLF
		cHTML += '        <div class="footer-info">' + CRLF
		cHTML += '            <div class="payment-box">' + CRLF
		cHTML += '                <p><strong>Condição de Pagamento:</strong>'+cCond+"-"+cDescCond+'</p>' + CRLF
		cHTML += '                <p><strong>Data de Emissão:</strong>'+DTOC(dDtEmis)+'</p>' + CRLF
		cHTML += '                <p><strong>Comprador:</strong>'+c_Comprador+'</p>' + CRLF
		cHTML += '                <p></p>' + CRLF
		cHTML += '            </div>' + CRLF
		cHTML += '            <div class="observations-box">' + CRLF
		cHTML += '                <p><strong>Observação Geral:</strong>'+cObspc+'</p>' + CRLF
		cHTML += '            </div>' + CRLF
		cHTML += '        </div>' + CRLF
		cHTML += '        <div class="delivery-box">' + CRLF
		cHTML += '            <p><strong>Endereço de Entrega:</strong> RUA PADRE BERNARDINO PESSOA, 238 - BOA VIAGEM - RECIFE / PE - CEP: 51020-210</p>' + CRLF
		cHTML += '            <p><strong>Endereço de Cobrança:</strong> RUA PADRE BERNARDINO PESSOA, 238 - BOA VIAGEM - RECIFE / PE - CEP: 51020-210</p>' + CRLF
		cHTML += '        </div>' + CRLF
		cHTML += '    </div>' + CRLF
		cHTML += '    </div>' + CRLF  // Fechar a última div da página
	End

	cHTML += '</body>' + CRLF
	cHTML += '</html>'
	// Gravar o arquivo HTML
//	MemoWrite(cFile, cHTML)

	// Imprimir o arquivo HTML usando o visualizador padrão (geralmente o navegador)
	//ShellExecute("print", cFile, "", GetTempPath(), 1)

Return(cHtml)

User Function fCabPchtml(cHTML,cAliasTrab)
	cHTML += '    <div class="page">' + CRLF
	cHTML += '    <div class="logo">' + CRLF
	cHTML += '        <img src="http://www.daluziluminacao.com.br/images/logo.png" alt="DaLuz Iluminação Logo">' + CRLF
	cHTML += '    </div>' + CRLF
	cHTML += '    <div class="header-info">' + CRLF
	cHTML += '        <div class="company-box">' + CRLF
	cHTML += '            <p><strong>Empresa: </strong> DALUZ ILUMINACAO</p>' + CRLF
	cHTML += '            <p><strong>Endereço: </strong> RUA PADRE BERNARDINO PESSOA, 238</p>' + CRLF
	cHTML += '            <p><strong>Cep: </strong> 51020-010 <strong>Cidade: </strong>  RECIFE <strong>Estado:  </strong> PE</p>' + CRLF
	cHTML += '            <p><strong>Telefone:</strong> 81 34659433 <strong>Fax:  </strong> 81 34659433</p>' + CRLF
	cHTML += '            <p><strong>CNPJ:</strong>  05.484.465/0001-89 <strong>IE:  </strong> 029898323</p>' + CRLF
	cHTML += '        </div>' + CRLF
	cHTML += '        <div class="supplier-box">' + CRLF
	cHTML += '            <p><strong>PEDIDO DE COMPRA:</strong>'+(cAliasTrab)->C7_NUM+'</p>' + CRLF
	cHTML += '            <p><strong>Fornecedor:</strong>'+(cAliasTrab)->A2_COD+"-"+(cAliasTrab)->A2_LOJA+'</p>' + CRLF
	cHTML += '            <p><strong>Razão Social:</strong>'+(cAliasTrab)->A2_NOME+'</p>' + CRLF
	cHTML += '            <p><strong>Nome Fantasia:</strong>'+(cAliasTrab)->A2_NREDUZ+'</p>' + CRLF
	cHTML += '            <p><strong>Endereço:</strong>'+(cAliasTrab)->A2_END+'<strong>Bairro:</strong>'+(cAliasTrab)->A2_BAIRRO+'</p>' + CRLF
	cHTML += '            <p><strong>Municipio:</strong>'+(cAliasTrab)->A2_MUN+'<strong>Estado:</strong>'+(cAliasTrab)->A2_EST+'<strong>CEP:</strong>'+TRANSFORM((cAliasTrab)->A2_CEP,"@R 99999-999")+'</p>' + CRLF
	cHTML += '            <p><strong>Telefone:</strong>'+(cAliasTrab)->A2_TEL+'<strong>Fax:</strong> </p>' + CRLF
	cHTML += '            <p><strong>CNPJ:</strong>'+TRANSFORM((cAliasTrab)->A2_CGC,"@R 99.999.999/9999-99")+'<strong>IE:</strong>'+(cAliasTrab)->A2_INSCR+'</p>' + CRLF
	cHTML += '        </div>' + CRLF
	cHTML += '    </div>' + CRLF
	cHTML += '' + CRLF
	cHTML += '    <table>' + CRLF
	cHTML += '        <thead>' + CRLF
	cHTML += '            <tr>' + CRLF
	cHTML += '                <th>Código</th>' + CRLF
	cHTML += '                <th>Descrição</th>' + CRLF
	cHTML += '                <th>UM</th>' + CRLF
	cHTML += '                <th>Quantidade</th>' + CRLF
	cHTML += '                <th>Preço</th>' + CRLF
	cHTML += '                <th>Total</th>' + CRLF
	cHTML += '            </tr>' + CRLF
	cHTML += '        </thead>' + CRLF
	cHTML += '        <tbody>' + CRLF
Return(cHTML)

User Function ConvHtmlPdf(ccompleto,cArqPDF)
	Local cHtmlFile :=  ccompleto//"C:\temp\seu_arquivo_gerado.html"  // Caminho do HTML gerado
	Local cPdfFile  := cArqPDF //"C:\temp\saida.pdf"               // Caminho para o PDF resultante
	Local cWkPath   := "\Protheus\bin\wkhtmltopdf\wkhtmltopdf.exe"  // Caminho do executável
	Local cCmd      := '"' + cWkPath + '" "' + cHtmlFile + '" "' + cPdfFile + '"'
	Local nRet      := 0

	// Executa a conversão no servidor
	nRet := WaitRunSrv(cCmd, .T., "C:\temp")  // Ou use ShellExecute se for client-side (mas evite em HTML)

	If nRet == 0
		MsgInfo("PDF gerado com sucesso em " + cPdfFile)

		// Para imprimir o PDF: Abra no navegador ou envie para spool
		StartURL(cPdfFile)  // Abre no navegador padrão para o usuário imprimir (Ctrl+P)
		// Ou use: ShellExecute("print", cPdfFile, "", "", 1)  // Imprime diretamente, mas depende do client
	Else
		MsgAlert("Erro na conversão: " + Str(nRet))
	EndIf

Return
//#INCLUDE "PROTHEUS.CH"
//#INCLUDE "FWPRINTSETUP.CH"
//#INCLUDE "TOPCONN.CH"
//#INCLUDE "TBICONN.CH"
//
//User Function fPrintPc()
//	Local oPrinter
//	Local oFont7 := TFont():New("Arial", , 7, , .F., , , , , .F.)
//	Local oFont8 := TFont():New("Arial", , 8, , .F., , , , , .F.)
//	Local oFont10 := TFont():New("Arial", , 10, , .F., , , , , .F.)
//	Local oFont12N := TFont():New("Arial", , 12, , .T., , , , , .F.)
//	Local nLineHeight := 50
//	Local nCol1 := 50
//	Local nCol2 := 600
//	Local nCol3 := 700
//	Local nCol4 := 900
//	Local nCol5 := 1100
//	Local nCol6 := 1300
//	Local nCol7 := 1800
//	Local nY := 100
//	Local aItems := {;
//		{"ILU3018.01", "ACESSORIO PEN BASE SOBREPOR", "UN", "1", "43.00", "43.00"},;
//		{"ILU2011.01", "ARANDELA NUDE 1XG13 NAO INCLUSA", "UN", "3", "67.00", "201.00"},;
//		{"ILU2344.01", "EMB DILLO Q 2XILED 200 25°", "UN", "1", "390.00", "390.00"},;
//		{"ILU2347.01", "EMB GAP Q LEDS - LED NAO INCLUSO", "UN", "3", "389.00", "1.167.00"},;
//		{"ILU2348.01", "EMB GAP Q LEDS - LED NAO INCLUSO", "UN", "1", "641.00", "641.00"},;
//		{"ILU2408.01", "EMB HOLE-POP ORIENTAVEL 1XGU10", "UN", "2", "53.00", "106.00"},;
//		{"ILU2407.01", "EMB HOLE-POP ORIENTAVEL 1XGU10 MINDIC", "UN", "2", "56.00", "112.00"},;
//		{"ILU2812.01", "EMB TUNE 1XGU10 MINI DICROICA BIVOLT LED NAO INCLUSO", "UN", "3", "101.00", "303.00"},;
//		{"ILU2356.01", "PERFIL SLIT EMB TETO/PAREDE 2M CORTAR EM OBRA PARA 1.70M", "UN", "1", "211.00", "211.00"},;
//		{"ILU2356.01", "PERFIL SLIT EMB TETO/PAREDE 2M CORTAR EM OBRA PARA 1.80M", "UN", "1", "211.00", "211.00"},;
//		{"ILU2367.01", "PERFIL SLIT EMB TETO/PAREDE 1.50M", "UN", "1", "170.00", "170.00"},;
//		{"ILU2271.01", "PERFIL SLOT PADRAO L-30 1.00M LED NAO INCLUSO", "UN", "2", "106.00", "212.00"},;
//		{"ILU2616.01", "PLAFON CI 1XGU10 DIC", "UN", "3", "72.00", "216.00"},;
//		{"ILU3014.01", "SPOT PEN 1XGU10 MINIDICROICA", "UN", "1", "108.00", "108.00"},;
//		{"ILU3010.01", "SPOT PEN 1XILED200 25° BIVOLT + DRIVER", "UN", "3", "170.00", "510.00"}}
//
//	oPrinter := FWMSPrinter():New("PedidoCompra", IMP_PDF, .F., , .T.)
//	oPrinter:SetResolution(300)
//	oPrinter:SetPortrait()
//	oPrinter:SetPaperSize(9) // A4
//	oPrinter:SetMargin(60,60,60,60)
//	oPrinter:Setup()
//
//	oPrinter:StartPage()
//
//	// Logo (assumindo caminho local, ajuste conforme necessário)
//	// oPrinter:SayBitmap(nY - 50, 800, "c:\logo.png", 200, 100)
//
//	// Top Header
//	oPrinter:Say(nY, nCol1, "BAIRRO: TAMARINEIRA", oFont10)
//	oPrinter:Say(nY, nCol7 - 200, "CEP: 52110090", oFont10)
//	nY += nLineHeight
//
//	oPrinter:Say(nY, nCol1, "- nc", oFont10)
//	nY += nLineHeight * 2
//
//	// Company Box
//	oPrinter:Box(nY, nCol1, nY + 300, nCol1 + 800)
//	oPrinter:Say(nY + 20, nCol1 + 10, "Empresa: DALUZ ILUMINACAO", oFont8)
//	oPrinter:Say(nY + 40, nCol1 + 10, "Endereço: RUA PADRE BERNARDINO PESSOA, 238", oFont8)
//	oPrinter:Say(nY + 60, nCol1 + 10, "Cep: 51020-010 Cidade: RECIFE Estado: PE", oFont8)
//	oPrinter:Say(nY + 80, nCol1 + 10, "Telefone: 81 34659433 Fax: 81 34659433", oFont8)
//	oPrinter:Say(nY + 100, nCol1 + 10, "CNPJ: 05.484.465/0001-89 IE: 029898323", oFont8)
//
//	// Supplier Box
//	oPrinter:Box(nY, nCol1 + 850, nY + 300, nCol1 + 1650)
//	oPrinter:Say(nY + 20, nCol1 + 860, "PEDIDO DE COMPRA: 910947/1 OK", oFont8)
//	oPrinter:Say(nY + 40, nCol1 + 860, "Fornecedor: 00647 - Loja: 01/047/1", oFont8)
//	oPrinter:Say(nY + 60, nCol1 + 860, "Razão Social: MMC ILUMINACAO IRELLI LTDA", oFont8)
//	oPrinter:Say(nY + 80, nCol1 + 860, "Nome Fantasia: MMC", oFont8)
//	oPrinter:Say(nY + 100, nCol1 + 860, "Endereço: ROD. JOSE FRANCISCO DA SILVA, 1505 Bairro: OSWALDO BARB", oFont8)
//	oPrinter:Say(nY + 120, nCol1 + 860, "Municipio: NOVA LIMA Estado: MG CEP: 34000000", oFont8)
//	oPrinter:Say(nY + 140, nCol1 + 860, "Telefone: 35891400 Fax: ", oFont8)
//	oPrinter:Say(nY + 160, nCol1 + 860, "CNPJ: 20.247.730/0001-07 IE: 448244440048", oFont8)
//	nY += 320
//
//	// Table Header
//	oPrinter:Box(nY, nCol1, nY + nLineHeight, nCol7)
//	oPrinter:Say(nY + 10, nCol1 + 10, "Código", oFont10)
//	oPrinter:Say(nY + 10, nCol2, "Descrição", oFont10)
//	oPrinter:Say(nY + 10, nCol3 + 100, "UM", oFont10)
//	oPrinter:Say(nY + 10, nCol4 + 50, "Quantidade", oFont10)
//	oPrinter:Say(nY + 10, nCol5 + 100, "Preço", oFont10)
//	oPrinter:Say(nY + 10, nCol6 + 100, "Total", oFont10)
//	nY += nLineHeight
//
//	// Table Rows
//	For i := 1 To Len(aItems)
//		oPrinter:Box(nY, nCol1, nY + nLineHeight, nCol7)
//		oPrinter:Say(nY + 10, nCol1 + 10, aItems[i][1], oFont8)
//		oPrinter:Say(nY + 10, nCol2, aItems[i][2], oFont8)
//		oPrinter:Say(nY + 10, nCol3 + 100, aItems[i][3], oFont8)
//		oPrinter:Say(nY + 10, nCol4 + 50, aItems[i][4], oFont8)
//		oPrinter:Say(nY + 10, nCol5 + 100, aItems[i][5], oFont8)
//		oPrinter:Say(nY + 10, nCol6 + 100, aItems[i][6], oFont8)
//		nY += nLineHeight
//	Next
//
//	// Total
//	oPrinter:Box(nY, nCol1, nY + nLineHeight, nCol7)
//	oPrinter:Say(nY + 10, nCol1 + 10, "Total:", oFont10)
//	oPrinter:Say(nY + 10, nCol4 + 50, "39", oFont8)
//	oPrinter:Say(nY + 10, nCol6 + 100, "9.464.00", oFont8)
//	nY += nLineHeight * 2
//
//	// Footer Payment Box
//	oPrinter:Box(nY, nCol1, nY + 150, nCol1 + 800)
//	oPrinter:Say(nY + 20, nCol1 + 10, "Condição de Pagamento: 001 - DINHEIRO", oFont8)
//	oPrinter:Say(nY + 40, nCol1 + 10, "Data de Emissão: 13/06/25", oFont8)
//	oPrinter:Say(nY + 60, nCol1 + 10, "Comprovador: Adriana Mota", oFont8)
//	oPrinter:Say(nY + 80, nCol1 + 10, "AVISTA/DINHEIRO", oFont8)
//
//	// Observations Box
//	oPrinter:Box(nY, nCol1 + 850, nY + 150, nCol1 + 1650)
//	oPrinter:Say(nY + 20, nCol1 + 860, "Observação Geral:", oFont8)
//
//	nY += 160
//
//	// Delivery Box
//	oPrinter:Box(nY, nCol1, nY + 100, nCol7)
//	oPrinter:Say(nY + 20, nCol1 + 10, "Endereço de Entrega: RUA PADRE BERNARDINO PESSOA, 238 - BOA VIAGEM - RECIFE / PE - CEP: 51020-210", oFont8)
//	oPrinter:Say(nY + 40, nCol1 + 10, "Endereço de Cobrança: RUA PADRE BERNARDINO PESSOA, 238 - BOA VIAGEM - RECIFE / PE - CEP: 51020-210", oFont8)
//
//	oPrinter:EndPage()
//	oPrinter:Preview()
//
//Return
