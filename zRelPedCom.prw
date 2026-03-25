//Bibliotecas
#Include "Totvs.ch"

/*/{Protheus.doc} fReportDef
Relatório Pedido de Compras
@author carlos.abrantes
@since 17/03/2026
@version 1.0
@type function
/*/

User Function zRelPedCom()
	Local aArea     := FWGetArea()
	Local _cUserOk  := AllTrim(GetMV('SN_RELPCOM',.F.,'000000|000860|000993'))
	Local aPergs    := {}
	Local dDataDe   := FirstDate(Date())
	Local dDataAte  := LastDate(Date())
	Local cProdDe   := Space(TamSX3('C7_PRODUTO')[1])
	Local cProdAte  := StrTran(cProdDe, ' ', 'Z')
	Local _lEstOk   := .F.

	IF !(__cUserID $ _cUserOk)
			_cPopUp := ' <font color="#A4A4A4" face="Arial" size="7">Atenção</font> '
			_cPopUp += ' <br> '
			_cPopUp += ' <font color="#FF0000" face="Arial" size="2">Relatório de Pedidos de Compra</font> '
			_cPopUp += ' <br>'
    		_cPopUp += ' <font color="#000000" face="Arial" size="2"> ##SN_RELPCOM: Usuário sem permissão </font> '	
			Alert(_cPopUp,'Sunnyvale') 
    	RETURN(_lEstOk)
	ENDIF

	//Adicionando os parametros do ParamBox
	aAdd(aPergs, {1, "Data De", dDataDe,  "", ".T.", "", ".T.", 80,  .F.}) // MV_PAR01
	aAdd(aPergs, {1, "Data Até", dDataAte,  "", ".T.", "", ".T.", 80,  .T.}) // MV_PAR02
    aAdd(aPergs, {1, "Produto De",  cProdDe,  "", ".T.", "SC7", ".T.", 80,  .F.}) // MV_PAR03
    aAdd(aPergs, {1, "Produto Até", cProdAte, "", ".T.", "SC7", ".T.", 80,  .T.}) // MV_PAR04

    //Se a pergunta for confirma, cria as definicoes do relatorio
    If ParamBox(aPergs, "Informe os parâmetros", , , , , , , , , .F., .F.)
        //MV_PAR05 := Val(cValToChar(MV_PAR05))
		Processa({|| fGeraExcel()})
	EndIf
	
	FWRestArea(aArea)
Return


Static Function fGeraExcel()
	Local cQryDad  	 := ""
	Local oFWMsExcel
	Local oExcel
	Local cArquivo   := GetTempPath() + "exporta.xml"
	Local cWorkSheet := "Base"
	Local cTitulo    := "Pedidos de Compra"
	Local nAtual 	 := 0
	Local nTotal     := 0

	//Montando consulta de dados
	cQryDad += "SELECT " + CRLF
	cQryDad += "    SC7.C7_NUM      AS NUM_PED, " + CRLF
	cQryDad += "    SC7.C7_ITEM     AS ITEM, " + CRLF
	cQryDad += "    SC7.C7_EMISSAO  AS DATA_EMISSAO, " + CRLF
	cQryDad += "    SC7.C7_FORNECE  AS FORNECEDOR, " + CRLF
	cQryDad += "    SC7.C7_PRODUTO  AS PRODUTO, " + CRLF
	cQryDad += "    SB1.B1_FABRIC 	AS PART_NUMBER, " + CRLF
	cQryDad += "    SC7.C7_DESCRI 	AS DESCRICAO, " + CRLF
	cQryDad += "    SC7.C7_UM     	AS UM, " + CRLF
	cQryDad += "    SC7.C7_QUANT  	AS QUANTIDADE, " + CRLF
	cQryDad += "    SC7.C7_PRECO  	AS VALOR_UNITARIO, " + CRLF
	cQryDad += "    SC7.C7_IPI    	AS IPI, " + CRLF
	cQryDad += "    SC7.C7_TOTAL  	AS VALOR_TOTAL, " + CRLF
	cQryDad += "    SC7.C7_DATPRF 	AS DATA_ENTREGA, " + CRLF
	cQryDad += "    SC7.C7_CC     	AS CENTRO_CUSTO, " + CRLF
	cQryDad += "    SC7.C7_NUMSC  	AS SC, " + CRLF
	cQryDad += "    SC7.C7_MOEDA  	AS MOEDA, " + CRLF
	cQryDad += "    SC7.C7_TXMOEDA AS TAXA_MOEDA, " + CRLF
	cQryDad += "    CASE " + CRLF
	cQryDad += "        WHEN SC7.C7_MOEDA = '1' THEN 'REAL' " + CRLF
	cQryDad += "        WHEN SC7.C7_MOEDA = '2' THEN 'DOLAR' " + CRLF
	cQryDad += "        WHEN SC7.C7_MOEDA = '3' THEN 'EURO' " + CRLF
	cQryDad += "        WHEN SC7.C7_MOEDA = '4' THEN 'LIBRA' " + CRLF
	cQryDad += "        WHEN SC7.C7_MOEDA = '5' THEN 'UFIR' " + CRLF
	cQryDad += "        ELSE 'OUTRA' " + CRLF
	cQryDad += "    END AS DESC_MOEDA " + CRLF
	cQryDad += "FROM " + RetSQLName('SC7') + " SC7 " + CRLF
	cQryDad += "LEFT JOIN " + RetSQLName('SB1') + " SB1 " + CRLF
	cQryDad += "    ON SB1.B1_FILIAL = '" + FWxFilial('SB1') + "' " + CRLF
	cQryDad += "    AND SB1.B1_COD   = SC7.C7_PRODUTO " + CRLF
	cQryDad += "    AND SB1.D_E_L_E_T_ = ' ' " + CRLF
	cQryDad += "WHERE " + CRLF
	cQryDad += "    SC7.C7_FILIAL = '" + FWxFilial('SC7') + "' " + CRLF
	cQryDad += "    AND SC7.D_E_L_E_T_ = ' ' "		+ CRLF
	cQryDad += " 	AND SC7.C7_EMISSAO >= '" + dToS(MV_PAR01) + "' "		+ CRLF
    cQryDad += " 	AND SC7.C7_EMISSAO <= '" + dToS(MV_PAR02) + "' "		+ CRLF
    cQryDad += " 	AND SC7.C7_PRODUTO  >= '" + MV_PAR03 + "' " + CRLF
    cQryDad += " 	AND SC7.C7_PRODUTO  <= '" + MV_PAR04 + "' " + CRLF
    cQryDad += "ORDER BY SC7.C7_EMISSAO, SC7.C7_NUM " + CRLF

	//Executando consulta e setando o total da regua
	PlsQuery(cQryDad, "QRY_DAD")
	TCSetField("QRY_DAD", "DATA_EMISSAO", "D")
	TCSetField("QRY_DAD", "DATA_ENTREGA", "D")
	DbSelectArea("QRY_DAD")

	//Cria a planilha do excel
	oFWMsExcel := FWMSExcel():New()

	//Criando a aba da planilha
	oFWMsExcel:AddworkSheet(cWorkSheet)

	//Criando a Tabela e as colunas
	oFWMsExcel:AddTable(cWorkSheet, cTitulo)
	oFWMsExcel:AddColumn(cWorkSheet, cTitulo, "Nº", 1, 1, .F.)
	oFWMsExcel:AddColumn(cWorkSheet, cTitulo, "Item", 1, 1, .F.)
	oFWMsExcel:AddColumn(cWorkSheet, cTitulo, "Data Emissao", 2, 1, .F.)
	oFWMsExcel:AddColumn(cWorkSheet, cTitulo, "Fornecedor", 1, 1, .F.)
	oFWMsExcel:AddColumn(cWorkSheet, cTitulo, "Produto", 1, 1, .F.)
	oFWMsExcel:AddColumn(cWorkSheet, cTitulo, "Part. Number", 1, 1, .F.)
	oFWMsExcel:AddColumn(cWorkSheet, cTitulo, "Descricao", 1, 1, .F.)
	oFWMsExcel:AddColumn(cWorkSheet, cTitulo, "UM.", 1, 1,.F.)
	oFWMsExcel:AddColumn(cWorkSheet, cTitulo, "Quantidade", 2, 2,.F.)
	oFWMsExcel:AddColumn(cWorkSheet, cTitulo, "Valor Unitario", 2, 2, .F.)
	oFWMsExcel:AddColumn(cWorkSheet, cTitulo, "IPI", 2, 2, .F.)
    oFWMsExcel:AddColumn(cWorkSheet, cTitulo, "Valor Total", 2, 2, .F.)
    oFWMsExcel:AddColumn(cWorkSheet, cTitulo, "Entrega", 2, 1, .F.)
	oFWMsExcel:AddColumn(cWorkSheet, cTitulo, "C.Custo", 1, 1, .F.)
	oFWMsExcel:AddColumn(cWorkSheet, cTitulo, "S.C.", 1, 1, .F.)
	oFWMsExcel:AddColumn(cWorkSheet, cTitulo, "Moeda", 1, 1, .F.)
	oFWMsExcel:AddColumn(cWorkSheet, cTitulo, "Taxa Moeda", 1, 1, .F.)

	//Definindo o tamanho da regua
	Count To nTotal
	ProcRegua(nTotal)
	QRY_DAD->(DbGoTop())

	//Percorrendo os dados da query
	While !(QRY_DAD->(EoF()))

		//Incrementando a regua
		nAtual++
		IncProc("Gerando planilha " + cValToChar(nAtual) + " de " + cValToChar(nTotal) + "...")

		//Adicionando uma nova linha
		oFWMsExcel:AddRow(cWorkSheet, cTitulo, {;
			QRY_DAD->NUM_PED,;
			QRY_DAD->ITEM,;
			QRY_DAD->DATA_EMISSAO,;
			QRY_DAD->FORNECEDOR,;
			QRY_DAD->PRODUTO,;
			QRY_DAD->PART_NUMBER,;
			QRY_DAD->DESCRICAO,;
			QRY_DAD->UM,;
			QRY_DAD->QUANTIDADE,;
			QRY_DAD->VALOR_UNITARIO,;
			QRY_DAD->IPI,;
			QRY_DAD->VALOR_TOTAL,;
			QRY_DAD->DATA_ENTREGA,;
			QRY_DAD->CENTRO_CUSTO,;
			QRY_DAD->SC,;
			QRY_DAD->DESC_MOEDA,;
			QRY_DAD->TAXA_MOEDA;
			})

		QRY_DAD->(DbSkip())
	EndDo
	QRY_DAD->(DbCloseArea())

	//Ativando o arquivo e gerando o xml
	oFWMsExcel:Activate()
	oFWMsExcel:GetXMLFile(cArquivo)

	//Abrindo o excel e arquivo xml
	oExcel := MsExcel():New()
	oExcel:WorkBooks:Open(cArquivo)
	oExcel:SetVisible(.T.)
	oExcel:Destroy()

	MsgInfo("Término","Info")

Return

