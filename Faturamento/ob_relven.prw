#Include "TOTVS.ch" //por causa da RGB()
#Include "TopConn.ch"
#Include "RPTDef.ch"
#Include "FWPrintSetup.ch"

//Constantes
//Constantes
#Define CRLF Chr(13) + Chr(10) //Carriage Return Line Feed

//Alinhamentos
#Define PAD_LEFT    0
#Define PAD_RIGHT   1
#Define PAD_CENTER  2
#Define PAD_JUSTIFY 3 //Opção disponível somente a partir da versão 1.6.2 da TOTVS Printer

//Cor(es)
Static nColorGray := RGB(110, 110, 110)
Static nBackgroundColor := RGB(240, 255, 255)


User Function ob_relven()

	ValidPerg()


return

Static Function RunMessage(oSay)

	Local oExcel := FWMsExcel():New()
	local cArq := 'RTCVTEA.xml'
	local cPath   := AllTrim(GetTempPath())
	Local cArquivo := cPath + cArq
	//local cArquivo := '\1arqtrab\RTCVTEA.xml'
	msgalert('oi')

	If MV_PAR04 == '2'
		QryCOMPANT(oExcel,1)
		QryVTOTA(oExcel,1)
		QryCOMPATU(oExcel,1)

		msgalert("Caminho do relatorio: " + cPath )
		//Criando o XML
		oExcel:Activate()
		oExcel:GetXMLFile(cArquivo)

		//Abrindo o excel e abrindo o arquivo xml

		If ApOleClient("msexcel")
			oExcel := MsExcel():New()
			oExcel:WorkBooks:Open(cArquivo)
			oExcel:SetVisible(.T.)
			oExcel:Destroy()
		ElseIf ExistDir("C:\Program Files (x86)\LibreOffice 5")
			WaitRun('C:\Program Files (x86)\LibreOffice 5\program\scalc.exe "'+cArquivo+'"', 1)
		ElseIf ExistDir("C:\Program Files (x86)\LibreOffice")
			WaitRun('C:\Program Files (x86)\LibreOffice\program\scalc.exe "'+cArquivo+'"', 1)
		ElseIf ExistDir("C:\Program Files\LibreOffice\program")
			WaitRun('C:\Program Files\LibreOffice\program\scalc.exe "'+cArquivo+'"', 1)
		Else
			//Senão, abre o XML pelo programa padrão
			ShellExecute("open", cArq, "", cPath, 1)
		EndIf

	else
		If MV_PAR05 == '1'
			QryCOMPANT(oExcel,2)
		endif
		If MV_PAR05 == '2'
			QryVTOTA(oExcel,2)
		endif
		If MV_PAR05 == '3'
			QryCOMPATU(oExcel,2)
		endif
	endif


Return


Static Function QryCOMPANT(oExcel,nTipo)

	Local cQuery := ""
	Local cPlano := "Vendas Antigas"
	Local cTitulo := " "
	Local nB := 0
	Local _aColunas := {"Codigo","Loja","Nome","Endereco","Municipio","Telefone","Contato","Data Ultima Compra","Valor Total Comprado"}
	//Local _aColunas := {"Codigo","Loja","Nome","Data Ultima Compra","Valor Total Comprado"}

	Local nTotal         := 0
	Local nCurrent       := 0

	Local cFileName      := 'ob_relv'+RetCodUsr()+'_' + dToS(Date()) + '_' + StrTran(Time(), ':', '-') + '.pdf'
	Private oPrintReport
	Private oBrushLin  := TBrush():New(/*uParam1*/, nBackgroundColor)
	Private cTimeReport    := Time()
	Private nPageCurrent   := 1
	Private cCompanyLogo   := searchLogo()
	//Linhas e colunas
	Private nReportLine    := 0
	Private nFooterLimit   := 580
	Private nLeftMargin    := 010
	Private nRightLimit    := 815
	Private nMiddleCol   := (nRightLimit - nLeftMargin) / 2
	//Colunas dos relatorio
	Private nColData1    := 0
	Private nColData2    := 0
	Private nColData3    := 0
	Private nColData4    := 0
	Private nColData5    := 0
	Private nColData6    := 0
	Private nColData7    := 0
	Private nColData8    := 0

	if mv_par05 <> '3'
		nColData1    := nLeftMargin
		nColData2    := nLeftMargin + 50
		nColData3    := nLeftMargin + 90
		nColData4    := nLeftMargin + 500
		nColData5    := nLeftMargin + 600
	else
		nColData1    := nLeftMargin
		nColData2    := nLeftMargin + 50
		nColData3    := nLeftMargin + 90
		nColData4    := nLeftMargin + 500
		nColData5    := nLeftMargin + 580
		nColData6    := nLeftMargin + 620
		nColData7    := nLeftMargin + 680
		nColData8    := nLeftMargin + 750

	endif

	//Declarando as fontes
	Private cFont                := 'Arial'
	Private oFontDetails         := TFont():New(cFont, /*uPar2*/, -11, /*uPar4*/, .F., /*uPar6*/, /*uPar7*/, /*uPar8*/, /*uPar9*/, .F.)
	Private oFontDetailsBold     := TFont():New(cFont, /*uPar2*/, -13, /*uPar4*/, .T., /*uPar6*/, /*uPar7*/, /*uPar8*/, /*uPar9*/, .F.)
	Private oFontFooter          := TFont():New(cFont, /*uPar2*/, -8,  /*uPar4*/, .F., /*uPar6*/, /*uPar7*/, /*uPar8*/, /*uPar9*/, .F.)
	Private oFontHeaderColumns   := TFont():New(cFont, /*uPar2*/, -15,  /*uPar4*/, .F., /*uPar6*/, /*uPar7*/, /*uPar8*/, /*uPar9*/, .F.)
	Private oFontTitle           := TFont():New(cFont, /*uPar2*/, -15, /*uPar4*/, .T., /*uPar6*/, /*uPar7*/, /*uPar8*/, /*uPar9*/, .F.)
	Private _aDados := {}

	cQuery := "SELECT "+CRLF
	cQuery += "    A1.A1_COD AS Codigo,"+CRLF
	cQuery += "    A1.A1_LOJA AS Loja,"+CRLF
	cQuery += "    A1.A1_NOME AS Nome,"+CRLF
	cQuery += "    A1.A1_END AS Ender,"+CRLF
	cQuery += "    A1.A1_MUN AS Mun,"+CRLF
	cQuery += "    A1.A1_TEL AS Fone,"+CRLF
	cQuery += "    A1.A1_CONTATO AS Contato,"+CRLF

	cQuery += "    UltimaCompra.DataUltimaCompra,"+CRLF
	cQuery += "    CAST(ISNULL(TotalComprado.ValorTotal, 0) AS NUMERIC(18,2)) AS ValorTotalComprado"+CRLF
	cQuery += "FROM "+CRLF
	cQuery += "     "+RetSQLName('SA1')+" A1"+CRLF
	cQuery += "OUTER APPLY ("+CRLF
	cQuery += "    SELECT MAX(F2.F2_EMISSAO) AS DataUltimaCompra"+CRLF
	cQuery += "    FROM "+RetSQLName('SF2')+" F2"+CRLF
	cQuery += "    WHERE F2.F2_CLIENTE = A1.A1_COD"+CRLF
	cQuery += "      AND F2.F2_LOJA = A1.A1_LOJA"+CRLF
	cQuery += "      AND F2.D_E_L_E_T_ = ''"+CRLF
	cQuery += ") AS UltimaCompra"+CRLF
	cQuery += "OUTER APPLY ("+CRLF
	cQuery += "    SELECT SUM(ISNULL(E1.E1_VALOR, 0)) AS ValorTotal"+CRLF
	cQuery += "    FROM "+RetSQLName('SE1')+" E1"+CRLF
	cQuery += "    WHERE E1.E1_CLIENTE = A1.A1_COD"+CRLF
	cQuery += "    AND E1.E1_LOJA = A1.A1_LOJA"+CRLF

	IF ALLTRIM(MV_PAR02) <> ''
		cQuery += "		 AND E1.E1_TIPO IN "+FormatIn(alltrim(MV_PAR02),',')+CRLF
	ENDIF
	IF ALLTRIM(MV_PAR03) <> ''
		cQuery += "		 AND E1.E1_TIPO NOT IN "+FormatIn(alltrim(MV_PAR03),',')+CRLF
	ENDIF

	cQuery += "      AND E1.D_E_L_E_T_ = ''"+CRLF
	cQuery += ") AS TotalComprado"+CRLF
	cQuery += "WHERE "+CRLF
	cQuery += "    A1.D_E_L_E_T_ = ''"+CRLF
	cQuery += "    AND A1.A1_COD BETWEEN  '"+MV_PAR06+"' AND  '"+MV_PAR07+"' "+CRLF
	cQuery += "    AND UltimaCompra.DataUltimaCompra IS NOT NULL"+CRLF
	cQuery += "    AND A1.A1_COD NOT IN ("+CRLF
	cQuery += "        SELECT F2_CLIENTE "+CRLF
	cQuery += "        FROM "+RetSQLName('SF2')+CRLF
	cQuery += "        WHERE F2_EMISSAO >= '"+DtoS(MV_PAR01)+"'"+CRLF
	cQuery += "          AND D_E_L_E_T_ = ''"+CRLF
	cQuery += "    )"+CRLF
	cQuery += "	AND ValorTotal > 0"+CRLF
	cQuery += "ORDER BY "+CRLF
	cQuery += "    ValorTotalComprado desc"+CRLF

	TCQuery cQuery New Alias "COMPANT"

	if nTipo == 1
		cTitulo := "Clientes que nao compraram depois de "+ dtoc(mv_par01) + ", quando foi a ultima compra e o total ja comprado"

		//Alterando atributos
		oExcel:SetFontSize(10)
		oExcel:SetFont("Arial")
		//oExcel:SetBgGeneralColor("#0000FF")
		oExcel:SetTitleBold(.T.)
		//oExcel:SetTitleFrColor("#F8F8FF")
		//oExcel:SetLineFrColor("#1E90FF")
		//oExcel:Set2LineFrColor("#00BFFF")

		oExcel:AddworkSheet(cPlano)
		oExcel:AddTable(cPlano,cTitulo)
		//Adicionando as colunas
		For nB:=1 to len(_aColunas)
			if nB <= 7
				//colunas de texto
				oExcel:AddColumn(cPlano,cTitulo,alltrim(_aColunas[nB]),1,1)
			else
				oExcel:AddColumn(cPlano,cTitulo,alltrim(_aColunas[nB]),1,3)
			endif
		next

		DbSelectArea('COMPANT')
		DbGoTop()
		Do While !eof()
			_aDados := {}
			//Local _aColunas := {"Codigo","Loja","Nome","Endereco","Municipio","Telefone","Contato","Telefone","Data Ultima Compra","Valor Total Comprado"}
	
			AADD(_aDados, COMPANT->Codigo)
			AADD(_aDados, COMPANT->Loja)
			AADD(_aDados, COMPANT->Nome)
			AADD(_aDados, COMPANT->Ender)
			AADD(_aDados, COMPANT->Mun)
			AADD(_aDados, COMPANT->Fone)
			AADD(_aDados, COMPANT->Contato)
			AADD(_aDados, StoD(COMPANT->DataUltimaCompra))
			AADD(_aDados, COMPANT->ValorTotalComprado)

			//Pulando Registro
			oExcel:AddRow(cPlano,cTitulo,_aDados)
			DbSelectArea('COMPANT')
			DbSkip()

		EndDo

	else  //relatorio
		DbSelectArea('COMPANT')
		COMPANT->(DbGoTop())
		Count to nTotal
		ProcRegua(nTotal)
		COMPANT->(DbGoTop())

		//Somente se tiver dados
		If ! COMPANT->(EoF())
			//Criando o objeto de impressao
			oPrintReport := FWMSPrinter():New(;
				cFileName,;      // cFilePrinter
			IMP_PDF,;        // nDevice
			.F.,;            // lAdjustToLegacy
			,;               // cPathInServer
			.T.,;            // lDisabeSetup
			,;               // lTReport
			@oPrintReport,;  // oPrintSetup
			,;               // cPrinter
			,;               // lServer
			,;               // lParam10
			,;               // lRaw
			.T.;             // lViewPDF
			)
			oPrintReport:cPathPDF := GetTempPath()
			oPrintReport:SetResolution(72)
			oPrintReport:SetLandscape()
			oPrintReport:SetPaperSize(DMPAPER_A4)
			oPrintReport:SetMargin(0, 0, 0, 0)

			//Imprime os dados
			printHeader()
			While ! COMPANT->(EoF())
				nCurrent++
				IncProc('Imprimindo registro ' + cValToChar(nCurrent) + ' de ' + cValToChar(nTotal) + '...')

				//Se atingiu o limite, quebra de pagina
				validPageBreak()

				//Faz o zebrado ao fundo
				If nCurrent % 2 == 0
					oPrintReport:FillRect({nReportLine - 2, nLeftMargin, nReportLine + 12, nRightLimit}, oBrushLin)
				EndIf

				//Imprime a linha atual
				oPrintReport:SayAlign(nReportLine, nColData1, Alltrim(Transform(COMPANT->Codigo, '@!')), oFontDetails, 50, 10, /*nClrText*/, PAD_LEFT, /*nAlignVert*/)
				oPrintReport:SayAlign(nReportLine, nColData2, Alltrim(Transform(COMPANT->Loja, '@!')), oFontDetails, 30, 10, /*nClrText*/, PAD_LEFT, /*nAlignVert*/)
				oPrintReport:SayAlign(nReportLine, nColData3, Alltrim(Transform(COMPANT->Nome, '@!')), oFontDetails, 100, 10, /*nClrText*/, PAD_LEFT, /*nAlignVert*/)
				oPrintReport:SayAlign(nReportLine, nColData4, Alltrim(dtoc(stod(COMPANT->DataUltimaCompra))), oFontDetails, 80, 10, /*nClrText*/, PAD_LEFT, /*nAlignVert*/)
				oPrintReport:SayAlign(nReportLine, nColData5, Alltrim(Transform(COMPANT->ValorTotalComprado, "@E 999,999,999.99")), oFontDetailsBold, 80, 10, /*nClrText*/, PAD_LEFT, /*nAlignVert*/)

				nReportLine += 15
				oPrintReport:Line(nReportLine-3, nLeftMargin, nReportLine-3, nRightLimit, nColorGray)

				//Se atingiu o limite, quebra de pagina
				validPageBreak()

				COMPANT->(DbSkip())
			EndDo

			//Imprime o último rodapé
			printFooter()

			oPrintReport:Preview()


		endif
	endif
	COMPANT->(DbCloseArea())

Return


Static Function searchLogo()
	Local cLogo       := FisxLogo('1')
Return cLogo



Static Function printHeader()
	Local cText        := ''
	Local nHeaderLine  := 015

	//Iniciando Pagina
	oPrintReport:StartPage()

	//Imprime o logo
	If File(cCompanyLogo)
		oPrintReport:SayBitmap(005, nLeftMargin, cCompanyLogo, 030, 030)
	EndIf

	//Cabecalho
	if mv_par05 == '1'
		cText := "Clientes que nao compraram depois de "+ dtoc(mv_par01) + ", quando foi a ultima compra e o total ja comprado"
	endif
	if mv_par05 == '2'
		cText := "Relatorio sintetico de Vendas a partir de "+dtoc(mv_par01)
	endif
	if mv_par05 == '3'
		cText := "Relatorio analitico de Vendas a partir de "+dtoc(mv_par01)'
	endif

	oPrintReport:SayAlign(nHeaderLine, nMiddleCol-350, cText, oFontTitle, 600, 20, /*nClrText*/, PAD_CENTER, /*nAlignVert*/)

	//Linha Separatoria
	nHeaderLine += 020
	oPrintReport:Line(nHeaderLine,   nLeftMargin, nHeaderLine,   nRightLimit)

	//Atualizando a linha inicial do relatorio
	nReportLine := nHeaderLine + 5

	if mv_par05 == '1'
		oPrintReport:SayAlign(nReportLine, nColData1, 'Codigo do Cliente', oFontHeaderColumns, 50, 10, /*nClrText*/, PAD_LEFT, /*nAlignVert*/)
		oPrintReport:SayAlign(nReportLine, nColData2, 'Loja do Cliente', oFontHeaderColumns, 30, 10, /*nClrText*/, PAD_LEFT, /*nAlignVert*/)
		oPrintReport:SayAlign(nReportLine, nColData3, 'Nome do Cliente', oFontHeaderColumns, 100, 10, /*nClrText*/, PAD_LEFT, /*nAlignVert*/)
		oPrintReport:SayAlign(nReportLine, nColData4, 'Data Ultima Compra', oFontHeaderColumns, 40, 10, /*nClrText*/, PAD_LEFT, /*nAlignVert*/)
		oPrintReport:SayAlign(nReportLine, nColData5, 'Valor Comprado', oFontHeaderColumns, 80, 10, /*nClrText*/, PAD_LEFT, /*nAlignVert*/)
	endif

	if mv_par05 == '2'
		oPrintReport:SayAlign(nReportLine, nColData1, 'Codigo do Cliente', oFontHeaderColumns, 50, 10, /*nClrText*/, PAD_LEFT, /*nAlignVert*/)
		oPrintReport:SayAlign(nReportLine, nColData2, 'Loja do Cliente', oFontHeaderColumns, 30, 10, /*nClrText*/, PAD_LEFT, /*nAlignVert*/)
		oPrintReport:SayAlign(nReportLine, nColData3, 'Nome do Cliente', oFontHeaderColumns, 100, 10, /*nClrText*/, PAD_LEFT, /*nAlignVert*/)
		oPrintReport:SayAlign(nReportLine, nColData5, 'Valor Comprado', oFontHeaderColumns, 80, 10, /*nClrText*/, PAD_LEFT, /*nAlignVert*/)
	endif

	if mv_par05 == '3'
		oPrintReport:SayAlign(nReportLine, nColData1, 'Codigo do Cliente', oFontHeaderColumns, 50, 10, /*nClrText*/, PAD_LEFT, /*nAlignVert*/)
		oPrintReport:SayAlign(nReportLine, nColData2, 'Loja do Cliente', oFontHeaderColumns, 30, 10, /*nClrText*/, PAD_LEFT, /*nAlignVert*/)
		oPrintReport:SayAlign(nReportLine, nColData3, 'Nome do Cliente', oFontHeaderColumns, 100, 10, /*nClrText*/, PAD_LEFT, /*nAlignVert*/)
		oPrintReport:SayAlign(nReportLine, nColData4, 'Numero', oFontHeaderColumns, 80, 10, /*nClrText*/, PAD_LEFT, /*nAlignVert*/)
		oPrintReport:SayAlign(nReportLine, nColData5, 'Parcela', oFontHeaderColumns, 80, 10, /*nClrText*/, PAD_LEFT, /*nAlignVert*/)
		oPrintReport:SayAlign(nReportLine, nColData6, 'Tipo NF', oFontHeaderColumns, 80, 10, /*nClrText*/, PAD_LEFT, /*nAlignVert*/)
		oPrintReport:SayAlign(nReportLine, nColData7, 'Valor', oFontHeaderColumns, 80, 10, /*nClrText*/, PAD_LEFT, /*nAlignVert*/)
		oPrintReport:SayAlign(nReportLine, nColData8, 'Valor em aberto', oFontHeaderColumns, 80, 10, /*nClrText*/, PAD_LEFT, /*nAlignVert*/)
	endif


	nReportLine += 15
Return

Static Function printFooter()
	Local nFooterLine := nFooterLimit
	Local cText       := ''

	//Linha Separatoria
	oPrintReport:Line(nFooterLine,   nLeftMargin, nFooterLine,   nRightLimit)
	nFooterLine += 3

	//Dados da Esquerda
	cText := dToC(dDataBase) + '     ' + cTimeReport + '     ' + FunName() + ' (ob_relv)     ' + UsrRetName(RetCodUsr())
	oPrintReport:SayAlign(nFooterLine, nLeftMargin, cText, oFontFooter, 500, 10, /*nClrText*/, PAD_LEFT, /*nAlignVert*/)

	//Direita
	cText := 'Pagina '+cValToChar(nPageCurrent)
	oPrintReport:SayAlign(nFooterLine, nRightLimit-40, cText, oFontFooter, 040, 10, /*nClrText*/, PAD_RIGHT, /*nAlignVert*/)

	//Finalizando a pagina e somando mais um
	oPrintReport:EndPage()
	nPageCurrent++
Return

Static Function validPageBreak()
	If nReportLine >= nFooterLimit - 10
		printFooter()
		printHeader()
	EndIf
Return

Static Function QryVTOTA(oExcel,nTipo)

	Local cQuery := ""
	Local cPlano := "Sintetico Vendas Atuais"
	Local cTitulo := " "
	Local nB := 0
	Local _aColunas := {"Codigo","Loja","Nome","Valor Total Titulos"}


	Local nTotal         := 0
	Local nCurrent       := 0

	Local cFileName      := 'ob_relv'+RetCodUsr()+'_' + dToS(Date()) + '_' + StrTran(Time(), ':', '-') + '.pdf'
	Private oPrintReport
	Private oBrushLin  := TBrush():New(/*uParam1*/, nBackgroundColor)
	Private cTimeReport    := Time()
	Private nPageCurrent   := 1
	Private cCompanyLogo   := searchLogo()
	//Linhas e colunas
	Private nReportLine    := 0
	Private nFooterLimit   := 580
	Private nLeftMargin    := 010
	Private nRightLimit    := 815
	Private nMiddleCol   := (nRightLimit - nLeftMargin) / 2

	//Colunas dos relatorio
	Private nColData1    := 0
	Private nColData2    := 0
	Private nColData3    := 0
	Private nColData4    := 0
	Private nColData5    := 0
	Private nColData6    := 0
	Private nColData7    := 0
	Private nColData8    := 0

	if mv_par05 <> '3'
		nColData1    := nLeftMargin
		nColData2    := nLeftMargin + 50
		nColData3    := nLeftMargin + 90
		nColData4    := nLeftMargin + 500
		nColData5    := nLeftMargin + 600
	else
		nColData1    := nLeftMargin
		nColData2    := nLeftMargin + 50
		nColData3    := nLeftMargin + 90
		nColData4    := nLeftMargin + 500
		nColData5    := nLeftMargin + 580
		nColData6    := nLeftMargin + 620
		nColData7    := nLeftMargin + 680
		nColData8    := nLeftMargin + 750

	endif

	//Declarando as fontes
	Private cFont                := 'Arial'
	Private oFontDetails         := TFont():New(cFont, /*uPar2*/, -11, /*uPar4*/, .F., /*uPar6*/, /*uPar7*/, /*uPar8*/, /*uPar9*/, .F.)
	Private oFontDetailsBold     := TFont():New(cFont, /*uPar2*/, -13, /*uPar4*/, .T., /*uPar6*/, /*uPar7*/, /*uPar8*/, /*uPar9*/, .F.)
	Private oFontFooter          := TFont():New(cFont, /*uPar2*/, -8,  /*uPar4*/, .F., /*uPar6*/, /*uPar7*/, /*uPar8*/, /*uPar9*/, .F.)
	Private oFontHeaderColumns   := TFont():New(cFont, /*uPar2*/, -15,  /*uPar4*/, .F., /*uPar6*/, /*uPar7*/, /*uPar8*/, /*uPar9*/, .F.)
	Private oFontTitle           := TFont():New(cFont, /*uPar2*/, -15, /*uPar4*/, .T., /*uPar6*/, /*uPar7*/, /*uPar8*/, /*uPar9*/, .F.)

	Private _aDados := {}

	cQuery := "SELECT "+CRLF
	cQuery += "    E1.E1_CLIENTE AS Codigo,"+CRLF
	cQuery += "    E1.E1_LOJA AS Loja,"+CRLF
	cQuery += "    MAX(A1.A1_NOME) AS Nome,"+CRLF
	cQuery += "    CAST(SUM(ISNULL(E1.E1_VALOR, 0)) AS NUMERIC(18,2)) AS VTotalTitulos"+CRLF
	cQuery += "FROM "+CRLF
	cQuery += "    "+RetSQLName('SE1')+" E1"+CRLF
	cQuery += "    INNER JOIN "+RetSQLName('SA1')+" A1 "+CRLF
	cQuery += "        ON A1.A1_COD = E1.E1_CLIENTE "+CRLF
	cQuery += "       AND A1.A1_LOJA = E1.E1_LOJA"+CRLF
	cQuery += "       AND A1.D_E_L_E_T_ = ''"+CRLF
	cQuery += "WHERE "+CRLF
	cQuery += "    E1.D_E_L_E_T_ = ''"+CRLF
	cQuery += "    AND E1.E1_EMISSAO >= '"+DtoS(MV_PAR01)+"'"+CRLF
	IF ALLTRIM(MV_PAR02) <> ''
		cQuery += "	AND E1.E1_TIPO IN "+FormatIn(alltrim(MV_PAR02),',')+CRLF
	endif
	cQuery += "       AND A1.A1_COD BETWEEN  '"+MV_PAR06+"' AND  '"+MV_PAR07+"' "+CRLF
	IF ALLTRIM(MV_PAR03) <> ''
		cQuery += "	AND E1.E1_TIPO NOT IN "+FormatIn(alltrim(MV_PAR03),',')+CRLF
	endif
	cQuery += "GROUP BY "+CRLF
	cQuery += "    E1.E1_CLIENTE, E1.E1_LOJA, E1.E1_TIPO"+CRLF
	cQuery += "ORDER BY "+CRLF
	cQuery += "    VTotalTitulos desc;"+CRLF

	TCQuery cQuery New Alias "VTOTA"
	if nTipo == 1
		cTitulo := "Relatorio sintetico vendas a partir de "+dtoc(mv_par01)

		//Alterando atributos
		oExcel:SetFontSize(10)
		oExcel:SetFont("Arial")
		//oExcel:SetBgGeneralColor("#0000FF")
		oExcel:SetTitleBold(.T.)
		//oExcel:SetTitleFrColor("#F8F8FF")
		//oExcel:SetLineFrColor("#1E90FF")
		//oExcel:Set2LineFrColor("#00BFFF")

		oExcel:AddworkSheet(cPlano)
		oExcel:AddTable(cPlano,cTitulo)
		//Adicionando as colunas
		For nB:=1 to len(_aColunas)
			if nB <= 3
				//colunas de texto
				oExcel:AddColumn(cPlano,cTitulo,alltrim(_aColunas[nB]),1,1)
			else
				oExcel:AddColumn(cPlano,cTitulo,alltrim(_aColunas[nB]),1,3)
			endif
		next

		DbSelectArea('VTOTA')
		DbGoTop()
		Do While !eof()
			_aDados := {}
			AADD(_aDados, VTOTA->Codigo)
			AADD(_aDados, VTOTA->Loja)
			AADD(_aDados, alltrim(VTOTA->Nome))
			AADD(_aDados, VTOTA->VTotalTitulos)

			//Pulando Registro
			oExcel:AddRow(cPlano,cTitulo,_aDados)
			DbSelectArea('VTOTA')
			DbSkip()
		EndDo
	else  //relatorio

		DbSelectArea('VTOTA')
		VTOTA->(DbGoTop())
		Count to nTotal
		ProcRegua(nTotal)
		VTOTA->(DbGoTop())

		//Somente se tiver dados
		If ! VTOTA->(EoF())
			//Criando o objeto de impressao
			oPrintReport := FWMSPrinter():New(;
				cFileName,;      // cFilePrinter
			IMP_PDF,;        // nDevice
			.F.,;            // lAdjustToLegacy
			,;               // cPathInServer
			.T.,;            // lDisabeSetup
			,;               // lTReport
			@oPrintReport,;  // oPrintSetup
			,;               // cPrinter
			,;               // lServer
			,;               // lParam10
			,;               // lRaw
			.T.;             // lViewPDF
			)
			oPrintReport:cPathPDF := GetTempPath()
			oPrintReport:SetResolution(72)
			oPrintReport:SetLandscape()
			oPrintReport:SetPaperSize(DMPAPER_A4)
			oPrintReport:SetMargin(0, 0, 0, 0)

			//Imprime os dados
			printHeader()
			While ! VTOTA->(EoF())
				nCurrent++
				IncProc('Imprimindo registro ' + cValToChar(nCurrent) + ' de ' + cValToChar(nTotal) + '...')

				//Se atingiu o limite, quebra de pagina
				validPageBreak()

				//Faz o zebrado ao fundo
				If nCurrent % 2 == 0
					oPrintReport:FillRect({nReportLine - 2, nLeftMargin, nReportLine + 12, nRightLimit}, oBrushLin)
				EndIf

				//Imprime a linha atual
				oPrintReport:SayAlign(nReportLine, nColData1, Alltrim(Transform(VTOTA->Codigo, '@!')), oFontDetails, 50, 10, /*nClrText*/, PAD_LEFT, /*nAlignVert*/)
				oPrintReport:SayAlign(nReportLine, nColData2, Alltrim(Transform(VTOTA->Loja, '@!')), oFontDetails, 30, 10, /*nClrText*/, PAD_LEFT, /*nAlignVert*/)
				oPrintReport:SayAlign(nReportLine, nColData3, Alltrim(Transform(VTOTA->Nome, '@!')), oFontDetails, 100, 10, /*nClrText*/, PAD_LEFT, /*nAlignVert*/)
				oPrintReport:SayAlign(nReportLine, nColData5, Alltrim(Transform(VTOTA->VTotalTitulos, "@E 999,999,999.99")), oFontDetailsBold, 80, 10, /*nClrText*/, PAD_LEFT, /*nAlignVert*/)

				nReportLine += 15
				oPrintReport:Line(nReportLine-3, nLeftMargin, nReportLine-3, nRightLimit, nColorGray)

				//Se atingiu o limite, quebra de pagina
				validPageBreak()

				VTOTA->(DbSkip())
			EndDo

			//Imprime o último rodapé
			printFooter()

			oPrintReport:Preview()


		endif
	endif
	VTOTA->(DbCloseArea())

Return


Static Function QryCOMPATU(oExcel,nTipo)

	Local cQuery := ""
	Local cPlano := "Vendas Atuais "
	Local cTitulo := " "
	Local nB := 0
	Local _aColunas := {"Codigo","Loja","Nome","Numero","Parcela","Tipo NF","Valor Total Titulos","Em aberto"}
	Local nTotal         := 0
	Local nCurrent       := 0

	Local cFileName      := 'ob_relv'+RetCodUsr()+'_' + dToS(Date()) + '_' + StrTran(Time(), ':', '-') + '.pdf'
	Private oPrintReport
	Private oBrushLin  := TBrush():New(/*uParam1*/, nBackgroundColor)
	Private cTimeReport    := Time()
	Private nPageCurrent   := 1
	Private cCompanyLogo   := searchLogo()
	//Linhas e colunas
	Private nReportLine    := 0
	Private nFooterLimit   := 580
	Private nLeftMargin    := 010
	Private nRightLimit    := 815
	Private nMiddleCol   := (nRightLimit - nLeftMargin) / 2

    //Colunas dos relatorio
	Private nColData1    := 0
	Private nColData2    := 0
	Private nColData3    := 0
	Private nColData4    := 0
	Private nColData5    := 0
	Private nColData6    := 0
	Private nColData7    := 0
	Private nColData8    := 0

	if mv_par05 <> '3'
		nColData1    := nLeftMargin
		nColData2    := nLeftMargin + 50
		nColData3    := nLeftMargin + 90
		nColData4    := nLeftMargin + 500
		nColData5    := nLeftMargin + 600
	else
		nColData1    := nLeftMargin
		nColData2    := nLeftMargin + 50
		nColData3    := nLeftMargin + 90
		nColData4    := nLeftMargin + 500
		nColData5    := nLeftMargin + 580
		nColData6    := nLeftMargin + 620
		nColData7    := nLeftMargin + 680
		nColData8    := nLeftMargin + 750

	endif

	//Declarando as fontes
	Private cFont                := 'Arial'
	Private oFontDetails         := TFont():New(cFont, /*uPar2*/, -11, /*uPar4*/, .F., /*uPar6*/, /*uPar7*/, /*uPar8*/, /*uPar9*/, .F.)
	Private oFontDetailsBold     := TFont():New(cFont, /*uPar2*/, -13, /*uPar4*/, .T., /*uPar6*/, /*uPar7*/, /*uPar8*/, /*uPar9*/, .F.)
	Private oFontFooter          := TFont():New(cFont, /*uPar2*/, -8,  /*uPar4*/, .F., /*uPar6*/, /*uPar7*/, /*uPar8*/, /*uPar9*/, .F.)
	Private oFontHeaderColumns   := TFont():New(cFont, /*uPar2*/, -15,  /*uPar4*/, .F., /*uPar6*/, /*uPar7*/, /*uPar8*/, /*uPar9*/, .F.)
	Private oFontTitle           := TFont():New(cFont, /*uPar2*/, -15, /*uPar4*/, .T., /*uPar6*/, /*uPar7*/, /*uPar8*/, /*uPar9*/, .F.)


	Private _aDados := {}

	cQuery := "SELECT "+CRLF
	cQuery += "    E1.E1_CLIENTE AS Codigo,"+CRLF
	cQuery += "    E1.E1_LOJA AS Loja,"+CRLF
	cQuery += "    A1.A1_NOME AS Nome,"+CRLF
	cQuery += "    E1.E1_NUM as Numero,"+CRLF
	cQuery += "    E1.E1_PARCELA AS Parcela,"+CRLF
	cQuery += "    E1.E1_TIPO AS TipoNF,"+CRLF
	cQuery += "    CAST(SUM(ISNULL(E1.E1_VALOR, 0)) AS NUMERIC(18,2)) AS VTotalTitulos,"+CRLF
	cQuery += "    CAST(SUM(ISNULL(E1.E1_SALDO, 0)) AS NUMERIC(18,2)) AS VTotalEmAberto"+CRLF
	cQuery += "FROM "+CRLF
	cQuery += "    "+RetSQLName('SE1')+" E1"+CRLF
	cQuery += "    INNER JOIN "+RetSQLName('SA1')+" A1 "+CRLF
	cQuery += "        ON A1.A1_COD = E1.E1_CLIENTE "+CRLF
	cQuery += "       AND A1.A1_LOJA = E1.E1_LOJA"+CRLF
	cQuery += "       AND A1.D_E_L_E_T_ = ''"+CRLF
	cQuery += "WHERE "+CRLF
	cQuery += "    E1.D_E_L_E_T_ = ''"+CRLF
	cQuery += "    AND E1.E1_EMISSAO >= '"+DtoS(MV_PAR01)+"'"+CRLF
	IF ALLTRIM(MV_PAR02) <> ''
		cQuery += "	AND E1.E1_TIPO IN "+FormatIn(alltrim(MV_PAR02),',')+CRLF
	ENDIF
	IF ALLTRIM(MV_PAR03) <> ''
		cQuery += "	AND E1.E1_TIPO NOT IN "+FormatIn(alltrim(MV_PAR03),',')+CRLF
	ENDIF
	cQuery += "   AND A1.A1_COD BETWEEN  '"+MV_PAR06+"' AND  '"+MV_PAR07+"' "+CRLF
	cQuery += "GROUP BY "+CRLF
	cQuery += "    E1.E1_CLIENTE, E1.E1_LOJA,A1_NOME,E1_PARCELA, E1.E1_TIPO,E1_NUM"+CRLF
	cQuery += "ORDER BY "+CRLF
	cQuery += "    Codigo;"+CRLF

	TCQuery cQuery New Alias "COMPATU"

	if nTipo == 1
		cTitulo := "Relatorio analitico de Vendas a partir de "+dtoc(mv_par01)

		//Alterando atributos
		oExcel:SetFontSize(10)
		oExcel:SetFont("Arial")
		//oExcel:SetBgGeneralColor("#0000FF")
		oExcel:SetTitleBold(.T.)
		//oExcel:SetTitleFrColor("#F8F8FF")
		//oExcel:SetLineFrColor("#1E90FF")
		//oExcel:Set2LineFrColor("#00BFFF")

		oExcel:AddworkSheet(cPlano)
		oExcel:AddTable(cPlano,cTitulo)
		//Adicionando as colunas
		For nB:=1 to len(_aColunas)
			if nB < 6
				//colunas de texto
				oExcel:AddColumn(cPlano,cTitulo,alltrim(_aColunas[nB]),1,1)
			else
				oExcel:AddColumn(cPlano,cTitulo,alltrim(_aColunas[nB]),1,3)
			endif
		next

		DbSelectArea('COMPATU')
		DbGoTop()
		Do While !eof()
			_aDados := {}
			AADD(_aDados, COMPATU->Codigo)
			AADD(_aDados, COMPATU->Loja)
			AADD(_aDados, alltrim(COMPATU->Nome))
			AADD(_aDados, alltrim(COMPATU->Numero))
			AADD(_aDados, alltrim(COMPATU->Parcela))
			AADD(_aDados, alltrim(COMPATU->TipoNF))
			AADD(_aDados, COMPATU->VTotalTitulos)
			AADD(_aDados, COMPATU->VTotalEmAberto)

			//Pulando Registro
			oExcel:AddRow(cPlano,cTitulo,_aDados)
			DbSelectArea('COMPATU')
			DbSkip()
		EndDo
	else  //relatorio

		DbSelectArea('COMPATU')
		COMPATU->(DbGoTop())
		Count to nTotal
		ProcRegua(nTotal)
		COMPATU->(DbGoTop())

		//Somente se tiver dados
		If ! COMPATU->(EoF())
			//Criando o objeto de impressao
			oPrintReport := FWMSPrinter():New(;
				cFileName,;      // cFilePrinter
			IMP_PDF,;        // nDevice
			.F.,;            // lAdjustToLegacy
			,;               // cPathInServer
			.T.,;            // lDisabeSetup
			,;               // lTReport
			@oPrintReport,;  // oPrintSetup
			,;               // cPrinter
			,;               // lServer
			,;               // lParam10
			,;               // lRaw
			.T.;             // lViewPDF
			)
			oPrintReport:cPathPDF := GetTempPath()
			oPrintReport:SetResolution(72)
			oPrintReport:SetLandscape()
			oPrintReport:SetPaperSize(DMPAPER_A4)
			oPrintReport:SetMargin(0, 0, 0, 0)

			//Imprime os dados
			printHeader()
			While ! COMPATU->(EoF())
				nCurrent++
				IncProc('Imprimindo registro ' + cValToChar(nCurrent) + ' de ' + cValToChar(nTotal) + '...')

				//Se atingiu o limite, quebra de pagina
				validPageBreak()

				//Faz o zebrado ao fundo
				If nCurrent % 2 == 0
					oPrintReport:FillRect({nReportLine - 2, nLeftMargin, nReportLine + 12, nRightLimit}, oBrushLin)
				EndIf

				//Imprime a linha atual
				oPrintReport:SayAlign(nReportLine, nColData1, Alltrim(Transform(COMPATU->Codigo, '@!')), oFontDetails, 50, 10, /*nClrText*/, PAD_LEFT, /*nAlignVert*/)
				oPrintReport:SayAlign(nReportLine, nColData2, Alltrim(Transform(COMPATU->Loja, '@!')), oFontDetails, 30, 10, /*nClrText*/, PAD_LEFT, /*nAlignVert*/)
				oPrintReport:SayAlign(nReportLine, nColData3, Alltrim(Transform(COMPATU->Nome, '@!')), oFontDetails, 100, 10, /*nClrText*/, PAD_LEFT, /*nAlignVert*/)
				oPrintReport:SayAlign(nReportLine, nColData4, Alltrim(Transform(COMPATU->Numero, '@!')), oFontDetails, 90, 10, /*nClrText*/, PAD_LEFT, /*nAlignVert*/)
				oPrintReport:SayAlign(nReportLine, nColData5, Alltrim(Transform(COMPATU->Parcela, '@!')), oFontDetails, 30, 10, /*nClrText*/, PAD_LEFT, /*nAlignVert*/)
				oPrintReport:SayAlign(nReportLine, nColData6, Alltrim(Transform(COMPATU->TipoNF, '@!')), oFontDetails, 30, 10, /*nClrText*/, PAD_LEFT, /*nAlignVert*/)
				oPrintReport:SayAlign(nReportLine, nColData7, Alltrim(Transform(COMPATU->VTotalTitulos, "@E 999,999,999.99")), oFontDetailsBold, 80, 10, /*nClrText*/, PAD_LEFT, /*nAlignVert*/)
				oPrintReport:SayAlign(nReportLine, nColData8, Alltrim(Transform(COMPATU->VTotalEmAberto, "@E 999,999,999.99")), oFontDetailsBold, 80, 10, /*nClrText*/, PAD_LEFT, /*nAlignVert*/)


				nReportLine += 15
				oPrintReport:Line(nReportLine-3, nLeftMargin, nReportLine-3, nRightLimit, nColorGray)

				//Se atingiu o limite, quebra de pagina
				validPageBreak()

				COMPATU->(DbSkip())
			EndDo

			//Imprime o último rodapé
			printFooter()

			oPrintReport:Preview()


		endif
	endif
	COMPATU->(DbCloseArea())

Return

static function ValidPerg()

	Local aPergs    := {}
	Local cRelat    := " "
	Local cFPara    := " "
	Local cTipo     := space(100)
	local cTipd     := space(100)
	local cClide    := space(6)
	local cCliate   := 'ZZZZZZ'
	Local dData     := CtoD('')


	aAdd(aPergs, {1, "Data", dData,  "", ".T.", "", ".T.", 50, .T.})
	aAdd(aPergs, {1, "Tipos = a ? (sep usando ,)", cTipo,  "", ".T.", "", ".T.", 100, .F.})
	aAdd(aPergs, {1, "Tipos <>  de : (sep usando ,)", cTipd,  "", ".T.", "", ".T.", 100, .F.})
	aAdd(aPergs, {2, "Tipo Relatorio",         cRelat, {"1=Relatorio","2=Excel"},     122, ".T.", .F.})
	aAdd(aPergs, {2, "Se tipo = Relatorio",    cFPara, {"1=Cli. que nao compram","2=Cli que compraram <Sintetico>","3=Cli. que compraram <Analitico>"},  122, ".T.", .F.})
	aAdd(aPergs, {1, "Cliente de ", cClide,  "", ".T.", "SA1", ".T.", 80,  .f.})
	aAdd(aPergs, {1, "Cliente ate", cCliate,  "", ".T.", "SA1", ".T.", 80,  .f.})

	/*
	aAdd(_aPergs, {1, "Numero de *"						,cNumDe	,	"", ".T.",     "",        ".T.", 120, .F.})
	aAdd(_aPergs, {1, "Numero ate *"					,cNumAt	,	"",	".T.",     "",        ".T.", 120, .F.})
	aAdd(_aPergs, {1, "Mes/Ano *"						,cMesAn	,	"",	".T.",     "",        ".T.", 120, .F.})
	aAdd(_aPergs, {2, "Relatorio *"			    		,cRelat	,	{"1=Fluxo de Caixa", "2=Endividamento"},	090,	 ".T.",    .F.})
	aAdd(_aPergs, {2, "Financeiro *"			    	,cFPara	,	{"1=Pagar", "2=Receber"},	090,	 ".T.",    .F.})
	*/

	If ParamBox(aPergs, "Parâmetros", , , , , , , , RetCodUsr()+"AUTODEV",.T., .T.)
		FwMsgRun(NIL, {|oSay| RunMessage(oSay)}, "Processando", "Processando relatorio...")
	Else
		MsgAlert("Processo Cancelado pelo usuario","Cancel")
		Return
	EndIf
return
