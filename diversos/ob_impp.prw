// Bibliotecas
#Include "Protheus.ch"
#Include "TopConn.ch"
#Include "RPTDef.ch"
#Include "FWPrintSetup.ch"

// Alinhamentos
#Define PAD_LEFT    0
#Define PAD_RIGHT   1
#Define PAD_CENTER  2

// Cores
#Define COR_CINZA   RGB(180, 180, 180)
#Define COR_PRETO   RGB(000, 000, 000)

// Posições do Array
/*
Static nPosMeto   := 1
Static nPosID     := 2
Static nPosUser   := 3
Static nPosStatus := 4
Static nPosDire   := 5
Static nPosHor    := 6
*/

// Posições do Array
Static nPosMeto   := 1
Static nPosID     := 2
Static nPosUser   := 3
Static nPosStatus := 4
Static nPosHor    := 5

/*/{Protheus.doc} zImpCSV
Função para importar informações do fornecedor via csv
@author Atilio
@since 07/06/2021
@version 1.0
@type function
/*/

User Function ob_zd6()
	AxCadastro("ZD6", "Registro apontamento", ".T.", ".T.")
Return

User Function ob_zD7()
	AxCadastro("ZD7", "Cadastro de terceiros", ".T.", ".T.")
Return

User Function ob_impp() //u_ob_impp
	Local aArea   := GetArea()
	Local cfilter := "Arquivos CSV   (*.CSV)  | *.CSV"
	Local lSave   := .F.

	Private cArqOri := ""

	cArqOri := U_OB_FilDlg(cfilter, lSave)

	If !Empty(cArqOri)
		If File(cArqOri) .And. Upper(SubStr(cArqOri, RAt('.', cArqOri) + 1, 3)) == 'CSV'
			FWMsgRun(, {|oSay| fImporta(oSay) }, "Processando", "Importando dados...")
		Else
			MsgStop("Arquivo e/ou extensão inválida!", "Atenção")
		EndIf
	EndIf

	RestArea(aArea)
Return


User Function OB_FilDlg(cfilter, lSave)
	Local tmp       := GetTempPath()
	Local targetDir := ""

	DEFINE DIALOG oDlg TITLE "**** Abrir Arquivo ****" FROM 0,0 TO 100,200 PIXEL

	oBtnFIle := TButton():New(002, 002, "Selecionar arquivo", oDlg, ;
		{|| targetDir := TFileDialog(cfilter, 'Selecao de Arquivos', , tmp, lSave), oDlg:End()}, ;
		100, 45, , , .F., .T., .F., , .F., , , .F.)

	ACTIVATE DIALOG oDlg CENTERED
	
Return targetDir

Static Function fImporta(oSay)

	Local aArea    := FWGetArea()
	Local nTotLinhas := 0
	Local cLinAtu    := ""
	Local nLinhaAtu  := 0
	Local aLinha     := {}
	Local oArqTab
	Local aLinhas
	private cDtade
	private cDtate
	private lPrmndta := .t.

	oSay:SetText("Abrindo arquivo...")
	ProcessMessages()

	oArqTab := FWFileReader():New(cArqOri)

	If (oArqTab:Open())
		// Se não for fim do arquivo
		If !(oArqTab:EoF())

			// Definindo o tamanho da régua

			aLinhas    := oArqTab:GetAllLines()
			nTotLinhas := Len(aLinhas)
			ProcRegua(nTotLinhas)

			// GoTop não funciona em algumas LIBs, então reabre
			oArqTab:Close()
			oArqTab := FWFileReader():New(cArqOri)
			oArqTab:Open()
			While (oArqTab:HasLine())
				nLinhaAtu++
				cLinAtu := oArqTab:GetLine()
				aLinha  := StrTokArr(cLinAtu, ";")
				oSay:SetText("Gravando registros...")
				ProcessMessages()

				if !"FALHOU" $ UPPER(cLinAtu) .and. !";id;" $ Lower(cLinAtu) .and. upper('Sucesso') $ upper(alltrim(aLinha[nPosStatus]))

					cData := aLinha[nPosHor]
					if lPrmndta
						cDtade := ctod(Left(StrTran(cData, "-", "/"), 10))
						cDtate := ctod(Left(StrTran(cData, "-", "/"), 10))
						lPrmndta:= .F.

					else
						if ctod(Left(StrTran(cData, "-", "/"), 10)) > cDtate
							cDtate := ctod(Left(StrTran(cData, "-", "/"), 10))
						endif
						if ctod(Left(StrTran(cData, "-", "/"), 10)) < cDtade
							cDtade := ctod(Left(StrTran(cData, "-", "/"), 10))
						endif
					endif

				endif
			EndDo

			oSay:SetText("Limpando tabela...")
			ProcessMessages()
			AtSql := " DELETE ZD6010 "
			AtSql += " WHERE ZD6_DATA BETWEEN '"+DTOS(cDtade)+"' AND '"+DTOS(cDtate)+"' and ZD6_CLASS <> '1' "

			TCSQLExec(AtSql)


			oArqTab:Close()
			oArqTab := FWFileReader():New(cArqOri)
			oArqTab:Open()
			While (oArqTab:HasLine())
				nLinhaAtu++
				cLinAtu := oArqTab:GetLine()
				aLinha  := StrTokArr(cLinAtu, ";")
				oSay:SetText("Gravando registros...")
				ProcessMessages()

				if !"FALHOU" $ UPPER(cLinAtu) .and. !";id;" $ Lower(cLinAtu) .and. upper('Sucesso') $ upper(alltrim(aLinha[nPosStatus]))
					cID   := aLinha[nPosID]
					cData := aLinha[nPosHor]
					cUser := aLinha[nPosUser]
					nMarc := Right(cData, 8)

					RecLock('ZD6',.t.)
					ZD6_ID   := cID
					ZD6_USER := cUser
					ZD6_DATA := ctod(Left(StrTran(cData, "-", "/"),10))
					ZD6_MARC := Val(StrTran(nMarc, ":", "."))
					MsUnlock()

				endif
			EndDo

			oArqTab:Close()
		Else
			MsgStop("Arquivo não pode ser aberto!", "Atenção")
		EndIf
	endif
	FWRestArea(aArea)

return


user  Function ob_Impm() //u_ob_Impm()
	ValidPerg()
Return


Static Function RunMessage(oSay)

	Local aLinha     := {}
	Local lPrim      := .T.
	local lReg       := .f.
	local cIdAdmin   := '  '


	// Linhas e colunas
	Private nLinAtu  := 000
	Private nTamLin  := 010
	Private nLinFin  := 820
	Private nColIni  := 010
	Private nColFin  := 550
	Private nColMeio := (nColFin - nColIni) / 2

	// Impressão
	Private oPrintPvt
	Private dDataGer := Date()
	Private cHoraGer := Time()
	Private nPagAtu  := 1
	Private cNomeUsr := UsrRetName(RetCodUsr())

	Private cNomeFont := "Arial"
	Private oFontDet  := TFont():New(cNomeFont, 9, -10, .T., .F., 5, .T., 5, .T., .F.)
	Private oFontDetN := TFont():New(cNomeFont, 9, -10, .T., .T., 5, .T., 5, .T., .F.)
	Private oFontRod  := TFont():New(cNomeFont, 9, -08, .T., .F., 5, .T., 5, .T., .F.)
	Private oFontTit  := TFont():New(cNomeFont, 9, -13, .T., .T., 5, .T., 5, .T., .F.)

	Private cUser      := ''
	Private cData      := ''
	Private cId        := ''
	Private cPriEnt    := '00:00'
	Private cPriSaida  := '00:00'
	Private cSegEntr   := '00:00'
	Private cSegSaida  := '00:00'
	Private cTerEnt    := '00:00'
	Private cTerSaida  := '00:00'
	Private cQuarEnt   := '00:00'
	Private cQuarSaida := '00:00'

	Private COL_USER := 015
	Private COL_DATA := 015
	Private COL_PRIM := COL_DATA + 60
	Private COL_SEGU := COL_PRIM + 40
	Private COL_TERC := COL_SEGU + 40
	Private COL_QUAR := COL_TERC + 40
	Private COL_QUIN := COL_QUAR + 40
	Private COL_SEXT := COL_QUIN + 40
	Private COL_SETI := COL_SEXT + 40
	Private COL_OITA := COL_SETI + 40
	Private COL_RESUL:= COL_OITA + 40

	Private limpaper := .T.
	Private nSubTot  := 0.00

	private cArquivo  := "relap_" + dToS(dDataGer) + "_" + StrTran(cHoraGer, ':', '-')


	cQry := "  SELECT  * FROM  ZD6010 ZD6"
	cQry += "  WHERE ZD6.D_E_L_E_T_ = '' AND ZD6_DATA between '"+ dtos(MV_PAR01) +"' and '"+dtos(MV_PAR02) +"'
	cQry += "  AND ZD6_CLASS <> '3'"
	If FWIsAdmin()
		cIdAdmin:= FWInputBox("Digite um id especifico",cIdAdmin)
		if alltrim(cIdAdmin) <> ''
			cQry += " AND ZD6_ID = '"+cIdAdmin+"'"
		endif
	EndIf
	cQry += " order by  ZD6_ID,ZD6_DATA,ZD6_MARC"

	TCQUERY cQry NEW ALIAS "TRBC"

	If FWIsAdmin()
		ShowLog(cQry)
	EndIf

	dbSelectArea('TRBC')
	dbgotop()
	Do while !eof() .and. !lReg
		lReg := .t.
		dbskip()
	enddo

	if lreg

		oPrintPvt := FWMSPrinter():New(cArquivo, IMP_PDF, .F., "", .T., , @oPrintPvt, "", , , , .T.)
		oPrintPvt:SetResolution(72)
		oPrintPvt:SetPortrait()
		oPrintPvt:SetPaperSize(DMPAPER_A4)
		oPrintPvt:SetMargin(60, 60, 60, 60)

		cUsebkp  := 'ZZZZZZZZZ'
		cIdBackup := 'ZZZ'
		cDataBkp := '99999999'
		nBat     := 1

		dbSelectArea('TRBC')
		dbgotop()
		While !eof()
			cUser := Alltrim(TRBC->ZD6_USER)
			cId   := Alltrim(TRBC->ZD6_ID)
			cData := TRBC->ZD6_DATA
			cMarc := HoraNumToStr(TRBC->ZD6_MARC)

			If Alltrim(cIdBackup) <> Alltrim(cId) .Or. cData <> cDataBkp
				nBat := 1

				If !lPrim
					//oPrintPvt:SayAlign(nLinAtu, COL_USER, Nome do contratado: cUsebkp              ,  oFontDet, 500, nTamLin, COR_PRETO, PAD_LEFT, 0)
					oPrintPvt:SayAlign(nLinAtu, COL_DATA, dtoc(stod(cDataBkp)) ,  oFontDet, 500, nTamLin, COR_PRETO, PAD_LEFT, 0)
					oPrintPvt:SayAlign(nLinAtu, COL_PRIM, cPriEnt              ,  oFontDet, 500, nTamLin, COR_PRETO, PAD_LEFT, 0)
					oPrintPvt:SayAlign(nLinAtu, COL_SEGU, cPriSaida            ,  oFontDet, 500, nTamLin, COR_PRETO, PAD_LEFT, 0)
					oPrintPvt:SayAlign(nLinAtu, COL_TERC, cSegEntr             ,  oFontDet, 500, nTamLin, COR_PRETO, PAD_LEFT, 0)
					oPrintPvt:SayAlign(nLinAtu, COL_QUAR, cSegSaida            ,  oFontDet, 500, nTamLin, COR_PRETO, PAD_LEFT, 0)
					oPrintPvt:SayAlign(nLinAtu, COL_QUIN, cTerEnt              ,  oFontDet, 500, nTamLin, COR_PRETO, PAD_LEFT, 0)
					oPrintPvt:SayAlign(nLinAtu, COL_SEXT, cTerSaida            ,  oFontDet, 500, nTamLin, COR_PRETO, PAD_LEFT, 0)
					oPrintPvt:SayAlign(nLinAtu, COL_SETI, cQuarEnt             ,  oFontDet, 500, nTamLin, COR_PRETO, PAD_LEFT, 0)
					oPrintPvt:SayAlign(nLinAtu, COL_OITA, cQuarSaida           ,  oFontDet, 500, nTamLin, COR_PRETO, PAD_LEFT, 0)

					cResult := bCalcRe(cPriEnt,cPriSaida,cSegEntr,cSegSaida,cTerEnt,cTerSaida,cQuarEnt,cQuarSaida)
					nSubTot+= HrSexaToDec(HrStrToNum(cResult))
					oPrintPvt:SayAlign(nLinAtu, COL_RESUL, cResult             ,  oFontDet, 500, nTamLin, COR_PRETO, PAD_LEFT, 0)

					if cTipBat == 'impar'
						oPrintPvt:SayAlign(nLinAtu, COL_RESUL + 30, "**verificar**"             ,  oFontDet, 500, nTamLin, COR_PRETO, PAD_LEFT, 0)
					endif


					// reinicializa variaveis
					cPriEnt    := '00:00'
					cPriSaida  := '00:00'
					cSegEntr   := '00:00'
					cSegSaida  := '00:00'
					cTerEnt    := '00:00'
					cTerSaida  := '00:00'
					cQuarEnt   := '00:00'
					cQuarSaida := '00:00'

					nLinAtu += nTamLin

					If (nLinAtu + nTamLin > nLinFin) .or. Alltrim(cIdBackup) <> Alltrim(cId)
						nLinAtu += nTamLin
						//--imprimir assinatura e totalizador
						oPrintPvt:SayAlign(nLinAtu, COL_SETI, 'Total:'   ,  oFontDet, 500, nTamLin, COR_PRETO, PAD_LEFT, 0)
						oPrintPvt:SayAlign(nLinAtu, COL_RESUL, DecToHora(nSubTot)   ,  oFontDet, 500, nTamLin, COR_PRETO, PAD_LEFT, 0)

						nLinAtu += nTamLin + 10
						oPrintPvt:SayAlign(nLinAtu, COL_DATA,'Lajeado '+dtoc(ddatabase) ,  oFontDet, 500, nTamLin, COR_PRETO, PAD_LEFT, 0)
						nLinAtu += nTamLin + 30
						oPrintPvt:SayAlign(nLinAtu, COL_DATA,'Ass Empresa ___________________________________. Ass fornecedor: ___________________________________'    ,  oFontDet, 500, nTamLin, COR_PRETO, PAD_LEFT, 0)

						fImpRod()
						fImpCab(cId+'-'+cUser)
						nSubTot:=0.00
					EndIf


				Else
					lPrim := .F.
					fImpCab(Alltrim(cId+'-'+cUser))
				EndIf

				cIdBackup  := cId
				cDataBkp := cData
				cPriEnt  := cMarc
				cTipBat := "impar"
			Else

				nBat++

				Do Case
				Case nBat == 2 ; cPriSaida   := cMarc ; cTipBat := "par"
				Case nBat == 3 ; cSegEntr    := cMarc ; cTipBat := "impar"
				Case nBat == 4 ; cSegSaida   := cMarc ; cTipBat := "par"
				Case nBat == 5 ; cTerEnt     := cMarc ; cTipBat := "impar"
				Case nBat == 6 ; cTerSaida   := cMarc ; cTipBat := "par"
				Case nBat == 7 ; cQuarEnt    := cMarc ; cTipBat := "impar"
				Case nBat == 8 ; cQuarSaida  := cMarc ; cTipBat := "par"
				EndCase

			EndIf

			dbSelectArea('TRBC')
			DbSkip()
		enddo


		If (nLinAtu + nTamLin > nLinFin)
			fImpRod()
			fImpCab()
		EndIf

		//Imprimindo a linha atual
		//oPrintPvt:SayAlign(nLinAtu, COL_USER, cUsebkp              ,  oFontDet, 500, nTamLin, COR_PRETO, PAD_LEFT, 0)
		oPrintPvt:SayAlign(nLinAtu, COL_DATA, dtoc(stod(cDataBkp)) ,  oFontDet, 500, nTamLin, COR_PRETO, PAD_LEFT, 0)
		oPrintPvt:SayAlign(nLinAtu, COL_PRIM, cPriEnt              ,  oFontDet, 500, nTamLin, COR_PRETO, PAD_LEFT, 0)
		oPrintPvt:SayAlign(nLinAtu, COL_SEGU, cPriSaida            ,  oFontDet, 500, nTamLin, COR_PRETO, PAD_LEFT, 0)
		oPrintPvt:SayAlign(nLinAtu, COL_TERC, cSegEntr             ,  oFontDet, 500, nTamLin, COR_PRETO, PAD_LEFT, 0)
		oPrintPvt:SayAlign(nLinAtu, COL_QUAR, cSegSaida            ,  oFontDet, 500, nTamLin, COR_PRETO, PAD_LEFT, 0)
		oPrintPvt:SayAlign(nLinAtu, COL_QUIN, cTerEnt              ,  oFontDet, 500, nTamLin, COR_PRETO, PAD_LEFT, 0)
		oPrintPvt:SayAlign(nLinAtu, COL_SEXT, cTerSaida            ,  oFontDet, 500, nTamLin, COR_PRETO, PAD_LEFT, 0)
		oPrintPvt:SayAlign(nLinAtu, COL_SETI, cQuarEnt             ,  oFontDet, 500, nTamLin, COR_PRETO, PAD_LEFT, 0)
		oPrintPvt:SayAlign(nLinAtu, COL_OITA, cQuarSaida           ,  oFontDet, 500, nTamLin, COR_PRETO, PAD_LEFT, 0)

		cResult := bCalcRe(cPriEnt,cPriSaida,cSegEntr,cSegSaida,cTerEnt,cTerSaida,cQuarEnt,cQuarSaida)
		oPrintPvt:SayAlign(nLinAtu, COL_RESUL, cResult             ,  oFontDet, 500, nTamLin, COR_PRETO, PAD_LEFT, 0)

		if cTipBat == 'impar'
			oPrintPvt:SayAlign(nLinAtu, COL_RESUL + 30, "**verificar**"             ,  oFontDet, 500, nTamLin, COR_PRETO, PAD_LEFT, 0)
		endif


		nSubTot+= HrSexaToDec(HrStrToNum(cResult))

		nLinAtu += nTamLin
		oPrintPvt:SayAlign(nLinAtu, COL_SETI, 'Total:'   ,  oFontDet, 500, nTamLin, COR_PRETO, PAD_LEFT, 0)
		oPrintPvt:SayAlign(nLinAtu, COL_RESUL, DecToHora(nSubTot)   ,  oFontDet, 500, nTamLin, COR_PRETO, PAD_LEFT, 0)

		nLinAtu += nTamLin + 10
		oPrintPvt:SayAlign(nLinAtu, COL_DATA,'Lajeado '+dtoc(ddatabase) ,  oFontDet, 500, nTamLin, COR_PRETO, PAD_LEFT, 0)
		nLinAtu += nTamLin + 30
		oPrintPvt:SayAlign(nLinAtu, COL_DATA,'Ass Empresa ___________________________________. Ass fornecedor: ___________________________________'    ,  oFontDet, 500, nTamLin, COR_PRETO, PAD_LEFT, 0)



		nLinAtu += nTamLin			//Se ainda tiver linhas sobrando na página, imprime o rodapé final
		If nLinAtu <= nLinFin
			fImpRod()
		EndIf

		//Mostrando o relatório
		oPrintPvt:Preview()
	endif
	TRBC->(DbCloseArea())
return


Static Function fImpRod()
	Local nLinRod   := nLinFin + nTamLin
	Local cTextoEsq := ''
	Local cTextoDir := ''

	//Linha Separatória
	oPrintPvt:Line(nLinRod, nColIni, nLinRod, nColFin, COR_CINZA)
	nLinRod += 3

	//Dados da Esquerda e Direita
	cTextoEsq := dToC(dDataGer) + "    " + cHoraGer + "    " + FunName() + "    " + cNomeUsr
	cTextoDir := "Página " + cValToChar(nPagAtu)

	//Imprimindo os textos
	oPrintPvt:SayAlign(nLinRod, nColIni,    cTextoEsq, oFontRod, 200, 05, COR_CINZA, PAD_LEFT,  0)
	oPrintPvt:SayAlign(nLinRod, nColFin-40, cTextoDir, oFontRod, 040, 05, COR_CINZA, PAD_RIGHT, 0)

	//Finalizando a página e somando mais um
	oPrintPvt:EndPage()
	nPagAtu++
Return

Static Function fImpCab(cUser)
	Local cTexto   := ""
	Local nLinCab  := 030
	local nLinTit  := 0

	//Iniciando Página
	oPrintPvt:StartPage()

	//Cabeçalho
	cTexto := "Registro de acesso a empresa"
	oPrintPvt:SayAlign(nLinCab, nColMeio - 120, cTexto, oFontTit, 240, 20, COR_CINZA, PAD_CENTER, 0)

	//Linha Separatória
	nLinCab += (nTamLin * 2)
	oPrintPvt:Line(nLinCab, nColIni, nLinCab, nColFin, COR_CINZA)

	//Cabeçalho das colunas
	nLinCab += nTamLin

	oPrintPvt:SayAlign(nLinCab, COL_USER, "Nome do contratado: "+cUser ,  oFontTit, 500, nTamLin, COR_PRETO, PAD_LEFT, 0)
	nLinCab += nTamLin

	//oPrintPvt:SayAlign(nLinCab, COL_USER, 'Usuario'   , oFontDetN, 0500, nTamLin, COR_PRETO, PAD_LEFT, 0)
	oPrintPvt:SayAlign(nLinCab, COL_DATA, 'Data'      , oFontDetN, 0500, nTamLin, COR_PRETO, PAD_LEFT, 0)
	oPrintPvt:SayAlign(nLinCab, COL_PRIM, '1 entrada' , oFontDetN, 0500, nTamLin, COR_PRETO, PAD_LEFT, 0)
	oPrintPvt:SayAlign(nLinCab, COL_SEGU, '1 Saida  ' , oFontDetN, 0500, nTamLin, COR_PRETO, PAD_LEFT, 0)
	oPrintPvt:SayAlign(nLinCab, COL_TERC, '2 Entrada' , oFontDetN, 0500, nTamLin, COR_PRETO, PAD_LEFT, 0)
	oPrintPvt:SayAlign(nLinCab, COL_QUAR, '2 Saida'   , oFontDetN, 0500, nTamLin, COR_PRETO, PAD_LEFT, 0)
	oPrintPvt:SayAlign(nLinCab, COL_QUIN, '3 Entrada' , oFontDetN, 0500, nTamLin, COR_PRETO, PAD_LEFT, 0)
	oPrintPvt:SayAlign(nLinCab, COL_SEXT, '3 Saida'   , oFontDetN, 0500, nTamLin, COR_PRETO, PAD_LEFT, 0)
	oPrintPvt:SayAlign(nLinCab, COL_SETI, '4 Entrada' , oFontDetN, 0500, nTamLin, COR_PRETO, PAD_LEFT, 0)
	oPrintPvt:SayAlign(nLinCab, COL_OITA, '4 Saida '  , oFontDetN, 0500, nTamLin, COR_PRETO, PAD_LEFT, 0)
	oPrintPvt:SayAlign(nLinCab, COL_RESUL, 'Quant   '  , oFontDetN, 0500, nTamLin, COR_PRETO, PAD_LEFT, 0)

	nLinCab += nTamLin

	//Atualizando a linha inicial do relatório
	nLinAtu := nLinCab + 3
Return


static function ValidPerg()

	Local aPergs    := {}
	Local dDataDe  := FirstDate(Date())
	Local dDataAt  := LastDate(Date())

	aAdd(aPergs, {1, "Data De",  dDataDe,  "", ".T.", "", ".T.", 80,  .F.})
	aAdd(aPergs, {1, "Data Até", dDataAt,  "", ".T.", "", ".T.", 80,  .T.})



	If ParamBox(aPergs, "Parâmetros", , , , , , , , RetCodUsr()+"AUTODED",.T., .T.)
		FwMsgRun(NIL, {|oSay| RunMessage(oSay)}, "Processando", "Processando relatorio...")
	Else
		MsgAlert("Processo Cancelado pelo usuario","Cancel")
	EndIf
return



static  Function HoraNumToStr(nHora)
	Local nHoras   := Int(nHora)                        // Parte inteira = horas
	Local nMinutos := Round((nHora - nHoras) * 100, 0)  // Parte decimal = minutos
	Local cHora    := ""

	cHora := StrZero(nHoras, 2) + ":" + StrZero(nMinutos, 2)

Return cHora



static function bCalcRe(cPriEnt,cPriSaida,cSegEntr,cSegSaida,cTerEnt,cTerSaida,cQuarEnt,cQuarSaida)
	local nTotal := 0.00
	local cResult
	n1Ent  := HrSexaToDec(HrStrToNum(cPriEnt))
	n1Said := HrSexaToDec(HrStrToNum(cPriSaida))
	n2Ent  := HrSexaToDec(HrStrToNum(cSegEntr))
	n2SAid := HrSexaToDec(HrStrToNum(cSegSaida))
	n3Ent  := HrSexaToDec(HrStrToNum(cTerEnt))
	n3SAid := HrSexaToDec(HrStrToNum(cTerSaida))
	n4Ent  := HrSexaToDec(HrStrToNum(cQuarEnt))
	n4SAid := HrSexaToDec(HrStrToNum(cQuarSaida))

	if n1Said > 0
		nTotal += n1Said - n1Ent
	endif

	if n2SAid > 0
		nTotal += n2SAid - n2Ent
	endif

	if n3SAid > 0
		nTotal += n3SAid - n3Ent
	endif

	if n4SAid > 0
		nTotal += n4SAid - n4Ent
	endif

	cResult := DecToHora(nTotal)

return(cResult)


static  Function HrStrToNum(cHora)
	Local nHora := 0
	Local nMin  := 0
	Local nRet  := 0

	// Validação básica
	If Empty(cHora) .Or. !( ":" $ cHora )
		Return 0
	EndIf

	// Extrai hora e minuto
	nHora := Val( SubStr(cHora, 1, 2) )
	nMin  := Val( SubStr(cHora, 4, 2) )

	// Converte para número no formato HH,MM
	nRet := nHora + (nMin / 100)

Return nRet


static  Function HrSexaToDec(nHoraSexa)
	Local nHora  := Int(nHoraSexa)
	Local nMin   := 0
	Local nDec   := 0

	// Pega os "minutos" após a vírgula
	nMin := Round((nHoraSexa - nHora) * 100, 0)

	// Converte minutos (base 60) para decimal
	nDec := nHora + (nMin / 60)

Return nDec


static Function DecToHora(nHoraDecimal)
	Local nHoras    := Int(nHoraDecimal)
	Local nMinutos  := Round((nHoraDecimal - nHoras) * 60, 0)
	Local cResultado := ""

	cResultado := StrZero(nHoras,iif(nHoras>=100,3,2)) + ":" + StrZero(nMinutos,2)

Return cResultado
