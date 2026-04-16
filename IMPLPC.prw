#include "PROTHEUS.CH"
#include "RWMAKE.CH"
#include "TOPCONN.CH"
#include "AP5MAIL.CH"
#include "JPEG.CH"
#include "APWIZARD.CH"


#define cEndLin     chr(13)+chr(10)
#define NLININI     2 // Pula o cabeçalho do CSV

/*/
-----------------------------------------------------------------------------------------
-- Programa:    | IMPLPC       | Autor: Leandro Muniz 		 | Data:  15/04/2026       --
-----------------------------------------------------------------------------------------
-- Descrição:   | Alteração de data de entrega de Pedidos de Compra (SC7) em lote      --
-----------------------------------------------------------------------------------------
/*/
user function IMPLPC()
    //wizard
    local oWizard
    local oPanel
    local aParamBox := {}
    local oTHButton1A
    private oTM2B

    private cIdUser   := RetCodUsr() 
    private lPermite  := .F.
    private lRet        := .F.
    private lNext       := .F.
    private aList       := {}
    private aRet        := {}
    private aCelulas    := {}
    private aFields     := {}
    private cAliasTMP   := "TMPPC"
    private oTempPC
    private oList       
    private cTe2B       := "" 
    private oFont1A 	:= TFont():New("Arial",,-11,,.F.,,,,,.F.,.F.)

   //Cria a pasta C:\temp, caso não exista
   if !ExistDir("C:\TEMP")
    MakeDir("C:\TEMP")
   endif

    // Estrutura do arquivo CSV 
	//             Campo		Variável	Coluna Tipo	Tamanho			Decimal
    AADD(aCelulas,{'C7_NUM'     ,"cPedido"  , 1, 'C', TamSX3("C7_NUM")[1], 0})
    AADD(aCelulas,{'C7_ITEM'    ,"cItem"    , 2, 'C', TamSX3("C7_ITEM")[1], 0})
    AADD(aCelulas,{'C7_PRODUTO' ,"cProd"    , 3, 'C', TamSX3("C7_PRODUTO")[1], 0})
    AADD(aCelulas,{'C7_DATENT'  ,"dNovaDat" , 4, 'D', 8, 0})

    // Definição do Wizard
    define WIZARD oWizard TITLE "Alteração de Datas PC" ;
        HEADER "Alteração de Pedidos de Compra";
        MESSAGE "Esta rotina altera a data de entrega dos itens de PC via CSV.";
    next {|| IIF(empty(aRet[1]), ;
            (MsgAlert("Favor selecionar o arquivo.","Atenção"),.F.),;
            (MsAguarde({|| PImpExcel()},OemtoAnsi( "Validando Dados...")), lRet))};
    FINISH {|| .F. };
    PANEL

    // Etapa 1 - Seleção do Arquivo
        oPanel := oWizard:GetPanel(1)
        aRET := {space(150)}
        
        aAdd(aParamBox,{6,"Arquivo CSV",Space(150),"","","",150,.T.,"CSV (*.csv)|*.csv"})
        ParamBox(aParamBox,"Importação...",@aRet,,,,,,oPanel)

        oTHButton1A := THButton():New(15,0,"Baixar layout arquivo csv",oPanel,{|| MsgRun("Gerando arquivo Layout...","Importação PC",{|| LayoutPC() })  },100,18,oFont1A,"LayoutPC")

        oPanel:Refresh()

    // Etapa 2 - Visualização dos Dados Inconsistentes
    CREATE PANEL oWizard ;
        HEADER "Inconsistências do Arquivo" ;
        MESSAGE "Os registros abaixo apresentam erros e não serão processados." ;
        next {|| IIF(lNext,.T.,(MsgAlert("Corrija as inconsistências para avançar."),.F.))};
        PANEL

        oPanel := oWizard:GetPanel(2)

        oTM2B := TMultiGet():New(01,01,bSETGET(cTe2B),oPanel,260,92,,,,,,.T.)
        oTM2B:Align := CONTROL_ALIGN_ALLCLIENT

    // Etapa 3 - Visualização dos Dados Válidos
    CREATE PANEL oWizard ;
		HEADER "Itens para Alteração" ;
		MESSAGE "Confira os dados que serão atualizados no Protheus." ;
		BACK {|| .T. } ;
		next {|| IIF(EMPTY(aList[1,1]),.F.,MsAguarde({|| fGravaPC()})),.T. } ;
		FINISH {|| .F. } ;
		PANEL

		oPanel := oWizard:GetPanel(3)

		aList := {{"","","","","","",0,0,""}}

		oPanCss := TPanelCss():New(0,0,"",oPanel,,.F.,.F.,,,15,15,.T.,.F.)
		oPanCss:Align := CONTROL_ALIGN_TOP

		@ 0,0 ListBox oList Fields HEADER "Seq", "Pedido","Item","Produto","Nova Data" Size 100,100 of oPanel pixel
		oList:SetArray(aList)
		oList:bLine := {||{  aList[oList:Nat,1],;
                           aList[oList:Nat,2],;
                           aList[oList:Nat,3],;
                           aList[oList:Nat,4],;
                           aList[oList:Nat,5]}}

		oList:LHSCROLL := .T.
		oList:Align := CONTROL_ALIGN_ALLCLIENT
		oList:Refresh()

	//Etapa 4 - Importação
	CREATE PANEL oWizard ;
    	HEADER "Importação Finalizada!" ;
	    MESSAGE "Importação  Finalizado!" ;
	    BACK {|| .F. } ;
	    NEXT {|| .F. } ;
	    FINISH {|| .T. } ;
	    PANEL
	   oPanel := oWizard:GetPanel(4)

	activate WIZARD oWizard centered

    if valtype(oTempPC) == "O"
        oTempPC:Delete()
    endif
Return

/*---------------------------------------------------------------------*
 | Func:  LayoutPC                                      	           |
 | Autor: Leandro Muniz                                                |
 | Data:  15/04/2026                                                   |
 | Desc:  Layout modelo arquivo importação		                       |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
static function LayoutPC()
    local cMask     := "All csv files (*.csv) " 
    local cTit      := "Salvar Layout"
    local nOpcoes   := GETF_LOCALHARD
    local cDir      := "C:\temp\layout_pc.csv"
    local cArq      := tFileDialog(cMask,cTit,0,cDir,.T.,nOpcoes)
    local nHandle
    local aLay		:= {}

    if Empty(cArq)
        return
    endif       

    aadd(aLay,{"PEDIDO"	     ,"C7_NUM"})
	aadd(aLay,{"ITEM"	     ,"C7_ITEM"})
	aadd(aLay,{"PRODUTO"	 ,"C7_PRODUTO"})
	aadd(aLay,{"DATA_ENTREGA","C7_DATENT"})
	nHandle := MsfCreate(cArq,0)
	
	if nHandle > 0

		// Grava o cabecalho do arquivo
		aEval(aLay, {|e, nX| fWrite(nHandle, e[1] + If(nX < Len(aLay), ";", "") ) } )
					
		fClose(nHandle)

		cMessage := "::: Instruções de Preenchimento  :::" + CRLF
		cMessage += CRLF
		cMessage += "01 - As colunas da planilha não devem ser alteradas."+ CRLF
		cMessage += "02 - Não devem ser inseridas novas colunas."+ CRLF
		cMessage += "03 - Caso o campo quantidade ou valor tenham casas decimais, devem ser separadas por vírgula."+ CRLF
		cMessage += "04 - O arquivo deverá ser salvo no formato CSV (separado por vírgulas)."+ CRLF
		

		EecView(cMessage)

		if ! ApOleClient( "MsExcel" )
			MsgAlert( "MsExcel nao instalado!" + CRLF + "Acesse o arquivo através do caminho informado.")
			return
		else		
			oExcelApp := MsExcel():New()
			oExcelApp:WorkBooks:Open( cArq ) // Abre a planilha salva
			oExcelApp:SetVisible(.T.)
		endif
	else
		MsgAlert("Falha na criação do arquivo!")
	endif
    
return

/*---------------------------------------------------------------------*
 | Func:  PImpExcel                                      	           |
 | Autor: Leandro Muniz                                                |
 | Data:  15/04/2026                                                   |
 | Desc:  Leitura, Validação e Carga na Temporária	                   |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
static function PImpExcel()
    local oFile     := FwFileReader():New(aRet[1])
    local aLinha    := {}
    local aTabela   := {}
    local nReg, nX
    local lFalha    := .F.
    local cVar, nCol, cTipo
    private cPedido, cItem, cProd, dNovaDat
    private aLog    := {}

    lRet := .T.
    cTe2B := ""

    if select(cAliasTMP) <> 0
        (cAliasTMP)->(dbCloseArea())
    endif

    oTempPC := FWTemporaryTable():New(cAliasTMP)
    aFields := {}
    aadd(aFields,{"SEQ",     "C", 6, 0})
    aadd(aFields,{"PEDIDO",  "C", TamSX3("C7_NUM")[1], 0})
    aadd(aFields,{"ITEM",    "C", TamSX3("C7_ITEM")[1], 0})
    aadd(aFields,{"PRODUTO", "C", TamSX3("C7_PRODUTO")[1], 0})
    aadd(aFields,{"DATA_OK", "D", 8, 0})
    oTempPC:SetFields(aFields)
    oTempPC:Create()

    if oFile:Open()
        aLinha := oFile:GetAllLines()
        AEval(aLinha, {|x| AAdd(aTabela, StrTokArr2(x, ";", .T.))})
        oFile:Close()
    else
        MsgAlert("Erro ao abrir o arquivo CSV.")
        lRet := .F.
        return .F.
    endif

    dbSelectArea("SC7")
    SC7->(dbSetOrder(1)) 

    for nReg := NLININI to Len(aTabela)
        lFalha := .F.
        
        for nX := 1 to len(aCelulas)
            cVar  := alltrim(aCelulas[nX,2])
            nCol  := aCelulas[nX,3]
            cTipo := aCelulas[nX,4]
            if cTipo == "C"
                &cVar := PadR(alltrim(aTabela[nReg, nCol]), aCelulas[nX,5])
            elseif cTipo == "D"
                &cVar := CtoD(aTabela[nReg, nCol])
            endif
        next nX

        if !SC7->(dbSeek(xFilial("SC7") + cPedido + cItem))
            aAdd(aLog, "Linha " + cValToChar(nReg) + ": Pedido " + cPedido + " / Item " + cItem + " nao encontrado.")
            lFalha := .T.
        elseif alltrim(SC7->C7_PRODUTO) <> alltrim(cProd)
                aAdd(aLog, "Linha " + cValToChar(nReg) + ": Produto " + alltrim(cProd) + " divergente.")
                lFalha := .T.
        elseif !fValidaSAJ(cIdUser, SC7->C7_GRUPCOM)
            aAdd(aLog, "Linha " + cValToChar(nReg) + ": Pedido " + cPedido + " / Item " + cItem + " não faz parte do seu grupo.")
            lFalha := .T.
        endif

        if !lFalha
            RecLock(cAliasTMP, .T.)
                (cAliasTMP)->SEQ     := StrZero(nReg-1, 6)
                (cAliasTMP)->PEDIDO  := cPedido
                (cAliasTMP)->ITEM    := cItem
                (cAliasTMP)->PRODUTO := cProd
                (cAliasTMP)->DATA_OK := dNovaDat
            (cAliasTMP)->(MsUnlock())
        endif
    next nReg

    aList := {}
    (cAliasTMP)->(dbGoTop())
    while (cAliasTMP)->(!Eof())
        aAdd(aList, {(cAliasTMP)->SEQ, (cAliasTMP)->PEDIDO, (cAliasTMP)->ITEM, (cAliasTMP)->PRODUTO, dToC((cAliasTMP)->DATA_OK)})
        (cAliasTMP)->(dbSkip())
    enddo

    if Len(aList) == 0
        aAdd(aList, {"", "", "", "", ""})
        lNext := .F.
    else
        if Len(aLog) > 0
            lNext := .F.
        else
            lNext := .T. 
        endif
    endif
    
    oList:SetArray(aList)
    oList:bLine := { || {aList[oList:nAt,1], aList[oList:nAt,2], aList[oList:nAt,3], aList[oList:nAt,4], aList[oList:nAt,5]} }
    oList:Refresh()

    if Len(aLog) > 0
        cTe2B := "ERROS ENCONTRADOS NO ARQUIVO:" + CRLF
        AEval(aLog, {|x| cTe2B += x + CRLF})
        oTM2B:Refresh() 
    endif

return .T.
/*---------------------------------------------------------------------*
 | Func:  fGravaPC                                      	           |
 | Autor: Leandro Muniz                                                |
 | Data:  15/04/2026                                                   |
 | Desc:  Grava alterações				                               |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
static function fGravaPC()
    local nTot      := 0
    local nAtual    := 0
    local cIdUser   := RetCodUsr() 
    local lPermite  := .F.

    (cAliasTMP)->(dbEval({|| nTot++}))
    (cAliasTMP)->(dbGoTop())
    
    ProcRegua(nTot)
    dbSelectArea("SC7")
    SC7->(dbSetOrder(1))

    while (cAliasTMP)->(!Eof())
        nAtual++
        IncProc("Processando: " + cValToChar(nAtual) + " / " + cValToChar(nTot))

        if SC7->(dbSeek(xFilial("SC7") + (cAliasTMP)->PEDIDO + (cAliasTMP)->ITEM))
            
            // Validação
            if (SC7->C7_USER == cIdUser) .or. lPermite .or. fValidaSAJ(cIdUser, SC7->C7_GRUPCOM)
                if RecLock("SC7", .F.)
                    SC7->C7_DATPRF := (cAliasTMP)->DATA_OK
                    SC7->(MsUnlock())
                    SC7->(dbCommit())
                endif
            endif
            
        endif
        (cAliasTMP)->(dbSkip())
    enddo

    MsgInfo("Processo de gravação finalizado!", "Fim")
return
/*---------------------------------------------------------------------*
 | Func:  fValidaSAJ                                      	           |
 | Autor: Leandro Muniz                                                |
 | Data:  15/04/2026                                                   |
 | Desc:  Verifica se o usuário logado pertence ao grupo do pedido     |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
static function fValidaSAJ(cIdUser, cGrupo)
    local lRet      := .F.
    local cQuery    := ""
    local cAliasQ   := GetNextAlias()

    if Empty(cGrupo)
        return .F.
    endif

    cQuery := " SELECT AJ_USER "
    cQuery += " FROM " + RetSqlName("SAJ") + " "
    cQuery += " WHERE D_E_L_E_T_ = '' "
    cQuery += " AND AJ_FILIAL = '" + xFilial("SAJ") + "' "
    cQuery += " AND AJ_GRCOM  = '" + cGrupo + "' "
    cQuery += " AND AJ_USER   = '" + cIdUser + "' "

    TCQuery cQuery New Alias (cAliasQ)

    if (cAliasQ)->(!EOF())
        lRet := .T.
    endif

    (cAliasQ)->(dbCloseArea())

return lRet


/*----------------------------------------------------------------------------------------*
*                                CONTROLE DE MANUTENÇÃO                                   *
* ----------------------------------------------------------------------------------------*
*|               |            |           |                                              |*
*| Responsável   | Data       | Chamado   | Breve Descritivo                             |*
*|_______________|____________|___________|______________________________________________|*
*|               |            |           |                                              |*
*|               |            |           |                                              |*
*|               |            |           |                                              |*
*|               |            |           |                                              |*
*|               |            |           |                                              |*
*|               |            |           |                                              |*
*|_______________|____________|___________|______________________________________________|*    
*                                                                                         *
------------------------------------------------------------------------------------------*/
