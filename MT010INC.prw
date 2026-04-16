#Include 'Protheus.ch'
#INCLUDE "RWMAKE.CH"
#INCLUDE "TBICONN.CH"

//ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
//±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
//±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
//±±ºPrograma  ³ MT010INC ºAutor  ³ Santiago           ºData  ³  22/04/13   º±±
//±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
//±±ºDesc.     ³ Ponto de Entrada para Replicar o Cadastro de Produtos na   º±±
//±±º          ³ Inclusao                                                   º±±
//±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
//±±ºUso       ³ TESTE REPLICA DE CADASTRO DE PRODUTO CHAMADO 'THALCN'      º±±
//		       ³ 														    º±±
//Inclusão	   ³ Zerar campos de Ponto de Pedido - Chamado 220879			º±±
//		       ³ bhora / lmsantos	-  18/03/2026							º±±
//±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
//±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
//ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß

User Function MT010INC()
	Local j := 2

	PRIVATE _aStru     := {}
	PRIVATE _cEmp      := space(3)
	PRIVATE _aSaveArea := {Alias(),recno(),IndexOrd()}
	PRIVATE _cAlias      := "SB1"
	PRIVATE _nIndexOrder := 1
	PRIVATE _cFilial     := Space(2)
	Private aRotAuto    := {}
	


	if funname() <> "RPC"
		DBSELECTAREA("SB1")

		If Inclui .and. cEmpAnt == "01"

			_aStru := SB1->(dbStruct())
			aRotAuto    := {}
			For J:=2 to Len(_aStru)
				_cArqTrab := _astru[j,1]
				AADD(aRotAuto,{_astru[j,1],&(_cArqTrab)})
			Next

			_cEmp        := "02"
			_cFiliaL     := "01" //SM0->M0_CODFIL
			_cAlias      := "SB1"
			_nIndexOrder := 1

			qout(funname() + " - " + _cAlias + " - " + _cemp+_cfilial)
			STARTJOB("U_IncProd",getenvserver(),.t.,_cEmp,_cFilial,@aRotAuto)
		endif

		If IsInCallStack("A010Copia") .Or. lCopia //Variável que traz quando é cópia   //MATA010 .Or. IsInCallStac("A010Copia")
			RecLock("SB1",.F.)
			B1_LFISCLS 	:= ""
			B1_LFISCLT 	:= ""
			B1_LFISCTL 	:= ""
			B1_LFISCCM :=  ""
			B1_XPISCO 	:= ""
			B1_XREVFIS	:= CTOD("  /  /  ") //Chamado - 92940
			B1_XDTNCM 	:= CTOD("  /  /  ")    
			B1_DTINCL 	:= dDataBase

			// Campos de Ponto de Pedido - Chamado 220879
			B1_QESUL    := 0
			B1_EMINSUC  := 0
			B1_QE       := 0
			B1_EMIN     := 0
			B1_QECTL    := 0
			B1_EMINCTL  := 0
			B1_QEUTR	:= 0
			B1_EMINUTR	:= 0
			B1_QECMT	:= 0
			B1_EMINCMT	:= 0
			B1_EMINTSA	:= 0
			B1_QETSA	:= 0
			B1_QESUD 	:= 0
			B1_EMINSUD 	:= 0

			MsUnlock()
		EndIf

	Endif
	dbSelectArea(_aSaveArea[1])
	dbSetOrder(_aSaveArea[3])
	dbGoto(_aSaveArea[2])

Return .T.

//±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
//±±ºPrograma  ³ IncProd  ºAutor  ³ Fmc                ºData ³  14/08/2014 º±±
//±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
User Function IncProd(_cEmpa,_cFiliala,aRotAuto)
	Local j := 0

	Private nOpc        :=  3 // inclusao
	Private lMsHelpAuto := .T. // se .t. direciona as mensagens de help para o arq. de log
	Private lMsErroAuto := .F. //necessario a criacao, pois sera atualizado quando houver alguma incosistencia nos parametros
	Private aAutoErro := {}

	qout("Abrindo Empresa " + _cEmpa + " Filial " + _cFiliala )
	RpcSetType( 3 )
	RpcSetEnv( _cEmpa, _cFiliala )

	DBSELECTAREA("SB1")
	RecLock("SB1",.T.)

	For j := 1 to len(aRotAuto)
		SB1->&(aRotAuto[J][1]) := aRotAuto[J][2]
	Next

	MsUnlock()

	qout("Fechando Empresa " + _cEmpa + " Filial " + _cFiliala )
	RESET environment


Return .t.
