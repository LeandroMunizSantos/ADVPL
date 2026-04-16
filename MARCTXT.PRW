#Include "TOTVS.ch"
#Include 'Protheus.ch'
#INCLUDE "rwmake.ch"

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³MARCTXT   ºAutor  Leandro Muniz         º Data ³  28/01/26  º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³Relatório de marcação x arquivo TXT						  º±±
±±º          ³				                                              º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ DEPARTAMENTO PESSOAL                                       º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

User Function MARCTXT()
	
	Local _cQuery		:= ""
	Local _aResultQry	:= {}
	Local _aStrut		:= {}
	local _cPerg 		:= "MARCTXT"

	Pergunte(_cPerg,.t.)
    
    cDataDe  := DtoS(MV_PAR01)
    cDataAte := DtoS(MV_PAR02)

    cDia := Right(cDataAte, 2)
    cMes := SubStr(cDataAte, 5, 2)
    cAno := Left(cDataAte, 4)
    If cDia == "30" .or. cDia == "31"
        If cMes == "12"
            cAno := StrZero( Val(cAno) + 1, 4 )
            cMes := "01"
            cDia := "01"
        Else
            cMes := StrZero( Val(cMes) + 1, 2 )
            cDia := "01"
        EndIf
        cDataAte := cAno + cMes + cDia
    Else
        cDataAte := StrZero( Val(cDataAte) + 1, 8 )
    EndIf

	
	aAdd(_aStrut,{"Matricula"		 ,"C",06,0})
	aAdd(_aStrut,{"Data da Marcação" ,"C",10,0})
	aAdd(_aStrut,{"Hora da Marcação" ,"C",05,0})
	aAdd(_aStrut,{"Numero do Relógio","C",03,0})
	aAdd(_aStrut,{"Intrgrado?"		 ,"C",03,0})

	
    	_cQuery := "SELECT"+CRLF
		_cQuery += "	RIGHT(CAST(C.CODIGO AS VARCHAR(20)), 6) AS Matricula, "+CRLF
		_cQuery += "	FORMAT(C.DATAHORA, 'dd/MM/yyyy') AS DataMarcacao, "+CRLF
		_cQuery += "	FORMAT(C.DATAHORA, 'HH:mm') AS HoraMarcacao, "+CRLF
		_cQuery += "	'P' + FORMAT(C.IDRELOGIO, '00') AS NumRelogio, "+CRLF
		_cQuery += "	'NÃO' AS Integrado, "+CRLF
		_cQuery += "	'' AS Del"+CRLF
		_cQuery += "FROM [LinkedApp].[SERVCOM_NET].[dbo].[COLETA] C "+CRLF
		_cQuery += "WHERE C.DATAHORA BETWEEN '" + cDataDe + "' AND '" + cDataAte + "' "+CRLF
		_cQuery += "	AND C.DATAHORA < DATEADD(DAY,1,'" + cDataAte + "') "+CRLF
		_cQuery += "	AND NOT EXISTS (SELECT 1 FROM SP8010 S WHERE S.D_E_L_E_T_ = '' AND S.P8_MAT COLLATE Latin1_General_BIN = RIGHT(CAST(C.CODIGO AS VARCHAR(20)), 6) COLLATE Latin1_General_BIN "+CRLF
		_cQuery += "	AND S.P8_DATA COLLATE Latin1_General_BIN = CONVERT(VARCHAR(8), DATEADD(MINUTE, DATEDIFF(MINUTE, 0, C.DATAHORA), 0), 112) COLLATE Latin1_General_BIN "+CRLF    
		_cQuery += "	AND CAST( "+CRLF    
		_cQuery += "		FLOOR(CAST(REPLACE(S.P8_HORA, ',', '.') AS DECIMAL(10,2))) * 100 + "+CRLF    
		_cQuery += "		((CAST(REPLACE(S.P8_HORA, ',', '.') AS DECIMAL(10,2)) - FLOOR(CAST(REPLACE(S.P8_HORA, ',', '.') AS DECIMAL(10,2)))) * 100) "+CRLF    
		_cQuery += "		AS INT"+CRLF    
		_cQuery += "	) = CAST(FORMAT(DATEADD(MINUTE, DATEDIFF(MINUTE, 0, C.DATAHORA), 0), 'HHmm') AS INT)"+CRLF    
		_cQuery += ")"+CRLF    
		_cQuery += "	AND LEFT(FORMAT(C.DATAHORA, 'dd/MM/yyyy'), 2) NOT IN ('21')"
		_cQuery += "ORDER BY DataMarcacao, Matricula, HoraMarcacao"+CRLF    	
	
	_aResultQry := U_QryArr(_cQuery)

	If Len(_aResultQry)> 0
		MsgRun("Favor Aguardar.....", "Exportando os Registros para o Excel" ,{|| DlgToExcel( { {"GETDADOS","MARCAÇÕES X ARQUIVO TXT" ,_aStrut,_aResultQry} } ) })
	Else
		MsgInfo("Não foram registros - MARCAÇÕES X ARQUIVO TXT","ATENÇÃO!")
	EndIF
	
Return
