#Include "TOTVS.ch"
#Include 'Protheus.ch'
#INCLUDE "rwmake.ch"

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³PREVREL   ºAutor  Leandro Muniz         º Data ³  13/03/26  º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³Relatório de comparativo de horario previsto x realizado	  º±±
±±º          ³				                                              º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ DEPARTAMENTO PESSOAL                                       º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

User Function PREVREL()
	
	Local _cQuery		:= ""
	Local _aResultQry	:= {}
	Local _aStrut		:= {}
	local _cPerg 		:= "PREVREL"

	Pergunte(_cPerg,.t.)

    Local cList := AllTrim(MV_PAR07)
    cList := StrTran(cList, ";", ",")
    cList := StrTran(cList, " ", "")
    If Right(cList,1) == ","
        cList := Left(cList, Len(cList)-1)
    EndIf
	
	aAdd(_aStrut,{"Filal"		     ,"C",02,0})
	aAdd(_aStrut,{"Matricula"        ,"C",06,0})
	aAdd(_aStrut,{"Nome"             ,"C",60,0})
	aAdd(_aStrut,{"CC"               ,"C",09,0})
	aAdd(_aStrut,{"Centro de Custo"	 ,"C",60,0})
    aAdd(_aStrut,{"Data"             ,"C",10,0})
    aAdd(_aStrut,{"Evento"           ,"C",03,0})
    aAdd(_aStrut,{"Desc. Evento"     ,"C",60,0})
    aAdd(_aStrut,{"Cod Inf"          ,"C",03,0})
    aAdd(_aStrut,{"Cod Abono"        ,"C",03,0})
    aAdd(_aStrut,{"Motivo Abono"     ,"C",60,0})
    aAdd(_aStrut,{"Inc. Intervalo"   ,"C",20,0})
    aAdd(_aStrut,{"Hora Prevista"    ,"C",30,0})
    aAdd(_aStrut,{"Hora Realizada"   ,"C",30,0})
    aAdd(_aStrut,{"Total de Horas"   ,"C",06,0})
	
    	_cQuery += "SELECT"+CRLF
    	_cQuery += "        RA.RA_FILIAL AS 'FILIAL',"+CRLF
    	_cQuery += "        RA.RA_MAT AS 'MATRICULA',"+CRLF
    	_cQuery += "        RA.RA_NOME AS 'NOME',"+CRLF
    	_cQuery += "        RA.RA_CC AS 'CC',"+CRLF
    	_cQuery += "        (SELECT TOP 1 CTT_DESC01 FROM CTT010 WHERE CTT_CUSTO = RA.RA_CC) AS 'CENTRO_CUSTO',"+CRLF
    	_cQuery += "        CONVERT(char(10), CONVERT(date, D.P8_DATAAPO), 103) AS 'DATA',"+CRLF
    	_cQuery += "        PC.PC_PD     AS 'TIPO_EVENTO',"+CRLF
    	_cQuery += "		(SELECT P9_DESC FROM SP9010 WHERE P9_CODIGO = PC.PC_PD AND D_E_L_E_T_ = '') AS 'DESC_EVENTO',"+CRLF
    	_cQuery += "        PC.PC_PDI,"+CRLF
    	_cQuery += "        PC.PC_ABONO,"+CRLF
    	_cQuery += "		(SELECT P6_DESC FROM SP6010 WHERE P6_CODIGO = PC.PC_ABONO AND D_E_L_E_T_ = '') AS 'DESC_ABONO',"+CRLF
    	_cQuery += "		(CASE "+CRLF
    	_cQuery += "			WHEN OA_P8.HORAS_REAL IS NOT NULL "+CRLF
    	_cQuery += "					AND OA_P8.HORAS_REAL <> ''"+CRLF
    	_cQuery += "			THEN"+CRLF
    	_cQuery += "				CASE "+CRLF
    	_cQuery += "					WHEN "+CRLF
    	_cQuery += "						PARSENAME(REPLACE(OA_R6.HORAS_PREV, ' ', '.'), 3) <>"+CRLF
    	_cQuery += "						PARSENAME(REPLACE(OA_P8.HORAS_REAL, ' ', '.'), 3)"+CRLF
    	_cQuery += "						OR "+CRLF
    	_cQuery += "						PARSENAME(REPLACE(OA_R6.HORAS_PREV, ' ', '.'), 2) <>"+CRLF
    	_cQuery += "						PARSENAME(REPLACE(OA_P8.HORAS_REAL, ' ', '.'), 2)"+CRLF
    	_cQuery += "					THEN 'Inconsistente'"+CRLF
    	_cQuery += "					ELSE ''"+CRLF
    	_cQuery += "				END"+CRLF
    	_cQuery += "			ELSE ''"+CRLF
    	_cQuery += "		END) AS 'Inc. Intervalo',"+CRLF
    	_cQuery += "        OA_R6.HORAS_PREV AS 'HORARIO_PREVISTO',"+CRLF
    	_cQuery += "        OA_P8.HORAS_REAL AS 'HORARIO_REALIZADO',"+CRLF
    	_cQuery += "        (QF.HH + ':' + QF.MM) AS 'TOTAL_HORAS',"+CRLF
    	_cQuery += "        '' AS Del"+CRLF
    	_cQuery += "    FROM SRA010 AS RA"+CRLF
    	_cQuery += "    INNER JOIN ("+CRLF
    	_cQuery += "        SELECT DISTINCT P8_MAT, P8_DATAAPO"+CRLF
    	_cQuery += "        FROM SP8010"+CRLF
    	_cQuery += "        WHERE D_E_L_E_T_ = ' '"+CRLF
    	_cQuery += "        	AND P8_TPMARCA IN ('1E','1S','2E','2S')"+CRLF
    	_cQuery += "		UNION"+CRLF
    	_cQuery += "			SELECT DISTINCT"+CRLF
    	_cQuery += "				PC_MAT  AS P8_MAT,"+CRLF
    	_cQuery += "				PC_DATA AS P8_DATAAPO"+CRLF
    	_cQuery += "			FROM SPC010"+CRLF
    	_cQuery += "			WHERE D_E_L_E_T_ = ' '"+CRLF
    	_cQuery += "				AND PC_PD = '008'"+CRLF
    	_cQuery += "    ) AS D"+CRLF
    	_cQuery += "        ON D.P8_MAT = RA.RA_MAT"+CRLF
    	_cQuery += "        "+CRLF
    	_cQuery += "	OUTER APPLY ("+CRLF
    	_cQuery += "		SELECT "+CRLF
    	_cQuery += "			STRING_AGG(F.HH + ':' + F.MM, ' ') "+CRLF
    	_cQuery += "				WITHIN GROUP (ORDER BY R.ord) AS HORAS_PREV"+CRLF
    	_cQuery += "		FROM SPJ010 AS PJ"+CRLF
    	_cQuery += "		CROSS APPLY ("+CRLF
    	_cQuery += "			SELECT ((DATEPART(WEEKDAY, D.P8_DATAAPO) + @@DATEFIRST - 1) % 7) + 1 AS dow"+CRLF
    	_cQuery += "		) AS DW"+CRLF
    	_cQuery += "		CROSS APPLY ("+CRLF
    	_cQuery += "			SELECT 1 AS ord, CAST(PJ.PJ_ENTRA1 AS varchar(20)) AS raw UNION ALL"+CRLF
    	_cQuery += "			SELECT 2,      CAST(PJ.PJ_SAIDA1 AS varchar(20))            UNION ALL"+CRLF
    	_cQuery += "			SELECT 3,      CAST(PJ.PJ_ENTRA2 AS varchar(20))            UNION ALL"+CRLF
    	_cQuery += "			SELECT 4,      CAST(PJ.PJ_SAIDA2 AS varchar(20))"+CRLF
    	_cQuery += "		) AS R"+CRLF
    	_cQuery += "		CROSS APPLY ("+CRLF
    	_cQuery += "			SELECT LTRIM(RTRIM(R.raw)) AS txt,"+CRLF
    	_cQuery += "				CASE "+CRLF
    	_cQuery += "					WHEN CHARINDEX(':', LTRIM(RTRIM(R.raw))) > 0 THEN CHARINDEX(':', LTRIM(RTRIM(R.raw)))"+CRLF
    	_cQuery += "					WHEN CHARINDEX(',', LTRIM(RTRIM(R.raw))) > 0 THEN CHARINDEX(',', LTRIM(RTRIM(R.raw)))"+CRLF
    	_cQuery += "					WHEN CHARINDEX('.', LTRIM(RTRIM(R.raw))) > 0 THEN CHARINDEX('.', LTRIM(RTRIM(R.raw)))"+CRLF
    	_cQuery += "					ELSE 0"+CRLF
    	_cQuery += "				END AS pos"+CRLF
    	_cQuery += "		) AS X"+CRLF
    	_cQuery += "		CROSS APPLY ("+CRLF
    	_cQuery += "			SELECT"+CRLF
    	_cQuery += "				RIGHT('00' + "+CRLF
    	_cQuery += "					NULLIF("+CRLF
    	_cQuery += "						CASE "+CRLF
    	_cQuery += "							WHEN X.txt IS NULL OR X.txt = '' THEN NULL"+CRLF
    	_cQuery += "							WHEN X.pos > 0 THEN LEFT(X.txt, X.pos - 1)"+CRLF
    	_cQuery += "							WHEN LEN(X.txt) >= 3 THEN LEFT(X.txt, LEN(X.txt) - 2)"+CRLF
    	_cQuery += "							ELSE X.txt"+CRLF
    	_cQuery += "						END"+CRLF
    	_cQuery += "					, ''), 2) AS HH,"+CRLF
    	_cQuery += "				RIGHT('00' + "+CRLF
    	_cQuery += "					CASE "+CRLF
    	_cQuery += "						WHEN X.txt IS NULL OR X.txt = '' THEN '00'"+CRLF
    	_cQuery += "						WHEN X.pos > 0 THEN"+CRLF
    	_cQuery += "							CASE "+CRLF
    	_cQuery += "								WHEN LEN(SUBSTRING(X.txt, X.pos + 1, 2)) = 0 THEN '00'"+CRLF
    	_cQuery += "								WHEN LEN(SUBSTRING(X.txt, X.pos + 1, 2)) = 1 THEN SUBSTRING(X.txt, X.pos + 1, 1) + '0'"+CRLF
    	_cQuery += "								ELSE LEFT(SUBSTRING(X.txt, X.pos + 1, 2), 2)"+CRLF
    	_cQuery += "							END"+CRLF
    	_cQuery += "						WHEN LEN(X.txt) >= 3 THEN RIGHT(X.txt, 2)"+CRLF
    	_cQuery += "						ELSE '00'"+CRLF
    	_cQuery += "					END"+CRLF
    	_cQuery += "				, 2) AS MM"+CRLF
    	_cQuery += "		) AS F"+CRLF
    	_cQuery += "		WHERE RA.RA_DEMISSA = '' "+CRLF
    	_cQuery += "			AND PJ.PJ_TURNO = (SELECT top 1 PF_TURNOPA FROM SPF010 PF WHERE D.P8_MAT = PF.PF_MAT AND PF.D_E_L_E_T_ = '' AND PF.PF_FILIAL = RA.RA_FILIAL AND D.P8_DATAAPO >= PF.PF_DATA ORDER BY PF.PF_DATA DESC)"+CRLF
    	_cQuery += "			AND PJ.PJ_DIA   = DW.dow"+CRLF
    	_cQuery += "			AND F.HH IS NOT NULL AND F.MM IS NOT NULL"+CRLF
    	_cQuery += "			AND TRY_CONVERT(int, F.HH) BETWEEN 0 AND 23"+CRLF
    	_cQuery += "			AND TRY_CONVERT(int, F.MM) BETWEEN 0 AND 59"+CRLF
    	_cQuery += "			AND PJ.D_E_L_E_T_ = ''"+CRLF
    	_cQuery += "	) AS OA_R6"+CRLF
    	_cQuery += "    OUTER APPLY ("+CRLF
    	_cQuery += "        SELECT "+CRLF
    	_cQuery += "        	STRING_AGG(HH + ':' + MM, ' ') "+CRLF
    	_cQuery += "        		WITHIN GROUP (ORDER BY OrdTP) AS HORAS_REAL"+CRLF
    	_cQuery += "        FROM ("+CRLF
    	_cQuery += "        	SELECT "+CRLF
    	_cQuery += "        		RIGHT('0' + CASE WHEN pos > 0 THEN LEFT(raw, pos - 1) ELSE raw END, 2) AS HH,"+CRLF
    	_cQuery += "        		CASE "+CRLF
    	_cQuery += "        			WHEN pos = 0 THEN '00'"+CRLF
    	_cQuery += "        			ELSE"+CRLF
    	_cQuery += "        				CASE "+CRLF
    	_cQuery += "        					WHEN LEN(SUBSTRING(raw, pos + 1, 2)) = 0 THEN '00'"+CRLF
    	_cQuery += "        					WHEN LEN(SUBSTRING(raw, pos + 1, 2)) = 1 THEN SUBSTRING(raw, pos + 1, 1) + '0'"+CRLF
    	_cQuery += "        					ELSE LEFT(SUBSTRING(raw, pos + 1, 2), 2)"+CRLF
    	_cQuery += "        				END"+CRLF
    	_cQuery += "        		END AS MM,"+CRLF
    	_cQuery += "        		OrdTP"+CRLF
    	_cQuery += "        	FROM ("+CRLF
    	_cQuery += "        		SELECT "+CRLF
    	_cQuery += "        			LTRIM(RTRIM(CAST(P8X.P8_HORA AS varchar(20)))) AS raw,"+CRLF
    	_cQuery += "        			CASE "+CRLF
    	_cQuery += "        				WHEN CHARINDEX(',', LTRIM(RTRIM(CAST(P8X.P8_HORA AS varchar(20))))) > 0"+CRLF
    	_cQuery += "        					THEN CHARINDEX(',', LTRIM(RTRIM(CAST(P8X.P8_HORA AS varchar(20)))))"+CRLF
    	_cQuery += "        				WHEN CHARINDEX('.', LTRIM(RTRIM(CAST(P8X.P8_HORA AS varchar(20))))) > 0"+CRLF
    	_cQuery += "        					THEN CHARINDEX('.', LTRIM(RTRIM(CAST(P8X.P8_HORA AS varchar(20)))))"+CRLF
    	_cQuery += "        				ELSE 0"+CRLF
    	_cQuery += "        			END AS pos,"+CRLF
    	_cQuery += "        			CASE P8X.P8_TPMARCA "+CRLF
    	_cQuery += "        				WHEN '1E' THEN 1 "+CRLF
    	_cQuery += "        				WHEN '1S' THEN 2 "+CRLF
    	_cQuery += "        				WHEN '2E' THEN 3 "+CRLF
    	_cQuery += "        				WHEN '2S' THEN 4 "+CRLF
    	_cQuery += "        				ELSE 99 "+CRLF
    	_cQuery += "        			END AS OrdTP,"+CRLF
    	_cQuery += "        			ROW_NUMBER() OVER ("+CRLF
    	_cQuery += "        				PARTITION BY P8X.P8_TPMARCA "+CRLF
    	_cQuery += "        				ORDER BY P8X.P8_DHORG, P8X.R_E_C_N_O_"+CRLF
    	_cQuery += "        			) AS rn"+CRLF
    	_cQuery += "        		FROM SP8010 AS P8X"+CRLF
    	_cQuery += "        		WHERE P8X.D_E_L_E_T_ = ' '"+CRLF
    	_cQuery += "        		AND P8X.P8_MAT     = D.P8_MAT"+CRLF
    	_cQuery += "        		AND P8X.P8_DATAAPO    = D.P8_DATAAPO"+CRLF
    	_cQuery += "        		AND P8X.P8_TPMARCA IN ('1E','1S','2E','2S')"+CRLF
    	_cQuery += "				AND (P8X.P8_APONTA = 'S' OR P8X.P8_FLAG = 'G')"+CRLF
    	_cQuery += "        	) AS S"+CRLF
    	_cQuery += "        	WHERE S.rn = 1"+CRLF
    	_cQuery += "        ) AS Z"+CRLF
    	_cQuery += "    ) AS OA_P8"+CRLF
    	_cQuery += "    INNER JOIN SPC010 AS PC"+CRLF
    	_cQuery += "        	ON PC.PC_MAT  = D.P8_MAT"+CRLF
    	_cQuery += "        AND PC.PC_DATA = D.P8_DATAAPO"+CRLF
    	_cQuery += "    CROSS APPLY ("+CRLF
    	_cQuery += "        SELECT "+CRLF
    	_cQuery += "        	LTRIM(RTRIM(CAST(PC.PC_QUANTC AS varchar(20)))) AS raw,"+CRLF
    	_cQuery += "        	CHARINDEX(',', LTRIM(RTRIM(CAST(PC.PC_QUANTC AS varchar(20))))) AS posC,"+CRLF
    	_cQuery += "        	CHARINDEX('.', LTRIM(RTRIM(CAST(PC.PC_QUANTC AS varchar(20))))) AS posD"+CRLF
    	_cQuery += "    ) AS Q0"+CRLF
    	_cQuery += "    CROSS APPLY ("+CRLF
    	_cQuery += "        SELECT "+CRLF
    	_cQuery += "        	CASE WHEN Q0.posC > 0 THEN Q0.posC"+CRLF
    	_cQuery += "        		WHEN Q0.posD > 0 THEN Q0.posD"+CRLF
    	_cQuery += "        		ELSE 0"+CRLF
    	_cQuery += "        	END AS pos,"+CRLF
    	_cQuery += "        	Q0.raw"+CRLF
    	_cQuery += "    ) AS Q1"+CRLF
    	_cQuery += "    CROSS APPLY ("+CRLF
    	_cQuery += "        SELECT"+CRLF
    	_cQuery += "        	RIGHT('0' + NULLIF(CASE WHEN Q1.pos > 0 THEN LEFT(Q1.raw, Q1.pos - 1) ELSE Q1.raw END, ''), 2) AS HH,"+CRLF
    	_cQuery += "        	CASE "+CRLF
    	_cQuery += "        		WHEN Q1.pos = 0 THEN '00'"+CRLF
    	_cQuery += "        		ELSE"+CRLF
    	_cQuery += "        			CASE "+CRLF
    	_cQuery += "        				WHEN LEN(SUBSTRING(Q1.raw, Q1.pos + 1, 2)) = 0 THEN '00'"+CRLF
    	_cQuery += "        				WHEN LEN(SUBSTRING(Q1.raw, Q1.pos + 1, 2)) = 1 THEN SUBSTRING(Q1.raw, Q1.pos + 1, 1) + '0'"+CRLF
    	_cQuery += "        				ELSE LEFT(SUBSTRING(Q1.raw, Q1.pos + 1, 2), 2)"+CRLF
    	_cQuery += "        			END"+CRLF
    	_cQuery += "        	END AS MM"+CRLF
    	_cQuery += "    ) AS QF"+CRLF
        _cQuery += "WHERE RA.RA_DEMISSA = ''"+CRLF
        _cQuery += "AND RA.RA_FILIAL BETWEEN '" + MV_PAR01 + "' AND '" + MV_PAR02 + "'"+CRLF
        _cQuery += "AND D.P8_MAT BETWEEN '" + MV_PAR03 + "' AND '" + MV_PAR04 + "'"+CRLF
        _cQuery += "AND RA.RA_CC BETWEEN '" + MV_PAR05 + "' AND '" + MV_PAR06 + "'"+CRLF
        _cQuery += "AND D.P8_DATAAPO BETWEEN '" + dtos(MV_PAR08) + "' AND '" + dtos(MV_PAR09) + "'"+CRLF
        _cQuery += "AND PC.PC_PD IN (" + cList + ")"+CRLF
        _cQuery += "AND PC.D_E_L_E_T_ = '' "+CRLF
        _cQuery += "AND RA.D_E_L_E_T_ = '' "+CRLF
        _cQuery += "ORDER BY D.P8_DATAAPO, FILIAL, MATRICULA"+CRLF	
	
	_aResultQry := U_QryArr(_cQuery)

	If Len(_aResultQry)> 0
		MsgRun("Favor Aguardar.....", "Exportando os Registros para o Excel" ,{|| DlgToExcel( { {"GETDADOS","HORARIO PREVISTO X REALIZADO" ,_aStrut,_aResultQry} } ) })
	Else
		MsgInfo("Não foram encontrados registros - HORARIO PREVISTO X REALIZADO","ATENÇÃO!")
	EndIF
	
Return
