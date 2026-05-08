#include "PROTHEUS.CH"

user function AVALCOT()
    local lRetPE  := .T.
    local cNumPC  := ""
    local cNumCot := ""
    local cDoc    := ""
    local aArea   := GetArea() 

    conout("AVALCOT - Iniciando processamento")

    if ValType(PARAMIXB) == "A" .and. Len(PARAMIXB) >= 2
        if ValType(PARAMIXB[1]) == "C"
            cNumPC := AllTrim(PARAMIXB[1])
        endif
        if ValType(PARAMIXB[2]) == "C"
            cNumCot := AllTrim(PARAMIXB[2])
        endif
    else
        conout("AVALCOT - Erro: PARAMIXB invalido ou incompleto. Abortando PE para evitar Erro 500.")
        return lRetPE
    endif

    conout("AVALCOT - cNumPC="  + cNumPC)
    conout("AVALCOT - cNumCot=" + cNumCot)

    // ---------------------------------------------------------------
    // Busca o número do contrato gravado na SC8 e propaga para SC7
    // ---------------------------------------------------------------
    if !Empty(cNumCot)
        DbSelectArea("SC8")
        SC8->(DbSetOrder(1)) // C8_FILIAL + C8_NUM
        
        if SC8->(DbSeek(FWxFilial("SC8") + cNumCot))
            if SC8->(FieldPos("C8_XCONTRA")) > 0
                cDoc := AllTrim(SC8->C8_XCONTRA)
                conout("AVALCOT - Contrato encontrado: " + cDoc)
            else
                conout("AVALCOT - Erro: Campo C8_XCONTRA nao existe no SX3.")
            endif
        endif
    endif

    RestArea(aArea) 
return lRetPE
