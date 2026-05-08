user function NFCGERDOC()
    local cOpc      := PARAMIXB[1]
    local aDados    := PARAMIXB[2]
    local lRet      := .T.
    local cMsg      := ""
    local aRet      := {}
    local aAreaSC8  := SC8->(GetArea())
    local aAreaSB1  := SB1->(GetArea())

    local cCotCod   := ""
    local cFornCod  := ""
    local cLoja     := ""
    local cTpDoc    := ""
    local aItVenc   := {}
    local aSupInfo  := {}

    // Variáveis de validação
    local nX             := 0
    local cCodProd       := ""
    local cContrat       := ""
    local cXcontrat      := ""
    local aProdveri      := {}
    local cPadraoContrat := "A"
    local cPadraoXcontr  := "S"

    // Variáveis para montar aItensCol
    local aItensCol  := {}
    local aItemAtual := {}
    local nI         := 0
    local nJ         := 0
    local aItensPE   := {}
    local nTotItens  := 0
    local nTotCampos := 0

    conout("NFCGERDOC - Iniciando processamento. Opcao: " + cValToChar(cOpc))

    // 1. LEITURA DOS PARÂMETROS
    if cOpc == "1"
        cCotCod  := aDados[1]
        cFornCod := aDados[2]
        cLoja    := aDados[3]
        cTpDoc   := aDados[5]
    elseif cOpc == "2"
        cCotCod  := aDados[1]
        aItVenc  := aDados[2]
        if Len(aItVenc) > 0
            aSupInfo := aItVenc[1]['supplierInfo']
            cFornCod := aSupInfo['suppliercode']
            cLoja    := aSupInfo['supplierstore']
            cTpDoc   := aSupInfo['documenttype']
        endif
    endif

    // 2. MONTA aItensCol direto da SC8
    aItensPE := {}
    DbSelectArea("SC8")
    SC8->(DbSetOrder(1))
    SC8->(DbSeek(FWxFilial("SC8") + cCotCod))

    while SC8->(!Eof()) .and. SC8->C8_NUM == cCotCod
        if SC8->C8_FORNECE == cFornCod .and. SC8->C8_LOJA == cLoja

            aAdd(aItensPE, {  ;
                SC8->C8_PRODUTO,  ;  // [1]
                SC8->C8_QUANT,    ;  // [2]
                SC8->C8_PRECO,    ;  // [3]
                SC8->C8_TXMOEDA,  ;  // [4]
                SC8->C8_VALIPI,   ;  // [5]
                SC8->C8_XDTINIC,   ;  // [6]
                SC8->C8_XDTVIGE,   ;  // [7]
                SC8->C8_XCC,      ;  // [8]
                SC8->C8_XOBS,     ;  // [9]
                SC8->C8_MOEDA     ;  // [10]
            })
        endif
        SC8->(DbSkip())
    enddo

    // 3. VALIDAÇÃO: BLOQUEIA SE JÁ TIVER PEDIDO VINCULADO
    DbSelectArea("SC8")
    SC8->(DbSetOrder(1))
    if SC8->(DbSeek(FWxFilial("SC8") + cCotCod))
        while SC8->(!Eof()) .and. SC8->C8_NUM == cCotCod .and. SC8->C8_FILIAL == FWxFilial("SC8")
            if !Empty(SC8->C8_NUMPED) .or. !Empty(SC8->C8_XCONTRA)
                lRet := .F.
                cMsg := "A cotacao " + cCotCod + " ja possui documento gerado (Pedido/Contrato)."
                aRet := {lRet, cMsg}
                SC8->(RestArea(aAreaSC8))
                return aRet
            endif
            SC8->(DbSkip())
        enddo
    endif

    // 4. VALIDAÇÃO DE PRODUTOS (SB1)
    if cOpc == "2" .and. Len(aItVenc) > 0
        DbSelectArea("SB1")
        SB1->(DbSetOrder(1))
        for nX := 1 to Len(aItVenc)
            cCodProd := AllTrim(aItVenc[nX]['productcode'])

            if aScan(aProdveri, {|x| x == cCodProd}) > 0
                lRet := .F.
                cMsg := "O produto " + cCodProd + " esta duplicado na analise."
                aRet := {lRet, cMsg}
                return aRet
            endif
            aAdd(aProdveri, cCodProd)

            if SB1->(DbSeek(FWxFilial("SB1") + cCodProd))
                if SB1->B1_CONTRAT != cPadraoContrat .or. SB1->B1_XCONTRA != cPadraoXcontr
                    if SB1->(RecLock("SB1", .F.))
                        SB1->B1_CONTRAT := cPadraoContrat
                        SB1->B1_XCONTRA := cPadraoXcontr
                        SB1->(MsUnlock())
                    endif
                endif
            endif
        next nX
    endif

    // 5. GERAÇÃO DO CONTRATO (MATA125)
    if lRet
        lRet := NFCEXEC125(cCotCod, cFornCod, cLoja, aItVenc, aItensCol, aItensPE)
        if !lRet
            cMsg := "Erro na geracao do Contrato de Parceria via MSExecAuto."
        endif
    endif

    aRet := {lRet, cMsg}
    SC8->(RestArea(aAreaSC8))
    SB1->(RestArea(aAreaSB1))
return aRet

/*----------------------------------------------------------------------*
 | Func:    NFCEXEC125 (Static)                                         |
 | Desc:    Executa o MATA125 via MSExecAuto                            |
 *----------------------------------------------------------------------*/
static function NFCEXEC125(cCotacao, cFornece, cLoja, aItensNFC, aItensCol, aItensPE)
    local aCab          := {}
    local aItens        := {}
    local aLinha        := {}
    local cDoc          := GetSxeNum("SC3","C3_NUM")
    local lRet          := .T.
    local nX            := 0
    local nPreco        := 0
    local nVlrTotal     := 0
    local cCodProd      := ""
    local cIdent        := ""
    local nQuant        := 0
    local aAreaSC8      := SC8->(GetArea())
    local aRecnosSC8    := {}
    local dDtIni        := CtoD("")
    local dDtFim        := CtoD("")
    private lMsErroAuto := .F.

    // 1. MONTAGEM DO CABEÇALHO (SC3)
    // Buscamos dados comuns na SC8 para o cabeçalho
    DbSelectArea("SC8")
    SC8->(DbSetOrder(1))
    SC8->(DbSeek(FWxFilial("SC8") + cCotacao))

    aadd(aCab, {"C3_FILIAL",   FWxFilial("SC3"),        nil})
    aadd(aCab, {"C3_NUM",      cDoc,                    nil})
    aadd(aCab, {"C3_EMISSAO",  dDataBase,               nil})
    aadd(aCab, {"C3_FORNECE",  cFornece,                nil})
    aadd(aCab, {"C3_LOJA",     cLoja,                   nil})
    aadd(aCab, {"C3_COND",     SC8->C8_COND,            nil})
    aadd(aCab, {"C3_MOEDA",    SC8->C8_MOEDA,           nil})
    aadd(aCab, {"C3_CONTATO",  SC8->C8_CONTATO,         nil})
    aadd(aCab, {"C3_TPFRETE",  SC8->C8_TPFRETE,         nil})

    // 2. MONTAGEM DOS ITENS (SC3)
    // Se aItensNFC estiver vazio (vencimento por proposta), buscamos todos da SC8 para aquele fornecedor
    if Len(aItensNFC) == 0
        SC8->(DbSetOrder(1)) // Filial + Num
        SC8->(DbSeek(FWxFilial("SC8") + cCotacao))
        while SC8->(!Eof()) .and. SC8->C8_NUM == cCotacao
            if SC8->C8_FORNECE == cFornece .and. SC8->C8_LOJA == cLoja
                aAdd(aRecnosSC8, SC8->(Recno()))
            endif
            SC8->(DbSkip())
        enddo
    else
        // Se veio por itens selecionados no NFC
        for nX := 1 to Len(aItensNFC)
            cCodProd := aItensNFC[nX]['productcode']
            cIdent   := aItensNFC[nX]['identifier'] // ID do item na cotação
            
            SC8->(DbSetOrder(4)) // Filial + Num + Ident + Prod
            if SC8->(DbSeek(FWxFilial("SC8") + cCotacao + cIdent + cCodProd))
                aAdd(aRecnosSC8, SC8->(Recno()))
            endif
        next nX
    endif

    // Processa os itens para montar o aItens do ExecAuto
    for nX := 1 to Len(aItensPE)
        aLinha := {}

        nQuant := aItensPE[nX][2]  // C8_QUANT

        // Cálculo do preço líquido
        if aItensPE[nX][10] == 2   // C8_MOEDA == 2
            nVlrTotal := (aItensPE[nX][3] / IIF(aItensPE[nX][4] == 0, 1, aItensPE[nX][4])) - aItensPE[nX][5]
        else
            nVlrTotal := aItensPE[nX][3] - aItensPE[nX][5]
        endif

        // Proteção contra divisão por zero
        if nQuant == 0
            nPreco := 0
        else
            nPreco := nVlrTotal / nQuant
        endif

        // Proteção contra datas vazias
        dDtIni := aItensPE[nX][6]
        dDtFim := aItensPE[nX][7]

        if Empty(dDtIni)
            dDtIni := dDataBase
        endif
        if Empty(dDtFim)
            dDtFim := dDataBase + 365
        endif

        aadd(aLinha, {"C3_ITEM",    StrZero(nX, 4),   nil})
        aadd(aLinha, {"C3_PRODUTO", aItensPE[nX][1],  nil})  // C8_PRODUTO
        aadd(aLinha, {"C3_QUANT",   nQuant,           nil})
        aadd(aLinha, {"C3_PRECO",   nPreco,           nil})
        aadd(aLinha, {"C3_TOTAL",   nVlrTotal,        nil})
        aadd(aLinha, {"C3_DATPRI",  dDtIni,           nil})  // Data inicial
        aadd(aLinha, {"C3_DATPRF",  dDtFim,           nil})  // Data final
        aadd(aLinha, {"C3_CC",      aItensPE[nX][8],  nil})  // C8_XCC
        aadd(aLinha, {"C3_OBS",     aItensPE[nX][9],  nil})  // C8_XOBS
        aadd(aItens, aLinha)
    next nX

    // 3. EXECUÇÃO DO MSEXECAUTO
    MSExecAuto({|x,y,z| MATA125(x,y,z)}, aCab, aItens, 3)

    if lMsErroAuto
        lRet := .F.
        MostraErro()
    else
        // 4. ATUALIZAÇÃO DOS VÍNCULOS (SC8 e SC1) EM CASO DE SUCESSO
        for nX := 1 to Len(aRecnosSC8)
            SC8->(DbGoTo(aRecnosSC8[nX]))
            if SC8->(RecLock("SC8", .F.))
                SC8->C8_XCONTRA := cDoc
                SC8->C8_NUMPED  := "CONTRAT" // Marca que virou contrato
                SC8->(MsUnlock())
            endif

            // Atualiza Solicitação de Compra (SC1)
            DbSelectArea("SC1")
            SC1->(DbSetOrder(1)) // C1_FILIAL+C1_NUM+C1_ITEM
            if SC1->(DbSeek(SC8->C8_FILIAL + SC8->C8_NUMSC + SC8->C8_ITEMSC))
                if SC1->(RecLock("SC1", .F.))
                    SC1->C1_QUJE    := SC1->C1_QUANT
                    SC1->C1_XCONTRA := cDoc
                    SC1->(MsUnlock())
                endif
            endif
        next nX
        MsgInfo("Contrato " + cDoc + " gerado com sucesso!")
    endif

    SC8->(RestArea(aAreaSC8))
return lRet
