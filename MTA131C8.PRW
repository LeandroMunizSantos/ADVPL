#include "PROTHEUS.CH"

user function MTA131C8()
   local cNumCot  := PARAMIXB[1]:OFORMMODELOWNER:ADATAMODEL[1][1][1][2]
   local cFornece := PARAMIXB[1]:ADATAMODEL[1][1][1][17]
   local cLoja    := PARAMIXB[1]:ADATAMODEL[1][1][1][18]
   local aItensPE := PARAMIXB[1]:ADATAMODEL[1][1]

   local aRetorno   := {}
   local aItensCol  := {}   // Array final com todos os itens coletados
   local aItemAtual := {}   // Array temporário de cada item
   local nMoeda     := 0
   local cMoedaNm   := ""
   local nI         := 0
   local nJ         := 0
   local nTotItens  := 0
   local nTotCampos := 0

   // Obtém a moeda configurada na cotação
   DbSelectArea("SC8")
   SC8->(DbSetOrder(1))
   SC8->(DbSeek(FWxFilial("SC8") + cNumCot))

   if !SC8->(Eof()) .and. SC8->C8_NUM == cNumCot
      nMoeda   := SC8->C8_MOEDA
      cMoedaNm := AllTrim(SuperGetMV("MV_MOEDA" + AllTrim(Str(nMoeda, 2))))
      aAdd(aRetorno, {"C8_XTPMOED", cMoedaNm})
   endif

   nTotItens := Len(aItensPE)

   for nI := 1 to nTotItens

      aItemAtual   := {}
      nTotCampos   := Len(aItensPE[nI])

      for nJ := 1 to nTotCampos
         aAdd(aItemAtual, aItensPE[nI][nJ])   // Coleta cada campo do item atual
      next nJ

      aAdd(aItensCol, aItemAtual)   // Adiciona o item completo no array final
   next nI

return aRetorno
