#Include 'Protheus.ch'

/****************************************************************************************
-----------------------------------------------------------------------------------------
-- Programa:        | M410VRES    | Autor:  Leandro Muniz          | Data:  23/04/2026  --
-----------------------------------------------------------------------------------------
-- Descrição:       | Ponto de entrada apos eliminação de residuo para usuario          --
--                  | justificar motivo da exclusão                                     --
-----------------------------------------------------------------------------------------
-- Chamada:         | MATA410, eliminação de residuo    							    --
-----------------------------------------------------------------------------------------
-- Funcionalidade:  | Criação de tela que abre automaticamente, ao usuario eliminar   	--
--                  | residuo de um pedido de venda, para justificar motivo pelo qual   --
--                  | esta sendo eliminado o residuo e povoando o motivo na tabela SC5. --
--                  |                 	                                                --
--                  |                 	                                                --
-----------------------------------------------------------------------------------------
*****************************************************************************************/

User Function M410VRES()

    Local cMotivo := Space(70)
    Local cNumPed := SC5->C5_NUM
    Local cUsuario := cUserName
    Local dHoje    := Date()

    If !PedJustificativa("Justificativa de Exclusão de Resíduo", cNumPed, @cMotivo)
        Return .F.
    EndIf

    RecLock("SC5", .F.)
        SC5->C5_XDMORES := cMotivo
        SC5->C5_USREXC  := AllTrim(cUsuario)
        SC5->C5_DTVEXC  := dHoje
    MsUnlock()

Return .T.

/*----------------------------------------------------------------------*
 | Func:  PedJustificativa                                     		    |
 | Autor: Leandro Muniz                                                 |
 | Data:  23/04/2026                                                    |
 | Desc:  Tela de input de justificativa de exclusão do residuo	        |
 | Obs.:  /                                                             |
 *----------------------------------------------------------------------*/

Static Function PedJustificativa(cTitulo, cNumRef, cMotivo)

    Local oDlg
    Local oSay
    Local oCombo
    Local lOk      := .F.
    Local lValido  := .F.
    Local cOpcao   := " "
    Local aMotivos := {}

    AAdd(aMotivos, " ")
    AAdd(aMotivos, "Falta de Estoque")
    AAdd(aMotivos, "Pedido abaixo do minimo para faturamento")

    While !lValido

        lOk    := .F.
        cOpcao := " "

        Define MsDialog oDlg Title cTitulo + " - " + cNumRef ;
            From 0,0 To 150,400 Pixel

            @ 10, 10 Say oSay Prompt "Selecione o motivo da eliminação de resíduo:" Of oDlg Pixel
            @ 25, 10 MsComboBox oCombo Var cOpcao Items aMotivos Size 180, 12 Of oDlg Pixel
            @ 50, 150 Button "Confirmar" Size 35, 12 Action (lOk := .T., oDlg:End()) Of oDlg Pixel

        Activate MsDialog oDlg Centered

        If !lOk
            Return .F.
        EndIf

        If AllTrim(cOpcao) == ""
            MsgStop("Selecione um motivo valido para prosseguir!", "Atenção")
        Else
            lValido := .T.
        EndIf

    EndDo

    cMotivo := AllTrim(cOpcao)

Return .T.
