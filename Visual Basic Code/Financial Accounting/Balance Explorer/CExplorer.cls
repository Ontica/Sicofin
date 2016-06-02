VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*** Empiria® ***********************************************************************************************
'*                                                                                                          *
'* Solución   : Empiria® Software Components                    Sistema : Financial Accounting              *
'* Componente : Balance Explorer (EFABalanceExplorer)           Parte   : User services                     *
'* Clase      : CExplorer                                       Patrón  : Stateless Class                   *
'* Fecha      : 28/Febrero/2002                                 Versión : 2.0       Versión patrón: 1.0     *
'*                                                                                                          *
'* Descripción: Clase principal del componente "Balance Explorer".                                          *
'*                                                                                                          *
'****************************************************** Copyright © La Vía Ontica, S.C. México, 1999-2002. **
Option Explicit
Private Const ClassId As String = "CExplorer"

Private Const cnMaxNumberOfColumns = 10

Private Const cnHeaderCol1 = _
        "<TR class=fullScrollMenuHeader valign=center>" & _
        "<TD colspan=6 class=fullScrollMenuTitle>Resultado de la consulta</TD>" & _
        "<TD colspan=<@COLSPAN@> align=right>" & _
        "<A href='' onclick='return(showOptionsWindow(document.all.divSelectedBalancesOptions));'>Operaciones con los saldos obtenidos</A>" & _
        "<img align=absbottom src='/empiria/images/invisible.gif'>" & _
        "</TD></TR>" & _
        "<TR class=applicationTableHeader valign=center>" & _
        "<TD nowrap><img src='/empiria/images/collapsed.gif' onclick='outline();' alt='Contraer todo'></TD>" & _
        "<TD nowrap><A href='' onclick='return orderBy(1);'>Número de cuenta</A><@MARK_1@></TD> " & _
        "<TD><A href='' onclick='return orderBy(2);'>Nombre de la cuenta</A><@MARK_2@></TD>" & _
        "<TD nowrap><A href='' onclick='return orderBy(3);'>Sect</A><@MARK_3@></TD>" & _
        "<TD nowrap><A href='' onclick='return orderBy(4);'>Mon</A><@MARK_4@></TD>"
Private Const cnHeaderCol2 = _
        "<TD nowrap align=right><A href='' onclick='return orderBy(6);'>Saldo inicial</A><@MARK_5@></TD>" & _
        "<TD nowrap align=right><A href='' onclick='return orderBy(7);'>Cargos</A><@MARK_6@></TD>" & _
        "<TD nowrap align=right><A href='' onclick='return orderBy(8);'>Abonos</A><@MARK_7@></TD>"
Private Const cnHeaderCol3 = _
        "<TD nowrap align=right><A href='' onclick='return orderBy(9);'>Saldo actual</A><@MARK_8@></TD>"
Private Const cnHeaderCol4 = _
        "<TD nowrap align=right><A href='' onclick='return orderBy(10);'>Saldo promedio</A><@MARK_9@></TD> "

Private Const cnCollapsedDivLine = _
          "<TBODY value=1><TR class=applicationTableRowDivisor valign=center>" & _
          "<TD nowrap colspan=<@COLSPAN@>><@SEPARATOR@>" & _
          "<img align=absbottom src='/empiria/images/collapsed.gif' onclick='outliner();'>" & _
          "<@NAME@></TD></TR>"
 
Private Const cnExpandedDivLine = _
          "<TBODY value=1><TR class=applicationTableRowDivisor>" & _
          "<TD nowrap colspan=<@COLSPAN@>><@SEPARATOR@>" & _
          "<img align=absbottom src='/empiria/images/expanded.gif' onclick='outliner();'>" & _
          "<@NAME@></TD></TR>" & vbCrLf
                    
Private Const cnCurrencySeparator = _
          "<TD nowrap><@CURRENCY@></TD>"
          
Private Const cnItemLine0 = _
          "<TR <@CLASS@> valign=center style='display:<@DISPLAY@>;'>" & _
          "<TD nowrap><@SEPARATOR@></TD>" & _
          "<TD nowrap><@PIVOT_ACCOUNT_NUMBER@></TD>" & _
          "<TD><@PIVOT_ACCOUNT_NAME@></TD>" & _
          "<TD nowrap><@SECTOR@></TD>" & _
          "<TD nowrap><@CURRENCY@></TD>"
       
Private Const cnItemLine2 = _
          "<TD nowrap align=right><@INIT_BALANCE@></A></TD>" & _
          "<TD nowrap align=right><@DEBIT@></A></TD>" & _
          "<TD nowrap align=right><@CREDIT@></A></TD>"
          
Private Const cnItemLine3 = _
          "<TD nowrap align=right><A href='' onclick='return(showIntegration(<@PARS@>));'><@BALANCE@></A></TD>"
          
Private Const cnItemLine4 = _
          "<TD nowrap align=right>" & _
          "<A href='' onclick='return(showAverage(<@PARS@>));'><@AVERAGE_BALANCE@></A>" & _
          "</TD>" & vbCrLf

'************************************************************************************************************
'* PROPIEDADES PÚBLICAS                                                                                     *
'************************************************************************************************************

'************************************************************************************************************
'* MÉTODOS PÚBLICOS                                                                                         *
'************************************************************************************************************

Public Function BalanceVouchersBody(sAppServer As String, vGralLedgers As Variant, bConsolidateGralLedgers As Boolean, _
                         sAccountPattern As String, dFromDate As Date, dToDate As Date, _
                         Optional sFromAccount As String, Optional sToAccount As String, _
                     Optional sFromSubLedgerAccount As String, Optional sToSubLedgerAccount As String, _
                     Optional nTransactionTypeId As Long, Optional nVoucherTypeId As Long, _
                     Optional nCurrencyTypeId As Long, Optional nAccountTypeId As Long, _
                     Optional sAccountNature As String, Optional nExchangeRateTypeId As Long, _
                     Optional nExchangeRateCurrencyId As Long, Optional dExchangeRateDate As Date, _
                     Optional nBalanceTypeToDisplay As Long, Optional bShowSubsAccounts As Boolean, _
                     Optional bGroupBySubLedgerAccount As Boolean, _
                     Optional bShowAverageCol As Boolean, Optional bShowOnlyLastLevels As Boolean, _
                     Optional sWhere As String = "", _
                     Optional nOrderColumn As Long, Optional sOrderBy As String = "") As String
  Dim oBalances As New EFABalanceEngine.CEngine, oBalanceRS As Recordset
  Dim sTemp As String, sTemp2 As String, sItemsSection As String, sHTML As String, sPattern As String
  Dim sItemSeparator As String, nColSpan As Long
  Dim nCurrentGralLedgerGroupId As Long, nCurrentGralLedgerId As Long
  Dim sGralLedgerNumber As String, sGralLedgerName As String
  Dim nCurrentCurrencyId As Long, sDisplay As String
  Dim nInitialBalanceSum As Double, nDebitSum As Double, nCreditSum As Double
  Dim nBalanceSum As Double, nAvgBalanceSum As Double
  Dim sSubLedgerAccountNumber As String, sSubLedgerAccountName As String, nLastSubledgerAccount As Long
  Dim nLastStdAccountId As Long, nLastCurrencyId As Long, nLastSectorId As Long, bChanges As Boolean
  '*******************************************************************************************************
  On Error GoTo ErrHandler
    sItemSeparator = Separator(3)
    sPattern = ItemsRowPattern(bShowSubsAccounts, bShowAverageCol)
    nColSpan = HeaderColumnsCount(bShowSubsAccounts, bShowAverageCol)
    
    If (Not bShowAverageCol) Then
      If bGroupBySubLedgerAccount Then
        bShowOnlyLastLevels = True
      End If
      Set oBalanceRS = oBalances.BalancesRS(sAppServer, sAccountPattern, vGralLedgers, _
                                      bConsolidateGralLedgers, bShowOnlyLastLevels, bShowSubsAccounts, _
                                      dFromDate, dToDate, sFromAccount, sToAccount, _
                                      nTransactionTypeId, nVoucherTypeId, nCurrencyTypeId, _
                                      nAccountTypeId, sAccountNature, nExchangeRateTypeId, _
                                      nExchangeRateCurrencyId, dExchangeRateDate, nBalanceTypeToDisplay, _
                                      sFromSubLedgerAccount, sToSubLedgerAccount, , , bGroupBySubLedgerAccount)
    Else
      Set oBalanceRS = oBalances.BalancesRS(sAppServer, sAccountPattern, vGralLedgers, _
                                      bConsolidateGralLedgers, bShowOnlyLastLevels, bShowSubsAccounts, _
                                      dFromDate, dToDate, sFromAccount, sToAccount, _
                                      nTransactionTypeId, nVoucherTypeId, nCurrencyTypeId, _
                                      nAccountTypeId, sAccountNature, nExchangeRateTypeId, _
                                      nExchangeRateCurrencyId, dExchangeRateDate, nBalanceTypeToDisplay, _
                                      sFromSubLedgerAccount, sToSubLedgerAccount, , , bGroupBySubLedgerAccount)
    End If
    With oBalanceRS
      Do
        If .EOF Then
          Exit Do
        End If
        nCurrentGralLedgerId = !id_mayor
        If (nCurrentGralLedgerId <> 0) Then
          GetGralLedgerValues sAppServer, nCurrentGralLedgerId, sGralLedgerNumber, sGralLedgerName
          sGralLedgerNumber = "[" & sGralLedgerNumber & "] "
        ElseIf Not IsArray(vGralLedgers) Then
          GetGralLedgerValues sAppServer, CLng(vGralLedgers), sGralLedgerNumber, sGralLedgerName
          sGralLedgerNumber = "[" & sGralLedgerNumber & "] "
        Else
          sGralLedgerName = "<b>Saldos consolidados</b>"
        End If
        If Not IsArray(vGralLedgers) Then
          sItemsSection = DivisionLine(nCurrentGralLedgerId, _
                                       "<b>" & sGralLedgerNumber & sGralLedgerName & "</b>", 2, nColSpan, False)
          sDisplay = "inline"
        ElseIf (IsArray(vGralLedgers)) And (bConsolidateGralLedgers) Then
          sItemsSection = DivisionLine(nCurrentGralLedgerId, _
                                       "<b>" & sGralLedgerNumber & sGralLedgerName & "</b>", 2, nColSpan, False)
          sDisplay = "inline"
        ElseIf (IsArray(vGralLedgers)) And (Not bConsolidateGralLedgers) Then
          sItemsSection = DivisionLine(nCurrentGralLedgerId, _
                                       "<b>" & sGralLedgerNumber & sGralLedgerName & "</b>", 2, nColSpan, True)
          sDisplay = "none"
        End If
        nCurrentCurrencyId = 0
        Do
          If .EOF Then
            Exit Do
          End If
          If (nCurrentGralLedgerGroupId <> 0) Or (nCurrentGralLedgerId <> !id_mayor) Then
            Exit Do
          End If
          If nCurrentCurrencyId = !id_moneda Then
            sTemp = sPattern
          Else
            sTemp = "<TR class=applicationTableHeader style='display:<@DISPLAY@>;'><TD></TD>" & _
                    "<TD colspan=" & nColSpan - 1 & ">" & !nombre_moneda & " (" & !clave_moneda & ")</TD></TR>"
            sTemp = Replace(sTemp, "<@DISPLAY@>", sDisplay)
            sItemsSection = sItemsSection & sTemp
            sTemp = sPattern
            nCurrentCurrencyId = !id_moneda
          End If
          
          sTemp = Replace(sTemp, "<@SEPARATOR@>", sItemSeparator)
          sTemp = Replace(sTemp, "<@DISPLAY@>", sDisplay)
          sTemp = Replace(sTemp, "<@SECTOR@>", !clave_sector)
          sTemp = Replace(sTemp, "<@CURRENCY@>", !clave_moneda)
          sTemp = Replace(sTemp, "<@PARS@>", !id_mayor & "," & !id_moneda & "," & !id_cuenta_estandar & "," & !id_cuenta_auxiliar & "," & !id_sector)
          If (Not bShowSubsAccounts) Then
            sTemp = Replace(sTemp, "<@CLASS@>", _
                                   IIf(Len(!numero_cuenta_estandar) = 4 And (!clave_sector = "00"), _
                                   "class=applicationTableSelectedRow", ""))
            sTemp = Replace(sTemp, "<@PIVOT_ACCOUNT_NUMBER@>", !numero_cuenta_estandar)
            sTemp = Replace(sTemp, "<@PIVOT_ACCOUNT_NAME@>", !nombre_cuenta)
            sTemp = Replace(sTemp, "<@INIT_BALANCE@>", Format$(!saldo_inicial, "Currency"))
            sTemp = Replace(sTemp, "<@DEBIT@>", Format$(!cargos, "Currency"))
            sTemp = Replace(sTemp, "<@CREDIT@>", Format$(!abonos, "Currency"))
            sTemp = Replace(sTemp, "<@BALANCE@>", Format$(!saldo_no_valorizado, "Currency"))
            If bShowAverageCol Then
              sTemp = Replace(sTemp, "<@AVERAGE_BALANCE@>", "SProm")
            End If
          ElseIf bShowSubsAccounts Then
            If (Not bGroupBySubLedgerAccount) Then
              bChanges = (!id_cuenta_estandar <> nLastStdAccountId) Or (!id_moneda <> nLastCurrencyId) Or _
                         (!id_sector <> nLastSectorId)
            Else
              bChanges = (!id_cuenta_auxiliar <> nLastSubledgerAccount) Or (!id_moneda <> nLastCurrencyId)
            End If
            If (bChanges) Then
              sItemsSection = Replace(sItemsSection, "<@INIT_BALANCE@>", Format$(nInitialBalanceSum, "Currency"), , 1)
              sItemsSection = Replace(sItemsSection, "<@DEBIT@>", Format$(nDebitSum, "Currency"), , 1)
              sItemsSection = Replace(sItemsSection, "<@CREDIT@>", Format$(nCreditSum, "Currency"), , 1)
              sItemsSection = Replace(sItemsSection, "<@BALANCE@>", Format$(nBalanceSum, "Currency"), , 1)
              sItemsSection = Replace(sItemsSection, "<@AVERAGE_BALANCE@>", Format$(nAvgBalanceSum, "Currency"), , 1)
              nInitialBalanceSum = 0: nDebitSum = 0: nCreditSum = 0: nBalanceSum = 0: nAvgBalanceSum = 0
              If (!id_cuenta_auxiliar <> 0) And (Not bGroupBySubLedgerAccount) Then
                sTemp2 = sTemp     'Agrega el renglón para la cuenta madre del auxiliar
                sTemp2 = Replace(sTemp2, "<@CLASS@>", "class=applicationTableSelectedRow")
                sTemp2 = Replace(sTemp2, "<@PIVOT_ACCOUNT_NUMBER@>", !numero_cuenta_estandar)
                sTemp2 = Replace(sTemp2, "<@PIVOT_ACCOUNT_NAME@>", !nombre_cuenta)
                sItemsSection = sItemsSection & sTemp2
              ElseIf (!id_cuenta_auxiliar <> 0) And (bGroupBySubLedgerAccount) Then
                sTemp2 = sTemp     'Agrega el renglón para la cuenta madre del auxiliar
                sTemp2 = Replace(sTemp2, "<@CLASS@>", "class=applicationTableSelectedRow")
                sTemp2 = Replace(sTemp2, "<@PIVOT_ACCOUNT_NUMBER@>", !numero_cuenta_auxiliar)
                sTemp2 = Replace(sTemp2, "<@PIVOT_ACCOUNT_NAME@>", !nombre_cuenta_auxiliar)
                sItemsSection = sItemsSection & sTemp2
              End If
              nLastStdAccountId = !id_cuenta_estandar
              nLastCurrencyId = !id_moneda
              nLastSectorId = !id_sector
              nLastSubledgerAccount = !id_cuenta_auxiliar
            End If
            sTemp = Replace(sTemp, "<@CLASS@>", _
                                   IIf(!id_cuenta_auxiliar = 0, _
                                   "class=applicationTableRowDivisor", ""))
            If (Not bGroupBySubLedgerAccount) Then
              sTemp = Replace(sTemp, "<@PIVOT_ACCOUNT_NUMBER@>", IIf(IsNull(!numero_cuenta_auxiliar), !numero_cuenta_estandar, !numero_cuenta_auxiliar))
              sTemp = Replace(sTemp, "<@PIVOT_ACCOUNT_NAME@>", IIf(IsNull(!numero_cuenta_auxiliar), !nombre_cuenta, "<i>" & !nombre_cuenta_auxiliar & "</i>"))
            ElseIf bGroupBySubLedgerAccount Then
              sTemp = Replace(sTemp, "<@PIVOT_ACCOUNT_NUMBER@>", !numero_cuenta_estandar)
              sTemp = Replace(sTemp, "<@PIVOT_ACCOUNT_NAME@>", "<i>" & !nombre_cuenta & "</i>")
            End If
            sTemp = Replace(sTemp, "<@INIT_BALANCE@>", Format$(!saldo_inicial, "Currency"))
            sTemp = Replace(sTemp, "<@DEBIT@>", Format$(!cargos, "Currency"))
            sTemp = Replace(sTemp, "<@CREDIT@>", Format$(!abonos, "Currency"))
            sTemp = Replace(sTemp, "<@BALANCE@>", Format$(!saldo_no_valorizado, "Currency"))
            If bShowAverageCol Then
              sTemp = Replace(sTemp, "<@AVERAGE_BALANCE@>", "SProm")
              nAvgBalanceSum = nAvgBalanceSum + !saldo_promedio
            End If
            nInitialBalanceSum = nInitialBalanceSum + !saldo_inicial
            nDebitSum = nDebitSum + !cargos
            nCreditSum = nCreditSum + !abonos
            nBalanceSum = nBalanceSum + !saldo_no_valorizado
          End If
          sItemsSection = sItemsSection & sTemp
          .MoveNext
        Loop
        If bShowSubsAccounts Then
          sItemsSection = Replace(sItemsSection, "<@INIT_BALANCE@>", Format$(nInitialBalanceSum, "Currency"), , 1)
          sItemsSection = Replace(sItemsSection, "<@DEBIT@>", Format$(nDebitSum, "Currency"), , 1)
          sItemsSection = Replace(sItemsSection, "<@CREDIT@>", Format$(nCreditSum, "Currency"), , 1)
          sItemsSection = Replace(sItemsSection, "<@BALANCE@>", Format$(nBalanceSum, "Currency"), , 1)
          sItemsSection = Replace(sItemsSection, "<@AVERAGE_BALANCE@>", Format$(nAvgBalanceSum, "Currency"), , 1)
          nInitialBalanceSum = 0: nDebitSum = 0: nCreditSum = 0: nBalanceSum = 0: nAvgBalanceSum = 0
        End If
        sHTML = sHTML & sItemsSection & "</TBODY>"
      Loop
    End With
    BalanceVouchersBody = sHTML
  Exit Function
ErrHandler:
  RaiseError ClassId, "BalanceVouchersBody", Err.Number
End Function

Public Function Body(sAppServer As String, vGralLedgers As Variant, bConsolidateGralLedgers As Boolean, _
                     sAccountPattern As String, dFromDate As Date, dToDate As Date, _
                     Optional sFromAccount As String, Optional sToAccount As String, _
                     Optional sFromSubLedgerAccount As String, Optional sToSubLedgerAccount As String, _
                     Optional nTransactionTypeId As Long, Optional nVoucherTypeId As Long, _
                     Optional nCurrencyTypeId As Long, Optional nAccountTypeId As Long, _
                     Optional sAccountNature As String, Optional nExchangeRateTypeId As Long, _
                     Optional nExchangeRateCurrencyId As Long, Optional dExchangeRateDate As Date, _
                     Optional nBalanceTypeToDisplay As Long, Optional bShowSubsAccounts As Boolean, _
                     Optional bGroupBySubLedgerAccount As Boolean, _
                     Optional bShowAverageCol As Boolean, Optional bShowOnlyLastLevels As Boolean, _
                     Optional sWhere As String = "", _
                     Optional nOrderColumn As Long, Optional sOrderBy As String = "") As String
  Dim oBalances As New EFABalanceEngine.CEngine, oBalanceRS As Recordset
  Dim sTemp As String, sTemp2 As String, sItemsSection As String, sHTML As String, sPattern As String
  Dim sItemSeparator As String, nColSpan As Long
  Dim nCurrentGralLedgerGroupId As Long, nCurrentGralLedgerId As Long
  Dim sGralLedgerNumber As String, sGralLedgerName As String
  Dim nCurrentCurrencyId As Long, sDisplay As String
  Dim nInitialBalanceSum As Double, nDebitSum As Double, nCreditSum As Double
  Dim nBalanceSum As Double, nAvgBalanceSum As Double
  Dim sSubLedgerAccountNumber As String, sSubLedgerAccountName As String, nLastSubledgerAccount As Long
  Dim nLastStdAccountId As Long, nLastCurrencyId As Long, nLastSectorId As Long, bChanges As Boolean
  '*******************************************************************************************************
  On Error GoTo ErrHandler
    sItemSeparator = Separator(3)
    sPattern = ItemsRowPattern(bShowSubsAccounts, bShowAverageCol)
    nColSpan = HeaderColumnsCount(bShowSubsAccounts, bShowAverageCol)
    
    If (Not bShowAverageCol) Then
      If bGroupBySubLedgerAccount Then
        bShowOnlyLastLevels = True
      End If
      Set oBalanceRS = oBalances.BalancesRS(sAppServer, sAccountPattern, vGralLedgers, _
                                      bConsolidateGralLedgers, bShowOnlyLastLevels, bShowSubsAccounts, _
                                      dFromDate, dToDate, sFromAccount, sToAccount, _
                                      nTransactionTypeId, nVoucherTypeId, nCurrencyTypeId, _
                                      nAccountTypeId, sAccountNature, nExchangeRateTypeId, _
                                      nExchangeRateCurrencyId, dExchangeRateDate, nBalanceTypeToDisplay, _
                                      sFromSubLedgerAccount, sToSubLedgerAccount, , , bGroupBySubLedgerAccount)
    Else
      Set oBalanceRS = oBalances.BalancesRS(sAppServer, sAccountPattern, vGralLedgers, _
                                      bConsolidateGralLedgers, bShowOnlyLastLevels, bShowSubsAccounts, _
                                      dFromDate, dToDate, sFromAccount, sToAccount, _
                                      nTransactionTypeId, nVoucherTypeId, nCurrencyTypeId, _
                                      nAccountTypeId, sAccountNature, nExchangeRateTypeId, _
                                      nExchangeRateCurrencyId, dExchangeRateDate, nBalanceTypeToDisplay, _
                                      sFromSubLedgerAccount, sToSubLedgerAccount, , , bGroupBySubLedgerAccount)
    End If
    With oBalanceRS
      Do
        If .EOF Then
          Exit Do
        End If
        nCurrentGralLedgerId = !id_mayor
        If (nCurrentGralLedgerId <> 0) Then
          GetGralLedgerValues sAppServer, nCurrentGralLedgerId, sGralLedgerNumber, sGralLedgerName
          sGralLedgerNumber = "[" & sGralLedgerNumber & "] "
        ElseIf Not IsArray(vGralLedgers) Then
          GetGralLedgerValues sAppServer, CLng(vGralLedgers), sGralLedgerNumber, sGralLedgerName
          sGralLedgerNumber = "[" & sGralLedgerNumber & "] "
        Else
          sGralLedgerName = "<b>Saldos consolidados</b>"
        End If
        If Not IsArray(vGralLedgers) Then
          sItemsSection = DivisionLine(nCurrentGralLedgerId, _
                                       "<b>" & sGralLedgerNumber & sGralLedgerName & "</b>", 2, nColSpan, False)
          sDisplay = "inline"
        ElseIf (IsArray(vGralLedgers)) And (bConsolidateGralLedgers) Then
          sItemsSection = DivisionLine(nCurrentGralLedgerId, _
                                       "<b>" & sGralLedgerNumber & sGralLedgerName & "</b>", 2, nColSpan, False)
          sDisplay = "inline"
        ElseIf (IsArray(vGralLedgers)) And (Not bConsolidateGralLedgers) Then
          sItemsSection = DivisionLine(nCurrentGralLedgerId, _
                                       "<b>" & sGralLedgerNumber & sGralLedgerName & "</b>", 2, nColSpan, True)
          sDisplay = "none"
        End If
        nCurrentCurrencyId = 0
        Do
          If .EOF Then
            Exit Do
          End If
          If (nCurrentGralLedgerGroupId <> 0) Or (nCurrentGralLedgerId <> !id_mayor) Then
            Exit Do
          End If
          If nCurrentCurrencyId = !id_moneda Then
            sTemp = sPattern
          Else
            sTemp = "<TR class=applicationTableHeader style='display:<@DISPLAY@>;'><TD></TD>" & _
                    "<TD colspan=" & nColSpan - 1 & ">" & !nombre_moneda & " (" & !clave_moneda & ")</TD></TR>"
            sTemp = Replace(sTemp, "<@DISPLAY@>", sDisplay)
            sItemsSection = sItemsSection & sTemp
            sTemp = sPattern
            nCurrentCurrencyId = !id_moneda
          End If
          
          sTemp = Replace(sTemp, "<@SEPARATOR@>", sItemSeparator)
          sTemp = Replace(sTemp, "<@DISPLAY@>", sDisplay)
          sTemp = Replace(sTemp, "<@SECTOR@>", !clave_sector)
          sTemp = Replace(sTemp, "<@CURRENCY@>", !clave_moneda)
          sTemp = Replace(sTemp, "<@PARS@>", !id_mayor & "," & !id_moneda & "," & !id_cuenta_estandar & "," & !id_cuenta_auxiliar & "," & !id_sector)
          If (Not bShowSubsAccounts) Then
            sTemp = Replace(sTemp, "<@CLASS@>", _
                                   IIf(Len(!numero_cuenta_estandar) = 4 And (!clave_sector = "00"), _
                                   "class=applicationTableSelectedRow", ""))
            sTemp = Replace(sTemp, "<@PIVOT_ACCOUNT_NUMBER@>", !numero_cuenta_estandar)
            sTemp = Replace(sTemp, "<@PIVOT_ACCOUNT_NAME@>", !nombre_cuenta)
            sTemp = Replace(sTemp, "<@INIT_BALANCE@>", Format$(!saldo_inicial, "Currency"))
            sTemp = Replace(sTemp, "<@DEBIT@>", Format$(!cargos, "Currency"))
            sTemp = Replace(sTemp, "<@CREDIT@>", Format$(!abonos, "Currency"))
            sTemp = Replace(sTemp, "<@BALANCE@>", Format$(!saldo_no_valorizado, "Currency"))
            If bShowAverageCol Then
              sTemp = Replace(sTemp, "<@AVERAGE_BALANCE@>", "SProm")
            End If
          ElseIf bShowSubsAccounts Then
            If (Not bGroupBySubLedgerAccount) Then
              bChanges = (!id_cuenta_estandar <> nLastStdAccountId) Or (!id_moneda <> nLastCurrencyId) Or _
                         (!id_sector <> nLastSectorId)
            Else
              bChanges = (!id_cuenta_auxiliar <> nLastSubledgerAccount) Or (!id_moneda <> nLastCurrencyId)
            End If
            If (bChanges) Then
              sItemsSection = Replace(sItemsSection, "<@INIT_BALANCE@>", Format$(nInitialBalanceSum, "Currency"), , 1)
              sItemsSection = Replace(sItemsSection, "<@DEBIT@>", Format$(nDebitSum, "Currency"), , 1)
              sItemsSection = Replace(sItemsSection, "<@CREDIT@>", Format$(nCreditSum, "Currency"), , 1)
              sItemsSection = Replace(sItemsSection, "<@BALANCE@>", Format$(nBalanceSum, "Currency"), , 1)
              sItemsSection = Replace(sItemsSection, "<@AVERAGE_BALANCE@>", Format$(nAvgBalanceSum, "Currency"), , 1)
              nInitialBalanceSum = 0: nDebitSum = 0: nCreditSum = 0: nBalanceSum = 0: nAvgBalanceSum = 0
              If (!id_cuenta_auxiliar <> 0) And (Not bGroupBySubLedgerAccount) Then
                sTemp2 = sTemp     'Agrega el renglón para la cuenta madre del auxiliar
                sTemp2 = Replace(sTemp2, "<@CLASS@>", "class=applicationTableSelectedRow")
                sTemp2 = Replace(sTemp2, "<@PIVOT_ACCOUNT_NUMBER@>", !numero_cuenta_estandar)
                sTemp2 = Replace(sTemp2, "<@PIVOT_ACCOUNT_NAME@>", !nombre_cuenta)
                sItemsSection = sItemsSection & sTemp2
              ElseIf (!id_cuenta_auxiliar <> 0) And (bGroupBySubLedgerAccount) Then
                sTemp2 = sTemp     'Agrega el renglón para la cuenta madre del auxiliar
                sTemp2 = Replace(sTemp2, "<@CLASS@>", "class=applicationTableSelectedRow")
                sTemp2 = Replace(sTemp2, "<@PIVOT_ACCOUNT_NUMBER@>", !numero_cuenta_auxiliar)
                sTemp2 = Replace(sTemp2, "<@PIVOT_ACCOUNT_NAME@>", !nombre_cuenta_auxiliar)
                sItemsSection = sItemsSection & sTemp2
              End If
              nLastStdAccountId = !id_cuenta_estandar
              nLastCurrencyId = !id_moneda
              nLastSectorId = !id_sector
              nLastSubledgerAccount = !id_cuenta_auxiliar
            End If
            sTemp = Replace(sTemp, "<@CLASS@>", _
                                   IIf(!id_cuenta_auxiliar = 0, _
                                   "class=applicationTableRowDivisor", ""))
            If (Not bGroupBySubLedgerAccount) Then
              sTemp = Replace(sTemp, "<@PIVOT_ACCOUNT_NUMBER@>", IIf(IsNull(!numero_cuenta_auxiliar), !numero_cuenta_estandar, !numero_cuenta_auxiliar))
              sTemp = Replace(sTemp, "<@PIVOT_ACCOUNT_NAME@>", IIf(IsNull(!numero_cuenta_auxiliar), !nombre_cuenta, "<i>" & !nombre_cuenta_auxiliar & "</i>"))
            ElseIf bGroupBySubLedgerAccount Then
              sTemp = Replace(sTemp, "<@PIVOT_ACCOUNT_NUMBER@>", !numero_cuenta_estandar)
              sTemp = Replace(sTemp, "<@PIVOT_ACCOUNT_NAME@>", "<i>" & !nombre_cuenta & "</i>")
            End If
            sTemp = Replace(sTemp, "<@INIT_BALANCE@>", Format$(!saldo_inicial, "Currency"))
            sTemp = Replace(sTemp, "<@DEBIT@>", Format$(!cargos, "Currency"))
            sTemp = Replace(sTemp, "<@CREDIT@>", Format$(!abonos, "Currency"))
            sTemp = Replace(sTemp, "<@BALANCE@>", Format$(!saldo_no_valorizado, "Currency"))
            If bShowAverageCol Then
              sTemp = Replace(sTemp, "<@AVERAGE_BALANCE@>", "SProm")
              nAvgBalanceSum = nAvgBalanceSum + !saldo_promedio
            End If
            nInitialBalanceSum = nInitialBalanceSum + !saldo_inicial
            nDebitSum = nDebitSum + !cargos
            nCreditSum = nCreditSum + !abonos
            nBalanceSum = nBalanceSum + !saldo_no_valorizado
          End If
          sItemsSection = sItemsSection & sTemp
          .MoveNext
        Loop
        If bShowSubsAccounts Then
          sItemsSection = Replace(sItemsSection, "<@INIT_BALANCE@>", Format$(nInitialBalanceSum, "Currency"), , 1)
          sItemsSection = Replace(sItemsSection, "<@DEBIT@>", Format$(nDebitSum, "Currency"), , 1)
          sItemsSection = Replace(sItemsSection, "<@CREDIT@>", Format$(nCreditSum, "Currency"), , 1)
          sItemsSection = Replace(sItemsSection, "<@BALANCE@>", Format$(nBalanceSum, "Currency"), , 1)
          sItemsSection = Replace(sItemsSection, "<@AVERAGE_BALANCE@>", Format$(nAvgBalanceSum, "Currency"), , 1)
          nInitialBalanceSum = 0: nDebitSum = 0: nCreditSum = 0: nBalanceSum = 0: nAvgBalanceSum = 0
        End If
        sHTML = sHTML & sItemsSection & "</TBODY>"
      Loop
    End With
    Body = sHTML
  Exit Function
ErrHandler:
  RaiseError ClassId, "Body", Err.Number
End Function

Public Function BuildSearchParametersString(Optional nGralLedgerGroup As Long, _
                                            Optional dFromApplicationDate As String, _
                                            Optional dToApplicationDate As String, _
                                            Optional dFromElaborationDate As String, _
                                            Optional dToElaborationDate As String, _
                                            Optional sVoucherNumber As String, _
                                            Optional sVoucherConcept As String, _
                                            Optional sAccounts As String, _
                                            Optional nTransactionTypeId As Long, _
                                            Optional nVoucherTypeId As Long, _
                                            Optional nBalancingType As Long) As String
  Dim sResult As String, sTemp As String
  '************************************************************************************
  On Error GoTo ErrHandler
    If (nGralLedgerGroup <> 0) Then
      sResult = ""
    End If
    
    sTemp = BuildDateIntervalString("fecha_afectacion", dFromApplicationDate, dToApplicationDate, False)
    If (Len(sTemp) <> 0) Then
      sResult = sResult & IIf(Len(sResult) <> 0, " AND ", "") & sTemp
    End If
    
    sTemp = BuildDateIntervalString("fecha_registro", dFromElaborationDate, dToElaborationDate, False)
    If (Len(sTemp) <> 0) Then
      sResult = sResult & IIf(Len(sResult) <> 0, " AND ", "") & sTemp
    End If
    
    sTemp = BuildLikeString("numero_transaccion", sVoucherNumber)
    If (Len(sTemp) <> 0) Then
      sResult = sResult & IIf(Len(sResult) <> 0, " AND ", "") & sTemp
    End If
    
    sTemp = BuildLikeString("concepto_transaccion", sVoucherConcept)
    If (Len(sTemp) <> 0) Then
      sResult = sResult & IIf(Len(sResult) <> 0, " AND ", "") & sTemp
    End If
    
    sTemp = BuildInString("numero_cuenta_estandar", sAccounts)
    If (Len(sTemp) <> 0) Then
      sResult = sResult & IIf(Len(sResult) <> 0, " AND ", "") & sTemp
    End If
    
    sTemp = BuildIdCompareString("id_tipo_transaccion", nTransactionTypeId)
    If (Len(sTemp) <> 0) Then
      sResult = sResult & IIf(Len(sResult) <> 0, " AND ", "") & sTemp
    End If
    
    sTemp = BuildIdCompareString("id_tipo_poliza", nVoucherTypeId)
    If (Len(sTemp) <> 0) Then
      sResult = sResult & IIf(Len(sResult) <> 0, " AND ", "") & sTemp
    End If
    
    If (nBalancingType = 1) Then
      sTemp = "(suma_cargos <> suma_abonos)"
      sResult = sResult & IIf(Len(sResult) <> 0, " AND ", "") & sTemp
    ElseIf (nBalancingType = 2) Then
      sTemp = "(suma_cargos = suma_abonos)"
      sResult = sResult & IIf(Len(sResult) <> 0, " AND ", "") & sTemp
    End If
    BuildSearchParametersString = sResult
  Exit Function
ErrHandler:
  RaiseError ClassId, "BuildSearchParametersString", Err.Number
End Function

Public Function Header(Optional bShowSubsAccounts As Boolean = False, _
                       Optional bGroupBySubLedgerAccount As Boolean = False, _
                       Optional bShowAverageCol As Boolean = False, _
                       Optional nSelectedColumn As Long) As String
  Dim sTemp As String, nColSpan As Long, i As Long
  '***********************************************************************
  On Error GoTo ErrHandler
    If bGroupBySubLedgerAccount Then
      sTemp = cnHeaderCol1
      nColSpan = 4
    ElseIf (Not bGroupBySubLedgerAccount) Then
      sTemp = cnHeaderCol1
      nColSpan = 3
    End If
    sTemp = sTemp & cnHeaderCol2
    sTemp = sTemp & cnHeaderCol3
    If bShowAverageCol Then
      sTemp = sTemp & cnHeaderCol4
      nColSpan = nColSpan + 1
    End If
    sTemp = sTemp & "</TR>"
    sTemp = Replace(sTemp, "<@COLSPAN@>", nColSpan)
    
    If (nSelectedColumn > 0) Then
      sTemp = Replace(sTemp, "<@MARK_" & nSelectedColumn & "@>", " <img src='/empiria/images/order_asc_red.gif'>")
    ElseIf (nSelectedColumn < 0) Then
      sTemp = Replace(sTemp, "<@MARK_" & Abs(nSelectedColumn) & "@>", " <img src='/empiria/images/order_desc_red.gif'>")
    End If
    For i = 1 To cnMaxNumberOfColumns
      sTemp = Replace(sTemp, "<@MARK_" & i & "@>", "")
    Next i
    Header = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "Header", Err.Number
End Function


'************************************************************************************************************
'* PROPIEDADES PRIVADAS                                                                                     *
'************************************************************************************************************

Private Property Get ColumnOrderBy(nSelectedColumn As Long) As String
  Dim sTemp As String
  '******************************************************************
  On Error GoTo ErrHandler
    Select Case Abs(nSelectedColumn)
      Case 1
        sTemp = "fecha_afectacion,"
      Case 2
        sTemp = "fecha_registro,"
      Case 3
        sTemp = "numero_transaccion,"
      Case 4
        sTemp = "tipo_transaccion, fecha_afectacion"
      Case 5
        sTemp = "tipo_poliza, fecha_afectacion"
      Case 6
        sTemp = "concepto_transaccion, fecha_afectacion"
      Case 7
        sTemp = "cargos, fecha_afectacion"
      Case 8
        sTemp = "abonos, fecha_afectacion"
      Case 9
        sTemp = "estado"
      Case Else
        sTemp = "fecha_afectacion"
    End Select
    If (nSelectedColumn < 0) Then
      sTemp = Replace(sTemp, ",", " DESC,")
    End If
    If Right$(sTemp, 1) = "," Then
      sTemp = Left$(sTemp, Len(sTemp) - 1)
    End If
    ColumnOrderBy = sTemp
  Exit Property
ErrHandler:
  RaiseError ClassId, "ColumnOrderBy", Err.Number
End Property

'************************************************************************************************************
'* MÉTODOS PRIVADOS                                                                                         *
'************************************************************************************************************

Private Function BuildDateIntervalString(sFieldName As String, dFromDate As String, dToDate As String, _
                                         Optional bIncludeTime As Boolean = False) As String
  On Error GoTo ErrHandler
    If (Not bIncludeTime) Then
      If ((Len(dFromDate) = 0) And (Len(dToDate) = 0)) Then
        BuildDateIntervalString = ""
      ElseIf ((Len(dFromDate) <> 0) And (Len(dToDate) = 0)) Then
        BuildDateIntervalString = " (" & sFieldName & " = trunc(to_date('" & dFromDate & "')) ) "
      ElseIf ((Len(dFromDate) = 0) And (Len(dToDate) <> 0)) Then
        BuildDateIntervalString = " (" & sFieldName & " = trunc(to_date('" & dToDate & "')) ) "
      ElseIf (Len(dFromDate) <> 0) And (Len(dToDate) <> 0) Then
        BuildDateIntervalString = " ( (" & sFieldName & " >= trunc(to_date('" & dFromDate & "')) ) AND " & _
                                  " (" & sFieldName & " <= trunc(to_date('" & dToDate & "')) ) ) "
      End If
    Else
      If ((Len(dFromDate) = 0) And (Len(dToDate) = 0)) Then
        BuildDateIntervalString = ""
      ElseIf ((Len(dFromDate) <> 0) And (Len(dToDate) = 0)) Then
        BuildDateIntervalString = _
          " (" & sFieldName & " = to_date(" & dFromDate & ", 'DD/MM/YYYY HH24:MI:SS') ) "
      ElseIf ((Len(dFromDate) = 0) And (Len(dToDate) <> 0)) Then
        BuildDateIntervalString = _
          " (" & sFieldName & " = to_date(" & dToDate & ", 'DD/MM/YYYY HH24:MI:SS') ) "
      ElseIf (Len(dFromDate) <> 0) And (Len(dToDate) <> 0) Then
        BuildDateIntervalString = _
          " ( (" & sFieldName & " >= to_date(" & dFromDate & ", 'DD/MM/YYYY HH24:MI:SS') ) AND " & _
          " (" & sFieldName & " <= to_date(" & dToDate & ", 'DD/MM/YYYY HH24:MI:SS') ) ) "
      End If
    End If
  Exit Function
ErrHandler:
  RaiseError ClassId, "BuildDateIntervalString", Err.Number
End Function

Private Function BuildIdCompareString(sFieldName As String, nIdNumber As Long) As String
  On Error GoTo ErrHandler
    If (nIdNumber = 0) Then
      BuildIdCompareString = ""
    ElseIf (nIdNumber > 0) Then
      BuildIdCompareString = " (" & sFieldName & " = " & nIdNumber & ") "
    ElseIf (nIdNumber < 0) Then
      BuildIdCompareString = " (" & sFieldName & " <> " & nIdNumber & ") "
    End If
  Exit Function
ErrHandler:
  RaiseError ClassId, "BuildIdCompareString", Err.Number
End Function

Private Function BuildInString(sFieldName As String, sValuesArray As String) As String
  On Error GoTo ErrHandler
    If (Len(sValuesArray) = 0) Then
      BuildInString = ""
      Exit Function
    End If
    BuildInString = " (" & sFieldName & " IN (" & sValuesArray & ") "
  Exit Function
ErrHandler:
  RaiseError ClassId, "BuildInString", Err.Number
End Function

Private Function BuildLikeString(sFieldName As String, sValue As String) As String
  Dim sTemp As String
  On Error GoTo ErrHandler
    If Len(sValue) = 0 Then
      BuildLikeString = ""
      Exit Function
    End If
    If InStr(1, sValue, "*") Or InStr(1, sValue, "?") Then
      sTemp = Replace(sValue, "*", "%")
      sTemp = Replace(sTemp, "?", "_")
      BuildLikeString = "(" & sFieldName & " LIKE '" & sTemp & "')"
    Else
      BuildLikeString = "(" & sFieldName & " = '" & sValue & "')"
    End If
  Exit Function
ErrHandler:
  RaiseError ClassId, "BuildLikeString", Err.Number
End Function

Private Function DivisionLine(nObjectId As Long, sObjectName As String, _
                              nLevel As Long, nColSpan As Long, bExpanded As Boolean) As String
  Dim sTemp As String
  '**************************************************************************
  On Error GoTo ErrHandler
    sTemp = IIf(bExpanded, cnExpandedDivLine, cnCollapsedDivLine)
    sTemp = Replace(sTemp, "<@LEVEL@>", nLevel)
    sTemp = Replace(sTemp, "<@SEPARATOR@>", Separator(nLevel))
    sTemp = Replace(sTemp, "<@NAME@>", sObjectName)
    sTemp = Replace(sTemp, "<@COLSPAN@>", nColSpan)
    DivisionLine = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "DivisionLine", Err.Number
End Function

Private Sub GetGralLedgerValues(sAppServer As String, nGralLedgerId As Long, _
                                ByRef sGralLedgerNumber As String, ByRef sGralLedgerName As String)
  Dim oRecordset As New Recordset
  Dim sSQL As String, sTemp As String
  '**************************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT nombre_mayor, numero_mayor " & _
           "FROM COF_Mayor " & _
           "WHERE (id_mayor = " & nGralLedgerId & ")"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      If Not (.BOF And .EOF) Then
        sGralLedgerNumber = !numero_mayor
        sGralLedgerName = !nombre_mayor
      Else
        sGralLedgerNumber = ""
        sGralLedgerName = ""
      End If
      .Close
    End With
  Exit Sub
ErrHandler:
  RaiseError ClassId, "GetGralLedgerValues", Err.Number
End Sub

Private Function HeaderColumnsCount(bShowSubsAccounts As Boolean, bShowAverageCol As Boolean) As Long
  Dim nCounter As Long
  '***************************************************************************************************
  On Error GoTo ErrHandler
    nCounter = 9
    If bShowSubsAccounts Then
      nCounter = nCounter + 1
    End If
    If bShowAverageCol Then
      nCounter = nCounter + 1
    End If
    HeaderColumnsCount = nCounter
  Exit Function
ErrHandler:
  RaiseError ClassId, "HeaderColumnsCount", Err.Number
End Function

Private Function ItemsRowPattern(bShowSubsAccounts As Boolean, bShowAverageCol As Boolean) As String
  Dim sTemp As String
  '*************************************************************************************************
  On Error GoTo ErrHandler
    sTemp = cnItemLine0
    sTemp = sTemp & cnItemLine2
    sTemp = sTemp & cnItemLine3
    If bShowAverageCol Then
      sTemp = sTemp & cnItemLine4
    End If
    ItemsRowPattern = sTemp & "</TR>" & vbCrLf
  Exit Function
ErrHandler:
  RaiseError ClassId, "ItemsRowPattern", Err.Number
End Function

Private Function Separator(nLevel As Long) As String
  Dim i As Long, sTemp As String
  '*************************************************
  On Error GoTo ErrHandler
    If nLevel <= 1 Then
      Separator = ""
      Exit Function
    End If
    sTemp = " &nbsp; "
    For i = 3 To nLevel
      sTemp = sTemp & "&nbsp;"
    Next i
    Separator = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "Separator", Err.Number
End Function
