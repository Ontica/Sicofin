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
'* Componente : Transactions Explorer (EFATransactionsExplorer) Parte   : User services                     *
'* Clase      : CExplorer                                       Patrón  : Stateless Class                   *
'* Fecha      : 31/Enero/2002                                   Versión : 1.0       Versión patrón: 1.0     *
'*                                                                                                          *
'* Descripción: Provee los servicios de interfaz con el usuario del explorador de transacciones (pólizas).  *
'*                                                                                                          *
'****************************************************** Copyright © La Vía Ontica, S.C. México, 1999-2002. **
Option Explicit
Private Const ClassId As String = "CExplorer"

Private Const cnHeaderCol0 = _
        "<TR class=fullScrollMenuHeader valign=center>" & _
        "<TD colspan=6 class=fullScrollMenuTitle>Bandeja de pólizas</TD>" & _
        "<TD colspan=<@COLSPAN@> align=right>" & _
        "<A href='' onclick='return(showOptionsWindow(document.all.divSelectedVouchersOptions));'>Pólizas seleccionadas</A>" & _
        "<img align=absbottom src='/empiria/images/invisible.gif'>" & _
        "</TD></TR>" & _
        "<TR class=applicationTableHeader valign=center>" & _
        "<TD nowrap><img src='../../images/collapsed.gif' onclick='outline();' alt='Contraer todo'></TD>" & _
        "<TD nowrap><A href='' onclick='return orderBy(1);'>Afectación</A><@MARK_1@></TD> " & _
        "<TD nowrap><A href='' onclick='return orderBy(2);'>Elaboración</A><@MARK_2@></TD> "
Private Const cnHeaderCol1 = _
        "<TD nowrap><A href='' onclick='return orderBy(3);'>Número</A><@MARK_3@></TD> "
Private Const cnHeaderCol2 = _
        "<TD nowrap><A href='' onclick='return orderBy(4);'>Tipo de transacción</A><@MARK_4@></TD> "
Private Const cnHeaderCol3 = _
        "<TD nowrap><A href='' onclick='return orderBy(5);'>Tipo de póliza</A><@MARK_5@></TD> "
Private Const cnHeaderCol4 = _
        "<TD nowrap><A href='' onclick='return orderBy(6);'>Concepto</A><@MARK_6@></TD> "
Private Const cnHeaderCol5 = _
        "<TD nowrap><A href='' onclick='return orderBy(7);'>Cargos</A><@MARK_7@></TD>" & _
        "<TD nowrap><A href='' onclick='return orderBy(8);'>Abonos</A><@MARK_8@></TD>"
Private Const cnHeaderCol6 = _
        "<TD nowrap><A href='' onclick='return orderBy(9);'>Elaboró</A><@MARK_9@></TD>" & _
        "<TD nowrap><A href='' onclick='return orderBy(10);'>Revisó</A><@MARK_10@></TD> "
Private Const cnHeaderCol7 = _
        "<TD nowrap><A href='' onclick='return orderBy(11);'>Estado</A><@MARK_11@></TD></TR>"

Private Const cnCollapsedDivLine = _
          "<TBODY value=<@LEVEL@>><TR class=applicationTableRowDivisor valign=center>" & _
          "<TD nowrap colspan=6><@SEPARATOR@>" & _
          "<img align=absbottom src='../../images/collapsed.gif' onclick='outliner();'>" & _
          "<@NAME@> (<@COUNTER@>)</TD><TD nowrap colspan=<@COL_SPAN@> align=right>" & _
          "<A href='' onclick='return(selectItems(this));'>Seleccionar</A> | " & _
          "<A href='' onclick='return(unSelectItems(this));'>Deseleccionar</A></TD></TR>"
 
Private Const cnExpandedDivLine = _
          "<TBODY value=<@LEVEL@>><TR class=applicationTableRowDivisor>" & _
          "<TD nowrap colspan=6 valign=top><@SEPARATOR@>" & _
          "<img align=absbottom src='../../images/expanded.gif' onclick='outliner();'>" & _
          "<@NAME@> (<@COUNTER@>)</TD><TD nowrap colspan=<@COL_SPAN@> align=right>" & _
          "<A href='' onclick='return(selectItems(this));'>Seleccionar</A> | " & _
          "<A href='' onclick='return(unSelectItems(this));'>Deseleccionar</A></TD></TR>" & vbCrLf
                    
Private Const cnItemLine0 = _
          "<TR valign=center style='display:inline;'>" & _
          "<TD nowrap><@SEPARATOR@>" & _
          "<INPUT type=checkbox value=<@ID2@> name=chkAllItems onclick='selectRow();'></TD>" & _
          "<TD nowrap><@APPLICATION_DATE@></TD>" & _
          "<TD nowrap><@ELABORATION_DATE@></TD>"

Private Const cnItemLine1 = _
          "<TD nowrap><@VOUCHER_NUMBER@></TD>"

Private Const cnItemLine2 = _
          "<TD><@TRANSACTION_TYPE@></A></TD>" & _
          "<TD><@VOUCHER_TYPE@></A></TD>" & _
          "<TD width=70%><A href='<@URL@>?id=<@ID@>' <@TARGET@>><@CONCEPT@></A></TD>"
          
Private Const cnItemLine3 = _
          "<TD nowrap align=right><@DEBIT_SUM@></TD>" & _
          "<TD nowrap align=right><@CREDIT_SUM@></TD>"
          
Private Const cnItemLine4 = _
          "<TD nowrap align=right><@ELABORATED_BY@></TD>" & _
          "<TD nowrap align=right><@CHECKED_BY@></TD>"

Private Const cnItemLine5 = _
          "<TD width=10%><@STATUS@></TD></TR>" & vbCrLf

Private Const cnMaxNumberOfColumns = 11
Private Const cnMaxNumberOfInboxes = 4

'************************************************************************************************************
'* PROPIEDADES PÚBLICAS                                                                                     *
'************************************************************************************************************

Public Property Get VoucherInboxName(tVoucherInbox As TEnumVoucherInboxes) As String
  On Error GoTo ErrHandler
    Select Case tVoucherInbox
      Case efaMyPendingVouchers
        VoucherInboxName = "Mis pólizas pendientes"
      Case efaMyPostedVouchers
        VoucherInboxName = "Mis pólizas actualizadas"
      Case efaGroupPendingVouchers
        VoucherInboxName = "Las pendientes en mis grupos de trabajo"
      Case efaGroupPostedVouchers
        VoucherInboxName = "Las actualizadas en mis grupos de trabajo"
      Case Else
        VoucherInboxName = "¡Cuidado! Bandeja no definida."
    End Select
  Exit Property
ErrHandler:
  RaiseError ClassId, "VoucherInboxName", Err.Number
End Property

'************************************************************************************************************
'* MÉTODOS PÚBLICOS                                                                                         *
'************************************************************************************************************

Public Function Body(sAppServer As String, nVoucherInbox As TEnumVoucherInboxes, nUserId As Long, _
                     Optional bIncludePostings As Boolean, Optional sWhere As String, _
                     Optional nOrderColumn As Long, Optional sOrderBy As String) As String
  Dim oTransactionsRS As Recordset
  Dim sTemp As String, sItemsSection As String, sHTML As String, sPattern As String
  Dim nGroupItems As Long, nItems As Long, sItemSeparator As String, nColSpan As Long
  Dim nCurrentGralLedgerGroupId As Long, nCurrentGralLedgerId As Long
  '*************************************************************************************************
  On Error GoTo ErrHandler
    sItemSeparator = Separator(3)
    nColSpan = 2
    If bIncludePostings Then
      nColSpan = nColSpan + 2
    End If
    If nVoucherInbox = efaGroupPostedVouchers Then
      nColSpan = nColSpan + 2
    End If
    sPattern = ItemsRowPattern(nVoucherInbox, bIncludePostings)
    Set oTransactionsRS = GetTransactions(sAppServer, nVoucherInbox, nUserId, bIncludePostings, _
                                          sWhere, nOrderColumn, sOrderBy)
    With oTransactionsRS
      Do
        If .EOF Then
          Exit Do
        End If
        nCurrentGralLedgerGroupId = !id_grupo_mayor
        nCurrentGralLedgerId = !id_mayor
        sHTML = sHTML & DivisionLine(nCurrentGralLedgerGroupId, _
                                     !nombre_grupo_mayor, 1, nColSpan, False) & "</TBODY>"
        nGroupItems = 0
        Do
          If .EOF Then
            Exit Do
          End If
          If (nCurrentGralLedgerGroupId <> !id_grupo_mayor) Then
            Exit Do
          End If
          nCurrentGralLedgerId = !id_mayor
          sItemsSection = DivisionLine(nCurrentGralLedgerId, _
                                      "[" & !numero_mayor & "] " & !nombre_mayor, 2, nColSpan, False)
          nItems = 0
          Do
            If .EOF Then
              Exit Do
            End If
            If (nCurrentGralLedgerGroupId <> !id_grupo_mayor) Or (nCurrentGralLedgerId <> !id_mayor) Then
              Exit Do
            End If
            sTemp = sPattern
            sTemp = Replace(sTemp, "<@ID@>", !id_transaccion)
            sTemp = Replace(sTemp, "<@SEPARATOR@>", sItemSeparator)
            sTemp = Replace(sTemp, "<@APPLICATION_DATE@>", Format$(!fecha_afectacion, "dd/mmm/yyyy"))
            sTemp = Replace(sTemp, "<@ELABORATION_DATE@>", Format$(!fecha_registro, "dd/mmm/yyyy hh:nn"))
            sTemp = Replace(sTemp, "<@TRANSACTION_TYPE@>", !tipo_transaccion)
            sTemp = Replace(sTemp, "<@VOUCHER_TYPE@>", !tipo_poliza)
            sTemp = Replace(sTemp, "<@CONCEPT@>", IIf(IsNull(!concepto_transaccion), "", !concepto_transaccion))
            sTemp = Replace(sTemp, "<@VOUCHER_NUMBER@>", !numero_transaccion)
            If (!esta_abierta) Then
              If (!id_autorizada_por = 0) And (!id_elaborada_por = nUserId) Then
                sTemp = Replace(sTemp, "<@ID2@>", (-1 * !id_transaccion))
                sTemp = Replace(sTemp, "<@URL@>", "voucher_editor.asp")
                sTemp = Replace(sTemp, "<@TARGET@>", "")
                sTemp = Replace(sTemp, "<@STATUS@>", "Pendiente")
              ElseIf ((!id_autorizada_por <> 0) And (!id_autorizada_por = nUserId)) Then
                sTemp = Replace(sTemp, "<@ID2@>", (-1 * !id_transaccion))
                sTemp = Replace(sTemp, "<@URL@>", "voucher_editor.asp")
                sTemp = Replace(sTemp, "<@TARGET@>", "")
                sTemp = Replace(sTemp, "<@STATUS@>", "Revisar póliza de " & !elaborada_por)
              ElseIf ((!id_autorizada_por <> 0) And (!id_elaborada_por = nUserId) And (!id_autorizada_por <> nUserId)) Then
                sTemp = Replace(sTemp, "<@ID2@>", (!id_transaccion))
                sTemp = Replace(sTemp, "<@URL@>", "pending_voucher_viewer.asp")
                sTemp = Replace(sTemp, "<@TARGET@>", "target='_blank'")
                sTemp = Replace(sTemp, "<@STATUS@>", "En revisión por: " & !autorizada_por)
              ElseIf (!id_autorizada_por <> nUserId) And (!id_elaborada_por <> nUserId) Then
                sTemp = Replace(sTemp, "<@ID2@>", (!id_transaccion))
                sTemp = Replace(sTemp, "<@URL@>", "pending_voucher_viewer.asp")
                sTemp = Replace(sTemp, "<@TARGET@>", "target='_blank'")
                If (!id_autorizada_por <> 0) Then
                  sTemp = Replace(sTemp, "<@STATUS@>", "En revisión por: " & !autorizada_por)
                Else
                  sTemp = Replace(sTemp, "<@STATUS@>", "En elaboración por:" & !elaborada_por)
                End If
              End If
            Else
              sTemp = Replace(sTemp, "<@ID2@>", !id_transaccion)
              sTemp = Replace(sTemp, "<@URL@>", "voucher_viewer.asp")
              sTemp = Replace(sTemp, "<@TARGET@>", "target='_blank'")
              sTemp = Replace(sTemp, "<@STATUS@>", "En el diario")
            End If
            If bIncludePostings Then
              sTemp = Replace(sTemp, "<@DEBIT_SUM@>", Format$(!cargos, "Currency"))
              sTemp = Replace(sTemp, "<@CREDIT_SUM@>", Format$(!abonos, "Currency"))
            End If
            sItemsSection = sItemsSection & sTemp
            nItems = nItems + 1
            .MoveNext
          Loop
          sHTML = sHTML & Replace(sItemsSection, "<@COUNTER@>", nItems)
          nGroupItems = nGroupItems + nItems
        Loop
        sHTML = Replace(sHTML, "<@COUNTER@>", nGroupItems)
      Loop
    End With
    Body = sHTML
  Exit Function
ErrHandler:
  RaiseError ClassId, "Body", Err.Number
End Function

Public Function CboBalancingTypes(Optional nSelectedBalancingType As Long = 0) As String
  Dim sTemp As String
  '*************************************************************************************
  On Error GoTo ErrHandler
    sTemp = "<OPTION value=0 selected> -- Todas las pólizas sin importar su balance-- </OPTION> " & _
            "<OPTION value=1>Sólo las pólizas que no están balanceadas</OPTION> " & _
            "<OPTION value=2>Sólo las pólizas balanceadas</OPTION>"
    If nSelectedBalancingType <> 0 Then
      sTemp = Replace(sTemp, _
                      "<OPTION value=" & nSelectedBalancingType & ">", _
                      "<OPTION SELECTED value=" & nSelectedBalancingType & ">")
    End If
    CboBalancingTypes = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "CboBalancingTypes", Err.Number
End Function

Public Function CboVoucherInboxes(Optional tSelectedVoucherInbox As TEnumVoucherInboxes = 0) As String
  Dim sTemp As String, i As Long
  '***************************************************************************************************
  On Error GoTo ErrHandler
    For i = 1 To cnMaxNumberOfInboxes
      sTemp = sTemp & "<OPTION value=" & i & ">" & VoucherInboxName(i) & "</OPTION>"
    Next i
    If (tSelectedVoucherInbox <> 0) Then
      sTemp = Replace(sTemp, _
                      "<OPTION value=" & tSelectedVoucherInbox & ">", _
                      "<OPTION SELECTED value=" & tSelectedVoucherInbox & ">")
    End If
    CboVoucherInboxes = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "CboVoucherInboxes", Err.Number
End Function

Public Function Header(sAppServer As String, nVoucherInbox As TEnumVoucherInboxes, _
                       Optional nSelectedColumn As Long, _
                       Optional bShowAmountsCols As Boolean = False) As String
  Dim sTemp As String, nColSpan As Long, i As Long
  '*********************************************************************************
  On Error GoTo ErrHandler
    sTemp = cnHeaderCol0
    Select Case nVoucherInbox
      Case efaMyPendingVouchers, efaGroupPendingVouchers
        sTemp = sTemp & cnHeaderCol2 & cnHeaderCol3 & cnHeaderCol4
      Case Else
        sTemp = sTemp & cnHeaderCol1 & cnHeaderCol2 & cnHeaderCol3 & cnHeaderCol4
    End Select
    nColSpan = 2
    If bShowAmountsCols Then
      sTemp = sTemp & cnHeaderCol5
      nColSpan = nColSpan + 2
    End If
    If nVoucherInbox = efaGroupPostedVouchers Then
      sTemp = sTemp & cnHeaderCol6
      nColSpan = nColSpan + 2
    End If
    sTemp = sTemp & cnHeaderCol7

    sTemp = Replace(sTemp, "<@COLSPAN@>", nColSpan)
    If (nSelectedColumn > 0) Then
      sTemp = Replace(sTemp, "<@MARK_" & nSelectedColumn & "@>", " <img src='../../images/order_asc_red.gif'>")
    ElseIf (nSelectedColumn < 0) Then
      sTemp = Replace(sTemp, "<@MARK_" & Abs(nSelectedColumn) & "@>", " <img src='../../images/order_desc_red.gif'>")
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

Private Function DivisionLine(nObjectId As Long, sObjectName As String, _
                              nLevel As Long, nColSpan As Long, bExpanded As Boolean) As String
  Dim sTemp As String
  '**************************************************************************
  On Error GoTo ErrHandler
    sTemp = IIf(bExpanded, cnExpandedDivLine, cnCollapsedDivLine)
    sTemp = Replace(sTemp, "<@LEVEL@>", nLevel)
    sTemp = Replace(sTemp, "<@SEPARATOR@>", Separator(nLevel))
    sTemp = Replace(sTemp, "<@NAME@>", sObjectName)
    sTemp = Replace(sTemp, "<@COL_SPAN@>", nColSpan)
    DivisionLine = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "DivisionLine", Err.Number
End Function

Private Function GetTransactions(sAppServer As String, nVoucherInbox As TEnumVoucherInboxes, _
                                 nUserId As Long, bIncludePostings As Boolean, _
                                 sWhere As String, nOrderColumn As Long, sOrderBy As String) As Recordset
  Dim oTransactionsDS As New EFATransactionsDS.CTransactions
  '******************************************************************************************************
  On Error GoTo ErrHandler
    With oTransactionsDS
      Select Case nVoucherInbox
        Case efaMyPendingVouchers
          Set GetTransactions = .GetPending(sAppServer, bIncludePostings, sWhere, sOrderBy)
        Case efaMyPostedVouchers
          Set GetTransactions = .GetPosted(sAppServer, bIncludePostings, sWhere, sOrderBy)
        Case efaGroupPendingVouchers
          Set GetTransactions = .GetPending(sAppServer, bIncludePostings, sWhere, sOrderBy)
        Case efaGroupPostedVouchers
          Set GetTransactions = .GetPosted(sAppServer, bIncludePostings, sWhere, sOrderBy)
      End Select
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetTransactions", Err.Number
End Function
        
Private Function ItemsRowPattern(nVoucherInbox As TEnumVoucherInboxes, bShowAmountsCols As Boolean) As String
  Dim sTemp As String
  '**********************************************************************************************************
  On Error GoTo ErrHandler
    sTemp = cnItemLine0
    Select Case nVoucherInbox
      Case efaMyPendingVouchers, efaGroupPendingVouchers
        sTemp = sTemp & cnItemLine2
      Case Else
        sTemp = sTemp & cnItemLine1 & cnItemLine2
    End Select
    If bShowAmountsCols Then
      sTemp = sTemp & cnItemLine3
    End If
    If nVoucherInbox = efaGroupPostedVouchers Then
      sTemp = sTemp & cnItemLine4
    End If
    sTemp = sTemp & cnItemLine5
    ItemsRowPattern = sTemp
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
