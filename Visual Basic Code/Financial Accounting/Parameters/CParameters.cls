VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*** Empiria® ***********************************************************************************************
'*                                                                                                          *
'* Solución   : Empiria® Software Components                    Sistema : Financial Accounting              *
'* Componente : Parameters (EFAParameters)                      Parte   : Business services                 *
'* Clase      : CParameters                                     Patrón  : Stateless class                   *
'* Fecha      : 31/Enero/2002                                   Versión : 1.0       Versión patrón: 1.0     *
'*                                                                                                          *
'* Descripción: Clase principal de los parámetros del sistema de contabilidad financiera.                   *
'*                                                                                                          *
'****************************************************** Copyright © La Vía Ontica, S.C. México, 1999-2002. **
Option Explicit
Private Const ClassId As String = "CParameters"

'************************************************************************************************************
'* MÉTODOS PÚBLICOS                                                                                         *
'************************************************************************************************************

Public Function GeneratedReportsPars(sAppServer As String, nGLGroupId As Long, nGralLedgerId As Long, _
                                     Optional dFromDate As Variant, Optional dToDate As Variant) As String
  Dim sTemp As String
  '*******************************************************************************************************
  On Error GoTo ErrHandler
    sTemp = GetGLGroupName(sAppServer, nGLGroupId)
    sTemp = sTemp & GetGLNumberAndName(sAppServer, nGralLedgerId)
    sTemp = sTemp & GetGLSigners(sAppServer, nGralLedgerId)
    sTemp = sTemp & GetDates(dFromDate, dToDate)
    GeneratedReportsPars = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "GeneratedReportsPars", Err.Number
End Function

Private Function GetDates(dFromDate As Variant, dToDate As Variant) As String
  Dim sTemp As String
  On Error GoTo ErrHandler
    If IsMissing(dFromDate) And IsMissing(dToDate) Then
      sTemp = "FromDate:=;ToDate:=;DateRange:=;"
    ElseIf (Not IsMissing(dFromDate)) And IsMissing(dToDate) Then
      sTemp = LongDate(dFromDate)
      sTemp = "FromDate:=" & sTemp & ";ToDate:=;DateRange:=;"
    ElseIf IsMissing(dFromDate) And (Not IsMissing(dToDate)) Then
      sTemp = LongDate(dToDate)
      sTemp = "FromDate:=;ToDate:=" & sTemp & ";DateRange:=;"
    ElseIf (Not IsMissing(dFromDate)) And (Not IsMissing(dToDate)) Then
      sTemp = "FromDate:=" & LongDate(dFromDate) & ";ToDate:=" & LongDate(dToDate) & ";"
      If CDate(dFromDate) = CDate(dToDate) Then
        sTemp = sTemp & "DateRange:=" & LongDate(dToDate) & ";"
      Else
        sTemp = sTemp & "DateRange:="
        If (Year(dFromDate) <> Year(dToDate)) Then
          sTemp = sTemp & "Del " & LongDate(dFromDate) & " al " & LongDate(dToDate)
        ElseIf (Month(dFromDate) <> Month(dToDate)) Then
          sTemp = sTemp & "Del " & LCase$(Format$(dFromDate, "d ""de"" mmmm")) & " al " & LongDate(dToDate)
        Else
          sTemp = sTemp & "Del " & Format$(dFromDate, "d") & " al " & LCase$(Format$(dToDate, "d"" de ""mmmm"" del ""yyyy"))
        End If
        sTemp = sTemp & ";"
      End If
    End If
    GetDates = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetDates", Err.Number
End Function

Private Function GetGLGroupName(sAppServer As String, nGLGroupId As Long) As String
  Dim oRecordset As New Recordset, sSQL As String
  '********************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT nombre_grupo_mayor FROM COF_Grupo_Mayor " & _
           "WHERE (id_grupo_mayor = " & nGLGroupId & ")"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      If Not (.BOF And .EOF) Then
        If IsNumeric(Left$(!nombre_grupo_mayor, 1)) Then
          GetGLGroupName = "GralLedgerGroup:=" & Mid$(!nombre_grupo_mayor, 11) & ";"
        Else
          GetGLGroupName = "GralLedgerGroup:=" & !nombre_grupo_mayor & ";"
        End If
      Else
        GetGLGroupName = "GralLedgerGroup:=Banco Nacional de Obras y Servicios Públicos, S.N.C.;"
      End If
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetGLGroupName", Err.Number
End Function

Private Function GetGLNumberAndName(sAppServer As String, nGralLedgerId As Long) As String
  Dim oRecordset As New Recordset, sSQL As String, sTemp As String
  '***************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT nombre_mayor, numero_mayor FROM COF_Mayor " & _
           "WHERE (id_mayor = " & nGralLedgerId & ")"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      If Not (.BOF And .EOF) Then
        sTemp = "GralLedgerNumber:=" & !numero_mayor & ";"
        sTemp = sTemp & "GralLedgerName:=" & !nombre_mayor & ";"
        sTemp = sTemp & "GralLedgerNumberAndName:=(" & !numero_mayor & ") " & !nombre_mayor & ";"
      Else
        sTemp = "GralLedgerNumber:=;GralLedgerName:=;GralLedgerNumberAndName:=;"
      End If
      .Close
    End With
    GetGLNumberAndName = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetGLNumberAndName", Err.Number
End Function

Private Function GetGLSigners(sAppServer As String, nGralLedgerId As Long) As String
  Dim sTemp As String
  On Error GoTo ErrHandler
    If nGralLedgerId <= 32 Then
      sTemp = sTemp & "SignNameA:=LIC. TOMAS RUIZ GONZALEZ;"
      sTemp = sTemp & "SignNameB:=C.P. NOEMI DAVID RODRIGUERA;"
      sTemp = sTemp & "SignNameC:=LIC. RUBEN VAZQUEZ VALDEZ;"
      sTemp = sTemp & "SignTitleA:=DIRECTOR GENERAL;"
      sTemp = sTemp & "SignTitleB:=DIRECTORA DE CREDITO Y CONTABILIDAD;"
      sTemp = sTemp & "SignTitleC:=SUBDIRECTOR DE CONTABILIDAD;"
    Else
      sTemp = sTemp & "SignNameA:=LIC. VICTOR MANUEL LOZANO PROA;"
      sTemp = sTemp & "SignNameB:=LIC. JOSE ANTONIO GARZA Y OCHOA;"
      sTemp = sTemp & "SignNameC:=;"
      sTemp = sTemp & "SignTitleA:=SUBDIRECTOR FIDUCIARIO;"
      sTemp = sTemp & "SignTitleB:=DELEGADO FIDUCIARIO;"
      sTemp = sTemp & "SignTitleC:=;"
    End If
    GetGLSigners = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetGLSigners", Err.Number
End Function

Private Function LongDate(dDate As Variant) As String
  On Error GoTo ErrHandler
    LongDate = LCase$(Format$(dDate, "d"" de ""mmmm"" del ""yyyy"))
  Exit Function
ErrHandler:
  RaiseError ClassId, "LongDate", Err.Number
End Function













'Private Function RoundedToName(sAppServer As String, nRoundedTo As Long) As String
'  On Error GoTo ErrHandler
'    Select Case nRoundedTo
'      Case 1
'        RoundedToName = "Cifras en pesos (sin centavos)"
'      Case 2
'        RoundedToName = "Cifras en miles de pesos"
'      Case 3
'        RoundedToName = "Cifras en millones de pesos"
'      Case Else
'        RoundedToName = "No redondear"
'    End Select
'  Exit Function
'ErrHandler:
'  RaiseError ClassId, "RoundedToName", Err.Number
'End Function
'
'Private Function SourceName(sAppServer As String, nGLGroupId As Long, nGralLedgerId As Long) As String
'  On Error GoTo ErrHandler
'    If (nGLGroupId <> 0) And (nGralLedgerId <> 0) Then
'      SourceName = GetGLParentGroupName(sAppServer, nGLGroupId, nGralLedgerId)
'    ElseIf (nGLGroupId <> 0) And (nGralLedgerId = 0) Then
'      SourceName = GetGLParentGroupName(sAppServer, nGLGroupId, Array(1, 2))
'    ElseIf (nGLGroupId = 0) And (nGralLedgerId <> 0) Then
'      SourceName = GetGLParentGroupName(sAppServer, nGLGroupId, Array(1, 2))
'    ElseIf (nGLGroupId = 0) And (nGralLedgerId = 0) Then
'      SourceName = GetGLParentGroupName(sAppServer, nGLGroupId, Array(1, 2))
'    End If
'  Exit Function
'ErrHandler:
'  RaiseError ClassId, "SourceName", Err.Number
'End Function
'
'Private Function GetGLParentGroupName(sAppServer As String, nGLGroupId As Long, vGL As Variant) As String
'  Dim oRecordset As New Recordset, sSQL As String, nGralLedgerId  As Long
'  '******************************************************************************************************
'  On Error GoTo ErrHandler
'    If IsArray(vGL) Then
'      sSQL = "SELECT COF_Grupo_Mayor2.id_grupo_mayor_padre, COF_Grupo_Mayor2.nombre_grupo_mayor Nombre " & _
'             "FROM COF_Grupo_Mayor, COF_Grupo_Mayor COF_Grupo_Mayor2 " & _
'             "WHERE (COF_Grupo_Mayor.id_grupo_mayor = " & nGLGroupId & ") AND " & _
'             "(COF_Grupo_Mayor.id_grupo_mayor_padre = COF_Grupo_Mayor2.id_grupo_mayor)"
'    Else
'      sSQL = "SELECT id_mayor id_grupo_mayor_padre, nombre_mayor Nombre, numero_mayor FROM COF_Mayor " & _
'             "WHERE (id_mayor = " & vGL & ")"
'    End If
'    With oRecordset
'      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
'      If Not (.BOF And .EOF) Then
'        If !id_grupo_mayor_padre <> 0 Then
'          If IsArray(vGL) Then
'            GetGLParentGroupName = Mid$(!nombre, 11)
'          Else
'            GetGLParentGroupName = "(" & !numero_mayor & ") " & !nombre
'          End If
'        Else
'          GetGLParentGroupName = "Consolidado"
'        End If
'      Else
'        GetGLParentGroupName = ""
'      End If
'      .Close
'    End With
'  Exit Function
'ErrHandler:
'  RaiseError ClassId, "GetGLParentGroupName", Err.Number
'End Function
'
'Private Function GetRoundLevelLabel(nRoundLevel As Long) As String
'  On Error GoTo ErrHandler
'    Select Case nRoundLevel
'      Case 0
'        GetRoundLevelLabel = "(Cifras en pesos)"
'      Case 1
'        GetRoundLevelLabel = "(Cifras en miles pesos)"
'      Case 2
'        GetRoundLevelLabel = "(Cifras en millones de pesos)"
'      Case Else
'        GetRoundLevelLabel = "NO DEFINIDO"
'    End Select
'  Exit Function
'ErrHandler:
'  RaiseError ClassId, "GetRoundLevelLabel", Err.Number
'End Function

