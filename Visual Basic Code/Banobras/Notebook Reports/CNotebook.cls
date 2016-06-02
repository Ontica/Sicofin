VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CNotebook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*** Sistema de contabilidad financiera (SICOFIN) ***********************************************************
'*                                                                                                          *
'* Solución   : Customer Components                             Sistema : Financial Accounting              *
'* Componente : Notebook (SCFNotebook)                          Parte   : Business services                 *
'* Clase      : CNotebook                                       Patrón  : Stateless Class                   *
'* Fecha      : 15/Enero/2002                                   Versión : 1.1       Versión patrón: 1.0     *
'*                                                                                                          *
'* Descripción: Genera los diferentes reportes que constituyen el cuaderno de delegaciones estatales.       *
'*                                                                                                          *
'****************************************************** Copyright © La Vía Ontica, S.C. México, 2001-2002. **
Option Explicit
Private Const ClassId As String = "CNoteBook"

Private Const cnPageLines As Long = 45

'************************************************************************************************************
'* PROPIEDADES PÚBLICAS                                                                                     *
'************************************************************************************************************

Public Property Get URLFilesPath() As String
  On Error GoTo ErrHandler
    URLFilesPath = GetConstant(cnURLFilesPath)
  Exit Property
ErrHandler:
  RaiseError ClassId, "URLFilesPath", Err.Number
End Property

'************************************************************************************************************
'* MÉTODOS PÚBLICOS                                                                                         *
'************************************************************************************************************

Public Sub Generate(sAppServer As String, nGl As Long, dIniDate As Date, dEndDate As Date, _
                    dExcDate As Date, sName1 As String, sPosition1 As String, _
                    sName2 As String, sPosition2 As String, Optional nExcRateType As Long = 46, _
                    Optional nExcRateCurrency As Long = 1, _
                    Optional sIndexFileName As Variant, Optional sVerticalFileName As Variant, _
                    Optional sHorizontalFileName As Variant)
  Dim oNotebookRulesFile As New CTextFile
  Dim oIndexFile As New CTextFile, oVerticalFile As New CTextFile, oHorizontalFile As New CTextFile
  Dim sLine As String, vLine As Variant, sAccount As String
  Dim sName As String, i As Long, sAccountsFile As String, sFolio As String, sTitle As String
  '************************************************************************************************
  On Error GoTo ErrHandler
    oIndexFile.Open_ GetConstant(cnGeneratedFilesPath)
    oVerticalFile.Open_ GetConstant(cnGeneratedFilesPath)
    oHorizontalFile.Open_ GetConstant(cnGeneratedFilesPath)
    oNotebookRulesFile.Open_ GetConstant(cnPatternsPath), GetConstant(cnNotebookReportsFile), True
    Do While Not oNotebookRulesFile.EOF
      sLine = oNotebookRulesFile.GetNextLine
      vLine = Split(sLine, ";")
      sFolio = vLine(0)
      sTitle = vLine(1)
      sAccount = vLine(2)
      For i = 3 To UBound(vLine)
        If (Trim$(vLine(i)) <> "") Then
           Select Case Trim(vLine(i))
               Case "7000"
                  Cuaderno_7000 oVerticalFile, sAppServer, dEndDate, sAccount, nGl, "7000", _
                                sName1, sPosition1, sName2, sPosition2, False
               Case "7001"
                  Cuaderno_7001 oVerticalFile, sAppServer, dEndDate, nGl, sAccount, _
                                sName1, sPosition1, sName2, sPosition2
               Case "7002"  'Horizontal
                  Cuaderno_7002 oHorizontalFile, sAppServer, nGl, dEndDate, sAccount, _
                                sName1, sPosition1, sName2, sPosition2, dExcDate, nExcRateType, nExcRateCurrency
               'Case "7006"   'Horizontal
               '   Cuaderno_7006 oHorizontalFile, sAppServer, dEndDate, dExcDate, _
                                Array("1305", "1307", "1308", "1310", "1311", "1312", "1314", "1315", "1316", "1317", "1319", "1320", "1323", "1324", "1337", "1350", "6378"), _
                                nGL, sName1, sPosition1, sName2, sPosition2, nExcRateType, nExcRateCurrency
               Case "7010"
                  Cuaderno_7010 oVerticalFile, sAppServer, dEndDate, sAccount, nGl, "7010", _
                                sName1, sPosition1, sName2, sPosition2
               Case "7011"
                  Cuaderno_7011 oVerticalFile, sAppServer, dEndDate, sAccount, nGl, "7011", _
                                sName1, sPosition1, sName2, sPosition2
               
               Case "7013"
                  Cuaderno_7013 oVerticalFile, sAppServer, dEndDate, sAccount, nGl, "7013", _
                                sName1, sPosition1, sName2, sPosition2
               Case "7014"
                 Cuaderno_7014 oVerticalFile, sAppServer, dEndDate, sAccount, nGl, "7014", _
                               sName1, sPosition1, sName2, sPosition2, False
               'Case "7016"
               '   Cuaderno_7016 oVerticalFile, sAppServer, nGL, "01", dIniDate, dEndDate, dExcDate, _
                                sName1, sPosition1, sName2, sPosition2, nExcRateType, nExcRateCurrency
               Case "7021"    'Horizontal
                  Cuaderno_7021 oHorizontalFile, sAppServer, dEndDate, nGl, _
                                sName1, sPosition1, sName2, sPosition2
               Case "7024"
                  Cuaderno_7024 oVerticalFile, sAppServer, dEndDate, sAccount, nGl, "7024", _
                                  sName1, sPosition1, sName2, sPosition2
               Case "7026"
                  REL_ANALITICA_CUENTAS oVerticalFile, sAppServer, nGl, dIniDate, dEndDate, _
                                        sName1, sPosition1, sName2, sPosition2
               Case "7028"
                  Cuaderno_7028 oVerticalFile, sAppServer, dEndDate, sAccount, nGl, "7028", _
                                sName1, sPosition1, sName2, sPosition2
          End Select
        End If
      Next i
    Loop
    oNotebookRulesFile.Close_
    sIndexFileName = oIndexFile.FileName
    oIndexFile.Close_
    sHorizontalFileName = oHorizontalFile.FileName
    oHorizontalFile.Close_
    sVerticalFileName = oVerticalFile.FileName
    oVerticalFile.Close_
  Exit Sub
ErrHandler:
  RaiseError ClassId, "Generate", Err.Number
End Sub

'************************************************************************************************************
'* MÉTODOS PRIVADOS                                                                                         *
'************************************************************************************************************

Private Function Cuaderno_7000(oTextFile As CTextFile, sAppServer As String, dCuteDate As Date, _
                              sAccount As String, nGl As Long, sRepNum As String, _
                              sNameA As String, sPuestoA As String, _
                              sNameB As String, sPuestoB As String, _
                              Optional bWithSector As Boolean = False, _
                              Optional bWithResumen As Boolean = False) As String
  Dim oRs           As ADODB.Recordset
  Dim oRSGenerator  As New SCFFixedReportsDS.CNotebookData
  'Dim oTextFile     As New CTextFile
  
  Dim sPrintLine    As String
  Dim sPreviousAcc  As String
  Dim sSector       As String
  Dim nLine         As Long
  Dim nPage         As Long
  Dim nTotalSubAcc  As Currency
  Dim nTotal        As Currency
  Dim nTotSectores  As Currency
  Dim nSpcSector    As Long
  
  Dim sCompleteAcc  As String
  Dim sGL           As String
  Dim sAccountName  As String
  Dim sSubAccName   As String
  Dim nTotalconLetra As String
  '***********************************************************************************************
  On Error GoTo ErrHandler
    Set oRs = oRSGenerator.GetRS_7000(sAppServer, nGl, dCuteDate, sAccount)
    sGL = oRSGenerator.GetObjName(sAppServer, nGl, "COF_MAYOR", "NOMBRE_MAYOR", "ID_MAYOR")
    sAccountName = sAccount & " - " & oRSGenerator.GetAccountName(sAppServer, sAccount, nGl)

    nTotal = 0
    nTotalSubAcc = 0
    nLine = 100
    nPage = 1
    
    If ((oRs.BOF) And (oRs.EOF)) Then
      PrintHeader7000 oTextFile, nLine, nPage, dCuteDate, sGL, sAccountName, sRepNum, _
                                 sNameA, sPuestoA, sNameB, sPuestoB
      oTextFile.PrintLine " NO EXISTE INFORMACION "
      Exit Function
    End If
    
    With oRs
      sPreviousAcc = !numero_cuenta_estandar
      sCompleteAcc = GetCuentaSubCuenta(!numero_cuenta_estandar)
      'nSpcSector = 5
      
      Do While Not .EOF()
        If (!numero_cuenta_estandar <> sPreviousAcc) Then
          sSubAccName = sAccount & " - " & oRSGenerator.GetAccountName(sAppServer, sPreviousAcc, nGl)
          PrintHeader7000 oTextFile, nLine, nPage, dCuteDate, sGL, sAccountName, sRepNum, _
                          sNameA, sPuestoA, sNameB, sPuestoB
          oTextFile.PrintLine " SUMA PARCIAL " & _
                              SpcFill(Mid(sCompleteAcc, 6), 21) & _
                              SpcFill(Mid(sSubAccName, 7), 25) & _
                              SpcFillAmount(nTotalSubAcc, 22)
          oTextFile.PrintLine ""
          nLine = nLine + 2
          nTotal = nTotal + nTotalSubAcc
          nTotalSubAcc = 0
        End If
        PrintHeader7000 oTextFile, nLine, nPage, dCuteDate, sGL, sAccountName, sRepNum, _
                        sNameA, sPuestoA, sNameB, sPuestoB
        sPrintLine = SpcFill(!fecha_ult, 10) & " " & _
                     SpcFill(!nombre_cuenta_auxiliar, 35) & "           " & _
                     SpcFillAmount(!Saldo_Actual, 25)
        
        oTextFile.PrintLine sPrintLine
        oTextFile.PrintLine Space(32) & SpcFill(!descripcion, 35)
        nLine = nLine + 1
        nTotalSubAcc = nTotalSubAcc + !Saldo_Actual
        sPreviousAcc = !numero_cuenta_estandar
        sCompleteAcc = GetCuentaSubCuenta(!numero_cuenta_estandar)
        
        sCompleteAcc = sCompleteAcc
        .MoveNext
      Loop
      oTextFile.PrintLine Space(26) & _
                          SpcFill(sCompleteAcc, 21) & _
                          SpcFillAmount(nTotalSubAcc, 22)
                  
      nTotal = nTotal + nTotalSubAcc
      
      oTextFile.PrintNewLine
      oTextFile.PrintLine Space(75 - nSpcSector) & "----------------------"
      oTextFile.PrintLine Space(45 - nSpcSector) & "SALDO IGUAL AL MAYOR ------> " & SpcFillAmount(nTotal, 22)
      oTextFile.PrintLine Space(75 - nSpcSector) & "======================"
    End With
    oRs.Close
    'nTotalconLetra = CantidadConLetra(ntotal)
    PrintSigns oTextFile, sNameA, sPuestoA, sNameB, sPuestoB, _
                    Len("================================================================================================================")
    'oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Cuaderno_7000", Err.Number
End Function

Private Function Cuaderno_7001(oTextFile As CTextFile, sAppServer As String, dCuteDate As Date, _
                              nGl As Long, sAccount As String, _
                              sNameA As String, sPuestoA As String, _
                              sNameB As String, sPuestoB As String) As String
  Dim oRs            As ADODB.Recordset
  Dim oRSGenerator   As New SCFFixedReportsDS.CNotebookData
  Dim sPreviousAcc   As String
  Dim sSector        As String
  Dim nLine          As Long
  Dim nPage          As Long
  Dim nTotalSubAcc   As Currency
  Dim nTotal         As Currency
  Dim nTotSectores   As Currency
  
  Dim sGL            As String
  Dim sAccountName   As String
  Dim nTotalconLetra As String
  '********************************************************************************
  On Error GoTo ErrHandler
    Set oRs = oRSGenerator.GeneraAuxiliares(sAppServer, dCuteDate, sAccount, nGl)
    sGL = oRSGenerator.GetObjName(sAppServer, nGl, "COF_MAYOR", "NOMBRE_MAYOR", "ID_MAYOR")
    sAccountName = sAccount & " - " & oRSGenerator.GetAccountName(sAppServer, sAccount, nGl)
    nTotal = 0
    nTotalSubAcc = 0
    nLine = 100
    nPage = 1
    
    If ((oRs.BOF) And (oRs.EOF)) Then
      PrintHeader7001 oTextFile, nLine, nPage, dCuteDate, sGL, _
                      sNameA, sPuestoA, sNameB, sPuestoB, sAccountName
      oTextFile.PrintLine " NO EXISTE INFORMACION "
      Exit Function
    End If
    
    With oRs
      sPreviousAcc = !numero_cuenta_estandar
      sSector = !clave_sector
      Do While Not .EOF()
        If (!numero_cuenta_estandar <> sPreviousAcc) Or (sSector <> !clave_sector) Then
          PrintHeader7001 oTextFile, nLine, nPage, dCuteDate, sGL, sNameA, sPuestoA, sNameB, sPuestoB, sAccountName
          oTextFile.PrintLine Space(40) & _
                              SpcFill(sPreviousAcc, 21) & " " & _
                              SpcFill(sSector, 4) & Space(17) & _
                              SpcFillAmount(nTotalSubAcc, 22)
          oTextFile.PrintLine ""
          nLine = nLine + 2
          nTotal = nTotal + nTotalSubAcc
          nTotalSubAcc = 0
        End If
        PrintHeader7001 oTextFile, nLine, nPage, dCuteDate, sGL, sNameA, sPuestoA, sNameB, sPuestoB, sAccountName
        oTextFile.PrintLine SpcFill(!NUMERO_CUENTA_AUXILIAR, 20) & " " & _
                            SpcFill(!clave_sector, 4) & " " & _
                            SpcFill(!nombre_cuenta_auxiliar, 35) & " " & _
                            SpcFillAmount(!Saldo_Actual, 22)
        nLine = nLine + 1
        nTotalSubAcc = nTotalSubAcc + !Saldo_Actual
        sPreviousAcc = !numero_cuenta_estandar
        sSector = !clave_sector
        .MoveNext
      Loop
      oTextFile.PrintLine Space(40) & _
                          SpcFill(sPreviousAcc, 21) & " " & _
                          SpcFill(sSector, 4) & Space(17) & _
                          SpcFillAmount(nTotalSubAcc, 22)
                  
      nTotal = nTotal + nTotalSubAcc
      
      oTextFile.PrintNewLine
      oTextFile.PrintLine Space(84) & "----------------------"
      oTextFile.PrintLine Space(54) & "SALDO IGUAL AL MAYOR ------> " & SpcFillAmount(nTotal, 22)
      oTextFile.PrintLine Space(84) & "======================"
    End With
    oRs.Close
    Set oRs = oRSGenerator.GeneraResumen(sAppServer, dCuteDate, "1103", nGl)
    oTextFile.PrintNewLine
    
    PrintHeader7001 oTextFile, nLine, nPage, dCuteDate, sGL, sNameA, sPuestoA, sNameB, sPuestoB, sAccountName
    
    oTextFile.PrintLine SpcFill(" ", 6) & "     RESUMEN DE SECTORES      "
    oTextFile.PrintLine SpcFill(" ", 6) & " SECTOR               IMPORTE "
    oTextFile.PrintLine SpcFill(" ", 6) & "------------------------------"

    With oRs
      Do
        If (.EOF) Then Exit Do
        oTextFile.PrintLine SpcFill(" ", 7) & SpcFill(!clave_sector, 6) & SpcFillAmount(!Saldo_Actual, 22)
        nTotSectores = nTotSectores + !Saldo_Actual
        .MoveNext
      Loop
    End With
    oTextFile.PrintLine SpcFill(" ", 6) & "------------------------------"
    oTextFile.PrintLine SpcFill(" ", 13) & SpcFillAmount(nTotSectores, 22)
    PrintFeed oTextFile, "1103"
    'nTotalconLetra = CantidadConLetra(nTotal)
    PrintSigns oTextFile, sNameA, sPuestoA, sNameB, sPuestoB, _
               Len("==========================================================================================================")
  Exit Function
ErrHandler:
  RaiseError ClassId, "Cuaderno_7001", Err.Number
End Function

Private Sub PrintHeader7001(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                            dCuteDate As Date, sGL As String, _
                            sNameA As String, sPuestoA As String, _
                            sNameB As String, sPuestoB As String, _
                            sAccountName As String)
  Dim sDate       As String
  '**********************************************************
  On Error GoTo ErrHandler
    If (nLine >= cnPageLines) Then
      If nLine <> 100 Then
         PrintSigns oTextFile, sNameA, sPuestoA, sNameB, sPuestoB, _
                    Len("==========================================================================================================")
      End If
      sDate = Format(dCuteDate, "d"" DE ""MMMM"" DE ""yyyy")
      
      oTextFile.PrintLine "TECOYOTITLA NUM 100             BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C.      FECHA: " & Date
      oTextFile.PrintLine "COLONIA FLORIDA                  " & SpcFill(sAccountName & sDate, 57) & "PAGINA:   " & nPage
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F.          ===================================================="
      oTextFile.PrintLine "BNO-670315-CDO"
      oTextFile.PrintLine "7001"
      oTextFile.PrintLine "                                  DELEGACION   : " & sGL
      oTextFile.PrintLine "=========================================================================================================="
      oTextFile.PrintLine "  AUXILIAR        SECTOR  DESCRIPCION                                     IMPORTE                 TOTAL"
      oTextFile.PrintLine "----------------------------------------------------------------------------------------------------------"
      nLine = 20
      nPage = nPage + 1
    End If
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintHeader7001", Err.Number
End Sub

Private Function Cuaderno_7011(oTextFile As CTextFile, sAppServer As String, dCuteDate As Date, _
                              sAccount As String, nGl As Long, sRepNum As String, _
                              sNameA As String, sPuestoA As String, _
                              sNameB As String, sPuestoB As String, _
                              Optional bWithSector As Boolean = False) As String
  Dim oRs            As ADODB.Recordset
  Dim oRSGenerator   As New SCFFixedReportsDS.CNotebookData
  Dim sPrintLine     As String
  Dim sPreviousAcc   As String
  Dim sSector        As String
  Dim nLine          As Long
  Dim nPage          As Long
  Dim nTotalSubAcc   As Currency
  Dim nTotal         As Currency
  Dim nTotSectores   As Currency
  Dim nSpcSector     As Long
  Dim sCompleteAcc   As String
  
  Dim sGL            As String
  Dim sAccountName   As String
  Dim nTotalconLetra As String
  '********************************************************************************
  On Error GoTo ErrHandler
    Set oRs = oRSGenerator.GeneraAuxiliares(sAppServer, dCuteDate, sAccount, nGl)
    sGL = oRSGenerator.GetObjName(sAppServer, nGl, "COF_MAYOR", "NOMBRE_MAYOR", "ID_MAYOR")
    sAccountName = sAccount & " - " & oRSGenerator.GetAccountName(sAppServer, sAccount, nGl)
   
    nTotal = 0
    nTotalSubAcc = 0
    nLine = 100
    nPage = 1
    
    If ((oRs.BOF) And (oRs.EOF)) Then
      PrintHeaderAuxFec oTextFile, nLine, nPage, dCuteDate, sGL, sAccountName, sRepNum, _
                        sNameA, sPuestoA, sNameB, sPuestoB, bWithSector
                        
      oTextFile.PrintLine " NO EXISTE INFORMACION "
      Exit Function
    End If
    
    With oRs
      sPreviousAcc = !numero_cuenta_estandar
      sSector = !clave_sector
      sCompleteAcc = GetCuentaSubCuenta(!numero_cuenta_estandar)
      sCompleteAcc = sCompleteAcc & !clave_sector
      
      If (bWithSector) Then
        nSpcSector = 0
      ElseIf Not (bWithSector) Then
        nSpcSector = 5
      End If
      
      Do While Not .EOF()
        If (!numero_cuenta_estandar <> sPreviousAcc) Or (sSector <> !clave_sector) Then
          PrintHeaderAuxFec oTextFile, nLine, nPage, dCuteDate, sGL, sAccountName, sRepNum, _
                            sNameA, sPuestoA, sNameB, sPuestoB, bWithSector
          oTextFile.PrintLine Space(51 - nSpcSector) & _
                              SpcFill(sCompleteAcc, 21) & " " & _
                              SpcFill(" ", 4) & Space(17) & _
                              SpcFillAmount(nTotalSubAcc, 22)
          oTextFile.PrintLine ""
          nLine = nLine + 2
          nTotal = nTotal + nTotalSubAcc
          nTotalSubAcc = 0
        End If
        
        PrintHeaderAuxFec oTextFile, nLine, nPage, dCuteDate, sGL, sAccountName, sRepNum, _
                          sNameA, sPuestoA, sNameB, sPuestoB, bWithSector
        
        If (bWithSector) Then
          sPrintLine = SpcFill(!fecha_ult, 10) & " " & _
                       SpcFill(IIf(IsNull(!NUMERO_CUENTA_AUXILIAR), " ", !NUMERO_CUENTA_AUXILIAR), 20) & " " & _
                       SpcFill(!clave_sector, 4) & " " & _
                       SpcFill(IIf(IsNull(!nombre_cuenta_auxiliar), " ", !nombre_cuenta_auxiliar), 35) & " " & _
                       SpcFillAmount(!Saldo_Actual, 22)
        ElseIf Not (bWithSector) Then
          sPrintLine = SpcFill(!fecha_ult, 10) & " " & _
                       SpcFill(IIf(IsNull(!NUMERO_CUENTA_AUXILIAR), " ", !NUMERO_CUENTA_AUXILIAR), 20) & " " & _
                       SpcFill(IIf(IsNull(!nombre_cuenta_auxiliar), " ", !nombre_cuenta_auxiliar), 35) & " " & _
                       SpcFillAmount(!Saldo_Actual, 22)
        End If
        
        oTextFile.PrintLine sPrintLine
        nLine = nLine + 1
        nTotalSubAcc = nTotalSubAcc + !Saldo_Actual
        sPreviousAcc = !numero_cuenta_estandar
        sSector = !clave_sector
        sCompleteAcc = GetCuentaSubCuenta(!numero_cuenta_estandar)
        sCompleteAcc = sCompleteAcc & !clave_sector
        .MoveNext
      Loop
      oTextFile.PrintLine Space(51 - nSpcSector) & _
                          SpcFill(sCompleteAcc, 21) & " " & _
                          SpcFill(" ", 4) & Space(17) & _
                          SpcFillAmount(nTotalSubAcc, 22)
                  
      nTotal = nTotal + nTotalSubAcc
      
      oTextFile.PrintNewLine
      oTextFile.PrintLine Space(95 - nSpcSector) & "----------------------"
      oTextFile.PrintLine Space(65 - nSpcSector) & "SALDO IGUAL AL MAYOR ------> " & SpcFillAmount(nTotal, 22)
      oTextFile.PrintLine Space(95 - nSpcSector) & "======================"
    End With
    oRs.Close
    
    If (bWithSector) Then
      Set oRs = oRSGenerator.GeneraResumen(sAppServer, dCuteDate, sAccount, nGl)
      oTextFile.PrintNewLine
    
      PrintHeaderAuxFec oTextFile, nLine, nPage, dCuteDate, sGL, sAccountName, sRepNum, _
                        sNameA, sPuestoA, sNameB, sPuestoB, bWithSector
     
      oTextFile.PrintLine SpcFill(" ", 6) & "     RESUMEN DE SECTORES      "
      oTextFile.PrintLine SpcFill(" ", 6) & " SECTOR               IMPORTE "
      oTextFile.PrintLine SpcFill(" ", 6) & "------------------------------"

      With oRs
        Do
          If (.EOF) Then Exit Do
          oTextFile.PrintLine SpcFill(" ", 7) & SpcFill(!clave_sector, 6) & SpcFillAmount(!Saldo_Actual, 22)
          nTotSectores = nTotSectores + !Saldo_Actual
          .MoveNext
        Loop
      End With
      oTextFile.PrintLine SpcFill(" ", 6) & "------------------------------"
      oTextFile.PrintLine SpcFill(" ", 13) & SpcFillAmount(nTotSectores, 22)
    ElseIf Not (bWithSector) Then
      '
    End If
    PrintFeed oTextFile, sAccount
    'nTotalconLetra = CantidadConLetra(nTotal)
    PrintSigns oTextFile, sNameA, sPuestoA, sNameB, sPuestoB, _
                    Len("=====================================================================================================================")
    'Cuaderno_7011 = oTextFile.FileName
    'oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Cuaderno_7011", Err.Number
End Function

Private Sub PrintHeaderAuxFec(oTextFile As CTextFile, _
                              nLine As Long, nPage As Long, _
                              dCuteDate As Date, _
                              sGL As String, sAccountName As String, _
                              sRepNum As String, _
                              sNameA As String, sPuestoA As String, _
                              sNameB As String, sPuestoB As String, _
                              Optional bWithSector As Boolean = False)
  Dim sDate       As String
  '**********************************************************
  On Error GoTo ErrHandler
    If (nLine >= cnPageLines) Then
      If nLine <> 100 Then
         PrintSigns oTextFile, sNameA, sPuestoA, sNameB, sPuestoB, _
                    Len("=====================================================================================================================")
         'oTextFile.PageFeed (nLine - 5)
      End If
      'sDate = " AL " & Day(dCuteDate) & " DE " & UCase(MonthName(Month(dCuteDate))) & " DE " & Year(dCuteDate)
      sDate = " AL " & Format(dCuteDate, "d"" DE ""MMMM"" DE ""yyyy")
      oTextFile.PrintLine "TECOYOTITLA NUM 100       BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C.             FECHA: " & Date
      oTextFile.PrintLine "COLONIA FLORIDA            " & SpcFill(sAccountName & sDate, 64) & "PAGINA:   " & nPage
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F.    ============================================================"
      oTextFile.PrintLine "BNO-670315-CDO"
      oTextFile.PrintLine sRepNum
      oTextFile.PrintLine "                                  DELEGACION   : " & sGL
      If (bWithSector) Then
        oTextFile.PrintLine "====================================================================================================================="
        oTextFile.PrintLine "  FECHA    AUXILIAR           SECTOR DESCRIPCION                                        IMPORTE                TOTAL "
        oTextFile.PrintLine "---------------------------------------------------------------------------------------------------------------------"
      ElseIf Not (bWithSector) Then
        oTextFile.PrintLine "================================================================================================================"
        oTextFile.PrintLine "  FECHA    AUXILIAR             DESCRIPCION                                        IMPORTE                TOTAL "
        oTextFile.PrintLine "----------------------------------------------------------------------------------------------------------------"
      End If
      'oTextFile.PrintLine
      'nLine = 10
      nLine = 20
      nPage = nPage + 1
    End If
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintHeaderAuxFec", Err.Number
End Sub

Private Function Cuaderno_7010(oTextFile As CTextFile, sAppServer As String, dCuteDate As Date, _
                                sAccount As String, nGl As Long, sRepNum As String, _
                                sNameA As String, sPuestoA As String, _
                                sNameB As String, sPuestoB As String, _
                                Optional bWithSector As Boolean = False, _
                                Optional bWithResumen As Boolean = False) As String
  Dim oRs           As ADODB.Recordset
  Dim oRSGenerator  As New SCFFixedReportsDS.CNotebookData
  'Dim oTextFile      As New CTextFile
  
  Dim sPrintLine    As String
  Dim sPreviousAcc  As String
  Dim sSector       As String
  Dim sSubAccount   As String
  Dim sCompleteAcc  As String
  Dim nLine         As Long
  Dim nPage         As Long
  Dim nTotalSubAcc  As Currency
  Dim nTotal        As Currency
  Dim nTotSectores  As Currency
  Dim nSpcSector    As Long
  
  Dim sGL           As String
  Dim sAccountName  As String
  Dim nTotalconLetra As String
  '********************************************************************************
  On Error GoTo ErrHandler
    'oTextFile.GenerateRandomFile
    'oTextFile.Open_
    'oTextFile.Open_ "c:\desarrollos\cuaderno\", "Cuaderno7010.txt"
    Set oRs = oRSGenerator.GetRS_7010(sAppServer, nGl, dCuteDate, sAccount)
    sGL = oRSGenerator.GetObjName(sAppServer, nGl, "COF_MAYOR", "NOMBRE_MAYOR", "ID_MAYOR")
    sAccountName = sAccount & " - " & oRSGenerator.GetAccountName(sAppServer, sAccount, nGl)
   
    nTotal = 0
    nTotalSubAcc = 0
    nLine = 100
    nPage = 1
    
    If ((oRs.BOF) And (oRs.EOF)) Then
      PrintHeader7010 oTextFile, nLine, nPage, dCuteDate, sGL, sAccountName, sRepNum, _
                      sNameA, sPuestoA, sNameB, sPuestoB, bWithSector
      oTextFile.PrintLine " NO EXISTE INFORMACION "
      Exit Function
    End If
    
    With oRs
      sPreviousAcc = !numero_cuenta_estandar
      sCompleteAcc = GetCuentaSubCuenta(!numero_cuenta_estandar, , sSubAccount)
      sCompleteAcc = sCompleteAcc & !clave_sector
      sSector = !clave_sector
      If (bWithSector) Then
        nSpcSector = 0
      ElseIf Not (bWithSector) Then
        nSpcSector = 5
      End If
      
      Do While Not .EOF()
        If (!numero_cuenta_estandar <> sPreviousAcc) Or (sSector <> !clave_sector) Then
          PrintHeader7010 oTextFile, nLine, nPage, dCuteDate, sGL, sAccountName, sRepNum, _
                          sNameA, sPuestoA, sNameB, sPuestoB, bWithSector
          oTextFile.PrintLine Space(51 - nSpcSector) & _
                              SpcFill(sCompleteAcc, 21) & " " & _
                              SpcFill(" ", 4) & Space(17) & _
                              SpcFillAmount(nTotalSubAcc, 22)
          oTextFile.PrintLine ""
          nLine = nLine + 2
          nTotal = nTotal + nTotalSubAcc
          nTotalSubAcc = 0
        End If
        
        PrintHeader7010 oTextFile, nLine, nPage, dCuteDate, sGL, sAccountName, sRepNum, _
                        sNameA, sPuestoA, sNameB, sPuestoB, bWithSector
        
        If (bWithSector) Then
          sPrintLine = SpcFill(IIf(IsNull(!FECHA_REGISTRO), " ", !FECHA_REGISTRO), 10) & " " & _
                       SpcFill(!NUMERO_CUENTA_AUXILIAR, 20) & " " & _
                       SpcFill(!clave_sector, 4) & " " & _
                       SpcFill(!nombre_cuenta_auxiliar, 35) & " " & _
                       SpcFillAmount(!Saldo_Actual, 22)
                       
        ElseIf Not (bWithSector) Then
          sPrintLine = SpcFill(IIf(IsNull(!FECHA_REGISTRO), " ", !FECHA_REGISTRO), 10) & " " & _
                       SpcFill(IIf(IsNull(!NUMERO_CUENTA_AUXILIAR), " ", !NUMERO_CUENTA_AUXILIAR), 20) & " " & _
                       SpcFill(IIf(IsNull(!nombre_cuenta_auxiliar), " ", !nombre_cuenta_auxiliar), 35) & " " & _
                       SpcFillAmount(!Saldo_Actual, 22)
        End If
        
        oTextFile.PrintLine sPrintLine
        nLine = nLine + 1
        nTotalSubAcc = nTotalSubAcc + !Saldo_Actual
        sPreviousAcc = !numero_cuenta_estandar
        sSector = !clave_sector
        sCompleteAcc = GetCuentaSubCuenta(!numero_cuenta_estandar, , sSubAccount)
        sSector = !clave_sector
        sCompleteAcc = sCompleteAcc & !clave_sector
        
        .MoveNext
      Loop
      oTextFile.PrintLine Space(51 - nSpcSector) & _
                          SpcFill(sCompleteAcc, 21) & " " & _
                          SpcFill(" ", 4) & Space(17) & _
                          SpcFillAmount(nTotalSubAcc, 22)
                  
      nTotal = nTotal + nTotalSubAcc
      
      oTextFile.PrintNewLine
      oTextFile.PrintLine Space(95 - nSpcSector) & "----------------------"
      oTextFile.PrintLine Space(65 - nSpcSector) & "SALDO IGUAL AL MAYOR ------> " & SpcFillAmount(nTotal, 22)
      oTextFile.PrintLine Space(95 - nSpcSector) & "======================"
    End With
    oRs.Close
    
    If (bWithResumen) Then
      Set oRs = oRSGenerator.GeneraResumen(sAppServer, dCuteDate, sAccount, nGl)
      oTextFile.PrintNewLine
    
      PrintHeader7010 oTextFile, nLine, nPage, dCuteDate, sGL, sAccountName, sRepNum, _
                      sNameA, sPuestoA, sNameB, sPuestoB, bWithSector
     
      oTextFile.PrintLine SpcFill(" ", 6) & "     RESUMEN DE SECTORES      "
      oTextFile.PrintLine SpcFill(" ", 6) & " SECTOR               IMPORTE "
      oTextFile.PrintLine SpcFill(" ", 6) & "------------------------------"

      With oRs
        Do
          If (.EOF) Then Exit Do
          oTextFile.PrintLine SpcFill(" ", 7) & SpcFill(!clave_sector, 6) & SpcFillAmount(!Saldo_Actual, 22)
          nTotSectores = nTotSectores + !Saldo_Actual
          .MoveNext
        Loop
      End With
      oTextFile.PrintLine SpcFill(" ", 6) & "------------------------------"
      oTextFile.PrintLine SpcFill(" ", 13) & SpcFillAmount(nTotSectores, 22)
    ElseIf Not (bWithResumen) Then
      '
    End If
    PrintFeed oTextFile, sAccount
    'nTotalconLetra = CantidadConLetra(nTotal)
    PrintSigns oTextFile, sNameA, sPuestoA, sNameB, sPuestoB, _
                    Len("================================================================================================================")
    'oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Cuaderno_7010", Err.Number
End Function

Private Sub PrintHeader7010(oTextFile As CTextFile, _
                              nLine As Long, nPage As Long, _
                              dCuteDate As Date, _
                              sGL As String, sAccountName As String, _
                              sRepNum As String, _
                              sNameA As String, sPuestoA As String, sNameB As String, sPuestoB As String, _
                              Optional bWithSector As Boolean = False)
  Dim sDate       As String
  '**********************************************************
  On Error GoTo ErrHandler
    If (nLine >= cnPageLines) Then
      If nLine <> 100 Then
        PrintSigns oTextFile, sNameA, sPuestoA, sNameB, sPuestoB, _
                    Len("================================================================================================================")
         'oTextFile.PageFeed (nLine - 5)
      End If
      
      'sDate = " AL " & Day(dCuteDate) & " DE " & UCase(MonthName(Month(dCuteDate))) & " DE " & Year(dCuteDate)
      sDate = " AL " & Format(dCuteDate, "d"" DE ""MMMM"" DE ""yyyy")
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine "TECOYOTITLA NUM 100             BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C.                FECHA: " & Date
      oTextFile.PrintLine "COLONIA FLORIDA                  " & SpcFill(sAccountName & sDate, 60) & "        PAGINA:   " & nPage
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F.          ===================================================="
      oTextFile.PrintLine "BNO-670315-CDO"
      oTextFile.PrintLine sRepNum
      oTextFile.PrintLine "                                  DELEGACION   : " & sGL
      If (bWithSector) Then
        oTextFile.PrintLine "====================================================================================================================="
        oTextFile.PrintLine "  FECHA    AUXILIAR           SECTOR DESCRIPCION                                        IMPORTE                TOTAL "
        oTextFile.PrintLine "---------------------------------------------------------------------------------------------------------------------"
      ElseIf Not (bWithSector) Then
        oTextFile.PrintLine "================================================================================================================"
        oTextFile.PrintLine "  FECHA    AUXILIAR             DESCRIPCION                                        IMPORTE                TOTAL "
        oTextFile.PrintLine "----------------------------------------------------------------------------------------------------------------"
      End If
      'oTextFile.PrintLine
      nLine = 24
      'nLine = 14
      nPage = nPage + 1
    End If
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintHeader7010", Err.Number
End Sub

Private Function Cuaderno_7013(oTextFile As CTextFile, sAppServer As String, dCuteDate As Date, _
                              sAccount As String, nGl As Long, sRepNum As String, _
                              sNameA As String, sPuestoA As String, _
                              sNameB As String, sPuestoB As String, _
                              Optional bWithSector As Boolean = False) As String
  Dim oRs           As ADODB.Recordset
  Dim oRSGenerator  As New SCFFixedReportsDS.CNotebookData
  'Dim oTextFile      As New CTextFile
  
  Dim sPrintLine    As String
  Dim sPreviousAcc  As String
  Dim sSector       As String
  Dim sCompleteAcc  As String
  Dim nLine         As Long
  Dim nPage         As Long
  Dim nTotalSubAcc  As Currency
  Dim nTotal        As Currency
  Dim nTotSectores  As Currency
  Dim nSpcSector    As Long
  Dim sAccNameRs    As String
  
  Dim sGL           As String
  Dim sAccountName  As String
  Dim nTotalconLetra As String
  '********************************************************************************
  On Error GoTo ErrHandler
    'oTextFile.GenerateRandomFile
    'oTextFile.Open_
    'oTextFile.Open_ "c:\desarrollos\cuaderno\", "Cuaderno7013.txt"
    Set oRs = oRSGenerator.GeneraAuxiliares(sAppServer, dCuteDate, sAccount, nGl)
    sGL = oRSGenerator.GetObjName(sAppServer, nGl, "COF_MAYOR", "NOMBRE_MAYOR", "ID_MAYOR")
    sAccountName = sAccount & " - " & oRSGenerator.GetAccountName(sAppServer, sAccount, nGl)
       
    nTotal = 0
    nTotalSubAcc = 0
    nLine = 100
    nPage = 1
    
    If ((oRs.BOF) And (oRs.EOF)) Then
      PrintHeader7013 oTextFile, nLine, nPage, dCuteDate, sGL, sAccountName, "7013", _
                      sNameA, sPuestoA, sNameB, sPuestoB
      oTextFile.PrintLine " NO EXISTE INFORMACION "
      Exit Function
    End If
    
    With oRs
      sPreviousAcc = !numero_cuenta_estandar
      sCompleteAcc = GetCuentaSubCuenta(!numero_cuenta_estandar)
      sCompleteAcc = sCompleteAcc & !clave_sector
      sSector = !clave_sector
      sAccNameRs = !nombre_cuenta_estandar
      Do While Not .EOF()
        If (!numero_cuenta_estandar <> sPreviousAcc) Or (sSector <> !clave_sector) Then
          PrintHeader7013 oTextFile, nLine, nPage, dCuteDate, sGL, sAccountName, "7013", _
                          sNameA, sPuestoA, sNameB, sPuestoB
          oTextFile.PrintLine Space(47) & _
                              SpcFill(sCompleteAcc, 21) & " " & _
                              SpcFill(" ", 4) & Space(17) & _
                              SpcFillAmount(nTotalSubAcc, 22)
          oTextFile.PrintLine Space(47) & SpcFill(sAccNameRs, 40)
          oTextFile.PrintNewLine
          nLine = nLine + 2
          nTotal = nTotal + nTotalSubAcc
          nTotalSubAcc = 0
        End If
        PrintHeader7013 oTextFile, nLine, nPage, dCuteDate, sGL, sAccountName, "7013", _
                        sNameA, sPuestoA, sNameB, sPuestoB
        oTextFile.PrintLine SpcFill(!NUMERO_CUENTA_AUXILIAR, 20) & " " & _
                            SpcFill(!fecha_ult, 10) & " " & _
                            SpcFill(!nombre_cuenta_auxiliar, 35) & " " & _
                            SpcFillAmount(!Saldo_Actual, 22)
        oTextFile.PrintLine Space(32) & "SECTOR : " & SpcFill(!clave_sector, 2)
        nLine = nLine + 1
        nTotalSubAcc = nTotalSubAcc + !Saldo_Actual
        sPreviousAcc = !numero_cuenta_estandar
        sSector = !clave_sector
        sAccNameRs = !nombre_cuenta_estandar
        sCompleteAcc = GetCuentaSubCuenta(!numero_cuenta_estandar)
        sCompleteAcc = sCompleteAcc & !clave_sector
        .MoveNext
      Loop
      oTextFile.PrintLine Space(47) & _
                          SpcFill(sCompleteAcc, 21) & " " & _
                          SpcFill(" ", 4) & Space(17) & _
                          SpcFillAmount(nTotalSubAcc, 22)
      oTextFile.PrintLine Space(47) & SpcFill(sAccNameRs, 40)
      nTotal = nTotal + nTotalSubAcc
      
      oTextFile.PrintNewLine
      oTextFile.PrintLine Space(91) & "----------------------"
      oTextFile.PrintLine Space(61) & "SALDO IGUAL AL MAYOR ------> " & SpcFillAmount(nTotal, 22)
      oTextFile.PrintLine Space(91) & "======================"
    End With
    oRs.Close
    Set oRs = oRSGenerator.GeneraResumen(sAppServer, dCuteDate, sAccount, nGl)
    oTextFile.PrintNewLine
    
    PrintHeader7013 oTextFile, nLine, nPage, dCuteDate, sGL, sAccountName, "7013", _
                    sNameA, sPuestoA, sNameB, sPuestoB
    
    oTextFile.PrintLine SpcFill(" ", 6) & "     RESUMEN DE SECTORES      "
    oTextFile.PrintLine SpcFill(" ", 6) & " SECTOR               IMPORTE "
    oTextFile.PrintLine SpcFill(" ", 6) & "------------------------------"

    With oRs
      Do
        If (.EOF) Then Exit Do
        oTextFile.PrintLine SpcFill(" ", 7) & SpcFill(!clave_sector, 6) & SpcFillAmount(!Saldo_Actual, 22)
        nTotSectores = nTotSectores + !Saldo_Actual
        .MoveNext
      Loop
    End With
    oTextFile.PrintLine SpcFill(" ", 6) & "------------------------------"
    oTextFile.PrintLine SpcFill(" ", 13) & SpcFillAmount(nTotSectores, 22)
    PrintFeed oTextFile, sAccount
    'nTotalconLetra = CantidadConLetra(nTotal)
    PrintSigns oTextFile, sNameA, sPuestoA, sNameB, sPuestoB, _
                    Len("=================================================================================================================")
    'oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Cuaderno_7013", Err.Number
End Function

Private Sub PrintHeader7013(oTextFile As CTextFile, _
                            nLine As Long, nPage As Long, _
                            dCuteDate As Date, _
                            sGL As String, sAccountName As String, _
                            sRepNum As String, _
                            sNameA As String, sPuestoA As String, _
                            sNameB As String, sPuestoB As String)
  Dim sDate       As String
  '**********************************************************
  On Error GoTo ErrHandler
    If (nLine >= cnPageLines) Then
      If nLine <> 100 Then
         PrintSigns oTextFile, sNameA, sPuestoA, sNameB, sPuestoB, _
                    Len("=================================================================================================================")
         ' oTextFile.PageFeed (nLine - 5)
      End If
      'sDate = " AL " & Day(dCuteDate) & " DE " & UCase(MonthName(Month(dCuteDate))) & " DE " & Year(dCuteDate)
      sDate = " AL " & Format(dCuteDate, "d"" DE ""MMMM"" DE ""yyyy")
      
      oTextFile.PrintLine "TECOYOTITLA NUM 100             BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C.      FECHA: " & Date
      oTextFile.PrintLine "COLONIA FLORIDA                  " & SpcFill(sAccountName & sDate, 57) & "PAGINA:   " & nPage
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F.          ===================================================="
      oTextFile.PrintLine "BNO-670315-CDO"
      oTextFile.PrintLine sRepNum
      oTextFile.PrintLine "                                  DELEGACION   : " & sGL
      oTextFile.PrintLine "================================================================================================================="
      oTextFile.PrintLine " NUM. DE CREDITO    FECHA REEST NOMBRE DEL ACREEDITADO                          IMPORTE                 TOTAL    "
      oTextFile.PrintLine "-----------------------------------------------------------------------------------------------------------------"
      nLine = 20
      'nLine = 10
      nPage = nPage + 1
    End If
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintHeader7013", Err.Number
End Sub

Private Function Cuaderno_7006(oTextFile As CTextFile, sAppServer As String, _
                              dCuteDate As Date, dExcDate As Date, _
                              vAccounts As Variant, nGl As Long, _
                              sNameA As String, sPuestoA As String, _
                              sNameB As String, sPuestoB As String, _
                              Optional nExcRateType As Long = 46, _
                              Optional nExcRateCurrency As Long = 1) As String
  Dim oRs           As ADODB.Recordset
  Dim oRSGenerator  As New SCFFixedReportsDS.CNotebookData
  'Dim oTextFile     As New CTextFile
                                                          
  Dim sAuxPrint     As String
  Dim sAuxName      As String
  Dim nPrintAcc(17) As Currency
  Dim nTotAcc(17)   As Currency
  Dim nSumaAux      As Currency
  Dim nSumaTot      As Currency
  Dim sGL           As String
  
  Dim nLine         As Long
  Dim nPage         As Long
  Dim i             As Long
  Dim nTotalconLetra As String
  '***************************************************************************************
  On Error GoTo ErrHandler
    'oTextFile.GenerateRandomFile
    'oTextFile.Open_
    '"C:\desarrollos\cuaderno\", "Cuaderno7006.txt"
    'Set oRs = oRSGenerator.GetRS_R7006(sAppServer, "31/08/2000", "31/08/2000", Array("1305", "1307", "1308", "1310", "1311", "1314", "1315", "1316", "1317", "1319", "1320", "1323", "1324", "1337", "1350", "6378"), 32)
    Set oRs = oRSGenerator.GetRS_R7006(sAppServer, dCuteDate, dExcDate, vAccounts, nGl, _
                                       nExcRateType, nExcRateCurrency)
    sGL = oRSGenerator.GetObjName(sAppServer, nGl, "COF_MAYOR", "NOMBRE_MAYOR", "ID_MAYOR")
    nLine = 100
    nPage = 1
    With oRs
      sAuxPrint = !NUMERO_CUENTA_AUXILIAR
      sAuxName = !nombre_cuenta_auxiliar
      Do While Not .EOF()
        If (sAuxPrint <> !NUMERO_CUENTA_AUXILIAR) Then
          PrintHeader7006 oTextFile, nLine, nPage, dCuteDate, sGL, "7006", _
                          sNameA, sPuestoA, sNameB, sPuestoB
                            
          oTextFile.PrintLine " " & sAuxPrint & "  " & sAuxName
          oTextFile.PrintLine SpcFillAmount(nPrintAcc(0), 17) & " " & SpcFillAmount(nPrintAcc(1), 17) & " " & _
                              SpcFillAmount(nPrintAcc(2), 17) & " " & SpcFillAmount(nPrintAcc(3), 17) & " " & _
                              SpcFillAmount(nPrintAcc(4), 17) & " " & SpcFillAmount(nPrintAcc(5), 17) & " " & _
                              SpcFillAmount(nPrintAcc(6), 17) & " " & SpcFillAmount(nPrintAcc(7), 17) & " " & _
                              SpcFillAmount(nPrintAcc(8), 17)
          oTextFile.PrintLine SpcFillAmount(nPrintAcc(9), 17) & " " & SpcFillAmount(nPrintAcc(10), 17) & " " & _
                              SpcFillAmount(nPrintAcc(11), 17) & " " & SpcFillAmount(nPrintAcc(12), 17) & " " & _
                              SpcFillAmount(nPrintAcc(13), 17) & " " & SpcFillAmount(nPrintAcc(14), 17) & " " & _
                              SpcFillAmount(nPrintAcc(15), 17) & " " & SpcFillAmount(nPrintAcc(16), 17) & " " & _
                              SpcFillAmount(nSumaAux, 17)
          nLine = nLine + 3
          nSumaAux = 0
          For i = LBound(vAccounts) To UBound(vAccounts)
            nPrintAcc(i) = 0
          Next i
        End If
        Select Case !numero_cuenta_estandar
             'Case "1305"
             Case vAccounts(0)
                        nPrintAcc(0) = !Saldo_Actual
                        nTotAcc(0) = nTotAcc(0) = !Saldo_Actual
             'Case "1307"
             Case vAccounts(1)
                        nPrintAcc(1) = !Saldo_Actual
                        nTotAcc(1) = nTotAcc(1) + !Saldo_Actual
             'Case "1308"
             Case vAccounts(2)
                        nPrintAcc(2) = !Saldo_Actual
                        nTotAcc(2) = nTotAcc(2) + !Saldo_Actual
             'Case "1310"
             Case vAccounts(3)
                        nPrintAcc(3) = !Saldo_Actual
                        nTotAcc(3) = nTotAcc(3) + !Saldo_Actual
             'Case "1311"
             Case vAccounts(4)
                        nPrintAcc(4) = !Saldo_Actual
                        nTotAcc(4) = nTotAcc(4) + !Saldo_Actual
             'Case "1312"
             Case vAccounts(5)
                        nPrintAcc(5) = !Saldo_Actual
                        nTotAcc(5) = nTotAcc(5) + !Saldo_Actual
             'Case "1314"
             Case vAccounts(6)
                        nPrintAcc(6) = !Saldo_Actual
                        nTotAcc(6) = nTotAcc(6) + !Saldo_Actual
             'Case "1315"
             Case vAccounts(7)
                        nPrintAcc(7) = !Saldo_Actual
                        nTotAcc(7) = nTotAcc(7) + !Saldo_Actual
             'Case "1316"
             Case vAccounts(8)
                        nPrintAcc(8) = !Saldo_Actual
                        nTotAcc(8) = nTotAcc(8) + !Saldo_Actual
             'Case "1317"
             Case vAccounts(9)
                        nPrintAcc(9) = !Saldo_Actual
                        nTotAcc(9) = nTotAcc(9) + !Saldo_Actual
             'Case "1319"
             Case vAccounts(10)
                        nPrintAcc(10) = !Saldo_Actual
                        nTotAcc(10) = nTotAcc(10) + !Saldo_Actual
             'Case "1320"
             Case vAccounts(11)
                        nPrintAcc(11) = !Saldo_Actual
                        nTotAcc(11) = nTotAcc(11) + !Saldo_Actual
             'Case "1323"
             Case vAccounts(12)
                        nPrintAcc(12) = !Saldo_Actual
                        nTotAcc(12) = nTotAcc(12) + !Saldo_Actual
             'Case "1324"
             Case vAccounts(13)
                        nPrintAcc(13) = !Saldo_Actual
                        nTotAcc(13) = nTotAcc(13) + !Saldo_Actual
             'Case "1337"
             Case vAccounts(14)
                        nPrintAcc(14) = !Saldo_Actual
                        nTotAcc(14) = nTotAcc(14) + !Saldo_Actual
             'Case "1350"
             Case vAccounts(15)
                        nPrintAcc(15) = !Saldo_Actual
                        nTotAcc(15) = nTotAcc(15) + !Saldo_Actual
             ' Case "6378"
             Case vAccounts(16)
                        nPrintAcc(16) = !Saldo_Actual
                        nTotAcc(16) = nTotAcc(16) + !Saldo_Actual
        End Select
        nSumaAux = nSumaAux + !Saldo_Actual
        nSumaTot = nSumaTot + !Saldo_Actual
        sAuxPrint = !NUMERO_CUENTA_AUXILIAR
        sAuxName = !nombre_cuenta_auxiliar
        .MoveNext
      Loop
      PrintHeader7006 oTextFile, nLine, nPage, dCuteDate, sGL, "7006", _
                      sNameA, sPuestoA, sNameB, sPuestoB
      oTextFile.PrintLine " " & sAuxPrint & "  " & sAuxName
      oTextFile.PrintLine SpcFillAmount(nPrintAcc(0), 17) & " " & SpcFillAmount(nPrintAcc(1), 17) & " " & _
                          SpcFillAmount(nPrintAcc(2), 17) & " " & SpcFillAmount(nPrintAcc(3), 17) & " " & _
                          SpcFillAmount(nPrintAcc(4), 17) & " " & SpcFillAmount(nPrintAcc(5), 17) & " " & _
                          SpcFillAmount(nPrintAcc(6), 17) & " " & SpcFillAmount(nPrintAcc(7), 17) & " " & _
                          SpcFillAmount(nPrintAcc(8), 17)
      oTextFile.PrintLine SpcFillAmount(nPrintAcc(9), 17) & " " & SpcFillAmount(nPrintAcc(10), 17) & " " & _
                          SpcFillAmount(nPrintAcc(11), 17) & " " & SpcFillAmount(nPrintAcc(12), 17) & " " & _
                          SpcFillAmount(nPrintAcc(12), 17) & " " & SpcFillAmount(nPrintAcc(14), 17) & " " & _
                          SpcFillAmount(nPrintAcc(15), 17) & " " & SpcFillAmount(nPrintAcc(16), 17) & " " & _
                          SpcFillAmount(nSumaAux, 17)
      
      oTextFile.PrintLine "==========================================================================================================================================================================="
      oTextFile.PrintLine SpcFillAmount(nTotAcc(0), 17) & " " & SpcFillAmount(nTotAcc(1), 17) & " " & _
                          SpcFillAmount(nTotAcc(2), 17) & " " & SpcFillAmount(nTotAcc(3), 17) & " " & _
                          SpcFillAmount(nTotAcc(4), 17) & " " & SpcFillAmount(nTotAcc(5), 17) & " " & _
                          SpcFillAmount(nTotAcc(6), 17) & " " & SpcFillAmount(nTotAcc(7), 17) & " " & _
                          SpcFillAmount(nTotAcc(8), 17)
      oTextFile.PrintLine SpcFillAmount(nTotAcc(9), 17) & " " & SpcFillAmount(nTotAcc(10), 17) & " " & _
                          SpcFillAmount(nTotAcc(11), 17) & " " & SpcFillAmount(nTotAcc(12), 17) & " " & _
                          SpcFillAmount(nTotAcc(13), 17) & " " & SpcFillAmount(nTotAcc(14), 17) & " " & _
                          SpcFillAmount(nTotAcc(15), 17) & " " & SpcFillAmount(nTotAcc(16), 17) & " " & _
                          SpcFillAmount(nSumaTot, 17)
      oTextFile.PrintLine "==========================================================================================================================================================================="
    End With
    
    PrintSigns oTextFile, sNameA, sPuestoA, sNameB, sPuestoB, _
                   Len("===========================================================================================================================================================================")
    'oTextFile.Close_
    Exit Function
ErrHandler:
  RaiseError ClassId, "Cuaderno_7006", Err.Number
End Function

Private Sub PrintHeader7006(oTextFile As CTextFile, _
                            nLine As Long, nPage As Long, _
                            dCuteDate As Date, _
                            sGL As String, _
                            sRepNum As String, _
                            sNameA As String, sPuestoA As String, _
                            sNameB As String, sPuestoB As String)
  Dim sDate       As String
  '**********************************************************
  On Error GoTo ErrHandler
    If (nLine >= cnPageLines - 3) Then
      If nLine <> 100 Then
        PrintSigns oTextFile, sNameA, sPuestoA, sNameB, sPuestoB, _
                   Len("===========================================================================================================================================================================")
        'oTextFile.PageFeed (nLine - 5)
      End If
      'sDate = " AL " & Day(dCuteDate) & " DE " & UCase(MonthName(Month(dCuteDate))) & " DE " & Year(dCuteDate)
      sDate = " AL " & Format(dCuteDate, "d"" DE ""MMMM"" DE ""yyyy")
      oTextFile.PrintLine "TECOYOTITLA NUM 100                                      BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C.                               FECHA: " & Date
      oTextFile.PrintLine "COLONIA FLORIDA                            " & SpcFill("RELACION DE RESPONSABILIDADES DATOS EN MONEDA NACIONAL" & sDate, 97) & "PAGINA:   " & nPage
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F.                    =================================================================================="
      oTextFile.PrintLine "BNO-670315-CDO"
      oTextFile.PrintLine sRepNum
      oTextFile.PrintLine "                                  DELEGACION   : " & sGL
      oTextFile.PrintLine "==========================================================================================================================================================================="
      oTextFile.PrintLine "No. DE CREDITO  NOMBRE DEL ACREDITADO"
      oTextFile.PrintLine "             1305               1307               1308               1310               1311               1312               1314               1315               1316  "
      oTextFile.PrintLine "             1317               1319               1320               1323               1324               1337               1350               6378           T O T A L "
      oTextFile.PrintLine "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
      'nLine = 12
      nLine = 22
      nPage = nPage + 1
    End If
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintHeader7006", Err.Number
End Sub

Private Function REL_ANALITICA_CUENTAS(oTextFile As CTextFile, sAppServer As String, nGl As Long, _
                                  dIniDate As Date, dEndDate As Date, _
                                  sNameA As String, sPuestoA As String, _
                                  sNameB As String, sPuestoB As String) As String
  Dim oRs           As ADODB.Recordset
  Dim oRSGenerator  As New SCFFixedReportsDS.CNotebookData
  'Dim oTextFile     As New CTextFile
  
  Dim sAccount      As String
  Dim sSubAccount   As String
  Dim sSubSubAcc    As String
  Dim nNumSpc       As Currency
  
  Dim sGL           As String
  Dim nPage         As Long
  Dim nLine         As Long
  Dim sText         As String
  Dim nTotGrp51     As Currency
  Dim nTotGrp52     As Currency
  Dim nDiferencia   As Currency
  Dim nTotalconLetra As String
  '********************************************************************************
  On Error GoTo ErrHandler
    'oTextFile.GenerateRandomFile
    'oTextFile.Open_
    'oTextFile.Open_ "c:\DEV2001\cuaderno\", "Cuaderno5.txt"
    Set oRs = oRSGenerator.GeneraRSAnalitico(sAppServer, "&&&&-&&-&&", nGl, dIniDate, dEndDate)
    'Set oRs = oRSGenerator.GeneraRSPRUEBA("&&&&-&&-&&", sAppServer, nGL, dIniDate, dEndDate, False)
    sGL = oRSGenerator.GetObjName(sAppServer, nGl, "COF_MAYOR", "NOMBRE_MAYOR", "ID_MAYOR")
    nLine = 100
    nPage = 1
    With oRs
      Do While Not (.EOF())
        DivideAcc !numero_cuenta_estandar, sAccount, sSubAccount, sSubSubAcc
        If (Len(sSubSubAcc) <> 0) Then
          nNumSpc = 61
        ElseIf (Len(sSubSubAcc) = 0) And (Len(sSubAccount) <> 0) Then
          nNumSpc = 33
        ElseIf (Len(sSubSubAcc) = 0) And (Len(sSubAccount) = 0) Then
          nNumSpc = 6
          If (Mid(!numero_cuenta_estandar, 1, 2) = "51") Then
            nTotGrp51 = nTotGrp51 + !Saldo_Actual
          ElseIf (Mid(!numero_cuenta_estandar, 1, 2) = "52") Then
            nTotGrp52 = nTotGrp52 + !Saldo_Actual
          End If
        End If
'        If ((sSubAccount <> "00") And Len(sSubSubAcc) = 0) Or (sSubSubAcc <> "00") Then
        If (sSubSubAcc <> "00") Then
          PrintHeader_RAC oTextFile, nLine, nPage, dEndDate, sGL, "RACB5", _
                          sNameA, sPuestoA, sNameB, sPuestoB
          oTextFile.PrintLine Space(1) & _
                              SpcFill(sAccount, 5) & Space(2) & _
                              SpcFill(sSubAccount, 2) & Space(4) & _
                              SpcFill(sSubSubAcc, 2) & Space(1) & _
                              Space(nNumSpc) & _
                              SpcFillAmount(!Saldo_Actual, 22)
                              '& !CLAVE_SECTOR
'        ElseIf ((sSubAccount = "00") And Len(sSubSubAcc) = 0) Or (sSubSubAcc = "00") Then
          ' siguiente
        End If
        sAccount = ""
        sSubAccount = ""
        sSubSubAcc = ""
        nLine = nLine + 1
        .MoveNext
      Loop
      PrintHeader_RAC oTextFile, nLine, nPage, dEndDate, sGL, "RACB5", _
                      sNameA, sPuestoA, sNameB, sPuestoB
      If ((nTotGrp52 - nTotGrp51) < 0) Then
        sText = "PERDIDA  $"
        nDiferencia = nTotGrp51 - nTotGrp52
      ElseIf ((nTotGrp52 - nTotGrp51) >= 0) Then
        sText = "UTILIDAD $"
        nDiferencia = nTotGrp52 - nTotGrp51
      End If
      nLine = nLine + 7
      PrintHeader_RAC oTextFile, nLine, nPage, dEndDate, sGL, "RACB5", _
                      sNameA, sPuestoA, sNameB, sPuestoB
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine Space(6) & _
                          "TOTAL GRUPO 52 $" & Space(1) & _
                          SpcFillAmount(nTotGrp52, 22) & Space(18) & sText & SpcFillAmount(nDiferencia, 22)
      oTextFile.PrintLine Space(6) & _
                          "TOTAL GRUPO 51 $" & Space(1) & _
                          SpcFillAmount(nTotGrp51, 22)
      oTextFile.PrintLine Space(20) & "-------------------------"
      oTextFile.PrintLine Space(21) & "$" & Space(1) & _
                          SpcFillAmount(nTotGrp52 - nTotGrp51, 22)
    End With
    'nTotalconLetra = CantidadConLetra(nTotGrp52 - nTotGrp51)
    PrintSigns oTextFile, sNameA, sPuestoA, sNameB, sPuestoB, _
                   Len("---------------------------------------------------------------------------------------------------------")
    'oTextFile.Close_
    Exit Function
ErrHandler:
  RaiseError ClassId, "REL_ANALITICA_CUENTAS", Err.Number
End Function

Private Sub PrintHeader_RAC(oTextFile As CTextFile, _
                            nLine As Long, nPage As Long, _
                            dCuteDate As Date, _
                            sGL As String, _
                            sRepNum As String, _
                            sNameA As String, sPuestoA As String, _
                            sNameB As String, sPuestoB As String)
 Dim sDate     As String
 '*********************************************************
 On Error GoTo ErrHandler
    If (nLine >= cnPageLines - 3) Then
      If nLine <> 100 Then
        PrintSigns oTextFile, sNameA, sPuestoA, sNameB, sPuestoB, _
                   Len("---------------------------------------------------------------------------------------------------------")
        'oTextFile.PageFeed (nLine - 5)
      End If
      'sDate = " AL " & Day(dCuteDate) & " DE " & UCase(MonthName(Month(dCuteDate))) & " DE " & Year(dCuteDate)
      sDate = " AL " & Format(dCuteDate, "d"" DE ""MMMM"" DE ""yyyy")
      oTextFile.PrintLine "TECOYOTITLA NUM 100       BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C.      FECHA: " & Date
      oTextFile.PrintLine "COLONIA FLORIDA                RELACION ANALITICA DE CUENTAS DE RESULTADOS          PAGINA: " & nPage
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F.        CON SALDOS ACUMULADOS" & sDate
      oTextFile.PrintLine "BNO-670315-CDO"
      'oTextFile.PrintLine sRepNum
      oTextFile.PrintLine "                                  DELEGACION   : " & sGL
      oTextFile.PrintLine ""
      oTextFile.PrintLine " CUENTA SUB  SUB                   SALDO                    SALDO                      IMPORTE     "
      oTextFile.PrintLine "        CTA  SUB CTA               CUENTA                 SUB CUENTA                SUB SUB CUENTA "
      oTextFile.PrintLine "---------------------------------------------------------------------------------------------------------"
      'nLine = 10
      nLine = 20
      nPage = nPage + 1
    End If
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintHeader_RAC", Err.Number
End Sub

Private Function Cuaderno_7014(oTextFile As CTextFile, sAppServer As String, dCuteDate As Date, _
                              sAccount As String, nGl As Long, sRepNum As String, _
                              sNameA As String, sPuestoA As String, _
                              sNameB As String, sPuestoB As String, _
                              Optional bWithSector As Boolean = False, _
                              Optional bWithResumen As Boolean = False) As String
  Dim oRs           As ADODB.Recordset
  Dim oRSGenerator  As New SCFFixedReportsDS.CNotebookData
  'Dim oTextFile      As New CTextFile
  
  Dim sPrintLine    As String
  Dim sPreviousAcc  As String
  Dim sSector       As String
  Dim nLine         As Long
  Dim nPage         As Long
  Dim nTotalSubAcc  As Currency
  Dim nTotal        As Currency
  Dim nTotSectores  As Currency
  Dim nSpcSector    As Long
  
  Dim sCompleteAcc  As String
  Dim sGL           As String
  Dim sAccountName  As String
  Dim nTotalconLetra As String
  '********************************************************************************
  On Error GoTo ErrHandler
    'oTextFile.GenerateRandomFile
    'oTextFile.Open_
    'oTextFile.Open_ "c:\desarrollos\cuaderno\", "Cuaderno7014.txt"
    Set oRs = oRSGenerator.GetRS_7014(sAppServer, nGl, dCuteDate, sAccount)
    sGL = oRSGenerator.GetObjName(sAppServer, nGl, "COF_MAYOR", "NOMBRE_MAYOR", "ID_MAYOR")
    sAccountName = sAccount & " - " & oRSGenerator.GetAccountName(sAppServer, sAccount, nGl)
   
    nTotal = 0
    nTotalSubAcc = 0
    nLine = 100
    nPage = 1
    
    If ((oRs.BOF) And (oRs.EOF)) Then
      PrintHeader7014 oTextFile, nLine, nPage, dCuteDate, sGL, sAccountName, sRepNum, _
                                 sNameA, sPuestoA, sNameB, sPuestoB, bWithSector
      oTextFile.PrintLine " NO EXISTE INFORMACION "
      Exit Function
    End If
    
    With oRs
      sPreviousAcc = !numero_cuenta_estandar
      sSector = !clave_sector
      sCompleteAcc = GetCuentaSubCuenta(!numero_cuenta_estandar)
      sCompleteAcc = sCompleteAcc & !clave_sector
      
      If (bWithSector) Then
        nSpcSector = 0
      ElseIf Not (bWithSector) Then
        nSpcSector = 5
      End If
      
      Do While Not .EOF()
        If (!numero_cuenta_estandar <> sPreviousAcc) Or (sSector <> !clave_sector) Then
          PrintHeader7014 oTextFile, nLine, nPage, dCuteDate, sGL, sAccountName, sRepNum, _
                          sNameA, sPuestoA, sNameB, sPuestoB, bWithSector
          oTextFile.PrintLine Space(51 - nSpcSector) & _
                              SpcFill(sCompleteAcc, 21) & " " & _
                              SpcFill(" ", 4) & Space(17) & _
                              SpcFillAmount(nTotalSubAcc, 22)
          oTextFile.PrintLine ""
          nLine = nLine + 2
          nTotal = nTotal + nTotalSubAcc
          nTotalSubAcc = 0
        End If
        PrintHeader7014 oTextFile, nLine, nPage, dCuteDate, sGL, sAccountName, sRepNum, _
                        sNameA, sPuestoA, sNameB, sPuestoB, bWithSector
        If (bWithSector) Then
          sPrintLine = SpcFill(!fecha_ult, 10) & " " & _
                       SpcFill(!NUMERO_CUENTA_AUXILIAR, 20) & " " & _
                       SpcFill(!clave_sector, 4) & " " & _
                       SpcFill(!nombre_cuenta_auxiliar, 35) & " " & _
                       SpcFillAmount(!Saldo_Actual, 22)
        ElseIf Not (bWithSector) Then
          sPrintLine = SpcFill(!fecha_ult, 10) & " " & _
                       SpcFill(!NUMERO_CUENTA_AUXILIAR, 20) & " " & _
                       SpcFill(!nombre_cuenta_auxiliar, 35) & " " & _
                       SpcFillAmount(!Saldo_Actual, 22)
        End If
        oTextFile.PrintLine sPrintLine
        oTextFile.PrintLine Space(37 - nSpcSector) & SpcFill(!descripcion, 35)
        nLine = nLine + 1
        nTotalSubAcc = nTotalSubAcc + !Saldo_Actual
        sPreviousAcc = !numero_cuenta_estandar
        sSector = !clave_sector
        sCompleteAcc = GetCuentaSubCuenta(!numero_cuenta_estandar)
        sCompleteAcc = sCompleteAcc & !clave_sector
        .MoveNext
      Loop
      oTextFile.PrintLine Space(51 - nSpcSector) & _
                          SpcFill(sCompleteAcc, 21) & " " & _
                          SpcFill(" ", 4) & Space(17) & _
                          SpcFillAmount(nTotalSubAcc, 22)
                  
      nTotal = nTotal + nTotalSubAcc
      
      oTextFile.PrintNewLine
      oTextFile.PrintLine Space(95 - nSpcSector) & "----------------------"
      oTextFile.PrintLine Space(65 - nSpcSector) & "SALDO IGUAL AL MAYOR ------> " & SpcFillAmount(nTotal, 22)
      oTextFile.PrintLine Space(95 - nSpcSector) & "======================"
    End With
    oRs.Close
    
    If (bWithResumen) Then
      Set oRs = oRSGenerator.GeneraResumen(sAppServer, dCuteDate, sAccount, nGl)
      oTextFile.PrintNewLine
    
      PrintHeader7014 oTextFile, nLine, nPage, dCuteDate, sGL, sAccountName, sRepNum, _
                      sNameA, sPuestoA, sNameB, sPuestoB, bWithSector
     
      oTextFile.PrintLine SpcFill(" ", 6) & "     RESUMEN DE SECTORES      "
      oTextFile.PrintLine SpcFill(" ", 6) & " SECTOR               IMPORTE "
      oTextFile.PrintLine SpcFill(" ", 6) & "------------------------------"

      With oRs
        Do
          If (.EOF) Then Exit Do
          oTextFile.PrintLine SpcFill(" ", 7) & SpcFill(!clave_sector, 6) & SpcFillAmount(!Saldo_Actual, 22)
          nTotSectores = nTotSectores + !Saldo_Actual
          .MoveNext
        Loop
      End With
      oTextFile.PrintLine SpcFill(" ", 6) & "------------------------------"
      oTextFile.PrintLine SpcFill(" ", 13) & SpcFillAmount(nTotSectores, 22)
    ElseIf Not (bWithResumen) Then
      '
    End If
    'nTotalconLetra = CantidadConLetra(nTotal)
    PrintSigns oTextFile, sNameA, sPuestoA, sNameB, sPuestoB, _
                    Len("================================================================================================================")
    'Cuaderno_7014 = oTextFile.FileName
    'oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Cuaderno_7014", Err.Number
End Function

Private Sub PrintHeader7014(oTextFile As CTextFile, _
                              nLine As Long, nPage As Long, _
                              dCuteDate As Date, _
                              sGL As String, sAccountName As String, _
                              sRepNum As String, _
                              sNameA As String, sPuestoA As String, _
                              sNameB As String, sPuestoB As String, _
                              Optional bWithSector As Boolean = False)
  Dim sDate       As String
  '**********************************************************
  On Error GoTo ErrHandler
    If (nLine >= cnPageLines) Then
      If nLine <> 100 Then
         PrintSigns oTextFile, sNameA, sPuestoA, sNameB, sPuestoB, _
                    Len("================================================================================================================")
         'oTextFile.PageFeed (nLine - 5)
      End If
      'sDate = " AL " & Day(dCuteDate) & " DE " & UCase(MonthName(Month(dCuteDate))) & " DE " & Year(dCuteDate)
      sDate = " AL " & Format(dCuteDate, "d"" DE ""MMMM"" DE ""yyyy")
      oTextFile.PrintLine "TECOYOTITLA NUM 100             BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C.      FECHA: " & Date
      oTextFile.PrintLine "COLONIA FLORIDA                  " & SpcFill(sAccountName & sDate, 57) & "PAGINA:   " & nPage
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F.          ===================================================="
      oTextFile.PrintLine "BNO-670315-CDO"
      oTextFile.PrintLine sRepNum
      oTextFile.PrintLine "                                  DELEGACION   : " & sGL
      If (bWithSector) Then
        oTextFile.PrintLine "====================================================================================================================="
        oTextFile.PrintLine "  FECHA    AUXILIAR           SECTOR DESCRIPCION                                        IMPORTE                TOTAL "
        oTextFile.PrintLine "---------------------------------------------------------------------------------------------------------------------"
      ElseIf Not (bWithSector) Then
        oTextFile.PrintLine "================================================================================================================"
        oTextFile.PrintLine "  FECHA    AUXILIAR             DESCRIPCION                                        IMPORTE                TOTAL "
        oTextFile.PrintLine "----------------------------------------------------------------------------------------------------------------"
      End If
      'oTextFile.PrintLine
      'nLine = 10
      nLine = 20
      nPage = nPage + 1
    End If
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintHeader7014", Err.Number
End Sub


Private Function Cuaderno_7021(oTextFile As CTextFile, sAppServer As String, _
                              dCuteDate As Date, _
                              vGL As Variant, _
                              sNameA As String, sPuestoA As String, _
                              sNameB As String, sPuestoB As String) As String
  Dim oRs              As ADODB.Recordset
  Dim oRSGenerator     As New SCFFixedReportsDS.CNotebookData
  'Dim oTextFile        As New CTextFile
                                                          
  Dim nTotRubro(18)    As Currency
  Dim sGL              As String
  Dim nGl              As Long
  Dim sAccountName     As String
  Dim sAccount         As String
  
  Dim nLine            As Long
  Dim nPage            As Long
  Dim i                As Long
  Dim nTotalconLetra As String
  '***********************************************************************
  On Error GoTo ErrHandler
    'oTextFile.GenerateRandomFile
    'oTextFile.Open_
    'oTextFile.Open_ "c:\desarrollos\cuaderno\", "Cuaderno7021.txt"
    Set oRs = oRSGenerator.GeneraRS_7021(sAppServer, dCuteDate, vGL)
    If (IsArray(vGL)) Then
      nGl = vGL(0)
    Else
      nGl = vGL
    End If
    sGL = oRSGenerator.GetObjName(sAppServer, nGl, "COF_MAYOR", "NOMBRE_MAYOR", "ID_MAYOR")
    sAccount = "1350"
    sAccountName = sAccount & " - " & oRSGenerator.GetAccountName(sAppServer, sAccount, nGl)
    
    nLine = 100
    nPage = 1
    With oRs
      Do While Not .EOF()
        PrintHeader7021 oTextFile, nLine, nPage, dCuteDate, sAccountName, sGL, "7021", _
                        sNameA, sPuestoA, sNameB, sPuestoB
        oTextFile.PrintLine SpcFill(!AUX, 18) & " " & SpcFill(!ACREDITADO, 30) & _
                            SpcFillAmount(!CAPITAL_VIG, 17) & " " & _
                            SpcFillAmount(!INT_VIGENTE, 17) & " " & _
                            SpcFillAmount(!CAPITAL_VIG + !INT_VIGENTE, 17) & " " & _
                            SpcFillAmount(!CAP_VENC_EXIGIBLE, 17) & " " & _
                            SpcFillAmount(!INTERESES_VENCIDS, 17) & " " & _
                            SpcFill(" ", 37) & " " & _
                            SpcFillAmount(!INTER_NORM_EJ_CSO, 17) & " " & _
                            SpcFillAmount(!INTER_MORATOR, 17)
        oTextFile.PrintLine SpcFill(" ", 49) & _
                            SpcFillAmount(!AMORT_X_TRA, 17) & " " & _
                            SpcFillAmount(!INTS_NOR_X_TRASP, 17) & " " & _
                            SpcFillAmount(!AMORT_X_TRA + !INTS_NOR_X_TRASP, 17) & " " & _
                            SpcFill(" ", 76) & _
                            SpcFillAmount(!INTER_NORM_EJ_ANT, 17) & " "
        oTextFile.PrintLine SpcFill(" ", 68) & _
                            SpcFillAmount(!INTS_COR_X_TRASP, 17) & " " & _
                            SpcFillAmount(!CAPITAL_VIG + !INT_VIGENTE + !AMORT_X_TRA + !INTS_NOR_X_TRASP, 17) & " " & _
                            SpcFillAmount(!CAP_VENC_NO_EXIGIBLE, 17) & " " & _
                            SpcFillAmount(!INTERESES_CORRIDS, 17) & " " & _
                            SpcFillAmount(!CAP_VENC_EXIGIBLE + !CAP_VENC_NO_EXIGIBLE + !INTERESES_VENCIDS + !INTERESES_CORRIDS, 17) & " " & _
                            SpcFillAmount((!CAPITAL_VIG + !INT_VIGENTE + !AMORT_X_TRA + !INTS_NOR_X_TRASP) - (!CAP_VENC_EXIGIBLE + !CAP_VENC_NO_EXIGIBLE + !INTERESES_VENCIDS + !INTERESES_CORRIDS), 17) & " " & _
                            SpcFillAmount(!INTER_CORR_EJ_CSO, 17) & " "
        oTextFile.PrintLine SpcFill(" ", 163) & _
                            SpcFillAmount(!INTER_CORR_EJ_ANT, 17)
        nTotRubro(0) = nTotRubro(0) + !CAPITAL_VIG
        nTotRubro(1) = nTotRubro(1) + !INT_VIGENTE
        nTotRubro(2) = nTotRubro(2) + !CAPITAL_VIG + !INT_VIGENTE
        nTotRubro(3) = nTotRubro(3) + !CAP_VENC_EXIGIBLE
        nTotRubro(4) = nTotRubro(4) + !INTERESES_VENCIDS
        nTotRubro(5) = nTotRubro(5) + !INTER_NORM_EJ_CSO
        nTotRubro(6) = nTotRubro(6) + !INTER_MORATOR
        nTotRubro(7) = nTotRubro(7) + !AMORT_X_TRA
        nTotRubro(8) = nTotRubro(8) + !INTS_NOR_X_TRASP
        nTotRubro(9) = nTotRubro(9) + !AMORT_X_TRA + !INTS_NOR_X_TRASP
        nTotRubro(10) = nTotRubro(10) + !INTER_NORM_EJ_ANT
        nTotRubro(11) = nTotRubro(11) + !INTS_COR_X_TRASP
        nTotRubro(12) = nTotRubro(12) + !CAPITAL_VIG + !INT_VIGENTE + !AMORT_X_TRA + !INTS_NOR_X_TRASP
        nTotRubro(13) = nTotRubro(13) + !CAP_VENC_NO_EXIGIBLE
        nTotRubro(14) = nTotRubro(14) + !INTERESES_CORRIDS
        nTotRubro(15) = nTotRubro(15) + !CAP_VENC_EXIGIBLE + !CAP_VENC_NO_EXIGIBLE + !INTERESES_VENCIDS + !INTERESES_CORRIDS
        nTotRubro(16) = nTotRubro(16) + (!CAPITAL_VIG + !INT_VIGENTE + !AMORT_X_TRA + !INTS_NOR_X_TRASP) - (!CAP_VENC_EXIGIBLE + !CAP_VENC_NO_EXIGIBLE + !INTERESES_VENCIDS + !INTERESES_CORRIDS)
        nTotRubro(17) = nTotRubro(17) + !INTER_CORR_EJ_CSO
        nTotRubro(17) = nTotRubro(17) + !INTER_CORR_EJ_ANT
        nLine = nLine + 4
        .MoveNext
      Loop
      nLine = nLine + 4
      PrintHeader7021 oTextFile, nLine, nPage, dCuteDate, sAccountName, sGL, "7021", _
                      sNameA, sPuestoA, sNameB, sPuestoB
      PrintTotals7021 oTextFile, nTotRubro
    End With
    PrintFeed oTextFile, sAccount
    'nTotalconLetra = CantidadConLetra(nTotsectore)
    PrintSigns oTextFile, sNameA, sPuestoA, sNameB, sPuestoB, _
                   Len("---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------")
    'oTextFile.Close_
    'Cuaderno_7021 = oTextFile.FileName
    Exit Function
ErrHandler:
  RaiseError ClassId, "Cuaderno_7021", Err.Number
End Function

Private Sub PrintHeader7021(oTextFile As CTextFile, _
                            nLine As Long, nPage As Long, _
                            dCuteDate As Date, sAccountName As String, _
                            sGL As String, _
                            sRepNum As String, _
                            sNameA As String, sPuestoA As String, _
                            sNameB As String, sPuestoB As String)
  Dim sDate       As String
  '**********************************************************
  On Error GoTo ErrHandler
    If (nLine >= cnPageLines - 3) Then
      If nLine <> 100 Then
        PrintSigns oTextFile, sNameA, sPuestoA, sNameB, sPuestoB, _
                   Len("---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------")
        'oTextFile.PageFeed (nLine - 5)
      End If
      'sDate = " AL " & Day(dCuteDate) & " DE " & UCase(MonthName(Month(dCuteDate))) & " DE " & Year(dCuteDate)
      sDate = " AL " & Format(dCuteDate, "d"" DE ""MMMM"" DE ""yyyy")
      oTextFile.PrintLine "TECOYOTITLA NUM 100                                                               BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C.                                                     FECHA: " & Date
      oTextFile.PrintLine "COLONIA FLORIDA                                                              " & SpcFill(sAccountName & sDate, 110) & "PAGINA:   " & nPage
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F.                                                    ===================================================================="
      oTextFile.PrintLine "BNO-670315-CDO                                                                               DELEGACION   : " & sGL
      oTextFile.PrintLine sRepNum
      oTextFile.PrintLine "MONEDA : MONEDA NACIONAL"
      oTextFile.PrintLine Space(183) & "C U E N T A S  D E  O R D E N"
      oTextFile.PrintLine "No. DE CREDITO  NOMBRE DEL ACREDITADO"
      oTextFile.PrintLine "AUXILIAR           ACREDITADO                   | CAPITAL VIGENTE  | INTERES VIGENTE  | CARTERA VIGENTE   |CAP VDO EXIGIBLE  | INTRS VENCIDOS  |                  |                  | INT NORM EJ CSO  | INTERES MORAT "
      oTextFile.PrintLine "                                                | AMORTZ POR TRASP | INT NORM POR TRA | CARTERA POR TRAS  |                  |                 |                  |                  | INT NORM EJ ANT  |"
      oTextFile.PrintLine "                                                |                  | INT CORR POR TRA | TOTAL VIGENTE     | CAP VDO NO EXIG  | INTRS CORRIDOS  |  TOTAL VENCIDO   |  TOTAL CARTERA   | INT CORR EJ CSO  |"
      oTextFile.PrintLine "                                                |                  |                  |                   |                  |                 |                  |                  | INT CORR EJ ANT  |"
      oTextFile.PrintLine "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
      'nLine = 12
      nLine = 22
      nPage = nPage + 1
    End If
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintHeader7021", Err.Number
End Sub

Private Sub PrintTotals7021(oTextFile As CTextFile, _
                            vTotals As Variant)
  Dim sDate       As String
  '**********************************************************
  On Error GoTo ErrHandler
    oTextFile.PrintLine Space(5) & " T O T A L E S " & Space(29) & "--------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    oTextFile.PrintLine Space(49) & _
                        SpcFillAmount(vTotals(0), 17) & " " & _
                        SpcFillAmount(vTotals(1), 17) & " " & _
                        SpcFillAmount(vTotals(2), 17) & " " & _
                        SpcFillAmount(vTotals(3), 17) & " " & _
                        SpcFillAmount(vTotals(4), 17) & " " & _
                        SpcFill(" ", 37) & " " & _
                        SpcFillAmount(vTotals(5), 17) & " " & _
                        SpcFillAmount(vTotals(6), 17)
    oTextFile.PrintLine SpcFill(" ", 49) & _
                        SpcFillAmount(vTotals(7), 17) & " " & _
                        SpcFillAmount(vTotals(8), 17) & " " & _
                        SpcFillAmount(vTotals(9), 17) & " " & _
                        SpcFill(" ", 76) & _
                        SpcFillAmount(vTotals(10), 17) & " "
    oTextFile.PrintLine SpcFill(" ", 68) & _
                        SpcFillAmount(vTotals(11), 17) & " " & _
                        SpcFillAmount(vTotals(12), 17) & " " & _
                        SpcFillAmount(vTotals(13), 17) & " " & _
                        SpcFillAmount(vTotals(14), 17) & " " & _
                        SpcFillAmount(vTotals(15), 17) & " " & _
                        SpcFillAmount(vTotals(16), 17) & " " & _
                        SpcFillAmount(vTotals(17), 17) & " "
    oTextFile.PrintLine SpcFill(" ", 163) & _
                        SpcFillAmount(vTotals(18), 17)
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintTotals7021", Err.Number
End Sub

Private Function Cuaderno_7002(oTextFile As CTextFile, sAppServer As String, nGl As Long, _
                              dCuteDate As Date, sAccount As String, _
                              sNameA As String, sPuestoA As String, _
                              sNameB As String, sPuestoB As String, dExcDate As Date, _
                              Optional nExcRateType As Long = 46, _
                              Optional nExcRateCurrency As Long = 1) As String
                              
  Dim oGeneraRS         As New SCFFixedReportsDS.CNotebookData
  'Dim oTextFile         As New CTextFile
  Dim oRs               As ADODB.Recordset
  Dim sDateTime         As String
  Dim sDate             As String
  Dim sGLDesc           As String
  Dim sAccountName      As String
  Dim sFull             As String
  Dim sTotconLetra      As String
  Dim nLine             As Long
  Dim nPage             As Long
  Dim nSum              As Double
  Dim x                 As Double
  Dim nTotalconLetra As String
  '**************************************
  On Error GoTo ErrorHandler
    sDateTime = Format(Date, "dd/mmm/yyyy") & "  " & Format(Time, "hh:mm:ss")
    nLine = 100
    nPage = 1
    sGLDesc = oGeneraRS.GetObjName(sAppServer, nGl, "COF_MAYOR", "NOMBRE_MAYOR", "ID_MAYOR")
    Set oRs = oGeneraRS.GetRS_R7002(sAppServer, nGl, dCuteDate, sAccount, dExcDate, nExcRateType, nExcRateCurrency)
    
    sAccountName = sAccount & " - " & oGeneraRS.GetAccountName(sAppServer, sAccount, nGl)
    If (oRs.BOF And oRs.EOF) Then
      PrintHeader7002 oTextFile, nLine, nPage, sGLDesc, sAccountName, dCuteDate, sDateTime, _
                      sNameA, sPuestoA, sNameB, sPuestoB
      oTextFile.PrintLine " NO EXISTE INFORMACION "
      Exit Function
    End If
    
    With oRs
      'sAccountName = sAccount & ".- " & !nom_cuenta
      Do While Not .EOF()
             PrintHeader7002 oTextFile, nLine, nPage, sGLDesc, sAccountName, dCuteDate, sDateTime, _
                             sNameA, sPuestoA, sNameB, sPuestoB
             sFull = GetSubCuenta(!numero_cuenta_estandar) & !clave_sector
             oTextFile.PrintLine SpcFill(sFull, 28) & " " & SpcFill(!nombre_cuenta_estandar, 50) & "  " & _
                                SpcFillAmount(!Importe, 15)
             nLine = nLine + 1
             nSum = nSum + !Importe
        
           .MoveNext
      Loop
      oTextFile.PrintLine "                                                                                                       --------------------- "
      oTextFile.PrintLine "                              SALDO IGUAL AL MAYOR ---------->     " & SpcFillAmount(nSum, 57)
      'sTotconLetra = CantidadConLetra(nSum)
      oTextFile.PrintLine "  " & sTotconLetra
      .Close
    End With
    PrintFeed oTextFile, sAccount
    'nTotalconLetra = CantidadConLetra(nSum)
    PrintSigns oTextFile, sNameA, sPuestoA, sNameB, sPuestoB, _
                Len("============================================================================================================================")
    Set oRs = Nothing
    'Cuaderno_7002 = oTextFile.FilePath & oTextFile.FileName

    'oTextFile.Close_
  Exit Function
ErrorHandler:
  RaiseError ClassId, "Cuaderno_7002", Err.Number
End Function

Private Sub PrintHeader7002(oTextFile As CTextFile, nLine As Long, nPage As Long, sGLDesc As String, _
                            sAccountName As String, dCuteDate As Date, sDateTime As String, _
                            sNameA As String, sPuestoA As String, sNameB As String, sPuestoB As String)
 Dim sDay    As String
 Dim sMonth  As String
 Dim syear   As String
 Dim sPaces  As String
 Dim sDate   As String
 
 Dim nLen    As Long
'**********************
 On Error GoTo ErrorHandler
 
 If (nLine >= cnPageLines) Then
      If nLine <> 100 Then
        nLen = Len("============================================================================================================================")
        PrintSigns oTextFile, sNameA, sPuestoA, sNameB, sPuestoB, nLen
      End If
      'sDate = " AL " & Day(dCuteDate) & " DE " & UCase(MonthName(Month(dCuteDate))) & " DE " & Year(dCuteDate)
      
      sDate = " AL " & Format(dCuteDate, "d"" DE ""MMMM"" DE ""yyyy")
      
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine "TECOYOTITLA NUM 100             BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C.                FECHA: " & Date
      oTextFile.PrintLine "COLONIA FLORIDA                  " & SpcFill(sAccountName & sDate, 60) & "        PAGINA:   " & nPage
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F.          ===================================================="
      oTextFile.PrintLine "BNO-670315-CDO "
      oTextFile.PrintLine "7002"
      oTextFile.PrintNewLine
      oTextFile.PrintLine "                                               DELEGACION   :  " & sGLDesc
      oTextFile.PrintLine "============================================================================================================================"
      oTextFile.PrintLine "NIVEL ANALITICO              D E S C R I P C I O N                           I M P O R T E                     T O T A L    "
      oTextFile.PrintLine "----------------------------------------------------------------------------------------------------------------------------"

      nPage = nPage + 1
      nLine = 24
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader7002", Err.Number
End Sub

Private Function Cuaderno_7016(oTextFile As CTextFile, sAppServer As String, vGL As Variant, _
                              sSubAcc As String, dIniDate As Date, _
                              dEndDate As Date, dExcDate As Date, _
                              sNameA As String, sPuestoA As String, _
                              sNameB As String, sPuestoB As String, _
                              Optional nExcRateType As Long = 46, _
                              Optional nExcRateCurrency As Long = 1) As String

  Dim oRs             As ADODB.Recordset
  'Dim oRsPosting      As ADODB.Recordset
  'Dim oRsAccNames     As ADODB.Recordset
  Dim oRSGenerator  As New SCFFixedReportsDS.CNotebookData
  'Dim oTextFile       As New CTextFile
  Dim sLongDate       As String
  
  Dim nTotYearA       As Currency
  Dim nTotAccA        As Currency
  Dim nTotYearB       As Currency
  Dim nTotAccB        As Currency
  Dim nLine           As Long
  Dim nPage           As Long
  Dim nCurr           As Long
  Dim nYear           As Long
  Dim nYearComp       As Long
  Dim nExcRate        As Currency
  
  Dim sDateTime       As String
  Dim sFileName       As String
  Dim sGL             As String
  Dim sName1801       As String
  Dim sName3103       As String
  Dim sNameSubAcc     As String
  Dim sCompleteAcc    As String
  
  Dim sCurrName       As String
  Dim nTotalconLetra As String
  '***********************************************************************
  On Error GoTo ErrorHandler
    sDateTime = Format(Date, "dd/mm/yyyy") & "  " & Format(Time, "hh:mm:ss")
    sLongDate = Format(dEndDate, "d") & " de " & Format(dEndDate, "MMMM") & " de " & Year(dEndDate)
    If (IsArray(vGL)) Then
      sGL = "0000 AGENCIA CONSOLIDADORA BANOBRAS"
    Else
      nLine = vGL
      sGL = oRSGenerator.GetObjName(sAppServer, nLine, "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
    End If
    
    sName1801 = oRSGenerator.GetAccountName(sAppServer, "1801-" & sSubAcc, nLine)
    sName1801 = sSubAcc & ".- " & sName1801
    sName3103 = oRSGenerator.GetAccountName(sAppServer, "3103-" & sSubAcc, nLine)
    sName3103 = sSubAcc & ".- " & sName3103
    sNameSubAcc = SpcFill(sName1801, 20) & Space(80) & SpcFill(sName3103, 20)
    nLine = 100
    nPage = 1

    Set oRs = oRSGenerator.GeneraRecordsetC7016(sAppServer, vGL, sSubAcc, dIniDate, dEndDate, dExcDate, nExcRateType, nExcRateCurrency)
    If ((oRs.BOF) And (oRs.EOF)) Then
      PrintHeaderC_7016 oTextFile, nLine, nPage, sLongDate, sGL, " ", sNameSubAcc, sDateTime, sNameA, _
                        sPuestoA, sNameB, sPuestoB
      oTextFile.PrintLine " NO EXISTE INFORMACION "
      Exit Function
    End If
    With oRs
      Do While Not .EOF()
        If IsNull(!FECHA_REGISTRO) Then
          nYear = 0
        Else
          nYear = Year(!FECHA_REGISTRO)
        End If
        If ((!ST_A <> 0)) Then
          PrintHeaderC_7016 oTextFile, nLine, nPage, sLongDate, sGL, " ", sNameSubAcc, sDateTime, sNameA, _
                            sPuestoA, sNameB, sPuestoB
          oTextFile.PrintLine SpcFill(IIf(IsNull(!FECHA_REGISTRO), " ", !FECHA_REGISTRO), 19) & "  " & _
                              SpcFill(!NUMERO_CUENTA_AUXILIAR, 21) & "  " & _
                              SpcFill(!nombre_cuenta_auxiliar, 38) & "  " & _
                              SpcFillAmount(!ST_A, 20) & "  " & Space(5) & _
                              SpcFillAmount(Round((!st_b / IIf((!ST_A = 0), 1, !ST_A)), 2) * (IIf((!ST_A = 0), 0, 100)), 6) & "%" & _
                              "  " & SpcFillAmount(!st_b, 20)
          nTotYearA = nTotYearA + !ST_A
          nTotYearB = nTotYearB + !st_b
          nTotAccA = nTotAccA + !ST_A
          nTotAccB = nTotAccB + !st_b
          nLine = nLine + 1
          .MoveNext
            If Not .EOF Then
              If IsNull(!FECHA_REGISTRO) Then
                nYearComp = 0
              Else
                nYearComp = Year(!FECHA_REGISTRO)
              End If
              If nYear <> nYearComp Then
                oTextFile.PrintLine Space(100) & "============================================"
                oTextFile.PrintLine Space(20) & SpcFill(" TOTAL AÑO: " & ZeroFill(nYear, 4), 35) & Space(29) & SpcFillAmount(nTotYearA, 20) & " " & Space(16) & SpcFillAmount(nTotYearB, 20)
                oTextFile.PrintNewLine
                nLine = nLine + 3
                nTotYearA = 0
                nTotYearB = 0
              End If
            End If
        Else
          .MoveNext
        End If
      Loop
        oTextFile.PrintLine Space(100) & "============================================"
        oTextFile.PrintLine Space(20) & SpcFill(" TOTAL AÑO: " & ZeroFill(nYear, 4), 35) & Space(29) & SpcFillAmount(nTotYearA, 20) & " " & Space(16) & SpcFillAmount(nTotYearB, 20)
        oTextFile.PrintNewLine
        oTextFile.PrintLine Space(100) & "============================================"
        sCompleteAcc = GetCuentaSubCuenta("1801" & sSubAcc)
        sCompleteAcc = sCompleteAcc
        oTextFile.PrintLine Space(20) & SpcFill(" TOTAL CUENTA " & sCompleteAcc, 35) & Space(29) & SpcFillAmount(nTotAccA, 20) & " " & Space(16) & SpcFillAmount(nTotAccB, 20)
      .Close
    End With
    Set oRs = Nothing
    'nTotalconLetra = CantidadConLetra(nTotAccB)
    PrintSigns oTextFile, sNameA, sPuestoA, sNameB, sPuestoB, _
                Len("================================================================================================================================================")
  Exit Function
ErrorHandler:
  RaiseError ClassId, "Cuaderno_7016", Err.Number
End Function

Private Sub PrintHeaderC_7016(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                                    dCuteDate As String, _
                                    sGLDesc As String, _
                                    skindacc As String, _
                                    sAccount As String, _
                                    sDateTime As String, _
                                    sNameA As String, sPuestoA As String, _
                                    sNameB As String, sPuestoB As String)
  
  On Error GoTo ErrorHandler
    If (nLine >= cnPageLines) Then
      If nLine <> 100 Then
        PrintSigns oTextFile, sNameA, sPuestoA, sNameB, sPuestoB, _
                Len("================================================================================================================================================")
      End If
      'Imprimir caracter de avance de pagina
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine sDateTime
      oTextFile.PrintLine "TECOYOTITLA NUM. 100                                                                                                                FECHA  : " & Date
      oTextFile.PrintLine "COL. FLORIDA C.P. 01030 MEXICO D.F.                                                                                                 PAGINA : " & nPage
      oTextFile.PrintLine "BNO-670315-CDO                           BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C "
      oTextFile.PrintLine "                       1801.- MOBILIARIO Y EQUIPO Y 3103.- DEPREC ACUMUL MOBILIARIO Y EQ AL " & dCuteDate
      oTextFile.PrintLine "BMG7016"
      oTextFile.PrintNewLine
      oTextFile.PrintLine "                                                         DELEGACION  : " & sGLDesc
      oTextFile.PrintLine "================================================================================================================================================"
      oTextFile.PrintLine "1801.- MOBILIARIO Y EQUIPO                                                                          3103.- DEPREC ACUMUL MOBILIARIO Y EQ "
      oTextFile.PrintLine sAccount
      oTextFile.PrintLine "================================================================================================================================================"
      oTextFile.PrintLine "FECHA DE             NUMERO DE                                                                 COSTO                             DEPRECIACION  "
      oTextFile.PrintLine "REGISTRO             INVENTARIO             DESCRIPCION                                        HISTORICO             %           ACUMULADA     "
      oTextFile.PrintLine "================================================================================================================================================"
      nPage = nPage + 1
      'nLine = 20
      nLine = 30
    End If
    Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeaderC_7016", Err.Number
End Sub

Private Function Cuaderno_7024(oTextFile As CTextFile, sAppServer As String, dCuteDate As Date, _
                                sAccount As String, nGl As Long, sRepNum As String, _
                                sNameA As String, sPuestoA As String, sNameB As String, sPuestoB As String, _
                                Optional bWithSector As Boolean = False, _
                                Optional bWithResumen As Boolean = False) As String
  Dim oRs           As ADODB.Recordset
  Dim oRSGenerator  As New SCFFixedReportsDS.CNotebookData
  'Dim oTextFile      As New CTextFile
  
  Dim sPrintLine    As String
  Dim sPreviousAcc  As String
  Dim sPrevSubAcc   As String
  Dim sSector       As String
  Dim nLine         As Long
  Dim nPage         As Long
  Dim nTotalSubAcc  As Currency
  Dim nTotalLevel   As Currency
  Dim nTotal        As Currency
  Dim nTotSectores  As Currency
  Dim nSpcSector    As Long
  
  Dim sGL           As String
  Dim sAccountName  As String
  Dim nTotalconLetra As String
  '********************************************************************************
  On Error GoTo ErrHandler
    'oTextFile.GenerateRandomFile
    'oTextFile.Open_
    'oTextFile.Open_ "c:\development\", "Cuaderno7024.txt"
    Set oRs = oRSGenerator.GetRS_7024(sAppServer, dCuteDate, sAccount, nGl)
    sGL = oRSGenerator.GetObjName(sAppServer, nGl, "COF_MAYOR", "NOMBRE_MAYOR", "ID_MAYOR")
    sAccountName = sAccount & " - " & oRSGenerator.GetAccountName(sAppServer, sAccount, nGl)
   
    nTotal = 0
    nTotalSubAcc = 0
    nLine = 100
    nPage = 1
    
    If ((oRs.BOF) And (oRs.EOF)) Then
      PrintHeader7024 oTextFile, nLine, nPage, dCuteDate, sGL, sAccountName, sRepNum, _
                      sNameA, sPuestoA, sNameB, sPuestoB
      oTextFile.PrintLine " NO EXISTE INFORMACION "
      Exit Function
    End If
    
    With oRs
      sPreviousAcc = !numero_cuenta_estandar
      sSector = !clave_sector
      sPrevSubAcc = Mid(!numero_cuenta_estandar, 6, 2)
      
      If (bWithSector) Then
        nSpcSector = 0
      ElseIf Not (bWithSector) Then
        nSpcSector = 5
      End If
      
      Do While Not .EOF()
        If (!numero_cuenta_estandar <> sPreviousAcc) Or (sSector <> !clave_sector) Then
          PrintHeader7024 oTextFile, nLine, nPage, dCuteDate, sGL, sAccountName, sRepNum, _
                          sNameA, sPuestoA, sNameB, sPuestoB
          oTextFile.PrintLine Space(36 - nSpcSector + 66) & "========================="
          oTextFile.PrintLine Space(14 - nSpcSector) & _
                              SpcFill(" TOTAL PARCIAL ----> ", 22) & _
                              SpcFill(Mid(sPreviousAcc, 6), 21) & " " & _
                              SpcFill(oRSGenerator.GetAccountName(sAppServer, sPreviousAcc, nGl), 44) & "  " & _
                              SpcFillAmount(nTotalSubAcc, 22)
          oTextFile.PrintLine ""
          nLine = nLine + 2
          nTotal = nTotal + nTotalSubAcc
          nTotalSubAcc = 0
        End If
        
        If (Mid(!numero_cuenta_estandar, 6, 2) <> sPrevSubAcc) Then
          PrintHeader7024 oTextFile, nLine, nPage, dCuteDate, sGL, sAccountName, sRepNum, _
                          sNameA, sPuestoA, sNameB, sPuestoB
          oTextFile.PrintLine Space(36 - nSpcSector + 66) & "========================="
          oTextFile.PrintLine Space(14 - nSpcSector) & _
                              SpcFill(" TOTAL PARCIAL POR NIVEL ----> ", 31) & _
                              SpcFill(Mid(sPreviousAcc, 6, 2), 3) & " " & _
                              SpcFill(oRSGenerator.GetAccountName(sAppServer, Mid(sPreviousAcc, 1, 7), nGl), 55) & _
                              SpcFillAmount(nTotalLevel, 22)
          oTextFile.PrintLine ""
          nLine = nLine + 2
          nTotalLevel = 0
        End If
        PrintHeader7024 oTextFile, nLine, nPage, dCuteDate, sGL, sAccountName, sRepNum, _
                        sNameA, sPuestoA, sNameB, sPuestoB
        
        
        sPrintLine = SpcFill(IIf(IsNull(!RFC), " ", !RFC), 10) & "  " & _
                     SpcFill(IIf(IsNull(!nombre_cuenta_auxiliar), " ", !nombre_cuenta_auxiliar), 48) & "  " & _
                     SpcFill(IIf(IsNull(!NUMERO_CUENTA_AUXILIAR), " ", !NUMERO_CUENTA_AUXILIAR), 18) & "  " & _
                     SpcFillAmount(!Saldo_Actual, 22)
        
        oTextFile.PrintLine sPrintLine
        oTextFile.PrintLine Space(10) & "  " & SpcFill(!DOMICILIO, 48)
        oTextFile.PrintLine Space(10) & "  " & SpcFill(!COLONIA, 48)
        oTextFile.PrintLine Space(10) & "  " & SpcFill(!ESTADO, 48)
        oTextFile.PrintNewLine
        
        nLine = nLine + 5
        nTotalSubAcc = nTotalSubAcc + !Saldo_Actual
        nTotalLevel = nTotalLevel + !Saldo_Actual
        sPreviousAcc = !numero_cuenta_estandar
        sPrevSubAcc = Mid(!numero_cuenta_estandar, 6, 2)
        sSector = !clave_sector
        .MoveNext
      Loop
      oTextFile.PrintLine Space(36 - nSpcSector + 66) & "========================="
      oTextFile.PrintLine Space(14 - nSpcSector) & _
                          SpcFill(" TOTAL PARCIAL ----> ", 22) & _
                          SpcFill(Mid(sPreviousAcc, 6), 21) & " " & _
                          SpcFill(oRSGenerator.GetAccountName(sAppServer, sPreviousAcc, nGl), 44) & "  " & _
                          SpcFillAmount(nTotalSubAcc, 22)
                          
      oTextFile.PrintNewLine
      oTextFile.PrintLine Space(36 - nSpcSector + 66) & "========================="
      oTextFile.PrintLine Space(14 - nSpcSector) & _
                              SpcFill(" TOTAL PARCIAL POR NIVEL ----> ", 31) & _
                              SpcFill(Mid(sPreviousAcc, 6, 2), 3) & " " & _
                              SpcFill(oRSGenerator.GetAccountName(sAppServer, Mid(sPreviousAcc, 1, 7), nGl), 55) & _
                              SpcFillAmount(nTotalLevel, 22)
      nTotal = nTotal + nTotalSubAcc
      
      oTextFile.PrintNewLine
      oTextFile.PrintLine Space(105 - nSpcSector) & "----------------------"
      oTextFile.PrintLine Space(75 - nSpcSector) & "SALDO IGUAL AL MAYOR ------> " & SpcFillAmount(nTotal, 22)
      oTextFile.PrintLine Space(105 - nSpcSector) & "======================"
    End With
    oRs.Close
    
    If (bWithResumen) Then
      Set oRs = oRSGenerator.GeneraResumen(sAppServer, dCuteDate, sAccount, nGl)
      oTextFile.PrintNewLine
    
      PrintHeader7024 oTextFile, nLine, nPage, dCuteDate, sGL, sAccountName, sRepNum, _
                      sNameA, sPuestoA, sNameB, sPuestoB
     
      oTextFile.PrintLine SpcFill(" ", 6) & "     RESUMEN DE SECTORES      "
      oTextFile.PrintLine SpcFill(" ", 6) & " SECTOR               IMPORTE "
      oTextFile.PrintLine SpcFill(" ", 6) & "------------------------------"

      With oRs
        Do
          If (.EOF) Then Exit Do
          oTextFile.PrintLine SpcFill(" ", 7) & SpcFill(!clave_sector, 6) & SpcFillAmount(!Saldo_Actual, 22)
          nTotSectores = nTotSectores + !Saldo_Actual
          .MoveNext
        Loop
      End With
      oTextFile.PrintLine SpcFill(" ", 6) & "------------------------------"
      oTextFile.PrintLine SpcFill(" ", 13) & SpcFillAmount(nTotSectores, 22)
    ElseIf Not (bWithResumen) Then
      '
    End If
    PrintFeed oTextFile, sAccount
    'nTotalconLetra = CantidadConLetra(nTotal)
    PrintSigns oTextFile, sNameA, sPuestoA, sNameB, sPuestoB, _
                   Len("==========================================================================================================================")
    'oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Cuaderno_7024", Err.Number
End Function

Private Sub PrintHeader7024(oTextFile As CTextFile, _
                            nLine As Long, nPage As Long, _
                            dCuteDate As Date, _
                            sGL As String, sAccountName As String, _
                            sRepNum As String, _
                            sNameA As String, sPuestoA As String, sNameB As String, sPuestoB As String)
  Dim sDate       As String
  '**********************************************************
  On Error GoTo ErrHandler
    If (nLine >= cnPageLines) Then
      If nLine <> 100 Then
        PrintSigns oTextFile, sNameA, sPuestoA, sNameB, sPuestoB, _
                   Len("==========================================================================================================================")
      End If
      sDate = " AL " & Format(dCuteDate, "d"" DE ""MMMM"" DE ""yyyy")
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine "TECOYOTITLA NUM 100            BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C.                FECHA  : " & Date
      oTextFile.PrintLine "COLONIA FLORIDA                 " & SpcFill(sAccountName & sDate, 68) & "PAGINA :  " & nPage
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F.         ================================================================"
      oTextFile.PrintLine "BNO-670315-CDO"
      oTextFile.PrintLine sRepNum
      oTextFile.PrintLine "                                  DELEGACION   : " & sGL
      oTextFile.PrintLine "=========================================================================================================================="
      oTextFile.PrintLine "            DESCRIPCION/                                                                                         "
      oTextFile.PrintLine "  RFC       DIRECCION                                         AUXILIAR                        IMPORTE               TOTAL "
      oTextFile.PrintLine "--------------------------------------------------------------------------------------------------------------------------"
      'oTextFile.PrintLine
      'nLine = 10
      nLine = 24
      nPage = nPage + 1
    End If
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintHeader7024", Err.Number
End Sub

Private Sub PrintIndexHeader(oTextFile As CTextFile, nLine As Long, nPage As Long, dCuteDate As Date, _
                             sGL As String)
  Dim sDate As String
  '**********************************************************
  On Error GoTo ErrHandler
    sDate = " AL " & Format(dCuteDate, "d"" DE ""MMMM"" DE ""yyyy")
    oTextFile.PrintNewLine
    oTextFile.PrintNewLine
    oTextFile.PrintNewLine
    oTextFile.PrintNewLine
    oTextFile.PrintLine "TECOYOTITLA NUM 100            BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C.                FECHA  : " & Date
    oTextFile.PrintLine "COLONIA FLORIDA                INDICE AL " & sDate & "PAGINA : 1"
    oTextFile.PrintLine "C.P. 01030 MEXICO D.F.         ================================================================"
    oTextFile.PrintLine "BNO-670315-CDO"
    oTextFile.PrintLine "                               DELEGACION: " & sGL
    oTextFile.PrintLine "=========================================================================================================================="
    oTextFile.PrintLine "   FOLIO    CUENTA             DESCRIPCION"
    oTextFile.PrintLine "=========================================================================================================================="
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintIndexHeader", Err.Number
End Sub

Private Sub PrintIndexRow(oTextFile As CTextFile, sFolio As String, sAccount As String, sAccountName As String)
  Dim sDate As String
  '************************************************************************************************************
  On Error GoTo ErrHandler
    oTextFile.PrintLine Space(6) & sFolio & Space(4) & SpcFill(sAccount, 19) & sAccountName
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintIndexRow", Err.Number
End Sub

Private Function Cuaderno_7028(oTextFile As CTextFile, sAppServer As String, dCuteDate As Date, _
                              sAccount As String, nGl As Long, sRepNum As String, _
                              sNameA As String, sPuestoA As String, _
                              sNameB As String, sPuestoB As String) As String
  Dim oRs             As ADODB.Recordset
  Dim oRSGenerator    As New SCFFixedReportsDS.CNotebookData
  'Dim oTextFile      As New CTextFile
  
  Dim sPrintLine      As String
  Dim sPreviousAcc    As String
  'Dim sSector      As String
  Dim nLine           As Long
  Dim nPage           As Long
  Dim nTotalSubAcc    As Currency
  Dim nTotal          As Currency
  Dim nTotSectores    As Currency
  Dim nSpcSector      As Long
  
  Dim sGL             As String
  Dim sAccountName    As String
  Dim sCompleteAcc    As String
  Dim nTotalconLetra  As String
  '********************************************************************************
  On Error GoTo ErrHandler
    'oTextFile.GenerateRandomFile
    'oTextFile.Open_
    'oTextFile.Open_ "c:\desarrollos\cuaderno\", "Cuaderno7028.txt"
    Set oRs = oRSGenerator.GetRS_7028(sAppServer, nGl, dCuteDate, sAccount)
    sGL = oRSGenerator.GetObjName(sAppServer, nGl, "COF_MAYOR", "NOMBRE_MAYOR", "ID_MAYOR")
    sAccountName = sAccount & " - " & oRSGenerator.GetAccountName(sAppServer, sAccount, nGl)
   
    nTotal = 0
    nTotalSubAcc = 0
    nLine = 100
    nPage = 1
    
    If ((oRs.BOF) And (oRs.EOF)) Then
      PrintHeader7028 oTextFile, nLine, nPage, dCuteDate, sGL, sAccountName, sRepNum, _
                      sNameA, sPuestoA, sNameB, sPuestoB
      oTextFile.PrintLine " NO EXISTE INFORMACION "
      Exit Function
    End If
    
    With oRs
      sPreviousAcc = !numero_cuenta_estandar
      sCompleteAcc = GetCuentaSubCuenta(!numero_cuenta_estandar)
      sCompleteAcc = sCompleteAcc & "00"

      Do While Not .EOF()
        PrintHeader7028 oTextFile, nLine, nPage, dCuteDate, sGL, sAccountName, sRepNum, _
                        sNameA, sPuestoA, sNameB, sPuestoB
        sPrintLine = SpcFill(IIf(IsNull(!fecha_oficio), " ", !fecha_oficio), 10) & " " & _
                     SpcFill(IIf(IsNull(!numero_oficio), " ", !numero_oficio), 20) & " " & _
                     SpcFill(IIf(IsNull(!NUMERO_CUENTA_AUXILIAR), " ", !NUMERO_CUENTA_AUXILIAR), 18) & " " & _
                     SpcFill(IIf(IsNull(!nombre_cuenta_auxiliar), " ", !nombre_cuenta_auxiliar), 28) & " " & _
                     SpcFillAmount(!Saldo_Actual, 22)
        oTextFile.PrintLine sPrintLine
        nLine = nLine + 1
        nTotalSubAcc = nTotalSubAcc + !Saldo_Actual
        sPreviousAcc = !numero_cuenta_estandar
        sCompleteAcc = GetCuentaSubCuenta(!numero_cuenta_estandar)
        sCompleteAcc = sCompleteAcc & "00"
        .MoveNext
      Loop
      oTextFile.PrintLine Space(58) & _
                          SpcFill(sCompleteAcc, 21) & " " & _
                          Space(20) & _
                          SpcFillAmount(nTotalSubAcc, 22)
                  ' SpcFill(sAccount, 21) & " " & _

      nTotal = nTotal + nTotalSubAcc
      
      oTextFile.PrintNewLine
      oTextFile.PrintLine Space(101) & "----------------------"
      oTextFile.PrintLine Space(71) & "SALDO IGUAL AL MAYOR ------> " & SpcFillAmount(nTotal, 22)
      oTextFile.PrintLine Space(101) & "======================"
    End With
    oRs.Close
    PrintFeed oTextFile, sAccount
    'nTotalconLetra = CantidadConLetra(nTotal)
    PrintSigns oTextFile, sNameA, sPuestoA, sNameB, sPuestoB, _
                   Len("==========================================================================================================================")
    'oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Cuaderno_7028", Err.Number
End Function

Private Sub PrintHeader7028(oTextFile As CTextFile, _
                            nLine As Long, nPage As Long, _
                            dCuteDate As Date, _
                            sGL As String, sAccountName As String, _
                            sRepNum As String, _
                            sNameA As String, sPuestoA As String, sNameB As String, sPuestoB As String)
  Dim sDate             As String
  '**********************************************************
  On Error GoTo ErrHandler
    If (nLine >= cnPageLines) Then
      If nLine <> 100 Then
        PrintSigns oTextFile, sNameA, sPuestoA, sNameB, sPuestoB, _
                   Len("==========================================================================================================================")
        'oTextFile.PageFeed (nLine - 5)
      End If
      'sDate = " AL " & Day(dCuteDate) & " DE " & UCase(MonthName(Month(dCuteDate))) & " DE " & Year(dCuteDate)
      sDate = " AL " & Format(dCuteDate, "d"" DE ""MMMM"" DE ""yyyy")
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      
      oTextFile.PrintLine FormatLine(130, "TECOYOTITLA NUM 100", "BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C.", "FECHA  : " & Date)
      oTextFile.PrintLine "COLONIA FLORIDA                 " & SpcFill(sAccountName & sDate, 68) & "PAGINA :  " & nPage
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F.         ================================================================"
      oTextFile.PrintLine "BNO-670315-CDO"
      oTextFile.PrintLine sRepNum
      oTextFile.PrintLine "                                  DELEGACION   : " & sGL
      oTextFile.PrintLine "=========================================================================================================================="
      oTextFile.PrintLine "  FECHA DE  NUMERO DE OFICIO                                                                                     "
      oTextFile.PrintLine "  OFICIO    C.N.B. y V.          AUXILIAR          DESCRIPCION                               IMPORTE                TOTAL "
      oTextFile.PrintLine "--------------------------------------------------------------------------------------------------------------------------"
      'oTextFile.PrintLine
      'nLine = 14
      nLine = 24
      nPage = nPage + 1
    End If
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintHeader7028", Err.Number
End Sub

Private Sub PrintHeader7000(oTextFile As CTextFile, _
                            nLine As Long, nPage As Long, _
                            dCuteDate As Date, _
                            sGL As String, sAccountName As String, _
                            sRepNum As String, _
                            sNameA As String, sPuestoA As String, _
                            sNameB As String, sPuestoB As String)
  Dim sDate       As String
  '**********************************************************
  On Error GoTo ErrHandler
    If (nLine >= cnPageLines) Then
      If nLine <> 100 Then
         PrintSigns oTextFile, sNameA, sPuestoA, sNameB, sPuestoB, _
                    Len("================================================================================================================")
         'oTextFile.PageFeed (nLine - 5)
      End If
      'sDate = " AL " & Day(dCuteDate) & " DE " & UCase(MonthName(Month(dCuteDate))) & " DE " & Year(dCuteDate)
      sDate = " AL " & Format(dCuteDate, "d"" DE ""MMMM"" DE ""yyyy")
      oTextFile.PrintLine "TECOYOTITLA NUM 100             BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C.      FECHA: " & Date
      oTextFile.PrintLine "COLONIA FLORIDA                  " & SpcFill(sAccountName & sDate, 57) & "PAGINA:   " & nPage
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F.          ===================================================="
      oTextFile.PrintLine "BNO-670315-CDO"
      oTextFile.PrintLine sRepNum
      oTextFile.PrintLine "                                  DELEGACION   : " & sGL
      oTextFile.PrintLine "========================================================================================================="
      oTextFile.PrintLine "  FECHA      DESCRIPCION                                      IMPORTE"
      oTextFile.PrintLine "---------------------------------------------------------------------------------------------------------"
      nLine = 20
      nPage = nPage + 1
    End If
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintHeader7000", Err.Number
End Sub



'********************************************************************************
'
'      Útilerias Generales
'
'********************************************************************************


Private Sub PrintFeed(oTextFile As CTextFile, sAccount As String)
  On Error GoTo ErrorHandler
    Select Case sAccount
      Case "1105", "1106", "1308", "1310", "1311"
        oTextFile.PrintNewLine
        oTextFile.PrintLine " CERTIFICAMOS QUE NO INCLUYE DOCUMENTOS CON 15 DIAS O MAS DE VENCIDOS NI TITULOS "
        oTextFile.PrintLine " CUYO PLAZO DE OPERACION NO SE AJUSTE A LOS MINIMOS Y MAXIMOS QUE PERMITE LA LEY "
        oTextFile.PrintNewLine
      Case "5105"
        oTextFile.PrintNewLine
        oTextFile.PrintLine " CERTIFICAMOS QUE FUE RETENIDO Y PAGADO EL IMPUESTO SOBRE LA RENTA DERIVADO DE LA PRESTACION DE "
        oTextFile.PrintLine " SERVICIOS PERSONALES SUBORDINADOS Y EN TODOS LOS CASOS SE CUBRIO EL SALARIO MINIMO BANCARIO "
        oTextFile.PrintNewLine
      Case "5106"
        oTextFile.PrintNewLine
        oTextFile.PrintLine " CERTIFICAMOS QUE OBRAN EN NUESTRO PODER LOS RECIBOS CORRESPONDIENTES CON TODOS "
        oTextFile.PrintLine " LOS REQUISITOS QUE MARCA LA LEY "
        oTextFile.PrintNewLine
      Case "5108", "5109"
        oTextFile.PrintNewLine
        oTextFile.PrintLine " CERTIFICAMOS QUE EN TODOS LOS CASOS SE RETUVO Y ENTERO EL IMPUESTO CORRESPONDIENTE Y "
        oTextFile.PrintLine " QUE OBRAN EN NUESTRO PODER LOS RECIBOS CON TODOS LOS REQUERIMIENTOS QUE MARCA LA LEY "
        oTextFile.PrintNewLine
      Case "5110", "5111"
        oTextFile.PrintNewLine
        oTextFile.PrintLine " CERTIFICAMOS QUE OBRAN EN NUESTRO PODER LOS COMPROBANTES CON TODOS LOS REQUERIMIENTOS "
        oTextFile.PrintLine " QUE MARCA LA LEY "
        oTextFile.PrintNewLine
    End Select
    
Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintFeed", Err.Number
End Sub

Private Sub PrintSigns(oTextFile As CTextFile, sNameA As String, sPuestoA As String, _
                        sNameB As String, sPuestoB As String, nLen As Long)
  Dim nNumSpcName        As Long
  Dim nNumSpcPos         As Long
  On Error GoTo ErrorHandler
    nNumSpcName = Round((nLen - (Len(sNameA) + Len(sNameB))) / 3)
    nNumSpcPos = Round((nLen - (Len(sPuestoA) + Len(sPuestoB))) / 3)
    oTextFile.PrintNewLine
    oTextFile.PrintNewLine
    oTextFile.PrintNewLine
    oTextFile.PrintNewLine
    oTextFile.PrintNewLine
    oTextFile.PrintNewLine
    oTextFile.PrintLine Space(nNumSpcName) & sNameA & Space(nNumSpcName) & sNameB
    oTextFile.PrintLine Space(nNumSpcPos) & sPuestoA & Space(nNumSpcPos) & sPuestoB
    oTextFile.PrintNewLine
    oTextFile.PrintNewLine
Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintSigns", Err.Number
End Sub

Private Function SpcFill(ByVal sStr As String, nLen As Long) As String
  On Error GoTo ErrorHandler
    If (Len(sStr) < nLen) Then
      SpcFill = sStr & Space(nLen - Len(sStr))
    Else
      SpcFill = Mid(sStr, 1, nLen)
    End If
  Exit Function
ErrorHandler:
  RaiseError ClassId, "SpcFill", Err.Number
End Function

Private Function SpcFillAmount(ByVal cAmount As String, _
                               nLen As Long, _
                               Optional nHowManyDec As Long = 2) As String
  Dim sFmtAmount As String
  Dim sSign      As String
  '************************************************************************
  On Error GoTo ErrorHandler
    sSign = " "
    If (cAmount < 0) Then
      cAmount = cAmount * -1
      sSign = "-"
    End If
    sFmtAmount = FormatNumber(cAmount, nHowManyDec)
    If (nLen - Len(sFmtAmount)) > 0 Then
      SpcFillAmount = Space(nLen - Len(sFmtAmount)) & sFmtAmount & sSign
    Else
      SpcFillAmount = sFmtAmount & sSign
    End If
  Exit Function
ErrorHandler:
  RaiseError ClassId, "SpcFillAmount", Err.Number
End Function

Private Sub DivideAcc(sNumCuenta As String, _
                      sAccount As String, _
                      sSubAccount As String, _
                      sSubSubAcc As String)
  Dim nLen As Long
  Dim nPos As Long
  '**************************************************************************************
  On Error GoTo ErrorHandler
  sAccount = Mid(sNumCuenta, 1, 4)
  If (Len(sNumCuenta) > 8) Then
    sSubAccount = Mid(sNumCuenta, 6, 2)
    sSubSubAcc = Mid(sNumCuenta, 9, 2)
'    If sSubSubAcc = "00" Then
'      sSubSubAcc = ""
'    End If
  ElseIf (Len(sNumCuenta) <= 8) And (Len(sNumCuenta) > 4) Then
    sSubAccount = Mid(sNumCuenta, 6, 2)
    sSubSubAcc = ""
  End If
  Exit Sub
ErrorHandler:
  RaiseError ClassId, "DivideAcc", Err.Number
End Sub

Private Function FormatLine(nLineLength As Long, sLeftString As String, _
                           sCenterString As String, sRightString As String)
  Dim sTemp As String, n1 As Long, n2 As Long
  '************************************************************************
  If (Len(sLeftString) <> 0) And (Len(sCenterString) <> 0) And (Len(sRightString) <> 0) Then
    If Len(sLeftString & sCenterString & sRightString) < nLineLength Then
      n1 = (nLineLength - Len(sCenterString)) \ 2 - Len(sLeftString)
      n2 = nLineLength - (Len(sLeftString & sCenterString & sRightString)) - n1
    Else
      n1 = 2
      n2 = 2
      sRightString = Mid(sRightString, Len(sRightString) - 2)
      sLeftString = Mid(sLeftString, Len(sLeftString) - 2)
    End If
    sTemp = sLeftString & Space(IIf(n1 < 0, 1, n1)) & sCenterString & Space(n2) & sRightString
  End If
  
  If (Len(sLeftString) = 0) And (Len(sCenterString) <> 0) And (Len(sRightString) = 0) Then
    If Len(sCenterString) < nLineLength Then
      n1 = (nLineLength - Len(sCenterString)) \ 2
      n2 = n1 + (n1 Mod 2)
    Else
      n1 = 2
      n2 = 2
      sCenterString = Mid(sCenterString, nLineLength - 4)
    End If
    sTemp = Space(n1) & sCenterString & Space(n2)
  End If
  
  If (Len(sLeftString) <> 0) And (Len(sCenterString) = 0) And (Len(sRightString) <> 0) Then
    If Len(sLeftString & sRightString) < nLineLength Then
      n1 = (nLineLength - Len(sLeftString & sRightString))
      'n2 = n1 + (n1 Mod 2)
    Else
      n1 = 2
      'n2 = 2
      sRightString = Mid(sRightString, Len(sRightString) - 2)
    End If
    sTemp = sLeftString & Space(n1) & sRightString
  End If
  FormatLine = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "FormatLine", Err.Number
End Function

Private Function GetSubCuenta1(ByVal sNumCuenta As String, _
                               Optional sCuenta As String = "", Optional sSubcuenta As String = "") As String
  Dim nLen As Long
  Dim nPos As Long
  '***********************
  On Error GoTo ErrorHandler
  sCuenta = Mid(sNumCuenta, 1, 4)
  sSubcuenta = Mid(sNumCuenta, 6, 17)
  Do
    If (InStr(sSubcuenta, "-") = 0) Then
      Exit Do
    End If
    nPos = InStr(sSubcuenta, "-")
    sSubcuenta = Mid(sSubcuenta, 1, nPos - 1) & Mid(sSubcuenta, nPos + 1, 17)
  Loop
  nLen = Len(sSubcuenta)
  If (nLen < 12) Then
    sSubcuenta = sSubcuenta & Mid("000000000000", 1, 12 - nLen)
  End If
  
  GetSubCuenta1 = sSubcuenta
  Exit Function
ErrorHandler:
  RaiseError ClassId, "GetSubcuenta", Err.Number
End Function

Private Function GetCuentaSubCuenta(ByVal sNumCuenta As String, _
                                   Optional sCuenta As String = "", Optional sSubcuenta As String = "") As String
  Dim nLen As Long
  Dim nPos As Long
  '***********************
  On Error GoTo ErrorHandler
  sCuenta = Mid(sNumCuenta, 1, 4)
  sSubcuenta = Mid(sNumCuenta, 6, 17)
  Do
    If (InStr(sSubcuenta, "-") = 0) Then
      Exit Do
    End If
    nPos = InStr(sSubcuenta, "-")
    sSubcuenta = Mid(sSubcuenta, 1, nPos - 1) & Mid(sSubcuenta, nPos + 1, 17)
  Loop
  nLen = Len(sSubcuenta)
  If (nLen < 12) Then
    sSubcuenta = sSubcuenta & Mid("000000000000", 1, 12 - nLen)
  End If
  
  GetCuentaSubCuenta = sCuenta & " " & sSubcuenta
  Exit Function
ErrorHandler:
  RaiseError ClassId, "GetCuentaSubcuenta", Err.Number
End Function

Private Function GetSubCuenta(ByVal sNumCuenta As String, _
                              Optional sCuenta As String = "", Optional sSubcuenta As String = "") As String
  Dim nLen As Long
  Dim nPos As Long
  '***********************
  On Error GoTo ErrorHandler
  sCuenta = Mid(sNumCuenta, 1, 4)
  sSubcuenta = Mid(sNumCuenta, 6, 17)
  Do
    If (InStr(sSubcuenta, "-") = 0) Then
      Exit Do
    End If
    nPos = InStr(sSubcuenta, "-")
    sSubcuenta = Mid(sSubcuenta, 1, nPos - 1) & Mid(sSubcuenta, nPos + 1, 17)
  Loop
  nLen = Len(sSubcuenta)
  If (nLen < 12) Then
    sSubcuenta = sSubcuenta & Mid("000000000000", 1, 12 - nLen)
  End If
  
  GetSubCuenta = sSubcuenta
  Exit Function
ErrorHandler:
  RaiseError ClassId, "GetSubCuenta", Err.Number
End Function

Private Function ZeroFill(ByVal sStr As String, nLen As Long) As String
  On Error GoTo ErrorHandler
    ZeroFill = Mid("00000000000000000000000", 1, nLen - Len(sStr)) & sStr
  Exit Function
ErrorHandler:
  RaiseError ClassId, "ZeroFill", Err.Number
End Function


'Private Function CantidadConLetra(ByVal Cantidad As Currency) As String
'  Const cteMilMillones = 1000000000
'  Dim Cadena As String
'  Dim Centavos As String
'  Dim Pesos As Long
'  Dim Temporal As Single
'
'  On Error GoTo ErrCantidadConLetra
'    If Cantidad >= cteMilMillones Then
'      ErrHandler aoxCantidadExcesiva, ClassId, "CantidadConLetra"
'    End If
'    Cantidad = Format$(Cantidad, "Currency")
'    If (Cantidad - Fix(Cantidad)) = 0 Then   'Cantidad - Fix(Cantidad) = Parte fraccionaria
'      Centavos = "00"
'    Else
'      Centavos = Trim$(Str$((Cantidad - Fix(Cantidad)) * 100))
'      If (Cantidad - Fix(Cantidad)) < 0.1 Then
'        Centavos = "0" & Centavos
'      End If
'    End If
'
'    Cadena = "("
'    Pesos = Fix(Cantidad)                              'Trunca la cantidad
'    If Pesos = 0 Then
'      If Centavos <> "00" Then
'        Cadena = Cadena & "Cero pesos " & Centavos & "/100 M.N.)"
'      ElseIf Centavos = "00" Then
'        Cadena = "**Sin Valor  Sin Valor  Sin Valor  Sin Valor**"
'      End If
'    ElseIf Pesos = 1 Then
'      Cadena = Cadena & "Un peso " & Centavos & "/100 M.N.)"
'    Else
'      If (0 < Pesos And Pesos < 1000) Then
'        Cadena = Cadena & CentenasConLetra(Pesos)
'      ElseIf 1000 <= Pesos And Pesos < 1000000 Then
'        Temporal = Pesos \ 1000
'        Cadena = Cadena & CentenasConLetra(Temporal) & " mil " & CentenasConLetra(Pesos Mod 1000)
'      ElseIf 1000000 <= Pesos And Pesos < 1000000000 Then
'        If Pesos < 2000000 Then
'          Cadena = Cadena & "un millón "
'        Else
'          Temporal = Pesos \ 1000000
'          Cadena = Cadena & CentenasConLetra(Temporal) & " millones "
'        End If
'        Temporal = Pesos \ 1000
'        Temporal = Temporal Mod 1000
'        If Temporal <> 0 Then                               'Si hay miles
'          Cadena = Cadena & CentenasConLetra(Temporal) & " mil "
'        End If
'        If (Pesos Mod 1000) <> 0 Then                       'Si hay centenas
'          Cadena = Cadena & CentenasConLetra(Pesos Mod 1000)
'        End If
'        If (Pesos Mod 1000 = 0) And (Temporal = 0) Then     'No hay ni centenas ni miles
'          Cadena = Cadena & " de "
'        End If
'      End If
'      Cadena = Cadena & " pesos " & Centavos & "/100 M.N.)"
'    End If
'    Cadena = SuprimeBlancos(Cadena)
'    CantidadConLetra = UCase$(Left$(Cadena, 2)) & Mid$(Cadena, 3)
'  Exit Function
'ErrCantidadConLetra:
'  ErrHandler aoiErrVB, ClassId, "CantidadConLetra"
'End Function
'
''//Métodos privados de la clase
'Private Function CentenasConLetra(ByVal Numero As Currency) As String
'  Dim Cadena As String
'  Static Tabla(1 To 9) As String
'  Dim Auxiliar As Single
'  Dim Temporal As Single
'
'  'Definición de la tabla
'  If Len(Tabla(1)) = 0 Then
'    Tabla(1) = "un     diez      ciento        once       veintiún"
'    Tabla(2) = "dos    veinte    doscientos    doce       veintidós"
'    Tabla(3) = "tres   treinta   trescientos   trece      veintitrés"
'    Tabla(4) = "cuatro cuarenta  cuatrocientos catorce    veinticuatro"
'    Tabla(5) = "cinco  cincuenta quinientos    quince     veinticinco"
'    Tabla(6) = "seis   sesenta   seiscientos   dieciséis  veintiséis"
'    Tabla(7) = "siete  setenta   setecientos   diecisiete veintisiete"
'    Tabla(8) = "ocho   ochenta   ochocientos   dieciocho  veintiocho"
'    Tabla(9) = "nueve  noventa   novecientos   diecinueve veintinueve"
'
'    'Tabla(1) = LoadResString(aosRenglonTablaCentenas1)    '"un     diez      ciento        once       veintiún"
'    'Tabla(2) = LoadResString(aosRenglonTablaCentenas2)    '"dos    veinte    doscientos    doce       veintidós"
'    'Tabla(3) = LoadResString(aosRenglonTablaCentenas3)    '"tres   treinta   trescientos   trece      veintitrés"
'    'Tabla(4) = LoadResString(aosRenglonTablaCentenas4)    '"cuatro cuarenta  cuatrocientos catorce    veinticuatro"
'    'Tabla(5) = LoadResString(aosRenglonTablaCentenas5)    '"cinco  cincuenta quinientos    quince     veinticinco"
'    'Tabla(6) = LoadResString(aosRenglonTablaCentenas6)    '"seis   sesenta   seiscientos   dieciséis  veintiséis"
'    'Tabla(7) = LoadResString(aosRenglonTablaCentenas7)    '"siete  setenta   setecientos   diecisiete veintisiete"
'    'Tabla(8) = LoadResString(aosRenglonTablaCentenas8)    '"ocho   ochenta   ochocientos   dieciocho  veintiocho"
'    'Tabla(9) = LoadResString(aosRenglonTablaCentenas9)    '"nueve  noventa   novecientos   diecinueve veintinueve"
'  End If
'
'  Cadena = aoEmptyString
'
'  Select Case Numero
'    Case 0
'    Case 100
'      Cadena = Cadena + "cien"
'    Case 1 To 99, 101 To 1000
'      Auxiliar = Numero \ 100               'Obtiene las centenas
'      If Auxiliar <> 0 Then                 'Si hay centenas
'        Cadena = Cadena & Mid$(Tabla(Auxiliar), 18, 13) & " "
'      End If
'      Auxiliar = Numero Mod 100             'Obtiene las decenas
'      Auxiliar = Auxiliar \ 10
'      Temporal = -1
'      Select Case Auxiliar
'        Case 0
'        Case 1                              'Del 11 al 19
'          Temporal = Numero Mod 10
'          If Temporal = 0 Then              'Si es el 10
'            Cadena = Cadena & Mid$(Tabla(1), 8, 6)
'          ElseIf Temporal <> 0 Then
'            Cadena = Cadena & Mid$(Tabla(Temporal), 32, 10)
'          End If
'        Case 2                              'Del 21 al 29
'          Temporal = Numero Mod 10
'          If Temporal = 0 Then              'Si es el 20
'            Cadena = Cadena & Mid$(Tabla(2), 8, 6)
'          ElseIf Temporal <> 0 Then
'            Cadena = Cadena & Mid$(Tabla(Temporal), 43, 12)
'          End If
'        Case Else
'          Cadena = Cadena + Mid$(Tabla(Auxiliar), 8, 9)
'          If Numero Mod 10 <> 0 Then
'            Cadena = Cadena & " y "
'          End If
'      End Select
'      If Temporal = -1 Then                  'No se le asignó nada
'        Auxiliar = Numero Mod 10             'Obtiene las unidades
'        If Auxiliar <> 0 Then
'          Cadena = Cadena & Mid$(Tabla(Auxiliar), 1, 6)
'        End If
'      End If
'  End Select
'  CentenasConLetra = SuprimeBlancos(Cadena)
'End Function


