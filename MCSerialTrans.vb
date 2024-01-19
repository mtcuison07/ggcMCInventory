'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
'
' Copyright 2006 and beyond
' All Rights Reserved
'
'     MC Serial Transaction object
'
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All  rights reserved. No part of this  software  €€  This Software is Owned by        €
' €  may be reproduced or transmitted in any form or  €€                                   €
' €  by   any   means,  electronic   or  mechanical,  €€    GUANZON MERCHANDISING CORP.    €
' €  including recording, or by information  storage  €€     Guanzon Bldg. Perez Blvd.     €
' €  and  retrieval  systems, without  prior written  €€           Dagupan City            €
' €  from the author.                                 €€  Tel No. 522-1085 ; 522-9275      €
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'
' ==========================================================================================
'  XerSys [ 01/15/2007 09:51 am ]
'     Started creating this object based on ggcMCInventory-VB6 edition.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Imports MySql.Data.MySqlClient
Imports ADODB
Imports ggcAppDriver

Public Class MCSerialTrans
    Public Const xeLocWarehouse As String = "0"
    Public Const xeLocBranch As String = "1"
    Public Const xeLocSupplier As String = "2"
    Public Const xeLocCustomer As String = "3"
    Public Const xeLocUnknown As String = "4"
    Public Const xeLocServiceCenter As String = "5"

    Private p_oApp As GRider
    Private p_oDTMstr As DataTable
    Private p_oDTDetl As DataTable

    Private p_sBranchCd As String
    Private p_sDestinat As String
    Private p_sClientID As String
    Private p_dTransact As Date
    Private p_sSourceCd As String
    Private p_sSourceNo As String
    Private p_bWareHous As Boolean
    Private p_bBackLoad As Boolean
    Private p_nEditMode As xeEditMode

    Private p_sCoCltID1 As String
    Private p_sCoCltID2 As String

    'Private p_bInitTran As Boolean

    Public Property AppDriver() As ggcAppDriver.GRider
        Get
            Return p_oApp
        End Get
        Set(value As ggcAppDriver.GRider)
            p_oApp = value
        End Set
    End Property

    Public Property Branch() As String
        Get
            Return p_sBranchCd
        End Get
        Set(value As String)
            p_sBranchCd = value
        End Set
    End Property

    Public Property Detail(Row As Integer, ByVal Index As String) As Object
        Get
            'If Not p_bInitTran Then Return ""
            Return p_oDTMstr(Row).Item(Index)
        End Get

        Set(value As Object)
            'If Not p_bInitTran Then Exit Property
            If Row < 0 Then Exit Property

            If Row > p_oDTMstr.Rows.Count Then
                Exit Property
            ElseIf Row > p_oDTMstr.Rows.Count - 1 Then
                AddDetail()
            End If

            Select Case LCase(Index)
                Case "sserialid"
                    p_oDTMstr(Row).Item(Index) = value
                Case "sbranchcd"
                Case "sengineno", "sframenox"
                    If p_sSourceCd = MCInventoryTrans.pxeMCAcceptance Then
                        p_oDTMstr(Row).Item(Index) = value
                    End If
                Case "smodelidx"
                    p_oDTMstr(Row).Item(Index) = value
                Case "scoloridx"
                    p_oDTMstr(Row).Item(Index) = value
                Case "smcinvidx"
                    p_oDTMstr(Row).Item(Index) = value
                Case "csoldstat"
                    p_oDTMstr(Row).Item(Index) = value
                    p_oDTMstr(Row).Item("cRepoMotr") = value
                Case "clocation"
                Case "cdeliverd"
                    p_oDTMstr(Row).Item(Index) = value
                Case "cregister"
                Case "swarrntno"
                    If p_sSourceCd = MCInventoryTrans.pxeMCSales Then p_oDTMstr(Row).Item(Index) = value
                Case "scompnyid"
                    p_oDTMstr(Row).Item(Index) = value
                Case "sclientid"
                Case "sfilenoxx", "splatenop", "splatenoh"
                    Select Case p_sSourceCd
                        Case MCInventoryTrans.pxeMCRegister, MCInventoryTrans.pxeMCTransfer, MCInventoryTrans.pxeMCTransferRenewal
                            p_oDTMstr(Row).Item(Index) = value
                    End Select
                Case "nledgerno"
                    p_oDTMstr(Row).Item(Index) = value
                Case "cbackload"
                    p_oDTMstr(Row).Item(Index) = value
                Case Else
                    MsgBox("Invalid Field Name Detected!" & vbCrLf & vbCrLf & _
                             "Please inform the SEG/SSG of Guanzon Group of Companies!", vbCritical, "Warning")
            End Select
        End Set
    End Property

    Public ReadOnly Property ItemCount() As Long
        Get
            Return p_oDTDetl.Rows.Count
        End Get
    End Property

    Function AcceptDelivery(SourceNo As String, _
                            TransDate As Date, _
                            UpdateMode As xeEditMode) As Boolean
        p_sSourceCd = MCInventoryTrans.pxeMCAcceptDelivery
        p_sSourceNo = SourceNo
        p_dTransact = TransDate
        p_nEditMode = UpdateMode

        AcceptDelivery = SaveTransaction
    End Function

    Function AssumedUnit(SourceNo As String, _
                         TransDate As Date, _
                         UpdateMode As xeEditMode) As Boolean
        p_sSourceCd = MCInventoryTrans.pxeMCAssume
        p_sSourceNo = SourceNo
        p_dTransact = TransDate
        p_nEditMode = UpdateMode

        AssumedUnit = SaveTransaction
    End Function

    Function Delivery(SourceNo As String, _
                      TransDate As Date, _
                      Branch As String, _
                      UpdateMode As xeEditMode) As Boolean
        p_sSourceCd = MCInventoryTrans.pxeMCDelivery
        p_sSourceNo = SourceNo
        p_sDestinat = Branch
        p_dTransact = TransDate
        p_nEditMode = UpdateMode

        Delivery = SaveTransaction
    End Function

    Function Purchase(SourceNo As String, _
                      TransDate As Date, _
                      UpdateMode As xeEditMode) As Boolean

        p_sSourceCd = MCInventoryTrans.pxeMCSalesInvoice
        p_sSourceNo = SourceNo
        p_dTransact = TransDate
        p_nEditMode = UpdateMode

        Purchase = SaveTransaction
    End Function

    Function PurchaseReceiving(SourceNo As String, _
                               TransDate As Date, _
                               UpdateMode As xeEditMode, _
                               Optional BackLoad As Object = False) As Boolean
        p_sSourceCd = MCInventoryTrans.pxeMCAcceptance
        p_sSourceNo = SourceNo
        p_dTransact = TransDate
        p_nEditMode = UpdateMode
        p_bBackLoad = BackLoad

        PurchaseReceiving = SaveTransaction
    End Function

    Function PurchaseReturn(SourceNo As String, _
                            TransDate As Date, _
                            UpdateMode As xeEditMode) As Boolean

        p_sSourceCd = MCInventoryTrans.pxeMCPurchaseReturn
        p_sSourceNo = SourceNo
        p_dTransact = TransDate
        p_nEditMode = UpdateMode

        PurchaseReturn = SaveTransaction

    End Function

    Function AcceptBackLoadFromBranch(SourceNo As String, _
                                        TransDate As Date, _
                                        UpdateMode As xeEditMode) As Boolean
        p_sSourceCd = MCInventoryTrans.pxeMCAcceptBackLoad
        p_sSourceNo = SourceNo
        p_dTransact = TransDate
        p_nEditMode = UpdateMode

        AcceptBackLoadFromBranch = SaveTransaction

    End Function

    Function BackLoadFromBranch(SourceNo As String, _
                                  TransDate As Date, _
                                  Branch As String, _
                                  UpdateMode As xeEditMode) As Boolean

        p_sSourceCd = MCInventoryTrans.pxeMCBackLoad
        p_sSourceNo = SourceNo
        p_sDestinat = Branch
        p_dTransact = TransDate

        BackLoadFromBranch = SaveTransaction
    End Function

    Function BackLoadTrucking(SourceNo As String, _
                               TransDate As Date, _
                               UpdateMode As xeEditMode) As Boolean

        p_sSourceCd = MCInventoryTrans.pxeMCBackLoadTrucking
        p_sSourceNo = SourceNo
        p_dTransact = TransDate
        p_nEditMode = UpdateMode

        BackLoadTrucking = SaveTransaction
    End Function

    Function Sales(SourceNo As String, _
                   TransDate As Date, _
                   ClientID As String, _
                   UpdateMode As xeEditMode) As Boolean
        Dim lasClient() As String

        If ClientID = "" Then
            MsgBox("Invalid Client Detected!!!" & vbCrLf & vbCrLf & _
                  "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning")
            Return False
        Else
            lasClient = Split(ClientID, "»")
            If UBound(lasClient) = 2 Then
                p_sCoCltID2 = lasClient(2)
            Else
                p_sCoCltID2 = ""
            End If

            If UBound(lasClient) >= 1 Then
                p_sCoCltID1 = lasClient(1)
            Else
                p_sCoCltID1 = ""
            End If

            p_sClientID = lasClient(0)
        End If

        p_sSourceCd = MCInventoryTrans.pxeMCSales
        p_sSourceNo = SourceNo
        p_dTransact = TransDate
        p_nEditMode = UpdateMode

        Sales = SaveTransaction

    End Function

    Function SalesReturn(SourceNo As String, _
                         TransDate As Date, _
                         UpdateMode As xeEditMode) As Boolean
    
        p_sSourceCd = MCInventoryTrans.pxeMCSalesReturn
        p_sSourceNo = SourceNo
        p_dTransact = TransDate
        p_nEditMode = UpdateMode

        SalesReturn = SaveTransaction

    End Function

    Function SalesReplacement(SourceNo As String, _
                               TransDate As Date, _
                               UpdateMode As xeEditMode) As Boolean

        p_sSourceCd = MCInventoryTrans.pxeMCSalesReplacement
        p_sSourceNo = SourceNo
        p_dTransact = TransDate
        p_nEditMode = UpdateMode

        SalesReplacement = SaveTransaction
    End Function

    Function Impound(SourceNo As String, _
                      TransDate As Date, _
                      UpdateMode As xeEditMode) As Boolean

        p_sSourceCd = MCInventoryTrans.pxeMCImpound
        p_sSourceNo = SourceNo
        p_dTransact = TransDate
        p_nEditMode = UpdateMode

        Impound = SaveTransaction()

    End Function

    Function Release(SourceNo As String, _
                      TransDate As Date, _
                      UpdateMode As xeEditMode) As Boolean

        p_sSourceCd = MCInventoryTrans.pxeMCRelease
        p_sSourceNo = SourceNo
        p_dTransact = TransDate
        p_nEditMode = UpdateMode
        Release = SaveTransaction
    End Function

    Function WarrantyPullOut(SourceNo As String, _
                               TransDate As Date, _
                               UpdateMode As xeEditMode) As Boolean

        p_sSourceCd = MCInventoryTrans.pxeMCWarrantyPullOut
        p_sSourceNo = SourceNo
        p_dTransact = TransDate
        p_nEditMode = UpdateMode

        WarrantyPullOut = SaveTransaction

    End Function

    Function WarrantyRelease(SourceNo As String, _
                               TransDate As Date, _
                               UpdateMode As xeEditMode) As Boolean

        p_sSourceCd = MCInventoryTrans.pxeMCWarrantyRelease
        p_sSourceNo = SourceNo
        p_dTransact = TransDate
        p_nEditMode = UpdateMode

        WarrantyRelease = SaveTransaction
    End Function

    Function Register(TransDate As Date, UpdateMode As xeEditMode) As Boolean

        p_sSourceCd = MCInventoryTrans.pxeMCRegister
        p_sSourceNo = ""
        p_dTransact = TransDate
        p_nEditMode = UpdateMode

        Register = SaveTransaction
    End Function

    Function Transfer(SourceNo As String, _
                         TransDate As Date, _
                         ClientID As String, _
                         UpdateMode As xeEditMode) As Boolean
        Dim lasClient() As String

        p_sSourceCd = MCInventoryTrans.pxeMCTransfer
        p_sSourceNo = SourceNo
        p_dTransact = TransDate

        lasClient = Split(ClientID, "»")
        If UBound(lasClient) = 2 Then
            p_sCoCltID2 = lasClient(2)
        Else
            p_sCoCltID2 = ""
        End If

        If UBound(lasClient) >= 1 Then
            p_sCoCltID1 = lasClient(1)
        Else
            p_sCoCltID1 = ""
        End If

        p_sClientID = lasClient(0)

        p_nEditMode = UpdateMode

        Transfer = SaveTransaction
    End Function

    Function TransferRenewal(SourceNo As String, _
                         TransDate As Date, _
                         ClientID As String, _
                         UpdateMode As xeEditMode) As Boolean
        Dim lasClient() As String

        p_sSourceCd = MCInventoryTrans.pxeMCTransferRenewal
        p_sSourceNo = SourceNo
        p_dTransact = TransDate
        lasClient = Split(ClientID, "»")
        If UBound(lasClient) = 2 Then
            p_sCoCltID2 = lasClient(2)
        Else
            p_sCoCltID2 = ""
        End If

        If UBound(lasClient) >= 1 Then
            p_sCoCltID1 = lasClient(1)
        Else
            p_sCoCltID1 = ""
        End If

        p_sClientID = lasClient(0)
        p_nEditMode = UpdateMode

        TransferRenewal = SaveTransaction
    End Function

    Function AcceptClearance(SourceNo As String, _
                            TransDate As Date, _
                            UpdateMode As xeEditMode) As Boolean

        p_sSourceCd = MCInventoryTrans.pxeMCClearRecvd
        p_sSourceNo = SourceNo
        p_dTransact = TransDate
        p_nEditMode = UpdateMode

        AcceptClearance = SaveTransaction
    End Function

    Function InitTransaction() As Boolean
        Call CreateTransaction()
        If AddDetail() = False Then Return False

        Call getBranch()

        InitTransaction = True
    End Function

    Private Function SaveTransaction() As Boolean
        Dim lnUnitInx As Integer, lnUnitOut As Integer
        Dim lnMstrRow As Long

        Select Case p_nEditMode
            Case xeEditMode.MODE_ADDNEW, xeEditMode.MODE_DELETE
            Case Else
                MsgBox("Invalid Update Mode Detected!", vbCritical, "Warning")
                Return False
        End Select

        Try
            If LoadTransaction() = False Then Return False

            ' before saving transaction create a temporary table first to
            '  be used in saving inventory
            If delTmpInventory() = False Then Return False

            If p_sSourceCd = MCInventoryTrans.pxeMCAcceptance Then
                Return SaveAcceptance()
            End If

            If p_nEditMode = xeEditMode.MODE_DELETE Then
                Return DeleteTransaction()
            End If

            With p_oDTMstr

                lnMstrRow = 0
                Do While .Rows.Count > lnMstrRow
                    lnUnitInx = 0
                    lnUnitOut = 0
                    Select Case p_sSourceCd
                        Case MCInventoryTrans.pxeMCAcceptDelivery
                            .Rows(lnMstrRow).Item("cLocation") = IIf(p_bWareHous, xeLocWarehouse, xeLocBranch)
                            .Rows(lnMstrRow).Item("sBranchCd") = p_sBranchCd
                            lnUnitInx = 1
                        Case MCInventoryTrans.pxeMCAcceptBackLoad
                            .Rows(lnMstrRow).Item("cLocation") = IIf(p_bWareHous, xeLocWarehouse, xeLocBranch)
                            .Rows(lnMstrRow).Item("sBranchCd") = p_sBranchCd
                            lnUnitInx = 1
                        Case MCInventoryTrans.pxeMCDelivery, MCInventoryTrans.pxeMCBackLoad
                            .Rows(lnMstrRow).Item("sBranchCd") = p_sDestinat
                            .Rows(lnMstrRow).Item("cLocation") = xeLocUnknown
                            .Rows(lnMstrRow).Item("cDeliverd") = 1
                            lnUnitOut = 1
                        Case MCInventoryTrans.pxeMCPurchaseReturn, MCInventoryTrans.pxeMCBackLoadTrucking
                            .Rows(lnMstrRow).Item("cLocation") = xeLocSupplier
                            lnUnitOut = 1
                        Case MCInventoryTrans.pxeMCSales
                            .Rows(lnMstrRow).Item("sBranchCd") = p_sBranchCd
                            .Rows(lnMstrRow).Item("cLocation") = xeLocCustomer
                            .Rows(lnMstrRow).Item("cSoldStat") = GRider.xeLogical_YES
                            .Rows(lnMstrRow).Item("sClientID") = p_sClientID
                            .Rows(lnMstrRow).Item("sCoCltID1") = p_sCoCltID1
                            .Rows(lnMstrRow).Item("sCoCltID2") = p_sCoCltID2

                            lnUnitOut = 1
                        Case MCInventoryTrans.pxeMCRelease
                            .Rows(lnMstrRow).Item("sBranchCd") = p_sBranchCd
                            .Rows(lnMstrRow).Item("cLocation") = xeLocCustomer
                            lnUnitOut = 1
                        Case MCInventoryTrans.pxeMCImpound
                            .Rows(lnMstrRow).Item("sBranchCd") = p_sBranchCd
                            .Rows(lnMstrRow).Item("cLocation") = IIf(p_bWareHous, xeLocWarehouse, xeLocBranch)
                            lnUnitInx = 1
                        Case MCInventoryTrans.pxeMCSalesReturn, MCInventoryTrans.pxeMCSalesReplacement
                            .Rows(lnMstrRow).Item("sBranchCd") = p_sBranchCd
                            .Rows(lnMstrRow).Item("cLocation") = xeLocBranch
                            .Rows(lnMstrRow).Item("cSoldStat") = .Rows(lnMstrRow).Item("cRegister")
                            lnUnitInx = 1
                        Case MCInventoryTrans.pxeMCWarrantyRelease
                            .Rows(lnMstrRow).Item("cLocation") = xeLocCustomer
                        Case MCInventoryTrans.pxeMCWarrantyPullOut
                            .Rows(lnMstrRow).Item("sBranchCd") = p_sBranchCd
                            .Rows(lnMstrRow).Item("cLocation") = xeLocBranch
                        Case MCInventoryTrans.pxeMCAssume
                            .Rows(lnMstrRow).Item("sBranchCd") = p_sBranchCd
                            .Rows(lnMstrRow).Item("cLocation") = xeLocCustomer
                            .Rows(lnMstrRow).Item("sClientID") = p_sClientID
                            .Rows(lnMstrRow).Item("sCoCltID1") = p_sCoCltID1
                            .Rows(lnMstrRow).Item("sCoCltID2") = p_sCoCltID2
                        Case MCInventoryTrans.pxeMCRegister
                            .Rows(lnMstrRow).Item("cRegistrd") = .Rows(lnMstrRow).Item("sFileNoxx") <> ""
                        Case MCInventoryTrans.pxeMCTransfer, MCInventoryTrans.pxeMCTransferRenewal
                            .Rows(lnMstrRow).Item("sClientID") = p_sClientID
                        Case MCInventoryTrans.pxeMCAdjustment
                            .Rows(lnMstrRow).Item("sBranchCd") = p_sBranchCd
                            lnUnitInx = 1
                    End Select

                    If addSerialTrans(lnMstrRow) = False Then Return False
                    If addTmpSerial(lnMstrRow, lnUnitInx, lnUnitOut) = False Then Return False
                    lnMstrRow = lnMstrRow + 1
                Loop
            End With

            SaveTransaction = saveInventory()
        Catch ex As Exception
            Throw ex
            Return False
        End Try

    End Function

    Private Function DeleteTransaction() As Boolean
        Dim lnDetlRow As Integer
        Try
            With p_oDTDetl
                lnDetlRow = 0
                Do While .Rows.Count > lnDetlRow
                    If delSerialTrans(lnDetlRow) = False Then Return False
                    lnDetlRow = lnDetlRow + 1
                Loop
            End With

            DeleteTransaction = deleteInventory()
        Catch ex As Exception
            Throw ex
            Return False
        End Try
    End Function

    Private Function DeleteAcceptance() As Boolean
        Dim lnDetlRow As Integer
        Try
            With p_oDTDetl
                lnDetlRow = 0
                Do While .Rows.Count > lnDetlRow
                    If delSerial(lnDetlRow) = False Then Return False

                    lnDetlRow = lnDetlRow + 1
                Loop
            End With

            DeleteAcceptance = deleteInventory()

        Catch ex As Exception
            Throw ex
            Return False
        End Try
    End Function

    Private Function saveInventory() As Boolean
        Dim loInventory As MCInventoryTrans
        Dim lsSQL As String
        Dim lnCtr As Integer
        Dim lnRow As Integer

        Try
            ' initialize all object to be used by this class
            loInventory = New MCInventoryTrans(p_oApp)
            With loInventory
                .Branch = p_sBranchCd
                .InitTransaction(p_sSourceCd)
            End With

            lsSQL = "SELECT sMCInvIDx" & _
                        ", cSoldStat" & _
                        ", SUM(nUnitInxx) xUnitInxx" & _
                        ", SUM(nUnitOutx) xUnitOutx" & _
                     " FROM tmpMC_Serial_Transaction" & _
                     " WHERE sSourceNo = " & strParm(p_sSourceNo) & _
                        " AND sSourceCd = " & strParm(p_sSourceCd) & _
                     " GROUP BY sMCInvIDx" & _
                        ", cSoldStat" & _
                     " ORDER BY sMCInvIDx"

            Dim loDta As DataTable
            loDta = p_oApp.ExecuteQuery(lsSQL)
            With loDta
                lnCtr = -1
                lnRow = 0
                Do While .Rows.Count > lnRow
                    If lsSQL <> loDta(lnRow).Item("sMCInvIDx") Then
                        lsSQL = loDta(lnRow).Item("sMCInvIDx")
                        lnCtr = lnCtr + 1
                        With loInventory
                            .Detail(lnCtr, "sMCInvIDx") = lsSQL
                            .Detail(lnCtr, "nQtyInxxx") = 0
                            .Detail(lnCtr, "nQtyInxxx") = 0
                            .Detail(lnCtr, "nRpoQtyIn") = 0
                            .Detail(lnCtr, "nRpoQtyOt") = 0
                        End With
                    End If

                    With loInventory
                        If loDta(lnRow).Item("cSoldStat") = GRider.xeLogical_YES Then
                            .Detail(lnCtr, "nRpoQtyIn") = .Detail(lnCtr, "nRpoQtyIn") + loDta(lnRow).Item("xUnitInxx")
                            .Detail(lnCtr, "nRpoQtyOt") = .Detail(lnCtr, "nRpoQtyOt") + loDta(lnRow).Item("xUnitOutx")
                        Else
                            .Detail(lnCtr, "nQtyInxxx") = .Detail(lnCtr, "nQtyInxxx") + loDta(lnRow).Item("xUnitInxx")
                            .Detail(lnCtr, "nQtyOutxx") = .Detail(lnCtr, "nQtyOutxx") + loDta(lnRow).Item("xUnitOutx")
                        End If
                    End With

                    lnRow = lnRow + 1
                Loop
            End With

            saveInventory = loInventory.SaveTransaction(p_sSourceNo, p_dTransact)
        Catch ex As Exception
            Throw ex
            Return False
        End Try

    End Function


    Private Function SaveAcceptance() As Boolean
        Try
            If p_nEditMode = xeEditMode.MODE_DELETE Then
                Return DeleteAcceptance()
            End If

            With p_oDTMstr
                Dim lnMstrRow As Integer
                lnMstrRow = 0
                Do While .Rows.Count > lnMstrRow
                    If .Rows(lnMstrRow).Item("sSerialID") = "" Then Exit Do

                    p_oDTMstr(lnMstrRow).Item("sBranchCd") = p_sBranchCd
                    If addSerialTrans(lnMstrRow) = False Then Return False
                    If addTmpSerial(lnMstrRow, 1, 0) = False Then Return False

                    lnMstrRow = lnMstrRow + 1
                Loop
            End With

            SaveAcceptance = saveInventory()
        Catch ex As Exception
            Throw ex
            Return False
        End Try
    End Function

    Private Function LoadTransaction() As Boolean
        'Test if the object was initialized
        'If p_bInitTran = False Then
        '    MsgBox("Object is not initialized!!!" & vbCrLf & vbCrLf & _
        '          "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning")
        '    Return False
        'End If

        Try
            Dim lsSQL As String
            lsSQL = "SELECT" & _
                        "  a.sSerialID" & _
                        ", a.sBranchCd" & _
                        ", a.sEngineNo" & _
                        ", a.sFrameNox" & _
                        ", a.sModelIDx" & _
                        ", a.sColorIDx" & _
                        ", a.sMCInvIDx" & _
                        ", a.cSoldStat" & _
                        ", a.cLocation" & _
                        ", a.cDeliverd" & _
                        ", a.cRegister" & _
                        ", a.sWarrntNo" & _
                        ", a.sCompnyID" & _
                        ", a.sClientID" & _
                        ", a.sFileNoxx" & _
                        ", a.sPlateNoP" & _
                        ", a.sPlateNoH" & _
                        ", b.sSourceNo" & _
                        ", b.sSourceCd" & _
                        ", a.nLedgerNo" & _
                        ", b.nLedgerNo xLedgerNo" & _
                        ", b.dTransact" & _
                        ", IFNULL(a.sCoCltID1, '') sCoCltID1" & _
                        ", IFNULL(a.sCoCltID2, '') sCoCltID2" & _
                     " FROM MC_Serial a" & _
                        ", MC_Serial_Ledger b" & _
                     " WHERE a.sSerialID = b.sSerialID"

            If p_nEditMode = xeEditMode.MODE_ADDNEW Then
                lsSQL = lsSQL & _
                         " AND 0 = 1"
            Else
                lsSQL = lsSQL & _
                         " AND b.sSourceNo = " & strParm(p_sSourceNo) & _
                         " AND b.sSourceCd = " & strParm(p_sSourceCd) & _
                      " ORDER BY a.sSerialID"
            End If

            p_oDTDetl = p_oApp.ExecuteQuery(lsSQL)
        Catch ex As Exception
            Throw ex
            Return False
        End Try

        LoadTransaction = True
    End Function

    Private Function deleteInventory() As Boolean
        Dim loInventory As MCInventoryTrans

        Try
            ' initialize all object to be used by this class
            loInventory = New MCInventoryTrans(p_oApp)
            With loInventory
                .Branch = p_sBranchCd
                .InitTransaction(p_sSourceCd)
                deleteInventory = .DeleteTransaction(p_sSourceNo, p_dTransact)
            End With
        Catch ex As Exception
            Throw ex
            Return False
        End Try
    End Function

    Private Function addSerial(lnMstrRow As Integer) As Boolean
        Dim lsSQL As String
        Dim lnRow As Long

        Try
            With p_oDTMstr
                .Rows(lnMstrRow).Item("nLedgerNo") = 0
                .Rows(lnMstrRow).Item("cLocation") = IIf(p_bWareHous, xeLocWarehouse, xeLocBranch)
                lsSQL = "INSERT INTO MC_Serial SET" & _
                            "  sSerialID = " & strParm(.Rows(lnMstrRow).Item("sSerialID")) & _
                            ", sBranchCd = " & strParm(p_sBranchCd) & _
                            ", sEngineNo = " & strParm(.Rows(lnMstrRow).Item("sEngineNo")) & _
                            ", sFrameNox = " & strParm(.Rows(lnMstrRow).Item("sFrameNox")) & _
                            ", sModelIDx = " & strParm(.Rows(lnMstrRow).Item("sModelIDx")) & _
                            ", sColorIDx = " & strParm(.Rows(lnMstrRow).Item("sColorIDx")) & _
                            ", sMCInvIDx = " & strParm(.Rows(lnMstrRow).Item("sMCInvIDx")) & _
                            ", cSoldStat = " & strParm(.Rows(lnMstrRow).Item("cSoldStat")) & _
                            ", cLocation = " & strParm(.Rows(lnMstrRow).Item("cLocation")) & _
                            ", cRegister = " & strParm(.Rows(lnMstrRow).Item("cRegister")) & _
                            ", cCSRValid = " & strParm(.Rows(lnMstrRow).Item("cCSRValid")) & _
                            ", cPNPClear = " & strParm(.Rows(lnMstrRow).Item("cPNPClear")) & _
                            ", sWarrntNo = " & strParm(.Rows(lnMstrRow).Item("sWarrntNo")) & _
                            ", sCompnyID = " & strParm(.Rows(lnMstrRow).Item("sCompnyID")) & _
                            ", sClientID = " & strParm(.Rows(lnMstrRow).Item("sClientID")) & _
                            ", sCoCltID1 = " & strParm(IFNull(.Rows(lnMstrRow).Item("sCoCltID1"))) & _
                            ", sCoCltID2 = " & strParm(IFNull(.Rows(lnMstrRow).Item("sCoCltID2"))) & _
                            ", sFileNoxx = " & strParm(.Rows(lnMstrRow).Item("sFileNoxx")) & _
                            ", sPlateNoP = " & strParm(.Rows(lnMstrRow).Item("sPlateNoP")) & _
                            ", sPlateNoH = " & strParm(.Rows(lnMstrRow).Item("sPlateNoH")) & _
                            ", cDeliverd = " & strParm(.Rows(lnMstrRow).Item("cDeliverd")) & _
                            ", nLedgerNo = " & strParm(Format(.Rows(lnMstrRow).Item("nLedgerNo") + 1, "00")) & _
                            ", sModified = " & strParm(p_oApp.UserID) & _
                            ", dModified = " & dateParm(p_oApp.getSysDate)

                lnRow = p_oApp.Execute(lsSQL, "MC_Serial", p_sBranchCd)
                If lnRow <= 0 Then
                    MsgBox("Unable to Update Motorcycle Serial Transaction!", vbCritical, "Warning")
                    Return False
                End If
            End With
        Catch ex As Exception
            Throw ex
            Return False
        End Try

        addSerial = True
    End Function

    Private Function addSerialTrans(lnMstrRow As Integer) As Boolean
        Dim lsSQL As String
        Dim lnRow As Long

        Try
            With p_oDTMstr
                If p_sSourceCd = MCInventoryTrans.pxeMCAcceptance Then
                    If .Rows(lnMstrRow).Item("cBackLoad") = GRider.xeLogical_YES Then
                        If updateSerial(lnMstrRow) = False Then Return False
                    Else
                        If addSerial(lnMstrRow) = False Then Return False
                    End If
                Else
                    ' nLedgerNo is required
                    If .Rows(lnMstrRow).Item("nLedgerNo") = 0 Then
                        MsgBox("Invalid Ledger Line Detected!", vbCritical, "Warning")
                        Return False
                    End If

                    If updateSerial(lnMstrRow) = False Then Return False
                End If

                lsSQL = "INSERT INTO MC_Serial_Ledger SET" & _
                            "  sSerialID = " & strParm(.Rows(lnMstrRow).Item("sSerialID")) & _
                            ", sBranchCd = " & strParm(.Rows(lnMstrRow).Item("sBranchCd")) & _
                            ", dTransact = " & dateParm(p_dTransact) & _
                            ", sSourceCd = " & strParm(p_sSourceCd) & _
                            ", sSourceNo = " & strParm(p_sSourceNo) & _
                            ", cSoldStat = " & strParm(.Rows(lnMstrRow).Item("cSoldStat")) & _
                            ", cLocation = " & strParm(.Rows(lnMstrRow).Item("cLocation")) & _
                            ", nLedgerNo = " & strParm(Format(.Rows(lnMstrRow).Item("nLedgerNo") + 1, "00")) & _
                            ", dModified = " & dateParm(p_oApp.getSysDate)

                lnRow = p_oApp.Execute(lsSQL, "MC_Serial_Ledger", p_sBranchCd)
                If lnRow <= 0 Then
                    MsgBox("Unable to Update Motorcycle Serial Transaction!", , "Warning")
                    Return False
                End If
            End With
        Catch ex As Exception
            Throw ex
            Return False
        End Try

        addSerialTrans = True
    End Function

    Private Function addTmpSerial(lnMstrRow As Integer, lnUnitInxx As Integer, lnUnitOutx As Integer) As Boolean
        Dim lsSQL As String
        Dim lnRow As Long

        Try
            With p_oDTMstr
                lsSQL = "INSERT INTO tmpMC_Serial_Transaction SET" & _
                            "  sSourceNo = " & strParm(p_sSourceNo) & _
                            ", sSourceCd = " & strParm(p_sSourceCd) & _
                            ", sSerialID = " & strParm(.Rows(lnMstrRow).Item("sSerialID")) & _
                            ", nUnitInxx = " & lnUnitInxx & _
                            ", nUnitOutx = " & lnUnitOutx & _
                            ", cSoldStat = " & strParm(.Rows(lnMstrRow).Item("cRepoMotr")) & _
                            ", sMCInvIDx = " & strParm(.Rows(lnMstrRow).Item("sMCInvIDx"))

                lnRow = p_oApp.ExecuteActionQuery(lsSQL)
                If lnRow <= 0 Then
                    MsgBox("Unable to Update Motorcycle Serial Transaction!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Warning")
                    Return False
                End If
            End With
        Catch ex As Exception
            Throw ex
            Return False
        End Try

        addTmpSerial = True
    End Function

    Private Function delSerial(lnDetlRow As Integer) As Boolean
        Dim lsSQL As String
        Dim lnRow As Long

        Try
            If delSerialTrans(lnDetlRow) = False Then Return False

            With p_oDTDetl
                If .Rows(lnDetlRow).Item("nLedgerNo") > 1 Then
                    MsgBox("Motorcycle that Has Other Transaction Can Not be Deleted!" & vbCrLf & _
                             "Cancel All its Other Transaction then Try Again!", vbCritical, "Warning")
                    Return False
                End If

                lsSQL = "DELETE FROM MC_Serial" & _
                         " WHERE sSerialID = " & strParm(.Rows(lnDetlRow).Item("sSerialID"))
            End With

            lnRow = p_oApp.Execute(lsSQL, "MC_Serial", p_sBranchCd)
            If lnRow <= 0 Then
                MsgBox("Unable to Delete Motorcycle Serial!", "Warning")
                Return False
            End If
        Catch ex As Exception
            Throw ex
            Return False
        End Try

        delSerial = True
    End Function

    Private Function delSerialTrans(lnDetlRow As Integer) As Boolean
        Dim lsSQL As String
        Dim lnRow As Long

        Try
            With p_oDTDetl
                If .Rows(lnDetlRow).Item("nLedgerNo") <> .Rows(lnDetlRow).Item("xLedgerNo") Then
                    MsgBox("Motorcycle Has Other Transaction!" & vbCrLf & _
                             "Modification is not Allowed!" & vbCrLf & _
                             "Cancel All its Other Transaction then Try Again!", vbCritical, "Warning")
                    Return False
                End If

                lsSQL = "DELETE FROM MC_Serial_Ledger" & _
                       " WHERE sSerialID = " & strParm(.Rows(lnDetlRow).Item("sSerialID")) & _
                         " AND sSourceCd = " & strParm(p_sSourceCd) & _
                         " AND sSourceNo = " & strParm(p_sSourceNo)

                lnRow = p_oApp.Execute(lsSQL, "MC_Serial_Ledger", p_sBranchCd)
                If lnRow <= 0 Then
                    MsgBox("Unable to Delete Motorcycle Serial!", , "Warning")
                    Return False
                End If

                If p_sSourceCd <> MCInventoryTrans.pxeMCAcceptance Then
                    lsSQL = "SELECT" & _
                                "  a.sSerialID" & _
                                ", a.sBranchCd" & _
                                ", a.cLocation" & _
                                ", a.cSoldStat" & _
                                ", a.sSourceCd" & _
                                ", a.sSourceNo" & _
                                ", b.nLedgerNo" & _
                             " FROM MC_Serial_Ledger a" & _
                                ", MC_Serial b" & _
                             " WHERE a.sSerialID = b.sSerialID" & _
                                " AND a.sSerialID = " & strParm(.Rows(lnDetlRow).Item("sSerialID")) & _
                             " ORDER BY a.nLedgerNo DESC" & _
                             " LIMIT 1"

                    Dim loDta As DataTable
                    loDta = p_oApp.ExecuteQuery(lsSQL)

                    With loDta
                        If .Rows.Count <= 0 Then
                            Return False
                        End If

                        'TODO: Check for the possible value of cLocation.
                        '      1 vs loDta(0).Item("cLocation")
                        lsSQL = "UPDATE MC_Serial SET" & _
                                    "  sBranchCd = " & strParm(loDta(0).Item("sBranchCd")) & _
                                    ", cLocation = " & strParm(1) & _
                                    ", cSoldStat = " & strParm(loDta(0).Item("cSoldStat")) & _
                                    ", nLedgerNo = " & strParm(Format(CInt(loDta(0).Item("nLedgerNo")) - 1, "00")) & _
                                    ", dModified = " & dateParm(p_oApp.getSysDate) & _
                                 " WHERE sSerialID = " & strParm(loDta(0).Item("sSerialID"))

                        lnRow = p_oApp.Execute(lsSQL, "MC_Serial", p_sBranchCd)
                        If lnRow <= 0 Then
                            MsgBox(lsSQL & vbCrLf & _
                                     "Unable to Update MC Serial Info!", vbCritical, "Warning")
                            Return False
                        End If
                    End With
                End If
            End With
        Catch ex As Exception
            Throw ex
            Return False
        End Try

        delSerialTrans = True
    End Function

    Private Function delTmpInventory() As Boolean
        Dim lsSQL
        Try
            lsSQL = "DELETE FROM tmpMC_Serial_Transaction" & _
                     " WHERE sSourceNo = " & strParm(p_sSourceNo) & _
                        " AND sSourceCd = " & strParm(p_sSourceCd)
            p_oApp.ExecuteActionQuery(lsSQL)

            delTmpInventory = True
        Catch ex As Exception
            Throw ex
            Return False
        End Try
    End Function

    Private Function updateSerial(lnMstrRow As Integer) As Boolean
        Dim lsSQL As String
        Dim lnRow As Integer

        Try
            With p_oDTMstr
                lsSQL = "UPDATE MC_Serial SET" & _
                            "  sBranchCd = " & strParm(.Rows(lnMstrRow).Item("sBranchCd")) & _
                            ", cLocation = " & strParm(.Rows(lnMstrRow).Item("cLocation")) & _
                            ", nLedgerNo = " & strParm(Format(.Rows(lnMstrRow).Item("nLedgerNo") + 1, "00"))

                ' check if this record has a valid value
                For lnRow = 4 To 9

                    If .Rows(lnMstrRow).Item(lnRow) <> "" Then
                        lsSQL = lsSQL & _
                                 ", " & .Columns(lnRow).ColumnName & " = " & strParm(.Rows(lnMstrRow).Item(lnRow))
                    End If
                Next

                For lnRow = 13 To 18
                    If .Rows(lnMstrRow).Item(lnRow) <> "" Then
                        lsSQL = lsSQL & _
                                 ", " & .Columns(lnRow).ColumnName & " = " & strParm(.Rows(lnMstrRow).Item(lnRow))
                    End If
                Next

                If .Rows(lnMstrRow).Item("sCoCltID1") <> "" Then
                    lsSQL = lsSQL & _
                             ", sCoCltID1 = " & strParm(.Rows(lnMstrRow).Item("sCoCltID1"))
                End If

                If .Rows(lnMstrRow).Item("sCoCltID2") <> "" Then
                    lsSQL = lsSQL & _
                             ", sCoCltID2 = " & strParm(.Rows(lnMstrRow).Item("sCoCltID1"))
                End If

                lsSQL = lsSQL & _
                            ", dModified = " & dateParm(p_oApp.getSysDate) & _
                         " WHERE sSerialID = " & strParm(.Rows(lnMstrRow).Item("sSerialID"))
                Debug.Print(lsSQL)
                lnRow = p_oApp.Execute(lsSQL, "MC_Serial", p_sBranchCd)
                If lnRow <= 0 Then
                    MsgBox("Unable to Update Motorcycle Serial Info!", vbCritical, "Warning")
                    Return False
                End If
            End With
        Catch ex As Exception
            Throw ex
            Return False
        End Try

        updateSerial = True
    End Function

    Private Sub CreateTransaction()
        p_oDTMstr = New DataTable
        p_oDTMstr.Columns.Add("sSerialID", Type.GetType("System.String"))
        p_oDTMstr.Columns.Add("sBranchCd", Type.GetType("System.String"))
        p_oDTMstr.Columns.Add("cLocation", Type.GetType("System.String"))
        p_oDTMstr.Columns.Add("nLedgerNo", Type.GetType("System.Int32"))
        p_oDTMstr.Columns.Add("sEngineNo", Type.GetType("System.String"))
        p_oDTMstr.Columns.Add("sFrameNox", Type.GetType("System.String"))
        p_oDTMstr.Columns.Add("sMCInvIDx", Type.GetType("System.String"))
        p_oDTMstr.Columns.Add("sModelIDx", Type.GetType("System.String"))
        p_oDTMstr.Columns.Add("sColorIDx", Type.GetType("System.String"))
        p_oDTMstr.Columns.Add("cSoldStat", Type.GetType("System.String"))
        p_oDTMstr.Columns.Add("cRegister", Type.GetType("System.String"))
        p_oDTMstr.Columns.Add("cCSRValid", Type.GetType("System.String"))
        p_oDTMstr.Columns.Add("cPNPClear", Type.GetType("System.String"))
        p_oDTMstr.Columns.Add("sWarrntNo", Type.GetType("System.String"))
        p_oDTMstr.Columns.Add("sClientID", Type.GetType("System.String"))
        p_oDTMstr.Columns.Add("sCompnyID", Type.GetType("System.String"))
        p_oDTMstr.Columns.Add("sFileNoxx", Type.GetType("System.String"))
        p_oDTMstr.Columns.Add("sPlateNoP", Type.GetType("System.String"))
        p_oDTMstr.Columns.Add("sPlateNoH", Type.GetType("System.String"))
        p_oDTMstr.Columns.Add("sModified", Type.GetType("System.String"))
        p_oDTMstr.Columns.Add("cDeliverd", Type.GetType("System.String"))
        p_oDTMstr.Columns.Add("cRepoMotr", Type.GetType("System.String"))
        p_oDTMstr.Columns.Add("cBackLoad", Type.GetType("System.String"))
        p_oDTMstr.Columns.Add("sCoCltID1", Type.GetType("System.String"))
        p_oDTMstr.Columns.Add("sCoCltID2", Type.GetType("System.String"))
    End Sub

    Private Function AddDetail() As Boolean
        p_oDTMstr.Rows.Add(p_oDTMstr.NewRow)
        InitRecord(p_oDTMstr.Rows.Count - 1)
        AddDetail = True
    End Function

    Private Sub InitRecord(lnRow As Integer)
        With p_oDTMstr
            .Rows(lnRow).Item("sSerialID") = ""
            .Rows(lnRow).Item("sBranchCd") = ""
            .Rows(lnRow).Item("sEngineNo") = ""
            .Rows(lnRow).Item("sFrameNox") = ""
            .Rows(lnRow).Item("sModelIDx") = ""
            .Rows(lnRow).Item("sColorIDx") = ""
            .Rows(lnRow).Item("sMCInvIDx") = ""
            .Rows(lnRow).Item("cSoldStat") = GRider.xeLogical_NO
            .Rows(lnRow).Item("cLocation") = IIf(p_bWareHous, xeLocWarehouse, xeLocBranch)
            .Rows(lnRow).Item("cRegister") = GRider.xeLogical_NO
            .Rows(lnRow).Item("cCSRValid") = GRider.xeLogical_NO
            .Rows(lnRow).Item("cPNPClear") = GRider.xeLogical_NO
            .Rows(lnRow).Item("sWarrntNo") = ""
            .Rows(lnRow).Item("sCompnyID") = ""
            .Rows(lnRow).Item("sClientID") = ""
            .Rows(lnRow).Item("sFileNoxx") = ""
            .Rows(lnRow).Item("sPlateNoP") = ""
            .Rows(lnRow).Item("sPlateNoH") = ""
            .Rows(lnRow).Item("cDeliverd") = 0
            .Rows(lnRow).Item("nLedgerNo") = 0
            .Rows(lnRow).Item("cBackLoad") = GRider.xeLogical_NO
            .Rows(lnRow).Item("sCoCltID1") = ""
            .Rows(lnRow).Item("sCoCltID2") = ""
        End With
    End Sub

    Private Sub getBranch()
        Dim loDta As DataTable
        Dim lsSQL As String

        lsSQL = "SELECT cWareHous" & _
               " FROM Branch" & _
               " WHERE sBranchCd = " & strParm(p_sBranchCd)
        loDta = p_oApp.ExecuteQuery(lsSQL)

        If loDta.Rows.Count > 0 Then
            p_bWareHous = loDta(0).Item("cWarehous") = GRider.xeLogical_YES
        End If
    End Sub

    Public Sub New(ByVal foRider As GRider)
        p_oApp = foRider
        p_sBranchCd = p_oApp.BranchCode
    End Sub
End Class
