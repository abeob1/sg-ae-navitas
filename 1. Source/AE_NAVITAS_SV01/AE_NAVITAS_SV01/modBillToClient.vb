Module modBillToClient

    Private dtBP As DataTable
    Private dtProject As DataTable
    Private dtItemCode As DataTable
    Private dtAcctCode As DataTable
    Private dtVatGroup As DataTable
    Private dtCusRefNo As DataTable
    Private dtFileName As DataTable
    

    Public Function ProcessBilltoClient(ByVal file As System.IO.FileInfo, ByVal odv As DataView, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = "ProcessBilltoClient"
        Dim sDBCode As String = String.Empty
        Dim sSQL As String = String.Empty

        Try

            sDBCode = odv(1)(0).ToString().Trim()
            Dim k As Integer = InStrRev(sDBCode, ":")
            sDBCode = Microsoft.VisualBasic.Right(sDBCode, Len(sDBCode) - k).Trim

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectToTargetCompany()", sFuncName)
            Console.WriteLine("Connecting Company")
            If ConnectToTargetCompany(p_oCompany, sDBCode, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_oCompany.Connected Then

                sSQL = "SELECT ""CardCode"" FROM ""OCRD"" "
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
                dtBP = ExecuteQueryReturnDataTable(sSQL, p_oCompany.CompanyDB)

                sSQL = "SELECT ""Code"",""Name"" FROM ""@AE_ITEMCODEMAPPING"" "
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
                dtItemCode = ExecuteQueryReturnDataTable(sSQL, p_oCompDef.sSAPDBName)

                sSQL = "SELECT ""OcrCode"" FROM ""OOCR"" "
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING  SQL :" & sSQL, sFuncName)
                dtProject = ExecuteQueryReturnDataTable(sSQL, p_oCompany.CompanyDB)

                sSQL = "SELECT ""AcctCode"" FROM ""OACT"" "
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING  SQL :" & sSQL, sFuncName)
                dtAcctCode = ExecuteQueryReturnDataTable(sSQL, p_oCompany.CompanyDB)

                sSQL = "SELECT ""ItemCode"",""VatGourpSa"" FROM ""OITM"" WHERE ""frozenFor""='N'"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING  SQL :" & sSQL, sFuncName)
                dtVatGroup = ExecuteQueryReturnDataTable(sSQL, p_oCompany.CompanyDB)

                sSQL = "SELECT T0.""GroupCode"" ,T0.""GroupName"" FROM ""OCRG"" T0 WHERE T0.""GroupType""='C' AND T0.""Locked""='N'"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING  SQL :" & sSQL, sFuncName)
                p_oDtBPGroup = ExecuteQueryReturnDataTable(sSQL, p_oCompany.CompanyDB)

                sSQL = "SELECT T0.""GroupNum"", T0.""PymntGroup"" FROM ""OCTG"" T0"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING  SQL :" & sSQL, sFuncName)
                p_oDtPayTerms = ExecuteQueryReturnDataTable(sSQL, p_oCompany.CompanyDB)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddInvoice_CreditNote()", sFuncName)

                Console.WriteLine("Adding Invoice/Credit Memo Document")

                If StartTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                If AddInvoice_CreditNote(file.Name, odv, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CommitTransaction", sFuncName)
                If CommitTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling FileMoveToArchive()", sFuncName)
                'FileMoveToArchive(file, file.FullName, RTN_SUCCESS)
                FileMoveToArchive_Success(file, file.FullName, file.Name, RTN_SUCCESS)

                'Insert Success Notificaiton into Table..
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
                AddDataToTable(p_oDtSuccess, file.Name, "Success")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("File successfully uploaded" & file.FullName, sFuncName)

            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting from SAP", sFuncName)
            ProcessBilltoClient = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message.ToString()
            Call WriteToLogFile(sErrDesc, sFuncName)

            'Insert Error Description into Table
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
            AddDataToTable(p_oDtError, file.Name, "Error", sErrDesc)
            'error condition

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling RollbackTransaction", sFuncName)
            If RollbackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling FileMoveToArchive()", sFuncName)
            FileMoveToArchive(file, file.FullName, RTN_ERROR)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ProcessBilltoClient = RTN_ERROR
        End Try

    End Function

    Private Function AddInvoice_CreditNote(ByVal sFileName As String, ByVal odv As DataView, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = "AddInvoice_CreditNote"

        Dim sBatchNo As String = String.Empty
        Dim sBatchPeriod As String = String.Empty
        Dim sBillType As String = String.Empty
        Dim dGrossTot As Double = 0.0
        Dim sItemCode As String = String.Empty
        Dim sCostCenter As String = String.Empty
        Dim sDBCode As String = String.Empty
        Dim sFullBatchNo As String = String.Empty
        Dim sFullBatchPeriod As String = String.Empty
        Dim sVatGroup As String = String.Empty
        Dim sCardName As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            sDBCode = odv(1)(0).ToString.Trim
            sFullBatchNo = odv(2)(0).ToString.Trim
            sFullBatchPeriod = odv(3)(0).ToString.Trim

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Read DB Code", sFuncName)
            Dim X As Integer = InStrRev(sDBCode, ":")
            sDBCode = Microsoft.VisualBasic.Right(sDBCode, Len(sDBCode) - X).Trim

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Reading Batch No", sFuncName)
            Dim k As Integer = InStrRev(sFullBatchNo, ":")
            sBatchNo = Microsoft.VisualBasic.Right(sFullBatchNo, Len(sFullBatchNo) - k).Trim

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Batch No : " & sBatchNo, sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Reading Batch Period", sFuncName)
            Dim m As Integer = InStrRev(sFullBatchPeriod, "to")
            sBatchPeriod = Microsoft.VisualBasic.Right(sFullBatchPeriod, Len(sFullBatchPeriod) - m - 1).Trim

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Batch Period : " & sBatchPeriod, sFuncName)

            For i As Integer = 6 To odv.Count - 1
                If Not odv(i)(2).ToString = String.Empty Then

                    dGrossTot = odv(i)(18).ToString

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting to create documents", sFuncName)

                    Dim sCardCode As String = String.Empty
                    Dim sFullCardCode As String = String.Empty
                    Dim sAcctCode As String = String.Empty
                    Dim iRetCode, iErrCode As Long
                    Dim oDoc As SAPbobsCOM.Documents = Nothing
                    Dim iCount As Integer = 0
                    Dim bIsError As Boolean = False
                    Dim bIsPanel As Boolean = False
                    Dim sSql As String = String.Empty
                    Dim sNumAtCard As String = String.Empty

                    If dGrossTot > 0.0 Then
                        sSql = "SELECT DISTINCT ""U_AI_APARUploadName"" FROM ""OINV"" WHERE IFNULL(""U_AI_APARUploadName"",'') <> ''"
                    Else
                        sSql = "SELECT DISTINCT ""U_AI_APARUploadName"" FROM ""ORIN"" WHERE IFNULL(""U_AI_APARUploadName"",'') <> ''"
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSql, sFuncName)
                    dtFileName = ExecuteQueryReturnDataTable(sSql, p_oCompany.CompanyDB)

                    dtFileName.DefaultView.RowFilter = "U_AI_APARUploadName = '" & sFileName & "'"
                    If dtFileName.DefaultView.Count > 0 Then
                        sErrDesc = "Interface file ::" & sFileName & " has already been uploaded"
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    End If

                    sFullCardCode = odv(i)(3).ToString.Trim

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("CardCode is " & sFullCardCode, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("CardCode Length is " & sFullCardCode.Length,sFuncName )

                    If sFullCardCode.Length > 15 Then
                        sCardCode = sFullCardCode.Substring(0, 15)
                    Else
                        sCardCode = sFullCardCode
                    End If

                    sAcctCode = odv(i)(5).ToString.Trim
                    'sCardCode = odv(i)(3).ToString.Trim
                    sCostCenter = odv(i)(4).ToString.Trim()
                    sCardName = odv(i)(2).ToString.Trim
                    sNumAtCard = odv(i)(0).ToString

                    dtBP.DefaultView.RowFilter = "CardCode = '" & sCardCode & "'"
                    If dtBP.DefaultView.Count = 0 Then
                        'sErrDesc = "Cardcode ::" & sCardCode & " provided does not exist in SAP."
                        'Call WriteToLogFile(sErrDesc, sFuncName)
                        'Throw New ArgumentException(sErrDesc)

                        If CheckBP(sFullCardCode, sCardCode, sCardName, "C", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                    End If

                    If sCostCenter.ToString() <> String.Empty Then
                        dtProject.DefaultView.RowFilter = "OcrCode = '" & sCostCenter & "'"
                        If dtProject.DefaultView.Count = 0 Then
                            sErrDesc = "Cost Center :: " & sCostCenter & " provided does not exist in SAP."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        End If
                    End If

                    If sAcctCode.ToString() <> String.Empty Then
                        dtAcctCode.DefaultView.RowFilter = "AcctCode = '" & sAcctCode & "'"
                        If dtAcctCode.DefaultView.Count = 0 Then
                            sErrDesc = "Account Code :: " & sAcctCode & " provided does not exist in SAP."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        End If
                    End If

                    If dGrossTot > 0.0 Then
                        Console.WriteLine("Creating A/R Invoice..")
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating A/R Invoice Object", sFuncName)
                        oDoc = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                    Else
                        Console.WriteLine("Creating A/R Credit Memo..")
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating A/R Credit Memo Object", sFuncName)
                        oDoc = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
                    End If

                    oDoc.CardCode = sCardCode
                    oDoc.DocDate = CDate(sBatchPeriod)
                    'oDoc.ImportFileNum = sBatchNo
                    oDoc.UserFields.Fields.Item("U_AI_BatchNo").Value = sBatchNo
                    oDoc.JournalMemo = sFullBatchPeriod

                    If dGrossTot > 0.0 Then
                        sSql = "SELECT DISTINCT ""NumAtCard"" FROM ""OINV"" WHERE IFNULL(""NumAtCard"",'') <> ''"
                    Else
                        sSql = "SELECT DISTINCT ""NumAtCard"" FROM ""ORIN"" WHERE IFNULL(""NumAtCard"",'') <> ''"
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSql, sFuncName)
                    dtCusRefNo = ExecuteQueryReturnDataTable(sSql, p_oCompany.CompanyDB)

                    If Not (sNumAtCard = String.Empty) Then
                        dtCusRefNo.DefaultView.RowFilter = "NumAtCard = '" & sNumAtCard & "'"
                        If dtCusRefNo.DefaultView.Count = 0 Then
                            oDoc.NumAtCard = sNumAtCard
                        Else
                            sErrDesc = "Customer Ref No :: " & sNumAtCard & " already exist in SAP."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        End If
                    End If

                    If Not (odv(i)(1).ToString = String.Empty) Then
                        oDoc.TaxDate = CDate(odv(i)(1).ToString)
                    End If

                    If Not (odv(i)(6).ToString = String.Empty) Then
                        oDoc.Comments = odv(i)(6).ToString
                    End If

                    oDoc.UserFields.Fields.Item("U_AI_APARUploadName").Value = sFileName

                    oDoc.BPL_IDAssignedToInvoice = "1"

                    iCount = iCount + 1

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Before Medical Claims " & odv(i)(13).ToString(), sFuncName)


                    If Not (odv(i)(13).ToString = String.Empty) Then

                        If (CDbl(odv(i)(13).ToString.Trim() <> 0)) Then

                            If iCount > 1 Then
                                oDoc.Lines.Add()
                            End If

                            dtItemCode.DefaultView.RowFilter = "Name = 'Medical Claims'"
                            If dtItemCode.DefaultView.Count = 0 Then
                                sErrDesc = "ItemCode ::'Medical Claims' provided does not exist in SAP(Mapping Table)."
                                Call WriteToLogFile(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            Else
                                sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                            End If

                            dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                            If dtVatGroup.DefaultView.Count = 0 Then
                                sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                                Call WriteToLogFile(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            Else
                                sVatGroup = dtVatGroup.DefaultView.Item(0)(1).ToString().Trim()
                            End If

                            oDoc.Lines.ItemCode = sItemCode
                            oDoc.Lines.UnitPrice = Math.Abs(CDbl(odv(i)(13).ToString.Trim()))

                            If Not (odv(i)(16).ToString.Trim() = String.Empty) Then
                                If (odv(i)(16).ToString.Trim = 0) Then
                                    oDoc.Lines.VatGroup = p_oCompDef.sARZeroRated
                                ElseIf (odv(i)(16).ToString.Trim = 7) Then
                                    oDoc.Lines.VatGroup = p_oCompDef.sARStdRated
                                Else
                                    oDoc.Lines.VatGroup = p_oCompDef.sARZeroRated
                                End If
                            Else
                                oDoc.Lines.VatGroup = p_oCompDef.sARZeroRated
                            End If
                            'oDoc.Lines.VatGroup = sVatGroup
                            If Not (sCostCenter = String.Empty) Then
                                oDoc.Lines.CostingCode = sCostCenter
                                oDoc.Lines.COGSCostingCode = sCostCenter
                            End If
                            If Not (sAcctCode = String.Empty) Then
                                oDoc.Lines.AccountCode = sAcctCode
                            End If

                            bIsPanel = True

                            iCount = iCount + 1
                        End If
                    End If


                    If Not (odv(i)(14).ToString = String.Empty) Then
                        If (CDbl(odv(i)(14).ToString.Trim() <> 0)) Then

                            If iCount > 1 Then
                                oDoc.Lines.Add()
                            End If

                            dtItemCode.DefaultView.RowFilter = "Name = 'TPA - FFS'"
                            If dtItemCode.DefaultView.Count = 0 Then
                                sErrDesc = "ItemCode ::'TPA - FFS' provided does not exist in SAP(Mapping Table)."
                                Call WriteToLogFile(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            Else
                                sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                            End If

                            dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                            If dtVatGroup.DefaultView.Count = 0 Then
                                sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                                Call WriteToLogFile(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            Else
                                sVatGroup = dtVatGroup.DefaultView.Item(0)(1).ToString().Trim()
                            End If

                            oDoc.Lines.ItemCode = sItemCode
                            oDoc.Lines.UnitPrice = Math.Abs(CDbl(odv(i)(14).ToString.Trim()))
                            oDoc.Lines.VatGroup = sVatGroup
                            If Not (sCostCenter = String.Empty) Then
                                oDoc.Lines.CostingCode = sCostCenter
                                oDoc.Lines.COGSCostingCode = sCostCenter
                            End If
                            bIsPanel = True

                            iCount = iCount + 1
                        End If
                    End If

                    If bIsPanel = False Then

                        If Not (odv(i)(7).ToString = String.Empty) Then

                            If (CDbl(odv(i)(7).ToString.Trim() <> 0)) Then


                                If iCount > 1 Then
                                    oDoc.Lines.Add()
                                End If

                                dtItemCode.DefaultView.RowFilter = "Name = 'GIRO Amount'"
                                If dtItemCode.DefaultView.Count = 0 Then
                                    sErrDesc = "ItemCode ::'GIRO Amount' provided does not exist in SAP(Mapping Table)."
                                    Call WriteToLogFile(sErrDesc, sFuncName)
                                    Throw New ArgumentException(sErrDesc)
                                Else
                                    sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                                End If

                                dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                                If dtVatGroup.DefaultView.Count = 0 Then
                                    sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                                    Call WriteToLogFile(sErrDesc, sFuncName)
                                    Throw New ArgumentException(sErrDesc)
                                Else
                                    sVatGroup = dtVatGroup.DefaultView.Item(0)(1).ToString().Trim()
                                End If

                                oDoc.Lines.ItemCode = sItemCode
                                oDoc.Lines.UnitPrice = Math.Abs(CDbl(odv(i)(7).ToString.Trim()))
                                oDoc.Lines.VatGroup = sVatGroup
                                If Not (sCostCenter = String.Empty) Then
                                    oDoc.Lines.CostingCode = sCostCenter
                                    oDoc.Lines.COGSCostingCode = sCostCenter
                                End If
                                If Not (sAcctCode = String.Empty) Then
                                    oDoc.Lines.AccountCode = sAcctCode
                                End If

                                iCount = iCount + 1
                            End If
                        End If

                        If Not (odv(i)(8).ToString = String.Empty) Then

                            If (CDbl(odv(i)(8).ToString.Trim() <> 0)) Then
                                If iCount > 1 Then
                                    oDoc.Lines.Add()
                                End If

                                dtItemCode.DefaultView.RowFilter = "Name = 'GIRO Fee'"
                                If dtItemCode.DefaultView.Count = 0 Then
                                    sErrDesc = "ItemCode ::'GIRO Fee' provided does not exist in SAP(Mapping Table)."
                                    Call WriteToLogFile(sErrDesc, sFuncName)
                                    Throw New ArgumentException(sErrDesc)
                                Else
                                    sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                                End If
                                dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                                If dtVatGroup.DefaultView.Count = 0 Then
                                    sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                                    Call WriteToLogFile(sErrDesc, sFuncName)
                                    Throw New ArgumentException(sErrDesc)
                                Else
                                    sVatGroup = dtVatGroup.DefaultView.Item(0)(1).ToString().Trim()
                                End If
                                oDoc.Lines.ItemCode = sItemCode
                                oDoc.Lines.UnitPrice = Math.Abs(CDbl(odv(i)(8).ToString.Trim()))
                                oDoc.Lines.VatGroup = sVatGroup
                                If Not (sCostCenter = String.Empty) Then
                                    oDoc.Lines.CostingCode = sCostCenter
                                    oDoc.Lines.COGSCostingCode = sCostCenter
                                End If

                                iCount = iCount + 1
                            End If
                        End If


                        If Not (odv(i)(9).ToString = String.Empty) Then
                            If (CDbl(odv(i)(9).ToString.Trim() <> 0)) Then
                                If iCount > 1 Then
                                    oDoc.Lines.Add()
                                End If

                                dtItemCode.DefaultView.RowFilter = "Name = 'GIRO Reprocessing Fee'"
                                If dtItemCode.DefaultView.Count = 0 Then
                                    sErrDesc = "ItemCode ::'GIRO Reprocessing Fee' provided does not exist in SAP(Mapping Table)."
                                    Call WriteToLogFile(sErrDesc, sFuncName)
                                    Throw New ArgumentException(sErrDesc)
                                Else
                                    sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                                End If
                                dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                                If dtVatGroup.DefaultView.Count = 0 Then
                                    sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                                    Call WriteToLogFile(sErrDesc, sFuncName)
                                    Throw New ArgumentException(sErrDesc)
                                Else
                                    sVatGroup = dtVatGroup.DefaultView.Item(0)(1).ToString().Trim()
                                End If

                                oDoc.Lines.ItemCode = sItemCode
                                oDoc.Lines.UnitPrice = Math.Abs(CDbl(odv(i)(9).ToString.Trim()))
                                oDoc.Lines.VatGroup = sVatGroup
                                If Not (sCostCenter = String.Empty) Then
                                    oDoc.Lines.CostingCode = sCostCenter
                                    oDoc.Lines.COGSCostingCode = sCostCenter
                                End If

                                iCount = iCount + 1
                            End If
                        End If


                        If Not (odv(i)(10).ToString = String.Empty) Then
                            If (CDbl(odv(i)(10).ToString.Trim() <> 0)) Then
                                If iCount > 1 Then
                                    oDoc.Lines.Add()
                                End If

                                dtItemCode.DefaultView.RowFilter = "Name = 'Cheque amount'"
                                If dtItemCode.DefaultView.Count = 0 Then
                                    sErrDesc = "ItemCode ::'Cheque amount' provided does not exist in SAP(Mapping Table)."
                                    Call WriteToLogFile(sErrDesc, sFuncName)
                                    Throw New ArgumentException(sErrDesc)
                                Else
                                    sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                                End If

                                dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                                If dtVatGroup.DefaultView.Count = 0 Then
                                    sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                                    Call WriteToLogFile(sErrDesc, sFuncName)
                                    Throw New ArgumentException(sErrDesc)
                                Else
                                    sVatGroup = dtVatGroup.DefaultView.Item(0)(1).ToString().Trim()
                                End If
                                oDoc.Lines.ItemCode = sItemCode
                                oDoc.Lines.UnitPrice = Math.Abs(CDbl(odv(i)(10).ToString.Trim()))
                                oDoc.Lines.VatGroup = sVatGroup
                                If Not (sCostCenter = String.Empty) Then
                                    oDoc.Lines.CostingCode = sCostCenter
                                    oDoc.Lines.COGSCostingCode = sCostCenter
                                End If
                                If Not (sAcctCode = String.Empty) Then
                                    oDoc.Lines.AccountCode = sAcctCode
                                End If

                                iCount = iCount + 1
                            End If
                        End If


                        If Not (odv(i)(11).ToString = String.Empty) Then
                            If (CDbl(odv(i)(11).ToString.Trim() <> 0)) Then
                                If iCount > 1 Then
                                    oDoc.Lines.Add()
                                End If

                                dtItemCode.DefaultView.RowFilter = "Name = 'Cheque Fee'"
                                If dtItemCode.DefaultView.Count = 0 Then
                                    sErrDesc = "ItemCode ::'Cheque Fee' provided does not exist in SAP(Mapping Table)."
                                    Call WriteToLogFile(sErrDesc, sFuncName)
                                    Throw New ArgumentException(sErrDesc)
                                Else
                                    sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                                End If
                                dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                                If dtVatGroup.DefaultView.Count = 0 Then
                                    sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                                    Call WriteToLogFile(sErrDesc, sFuncName)
                                    Throw New ArgumentException(sErrDesc)
                                Else
                                    sVatGroup = dtVatGroup.DefaultView.Item(0)(1).ToString().Trim()
                                End If
                                oDoc.Lines.ItemCode = sItemCode
                                oDoc.Lines.UnitPrice = Math.Abs(CDbl(odv(i)(11).ToString.Trim()))
                                oDoc.Lines.VatGroup = sVatGroup
                                If Not (sCostCenter = String.Empty) Then
                                    oDoc.Lines.CostingCode = sCostCenter
                                    oDoc.Lines.COGSCostingCode = sCostCenter
                                End If

                                iCount = iCount + 1
                            End If
                        End If


                        If Not (odv(i)(12).ToString = String.Empty) Then
                            If (CDbl(odv(i)(12).ToString.Trim() <> 0)) Then
                                If iCount > 1 Then
                                    oDoc.Lines.Add()
                                End If

                                dtItemCode.DefaultView.RowFilter = "Name = 'Cheque Reprocessing Fee'"
                                If dtItemCode.DefaultView.Count = 0 Then
                                    sErrDesc = "ItemCode ::'Cheque Reprocessing Fee' provided does not exist in SAP(Mapping Table)."
                                    Call WriteToLogFile(sErrDesc, sFuncName)
                                    Throw New ArgumentException(sErrDesc)
                                Else
                                    sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                                End If
                                dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                                If dtVatGroup.DefaultView.Count = 0 Then
                                    sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                                    Call WriteToLogFile(sErrDesc, sFuncName)
                                    Throw New ArgumentException(sErrDesc)
                                Else
                                    sVatGroup = dtVatGroup.DefaultView.Item(0)(1).ToString().Trim()
                                End If
                                oDoc.Lines.ItemCode = sItemCode
                                oDoc.Lines.UnitPrice = Math.Abs(CDbl(odv(i)(12).ToString.Trim()))
                                oDoc.Lines.VatGroup = sVatGroup
                                If Not (sCostCenter = String.Empty) Then
                                    oDoc.Lines.CostingCode = sCostCenter
                                    oDoc.Lines.COGSCostingCode = sCostCenter
                                End If

                                iCount = iCount + 1
                            End If
                        End If

                        If Not (odv(i)(15).ToString = String.Empty) Then
                            If (CDbl(odv(i)(15).ToString.Trim() <> 0)) Then
                                If iCount > 1 Then
                                    oDoc.Lines.Add()
                                End If

                                dtItemCode.DefaultView.RowFilter = "Name = 'Card Replacement'"
                                If dtItemCode.DefaultView.Count = 0 Then
                                    sErrDesc = "ItemCode ::'Card Replacement' provided does not exist in SAP(Mapping Table)."
                                    Call WriteToLogFile(sErrDesc, sFuncName)
                                    Throw New ArgumentException(sErrDesc)
                                Else
                                    sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                                End If
                                dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                                If dtVatGroup.DefaultView.Count = 0 Then
                                    sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                                    Call WriteToLogFile(sErrDesc, sFuncName)
                                    Throw New ArgumentException(sErrDesc)
                                Else
                                    sVatGroup = dtVatGroup.DefaultView.Item(0)(1).ToString().Trim()
                                End If
                                oDoc.Lines.ItemCode = sItemCode
                                oDoc.Lines.UnitPrice = Math.Abs(CDbl(odv(i)(15).ToString.Trim()))
                                oDoc.Lines.VatGroup = sVatGroup
                                If Not (sCostCenter = String.Empty) Then
                                    oDoc.Lines.CostingCode = sCostCenter
                                    oDoc.Lines.COGSCostingCode = sCostCenter
                                End If

                                iCount = iCount + 1
                            End If
                        End If
                    End If


                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Document", sFuncName)

                    iRetCode = oDoc.Add()

                    If iRetCode <> 0 Then
                        p_oCompany.GetLastError(iErrCode, sErrDesc)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        Dim iDocNo As Integer
                        p_oCompany.GetNewObjectCode(iDocNo)
                        Console.WriteLine("Document Created successfully :: " & iDocNo)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oDoc)
                    End If

                End If
            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            AddInvoice_CreditNote = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message.ToString
            Call WriteToLogFile_Debug(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            AddInvoice_CreditNote = RTN_ERROR
        End Try

    End Function

End Module
