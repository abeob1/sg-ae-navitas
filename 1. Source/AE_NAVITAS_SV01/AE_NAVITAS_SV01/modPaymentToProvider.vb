Module modPaymentToProvider

    Private dtBP As DataTable
    Private dtProject As DataTable
    Private dtAcctCode As DataTable
    Private dtFileName As DataTable

    Public Function ProcessPayToProvider(ByVal file As System.IO.FileInfo, ByVal odv As DataView, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = "ProcessPayToProvider"
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

                sSQL = "SELECT ""OcrCode"" FROM ""OOCR"" "
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING  SQL :" & sSQL, sFuncName)
                dtProject = ExecuteQueryReturnDataTable(sSQL, p_oCompany.CompanyDB)

                sSQL = "SELECT ""AcctCode"" FROM ""OACT"" "
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING  SQL :" & sSQL, sFuncName)
                dtAcctCode = ExecuteQueryReturnDataTable(sSQL, p_oCompany.CompanyDB)

                sSQL = "SELECT ""CardCode"" FROM ""OCRD"""
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
                dtBP = ExecuteQueryReturnDataTable(sSQL, p_oCompany.CompanyDB)

                sSQL = "SELECT T0.""GroupCode"" ,T0.""GroupName"" FROM ""OCRG"" T0 WHERE T0.""GroupType""='S' AND T0.""Locked""='N'"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING  SQL :" & sSQL, sFuncName)
                p_oDtBPGroup = ExecuteQueryReturnDataTable(sSQL, p_oCompany.CompanyDB)

                sSQL = "SELECT T0.""GroupNum"", T0.""PymntGroup"" FROM ""OCTG"" T0"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING  SQL :" & sSQL, sFuncName)
                p_oDtPayTerms = ExecuteQueryReturnDataTable(sSQL, p_oCompany.CompanyDB)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ProcessExcelDatas()", sFuncName)

                If ProcessExcelDatas(file, file.Name, odv, sDBCode, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting from SAP", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

            ProcessPayToProvider = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message.ToString()
            Call WriteToLogFile(sErrDesc, sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ProcessPayToProvider = RTN_ERROR
        End Try

    End Function

    Private Function ProcessExcelDatas(ByVal file As System.IO.FileInfo, ByVal sFileName As String, ByVal oDv As DataView, ByVal sDBCode As String, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "ProcessExcelDatas"
        Dim sBatchNo As String = String.Empty
        Dim sBatchPeriod As String = String.Empty
        Dim oDTGiro As DataTable
        Dim oDTCheque As DataTable
        Dim bTransStarted As Boolean = False
        Dim sFullBatchPeriod As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            oDTGiro = oDv.Table.Clone
            oDTCheque = oDv.Table.Clone
            oDTGiro.Clear()
            oDTCheque.Clear()

            sBatchNo = oDv(2)(0).ToString()
            sFullBatchPeriod = oDv(3)(0).ToString()

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Reading Batch no", sFuncName)
            Dim m As Integer = InStrRev(sBatchNo, ":")
            sBatchNo = Microsoft.VisualBasic.Right(sBatchNo, Len(sBatchNo) - m).Trim

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Reading Batch Period", sFuncName)
            Dim n As Integer = InStrRev(sFullBatchPeriod, "to")
            sBatchPeriod = Microsoft.VisualBasic.Right(sFullBatchPeriod, Len(sFullBatchPeriod) - n - 1).Trim

            For iRow As Integer = 6 To oDv.Count - 1
                If (oDv(iRow)(7).ToString.Trim = String.Empty) Then
                    sErrDesc = "Payment Mode not found in some rows"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                    Throw New ArgumentException(sErrDesc)
                End If
            Next

            '**********************CREATE OUTGOING PAYMENT FOR GIRO ITEMS**********************************

            Dim oDTGRIOGrouped As DataTable = Nothing
            Dim oDVGIRODetl As DataView = New DataView

            oDTGRIOGrouped = oDv.Table.DefaultView.ToTable(True, "F8")

            For intRow As Integer = 0 To oDTGRIOGrouped.Rows.Count - 1
                If Not (oDTGRIOGrouped.Rows(intRow).Item(0).ToString.Trim() = String.Empty Or oDTGRIOGrouped.Rows(intRow).Item(0).ToString.ToUpper.Trim = "PAYMENT MODE") Then
                    oDv.RowFilter = "F8 = '" & oDTGRIOGrouped.Rows(intRow).Item(0).ToString.Trim() & "'"
                End If
            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Splitting Dataview based on Payment Mode - GIRO", sFuncName)

            For Each row As DataRow In oDv.Table.Rows
                If row.Item(7).ToString.Trim.ToUpper = "GIRO" Then
                    oDTGiro.ImportRow(row)
                End If
            Next

            oDVGIRODetl = New DataView(oDTGiro)

            oDTGRIOGrouped = oDVGIRODetl.Table.DefaultView.ToTable(True, "F2")

            For intRow As Integer = 0 To oDTGRIOGrouped.Rows.Count - 1
                If Not (oDTGRIOGrouped.Rows(intRow).Item(0).ToString.Trim() = String.Empty Or oDTGRIOGrouped.Rows(intRow).Item(0).ToString.ToUpper().Trim() = "PAYMENT MODE") Then
                    oDVGIRODetl.RowFilter = "F2 = '" & oDTGRIOGrouped.Rows(intRow).Item(0).ToString.Trim() & "'"

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateOutGoingPayment_GIRO()", sFuncName)
                    Console.WriteLine("Creating outgoing payment document")

                    If bTransStarted = False Then
                        If StartTransaction(sErrDesc) = RTN_SUCCESS Then
                            bTransStarted = True
                        Else
                            Throw New ArgumentException(sErrDesc)
                        End If
                    End If

                    If bTransStarted = True Then
                        If CreateOutGoingPayment_GIRO(oDVGIRODetl, file.Name, sDBCode, sBatchNo, sBatchPeriod, sFullBatchPeriod, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If

                End If
            Next

            '**********************GROUPING CHEQUE ITEMS AND CREATING OUTGOING PAYMENT*******************
            Dim oDTChkGrouped As DataTable = Nothing
            Dim oDVChkDetl As DataView = New DataView

            oDTChkGrouped = oDv.Table.DefaultView.ToTable(True, "F8")

            For intRow As Integer = 0 To oDTChkGrouped.Rows.Count - 1
                If Not (oDTChkGrouped.Rows(intRow).Item(0).ToString.Trim() = String.Empty Or oDTChkGrouped.Rows(intRow).Item(0).ToString.ToUpper.Trim = "PAYMENT MODE") Then
                    oDv.RowFilter = "F8 = '" & oDTChkGrouped.Rows(intRow).Item(0).ToString & "'"
                End If
            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Splitting Dataview based on Payment Mode - Cheque", sFuncName)

            For Each row As DataRow In oDv.Table.Rows
                If row.Item(7).ToString.Trim.ToUpper = "CHEQUE" Then
                    oDTCheque.ImportRow(row)
                End If
            Next

            oDVChkDetl = New DataView(oDTCheque)

            For intRow As Integer = 0 To oDTChkGrouped.Rows.Count - 1
                If Not (oDTChkGrouped.Rows(intRow).Item(0).ToString.Trim() = String.Empty Or oDTChkGrouped.Rows(intRow).Item(0).ToString.ToUpper.Trim = "PAYMENT MODE") Then
                    oDVChkDetl.RowFilter = "F8 = '" & oDTChkGrouped.Rows(intRow).Item(0).ToString & "'"

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateOutGoingPayment_Cheque()", sFuncName)

                    If bTransStarted = False Then
                        If StartTransaction(sErrDesc) = RTN_SUCCESS Then
                            bTransStarted = True
                        Else
                            Throw New ArgumentException(sErrDesc)
                        End If
                    End If

                    If bTransStarted = True Then
                        If CreateOutGoingPayment_Cheque(oDVChkDetl, file.Name, sDBCode, sBatchNo, sBatchPeriod, sFullBatchPeriod, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If
                End If
            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CommitTransaction", sFuncName)
            If CommitTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling FileMoveToArchive()", sFuncName)
            'FileMoveToArchive(file, file.FullName, RTN_SUCCESS)
            FileMoveToArchive_Success(file, file.FullName, file.Name, RTN_SUCCESS)

            'Insert Error Description into Table
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
            AddDataToTable(p_oDtSuccess, file.Name, "Success")
            'error condition

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            ProcessExcelDatas = RTN_SUCCESS
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
            ProcessExcelDatas = RTN_ERROR
        End Try

    End Function

    Private Function CreateOutGoingPayment_GIRO(ByVal odv As DataView, ByVal sFileName As String, _
                                                ByVal sDBCode As String, ByVal sBatchNo As String, _
                                                ByVal sBatchPeriod As String, ByVal sFullBatchPeriod As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = "CreateOutGoingPayment_GIRO"
        Dim iRetcode, iErrCode As Long
        Dim iCount As Integer = 1
        Dim sTrnsAcct As String = String.Empty
        Dim dTotPaymentAmt As Double = 0.0
        Dim dTotalAmt As Double = 0.0
        Dim dTPACol As Double = 0.0
        Dim dGSTCol As Double = 0.0
        Dim dReimbCol As Double = 0.0
        Dim bIsLineAdded As Boolean = False
        Dim sCardCode As String = String.Empty
        Dim sFullCardCode As String = String.Empty
        Dim sCardName As String = String.Empty
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim oRecordSet1 As SAPbobsCOM.Recordset
        Dim sSql As String = String.Empty

        Try
            sSql = "SELECT DISTINCT ""U_AI_APARUploadName"" FROM ""OVPM"" WHERE IFNULL(""U_AI_APARUploadName"",'') <> ''"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSql, sFuncName)
            dtFileName = ExecuteQueryReturnDataTable(sSql, p_oCompany.CompanyDB)

            dtFileName.DefaultView.RowFilter = "U_AI_APARUploadName = '" & sFileName & "'"
            If dtFileName.DefaultView.Count > 0 Then
                sErrDesc = "Interface file ::" & sFileName & " has already been uploaded"
                Call WriteToLogFile(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            Dim oPayments As SAPbobsCOM.Payments
            oPayments = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments)
            oPayments.DocType = SAPbobsCOM.BoRcptTypes.rSupplier

            sFullCardCode = odv(0)(1).ToString.Trim()

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("CardCode is " & sFullCardCode, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("CardCode Length is " & sFullCardCode.Length, sFuncName)

            If sFullCardCode.Length > 15 Then
                sCardCode = sFullCardCode.Substring(0, 15)
            Else
                sCardCode = sFullCardCode
            End If

            sCardName = odv(0)(0).ToString.Trim()

            dtBP.DefaultView.RowFilter = "CardCode = '" & sCardCode & "'"
            If dtBP.DefaultView.Count = 0 Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CheckBP", sFuncName)
                If CheckBP(sFullCardCode, sCardCode, sCardName, "V", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Create_OutgoingPayment", sFuncName)
                If Create_OutgoingPayment(odv, sFileName, sDBCode, sBatchNo, sBatchPeriod, "GIRO", sFullBatchPeriod, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                Return RTN_SUCCESS
            End If

            oPayments.CardCode = sCardCode
            '' oPayments.CardName = sCardName

            oPayments.DocDate = CDate(sBatchPeriod)
            oPayments.CounterReference = odv(0)(9).ToString.Trim
            oPayments.Remarks = sBatchNo
            oPayments.JournalRemarks = sFullBatchPeriod
            oPayments.UserFields.Fields.Item("U_AI_APARUploadName").Value = sFileName

            oPayments.BPLID = "1"

            For i As Integer = 0 To odv.Count - 1
                If Not (odv(i)(0).ToString = String.Empty) Then
                    If Not (odv(i)(4).ToString.Trim = String.Empty) Then
                        dReimbCol = odv(i)(4).ToString.Trim
                    Else
                        dReimbCol = 0.0
                    End If
                    'If Not (odv(i)(5).ToString.Trim = String.Empty) Then
                    '    dTPACol = odv(i)(5).ToString.Trim
                    'Else
                    '    dTPACol = 0.0
                    'End If
                    'If Not (odv(i)(6).ToString.Trim = String.Empty) Then
                    '    dGSTCol = odv(i)(6).ToString.Trim
                    'Else
                    '    dGSTCol = 0.0
                    'End If

                    dTotalAmt = dTotalAmt + dReimbCol
                End If
            Next

            dTotPaymentAmt = dTotalAmt

            If dTotPaymentAmt > 0.0 Then
                Dim sClincCode As String
                sClincCode = odv(0)(1).ToString

                ''Dim sSql As String
                Dim sBaseRef As String = String.Empty
                Dim dTransAmt As Double = 0.0
                Dim sTransType As String = String.Empty

                '**************GETTING ONLY DEBIT VALUES**********************

                sSql = "SELECT CASE WHEN ""TransType""=46 THEN ""TransId"" ELSE ""CreatedBy"" END ""CreatedBy"", ""BalDueDeb"" *-1 ""Total"",""TransType"" "
                sSql = sSql & " FROM ""JDT1"" WHERE ""ShortName"" = '" & sClincCode & "' and ""BalDueDeb"" - ""BalDueCred"" <> 0 "
                sSql = sSql & " AND ""BalDueDeb"" > 0"
                sSql = sSql & " ORDER BY ""DueDate"",""BaseRef"" "
                oRecordSet = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Query : " & sSql, sFuncName)

                oRecordSet.DoQuery(sSql)
                If Not (oRecordSet.BoF And oRecordSet.EoF) Then
                    oRecordSet.MoveFirst()
                    Do Until oRecordSet.EoF
                        sBaseRef = oRecordSet.Fields.Item("CreatedBy").Value
                        dTransAmt = oRecordSet.Fields.Item("Total").Value
                        sTransType = oRecordSet.Fields.Item("TransType").Value
                        If dTotalAmt > 0.0 Then
                            If dTransAmt < 0.0 Then
                                'If iCount > 1 Then
                                '    oPayments.Invoices.Add()
                                'End If
                                oPayments.Invoices.DocEntry = sBaseRef
                                If sTransType = "18" Then
                                    oPayments.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_PurchaseInvoice
                                    oPayments.Invoices.DocLine = 0
                                ElseIf sTransType = "19" Then
                                    oPayments.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_PurchaseCreditNote
                                    oPayments.Invoices.DocLine = 0
                                ElseIf sTransType = "46" Then
                                    oPayments.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_JournalEntry
                                    oPayments.Invoices.DocLine = 1
                                End If

                                oPayments.Invoices.SumApplied = dTransAmt
                                dTotalAmt = dTotalAmt - dTransAmt

                                If Not (odv(0)(2).ToString.Trim = String.Empty) Then
                                    oPayments.Invoices.DistributionRule = odv(0)(2).ToString.Trim
                                End If

                                oPayments.Invoices.UserFields.Fields.Item("U_AI_PayMode").Value = odv(0)(7).ToString.Trim
                                oPayments.Invoices.UserFields.Fields.Item("U_AI_PayeeAcNo").Value = odv(0)(8).ToString.Trim

                                bIsLineAdded = True

                                oPayments.Invoices.Add()
                                ''iCount = iCount + 1
                            End If
                        End If
                        oRecordSet.MoveNext()
                    Loop

                End If



                ''  System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)


                '*****************GETTING CREDIT VALUES*********************
                sSql = "SELECT CASE WHEN ""TransType""=46 THEN ""TransId"" ELSE ""CreatedBy"" END ""CreatedBy"", ""BalDueCred"" ""Total"",""TransType"" "
                sSql = sSql & " FROM ""JDT1"" WHERE ""ShortName"" = '" & sClincCode & "' and ""BalDueDeb"" - ""BalDueCred"" <> 0 "
                sSql = sSql & " AND ""BalDueCred"" > 0"
                sSql = sSql & " ORDER BY ""DueDate"",""BaseRef"" "
                oRecordSet1 = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Query : " & sSql, sFuncName)

                oRecordSet1.DoQuery(sSql)
                If Not (oRecordSet1.BoF And oRecordSet1.EoF) Then
                    oRecordSet1.MoveFirst()
                    Do Until oRecordSet1.EoF
                        sBaseRef = oRecordSet1.Fields.Item("CreatedBy").Value
                        dTransAmt = oRecordSet1.Fields.Item("Total").Value
                        sTransType = oRecordSet1.Fields.Item("TransType").Value
                        If dTotalAmt > 0.0 Then
                            'If iCount > 1 Then
                            '    oPayments.Invoices.Add()
                            'End If
                            oPayments.Invoices.DocEntry = sBaseRef
                            If sTransType = "18" Then
                                oPayments.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_PurchaseInvoice
                                oPayments.Invoices.DocLine = 0
                            ElseIf sTransType = "19" Then
                                oPayments.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_PurchaseCreditNote
                                oPayments.Invoices.DocLine = 0
                            ElseIf sTransType = "46" Then
                                oPayments.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_JournalEntry
                                oPayments.Invoices.DocLine = 1
                            End If
                            If dTotalAmt > dTransAmt Then
                                oPayments.Invoices.SumApplied = dTransAmt
                                dTotalAmt = dTotalAmt - dTransAmt
                            Else
                                oPayments.Invoices.SumApplied = dTotalAmt
                                dTotalAmt = dTotalAmt - dTotalAmt
                            End If
                            If Not (odv(0)(2).ToString.Trim = String.Empty) Then
                                oPayments.Invoices.DistributionRule = odv(0)(2).ToString.Trim
                            End If

                            oPayments.Invoices.UserFields.Fields.Item("U_AI_PayMode").Value = odv(0)(7).ToString.Trim
                            oPayments.Invoices.UserFields.Fields.Item("U_AI_PayeeAcNo").Value = odv(0)(8).ToString.Trim

                            ''iCount = iCount + 1
                            oPayments.Invoices.Add()
                            bIsLineAdded = True
                        Else
                            Exit Do
                        End If
                        oRecordSet1.MoveNext()
                    Loop
                End If
                ''System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)

                If oRecordSet.RecordCount = 0 AndAlso oRecordSet1.RecordCount = 0 Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Create_OutgoingPayment", sFuncName)
                    If Create_OutgoingPayment(odv, sFileName, sDBCode, sBatchNo, sBatchPeriod, "GIRO", sFullBatchPeriod, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    Return RTN_SUCCESS
                End If

                If dTotPaymentAmt > 0.0 And bIsLineAdded = True Then
                    sTrnsAcct = GetBankTrnsAcct(sDBCode)

                    oPayments.TransferAccount = sTrnsAcct
                    oPayments.TransferDate = CDate(sBatchPeriod)
                    oPayments.TransferSum = dTotPaymentAmt
                End If

                If bIsLineAdded = True Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Document", sFuncName)

                    iRetcode = oPayments.Add()

                    If iRetcode <> 0 Then
                        p_oCompany.GetLastError(iErrCode, sErrDesc)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        iCount = 0
                        Dim iDocNo As Integer
                        p_oCompany.GetNewObjectCode(iDocNo)
                        Console.WriteLine("Document Created Successfully :: " & iDocNo)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oPayments)
                    End If
                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateOutGoingPayment_GIRO = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message.ToString()
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Completed with ERROR", sFuncName)
            CreateOutGoingPayment_GIRO = RTN_ERROR
        End Try

    End Function

    Private Function CreateOutGoingPayment_Cheque(ByVal odv As DataView, ByVal sFileName As String, _
                                                  ByVal sDBCode As String, ByVal sBatchNo As String, _
                                                  ByVal sBatchPeriod As String, ByVal sFullBatchPeriod As String, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateOutGoingPayment_Cheque"
        Dim iCount As Integer = 0
        Dim iRetCode, iErrCode As Long
        Dim dTotPaymentAmt As Double = 0.0
        Dim sBankCntryCode As String = String.Empty
        Dim sBankCode As String = String.Empty
        Dim sChkGLAcct As String = String.Empty
        Dim sChkAcct As String = String.Empty
        Dim dTPACol As Double = 0.0
        Dim dGSTCol As Double = 0.0
        Dim dReimbCol As Double = 0.0
        Dim dTotalAmt As Double = 0.0
        Dim sSql As String
        Dim sCardCode As String = String.Empty
        Dim sFullCardCode As String = String.Empty
        Dim sCardName As String = String.Empty
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim oRecordSet1 As SAPbobsCOM.Recordset
        Try

            sSql = "SELECT DISTINCT ""U_AI_APARUploadName"" FROM ""OVPM"" WHERE IFNULL(""U_AI_APARUploadName"",'') <> ''"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSql, sFuncName)
            dtFileName = ExecuteQueryReturnDataTable(sSql, p_oCompany.CompanyDB)

            dtFileName.DefaultView.RowFilter = "U_AI_APARUploadName = '" & sFileName & "'"
            If dtFileName.DefaultView.Count > 0 Then
                sErrDesc = "Interface file ::" & sFileName & " has already been uploaded"
                Call WriteToLogFile(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            Dim oPayments As SAPbobsCOM.IPayments

            For i As Integer = 0 To odv.Count - 1
                If Not (odv(i)(1).ToString = String.Empty) Then

                    Console.WriteLine("Creating outgoing payment document - Cheque")

                    If Not (odv(i)(4).ToString.Trim = String.Empty) Then
                        dReimbCol = odv(i)(4).ToString.Trim
                    Else
                        dReimbCol = 0.0
                    End If

                    dTotalAmt = dReimbCol
                    dTotPaymentAmt = dReimbCol

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                    oPayments = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments)
                    oPayments.DocType = SAPbobsCOM.BoRcptTypes.rSupplier

                    sFullCardCode = odv(0)(1).ToString.Trim()

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("CardCode is " & sFullCardCode, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("CardCode Length is " & sFullCardCode.Length, sFuncName)

                    If sFullCardCode.Length > 15 Then
                        sCardCode = sFullCardCode.Substring(0, 15)
                    Else
                        sCardCode = sFullCardCode
                    End If

                    sCardName = odv(0)(0).ToString.Trim()

                    dtBP.DefaultView.RowFilter = "CardCode = '" & sCardCode & "'"
                    If dtBP.DefaultView.Count = 0 Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CheckBP", sFuncName)
                        If CheckBP(sFullCardCode, sCardCode, sCardName, "V", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Create_OutgoingPayment", sFuncName)
                        If Create_OutgoingPayment(odv, sFileName, sDBCode, sBatchNo, sBatchPeriod, "CHEQUE", sFullBatchPeriod, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        Return RTN_SUCCESS
                    End If

                    oPayments.CardCode = sCardCode
                    oPayments.DocDate = CDate(sBatchPeriod)
                    oPayments.CounterReference = odv(i)(9).ToString.Trim
                    oPayments.Remarks = sBatchNo
                    oPayments.UserFields.Fields.Item("U_AI_APARUploadName").Value = sFileName

                    oPayments.BPLID = "1"

                    iCount = 0
                    iCount = iCount + 1

                    Dim bIsLineAdded As Boolean = False
                    Dim sBaseRef As String = String.Empty
                    Dim dTransAmt As Double = 0.0
                    Dim sTransType As String = String.Empty
                    Dim sClincCode As String
                    sClincCode = odv(i)(1).ToString

                    '**************GETTING ONLY DEBIT VALUES**********************

                    sSql = "SELECT CASE WHEN ""TransType""=46 THEN ""TransId"" ELSE ""CreatedBy"" END ""CreatedBy"", ""BalDueDeb"" *-1 ""Total"",""TransType"" "
                    sSql = sSql & " FROM ""JDT1"" WHERE ""ShortName"" = '" & sClincCode & "' and ""BalDueDeb"" - ""BalDueCred"" <> 0 "
                    sSql = sSql & " AND ""BalDueDeb"" > 0"
                    sSql = sSql & " ORDER BY ""DueDate"",""BaseRef"" "
                    oRecordSet = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecordSet.DoQuery(sSql)
                    If Not (oRecordSet.BoF And oRecordSet.EoF) Then
                        oRecordSet.MoveFirst()
                        Do Until oRecordSet.EoF
                            sBaseRef = oRecordSet.Fields.Item("CreatedBy").Value
                            dTransAmt = oRecordSet.Fields.Item("Total").Value
                            sTransType = oRecordSet.Fields.Item("TransType").Value
                            If dTotalAmt > 0.0 Then
                                If dTransAmt < 0.0 Then
                                    If iCount > 1 Then
                                        oPayments.Invoices.Add()
                                    End If
                                    oPayments.Invoices.DocEntry = sBaseRef
                                    If sTransType = "18" Then
                                        oPayments.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_PurchaseInvoice
                                        oPayments.Invoices.DocLine = 0
                                    ElseIf sTransType = "19" Then
                                        oPayments.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_PurchaseCreditNote
                                        oPayments.Invoices.DocLine = 0
                                    ElseIf sTransType = "46" Then
                                        oPayments.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_JournalEntry
                                        oPayments.Invoices.DocLine = 1
                                    End If
                                    oPayments.Invoices.SumApplied = dTransAmt
                                    dTotalAmt = dTotalAmt - dTransAmt

                                    If Not (odv(i)(2).ToString.Trim = String.Empty) Then
                                        oPayments.Invoices.DistributionRule = odv(i)(2).ToString.Trim
                                    End If

                                    oPayments.Invoices.UserFields.Fields.Item("U_AI_PayMode").Value = odv(i)(7).ToString.Trim
                                    oPayments.Invoices.UserFields.Fields.Item("U_AI_PayeeAcNo").Value = odv(i)(8).ToString.Trim

                                    iCount = iCount + 1
                                    bIsLineAdded = True

                                End If
                            End If
                            oRecordSet.MoveNext()
                        Loop
                    End If
                    ''System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)

                    '*****************GETTING CREDIT VALUES*********************
                    sSql = "SELECT CASE WHEN ""TransType""=46 THEN ""TransId"" ELSE ""CreatedBy"" END ""CreatedBy"", ""BalDueCred"" ""Total"",""TransType"" "
                    sSql = sSql & " FROM ""JDT1"" WHERE ""ShortName"" = '" & sClincCode & "' and ""BalDueDeb"" - ""BalDueCred"" <> 0 "
                    sSql = sSql & " AND ""BalDueCred"" > 0"
                    sSql = sSql & " ORDER BY ""DueDate"",""BaseRef"" "
                    oRecordSet1 = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecordSet1.DoQuery(sSql)
                    If Not (oRecordSet1.BoF And oRecordSet1.EoF) Then
                        oRecordSet1.MoveFirst()
                        Do Until oRecordSet1.EoF
                            sBaseRef = oRecordSet1.Fields.Item("CreatedBy").Value
                            dTransAmt = oRecordSet1.Fields.Item("Total").Value
                            sTransType = oRecordSet1.Fields.Item("TransType").Value
                            If dTotalAmt > 0.0 Then
                                If iCount > 1 Then
                                    oPayments.Invoices.Add()
                                End If
                                oPayments.Invoices.DocEntry = sBaseRef

                                If sTransType = "18" Then
                                    oPayments.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_PurchaseInvoice
                                    oPayments.Invoices.DocLine = 0
                                ElseIf sTransType = "19" Then
                                    oPayments.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_PurchaseCreditNote
                                    oPayments.Invoices.DocLine = 0
                                ElseIf sTransType = "46" Then
                                    oPayments.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_JournalEntry
                                    oPayments.Invoices.DocLine = 1
                                End If
                                If dTotalAmt > dTransAmt Then
                                    oPayments.Invoices.SumApplied = dTransAmt
                                    dTotalAmt = dTotalAmt - dTransAmt
                                Else
                                    oPayments.Invoices.SumApplied = dTotalAmt
                                    dTotalAmt = dTotalAmt - dTotalAmt
                                End If

                                If Not (odv(i)(2).ToString.Trim = String.Empty) Then
                                    oPayments.Invoices.DistributionRule = odv(i)(2).ToString.Trim
                                End If

                                oPayments.Invoices.UserFields.Fields.Item("U_AI_PayMode").Value = odv(i)(7).ToString.Trim
                                oPayments.Invoices.UserFields.Fields.Item("U_AI_PayeeAcNo").Value = odv(i)(8).ToString.Trim

                                iCount = iCount + 1
                                bIsLineAdded = True

                            Else
                                Exit Do
                            End If
                            oRecordSet1.MoveNext()
                        Loop
                    End If

                    If oRecordSet.RecordCount = 0 AndAlso oRecordSet1.RecordCount = 0 Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Create_OutgoingPayment", sFuncName)
                        If Create_OutgoingPayment(odv, sFileName, sDBCode, sBatchNo, sBatchPeriod, "CHEQUE", sFullBatchPeriod, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        Return RTN_SUCCESS
                    End If
                    ''  System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)

                    sBankCntryCode = GetCountryCode(sDBCode)
                    sBankCode = GetBankCode(sDBCode)
                    sChkGLAcct = GetCheckGLAcct(sDBCode)
                    sChkAcct = GetBankAcct(sDBCode)

                    oPayments.Checks.CountryCode = sBankCntryCode
                    oPayments.Checks.BankCode = sBankCode
                    oPayments.Checks.AccounttNum = sChkAcct
                    oPayments.Checks.CheckSum = dTotPaymentAmt
                    oPayments.Checks.CheckAccount = sChkGLAcct
                    oPayments.Checks.DueDate = CDate(sBatchPeriod)
                    oPayments.CashSum = 0
                    oPayments.TransferSum = 0

                    If bIsLineAdded = True Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Document", sFuncName)

                        iRetCode = oPayments.Add()

                        If iRetCode <> 0 Then
                            p_oCompany.GetLastError(iErrCode, sErrDesc)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            Dim iDocNo As Integer
                            p_oCompany.GetNewObjectCode(iDocNo)
                            Console.WriteLine("Document Created Successfully :: " & iDocNo)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oPayments)
                            iCount = 0
                        End If
                    End If

                End If
            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateOutGoingPayment_Cheque = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message.ToString()
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Completed with ERROR", sFuncName)
            CreateOutGoingPayment_Cheque = RTN_ERROR
        End Try

    End Function

    Private Function Create_OutgoingPayment(ByVal oDV As DataView, ByVal sFileName As String, _
                                            ByVal sDBCode As String, ByVal sBatchNo As String, _
                                            ByVal sBatchPeriod As String, ByVal sType As String, _
                                            ByVal sFullBatchPeriod As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim oPayment As SAPbobsCOM.Payments
        Dim lRetCode As Double
        Dim sCardCode As String = String.Empty
        Dim sFullCardcode As String = String.Empty
        Dim sCardName As String = String.Empty
        Dim sBankCntryCode As String = String.Empty
        Dim sBankCode As String = String.Empty
        Dim sChkGLAcct As String = String.Empty
        Dim sChkAcct As String = String.Empty
        Dim dReimbCol As Double = 0.0
        Dim dTotalAmt As Double = 0.0
        Dim sTrnsAcct As String = String.Empty
        Dim sSql As String = String.Empty

        Try

            sSql = "SELECT DISTINCT ""U_AI_APARUploadName"" FROM ""OVPM"" WHERE IFNULL(""U_AI_APARUploadName"",'') <> ''"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSql, sFuncName)
            dtFileName = ExecuteQueryReturnDataTable(sSql, p_oCompany.CompanyDB)

            dtFileName.DefaultView.RowFilter = "U_AI_APARUploadName = '" & sFileName & "'"
            If dtFileName.DefaultView.Count > 0 Then
                sErrDesc = "Interface file ::" & sFileName & " has already been uploaded"
                Call WriteToLogFile(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            sFuncName = "Create_OutgoingPayment"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oPayment = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments)

            sFullCardcode = oDV(0)(1).ToString.Trim()

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("CardCode is " & sFullCardcode, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("CardCode length is " & sFullCardcode.Length, sFuncName)

            If sFullCardcode.Length > 15 Then
                sCardCode = sFullCardcode.Substring(0, 15)
            Else
                sCardCode = sFullCardcode
            End If

            sCardName = oDV(0)(0).ToString.Trim()

            If sType.ToString().Trim().ToUpper() = "CHEQUE" Then

                Console.WriteLine("Creating Outgoing payment for CHEQUE")

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating Outgoing payment for CHEQUE", sFuncName)

                For iCount As Integer = 0 To oDV.Count - 1
                    If Not (oDV(iCount)(1).ToString = String.Empty) Then

                        If Not (oDV(iCount)(4).ToString.Trim = String.Empty) Then
                            dReimbCol = oDV(iCount)(4).ToString.Trim
                        Else
                            dReimbCol = 0.0
                        End If

                        oPayment.CardCode = sCardCode
                        oPayment.DocDate = CDate(sBatchPeriod)
                        oPayment.CounterReference = oDV(iCount)(9).ToString.Trim
                        oPayment.Remarks = sBatchNo
                        oPayment.JournalRemarks = sFullBatchPeriod
                        oPayment.UserFields.Fields.Item("U_AI_APARUploadName").Value = sFileName
                        oPayment.BPLID = "1"

                        sBankCntryCode = GetCountryCode(sDBCode)
                        sBankCode = GetBankCode(sDBCode)
                        sChkGLAcct = GetCheckGLAcct(sDBCode)
                        sChkAcct = GetBankAcct(sDBCode)

                        oPayment.Checks.CountryCode = sBankCntryCode
                        oPayment.Checks.BankCode = sBankCode
                        oPayment.Checks.AccounttNum = sChkAcct
                        oPayment.Checks.CheckSum = dReimbCol
                        oPayment.Checks.CheckAccount = sChkGLAcct
                        oPayment.Checks.DueDate = CDate(sBatchPeriod)
                        oPayment.CashSum = 0
                        oPayment.TransferSum = 0

                        lRetCode = oPayment.Add()

                        If lRetCode <> 0 Then
                            sErrDesc = p_oCompany.GetLastErrorDescription()
                            Throw New ArgumentException(sErrDesc)
                        End If
                    End If
                Next
            ElseIf sType.ToString().Trim().ToUpper() = "GIRO" Then
                Console.WriteLine("Creating Outgoing payment for GIRO")

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating Outgoing payment for GIRO", sFuncName)

                oPayment.CardCode = sCardCode
                oPayment.DocDate = CDate(sBatchPeriod)
                oPayment.CounterReference = oDV(0)(9).ToString.Trim
                oPayment.Remarks = sBatchNo
                oPayment.JournalRemarks = sFullBatchPeriod
                oPayment.UserFields.Fields.Item("U_AI_APARUploadName").Value = sFileName
                oPayment.BPLID = "1"

                For iCount As Integer = 0 To oDV.Count - 1
                    If Not (oDV(iCount)(1).ToString = String.Empty) Then
                        If Not (oDV(iCount)(4).ToString.Trim = String.Empty) Then
                            dReimbCol = oDV(iCount)(4).ToString.Trim
                        Else
                            dReimbCol = 0.0
                        End If
                        dTotalAmt = dTotalAmt + dReimbCol
                    End If
                Next

                If dTotalAmt > 0.0 Then
                    sTrnsAcct = GetBankTrnsAcct(sDBCode)

                    oPayment.TransferAccount = sTrnsAcct
                    oPayment.TransferDate = CDate(sBatchPeriod)
                    oPayment.TransferSum = dTotalAmt

                    lRetCode = oPayment.Add()

                    If lRetCode <> 0 Then
                        sErrDesc = p_oCompany.GetLastErrorDescription()
                        Throw New ArgumentException(sErrDesc)
                    End If
                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Create_OutgoingPayment = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message.ToString()
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Completed with ERROR", sFuncName)
            Create_OutgoingPayment = RTN_ERROR
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oPayment)
        End Try
    End Function

End Module
