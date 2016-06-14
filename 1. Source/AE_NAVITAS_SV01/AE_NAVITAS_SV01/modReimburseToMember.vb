Module modReimburseToMember
    Private dtBP As DataTable
    Private dtProject As DataTable
    Private dtItemCode As DataTable
    Private dtAcctCode As DataTable
    Private dtHouseBnkAct As DataTable

    Public Function ProcessReimbToMember(ByVal file As System.IO.FileInfo, ByVal odv As DataView, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = "ProcessReimbToMember"
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

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ProcessExcelDatas()", sFuncName)

                If ProcessExcelDatas(file, file.Name, odv, sDBCode, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            End If

            ProcessReimbToMember = RTN_SUCCESS

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

        Catch ex As Exception
            sErrDesc = ex.Message.ToString()
            Call WriteToLogFile(sErrDesc, sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ProcessReimbToMember = RTN_ERROR
        End Try

    End Function

    Private Function ProcessExcelDatas(ByVal file As System.IO.FileInfo, ByVal sFileName As String, ByVal oDv As DataView, ByVal sDBCode As String, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "ProcessExcelDatas"
        Dim sBatchNo As String = String.Empty
        Dim sBatchPeriod As String = String.Empty
        Dim bTransStarted As Boolean = False
        Dim oDVChkDetl As DataView = New DataView
        Dim oDVCPFDetl As DataView = New DataView

        Dim sFullBatchPeriod As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            sBatchNo = oDv(2)(0).ToString()
            sFullBatchPeriod = oDv(3)(0).ToString()

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Reading Batch no", sFuncName)
            Dim m As Integer = InStrRev(sBatchNo, ":")
            sBatchNo = Microsoft.VisualBasic.Right(sBatchNo, Len(sBatchNo) - m).Trim

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Reading Batch Period", sFuncName)
            Dim n As Integer = InStrRev(sFullBatchPeriod, "to")
            sBatchPeriod = Microsoft.VisualBasic.Right(sFullBatchPeriod, Len(sFullBatchPeriod) - n - 1).Trim

            Dim oDTGRIOGrouped As DataTable = Nothing
            Dim oDVGIRODetl As DataView = New DataView

            For iRow As Integer = 6 To oDv.Count - 1
                If (oDv(iRow)(7).ToString.Trim = String.Empty) Then
                    sErrDesc = "Payment Mode not found in some rows"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                    Throw New ArgumentException(sErrDesc)
                End If
            Next

            oDTGRIOGrouped = oDv.Table.DefaultView.ToTable(True, "F8")

            For intRow As Integer = 0 To oDTGRIOGrouped.Rows.Count - 1
                If Not (oDTGRIOGrouped.Rows(intRow).Item(0).ToString.Trim() = String.Empty Or oDTGRIOGrouped.Rows(intRow).Item(0).ToString.ToUpper.Trim = "PAYMENT MODE") Then
                    oDv.RowFilter = "F8 = '" & oDTGRIOGrouped.Rows(intRow).Item(0).ToString & "'"

                    If oDTGRIOGrouped.Rows(intRow).Item(0).ToString.ToUpper() = "GIRO" Then
                        oDVGIRODetl = New DataView(oDv.ToTable())
                    ElseIf oDTGRIOGrouped.Rows(intRow).Item(0).ToString.ToUpper() = "CHEQUE" Then
                        oDVChkDetl = New DataView(oDv.ToTable())
                    ElseIf oDTGRIOGrouped.Rows(intRow).Item(0).ToString.ToUpper() = "CPF" Then
                        oDVCPFDetl = New DataView(oDv.ToTable())
                    End If

                End If
            Next

            'oDVChkDetl.RowFilter = "F9 = '" & String.Empty & "' "
            oDVChkDetl.RowFilter = "ISNULL(F9,'') = ''"

            If oDVChkDetl.Count > 0 Then
                sErrDesc = "Cheque number not found in some rows"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            oDVChkDetl.RowFilter = Nothing

            If oDVGIRODetl.Count > 0 Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Splitting Dataview based on Payment Mode - GIRO", sFuncName)

                oDTGRIOGrouped = oDVGIRODetl.Table.DefaultView.ToTable(True, "F10")

                For intRow As Integer = 0 To oDTGRIOGrouped.Rows.Count - 1
                    If Not (oDTGRIOGrouped.Rows(intRow).Item(0).ToString.Trim() = String.Empty Or oDTGRIOGrouped.Rows(intRow).Item(0).ToString.ToUpper().Trim() = "PAYMENT MODE") Then

                        oDVGIRODetl.RowFilter = "F10 = '" & oDTGRIOGrouped.Rows(intRow).Item(0).ToString.Trim() & "' "

                        Dim odtGIRORem As DataTable = New DataTable
                        Dim sGIRORem As String = String.Empty

                        Dim odvGIROFil As DataView = New DataView(oDVGIRODetl.ToTable())

                        If odvGIROFil.Count > 0 Then
                            If bTransStarted = False Then
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling StartTransaction()", sFuncName)

                                If StartTransaction(sErrDesc) = RTN_SUCCESS Then
                                    bTransStarted = True
                                Else
                                    Throw New ArgumentException(sErrDesc)
                                End If
                            End If

                            If bTransStarted = True Then
                                Console.WriteLine("Creating outgoing payment document - GIRO")
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateOutGoingPayment_GIRO()", sFuncName)

                                If CreateOutGoingPayment_GIRO(odvGIROFil, file.Name, sDBCode, sBatchNo, sBatchPeriod, sFullBatchPeriod, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            Else
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No records found.", sFuncName)
                            End If
                        End If
                    End If
                Next
            End If

            '**********************GROUPING CHEQUE ITEMS AND CREATING OUTGOING PAYMENT*******************
            If oDVChkDetl.Count > 0 Then
                If bTransStarted = False Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling StartTransaction()", sFuncName)

                    If StartTransaction(sErrDesc) = RTN_SUCCESS Then
                        bTransStarted = True
                    Else
                        Throw New ArgumentException(sErrDesc)
                    End If
                End If

                If bTransStarted = True Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateOutGoingPayment_Cheque()", sFuncName)
                    Console.WriteLine("Creating outgoing payment document - Cheque")

                    If CreateOutGoingPayment_Cheque(oDVChkDetl, file.Name, sDBCode, sBatchNo, sBatchPeriod, sFullBatchPeriod, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                End If
            End If
          
            ''*********************CREATING OUTGOING PAYMENT FOR CPF ITEMS*******************
            If oDVCPFDetl.Count > 0 Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Splitting Dataview based on Payment Mode - CPF", sFuncName)

                oDTGRIOGrouped = oDVCPFDetl.Table.DefaultView.ToTable(True, "F10")

                For intRow As Integer = 0 To oDTGRIOGrouped.Rows.Count - 1
                    If Not (oDTGRIOGrouped.Rows(intRow).Item(0).ToString.Trim() = String.Empty Or oDTGRIOGrouped.Rows(intRow).Item(0).ToString.ToUpper().Trim() = "PAYMENT MODE") Then
                        oDVCPFDetl.RowFilter = "F10 = '" & oDTGRIOGrouped.Rows(intRow).Item(0).ToString.Trim() & "' "

                        Dim odtCPFRem As DataTable = New DataTable
                        Dim sCPFRem As String = String.Empty

                        Dim odvCPFFil As DataView = New DataView(oDVCPFDetl.ToTable())

                        If odvCPFFil.Count > 0 Then
                            If bTransStarted = False Then
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling StartTransaction()", sFuncName)

                                If StartTransaction(sErrDesc) = RTN_SUCCESS Then
                                    bTransStarted = True
                                Else
                                    Throw New ArgumentException(sErrDesc)
                                End If
                            End If

                            If bTransStarted = True Then
                                Console.WriteLine("Creating outgoing payment document - CPF")
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateOutGoingPayment_CPF()", sFuncName)

                                If CreateOutGoingPayment_CPF(odvCPFFil, file.Name, sDBCode, sBatchNo, sBatchPeriod, sFullBatchPeriod, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            Else
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No records found.", sFuncName)
                            End If
                        End If
                    End If
                Next

            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CommitTransaction", sFuncName)
            If CommitTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling FileMoveToArchive()", sFuncName)
            'FileMoveToArchive(file, file.FullName, RTN_SUCCESS)
            FileMoveToArchive_Success(file, file.FullName, file.Name, RTN_SUCCESS)

            'Insert Success Notificaiton into Table..
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
            AddDataToTable(p_oDtSuccess, file.Name, "Success")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("File successfully uploaded" & file.FullName, sFuncName)

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

    Private Function CreateOutGoingPayment_GIRO(ByVal odv As DataView, ByVal sFileName As String _
                                           , ByVal sDBCode As String, ByVal sBatchNo As String _
                                           , ByVal sBatchPeriod As String, ByVal sFullBatchPeriod As String, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateOutGoingPayment_GIRO"
        Dim iRetcode, iErrCode As Long
        Dim iCount As Integer = 1
        Dim sTrnsAcct As String = String.Empty
        Dim dTransAmount As Double = 0.0
        Dim dTotalAmt As Double = 0.0
        Dim sCardName As String = String.Empty
        Dim sAddress As String = String.Empty
        Dim sAcctCode As String = String.Empty
        Dim sCardCode As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            Dim oPayments As SAPbobsCOM.IPayments
            oPayments = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments)
            oPayments.DocType = SAPbobsCOM.BoRcptTypes.rAccount

            sCardCode = odv(0)(3).ToString.Trim()
            'sCardName = odv(0)(0).ToString.Trim()
            'sAddress = odv(0)(1).ToString.Trim()

            sCardName = String.Empty
            sAddress = String.Empty

            If sCardName = String.Empty Then
                oPayments.CardName = "GIRO Payment"
            Else
                oPayments.CardName = sCardName
            End If

            If sAddress = String.Empty Then
                oPayments.Address = "GIRO Payment"
            Else
                oPayments.Address = sAddress
            End If

            oPayments.DocDate = CDate(sBatchPeriod)
            oPayments.DueDate = CDate(sBatchPeriod)
            oPayments.CounterReference = odv(0)(9).ToString.Trim
            oPayments.Remarks = sBatchNo
            oPayments.JournalRemarks = sFullBatchPeriod
            oPayments.UserFields.Fields.Item("U_AI_APARUploadName").Value = sFileName

            oPayments.BPLID = "1"

            For i As Integer = 0 To odv.Count - 1
                '' If Not (odv(i)(0).ToString = String.Empty) Then

                If Not (odv(i)(6).ToString.Trim = String.Empty) Then
                    dTransAmount = odv(i)(6).ToString.Trim
                Else
                    dTransAmount = 0.0
                End If

                If dTransAmount = 0.0 Then Continue For

                If iCount > 1 Then
                    oPayments.AccountPayments.Add()
                End If

                sAcctCode = odv(i)(5).ToString.Trim

                If sAcctCode <> String.Empty Then
                    oPayments.AccountPayments.AccountCode = sAcctCode
                End If

                If Not (odv(i)(4).ToString.Trim = String.Empty) Then
                    oPayments.AccountPayments.ProfitCenter = odv(i)(4).ToString.Trim
                End If
                oPayments.AccountPayments.UserFields.Fields.Item("U_AI_CustName").Value = odv(i)(2).ToString.Trim
                oPayments.AccountPayments.UserFields.Fields.Item("U_AI_CustCode").Value = odv(i)(3).ToString.Trim
                oPayments.AccountPayments.UserFields.Fields.Item("U_AI_PayMode").Value = odv(i)(7).ToString.Trim
                oPayments.AccountPayments.UserFields.Fields.Item("U_AI_PayeeAcNo").Value = odv(i)(8).ToString.Trim

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("AMT IS. " & dTransAmount & "Line Number. " & i, sFuncName)

                dTotalAmt = dTotalAmt + dTransAmount

                oPayments.AccountPayments.GrossAmount = dTransAmount

                iCount = iCount + 1
                ''End If
            Next

            If dTotalAmt > 0.0 Then
                'sTrnsAcct = GetBankTrnsAcct(sDBCode)
                If Not sCardCode = String.Empty Then
                    sTrnsAcct = GetHouseBankAccout_GIRO(sCardCode, p_oCompany.CompanyDB)
                    If sTrnsAcct = String.Empty Then
                        sTrnsAcct = GetBankTrnsAcct(sDBCode)
                    End If
                Else
                    sTrnsAcct = GetBankTrnsAcct(sDBCode)
                End If

                oPayments.TransferAccount = sTrnsAcct
                oPayments.TransferDate = CDate(sBatchPeriod)
                oPayments.TransferSum = dTotalAmt
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("DB CODE IS " & sDBCode, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("TOTAL TRANSC. AMT IS " & dTotalAmt, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("GIRO GL ACCT " & sTrnsAcct, sFuncName)


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

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateOutGoingPayment_GIRO = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message.ToString()
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateOutGoingPayment_GIRO = RTN_ERROR
        End Try

    End Function

    Private Function CreateOutGoingPayment_Cheque(ByVal odv As DataView, ByVal sFileName As String, ByVal sDBCode As String, ByVal sBatchNo As String, ByVal sBatchPeriod As String, ByVal sFullBatchPeriod As String, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateOutGoingPayment_Cheque"
        Dim iCount As Integer = 0
        Dim iRetCode, iErrCode As Long
        Dim dAmount As Double = 0.0
        Dim sBankCntryCode As String = String.Empty
        Dim sBankCode As String = String.Empty
        Dim sChkGLAcct As String = String.Empty
        Dim sChkAcct As String = String.Empty
        Dim sAcctCode As String = String.Empty
        Dim iCheckNum As Integer = 0
        Dim sCardCode As String = String.Empty
        Dim sSql As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            Dim oPayments As SAPbobsCOM.IPayments


            For i As Integer = 0 To odv.Count - 1
                If Not (odv(i)(0).ToString = String.Empty) Then

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

                    oPayments = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments)
                    oPayments.DocType = SAPbobsCOM.BoRcptTypes.rAccount

                    sCardCode = odv(i)(3).ToString.Trim

                    oPayments.CardName = odv(i)(0).ToString.Trim
                    oPayments.Address = odv(i)(1).ToString.Trim
                    oPayments.DocDate = CDate(sBatchPeriod)
                    oPayments.DueDate = CDate(sBatchPeriod)
                    oPayments.CounterReference = odv(i)(9).ToString.Trim
                    oPayments.Remarks = sBatchNo
                    oPayments.JournalRemarks = sFullBatchPeriod
                    oPayments.UserFields.Fields.Item("U_AI_APARUploadName").Value = sFileName

                    oPayments.BPLID = "1"

                    iCount = 0
                    iCount = iCount + 1
                    If iCount > 1 Then
                        oPayments.AccountPayments.Add()
                    End If

                    sAcctCode = odv(i)(5).ToString.Trim
                    If sAcctCode <> String.Empty Then
                        oPayments.AccountPayments.AccountCode = sAcctCode
                    End If

                    If Not (odv(i)(4).ToString.Trim = String.Empty) Then
                        oPayments.AccountPayments.ProfitCenter = odv(i)(4).ToString.Trim
                    End If
                    oPayments.AccountPayments.UserFields.Fields.Item("U_AI_CustName").Value = odv(i)(2).ToString.Trim
                    oPayments.AccountPayments.UserFields.Fields.Item("U_AI_CustCode").Value = odv(i)(3).ToString.Trim
                    oPayments.AccountPayments.UserFields.Fields.Item("U_AI_PayMode").Value = odv(i)(7).ToString.Trim
                    oPayments.AccountPayments.UserFields.Fields.Item("U_AI_PayeeAcNo").Value = odv(i)(8).ToString.Trim
                    oPayments.AccountPayments.GrossAmount = CDbl(odv(i)(6).ToString.Trim)

                    sBankCode = GetBankCode(sDBCode)
                    sBankCntryCode = GetCountryCode(sDBCode)
                    sChkGLAcct = GetCheckGLAcct(sDBCode)
                    If Not sCardCode = String.Empty Then
                        If sCardCode.Trim.ToLower = p_oCompDef.sCaiCancerCode.ToLower Then
                            sChkAcct = p_oCompDef.sCaiCancerBankAct
                            sChkGLAcct = p_oCompDef.sCaiaCancerGLCode
                        Else
                            sChkAcct = GetHouseBankAccout(sCardCode, p_oCompany.CompanyDB)
                            If sChkAcct = String.Empty Then
                                sChkAcct = GetBankAcct(sDBCode)
                            End If
                        End If
                    Else
                        sChkAcct = GetBankAcct(sDBCode)
                    End If


                    oPayments.Checks.CountryCode = sBankCntryCode
                    oPayments.Checks.BankCode = sBankCode
                    oPayments.Checks.AccounttNum = sChkAcct
                    oPayments.Checks.CheckSum = CDbl(odv(i)(6).ToString.Trim)
                    oPayments.Checks.CheckAccount = sChkGLAcct
                    oPayments.Checks.CheckNumber = odv(i)(8).ToString.Trim
                    oPayments.Checks.DueDate = CDate(sBatchPeriod)
                    oPayments.CashSum = 0
                    oPayments.TransferSum = 0

                    iCount = iCount + 1

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

    Private Function CreateOutGoingPayment_CPF(ByVal odv As DataView, ByVal sFileName As String _
                                           , ByVal sDBCode As String, ByVal sBatchNo As String _
                                           , ByVal sBatchPeriod As String, ByVal sFullBatchPeriod As String, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateOutGoingPayment_CPF"
        Dim iRetcode, iErrCode As Long
        Dim iCount As Integer = 1
        Dim sTrnsAcct As String = String.Empty
        Dim dTransAmount As Double = 0.0
        Dim dTotalAmt As Double = 0.0
        Dim sCardName As String = String.Empty
        Dim sAddress As String = String.Empty
        Dim sAcctCode As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            Dim oPayments As SAPbobsCOM.IPayments
            oPayments = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments)
            oPayments.DocType = SAPbobsCOM.BoRcptTypes.rAccount

            'sCardName = odv(0)(0).ToString.Trim()
            'sAddress = odv(0)(1).ToString.Trim()

            sCardName = String.Empty
            sAddress = String.Empty

            If sCardName = String.Empty Then
                oPayments.CardName = "CPF Payment"
            Else
                oPayments.CardName = sCardName
            End If

            If sAddress = String.Empty Then
                oPayments.Address = "CPF Payment"
            Else
                oPayments.Address = sAddress
            End If

            oPayments.DocDate = CDate(sBatchPeriod)
            oPayments.DueDate = CDate(sBatchPeriod)
            oPayments.CounterReference = odv(0)(9).ToString.Trim
            oPayments.Remarks = sBatchNo
            oPayments.JournalRemarks = sFullBatchPeriod
            oPayments.UserFields.Fields.Item("U_AI_APARUploadName").Value = sFileName

            oPayments.BPLID = "1"

            For i As Integer = 0 To odv.Count - 1
                '' If Not (odv(i)(0).ToString = String.Empty) Then

                If Not (odv(i)(6).ToString.Trim = String.Empty) Then
                    dTransAmount = odv(i)(6).ToString.Trim
                Else
                    dTransAmount = 0.0
                End If

                If dTransAmount = 0.0 Then Continue For

                If iCount > 1 Then
                    oPayments.AccountPayments.Add()
                End If

                sAcctCode = odv(i)(5).ToString.Trim

                If sAcctCode <> String.Empty Then
                    oPayments.AccountPayments.AccountCode = sAcctCode
                End If

                If Not (odv(i)(4).ToString.Trim = String.Empty) Then
                    oPayments.AccountPayments.ProfitCenter = odv(i)(4).ToString.Trim
                End If
                oPayments.AccountPayments.UserFields.Fields.Item("U_AI_CustName").Value = odv(i)(2).ToString.Trim
                oPayments.AccountPayments.UserFields.Fields.Item("U_AI_CustCode").Value = odv(i)(3).ToString.Trim
                oPayments.AccountPayments.UserFields.Fields.Item("U_AI_PayMode").Value = odv(i)(7).ToString.Trim
                oPayments.AccountPayments.UserFields.Fields.Item("U_AI_PayeeAcNo").Value = odv(i)(8).ToString.Trim

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("AMT IS. " & dTransAmount & "Line Number. " & i, sFuncName)

                dTotalAmt = dTotalAmt + dTransAmount

                oPayments.AccountPayments.GrossAmount = dTransAmount

                iCount = iCount + 1
                ''End If
            Next

            If dTotalAmt > 0.0 Then
                sTrnsAcct = GetBankTrnsAcct(sDBCode)

                oPayments.TransferAccount = sTrnsAcct
                oPayments.TransferDate = CDate(sBatchPeriod)
                oPayments.TransferSum = dTotalAmt
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Document CPF Payment", sFuncName)

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

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateOutGoingPayment_CPF = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message.ToString()
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateOutGoingPayment_CPF = RTN_ERROR
        End Try

    End Function

End Module
