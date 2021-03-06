﻿Module modAPInvToProvider

    Private dtBP As DataTable
    Private dtCostCenter As DataTable
    Private dtItemCode As DataTable
    Private dtAcctCode As DataTable
    Private dtVendRefNo As DataTable

    Public Function ProcessAPInvToProvider(ByVal file As System.IO.FileInfo, ByVal odv As DataView, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = "ProcessAPInvToProvider"
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
                dtCostCenter = ExecuteQueryReturnDataTable(sSQL, p_oCompany.CompanyDB)

                sSQL = "SELECT ""AcctCode"" FROM ""OACT"" "
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING  SQL :" & sSQL, sFuncName)
                dtAcctCode = ExecuteQueryReturnDataTable(sSQL, p_oCompany.CompanyDB)

                sSQL = "SELECT T0.""GroupCode"" ,T0.""GroupName"" FROM ""OCRG"" T0 WHERE  T0.""Locked""='N'"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING  SQL :" & sSQL, sFuncName)
                p_oDtBPGroup = ExecuteQueryReturnDataTable(sSQL, p_oCompany.CompanyDB)

                sSQL = "SELECT T0.""GroupNum"", T0.""PymntGroup"" FROM ""OCTG"" T0"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING  SQL :" & sSQL, sFuncName)
                p_oDtPayTerms = ExecuteQueryReturnDataTable(sSQL, p_oCompany.CompanyDB)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddAPInvoice_CreditNote()", sFuncName)

                Console.WriteLine("Creating A/p / Credit Memo Document")
                If ProcessExcelDatas(file, file.Name, odv, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            ProcessAPInvToProvider = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message.ToString()
            Call WriteToLogFile(sErrDesc, sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ProcessAPInvToProvider = RTN_ERROR
        End Try

    End Function

    Private Function ProcessExcelDatas(ByVal file As System.IO.FileInfo, ByVal sFileName As String, ByVal oDv As DataView, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "ProcessExcelDatas"
        Dim sBatchNo As String = String.Empty
        Dim sBatchPeriod As String = String.Empty
        Dim sFullBatchPeriod As String = String.Empty
        Dim bTransStarted As Boolean = False

        Dim dReimbAmt As Double = 0.0
        Dim oDtInvoiceItems As DataTable
        Dim oDtCreditMemoItems As DataTable
        oDtInvoiceItems = oDv.Table.Clone
        oDtCreditMemoItems = oDv.Table.Clone
        oDtInvoiceItems.Clear()
        oDtCreditMemoItems.Clear()

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

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Create seperate views for invoice and credit memo datas", sFuncName)

            For Each row As DataRow In oDv.Table.Rows
                If Not (row.Item(5).ToString = String.Empty And row.Item(1).ToString = String.Empty) Then
                    Try
                        dReimbAmt = row.Item(5).ToString
                        If dReimbAmt > 0.0 Then
                            oDtInvoiceItems.ImportRow(row)
                        Else
                            oDtCreditMemoItems.ImportRow(row)
                        End If
                    Catch ex As Exception
                    End Try
                End If
            Next

            Dim odvInvView As DataView
            odvInvView = New DataView(oDtInvoiceItems)

            Dim odvCrdtView As DataView
            odvCrdtView = New DataView(oDtCreditMemoItems)

            '**********************GROUP EXCEL DATAS BASED ON VENDOR CODE - A/P INVOICE*******************
            Dim oDTInvGrpData As DataTable = Nothing
            Dim sCardCode As String = String.Empty

            oDTInvGrpData = odvInvView.Table.DefaultView.ToTable(True, "F3")

            For intRow As Integer = 0 To oDTInvGrpData.Rows.Count - 1
                If Not (oDTInvGrpData.Rows(intRow).Item(0).ToString.Trim() = String.Empty Or oDTInvGrpData.Rows(intRow).Item(0).ToString.ToUpper.Trim = "CLINIC CODE") Then

                    sCardCode = oDTInvGrpData.Rows(intRow).Item(0).ToString.Trim()

                    If sCardCode.ToUpper() = "VHONG WHYE" Then
                        MsgBox("VHONG WHYE")
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Customer Code Before Filtering. CardCode : " & sCardCode, sFuncName)
                    odvInvView.RowFilter = "F3 = '" & sCardCode & "'"

                    If bTransStarted = False Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling StartTransaction()", sFuncName)
                        If StartTransaction(sErrDesc) = RTN_SUCCESS Then
                            bTransStarted = True
                        Else
                            Throw New ArgumentException(sErrDesc)
                        End If
                    End If

                    If bTransStarted = True Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddInvoiceDoc()", sFuncName)

                        If AddInvoiceDoc(odvInvView, file.Name, sBatchNo, sBatchPeriod, sFullBatchPeriod, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If
                End If
            Next

            '**********************GROUP EXCEL DATAS BASED ON VENDOR CODE - A/P CREDIT MEMO*******************
            Dim oDTCrdtMemoGrpData As DataTable = Nothing

            oDTCrdtMemoGrpData = odvCrdtView.Table.DefaultView.ToTable(True, "F3")

            For intRow As Integer = 0 To oDTCrdtMemoGrpData.Rows.Count - 1
                If Not (oDTCrdtMemoGrpData.Rows(intRow).Item(0).ToString.Trim() = String.Empty Or oDTCrdtMemoGrpData.Rows(intRow).Item(0).ToString.ToUpper.Trim = "CLINIC CODE") Then
                    odvCrdtView.RowFilter = "F3 = '" & oDTCrdtMemoGrpData.Rows(intRow).Item(0).ToString & "'"

                    If bTransStarted = False Then

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling StartTransaction()", sFuncName)

                        If StartTransaction(sErrDesc) = RTN_SUCCESS Then
                            bTransStarted = True
                        Else
                            Throw New ArgumentException(sErrDesc)
                        End If
                    End If

                    If bTransStarted = True Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddCreditMemoDoc()", sFuncName)

                        If AddCreditMemoDoc(odvCrdtView, file.Name, sBatchNo, sBatchPeriod, sFullBatchPeriod, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If

                End If
            Next

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

    Private Function AddInvoiceDoc(ByVal odv As DataView, ByVal sFileName As String, ByVal sBatchNo As String, ByVal sBatchPeriod As String, ByVal sFullBatchPeriod As String, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "AddInvoiceDoc"
        Dim sCardCode As String = String.Empty
        Dim sFullCardCode As String = String.Empty
        Dim oDoc As SAPbobsCOM.Documents = Nothing
        Dim iCount As Integer = 0
        Dim bLineAdded As Boolean = False
        Dim iRetcode, iErrCode As Integer
        Dim sCardName As String = String.Empty
        Dim sNumAtCard As String = String.Empty
        Dim sSql As String = String.Empty

        Try

            Console.WriteLine("Creating A/p Invoice..")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating Purchase Invoice Object", sFuncName)
            oDoc = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)

            sFullCardCode = odv(0)(2).ToString.Trim

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("CardCode is " & sFullCardCode, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("CardCode Length is " & sFullCardCode.Length, sFuncName)

            If sFullCardCode.Length > 15 Then
                sCardCode = sFullCardCode.Substring(0, 15).ToUpper()
            Else
                sCardCode = sFullCardCode.ToUpper()
            End If

            sSql = "SELECT ""CardCode"" FROM " & p_oCompany.CompanyDB & ".""OCRD"" WHERE UPPER(""CardCode"") = '" & sCardCode & "'"
            Dim oDs As DataSet
            oDs = ExecuteSQLQuery(sSql)
            If oDs.Tables(0).Rows.Count > 0 Then
                sCardCode = oDs.Tables(0).Rows(0).Item("CardCode").ToString
            End If

            sCardName = odv(0)(1).ToString.Trim

            dtBP.DefaultView.RowFilter = "CardCode = '" & sCardCode & "'"
            If dtBP.DefaultView.Count = 0 Then
                'sErrDesc = "Cardcode ::" & sCardCode & " provided does not exist in SAP."
                'Call WriteToLogFile(sErrDesc, sFuncName)
                'Throw New ArgumentException(sErrDesc)
                If CheckBP(sFullCardCode, sCardCode, sCardName, "V", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            End If

            sNumAtCard = odv(0)(0).ToString.Trim

            sSql = "SELECT DISTINCT ""NumAtCard"" FROM ""OPCH"" WHERE IFNULL(""NumAtCard"",'') <> '' AND ""CardCode"" = '" & sCardCode & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSql, sFuncName)
            dtVendRefNo = ExecuteQueryReturnDataTable(sSql, p_oCompany.CompanyDB)

            If Not (sNumAtCard = String.Empty) Then
                dtVendRefNo.DefaultView.RowFilter = "NumAtCard = '" & sNumAtCard & "'"
                If dtVendRefNo.DefaultView.Count = 0 Then
                    oDoc.NumAtCard = sNumAtCard
                Else
                    sErrDesc = "Vendor Ref No :: " & sNumAtCard & " already exist in SAP."
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    Throw New ArgumentException(sErrDesc)
                End If
            End If


            oDoc.CardCode = sCardCode
            oDoc.DocDate = CDate(sBatchPeriod)
            oDoc.Comments = sBatchNo
            oDoc.JournalMemo = sFullBatchPeriod
            oDoc.NumAtCard = sNumAtCard
            oDoc.UserFields.Fields.Item("U_AI_APARUploadName").Value = sFileName
            oDoc.UserFields.Fields.Item("U_AI_InvRefNo").Value = sCardName

            oDoc.BPL_IDAssignedToInvoice = "1"

            iCount = iCount + 1

            For i As Integer = 0 To odv.Count - 1
                Dim sItemCode As String = String.Empty
                Dim sCostCenter As String = String.Empty
                Dim sAcctCode As String = String.Empty
                Dim dReimbAmt As Double = 0.0

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Customer Code " & sCardCode, sFuncName)

                sAcctCode = odv(i)(4).ToString.Trim
                sCostCenter = odv(i)(3).ToString.Trim()

                If sCostCenter.ToString() <> String.Empty Then
                    dtCostCenter.DefaultView.RowFilter = "OcrCode = '" & sCostCenter & "'"
                    If dtCostCenter.DefaultView.Count = 0 Then
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


                Try
                    If Not (odv(i)(5).ToString = String.Empty) Then
                        dReimbAmt = odv(i)(5).ToString
                    Else
                        dReimbAmt = 0.0
                    End If
                Catch ex As Exception

                End Try

                If Not (dReimbAmt = 0.0) Then
                    If iCount > 1 Then
                        oDoc.Lines.Add()
                    End If

                    dtItemCode.DefaultView.RowFilter = "Name = 'Reimbursement Amount'"
                    If dtItemCode.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode ::'Reimbursement Amount' provided does not exist in SAP(Mapping Table)."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    oDoc.Lines.ItemCode = sItemCode
                    oDoc.Lines.UnitPrice = Math.Abs(dReimbAmt)

                    If Not (odv(i)(7).ToString.Trim() = String.Empty) Then
                        If (odv(i)(7).ToString.Trim = 0) Then
                            oDoc.Lines.VatGroup = p_oCompDef.sAPZeroRated
                        ElseIf (odv(i)(7).ToString.Trim = 7) Then
                            oDoc.Lines.VatGroup = p_oCompDef.sAPStdRated
                        Else
                            oDoc.Lines.VatGroup = p_oCompDef.sAPZeroRated
                        End If
                    Else
                        oDoc.Lines.VatGroup = p_oCompDef.sAPZeroRated
                    End If
                    'oDoc.Lines.TaxCode = "ZI"

                    If Not (sCostCenter = String.Empty) Then
                        oDoc.Lines.CostingCode = sCostCenter
                        oDoc.Lines.COGSCostingCode = sCostCenter
                    End If

                    If Not (sAcctCode = String.Empty) Then
                        oDoc.Lines.AccountCode = sAcctCode
                    End If

                    If Not (odv(i)(10).ToString.Trim = String.Empty) Then
                        oDoc.Comments = sBatchNo & "-" & odv(0)(10).ToString.Trim
                    End If

                    bLineAdded = True

                    iCount = iCount + 1
                End If

            Next

            If bLineAdded = True Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Document", sFuncName)

                iRetcode = oDoc.Add()

                If iRetcode <> 0 Then
                    p_oCompany.GetLastError(iErrCode, sErrDesc)
                    Throw New ArgumentException(sErrDesc)
                Else
                    Dim iDocNo As Integer
                    p_oCompany.GetNewObjectCode(iDocNo)
                    Console.WriteLine("Document Created successfully :: " & iDocNo)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDoc)
                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            AddInvoiceDoc = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message.ToString
            Call WriteToLogFile_Debug(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            AddInvoiceDoc = RTN_ERROR
        End Try
    End Function

    Private Function AddCreditMemoDoc(ByVal odv As DataView, ByVal sFileName As String, ByVal sBatchNo As String, ByVal sBatchPeriod As String, ByVal sFullBatchPeriod As String, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "AddCreditMemoDoc"
        Dim sCardCode As String = String.Empty
        Dim sFullCardCode As String = String.Empty
        Dim oDoc As SAPbobsCOM.Documents = Nothing
        Dim iCount As Integer = 0
        Dim bLineAdded As Boolean = False
        Dim iRetcode, iErrCode As Integer
        Dim sCardName As String = String.Empty
        Dim sNumAtCard As String = String.Empty
        Dim sSql As String = String.Empty

        Try

            Console.WriteLine("Creating A/P Credit Memo..")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating A/p Credit Memo Object", sFuncName)
            oDoc = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)

            sFullCardCode = odv(0)(2).ToString.Trim

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("CardCode is " & sFullCardCode, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("CardCode Length is " & sFullCardCode.Length, sFuncName)

            If sFullCardCode.Length > 15 Then
                sCardCode = sFullCardCode.Substring(0, 15).ToUpper()
            Else
                sCardCode = sFullCardCode.ToUpper()
            End If

            sSql = "SELECT ""CardCode"" FROM " & p_oCompany.CompanyDB & ".""OCRD"" WHERE UPPER(""CardCode"") = '" & sCardCode & "'"
            Dim oDs As DataSet
            oDs = ExecuteSQLQuery(sSql)
            If oDs.Tables(0).Rows.Count > 0 Then
                sCardCode = oDs.Tables(0).Rows(0).Item("CardCode").ToString
            End If

            sCardName = odv(0)(1).ToString.Trim

            dtBP.DefaultView.RowFilter = "CardCode = '" & sCardCode & "'"
            If dtBP.DefaultView.Count = 0 Then
                If CheckBP(sFullCardCode, sCardCode, sCardName, "V", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            End If

            sNumAtCard = odv(0)(0).ToString.Trim

            sSql = "SELECT DISTINCT ""NumAtCard"" FROM ""ORPC"" WHERE IFNULL(""NumAtCard"",'') <> '' AND ""CardCode"" = '" & sCardCode & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSql, sFuncName)
            dtVendRefNo = ExecuteQueryReturnDataTable(sSql, p_oCompany.CompanyDB)

            If Not (sNumAtCard = String.Empty) Then
                dtVendRefNo.DefaultView.RowFilter = "NumAtCard = '" & sNumAtCard & "'"
                If dtVendRefNo.DefaultView.Count = 0 Then
                    oDoc.NumAtCard = sNumAtCard
                Else
                    sErrDesc = "Vendor Ref No :: " & sNumAtCard & " already exist in SAP."
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    Throw New ArgumentException(sErrDesc)
                End If
            End If

            oDoc.CardCode = sCardCode
            oDoc.DocDate = CDate(sBatchPeriod)
            oDoc.Comments = sBatchNo
            oDoc.JournalMemo = sFullBatchPeriod
            oDoc.NumAtCard = sNumAtCard
            If Not (odv(0)(10).ToString.Trim = String.Empty) Then
                oDoc.Comments = sBatchNo & "-" & odv(0)(10).ToString.Trim
            End If

            oDoc.UserFields.Fields.Item("U_AI_APARUploadName").Value = sFileName
            oDoc.UserFields.Fields.Item("U_AI_InvRefNo").Value = sCardName

            oDoc.BPL_IDAssignedToInvoice = "1"

            iCount = iCount + 1

            For i As Integer = 0 To odv.Count - 1
                Dim sItemCode As String = String.Empty
                Dim sCostCenter As String = String.Empty
                Dim sAcctCode As String = String.Empty
                Dim dReimbAmt As Double = 0.0

                sAcctCode = odv(i)(4).ToString.Trim
                sCostCenter = odv(i)(3).ToString.Trim()

                If sCostCenter.ToString() <> String.Empty Then
                    dtCostCenter.DefaultView.RowFilter = "OcrCode = '" & sCostCenter & "'"
                    If dtCostCenter.DefaultView.Count = 0 Then
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

                Try
                    If Not (odv(i)(5).ToString = String.Empty) Then
                        dReimbAmt = odv(i)(5).ToString
                    Else
                        dReimbAmt = 0.0
                    End If
                Catch ex As Exception

                End Try

                If Not (dReimbAmt = 0.0) Then
                    If iCount > 1 Then
                        oDoc.Lines.Add()
                    End If

                    dtItemCode.DefaultView.RowFilter = "Name = 'Reimbursement Amount'"
                    If dtItemCode.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode ::'Reimbursement Amount' provided does not exist in SAP(Mapping Table)."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    oDoc.Lines.ItemCode = sItemCode
                    oDoc.Lines.UnitPrice = Math.Abs(dReimbAmt)

                    If Not (odv(i)(7).ToString.Trim() = String.Empty) Then
                        If (odv(i)(7).ToString.Trim = 0) Then
                            oDoc.Lines.VatGroup = p_oCompDef.sAPZeroRated
                        ElseIf (odv(i)(7).ToString.Trim = 7) Then
                            oDoc.Lines.VatGroup = p_oCompDef.sAPStdRated
                        Else
                            oDoc.Lines.VatGroup = p_oCompDef.sAPZeroRated
                        End If
                    Else
                        oDoc.Lines.VatGroup = p_oCompDef.sAPZeroRated
                    End If

                    'oDoc.Lines.TaxCode = "ZI"
                    If Not (sCostCenter = String.Empty) Then
                        oDoc.Lines.CostingCode = sCostCenter
                        oDoc.Lines.COGSCostingCode = sCostCenter
                    End If
                    If Not (sAcctCode = String.Empty) Then
                        oDoc.Lines.AccountCode = sAcctCode
                    End If

                    bLineAdded = True

                    iCount = iCount + 1
                End If

            Next

            If bLineAdded = True Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Document", sFuncName)

                iRetcode = oDoc.Add()

                If iRetcode <> 0 Then
                    p_oCompany.GetLastError(iErrCode, sErrDesc)
                    Throw New ArgumentException(sErrDesc)
                Else
                    Dim iDocNo As Integer
                    p_oCompany.GetNewObjectCode(iDocNo)
                    Console.WriteLine("Document Created successfully :: " & iDocNo)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDoc)
                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            AddCreditMemoDoc = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message.ToString
            Call WriteToLogFile_Debug(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            AddCreditMemoDoc = RTN_ERROR
        End Try
    End Function

End Module
