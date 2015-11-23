Imports System.Data.SqlClient

Module modAutoPostDO
    Sub Main()
        Dim sFuncName As String = String.Empty
        Dim sErrDesc As String = String.Empty
        Dim oDSDODrafts As DataSet = Nothing
        Dim sQueryString As String = String.Empty
        Dim sDocKey As String = String.Empty

        Try
            p_iDebugMode = DEBUG_ON

            sFuncName = "Main()"
            Console.WriteLine("Starting Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            'Getting the Parameter Values from App Cofig File
            Console.WriteLine("Calling GetSystemIntializeInfo() ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling GetSystemIntializeInfo()", sFuncName)
            If GetSystemIntializeInfo(p_oCompDef, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            p_sSAPConnString = String.Empty

            p_sSAPConnString = "Data Source=" & p_oCompDef.p_sServerName & ";Initial Catalog=" & p_oCompDef.p_sDataBaseName & ";User ID=" & p_oCompDef.p_sDBUserName & "; Password=" & p_oCompDef.p_sDBPassword

            'For Converting the DO Drafts to Actual for EASI: 

            '================================================   Starting the Function   ========================================================

            'Fetching the Values from the  SAP

            sQueryString = "select T0.DocEntry,T0.DocDate   from ODRF T0 WITH (NOLOCK)  " & _
                    " INNER JOIN DRF1 T1 WITH (NOLOCK) ON T0.DocEntry =T1.DocEntry " & _
                    " where  CardCode ='" & P_sCardCode & "' AND DocStatus ='O' AND T1.BaseType ='13' " & _
                    " GROUP BY T0.DocEntry ,T0.DocDate,T1.WhsCode ORDER BY T0.DocDate,T1.WhsCode ASC"

            'sQueryString = "select T0.DocEntry,T0.DocDate   from ODRF T0 WITH (NOLOCK)  " & _
            '        " INNER JOIN DRF1 T1 WITH (NOLOCK) ON T0.DocEntry =T1.DocEntry " & _
            '        " where T0.DocEntry ='1488' and CardCode ='" & P_sCardCode & "' AND DocStatus ='O'" & _
            '        " GROUP BY T0.DocEntry ,T0.DocDate ORDER BY T0.DocDate ASC "


            Console.WriteLine("Calling Get_DataSet() from SAP DB ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Get_DataSet()", sFuncName)

            'Getting the Query Result in DataSet
            oDSDODrafts = Get_DataSet(sQueryString, p_sSAPConnString, sErrDesc)


            If Not oDSDODrafts Is Nothing Then

                'Function to connect the Company
                If p_oCompany Is Nothing Then
                    Console.WriteLine("Calling ConnectToTargetCompany() ", sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectToTargetCompany()", sFuncName)
                    If ConnectToTargetCompany(p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                End If

                'If Company Connected then Call the Function to Convert Actual Document from Drafts
                If Not p_oCompany Is Nothing Then
                    For iDataCount As Integer = 0 To oDSDODrafts.Tables(0).Rows.Count - 1

                        sDocKey = oDSDODrafts.Tables(0).Rows(iDataCount)(0).ToString().Trim()

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling UpdateBatchNumber()", sFuncName)
                        If UpdateBatchNumber(sDocKey, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Continue For

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConvertDraftToDocument()", sFuncName)
                        If ConvertDraftToDocument(sDocKey, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Continue For

                    Next
                   
                End If
            Else
                Console.WriteLine("No DO Draft's Found in SAP... ", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No DO Draft's Found in SAP... ", sFuncName)
            End If

            '================================================= End the Function ================================================================

            Console.WriteLine("Completed with SUCCESS ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            Console.WriteLine("Completed with ERROR : " & sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)

        Finally

            'Releasing the SAP Objects:
            If Not p_oCompany Is Nothing Then
                p_oCompany.Disconnect()
                p_oCompany = Nothing
            End If

        End Try
    End Sub

    Function Get_DataSet(ByVal sQueryString As String, ByVal sConnString As String, ByRef sErrDesc As String) As DataSet

        Dim oDataSet As DataSet
        Dim oDataAdapter As SqlDataAdapter
        Dim oDataTable As DataTable
        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "EventOrder_DataSet()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Query : " & sQueryString, sFuncName)

            oDataSet = New DataSet

            'To Get the Dataset Based on the Query String :

            oDataAdapter = New SqlDataAdapter(sQueryString, sConnString)
            oDataTable = New DataTable
            oDataAdapter.Fill(oDataTable)
            oDataSet.Tables.Add(oDataTable)

            If oDataSet.Tables(0).Rows.Count > 0 Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                Return oDataSet
            Else
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                Return Nothing
            End If


        Catch ex As Exception
            sErrDesc = ex.Message.ToString()
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Return Nothing
        End Try
    End Function

    Function UpdateBatchNumber(ByVal sDocKey As String, ByRef oDICompany As SAPbobsCOM.Company _
                               , ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim lRetCode As Double
        Dim oDraft As SAPbobsCOM.Documents
        Dim iRowCount As Int32 = 0
        Dim sItemCode As String = String.Empty
        Dim sWhscode As String = String.Empty
        Dim sManBatch As String = String.Empty
        Dim sQuery As String = String.Empty
        Dim dReceiptQty As Double = 0
        Dim oDTBatch As DataTable = New DataTable
        Dim dBalBatchQty As Double = 0
        Dim dBatchQty As Double = 0
        Dim dBalReceiptQty As Double = 0

        Try
            sFuncName = "UpdateBatchNumber()"
            Console.WriteLine("Please wait... while updating the batch no. Draft No :  " & sDocKey, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            oDraft = oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)

            oDraft.GetByKey(sDocKey)

            iRowCount = oDraft.Lines.Count

            For iRow As Integer = 0 To iRowCount - 1

                oDraft.Lines.SetCurrentLine(iRow)
                sItemCode = oDraft.Lines.ItemCode
                sWhscode = oDraft.Lines.WarehouseCode
                dReceiptQty = oDraft.Lines.Quantity

                dBalReceiptQty = dReceiptQty

                sQuery = "SELECT T0.[ManBtchNum] FROM OITM T0 WITH(NOLOCK) WHERE T0.[ItemCode] ='" & sItemCode & "'"

                sManBatch = GetSingleValue(sQuery, oDICompany, sErrDesc)

                If sManBatch.ToString().ToUpper() = "Y" Then

                    sQuery = "SELECT BatchNum ,Quantity , SysNumber  FROM OIBT WITH (NOLOCK) WHERE ItemCode ='" & sItemCode & "' and Quantity >0 " & _
                                           "AND WhsCode ='" & sWhscode & "' ORDER BY InDate ASC "

                    oDTBatch = ExecuteSQLQuery_DT(sQuery, sErrDesc)
                    If Not oDTBatch Is Nothing Then

                        For iBCount As Integer = 0 To oDTBatch.Rows.Count - 1

                            If dBalBatchQty >= dReceiptQty And iBCount > 0 Then Exit For

                            oDraft.Lines.BatchNumbers.BatchNumber = oDTBatch.Rows(iBCount)("BatchNum").ToString().Trim()
                            oDraft.Lines.BatchNumbers.AddmisionDate = Convert.ToDateTime(DateTime.Today).Date

                            dBatchQty = Convert.ToDouble(oDTBatch.Rows(iBCount)("Quantity").ToString().Trim())

                            If (dBalReceiptQty > dBatchQty) Then
                                oDraft.Lines.BatchNumbers.Quantity = dBatchQty
                            Else
                                oDraft.Lines.BatchNumbers.Quantity = dBalReceiptQty
                            End If

                            oDraft.Lines.BatchNumbers.Add()

                            dBalBatchQty += dBatchQty
                            dBalReceiptQty = dReceiptQty - dBatchQty

                        Next
                    Else
                        sErrDesc = "There is no batches available in SAP. ItemCode : " & sItemCode
                        Throw New ArgumentException(sErrDesc)

                    End If

                End If
            Next

            lRetCode = oDraft.Update()

            If lRetCode <> 0 Then
                sErrDesc = oDICompany.GetLastErrorDescription()
                Call WriteToLogFile(sErrDesc, sFuncName)
                Console.WriteLine("Completed With ERROR. Draft No :  " & sDocKey & "Error : " & sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With ERROR. Draft No :  " & sDocKey & "Error : " & sErrDesc, sFuncName)
            Else
                Console.WriteLine("Batch updated Successfully. Draft No :  " & sDocKey & "Error : " & sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With SUCCESS. Draft No :  " & sDocKey, sFuncName)

            End If
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            UpdateBatchNumber = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message.ToString()
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            UpdateBatchNumber = RTN_ERROR
        End Try
    End Function

    Function ConvertDraftToDocument(ByRef sDocKey As String, ByRef oDICompany As SAPbobsCOM.Company, _
                                    ByRef sErrDesc As String)
        Dim sFuncName As String = String.Empty
        Dim lRetCode As Double
        Dim oDraft As SAPbobsCOM.Documents

        Try
            sFuncName = "ConvertDraftToDocument()"
            Console.WriteLine("Statring Function... ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oDraft = oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)

            Console.WriteLine("Please wait... while Converting... Draft No :  " & sDocKey, sFuncName)

            oDraft.GetByKey(sDocKey)

            lRetCode = oDraft.SaveDraftToDocument()

            If lRetCode <> 0 Then
                sErrDesc = oDICompany.GetLastErrorDescription()
                Call WriteToLogFile(sErrDesc, sFuncName)
                Console.WriteLine("Completed With ERROR. Draft No :  " & sDocKey & "Error : " & sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With ERROR. Draft No :  " & sDocKey & "Error : " & sErrDesc, sFuncName)
            Else
                Console.WriteLine("Coverted draft to actual successfully. Draft No :  " & sDocKey & "Error : " & sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With SUCCESS. Draft No :  " & sDocKey, sFuncName)

            End If
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            ConvertDraftToDocument = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message.ToString()
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ConvertDraftToDocument = RTN_ERROR
        End Try

    End Function

    'Function ConvertDraftToDocument(ByRef oDSDODrafts As DataSet, ByRef oDICompany As SAPbobsCOM.Company, _
    '                                ByRef sErrDesc As String)
    '    Dim sFuncName As String = String.Empty
    '    Dim lRetCode As Double
    '    Dim oDraft As SAPbobsCOM.Documents
    '    Dim sDraftKey As String = String.Empty

    '    Try
    '        sFuncName = "ConvertDraftToDocument()"
    '        Console.WriteLine("Statring Function... ", sFuncName)
    '        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

    '        oDraft = oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)

    '        For iRowCount As Integer = 0 To oDSDODrafts.Tables(0).Rows.Count - 1

    '            sDraftKey = oDSDODrafts.Tables(0).Rows(0)(0).ToString()

    '            Console.WriteLine("Converting... Draft No :  " & sDraftKey, sFuncName)

    '            If oDraft.GetByKey(sDraftKey) Then

    '                lRetCode = oDraft.SaveDraftToDocument()

    '                If lRetCode <> 0 Then
    '                    sErrDesc = oDICompany.GetLastErrorDescription()
    '                    Call WriteToLogFile(sErrDesc, sFuncName)
    '                    Console.WriteLine("Completed With ERROR. Draft No :  " & sDraftKey & "Error : " & sErrDesc, sFuncName)
    '                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With ERROR. Draft No :  " & sDraftKey & "Error : " & sErrDesc, sFuncName)
    '                Else
    '                    Console.WriteLine("Completed With SUCCESS. Draft No :  " & sDraftKey & "Error : " & sErrDesc, sFuncName)
    '                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With SUCCESS. Draft No :  " & sDraftKey, sFuncName)

    '                End If
    '            Else
    '                Console.WriteLine("Draft Doesn't Exist in SAP Draft No :  " & sDraftKey & "Error : " & sErrDesc, sFuncName)
    '                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Draft Doesn't Exist in SAP Draft No : " & sDraftKey, sFuncName)

    '            End If
    '        Next
    '        Console.WriteLine("Converting Completed With SUCCESS", sFuncName)
    '        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
    '        ConvertDraftToDocument = RTN_SUCCESS

    '    Catch ex As Exception
    '        sErrDesc = ex.Message.ToString()
    '        Call WriteToLogFile(sErrDesc, sFuncName)
    '        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
    '        ConvertDraftToDocument = RTN_ERROR
    '    End Try

    'End Function

End Module
