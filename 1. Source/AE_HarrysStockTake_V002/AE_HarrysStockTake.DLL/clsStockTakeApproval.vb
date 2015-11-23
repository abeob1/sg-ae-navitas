
Imports System.Data.SqlClient

Public Class clsStockTakeApproval

    Public Function StockTakeApproval(ByRef oDSDataHeader As DataSet, ByRef oDICompany As SAPbobsCOM.Company, _
                                      ByVal sConnString As String, ByVal sSAPConnString As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim sQueryString As String = String.Empty
        Dim sDocEntry As String = String.Empty
        Dim oDSCompanyList As DataSet = Nothing
        Dim sDBName As String = String.Empty

        Dim sSAPConString = String.Empty

        Try
            sFuncName = "StockTakeApproval()"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)


            For iRowCount As Integer = 0 To oDSDataHeader.Tables(0).Rows.Count - 1

                sDBName = oDSDataHeader.Tables(0).Rows(iRowCount)("DBName").ToString()

                sQueryString = "select * from [@WEB_CMPDET] where U_DBName ='" & sDBName & "'"

                oDSCompanyList = Get_DataSet(sQueryString, sSAPConnString, sErrDesc)

                If oDSCompanyList Is Nothing Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("There is No Company Details in WEB Company Details " & sDBName, sFuncName)
                    Continue For

                End If

                sSAPConString = oDSCompanyList.Tables(0).Rows(0)("U_ConnString").ToString()

                oDICompany = ConnectToCompany(oDICompany, oDSCompanyList.Tables(0).Rows(0)("U_SAPUserName").ToString(), _
                                              oDSCompanyList.Tables(0).Rows(0)("U_SAPPassword").ToString(), _
                                              oDSCompanyList.Tables(0).Rows(0)("U_Server").ToString(), _
                                              oDSCompanyList.Tables(0).Rows(0)("U_LicenseServer").ToString(), _
                                              oDSCompanyList.Tables(0).Rows(0)("U_DBName").ToString(), _
                                              oDSCompanyList.Tables(0).Rows(0)("U_DBUserName").ToString(), _
                                              oDSCompanyList.Tables(0).Rows(0)("U_DBPassword").ToString(), sErrDesc)

                If oDICompany Is Nothing Then Throw New ArgumentException(sErrDesc)

                If (p_iDebugMode = DEBUG_ON) Then WriteToLogFile_Debug("Start the SAP Transaction", sFuncName)

                If Not oDICompany.InTransaction Then oDICompany.StartTransaction()

                sDocEntry = oDSDataHeader.Tables(0).Rows(iRowCount)("DocEntry").ToString()

                sQueryString = " SELECT T0.ItemCode ,T0.WhsCode ,T0.Quantity ,T0.Variance [U_Variance],T1.DocDate  FROM STK1 T0 WITH (NOLOCK)  " & _
                                " INNER JOIN OSTK T1 WITH (NOLOCK) ON T0.DocEntry =T1.DocEntry WHERE T1.SAPSyncStatus ='N' " & _
                                " AND DocStatus ='O' AND T1.DocEntry ='" & sDocEntry & "'"

                Dim oDataDetail As DataSet = Get_DataSet(sQueryString, sConnString, sErrDesc)

                '   If (p_iDebugMode = DEBUG_ON) Then WriteToLogFile_Debug("Calling Create_GoodsReceipt()", sFuncName)

                ' If Create_GoodsReceipt(oDataDetail, oDICompany, sConnString, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                If (p_iDebugMode = DEBUG_ON) Then WriteToLogFile_Debug("Calling Create_StockTakeApproval()", sFuncName)

                If Create_StockTakeApproval(oDataDetail, oDICompany, sDocEntry, sConnString, sErrDesc) <> RTN_SUCCESS Then
                    If (p_iDebugMode = DEBUG_ON) Then WriteToLogFile_Debug("Rollback the SAP Transaction", sFuncName)
                    If oDICompany.InTransaction = True Then oDICompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)

                    If (p_iDebugMode = DEBUG_ON) Then WriteToLogFile_Debug("Calling Update_Status_IntDB() ", sFuncName)
                    Update_Status_IntDB(sDocEntry, sErrDesc, "FAIL", sConnString, sErrDesc)

                    'Releasing the SAP Objects:
                    If Not oDICompany Is Nothing Then
                        oDICompany.Disconnect()
                        oDICompany = Nothing
                    End If

                    Continue For

                End If

                If (p_iDebugMode = DEBUG_ON) Then WriteToLogFile_Debug("Calling Update_Status()", sFuncName)
                If Update_StockTakeStatus(sDocEntry, sConnString, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                If (p_iDebugMode = DEBUG_ON) Then WriteToLogFile_Debug("Calling Update_Status_IntDB()", sFuncName)
                If Update_Status_IntDB(sDocEntry, "SUCCESS", "SUCCESS", sConnString, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                If (p_iDebugMode = DEBUG_ON) Then WriteToLogFile_Debug("Committed the SAP Transaction", sFuncName)

                If oDICompany.InTransaction = True Then oDICompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)

                'Releasing the SAP Objects:
                If Not oDICompany Is Nothing Then
                    oDICompany.Disconnect()
                    oDICompany = Nothing
                End If

            Next

            ' End If

            If (p_iDebugMode = DEBUG_ON) Then WriteToLogFile_Debug("Completed With SUCCESS", sFuncName)

            StockTakeApproval = RTN_SUCCESS


        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)

            If (p_iDebugMode = DEBUG_ON) Then WriteToLogFile_Debug("Rollback the SAP Transaction", sFuncName)
            If oDICompany.InTransaction = True Then oDICompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)

            If (p_iDebugMode = DEBUG_ON) Then WriteToLogFile_Debug("Calling Update_Status_IntDB() ", sFuncName)
            Update_Status_IntDB(sDocEntry, sErrDesc, "FAIL", sConnString, sErrDesc)
            StockTakeApproval = RTN_ERROR

        Finally

            If Not oDICompany Is Nothing Then
                oDICompany.Disconnect()
                oDICompany = Nothing
            End If

        End Try


    End Function

    Public Function Create_StockTakeApproval(ByVal oDSDataset As DataSet, ByVal oDICompany As SAPbobsCOM.Company, _
                                             ByVal sDODraftNum As String, ByVal sConnString As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim sQueryString As String = String.Empty
        Dim oDTHeader As DataTable = New DataTable
        Dim oStockTaking As SAPbobsCOM.StockTaking
        Dim sBinCode As String = String.Empty
        Dim sWhsCode As String = String.Empty
        Dim dtCountDate As DateTime

        Try

            sFuncName = "StockTakeApproval()"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            oDTHeader = oDSDataset.Tables(0)

            oStockTaking = oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTakings)

            Dim oCS As SAPbobsCOM.CompanyService = oDICompany.GetCompanyService()
            Dim oICS As SAPbobsCOM.InventoryCountingsService = CType((oCS.GetBusinessService(SAPbobsCOM.ServiceTypes.InventoryCountingsService)), SAPbobsCOM.InventoryCountingsService)
            Dim oIC As SAPbobsCOM.InventoryCounting = oICS.GetDataInterface(SAPbobsCOM.InventoryCountingsServiceDataInterfaces.icsInventoryCounting)

            sWhsCode = oDTHeader.Rows(0)("WhsCode").ToString()
            Dim sPostingDate() As String = oDTHeader.Rows(0)("DocDate").ToString().Split("")
            dtCountDate = Convert.ToDateTime(sPostingDate(0))


            If sWhsCode.ToUpper().Trim().ToUpper() = "01CKT" Then
                sQueryString = "SELECT AbsEntry from OBIN WHERE BinCode ='" & sBINLocation & "'"
                sBinCode = GetSingleValue(sQueryString, oDICompany, sErrDesc)
            End If

            For iRow As Integer = 0 To oDTHeader.Rows.Count - 1

                oIC.CountDate = dtCountDate
                oIC.UserFields.Item("U_stocktakeDODraft").Value = sDODraftNum

                Dim oICLS As SAPbobsCOM.InventoryCountingLines = oIC.InventoryCountingLines
                Dim oICL As SAPbobsCOM.InventoryCountingLine = oICLS.Add()

                oICL.ItemCode = oDTHeader.Rows(iRow)("ItemCode").ToString()
                oICL.CountedQuantity = CDbl(oDTHeader.Rows(iRow)("Quantity").ToString())
                oICL.WarehouseCode = sWhsCode
                oICL.CostingCode = sWhsCode

                If sWhsCode.ToUpper().Trim().ToUpper() = "01CKT" Then
                    oICL.BinEntry = CInt(sBinCode)
                End If

                oICL.Counted = SAPbobsCOM.BoYesNoEnum.tYES

                ''=========================================================================================

                'oStockTaking.ItemCode = oDTHeader.Rows(iRow)("ItemCode").ToString()
                'oStockTaking.WarehouseCode = oDTHeader.Rows(iRow)("WhsCode").ToString()
                'oStockTaking.Counted = CDbl(oDTHeader.Rows(iRow)("Quantity").ToString())

                'oStockTaking.Add()

                ''=========================================================================================

            Next

            Dim oICP As SAPbobsCOM.InventoryCountingParams = oICS.Add(oIC)

            If (p_iDebugMode = DEBUG_ON) Then WriteToLogFile_Debug("Completed With SUCCESS", sFuncName)

            Create_StockTakeApproval = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            Console.WriteLine("Completed with ERROR : " & sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Create_StockTakeApproval = RTN_ERROR
        End Try
    End Function

    Public Function Create_GoodsReceipt(ByVal oDataSet As DataSet, ByVal oDICompany As SAPbobsCOM.Company _
                                        , ByVal sConnString As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim dtDocDate As Date
        Dim oDTHeader As DataTable = New DataTable
        Dim lRetCode As Double = 0
        Dim sItemCode As String = String.Empty
        Dim sWhscode As String = String.Empty

        Dim oDSData As DataSet = New DataSet
        Dim oDTData As DataTable = New DataTable
        Dim dQuantity As Double
        Dim bIsDataExist As Boolean = False


        Try

            sFuncName = "Create_GoodsReceipt()"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            Dim oGoodsReceipt As SAPbobsCOM.Documents
            oGoodsReceipt = oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry)

            dtDocDate = Convert.ToDateTime(DateTime.Today)
            oDTHeader = oDataSet.Tables(0)

            oGoodsReceipt.DocDate = dtDocDate
            oGoodsReceipt.DocDueDate = dtDocDate
            oGoodsReceipt.TaxDate = dtDocDate

            For iRow As Integer = 0 To oDTHeader.Rows.Count - 1

                'oDSData = Get_DataSet("", sConnString, sErrDesc)

                If (Convert.ToDouble(oDTHeader.Rows(iRow)("U_Variance").ToString()) < 0) Then Continue For
                If (Convert.ToDouble(oDTHeader.Rows(iRow)("Quantity").ToString()) < 0) Then Continue For

                bIsDataExist = True


                sItemCode = oDTHeader.Rows(iRow)("ItemCode").ToString()
                sWhscode = oDTHeader.Rows(iRow)("WhsCode").ToString()
                dQuantity = Convert.ToDouble(oDTHeader.Rows(iRow)("Quantity").ToString())


                oGoodsReceipt.Lines.ItemCode = sItemCode
                oGoodsReceipt.Lines.WarehouseCode = sWhscode
                oGoodsReceipt.Lines.Quantity = Math.Abs(dQuantity)
                oGoodsReceipt.Lines.AccountCode = "130252"
                oGoodsReceipt.Lines.CostingCode = sWhscode
                oGoodsReceipt.Lines.COGSCostingCode = sWhscode


                If (p_iDebugMode = DEBUG_ON) Then WriteToLogFile_Debug("Calling GetSingleValue() for Checking the Batch Item", sFuncName)

                If (GetSingleValue("select ManBtchNum  from OITM where ItemCode='" & sItemCode & "'",
                                                    oDICompany, sErrDesc).ToString().ToUpper() = "Y") Then
                    oGoodsReceipt.Lines.BatchNumbers.BatchNumber = sWhscode
                    oGoodsReceipt.Lines.BatchNumbers.Quantity = dQuantity
                    oGoodsReceipt.Lines.BatchNumbers.AddmisionDate = Convert.ToDateTime(DateTime.Today)
                    oGoodsReceipt.Lines.BatchNumbers.Add()

                End If

                oGoodsReceipt.Lines.Add()

            Next

            If bIsDataExist = True Then
                lRetCode = oGoodsReceipt.Add()
                If lRetCode <> 0 Then
                    sErrDesc = oDICompany.GetLastErrorDescription()
                    Throw New ArgumentException(sErrDesc)
                End If

            End If

            If (p_iDebugMode = DEBUG_ON) Then WriteToLogFile_Debug("Completed With SUCCESS", sFuncName)

            Create_GoodsReceipt = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            Console.WriteLine("Completed with ERROR : " & sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Create_GoodsReceipt = RTN_ERROR
        End Try


    End Function

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

    Function Update_StockTakeStatus(ByVal sDocEntry As String, ByVal sConnString As String, ByRef sErrDesc As String) As Long
        Dim oDataSet As DataSet
        Dim oDataAdapter As SqlDataAdapter
        Dim oDataTable As DataTable
        Dim sFuncName As String = String.Empty
        Dim sQueryString As String = String.Empty

        Try
            sFuncName = "Update_StockTakeStatus()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            sQueryString = "UPDATE OSTK SET DocStatus ='C',SAPSyncStatus ='Y' WHERE DocEntry ='" & sDocEntry & "' AND DocStatus ='O'"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Query : " & sQueryString, sFuncName)

            oDataSet = New DataSet

            'To Get the Dataset Based on the Query String :

            oDataAdapter = New SqlDataAdapter(sQueryString, sConnString)
            oDataTable = New DataTable
            oDataAdapter.Fill(oDataTable)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Update_StockTakeStatus = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message.ToString()
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Update_StockTakeStatus = RTN_ERROR
        End Try
    End Function

    Function Update_Status_IntDB(ByVal sDocEntry As String, ByVal sErrMsg As String, _
                                 ByVal sStatus As String, ByVal sConnString As String, ByRef sErrDesc As String) As Long
        Dim oDataSet As DataSet
        Dim oDataAdapter As SqlDataAdapter
        Dim oDataTable As DataTable
        Dim sFuncName As String = String.Empty
        Dim sQueryString As String = String.Empty

        Try
            sFuncName = "Update_Status_IntDB()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            sErrMsg = Replace(sErrMsg, "'", "")

            sQueryString = "update StockTakeApproval set [ReceiveDate]=GETDATE(),[ErrMsg]='" & sErrMsg & "',[Status]='" & sStatus & "' where  [DocEntry]='" & sDocEntry & "'"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Query : " & sQueryString, sFuncName)

            oDataSet = New DataSet

            'To Get the Dataset Based on the Query String :

            oDataAdapter = New SqlDataAdapter(sQueryString, sConnString)
            oDataTable = New DataTable
            oDataAdapter.Fill(oDataTable)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Update_Status_IntDB = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message.ToString()
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Update_Status_IntDB = RTN_ERROR
        End Try
    End Function

    Function ConvertDraftToDocument(ByRef oDSDODrafts As DataSet, ByRef oDICompany As SAPbobsCOM.Company, _
                                    ByRef sErrDesc As String)
        Dim sFuncName As String = String.Empty
        Dim lRetCode As Double
        Dim oDraft As SAPbobsCOM.Documents

        Try
            sFuncName = "ConvertDraftToDocument()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oDraft = oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)


            For iRowCount As Integer = 0 To oDSDODrafts.Tables(0).Rows.Count - 1

                If oDraft.GetByKey(oDSDODrafts.Tables(0).Rows(0)(0).ToString()) Then

                    lRetCode = oDraft.SaveDraftToDocument()

                    If lRetCode <> 0 Then
                        sErrDesc = oDICompany.GetLastErrorDescription()
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                    End If
                Else

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("There is No Records Found in Draft : ", sFuncName)

                End If

            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            ConvertDraftToDocument = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message.ToString()
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ConvertDraftToDocument = RTN_ERROR

        Finally
            'Releasing the SAP Objects:
            If Not oDICompany Is Nothing Then
                oDICompany.Disconnect()
                oDICompany = Nothing
            End If
        End Try

    End Function

    Public Function ConnectToCompany(ByRef oCompany As SAPbobsCOM.Company, ByVal sUserName As String, ByVal sPassword As String, _
                                    ByVal sServer As String, ByVal sLicenseServer As String, ByVal sDBName As String, _
                                   ByVal sDBUserName As String, ByVal sDBPassword As String, ByRef sErrDesc As String) As SAPbobsCOM.Company


        Dim sFuncName As String = String.Empty
        Dim iRetValue As Integer = -1
        Dim iErrCode As Integer = -1
        Dim sSQL As String = String.Empty
        Dim oDs As New DataSet

        Try
            sFuncName = "ConnectToCompany()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing the Company Object", sFuncName)

            oCompany = New SAPbobsCOM.Company

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assigning the representing database name", sFuncName)

            oCompany.Server = sServer
            oCompany.LicenseServer = sLicenseServer
            oCompany.DbUserName = sDBUserName
            oCompany.DbPassword = sDBPassword

            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English
            oCompany.UseTrusted = False
            oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008

            oCompany.CompanyDB = sDBName
            oCompany.UserName = sUserName
            oCompany.Password = sPassword

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connecting to the Company Database.", sFuncName)

            iRetValue = oCompany.Connect()

            If iRetValue <> 0 Then
                oCompany.GetLastError(iErrCode, sErrDesc)

                sErrDesc = String.Format("Connection to Database ({0}) {1} {2} {3}", _
                    oCompany.CompanyDB, System.Environment.NewLine, _
                                vbTab, sErrDesc)

                Throw New ArgumentException(sErrDesc)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            ConnectToCompany = oCompany
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ConnectToCompany = Nothing
        End Try
    End Function

    Public Function GetSingleValue(ByVal Query As String, ByRef p_oDICompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As String

        ' ***********************************************************************************
        '   Function   :    GetSingleValue()
        '   Purpose    :    This function is handles - Return single value based on Query
        '   Parameters :    ByVal Query As String
        '                       sDate = Passing Query 
        '                   ByRef oCompany As SAPbobsCOM.Company
        '                       oCompany = Passing the Company which has been connected
        '                   ByRef sErrDesc As String
        '                       sErrDesc=Error Description to be returned to calling function
        '   Author     :    SRINIVASAN
        '   Date       :    15/08/2014 
        '   Change     :   
        '                   
        ' ***********************************************************************************

        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "GetSingleValue()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Query : " & Query, sFuncName)

            Dim objRS As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRS.DoQuery(Query)
            If objRS.RecordCount > 0 Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
                GetSingleValue = RTN_SUCCESS

                Return objRS.Fields.Item(0).Value.ToString
            End If
        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            WriteToLogFile(ex.Message, sFuncName, sLogFolderPath)
            GetSingleValue = RTN_SUCCESS
            Return ""
        End Try
        Return Nothing
    End Function

End Class
