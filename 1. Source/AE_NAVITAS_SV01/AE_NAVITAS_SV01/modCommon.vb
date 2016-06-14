Imports System.Configuration
Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.Xml
Imports System.Data.OleDb
Imports System.IO

Imports Excel = Microsoft.Office.Interop.Excel


Module modCommon

#Region "Connection Object [Connect to DI Company]"

#Region "Get Company Initialization info"

    Public Function GetCompanyInfo(ByRef oCompDef As CompanyDefault, ByRef sErrDesc As String) As Long
        Dim sFunctName As String = String.Empty
        Dim sConnection As String = String.Empty

        Try
            sFunctName = "Get Company Initialization"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Company Initialization", sFunctName)


            oCompDef.sServer = String.Empty

            oCompDef.sLicenceServer = String.Empty
            oCompDef.sSAPUser = String.Empty
            oCompDef.sSAPPwd = String.Empty
            oCompDef.sSAPDBName = String.Empty
            oCompDef.sDBUser = String.Empty
            oCompDef.sDBPwd = String.Empty
            oCompDef.sDSN = String.Empty

            oCompDef.sInboxDir = String.Empty
            oCompDef.sSuccessDir = String.Empty
            oCompDef.sFailDir = String.Empty
            oCompDef.sLogPath = String.Empty
            oCompDef.sDebug = String.Empty

            oCompDef.sARZeroRated = String.Empty
            oCompDef.sARStdRated = String.Empty
            oCompDef.sAPZeroRated = String.Empty
            oCompDef.sAPStdRated = String.Empty
            oCompDef.sTpaOthersGlCode = String.Empty
            oCompDef.sTpaProcFeeGlCode = String.Empty

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("Server")) Then
                oCompDef.sServer = ConfigurationManager.AppSettings("Server")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("LicenceServer")) Then
                oCompDef.sLicenceServer = ConfigurationManager.AppSettings("LicenceServer")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SAPDBName")) Then
                oCompDef.sSAPDBName = ConfigurationManager.AppSettings("SAPDBName")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SAPUserName")) Then
                oCompDef.sSAPUser = ConfigurationManager.AppSettings("SAPUserName")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SAPPassword")) Then
                oCompDef.sSAPPwd = ConfigurationManager.AppSettings("SAPPassword")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBUser")) Then
                oCompDef.sDBUser = ConfigurationManager.AppSettings("DBUser")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBPwd")) Then
                oCompDef.sDBPwd = ConfigurationManager.AppSettings("DBPwd")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("InboxDir")) Then
                oCompDef.sInboxDir = ConfigurationManager.AppSettings("InboxDir")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SuccessDir")) Then
                oCompDef.sSuccessDir = ConfigurationManager.AppSettings("SuccessDir")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("FailDir")) Then
                oCompDef.sFailDir = ConfigurationManager.AppSettings("FailDir")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("LogPath")) Then
                oCompDef.sLogPath = ConfigurationManager.AppSettings("LogPath")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("EmailFrom")) Then
                oCompDef.sEmailFrom = ConfigurationManager.AppSettings("EmailFrom")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("EmailTo")) Then
                oCompDef.sEmailTo = ConfigurationManager.AppSettings("EmailTo")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("EmailSubject")) Then
                oCompDef.sEmailSubject = ConfigurationManager.AppSettings("EmailSubject")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SMTPServer")) Then
                oCompDef.sSMTPServer = ConfigurationManager.AppSettings("SMTPServer")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SMTPPort")) Then
                oCompDef.sSMTPPort = ConfigurationManager.AppSettings("SMTPPort")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SMTPUser")) Then
                oCompDef.sSMTPUser = ConfigurationManager.AppSettings("SMTPUser")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SMTPPassword")) Then
                oCompDef.sSMTPPassword = ConfigurationManager.AppSettings("SMTPPassword")
            End If


            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("CustGroup")) Then
                oCompDef.sCustomerGroup = ConfigurationManager.AppSettings("CustGroup")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("CustPayTerms")) Then
                oCompDef.sCustPayTerm = ConfigurationManager.AppSettings("CustPayTerms")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("VendorGroup")) Then
                oCompDef.sVendorGroup = ConfigurationManager.AppSettings("VendorGroup")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("VendPayTerms")) Then
                oCompDef.sVendPayTerm = ConfigurationManager.AppSettings("VendPayTerms")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("ARZeroRated")) Then
                oCompDef.sARZeroRated = ConfigurationManager.AppSettings("ARZeroRated")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("ARStdRated")) Then
                oCompDef.sARStdRated = ConfigurationManager.AppSettings("ARStdRated")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("APZeroRated")) Then
                oCompDef.sAPZeroRated = ConfigurationManager.AppSettings("APZeroRated")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("APStdRated")) Then
                oCompDef.sAPStdRated = ConfigurationManager.AppSettings("APStdRated")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("CaiaCancerCode")) Then
                oCompDef.sCaiCancerCode = ConfigurationManager.AppSettings("CaiaCancerCode")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("CaiaCancerBank")) Then
                oCompDef.sCaiCancerBankAct = ConfigurationManager.AppSettings("CaiaCancerBank")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("CaiaCancerGLCode")) Then
                oCompDef.sCaiaCancerGLCode = ConfigurationManager.AppSettings("CaiaCancerGLCode")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("TPAOthersGLAcct")) Then
                oCompDef.sTpaOthersGlCode = ConfigurationManager.AppSettings("TPAOthersGLAcct")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("TPAProcessFeeGLAcct")) Then
                oCompDef.sTpaProcFeeGlCode = ConfigurationManager.AppSettings("TPAProcessFeeGLAcct")
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with Success", sFunctName)
            GetCompanyInfo = RTN_SUCCESS

        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with Error", sFunctName)
            GetCompanyInfo = RTN_ERROR
        End Try

    End Function
#End Region

    Public Function CompanyConnection(ByRef oCompany As SAPbobsCOM.Company, ByVal sDBName As String, ByRef sErrDesc As String) As Long

        Dim sFunctName As String = String.Empty
        Dim iRetValue As Integer = -1
        Dim iErrCode As Integer = -1

        Try
            sFunctName = "Company Connection"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Company Connection", sFunctName)

            oCompany = New SAPbobsCOM.Company

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assigning DB values", sFunctName)

            oCompany.LicenseServer = p_oCompDef.sLicenceServer
            oCompany.Server = p_oCompDef.sServer
            oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012
            oCompany.CompanyDB = p_oCompDef.sSAPDBName
            oCompany.UserName = p_oCompDef.sSAPUser
            oCompany.Password = p_oCompDef.sSAPPwd
            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English
            oCompany.UseTrusted = False

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connecting to database", sFunctName)
            iRetValue = oCompany.Connect()

            If iRetValue <> 0 Then
                oCompany.GetLastError(iErrCode, sErrDesc)

                sErrDesc = String.Format("Connection to database({0}) {1} {2} {3}", oCompany.CompanyDB, System.Environment.NewLine, vbTab, sErrDesc)

                Throw New ArgumentException(sErrDesc)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Company connection Sucess", sFunctName)
            CompanyConnection = RTN_SUCCESS
        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while connecting to Company", sFunctName)
            CompanyConnection = RTN_ERROR
        End Try

    End Function

    Public Function ConnectToTargetCompany(ByRef oCompany As SAPbobsCOM.Company, _
                                            ByVal sDBCode As String, _
                                            ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   ConnectToTargetCompany()
        '   Purpose     :   This function will be providing to proceed the connectivity of 
        '                   using SAP DIAPI function
        '               
        '   Parameters  :   ByRef oCompany As SAPbobsCOM.Company
        '                       oCompany =  set the SAP DI Company Object
        '                   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author      :   SRI
        '   Date        :   October 2013
        ' **********************************************************************************

        Dim sFuncName As String = String.Empty
        Dim iRetValue As Integer = -1
        Dim iErrCode As Integer = -1
        Dim sSQL As String = String.Empty
        Dim oDs As New DataSet
        Dim sSAPUser As String = String.Empty
        Dim sSAPPWd As String = String.Empty
        Dim sTrgtDBName As String = String.Empty


        Try
            sFuncName = "ConnectToTargetCompany()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            sSQL = "SELECT * FROM ""@AI_TB01_COMPANYDATA""  WHERE ""Code"" = '" & sDBCode & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL" & sSQL, sFuncName)

            oDs = ExecuteSQLQuery(sSQL)

            If oDs.Tables(0).Rows.Count > 0 Then

                sTrgtDBName = oDs.Tables(0).Rows(0).Item("Name").ToString
                sSAPUser = oDs.Tables(0).Rows(0).Item("U_SAPUSER").ToString
                sSAPPWd = oDs.Tables(0).Rows(0).Item("U_SAPPASSWORD").ToString

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing the Company Object", sFuncName)
                oCompany = New SAPbobsCOM.Company

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assigning the representing database name : " & sTrgtDBName, sFuncName)
                oCompany.Server = p_oCompDef.sServer

                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB

                oCompany.LicenseServer = p_oCompDef.sLicenceServer
                oCompany.CompanyDB = sTrgtDBName
                oCompany.UserName = sSAPUser
                oCompany.Password = sSAPPWd

                oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English

                oCompany.UseTrusted = False

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connecting to the Company Database.", sFuncName)
                iRetValue = oCompany.Connect()

                If iRetValue <> 0 Then
                    oCompany.GetLastError(iErrCode, sErrDesc)

                    sErrDesc = String.Format("Connection to Database ({0}) {1} {2} {3}", _
                        oCompany.CompanyDB, System.Environment.NewLine, _
                                    vbTab, sErrDesc)

                    Throw New ArgumentException(sErrDesc)
                End If
            Else
                sErrDesc = "No Database login information found in COMPANYDATA Table. Please check"
                Throw New ArgumentException(sErrDesc)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connection established with " & oCompany.CompanyName, sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            ConnectToTargetCompany = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling EmailTemplate_GeneralError()", sFuncName)
            EmailTemplate_GeneralError("Unable to connect to SAP." & " " & sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ConnectToTargetCompany = RTN_ERROR
        End Try
    End Function

#End Region

#Region "Start Transaction"
    Public Function StartTransaction(ByRef sErrDesc As String) As Long
        ' ***********************************************************************************
        '   Function   :    StartTransaction()
        '   Purpose    :    Start DI Company Transaction
        '
        '   Parameters :    ByRef sErrDesc As String
        '                       sErrDesc = Error Description to be returned to calling function
        '   Return     :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author     :   Jeeva
        '   Date       :   03 Aug 2015
        '   Change     :
        ' ***********************************************************************************

        Dim sFuncName As String = "StartTransaction"
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Transaction", sFuncName)

            If p_oCompany.InTransaction Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback hanging transactions", sFuncName)
                p_oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If

            p_oCompany.StartTransaction()

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Trancation Started Successfully", sFuncName)
            StartTransaction = RTN_SUCCESS

        Catch ex As Exception
            Call WriteToLogFile_Debug(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while starting Trancation", sFuncName)
            StartTransaction = RTN_ERROR
        End Try

    End Function
#End Region

#Region "Commit Transaction"
    Public Function CommitTransaction(ByRef sErrDesc As String) As Long
        ' ***********************************************************************************
        '   Function   :    CommitTransaction()
        '   Purpose    :    Commit DI Company Transaction
        '
        '   Parameters :    ByRef sErrDesc As String
        '                       sErrDesc=Error Description to be returned to calling function
        '   Return     :    0 - FAILURE
        '                   1 - SUCCESS
        '   Author     :    Jeeva
        '   Date       :    03 Aug 2015
        '   Change     :
        ' ***********************************************************************************
        Dim sFuncName As String = "CommitTransaction"
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            If p_oCompany.InTransaction Then
                p_oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            Else
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Transaction is Active", sFuncName)
            End If

            CommitTransaction = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Commit Transaction Complete", sFuncName)
        Catch ex As Exception
            Call WriteToLogFile_Debug(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while committing Transaciton", sFuncName)
            CommitTransaction = RTN_ERROR
        End Try
    End Function
#End Region

#Region "Rollback Transaction"
    Public Function RollbackTransaction(ByRef sErrDesc As String) As Long
        ' ***********************************************************************************
        '   Function   :    RollbackTransaction()
        '   Purpose    :    Rollback DI Company Transaction
        '
        '   Parameters :    ByRef sErrDesc As String
        '                       sErrDesc = Error Description to be returned to calling function
        '   Return     :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author     :   Jeeva
        '   Date       :   31 July 2015
        '   Change     :
        ' ***********************************************************************************
        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "RollbackTransaction()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_oCompany.InTransaction Then
                p_oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            Else
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No transaction is active", sFuncName)
            End If
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with Success", sFuncName)
            RollbackTransaction = RTN_SUCCESS
        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with Error", sFuncName)
            RollbackTransaction = RTN_ERROR
        End Try

    End Function
#End Region

    Public Function GetDataViewFromExcel(ByVal CurrFileToUpload As String, ByVal sExtension As String) As DataView

        Dim conStr As String = ""
        Dim sFuncName As String = String.Empty
        Dim dv As DataView

        Try
            sFuncName = "GetDataViewFromExcel"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            Select Case sExtension
                Case ".xls"
                    'Excel 97-03
                    conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CurrFileToUpload & ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1'"
                    Exit Select
                Case ".xlsx"
                    'Excel 07
                    conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CurrFileToUpload & ";Extended Properties='Excel 12.0;HDR=NO;IMEX=1'"
                    Exit Select
            End Select

            Dim connExcel As New OleDbConnection(conStr)
            Dim cmdExcel As New OleDbCommand()
            Dim oda As New OleDbDataAdapter()
            Dim dt As New DataTable()

            cmdExcel.Connection = connExcel

            'Get the name of First Sheet
            connExcel.Open()
            Dim dtExcelSchema As DataTable
            dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
            Dim SheetName As String = dtExcelSchema.Rows(0)("TABLE_NAME").ToString()
            connExcel.Close()

            'Read Data from First Sheet
            connExcel.Open()
            cmdExcel.CommandText = "SELECT * From [" & SheetName & "]"
            oda.SelectCommand = cmdExcel
            dt = New DataTable("Data")
            oda.Fill(dt)
            connExcel.Close()

            dv = New DataView(dt)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With SUCCESS", sFuncName)

            Return dv


        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while reading content of " & ex.Message, sFuncName)
            Call WriteToLogFile_Debug(ex.Message, sFuncName)
            Return Nothing
        End Try


    End Function

    Public Function CreateDataTable(ByVal ParamArray oColumnName() As String) As DataTable
        Dim oDataTable As DataTable = New DataTable()

        Dim oDataColumn As DataColumn

        For i As Integer = LBound(oColumnName) To UBound(oColumnName)
            oDataColumn = New DataColumn()
            oDataColumn.DataType = Type.GetType("System.String")
            oDataColumn.ColumnName = oColumnName(i).ToString
            oDataTable.Columns.Add(oDataColumn)
        Next

        Return oDataTable

    End Function

    Public Sub AddDataToTable(ByVal oDt As DataTable, ByVal ParamArray sColumnValue() As String)
        Dim oRow As DataRow = Nothing
        oRow = oDt.NewRow()
        For i As Integer = LBound(sColumnValue) To UBound(sColumnValue)
            oRow(i) = sColumnValue(i).ToString
        Next
        oDt.Rows.Add(oRow)
    End Sub

    Public Function CheckBP(ByVal sFullBPCode As String, ByVal sBPCode As String, ByRef sBPName As String, ByVal sType As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim oBP As SAPbobsCOM.BusinessPartners
        Dim lRetCode, lErrCode As Long
        Dim sGroupCode As String = String.Empty
        Dim sPayTerms As String = String.Empty


        Try
            sFuncName = "CheckBP()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oBP = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Customer Code : " & sBPCode & ". Customer Name : " & sBPName, sFuncName)

            If oBP.GetByKey(sBPCode) = False Then

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("BP not exist in SAP", sFuncName)

                oBP.CardCode = sBPCode.ToUpper()
                oBP.CardName = sBPName

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Before cheking whether customer or vendor.", sFuncName)

                If sType.ToString().Trim().ToUpper() = "C" Then

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inside the customer", sFuncName)

                    oBP.CardType = SAPbobsCOM.BoCardTypes.cCustomer

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Before filter the BP group", sFuncName)
                    p_oDtBPGroup.DefaultView.RowFilter = "GroupName = '" & p_oCompDef.sCustomerGroup & "'"
                    If p_oDtBPGroup.DefaultView.Count = 0 Then
                        sErrDesc = "Customer Group ::" & p_oCompDef.sCustomerGroup & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sGroupCode = p_oDtBPGroup.DefaultView.Item(0)(0).ToString().Trim()
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Filtered BP group", sFuncName)


                    p_oDtPayTerms.DefaultView.RowFilter = "PymntGroup = '" & p_oCompDef.sCustPayTerm & "'"
                    If p_oDtPayTerms.DefaultView.Count = 0 Then
                        sErrDesc = "Payment Terms ::" & p_oCompDef.sCustPayTerm & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sPayTerms = p_oDtPayTerms.DefaultView.Item(0)(0).ToString().Trim()
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Filtered payentterms", sFuncName)
                    oBP.GroupCode = sGroupCode
                    oBP.PayTermsGrpCode = sPayTerms
                    oBP.UserFields.Fields.Item("U_PortalCode").Value = sFullBPCode

                ElseIf sType.ToString().Trim().ToUpper() = "V" Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inside the customer", sFuncName)
                    oBP.CardType = SAPbobsCOM.BoCardTypes.cSupplier

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Before filter the BP group", sFuncName)

                    p_oDtBPGroup.DefaultView.RowFilter = "GroupName = '" & p_oCompDef.sVendorGroup & "'"
                    If p_oDtBPGroup.DefaultView.Count = 0 Then
                        sErrDesc = "Vendor Group ::" & p_oCompDef.sVendorGroup & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sGroupCode = p_oDtBPGroup.DefaultView.Item(0)(0).ToString().Trim()
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Filtered BP group", sFuncName)
                    p_oDtPayTerms.DefaultView.RowFilter = "PymntGroup = '" & p_oCompDef.sVendPayTerm & "'"
                    If p_oDtPayTerms.DefaultView.Count = 0 Then
                        sErrDesc = "Payment Terms ::" & p_oCompDef.sVendPayTerm & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sPayTerms = p_oDtPayTerms.DefaultView.Item(0)(0).ToString().Trim()
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Filtered payentterms", sFuncName)
                    oBP.GroupCode = sGroupCode
                    oBP.PayTermsGrpCode = sPayTerms
                    oBP.UserFields.Fields.Item("U_PortalCode").Value = sFullBPCode
                End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding BP.", sFuncName)
                lRetCode = oBP.Add

                If lRetCode <> 0 Then
                    p_oCompany.GetLastError(lErrCode, sErrDesc)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding BP failed.", sFuncName)
                    Throw New ArgumentException(sErrDesc)
                Else
                    p_oCompany.GetNewObjectCode(sBPCode)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS.", sFuncName)
                End If
            Else
                sBPCode = sBPCode
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS.", sFuncName)
            End If
            CheckBP = RTN_SUCCESS
        Catch ex As Exception
            CheckBP = RTN_ERROR
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrDesc, sFuncName)
        Finally
            oBP = Nothing
        End Try

    End Function

#Region "Get Data View From Excel"
    Public Function GetDataViewFromExcel_Backup(ByVal CurrFileToUpload As String, ByVal sSheet As String) As DataView

        Dim sConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CurrFileToUpload & ";Extended Properties='Excel 12.0;HDR=NO;IMEX=1'"

        Dim objConn As New System.Data.OleDb.OleDbConnection(sConnectionString)
        Dim da As OleDb.OleDbDataAdapter
        Dim dt As DataTable
        Dim dv As DataView
        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "GetDataViewFromExcel"

            da = New System.Data.OleDb.OleDbDataAdapter("SELECT * FROM [" & sSheet & "$]", objConn)
            dt = New DataTable("BillToClient")

            da.Fill(dt)
            dv = New DataView(dt)
            Return dv
        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while reading content of " & ex.Message, sFuncName)
            Call WriteToLogFile_Debug(ex.Message, sFuncName)
            Return Nothing
        End Try

    End Function
#End Region

#Region "Execute SQL Query"
    Public Function ExecuteQueryReturnDataTable(ByVal sQueryString As String, ByVal sCompanyDB As String) As DataTable

        Dim sFuncName As String = "ExecuteQueryReturnDataTable"
        'Dim sConstr As String = "Data Source=" & p_oCompDef.sServer & ";Initial Catalog=" & sCompanyDB & ";User ID=" & p_oCompDef.sDBUser & "; Password=" & p_oCompDef.sDBPwd & ""
        Dim sConstr As String = "DRIVER={HDBODBC32};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";SERVERNODE=" & p_oCompDef.sServer & ";CS=" & sCompanyDB

        Dim oCmd As New Odbc.OdbcCommand
        Dim oDS As DataSet = New DataSet
        Dim oDbProviderFactoryObj As DbProviderFactory = DbProviderFactories.GetFactory("System.Data.Odbc")
        Dim Con As DbConnection = oDbProviderFactoryObj.CreateConnection()
        Dim dtDetail As DataTable = New DataTable


        Try
            Con.ConnectionString = sConstr
            Con.Open()

            oCmd.CommandText = CommandType.Text
            oCmd.CommandText = sQueryString
            oCmd.Connection = Con
            oCmd.CommandTimeout = 0

            Dim da As New Odbc.OdbcDataAdapter(oCmd)
            da.Fill(dtDetail)
            dtDetail.TableName = "Data"

            'oCmd.CommandType = CommandType.Text
            'oCmd.CommandText = sQueryString
            'oCmd.Connection = oCon
            'If oCon.State = ConnectionState.Closed Then
            '    oCon.Open()
            'End If

            'oSQLAdapter.SelectCommand = oCmd

            'oSQLAdapter.Fill(dtDetail)
            'dtDetail.TableName = "Data"

        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("ExecuteSQL Query Error", sFuncName)
            Throw New Exception(ex.Message)
        Finally
            Con.Dispose()
        End Try

        ExecuteQueryReturnDataTable = dtDetail

    End Function

    Public Function ExecuteSQLQuery(ByVal sSql As String) As DataSet
        Dim sFuncName As String = "ExecuteSQLQuery"
        Dim sErrDesc As String = String.Empty

        Dim cmd As New Odbc.OdbcCommand
        Dim ods As New DataSet
        'Dim oSQLCommand As SqlCommand = Nothing
        'Dim oSQLAdapter As New SqlDataAdapter
        Dim oDbProviderFactoryObj As DbProviderFactory = DbProviderFactories.GetFactory("System.Data.Odbc")
        Dim Con As DbConnection = oDbProviderFactoryObj.CreateConnection()

        Try

            Con.ConnectionString = "DRIVER={HDBODBC32};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";SERVERNODE=" & p_oCompDef.sServer & ";CS=" & p_oCompDef.sSAPDBName
            Con.Open()

            cmd.CommandType = CommandType.Text
            cmd.CommandText = sSql
            cmd.Connection = Con
            cmd.CommandTimeout = 0
            Dim da As New Odbc.OdbcDataAdapter(cmd)
            da.Fill(ods)
        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("ExecuteSQL Query Error", sFuncName)
            Throw New Exception(ex.Message)
        Finally
            Con.Dispose()
        End Try
        Return ods
    End Function

    Public Function ExecuteTargetCompSQLQuery(ByVal sSql As String, ByVal sCompanyDB As String) As DataSet
        Dim sFuncName As String = "ExecuteSQLQuery"
        Dim sErrDesc As String = String.Empty

        Dim cmd As New Odbc.OdbcCommand
        Dim ods As New DataSet
        'Dim oSQLCommand As SqlCommand = Nothing
        'Dim oSQLAdapter As New SqlDataAdapter
        Dim oDbProviderFactoryObj As DbProviderFactory = DbProviderFactories.GetFactory("System.Data.Odbc")
        Dim Con As DbConnection = oDbProviderFactoryObj.CreateConnection()

        Try

            Con.ConnectionString = "DRIVER={HDBODBC32};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";SERVERNODE=" & p_oCompDef.sServer & ";CS=" & sCompanyDB
            Con.Open()

            cmd.CommandType = CommandType.Text
            cmd.CommandText = sSql
            cmd.Connection = Con
            cmd.CommandTimeout = 0
            Dim da As New Odbc.OdbcDataAdapter(cmd)
            da.Fill(ods)
        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("ExecuteSQL Query Error", sFuncName)
            Throw New Exception(ex.Message)
        Finally
            Con.Dispose()
        End Try
        Return ods
    End Function
#End Region

#Region "Get Country Code for cheque payment"
    Public Function GetCountryCode(ByRef sDBName As String) As String
        Dim sSql As String
        Dim ods As DataSet
        Dim sFuncName As String = "GetCountryCode"
        Dim sCountCode As String = String.Empty

        sSql = "SELECT ""U_CheckBankCountry"" FROM ""@AI_TB01_COMPANYDATA"" WHERE ""Code"" = '" & sDBName & "'"
        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL", sFuncName)

        ods = ExecuteSQLQuery(sSql)

        If ods.Tables(0).Rows.Count > 0 Then
            sCountCode = ods.Tables(0).Rows(0).Item("U_CheckBankCountry").ToString
        End If

        Return sCountCode
    End Function
#End Region

#Region "Get Bank Code"
    Public Function GetBankCode(ByRef sDBName As String) As String
        Dim sFuncName As String = "GetBankCode"
        Dim sSql As String
        Dim oDs As DataSet
        Dim sBnkCode As String = String.Empty

        sSql = "SELECT ""U_CheckBankCode"" FROM ""@AI_TB01_COMPANYDATA"" WHERE ""Code"" = '" & sDBName & "'"
        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL", sFuncName)

        oDs = ExecuteSQLQuery(sSql)

        If oDs.Tables(0).Rows.Count > 0 Then
            sBnkCode = oDs.Tables(0).Rows(0).Item("U_CheckBankCode").ToString
        End If

        Return sBnkCode
    End Function
#End Region

#Region "Get Cheque GL Account"
    Public Function GetCheckGLAcct(ByRef sDBName As String) As String
        Dim sFuncName As String = "GetCheckGLAcct"
        Dim sSql As String
        Dim ods As DataSet
        Dim sGLAcct As String = String.Empty

        sSql = "SELECT ""U_CheckGLAccount"" FROM ""@AI_TB01_COMPANYDATA"" WHERE ""Code"" = '" & sDBName & "'"
        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL", sFuncName)

        ods = ExecuteSQLQuery(sSql)

        If ods.Tables(0).Rows.Count > 0 Then
            sGLAcct = ods.Tables(0).Rows(0).Item("U_CheckGLAccount").ToString
        End If

        Return sGLAcct
    End Function
#End Region

#Region "Get Bank Account for Cheque"
    Public Function GetBankAcct(ByRef sDBName As String) As String
        Dim sFuncName As String = "GetBankAcct"
        Dim sSql As String
        Dim ods As DataSet
        Dim sBankAcct As String = String.Empty

        sSql = "SELECT ""U_CheckBankAccount"" FROM ""@AI_TB01_COMPANYDATA"" WHERE ""Code"" = '" & sDBName & "'"
        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL", sFuncName)

        ods = ExecuteSQLQuery(sSql)

        If ods.Tables(0).Rows.Count > 0 Then
            sBankAcct = ods.Tables(0).Rows(0).Item("U_CheckBankAccount").ToString
        End If

        Return sBankAcct
    End Function
#End Region

#Region "Get House Bank Account Based on CardCode"
    Public Function GetHouseBankAccout(ByVal sCardCode As String, ByVal sCompanyDB As String) As String
        Dim sFuncName As String = "GetHouseBankAccout"
        Dim sSql As String
        Dim ods As DataSet
        Dim sHouseBnkAcct As String = String.Empty

        sSql = "SELECT ""HousBnkAct"",""CardCode"" FROM ""OCRD"" WHERE ""CardCode"" =  '" & sCardCode & "' AND IFNULL(""HousBnkAct"",'') <> ''"
        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL" & sSql, sFuncName)

        ods = ExecuteTargetCompSQLQuery(sSql, sCompanyDB)

        If ods.Tables(0).Rows.Count > 0 Then
            sHouseBnkAcct = ods.Tables(0).Rows(0).Item("HousBnkAct").ToString
        End If

        Return sHouseBnkAcct
    End Function

    Public Function GetHouseBankAccout_GIRO(ByVal sCardCode As String, ByVal sCompanyDB As String) As String
        Dim sFuncName As String = "GetHouseBankAccout"
        Dim sSql As String
        Dim ods As DataSet
        Dim sHouseBnkGLAcct As String = String.Empty

        'sSql = "SELECT ""HousBnkAct"",""CardCode"" FROM ""OCRD"" WHERE ""CardCode"" =  '" & sCardCode & "' AND IFNULL(""HousBnkAct"",'') <> ''"
        sSql = "SELECT B.""GLAccount"" FROM ""OCRD"" A INNER JOIN ""DSC1"" B ON B.""BankCode"" = A.""HouseBank"" AND B.""Country"" = A.""HousBnkCry"" AND B.""Branch"" = A.""HousBnkBrn"" " & _
               " WHERE A.""CardCode"" =  '" & sCardCode & "' AND IFNULL(A.""HousBnkAct"",'') <> '' AND B.""Account"" = A.""HousBnkAct"" "
        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL" & sSql, sFuncName)

        ods = ExecuteTargetCompSQLQuery(sSql, sCompanyDB)

        If ods.Tables(0).Rows.Count > 0 Then
            sHouseBnkGLAcct = ods.Tables(0).Rows(0).Item("GLAccount").ToString
        End If

        Return sHouseBnkGLAcct
    End Function
#End Region

#Region "Get Bank Transfer Account"
    Public Function GetBankTrnsAcct(ByRef sDBName As String) As String
        Dim sFuncName As String = "GetBankTrnsAcct"
        Dim sSql As String = String.Empty
        Dim ods As DataSet
        Dim sTrnsAcct As String = String.Empty

        sSql = "SELECT ""U_GIROGLAccount"" FROM " & p_oCompDef.sSAPDBName & ".""@AI_TB01_COMPANYDATA"" WHERE ""Code"" = '" & sDBName & "'"
        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Query : " & sSql, sFuncName)

        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL", sFuncName)

        ods = ExecuteSQLQuery(sSql)

        If ods.Tables(0).Rows.Count > 0 Then
            sTrnsAcct = ods.Tables(0).Rows(0).Item("U_GIROGLAccount").ToString
        End If

        Return sTrnsAcct
    End Function
#End Region

    Public Sub FileMoveToArchive(ByVal oFile As System.IO.FileInfo, ByVal CurrFileToUpload As String, ByVal iStatus As Integer)

        'Event      :   FileMoveToArchive
        'Purpose    :   For Renaming the file with current time stamp & moving to archive folder
        'Author     :   SRI 
        'Date       :   24 NOV 2013

        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "FileMoveToArchive"

            'Dim RenameCurrFileToUpload = Replace(CurrFileToUpload.ToUpper, ".CSV", "") & "_" & Format(Now, "yyyyMMddHHmmss") & ".csv"
            Dim RenameCurrFileToUpload As String = Mid(oFile.Name, 1, oFile.Name.Length - 4) & "_" & Now.ToString("yyyyMMddhhmmss") & ".xls"

            If iStatus = RTN_SUCCESS Then
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Moving Excel file to success folder", sFuncName)
                oFile.MoveTo(p_oCompDef.sSuccessDir & "\" & RenameCurrFileToUpload)
            Else
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Moving Excel file to Fail folder", sFuncName)
                oFile.MoveTo(p_oCompDef.sFailDir & "\" & RenameCurrFileToUpload)
            End If
        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error in renaming/copying/moving", sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try
    End Sub

    Public Sub FileMoveToArchive_Success(ByVal oFile As System.IO.FileInfo, ByVal CurrFileToUpload As String, ByVal sFileName As String, ByVal iStatus As Integer)

        'Event      :   FileMoveToArchive
        'Purpose    :   For Renaming the file with current time stamp & moving to archive folder
        'Author     :   SRI 
        'Date       :   24 NOV 2013

        Dim sFuncName As String = String.Empty
        Dim sFolderName As String = String.Empty
        Dim sFilePath As String = String.Empty

        Try
            sFuncName = "FileMoveToArchive"

            'Dim RenameCurrFileToUpload = Replace(CurrFileToUpload.ToUpper, ".CSV", "") & "_" & Format(Now, "yyyyMMddHHmmss") & ".csv"
            Dim RenameCurrFileToUpload As String = Mid(oFile.Name, 1, oFile.Name.Length - 4) & "_" & Now.ToString("yyyyMMddhhmmss") & ".xls"

            If iStatus = RTN_SUCCESS Then
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Moving Excel file to success folder", sFuncName)

                sFolderName = sFileName.Substring(0, 3)
                sFilePath = p_oCompDef.sSuccessDir & "\" & sFolderName

                If Not Directory.Exists(sFilePath) Then
                    Directory.CreateDirectory(sFilePath)
                End If

                oFile.MoveTo(sFilePath & "\" & RenameCurrFileToUpload)
            Else
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Moving Excel file to Fail folder", sFuncName)
                oFile.MoveTo(p_oCompDef.sFailDir & "\" & RenameCurrFileToUpload)
            End If
        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error in renaming/copying/moving", sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try
    End Sub

    Public Function IsXLBookOpen(strName As String) As Boolean

        'Function designed to test if a specific Excel
        'workbook is open or not.
        Dim i As Long
        Dim XLAppFx As Excel.Application
        Dim NotOpen As Boolean
        Dim sFuncName As String = "IsXLBookOpen"


        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

        'Find/create an Excel instance
        On Error Resume Next
        XLAppFx = GetObject(, "Excel.Application")
        If Err.Number = 429 Then
            NotOpen = True
            XLAppFx = CreateObject("Excel.Application")
            Err.Clear()
        End If

        'Loop through all open workbooks in such instance

        For i = XLAppFx.Workbooks.Count To 1 Step -1

            If XLAppFx.Workbooks(i).Name = strName Then
                'Perform check to see if name was found
                IsXLBookOpen = True
                Exit For
            Else
                'Set all to False
                IsXLBookOpen = False
            End If
        Next i

        'Close if was closed
        If NotOpen Then XLAppFx.Quit()

        'Release the instance
        XLAppFx = Nothing
        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

    End Function


End Module