

Module ModStockTakeApproval
    Sub Main()

        Dim sFuncName As String = String.Empty
        Dim sErrDesc As String = String.Empty
        Dim oStockTakeApp As AE_HarrysStockTake.DLL.clsStockTakeApproval = New AE_HarrysStockTake.DLL.clsStockTakeApproval()

        Dim oDataSetSO As DataSet = Nothing
        Dim oDSDataHeader As DataSet = Nothing
        Dim oDSDODrafts As DataSet = Nothing
        Dim sQueryString As String = String.Empty


        Try

            sFuncName = "Main()"
            Console.WriteLine("Starting Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            'Getting the Parameter Values from App Cofig File
            Console.WriteLine("Calling GetSystemIntializeInfo() ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling GetSystemIntializeInfo()", sFuncName)
            If GetSystemIntializeInfo(p_oCompDef, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            P_sConString = String.Empty
            p_sSAPConnString = String.Empty

            P_sConString = "Data Source=" & p_oCompDef.p_sServerName & ";Initial Catalog=" & p_oCompDef.p_sIntDBName & ";User ID=" & p_oCompDef.p_sDBUserName & "; Password=" & p_oCompDef.p_sDBPassword

            AE_HarrysStockTake.DLL.P_sConString = P_sConString

            p_sSAPConnString = "Data Source=" & p_oCompDef.p_sServerName & ";Initial Catalog=" & p_oCompDef.p_sDataBaseName & ";User ID=" & p_oCompDef.p_sDBUserName & "; Password=" & p_oCompDef.p_sDBPassword


            'For Stock Take Approval Process:
            '================================================   Starting the Function   ==============================================

            sQueryString = "SELECT [DocEntry],[WhsCode],[DBName] FROM [StockTakeApproval] " & _
                 " where isnull([ReceiveDate],'')='' and isnull([ErrMsg],'')='' and ISNULL([Status] ,'')=''"

            Console.WriteLine("Calling Get_DataSet() from Integration DB ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Get_DataSet()", sFuncName)

            oDSDataHeader = oStockTakeApp.Get_DataSet(sQueryString, P_sConString, sErrDesc)

            If Not oDSDataHeader Is Nothing Then

                Console.WriteLine("Calling StockTakeApproval() ", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectToTargetCompany()", sFuncName)
                If oStockTakeApp.StockTakeApproval(oDSDataHeader, p_oCompany, P_sConString, p_sSAPConnString, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            End If

            '================================================= End the Function ================================================================


            'For Converting the DO Drafts to Actual for EASI: 

            '================================================   Starting the Function   ========================================================

            ''Fetching the Values from the  SAP

            'sQueryString = "select T0.DocEntry,T0.DocDate   from ODRF T0 WITH (NOLOCK)  " & _
            '        " INNER JOIN DRF1 T1 WITH (NOLOCK) ON T0.DocEntry =T1.DocEntry " & _
            '        " where CardCode ='" & P_sCardCode & "' AND DocStatus ='O' AND T1.BaseType ='13' " & _
            '        " GROUP BY T0.DocEntry ,T0.DocDate ORDER BY T0.DocDate ASC "


            'Console.WriteLine("Calling Get_DataSet() from Integration DB ", sFuncName)
            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Get_DataSet()", sFuncName)

            ''Getting the Query Result in DataSet
            'oDSDODrafts = oStockTakeApp.Get_DataSet(sQueryString, p_sSAPConnString, sErrDesc)


            'If Not oDSDODrafts Is Nothing Then

            '    'Function to connect the Company
            '    If p_oCompany Is Nothing Then
            '        Console.WriteLine("Calling ConnectToTargetCompany() ", sFuncName)
            '        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectToTargetCompany()", sFuncName)
            '        If ConnectToTargetCompany(p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            '    End If

            '    'If Company Connected then Call the Function to Convert Actual Document from Drafts
            '    If Not p_oCompany Is Nothing Then
            '        Console.WriteLine("Calling StockTakeApproval() ", sFuncName)
            '        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectToTargetCompany()", sFuncName)
            '        If oStockTakeApp.ConvertDraftToDocument(oDSDODrafts, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            '    End If


            'End If

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
End Module
