Module modProcess

#Region "Start"
    Public Sub Start()
        Dim sFuncName As String = "Start()"
        Dim sErrDesc As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("calling ReadExcel()", sFuncName)

            Console.WriteLine("Reading Excel values")

            UploadFiles(sErrDesc)

            'Send Error Email if Datable has rows.
            If p_oDtError.Rows.Count > 0 Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling EmailTemplate_Error()", sFuncName)
                EmailTemplate_Error()
            End If
            p_oDtError.Rows.Clear()

            'Send Success Email if Datable has rows..
            If p_oDtSuccess.Rows.Count > 0 Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling EmailTemplate_Success()", sFuncName)
                EmailTemplate_Success()
            End If
            p_oDtSuccess.Rows.Clear()


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            End
        End Try

    End Sub
#End Region

#Region "Read Excel Files"

    Private Function UploadFiles(ByRef sErrDesc As String) As Long

        Dim sFuncName As String = "UploadFiles"
        Dim bIsFileExists As Boolean = False
        Dim oDVData As DataView = New DataView

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Upload funciton", sFuncName)

            p_oDtSuccess = CreateDataTable("FileName", "Status")
            p_oDtError = CreateDataTable("FileName", "Status", "ErrDesc")
            p_oDtReport = CreateDataTable("Type", "DocEntry", "BPCode", "Owner")

            Dim DirInfo As New System.IO.DirectoryInfo(p_oCompDef.sInboxDir)
            Dim files() As System.IO.FileInfo

            files = DirInfo.GetFiles("*.xls")

            For Each file As System.IO.FileInfo In files
                sErrDesc = String.Empty
                bIsFileExists = True

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("File Name is: " & file.Name.ToUpper, sFuncName)
                Console.WriteLine("Reading File: " & file.Name.ToUpper)

                Dim sFileType As String = String.Empty
                Dim sFileName As String = String.Empty
                sFileName = file.FullName

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling IsXLBookOpen()", sFuncName)

                If IsXLBookOpen(file.Name) = True Then
                    sErrDesc = "File is in use. Please close the document. File Name : " & sFileName
                    Console.WriteLine(sErrDesc)
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug(sErrDesc, sFuncName)

                    'Insert Error Description into Table
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
                    AddDataToTable(p_oDtError, file.Name, "Error", sErrDesc)

                    Continue For
                End If

                Dim k As Integer = sFileName.IndexOf("_")
                sFileType = sFileName.Substring(k, Len(sFileName) - k)

                If sFileType.Contains("_AR_") Then

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Read Excel file into Dataview", sFuncName)
                    oDVData = GetDataViewFromExcel(file.FullName, file.Extension)

                    If Not oDVData Is Nothing Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteBilltoClient()", sFuncName)
                        Console.WriteLine("Processing Bill to Client file " & sFileName)
                        If ProcessBilltoClient(file, oDVData, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Data's found in excel. File Name :" & file.Name, sFuncName)
                        Continue For
                    End If

                ElseIf sFileType.Contains("_PM_") Then

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Read Excel file into Dataview", sFuncName)
                    oDVData = GetDataViewFromExcel(file.FullName, file.Extension)

                    If Not oDVData Is Nothing Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteBilltoClient()", sFuncName)
                        If ProcessReimbToMember(file, oDVData, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Data's found in excel. File Name :" & file.Name, sFuncName)
                        Continue For
                    End If

                ElseIf sFileType.Contains("_AP_") Then


                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Read AP Invoice to Dataview", sFuncName)
                    oDVData = GetDataViewFromExcel(sFileName, file.Extension)

                    If Not oDVData Is Nothing Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteBilltoClient()", sFuncName)
                        If ProcessAPInvToProvider(file, oDVData, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Data's found in excel. File Name :" & file.Name, sFuncName)
                        Continue For
                    End If
                ElseIf sFileType.Contains("_PP_") Then

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Read Payment to Provider to Dataview", sFuncName)
                    oDVData = GetDataViewFromExcel(sFileName, file.Extension)

                    If Not oDVData Is Nothing Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Dataview return null", sFuncName)
                        If ProcessPayToProvider(file, oDVData, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Data's found in excel. file name :" & file.Name, sFuncName)
                        Continue For
                    End If
                End If
            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function Completed Successfully.", sFuncName)
            UploadFiles = RTN_SUCCESS

        Catch ex As Exception
            UploadFiles = RTN_ERROR
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error in Uplodiang AR file.", sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try
    End Function
#End Region

End Module
