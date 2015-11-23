Option Explicit On

Imports System.IO
Imports System.Configuration

Module modLog

    Public Structure CompanyDefault

        Public p_sServerName As String
        Public p_sLicServerName As String
        Public p_sDataBaseName As String
        Public p_sIntDBName As String
        Public p_sDBUserName As String
        Public p_sDBPassword As String
        Public p_sSAPUserName As String
        Public p_sSAPPassword As String
        Public p_sSQLType As Int16
        Public p_sLogDir As String
        Public p_sDebug As String
        Public p_sBINLocation As String

    End Structure


    Public Const RTN_SUCCESS As Int16 = 1
    Public Const RTN_ERROR As Int16 = 0
    ' Debug Value Variable Control
    Public Const DEBUG_ON As Int16 = 1
    Public Const DEBUG_OFF As Int16 = 0

    ' Global variables group
    Public p_iDebugMode As Int16
    Public p_iErrDispMethod As Int16
    Public p_iDeleteDebugLog As Int16
    Public p_oCompDef As CompanyDefault

    ' Public p_oCompany As SAPbobsCOM.Company
    Public p_oDataView As DataView = Nothing
    Public P_sConString As String = String.Empty
    Public p_sSAPConnString As String = String.Empty
    Public p_oCompany As SAPbobsCOM.Company
    Public P_sCardCode As String = String.Empty

    '***************************************
    'Name       :   modLog
    'Descrption :   Contains function for log errors and Application related information
    'Author     :   JOHN
    'Created    :   MAY 2014
    '***************************************

    Private Const MAXFILESIZE_IN_MB As Int16 = 5 '(2 MB)
    Private Const LOG_FILE_ERROR As String = "ErrorLog"
    Private Const LOG_FILE_ERROR_ARCH As String = "ErrorLog_"
    Private Const LOG_FILE_DEBUG As String = "DebugLog"
    Private Const LOG_FILE_DEBUG_ARCH As String = "DebugLog_"
    Private Const FILE_SIZE_CHECK_ENABLE As Int16 = 1
    Private Const FILE_SIZE_CHECK_DISABLE As Int16 = 0

    Public Function WriteToLogFile(ByVal strErrText As String, ByVal strSourceName As String, Optional ByVal intCheckFileForDelete As Int16 = 1) As Long

        ' **********************************************************************************
        '   Function   :    WriteToLogFile()
        '   Purpose    :    This function checks if given input file name exists or not
        '
        '   Parameters :    ByVal strErrText As String
        '                       strErrText = Text to be written to the log
        '                   ByVal intLogType As Integer
        '                       intLogType = Log type (1 - Log ; 2 - Error ; 0 - None)
        '                   ByVal strSourceName As String
        '                       strSourceName = Function name calling this function
        '                   Optional ByVal intCheckFileForDelete As Integer
        '                       intCheckFileForDelete = Flag to indicate if file size need to be checked before logging (0 - No check ; 1 - Check)
        '   Return     :    0 - FAILURE
        '                   1 - SUCCESS
        '   Author     :    JOHN
        '   Date       :    MAY 2014
        ' **********************************************************************************

        Dim oStreamWriter As StreamWriter = Nothing
        Dim strFileName As String = String.Empty
        Dim strArchFileName As String = String.Empty
        Dim strTempString As String = String.Empty
        Dim lngFileSizeInMB As Double

        Try
            strTempString = Space(IIf(Len(strSourceName) > 30, 0, 30 - Len(strSourceName)))
            strSourceName = strTempString & strSourceName
            strErrText = "[" & Format(Now, "MM/dd/yyyy HH:mm:ss") & "]" & "[" & strSourceName & "] " & strErrText


            strFileName = p_oCompDef.p_sLogDir & "\" & LOG_FILE_ERROR & ".log"
            strArchFileName = p_oCompDef.p_sLogDir & "\" & LOG_FILE_ERROR_ARCH & Format(Now(), "yyMMddHHMMss") & ".log"


            'strFileName = System.IO.Directory.GetCurrentDirectory() & "\" & LOG_FILE_ERROR & ".log"
            'strArchFileName = System.IO.Directory.GetCurrentDirectory() & "\" & LOG_FILE_ERROR_ARCH & Format(Now(), "yyMMddHHMMss") & ".log"

            If intCheckFileForDelete = FILE_SIZE_CHECK_ENABLE Then
                If File.Exists(strFileName) Then
                    lngFileSizeInMB = (FileLen(strFileName) / 1024) / 1024
                    If lngFileSizeInMB >= MAXFILESIZE_IN_MB Then
                        'If intCheckDeleteDebugLog=1 then remove all debug_log file
                        If p_iDeleteDebugLog = 1 Then
                            For Each sFileName As String In Directory.GetFiles(System.IO.Directory.GetCurrentDirectory(), LOG_FILE_DEBUG_ARCH & "*")
                                File.Delete(sFileName)
                            Next
                        End If
                        File.Move(strFileName, strArchFileName)
                    End If
                End If
            End If
            oStreamWriter = File.AppendText(strFileName)
            oStreamWriter.WriteLine(strErrText)
            WriteToLogFile = RTN_SUCCESS
        Catch exc As Exception
            WriteToLogFile = RTN_ERROR
        Finally
            If Not IsNothing(oStreamWriter) Then
                oStreamWriter.Flush()
                oStreamWriter.Close()
                oStreamWriter = Nothing
            End If
        End Try

    End Function

    Public Function WriteToLogFile_Debug(ByVal strErrText As String, ByVal strSourceName As String, Optional ByVal intCheckFileForDelete As Int16 = 1) As Long
        ' **********************************************************************************
        '   Function   :    WriteToLogFile_Debug()
        '   Purpose    :    This function checks if given input file name exists or not
        '
        '   Parameters :    ByVal strErrText As String
        '                       strErrText = Text to be written to the log
        '                   ByVal intLogType As Integer
        '                       intLogType = Log type (1 - Log ; 2 - Error ; 0 - None)
        '                   ByVal strSourceName As String
        '                       strSourceName = Function name calling this function
        '                   Optional ByVal intCheckFileForDelete As Integer
        '                       intCheckFileForDelete = Flag to indicate if file size need to be checked before logging (0 - No check ; 1 - Check)
        '   Return     :    0 - FAILURE
        '                   1 - SUCCESS
        '   Author     :    JOHN
        '   Date       :    MAY 2013
        '   Changes    : 
        '                   
        ' **********************************************************************************

        Dim oStreamWriter As StreamWriter = Nothing
        Dim strFileName As String = String.Empty
        Dim strArchFileName As String = String.Empty
        Dim strTempString As String = String.Empty
        Dim lngFileSizeInMB As Double
        Dim iFileCount As Integer = 0

        Try
            strTempString = Space(IIf(Len(strSourceName) > 30, 0, 30 - Len(strSourceName)))
            strSourceName = strTempString & strSourceName
            strErrText = "[" & Format(Now, "MM/dd/yyyy HH:mm:ss") & "]" & "[" & strSourceName & "] " & strErrText

            strFileName = p_oCompDef.p_sLogDir & "\" & LOG_FILE_DEBUG & ".log"
            strArchFileName = p_oCompDef.p_sLogDir & "\" & LOG_FILE_DEBUG_ARCH & Format(Now(), "yyMMddHHMMss") & ".log"


            'strFileName = System.IO.Directory.GetCurrentDirectory() & "\" & LOG_FILE_DEBUG & ".log"
            'strArchFileName = System.IO.Directory.GetCurrentDirectory() & "\" & LOG_FILE_DEBUG_ARCH & Format(Now(), "yyMMddHHMMss") & ".log"

            If intCheckFileForDelete = FILE_SIZE_CHECK_ENABLE Then
                If File.Exists(strFileName) Then
                    lngFileSizeInMB = (FileLen(strFileName) / 1024) / 1024
                    If lngFileSizeInMB >= MAXFILESIZE_IN_MB Then
                        'If intCheckDeleteDebugLog=1 then remove all debug_log file
                        If p_iDeleteDebugLog = 1 Then
                            For Each sFileName As String In Directory.GetFiles(System.IO.Directory.GetCurrentDirectory(), LOG_FILE_DEBUG_ARCH & "*")
                                File.Delete(sFileName)
                            Next
                        End If
                        File.Move(strFileName, strArchFileName)
                    End If
                End If
            End If
            oStreamWriter = File.AppendText(strFileName)
            oStreamWriter.WriteLine(strErrText)
            WriteToLogFile_Debug = RTN_SUCCESS
        Catch exc As Exception
            WriteToLogFile_Debug = RTN_ERROR
        Finally
            If Not IsNothing(oStreamWriter) Then
                oStreamWriter.Flush()
                oStreamWriter.Close()
                oStreamWriter = Nothing
            End If
        End Try

    End Function

    Public Function GetSystemIntializeInfo(ByRef oCompDef As CompanyDefault, ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   GetSystemIntializeInfo()
        '   Purpose     :   This function will be providing information about the initialing variables
        '               
        '   Parameters  :   ByRef oCompDef As CompanyDefault
        '                       oCompDef =  set the Company Default structure
        '                   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author      :   JOHN
        '   Date        :   MAY 2014
        ' **********************************************************************************

        Dim sFuncName As String = String.Empty
        Dim sConnection As String = String.Empty
        Dim sSqlstr As String = String.Empty
        Try

            sFuncName = "GetSystemIntializeInfo()"
            Console.WriteLine("Starting Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)


            oCompDef.p_sServerName = String.Empty
            oCompDef.p_sLicServerName = String.Empty
            oCompDef.p_sDBUserName = String.Empty
            oCompDef.p_sDBPassword = String.Empty

            oCompDef.p_sDataBaseName = String.Empty
            oCompDef.p_sSAPUserName = String.Empty
            oCompDef.p_sSAPPassword = String.Empty

            oCompDef.p_sLogDir = String.Empty
            oCompDef.p_sDebug = String.Empty
            oCompDef.p_sIntDBName = String.Empty
            oCompDef.p_sBINLocation = String.Empty



            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("Server")) Then
                oCompDef.p_sServerName = ConfigurationManager.AppSettings("Server")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("LicenseServer")) Then
                oCompDef.p_sLicServerName = ConfigurationManager.AppSettings("LicenseServer")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SAPDBName")) Then
                oCompDef.p_sDataBaseName = ConfigurationManager.AppSettings("SAPDBName")
                AE_HarrysStockTake.DLL.P_sSAPDBName = oCompDef.p_sDataBaseName
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SAPUserName")) Then
                oCompDef.p_sSAPUserName = ConfigurationManager.AppSettings("SAPUserName")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SAPPassword")) Then
                oCompDef.p_sSAPPassword = ConfigurationManager.AppSettings("SAPPassword")
            End If


            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBUser")) Then
                oCompDef.p_sDBUserName = ConfigurationManager.AppSettings("DBUser")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBPwd")) Then
                oCompDef.p_sDBPassword = ConfigurationManager.AppSettings("DBPwd")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SQLType")) Then
                oCompDef.p_sSQLType = ConfigurationManager.AppSettings("SQLType")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("IntegrationDBName")) Then
                oCompDef.p_sIntDBName = ConfigurationManager.AppSettings("IntegrationDBName")
                AE_HarrysStockTake.DLL.P_sStagingDBName = oCompDef.p_sIntDBName
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("BinLocation")) Then
                oCompDef.p_sBINLocation = ConfigurationManager.AppSettings("BinLocation")
                AE_HarrysStockTake.DLL.sBINLocation = oCompDef.p_sBINLocation
            End If


            ' folder
            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("LogDir")) Then
                oCompDef.p_sLogDir = ConfigurationManager.AppSettings("LogDir")
                AE_HarrysStockTake.DLL.sLogFolderPath = oCompDef.p_sLogDir
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("Debug")) Then
                oCompDef.p_sDebug = ConfigurationManager.AppSettings("Debug")
                If p_oCompDef.p_sDebug.ToUpper = "ON" Then
                    p_iDebugMode = 1
                Else
                    p_iDebugMode = 0
                End If
            Else
                p_iDebugMode = 0
            End If

            AE_HarrysStockTake.DLL.p_iDebugMode = p_iDebugMode

            'IntegrationDBName

            Console.WriteLine("Completed with SUCCESS ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            GetSystemIntializeInfo = RTN_SUCCESS

        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            Console.WriteLine("Completed with ERROR ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            GetSystemIntializeInfo = RTN_ERROR
        End Try
    End Function

    Public Function ConnectToTargetCompany(ByRef oCompany As SAPbobsCOM.Company, _
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
        '   Author      :   JOHN
        '   Date        :   MAY 2013 21
        ' **********************************************************************************

        Dim sFuncName As String = String.Empty
        Dim iRetValue As Integer = -1
        Dim iErrCode As Integer = -1
        Dim sSQL As String = String.Empty
        Dim oDs As New DataSet

        Try
            sFuncName = "ConnectToTargetCompany()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing the Company Object", sFuncName)

            oCompany = New SAPbobsCOM.Company

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assigning the representing database name", sFuncName)

            oCompany.Server = p_oCompDef.p_sServerName
            oCompany.LicenseServer = p_oCompDef.p_sLicServerName
            oCompany.DbUserName = p_oCompDef.p_sDBUserName
            oCompany.DbPassword = p_oCompDef.p_sDBPassword
            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English
            oCompany.UseTrusted = False

            If p_oCompDef.p_sSQLType = 2012 Then
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012
            ElseIf p_oCompDef.p_sSQLType = 2008 Then
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008
            End If

            oCompany.CompanyDB = p_oCompDef.p_sDataBaseName
            oCompany.UserName = p_oCompDef.p_sSAPUserName
            oCompany.Password = p_oCompDef.p_sSAPPassword

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
            ConnectToTargetCompany = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ConnectToTargetCompany = RTN_ERROR
        End Try
    End Function

End Module
