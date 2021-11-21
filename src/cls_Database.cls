VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Database"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------------------------
'#Name : cls_Database
'
'#Description :
'   This module is the hear & soul of the application Backend. It's mixing ADO function & DAO function
'   - Database connection
'   - Query Execution
'   - Database Design
'   - ETL Creation
'
'#WARNING :
'
'#Contributors :
'   - Victor Hüni
'   - Jérôme Favre-Rochex
'
'#References :
'   - Microsoft ActiveX Data Objects 6.1 Library
'   - Microsoft Scripting Runtime
'   - Microsoft Office 16.0 Access Database Engine Object (DAO)
'   - Microsoft ADO Ext. 6.0 for DDL and Security
'----------------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit

'====================================================================================================================================================
'   MEMBERS
'===================================================================================================================================================

'Custom Error Type
Private Const ERROR_UNSUPPORTED_TYPE_PARAM As Long = vbObjectError + 513

'Logs
Private logs As cls_Logs

'Data Sources Types
Private m_DataSourceTypesDisplay As Variant
Private m_DataSourceTypes As Variant
Private m_VBADataSourceTypes As Variant

'Folders
Private Const SEP_DB As String = "-"
Private Const SEP_FOLDER As String = "\"

Private Const DATA_FOLDER As String = "\Data"


'Root Build Path - WARNING : Require the wsAPropos Sheet Correctly configured
Private m_Root As String
Private m_Site As String
Private m_AppName As String
Private m_Env As String
Private m_Ext As String
Private m_DbName As String
Private m_DbFilePath As String

'Connection members
Private m_CnStr As String
Private m_oCn As ADODB.Connection
Private m_rsSource As ADODB.Recordset
Private m_rsTarget As ADODB.Recordset

Private m_SourceQuery As String
Private m_UpdateQuery As String
Private m_InsertQuery As String

'Output parameters
Private m_IsSelectRecordSetEmpty As Boolean
Private m_AffectedRows As Long
Private m_LastId As Long
Private m_Errors As String

'====================================================================================================================================================
'   PROPERTIES
'====================================================================================================================================================

'Root Build Path
Public Property Get Root() As String: Root = m_Root: End Property
Private Property Let Root(ByVal p_Root As String): m_Root = p_Root: End Property

Public Property Get Site() As String: Site = m_Site: End Property
Private Property Let Site(ByVal p_site As String): m_Site = p_site: End Property

Public Property Get AppName() As String: AppName = m_AppName: End Property
Private Property Let AppName(ByVal p_AppName As String): m_AppName = p_AppName: End Property

Public Property Get Env() As String: Env = m_Env: End Property
Private Property Let Env(ByVal p_Env As String): m_Env = p_Env: End Property

Public Property Get Ext() As String: Ext = m_Ext: End Property
Private Property Let Ext(ByVal p_Ext As String): m_Ext = p_Ext: End Property

Public Property Get DbName() As String: DbName = m_DbName: End Property
Private Property Let DbName(ByVal p_DbName As String): m_DbName = p_DbName: End Property

Public Property Get DbFilePath() As String: DbFilePath = m_DbFilePath: End Property
Private Property Let DbFilePath(ByVal p_DbFilePath As String): m_DbFilePath = p_DbFilePath: End Property


'Connection properties
Private Property Get CnStr() As String: CnStr = m_CnStr: End Property
Private Property Let CnStr(ByVal p_CnStr As String): m_CnStr = p_CnStr: End Property

Private Property Get oCn() As ADODB.Connection: Set oCn = m_oCn: End Property
Private Property Set oCn(ByVal p_oCn As ADODB.Connection): Set m_oCn = p_oCn: End Property

Public Property Get rsSource() As ADODB.Recordset: Set rsSource = m_rsSource: End Property
Private Property Set rsSource(ByVal p_rsSource As ADODB.Recordset): Set m_rsSource = p_rsSource: End Property

Public Property Get rsTarget() As ADODB.Recordset: Set rsTarget = m_rsTarget: End Property
Private Property Set rsTarget(ByVal p_rsTarget As ADODB.Recordset): Set m_rsTarget = p_rsTarget: End Property

'Queries for ETL
Public Property Get UpdateQuery() As String: UpdateQuery = m_UpdateQuery: End Property
Public Property Let UpdateQuery(ByVal p_UpdateQuery As String): m_UpdateQuery = p_UpdateQuery: End Property

Public Property Get InsertQuery() As String: InsertQuery = m_InsertQuery: End Property
Public Property Let InsertQuery(ByVal p_InsertQuery As String): m_InsertQuery = p_InsertQuery: End Property

Public Property Get SourceQuery() As String: SourceQuery = m_SourceQuery: End Property
Public Property Let SourceQuery(ByVal p_SourceQuery As String): m_SourceQuery = p_SourceQuery: End Property


'Output properties
Public Property Get IsSelectRecordSetEmpty() As Boolean: IsSelectRecordSetEmpty = m_IsSelectRecordSetEmpty: End Property
Private Property Let IsSelectRecordSetEmpty(ByVal p_IsSelectRecordSetEmpty As Boolean): m_IsSelectRecordSetEmpty = p_IsSelectRecordSetEmpty: End Property

Public Property Get AffectedRows() As Long: AffectedRows = m_AffectedRows: End Property
Private Property Let AffectedRows(ByVal p_AffectedRows As Long): m_AffectedRows = p_AffectedRows: End Property

Public Property Get LastId() As Long: LastId = m_LastId: End Property
Private Property Let LastId(ByVal p_LastId As Long): m_LastId = p_LastId: End Property

Public Property Get Errors() As Long: Errors = m_Errors: End Property
Private Property Let Errors(ByVal p_Errors As Long): m_Errors = p_Errors: End Property

'====================================================================================================================================================
'   EVENT HOOKS
'====================================================================================================================================================

Private Sub Class_Initialize()
    
    Dim i As Integer

    Set logs = New cls_Logs

    m_DataSourceTypes = Array()
    
    m_DataSourceTypesDisplay = Array("TXT", "CSV", "XLSX", "XLSM", "ACCDB")
    m_DataSourceTypes = Array(".txt", ".csv", ".xlsx", ".xlsm", ".accdb")
    m_VBADataSourceTypes = Array(xlTextWindows, xlCSVMSDOS, xlWorkbookDefault, xlOpenXMLWorkbookMacroEnabled) 'From XlFileFormat enumeration

    'Initialize Default database from the wsApropos file
    Root = wsAPropos.range("Root").Value
    Site = LCase(wsAPropos.range("Site").Value)
    AppName = LCase(wsAPropos.range("AppName").Value)
    Env = LCase(wsAPropos.range("Env").Value)
    Ext = wsAPropos.range("Ext").Value
    
    For i = 0 To UBound(m_DataSourceTypesDisplay)
        If Ext = m_DataSourceTypesDisplay(i) Then
            Ext = m_DataSourceTypes(i)
            DbName = Site & SEP_DB & AppName & SEP_DB & Env
            DbFilePath = Root & DATA_FOLDER
            Call InitCnStr(Me.GetDBFullPath, i)
            Exit For
        End If
    Next i

End Sub

Private Sub Class_Terminate()

    Set logs = Nothing
    Set oCn = Nothing
    Set rsSource = Nothing
    Set rsTarget = Nothing
    CnStr = ""
    m_DataSourceTypes = Array()
    
End Sub

'====================================================================================================================================================
'   CONNECTIONS METHODS
'====================================================================================================================================================

'----------------------------------------------------------------------------------------------------------------------------------------------------
'#Name:             GetDBFullPath
'
'#Description:      Get the full path of the DB with her extension
'
'#Contributors :    Victor Hüni
'#Creation Date :   30.04.2021
'
'#Dependencies :    Class_Initialize
'
'#Return :
'   [String] : Filepath to the DB
'
'#Example :
'
'#Source :
'----------------------------------------------------------------------------------------------------------------------------------------------------
Public Function GetDBFullPath() As String
    GetDBFullPath = DbFilePath & SEP_FOLDER & DbName & Ext
End Function


'----------------------------------------------------------------------------------------------------------------------------------------------------
'#Name:             InitCnStr
'
'#Description:      Create the correct connection string for the correct data source
'
'#WARNING:          p_type needs to have standards value stored in a enum that should be copy/paste in the mod_GlobalRef when using this classes
'                   p_path should contains the name of the folder only for a TXT, CSV source, and should contains path + name for an ACCDB & XLSX source
'
'#Contributors :    Jérôme Favre-Rochex, Victor Hüni
'#Creation Date :   01.10.2020
'#Version :         1.2.0
'
'#References :      Microsoft ActiveX Data Objects 6.1 Library, Microsoft Scripting Runtime
'#Dependencies :
'#Related :
'
'#Parameters :
'   p_path [String] : Path to the data source
'   p_type [String] : Type of the data source (ACCDB, XLSX, TXT...)
'
'
'#Return :
'   ADODB.DataTypeEnum : This is the new type that should be use for this parameters
'
'#Example :
'
'#Source :
'----------------------------------------------------------------------------------------------------------------------------------------------------
Public Function InitCnStr(ByVal p_path As String, ByVal p_Type As Integer) As Boolean
    
    Dim fso As New FileSystemObject
    
    'Check Types
    Select Case m_DataSourceTypes(p_Type)
    
        'TXT
        Case m_DataSourceTypes(e_DataSourceTypes.txt)
            If (fso.FolderExists(p_path)) Then
                CnStr = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & p_path & ";Extended Properties='text;HDR=YES;FMT=Delimited';"""
                InitCnStr = True
            Else
                InitCnStr = False
            End If
        
        Case m_DataSourceTypes(e_DataSourceTypes.CSV)
            If (fso.FolderExists(p_path)) Then
                CnStr = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & p_path & ";Extended Properties='text;HDR=YES;FMT=Delimited';"""
                InitCnStr = True
            Else
                InitCnStr = False
            End If
            
            'XLSM
        Case m_DataSourceTypes(e_DataSourceTypes.xlsm)
            If (fso.FileExists(p_path)) Then
                CnStr = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & p_path & ";Extended Properties=""Excel 12.0 Macro;HDR=YES;FMT=Delimited;"""
                InitCnStr = True
            Else
                InitCnStr = False
            End If
        
            'ACCDB
        Case m_DataSourceTypes(e_DataSourceTypes.ACCDB)
            If (fso.FileExists(p_path)) Then
                CnStr = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & p_path & ";Persist Security Info=True;"
                InitCnStr = True
            Else
                InitCnStr = False
            End If
    
    End Select
   
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------
'#Name :            OpenConnection
'
'#Description :     Open the connection with the data source
'
'#WARNING :         This methods needs that the CnStr member has been initialized
'
'#Contributors :    Jérôme Favre-Rochex, Victor Hüni
'#Creation Date :   01.10.2020
'#Version :         1.2.0
'
'#References :      Microsoft ActiveX Data Objects 6.1 Library
'#Dependencies :
'#Related :
'
'#Parameters :
'#Return :
'   Boolean : True = Connnection successful / False =  Connection Failed
'
'#Example :
'
'#Source :
'----------------------------------------------------------------------------------------------------------------------------------------------------
Private Function OpenConnection() As Boolean

    Dim initiateNew As Boolean

    On Error GoTo error_handler:
    
    initiateNew = True
    
    If Not oCn Is Nothing Then initiateNew = Not (SameCnStrUsed)
    
    If initiateNew Then
        Set oCn = New ADODB.Connection
        If Len(CnStr) > 0 Then                      'If connction string is set
            oCn.Open CnStr                          'Open connection
        End If
    End If

    OpenConnection = (oCn.state = adStateOpen)  'return true if connection initiate & false if not
   
    Exit Function

error_handler:
    logs.Log_Erreur ("OpenConnection")
    OpenConnection = False

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------
'#Name :            CloseConnection
'
'#Description :     Close the connection with the data source
'
'#WARNING :
'
'#Contributors :    Victor Hüni
'#Creation Date :   01.10.2020
'#Version :         1.2.0
'
'#References :      Microsoft ActiveX Data Objects 6.1 Library
'#Dependencies :    cls_Logs
'#Related :
'
'#Parameters :
'#Return :
'   Boolean : True = Connnection closing successful / False =  Connection closing Failed
'
'#Example :
'
'#Source :
'----------------------------------------------------------------------------------------------------------------------------------------------------
Private Function CloseConnection() As Boolean
          
    If oCn.state = adStateOpen Then     'If the connection is open
        oCn.Close                       'Close it
        Set oCn = Nothing               'Kill the connection object
    End If
    
    CloseConnection = True

End Function

Private Function getCnStrPath(ByVal p_CnStr As String)
    Dim arrCnStr As Variant
    Dim el As Variant
    arrCnStr = Split(oCn.ConnectionString, ";")
    For Each el In arrCnStr
        If InStr(1, el, "Data Source", vbTextCompare) > 0 Then
            getCnStrPath = Split(el, "=")(1)
            Exit For
        End If
    Next el
End Function

Private Function SameCnStrUsed() As Boolean
    SameCnStrUsed = (oCn.state = adStateOpen And getCnStrPath(CnStr) = getCnStrPath(oCn.ConnectionString))
End Function


'====================================================================================================================================================
'   QUERIES METHODS
'====================================================================================================================================================

Public Function TransformParamsInQuestionsMarks(ByRef arrParam As Variant) As String

    Dim i As Long
    Dim questionsMarks As String
    
    i = UBound(arrParam) - 1
    questionsMarks = Replace(space(i), " ", "?,")
    questionsMarks = questionsMarks & "?"
    
    TransformParamsInQuestionsMarks = questionsMarks
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------
'#Name :            GetParameterType
'
'#Description :     Format a parameters list with the right type cast.
'                   This list should be then use when preparing the SQL Statement via an ADODB.Command Object
'
'#Warning :         Generate an custom error that could be handle. This error is generate each time a vba type isn't adapt to ADODB
'
'#Contributors :    Victor Hüni
'#Creation Date :   01.10.2020
'#Version :         1.0.0
'
'#References :      Microsoft ActiveX Data Objects 6.1 Library
'#Dependencies :    cls_Logs
'#Related :
'
'#Parameters :
'   inputParam [Variant]   : Parameters that need to be casted before added into the comamand object
'
'#Return :
'   ADODB.DataTypeEnum : This is the new type that should be use for this parameters
'
'#Example :
'
'#Source :
'----------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetParameterType(inputParam As Variant) As ADODB.DataTypeEnum

    On Error GoTo UnsupportedType:
    
    Select Case VarType(inputParam)
        Case vbEmpty            '    0  Empty (uninitialized).
            Err.Raise ERROR_UNSUPPORTED_TYPE_PARAM, , "Non Blocking Error. Unsupported Type (vbEmpty) defined for a param in prepared query"
        Case vbNull             '    1  Null (no valid data).
            Err.Raise ERROR_UNSUPPORTED_TYPE_PARAM, , "Non Blocking Error. Unsupported Type (vbNull) defined for a param in prepared query"
        Case vbInteger          '    2  Integer.
            GetParameterType = adInteger
        Case vbLong             '    3  Long integer.
            GetParameterType = adInteger
        Case vbSingle           '    4  Single-precision floating-point number.
            GetParameterType = adSingle
        Case vbDouble           '    5  Double-precision floating-point number.
            GetParameterType = adDouble
        Case vbCurrency         '    6  Currency.
            GetParameterType = adDecimal
        Case vbDate             '    7  Date.
            GetParameterType = adDate
        Case vbString           '    8  String.
            GetParameterType = adVarWChar 'Assumes strings are not long
        Case vbObject           '    9  Object.
            Err.Raise ERROR_UNSUPPORTED_TYPE_PARAM, , "Non Blocking Error. Unsupported Type (vbObject) defined for a param in prepared query"
        Case vbError            '   10  Error.
            Err.Raise ERROR_UNSUPPORTED_TYPE_PARAM, , "Non Blocking Error. Unsupported Type (vbError) defined for a param in prepared query"
        Case vbBoolean          '   11  Boolean.
            GetParameterType = adBoolean
        Case vbVariant          '   12  Variant (used only with arrays of variants).
            Err.Raise ERROR_UNSUPPORTED_TYPE_PARAM, , "Non Blocking Error. Unsupported Type (vbVariant) defined for a param in prepared query"
        Case vbDataObject       '   13  A data access object.
            Err.Raise ERROR_UNSUPPORTED_TYPE_PARAM, , "Non Blocking Error. Unsupported Type (vbDataObject) defined for a param in prepared query"
        Case vbDecimal          '   14  Decimal.
            GetParameterType = adDecimal
        Case vbByte             '   17  Byte.
            GetParameterType = adChar
        Case vbUserDefinedType  '   36  Variants that contain user-defined types.
            Err.Raise ERROR_UNSUPPORTED_TYPE_PARAM, , "Non Blocking Error. Unsupported Type (vbUserDefinedType) defined for a param in prepared query"
        Case vbArray            '   8192  Array.
            Err.Raise ERROR_UNSUPPORTED_TYPE_PARAM, , "Non Blocking Error. Unsupported Type (vbArray) defined for a param in prepared query"
        Case Else               '   For example array + variant = 8204
            Err.Raise ERROR_UNSUPPORTED_TYPE_PARAM, , "Non Blocking Error. Unsupported Type (Mixed Type) defined for a param in prepared query"
    End Select
    
    Exit Function
    
UnsupportedType:
    'Logs.Log_Erreur ("GetParameterType")
    Err.Clear
    GetParameterType = adVarWChar
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------
'#Name :            ExecuteParameters
'
'#Description :     Prepare and execute an SQL Command, either a stored procedure OR a paramterized query
'
'#WARNING:
'
'#Contributors :    Victor Hüni, Jérôme Favre-Rochex
'#Creation Date :   16.12.2020
'#Version :         2.0.0
'
'#References :      Microsoft ActiveX Data Objects 6.1 Library
'                   Microsoft ADO Ext. 6.0 for DDL and Security
'
'#Dependencies :    cls_Logs
'                   GetParameterType
'                   OpenConnection
'                   CloseConnection
'#Related :
'
'#Parameters :
'   p_SqlCommand [String]       : This is the SQL Statement parametrable with ? to replace the value to be added
'   IsStoredProcedure [Boolean] : Allow the procedure to know if this only a stored procedure to be excuted and there for there no need to prepare the stametent
'   Params [Variant]            : The list of paramters to replace the ? in the SQL Statement. The ParamArray keywords allow a array without a fixed size as a parameter.
'
'#Return :
'   Boolean : True = Query execution successful / False =  Error during the query execution
'
'#Example :
'
'#Source :
'----------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ExecuteParameters(ByVal p_SqlCommand As String, ByVal IsStoredProcedure As Boolean, ByVal ConnHasToBeClosed As Boolean, ByVal rsSourceOrTarget As Boolean, ParamArray params() As Variant) As Boolean

    Dim objCmd As New ADODB.Command
    Dim inputParam As Variant
    Dim elem As Variant
    Dim lAffectedRows As Long
    Dim rsToUse As New ADODB.Recordset
    
    On Error GoTo error_handler:
    
    If rsSourceOrTarget Then
        Set rsSource = New ADODB.Recordset
    Else
        Set rsTarget = New ADODB.Recordset
    End If
    
    'Init connection --------------------------------------------------------------------------------------------------------------------------------
    If Not OpenConnection Then                   'Try to open the connection with the data source - Require a connection string set
        ExecuteParameters = False
        Exit Function
    End If

    Set objCmd.ActiveConnection = oCn         'Initiate the active connection of the obj comamnd
    
    objCmd.CommandText = p_SqlCommand
    
    'Check for command type -------------------------------------------------------------------------------------------------------------------------
    If IsStoredProcedure Then
    
        objCmd.CommandType = adCmdStoredProc    'If it a stored procedure, then I only set to parameters in the object command. Avoid the parametrization of the sql command
        
    Else
    
        objCmd.CommandType = adCmdText          'If it a simple query, I initiate the correct command type and lauch the parametrization
        
        'Add parameters -----------------------------------------------------------------------------------------------------------------------------
        For Each inputParam In params           'Loop through the potential ParamArray of Parameters
            If IsArray(inputParam) Then         'If the element in the ParamArray is an Array itself, we iterate opver its elements
                For Each elem In inputParam
                    elem = IIf(elem = "", Null, elem)
                    Set elem = objCmd.CreateParameter(, GetParameterType(elem), adParamInput, Len(IIf(IsNull(elem), " ", elem)), elem) 'Each element of this sub array is added as a parameters
                    objCmd.Parameters.Append elem                                                                                      'After cast with the correct type
                Next elem
            Else
                inputParam = IIf(inputParam = "", Null, inputParam)
                Set inputParam = objCmd.CreateParameter(, GetParameterType(inputParam), adParamInput, Len(IIf(IsNull(inputParam), " ", inputParam)), inputParam)    'Each element of this sub array is added as a parameters
                objCmd.Parameters.Append inputParam                                                                                                                 'After cast with the correct type
            End If
        Next inputParam
        
    End If
    
    'Execute Query/Stored Procedure------------------------------------------------------------------------------------------------------------------
    If left(p_SqlCommand, 6) = "SELECT" Or IsStoredProcedure Then
        
        'SELECT -------------------------------------------------------------------------------------------------------------------------------------

        
        With rsToUse
            .CursorLocation = adUseClient   'Required to then disconnect the Recordset
            .Open objCmd, CursorType:=adOpenDynamic, LockType:=adLockBatchOptimistic ' adOpenForwardOnly adOpenDynamic
            .ActiveConnection = Nothing
        End With
        
        IsSelectRecordSetEmpty = (rsToUse.EOF Or rsToUse.BOF)
        
        If rsSourceOrTarget Then
            Set rsSource = rsToUse
        Else
            Set rsTarget = rsToUse
        End If
        
        If G_DEBUG Then
            Dim arr As Variant
            Dim i As Long
            wsDebug.Cells.ClearContents
            For i = 1 To rsToUse.Fields.Count
                wsDebug.Cells(1, i) = rsToUse.Fields(i - 1).name
            Next i
            arr = TransposeArray(rsToUse.GetRows)
            range(wsDebug.Cells(2, 1), wsDebug.Cells(UBound(arr, 1) + 2, UBound(arr, 2) + 1)) = arr
        End If
        
    Else
    
        'INSERT/UDATE/DELETE ------------------------------------------------------------------------------------------------------------------------
        If rsSourceOrTarget Then
            Set rsSource = objCmd.Execute(lAffectedRows)
        Else
            Set rsTarget = objCmd.Execute(lAffectedRows)
        End If
        
        AffectedRows = lAffectedRows
        LastId = GetLastIdInserted
        
    End If
    
    ExecuteParameters = True
    
    Set objCmd = Nothing
    
    If ConnHasToBeClosed Then
        CloseConnection
    End If

    Exit Function

error_handler:
    CloseConnection
    Errors = Err.Number
    logs.Log_Erreur "ExecuteParameters", p_SqlCommand
    ExecuteParameters = False
    
End Function

Private Sub ExecuteParametersDAO(ByRef oDB As DAO.Database, ByVal strSQL As String, ByVal dicParamsValues As Dictionary)

    Dim qdf As DAO.QueryDef
    Dim i As Long
    Dim k As Variant
    
    Set qdf = oDB.CreateQueryDef(vbNullString, strSQL)
    
    With qdf
        For Each k In dicParamsValues.Keys
            'Debug.Print "param name : " & k & " --- value : " & dicParamsValues(k)
            .Parameters(k).Value = dicParamsValues(k)
        Next k
        .Execute dbFailOnError
    End With
        
    
End Sub
Private Sub ExecuteTransaction(ByVal arrSQL As Variant)
 
    Dim Acc As New DAO.DBEngine
    Dim wrkCurrent As DAO.Workspace
    Dim oDB As DAO.Database
    Dim rstArticles As DAO.Recordset
    Dim errLoop As Error
    Dim i As Long, k As Long
    
    
    'On Error GoTo ErrorHandler
    Set wrkCurrent = Acc.Workspaces(0)
    Set oDB = Acc.OpenDatabase(Me.GetDBFullPath, False, False)
    
    wrkCurrent.BeginTrans
    
    ' Run action query to restore data. Trap for errors,
    ' checking the Errors collection if necessary.
    'On Error GoTo Err_Execute
        For i = 0 To UBound(arrSQL)
            For k = 0 To UBound(arrSQL, 2) Step 2
                Call ExecuteParametersDAO(oDB, arrSQL(i, k), arrSQL(i, k + 1))
            Next k
        Next i
    'On Error GoTo 0
    wrkCurrent.CommitTrans

    oDB.Close
    wrkCurrent.Close
    
    Set oDB = Nothing
    Set wrkCurrent = Nothing
    
    Exit Sub
    
     
Err_Execute:
    ' wrkCurrent.Rollback
    
     ' Notify user of any errors that result from
     ' executing the query.
     If DBEngine.Errors.Count > 0 Then
     For Each errLoop In DBEngine.Errors
        MsgBox "Error number: " & errLoop.Number & vbCr & _
        errLoop.Description
     Next errLoop
     End If
     
     Resume Next
    
ErrorHandler:
    MsgBox "Error #: " & Err.Number & vbCrLf & vbCrLf & Err.Description
    
End Sub


Function LogsQuery(ByVal p_log_matricule As String, ByVal p_log_type As String, ByVal p_modified_table As String, ByVal p_modified_field As String, ByVal p_record_pk As String, p_new_value As String) As Boolean
    
    Dim res As Boolean
    
    res = False
    
    If ExecuteParameters("INSERT INTO f_logs (log_date, log_matricule, log_type, modified_table, modified_field, record_pk, new_value) VALUES (?, ?, ?, ?, ?, ?, ?)", False, False, True, CDate(Now), p_log_matricule, p_log_type, p_modified_table, p_modified_field, p_record_pk, p_new_value) Then
        If AffectedRows > 0 Then
            res = True
        End If
    End If
    
    LogsQuery = res
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------
'#Name :            GetLastIdInserted
'
'#Description :     Use the T-SQL parameters @@identity to get the last id inserted in the database with the same connection
'                   This functions is normally used right after an insert to avoid any errors and getting a wrong id.
'
'#WARNING :         This function could not use the ExecuteParameters Function due to it's specificity it has to open it's own recordset
'                   A select max
'
'#Contributors :    Victor Hüni
'#Creation Date :   01.10.2020
'#Version :         1.0.0
'
'#References :      Microsoft ActiveX Data Objects 6.1 Library
'#Dependencies :
'#Related :
'
'#Parameters :
'
'#Return :
'   Long :          The Primary key last inserted
'
'#Example :
'
'#Source :
'----------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetLastIdInserted() As Long

    Dim rs As New ADODB.Recordset
    
    rs.Open "SELECT @@IDENTITY AS lastID", oCn
    GetLastIdInserted = rs.Fields("lastID")
    rs.Close
    Set rs = Nothing

End Function

'====================================================================================================================================================
'   DATABASE DESIGN
'====================================================================================================================================================
'----------------------------------------------------------------------------------------------------------------------------------------------------
'#Name :            CompactAndRepairAccessDB
'
'#Description :     Execute the compact & repair function of Access on a ACCDB file to manage the size and avoid data corruption
'
'#WARNING :         This Function copy/paste the db and compact the copy AND then delete the old one
'
'#Contributors :    Victor Hüni, Jérôme Favre-Rochex
'#Creation Date :   01.10.2020
'#Version :         1.0.0
'
'#References :      Microsoft Office 16.0 Access Database Engine Object (DAO)
'#Dependencies :
'#Related :
'
'#Parameters :
'   p_DbFilePath [String]   : Path of the accdb files
'   p_DbName [String]       : Name of the accdb files
'
'#Return :
'
'#Example :
'
'#Source :
'----------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub CompactAndRepairAccessDB(ByVal p_DbFilePath As String, ByVal p_DbName As String)

    Dim Acc As New DAO.DBEngine
    Dim dbPath As String
    Dim dbPathX As String
    
    dbPath = p_DbFilePath & p_DbName
    dbPathX = p_DbFilePath & "tmp" & m_DataSourceTypes(e_DataSourceTypes.ACCDB)

    On Error Resume Next
    
    Acc.CompactDatabase dbPath, dbPathX
    Set Acc = Nothing
    Kill dbPath
    Name dbPathX As dbPath
    
    Call Err.Clear
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------
'#Name :            AddColumnToTable
'
'#Description :     ALlow the developper to add a column to a table
'
'#WARNING :         Type should respect the rules : https://docs.microsoft.com/en-us/sql/odbc/microsoft/microsoft-access-data-types?view=sql-server-ver15
'
'#Contributors :    Victor Hüni
'#Creation Date :   19.08.2020
'#Version :         1.0.0
'
'#References :
'#Dependencies :
'#Related :
'
'#Parameters :
'   p_Table [String]        : Table Name
'   p_ColumnName [String]   : Name of the new column
'   p_Type [String]         : Type of the new column
'
'#Return :
'   Boolean : True - Column added to the table / False - Issue when executing the query
'
'#Example :
'
'#Source :
'----------------------------------------------------------------------------------------------------------------------------------------------------
Public Function AddColumnToTable(ByVal p_Table As String, ByVal p_ColumnName As String, ByVal p_Type As String) As Boolean
    AddColumnToTable = ExecuteParameters("ALTER TABLE " & p_Table & " ADD COLUMN " & p_ColumnName & " " & p_Type, False, True, True)
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------
'#Name :            DeleteTableRelationships
'
'#Description :     Delete ALL relation between 2 tables
'
'#WARNING :         This delete ALL relations not only the PRIMARY one
'                   Always return true to no block the code for a non existing relation
'
'#Contributors :    Victor Hüni
'#Creation Date :   30.04.2021
'#Version :         1.0.0
'
'#References :      Microsoft Office 16.0 Access Database Engine Object (DAO)
'#Dependencies :
'#Related :
'
'#Parameters :
'   p_ParentTable [String]  :   Parent Table Name within the DB
'   p_ChildTable [String]   :   Child Table Name within the DB
'
'#Return :
'   [Boolean] : Always return True. Relationship deleted with success
'
'#Example :
'
'#Source :https://www.devhut.net/2020/01/08/ms-access-vba-delete-table-relationships/ BY Daniel Pineault, CARDA Consultants Inc.
'----------------------------------------------------------------------------------------------------------------------------------------------------
Public Function DeleteTableRelationships(p_ParentTable As String, p_ChildTable As String) As Boolean

    Dim Acc As New DAO.DBEngine
    Dim oDB As DAO.Database
    Dim oRlt As DAO.Relation
 
    On Error GoTo error_handler
 
    Set oDB = Acc.OpenDatabase(Me.GetDBFullPath, False, False) 'Init Database in the DAO workspace
    
    For Each oRlt In oDB.Relations                  'Loop through all relations declared wihtin the database
    
        If (oRlt.Table = p_ParentTable And oRlt.ForeignTable = p_ChildTable) Or (oRlt.Table = p_ChildTable And oRlt.ForeignTable = p_ParentTable) Then  'If a relation is defined between the two table in either sense
        
            oDB.Relations.Delete (oRlt.name)    'Delete the relation
            oDB.Relations.Refresh               'Refresh the relations definition model for the deletion to take effect

            logs.Log_Action "DeleteTableRelationships", Actions.Supprimer, Array("Suppression de relation: ", p_ParentTable, p_ChildTable)
            
        End If
        
    Next oRlt
    
    DeleteTableRelationships = True 'To not block the code, either if there is no relation deleted OR if there are severals, return True
    
    Exit Function
    
error_handler:
    On Error GoTo 0
    If Not Acc Is Nothing Then Set Acc = Nothing
    If Not oRlt Is Nothing Then Set oRlt = Nothing
    If Not oDB Is Nothing Then Set oDB = Nothing
    Errors = Err.Number
    logs.Log_Erreur "DeleteTableRelationships", "No Query", p_ParentTable, p_ChildTable
    DeleteTableRelationships = False
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------
'#Name :            CreateTableRelationships
'
'#Description :     Execute the compact & repair function of Access on a ACCDB file to manage the size and avoid data corruption
'
'#WARNING :         This Function copy/paste the db and compact the copy AND then delete the old one
'                   The field used as an index should have the same name in both table
'                   fieldname could be an array to create a composite relation
'
'#Contributors :    Victor Hüni
'#Creation Date :   30.04.2021
'#Version :         1.0.0
'
'#References :      Microsoft Office 16.0 Access Database Engine Object (DAO)
'#Dependencies :
'#Related :
'
'#Parameters :
'   p_ParentTable [String]  :   Parent Table Name within the DB
'   p_ChildTable [String]   :   Child Table Name within the DB
'   p_FieldName  [Variant]  :   Field in both table that should be related
'   p_Attributes [Long]     :   RelationAttributeEnum enumeration(DAO)
'        dbRelationDeleteCascade     4096   Deletions cascade
'        dbRelationDontEnforce          2   Relationship not enforced (no referential integrity)
'        dbRelationInherited            4   Relationship exists in the database containing the two linked tables
'        dbRelationLeft          16777216   Microsoft Access only. In Design view, display a LEFT JOIN as the default join type.
'        dbRelationRight         33554432   Microsoft Access only. In Design view, display a RIGHT JOIN as the default join type.
'        dbRelationUnique               1   One-to-one relationship
'        dbRelationUpdateCascade      256   Updates cascade
'
'#Return :
'   [Boolean] : Relationship created with success OR Not
'
'#Example :
'
'#Source :
'    https://docs.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/relationattributeenum-enumeration-dao
'    https://docs.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/database-createrelation-method-dao
'----------------------------------------------------------------------------------------------------------------------------------------------------
Public Function CreateTableRelationships(p_ParentTable As String, p_ChildTable As String, ByVal p_FieldName As Scripting.Dictionary, Optional p_Attributes As Long = 0, Optional p_Debug As Boolean = False) As Boolean
    
    Dim Acc As New DAO.DBEngine
    Dim oDB As DAO.Database
    Dim oRlt As DAO.Relation
    Dim tdfParentTable As DAO.TableDef
    Dim tdfChildTable As DAO.TableDef
    Dim oFld As DAO.Field
    Dim fld As Variant
    Dim idxStr As String
    
    On Error GoTo error_handler
    
    Set oDB = Acc.OpenDatabase(Me.DbFilePath & SEP_FOLDER & Me.DbName & Me.Ext, False, False)
    
    Set tdfParentTable = oDB.TableDefs(p_ParentTable)   'Get Parent table definition
    Set tdfChildTable = oDB.TableDefs(p_ChildTable)     'Get Child table definition

    Set oRlt = oDB.CreateRelation(tdfParentTable.name & tdfChildTable.name, tdfParentTable.name, tdfChildTable.name, p_Attributes) 'Initiate the relation between the two tables
    
    For Each fld In p_FieldName.Keys         'Loop throught the potential array of index
    
        idxStr = idxStr & ";" & fld & ", " & p_FieldName(fld)
        
        Set oFld = oRlt.CreateField(fld)    'Create the field in the relation
        oFld.ForeignName = p_FieldName(fld) 'Define foreign key
        oRlt.Fields.Append oFld             'Add the field
        oDB.Relations.Refresh               'Refresh the realtions to make the new fields available in the data model
        
    Next fld
    
    oDB.Relations.Append oRlt               'Add relations to the data model
    oDB.Relations.Refresh                   'Refresh relations to make the new one available
    
    logs.Log_Action "CreateTableRelationships", Actions.Creer, Array("Création de relation: DB: " & DbName, p_ParentTable, p_ChildTable, "Index : " & idxStr)
    
    
    If p_Debug Then
    
        ' Print report
        Debug.Print "Properties of " & oRlt.name & " Relation"
        Debug.Print " Table = " & oRlt.Table
        Debug.Print " ForeignTable = " & oRlt.ForeignTable
        Debug.Print "Fields of " & oRlt.name & " Relation"
        
        For Each fld In p_FieldName.Keys
            With oRlt.Fields(fld)
                Debug.Print " " & .name
                Debug.Print " Name = " & .name
                Debug.Print " ForeignName = " & .ForeignName
            End With
        Next fld
    End If
    
    'Ferme la base de données
    oDB.Close
    Set oDB = Nothing

    CreateTableRelationships = True
    
    Exit Function
    
error_handler:
    On Error GoTo 0
    If Not Acc Is Nothing Then Set Acc = Nothing
    If Not oRlt Is Nothing Then Set oRlt = Nothing
    If Not oDB Is Nothing Then Set oDB = Nothing
    Errors = Err.Number
    logs.Log_Erreur "CreateTableRelationships", "No Query", p_ParentTable, p_ChildTable, idxStr
    CreateTableRelationships = False
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------
'#Name :            GetPrimaryKey
'
'#Description :     Get the Primary key of a table
'
'#WARNING :         Doesn't work for Composite primary key
'
'#Contributors :    Victor Hüni
'#Creation Date :   30.04.2021
'#Version :         1.0.0
'
'#References :      Microsoft Office 16.0 Access Database Engine Object (DAO)
'#Dependencies :
'#Related :
'
'#Parameters :
'   p_TableName [String] : Name of the table for which we want to know the PK
'
'#Return :
'   [string] : Name of the PK fields
'
'#Example :
'
'#Source :
'----------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetPrimaryKey(ByVal p_TableName As String, Optional p_Debug As Boolean = False) As Variant

    Dim Acc As New DAO.DBEngine
    Dim oDB As DAO.Database
    Dim tdfCurrent As TableDef
    Dim idxLoop As DAO.index
    Dim prpLoop As DAO.Property
     
    Set oDB = Acc.OpenDatabase(Me.DbFilePath & SEP_FOLDER & Me.DbName & Me.Ext, False, False)
    Set tdfCurrent = oDB.TableDefs(p_TableName)
    With tdfCurrent
       
        For Each idxLoop In .Indexes    'Enumerate Indexes collection.
            With idxLoop
                If p_Debug Then
                    Debug.Print " " & .name
                    Debug.Print " Fields"                       ' Enumerate Fields collection of each Index object
                    Debug.Print CStr(.Fields)
                    Debug.Print " Properties"
                    For Each prpLoop In .Properties             'Enumerate Properties collection of each Index object.
                        Debug.Print " " & prpLoop.name & " = " & IIf(prpLoop = "", "[empty]", prpLoop)
                    Next prpLoop
                End If
                
                For Each prpLoop In .Properties                 ' Enumerate Properties collection of each Index object.
                    If prpLoop.name = "Primary" Then            'If the current property is a PK
                        If prpLoop Then                         'And the value of the propertie is true
                            GetPrimaryKey = Split(Replace(Right(.Fields, Len(.Fields) - 1), ";", ""), "+") 'Return the fields name without the + added at the beginning
                            oDB.Close
                            Exit Function
                        End If
                    End If
                Next prpLoop

            End With
        Next idxLoop
    End With
     
    oDB.Close
     
End Function

'====================================================================================================================================================
'   ETL METHODS
'====================================================================================================================================================

Public Function TransferDataWithTransaction(ByVal p_SourcePath As String, ByVal p_SourceType As Integer, ByVal p_TargetPath As String, ByVal p_TargetTable As String, ByVal p_TargetFields As String, ByVal p_Debug As Boolean, ByVal IsInsertOnly As Boolean, ByVal SourceCouldBeEmpty As Boolean) As Boolean

    Dim rs As ADODB.Recordset
    Dim res As Boolean

    res = False
    
    'If SourceQuery & the connection string have been set
    If Not (SourceQuery = vbNullString) And _
       Not (InsertQuery = vbNullString) And _
       IIf(IsInsertOnly, (UpdateQuery = vbNullString), Not (UpdateQuery = vbNullString)) And _
       InitCnStr(p_SourcePath, p_SourceType) Then
        
        'Get Source Data
        If ExecuteParameters(SourceQuery, False, False, True) Then
            If Not (IsSelectRecordSetEmpty) Then
                Set rs = rsSource
                rs.MoveFirst
                   
                'Display SourceQuery recordset in the debug sheet
                If p_Debug Then
                    Dim arr As Variant
                    Dim i As Long
                    wsDebug.Cells.ClearContents
                    For i = 1 To rsSource.Fields.Count
                        wsDebug.Cells(1, i) = rsSource.Fields(i - 1).name
                    Next i
                    arr = TransposeArray(rs.GetRows)
                    range(wsDebug.Cells(2, 1), wsDebug.Cells(UBound(arr, 1) + 2, UBound(arr, 2) + 1)) = arr
                End If
                
                'Transfer the source recordset data to the target
                res = InsertUpdateRecordsetInAccdbWithTransaction(p_TargetPath, p_TargetTable, p_TargetFields, IsInsertOnly)
                    
            Else
                res = SourceCouldBeEmpty
            End If
        End If
    End If
    
    TransferDataWithTransaction = res

End Function

Private Function InsertUpdateRecordsetInAccdbWithTransaction(ByVal p_target_path As String, ByVal target_table_name As String, ByVal target_field_list As String, Optional IsInsertOnly As Boolean) As Boolean

    Dim rs As Boolean
    
    Dim el As Variant
    Dim i_pk As Long, i_field As Long, i_record As Long
    Dim condition As String, conditionString As String
    Dim dictParamsValue As Dictionary
    Dim arrSQL As Variant
    
    Dim pk As Variant
    Dim selectRecord As String
    Dim arrWhereRecords As Variant
    Dim whereRecord As String
    Dim arrValues As Variant
    Dim checkQuery As String
    
    Dim recordCounter As Long
    
    rs = False
    ReDim arrSQL(0 To 1, 0 To 0)

    On Error GoTo err_hdl:

    'Init Connection String -------------------------------------------------------------------------------------------------------------------------
    If Not InitCnStr(p_target_path, e_DataSourceTypes.ACCDB) Then GoTo err_hdl

    'Open Connection --------------------------------------------------------------------------------------------------------------------------------
    If Not OpenConnection Then GoTo err_hdl

    'If there is something to insert/update
    If rsSource Is Nothing Then GoTo err_hdl
    If Not (rsSource.EOF And rsSource.BOF) Then
        rsSource.MoveLast
        rsSource.MoveFirst 'Start fromt he beginning
    Else
        GoTo err_hdl
    End If

    pk = GetPrimaryKey(target_table_name) 'Get primary key of the target table
    If Not IsEmpty(pk) Then
        arrWhereRecords = pk
    End If
    
    'ReDim arrSQL(0 To rsSource.RecordCount - 1, 0 To 1)
    
        
    'Loop through all source records
    Do Until rsSource.EOF
        'Reset the Record empty to be sure to insert the next record if there is no check query executed
        m_IsSelectRecordSetEmpty = True
        
        'Build Check Query
        'TODO: Avoid check if not using pk. IsInsertOnly?
        If Not IsEmpty(pk) Then
            ReDim arrValues(UBound(pk))
            selectRecord = Join(pk, ",")
            For i_pk = 0 To UBound(pk)
                arrValues(i_pk) = rsSource.Fields(pk(i_pk)).Value
                arrWhereRecords(i_pk) = pk(i_pk) & " = ? "
            Next i_pk
            whereRecord = Join(arrWhereRecords, "AND ")
            Call ExecuteParameters("SELECT " & selectRecord & " FROM " & target_table_name & " WHERE " & whereRecord, False, False, False, arrValues)
        End If
        
        'Build hash map with Parameters Name AS key & the value of the field AS value
        Set dictParamsValue = New Dictionary
        For i_field = 1 To rsSource.Fields.Count
            Call dictParamsValue.Add("which_" & rsSource.Fields(i_field - 1).name, rsSource.Fields(i_field - 1).Value)
        Next i_field
        
        If IsInsertOnly Then
            
            If m_IsSelectRecordSetEmpty Then
                'INSERT : If the record is Not In recordset
                If (Trim(arrSQL(0, 0)) <> "") Then ReDim Preserve arrSQL(0 To 1, 0 To UBound(arrSQL, 2) + 1)
                arrSQL(0, UBound(arrSQL, 2)) = InsertQuery
                Set arrSQL(1, UBound(arrSQL, 2)) = dictParamsValue
            End If
            
        Else
        
            If Not (m_IsSelectRecordSetEmpty) Then
                'UPDATE : If the record is In recordset
                
                'Avoid creating update that could be harmful due to the field concerned
                '(If I insert only pk field, I could not update a pk field without generating a error)
                
                If (Trim(arrSQL(0, 0)) <> "") Then ReDim Preserve arrSQL(0 To 1, 0 To UBound(arrSQL, 2) + 1)
                arrSQL(0, UBound(arrSQL, 2)) = UpdateQuery
                Set arrSQL(1, UBound(arrSQL, 2)) = dictParamsValue
            Else
                'INSERT : If the record is Not In recordset
                If (Trim(arrSQL(0, 0)) <> "") Then ReDim Preserve arrSQL(0 To 1, 0 To UBound(arrSQL, 2) + 1)
                arrSQL(0, UBound(arrSQL, 2)) = InsertQuery
                Set arrSQL(1, UBound(arrSQL, 2)) = dictParamsValue
            End If
            
        End If

        If Not rsTarget Is Nothing Then
            rsTarget.Close
            Set rsTarget = Nothing
        End If
    
         rsSource.MoveNext
    Loop



    'Close Source ---------------------------------------------------------------------------------------------------------------------------
    rsSource.Close                                      'disconnect the recordset and close the connection
    Set rsSource = Nothing
    CloseConnection
    
    If (Trim(arrSQL(0, 0)) <> "") Then
    
        arrSQL = Transpose2ColArray(arrSQL)
        ExecuteTransaction (arrSQL)
    End If

    rs = True
    InsertUpdateRecordsetInAccdbWithTransaction = rs
    Exit Function
    
err_hdl:
    logs.Log_Erreur "InsertUpdateRecordsetInAccdbWithTransaction" '"Query Source :" & CStr(rsSource.source) & " - Query Target : " & CStr(rsTarget.source)
    InsertUpdateRecordsetInAccdbWithTransaction = rs


End Function

Private Function FieldExists(rs As Recordset, sFieldName As String) As Boolean
    Dim fld As Field
    FieldExists = False
    For Each fld In rs.Fields
        If UCase(fld.name) = UCase(sFieldName) Then
            FieldExists = True
            Exit Function
        End If
    Next fld
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------
'#Name :            ExportRecordsetInEasyFormat
'#Description :     Export a recordsetin a new csv file & save it
'#Warning :         Need only one list object on the selected source tab
'#Contributors :    Victor Hüni
'#Creation Date :   15.12.2020
'#Version :         1.0.0
'
'#References :      None
'#Dependencies :    mod_GlobalRef
'#Related :
'
'#Parameters :
'   src_wb [Workbook]   : Xlsx workbook where is stored the source worksheet
'   src_ws [Worksheet]  : source worksheet to be copy
'#Return :
'   Boolean : True = Success / False = Error (go Check the Logs Folder)
'
'#Example :
'
'#Source :
'----------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ExportRecordsetInEasyFormat(target_path As String, ByVal p_Type As Long) As Boolean

    Dim wbNew As Workbook
    Dim wsNew As Worksheet
    Dim i As Integer
    Dim rs As Boolean
       
    rs = False
    
    'Open Connection --------------------------------------------------------------------------------------------------------------------------------
    If OpenConnection Then
    
        If Not Me.rsSource Is Nothing Then                      'If the source recordset is empty, no data to export, exit function
        
            'Begin export ---------------------------------------------------------------------------------------------------------------------------
            Me.rsSource.MoveFirst                               'Start fromt he beginning

            'Init Target--------------------------------------------------------------------------------------------------
            Set wbNew = Workbooks.Add
            With wbNew
                .Activate
                Set wsNew = .Sheets(1)                   'Get the standard Sheet object from the destination workbook
                'Copy Source to Target ----------------------------------------------------------------------------------------------------------------------
                For i = 0 To rsSource.Fields.Count - 1
                    Cells(1, i + 1).Value = rsSource.Fields(i).name
                Next
                range("A2").CopyFromRecordset Me.rsSource
                
                'TODO: Check , Local:=True
                .SaveAs Filename:=target_path, FileFormat:=m_VBADataSourceTypes(p_Type), CreateBackup:=False
                            
                .Close
            End With
            rs = True
        End If
    End If
    
    ExportRecordsetInEasyFormat = rs
    Exit Function

error_handler:
    Errors = Err.Number
    logs.Log_Erreur "ExportRecordsetInEasyFormat", "No Query"
    ExportRecordsetInEasyFormat = False

End Function


'====================================================================================================================================================
'   MISCELLANOUS METHODS
'====================================================================================================================================================

'----------------------------------------------------------------------------------------------------------------------------------------------------
'#Name :            Transpose2ColArray
'
'#Description :     Use to Transpose a 2 col array when preparing a array of transaction
'
'#WARNING :
'
'#Contributors :    Victor Hüni
'#Creation Date :   19.08.2020
'#Version :         1.0.0
'
'#References :
'#Dependencies :
'#Related :
'
'#Parameters :
'   p_Array [Variant]   : Array to transpose
'
'#Return :
'   Variant : transposed Array
'
'#Example :
'
'#Source :
'----------------------------------------------------------------------------------------------------------------------------------------------------
Private Function Transpose2ColArray(ByVal p_Array As Variant) As Variant
    
    If Not (IsArrayEmpty(p_Array)) And IsArray(p_Array) Then
        'Execute all query build previously
        Dim X As Long
        Dim XUpper As Long
        Dim tempArray As Variant
        
        XUpper = UBound(p_Array, 2)
    
        ReDim tempArray(0 To XUpper, 0 To 1)
        
        For X = 0 To XUpper
            tempArray(X, 0) = p_Array(0, X)
            Set tempArray(X, 1) = p_Array(1, X)
            
        Next X
        Transpose2ColArray = tempArray
    Else
        Debug.Print "Error -> " & "Transpose2ColArray"
    End If

End Function

