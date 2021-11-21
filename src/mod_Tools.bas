Attribute VB_Name = "mod_Tools"
'----------------------------------------------------------------------------------------------------------------------------------------------------
'#Name : XxxYyy
'#Description :
'#WARNING :
'#Contributors :
'#Version : v1.0.1
' 09.04.21: modif GetUNCPath to get correct path on local drive
'
'#References : Reference_Name
'#Dependencies : classes à importer
'#Related :
'   FunctionName :
'   ProcedureName :
'
'#Source :
'----------------------------------------------------------------------------------------------------------------------------------------------------

Option Explicit

'====================================================================================================================================================
'   FILESYSTEM FUNCTIONS
'====================================================================================================================================================

'----------------------------------------------------------------------------------------------------------------------------------------------------
'#Name : GetUNCPath
'#Description :
'#WARNING :
'#Contributors :
'
'#Dependencies : classes à importer
'#Related :
'   SimilarFunction
'
'#Parameters :
'   Name1 [Type1] : Description1
'   Name2 [Type2] : Description2
'
'#Return : Type
'
'#Example : call XxxYyy (Name1, Name2)
'
'#Source :
'----------------------------------------------------------------------------------------------------------------------------------------------------
Function GetUNCPath(strMappedDrive As String) As String
    
    'declarations
    Dim objFso  As New FileSystemObject
    Dim strDrive As String, strShare As String, uncPath As String

    'init
    On Error Resume Next
    If (objFso.GetFolder(strMappedDrive).Drive.DriveLetter <> "") Then
        'Separate the mapped letter from any following sub-folders
        strDrive = objFso.GetDriveName(strMappedDrive)
        
        'find the UNC share name from the mapped letter
        strShare = objFso.Drives(strDrive & "\").ShareName
        
        'The Replace function allows for sub-folders
        'of the mapped drive
        If (strShare <> "") Then
            uncPath = Replace(strMappedDrive, strDrive, strShare)
        Else
            uncPath = strMappedDrive
        End If
    Else
        'don't change path
        uncPath = strMappedDrive
    End If
    
    'clean
    Call Err.Clear
    On Error GoTo 0
    Set objFso = Nothing 'Destroy the object
    
    'return
    GetUNCPath = uncPath

End Function '--- GetUNCPath ---'

'----------------------------------------------------------------------------------------------------------------------------------------------------
'#Name :            GetFileInformation
'#Description :     Get information of files from a source folder
'#Warning :         Need Scripting Runtime Reference
'#Contributors :    Jérôme Favre-Rochex
'#Creation Date :   06.05.2021
'#Version :         1.0.0
'
'#References :      Microsoft Scripting Runtime
'#Dependencies :    mod_GlobalRef
'#Related :
'
'#Parameters :
'   src_path [string]   : Source FilePath
'   file_name [string]  : Filename Model
'#Return :
'   Variant : List of files with information
'
'#Example :
'
'#Source :
'----------------------------------------------------------------------------------------------------------------------------------------------------
Function GetFileInformation(ByVal src_path As String, ByVal file_model As String) As Variant

    'declarations
    Dim fso As New FileSystemObject
    Dim fileTmp As File
    Dim res As Variant

    'initialisation
    ReDim res(0 To 2, 0 To 0)

    'check path
    If (fso.FolderExists(src_path)) Then
        For Each fileTmp In fso.GetFolder(src_path).Files
            'check name
            If (InStr(1, LCase(fileTmp.name), LCase(file_model)) <> 0 And InStr(1, fileTmp.name, "~") = 0) Then
                'redim if necessary
                If (Trim(res(0, 0)) <> "") Then ReDim Preserve res(LBound(res, 1) To UBound(res, 1), LBound(res, 2) To UBound(res, 2) + 1)
                
                'set values
                res(0, UBound(res, 2)) = fileTmp.name
                res(1, UBound(res, 2)) = Format(fileTmp.DateLastModified, "dd.mm.yyyy")
                res(2, UBound(res, 2)) = Format(fileTmp.DateLastModified, "hh:MM:ss")
            End If
        Next fileTmp
    End If
    
    'invert
    res = TransposeArray(res)
    
    'return
    GetFileInformation = res
    
End Function '--- GetFileInformation ---'

'----------------------------------------------------------------------------------------------------------------------------------------------------
'#Name :            CopyingAFile
'#Description :     Copying any tipe of file from a source folder to a destination folder
'#Warning :         Need Scripting Runtime Reference
'#Contributors :    Victor Hüni
'#Creation Date :   15.12.2020
'#Version :         1.0.0
'
'#References :      Microsoft Scripting Runtime
'#Dependencies :    mod_GlobalRef
'#Related :
'
'#Parameters :
'   src_path [string]   : Source FilePath
'   src_name [string]   : Source Filename
'   dst_path [string]   : Destination FilePath
'   dst [string]        : Destination FileName
'   ext_files [string]  : Extension to add to the src & dst filename (.xlsx, .csv. accdb...)
'#Return :
'   Boolean : True = Success / False = Error (go Check the Logs Folder)
'
'#Example :
'
'#Source :
'----------------------------------------------------------------------------------------------------------------------------------------------------
Function CopyingAFile(ByVal src_path As String, ByVal src_name As String, ByVal dst_path As String, ByVal dst_name As String, ByVal ext_files As String) As Boolean

    Dim fso
    Dim sFile As String
    Dim sDFile As String
    Dim sSFolder As String
    Dim sDFolder As String
    Dim logs As New cls_Logs

    On Error GoTo err_hdl

    'Init Source-------------------------------------------------------------------------------------------------------
    sFile = src_name & IIf(InStr(1, src_name, ext_files) > 0, "", ext_files)    'This is Your File Name which you want to Copy
    sSFolder = src_path                                                         'Change to match the source folder path

    'Init Destination--------------------------------------------------------------------------------------------------
    sDFile = dst_name & IIf(InStr(1, dst_name, ext_files) > 0, "", ext_files)   'This is Your File Name which you want to Copy
    sDFolder = dst_path                                                         'Change to match the destination folder path

    'Copy File---------------------------------------------------------------------------------------------------------
    Set fso = CreateObject("Scripting.FileSystemObject")                    'Create Filesystem object
    If Not fso.FileExists(sSFolder & SEP_FOLDER & sFile) Then               'Checking If File Is Located in the Source Folder
        MsgBox "Specified File Not Found", vbInformation, "Not Found"       'Error Handling
        CopyingAFile = False
        Exit Function
    End If

    fso.CopyFile (sSFolder & SEP_FOLDER & sFile), sDFolder & SEP_FOLDER & sDFile, True    'Copying files
    'MsgBox "Specified File Copied Successfully", vbInformation, "Done!"    'Success Messages
    
    'Check if destination file exist
    If Not fso.FileExists(sDFolder & SEP_FOLDER & sDFile) Then               'Checking If File Is Located in the Source Folder
        MsgBox "Copy of the File Not Found", vbInformation, "Not Found"      'Error Handling
        CopyingAFile = False
        Exit Function
    End If
    
    CopyingAFile = True
    Exit Function
    
err_hdl:
    Set logs = New cls_Logs
    logs.Log_Erreur "mod_Tools.CopyingAFile", "No Query"
    Set logs = Nothing
    CopyingAFile = False

End Function

'====================================================================================================================================================
'   OPTIMIZATION FUNCTIONS
'====================================================================================================================================================

'----------------------------------------------------------------------------------------------------------------------------------------------------
'#Name :            SwitchOff
'#Description :
'   Switch On or Off all excel process that slow Excel Macros exectution
'       - Calcultation
'       - Screen Updating
'       - Animations
'       - Pages Break
'       - Animations
'
'#Warning :         This function should be use at the begin of the procedure where you to disable everything and MUST BE ADDED AT THE END TO REVERT THE EFFECT
'#Contributors :    Victor Hüni
'#Creation Date :   15.12.2020
'#Version :         1.0.0
'
'#References :      None
'#Dependencies :    None
'#Related :
'
'#Parameters :
'   bSwitchOff [Boolean]   : True = Disable Everything, False = Renable Everything

'#Return : None
'
'#Example :
'
'#Source : https://techcommunity.microsoft.com/t5/excel/9-quick-tips-to-improve-your-vba-macro-performance/m-p/173687#M602
'----------------------------------------------------------------------------------------------------------------------------------------------------
Sub SwitchOff(bSwitchOff As Boolean)

    Dim ws As Worksheet
    
    With Application
    
        If bSwitchOff Then
        
            ' OFF
            .Calculation = xlCalculationManual
            .ScreenUpdating = False
            .EnableAnimations = False
            .EnableEvents = False
            
            
            'switch off display pagebreaks for all worksheets
            For Each ws In ActiveWorkbook.Worksheets
            
                ws.DisplayPageBreaks = False
            
            Next ws
        
        Else
            
            ' ON
            .Calculation = xlAutomatic
            .ScreenUpdating = True
            .EnableAnimations = True
            .EnableEvents = True
        
        End If
    
    End With

End Sub

'====================================================================================================================================================
'   INFO FUNCTIONS
'====================================================================================================================================================
Function Msgbox_Perso(Optional ByVal p_Title As String, Optional ByVal p_message As String, Optional ByVal p_Type As String = TYP_INFORMATION) As Variant
    
    'declarations
    Dim usf As usf_Information
    Dim res As Integer
    
    'init
    Set usf = New usf_Information
    
    'parameters
    If (Not IsMissing(p_Title)) Then usf.title_header.Caption = p_Title
    If (Not IsMissing(p_message)) Then usf.txt_message.Text = p_message
    
    'display buttons
    Select Case p_Type
        Case TYP_QUESTION
            usf.btn_validate.Visible = False
            usf.btn_yes.Visible = True
            usf.btn_no.Visible = True
        Case Else
            usf.btn_validate.Visible = True
            usf.btn_yes.Visible = False
            usf.btn_no.Visible = False
    End Select
    
    'execution
    Call usf.Show(vbModal)
    res = usf.m_return
    Call Unload(usf)
    
    'return
    Msgbox_Perso = res
    
End Function '--- msgbox_perso ---'

'====================================================================================================================================================
'   STRING MANIPULATION FUNCTIONS
'====================================================================================================================================================
Public Function ParseNum(strInput As String) As Long

    Dim matches As IMatchCollection2
    Dim match As IMatch2
    Dim regex As IRegExp2
    
    Set regex = CreateObject("vbscript.regexp")     'Set the RegExp object
    regex.Pattern = "-?\d*\.?\d+"                    'Initiate the matching pattern, this optional minus sign, all numerical carater, optional decimal point, and all numerical caracter after this
    regex.Global = True
    regex.IgnoreCase = True
    Set matches = regex.Execute(strInput)           'Save the matching part of the string in a Collection
    
    If matches.Count = 1 Then                       'If the string has only one part with numerical values
        ParseNum = CLng(matches.Item(0).Value)
    Else                                            'If there is nore than one part of the string with numerical value
        For Each match In matches
            'Function not finisehd. should return an array after looping through the collection
        Next
    End If

End Function

Public Function ParseNumInParenthesis(strInput As String) As Double

    Dim matches As IMatchCollection2
    Dim match As IMatch2
    Dim regex As IRegExp2
    
    Set regex = CreateObject("vbscript.regexp")     'Set the RegExp object
    regex.Pattern = "\([^()]+\)"                    'Initiate the matching pattern, this optional minus sign, all numerical carater, optional decimal point, and all numerical caracter after this
    regex.Global = True
    regex.IgnoreCase = True
    Set matches = regex.Execute(strInput)           'Save the matching part of the string in a Collection
    
    If matches.Count = 1 Then                       'If the string has only one part with numerical values
        ParseNumInParenthesis = Val(Replace(matches.Item(0).Value, "(", ""))
    Else                                            'If there is nore than one part of the string with numerical value
        For Each match In matches
            'Function not finisehd. should return an array after looping through the collection
        Next
    End If

End Function
