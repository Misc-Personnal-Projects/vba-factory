VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_XxxsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    '----------------------------------------------------------------------------------------------------------------------------------------------------
'#Name : <ClassName>
'#Description :
'#WARNING :
'#Contributors :
'#Version :
'
'#References : Reference_Name
'#Dependencies : classes à importer
'#Related :
'   MethodsName
'
'#Members :
'   Name1 [Type1] : Description1
'   Name2 [Type2] : Description2
'
'#Source :
'----------------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit

'README :
'1) Replace All Xxx by the name of your Object with a Capital Letter at the Begining
'2) Replace All YYY By the name of your Object all in capital letters
'3) Uncomment line with Remplacer ID by your unique value member

Private m_Xxxs As collection

'Logs
'Update the log message according to you need
Private logs As cls_Logs
Private Const SOURCE_LOG As String = "cls_XxxsList"
Private Const INSERT_LOG_MSG As String = "Xxx Cré(e)"
Private Const UPDATE_LOG_MSG As String = "Xxx Mis(e) à jour"
Private Const DELETE_LOG_MSG As String = "Xxx Supprimé(e)"

'Update this number the number of member in the cls_Xxx Class (Start From 0)
Private Const SUB_CLASS_MEMBERS_INDEX As Integer = 2

'Query
Private Const GET_ALL_YYYS As String = vbNullString
Private Const GET_YYYS_BY_ID As String = vbNullString
Private Const GET_YYYS_BY_FILTER As String = "SELECT * FROM [ZZZ] WHERE primary_key IS NOT NULL [CONDITIONS]"
Private Const GET_EXT_DATA As String = "SELECT * FROM [ZZZ] LEFT JOIN [WWW] ON YYY.pk = WWW.pk WHERE pk IN  ([QUESTIONS_MARKS])"
Private Const CREATE_NEW_YYY As String = vbNullString
Private Const UPDATE_YYY As String = vbNullString
Private Const DELETE_YYY As String = vbNullString
Private Const CHECK_IF_YYY_EXIST As String = vbNullString

Private Sub Class_Initialize()
    Set m_Xxxs = New collection
    Set logs = New cls_Logs
End Sub

Private Sub Class_Terminate()
    Set m_Xxxs = Nothing
    Set logs = Nothing
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_Xxxs.[_NewEnum]
End Property

Public Sub Add(p_Xxx As cls_Xxx)
    'm_Xxxs.Add p_Xxx, CStr(p_Xxx.Id) -> Remplacer ID par your unique value
End Sub
 
Public Sub Remove(index As Variant)
    m_Xxxs.Remove CStr(index)
End Sub

Public Property Get Item(index As Variant) As cls_Xxx
Attribute Item.VB_UserMemId = 0
    Set Item = m_Xxxs.Item(CStr(index))
End Property

Public Property Get Count() As Long
    Count = m_Xxxs.Count
End Property

Public Sub Clear()
    Set m_Xxxs = New collection
End Sub

Public Function Contains(key As Variant) As Boolean

    Dim Obj As Variant
    On Error GoTo Err
    
    Set Obj = Me.Item(key)
    Contains = True
    Exit Function
    
Err:
    Err.Clear
    On Error GoTo 0
    Contains = False

End Function



Public Function GetAllRecords() As Boolean
    
    Dim Xxx As cls_Xxx
    Dim db As New cls_Database
    Dim rs As ADODB.Recordset
    Dim res As Boolean
    
    Me.Clear
    
    res = False
    
    If db.ExecuteParameters(GET_ALL_YYYS, False, True, True) Then
       If Not (db.IsSelectRecordSetEmpty) Then
           Set rs = db.rsSource
           rs.MoveFirst
           Do While Not rs.EOF
               Set Xxx = New cls_Xxx
               With Xxx
                   '<GETRECORDSET>
               End With
               Me.Add Xxx
               rs.MoveNext
           Loop
           res = True
       End If
    End If
    
    GetAllRecords = res
    
End Function

Public Function GetList() As Variant

    Dim arrRes As Variant
    Dim Xxx As cls_Xxx
    Dim i As Long
    
    If Me.Count > 0 Then
        ReDim arrRes(0 To Me.Count - 1, 0 To SUB_CLASS_MEMBERS_INDEX)
        i = 0
        For Each Xxx In Me
            With Xxx
                '<GETOBJECT>
            End With
            i = i + 1
        Next Xxx

    Else
        ReDim arrRes(0 To 0, 0 To SUB_CLASS_MEMBERS_INDEX)
    End If
    
    GetList = arrRes
    
End Function

Public Function GetXxxsByFilters(Optional ByVal p_WhereRequest As String = "") As Variant
    
    Dim db As New cls_Database
    Dim res As Variant
    
    res = Array()

    With db
        If .ExecuteParameters(Replace(GET_YYYS_BY_FILTER, "[CONDITIONS]", p_WhereRequest), False, True, True) Then
            If Not (.IsSelectRecordSetEmpty) Then
                res = TransposeArray(.rsSource.GetRows)
            End If
        End If
    End With

    GetXxxsByFilters = res

End Function

Public Function GetIdList() As Variant
    
    Dim arrRes As Variant
    Dim Xxx As New cls_Xxx
    Dim i As Long
    
    If Me.Count > 0 Then
        ReDim arrRes(0 To Me.Count - 1)
        i = 0
        For Each Xxx In Me
            'arrRes(i) = Xxx.Id -> Remplacer ID par your unique value
            i = i + 1
        Next Xxx

    Else
        ReDim arrRes(0 To 0)
    End If
    
    GetIdList = arrRes

End Function

Public Function GetExternalData() As Variant
    
    Dim Yyy As Object
    Dim db As New cls_Database
    Dim rs As ADODB.Recordset
   
    Dim idArr As Variant
    Dim questionsMarks As String
    Dim strQuery As String
    
   
    Dim res As Boolean
    res = False
    
    idArr = Array()
    idArr = GetIdList()
    
    If IsArrayEmpty(idArr) Then GoTo Exit_Function
        
    questionsMarks = db.TransformParamsInQuestionsMarks(idArr)
    strQuery = Replace(GET_EXT_DATA, "[QUESTIONS_MARKS]", questionsMarks)
    
    If db.ExecuteParameters(strQuery, False, True, True, idArr) Then
        If Not (db.IsSelectRecordSetEmpty) Then
            Set rs = db.rsSource
            rs.MoveFirst
            res = TransposeArray(rs.GetRows)
            rs.MoveFirst
            Do While Not rs.EOF
                'OTHER OBJECT TO CREATE : Set Xxx = New cls_Yyy
                With Yyy
                   'Assigne recordset field to class member
                   'Use the generators
                End With
                'ADD OBJECT TO NESTED LIST : Me.Item(rs.Fields("pk")).Yyys.Add Yyy
                Set Yyy = Nothing
                rs.MoveNext
            Loop
        End If
    
    End If

Exit_Function:
    GetExternalData = res

End Function

Public Function InsertRecord(p_Xxx As cls_Xxx) As Long
   
    Dim cls_Xxx As cls_Xxx
    Dim db As New cls_Database
    Dim res As Long
    
    res = 0
    
    'Add p_Xxx member value to the ExecuteParameters Array
    If db.ExecuteParameters(CREATE_NEW_YYY, False, True, True) Then
        If db.AffectedRows > 0 Then
            'Add
            '<ADDRECORD>
            
            'log
            '<LOGADDRECORD>
        End If
    End If
    
    InsertRecord = res

End Function

Public Function UpateRecord(p_Xxx As cls_Xxx) As Boolean
   
    Dim db As New cls_Database
    Dim res As Boolean
    
    res = False
    
    'Add p_Xxx member value to the ExecuteParameters Array
    If db.ExecuteParameters(UPDATE_YYY, False, True, True) Then
        If db.AffectedRows > 0 Then
            'Update
            '<UPDATERECORD>
            
            'log
            '<LOGUPDATERECORD>
            
            'validation
            res = db.LastId
        End If
    End If
    
    UpateRecord = res

End Function


Public Function DeleteRecord(p_Id As Long) As Boolean
   
    Dim cls_Xxx As cls_Xxx
    Dim db As New cls_Database
    Dim res As Boolean
    
    If db.ExecuteParameters(DELETE_YYY, False, True, True, p_Id) Then
    
        If db.AffectedRows > 0 Then
            'Delete
            '<DELETERECORD>
            
            'log
            '<LOGDELETERECORD>
            
            'validation
            res = db.LastId
        End If
        
    Else
    
        If db.Errors = -2147467259 Then
            Call Msgbox_Perso("Suppression Élément", "Impossible de supprimer un élément déjà assigné")
        End If
        res = False
        
    End If
    
    DeleteRecord = res

End Function

Public Function CheckIfExists(p_Value As String) As Long

    Dim db As New cls_Database
    Dim rs As ADODB.Recordset
    Dim res As Long
        
    res = 0
        
    If db.ExecuteParameters(CHECK_IF_YYY_EXIST, False, True, True, p_Value) Then
        If Not (db.IsSelectRecordSetEmpty) Then
            Set rs = db.rsSource
            res = rs.Fields(0).Value
        End If
    End If
    
    CheckIfExists = res
    
End Function




