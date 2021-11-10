VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Xxx"
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

'Members
'<Members>


'Property
'<Getters&Setters>

'----------------------------------------------------------------------------------------------------------------------------------------------------
'#Name : SetParameters
'#Description : Define object parameters with parameter values passed by reference
'#WARNING :
'#Contributors : Jérôme Favre-Rochex, Victor Hüni
'
'#Dependencies :
'#Related :
'
'#Parameters :
'   params [Variant] : list of parameter values
'
'#Return : true if argument params contains the right number of values
'
'#Example :
'
'#Source :
'----------------------------------------------------------------------------------------------------------------------------------------------------
Public Function SetParameters(ByRef params As Variant) As Boolean
    
    'declaration
    Dim IsSet As Boolean
    
    'init
    IsSet = False
    
    'check
    If (UBound(params, 2) = 2) Then
        'init
        '<Setparams>
        IsSet = True
    End If
    
    'return
    SetParameters = IsSet

End Function '--- SetParameters ---'


'----------------------------------------------------------------------------------------------------------------------------------------------------
'#Name : GetParameters
'#Description : Return object parameters with array of parameter values passed by reference
'#WARNING :
'#Contributors : Jérôme Favre-Rochex, Victor Hüni
'
'#Dependencies :
'#Related :
'
'#Parameters :
'
'#Return : Array of parameter values
'
'#Example :
'
'#Source :
'----------------------------------------------------------------------------------------------------------------------------------------------------
Public Function GetParameters() As Variant
    
    'declaration
    Dim params As Variant
    
    'init
    ReDim params(0 To 1, 0 To 2)
    
    'get parameters
    '<Getparams>
    'return
    GetParameters = params

End Function '--- GetParameters ---'


'----------------------------------------------------------------------------------------------------------------------------------------------------
'#Name : GetIdentity
'#Description : Return identity card of the instance
'#WARNING :
'#Contributors : Jérôme Favre-Rochex, Victor Hüni
'
'#Dependencies :
'#Related :
'
'#Parameters :
'
'#Return : String with all members value
'
'#Example :
'
'#Source :
'----------------------------------------------------------------------------------------------------------------------------------------------------
Public Function GetIdentity() As String
    
    'declaration
    Dim identity As String
    
    'init
    identity = ""
    
    'get parameters
    '<Identity>
    
    'return
    GetIdentity = identity

End Function '--- GetIdentity ---'

