VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Feedback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const DATA_FOLDER As String = "\Logs\"


Dim logs As cls_Logs

Private m_Root As String
Private m_Site As String
Private m_AppName As String
Private m_Env As String

Private m_LogName As String
Private m_FeedbackFilePath As String


Public Property Get Root() As String: Root = m_Root: End Property
Public Property Let Root(ByVal p_Root As String): m_Root = p_Root: End Property

Public Property Get Site() As String: Site = m_Site: End Property
Public Property Let Site(ByVal p_site As String): m_Site = p_site: End Property

Public Property Get AppName() As String: AppName = m_AppName: End Property
Public Property Let AppName(ByVal p_AppName As String): m_AppName = p_AppName: End Property

Public Property Get Env() As String: Env = m_Env: End Property
Public Property Let Env(ByVal p_Env As String): m_Env = p_Env: End Property

Public Property Get LogName() As String: LogName = m_LogName: End Property
Public Property Let LogName(ByVal p_LogName As String): m_LogName = p_LogName: End Property

Public Property Get FeedbackFilePath() As String: FeedbackFilePath = m_FeedbackFilePath: End Property
Public Property Let FeedbackFilePath(ByVal p_FeedbackFilePath As String): m_FeedbackFilePath = p_FeedbackFilePath: End Property



Private Sub Class_Initialize()
    Set logs = New cls_Logs
    
    Root = wsAPropos.range("Root").Value
    Site = LCase(wsAPropos.range("Site").Value)
    AppName = LCase(wsAPropos.range("AppName").Value)
    Env = LCase(wsAPropos.range("Env").Value)
    
    FeedbackFilePath = Root & DATA_FOLDER
    LogName = "feedback.lfo"
    
    logs.Dossier = FeedbackFilePath
    logs.Fichier = LogName
    logs.Entetes = "login|date|heure|remarque|scenario|tache"
    
End Sub


Public Sub Get_Comment()

    Dim ls As Variant
    ReDim ls(0 To 5)
    
    ReDim g_transfert(0 To 3)
    Call usf_Feedback.Show(vbModal)
    
    'check
    If (Not g_transfert(3)) Then Exit Sub
    
    ls(0) = Environ("username")
    ls(1) = Date
    ls(2) = Time
    ls(3) = g_transfert(0)
    ls(4) = g_transfert(1)
    ls(5) = g_transfert(2)
    Call logs.Log_Perso(ls)

End Sub

