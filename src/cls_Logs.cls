VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Logs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--- cls_logs ---'
'#version:1
'
'#references
'Microsoft Scripting Runtime
'
'#description
'classe qui permets de gérer les logs dans un fichier
'elle commence par créer le dossier et les fichiers si ils n'existent pas,
'puis on enregistre dans ces fichiers leurs logs associés (erreurs, actions,...)
'
'
'Enum à copier/coller dans le Références Global de l'application maître
'Afin de l'utiliser lors de l'appel de la méthode Log_Action
'Public Enum Actions
'    Ouvrir = 0
'    Creer = 1
'    Modifier = 2
'    Annuler = 3
'    Supprimer = 4
'    Sauvegarder = 5
'    Envoyer = 6
'    Fermer = 7
'    Sélectionner = 8
'End Enum
'
'#exemple
'--- méthode test ---'
'Sub Test()
'
'    'déclarations
'    Dim logs As cls_Logs
'    Dim i_test As Integer
'
'    'gestion des erreurs
'    On Error GoTo end_err
'
'    'création d'une instance de cls_log
'    Set logs = New cls_Logs
'    Call logs.Log_Action("Test action", Actions.Ouvrir, Array("1", "2", "3"))
'
'    'création erreur
'    i_test = 65 + 342345 * 23343 * 998
'
'    Exit Sub
'
'end_err:
'    'log de l'erreur
'    Call logs.Log_Erreur("Test", "select", "infi1", "info2", Trim(Time))
'
'End Sub '--- Test ---'

'----------------------------------------------------------------------------------------------------------------------------------------------------
'#Name : XxxYyy
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

'dossier qui contient les logs
Private Const m_dossier As String = "Logs"

'fichiers logs
Private Const M_FICHIER_ERREURS As String = "logs_erreurs.lfo"                 'logs erreurs
Private Const M_FICHIER_ACTIONS As String = "logs_actions.lfo"                  'logs actions

'séparateurs
Private Const SEP_LOG As String = "|"
Private Const SEP_DETAILS As String = ";"
Private Const SEP_DOSSIER As String = "\"

'taille des archives en Bytes
Private Const SIZE_ARCHIVE As Long = 52428800

'entêtes des fichiers logs
Private Const M_HEADS_ERREURS As String = "date|heure|login|Outil|fichier|emplacement|erreur|query|info_1|info_2|info_3"
Private Const M_HEADS_ACTIONS As String = "date|heure|login|Outil|fichier|emplacement|action|details"

'variable de la classe
Private m_login As String

Private m_Actions As Variant

'variable pour logs personnalisé
Private perso_dossier As String
Private perso_fichier As String
Private perso_heads As String

'structure erreur
Private Type Erreur
    source As String
    Number As Long
    Description As String
End Type

'-------------------------------------------------------------------------------------
'#description
'propriété permettant d'affecter un dossier perso
'-------------------------------------------------------------------------------------
Property Let Dossier(ByVal s_valeur As String)

    'affectation du dossier au log perso
    perso_dossier = Trim(s_valeur)
    If (Right(perso_dossier, 1) <> SEP_DOSSIER) Then perso_dossier = perso_dossier & SEP_DOSSIER
    
End Property '--- Dossier ---'


'-------------------------------------------------------------------------------------
'#description
'propriété permettant d'affecter un fichier perso
'-------------------------------------------------------------------------------------
Property Let Fichier(ByVal s_valeur As String)

    'affectation du dossier au log perso
    perso_fichier = Trim(s_valeur)
    
End Property '--- Fichier ---'


'-------------------------------------------------------------------------------------
'#description
'propriété permettant d'affecter une entête perso
'-------------------------------------------------------------------------------------
Property Let Entetes(ByVal s_valeur As String)

    'affectation de l'entête au log perso
    perso_heads = s_valeur
    
End Property '--- Entetes ---'


'-------------------------------------------------------------------------------------
'#description
'méthode d'initialisation de la classe
'
'#important
'cette méthode équivalente à un constructeur nous permets de récupérer l'utilisateur
'actif et d'initialiser les fichiers logs
'-------------------------------------------------------------------------------------
Private Sub Class_Initialize()

    'utilisateur actif
    m_login = Environ("username")
    
    m_Actions = Array("Ouvrir", "Creer", "Modifier", "Annuler", "Supprimer", "Sauvegarder", "Envoyer", "Fermer", "Sélectionner")
        
    'valeurs personnalisées
    perso_dossier = ""
    perso_fichier = ""
    perso_heads = ""
        
End Sub '--- Class_Initialize ---'

'-------------------------------------------------------------------------------------
'#description
'méthode de terminaison de la classe
'
'#important
'cette méthode équivalente à un destructeur nous permets de nettoyer
'les valeurs
'-------------------------------------------------------------------------------------
Private Sub Class_Terminate()
    
    'utilisateur actif
    m_login = ""
    
End Sub '--- Class_Terminate ---'

'-------------------------------------------------------------------------------------
'#description
'méthode qui log une erreur
'
'#paramètre
'p_location: nom de la méthode ou fonction qui envoi l'erreur
'
'#important
'cette méthode sauvegarde l'erreur puis la supprime
'-------------------------------------------------------------------------------------
Sub Log_Erreur(ByVal p_location As String, Optional p_Query As String = "", Optional info_1 As String = "", Optional info_2 As String = "", Optional info_3 As String = "")

    'déclarations
    Dim s_fichier As String
    Dim s_error As Erreur
        
    'sauvegarde l'erreur
    s_error.source = Err.source
    s_error.Number = Err.Number
    s_error.Description = Err.Description
    Call Err.Clear
    
    'init nom du fichier
    s_fichier = ThisWorkbook.path & SEP_DOSSIER & m_dossier & SEP_DOSSIER & M_FICHIER_ERREURS
        
    'vérification des fichiers
    If (Not fn_initialisation) Then Exit Sub
    
    'ajoute le log erreur
    Call mt_textfile_append(s_fichier, _
        Format(Date, "dd.mm.yyyy") & SEP_LOG & _
        Format(Time, "hh:mm:ss") & SEP_LOG & _
        m_login & SEP_LOG & _
        ThisWorkbook.VBProject.name & SEP_LOG & _
        ThisWorkbook.name & SEP_LOG & _
        p_location & SEP_LOG & _
        "(" & s_error.Number & ") " & s_error.Description & SEP_LOG & _
        p_Query & SEP_LOG & _
        info_1 & SEP_LOG & _
        info_2 & SEP_LOG & _
        info_3)

    'clean l'erreur
    Call Err.Clear

End Sub '--- Log_Erreur ---'

'-------------------------------------------------------------------------------------
'#description
'méthode qui log une action
'
'#paramètre
'p_location: nom de la méthode ou fonction qui envoi l'action
'p_action: action lancée
'
'#important
'cette méthode sauvegarder une action avec des points de répères
'numérotés (index) et temporels (timer)
'-------------------------------------------------------------------------------------
Sub Log_Action(ByVal p_location As String, ByVal p_Action As Integer, ByRef p_details As Variant)

    'déclarations
    Dim s_fichier As String
    
    'vérification des fichiers
    If (Not fn_initialisation) Then Exit Sub

    'init nom du fichier
    s_fichier = ThisWorkbook.path & SEP_DOSSIER & m_dossier & SEP_DOSSIER & M_FICHIER_ACTIONS
        
    'ajoute le log erreur
    Call mt_textfile_append(s_fichier, _
        Format(Date, "dd.mm.yyyy") & SEP_LOG & _
        Format(Time, "hh:mm:ss") & SEP_LOG & _
        m_login & SEP_LOG & _
        ThisWorkbook.VBProject.name & SEP_LOG & _
        ThisWorkbook.name & SEP_LOG & _
        p_location & SEP_LOG & _
        m_Actions(p_Action) & SEP_LOG & _
        Join(p_details, SEP_DETAILS))

End Sub '--- Log_Action ---'

'-------------------------------------------------------------------------------------
'#description
'méthode qui log des données personnalisées
'
'#paramètre
'p_datas: tableau des valeurs à ajouter en ligne
'
'#important
'le tableau ne doit avoir qu'une seule dimension car un join est utilisé
'pour créer la ligne
'-------------------------------------------------------------------------------------
Sub Log_Perso(ByRef p_datas As Variant)

    'déclarations
    Dim s_fichier As String
    
    'vérification des fichiers
    If (Not fn_initialisation_perso) Then Exit Sub

    'init nom du fichier
    s_fichier = perso_dossier & perso_fichier

    'ajoute le log perso
    Call mt_textfile_append(s_fichier, Join(p_datas, SEP_LOG))

End Sub '--- Log_Perso ---'


'-------------------------------------------------------------------------------------
'#description
'fonction qui initialise le dossier et les fichiers logs
'-------------------------------------------------------------------------------------
Private Function fn_initialisation() As Boolean

    'declarations
    Dim s_dossier  As String, s_fichier As String
    Dim b_res As Boolean
    
    'gestion des erreurs
    On Error GoTo fn_initialisation_end
    
    'initialisation
    b_res = False
    
    'verification dossier
    s_dossier = ThisWorkbook.path & SEP_DOSSIER & m_dossier
    If (Dir(s_dossier, vbDirectory) = "") Then Call MkDir(s_dossier)
    s_dossier = s_dossier & SEP_DOSSIER
    
    'LOGS ERREURS
    
    'fichier log erreur
    s_fichier = s_dossier & M_FICHIER_ERREURS
    'verification fichier
    If (Dir(s_fichier, vbNormal) = "") Then Call mt_textfile_create(s_fichier, M_HEADS_ERREURS)
    ' Archive file at certain size
    mt_archive_textfile (s_fichier)
    
    'LOGS ACTIONS
    
    'fichier log action
    s_fichier = s_dossier & M_FICHIER_ACTIONS
    'verification fichier
    If (Dir(s_fichier, vbNormal) = "") Then Call mt_textfile_create(s_fichier, M_HEADS_ACTIONS)
    ' Archive file at certain size
    mt_archive_textfile (s_fichier)


    'validation
    b_res = True
    
fn_initialisation_end:
    'clean l'erreur
    Call Err.Clear
    
    'return
    fn_initialisation = b_res
    
End Function '--- fn_initialisation ---'

'-------------------------------------------------------------------------------------
'#description
'fonction qui initialise le dossier et le fichier log personnalisé
'-------------------------------------------------------------------------------------
Private Function fn_initialisation_perso() As Boolean

    'declarations
    Dim b_res As Boolean
    Dim s_fichier As String
    
    'gestion des erreurs
    On Error GoTo fn_initialisation_perso_end
    
    'initialisation
    b_res = False
    
    'verification dossier perso
    If (Trim(perso_dossier) = "") Then GoTo fn_initialisation_perso_end
    If (Dir(perso_dossier, vbDirectory) = "") Then GoTo fn_initialisation_perso_end
    If (Right(perso_dossier, 1) <> SEP_DOSSIER) Then perso_dossier = perso_dossier & SEP_DOSSIER
        
    'fichier log perso avec entêtes
    If (Trim(perso_fichier) = "" Or Trim(perso_heads) = "") Then GoTo fn_initialisation_perso_end
        
    s_fichier = perso_dossier & perso_fichier
        
    'verification fichier
    If (Dir(s_fichier, vbNormal) = "") Then Call mt_textfile_create(perso_dossier & perso_fichier, perso_heads)
        
    mt_archive_textfile (s_fichier)

    'validation
    b_res = True
    
fn_initialisation_perso_end:
    'clean l'erreur
    Call Err.Clear
    
    'return
    fn_initialisation_perso = b_res
    
End Function '--- fn_initialisation_perso ---'

'-------------------------------------------------------------------------------------
'#description
'méthode qui gère la création d'un fichier type texte
'-------------------------------------------------------------------------------------
Private Sub mt_textfile_create(ByVal p_path As String, ByVal p_entetes As String)

    'déclarations
    Dim TextFile As Integer

    'détermine le numéro de fichier texte à utiliser
    TextFile = FreeFile

    'création du fichier
    Open p_path For Output As TextFile

    'ajoute les entêtes
    Print #TextFile, p_entetes

    'sauvegarde et ferme le fichier
    Close TextFile

End Sub '--- mt_textfile_create ---'

'-------------------------------------------------------------------------------------
'#description
'méthode qui gère l'ajout de données à un fichier type texte
'-------------------------------------------------------------------------------------
Private Sub mt_textfile_append(ByVal p_path As String, ByVal p_msg As String)

    'déclarations
    Dim TextFile As Integer

    'détermine le numéro de fichier texte à utiliser
    TextFile = FreeFile

    'ouverture du fichier
    Open p_path For Append As TextFile

    'ajoute le message
    Print #TextFile, p_msg

    'sauvegarde et ferme le fichier
    Close TextFile

End Sub '--- mt_textfile_append ---'

Private Sub mt_archive_textfile(ByVal p_fichier As String)

    ' Archive file at certain size
    If FileLen(p_fichier) > SIZE_ARCHIVE Then
        FileCopy p_fichier, Replace(p_fichier, ".lfo", Format(Now, "ddmmyyyy hhmmss.lfo"))
        Kill p_fichier
    End If
End Sub

