Attribute VB_Name = "mod_GlobalRef"
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
'   FunctionName
'   ProcedureName
'
'#Source :
'----------------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit

Public Const G_DEBUG As Boolean = False 'ATTENTION - Should be at false in prod

Public Const TYP_TXT As String = "TEXT"
Public Const TYP_XL As String = "EXCEL"

Public Const ACT_INSERT As String = "Ajouter" 'usf_Administration
Public Const ACT_MODIFY As String = "Modifier" 'usf_Administration
Public Const ACT_REMOVE As String = "Supprimer" 'usf_Administration
Public Const ACT_SELECT As String = "Sélectionner" 'usf_Administration
Public Const ACT_LIMITED As String = "Limiter" 'usf_Administration


'====================================================================================================================================================
'   ACTION TYPES (cls_Logs)
'====================================================================================================================================================
Public Enum Actions
    Ouvrir = 0
    Creer = 1
    Modifier = 2
    Annuler = 3
    Supprimer = 4
    Sauvegarder = 5
    Envoyer = 6
    Fermer = 7
    Sélectionner = 8
End Enum

'====================================================================================================================================================
'   TYPES & EXTENSION (cls_Database)
'====================================================================================================================================================
Public Enum e_DataSourceTypes
    txt = 0
    CSV = 1
    XLSX = 2
    xlsm = 3
    ACCDB = 4
End Enum


'====================================================================================================================================================
'   PARAMETERS INDEX
'====================================================================================================================================================
Public Enum e_ParameterIndex
    ParaName = 0
    ParaValue = 1
    ParaEditable = 2
End Enum

'====================================================================================================================================================
'   REFERENCES
'====================================================================================================================================================
Public Const REF_VERS As String = "ref_version"

'====================================================================================================================================================
'   FOLDERS
'====================================================================================================================================================
Public Const DATA_FOLDER As String = "\Data"
Public Const DATA_TEMPLATE_FOLDER As String = "\Data\Templates"
Public Const EXTRACTIONS_FOLDER As String = "\Data\Extractions"
Public Const SEP_FOLDER As String = "\"

'====================================================================================================================================================
'   LAYOUT
'====================================================================================================================================================

'usf_Administration
Public Const COLOR_FONT_INACTIVE As Long = 1973790
Public Const COLOR_FONT_ACTIVE As Long = 16119285

Public Const HEIGHT_ADMIN_ACTIVE As Double = 480
Public Const HEIGHT_ADMIN_INACTIVE As Double = 110
Public Const TOP_ADMIN_FRAME As Double = 100
Public Const HEIGHT_SEARCH_RH_FILTERS As Double = 210
Public Const HEIGHT_SEARCH_RH_RESULTS As Double = 460
Public Const HEIGHT_DETAIL_BASE As Double = 480
Public Const HEIGHT_SEARCH_BASE As Double = 200

'Msgbox_Perso
Public Const TYP_INFORMATION As String = "Information"
Public Const TYP_QUESTION As String = "Question"

'====================================================================================================================================================
'   COLORS
'====================================================================================================================================================
Public Const COLOR_WHITE As Long = 16777215
Public Const COLOR_BLACK As Long = 0

Public Const COLOR_VAR_CONTAINER As Long = 7895160
Public Const COLOR_VAR_ELEMENTS As Long = 15000804

Public Const COLOR_VIEW_ARTICLES As Long = 3760128
Public Const COLOR_VIEW_EXT As Long = 6969932
Public Const COLOR_VIEW_INT As Long = 2916003

'====================================================================================================================================================
'   GUEST
'====================================================================================================================================================
Public Const GUEST_FULLNAME_TXT As String = "Invité"
Public Const ANONYM_PHOTO As String = "img_silhouette.gif"
Public Const ANONYM_FIELD_TXT As String = "xxxx"
Public Const ANONYM_FIELD_VAL As Long = 0

'====================================================================================================================================================
'   ROLES
'====================================================================================================================================================
Public Const ADMINISTRATEUR As Integer = 7 'usf_DetailsUser
Public Const DEMANDEUR As Integer = 1 'cls_UserList

'====================================================================================================================================================
'   FEATURES
'====================================================================================================================================================
Public Const ADMINISTRER_LES_DROITS_DUTILISATION As String = 1

'====================================================================================================================================================
'   SIMPLE DIM
'====================================================================================================================================================
Public Const ELT_ROLE As String = "Rôle"
Public Const ELT_RISK As String = "Risque"
Public Const ELT_PROJET As String = "Projet"

'====================================================================================================================================================
'   VARIABLES
'====================================================================================================================================================
Public g_activeIndex As Variant 'usf_Administration

'usf_DatePicker
Public g_date As Date

'usf_Feedback
Public g_transfert As Variant
Public g_action As String
Public g_source As String

'main worksheet
Public g_line_number As Long

'SimpleDim
Public g_tmpDim As cls_SimpleDim 'usf_Administration
Public g_tmpDim_list As cls_SimpleDimList 'usf_Administration

'User
Public g_active_user As cls_User 'RefreshDataIntranet
Public g_roles As cls_SimpleDimList 'usf_DetailsUser
Public g_TmpUser As cls_User 'usf_Administration
Public g_Users As cls_UsersList 'usf_Administration



