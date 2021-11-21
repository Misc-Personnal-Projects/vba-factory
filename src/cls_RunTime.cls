VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_RunTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
'https://stackoverflow.com/questions/198409/how-do-you-test-running-time-of-vba-code
'https://docs.microsoft.com/fr-ch/windows/win32/sysinfo/acquiring-high-resolution-time-stamps?redirectedfrom=MSDN
'----------------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit

'API WINDOWS
Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

#If VBA7 Then
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
#Else
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
#End If

Private m_CounterStart As LARGE_INTEGER
Private m_CounterEnd As LARGE_INTEGER
Private m_crFrequency As Double

Private Const TWO_32 = 4294967296# ' = 256# * 256# * 256# * 256#


'MASTER DICTIONNARY
Private m_Steps As Dictionary

'PROGRESS
Private m_Progression As usf_Progression
Private m_TotalSteps As Long
Private m_StepsCount As Long
Private m_ProcessName As String

'TIME:
Private m_LastTime As Double
Private m_SecondElapsed As Double
Private m_StepDuration As Double

'LOGS
Private m_Logs As cls_Logs

Public Property Get SecondElapsed() As Double: SecondElapsed = m_SecondElapsed: End Property
Public Property Get StepDuration() As Double: StepDuration = m_StepDuration: End Property

Private Sub Class_Initialize()
    Dim PerfFrequency As LARGE_INTEGER
    
    'PROGRESS :
    Set m_Steps = New Dictionary
    Set m_Progression = New usf_Progression
    m_Progression.ProgressionBar.ValMin = 0
    
    'TIME :
    QueryPerformanceFrequency PerfFrequency
    m_crFrequency = LI2Double(PerfFrequency)
    
    'LOGS :
    Set m_Logs = New cls_Logs
    m_Logs.Entetes = LOG_FILE_RUNTIME_HEADERS
    m_Logs.Dossier = ThisWorkbook.path & LOG_FOLDER
    m_Logs.Fichier = LOG_FILE_RUNTIME
End Sub

Private Sub Class_Terminate()

    'PROGRESS
    Call Unload(m_Progression)
    
    'TIME : Save steps values in dictionnary
    m_SecondElapsed = getSecondElapsed
    
    'LOGS :
    Call m_Logs.Log_Perso(Array(Format(Date, "dd.mm.yyyy"), Format(Time, "hh:mm:ss"), Environ("username"), ThisWorkbook.VBProject.name, UCase(m_ProcessName), "--------------------- END ----------------------------------------------------------------", m_SecondElapsed))

    Set m_Steps = Nothing
    Set m_Progression = Nothing
    Set m_Logs = Nothing

End Sub

Private Function LI2Double(LI As LARGE_INTEGER) As Double
    Dim Low As Double
    
    Low = LI.lowpart
    If Low < 0 Then
        Low = Low + TWO_32
    End If
    
    LI2Double = LI.highpart * TWO_32 + Low
End Function

Public Sub StartCounter(p_TotalSteps As Long, p_ProcessName As String)
    
    'PROGRESS :
    m_TotalSteps = p_TotalSteps
    m_ProcessName = p_ProcessName
    m_Progression.title_header = p_ProcessName
    m_Progression.ProgressionBar.ValMax = m_TotalSteps
    Call m_Progression.Show(vbModeless)
    
    'TIME : Save the start time
    QueryPerformanceCounter m_CounterStart
    
    'LOGS :
    Call m_Logs.Log_Perso(Array(Format(Date, "dd.mm.yyyy"), Format(Time, "hh:mm:ss"), Environ("username"), ThisWorkbook.VBProject.name, UCase(p_ProcessName), "--------------------- START ----------------------------------------------------------------"))

End Sub

Function TimeElapsed() As Double
    Dim crStart As Double
    Dim crStop As Double
    
    QueryPerformanceCounter m_CounterEnd
    crStart = LI2Double(m_CounterStart)
    crStop = LI2Double(m_CounterEnd)
    TimeElapsed = 1000# * (crStop - crStart) / m_crFrequency
End Function

Function SaveStep(p_StepName As String, p_Status As Boolean)

    Dim Status As String
    Status = IIf(p_Status, "SUCCESS", "FAILED")
    
    'PROGRESS : load userform to display progression
    m_StepsCount = m_Steps.Count
    m_Progression.lbl_message = p_StepName & "Steps :" & m_StepsCount & "/" & m_TotalSteps & "..."
    m_Progression.ProgressionBar.Val = m_StepsCount
    Call m_Progression.Repaint
    
    'TIME : Save steps values in dictionnary
    m_SecondElapsed = getSecondElapsed
    m_StepDuration = m_SecondElapsed - m_LastTime
    
    m_Steps.Add p_StepName, Array(m_SecondElapsed, m_StepDuration, Status)
    m_LastTime = m_Steps.Item(p_StepName)(0)
    
    'LOGS :
    Call m_Logs.Log_Perso(Array(Format(Date, "dd.mm.yyyy"), Format(Time, "hh:mm:ss"), Environ("username"), ThisWorkbook.VBProject.name, m_ProcessName, p_StepName, CStr(SecondElapsed), CStr(Round(m_StepDuration, 2)), CStr(Status)))


End Function

Function getSecondElapsed() As Double
    getSecondElapsed = Round(TimeElapsed / 1000, 2)
End Function

Function getRuntimeReport()
    
    Dim step As Variant
    
    For Each step In m_Steps.Keys
        Debug.Print "--------------------------------------"
        Debug.Print step
        Debug.Print "--------------------------------------"
        Debug.Print "Time elapsed : " & Round(m_Steps(step)(0), 2)
        Debug.Print "Step Duration : " & Round(m_Steps(step)(1), 2)
        Debug.Print "Status : " & m_Steps(step)(2)
        Debug.Print "--------------------------------------" & vbNewLine
    Next step
    
End Function


