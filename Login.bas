Attribute VB_Name = "Login"

Public Const CTF_COINIT = &H8
Public Const CTF_INSIST = &H1
Public Const CTF_PROCESS_REF = &H4
Public Const CTF_THREAD_REF = &H2

Declare Function SHCreateThread Lib "shlwapi.dll" (ByVal pfnThreadProc As Long, pData As Any, ByVal dwFlags As Long, ByVal pfnCallback As Long) As Long

Public StopThread As Boolean




Public Sub myThread()
    
    'This would normally lock the form
    Do While Not StopThread
    
        LoadState (1)
        'LoadForms
        
        'StopThread = True
        
    Loop
    
    
End Sub
