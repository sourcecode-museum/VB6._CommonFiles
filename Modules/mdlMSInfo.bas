Attribute VB_Name = "mdlMSInfo"
Option Explicit

' Reg Key Security Options...
Private Const READ_CONTROL = &H20000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const ERROR_SUCCESS = 0
Private Const REG_SZ = 1                         ' Unicode nul terminated string
Private Const REG_DWORD = 4                      ' 32-bit number

Private Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Private Const gREGVALSYSINFOLOC = "MSINFO"
Private Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Private Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Private Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
   Dim i As Long              ' Loop Counter
   Dim rc As Long             ' Return Code
   Dim hKey As Long           ' Handle To An Open Registry Key
   Dim hDepth As Long
   Dim KeyValType As Long     ' Data Type Of A Registry Key
   Dim tmpVal As String       ' Tempory Storage For A Registry Key Value
   Dim KeyValSize As Long     ' Size Of Registry Key Variable
   
   '------------------------------------------------------------
   ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
   '------------------------------------------------------------
   
   ' Open Registry Key
   rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)
    
   If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
   
   tmpVal = String$(1024, 0)                             ' Allocate Variable Space
   KeyValSize = 1024                                       ' Mark Variable Size
   
   '------------------------------------------------------------
   ' Retrieve Registry Key Value...
   '------------------------------------------------------------
   rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                        KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
   If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
   If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
       tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
   Else                                                    ' WinNT Does NOT Null Terminate String...
       tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
   End If
   
   '------------------------------------------------------------
   ' Determine Key Value Type For Conversion...
   '------------------------------------------------------------
   Select Case KeyValType                                  ' Search Data Types...
      Case REG_SZ                                             ' String Registry Key Data Type
         KeyVal = tmpVal                                     ' Copy String Value
      Case REG_DWORD                                          ' Double Word Registry Key Data Type
         For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
         Next
         KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
   End Select
    
   GetKeyValue = True                                      ' Return Success
   rc = RegCloseKey(hKey)                                  ' Close Registry Key
   Exit Function                                           ' Exit
   
GetKeyError:      ' Cleanup After An Error Has Occured...
   KeyVal = ""                                             ' Set Return Val To Empty String
   GetKeyValue = False                                     ' Return Failure
   rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Public Sub ShowMSInfo32()
   Dim rc As Long
   Dim SysInfoPath As String
   
   On Error GoTo SysInfoErr
     
   'Tentando Obter Informação do Caminho do Programa de Sistema e Nome do Registro. . .
   If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
   'Try To Get System Info Program Path Only From Registry...
   ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
      If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
          SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
      Else     'Error - File Can Not Be Found...
         GoTo SysInfoErr
      End If
   Else        'Error - Registry Entry Can Not Be Found...
      GoTo SysInfoErr
   End If
    
   Call Shell(SysInfoPath, vbNormalFocus)
    
   Exit Sub
   
SysInfoErr:
   MGShowInfo "Informação de Sistema não está disponível neste momento."
End Sub

Public Sub StatusMemoria(ByRef FisicaKB As Double, _
                        ByRef FisPercentual As Double, _
                        ByRef VirtualPercentual As Double)
   Dim VarMemoria As MEMORYSTATUS
   Dim nWidth As Integer
   
   VarMemoria.dwLength = Len(VarMemoria)
   GlobalMemoryStatus VarMemoria
   
   With VarMemoria
      FisicaKB = (.dwAvailPhys / 1024)
      FisPercentual = .dwAvailPhys / .dwTotalPhys
      VirtualPercentual = .dwAvailVirtual / .dwTotalVirtual
   End With
End Sub

