Attribute VB_Name = "modCommonFunctions"
#Const FinalMode = 1
Option Explicit
'***********************
'* Win32 Constants . . .
'***********************

Private Const INFINITE As Long = &HFFFF
Private Const TH32CS_SNAPPROCESS As Long = 2&
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const WINAPI_TRUE = 1
Private Const PROCESS_TERMINATE = 1
Private Const CREATE_SUSPENDED As Long = &H4

Private Const STARTF_USESHOWWINDOW = &H1
Private Enum enSW
SW_HIDE = 0
SW_NORMAL = 1
SW_MAXIMIZE = 3
SW_MINIMIZE = 6
End Enum

Private Type PROCESS_INFORMATION
hProcess As Long
hThread As Long
dwProcessId As Long
dwThreadId As Long
End Type

Private Type STARTUPINFO
cb As Long
lpReserved As Long
lpDesktop As Long
lpTitle As Long
dwX As Long
dwY As Long
dwXSize As Long
dwYSize As Long
dwXCountChars As Long
dwYCountChars As Long
dwFillAttribute As Long
dwFlags As Long
wShowWindow As Integer
cbReserved2 As Integer
lpReserved2 As Byte
hStdInput As Long
hStdOutput As Long
hStdError As Long
End Type

Private Type SECURITY_ATTRIBUTES
nLength As Long
lpSecurityDescriptor As Long
bInheritHandle As Long
End Type

Private Enum enPriority_Class
NORMAL_PRIORITY_CLASS = &H20
IDLE_PRIORITY_CLASS = &H40
HIGH_PRIORITY_CLASS = &H80
End Enum

Private Const GW_HWNDFIRST& = 0
Private Const HWND_NOTOPMOST& = -2
Private Const HWND_TOPMOST& = -1
Private Const SWP_NOMOVE& = &H2
Private Const SWP_NOSIZE& = &H1

Private Const PROCESS_VM_READ = (&H10)
Private Const PROCESS_VM_WRITE = (&H20)
Private Const PROCESS_VM_OPERATION = (&H8)
Private Const PROCESS_QUERY_INFORMATION = (&H400)
Private Const PROCESS_READ_WRITE_QUERY = PROCESS_VM_READ + PROCESS_VM_WRITE + PROCESS_VM_OPERATION + PROCESS_QUERY_INFORMATION






      'constants required by Shell_NotifyIcon API call:
      Public Const NIM_ADD = &H0
      Public Const NIM_MODIFY = &H1
      Public Const NIM_DELETE = &H2
      Public Const NIF_MESSAGE = &H1
      Public Const NIF_ICON = &H2
      Public Const NIF_TIP = &H4
      Public Const WM_MOUSEMOVE = &H200
      Public Const WM_LBUTTONDOWN = &H201     'Button down
      Public Const WM_LBUTTONUP = &H202       'Button up
      Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
      Public Const WM_RBUTTONDOWN = &H204     'Button down
      Public Const WM_RBUTTONUP = &H205       'Button up
      Public Const WM_RBUTTONDBLCLK = &H206   'Double-click
      
#If Win32 Then
  Public Declare Function GetTickCount Lib "Kernel32" () As Long
#Else
  Public Declare Function GetTickCount Lib "user" () As Long
#End If

Public Const ProxyVersion = "41.8" ' Equivalent Blackd Proxy version
Public Const myNumericVersion = 41800 ' Equivalent Blackd Proxy numeric version
Public Const SafeVersion = "2.3.4" ' BSC version
Public Const myNumericSafeVersion = 233 ' BSC numeric version
Public Const myAuthProtocol = 2 ' authetication protocol - NOT USED at this moment

' authentication key - not used at this moment
Public Const longsecretkey = "pfiwmvjgjikdfzasdruieopqwfhgkvvbnmklpofufrhufhuhsqaewftswgyguuhbvxhchufudhgoipopeqwiueifhjhsfdzvvcdvhfhfruyiurtuiuwfewqweffswqdepoffr"
Public blnConnected As Boolean
Public webReceived As String

Public InvalidateIt As Long
Public BpUserEmail As String
Public CurrBlackdServer As String
Public CurrBlackdServer_folder As String
Public cteLoginServerIP01 As String
Public cteLoginServerIP02 As String
Public cteLoginServerIP03 As String
Public cteLoginServerIP04 As String
Public cteLoginServerIP05 As String
Public cteLoginServerIP06 As String

Public cteLoginServerIP07 As String
Public cteLoginServerIP08 As String
Public cteLoginServerIP09 As String
Public cteLoginServerIP10 As String
Public cteLoginServerIP11 As String
Public cteLoginServerIP12 As String
Public cteLoginServerIP13 As String
Public cteLoginServerIP14 As String
Public cteLoginServerIP15 As String
Public cteLoginServerIP16 As String


Public cte_MIRROR1_folder As String
Public cte_MIRROR2_folder As String
Public cte_MIRROR1 As String
Public cte_MIRROR2 As String

Public TibiaVersionLong As Long
Public TibiaVersion As String
Public tibiaclassname As String
Public DefaultTibiaFolder As String
Public OverwriteTibiaExePath As String

Public XRAY_floors_ABOVE As Long
Public XRAY_floors_BELOW As Long
Public XRAY_key1_1 As Long
Public XRAY_key1_2 As Long
Public XRAY_key2_1 As Long
Public XRAY_key2_2 As Long
Public XRAY_key3_1 As Long
Public XRAY_key3_2 As Long


Public LAST_BATTLELISTPOS As Long

Public MustUnload As Boolean
Public nid As NOTIFYICONDATA
      Public Type NOTIFYICONDATA
       cbSize As Long
       hWnd As Long
       uId As Long
       uFlags As Long
       uCallBackMessage As Long
       hIcon As Long
       szTip As String * 64
      End Type
      

Public Now_we_define As String
Public HPmanadelay1 As Long
Public HPmanadelay2 As Long
Public HPrandpercent As Long


Public LEVELSPY_NOP As Long
Public LEVELSPY_ABOVE As Long
Public LEVELSPY_BELOW As Long
Public LIGHT_NOP As Long
Public LIGHT_AMOUNT As Long



Public PLAYER_Z As Long

Public MAP_POINTER_ADDR As Long
Public OFFSET_POINTER_ADDR As Long ' +1C

Public adrNChar As Long
Public CharDist As Long
Public NameDist As Long
Public OutfitDist As Long
Public adrNum As Long

Public adrXOR As Long
Public adrMyHP As Long
Public adrMyMaxHP As Long
Public adrMyMana As Long
Public adrMyMaxMana As Long
Public adrMySoul As Long

Public MyHP As Long
Public MyMaxHP As Long
Public MyMana As Long
Public MyMaxMana As Long
Public MySoul As Long
Public MyHPpercent As Long
Public Mymanapercent As Long
    
Public adrConnected As Long



Public HPmanaLimit0 As Long
Public HPmanaLimit1  As Long
Public HPmanaLimit2 As Long
Public HPmanaLimit3 As Long
Public HPmanaLimit4 As Long
Public HPmanaLimit5 As Long
Public HPmanaAction0 As String
Public HPmanaAction1 As String
Public HPmanaAction2 As String
Public HPmanaAction3 As String
Public HPmanaAction4 As String
Public HPmanaAction5 As String

Public Message_Tittle As String
Public Message_Message As String
Public LanguageFile As String
Public TibiaIsConnected As Boolean



Public LastWinsockError As String

Public LightDist As Long
Public LightColourDist As Long

Public LightColour As Byte
Public LightIntensity As Byte
Public LightEnabled As Boolean

Public LightRefreshDelay As Long
Public TibiaExePath As String




Public tileID_Blank As Long
Public tileID_WallBugItem As Long
Public tileID_SD As Long
Public tileID_HMM As Long
Public tileID_Explosion As Long
Public tileID_IH As Long
Public tileID_UH As Long

Public tileID_fireball As Long
Public tileID_stalagmite As Long
Public tileID_icicle As Long

'items
Public tileID_Bag As Long
Public tileID_Backpack As Long
Public tileID_Oracle As Long
Public tileID_FishingRod As Long
Public tileID_Rope As Long
Public tileID_LightRope As Long
Public tileID_Shovel As Long
Public tileID_LightShovel As Long

'water
Public tileID_waterEmpty As Long
Public tileID_waterWithFish As Long
Public tileID_waterEmptyEnd As Long
Public tileID_waterWithFishEnd As Long

Public TimesWarnedAboutRelog As Long

' blocking objects
Public tileID_blockingBox As Long

' to up floor
Public tileID_rampToNorth As Long
Public tileID_rampToSouth As Long
Public tileID_ladderToUp As Long
Public tileID_holeInCelling As Long
Public tileID_stairsToUp As Long
Public tileID_woodenStairstoUp As Long

Public tileID_desertRamptoUp As Long

Public tileID_rampToRightCycMountain As Long
Public tileID_rampToLeftCycMountain As Long

Public tileID_jungleStairsToNorth As Long
Public tileID_jungleStairsToLeft As Long


' to down
Public tileID_grassCouldBeHole As Long
Public tileID_pitfall As Long
Public tileID_openHole As Long
Public tileID_openHole2 As Long
Public tileID_trapdoor As Long
Public tileID_trapdoor2 As Long
Public tileID_sewerGate As Long
Public tileID_stairsToDown As Long
Public tileID_stairsToDown2 As Long
Public tileID_woodenStairstoDown As Long
Public tileID_rampToDown As Long
Public tileID_closedHole As Long
Public tileID_desertLooseStonePile As Long
Public tileID_OpenDesertLooseStonePile As Long
Public tileID_trapdoorKazordoon As Long
Public tileID_stairsToDownKazordoon As Long
Public tileID_stairsToDownThais As Long
Public tileID_down1 As Long
Public tileID_down2 As Long
Public tileID_down3 As Long

'FOOD
Public tileID_firstFoodTileID As Long
Public tileID_lastFoodTileID As Long
Public tileID_firstMushroomTileID As Long
Public tileID_lastMushroomTileID As Long


'FIELD RANGE1
Public tileID_firstFieldRangeStart As Long
Public tileID_firstFieldRangeEnd As Long
Public tileID_secondFieldRangeStart As Long
Public tileID_secondFieldRangeEnd As Long

Public tileID_campFire1 As Long
Public tileID_campFire2 As Long

'WALKABLE FIELDS
Public tileID_walkableFire1 As Long
Public tileID_walkableFire2 As Long
Public tileID_walkableFire3 As Long

'inside depot item
Public tileID_depotChest As Long

' flasks mana
Public tileID_flask As Long

Public tileID_health_potion As Long
Public tileID_strong_health_potion As Long
Public tileID_small_health_potion As Long
Public tileID_great_health_potion As Long
Public tileID_mana_potion As Long
Public tileID_strong_mana_potion As Long
Public tileID_great_mana_potion As Long

Public tileID_ultimate_health_potion As Long
Public tileID_great_spirit_potion As Long

Public byteNothing As Byte
Public byteMana As Byte
Public byteLife As Byte

Public blank1 As Byte
Public blank2 As Byte


Public TilesAvailable As Boolean

Public firstValidOutfit As Long
Public lastValidOutfit As Long

Public useDynamicOffsetBool As Boolean

Public TIBIA_LASTPID As Long
Public TIBIA_LASTOFFSET As Long
Public TIBIA_LASTBASE As Long

            Public Declare Function SetForegroundWindow Lib "user32" _
      (ByVal hWnd As Long) As Long
      Public Declare Function Shell_NotifyIcon Lib "shell32" _
      Alias "Shell_NotifyIconA" _
      (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
      Public Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As _
String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "Kernel32" Alias _
"WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As _
Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

'Find a child window with a given class name and caption
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

' get windows with current focus
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function OpenProcess Lib "Kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WriteProcessMemory Lib "Kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function ReadProcessMemory Lib "Kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long




Public Sub LogStatusOnFile(file_name As String)
  Dim fn As Integer
  Dim a As Long
  Dim writeThis As String
  On Error GoTo ignoreit
  a = 0
  fn = FreeFile
  Open App.Path & "\" & file_name For Append As #fn
    writeThis = vbCrLf & "ADITIONAL DETAILS:" & vbCrLf
    Print #fn, writeThis
    writeThis = "TibiaVersionLong=" & CStr(TibiaVersionLong)
    Print #fn, writeThis
    writeThis = "TibiaVersion=" & TibiaVersion
  Close #fn
  Exit Sub
ignoreit:
  a = -1
End Sub
Public Sub OverwriteOnFile(file_name As String, strText As String)
  Dim fn As Integer
  Dim errheader As String
  Dim writeThis As String
  Dim a As Long
  On Error GoTo ignoreit
  a = 0
  fn = FreeFile
  If file_name = "errors.txt" Then
    errheader = "[" & Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss") & " using version " & SafeVersion & " , with config.int v" & CStr(TibiaVersionLong) & " ] "
    writeThis = errheader & strText
  Else
    writeThis = strText
  End If
  Open App.Path & "\" & file_name For Output As #fn
    Print #fn, writeThis
  Close #fn
  If file_name = "errors.txt" Then
    LogStatusOnFile "errors.txt"
    'frmMenu.Caption = "ERROR - Check errors.txt for details"
  End If
  Exit Sub
ignoreit:
  a = -1
End Sub
Public Sub OverwriteOnFileSimple(file_name As String, strText As String)
  Dim fn As Integer
  Dim writeThis As String
  Dim a As Long
  On Error GoTo ignoreit
  a = 0
  fn = FreeFile
    writeThis = strText
  Open App.Path & "\" & file_name For Output As #fn
    Print #fn, writeThis
  Close #fn

  Exit Sub
ignoreit:
  a = -1
End Sub
Public Sub LogOnFile(file_name As String, strText As String)
  Dim fn As Integer
  Dim errheader As String
  Dim writeThis As String
  Dim a As Long
  On Error GoTo ignoreit
  a = 0
  fn = FreeFile
  If file_name = "errors.txt" Then
    errheader = "[" & Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss") & " using version " & SafeVersion & " , with config.int v" & CStr(TibiaVersionLong) & " ] "
    writeThis = errheader & strText
   ' frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & writeThis
  Else
    writeThis = strText
  End If
  If Len(file_name) > 4 Then
    If left$(file_name, 4) = "log_" Then
      Open App.Path & "\mylogs\" & file_name For Append As #fn
    Else
      Open App.Path & "\" & file_name For Append As #fn
    End If
  Else
  Open App.Path & "\" & file_name For Append As #fn
  End If
    Print #fn, writeThis
  Close #fn

  Exit Sub
ignoreit:
  a = -1
End Sub
Public Function encriptionSumChr(ByVal chr1 As String, ByVal chr2 As String, Optional dosum As Boolean = True) As String
    Dim asc1 As Long
    Dim asc2 As Long
    Dim newchr3 As String
    Dim res As Long
    Dim i As Long
    asc1 = CLng(AscB(chr1))
    asc2 = CLng(AscB(chr2))
    If dosum = True Then
        res = asc1
        For i = 1 To asc2
            res = res + 1
            If res > 127 Then
                res = 0
            End If
        Next i
    Else
        res = asc1
        For i = 1 To asc2
            res = res - 1
            If res < 0 Then
                res = 127
            End If
        Next i
    End If
    newchr3 = Chr$(res)
    encriptionSumChr = newchr3
End Function


Public Function stringOut2(str As String) As String
  Dim s As String
  Dim l As Long
  Dim i As Long
  Dim tmp1 As String
  Dim tmp2 As Byte
  l = Len(str) / 3
  s = ""
  For i = 1 To l
    tmp1 = Mid(str, -2 + (i * 3), 3)
    tmp2 = CByte(CLng(tmp1))
    s = s & Chr(tmp2)
  Next i
  stringOut2 = s
End Function

Public Function stringout3(str As String) As String
    Dim randomDesp As Long
    Dim lenstr As String
    Dim cranchar As String
    Dim finalPass As String
    Dim i As Long
    Dim res As String
    randomDesp = CLng(left$(str, 1))
    lenstr = Len(str)
    finalPass = ""
    For i = 2 To lenstr
       cranchar = Mid$(longsecretkey, randomDesp + i - 1, 1)
       finalPass = finalPass & encriptionSumChr(Mid$(str, i, 1), cranchar, False)
    Next i
    res = finalPass
    res = stringOut2(finalPass)
    stringout3 = res
End Function

Public Function GoodHex(B As Byte) As String
  Dim res As String
  res = Hex(B)
  If Len(res) = 1 Then
    GoodHex = "0" & res 'add a zero if VB conversion only return 1 character
  Else
    GoodHex = res
  End If
End Function

Public Function SafeLong(strThing As String) As Long
    Dim res As Long
    
    On Error GoTo goterr
    res = CLng(strThing)
    SafeLong = res
    Exit Function
goterr:
    SafeLong = 0
End Function

Public Function getfromINI(ByRef par1 As String, ByRef par2 As String, _
 ByRef par3 As String, ByRef par4 As String, ByRef par5 As Long, ByRef par6 As String) As Long
    getfromINI = GetPrivateProfileString(par1, par2, par3, par4, par5, par6)
End Function

Public Function setToINI(ByRef par1 As String, ByRef par2 As String, _
 ByRef par3 As String, ByRef par4 As String)
    setToINI = WritePrivateProfileString(par1, par2, par3, par4)
End Function

Public Function GetFocusedTibiaPID() As Long
    Dim res As Boolean
    Dim pIDfocusedWindow As Long
    Dim tibiaclient As Long
    
    pIDfocusedWindow = GetForegroundWindow()
    tibiaclient = 0
    Do
        tibiaclient = FindWindowEx(0, tibiaclient, tibiaclassname, vbNullString)
        If tibiaclient = 0 Then
          Exit Do
        Else
            If pIDfocusedWindow = tibiaclient Then
                GetFocusedTibiaPID = tibiaclient
                Exit Function
            End If
        End If
    Loop
    GetFocusedTibiaPID = 0
End Function


Public Function GetFirstTibiaPID() As Long
    Dim tibiaclient As Long
    tibiaclient = 0
    tibiaclient = FindWindowEx(0, tibiaclient, tibiaclassname, vbNullString)
    GetFirstTibiaPID = tibiaclient
End Function

Public Function Memory_ReadByte(ByVal Address As Long, process_Hwnd As Long) As Byte
  
   ' Declare some variables we need
   Dim pid As Long         ' Used to hold the Process Id
   Dim phandle As Long     ' Holds the Process Handle
   Dim valbuffer As Byte   ' Byte
    
    

  
   ' First get a handle to the "game" window
   If (process_Hwnd = 0) Then Exit Function
   
   


   
   ' We can now get the pid
   GetWindowThreadProcessId process_Hwnd, pid
   
   
   ' Use the pid to get a Process Handle
   'phandle = OpenProcess(PROCESS_VM_READ, False, pid)
   phandle = OpenProcess(PROCESS_READ_WRITE_QUERY, False, pid)
   If (phandle = 0) Then Exit Function
   
   If useDynamicOffsetBool = True Then
     If process_Hwnd <> TIBIA_LASTPID Then
        TIBIA_LASTPID = process_Hwnd
        TIBIA_LASTBASE = getProcessBase(phandle, tibiaModuleRegionSize, False)
        If TIBIA_LASTBASE = 0 Then
          Debug.Print "Address Error"
          TIBIA_LASTBASE = &H400000
        End If
        TIBIA_LASTOFFSET = TIBIA_LASTBASE - &H400000
     End If
     Address = Address + TIBIA_LASTOFFSET
   End If
   
   ' Read Long
   ReadProcessMemory phandle, Address, valbuffer, 1, 0&
       
   ' Return
   Memory_ReadByte = valbuffer
   
   ' Close the Process Handle
   CloseHandle phandle
  
End Function
Public Sub Memory_WriteByte(ByVal Address As Long, valbuffer As Byte, process_Hwnd As Long)

   'Declare some variables we need
   Dim pid As Long         ' Used to hold the Process Id
   Dim phandle As Long     ' Holds the Process Handle
   
   ' First get a handle to the "game" window
   If (process_Hwnd = 0) Then Exit Sub
   
   ' We can now get the pid
   GetWindowThreadProcessId process_Hwnd, pid
   
   
   ' Use the pid to get a Process Handle
   phandle = OpenProcess(PROCESS_READ_WRITE_QUERY, False, pid)
   If (phandle = 0) Then Exit Sub
   
   
   
   If useDynamicOffsetBool = True Then
     If process_Hwnd <> TIBIA_LASTPID Then
        TIBIA_LASTPID = process_Hwnd
        TIBIA_LASTBASE = getProcessBase(phandle, tibiaModuleRegionSize, False)
        If TIBIA_LASTBASE = 0 Then
          Debug.Print "Address Error"
          TIBIA_LASTBASE = &H400000
        End If
        TIBIA_LASTOFFSET = TIBIA_LASTBASE - &H400000
     End If
     Address = Address + TIBIA_LASTOFFSET
   End If
   
   
   ' Write Long
   WriteProcessMemory phandle, Address, valbuffer, 1, 0&
   
   ' Close the Process Handle
   CloseHandle phandle

End Sub

Public Sub Memory_WriteLong(ByVal Address As Long, valbuffer As Long, process_Hwnd As Long)

   'Declare some variables we need
   Dim pid As Long         ' Used to hold the Process Id
   Dim phandle As Long     ' Holds the Process Handle
   
   ' First get a handle to the "game" window
   If (process_Hwnd = 0) Then Exit Sub
   
   ' We can now get the pid
   GetWindowThreadProcessId process_Hwnd, pid
   
   
   ' Use the pid to get a Process Handle
   phandle = OpenProcess(PROCESS_READ_WRITE_QUERY, False, pid)
   If (phandle = 0) Then Exit Sub
   
   
   
   If useDynamicOffsetBool = True Then
     If process_Hwnd <> TIBIA_LASTPID Then
        TIBIA_LASTPID = process_Hwnd
        TIBIA_LASTBASE = getProcessBase(phandle, tibiaModuleRegionSize, False)
        If TIBIA_LASTBASE = 0 Then
          Debug.Print "Address Error"
          TIBIA_LASTBASE = &H400000
        End If
        TIBIA_LASTOFFSET = TIBIA_LASTBASE - &H400000
     End If
     Address = Address + TIBIA_LASTOFFSET
   End If
   
   
   ' Write Long
   WriteProcessMemory phandle, Address, valbuffer, 2, 0&
   
   ' Close the Process Handle
   CloseHandle phandle

End Sub


Public Function Memory_ReadLong(ByVal Address As Long, process_Hwnd As Long) As Long
  
   ' Declare some variables we need
   Dim pid As Long         ' Used to hold the Process Id
   Dim phandle As Long     ' Holds the Process Handle
   Dim valbuffer As Long   ' Long
    
   ' First get a handle to the "game" window
   If (process_Hwnd = 0) Then Exit Function
   
   ' We can now get the pid
   GetWindowThreadProcessId process_Hwnd, pid
   
   ' Use the pid to get a Process Handle
   'phandle = OpenProcess(PROCESS_VM_READ, False, pid)
   phandle = OpenProcess(PROCESS_READ_WRITE_QUERY, False, pid)
   If (phandle = 0) Then Exit Function
   
   If useDynamicOffsetBool = True Then
     If process_Hwnd <> TIBIA_LASTPID Then
        TIBIA_LASTPID = process_Hwnd
        TIBIA_LASTBASE = getProcessBase(phandle, tibiaModuleRegionSize, False)
        If TIBIA_LASTBASE = 0 Then
          Debug.Print "Address Error"
          TIBIA_LASTBASE = &H400000
        End If
        TIBIA_LASTOFFSET = TIBIA_LASTBASE - &H400000
     End If
     Address = Address + TIBIA_LASTOFFSET
   End If
   
   ' Read Long
   ReadProcessMemory phandle, Address, valbuffer, 4, 0&
       
   ' Return
   Memory_ReadLong = valbuffer
   
   ' Close the Process Handle
   CloseHandle phandle
  
End Function

Public Function MyBattleListPositionByPID(tibiaclient As Long) As Long
  Dim c1 As Long
  Dim id As Double
  Dim res As Long
  Dim myID As Double
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  res = -1
  myID = CDbl(Memory_ReadLong(adrNum, tibiaclient))
  For c1 = 0 To LAST_BATTLELISTPOS
        id = CDbl(Memory_ReadLong(adrNChar + (CharDist * c1), tibiaclient))
        If myID = id Then
          res = c1
          Exit For
        End If
  Next c1
  MyBattleListPositionByPID = res
  Exit Function
goterr:
  MsgBox "Critical error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "MyBattleListPositionByPID"
  End
End Function

Public Function GetNameFromID(tibiaclient As Long, dblID As Double) As String
    Dim res As String
    Dim pos As Long
    res = ""
    pos = BattleListPositionOfID(tibiaclient, dblID)
    If pos = -1 Then
        res = "*Unknown* (ID " & dblID & ")"
    Else
        res = GetNameFromBattleListPos(tibiaclient, pos)
    End If
    GetNameFromID = res
End Function
Public Function BattleListPositionOfID(tibiaclient As Long, dblID As Double) As Long
  Dim c1 As Long
  Dim id As Double
  Dim res As Long
  Dim myID As Double
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  res = -1
  myID = dblID
  For c1 = 0 To LAST_BATTLELISTPOS
        id = CDbl(Memory_ReadLong(adrNChar + (CharDist * c1), tibiaclient))
        If myID = id Then
          res = c1
          Exit For
        End If
  Next c1
  BattleListPositionOfID = res
  Exit Function
goterr:
  MsgBox "Critical error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "MyBattleListPositionByPID"
  End
End Function
' debug.print GetNameFromBattleListPos(GetFirstTibiaPID(),0)
Public Function GetNameFromBattleListPos(ByVal tibiaclient As Long, _
    ByVal bPos As Long)
    Dim i As Long
    Dim resb As Byte
    Dim res As String
    res = ""
    For i = 4 To 54
        resb = GetElementFromBattleListPos(tibiaclient, bPos, i)
        If resb = 0 Then
            Exit For
        Else
            res = res & Chr(resb)
        End If
    Next i
    GetNameFromBattleListPos = res
End Function
Public Function GetElementFromBattleListPos(ByVal tibiaclient As Long, _
 ByVal bPos As Long, ByVal bElement As Long) As Byte
    Dim res As Byte
    #If FinalMode Then
    On Error GoTo goterr
    #End If
    res = Memory_ReadByte((adrNChar + (bPos * CharDist) + bElement), tibiaclient)
    GetElementFromBattleListPos = res
  Exit Function
goterr:
  MsgBox "Critical error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "GetElementFromBattleListPos"
  End
End Function

Public Function MaxV(ByVal v1 As Long, ByVal v2 As Long) As Long
    If v1 >= v2 Then
        MaxV = v1
    Else
        MaxV = v2
    End If
End Function

Public Function HighByteOfLong(Address As Long) As Byte
On Error GoTo goterr
  Dim h As Byte
  h = CByte(Address \ 256) ' high byte
  HighByteOfLong = h
goterr:
  HighByteOfLong = &H0
End Function

Public Function LowByteOfLong(Address As Long) As Byte
On Error GoTo goterr
  Dim h As Byte
  Dim l As Byte
  h = CByte(Address \ 256)
  l = CByte(Address - (CLng(h) * 256)) ' low byte
  LowByteOfLong = l
  Exit Function
goterr:
  LowByteOfLong = &H0
End Function

Public Function GetTheLong(byte1 As Byte, byte2 As Byte) As Long
  'get the long from 2 consecutive bytes in a tibia packet
  Dim res As Long
  res = CLng(byte2) * 256 + CLng(byte1)
  GetTheLong = res
End Function



Public Function MyFileExists(FileName As String) As Boolean
    On Error GoTo ErrorHandler
    ' get the attributes and ensure that it isn't a directory
    MyFileExists = (GetAttr(FileName) And vbDirectory) = 0
ErrorHandler:
    ' if an error occurs, this function returns False
End Function


Public Sub ReadTileIDListFromIni(ByRef thing() As Long, ByRef name As String, ByRef here As String, ByRef defaultV As String)
  ' read a tileID from ini
  Dim strInfo As String
  Dim i As Integer

  strInfo = String$(255, 0)
  i = GetPrivateProfileString("tileIDs", name, "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = left(strInfo, i)
    FillTheListFromString thing, strInfo
  Else
    FillTheListFromString thing, defaultV
  End If
End Sub

Public Sub FillTheListFromString(ByRef theList() As Long, ByRef theString As String)
  Dim remainingString As String
  Dim aTile As String
  Dim pos As Long
  Dim lonS As Long
  Dim listPos As Long
  Dim currChar As String
  On Error GoTo letsIgnoreIt
  lonS = Len(theString)
  pos = 1
  listPos = 0
  Do
    If pos > lonS Then
      theList(listPos) = 0
      Exit Do
    Else
      currChar = Mid$(theString, pos, 1)
      If (currChar = ",") Or (currChar = " ") Then
        pos = pos + 1
      Else
        If (pos + 5) <= (lonS + 1) Then
          aTile = Mid$(theString, pos, 5)
          theList(listPos) = GetTheLongFromFiveChr(aTile)
          listPos = listPos + 1
        End If
        pos = pos + 5
      End If
    End If
  Loop
  Exit Sub
letsIgnoreIt:
  theList(0) = 0
End Sub

Public Function GetTheLongFromFiveChr(str As String) As Long
  On Error GoTo goterr
  Dim b1 As Byte
  Dim b2 As Byte
  Dim b3 As Byte
  Dim b4 As Byte
  Dim res As Long
  res = -1
  
  If Len(str) > 4 Then
    b1 = FromHexToDec(Mid(str, 1, 1))
    b2 = FromHexToDec(Mid(str, 2, 1))
    b3 = FromHexToDec(Mid(str, 4, 1))
    b4 = FromHexToDec(Mid(str, 5, 1))
    res = (CLng(b2)) + (CLng(b1) * 16) + (CLng(b4) * 256) + (CLng(b3) * 4096)
  End If
  GetTheLongFromFiveChr = res
  Exit Function
goterr:
  GetTheLongFromFiveChr = -1 'new in 8.21 +
End Function


Public Function FromHexToDec(str As String) As Byte
  Dim res As Byte
  ' converts 1 character string
  ' to a byte
  res = 16 'reserved to error
  Select Case str
  Case "0"
    res = 0
  Case "1"
    res = 1
  Case "2"
    res = 2
  Case "3"
    res = 3
  Case "4"
    res = 4
  Case "5"
    res = 5
  Case "6"
    res = 6
  Case "7"
    res = 7
  Case "8"
    res = 8
  Case "9"
    res = 9
  Case "A", "a"
    res = 10
  Case "B", "b"
    res = 11
  Case "C", "c"
    res = 12
  Case "D", "d"
    res = 13
  Case "E", "e"
    res = 14
  Case "F", "f"
    res = 15
  End Select
  FromHexToDec = res
End Function


Public Sub ReadTileIDFromIni(ByRef thing As Long, ByRef name As String, ByRef here As String, ByRef defaultV As String)
  ' read a tileID from ini
  Dim strInfo As String
  Dim lonInfo As Long
  Dim i As Integer
  strInfo = String$(50, 0)
  i = GetPrivateProfileString("tileIDs", name, "", strInfo, Len(strInfo), here)
  If i > 0 Then
    strInfo = left(strInfo, i)
    lonInfo = GetTheLongFromFiveChr(strInfo)
    thing = lonInfo
  Else
    thing = GetTheLongFromFiveChr(defaultV)
  End If
End Sub


Public Function TibiaDatExists() As Boolean
    On Error GoTo goterr
  Dim tibiadathere As String
  tibiadathere = TibiaExePath & "tibia.dat"
  If MyFileExists(tibiadathere) = False Then
    TibiaDatExists = False
    Exit Function
  End If
  TibiaDatExists = True
  Exit Function
goterr:
  DBGtileError = "Error number = " & CStr(Err.Number) & vbCrLf & " ; Error description = " & Err.Description & " ; Path = " & tibiadathere
  TibiaDatExists = False
End Function
