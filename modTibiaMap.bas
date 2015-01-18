Attribute VB_Name = "modTibiaMap"
Option Explicit
#Const FinalMode = 1
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



Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function OpenProcess Lib "Kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WriteProcessMemory Lib "Kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function ReadProcessMemory Lib "Kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long


'Public Const MAP_POINTER_ADDR As Long = &H64A048
Public Const MAX_ITEMS_PER_TILE As Long = 10


' (size of 1 tibia map square info = 168)
Private Const SizeOfTibiaMap As Long = 338688 '  168*14*18*8  =   168*2016
'Private Const SizeOfTibiaMapOld As Long = 258048 '  128*14*18*8  =   128*2016
Private Const SizeOfTibiaMap3 As Long = 661248 '  328*14*18*8  =   328*2016
Private Const SizeOfTibiaMap4 As Long = 822528 '  408*14*18*8  =   408*2016
Private Const SizeOfTibiaMap5 As Long = 741888 '  368*14*18*8  =   368*2016

Private Const SizeOfTibiaMapOld As Long = 346752 '  128*14*18*8  =   172*2016
Public Type MapItemOld
    id As Long 'tile id
    data1 As Long ' creature/player id
    data2 As Long
End Type
Public Type MapItem
    id As Long 'tile id
    data1 As Long ' creature/player id
    data2 As Long
End Type
Public Type MapItem2 '12
    data1 As Long ' creature/player id ' 4 bytes
    id As Long 'tile id           ' 4 bytes
    data2 As Long                ' 4 bytes
End Type

Public Type MapItem3 ' 28
    data1 As Long ' creature/player id ' 4 bytes
    id As Long 'tile id ' 4 bytes
    data2 As Long ' 4 bytes
    data3(15) As Byte ' 16 new bytes
End Type

Public Type MapItem4 ' 36
    data1 As Long ' creature/player id ' 4 bytes
    id As Long 'tile id ' 4 bytes
    data2 As Long ' 4 bytes
    data3 As Long ' 4 bytes
    data4 As Long ' 4 bytes
    data5 As Long ' 4 bytes
    data6 As Long ' 4 bytes
    data7 As Long ' 4 bytes
    data8 As Long ' 4 bytes
End Type

Public Type MapItem5 ' 36
    data1 As Long ' creature/player id ' 4 bytes
    id As Long 'tile id ' 4 bytes
    data2 As Long ' 4 bytes
    data3 As Long ' 4 bytes
    data4 As Long ' 4 bytes
    data5 As Long ' 4 bytes
    data6 As Long ' 4 bytes
    data7 As Long ' 4 bytes
End Type

Public Type MapTileOld
    count As Long
    items(0 To MAX_ITEMS_PER_TILE - 1) As MapItemOld
    order(MAX_ITEMS_PER_TILE - 1) As Long ' 40 bytes
    padding As Long
    padding2 As Long
End Type

Public Type MapTile
    count As Long
    items(0 To MAX_ITEMS_PER_TILE - 1) As MapItem
    order(MAX_ITEMS_PER_TILE - 1) As Long ' 40 bytes
    padding As Long
End Type

Public Type MapTile2 ' total = 168 bytes
    count As Long ' 4 bytes
    padding As Long ' 4 bytes
    order(MAX_ITEMS_PER_TILE - 1) As Long ' 40 bytes
    items(0 To MAX_ITEMS_PER_TILE - 1) As MapItem2 ' 120 bytes
End Type

Public Type MapTile3 ' total = 328 bytes
    count As Long ' 4 bytes
    padding As Long ' 4 bytes
    order(MAX_ITEMS_PER_TILE - 1) As Long ' 40 bytes
    items(0 To MAX_ITEMS_PER_TILE - 1) As MapItem3 ' 280 bytes
End Type


Public Type MapTile4 ' total = 408 bytes
    count As Long ' 4 bytes
    padding As Long ' 4 bytes
    order(MAX_ITEMS_PER_TILE - 1) As Long ' 40 bytes
    items(0 To MAX_ITEMS_PER_TILE - 1) As MapItem4 ' 360 bytes
End Type

Public Type MapTile5 ' total = 368 bytes
    count As Long ' 4 bytes
    padding As Long ' 4 bytes
    order(MAX_ITEMS_PER_TILE - 1) As Long ' 40 bytes
    items(0 To MAX_ITEMS_PER_TILE - 1) As MapItem5 ' 320 bytes
End Type

Public MapTiles(0 To 2015) As MapTile
Public MapTiles2(0 To 2015) As MapTile2
Public MapTiles3(0 To 2015) As MapTile3
Public MapTiles4(0 To 2015) As MapTile4
Public MapTiles5(0 To 2015) As MapTile5
Public MapTilesOld(0 To 2015) As MapTileOld
Private playerId As Long
Private playerZ As Long

Public mapIsValid As Boolean

Private m_OffsetX As Long
Private m_OffsetY As Long
Private m_OffsetZ As Long
Private m_OffsetMove As Long

Private realOffset As Long
Private mapPlayerZ As Long
Private mapDebugOffZ As Long
Private myOffMove As Long
Private mapLooksValid As Boolean



Public Function DoMapIntegrityCheck() As Boolean
    Dim res As Boolean
    Dim i As Long
    Dim j As Long
    res = False
    Debug.Print "Searching player in map (ID=" & CStr(playerId) & ") ..."
    For i = 0 To 2015
        For j = 0 To MAX_ITEMS_PER_TILE - 1
            If TibiaVersionLong >= 1050 Then
                'map3
                If MapTiles5(i).items(j).data1 = playerId Then
                    res = True
                    Debug.Print "Player found at MapTiles5(" & CStr(i) & ").items(" & CStr(j) & ").data1"
                End If
                If MapTiles5(i).items(j).data2 = playerId Then
                    res = True
                    Debug.Print "Player found at MapTiles5(" & CStr(i) & ").items(" & CStr(j) & ").data2"
                End If
                If MapTiles5(i).items(j).id = playerId Then
                    res = True
                    Debug.Print "Player found at MapTiles5(" & CStr(i) & ").items(" & CStr(j) & ").id"
                End If
            ElseIf TibiaVersionLong >= 1021 Then
                'map3
                If MapTiles4(i).items(j).data1 = playerId Then
                    res = True
                    Debug.Print "Player found at MapTiles4(" & CStr(i) & ").items(" & CStr(j) & ").data1"
                End If
                If MapTiles4(i).items(j).data2 = playerId Then
                    res = True
                    Debug.Print "Player found at MapTiles4(" & CStr(i) & ").items(" & CStr(j) & ").data2"
                End If
                If MapTiles4(i).items(j).id = playerId Then
                    res = True
                    Debug.Print "Player found at MapTiles4(" & CStr(i) & ").items(" & CStr(j) & ").id"
                End If
            ElseIf TibiaVersionLong >= 990 Then
                'map3
                If MapTiles3(i).items(j).data1 = playerId Then
                    res = True
                    Debug.Print "Player found at MapTiles3(" & CStr(i) & ").items(" & CStr(j) & ").data1"
                End If
                If MapTiles3(i).items(j).data2 = playerId Then
                    res = True
                    Debug.Print "Player found at MapTiles3(" & CStr(i) & ").items(" & CStr(j) & ").data2"
                End If
                If MapTiles3(i).items(j).id = playerId Then
                    res = True
                    Debug.Print "Player found at MapTiles3(" & CStr(i) & ").items(" & CStr(j) & ").id"
                End If
            ElseIf TibiaVersionLong >= 942 Then
                'map2
                If MapTiles2(i).items(j).data1 = playerId Then
                    res = True
                    Debug.Print "Player found at MapTiles2(" & CStr(i) & ").items(" & CStr(j) & ").data1"
                End If
                If MapTiles2(i).items(j).data2 = playerId Then
                    res = True
                    Debug.Print "Player found at MapTiles2(" & CStr(i) & ").items(" & CStr(j) & ").data2"
                End If
                If MapTiles2(i).items(j).id = playerId Then
                    res = True
                    Debug.Print "Player found at MapTiles2(" & CStr(i) & ").items(" & CStr(j) & ").id"
                End If
            ElseIf TibiaVersionLong > 772 Then
                'map1
                If MapTiles(i).items(j).data1 = playerId Then
                    res = True
                    Debug.Print "Player found at MapTiles(" & CStr(i) & ").items(" & CStr(j) & ").data1"
                End If
                If MapTiles(i).items(j).data2 = playerId Then
                    res = True
                    Debug.Print "Player found at MapTiles(" & CStr(i) & ").items(" & CStr(j) & ").data2"
                End If
                If MapTiles(i).items(j).id = playerId Then
                    res = True
                    Debug.Print "Player found at MapTilesOld(" & CStr(i) & ").items(" & CStr(j) & ").id"
                End If
            Else
                If MapTilesOld(i).items(j).data1 = playerId Then
                    res = True
                    Debug.Print "Player found at MapTilesOld(" & CStr(i) & ").items(" & CStr(j) & ").data1"
                End If
                If MapTilesOld(i).items(j).data2 = playerId Then
                    res = True
                    Debug.Print "Player found at MapTilesOld(" & CStr(i) & ").items(" & CStr(j) & ").data2"
                End If
                If MapTilesOld(i).items(j).id = playerId Then
                    res = True
                    Debug.Print "Player found at MapTiles(" & CStr(i) & ").items(" & CStr(j) & ").id"
                End If
            End If
        Next j
    Next i
    If res = False Then
        Debug.Print "WARNING: Player not found in map"
    End If
  
    DoMapIntegrityCheck = res

End Function
Public Sub UpdatePlayerData(ByVal tibiaclient As Long)
    playerId = Memory_ReadLong(adrNum, tibiaclient)
    playerZ = CLng(Memory_ReadByte(PLAYER_Z, tibiaclient))
    #If FinalMode = 0 Then
    Debug.Print "Reading floor " & CStr(playerZ)
    #End If
End Sub

Public Sub UpdateMapArray(ByVal process_Hwnd As Long)

   ' Declare some variables we need
   Dim addoffset As Long
   Dim pid As Long         ' Used to hold the Process Id
   Dim phandle As Long     ' Holds the Process Handle
   Dim valbuffer As Long   ' Long
   Dim hMapAddress As Long
   Dim hOffaddress As Long
   Dim zthing As Byte
   ' First get a handle to the "game" window
   If (process_Hwnd = 0) Then
        Exit Sub
   End If
   
   ' We can now get the pid
   GetWindowThreadProcessId process_Hwnd, pid
   
   ' Use the pid to get a Process Handle
   phandle = OpenProcess(PROCESS_VM_READ, False, pid)
   If (phandle = 0) Then
    Exit Sub
   End If
   
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
     'Address = Address + TIBIA_LASTOFFSET
   End If

    'read player Z
   ' ReadProcessMemory phandle, TIBIA_LASTOFFSET + PLAYER_Z, zthing, 1, 0&
    ReadProcessMemory phandle, TIBIA_LASTOFFSET + PLAYER_Z, zthing, 1, 0&
    mapPlayerZ = CLng(zthing)
    
   ' Read offset
    myOffMove = 0
    'ReadProcessMemory phandle, TIBIA_LASTOFFSET + OFFSET_POINTER_ADDR, hOffaddress, 4, 0&
    ReadProcessMemory phandle, TIBIA_LASTOFFSET + OFFSET_POINTER_ADDR, hOffaddress, 4, 0&
    If mapPlayerZ <= 7 Then
       addoffset = 464 + ((7 - mapPlayerZ) * 1008)
       ReadProcessMemory phandle, hOffaddress + addoffset, myOffMove, 4, 0&
    Else
       addoffset = 464 + (2 * 1008)
       ReadProcessMemory phandle, hOffaddress + addoffset, myOffMove, 4, 0&
    End If
  
    'Debug.Print "Addoffset = " & Hex(addoffset) & " myOffMove: Dec=" & myOffMove & " Hex = " & Hex(myOffMove)
    realOffset = 0
   ' Read map
    ReadProcessMemory phandle, TIBIA_LASTOFFSET + MAP_POINTER_ADDR, hMapAddress, 4, 0&
    If TibiaVersionLong >= 1050 Then
      ReadProcessMemory phandle, hMapAddress, MapTiles5(0), SizeOfTibiaMap5, 0&
    ElseIf TibiaVersionLong >= 1021 Then
      ReadProcessMemory phandle, hMapAddress, MapTiles4(0), SizeOfTibiaMap4, 0&
    ElseIf TibiaVersionLong >= 990 Then
      ReadProcessMemory phandle, hMapAddress, MapTiles3(0), SizeOfTibiaMap3, 0&
    ElseIf TibiaVersionLong >= 942 Then
      ReadProcessMemory phandle, hMapAddress, MapTiles2(0), SizeOfTibiaMap, 0&
    ElseIf TibiaVersionLong > 772 Then
      ReadProcessMemory phandle, hMapAddress, MapTiles(0), SizeOfTibiaMap, 0&
    Else
      ReadProcessMemory phandle, hMapAddress, MapTilesOld(0), SizeOfTibiaMapOld, 0&
    End If
  
    #If FinalMode = 0 Then
    Debug.Print ">> My offmove = " & myOffMove ' a low number, example = 197
    Debug.Print ">> Map starts at address " & hMapAddress
    mapLooksValid = DoMapIntegrityCheck()
    #End If
   ' Close the Process Handle
   CloseHandle phandle
End Sub

Public Sub UpdateMap(ByVal hProcess As Long)
    If hProcess = 0 Then
        mapIsValid = False
    Else
        UpdatePlayerData hProcess
        UpdateMapArray hProcess
        GetPlayerCenter3 ' if it fails in future versions then change to safe method  GetPlayerCenter3
        mapIsValid = True
    End If
End Sub

Private Sub GetPlayerCenter3()


    Dim x As Long
    Dim y As Long
    Dim px As Long
    Dim py As Long
    Dim pz As Long
    Dim i As Long
    Dim currentTile As Long
    Dim lim As Long
    Dim totalFound As Long
    Dim thing1 As Long
    Dim thing2 As Long
    Dim thing3 As Long
    
    totalFound = 0
    currentTile = 0
    m_OffsetMove = 0
    For pz = 0 To 7
        For py = 0 To 13
            For px = 0 To 17
                If currentTile = myOffMove Then ' This can be optimized a lot, but I am lazy at this moment...
                    m_OffsetX = px - 8
                    If (m_OffsetX < 0) Then
                       m_OffsetX = m_OffsetX + 18
                    End If
                    m_OffsetY = py - 6
                    If (m_OffsetY < 0) Then
                       m_OffsetY = m_OffsetY + 14
                    End If
                    m_OffsetZ = pz - (15 - playerZ)
                    m_OffsetMove = currentTile
                    totalFound = totalFound + 1
                    'Debug.Print "total found=" & totalFound & " ; LAST FOUND HERE: offx=" & CStr(m_OffsetX) & " offy=" & CStr(m_OffsetY) & " offz=" & CStr(m_OffsetZ) & " offMove=" & CStr(m_OffsetMove) & vbCrLf
                    Exit Sub
                End If
                currentTile = currentTile + 1
            Next px
        Next py
    Next pz
    Debug.Print "total found=" & totalFound & " ; LAST FOUND HERE: offx=" & CStr(m_OffsetX) & " offy=" & CStr(m_OffsetY) & " offz=" & CStr(m_OffsetZ) & " offMove=" & CStr(m_OffsetMove) & vbCrLf
 End Sub



Private Sub GetPlayerCenter()
    Dim x As Long
    Dim y As Long
    Dim px As Long
    Dim py As Long
    Dim pz As Long
    Dim i As Long
    Dim currentTile As Long
    Dim lim As Long
    Dim xoff As Long
    Dim yoff As Long
    Dim zoff As Long
    Dim therest As Long
    ' new method to locate center, based on the offset
    zoff = (realOffset \ 252)
    therest = realOffset - (zoff * 252)
    xoff = therest Mod 18
    yoff = therest \ 18
    therest = myOffMove - xoff - (yoff * 18)
    zoff = therest \ (18 * 14)
    zoff = zoff - 15 + mapPlayerZ
    zoff = ObtainOffsetZ(xoff, yoff, mapPlayerZ, myOffMove)
    mapDebugOffZ = zoff
    m_OffsetX = xoff
    m_OffsetY = yoff
    m_OffsetZ = zoff
    m_OffsetMove = myOffMove
    'Debug.Print "realoffset=" & realOffset & vbCrLf & "xOff=" & CStr(xoff) & " yOff=" & CStr(yoff) & " zoff=" & CStr(zoff) & " offMov1=" & CStr(myOffMove)
   
'    Exit Sub
'    currentTile = 0
'    m_OffsetMove = 0
'    For pz = 0 To 7
'        For py = 0 To 13
'            For px = 0 To 17
'                lim = MapTiles(currentTile).count - 1
'                For i = 0 To lim
'                    If MapTiles(currentTile).items(i).id = &H63 Then
'                        If MapTiles(currentTile).items(i).data1 = playerId Then
'                             m_OffsetX = px - 8
'                             If (m_OffsetX < 0) Then
'                                m_OffsetX = m_OffsetX + 18
'                             End If
'                             m_OffsetY = py - 6
'                             If (m_OffsetY < 0) Then
'                                m_OffsetY = m_OffsetY + 14
'                             End If
'                             m_OffsetZ = pz - (15 - playerZ)
'                             m_OffsetMove = currentTile
'                        End If
'                    End If
'                Next i
'                currentTile = currentTile + 1
'            Next px
'        Next py
'    Next pz
'   ' m_OffsetZ = (m_OffsetZ + 16) Mod 8
'    Debug.Print "offx=" & CStr(m_OffsetX) & " offy=" & CStr(m_OffsetY) & " offz=" & CStr(m_OffsetZ) & " offMove=" & CStr(m_OffsetMove) & vbCrLf
End Sub

Private Sub GetPlayerCenter2()
    Dim x As Long
    Dim y As Long
    Dim px As Long
    Dim py As Long
    Dim pz As Long
    Dim i As Long
    Dim currentTile As Long
    Dim lim As Long
    Dim totalFound As Long
    Dim thing1 As Long
    Dim thing2 As Long
    Dim thing3 As Long
    
    totalFound = 0
    currentTile = 0
    m_OffsetMove = 0
    For pz = 0 To 7
        For py = 0 To 13
            For px = 0 To 17
                If TibiaVersionLong >= 1050 Then
                    thing1 = MapTiles5(currentTile).count
                ElseIf TibiaVersionLong >= 1021 Then
                    thing1 = MapTiles4(currentTile).count
                ElseIf TibiaVersionLong >= 990 Then
                    thing1 = MapTiles3(currentTile).count
                ElseIf TibiaVersionLong >= 942 Then
                    thing1 = MapTiles2(currentTile).count
                ElseIf TibiaVersionLong > 772 Then
                    thing1 = MapTiles(currentTile).count
                Else
                    thing1 = MapTilesOld(currentTile).count
                End If
                lim = thing1 - 1
                For i = 0 To lim
                    If TibiaVersionLong >= 1050 Then
                        thing2 = MapTiles5(currentTile).items(i).id
                        thing3 = MapTiles5(currentTile).items(i).data1
                    ElseIf TibiaVersionLong >= 1021 Then
                        thing2 = MapTiles4(currentTile).items(i).id
                        thing3 = MapTiles4(currentTile).items(i).data1
                    ElseIf TibiaVersionLong >= 990 Then
                        thing2 = MapTiles3(currentTile).items(i).id
                        thing3 = MapTiles3(currentTile).items(i).data1
                    ElseIf TibiaVersionLong >= 942 Then
                        thing2 = MapTiles2(currentTile).items(i).id
                        thing3 = MapTiles2(currentTile).items(i).data1
                    ElseIf TibiaVersionLong > 772 Then
                        thing2 = MapTiles(currentTile).items(i).id
                        thing3 = MapTiles(currentTile).items(i).data1
                    Else
                        thing2 = MapTilesOld(currentTile).items(i).id
                        thing3 = MapTilesOld(currentTile).items(i).data1
                    End If
                    If thing2 = &H63 Then
                        If thing3 = playerId Then
                             m_OffsetX = px - 8
                             If (m_OffsetX < 0) Then
                                m_OffsetX = m_OffsetX + 18
                             End If
                             m_OffsetY = py - 6
                             If (m_OffsetY < 0) Then
                                m_OffsetY = m_OffsetY + 14
                             End If
                             m_OffsetZ = pz - (15 - playerZ)
                             m_OffsetMove = currentTile
                             totalFound = totalFound + 1
                        End If
                    End If
                Next i
                currentTile = currentTile + 1
            Next px
        Next py
    Next pz
    Debug.Print "total found=" & totalFound & " ; LAST FOUND HERE: offx=" & CStr(m_OffsetX) & " offy=" & CStr(m_OffsetY) & " offz=" & CStr(m_OffsetZ) & " offMove=" & CStr(m_OffsetMove) & vbCrLf
End Sub

Public Function ObtainOffsetZ(ByVal offX, ByVal offY, ByVal Z, ByVal offmove) As Long
    Dim x As Long
    Dim y As Long
    Dim px As Long
    Dim py As Long
    Dim pz As Long
    Dim lim As Long
    Dim PosX As Long
    Dim PosY As Long
    Dim PosZ As Long
    Dim compP As Long
    Dim arrayPos As Long
    Dim resOff As Long
    Dim sol As Long
    Dim i As Long
    For i = -7 To 7
    x = 0 + 8
    y = 0 + 6
    PosX = x + offX
    If (PosX > 17) Then
        PosX = PosX - 18
    End If
    PosY = y + offY
    If (PosY > 13) Then
        PosY = PosY - 14
    End If
    PosZ = (15 - Z) + i
    PosZ = (PosZ + 16) Mod 8
    arrayPos = PosX + (PosY * 18) + (PosZ * 14 * 18)
    If arrayPos = offmove Then
        sol = i
        ObtainOffsetZ = sol
        Exit Function
    End If
    Next i
End Function
Public Function GetMapTile(ByVal incX As Long, ByVal incY As Long, ByVal Z As Long) As MapTile
    Dim x As Long
    Dim y As Long
    Dim px As Long
    Dim py As Long
    Dim pz As Long
    Dim lim As Long
    Dim PosX As Long
    Dim PosY As Long
    Dim PosZ As Long
    Dim compP As Long
    Dim arrayPos As Long
    x = incX + 8
    y = incY + 6
    PosX = x + m_OffsetX
    If (PosX > 17) Then
        PosX = PosX - 18
    End If
    PosY = y + m_OffsetY
    If (PosY > 13) Then
        PosY = PosY - 14
    End If
    PosZ = (15 - Z) + m_OffsetZ
    PosZ = (PosZ + 16) Mod 8

    arrayPos = PosX + (PosY * 18) + (PosZ * 14 * 18)
    'compP = PosX + (PosY * 18) + (mapDebugOffZ * 14 * 18)
    'Debug.Print arrayPos
    'Debug.Print m_OffsetMove
    
    GetMapTile = MapTiles(arrayPos)
End Function

Public Function GetMapTile2(ByVal incX As Long, ByVal incY As Long, ByVal Z As Long) As MapTile2
    Dim x As Long
    Dim y As Long
    Dim px As Long
    Dim py As Long
    Dim pz As Long
    Dim lim As Long
    Dim PosX As Long
    Dim PosY As Long
    Dim PosZ As Long
    Dim compP As Long
    Dim arrayPos As Long
    x = incX + 8
    y = incY + 6
    PosX = x + m_OffsetX
    If (PosX > 17) Then
        PosX = PosX - 18
    End If
    PosY = y + m_OffsetY
    If (PosY > 13) Then
        PosY = PosY - 14
    End If
    PosZ = (15 - Z) + m_OffsetZ
    PosZ = (PosZ + 16) Mod 8

    arrayPos = PosX + (PosY * 18) + (PosZ * 14 * 18)
    'compP = PosX + (PosY * 18) + (mapDebugOffZ * 14 * 18)
    'Debug.Print arrayPos
    'Debug.Print m_OffsetMove
    
    GetMapTile2 = MapTiles2(arrayPos)
End Function
Public Function GetMapTile5(ByVal incX As Long, ByVal incY As Long, ByVal Z As Long) As MapTile5
    Dim x As Long
    Dim y As Long
    Dim px As Long
    Dim py As Long
    Dim pz As Long
    Dim lim As Long
    Dim PosX As Long
    Dim PosY As Long
    Dim PosZ As Long
    Dim compP As Long
    Dim arrayPos As Long
    x = incX + 8
    y = incY + 6
    PosX = x + m_OffsetX
    If (PosX > 17) Then
        PosX = PosX - 18
    End If
    PosY = y + m_OffsetY
    If (PosY > 13) Then
        PosY = PosY - 14
    End If
    PosZ = (15 - Z) + m_OffsetZ
    PosZ = (PosZ + 16) Mod 8

    arrayPos = PosX + (PosY * 18) + (PosZ * 14 * 18)
    'compP = PosX + (PosY * 18) + (mapDebugOffZ * 14 * 18)
    'Debug.Print arrayPos
    'Debug.Print m_OffsetMove
    
    GetMapTile5 = MapTiles5(arrayPos)
End Function

Public Function GetMapTile4(ByVal incX As Long, ByVal incY As Long, ByVal Z As Long) As MapTile4
    Dim x As Long
    Dim y As Long
    Dim px As Long
    Dim py As Long
    Dim pz As Long
    Dim lim As Long
    Dim PosX As Long
    Dim PosY As Long
    Dim PosZ As Long
    Dim compP As Long
    Dim arrayPos As Long
    x = incX + 8
    y = incY + 6
    PosX = x + m_OffsetX
    If (PosX > 17) Then
        PosX = PosX - 18
    End If
    PosY = y + m_OffsetY
    If (PosY > 13) Then
        PosY = PosY - 14
    End If
    PosZ = (15 - Z) + m_OffsetZ
    PosZ = (PosZ + 16) Mod 8

    arrayPos = PosX + (PosY * 18) + (PosZ * 14 * 18)
    'compP = PosX + (PosY * 18) + (mapDebugOffZ * 14 * 18)
    'Debug.Print arrayPos
    'Debug.Print m_OffsetMove
    
    GetMapTile4 = MapTiles4(arrayPos)
End Function


Public Function GetMapTile3(ByVal incX As Long, ByVal incY As Long, ByVal Z As Long) As MapTile3
    Dim x As Long
    Dim y As Long
    Dim px As Long
    Dim py As Long
    Dim pz As Long
    Dim lim As Long
    Dim PosX As Long
    Dim PosY As Long
    Dim PosZ As Long
    Dim compP As Long
    Dim arrayPos As Long
    x = incX + 8
    y = incY + 6
    PosX = x + m_OffsetX
    If (PosX > 17) Then
        PosX = PosX - 18
    End If
    PosY = y + m_OffsetY
    If (PosY > 13) Then
        PosY = PosY - 14
    End If
    PosZ = (15 - Z) + m_OffsetZ
    PosZ = (PosZ + 16) Mod 8

    arrayPos = PosX + (PosY * 18) + (PosZ * 14 * 18)
    'compP = PosX + (PosY * 18) + (mapDebugOffZ * 14 * 18)
    'Debug.Print arrayPos
    'Debug.Print m_OffsetMove
    
    GetMapTile3 = MapTiles3(arrayPos)
End Function


Public Function GetMapTileOld(ByVal incX As Long, ByVal incY As Long, ByVal Z As Long) As MapTileOld
    Dim x As Long
    Dim y As Long
    Dim px As Long
    Dim py As Long
    Dim pz As Long
    Dim lim As Long
    Dim PosX As Long
    Dim PosY As Long
    Dim PosZ As Long
    Dim compP As Long
    Dim arrayPos As Long
    x = incX + 8
    y = incY + 6
    PosX = x + m_OffsetX
    If (PosX > 17) Then
        PosX = PosX - 18
    End If
    PosY = y + m_OffsetY
    If (PosY > 13) Then
        PosY = PosY - 14
    End If
    PosZ = (15 - Z) + m_OffsetZ
    PosZ = (PosZ + 16) Mod 8

    arrayPos = PosX + (PosY * 18) + (PosZ * 14 * 18)
    'compP = PosX + (PosY * 18) + (mapDebugOffZ * 14 * 18)
    'Debug.Print arrayPos
    'Debug.Print m_OffsetMove
    
    GetMapTileOld = MapTilesOld(arrayPos)
End Function
