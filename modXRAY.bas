Attribute VB_Name = "modXRAY"
Option Explicit
#Const FinalMode = 1


'======Constants=======
'API constants
Private Const PROCESS_ALL_ACCESS = &H1F0FFF
'Statusbar
Const STATUSBAR_DURATION = 50
'Levelspy
Const LEVELSPY_NOP_DEFAULT = 49451
Const LEVELSPY_ABOVE_DEFAULT = 7
Const LEVELSPY_BELOW_DEFAULT = 2
Const LEVELSPY_MIN = 0
Const LEVELSPY_MAX = 7
'name spy
Const NAMESPY_NOP_DEFAULT = 19573
Const NAMESPY_NOP2_DEFAULT = 17013
'z-axis
Const Z_AXIS_DEFAULT = 7 'default ground level

' some colors for map
Public Const ColourField = &H80&
Public Const ColourPath = &HFFFFC0
Public Const ColourNothing = &H111111


' some more colors for map
Public Const ColourGround = &HC0FFC0
Public Const ColourPlayer = &H80C0FF



' some more colors for map
Public Const ColourUnknown = &H8000000F
Public Const ColourWithInfo = &HC0FFC0
Public Const ColourWithMe = &H8080FF
Public Const ColourSelected = &HC0FFFF
Public Const ColourSomething = &H808080
Public Const ColourSomething2 = &H80&
Public Const ColourSomething3 = &HFFC0FF
Public Const ColourSomething4 = &HFF80FF



' some more colors for map
Public Const ColourSomething5 = &HFF00FF
Public Const ColourSomething6 = &HC000C0
Public Const ColourSomething7 = &HFFC0C0
Public Const ColourDown = &H800080
Public Const ColourUp = &HFF00FF
Public Const ColourBlockMoveable = &HC000&
Public Const ColourWater = &HFF8080
Public Const ColourFish = &HFFFF00



Private bLevelSpy As Boolean

Public Function MemoryChangeFloor(pid As Long, relfloornumber As String) As Long 'receives mc id and relative floor increase desired
    On Error GoTo goterr
    Dim floornumber As Long
    Dim relChange As Long
    Dim ammountOfChanges As Long
    Dim i As Long
    'MemoryChangeFloor = -1 ' not working yet
    'Exit Function
    If IsNumeric(relfloornumber) = False Then
        MemoryChangeFloor = -1 'failure (bad parameters)
        Exit Function
    End If
    relChange = CLng(relfloornumber)
    ammountOfChanges = Abs(relChange)
    levelSpy_Off pid
    If ammountOfChanges > 0 Then
        Call WriteNops(pid, LEVELSPY_NOP, 2)
        'Initialize Level spying
        LevelSpy_Init pid
        'Set boolean
        bLevelSpy = True
        
        'full light
        If LIGHT_NOP = 0 Then
            SetTibiaPermaLight pid, True
        Else
            Call WriteNops(pid, LIGHT_NOP, 2)
            Call writeBytes(pid, LIGHT_AMOUNT, 255, 1)
        End If
    Else
        If LIGHT_NOP = 0 Then
            SetTibiaPermaLight pid, False
        End If
    End If
    For i = 1 To ammountOfChanges
        If relChange > 0 Then
            levelSpy_Down pid
        Else
            levelSpy_Up pid
        End If
    Next i
    MemoryChangeFloor = 0 ' sucess
    Exit Function
goterr:
    MemoryChangeFloor = -1 ' failure (unknown)
End Function




Public Sub levelSpy_Off(pid As Long)
'disable level spying by restoring default values
Call writeBytes(pid, LEVELSPY_NOP, LEVELSPY_NOP_DEFAULT, 2)
Call writeBytes(pid, LEVELSPY_ABOVE, LEVELSPY_ABOVE_DEFAULT, 1)
Call writeBytes(pid, LEVELSPY_BELOW, LEVELSPY_BELOW_DEFAULT, 1)
'Set boolean
bLevelSpy = False
End Sub
 
Public Sub WriteSpecial3Nops(ByVal pid As Long, ByVal adr As Long, ByVal b1 As Byte, ByVal b2 As Byte, ByVal b3 As Byte)
Memory_WriteByte adr, b1, pid
Memory_WriteByte adr + 1, b2, pid
Memory_WriteByte adr + 2, b3, pid
End Sub
Public Sub WriteNops(pid As Long, Address As Long, Nops As Integer)

'Get Process Handle
Dim ProcessHandle As Long
ProcessHandle = pid

'Write Memory
Dim i, j As Integer
i = 0: j = 0
For i = 1 To Nops
Const nop = &H90
Memory_WriteByte Address + j, nop, ProcessHandle
j = j + 1
Next i
'Close process handle

End Sub

Private Sub writeBytes(pid As Long, Address As Long, Value As Long, bytes As Integer)
'Get Process Handle
Dim ProcessHandle As Long
ProcessHandle = pid
'write to memory
If bytes = 1 Then
  'Debug.Print "Writting 1 byte [" & CStr(ProcessHandle) & "] at address & " & CStr(Address) & " :" & CStr(Value)
  Memory_WriteByte Address, CByte(Value), ProcessHandle
Else
  'Debug.Print "Writting 2 byte [" & CStr(ProcessHandle) & "] at address & " & CStr(Address) & " :" & CStr(Value)
  
  ' experimental
  'Memory_WriteLong Address, Value, ProcessHandle
  Memory_WriteByte Address, LowByteOfLong(Value), ProcessHandle
  Memory_WriteByte Address + 1, HighByteOfLong(Value), ProcessHandle
End If
End Sub

'Initialize level spying
Public Sub LevelSpy_Init(pid As Long)
'Get player Z
Dim playerZ As Integer
playerZ = readBytes(pid, PLAYER_Z, 1)
'Set levelspy to current level
If (playerZ <= Z_AXIS_DEFAULT) Then
    'Above ground
    Call writeBytes(pid, LEVELSPY_ABOVE, Z_AXIS_DEFAULT - playerZ, 1)
    Call writeBytes(pid, LEVELSPY_BELOW, LEVELSPY_BELOW_DEFAULT, 1)
Else
    'Below Ground
    Call writeBytes(pid, LEVELSPY_ABOVE, LEVELSPY_ABOVE_DEFAULT, 1)
    Call writeBytes(pid, LEVELSPY_BELOW, LEVELSPY_BELOW_DEFAULT, 1)
End If
End Sub

'Increase spy level
Public Sub levelSpy_Up(pid As Long)
'Levelspy must be on

'Get player z
Dim playerZ As Integer
playerZ = readBytes(pid, PLAYER_Z, 1)
'Ground level
Dim groundLevel As Long
groundLevel = 0
If playerZ <= Z_AXIS_DEFAULT Then
    groundLevel = LEVELSPY_ABOVE ' above ground
Else
    groundLevel = LEVELSPY_BELOW ' below ground
End If
    
'Get Current level
Dim currentLevel As Integer
currentLevel = readBytes(pid, groundLevel, 1)
If currentLevel >= LEVELSPY_MAX Then
    Call writeBytes(pid, groundLevel, LEVELSPY_MIN, 1) ' Loop back to start
Else
    Call writeBytes(pid, groundLevel, currentLevel + 1, 1) ' increase spy level
    
'Set statusbar
'setStatusBar ("Level Spy: Up")
End If
End Sub

'Decrease spy level
Public Sub levelSpy_Down(pid As Long)
'Levelspy must be on
If bLevelSpy = False Then
'setStatusBar ("Please Enable Level Spy first!")
Exit Sub
End If
'Get player z
Dim playerZ As Integer
playerZ = readBytes(pid, PLAYER_Z, 1)
'Ground level
Dim groundLevel As Long
groundLevel = 0
If playerZ <= Z_AXIS_DEFAULT Then
    groundLevel = LEVELSPY_ABOVE ' above ground
Else
    groundLevel = LEVELSPY_BELOW ' below ground
End If
    
'Get Current level
Dim currentLevel As Integer
currentLevel = readBytes(pid, groundLevel, 1)
If currentLevel <= LEVELSPY_MIN Then
    Call writeBytes(pid, groundLevel, LEVELSPY_MAX, 1) ' Loop back to start
Else
    Call writeBytes(pid, groundLevel, currentLevel - 1, 1) ' increase spy level
    
'Set statusbar
'setStatusBar ("Level Spy: Down")
End If
End Sub


Public Function readBytes(pid As Long, Address As Long, bytes As Integer) As Long
'Get Process Handle
Dim ProcessHandle As Long
Dim b1 As Byte
Dim b2 As Byte
ProcessHandle = pid

'read memory
Dim buffer As Long
buffer = 0
If bytes = 1 Then
    readBytes = Memory_ReadByte(Address, ProcessHandle)
Else
    b1 = Memory_ReadByte(Address, ProcessHandle)
    b2 = Memory_ReadByte(Address + 1, ProcessHandle)
    readBytes = GetTheLong(b1, b2)
End If
'Call ReadProcessMemory(processHandle, Address, buffer, bytes, 0)
'Close handle
'CloseHandle (processHandle)
'readBytes = buffer
End Function
