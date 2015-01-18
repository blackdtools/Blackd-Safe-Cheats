Attribute VB_Name = "modHotkeys"
#Const FinalMode = 1
Option Explicit

Public Const lastHotkey As Long = 2
   Public Const KEY_TOGGLED As Integer = &H1
   Public Const KEY_PRESSED As Integer = &H1000



Public Type TypeHotkey
  key1 As Byte
  key2 As Byte
  command As String
  usable As Boolean
End Type



Public Hotkeys(0 To lastHotkey) As TypeHotkey
Public espectingHotkey As Boolean
Public DefiningHotkeyLine As Long
Public DefiningHotkeySub As Long

Public debugdxError As String

Private Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer


Public HotkeysAreUsable As Boolean


Public Function TranslateHotkeyID2(HotkeyID As Byte) As String
  Dim res As String
  Select Case HotkeyID
  Case 0
    res = BString(47)
'  Case 27
'    res = "ESCAPE"
'  Case 49
'    res = "1"
'  Case 50
'    res = "2"
'  Case 51
'    res = "3"
'  Case 52
'    res = "4"
'  Case 53
'   res = "5"
'  Case 54
'    res = "6"
'  Case 55
'    res = "7"
'  Case 56
'    res = "8"
'  Case 57
'   res = "9"
'  Case 48
'    res = "0"
'  Case 8
'    res = "BACKSPACE"
'  Case 9
'    res = "TAB"
'  Case 81
'    res = "Q"
'  Case 87
'    res = "W"
'  Case 69
'    res = "E" ' hasta aqui
'  Case 19
'    res = "R"
'  Case 20
'    res = "T"
'  Case 21
'    res = "Y"
'  Case 22
'    res = "U"
'  Case 23
'    res = "I"
'  Case 24
'    res = "O"
'  Case 25
'    res = "P"
'  Case 28
'    res = "ENTER"
'  Case 29
'    res = "L-CONTROL"
'  Case 30
'    res = "A"
'  Case 31
'    res = "S"
'  Case 32
'    res = "D"
'  Case 33
'    res = "F"
'  Case 34
'    res = "G"
'  Case 35
'    res = "H"
'  Case 36
'    res = "J"
'  Case 37
'    res = "K"
'  Case 38
'    res = "L"
'  Case 42
'    res = "L-SHIFT"
'  Case 44
'    res = "Z"
'  Case 45
'    res = "X"
'  Case 46
'    res = "C"
'  Case 47
'    res = "V"
'  Case 48
'    res = "B"
'  Case 49
'    res = "N"
'  Case 50
'    res = "M"
'  Case 51
'    res = ","
'  Case 52
'    res = "."
'  Case 53
'    res = "-"
'  Case 54
'    res = "R-SHIFT"
'  Case 55
'    res = "PAD *"
'  Case 56
'    res = "L-ALT"
'  Case 57
'    res = "SPACE"
'  Case 58
'    res = "CAPS"
'  Case 59
'    res = "F1"
'  Case 60
'    res = "F2"
'  Case 61
'    res = "F3"
'  Case 62
'    res = "F4"
'  Case 63
'    res = "F5"
'  Case 64
'    res = "F6"
'  Case 65
'    res = "F7"
'  Case 66
'    res = "F8"
'  Case 67
'    res = "F9"
'  Case 68
'    res = "F10"
'  Case 69
'    res = "PAD LOCK"
'  Case 70
'    res = "LOCK"
'  Case 71
'    res = "PAD 7"
'  Case 72
'    res = "PAD 8"
'  Case 73
'    res = "PAD 9"
'  Case 74
'    res = "PAD -"
'  Case 75
'    res = "PAD 4"
'  Case 76
'    res = "PAD 5"
'  Case 77
'    res = "PAD 6"
'  Case 78
'    res = "PAD +"
'  Case 79
'    res = "PAD 1"
'  Case 80
'    res = "PAD 2"
'  Case 81
'    res = "PAD 3"
'  Case 82
'    res = "PAD 0"
'  Case 83
'    res = "PAD ."
'  Case 87
'    res = "F11"
'  Case 88
'    res = "F12"
'  Case 156
'    res = "PAD ENTER"
'  Case 157
'    res = "R-CONTROL"
'  Case 197
'    res = "PAUSE"
'  Case 181
'    res = "PAD /"
'  Case 183
'    res = "PRINT"
'  Case 184
'    res = "R-ALT"
'  Case 199
'    res = "HOME"
'  Case 200
'    res = "UP ARROW"
'  Case 201
'    res = "PAG UP"
'  Case 203
'    res = "LEFT ARROW"
'  Case 205
'    res = "RIGHT ARROW"
'  Case 207
'    res = "END"
'  Case 208
'    res = "DOWN ARROW"
'  Case 209
'    res = "PAG DOWN"
'  Case 210
'    res = "INSERT"
'  Case 211
'    res = "DELETE"
'  Case 219
'    res = "WINDOWS"
'  Case 221
'    res = "MENU"
  Case Else
    res = BString(61) & " @" & CStr(CInt(HotkeyID))
  End Select
  TranslateHotkeyID2 = res
End Function



Public Function InitDI() As String
  #If FinalMode Then
    On Error GoTo justend
  #End If
  Dim res As String
  res = ""
  HotkeysAreUsable = False

  HotkeysAreUsable = True
  Exit Function
justend:
  debugdxError = "Error number: " & Err.Number & " ; Error description: " & Err.Description
  InitDI = res
  Exit Function
End Function
