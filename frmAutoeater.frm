VERSION 5.00
Begin VB.Form frmAutoeater 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autoeater"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4800
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmAutoeater.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3525
   ScaleWidth      =   4800
   Begin VB.Timer timerTicTac 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3240
      Top             =   840
   End
   Begin VB.TextBox txtFrom 
      Height          =   315
      Left            =   840
      TabIndex        =   3
      Text            =   "40"
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox txtTo 
      Height          =   315
      Left            =   3120
      TabIndex        =   2
      Text            =   "60"
      Top             =   1440
      Width           =   615
   End
   Begin VB.ComboBox cmbHotkey 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "APPLY"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblInfo1 
      BackColor       =   &H00000000&
      Caption         =   "Use this tool to eat food in Tibia from time to time or to do something else and act as anti-iddle."
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   12
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label lblInfo2 
      BackColor       =   &H00000000&
      Caption         =   "To avoid detection, the action should be reapeated randomly in a variable time..."
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label lblFrom 
      BackColor       =   &H00000000&
      Caption         =   "from"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblSeconds1 
      BackColor       =   &H00000000&
      Caption         =   "seconds"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1560
      TabIndex        =   9
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblTo 
      BackColor       =   &H00000000&
      Caption         =   "to"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblSeconds2 
      BackColor       =   &H00000000&
      Caption         =   "seconds"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblInfo3 
      BackColor       =   &H00000000&
      Caption         =   "Warning: the hotkey only will be pressed if a Tibia window is focused!"
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      Width           =   4455
   End
   Begin VB.Label lblHotkey 
      BackColor       =   &H00000000&
      Caption         =   "Tibia hotkey to press:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblCurrent 
      BackColor       =   &H00000000&
      Caption         =   "Current timer: 40-60 seconds"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   3375
   End
End
Attribute VB_Name = "frmAutoeater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit

Private Const tibiaclassname As String = "TibiaClient"
Private Const defaultFHotkey As Long = 13 'none
Private Const defaultTimerFrom As Double = 40
Private Const defaultTimerTo As Double = 60
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
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

#If Win32 Then
  Private Declare Function GetTickCount Lib "Kernel32" () As Long
#Else
  Private Declare Function GetTickCount Lib "user" () As Long
#End If

Private NextHotkeyTime As Long
Private TimerFrom As Double
Private TimerTo As Double

Private Function GetFocusedTibiaPID() As Long
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

Private Sub cmdApply_Click()
    On Error GoTo goterr
    Dim v1 As Double
    Dim v2 As Double
    v1 = Round(CDbl(txtFrom.Text), 4)
    v2 = Round(CDbl(txtTo.Text), 4)
    If v1 < 0.1 Then
        GoTo goterr
    End If
    If v2 > 3600 Then
        GoTo goterr
    End If
    If v1 > v2 Then
        GoTo goterr
    End If
    If v2 < v1 Then
        GoTo goterr
    End If
    TimerFrom = v1
    TimerTo = v2
    txtFrom.Text = CStr(TimerFrom)
    txtTo.Text = CStr(TimerTo)
    lblCurrent.Caption = BString(73) & " " & CStr(TimerFrom) & "-" & _
     CStr(TimerTo) & " " & BString(71)
    NextHotkeyTime = 0
    Exit Sub
goterr:
    txtFrom.Text = CStr(TimerFrom)
    txtTo.Text = CStr(TimerTo)
End Sub

Private Sub Form_Load()
    NextHotkeyTime = 0
    TimerFrom = defaultTimerFrom
    TimerTo = defaultTimerTo
    txtFrom.Text = CStr(TimerFrom)
    txtTo.Text = CStr(TimerTo)
    Me.cmbHotkey.Clear
    Me.cmbHotkey.AddItem "F1"
    Me.cmbHotkey.AddItem "F2"
    Me.cmbHotkey.AddItem "F3"
    Me.cmbHotkey.AddItem "F4"
    Me.cmbHotkey.AddItem "F5"
    Me.cmbHotkey.AddItem "F6"
    Me.cmbHotkey.AddItem "F7"
    Me.cmbHotkey.AddItem "F8"
    Me.cmbHotkey.AddItem "F9"
    Me.cmbHotkey.AddItem "F10"
    Me.cmbHotkey.AddItem "F11"
    Me.cmbHotkey.AddItem "F12"
    Me.cmbHotkey.AddItem "--"
    cmbHotkey.ListIndex = defaultFHotkey - 1
    timerTicTac.Enabled = True
End Sub

Private Function randomNumberBetween(limite_inferior As Long, limite_superior As Long) As Long
  Dim res As Long
  Randomize
  If limite_inferior <= limite_superior Then
    res = Int((limite_superior - limite_inferior + 1) * Rnd + limite_inferior)
  Else
    res = Int((limite_inferior - limite_superior + 1) * Rnd + limite_superior)
  End If
  If (res >= limite_inferior) And (res <= limite_superior) Then
     randomNumberBetween = res
  Else
     MsgBox "Critical error: random number between " & CStr(limite_inferior) & " and " & CStr(limite_superior) & " returned " & CStr(res), vbOKOnly + vbCritical, "randomNumberBetween"
     End
  End If
End Function





Private Sub timerTicTac_Timer()
    Dim pressFKey As Long
    Dim gtc As Long
    Dim rtime As Long
    Dim strAction As String
    pressFKey = cmbHotkey.ListIndex + 1
    If ((pressFKey >= 1) And (pressFKey <= 12)) Then
      gtc = GetTickCount()
      If (gtc >= NextHotkeyTime) Then
        rtime = randomNumberBetween(CLng(TimerFrom * 1000), CLng(TimerTo * 1000))
        NextHotkeyTime = gtc + rtime
        If GetFocusedTibiaPID() <> 0 Then
            strAction = "{F" & CStr(pressFKey) & "}"
            SendKeysSAFE strAction
            
        End If
      End If
    End If
End Sub

Public Sub UpdateALL()
cmbHotkey.ListIndex = AutoeaterKey - 1
TimerFrom = CDbl(AutoeaterTimerFrom) / 1000
TimerTo = CDbl(AutoeaterTimerTo) / 1000
txtFrom.Text = CStr(TimerFrom)
txtTo.Text = CStr(TimerTo)
Me.UpdateLanguage
End Sub


Public Sub UpdatePublicVars()
 AutoeaterKey = cmbHotkey.ListIndex + 1
 AutoeaterTimerFrom = CLng(TimerFrom * 1000)
 AutoeaterTimerTo = CLng(TimerTo * 1000)
End Sub

Public Sub UpdateLanguage()
    Me.Caption = BString(67)
    
    Me.lblCurrent.Caption = BString(73) & " " & CStr(TimerFrom) & "-" & _
      CStr(TimerTo) & " " & BString(71)
    Me.lblInfo1.Caption = BString(68)
    Me.lblInfo2.Caption = BString(69)
    Me.lblFrom.Caption = BString(70)
    Me.lblTo.Caption = BString(72)
    Me.lblHotkey.Caption = BString(74)
    Me.lblInfo3.Caption = BString(75)
    Me.cmdApply.Caption = BString(76)
    Me.lblSeconds1.Caption = BString(71)
    Me.lblSeconds2.Caption = BString(71)
End Sub
