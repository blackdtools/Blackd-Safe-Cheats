VERSION 5.00
Begin VB.Form frmXRAY 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "XRAY"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6105
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmXRAY.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3165
   ScaleWidth      =   6105
   Begin VB.TextBox txtFloors3 
      Height          =   375
      Left            =   4560
      TabIndex        =   25
      Text            =   "0"
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox txtKey31 
      Height          =   375
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "<nothing>"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdDel31 
      BackColor       =   &H00808080&
      Caption         =   "X"
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1920
      Width           =   255
   End
   Begin VB.TextBox txtKey32 
      Height          =   375
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "<nothing>"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdDel32 
      BackColor       =   &H00808080&
      Caption         =   "X"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton cmdTest3 
      Caption         =   "Test"
      Height          =   375
      Left            =   5280
      TabIndex        =   20
      Top             =   1920
      Width           =   735
   End
   Begin VB.ComboBox cmbFloors2 
      Height          =   330
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   1440
      Width           =   615
   End
   Begin VB.ComboBox cmbFloors1 
      Height          =   330
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   960
      Width           =   615
   End
   Begin VB.Timer timerHotkeys 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5520
      Top             =   240
   End
   Begin VB.CommandButton cmdTest2 
      Caption         =   "Test"
      Height          =   375
      Left            =   5280
      TabIndex        =   13
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton cmdTest1 
      Caption         =   "Test"
      Height          =   375
      Left            =   5280
      TabIndex        =   12
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton cmdDel22 
      BackColor       =   &H00808080&
      Caption         =   "X"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1440
      Width           =   255
   End
   Begin VB.TextBox txtKey22 
      Height          =   375
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "<nothing>"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdDel21 
      BackColor       =   &H00808080&
      Caption         =   "X"
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1440
      Width           =   255
   End
   Begin VB.TextBox txtKey21 
      Height          =   375
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "<nothing>"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdDel12 
      BackColor       =   &H00808080&
      Caption         =   "X"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Width           =   255
   End
   Begin VB.TextBox txtKey12 
      Height          =   375
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "<nothing>"
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdDel11 
      BackColor       =   &H00808080&
      Caption         =   "X"
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   255
   End
   Begin VB.TextBox txtKey11 
      Height          =   375
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "<nothing>"
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblFloors 
      BackColor       =   &H00000000&
      Caption         =   "Floors:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   17
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblFloor3 
      BackColor       =   &H00000000&
      Caption         =   "Reset:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblDesc 
      BackColor       =   &H00000000&
      Caption         =   "Click on textboxes and press a key to define a hotkey for virtual floor change (inspect floors above or below you)"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   5175
   End
   Begin VB.Label lblMessage 
      BackColor       =   &H00000000&
      Caption         =   "Warning: Hotkeys are not usable because Directx failed to initialize"
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   5895
   End
   Begin VB.Label lblFloor2 
      BackColor       =   &H00000000&
      Caption         =   "Floor below:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblKey2 
      BackColor       =   &H00000000&
      Caption         =   "Key 2 (optional)"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblKey1 
      BackColor       =   &H00000000&
      Caption         =   "Key 1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblFloor1 
      BackColor       =   &H00000000&
      Caption         =   "Floor above:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "frmXRAY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit


Private meloading As Long
Private meready As Long


Private Sub cmbFloors1_Click()
    If meloading = 1 Then
        XRAY_floors_ABOVE = cmbFloors1.Text
    End If
End Sub

Private Sub cmbFloors2_Click()
    If meloading = 1 Then
        XRAY_floors_BELOW = cmbFloors2.Text
    End If
End Sub

Private Sub cmdDel11_Click()
    espectingHotkey = False
    Me.txtKey11.Text = BString(47)
    Hotkeys(0).key1 = 0
    XRAY_key1_1 = 0
End Sub
Private Sub cmdDel12_Click()
    espectingHotkey = False
    Me.txtKey12.Text = BString(47)
    Hotkeys(0).key2 = 0
    XRAY_key1_2 = 0
End Sub
Private Sub cmdDel21_Click()
    espectingHotkey = False
    Me.txtKey21.Text = BString(47)
    Hotkeys(1).key1 = 0
    XRAY_key2_1 = 0
End Sub
Private Sub cmdDel22_Click()
    espectingHotkey = False
    Me.txtKey22.Text = BString(47)
    Hotkeys(1).key2 = 0
    XRAY_key2_2 = 0
End Sub

Private Sub Xray1(focused As Boolean)
    Dim pid As Long
    Dim tmpLong As Long
    Dim res As Long
    If focused = True Then
        pid = GetFocusedTibiaPID()
    Else
        pid = GetFirstTibiaPID()
    End If
    If ((pid = 0) And (focused = False)) Then
        Me.lblMessage = BString(48)
        Me.lblMessage.Visible = True
    Else
        Me.lblMessage.Visible = False
        tmpLong = Memory_ReadLong(adrConnected, pid)
        If tmpLong >= 8 Then
            ' connected
            res = MemoryChangeFloor(pid, Me.cmbFloors1.Text)
        End If
    End If
End Sub

Private Sub Xray2(focused As Boolean)
    Dim pid As Long
    Dim tmpLong As Long
    Dim res As Long
    If focused = True Then
        pid = GetFocusedTibiaPID()
    Else
        pid = GetFirstTibiaPID()
    End If
    If ((pid = 0) And (focused = False)) Then
        Me.lblMessage = BString(48)
        Me.lblMessage.Visible = True
    Else
        Me.lblMessage.Visible = False
        tmpLong = Memory_ReadLong(adrConnected, pid)
        If tmpLong >= 8 Then
            ' connected
            res = MemoryChangeFloor(pid, Me.cmbFloors2.Text)
        End If
    End If
End Sub

Private Sub Xray3(focused As Boolean)
    Dim pid As Long
    Dim tmpLong As Long
    Dim res As Long
    If focused = True Then
        pid = GetFocusedTibiaPID()
    Else
        pid = GetFirstTibiaPID()
    End If
    If ((pid = 0) And (focused = False)) Then
        Me.lblMessage = BString(48)
        Me.lblMessage.Visible = True
    Else
        Me.lblMessage.Visible = False
        tmpLong = Memory_ReadLong(adrConnected, pid)
        If tmpLong >= 8 Then
            ' connected
            res = MemoryChangeFloor(pid, Me.txtFloors3.Text)
        End If
    End If
End Sub

Private Sub cmdDel31_Click()
    espectingHotkey = False
    Me.txtKey31.Text = BString(47)
    Hotkeys(2).key1 = 0
    XRAY_key3_1 = 0
End Sub

Private Sub cmdDel32_Click()
    espectingHotkey = False
    Me.txtKey32.Text = BString(47)
    Hotkeys(2).key2 = 0
    XRAY_key3_2 = 0
End Sub

Private Sub cmdTest1_Click()
    Xray1 False
End Sub

Private Sub cmdTest2_Click()
    Xray2 False
End Sub

Private Sub cmdTest3_Click()
    Xray3 False
End Sub

Public Sub UpdateXRAY_Language()
   lblDesc.Caption = BString(49)
   lblKey1.Caption = BString(50)
   lblKey2.Caption = BString(51)
   lblFloors.Caption = BString(52)
   lblFloor1.Caption = BString(53)
   lblFloor2.Caption = BString(54)
   lblFloor3.Caption = BString(55)
   cmdTest1.Caption = BString(56)
   cmdTest2.Caption = BString(56)
   cmdTest3.Caption = BString(56)
   Me.Caption = BString(58)
End Sub
Private Sub Form_Load()
    Dim sRes As String
    meloading = 0
    meready = 0
    espectingHotkey = False

    Hotkeys(0).command = "xray1"

    Hotkeys(1).command = "xray2"

    Hotkeys(2).command = "xray3"
    
    Me.cmbFloors1.Clear
    Me.cmbFloors1.AddItem "-1"
    Me.cmbFloors1.AddItem "-2"
    Me.cmbFloors1.AddItem "-3"
    Me.cmbFloors1.AddItem "-4"
    Me.cmbFloors1.AddItem "-5"
    Me.cmbFloors1.AddItem "-6"
    Me.cmbFloors1.AddItem "-7"
    Me.cmbFloors1.Text = "-1"
    Me.cmbFloors2.Clear
    Me.cmbFloors2.AddItem "1"
    Me.cmbFloors2.AddItem "2"
    Me.cmbFloors2.Text = "1"
  sRes = InitDI()
    meloading = 1
  If sRes = "" Then
    lblMessage.Visible = False
    timerHotkeys.Enabled = True
  Else
    lblMessage.Visible = True
    lblMessage.Caption = "Failed to intialize. Hotkeys disabled."
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MustUnload = False Then
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub ClearKeys()
    Dim iRes As Integer
    Dim i As Long
    For i = 0 To 255
        iRes = GetAsyncKeyState(i)
    Next i
End Sub

Private Sub timerHotkeys_Timer()
    Dim i As Long
    Dim iRes As Integer
    Dim pressed As Byte
    Dim something As Boolean
    Dim pressedKeys(0 To 255) As Boolean
    If ((HotkeysAreUsable = True) And (meready = 1)) Then
        For i = 0 To 255
            If GetAsyncKeyState(i) Then
                pressedKeys(i) = True
            Else
                pressedKeys(i) = False
            End If
        Next i
        pressedKeys(0) = True
        pressedKeys(255) = False
        If espectingHotkey = True Then
            pressed = 255
            something = False
            For i = 0 To 255
              Select Case i
              Case 0, 1, 144, 255
              
              Case Else
                    'If GetAsyncKeyState(i) And KEY_PRESSED Thenç
                    If pressedKeys(i) Then
                        pressed = i
                        something = True
                        Exit For
                    End If
                End Select
            Next i
            If something = True Then
                If espectingHotkey = True Then
                    espectingHotkey = False
                    If DefiningHotkeySub = 0 Then
                        Hotkeys(DefiningHotkeyLine).key1 = pressed
                    Else
                        Hotkeys(DefiningHotkeyLine).key2 = pressed
                    End If
                    If ((DefiningHotkeyLine = 0) And (DefiningHotkeySub = 0)) Then
                        Me.txtKey11.Text = TranslateHotkeyID2(pressed)
                        XRAY_key1_1 = pressed
                    End If
                    If ((DefiningHotkeyLine = 0) And (DefiningHotkeySub = 1)) Then
                        Me.txtKey12.Text = TranslateHotkeyID2(pressed)
                        XRAY_key1_2 = pressed
                    End If
                    If ((DefiningHotkeyLine = 1) And (DefiningHotkeySub = 0)) Then
                        Me.txtKey21.Text = TranslateHotkeyID2(pressed)
                        XRAY_key2_1 = pressed
                    End If
                    If ((DefiningHotkeyLine = 1) And (DefiningHotkeySub = 1)) Then
                        Me.txtKey22.Text = TranslateHotkeyID2(pressed)
                        XRAY_key2_2 = pressed
                    End If
                    If ((DefiningHotkeyLine = 2) And (DefiningHotkeySub = 0)) Then
                        Me.txtKey31.Text = TranslateHotkeyID2(pressed)
                        XRAY_key3_1 = pressed
                    End If
                    If ((DefiningHotkeyLine = 2) And (DefiningHotkeySub = 1)) Then
                        Me.txtKey32.Text = TranslateHotkeyID2(pressed)
                        XRAY_key3_2 = pressed
                    End If
                    'MsgBox TranslateHotkeyID2(pressed)
                End If
            End If
        Else
            For i = 0 To lastHotkey
                If (Not ((Hotkeys(i).key1 = 0) And (Hotkeys(i).key2 = 0))) Then
                    If (pressedKeys(Hotkeys(i).key1) And pressedKeys(Hotkeys(i).key2)) Then
                        Select Case Hotkeys(i).command
                        Case "xray1"
                            Xray1 True
                        Case "xray2"
                            Xray2 True
                        Case "xray3"
                            Xray3 True
                        End Select
                        Exit Sub
                    End If
                End If
                
            Next i
        End If
    End If
End Sub

Public Sub ShowCurrentKeys()
    Me.txtKey11.Text = TranslateHotkeyID2(Hotkeys(0).key1)
    Me.txtKey12.Text = TranslateHotkeyID2(Hotkeys(0).key2)
    Me.txtKey21.Text = TranslateHotkeyID2(Hotkeys(1).key1)
    Me.txtKey22.Text = TranslateHotkeyID2(Hotkeys(1).key2)
    Me.txtKey31.Text = TranslateHotkeyID2(Hotkeys(2).key1)
    Me.txtKey32.Text = TranslateHotkeyID2(Hotkeys(2).key2)
    Me.cmbFloors1.Text = XRAY_floors_ABOVE
    Me.cmbFloors2.Text = XRAY_floors_BELOW
    meready = 1
End Sub

Private Sub txtKey11_Click()
    ClearKeys
    espectingHotkey = True
    ShowCurrentKeys
    Me.txtKey11.Text = BString(60)
    DefiningHotkeyLine = 0
    DefiningHotkeySub = 0
End Sub


Private Sub txtKey12_Click()
    ClearKeys
    espectingHotkey = True
    ShowCurrentKeys
    Me.txtKey12.Text = BString(60)
    DefiningHotkeyLine = 0
    DefiningHotkeySub = 1
End Sub

Private Sub txtKey21_Click()
    ClearKeys
    espectingHotkey = True
    ShowCurrentKeys
    Me.txtKey21.Text = BString(60)
    DefiningHotkeyLine = 1
    DefiningHotkeySub = 0
End Sub

Private Sub txtKey22_Click()
    ClearKeys
    espectingHotkey = True
    ShowCurrentKeys
    Me.txtKey22.Text = BString(60)
    DefiningHotkeyLine = 1
    DefiningHotkeySub = 1
End Sub

Private Sub txtKey31_Click()
    ClearKeys
    espectingHotkey = True
    ShowCurrentKeys
    Me.txtKey31.Text = BString(60)
    DefiningHotkeyLine = 2
    DefiningHotkeySub = 0
End Sub

Private Sub txtKey32_Click()
    ClearKeys
    espectingHotkey = True
    ShowCurrentKeys
    Me.txtKey32.Text = BString(60)
    DefiningHotkeyLine = 2
    DefiningHotkeySub = 1
End Sub
