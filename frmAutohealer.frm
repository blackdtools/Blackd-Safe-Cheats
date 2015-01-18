VERSION 5.00
Begin VB.Form frmMenuAutohealer 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HP and Mana"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5295
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmAutohealer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4590
   ScaleWidth      =   5295
   Begin VB.Timer timerHPmana 
      Interval        =   10
      Left            =   4440
      Top             =   480
   End
   Begin VB.Frame fraMana 
      BackColor       =   &H00808000&
      Caption         =   "Mana recharge"
      Height          =   1815
      Left            =   0
      TabIndex        =   8
      Top             =   2760
      Width           =   5295
      Begin VB.ComboBox cmbHotkey 
         Height          =   315
         Index           =   5
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ComboBox cmbLimit 
         Height          =   315
         Index           =   5
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1320
         Width           =   735
      End
      Begin VB.ComboBox cmbHotkey 
         Height          =   315
         Index           =   4
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox cmbLimit 
         Height          =   315
         Index           =   4
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   840
         Width           =   735
      End
      Begin VB.ComboBox cmbLimit 
         Height          =   315
         Index           =   3
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   360
         Width           =   735
      End
      Begin VB.ComboBox cmbHotkey 
         Height          =   315
         Index           =   3
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblThen 
         BackColor       =   &H00808000&
         Caption         =   "then press key"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   2520
         TabIndex        =   26
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblIf 
         BackColor       =   &H00808000&
         Caption         =   "If Mana is less than"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   25
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblThen 
         BackColor       =   &H00808000&
         Caption         =   "then press key"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   2520
         TabIndex        =   22
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblIf 
         BackColor       =   &H00808000&
         Caption         =   "If Mana is less than"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblIf 
         BackColor       =   &H00808000&
         Caption         =   "If Mana is less than"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblThen 
         BackColor       =   &H00808000&
         Caption         =   "then press key"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   2520
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame fraHP 
      BackColor       =   &H00000080&
      Caption         =   "HP recharge"
      Height          =   1815
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   5295
      Begin VB.ComboBox cmbHotkey 
         Height          =   315
         Index           =   2
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ComboBox cmbLimit 
         Height          =   315
         Index           =   2
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1320
         Width           =   735
      End
      Begin VB.ComboBox cmbLimit 
         Height          =   330
         Index           =   1
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   840
         Width           =   735
      End
      Begin VB.ComboBox cmbLimit 
         Height          =   330
         Index           =   0
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   360
         Width           =   735
      End
      Begin VB.ComboBox cmbHotkey 
         Height          =   330
         Index           =   1
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox cmbHotkey 
         Height          =   330
         Index           =   0
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblThen 
         BackColor       =   &H00000080&
         Caption         =   "then press key"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   18
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblIf 
         BackColor       =   &H00000080&
         Caption         =   "If HP is less than"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblThen 
         BackColor       =   &H00000080&
         Caption         =   "then press key"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   6
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblIf 
         BackColor       =   &H00000080&
         Caption         =   "If HP is less than"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblThen 
         BackColor       =   &H00000080&
         Caption         =   "then press key"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   2520
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblIf 
         BackColor       =   &H00000080&
         Caption         =   "If HP is less than"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Label lblMainLabel 
      BackColor       =   &H00000000&
      Caption         =   $"frmAutohealer.frx":0442
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frmMenuAutohealer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit

Private TimeNextRecharge(0 To 5) As Long


Private Sub Form_Load()
    Dim i As Integer
    MyHP = 0
    MyMaxHP = 0
    MyMana = 0
    MyMaxMana = 0
    MySoul = 0
    MyHPpercent = 255
    Mymanapercent = 255
    For i = 0 To 5
        TimeNextRecharge(i) = 0
        Me.cmbLimit(i).Clear
        Me.cmbLimit(i).AddItem "100"
        Me.cmbLimit(i).AddItem "95"
        Me.cmbLimit(i).AddItem "90"
        Me.cmbLimit(i).AddItem "85"
        Me.cmbLimit(i).AddItem "80"
        Me.cmbLimit(i).AddItem "75"
        Me.cmbLimit(i).AddItem "70"
        Me.cmbLimit(i).AddItem "65"
        Me.cmbLimit(i).AddItem "60"
        Me.cmbLimit(i).AddItem "55"
        Me.cmbLimit(i).AddItem "50"
        Me.cmbLimit(i).AddItem "45"
        Me.cmbLimit(i).AddItem "40"
        Me.cmbLimit(i).AddItem "35"
        Me.cmbLimit(i).AddItem "30"
        Me.cmbLimit(i).AddItem "25"
        Me.cmbLimit(i).AddItem "20"
        Me.cmbLimit(i).AddItem "15"
        Me.cmbLimit(i).AddItem "10"
        Me.cmbLimit(i).AddItem "5"
        Me.cmbLimit(i).AddItem "0"
        Me.cmbLimit(i).Text = "0"
        Me.cmbHotkey(i).Clear
        Me.cmbHotkey(i).AddItem "F1"
        Me.cmbHotkey(i).AddItem "F2"
        Me.cmbHotkey(i).AddItem "F3"
        Me.cmbHotkey(i).AddItem "F4"
        Me.cmbHotkey(i).AddItem "F5"
        Me.cmbHotkey(i).AddItem "F6"
        Me.cmbHotkey(i).AddItem "F7"
        Me.cmbHotkey(i).AddItem "F8"
        Me.cmbHotkey(i).AddItem "F9"
        Me.cmbHotkey(i).AddItem "F10"
        Me.cmbHotkey(i).AddItem "F11"
        Me.cmbHotkey(i).AddItem "F12"
        Me.cmbHotkey(i).AddItem "--"
        Me.cmbHotkey(i).Text = "--"
    Next i
End Sub

Private Sub WriteMyBattlelistInfo(pid As Long)
    ' for debug
    Dim bPos As Long
    Dim i As Long
    Dim limitI As Long
    Dim strAll As String
    Dim bRes As Byte
    strAll = "Battlelist:"
    limitI = CLng(CharDist)
    bPos = MyBattleListPositionByPID(pid)
    If bPos >= 0 Then
        For i = 0 To limitI
            bRes = GetElementFromBattleListPos(pid, bPos, i)
            strAll = strAll & " " & GoodHex(bRes)
        Next i
    End If
    'Debug.Print strAll
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MustUnload = False Then
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub TestRechargeVars()
    Dim limitR(0 To 5) As Long
    Dim i As Long
    Dim gtc As Long
    Dim rFactor As Long
    Dim strAction As String
    Dim theAction As String
    Dim theLevel As Long
    Dim bestHP As Long
    Dim bestMANA As Long
    Dim am As Long
    gtc = GetTickCount()
    am = 150
    For i = 0 To 2
        limitR(i) = SafeLong(Me.cmbLimit(i).Text)
        If (MyHPpercent <= limitR(i)) And (limitR(i) > 0) Then
        If Me.cmbHotkey(i).Text <> "--" Then
            If limitR(i) < am Then
                bestHP = i
                am = limitR(i)
            End If
        End If
        End If
    Next i
    am = 150
    For i = 3 To 5
        limitR(i) = SafeLong(Me.cmbLimit(i).Text)
        If (Mymanapercent <= limitR(i)) And (limitR(i) > 0) Then
        If Me.cmbHotkey(i).Text <> "--" Then
            If limitR(i) < am Then
                bestMANA = i
                am = limitR(i)
            End If
        End If
        End If
    Next i
    
    
    For i = 0 To 2 ' hp
        If i = bestHP Then
            limitR(i) = SafeLong(Me.cmbLimit(i).Text)
            If (MyHPpercent <= limitR(i)) And (limitR(i) > 0) Then
                If Me.cmbHotkey(i).Text <> "--" Then
                    If TimeNextRecharge(i) = 0 Then
                        rFactor = CLng((HPmanadelay1 * HPrandpercent) / 100)
                        TimeNextRecharge(i) = gtc + randomNumberBetween(MaxV(0, HPmanadelay1 - rFactor), HPmanadelay1 + rFactor)
                    End If
                End If
            End If
        End If
    Next i
    For i = 3 To 5 ' mana
        If i = bestMANA Then
            limitR(i) = SafeLong(Me.cmbLimit(i).Text)
            If (Mymanapercent <= limitR(i)) And (limitR(i) > 0) Then
                If Me.cmbHotkey(i).Text <> "--" Then
                    If TimeNextRecharge(i) = 0 Then
                        rFactor = CLng((HPmanadelay1 * HPrandpercent) / 100)
                        TimeNextRecharge(i) = gtc + randomNumberBetween(MaxV(0, HPmanadelay1 - rFactor), HPmanadelay1 + rFactor)
                    End If
                End If
            End If
        End If
    Next i
    
    theAction = ""
    theLevel = 150
    For i = 0 To 2
        If i = bestHP Then
        If (TimeNextRecharge(i) <= gtc) And (TimeNextRecharge(i) > 0) Then
            If (MyHPpercent <= limitR(i)) And (limitR(i) > 0) Then
                If Me.cmbHotkey(i).Text <> "--" Then
                    rFactor = CLng((HPmanadelay2 * HPrandpercent) / 100)
                    TimeNextRecharge(i) = gtc + randomNumberBetween(MaxV(0, HPmanadelay2 - rFactor), HPmanadelay2 + rFactor)
                    strAction = "{" & Me.cmbHotkey(i).Text & "}"
                    'Debug.Print "[" & gtc & "] action " & i
                    If limitR(i) < theLevel Then
                        theAction = strAction
                        theLevel = limitR(i)
                    End If
                End If
            Else
                TimeNextRecharge(i) = 0
            End If
        End If
        End If
    Next i
    If theAction <> "" Then
        SendKeys theAction
        DoEvents
        Exit Sub
    End If
    
    theAction = ""
    theLevel = 150
    For i = 3 To 5
        If i = bestMANA Then
        If (TimeNextRecharge(i) <= gtc) And (TimeNextRecharge(i) > 0) Then
            If (Mymanapercent <= limitR(i)) And (limitR(i) > 0) Then
                If Me.cmbHotkey(i).Text <> "--" Then
                    rFactor = CLng((HPmanadelay2 * HPrandpercent) / 100)
                    TimeNextRecharge(i) = gtc + randomNumberBetween(MaxV(0, HPmanadelay2 - rFactor), HPmanadelay2 + rFactor)
                    strAction = "{" & Me.cmbHotkey(i).Text & "}"
                    'Debug.Print "[" & gtc & "] action " & i
                    If limitR(i) < theLevel Then
                        theAction = strAction
                        theLevel = limitR(i)
                    End If
                End If
            Else
                TimeNextRecharge(i) = 0
            End If
        End If
        End If
    Next i
    
    If theAction <> "" Then
        SendKeys theAction
        DoEvents
        Exit Sub
    End If
End Sub


Private Sub timerHPmana_Timer()
    On Error GoTo goterr
    Dim lngTibiaPID As Long
    Dim tmpLong As Long
    Dim valueXOR As Long
    lngTibiaPID = GetFocusedTibiaPID()
    If lngTibiaPID <> 0 Then
        tmpLong = Memory_ReadLong(adrConnected, lngTibiaPID)
        If tmpLong >= 8 Then
            TibiaIsConnected = True
            MyHP = Memory_ReadLong(adrMyHP, lngTibiaPID)
            MyMaxHP = Memory_ReadLong(adrMyMaxHP, lngTibiaPID)
            MyMana = Memory_ReadLong(adrMyMana, lngTibiaPID)
            MyMaxMana = Memory_ReadLong(adrMyMaxMana, lngTibiaPID)
            MySoul = Memory_ReadLong(adrMySoul, lngTibiaPID)
            If TibiaVersionLong >= 943 Then
                valueXOR = Memory_ReadLong(adrXOR, lngTibiaPID)
                MyHP = valueXOR Xor MyHP
                MyMaxHP = valueXOR Xor MyMaxHP
                MyMana = valueXOR Xor MyMana
                MyMaxMana = valueXOR Xor MyMaxMana
                
                ' Soul points does not require XOR
                ' Debug.Print "DEBUG MySoul=" & (MySoul)
            End If
            If MyMaxHP = 0 Then
                MyHPpercent = 255
            Else
                MyHPpercent = CLng(Round(CDbl(MyHP * 100) / CDbl(MyMaxHP)))
            End If
            If MyMaxMana = 0 Then
                Mymanapercent = 255
            Else
                Mymanapercent = CLng(Round(CDbl(MyMana * 100) / CDbl(MyMaxMana)))
            End If
            TestRechargeVars
        Else
            TibiaIsConnected = False
        End If
    End If
    Exit Sub
goterr:
    MyHP = 0
    MyMaxHP = 0
    MyMana = 0
    MyMaxMana = 0
End Sub
