VERSION 5.00
Begin VB.Form frmLight 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Light"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4545
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmLight.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1410
   ScaleWidth      =   4545
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3840
      Top             =   120
   End
   Begin VB.HScrollBar scrollColour 
      Height          =   255
      Left            =   1440
      Max             =   255
      TabIndex        =   3
      Top             =   960
      Value           =   215
      Width           =   2055
   End
   Begin VB.HScrollBar scrollLevel 
      Height          =   255
      Left            =   1440
      Max             =   15
      TabIndex        =   1
      Top             =   600
      Value           =   15
      Width           =   2055
   End
   Begin VB.CheckBox chkEnable 
      BackColor       =   &H00000000&
      Caption         =   "Enable light in focused Tibia client"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label lblLevel 
      BackColor       =   &H00000000&
      Caption         =   "100 %"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   6
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblColour 
      BackColor       =   &H00000000&
      Caption         =   "215"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblColourID 
      BackColor       =   &H00000000&
      Caption         =   "Light colour ID:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblIntensity 
      BackColor       =   &H00000000&
      Caption         =   "Light level:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "frmLight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit


Public Sub UpdateLight()
    Dim tibiaclient As Long
    Dim hWndDesktop As Long
    Dim i As Integer
    Dim j As Integer
    Dim num As Integer
    Dim num2 As Integer
    Dim posM As Integer
    Dim writeStr As String
    Dim writeChr As String
    Dim myID As Long
    Dim tmpID As Long
    Dim bPos As Long
    Dim lastPos As Long
    Dim lngTibiaPID As Long
    Dim tmpLong As Long
    Dim b1 As Byte
    Dim b2 As Byte
    Dim zz As Long
    
    TibiaIsConnected = False
    lngTibiaPID = GetFocusedTibiaPID()
    If lngTibiaPID <> 0 Then
        tmpLong = Memory_ReadLong(adrConnected, lngTibiaPID)
        If tmpLong >= 8 Then
            TibiaIsConnected = True
        Else
            TibiaIsConnected = False
        End If
    End If
    If TibiaIsConnected = True Then
        tibiaclient = lngTibiaPID
        myID = Memory_ReadLong(adrNum, tibiaclient)
       ' Debug.Print ("HELMET 10.82 =" & Memory_ReadLong(&HB72DE0, tibiaclient))
        
        lastPos = -1
        lastPos = MyBattleListPositionByPID(tibiaclient)
      
        'b1 = Memory_ReadByte((adrNChar + (lastPos * CharDist) + LightDist), tibiaclient)
        'b2 = Memory_ReadByte((adrNChar + (lastPos * CharDist) + LightColourDist), tibiaclient)
        
       ' Debug.Print Hex(b1) & " " & _
                    Hex(b2)
        Memory_WriteByte (adrNChar + (lastPos * CharDist) + LightDist), LightIntensity, tibiaclient
        Memory_WriteByte (adrNChar + (lastPos * CharDist) + LightColourDist), LightColour, tibiaclient
    End If
End Sub


Private Sub chkEnable_Click()
    If chkEnable.Value = 1 Then
      LightEnabled = True
    Else
      LightEnabled = False
    End If
End Sub

Private Sub Form_Load()
    LightColour = 215
    LightIntensity = 15
    If (LightEnabled = True) Then
      chkEnable.Value = 1

    Else
      chkEnable.Value = 0
    End If
    LightRefreshDelay = 1000
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MustUnload = False Then
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub Label2_Click()

End Sub

Private Sub scrollColour_Change()
    lblColour.Caption = CStr(scrollColour.Value)
    LightColour = CByte(scrollColour.Value)
End Sub

Private Sub scrollLevel_Change()
    lblLevel.Caption = CStr(Round((scrollLevel.Value / 15) * 100)) & " %"
    LightIntensity = CByte(scrollLevel.Value)
End Sub

Private Sub Timer1_Timer()
    If LightEnabled = True Then
        UpdateLight
    End If
End Sub

Public Sub UpdateControlValues()
    Dim lngLev As Long
    If LightEnabled = True Then
        Me.chkEnable.Value = 1
    Else
        Me.chkEnable.Value = 0
    End If
    scrollLevel.Value = LightIntensity
    lngLev = Round(LightColour)
    scrollColour.Value = lngLev
    lblColour.Caption = CStr(lngLev)
    Timer1.Enabled = False
    Timer1.Interval = LightRefreshDelay
    Timer1.Enabled = True
    
    chkEnable.Caption = BString(44)
    lblIntensity.Caption = BString(45)
    lblColourID.Caption = BString(46)
    
End Sub
