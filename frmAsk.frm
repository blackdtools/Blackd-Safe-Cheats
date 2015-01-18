VERSION 5.00
Begin VB.Form frmAsk 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Define the value..."
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4800
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmAsk.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1470
   ScaleWidth      =   4800
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Save"
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtValue 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblResult 
      BackColor       =   &H00000000&
      Caption         =   "---"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblQuestion 
      BackColor       =   &H00000000&
      Caption         =   "Question"
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmAsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Public Sub UpdateQuestion()
    '...
    lblResult.Caption = ""
    Me.Caption = BString(25)
    Me.cmdSave.Caption = BString(35)
    Me.cmdCancel.Caption = BString(36)
    
    Select Case Now_we_define
    Case "HPmanadelay1"
        Me.lblQuestion.Caption = BString(26)
        Me.txtValue.Text = CStr(HPmanadelay1)
    Case "HPmanadelay2"
        Me.lblQuestion.Caption = BString(27)
        Me.txtValue.Text = CStr(HPmanadelay2)
    Case "HPrandpercent"
        Me.lblQuestion.Caption = BString(28)
        Me.txtValue.Text = CStr(HPrandpercent)
    Case "LightRefreshDelay"
        Me.lblQuestion.Caption = BString(34)
        Me.txtValue.Text = CStr(LightRefreshDelay)
    Case Else
        Me.lblQuestion.Caption = "Error: undefined question"
        Me.txtValue.Text = "ERROR"
    End Select
End Sub


Private Sub cmdSave_Click()
    Dim strRes As String
    Dim lngRes As Long
    strRes = Me.txtValue.Text
    Select Case Now_we_define
    Case "HPmanadelay1"
        If IsNumeric(strRes) = False Then
            lblResult.Caption = "Must be numeric"
            Exit Sub
        End If
        lngRes = SafeLong(strRes)
        If lngRes < 0 Then
            lblResult.Caption = "Must be positive"
            Exit Sub
        End If
        If lngRes > 10000 Then
            lblResult.Caption = "Too high value"
            Exit Sub
        End If
        HPmanadelay1 = lngRes
        Me.Hide
    Case "HPmanadelay2"
        If IsNumeric(strRes) = False Then
            lblResult.Caption = "Must be numeric"
            Exit Sub
        End If
        lngRes = SafeLong(strRes)
        If lngRes < 0 Then
            lblResult.Caption = "Must be positive"
            Exit Sub
        End If
        If lngRes > 10000 Then
            lblResult.Caption = "Too high value"
            Exit Sub
        End If
        HPmanadelay2 = lngRes
        Me.Hide
    Case "HPrandpercent"
        If IsNumeric(strRes) = False Then
            lblResult.Caption = "Must be numeric"
            Exit Sub
        End If
        lngRes = SafeLong(strRes)
        If lngRes < 0 Then
            lblResult.Caption = "Must be positive"
            Exit Sub
        End If
        If lngRes > 50 Then
            lblResult.Caption = "Too high value"
            Exit Sub
        End If
        HPrandpercent = lngRes
        Me.Hide
    Case "LightRefreshDelay"
        If IsNumeric(strRes) = False Then
            lblResult.Caption = "Must be numeric"
            Exit Sub
        End If
        lngRes = SafeLong(strRes)
        If lngRes < 0 Then
            lblResult.Caption = "Must be positive"
            Exit Sub
        End If
        If lngRes < 10 Then
            lblResult.Caption = "Too low value"
            Exit Sub
        End If
        If lngRes > 60000 Then
            lblResult.Caption = "Too high value"
            Exit Sub
        End If
        LightRefreshDelay = lngRes
        frmLight.UpdateControlValues
        Me.Hide
    Case Else
        Me.Hide
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MustUnload = False Then
        Cancel = True
        Me.Hide
    End If
End Sub
