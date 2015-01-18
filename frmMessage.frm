VERSION 5.00
Begin VB.Form frmMessage 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Information"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmMessage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ok"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit

Public Sub UpdateDisplay()
    Me.Caption = Message_Tittle
    Me.txtInfo.Locked = False
    Me.txtInfo.Text = Message_Message
    Me.txtInfo.Locked = True
End Sub
Private Sub cmdOk_Click()
    Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MustUnload = False Then
        Cancel = True
        Me.Hide
    End If
End Sub
