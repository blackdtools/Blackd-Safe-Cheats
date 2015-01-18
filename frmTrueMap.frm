VERSION 5.00
Begin VB.Form frmTrue 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Truemap"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4650
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmTrueMap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4335
   ScaleWidth      =   4650
   Begin VB.CommandButton cmdUpdateMyFloor 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Watch my floor"
      Height          =   255
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   270
      Top             =   50
      Width           =   2055
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Watch selected floor"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   269
      Top             =   50
      Width           =   2055
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   251
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   268
      Top             =   3480
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   250
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   267
      Top             =   3480
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   249
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   266
      Top             =   3480
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   248
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   265
      Top             =   3480
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   247
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   264
      Top             =   3480
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   246
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   263
      Top             =   3480
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   245
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   262
      Top             =   3480
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   244
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   261
      Top             =   3480
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   243
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   260
      Top             =   3480
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   242
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   259
      Top             =   3480
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   241
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   258
      Top             =   3480
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   240
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   257
      Top             =   3480
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   239
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   256
      Top             =   3480
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   238
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   255
      Top             =   3480
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   237
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   254
      Top             =   3480
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   236
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   253
      Top             =   3480
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   235
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   252
      Top             =   3480
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   234
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   251
      Top             =   3480
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   233
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   250
      Top             =   3240
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   232
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   249
      Top             =   3240
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   231
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   248
      Top             =   3240
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   230
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   247
      Top             =   3240
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   229
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   246
      Top             =   3240
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   228
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   245
      Top             =   3240
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   227
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   244
      Top             =   3240
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   226
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   243
      Top             =   3240
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   225
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   242
      Top             =   3240
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   224
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   241
      Top             =   3240
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   223
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   240
      Top             =   3240
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   222
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   239
      Top             =   3240
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   221
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   238
      Top             =   3240
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   220
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   237
      Top             =   3240
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   219
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   236
      Top             =   3240
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   218
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   235
      Top             =   3240
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   217
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   234
      Top             =   3240
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   216
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   233
      Top             =   3240
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   215
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   232
      Top             =   3000
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   214
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   231
      Top             =   3000
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   213
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   230
      Top             =   3000
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   212
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   229
      Top             =   3000
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   211
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   228
      Top             =   3000
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   210
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   227
      Top             =   3000
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   209
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   226
      Top             =   3000
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   208
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   225
      Top             =   3000
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   207
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   224
      Top             =   3000
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   206
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   223
      Top             =   3000
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   205
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   222
      Top             =   3000
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   204
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   221
      Top             =   3000
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   203
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   220
      Top             =   3000
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   202
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   219
      Top             =   3000
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   201
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   218
      Top             =   3000
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   200
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   217
      Top             =   3000
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   199
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   216
      Top             =   3000
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   198
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   215
      Top             =   3000
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   197
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   214
      Top             =   2760
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   196
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   213
      Top             =   2760
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   195
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   212
      Top             =   2760
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   194
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   211
      Top             =   2760
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   193
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   210
      Top             =   2760
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   192
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   209
      Top             =   2760
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   191
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   208
      Top             =   2760
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   190
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   207
      Top             =   2760
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   189
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   206
      Top             =   2760
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   188
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   205
      Top             =   2760
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   187
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   204
      Top             =   2760
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   186
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   203
      Top             =   2760
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   185
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   202
      Top             =   2760
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   184
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   201
      Top             =   2760
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   183
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   200
      Top             =   2760
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   182
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   199
      Top             =   2760
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   181
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   198
      Top             =   2760
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   180
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   197
      Top             =   2760
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   179
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   196
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   178
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   195
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   177
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   194
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   176
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   193
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   175
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   192
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   174
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   191
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   173
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   190
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   172
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   189
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   171
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   188
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   170
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   187
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   169
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   186
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   168
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   185
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   167
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   184
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   166
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   183
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   165
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   182
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   164
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   181
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   163
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   180
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   162
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   179
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   161
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   178
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   160
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   177
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   159
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   176
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   158
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   175
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   157
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   174
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   156
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   173
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   155
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   172
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   154
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   171
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   153
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   170
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   152
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   169
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   151
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   168
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   150
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   167
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   149
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   166
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   148
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   165
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   147
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   164
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   146
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   163
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   145
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   162
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   144
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   161
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   143
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   160
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   142
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   159
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   141
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   158
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   140
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   157
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   139
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   156
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   138
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   155
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   137
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   154
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   136
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   153
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   135
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   152
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   134
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   151
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   133
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   150
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   132
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   149
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   131
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   148
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   130
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   147
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   129
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   146
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   128
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   145
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   127
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   144
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   126
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   143
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   125
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   142
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   124
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   141
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   123
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   140
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   122
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   139
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   121
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   138
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   120
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   137
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   119
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   136
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   118
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   135
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   117
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   134
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   116
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   133
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   115
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   132
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   114
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   131
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   113
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   130
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   112
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   129
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   111
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   128
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   110
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   127
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   109
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   126
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   108
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   125
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   107
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   124
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   106
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   123
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   105
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   122
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   104
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   121
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   103
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   120
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   102
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   119
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   101
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   118
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   100
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   117
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   99
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   116
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   98
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   115
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   97
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   114
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   96
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   113
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   95
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   112
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   94
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   111
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   93
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   110
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   92
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   109
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   91
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   108
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   90
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   107
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   89
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   106
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   88
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   105
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   87
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   104
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   86
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   103
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   85
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   102
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   84
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   101
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   83
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   100
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   82
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   99
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   81
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   98
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   80
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   79
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   96
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   78
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   77
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   76
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   75
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   74
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   91
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   73
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   72
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   71
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   70
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   87
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   69
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   86
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   68
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   67
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   66
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   65
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   64
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   63
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   62
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   61
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   60
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   59
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   58
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   57
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   56
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   55
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   54
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   53
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   52
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   51
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   50
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   49
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   48
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   47
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   46
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   45
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   44
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   43
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   42
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   41
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   40
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   39
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   38
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   37
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   36
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   35
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   34
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   33
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   32
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   31
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   30
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   29
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   28
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   27
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   26
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   25
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   24
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   23
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   22
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   21
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   20
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   19
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   18
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdFloor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   360
      Width           =   300
   End
   Begin VB.CommandButton cmdFloor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   555
      Width           =   300
   End
   Begin VB.CommandButton cmdFloor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   2
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   765
      Width           =   300
   End
   Begin VB.CommandButton cmdFloor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   3
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   960
      Width           =   300
   End
   Begin VB.CommandButton cmdFloor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   4
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1155
      Width           =   300
   End
   Begin VB.CommandButton cmdFloor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   5
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1365
      Width           =   300
   End
   Begin VB.CommandButton cmdFloor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   6
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1560
      Width           =   300
   End
   Begin VB.CommandButton cmdFloor 
      BackColor       =   &H00C0FFFF&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   7
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1755
      Width           =   300
   End
   Begin VB.CommandButton cmdFloor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   8
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1965
      Width           =   300
   End
   Begin VB.CommandButton cmdFloor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   9
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      Width           =   300
   End
   Begin VB.CommandButton cmdFloor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   10
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2355
      Width           =   300
   End
   Begin VB.CommandButton cmdFloor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   11
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2565
      Width           =   300
   End
   Begin VB.CommandButton cmdFloor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   12
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   300
   End
   Begin VB.CommandButton cmdFloor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   13
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2955
      Width           =   300
   End
   Begin VB.CommandButton cmdFloor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   14
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3165
      Width           =   300
   End
   Begin VB.CommandButton cmdFloor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   15
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3360
      Width           =   300
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   17
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   16
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   14
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton gridMap 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox txtSelected 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   0
      TabIndex        =   16
      Top             =   3720
      Width           =   4650
   End
End
Attribute VB_Name = "frmTrue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit

Private gridMap_col As Long
Private gridMap_row As Long
Private mapFloorSelected As Long
Private GridCaption(0 To 251) As String
Private CharCount(0 To 251) As Long
'...
Public Sub UpdateLanguage()
   Me.Caption = BString(62)
   Me.cmdUpdate.Caption = BString(63)
   Me.cmdUpdateMyFloor.Caption = BString(64)
End Sub

Public Function ColourPriority(Colour As ColorConstants) As Integer
  Select Case Colour
  Case ColourNothing
    ColourPriority = 0
  Case ColourGround ' ground
    ColourPriority = 1
  Case ColourWater 'water
    ColourPriority = 2
  Case ColourFish ' with fish
    ColourPriority = 3
  Case ColourBlockMoveable ' blocking , but moveable
    ColourPriority = 4
  Case ColourSomething 'blocking + not moveable
    ColourPriority = 5
  Case ColourField 'field
    ColourPriority = 6
  Case ColourDown ' ladder down
    ColourPriority = 7
  Case ColourUp ' ladder up
    ColourPriority = 8
  Case ColourPlayer
    ColourPriority = 50
  Case ColourWithMe
    ColourPriority = 99
  End Select
End Function
Public Sub DrawFloor(Optional selectedFloor = True)
    '....
    Dim drawThisFloor As Long
    Dim tibiaclient As Long
    Dim cellPos As Long
    Dim Z As Long
    Dim PosX As Long
    Dim PosY As Long
    Dim MT As MapTile
    Dim tmpID As Double
    Dim poss As Long
    Dim tileID As Long
    Dim tmpName As String
    Dim withID As Boolean
    Dim theExtra As String
    Dim MTc As Long
    Dim tmpam As Byte
    withID = True
    tibiaclient = 0
    tibiaclient = GetFirstTibiaPID()
    If tibiaclient = 0 Then
        Me.txtSelected.Text = BString(65)
        Exit Sub
    End If
    Z = CLng(Memory_ReadByte(PLAYER_Z, tibiaclient))
    If selectedFloor = True Then
        drawThisFloor = mapFloorSelected
    Else
        drawThisFloor = Z
    End If
    UpdateMap tibiaclient
    Me.AutoRedraw = False
    
    cellPos = 0
    For PosY = -6 To 7
        For PosX = -8 To 9
            MT = GetMapTile(PosX, PosY, drawThisFloor)

            
            
            
            
            
            
            
            
      gridMap(cellPos).BackColor = ColourNothing
      GridCaption(cellPos) = ""
      CharCount(cellPos) = 0
      gridMap(cellPos).Caption = ""
      MTc = MT.count - 1
      If MTc > 9 Then
        MTc = 9
      End If
      For poss = 0 To MTc

        tileID = MT.items(poss).id
        If MT.items(poss).data1 < 255 Then
        tmpam = CByte(MT.items(poss).data1)
        Else
        tmpam = 0
        End If
        
        #If FinalMode = 0 Then
         GridCaption(cellPos) = GridCaption(cellPos) & _
         "[" & GoodHex(LowByteOfLong(tileID)) & " " & GoodHex(HighByteOfLong(tileID)) & " " & GoodHex(tmpam) & "]"
        #End If
        If ((tileID <= 0) Or (tileID > highestDatTile)) Then
           Exit For
        ElseIf tileID = &H63 Then
          gridMap(cellPos).BackColor = ColourPlayer
          If GridCaption(cellPos) = "" Then
          tmpID = CDbl(MT.items(poss).data1)
          If tmpID = 0 Then
            GridCaption(cellPos) = "tileid " & CStr(tileID) & "??"
            CharCount(cellPos) = CharCount(cellPos) + 1
            gridMap(cellPos).Caption = CStr(CharCount(cellPos))
            
          Else
            theExtra = ""
            If withID = True Then
              theExtra = " (" & CStr(tmpID) & ")"
            End If
            GridCaption(cellPos) = GetNameFromID(tibiaclient, tmpID) & theExtra
            CharCount(cellPos) = CharCount(cellPos) + 1
            gridMap(cellPos).Caption = CStr(CharCount(cellPos))
          End If
          Else
            tmpID = CDbl(MT.items(poss).data1)
            If tmpID = 0 Then
              tmpName = "tileid " & CStr(tileID) & "??"
            Else
               tmpName = GetNameFromID(tibiaclient, tmpID)
               
            End If
            theExtra = ""
            If withID = True Then
              theExtra = " (" & CStr(tmpID) & ")"
            End If
            GridCaption(cellPos) = GridCaption(cellPos) & " , " & tmpName & theExtra
            CharCount(cellPos) = CharCount(cellPos) + 1
            gridMap(cellPos).Caption = CStr(CharCount(cellPos))
          End If
           'gotMobiles = True
        ElseIf poss = 0 Then
          If (tileID <> &H0) Then
            gridMap(cellPos).BackColor = ColourGround

            If ColourPriority(ColourUp) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).floorChangeUP Then
                gridMap(cellPos).BackColor = ColourUp
              End If
            End If
            If ColourPriority(ColourDown) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).floorChangeDOWN Then
                gridMap(cellPos).BackColor = ColourDown
              End If
            End If
            
            If ColourPriority(ColourWater) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).isWater = True Then
                gridMap(cellPos).BackColor = ColourWater
                If DatTiles(tileID).haveFish = True And _
                 ColourPriority(ColourFish) > ColourPriority(gridMap(cellPos).BackColor) Then
                  gridMap(cellPos).BackColor = ColourFish
                End If
              End If
            End If
          End If
          If DatTiles(tileID).isWater = False Then
            If ColourPriority(ColourSomething2) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).blocking = True And DatTiles(tileID).notMoveable = False Then
                gridMap(cellPos).BackColor = ColourSomething2
              End If
            End If
            If ColourPriority(ColourSomething) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).blocking = True And DatTiles(tileID).notMoveable = True Then
                gridMap(cellPos).BackColor = ColourSomething
              End If
            End If
          End If
            If ColourPriority(ColourField) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).isField Then
               gridMap(cellPos).BackColor = ColourField
              End If
            End If
        Else
          If ((tileID > 99) And (tileID <= highestDatTile)) Then
            If DatTiles(tileID).blocking = False Then

                If ColourPriority(ColourGround) > ColourPriority(gridMap(cellPos).BackColor) Then
                  gridMap(cellPos).BackColor = ColourGround
                End If
                If ColourPriority(ColourUp) > ColourPriority(gridMap(cellPos).BackColor) Then
                  If DatTiles(tileID).floorChangeUP Then
                    gridMap(cellPos).BackColor = ColourUp
                  End If
                End If
               If ColourPriority(ColourDown) > ColourPriority(gridMap(cellPos).BackColor) Then
                 If DatTiles(tileID).floorChangeDOWN Then
                   gridMap(cellPos).BackColor = ColourDown
                 End If
               End If
               If ColourPriority(ColourField) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).isField Then
                gridMap(cellPos).BackColor = ColourField
              End If
            End If
            Else

                If DatTiles(tileID).notMoveable = True Then
                  ' blocking and not moveable
                  If ColourPriority(ColourSomething) > ColourPriority(gridMap(cellPos).BackColor) Then
                    gridMap(cellPos).BackColor = ColourSomething
                  End If
                  
            If ColourPriority(ColourUp) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).floorChangeUP Then
                gridMap(cellPos).BackColor = ColourUp
              End If
            End If
            If ColourPriority(ColourDown) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).floorChangeDOWN Then
                gridMap(cellPos).BackColor = ColourDown
              End If
            End If
                
                Else ' blocking but moveable
                  If ColourPriority(ColourBlockMoveable) > ColourPriority(gridMap(cellPos).BackColor) Then
                    gridMap(cellPos).BackColor = ColourBlockMoveable
                  End If
                
                         If ColourPriority(ColourUp) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).floorChangeUP Then
                gridMap(cellPos).BackColor = ColourUp
              End If
            End If
            If ColourPriority(ColourDown) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).floorChangeDOWN Then
                gridMap(cellPos).BackColor = ColourDown
              End If
            End If
                
                
                End If
            End If
          ElseIf tileID = 0 Then
             ' end of stack
             Exit For
          End If
        End If
      Next poss
            
            
            
            
            
            
            
            
            
           
            cellPos = cellPos + 1
        Next PosX
    Next PosY
    Me.AutoRedraw = True
    Me.txtSelected.Text = "Done"
End Sub


'Private Sub debugtiles()
'    Dim i As Long
'    For i = 0 To 2015
 '       If MapTiles2(i).count > 0 Then
 '           ' Debug.Print "?"
 '           If MapTiles2(i).items(0).data1 > 0 Then
 '              Debug.Print "!"
 '           End If
'        End If
'    Next i
'End Sub
Public Sub DrawFloor2(Optional selectedFloor = True)
    '....

    Dim drawThisFloor As Long
    Dim tibiaclient As Long
    Dim cellPos As Long
    Dim Z As Long
    Dim PosX As Long
    Dim PosY As Long
    Dim MT As MapTile2
    Dim tmpID As Double
    Dim poss As Long
    Dim tileID As Long
    Dim tmpName As String
    Dim withID As Boolean
    Dim theExtra As String
    Dim MTc As Long
    Dim tmpam As Byte
    withID = True
    tibiaclient = 0
    tibiaclient = GetFirstTibiaPID()
    If tibiaclient = 0 Then
        Me.txtSelected.Text = BString(65)
        Exit Sub
    End If
    Z = CLng(Memory_ReadByte(PLAYER_Z, tibiaclient))
    If selectedFloor = True Then
        drawThisFloor = mapFloorSelected
    Else
        drawThisFloor = Z
    End If
    UpdateMap tibiaclient
    
    
    
    'debugtiles
    'Exit Sub
    
    
    Me.AutoRedraw = False
    
    cellPos = 0
    For PosY = -6 To 7
        For PosX = -8 To 9
            MT = GetMapTile2(PosX, PosY, drawThisFloor)
'            If (PosX = -1) And (PosY = 0) Then
'                'Debug.Print "bien"
'            End If
            
            
            
            
            
            
            
            
            
      gridMap(cellPos).BackColor = ColourNothing
      GridCaption(cellPos) = ""
      CharCount(cellPos) = 0
      gridMap(cellPos).Caption = ""
      MTc = MT.count - 1
      'If MTc > 0 Then
       ' MsgBox "debug"
      'End If
      For poss = 0 To MTc

        tileID = MT.items(poss).id
        If MT.items(poss).data1 < 255 Then
        tmpam = CByte(MT.items(poss).data1)
        Else
        tmpam = 0
        End If
        
        #If FinalMode = 0 Then
         GridCaption(cellPos) = GridCaption(cellPos) & _
         "[" & GoodHex(LowByteOfLong(tileID)) & " " & GoodHex(HighByteOfLong(tileID)) & " " & GoodHex(tmpam) & "]"
        #End If
        If ((tileID <= 0) Or (tileID > highestDatTile)) Then
           Exit For
        ElseIf tileID = &H63 Then
          gridMap(cellPos).BackColor = ColourPlayer
          If GridCaption(cellPos) = "" Then
          tmpID = CDbl(MT.items(poss).data1)
          If tmpID = 0 Then
            GridCaption(cellPos) = "tileid " & CStr(tileID) & "??"
            CharCount(cellPos) = CharCount(cellPos) + 1
            gridMap(cellPos).Caption = CStr(CharCount(cellPos))
            
          Else
            theExtra = ""
            If withID = True Then
              theExtra = " (" & CStr(tmpID) & ")"
            End If
            GridCaption(cellPos) = GetNameFromID(tibiaclient, tmpID) & theExtra
            CharCount(cellPos) = CharCount(cellPos) + 1
            gridMap(cellPos).Caption = CStr(CharCount(cellPos))
          End If
          Else
            tmpID = CDbl(MT.items(poss).data1)
            If tmpID = 0 Then
              tmpName = "tileid " & CStr(tileID) & "??"
            Else
               tmpName = GetNameFromID(tibiaclient, tmpID)
               
            End If
            theExtra = ""
            If withID = True Then
              theExtra = " (" & CStr(tmpID) & ")"
            End If
            GridCaption(cellPos) = GridCaption(cellPos) & " , " & tmpName & theExtra
            CharCount(cellPos) = CharCount(cellPos) + 1
            gridMap(cellPos).Caption = CStr(CharCount(cellPos))
          End If
           'gotMobiles = True
        ElseIf poss = 0 Then
          If tileID <> &H0 Then
            gridMap(cellPos).BackColor = ColourGround

            If ColourPriority(ColourUp) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).floorChangeUP Then
                gridMap(cellPos).BackColor = ColourUp
              End If
            End If
            If ColourPriority(ColourDown) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).floorChangeDOWN Then
                gridMap(cellPos).BackColor = ColourDown
              End If
            End If
            
            If ColourPriority(ColourWater) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).isWater = True Then
                gridMap(cellPos).BackColor = ColourWater
                If DatTiles(tileID).haveFish = True And _
                 ColourPriority(ColourFish) > ColourPriority(gridMap(cellPos).BackColor) Then
                  gridMap(cellPos).BackColor = ColourFish
                End If
              End If
            End If
          End If
          If DatTiles(tileID).isWater = False Then
            If ColourPriority(ColourSomething2) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).blocking = True And DatTiles(tileID).notMoveable = False Then
                gridMap(cellPos).BackColor = ColourSomething2
              End If
            End If
            If ColourPriority(ColourSomething) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).blocking = True And DatTiles(tileID).notMoveable = True Then
                gridMap(cellPos).BackColor = ColourSomething
              End If
            End If
          End If
            If ColourPriority(ColourField) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).isField Then
               gridMap(cellPos).BackColor = ColourField
              End If
            End If
        Else
          If ((tileID > 99) And (tileID <= highestDatTile)) Then
            If DatTiles(tileID).blocking = False Then

                If ColourPriority(ColourGround) > ColourPriority(gridMap(cellPos).BackColor) Then
                  gridMap(cellPos).BackColor = ColourGround
                End If
                If ColourPriority(ColourUp) > ColourPriority(gridMap(cellPos).BackColor) Then
                  If DatTiles(tileID).floorChangeUP Then
                    gridMap(cellPos).BackColor = ColourUp
                  End If
                End If
               If ColourPriority(ColourDown) > ColourPriority(gridMap(cellPos).BackColor) Then
                 If DatTiles(tileID).floorChangeDOWN Then
                   gridMap(cellPos).BackColor = ColourDown
                 End If
               End If
               If ColourPriority(ColourField) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).isField Then
                gridMap(cellPos).BackColor = ColourField
              End If
            End If
            Else

                If DatTiles(tileID).notMoveable = True Then
                  ' blocking and not moveable
                  If ColourPriority(ColourSomething) > ColourPriority(gridMap(cellPos).BackColor) Then
                    gridMap(cellPos).BackColor = ColourSomething
                  End If
                  
            If ColourPriority(ColourUp) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).floorChangeUP Then
                gridMap(cellPos).BackColor = ColourUp
              End If
            End If
            If ColourPriority(ColourDown) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).floorChangeDOWN Then
                gridMap(cellPos).BackColor = ColourDown
              End If
            End If
                
                Else ' blocking but moveable
                  If ColourPriority(ColourBlockMoveable) > ColourPriority(gridMap(cellPos).BackColor) Then
                    gridMap(cellPos).BackColor = ColourBlockMoveable
                  End If
                
                         If ColourPriority(ColourUp) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).floorChangeUP Then
                gridMap(cellPos).BackColor = ColourUp
              End If
            End If
            If ColourPriority(ColourDown) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).floorChangeDOWN Then
                gridMap(cellPos).BackColor = ColourDown
              End If
            End If
                
                
                End If
            End If
          ElseIf tileID = 0 Then
             ' end of stack
             Exit For
          End If
        End If
      Next poss
            
            
            
            
            
            
            
            
            
           
            cellPos = cellPos + 1
        Next PosX
    Next PosY
    Me.AutoRedraw = True
    Me.txtSelected.Text = "Done"
End Sub


Private Function SafeByte(cl As Long) As Byte
    If cl > 255 Then
    SafeByte = &HFF
    ElseIf cl < 0 Then
    SafeByte = &H0
    Else
    SafeByte = CByte(cl)
    End If
End Function

Public Sub DrawFloor3(Optional selectedFloor = True)
    '....

    Dim drawThisFloor As Long
    Dim tibiaclient As Long
    Dim cellPos As Long
    Dim Z As Long
    Dim PosX As Long
    Dim PosY As Long
    Dim MT As MapTile3
    Dim tmpID As Double
    Dim poss As Long
    Dim tileID As Long
    Dim tmpName As String
    Dim withID As Boolean
    Dim theExtra As String
    Dim MTc As Long
    Dim tmpam As Byte
    withID = True
    tibiaclient = 0
    tibiaclient = GetFirstTibiaPID()
    If tibiaclient = 0 Then
        Me.txtSelected.Text = BString(65)
        Exit Sub
    End If
    Z = CLng(Memory_ReadByte(PLAYER_Z, tibiaclient))
    If selectedFloor = True Then
        drawThisFloor = mapFloorSelected
    Else
        drawThisFloor = Z
    End If
    UpdateMap tibiaclient
    
    
    
    'debugtiles
    'Exit Sub
    
    
    Me.AutoRedraw = False
    
    cellPos = 0
    For PosY = -6 To 7
        For PosX = -8 To 9
            MT = GetMapTile3(PosX, PosY, drawThisFloor)
'            If (PosX = -1) And (PosY = 0) Then
'                'Debug.Print "bien"
'            End If
            
            
            
            
            
            
            
            
            
      gridMap(cellPos).BackColor = ColourNothing
      GridCaption(cellPos) = ""
      CharCount(cellPos) = 0
      gridMap(cellPos).Caption = ""
      MTc = MT.count - 1
      'If MTc > 0 Then
       ' MsgBox "debug"
      'End If
      For poss = 0 To MTc
        If poss < UBound(MT.items) Then
        
        
        tileID = MT.items(poss).id
        If MT.items(poss).data1 < 255 Then
        tmpam = SafeByte(MT.items(poss).data1)
        Else
        tmpam = 0
        End If
        
        #If FinalMode = 0 Then
         GridCaption(cellPos) = GridCaption(cellPos) & _
         "[" & GoodHex(LowByteOfLong(tileID)) & " " & GoodHex(HighByteOfLong(tileID)) & " " & GoodHex(tmpam) & "]"
        #End If
        If ((tileID <= 0) Or (tileID > highestDatTile)) Then
           Exit For
        ElseIf tileID = &H63 Then
          gridMap(cellPos).BackColor = ColourPlayer
          If GridCaption(cellPos) = "" Then
          tmpID = CDbl(MT.items(poss).data1)
          If tmpID = 0 Then
            GridCaption(cellPos) = "tileid " & CStr(tileID) & "??"
            CharCount(cellPos) = CharCount(cellPos) + 1
            gridMap(cellPos).Caption = CStr(CharCount(cellPos))
            
          Else
            theExtra = ""
            If withID = True Then
              theExtra = " (" & CStr(tmpID) & ")"
            End If
            GridCaption(cellPos) = GetNameFromID(tibiaclient, tmpID) & theExtra
            CharCount(cellPos) = CharCount(cellPos) + 1
            gridMap(cellPos).Caption = CStr(CharCount(cellPos))
          End If
          Else
            tmpID = CDbl(MT.items(poss).data1)
            If tmpID = 0 Then
              tmpName = "tileid " & CStr(tileID) & "??"
            Else
               tmpName = GetNameFromID(tibiaclient, tmpID)
               
            End If
            theExtra = ""
            If withID = True Then
              theExtra = " (" & CStr(tmpID) & ")"
            End If
            GridCaption(cellPos) = GridCaption(cellPos) & " , " & tmpName & theExtra
            CharCount(cellPos) = CharCount(cellPos) + 1
            gridMap(cellPos).Caption = CStr(CharCount(cellPos))
          End If
           'gotMobiles = True
        ElseIf poss = 0 Then
          If tileID <> &H0 Then
            gridMap(cellPos).BackColor = ColourGround

            If ColourPriority(ColourUp) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).floorChangeUP Then
                gridMap(cellPos).BackColor = ColourUp
              End If
            End If
            If ColourPriority(ColourDown) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).floorChangeDOWN Then
                gridMap(cellPos).BackColor = ColourDown
              End If
            End If
            
            If ColourPriority(ColourWater) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).isWater = True Then
                gridMap(cellPos).BackColor = ColourWater
                If DatTiles(tileID).haveFish = True And _
                 ColourPriority(ColourFish) > ColourPriority(gridMap(cellPos).BackColor) Then
                  gridMap(cellPos).BackColor = ColourFish
                End If
              End If
            End If
          End If
          If DatTiles(tileID).isWater = False Then
            If ColourPriority(ColourSomething2) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).blocking = True And DatTiles(tileID).notMoveable = False Then
                gridMap(cellPos).BackColor = ColourSomething2
              End If
            End If
            If ColourPriority(ColourSomething) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).blocking = True And DatTiles(tileID).notMoveable = True Then
                gridMap(cellPos).BackColor = ColourSomething
              End If
            End If
          End If
            If ColourPriority(ColourField) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).isField Then
               gridMap(cellPos).BackColor = ColourField
              End If
            End If
        Else
          If ((tileID > 99) And (tileID <= highestDatTile)) Then
            If DatTiles(tileID).blocking = False Then

                If ColourPriority(ColourGround) > ColourPriority(gridMap(cellPos).BackColor) Then
                  gridMap(cellPos).BackColor = ColourGround
                End If
                If ColourPriority(ColourUp) > ColourPriority(gridMap(cellPos).BackColor) Then
                  If DatTiles(tileID).floorChangeUP Then
                    gridMap(cellPos).BackColor = ColourUp
                  End If
                End If
               If ColourPriority(ColourDown) > ColourPriority(gridMap(cellPos).BackColor) Then
                 If DatTiles(tileID).floorChangeDOWN Then
                   gridMap(cellPos).BackColor = ColourDown
                 End If
               End If
               If ColourPriority(ColourField) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).isField Then
                gridMap(cellPos).BackColor = ColourField
              End If
            End If
            Else

                If DatTiles(tileID).notMoveable = True Then
                  ' blocking and not moveable
                  If ColourPriority(ColourSomething) > ColourPriority(gridMap(cellPos).BackColor) Then
                    gridMap(cellPos).BackColor = ColourSomething
                  End If
                  
            If ColourPriority(ColourUp) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).floorChangeUP Then
                gridMap(cellPos).BackColor = ColourUp
              End If
            End If
            If ColourPriority(ColourDown) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).floorChangeDOWN Then
                gridMap(cellPos).BackColor = ColourDown
              End If
            End If
                
                Else ' blocking but moveable
                  If ColourPriority(ColourBlockMoveable) > ColourPriority(gridMap(cellPos).BackColor) Then
                    gridMap(cellPos).BackColor = ColourBlockMoveable
                  End If
                
                         If ColourPriority(ColourUp) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).floorChangeUP Then
                gridMap(cellPos).BackColor = ColourUp
              End If
            End If
            If ColourPriority(ColourDown) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).floorChangeDOWN Then
                gridMap(cellPos).BackColor = ColourDown
              End If
            End If
                
                
                End If
            End If
          ElseIf tileID = 0 Then
             ' end of stack
             Exit For
          End If
        End If
        Else
            Debug.Print "error. bad map read"
            Exit For
        End If
      Next poss
            
            
            
            
            
            
            
            
            
           
            cellPos = cellPos + 1
        Next PosX
    Next PosY
    Me.AutoRedraw = True
    Me.txtSelected.Text = "Done"
End Sub


Public Sub DrawFloor5(Optional selectedFloor = True)
    Dim drawThisFloor As Long
    Dim tibiaclient As Long
    Dim cellPos As Long
    Dim Z As Long
    Dim PosX As Long
    Dim PosY As Long
    Dim MT As MapTile5
    Dim tmpID As Double
    Dim poss As Long
    Dim tileID As Long
    Dim tmpName As String
    Dim withID As Boolean
    Dim theExtra As String
    Dim MTc As Long
    Dim tmpam As Byte
    withID = True
    tibiaclient = 0
    tibiaclient = GetFirstTibiaPID()
    If tibiaclient = 0 Then
        Me.txtSelected.Text = BString(65)
        Exit Sub
    End If
    Z = CLng(Memory_ReadByte(PLAYER_Z, tibiaclient))
    If selectedFloor = True Then
        drawThisFloor = mapFloorSelected
    Else
        drawThisFloor = Z
    End If
    UpdateMap tibiaclient
    Me.AutoRedraw = False
    cellPos = 0
    For PosY = -6 To 7
        For PosX = -8 To 9
            MT = GetMapTile5(PosX, PosY, drawThisFloor)
            gridMap(cellPos).BackColor = ColourNothing
            GridCaption(cellPos) = ""
            CharCount(cellPos) = 0
            gridMap(cellPos).Caption = ""
            MTc = MT.count - 1
            For poss = 0 To MTc
                If poss <= UBound(MT.items) Then
                    tileID = MT.items(poss).id
                    If MT.items(poss).data1 < 255 Then
                        tmpam = SafeByte(MT.items(poss).data1)
                    Else
                        tmpam = 0
                    End If
                    #If FinalMode = 0 Then
                        GridCaption(cellPos) = GridCaption(cellPos) & _
                        "[" & GoodHex(LowByteOfLong(tileID)) & " " & GoodHex(HighByteOfLong(tileID)) & " " & GoodHex(tmpam) & "]"
                    #End If
                    If ((tileID <= 0) Or (tileID > highestDatTile)) Then
                        Exit For
                    ElseIf tileID = &H63 Then
                        gridMap(cellPos).BackColor = ColourPlayer
                        If GridCaption(cellPos) = "" Then
                            tmpID = CDbl(MT.items(poss).data1)
                            If tmpID = 0 Then
                                GridCaption(cellPos) = "tileid " & CStr(tileID) & "??"
                                CharCount(cellPos) = CharCount(cellPos) + 1
                                gridMap(cellPos).Caption = CStr(CharCount(cellPos))
                            Else
                                theExtra = ""
                                If withID = True Then
                                    theExtra = " (" & CStr(tmpID) & ")"
                                End If
                                GridCaption(cellPos) = GetNameFromID(tibiaclient, tmpID) & theExtra
                                CharCount(cellPos) = CharCount(cellPos) + 1
                                gridMap(cellPos).Caption = CStr(CharCount(cellPos))
                            End If
                        Else
                            tmpID = CDbl(MT.items(poss).data1)
                            If tmpID = 0 Then
                                tmpName = "tileid " & CStr(tileID) & "??"
                            Else
                                tmpName = GetNameFromID(tibiaclient, tmpID)
                            End If
                            theExtra = ""
                            If withID = True Then
                                theExtra = " (" & CStr(tmpID) & ")"
                            End If
                            GridCaption(cellPos) = GridCaption(cellPos) & " , " & tmpName & theExtra
                            CharCount(cellPos) = CharCount(cellPos) + 1
                            gridMap(cellPos).Caption = CStr(CharCount(cellPos))
                        End If
                        'gotMobiles = True
                    ElseIf poss = 0 Then
                        If tileID <> &H0 Then
                            gridMap(cellPos).BackColor = ColourGround
                            If ColourPriority(ColourUp) > ColourPriority(gridMap(cellPos).BackColor) Then
                                If DatTiles(tileID).floorChangeUP Then
                                    gridMap(cellPos).BackColor = ColourUp
                                End If
                            End If
                            If ColourPriority(ColourDown) > ColourPriority(gridMap(cellPos).BackColor) Then
                                If DatTiles(tileID).floorChangeDOWN Then
                                    gridMap(cellPos).BackColor = ColourDown
                                End If
                            End If
                            If ColourPriority(ColourWater) > ColourPriority(gridMap(cellPos).BackColor) Then
                                If DatTiles(tileID).isWater = True Then
                                    gridMap(cellPos).BackColor = ColourWater
                                    If DatTiles(tileID).haveFish = True And ColourPriority(ColourFish) > ColourPriority(gridMap(cellPos).BackColor) Then
                                        gridMap(cellPos).BackColor = ColourFish
                                    End If
                                End If
                            End If
                        End If
                        If DatTiles(tileID).isWater = False Then
                            If ColourPriority(ColourSomething2) > ColourPriority(gridMap(cellPos).BackColor) Then
                                If DatTiles(tileID).blocking = True And DatTiles(tileID).notMoveable = False Then
                                    gridMap(cellPos).BackColor = ColourSomething2
                                End If
                            End If
                            If ColourPriority(ColourSomething) > ColourPriority(gridMap(cellPos).BackColor) Then
                                If DatTiles(tileID).blocking = True And DatTiles(tileID).notMoveable = True Then
                                    gridMap(cellPos).BackColor = ColourSomething
                                End If
                            End If
                        End If
                        If ColourPriority(ColourField) > ColourPriority(gridMap(cellPos).BackColor) Then
                            If DatTiles(tileID).isField Then
                                gridMap(cellPos).BackColor = ColourField
                            End If
                        End If
                    Else
                        If ((tileID > 99) And (tileID <= highestDatTile)) Then
                            If DatTiles(tileID).blocking = False Then
                                If ColourPriority(ColourGround) > ColourPriority(gridMap(cellPos).BackColor) Then
                                    gridMap(cellPos).BackColor = ColourGround
                                End If
                                If ColourPriority(ColourUp) > ColourPriority(gridMap(cellPos).BackColor) Then
                                    If DatTiles(tileID).floorChangeUP Then
                                        gridMap(cellPos).BackColor = ColourUp
                                    End If
                                End If
                                If ColourPriority(ColourDown) > ColourPriority(gridMap(cellPos).BackColor) Then
                                    If DatTiles(tileID).floorChangeDOWN Then
                                        gridMap(cellPos).BackColor = ColourDown
                                    End If
                                End If
                                If ColourPriority(ColourField) > ColourPriority(gridMap(cellPos).BackColor) Then
                                    If DatTiles(tileID).isField Then
                                       gridMap(cellPos).BackColor = ColourField
                                    End If
                                End If
                            Else
                                If DatTiles(tileID).notMoveable = True Then
                                    ' blocking and not moveable
                                    If ColourPriority(ColourSomething) > ColourPriority(gridMap(cellPos).BackColor) Then
                                        gridMap(cellPos).BackColor = ColourSomething
                                    End If
                                    If ColourPriority(ColourUp) > ColourPriority(gridMap(cellPos).BackColor) Then
                                        If DatTiles(tileID).floorChangeUP Then
                                            gridMap(cellPos).BackColor = ColourUp
                                        End If
                                    End If
                                    If ColourPriority(ColourDown) > ColourPriority(gridMap(cellPos).BackColor) Then
                                        If DatTiles(tileID).floorChangeDOWN Then
                                            gridMap(cellPos).BackColor = ColourDown
                                        End If
                                    End If
                                Else ' blocking but moveable
                                    If ColourPriority(ColourBlockMoveable) > ColourPriority(gridMap(cellPos).BackColor) Then
                                        gridMap(cellPos).BackColor = ColourBlockMoveable
                                    End If
                                    If ColourPriority(ColourUp) > ColourPriority(gridMap(cellPos).BackColor) Then
                                        If DatTiles(tileID).floorChangeUP Then
                                            gridMap(cellPos).BackColor = ColourUp
                                        End If
                                    End If
                                    If ColourPriority(ColourDown) > ColourPriority(gridMap(cellPos).BackColor) Then
                                        If DatTiles(tileID).floorChangeDOWN Then
                                            gridMap(cellPos).BackColor = ColourDown
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf tileID = 0 Then
                        ' end of stack
                            Exit For
                        End If
                    End If
                Else ' (poss<>0) AND (poss > UBound(MT.items))
                    Debug.Print "error. bad map read: poss=" & CStr(poss) & " ubound = " & CStr(UBound(MT.items)) & " MTC = " & CStr(MTc)
                    Exit For
                End If
            Next poss
            cellPos = cellPos + 1
        Next PosX
    Next PosY
    Me.AutoRedraw = True
    Me.txtSelected.Text = "Done"
End Sub


Public Sub DrawFloor4(Optional selectedFloor = True)
    '....

    Dim drawThisFloor As Long
    Dim tibiaclient As Long
    Dim cellPos As Long
    Dim Z As Long
    Dim PosX As Long
    Dim PosY As Long
    Dim MT As MapTile4
    Dim tmpID As Double
    Dim poss As Long
    Dim tileID As Long
    Dim tmpName As String
    Dim withID As Boolean
    Dim theExtra As String
    Dim MTc As Long
    Dim tmpam As Byte
    withID = True
    tibiaclient = 0
    tibiaclient = GetFirstTibiaPID()
    If tibiaclient = 0 Then
        Me.txtSelected.Text = BString(65)
        Exit Sub
    End If
    Z = CLng(Memory_ReadByte(PLAYER_Z, tibiaclient))
    If selectedFloor = True Then
        drawThisFloor = mapFloorSelected
    Else
        drawThisFloor = Z
    End If
    UpdateMap tibiaclient
    
    
    
    'debugtiles
    'Exit Sub
    
    
    Me.AutoRedraw = False
    
    cellPos = 0
    For PosY = -6 To 7
        For PosX = -8 To 9
            MT = GetMapTile4(PosX, PosY, drawThisFloor)
'            If (PosX = -1) And (PosY = 0) Then
'                'Debug.Print "bien"
'            End If
            
            
            
            
            
            
            
            
            
      gridMap(cellPos).BackColor = ColourNothing
      GridCaption(cellPos) = ""
      CharCount(cellPos) = 0
      gridMap(cellPos).Caption = ""
      MTc = MT.count - 1
      'If MTc > 0 Then
       ' MsgBox "debug"
      'End If
      For poss = 0 To MTc
        If poss < UBound(MT.items) Then
        
        
        tileID = MT.items(poss).id
        If MT.items(poss).data1 < 255 Then
        tmpam = SafeByte(MT.items(poss).data1)
        Else
        tmpam = 0
        End If
        
        #If FinalMode = 0 Then
         GridCaption(cellPos) = GridCaption(cellPos) & _
         "[" & GoodHex(LowByteOfLong(tileID)) & " " & GoodHex(HighByteOfLong(tileID)) & " " & GoodHex(tmpam) & "]"
        #End If
        If ((tileID <= 0) Or (tileID > highestDatTile)) Then
           Exit For
        ElseIf tileID = &H63 Then
          gridMap(cellPos).BackColor = ColourPlayer
          If GridCaption(cellPos) = "" Then
          tmpID = CDbl(MT.items(poss).data1)
          If tmpID = 0 Then
            GridCaption(cellPos) = "tileid " & CStr(tileID) & "??"
            CharCount(cellPos) = CharCount(cellPos) + 1
            gridMap(cellPos).Caption = CStr(CharCount(cellPos))
            
          Else
            theExtra = ""
            If withID = True Then
              theExtra = " (" & CStr(tmpID) & ")"
            End If
            GridCaption(cellPos) = GetNameFromID(tibiaclient, tmpID) & theExtra
            CharCount(cellPos) = CharCount(cellPos) + 1
            gridMap(cellPos).Caption = CStr(CharCount(cellPos))
          End If
          Else
            tmpID = CDbl(MT.items(poss).data1)
            If tmpID = 0 Then
              tmpName = "tileid " & CStr(tileID) & "??"
            Else
               tmpName = GetNameFromID(tibiaclient, tmpID)
               
            End If
            theExtra = ""
            If withID = True Then
              theExtra = " (" & CStr(tmpID) & ")"
            End If
            GridCaption(cellPos) = GridCaption(cellPos) & " , " & tmpName & theExtra
            CharCount(cellPos) = CharCount(cellPos) + 1
            gridMap(cellPos).Caption = CStr(CharCount(cellPos))
          End If
           'gotMobiles = True
        ElseIf poss = 0 Then
          If tileID <> &H0 Then
            gridMap(cellPos).BackColor = ColourGround

            If ColourPriority(ColourUp) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).floorChangeUP Then
                gridMap(cellPos).BackColor = ColourUp
              End If
            End If
            If ColourPriority(ColourDown) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).floorChangeDOWN Then
                gridMap(cellPos).BackColor = ColourDown
              End If
            End If
            
            If ColourPriority(ColourWater) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).isWater = True Then
                gridMap(cellPos).BackColor = ColourWater
                If DatTiles(tileID).haveFish = True And _
                 ColourPriority(ColourFish) > ColourPriority(gridMap(cellPos).BackColor) Then
                  gridMap(cellPos).BackColor = ColourFish
                End If
              End If
            End If
          End If
          If DatTiles(tileID).isWater = False Then
            If ColourPriority(ColourSomething2) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).blocking = True And DatTiles(tileID).notMoveable = False Then
                gridMap(cellPos).BackColor = ColourSomething2
              End If
            End If
            If ColourPriority(ColourSomething) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).blocking = True And DatTiles(tileID).notMoveable = True Then
                gridMap(cellPos).BackColor = ColourSomething
              End If
            End If
          End If
            If ColourPriority(ColourField) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).isField Then
               gridMap(cellPos).BackColor = ColourField
              End If
            End If
        Else
          If ((tileID > 99) And (tileID <= highestDatTile)) Then
            If DatTiles(tileID).blocking = False Then

                If ColourPriority(ColourGround) > ColourPriority(gridMap(cellPos).BackColor) Then
                  gridMap(cellPos).BackColor = ColourGround
                End If
                If ColourPriority(ColourUp) > ColourPriority(gridMap(cellPos).BackColor) Then
                  If DatTiles(tileID).floorChangeUP Then
                    gridMap(cellPos).BackColor = ColourUp
                  End If
                End If
               If ColourPriority(ColourDown) > ColourPriority(gridMap(cellPos).BackColor) Then
                 If DatTiles(tileID).floorChangeDOWN Then
                   gridMap(cellPos).BackColor = ColourDown
                 End If
               End If
               If ColourPriority(ColourField) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).isField Then
                gridMap(cellPos).BackColor = ColourField
              End If
            End If
            Else

                If DatTiles(tileID).notMoveable = True Then
                  ' blocking and not moveable
                  If ColourPriority(ColourSomething) > ColourPriority(gridMap(cellPos).BackColor) Then
                    gridMap(cellPos).BackColor = ColourSomething
                  End If
                  
            If ColourPriority(ColourUp) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).floorChangeUP Then
                gridMap(cellPos).BackColor = ColourUp
              End If
            End If
            If ColourPriority(ColourDown) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).floorChangeDOWN Then
                gridMap(cellPos).BackColor = ColourDown
              End If
            End If
                
                Else ' blocking but moveable
                  If ColourPriority(ColourBlockMoveable) > ColourPriority(gridMap(cellPos).BackColor) Then
                    gridMap(cellPos).BackColor = ColourBlockMoveable
                  End If
                
                         If ColourPriority(ColourUp) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).floorChangeUP Then
                gridMap(cellPos).BackColor = ColourUp
              End If
            End If
            If ColourPriority(ColourDown) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).floorChangeDOWN Then
                gridMap(cellPos).BackColor = ColourDown
              End If
            End If
                
                
                End If
            End If
          ElseIf tileID = 0 Then
             ' end of stack
             Exit For
          End If
        End If
        Else
            Debug.Print "error. bad map read"
            Exit For
        End If
      Next poss
            
            
            
            
            
            
            
            
            
           
            cellPos = cellPos + 1
        Next PosX
    Next PosY
    Me.AutoRedraw = True
    Me.txtSelected.Text = "Done"
End Sub


Public Sub SetButtonColours(Optional init As Boolean = False, Optional tc As Long = 0)
 Dim Z As Long
 Dim i As Long
 Dim tibiaclient As Long
 tibiaclient = 0
 If init = False Then
    tibiaclient = GetFirstTibiaPID()
 Else
    tibiaclient = tc
 End If
 If tibiaclient = 0 Then
   Z = 7
 Else
   Z = CLng(Memory_ReadByte(PLAYER_Z, tibiaclient))
 End If
 If Z <= 7 Then
   For i = 0 To 7
     If i = Z Then
       cmdFloor(i).BackColor = ColourWithMe
     Else
       cmdFloor(i).BackColor = ColourWithInfo
     End If
   Next i
   For i = 8 To 15
   cmdFloor(i).BackColor = ColourUnknown
   Next i
 Else
   For i = 0 To 15
     cmdFloor(i).BackColor = ColourUnknown
   Next i
   For i = Z - 2 To Z + 2
     If i = Z Then
       cmdFloor(i).BackColor = ColourWithMe
     ElseIf i <= 15 Then
       cmdFloor(i).BackColor = ColourWithInfo
     End If
   Next i
 End If
 cmdFloor(mapFloorSelected).BackColor = ColourSelected
End Sub

Private Sub cmdFloor_Click(Index As Integer)
 Dim tibiaclient As Long
 tibiaclient = 0
 
 tibiaclient = GetFirstTibiaPID()
 SetButtonColours True, tibiaclient
 If tibiaclient = 0 Then
    Me.txtSelected.Text = BString(65)
 Else
    mapFloorSelected = CLng(Index)
    If TibiaVersionLong >= 1021 Then
        DrawFloor5 True
    ElseIf TibiaVersionLong >= 1021 Then
        DrawFloor4 True
    ElseIf TibiaVersionLong >= 990 Then
        DrawFloor3 True
    ElseIf TibiaVersionLong >= 942 Then
        DrawFloor2 True
    ElseIf TibiaVersionLong > 772 Then
        DrawFloor True
    Else
        DrawFloorOld True
    End If
 End If
End Sub

Private Sub cmdUpdate_Click()
    If TibiaVersionLong >= 1050 Then
        DrawFloor5 True
    ElseIf TibiaVersionLong >= 1021 Then
        DrawFloor4 True
    ElseIf TibiaVersionLong >= 990 Then
        DrawFloor3 True
    ElseIf TibiaVersionLong >= 942 Then
        DrawFloor2 True
    ElseIf TibiaVersionLong > 772 Then
        DrawFloor True
    Else
        DrawFloorOld True
    End If
End Sub


Private Sub cmdUpdateMyFloor_Click()
    If TibiaVersionLong >= 1050 Then
        DrawFloor5 False
    ElseIf TibiaVersionLong >= 1021 Then
        DrawFloor4 False
    ElseIf TibiaVersionLong >= 990 Then
        DrawFloor3 False
    ElseIf TibiaVersionLong >= 942 Then
        DrawFloor2 False
    ElseIf TibiaVersionLong > 772 Then
        DrawFloor True
    Else
        DrawFloorOld True
    End If
End Sub

Private Sub Form_Load()
    gridMap_col = 0
    gridMap_row = 0
    mapFloorSelected = 7
    SetButtonColours True, 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MustUnload = False Then
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub gridMap_Click(Index As Integer)
    gridMap_col = CLng(Index Mod 18)
    gridMap_row = CLng(Index \ 18)
    
    txtSelected.Text = "col=" & gridMap_col & "row=" & gridMap_row & " : " & GridCaption(Index)
End Sub


















Public Sub DrawFloorOld(Optional selectedFloor = True)
    ' Tibia 7.72 or lower
    Dim drawThisFloor As Long
    Dim tibiaclient As Long
    Dim cellPos As Long
    Dim Z As Long
    Dim PosX As Long
    Dim PosY As Long
    Dim MT As MapTileOld
    Dim tmpID As Double
    Dim poss As Long
    Dim tileID As Long
    Dim tmpName As String
    Dim withID As Boolean
    Dim theExtra As String
    Dim MTc As Long
    Dim tmpam As Byte
    Dim tmpamSTR As String
    withID = True
    tibiaclient = 0
    tibiaclient = GetFirstTibiaPID()
    If tibiaclient = 0 Then
        Me.txtSelected.Text = BString(65)
        Exit Sub
    End If
    Z = CLng(Memory_ReadByte(PLAYER_Z, tibiaclient))
    If selectedFloor = True Then
        drawThisFloor = mapFloorSelected
    Else
        drawThisFloor = Z
    End If
    UpdateMap tibiaclient
    Me.AutoRedraw = False
    
    cellPos = 0
    For PosY = -6 To 7
        For PosX = -8 To 9
            MT = GetMapTileOld(PosX, PosY, drawThisFloor)

            
            
            
            
            
            
            
            
      gridMap(cellPos).BackColor = ColourNothing
      GridCaption(cellPos) = ""
      CharCount(cellPos) = 0
      gridMap(cellPos).Caption = ""
      MTc = MT.count - 1
      If MTc > 9 Then
        MTc = 9
      End If
      For poss = 0 To MTc

        tileID = MT.items(poss).id
        If (MT.items(poss).data1 >= 0) And (MT.items(poss).data1 < 255) Then
        tmpam = CByte(MT.items(poss).data1)
        tmpamSTR = " " & GoodHex(tmpam)
        Else
        tmpam = 0
        tmpamSTR = ""
        End If
        
        #If FinalMode = 0 Then
         GridCaption(cellPos) = GridCaption(cellPos) & _
         "[" & GoodHex(LowByteOfLong(tileID)) & " " & GoodHex(HighByteOfLong(tileID)) & tmpamSTR & "]"
        #End If
        If ((tileID <= 0) Or (tileID > highestDatTile)) Then
           Exit For
        ElseIf tileID = &H63 Then
          gridMap(cellPos).BackColor = ColourPlayer
          If GridCaption(cellPos) = "" Then
          tmpID = CDbl(MT.items(poss).data1)
          If tmpID = 0 Then
            GridCaption(cellPos) = "tileid " & CStr(tileID) & "??"
            CharCount(cellPos) = CharCount(cellPos) + 1
            gridMap(cellPos).Caption = CStr(CharCount(cellPos))
            
          Else
            theExtra = ""
            If withID = True Then
              theExtra = " (" & CStr(tmpID) & ")"
            End If
            GridCaption(cellPos) = GetNameFromID(tibiaclient, tmpID) & theExtra
            CharCount(cellPos) = CharCount(cellPos) + 1
            gridMap(cellPos).Caption = CStr(CharCount(cellPos))
          End If
          Else
            tmpID = CDbl(MT.items(poss).data1)
            If tmpID = 0 Then
              tmpName = "tileid " & CStr(tileID) & "??"
            Else
               tmpName = GetNameFromID(tibiaclient, tmpID)
               
            End If
            theExtra = ""
            If withID = True Then
              theExtra = " (" & CStr(tmpID) & ")"
            End If
            GridCaption(cellPos) = GridCaption(cellPos) & " , " & tmpName & theExtra
            CharCount(cellPos) = CharCount(cellPos) + 1
            gridMap(cellPos).Caption = CStr(CharCount(cellPos))
          End If
           'gotMobiles = True
        ElseIf poss = 0 Then
          If (tileID <> &H0) Then
            gridMap(cellPos).BackColor = ColourGround

            If ColourPriority(ColourUp) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).floorChangeUP Then
                gridMap(cellPos).BackColor = ColourUp
              End If
            End If
            If ColourPriority(ColourDown) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).floorChangeDOWN Then
                gridMap(cellPos).BackColor = ColourDown
              End If
            End If
            
            If ColourPriority(ColourWater) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).isWater = True Then
                gridMap(cellPos).BackColor = ColourWater
                If DatTiles(tileID).haveFish = True And _
                 ColourPriority(ColourFish) > ColourPriority(gridMap(cellPos).BackColor) Then
                  gridMap(cellPos).BackColor = ColourFish
                End If
              End If
            End If
          End If
          If DatTiles(tileID).isWater = False Then
            If ColourPriority(ColourSomething2) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).blocking = True And DatTiles(tileID).notMoveable = False Then
                gridMap(cellPos).BackColor = ColourSomething2
              End If
            End If
            If ColourPriority(ColourSomething) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).blocking = True And DatTiles(tileID).notMoveable = True Then
                gridMap(cellPos).BackColor = ColourSomething
              End If
            End If
          End If
            If ColourPriority(ColourField) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).isField Then
               gridMap(cellPos).BackColor = ColourField
              End If
            End If
        Else
          If ((tileID > 99) And (tileID <= highestDatTile)) Then
            If DatTiles(tileID).blocking = False Then

                If ColourPriority(ColourGround) > ColourPriority(gridMap(cellPos).BackColor) Then
                  gridMap(cellPos).BackColor = ColourGround
                End If
                If ColourPriority(ColourUp) > ColourPriority(gridMap(cellPos).BackColor) Then
                  If DatTiles(tileID).floorChangeUP Then
                    gridMap(cellPos).BackColor = ColourUp
                  End If
                End If
               If ColourPriority(ColourDown) > ColourPriority(gridMap(cellPos).BackColor) Then
                 If DatTiles(tileID).floorChangeDOWN Then
                   gridMap(cellPos).BackColor = ColourDown
                 End If
               End If
               If ColourPriority(ColourField) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).isField Then
                gridMap(cellPos).BackColor = ColourField
              End If
            End If
            Else

                If DatTiles(tileID).notMoveable = True Then
                  ' blocking and not moveable
                  If ColourPriority(ColourSomething) > ColourPriority(gridMap(cellPos).BackColor) Then
                    gridMap(cellPos).BackColor = ColourSomething
                  End If
                  
            If ColourPriority(ColourUp) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).floorChangeUP Then
                gridMap(cellPos).BackColor = ColourUp
              End If
            End If
            If ColourPriority(ColourDown) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).floorChangeDOWN Then
                gridMap(cellPos).BackColor = ColourDown
              End If
            End If
                
                Else ' blocking but moveable
                  If ColourPriority(ColourBlockMoveable) > ColourPriority(gridMap(cellPos).BackColor) Then
                    gridMap(cellPos).BackColor = ColourBlockMoveable
                  End If
                
                         If ColourPriority(ColourUp) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).floorChangeUP Then
                gridMap(cellPos).BackColor = ColourUp
              End If
            End If
            If ColourPriority(ColourDown) > ColourPriority(gridMap(cellPos).BackColor) Then
              If DatTiles(tileID).floorChangeDOWN Then
                gridMap(cellPos).BackColor = ColourDown
              End If
            End If
                
                
                End If
            End If
          ElseIf tileID = 0 Then
             ' end of stack
             Exit For
          End If
        End If
      Next poss
            
            
            
            
            
            
            
            
            
           
            cellPos = cellPos + 1
        Next PosX
    Next PosY
    Me.AutoRedraw = True
    Me.txtSelected.Text = "Done"
End Sub


