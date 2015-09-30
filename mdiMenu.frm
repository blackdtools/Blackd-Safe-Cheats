VERSION 5.00
Begin VB.MDIForm mdiMenu 
   BackColor       =   &H00000000&
   Caption         =   "Blackd Safe Cheats 2.2.1"
   ClientHeight    =   5400
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   6615
   Icon            =   "mdiMenu.frx":0000
   LinkTopic       =   "Main menu"
   Picture         =   "mdiMenu.frx":058A
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrDebug 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   4080
      Top             =   1680
   End
   Begin VB.Menu mPopupSys 
      Caption         =   "&SysTray"
      Visible         =   0   'False
      Begin VB.Menu mPopupShow 
         Caption         =   "Show BSC"
      End
      Begin VB.Menu mNothing3 
         Caption         =   "-"
      End
      Begin VB.Menu mShowAllTibia 
         Caption         =   "Show Tibia"
      End
      Begin VB.Menu mHideAllTibia 
         Caption         =   "Hide Tibia"
      End
      Begin VB.Menu mNothing4 
         Caption         =   "-"
      End
      Begin VB.Menu mDoNothing 
         Caption         =   "Do Nothing"
      End
   End
   Begin VB.Menu mCheats 
      Caption         =   "Open cheats"
      Begin VB.Menu mOpenHPmana 
         Caption         =   "HP and mana"
      End
      Begin VB.Menu mOpenLight 
         Caption         =   "Light"
      End
      Begin VB.Menu mXRAY 
         Caption         =   "XRAY"
      End
      Begin VB.Menu mTruemap 
         Caption         =   "Truemap"
      End
      Begin VB.Menu mAutoeater 
         Caption         =   "Autoeater"
      End
   End
   Begin VB.Menu mSettings 
      Caption         =   "Settings"
      Begin VB.Menu mLoadSettings 
         Caption         =   "Load settings from file..."
      End
      Begin VB.Menu mSaveSettings 
         Caption         =   "Save settings to file..."
      End
      Begin VB.Menu mSettingsHPmana 
         Caption         =   "HP and mana"
         Begin VB.Menu mHPdelay1 
            Caption         =   "Define delay between low status and recharge"
         End
         Begin VB.Menu mHPdelay2 
            Caption         =   "Define delay between recharges"
         End
         Begin VB.Menu mHPrand 
            Caption         =   "Define percent of delay time randomized"
         End
      End
      Begin VB.Menu mSettingsLight 
         Caption         =   "Light"
         Begin VB.Menu mDefineLightDelay 
            Caption         =   "Define time between each light update"
         End
      End
      Begin VB.Menu mTibiaPath 
         Caption         =   "Set Tibia Path..."
         Visible         =   0   'False
      End
      Begin VB.Menu mNothing 
         Caption         =   "-"
      End
      Begin VB.Menu mReloadDefault 
         Caption         =   "Reload settings from default.ini"
      End
      Begin VB.Menu mNothing2 
         Caption         =   "-"
      End
      Begin VB.Menu mSetLanguageFile 
         Caption         =   "Set language file..."
      End
   End
   Begin VB.Menu mLinks 
      Caption         =   "Links"
      Begin VB.Menu mBuyGold 
         Caption         =   "Buy Tibia gold"
      End
      Begin VB.Menu mShootFruits 
         Caption         =   "Play Shoot Fruits (flash game)"
      End
      Begin VB.Menu mForum 
         Caption         =   "Blackdtools.com forum"
      End
   End
   Begin VB.Menu mAbout 
      Caption         =   "About"
      Begin VB.Menu mLatestchanges 
         Caption         =   "Latest changes"
      End
      Begin VB.Menu mCopyright 
         Caption         =   "Copyright"
      End
   End
   Begin VB.Menu mDebug 
      Caption         =   "Debug"
      Visible         =   0   'False
      Begin VB.Menu mTestFocus 
         Caption         =   "Test Focus"
      End
   End
   Begin VB.Menu mUntested 
      Caption         =   "Untested Cheats!"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "mdiMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const FinalMode = 1
Option Explicit
Private Const SW_NORMAL = 1
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Function FixPoint(ByVal strThing As String) As String
    strThing = Replace(strThing, ".", ",")
    FixPoint = strThing
End Function
Private Function SpecialFixPoint(ByVal strThing As String) As Double
    Dim strTemp As String
    Dim dblTemp As Double
    Dim dblRes As String
    dblTemp = CDbl(FixPoint("1.00"))
    strTemp = FixPoint(strThing)
    dblRes = (100 * CDbl(strTemp)) / dblTemp
    SpecialFixPoint = dblRes
End Function

Private Function GetTibiaVersionLong(ByVal TibiaVersion As String) As Long
    Dim thePoint As Long
    Dim partLeft As String
    Dim partRight As String
    Dim result As String
    Dim lngResult As Long
    thePoint = InStr(1, TibiaVersion, ".", vbTextCompare)
    If thePoint <= 0 Then
        MsgBox "Error at GetTibiaVersionLong(" & TibiaVersion & ")", vbOKOnly + vbCritical, "Critical Error"
        End
    End If
    partLeft = left$(TibiaVersion, thePoint - 1)
    partRight = Right$(TibiaVersion, Len(TibiaVersion) - thePoint)
    If Len(partRight) = 2 Then
        result = partLeft & partRight
    Else
        result = partLeft & partRight & "0"
    End If
    lngResult = CLng(result)
    GetTibiaVersionLong = lngResult
End Function

Private Function LoadConfig(strPath As String) As String
  #If FinalMode Then
    On Error GoTo goterr
  #End If
    Dim strInfo As String
    Dim i As Long
    Dim lonInfo As Long
    Dim strThing As String
    Dim here As String
    strInfo = String$(10, 0)
    strThing = "TibiaVersion"
    i = getfromINI("Tibia", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        TibiaVersion = strInfo
       
    Else
        LoadConfig = "Could not read the value of " & strThing
        Exit Function
    End If
    TibiaVersionLong = GetTibiaVersionLong(TibiaVersion)
'    TibiaVersionLong = CLng(Round(CDbl(FixPoint(TibiaVersion)) * 100))
'    Select Case TibiaVersion
'    Case "8.55"
'        TibiaVersionLong = 855
'    Case "8.56"
'        TibiaVersionLong = 856
'    Case "8.57"
'        TibiaVersionLong = 857
'    Case "8.6"
'        TibiaVersionLong = 860
'    Case "8.61"
'        TibiaVersionLong = 861
'    Case "8.62"
'        TibiaVersionLong = 862
'    Case "8.7"
'        TibiaVersionLong = 870
'    Case "8.71"
'        TibiaVersionLong = 871
'    Case "8.72"
'        TibiaVersionLong = 872
'    Case "8.73"
'        TibiaVersionLong = 873
'    Case "8.74"
'        TibiaVersionLong = 874
'    Case "9.00"
'        TibiaVersionLong = 900
'    Case "9"
'        TibiaVersionLong = 900
'    Case "9.0"
'        TibiaVersionLong = 900
'    Case "9.1"
'        TibiaVersionLong = 910
'    Case "9.2"
'        TibiaVersionLong = 920
'    Case "9.31"
'        TibiaVersionLong = 931
'    Case "9.4"
'        TibiaVersionLong = 940
'    Case "9.41"
'        TibiaVersionLong = 941
'    Case "9.42"
'        TibiaVersionLong = 942
'    Case "9.43"
'        TibiaVersionLong = 943
'    Case "9.44"
'        TibiaVersionLong = 944
'    Case "9.45"
'        TibiaVersionLong = 945
'    Case "9.46"
'        TibiaVersionLong = 946
'    Case "9.5"
'        TibiaVersionLong = 950
'    Case "9.50"
'        TibiaVersionLong = 950
'    Case "9.51"
'        TibiaVersionLong = 951
'    Case "9.52"
'        TibiaVersionLong = 952
'    Case "9.53"
'        TibiaVersionLong = 953
'    Case "8,55"
'        TibiaVersionLong = 855
'    Case "8,56"
'        TibiaVersionLong = 856
'    Case "8,57"
'        TibiaVersionLong = 857
'    Case "8,6"
'        TibiaVersionLong = 860
'    Case "8,61"
'        TibiaVersionLong = 861
'    Case "8,62"
'        TibiaVersionLong = 862
'    Case "8,7"
'        TibiaVersionLong = 870
'    Case "8,71"
'        TibiaVersionLong = 871
'    Case "8,72"
'        TibiaVersionLong = 872
'    Case "8,73"
'        TibiaVersionLong = 873
'    Case "8,74"
'        TibiaVersionLong = 874
'    Case "9"
'        TibiaVersionLong = 900
'    Case "9,0"
'        TibiaVersionLong = 900
'    Case "9,00"
'        TibiaVersionLong = 900
'    Case "9,1"
'        TibiaVersionLong = 910
'    Case "9,2"
'        TibiaVersionLong = 920
'    Case "9,31"
'        TibiaVersionLong = 931
'    Case "9,4"
'        TibiaVersionLong = 940
'    Case "9,41"
'        TibiaVersionLong = 941
'    Case "9,42"
'        TibiaVersionLong = 942
'    Case "9,43"
'        TibiaVersionLong = 943
'    Case "9,44"
'        TibiaVersionLong = 944
'    Case "9,45"
'        TibiaVersionLong = 945
'    Case "9,46"
'        TibiaVersionLong = 946
'    Case "9,5"
'        TibiaVersionLong = 950
'    Case "9,50"
'        TibiaVersionLong = 950
'    Case "9,51"
'        TibiaVersionLong = 951
'    Case "9,52"
'        TibiaVersionLong = 952
'    Case "9,53"
'        TibiaVersionLong = 953
'    Case Else
'        TibiaVersionLong = SpecialFixPoint(TibiaVersion)
'    End Select
    
    strInfo = String$(50, 0)
    strThing = "tibiaclassname"
    i = getfromINI("Tibia", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        tibiaclassname = strInfo
    Else
        LoadConfig = "Could not read the value of " & strThing
        Exit Function
    End If
    If LCase(tibiaclassname) = "tibiaclient" Then
     mdiMenu.Caption = "Blackd Safe Cheats " & SafeVersion & " for Tibia " & TibiaVersion
    Else
     mdiMenu.Caption = "Blackd Safe Cheats " & SafeVersion & " for Tibia " & TibiaVersion & " PREVIEW"
    End If
    
    
    
    strInfo = String$(50, 0)
    strThing = "DefaultTibiaFolder"
    i = getfromINI("Tibia", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        DefaultTibiaFolder = strInfo
    Else
        DefaultTibiaFolder = "Tibia"
    End If
    
    strInfo = String$(50, 0)
    strThing = "OverwriteTibiaExePath"
    i = getfromINI("Tibia", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        OverwriteTibiaExePath = strInfo
    Else
        OverwriteTibiaExePath = ""
    End If
    
    strInfo = String$(10, 0)
    strThing = "LEVELSPY_NOP"
    i = getfromINI("Tibia", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        LEVELSPY_NOP = lonInfo
    Else
        LoadConfig = "Could not read the value of " & strThing
        Exit Function
    End If
    
    strInfo = String$(10, 0)
    strThing = "LEVELSPY_ABOVE"
    i = getfromINI("Tibia", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        LEVELSPY_ABOVE = lonInfo
    Else
        LoadConfig = "Could not read the value of " & strThing
        Exit Function
    End If
    
    strInfo = String$(10, 0)
    strThing = "LEVELSPY_BELOW"
    i = getfromINI("Tibia", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        LEVELSPY_BELOW = lonInfo
    Else
        LoadConfig = "Could not read the value of " & strThing
        Exit Function
    End If
    
    strInfo = String$(10, 0)
    strThing = "LIGHT_NOP"
    i = getfromINI("Tibia", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        LIGHT_NOP = lonInfo
    Else
        LIGHT_NOP = &H0
    End If
    
    
    strInfo = String$(10, 0)
    strThing = "LIGHT_AMOUNT"
    i = getfromINI("Tibia", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        LIGHT_AMOUNT = lonInfo
    Else
        LIGHT_AMOUNT = 0
    End If
    
    strInfo = String$(255, 0)
    strThing = "LIGHT_TRICK_CODE"
    i = getfromINI("Tibia", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        LIGHT_TRICK_CODE = strInfo
    Else
        LIGHT_TRICK_CODE = ""
    End If
    ' Changes from ...
'00526662      8B1D E0BF8100    MOV EBX,DWORD PTR DS:[81BFE0]
'00526668      85DB             TEST EBX,EBX
'0052666A      79 04            JNS SHORT 00526670

    ' to... (in tibia 10.5)
'00526662      BB FF000000      MOV EBX,0FF
'00526667      EB 11            JMP SHORT 0052667A
'00526669      90               NOP
'0052666A      90               NOP
'0052666A      90               NOP


    
    strInfo = String$(10, 0)
    strThing = "LIGHT_TRICK_ADR"
    i = getfromINI("Tibia", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        LIGHT_TRICK_ADR = lonInfo
    Else
        LIGHT_TRICK_ADR = 0
    End If
    
    
    
    

    
    strInfo = String$(10, 0)
    strThing = "adrNChar"
    i = getfromINI("Tibia", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        adrNChar = lonInfo
    Else
        LoadConfig = "Could not read the value of " & strThing
        Exit Function
    End If
    
    strInfo = String$(10, 0)
    strThing = "MAP_POINTER_ADDR"
    i = getfromINI("Tibia", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        MAP_POINTER_ADDR = lonInfo
    Else
        LoadConfig = "Could not read the value of " & strThing
        Exit Function
    End If
    
    
    strInfo = String$(10, 0)
    strThing = "OFFSET_POINTER_ADDR"
    i = getfromINI("Tibia", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        OFFSET_POINTER_ADDR = lonInfo
    Else
        OFFSET_POINTER_ADDR = MAP_POINTER_ADDR + &H1C
    End If

    
    strInfo = String$(10, 0)
    strThing = "CharDist"
    i = getfromINI("Tibia", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        CharDist = lonInfo
    Else
        LoadConfig = "Could not read the value of " & strThing
        Exit Function
    End If
    
    strInfo = String$(10, 0)
    strThing = "NameDist"
    i = getfromINI("Tibia", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        NameDist = lonInfo
    Else
        LoadConfig = "Could not read the value of " & strThing
        Exit Function
    End If
    
    strInfo = String$(10, 0)
    strThing = "OutfitDist"
    i = getfromINI("Tibia", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        OutfitDist = lonInfo
    Else
        LoadConfig = "Could not read the value of " & strThing
        Exit Function
    End If

    strInfo = String$(10, 0)
    strThing = "adrNum"
    i = getfromINI("Tibia", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        adrNum = lonInfo
    Else
        LoadConfig = "Could not read the value of " & strThing
        Exit Function
    End If

    strInfo = String$(10, 0)
    strThing = "LightDist"
    i = getfromINI("Tibia", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        LightDist = lonInfo
    Else
        LoadConfig = "Could not read the value of " & strThing
        Exit Function
    End If
    
    strInfo = String$(10, 0)
    strThing = "LightColourDist"
    i = getfromINI("Tibia", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        LightColourDist = lonInfo
    Else
        LoadConfig = "Could not read the value of " & strThing
        Exit Function
    End If

    
    
    
    

    strInfo = String$(10, 0)
    strThing = "adrMyHP"
    i = getfromINI("Tibia", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        adrMyHP = lonInfo
    Else
        LoadConfig = "Could not read the value of " & strThing
        Exit Function
    End If
    
    strInfo = String$(10, 0)
    strThing = "adrMyMaxHP"
    i = getfromINI("Tibia", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        adrMyMaxHP = lonInfo
    Else
        LoadConfig = "Could not read the value of " & strThing
        Exit Function
    End If
    
    strInfo = String$(10, 0)
    strThing = "adrMyMana"
    i = getfromINI("Tibia", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        adrMyMana = lonInfo
    Else
        LoadConfig = "Could not read the value of " & strThing
        Exit Function
    End If
    
    strInfo = String$(10, 0)
    strThing = "adrMyMaxMana"
    i = getfromINI("Tibia", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        adrMyMaxMana = lonInfo
    Else
        LoadConfig = "Could not read the value of " & strThing
        Exit Function
    End If
    
        strInfo = String$(10, 0)
    strThing = "adrMySoul"
    i = getfromINI("Tibia", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        adrMySoul = lonInfo
    Else
        LoadConfig = "Could not read the value of " & strThing
        Exit Function
    End If
    
    strInfo = String$(10, 0)
    strThing = "adrXOR"
    i = getfromINI("Tibia", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        adrXOR = lonInfo
    Else
        adrXOR = &H944008
    End If
    

    strInfo = String$(10, 0)
    strThing = "adrConnected"
    i = getfromINI("Tibia", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        adrConnected = lonInfo
    Else
        LoadConfig = "Could not read the value of " & strThing
        Exit Function
    End If

    
    strInfo = String$(10, 0)
    strThing = "PLAYER_Z"
    i = getfromINI("Tibia", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        PLAYER_Z = lonInfo
    Else
        LoadConfig = "Could not read the value of " & strThing
        Exit Function
    End If





  
    strInfo = String$(10, 0)
    strThing = "MAXTILEIDLISTSIZE"
    i = getfromINI("Tibia", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        MAXTILEIDLISTSIZE = lonInfo
    Else
        MAXTILEIDLISTSIZE = 50
    End If
    
    strInfo = String$(10, 0)
    strThing = "MAXDATTILES"
    i = getfromINI("Tibia", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        MAXDATTILES = lonInfo
    Else
        MAXDATTILES = 11000
    End If
  
  ReDim AditionalStairsToDownFloor(0 To MAXTILEIDLISTSIZE)
  ReDim AditionalStairsToUpFloor(0 To MAXTILEIDLISTSIZE)
  ReDim AditionalRequireRope(0 To MAXTILEIDLISTSIZE)
  ReDim AditionalRequireShovel(0 To MAXTILEIDLISTSIZE)
  ReDim DatTiles(0 To MAXDATTILES)
    
  ' Read some tile ID values from the ini :
  
  here = strPath
  ' runes
  ReadTileIDListFromIni AditionalStairsToUpFloor, "AditionalStairsToUpFloor", here, "AC 07,AE 07,AA 07,94 08,96 08,90 08,92 08"
  ReadTileIDListFromIni AditionalStairsToDownFloor, "AditionalStairsToDownFloor", here, ""
  ReadTileIDListFromIni AditionalRequireRope, "AditionalRequireRope", here, ""
  ReadTileIDListFromIni AditionalRequireShovel, "AditionalRequireShovel", here, ""
  
  ReadTileIDFromIni tileID_Blank, "tileID_Blank", here, "0D 0C"

  ReadTileIDFromIni tileID_WallBugItem, "tileID_WallBugItem", here, "4E 10"
  
  blank1 = LowByteOfLong(tileID_Blank)
  blank2 = HighByteOfLong(tileID_Blank)
  
  ReadTileIDFromIni tileID_SD, "tileID_SD", here, "53 0C"
  ReadTileIDFromIni tileID_HMM, "tileID_HMM", here, "40 0C"
  ReadTileIDFromIni tileID_Explosion, "tileID_Explosion", here, "42 0C"
  ReadTileIDFromIni tileID_IH, "tileID_IH", here, "12 0C"
  ReadTileIDFromIni tileID_UH, "tileID_UH", here, "1A 0C"
  
  ReadTileIDFromIni tileID_fireball, "tileID_fireball", here, "75 0C"
  ReadTileIDFromIni tileID_stalagmite, "tileID_stalagmite", here, "6B 0C"
  ReadTileIDFromIni tileID_icicle, "tileID_icicle", here, "56 0C"
  
  ' items
  ReadTileIDFromIni tileID_Bag, "tileID_Bag", here, "E7 0A"
  ReadTileIDFromIni tileID_Backpack, "tileID_Backpack", here, "E8 0A"
  ReadTileIDFromIni tileID_Oracle, "tileID_Oracle", here, "DA 07"
  ReadTileIDFromIni tileID_FishingRod, "tileID_FishingRod", here, "5D 0D"
 
  ReadTileIDFromIni tileID_Rope, "tileID_Rope", here, "7D 0B"
  ReadTileIDFromIni tileID_LightRope, "tileID_LightRope", here, "86 02"
  ReadTileIDFromIni tileID_Shovel, "tileID_Shovel", here, "43 0D"
  ReadTileIDFromIni tileID_LightShovel, "tileID_LightShovel", here, "4E 16"

  ' water
  ReadTileIDFromIni tileID_waterEmpty, "tileID_waterEmpty", here, "5B 02"
  ReadTileIDFromIni tileID_waterWithFish, "tileID_waterWithFish", here, "59 02"
  
  ReadTileIDFromIni tileID_waterEmptyEnd, "tileID_waterEmptyEnd", here, "5B 02"
  ReadTileIDFromIni tileID_waterWithFishEnd, "tileID_waterWithFishEnd", here, "59 02"
  
  ' blocking table
  ReadTileIDFromIni tileID_blockingBox, "tileID_blockingBox", here, "A5 09"
  
  ' to UP floor
  ReadTileIDFromIni tileID_stairsToUp, "tileID_stairsToUp", here, "88 07"
  ReadTileIDFromIni tileID_woodenStairstoUp, "tileID_woodenStairstoUp", here, "93 07"
  
  ReadTileIDFromIni tileID_desertRamptoUp, "tileID_desertRamptoUp", here, "A8 07"
  
  ReadTileIDFromIni tileID_rampToNorth, "tileID_rampToNorth", here, "91 07"
  ReadTileIDFromIni tileID_rampToSouth, "tileID_rampToSouth", here, "8F 07"
 
  ReadTileIDFromIni tileID_rampToRightCycMountain, "tileID_rampToRightCycMountain", here, "8B 07"
  ReadTileIDFromIni tileID_rampToLeftCycMountain, "tileID_rampToLeftCycMountain", here, "8D 07"
  
  
  ReadTileIDFromIni tileID_jungleStairsToNorth, "tileID_jungleStairsToNorth", here, "B9 07"
  ReadTileIDFromIni tileID_jungleStairsToLeft, "tileID_jungleStairsToLeft", here, "BA 07"
  
  ' + requires rightClick
  ReadTileIDFromIni tileID_ladderToUp, "tileID_ladderToUp", here, "89 07"
  
  ' + requires rope
  ReadTileIDFromIni tileID_holeInCelling, "tileID_holeInCelling", here, "80 01"
  
  ' to DOWN
  ReadTileIDFromIni tileID_grassCouldBeHole, "tileID_grassCouldBeHole", here, "25 01"
  ReadTileIDFromIni tileID_pitfall, "tileID_pitfall", here, "26 01"

  ReadTileIDFromIni tileID_openHole, "tileID_openHole", here, "44 02"
  ReadTileIDFromIni tileID_OpenDesertLooseStonePile, "tileID_OpenDesertLooseStonePile", here, "51 02"
  
  
  ReadTileIDFromIni tileID_trapdoor, "tileID_trapdoor", here, "71 01"
  ReadTileIDFromIni tileID_down1, "tileID_down1", here, "72 01"
  
  ReadTileIDFromIni tileID_openHole2, "tileID_openHole2", here, "7F 01"
  
  ReadTileIDFromIni tileID_trapdoor2, "tileID_trapdoor2", here, "98 01"
  ReadTileIDFromIni tileID_down2, "tileID_down2", here, "99 01"
  ReadTileIDFromIni tileID_stairsToDownKazordoon, "tileID_stairsToDownKazordoon", here, "9A 01"
  ReadTileIDFromIni tileID_stairsToDownThais, "tileID_stairsToDownThais", here, "9B 01"
  
  ReadTileIDFromIni tileID_trapdoorKazordoon, "tileID_trapdoorKazordoon", here, "AB 01"
  ReadTileIDFromIni tileID_down3, "tileID_down3", here, "AC 01"
  ReadTileIDFromIni tileID_stairsToDown, "tileID_stairsToDown", here, "AD 01"
  
  ReadTileIDFromIni tileID_stairsToDown2, "tileID_stairsToDown2", here, "B0 01"
  ReadTileIDFromIni tileID_woodenStairstoDown, "tileID_woodenStairstoDown", here, "B1 01"
  
  ReadTileIDFromIni tileID_rampToDown, "tileID_rampToDown", here, "CB 01"

  ' + requires rightClick
  ReadTileIDFromIni tileID_sewerGate, "tileID_sewerGate", here, "AE 01"

  ' + requires shovel
  ReadTileIDFromIni tileID_closedHole, "tileID_closedHole", here, "43 02"
  ReadTileIDFromIni tileID_desertLooseStonePile, "tileID_desertLooseStonePile", here, "50 02"
  
  ' FOOD
  ReadTileIDFromIni tileID_firstFoodTileID, "tileID_firstFoodTileID", here, "BB 0D"
  ReadTileIDFromIni tileID_lastFoodTileID, "tileID_lastFoodTileID", here, "D9 0D"
  ReadTileIDFromIni tileID_firstMushroomTileID, "tileID_firstMushroomTileID", here, "4A 0E"
  ReadTileIDFromIni tileID_lastMushroomTileID, "tileID_lastMushroomTileID", here, "4E 0E"
  
  'FIELD RANGE1
  ReadTileIDFromIni tileID_firstFieldRangeStart, "tileID_firstFieldRangeStart", here, "31 08"
  ReadTileIDFromIni tileID_firstFieldRangeEnd, "tileID_firstFieldRangeEnd", here, "3A 08"
  ReadTileIDFromIni tileID_secondFieldRangeStart, "tileID_secondFieldRangeStart", here, "3E 08"
  ReadTileIDFromIni tileID_secondFieldRangeEnd, "tileID_secondFieldRangeEnd", here, "45 08"

  ReadTileIDFromIni tileID_campFire1, "tileID_campFire1", here, "20 20"
  ReadTileIDFromIni tileID_campFire2, "tileID_campFire2", here, "20 20"

  'WALKABLE FIELDS
  ReadTileIDFromIni tileID_walkableFire1, "tileID_walkableFire1", here, "33 08"
  ReadTileIDFromIni tileID_walkableFire2, "tileID_walkableFire2", here, "38 08"
  ReadTileIDFromIni tileID_walkableFire3, "tileID_walkableFire3", here, "40 08"
  
  ' Depot chest
  ReadTileIDFromIni tileID_depotChest, "tileID_depotChest", here, "70 0D"
  
  ' flasks - mana fluids
  ReadTileIDFromIni tileID_flask, "tileID_flask", here, "3A 0B"
  
  
  ReadTileIDFromIni tileID_health_potion, "tileID_health_potion", here, "0A 01"
  ReadTileIDFromIni tileID_strong_health_potion, "tileID_strong_health_potion", here, "EC 00"
  ReadTileIDFromIni tileID_great_health_potion, "tileID_great_health_potion", here, "EF 00"
  ReadTileIDFromIni tileID_small_health_potion, "tileID_small_health_potion", here, "C4 1E"
  ReadTileIDFromIni tileID_mana_potion, "tileID_mana_potion", here, "0C 01"
  ReadTileIDFromIni tileID_strong_mana_potion, "tileID_strong_mana_potion", here, "ED 00"
  ReadTileIDFromIni tileID_great_mana_potion, "tileID_great_mana_potion", here, "EE 00"
  
  ReadTileIDFromIni tileID_ultimate_health_potion, "tileID_ultimate_health_potion", here, "DB 1D"
  ReadTileIDFromIni tileID_great_spirit_potion, "tileID_great_spirit_potion", here, "DA 1D"
  
  
    strInfo = String$(10, 0)
    strThing = "byteNothing"
    i = getfromINI("tileIDs", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        byteNothing = lonInfo
    Else
        byteNothing = &H0
    End If
    
    strInfo = String$(10, 0)
    strThing = "byteMana"
    i = getfromINI("tileIDs", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        byteMana = lonInfo
    Else
        byteMana = &H7
    End If
    
    strInfo = String$(10, 0)
    strThing = "byteLife"
    i = getfromINI("tileIDs", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        byteLife = lonInfo
    Else
        byteLife = &HB
    End If
    
    
    strInfo = String$(10, 0)
    strThing = "LAST_BATTLELISTP2OS"
    i = getfromINI("Tibia", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        LAST_BATTLELISTPOS = lonInfo
    Else
        LAST_BATTLELISTPOS = 147
    End If

    strInfo = String$(10, 0)
    strThing = "tibiaModuleRegionSize"
    i = getfromINI("Tibia", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        tibiaModuleRegionSize = lonInfo
    Else
        tibiaModuleRegionSize = &H2C3000
    End If
    
    strInfo = String$(10, 0)
    strThing = "useDynamicOffset"
    i = getfromINI("Tibia", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        useDynamicOffset = strInfo
    Else
        useDynamicOffset = "no"
    End If
    If useDynamicOffset = "yes" Then
      useDynamicOffsetBool = True
    Else
      useDynamicOffsetBool = False
    End If

    LoadConfig = ""
    Exit Function
goterr:
    LoadConfig = "LoadConfig: Got error code " & Err.Number & ": " & Err.Description
End Function
Private Function LoadSettings(strPath As String) As String
  #If FinalMode = 1 Then
  On Error GoTo goterr
  #End If
    Dim strInfo As String
    Dim i As Long
    Dim lonInfo As Long
    Dim strThing As String
    Dim here As String
    here = strPath
    
    strInfo = String$(255, 0)
    strThing = "LanguageFile"
    i = getfromINI("Misc", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        LanguageFile = strInfo
    Else
        LoadSettings = "Could not read the value of " & strThing
        Exit Function
    End If
    
    
    strInfo = String$(10, 0)
    strThing = "HPmanadelay1"
    i = getfromINI("HPmana", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        HPmanadelay1 = lonInfo
    Else
        LoadSettings = "Could not read the value of " & strThing
        Exit Function
    End If
    
    strInfo = String$(10, 0)
    strThing = "HPmanadelay2"
    i = getfromINI("HPmana", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        HPmanadelay2 = lonInfo
    Else
        LoadSettings = "Could not read the value of " & strThing
        Exit Function
    End If
    
    strInfo = String$(10, 0)
    strThing = "HPrandpercent"
    i = getfromINI("HPmana", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        HPrandpercent = lonInfo
    Else
        LoadSettings = "Could not read the value of " & strThing
        Exit Function
    End If
    
    strInfo = String$(10, 0)
    strThing = "HPmanaLimit0"
    i = getfromINI("HPmana", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        HPmanaLimit0 = lonInfo
    Else
        LoadSettings = "Could not read the value of " & strThing
        Exit Function
    End If
    
    strInfo = String$(10, 0)
    strThing = "HPmanaLimit1"
    i = getfromINI("HPmana", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        HPmanaLimit1 = lonInfo
    Else
        LoadSettings = "Could not read the value of " & strThing
        Exit Function
    End If
    
    strInfo = String$(10, 0)
    strThing = "HPmanaLimit2"
    i = getfromINI("HPmana", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        HPmanaLimit2 = lonInfo
    Else
        LoadSettings = "Could not read the value of " & strThing
        Exit Function
    End If
    
    strInfo = String$(10, 0)
    strThing = "HPmanaLimit3"
    i = getfromINI("HPmana", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        HPmanaLimit3 = lonInfo
    Else
        LoadSettings = "Could not read the value of " & strThing
        Exit Function
    End If

    strInfo = String$(10, 0)
    strThing = "HPmanaLimit4"
    i = getfromINI("HPmana", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        HPmanaLimit4 = lonInfo
    Else
        HPmanaLimit4 = 0
        'LoadSettings = "Could not read the value of " & strThing
        'Exit Function
    End If
    
    strInfo = String$(10, 0)
    strThing = "HPmanaLimit5"
    i = getfromINI("HPmana", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        HPmanaLimit5 = lonInfo
    Else
        HPmanaLimit5 = 0
        'LoadSettings = "Could not read the value of " & strThing
        'Exit Function
    End If

    strInfo = String$(50, 0)
    strThing = "HPmanaAction0"
    i = getfromINI("HPmana", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        HPmanaAction0 = strInfo
    Else
        LoadSettings = "Could not read the value of " & strThing
        Exit Function
    End If
    
    strInfo = String$(50, 0)
    strThing = "HPmanaAction1"
    i = getfromINI("HPmana", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        HPmanaAction1 = strInfo
    Else
        LoadSettings = "Could not read the value of " & strThing
        Exit Function
    End If
    
    strInfo = String$(50, 0)
    strThing = "HPmanaAction2"
    i = getfromINI("HPmana", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        HPmanaAction2 = strInfo
    Else
        LoadSettings = "Could not read the value of " & strThing
        Exit Function
    End If
    
    strInfo = String$(50, 0)
    strThing = "HPmanaAction3"
    i = getfromINI("HPmana", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        HPmanaAction3 = strInfo
    Else
        LoadSettings = "Could not read the value of " & strThing
        Exit Function
    End If


    strInfo = String$(50, 0)
    strThing = "HPmanaAction4"
    i = getfromINI("HPmana", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        HPmanaAction4 = strInfo
    Else
        HPmanaAction4 = "--"
        'LoadSettings = "Could not read the value of " & strThing
        'Exit Function
    End If
    
    strInfo = String$(50, 0)
    strThing = "HPmanaAction5"
    i = getfromINI("HPmana", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        HPmanaAction5 = strInfo
    Else
        HPmanaAction5 = "--"
        'LoadSettings = "Could not read the value of " & strThing
        'Exit Function
    End If

    strInfo = String$(10, 0)
    strThing = "LightIntensity"
    i = getfromINI("Light", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        LightIntensity = lonInfo
    Else
        LoadSettings = "Could not read the value of " & strThing
        Exit Function
    End If
    
    strInfo = String$(10, 0)
    strThing = "LightColour"
    i = getfromINI("Light", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        LightColour = lonInfo
    Else
        LoadSettings = "Could not read the value of " & strThing
        Exit Function
    End If

    strInfo = String$(10, 0)
    strThing = "LightEnabled"
    i = getfromINI("Light", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        If lonInfo = 1 Then
            LightEnabled = True
        Else
            LightEnabled = False
        End If
    Else
        LightEnabled = False
        LoadSettings = "Could not read the value of " & strThing
        Exit Function
    End If


    strInfo = String$(10, 0)
    strThing = "LightRefreshDelay"
    i = getfromINI("Light", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        LightRefreshDelay = lonInfo
    Else
        LoadSettings = "Could not read the value of " & strThing
        Exit Function
    End If
    
    
    strInfo = String$(10, 0)
    strThing = "XRAY_floors_ABOVE"
    i = getfromINI("XRAY", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        XRAY_floors_ABOVE = lonInfo
    Else
        XRAY_floors_ABOVE = -1
    End If
    
    strInfo = String$(10, 0)
    strThing = "XRAY_floors_BELOW"
    i = getfromINI("XRAY", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        XRAY_floors_BELOW = lonInfo
    Else
        XRAY_floors_BELOW = 1
    End If
    
    strInfo = String$(10, 0)
    strThing = "XRAY_key1_1"
    i = getfromINI("XRAY", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        XRAY_key1_1 = lonInfo
    Else
        XRAY_key1_1 = 0
    End If
    
    strInfo = String$(10, 0)
    strThing = "XRAY_key1_2"
    i = getfromINI("XRAY", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        XRAY_key1_2 = lonInfo
    Else
        XRAY_key1_2 = 0
    End If

    strInfo = String$(10, 0)
    strThing = "XRAY_key2_1"
    i = getfromINI("XRAY", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        XRAY_key2_1 = lonInfo
    Else
        XRAY_key2_1 = 0
    End If
    
    strInfo = String$(10, 0)
    strThing = "XRAY_key2_2"
    i = getfromINI("XRAY", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        XRAY_key2_2 = lonInfo
    Else
        XRAY_key2_2 = 0
    End If
    
    strInfo = String$(10, 0)
    strThing = "XRAY_key3_1"
    i = getfromINI("XRAY", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        XRAY_key3_1 = lonInfo
    Else
        XRAY_key3_1 = 0
    End If
    
    strInfo = String$(10, 0)
    strThing = "XRAY_key2_2"
    i = getfromINI("XRAY", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        XRAY_key3_2 = lonInfo
    Else
        XRAY_key3_2 = 0
    End If
    
    
    
    
    strInfo = String$(10, 0)
    strThing = "AutoeaterKey"
    i = getfromINI("Autoeater", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        AutoeaterKey = lonInfo
    Else
        AutoeaterKey = 13
    End If
    
    strInfo = String$(20, 0)
    strThing = "AutoeaterTimerFrom"
    i = getfromINI("Autoeater", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        AutoeaterTimerFrom = lonInfo
    Else
        AutoeaterTimerFrom = 40000
    End If
    
    strInfo = String$(20, 0)
    strThing = "AutoeaterTimerTo"
    i = getfromINI("Autoeater", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        lonInfo = CLng(strInfo)
        AutoeaterTimerTo = lonInfo
    Else
        AutoeaterTimerTo = 60000
    End If
    
    
    strInfo = String$(255, 0)
    strThing = "TibiaExePath"
    i = getfromINI("Misc", strThing, "", strInfo, Len(strInfo), strPath)
    If i > 0 Then
        strInfo = left(strInfo, i)
        TibiaExePath = strInfo
    Else
        TibiaExePath = ""

    End If
    If Not (OverwriteTibiaExePath = "") Then
        TibiaExePath = OverwriteTibiaExePath
    ElseIf TibiaExePath = "" Then
        TibiaExePath = autoGetTibiaFolder()
    End If
    If (Not (TibiaExePath = "")) Then
        If (Not (Right$(TibiaExePath, 1) = "\")) Then
            TibiaExePath = TibiaExePath & "\"
        End If
    End If
    
    
    TryReadTiles
    
    Hotkeys(0).key1 = XRAY_key1_1
    Hotkeys(0).key2 = XRAY_key1_2
    Hotkeys(1).key1 = XRAY_key2_1
    Hotkeys(1).key2 = XRAY_key2_2
    Hotkeys(2).key1 = XRAY_key3_1
    Hotkeys(2).key2 = XRAY_key3_2
    
    
    
    
    
    

    'Debug.Print DatTiles(1294).blocking
    
    LoadSettings = ""
    Exit Function
goterr:
    LoadSettings = "LoadSettings: Got error code " & Err.Number & ": " & Err.Description
End Function


Private Function LoadLanguage(strPath As String) As String
    On Error GoTo goterr
    Dim strInfo As String
    Dim i As Long
    Dim lonInfo As Long
    Dim B As Long
    Dim strThing As String
    LoadDefaultStrings
    For B = 0 To LastBstring
        strInfo = String$(255, 0)
        strThing = "S" & CStr(B)
        i = getfromINI("Lang", strThing, "", strInfo, Len(strInfo), strPath)
        If i > 0 Then
            strInfo = left(strInfo, i)
            BString(B) = strInfo
        End If
    Next B
    BString(17) = FixVBCRLF(BString(17))
    BString(59) = FixVBCRLF(BString(59))
    BString(66) = FixVBCRLF(BString(66))
    BString(77) = FixVBCRLF(BString(77))

    LoadLanguage = ""
    Exit Function
goterr:
    LoadLanguage = "LoadLanguage: Got error code " & Err.Number & ": " & Err.Description
End Function
Private Function SaveSettings(ByVal strPath As String) As String
    On Error GoTo goterr
    Dim strInfo As String
    Dim i As Long
    strInfo = CStr(HPmanadelay1)
    i = setToINI("HPmana", "HPmanadelay1", strInfo, strPath)
    strInfo = CStr(HPmanadelay2)
    i = setToINI("HPmana", "HPmanadelay2", strInfo, strPath)
    strInfo = CStr(HPrandpercent)
    i = setToINI("HPmana", "HPrandpercent", strInfo, strPath)
    
    strInfo = CStr(frmMenuAutohealer.cmbLimit(0).Text)
    i = setToINI("HPmana", "HPmanaLimit0", strInfo, strPath)
    strInfo = CStr(frmMenuAutohealer.cmbLimit(1).Text)
    i = setToINI("HPmana", "HPmanaLimit1", strInfo, strPath)
    strInfo = CStr(frmMenuAutohealer.cmbLimit(2).Text)
    i = setToINI("HPmana", "HPmanaLimit2", strInfo, strPath)
    strInfo = CStr(frmMenuAutohealer.cmbLimit(3).Text)
    i = setToINI("HPmana", "HPmanaLimit3", strInfo, strPath)
    strInfo = CStr(frmMenuAutohealer.cmbLimit(4).Text)
    i = setToINI("HPmana", "HPmanaLimit4", strInfo, strPath)
    strInfo = CStr(frmMenuAutohealer.cmbLimit(5).Text)
    i = setToINI("HPmana", "HPmanaLimit5", strInfo, strPath)
    
    strInfo = CStr(frmMenuAutohealer.cmbHotkey(0).Text)
    i = setToINI("HPmana", "HPmanaAction0", strInfo, strPath)
    strInfo = CStr(frmMenuAutohealer.cmbHotkey(1).Text)
    i = setToINI("HPmana", "HPmanaAction1", strInfo, strPath)
    strInfo = CStr(frmMenuAutohealer.cmbHotkey(2).Text)
    i = setToINI("HPmana", "HPmanaAction2", strInfo, strPath)
    strInfo = CStr(frmMenuAutohealer.cmbHotkey(3).Text)
    i = setToINI("HPmana", "HPmanaAction3", strInfo, strPath)
    strInfo = CStr(frmMenuAutohealer.cmbHotkey(4).Text)
    i = setToINI("HPmana", "HPmanaAction4", strInfo, strPath)
    strInfo = CStr(frmMenuAutohealer.cmbHotkey(5).Text)
    i = setToINI("HPmana", "HPmanaAction5", strInfo, strPath)
    
    strInfo = CStr(LanguageFile)
    i = setToINI("Misc", "LanguageFile", strInfo, strPath)
    
    strInfo = CStr(LightIntensity)
    i = setToINI("Light", "LightIntensity", strInfo, strPath)
    strInfo = CStr(LightColour)
    i = setToINI("Light", "LightColour", strInfo, strPath)
    
    If LightEnabled = True Then
        strInfo = "1"
    Else
        strInfo = "0"
    End If
    i = setToINI("Light", "LightEnabled", strInfo, strPath)
    
    strInfo = CStr(LightRefreshDelay)
    i = setToINI("Light", "LightRefreshDelay", strInfo, strPath)
    
    
    
    
    strInfo = CStr(XRAY_floors_ABOVE)
    i = setToINI("XRAY", "XRAY_floors_ABOVE", strInfo, strPath)
    
    strInfo = CStr(XRAY_floors_BELOW)
    i = setToINI("XRAY", "XRAY_floors_BELOW", strInfo, strPath)
    
    strInfo = CStr(XRAY_key1_1)
    i = setToINI("XRAY", "XRAY_key1_1", strInfo, strPath)
    
    strInfo = CStr(XRAY_key1_2)
    i = setToINI("XRAY", "XRAY_key1_2", strInfo, strPath)
    
    strInfo = CStr(XRAY_key2_1)
    i = setToINI("XRAY", "XRAY_key2_1", strInfo, strPath)
    
    strInfo = CStr(XRAY_key2_2)
    i = setToINI("XRAY", "XRAY_key2_2", strInfo, strPath)
    
    strInfo = CStr(XRAY_key3_1)
    i = setToINI("XRAY", "XRAY_key3_1", strInfo, strPath)
    
    strInfo = CStr(XRAY_key3_2)
    i = setToINI("XRAY", "XRAY_key3_2", strInfo, strPath)

    frmAutoeater.UpdatePublicVars
    strInfo = CStr(AutoeaterKey)
    i = setToINI("Autoeater", "AutoeaterKey", strInfo, strPath)
    
    strInfo = CStr(AutoeaterTimerFrom)
    i = setToINI("Autoeater", "AutoeaterTimerFrom", strInfo, strPath)
    
    strInfo = CStr(AutoeaterTimerTo)
    i = setToINI("Autoeater", "AutoeaterTimerTo", strInfo, strPath)



    strInfo = CStr(TibiaExePath)
    i = setToINI("Misc", "TibiaExePath", strInfo, strPath)
    
    SaveSettings = ""
    Exit Function
goterr:
    SaveSettings = "Got error code " & Err.Number & ": " & Err.Description
End Function

Private Sub UpdateFormsFromVars()
    frmMenuAutohealer.cmbLimit(0).Text = CStr(HPmanaLimit0)
    frmMenuAutohealer.cmbLimit(1).Text = CStr(HPmanaLimit1)
    frmMenuAutohealer.cmbLimit(2).Text = CStr(HPmanaLimit2)
    frmMenuAutohealer.cmbLimit(3).Text = CStr(HPmanaLimit3)
    frmMenuAutohealer.cmbHotkey(0).Text = HPmanaAction0
    frmMenuAutohealer.cmbHotkey(1).Text = HPmanaAction1
    frmMenuAutohealer.cmbHotkey(2).Text = HPmanaAction2
    frmMenuAutohealer.cmbHotkey(3).Text = HPmanaAction3
    frmLight.UpdateControlValues
    frmXRAY.ShowCurrentKeys
    frmAutoeater.UpdateALL
End Sub

Private Sub mAutoeater_Click()
    frmAutoeater.Show
End Sub

Private Sub mBuyGold_Click()
    Dim a
    a = ShellExecute(Me.hWnd, "Open", "http://www.blackdtools.com/worldtrade.php?source=BSC", &O0, &O0, SW_NORMAL)

End Sub

Private Sub mCopyright_Click()
    Message_Tittle = "Copyright"
    Message_Message = "(R) Daniel Peña Vázquez, alias Blackd" & vbCrLf & _
    "( www.blackdtools.com )" & vbCrLf & _
    "2010" & vbCrLf & _
    "daniel@blackdtools.com"
    frmMessage.UpdateDisplay
    frmMessage.Show
End Sub


Private Sub mDefineLightDelay_Click()
    Now_we_define = "LightRefreshDelay"
    frmAsk.Show
    frmAsk.UpdateQuestion
End Sub

Private Sub MDIForm_Load()
    Dim strRes As String
    Dim strPath As String
    Dim strFPath As String
    'Dim testit As MapItem4
    'Debug.Print "s=" & Len(testit)
    
        Me.Show

    MustUnload = False
    
    SetAllPrivilegesForMe
    
    InitTibiaPermaLight
    'Unload frmAuth
    TIBIA_LASTPID = 0
    strPath = App.Path
    If Right$(strPath, 1) <> "\" Then
        strPath = strPath & "\"
    End If
    strFPath = strPath & "config.int"
    strRes = LoadConfig(strFPath)
    If strRes <> "" Then
        MsgBox strRes, vbOKOnly + vbCritical, "Error loading " & strFPath
        End
    End If
    strFPath = strPath & "default.ini"
    strRes = LoadSettings(strFPath)
    If strRes <> "" Then
        MsgBox strRes, vbOKOnly + vbCritical, "WARNING, loading " & strFPath
    End If
    strFPath = strPath & LanguageFile
    strRes = LoadLanguage(strFPath)
    If strRes <> "" Then
        MsgBox strRes, vbOKOnly + vbCritical, "Error loading " & strFPath
        End
    End If

    Load frmAsk
    frmAsk.Hide
    Load frmMessage
    frmMessage.Hide
    Load frmMenuAutohealer
    frmMenuAutohealer.Hide
    Load frmLight
    frmLight.Hide
    Load frmXRAY
    frmXRAY.Hide
    Load frmTrue
    frmTrue.Hide
    Load frmAutoeater
    frmAutoeater.Hide
    UpdateFormsFromVars
    
    
    UpdateMenuText
    Me.Show
    With nid
      .cbSize = Len(nid)
      .hWnd = Me.hWnd
      .uId = vbNull
      .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
      .uCallBackMessage = WM_MOUSEMOVE
      .hIcon = Me.Icon
      .szTip = "Blackd Proxy" & vbNullChar
    End With
      Shell_NotifyIcon NIM_ADD, nid
      
    
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

  'this procedure receives the callbacks from the System Tray icon.
  Dim result As Long
  Dim msg As Long
  'the value of X will vary depending upon the scalemode setting
  ' If mdiMenu.for.ScaleMode = vbPixels Then
    'msg = X
  'Else
   msg = x / Screen.TwipsPerPixelX
 ' End If
  
  Select Case msg
  'Case WM_LBUTTONUP        '514 restore form window
  '  Me.WindowState = vbNormal
  '  result = SetForegroundWindow(Me.hWnd)
  '  Me.Hide
  '  Me.Show
  'Case WM_LBUTTONDBLCLK    '515 restore form window
  '  Me.WindowState = vbNormal
  '  result = SetForegroundWindow(Me.hWnd)
  '  Me.Show
  Case WM_RBUTTONUP        '517 display popup menu
    result = SetForegroundWindow(Me.hWnd)
    Me.PopupMenu Me.mPopupSys
  End Select
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MustUnload = True
    Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub MDIForm_Resize()
    If Me.WindowState = 1 Then
        Me.Hide
    End If
End Sub



Private Sub mForum_Click()
    Dim a
    a = ShellExecute(Me.hWnd, "Open", "http://www.blackdtools.com/forum/index.php", &O0, &O0, SW_NORMAL)
End Sub

Private Sub mHPdelay1_Click()
    Now_we_define = "HPmanadelay1"
    frmAsk.Show
    frmAsk.UpdateQuestion
End Sub

Private Sub mHPdelay2_Click()
    Now_we_define = "HPmanadelay2"
    frmAsk.Show
    frmAsk.UpdateQuestion
End Sub

Private Sub mHPrand_Click()
    Now_we_define = "HPrandpercent"
    frmAsk.Show
    frmAsk.UpdateQuestion
End Sub



Private Sub LoadSettingsFromFile()
    On Error GoTo goterr
    
    Dim strRes As String
    Dim strRes2 As String
    Dim sOpen As SelectedFile
    Dim count As Integer
    Dim FileList As String
    
    FileDialog.sFilter = "Ini (*.ini)" & Chr$(0) & "*.ini"
    
    ' See Standard CommonDialog Flags for all options
    FileDialog.flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY
    'FileDialog.sDlgTitle = "Show Open"
    FileDialog.sInitDir = App.Path
    sOpen = ShowOpen(Me.hWnd)
    If Err.Number <> 32755 And sOpen.bCanceled = False Then
        strRes = sOpen.sLastDirectory & sOpen.sFiles(1)
        strRes2 = LoadSettings(strRes)
        If strRes2 = "" Then
            UpdateFormsFromVars
            MsgBox BString(30) & strRes, vbOKOnly + vbInformation, BString(29)
        Else
            MsgBox strRes2, vbOKOnly + vbCritical, "Error loading " & strRes
        End If
    End If
    Exit Sub
            
    
'    Dim strRes As String
'    Dim strRes2 As String
'    cdlg.Filter = "Ini (*.ini) | *.ini"
'    cdlg.InitDir = App.Path
'    cdlg.ShowOpen
'    strRes = cdlg.FileName
'    If strRes = "" Then
'    ' User canceled.
'    Else
'        strRes2 = LoadSettings(strRes)
'        If strRes2 = "" Then
'            UpdateFormsFromVars
'            MsgBox BString(30) & strRes, vbOKOnly + vbInformation, BString(29)
'        Else
'            MsgBox strRes2, vbOKOnly + vbCritical, "Error loading " & strRes
'        End If
'    End If
'    Exit Sub
goterr:
    If Err.Number <> 32755 Then
        MsgBox "Unexpected error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "LoadSettingsFromFile"
    End If
End Sub

Private Sub SaveSettingsToFile()
    On Error GoTo goterr
    
    Dim strRes As String
    Dim strRes2 As String
    Dim sOpen As SelectedFile
    Dim count As Integer
    Dim FileList As String
    
    FileDialog.sFilter = "Ini (*.ini)" & Chr$(0) & "*.ini"
    
    ' See Standard CommonDialog Flags for all options
    FileDialog.flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY
    'FileDialog.sDlgTitle = "Show Open"
    FileDialog.sInitDir = App.Path
    sOpen = ShowSave(Me.hWnd)
    If Err.Number <> 32755 And sOpen.bCanceled = False Then
        strRes = sOpen.sLastDirectory & sOpen.sFiles(1)
        If Right$(strRes, 4) <> ".ini" Then
            strRes = strRes & ".ini"
        End If
        strRes2 = SaveSettings(strRes)
        If strRes2 = "" Then
            MsgBox BString(32) & strRes, vbOKOnly + vbInformation, BString(31)
        Else
            MsgBox strRes2, vbOKOnly + vbCritical, "Error saving " & strRes
        End If
    End If
    Exit Sub
    
    
'    Dim strRes As String
'    Dim strRes2 As String
'    cdlg.Filter = "Ini (*.ini) | *.ini"
'    cdlg.InitDir = App.Path
'    cdlg.ShowSave
'    strRes = cdlg.FileName
'    If strRes = "" Then
'    ' User canceled.
'    Else
'        strRes2 = SaveSettings(strRes)
'        If strRes2 = "" Then
'            MsgBox BString(32) & strRes, vbOKOnly + vbInformation, BString(31)
'        Else
'            MsgBox strRes2, vbOKOnly + vbCritical, "Error loading " & strRes
'        End If
'    End If
'    Exit Sub
goterr:
    If Err.Number <> 32755 Then
        MsgBox "Unexpected error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    End If
    
End Sub

Private Sub mLatestchanges_Click()
    Message_Tittle = BString(11)
    Message_Message = BString(77) & BString(66) & BString(59) & BString(17)
    
    frmMessage.UpdateDisplay
    frmMessage.Show
End Sub


Private Sub mLoadSettings_Click()
    Dim strRes As String
    Dim strPath As String
    Dim strFPath As String
    LoadSettingsFromFile

    strPath = App.Path
    If Right$(strPath, 1) <> "\" Then
        strPath = strPath & "\"
    End If
    strFPath = strPath & LanguageFile
    strRes = LoadLanguage(strFPath)
    If strRes <> "" Then
        MsgBox strRes, vbOKOnly + vbCritical, "Error loading " & strFPath
        End
    End If
    DoEvents
    UpdateMenuText
End Sub

Private Sub mOpenHPmana_Click()
    frmMenuAutohealer.Show
End Sub

Private Sub mOpenLight_Click()
    frmLight.Show
End Sub

Private Sub mPopupShow_Click()
  Dim result As Long
  Me.WindowState = vbNormal
  result = SetForegroundWindow(Me.hWnd)
  Me.Show
End Sub

Private Sub mHideAllTibia_Click()
    SetTibiaClientsVisible False
End Sub

Private Sub mShootFruits_Click()
    Dim a
    a = ShellExecute(Me.hWnd, "Open", "http://shootfruits.com", &O0, &O0, SW_NORMAL)

End Sub

Private Sub mShowAllTibia_Click()
    SetTibiaClientsVisible True
End Sub

Private Sub mReloadDefault_Click()
    Dim strRes As String
    Dim strRes2 As String
    Dim strPath As String
    Dim strFPath As String
    strPath = App.Path
    If Right$(strPath, 1) <> "\" Then
        strPath = strPath & "\"
    End If
    strRes = strPath & "default.ini"
    strRes2 = LoadSettings(strRes)
    If strRes2 = "" Then
        UpdateFormsFromVars
        strFPath = strPath & LanguageFile
        strRes = LoadLanguage(strFPath)
        If strRes <> "" Then
            MsgBox strRes, vbOKOnly + vbCritical, "Error loading " & strFPath
            End
        End If
        UpdateMenuText
        DoEvents
        MsgBox BString(30) & strRes, vbOKOnly + vbInformation, BString(29)
    Else
        MsgBox strRes2, vbOKOnly + vbCritical, "Error loading " & strRes
    End If
    
    
    
    
End Sub

Private Sub mSaveSettings_Click()
    SaveSettingsToFile
End Sub

Private Sub mSetLanguageFile_Click()
    LoadLangFromFile
End Sub

Private Sub LoadLangFromFile()
    On Error GoTo goterr
    
    
    Dim strRes As String
    Dim strRes2 As String
    Dim sOpen As SelectedFile
    Dim count As Integer
    Dim FileList As String
    
    FileDialog.sFilter = "Txt (*.txt)" & Chr$(0) & "*.txt"
    
    ' See Standard CommonDialog Flags for all options
    FileDialog.flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY
    'FileDialog.sDlgTitle = "Show Open"
    FileDialog.sInitDir = App.Path
    sOpen = ShowOpen(Me.hWnd)
    If Err.Number <> 32755 And sOpen.bCanceled = False Then
        strRes = sOpen.sLastDirectory & sOpen.sFiles(1)
        LanguageFile = sOpen.sFiles(1)
        strRes2 = LoadLanguage(strRes)
        If strRes2 = "" Then
            UpdateMenuText
            'MsgBox BString(30) & strRes, vbOKOnly + vbInformation, BString(29)
        Else
            MsgBox strRes2, vbOKOnly + vbCritical, "Error loading " & strRes
        End If
    End If
    Exit Sub
    
'    Dim strRes As String
'    Dim strRes2 As String
'
'    cdlg.Filter = "Txt (*.txt) | *.txt"
'    cdlg.InitDir = App.Path
'    cdlg.ShowOpen
'    strRes = cdlg.FileName
'    If strRes = "" Then
'    ' User canceled.
'    Else
'        LanguageFile = cdlg.FileTitle
'        strRes2 = LoadLanguage(strRes)
'        If strRes2 = "" Then
'            UpdateMenuText
'            'MsgBox BString(30) & strRes, vbOKOnly + vbInformation, BString(29)
'        Else
'            MsgBox strRes2, vbOKOnly + vbCritical, "Error loading " & strRes
'        End If
'    End If
'    Exit Sub
goterr:
    If Err.Number <> 32755 Then
        MsgBox "Unexpected error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "LoadLangFromFile"
    End If
End Sub



Public Function autoGetTibiaFolder() As String
    On Error GoTo goterr
    Dim tpath As String
    tpath = GetProgFolder()
    If Right$(tpath, 1) <> "\" Then
        tpath = tpath & "\"
    End If
    tpath = tpath & DefaultTibiaFolder & "\"
    If MyFileExists(tpath & "Tibia.exe") = True Then
        autoGetTibiaFolder = tpath
    Else
        autoGetTibiaFolder = ""
    End If
    Exit Function
goterr:
    autoGetTibiaFolder = ""
End Function

Private Sub LoadTibiaPathFromFile()
  #If FinalMode Then
    On Error GoTo goterr
  #End If
    
    
    Dim strRes As String
    Dim strRes2 As String
    Dim sOpen As SelectedFile
    Dim count As Integer
    Dim FileList As String
    
    FileDialog.sFilter = "Tibia.exe (Tibia.exe)" & Chr$(0) & "Tibia.exe"
    
    ' See Standard CommonDialog Flags for all options
    FileDialog.flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY
    'FileDialog.sDlgTitle = "Show Open"
    If TibiaExePath = "" Then
        FileDialog.sInitDir = left$(App.Path, 2)
    Else
        FileDialog.sInitDir = TibiaExePath
    End If
    sOpen = ShowOpen(Me.hWnd)
    If Err.Number <> 32755 And sOpen.bCanceled = False Then
        strRes = sOpen.sLastDirectory & sOpen.sFiles(1)
        TibiaExePath = sOpen.sLastDirectory
        MsgBox "Ok. Tibia path redefined: " & vbCrLf & TibiaExePath, vbOKOnly + vbInformation, "Ok"
        
    End If
    Exit Sub
    
'    Dim strRes As String
'    Dim strRes2 As String
'
'    cdlg.Filter = "Txt (*.txt) | *.txt"
'    cdlg.InitDir = App.Path
'    cdlg.ShowOpen
'    strRes = cdlg.FileName
'    If strRes = "" Then
'    ' User canceled.
'    Else
'        LanguageFile = cdlg.FileTitle
'        strRes2 = LoadLanguage(strRes)
'        If strRes2 = "" Then
'            UpdateMenuText
'            'MsgBox BString(30) & strRes, vbOKOnly + vbInformation, BString(29)
'        Else
'            MsgBox strRes2, vbOKOnly + vbCritical, "Error loading " & strRes
'        End If
'    End If
'    Exit Sub
goterr:
    If Err.Number <> 32755 Then
        MsgBox "Unexpected error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "LoadTibiaPathFromFile"
    End If
End Sub


Private Sub mTestFocus_Click()
    MsgBox "This will test if focus is correctly detected. Steps:" & vbCrLf & _
    "1. Open Tibia" & vbCrLf & _
    "2. Press the accept button on this message" & vbCrLf & _
    "3. Select fast your Tibia client" & vbCrLf & _
    "4. Come back to this menu 10 seconds later and you will see the results." & vbCrLf & _
    "5. Send the results to daniel@blackdtools.com", vbOKOnly + vbInformation, "Debug"
    tmrDebug.Enabled = True
End Sub

Private Sub mTibiaPath_Click()
    LoadTibiaPathFromFile
End Sub

Private Sub mTruemap_Click()
    frmTrue.Show
End Sub

Private Sub mXRAY_Click()
    frmXRAY.Show
End Sub

Private Sub tmrDebug_Timer()
    Dim pid As Long
    Dim fid As Long
    tmrDebug.Enabled = False
    Message_Tittle = "Debug result"
    fid = GetForegroundWindow()
    pid = FindWindowEx(0, 0, tibiaclassname, vbNullString)
    Message_Message = "Focus id: " & CStr(fid) & vbCrLf & _
    "Tibia id: " & CStr(pid) & vbCrLf & _
    "Please send this debug result to" & vbCrLf & _
    "daniel@blackdtools.com"
    frmMessage.UpdateDisplay
    frmMessage.Show
End Sub


Private Sub TryReadTiles()
    Dim res As Long
    Dim tibiadathere As String
  TilesAvailable = False
  If TibiaExePath = "" Then
    MsgBox "Unable to find Tibia in default folder, please set the folder manually. Some cheats might not work until that.", vbOKOnly + vbExclamation, "Warning"
    Exit Sub
  End If
  If TibiaVersionLong < 710 Then
    MsgBox "Internal error. Found TibiaVersion=" & CStr(TibiaVersion) & " ; TibiaVersionLong=" & CStr(TibiaVersionLong) & vbCrLf & _
     "Please report this message to daniel@blackdtools.com"
    Exit Sub
  End If
  tibiadathere = TibiaExePath & "tibia.dat"
  If TibiaVersionLong <= 750 Then
    firstValidOutfit = 2
    lastValidOutfit = 142
    res = LoadDatFile(tibiadathere)
  ElseIf TibiaVersionLong < 773 Then
    firstValidOutfit = 2
    lastValidOutfit = 142
    res = LoadDatFile2(tibiadathere)
  ElseIf TibiaVersionLong < 860 Then
    firstValidOutfit = 2
    lastValidOutfit = 160
    res = LoadDatFile3(tibiadathere)
  ElseIf TibiaVersionLong < 872 Then
    firstValidOutfit = 2
    lastValidOutfit = 160
    res = LoadDatFile4(tibiadathere)
  ElseIf TibiaVersionLong < 940 Then
    firstValidOutfit = 2
    lastValidOutfit = 160
    res = LoadDatFile5(tibiadathere)
  ElseIf TibiaVersionLong < 960 Then
    firstValidOutfit = 2
    lastValidOutfit = 160
    res = LoadDatFile6(tibiadathere)
  ElseIf TibiaVersionLong < 994 Then
    firstValidOutfit = 2
    lastValidOutfit = 160
    res = LoadDatFile7(tibiadathere)
  ElseIf TibiaVersionLong < 1021 Then
    firstValidOutfit = 2
    lastValidOutfit = 160
    res = LoadDatFile8(tibiadathere)
  ElseIf TibiaVersionLong < 1050 Then
    firstValidOutfit = 2
    lastValidOutfit = 160
    res = LoadDatFile9(tibiadathere)
  ElseIf TibiaVersionLong < 1058 Then
    firstValidOutfit = 2
    lastValidOutfit = 160
    res = LoadDatFile10(tibiadathere)
  Else
    firstValidOutfit = 2
    lastValidOutfit = 160
    res = LoadDatFile11(tibiadathere)
    End If
  If ((res = -1) Or (res = -2)) Then
    MsgBox "Non compatible tibia.dat file , error " & CStr(res) & vbCrLf & "YOU WILL HAVE TO DOWNLOAD AN UPDATE HERE:" & vbCrLf & "http://www.blackdtools.com/updates.php", vbOKOnly, "You will have to download an update"
    'LogOnFile "debug.txt", "Terminated becouse incompatible tibia.dat (-3)"
    Exit Sub
  End If
  If (res = -3) Then
    MsgBox "Too many tiles found in tibia.dat , please increase MAXDATTILES in your settings.ini" & CStr(res), vbOKOnly, "Problem with config" & CStr(TibiaVersionLong)
    'LogOnFile "debug.txt", "Terminated becouse incompatible tibia.dat (-3)"
    Exit Sub
  End If
  If (res = -4) Then
    MsgBox "Outstanding error -4 while reading tibia.dat: " & vbCrLf & DBGtileError, vbOKOnly, "Problem with config" & CStr(TibiaVersionLong)
    'LogOnFile "debug.txt", "Terminated becouse incompatible tibia.dat (-3)"
    Exit Sub
  End If
  If (res = -5) Then
    MsgBox "Bug caught: " & vbCrLf & DBGtileError, vbOKOnly, "Debug report"
    'LogOnFile "debug.txt", "Terminated becouse incompatible tibia.dat (-3)"
    Exit Sub
  End If
  TilesAvailable = True
End Sub
