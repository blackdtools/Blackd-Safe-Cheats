Attribute VB_Name = "modLanguage"
#Const FinalMode = 1
Option Explicit
Public Const LastBstring As Long = 80
Public BString(0 To LastBstring) As String
Public Sub LoadDefaultStrings()
    BString(0) = "Open Cheats"
    BString(1) = "HP and Mana"
    BString(2) = "Light"
    BString(3) = "Settings"
    BString(4) = "Load settings from file..."
    BString(5) = "HP and Mana"
    BString(6) = "Reload settings from default.ini"
    BString(7) = "Define delay between low status and recharge"
    BString(8) = "Define delay between recharges"
    BString(9) = "Define percent of delay time randomized"
    BString(10) = "About"
    BString(11) = "Latest changes"
    BString(12) = "Copyright"
    BString(13) = "Save settings to file..."
    BString(14) = "Show BSC"
    BString(15) = "Do Nothing"
    BString(16) = "HP and Mana"
    BString(17) = FixVBCRLF("Version 1.0$First version released!$For now it only includes few cheats but they are 100% safe!$Cheats included:$-HP and Mana autorecharger$-Light")
    BString(18) = "Ok"
    BString(19) = "Here you can define up to 6 conditions for HP / mana auto recharge. Remember to prepare the hotkeys in your Tibia Client! The Tibia client should be active and focused for correct recharge."
    BString(20) = "HP recharge"
    BString(21) = "Mana recharge"
    BString(22) = "If HP is less than"
    BString(23) = "then press key"
    BString(24) = "If Mana is less than"
    BString(25) = "Define the value..."
    BString(26) = "Please define the delay that should happen between the moment when recharge is needed and the moment when the key is pressed (at the start of a recharge process) Time goes in milliseconds"
    BString(27) = "Please define the delay between recharges when a lot of recharges are needed. Time goes in milliseconds"
    BString(28) = "Please define the percent of delay that should be randomized to look more human. Choose from 0 to 50 %"
    BString(29) = "Settings loaded"
    BString(30) = "Loaded sucesfully:"
    BString(31) = "Settings saved"
    BString(32) = "Saved sucesfully:"
    BString(33) = "Set language file..."
    BString(34) = "Please define the delay between each light update (if light is enabled) Time goes in milliseconds."
    BString(35) = "Save"
    BString(36) = "Cancel"
    BString(37) = "Light"
    BString(38) = "Define time between each light update"
    BString(39) = "Links"
    BString(40) = "Buy Tibia gold"
    BString(41) = "blackdtools.com forum"
    BString(42) = "Show Tibia"
    BString(43) = "Hide Tibia"
    BString(44) = "Enable light in focused Tibia client"
    BString(45) = "Light level:"
    BString(46) = "Light colour ID:"
    
    BString(47) = "<nothing>"
    BString(48) = "No Tibia client connected!"
    BString(49) = "Click on textboxes and press a key to define a hotkey for virtual floor change (inspect floors above or below you)"
    BString(50) = "Key 1"
    BString(51) = "Key 2 (optional)"
    BString(52) = "Floors:"
    BString(53) = "Floor above:"
    BString(54) = "Floor below:"
    BString(55) = "Reset:"
    BString(56) = "Test"
    BString(57) = "Untested Cheats!"
    BString(58) = "XRAY"
    BString(59) = FixVBCRLF("Version 1.1$Added XRAY. It will allow you to inspect floors above or below. Up to 7 floors above and up to 2 floors below. Restriction: at base floor level you won't be able to see underground floors (by Tibia design)$")
    BString(60) = "<PRESS KEY>"
    BString(61) = "KEY"
    
    BString(62) = "Truemap"
    BString(63) = "Watch selected floor"
    BString(64) = "Watch my floor"
    BString(65) = "Tibia client could not be found"
    BString(66) = FixVBCRLF("Version 1.2$Added Truemap. Preview floors above or below without screen modification (without risk of crash)$")


    BString(67) = "Autoeater"
    BString(68) = "Use this tool to eat food in Tibia from time to time or to do something else and act as anti-iddle."
    BString(69) = "To avoid detection, the action should be reapeated randomly in a variable time..."
    BString(70) = "from"
    BString(71) = "seconds"
    BString(72) = "to"
    BString(73) = "Current timer:"
    BString(74) = "Tibia hotkey to press:"
    BString(75) = "Warning: the hotkey only will be pressed if a Tibia window is focused!"
    BString(76) = "APPLY"
    
    BString(77) = FixVBCRLF("S77=Version 1.3$Added Autoeater. It allows eating food by pressing Tibia hotkeys. It can also act as anti-iddle.$")
End Sub

Public Function FixVBCRLF(strVar As String) As String
    On Error GoTo goterr
    Dim res As String
    If strVar = "" Then
        res = ""
    Else
        res = Replace(strVar, "$", vbCrLf)
    End If
    FixVBCRLF = res
    Exit Function
goterr:
    MsgBox "Got error code " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error at FixVBCRLF"
    End
End Function

Public Sub UpdateMenuText()
    Dim i As Long
    mdiMenu.mCheats.Caption = BString(0)
    mdiMenu.mOpenHPmana.Caption = BString(1)
    mdiMenu.mOpenLight.Caption = BString(2)
    mdiMenu.mSettings.Caption = BString(3)
    mdiMenu.mLoadSettings.Caption = BString(4)
    mdiMenu.mSettingsHPmana.Caption = BString(5)
    mdiMenu.mReloadDefault.Caption = BString(6)
    mdiMenu.mHPdelay1.Caption = BString(7)
    mdiMenu.mHPdelay2.Caption = BString(8)
    mdiMenu.mHPrand.Caption = BString(9)
    mdiMenu.mAbout.Caption = BString(10)
    mdiMenu.mLatestchanges.Caption = BString(11)
    mdiMenu.mCopyright.Caption = BString(12)
    mdiMenu.mSaveSettings.Caption = BString(13)
    mdiMenu.mPopupShow.Caption = BString(14)
    mdiMenu.mDoNothing.Caption = BString(15)
    frmMenuAutohealer.Caption = BString(16)
    frmAsk.Hide
    frmMessage.Hide
    frmMessage.cmdOk.Caption = BString(18)
    frmMenuAutohealer.lblMainLabel = BString(19)
    frmMenuAutohealer.fraHP.Caption = BString(20)
    frmMenuAutohealer.fraMana.Caption = BString(21)
    For i = 0 To 2
        frmMenuAutohealer.lblIf(i).Caption = BString(22)
        frmMenuAutohealer.lblThen(i).Caption = BString(23)
    Next i
    For i = 3 To 5
        frmMenuAutohealer.lblIf(i).Caption = BString(24)
        frmMenuAutohealer.lblThen(i).Caption = BString(23)
    Next i
    mdiMenu.mSetLanguageFile.Caption = BString(33)
    frmLight.Caption = BString(2)
    mdiMenu.mSettingsLight.Caption = BString(37)
    mdiMenu.mDefineLightDelay.Caption = BString(38)
    mdiMenu.mLinks.Caption = BString(39)
    mdiMenu.mBuyGold.Caption = BString(40)
    mdiMenu.mForum.Caption = BString(41)
    mdiMenu.mShowAllTibia.Caption = BString(42)
    mdiMenu.mHideAllTibia.Caption = BString(43)
    frmLight.chkEnable.Caption = BString(44)
    frmLight.lblIntensity.Caption = BString(45)
    frmLight.lblColourID.Caption = BString(46)
    
    frmXRAY.UpdateXRAY_Language
    frmXRAY.ShowCurrentKeys
    mdiMenu.mUntested.Caption = BString(57)
    mdiMenu.mXRAY.Caption = BString(58)
    mdiMenu.mTruemap.Caption = BString(62)
    frmTrue.UpdateLanguage
    
    mdiMenu.mAutoeater.Caption = BString(67)
    frmAutoeater.UpdateLanguage
End Sub
