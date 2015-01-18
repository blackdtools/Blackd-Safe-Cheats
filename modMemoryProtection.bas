Attribute VB_Name = "modMemoryProtection"
#Const FinalMode = 1
Option Explicit

Private Declare Function apiGetClassName Lib "user32" Alias _
                "GetClassNameA" (ByVal hWnd As Long, _
                ByVal lpClassName As String, _
                ByVal nMaxCount As Long) As Long
Private Declare Function apiGetDesktopWindow Lib "user32" Alias _
                "GetDesktopWindow" () As Long
Private Declare Function apiGetWindow Lib "user32" Alias _
                "GetWindow" (ByVal hWnd As Long, _
                ByVal wCmd As Long) As Long
Private Declare Function apiGetWindowLong Lib "user32" Alias _
                "GetWindowLongA" (ByVal hWnd As Long, ByVal _
                nIndex As Long) As Long
Private Declare Function apiGetWindowText Lib "user32" Alias _
                "GetWindowTextA" (ByVal hWnd As Long, ByVal _
                lpString As String, ByVal aint As Long) As Long
Private Const mcGWCHILD = 5
Private Const mcGWHWNDNEXT = 2
Private Const mcGWLSTYLE = (-16)
Private Const mcWSVISIBLE = &H10000000
Private Const mconMAXLEN = 255

Public Function DoTheCheck() As Boolean
  Dim lngx As Long
  Dim lngLen As Long
  Dim lngStyle As Long
  Dim strCaption As String
  Dim currClass As String
  Dim currCaption As String
  Dim currCapLen As Long
  Dim res As Boolean
  'DoTheCheck = False
  'Exit Function
  res = False
  lngx = apiGetDesktopWindow()
  'Return the first child to Desktop
  lngx = apiGetWindow(lngx, mcGWCHILD)
  Do While Not lngx = 0
    strCaption = fGetCaption(lngx)
    If Len(strCaption) > 0 Then
      lngStyle = apiGetWindowLong(lngx, mcGWLSTYLE)
      'enum visible windows only
      If lngStyle And mcWSVISIBLE Then
        currClass = fGetClassName(lngx)
        currCaption = fGetCaption(lngx)
        currCaption = LCase(currCaption)
        currCapLen = Len(currCaption)
        If currCapLen > 2 Then
          If left(currCaption, 3) = "wpe" Then
            res = True
          End If
        End If
        If currCapLen > 6 Then
          If left(currCaption, 7) = "tsearch" Then
            res = True
          End If
        End If
      End If
    End If
    lngx = apiGetWindow(lngx, mcGWHWNDNEXT)
  Loop
  DoTheCheck = res
End Function

Function fEnumWindows()
Dim lngx As Long
Dim lngLen As Long
Dim lngStyle As Long
Dim strCaption As String
    
    lngx = apiGetDesktopWindow()
    'Return the first child to Desktop
    lngx = apiGetWindow(lngx, mcGWCHILD)
    
    Do While Not lngx = 0
        strCaption = fGetCaption(lngx)
        If Len(strCaption) > 0 Then
            lngStyle = apiGetWindowLong(lngx, mcGWLSTYLE)
            'enum visible windows only
            If lngStyle And mcWSVISIBLE Then
                'Debug.Print "Class = " & fGetClassName(lngx),
                'Debug.Print "Caption = " & fGetCaption(lngx)
            End If
        End If
        lngx = apiGetWindow(lngx, mcGWHWNDNEXT)
    Loop
End Function
Private Function fGetClassName(hWnd As Long) As String
    Dim strBuffer As String
    Dim intCount As Integer
   
    strBuffer = String$(mconMAXLEN - 1, 0)
    intCount = apiGetClassName(hWnd, strBuffer, mconMAXLEN)
    If intCount > 0 Then
        fGetClassName = left$(strBuffer, intCount)
    End If
End Function

Private Function fGetCaption(hWnd As Long) As String
    Dim strBuffer As String
    Dim intCount As Integer

    strBuffer = String$(mconMAXLEN - 1, 0)
    intCount = apiGetWindowText(hWnd, strBuffer, mconMAXLEN)
    If intCount > 0 Then
        fGetCaption = left$(strBuffer, intCount)
    End If
End Function







