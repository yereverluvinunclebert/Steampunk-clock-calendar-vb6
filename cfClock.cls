VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cfClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule IntegerDataType, ModuleWithoutFolder
Option Explicit 'simple Alpha-Form (...to be filled with (mostly Img-) Content via .Widgets.Add)

Public WithEvents clockForm As cWidgetForm
Attribute clockForm.VB_VarHelpID = -1

Public collPSDNonUIElements As cCollection, FX As Long, FY As Long, FZ As Single  'available on the outside, even after the Form was closed

Private PSDWidth As Long
Private PSDHeight As Long

Private mMute As Integer ' property
'Private mSystemAudioLevel As Currency ' property
'
'Private RenderDev As cMMDevice
'Private WithEvents SystemVolume As cSimpleAudioVolume 'use cSimpleAudioVoume with events

'Private lblPercent As cwLabel
'Private lblPercentShadow As cwLabel

'Slider-related Private-Vars
Private WithEvents SliderSet As cWidgetBase
Attribute SliderSet.VB_VarHelpID = -1
Private sldFixedTop As Double
Private sldXMax As Double
Private sldXMin As Double
Private sldX As Double
Private mVolumePerc As Currency

'Public WithEvents tmrSampleAudioVolume As cTimer
'Public WithEvents tmrSampleAudioMute As cTimer

'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : olaf schmidt
' Date      : 31/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
   On Error GoTo Class_Initialize_Error

    Set collPSDNonUIElements = New_c.Collection(False)
    
    'retrieve (and Set) the Default-Renderdevice
'    Set RenderDev = New_c.MMDeviceEnumerator.GetDefaultAudioEndpoint(eRender, eMultimedia)
'
'    'retrieve (and Set) a cSimpleAudioVolume instance, to get access to Volume-Handling
'    Set SystemVolume = RenderDev.Activate(AudioEndpointVolume)
'
'    Set tmrSampleAudioVolume = Nothing
'    Set tmrSampleAudioVolume = New_c.Timer(1000, True) ' interval is one second.
'
'    Set tmrSampleAudioMute = Nothing
'    Set tmrSampleAudioMute = New_c.Timer(1000, True) ' interval is one second.

   On Error GoTo 0
   Exit Sub

Class_Initialize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cfClock"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : InitFromPSD
' Author    : olaf schmidt
' Date      : 31/07/2023
' Purpose   :    loads up the Cairo imageList with keyed entries and a surface from each and every PSD layer
'                adds path X,Y and alpha to the excluded collection
'                creates widgets with corresponding keynames, locations and sizes as per PSD for each layer
'                cwOverlay class puts it on the screen
'---------------------------------------------------------------------------------------
'
Public Sub InitFromPSD(ByRef PSD_FileNameOrByteArray As Variant)
    
    Const AlphaWithTaskbarEntry As Integer = 6
    Const AlphaNoTaskbarEntry As Integer = 7
  
    Dim I As Long: I = 0
    Dim W As cWidgetBase
   
    On Error GoTo InitFromPSD_Error

    'create the Alpha-capable Top-Level-Form
    Set clockForm = Cairo.WidgetForms.Create(IIf(App.LogMode, AlphaNoTaskbarEntry, AlphaWithTaskbarEntry), widgetName, True, 1, 1)
        clockForm.WidgetRoot.BackColor = -1
 
    With New_c.SimplePSD(PSD_FileNameOrByteArray)  'create a new PSD-Parser.instance (and load the passed content)
        PSDWidth = .Width
        PSDHeight = .Height       'store the original Psd-Pixel-Width/Height in Private vars (as the base from which we calculate the zoomed Form-Width/Height)
    
        For I = 0 To .LayersCount - 1 ' loop through each of the Layers in the PSD
            If .LayerByteSize(I) Then  ' this is a true Alpha-Surface-Layer and not just a PSD layer 'group'
                If .LayerAlphaPercent(I) > 0 Then ' only handles layers that have an opacity greater than 0 - need to note this for the future, this will cause a problem!
                    'add each current Layer path and surface object into the global ImageList collection (using LayerPath as the ImageKey)
                    Cairo.ImageList.AddSurface .LayerPath(I), .LayerSurface(I)
                    
                    If collPSDNonUIElements.Exists(.LayerPath(I)) Then ' we do not create Widgets for those in this exclude list.
                        
                        'we add layer info. (used in cwOverlay) to the "excluded paths" previously defined in Sub Main,
                        collPSDNonUIElements(.LayerPath(I)) = Array(.LayerX(I), .LayerY(I), Val(gblOpacity) / 100)  'here we update the so far empty slots with the PSD-offsets
                
                    Else 'everything not in the exclude-list will get its own Widget-instance
                        Set W = clockForm.Widgets.Add(New cwAlphaImg, LCase$(.LayerPath(I)), .LayerX(I), .LayerY(I), .LayerWidth(I), .LayerHeight(I)).Widget
                        'Set W = clockForm.Widgets.Add(New_W("cwAlphaImg"), LCase$(.LayerPath(I)), .LayerX(I), .LayerY(I), .LayerWidth(I), .LayerHeight(I)).Widget
                        W.ImageKey = W.Key 'W.Key equals ImageList-Key, set above - and LayerPath(i) at this point ... set it also as the ImageKey of our new created Widget

                        W.Alpha = 0
                    
                        ' note: the clickable layers characteristics are set in adjustMainControls
                        
                        ' all non-clickable Layer-Widgets will be -1 or "non-hoverable" and "fully click-through"
                        W.HoverColor = -1 ' task: might change this later when new ver or RC6 arrives
                        If gblEnableTooltips = "1" Then W.ToolTip = "Ctrl + mouse scrollwheel up/down to resize, you can also drag me to a new position."
                        W.MousePointer = IDC_SIZEALL

                    End If
                End If
            End If
        Next I
    End With '<-- the Parser-instance will be destroyed here (freeing the Memory, the internal PSD-Layers have occupied)
        
    Set overlayWidget = clockForm.Widgets.Add(New cwOverlay, "Overlay", 0, 0, PSDWidth, PSDHeight)
    Set overlayWidget.cCollOverlayImages = collPSDNonUIElements 'make the excluded PSD-Surface-Paths known to the Overlay-Widget
        overlayWidget.cxPerc = 0.4940476 'define the relative center (within cwOverlay, which has the same size as the Form and the PSD...) this is used for clock-hand-positioning
        overlayWidget.cyPerc = 0.4727272
        
    'the next 5 lines will init a "moving Slider" (which was defined as "sliderset"-layer in the PSD)
    Set SliderSet = clockForm.Widgets("sliderset").Widget 'store the already (in PSD-Loop created) Widget-reference in an explicit "WithEvents-Variable"
        
    With SliderSet
        .Moveable = True 'make it "moveable"
        .HoverColor = 0
        .MousePointer = IDC_HAND
        .Tag = 0.01
    End With
    
    sldFixedTop = SliderSet.Top   'and store the initial position-data for y-direction (which will not be changed)
    sldXMax = SliderSet.Left      'as well as the x-direction related pos (which is initially at the "100%-Volume, right-endswitch" )
    sldXMin = sldXMax - 300       'the Min-Value is simply the Max-Value minus Slider-Range (the 300pxl-range was determined by "probing")

   On Error GoTo 0
   Exit Sub

InitFromPSD_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InitFromPSD of Class Module cfClock"
End Sub
'
'---------------------------------------------------------------------------------------
' Procedure : clockForm_BubblingEvent
' Author    : beededea
' Date      : 31/07/2023
' Purpose   : central handling of all Click-Events
'---------------------------------------------------------------------------------------
'
Private Sub clockForm_BubblingEvent(ByRef Sender As Object, ByRef EventName As String, ByRef P1 As Variant, ByRef P2 As Variant, ByRef P3 As Variant, ByRef P4 As Variant, ByRef P5 As Variant, ByRef P6 As Variant, ByRef P7 As Variant)
   
    Dim stdMessage As String: stdMessage = vbNullString
    Dim stdSubject As String: stdSubject = vbNullString
    
    On Error GoTo clockForm_BubblingEvent_Error

    If gblIgnoreMouse = "1" Then Exit Sub
    
    stdMessage = "Right Click to open the menu and the preferences. CTRL+ Mouse scrollwheel UP/DOWN to resize. "
    stdSubject = "You can turn off the balloon tooltips in the preferences."
    
    If EventName = "W_MouseMove" Then
        Select Case LCase$(Sender.Widget.Key)
            Case "helppin"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "This button opens the one-page help file for this widget. There is a further detailed CHM help file available from the right-click menu that contains all the information that you need to run this widget.", _
                     TTIconInfo, widgetName & " Help Pin Help", , , , True
            Case "lockingpin"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "Click me to lock the widget in place. This lock can be released by clicking again on the locking pin, or in the preferences. A locking sound will denote the button's operation.", _
                     TTIconInfo, widgetName & " Locking pin Help", , , , True
            Case "bell"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "This is the mute button. Click me to toggle sound on/off.", _
                     TTIconInfo, widgetName & " Sound Toggle Help", , , , True
            Case "lockingpinunlocked"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "Click me to lock the widget in place. This lock can be released by clicking again on the locking pin, or in the preferences. A locking sound will denote the button's operation.", _
                     TTIconInfo, widgetName & " Locking pin Help", , , , True
            Case "indicatorred"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "This is the mute indicator. A green lamp means that the sound is unmuted. Red lamp means that mute is ON.", _
                     TTIconInfo, widgetName & " Locking pin Help", , , , True
            Case "indicatorgreen"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "This is the mute indicator. Sound is unmuted.", _
                     TTIconInfo, widgetName & " Locking pin Help", , , , True
            Case "sliderset"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "This is the Time Slider. Click and drag me left/right to alter time.", _
                     TTIconInfo, widgetName & " Locking pin Help", , , , True

            Case "speaker", "cable", "bar", "pipes", "cablewheelset"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, stdMessage _
                    & vbCrLf & stdSubject, _
                    TTIconInfo, widgetName & " Configuration Help", , , , True
        End Select
    End If
    
    If EventName = "W_MouseDown" Then
        Select Case LCase$(Sender.Widget.Key)
            Case "lockingpin"
                Call lockingpin_mouseDown(Sender)
            Case "helppin"
                Call helppin_mouseDown(Sender)
            Case "lockingpinunlocked"
                Call lockingpin_mouseDown(Sender)
            Case "bell"
                'Call bell_mouseDown(Sender)
            Case "sliderset"
                overlayWidget.Locked = True
                'fClock.tmrSampleAudioVolume.Enabled = False
            Case Else
                Debug.Print Sender.Widget.Key
        End Select
    End If
    
    If EventName = "W_MouseUp" Then
        Select Case LCase$(Sender.Widget.Key)
            Case "bell"
                'Call bell_mouseUp(Sender)
            Case "sliderset"
                overlayWidget.Locked = False
                'fClock.tmrSampleAudioVolume.Enabled = True
            Case Else
                Debug.Print Sender.Widget.Key
        End Select
    End If
   
   On Error GoTo 0
   Exit Sub

clockForm_BubblingEvent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure clockForm_BubblingEvent of Class Module cfClock"
End Sub




'---------------------------------------------------------------------------------------
' Procedure : lockingpin_mouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lockingpin_mouseDown(ByRef Sender As Object)
   On Error GoTo lockingpin_mouseDown_Error
   
    Dim fileToPlay As String: fileToPlay = vbNullString

    fileToPlay = "lock.wav"
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    If overlayWidget.Locked = False Then
        Sender.Widget.Alpha = 0
        fClock.clockForm.Widgets("lockingpinunlocked").Widget.Alpha = Val(gblOpacity) / 100
    Else
        Sender.Widget.Alpha = Val(gblOpacity) / 100
        fClock.clockForm.Widgets("lockingpinunlocked").Widget.Alpha = 0
    End If
    Sender.Widget.Refresh ' this removes the delay in the lockingpin depressing
    Call lockWidget
    
   On Error GoTo 0
   Exit Sub

lockingpin_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lockingpin_mouseDown of Class Module cfClock"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : helppin_mouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub helppin_mouseDown(ByRef Sender As Object)
   On Error GoTo helppin_mouseDown_Error
   
    Dim fileToPlay As String: fileToPlay = vbNullString

    fileToPlay = "lock.wav"
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
'    If overlayWidget.Locked = False Then
'        Sender.Widget.Alpha = 0
'        fClock.clockForm.Widgets("helppinunlocked").Widget.Alpha = Val(gblOpacity) / 100
'    Else
'        Sender.Widget.Alpha = Val(gblOpacity) / 100
'        fClock.clockForm.Widgets("helppinunlocked").Widget.Alpha = 0
'    End If
    Sender.Widget.Refresh ' this removes the delay in the helppin depressing
    
    Call helpSplash
    
   On Error GoTo 0
   Exit Sub

helppin_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure helppin_mouseDown of Class Module cfClock"

End Sub

''---------------------------------------------------------------------------------------
'' Procedure : bell_mouseDown
'' Author    : beededea
'' Date      : 03/08/2023
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub bell_mouseDown(ByRef Sender As Object)
'   On Error GoTo bell_mouseDown_Error
'
'    Dim fileToPlay As String: fileToPlay = vbNullString
'
'    'tmrSampleAudioMute.Enabled = False
'
'    fileToPlay = "buzzer.wav"
'
'    If Mute = 1 Then
'        Mute = 0
'    Else
'        Mute = 1
'    End If
'
'    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
'        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
'    End If
'
'    Sender.Widget.Alpha = 0
'    Sender.Widget.Refresh ' this removes the delay in the bell depressing
'
''    tmrSampleAudioMute.Enabled = True
'
'   On Error GoTo 0
'   Exit Sub
'
'bell_mouseDown_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure bell_mouseDown of Class Module cfClock"
'
'End Sub


''---------------------------------------------------------------------------------------
'' Procedure : bell_mouseUp
'' Author    : beededea
'' Date      : 03/08/2023
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub bell_mouseUp(ByRef Sender As Object)
'   On Error GoTo bell_mouseUp_Error
'
'    Sender.Widget.Alpha = Val(gblOpacity) / 100
'    Sender.Widget.Refresh ' this removes the delay in the bell depressing
'
'   On Error GoTo 0
'   Exit Sub
'
'bell_mouseUp_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure bell_mouseUp of Class Module cfClock"
'
'End Sub





'---------------------------------------------------------------------------------------
' Procedure : clockForm_DblClick
' Author    : beededea
' Date      : 05/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub clockForm_DblClick()
    Dim userprof As String: userprof = vbNullString
    Dim thisCommand As String: thisCommand = vbNullString
    
    On Error GoTo clockForm_DblClick_Error

    '    If gblIgnoreMouse = "1" Then Exit Sub

    If LTrim$(gblDblClickCommand) = vbNullString Then Exit Sub
    
    thisCommand = gblDblClickCommand
        
    If InStr(thisCommand, "%userprofile%") Then
        userprof = Environ$("USERPROFILE")
        thisCommand = Replace(thisCommand, "%userprofile%", userprof)
    End If
    
    ' .91 DAEB 08/12/2022 frmMain.frm SteamyDock responds to %systemroot% environment variables during runCommand
    If InStr(thisCommand, "%systemroot%") Then
        userprof = Environ$("SYSTEMROOT")
        thisCommand = Replace(thisCommand, "%systemroot%", userprof)
    End If
    
    If SHIFT_1 = True Then
        SHIFT_1 = False
        Call ShellExecute(fClock.clockForm.hwnd, "Open", gblOpenFile, vbNullString, App.path, 1)
    Else
        Call ShellExecute(fClock.clockForm.hwnd, "runas", thisCommand, vbNullString, App.path, 1)
    End If

    On Error GoTo 0
    Exit Sub

clockForm_DblClick_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure clockForm_DblClick of Class Module cfMain"
            Resume Next
          End If
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : clockForm_KeyDown
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : get F5 and SHIFT keypresses
'---------------------------------------------------------------------------------------
'
Private Sub clockForm_KeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer)
    On Error GoTo clockForm_KeyDown_Error

    Call getKeyPress(KeyCode, Shift) ' will not catch Shift or CTRL keypresses on their own, must be combination with another key.

    On Error GoTo 0
    Exit Sub

clockForm_KeyDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure clockForm_KeyDown of Class Module cfMain"
End Sub
''
'---------------------------------------------------------------------------------------
' Procedure : clockForm_MouseUp
' Author    : beededea
' Date      : 05/05/2023
' Purpose   : save the form x,y position when ever the globe/form is dragged
'---------------------------------------------------------------------------------------
'
Private Sub clockForm_MouseUp(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
    On Error GoTo clockForm_MouseUp_Error
    
    Call savePosition

    On Error GoTo 0
    Exit Sub

clockForm_MouseUp_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure clockForm_MouseUp of Class Module cfMain"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : clockForm_MouseMove
' Author    : beededea
' Date      : 31/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub clockForm_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
  Static x0 As Single, y0 As Single: If Button = 0 Then x0 = x: y0 = y 'just store the offsets when no button is down
  
   On Error GoTo clockForm_MouseMove_Error
    
    If overlayWidget.Locked = True Then Exit Sub
    If gblIgnoreMouse = "1" Then Exit Sub

    If Button = vbLeftButton And Not clockForm.ActiveWidget Is Nothing Then 'Form-Dragging (under certain conditions)

      FX = (x - x0) * clockForm.WidgetRoot.Zoom + clockForm.Left
      FY = (y - y0) * clockForm.WidgetRoot.Zoom + clockForm.Top
      clockForm.Move FX, FY
    End If


   On Error GoTo 0
   Exit Sub

clockForm_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure clockForm_MouseMove of Class Module cfClock"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : clockForm_MouseWheel
' Author    : beededea
' Date      : 31/07/2023
' Purpose   : rotate forward positive = smaller
'---------------------------------------------------------------------------------------
'
Private Sub clockForm_MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Single, ByVal Ypos As Single)
   Dim RotationDirection As Long: RotationDirection = 0
   
   On Error GoTo clockForm_MouseWheel_Error

    If gblScrollWheelDirection = "1" Then
        If Rotation > 0 Then
            RotationDirection = 0 - Abs(Rotation)
        Else
            RotationDirection = Abs(Rotation)
        End If
    Else
        RotationDirection = Rotation
    End If
    
   If MouseKeys = 8 Then Call AdjustZoom(FZ + RotationDirection / 2400)   'change the Zoom only, when the Ctrl-Key is down

   On Error GoTo 0
   Exit Sub

clockForm_MouseWheel_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure clockForm_MouseWheel of Class Module cfClock"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : AdjustZoom
' Author    : olaf schmidt
' Date      : 31/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub AdjustZoom(ByVal NewZoom As Single)
   On Error GoTo AdjustZoom_Error

    FZ = NewZoom
    If FZ < 0.05 Then FZ = 0.05 Else If FZ > 3 Then FZ = 3 'keep FZ within a sane ZoomInterval
    clockForm.WidgetRoot.Zoom = FZ '* clockForm.WidgetRoot.CurrentMonitor.Zoom '<-- if the multiplicator is activated, we'd have DPI-awareness on high-res monitors
    clockForm.Move clockForm.Left, clockForm.Top, PSDWidth * clockForm.WidgetRoot.Zoom, PSDHeight * clockForm.WidgetRoot.Zoom
    
    ' when resized from the gauge scroll up/down it needs to write it back so the size is correct when prefs opened.
    gblGaugeSize = LTrim$(Str$(NewZoom * 100))
    If widgetPrefs.IsVisible = True Then widgetPrefs.sliGaugeSize = Val(gblGaugeSize)

   On Error GoTo 0
   Exit Sub

AdjustZoom_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure AdjustZoom of Class Module cfClock"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : W_MouseDown
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub clockForm_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo W_MouseDown_Error

    If Button = vbRightButton Then
        Call menuForm.PopupMenu(menuForm.mnuMainMenu)
    Else
        If gblIgnoreMouse = "1" Then Exit Sub
        ' do whatever you want with a mouseDown here, not doing anything at the moment but we will...
    End If

   On Error GoTo 0
   Exit Sub

W_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_MouseDown of Class Module cwOverlay"
End Sub

' ----------------------------------------------------------------
' Procedure Name: SliderSet_Moving
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 26/05/2024
' ----------------------------------------------------------------
Private Sub SliderSet_Moving() 'this is an explcit Event-Handler on the single Widget-instance (derived in the routine above)
    On Error GoTo SliderSet_Moving_Error
     sldX = SliderSet.Left
     If sldX < sldXMin Then sldX = sldXMin Else If sldX > sldXMax Then sldX = sldXMax 'keep sldX "in range"
    
     VolumePerc = (sldX - sldXMin) / (sldXMax - sldXMin) 'assign the new Value to the Public Prop, to trigger all the necessary refreshs
    
    On Error GoTo 0
    Exit Sub

SliderSet_Moving_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SliderSet_Moving, line " & Erl & "."

End Sub
''---------------------------------------------------------------------------------------
'' Property  : SystemAudioLevel
'' Author    : beededea
'' Date      : 17/05/2023
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Property Let SystemAudioLevel(ByVal newValue As Currency)
'   On Error GoTo SystemAudioLevelLet_Error
'
'   If mSystemAudioLevel <> newValue Then mSystemAudioLevel = newValue Else Exit Property
'
'   On Error GoTo 0
'   Exit Property
'
'SystemAudioLevelLet_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property SystemAudioLevel of Class Module cwGlobe"
'End Property
'
'' ----------------------------------------------------------------
'' Procedure Name: SystemAudioLevel
'' Purpose:
'' Procedure Kind: Property (Get)
'' Procedure Access: Public
'' Return Type: Currency
'' Author: beededea
'' Date: 26/05/2024
'' ----------------------------------------------------------------
'Public Property Get SystemAudioLevel() As Currency
'
'    On Error GoTo SystemAudioLevel_Error
'
'    mSystemAudioLevel = SystemVolume.GetMasterVolume
'    SystemAudioLevel = mSystemAudioLevel
'
'
'    On Error GoTo 0
'    Exit Property
'
'SystemAudioLevel_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SystemAudioLevel, line " & Erl & "."
'
'End Property

''---------------------------------------------------------------------------------------
'' Property  : Mute
'' Author    : beededea
'' Date      : 10/05/2023
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Property Let Mute(ByVal newValue As Integer)
'   On Error GoTo MuteLet_Error
'
'    If mMute <> newValue Then mMute = newValue Else Exit Property
'
'    Mute = mMute
'
'    SystemVolume.SetMute mMute
'
'    Call setLamps
'
'   On Error GoTo 0
'   Exit Property
'
'MuteLet_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Mute of Class Module cwGlobe"
'End Property
'
''---------------------------------------------------------------------------------------
'' Property  : Mute
'' Author    : beededea
'' Date      : 17/05/2023
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Property Get Mute() As Integer
'   On Error GoTo MuteGet_Error
'
'    Mute = mMute
'
'    SystemVolume.SetMute mMute
'
'    Call setLamps
'
'   On Error GoTo 0
'   Exit Property
'
'MuteGet_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Mute of Class Module cwGlobe"
'End Property


' ----------------------------------------------------------------
' Procedure Name: VolumePerc
' Purpose:
' Procedure Kind: Property (Get)
' Procedure Access: Public
' Return Type: Currency
' Author: beededea
' Date: 26/05/2024
' ----------------------------------------------------------------
Public Property Get VolumePerc() As Currency
    
    On Error GoTo VolumePerc_Error
    
    VolumePerc = mVolumePerc
    
    On Error GoTo 0
    Exit Property

VolumePerc_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure VolumePerc, line " & Erl & "."

End Property

' ----------------------------------------------------------------
' Procedure Name: VolumePerc
' Purpose:
' Procedure Kind: Property (Let)
' Procedure Access: Public
' Parameter RHS (Currency):
' Author: beededea
' Date: 26/05/2024
' ----------------------------------------------------------------
Public Property Let VolumePerc(ByVal RHS As Currency)
    On Error GoTo VolumePerc_Error
    
    If RHS < 0 Then RHS = 0 Else If RHS > 1 Then RHS = 1 'keep RHS "in range"
    sldX = RHS * (sldXMax - sldXMin) + sldXMin
    overlayWidget.sldCurWidth = sldXMax - sldX 'inform the Overlay about the new current "Cable-Width"
    SliderSet.Move sldX, sldFixedTop 'move the SliderSet-widget accordingly
    If mVolumePerc <> RHS Then mVolumePerc = RHS Else Exit Property  'only refresh, when there was a change to the previous value
    
    If Not Cairo.ToolTipForm Is Nothing Then Cairo.ToolTipForm.Visible = False
    clockForm.Refresh
    
'    Call HandleVolumeChangeForHardware
    
    On Error GoTo 0
    Exit Property

VolumePerc_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure VolumePerc, line " & Erl & "."

End Property


'' ----------------------------------------------------------------
'' Procedure Name: HandleVolumeChangeForHardware
'' Purpose:
'' Procedure Kind: Sub
'' Procedure Access: Private
'' Author: beededea
'' Date: 27/05/2024
'' ----------------------------------------------------------------
'Private Sub HandleVolumeChangeForHardware()
'
'    On Error GoTo HandleVolumeChangeForHardware_Error
'
'    SystemVolume.SetMasterVolume VolumePerc
'
'    On Error GoTo 0
'    Exit Sub
'
'HandleVolumeChangeForHardware_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure HandleVolumeChangeForHardware, line " & Erl & "."
'
'End Sub


    
''---------------------------------------------------------------------------------------
'' Procedure : tmrSampleAudioVolume_Timer
'' Author    :
'' Date      : 30/07/2023
'' Purpose   : regular check to reflect changes to system volume
''---------------------------------------------------------------------------------------
''
'Public Sub tmrSampleAudioVolume_Timer()
'
'    On Error GoTo tmrSampleAudioVolume_Timer_Error
'
'    If mSystemAudioLevel = SystemVolume.GetMasterVolume Then Exit Sub
'
'    mSystemAudioLevel = SystemVolume.GetMasterVolume
'
'    sldX = mSystemAudioLevel * (sldXMax - sldXMin) + sldXMin
'    overlayWidget.sldCurWidth = sldXMax - sldX 'inform the Overlay about the new current "Cable-Width"
'    SliderSet.Move sldX, sldFixedTop 'move the SliderSet-widget accordingly
'
'    fClock.clockForm.Refresh
'
'   On Error GoTo 0
'   Exit Sub
'
'tmrSampleAudioVolume_Timer_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrSampleAudioVolume_Timer of Class Module cwOverlayTemp"
'End Sub

    
''---------------------------------------------------------------------------------------
'' Procedure : tmrSampleAudioMute_Timer
'' Author    :
'' Date      : 30/07/2023
'' Purpose   : regular check to reflect changes to system volume
''---------------------------------------------------------------------------------------
''
'Public Sub tmrSampleAudioMute_Timer()
'
'    On Error GoTo tmrSampleAudioMute_Timer_Error
'
'    If mMute = SystemVolume.GetMute Then Exit Sub
'
'    mMute = SystemVolume.GetMute
'
'    Call setLamps
'
'   On Error GoTo 0
'   Exit Sub
'
'tmrSampleAudioMute_Timer_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrSampleAudioMute_Timer of Class Module cwOverlayTemp"
'End Sub


' ----------------------------------------------------------------
' Procedure Name: setLamps
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 28/05/2024
' ----------------------------------------------------------------
Private Sub setLamps()
    
    On Error GoTo setLamps_Error
'    If mMute = 1 Then
'        fClock.clockForm.Widgets("indicatorgreen").Widget.Alpha = 0
'        fClock.clockForm.Widgets("indicatorred").Widget.Alpha = Val(gblOpacity) / 100
'    Else
'        fClock.clockForm.Widgets("indicatorgreen").Widget.Alpha = Val(gblOpacity) / 100
'        fClock.clockForm.Widgets("indicatorred").Widget.Alpha = 0
'    End If

    fClock.clockForm.Refresh
    
    On Error GoTo 0
    Exit Sub

setLamps_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setLamps, line " & Erl & "."

End Sub
