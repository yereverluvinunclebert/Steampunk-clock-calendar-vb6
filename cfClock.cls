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
Option Explicit

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" _
 (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, _
  ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Public WithEvents clockForm As cWidgetForm
Attribute clockForm.VB_VarHelpID = -1

Public collPSDNonUIElements As cCollection, FX As Long, FY As Long, FZ As Single  'available on the outside, even after the Form was closed

Private PSDWidth As Long
Private PSDHeight As Long

' properties
Private mMute As Integer
Private mSliderFreed As Boolean
Private msetToggleEnabled As Boolean
Private mMuteToggleEnabled As Boolean
Private mPendulumToggleEnabled As Boolean
Private mWeekdayToggleEnabled As Boolean
Private mDisplayScreenToggleEnabled As Boolean
Private mTimeMachineToggleEnabled As Boolean
Private mBackToggleEnabled As Boolean
Private mClapperEnabled As Boolean
Private mChainEnabled As Boolean
Private mCrankEnabled As Boolean
Private mTicking As Boolean
Private mPendulumEnabled As Boolean
Private mAlarmToggle1Enabled As Boolean
Private mAlarmToggle2Enabled As Boolean
Private mAlarmToggle3Enabled As Boolean
Private mAlarmToggle4Enabled As Boolean
Private mAlarmToggle5Enabled As Boolean

Public mShowHelp As Boolean

' widgets
Private WithEvents helpbottom As cWidgetBase
Attribute helpbottom.VB_VarHelpID = -1
Private WithEvents sliderset As cWidgetBase
Attribute sliderset.VB_VarHelpID = -1
Private WithEvents screentop As cWidgetBase
Attribute screentop.VB_VarHelpID = -1
Private WithEvents displayscreen As cWidgetBase
Attribute displayscreen.VB_VarHelpID = -1
Private WithEvents screenbackground As cWidgetBase
Attribute screenbackground.VB_VarHelpID = -1
Private WithEvents displayscreentoggle As cWidgetBase
Attribute displayscreentoggle.VB_VarHelpID = -1
Private WithEvents weekdaytoggle As cWidgetBase
Attribute weekdaytoggle.VB_VarHelpID = -1
Private WithEvents helptoggle As cWidgetBase
Attribute helptoggle.VB_VarHelpID = -1
Private WithEvents pendulumtoggle As cWidgetBase
Attribute pendulumtoggle.VB_VarHelpID = -1
Private WithEvents mutetoggle As cWidgetBase
Attribute mutetoggle.VB_VarHelpID = -1
Private WithEvents alarmtoggle As cWidgetBase
Attribute alarmtoggle.VB_VarHelpID = -1
Private WithEvents settoggle As cWidgetBase
Attribute settoggle.VB_VarHelpID = -1
Private WithEvents timemachinetoggle As cWidgetBase
Attribute timemachinetoggle.VB_VarHelpID = -1
Private WithEvents backtoggle As cWidgetBase
Attribute backtoggle.VB_VarHelpID = -1
Private WithEvents alarmtoggle1 As cWidgetBase
Attribute alarmtoggle1.VB_VarHelpID = -1
Private WithEvents alarmtoggle2 As cWidgetBase
Attribute alarmtoggle2.VB_VarHelpID = -1
Private WithEvents alarmtoggle3 As cWidgetBase
Attribute alarmtoggle3.VB_VarHelpID = -1
Private WithEvents alarmtoggle4 As cWidgetBase
Attribute alarmtoggle4.VB_VarHelpID = -1
Private WithEvents alarmtoggle5 As cWidgetBase
Attribute alarmtoggle5.VB_VarHelpID = -1
Private WithEvents clapperLeft As cWidgetBase
Attribute clapperLeft.VB_VarHelpID = -1
Private WithEvents clapperright As cWidgetBase
Attribute clapperright.VB_VarHelpID = -1
Private WithEvents labellayer As cWidgetBase
Attribute labellayer.VB_VarHelpID = -1
Private WithEvents chain As cWidgetBase
Attribute chain.VB_VarHelpID = -1
Private WithEvents crankup As cWidgetBase
Attribute crankup.VB_VarHelpID = -1
Private WithEvents crankdown As cWidgetBase
Attribute crankdown.VB_VarHelpID = -1
Private WithEvents weekdayredcover As cWidgetBase
Attribute weekdayredcover.VB_VarHelpID = -1
Private WithEvents dropdown As cWidgetBase
Attribute dropdown.VB_VarHelpID = -1
Private WithEvents bellset As cWidgetBase
Attribute bellset.VB_VarHelpID = -1
Private WithEvents helpdropdown As cWidgetBase
Attribute helpdropdown.VB_VarHelpID = -1
Private WithEvents glow As cWidgetBase
Attribute glow.VB_VarHelpID = -1

'general
Private sldFixedTop As Double
Private sldXMax As Double
Private sldXMin As Double
Private sldX As Double
Private mTimeShiftPercent As Currency
Private helpMoved As Boolean
Private muteInitiatingControl As String
Private pendulumInitiatingControl As String
Private alarmInitiatingControl As String
Private slidersetMousedown As Boolean

Private WithEvents tmrZZZZ As cTimer
Attribute tmrZZZZ.VB_VarHelpID = -1
Private WithEvents tmrTicking As cTimer
Attribute tmrTicking.VB_VarHelpID = -1


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
    
    Set tmrZZZZ = Nothing
    Set tmrZZZZ = New_c.Timer(300, False)
    
    Set tmrTicking = Nothing
    Set tmrTicking = New_c.Timer(300, False)


   On Error GoTo 0
   Exit Sub

Class_Initialize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cfClock"
End Sub

'---------------------------------------------------------------------------------------
' Property : ShowHelp
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get ShowHelp() As Boolean
   On Error GoTo ShowHelpGet_Error

   ShowHelp = mShowHelp

   On Error GoTo 0
   Exit Property

ShowHelpGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property ShowHelp of Class Module cfClock"
End Property
'---------------------------------------------------------------------------------------
' Property  : ShowHelp
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let ShowHelp(ByVal newValue As Boolean)

    Static helpbottomTop As Long
    Static helpbottomLeft As Long
    Static labellayerTop As Long
    Static labellayerLeft As Long
    Dim fileToPlay As String: fileToPlay = vbNullString

    On Error GoTo ShowHelpLet_Error

    
    helpbottomTop = helpbottom.Top ' 36
    helpbottomLeft = helpbottom.Left '220
    
    labellayerTop = labellayer.Top ' 10
    labellayerLeft = labellayer.Left '2
        
    If mShowHelp = True Then

         If helpMoved = True Then  ' now copes with f5 reloads
             helpbottom.Move 220, helpbottomTop
             labellayer.Move 2, labellayerTop
             helpMoved = False
         End If
         
         helpbottom.Alpha = Val(gblOpacity) / 100
         labellayer.Alpha = Val(gblOpacity) / 100
         
    Else

        ' the next two lines are a temporary hack.
        ' when widgets are made lapha = 0 they are still clickable, I can move these to the overlay or I
        ' can just move them out of the way so that any click on an invisible element does not interfere with one 'underneath'
        
         helpbottom.Alpha = 0
         labellayer.Alpha = 0

        If helpMoved = False Then
            helpbottom.Move helpbottomLeft + 700, helpbottomTop
            labellayer.Move labellayerLeft + 400, labellayerTop ' the 400 is a magical number that moves them just out of the way without causing tearing
            helpMoved = True
        End If
    
    End If
    
    If mShowHelp <> newValue Then mShowHelp = newValue Else Exit Property
    ShowHelp = mShowHelp
    
    clockForm.Refresh
   
    fileToPlay = "pagefumble.wav"
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If

   On Error GoTo 0
   Exit Property

ShowHelpLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property ShowHelp of Class Module cfClock"
End Property

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
        overlayWidget.cxPerc = 0.645 'define the relative center (within cwOverlay, which has the same size as the Form and the PSD...) this is used for clock-hand-positioning
        overlayWidget.cyPerc = 0.334
        
    'the next 5 lines will init a "moving Slider" (which was defined as "sliderset"-layer in the PSD)
    Set sliderset = clockForm.Widgets("sliderset").Widget 'store the already (in PSD-Loop created) Widget-reference in an explicit "WithEvents-Variable"
        
    With sliderset
        .Moveable = False 'make it not moveable by default
        .HoverColor = 0
        .MousePointer = IDC_HAND
        .Tag = 0.01
    End With
    
    sldFixedTop = sliderset.Top   'and store the initial position-data for y-direction (which will not be changed)
    sldXMax = sliderset.Left      'as well as the x-direction related pos (which is initially at the "100%-Volume, right-endswitch" )
    sldXMin = sldXMax - 116       'the Min-Value is simply the Max-Value minus Slider-Range (the 300pxl-range was determined by "probing")

    Set helpbottom = clockForm.Widgets("helpbottom").Widget
    Set screentop = clockForm.Widgets("screentop").Widget
    Set displayscreen = clockForm.Widgets("displayscreen").Widget
    Set displayscreentoggle = clockForm.Widgets("displayscreentoggle").Widget
    Set screenbackground = clockForm.Widgets("screenbackground").Widget
    Set weekdaytoggle = clockForm.Widgets("weekdaytoggle").Widget
    Set helptoggle = clockForm.Widgets("helptoggle").Widget
    Set pendulumtoggle = clockForm.Widgets("pendulumtoggle").Widget
    Set mutetoggle = clockForm.Widgets("mutetoggle").Widget
    Set alarmtoggle = clockForm.Widgets("alarmtoggle").Widget
    Set settoggle = clockForm.Widgets("settoggle").Widget
    Set timemachinetoggle = clockForm.Widgets("timemachinetoggle").Widget
    Set backtoggle = clockForm.Widgets("backtoggle").Widget
    Set alarmtoggle1 = clockForm.Widgets("alarmtoggle1").Widget
    Set alarmtoggle2 = clockForm.Widgets("alarmtoggle2").Widget
    Set alarmtoggle3 = clockForm.Widgets("alarmtoggle3").Widget
    Set alarmtoggle4 = clockForm.Widgets("alarmtoggle4").Widget
    Set alarmtoggle5 = clockForm.Widgets("alarmtoggle5").Widget
    Set clapperLeft = clockForm.Widgets("clapperleft").Widget
    Set clapperright = clockForm.Widgets("clapperright").Widget
    Set labellayer = clockForm.Widgets("labellayer").Widget
    Set chain = clockForm.Widgets("chain").Widget
    Set crankup = clockForm.Widgets("crankup").Widget
    Set crankdown = clockForm.Widgets("crankdown").Widget
    Set weekdayredcover = clockForm.Widgets("weekdayredcover").Widget
    Set dropdown = clockForm.Widgets("dropdown").Widget
    Set bellset = clockForm.Widgets("bellset").Widget
    Set helpdropdown = clockForm.Widgets("helpdropdown").Widget
    Set glow = clockForm.Widgets("glow").Widget
    
    
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
Private Sub clockForm_BubblingEvent(ByRef sender As Object, ByRef EventName As String, ByRef P1 As Variant, ByRef P2 As Variant, ByRef P3 As Variant, ByRef P4 As Variant, ByRef P5 As Variant, ByRef P6 As Variant, ByRef P7 As Variant)
   
    Dim stdMessage As String: stdMessage = vbNullString
    Dim stdSubject As String: stdSubject = vbNullString
    Dim fileToPlay As String: fileToPlay = vbNullString
    
    On Error GoTo clockForm_BubblingEvent_Error

    If gblIgnoreMouse = "1" Then Exit Sub
    
    stdMessage = "Right Click to open the menu and the preferences. CTRL+ Mouse scrollwheel UP/DOWN to resize. "
    stdSubject = "You can turn off the balloon tooltips in the preferences."
    
    If EventName = "W_MouseMove" Then
        Select Case LCase$(sender.Widget.Key)
            Case "sliderset"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "This is the Time Slider. Click and drag me left/right to alter time.", _
                     TTIconInfo, widgetName & " Locking pin Help", , , , True
            Case "helpbottom"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "This is the Widget Help. Click on me to make me go away.", _
                     TTIconInfo, widgetName & " Help", , , , True
            Case "timemachinetoggle"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "Toggle the time machine images.", _
                     TTIconInfo, widgetName & " Time Toggle Help", , , , True
            Case "displayscreentoggle"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "Toggle back screen.", _
                     TTIconInfo, widgetName & " Back Screen Toggle Help", , , , True
            Case "weekdaytoggle"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "Toggle weekday indicator.", _
                     TTIconInfo, widgetName & " Weekday Toggle Help", , , , True
            Case "helptoggle"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "Click me to show the help.", _
                     TTIconInfo, widgetName & " Help Toggle Help", , , , True
            Case "alarmtoggle"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "Click here to edit an alarm date and time.", _
                     TTIconInfo, widgetName & " Alarm Toggle Help", , , , True
            Case "settoggle"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "When you are editing an alarm, press this toggle to confirm any changes.", _
                     TTIconInfo, widgetName & " Loud Toggle Help", , , , True
            Case "mutetoggle"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "Mute chimes and sounds on/off.", _
                     TTIconInfo, widgetName & " Mute Toggle Help", , , , True
            Case "pendulumtoggle"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "Stop/start the pendulum.", _
                     TTIconInfo, widgetName & " Pendulum Toggle Help", , , , True
            Case "pendulum"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "Stop/start the pendulum.", _
                     TTIconInfo, widgetName & " Pendulum Help", , , , True
            Case "displayscreen"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "To lower kinematoscope press S toggle.", _
                     TTIconInfo, widgetName & " Kinematoscope Toggle Help", , , , True
            Case "bellset"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "Click here to set or edit an alarm.", _
                     TTIconInfo, widgetName & " Alarm Toggle Help", , , , True
            Case "clapperleft"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "Mute the chimes.", _
                     TTIconInfo, widgetName & " Chime Mute Toggle Help", , , , True
            Case "clapperright"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "Enable the chimes.", _
                     TTIconInfo, widgetName & " Chime Enable Toggle Help", , , , True
            Case "dropdown"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "Click me for information.", _
                     TTIconInfo, widgetName & " Dropdown Toggle Help", , , , True
            Case "backtoggle"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "Toggle the Back Screen to make the text easier to read.", _
                     TTIconInfo, widgetName & " Back Screen Toggle Help", , , , True
            Case "alarmtoggle1"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "Click to set alarm number one.", _
                     TTIconInfo, widgetName & " Alarm 1 Toggle Help", , , , True
            Case "alarmtoggle2"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "Click to set alarm number two.", _
                     TTIconInfo, widgetName & " Alarm 2 Toggle Help", , , , True
            Case "alarmtoggle3"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "Click to set alarm number three.", _
                     TTIconInfo, widgetName & " Alarm 3 Toggle Help", , , , True
            Case "alarmtoggle4"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "Click to set alarm number four.", _
                     TTIconInfo, widgetName & " Alarm 4 Toggle Help", , , , True
            Case "alarmtoggle5"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "Click to set alarm number five.", _
                     TTIconInfo, widgetName & " Alarm 5 Toggle Help", , , , True
            Case "chain"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "Pulling the chain will silence the ticks.", _
                     TTIconInfo, widgetName & " Chain Help", , , , True
            Case "crankup"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "Crank me down to quieten the whole clock.", _
                     TTIconInfo, widgetName & " Crank Help", , , , True
            Case "crankdown"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "Crank up the sound! Crank me up to make more sound throughout the whole clock.", _
                     TTIconInfo, widgetName & " Crank Help", , , , True
            Case "lockingpinlocked"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "Click here to unlock the clock and make it moveable using mouse.", _
                     TTIconInfo, widgetName & " Locking Pin Help", , , , True
            Case "lockingpin"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "Click here to lock the clock in place on the desktop.", _
                     TTIconInfo, widgetName & " Locking Pin Help", , , , True
             Case "helpdropdown"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "Click here to lock the clock in place on the desktop.", _
                     TTIconInfo, widgetName & " Dropdown Pin Help", , , , True
            Case "monday"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "Click here to lower the weekday flag.", _
                     TTIconInfo, widgetName & " Weekday Indicator Help", , , , True
                
                If clockForm.Widgets(fDayOfWeek).Widget.Alpha <> 0 Then weekdayredcover.Alpha = Val(gblOpacity) / 100
           
           Case "weekdayredcover"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip clockForm.hwnd, "Click here to lower the weekday flag.", _
                     TTIconInfo, widgetName & " Weekday Indicator Help", , , , True
        End Select
    End If
    
    If EventName = "W_MouseLeave" Then
        Select Case LCase$(sender.Widget.Key)
        Case "monday"
            weekdayredcover.Alpha = 0
        Case "tuesday"
            weekdayredcover.Alpha = 0
        Case "wednesday"
            weekdayredcover.Alpha = 0
        Case "thursday"
            weekdayredcover.Alpha = 0
        Case "friday"
            weekdayredcover.Alpha = 0
        Case "satday"
            weekdayredcover.Alpha = 0
        Case "sunday"
            weekdayredcover.Alpha = 0
        End Select
        
        weekdayredcover.Refresh
        
    End If
    
    If EventName = "W_MouseDown" Then
        Select Case LCase$(sender.Widget.Key)
            Case "helpdropdown"
                Call helpdropdown_myMouseDown
            Case "lockingpin"
                If overlayWidget.Locked = True Then
                    Call unlockit
                Else
                    Call lockit
                End If
             Case "grommet"
                If overlayWidget.Locked = True Then
                    Call unlockit
                Else
                    Call lockit
                End If
            Case "alarmtoggle1"
                Call alarmtoggle1_myMouseDown
            Case "alarmtoggle2"
                Call alarmtoggle2_myMouseDown
            Case "alarmtoggle3"
                Call alarmtoggle3_myMouseDown
            Case "alarmtoggle4"
                Call alarmtoggle4_myMouseDown
            Case "alarmtoggle5"
                Call alarmtoggle5_myMouseDown
            Case "alarmtoggle"
                Call alarmtoggle_myMouseDown
            Case "backtoggle"
                Call backtoggle_myMouseDown
            Case "timemachinetoggle"
                Call timeMachineToggle_myMouseDown
            Case "displayscreentoggle"
                Call displayScreenToggle_myMouseDown
            Case "pendulumtoggle"
                Call pendulumtoggle_myMouseDown
            Case "pendulum"
                Call pendulum_myMouseDown
            Case "mutetoggle"
                Call mutetoggle_myMouseDown
            Case "settoggle"
                Call settoggle_myMouseDown
            Case "weekdaytoggle"
                Call weekdayToggle_myMouseDown
            Case "helptoggle"
                Call helptoggle_myMouseDown
            Case "helpbottom"
                Call helpbottom_myMouseDown
            Case "clapperleft"
                Call clapperleft_myMouseDown
            Case "clapperright"
                Call clapperright_myMouseDown
            Case "chain"
                Call chain_myMouseDown
            Case "crankup"
                Call crankup_myMouseDown
            Case "crankdown"
                Call crankdown_myMouseDown
            Case "monday"
                Call weekdayredcover_myMouseDown
            Case "tuesday"
                Call weekdayredcover_myMouseDown
            Case "wednesday"
                Call weekdayredcover_myMouseDown
            Case "thursday"
                Call weekdayredcover_myMouseDown
            Case "friday"
                Call weekdayredcover_myMouseDown
            Case "saturday"
                Call weekdayredcover_myMouseDown
            Case "sunday"
                Call weekdayredcover_myMouseDown
            Case "dropdown"
                Call dropdown_myMouseDown
            Case "bellset"
                Call bellset_myMouseDown
           Case "sliderset"
                Call sliderset_myMouseDown

            Case Else
                Debug.Print sender.Widget.Key
        End Select
    End If
    
    If EventName = "W_MouseUp" Then
        Select Case LCase$(sender.Widget.Key)
            Case "bell"
                'Call bell_mouseUp(Sender)
            Case "sliderset"
    
                fileToPlay = "zzzz.wav"
                If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
                    PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
                End If
                
                overlayWidget.Locked = False
                If SliderFreed = True Then fClock.timeShiftPercent = 0.5
                tmrZZZZ.Enabled = False
            Case "sliderset"
                slidersetMousedown = False ' never called!
                
            Case Else
                Debug.Print sender.Widget.Key
        End Select
    End If
            
   On Error GoTo 0
   Exit Sub

clockForm_BubblingEvent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure clockForm_BubblingEvent of Class Module cfClock"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : sliderset_myMouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :  - note: - can't use a std MouseDown as
'             declaring a widget as cWidgetBase 'WithEvents' does not provide a std MouseDown with a sender object
'---------------------------------------------------------------------------------------
'
Private Sub sliderset_myMouseDown()
    Dim fileToPlay As String: fileToPlay = vbNullString
   
    On Error GoTo sliderset_myMouseDown_Error
    
    overlayWidget.Locked = True

    slidersetMousedown = True
                
    If SliderFreed = False Then
        fileToPlay = "buzzer.wav"
        If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
            PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
        End If
    End If
     
   On Error GoTo 0
   Exit Sub

sliderset_myMouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliderset_myMouseDown of Class Module cfClock"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : pendulum_myMouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :  - note: - can't use a std MouseDown as
'             declaring a widget as cWidgetBase 'WithEvents' does not provide a std MouseDown with a sender object
'---------------------------------------------------------------------------------------
'
Private Sub pendulum_myMouseDown()
    
    On Error GoTo pendulum_myMouseDown_Error
    
    pendulumInitiatingControl = "pendulum"

    If pendulumEnabled = True Then
        pendulumEnabled = False
    Else
        pendulumEnabled = True
    End If
    
   On Error GoTo 0
   Exit Sub

pendulum_myMouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure pendulum_myMouseDown of Class Module cfClock"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : bellset_myMouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :  - note: - can't use a std MouseDown as
'             declaring a widget as cWidgetBase 'WithEvents' does not provide a std MouseDown with a sender object
'---------------------------------------------------------------------------------------
'
Private Sub bellset_myMouseDown()
    Dim fileToPlay As String: fileToPlay = vbNullString
    
    On Error GoTo bellset_myMouseDown_Error
    
    fileToPlay = "singleBell.wav"
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If

   On Error GoTo 0
   Exit Sub

bellset_myMouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure bellset_myMouseDown of Class Module cfClock"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : dropdown_MouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :  - note: - can't use a std MouseDown as
'             declaring a widget as cWidgetBase 'WithEvents' does not provide a std MouseDown with a sender object
'---------------------------------------------------------------------------------------
'
Private Sub dropdown_myMouseDown()

    Dim fileToPlay As String: fileToPlay = vbNullString
    Dim alphaVal As Integer: alphaVal = 0
    Dim thisDayOfWeek As String: thisDayOfWeek = vbNullString
    
    Dim dropdownTop As Long: dropdownTop = 0
    Dim dropdownLeft As Long: dropdownLeft = 0
    
    Static dropdownEnabled As Boolean ' not using a property for this as it is not a stored value for restart
    
    On Error GoTo dropdown_MouseDown_Error
    
    dropdownTop = dropdown.Top
    dropdownLeft = dropdown.Left
            
    alphaVal = Val(gblOpacity) / 100
    
    If dropdownEnabled = False Then
        fileToPlay = "rollerblind.wav"
        dropdown.Move dropdownLeft, dropdownTop + 175
        dropdownEnabled = True
    Else
        fileToPlay = "rollerblindup.wav"
        dropdown.Move dropdownLeft, dropdownTop - 175
        dropdownEnabled = False
    End If
    
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    clockForm.Refresh

   On Error GoTo 0
   Exit Sub

dropdown_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure dropdown_MouseDown of Class Module cfClock"

End Sub

'

'---------------------------------------------------------------------------------------
' Procedure : alarmtoggle1_myMouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :  - note: - can't use a std MouseDown as
'             declaring a widget as cWidgetBase 'WithEvents' does not provide a std MouseDown with a sender object
'---------------------------------------------------------------------------------------
'
Private Sub alarmtoggle1_myMouseDown()

    Dim fileToPlay As String: fileToPlay = vbNullString

    On Error GoTo alarmtoggle1_myMouseDown_Error
    
    fileToPlay = "clunk.wav"
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    If alarmtoggle1Enabled = False Then
        alarmtoggle1Enabled = True
    Else
        alarmtoggle1Enabled = False
    End If
    
    clockForm.Refresh

   On Error GoTo 0
   Exit Sub

alarmtoggle1_myMouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure alarmtoggle1_myMouseDown of Class Module cfClock"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : alarmtoggle2_myMouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :  - note: - can't use a std MouseDown as
'             declaring a widget as cWidgetBase 'WithEvents' does not provide a std MouseDown with a sender object
'---------------------------------------------------------------------------------------
'
Private Sub alarmtoggle2_myMouseDown()

    Dim fileToPlay As String: fileToPlay = vbNullString

    On Error GoTo alarmtoggle2_myMouseDown_Error
    
    fileToPlay = "clunk.wav"
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
        
    If alarmtoggle2Enabled = False Then
        alarmtoggle2Enabled = True
    Else
        alarmtoggle2Enabled = False
    End If
    
    clockForm.Refresh

   On Error GoTo 0
   Exit Sub

alarmtoggle2_myMouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure alarmtoggle2_myMouseDown of Class Module cfClock"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : alarmtoggle3_myMouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :  - note: - can't use a std MouseDown as
'             declaring a widget as cWidgetBase 'WithEvents' does not provide a std MouseDown with a sender object
'---------------------------------------------------------------------------------------
'
Private Sub alarmtoggle3_myMouseDown()

    Dim fileToPlay As String: fileToPlay = vbNullString

    On Error GoTo alarmtoggle3_myMouseDown_Error
    
    fileToPlay = "clunk.wav"
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
        
    If alarmtoggle3Enabled = False Then
        alarmtoggle3Enabled = True
    Else
        alarmtoggle3Enabled = False
    End If
    
    clockForm.Refresh

   On Error GoTo 0
   Exit Sub

alarmtoggle3_myMouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure alarmtoggle3_myMouseDown of Class Module cfClock"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : alarmtoggle4_myMouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :  - note: - can't use a std MouseDown as
'             declaring a widget as cWidgetBase 'WithEvents' does not provide a std MouseDown with a sender object
'---------------------------------------------------------------------------------------
'
Private Sub alarmtoggle4_myMouseDown()

    Dim fileToPlay As String: fileToPlay = vbNullString

    On Error GoTo alarmtoggle4_myMouseDown_Error
    
    fileToPlay = "clunk.wav"
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    If alarmtoggle4Enabled = False Then
        alarmtoggle4Enabled = True
    Else
        alarmtoggle4Enabled = False
    End If
    
    clockForm.Refresh

   On Error GoTo 0
   Exit Sub

alarmtoggle4_myMouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure alarmtoggle4_myMouseDown of Class Module cfClock"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : alarmtoggle5_myMouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :  - note: - can't use a std MouseDown as
'             declaring a widget as cWidgetBase 'WithEvents' does not provide a std MouseDown with a sender object
'---------------------------------------------------------------------------------------
'
Private Sub alarmtoggle5_myMouseDown()

    Dim fileToPlay As String: fileToPlay = vbNullString

    On Error GoTo alarmtoggle5_myMouseDown_Error
    
    fileToPlay = "clunk.wav"
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    If alarmtoggle5Enabled = False Then
        alarmtoggle5Enabled = True
    Else
        alarmtoggle5Enabled = False
    End If
    
    clockForm.Refresh

   On Error GoTo 0
   Exit Sub

alarmtoggle5_myMouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure alarmtoggle5_myMouseDown of Class Module cfClock"

End Sub





'---------------------------------------------------------------------------------------
' Procedure : weekdayredcover_myMouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :  - note: - can't use a std MouseDown as
'             declaring a widget as cWidgetBase 'WithEvents' does not provide a std MouseDown with a sender object
'---------------------------------------------------------------------------------------
'
Private Sub weekdayredcover_myMouseDown()

    Dim fileToPlay As String: fileToPlay = vbNullString
    Dim alphaVal As Integer: alphaVal = 0
    Dim thisDayOfWeek As String: thisDayOfWeek = vbNullString
    
    Dim weekdayToggleTop As Long: weekdayToggleTop = 0
    Dim weekdayToggleLeft As Long: weekdayToggleLeft = 0
    
    On Error GoTo weekdayredcover_myMouseDown_Error
    
    weekdayToggleTop = weekdaytoggle.Top
    weekdayToggleLeft = weekdaytoggle.Left
    
    thisDayOfWeek = fDayOfWeek
        
    alphaVal = Val(gblOpacity) / 100
    
    fileToPlay = "till.wav"
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    clockForm.Widgets(thisDayOfWeek).Widget.Alpha = 0
    weekdayToggleEnabled = False

    clockForm.Refresh

   On Error GoTo 0
   Exit Sub

weekdayredcover_myMouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure weekdayredcover_myMouseDown of Class Module cfClock"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : chain_myMouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :  - note: - can't use a std MouseDown as
'             declaring a widget as cWidgetBase 'WithEvents' does not provide a std MouseDown with a sender object
'---------------------------------------------------------------------------------------
'
Private Sub chain_myMouseDown()
    Dim lngReturn As Long: lngReturn = 0
    Dim fileToPlay As String: fileToPlay = vbNullString

    On Error GoTo chain_myMouseDown_Error
    
    fileToPlay = "clunk.wav"
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME
    End If

    If gblChainEnabled = "False" Then
        chainEnabled = True
    Else
        chainEnabled = False
    End If
    
    clockForm.Refresh

   On Error GoTo 0
   Exit Sub

chain_myMouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chain_myMouseDown of Class Module cfClock"

End Sub




'---------------------------------------------------------------------------------------
' Procedure : backtoggle_myMouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :  - note: - can't use a std MouseDown as
'             declaring a widget as cWidgetBase 'WithEvents' does not provide a std MouseDown with a sender object
'---------------------------------------------------------------------------------------
'
Private Sub backtoggle_myMouseDown()
   On Error GoTo backtoggle_myMouseDown_Error

    Dim fileToPlay As String: fileToPlay = vbNullString
    
    fileToPlay = "clunk.wav"
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
        
    If gblBackToggleEnabled = "False" Then
        backToggleEnabled = True
    Else
        backToggleEnabled = False
    End If
    
    clockForm.Refresh

   On Error GoTo 0
   Exit Sub

backtoggle_myMouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure backtoggle_myMouseDown of Class Module cfClock"

End Sub
    

    
'---------------------------------------------------------------------------------------
' Procedure : alarmtoggle_myMouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :  - note: - can't use a std MouseDown as
'             declaring a widget as cWidgetBase 'WithEvents' does not provide a std MouseDown with a sender object
'---------------------------------------------------------------------------------------
'
Private Sub alarmtoggle_myMouseDown()

    Dim fileToPlay As String: fileToPlay = vbNullString
    Dim alarmtoggleTop As Long: alarmtoggleTop = 0
    Dim alarmtoggleLeft As Long: alarmtoggleLeft = 0
    
    Static alarmtoggleEnabled As Boolean ' not using a property for this as it is not a stored value for restart
    
    Dim alarmToggleIn As Long: alarmToggleIn = 0
    Dim alarmToggleOut As Long: alarmToggleOut = 0
    
    On Error GoTo alarmtoggle_myMouseDown_Error
    
    alarmtoggleTop = alarmtoggle.Top
    alarmToggleOut = 359
    alarmToggleIn = 363
    
    slidersetMousedown = False

    fileToPlay = "clunk.wav"
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME
    End If
    
    If alarmtoggleEnabled = False Then
        alarmtoggle.Move alarmToggleIn, alarmtoggleTop
        alarmtoggleEnabled = True
        SliderFreed = True
        ' set the slider position to the start point
        fClock.timeShiftPercent = 0.5
        
    Else
        alarmtoggle.Move alarmToggleOut, alarmtoggleTop
        
        alarmtoggleEnabled = False
        SliderFreed = False
        fClock.timeShiftPercent = 0
        
        
    End If
    clockForm.Refresh

   On Error GoTo 0
   Exit Sub

alarmtoggle_myMouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure alarmtoggle_myMouseDown of Class Module cfClock"

End Sub
    

'---------------------------------------------------------------------------------------
' Procedure : timeMachineToggle_myMouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :  - note: - can't use a std MouseDown as
'             declaring a widget as cWidgetBase 'WithEvents' does not provide a std MouseDown with a sender object
'---------------------------------------------------------------------------------------
'
Private Sub timeMachineToggle_myMouseDown()

    Dim fileToPlay As String: fileToPlay = vbNullString

    On Error GoTo timeMachineToggle_myMouseDown_Error
    
    fileToPlay = "clunk.wav"
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    If gblTimeMachineToggleEnabled = "False" Then
        timeMachineToggleEnabled = True
    Else
        timeMachineToggleEnabled = False
    End If
        
    clockForm.Refresh

   On Error GoTo 0
   Exit Sub

timeMachineToggle_myMouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure timeMachineToggle_myMouseDown of Class Module cfClock"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : weekdayToggle_myMouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub weekdayToggle_myMouseDown()
   On Error GoTo weekdayToggle_myMouseDown_Error

    Dim fileToPlay As String: fileToPlay = vbNullString

    fileToPlay = "clunk.wav"
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME
    End If
    
    If gblWeekdayToggleEnabled = "False" Then
        weekdayToggleEnabled = True
    Else
        weekdayToggleEnabled = False
    End If
        
    clockForm.Refresh

   On Error GoTo 0
   Exit Sub

weekdayToggle_myMouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure weekdayToggle_myMouseDown of Class Module cfClock"

End Sub


    
'---------------------------------------------------------------------------------------
' Procedure : settoggle_myMouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub settoggle_myMouseDown()

    Dim settoggleTop As Long: settoggleTop = 0
    Dim settoggleLeft As Long: settoggleLeft = 0

    Dim fileToPlay As String: fileToPlay = vbNullString
    
    Static settoggleEnabled As Boolean ' not using a property for this as it is not a stored value for restart
    
    Dim setToggleIn As Long: setToggleIn = 0
    Dim setToggleOut As Long: setToggleOut = 0
    
    On Error GoTo settoggle_myMouseDown_Error
    
    settoggleTop = settoggle.Top
    setToggleOut = 359
    setToggleIn = 363

    fileToPlay = "clunk.wav"
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    If settoggleEnabled = False Then
        settoggle.Move setToggleIn, settoggleTop
        settoggleEnabled = True
    Else
        settoggle.Move setToggleOut, settoggleTop
        settoggleEnabled = False
    End If
    
    clockForm.Refresh


   On Error GoTo 0
   Exit Sub

settoggle_myMouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure settoggle_myMouseDown of Class Module cfClock"

End Sub
    
'---------------------------------------------------------------------------------------
' Procedure : mutetoggle_myMouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mutetoggle_myMouseDown()
    
    On Error GoTo mutetoggle_myMouseDown_Error
    
    muteInitiatingControl = "muteToggle"
     
    If gblMuteToggleEnabled = "False" Then
        muteToggleEnabled = True
    Else
        muteToggleEnabled = False
    End If
    clockForm.Refresh

   On Error GoTo 0
   Exit Sub

mutetoggle_myMouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mutetoggle_myMouseDown of Class Module cfClock"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : pendulumtoggle_myMouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub pendulumtoggle_myMouseDown()

    On Error GoTo pendulumtoggle_myMouseDown_Error

    pendulumInitiatingControl = "pendulumToggle"

    If gblPendulumToggleEnabled = "False" Then
        pendulumToggleEnabled = True
    Else
        pendulumToggleEnabled = False
    End If
    clockForm.Refresh

   On Error GoTo 0
   Exit Sub

pendulumtoggle_myMouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure pendulumtoggle_myMouseDown of Class Module cfClock"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : displayScreenToggle_myMouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub displayScreenToggle_myMouseDown()
   On Error GoTo displayScreenToggle_myMouseDown_Error
   
    Dim fileToPlay As String: fileToPlay = vbNullString
    Dim displayScreenToggleTop As Long: displayScreenToggleTop = 0
    Dim displayScreenToggleLeft As Long: displayScreenToggleLeft = 0
    Dim alphaVal As Integer: alphaVal = 0
    
    displayScreenToggleTop = displayscreentoggle.Top
    displayScreenToggleLeft = displayscreentoggle.Left
    alphaVal = Val(gblOpacity) / 100

    fileToPlay = "clunk.wav"
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
  
    If gblDisplayScreenToggleEnabled = "False" Then
        displayScreenToggleEnabled = True
    Else
        displayScreenToggleEnabled = False
    End If
    
    clockForm.Refresh
    
   On Error GoTo 0
   Exit Sub

displayScreenToggle_myMouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure displayScreenToggle_myMouseDown of Class Module cfClock"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : helptoggle_myMouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub helptoggle_myMouseDown()
   On Error GoTo helptoggle_myMouseDown_Error

    Dim fileToPlay As String: fileToPlay = vbNullString
    Dim helptoggleTop As Long: helptoggleTop = 0
    Dim helptoggleLeft As Long: helptoggleLeft = 0
    Dim alphaVal As Integer: alphaVal = 0

    helptoggleTop = helptoggle.Top
    helptoggleLeft = helptoggle.Left
    alphaVal = Val(gblOpacity) / 100
    
    If helpdropdown.Alpha = 0 Then
        fileToPlay = "rollerblinddown.wav"
        helptoggle.Alpha = 0
        clockForm.Widgets("woodenbar").Widget.Alpha = alphaVal
        clockForm.Widgets("help1toggle").Widget.Alpha = alphaVal
        clockForm.Widgets("helpdropdown").Widget.Alpha = alphaVal
        clockForm.Widgets("txthelp").Widget.Alpha = alphaVal
    Else
        fileToPlay = "rollerblindup.wav"
        helptoggle.Alpha = alphaVal
        clockForm.Widgets("woodenbar").Widget.Alpha = 0
        clockForm.Widgets("help1toggle").Widget.Alpha = 0
        clockForm.Widgets("helpdropdown").Widget.Alpha = 0
        clockForm.Widgets("txthelp").Widget.Alpha = 0
    End If
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    clockForm.Refresh

   On Error GoTo 0
   Exit Sub

helptoggle_myMouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure helptoggle_myMouseDown of Class Module cfClock"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : helptoggle_myMouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub helpdropdown_myMouseDown()
   On Error GoTo helpdropdown_myMouseDown_Error

    Dim fileToPlay As String: fileToPlay = vbNullString
    Dim helpdropdownTop As Long: helpdropdownTop = 0
    Dim helpdropdownLeft As Long: helpdropdownLeft = 0
    Dim alphaVal As Integer: alphaVal = 0

    If helpdropdown.Alpha <> 0 Then
        helpdropdownTop = helpdropdown.Top
        helpdropdownLeft = helpdropdown.Left
        alphaVal = Val(gblOpacity) / 100
    
        fileToPlay = "rollerblindup.wav"
        If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
            PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
        End If
        
        clockForm.Widgets("woodenbar").Widget.Alpha = 0
        clockForm.Widgets("help1toggle").Widget.Alpha = 0
        clockForm.Widgets("helpdropdown").Widget.Alpha = 0
        clockForm.Widgets("txthelp").Widget.Alpha = 0
        helptoggle.Alpha = alphaVal
        
        clockForm.Refresh
    End If

   On Error GoTo 0
   Exit Sub

helpdropdown_myMouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure helpdropdown_myMouseDown of Class Module cfClock"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : lockingpin_mouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lockit()
   On Error GoTo lockingpin_mouseDown_Error

    clockForm.Widgets("lockingpinlocked").Widget.Alpha = Val(gblOpacity) / 100
    clockForm.Widgets("lockingpin").Widget.Alpha = Val(gblOpacity) / 100

    clockForm.Widgets("lockingpinlocked").Widget.Visible = True
    clockForm.Widgets("lockingpin").Widget.Visible = False

    clockForm.Refresh
    Call toggleWidgetLock
    
   On Error GoTo 0
   Exit Sub

lockingpin_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lockingpin_mouseDown of Class Module cfClock"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : lockingpin_mouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub unlockit()
   On Error GoTo lockingpin_mouseDown_Error

    
    clockForm.Widgets("lockingpinlocked").Widget.Alpha = Val(gblOpacity) / 100
    clockForm.Widgets("lockingpin").Widget.Alpha = Val(gblOpacity) / 100

    clockForm.Widgets("lockingpinlocked").Widget.Visible = False
    clockForm.Widgets("lockingpin").Widget.Visible = True

    clockForm.Refresh
    Call toggleWidgetLock
    
   On Error GoTo 0
   Exit Sub

lockingpin_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lockingpin_mouseDown of Class Module cfClock"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : helpbottom_mouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   : make the elements that are displayed on startup seem to disappear (moved)
'---------------------------------------------------------------------------------------
'
Private Sub helpbottom_myMouseDown()
'    Dim helpbottomTop As Long: helpbottomTop = 0
'    Dim helpbottomLeft As Long: helpbottomLeft = 0
'    Dim labellayerTop As Long: labellayerTop = 0
'    Dim labellayerLeft As Long: labellayerLeft = 0
'    Dim alphaVal As Integer: alphaVal = 0

'    Dim fileToPlay As String: fileToPlay = vbNullString
    
    On Error GoTo helpbottom_mouseDown_Error
'
'    helpbottomTop = helpbottom.Top
'    helpbottomLeft = helpbottom.Left
'
'    labellayerTop = labellayer.Top
'    labellayerLeft = labellayer.Left

    If helpbottom.Alpha <> 0 Then ' cope with the click on the invisible helpbottom layer
        
        ShowHelp = False
                
        'If startupFlg = True Then
            clockForm.Widgets("monday").Widget.Alpha = 0
            If alarmtoggle1Enabled = False Then clockForm.Widgets("alarm1till").Widget.Alpha = 0
            clockForm.Widgets("redarlarmcover").Widget.Alpha = 0
        'End If
        
'        fileToPlay = "pagefumble.wav"
'        If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
'            PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
'        End If
                
        'clockForm.Refresh
    End If
        
   On Error GoTo 0
   Exit Sub

helpbottom_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure helpbottom_mouseDown of Class Module cfClock"

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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_MouseDown of Class Module cfClock"
End Sub

' ----------------------------------------------------------------
' Procedure Name: sliderset_Moving
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 26/05/2024
' ----------------------------------------------------------------
Private Sub sliderset_Moving() 'this is an explcit Event-Handler on the single Widget-instance (derived in the routine above)
   
    On Error GoTo sliderset_Moving_Error
         
     sldX = sliderset.Left
     If sldX < sldXMin Then sldX = sldXMin Else If sldX > sldXMax Then sldX = sldXMax 'keep sldX "in range"
     
     If tmrZZZZ.Enabled = False And slidersetMousedown = True Then tmrZZZZ.Enabled = True
    
     timeShiftPercent = (sldX - sldXMin) / (sldXMax - sldXMin) 'assign the new Value to the Public Prop, to trigger all the necessary refreshs
    
    On Error GoTo 0
    Exit Sub

sliderset_Moving_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliderset_Moving, line " & Erl & "."

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
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property SystemAudioLevel of Class Module cfClock"
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
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Mute of Class Module cfClock"
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
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Mute of Class Module cfClock"
'End Property

'---------------------------------------------------------------------------------------
' Property  : SliderFreed
' Author    : beededea
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let SliderFreed(ByVal newValue As Boolean)
   On Error GoTo SliderFreedLet_Error

    If mSliderFreed <> newValue Then mSliderFreed = newValue Else Exit Property

    SliderFreed = mSliderFreed
   
    If mSliderFreed = True Then
        sliderset.Moveable = True
        glow.Moveable = True
    Else
        sliderset.Moveable = False
        glow.Moveable = False
    End If
   
   On Error GoTo 0
   Exit Property

SliderFreedLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property SliderFreed of Class Module cfClock"
End Property

'---------------------------------------------------------------------------------------
' Property  : SliderFreed
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get SliderFreed() As Boolean
   On Error GoTo SliderFreedGet_Error

    SliderFreed = mSliderFreed

   On Error GoTo 0
   Exit Property

SliderFreedGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property SliderFreed of Class Module cfClock"
End Property


' ----------------------------------------------------------------
' Procedure Name: timeShiftPercent
' Purpose:
' Procedure Kind: Property (Get)
' Procedure Access: Public
' Return Type: Currency
' Author: beededea
' Date: 26/05/2024
' ----------------------------------------------------------------
Public Property Get timeShiftPercent() As Currency
    
    On Error GoTo timeShiftPercent_Error
    
    timeShiftPercent = mTimeShiftPercent
    
    On Error GoTo 0
    Exit Property

timeShiftPercent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure timeShiftPercent, line " & Erl & "."

End Property

' ----------------------------------------------------------------
' Procedure Name: timeShiftPercent
' Purpose:
' Procedure Kind: Property (Let)
' Procedure Access: Public
' Parameter RHS (Currency):
' Author: beededea
' Date: 26/05/2024
' ----------------------------------------------------------------
Public Property Let timeShiftPercent(ByVal RHS As Currency)
    On Error GoTo timeShiftPercent_Error
    
    If RHS < 0 Then RHS = 0 Else If RHS > 1 Then RHS = 1 'keep RHS "in range"
    sldX = RHS * (sldXMax - sldXMin) + sldXMin
    overlayWidget.sldCurWidth = sldXMax - sldX 'inform the Overlay about the new current "Cable-Width"
    
    sliderset.Move sldX, sldFixedTop 'move the sliderset-widget accordingly
    glow.Move sldX + 29, sldFixedTop - 13 'move the glow along with the sliderset-widget accordingly
    
    If mTimeShiftPercent <> RHS Then mTimeShiftPercent = RHS Else Exit Property  'only refresh, when there was a change to the previous value
    
    If Not Cairo.ToolTipForm Is Nothing Then Cairo.ToolTipForm.Visible = False
    clockForm.Refresh
    
'    Call HandleVolumeChangeForHardware
    
    On Error GoTo 0
    Exit Property

timeShiftPercent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure timeShiftPercent, line " & Erl & "."

End Property





'---------------------------------------------------------------------------------------
' Property  : displayScreenToggleEnabled
' Author    : beededea
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let displayScreenToggleEnabled(ByVal newValue As Boolean)
   
    Dim displayScreenToggleTop As Long: displayScreenToggleTop = 0
    Dim displayScreenToggleIn As Long: displayScreenToggleIn = 0
    Dim displayScreenToggleOut As Long: displayScreenToggleOut = 0
    Dim fileToPlay As String: fileToPlay = vbNullString
    Dim alphaVal As Integer: alphaVal = 0
    Dim thisDayOfWeek As String: thisDayOfWeek = vbNullString
    
    On Error GoTo displayScreenToggleEnabledLet_Error
    
    displayScreenToggleTop = displayscreentoggle.Top
    displayScreenToggleOut = 393
    displayScreenToggleIn = 397
    
    alphaVal = Val(gblOpacity) / 100

    If newValue = True Then
        displayscreentoggle.Move displayScreenToggleIn, displayScreenToggleTop
        displayscreen.Alpha = alphaVal
        screentop.Alpha = 0
        timemachinetoggle.Alpha = alphaVal
        backtoggle.Alpha = alphaVal
        
        If backToggleEnabled = True Then screenbackground.Alpha = alphaVal
        gblDisplayScreenToggleEnabled = "True"
    Else
        displayscreentoggle.Move displayScreenToggleOut, displayScreenToggleTop
        displayscreen.Alpha = 0
        screentop.Alpha = alphaVal
        timemachinetoggle.Alpha = 0
        backtoggle.Alpha = 0
        screenbackground.Alpha = 0
        gblDisplayScreenToggleEnabled = "False"
    End If
    
    sPutINISetting "Software\SteampunkClockCalendar", "displayScreenToggleEnabled", gblDisplayScreenToggleEnabled, gblSettingsFile

    If mDisplayScreenToggleEnabled <> newValue Then mDisplayScreenToggleEnabled = newValue Else Exit Property
    displayScreenToggleEnabled = mDisplayScreenToggleEnabled
    
    clockForm.Refresh
     
    fileToPlay = "clunk.wav"
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME
    End If

      
   On Error GoTo 0
   Exit Property

displayScreenToggleEnabledLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property displayScreenToggleEnabled of Class Module cfClock"
End Property


'---------------------------------------------------------------------------------------
' Property  : displayScreenToggleEnabled
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get displayScreenToggleEnabled() As Boolean
   On Error GoTo displayScreenToggleEnabledGet_Error

    displayScreenToggleEnabled = mDisplayScreenToggleEnabled

   On Error GoTo 0
   Exit Property

displayScreenToggleEnabledGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property displayScreenToggleEnabled of Class Module cfClock"
End Property

'---------------------------------------------------------------------------------------
' Procedure : clapperleft_myMouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :  - note: - can't use a std MouseDown as
'             declaring a widget as cWidgetBase 'WithEvents' does not provide a std MouseDown with a sender object
'---------------------------------------------------------------------------------------
'
Private Sub clapperleft_myMouseDown()
    'Dim fileToPlay As String: fileToPlay = vbNullString
    
    On Error GoTo clapperleft_myMouseDown_Error

    muteInitiatingControl = "clapper"

'    fileToPlay = "clunk.wav"
'
'    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
'        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
'    End If
        
    If clapperEnabled = True Then
        clapperEnabled = False
    Else
        clapperEnabled = True
    End If
    
    clockForm.Refresh

   On Error GoTo 0
   Exit Sub

clapperleft_myMouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure clapperleft_myMouseDown of Class Module cfClock"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : clapperright_myMouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :  - note: - can't use a std MouseDown as
'             declaring a widget as cWidgetBase 'WithEvents' does not provide a std MouseDown with a sender object
'---------------------------------------------------------------------------------------
'
Private Sub clapperright_myMouseDown()

    On Error GoTo clapperright_myMouseDown_Error
    
    muteInitiatingControl = "clapper"

    If clapperEnabled = True Then
        clapperEnabled = False
    Else
        clapperEnabled = True
    End If
    
    clockForm.Refresh

   On Error GoTo 0
   Exit Sub

clapperright_myMouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure clapperright_myMouseDown of Class Module cfClock"

End Sub

'---------------------------------------------------------------------------------------
' Property  : clapperEnabled
' Author    : beededea
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let clapperEnabled(ByVal newValue As Boolean)
    Dim fileToPlay As String: fileToPlay = vbNullString
    Dim alphaVal As Integer: alphaVal = 0

    On Error GoTo clapperEnabledLet_Error
    
    alphaVal = Val(gblOpacity) / 100
 
    If newValue = True Then
        gblClapperEnabled = "True"
        clapperright.Alpha = 0
        clapperLeft.Alpha = alphaVal
        If muteInitiatingControl = "clapper" Then
            fileToPlay = "singleBell.wav"
            If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
                PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
            End If
        End If
    Else
        gblClapperEnabled = "False"
        'If muteInitiatingControl = "clapper" Then muteToggleEnabled = True
        clapperright.Alpha = alphaVal
        clapperLeft.Alpha = 0

    End If
    
    sPutINISetting "Software\SteampunkClockCalendar", "clapperEnabled", gblClapperEnabled, gblSettingsFile

    If mClapperEnabled <> newValue Then mClapperEnabled = newValue Else Exit Property
    clapperEnabled = mClapperEnabled
    
    clockForm.Refresh
     
   On Error GoTo 0
   Exit Property

clapperEnabledLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property clapperEnabled of Class Module cfClock"
End Property

'---------------------------------------------------------------------------------------
' Property  : clapperEnabled
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get clapperEnabled() As Boolean
   On Error GoTo clapperEnabledGet_Error

    clapperEnabled = mClapperEnabled

   On Error GoTo 0
   Exit Property

clapperEnabledGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property clapperEnabled of Class Module cfClock"
End Property

'---------------------------------------------------------------------------------------
' Property  : backToggleEnabled
' Author    : beededea
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let backToggleEnabled(ByVal newValue As Boolean)
   
    Dim backToggleTop As Long: backToggleTop = 0
    Dim backToggleIn As Long: backToggleIn = 0
    Dim backToggleOut As Long: backToggleOut = 0
    Dim fileToPlay As String: fileToPlay = vbNullString
    Dim alphaVal As Integer: alphaVal = 0
    Dim thisDayOfWeek As String: thisDayOfWeek = vbNullString
    
    On Error GoTo backToggleEnabledLet_Error
        
    backToggleTop = backtoggle.Top
    backToggleOut = 654
    backToggleIn = 650
    
    alphaVal = Val(gblOpacity) / 100

    If newValue = True Then
        backtoggle.Move backToggleIn, backToggleTop
        gblBackToggleEnabled = "True"
        clockForm.Widgets("screenbackground").Widget.Alpha = alphaVal
    Else
        backtoggle.Move backToggleOut, backToggleTop
        gblBackToggleEnabled = "False"
        clockForm.Widgets("screenbackground").Widget.Alpha = 0
    End If
    
    sPutINISetting "Software\SteampunkClockCalendar", "backToggleEnabled", gblBackToggleEnabled, gblSettingsFile

    If mBackToggleEnabled <> newValue Then mBackToggleEnabled = newValue Else Exit Property
    backToggleEnabled = mBackToggleEnabled
    
    clockForm.Refresh
     
   On Error GoTo 0
   Exit Property

backToggleEnabledLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property backToggleEnabled of Class Module cfClock"
End Property

'---------------------------------------------------------------------------------------
' Property  : backToggleEnabled
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get backToggleEnabled() As Boolean
   On Error GoTo backToggleEnabledGet_Error

    backToggleEnabled = mBackToggleEnabled

   On Error GoTo 0
   Exit Property

backToggleEnabledGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property backToggleEnabled of Class Module cfClock"
End Property


' ---------------------------------------------------------------------------------------
' Procedure : crankdown_myMouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :  - note: - can't use a std MouseDown as
'             declaring a widget as cWidgetBase 'WithEvents' does not provide a std MouseDown with a sender object
'---------------------------------------------------------------------------------------
'
Private Sub crankdown_myMouseDown()

    Dim fileToPlay As String: fileToPlay = vbNullString
            
    On Error GoTo crankdown_myMouseDown_Error
    
    fileToPlay = "crank.wav"
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    crankEnabled = True

    clockForm.Refresh

   On Error GoTo 0
   Exit Sub

crankdown_myMouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure crankdown_myMouseDown of Class Module cfClock"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : crankup_myMouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :  - note: - can't use a std MouseDown as
'             declaring a widget as cWidgetBase 'WithEvents' does not provide a std MouseDown with a sender object
'---------------------------------------------------------------------------------------
'
Private Sub crankup_myMouseDown()

    Dim fileToPlay As String: fileToPlay = vbNullString
    
    On Error GoTo crankup_myMouseDown_Error
    
    fileToPlay = "crank.wav"
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    crankEnabled = False

    clockForm.Refresh

   On Error GoTo 0
   Exit Sub

crankup_myMouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure crankup_myMouseDown of Class Module cfClock"

End Sub




'---------------------------------------------------------------------------------------
' Property  : crankEnabled
' Author    : beededea
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let crankEnabled(ByVal newValue As Boolean)
    Dim alphaVal As Integer: alphaVal = 0
  
    On Error GoTo crankEnabledLet_Error
    
    alphaVal = Val(gblOpacity) / 100
    
    If newValue = True Then
        crankup.Alpha = alphaVal
        crankdown.Alpha = 0
        gblCrankEnabled = "True"
    Else
        gblCrankEnabled = "False"
        crankup.Alpha = 0
        crankdown.Alpha = alphaVal
    End If
       
    sPutINISetting "Software\SteampunkClockCalendar", "crankEnabled", gblCrankEnabled, gblSettingsFile

    If mCrankEnabled <> newValue Then mCrankEnabled = newValue Else Exit Property
    crankEnabled = mCrankEnabled
    
    clockForm.Refresh
     
   On Error GoTo 0
   Exit Property

crankEnabledLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property crankEnabled of Class Module cfClock"
End Property

'---------------------------------------------------------------------------------------
' Property  : crankEnabled
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get crankEnabled() As Boolean
   On Error GoTo crankEnabledGet_Error

    crankEnabled = mCrankEnabled

   On Error GoTo 0
   Exit Property

crankEnabledGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property crankEnabled of Class Module cfClock"
End Property




'---------------------------------------------------------------------------------------
' Property  : chainEnabled
' Author    : beededea
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let chainEnabled(ByVal newValue As Boolean)
   
    Dim chainTop As Long: chainTop = 0
    Dim chainLeft As Long: chainLeft = 0

    On Error GoTo chainEnabledLet_Error
    
    chainTop = chain.Top
    chainLeft = chain.Left
    
    If newValue = True Then
        chain.Move chainLeft, chainTop + 10
        gblChainEnabled = "True"
        ticking = True
    Else
        chain.Move chainLeft, chainTop - 10
        gblChainEnabled = "False"
        ticking = False
    End If
       
    sPutINISetting "Software\SteampunkClockCalendar", "chainEnabled", gblChainEnabled, gblSettingsFile

    If mChainEnabled <> newValue Then mChainEnabled = newValue Else Exit Property
    chainEnabled = mChainEnabled
    
    clockForm.Refresh
     
   On Error GoTo 0
   Exit Property

chainEnabledLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property chainEnabled of Class Module cfClock"
End Property

'---------------------------------------------------------------------------------------
' Property  : chainEnabled
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get chainEnabled() As Boolean
   On Error GoTo chainEnabledGet_Error

    chainEnabled = mChainEnabled

   On Error GoTo 0
   Exit Property

chainEnabledGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property chainEnabled of Class Module cfClock"
End Property
'---------------------------------------------------------------------------------------
' Property  : weekdayToggleEnabled
' Author    : beededea
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let weekdayToggleEnabled(ByVal newValue As Boolean)
   
    Dim weekdayToggleTop As Long: weekdayToggleTop = 0
    Dim weekdayToggleIn As Long: weekdayToggleIn = 0
    Dim weekdayToggleOut As Long: weekdayToggleOut = 0
    Dim fileToPlay As String: fileToPlay = vbNullString
    Dim alphaVal As Integer: alphaVal = 0
    Dim thisDayOfWeek As String: thisDayOfWeek = vbNullString
    
    On Error GoTo weekdayToggleEnabledLet_Error
        
    weekdayToggleTop = weekdaytoggle.Top
    weekdayToggleOut = 397
    weekdayToggleIn = 401
    
    weekdayToggleTop = weekdaytoggle.Top
    
    alphaVal = Val(gblOpacity) / 100
    thisDayOfWeek = fDayOfWeek

    If newValue = True Then
        weekdaytoggle.Move weekdayToggleIn, weekdayToggleTop
        clockForm.Widgets(thisDayOfWeek).Widget.Alpha = alphaVal
        gblWeekdayToggleEnabled = "True"
    Else
        weekdaytoggle.Move weekdayToggleOut, weekdayToggleTop
        clockForm.Widgets(thisDayOfWeek).Widget.Alpha = 0
        weekdayredcover.Alpha = 0 ' sometimes you can fool the mouseleave, this recovers the situation
        gblWeekdayToggleEnabled = "False"
    End If
    
    sPutINISetting "Software\SteampunkClockCalendar", "weekdayToggleEnabled", gblWeekdayToggleEnabled, gblSettingsFile

    If mWeekdayToggleEnabled <> newValue Then mWeekdayToggleEnabled = newValue Else Exit Property
    weekdayToggleEnabled = mWeekdayToggleEnabled
    
    clockForm.Refresh
     
    fileToPlay = "till.wav"
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME
    End If

      
   On Error GoTo 0
   Exit Property

weekdayToggleEnabledLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property weekdayToggleEnabled of Class Module cfClock"
End Property

'---------------------------------------------------------------------------------------
' Property  : weekdayToggleEnabled
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get weekdayToggleEnabled() As Boolean
   On Error GoTo weekdayToggleEnabledGet_Error

    weekdayToggleEnabled = mWeekdayToggleEnabled

   On Error GoTo 0
   Exit Property

weekdayToggleEnabledGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property weekdayToggleEnabled of Class Module cfClock"
End Property


'---------------------------------------------------------------------------------------
' Property  : timeMachineToggleEnabled
' Author    : beededea
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let timeMachineToggleEnabled(ByVal newValue As Boolean)
   
    Dim timeMachineToggleTop As Long: timeMachineToggleTop = 0
    Dim timeMachineToggleLeft As Long: timeMachineToggleLeft = 0

    Dim timeMachineToggleIn As Long: timeMachineToggleIn = 0
    Dim timeMachineToggleOut As Long: timeMachineToggleOut = 0
    Dim fileToPlay As String: fileToPlay = vbNullString
    
    On Error GoTo timeMachineToggleEnabledLet_Error
    
    timeMachineToggleTop = timemachinetoggle.Top
    timeMachineToggleOut = 381
    timeMachineToggleIn = 385

    If newValue = True Then
        timemachinetoggle.Move timeMachineToggleIn, timeMachineToggleTop
        overlayWidget.TimeMachine = True
        gblTimeMachineToggleEnabled = "True"
    Else
        timemachinetoggle.Move timeMachineToggleOut, timeMachineToggleTop
        overlayWidget.TimeMachine = False
        gblTimeMachineToggleEnabled = "False"
    End If
    
    sPutINISetting "Software\SteampunkClockCalendar", "timeMachineToggleEnabled", gblTimeMachineToggleEnabled, gblSettingsFile

    If mTimeMachineToggleEnabled <> newValue Then mTimeMachineToggleEnabled = newValue Else Exit Property
    timeMachineToggleEnabled = mTimeMachineToggleEnabled
    
    clockForm.Refresh
         
   On Error GoTo 0
   Exit Property

timeMachineToggleEnabledLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property timeMachineToggleEnabled of Class Module cfClock"
End Property


'---------------------------------------------------------------------------------------
' Property  : timeMachineToggleEnabled
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get timeMachineToggleEnabled() As Boolean
   On Error GoTo timeMachineToggleEnabledGet_Error

    timeMachineToggleEnabled = mTimeMachineToggleEnabled

   On Error GoTo 0
   Exit Property

timeMachineToggleEnabledGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property timeMachineToggleEnabled of Class Module cfClock"
End Property

'---------------------------------------------------------------------------------------
' Property  : setToggleEnabled
' Author    : beededea
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let settoggleEnabled(ByVal newValue As Boolean)
   
    Dim settoggleTop As Long: settoggleTop = 0
    Dim setToggleIn As Long: setToggleIn = 0
    Dim setToggleOut As Long: setToggleOut = 0
    
    On Error GoTo setToggleEnabledLet_Error
    
    setToggleOut = 359
    setToggleIn = 363
    settoggleTop = settoggle.Top
    
    If newValue = True Then
        settoggle.Move setToggleIn, settoggleTop
        gblsetToggleEnabled = "True"
    Else
        settoggle.Move setToggleOut, settoggleTop
        gblsetToggleEnabled = "False"
    End If
    
    sPutINISetting "Software\SteampunkClockCalendar", "setToggleEnabled", gblsetToggleEnabled, gblSettingsFile

    If msetToggleEnabled <> newValue Then msetToggleEnabled = newValue Else Exit Property
    settoggleEnabled = msetToggleEnabled

    clockForm.Refresh
      
   On Error GoTo 0
   Exit Property

setToggleEnabledLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property setToggleEnabled of Class Module cfClock"
End Property

'---------------------------------------------------------------------------------------
' Property  : setToggleEnabled
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get settoggleEnabled() As Boolean
   On Error GoTo setToggleEnabledGet_Error

    settoggleEnabled = msetToggleEnabled

   On Error GoTo 0
   Exit Property

setToggleEnabledGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property setToggleEnabled of Class Module cfClock"
End Property


'---------------------------------------------------------------------------------------
' Property  : muteToggleEnabled
' Author    : beededea
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let muteToggleEnabled(ByVal newValue As Boolean)
   
    Dim muteToggleTop As Long: muteToggleTop = 0
    Dim muteToggleLeft As Long: muteToggleLeft = 0
    Dim fileToPlay As String: fileToPlay = vbNullString
    
    Dim muteToggleIn As Long: muteToggleIn = 0
    Dim muteToggleOut As Long: muteToggleOut = 0
    Dim lngReturn As Long: lngReturn = 0
    
    On Error GoTo muteToggleEnabledLet_Error
    
    muteToggleTop = mutetoggle.Top
    muteToggleOut = 359
    muteToggleIn = 363
    
    fileToPlay = "clunk.wav"

    If newValue = True Then
        mutetoggle.Move muteToggleIn, muteToggleTop
        gblMuteToggleEnabled = "True"
        ticking = False
        gblEnableSounds = "0"
        widgetPrefs.chkEnableSounds.Value = 0
    Else
        mutetoggle.Move muteToggleOut, muteToggleTop
        gblMuteToggleEnabled = "False"
        gblEnableSounds = "1"
        If ticking = True Then ticking = True
        
        
        widgetPrefs.chkEnableSounds.Value = 1
    End If
    
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    sPutINISetting "Software\SteampunkClockCalendar", "muteToggleEnabled", gblMuteToggleEnabled, gblSettingsFile

    If mMuteToggleEnabled <> newValue Then mMuteToggleEnabled = newValue Else Exit Property
    muteToggleEnabled = mMuteToggleEnabled

    clockForm.Refresh
      
   On Error GoTo 0
   Exit Property

muteToggleEnabledLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property muteToggleEnabled of Class Module cfClock"
End Property

'---------------------------------------------------------------------------------------
' Property  : muteToggleEnabled
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get muteToggleEnabled() As Boolean
   On Error GoTo muteToggleEnabledGet_Error

    muteToggleEnabled = mMuteToggleEnabled

   On Error GoTo 0
   Exit Property

muteToggleEnabledGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property muteToggleEnabled of Class Module cfClock"
End Property




'---------------------------------------------------------------------------------------
' Property  : pendulumToggleEnabled
' Author    : beededea
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let pendulumToggleEnabled(ByVal newValue As Boolean)
       
    Dim pendulumToggleTop As Long: pendulumToggleTop = 0
    Dim pendulumToggleIn As Long: pendulumToggleIn = 0
    Dim pendulumToggleOut As Long: pendulumToggleOut = 0
    Dim fileToPlay As String: fileToPlay = vbNullString
    
    On Error GoTo pendulumToggleEnabledLet_Error
    
    pendulumToggleTop = pendulumtoggle.Top
    pendulumToggleOut = 359
    pendulumToggleIn = 363
    
    If newValue = True Then
        pendulumtoggle.Move pendulumToggleIn, pendulumToggleTop
        gblPendulumToggleEnabled = "True"
        If pendulumInitiatingControl = "pendulumToggle" Then pendulumEnabled = True
    Else
        pendulumtoggle.Move pendulumToggleOut, pendulumToggleTop
        gblPendulumToggleEnabled = "False"
        If pendulumInitiatingControl = "pendulumToggle" Then pendulumEnabled = False
    End If
        
    fileToPlay = "clunk.wav"
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    sPutINISetting "Software\SteampunkClockCalendar", "pendulumToggleEnabled", gblPendulumToggleEnabled, gblSettingsFile

    If mPendulumToggleEnabled <> newValue Then mPendulumToggleEnabled = newValue Else Exit Property
    pendulumToggleEnabled = mPendulumToggleEnabled
    
    clockForm.Refresh
      
   On Error GoTo 0
   Exit Property

pendulumToggleEnabledLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property pendulumToggleEnabled of Class Module cfClock"
End Property

'---------------------------------------------------------------------------------------
' Property  : pendulumToggleEnabled
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get pendulumToggleEnabled() As Boolean
   On Error GoTo pendulumToggleGet_Error

    pendulumToggleEnabled = mPendulumToggleEnabled

   On Error GoTo 0
   Exit Property

pendulumToggleGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property pendulumToggleEnabled of Class Module cfClock"
End Property



'---------------------------------------------------------------------------------------
' Property  : pendulumEnabled
' Author    : beededea
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let pendulumEnabled(ByVal newValue As Boolean)
       
    On Error GoTo pendulumEnabledLet_Error
    
    If newValue = True Then
        gblPendulumEnabled = "True"
        If pendulumInitiatingControl = "pendulum" Then pendulumToggleEnabled = False ' prevents recurring calls
    Else
        gblPendulumEnabled = "False"
        If pendulumInitiatingControl = "pendulum" Then pendulumToggleEnabled = True
    End If
            
    sPutINISetting "Software\SteampunkClockCalendar", "pendulumEnabled", gblPendulumEnabled, gblSettingsFile

    If mPendulumEnabled <> newValue Then mPendulumEnabled = newValue Else Exit Property
    pendulumEnabled = mPendulumEnabled

    clockForm.Refresh
      
   On Error GoTo 0
   Exit Property

pendulumEnabledLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property pendulumEnabled of Class Module cfClock"
End Property

'---------------------------------------------------------------------------------------
' Property  : pendulumEnabled
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get pendulumEnabled() As Boolean
   On Error GoTo pendulumGet_Error

    pendulumEnabled = mPendulumEnabled

   On Error GoTo 0
   Exit Property

pendulumGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property pendulumEnabled of Class Module cfClock"
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
'    SystemVolume.SetMasterVolume timeShiftPercent
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
'    sliderset.Move sldX, sldFixedTop 'move the sliderset-widget accordingly
'
'    fClock.clockForm.Refresh
'
'   On Error GoTo 0
'   Exit Sub
'
'tmrSampleAudioVolume_Timer_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrSampleAudioVolume_Timer of Class Module cfClockTemp"
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
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrSampleAudioMute_Timer of Class Module cfClockTemp"
'End Sub




' ----------------------------------------------------------------
' Procedure Name: fDayOfWeek
' Purpose:
' Procedure Kind: Function
' Procedure Access: Private
' Return Type: String
' Author: beededea
' Date: 17/06/2024
' ----------------------------------------------------------------
Private Function fDayOfWeek() As String
    On Error GoTo fDayOfWeek_Error
     Dim vb6DateTime As Date
     
     vb6DateTime = Date

     Select Case DatePart("w", vb6DateTime)
         Case vbSunday
             fDayOfWeek = "Sunday"
         Case vbMonday
             fDayOfWeek = "Monday"
         Case vbTuesday
             fDayOfWeek = "Tuesday"
         Case vbWednesday
             fDayOfWeek = "Wednesday"
         Case vbThursday
             fDayOfWeek = "Thursday"
         Case vbFriday
             fDayOfWeek = "Friday"
         Case vbSaturday
             fDayOfWeek = "Saturday"
     End Select
     
    
    On Error GoTo 0
    Exit Function

fDayOfWeek_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fDayOfWeek, line " & Erl & "."

End Function
     


'---------------------------------------------------------------------------------------
' Property  : alarmtoggle1Enabled
' Author    : beededea
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let alarmtoggle1Enabled(ByVal newValue As Boolean)

    Dim alarmtoggle1Top As Long: alarmtoggle1Top = 0
    Dim alarmtoggle1In As Long: alarmtoggle1In = 0
    Dim alarmtoggle1Out As Long: alarmtoggle1Out = 0
    Dim alphaVal As Integer: alphaVal = 0
   
    On Error GoTo alarmtoggle1EnabledLet_Error
    
    alarmtoggle1Top = alarmtoggle1.Top
    alarmtoggle1Out = 664
    alarmtoggle1In = 654
    
    alarmInitiatingControl = "alarm1"
    
    alphaVal = Val(gblOpacity) / 100
 
    If newValue = True Then
        gblAlarmToggle1Enabled = "True"
        alarmtoggle1.Move alarmtoggle1Out, alarmtoggle1Top
        Call hideTills
        fClock.clockForm.Widgets("alarm1till").Widget.Alpha = alphaVal
        
        alarmtoggle1.ImageKey = "alarmtoggle1armed"
'        If alarmtoggle1InitiatingControl = "alarmtoggle1" Then alarmtoggle1ToggleEnabled = False ' prevents recurring calls
    Else
        gblAlarmToggle1Enabled = "False"
        alarmtoggle1.Move alarmtoggle1In, alarmtoggle1Top
        Call hideTills
        fClock.clockForm.Widgets("alarm1till").Widget.Alpha = 0

        alarmtoggle1.ImageKey = "alarmtoggle1"
'        If alarmtoggle1InitiatingControl = "alarmtoggle1" Then alarmtoggle1ToggleEnabled = True
    End If
    
    sPutINISetting "Software\SteampunkClockCalendar", "alarmToggle1Enabled", gblAlarmToggle1Enabled, gblSettingsFile

    If mAlarmToggle1Enabled <> newValue Then mAlarmToggle1Enabled = newValue Else Exit Property
    alarmtoggle1Enabled = mAlarmToggle1Enabled
    
    clockForm.Refresh
     
   On Error GoTo 0
   Exit Property

alarmtoggle1EnabledLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property alarmtoggle1Enabled of Class Module cfClock"
End Property

'---------------------------------------------------------------------------------------
' Property  : alarmtoggle1Enabled
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get alarmtoggle1Enabled() As Boolean
   On Error GoTo alarmtoggle1EnabledGet_Error

    alarmtoggle1Enabled = mAlarmToggle1Enabled

   On Error GoTo 0
   Exit Property

alarmtoggle1EnabledGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property alarmtoggle1Enabled of Class Module cfClock"
End Property



'---------------------------------------------------------------------------------------
' Property  : alarmtoggle2Enabled
' Author    : beededea
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let alarmtoggle2Enabled(ByVal newValue As Boolean)

    Dim alarmtoggle2Top As Long: alarmtoggle2Top = 0
    Dim alarmtoggle2In As Long: alarmtoggle2In = 0
    Dim alarmtoggle2Out As Long: alarmtoggle2Out = 0
    Dim alphaVal As Integer: alphaVal = 0
   
    On Error GoTo alarmtoggle2EnabledLet_Error
    
    alarmtoggle2Top = alarmtoggle2.Top
    alarmtoggle2Out = 678
    alarmtoggle2In = 670

    alarmInitiatingControl = "alarm2"
        
    alphaVal = Val(gblOpacity) / 100
 
    If newValue = True Then
        gblAlarmToggle2Enabled = "True"
        alarmtoggle2.Move alarmtoggle2Out, alarmtoggle2Top
        Call hideTills
        fClock.clockForm.Widgets("alarm2till").Widget.Alpha = alphaVal
        
        alarmtoggle2.ImageKey = "alarmtoggle2armed"

'        If alarmtoggle2InitiatingControl = "alarmtoggle2" Then alarmtoggle2ToggleEnabled = False ' prevents recurring calls
    Else
        gblAlarmToggle2Enabled = "False"
        alarmtoggle2.Move alarmtoggle2In, alarmtoggle2Top
        Call hideTills
        fClock.clockForm.Widgets("alarm2till").Widget.Alpha = 0
        
        alarmtoggle2.ImageKey = "alarmtoggle2"

'        If alarmtoggle2InitiatingControl = "alarmtoggle2" Then alarmtoggle2ToggleEnabled = True
    End If
    
    sPutINISetting "Software\SteampunkClockCalendar", "alarmtoggle2Enabled", gblAlarmToggle2Enabled, gblSettingsFile

    If mAlarmToggle2Enabled <> newValue Then
        mAlarmToggle2Enabled = newValue
        alarmtoggle2Enabled = mAlarmToggle2Enabled
    Else
        Exit Property
    End If
    
    clockForm.Refresh
     
   On Error GoTo 0
   Exit Property

alarmtoggle2EnabledLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property alarmtoggle2Enabled of Class Module cfClock"
End Property

'---------------------------------------------------------------------------------------
' Property  : alarmtoggle2Enabled
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get alarmtoggle2Enabled() As Boolean
   On Error GoTo alarmtoggle2EnabledGet_Error

    alarmtoggle2Enabled = mAlarmToggle2Enabled

   On Error GoTo 0
   Exit Property

alarmtoggle2EnabledGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property alarmtoggle2Enabled of Class Module cfClock"
End Property



'---------------------------------------------------------------------------------------
' Property  : alarmtoggle3Enabled
' Author    : beededea
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let alarmtoggle3Enabled(ByVal newValue As Boolean)

    Dim alarmtoggle3Top As Long: alarmtoggle3Top = 0
    Dim alarmtoggle3In As Long: alarmtoggle3In = 0
    Dim alarmtoggle3Out As Long: alarmtoggle3Out = 0
    Dim alphaVal As Integer: alphaVal = 0
   
    On Error GoTo alarmtoggle3EnabledLet_Error
    
    alarmtoggle3Top = alarmtoggle3.Top
    alarmtoggle3Out = 690
    alarmtoggle3In = 682
        
    alarmInitiatingControl = "alarm3"
    
    alphaVal = Val(gblOpacity) / 100
 
    If newValue = True Then
        gblAlarmToggle3Enabled = "True"
        alarmtoggle3.Move alarmtoggle3Out, alarmtoggle3Top
        Call hideTills
        fClock.clockForm.Widgets("alarm3till").Widget.Alpha = alphaVal
        
        alarmtoggle3.ImageKey = "alarmtoggle3armed"

'        If alarmtoggle3InitiatingControl = "alarmtoggle3" Then alarmtoggle3ToggleEnabled = False ' prevents recurring calls
    Else
        gblAlarmToggle3Enabled = "False"
        alarmtoggle3.Move alarmtoggle3In, alarmtoggle3Top
        Call hideTills
        fClock.clockForm.Widgets("alarm3till").Widget.Alpha = 0
        
        alarmtoggle3.ImageKey = "alarmtoggle3"
'        If alarmtoggle3InitiatingControl = "alarmtoggle3" Then alarmtoggle3ToggleEnabled = True
    End If
    
    sPutINISetting "Software\SteampunkClockCalendar", "alarmtoggle3Enabled", gblAlarmToggle3Enabled, gblSettingsFile

    If mAlarmToggle3Enabled <> newValue Then mAlarmToggle3Enabled = newValue Else Exit Property
    alarmtoggle3Enabled = mAlarmToggle3Enabled
    
    clockForm.Refresh
     
   On Error GoTo 0
   Exit Property

alarmtoggle3EnabledLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property alarmtoggle3Enabled of Class Module cfClock"
End Property

'---------------------------------------------------------------------------------------
' Property  : alarmtoggle3Enabled
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get alarmtoggle3Enabled() As Boolean
   On Error GoTo alarmtoggle3EnabledGet_Error

    alarmtoggle3Enabled = mAlarmToggle3Enabled

   On Error GoTo 0
   Exit Property

alarmtoggle3EnabledGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property alarmtoggle3Enabled of Class Module cfClock"
End Property


'---------------------------------------------------------------------------------------
' Property  : alarmtoggle4Enabled
' Author    : beededea
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let alarmtoggle4Enabled(ByVal newValue As Boolean)

    Dim alarmtoggle4Top As Long: alarmtoggle4Top = 0
    Dim alarmtoggle4In As Long: alarmtoggle4In = 0
    Dim alarmtoggle4Out As Long: alarmtoggle4Out = 0
    Dim alphaVal As Integer: alphaVal = 0
   
    On Error GoTo alarmtoggle4EnabledLet_Error
    
    alarmtoggle4Top = alarmtoggle4.Left
    alarmtoggle4Top = alarmtoggle4.Top
    alarmtoggle4Out = 690
    alarmtoggle4In = 682

    alarmInitiatingControl = "alarm4"
    
    alphaVal = Val(gblOpacity) / 100
 
    If newValue = True Then
        gblAlarmToggle4Enabled = "True"
        alarmtoggle4.Move alarmtoggle4Out, alarmtoggle4Top
        Call hideTills
        fClock.clockForm.Widgets("alarm4till").Widget.Alpha = alphaVal
        
        alarmtoggle4.ImageKey = "alarmtoggle4armed"

'        If alarmtoggle4InitiatingControl = "alarmtoggle4" Then alarmtoggle4ToggleEnabled = False ' prevents recurring calls
    Else
        gblAlarmToggle4Enabled = "False"
        alarmtoggle4.Move alarmtoggle4In, alarmtoggle4Top
        Call hideTills
        fClock.clockForm.Widgets("alarm4till").Widget.Alpha = 0
        
        alarmtoggle4.ImageKey = "alarmtoggle4"
        
'        If alarmtoggle4InitiatingControl = "alarmtoggle4" Then alarmtoggle4ToggleEnabled = True
    End If
    
    sPutINISetting "Software\SteampunkClockCalendar", "alarmtoggle4Enabled", gblAlarmToggle4Enabled, gblSettingsFile

    If mAlarmToggle4Enabled <> newValue Then mAlarmToggle4Enabled = newValue Else Exit Property
    alarmtoggle4Enabled = mAlarmToggle4Enabled
    
    clockForm.Refresh
     
   On Error GoTo 0
   Exit Property

alarmtoggle4EnabledLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property alarmtoggle4Enabled of Class Module cfClock"
End Property

'---------------------------------------------------------------------------------------
' Property  : alarmtoggle4Enabled
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get alarmtoggle4Enabled() As Boolean
   On Error GoTo alarmtoggle4EnabledGet_Error

    alarmtoggle4Enabled = mAlarmToggle4Enabled

   On Error GoTo 0
   Exit Property

alarmtoggle4EnabledGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property alarmtoggle4Enabled of Class Module cfClock"
End Property



'---------------------------------------------------------------------------------------
' Property  : alarmtoggle5Enabled
' Author    : beededea
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let alarmtoggle5Enabled(ByVal newValue As Boolean)

    Dim alarmtoggle5Top As Long: alarmtoggle5Top = 0
    Dim alarmtoggle5In As Long: alarmtoggle5In = 0
    Dim alarmtoggle5Out As Long: alarmtoggle5Out = 0
    Dim alphaVal As Integer: alphaVal = 0
   
    On Error GoTo alarmtoggle5EnabledLet_Error
    
    alarmtoggle5Top = alarmtoggle5.Left
    alarmtoggle5Top = alarmtoggle5.Top
    alarmtoggle5Out = 678
    alarmtoggle5In = 670
    
    alarmInitiatingControl = "alarm5"
        
    alphaVal = Val(gblOpacity) / 100
 
    If newValue = True Then
        gblAlarmToggle5Enabled = "True"
        alarmtoggle5.Move alarmtoggle5Out, alarmtoggle5Top
        Call hideTills
        fClock.clockForm.Widgets("alarm5till").Widget.Alpha = alphaVal
        
        alarmtoggle5.ImageKey = "alarmtoggle5armed"
        
'        If alarmtoggle5InitiatingControl = "alarmtoggle5" Then alarmtoggle5ToggleEnabled = False ' prevents recurring calls
    Else
        gblAlarmToggle5Enabled = "False"
        alarmtoggle5.Move alarmtoggle5In, alarmtoggle5Top
        Call hideTills
        fClock.clockForm.Widgets("alarm5till").Widget.Alpha = 0
        
        alarmtoggle5.ImageKey = "alarmtoggle5"
        
'        If alarmtoggle5InitiatingControl = "alarmtoggle5" Then alarmtoggle5ToggleEnabled = True
    End If
    
    sPutINISetting "Software\SteampunkClockCalendar", "alarmtoggle5Enabled", gblAlarmToggle5Enabled, gblSettingsFile

    If mAlarmToggle5Enabled <> newValue Then mAlarmToggle5Enabled = newValue Else Exit Property
    alarmtoggle5Enabled = mAlarmToggle5Enabled
    
    clockForm.Refresh
     
   On Error GoTo 0
   Exit Property

alarmtoggle5EnabledLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property alarmtoggle5Enabled of Class Module cfClock"
End Property

'---------------------------------------------------------------------------------------
' Property  : alarmtoggle5Enabled
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get alarmtoggle5Enabled() As Boolean
   On Error GoTo alarmtoggle5EnabledGet_Error

    alarmtoggle5Enabled = mAlarmToggle5Enabled

   On Error GoTo 0
   Exit Property

alarmtoggle5EnabledGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property alarmtoggle5Enabled of Class Module cfClock"
End Property



' ----------------------------------------------------------------
' Procedure Name: hideTills
' Purpose:
' Procedure Kind: sub
' Procedure Access: Private
' Return Type: String
' Author: beededea
' Date: 17/06/2024
' ----------------------------------------------------------------
Private Sub hideTills()
    On Error GoTo fDayOfWeek_Error
            
    If alarmInitiatingControl <> "alarm1" And alarmtoggle1Enabled = True Then alarmtoggle1Enabled = False
    If alarmInitiatingControl <> "alarm2" And alarmtoggle2Enabled = True Then alarmtoggle2Enabled = False
    If alarmInitiatingControl <> "alarm3" And alarmtoggle3Enabled = True Then alarmtoggle3Enabled = False
    If alarmInitiatingControl <> "alarm4" And alarmtoggle4Enabled = True Then alarmtoggle4Enabled = False
    If alarmInitiatingControl <> "alarm5" And alarmtoggle5Enabled = True Then alarmtoggle5Enabled = False
    
    On Error GoTo 0
    Exit Sub

fDayOfWeek_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fDayOfWeek, line " & Erl & "."

End Sub




' ----------------------------------------------------------------
' Procedure Name: tmrZZZZ_Timer
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 21/06/2024
' ----------------------------------------------------------------
Private Sub tmrZZZZ_Timer()
    Dim fileToPlay As String: fileToPlay = vbNullString
    
    On Error GoTo tmrZZZZ_Timer_Error
    
    fileToPlay = "longzzzz.wav"
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    tmrZZZZ.Enabled = False
    
   On Error GoTo 0
   Exit Sub

tmrZZZZ_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrZZZZ_Timer of Class Module cfClock"
End Sub






' ----------------------------------------------------------------
' Procedure Name: tmrTicking_Timer
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 21/06/2024
' ----------------------------------------------------------------
Private Sub tmrTicking_Timer()
    Dim fileToPlay As String: fileToPlay = vbNullString
    Dim lngReturn As Long: lngReturn = 0
    
    On Error GoTo tmrTicking_Timer_Error
    
    tmrTicking.Enabled = False
    tmrTicking.Interval = 30000 ' 30 secs
    
    fileToPlay = App.path & "\resources\sounds\" & "ticktock.wav"
    If gblEnableSounds = "1" And fFExists(fileToPlay) Then
    
        lngReturn = mciSendString("Play " & fileToPlay, "", 0, 0)
        'PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
        
    tmrTicking.Enabled = True
    
   On Error GoTo 0
   Exit Sub

tmrTicking_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrTicking_Timer of Class Module cfClock"
End Sub






'---------------------------------------------------------------------------------------
' Property  : ticking
' Author    : beededea
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let ticking(ByVal newValue As Boolean)
    Dim fileToPlay As String: fileToPlay = vbNullString
    Dim alphaVal As Integer: alphaVal = 0
    Dim mciFileToPlay As String: mciFileToPlay = vbNullString
    Dim lngReturn As Long: lngReturn = 0
    
    On Error GoTo tickingLet_Error
    
    alphaVal = Val(gblOpacity) / 100
    
    If newValue = True Then
        tmrTicking.Interval = 300
        tmrTicking.Enabled = True
     Else
        tmrTicking.Enabled = False
        
        ' stop MCI from playing
        fileToPlay = App.path & "\resources\sounds\" & "ticktock.wav"
        If fFExists(fileToPlay) Then
            lngReturn = mciSendString("stop " & fileToPlay, "", 0, 0)
        End If
    End If

    If mTicking <> newValue Then mTicking = newValue Else Exit Property
    ticking = mTicking
    
    clockForm.Refresh
     
   On Error GoTo 0
   Exit Property

tickingLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property ticking of Class Module cfClock"
End Property

'---------------------------------------------------------------------------------------
' Property  : ticking
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get ticking() As Boolean
   On Error GoTo tickingGet_Error

    ticking = mTicking

   On Error GoTo 0
   Exit Property

tickingGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property ticking of Class Module cfClock"
End Property
