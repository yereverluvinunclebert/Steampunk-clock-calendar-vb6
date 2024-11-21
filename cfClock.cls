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
'---------------------------------------------------------------------------------------
' Module    : cfClock
' Author    : beededea
' Date      : 18/10/2024
' Purpose   :
'---------------------------------------------------------------------------------------

'@IgnoreModule IntegerDataType, ModuleWithoutFolder
Option Explicit

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" _
 (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, _
  ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Public WithEvents clockForm As cWidgetForm
Attribute clockForm.VB_VarHelpID = -1

'Private WithEvents tmrDisplay As cTimer
Public WithEvents tmrZZZZSound As cTimer
Attribute tmrZZZZSound.VB_VarHelpID = -1
Public WithEvents tmrTILLSound As cTimer
Attribute tmrTILLSound.VB_VarHelpID = -1
Public WithEvents tmrTickingSound As cTimer
Attribute tmrTickingSound.VB_VarHelpID = -1

Public collPSDNonUIElements As cCollection
Public FX As Long
Public FY As Long
Public FZ As Single  'available on the outside, even after the Form is closed

' private member properties

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
Private mCrankedRaised As Boolean
Private mTicking As Boolean
Private mPendulumEnabled As Boolean
Private mAlarmtoggleEnabled As Boolean
Private mAlarmToggle1Enabled As Boolean
Private mAlarmToggle2Enabled As Boolean
Private mAlarmToggle3Enabled As Boolean
Private mAlarmToggle4Enabled As Boolean
Private mAlarmToggle5Enabled As Boolean
Private mtimeShiftValue As Currency

Public mShowHelp As Boolean 'change to property get/set

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
Private WithEvents helpToggle As cWidgetBase
Attribute helpToggle.VB_VarHelpID = -1
Private WithEvents help1toggle As cWidgetBase
Attribute help1toggle.VB_VarHelpID = -1
Private WithEvents help2toggle As cWidgetBase
Attribute help2toggle.VB_VarHelpID = -1
Private WithEvents help3toggle As cWidgetBase
Attribute help3toggle.VB_VarHelpID = -1

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
Private WithEvents weekdayred As cWidgetBase
Attribute weekdayred.VB_VarHelpID = -1
Private WithEvents weekdaytill As cWidgetBase
Attribute weekdaytill.VB_VarHelpID = -1
Private WithEvents weekdayshadow As cWidgetBase
Attribute weekdayshadow.VB_VarHelpID = -1
Private WithEvents dropdown As cWidgetBase
Attribute dropdown.VB_VarHelpID = -1
Private WithEvents bellset As cWidgetBase
Attribute bellset.VB_VarHelpID = -1
Private WithEvents helpdropdown As cWidgetBase
Attribute helpdropdown.VB_VarHelpID = -1
Private WithEvents pendulumtransparent As cWidgetBase
Attribute pendulumtransparent.VB_VarHelpID = -1
Private WithEvents glow As cWidgetBase
Attribute glow.VB_VarHelpID = -1
Private WithEvents clockset As cWidgetBase
Attribute clockset.VB_VarHelpID = -1

Private lblTerminalText As cwLabel

' test labels, textboxes causing IDE crash
Private newLabel As cwLabel
Attribute newLabel.VB_VarHelpID = -1
Private WithEvents textBox As cwTextBox
Attribute textBox.VB_VarHelpID = -1

'general vars

Private pvtSldFixedTop As Double
Private pvtSldXMax As Double
Private pvtSldXMin As Double
Private pvtSldX As Double
Private pvtHelpMoved As Boolean
Private pvtMuteInitiatingControl As String
Private pvtAlarmInitiatingControl As String
Private pvtSlidersetMousedown As Boolean
Private pvtPSDWidth As Long
Private pvtPSDHeight As Long


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
    
    Set tmrZZZZSound = Nothing
    Set tmrZZZZSound = New_c.Timer(100, False)
        
    Set tmrTILLSound = Nothing
    Set tmrTILLSound = New_c.Timer(300, False)
    
    Set tmrTickingSound = Nothing
    Set tmrTickingSound = New_c.Timer(300, False)
    
'    Set tmrDisplay = Nothing
'    Set tmrDisplay = New_c.Timer(1000, True)

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

         If pvtHelpMoved = True Then  ' now copes with f5 reloads
             helpbottom.Move 220, helpbottomTop
             labellayer.Move 2, labellayerTop
             pvtHelpMoved = False
         End If
         
         helpbottom.Alpha = Val(gblOpacity) / 100
         labellayer.Alpha = Val(gblOpacity) / 100
         
    Else

        ' the next two lines are a temporary hack.
        ' when widgets are made lapha = 0 they are still clickable, I can move these to the overlay or I
        ' can just move them out of the way so that any click on an invisible element does not interfere with one 'underneath'
        
         helpbottom.Alpha = 0
         labellayer.Alpha = 0

        If pvtHelpMoved = False Then
            helpbottom.Move helpbottomLeft + 700, helpbottomTop
            labellayer.Move labellayerLeft + 400, labellayerTop ' the 400 is a magical number that moves them just out of the way without causing tearing
            pvtHelpMoved = True
        End If
    
    End If
    
    If mShowHelp <> newValue Then mShowHelp = newValue Else Exit Property
    ShowHelp = mShowHelp
    
    clockForm.Refresh
       
    If gblVolumeBoost = "1" Then
        fileToPlay = "pagefumble.wav"
    Else
        fileToPlay = "pagefumble-quiet.wav"
    End If
    
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
' Author    : olaf schmidt and me
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
    Dim grey As String: grey = vbNullString
    Dim someOpacity As Double: someOpacity = 0
    Dim thisZorder As Long: thisZorder = 0
    
    On Error GoTo InitFromPSD_Error

    grey = RGB(18, 18, 18)
    
    someOpacity = Val(gblOpacity) / 100

    'create the Top-Level-Form
    Set clockForm = Cairo.WidgetForms.Create(IIf(App.LogMode, AlphaNoTaskbarEntry, AlphaWithTaskbarEntry), widgetName, True, 1, 1)
        clockForm.WidgetRoot.BackColor = -1 ' transparent
 
    With New_c.SimplePSD(PSD_FileNameOrByteArray)  'create a new PSD-Parser.instance (and load the passed content)
        pvtPSDWidth = .Width
        pvtPSDHeight = .Height       'store the original Psd-Pixel-Width/Height in Private vars (as the base from which we calculate the zoomed Form-Width/Height)
    
        For I = 0 To .LayersCount - 1 ' loop through each of the Layers in the PSD
            If .LayerByteSize(I) Then  ' this is a true Alpha-Surface-Layer and not just a PSD layer 'group'
                If .LayerAlphaPercent(I) > 0 Then ' only handles layers that have an opacity greater than 0 - need to note this for the future, this will cause a problem!
                    'add each current Layer path and surface object into the global ImageList collection (using LayerPath as the ImageKey)
                    Cairo.ImageList.AddSurface .LayerPath(I), .LayerSurface(I)
                    
                    If collPSDNonUIElements.Exists(.LayerPath(I)) Then ' we do not create Widgets for those in this exclude list.
                        
                        'we add layer info. (used in cwOverlay) to the "excluded paths" previously defined in Sub Main,
                        collPSDNonUIElements(.LayerPath(I)) = Array(.LayerX(I), .LayerY(I), someOpacity)  'here we update the so far empty slots with the PSD-offsets
                
                    Else 'everything not in the exclude-list will get its own Widget-instance
                        Set W = clockForm.Widgets.Add(New cwAlphaImg, LCase$(.LayerPath(I)), .LayerX(I), .LayerY(I), .LayerWidth(I), .LayerHeight(I)).Widget
                        
                        W.ImageKey = W.Key 'W.Key equals ImageList-Key, set above - and LayerPath(i) at this point ... set it also as the ImageKey of our new created Widget

                        W.Alpha = 0
                    
                        ' note: the clickable layers characteristics are set in adjustMainControls
                        
                        ' all non-clickable Layer-Widgets will be -1 or "non-hoverable" and "fully click-through"
                        W.HoverColor = -1 ' task: might change this later when new ver or RC6 arrives
                        If gblEnableTooltips = "True" Then W.ToolTip = "Ctrl + mouse scrollwheel up/down to resize, you can also drag me to a new position."
                        W.MousePointer = IDC_SIZEALL
                        
                    End If
                End If
            End If
        Next I
    End With '<-- the Parser-instance will be destroyed here (freeing the Memory, the internal PSD-Layers have occupied)
        
    Set overlayWidget = clockForm.Widgets.Add(New cwOverlay, "Overlay", 0, 0, pvtPSDWidth, pvtPSDHeight)
    Set overlayWidget.cCollOverlayImages = collPSDNonUIElements 'make the excluded PSD-Surface-Paths known to the Overlay-Widget
        overlayWidget.cxPerc = 0.645 'define the relative center (within cwOverlay, which has the same size as the Form and the PSD...) this is used for clock-hand-positioning
        overlayWidget.cyPerc = 0.334

    thisZorder = clockForm.Widgets.GetOneBasedChildIndexByKey("displayScreen")

    ' set some Font and ForeColor Properties and draw some Text on the Widget
    Set lblTerminalText = clockForm.Widgets.Add(New_W("cwLabel"), "lblTerminalText", 423, 74, 220, 165, True, thisZorder + 1)

    With lblTerminalText
        .Widget.FontSize = gblDisplayScreenFontSize
        .Widget.FontName = gblDisplayScreenFont
        .Widget.ForeColor = grey
        .Widget.Alpha = someOpacity
        .Widget.FontBold = False
        .BorderWidth = 0
        .WordWrap = True
        .Alignment = 0
        .Widget.BackColor = -1
    End With
    
    gblFClockAvailable = True
    
    'the next 5 lines will init a "moving Slider" (which was defined as "sliderset"-layer in the PSD)
    Set sliderset = clockForm.Widgets("sliderset").Widget 'store the already (in PSD-Loop created) Widget-reference in an explicit "WithEvents-Variable"
        
    With sliderset
        .Moveable = False 'make it not moveable by default
        .HoverColor = 0
        .MousePointer = IDC_HAND
        .Tag = 0.01
    End With
    
    
    pvtSldFixedTop = sliderset.Top   'and store the initial position-data for y-direction (which will not be changed)
    pvtSldXMax = sliderset.Left      'as well as the x-direction related pos (which is initially at the "100%-Volume, right-endswitch" )
    pvtSldXMin = pvtSldXMax - 116       'the Min-Value is simply the Max-Value minus Slider-Range (the 300pxl-range was determined by "probing")

    Set helpbottom = clockForm.Widgets("helpbottom").Widget
    Set screentop = clockForm.Widgets("screentop").Widget
    Set displayscreen = clockForm.Widgets("displayscreen").Widget
    
    Set displayscreentoggle = clockForm.Widgets("displayscreentoggle").Widget
    Set screenbackground = clockForm.Widgets("screenbackground").Widget
    Set weekdaytoggle = clockForm.Widgets("weekdaytoggle").Widget
    Set helpToggle = clockForm.Widgets("helptoggle").Widget
    Set help1toggle = clockForm.Widgets("help1toggle").Widget
    Set help2toggle = clockForm.Widgets("help2toggle").Widget
    Set help3toggle = clockForm.Widgets("help3toggle").Widget
    
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
    Set weekdayred = clockForm.Widgets("weekdayred").Widget
    Set weekdaytill = clockForm.Widgets("weekdaytill").Widget
    Set weekdayshadow = clockForm.Widgets("weekdayshadow").Widget
    Set dropdown = clockForm.Widgets("dropdown").Widget
    Set bellset = clockForm.Widgets("bellset").Widget
    Set helpdropdown = clockForm.Widgets("helpdropdown").Widget
    Set pendulumtransparent = clockForm.Widgets("pendulumtransparent").Widget
    Set clockset = clockForm.Widgets("clockset").Widget
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
'    Dim fileToPlay As String: fileToPlay = vbNullString
    
    On Error GoTo clockForm_BubblingEvent_Error

    If gblIgnoreMouse = "1" Then Exit Sub
    
    stdMessage = "Right Click to open the menu and the preferences. CTRL+ Mouse scrollwheel UP/DOWN to resize. "
    stdSubject = "You can turn off the balloon tooltips in the preferences."
    
    If EventName = "W_MouseMove" Then
        Select Case LCase$(sender.Widget.Key)
            Case "topdigitalclock"
                Call clearWeekdayTill
            Case "sliderset"
                If gblEnableBalloonTooltips = "True" Then CreateToolTip clockForm.hWnd, "This is the Time Slider. Click and drag me left/right to alter time.", _
                     TTIconInfo, widgetName & " Locking pin Help", , , , True
            Case "helpbottom"
                If gblEnableBalloonTooltips = "True" Then CreateToolTip clockForm.hWnd, "This is the Widget Help. Click on me to make me go away.", _
                     TTIconInfo, widgetName & " Help", , , , True
            Case "timemachinetoggle"
                If gblEnableBalloonTooltips = "True" Then CreateToolTip clockForm.hWnd, "Toggle the time machine images.", _
                     TTIconInfo, widgetName & " Time Toggle Help", , , , True
            Case "displayscreentoggle"
                If gblEnableBalloonTooltips = "True" Then CreateToolTip clockForm.hWnd, "Toggle back screen.", _
                     TTIconInfo, widgetName & " Back Screen Toggle Help", , , , True
            Case "weekdaytoggle"
                If gblEnableBalloonTooltips = "True" Then CreateToolTip clockForm.hWnd, "Toggle weekday indicator.", _
                     TTIconInfo, widgetName & " Weekday Toggle Help", , , , True
            Case "helptoggle"
                If gblEnableBalloonTooltips = "True" Then CreateToolTip clockForm.hWnd, "Click me to show the help dropdown canvas.", _
                     TTIconInfo, widgetName & " Help Toggle Help", , , , True
            Case "help1toggle"
                If gblEnableBalloonTooltips = "True" Then CreateToolTip clockForm.hWnd, "Click me to show help dropdown canvas number two.", _
                     TTIconInfo, widgetName & " Help Number 1 Help Toggle", , , , True
            Case "help2toggle"
                If gblEnableBalloonTooltips = "True" Then CreateToolTip clockForm.hWnd, "Click me to show help dropdown canvas number three.", _
                     TTIconInfo, widgetName & " Help Number 2 Help Toggle", , , , True
            Case "help3toggle"
                If gblEnableBalloonTooltips = "True" Then CreateToolTip clockForm.hWnd, "Click me to show help dropdown canvas number one.", _
                     TTIconInfo, widgetName & " Help Number 3 Help Toggle", , , , True
            
            Case "alarmtoggle"
                If gblEnableBalloonTooltips = "True" Then CreateToolTip clockForm.hWnd, "Click here to enter alarm mode.", _
                     TTIconInfo, widgetName & " Alarm Toggle Help", , , , True
            Case "settoggle"
                If gblEnableBalloonTooltips = "True" Then CreateToolTip clockForm.hWnd, "In alarm mode, when you are editing an alarm, press this toggle (or the bell) to save the alarm.", _
                     TTIconInfo, widgetName & " Loud Toggle Help", , , , True
            Case "mutetoggle"
                If gblEnableBalloonTooltips = "True" Then CreateToolTip clockForm.hWnd, "Mute chimes and ALL sounds on/off.", _
                     TTIconInfo, widgetName & " Mute Toggle Help", , , , True
            Case "pendulumtoggle"
                If gblEnableBalloonTooltips = "True" Then CreateToolTip clockForm.hWnd, "Stop/start the pendulum by clicking on it.", _
                     TTIconInfo, widgetName & " Pendulum Toggle Help", , , , True
            Case "clockset"
                If gblEnableBalloonTooltips = "True" Then CreateToolTip clockForm.hWnd, "This is just the analogue clock face, clicking here does nothing at the moment.", _
                     TTIconInfo, widgetName & " Clock face Help", , , , True
            Case "pendulumtransparent"
                If gblEnableBalloonTooltips = "True" Then CreateToolTip clockForm.hWnd, "Stop/start the pendulum.", _
                     TTIconInfo, widgetName & " Pendulum Help", , , , True
            Case "displayscreen"
                If gblEnableBalloonTooltips = "True" Then CreateToolTip clockForm.hWnd, "To lower kinematoscope press S toggle. " & vbCrLf & vbCrLf & " Scroll up the display screen text by double-clicking on it - each time sending a carriage return to the screen.", _
                    TTIconInfo, widgetName & " Kinematoscope Toggle Help", , , , True
                Call clearWeekdayTill
                With clockForm.Widgets("redalarmcover").Widget
                    .Alpha = 0
                    .Refresh
                End With
            Case "bellset"
                If gblEnableBalloonTooltips = "True" Then CreateToolTip clockForm.hWnd, "Click here to set or edit an alarm.", _
                     TTIconInfo, widgetName & " Alarm Toggle Help", , , , True
                With clockForm.Widgets("redalarmcover").Widget
                    .Alpha = 0
                    .Refresh
                End With
            Case "clapperleft"
                If gblEnableBalloonTooltips = "True" Then CreateToolTip clockForm.hWnd, "Chiming is currently enabled, click to mute the clock chimes.", _
                     TTIconInfo, widgetName & " Chime Mute Toggle Help", , , , True
            Case "clapperright"
                If gblEnableBalloonTooltips = "True" Then CreateToolTip clockForm.hWnd, "Chiming is currently disabled, click to enable the clock chimes..", _
                     TTIconInfo, widgetName & " Chime Enable Toggle Help", , , , True
            Case "dropdown"
                If gblEnableBalloonTooltips = "True" Then CreateToolTip clockForm.hWnd, "Click me for information.", _
                     TTIconInfo, widgetName & " Dropdown Toggle Help", , , , True
            Case "backtoggle"
                If gblEnableBalloonTooltips = "True" Then CreateToolTip clockForm.hWnd, "Toggle the Back Screen to make the text easier to read.", _
                     TTIconInfo, widgetName & " Back Screen Toggle Help", , , , True
            Case "alarmtoggle1"
                If gblEnableBalloonTooltips = "True" Then CreateToolTip clockForm.hWnd, "Click to set or view alarm number one.", _
                     TTIconInfo, widgetName & " Alarm 1 Toggle Help", , , , True
            Case "alarmtoggle2"
                If gblEnableBalloonTooltips = "True" Then CreateToolTip clockForm.hWnd, "Click to set or view alarm number two.", _
                     TTIconInfo, widgetName & " Alarm 2 Toggle Help", , , , True
            Case "alarmtoggle3"
                If gblEnableBalloonTooltips = "True" Then CreateToolTip clockForm.hWnd, "Click to set or view alarm number three.", _
                     TTIconInfo, widgetName & " Alarm 3 Toggle Help", , , , True
            Case "alarmtoggle4"
                If gblEnableBalloonTooltips = "True" Then CreateToolTip clockForm.hWnd, "Click to set or view alarm number four.", _
                     TTIconInfo, widgetName & " Alarm 4 Toggle Help", , , , True
            Case "alarmtoggle5"
                If gblEnableBalloonTooltips = "True" Then CreateToolTip clockForm.hWnd, "Click to set or view alarm number five.", _
                     TTIconInfo, widgetName & " Alarm 5 Toggle Help", , , , True
            Case "chain"
                If gblEnableBalloonTooltips = "True" Then CreateToolTip clockForm.hWnd, "Pulling the chain will silence the ticks.", _
                     TTIconInfo, widgetName & " Chain Help", , , , True
            Case "crankup"
                If gblEnableBalloonTooltips = "True" Then CreateToolTip clockForm.hWnd, "Crank me down to quieten the whole clock.", _
                     TTIconInfo, widgetName & " Crank Help", , , , True
            Case "crankdown"
                If gblEnableBalloonTooltips = "True" Then CreateToolTip clockForm.hWnd, "Crank up the sound! Crank me up to make more sound throughout the whole clock.", _
                     TTIconInfo, widgetName & " Crank Help", , , , True
            Case "lockingpinlocked"
                If gblEnableBalloonTooltips = "True" Then CreateToolTip clockForm.hWnd, "Click here to unlock the clock and make it moveable using mouse.", _
                     TTIconInfo, widgetName & " Locking Pin Help", , , , True
            Case "lockingpin"
                If gblEnableBalloonTooltips = "True" Then CreateToolTip clockForm.hWnd, "Click here to lock the clock in place on the desktop.", _
                     TTIconInfo, widgetName & " Locking Pin Help", , , , True
             Case "helpdropdown"
                If gblEnableBalloonTooltips = "True" Then CreateToolTip clockForm.hWnd, "Click here to lock the clock in place on the desktop.", _
                     TTIconInfo, widgetName & " Dropdown Pin Help", , , , True
                     
            Case "weekdaytill" ' this does the job for all the weekday tills underneath
                If gblEnableBalloonTooltips = "True" Then CreateToolTip clockForm.hWnd, "Click here to lower the weekday flag.", _
                     TTIconInfo, widgetName & " Weekday Indicator Help", , , , True
                ' set the weekday red cover to solid
                If clockForm.Widgets(fDayOfWeek).Widget.Alpha <> 0 Then
                    weekdayred.Alpha = Val(gblOpacity) / 100
                    weekdayred.Refresh
                End If
                With clockForm.Widgets("redalarmcover").Widget
                    .Alpha = 0
                    .Refresh
                End With
                
            ' monday - sunday ensuring the weekday till goes red even when hovering over the day text in=mages.
            Case "monday"
                If weekdayToggleEnabled = "true" Then
                    weekdayred.Alpha = Val(gblOpacity) / 100
                    weekdayred.Refresh
                End If
            Case "tuesday"
                If weekdayToggleEnabled = "true" Then
                    weekdayred.Alpha = Val(gblOpacity) / 100
                    weekdayred.Refresh
                End If
            Case "wednesday"
                If weekdayToggleEnabled = "true" Then
                    weekdayred.Alpha = Val(gblOpacity) / 100
                    weekdayred.Refresh
                End If
            Case "thursday"
                If weekdayToggleEnabled = "true" Then
                    weekdayred.Alpha = Val(gblOpacity) / 100
                    weekdayred.Refresh
                End If
            Case "friday"
                If weekdayToggleEnabled = "true" Then
                    weekdayred.Alpha = Val(gblOpacity) / 100
                    weekdayred.Refresh
                End If
            Case "saturday"
                If weekdayToggleEnabled = "true" Then
                    weekdayred.Alpha = Val(gblOpacity) / 100
                    weekdayred.Refresh
                End If
            Case "sunday"
                If weekdayToggleEnabled = "true" Then
                    weekdayred.Alpha = Val(gblOpacity) / 100
                    weekdayred.Refresh
                End If
                
            Case "redalarmcover"
                If gblEnableBalloonTooltips = "True" Then CreateToolTip clockForm.hWnd, "Click this alarm flag to delete this alarm.", _
                     TTIconInfo, widgetName & " Alarm Indicator Help", , , , True
                 weekdayred.Alpha = 0
                 weekdayred.Refresh
                If alarmtoggle1Enabled = True _
                    Or alarmtoggle2Enabled = True _
                    Or alarmtoggle3Enabled = True _
                    Or alarmtoggle4Enabled = True _
                    Or alarmtoggle5Enabled = True Then
                    With clockForm.Widgets("redalarmcover").Widget
                        .Alpha = Val(gblOpacity) / 100
                        .Refresh
                    End With
                End If
                
                
            Case "weekdayred"
                If gblEnableBalloonTooltips = "True" Then CreateToolTip clockForm.hWnd, "Click here to lower the weekday flag.", _
                     TTIconInfo, widgetName & " Weekday Indicator Help", , , , True
            Case Else
        
        End Select
    End If
    
    If EventName = "W_MouseLeave" Then
        Select Case LCase$(sender.Widget.Key)
        Case "weekdayred"
            weekdayred.Alpha = 0
            weekdayred.Refresh

        Case "redalarmcover"
             With clockForm.Widgets("redalarmcover").Widget
                .Alpha = 0
                .Refresh
            End With
        Case Else
        
        End Select
    End If
    
    If EventName = "W_MouseDown" Then
        Select Case LCase$(sender.Widget.Key)
            Case "helpdropdown"
                Call helpdropdown_myMouseDown

            Case "pendulumtransparent"
                Call pendulumtransparent_myMouseDown
                
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
'            Case "pendulum"
'                Call pendulum_myMouseDown
            Case "mutetoggle"
                Call mutetoggle_myMouseDown
            Case "settoggle"
                Call settoggle_myMouseDown
            Case "weekdaytoggle"
                Call weekdayToggle_myMouseDown
            Case "helptoggle"
                Call helptoggle_myMouseDown
'            Case "help1toggle" ' help3toggle is the topmost so it receives all the clicks
'                Call help1toggle_myMouseDown
'            Case "help2toggle" ' help3toggle is the topmost so it receives all the clicks
'                Call help2toggle_myMouseDown
             Case "help3toggle"
                Call help3toggle_myMouseDown
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
                Call weekdayred_myMouseDown
            Case "tuesday"
                Call weekdayred_myMouseDown
            Case "wednesday"
                Call weekdayred_myMouseDown
            Case "thursday"
                Call weekdayred_myMouseDown
            Case "friday"
                Call weekdayred_myMouseDown
            Case "saturday"
                Call weekdayred_myMouseDown
            Case "sunday"
                Call weekdayred_myMouseDown
            Case "dropdown"
                Call dropdown_myMouseDown
            Case "bellset"
                Call bellset_myMouseDown
            Case "sliderset"
                Call sliderset_myMouseDown
            Case "redalarmcover"
                Call redalarmcover_myMouseDown
            Case Else
                Debug.Print sender.Widget.Key
        End Select
    End If
    
    If EventName = "W_DblClick" Then
        Select Case LCase$(sender.Widget.Key)
            Case "lblterminaltext"
                Call displayscreen_myDblClick
            Case "displayscreen"
                Call displayscreen_myDblClick
            Case "backscreen"
                Call displayscreen_myDblClick
        End Select
    End If
    
    If EventName = "W_MouseUp" Then
        Select Case LCase$(sender.Widget.Key)
            Case "bell"
                'Call bell_mouseUp(Sender)
            Case "sliderset"
    
'                fileToPlay = "zzzz.wav"
'                If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
'                    PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
'                End If
                
               overlayWidget.Locked = False
                If SliderFreed = True Then
                    fClock.timeShiftValue = 0.5
                    overlayWidget.tmrTimeShift.Enabled = False
                End If

                tmrZZZZSound.Enabled = True
            Case "sliderset"
                pvtSlidersetMousedown = False ' never called!
                
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
' Procedure : displayscreen_myDblClick
' Author    : beededea
' Date      : 27/10/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub displayscreen_myDblClick()
    Dim fileToPlay As String: fileToPlay = vbNullString
   
    On Error GoTo displayscreen_myDblClick_Error

    If gblWidgetFunctions = "0" Or gblIgnoreMouse = "1" Then Exit Sub
    
    Call screenWrite(".")
    
    If gblVolumeBoost = "1" Then
        fileToPlay = "creturn.wav"
    Else
        fileToPlay = "creturn-quiet.wav"
    End If

    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If

   On Error GoTo 0
   Exit Sub

displayscreen_myDblClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure displayscreen_myDblClick of Class Module cfClock"
                
End Sub
'---------------------------------------------------------------------------------------
' Procedure : redalarmcover_myMouseDown
' Author    : beededea
' Date      : 03/08/2023

' Purpose   :  - note: - can't use a std MouseDown as
'             declaring a widget as cWidgetBase 'WithEvents' does not provide a std MouseDown with a sender object
'---------------------------------------------------------------------------------------
'
Private Sub redalarmcover_myMouseDown()
    Dim fileToPlay As String: fileToPlay = vbNullString
    Dim anAlarmDeleted As Boolean: anAlarmDeleted = False
   
    On Error GoTo redalarmcover_myMouseDown_Error
    
    If gblWidgetFunctions = "0" Or gblIgnoreMouse = "1" Then Exit Sub
    
    ' first thing to do is to stop the alarm timer ringing the bell, if running
    gblAlarmFlgRaised = False
    overlayWidget.tmrAlarmRinging.Enabled = False
       
    If fClock.clockForm.Widgets("redalarmcover").Widget.Alpha = 0 Then Exit Sub ' transparent widgets still respond to click events
    
    If gblVolumeBoost = "1" Then
        fileToPlay = "clunk.wav"
    Else
        fileToPlay = "clunk-quiet.wav"
    End If

    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    If gblRaisedAlarm = 1 And gblAlarm1Date <> "Alarm not yet set" Then
        Call deleteAlarm1(anAlarmDeleted)
       
    ElseIf gblRaisedAlarm = 2 And gblAlarm2Date <> "Alarm not yet set" Then
        Call deleteAlarm2(anAlarmDeleted)
    ElseIf gblRaisedAlarm = 3 And gblAlarm3Date <> "Alarm not yet set" Then
        Call deleteAlarm3(anAlarmDeleted)
    ElseIf gblRaisedAlarm = 4 And gblAlarm4Date <> "Alarm not yet set" Then
        Call deleteAlarm4(anAlarmDeleted)
    ElseIf gblRaisedAlarm = 5 And gblAlarm5Date <> "Alarm not yet set" Then
        Call deleteAlarm5(anAlarmDeleted)
    Else
        If alarmtoggle1Enabled = True Then alarmtoggle1Enabled = False
        If alarmtoggle2Enabled = True Then alarmtoggle2Enabled = False
        If alarmtoggle3Enabled = True Then alarmtoggle3Enabled = False
        If alarmtoggle4Enabled = True Then alarmtoggle4Enabled = False
        If alarmtoggle5Enabled = True Then alarmtoggle5Enabled = False
    End If
        
    With clockForm.Widgets("redalarmcover").Widget
        .Alpha = 0
        .Refresh
    End With

    If anAlarmDeleted = True Then
        If gblVolumeBoost = "1" Then
            fileToPlay = "till.wav"
        Else
            fileToPlay = "till-quiet.wav"
        End If
    Else
        If gblVolumeBoost = "1" Then
            fileToPlay = "clunk.wav"
        Else
            fileToPlay = "clunk-quiet.wav"
        End If
    End If
    
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
        
    If alarmtoggleEnabled = True Then
        alarmtoggleEnabled = False
    End If
     
   On Error GoTo 0
   Exit Sub

redalarmcover_myMouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure redalarmcover_myMouseDown of Class Module cfClock"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : deleteAlarm1
' Author    : beededea
' Date      : 29/09/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub deleteAlarm1(ByRef anAlarmDeleted As Boolean)
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString

    ' msgbox for deletion of alarm
    On Error GoTo deleteAlarm1_Error
    
    anAlarmDeleted = False

    answer = vbYes
    answerMsg = " Delete this alarm?"
    answer = msgBoxA(answerMsg, vbExclamation + vbYesNo, "Alarm Deletion", False)

    If answer = vbYes Then
        ' read the date and time variables for alarms
        If alarmtoggle1Enabled = True And gblAlarm1Date <> "Alarm not yet set" Then
            gblAlarm1Date = "Alarm not yet set"
            gblAlarm1Time = ""
            gblAlarm1 = ""
            sPutINISetting "Software\SteampunkClockCalendar", "alarm1Date", gblAlarm1Date, gblSettingsFile
            sPutINISetting "Software\SteampunkClockCalendar", "alarm1Time", gblAlarm1Time, gblSettingsFile
            alarmtoggle1Enabled = False
            
            ' change the alarm toggle to blank toggle
            fClock.clockForm.Widgets("alarmtoggle1").Widget.ImageKey = "alarmtoggle1"
            If widgetPrefs.IsVisible = True Then
                widgetPrefs.txtAlarm1Date.Text = gblAlarm1Date
                widgetPrefs.txtAlarm1Time.Text = gblAlarm1Time
            End If
            
            anAlarmDeleted = True
            
            gblAlarm1FlgRaised = False
             
        End If
    End If

   On Error GoTo 0
   Exit Sub

deleteAlarm1_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure deleteAlarm1 of Class Module cfClock"
    End Sub
    
'---------------------------------------------------------------------------------------
' Procedure : deleteAlarm2
' Author    : beededea
' Date      : 29/09/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub deleteAlarm2(ByRef anAlarmDeleted As Boolean)
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString

    ' msgbox for deletion of alarm
    On Error GoTo deleteAlarm2_Error

    anAlarmDeleted = False
        
    answer = vbYes
    answerMsg = " Delete this alarm?"
    answer = msgBoxA(answerMsg, vbExclamation + vbYesNo, "Alarm Deletion", False)

    If answer = vbYes Then
        ' read the date and time variables for alarms
        If alarmtoggle2Enabled = True And gblAlarm2Date <> "Alarm not yet set" Then
            gblAlarm2Date = "Alarm not yet set"
            gblAlarm2Time = ""
            gblAlarm2 = ""
            sPutINISetting "Software\SteampunkClockCalendar", "alarm2Date", gblAlarm2Date, gblSettingsFile
            sPutINISetting "Software\SteampunkClockCalendar", "alarm2Time", gblAlarm2Time, gblSettingsFile
            alarmtoggle2Enabled = False
            fClock.clockForm.Widgets("alarmtoggle2").Widget.ImageKey = "alarmtoggle2"
            widgetPrefs.txtAlarm2Date.Text = gblAlarm2Date
            widgetPrefs.txtAlarm2Time.Text = gblAlarm2Time
           
            anAlarmDeleted = True
            gblAlarm2FlgRaised = False
        End If
    End If

   On Error GoTo 0
   Exit Sub

deleteAlarm2_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure deleteAlarm2 of Class Module cfClock"
    End Sub
    
'---------------------------------------------------------------------------------------
' Procedure : deleteAlarm3
' Author    : beededea
' Date      : 29/09/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub deleteAlarm3(ByRef anAlarmDeleted As Boolean)
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString

    ' msgbox for deletion of alarm
    On Error GoTo deleteAlarm3_Error

    anAlarmDeleted = False
        
    answer = vbYes
    answerMsg = " Delete this alarm?"
    answer = msgBoxA(answerMsg, vbExclamation + vbYesNo, "Alarm Deletion", False)

    If answer = vbYes Then
        ' read the date and time variables for alarms
        If alarmtoggle3Enabled = True And gblAlarm3Date <> "Alarm not yet set" Then
            gblAlarm3Date = "Alarm not yet set"
            gblAlarm3Time = ""
            gblAlarm3 = ""
            sPutINISetting "Software\SteampunkClockCalendar", "alarm3Date", gblAlarm3Date, gblSettingsFile
            sPutINISetting "Software\SteampunkClockCalendar", "alarm3Time", gblAlarm3Time, gblSettingsFile
            alarmtoggle3Enabled = False
            fClock.clockForm.Widgets("alarmtoggle3").Widget.ImageKey = "alarmtoggle3"
            widgetPrefs.txtAlarm3Date.Text = gblAlarm3Date
            widgetPrefs.txtAlarm3Time.Text = gblAlarm3Time
            
            anAlarmDeleted = True
            gblAlarm3FlgRaised = False
        End If
    End If

   On Error GoTo 0
   Exit Sub

deleteAlarm3_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure deleteAlarm3 of Class Module cfClock"
    End Sub
    
'---------------------------------------------------------------------------------------
' Procedure : deleteAlarm4
' Author    : beededea
' Date      : 29/09/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
    Private Sub deleteAlarm4(ByRef anAlarmDeleted As Boolean)
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString

    ' msgbox for deletion of alarm
    On Error GoTo deleteAlarm4_Error

    anAlarmDeleted = False
        
    answer = vbYes
    answerMsg = " Delete this alarm?"
    answer = msgBoxA(answerMsg, vbExclamation + vbYesNo, "Alarm Deletion", False)

    If answer = vbYes Then
        ' read the date and time variables for alarms
        If alarmtoggle4Enabled = True And gblAlarm4Date <> "Alarm not yet set" Then
            gblAlarm4Date = "Alarm not yet set"
            gblAlarm4Time = ""
            gblAlarm4 = ""
            sPutINISetting "Software\SteampunkClockCalendar", "alarm4Date", gblAlarm4Date, gblSettingsFile
            sPutINISetting "Software\SteampunkClockCalendar", "alarm4Time", gblAlarm4Time, gblSettingsFile
            alarmtoggle4Enabled = False
            fClock.clockForm.Widgets("alarmtoggle4").Widget.ImageKey = "alarmtoggle4"
            widgetPrefs.txtAlarm4Date.Text = gblAlarm4Date
            widgetPrefs.txtAlarm4Time.Text = gblAlarm4Time
            
            anAlarmDeleted = True
            gblAlarm4FlgRaised = False
        End If
    End If

   On Error GoTo 0
   Exit Sub

deleteAlarm4_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure deleteAlarm4 of Class Module cfClock"
    End Sub
    
'---------------------------------------------------------------------------------------
' Procedure : deleteAlarm5
' Author    : beededea
' Date      : 29/09/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub deleteAlarm5(ByRef anAlarmDeleted As Boolean)
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString

    ' msgbox for deletion of alarm
    On Error GoTo deleteAlarm5_Error

    anAlarmDeleted = False
        
    answer = vbYes
    answerMsg = " Delete this alarm?"
    answer = msgBoxA(answerMsg, vbExclamation + vbYesNo, "Alarm Deletion", False)

    If answer = vbYes Then
        ' read the date and time variables for alarms
        If alarmtoggle5Enabled = True And gblAlarm5Date <> "Alarm not yet set" Then
            gblAlarm5Date = "Alarm not yet set"
            gblAlarm5Time = ""
            gblAlarm5 = ""
            sPutINISetting "Software\SteampunkClockCalendar", "alarm5Date", gblAlarm5Date, gblSettingsFile
            sPutINISetting "Software\SteampunkClockCalendar", "alarm5Time", gblAlarm5Time, gblSettingsFile
            alarmtoggle5Enabled = False
            fClock.clockForm.Widgets("alarmtoggle5").Widget.ImageKey = "alarmtoggle5"
            widgetPrefs.txtAlarm5Date.Text = gblAlarm5Date
            widgetPrefs.txtAlarm5Time.Text = gblAlarm5Time

            anAlarmDeleted = True
            gblAlarm5FlgRaised = False
        End If
    End If

   On Error GoTo 0
   Exit Sub

deleteAlarm5_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure deleteAlarm5 of Class Module cfClock"
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
    
    If gblWidgetFunctions = "0" Or gblIgnoreMouse = "1" Then Exit Sub
    
    If gblVolumeBoost = "1" Then
        fileToPlay = "singleBell.wav"
    Else
        fileToPlay = "singleBell-quiet.wav"
    End If
    
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
    Dim alphaVal As Single: alphaVal = 0
    Dim thisDayOfWeek As String: thisDayOfWeek = vbNullString
    
    Dim dropdownTop As Long: dropdownTop = 0
    Dim dropdownLeft As Long: dropdownLeft = 0
    
    Static dropdownEnabled As Boolean ' not using a property for this as it is not a stored value for restart
    
    On Error GoTo dropdown_MouseDown_Error
    
     If gblWidgetFunctions = "0" Or gblIgnoreMouse = "1" Then Exit Sub
    
   dropdownTop = dropdown.Top
    dropdownLeft = dropdown.Left
            
    alphaVal = Val(gblOpacity) / 100
    
    If dropdownEnabled = False Then
        Call screenWrite("Dropdown information canvas rolling down")
        If gblVolumeBoost = "1" Then
            fileToPlay = "rollerblind.wav"
        Else
            fileToPlay = "rollerblind-quiet.wav"
        End If
        
        dropdown.Move dropdownLeft, dropdownTop + 175
        dropdownEnabled = True
    Else
        Call screenWrite("Dropdown information canvas rolling up")
        If gblVolumeBoost = "1" Then
            fileToPlay = "rollerblindup.wav"
        Else
            fileToPlay = "rollerblindup-quiet.wav"
        End If

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
    
    If gblWidgetFunctions = "0" Or gblIgnoreMouse = "1" Then Exit Sub
    
    gblShowAlarmDateTime = True ' var that allows the stored alarm date/time to be displayed

    If gblVolumeBoost = "1" Then
        fileToPlay = "till.wav"
    Else
        fileToPlay = "till-quiet.wav"
    End If
    
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
    
    If gblWidgetFunctions = "0" Or gblIgnoreMouse = "1" Then Exit Sub
    
    gblShowAlarmDateTime = True ' var that allows the stored alarm date/time to be displayed 'alarmtoggle2"
    
    If gblVolumeBoost = "1" Then
        fileToPlay = "till.wav"
    Else
        fileToPlay = "till-quiet.wav"
    End If
    
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
    
    If gblWidgetFunctions = "0" Or gblIgnoreMouse = "1" Then Exit Sub
    
    gblShowAlarmDateTime = True ' var that allows the stored alarm date/time to be displayed 'alarmtoggle3"
    
    If gblVolumeBoost = "1" Then
        fileToPlay = "till.wav"
    Else
        fileToPlay = "till-quiet.wav"
    End If
    
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
    
    If gblWidgetFunctions = "0" Or gblIgnoreMouse = "1" Then Exit Sub
    
    gblShowAlarmDateTime = True ' var that allows the stored alarm date/time to be displayed 'alarmtoggle4"
    
    If gblVolumeBoost = "1" Then
        fileToPlay = "till.wav"
    Else
        fileToPlay = "till-quiet.wav"
    End If
    
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
    
    If gblWidgetFunctions = "0" Or gblIgnoreMouse = "1" Then Exit Sub
    
    gblShowAlarmDateTime = True ' var that allows the stored alarm date/time to be displayed 'alarmtoggle5"
    
    If gblVolumeBoost = "1" Then
        fileToPlay = "till.wav"
    Else
        fileToPlay = "till-quiet.wav"
    End If
    
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
' Procedure : weekdayred_myMouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :  - note: - can't use a std MouseDown as
'             declaring a widget as cWidgetBase 'WithEvents' does not provide a std MouseDown with a sender object
'---------------------------------------------------------------------------------------
'
Private Sub weekdayred_myMouseDown()

    Dim fileToPlay As String: fileToPlay = vbNullString
    Dim alphaVal As Single: alphaVal = 0
    Dim thisDayOfWeek As String: thisDayOfWeek = vbNullString
    
    Dim weekdayToggleTop As Long: weekdayToggleTop = 0
    Dim weekdayToggleLeft As Long: weekdayToggleLeft = 0
    
    On Error GoTo weekdayred_myMouseDown_Error
    
    If gblWidgetFunctions = "0" Or gblIgnoreMouse = "1" Then Exit Sub
    
    If fClock.clockForm.Widgets("weekdayred").Widget.Alpha = 0 Then Exit Sub
    
    weekdayToggleTop = weekdaytoggle.Top
    weekdayToggleLeft = weekdaytoggle.Left
    
    thisDayOfWeek = fDayOfWeek
        
    alphaVal = Val(gblOpacity) / 100
    
    If gblVolumeBoost = "1" Then
        fileToPlay = "till.wav"
    Else
        fileToPlay = "till-quiet.wav"
    End If
    
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    clockForm.Widgets(thisDayOfWeek).Widget.Alpha = 0
    weekdayToggleEnabled = False

    clockForm.Refresh

   On Error GoTo 0
   Exit Sub

weekdayred_myMouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure weekdayred_myMouseDown of Class Module cfClock"

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
    
    If gblWidgetFunctions = "0" Or gblIgnoreMouse = "1" Then Exit Sub
    
    If gblVolumeBoost = "1" Then
        fileToPlay = "chainpull.wav"
    Else
        fileToPlay = "chainpull-quiet.wav"
    End If

    If gblChainEnabled = "False" Then
        chainEnabled = True
    Else
        chainEnabled = False
    End If
         
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME
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
    
    If gblWidgetFunctions = "0" Or gblIgnoreMouse = "1" Then Exit Sub
    
    If gblVolumeBoost = "1" Then
        fileToPlay = "clunk.wav"
    Else
        fileToPlay = "clunk-quiet.wav"
    End If
     
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
    
    On Error GoTo alarmtoggle_myMouseDown_Error
    
    If gblWidgetFunctions = "0" Or gblIgnoreMouse = "1" Then Exit Sub
    
    pvtSlidersetMousedown = False

    If gblVolumeBoost = "1" Then
        fileToPlay = "clunk.wav"
    Else
        fileToPlay = "clunk-quiet.wav"
    End If
     
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME
    End If
        
    If alarmtoggleEnabled = False Then
        gblShowAlarmDateTime = False
        alarmtoggleEnabled = True
    Else
        gblShowAlarmDateTime = True ' var that allows the stored alarm date/time to be displayed
        alarmtoggleEnabled = False
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
    
    If gblWidgetFunctions = "0" Or gblIgnoreMouse = "1" Then Exit Sub
    
    If gblVolumeBoost = "1" Then
        fileToPlay = "clunk.wav"
    Else
        fileToPlay = "clunk-quiet.wav"
    End If
     
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

    If gblWidgetFunctions = "0" Or gblIgnoreMouse = "1" Then Exit Sub
    
    If gblVolumeBoost = "1" Then
        fileToPlay = "clunk.wav"
    Else
        fileToPlay = "clunk-quiet.wav"
    End If
     
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
    Dim setToggleIn As Long: setToggleIn = 0
    Dim setToggleOut As Long: setToggleOut = 0

    Dim fileToPlay As String: fileToPlay = vbNullString
    
    On Error GoTo settoggle_myMouseDown_Error
    
    If gblWidgetFunctions = "0" Or gblIgnoreMouse = "1" Then Exit Sub
    
    settoggleTop = settoggle.Top
    setToggleOut = 359
    setToggleIn = 363
    
    settoggle.Move setToggleIn, settoggleTop
    clockForm.Refresh
    
    Call screenWrite("setting any alarm")

    If gblVolumeBoost = "1" Then
        fileToPlay = "clunk.wav"
    Else
        fileToPlay = "clunk-quiet.wav"
    End If
     
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If

    Sleep 250
    
    settoggle.Move setToggleOut, settoggleTop
    
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
    
    pvtMuteInitiatingControl = "muteToggle"
     
    If gblWidgetFunctions = "0" Or gblIgnoreMouse = "1" Then Exit Sub
    
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

    If gblWidgetFunctions = "0" Or gblIgnoreMouse = "1" Then Exit Sub
    
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
    Dim alphaVal As Single: alphaVal = 0
    
    If gblWidgetFunctions = "0" Or gblIgnoreMouse = "1" Then Exit Sub
    
    displayScreenToggleTop = displayscreentoggle.Top
    displayScreenToggleLeft = displayscreentoggle.Left
    alphaVal = Val(gblOpacity) / 100

    If gblVolumeBoost = "1" Then
        fileToPlay = "clunk.wav"
    Else
        fileToPlay = "clunk-quiet.wav"
    End If
     
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
    Dim alphaVal As Single: alphaVal = 0
    Dim helpToggleFlg As Boolean: helpToggleFlg = False

    If gblWidgetFunctions = "0" Or gblIgnoreMouse = "1" Then Exit Sub
    
    helptoggleTop = helpToggle.Top
    helptoggleLeft = helpToggle.Left
    alphaVal = Val(gblOpacity) / 100
    
    If helpdropdown.Alpha = 0 Then
        If gblVolumeBoost = "1" Then
            fileToPlay = "rollerblinddown.wav"
        Else
            fileToPlay = "rollerblinddown-quiet.wav"
        End If
        helpToggle.Alpha = 0
        clockForm.Widgets("woodenbar").Widget.Alpha = alphaVal
        clockForm.Widgets("help1toggle").Widget.Alpha = alphaVal
        clockForm.Widgets("helpdropdown").Widget.Alpha = alphaVal
        clockForm.Widgets("txthelp").Widget.Alpha = alphaVal
        
        helpToggleFlg = True
    Else
        If gblVolumeBoost = "1" Then
            fileToPlay = "rollerblindup.wav"
        Else
            fileToPlay = "rollerblindup-quiet.wav"
        End If
    
        helpToggle.Alpha = alphaVal
        clockForm.Widgets("woodenbar").Widget.Alpha = 0
        clockForm.Widgets("help1toggle").Widget.Alpha = 0
        clockForm.Widgets("help2toggle").Widget.Alpha = 0
        clockForm.Widgets("help3toggle").Widget.Alpha = 0
        
        clockForm.Widgets("helpdropdown").Widget.Alpha = 0
        clockForm.Widgets("txthelp").Widget.Alpha = 0
        clockForm.Widgets("txtalarm").Widget.Alpha = 0
        clockForm.Widgets("txtlinks").Widget.Alpha = 0
        helpToggleFlg = False

    End If
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    Call screenWrite("Help Dropdown canvas No. 1 visible " & helpToggleFlg)
    clockForm.Refresh

   On Error GoTo 0
   Exit Sub

helptoggle_myMouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure helptoggle_myMouseDown of Class Module cfClock"

End Sub


''---------------------------------------------------------------------------------------
'' Procedure : help1toggle_myMouseDown
'' Author    : beededea
'' Date      : 03/08/2023
'' Purpose   : help3toggle is the topmost so it receives all the clicks
''---------------------------------------------------------------------------------------
''
'Private Sub help1toggle_myMouseDown()
'   On Error GoTo help1toggle_myMouseDown_Error
'
'    Dim fileToPlay As String: fileToPlay = vbNullString
'    Dim alphaVal As Single: alphaVal = 0
'    Dim helpToggleFlg As Boolean: helpToggleFlg = False
'
'    alphaVal = Val(gblOpacity) / 100
'
'    If help1toggle.Alpha > 0 Then
'        If gblVolumeBoost = "1" Then
'            fileToPlay = "rollerblinddown.wav"
'        Else
'            fileToPlay = "rollerblinddown-quiet.wav"
'        End If
'        helpToggleFlg = True
'
'        help1toggle.Alpha = 0
'        help2toggle.Alpha = alphaVal
'        clockForm.Widgets("txthelp").Widget.Alpha = 0
'        clockForm.Widgets("txtalarm").Widget.Alpha = alphaVal
'
'    End If
'    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
'        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
'    End If
'
'    Call screenWrite("Help Dropdown canvas No. 2 visible " & helpToggleFlg)
'    clockForm.Refresh
'
'   On Error GoTo 0
'   Exit Sub
'
'help1toggle_myMouseDown_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure help1toggle_myMouseDown of Class Module cfClock"
'
'End Sub
'
'
'
'
''---------------------------------------------------------------------------------------
'' Procedure : help2toggle_myMouseDown
'' Author    : beededea
'' Date      : 03/08/2023
'' Purpose   : help3toggle is the topmost so it receives all the clicks
''---------------------------------------------------------------------------------------
''
'Private Sub help2toggle_myMouseDown()
'   On Error GoTo help2toggle_myMouseDown_Error
'
'    Dim fileToPlay As String: fileToPlay = vbNullString
'    Dim alphaVal As Single: alphaVal = 0
'    Dim helpToggleFlg As Boolean: helpToggleFlg = False
'
'    alphaVal = Val(gblOpacity) / 100
'
'    If help2toggle.Alpha > 0 Then
'        If gblVolumeBoost = "1" Then
'            fileToPlay = "rollerblinddown.wav"
'        Else
'            fileToPlay = "rollerblinddown-quiet.wav"
'        End If
'        helpToggleFlg = True
'
'        help2toggle.Alpha = 0
'        help3toggle.Alpha = alphaVal
'        clockForm.Widgets("txtalarm").Widget.Alpha = 0
'        clockForm.Widgets("txtlinks").Widget.Alpha = alphaVal
'
'    End If
'    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
'        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
'    End If
'
'    Call screenWrite("Help Dropdown canvas No. 3 visible " & helpToggleFlg)
'    clockForm.Refresh
'
'   On Error GoTo 0
'   Exit Sub
'
'help2toggle_myMouseDown_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure help2toggle_myMouseDown of Class Module cfClock"
'
'End Sub



'---------------------------------------------------------------------------------------
' Procedure : help3toggle_myMouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   : help3toggle is the topmost so it receives all the clicks
'---------------------------------------------------------------------------------------
'
Private Sub help3toggle_myMouseDown()
   On Error GoTo help3toggle_myMouseDown_Error

    Dim fileToPlay As String: fileToPlay = vbNullString
    Dim alphaVal As Single: alphaVal = 0
    Dim helpToggleFlg As Boolean: helpToggleFlg = False

    If gblWidgetFunctions = "0" Or gblIgnoreMouse = "1" Then Exit Sub
    
    alphaVal = Val(gblOpacity) / 100
    
    If gblVolumeBoost = "1" Then
        fileToPlay = "rollerblinddown.wav"
    Else
        fileToPlay = "rollerblinddown-quiet.wav"
    End If
    helpToggleFlg = True

    If help1toggle.Alpha > 0 Then ' equates to a click on the 1 toggle
        help1toggle.Alpha = 0
        help2toggle.Alpha = alphaVal
        clockForm.Widgets("txthelp").Widget.Alpha = 0
        clockForm.Widgets("txtalarm").Widget.Alpha = alphaVal
        Call screenWrite("Help Dropdown canvas No. 2 visible " & helpToggleFlg)
    ElseIf help2toggle.Alpha > 0 Then ' equates to a click on the 1 toggle
        help2toggle.Alpha = 0
        help3toggle.Alpha = alphaVal
        clockForm.Widgets("txtalarm").Widget.Alpha = 0
        clockForm.Widgets("txtlinks").Widget.Alpha = alphaVal
        Call screenWrite("Help Dropdown canvas No. 3 visible " & helpToggleFlg)
    ElseIf help3toggle.Alpha > 0 Then
        help3toggle.Alpha = 0
        help1toggle.Alpha = alphaVal
        clockForm.Widgets("txtlinks").Widget.Alpha = 0
        clockForm.Widgets("txthelp").Widget.Alpha = alphaVal
        Call screenWrite("Help Dropdown canvas No. 1 visible " & helpToggleFlg)
    End If
    
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    clockForm.Refresh

   On Error GoTo 0
   Exit Sub

help3toggle_myMouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure help3toggle_myMouseDown of Class Module cfClock"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : pendulumtransparent_myMouseDown
' Author    : beededea
' Date      : 29/07/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub pendulumtransparent_myMouseDown()
    Dim fileToPlay As String: fileToPlay = vbNullString

    On Error GoTo pendulumtransparent_myMouseDown_Error
    
    If gblWidgetFunctions = "0" Or gblIgnoreMouse = "1" Then Exit Sub
    
    If gblVolumeBoost = "1" Then
        fileToPlay = "tick.wav"
    Else
        fileToPlay = "tick-quiet.wav"
    End If
    
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If

    If overlayWidget.SwingPendulum = False Then
        fClock.pendulumToggleEnabled = True
        overlayWidget.SwingPendulum = True
        gblTogglePendulum = "1"
    Else
        fClock.pendulumToggleEnabled = False
        overlayWidget.SwingPendulum = False
        gblTogglePendulum = "0"
    End If

   On Error GoTo 0
   Exit Sub

pendulumtransparent_myMouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure pendulumtransparent_myMouseDown of Class Module cfClock"

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
    Dim alphaVal As Single: alphaVal = 0

    If gblWidgetFunctions = "0" Or gblIgnoreMouse = "1" Then Exit Sub
    
    If helpdropdown.Alpha <> 0 Then
        Call screenWrite("Dropdown canvas rolling up")
        helpdropdownTop = helpdropdown.Top
        helpdropdownLeft = helpdropdown.Left
        alphaVal = Val(gblOpacity) / 100
    
        If gblVolumeBoost = "1" Then
            fileToPlay = "rollerblindup.wav"
        Else
            fileToPlay = "rollerblindup-quiet.wav"
        End If
        If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
            PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
        End If
        
        clockForm.Widgets("woodenbar").Widget.Alpha = 0
        clockForm.Widgets("help1toggle").Widget.Alpha = 0
        clockForm.Widgets("help2toggle").Widget.Alpha = 0
        clockForm.Widgets("help3toggle").Widget.Alpha = 0
        
        clockForm.Widgets("helpdropdown").Widget.Alpha = 0
        clockForm.Widgets("txthelp").Widget.Alpha = 0
        clockForm.Widgets("txtalarm").Widget.Alpha = 0
        clockForm.Widgets("txtlinks").Widget.Alpha = 0
        helpToggle.Alpha = alphaVal
        
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
     If gblWidgetFunctions = "0" Or gblIgnoreMouse = "1" Then Exit Sub
    
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
' Procedure : helpbottom_myMouseDown
' Author    : beededea
' Date      : 04/08/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub helpbottom_myMouseDown()
'    Dim helpbottomTop As Long: helpbottomTop = 0
'    Dim helpbottomLeft As Long: helpbottomLeft = 0
'    Dim labellayerTop As Long: labellayerTop = 0
'    Dim labellayerLeft As Long: labellayerLeft = 0
'    Dim alphaVal As single: alphaVal = 0

'    Dim fileToPlay As String: fileToPlay = vbNullString
    
    On Error GoTo helpbottom_myMouseDown_Error
'
'    helpbottomTop = helpbottom.Top
'    helpbottomLeft = helpbottom.Left
'
'    labellayerTop = labellayer.Top
'    labellayerLeft = labellayer.Left

     If gblWidgetFunctions = "0" Or gblIgnoreMouse = "1" Then Exit Sub
    
     If helpbottom.Alpha <> 0 Then ' cope with the click on the invisible helpbottom layer
        
        ShowHelp = False
                
        'If startupFlg = True Then
            'clockForm.Widgets("monday").Widget.Alpha = 0
            'f alarmtoggle1Enabled = False Then clockForm.Widgets("alarm1till").Widget.Alpha = 0
            'clockForm.Widgets("redarlarmcover").Widget.Alpha = 0
        'End If
        
'        fileToPlay = "pagefumble.wav"
'        If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
'            PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
'        End If
                
        'clockForm.Refresh
    End If
        
   On Error GoTo 0
   Exit Sub

helpbottom_myMouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure helpbottom_myMouseDown of Class Module cfClock"

    On Error GoTo 0
    Exit Sub

End Sub



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

    If gblWidgetFunctions = "0" Or gblIgnoreMouse = "1" Then Exit Sub
    
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
        Call ShellExecute(fClock.clockForm.hWnd, "Open", gblOpenFile, vbNullString, App.path, 1)
    Else
        Call ShellExecute(fClock.clockForm.hWnd, "runas", thisCommand, vbNullString, App.path, 1)
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


'---------------------------------------------------------------------------------------
' Procedure : clockForm_MouseUp
' Author    : beededea
' Date      : 05/05/2023
' Purpose   :    if this is a multi monitor system then:
'                we note the monitor ID at clockForm form_load and store as the clockFormMonitorID, for widgetPrefs widgetPrefsMonitorID
'                on clockFromMouseUp (drag complete) we sample the monitor ID again -only works, for the prefs form we have to have a 200ms timer that tests for form .left and .top x,y change (no MOVED event and mouse UP does not work on the form itself when the titlebar dragged)
'                if the monitor has changed then sample the physical monitor resolution
'                if the resolution is different then calculate new size proportion
'                alter the size of the form to correspond to the monitor
'---------------------------------------------------------------------------------------
'
Private Sub clockForm_MouseUp(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)

    On Error GoTo clockForm_MouseUp_Error
    
    frmTimer.tmrScreenResolution.Interval = 4500

    ' save the form x,y position whenever the globe/form is dragged
    Call savePosition

    Call positionClockByMonitorSize
    
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
      
      Debug.Print "FX " & FX
      Debug.Print "FY " & FY
      
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
    clockForm.Move clockForm.Left, clockForm.Top, pvtPSDWidth * clockForm.WidgetRoot.Zoom, pvtPSDHeight * clockForm.WidgetRoot.Zoom
    
    ' when resized from the gauge scroll up/down it needs to write it back so the size is correct when prefs opened.
    gblGaugeSize = CStr(NewZoom * 100)
    If widgetPrefs.IsVisible = True Then widgetPrefs.sliGaugeSize = Val(gblGaugeSize)
    
    If clockMonitorStruct.IsPrimary = True Then
        gblClockPrimaryHeightRatio = CStr(NewZoom)
    Else
        gblClockSecondaryHeightRatio = CStr(NewZoom)
    End If

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

    frmTimer.tmrScreenResolution.Enabled = False
    frmTimer.tmrScreenResolution.Interval = 200
    frmTimer.tmrScreenResolution.Enabled = True
    
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
    
    If gblWidgetFunctions = "0" Or gblIgnoreMouse = "1" Then Exit Sub
    
    overlayWidget.Locked = True

    pvtSlidersetMousedown = True
                
    If SliderFreed = False Then
        If gblVolumeBoost = "1" Then
            fileToPlay = "buzzer.wav"
        Else
            fileToPlay = "buzzer-quiet.wav"
        End If
        
        Call screenWrite("slider mechanism held, use A toggle to release")
        
        If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
            PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
        End If
    End If
     
   On Error GoTo 0
   Exit Sub

sliderset_myMouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliderset_myMouseDown of Class Module cfClock"

End Sub

' ----------------------------------------------------------------
' Procedure Name: sliderset_Moving
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 26/05/2024
' ----------------------------------------------------------------
Private Sub sliderset_Moving() ' this is an explicit Event-Handler derived from the RC Widget-instance with the property .moveable = true
    Dim a As Double
    
    Dim steps As Integer: steps = 10
    Dim min As Double: min = pvtSldXMin
    Dim max As Double: max = pvtSldXMax
    Dim value As Double: value = 0
    Dim I As Integer
    
'    Dim sliderwidth As Integer: sliderwidth = 0
'    Dim currpos As Integer: perc = 0
    Dim perc As Integer: perc = 0
    Dim timeDeviation As Currency: timeDeviation = 0
    
    On Error GoTo sliderset_Moving_Error
         
    pvtSldX = sliderset.Left
    If pvtSldX < pvtSldXMin Then pvtSldX = pvtSldXMin Else If pvtSldX > pvtSldXMax Then pvtSldX = pvtSldXMax 'keep pvtSldX "in range"
    
    If tmrZZZZSound.Enabled = False And pvtSlidersetMousedown = True Then tmrZZZZSound.Enabled = True
    
    timeShiftValue = (pvtSldX - pvtSldXMin) / (pvtSldXMax - pvtSldXMin) 'assign the new Value to the Public Prop, to trigger all the necessary refreshs
    
    ' if the slider is enabled and the timeShiftValue <> 0.5 then start the timer that shifts the digital clocks forwards or backwards
    If SliderFreed = True And timeShiftValue <> 0.5 Then
        overlayWidget.tmrTimeShift.Enabled = True
    Else
        overlayWidget.tmrTimeShift.Enabled = False
    End If
    
    ' Calculate the percentage
     perc = timeShiftValue * 100

'    ' Determine the percentage from the middle of the slider
    If perc <= 50 Then ' Negative value
        perc = perc - 50
        perc = perc * 2
        'perc = Math.Abs(perc)
    ElseIf perc > 50 Then ' Positive
        perc = perc - 50
        perc = perc * 2
    End If

'    ' Given a percentage calculates the logarithmic value
    timeDeviation = fLogSlider(perc)
    
    'create a property in cwOverlay and pass the time deviation to it.
    
    Debug.Print ("%KON-I-INFO, perc before " & perc)
    Debug.Print ("%KON-I-INFO, timeDeviation " & timeDeviation)
    
    On Error GoTo 0
    Exit Sub

sliderset_Moving_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliderset_Moving, line " & Erl & "."

End Sub

'---------------------------------------------------------------------------------------
' Procedure : LogSlider
' Author    : beededea
' Date      : 14/11/2024
' Purpose   : this function determines the logarithmic value of the time slider
'---------------------------------------------------------------------------------------
'
Function fLogSlider(thisPerc As Integer) As Currency

    Dim minp As Double: minp = 0
    Dim maxp As Double: maxp = 0
    Dim minv As Double: minv = 0
    Dim maxv As Double: maxv = 0
    Dim logScale As Double: logScale = 0
    Dim thisValue As Double: thisValue = 0
    
    On Error GoTo LogSlider_Error
    
    ' thisPerc will be between 0 and 100
    maxp = 100

    ' The result should be between 100 and 100000
    minv = Log(100)
    maxv = Log(50000#)

    ' calculate adjustment factor
    logScale = (maxv - minv) / (maxp - minp)
    thisValue = Exp(minv + logScale * (thisPerc - minp))
    'If (debugFlg = 1) Then Console.WriteLine("%KON-I-INFO, thisPerc " & thisPerc)

    fLogSlider = thisValue  ' return

   On Error GoTo 0
   Exit Function

LogSlider_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure LogSlider of Class Module cfClock"
End Function



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
' Procedure Name: timeShiftValue
' Purpose:
' Procedure Kind: Property (Get)
' Procedure Access: Public
' Return Type: Currency
' Author: beededea
' Date: 26/05/2024
' ----------------------------------------------------------------
Public Property Get timeShiftValue() As Currency
    
    On Error GoTo timeShiftValue_Error
    
    timeShiftValue = mtimeShiftValue
    
    On Error GoTo 0
    Exit Property

timeShiftValue_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure timeShiftValue, line " & Erl & "."

End Property

' ----------------------------------------------------------------
' Procedure Name: timeShiftValue
' Purpose:
' Procedure Kind: Property (Let)
' Procedure Access: Public
' Parameter RHS (Currency):
' Author: beededea
' Date: 26/05/2024
' ----------------------------------------------------------------
Public Property Let timeShiftValue(ByVal RHS As Currency)
    On Error GoTo timeShiftValue_Error
    
    If RHS < 0 Then RHS = 0 Else If RHS > 1 Then RHS = 1 'keep RHS "in range"
    pvtSldX = RHS * (pvtSldXMax - pvtSldXMin) + pvtSldXMin
    overlayWidget.sldCurWidth = pvtSldXMax - pvtSldX 'inform the Overlay about the new current "Cable-Width"
    
    sliderset.Move pvtSldX, pvtSldFixedTop 'move the sliderset-widget accordingly
    glow.Move pvtSldX + 29, pvtSldFixedTop - 13 'move the glow along with the sliderset-widget accordingly
    
    If mtimeShiftValue <> RHS Then mtimeShiftValue = RHS Else Exit Property  'only refresh, when there was a change to the previous value
    
    If Not Cairo.ToolTipForm Is Nothing Then Cairo.ToolTipForm.Visible = False
    clockForm.Refresh
    
    On Error GoTo 0
    Exit Property

timeShiftValue_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure timeShiftValue, line " & Erl & "."

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
    Dim alphaVal As Single: alphaVal = 0
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
        'fClock.clockForm.Widgets("lblTerminalText").Widget.Alpha = alphaVal
        'clockForm.Widgets("lblTerminalText").Widget.Alpha = alphaVal
    Else
        
        displayscreentoggle.Move displayScreenToggleOut, displayScreenToggleTop
        displayscreen.Alpha = 0
        screentop.Alpha = alphaVal
        timemachinetoggle.Alpha = 0
        backtoggle.Alpha = 0
        screenbackground.Alpha = 0
        gblDisplayScreenToggleEnabled = "False"
        'clockForm.Widgets("lblTerminalText").Widget.Alpha = alphaVal
        'fClock.clockForm.Widgets("lblTerminalText").Widget.Alpha = 0
    End If
    
    sPutINISetting "Software\SteampunkClockCalendar", "displayScreenToggleEnabled", gblDisplayScreenToggleEnabled, gblSettingsFile

    If mDisplayScreenToggleEnabled <> newValue Then mDisplayScreenToggleEnabled = newValue Else Exit Property
    displayScreenToggleEnabled = mDisplayScreenToggleEnabled
    
    Call screenWrite("raised display screen = " & mDisplayScreenToggleEnabled)
    
    clockForm.Refresh
     
    If gblVolumeBoost = "1" Then
        fileToPlay = "clunk.wav"
    Else
        fileToPlay = "clunk-quiet.wav"
    End If
     
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
    Dim fileToPlay As String: fileToPlay = vbNullString
    
    On Error GoTo clapperleft_myMouseDown_Error

    If gblWidgetFunctions = "0" Or gblIgnoreMouse = "1" Then Exit Sub
    
    pvtMuteInitiatingControl = "clapper"

    If gblVolumeBoost = "1" Then
        fileToPlay = "clunk.wav"
    Else
        fileToPlay = "clunk-quiet.wav"
    End If

    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
        
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
    
    If gblWidgetFunctions = "0" Or gblIgnoreMouse = "1" Then Exit Sub
    
    pvtMuteInitiatingControl = "clapper"

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
    Dim alphaVal As Single: alphaVal = 0

    On Error GoTo clapperEnabledLet_Error
    
    alphaVal = Val(gblOpacity) / 100
 
    If newValue = True Then
        gblEnableChimes = "1"
        gblClapperEnabled = "True"
        clapperright.Alpha = 0
        clapperLeft.Alpha = alphaVal
        widgetPrefs.chkEnableChimes.value = 1
        If pvtMuteInitiatingControl = "clapper" Then
                If gblVolumeBoost = "1" Then
                    fileToPlay = "singleBell.wav"
                Else
                    fileToPlay = "singleBell-quiet.wav"
                End If
                If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
                PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
            End If
        End If
    Else
        gblEnableChimes = "1"
        gblClapperEnabled = "False"
        'If pvtMuteInitiatingControl = "clapper" Then muteToggleEnabled = True
        clapperright.Alpha = alphaVal
        clapperLeft.Alpha = 0
        widgetPrefs.chkEnableChimes.value = 0
    End If
    
    sPutINISetting "Software\SteampunkClockCalendar", "clapperEnabled", gblClapperEnabled, gblSettingsFile

    If mClapperEnabled <> newValue Then mClapperEnabled = newValue Else Exit Property
    clapperEnabled = mClapperEnabled
    
    Call screenWrite("Clock chime enabled is set to " & mClapperEnabled)
    If mClapperEnabled = False Then Call screenWrite("All chimes and alarms will be disabled ")
    
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
    Dim alphaVal As Single: alphaVal = 0
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
    
    Call screenWrite("back screen raised = " & mBackToggleEnabled)

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

    If gblWidgetFunctions = "0" Or gblIgnoreMouse = "1" Then Exit Sub
    
    If gblVolumeBoost = "1" Then
        fileToPlay = "crank.wav"
    Else
        fileToPlay = "crank-quiet.wav"
    End If
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    crankedRaised = True

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
    
    If gblWidgetFunctions = "0" Or gblIgnoreMouse = "1" Then Exit Sub
    
    If gblVolumeBoost = "1" Then
        fileToPlay = "crank.wav"
    Else
        fileToPlay = "crank-quiet.wav"
    End If
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    crankedRaised = False

    clockForm.Refresh

   On Error GoTo 0
   Exit Sub

crankup_myMouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure crankup_myMouseDown of Class Module cfClock"

End Sub




'---------------------------------------------------------------------------------------
' Property  : crankedRaised
' Author    : beededea
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let crankedRaised(ByVal newValue As Boolean)
    Dim alphaVal As Single: alphaVal = 0
  
    On Error GoTo crankedRaisedLet_Error
    
    alphaVal = Val(gblOpacity) / 100
    
    If newValue = True Then
        crankup.Alpha = alphaVal
        crankdown.Alpha = 0
        gblCrankEnabled = "True"
        widgetPrefs.chkVolumeBoost.value = 1
        gblVolumeBoost = "1"
    Else
        gblCrankEnabled = "False"
        crankup.Alpha = 0
        crankdown.Alpha = alphaVal
        widgetPrefs.chkVolumeBoost.value = 0
        gblVolumeBoost = "0"
    End If
       
    sPutINISetting "Software\SteampunkClockCalendar", "crankEnabled", gblCrankEnabled, gblSettingsFile

    If mCrankedRaised <> newValue Then mCrankedRaised = newValue Else Exit Property
    crankedRaised = mCrankedRaised
    
    Call screenWrite("volume crank raised = " & mCrankedRaised)
    
    clockForm.Refresh
     
   On Error GoTo 0
   Exit Property

crankedRaisedLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property crankedRaised of Class Module cfClock"
End Property

'---------------------------------------------------------------------------------------
' Property  : crankedRaised
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get crankedRaised() As Boolean
   On Error GoTo crankedRaisedGet_Error

    crankedRaised = mCrankedRaised

   On Error GoTo 0
   Exit Property

crankedRaisedGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property crankedRaised of Class Module cfClock"
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
        widgetPrefs.chkEnableTicks.value = 1
    Else
        chain.Move chainLeft, chainTop - 10
        gblChainEnabled = "False"
        ticking = False
        widgetPrefs.chkEnableTicks.value = 0
    End If
       
    sPutINISetting "Software\SteampunkClockCalendar", "chainEnabled", gblChainEnabled, gblSettingsFile

    If mChainEnabled <> newValue Then mChainEnabled = newValue Else Exit Property
    chainEnabled = mChainEnabled
    Call screenWrite("ticking sound status = " & mChainEnabled)
    
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
    Dim alphaVal As Single: alphaVal = 0
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
        clockForm.Widgets("weekdaytill").Widget.Alpha = alphaVal
        clockForm.Widgets("weekdayshadow").Widget.Alpha = alphaVal
        clockForm.Widgets(thisDayOfWeek).Widget.Alpha = alphaVal
        gblWeekdayToggleEnabled = "True"
    Else
        weekdaytoggle.Move weekdayToggleOut, weekdayToggleTop
        clockForm.Widgets("weekdayred").Widget.Alpha = 0
        clockForm.Widgets("weekdaytill").Widget.Alpha = 0
        clockForm.Widgets("weekdayshadow").Widget.Alpha = 0
        clockForm.Widgets(thisDayOfWeek).Widget.Alpha = 0
        gblWeekdayToggleEnabled = "False"
    End If
    
    sPutINISetting "Software\SteampunkClockCalendar", "weekdayToggleEnabled", gblWeekdayToggleEnabled, gblSettingsFile

    If mWeekdayToggleEnabled <> newValue Then mWeekdayToggleEnabled = newValue Else Exit Property
    weekdayToggleEnabled = mWeekdayToggleEnabled
    
    Call screenWrite("weekday indicator raised = " & mWeekdayToggleEnabled)
    
    clockForm.Refresh
     
    If gblVolumeBoost = "1" Then
        fileToPlay = "till.wav"
    Else
        fileToPlay = "till-quiet.wav"
    End If
    
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
' Property  : alarmtoggleEnabled
' Author    : beededea
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let alarmtoggleEnabled(ByVal newValue As Boolean)
   
    Dim alarmtoggleTop As Long: alarmtoggleTop = 0
    Dim alarmtoggleLeft As Long: alarmtoggleLeft = 0
    Dim alarmToggleIn As Long: alarmToggleIn = 0
    Dim alarmToggleOut As Long: alarmToggleOut = 0
    Dim fileToPlay As String: fileToPlay = vbNullString
    
    On Error GoTo alarmtoggleEnabledLet_Error
    
    alarmtoggleTop = alarmtoggle.Top
    alarmToggleOut = 359
    alarmToggleIn = 363
    
    If alarmtoggleEnabled = True Then
        
        Call screenWrite("Entering Alarm Mode")
        Call screenWrite("Slider Mechanism Released, click bell to set")
       
        alarmtoggle.Move alarmToggleIn, alarmtoggleTop
        SliderFreed = True
        
        tmrZZZZSound.Enabled = True ' does a little buzz after 1/10th of a second, not cutting off the preceding 'clunk' prematurely
        
        ' set the slider position to the start point
        fClock.timeShiftValue = 0.5
        
        ' if an alarm is not selected then choose the first available alarm
        '     raise the alarm till and alarm toggle
        ' if an alarm is already raised then we edit that

        If alarmtoggle1Enabled = False _
            And alarmtoggle2Enabled = False _
            And alarmtoggle3Enabled = False _
            And alarmtoggle4Enabled = False _
            And alarmtoggle5Enabled = False Then
                Sleep 500
                If gblAlarm1Date = "Alarm not yet set" Then
                    alarmtoggle1Enabled = True
                ElseIf gblAlarm2Date = "Alarm not yet set" Then
                    alarmtoggle2Enabled = True
                ElseIf gblAlarm3Date = "Alarm not yet set" Then
                    alarmtoggle3Enabled = True
                ElseIf gblAlarm4Date = "Alarm not yet set" Then
                    alarmtoggle4Enabled = True
                ElseIf gblAlarm5Date = "Alarm not yet set" Then
                    alarmtoggle5Enabled = True
                Else
                    alarmtoggle1Enabled = True
                End If
        End If
        
        If weekdayToggleEnabled = True Then
            weekdayToggleEnabled = False ' makes a till sound
        Else
            tmrTILLSound.Enabled = True ' does a little buzz after 9/10th of a second, not cutting off the preceding 'zzzz' prematurely
        End If
        
    Else
        Call screenWrite("Cancelling Alarm Edit Mode")
        Call screenWrite("Slider Mechanism Placed in Rest Position")
        
        alarmtoggle.Move alarmToggleOut, alarmtoggleTop
        SliderFreed = False
        fClock.timeShiftValue = 0
        
        alarmtoggle1Enabled = False
        alarmtoggle2Enabled = False
        alarmtoggle3Enabled = False
        alarmtoggle4Enabled = False
        alarmtoggle5Enabled = False

    End If
        
    If mAlarmtoggleEnabled <> newValue Then mAlarmtoggleEnabled = newValue Else Exit Property
    alarmtoggleEnabled = mAlarmtoggleEnabled
    
    clockForm.Refresh
      
   On Error GoTo 0
   Exit Property

alarmtoggleEnabledLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property alarmtoggleEnabled of Class Module cfClock"
End Property

'---------------------------------------------------------------------------------------
' Property  : alarmtoggleEnabled
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get alarmtoggleEnabled() As Boolean
   On Error GoTo alarmtoggleEnabledGet_Error

    alarmtoggleEnabled = mAlarmtoggleEnabled

   On Error GoTo 0
   Exit Property

alarmtoggleEnabledGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property alarmtoggleEnabled of Class Module cfClock"
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
    
    Call screenWrite("time display raised = " & mTimeMachineToggleEnabled)
    
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
    
    If gblVolumeBoost = "1" Then
        fileToPlay = "clunk.wav"
    Else
        fileToPlay = "clunk-quiet.wav"
    End If
         
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If

    If newValue = True Then
        mutetoggle.Move muteToggleIn, muteToggleTop
        gblMuteToggleEnabled = "True"
        ticking = False
        gblEnableSounds = "0"
        widgetPrefs.chkEnableSounds.value = 0
        
        ' stop MCI from playing
        If gblVolumeBoost = "1" Then
            fileToPlay = "ticktock.wav"
        Else
            fileToPlay = "ticktock-quiet.wav"
        End If
        
        fileToPlay = App.path & "\resources\sounds\" & fileToPlay
        If fFExists(fileToPlay) Then
            lngReturn = mciSendString("stop " & fileToPlay, "", 0, 0)
        End If
    Else
        mutetoggle.Move muteToggleOut, muteToggleTop
        gblMuteToggleEnabled = "False"
        gblEnableSounds = "1"
        If ticking = True Then ticking = True
        widgetPrefs.chkEnableSounds.value = 1
    End If
    
    sPutINISetting "Software\SteampunkClockCalendar", "muteToggleEnabled", gblMuteToggleEnabled, gblSettingsFile

    If mMuteToggleEnabled <> newValue Then mMuteToggleEnabled = newValue Else Exit Property
    muteToggleEnabled = mMuteToggleEnabled
    
    Call screenWrite("sound mute is set to " & gblMuteToggleEnabled)

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
        overlayWidget.SwingPendulum = True
        gblPendulumToggleEnabled = "True"
        widgetPrefs.chkTogglePendulum.value = 1
    Else
        pendulumtoggle.Move pendulumToggleOut, pendulumToggleTop
        overlayWidget.SwingPendulum = False
        gblPendulumToggleEnabled = "False"
        widgetPrefs.chkTogglePendulum.value = 0
    End If
        
    If gblVolumeBoost = "1" Then
        fileToPlay = "clunk.wav"
    Else
        fileToPlay = "clunk-quiet.wav"
    End If
     
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    sPutINISetting "Software\SteampunkClockCalendar", "pendulumToggleEnabled", gblPendulumToggleEnabled, gblSettingsFile

    If mPendulumToggleEnabled <> newValue Then mPendulumToggleEnabled = newValue Else Exit Property
    pendulumToggleEnabled = mPendulumToggleEnabled
    
    Call screenWrite("pendulum swing " & mPendulumToggleEnabled)
    
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
    Dim alphaVal As Single: alphaVal = 0

    On Error GoTo alarmtoggle1EnabledLet_Error
    
    alarmtoggle1Top = alarmtoggle1.Top
    alarmtoggle1Out = 664
    alarmtoggle1In = 654
    
    pvtAlarmInitiatingControl = "alarmtoggle1"
    
    alphaVal = Val(gblOpacity) / 100
 
    ' NOTE: the code for displaying the alarm value when alarmtoggle1Enabled = true, is contained within the
    '       subroutine populateDigitalClockCache, this checks the current state of the alarm toggle and shows the alarm.
    
    If newValue = True Then
        gblAlarmToggle1Enabled = "True"
        gblRaisedAlarm = 1
        alarmtoggle1.Move alarmtoggle1Out, alarmtoggle1Top
        Call hideAlarmTills
        fClock.clockForm.Widgets("alarm1till").Widget.Alpha = alphaVal
    Else
        gblRaisedAlarm = 0
        gblAlarmToggle1Enabled = "False"
        alarmtoggle1.Move alarmtoggle1In, alarmtoggle1Top
        Call hideAlarmTills
        fClock.clockForm.Widgets("alarm1till").Widget.Alpha = 0
        fClock.clockForm.Widgets("redalarmcover").Widget.Alpha = 0
    End If
    
    sPutINISetting "Software\SteampunkClockCalendar", "alarmToggle1Enabled", gblAlarmToggle1Enabled, gblSettingsFile

    If mAlarmToggle1Enabled <> newValue Then mAlarmToggle1Enabled = newValue Else Exit Property
    alarmtoggle1Enabled = mAlarmToggle1Enabled
    
    If mAlarmToggle1Enabled = True Then
        Call screenWrite("Viewing Alarm - alarmtoggle1")
    Else
        Call screenWrite("Cancelling Alarm View Mode")
    End If
    
    triggerDigitalClockPopulation = True
    overlayWidget.tmrDigitRotator.Enabled = True
    
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
    Dim alphaVal As Single: alphaVal = 0
   
    On Error GoTo alarmtoggle2EnabledLet_Error
    
    alarmtoggle2Top = alarmtoggle2.Top
    alarmtoggle2Out = 678
    alarmtoggle2In = 670

    pvtAlarmInitiatingControl = "alarmtoggle2"
        
    alphaVal = Val(gblOpacity) / 100
 
    ' NOTE: the code for displaying the alarm value when alarmtoggle1Enabled = true, is contained within the
    '       subroutine populateDigitalClockCache, this checks the current state if the alarm toggle and shows the alarm.
    
    If newValue = True Then
        gblAlarmToggle2Enabled = "True"
        gblRaisedAlarm = 2
        alarmtoggle2.Move alarmtoggle2Out, alarmtoggle2Top
        Call hideAlarmTills
        fClock.clockForm.Widgets("alarm2till").Widget.Alpha = alphaVal
    Else
        gblAlarmToggle2Enabled = "False"
        gblRaisedAlarm = 0
        alarmtoggle2.Move alarmtoggle2In, alarmtoggle2Top
        Call hideAlarmTills
        fClock.clockForm.Widgets("alarm2till").Widget.Alpha = 0
        fClock.clockForm.Widgets("redalarmcover").Widget.Alpha = 0
    End If
    
    sPutINISetting "Software\SteampunkClockCalendar", "alarmtoggle2Enabled", gblAlarmToggle2Enabled, gblSettingsFile

    If mAlarmToggle2Enabled <> newValue Then
        mAlarmToggle2Enabled = newValue
        alarmtoggle2Enabled = mAlarmToggle2Enabled
    Else
        Exit Property
    End If
    
        
    If mAlarmToggle2Enabled = True Then
        Call screenWrite("Viewing Alarm - alarmtoggle2")
    Else
        Call screenWrite("Cancelling Alarm View Mode")
    End If
    
    triggerDigitalClockPopulation = True
    overlayWidget.tmrDigitRotator.Enabled = True
    
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
    Dim alphaVal As Single: alphaVal = 0
   
    On Error GoTo alarmtoggle3EnabledLet_Error
    
    alarmtoggle3Top = alarmtoggle3.Top
    alarmtoggle3Out = 690
    alarmtoggle3In = 682
        
    pvtAlarmInitiatingControl = "alarmtoggle3"
    
    alphaVal = Val(gblOpacity) / 100
 
    ' NOTE: the code for displaying the alarm value when alarmtoggle1Enabled = true, is contained within the
    '       subroutine populateDigitalClockCache, this checks the current state if the alarm toggle and shows the alarm.
    
    If newValue = True Then
        gblAlarmToggle3Enabled = "True"
        gblRaisedAlarm = 3
        alarmtoggle3.Move alarmtoggle3Out, alarmtoggle3Top
        Call hideAlarmTills
        fClock.clockForm.Widgets("alarm3till").Widget.Alpha = alphaVal
    Else
        gblAlarmToggle3Enabled = "False"
        gblRaisedAlarm = 0
        alarmtoggle3.Move alarmtoggle3In, alarmtoggle3Top
        Call hideAlarmTills
        fClock.clockForm.Widgets("alarm3till").Widget.Alpha = 0
        fClock.clockForm.Widgets("redalarmcover").Widget.Alpha = 0
    End If
    
    sPutINISetting "Software\SteampunkClockCalendar", "alarmtoggle3Enabled", gblAlarmToggle3Enabled, gblSettingsFile

    If mAlarmToggle3Enabled <> newValue Then mAlarmToggle3Enabled = newValue Else Exit Property
    alarmtoggle3Enabled = mAlarmToggle3Enabled
    
    If mAlarmToggle3Enabled = True Then
        Call screenWrite("Viewing Alarm - alarmtoggle3")
    Else
        Call screenWrite("Cancelling Alarm View Mode")
    End If
    
    triggerDigitalClockPopulation = True
    overlayWidget.tmrDigitRotator.Enabled = True
    
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
    Dim alphaVal As Single: alphaVal = 0
   
    On Error GoTo alarmtoggle4EnabledLet_Error
    
    alarmtoggle4Top = alarmtoggle4.Left
    alarmtoggle4Top = alarmtoggle4.Top
    alarmtoggle4Out = 690
    alarmtoggle4In = 682

     pvtAlarmInitiatingControl = "alarmtoggle4"
    
    alphaVal = Val(gblOpacity) / 100
 
    ' NOTE: the code for displaying the alarm value when alarmtoggle1Enabled = true, is contained within the
    '       subroutine populateDigitalClockCache, this checks the current state if the alarm toggle and shows the alarm.
    
    If newValue = True Then
        gblAlarmToggle4Enabled = "True"
        gblRaisedAlarm = 4
        alarmtoggle4.Move alarmtoggle4Out, alarmtoggle4Top
        Call hideAlarmTills
        fClock.clockForm.Widgets("alarm4till").Widget.Alpha = alphaVal
    Else
        gblAlarmToggle4Enabled = "False"
        gblRaisedAlarm = 0
        alarmtoggle4.Move alarmtoggle4In, alarmtoggle4Top
        Call hideAlarmTills
        fClock.clockForm.Widgets("alarm4till").Widget.Alpha = 0
        fClock.clockForm.Widgets("redalarmcover").Widget.Alpha = 0
    End If
    
    sPutINISetting "Software\SteampunkClockCalendar", "alarmtoggle4Enabled", gblAlarmToggle4Enabled, gblSettingsFile

    If mAlarmToggle4Enabled <> newValue Then mAlarmToggle4Enabled = newValue Else Exit Property
    alarmtoggle4Enabled = mAlarmToggle4Enabled
        
    If mAlarmToggle4Enabled = True Then
        Call screenWrite("Viewing Alarm - alarmtoggle4")
    Else
        Call screenWrite("Cancelling Alarm View Mode")
    End If
    
    triggerDigitalClockPopulation = True
    overlayWidget.tmrDigitRotator.Enabled = True
    
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
    Dim alphaVal As Single: alphaVal = 0
   
    On Error GoTo alarmtoggle5EnabledLet_Error
    
    alarmtoggle5Top = alarmtoggle5.Left
    alarmtoggle5Top = alarmtoggle5.Top
    alarmtoggle5Out = 680
    alarmtoggle5In = 672
    
    pvtAlarmInitiatingControl = "alarmtoggle5"
        
    alphaVal = Val(gblOpacity) / 100
 
    ' NOTE: the code for displaying the alarm value when alarmtoggle1Enabled = true, is contained within the
    '       subroutine populateDigitalClockCache, this checks the current state if the alarm toggle and shows the alarm.
    
    If newValue = True Then
        gblAlarmToggle5Enabled = "True"
        gblRaisedAlarm = 5
        alarmtoggle5.Move alarmtoggle5Out, alarmtoggle5Top
        Call hideAlarmTills
        fClock.clockForm.Widgets("alarm5till").Widget.Alpha = alphaVal
    Else
        gblAlarmToggle5Enabled = "False"
        gblRaisedAlarm = 0
        alarmtoggle5.Move alarmtoggle5In, alarmtoggle5Top
        Call hideAlarmTills
        fClock.clockForm.Widgets("alarm5till").Widget.Alpha = 0
        fClock.clockForm.Widgets("redalarmcover").Widget.Alpha = 0
    End If
    
    sPutINISetting "Software\SteampunkClockCalendar", "alarmtoggle5Enabled", gblAlarmToggle5Enabled, gblSettingsFile

    If mAlarmToggle5Enabled <> newValue Then mAlarmToggle5Enabled = newValue Else Exit Property
    alarmtoggle5Enabled = mAlarmToggle5Enabled
    
    If mAlarmToggle5Enabled = True Then
        Call screenWrite("Viewing Alarm - alarmtoggle5")
    Else
        Call screenWrite("Cancelling Alarm View Mode")
    End If
    
    triggerDigitalClockPopulation = True
    overlayWidget.tmrDigitRotator.Enabled = True
    
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
' Procedure Name: hideAlarmTills
' Purpose:
' Procedure Kind: sub
' Procedure Access: Private
' Return Type: String
' Author: beededea
' Date: 17/06/2024
' ----------------------------------------------------------------
Private Sub hideAlarmTills()
    On Error GoTo fDayOfWeek_Error
            
    If pvtAlarmInitiatingControl <> "alarmtoggle1" And alarmtoggle1Enabled = True Then alarmtoggle1Enabled = False
    If pvtAlarmInitiatingControl <> "alarmtoggle2" And alarmtoggle2Enabled = True Then alarmtoggle2Enabled = False
    If pvtAlarmInitiatingControl <> "alarmtoggle3" And alarmtoggle3Enabled = True Then alarmtoggle3Enabled = False
    If pvtAlarmInitiatingControl <> "alarmtoggle4" And alarmtoggle4Enabled = True Then alarmtoggle4Enabled = False
    If pvtAlarmInitiatingControl <> "alarmtoggle5" And alarmtoggle5Enabled = True Then alarmtoggle5Enabled = False
    
    On Error GoTo 0
    Exit Sub

fDayOfWeek_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure hideAlarmTills, line " & Erl & "."

End Sub

' ----------------------------------------------------------------
' Procedure Name: tmrTILLSound_Timer
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 21/06/2024
' ----------------------------------------------------------------
Private Sub tmrTILLSound_Timer()
    Dim fileToPlay As String: fileToPlay = vbNullString
    
    On Error GoTo tmrTILLSound_Timer_Error
    
    If gblVolumeBoost = "1" Then
        fileToPlay = "till.wav"
    Else
        fileToPlay = "till-quiet.wav"
    End If
    
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    tmrTILLSound.Enabled = False
    
   On Error GoTo 0
   Exit Sub

tmrTILLSound_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrTILLSound_Timer of Class Module cfClock"
End Sub



' ----------------------------------------------------------------
' Procedure Name: tmrZZZZSound_Timer
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 21/06/2024
' ----------------------------------------------------------------
Private Sub tmrZZZZSound_Timer()
    Dim fileToPlay As String: fileToPlay = vbNullString
    
    On Error GoTo tmrZZZZSound_Timer_Error
    
    If gblVolumeBoost = "1" Then
        fileToPlay = "zzzz.wav"
    Else
        fileToPlay = "zzzz-quiet.wav"
    End If
    
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    tmrZZZZSound.Enabled = False
    
   On Error GoTo 0
   Exit Sub

tmrZZZZSound_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrZZZZSound_Timer of Class Module cfClock"
End Sub

' ----------------------------------------------------------------
' Procedure Name: tmrTickingSound_Timer
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 21/06/2024
' ----------------------------------------------------------------
Private Sub tmrTickingSound_Timer()
    Dim fileToPlay As String: fileToPlay = vbNullString
    Dim lngReturn As Long: lngReturn = 0
    
    On Error GoTo tmrTickingSound_Timer_Error
    
    tmrTickingSound.Enabled = False
    tmrTickingSound.Interval = 30000 ' 30 secs
    
    If gblVolumeBoost = "1" Then
        fileToPlay = "ticktock.wav"
    Else
        fileToPlay = "ticktock-quiet.wav"
    End If
        
    fileToPlay = App.path & "\resources\sounds\" & fileToPlay
    If gblEnableSounds = "1" And fFExists(fileToPlay) Then
    
        lngReturn = mciSendString("Play " & fileToPlay, "", 0, 0)
        'PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
        
    tmrTickingSound.Enabled = True
    
   On Error GoTo 0
   Exit Sub

tmrTickingSound_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrTickingSound_Timer of Class Module cfClock"
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
    Dim alphaVal As Single: alphaVal = 0
    Dim mciFileToPlay As String: mciFileToPlay = vbNullString
    Dim lngReturn As Long: lngReturn = 0
    
    On Error GoTo tickingLet_Error
    
    alphaVal = Val(gblOpacity) / 100
    
    If newValue = True Then
        tmrTickingSound.Interval = 300
        tmrTickingSound.Enabled = True
     Else
        tmrTickingSound.Enabled = False
        
        ' stop MCI from playing
        If gblVolumeBoost = "1" Then
            fileToPlay = "ticktock.wav"
        Else
            fileToPlay = "ticktock-quiet.wav"
        End If
        
        fileToPlay = App.path & "\resources\sounds\" & fileToPlay
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


'---------------------------------------------------------------------------------------
' Procedure : clearWeekdayTill
' Author    : beededea
' Date      : 04/08/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub clearWeekdayTill()
    On Error GoTo clearWeekdayTill_Error

    weekdayred.Alpha = 0 ' sometimes you can fool the mouseleave, this helps to recover the situation
    weekdayred.Refresh


    On Error GoTo 0
    Exit Sub

clearWeekdayTill_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure clearWeekdayTill of Class Module cfClock"
            Resume Next
          End If
    End With

End Sub

'Private Sub tmrDisplay_Timer()

'    Dim interimText As String
'    Dim a As Integer: a = 0
'
'    ' print the console text on the screen
'    If gblDisplayScreenToggleEnabled = "True" Then
'        For a = 14 To 0 Step -1
'            interimText = interimText & gblTerminalRows(a) & vbCrLf
'
''            CC.DrawText 423, position, 220, 20, gblTerminalRows(a), , , 2, True, , someOpacity
''            position = position - interval
'        Next a
'
''        CC.DrawText 423, 203, 200, 20, ".......", , , 2, True, , someOpacity
''        CC.DrawText 423, 211, 200, 20, ".......", , , 2, True, , someOpacity
''        CC.DrawText 423, 218, 200, 20, ".......", , , 2, True, , someOpacity
''        CC.DrawText 423, 227, 200, 20, ".......", , , 2, True, , someOpacity
'    End If
'
'    fClock.clockForm.Widgets("lblTerminalText").Caption = interimText

'End Sub


'---------------------------------------------------------------------------------------
' Procedure : IsVisible
' Author    : beededea
' Date      : 08/05/2023
' Purpose   : calling a manual property to a form allows external checks to the form to
'             determine whether it is loaded, without also activating the form automatically.
'---------------------------------------------------------------------------------------
'
Public Property Get IsVisible() As Boolean
    On Error GoTo IsVisible_Error

    'If gblPrefsLoadedFlg Then
        If fClock.clockForm.WindowState = vbNormal Then
            IsVisible = fClock.clockForm.Visible
        Else
            IsVisible = False
        End If
'    Else
'        IsVisible = False
'    End If

    On Error GoTo 0
    Exit Property

IsVisible_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsVisible of Form widgetPrefs"
            Resume Next
          End If
    End With
End Property
