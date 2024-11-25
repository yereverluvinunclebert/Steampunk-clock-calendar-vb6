Attribute VB_Name = "modMain"
'@IgnoreModule IntegerDataType, ModuleWithoutFolder
' clockForm_BubblingEvent ' leaving that here so I can copy/paste to find it

Option Explicit

'------------------------------------------------------ STARTS
' for SetWindowPos z-ordering
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const HWND_TOP As Long = 0 ' for SetWindowPos z-ordering
Public Const HWND_TOPMOST As Long = -1
Public Const HWND_BOTTOM As Long = 1
Private Const SWP_NOMOVE  As Long = &H2
Private Const SWP_NOSIZE  As Long = &H1
Public Const OnTopFlags  As Long = SWP_NOMOVE Or SWP_NOSIZE
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
' to set the full window Opacity
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const WS_EX_LAYERED  As Long = &H80000
Private Const GWL_EXSTYLE  As Long = (-20)
Private Const LWA_COLORKEY  As Long = &H1       'to transparent
Private Const LWA_ALPHA  As Long = &H2          'to semi transparent
'------------------------------------------------------ ENDS

Public fMain As New cfMain
Public aboutWidget As cwAbout
Public licenceWidget As cwLicence

Public revealWidgetTimerCount As Integer
 
Public fClock As New cfClock
Public overlayWidget As cwOverlay
Public widgetName As String
Private B() As Byte

'---------------------------------------------------------------------------------------
' Procedure : Main
' Author    : beededea
' Date      : 27/04/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Main()
   On Error GoTo Main_Error
    
   Call mainRoutine(False)

   On Error GoTo 0
   Exit Sub

Main_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Main of Module modMain"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : main_routine
' Author    : beededea
' Date      : 27/06/2023
' Purpose   : called by sub main() to allow this routine to be called elsewhere,
'             a reload for example.
'---------------------------------------------------------------------------------------
'
Public Sub mainRoutine(ByVal restart As Boolean)

    Dim extractCommand As String: extractCommand = vbNullString
    Dim thisPSDFullPath As String: thisPSDFullPath = vbNullString
    Dim licenceState As Integer: licenceState = 0
    Dim clockFormMonitorID As Long: clockFormMonitorID = 0

    On Error GoTo main_routine_Error
    
    ' initialise global vars
    Call initialiseGlobalVars
    
    #If TWINBASIC Then
        gblCodingEnvironment = "TwinBasic"
    #Else
        gblCodingEnvironment = "VB6"
    #End If
    
    'MsgBox gblCodingEnvironment
    
    widgetName = "Steampunk Clock Calendar"
    thisPSDFullPath = App.path & "\Res\Steampunk Clock Calendar.psd"
    
    Call screenWrite("Steampunk O/S ver 1.0 (clockwork 0.1 hz)")
    Call screenWrite("A " & gblCodingEnvironment & " (WoW64) and RC6 creation")
    Call screenWrite("Copyright 2024, BrickMoon Interplanetary Enterprises")
    
    menuForm.mnuAbout.Caption = "About Steampunk Clock Calendar Cairo " & gblCodingEnvironment & " widget"

    ' resolve VB6 sizing width bug
    Call determineScreenDimensions

    gblPrefsPrimaryHeightTwips = 1000 * screenTwipsPerPixelY
    startupFlg = True
    
    extractCommand = Command$ ' capture any parameter passed, remove if a soft reload
    If restart = True Then extractCommand = vbNullString
    
    'add Resources to the global ImageList
    Call addImagesToImageList
    
    ' check the Windows version
    classicThemeCapable = fTestClassicThemeCapable
  
    ' get this tool's entry in the trinkets settings file and assign the app.path
    Call getTrinketsFile
    
    ' get the location of this tool's settings file (appdata)
    Call getToolSettingsFile
    
    ' read the dock settings from the new configuration file
    Call readSettingsFile("Software\SteampunkClockCalendar", gblSettingsFile)
    
    ' validate the inputs of any data from the input settings file
    Call validateInputs
    
    If gbl24HourClockMode = "1" Then
        Call screenWrite("Running startup " & "24hr mode")
    Else
        Call screenWrite("Running startup " & "12hr mode")
    End If
    
    ' check first usage via licence acceptance value and then set initial DPI awareness
    licenceState = fLicenceState()
    If licenceState = 0 Then
        Call testDPIAndSetInitialAwareness ' determine High DPI awareness or not by default on first run
    Else
        Call setDPIaware ' determine the user settings for DPI awareness, for this program and all its forms.
    End If

    'load the collection for storing the overlay surfaces with its relevant keys direct from the PSD
    If restart = False Then Call loadExcludePathCollection ' no need to reload the collPSDNonUIElements layer name keys on a reload
    
    ' start the load of the PSD file using the RC6 PSD-Parser.instance
    Call fClock.InitFromPSD(thisPSDFullPath)  ' no optional close layer as 3rd param
            
    ' initialise and create the three main RC forms on the current display
    Call createRCFormsOnCurrentDisplay
    
    ' place the form at the saved location and configure all the form elements
    Call makeVisibleFormElements
        
    ' run the functions that are also called at reload time.
    Call adjustMainControls ' this needs to be here after the initialisation of the Cairo forms and widgets
    
    ' move/hide onto/from the main screen
    Call mainScreen
        
    ' if the program is run in unhide mode, write the settings and exit
    Call handleUnhideMode(extractCommand)
    
    ' if the parameter states re-open prefs then shows the prefs
    If extractCommand = "prefs" Then
        Call makeProgramPreferencesAvailable
        extractCommand = vbNullString
    End If
    
    'load the preferences form but don't yet show it, speeds up access to the prefs via the menu
    Load widgetPrefs
    
    'load the message form but don't yet show it, speeds up access to the message form when needed.
    Load frmMessage
    
    ' display licence screen on first usage
    Call showLicence(fLicenceState)
    
    ' make the prefs appear on the first time running
    Call checkFirstTime
 
    ' configure any global timers here
    Call configureTimers
    
    ' note the monitor primary at clockForm form_load and store oldClockFormMonitorPrimary
    clockMonitorStruct = cWidgetFormScreenProperties(fClock.clockForm, clockFormMonitorID)
    oldClockFormMonitorPrimary = clockMonitorStruct.IsPrimary

    startupFlg = False
        
    ' RC message pump will auto-exit when Cairo Forms > 0 so we run it only when 0, this prevents message interruption
    ' when running twice on reload.
    If Cairo.WidgetForms.Count = 0 Then Cairo.WidgetForms.EnterMessageLoop
     
   On Error GoTo 0
   Exit Sub

main_routine_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure main_routine of Module modMain at "
    
End Sub
 
'---------------------------------------------------------------------------------------
' Procedure : checkFirstTime
' Author    : beededea
' Date      : 12/05/2023
' Purpose   : check for first time running, first time run shows prefs
'---------------------------------------------------------------------------------------
'
Private Sub checkFirstTime()

   On Error GoTo checkFirstTime_Error

    If gblFirstTimeRun = "true" Then
        'MsgBox "checkFirstTime"

        Call makeProgramPreferencesAvailable
        gblFirstTimeRun = "false"
        sPutINISetting "Software\SteampunkClockCalendar", "firstTimeRun", gblFirstTimeRun, gblSettingsFile
    End If

   On Error GoTo 0
   Exit Sub

checkFirstTime_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkFirstTime of Module modMain"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : initialiseGlobalVars
' Author    : beededea
' Date      : 12/05/2023
' Purpose   : initialise global vars
'---------------------------------------------------------------------------------------
'
Private Sub initialiseGlobalVars()
      
    On Error GoTo initialiseGlobalVars_Error
    
    monitorCount = 0

    ' general
    gblStartup = vbNullString
    gblWidgetFunctions = vbNullString
    
    gblSmoothSecondHand = vbNullString

    ' config
    gblEnableTooltips = vbNullString
    gblEnableBalloonTooltips = vbNullString
    
    gblEnablePrefsBalloonTooltips = vbNullString
    gblEnablePrefsTooltips = vbNullString
    
    gblShowTaskbar = vbNullString
    gblShowHelp = vbNullString
    gblTogglePendulum = vbNullString
    gbl24HourClockMode = vbNullString
    
    gblDpiAwareness = vbNullString
    
    gblGaugeSize = vbNullString
    gblScrollWheelDirection = vbNullString
    
    ' position
    gblAspectHidden = vbNullString
    gblWidgetPosition = vbNullString
    gblWidgetLandscape = vbNullString
    gblWidgetPortrait = vbNullString
    gblLandscapeFormHoffset = vbNullString
    gblLandscapeFormVoffset = vbNullString
    gblPortraitHoffset = vbNullString
    gblPortraitYoffset = vbNullString
    gblvLocationPercPrefValue = vbNullString
    gblhLocationPercPrefValue = vbNullString
    
    ' sounds
    gblEnableSounds = vbNullString
    gblEnableTicks = vbNullString
    gblEnableChimes = vbNullString
    gblVolumeBoost = vbNullString
    
    
    ' development
    gblDebug = vbNullString
    gblDblClickCommand = vbNullString
    gblOpenFile = vbNullString
    gblDefaultEditor = vbNullString
         
    ' font
    gblClockFont = vbNullString
    gblPrefsFont = vbNullString
    gblPrefsFontSizeHighDPI = vbNullString
    gblPrefsFontSizeLowDPI = vbNullString
    gblPrefsFontItalics = vbNullString
    gblPrefsFontColour = vbNullString
    
    gblDisplayScreenFont = vbNullString
    gblDisplayScreenFontSize = vbNullString
    gblDisplayScreenFontItalics = vbNullString
    gblDisplayScreenFontColour = vbNullString
    
    
    ' window
    gblWindowLevel = vbNullString
    gblPreventDragging = vbNullString
    gblOpacity = vbNullString
    gblWidgetHidden = vbNullString
    gblHidingTime = vbNullString
    gblIgnoreMouse = vbNullString
    gblFirstTimeRun = vbNullString
    gblMultiMonitorResize = vbNullString
    
    
    ' general storage variables declared
    gblSettingsDir = vbNullString
    gblSettingsFile = vbNullString
    
    gblTrinketsDir = vbNullString
    gblTrinketsFile = vbNullString
    
    gblClockHighDpiXPos = vbNullString
    gblClockHighDpiYPos = vbNullString
    
    gblClockLowDpiXPos = vbNullString
    gblClockLowDpiYPos = vbNullString
    
    gblLastSelectedTab = vbNullString
    gblSkinTheme = vbNullString
    
    gblsetToggleEnabled = vbNullString
    gblMuteToggleEnabled = vbNullString
    gblPendulumToggleEnabled = vbNullString
    gblPendulumEnabled = vbNullString
    gblWeekdayToggleEnabled = vbNullString
    gblDisplayScreenToggleEnabled = vbNullString
    gblTimeMachineToggleEnabled = vbNullString
    gblBackToggleEnabled = vbNullString
    gblClapperEnabled = vbNullString
    gblChainEnabled = vbNullString
    gblCrankEnabled = vbNullString
    gblAlarmToggle1Enabled = vbNullString
    gblAlarmToggle2Enabled = vbNullString
    gblAlarmToggle3Enabled = vbNullString
    gblAlarmToggle4Enabled = vbNullString
    gblAlarmToggle5Enabled = vbNullString
    
    gblAlarm1 = vbNullString
    gblAlarm2 = vbNullString
    gblAlarm3 = vbNullString
    gblAlarm4 = vbNullString
    gblAlarm5 = vbNullString
    
    gblAlarm1Date = vbNullString
    gblAlarm2Date = vbNullString
    gblAlarm3Date = vbNullString
    gblAlarm4Date = vbNullString
    gblAlarm5Date = vbNullString
    
    gblAlarm1Time = vbNullString
    gblAlarm2Time = vbNullString
    gblAlarm3Time = vbNullString
    gblAlarm4Time = vbNullString
    gblAlarm5Time = vbNullString
    
    ' general variables declared
    'toolSettingsFile = vbNullString
    classicThemeCapable = False
    storeThemeColour = 0
    windowsVer = vbNullString
    
    ' vars to obtain correct screen width (to correct VB6 bug) STARTS
    screenTwipsPerPixelX = 0
    screenTwipsPerPixelY = 0
    physicalScreenWidthTwips = 0
    physicalScreenHeightTwips = 0
    physicalScreenHeightPixels = 0
    physicalScreenWidthPixels = 0
    
    virtualScreenHeightPixels = 0
    virtualScreenWidthPixels = 0
    
    oldPhysicalScreenHeightPixels = 0
    oldPhysicalScreenWidthPixels = 0
    
    gblPrefsPrimaryHeightTwips = vbNullString
    gblPrefsSecondaryHeightTwips = vbNullString
    gblClockPrimaryHeightRatio = vbNullString
    gblClockSecondaryHeightRatio = vbNullString
    
    gblMessageAHeightTwips = vbNullString
    gblMessageAWidthTwips = vbNullString
    
    ' key presses
    CTRL_1 = False
    SHIFT_1 = False
    
    ' other globals
    debugFlg = 0
    minutesToHide = 0
    aspectRatio = vbNullString
    revealWidgetTimerCount = 0
    oldgblSettingsModificationTime = #1/1/2000 12:00:00 PM#
    
    gblCodingEnvironment = vbNullString

   On Error GoTo 0
   Exit Sub

initialiseGlobalVars_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure initialiseGlobalVars of Module modMain"
    
End Sub

        
'---------------------------------------------------------------------------------------
' Procedure : addImagesToImageList
' Author    : beededea
' Date      : 27/04/2023
' Purpose   : add image Resources to the global ImageList
'---------------------------------------------------------------------------------------
'
Private Sub addImagesToImageList()
    'Dim useloop As Integer: useloop = 0
    
    On Error GoTo addImagesToImageList_Error

'    add Resources to the global ImageList that are not being pulled from the PSD directly
    
    Cairo.ImageList.AddImage "about", App.path & "\Resources\images\about.png"
    'Cairo.ImageList.AddImage "help", App.path & "\Resources\images\diesel-volume-help.png"
    Cairo.ImageList.AddImage "licence", App.path & "\Resources\images\frame.png"
    Cairo.ImageList.AddImage "frmIcon", App.path & "\Resources\images\Icon.png"
    
    ' prefs icons
    
    Cairo.ImageList.AddImage "about-icon-dark", App.path & "\Resources\images\about-icon-dark-1010.jpg"
    Cairo.ImageList.AddImage "about-icon-light", App.path & "\Resources\images\about-icon-light-1010.jpg"
    Cairo.ImageList.AddImage "config-icon-dark", App.path & "\Resources\images\config-icon-dark-1010.jpg"
    Cairo.ImageList.AddImage "config-icon-light", App.path & "\Resources\images\config-icon-light-1010.jpg"
    Cairo.ImageList.AddImage "development-icon-light", App.path & "\Resources\images\development-icon-light-1010.jpg"
    Cairo.ImageList.AddImage "development-icon-dark", App.path & "\Resources\images\development-icon-dark-1010.jpg"
    Cairo.ImageList.AddImage "general-icon-dark", App.path & "\Resources\images\general-icon-dark-1010.jpg"
    Cairo.ImageList.AddImage "general-icon-light", App.path & "\Resources\images\general-icon-light-1010.jpg"
    Cairo.ImageList.AddImage "sounds-icon-light", App.path & "\Resources\images\sounds-icon-light-1010.jpg"
    Cairo.ImageList.AddImage "sounds-icon-dark", App.path & "\Resources\images\sounds-icon-dark-1010.jpg"
    Cairo.ImageList.AddImage "windows-icon-light", App.path & "\Resources\images\windows-icon-light-1010.jpg"
    Cairo.ImageList.AddImage "windows-icon-dark", App.path & "\Resources\images\windows-icon-dark-1010.jpg"
    Cairo.ImageList.AddImage "font-icon-dark", App.path & "\Resources\images\font-icon-dark-1010.jpg"
    Cairo.ImageList.AddImage "font-icon-light", App.path & "\Resources\images\font-icon-light-1010.jpg"
    Cairo.ImageList.AddImage "position-icon-light", App.path & "\Resources\images\position-icon-light-1010.jpg"
    Cairo.ImageList.AddImage "position-icon-dark", App.path & "\Resources\images\position-icon-dark-1010.jpg"
    
    Cairo.ImageList.AddImage "general-icon-dark-clicked", App.path & "\Resources\images\general-icon-dark-600-clicked.jpg"
    Cairo.ImageList.AddImage "config-icon-dark-clicked", App.path & "\Resources\images\config-icon-dark-600-clicked.jpg"
    Cairo.ImageList.AddImage "font-icon-dark-clicked", App.path & "\Resources\images\font-icon-dark-600-clicked.jpg"
    Cairo.ImageList.AddImage "sounds-icon-dark-clicked", App.path & "\Resources\images\sounds-icon-dark-600-clicked.jpg"
    Cairo.ImageList.AddImage "position-icon-dark-clicked", App.path & "\Resources\images\position-icon-dark-600-clicked.jpg"
    Cairo.ImageList.AddImage "development-icon-dark-clicked", App.path & "\Resources\images\development-icon-dark-600-clicked.jpg"
    Cairo.ImageList.AddImage "windows-icon-dark-clicked", App.path & "\Resources\images\windows-icon-dark-600-clicked.jpg"
    Cairo.ImageList.AddImage "about-icon-dark-clicked", App.path & "\Resources\images\about-icon-dark-600-clicked.jpg"
    
    Cairo.ImageList.AddImage "general-icon-light-clicked", App.path & "\Resources\images\general-icon-light-600-clicked.jpg"
    Cairo.ImageList.AddImage "config-icon-light-clicked", App.path & "\Resources\images\config-icon-light-600-clicked.jpg"
    Cairo.ImageList.AddImage "font-icon-light-clicked", App.path & "\Resources\images\font-icon-light-600-clicked.jpg"
    Cairo.ImageList.AddImage "sounds-icon-light-clicked", App.path & "\Resources\images\sounds-icon-light-600-clicked.jpg"
    Cairo.ImageList.AddImage "position-icon-light-clicked", App.path & "\Resources\images\position-icon-light-600-clicked.jpg"
    Cairo.ImageList.AddImage "development-icon-light-clicked", App.path & "\Resources\images\development-icon-light-600-clicked.jpg"
    Cairo.ImageList.AddImage "windows-icon-light-clicked", App.path & "\Resources\images\windows-icon-light-600-clicked.jpg"
    Cairo.ImageList.AddImage "about-icon-light-clicked", App.path & "\Resources\images\about-icon-light-600-clicked.jpg"
    
    Cairo.ImageList.AddImage "num0large", App.path & "\Resources\images\num0large.png"
    Cairo.ImageList.AddImage "num1large", App.path & "\Resources\images\num1large.png"
    Cairo.ImageList.AddImage "num2large", App.path & "\Resources\images\num2large.png"
    Cairo.ImageList.AddImage "num3large", App.path & "\Resources\images\num3large.png"
    Cairo.ImageList.AddImage "num4large", App.path & "\Resources\images\num4large.png"
    Cairo.ImageList.AddImage "num5large", App.path & "\Resources\images\num5large.png"
    Cairo.ImageList.AddImage "num6large", App.path & "\Resources\images\num6large.png"
    Cairo.ImageList.AddImage "num7large", App.path & "\Resources\images\num7large.png"
    Cairo.ImageList.AddImage "num8large", App.path & "\Resources\images\num8large.png"
    Cairo.ImageList.AddImage "num9large", App.path & "\Resources\images\num9large.png"
    
    Cairo.ImageList.AddImage "jan", App.path & "\Resources\images\jan.png"
    Cairo.ImageList.AddImage "feb", App.path & "\Resources\images\feb.png"
    Cairo.ImageList.AddImage "mar", App.path & "\Resources\images\mar.png"
    Cairo.ImageList.AddImage "apr", App.path & "\Resources\images\apr.png"
    Cairo.ImageList.AddImage "may", App.path & "\Resources\images\may.png"
    Cairo.ImageList.AddImage "jun", App.path & "\Resources\images\jun.png"
    Cairo.ImageList.AddImage "jul", App.path & "\Resources\images\jul.png"
    Cairo.ImageList.AddImage "aug", App.path & "\Resources\images\aug.png"
    Cairo.ImageList.AddImage "sep", App.path & "\Resources\images\sep.png"
    Cairo.ImageList.AddImage "oct", App.path & "\Resources\images\oct.png"
    Cairo.ImageList.AddImage "nov", App.path & "\Resources\images\nov.png"
    Cairo.ImageList.AddImage "dec", App.path & "\Resources\images\dec.png"
    
    Cairo.ImageList.AddImage "small0", App.path & "\Resources\images\small0.png"
    Cairo.ImageList.AddImage "small1", App.path & "\Resources\images\small1.png"
    Cairo.ImageList.AddImage "small2", App.path & "\Resources\images\small2.png"
    Cairo.ImageList.AddImage "small3", App.path & "\Resources\images\small3.png"
    Cairo.ImageList.AddImage "small4", App.path & "\Resources\images\small4.png"
    Cairo.ImageList.AddImage "small5", App.path & "\Resources\images\small5.png"
    Cairo.ImageList.AddImage "small6", App.path & "\Resources\images\small6.png"
    Cairo.ImageList.AddImage "small7", App.path & "\Resources\images\small7.png"
    Cairo.ImageList.AddImage "small8", App.path & "\Resources\images\small8.png"
    Cairo.ImageList.AddImage "small9", App.path & "\Resources\images\small9.png"
    
    Cairo.ImageList.AddImage "AM", App.path & "\Resources\images\AM.png"
    Cairo.ImageList.AddImage "PM", App.path & "\Resources\images\PM.png"
    
    'Cairo.ImageList.AddImage "month", App.path & "\Resources\images\jan.png"
    
    Cairo.ImageList.AddImage "hourCache1", App.path & "\Resources\images\small0.png"
    Cairo.ImageList.AddImage "hourCache2", App.path & "\Resources\images\small0.png"
    Cairo.ImageList.AddImage "minuteCache1", App.path & "\Resources\images\small0.png"
    Cairo.ImageList.AddImage "minuteCache2", App.path & "\Resources\images\small0.png"
    Cairo.ImageList.AddImage "dayCache1", App.path & "\Resources\images\num0large.png"
    Cairo.ImageList.AddImage "dayCache2", App.path & "\Resources\images\num1large.png"
    Cairo.ImageList.AddImage "monthCache", App.path & "\Resources\images\jan.png"
    Cairo.ImageList.AddImage "yearCache1", App.path & "\Resources\images\num1large.png"
    Cairo.ImageList.AddImage "yearCache2", App.path & "\Resources\images\num9large.png"
    Cairo.ImageList.AddImage "yearCache3", App.path & "\Resources\images\num7large.png"
    Cairo.ImageList.AddImage "yearCache4", App.path & "\Resources\images\num0large.png"
    Cairo.ImageList.AddImage "AMPMCache", App.path & "\Resources\images\AM.png"
    
   On Error GoTo 0
   Exit Sub

addImagesToImageList_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure addImagesToImageList of Module modMain, an image has probably been accidentally deleted from the resources/images folder."

End Sub
'---------------------------------------------------------------------------------------
' Procedure : adjustMainControls
' Author    : beededea
' Date      : 27/04/2023
' Purpose   : called at runtime and on restart, sets the characteristics of the gauge, individual controls and menus
'---------------------------------------------------------------------------------------
'
Public Sub adjustMainControls()
   
   On Error GoTo adjustMainControls_Error

    ' validate the inputs of any data from the input settings file
    Call validateInputs
    
    ' initial call just to obtain initial physical screen monitor ID
    Call positionClockByMonitorSize
        
    ' set the initial size
    If monitorCount > 1 And (LTrim$(gblMultiMonitorResize) = "1" Or LTrim$(gblMultiMonitorResize) = "2") Then
        If clockMonitorStruct.IsPrimary = True Then
            Call fClock.AdjustZoom(Val(gblClockPrimaryHeightRatio))
        Else
            Call fClock.AdjustZoom(Val(gblClockSecondaryHeightRatio))
        End If
    Else
        fClock.AdjustZoom Val(gblGaugeSize) / 100
    End If
    
'    overlayWidget.ZoomDirection = gblScrollWheelDirection
    
    If gblWidgetFunctions = "1" Then
        menuForm.mnuSwitchOff.Checked = False
        menuForm.mnuTurnFunctionsOn.Checked = True
    Else
        menuForm.mnuSwitchOff.Checked = True
        menuForm.mnuTurnFunctionsOn.Checked = False
    End If
    
    If gblDefaultEditor <> vbNullString And gblDebug = "1" Then
        menuForm.mnuEditWidget.Caption = "Edit Widget using " & gblDefaultEditor
        menuForm.mnuEditWidget.Visible = True
    Else
        menuForm.mnuEditWidget.Visible = False
    End If
    
    If gblShowTaskbar = "0" Then
        fClock.clockForm.ShowInTaskbar = False
    Else
        fClock.clockForm.ShowInTaskbar = True
    End If

    
    ' set the visibility and characteristics of the interactive areas
    ' the alpha is already set to zero for all layers found in the PSD, we now turn them back on as we require
    
    With fClock.clockForm.Widgets("helpbottom").Widget
        .HoverColor = 0
        .MousePointer = IDC_HAND
        .Alpha = 0
        .Tag = 0.01
    End With
    
    With fClock.clockForm.Widgets("dropdown").Widget
        .HoverColor = 0
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.01
    End With

    With fClock.clockForm.Widgets("bottombox").Widget
        .HoverColor = 0
        .MousePointer = IDC_ARROW
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.01
    End With

'    With fClock.clockForm.Widgets("pendulum").Widget
'        .HoverColor = 0
'        .MousePointer = IDC_HAND
'        .Alpha = Val(gblOpacity) / 100
'        .Tag = 0.01
'    End With


    With fClock.clockForm.Widgets("heatercoil").Widget
        .HoverColor = 0
        .MousePointer = IDC_ARROW
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.01
    End With

'    With fClock.clockForm.Widgets("orangeheaterglow").Widget
'        .HoverColor = 0
'        .MousePointer = IDC_ARROW
'        .Alpha = Val(gblOpacity) / 100
'        .Tag = 0.01
'    End With
    
'    With fClock.clockForm.Widgets("meridien").Widget
'        .HoverColor = 0
'        .MousePointer = IDC_ARROW
'        .Alpha = Val(gblOpacity) / 100
'        .Tag = 0.01
'    End With
'
'    With fClock.clockForm.Widgets("anti").Widget
'        .HoverColor = 0
'        .MousePointer = IDC_ARROW
'        .Alpha = Val(gblOpacity) / 100
'        .Tag = 0.01
'    End With

    With fClock.clockForm.Widgets("backtoggle").Widget
        .HoverColor = 0
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.01
    End With

    With fClock.clockForm.Widgets("alarmtoggle3").Widget
        .HoverColor = 0
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.01
    End With

    With fClock.clockForm.Widgets("alarmtoggle2").Widget
        .HoverColor = 0
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.01
    End With

    With fClock.clockForm.Widgets("alarmtoggle1").Widget
        .HoverColor = 0
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.01
    End With

    With fClock.clockForm.Widgets("alarmtoggle4").Widget
        .HoverColor = 0
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.01
    End With
    
    With fClock.clockForm.Widgets("alarmtoggle5").Widget
        .HoverColor = 0
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.01
    End With

    With fClock.clockForm.Widgets("maincasingsurround").Widget
        .HoverColor = 0
        .MousePointer = IDC_ARROW
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.01
    End With

    With fClock.clockForm.Widgets("crankdown").Widget
        .HoverColor = 0
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.01
    End With

    With fClock.clockForm.Widgets("crankup").Widget
        .HoverColor = 0
        .MousePointer = IDC_HAND
        .Alpha = 0
        .Tag = 0.01
    End With

    With fClock.clockForm.Widgets("displayscreen").Widget
        .HoverColor = 0
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.01
    End With

    With fClock.clockForm.Widgets("alarm1till").Widget
        .HoverColor = 0
        .MousePointer = IDC_HAND
        .Alpha = 0
        .Tag = 0.01
    End With
    
    With fClock.clockForm.Widgets("redalarmcover").Widget
        .HoverColor = 0
        .MousePointer = IDC_HAND
        .Alpha = 0
        .Tag = 0.01
    End With

    With fClock.clockForm.Widgets("displayscreentoggle").Widget
        .HoverColor = 0
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.01
    End With

    With fClock.clockForm.Widgets("timemachinetoggle").Widget
        .HoverColor = 0
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.01
    End With
    
    With fClock.clockForm.Widgets("weekdaytoggle").Widget
        .HoverColor = 0
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.01
    End With
    
    
    With fClock.clockForm.Widgets("alarmtoggle").Widget
        .HoverColor = 0
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.01
    End With

    With fClock.clockForm.Widgets("helptoggle").Widget
        .HoverColor = 0
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.01
    End With
    

    With fClock.clockForm.Widgets("help1toggle").Widget
        .HoverColor = 0
        .MousePointer = IDC_HAND
        .Alpha = 0
        .Tag = 0.01
    End With
    
    With fClock.clockForm.Widgets("help2toggle").Widget
        .HoverColor = 0
        .MousePointer = IDC_HAND
        .Alpha = 0
        .Tag = 0.01
    End With
    
    
    
    With fClock.clockForm.Widgets("help3toggle").Widget
        .HoverColor = 0
        .MousePointer = IDC_HAND
        .Alpha = 0
        .Tag = 0.01
    End With
    
    With fClock.clockForm.Widgets("settoggle").Widget
        .HoverColor = 0
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.01
    End With
    
    With fClock.clockForm.Widgets("mutetoggle").Widget
        .HoverColor = 0
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.01
    End With
    
    With fClock.clockForm.Widgets("pendulumtoggle").Widget
        .HoverColor = 0
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.01
    End With

    With fClock.clockForm.Widgets("monday").Widget
        .HoverColor = 0
        .MousePointer = IDC_HAND
        .Alpha = 0
        .Tag = 0.01
    End With

    With fClock.clockForm.Widgets("topdigitalclock").Widget
        .HoverColor = 0
        .MousePointer = IDC_ARROW
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.01
    End With
    

    With fClock.clockForm.Widgets("cablewheel").Widget
        .HoverColor = 0
        .MousePointer = IDC_ARROW
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.01
    End With

    With fClock.clockForm.Widgets("clapperleft").Widget
        .HoverColor = 0
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.01
    End With

    With fClock.clockForm.Widgets("clapperright").Widget
        .HoverColor = 0
        .MousePointer = IDC_HAND
        .Alpha = 0
        .Tag = 0.01
    End With

    With fClock.clockForm.Widgets("bar").Widget
        .HoverColor = 0
        .MousePointer = IDC_ARROW
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.01
    End With

    With fClock.clockForm.Widgets("glow").Widget
        .HoverColor = 0
        .MousePointer = IDC_ARROW
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.01
    End With

    With fClock.clockForm.Widgets("sliderset").Widget
        .HoverColor = 0
        .MousePointer = IDC_SIZEALL
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.01
    End With

    With fClock.clockForm.Widgets("chain").Widget
        .HoverColor = 0
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.01
    End With

    With fClock.clockForm.Widgets("clockset").Widget
        .HoverColor = 0
        .MousePointer = IDC_ARROW
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.01
    End With

    With fClock.clockForm.Widgets("bellset").Widget
        .HoverColor = 0
        .MousePointer = IDC_ARROW
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.01
    End With

    With fClock.clockForm.Widgets("labellayer").Widget
        .HoverColor = 0
        .MousePointer = IDC_ARROW
        .Alpha = 0
        .Tag = 0.01
    End With

    With fClock.clockForm.Widgets("lockingpin").Widget
        .HoverColor = 0
        .MousePointer = IDC_ARROW
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.01
    End With

    With fClock.clockForm.Widgets("grommet").Widget
        .HoverColor = 0
        .MousePointer = IDC_ARROW
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.01
    End With

    With fClock.clockForm.Widgets("weekdaytill").Widget
        .HoverColor = 0
        .MousePointer = IDC_ARROW
        .Alpha = 0
        .Tag = 0.01
    End With

    With fClock.clockForm.Widgets("weekdayshadow").Widget
        .HoverColor = 0
        .MousePointer = IDC_ARROW
        .Alpha = 0
        .Tag = 0.01
    End With
'
'    With fClock.clockForm.Widgets("deletealarm").Widget
'        .HoverColor = 0
'        .MousePointer = IDC_ARROW
'        .Alpha = 0
'        .Tag = 0.01
'    End With
'
'    With fClock.clockForm.Widgets("tickdeletealarm").Widget
'        .HoverColor = 0
'        .MousePointer = IDC_ARROW
'        .Alpha = 0
'        .Tag = 0.01
'    End With

    With fClock.clockForm.Widgets("helpdropdown").Widget
        .HoverColor = 0
        .MousePointer = IDC_ARROW
        .Alpha = 0
        .Tag = 0.01
    End With

    With fClock.clockForm.Widgets("pendulumtransparent").Widget
        .HoverColor = 0
        .MousePointer = IDC_ARROW
        .Alpha = 0.01
        .Tag = 0.01
    End With
    
    ' set the lock state of the clock
    If gblPreventDragging = "0" Then
        menuForm.mnuLockWidget.Checked = False
        overlayWidget.Locked = False
        fClock.clockForm.Widgets("lockingpin").Widget.Alpha = Val(gblOpacity) / 100
        fClock.clockForm.Widgets("lockingpinlocked").Widget.Alpha = 0
    Else
        menuForm.mnuLockWidget.Checked = True
        overlayWidget.Locked = True ' this is just here for continuity's sake, it is also set at the time the control is selected
        fClock.clockForm.Widgets("lockingpin").Widget.Alpha = 0
        fClock.clockForm.Widgets("lockingpinlocked").Widget.Alpha = Val(gblOpacity) / 100
    End If
    
    ' ensure the background help displays on startup
    If gblShowHelp = "0" Then
        fClock.ShowHelp = False
    Else
        fClock.ShowHelp = True
    End If
    
    ' ensure the background help displays on startup
    If gblTogglePendulum = "0" Then
        overlayWidget.SwingPendulum = False
    Else
        overlayWidget.SwingPendulum = True
    End If
    
'    If gbl24HourClockMode = "1" Then
'
'    End If

    ' set the slider position to the start point
    fClock.timeShiftValue = 0
    
    ' set the position of the various UI toggles at startup
    If gblsetToggleEnabled = "True" Then
        fClock.settoggleEnabled = True
    End If
    
    If gblMuteToggleEnabled = "True" Then
        gblEnableSounds = "0"
        fClock.muteToggleEnabled = True
    Else
        gblEnableSounds = "False"
        fClock.muteToggleEnabled = False
    End If
    
    If gblEnableTicks = "1" Then
        fClock.chainEnabled = True
        fClock.ticking = True
    Else
        fClock.chainEnabled = False
        fClock.ticking = False
    End If
    
    If gblClapperEnabled = "True" Then
        fClock.clapperEnabled = True
    Else
        fClock.clapperEnabled = False
    End If
        
    If gblEnableChimes = "1" Then
        fClock.clapperEnabled = True
    Else
        fClock.clapperEnabled = False
    End If
    
    If gblCrankEnabled = "True" Then
        fClock.crankedRaised = True
    Else
        fClock.crankedRaised = False
    End If
     
'    If gblVolumeBoost = "1" Then
'        fClock.crankedRaised = True
'    Else
'        fClock.crankedRaised = False
'    End If
      
    If gblPendulumToggleEnabled = "True" Then
        fClock.pendulumToggleEnabled = True
    End If
    
    If gblWeekdayToggleEnabled = "True" Then
        fClock.weekdayToggleEnabled = True
    End If
    
    If gblDisplayScreenToggleEnabled = "True" Then
        fClock.displayScreenToggleEnabled = True
    End If
    
    If gblTimeMachineToggleEnabled = "True" Then
        fClock.timeMachineToggleEnabled = True
    End If
    
    If gblBackToggleEnabled = "True" Then
        fClock.backToggleEnabled = True
    End If
     
    fClock.clapperEnabled = CBool(gblClapperEnabled)
    
    fClock.alarmtoggleEnabled = False
    
    If gblAlarmToggle1Enabled = "True" Then
        fClock.alarmtoggle1Enabled = True
    End If
    
    If gblAlarmToggle2Enabled = "True" Then
        fClock.alarmtoggle2Enabled = True
    End If
    
    If gblAlarmToggle3Enabled = "True" Then
        fClock.alarmtoggle3Enabled = True
    End If
    
    If gblAlarmToggle4Enabled = "True" Then
        fClock.alarmtoggle4Enabled = True
    End If
    
    If gblAlarmToggle5Enabled = "True" Then
        fClock.alarmtoggle5Enabled = True
    End If
    
    If gblAlarm1Date <> "Alarm not yet set" Then
        fClock.clockForm.Widgets("alarmtoggle1").Widget.ImageKey = "alarmtoggle1armed"
    End If

    If gblAlarm2Date <> "Alarm not yet set" Then
        fClock.clockForm.Widgets("alarmtoggle2").Widget.ImageKey = "alarmtoggle2armed"
    End If
    
    If gblAlarm3Date <> "Alarm not yet set" Then
        fClock.clockForm.Widgets("alarmtoggle3").Widget.ImageKey = "alarmtoggle3armed"
    End If
    
    If gblAlarm4Date <> "Alarm not yet set" Then
        fClock.clockForm.Widgets("alarmtoggle4").Widget.ImageKey = "alarmtoggle4armed"
    End If

    If gblAlarm5Date <> "Alarm not yet set" Then
        fClock.clockForm.Widgets("alarmtoggle5").Widget.ImageKey = "alarmtoggle5armed"
    End If

    overlayWidget.MyOpacity = Val(gblOpacity) / 100

    ' set the z-ordering of the window
    Call setAlphaFormZordering
    
    ' set the tooltips on the main screen
    Call setMainTooltips
    
    ' set the hiding time for the hiding timer, can't read the minutes from comboxbox as the prefs isn't yet open
    Call setHidingTime

    If minutesToHide > 0 Then menuForm.mnuHideWidget.Caption = "Hide Widget for " & minutesToHide & " min."
    
    ' refresh the form in order to show the above changes immediately
    fClock.clockForm.Refresh

   On Error GoTo 0
   Exit Sub

adjustMainControls_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure adjustMainControls of Module modMain " _
        & " Most likely one of the layers above is named incorrectly."

End Sub

'---------------------------------------------------------------------------------------
' Procedure : setAlphaFormZordering
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : set the z-ordering of the window
'---------------------------------------------------------------------------------------
'
Public Sub setAlphaFormZordering()

   On Error GoTo setAlphaFormZordering_Error

    If Val(gblWindowLevel) = 0 Then
        Call SetWindowPos(fClock.clockForm.hWnd, HWND_BOTTOM, 0&, 0&, 0&, 0&, OnTopFlags)
    ElseIf Val(gblWindowLevel) = 1 Then
        Call SetWindowPos(fClock.clockForm.hWnd, HWND_TOP, 0&, 0&, 0&, 0&, OnTopFlags)
    ElseIf Val(gblWindowLevel) = 2 Then
        Call SetWindowPos(fClock.clockForm.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, OnTopFlags)
    End If

   On Error GoTo 0
   Exit Sub

setAlphaFormZordering_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setAlphaFormZordering of Module modMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : readSettingsFile
' Author    : beededea
' Date      : 12/05/2020
' Purpose   : read the application's setting file and assign values to public vars
'---------------------------------------------------------------------------------------
'
Public Sub readSettingsFile(ByVal Location As String, ByVal gblSettingsFile As String)
    On Error GoTo readSettingsFile_Error

    If fFExists(gblSettingsFile) Then
        
        ' general
        gblStartup = fGetINISetting(Location, "startup", gblSettingsFile)
        gblWidgetFunctions = fGetINISetting(Location, "widgetFunctions", gblSettingsFile)
        
        gblSmoothSecondHand = fGetINISetting(Location, "smoothSecondHand", gblSettingsFile)
        
'        gblClockFaceSwitchPref = fGetINISetting(location, "clockFaceSwitchPref", gblSettingsFile)
        'gblSecondaryGaugeTimeZone = fGetINISetting(location, "secondaryGaugeTimeZone", gblSettingsFile)
        'gblSecondaryDaylightSaving = fGetINISetting(location, "secondaryDaylightSaving", gblSettingsFile)

        ' configuration
        gblEnableTooltips = fGetINISetting(Location, "enableTooltips", gblSettingsFile)
        gblEnableBalloonTooltips = fGetINISetting(Location, "enableBalloonTooltips", gblSettingsFile)
        
        gblEnablePrefsTooltips = fGetINISetting(Location, "enablePrefsTooltips", gblSettingsFile)
        gblEnablePrefsBalloonTooltips = fGetINISetting(Location, "enablePrefsBalloonTooltips", gblSettingsFile)
        
        gblShowTaskbar = fGetINISetting(Location, "showTaskbar", gblSettingsFile)
        gblShowHelp = fGetINISetting(Location, "showHelp", gblSettingsFile)
        gblTogglePendulum = fGetINISetting(Location, "togglePendulum", gblSettingsFile)
        gbl24HourClockMode = fGetINISetting(Location, "24HourClockMode", gblSettingsFile)
        
        gblDpiAwareness = fGetINISetting(Location, "dpiAwareness", gblSettingsFile)
        
        
        gblGaugeSize = fGetINISetting(Location, "gaugeSize", gblSettingsFile)
        gblScrollWheelDirection = fGetINISetting(Location, "scrollWheelDirection", gblSettingsFile)
        
        ' position
        gblAspectHidden = fGetINISetting(Location, "aspectHidden", gblSettingsFile)
        gblWidgetPosition = fGetINISetting(Location, "widgetPosition", gblSettingsFile)
        gblWidgetLandscape = fGetINISetting(Location, "widgetLandscape", gblSettingsFile)
        gblWidgetPortrait = fGetINISetting(Location, "widgetPortrait", gblSettingsFile)
        gblLandscapeFormHoffset = fGetINISetting(Location, "landscapeHoffset", gblSettingsFile)
        gblLandscapeFormVoffset = fGetINISetting(Location, "landscapeYoffset", gblSettingsFile)
        gblPortraitHoffset = fGetINISetting(Location, "portraitHoffset", gblSettingsFile)
        gblPortraitYoffset = fGetINISetting(Location, "portraitYoffset", gblSettingsFile)
        gblvLocationPercPrefValue = fGetINISetting(Location, "vLocationPercPrefValue", gblSettingsFile)
        gblhLocationPercPrefValue = fGetINISetting(Location, "hLocationPercPrefValue", gblSettingsFile)

        ' font
        gblClockFont = fGetINISetting(Location, "clockFont", gblSettingsFile)
        gblPrefsFont = fGetINISetting(Location, "prefsFont", gblSettingsFile)
        gblPrefsFontSizeHighDPI = fGetINISetting(Location, "prefsFontSizeHighDPI", gblSettingsFile)
        gblPrefsFontSizeLowDPI = fGetINISetting(Location, "prefsFontSizeLowDPI", gblSettingsFile)
        gblPrefsFontItalics = fGetINISetting(Location, "prefsFontItalics", gblSettingsFile)
        gblPrefsFontColour = fGetINISetting(Location, "prefsFontColour", gblSettingsFile)
    
        gblDisplayScreenFont = fGetINISetting(Location, "displayScreenFont", gblSettingsFile)
        gblDisplayScreenFontSize = fGetINISetting(Location, "displayScreenFontSize", gblSettingsFile)
        gblDisplayScreenFontItalics = fGetINISetting(Location, "displayScreenFontItalics", gblSettingsFile)
        gblDisplayScreenFontColour = fGetINISetting(Location, "displayScreenFontColour", gblSettingsFile)
       
        ' sound
        gblEnableSounds = fGetINISetting(Location, "enableSounds", gblSettingsFile)
        gblEnableTicks = fGetINISetting(Location, "enableTicks", gblSettingsFile)
        gblEnableChimes = fGetINISetting(Location, "enableChimes", gblSettingsFile)
        gblVolumeBoost = fGetINISetting(Location, "volumeBoost", gblSettingsFile)
        
        
        ' development
        gblDebug = fGetINISetting(Location, "debug", gblSettingsFile)
        gblDblClickCommand = fGetINISetting(Location, "dblClickCommand", gblSettingsFile)
        gblOpenFile = fGetINISetting(Location, "openFile", gblSettingsFile)
        gblDefaultEditor = fGetINISetting(Location, "defaultEditor", gblSettingsFile)
        
        ' other
        gblClockHighDpiXPos = fGetINISetting("Software\SteampunkClockCalendar", "clockHighDpiXPos", gblSettingsFile)
        gblClockHighDpiYPos = fGetINISetting("Software\SteampunkClockCalendar", "clockHighDpiYPos", gblSettingsFile)
        
        gblClockLowDpiXPos = fGetINISetting("Software\SteampunkClockCalendar", "clockLowDpiXPos", gblSettingsFile)
        gblClockLowDpiYPos = fGetINISetting("Software\SteampunkClockCalendar", "clockLowDpiYPos", gblSettingsFile)
        
        gblLastSelectedTab = fGetINISetting(Location, "lastSelectedTab", gblSettingsFile)
        gblSkinTheme = fGetINISetting(Location, "skinTheme", gblSettingsFile)
        
        ' window
        gblWindowLevel = fGetINISetting(Location, "windowLevel", gblSettingsFile)
        gblPreventDragging = fGetINISetting(Location, "preventDragging", gblSettingsFile)
        gblOpacity = fGetINISetting(Location, "opacity", gblSettingsFile)
        
        ' we do not want the widget to hide at startup
        'gblWidgetHidden = fGetINISetting(location, "widgetHidden", gblSettingsFile)
        gblWidgetHidden = "0"
        
        gblHidingTime = fGetINISetting(Location, "hidingTime", gblSettingsFile)
        gblIgnoreMouse = fGetINISetting(Location, "ignoreMouse", gblSettingsFile)
        gblMultiMonitorResize = fGetINISetting(Location, "multiMonitorResize", gblSettingsFile)
        
         
        gblFirstTimeRun = fGetINISetting(Location, "firstTimeRun", gblSettingsFile)
        
        gblsetToggleEnabled = fGetINISetting(Location, "setToggleEnabled", gblSettingsFile)
        gblMuteToggleEnabled = fGetINISetting(Location, "muteToggleEnabled", gblSettingsFile)
        gblPendulumToggleEnabled = fGetINISetting(Location, "pendulumToggleEnabled", gblSettingsFile)
        gblPendulumEnabled = fGetINISetting(Location, "pendulumEnabled", gblSettingsFile)
        gblWeekdayToggleEnabled = fGetINISetting(Location, "weekdayToggleEnabled", gblSettingsFile)
        gblDisplayScreenToggleEnabled = fGetINISetting(Location, "displayScreenToggleEnabled", gblSettingsFile)
        gblTimeMachineToggleEnabled = fGetINISetting(Location, "timeMachineToggleEnabled", gblSettingsFile)
        gblBackToggleEnabled = fGetINISetting(Location, "backToggleEnabled", gblSettingsFile)
        gblClapperEnabled = fGetINISetting(Location, "clapperEnabled", gblSettingsFile)
        gblChainEnabled = fGetINISetting(Location, "chainEnabled", gblSettingsFile)
        gblCrankEnabled = fGetINISetting(Location, "crankEnabled", gblSettingsFile)
        gblAlarmToggle1Enabled = fGetINISetting(Location, "alarmToggle1Enabled", gblSettingsFile)
        gblAlarmToggle2Enabled = fGetINISetting(Location, "alarmToggle2Enabled", gblSettingsFile)
        gblAlarmToggle3Enabled = fGetINISetting(Location, "alarmToggle3Enabled", gblSettingsFile)
        gblAlarmToggle4Enabled = fGetINISetting(Location, "alarmToggle4Enabled", gblSettingsFile)
        gblAlarmToggle5Enabled = fGetINISetting(Location, "alarmToggle5Enabled", gblSettingsFile)
        gblAlarm1Date = fGetINISetting(Location, "alarm1Date", gblSettingsFile)
        gblAlarm2Date = fGetINISetting(Location, "alarm2Date", gblSettingsFile)
        gblAlarm3Date = fGetINISetting(Location, "alarm3Date", gblSettingsFile)
        gblAlarm4Date = fGetINISetting(Location, "alarm4Date", gblSettingsFile)
        gblAlarm5Date = fGetINISetting(Location, "alarm5Date", gblSettingsFile)
        
        gblAlarm1Time = fGetINISetting(Location, "alarm1Time", gblSettingsFile)
        gblAlarm2Time = fGetINISetting(Location, "alarm2Time", gblSettingsFile)
        gblAlarm3Time = fGetINISetting(Location, "alarm3Time", gblSettingsFile)
        gblAlarm4Time = fGetINISetting(Location, "alarm4Time", gblSettingsFile)
        gblAlarm5Time = fGetINISetting(Location, "alarm5Time", gblSettingsFile)
                           
        gblClockSecondaryHeightRatio = fGetINISetting(Location, "clockSecondaryHeightRatio", gblSettingsFile)
        gblClockPrimaryHeightRatio = fGetINISetting(Location, "clockPrimaryHeightRatio", gblSettingsFile)
        
        gblMessageAHeightTwips = fGetINISetting(Location, "messageAHeightTwips", gblSettingsFile)
        gblMessageAWidthTwips = fGetINISetting(Location, "messageAWidthTwips ", gblSettingsFile)
        
    End If
        
    gblAlarm1 = gblAlarm1Date & " " & gblAlarm1Time
    gblAlarm2 = gblAlarm2Date & " " & gblAlarm2Time
    gblAlarm3 = gblAlarm3Date & " " & gblAlarm3Time
    gblAlarm4 = gblAlarm4Date & " " & gblAlarm4Time
    gblAlarm5 = gblAlarm5Date & " " & gblAlarm5Time

   On Error GoTo 0
   Exit Sub

readSettingsFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readSettingsFile of Module common2"

End Sub


    
'---------------------------------------------------------------------------------------
' Procedure : validateInputs
' Author    : beededea
' Date      : 17/06/2020
' Purpose   : validate the relevant entries from the settings.ini file in user appdata
'---------------------------------------------------------------------------------------
'
Public Sub validateInputs()
    
   On Error GoTo validateInputs_Error
            
        ' general
        If gblWidgetFunctions = vbNullString Then gblWidgetFunctions = "1" ' always turn
'        If gblAnimationInterval = vbNullString Then gblAnimationInterval = "130"
        If gblStartup = vbNullString Then gblStartup = "1"
        
        If gblSmoothSecondHand = vbNullString Then gblSmoothSecondHand = "0"
        
        'If gblClockFaceSwitchPref = vbNullString Then gblClockFaceSwitchPref = "0"

        'If gblSecondaryGaugeTimeZone = vbNullString Then gblSecondaryGaugeTimeZone = "1"
        'If gblSecondaryDaylightSaving = vbNullString Then gblSecondaryDaylightSaving = "1"

        ' Configuration
        If gblEnableTooltips = vbNullString Then gblEnableTooltips = "False"
        If gblEnableBalloonTooltips = vbNullString Then gblEnableBalloonTooltips = "True"
        
        If gblEnablePrefsTooltips = vbNullString Then gblEnablePrefsTooltips = "false"
        If gblEnablePrefsBalloonTooltips = vbNullString Then gblEnablePrefsBalloonTooltips = "True"
        
        If gblShowTaskbar = vbNullString Then gblShowTaskbar = "0"
        If gblShowHelp = vbNullString Then gblShowHelp = "0"
        If gblTogglePendulum = vbNullString Then gblTogglePendulum = "0"
        If gbl24HourClockMode = vbNullString Then gbl24HourClockMode = "1"
        
        If gblDpiAwareness = vbNullString Then gblDpiAwareness = "0"
        If gblGaugeSize = vbNullString Then gblGaugeSize = "25"
        If gblScrollWheelDirection = vbNullString Then gblScrollWheelDirection = "1"
               
        ' fonts
        If gblPrefsFont = vbNullString Then gblPrefsFont = "times new roman"
        If gblPrefsFontSizeHighDPI = vbNullString Then gblPrefsFontSizeHighDPI = "8"
        If gblPrefsFontSizeLowDPI = vbNullString Then gblPrefsFontSizeLowDPI = "8"
        If gblPrefsFontItalics = vbNullString Then gblPrefsFontItalics = "false"
        If gblPrefsFontColour = vbNullString Then gblPrefsFontColour = "0"

        If gblClockFont = vbNullString Then gblClockFont = gblPrefsFont

        If gblDisplayScreenFont = vbNullString Then gblDisplayScreenFont = "courier new"
        If gblDisplayScreenFont = "Courier  New" Then gblDisplayScreenFont = "courier new"
        If gblDisplayScreenFontSize = vbNullString Then gblDisplayScreenFontSize = "5"
        If gblDisplayScreenFontItalics = vbNullString Then gblDisplayScreenFontItalics = "false"
        If gblDisplayScreenFontColour = vbNullString Then gblDisplayScreenFontColour = "0"

        ' sounds
        
        If gblEnableSounds = vbNullString Then gblEnableSounds = "1"
        If gblEnableTicks = vbNullString Then gblEnableTicks = "0"
        If gblEnableChimes = vbNullString Then gblEnableChimes = "0"
        If gblVolumeBoost = vbNullString Then gblVolumeBoost = "0"
        
        
        ' position
        If gblAspectHidden = vbNullString Then gblAspectHidden = "0"
        If gblWidgetPosition = vbNullString Then gblWidgetPosition = "0"
        If gblWidgetLandscape = vbNullString Then gblWidgetLandscape = "0"
        If gblWidgetPortrait = vbNullString Then gblWidgetPortrait = "0"
        If gblLandscapeFormHoffset = vbNullString Then gblLandscapeFormHoffset = vbNullString
        If gblLandscapeFormVoffset = vbNullString Then gblLandscapeFormVoffset = vbNullString
        If gblPortraitHoffset = vbNullString Then gblPortraitHoffset = vbNullString
        If gblPortraitYoffset = vbNullString Then gblPortraitYoffset = vbNullString
        If gblvLocationPercPrefValue = vbNullString Then gblvLocationPercPrefValue = vbNullString
        If gblhLocationPercPrefValue = vbNullString Then gblhLocationPercPrefValue = vbNullString
                
        ' development
        If gblDebug = vbNullString Then gblDebug = "0"
        If gblDblClickCommand = vbNullString And gblFirstTimeRun = "True" Then gblDblClickCommand = "mmsys.cpl"
        If gblOpenFile = vbNullString Then gblOpenFile = vbNullString
        If gblDefaultEditor = vbNullString Then gblDefaultEditor = vbNullString
        
        ' window
        If gblWindowLevel = vbNullString Then gblWindowLevel = "1" 'WindowLevel", gblSettingsFile)
        If gblOpacity = vbNullString Then gblOpacity = "100"
        If gblWidgetHidden = vbNullString Then gblWidgetHidden = "0"
        If gblHidingTime = vbNullString Then gblHidingTime = "0"
        If gblIgnoreMouse = vbNullString Then gblIgnoreMouse = "0"
        If gblPreventDragging = vbNullString Then gblPreventDragging = "0"
        If gblMultiMonitorResize = vbNullString Then gblMultiMonitorResize = "0"
        
        
        ' other
        If gblFirstTimeRun = vbNullString Then gblFirstTimeRun = "true"
        If gblLastSelectedTab = vbNullString Then gblLastSelectedTab = "general"
        If gblSkinTheme = vbNullString Then gblSkinTheme = "dark"
        
        
        ' clock UI element state
        If gblsetToggleEnabled = vbNullString Then gblsetToggleEnabled = "False"
        If gblMuteToggleEnabled = vbNullString Then gblMuteToggleEnabled = "False"
        If gblPendulumToggleEnabled = vbNullString Then gblPendulumToggleEnabled = "False"
        If gblPendulumEnabled = vbNullString Then gblPendulumEnabled = "False"
        
        
        If gblWeekdayToggleEnabled = vbNullString Then gblWeekdayToggleEnabled = "False"
        If gblDisplayScreenToggleEnabled = vbNullString Then gblDisplayScreenToggleEnabled = "True"
        If gblTimeMachineToggleEnabled = vbNullString Then gblTimeMachineToggleEnabled = "False"
        If gblBackToggleEnabled = vbNullString Then gblBackToggleEnabled = "False"
        If gblClapperEnabled = vbNullString Then gblClapperEnabled = "True"
        If gblChainEnabled = vbNullString Then gblChainEnabled = "True"
        If gblCrankEnabled = vbNullString Then gblCrankEnabled = "True"
        If gblAlarmToggle1Enabled = vbNullString Then gblAlarmToggle1Enabled = "False"
        If gblAlarmToggle2Enabled = vbNullString Then gblAlarmToggle2Enabled = "False"
        If gblAlarmToggle3Enabled = vbNullString Then gblAlarmToggle3Enabled = "False"
        If gblAlarmToggle4Enabled = vbNullString Then gblAlarmToggle4Enabled = "False"
        If gblAlarmToggle5Enabled = vbNullString Then gblAlarmToggle5Enabled = "False"
        
        If gblAlarm1Date = vbNullString Then gblAlarm1Date = "Alarm not yet set"
        If gblAlarm2Date = vbNullString Then gblAlarm2Date = "Alarm not yet set"
        If gblAlarm3Date = vbNullString Then gblAlarm3Date = "Alarm not yet set"
        If gblAlarm4Date = vbNullString Then gblAlarm4Date = "Alarm not yet set"
        If gblAlarm5Date = vbNullString Then gblAlarm5Date = "Alarm not yet set"
        
        If gblClockPrimaryHeightRatio = "" Then gblClockPrimaryHeightRatio = "1"
        If gblClockSecondaryHeightRatio = "" Then gblClockSecondaryHeightRatio = "1"
        
        
   On Error GoTo 0
   Exit Sub

validateInputs_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure validateInputs of form modMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : getTrinketsFile
' Author    : beededea
' Date      : 17/10/2019
' Purpose   : get this tool's entry in the trinkets settings file and assign the app.path
'---------------------------------------------------------------------------------------
'
Private Sub getTrinketsFile()
    On Error GoTo getTrinketsFile_Error
    
    Dim iFileNo As Integer: iFileNo = 0
    
    gblTrinketsDir = fSpecialFolder(feUserAppData) & "\trinkets" ' just for this user alone
    gblTrinketsFile = gblTrinketsDir & "\" & widgetName & ".ini"
        
    'if the folder does not exist then create the folder
    If Not fDirExists(gblTrinketsDir) Then
        MkDir gblTrinketsDir
    End If

    'if the settings.ini does not exist then create the file by copying
    If Not fFExists(gblTrinketsFile) Then

        iFileNo = FreeFile
        'open the file for writing
        Open gblTrinketsFile For Output As #iFileNo
        Write #iFileNo, App.path & "\" & App.EXEName & ".exe"
        Write #iFileNo,
        Close #iFileNo
    End If
    
   On Error GoTo 0
   Exit Sub

getTrinketsFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getTrinketsFile of Form modMain"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : getToolSettingsFile
' Author    : beededea
' Date      : 17/10/2019
' Purpose   : get this tool's settings file and assign to a global var
'---------------------------------------------------------------------------------------
'
Private Sub getToolSettingsFile()
    On Error GoTo getToolSettingsFile_Error
    ''If debugflg = 1  Then Debug.Print "%getToolSettingsFile"
    
    Dim iFileNo As Integer: iFileNo = 0
    
    gblSettingsDir = fSpecialFolder(feUserAppData) & "\SteampunkClockCalendar" ' just for this user alone
    gblSettingsFile = gblSettingsDir & "\settings.ini"
        
    'if the folder does not exist then create the folder
    If Not fDirExists(gblSettingsDir) Then
        MkDir gblSettingsDir
    End If

    'if the settings.ini does not exist then create the file by copying
    If Not fFExists(gblSettingsFile) Then

        iFileNo = FreeFile
        'open the file for writing
        Open gblSettingsFile For Output As #iFileNo
        Close #iFileNo
    End If
    
   On Error GoTo 0
   Exit Sub

getToolSettingsFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getToolSettingsFile of Form modMain"

End Sub



'
'---------------------------------------------------------------------------------------
' Procedure : configureTimers
' Author    : beededea
' Date      : 07/05/2023
' Purpose   : configure any global timers here
'---------------------------------------------------------------------------------------
'
Private Sub configureTimers()

    On Error GoTo configureTimers_Error
    
    oldgblSettingsModificationTime = FileDateTime(gblSettingsFile)

    frmTimer.tmrScreenResolution.Enabled = True
    frmTimer.unhideTimer.Enabled = True

    On Error GoTo 0
    Exit Sub

configureTimers_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure configureTimers of Module modMain"
            Resume Next
          End If
    End With
 
End Sub

'
'---------------------------------------------------------------------------------------
' Procedure : setHidingTime
' Author    : beededea
' Date      : 07/05/2023
' Purpose   : set the hiding time for the hiding timer, can't read the minutes from comboxbox as the prefs isn't yet open
'---------------------------------------------------------------------------------------
'
Private Sub setHidingTime()
    
    On Error GoTo setHidingTime_Error

    If gblHidingTime = "0" Then minutesToHide = 1
    If gblHidingTime = "1" Then minutesToHide = 5
    If gblHidingTime = "2" Then minutesToHide = 10
    If gblHidingTime = "3" Then minutesToHide = 20
    If gblHidingTime = "4" Then minutesToHide = 30
    If gblHidingTime = "5" Then minutesToHide = 60

    On Error GoTo 0
    Exit Sub

setHidingTime_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setHidingTime of Module modMain"
            Resume Next
          End If
    End With

End Sub

'---------------------------------------------------------------------------------------
' Procedure : createRCFormsOnCurrentDisplay
' Author    : beededea
' Date      : 07/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub createRCFormsOnCurrentDisplay()
    On Error GoTo createRCFormsOnCurrentDisplay_Error

    With New_c.Displays(1) 'get the current Display
      Call fMain.initAndShowAboutForm(.WorkLeft, .WorkTop, 1000, 1000, widgetName)
    End With

    With New_c.Displays(1) 'get the current Display
      Call fMain.initAndShowLicenceForm(.WorkLeft, .WorkTop, 1000, 1000, widgetName)
    End With
    
        On Error GoTo 0
    Exit Sub

createRCFormsOnCurrentDisplay_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure createRCFormsOnCurrentDisplay of Module modMain"
            Resume Next
          End If
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : handleUnhideMode
' Author    : beededea
' Date      : 13/05/2023
' Purpose   : when run in 'unhide' mode it writes the settings file then exits, the other
'             running but hidden process will unhide itself by timer.
'---------------------------------------------------------------------------------------
'
Private Sub handleUnhideMode(ByVal thisUnhideMode As String)
    
    On Error GoTo handleUnhideMode_Error

    If thisUnhideMode = "unhide" Then     'parse the command line
        gblUnhide = "true"
        sPutINISetting "Software\SteampunkClockCalendar", "unhide", gblUnhide, gblSettingsFile
        Call thisForm_Unload
        End
    End If

    On Error GoTo 0
    Exit Sub

handleUnhideMode_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure handleUnhideMode of Module modMain"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : loadExcludePathCollection
' Author    : beededea
' Date      : 30/07/2023
' Purpose   : Do not create Widgets for those in the exclude list.
'             all non UI-interacting elements (no mouse events) must be inserted here
'---------------------------------------------------------------------------------------
'
Private Sub loadExcludePathCollection()

    'all of these will be rendered in cwOverlay in the same order as below
    On Error GoTo loadExcludePathCollection_Error

    With fClock.collPSDNonUIElements ' the exclude list

        .Add Empty, "secondhand"
        .Add Empty, "hourhand"
        .Add Empty, "minutehand"
        .Add Empty, "hole"
        .Add Empty, "cable"
        .Add Empty, "pendulum"
    End With
    

   On Error GoTo 0
   Exit Sub

loadExcludePathCollection_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadExcludePathCollection of Module modMain"

End Sub


''---------------------------------------------------------------------------------------
'' Procedure : ExportPngs
'' Author    : Olaf
'' Date      : 06/08/2023
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Sub exportPngs(PSD_FileNameOrByteArray, ByVal pngFolder As String)
'   On Error GoTo ExportPngs_Error
'
'  New_c.FSO.EnsurePath pngFolder 'make sure the PngFolder-Path "materializes itself" in the FileSystem
'  New_c.FSO.EnsurePathEndSep pngFolder 'add a backslash to the PngFolder-param (in case it was missing)
'
'  With New_c.SimplePSD(PSD_FileNameOrByteArray)  'create a new PSD-Parser.instance (and load the passed content)
'    Dim i As Long
'    For i = 0 To .LayersCount - 1 'loop over all the Layers in the PSD
'      If .LayerByteSize(i) Then   'this is an Alpha-Surface-Layer with "meat" (and not a group-specification)
'         .LayerSurface(i).WriteContentToPngFile pngFolder & Replace(.LayerPath(i), "/", "_") & ".png"
'      End If
'    Next
'  End With
'
'   On Error GoTo 0
'   Exit Sub
'
'ExportPngs_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ExportPngs of Module modMain"
'End Sub







     



' .74 DAEB 22/05/2022 rDIConConfig.frm Msgbox replacement that can be placed on top of the form instead as the middle of the screen, see Steamydock for a potential replacement?
'---------------------------------------------------------------------------------------
' Procedure : msgBoxA
' Author    : beededea
' Date      : 20/05/2022
' Purpose   : ans = msgBoxA("main message", vbOKOnly, "title bar message", False)
'---------------------------------------------------------------------------------------
'
Public Function msgBoxA(ByVal msgBoxPrompt As String, Optional ByVal msgButton As VbMsgBoxResult, Optional ByVal msgTitle As String, Optional ByVal msgShowAgainChkBox As Boolean = False, Optional ByRef msgContext As String = "none") As Integer
     
    ' set the defined properties of a form
    On Error GoTo msgBoxA_Error

    frmMessage.propMessage = msgBoxPrompt
    frmMessage.propTitle = msgTitle
    frmMessage.propShowAgainChkBox = msgShowAgainChkBox
    frmMessage.propButtonVal = msgButton
    frmMessage.propMsgContext = msgContext
    Call frmMessage.Display ' run a subroutine in the form that displays the form

    msgBoxA = frmMessage.propReturnedValue

    On Error GoTo 0
    Exit Function

msgBoxA_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure msgBoxA of Module mdlMain"
            Resume Next
          End If
    End With

End Function






