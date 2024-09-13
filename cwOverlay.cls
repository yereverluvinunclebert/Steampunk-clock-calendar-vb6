VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwOverlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule IntegerDataType, ModuleWithoutFolder
'---------------------------------------------------------------------------------------
' Module    : cwOverlay
' Author    : beededea
' Date      : 31/07/2023
' Purpose   : create a single fully click-through Widget, which allows for all kind of
'             things to render on top of the PSD-generated Widgets created in cfClock
'---------------------------------------------------------------------------------------

Option Explicit

Public cCollOverlayImages As cCollection
Public sldCurWidth As Double

Public cxPerc As Single
Public cyPerc As Single

Private hdeg As Double
Private mdeg As Double
Private SDeg As Double
Private Sec As Integer

Private WithEvents tmrClock As cTimer
Attribute tmrClock.VB_VarHelpID = -1
Private WithEvents tmrPendulum As cTimer
Attribute tmrPendulum.VB_VarHelpID = -1
Public WithEvents tmrTimeShift As cTimer
Attribute tmrTimeShift.VB_VarHelpID = -1
Private WithEvents tmrDigitRotator As cTimer
Attribute tmrDigitRotator.VB_VarHelpID = -1

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1

Private firstRun As Boolean
Private thisHour As Integer
Private thisMin As Integer
Private thisSec As Integer

Private pendulumAcceleration As Double
Private gravity As Double
Private pendulumAngle As Double
Private pendulumVelocity As Double

Private mLocked As Boolean ' property Locked
Private mTimeMachine As Boolean
Private mHidden As Boolean ' property Hidden
Private mSwingPendulum As Boolean



'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : olaf schmidt
' Date      : 31/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    Dim intervalValue As Integer: intervalValue = 0
    On Error GoTo Class_Initialize_Error

    Set W = Cairo.WidgetBase
        W.BackColor = -1            ' we don't use any BackColor here, transparent
        W.ImplementsHitTest = True  ' essential

    intervalValue = 1000
    firstRun = True
    
    ' variables set for the pendulum
    ' g = -0.1, a = 0.6, v = 2.5 uses 3% of cpu with a 40ms timer

    gravity = -0.2
    pendulumAcceleration = 0.6
    pendulumVelocity = 3.5
    
    Set tmrClock = Nothing
    Set tmrClock = New_c.Timer(intervalValue, True)
    
    tmrClock.Enabled = True
    
    Set tmrPendulum = Nothing
    Set tmrPendulum = New_c.Timer(70, False)
    
    tmrPendulum.Enabled = False
    
    Set tmrTimeShift = Nothing
    Set tmrTimeShift = New_c.Timer(100, False)
    
    tmrTimeShift.Enabled = False
   
    Set tmrDigitRotator = Nothing
    Set tmrDigitRotator = New_c.Timer(100, False)
    
    tmrDigitRotator.Enabled = False
   
   On Error GoTo 0
   Exit Sub

Class_Initialize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cwOverlay"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Widget
' Author    : olaf schmidt
' Date      : 21/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Widget() As cWidgetBase
   On Error GoTo Widget_Error

    Set Widget = W

   On Error GoTo 0
   Exit Property

Widget_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Widget of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Procedure : Widgets
' Author    : olaf schmidt
' Date      : 21/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Widgets() As cWidgets
   On Error GoTo Widgets_Error

    Set Widgets = W.Widgets

   On Error GoTo 0
   Exit Property

Widgets_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Widgets of Class Module cwOverlay"
End Property

'---------------------------------------------------------------------------------------
' Property : Locked
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Locked() As Boolean
   On Error GoTo LockedGet_Error

   Locked = mLocked

   On Error GoTo 0
   Exit Property

LockedGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Locked of Class Module cwGlobe"
End Property
'---------------------------------------------------------------------------------------
' Property  : Locked
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Locked(ByVal newValue As Boolean)
   On Error GoTo LockedLet_Error

   If mLocked <> newValue Then mLocked = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

LockedLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Locked of Class Module cwGlobe"
End Property


'---------------------------------------------------------------------------------------
' Property : SwingPendulum
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get SwingPendulum() As Boolean
   On Error GoTo SwingPendulumGet_Error

   SwingPendulum = mSwingPendulum

   On Error GoTo 0
   Exit Property

SwingPendulumGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property SwingPendulum of Class Module cfClock"
End Property
'---------------------------------------------------------------------------------------
' Property  : SwingPendulum
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let SwingPendulum(ByVal newValue As Boolean)

    On Error GoTo SwingPendulumLet_Error
        
    If mSwingPendulum = True Then
        tmrPendulum.Enabled = True
    Else
        tmrPendulum.Enabled = False
    End If
    
    If mSwingPendulum <> newValue Then mSwingPendulum = newValue Else Exit Property
    SwingPendulum = mSwingPendulum
    
    W.Refresh

   On Error GoTo 0
   Exit Property

SwingPendulumLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property SwingPendulum of Class Module cfClock"
End Property

'---------------------------------------------------------------------------------------
' Property : TimeMachine
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get TimeMachine() As Boolean
   On Error GoTo TimeMachineGet_Error

   TimeMachine = mTimeMachine

   On Error GoTo 0
   Exit Property

TimeMachineGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property TimeMachine of Class Module cwGlobe"
End Property
'---------------------------------------------------------------------------------------
' Property  : TimeMachine
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let TimeMachine(ByVal newValue As Boolean)
   On Error GoTo TimeMachineLet_Error

   If mTimeMachine <> newValue Then mTimeMachine = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

TimeMachineLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property TimeMachine of Class Module cwGlobe"
End Property

'---------------------------------------------------------------------------------------
' Property  : Hidden
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Hidden() As Boolean
   On Error GoTo HiddenGet_Error

   Hidden = mHidden

   On Error GoTo 0
   Exit Property

HiddenGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Hidden of Class Module cwGlobe"
End Property

    


'---------------------------------------------------------------------------------------
' Property  : Hidden
' Author    : beededea
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Hidden(ByVal newValue As Boolean)
   On Error GoTo HiddenLet_Error

    If mHidden <> newValue Then mHidden = newValue Else Exit Property
    If mHidden = True Then
        W.Visible = False
        
        W.Refresh
    Else
        W.Visible = True
        
        W.Refresh
    End If

   On Error GoTo 0
   Exit Property

HiddenLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Hidden of Class Module cwGlobe"
End Property
'
'---------------------------------------------------------------------------------------------
' Procedure : W_HitTest
' Author    : olaf schmidt
' Date      : 01/08/2023
' Purpose   : This is run continuously for the whole widget.
'             For this Overlay-Widget alone, we make everything "fully click-through"
'---------------------------------------------------------------------------------------------
'
Private Sub W_HitTest(ByVal x As Single, ByVal y As Single, HitResultHit As Boolean) 'ensure ClickThrough-behaviour in ImagePixels which are "fully Alpha"
   On Error GoTo W_HitTest_Error

   HitResultHit = False

   On Error GoTo 0
   Exit Sub

W_HitTest_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_HitTest of Class Module cwOverlay"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : W_Paint
' Author    : olaf schmidt
' Date      : 01/08/2023
' Purpose   : this is run when the graphics are initially placed on the screen, refreshing and when a mouse cursor is interacting
'---------------------------------------------------------------------------------------
'
Private Sub W_Paint(CC As RC6.cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
    Dim I As Long: I = 0
    Dim Arr As Variant
    Dim Key As String: Key = vbNullString
    Dim Pat As cCairoPattern
    Dim someOpacity As Double: someOpacity = 0
    Dim grey As String: grey = vbNullString
    Dim interval As Integer: interval = 0
    Dim position As Integer: position = 0
    Dim a As Integer: a = 0
    
    grey = RGB(18, 18, 18)
    
    'set some Font- and ForeColor-Properties and draw some Text on the Widget
    W.FontName = "courier new"
    W.ForeColor = grey
    W.FontSize = 5
    'W.FontBold = True
    W.SelectFontSettingsInto CC
    W.Refresh
    
    someOpacity = 0.6
    position = 194
    interval = 9
       
    On Error GoTo W_Paint_Error
  
    For I = 0 To cCollOverlayImages.Count - 1 ' this is the ExcludedPaths collection populated in cfClock by duplicating collPSDNonUIElements.
      Arr = cCollOverlayImages.ItemByIndex(I)
      Key = cCollOverlayImages.KeyByIndex(I)
      
      If IsArray(Arr) Then
             CC.Save
             'derive a Pattern-Obj from each retrieved ImageList-Surface-reference
             Set Pat = Cairo.ImageList(Key).CreateSurfacePattern
            
            ' for specific keys, perform actions &c
             Select Case Key
                 Case "hourhand"
                    Set Pat = Cairo.ImageList(Key).CreateSurfacePattern
                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(4, 4) 'shift to the "rotation-point" within the Pattern
                    CC.TranslateDrawings W.Width * cxPerc, W.Height * cyPerc 'shift to the center-point of the Overlay-Widget-area

                    'CC.TranslateDrawings 645, 333 'shift to the center-point of the Overlay-Widget-area
                    CC.RotateDrawingsDeg -90 + hdeg  'now we can rotate (the leading negative degrees are the rotation-offsets for 12 o'clock)

                 Case "minutehand"  'minutes-shadow- and minutes-hand
                    Set Pat = Cairo.ImageList(Key).CreateSurfacePattern
                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(3, 3) 'shift to the "rotation-point" within the Pattern
                    CC.TranslateDrawings 646.5, 332 'shift to the center-point of the Overlay-Widget-area

                    CC.RotateDrawingsDeg -90 + mdeg 'now we can rotate (the leading negative degrees are the rotation-offsets for 12 o'clock)

                 Case "secondhand" 'seconds-shadow- and seconds-hand
                    Set Pat = Cairo.ImageList(Key).CreateSurfacePattern
                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(3, 3) 'shift to the "rotation-point" within the Pattern
                    CC.TranslateDrawings 646.5, 332 'shift to the center-point of the Overlay-Widget-area

                    CC.RotateDrawingsDeg -90 + SDeg 'now we can rotate (the leading negative degrees are the rotation-offsets for 12 o'clock)

                 Case "hole"
                    Set Pat = Cairo.ImageList(Key).CreateSurfacePattern
                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(4, 4) 'shift to the "rotation-point" within the Pattern

                    CC.TranslateDrawings 646, 332 'shift to the center-point of the Overlay-Widget-area

                Case "cable"
                    Set Pat = Cairo.ImageList(Key).CropSurface(0, 0, 200, 27).CreateSurfacePattern 'cut-out a nice (repeatable) part of the original "cable"-image (via CropSurface)
                    
                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(sldCurWidth, -Arr(1)) 'sldCurWidth in the x-Coord will ensure that nice "spooling-effect"
                        Pat.Extend = CAIRO_EXTEND_REPEAT

                    'since we've set CAIRO_EXTEND_REPEAT, we now have to clip the range it is "allowed to extend to"
                    CC.ClipExplicit Arr(0) - sldCurWidth, Arr(1), sldCurWidth + 15, 27
                
                Case "pendulum"
                    Set Pat = Cairo.ImageList(Key).CreateSurfacePattern
                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(21, 0) 'shift to the "rotation-point" within the Pattern
                    CC.TranslateDrawings 624, 325

                    CC.RotateDrawingsDeg pendulumAngle 'now we can rotate (the leading negative degrees are the rotation-offsets for 12 o'clock)
                
                Case Else 'plain output for everything else
                       
                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(-Arr(0), -Arr(1)) 'here we only shift the Pattern via the original PSD-Offsets (stored in Arr(0) and Arr(1))
                    
             End Select
    
            CC.Paint W.AlphaInherited * Arr(2), Pat  'finally we render the Pattern with the current PSD-LayerAlpha, present in Arr(2)
            CC.Restore
      End If
    Next
     
    ' print the console text on the screen
    If gblDisplayScreenToggleEnabled = "True" Then
        For a = 0 To 14
            CC.DrawText 423, position, 220, 20, gblTerminalRows(a), , , 2, True, , someOpacity
            position = position - interval
        Next a

        CC.DrawText 423, 203, 200, 20, ".......", , , 2, True, , someOpacity
        CC.DrawText 423, 211, 200, 20, ".......", , , 2, True, , someOpacity
        CC.DrawText 423, 218, 200, 20, ".......", , , 2, True, , someOpacity
        CC.DrawText 423, 227, 200, 20, ".......", , , 2, True, , someOpacity
    End If

    ' if first time running then populate the dgitial clocks or if one of the alarms is currently being viewed then populate the clocks with the alarm digits.
    If triggerDigitalClockPopulation = True Then
        Call populateDigitalClockCache(CC)
        triggerDigitalClockPopulation = False
    End If
    
    If thisSec <= 1 Then
        Call populateDigitalMinsCache(CC)
        If thisMin = 0 Then
            Call populateDigitalHoursCache(CC)
            If thisHour = 0 Then
                Call populateDigitalClockCache(CC)
            End If
        End If
    End If

    
    ' test for alarms, using globals to reduce the cpu cycles
    If gblAlarmToggle1Enabled = "True" Or gblAlarmToggle2Enabled = "True" Or gblAlarmToggle3Enabled = "True" Or gblAlarmToggle4Enabled = "True" Or gblAlarmToggle5Enabled = "True" Then Call populateDigitalClockCache(CC)
    
    ' the above routines populate the digital clock cache, this routine shows the digital clocks
    Call DrawDigitalClocks(CC)
     
    If mTimeMachine = True Then Call drawTimeMachineImages(CC)

   On Error GoTo 0
   Exit Sub

W_Paint_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_Paint of Class Module cwOverlay"
End Sub
 
'---------------------------------------------------------------------------------------
' Procedure : drawTimeMachineImages
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub drawTimeMachineImages(ByRef CC As cCairoContext)
    
    



End Sub
 
'---------------------------------------------------------------------------------------
' Procedure : drawAMPMDigitalDay
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub drawAMPMDigitalDay(ByRef CC As cCairoContext, Optional ByVal alarmHour As String)
    ' no error handling code - for speed
    Dim localHour As Integer: localHour = 0
    
    If alarmHour <> vbNullString Then
        localHour = CInt(alarmHour)
    Else
        localHour = thisHour
    End If
    
    If localHour < 12 Then ' uses vb6 coerce a string to a numeric without calling a cast function, quicker?
         CC.RenderSurfaceContent "AM", 621, 387, , , , 1
    Else
         CC.RenderSurfaceContent "PM", 621, 387, , , , 1
    End If


End Sub
'---------------------------------------------------------------------------------------
' Procedure : oldDrawDigitalClocks
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub oldDrawDigitalClocks(ByRef CC As cCairoContext)
    Dim theDay As String: theDay = vbNullString
    Dim theMonth As String: theMonth = vbNullString
    Dim theYear As String: theYear = vbNullString
    Dim theHours As String: theHours = vbNullString
    Dim theMins As String: theMins = vbNullString
    
    If fClock.alarmtoggle1Enabled = True Then
        ' read the date and time variables for alarm 1
        If gblAlarm1Date = "Alarm not yet set" Then
            theDay = "01"
            theMonth = "01"
            theYear = "1970"
            theHours = "00"
            theMins = "00"
        Else
            theDay = Left$(gblAlarm1Date, 2)
            theMonth = Mid$(gblAlarm1Date, 4, 2)
            theYear = Mid$(gblAlarm1Date, 7, 4)
            theHours = Left$(gblAlarm1Time, 2)
            theMins = Mid$(gblAlarm1Time, 4, 2)
        End If
    End If
        
    If fClock.alarmtoggle2Enabled = True Then
        ' read the date and time variables for alarm 2
        If gblAlarm2Date = "Alarm not yet set" Then
            theDay = "01"
            theMonth = "01"
            theYear = "1970"
            theHours = "00"
            theMins = "00"
        Else
            theDay = Left$(gblAlarm2Date, 2)
            theMonth = Mid$(gblAlarm2Date, 4, 2)
            theYear = Mid$(gblAlarm2Date, 7, 4)
            theHours = Left$(gblAlarm2Time, 2)
            theMins = Mid$(gblAlarm2Time, 4, 2)
        End If
    End If
    
    If fClock.alarmtoggle3Enabled = True Then
        ' read the date and time variables for alarm 3
        If gblAlarm3Date = "Alarm not yet set" Then
            theDay = "01"
            theMonth = "01"
            theYear = "1970"
            theHours = "00"
            theMins = "00"
        Else
            theDay = Left$(gblAlarm3Date, 2)
            theMonth = Mid$(gblAlarm3Date, 4, 2)
            theYear = Mid$(gblAlarm3Date, 7, 4)
            theHours = Left$(gblAlarm3Time, 2)
            theMins = Mid$(gblAlarm3Time, 4, 2)
        End If
    End If
    
    If fClock.alarmtoggle4Enabled = True Then
        ' read the date and time variables for alarm 4
        If gblAlarm4Date = "Alarm not yet set" Then
            theDay = "01"
            theMonth = "01"
            theYear = "1970"
            theHours = "00"
            theMins = "00"
        Else
            theDay = Left$(gblAlarm4Date, 2)
            theMonth = Mid$(gblAlarm4Date, 4, 2)
            theYear = Mid$(gblAlarm4Date, 7, 4)
            theHours = Left$(gblAlarm4Time, 2)
            theMins = Mid$(gblAlarm4Time, 4, 2)
        End If
    End If
    
    If fClock.alarmtoggle5Enabled = True Then
         If gblAlarm4Date = "Alarm not yet set" Then
            theDay = "01"
            theMonth = "01"
            theYear = "1970"
            theHours = "00"
            theMins = "00"
        Else       ' read the date and time variables for alarm 5
            theDay = Left$(gblAlarm5Date, 2)
            theMonth = Mid$(gblAlarm5Date, 4, 2)
            theYear = Mid$(gblAlarm5Date, 7, 4)
            theHours = Left$(gblAlarm5Time, 2)
            theMins = Mid$(gblAlarm5Time, 4, 2)
        End If
    End If
    
    'Call displayDigitalDayofMonth(CC, theDay)
    'Call displayDigitalMonth(CC, theMonth)
    'Call displayDigitalYear(CC, theYear)
'    Call displayDigitalHours(CC, theHours)
'    Call displayDigitalMins(CC, theMins)
    'Call drawAMPMDigitalDay(CC, theHours)
   
End Sub





' ----------------------------------------------------------------
' Procedure Name: displayDigitalHours
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Parameter CC (cCairoContext):
' Author: beededea
' Date: 24/06/2024
' ----------------------------------------------------------------
Private Sub displayDigitalHours(ByRef CC As cCairoContext, Optional ByVal alarmHour As String)
        
    Dim hourString As String: hourString = vbNullString
    Dim localHour As Integer: localHour = 0
    
    'debug
'    alarmHour = "12"
       
    If alarmHour <> vbNullString Then
        localHour = CInt(alarmHour)
        hourString = CStr(localHour)
    Else
        localHour = CInt(thisHour)
        hourString = CStr(thisHour)
    End If
    
    If localHour < 10 Then ' uses vb6 coerce a string to a numeric without calling a cast function, quicker?
        hourString = "0" & localHour '
    End If
    
    ' hour digits for digital display
    
    '  Cairo.ImageList.AddSurface "jan", Cairo.ImageList("month")
    
    Select Case Left$(hourString, 1)
        Case 0
            CC.RenderSurfaceContent "small0", 453, 266, , , , 1
        Case 1
            CC.RenderSurfaceContent "small1", 453, 266, , , , 1
        Case 2
            CC.RenderSurfaceContent "small2", 453, 266, , , , 1
        Case 3
            CC.RenderSurfaceContent "small3", 453, 266, , , , 1
        Case 4
            CC.RenderSurfaceContent "small4", 453, 266, , , , 1
        Case 5
            CC.RenderSurfaceContent "small5", 453, 266, , , , 1
        Case 6
            CC.RenderSurfaceContent "small6", 453, 266, , , , 1
        Case 7
            CC.RenderSurfaceContent "small7", 453, 266, , , , 1
        Case 8
            CC.RenderSurfaceContent "small8", 453, 266, , , , 1
        Case 9
            CC.RenderSurfaceContent "small9", 453, 266, , , , 1
    End Select
        
    Select Case Mid$(hourString, 2, 1)
        Case 0
            CC.RenderSurfaceContent "small0", 474, 266, , , , 1
        Case 1
            CC.RenderSurfaceContent "small1", 474, 266, , , , 1
        Case 2
            CC.RenderSurfaceContent "small2", 474, 266, , , , 1
        Case 3
            CC.RenderSurfaceContent "small3", 474, 266, , , , 1
        Case 4
            CC.RenderSurfaceContent "small4", 474, 266, , , , 1
        Case 5
            CC.RenderSurfaceContent "small5", 474, 266, , , , 1
        Case 6
            CC.RenderSurfaceContent "small6", 474, 266, , , , 1
        Case 7
            CC.RenderSurfaceContent "small7", 474, 266, , , , 1
        Case 8
            CC.RenderSurfaceContent "small8", 474, 266, , , , 1
        Case 9
            CC.RenderSurfaceContent "small9", 474, 266, , , , 1
    End Select
    

End Sub
' ----------------------------------------------------------------
' Procedure Name: displayDigitalMins
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Parameter CC (cCairoContext):
' Author: beededea
' Date: 24/06/2024
' ----------------------------------------------------------------
Private Sub displayDigitalMins(ByRef CC As cCairoContext, Optional ByVal alarmMins As String)

    Dim minString As String: minString = vbNullString
    Dim localMins As Integer: localMins = 0
    
    'debug
'    alarmMins = "37"
       
    If alarmMins <> vbNullString Then
        localMins = CInt(alarmMins)
        minString = CStr(localMins)
    Else
        localMins = CInt(thisMin)
        minString = CStr(thisMin)
    End If
    
    If localMins < 10 Then ' uses vb6 coerce a string to a numeric without calling a cast function, quicker?
        minString = "0" & localMins '
    End If

    ' Min digits for digital display
    
    Select Case Left$(minString, 1)
        Case 0
            CC.RenderSurfaceContent "small0", 502, 266, , , , 1
        Case 1
            CC.RenderSurfaceContent "small1", 502, 266, , , , 1
        Case 2
            CC.RenderSurfaceContent "small2", 502, 266, , , , 1
        Case 3
            CC.RenderSurfaceContent "small3", 502, 266, , , , 1
        Case 4
            CC.RenderSurfaceContent "small4", 502, 266, , , , 1
        Case 5
            CC.RenderSurfaceContent "small5", 502, 266, , , , 1
        Case 6
            CC.RenderSurfaceContent "small6", 502, 266, , , , 1
        Case 7
            CC.RenderSurfaceContent "small7", 502, 266, , , , 1
        Case 8
            CC.RenderSurfaceContent "small8", 502, 266, , , , 1
        Case 9
            CC.RenderSurfaceContent "small9", 502, 266, , , , 1
    End Select
        
    Select Case Mid$(minString, 2, 1)
        Case 0
            CC.RenderSurfaceContent "small0", 523, 266, , , , 1
        Case 1
            CC.RenderSurfaceContent "small1", 523, 266, , , , 1
        Case 2
            CC.RenderSurfaceContent "small2", 523, 266, , , , 1
        Case 3
            CC.RenderSurfaceContent "small3", 523, 266, , , , 1
        Case 4
            CC.RenderSurfaceContent "small4", 523, 266, , , , 1
        Case 5
            CC.RenderSurfaceContent "small5", 523, 266, , , , 1
        Case 6
            CC.RenderSurfaceContent "small6", 523, 266, , , , 1
        Case 7
            CC.RenderSurfaceContent "small7", 523, 266, , , , 1
        Case 8
            CC.RenderSurfaceContent "small8", 523, 266, , , , 1
        Case 9
            CC.RenderSurfaceContent "small9", 523, 266, , , , 1
    End Select
    

End Sub
' ----------------------------------------------------------------
' Procedure Name: displayDigitalMonth
' Purpose: month digits for digital display
' Procedure Kind: Sub
' Procedure Access: Private
' Parameter CC (cCairoContext):
' Author: beededea
' Date: 24/06/2024
' ----------------------------------------------------------------
Private Sub displayDigitalMonth(ByRef CC As cCairoContext, Optional ByVal alarmMonth As String)

    Dim thismonth As Integer: thismonth = 0

    thismonth = Month(Now) ' coerce date variant to integer
       
    If alarmMonth <> vbNullString Then
        thismonth = CInt(alarmMonth)
    Else
        thismonth = Month(Now) ' coerce date variant to integer
    End If

    Select Case thismonth
        Case 1
            CC.RenderSurfaceContent "jan", 442, 312, , , , 1
        Case 2
            CC.RenderSurfaceContent "feb", 442, 312, , , , 1
        Case 3
            CC.RenderSurfaceContent "mar", 442, 312, , , , 1
        Case 4
            CC.RenderSurfaceContent "apr", 442, 312, , , , 1
        Case 5
            CC.RenderSurfaceContent "may", 442, 312, , , , 1
        Case 6
            CC.RenderSurfaceContent "jun", 442, 312, , , , 1
        Case 7
            CC.RenderSurfaceContent "jul", 442, 312, , , , 1
        Case 8
            CC.RenderSurfaceContent "aug", 442, 312, , , , 1
        Case 9
            CC.RenderSurfaceContent "sep", 442, 312, , , , 1
        Case 10
            CC.RenderSurfaceContent "oct", 442, 312, , , , 1
        Case 11
            CC.RenderSurfaceContent "nov", 442, 312, , , , 1
        Case 12
            CC.RenderSurfaceContent "dec", 442, 312, , , , 1
    End Select
    
    
     

   On Error GoTo 0
   Exit Sub

displayDigitalMonth_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure displayDigitalMonth of Class Module cwOverlay"

End Sub


' ----------------------------------------------------------------
' Procedure Name: displayDigitalDayofMonth
' Purpose: ' day digits for digital display
' Procedure Kind: Sub
' Procedure Access: Private
' Parameter CC (cCairoContext):
' Author: beededea
' Date: 24/06/2024
' ----------------------------------------------------------------
Private Sub displayDigitalDayofMonth(ByRef CC As cCairoContext, Optional ByVal alarmDay As String)
    
    Dim dayString As String: dayString = vbNullString
    Dim localDay As Integer: localDay = 0
    Dim thisDay As Integer: thisDay = 0

    thisDay = Day(Now)
       
    If alarmDay <> vbNullString Then
        localDay = CInt(alarmDay)
        dayString = CStr(localDay)
    Else
        localDay = CInt(thisDay)
        dayString = CStr(thisDay)
    End If
    
    If localDay < 10 Then ' uses vb6 coerce a string to a numeric without calling a cast function, quicker?
        dayString = "0" & localDay '
    End If
    
    
    
    Select Case Left$(dayString, 1)
        Case 0
            CC.RenderSurfaceContent "num0large", 404, 312, , , , 1
        Case 1
            CC.RenderSurfaceContent "num1large", 404, 312, , , , 1
        Case 2
            CC.RenderSurfaceContent "num2large", 404, 312, , , , 1
        Case 3
            CC.RenderSurfaceContent "num3large", 404, 312, , , , 1
        Case 4
            CC.RenderSurfaceContent "num4large", 404, 312, , , , 1
        Case 5
            CC.RenderSurfaceContent "num5large", 404, 312, , , , 1
        Case 6
            CC.RenderSurfaceContent "num6large", 404, 312, , , , 1
        Case 7
            CC.RenderSurfaceContent "num7large", 404, 312, , , , 1
        Case 8
            CC.RenderSurfaceContent "num8large", 404, 312, , , , 1
        Case 9
            CC.RenderSurfaceContent "num9large", 404, 312, , , , 1
    End Select
    
    Select Case Mid$(dayString, 2, 1)
        Case 0
            CC.RenderSurfaceContent "num0large", 423, 312, , , , 1
        Case 1
            CC.RenderSurfaceContent "num1large", 423, 312, , , , 1
        Case 2
            CC.RenderSurfaceContent "num2large", 423, 312, , , , 1
        Case 3
            CC.RenderSurfaceContent "num3large", 423, 312, , , , 1
        Case 4
            CC.RenderSurfaceContent "num4large", 423, 312, , , , 1
        Case 5
            CC.RenderSurfaceContent "num5large", 423, 312, , , , 1
        Case 6
            CC.RenderSurfaceContent "num6large", 423, 312, , , , 1
        Case 7
            CC.RenderSurfaceContent "num7large", 423, 312, , , , 1
        Case 8
            CC.RenderSurfaceContent "num8large", 423, 312, , , , 1
        Case 9
            CC.RenderSurfaceContent "num9large", 423, 312, , , , 1
    End Select
End Sub


' ----------------------------------------------------------------
' Procedure Name: displayDigitalYear
' Purpose: year digits for digital display
' Procedure Kind: Sub
' Procedure Access: Private
' Parameter CC (cCairoContext):
' Author: beededea
' Date: 24/06/2024
' ----------------------------------------------------------------
Private Sub displayDigitalYear(ByRef CC As cCairoContext, Optional ByVal alarmYear As String)

    Dim yearString As String: yearString = vbNullString
    
    'alarmYear = "1942"
           
    If alarmYear <> vbNullString Then
        yearString = alarmYear
    Else
        yearString = Year(Now) ' coerce date variant to integer
    End If

    Select Case Left$(yearString, 1)
        Case 0
            CC.RenderSurfaceContent "num0large", 510, 312, , , , 1
        Case 1
            CC.RenderSurfaceContent "num1large", 510, 312, , , , 1
        Case 2
            CC.RenderSurfaceContent "num2large", 510, 312, , , , 1
        Case 3
            CC.RenderSurfaceContent "num3large", 510, 312, , , , 1
        Case 4
            CC.RenderSurfaceContent "num4large", 510, 312, , , , 1
        Case 5
            CC.RenderSurfaceContent "num5large", 510, 312, , , , 1
        Case 6
            CC.RenderSurfaceContent "num6large", 510, 312, , , , 1
        Case 7
            CC.RenderSurfaceContent "num7large", 510, 312, , , , 1
        Case 8
            CC.RenderSurfaceContent "num8large", 510, 312, , , , 1
        Case 9
            CC.RenderSurfaceContent "num9large", 510, 312, , , , 1
    End Select
    
    Select Case Mid$(yearString, 2, 1)
        Case 0
            CC.RenderSurfaceContent "num0large", 529, 312, , , , 1
        Case 1
            CC.RenderSurfaceContent "num1large", 529, 312, , , , 1
        Case 2
            CC.RenderSurfaceContent "num2large", 529, 312, , , , 1
        Case 3
            CC.RenderSurfaceContent "num3large", 529, 312, , , , 1
        Case 4
            CC.RenderSurfaceContent "num4large", 529, 312, , , , 1
        Case 5
            CC.RenderSurfaceContent "num5large", 529, 312, , , , 1
        Case 6
            CC.RenderSurfaceContent "num6large", 529, 312, , , , 1
        Case 7
            CC.RenderSurfaceContent "num7large", 529, 312, , , , 1
        Case 8
            CC.RenderSurfaceContent "num8large", 529, 312, , , , 1
        Case 9
            CC.RenderSurfaceContent "num9large", 529, 312, , , , 1
    End Select
        
    Select Case Mid$(yearString, 3, 1)
        Case 0
            CC.RenderSurfaceContent "num0large", 548, 312, , , , 1
        Case 1
            CC.RenderSurfaceContent "num1large", 548, 312, , , , 1
        Case 2
            CC.RenderSurfaceContent "num2large", 548, 312, , , , 1
        Case 3
            CC.RenderSurfaceContent "num3large", 548, 312, , , , 1
        Case 4
            CC.RenderSurfaceContent "num4large", 548, 312, , , , 1
        Case 5
            CC.RenderSurfaceContent "num5large", 548, 312, , , , 1
        Case 6
            CC.RenderSurfaceContent "num6large", 548, 312, , , , 1
        Case 7
            CC.RenderSurfaceContent "num7large", 548, 312, , , , 1
        Case 8
            CC.RenderSurfaceContent "num8large", 548, 312, , , , 1
        Case 9
            CC.RenderSurfaceContent "num9large", 548, 312, , , , 1
    End Select
    
    Select Case Mid$(yearString, 4, 1)
        Case 0
            CC.RenderSurfaceContent "num0large", 567, 312, , , , 1
        Case 1
            CC.RenderSurfaceContent "num1large", 567, 312, , , , 1
        Case 2
            CC.RenderSurfaceContent "num2large", 567, 312, , , , 1
        Case 3
            CC.RenderSurfaceContent "num3large", 567, 312, , , , 1
        Case 4
            CC.RenderSurfaceContent "num4large", 567, 312, , , , 1
        Case 5
            CC.RenderSurfaceContent "num5large", 567, 312, , , , 1
        Case 6
            CC.RenderSurfaceContent "num6large", 567, 312, , , , 1
        Case 7
            CC.RenderSurfaceContent "num7large", 567, 312, , , , 1
        Case 8
            CC.RenderSurfaceContent "num8large", 567, 312, , , , 1
        Case 9
            CC.RenderSurfaceContent "num9large", 567, 312, , , , 1
    End Select
End Sub



' ----------------------------------------------------------------
' Procedure Name: tmrTimeShift_Timer
' Purpose:  timer controlling the pendulum only,
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 21/06/2024
' ----------------------------------------------------------------
Private Sub tmrTimeShift_Timer()
    
    On Error GoTo tmrTimeShift_Timer_Error

    W.Parent.Refresh ' causes a paint event to fire
    
   On Error GoTo 0
   Exit Sub

tmrTimeShift_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrTimeShift_Timer of Class Module cwOverlay"
End Sub
' ----------------------------------------------------------------
' Procedure Name: tmrPendulum_Timer
' Purpose:  timer controlling the pendulum only,
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 21/06/2024
' ----------------------------------------------------------------
Private Sub tmrPendulum_Timer()
    
    On Error GoTo tmrPendulum_Timer_Error

    pendulumAcceleration = gravity * pendulumAngle
    pendulumVelocity = pendulumVelocity + pendulumAcceleration
    pendulumAngle = pendulumAngle + pendulumVelocity

    W.Parent.Refresh ' causes a paint event to fire
    
   On Error GoTo 0
   Exit Sub

tmrPendulum_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrPendulum_Timer of Class Module cwOverlay"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : tmrDigitRotator_Timer
' Author    : beededea
' Date      : 13/09/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub tmrDigitRotator_Timer()

   On Error GoTo tmrDigitRotator_Timer_Error

    ' which digit?
    ' get current value
    ' start at zero and increment to current value
    ' if the value is
    ' set the current displayed digit in the cache
    ' trigger a display
    
   On Error GoTo 0
   Exit Sub

tmrDigitRotator_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrDigitRotator_Timer of Class Module cwOverlay"
End Sub

' ----------------------------------------------------------------
' Procedure Name: tmrClock_Timer
' Purpose: main timer controlling the analogue clock only,
' very few routines are called from this timer but instead the subsequent refresh will trigger a PAINT event
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 21/06/2024
' ----------------------------------------------------------------
Private Sub tmrClock_Timer()
    
    On Error GoTo tmrClock_Timer_Error
    
    ' just for the analogue clock
    thisHour = Hour(Now)
    thisMin = Minute(Now)
    thisSec = Second(Now)
    
    ' calculate angles of hour and second hand
    hdeg = 30 * (thisHour Mod 12 + (thisMin) / 60) ' global vars used in the paint event
    mdeg = 6 * ((thisMin) + (thisSec / 60))
    Sec = thisSec

    If SDeg <> 6 * Sec Then
        SDeg = 6 * Sec
    Else
        Exit Sub  'only refresh when SDeg has changed to a new value
    End If
    
    ' only clear any existing weekday indicator in the first three seconds of the day
    If fClock.weekdayToggleEnabled = True Then
        If thisHour <= 1 And thisMin = 1 And thisSec < 3 Then
            fClock.clockForm.Widgets("monday").Widget.Alpha = 0
            fClock.clockForm.Widgets("tuesday").Widget.Alpha = 0
            fClock.clockForm.Widgets("wednesday").Widget.Alpha = 0
            fClock.clockForm.Widgets("thursday").Widget.Alpha = 0
            fClock.clockForm.Widgets("friday").Widget.Alpha = 0
            fClock.clockForm.Widgets("saturday").Widget.Alpha = 0
            fClock.clockForm.Widgets("sunday").Widget.Alpha = 0
        End If
        fClock.clockForm.Widgets(fDayOfWeek).Widget.Alpha = 1
    End If
    
        ' this should be called once then every hour on the hour, populate the cache
    If firstRun = True Then
        triggerDigitalClockPopulation = True
        firstRun = False
    End If
          
    ' very few routines are called from this timer but instead this subsequent refresh will trigger a PAINT event
    
    W.Parent.Refresh ' causes a paint event to fire - lots happens in the paint event
    
   On Error GoTo 0
   Exit Sub

tmrClock_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrClock_Timer of Class Module cwOverlay"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : DrawDigitalClocks
' Author    : beededea
' Date      : 06/09/2024
' Purpose   : display the hours and mins
'---------------------------------------------------------------------------------------
'
Sub DrawDigitalClocks(CC)
   On Error GoTo DrawDigitalClocks_Error
      
      
    CC.RenderSurfaceContent "hourCache1", 453, 266, , , , 1
    CC.RenderSurfaceContent "hourCache2", 474, 266, , , , 1
              
    CC.RenderSurfaceContent "minuteCache1", 502, 266, , , , 1
    CC.RenderSurfaceContent "minuteCache2", 523, 266, , , , 1
'
'    ' display the day, month and years
    CC.RenderSurfaceContent "dayCache1", 404, 312, , , , 1
    CC.RenderSurfaceContent "dayCache2", 423, 312, , , , 1
'
    CC.RenderSurfaceContent "monthCache", 442, 312, , , , 1

    CC.RenderSurfaceContent "yearCache1", 510, 312, , , , 1
    CC.RenderSurfaceContent "yearCache2", 529, 312, , , , 1
    CC.RenderSurfaceContent "yearCache3", 548, 312, , , , 1
    CC.RenderSurfaceContent "yearCache4", 567, 312, , , , 1
    
    CC.RenderSurfaceContent "AMPMCache", 621, 387, , , , 1

   On Error GoTo 0
   Exit Sub

DrawDigitalClocks_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DrawDigitalClocks of Class Module cwOverlay"
End Sub
    


' ----------------------------------------------------------------
' Procedure Name: populateDigitalHoursCache
' Purpose:
' Procedure Kind: Sub
' Procedure ACCess: Private
' Parameter CC (CCCContext):
' Author: beededea
' Date: 24/06/2024
' ----------------------------------------------------------------
Private Sub populateDigitalHoursCache(ByRef CC As cCairoContext, Optional ByVal alarmHour As String)

    Dim hourString As String: hourString = vbNullString
    Dim localHour As Integer: localHour = 0

   On Error GoTo populateDigitalHoursCache_Error

    If alarmHour <> vbNullString Then
        localHour = CInt(alarmHour)
        hourString = CStr(localHour)
    Else
        localHour = CInt(thisHour)
        hourString = CStr(thisHour)
    End If

    If localHour < 10 Then ' uses vb6 coerce a string to a numeric without calling a cast function, quicker?
        hourString = "0" & localHour '
    End If
    
    ' hour digits for digital display
    Select Case Left$(hourString, 1)
        Case 0
            Cairo.ImageList.AddSurface "hourCache1", Cairo.ImageList("small0")
        Case 1
            Cairo.ImageList.AddSurface "hourCache1", Cairo.ImageList("small1")
        Case 2
            Cairo.ImageList.AddSurface "hourCache1", Cairo.ImageList("small2")
        Case 3
            Cairo.ImageList.AddSurface "hourCache1", Cairo.ImageList("small3")
        Case 4
            Cairo.ImageList.AddSurface "hourCache1", Cairo.ImageList("small4")
        Case 5
            Cairo.ImageList.AddSurface "hourCache1", Cairo.ImageList("small5")
        Case 6
            Cairo.ImageList.AddSurface "hourCache1", Cairo.ImageList("small6")
        Case 7
            Cairo.ImageList.AddSurface "hourCache1", Cairo.ImageList("small7")
        Case 8
            Cairo.ImageList.AddSurface "hourCache1", Cairo.ImageList("small8")
        Case 9
            Cairo.ImageList.AddSurface "hourCache1", Cairo.ImageList("small9")
    End Select

    Select Case Mid$(hourString, 2, 1)
        Case 0
            Cairo.ImageList.AddSurface "hourCache2", Cairo.ImageList("small0")
        Case 1
            Cairo.ImageList.AddSurface "hourCache2", Cairo.ImageList("small1")
        Case 2
            Cairo.ImageList.AddSurface "hourCache2", Cairo.ImageList("small2")
        Case 3
            Cairo.ImageList.AddSurface "hourCache2", Cairo.ImageList("small3")
        Case 4
            Cairo.ImageList.AddSurface "hourCache2", Cairo.ImageList("small4")
        Case 5
            Cairo.ImageList.AddSurface "hourCache2", Cairo.ImageList("small5")
        Case 6
            Cairo.ImageList.AddSurface "hourCache2", Cairo.ImageList("small6")
        Case 7
            Cairo.ImageList.AddSurface "hourCache2", Cairo.ImageList("small7")
        Case 8
            Cairo.ImageList.AddSurface "hourCache2", Cairo.ImageList("small8")
        Case 9
            Cairo.ImageList.AddSurface "hourCache2", Cairo.ImageList("small9")
    End Select


   On Error GoTo 0
   Exit Sub

populateDigitalHoursCache_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure populateDigitalHoursCache of Class Module cwOverlay"

End Sub

' ----------------------------------------------------------------
' Procedure Name: populateDigitalMinsCache
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Parameter CC (cCairoContext):
' Author: beededea
' Date: 24/06/2024
' ----------------------------------------------------------------
Private Sub populateDigitalMinsCache(ByRef CC As cCairoContext, Optional ByVal alarmMins As String)

    Dim minString As String: minString = vbNullString
    Dim localMins As Integer: localMins = 0
       
    On Error GoTo populateDigitalMinsCache_Error

    If alarmMins <> vbNullString Then
        localMins = CInt(alarmMins)
        minString = CStr(localMins)
    Else
        localMins = CInt(thisMin)
        minString = CStr(thisMin)
    End If
    
    If localMins < 10 Then ' uses vb6 coerce a string to a numeric without calling a cast function, quicker?
        minString = "0" & localMins '
    End If

    ' Min digits for digital display
        
    Select Case Left$(minString, 1)
        Case 0
            Cairo.ImageList.AddSurface "minuteCache1", Cairo.ImageList("small0")
        Case 1
            Cairo.ImageList.AddSurface "minuteCache1", Cairo.ImageList("small1")
        Case 2
            Cairo.ImageList.AddSurface "minuteCache1", Cairo.ImageList("small2")
        Case 3
            Cairo.ImageList.AddSurface "minuteCache1", Cairo.ImageList("small3")
        Case 4
            Cairo.ImageList.AddSurface "minuteCache1", Cairo.ImageList("small4")
        Case 5
            Cairo.ImageList.AddSurface "minuteCache1", Cairo.ImageList("small5")
        Case 6
            Cairo.ImageList.AddSurface "minuteCache1", Cairo.ImageList("small6")
        Case 7
            Cairo.ImageList.AddSurface "minuteCache1", Cairo.ImageList("small7")
        Case 8
            Cairo.ImageList.AddSurface "minuteCache1", Cairo.ImageList("small8")
        Case 9
            Cairo.ImageList.AddSurface "minuteCache1", Cairo.ImageList("small9")
    End Select
        
    Select Case Mid$(minString, 2, 1)
        Case 0
            Cairo.ImageList.AddSurface "minuteCache2", Cairo.ImageList("small0")
        Case 1
            Cairo.ImageList.AddSurface "minuteCache2", Cairo.ImageList("small1")
        Case 2
            Cairo.ImageList.AddSurface "minuteCache2", Cairo.ImageList("small2")
        Case 3
            Cairo.ImageList.AddSurface "minuteCache2", Cairo.ImageList("small3")
        Case 4
            Cairo.ImageList.AddSurface "minuteCache2", Cairo.ImageList("small4")
        Case 5
            Cairo.ImageList.AddSurface "minuteCache2", Cairo.ImageList("small5")
        Case 6
            Cairo.ImageList.AddSurface "minuteCache2", Cairo.ImageList("small6")
        Case 7
            Cairo.ImageList.AddSurface "minuteCache2", Cairo.ImageList("small7")
        Case 8
            Cairo.ImageList.AddSurface "minuteCache2", Cairo.ImageList("small8")
        Case 9
            Cairo.ImageList.AddSurface "minuteCache2", Cairo.ImageList("small9")
    End Select
    

   On Error GoTo 0
   Exit Sub

populateDigitalMinsCache_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure populateDigitalMinsCache of Class Module cwOverlay"

End Sub



' ----------------------------------------------------------------
' Procedure Name: populateDigitalYearCache
' Purpose: year digits for digital display
' Procedure Kind: Sub
' Procedure Access: Private
' Parameter CC (cCairoContext):
' Author: beededea
' Date: 24/06/2024
' ----------------------------------------------------------------
Private Sub populateDigitalYearCache(ByRef CC As cCairoContext, Optional ByVal alarmYear As String)

    Dim yearString As String: yearString = vbNullString
           
    On Error GoTo populateDigitalYearCache_Error

    If alarmYear <> vbNullString Then
        yearString = alarmYear
    Else
        yearString = Year(Now) ' coerce date variant to integer
    End If

    Select Case Left$(yearString, 1)
        Case 0
            Cairo.ImageList.AddSurface "yearCache1", Cairo.ImageList("num0large")
        Case 1
            Cairo.ImageList.AddSurface "yearCache1", Cairo.ImageList("num1large")
        Case 2
            Cairo.ImageList.AddSurface "yearCache1", Cairo.ImageList("num2large")
        Case 3
            Cairo.ImageList.AddSurface "yearCache1", Cairo.ImageList("num3large")
        Case 4
            Cairo.ImageList.AddSurface "yearCache1", Cairo.ImageList("num4large")
        Case 5
            Cairo.ImageList.AddSurface "yearCache1", Cairo.ImageList("num5large")
        Case 6
            Cairo.ImageList.AddSurface "yearCache1", Cairo.ImageList("num6large")
        Case 7
            Cairo.ImageList.AddSurface "yearCache1", Cairo.ImageList("num7large")
        Case 8
            Cairo.ImageList.AddSurface "yearCache1", Cairo.ImageList("num8large")
        Case 9
            Cairo.ImageList.AddSurface "yearCache1", Cairo.ImageList("num9large")
    End Select
    
    Select Case Mid$(yearString, 2, 1)
        Case 0
            Cairo.ImageList.AddSurface "yearCache2", Cairo.ImageList("num0large")
        Case 1
            Cairo.ImageList.AddSurface "yearCache2", Cairo.ImageList("num1large")
        Case 2
            Cairo.ImageList.AddSurface "yearCache2", Cairo.ImageList("num2large")
        Case 3
            Cairo.ImageList.AddSurface "yearCache2", Cairo.ImageList("num3large")
        Case 4
            Cairo.ImageList.AddSurface "yearCache2", Cairo.ImageList("num4large")
        Case 5
            Cairo.ImageList.AddSurface "yearCache2", Cairo.ImageList("num5large")
        Case 6
            Cairo.ImageList.AddSurface "yearCache2", Cairo.ImageList("num6large")
        Case 7
            Cairo.ImageList.AddSurface "yearCache2", Cairo.ImageList("num7large")
        Case 8
            Cairo.ImageList.AddSurface "yearCache2", Cairo.ImageList("num8large")
        Case 9
            Cairo.ImageList.AddSurface "yearCache2", Cairo.ImageList("num9large")
    End Select
        
    Select Case Mid$(yearString, 3, 1)
        Case 0
            Cairo.ImageList.AddSurface "yearCache3", Cairo.ImageList("num0large")
        Case 1
            Cairo.ImageList.AddSurface "yearCache3", Cairo.ImageList("num1large")
        Case 2
            Cairo.ImageList.AddSurface "yearCache3", Cairo.ImageList("num2large")
        Case 3
            Cairo.ImageList.AddSurface "yearCache3", Cairo.ImageList("num3large")
        Case 4
            Cairo.ImageList.AddSurface "yearCache3", Cairo.ImageList("num4large")
        Case 5
            Cairo.ImageList.AddSurface "yearCache3", Cairo.ImageList("num5large")
        Case 6
            Cairo.ImageList.AddSurface "yearCache3", Cairo.ImageList("num6large")
        Case 7
            Cairo.ImageList.AddSurface "yearCache3", Cairo.ImageList("num7large")
        Case 8
            Cairo.ImageList.AddSurface "yearCache3", Cairo.ImageList("num8large")
        Case 9
            Cairo.ImageList.AddSurface "yearCache3", Cairo.ImageList("num9large")
    End Select
    
    Select Case Mid$(yearString, 4, 1)
        Case 0
            Cairo.ImageList.AddSurface "yearCache4", Cairo.ImageList("num0large")
        Case 1
            Cairo.ImageList.AddSurface "yearCache4", Cairo.ImageList("num1large")
        Case 2
            Cairo.ImageList.AddSurface "yearCache4", Cairo.ImageList("num2large")
        Case 3
            Cairo.ImageList.AddSurface "yearCache4", Cairo.ImageList("num3large")
        Case 4
            Cairo.ImageList.AddSurface "yearCache4", Cairo.ImageList("num4large")
        Case 5
            Cairo.ImageList.AddSurface "yearCache4", Cairo.ImageList("num5large")
        Case 6
            Cairo.ImageList.AddSurface "yearCache4", Cairo.ImageList("num6large")
        Case 7
            Cairo.ImageList.AddSurface "yearCache4", Cairo.ImageList("num7large")
        Case 8
            Cairo.ImageList.AddSurface "yearCache4", Cairo.ImageList("num8large")
        Case 9
            Cairo.ImageList.AddSurface "yearCache4", Cairo.ImageList("num9large")
    End Select

   On Error GoTo 0
   Exit Sub

populateDigitalYearCache_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure populateDigitalYearCache of Class Module cwOverlay"
End Sub

' ----------------------------------------------------------------
' Procedure Name: populateDigitalMonthCache
' Purpose: month digits for digital display
' Procedure Kind: Sub
' Procedure Access: Private
' Parameter CC (cCairoContext):
' Author: beededea
' Date: 24/06/2024
' ----------------------------------------------------------------
Private Sub populateDigitalMonthCache(ByRef CC As cCairoContext, Optional ByVal alarmMonth As String)

    Dim thismonth As Integer: thismonth = 0

    On Error GoTo populateDigitalMonthCache_Error

    thismonth = Month(Now) ' coerce date variant to integer
       
    If alarmMonth <> vbNullString Then
        thismonth = CInt(alarmMonth)
    Else
        thismonth = Month(Now) ' coerce date variant to integer
    End If

    Select Case thismonth
        Case 1
            Cairo.ImageList.AddSurface "monthCache", Cairo.ImageList("jan")
        Case 2
            Cairo.ImageList.AddSurface "monthCache", Cairo.ImageList("feb")
        Case 3
            Cairo.ImageList.AddSurface "monthCache", Cairo.ImageList("mar")
        Case 4
            Cairo.ImageList.AddSurface "monthCache", Cairo.ImageList("apr")
        Case 5
            Cairo.ImageList.AddSurface "monthCache", Cairo.ImageList("may")
        Case 6
            Cairo.ImageList.AddSurface "monthCache", Cairo.ImageList("jun")
        Case 7
            Cairo.ImageList.AddSurface "monthCache", Cairo.ImageList("jul")
        Case 8
            Cairo.ImageList.AddSurface "monthCache", Cairo.ImageList("aug")
        Case 9
            Cairo.ImageList.AddSurface "monthCache", Cairo.ImageList("sep")
        Case 10
            Cairo.ImageList.AddSurface "monthCache", Cairo.ImageList("oct")
        Case 11
            Cairo.ImageList.AddSurface "monthCache", Cairo.ImageList("nov")
        Case 12
            Cairo.ImageList.AddSurface "monthCache", Cairo.ImageList("dec")
    End Select
    
   On Error GoTo 0
   Exit Sub

populateDigitalMonthCache_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure populateDigitalMonthCache of Class Module cwOverlay"

End Sub



' ----------------------------------------------------------------
' Procedure Name: populateDigitalDayofMonthCache
' Purpose: ' day digits for digital display
' Procedure Kind: Sub
' Procedure Access: Private
' Parameter CC (cCairoContext):
' Author: beededea
' Date: 24/06/2024
' ----------------------------------------------------------------
Private Sub populateDigitalDayofMonthCache(ByRef CC As cCairoContext, Optional ByVal alarmDay As String)
    
    Dim dayString As String: dayString = vbNullString
    Dim localDay As Integer: localDay = 0
    Dim thisDay As Integer: thisDay = 0

   On Error GoTo populateDigitalDayofMonthCache_Error

    thisDay = Day(Now)
       
    If alarmDay <> vbNullString Then
        localDay = CInt(alarmDay)
        dayString = CStr(localDay)
    Else
        localDay = CInt(thisDay)
        dayString = CStr(thisDay)
    End If
    
    If localDay < 10 Then ' uses vb6 coerce a string to a numeric without calling a cast function, quicker?
        dayString = "0" & localDay '
    End If
        
    Select Case Left$(dayString, 1)
        Case 0
           Cairo.ImageList.AddSurface "dayCache1", Cairo.ImageList("num0large")
        Case 1
           Cairo.ImageList.AddSurface "dayCache1", Cairo.ImageList("num1large")
        Case 2
           Cairo.ImageList.AddSurface "dayCache1", Cairo.ImageList("num2large")
        Case 3
           Cairo.ImageList.AddSurface "dayCache1", Cairo.ImageList("num3large")
        Case 4
           Cairo.ImageList.AddSurface "dayCache1", Cairo.ImageList("num4large")
        Case 5
           Cairo.ImageList.AddSurface "dayCache1", Cairo.ImageList("num5large")
        Case 6
           Cairo.ImageList.AddSurface "dayCache1", Cairo.ImageList("num6large")
        Case 7
           Cairo.ImageList.AddSurface "dayCache1", Cairo.ImageList("num7large")
        Case 8
           Cairo.ImageList.AddSurface "dayCache1", Cairo.ImageList("num8large")
        Case 9
           Cairo.ImageList.AddSurface "dayCache1", Cairo.ImageList("num9large")
    End Select
    
    Select Case Mid$(dayString, 2, 1)
        Case 0
           Cairo.ImageList.AddSurface "dayCache2", Cairo.ImageList("num0large")
        Case 1
           Cairo.ImageList.AddSurface "dayCache2", Cairo.ImageList("num1large")
        Case 2
           Cairo.ImageList.AddSurface "dayCache2", Cairo.ImageList("num2large")
        Case 3
           Cairo.ImageList.AddSurface "dayCache2", Cairo.ImageList("num3large")
        Case 4
           Cairo.ImageList.AddSurface "dayCache2", Cairo.ImageList("num4large")
        Case 5
           Cairo.ImageList.AddSurface "dayCache2", Cairo.ImageList("num5large")
        Case 6
           Cairo.ImageList.AddSurface "dayCache2", Cairo.ImageList("num6large")
        Case 7
           Cairo.ImageList.AddSurface "dayCache2", Cairo.ImageList("num7large")
        Case 8
           Cairo.ImageList.AddSurface "dayCache2", Cairo.ImageList("num8large")
        Case 9
           Cairo.ImageList.AddSurface "dayCache2", Cairo.ImageList("num9large")
    End Select

   On Error GoTo 0
   Exit Sub

populateDigitalDayofMonthCache_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure populateDigitalDayofMonthCache of Class Module cwOverlay"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : populateAMPMDigitalDayCache
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub populateAMPMDigitalDayCache(ByRef CC As cCairoContext, Optional ByVal alarmHour As String)

    Dim localHour As Integer: localHour = 0
    
    On Error GoTo populateAMPMDigitalDayCache_Error

    If alarmHour <> vbNullString Then
        localHour = CInt(alarmHour)
    Else
        localHour = thisHour
    End If
    
    If localHour < 12 Then ' uses vb6 coerce a string to a numeric without calling a cast function, quicker?
         Cairo.ImageList.AddSurface "AMPMCache", Cairo.ImageList("AM")
    Else
         Cairo.ImageList.AddSurface "AMPMCache", Cairo.ImageList("PM")
    End If

   On Error GoTo 0
   Exit Sub

populateAMPMDigitalDayCache_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure populateAMPMDigitalDayCache of Class Module cwOverlay"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : populateDigitalClockCache
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub populateDigitalClockCache(ByRef CC As cCairoContext)
    Dim theDay As String: theDay = vbNullString
    Dim theMonth As String: theMonth = vbNullString
    Dim theYear As String: theYear = vbNullString
    Dim theHours As String: theHours = vbNullString
    Dim theMins As String: theMins = vbNullString
    
    On Error GoTo populateDigitalClockCache_Error

    If gblAlarmToggle1Enabled = "True" Then
        ' read the date and time variables for alarm 1
        If gblAlarm1Date = "Alarm not yet set" Then
            theDay = "01"
            theMonth = "01"
            theYear = "1970"
            theHours = "00"
            theMins = "00"
        Else
            theDay = Left$(gblAlarm1Date, 2)
            theMonth = Mid$(gblAlarm1Date, 4, 2)
            theYear = Mid$(gblAlarm1Date, 7, 4)
            theHours = Left$(gblAlarm1Time, 2)
            theMins = Mid$(gblAlarm1Time, 4, 2)
        End If
    End If
        
    If gblAlarmToggle2Enabled = "True" Then
        ' read the date and time variables for alarm 2
        If gblAlarm2Date = "Alarm not yet set" Then
            theDay = "01"
            theMonth = "01"
            theYear = "1970"
            theHours = "00"
            theMins = "00"
        Else
            theDay = Left$(gblAlarm2Date, 2)
            theMonth = Mid$(gblAlarm2Date, 4, 2)
            theYear = Mid$(gblAlarm2Date, 7, 4)
            theHours = Left$(gblAlarm2Time, 2)
            theMins = Mid$(gblAlarm2Time, 4, 2)
        End If
    End If
    
    If gblAlarmToggle3Enabled = "True" Then
        ' read the date and time variables for alarm 3
        If gblAlarm3Date = "Alarm not yet set" Then
            theDay = "01"
            theMonth = "01"
            theYear = "1970"
            theHours = "00"
            theMins = "00"
        Else
            theDay = Left$(gblAlarm3Date, 2)
            theMonth = Mid$(gblAlarm3Date, 4, 2)
            theYear = Mid$(gblAlarm3Date, 7, 4)
            theHours = Left$(gblAlarm3Time, 2)
            theMins = Mid$(gblAlarm3Time, 4, 2)
        End If
    End If
    
    If gblAlarmToggle4Enabled = "True" Then
        ' read the date and time variables for alarm 4
        If gblAlarm4Date = "Alarm not yet set" Then
            theDay = "01"
            theMonth = "01"
            theYear = "1970"
            theHours = "00"
            theMins = "00"
        Else
            theDay = Left$(gblAlarm4Date, 2)
            theMonth = Mid$(gblAlarm4Date, 4, 2)
            theYear = Mid$(gblAlarm4Date, 7, 4)
            theHours = Left$(gblAlarm4Time, 2)
            theMins = Mid$(gblAlarm4Time, 4, 2)
        End If
    End If
    
    If gblAlarmToggle5Enabled = "True" Then
         If gblAlarm4Date = "Alarm not yet set" Then
            theDay = "01"
            theMonth = "01"
            theYear = "1970"
            theHours = "00"
            theMins = "00"
        Else       ' read the date and time variables for alarm 5
            theDay = Left$(gblAlarm5Date, 2)
            theMonth = Mid$(gblAlarm5Date, 4, 2)
            theYear = Mid$(gblAlarm5Date, 7, 4)
            theHours = Left$(gblAlarm5Time, 2)
            theMins = Mid$(gblAlarm5Time, 4, 2)
        End If
    End If
    
    Call populateDigitalDayofMonthCache(CC, theDay)
    Call populateDigitalMonthCache(CC, theMonth)
    Call populateDigitalYearCache(CC, theYear)
    Call populateDigitalHoursCache(CC, theHours)
    Call populateDigitalMinsCache(CC, theMins)
    Call populateAMPMDigitalDayCache(CC, theHours)

   On Error GoTo 0
   Exit Sub

populateDigitalClockCache_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure populateDigitalClockCache of Class Module cwOverlay"
End Sub


