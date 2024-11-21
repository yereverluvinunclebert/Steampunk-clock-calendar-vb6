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

' public vars

Public cCollOverlayImages As cCollection 'collection made available publically, can't easily pass collections via property get/set

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1

Public sldCurWidth As Double ' need to add set/get for these public vars
Public cxPerc As Single      ' need to add set/get for these public vars
Public cyPerc As Single      ' need to add set/get for these public vars

' timer vars

Public WithEvents tmrClock As cTimer
Attribute tmrClock.VB_VarHelpID = -1
Public WithEvents tmrPendulum As cTimer
Attribute tmrPendulum.VB_VarHelpID = -1
Public WithEvents tmrTolling As cTimer
Attribute tmrTolling.VB_VarHelpID = -1
Public WithEvents tmrCountdownToToll As cTimer
Attribute tmrCountdownToToll.VB_VarHelpID = -1

Public WithEvents tmrTimeShift As cTimer    ' need to add set/get for these public timers
Attribute tmrTimeShift.VB_VarHelpID = -1
Public WithEvents tmrDigitRotator As cTimer ' need to add set/get for these
Attribute tmrDigitRotator.VB_VarHelpID = -1
Public WithEvents tmrAlarmRinging As cTimer ' need to add set/get for these public timers
Attribute tmrAlarmRinging.VB_VarHelpID = -1

' class member vars

Private mLocked As Boolean ' property Locked
Private mMyOpacity As Single ' property opacity
Private mTimeMachine As Boolean
Private mHidden As Boolean ' property Hidden
Private mSwingPendulum As Boolean

' private vars general usage

Private pvtHourDegrees As Double
Private pvtMinsDegrees As Double
Private pvtSecsDegrees As Double
Private pvtSeconds As Integer

Private pvtFirstRun As Boolean
Private pvtThisHour As Integer
Private pvtThisMins As Integer
Private pvtThisSecs As Integer

Private pvtPendulumAcceleration As Double
Private pvtGravity As Double
Private pvtPendulumAngle As Double
Private pvtPendulumVelocity As Double

Private pvtRotatorLimit As Integer
Private pvtMinsRotatorString As String
Private pvtHoursRotatorString As String
Private pvtDaysRotatorString As String
Private pvtMonthRotatorString As String
Private pvtYearRotatorString As String

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
    pvtFirstRun = True
    pvtRotatorLimit = 0
    
    ' variables set for the pendulum
    ' g = -0.1, a = 0.6, v = 2.5 uses 3% of cpu with a 40ms timer

    pvtGravity = -0.2
    pvtPendulumAcceleration = 0.6
    pvtPendulumVelocity = 3.5
    
    pvtMinsRotatorString = "00"
    pvtHoursRotatorString = "00"
    pvtDaysRotatorString = "00"
    pvtMonthRotatorString = "0"
    pvtYearRotatorString = "0000"

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
    Set tmrDigitRotator = New_c.Timer(50, False)
    
    tmrDigitRotator.Enabled = False
    
    Set tmrCountdownToToll = Nothing
    Set tmrCountdownToToll = New_c.Timer(2000, False)
    
    tmrCountdownToToll.Enabled = False
         
    Set tmrTolling = Nothing
    Set tmrTolling = New_c.Timer(2000, False)
    
    tmrTolling.Enabled = False
       
    Set tmrAlarmRinging = Nothing
    Set tmrAlarmRinging = New_c.Timer(1000, False)
    
    tmrAlarmRinging.Enabled = False
    
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Locked of Class Module cwoverlay"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Locked of Class Module cwoverlay"
End Property

'---------------------------------------------------------------------------------------
' Procedure : MyOpacity
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let MyOpacity(ByVal newValue As Double)
   On Error GoTo MyOpacityLet_Error

   If mMyOpacity <> newValue Then mMyOpacity = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

MyOpacityLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MyOpacity of Class Module cwAbout"
End Property
'---------------------------------------------------------------------------------------
' Procedure : MyOpacity
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get MyOpacity() As Double
   On Error GoTo MyOpacityGet_Error

   MyOpacity = mMyOpacity

   On Error GoTo 0
   Exit Property

MyOpacityGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MyOpacity of Class Module cwAbout"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property TimeMachine of Class Module cwoverlay"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property TimeMachine of Class Module cwoverlay"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Hidden of Class Module cwoverlay"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Hidden of Class Module cwoverlay"
End Property

 
'---------------------------------------------------------------------------------------
' Procedure : drawTimeMachineImages
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'ByRef CC As cCairoContext
Private Sub drawTimeMachineImages(ByRef CC As cCairoContext)
    
    



End Sub
 


'---------------------------------------------------------------------------------------
' Procedure : DrawDigitalClocks
' Author    : beededea
' Date      : 06/09/2024
' Purpose   : display the hours and mins
'---------------------------------------------------------------------------------------
'
Sub DrawDigitalClocks(ByRef CC As cCairoContext)
   On Error GoTo DrawDigitalClocks_Error
      
      
    CC.RenderSurfaceContent "hourCache1", 453, 266, , , , MyOpacity
    CC.RenderSurfaceContent "hourCache2", 474, 266, , , , MyOpacity
              
    CC.RenderSurfaceContent "minuteCache1", 502, 266, , , , MyOpacity
    CC.RenderSurfaceContent "minuteCache2", 523, 266, , , , MyOpacity
'
'    ' display the day, month and years
    CC.RenderSurfaceContent "dayCache1", 404, 312, , , , MyOpacity
    CC.RenderSurfaceContent "dayCache2", 423, 312, , , , MyOpacity
'
    CC.RenderSurfaceContent "monthCache", 442, 312, , , , MyOpacity

    CC.RenderSurfaceContent "yearCache1", 510, 312, , , , MyOpacity
    CC.RenderSurfaceContent "yearCache2", 529, 312, , , , MyOpacity
    CC.RenderSurfaceContent "yearCache3", 548, 312, , , , MyOpacity
    CC.RenderSurfaceContent "yearCache4", 567, 312, , , , MyOpacity
    
    CC.RenderSurfaceContent "AMPMCache", 621, 387, , , , MyOpacity

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
       
    ' if the rotator is functioning then display the contents from that...
    If tmrDigitRotator.Enabled = True Then
        localHour = CInt(pvtHoursRotatorString)
        hourString = pvtHoursRotatorString
    Else
        If alarmHour <> vbNullString Then
            localHour = CInt(alarmHour)
            hourString = CStr(localHour)
        Else
            localHour = CInt(pvtThisHour)
            hourString = CStr(pvtThisHour)
        End If
        
        ' if 24hr mode then convert
        If gbl24HourClockMode = "0" Then
            If localHour > 12 Then localHour = (localHour - 12)
        
            hourString = CStr(localHour)
       End If
    End If
        
    ' add a preceding zero to the time string
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

    ' if the rotator is functioning then display the contents from that...
    If tmrDigitRotator.Enabled = True Then
        localMins = CInt(pvtMinsRotatorString)
        minString = pvtHoursRotatorString
    Else
        If alarmMins <> vbNullString Then
            localMins = CInt(alarmMins)
            minString = CStr(localMins)
        Else
            localMins = CInt(pvtThisMins)
            minString = CStr(pvtThisMins)
        End If
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

    ' if the rotator is functioning then display the contents from that...
    If tmrDigitRotator.Enabled = True Then
        yearString = pvtYearRotatorString
    Else
        If alarmYear <> vbNullString Then
            yearString = alarmYear
        Else
            yearString = Year(Now) ' coerce date variant to integer
        End If
    End If
    
'    If Len(yearString) = 3 Then
'        MsgBox "hmmmm"
'    End If

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

    Dim thisMonth As Integer: thisMonth = 0

    On Error GoTo populateDigitalMonthCache_Error
    
    ' if the rotator is functioning then display the contents from that...
    If tmrDigitRotator.Enabled = True Then
        thisMonth = CInt(pvtMonthRotatorString)
    Else
        thisMonth = Month(Now) ' coerce date variant to integer
           
        If alarmMonth <> vbNullString Then
            thisMonth = CInt(alarmMonth)
        Else
            thisMonth = Month(Now) ' coerce date variant to integer
        End If
    End If

    Select Case thisMonth
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
   
    ' if the rotator is functioning then display the contents from that...
    If tmrDigitRotator.Enabled = True Then
        localDay = CInt(pvtDaysRotatorString)
        dayString = pvtDaysRotatorString
    Else
        thisDay = Day(Now)
           
        If alarmDay <> vbNullString Then
            localDay = CInt(alarmDay)
            dayString = CStr(localDay)
        Else
            localDay = CInt(thisDay)
            dayString = CStr(thisDay)
        End If
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

Private Sub populateAMPMDigitalDayCache(ByRef CC As cCairoContext, Optional ByVal alarmHour As String)

    Dim localHour As Integer: localHour = 0
    
    On Error GoTo populateAMPMDigitalDayCache_Error

    If alarmHour <> vbNullString Then
        localHour = CInt(alarmHour)
    Else
        localHour = pvtThisHour
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
Private Sub populateDigitalClockCache(ByRef CC As cCairoContext)
    Dim theDay As String: theDay = vbNullString
    Dim theMonth As String: theMonth = vbNullString
    Dim theYear As String: theYear = vbNullString
    Dim theHours As String: theHours = vbNullString
    Dim theMins As String: theMins = vbNullString
    
    On Error GoTo populateDigitalClockCache_Error
    
    If gblShowAlarmDateTime = True Then ' don't always want the stored dates/times to show when setting a new alarm ' replace with a property get/set?
        If gblAlarmToggle1Enabled = "True" Then
            ' read the date and time variables for alarmtoggle1
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
            ' read the date and time variables for alarmtoggle2
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
            ' read the date and time variables for alarmtoggle3
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
            ' read the date and time variables for alarmtoggle4
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
             If gblAlarm5Date = "Alarm not yet set" Then
                theDay = "01"
                theMonth = "01"
                theYear = "1970"
                theHours = "00"
                theMins = "00"
            Else       ' read the date and time variables for alarmtoggle5
                theDay = Left$(gblAlarm5Date, 2)
                theMonth = Mid$(gblAlarm5Date, 4, 2)
                theYear = Mid$(gblAlarm5Date, 7, 4)
                theHours = Left$(gblAlarm5Time, 2)
                theMins = Mid$(gblAlarm5Time, 4, 2)
            End If
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
'---------------------------------------------------------------------------------------
' Procedure : tmrDigitRotator_Timer
' Author    : beededea
' Date      : 13/09/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub tmrDigitRotator_Timer()
    Static rotatorTimerCount As Integer
    
    On Error GoTo tmrDigitRotator_Timer_Error
    
    pvtRotatorLimit = 24

    rotatorTimerCount = rotatorTimerCount + 1
    
    ' start at zero and increment a counter to current value
    ' if the value is equal to the counter, stop the timer
    
    If rotatorTimerCount <= pvtRotatorLimit Then
        ' set the current displayed digit in the cache
        ' we can use populateDigitalMinsCache passing a value as a param to alarmMins
        
        pvtMinsRotatorString = CStr(rotatorTimerCount)
        pvtHoursRotatorString = CStr(rotatorTimerCount)
        pvtDaysRotatorString = CStr(rotatorTimerCount)
        If rotatorTimerCount <= 9 Then
            pvtYearRotatorString = "190" & CStr(rotatorTimerCount)
        Else
            pvtYearRotatorString = "19" & CStr(rotatorTimerCount)
        End If
        pvtMonthRotatorString = CStr(rotatorTimerCount)

    Else
        tmrDigitRotator.Enabled = False
        rotatorTimerCount = 0
        
'        pvtMinsRotatorString = "00"
'        pvtHoursRotatorString = "00"
'        pvtDaysRotatorString = "00"
'        pvtYearRotatorString = "0000"
        
    End If
         
    ' trigger a paint event
        
    triggerDigitalClockPopulation = True
    overlayWidget.Widget.Parent.Refresh
   
   On Error GoTo 0
   Exit Sub

tmrDigitRotator_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrDigitRotator_Timer of Class Module cwOverlay"
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

    pvtPendulumAcceleration = pvtGravity * pvtPendulumAngle
    pvtPendulumVelocity = pvtPendulumVelocity + pvtPendulumAcceleration
    pvtPendulumAngle = pvtPendulumAngle + pvtPendulumVelocity

    W.Parent.Refresh ' causes a paint event to fire
    
   On Error GoTo 0
   Exit Sub

tmrPendulum_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrPendulum_Timer of Class Module cwOverlay"
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

    Dim fileToPlay As String: fileToPlay = vbNullString
    
    On Error GoTo tmrClock_Timer_Error

    ' just for the analogue clock
    pvtThisHour = Hour(Now)
    pvtThisMins = Minute(Now)
    pvtThisSecs = Second(Now)
    
    ' calculate angles of hour and second hand
    pvtHourDegrees = 30 * (pvtThisHour Mod 12 + (pvtThisMins) / 60) ' global vars used in the paint event
    pvtMinsDegrees = 6 * ((pvtThisMins) + (pvtThisSecs / 60))
    pvtSeconds = pvtThisSecs

    If pvtSecsDegrees <> 6 * pvtSeconds Then
        pvtSecsDegrees = 6 * pvtSeconds
    Else
        Exit Sub  'only refresh when pvtSecsDegrees has changed to a new value
    End If
    
    ' only clear any existing weekday indicator in the first three seconds of the day
    If fClock.weekdayToggleEnabled = True Then
        If pvtThisHour <= 1 And pvtThisMins = 1 And pvtThisSecs < 3 Then
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
    If pvtFirstRun = True Then
        triggerDigitalClockPopulation = True
        pvtFirstRun = False
    End If
    
    If fClock.clapperEnabled = True Then

        If pvtThisMins = 14 And pvtThisSecs = 58 Then ' 2 seconds before
            If gblVolumeBoost = "1" Then
                fileToPlay = "quarterchime.wav"
            Else
                fileToPlay = "quarterchime-quiet.wav"
            End If
            If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
                PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
            End If

           screenWrite ("chiming quarter past the hour at " & CStr(pvtThisHour) & ":15")

        ElseIf pvtThisMins = 29 And (pvtThisSecs = 54) Then ' 6 seconds before
    
            If gblVolumeBoost = "1" Then
                fileToPlay = "halfchime.wav"
            Else
                fileToPlay = "halfchime-quiet.wav"
            End If
            If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
                PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
            End If

            screenWrite ("chiming half past the hour at " & CStr(pvtThisHour) & ":30")

        ElseIf pvtThisMins = 45 And (pvtThisSecs = 51) Then '9 seconds before

            If gblVolumeBoost = "1" Then
                fileToPlay = "threequarterchime.wav"
            Else
                fileToPlay = "threequarterchime-quiet.wav"
            End If
            If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
                PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
            End If

            screenWrite ("chiming quarter to the hour at " & CStr(pvtThisHour) & ":45")

        ElseIf pvtThisMins = 59 And (pvtThisSecs = 42) Then       ' 18 seconds before the o'clock

            If gblVolumeBoost = "1" Then
                fileToPlay = "fullchime.wav"
            Else
                fileToPlay = "fullchime-quiet.wav"
            End If
            If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
                PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
            End If
            
            screenWrite ("chiming full hour ")
            
            tmrCountdownToToll.Enabled = True
            
        End If
    
        Call ShallWeRingAlarm
    End If
          
    ' very few routines are called from this timer but instead this subsequent refresh will trigger a PAINT event
    W.Parent.Refresh ' causes a paint event to fire - lots happens in the paint event
    
   On Error GoTo 0
   Exit Sub

tmrClock_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrClock_Timer of Class Module cwOverlay"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ShallWeRingAlarm
' Author    : beededea
' Date      : 19/10/2024
' Purpose   : check the alarms and raise a flag
'---------------------------------------------------------------------------------------
'
Private Sub ShallWeRingAlarm()
    
    Dim lngSecondsGap As Double: lngSecondsGap = 0

    On Error GoTo ShallWeRingAlarm_Error
    
    ' if the clapper has been moved aside then the alarms are silenced
    If gblEnableChimes = "0" Then Exit Sub

    ' if an alarm is currently being raised then the function is immediately exited
    ' ie. don't raise a second alarm
    If gblAlarmFlgRaised = True Then Exit Sub
    
    ' no alarms even entered in the prefs so exit
    If gblAlarm1 = "" And gblAlarm2 = "" And gblAlarm3 = "" And gblAlarm4 = "" And gblAlarm5 = "" Then Exit Sub
    
    If IsDate(gblAlarm1) And gblAlarm1 <> "00:00:00" And gblAlarm1FlgRaised = False Then
       
       'if current time is beyond the alarm time
       lngSecondsGap = DateDiff("s", Now(), gblAlarm1)
       
       If lngSecondsGap <= 0 Then
       
          ' prevent another alarm being raised
            gblAlarmFlgRaised = True
            gblAlarm1FlgRaised = True
            
            If fClock.ShowHelp = True Then fClock.ShowHelp = False
            
            ' raise the alarm till
            fClock.alarmtoggle1Enabled = True
            
            screenWrite ("Alarm 1 Raised! " & CStr(gblAlarm1))
            screenWrite ("Press alarm till to cancel raised alarm ")
            
            ' the alarm entries are removed here and written to the config straight away, preventing another accidental alarm being raised.
            
            widgetPrefs.txtAlarm1Date.Text = "Alarm not yet set"
            widgetPrefs.txtAlarm1Time.Text = ""
            If widgetPrefs.IsVisible = True Then
                widgetPrefs.txtAlarm1Date.Refresh
                widgetPrefs.txtAlarm1Time.Refresh
            End If
            
            gblAlarm1Date = vbNullString
            gblAlarm1Time = vbNullString
            
            sPutINISetting "Software\SteampunkClockCalendar", "alarm1Date", gblAlarm1Date, gblSettingsFile
            sPutINISetting "Software\SteampunkClockCalendar", "alarm1Time", gblAlarm1Time, gblSettingsFile
          
            ' start the alarm timer ringing the bell
            tmrAlarmRinging.Enabled = True
          
            Exit Sub
        End If
    End If
    
    If IsDate(gblAlarm2) And gblAlarm2 <> "00:00:00" And gblAlarm2FlgRaised = False Then
       
       'if current time is beyond the alarm time
       lngSecondsGap = DateDiff("s", Now(), gblAlarm2)
       
       If lngSecondsGap <= 0 Then
       
          ' prevent another alarm being raised
            gblAlarmFlgRaised = True
            gblAlarm2FlgRaised = True
            
            If fClock.ShowHelp = True Then fClock.ShowHelp = False
            
            ' raise the alarm till
            fClock.alarmtoggle2Enabled = True
            
            screenWrite ("Alarm 2 Raised! " & CStr(gblAlarm2))
            screenWrite ("Press alarm till to cancel raised alarm ")
            
            ' the alarm entries are removed here and written to the config straight away, preventing another accidental alarm being raised.
            
            widgetPrefs.txtAlarm2Date.Text = "Alarm not yet set"
            widgetPrefs.txtAlarm2Time.Text = ""
            If widgetPrefs.IsVisible = True Then
                widgetPrefs.txtAlarm2Date.Refresh
                widgetPrefs.txtAlarm2Time.Refresh
            End If
            
            gblAlarm2Date = vbNullString
            gblAlarm2Time = vbNullString
            
            sPutINISetting "Software\SteampunkClockCalendar", "alarm2Date", gblAlarm2Date, gblSettingsFile
            sPutINISetting "Software\SteampunkClockCalendar", "alarm2Time", gblAlarm2Time, gblSettingsFile
          
            ' start the alarm timer ringing the bell
            tmrAlarmRinging.Enabled = True
          
            Exit Sub
        End If
    End If
    
        
    If IsDate(gblAlarm3) And gblAlarm3 <> "00:00:00" And gblAlarm3FlgRaised = False Then
       
       'if current time is beyond the alarm time
       lngSecondsGap = DateDiff("s", Now(), gblAlarm3)
       
       If lngSecondsGap <= 0 Then
       
          ' prevent another alarm being raised
            gblAlarmFlgRaised = True
            gblAlarm3FlgRaised = True
            
            If fClock.ShowHelp = True Then fClock.ShowHelp = False
            
            ' raise the alarm till
            fClock.alarmtoggle3Enabled = True
            
            screenWrite ("Alarm 3 Raised! " & CStr(gblAlarm3))
            screenWrite ("Press alarm till to cancel raised alarm ")
            
            ' the alarm entries are removed here and written to the config straight away, preventing another accidental alarm being raised.
            
            widgetPrefs.txtAlarm3Date.Text = "Alarm not yet set"
            widgetPrefs.txtAlarm3Time.Text = ""
            If widgetPrefs.IsVisible = True Then
                widgetPrefs.txtAlarm3Date.Refresh
                widgetPrefs.txtAlarm3Time.Refresh
            End If
            
            gblAlarm3Date = vbNullString
            gblAlarm3Time = vbNullString
            
            sPutINISetting "Software\SteampunkClockCalendar", "alarm3Date", gblAlarm3Date, gblSettingsFile
            sPutINISetting "Software\SteampunkClockCalendar", "alarm3Time", gblAlarm3Time, gblSettingsFile
          
            ' start the alarm timer ringing the bell
            tmrAlarmRinging.Enabled = True
          
            Exit Sub
        End If
    End If
    
    If IsDate(gblAlarm4) And gblAlarm4 <> "00:00:00" And gblAlarm4FlgRaised = False Then
       
       'if current time is beyond the alarm time
       lngSecondsGap = DateDiff("s", Now(), gblAlarm4)
       
       If lngSecondsGap <= 0 Then
       
          ' prevent another alarm being raised
            gblAlarmFlgRaised = True
            gblAlarm4FlgRaised = True
            
            If fClock.ShowHelp = True Then fClock.ShowHelp = False
            
            ' raise the alarm till
            fClock.alarmtoggle4Enabled = True
            
            screenWrite ("Alarm 4 Raised! " & CStr(gblAlarm4))
            screenWrite ("Press alarm till to cancel raised alarm ")
            
            ' the alarm entries are removed here and written to the config straight away, preventing another accidental alarm being raised.
            
            widgetPrefs.txtAlarm4Date.Text = "Alarm not yet set"
            widgetPrefs.txtAlarm4Time.Text = ""
            If widgetPrefs.IsVisible = True Then
                widgetPrefs.txtAlarm4Date.Refresh
                widgetPrefs.txtAlarm4Time.Refresh
            End If
            
            gblAlarm4Date = vbNullString
            gblAlarm4Time = vbNullString
            
            sPutINISetting "Software\SteampunkClockCalendar", "alarm4Date", gblAlarm4Date, gblSettingsFile
            sPutINISetting "Software\SteampunkClockCalendar", "alarm4Time", gblAlarm4Time, gblSettingsFile
          
            ' start the alarm timer ringing the bell
            tmrAlarmRinging.Enabled = True
          
            Exit Sub
        End If
    End If
        
    If IsDate(gblAlarm5) And gblAlarm5 <> "00:00:00" And gblAlarm5FlgRaised = False Then
       
       'if current time is beyond the alarm time
       lngSecondsGap = DateDiff("s", Now(), gblAlarm5)
       
       If lngSecondsGap <= 0 Then
       
          ' prevent another alarm being raised
            gblAlarmFlgRaised = True
            gblAlarm5FlgRaised = True
            
            If fClock.ShowHelp = True Then fClock.ShowHelp = False
            
            ' raise the alarm till
            fClock.alarmtoggle5Enabled = True
            
            screenWrite ("Alarm 5 Raised! " & CStr(gblAlarm5))
            screenWrite ("Press alarm till to cancel raised alarm ")
            
            ' the alarm entries are removed here and written to the config straight away, preventing another accidental alarm being raised.
            
            widgetPrefs.txtAlarm5Date.Text = "Alarm not yet set"
            widgetPrefs.txtAlarm5Time.Text = ""
            If widgetPrefs.IsVisible = True Then
                widgetPrefs.txtAlarm5Date.Refresh
                widgetPrefs.txtAlarm5Time.Refresh
            End If
            
            gblAlarm5Date = vbNullString
            gblAlarm5Time = vbNullString
            
            sPutINISetting "Software\SteampunkClockCalendar", "alarm5Date", gblAlarm5Date, gblSettingsFile
            sPutINISetting "Software\SteampunkClockCalendar", "alarm5Time", gblAlarm5Time, gblSettingsFile
          
            ' start the alarm timer ringing the bell
            tmrAlarmRinging.Enabled = True
          
            Exit Sub
        End If
    End If
                
   On Error GoTo 0
   Exit Sub

ShallWeRingAlarm_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ShallWeRingAlarm of Class Module cwOverlay"
 End Sub


'---------------------------------------------------------------------------------------
' Procedure : tmrAlarmRinging_Timer
' Author    : beededea
' Date      : 18/10/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub tmrAlarmRinging_Timer()

   On Error GoTo tmrAlarmRinging_Timer_Error
   
    Dim fileToPlay As String: fileToPlay = vbNullString

       If gblVolumeBoost = "1" Then
           fileToPlay = "alarmbells.wav"
       Else
           fileToPlay = "alarmbells-quiet.wav"
       End If
       If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
           PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
       End If

   On Error GoTo 0
   Exit Sub

tmrAlarmRinging_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrAlarmRinging_Timer of Class Module cwOverlay"

End Sub


   
'---------------------------------------------------------------------------------------
' Procedure : tmrCountdownToToll_Timer
' Author    : beededea
' Date      : 18/10/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub tmrCountdownToToll_Timer()
    Static countdownToToll As Integer

    On Error GoTo tmrCountdownToToll_Timer_Error

    If (countdownToToll <= 6) Then
        countdownToToll = countdownToToll + 1
        screenWrite ("Countdown to tolling the hour =  " & 8 - countdownToToll)
    Else
        countdownToToll = 0
        tmrCountdownToToll.Enabled = False
        tmrTolling.Enabled = True
    End If

   On Error GoTo 0
   Exit Sub

tmrCountdownToToll_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrCountdownToToll_Timer of Class Module cwOverlay"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : tmrTolling_Timer
' Author    : beededea
' Date      : 18/10/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub tmrTolling_Timer()

    On Error GoTo tmrTolling_Timer_Error

     Call tollbell

   On Error GoTo 0
   Exit Sub

tmrTolling_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrTolling_Timer of Class Module cwOverlay"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : tollbell
' Author    : beededea
' Date      : 18/10/2024
' Purpose   : function to toll the bell x no. of times, called by the tmrTolling
'---------------------------------------------------------------------------------------
'
Private Sub tollbell()
    Dim fileToPlay As String: fileToPlay = vbNullString
    Dim tollhours As Integer: tollhours = 0
    Dim timeString As String: timeString = vbNullString
    
    Static tollCount As Integer
    
    On Error GoTo tollbell_Error

    tollhours = Hour(Now)
    
    If tollhours = 0 Then
        tollhours = 12
        timeString = " midnight"
    Else
        timeString = " o'clock"
    End If
    
    'stop it striking more than 12 times even when on a 24hr clock
    If tollhours > 12 Then tollhours = tollhours - 12
     
    ' do not ask for whom the bell tolls, it tolls for thee...when the counter reaches 8
    tollCount = tollCount + 1

    'play the longer bell on the last toll
    If (tollCount < tollhours) Then
        If gblVolumeBoost = "1" Then
            fileToPlay = "belltoll.wav"
        Else
            fileToPlay = "belltoll-quiet.wav"
        End If
    Else
        If gblVolumeBoost = "1" Then
            fileToPlay = "belltollLong.wav"
        Else
            fileToPlay = "belltollLong-quiet.wav"
        End If
    End If
    
    If tollCount = 1 Then screenWrite ("Chiming using " & fileToPlay & " at " & Now())
    
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    screenWrite ("Tolling bell for " & tollhours & " o'clock, chime no." & tollCount)
             
    ' turn off the belltimer to stop the bell from tolling
    If (tollCount >= tollhours) Then
        tmrTolling.Enabled = False
        tollCount = 0
        screenWrite ("Just struck " & tollhours & timeString)
     End If

   On Error GoTo 0
   Exit Sub

tollbell_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tollbell of Class Module cwOverlay"
End Sub



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

Private Sub W_Paint(ByRef CC As cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
    Dim I As Long: I = 0
    Dim Arr As Variant
    Dim Key As String: Key = vbNullString
    Dim Pat As cCairoPattern

    W.Refresh
    
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
                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(4, 30) 'shift to the "rotation-point" within the Pattern
                    CC.TranslateDrawings 645, 333.5 'shift to the center-point of the Overlay-Widget-area

                    CC.RotateDrawingsDeg pvtHourDegrees   'now we can rotate (the leading negative degrees are the rotation-offsets for 12 o'clock)

                 Case "minutehand"  'minutes-shadow- and minutes-hand
                    Set Pat = Cairo.ImageList(Key).CreateSurfacePattern
                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(3, 41) 'shift to the "rotation-point" within the Pattern
                    CC.TranslateDrawings 645, 333.5 'shift to the center-point of the Overlay-Widget-area

                    CC.RotateDrawingsDeg pvtMinsDegrees 'now we can rotate (the leading negative degrees are the rotation-offsets for 12 o'clock)

                 Case "secondhand" 'seconds-shadow- and seconds-hand
                    Set Pat = Cairo.ImageList(Key).CreateSurfacePattern
                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(3, 42) 'shift to the "rotation-point" within the Pattern
                    CC.TranslateDrawings 645, 333.5 'shift to the center-point of the Overlay-Widget-area

                    CC.RotateDrawingsDeg pvtSecsDegrees 'now we can rotate (the leading negative degrees are the rotation-offsets for 12 o'clock)

                 Case "hole"
                    Set Pat = Cairo.ImageList(Key).CreateSurfacePattern
                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(4, 4) 'shift to the "rotation-point" within the Pattern

                    CC.TranslateDrawings 645, 333.5 'shift to the center-point of the Overlay-Widget-area

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

                    CC.RotateDrawingsDeg pvtPendulumAngle 'now we can rotate (the leading negative degrees are the rotation-offsets for 12 o'clock)
                
                Case Else 'plain output for everything else
                       
                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(-Arr(0), -Arr(1)) 'here we only shift the Pattern via the original PSD-Offsets (stored in Arr(0) and Arr(1))
                    
             End Select
    
            CC.Paint W.AlphaInherited * Arr(2), Pat  'finally we render the Pattern with the current PSD-LayerAlpha, present in Arr(2)
            CC.Restore
      End If
    Next
    
    ' if first time running then populate the digital clocks.
    If triggerDigitalClockPopulation = True Then
        Call populateDigitalClockCache(CC)
        triggerDigitalClockPopulation = False
    Else ' update the correct digital display when a time change occurs
        If pvtThisSecs <= 1 Then
            Call populateDigitalMinsCache(CC)
            If pvtThisMins = 0 Then
                Call populateDigitalHoursCache(CC)
                If pvtThisHour = 0 Then
                    Call populateDigitalClockCache(CC)
                End If
            End If
        End If
    End If
        
    ' if one of the alarms is currently being viewed then populate the clocks with the alarm digits
    ' test for alarms, using globals to reduce the cpu cycles, if found, then populate the cache images
    If gblAlarmToggle1Enabled = "True" Or gblAlarmToggle2Enabled = "True" Or gblAlarmToggle3Enabled = "True" Or gblAlarmToggle4Enabled = "True" Or gblAlarmToggle5Enabled = "True" Then Call populateDigitalClockCache(CC)
    
    ' the above routines populate the digital clock cache, this routine shows the digital clocks
    Call DrawDigitalClocks(CC)
     
    If mTimeMachine = True Then Call drawTimeMachineImages(CC)

   On Error GoTo 0
   Exit Sub

W_Paint_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_Paint of Class Module cwOverlay"
End Sub
