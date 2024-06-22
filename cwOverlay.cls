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
' Purpose   : create a fully click-through Widget, which allows for all kind of
'             things to render on top of the PSD-generated Widgets created in cfClock
'---------------------------------------------------------------------------------------

Option Explicit

'these Public Properties are set in the Parent (here, cfClock)
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
Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1

Private firstRun As Boolean
Private thisHour As Integer
Private thisMin As Integer
Private thisSec As Integer

Private mTicking As Boolean ' property Ticking
Private mLocked As Boolean ' property Locked
Private mHidden As Boolean ' property Hidden

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
        W.BackColor = -1  'we don't use any BackColor here, transparent
        W.ImplementsHitTest = True 'when at False, no HitTest-Event would be triggered

    intervalValue = 1000
    firstRun = True
    
    Set tmrClock = Nothing
    Set tmrClock = New_c.Timer(intervalValue, True)
    
    tmrClock.Enabled = True
            
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

    
''---------------------------------------------------------------------------------------
'' Procedure : tmrClock_Timer
'' Author    :
'' Date      : 30/07/2023
'' Purpose   : update the degree-Variables for the clock Hour, Minute and Second Hands
''---------------------------------------------------------------------------------------
''
'Private Sub tmrClock_Timer()
'
'    Static T As Date: T = Time
'
'    On Error GoTo tmrClock_Timer_Error
'
'    ' time calc for main gauge
'    hdeg = 30 * (Hour(T) Mod 12 + (Minute(T)) / 60)
'    mdeg = 6 * ((Minute(T)) + (Second(T) / 60))
'
'    Sec = Second(T)
'
'    If SDeg <> 6 * Sec Then
'        SDeg = 6 * Sec
'    Else
'        Exit Sub  'only refresh when SDeg has changed to a new value
'    End If
'
'    W.Parent.Refresh ' causes a paint event to fire
'
'   On Error GoTo 0
'   Exit Sub
'
'tmrClock_Timer_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrClock_Timer of Class Module cwOverlay"
'End Sub




' ----------------------------------------------------------------
' Procedure Name: tmrClock_Timer
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 21/06/2024
' ----------------------------------------------------------------
Private Sub tmrClock_Timer()


'    Static thisHour As Integer
'    Static thisMin As Integer
'    Static thisSec As Integer
    
    On Error GoTo tmrClock_Timer_Error
    
    thisHour = Hour(Now)
    thisMin = Minute(Now)
    thisSec = Second(Now)
    
    ' time calc for main gauge
    hdeg = 30 * (thisHour Mod 12 + (thisMin) / 60)
    mdeg = 6 * ((thisMin) + (thisSec / 60))
    
    Sec = thisSec

    If SDeg <> 6 * Sec Then
        SDeg = 6 * Sec
    Else
        Exit Sub  'only refresh when SDeg has changed to a new value
    End If



    
    W.Parent.Refresh ' causes a paint event to fire
    
   On Error GoTo 0
   Exit Sub

tmrClock_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrClock_Timer of Class Module cwOverlay"
End Sub
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
Private Sub W_HitTest(ByVal x As Single, ByVal Y As Single, HitResultHit As Boolean) 'ensure ClickThrough-behaviour in ImagePixels which are "fully Alpha"
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
       
    On Error GoTo W_Paint_Error
  
    For I = 0 To cCollOverlayImages.Count - 1 ' this is the ExcludedPaths collection populated in cfClock by duplicating collPSDNonUIElements.
      Arr = cCollOverlayImages.ItemByIndex(I)
      Key = cCollOverlayImages.KeyByIndex(I)
      
      If IsArray(Arr) Then
             CC.Save
             'derive a Pattern-Obj from each retrieved ImageList-Surface-reference
             Set Pat = Cairo.ImageList(Key).CreateSurfacePattern
            
            ' for specific keys perform actions &c
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
                    Set Pat = Cairo.ImageList(Key).CropSurface(10, 0, 140, 27).CreateSurfacePattern 'cut-out a nice (repeatable) part of the original "cable"-image (via CropSurface)

                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(sldCurWidth, -Arr(1)) 'sldCurWidth in the x-Coord will ensure that nice "spooling-effect"
                        Pat.Extend = CAIRO_EXTEND_REPEAT

                    'since we've set CAIRO_EXTEND_REPEAT, we now have to clip the range it is "allowed to extend to"
                    CC.ClipExplicit Arr(0) + 10 - sldCurWidth, Arr(1), 249 + sldCurWidth, 27

                Case Else 'plain output for everything else ... this includes the reflection-layers,
                       
                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(-Arr(0), -Arr(1)) 'here we only shift the Pattern via the original PSD-Offsets (stored in Arr(0) and Arr(1))
                    
             End Select
            CC.Paint W.AlphaInherited * Arr(2), Pat  'finally we render the Pattern with the current PSD-LayerAlpha, present in Arr(2)
            CC.Restore
      End If
     Next
     
    Draw CC, dx_Aligned, dy_Aligned

   On Error GoTo 0
   Exit Sub

W_Paint_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_Paint of Class Module cwOverlay"
End Sub
 
 
'---------------------------------------------------------------------------------------
' Procedure : Draw
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Draw(ByRef CC As cCairoContext, ByVal dx As Single, ByVal dy As Single)

    Static hourString As String
    Static minString As String
    
    On Error GoTo Draw_Error
    
    ' only check whether to change the minute within the last second of the minute
    ' due to CPU usage and delay, we use a 3 sec window at the top of the minute
    Call displayDigitalDayofMonth(CC)
    Call displayDigitalYear(CC)
     
'    If (thisSec >= 59 Or thisSec <= 1) Or firstRun = True Then
'       ' min digits for digital display
'        Debug.Print "thissec - " & thisSec
'
'        If thisMin < 10 Then
'            minString = "0" & thisMin ' uses vb6 coerce a string to a numeric without calling a cast function, quicker?
'        Else
'            minString = thisMin
'        End If
'
''        Call showMin1(CC, Left$(minString, 1))
''        Call showMin2(Mid$(minString, 2, 1))
'
'        ' only check the hour within the last second of the 59th minute
'        If thisMin > 59 Or firstRun = True Then
'
'            ' hour digits for digital display
'
'            If thisHour < 10 Then ' uses vb6 coerce a string to a numeric without calling a cast function, quicker?
'                hourString = "0" & thisHour '
'            Else
'                hourString = thisHour
'            End If
'
''            Call showHour1(Left$(hourString, 1))
''            Call showHour2(Mid$(hourString, 2, 1))
'
'            ' only check the month and year within the last second of the day
'            If thisHour > 23 Or firstRun = True Then
'
'                Call displayDigitalDay(CC)
''                Call displayDigitalYear
'                firstRun = False
'            End If
'        End If
'    End If
    
    CC.RenderSurfaceContent "jan", 442, 312, , , , 1
'    CC.RenderSurfaceContent "num2large", 423, 310, , , , 1
'
'    CC.RenderSurfaceContent "num1large", 511, 310, , , , 1
'    CC.RenderSurfaceContent "num8large", 530, 310, , , , 1
'    CC.RenderSurfaceContent "num8large", 549, 310, , , , 1
'    CC.RenderSurfaceContent "num9large", 568, 310, , , , 1

    W.Parent.Refresh
    
   On Error GoTo 0
   Exit Sub

Draw_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Draw of Class Module cwLicence"

End Sub



Private Sub displayDigitalDayofMonth(ByRef CC As cCairoContext)

'    ' day digits for digital display

    Dim dDigit1 As String: dDigit1 = vbNullString
    Dim dDigit2 As String: dDigit2 = vbNullString
    Dim dayString As String: dayString = vbNullString
    Dim thisDay As Integer: thisDay = 0

    thisDay = Day(Now)

    If thisDay < 10 Then
        dayString = "0" & thisDay ' uses vb6 coerce a variant to a numeric without calling a cast function, quicker?
    Else
        dayString = thisDay
    End If

    dDigit1 = Left$(dayString, 1)
    dDigit2 = Mid$(dayString, 2, 1)
'
    Select Case dDigit1
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
    
    Select Case dDigit2
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

Private Sub displayDigitalYear(ByRef CC As cCairoContext)

    Dim yDigit1 As String: yDigit1 = vbNullString
    Dim yDigit2 As String: yDigit2 = vbNullString
    Dim yDigit3 As String: yDigit3 = vbNullString
    Dim yDigit4 As String: yDigit4 = vbNullString
    Dim thisYear As Integer: thisYear = 0
    Dim yearString As String: yearString = vbNullString
    Static thisDate As Date: thisDate = Now

    ' year digits for digital display

    yearString = Year(thisDate)

    yDigit1 = Left$(yearString, 1)
    yDigit2 = Mid$(yearString, 2, 1)
    yDigit3 = Mid$(yearString, 3, 1)
    yDigit4 = Mid$(yearString, 4, 1)

    Select Case yDigit1
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
    
    Select Case yDigit2
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
        
    Select Case yDigit3
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
    
    Select Case yDigit4
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



'Sub showMin1(ByRef CC As cCairoContext, ByVal mDigit1 As Integer)
'
'    Select Case mDigit1
'        Case 0
'            CC.RenderSurfaceContent "num0large", 404, 312, , , , 1
'        Case 1
'            CC.RenderSurfaceContent "num1large", 404, 312, , , , 1
'        Case 2
'            CC.RenderSurfaceContent "num2large", 404, 312, , , , 1
'        Case 3
'            CC.RenderSurfaceContent "num3large", 404, 312, , , , 1
'        Case 4
'            CC.RenderSurfaceContent "num4large", 404, 312, , , , 1
'        Case 5
'            CC.RenderSurfaceContent "num5large", 404, 312, , , , 1
'        Case 6
'            CC.RenderSurfaceContent "num6large", 404, 312, , , , 1
'        Case 7
'            CC.RenderSurfaceContent "num7large", 404, 312, , , , 1
'        Case 8
'            CC.RenderSurfaceContent "num8large", 404, 312, , , , 1
'        Case 9
'            CC.RenderSurfaceContent "num9large", 404, 312, , , , 1
'    End Select
'
'End Sub
