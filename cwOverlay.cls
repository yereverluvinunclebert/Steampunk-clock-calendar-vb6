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

Private HDeg As Double
Private MDeg As Double
Private SDeg As Double
Private Sec As Integer

Private WithEvents tmrClock As cTimer
Attribute tmrClock.VB_VarHelpID = -1
Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1

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

    
'---------------------------------------------------------------------------------------
' Procedure : tmrClock_Timer
' Author    : olaf schmidt
' Date      : 30/07/2023
' Purpose   : update the degree-Variables for the clock Hour, Minute and Second Hands
'---------------------------------------------------------------------------------------
'
Private Sub tmrClock_Timer()
 
    Static T As Date: T = Time
    
    On Error GoTo tmrClock_Timer_Error
    
    ' time calc for main gauge
    HDeg = 30 * (Hour(T) Mod 12 + (Minute(T) + tzDelta) / 60)
    MDeg = 6 * ((Minute(T) + tzDelta) + (Second(T) / 60))
    
    Sec = Second(T)

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
                    CC.RotateDrawingsDeg -90 + HDeg  'now we can rotate (the leading negative degrees are the rotation-offsets for 12 o'clock)
                 
                 Case "minutehand"  'minutes-shadow- and minutes-hand
                    Set Pat = Cairo.ImageList(Key).CreateSurfacePattern
                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(3, 3) 'shift to the "rotation-point" within the Pattern
                    CC.TranslateDrawings 646.5, 332 'shift to the center-point of the Overlay-Widget-area
                    
                    CC.RotateDrawingsDeg -90 + MDeg 'now we can rotate (the leading negative degrees are the rotation-offsets for 12 o'clock)
                    
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

   On Error GoTo 0
   Exit Sub

W_Paint_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_Paint of Class Module cwOverlay"
End Sub
 
 
