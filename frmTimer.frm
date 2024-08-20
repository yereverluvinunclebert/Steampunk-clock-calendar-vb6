VERSION 5.00
Begin VB.Form frmTimer 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmTimer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer settingsTimer 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   90
      Tag             =   "settingsTimer for reading external changes to prefs"
      Top             =   1095
   End
   Begin VB.Timer tmrScreenResolution 
      Enabled         =   0   'False
      Interval        =   4500
      Left            =   90
      Top             =   615
   End
   Begin VB.Timer revealWidgetTimer 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   90
      Top             =   135
   End
   Begin VB.Label Label3 
      Caption         =   "Note: this invisible form is also the container for the large 128x128px project icon"
      Height          =   435
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   4125
   End
   Begin VB.Label Label2 
      Caption         =   "settingsTimer for reading external changes to prefs"
      Height          =   195
      Left            =   705
      TabIndex        =   2
      Top             =   1170
      Width           =   3645
   End
   Begin VB.Label Label1 
      Caption         =   "ScreenResolutionTimer for handling rotation of the screen"
      Height          =   195
      Left            =   705
      TabIndex        =   1
      Top             =   735
      Width           =   3570
   End
   Begin VB.Label Label 
      Caption         =   "revealWidgetTimer for revealing after a hide."
      Height          =   195
      Left            =   690
      TabIndex        =   0
      Top             =   270
      Width           =   3480
   End
End
Attribute VB_Name = "frmTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule ModuleWithoutFolder
Option Explicit












'---------------------------------------------------------------------------------------
' Procedure : revealWidgetTimer_Timer
' Author    : beededea
' Date      : 05/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub revealWidgetTimer_Timer()
    On Error GoTo revealWidgetTimer_Timer_Error

    revealWidgetTimerCount = revealWidgetTimerCount + 1
    If revealWidgetTimerCount >= (minutesToHide * 12) Then
        revealWidgetTimerCount = 0

        fClock.clockForm.Visible = True
        revealWidgetTimer.Enabled = False
        gblWidgetHidden = "0"
        sPutINISetting "Software\SteampunkClockCalendar", "widgetHidden", gblWidgetHidden, gblSettingsFile
    End If

    On Error GoTo 0
    Exit Sub

revealWidgetTimer_Timer_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure revealWidgetTimer_Timer of Form frmTimer"
            Resume Next
          End If
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : tmrScreenResolution_Timer
' Author    : beededea
' Date      : 05/05/2023
' Purpose   : for handling rotation of the screen in tablet mode or a resolution change
'             possibly due to an old game in full screen mode.
'---------------------------------------------------------------------------------------
'
Private Sub tmrScreenResolution_Timer()

    Dim resizeProportion As Single: resizeProportion = 0
    
    On Error GoTo tmrScreenResolution_Timer_Error

    physicalScreenHeightPixels = GetDeviceCaps(Me.hdc, VERTRES)
    physicalScreenWidthPixels = GetDeviceCaps(Me.hdc, HORZRES)
    
    virtualScreenWidthPixels = fVirtualScreenWidth(True)
    virtualScreenHeightPixels = fVirtualScreenHeight(True)
    
    
'        If monitorCount > 1 Then
'        monitorStruct = formScreenProperties(fClock.clockForm)
'        a = monitorStruct.Width
'        a = monitorStruct.Height
'
'    End If
    
    ' if current position is beyond monitor
    If fClock.clockForm.Left > 3840 Then
               
        ' now calculate the size of the widget according to the screen height.
'        resizeProportion = 700 / oldPhysicalScreenHeightPixels
'        Call fClock.AdjustZoom(resizeProportion)
        
    End If
    
    ' will be used to check for orientation changes
    If (oldPhysicalScreenHeightPixels <> physicalScreenHeightPixels) Or (oldPhysicalScreenWidthPixels <> physicalScreenWidthPixels) Then
        
        ' move/hide onto/from the main screen and position per orientation portrait/landscape
        Call mainScreen
'
        'store the resolution change
        oldPhysicalScreenHeightPixels = physicalScreenHeightPixels
        oldPhysicalScreenWidthPixels = physicalScreenWidthPixels
    End If

    On Error GoTo 0
    Exit Sub

tmrScreenResolution_Timer_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrScreenResolution_Timer of Form frmTimer"
            Resume Next
          End If
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : settingsTimer_Timer
' Author    : beededea
' Date      : 13/05/2023
' Purpose   : if the unhide setting is set by another process it will unhide the widget
'---------------------------------------------------------------------------------------
'
Private Sub settingsTimer_Timer()
    
    On Error GoTo settingsTimer_Timer_Error

    gblUnhide = fGetINISetting("Software\SteampunkClockCalendar", "unhide", gblSettingsFile)

    If gblUnhide = "true" Then
        'overlayWidget.Hidden = False
        fClock.clockForm.Visible = True
        sPutINISetting "Software\SteampunkClockCalendar", "unhide", vbNullString, gblSettingsFile
    End If

    On Error GoTo 0
    Exit Sub

settingsTimer_Timer_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure settingsTimer_Timer of Form frmTimer"
            Resume Next
          End If
    End With
End Sub

