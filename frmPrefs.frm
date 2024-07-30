VERSION 5.00
Object = "{BCE37951-37DF-4D69-A8A3-2CFABEE7B3CC}#1.0#0"; "CCRSlider.ocx"
Begin VB.Form widgetPrefs 
   AutoRedraw      =   -1  'True
   Caption         =   "Steampunk Clock Calendar Preferences"
   ClientHeight    =   10545
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   8880
   Icon            =   "frmPrefs.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   10693.53
   ScaleMode       =   0  'User
   ScaleWidth      =   8880
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame fraDevelopmentButton 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   5490
      TabIndex        =   41
      Top             =   0
      Width           =   1065
      Begin VB.Label lblDevelopment 
         Caption         =   "Development"
         Height          =   240
         Left            =   45
         TabIndex        =   42
         Top             =   855
         Width           =   960
      End
      Begin VB.Image imgDevelopment 
         Height          =   600
         Left            =   150
         Picture         =   "frmPrefs.frx":0ECA
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
      Begin VB.Image imgDevelopmentClicked 
         Height          =   600
         Left            =   160
         Picture         =   "frmPrefs.frx":1482
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
   End
   Begin VB.Timer positionTimer 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1170
      Top             =   9690
   End
   Begin VB.CheckBox chkEnableResizing 
      Caption         =   "Enable Corner Resize"
      Height          =   210
      Left            =   3240
      TabIndex        =   130
      Top             =   10125
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Frame fraAboutButton 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   7695
      TabIndex        =   97
      Top             =   0
      Width           =   975
      Begin VB.Label lblAbout 
         Caption         =   "About"
         Height          =   240
         Index           =   0
         Left            =   255
         TabIndex        =   98
         Top             =   855
         Width           =   615
      End
      Begin VB.Image imgAbout 
         Height          =   600
         Left            =   160
         Picture         =   "frmPrefs.frx":1808
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
      Begin VB.Image imgAboutClicked 
         Height          =   600
         Left            =   160
         Picture         =   "frmPrefs.frx":1D90
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
   End
   Begin VB.Frame fraConfigButton 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   1215
      TabIndex        =   43
      Top             =   -15
      Width           =   930
      Begin VB.Label lblConfig 
         Caption         =   "Config."
         Height          =   240
         Index           =   0
         Left            =   270
         TabIndex        =   44
         Top             =   855
         Width           =   510
      End
      Begin VB.Image imgConfig 
         Height          =   600
         Left            =   165
         Picture         =   "frmPrefs.frx":227B
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
      Begin VB.Image imgConfigClicked 
         Height          =   600
         Left            =   165
         Picture         =   "frmPrefs.frx":285A
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
   End
   Begin VB.Frame fraPositionButton 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   4410
      TabIndex        =   39
      Top             =   0
      Width           =   930
      Begin VB.Label lblPosition 
         Caption         =   "Position"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   40
         Top             =   855
         Width           =   615
      End
      Begin VB.Image imgPosition 
         Height          =   600
         Left            =   160
         Picture         =   "frmPrefs.frx":2D5F
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
      Begin VB.Image imgPositionClicked 
         Height          =   600
         Left            =   160
         Picture         =   "frmPrefs.frx":3330
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   360
      Left            =   6090
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Save the changes you have made to the preferences"
      Top             =   10020
      Width           =   1320
   End
   Begin VB.CommandButton btnHelp 
      Caption         =   "Help"
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Open the help utility"
      Top             =   10035
      Width           =   1320
   End
   Begin VB.Frame fraSoundsButton 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   3315
      TabIndex        =   11
      Top             =   -15
      Width           =   930
      Begin VB.Label lblSounds 
         Caption         =   "Sounds"
         Height          =   240
         Left            =   210
         TabIndex        =   12
         Top             =   870
         Width           =   615
      End
      Begin VB.Image imgSounds 
         Height          =   600
         Left            =   160
         Picture         =   "frmPrefs.frx":36CE
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
      Begin VB.Image imgSoundsClicked 
         Height          =   600
         Left            =   160
         Picture         =   "frmPrefs.frx":3C8D
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
   End
   Begin VB.Timer themeTimer 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   660
      Top             =   9705
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      Height          =   360
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Close the utility"
      Top             =   10020
      Width           =   1320
   End
   Begin VB.Frame fraWindowButton 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   6615
      TabIndex        =   4
      Top             =   0
      Width           =   975
      Begin VB.Label lblWindow 
         Caption         =   "Window"
         Height          =   240
         Left            =   180
         TabIndex        =   5
         Top             =   855
         Width           =   615
      End
      Begin VB.Image imgWindow 
         Height          =   600
         Left            =   160
         Picture         =   "frmPrefs.frx":415D
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
      Begin VB.Image imgWindowClicked 
         Height          =   600
         Left            =   160
         Picture         =   "frmPrefs.frx":4627
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
   End
   Begin VB.Frame fraFontsButton 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   2280
      TabIndex        =   2
      Top             =   0
      Width           =   930
      Begin VB.Label lblFonts 
         Caption         =   "Fonts"
         Height          =   240
         Left            =   270
         TabIndex        =   3
         Top             =   855
         Width           =   510
      End
      Begin VB.Image imgFonts 
         Height          =   600
         Left            =   180
         Picture         =   "frmPrefs.frx":49D3
         Stretch         =   -1  'True
         Top             =   195
         Width           =   600
      End
      Begin VB.Image imgFontsClicked 
         Height          =   600
         Left            =   180
         Picture         =   "frmPrefs.frx":4F29
         Stretch         =   -1  'True
         Top             =   195
         Width           =   600
      End
   End
   Begin VB.Frame fraGeneralButton 
      Height          =   1140
      Left            =   240
      TabIndex        =   0
      Top             =   -15
      Width           =   930
      Begin VB.Image imgGeneral 
         Appearance      =   0  'Flat
         Height          =   600
         Left            =   165
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
      Begin VB.Label lblGeneral 
         Caption         =   "General"
         Height          =   240
         Index           =   0
         Left            =   195
         TabIndex        =   1
         Top             =   855
         Width           =   705
      End
      Begin VB.Image imgGeneralClicked 
         Height          =   600
         Left            =   165
         Stretch         =   -1  'True
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Frame fraSounds 
      Caption         =   "Sounds"
      Height          =   3330
      Left            =   855
      TabIndex        =   13
      Top             =   1230
      Visible         =   0   'False
      Width           =   7965
      Begin VB.Frame fraSoundsInner 
         BorderStyle     =   0  'None
         Height          =   2760
         Left            =   765
         TabIndex        =   25
         Top             =   285
         Width           =   6420
         Begin VB.CheckBox chkEnableBoost 
            Caption         =   "Decide whether the various sounds will be boosted."
            Height          =   225
            Left            =   1485
            TabIndex        =   181
            ToolTipText     =   "Check this box to enable or disable all of the sounds used during any animation on the main screen."
            Top             =   1680
            Width           =   4365
         End
         Begin VB.CheckBox chkEnableChimes 
            Caption         =   "Decide whether the clock chimes will sound."
            Height          =   225
            Left            =   1485
            TabIndex        =   179
            ToolTipText     =   "Check this box to enable or disable all of the sounds used during any animation on the main screen."
            Top             =   1215
            Width           =   4365
         End
         Begin VB.CheckBox chkEnableTicks 
            Caption         =   "Decide whether the clock has the tick sounds enabled."
            Height          =   225
            Left            =   1485
            TabIndex        =   177
            ToolTipText     =   "Check this box to enable or disable all of the sounds used during any animation on the main screen."
            Top             =   765
            Width           =   4365
         End
         Begin VB.CheckBox chkEnableSounds 
            Caption         =   "Enable ALL sounds for the whole widget."
            Height          =   225
            Left            =   1485
            TabIndex        =   36
            ToolTipText     =   "Check this box to enable or disable all of the sounds used during any animation on the main screen."
            Top             =   285
            Width           =   4485
         End
         Begin VB.Label lblSoundsTab 
            Caption         =   "Determine the sound of UI elements, clock tick and chiming volumes. Set the overall volume to loud or quiet."
            Height          =   540
            Index           =   4
            Left            =   1470
            TabIndex        =   183
            Tag             =   "lblSharedInputFile"
            Top             =   2100
            Width           =   4680
         End
         Begin VB.Label lblSoundsTab 
            Caption         =   "Volume Boost :"
            Height          =   255
            Index           =   2
            Left            =   300
            TabIndex        =   182
            Tag             =   "lblSharedInputFile"
            Top             =   1680
            Width           =   1185
         End
         Begin VB.Label lblSoundsTab 
            Caption         =   "Chimes :"
            Height          =   255
            Index           =   1
            Left            =   765
            TabIndex        =   180
            Tag             =   "lblSharedInputFile"
            Top             =   1215
            Width           =   765
         End
         Begin VB.Label lblSoundsTab 
            Caption         =   "Ticks :"
            Height          =   255
            Index           =   0
            Left            =   885
            TabIndex        =   178
            Tag             =   "lblSharedInputFile"
            Top             =   765
            Width           =   765
         End
         Begin VB.Label lblSoundsTab 
            Caption         =   "Audio :"
            Height          =   255
            Index           =   3
            Left            =   885
            TabIndex        =   95
            Tag             =   "lblSharedInputFile"
            Top             =   285
            Width           =   765
         End
      End
   End
   Begin VB.Frame fraFonts 
      Caption         =   "Fonts"
      Height          =   4320
      Left            =   240
      TabIndex        =   9
      Top             =   1230
      Width           =   7335
      Begin VB.Frame fraFontsInner 
         BorderStyle     =   0  'None
         Height          =   3855
         Left            =   690
         TabIndex        =   26
         Top             =   360
         Width           =   6105
         Begin VB.CommandButton btnResetMessages 
            Caption         =   "Reset"
            Height          =   300
            Left            =   1710
            Style           =   1  'Graphical
            TabIndex        =   147
            Top             =   3405
            Width           =   885
         End
         Begin VB.TextBox txtPrefsFontCurrentSize 
            Height          =   315
            Left            =   4200
            Locked          =   -1  'True
            TabIndex        =   131
            ToolTipText     =   "Disabled for manual input. Shows the current font size when form resizing is enabled."
            Top             =   1065
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.TextBox txtPrefsFontSize 
            Height          =   315
            Left            =   1710
            Locked          =   -1  'True
            TabIndex        =   29
            Text            =   "8"
            ToolTipText     =   "Disabled for manual input. Choose a font size using the font selector to be used within this preferences window only"
            Top             =   1065
            Width           =   510
         End
         Begin VB.CommandButton btnPrefsFont 
            Caption         =   "Font"
            Height          =   300
            Left            =   5025
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "The Font Selector."
            Top             =   75
            Width           =   585
         End
         Begin VB.TextBox txtPrefsFont 
            Height          =   315
            Left            =   1710
            Locked          =   -1  'True
            TabIndex        =   27
            Text            =   "Times New Roman"
            ToolTipText     =   "Disabled for manual input. Choose a font via the font selector to be used only for this preferences window"
            Top             =   90
            Width           =   3285
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Hidden message boxes can be reactivated by pressing this reset button."
            Height          =   480
            Index           =   4
            Left            =   2700
            TabIndex        =   148
            ToolTipText     =   "We suggest Centurion Light SF at 8pt - which you will find in the FCW program folder"
            Top             =   3345
            Width           =   3360
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Reset Pop ups :"
            Height          =   300
            Index           =   1
            Left            =   435
            TabIndex        =   146
            Tag             =   "lblPrefsFont"
            Top             =   3450
            Width           =   1470
         End
         Begin VB.Label lblCurrentFontsTab 
            Caption         =   "Resized Font"
            Height          =   315
            Left            =   4950
            TabIndex        =   132
            ToolTipText     =   "Disabled for manual input. Shows the current font size when form resizing is enabled."
            Top             =   1110
            Visible         =   0   'False
            Width           =   2400
         End
         Begin VB.Label lblFontsTab 
            Caption         =   $"frmPrefs.frx":53C2
            Height          =   1710
            Index           =   0
            Left            =   1725
            TabIndex        =   96
            ToolTipText     =   "We suggest Centurion Light SF at 8pt - which you will find in the FCW program folder"
            Top             =   1455
            Width           =   4455
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "The chosen font size *"
            Height          =   480
            Index           =   7
            Left            =   2370
            TabIndex        =   33
            ToolTipText     =   "Choose a font size that fits the text boxes"
            Top             =   1095
            Width           =   2400
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Base Font Size :"
            Height          =   330
            Index           =   3
            Left            =   435
            TabIndex        =   32
            Tag             =   "lblPrefsFontSize"
            Top             =   1095
            Width           =   1230
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Prefs Window Font:"
            Height          =   300
            Index           =   2
            Left            =   15
            TabIndex        =   31
            Tag             =   "lblPrefsFont"
            ToolTipText     =   "We suggest Centurion Light SF at 8pt - which you will find in the FCW program folder"
            Top             =   120
            Width           =   1665
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Choose a font to be used for the text in this preferences window, gauge tooltips and message boxes *"
            Height          =   480
            Index           =   6
            Left            =   1695
            TabIndex        =   30
            ToolTipText     =   "We suggest Centurion Light SF at 8pt - which you will find in the FCW program folder"
            Top             =   480
            Width           =   4035
         End
      End
   End
   Begin VB.Frame fraWindow 
      Caption         =   "Window"
      Height          =   6300
      Left            =   405
      TabIndex        =   10
      Top             =   1515
      Width           =   8280
      Begin VB.Frame fraWindowInner 
         BorderStyle     =   0  'None
         Height          =   5775
         Left            =   1095
         TabIndex        =   14
         Top             =   345
         Width           =   5715
         Begin VB.Frame fraHiding 
            BorderStyle     =   0  'None
            Height          =   2010
            Left            =   480
            TabIndex        =   119
            Top             =   2325
            Width           =   5130
            Begin VB.ComboBox cmbHidingTime 
               Height          =   315
               Left            =   825
               Style           =   2  'Dropdown List
               TabIndex        =   122
               Top             =   1680
               Width           =   3720
            End
            Begin VB.CheckBox chkWidgetHidden 
               Caption         =   "Hiding Widget *"
               Height          =   225
               Left            =   855
               TabIndex        =   120
               Top             =   225
               Width           =   2955
            End
            Begin VB.Label lblWindowLevel 
               Caption         =   "Hiding :"
               Height          =   345
               Index           =   2
               Left            =   90
               TabIndex        =   123
               Top             =   210
               Width           =   720
            End
            Begin VB.Label lblWindowLevel 
               Caption         =   $"frmPrefs.frx":5500
               Height          =   975
               Index           =   1
               Left            =   855
               TabIndex        =   121
               Top             =   600
               Width           =   3900
            End
         End
         Begin VB.ComboBox cmbWindowLevel 
            Height          =   315
            Left            =   1305
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   0
            Width           =   3720
         End
         Begin VB.CheckBox chkIgnoreMouse 
            Caption         =   "Ignore Mouse *"
            Height          =   225
            Left            =   1335
            TabIndex        =   15
            ToolTipText     =   "Checking this box causes the program to ignore all mouse events."
            Top             =   1500
            Width           =   2535
         End
         Begin vb6projectCCRSlider.Slider sliOpacity 
            Height          =   390
            Left            =   1200
            TabIndex        =   16
            ToolTipText     =   "Set the transparency of the Program."
            Top             =   4560
            Width           =   3870
            _ExtentX        =   6826
            _ExtentY        =   688
            Min             =   20
            Max             =   100
            Value           =   20
            SelStart        =   20
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "This setting controls the relative layering of this widget. You may use it to place it on top of other windows or underneath. "
            Height          =   660
            Index           =   3
            Left            =   1320
            TabIndex        =   128
            Top             =   570
            Width           =   3810
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Window Level :"
            Height          =   345
            Index           =   0
            Left            =   0
            TabIndex        =   24
            Top             =   60
            Width           =   1740
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "20%"
            Height          =   315
            Index           =   7
            Left            =   1290
            TabIndex        =   23
            Top             =   5070
            Width           =   345
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "100%"
            Height          =   315
            Index           =   9
            Left            =   4650
            TabIndex        =   22
            Top             =   5070
            Width           =   405
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Opacity"
            Height          =   315
            Index           =   8
            Left            =   2775
            TabIndex        =   21
            Top             =   5070
            Width           =   840
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Opacity:"
            Height          =   315
            Index           =   6
            Left            =   555
            TabIndex        =   20
            Top             =   4620
            Width           =   780
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Set the program transparency level."
            Height          =   330
            Index           =   5
            Left            =   1335
            TabIndex        =   19
            Top             =   5385
            Width           =   3810
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Checking this box causes the program to ignore all mouse events except right click menu interactions."
            Height          =   660
            Index           =   4
            Left            =   1320
            TabIndex        =   18
            Top             =   1890
            Width           =   3810
         End
      End
   End
   Begin VB.Frame fraDevelopment 
      Caption         =   "Development"
      Height          =   6210
      Left            =   240
      TabIndex        =   47
      Top             =   1200
      Visible         =   0   'False
      Width           =   8520
      Begin VB.Frame fraDevelopmentInner 
         BorderStyle     =   0  'None
         Height          =   5595
         Left            =   870
         TabIndex        =   48
         Top             =   300
         Width           =   7455
         Begin VB.Frame fraDefaultEditor 
            BorderStyle     =   0  'None
            Height          =   2370
            Left            =   75
            TabIndex        =   133
            Top             =   3165
            Width           =   7290
            Begin VB.CommandButton btnDefaultEditor 
               Caption         =   "..."
               Height          =   300
               Left            =   5115
               Style           =   1  'Graphical
               TabIndex        =   135
               ToolTipText     =   "Click to select the .vbp file to edit the program - You need to have access to the source!"
               Top             =   210
               Width           =   315
            End
            Begin VB.TextBox txtDefaultEditor 
               Height          =   315
               Left            =   1440
               TabIndex        =   134
               Text            =   " eg. E:\vb6\fire call\FireCallWin.vbp"
               Top             =   195
               Width           =   3660
            End
            Begin VB.Label lblGitHub 
               Caption         =   $"frmPrefs.frx":55A3
               ForeColor       =   &H8000000D&
               Height          =   915
               Left            =   1560
               TabIndex        =   139
               ToolTipText     =   "Double Click to visit github"
               Top             =   1440
               Width           =   4935
            End
            Begin VB.Label lblDebug 
               Caption         =   $"frmPrefs.frx":566A
               Height          =   930
               Index           =   9
               Left            =   1545
               TabIndex        =   137
               Top             =   690
               Width           =   4785
            End
            Begin VB.Label lblDebug 
               Caption         =   "Default Editor :"
               Height          =   255
               Index           =   7
               Left            =   285
               TabIndex        =   136
               Tag             =   "lblSharedInputFile"
               Top             =   225
               Width           =   1350
            End
         End
         Begin VB.TextBox txtDblClickCommand 
            Height          =   315
            Left            =   1515
            TabIndex        =   59
            ToolTipText     =   "Enter a Windows command for the gauge to operate when double-clicked."
            Top             =   1095
            Width           =   3660
         End
         Begin VB.CommandButton btnOpenFile 
            Caption         =   "..."
            Height          =   300
            Left            =   5175
            Style           =   1  'Graphical
            TabIndex        =   56
            ToolTipText     =   "Click to select a particular file for the gauge to run or open when double-clicked."
            Top             =   2250
            Width           =   315
         End
         Begin VB.TextBox txtOpenFile 
            Height          =   315
            Left            =   1515
            TabIndex        =   55
            ToolTipText     =   "Enter a particular file for the gauge to run or open when double-clicked."
            Top             =   2235
            Width           =   3660
         End
         Begin VB.ComboBox cmbDebug 
            Height          =   315
            ItemData        =   "frmPrefs.frx":570E
            Left            =   1530
            List            =   "frmPrefs.frx":5710
            Style           =   2  'Dropdown List
            TabIndex        =   52
            ToolTipText     =   "Choose to set debug mode."
            Top             =   -15
            Width           =   2160
         End
         Begin VB.Label lblDebug 
            Caption         =   "DblClick Command :"
            Height          =   510
            Index           =   1
            Left            =   -15
            TabIndex        =   61
            Tag             =   "lblPrefixString"
            Top             =   1155
            Width           =   1545
         End
         Begin VB.Label lblConfigurationTab 
            Caption         =   "Shift+double-clicking on the widget image will open this file. "
            Height          =   375
            Index           =   6
            Left            =   1560
            TabIndex        =   60
            Top             =   2730
            Width           =   3705
         End
         Begin VB.Label lblDebug 
            Caption         =   "Default command to run when the gauge receives a double-click eg.  mmsys.cpl to run the sounds utility."
            Height          =   570
            Index           =   5
            Left            =   1590
            TabIndex        =   58
            Tag             =   "lblSharedInputFileDesc"
            Top             =   1605
            Width           =   4410
         End
         Begin VB.Label lblDebug 
            Caption         =   "Open File :"
            Height          =   255
            Index           =   4
            Left            =   645
            TabIndex        =   57
            Tag             =   "lblSharedInputFile"
            Top             =   2280
            Width           =   1350
         End
         Begin VB.Label lblDebug 
            Caption         =   "Turning on the debugging will provide extra information in the debug window.  *"
            Height          =   495
            Index           =   2
            Left            =   1545
            TabIndex        =   54
            Tag             =   "lblAlarmSoundDesc"
            Top             =   450
            Width           =   4455
         End
         Begin VB.Label lblDebug 
            Caption         =   "Debug :"
            Height          =   375
            Index           =   0
            Left            =   855
            TabIndex        =   53
            Tag             =   "lblAlarmSound"
            Top             =   45
            Width           =   1740
         End
      End
   End
   Begin VB.Frame fraConfig 
      Caption         =   "Configuration"
      Height          =   6690
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   7605
      Begin VB.Frame fraConfigInner 
         BorderStyle     =   0  'None
         Height          =   6120
         Left            =   435
         TabIndex        =   34
         Top             =   435
         Width           =   6450
         Begin VB.CheckBox chkShowHelp 
            Caption         =   "Show Help on Widget Start"
            Height          =   225
            Left            =   2010
            TabIndex        =   151
            ToolTipText     =   "Check the box to show the widget in the taskbar"
            Top             =   4515
            Width           =   3405
         End
         Begin VB.CheckBox chkEnablePrefsTooltips 
            Caption         =   "Enable Preference Utility Tooltips *"
            Height          =   225
            Left            =   2010
            TabIndex        =   142
            ToolTipText     =   "Check the box to enable tooltips for all controls in the preferences utility"
            Top             =   3720
            Width           =   3345
         End
         Begin VB.CheckBox chkDpiAwareness 
            Caption         =   "DPI Awareness Enable *"
            Height          =   285
            Left            =   2010
            TabIndex        =   140
            ToolTipText     =   "Check the box to make the program DPI aware. RESTART required."
            Top             =   4860
            Width           =   3405
         End
         Begin VB.CheckBox chkShowTaskbar 
            Caption         =   "Show Widget in Taskbar"
            Height          =   225
            Left            =   2010
            TabIndex        =   138
            ToolTipText     =   "Check the box to show the widget in the taskbar"
            Top             =   4110
            Width           =   3405
         End
         Begin VB.ComboBox cmbScrollWheelDirection 
            Height          =   315
            Left            =   1995
            Style           =   2  'Dropdown List
            TabIndex        =   86
            ToolTipText     =   "To change the direction of the mouse scroll wheel when resiziing the globe gauge."
            Top             =   1695
            Width           =   2490
         End
         Begin VB.CheckBox chkEnableBalloonTooltips 
            Caption         =   "Enable Balloon Tooltips on all Controls *"
            Height          =   225
            Left            =   2010
            TabIndex        =   38
            ToolTipText     =   "Check the box to enable larger balloon tooltips for all controls on the main program"
            Top             =   3345
            Width           =   4035
         End
         Begin VB.CheckBox chkEnableTooltips 
            Caption         =   "Enable Main Program Tooltips"
            Height          =   225
            Left            =   2010
            TabIndex        =   35
            ToolTipText     =   "Check the box to enable tooltips for all controls on the main program"
            Top             =   2910
            Width           =   3345
         End
         Begin vb6projectCCRSlider.Slider sliGaugeSize 
            Height          =   390
            Left            =   1920
            TabIndex        =   94
            ToolTipText     =   "Adjust to a percentage of the original size. You can also use Ctrl+Mousewheel."
            Top             =   60
            Width           =   3870
            _ExtentX        =   6826
            _ExtentY        =   688
            Min             =   5
            Max             =   200
            Value           =   5
            TickFrequency   =   3
            SelStart        =   5
         End
         Begin VB.Label lblConfiguration 
            Caption         =   $"frmPrefs.frx":5712
            Height          =   915
            Index           =   0
            Left            =   1980
            TabIndex        =   141
            Top             =   5220
            Width           =   4335
         End
         Begin VB.Label lblConfiguration 
            Caption         =   "The scroll-wheel resizing direction can be determined here. The direction chosen causes the gauge to grow. *"
            Height          =   690
            Index           =   6
            Left            =   2025
            TabIndex        =   118
            Top             =   2115
            Width           =   3930
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "160"
            Height          =   315
            Index           =   4
            Left            =   4740
            TabIndex        =   90
            Top             =   555
            Width           =   345
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "120"
            Height          =   315
            Index           =   3
            Left            =   3990
            TabIndex        =   89
            Top             =   555
            Width           =   345
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "50"
            Height          =   315
            Index           =   1
            Left            =   2790
            TabIndex        =   88
            Top             =   555
            Width           =   345
         End
         Begin VB.Label lblConfiguration 
            Caption         =   "Mouse Wheel Resize :"
            Height          =   345
            Index           =   3
            Left            =   255
            TabIndex        =   87
            ToolTipText     =   "To change the direction of the mouse scroll wheel when resiziing the globe gauge."
            Top             =   1740
            Width           =   2055
         End
         Begin VB.Label lblConfiguration 
            Caption         =   "Adjust to a percentage of the original size. You can also use Ctrl+Mousewheel. Immediate. *"
            Height          =   555
            Index           =   2
            Left            =   2070
            TabIndex        =   85
            ToolTipText     =   "Adjust to a percentage of the original size. You can also use Ctrl+Mousewheel."
            Top             =   870
            Width           =   3810
         End
         Begin VB.Label lblConfiguration 
            Caption         =   "Gauge Size :"
            Height          =   315
            Index           =   1
            Left            =   885
            TabIndex        =   84
            Top             =   105
            Width           =   975
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "80"
            Height          =   315
            Index           =   2
            Left            =   3360
            TabIndex        =   83
            Top             =   555
            Width           =   360
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "200 (%)"
            Height          =   315
            Index           =   5
            Left            =   5385
            TabIndex        =   82
            Top             =   555
            Width           =   735
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "5"
            Height          =   315
            Index           =   0
            Left            =   2085
            TabIndex        =   81
            Top             =   555
            Width           =   345
         End
      End
   End
   Begin VB.Frame fraGeneral 
      Caption         =   "General"
      ForeColor       =   &H80000008&
      Height          =   6615
      Left            =   75
      TabIndex        =   49
      Top             =   1200
      Visible         =   0   'False
      Width           =   7995
      Begin VB.Frame fraGeneralInner 
         BorderStyle     =   0  'None
         Height          =   6240
         Left            =   435
         TabIndex        =   50
         Top             =   300
         Width           =   6750
         Begin VB.CheckBox chkTogglePendulum 
            Caption         =   "Toggle the pendulum animation *"
            Height          =   465
            Left            =   1995
            TabIndex        =   175
            ToolTipText     =   "Check this box to enable the automatic start of the program when Windows is started."
            Top             =   1755
            Width           =   4020
         End
         Begin VB.Frame Frame 
            BorderStyle     =   0  'None
            Height          =   3795
            Left            =   -105
            TabIndex        =   152
            Top             =   2505
            Width           =   7065
            Begin VB.CommandButton btnVerifyDateTime5 
               Caption         =   ">"
               Height          =   360
               Left            =   4410
               Style           =   1  'Graphical
               TabIndex        =   167
               ToolTipText     =   "Verify Date Time for alarm number 1"
               Top             =   2175
               Width           =   300
            End
            Begin VB.TextBox txtAlarm5Time 
               Height          =   360
               Left            =   3465
               TabIndex        =   166
               Top             =   2160
               Width           =   915
            End
            Begin VB.CommandButton btnVerifyDateTime4 
               Caption         =   ">"
               Height          =   360
               Left            =   4410
               Style           =   1  'Graphical
               TabIndex        =   165
               ToolTipText     =   "Verify Date Time for alarm number 1"
               Top             =   1665
               Width           =   300
            End
            Begin VB.TextBox txtAlarm4Time 
               Height          =   360
               Left            =   3465
               TabIndex        =   164
               Top             =   1650
               Width           =   915
            End
            Begin VB.CommandButton btnVerifyDateTime3 
               Caption         =   ">"
               Height          =   360
               Left            =   4410
               Style           =   1  'Graphical
               TabIndex        =   163
               ToolTipText     =   "Verify Date Time for alarm number 1"
               Top             =   1185
               Width           =   300
            End
            Begin VB.TextBox txtAlarm3Time 
               Height          =   360
               Left            =   3465
               TabIndex        =   162
               Top             =   1170
               Width           =   915
            End
            Begin VB.CommandButton btnVerifyDateTime2 
               Caption         =   ">"
               Height          =   360
               Left            =   4410
               Style           =   1  'Graphical
               TabIndex        =   161
               ToolTipText     =   "Verify Date Time for alarm number 1"
               Top             =   705
               Width           =   300
            End
            Begin VB.TextBox txtAlarm2Time 
               Height          =   360
               Left            =   3465
               TabIndex        =   160
               Top             =   690
               Width           =   915
            End
            Begin VB.TextBox txtAlarm1Time 
               Height          =   360
               Left            =   3465
               TabIndex        =   159
               Top             =   210
               Width           =   915
            End
            Begin VB.CommandButton btnVerifyDateTime1 
               Caption         =   ">"
               Height          =   360
               Left            =   4410
               Style           =   1  'Graphical
               TabIndex        =   158
               ToolTipText     =   "Verify Date Time for alarm number 1"
               Top             =   225
               Width           =   300
            End
            Begin VB.TextBox txtAlarm5Date 
               Height          =   360
               Left            =   2100
               TabIndex        =   157
               Top             =   2160
               Width           =   1320
            End
            Begin VB.TextBox txtAlarm4Date 
               Height          =   360
               Left            =   2100
               TabIndex        =   156
               Top             =   1665
               Width           =   1320
            End
            Begin VB.TextBox txtAlarm3Date 
               Height          =   360
               Left            =   2100
               TabIndex        =   155
               Top             =   1185
               Width           =   1320
            End
            Begin VB.TextBox txtAlarm2Date 
               Height          =   360
               Left            =   2100
               TabIndex        =   154
               Top             =   705
               Width           =   1320
            End
            Begin VB.TextBox txtAlarm1Date 
               Height          =   360
               Left            =   2100
               TabIndex        =   153
               Top             =   210
               Width           =   1320
            End
            Begin VB.Label lblGeneral 
               Caption         =   $"frmPrefs.frx":57C6
               Height          =   960
               Index           =   10
               Left            =   2070
               TabIndex        =   173
               Top             =   2670
               Width           =   4230
            End
            Begin VB.Label lblGeneral 
               Caption         =   "Alarm No. 5 :"
               Height          =   375
               Index           =   9
               Left            =   1005
               TabIndex        =   172
               Tag             =   "lblRefreshInterval"
               Top             =   2220
               Width           =   1095
            End
            Begin VB.Label lblGeneral 
               Caption         =   "Alarm No. 4 :"
               Height          =   375
               Index           =   8
               Left            =   1005
               TabIndex        =   171
               Tag             =   "lblRefreshInterval"
               Top             =   1725
               Width           =   1095
            End
            Begin VB.Label lblGeneral 
               Caption         =   "Alarm No. 3 :"
               Height          =   375
               Index           =   7
               Left            =   1005
               TabIndex        =   170
               Tag             =   "lblRefreshInterval"
               Top             =   1245
               Width           =   1095
            End
            Begin VB.Label lblGeneral 
               Caption         =   "Alarm No. 2 :"
               Height          =   375
               Index           =   5
               Left            =   1005
               TabIndex        =   169
               Tag             =   "lblRefreshInterval"
               Top             =   765
               Width           =   1095
            End
            Begin VB.Label lblGeneral 
               Caption         =   "Alarm No. 1 :"
               Height          =   375
               Index           =   4
               Left            =   1005
               TabIndex        =   168
               Tag             =   "lblRefreshInterval"
               Top             =   270
               Width           =   1095
            End
         End
         Begin VB.CheckBox chkWidgetFunctions 
            Caption         =   "Toggle to enable/disable"
            Height          =   390
            Left            =   1995
            TabIndex        =   51
            ToolTipText     =   "When checked this box enables the spinning earth functionality. That's it!"
            Top             =   165
            Width           =   3405
         End
         Begin VB.CheckBox chkGenStartup 
            Caption         =   "Run the Steampunk Clock Calendar Widget at Windows Startup "
            Height          =   465
            Left            =   1995
            TabIndex        =   91
            ToolTipText     =   "Check this box to enable the automatic start of the program when Windows is started."
            Top             =   1230
            Width           =   4020
         End
         Begin VB.Label lblGeneral 
            Caption         =   "Pendulum Swing :"
            Height          =   375
            Index           =   3
            Left            =   480
            TabIndex        =   176
            Tag             =   "lblRefreshInterval"
            Top             =   1875
            Width           =   1740
         End
         Begin VB.Label lblGeneral 
            Caption         =   "Alarm Dates and times shown below"
            Height          =   435
            Index           =   2
            Left            =   2010
            TabIndex        =   174
            Top             =   2280
            Width           =   4215
         End
         Begin VB.Label lblGeneral 
            Caption         =   "When checked this box enables the functionality of this widget to control audio - That's it! *"
            Height          =   465
            Index           =   1
            Left            =   1995
            TabIndex        =   150
            Top             =   675
            Width           =   4020
         End
         Begin VB.Label lblGeneral 
            Caption         =   "Widget Functions :"
            Height          =   315
            Index           =   6
            Left            =   495
            TabIndex        =   93
            Top             =   255
            Width           =   1395
         End
         Begin VB.Label lblGeneral 
            Caption         =   "Auto Start :"
            Height          =   375
            Index           =   11
            Left            =   975
            TabIndex        =   92
            Tag             =   "lblRefreshInterval"
            Top             =   1350
            Width           =   1740
         End
      End
   End
   Begin VB.Frame fraAbout 
      Caption         =   "About"
      Height          =   8580
      Left            =   255
      TabIndex        =   99
      Top             =   1185
      Visible         =   0   'False
      Width           =   8520
      Begin VB.Frame fraScrollbarCover 
         BorderStyle     =   0  'None
         Height          =   6435
         Left            =   7950
         TabIndex        =   113
         Top             =   1995
         Width           =   420
      End
      Begin VB.TextBox txtAboutText 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   6135
         Left            =   300
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   112
         Text            =   "frmPrefs.frx":5870
         Top             =   2205
         Width           =   8010
      End
      Begin VB.CommandButton btnAboutDebugInfo 
         Caption         =   "Debug &Info."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6780
         Style           =   1  'Graphical
         TabIndex        =   103
         ToolTipText     =   "This gives access to the debugging tool"
         Top             =   1110
         Width           =   1470
      End
      Begin VB.CommandButton btnFacebook 
         Caption         =   "&Facebook"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6780
         Style           =   1  'Graphical
         TabIndex        =   102
         ToolTipText     =   "This will link you to the Rocket/Steamy dock users Group"
         Top             =   735
         Width           =   1470
      End
      Begin VB.CommandButton btnUpdate 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6780
         Style           =   1  'Graphical
         TabIndex        =   101
         ToolTipText     =   "Here you can visit the update location where you can download new versions of the programs."
         Top             =   360
         Width           =   1470
      End
      Begin VB.CommandButton btnDonate 
         Caption         =   "&Donate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6780
         Style           =   1  'Graphical
         TabIndex        =   100
         ToolTipText     =   "Opens a browser window and sends you to our donate page on Amazon"
         Top             =   1485
         Width           =   1470
      End
      Begin VB.Label lblDotDot 
         BackStyle       =   0  'Transparent
         Caption         =   ".        ."
         BeginProperty Font 
            Name            =   "Centurion Light SF"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2940
         TabIndex        =   117
         Top             =   510
         Width           =   495
      End
      Begin VB.Label lblRevisionNum 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Centurion Light SF"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3450
         TabIndex        =   116
         Top             =   510
         Width           =   525
      End
      Begin VB.Label lblMajorVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Centurion Light SF"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2730
         TabIndex        =   115
         Top             =   510
         Width           =   225
      End
      Begin VB.Label lblMinorVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Centurion Light SF"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3090
         TabIndex        =   114
         Top             =   510
         Width           =   225
      End
      Begin VB.Label lblAbout 
         Caption         =   "Dean Beedell © 2023"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   2715
         TabIndex        =   111
         Top             =   855
         Width           =   2175
      End
      Begin VB.Label lblAbout 
         Caption         =   "Originator"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   1050
         TabIndex        =   110
         Top             =   855
         Width           =   795
      End
      Begin VB.Label lblAbout 
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   1065
         TabIndex        =   109
         Top             =   495
         Width           =   795
      End
      Begin VB.Label lblAbout 
         Caption         =   "Dean Beedell © 2023"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   2715
         TabIndex        =   108
         Top             =   1215
         Width           =   2175
      End
      Begin VB.Label lblAbout 
         Caption         =   "Current Developer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   1050
         TabIndex        =   107
         Top             =   1215
         Width           =   1470
      End
      Begin VB.Label lblAbout 
         Caption         =   "Target"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   1050
         TabIndex        =   106
         Top             =   1560
         Width           =   1470
      End
      Begin VB.Label lblAbout 
         Caption         =   "Windows Vista, 7, 8, 10  && 11"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   2715
         TabIndex        =   105
         Top             =   1560
         Width           =   3735
      End
      Begin VB.Label lblAbout 
         Caption         =   "(32bit WoW64)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   3900
         TabIndex        =   104
         Top             =   510
         Width           =   1245
      End
   End
   Begin VB.Frame fraPosition 
      Caption         =   "Position"
      Height          =   7440
      Left            =   255
      TabIndex        =   45
      Top             =   1230
      Visible         =   0   'False
      Width           =   8520
      Begin VB.Frame fraPositionInner 
         BorderStyle     =   0  'None
         Height          =   6960
         Left            =   150
         TabIndex        =   46
         Top             =   300
         Width           =   7680
         Begin VB.TextBox txtLandscapeHoffset 
            Height          =   315
            Left            =   2250
            TabIndex        =   73
            Top             =   4425
            Width           =   2130
         End
         Begin VB.CheckBox chkPreventDragging 
            Caption         =   "Widget Position Locked. *"
            Height          =   225
            Left            =   2265
            TabIndex        =   126
            ToolTipText     =   "Checking this box turns off the ability to drag the program with the mouse, locking it in position."
            Top             =   3465
            Width           =   2505
         End
         Begin VB.TextBox txtPortraitYoffset 
            Height          =   315
            Left            =   2250
            TabIndex        =   79
            ToolTipText     =   "Enter a prefix/nickname for outgoing messages."
            Top             =   6465
            Width           =   2130
         End
         Begin VB.TextBox txtPortraitHoffset 
            Height          =   315
            Left            =   2250
            TabIndex        =   77
            ToolTipText     =   "Enter a prefix/nickname for outgoing messages."
            Top             =   6000
            Width           =   2130
         End
         Begin VB.TextBox txtLandscapeVoffset 
            Height          =   315
            Left            =   2250
            TabIndex        =   75
            ToolTipText     =   "Enter a prefix/nickname for outgoing messages."
            Top             =   4875
            Width           =   2130
         End
         Begin VB.ComboBox cmbWidgetLandscape 
            Height          =   315
            Left            =   2250
            Style           =   2  'Dropdown List
            TabIndex        =   71
            ToolTipText     =   "Choose the alarm sound."
            Top             =   3930
            Width           =   2160
         End
         Begin VB.ComboBox cmbWidgetPortrait 
            Height          =   315
            Left            =   2250
            Style           =   2  'Dropdown List
            TabIndex        =   68
            ToolTipText     =   "Choose the alarm sound."
            Top             =   5505
            Width           =   2160
         End
         Begin VB.ComboBox cmbWidgetPosition 
            Height          =   315
            Left            =   2265
            Style           =   2  'Dropdown List
            TabIndex        =   65
            ToolTipText     =   "Choose the alarm sound."
            Top             =   2100
            Width           =   2160
         End
         Begin VB.ComboBox cmbAspectHidden 
            Height          =   315
            Left            =   2265
            Style           =   2  'Dropdown List
            TabIndex        =   62
            ToolTipText     =   "Choose the alarm sound."
            Top             =   0
            Width           =   2160
         End
         Begin VB.Label lblPosition 
            Caption         =   "(px)"
            Height          =   300
            Index           =   7
            Left            =   4530
            TabIndex        =   144
            Tag             =   "lblPrefixString"
            Top             =   6495
            Width           =   435
         End
         Begin VB.Label lblPosition 
            Caption         =   "(px)"
            Height          =   300
            Index           =   5
            Left            =   4530
            TabIndex        =   143
            Tag             =   "lblPrefixString"
            Top             =   6045
            Width           =   435
         End
         Begin VB.Label lblPosition 
            Caption         =   "*"
            Height          =   255
            Index           =   1
            Left            =   4545
            TabIndex        =   129
            Tag             =   "lblAlarmSound"
            Top             =   45
            Width           =   345
         End
         Begin VB.Label lblPosition 
            Caption         =   "(px)"
            Height          =   300
            Index           =   4
            Left            =   4530
            TabIndex        =   125
            Tag             =   "lblPrefixString"
            Top             =   4905
            Width           =   435
         End
         Begin VB.Label lblPosition 
            Caption         =   "(px)"
            Height          =   300
            Index           =   2
            Left            =   4530
            TabIndex        =   124
            Tag             =   "lblPrefixString"
            Top             =   4500
            Width           =   390
         End
         Begin VB.Label lblPosition 
            Caption         =   "Portrait Top Y pos :"
            Height          =   510
            Index           =   17
            Left            =   645
            TabIndex        =   80
            Tag             =   "lblPrefixString"
            Top             =   6480
            Width           =   2175
         End
         Begin VB.Label lblPosition 
            Caption         =   "Portrait Left X pos :"
            Height          =   510
            Index           =   16
            Left            =   660
            TabIndex        =   78
            Tag             =   "lblPrefixString"
            Top             =   6015
            Width           =   2175
         End
         Begin VB.Label lblPosition 
            Caption         =   "Landscape Top Y pos :"
            Height          =   510
            Index           =   15
            Left            =   420
            TabIndex        =   76
            Tag             =   "lblPrefixString"
            Top             =   4905
            Width           =   2175
         End
         Begin VB.Label lblPosition 
            Caption         =   "Landscape Left X pos :"
            Height          =   510
            Index           =   14
            Left            =   420
            TabIndex        =   74
            Tag             =   "lblPrefixString"
            Top             =   4455
            Width           =   2175
         End
         Begin VB.Label lblPosition 
            Caption         =   "Locked in Landscape :"
            Height          =   435
            Index           =   13
            Left            =   450
            TabIndex        =   72
            Tag             =   "lblAlarmSound"
            Top             =   3975
            Width           =   2115
         End
         Begin VB.Label lblPosition 
            Caption         =   $"frmPrefs.frx":6827
            Height          =   3435
            Index           =   12
            Left            =   5145
            TabIndex        =   70
            Tag             =   "lblAlarmSoundDesc"
            Top             =   3480
            Width           =   2520
         End
         Begin VB.Label lblPosition 
            Caption         =   "Locked in Portrait :"
            Height          =   375
            Index           =   11
            Left            =   690
            TabIndex        =   69
            Tag             =   "lblAlarmSound"
            Top             =   5550
            Width           =   2040
         End
         Begin VB.Label lblPosition 
            Caption         =   $"frmPrefs.frx":69F9
            Height          =   705
            Index           =   10
            Left            =   2250
            TabIndex        =   67
            Tag             =   "lblAlarmSoundDesc"
            Top             =   2550
            Width           =   5325
         End
         Begin VB.Label lblPosition 
            Caption         =   "Widget Position by Percent:"
            Height          =   375
            Index           =   8
            Left            =   195
            TabIndex        =   66
            Tag             =   "lblAlarmSound"
            Top             =   2145
            Width           =   2355
         End
         Begin VB.Label lblPosition 
            Caption         =   $"frmPrefs.frx":6A98
            Height          =   3045
            Index           =   6
            Left            =   2265
            TabIndex        =   64
            Tag             =   "lblAlarmSoundDesc"
            Top             =   450
            Width           =   5370
         End
         Begin VB.Label lblPosition 
            Caption         =   "Aspect Ratio Hidden Mode :"
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   63
            Tag             =   "lblAlarmSound"
            Top             =   45
            Width           =   2145
         End
      End
   End
   Begin VB.Label lblDragCorner 
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   8670
      TabIndex        =   149
      ToolTipText     =   "drag me"
      Top             =   10335
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label lblSize 
      Caption         =   "Size in twips"
      Height          =   285
      Left            =   1875
      TabIndex        =   145
      Top             =   9780
      Visible         =   0   'False
      Width           =   4170
   End
   Begin VB.Label lblAsterix 
      Caption         =   "All controls marked with a * take effect immediately."
      Height          =   300
      Left            =   1920
      TabIndex        =   127
      Top             =   10155
      Width           =   3870
   End
   Begin VB.Menu prefsMnuPopmenu 
      Caption         =   "The main menu"
      Visible         =   0   'False
      Begin VB.Menu mnuAbout 
         Caption         =   "About Panzer Earth Widget"
      End
      Begin VB.Menu blank 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCoffee 
         Caption         =   "Donate a coffee with KoFi"
      End
      Begin VB.Menu mnuSupport 
         Caption         =   "Contact Support"
      End
      Begin VB.Menu blank3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuButton 
         Caption         =   "Theme Colours"
         Begin VB.Menu mnuLight 
            Caption         =   "Light Theme Enable"
         End
         Begin VB.Menu mnuDark 
            Caption         =   "High Contrast Theme Enable"
         End
         Begin VB.Menu mnuAuto 
            Caption         =   "Auto Theme Selection"
         End
      End
      Begin VB.Menu mnuLicenceA 
         Caption         =   "Display Licence Agreement"
      End
      Begin VB.Menu mnuClosePreferences 
         Caption         =   "Close Preferences"
      End
   End
End
Attribute VB_Name = "widgetPrefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule IntegerDataType, ModuleWithoutFolder

' clockForm_BubblingEvent ' leaving that here so I can copy/paste to find it

'---------------------------------------------------------------------------------------
' Module    : widgetPrefs
' Author    : beededea
' Date      : 17/08/2022
' Purpose   : VB6 standard form to display the prefs
'---------------------------------------------------------------------------------------

Option Explicit

'------------------------------------------------------ STARTS
' Constants and APIs to create and subclass the dragCorner
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTBOTTOMRIGHT = 17
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
' Constants defined for setting a theme to the prefs
Private Const COLOR_BTNFACE As Long = 15

' APIs declared for setting a theme to the prefs
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function IsThemeActive Lib "uxtheme" () As Boolean
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
' Private Types for determining prefs sizing
Private gblPrefsLoadedFlg As Boolean
Private prefsDynamicSizingFlg As Boolean
Private lastFormHeight As Long
Private Const cPrefsFormHeight As Long = 11055
Private Const cPrefsFormWidth  As Long = 9090
'------------------------------------------------------ ENDS

Private prefsStartupFlg As Boolean
Private gblAllowSizeChangeFlg As Boolean




'---------------------------------------------------------------------------------------
' Procedure : btnVerifyDateTime1_Click
' Author    : beededea
' Date      : 24/07/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnVerifyDateTime1_Click()
    Dim alarmTimeStatus As Boolean: alarmTimeStatus = False
    Dim alarmDateStatus As Boolean: alarmDateStatus = False
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString
    
    On Error GoTo btnVerifyDateTime1_Click_Error
   
    If txtAlarm1Date.Text = "Alarm not yet set" Then
    
        answerMsg = "Alarm not yet set!"
        answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Message", True, "btnVerifyDateTime1_Click1")
        
        Exit Sub
    End If
    
    If txtAlarm1Date.Text <> vbNullString Then
        alarmDateStatus = fVerifyAlarmDate(txtAlarm1Date.Text)
        alarmTimeStatus = fVerifyAlarmTime(txtAlarm1Date.Text, txtAlarm1Time.Text)
    End If
    If alarmDateStatus = True Then
        txtAlarm1Date.BackColor = vbWhite
    Else
        txtAlarm1Date.BackColor = vbRed
        
        answerMsg = "Alarm date format invalid"
        answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Message", True, "btnVerifyDateTime1_Click2")
        
        Exit Sub
    End If
    If alarmTimeStatus = True Then
        txtAlarm1Time.BackColor = vbWhite
    Else
        txtAlarm1Time.BackColor = vbRed
        
        answerMsg = "Alarm time format invalid"
        answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Message", True, "btnVerifyDateTime1_Click3")
                
    End If
    
    If alarmDateStatus = True And alarmTimeStatus = True Then
    
        answerMsg = "Alarm date and time formats both valid and in the future"
        answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Information", True, "btnVerifyDateTime1_Click4")
        
    End If
    
   On Error GoTo 0
   Exit Sub

btnVerifyDateTime1_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnVerifyDateTime1_Click of Form widgetPrefs"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : btnVerifyDateTime2_Click
' Author    : beededea
' Date      : 24/07/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnVerifyDateTime2_Click()
    Dim alarmTimeStatus As Boolean: alarmTimeStatus = False
    Dim alarmDateStatus As Boolean: alarmDateStatus = False
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString
    
    On Error GoTo btnVerifyDateTime2_Click_Error
   
    If txtAlarm2Date.Text = "Alarm not yet set" Then
    
        answerMsg = "Alarm not yet set!"
        answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Message", True, "btnVerifyDateTime2_Click1")
        
        Exit Sub
    End If
    
    If txtAlarm2Date.Text <> vbNullString Then
        alarmDateStatus = fVerifyAlarmDate(txtAlarm2Date.Text)
        alarmTimeStatus = fVerifyAlarmTime(txtAlarm2Date.Text, txtAlarm2Time.Text)
    End If
    If alarmDateStatus = True Then
        txtAlarm2Date.BackColor = vbWhite
    Else
        txtAlarm2Date.BackColor = vbRed
        
        answerMsg = "Alarm date format invalid"
        answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Message", True, "btnVerifyDateTime2_Click2")
        
        Exit Sub
    End If
    If alarmTimeStatus = True Then
        txtAlarm2Time.BackColor = vbWhite
    Else
        txtAlarm2Time.BackColor = vbRed
        
        answerMsg = "Alarm time format invalid"
        answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Message", True, "btnVerifyDateTime2_Click3")
                
    End If
    
    If alarmDateStatus = True And alarmTimeStatus = True Then
    
        answerMsg = "Alarm date and time formats both valid and in the future"
        answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Information", True, "btnVerifyDateTime2_Click4")
        
    End If
    
   On Error GoTo 0
   Exit Sub

btnVerifyDateTime2_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnVerifyDateTime2_Click of Form widgetPrefs"
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : fVerifyAlarmDate
' Author    : beededea
' Date      : 24/07/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function fVerifyAlarmDate(ByVal datefield As String) As Boolean
    Dim goodDate As Boolean: goodDate = False
    Dim futureDate As Double: futureDate = 0
    
   On Error GoTo fVerifyAlarmDate_Error

    goodDate = IsDate(datefield)
    If goodDate = False Then Exit Function
    
    futureDate = DateDiff("s", Date, datefield)
    If futureDate < 0 Then Exit Function
    
    fVerifyAlarmDate = True

   On Error GoTo 0
   Exit Function

fVerifyAlarmDate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fVerifyAlarmDate of Form widgetPrefs"

End Function
'---------------------------------------------------------------------------------------
' Procedure : fVerifyAlarmTime
' Author    : beededea
' Date      : 24/07/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function fVerifyAlarmTime(ByVal datefield As String, ByVal timefield As String) As Boolean
    Dim goodTime As Boolean: goodTime = False
    Dim futureTime As Double: futureTime = 0
    
    On Error GoTo fVerifyAlarmTime_Error

    goodTime = IsDate(timefield)
    If goodTime = False Then Exit Function
    
    futureTime = DateDiff("s", Time, datefield & " " & timefield)
    If futureTime < 0 Then Exit Function
    
    fVerifyAlarmTime = True

   On Error GoTo 0
   Exit Function

fVerifyAlarmTime_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fVerifyAlarmTime of Form widgetPrefs"

End Function


'---------------------------------------------------------------------------------------
' Procedure : btnVerifyDateTime3_Click
' Author    : beededea
' Date      : 24/07/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnVerifyDateTime3_Click()
    Dim alarmTimeStatus As Boolean: alarmTimeStatus = False
    Dim alarmDateStatus As Boolean: alarmDateStatus = False
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString
    
    On Error GoTo btnVerifyDateTime3_Click_Error
   
    If txtAlarm3Date.Text = "Alarm not yet set" Then
    
        answerMsg = "Alarm not yet set!"
        answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Message", True, "btnVerifyDateTime3_Click1")
        
        Exit Sub
    End If
    
    If txtAlarm3Date.Text <> vbNullString Then
        alarmDateStatus = fVerifyAlarmDate(txtAlarm3Date.Text)
        alarmTimeStatus = fVerifyAlarmTime(txtAlarm3Date.Text, txtAlarm3Time.Text)
    End If
    If alarmDateStatus = True Then
        txtAlarm3Date.BackColor = vbWhite
    Else
        txtAlarm3Date.BackColor = vbRed
        
        answerMsg = "Alarm date format invalid"
        answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Message", True, "btnVerifyDateTime3_Click2")
        
        Exit Sub
    End If
    If alarmTimeStatus = True Then
        txtAlarm3Time.BackColor = vbWhite
    Else
        txtAlarm3Time.BackColor = vbRed
        
        answerMsg = "Alarm time format invalid"
        answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Message", True, "btnVerifyDateTime3_Click3")
                
    End If
    
    If alarmDateStatus = True And alarmTimeStatus = True Then
    
        answerMsg = "Alarm date and time formats both valid and in the future"
        answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Information", True, "btnVerifyDateTime3_Click4")
        
    End If
    
   On Error GoTo 0
   Exit Sub

btnVerifyDateTime3_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnVerifyDateTime3_Click of Form widgetPrefs"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : btnVerifyDateTime4_Click
' Author    : beededea
' Date      : 24/07/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnVerifyDateTime4_Click()
    Dim alarmTimeStatus As Boolean: alarmTimeStatus = False
    Dim alarmDateStatus As Boolean: alarmDateStatus = False
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString
    
    On Error GoTo btnVerifyDateTime4_Click_Error
   
    If txtAlarm4Date.Text = "Alarm not yet set" Then
    
        answerMsg = "Alarm not yet set!"
        answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Message", True, "btnVerifyDateTime4_Click1")
        
        Exit Sub
    End If
    
    If txtAlarm4Date.Text <> vbNullString Then
        alarmDateStatus = fVerifyAlarmDate(txtAlarm4Date.Text)
        alarmTimeStatus = fVerifyAlarmTime(txtAlarm4Date.Text, txtAlarm4Time.Text)
    End If
    If alarmDateStatus = True Then
        txtAlarm4Date.BackColor = vbWhite
    Else
        txtAlarm4Date.BackColor = vbRed
        
        answerMsg = "Alarm date format invalid"
        answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Message", True, "btnVerifyDateTime4_Click2")
        
        Exit Sub
    End If
    If alarmTimeStatus = True Then
        txtAlarm4Time.BackColor = vbWhite
    Else
        txtAlarm4Time.BackColor = vbRed
        
        answerMsg = "Alarm time format invalid"
        answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Message", True, "btnVerifyDateTime4_Click3")
                
    End If
    
    If alarmDateStatus = True And alarmTimeStatus = True Then
    
        answerMsg = "Alarm date and time formats both valid and in the future"
        answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Information", True, "btnVerifyDateTime4_Click4")
        
    End If
    
   On Error GoTo 0
   Exit Sub

btnVerifyDateTime4_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnVerifyDateTime4_Click of Form widgetPrefs"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : btnVerifyDateTime5_Click
' Author    : beededea
' Date      : 24/07/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnVerifyDateTime5_Click()
    Dim alarmTimeStatus As Boolean: alarmTimeStatus = False
    Dim alarmDateStatus As Boolean: alarmDateStatus = False
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString
    
    On Error GoTo btnVerifyDateTime5_Click_Error
   
    If txtAlarm5Date.Text = "Alarm not yet set" Then
    
        answerMsg = "Alarm not yet set!"
        answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Message", True, "btnVerifyDateTime5_Click1")
        
        Exit Sub
    End If
    
    If txtAlarm5Date.Text <> vbNullString Then
        alarmDateStatus = fVerifyAlarmDate(txtAlarm5Date.Text)
        alarmTimeStatus = fVerifyAlarmTime(txtAlarm5Date.Text, txtAlarm5Time.Text)
    End If
    If alarmDateStatus = True Then
        txtAlarm5Date.BackColor = vbWhite
    Else
        txtAlarm5Date.BackColor = vbRed
        
        answerMsg = "Alarm date format invalid"
        answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Message", True, "btnVerifyDateTime5_Click2")
        
        Exit Sub
    End If
    If alarmTimeStatus = True Then
        txtAlarm5Time.BackColor = vbWhite
    Else
        txtAlarm5Time.BackColor = vbRed
        
        answerMsg = "Alarm time format invalid"
        answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Message", True, "btnVerifyDateTime5_Click3")
                
    End If
    
    If alarmDateStatus = True And alarmTimeStatus = True Then
    
        answerMsg = "Alarm date and time formats both valid and in the future"
        answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Information", True, "btnVerifyDateTime5_Click4")
        
    End If
    
   On Error GoTo 0
   Exit Sub

btnVerifyDateTime5_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnVerifyDateTime5_Click of Form widgetPrefs"
    
End Sub

Private Sub chkEnableTicks_Click()
    btnSave.Enabled = True ' enable the save button
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkShowHelp_Click
' Author    : beededea
' Date      : 03/07/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkShowHelp_Click()
   On Error GoTo chkShowHelp_Click_Error

    btnSave.Enabled = True ' enable the save button
    If chkShowHelp.Value = 1 Then
        gblShowHelp = "1"
    Else
        gblShowHelp = "0"
    End If

   On Error GoTo 0
   Exit Sub

chkShowHelp_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkShowHelp_Click of Form widgetPrefs"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : chkTogglePendulum_Click
' Author    : beededea
' Date      : 29/07/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkTogglePendulum_Click()
    
    btnSave.Enabled = True ' enable the save button
    
    If chkTogglePendulum.Value = 1 Then
        overlayWidget.SwingPendulum = True
        gblTogglePendulum = "1"
    Else
        overlayWidget.SwingPendulum = False
        gblTogglePendulum = "0"
    End If
    
    sPutINISetting "Software\SteampunkClockCalendar", "togglePendulum", gblTogglePendulum, gblSettingsFile


   On Error GoTo 0
   Exit Sub

chkTogglePendulum_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkTogglePendulum_Click of Form widgetPrefs"

End Sub

' ----------------------------------------------------------------
' Procedure Name: Form_Initialize
' Purpose:
' Procedure Kind: Constructor (Initialize)
' Procedure Access: Private
' Author: beededea
' Date: 05/10/2023
' ----------------------------------------------------------------
Private Sub Form_Initialize()
    On Error GoTo Form_Initialize_Error
    
    gblPrefsLoadedFlg = False
    prefsDynamicSizingFlg = False
    lastFormHeight = 0

    On Error GoTo 0
    Exit Sub

Form_Initialize_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Initialize of Form widgetPrefs"
    
    End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : beededea
' Date      : 25/04/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()
    Dim prefsFormHeight As Long: prefsFormHeight = 0

    On Error GoTo Form_Load_Error
        
    prefsStartupFlg = True ' this is used to prevent some control initialisations from running code at startup
    prefsDynamicSizingFlg = False
    gblPrefsLoadedFlg = True ' this is a variable tested by an added form property to indicate whether the form is loaded or not
    gblWindowLevelWasChanged = False
    prefsFormHeight = prefsCurrentHeight
    
    With lblDragCorner
      .ForeColor = &H80000015
      .BackStyle = vbTransparent
      .AutoSize = True
      .Font.Size = 12
      .Font.Name = "Marlett"
      .Caption = "o"
      .Font.Bold = False
      .Visible = False
    End With
     
    btnSave.Enabled = False ' disable the save button

    If gblDpiAwareness = "1" Then
        prefsDynamicSizingFlg = True
        chkEnableResizing.Value = 1
        lblDragCorner.Visible = True
    End If
      
    ' read the last saved position from the settings.ini
    Call readPrefsPosition
        
    ' determine the frame heights in dynamic sizing or normal mode
    Call setframeHeights
    
    ' set the tooltips on the prefs screen
    Call setPrefsTooltips
    
    ' set the text in any labels that need a vbCrLf to space the text
    Call setPrefsLabels
    
    ' populate all the comboboxes in the prefs form
    Call populatePrefsComboBoxes
        
    ' adjust all the preferences and main program controls
    Call adjustPrefsControls
    
    ' adjust the theme used by the prefs alone
    Call adjustPrefsTheme
    
    ' size and position the frames and buttons
    Call positionPrefsFramesButtons
    
    ' make the last used tab appear on startup
    Call showLastTab
    
    ' load the about text and load into prefs
    Call loadPrefsAboutText
    
    ' load the preference icons from a previously populated CC imageList
    Call loadHigherResPrefsImages
    
    ' now cause a form_resize event and set the height of the whole form
    If gblDpiAwareness = "1" Then
        If prefsFormHeight < screenHeightTwips Then
            Me.Height = prefsFormHeight
        Else
            Me.Height = screenHeightTwips - 1000
        End If
    End If
    
    ' position the prefs on the current monitor
    Call positionPrefsMonitor
    
    ' set the Z order of the prefs form
    Call setPrefsFormZordering
    
    ' start the timer that records the prefs position every 10 seconds
    positionTimer.Enabled = True
    
    ' end the startup by un-setting the start flag
    prefsStartupFlg = False
    
    btnSave.Enabled = False

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form widgetPrefs"

End Sub

'
'---------------------------------------------------------------------------------------
' Procedure : positionPrefsMonitor
' Author    : beededea
' Date      : 06/10/2023
' Purpose   : position the prefs on the curent monitor
'---------------------------------------------------------------------------------------
'
Public Sub positionPrefsMonitor()

    Dim formLeftTwips As Long: formLeftTwips = 0
    Dim formTopTwips As Long: formTopTwips = 0
    Dim monitorCount As Long: monitorCount = 0
    
    On Error GoTo positionPrefsMonitor_Error
    
    If gblDpiAwareness = "1" Then
        formLeftTwips = Val(gblFormHighDpiXPosTwips)
        formTopTwips = Val(gblFormHighDpiYPosTwips)
    Else
        formLeftTwips = Val(gblFormLowDpiXPosTwips)
        formTopTwips = Val(gblFormLowDpiYPosTwips)
    End If
    
    If formLeftTwips = 0 Then
        If ((fClock.clockForm.Left + fClock.clockForm.Width) * screenTwipsPerPixelX) + 200 + widgetPrefs.Width > screenWidthTwips Then
            widgetPrefs.Left = (fClock.clockForm.Left * screenTwipsPerPixelX) - (widgetPrefs.Width + 200)
        End If
    End If

    ' if a current location not stored then position to the middle of the screen
    If formLeftTwips <> 0 Then
        Me.Left = formLeftTwips
    Else
        Me.Left = screenWidthTwips / 2 - Me.Width / 2
    End If
    
    If formTopTwips <> 0 Then
        Me.Top = formTopTwips
    Else
        Me.Top = Screen.Height / 2 - Me.Height / 2
    End If
    
    monitorCount = fGetMonitorCount
    If monitorCount > 1 Then Call adjustFormPositionToCorrectMonitor(Me.hwnd, formLeftTwips, formTopTwips)
    
    ' calculate the on-screen widget position
    If Me.Left < 0 Then
        Me.Left = 10
    End If
    If Me.Top < 0 Then
        Me.Top = 0
    End If
    If Me.Left > screenWidthTwips - 2500 Then
        Me.Left = screenWidthTwips - 2500
    End If
    If Me.Top > screenHeightTwips - 2500 Then
        Me.Top = screenHeightTwips - 2500
    End If
    
    On Error GoTo 0
    Exit Sub

positionPrefsMonitor_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure positionPrefsMonitor of Form widgetPrefs"
End Sub
    
'---------------------------------------------------------------------------------------
' Procedure : btnResetMessages_MouseMove
' Author    : beededea
' Date      : 01/10/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnResetMessages_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo btnResetMessages_MouseMove_Error

    If gblEnableBalloonTooltips = "1" Then CreateToolTip btnResetMessages.hwnd, "The various pop-up messages that this program generates can be manually hidden. This button restores them to their original visible state.", _
                  TTIconInfo, "Help on the message reset button", , , , True

    On Error GoTo 0
    Exit Sub

btnResetMessages_MouseMove_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnResetMessages_MouseMove of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkDpiAwareness_Click
' Author    : beededea
' Date      : 14/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkDpiAwareness_Click()
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString

    On Error GoTo chkDpiAwareness_Click_Error

    btnSave.Enabled = True ' enable the save button
    If prefsStartupFlg = False Then ' don't run this on startup
                    
        answer = vbYes
        answerMsg = "You must close this widget and HARD restart it, in order to change the widget's DPI awareness (a simple soft reload just won't cut it), do you want me to close and restart this widget? I can do it now for you."
        answer = msgBoxA(answerMsg, vbYesNo, "DpiAwareness Confirmation", True, "chkDpiAwarenessRestart")
        
        If chkDpiAwareness.Value = 0 Then
            gblDpiAwareness = "0"
        Else
            gblDpiAwareness = "1"
        End If

        sPutINISetting "Software\SteampunkClockCalendar", "dpiAwareness", gblDpiAwareness, gblSettingsFile
        
        If answer = vbNo Then
            answer = vbYes
            answerMsg = "OK, the widget is still DPI aware until you restart. Some forms may show abnormally."
            answer = msgBoxA(answerMsg, vbOKOnly, "DpiAwareness Notification", True, "chkDpiAwarenessAbnormal")
        
            Exit Sub
        Else

            sPutINISetting "Software\SteampunkClockCalendar", "dpiAwareness", gblDpiAwareness, gblSettingsFile
            'Call reloadWidget ' this is insufficient, image controls still fail to resize and autoscale correctly
            Call hardRestart
        End If

    End If

   On Error GoTo 0
   Exit Sub

chkDpiAwareness_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkDpiAwareness_Click of Form widgetPrefs"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : chkEnablePrefsTooltips_Click
' Author    : beededea
' Date      : 07/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkEnablePrefsTooltips_Click()

   On Error GoTo chkEnablePrefsTooltips_Click_Error

    btnSave.Enabled = True ' enable the save button
    
    If prefsStartupFlg = False Then
        If chkEnablePrefsTooltips.Value = 1 Then
            gblEnablePrefsTooltips = "1"
        Else
            gblEnablePrefsTooltips = "0"
        End If
        
        sPutINISetting "Software\SteampunkClockCalendar", "enablePrefsTooltips", gblEnablePrefsTooltips, gblSettingsFile

    End If
    
    ' set the tooltips on the prefs screen
    Call setPrefsTooltips

   On Error GoTo 0
   Exit Sub

chkEnablePrefsTooltips_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkEnablePrefsTooltips_Click of Form widgetPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkEnableTooltips_MouseMove
' Author    : beededea
' Date      : 05/10/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkEnableTooltips_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo chkEnableTooltips_MouseMove_Error

    If gblEnableBalloonTooltips = "1" Then CreateToolTip chkEnableTooltips.hwnd, "There is a problem with the current tooltips on the clock itself as they resize along with the program graphical elements, meaning that they cannot be seen, there is also a problem with tooltip handling different fonts, hoping to get Olaf to fix these soon. My suggestion is to turn them off for the moment.", _
                  TTIconInfo, "Help on the Program tooltip problem", , , , True

    On Error GoTo 0
    Exit Sub

chkEnableTooltips_MouseMove_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkEnableTooltips_MouseMove of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkShowTaskbar_Click
' Author    : beededea
' Date      : 19/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkShowTaskbar_Click()

   On Error GoTo chkShowTaskbar_Click_Error

    btnSave.Enabled = True ' enable the save button
    If chkShowTaskbar.Value = 1 Then
        gblShowTaskbar = "1"
    Else
        gblShowTaskbar = "0"
    End If

   On Error GoTo 0
   Exit Sub

chkShowTaskbar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkShowTaskbar_Click of Form widgetPrefs"
End Sub





'---------------------------------------------------------------------------------------
' Procedure : btnResetMessages_Click
' Author    : beededea
' Date      : 01/10/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnResetMessages_Click()

    On Error GoTo btnResetMessages_Click_Error
        
    ' Clear all the message box "show again" entries in the registry
    Call clearAllMessageBoxRegistryEntries
    
    MsgBox "Message boxes fully reset, confirmation pop-ups will continue as normal."

    On Error GoTo 0
    Exit Sub

btnResetMessages_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnResetMessages_Click of Form widgetPrefs"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : btnAboutDebugInfo_Click
' Author    : beededea
' Date      : 03/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnAboutDebugInfo_Click()

   On Error GoTo btnAboutDebugInfo_Click_Error
   'If debugflg = 1 Then Debug.Print "%btnAboutDebugInfo_Click"

    'mnuDebug_Click
    MsgBox "The debug mode is not yet enabled."

   On Error GoTo 0
   Exit Sub

btnAboutDebugInfo_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnAboutDebugInfo_Click of form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnDonate_Click
' Author    : beededea
' Date      : 13/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnDonate_Click()
   On Error GoTo btnDonate_Click_Error

    Call mnuCoffee_ClickEvent

   On Error GoTo 0
   Exit Sub

btnDonate_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnDonate_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnFacebook_Click
' Author    : beededea
' Date      : 29/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnFacebook_Click()
   On Error GoTo btnFacebook_Click_Error
   'If debugflg = 1 Then DebugPrint "%btnFacebook_Click"

    Call menuForm.mnuFacebook_Click
    

   On Error GoTo 0
   Exit Sub

btnFacebook_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnFacebook_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnOpenFile_Click
' Author    : beededea
' Date      : 02/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnOpenFile_Click()
    Dim retFileName As String: retFileName = vbNullString
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString

    On Error GoTo btnOpenFile_Click_Error


    
    Call addTargetFile(txtOpenFile.Text, retFileName)
    
    If retFileName <> vbNullString Then
        txtOpenFile.Text = retFileName ' strips the buffered bit, leaving just the filename
    End If
    
    If retFileName = vbNullString Then
        Exit Sub
    End If
    
    If Not fFExists(retFileName) Then
        'answer = MsgBox("The file doesn't currently exist, do you want me to create the chosen file, " & "   -  are you sure?", vbYesNo)
        answer = vbYes
        answerMsg = "The file doesn't currently exist, do you want me to create the chosen file, " & "   -  are you sure?"
        answer = msgBoxA(answerMsg, vbYesNo, "Create file confirmation", False)
        If answer = vbNo Then
            Exit Sub
        End If
    
        'create new
        Open retFileName For Output As #1
        Close #1
    End If

    On Error GoTo 0
    Exit Sub

btnOpenFile_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnOpenFile_Click of Form widgetPrefs"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnUpdate_Click
' Author    : beededea
' Date      : 29/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnUpdate_Click()
   On Error GoTo btnUpdate_Click_Error
   'If debugflg = 1 Then DebugPrint "%btnUpdate_Click"

    'MsgBox "The update button is not yet enabled."
    menuForm.mnuLatest_Click

   On Error GoTo 0
   Exit Sub

btnUpdate_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnUpdate_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkWidgetFunctions_Click
' Author    : beededea
' Date      : 30/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkWidgetFunctions_Click()
    On Error GoTo chkWidgetFunctions_Click_Error

    btnSave.Enabled = True ' enable the save button

    On Error GoTo 0
    Exit Sub

chkWidgetFunctions_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkWidgetFunctions_Click of Form widgetPrefs"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : chkGenStartup_Click
' Author    : beededea
' Date      : 30/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkGenStartup_Click()
    On Error GoTo chkGenStartup_Click_Error

    btnSave.Enabled = True ' enable the save button

    On Error GoTo 0
    Exit Sub

chkGenStartup_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkGenStartup_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnDefaultEditor_Click
' Author    : beededea
' Date      : 02/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnDefaultEditor_Click()
    Dim retFileName As String: retFileName = vbNullString
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString

    On Error GoTo btnDefaultEditor_Click_Error

    Call addTargetFile(txtDefaultEditor.Text, retFileName)
    
    If retFileName <> vbNullString Then
        txtDefaultEditor.Text = retFileName ' strips the buffered bit, leaving just the filename
    End If
    
    If retFileName = vbNullString Then
        Exit Sub
    End If
    
    If Not fFExists(retFileName) Then
        answer = vbYes
        answerMsg = "The file doesn't currently exist, do you want me to create the chosen file, " & "   -  are you sure?"
        answer = msgBoxA(answerMsg, vbYesNo, "Default Editor Confirmation", False)
        If answer = vbNo Then
            Exit Sub
        End If
    
        'create new
        Open retFileName For Output As #1
        Close #1
    End If

    On Error GoTo 0
    Exit Sub

btnDefaultEditor_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnDefaultEditor_Click of Form widgetPrefs"
            Resume Next
          End If
    End With
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkEnableBalloonTooltips_Click
' Author    : beededea
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkEnableBalloonTooltips_Click()
   On Error GoTo chkEnableBalloonTooltips_Click_Error

    btnSave.Enabled = True ' enable the save button
    If chkEnableBalloonTooltips.Value = 1 Then
        gblEnableBalloonTooltips = "1"
    Else
        gblEnableBalloonTooltips = "0"
    End If

   On Error GoTo 0
   Exit Sub

chkEnableBalloonTooltips_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkEnableBalloonTooltips_Click of Form widgetPrefs"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : chkIgnoreMouse_Click
' Author    : beededea
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkIgnoreMouse_Click()
   On Error GoTo chkIgnoreMouse_Click_Error

    If chkIgnoreMouse.Value = 0 Then
        gblIgnoreMouse = "0"
    Else
        gblIgnoreMouse = "1"
    End If

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

chkIgnoreMouse_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkIgnoreMouse_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkPreventDragging_Click
' Author    : beededea
' Date      : 02/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkPreventDragging_Click()
    On Error GoTo chkPreventDragging_Click_Error

    btnSave.Enabled = True ' enable the save button
    ' immediately make the widget locked in place
    If chkPreventDragging.Value = 0 Then
        overlayWidget.Locked = False
        gblPreventDragging = "0"
        menuForm.mnuLockWidget.Checked = False
        If aspectRatio = "landscape" Then
            txtLandscapeHoffset.Text = vbNullString
            txtLandscapeVoffset.Text = vbNullString
        Else
            txtPortraitHoffset.Text = vbNullString
            txtPortraitYoffset.Text = vbNullString
        End If
    Else
        overlayWidget.Locked = True
        gblPreventDragging = "1"
        menuForm.mnuLockWidget.Checked = True
        If aspectRatio = "landscape" Then
            txtLandscapeHoffset.Text = fClock.clockForm.Left
            txtLandscapeVoffset.Text = fClock.clockForm.Top
        Else
            txtPortraitHoffset.Text = fClock.clockForm.Left
            txtPortraitYoffset.Text = fClock.clockForm.Top
        End If
    End If

    On Error GoTo 0
    Exit Sub

chkPreventDragging_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkPreventDragging_Click of Form widgetPrefs"
            Resume Next
          End If
    End With
    
End Sub



'---------------------------------------------------------------------------------------
' Procedure : chkWidgetHidden_Click
' Author    : beededea
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkWidgetHidden_Click()
   On Error GoTo chkWidgetHidden_Click_Error

    If chkWidgetHidden.Value = 0 Then
        'overlayWidget.Hidden = False
        fClock.clockForm.Visible = True

        frmTimer.revealWidgetTimer.Enabled = False
        gblWidgetHidden = "0"
    Else
        'overlayWidget.Hidden = True
        fClock.clockForm.Visible = False


        frmTimer.revealWidgetTimer.Enabled = True
        gblWidgetHidden = "1"
    End If
    
    sPutINISetting "Software\SteampunkClockCalendar", "widgetHidden", gblWidgetHidden, gblSettingsFile
    
    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

chkWidgetHidden_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkWidgetHidden_Click of Form widgetPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbAspectHidden_Click
' Author    : beededea
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbAspectHidden_Click()

   On Error GoTo cmbAspectHidden_Click_Error

    If cmbAspectHidden.ListIndex = 1 And aspectRatio = "portrait" Then
        'overlayWidget.Hidden = True
        fClock.clockForm.Visible = False
    ElseIf cmbAspectHidden.ListIndex = 2 And aspectRatio = "landscape" Then
        'overlayWidget.Hidden = True
        fClock.clockForm.Visible = False
    Else
        'overlayWidget.Hidden = False
        fClock.clockForm.Visible = True
    End If

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

cmbAspectHidden_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbAspectHidden_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbDebug_Click
' Author    : beededea
' Date      : 02/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbDebug_Click()
    On Error GoTo cmbDebug_Click_Error

    btnSave.Enabled = True ' enable the save button
    If cmbDebug.ListIndex = 0 Then
        txtDefaultEditor.Text = "eg. E:\vb6\Steampunk Clock Calendar\Steampunk Clock Calendar.vbp"
        txtDefaultEditor.Enabled = False
        lblDebug(7).Enabled = False
        btnDefaultEditor.Enabled = False
        lblDebug(9).Enabled = False
    Else
        txtDefaultEditor.Text = gblDefaultEditor
        txtDefaultEditor.Enabled = True
        lblDebug(7).Enabled = True
        btnDefaultEditor.Enabled = True
        lblDebug(9).Enabled = True
    End If

    On Error GoTo 0
    Exit Sub

cmbDebug_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbDebug_Click of Form widgetPrefs"
            Resume Next
          End If
    End With

End Sub



Private Sub cmbHidingTime_Click()
    btnSave.Enabled = True ' enable the save button
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbScrollWheelDirection_Click
' Author    : beededea
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbScrollWheelDirection_Click()
   On Error GoTo cmbScrollWheelDirection_Click_Error

    btnSave.Enabled = True ' enable the save button
    'overlayWidget.ZoomDirection = cmbScrollWheelDirection.List(cmbScrollWheelDirection.ListIndex)

   On Error GoTo 0
   Exit Sub

cmbScrollWheelDirection_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbScrollWheelDirection_Click of Form widgetPrefs"
End Sub



Private Sub cmbWidgetLandscape_Click()
    btnSave.Enabled = True ' enable the save button
End Sub

Private Sub cmbWidgetPortrait_Click()
    btnSave.Enabled = True ' enable the save button
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbWidgetPosition_Click
' Author    : beededea
' Date      : 02/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbWidgetPosition_Click()
    On Error GoTo cmbWidgetPosition_Click_Error

    btnSave.Enabled = True ' enable the save button
    If cmbWidgetPosition.ListIndex = 1 Then
        cmbWidgetLandscape.ListIndex = 0
        cmbWidgetPortrait.ListIndex = 0
        cmbWidgetLandscape.Enabled = False
        cmbWidgetPortrait.Enabled = False
        txtLandscapeHoffset.Enabled = False
        txtLandscapeVoffset.Enabled = False
        txtPortraitHoffset.Enabled = False
        txtPortraitYoffset.Enabled = False
        
    Else
        cmbWidgetLandscape.Enabled = True
        cmbWidgetPortrait.Enabled = True
        txtLandscapeHoffset.Enabled = True
        txtLandscapeVoffset.Enabled = True
        txtPortraitHoffset.Enabled = True
        txtPortraitYoffset.Enabled = True
    End If

    On Error GoTo 0
    Exit Sub

cmbWidgetPosition_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbWidgetPosition_Click of Form widgetPrefs"
            Resume Next
          End If
    End With
End Sub




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

    If gblPrefsLoadedFlg Then
        If Me.WindowState = vbNormal Then
            IsVisible = Me.Visible
        Else
            IsVisible = False
        End If
    Else
        IsVisible = False
    End If

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


'---------------------------------------------------------------------------------------
' Procedure : showLastTab
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : make the last used tab appear on startup
'---------------------------------------------------------------------------------------
'
Private Sub showLastTab()

   On Error GoTo showLastTab_Error

    If gblLastSelectedTab = "general" Then Call picButtonMouseUpEvent("general", imgGeneral, imgGeneralClicked, fraGeneral, fraGeneralButton)  ' was imgGeneralMouseUpEvent
    If gblLastSelectedTab = "config" Then Call picButtonMouseUpEvent("config", imgConfig, imgConfigClicked, fraConfig, fraConfigButton)     ' was imgConfigMouseUpEvent
    If gblLastSelectedTab = "position" Then Call picButtonMouseUpEvent("position", imgPosition, imgPositionClicked, fraPosition, fraPositionButton)
    If gblLastSelectedTab = "development" Then Call picButtonMouseUpEvent("development", imgDevelopment, imgDevelopmentClicked, fraDevelopment, fraDevelopmentButton)
    If gblLastSelectedTab = "fonts" Then Call picButtonMouseUpEvent("fonts", imgFonts, imgFontsClicked, fraFonts, fraFontsButton)
    If gblLastSelectedTab = "sounds" Then Call picButtonMouseUpEvent("sounds", imgSounds, imgSoundsClicked, fraSounds, fraSoundsButton)
    If gblLastSelectedTab = "window" Then Call picButtonMouseUpEvent("window", imgWindow, imgWindowClicked, fraWindow, fraWindowButton)
    If gblLastSelectedTab = "about" Then Call picButtonMouseUpEvent("about", imgAbout, imgAboutClicked, fraAbout, fraAboutButton)

   On Error GoTo 0
   Exit Sub

showLastTab_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure showLastTab of Form widgetPrefs"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : positionPrefsFramesButtons
' Author    : beededea
' Date      : 01/05/2023
' Purpose   : size and position the frames and buttons. Note we are NOT using control
'             arrays so the form can be converted to Cairo forms later.
'---------------------------------------------------------------------------------------
'
' for the future, when reading multiple buttons from XML config.
' read the XML prefs group and identify prefgroups - <prefGroup name="general" and count them.
'
' for each group read all the controls and identify those in the group - ie. preference group =
' for each specific group, identify the group image, title and order
' read those into an array
' use a for-loop (can't use foreach unless you read the results into a collection, foreach requires use of variant
'   elements as foreach needs an object or variant type to operate.
' create a group, fraHiding, image and text element and order in a class of yWidgetGroup
' create a button of yWidgetGroup for each group
' loop through each line and identify the controls belonging to the group

' for the moment though, we will do it manually
'
Private Sub positionPrefsFramesButtons()
    On Error GoTo positionPrefsFramesButtons_Error

    Dim frameWidth As Integer: frameWidth = 0
    Dim frameTop As Integer: frameTop = 0
    Dim frameLeft As Integer: frameLeft = 0
    Dim buttonTop As Integer:    buttonTop = 0
    'Dim currentFrameHeight As Integer: currentFrameHeight = 0
    Dim rightHandAlignment As Long: rightHandAlignment = 0
    Dim leftHandGutterWidth As Long: leftHandGutterWidth = 0
    
    ' align frames rightmost and leftmost to the buttons at the top
    buttonTop = -15
    frameTop = 1150
    leftHandGutterWidth = 240
    frameLeft = leftHandGutterWidth ' use the first frame leftmost as reference
    rightHandAlignment = fraAboutButton.Left + fraAboutButton.Width ' use final button rightmost as reference
    frameWidth = rightHandAlignment - frameLeft
    fraScrollbarCover.Left = rightHandAlignment - 690
    Me.Width = rightHandAlignment + leftHandGutterWidth + 75 ' (not quite sure why we need the 75 twips padding)
    
    ' align the top buttons
    fraGeneralButton.Top = buttonTop
    fraConfigButton.Top = buttonTop
    fraFontsButton.Top = buttonTop
    fraSoundsButton.Top = buttonTop
    fraPositionButton.Top = buttonTop
    fraDevelopmentButton.Top = buttonTop
    fraWindowButton.Top = buttonTop
    fraAboutButton.Top = buttonTop
    
    ' align the frames
    fraGeneral.Top = frameTop
    fraConfig.Top = frameTop
    fraFonts.Top = frameTop
    fraSounds.Top = frameTop
    fraPosition.Top = frameTop
    fraDevelopment.Top = frameTop
    fraWindow.Top = frameTop
    fraAbout.Top = frameTop
    
    fraGeneral.Left = frameLeft
    fraConfig.Left = frameLeft
    fraSounds.Left = frameLeft
    fraPosition.Left = frameLeft
    fraFonts.Left = frameLeft
    fraDevelopment.Left = frameLeft
    fraWindow.Left = frameLeft
    fraAbout.Left = frameLeft
    
    fraGeneral.Width = frameWidth
    fraConfig.Width = frameWidth
    fraSounds.Width = frameWidth
    fraPosition.Width = frameWidth
    fraFonts.Width = frameWidth
    fraWindow.Width = frameWidth
    fraDevelopment.Width = frameWidth
    fraAbout.Width = frameWidth
    
    ' set the base visibility of the frames
    fraGeneral.Visible = True
    fraConfig.Visible = False
    fraSounds.Visible = False
    fraPosition.Visible = False
    fraFonts.Visible = False
    fraWindow.Visible = False
    fraDevelopment.Visible = False
    fraAbout.Visible = False
    
    fraGeneralButton.BorderStyle = 1

    btnClose.Left = fraWindow.Left + fraWindow.Width - btnClose.Width
    btnSave.Left = btnClose.Left - btnSave.Width - 50
    btnHelp.Left = frameLeft
    

   On Error GoTo 0
   Exit Sub

positionPrefsFramesButtons_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure positionPrefsFramesButtons of Form widgetPrefs"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : btnClose_Click
' Author    : beededea
' Date      : 02/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnClose_Click()
   On Error GoTo btnClose_Click_Error

    btnSave.Enabled = False ' disable the save button
    Me.Hide
    Me.themeTimer.Enabled = False
    
    Call writePrefsPosition

   On Error GoTo 0
   Exit Sub

btnClose_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnClose_Click of Form widgetPrefs"
End Sub
'
'---------------------------------------------------------------------------------------
' Procedure : btnHelp_Click
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : display the help file
'---------------------------------------------------------------------------------------
'
Private Sub btnHelp_Click()
    
    On Error GoTo btnHelp_Click_Error
    
        If fFExists(App.path & "\help\Help.chm") Then
            Call ShellExecute(Me.hwnd, "Open", App.path & "\help\Help.chm", vbNullString, App.path, 1)
        Else
            MsgBox ("%Err-I-ErrorNumber 11 - The help file - Help.chm - is missing from the help folder.")
        End If

   On Error GoTo 0
   Exit Sub

btnHelp_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnHelp_Click of Form widgetPrefs"
End Sub
'
'---------------------------------------------------------------------------------------
' Procedure : btnSave_Click
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : save the values from all the tabs
'---------------------------------------------------------------------------------------
'
Private Sub btnSave_Click()

    Dim alarmTest As Boolean: alarmTest = False
    
    On Error GoTo btnSave_Click_Error

    ' configuration
    gblEnableTooltips = LTrim$(Str$(chkEnableTooltips.Value))
    gblEnablePrefsTooltips = LTrim$(Str$(chkEnablePrefsTooltips.Value))
    gblEnableBalloonTooltips = LTrim$(Str$(chkEnableBalloonTooltips.Value))
    gblShowTaskbar = LTrim$(Str$(chkShowTaskbar.Value))
    gblShowHelp = LTrim$(Str$(chkShowHelp.Value))
    
    gblTogglePendulum = LTrim$(Str$(chkTogglePendulum.Value))
    
    gblDpiAwareness = LTrim$(Str$(chkDpiAwareness.Value))
    gblGaugeSize = LTrim$(Str$(sliGaugeSize.Value))
    gblScrollWheelDirection = LTrim$(Str$(cmbScrollWheelDirection.ListIndex))
    
    ' general
    gblWidgetFunctions = LTrim$(Str$(chkWidgetFunctions.Value))
    gblStartup = LTrim$(Str$(chkGenStartup.Value))
    
    ' Validate all the alarm variables
    alarmTest = validateAlarmVars
    If alarmTest = False Then Exit Sub ' END the save if the alarms are malformed
    
    ' sounds
    gblEnableSounds = LTrim$(Str$(chkEnableSounds.Value))
    gblEnableTicks = LTrim$(Str$(chkEnableTicks.Value))
    
    'development
    gblDebug = LTrim$(Str$(cmbDebug.ListIndex))
    gblDblClickCommand = txtDblClickCommand.Text
    gblOpenFile = txtOpenFile.Text
    gblDefaultEditor = txtDefaultEditor.Text
    
    ' position
    gblAspectHidden = LTrim$(Str$(cmbAspectHidden.ListIndex))
    gblWidgetPosition = LTrim$(Str$(cmbWidgetPosition.ListIndex))
    gblWidgetLandscape = LTrim$(Str$(cmbWidgetLandscape.ListIndex))
    gblWidgetPortrait = LTrim$(Str$(cmbWidgetPortrait.ListIndex))
    gblLandscapeFormHoffset = txtLandscapeHoffset.Text
    gblLandscapeFormVoffset = txtLandscapeVoffset.Text
    gblPortraitHoffset = txtPortraitHoffset.Text
    gblPortraitYoffset = txtPortraitYoffset.Text
    
'    gblvLocationPercPrefValue
'    gblhLocationPercPrefValue

    ' fonts
    gblPrefsFont = txtPrefsFont.Text
    gblClockFont = gblPrefsFont
    
    ' the sizing is not saved here again as it saved during the setting phase.
    
'    If gblDpiAwareness = "1" Then
'        gblPrefsFontSizeHighDPI = txtPrefsFontSize.Text
'    Else
'        gblPrefsFontSizeLowDPI = txtPrefsFontSize.Text
'    End If
    'gblPrefsFontItalics = txtFontSize.Text

    ' Windows
    gblWindowLevel = LTrim$(Str$(cmbWindowLevel.ListIndex))
    gblPreventDragging = LTrim$(Str$(chkPreventDragging.Value))
    gblOpacity = LTrim$(Str$(sliOpacity.Value))
    gblWidgetHidden = LTrim$(Str$(chkWidgetHidden.Value))
    gblHidingTime = LTrim$(Str$(cmbHidingTime.ListIndex))
    gblIgnoreMouse = LTrim$(Str$(chkIgnoreMouse.Value))
            
    
    'development
    gblDebug = LTrim$(Str$(cmbDebug.ListIndex))
    gblDblClickCommand = txtDblClickCommand.Text
    gblOpenFile = txtOpenFile.Text
    gblDefaultEditor = txtDefaultEditor.Text
            
    If gblStartup = "1" Then
        Call savestring(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "SteampunkClockCalendar", """" & App.path & "\" & "Steampunk Clock Calendar.exe""")
    Else
        Call savestring(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "SteampunkClockCalendar", vbNullString)
    End If

    ' save the values from the general tab
    If fFExists(gblSettingsFile) Then
        sPutINISetting "Software\SteampunkClockCalendar", "enableTooltips", gblEnableTooltips, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "enablePrefsTooltips", gblEnablePrefsTooltips, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "enableBalloonTooltips", gblEnableBalloonTooltips, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "showTaskbar", gblShowTaskbar, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "showHelp", gblShowHelp, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "togglePendulum", gblTogglePendulum, gblSettingsFile
        
        sPutINISetting "Software\SteampunkClockCalendar", "dpiAwareness", gblDpiAwareness, gblSettingsFile
        
        
        sPutINISetting "Software\SteampunkClockCalendar", "gaugeSize", gblGaugeSize, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "scrollWheelDirection", gblScrollWheelDirection, gblSettingsFile
                
        sPutINISetting "Software\SteampunkClockCalendar", "widgetFunctions", gblWidgetFunctions, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "smoothSecondHand", gblSmoothSecondHand, gblSettingsFile
              
        sPutINISetting "Software\SteampunkClockCalendar", "aspectHidden", gblAspectHidden, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "widgetPosition", gblWidgetPosition, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "widgetLandscape", gblWidgetLandscape, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "widgetPortrait", gblWidgetPortrait, gblSettingsFile

        sPutINISetting "Software\SteampunkClockCalendar", "prefsFont", gblPrefsFont, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "clockFont", gblClockFont, gblSettingsFile
        
        sPutINISetting "Software\SteampunkClockCalendar", "prefsFontSizeHighDPI", gblPrefsFontSizeHighDPI, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "prefsFontSizeLowDPI", gblPrefsFontSizeLowDPI, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "prefsFontItalics", gblPrefsFontItalics, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "prefsFontColour", gblPrefsFontColour, gblSettingsFile

        'save the values from the Windows Config Items
        sPutINISetting "Software\SteampunkClockCalendar", "windowLevel", gblWindowLevel, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "preventDragging", gblPreventDragging, gblSettingsFile
        
        sPutINISetting "Software\SteampunkClockCalendar", "opacity", gblOpacity, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "widgetHidden", gblWidgetHidden, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "hidingTime", gblHidingTime, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "ignoreMouse", gblIgnoreMouse, gblSettingsFile
        
        sPutINISetting "Software\SteampunkClockCalendar", "startup", gblStartup, gblSettingsFile

        sPutINISetting "Software\SteampunkClockCalendar", "enableSounds", gblEnableSounds, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "enableTicks", gblEnableTicks, gblSettingsFile
        
        sPutINISetting "Software\SteampunkClockCalendar", "lastSelectedTab", gblLastSelectedTab, gblSettingsFile
        
        sPutINISetting "Software\SteampunkClockCalendar", "debug", gblDebug, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "dblClickCommand", gblDblClickCommand, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "openFile", gblOpenFile, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "defaultEditor", gblDefaultEditor, gblSettingsFile
        
        sPutINISetting "Software\SteampunkClockCalendar", "clockHighDpiXPos", gblClockHighDpiXPos, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "clockHighDpiYPos", gblClockHighDpiYPos, gblSettingsFile
        
        sPutINISetting "Software\SteampunkClockCalendar", "clockLowDpiXPos", gblClockLowDpiXPos, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "clockLowDpiYPos", gblClockLowDpiYPos, gblSettingsFile
        
        sPutINISetting "Software\SteampunkClockCalendar", "setToggleEnabled", gblsetToggleEnabled, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "muteToggleEnabled", gblMuteToggleEnabled, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "pendulumToggleEnabled", gblPendulumToggleEnabled, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "pendulumEnabled", gblPendulumEnabled, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "weekdayToggleEnabled", gblWeekdayToggleEnabled, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "displayScreenToggleEnabled", gblDisplayScreenToggleEnabled, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "timeMachineToggleEnabled", gblTimeMachineToggleEnabled, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "backToggleEnabled", gblBackToggleEnabled, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "clapperEnabled", gblClapperEnabled, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "chainEnabled", gblChainEnabled, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "crankEnabled", gblCrankEnabled, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "alarmToggle1Enabled", gblAlarmToggle1Enabled, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "alarmToggle2Enabled", gblAlarmToggle2Enabled, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "alarmToggle3Enabled", gblAlarmToggle3Enabled, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "alarmToggle4Enabled", gblAlarmToggle4Enabled, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "alarmToggle5Enabled", gblAlarmToggle5Enabled, gblSettingsFile
        
        sPutINISetting "Software\SteampunkClockCalendar", "alarm1Date", gblAlarm1Date, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "alarm2Date", gblAlarm2Date, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "alarm3Date", gblAlarm3Date, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "alarm4Date", gblAlarm4Date, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "alarm5Date", gblAlarm5Date, gblSettingsFile
               
        
        sPutINISetting "Software\SteampunkClockCalendar", "alarm1Time", gblAlarm1Time, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "alarm2Time", gblAlarm2Time, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "alarm3Time", gblAlarm3Time, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "alarm4Time", gblAlarm4Time, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "alarm5Time", gblAlarm5Time, gblSettingsFile
               
    End If
    
    ' set the tooltips on the prefs screen
    Call setPrefsTooltips

    ' sets the characteristics of the gauge and menus immediately after saving
    Call adjustMainControls
    
    Me.SetFocus
    btnSave.Enabled = False ' disable the save button showing it has successfully saved
    
    ' reload here if the gblWindowLevel Was Changed
    If gblWindowLevelWasChanged = True Then
        gblWindowLevelWasChanged = False
        Call reloadWidget
    End If
    
   On Error GoTo 0
   Exit Sub

btnSave_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnSave_Click of Form widgetPrefs"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : validateAlarmVars
' Author    : beededea
' Date      : 24/07/2024
' Purpose   : Validate all the alarm variables and FAIL if bad.
'---------------------------------------------------------------------------------------
'
Private Function validateAlarmVars() As Boolean
        
    Dim alarmTimeStatus As Boolean: alarmTimeStatus = False
    Dim alarmDateStatus As Boolean: alarmDateStatus = False
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString
     
    On Error GoTo validateAlarmVars_Error

    validateAlarmVars = True
        
    If txtAlarm1Date.Text = vbNullString Then txtAlarm1Date.Text = "Alarm not yet set"
    If txtAlarm2Date.Text = vbNullString Then txtAlarm2Date.Text = "Alarm not yet set"
    If txtAlarm3Date.Text = vbNullString Then txtAlarm3Date.Text = "Alarm not yet set"
    If txtAlarm4Date.Text = vbNullString Then txtAlarm4Date.Text = "Alarm not yet set"
    If txtAlarm5Date.Text = vbNullString Then txtAlarm5Date.Text = "Alarm not yet set"
        
    If txtAlarm1Date.Text <> "Alarm not yet set" Then
        alarmDateStatus = fVerifyAlarmDate(txtAlarm1Date.Text)
        alarmTimeStatus = fVerifyAlarmTime(txtAlarm1Date.Text, txtAlarm1Time.Text)
        If alarmDateStatus = False Or alarmTimeStatus = False Then
            btnSave.Enabled = False
            
            answerMsg = "Alarm no.1 is invalid, saving FAILED. Please correct and re-save."
            answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Message", True, "validateAlarmVars1")
            
            validateAlarmVars = False
            Exit Function
        End If
    End If
    gblAlarm1Date = txtAlarm1Date.Text
    gblAlarm1Time = txtAlarm1Time.Text
            
    If txtAlarm2Date.Text <> "Alarm not yet set" Then
        alarmDateStatus = fVerifyAlarmDate(txtAlarm2Date.Text)
        alarmTimeStatus = fVerifyAlarmTime(txtAlarm2Date.Text, txtAlarm2Time.Text)
        If alarmDateStatus = False Or alarmTimeStatus = False Then
            btnSave.Enabled = False
            
            answerMsg = "Alarm no.2 is invalid, saving FAILED. Please correct and re-save."
            answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Message", True, "validateAlarmVars2")
            
            validateAlarmVars = False
            Exit Function
        End If
    End If
    gblAlarm2Date = txtAlarm2Date.Text
    gblAlarm2Time = txtAlarm2Time.Text
    
    If txtAlarm3Date.Text <> "Alarm not yet set" Then
        alarmDateStatus = fVerifyAlarmDate(txtAlarm3Date.Text)
        alarmTimeStatus = fVerifyAlarmTime(txtAlarm3Date.Text, txtAlarm3Time.Text)
        If alarmDateStatus = False Or alarmTimeStatus = False Then
            btnSave.Enabled = False
            
            answerMsg = "Alarm no.3 is invalid, saving FAILED. Please correct and re-save."
            answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Message", True, "validateAlarmVars3")
            
            validateAlarmVars = False
            Exit Function
        End If
    End If
    gblAlarm3Date = txtAlarm3Date.Text
    gblAlarm3Time = txtAlarm3Time.Text
    
    If txtAlarm4Date.Text <> "Alarm not yet set" Then
        alarmDateStatus = fVerifyAlarmDate(txtAlarm4Date.Text)
        alarmTimeStatus = fVerifyAlarmTime(txtAlarm4Date.Text, txtAlarm4Time.Text)
        If alarmDateStatus = False Or alarmTimeStatus = False Then
            btnSave.Enabled = False
            
            answerMsg = "Alarm no.4 is invalid, saving FAILED. Please correct and re-save."
            answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Message", True, "validateAlarmVars4")
            
            validateAlarmVars = False
            Exit Function
        End If
    End If
    gblAlarm4Date = txtAlarm4Date.Text
    gblAlarm4Time = txtAlarm4Time.Text
            
    If txtAlarm5Date.Text <> "Alarm not yet set" Then
        alarmDateStatus = fVerifyAlarmDate(txtAlarm5Date.Text)
        alarmTimeStatus = fVerifyAlarmTime(txtAlarm5Date.Text, txtAlarm5Time.Text)
        If alarmDateStatus = False Or alarmTimeStatus = False Then
            btnSave.Enabled = False
            
            answerMsg = "Alarm no.5 is invalid, saving FAILED. Please correct and re-save."
            answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Message", True, "validateAlarmVars5")
            
            validateAlarmVars = False
            Exit Function
        End If
    End If
    gblAlarm5Date = txtAlarm5Date.Text
    gblAlarm5Time = txtAlarm5Time.Text

   On Error GoTo 0
   Exit Function

validateAlarmVars_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure validateAlarmVars of Form widgetPrefs"
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : chkEnableTooltips_Click
' Author    : beededea
' Date      : 19/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkEnableTooltips_Click()
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString
    On Error GoTo chkEnableTooltips_Click_Error

    btnSave.Enabled = True ' enable the save button
    
    If prefsStartupFlg = False Then
        If chkEnableTooltips.Value = 1 Then
            gblEnableTooltips = "1"
        Else
            gblEnableTooltips = "0"
        End If
        
        sPutINISetting "Software\SteampunkClockCalendar", "enableTooltips", gblEnableTooltips, gblSettingsFile

        answer = vbYes
        answerMsg = "You must soft reload this widget, in order to change the tooltip setting, do you want me to reload this widget? I can do it now for you."
        answer = msgBoxA(answerMsg, vbYesNo, "Request to Enable Tooltips", True, "chkEnableTooltipsClick")
        If answer = vbNo Then
            Exit Sub
        Else
            Call reloadWidget
        End If
    End If


   On Error GoTo 0
   Exit Sub

chkEnableTooltips_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkEnableTooltips_Click of Form widgetPrefs"

End Sub

Private Sub chkEnableSounds_Click()
    btnSave.Enabled = True ' enable the save button
End Sub

'Private Sub cmbRefreshInterval_Click()
'    btnSave.Enabled = True ' enable the save button
'End Sub

' ----------------------------------------------------------------
' Procedure Name: cmbWindowLevel_Click
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 28/05/2024
' ----------------------------------------------------------------
Private Sub cmbWindowLevel_Click()
    On Error GoTo cmbWindowLevel_Click_Error
    btnSave.Enabled = True ' enable the save button
    If prefsStartupFlg = False Then gblWindowLevelWasChanged = True
    
    On Error GoTo 0
    Exit Sub

cmbWindowLevel_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbWindowLevel_Click, line " & Erl & "."

End Sub
'---------------------------------------------------------------------------------------
' Procedure : btnPrefsFont_Click
' Author    : beededea
' Date      : 02/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPrefsFont_Click()

    Dim fntFont As String: fntFont = vbNullString
    Dim fntSize As Integer: fntSize = 0
    Dim fntWeight As Integer: fntWeight = 0
    Dim fntStyle As Boolean: fntStyle = False
    Dim fntColour As Long: fntColour = 0
    Dim fntItalics As Boolean: fntItalics = False
    Dim fntUnderline As Boolean: fntUnderline = False
    Dim fntFontResult As Boolean: fntFontResult = False
    
    On Error GoTo btnPrefsFont_Click_Error

    btnSave.Enabled = True ' enable the save button
    
    ' set the preliminary vars to feed and populate the changefont routine
    fntFont = gblPrefsFont
    ' gblClockFont
    
    If gblDpiAwareness = "1" Then
        fntSize = Val(gblPrefsFontSizeHighDPI)
    Else
        fntSize = Val(gblPrefsFontSizeLowDPI)
    End If
    
    If fntSize = 0 Then fntSize = 8
    fntItalics = CBool(gblPrefsFontItalics)
    fntColour = CLng(gblPrefsFontColour)
        
    Call changeFont(widgetPrefs, True, fntFont, fntSize, fntWeight, fntStyle, fntColour, fntItalics, fntUnderline, fntFontResult)
    
    gblPrefsFont = CStr(fntFont)
    gblClockFont = gblPrefsFont
    
    If gblDpiAwareness = "1" Then
        gblPrefsFontSizeHighDPI = CStr(fntSize)
        Call Form_Resize
    Else
        gblPrefsFontSizeLowDPI = CStr(fntSize)
    End If
    
    gblPrefsFontItalics = CStr(fntItalics)
    gblPrefsFontColour = CStr(fntColour)

    If fFExists(gblSettingsFile) Then ' does the tool's own settings.ini exist?
        sPutINISetting "Software\SteampunkClockCalendar", "prefsFont", gblPrefsFont, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "clockFont", gblClockFont, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "prefsFontSizeHighDPI", gblPrefsFontSizeHighDPI, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "prefsFontSizeLowDPI", gblPrefsFontSizeLowDPI, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "prefsFontItalics", gblPrefsFontItalics, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "PrefsFontColour", gblPrefsFontColour, gblSettingsFile
    End If
    
    If fntFont = vbNullString Then fntFont = "arial"
    txtPrefsFont.Text = fntFont
    txtPrefsFont.Font.Name = fntFont
    'txtPrefsFont.Font.Size = fntSize
    txtPrefsFont.Font.Italic = fntItalics
    txtPrefsFont.ForeColor = fntColour
    
    txtPrefsFontSize.Text = fntSize

   On Error GoTo 0
   Exit Sub

btnPrefsFont_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnPrefsFont_Click of Form widgetPrefs"

End Sub




'---------------------------------------------------------------------------------------
' Procedure : adjustPrefsControls
' Author    : beededea
' Date      : 12/05/2020
' Purpose   : adjust the controls so their startup position matches the last write of the config file
'---------------------------------------------------------------------------------------
'
Private Sub adjustPrefsControls()
    
    Dim fntWeight As Integer: fntWeight = 0
    Dim fntStyle As Boolean: fntStyle = False
    Dim sliGaugeSizeOldValue As Long: sliGaugeSizeOldValue = 0
    
    On Error GoTo adjustPrefsControls_Error
            
    ' general tab
    chkWidgetFunctions.Value = Val(gblWidgetFunctions)
    chkGenStartup.Value = Val(gblStartup)
        
    widgetPrefs.txtAlarm1Date.Text = gblAlarm1Date
    widgetPrefs.txtAlarm2Date.Text = gblAlarm2Date
    widgetPrefs.txtAlarm3Date.Text = gblAlarm3Date
    widgetPrefs.txtAlarm4Date.Text = gblAlarm4Date
    widgetPrefs.txtAlarm5Date.Text = gblAlarm5Date
        
    widgetPrefs.txtAlarm1Time.Text = gblAlarm1Time
    widgetPrefs.txtAlarm2Time.Text = gblAlarm2Time
    widgetPrefs.txtAlarm3Time.Text = gblAlarm3Time
    widgetPrefs.txtAlarm4Time.Text = gblAlarm4Time
    widgetPrefs.txtAlarm5Time.Text = gblAlarm5Time
    
    ' configuration tab
   
    ' check whether the size has been previously altered via ctrl+mousewheel on the widget
    sliGaugeSizeOldValue = sliGaugeSize.Value
    sliGaugeSize.Value = Val(gblGaugeSize)
    If sliGaugeSize.Value <> sliGaugeSizeOldValue Then
        btnSave.Visible = True
    End If
    
    cmbScrollWheelDirection.ListIndex = Val(gblScrollWheelDirection)
    chkEnableTooltips.Value = Val(gblEnableTooltips)
    chkEnableBalloonTooltips.Value = Val(gblEnableBalloonTooltips)
    chkShowTaskbar.Value = Val(gblShowTaskbar)
    chkShowHelp.Value = Val(gblShowHelp)
    chkTogglePendulum.Value = Val(gblTogglePendulum)
    
    chkDpiAwareness.Value = Val(gblDpiAwareness)
    
    chkEnablePrefsTooltips.Value = Val(gblEnablePrefsTooltips)
    
    ' sounds tab
    chkEnableSounds.Value = Val(gblEnableSounds)
    chkEnableTicks.Value = Val(gblEnableTicks)

    ' development
    cmbDebug.ListIndex = Val(gblDebug)
    txtDblClickCommand.Text = gblDblClickCommand
    txtOpenFile.Text = gblOpenFile
    txtDefaultEditor.Text = gblDefaultEditor
    lblGitHub.Caption = "You can find the code for the Steampunk Clock Calendar on github, visit by double-clicking this link https://github.com/yereverluvinunclebert/ Steampunk-Clock-Calendar"
     
     ' fonts tab
    If gblPrefsFont <> vbNullString Then
        txtPrefsFont.Text = gblPrefsFont
        If gblDpiAwareness = "1" Then
            Call changeFormFont(widgetPrefs, gblPrefsFont, Val(gblPrefsFontSizeHighDPI), fntWeight, fntStyle, gblPrefsFontItalics, gblPrefsFontColour)
            txtPrefsFontSize.Text = gblPrefsFontSizeHighDPI
        Else
            Call changeFormFont(widgetPrefs, gblPrefsFont, Val(gblPrefsFontSizeLowDPI), fntWeight, fntStyle, gblPrefsFontItalics, gblPrefsFontColour)
            txtPrefsFontSize.Text = gblPrefsFontSizeLowDPI
        End If
    End If
    
    
    ' position tab
    cmbAspectHidden.ListIndex = Val(gblAspectHidden)
    cmbWidgetPosition.ListIndex = Val(gblWidgetPosition)
        
    If gblPreventDragging = "1" Then
        If aspectRatio = "landscape" Then
'            txtLandscapeHoffset.Text = fClock.clockForm.Left
'            txtLandscapeVoffset.Text = fClock.clockForm.Top
            If gblDpiAwareness = "1" Then
                txtLandscapeHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & gblClockHighDpiXPos & "px"
                txtLandscapeVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & gblClockHighDpiYPos & "px"
            Else
                txtLandscapeHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & gblClockLowDpiXPos & "px"
                txtLandscapeVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & gblClockLowDpiYPos & "px"
            End If
        Else
'            txtPortraitHoffset.Text = fClock.clockForm.Left
'            txtPortraitYoffset.Text = fClock.clockForm.Top
            If gblDpiAwareness = "1" Then
                txtPortraitHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & gblClockHighDpiXPos & "px"
                txtPortraitYoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & gblClockHighDpiYPos & "px"
            Else
                txtPortraitHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & gblClockLowDpiXPos & "px"
                txtPortraitYoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & gblClockLowDpiYPos & "px"
            End If
        End If
    End If
    
    'cmbWidgetLandscape
    cmbWidgetLandscape.ListIndex = Val(gblWidgetLandscape)
    cmbWidgetPortrait.ListIndex = Val(gblWidgetPortrait)
    txtLandscapeHoffset.Text = gblLandscapeFormHoffset
    txtLandscapeVoffset.Text = gblLandscapeFormVoffset
    txtPortraitHoffset.Text = gblPortraitHoffset
    txtPortraitYoffset.Text = gblPortraitYoffset

    ' Windows tab
    cmbWindowLevel.ListIndex = Val(gblWindowLevel)
    chkIgnoreMouse.Value = Val(gblIgnoreMouse)
    chkPreventDragging.Value = Val(gblPreventDragging)
    sliOpacity.Value = Val(gblOpacity)
    chkWidgetHidden.Value = Val(gblWidgetHidden)
    cmbHidingTime.ListIndex = Val(gblHidingTime)
        
   On Error GoTo 0
   Exit Sub

adjustPrefsControls_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure adjustPrefsControls of Form widgetPrefs on line " & Erl

End Sub
'
'---------------------------------------------------------------------------------------
' Procedure : populatePrefsComboBoxes
' Author    : beededea
' Date      : 10/09/2022
' Purpose   : all combo boxes in the prefs are populated here with default values
'           : done by preference here rather than in the IDE
'---------------------------------------------------------------------------------------

Private Sub populatePrefsComboBoxes()
    
    On Error GoTo populatePrefsComboBoxes_Error
    
    cmbScrollWheelDirection.AddItem "up", 0
    cmbScrollWheelDirection.ItemData(0) = 0
    cmbScrollWheelDirection.AddItem "down", 1
    cmbScrollWheelDirection.ItemData(1) = 1
    
    cmbAspectHidden.AddItem "none", 0
    cmbAspectHidden.ItemData(0) = 0
    cmbAspectHidden.AddItem "portrait", 1
    cmbAspectHidden.ItemData(1) = 1
    cmbAspectHidden.AddItem "landscape", 2
    cmbAspectHidden.ItemData(2) = 2

    cmbWidgetPosition.AddItem "disabled", 0
    cmbWidgetPosition.ItemData(0) = 0
    cmbWidgetPosition.AddItem "enabled", 1
    cmbWidgetPosition.ItemData(1) = 1
    
    cmbWidgetLandscape.AddItem "disabled", 0
    cmbWidgetLandscape.ItemData(0) = 0
    cmbWidgetLandscape.AddItem "enabled", 1
    cmbWidgetLandscape.ItemData(1) = 1
    
    cmbWidgetPortrait.AddItem "disabled", 0
    cmbWidgetPortrait.ItemData(0) = 0
    cmbWidgetPortrait.AddItem "enabled", 1
    cmbWidgetPortrait.ItemData(1) = 1
    
    cmbDebug.AddItem "Debug OFF", 0
    cmbDebug.ItemData(0) = 0
    cmbDebug.AddItem "Debug ON", 1
    cmbDebug.ItemData(1) = 1
    
    ' populate comboboxes in the windows tab
    cmbWindowLevel.AddItem "Keep on top of other windows", 0
    cmbWindowLevel.ItemData(0) = 0
    cmbWindowLevel.AddItem "Normal", 0
    cmbWindowLevel.ItemData(1) = 1
    cmbWindowLevel.AddItem "Keep below all other windows", 0
    cmbWindowLevel.ItemData(2) = 2

    ' populate the hiding timer combobox
    cmbHidingTime.AddItem "1 minute", 0
    cmbHidingTime.ItemData(0) = 1
    cmbHidingTime.AddItem "5 minutes", 1
    cmbHidingTime.ItemData(1) = 5
    cmbHidingTime.AddItem "10 minutes", 2
    cmbHidingTime.ItemData(2) = 10
    cmbHidingTime.AddItem "20 minutes", 3
    cmbHidingTime.ItemData(3) = 20
    cmbHidingTime.AddItem "30 minutes", 4
    cmbHidingTime.ItemData(4) = 30
    cmbHidingTime.AddItem "I hour", 5
    cmbHidingTime.ItemData(5) = 60
    
    On Error GoTo 0
    Exit Sub

populatePrefsComboBoxes_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure populatePrefsComboBoxes of Form widgetPrefs"
            Resume Next
          End If
    End With
                
End Sub

'---------------------------------------------------------------------------------------
' Procedure : readFileWriteComboBox
' Author    : beededea
' Date      : 28/07/2023
' Purpose   : Open and load the Array with the timezones text File
'---------------------------------------------------------------------------------------
'
Private Sub readFileWriteComboBox(ByRef thisComboBox As Control, ByVal thisFileName As String)
    Dim strArr() As String
    Dim lngCount As Long: lngCount = 0
    Dim lngIdx As Long: lngIdx = 0
        
    On Error GoTo readFileWriteComboBox_Error

    If fFExists(thisFileName) = True Then
       ' the files must be DOS CRLF delineated
       Open thisFileName For Input As #1
           strArr() = Split(Input(LOF(1), 1), vbCrLf)
       Close #1
    
       lngCount = UBound(strArr)
    
       thisComboBox.Clear
       For lngIdx = 0 To lngCount
           thisComboBox.AddItem strArr(lngIdx)
       Next lngIdx
    End If

   On Error GoTo 0
   Exit Sub

readFileWriteComboBox_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readFileWriteComboBox of Form widgetPrefs"

End Sub

'
'---------------------------------------------------------------------------------------
' Procedure : clearBorderStyle
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : removes all styling from the icon frames and makes the major frames below invisible too, not using control arrays.
'---------------------------------------------------------------------------------------
'
Private Sub clearBorderStyle()

   On Error GoTo clearBorderStyle_Error

    fraGeneral.Visible = False
    fraConfig.Visible = False
    fraFonts.Visible = False
    fraWindow.Visible = False
    fraPosition.Visible = False
    fraDevelopment.Visible = False
    fraSounds.Visible = False
    fraAbout.Visible = False

    fraGeneralButton.BorderStyle = 0
    fraConfigButton.BorderStyle = 0
    fraDevelopmentButton.BorderStyle = 0
    fraPositionButton.BorderStyle = 0
    fraFontsButton.BorderStyle = 0
    fraWindowButton.BorderStyle = 0
    fraSoundsButton.BorderStyle = 0
    fraAboutButton.BorderStyle = 0

   On Error GoTo 0
   Exit Sub

clearBorderStyle_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure clearBorderStyle of Form widgetPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Resize
' Author    : beededea
' Date      : 30/05/2023
' Purpose   : If the form is NOT to be resized then restrain the height/width. Otherwise,
'             maintain the aspect ratio. When minimised and a resize is called then simply exit.
'---------------------------------------------------------------------------------------
'
Private Sub Form_Resize()
    Dim ratio As Double: ratio = 0
    Dim currentFont As Long: currentFont = 0
    
    On Error GoTo Form_Resize_Error
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    ' move the drag corner label along with the form's bottom right corner
    lblDragCorner.Move Me.ScaleLeft + Me.ScaleWidth - (lblDragCorner.Width + 40), _
               Me.ScaleTop + Me.ScaleHeight - (lblDragCorner.Height + 40)

    ratio = cPrefsFormHeight / cPrefsFormWidth
    
    If prefsDynamicSizingFlg = True Then
        
        If gblDpiAwareness = "1" Then
            currentFont = gblPrefsFontSizeHighDPI
        Else
            currentFont = gblPrefsFontSizeLowDPI
        End If
        Call resizeControls(Me, prefsControlPositions(), prefsCurrentWidth, prefsCurrentHeight, currentFont)
        Call tweakPrefsControlPositions(Me, prefsCurrentWidth, prefsCurrentHeight)
        
        Me.Width = Me.Height / ratio ' maintain the aspect ratio

        Call loadHigherResPrefsImages
    Else
        If Me.WindowState = 0 Then ' normal
            If Me.Width > 9090 Then Me.Width = 9090
            If Me.Width < 9085 Then Me.Width = 9090
            If lastFormHeight <> 0 Then Me.Height = lastFormHeight
        End If
    End If
        
    'lblSize.Caption = "topIconWidth = " & topIconWidth & " imgGeneral width = " & imgGeneral.Width
    
    On Error GoTo 0
    Exit Sub

Form_Resize_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Resize of Form widgetPrefs"
            Resume Next
          End If
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : tweakPrefsControlPositions
' Author    : beededea
' Date      : 22/09/2023
' Purpose   : final tweak the bottom frame top and left positions
'---------------------------------------------------------------------------------------
'
Private Sub tweakPrefsControlPositions(ByVal thisForm As Form, ByVal m_FormWid As Single, ByVal m_FormHgt As Single)

    ' not sure why but the resizeControls routine can lead to incorrect positioning of frames and buttons
    Dim x_scale As Single: x_scale = 0
    Dim y_scale As Single: y_scale = 0
    
    On Error GoTo tweakPrefsControlPositions_Error

    ' Get the form's current scale factors.
    x_scale = thisForm.ScaleWidth / m_FormWid
    y_scale = thisForm.ScaleHeight / m_FormHgt

    fraGeneral.Left = fraGeneralButton.Left
    fraConfig.Left = fraGeneralButton.Left
    fraSounds.Left = fraGeneralButton.Left
    fraPosition.Left = fraGeneralButton.Left
    fraFonts.Left = fraGeneralButton.Left
    fraDevelopment.Left = fraGeneralButton.Left
    fraWindow.Left = fraGeneralButton.Left
    fraAbout.Left = fraGeneralButton.Left
         
    'fraGeneral.Top = fraGeneralButton.Top
    fraConfig.Top = fraGeneral.Top
    fraSounds.Top = fraGeneral.Top
    fraPosition.Top = fraGeneral.Top
    fraFonts.Top = fraGeneral.Top
    fraDevelopment.Top = fraGeneral.Top
    fraWindow.Top = fraGeneral.Top
    fraAbout.Top = fraGeneral.Top
    
    ' final tweak the bottom button positions
    
    btnHelp.Top = fraGeneral.Top + fraGeneral.Height + (200 * y_scale)
    btnSave.Top = btnHelp.Top
    btnClose.Top = btnHelp.Top
    
    btnClose.Left = fraWindow.Left + fraWindow.Width - btnClose.Width
    btnSave.Left = btnClose.Left - btnSave.Width - (150 * x_scale)
    btnHelp.Left = fraGeneral.Left

    txtPrefsFontCurrentSize.Text = y_scale * txtPrefsFontCurrentSize.FontSize
    
   On Error GoTo 0
   Exit Sub

tweakPrefsControlPositions_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tweakPrefsControlPositions of Form widgetPrefs"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : Form_Unload
' Author    : beededea
' Date      : 27/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Unload(Cancel As Integer)
   On Error GoTo Form_Unload_Error

    gblPrefsLoadedFlg = False
    
    Call writePrefsPosition
    
    Call DestroyToolTip

   On Error GoTo 0
   Exit Sub

Form_Unload_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form widgetPrefs"
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    fraScrollbarCover.Visible = True

End Sub
Private Sub fraAbout_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraAbout_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    fraScrollbarCover.Visible = True
    If gblEnableBalloonTooltips = "1" Then CreateToolTip fraAbout.hwnd, "The About tab tells you all about this program and its creation using VB6.", _
                  TTIconInfo, "Help on the About Tab", , , , True
End Sub
Private Sub fraConfigInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraConfigInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnableBalloonTooltips = "1" Then CreateToolTip fraConfigInner.hwnd, "The configuration panel is the location for optional configuration items. These items change how the volume control operates, configure them to suit your needs and your mode of operation.", _
                  TTIconInfo, "Help on Configuration", , , , True

End Sub
Private Sub fraConfig_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraConfig_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnableBalloonTooltips = "1" Then CreateToolTip fraConfig.hwnd, "The configuration panel is the location for optional configuration items. These items change how the volume control operates, configure them to suit your needs and your mode of operation.", _
                  TTIconInfo, "Help on Configuration", , , , True

End Sub

Private Sub fraDefaultEditor_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblGitHub.ForeColor = &H80000012
End Sub

Private Sub fraDevelopment_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraDevelopment_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnableBalloonTooltips = "1" Then CreateToolTip fraDevelopment.hwnd, "This tab contains elements that will assist in debugging and developing this program further. ", _
                  TTIconInfo, "Help on the Development Tab", , , , True
End Sub


Private Sub fraDevelopmentInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraDevelopmentInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnableBalloonTooltips = "1" Then CreateToolTip fraDevelopmentInner.hwnd, "This tab contains elements that will assist in debugging and developing this program further. ", _
                  TTIconInfo, "Help on the Development Tab", , , , True

End Sub
Private Sub fraFonts_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnableBalloonTooltips = "1" Then CreateToolTip fraFonts.hwnd, "This tab allows you to set a specific font for the preferences only as there are no textual elements in the main program. We suggest Centurion Light SF at 8pt, which you will find bundled in the gbl program folder. Choose a small 8pt font for each.", _
                  TTIconInfo, "Help on Setting the Fonts", , , , True

End Sub
Private Sub fraFontsInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraFontsInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnableBalloonTooltips = "1" Then CreateToolTip fraFontsInner.hwnd, "This tab allows you to set a specific font for the preferences only as there are no textual elements in the main program. We suggest Centurion Light SF at 8pt, which you will find bundled in the gbl program folder. Choose a small 8pt font for each.", _
                  TTIconInfo, "Help on Setting the Fonts", , , , True
End Sub
'Private Sub fraConfigurationButtonInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    If Button = 2 Then
'        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
'    End If
'End Sub
Private Sub fraGeneral_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraGeneral_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnableBalloonTooltips = "1" Then CreateToolTip fraGeneral.hwnd, "The General Panel contains the most important user-configurable items required for the program to operate correctly.", _
                  TTIconInfo, "Help on Essential Configuration", , , , True
End Sub

Private Sub fraGeneralInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraGeneralInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnableBalloonTooltips = "1" Then CreateToolTip fraGeneralInner.hwnd, "The General Panel contains the most important user-configurable items required for the program to operate correctly.", _
                  TTIconInfo, "Help on Essential Configuration", , , , True
End Sub

Private Sub fraPosition_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     If gblEnableBalloonTooltips = "1" Then CreateToolTip fraPosition.hwnd, "This tab allows you to determine the X and Y positioning of your widget in landscape and portrait screen modes. Best left well alone unless you use Windows on a tablet.", _
                  TTIconInfo, "Help on Tablet Positioning", , , , True
End Sub
Private Sub fraPositionInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraPositionInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnableBalloonTooltips = "1" Then CreateToolTip fraPositionInner.hwnd, "This tab allows you to determine the X and Y positioning of your widget in landscape and portrait screen modes. Best left well alone unless you use Windows on a tablet.", _
                  TTIconInfo, "Help on Tablet Positioning", , , , True
End Sub

Private Sub fraScrollbarCover_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    fraScrollbarCover.Visible = False

End Sub
Private Sub fraSounds_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraSounds_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 If gblEnableBalloonTooltips = "1" Then CreateToolTip fraSounds.hwnd, "The sound panel allows you to configure the sounds that occur within gbl. Some of the animations have associated sounds, you can control these here..", _
                  TTIconInfo, "Help on Configuring Sounds", , , , True
End Sub
Private Sub fraSoundsInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraSoundsInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     If gblEnableBalloonTooltips = "1" Then CreateToolTip fraSoundsInner.hwnd, "The sound panel allows you to configure the sounds that occur within gbl. Some of the animations have associated sounds, you can control these here..", _
                  TTIconInfo, "Help on Configuring Sounds", , , , True
End Sub

Private Sub fraWindow_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraWindow_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     If gblEnableBalloonTooltips = "1" Then CreateToolTip fraWindow.hwnd, "The Opacity and Window Level of the program are rather strange characteristics to change in a Windows program, however this widget is a copy of a Yahoo Widget of the same name. All widgets have similar window tab options including the capability to change the opacity and window level. Whether these options are useful to you or anyone is a moot point but as this tool aims to replicate the YWE version functionality it has been reproduced here. It is here as more of an experiment as to how to implement a feature, one carried over from the Yahoo Widget (javascript) version of this program.", _
                  TTIconInfo, "Help on YWE Quirk Mode Options", , , , True
End Sub
Private Sub fraWindowInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraWindowInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     If gblEnableBalloonTooltips = "1" Then CreateToolTip fraWindowInner.hwnd, "The Opacity and Window Level of the program are rather strange characteristics to change in a Windows program, however this widget is a copy of a Yahoo Widget of the same name. All widgets have similar window tab options including the capability to change the opacity and window level. Whether these options are useful to you or anyone is a moot point but as this tool aims to replicate the YWE version functionality it has been reproduced here. It is here as more of an experiment as to how to implement a feature, one carried over from the Yahoo Widget (javascript) version of this program.", _
                  TTIconInfo, "Help on YWE Quirk Mode Options", , , , True
End Sub

Private Sub imgGeneral_Click()
    imgGeneral.Visible = False
    imgGeneralClicked.Visible = True
End Sub

Private Sub imgGeneral_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picButtonMouseUpEvent("general", imgGeneral, imgGeneralClicked, fraGeneral, fraGeneralButton) ' was imgGeneralMouseUpEvent
End Sub



'---------------------------------------------------------------------------------------
' Procedure : lblGitHub_dblClick
' Author    : beededea
' Date      : 14/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblGitHub_dblClick()
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString
    
    On Error GoTo lblGitHub_dblClick_Error

    answer = vbYes
    answerMsg = "This option opens a browser window and take you straight to Github. Proceed?"
    answer = msgBoxA(answerMsg, vbExclamation + vbYesNo, "Proceed to Github? ", True, "lblGitHubDblClick")
    If answer = vbYes Then
       Call ShellExecute(Me.hwnd, "Open", "https://github.com/yereverluvinunclebert/Steampunk-Clock-Calendar-VB6", vbNullString, App.path, 1)
    End If

   On Error GoTo 0
   Exit Sub

lblGitHub_dblClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblGitHub_dblClick of Form widgetPrefs"
End Sub

Private Sub lblGitHub_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblGitHub.ForeColor = &H8000000D
End Sub



Private Sub sliOpacity_Change()
    btnSave.Enabled = True ' enable the save button
End Sub

'---------------------------------------------------------------------------------------
' Procedure : txtAboutText_MouseDown
' Author    : beededea
' Date      : 30/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtAboutText_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo txtAboutText_MouseDown_Error

    If Button = vbRightButton Then
        txtAboutText.Enabled = False
        txtAboutText.Enabled = True
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If

    On Error GoTo 0
    Exit Sub

txtAboutText_MouseDown_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure txtAboutText_MouseDown of Form widgetPrefs"
End Sub

Private Sub txtAboutText_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    fraScrollbarCover.Visible = False
End Sub

Private Sub imgAbout_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgAbout.Visible = False
    imgAboutClicked.Visible = True
End Sub
Private Sub imgAbout_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picButtonMouseUpEvent("about", imgAbout, imgAboutClicked, fraAbout, fraAboutButton)
End Sub

Private Sub imgDevelopment_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgDevelopment.Visible = False
    imgDevelopmentClicked.Visible = True
End Sub

Private Sub imgDevelopment_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picButtonMouseUpEvent("development", imgDevelopment, imgDevelopmentClicked, fraDevelopment, fraDevelopmentButton)
End Sub

Private Sub imgFonts_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgFonts.Visible = False
    imgFontsClicked.Visible = True
End Sub

Private Sub imgFonts_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picButtonMouseUpEvent("fonts", imgFonts, imgFontsClicked, fraFonts, fraFontsButton)
End Sub

Private Sub imgConfig_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgConfig.Visible = False
    imgConfigClicked.Visible = True
End Sub

Private Sub imgConfig_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picButtonMouseUpEvent("config", imgConfig, imgConfigClicked, fraConfig, fraConfigButton) ' was imgConfigMouseUpEvent
End Sub

Private Sub imgGeneral_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgGeneral.Visible = False
    imgGeneralClicked.Visible = True
End Sub


Private Sub imgPosition_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgPosition.Visible = False
    imgPositionClicked.Visible = True
End Sub

Private Sub imgPosition_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picButtonMouseUpEvent("position", imgPosition, imgPositionClicked, fraPosition, fraPositionButton)
End Sub

Private Sub imgSounds_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    '
    imgSounds.Visible = False
    imgSoundsClicked.Visible = True
End Sub

Private Sub imgSounds_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Call imgSoundsMouseUpEvent
    Call picButtonMouseUpEvent("sounds", imgSounds, imgSoundsClicked, fraSounds, fraSoundsButton)
End Sub

Private Sub imgWindow_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgWindow.Visible = False
    imgWindowClicked.Visible = True
End Sub

Private Sub imgWindow_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picButtonMouseUpEvent("window", imgWindow, imgWindowClicked, fraWindow, fraWindowButton)
End Sub

'Private Sub sliAnimationInterval_Change()
'    'overlayWidget.RotationSpeed = sliAnimationInterval.Value
'    btnSave.Enabled = True ' enable the save button
'
'End Sub

'Private Sub 'sliWidgetSkew_Click()
'    btnSave.Enabled = True ' enable the save button
'    'overlayWidget.GlobeSkewDeg = 'sliWidgetSkew.Value
'End Sub


Private Sub sliGaugeSize_GotFocus()
    gblAllowSizeChangeFlg = True
End Sub

Private Sub sliGaugeSize_LostFocus()
    gblAllowSizeChangeFlg = False
End Sub
'---------------------------------------------------------------------------------------
' Procedure : sliGaugeSize_Change
' Author    : beededea
' Date      : 30/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliGaugeSize_Change()
    On Error GoTo sliGaugeSize_Change_Error

    btnSave.Enabled = True ' enable the save button
    
    If gblAllowSizeChangeFlg = True Then Call fClock.AdjustZoom(sliGaugeSize.Value / 100)

    On Error GoTo 0
    Exit Sub

sliGaugeSize_Change_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliGaugeSize_Change of Form widgetPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliOpacity_Change
' Author    : beededea
' Date      : 15/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliOpacity_Click()
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString
    
    On Error GoTo sliOpacity_Change_Error

    btnSave.Enabled = True ' enable the save button

    If prefsStartupFlg = False Then
        gblOpacity = LTrim$(Str$(sliOpacity.Value))
    
        sPutINISetting "Software\SteampunkClockCalendar", "opacity", gblOpacity, gblSettingsFile
        
        'Call setOpacity(sliOpacity.Value) ' this works but reveals the background form itself
        
        answer = vbYes
        answerMsg = "You must perform a hard reload on this widget in order to change the widget's opacity, do you want me to do it for you now?"
        answer = msgBoxA(answerMsg, vbYesNo, "Hard Reload Request", True, "sliOpacityClick")
        If answer = vbNo Then
            Exit Sub
        Else
            Call hardRestart
        End If
    End If

   On Error GoTo 0
   Exit Sub

sliOpacity_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliOpacity_Change of Form widgetPrefs"
End Sub




'---------------------------------------------------------------------------------------
' Procedure : Form_MouseDown
' Author    : beededea
' Date      : 14/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo Form_MouseDown_Error

    If Button = 2 Then

        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
        
    End If

   On Error GoTo 0
   Exit Sub

Form_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_MouseDown of Form widgetPrefs"
End Sub


Private Sub fraFonts_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub


Private Sub txtAlarm1Date_Change()
    btnSave.Enabled = True ' enable the save button
End Sub

Private Sub txtAlarm1Date_Click()
    If txtAlarm1Date.Text = "Alarm not yet set" Then txtAlarm1Date.Text = vbNullString

End Sub

Private Sub txtAlarm1Time_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtAlarm2Date_Change()
    btnSave.Enabled = True ' enable the save button
End Sub

Private Sub txtAlarm2Date_Click()
    If txtAlarm2Date.Text = "Alarm not yet set" Then txtAlarm2Date.Text = vbNullString

End Sub

Private Sub txtAlarm2Time_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtAlarm3Date_Change()
    btnSave.Enabled = True ' enable the save button
End Sub

Private Sub txtAlarm3Date_Click()
    If txtAlarm3Date.Text = "Alarm not yet set" Then txtAlarm3Date.Text = vbNullString

End Sub

Private Sub txtAlarm3Time_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtAlarm4Date_Change()
    btnSave.Enabled = True ' enable the save button
End Sub



Private Sub txtAlarm4Date_Click()
    If txtAlarm4Date.Text = "Alarm not yet set" Then txtAlarm4Date.Text = vbNullString

End Sub

Private Sub txtAlarm4Time_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtAlarm5Date_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtAlarm5Date_Click()
    If txtAlarm5Date.Text = "Alarm not yet set" Then txtAlarm5Date.Text = vbNullString
End Sub

Private Sub txtAlarm5Time_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtDblClickCommand_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtDefaultEditor_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtLandscapeHoffset_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtLandscapeVoffset_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtOpenFile_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtPortraitHoffset_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtPortraitYoffset_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtPrefsFont_Change()
    btnSave.Enabled = True ' enable the save button
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuAbout_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAbout_Click()
    
    On Error GoTo mnuAbout_Click_Error

    Call aboutClickEvent

    On Error GoTo 0
    Exit Sub

mnuAbout_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAbout_Click of form menuForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : setPrefsTooltips
' Author    : beededea
' Date      : 27/04/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub setPrefsTooltips()

   On Error GoTo setPrefsTooltips_Error

    If chkEnablePrefsTooltips.Value = 1 Then
        imgConfig.ToolTipText = "Opens the configuration tab"
        imgConfigClicked.ToolTipText = "Opens the configuration tab"
        imgDevelopment.ToolTipText = "Opens the Development tab"
        imgDevelopmentClicked.ToolTipText = "Opens the Development tab"
        imgPosition.ToolTipText = "Opens the Position tab"
        imgPositionClicked.ToolTipText = "Opens the Position tab"
        btnSave.ToolTipText = "Save the changes you have made to the preferences"
        btnHelp.ToolTipText = "Open the help utility"
        imgSounds.ToolTipText = "Opens the Sounds tab"
        imgSoundsClicked.ToolTipText = "Opens the Sounds tab"
        btnClose.ToolTipText = "Close the utility"
        imgWindow.ToolTipText = "Opens the Window tab"
        imgWindowClicked.ToolTipText = "Opens the Window tab"
        lblWindow.ToolTipText = "Opens the Window tab"
        imgFonts.ToolTipText = "Opens the Fonts tab"
        imgFonts.ToolTipText = "Opens the Fonts tab"
        imgFontsClicked.ToolTipText = "Opens the Fonts tab"
        imgGeneral.ToolTipText = "Opens the general tab"
        imgGeneralClicked.ToolTipText = "Opens the general tab"
        lblPosition(6).ToolTipText = "Tablets only. Don't fiddle with this unless you really know what you are doing. Here you can choose whether this the volume control widget is hidden by default in either landscape or portrait mode or not at all. This option allows you to have certain widgets that do not obscure the screen in either landscape or portrait. If you accidentally set it so you can't find your widget on screen then change the setting here to NONE."
        chkGenStartup.ToolTipText = "Check this box to enable the automatic start of the program when Windows is started."
        chkWidgetFunctions.ToolTipText = "When checked this box enables the spinning earth functionality. Any adjustment takes place instantly. "

        txtPortraitYoffset.ToolTipText = "Field to hold the vertical offset for the widget position in portrait mode."
        txtPortraitHoffset.ToolTipText = "Field to hold the horizontal offset for the widget position in portrait mode."
        txtLandscapeVoffset.ToolTipText = "Field to hold the horizontal offset for the widget position in landscape mode."
        txtLandscapeHoffset.ToolTipText = "Field to hold the horizontal offset for the widget position in landscape mode."
        cmbWidgetLandscape.ToolTipText = "The widget can be locked into landscape mode, it ensures that the widget always appears where you want it to. Using the fields below, you can assign a default x/y position for Landscape mode. "
        cmbWidgetPortrait.ToolTipText = "The widget can be locked into portrait mode, it ensures that the widget always appears where you want it to. Using the fields below, you can assign a default x/y position for portrait mode. "
        cmbWidgetPosition.ToolTipText = "Tablets only. The widget can be positioned proportionally when switching between portrait/landscape. If you want to enable this, disable the options below."
        cmbAspectHidden.ToolTipText = " Here you can choose whether the widget is hidden by default in either landscape or portrait mode or not at all. This allows you to have certain widgets that do not obscure the screen in one mode or another. If you accidentally set it so you can't find it on screen then change the setting here to none."
        chkEnableSounds.ToolTipText = "Check this box to enable or disable all of the sounds used during any animation on the main steampunk GUI as well as all other chimes, tick sounds."
        chkEnableTicks.ToolTipText = "Enables or disables just the sound of the clock ticking."
        chkEnableChimes.ToolTipText = "Enables or disables just the clock chimes."
        chkEnableBoost.ToolTipText = "Sets the volume of the various sound elements, you can boost from quiet to loud."
        btnDefaultEditor.ToolTipText = "Click to select the .vbp file to edit the program - You need to have access to the source!"
        txtDblClickCommand.ToolTipText = "Enter a Windows command for the gauge to operate when double-clicked."
        btnOpenFile.ToolTipText = "Click to select a particular file for the gauge to run or open when double-clicked."
        txtOpenFile.ToolTipText = "Enter a particular file for the gauge to run or open when double-clicked."
        cmbDebug.ToolTipText = "Choose to set debug mode."
        txtPrefsFontSize.ToolTipText = "Disabled for manual input. Choose a font size using the font selector to be used within this preferences window only"
        btnPrefsFont.ToolTipText = "The Font Selector."
        txtPrefsFont.ToolTipText = "Disabled for manual input. Choose a font via the font selector to be used only for this preferences window"
        txtPrefsFontSize.ToolTipText = "Disabled for manual input. Choose a font size via the font selector that fits the text boxes"
        cmbWindowLevel.ToolTipText = "You can determine the window position here. Set to bottom to keep the widget below other windows."
        cmbHidingTime.ToolTipText = "."
        chkEnableResizing.ToolTipText = "Provides an alternative method of supporting high DPI screens."
        chkPreventDragging.ToolTipText = "Checking this box turns off the ability to drag the program with the mouse. The locking in position effect takes place instantly."
        chkIgnoreMouse.ToolTipText = "Checking this box causes the program to ignore all mouse events."
        sliOpacity.ToolTipText = "Set the transparency of the program. Any change in opacity takes place instantly."
        cmbScrollWheelDirection.ToolTipText = "To change the direction of the mouse scroll wheel when resiziing the globe gauge."
        chkEnableBalloonTooltips.ToolTipText = "Check the box to enable larger balloon tooltips for all controls on the main program"
        chkShowTaskbar.ToolTipText = "Check the box to show the widget in the taskbar"
        chkShowHelp.ToolTipText = "Check the box to show the help page on startup"
        chkEnableTooltips.ToolTipText = "Check the box to enable tooltips for all controls on the main program"
        sliGaugeSize.ToolTipText = "Adjust to a percentage of the original size. Any adjustment in size takes place instantly (you can also use Ctrl+Mousewheel hovering over the globe itself)."
        btnFacebook.ToolTipText = "This will link you to the our Steampunk/Dieselpunk program users Group."
        imgAbout.ToolTipText = "Opens the About tab"
        btnAboutDebugInfo.ToolTipText = "This gives access to the debugging tool"
        btnDonate.ToolTipText = "Buy me a Kofi! This button opens a browser window and connects to Kofi donation page"
        btnUpdate.ToolTipText = "Here you can visit the update location where you can download new versions of the programs."
        lblFontsTab(0).ToolTipText = "We suggest Centurion Light SF at 8pt - which you will find in the FCW program folder"
        lblFontsTab(6).ToolTipText = "We suggest Centurion Light SF at 8pt - which you will find in the FCW program folder"
        lblFontsTab(7).ToolTipText = "Choose a font size that fits the text boxes"
        txtPrefsFontCurrentSize.ToolTipText = "Disabled for manual input. Shows the current font size when form resizing is enabled."
        lblCurrentFontsTab.ToolTipText = "Disabled for manual input. Shows the current font size when form resizing is enabled."
        
        chkDpiAwareness.ToolTipText = " Check the box to make the program DPI aware. RESTART required."
        chkEnablePrefsTooltips.ToolTipText = "Check the box to enable tooltips for all controls in the preferences utility"
        btnResetMessages.ToolTipText = "This button restores the pop-up messages to their original visible state."
        
        txtAlarm1Date.ToolTipText = "This is alarm number one, set the date here using the alarm toggle and slider or by editing any values shown here."
        txtAlarm2Date.ToolTipText = "This is alarm number two, set the date here using the alarm toggle and slider or by editing any values shown here."
        txtAlarm3Date.ToolTipText = "This is alarm number three, set the date here using the alarm toggle and slider or by editing any values shown here."
        txtAlarm4Date.ToolTipText = "This is alarm number four, set the date here using the alarm toggle and slider or by editing any values shown here."
        txtAlarm5Date.ToolTipText = "This is alarm number five, set the date here using the alarm toggle and slider or by editing any values shown here."
        
        txtAlarm1Time.ToolTipText = "This is alarm number one, set the time here using the alarm toggle and slider or by editing any values shown here."
        txtAlarm2Time.ToolTipText = "This is alarm number two, set the time here using the alarm toggle and slider or by editing any values shown here."
        txtAlarm3Time.ToolTipText = "This is alarm number three, set the time here using the alarm toggle and slider or by editing any values shown here."
        txtAlarm4Time.ToolTipText = "This is alarm number four, set the time here using the alarm toggle and slider or by editing any values shown here."
        txtAlarm5Time.ToolTipText = "This is alarm number five, set the time here using the alarm toggle and slider or by editing any values shown here."
        
        btnVerifyDateTime1.ToolTipText = "Verify Date Time for alarm number 1"
        btnVerifyDateTime2.ToolTipText = "Verify Date Time for alarm number 2"
        btnVerifyDateTime3.ToolTipText = "Verify Date Time for alarm number 3"
        btnVerifyDateTime4.ToolTipText = "Verify Date Time for alarm number 4"
        btnVerifyDateTime5.ToolTipText = "Verify Date Time for alarm number 5"
        

    Else
        imgConfig.ToolTipText = vbNullString
        imgConfigClicked.ToolTipText = vbNullString
        imgDevelopment.ToolTipText = vbNullString
        imgDevelopmentClicked.ToolTipText = vbNullString
        imgPosition.ToolTipText = vbNullString
        imgPositionClicked.ToolTipText = vbNullString
        btnSave.ToolTipText = vbNullString
        btnHelp.ToolTipText = vbNullString
        imgSounds.ToolTipText = vbNullString
        imgSoundsClicked.ToolTipText = vbNullString
        btnClose.ToolTipText = vbNullString
        imgWindow.ToolTipText = vbNullString
        imgWindowClicked.ToolTipText = vbNullString
        imgFonts.ToolTipText = vbNullString
        imgFontsClicked.ToolTipText = vbNullString
        imgGeneral.ToolTipText = vbNullString
        imgGeneralClicked.ToolTipText = vbNullString
        chkGenStartup.ToolTipText = vbNullString
        chkWidgetFunctions.ToolTipText = vbNullString

        txtPortraitYoffset.ToolTipText = vbNullString
        txtPortraitHoffset.ToolTipText = vbNullString
        txtLandscapeVoffset.ToolTipText = vbNullString
        txtLandscapeHoffset.ToolTipText = vbNullString
        cmbWidgetLandscape.ToolTipText = vbNullString
        cmbWidgetPortrait.ToolTipText = vbNullString
        cmbWidgetPosition.ToolTipText = vbNullString
        cmbAspectHidden.ToolTipText = vbNullString
        chkEnableSounds.ToolTipText = vbNullString
        chkEnableTicks.ToolTipText = vbNullString
        chkEnableChimes.ToolTipText = vbNullString
        chkEnableBoost.ToolTipText = vbNullString
        
        btnDefaultEditor.ToolTipText = vbNullString
        txtDblClickCommand.ToolTipText = vbNullString
        btnOpenFile.ToolTipText = vbNullString
        txtOpenFile.ToolTipText = vbNullString
        cmbDebug.ToolTipText = vbNullString
        txtPrefsFontSize.ToolTipText = vbNullString
        btnPrefsFont.ToolTipText = vbNullString
        txtPrefsFont.ToolTipText = vbNullString
        cmbWindowLevel.ToolTipText = vbNullString
        cmbHidingTime.ToolTipText = vbNullString
        chkEnableResizing.ToolTipText = vbNullString
        chkPreventDragging.ToolTipText = vbNullString
        chkIgnoreMouse.ToolTipText = vbNullString
        sliOpacity.ToolTipText = vbNullString
        cmbScrollWheelDirection.ToolTipText = vbNullString
        chkEnableBalloonTooltips.ToolTipText = vbNullString
        chkShowTaskbar.ToolTipText = vbNullString
        chkShowHelp.ToolTipText = vbNullString
        chkEnableTooltips.ToolTipText = vbNullString
        sliGaugeSize.ToolTipText = vbNullString
        btnFacebook.ToolTipText = vbNullString
        imgAbout.ToolTipText = vbNullString
        btnAboutDebugInfo.ToolTipText = vbNullString
        btnDonate.ToolTipText = vbNullString
        btnUpdate.ToolTipText = vbNullString
        lblFontsTab(0).ToolTipText = vbNullString
        lblFontsTab(6).ToolTipText = vbNullString
        lblFontsTab(7).ToolTipText = vbNullString
        txtPrefsFontCurrentSize.ToolTipText = vbNullString
        lblCurrentFontsTab.ToolTipText = vbNullString
        
        chkDpiAwareness.ToolTipText = vbNullString
        chkEnablePrefsTooltips.ToolTipText = vbNullString
        btnResetMessages.ToolTipText = vbNullString
    
        txtAlarm1Date.ToolTipText = vbNullString
        txtAlarm2Date.ToolTipText = vbNullString
        txtAlarm3Date.ToolTipText = vbNullString
        txtAlarm4Date.ToolTipText = vbNullString
        txtAlarm5Date.ToolTipText = vbNullString
    
        btnVerifyDateTime1.ToolTipText = vbNullString
        btnVerifyDateTime2.ToolTipText = vbNullString
        btnVerifyDateTime3.ToolTipText = vbNullString
        btnVerifyDateTime4.ToolTipText = vbNullString
        btnVerifyDateTime5.ToolTipText = vbNullString
    End If

   On Error GoTo 0
   Exit Sub

setPrefsTooltips_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setPrefsTooltips of Form widgetPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : setPrefsLabels
' Author    : beededea
' Date      : 27/09/2023
' Purpose   : set the text in any labels that need a vbCrLf to space the text
'---------------------------------------------------------------------------------------
'
Private Sub setPrefsLabels()

    On Error GoTo setPrefsLabels_Error

    lblFontsTab(0).Caption = "When resizing the form (drag bottom right) the font size will in relation to form height. The base font determines the initial size, the resulting resized font will dynamically change." & vbCrLf & vbCrLf & _
        "Next time you open the prefs it will revert to the default." & vbCrLf & vbCrLf & _
        "My preferred font for this utility is Centurion Light SF at 8pt size."

    On Error GoTo 0
    Exit Sub

setPrefsLabels_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setPrefsLabels of Form widgetPrefs"
        
End Sub

'---------------------------------------------------------------------------------------
' Procedure : DestroyToolTip
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : It's not a bad idea to put this in the Form_Unload event just to make sure.
'---------------------------------------------------------------------------------------
'
Public Sub DestroyToolTip()
    
   On Error GoTo DestroyToolTip_Error

    If hwndTT <> 0& Then DestroyWindow hwndTT
    hwndTT = 0&

   On Error GoTo 0
   Exit Sub

DestroyToolTip_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DestroyToolTip of Form widgetPrefs"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : loadPrefsAboutText
' Author    : beededea
' Date      : 12/03/2020
' Purpose   : The text for the about page is stored here
'---------------------------------------------------------------------------------------
'
Private Sub loadPrefsAboutText()
    On Error GoTo loadPrefsAboutText_Error
    'If debugflg = 1 Then Debug.Print "%loadPrefsAboutText"
    
    lblMajorVersion.Caption = App.Major
    lblMinorVersion.Caption = App.Minor
    lblRevisionNum.Caption = App.Revision
    
    Call LoadFileToTB(txtAboutText, App.path & "\resources\txt\about.txt", False)

   On Error GoTo 0
   Exit Sub

loadPrefsAboutText_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadPrefsAboutText of Form widgetPrefs"
    
End Sub

'
'---------------------------------------------------------------------------------------
' Procedure : picButtonMouseUpEvent
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : capture the icon button clicks avoiding creating a control array
'---------------------------------------------------------------------------------------
'
Private Sub picButtonMouseUpEvent(ByVal thisTabName As String, ByRef thisPicName As Image, ByRef thisPicNameClicked As Image, ByRef thisFraName As Frame, Optional ByRef thisFraButtonName As Frame)
    
    On Error GoTo picButtonMouseUpEvent_Error
    
    Dim padding As Long: padding = 0
    Dim borderWidth As Long: borderWidth = 0
    Dim captionHeight As Long: captionHeight = 0
    Dim y_scale As Single: y_scale = 0
    
    thisPicNameClicked.Visible = False
    thisPicName.Visible = True
      
    btnSave.Visible = False
    btnClose.Visible = False
    btnHelp.Visible = False
    
    Call clearBorderStyle

    gblLastSelectedTab = thisTabName
    sPutINISetting "Software\SteampunkClockCalendar", "lastSelectedTab", gblLastSelectedTab, gblSettingsFile

    thisFraName.Visible = True
    thisFraButtonName.BorderStyle = 1

    ' Get the form's current scale factors.
    y_scale = Me.ScaleHeight / prefsCurrentHeight
    
    If gblDpiAwareness = "1" Then
        btnHelp.Top = fraGeneral.Top + fraGeneral.Height + (200 * y_scale)
    Else
        btnHelp.Top = thisFraName.Top + thisFraName.Height + (200 * y_scale)
    End If
    
    btnSave.Top = btnHelp.Top
    btnClose.Top = btnSave.Top
    
    btnSave.Visible = True
    btnClose.Visible = True
    btnHelp.Visible = True
    
    lblAsterix.Top = btnSave.Top + 50
    lblSize.Top = lblAsterix.Top - 300
    
    chkEnableResizing.Top = btnSave.Top + 50
    'chkEnableResizing.Left = lblAsterix.Left
    
    borderWidth = (Me.Width - Me.ScaleWidth) / 2
    captionHeight = Me.Height - Me.ScaleHeight - borderWidth
        
    ' under windows 10+ the internal window calcs are all wrong due to the bigger title bars
    If prefsDynamicSizingFlg = False Then
        padding = 200 ' add normal padding below the help button to position the bottom of the form

        lastFormHeight = btnHelp.Top + btnHelp.Height + captionHeight + borderWidth + padding
        Me.Height = lastFormHeight
    End If
    
    If gblDpiAwareness = "0" Then
        If thisTabName = "about" Then
            lblAsterix.Visible = False
            chkEnableResizing.Visible = True
        Else
            lblAsterix.Visible = True
            chkEnableResizing.Visible = False
        End If
    End If
    
   On Error GoTo 0
   Exit Sub

picButtonMouseUpEvent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picButtonMouseUpEvent of Form widgetPrefs"

End Sub




'---------------------------------------------------------------------------------------
' Procedure : themeTimer_Timer
' Author    : beededea
' Date      : 13/06/2020
' Purpose   : a timer to apply a theme automatically
'---------------------------------------------------------------------------------------
'
Private Sub themeTimer_Timer()
        
    Dim SysClr As Long: SysClr = 0

    On Error GoTo themeTimer_Timer_Error

    SysClr = GetSysColor(COLOR_BTNFACE)

    If SysClr <> storeThemeColour Then
        Call setThemeColour
    End If

   On Error GoTo 0
   Exit Sub

themeTimer_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure themeTimer_Timer of Form widgetPrefs"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : mnuCoffee_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   : main menu item to buy the developer a coffee
'---------------------------------------------------------------------------------------
'
Private Sub mnuCoffee_Click()
    On Error GoTo mnuCoffee_Click_Error
    
    Call mnuCoffee_ClickEvent

    On Error GoTo 0
    Exit Sub
mnuCoffee_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuCoffee_Click of Form widgetPrefs"
End Sub


'
'---------------------------------------------------------------------------------------
' Procedure : mnuLicenceA_Click
' Author    : beededea
' Date      : 17/08/2022
' Purpose   : menu option to show licence
'---------------------------------------------------------------------------------------
'
Private Sub mnuLicenceA_Click()
    On Error GoTo mnuLicenceA_Click_Error

    Call mnuLicence_ClickEvent

    On Error GoTo 0
    Exit Sub

mnuLicenceA_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuLicenceA_Click of Form widgetPrefs"
            Resume Next
          End If
    End With

End Sub



'---------------------------------------------------------------------------------------
' Procedure : mnuSupport_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   : menu option to open support page
'---------------------------------------------------------------------------------------
'
Private Sub mnuSupport_Click()
    
    On Error GoTo mnuSupport_Click_Error

    Call mnuSupport_ClickEvent

    On Error GoTo 0
    Exit Sub

mnuSupport_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuSupport_Click of Form widgetPrefs"
End Sub




Private Sub mnuClosePreferences_Click()
    Call btnClose_Click
End Sub


'---------------------------------------------------------------------------------------
' Procedure : mnuAuto_Click
' Author    : beededea
' Date      : 19/05/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAuto_Click()
    
   On Error GoTo mnuAuto_Click_Error

    If themeTimer.Enabled = True Then
            MsgBox "Automatic Theme Selection is now Disabled"
            mnuAuto.Caption = "Auto Theme Disabled - Click to Enable"
            mnuAuto.Checked = False
            
            themeTimer.Enabled = False
    Else
            MsgBox "Auto Theme Selection Enabled. If the o/s theme changes the utility should automatically skin the utility to suit the theme."
            mnuAuto.Caption = "Auto Theme Enabled - Click to Disable"
            mnuAuto.Checked = True
            
            themeTimer.Enabled = True
            Call setThemeColour
    End If

   On Error GoTo 0
   Exit Sub

mnuAuto_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAuto_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuDark_Click
' Author    : beededea
' Date      : 19/05/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuDark_Click()
   On Error GoTo mnuDark_Click_Error

    mnuAuto.Caption = "Auto Theme Disabled - Click to Enable"
    mnuAuto.Checked = False
    mnuDark.Caption = "Dark Theme Enabled"
    mnuLight.Caption = "Light Theme Enable"
    themeTimer.Enabled = False
    
    gblSkinTheme = "dark"

    Call setThemeShade(212, 208, 199)

   On Error GoTo 0
   Exit Sub

mnuDark_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuDark_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuLight_Click
' Author    : beededea
' Date      : 19/05/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuLight_Click()
    'MsgBox "Auto Theme Selection Manually Disabled"
   On Error GoTo mnuLight_Click_Error
    
    mnuAuto.Caption = "Auto Theme Disabled - Click to Enable"
    mnuAuto.Checked = False
    mnuDark.Caption = "Dark Theme Enable"
    mnuLight.Caption = "Light Theme Enabled"
    themeTimer.Enabled = False
    
    gblSkinTheme = "light"

    Call setThemeShade(240, 240, 240)

   On Error GoTo 0
   Exit Sub

mnuLight_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuLight_Click of Form widgetPrefs"
End Sub




'
'---------------------------------------------------------------------------------------
' Procedure : setThemeShade
' Author    : beededea
' Date      : 06/05/2023
' Purpose   : set the theme shade, Windows classic dark/new lighter theme colours
'---------------------------------------------------------------------------------------
'
Private Sub setThemeShade(ByVal redC As Integer, ByVal greenC As Integer, ByVal blueC As Integer)
    
    Dim Ctrl As Control
    
    On Error GoTo setThemeShade_Error

    ' RGB(redC, greenC, blueC) is the background colour used by the lighter themes
    
    Me.BackColor = RGB(redC, greenC, blueC)
    
    ' all buttons must be set to graphical
    For Each Ctrl In Me.Controls
        If (TypeOf Ctrl Is CommandButton) Or (TypeOf Ctrl Is CheckBox) Or (TypeOf Ctrl Is Label) Or (TypeOf Ctrl Is OptionButton) Or (TypeOf Ctrl Is Frame) Then
          Ctrl.BackColor = RGB(redC, greenC, blueC)
        End If
    Next
    
    If redC = 212 Then
        'classicTheme = True
        mnuLight.Checked = False
        mnuDark.Checked = True
        
        Call setPrefsIconImagesDark
        
    Else
        'classicTheme = False
        mnuLight.Checked = True
        mnuDark.Checked = False
        
        Call setPrefsIconImagesLight
                
    End If
    
    'now change the color of the sliders.
'    widgetPrefs.sliAnimationInterval.BackColor = RGB(redC, greenC, blueC)
    'widgetPrefs.'sliWidgetSkew.BackColor = RGB(redC, greenC, blueC)
    sliGaugeSize.BackColor = RGB(redC, greenC, blueC)
    sliOpacity.BackColor = RGB(redC, greenC, blueC)
    txtAboutText.BackColor = RGB(redC, greenC, blueC)
    
    sPutINISetting "Software\SteampunkClockCalendar", "skinTheme", gblSkinTheme, gblSettingsFile

    On Error GoTo 0
    Exit Sub

setThemeShade_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setThemeShade of Module Module1"
            Resume Next
          End If
    End With
End Sub



'---------------------------------------------------------------------------------------
' Procedure : setThemeColour
' Author    : beededea
' Date      : 19/09/2019
' Purpose   : if the o/s is capable of supporting the classic theme it tests every 10 secs
'             to see if a theme has been switched
'
'---------------------------------------------------------------------------------------
'
Private Sub setThemeColour()
    
    Dim SysClr As Long: SysClr = 0
    
   On Error GoTo setThemeColour_Error
   'If debugflg = 1  Then Debug.Print "%setThemeColour"

    If IsThemeActive() = False Then
        'MsgBox "Windows Classic Theme detected"
        'set themed buttons to none
        Call setThemeShade(212, 208, 199)
        SysClr = GetSysColor(COLOR_BTNFACE)
        gblSkinTheme = "dark"
        
        mnuDark.Caption = "Dark Theme Enabled"
        mnuLight.Caption = "Light Theme Enable"

    Else
        Call setModernThemeColours
        mnuDark.Caption = "Dark Theme Enable"
        mnuLight.Caption = "Light Theme Enabled"
    End If

    storeThemeColour = SysClr

   On Error GoTo 0
   Exit Sub

setThemeColour_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setThemeColour of module module1"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : adjustPrefsTheme
' Author    : beededea
' Date      : 25/04/2023
' Purpose   : adjust the theme used by the prefs alone
'---------------------------------------------------------------------------------------
'
Private Sub adjustPrefsTheme()
   On Error GoTo adjustPrefsTheme_Error

    If gblSkinTheme <> vbNullString Then
        If gblSkinTheme = "dark" Then
            Call setThemeShade(212, 208, 199)
        Else
            Call setThemeShade(240, 240, 240)
        End If
    Else
        If classicThemeCapable = True Then
            mnuAuto.Caption = "Auto Theme Enabled - Click to Disable"
            themeTimer.Enabled = True
        Else
            gblSkinTheme = "light"
            Call setModernThemeColours
        End If
    End If

   On Error GoTo 0
   Exit Sub

adjustPrefsTheme_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure adjustPrefsTheme of Form widgetPrefs"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : setModernThemeColours
' Author    : beededea
' Date      : 02/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setModernThemeColours()
         
    Dim SysClr As Long: SysClr = 0
    
    On Error GoTo setModernThemeColours_Error
    
    'the volume controlPrefs.mnuAuto.Caption = "Auto Theme Selection Cannot be Enabled"

    'MsgBox "Windows Alternate Theme detected"
    SysClr = GetSysColor(COLOR_BTNFACE)
    If SysClr = 13160660 Then
        Call setThemeShade(212, 208, 199)
        gblSkinTheme = "dark"
    Else ' 15790320
        Call setThemeShade(240, 240, 240)
        gblSkinTheme = "light"
    End If

   On Error GoTo 0
   Exit Sub

setModernThemeColours_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setModernThemeColours of Module Module1"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : loadHigherResPrefsImages
' Author    : beededea
' Date      : 18/06/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub loadHigherResPrefsImages()
    
    On Error GoTo loadHigherResPrefsImages_Error
      
    If Me.WindowState = vbMinimized Then Exit Sub
        
    If mnuDark.Checked = True Then
        Call setPrefsIconImagesDark
    Else
        Call setPrefsIconImagesLight
    End If
    
   On Error GoTo 0
   Exit Sub

loadHigherResPrefsImages_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadHigherResPrefsImages of Form widgetPrefs"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : positionTimer_Timer
' Author    : beededea
' Date      : 27/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub positionTimer_Timer()
    ' save the current X and y position of this form to allow repositioning when restarting
    On Error GoTo positionTimer_Timer_Error
   
    Call writePrefsPosition

   On Error GoTo 0
   Exit Sub

positionTimer_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure positionTimer_Timer of Form widgetPrefs"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : chkEnableResizing_Click
' Author    : beededea
' Date      : 27/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkEnableResizing_Click()
   On Error GoTo chkEnableResizing_Click_Error

    If chkEnableResizing.Value = 1 Then
        prefsDynamicSizingFlg = True
        txtPrefsFontCurrentSize.Visible = True
        lblCurrentFontsTab.Visible = True
        'Call writePrefsPosition
        chkEnableResizing.Caption = "Disable Corner Resizing"
    Else
        prefsDynamicSizingFlg = False
        txtPrefsFontCurrentSize.Visible = False
        lblCurrentFontsTab.Visible = False
        Unload widgetPrefs
        Me.Show
        Call readPrefsPosition
        chkEnableResizing.Caption = "Enable Corner Resizing"
    End If
    
    Call setframeHeights

   On Error GoTo 0
   Exit Sub

chkEnableResizing_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkEnableResizing_Click of Form widgetPrefs"

End Sub

Private Sub chkEnableResizing_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnableBalloonTooltips = "1" Then CreateToolTip chkEnableResizing.hwnd, "This allows you to resize the whole prefs window by dragging the bottom right corner of the window. It provides an alternative method of supporting high DPI screens.", _
                  TTIconInfo, "Help on Resizing", , , , True
End Sub
 



'---------------------------------------------------------------------------------------
' Procedure : setframeHeights
' Author    : beededea
' Date      : 28/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setframeHeights()
   On Error GoTo setframeHeights_Error

    If prefsDynamicSizingFlg = True Then
        fraGeneral.Height = fraAbout.Height
        fraFonts.Height = fraAbout.Height
        fraConfig.Height = fraAbout.Height
        fraSounds.Height = fraAbout.Height
        fraPosition.Height = fraAbout.Height
        fraDevelopment.Height = fraAbout.Height
        fraWindow.Height = fraAbout.Height
        
        fraGeneral.Width = fraAbout.Width
        fraFonts.Width = fraAbout.Width
        fraConfig.Width = fraAbout.Width
        fraSounds.Width = fraAbout.Width
        fraPosition.Width = fraAbout.Width
        fraDevelopment.Width = fraAbout.Width
        fraWindow.Width = fraAbout.Width
    
        'If gblDpiAwareness = "1" Then
            ' save the initial positions of ALL the controls on the prefs form
            Call SaveSizes(widgetPrefs, prefsControlPositions(), prefsCurrentWidth, prefsCurrentHeight)
        'End If
    Else
        fraGeneral.Height = 6708
        fraConfig.Height = 6784
        fraSounds.Height = 3376
        fraPosition.Height = 7544
        fraFonts.Height = 4481
        fraWindow.Height = 6388
        fraDevelopment.Height = 6297
        fraAbout.Height = 8700
    End If
    
    


   On Error GoTo 0
   Exit Sub

setframeHeights_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setframeHeights of Form widgetPrefs"
End Sub




'---------------------------------------------------------------------------------------
' Procedure : setPrefsIconImagesDark
' Author    : beededea
' Date      : 22/06/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setPrefsIconImagesDark()
    
    On Error GoTo setPrefsIconImagesDark_Error
    
    Set imgGeneral.Picture = Cairo.ImageList("general-icon-dark").Picture
    Set imgConfig.Picture = Cairo.ImageList("config-icon-dark").Picture
    Set imgFonts.Picture = Cairo.ImageList("font-icon-dark").Picture
    Set imgSounds.Picture = Cairo.ImageList("sounds-icon-dark").Picture
    Set imgPosition.Picture = Cairo.ImageList("position-icon-dark").Picture
    Set imgDevelopment.Picture = Cairo.ImageList("development-icon-dark").Picture
    Set imgWindow.Picture = Cairo.ImageList("windows-icon-dark").Picture
    Set imgAbout.Picture = Cairo.ImageList("about-icon-dark").Picture
'
    Set imgGeneralClicked.Picture = Cairo.ImageList("general-icon-dark-clicked").Picture
    Set imgConfigClicked.Picture = Cairo.ImageList("config-icon-dark-clicked").Picture
    Set imgFontsClicked.Picture = Cairo.ImageList("font-icon-dark-clicked").Picture
    Set imgSoundsClicked.Picture = Cairo.ImageList("sounds-icon-dark-clicked").Picture
    Set imgPositionClicked.Picture = Cairo.ImageList("position-icon-dark-clicked").Picture
    Set imgDevelopmentClicked.Picture = Cairo.ImageList("development-icon-dark-clicked").Picture
    Set imgWindowClicked.Picture = Cairo.ImageList("windows-icon-dark-clicked").Picture
    Set imgAboutClicked.Picture = Cairo.ImageList("about-icon-dark-clicked").Picture

   On Error GoTo 0
   Exit Sub

setPrefsIconImagesDark_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setPrefsIconImagesDark of Form widgetPrefs"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : setPrefsIconImagesLight
' Author    : beededea
' Date      : 22/06/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setPrefsIconImagesLight()
    
    On Error GoTo setPrefsIconImagesLight_Error
    
    Set imgGeneral.Picture = Cairo.ImageList("general-icon-light").Picture
    Set imgConfig.Picture = Cairo.ImageList("config-icon-light").Picture
    Set imgFonts.Picture = Cairo.ImageList("font-icon-light").Picture
    Set imgSounds.Picture = Cairo.ImageList("sounds-icon-light").Picture
    Set imgPosition.Picture = Cairo.ImageList("position-icon-light").Picture
    Set imgDevelopment.Picture = Cairo.ImageList("development-icon-light").Picture
    Set imgWindow.Picture = Cairo.ImageList("windows-icon-light").Picture
    Set imgAbout.Picture = Cairo.ImageList("about-icon-light").Picture
'
    Set imgGeneralClicked.Picture = Cairo.ImageList("general-icon-light-clicked").Picture
    Set imgConfigClicked.Picture = Cairo.ImageList("config-icon-light-clicked").Picture
    Set imgFontsClicked.Picture = Cairo.ImageList("font-icon-light-clicked").Picture
    Set imgSoundsClicked.Picture = Cairo.ImageList("sounds-icon-light-clicked").Picture
    Set imgPositionClicked.Picture = Cairo.ImageList("position-icon-light-clicked").Picture
    Set imgDevelopmentClicked.Picture = Cairo.ImageList("development-icon-light-clicked").Picture
    Set imgWindowClicked.Picture = Cairo.ImageList("windows-icon-light-clicked").Picture
    Set imgAboutClicked.Picture = Cairo.ImageList("about-icon-light-clicked").Picture

   On Error GoTo 0
   Exit Sub

setPrefsIconImagesLight_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setPrefsIconImagesLight of Form widgetPrefs"

End Sub

Private Sub txtPrefsFontCurrentSize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnableBalloonTooltips = "1" Then CreateToolTip txtPrefsFontCurrentSize.hwnd, "This is a read-only text box. It displays the current font as set when dynamic form resizing is enabled. Drag the right hand corner of the window downward and the form will auto-resize. This text box will display the resized font currently in operation for informational purposes only.", _
                  TTIconInfo, "Help on Setting the Font size Dynamically", , , , True
End Sub



'---------------------------------------------------------------------------------------
' Procedure : IsWinNTPlus
' Author    : Randy Birch
' Date      : 13/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Private Function IsWinNTPlus() As Boolean
'
'   'returns True if running WinNT or better
'   On Error GoTo IsWinNTPlus_Error
'
'   #If Win32 Then
'
'      Dim OSV As OSVERSIONINFO
'
'      OSV.OSVSize = Len(OSV)
'
'      If GetVersionEx(OSV) = 1 Then
'
'         IsWinNTPlus = (OSV.PlatformID = VER_PLATFORM_WIN32_NT)
'
'      End If
'
'   #End If
'
'   On Error GoTo 0
'   Exit Function
'
'IsWinNTPlus_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsWinNTPlus of Form widgetPrefs"
'
'End Function

'---------------------------------------------------------------------------------------
' Procedure : lblDragCorner_MouseDown
' Author    : beededea
' Date      : 07/10/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblDragCorner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error GoTo lblDragCorner_MouseDown_Error
    
    If Button = vbLeftButton Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, 0
    End If
    
    On Error GoTo 0
    Exit Sub

lblDragCorner_MouseDown_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblDragCorner_MouseDown of Form widgetPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : lblDragCorner_MouseMove
' Author    : beededea
' Date      : 07/10/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblDragCorner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error GoTo lblDragCorner_MouseMove_Error

    lblDragCorner.MousePointer = 8

    On Error GoTo 0
    Exit Sub

lblDragCorner_MouseMove_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblDragCorner_MouseMove of Form widgetPrefs"
   
End Sub


'---------------------------------------------------------------------------------------
' Procedure : setPrefsFormZordering
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : set the z-ordering of the window
'---------------------------------------------------------------------------------------
'
Public Sub setPrefsFormZordering()

   On Error GoTo setPrefsFormZordering_Error

'    If Val(gblWindowLevel) = 0 Then
'        Call SetWindowPos(Me.hwnd, HWND_BOTTOM, 0&, 0&, 0&, 0&, OnTopFlags)
'    ElseIf Val(gblWindowLevel) = 1 Then
'        Call SetWindowPos(Me.hwnd, HWND_TOP, 0&, 0&, 0&, 0&, OnTopFlags)
'    ElseIf Val(gblWindowLevel) = 2 Then
'        Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, OnTopFlags)
'    End If

   On Error GoTo 0
   Exit Sub

setPrefsFormZordering_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setPrefsFormZordering of Module modMain"
End Sub


'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'  --- All folded content will be temporary put under this lines ---
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'CODEFOLD STORAGE:
'CODEFOLD STORAGE END:
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'--- If you're Subclassing: Move the CODEFOLD STORAGE up as needed ---
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\


