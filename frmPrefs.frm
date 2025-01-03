VERSION 5.00
Object = "{BCE37951-37DF-4D69-A8A3-2CFABEE7B3CC}#1.0#0"; "CCRSlider.ocx"
Begin VB.Form widgetPrefs 
   Caption         =   "Steampunk Clock Calendar Preferences"
   ClientHeight    =   10545
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   8880
   Icon            =   "frmPrefs.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   10693.53
   ScaleMode       =   0  'User
   ScaleWidth      =   8880
   Visible         =   0   'False
   Begin VB.Timer tmrPrefsMonitorSaveHeight 
      Interval        =   5000
      Left            =   -90
      Top             =   5220
   End
   Begin VB.Timer tmrPrefsScreenResolution 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   -90
      Top             =   6420
   End
   Begin VB.Frame fraDevelopmentButton 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   5490
      TabIndex        =   38
      Top             =   0
      Width           =   1065
      Begin VB.Label lblDevelopment 
         Caption         =   "Development"
         Height          =   240
         Left            =   45
         TabIndex        =   39
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
   Begin VB.Timer tmrWritePosition 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   -75
      Top             =   6990
   End
   Begin VB.CheckBox chkEnableResizing 
      Caption         =   "Enable Corner Resize"
      Height          =   210
      Left            =   3240
      TabIndex        =   125
      Top             =   10125
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Frame fraAboutButton 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   7695
      TabIndex        =   92
      Top             =   0
      Width           =   975
      Begin VB.Label lblAbout 
         Caption         =   "About"
         Height          =   240
         Index           =   0
         Left            =   255
         TabIndex        =   93
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
      TabIndex        =   40
      Top             =   -15
      Width           =   930
      Begin VB.Label lblConfig 
         Caption         =   "Config."
         Height          =   240
         Index           =   0
         Left            =   270
         TabIndex        =   41
         Top             =   855
         Width           =   510
      End
      Begin VB.Image imgConfig 
         Height          =   600
         Left            =   165
         Picture         =   "frmPrefs.frx":227B
         Stretch         =   -1  'True
         Top             =   240
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
      TabIndex        =   36
      Top             =   0
      Width           =   930
      Begin VB.Label lblPosition 
         Caption         =   "Position"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   37
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
      Top             =   10035
      Width           =   1320
   End
   Begin VB.CommandButton btnHelp 
      Caption         =   "Help"
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   35
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
      Left            =   -90
      Top             =   5835
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
   Begin VB.Frame fraGeneral 
      Caption         =   "General"
      ForeColor       =   &H80000008&
      Height          =   7275
      Left            =   75
      TabIndex        =   46
      Top             =   1380
      Visible         =   0   'False
      Width           =   7995
      Begin VB.Frame fraGeneralInner 
         BorderStyle     =   0  'None
         Height          =   6555
         Left            =   435
         TabIndex        =   47
         Top             =   300
         Width           =   7395
         Begin VB.CheckBox chk24HourClockMode 
            Caption         =   "Toggle 24hr clock mode ON *"
            Height          =   465
            Left            =   1995
            TabIndex        =   181
            ToolTipText     =   "Check this box to enable the automatic start of the program when Windows is started."
            Top             =   2205
            Width           =   4020
         End
         Begin VB.CheckBox chkTogglePendulum 
            Caption         =   "Toggle the pendulum animation"
            Height          =   465
            Left            =   1980
            TabIndex        =   169
            ToolTipText     =   "Check this box to enable the automatic start of the program when Windows is started."
            Top             =   1725
            Width           =   4020
         End
         Begin VB.Frame Frame 
            BorderStyle     =   0  'None
            Height          =   3795
            Left            =   -1005
            TabIndex        =   146
            Top             =   3285
            Width           =   8325
            Begin VB.CommandButton btnNow 
               Caption         =   "o"
               Height          =   345
               Index           =   5
               Left            =   2115
               Style           =   1  'Graphical
               TabIndex        =   232
               Top             =   2160
               Width           =   330
            End
            Begin VB.CommandButton btnNow 
               Caption         =   "o"
               Height          =   345
               Index           =   4
               Left            =   2115
               Style           =   1  'Graphical
               TabIndex        =   231
               Top             =   1665
               Width           =   330
            End
            Begin VB.CommandButton btnNow 
               Caption         =   "o"
               Height          =   345
               Index           =   3
               Left            =   2115
               Style           =   1  'Graphical
               TabIndex        =   230
               Top             =   1185
               Width           =   330
            End
            Begin VB.CommandButton btnNow 
               Caption         =   "o"
               Height          =   345
               Index           =   2
               Left            =   2115
               Style           =   1  'Graphical
               TabIndex        =   229
               Top             =   705
               Width           =   330
            End
            Begin VB.CommandButton btnclearAlarm 
               Caption         =   "Clear"
               Height          =   345
               Index           =   5
               Left            =   7680
               Style           =   1  'Graphical
               TabIndex        =   228
               Top             =   2145
               Width           =   540
            End
            Begin VB.CommandButton btnclearAlarm 
               Caption         =   "Clear"
               Height          =   345
               Index           =   4
               Left            =   7680
               Style           =   1  'Graphical
               TabIndex        =   227
               Top             =   1650
               Width           =   540
            End
            Begin VB.CommandButton btnclearAlarm 
               Caption         =   "Clear"
               Height          =   345
               Index           =   3
               Left            =   7680
               Style           =   1  'Graphical
               TabIndex        =   226
               Top             =   1170
               Width           =   540
            End
            Begin VB.CommandButton btnclearAlarm 
               Caption         =   "Clear"
               Height          =   345
               Index           =   2
               Left            =   7680
               Style           =   1  'Graphical
               TabIndex        =   225
               Top             =   690
               Width           =   540
            End
            Begin VB.CommandButton btnclearAlarm 
               Caption         =   "Clear"
               Height          =   345
               Index           =   1
               Left            =   7695
               Style           =   1  'Graphical
               TabIndex        =   224
               Top             =   180
               Width           =   540
            End
            Begin VB.CommandButton btnNow 
               Caption         =   "o"
               Height          =   345
               Index           =   1
               Left            =   2115
               Style           =   1  'Graphical
               TabIndex        =   223
               Top             =   210
               Width           =   330
            End
            Begin VB.ComboBox cmbAlarm5Minutes 
               Height          =   315
               Left            =   5340
               Style           =   2  'Dropdown List
               TabIndex        =   220
               Top             =   2160
               Width           =   525
            End
            Begin VB.ComboBox cmbAlarm4Minutes 
               Height          =   315
               Left            =   5340
               Style           =   2  'Dropdown List
               TabIndex        =   219
               Top             =   1680
               Width           =   525
            End
            Begin VB.ComboBox cmbAlarm3Minutes 
               Height          =   315
               Left            =   5340
               Style           =   2  'Dropdown List
               TabIndex        =   218
               Top             =   1215
               Width           =   525
            End
            Begin VB.ComboBox cmbAlarm2Minutes 
               Height          =   315
               Left            =   5340
               Style           =   2  'Dropdown List
               TabIndex        =   217
               Top             =   720
               Width           =   525
            End
            Begin VB.ComboBox cmbAlarm5Hours 
               Height          =   315
               Left            =   4785
               Style           =   2  'Dropdown List
               TabIndex        =   216
               Top             =   2175
               Width           =   525
            End
            Begin VB.ComboBox cmbAlarm4Hours 
               Height          =   315
               Left            =   4785
               Style           =   2  'Dropdown List
               TabIndex        =   215
               Top             =   1680
               Width           =   525
            End
            Begin VB.ComboBox cmbAlarm3Hours 
               Height          =   315
               Left            =   4785
               Style           =   2  'Dropdown List
               TabIndex        =   214
               Top             =   1215
               Width           =   525
            End
            Begin VB.ComboBox cmbAlarm2Hours 
               Height          =   315
               Left            =   4785
               Style           =   2  'Dropdown List
               TabIndex        =   213
               Top             =   720
               Width           =   525
            End
            Begin VB.ComboBox cmbAlarm5Year 
               Height          =   315
               Left            =   3945
               Style           =   2  'Dropdown List
               TabIndex        =   212
               Top             =   2175
               Width           =   750
            End
            Begin VB.ComboBox cmbAlarm4Year 
               Height          =   315
               Left            =   3945
               Style           =   2  'Dropdown List
               TabIndex        =   211
               Top             =   1680
               Width           =   750
            End
            Begin VB.ComboBox cmbAlarm3Year 
               Height          =   315
               Left            =   3945
               Style           =   2  'Dropdown List
               TabIndex        =   210
               Top             =   1215
               Width           =   750
            End
            Begin VB.ComboBox cmbAlarm2Year 
               Height          =   315
               Left            =   3945
               Style           =   2  'Dropdown List
               TabIndex        =   209
               Top             =   720
               Width           =   750
            End
            Begin VB.ComboBox cmbAlarm5Month 
               Height          =   315
               Left            =   3060
               Style           =   2  'Dropdown List
               TabIndex        =   208
               Top             =   2175
               Width           =   870
            End
            Begin VB.ComboBox cmbAlarm4Month 
               Height          =   315
               Left            =   3060
               Style           =   2  'Dropdown List
               TabIndex        =   207
               Top             =   1680
               Width           =   870
            End
            Begin VB.ComboBox cmbAlarm3Month 
               Height          =   315
               Left            =   3060
               Style           =   2  'Dropdown List
               TabIndex        =   206
               Top             =   1215
               Width           =   870
            End
            Begin VB.ComboBox cmbAlarm5Day 
               Height          =   315
               Left            =   2505
               Style           =   2  'Dropdown List
               TabIndex        =   205
               Top             =   2175
               Width           =   525
            End
            Begin VB.ComboBox cmbAlarm4Day 
               Height          =   315
               Left            =   2505
               Style           =   2  'Dropdown List
               TabIndex        =   204
               Top             =   1680
               Width           =   525
            End
            Begin VB.ComboBox cmbAlarm3Day 
               Height          =   315
               Left            =   2505
               Style           =   2  'Dropdown List
               TabIndex        =   203
               Top             =   1215
               Width           =   525
            End
            Begin VB.ComboBox cmbAlarm1Year 
               Height          =   315
               Left            =   3945
               Style           =   2  'Dropdown List
               TabIndex        =   202
               Top             =   210
               Width           =   750
            End
            Begin VB.ComboBox cmbAlarm1Minutes 
               Height          =   315
               Left            =   5340
               Style           =   2  'Dropdown List
               TabIndex        =   201
               Top             =   210
               Width           =   525
            End
            Begin VB.ComboBox cmbAlarm1Hours 
               Height          =   315
               Left            =   4785
               Style           =   2  'Dropdown List
               TabIndex        =   200
               Top             =   210
               Width           =   525
            End
            Begin VB.ComboBox cmbAlarm1Month 
               Height          =   315
               Left            =   3060
               Style           =   2  'Dropdown List
               TabIndex        =   199
               Top             =   210
               Width           =   870
            End
            Begin VB.ComboBox cmbAlarm2Day 
               Height          =   315
               Left            =   2505
               Style           =   2  'Dropdown List
               TabIndex        =   198
               Top             =   720
               Width           =   525
            End
            Begin VB.ComboBox cmbAlarm2Month 
               Height          =   315
               Left            =   3060
               Style           =   2  'Dropdown List
               TabIndex        =   197
               Top             =   720
               Width           =   870
            End
            Begin VB.ComboBox cmbAlarm1Day 
               Height          =   315
               Left            =   2505
               Style           =   2  'Dropdown List
               TabIndex        =   196
               Top             =   210
               Width           =   525
            End
            Begin VB.CommandButton btnVerifyDateTime5 
               Caption         =   ">"
               Height          =   315
               Left            =   5895
               Style           =   1  'Graphical
               TabIndex        =   161
               ToolTipText     =   "Verify Date Time for alarm number 1"
               Top             =   2160
               Width           =   300
            End
            Begin VB.TextBox txtAlarm5Time 
               Height          =   315
               Left            =   7140
               Locked          =   -1  'True
               TabIndex        =   160
               Top             =   2160
               Width           =   500
            End
            Begin VB.CommandButton btnVerifyDateTime4 
               Caption         =   ">"
               Height          =   315
               Left            =   5895
               Style           =   1  'Graphical
               TabIndex        =   158
               ToolTipText     =   "Verify Date Time for alarm number 1"
               Top             =   1650
               Width           =   300
            End
            Begin VB.TextBox txtAlarm4Time 
               Height          =   315
               Left            =   7140
               Locked          =   -1  'True
               TabIndex        =   157
               Top             =   1665
               Width           =   500
            End
            Begin VB.CommandButton btnVerifyDateTime3 
               Caption         =   ">"
               Height          =   315
               Left            =   5895
               Style           =   1  'Graphical
               TabIndex        =   155
               ToolTipText     =   "Verify Date Time for alarm number 1"
               Top             =   1185
               Width           =   300
            End
            Begin VB.TextBox txtAlarm3Time 
               Height          =   315
               Left            =   7140
               Locked          =   -1  'True
               TabIndex        =   154
               Top             =   1185
               Width           =   500
            End
            Begin VB.CommandButton btnVerifyDateTime2 
               Caption         =   ">"
               Height          =   315
               Left            =   5895
               Style           =   1  'Graphical
               TabIndex        =   152
               ToolTipText     =   "Verify Date Time for alarm number 1"
               Top             =   705
               Width           =   300
            End
            Begin VB.TextBox txtAlarm2Time 
               Height          =   315
               Left            =   7140
               Locked          =   -1  'True
               TabIndex        =   151
               Top             =   705
               Width           =   500
            End
            Begin VB.TextBox txtAlarm1Time 
               Height          =   315
               Left            =   7140
               Locked          =   -1  'True
               TabIndex        =   148
               Top             =   210
               Width           =   495
            End
            Begin VB.CommandButton btnVerifyDateTime1 
               Caption         =   ">"
               Height          =   300
               Left            =   5895
               Style           =   1  'Graphical
               TabIndex        =   149
               ToolTipText     =   "Verify Date Time for alarm number 1"
               Top             =   225
               Width           =   300
            End
            Begin VB.TextBox txtAlarm5Date 
               Height          =   315
               Left            =   6255
               Locked          =   -1  'True
               TabIndex        =   159
               Top             =   2160
               Width           =   825
            End
            Begin VB.TextBox txtAlarm4Date 
               Height          =   315
               Left            =   6255
               Locked          =   -1  'True
               TabIndex        =   156
               Top             =   1665
               Width           =   825
            End
            Begin VB.TextBox txtAlarm3Date 
               Height          =   315
               Left            =   6255
               Locked          =   -1  'True
               TabIndex        =   153
               Top             =   1185
               Width           =   825
            End
            Begin VB.TextBox txtAlarm2Date 
               Height          =   315
               Left            =   6255
               Locked          =   -1  'True
               TabIndex        =   150
               Top             =   705
               Width           =   825
            End
            Begin VB.TextBox txtAlarm1Date 
               Height          =   315
               Left            =   6255
               Locked          =   -1  'True
               TabIndex        =   147
               Top             =   210
               Width           =   825
            End
            Begin VB.Label lblGeneral 
               Caption         =   $"frmPrefs.frx":53C2
               Height          =   960
               Index           =   10
               Left            =   2070
               TabIndex        =   167
               Top             =   2670
               Width           =   4230
            End
            Begin VB.Label lblGeneral 
               Caption         =   "Alarm No. 5 :"
               Height          =   375
               Index           =   9
               Left            =   1005
               TabIndex        =   166
               Tag             =   "lblRefreshInterval"
               Top             =   2220
               Width           =   1095
            End
            Begin VB.Label lblGeneral 
               Caption         =   "Alarm No. 4 :"
               Height          =   375
               Index           =   8
               Left            =   1005
               TabIndex        =   165
               Tag             =   "lblRefreshInterval"
               Top             =   1725
               Width           =   1095
            End
            Begin VB.Label lblGeneral 
               Caption         =   "Alarm No. 3 :"
               Height          =   375
               Index           =   7
               Left            =   1005
               TabIndex        =   164
               Tag             =   "lblRefreshInterval"
               Top             =   1245
               Width           =   1095
            End
            Begin VB.Label lblGeneral 
               Caption         =   "Alarm No. 2 :"
               Height          =   375
               Index           =   5
               Left            =   1005
               TabIndex        =   163
               Tag             =   "lblRefreshInterval"
               Top             =   765
               Width           =   1095
            End
            Begin VB.Label lblGeneral 
               Caption         =   "Alarm No. 1 :"
               Height          =   375
               Index           =   4
               Left            =   1005
               TabIndex        =   162
               Tag             =   "lblRefreshInterval"
               Top             =   270
               Width           =   1095
            End
         End
         Begin VB.CheckBox chkWidgetFunctions 
            Caption         =   "Toggle to enable/disable"
            Height          =   390
            Left            =   1995
            TabIndex        =   48
            ToolTipText     =   "When checked this box enables the spinning earth functionality. That's it!"
            Top             =   165
            Width           =   3405
         End
         Begin VB.CheckBox chkGenStartup 
            Caption         =   "Run the Steampunk Clock Calendar Widget at Windows Startup "
            Height          =   465
            Left            =   1995
            TabIndex        =   88
            ToolTipText     =   "Check this box to enable the automatic start of the program when Windows is started."
            Top             =   1230
            Width           =   4020
         End
         Begin VB.Label lblGeneral 
            Caption         =   "day        month          year            hr        min"
            Height          =   375
            Index           =   13
            Left            =   1575
            TabIndex        =   233
            Tag             =   "lblRefreshInterval"
            Top             =   3075
            Width           =   4485
         End
         Begin VB.Label lblGeneral 
            Caption         =   "Digital Clock :"
            Height          =   375
            Index           =   12
            Left            =   480
            TabIndex        =   182
            Tag             =   "lblRefreshInterval"
            Top             =   2310
            Width           =   1740
         End
         Begin VB.Label lblGeneral 
            Caption         =   "Pendulum Swing :"
            Height          =   375
            Index           =   3
            Left            =   480
            TabIndex        =   170
            Tag             =   "lblRefreshInterval"
            Top             =   1845
            Width           =   1740
         End
         Begin VB.Label lblGeneral 
            Caption         =   "Alarm Dates and times shown below"
            Height          =   435
            Index           =   2
            Left            =   2010
            TabIndex        =   168
            Top             =   2700
            Width           =   4215
         End
         Begin VB.Label lblGeneral 
            Caption         =   "When checked this box enables the functionality of this widget to control audio - That's it! *"
            Height          =   465
            Index           =   1
            Left            =   1995
            TabIndex        =   144
            Top             =   675
            Width           =   4020
         End
         Begin VB.Label lblGeneral 
            Caption         =   "Widget Functions :"
            Height          =   315
            Index           =   6
            Left            =   495
            TabIndex        =   90
            Top             =   255
            Width           =   1395
         End
         Begin VB.Label lblGeneral 
            Caption         =   "Auto Start :"
            Height          =   375
            Index           =   11
            Left            =   975
            TabIndex        =   89
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
      TabIndex        =   94
      Top             =   1170
      Visible         =   0   'False
      Width           =   8520
      Begin VB.CommandButton btnGithubHome 
         Caption         =   "Github &Home"
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
         Left            =   6765
         Style           =   1  'Graphical
         TabIndex        =   195
         ToolTipText     =   "Here you can visit the update location where you can download new versions of the programs."
         Top             =   300
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
         Left            =   6765
         Style           =   1  'Graphical
         TabIndex        =   95
         ToolTipText     =   "Opens a browser window and sends you to our donate page on Amazon"
         Top             =   1800
         Width           =   1470
      End
      Begin VB.Frame fraScrollbarCover 
         BorderStyle     =   0  'None
         Height          =   6225
         Left            =   7980
         TabIndex        =   108
         Top             =   2205
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
         TabIndex        =   107
         Text            =   "frmPrefs.frx":547F
         Top             =   2205
         Width           =   7935
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
         Left            =   6765
         Style           =   1  'Graphical
         TabIndex        =   98
         ToolTipText     =   "This gives access to the debugging tool"
         Top             =   1425
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
         Left            =   6765
         Style           =   1  'Graphical
         TabIndex        =   97
         ToolTipText     =   "This will link you to the Rocket/Steamy dock users Group"
         Top             =   1050
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
         Left            =   6765
         Style           =   1  'Graphical
         TabIndex        =   96
         ToolTipText     =   "Here you can visit the update location where you can download new versions of the programs."
         Top             =   675
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
         TabIndex        =   112
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
         TabIndex        =   111
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
         TabIndex        =   110
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
         TabIndex        =   109
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
         TabIndex        =   106
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
         TabIndex        =   105
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
         TabIndex        =   104
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
         TabIndex        =   103
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
         TabIndex        =   102
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
         TabIndex        =   101
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
         TabIndex        =   100
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
         TabIndex        =   99
         Top             =   510
         Width           =   2655
      End
   End
   Begin VB.Frame fraConfig 
      Caption         =   "Configuration"
      Height          =   7035
      Left            =   225
      TabIndex        =   8
      Top             =   1230
      Width           =   7605
      Begin VB.Frame fraConfigInner 
         BorderStyle     =   0  'None
         Height          =   6345
         Left            =   435
         TabIndex        =   33
         Top             =   435
         Width           =   6450
         Begin vb6projectCCRSlider.Slider sliGaugeSize 
            Height          =   390
            Left            =   1920
            TabIndex        =   221
            Top             =   30
            Width           =   3870
            _ExtentX        =   6826
            _ExtentY        =   688
            Min             =   5
            Max             =   200
            Value           =   5
            TickFrequency   =   3
            SelStart        =   5
         End
         Begin VB.Frame fraPrefsTooltips 
            BorderStyle     =   0  'None
            Height          =   870
            Left            =   1860
            TabIndex        =   186
            Top             =   3405
            Width           =   3150
            Begin VB.OptionButton optEnablePrefsBalloonTooltips 
               Caption         =   "Prefs - Enable Balloon Tooltips *"
               Height          =   195
               Left            =   135
               TabIndex        =   188
               Top             =   150
               Width           =   2760
            End
            Begin VB.OptionButton optEnablePrefsTooltips 
               Caption         =   "Prefs - Enable SquareTooltips *"
               Height          =   195
               Left            =   135
               TabIndex        =   187
               Top             =   480
               Width           =   2970
            End
         End
         Begin VB.Frame fraMainTooltips 
            BorderStyle     =   0  'None
            Height          =   645
            Left            =   1995
            TabIndex        =   183
            Top             =   2775
            Width           =   4095
            Begin VB.OptionButton optEnableBalloonTooltips 
               Caption         =   "Clock - Enable Balloon Tooltips *"
               Height          =   315
               Left            =   0
               TabIndex        =   185
               Top             =   -30
               Width           =   3150
            End
            Begin VB.OptionButton optEnableTooltips 
               Caption         =   "Clock - Enable Square Tooltips"
               Height          =   300
               Left            =   0
               TabIndex        =   184
               Top             =   300
               Width           =   2790
            End
         End
         Begin VB.CheckBox chkShowHelp 
            Caption         =   "Show Help on Widget Start"
            Height          =   225
            Left            =   2010
            TabIndex        =   145
            ToolTipText     =   "Check the box to show the widget in the taskbar"
            Top             =   4725
            Width           =   3405
         End
         Begin VB.CheckBox chkDpiAwareness 
            Caption         =   "DPI Awareness Enable *"
            Height          =   285
            Left            =   2010
            TabIndex        =   135
            ToolTipText     =   "Check the box to make the program DPI aware. RESTART required."
            Top             =   5070
            Width           =   3405
         End
         Begin VB.CheckBox chkShowTaskbar 
            Caption         =   "Show Widget in Taskbar"
            Height          =   225
            Left            =   2010
            TabIndex        =   133
            ToolTipText     =   "Check the box to show the widget in the taskbar"
            Top             =   4350
            Width           =   3405
         End
         Begin VB.ComboBox cmbScrollWheelDirection 
            Height          =   315
            Left            =   1995
            Style           =   2  'Dropdown List
            TabIndex        =   83
            ToolTipText     =   "To change the direction of the mouse scroll wheel when resiziing the globe gauge."
            Top             =   1695
            Width           =   2490
         End
         Begin VB.Label lblConfiguration 
            Caption         =   $"frmPrefs.frx":6436
            Height          =   915
            Index           =   0
            Left            =   1980
            TabIndex        =   136
            Top             =   5430
            Width           =   4335
         End
         Begin VB.Label lblConfiguration 
            Caption         =   "The scroll-wheel resizing direction can be determined here. The direction chosen causes the gauge to grow. *"
            Height          =   675
            Index           =   6
            Left            =   2025
            TabIndex        =   113
            Top             =   2115
            Width           =   3990
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "160"
            Height          =   315
            Index           =   4
            Left            =   4740
            TabIndex        =   87
            Top             =   555
            Width           =   345
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "120"
            Height          =   315
            Index           =   3
            Left            =   3990
            TabIndex        =   86
            Top             =   555
            Width           =   345
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "50"
            Height          =   315
            Index           =   1
            Left            =   2790
            TabIndex        =   85
            Top             =   555
            Width           =   345
         End
         Begin VB.Label lblConfiguration 
            Caption         =   "Mouse Wheel Resize :"
            Height          =   345
            Index           =   3
            Left            =   255
            TabIndex        =   84
            ToolTipText     =   "To change the direction of the mouse scroll wheel when resiziing the globe gauge."
            Top             =   1740
            Width           =   2055
         End
         Begin VB.Label lblConfiguration 
            Caption         =   "Adjust to a percentage of the original size. You can also use Ctrl+Mousewheel. Immediate. *"
            Height          =   555
            Index           =   2
            Left            =   2070
            TabIndex        =   82
            ToolTipText     =   "Adjust to a percentage of the original size. You can also use Ctrl+Mousewheel."
            Top             =   870
            Width           =   3810
         End
         Begin VB.Label lblConfiguration 
            Caption         =   "Gauge Size :"
            Height          =   315
            Index           =   1
            Left            =   885
            TabIndex        =   81
            Top             =   105
            Width           =   975
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "80"
            Height          =   315
            Index           =   2
            Left            =   3360
            TabIndex        =   80
            Top             =   555
            Width           =   360
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "200 (%)"
            Height          =   315
            Index           =   5
            Left            =   5385
            TabIndex        =   79
            Top             =   555
            Width           =   735
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "5"
            Height          =   315
            Index           =   0
            Left            =   2085
            TabIndex        =   78
            Top             =   555
            Width           =   345
         End
      End
   End
   Begin VB.Frame fraWindow 
      Caption         =   "Window"
      Height          =   7935
      Left            =   225
      TabIndex        =   10
      Top             =   1620
      Width           =   8235
      Begin VB.Frame fraWindowInner 
         BorderStyle     =   0  'None
         Height          =   7500
         Left            =   165
         TabIndex        =   14
         Top             =   345
         Width           =   7005
         Begin vb6projectCCRSlider.Slider sliOpacity 
            Height          =   390
            Left            =   2115
            TabIndex        =   222
            Top             =   4575
            Width           =   3870
            _ExtentX        =   6826
            _ExtentY        =   688
            Min             =   20
            Max             =   100
            Value           =   20
            SelStart        =   20
         End
         Begin VB.ComboBox cmbMultiMonitorResize 
            Height          =   315
            Left            =   2220
            Style           =   2  'Dropdown List
            TabIndex        =   180
            Top             =   5805
            Width           =   3720
         End
         Begin VB.Frame fraHiding 
            BorderStyle     =   0  'None
            Height          =   2010
            Left            =   1395
            TabIndex        =   114
            Top             =   2325
            Width           =   5130
            Begin VB.ComboBox cmbHidingTime 
               Height          =   315
               Left            =   825
               Style           =   2  'Dropdown List
               TabIndex        =   117
               Top             =   1605
               Width           =   3720
            End
            Begin VB.CheckBox chkWidgetHidden 
               Caption         =   "Hiding Widget *"
               Height          =   225
               Left            =   855
               TabIndex        =   115
               Top             =   210
               Width           =   2955
            End
            Begin VB.Label lblWindowLevel 
               Caption         =   "Hiding :"
               Height          =   345
               Index           =   2
               Left            =   90
               TabIndex        =   118
               Top             =   210
               Width           =   720
            End
            Begin VB.Label lblWindowLevel 
               Caption         =   $"frmPrefs.frx":64EA
               Height          =   975
               Index           =   1
               Left            =   855
               TabIndex        =   116
               Top             =   600
               Width           =   3900
            End
         End
         Begin VB.ComboBox cmbWindowLevel 
            Height          =   315
            Left            =   2220
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   0
            Width           =   3720
         End
         Begin VB.CheckBox chkIgnoreMouse 
            Caption         =   "Ignore Mouse *"
            Height          =   225
            Left            =   2250
            TabIndex        =   15
            ToolTipText     =   "Checking this box causes the program to ignore all mouse events."
            Top             =   1500
            Width           =   2535
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Multi-Monitor Resizing :"
            Height          =   255
            Index           =   11
            Left            =   375
            TabIndex        =   179
            Top             =   5835
            Width           =   1830
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   $"frmPrefs.frx":658D
            Height          =   1140
            Index           =   10
            Left            =   2235
            TabIndex        =   178
            Top             =   6255
            Width           =   4050
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "This setting controls the relative layering of this widget. You may use it to place it on top of other windows or underneath. "
            Height          =   660
            Index           =   3
            Left            =   2235
            TabIndex        =   123
            Top             =   570
            Width           =   3810
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Window Level :"
            Height          =   345
            Index           =   0
            Left            =   915
            TabIndex        =   23
            Top             =   60
            Width           =   1740
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "20%"
            Height          =   315
            Index           =   7
            Left            =   2205
            TabIndex        =   22
            Top             =   5070
            Width           =   345
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "100%"
            Height          =   315
            Index           =   9
            Left            =   5565
            TabIndex        =   21
            Top             =   5070
            Width           =   405
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "60%"
            Height          =   315
            Index           =   8
            Left            =   3975
            TabIndex        =   20
            Top             =   5070
            Width           =   840
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Opacity:"
            Height          =   315
            Index           =   6
            Left            =   1470
            TabIndex        =   19
            Top             =   4620
            Width           =   780
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Set the program transparency level."
            Height          =   330
            Index           =   5
            Left            =   2250
            TabIndex        =   18
            Top             =   5385
            Width           =   3810
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Checking this box causes the program to ignore all mouse events except right click menu interactions."
            Height          =   660
            Index           =   4
            Left            =   2235
            TabIndex        =   17
            Top             =   1890
            Width           =   3810
         End
      End
   End
   Begin VB.Frame fraFonts 
      Caption         =   "Fonts"
      Height          =   4470
      Left            =   240
      TabIndex        =   9
      Top             =   1230
      Width           =   8280
      Begin VB.Frame fraFontsInner 
         BorderStyle     =   0  'None
         Height          =   3810
         Left            =   690
         TabIndex        =   25
         Top             =   360
         Width           =   6105
         Begin VB.TextBox txtDisplayScreenFont 
            Height          =   315
            Left            =   1695
            Locked          =   -1  'True
            TabIndex        =   191
            Text            =   "Courier  New"
            Top             =   1830
            Width           =   3285
         End
         Begin VB.CommandButton btnDisplayScreenFont 
            Caption         =   "Font"
            Height          =   300
            Left            =   5010
            Style           =   1  'Graphical
            TabIndex        =   190
            Top             =   1830
            Width           =   585
         End
         Begin VB.TextBox txtDisplayScreenFontSize 
            Height          =   315
            Left            =   1695
            Locked          =   -1  'True
            TabIndex        =   189
            Text            =   "8"
            Top             =   2370
            Width           =   510
         End
         Begin VB.CommandButton btnResetMessages 
            Caption         =   "Reset"
            Height          =   300
            Left            =   1680
            Style           =   1  'Graphical
            TabIndex        =   141
            Top             =   3270
            Width           =   885
         End
         Begin VB.TextBox txtPrefsFontCurrentSize 
            Height          =   315
            Left            =   4200
            Locked          =   -1  'True
            TabIndex        =   126
            ToolTipText     =   "Disabled for manual input. Shows the current font size when form resizing is enabled."
            Top             =   1065
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.TextBox txtPrefsFontSize 
            Height          =   315
            Left            =   1710
            Locked          =   -1  'True
            TabIndex        =   28
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
            TabIndex        =   27
            Top             =   90
            Width           =   585
         End
         Begin VB.TextBox txtPrefsFont 
            Height          =   315
            Left            =   1710
            Locked          =   -1  'True
            TabIndex        =   26
            Text            =   "Times New Roman"
            Top             =   90
            Width           =   3285
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Choose a font to be used for the text in the console display screen on the main clock *"
            Height          =   480
            Index           =   9
            Left            =   2415
            TabIndex        =   194
            Top             =   2355
            Width           =   4035
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Display Console Font :"
            Height          =   300
            Index           =   8
            Left            =   0
            TabIndex        =   193
            Tag             =   "lblPrefsFont"
            Top             =   1890
            Width           =   1665
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Console  Font Size :"
            Height          =   330
            Index           =   5
            Left            =   165
            TabIndex        =   192
            Tag             =   "lblPrefsFontSize"
            Top             =   2400
            Width           =   1590
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Hidden message boxes can be reactivated by pressing this reset button."
            Height          =   480
            Index           =   4
            Left            =   2670
            TabIndex        =   142
            Top             =   3210
            Width           =   3360
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Reset Pop ups :"
            Height          =   300
            Index           =   1
            Left            =   405
            TabIndex        =   140
            Tag             =   "lblPrefsFont"
            Top             =   3315
            Width           =   1470
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Resized Font"
            Height          =   315
            Index           =   10
            Left            =   4920
            TabIndex        =   127
            Top             =   1110
            Visible         =   0   'False
            Width           =   2400
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "The chosen font size *"
            Height          =   480
            Index           =   7
            Left            =   2310
            TabIndex        =   32
            Top             =   1095
            Width           =   2400
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Base Font Size :"
            Height          =   330
            Index           =   3
            Left            =   435
            TabIndex        =   31
            Tag             =   "lblPrefsFontSize"
            Top             =   1095
            Width           =   1230
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Prefs Utility Font :"
            Height          =   300
            Index           =   2
            Left            =   360
            TabIndex        =   30
            Tag             =   "lblPrefsFont"
            Top             =   120
            Width           =   1665
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Choose a font to be used for the text in this preferences window, gauge tooltips and message boxes *"
            Height          =   480
            Index           =   6
            Left            =   1695
            TabIndex        =   29
            Top             =   480
            Width           =   4035
         End
      End
   End
   Begin VB.Frame fraDevelopment 
      Caption         =   "Development"
      Height          =   6210
      Left            =   240
      TabIndex        =   44
      Top             =   1200
      Visible         =   0   'False
      Width           =   8520
      Begin VB.Frame fraDevelopmentInner 
         BorderStyle     =   0  'None
         Height          =   5595
         Left            =   870
         TabIndex        =   45
         Top             =   300
         Width           =   7455
         Begin VB.Frame fraDefaultEditor 
            BorderStyle     =   0  'None
            Height          =   2370
            Left            =   75
            TabIndex        =   128
            Top             =   3165
            Width           =   7290
            Begin VB.CommandButton btnDefaultEditor 
               Caption         =   "..."
               Height          =   300
               Left            =   5115
               Style           =   1  'Graphical
               TabIndex        =   130
               ToolTipText     =   "Click to select the .vbp file to edit the program - You need to have access to the source!"
               Top             =   210
               Width           =   315
            End
            Begin VB.TextBox txtDefaultEditor 
               Height          =   315
               Left            =   1440
               TabIndex        =   129
               Text            =   " eg. E:\vb6\fire call\FireCallWin.vbp"
               Top             =   195
               Width           =   3660
            End
            Begin VB.Label lblGitHub 
               Caption         =   $"frmPrefs.frx":66A4
               ForeColor       =   &H8000000D&
               Height          =   915
               Left            =   1560
               TabIndex        =   134
               ToolTipText     =   "Double Click to visit github"
               Top             =   1440
               Width           =   4935
            End
            Begin VB.Label lblDebug 
               Caption         =   $"frmPrefs.frx":676B
               Height          =   930
               Index           =   9
               Left            =   1545
               TabIndex        =   132
               Top             =   690
               Width           =   4785
            End
            Begin VB.Label lblDebug 
               Caption         =   "Default Editor :"
               Height          =   255
               Index           =   7
               Left            =   285
               TabIndex        =   131
               Tag             =   "lblSharedInputFile"
               Top             =   225
               Width           =   1350
            End
         End
         Begin VB.TextBox txtDblClickCommand 
            Height          =   315
            Left            =   1515
            TabIndex        =   56
            ToolTipText     =   "Enter a Windows command for the gauge to operate when double-clicked."
            Top             =   1095
            Width           =   3660
         End
         Begin VB.CommandButton btnOpenFile 
            Caption         =   "..."
            Height          =   300
            Left            =   5175
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Click to select a particular file for the gauge to run or open when double-clicked."
            Top             =   2250
            Width           =   315
         End
         Begin VB.TextBox txtOpenFile 
            Height          =   315
            Left            =   1515
            TabIndex        =   52
            ToolTipText     =   "Enter a particular file for the gauge to run or open when double-clicked."
            Top             =   2235
            Width           =   3660
         End
         Begin VB.ComboBox cmbDebug 
            Height          =   315
            ItemData        =   "frmPrefs.frx":680F
            Left            =   1530
            List            =   "frmPrefs.frx":6811
            Style           =   2  'Dropdown List
            TabIndex        =   49
            ToolTipText     =   "Choose to set debug mode."
            Top             =   -15
            Width           =   2160
         End
         Begin VB.Label lblDebug 
            Caption         =   "DblClick Command :"
            Height          =   510
            Index           =   1
            Left            =   -15
            TabIndex        =   58
            Tag             =   "lblPrefixString"
            Top             =   1155
            Width           =   1545
         End
         Begin VB.Label lblConfigurationTab 
            Caption         =   "Shift+double-clicking on the widget image will open this file. "
            Height          =   375
            Index           =   6
            Left            =   1560
            TabIndex        =   57
            Top             =   2730
            Width           =   3705
         End
         Begin VB.Label lblDebug 
            Caption         =   "Default command to run when the gauge receives a double-click eg.  mmsys.cpl to run the sounds utility."
            Height          =   570
            Index           =   5
            Left            =   1590
            TabIndex        =   55
            Tag             =   "lblSharedInputFileDesc"
            Top             =   1605
            Width           =   4410
         End
         Begin VB.Label lblDebug 
            Caption         =   "Open File :"
            Height          =   255
            Index           =   4
            Left            =   645
            TabIndex        =   54
            Tag             =   "lblSharedInputFile"
            Top             =   2280
            Width           =   1350
         End
         Begin VB.Label lblDebug 
            Caption         =   "Turning on the debugging will provide extra information in the debug window.  *"
            Height          =   495
            Index           =   2
            Left            =   1545
            TabIndex        =   51
            Tag             =   "lblAlarmSoundDesc"
            Top             =   450
            Width           =   4455
         End
         Begin VB.Label lblDebug 
            Caption         =   "Debug :"
            Height          =   375
            Index           =   0
            Left            =   855
            TabIndex        =   50
            Tag             =   "lblAlarmSound"
            Top             =   45
            Width           =   1740
         End
      End
   End
   Begin VB.Frame fraPosition 
      Caption         =   "Position"
      Height          =   7440
      Left            =   255
      TabIndex        =   42
      Top             =   1230
      Visible         =   0   'False
      Width           =   8520
      Begin VB.Frame fraPositionInner 
         BorderStyle     =   0  'None
         Height          =   6960
         Left            =   150
         TabIndex        =   43
         Top             =   300
         Width           =   7680
         Begin VB.TextBox txtLandscapeHoffset 
            Height          =   315
            Left            =   2250
            TabIndex        =   70
            Top             =   4425
            Width           =   2130
         End
         Begin VB.CheckBox chkPreventDragging 
            Caption         =   "Widget Position Locked. *"
            Height          =   225
            Left            =   2265
            TabIndex        =   121
            ToolTipText     =   "Checking this box turns off the ability to drag the program with the mouse, locking it in position."
            Top             =   3465
            Width           =   2505
         End
         Begin VB.TextBox txtPortraitYoffset 
            Height          =   315
            Left            =   2250
            TabIndex        =   76
            ToolTipText     =   "Enter a prefix/nickname for outgoing messages."
            Top             =   6465
            Width           =   2130
         End
         Begin VB.TextBox txtPortraitHoffset 
            Height          =   315
            Left            =   2250
            TabIndex        =   74
            ToolTipText     =   "Enter a prefix/nickname for outgoing messages."
            Top             =   6000
            Width           =   2130
         End
         Begin VB.TextBox txtLandscapeVoffset 
            Height          =   315
            Left            =   2250
            TabIndex        =   72
            ToolTipText     =   "Enter a prefix/nickname for outgoing messages."
            Top             =   4875
            Width           =   2130
         End
         Begin VB.ComboBox cmbWidgetLandscape 
            Height          =   315
            Left            =   2250
            Style           =   2  'Dropdown List
            TabIndex        =   68
            Top             =   3930
            Width           =   2160
         End
         Begin VB.ComboBox cmbWidgetPortrait 
            Height          =   315
            Left            =   2250
            Style           =   2  'Dropdown List
            TabIndex        =   65
            ToolTipText     =   "Choose the alarm sound."
            Top             =   5505
            Width           =   2160
         End
         Begin VB.ComboBox cmbWidgetPosition 
            Height          =   315
            Left            =   2265
            Style           =   2  'Dropdown List
            TabIndex        =   62
            ToolTipText     =   "Choose the alarm sound."
            Top             =   2100
            Width           =   2160
         End
         Begin VB.ComboBox cmbAspectHidden 
            Height          =   315
            Left            =   2265
            Style           =   2  'Dropdown List
            TabIndex        =   59
            ToolTipText     =   "Choose the alarm sound."
            Top             =   0
            Width           =   2160
         End
         Begin VB.Label lblPosition 
            Caption         =   "(px)"
            Height          =   300
            Index           =   7
            Left            =   4530
            TabIndex        =   138
            Tag             =   "lblPrefixString"
            Top             =   6495
            Width           =   435
         End
         Begin VB.Label lblPosition 
            Caption         =   "(px)"
            Height          =   300
            Index           =   5
            Left            =   4530
            TabIndex        =   137
            Tag             =   "lblPrefixString"
            Top             =   6045
            Width           =   435
         End
         Begin VB.Label lblPosition 
            Caption         =   "*"
            Height          =   255
            Index           =   1
            Left            =   4545
            TabIndex        =   124
            Tag             =   "lblAlarmSound"
            Top             =   45
            Width           =   345
         End
         Begin VB.Label lblPosition 
            Caption         =   "(px)"
            Height          =   300
            Index           =   4
            Left            =   4530
            TabIndex        =   120
            Tag             =   "lblPrefixString"
            Top             =   4905
            Width           =   435
         End
         Begin VB.Label lblPosition 
            Caption         =   "(px)"
            Height          =   300
            Index           =   2
            Left            =   4530
            TabIndex        =   119
            Tag             =   "lblPrefixString"
            Top             =   4500
            Width           =   390
         End
         Begin VB.Label lblPosition 
            Caption         =   "Portrait Top Y pos :"
            Height          =   510
            Index           =   17
            Left            =   645
            TabIndex        =   77
            Tag             =   "lblPrefixString"
            Top             =   6480
            Width           =   2175
         End
         Begin VB.Label lblPosition 
            Caption         =   "Portrait Left X pos :"
            Height          =   510
            Index           =   16
            Left            =   660
            TabIndex        =   75
            Tag             =   "lblPrefixString"
            Top             =   6015
            Width           =   2175
         End
         Begin VB.Label lblPosition 
            Caption         =   "Landscape Top Y pos :"
            Height          =   510
            Index           =   15
            Left            =   420
            TabIndex        =   73
            Tag             =   "lblPrefixString"
            Top             =   4905
            Width           =   2175
         End
         Begin VB.Label lblPosition 
            Caption         =   "Landscape Left X pos :"
            Height          =   510
            Index           =   14
            Left            =   420
            TabIndex        =   71
            Tag             =   "lblPrefixString"
            Top             =   4455
            Width           =   2175
         End
         Begin VB.Label lblPosition 
            Caption         =   "Locked in Landscape :"
            Height          =   435
            Index           =   13
            Left            =   450
            TabIndex        =   69
            Tag             =   "lblAlarmSound"
            Top             =   3975
            Width           =   2115
         End
         Begin VB.Label lblPosition 
            Caption         =   $"frmPrefs.frx":6813
            Height          =   3435
            Index           =   12
            Left            =   5145
            TabIndex        =   67
            Tag             =   "lblAlarmSoundDesc"
            Top             =   3480
            Width           =   2520
         End
         Begin VB.Label lblPosition 
            Caption         =   "Locked in Portrait :"
            Height          =   375
            Index           =   11
            Left            =   690
            TabIndex        =   66
            Tag             =   "lblAlarmSound"
            Top             =   5550
            Width           =   2040
         End
         Begin VB.Label lblPosition 
            Caption         =   $"frmPrefs.frx":69E5
            Height          =   705
            Index           =   10
            Left            =   2250
            TabIndex        =   64
            Tag             =   "lblAlarmSoundDesc"
            Top             =   2550
            Width           =   5325
         End
         Begin VB.Label lblPosition 
            Caption         =   "Widget Position by Percent:"
            Height          =   375
            Index           =   8
            Left            =   195
            TabIndex        =   63
            Tag             =   "lblAlarmSound"
            Top             =   2145
            Width           =   2355
         End
         Begin VB.Label lblPosition 
            Caption         =   $"frmPrefs.frx":6A84
            Height          =   3045
            Index           =   6
            Left            =   2265
            TabIndex        =   61
            Tag             =   "lblAlarmSoundDesc"
            Top             =   450
            Width           =   5370
         End
         Begin VB.Label lblPosition 
            Caption         =   "Aspect Ratio Hidden Mode :"
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   60
            Tag             =   "lblAlarmSound"
            Top             =   45
            Width           =   2145
         End
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
         TabIndex        =   24
         Top             =   285
         Width           =   6420
         Begin VB.CheckBox chkVolumeBoost 
            Caption         =   "Decide whether the various sounds will be boosted."
            Height          =   225
            Left            =   1485
            TabIndex        =   175
            ToolTipText     =   "Check this box to enable or disable all of the sounds used during any animation on the main screen."
            Top             =   1680
            Width           =   4365
         End
         Begin VB.CheckBox chkEnableChimes 
            Caption         =   "Decide whether the clock chimes will sound."
            Height          =   225
            Left            =   1485
            TabIndex        =   173
            ToolTipText     =   "Check this box to enable or disable all of the sounds used during any animation on the main screen."
            Top             =   1215
            Width           =   4365
         End
         Begin VB.CheckBox chkEnableTicks 
            Caption         =   "Decide whether the clock has the tick sounds enabled."
            Height          =   225
            Left            =   1485
            TabIndex        =   171
            ToolTipText     =   "Check this box to enable or disable all of the sounds used during any animation on the main screen."
            Top             =   765
            Width           =   4365
         End
         Begin VB.CheckBox chkEnableSounds 
            Caption         =   "Enable ALL sounds for the whole widget."
            Height          =   225
            Left            =   1485
            TabIndex        =   34
            ToolTipText     =   "Check this box to enable or disable all of the sounds used during any animation on the main screen."
            Top             =   285
            Width           =   4485
         End
         Begin VB.Label lblSoundsTab 
            Caption         =   "Determine the sound of UI elements, clock tick and chiming volumes. Set the overall volume to loud or quiet."
            Height          =   540
            Index           =   4
            Left            =   1470
            TabIndex        =   177
            Tag             =   "lblSharedInputFile"
            Top             =   2100
            Width           =   4680
         End
         Begin VB.Label lblSoundsTab 
            Caption         =   "Volume Boost :"
            Height          =   255
            Index           =   2
            Left            =   300
            TabIndex        =   176
            Tag             =   "lblSharedInputFile"
            Top             =   1680
            Width           =   1185
         End
         Begin VB.Label lblSoundsTab 
            Caption         =   "Chimes :"
            Height          =   255
            Index           =   1
            Left            =   765
            TabIndex        =   174
            Tag             =   "lblSharedInputFile"
            Top             =   1215
            Width           =   765
         End
         Begin VB.Label lblSoundsTab 
            Caption         =   "Ticks :"
            Height          =   255
            Index           =   0
            Left            =   885
            TabIndex        =   172
            Tag             =   "lblSharedInputFile"
            Top             =   765
            Width           =   765
         End
         Begin VB.Label lblSoundsTab 
            Caption         =   "Audio :"
            Height          =   255
            Index           =   3
            Left            =   885
            TabIndex        =   91
            Tag             =   "lblSharedInputFile"
            Top             =   285
            Width           =   765
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
      Left            =   8700
      TabIndex        =   143
      ToolTipText     =   "drag me"
      Top             =   10350
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label lblSize 
      Caption         =   "Size in twips"
      Height          =   285
      Left            =   1875
      TabIndex        =   139
      Top             =   9780
      Visible         =   0   'False
      Width           =   4170
   End
   Begin VB.Label lblAsterix 
      Caption         =   "All controls marked with a * take effect immediately."
      Height          =   300
      Left            =   1920
      TabIndex        =   122
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
  (ByVal hWnd As Long, _
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
'Private gblPrefsLoadedFlg As Boolean
Private prefsDynamicSizingFlg As Boolean
Private lastFormHeight As Long
Private Const cPrefsFormHeight As Long = 11055
Private Const cPrefsFormWidth  As Long = 9090

'    gblPrefsCurrentWidth = 9075
'    gblPrefsCurrentHeight = 16450
'------------------------------------------------------ ENDS

Private prefsStartupFlg As Boolean
Private gblAllowSizeChangeFlg As Boolean

' module level balloon tooltip variables for subclassed comboBoxes ONLY.
Private gCmbMultiMonitorResizeBalloonTooltip As String
Private gCmbScrollWheelDirectionBalloonTooltip As String
Private gCmbWindowLevelBalloonTooltip As String
Private gCmbHidingTimeBalloonTooltip As String
Private gCmbAspectHiddenBalloonTooltip As String
Private gCmbWidgetPositionBalloonTooltip As String
Private gCmbWidgetLandscapeBalloonTooltip As String
Private gCmbWidgetPortraitBalloonTooltip As String
Private gCmbDebugBalloonTooltip As String
Private gCmbAlarmDayBalloonTooltip As String
Private gCmbAlarmMonthBalloonTooltip As String
Private gCmbAlarmYearBalloonTooltip As String
Private gCmbAlarmHoursBalloonTooltip As String
Private gCmbAlarmMinutesBalloonTooltip As String

Private gPrefsFormResizedByDrag As Boolean

Private mIsLoaded As Boolean

Private Sub btnAboutDebugInfo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip btnAboutDebugInfo.hWnd, "Here you can switch on Debug mode, not yet functional for this widget.", _
                  TTIconInfo, "Help on the Debug Info. Buttton", , , , True
End Sub

Private Sub btnclearAlarm_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip btnclearAlarm(Index).hWnd, "Clear all alarm fields in number " & Index & " alarm.", _
                  TTIconInfo, "Help on the Alarm Clear Buttton", , , , True
End Sub

Private Sub btnClose_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip btnClose.hWnd, "Close the Preference Utility", _
                  TTIconInfo, "Help on the Close Buttton", , , , True
End Sub

Private Sub btnDefaultEditor_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip btnDefaultEditor.hWnd, "Clicking on this button will cause a file explorer window to appear allowing you to select a Visual Basic Project (VBP) file for opening via the right click menu edit option. Once selected the adjacent text field will be automatically filled with the chosen path and file.", _
                  TTIconInfo, "Help on the VBP File Explorer Button", , , , True
End Sub

Private Sub btnDisplayScreenFont_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip btnDisplayScreenFont.hWnd, "This is the font selector button, if you click it the font selection window will pop up for you to select your chosen font. When resizing the main clock the display screen font size will change in relation to clock size. The base font determines the initial size, the resulting resized font will dynamically change. ", _
                  TTIconInfo, "Help on the Font Selector Button", , , , True
End Sub

Private Sub btnDonate_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip btnDonate.hWnd, "Here you can visit my KofI page and donate a Coffee if you like my creations.", _
                  TTIconInfo, "Help on the Donate Buttton", , , , True
End Sub

Private Sub btnFacebook_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip btnFacebook.hWnd, "Here you can visit the Facebook page for the steampunk Widget community.", _
                  TTIconInfo, "Help on the Update Buttton", , , , True
End Sub

Private Sub btnGithubHome_Click()
    Call menuForm.mnuGithubHome_Click
End Sub

Private Sub btnGithubHome_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip btnGithubHome.hWnd, "Here you can visit the widget's home page on github, when you click the button it will open a browser window and take you to the github home page.", _
                  TTIconInfo, "Help on the Update Buttton", , , , True
End Sub

Private Sub btnHelp_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip btnHelp.hWnd, "Opens the help document, this will open as a compiled HTML file.", _
                  TTIconInfo, "Help on the Help Buttton", , , , True
End Sub



Private Sub btnOpenFile_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip btnOpenFile.hWnd, "Clicking on this button will cause a file explorer window to appear allowing you to select any file you would like to execute on a shift+DBlClick. Once selected the adjacent text field will be automatically filled with the chosen path and file.", _
                  TTIconInfo, "Help on the shift+DBlClick File Explorer Button", , , , True
End Sub

Private Sub btnPrefsFont_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip btnPrefsFont.hWnd, "This is the font selector button, if you click it the font selection window will pop up for you to select your chosen font. Centurion Light SF is a good one and my personal favourite. When resizing the form (drag bottom right) the font size will change in relation to form height. The base font determines the initial size, the resulting resized font will dynamically change. ", _
                  TTIconInfo, "Help on Setting the Font Selector Button", , , , True
End Sub

Private Sub btnSave_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip btnSave.hWnd, "Save the changes you have made to the preferences", _
                  TTIconInfo, "Help on the Save Buttton", , , , True
End Sub



Private Sub btnUpdate_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip btnUpdate.hWnd, "Here you can able to download a new version of the program from github, when you click the button it will open a browser window and take you to the github page.", _
                  TTIconInfo, "Help on the Update Buttton", , , , True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnVerifyDateTime1_Click
' Author    : beededea
' Date      : 24/07/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnVerifyDateTime1_Click()
    Dim alarmTimeStatus As Boolean: alarmTimeStatus = False
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString
    
    Dim dayNo As Integer: dayNo = 0
    Dim dayString As String: dayString = vbNullString
    Dim monthNo As Integer: monthNo = 0
    Dim monthString As String: monthString = vbNullString
    Dim yearNo As Integer: yearNo = 0
    Dim yearString As String: yearString = vbNullString
    
    On Error GoTo btnVerifyDateTime1_Click_Error
    
    If cmbAlarm1Day.ListIndex <= 0 Or cmbAlarm1Month.ListIndex <= 0 Or cmbAlarm1Year.ListIndex <= 0 Or cmbAlarm1Hours.ListIndex <= 0 Or cmbAlarm1Minutes.ListIndex <= 0 Then Exit Sub

    dayNo = cmbAlarm1Day.List(cmbAlarm1Day.ListIndex)
    If dayNo <= 9 Then
        dayString = "0" & CStr(dayNo)
    Else
        dayString = CStr(dayNo)
    End If

    monthNo = cmbAlarm1Month.ListIndex
    If monthNo <= 9 Then
        monthString = "0" & CStr(monthNo)
    Else
        monthString = CStr(monthNo)
    End If

    yearNo = cmbAlarm1Year.List(cmbAlarm1Year.ListIndex)
    yearString = CStr(yearNo)
    
    txtAlarm1Date.Text = dayString & "/" & monthString & "/" & yearString
    txtAlarm1Time.Text = cmbAlarm1Hours.List(cmbAlarm1Hours.ListIndex) & ":" & cmbAlarm1Minutes.List(cmbAlarm1Minutes.ListIndex)
   
    If txtAlarm1Date.Text = "Alarm not yet set" Or txtAlarm1Time.Text = vbNullString Then
        answerMsg = "Alarm not yet set!"
        answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Message", False, "btnVerifyDateTime1_Click1")
        Exit Sub
    End If
    
    If txtAlarm1Date.Text <> vbNullString Then
        alarmTimeStatus = fVerifyAlarmTime(txtAlarm1Date.Text, txtAlarm1Time.Text)
    End If
    
    If alarmTimeStatus = True Then
        txtAlarm1Date.BackColor = vbWhite
        txtAlarm1Time.BackColor = vbWhite
        answerMsg = "Alarm date and time formats both valid and in the future"
        answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Information", False, "btnVerifyDateTime1_Click4")
    Else
        txtAlarm1Date.BackColor = vbRed
        txtAlarm1Time.BackColor = vbRed
        
        answerMsg = "Alarm time format invalid"
        answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Message", False, "btnVerifyDateTime1_Click3")
    End If
    
   On Error GoTo 0
   Exit Sub

btnVerifyDateTime1_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnVerifyDateTime1_Click of Form widgetPrefs"
    
End Sub


Private Sub btnVerifyDateTime1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip btnVerifyDateTime1.hWnd, "Validate a future date and time here in 24hr military format, eg: 23/11/2036 23:45. If the date and time are not valid nor in the future the invalid fields will be highlighted in red.", _
                  TTIconInfo, "Help on Validating Alarm Number One", , , , True
End Sub
Private Sub btnVerifyDateTime2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip btnVerifyDateTime2.hWnd, "Validate a future date and time here in 24hr military format, eg: 23/11/2036 23:45. If the date and time are not valid nor in the future the invalid fields will be highlighted in red.", _
                  TTIconInfo, "Help on Validating Alarm Number Two", , , True
End Sub
Private Sub btnVerifyDateTime3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip btnVerifyDateTime3.hWnd, "Validate a future date and time here in 24hr military format, eg: 23/11/2036 23:45. If the date and time are not valid nor in the future the invalid fields will be highlighted in red.", _
                  TTIconInfo, "Help on Validating Alarm Number Three", , , , True
End Sub
Private Sub btnVerifyDateTime4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip btnVerifyDateTime4.hWnd, "Validate a future date and time here in 24hr military format, eg: 23/11/2036 23:45. If the date and time are not valid nor in the future the invalid fields will be highlighted in red.", _
                  TTIconInfo, "Help on Validating Alarm Number Four", , , , True
End Sub
Private Sub btnVerifyDateTime5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip btnVerifyDateTime5.hWnd, "Validate a future date and time here in 24hr military format, eg: 23/11/2036 23:45. If the date and time are not valid nor in the future the invalid fields will be highlighted in red.", _
                  TTIconInfo, "Help on Validating Alarm Number Five", , , , True
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
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString
    
    Dim dayNo As Integer: dayNo = 0
    Dim dayString As String: dayString = vbNullString
    Dim monthNo As Integer: monthNo = 0
    Dim monthString As String: monthString = vbNullString
    Dim yearNo As Integer: yearNo = 0
    Dim yearString As String: yearString = vbNullString
    
    On Error GoTo btnVerifyDateTime2_Click_Error

    If cmbAlarm2Day.ListIndex <= 0 Or cmbAlarm2Month.ListIndex <= 0 Or cmbAlarm2Year.ListIndex <= 0 Or cmbAlarm2Hours.ListIndex <= 0 Or cmbAlarm2Minutes.ListIndex <= 0 Then Exit Sub
    
    dayNo = cmbAlarm2Day.List(cmbAlarm2Day.ListIndex)
    If dayNo <= 9 Then
        dayString = "0" & CStr(dayNo)
    Else
        dayString = CStr(dayNo)
    End If

    monthNo = cmbAlarm2Month.ListIndex
    If monthNo <= 9 Then
        monthString = "0" & CStr(monthNo)
    Else
        monthString = CStr(monthNo)
    End If

    yearNo = cmbAlarm2Year.List(cmbAlarm2Year.ListIndex)
    yearString = CStr(yearNo)
    
    txtAlarm2Date.Text = dayString & "/" & monthString & "/" & yearString
    txtAlarm2Time.Text = cmbAlarm2Hours.List(cmbAlarm2Hours.ListIndex) & ":" & cmbAlarm2Minutes.List(cmbAlarm2Minutes.ListIndex)
   
    If txtAlarm2Date.Text = "Alarm not yet set" Or txtAlarm2Time.Text = vbNullString Then
        answerMsg = "Alarm not yet set!"
        answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Message", False, "btnVerifyDateTime2_Click1")
        Exit Sub
    End If
    
    If txtAlarm2Date.Text <> vbNullString Then
        alarmTimeStatus = fVerifyAlarmTime(txtAlarm2Date.Text, txtAlarm2Time.Text)
    End If
    
    If alarmTimeStatus = True Then
        txtAlarm2Date.BackColor = vbWhite
        txtAlarm2Time.BackColor = vbWhite
        answerMsg = "Alarm date and time formats both valid and in the future"
        answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Information", False, "btnVerifyDateTime2_Click4")
    Else
        txtAlarm2Date.BackColor = vbRed
        txtAlarm2Time.BackColor = vbRed
        
        answerMsg = "Alarm time format invalid"
        answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Message", False, "btnVerifyDateTime2_Click3")
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
    
    futureDate = DateDiff("s", Now(), datefield)
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
    
    futureTime = DateDiff("s", Now(), datefield & " " & timefield)
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
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString
    
    Dim dayNo As Integer: dayNo = 0
    Dim dayString As String: dayString = vbNullString
    Dim monthNo As Integer: monthNo = 0
    Dim monthString As String: monthString = vbNullString
    Dim yearNo As Integer: yearNo = 0
    Dim yearString As String: yearString = vbNullString
    
    On Error GoTo btnVerifyDateTime3_Click_Error

    If cmbAlarm3Day.ListIndex <= 0 Or cmbAlarm3Month.ListIndex <= 0 Or cmbAlarm3Year.ListIndex <= 0 Or cmbAlarm3Hours.ListIndex <= 0 Or cmbAlarm3Minutes.ListIndex <= 0 Then Exit Sub
    
    dayNo = cmbAlarm3Day.List(cmbAlarm3Day.ListIndex)
    If dayNo <= 9 Then
        dayString = "0" & CStr(dayNo)
    Else
        dayString = CStr(dayNo)
    End If

    monthNo = cmbAlarm3Month.ListIndex
    If monthNo <= 9 Then
        monthString = "0" & CStr(monthNo)
    Else
        monthString = CStr(monthNo)
    End If

    yearNo = cmbAlarm3Year.List(cmbAlarm3Year.ListIndex)
    yearString = CStr(yearNo)
    
    txtAlarm3Date.Text = dayString & "/" & monthString & "/" & yearString
    txtAlarm3Time.Text = cmbAlarm3Hours.List(cmbAlarm3Hours.ListIndex) & ":" & cmbAlarm3Minutes.List(cmbAlarm3Minutes.ListIndex)
   
    If txtAlarm3Date.Text = "Alarm not yet set" Or txtAlarm3Time.Text = vbNullString Then
        answerMsg = "Alarm not yet set!"
        answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Message", False, "btnVerifyDateTime3_Click1")
        Exit Sub
    End If
    
    If txtAlarm3Date.Text <> vbNullString Then
        alarmTimeStatus = fVerifyAlarmTime(txtAlarm3Date.Text, txtAlarm3Time.Text)
    End If
    
    If alarmTimeStatus = True Then
        txtAlarm3Date.BackColor = vbWhite
        txtAlarm3Time.BackColor = vbWhite
        answerMsg = "Alarm date and time formats both valid and in the future"
        answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Information", False, "btnVerifyDateTime3_Click4")
    Else
        txtAlarm3Date.BackColor = vbRed
        txtAlarm3Time.BackColor = vbRed
        
        answerMsg = "Alarm time format invalid"
        answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Message", False, "btnVerifyDateTime3_Click3")
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
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString
    
    Dim dayNo As Integer: dayNo = 0
    Dim dayString As String: dayString = vbNullString
    Dim monthNo As Integer: monthNo = 0
    Dim monthString As String: monthString = vbNullString
    Dim yearNo As Integer: yearNo = 0
    Dim yearString As String: yearString = vbNullString
    
    On Error GoTo btnVerifyDateTime4_Click_Error
   
    If cmbAlarm4Day.ListIndex <= 0 Or cmbAlarm4Month.ListIndex <= 0 Or cmbAlarm4Year.ListIndex <= 0 Or cmbAlarm4Hours.ListIndex <= 0 Or cmbAlarm4Minutes.ListIndex <= 0 Then Exit Sub
    
    dayNo = cmbAlarm4Day.List(cmbAlarm4Day.ListIndex)
    If dayNo <= 9 Then
        dayString = "0" & CStr(dayNo)
    Else
        dayString = CStr(dayNo)
    End If

    monthNo = cmbAlarm4Month.ListIndex
    If monthNo <= 9 Then
        monthString = "0" & CStr(monthNo)
    Else
        monthString = CStr(monthNo)
    End If

    yearNo = cmbAlarm4Year.List(cmbAlarm4Year.ListIndex)
    yearString = CStr(yearNo)
    
    txtAlarm4Date.Text = dayString & "/" & monthString & "/" & yearString
    txtAlarm4Time.Text = cmbAlarm4Hours.List(cmbAlarm4Hours.ListIndex) & ":" & cmbAlarm4Minutes.List(cmbAlarm4Minutes.ListIndex)
   
    If txtAlarm4Date.Text = "Alarm not yet set" Or txtAlarm4Time.Text = vbNullString Then
        answerMsg = "Alarm not yet set!"
        answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Message", False, "btnVerifyDateTime4_Click1")
        Exit Sub
    End If
    
    If txtAlarm4Date.Text <> vbNullString Then
        alarmTimeStatus = fVerifyAlarmTime(txtAlarm4Date.Text, txtAlarm4Time.Text)
    End If
    
    If alarmTimeStatus = True Then
        txtAlarm4Date.BackColor = vbWhite
        txtAlarm4Time.BackColor = vbWhite
        answerMsg = "Alarm date and time formats both valid and in the future"
        answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Information", False, "btnVerifyDateTime4_Click4")
    Else
        txtAlarm4Date.BackColor = vbRed
        txtAlarm4Time.BackColor = vbRed
        
        answerMsg = "Alarm time format invalid"
        answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Message", False, "btnVerifyDateTime4_Click3")
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
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString
      
    Dim dayNo As Integer: dayNo = 0
    Dim dayString As String: dayString = vbNullString
    Dim monthNo As Integer: monthNo = 0
    Dim monthString As String: monthString = vbNullString
    Dim yearNo As Integer: yearNo = 0
    Dim yearString As String: yearString = vbNullString
    
    On Error GoTo btnVerifyDateTime5_Click_Error
   
    If cmbAlarm5Day.ListIndex <= 0 Or cmbAlarm5Month.ListIndex <= 0 Or cmbAlarm5Year.ListIndex <= 0 Or cmbAlarm5Hours.ListIndex <= 0 Or cmbAlarm5Minutes.ListIndex <= 0 Then Exit Sub
    
    dayNo = cmbAlarm5Day.List(cmbAlarm5Day.ListIndex)
    If dayNo <= 9 Then
        dayString = "0" & CStr(dayNo)
    Else
        dayString = CStr(dayNo)
    End If

    monthNo = cmbAlarm5Month.ListIndex
    If monthNo <= 9 Then
        monthString = "0" & CStr(monthNo)
    Else
        monthString = CStr(monthNo)
    End If

    yearNo = cmbAlarm5Year.List(cmbAlarm5Year.ListIndex)
    yearString = CStr(yearNo)
    
    txtAlarm5Date.Text = dayString & "/" & monthString & "/" & yearString
    txtAlarm5Time.Text = cmbAlarm5Hours.List(cmbAlarm5Hours.ListIndex) & ":" & cmbAlarm5Minutes.List(cmbAlarm5Minutes.ListIndex)
    
    If txtAlarm5Date.Text = "Alarm not yet set" Or txtAlarm5Time.Text = vbNullString Then
        answerMsg = "Alarm not yet set!"
        answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Message", False, "btnVerifyDateTime5_Click1")
        Exit Sub
    End If
    
    If txtAlarm5Date.Text <> vbNullString Then
        alarmTimeStatus = fVerifyAlarmTime(txtAlarm5Date.Text, txtAlarm5Time.Text)
    End If
    
    If alarmTimeStatus = True Then
        txtAlarm5Date.BackColor = vbWhite
        txtAlarm5Time.BackColor = vbWhite
        answerMsg = "Alarm date and time formats both valid and in the future"
        answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Information", False, "btnVerifyDateTime5_Click4")
    Else
        txtAlarm5Date.BackColor = vbRed
        txtAlarm5Time.BackColor = vbRed
        
        answerMsg = "Alarm time format invalid"
        answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Message", False, "btnVerifyDateTime5_Click3")
    End If
    
   On Error GoTo 0
   Exit Sub

btnVerifyDateTime5_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnVerifyDateTime5_Click of Form widgetPrefs"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chk24HourClockMode_Click
' Author    : beededea
' Date      : 15/10/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chk24HourClockMode_Click()
    
   On Error GoTo chk24HourClockMode_Click_Error

    btnSave.Enabled = True ' enable the save button
    
    If chk24HourClockMode.Value = 1 Then
        gbl24HourClockMode = "1"
    Else
        gbl24HourClockMode = "0"

    End If
    
    triggerDigitalClockPopulation = True
    overlayWidget.tmrDigitRotator.Enabled = True
    
    sPutINISetting "Software\SteampunkClockCalendar", "24HourClockMode", gbl24HourClockMode, gblSettingsFile

   On Error GoTo 0
   Exit Sub

chk24HourClockMode_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chk24HourClockMode_Click of Form widgetPrefs"

End Sub

Private Sub chk24HourClockMode_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip chk24HourClockMode.hWnd, "Decide whether the digital clock display should use military/twenty-four hour time or the 12-hour clock.", _
                  TTIconInfo, "Help on the Digital Clock 24hour Mode", , , , True
End Sub

Private Sub chkDpiAwareness_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip chkDpiAwareness.hWnd, "Check here to make the program DPI aware. NOT required on small to medium screens that are less than 1920 bytes wide. Try it and see which suits your system. RESTART required.", _
                  TTIconInfo, "Help on DPI Awareness Mode", , , , True
End Sub




Private Sub optEnableBalloonTooltips_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip optEnableBalloonTooltips.hWnd, "This setting enables/disables the enhanced balloon tooltips for elements within the Steampunk GUI. These tooltips are multi-line and in general more attractive, note that their font size will match the Windows system font size.", _
                  TTIconInfo, "Help on Balloon Tooltips on the GUI", , , , True
End Sub

Private Sub chkEnableChimes_Click()
    btnSave.Enabled = True ' enable the save button
End Sub

Private Sub chkEnableChimes_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip chkEnableChimes.hWnd, "Check this box to enable or disable the quarterly, half-hour and hourly chimes issued by the clock.", _
                  TTIconInfo, "Help on Enabling/Disabling Clock Chimes", , , , True
End Sub





'---------------------------------------------------------------------------------------
' Procedure : optEnablePrefsBalloonTooltips_Click
' Author    : beededea
' Date      : 06/10/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub optEnablePrefsBalloonTooltips_Click()

   On Error GoTo optEnablePrefsBalloonTooltips_Click_Error

    btnSave.Enabled = True ' enable the save button
    If optEnablePrefsBalloonTooltips.Value = True Then
        gblEnablePrefsBalloonTooltips = "True"
    Else
        gblEnablePrefsBalloonTooltips = "False"
    End If
         
    sPutINISetting "Software\SteampunkClockCalendar", "enablePrefsBalloonTooltips", gblEnablePrefsBalloonTooltips, gblSettingsFile
    
    ' set the tooltips on the prefs screen
    Call setPrefsTooltips

   On Error GoTo 0
   Exit Sub

optEnablePrefsBalloonTooltips_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optEnablePrefsBalloonTooltips_Click of Form widgetPrefs"
End Sub

Private Sub optEnablePrefsBalloonTooltips_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' NOTE: see setPrefsTooltips for cancelling this particular tooltip.
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip optEnablePrefsBalloonTooltips.hWnd, "This setting enables/disables the enhanced balloon tooltips for elements within this Preference Utility. These tooltips are multi-line and in general more attractive, note that their font size will match the Windows system font size.", _
                  TTIconInfo, "Help on Balloon Tooltips on the Preference Utility", , , , True
End Sub

Private Sub optEnablePrefsTooltips_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip optEnablePrefsTooltips.hWnd, "This setting enables/disables the native " & gblCodingEnvironment & " -style tooltips for elements within this Preference Utility.", _
                  TTIconInfo, "Help on  " & gblCodingEnvironment & " Native Style Tooltips on the Preference Utility", , , , True
End Sub




Private Sub chkEnableSounds_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip chkEnableSounds.hWnd, "Check this box to enable or disable all of the sounds used during any animation on the main steampunk GUI, as well as all other chimes, tick sounds &c.", _
                  TTIconInfo, "Help on Enabling/Disabling Sounds", , , , True
End Sub

Private Sub chkEnableTicks_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip chkEnableTicks.hWnd, "Check this box to enable or disable only the sound of the clock ticking.", _
                  TTIconInfo, "Help on Enabling/Disabling Clock Ticking Sound", , , , True
End Sub

Private Sub chkGenStartup_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip chkGenStartup.hWnd, "Check this box to enable the automatic start of the program when Windows is started.", _
                  TTIconInfo, "Help on the Widget Automatic Start Toggle", , , , True
End Sub

Private Sub chkIgnoreMouse_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip chkIgnoreMouse.hWnd, "Checking this box causes the program to ignore all mouse events. A strange option, a left-over from the Yahoo Widgets days that offered this additional option. Replicated here as a homage to the old widget platform.", _
                  TTIconInfo, "Help on the Ignore Mouse optEnableTooltips", , , , True
End Sub

Private Sub chkPreventDragging_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip chkPreventDragging.hWnd, "Checking this box causes the program to lock in place and ignore all attempts to move it with the mouse. " & vbCrLf & vbCrLf & _
        "The widget can be locked into a certain position in either landscape/portrait mode, ensuring that the widget always appears exactly where you want it to.  " & vbCrLf & vbCrLf & _
        "Using the fields adjacent, you can assign a default x/y position for both Landscape or Portrait mode.  " & vbCrLf & vbCrLf & _
        "When the widget is locked in place (using the Widget Position Locked option in the Window Tab), this value is set automatically.  " & vbCrLf & vbCrLf & _
        "A strange option, a left-over from the Yahoo Widgets days that offered this additional option. Replicated here as a homage to the old widget platform.", _
                  TTIconInfo, "Help on the Lock in Place optEnableTooltips", , , , True
End Sub

Private Sub chkShowHelp_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip chkShowHelp.hWnd, "Checking this box causes the rather attractive help canvas to appear every time the widget is started.", _
                  TTIconInfo, "Help on the Ignore Mouse optEnableTooltips", , , , True
End Sub

Private Sub chkShowTaskbar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip chkShowTaskbar.hWnd, "Check the box to show the widget in the Windows taskbar. A typical user may have multiple desktop widgets and it makes no sense to fill the taskbar with taskbar entries, this option allows you to enable a single one or two at your whim.", _
                  TTIconInfo, "Help on the Showing Entries in the Taskbar", , , , True
End Sub

Private Sub chkTogglePendulum_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip chkTogglePendulum.hWnd, "Decide whether the pendulum swings or not. The math behind the graphics of a swinging pendulum requires processor power. Older single-core machines may experience a lot of cpu usage.", _
                  TTIconInfo, "Help on the Pendulum Toggle", , , , True
End Sub

Private Sub chkVolumeBoost_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip chkVolumeBoost.hWnd, "Check this box to enable or disable the sound boost for all sounds made by this program.", _
                  TTIconInfo, "Help on Enabling/Disabling the Sound Boost", , , , True
End Sub

Private Sub chkWidgetFunctions_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip chkWidgetFunctions.hWnd, "When checked this box enables the spinning earth functionality. Any adjustment takes place instantly.", _
                  TTIconInfo, "Help on the Widget Function Toggle", , , , True
End Sub

Private Sub chkWidgetHidden_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip chkWidgetHidden.hWnd, "Checking this box causes the program to hide for a certain number of minutes. More useful from the widget's right click menu where you can hide the widget at will. Seemingly, a strange option, a left-over from the Yahoo Widgets days that offered this additional option. Replicated here as a homage to the old widget platform.", _
                  TTIconInfo, "Help on the Hidden optEnableTooltips", , , , True
End Sub

Private Sub cmbMultiMonitorResize_Click()
    btnSave.Enabled = True ' enable the save button
    
    If prefsStartupFlg = True Then Exit Sub
    
    gblMultiMonitorResize = CStr(cmbMultiMonitorResize.ListIndex)
    
    If cmbMultiMonitorResize.ListIndex = 2 Then
        If prefsMonitorStruct.IsPrimary = True Then
            gblClockPrimaryHeightRatio = fClock.clockForm.WidgetRoot.Zoom
            sPutINISetting "Software\SteampunkClockCalendar", "clockPrimaryHeightRatio", gblClockPrimaryHeightRatio, gblSettingsFile
            
            'gblPrefsPrimaryHeightTwips = Trim$(cstr(widgetPrefs.Height))
            sPutINISetting "Software\SteampunkClockCalendar", "prefsPrimaryHeightTwips", gblPrefsPrimaryHeightTwips, gblSettingsFile
        Else
            gblClockSecondaryHeightRatio = fClock.clockForm.WidgetRoot.Zoom
            sPutINISetting "Software\SteampunkClockCalendar", "clockSecondaryHeightRatio", gblClockSecondaryHeightRatio, gblSettingsFile
            
            'gblPrefsSecondaryHeightTwips = Trim$(cstr(widgetPrefs.Height))
            sPutINISetting "Software\SteampunkClockCalendar", "prefsSecondaryHeightTwips", gblPrefsSecondaryHeightTwips, gblSettingsFile
        End If
    End If
End Sub

Private Sub chkVolumeBoost_Click()
    btnSave.Enabled = True ' enable the save button
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
        'fClock.pendulumToggleEnabled = True
        'overlayWidget.SwingPendulum = True
        gblTogglePendulum = "1"
    Else
        'fClock.pendulumToggleEnabled = False
        'overlayWidget.SwingPendulum = False
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
    
    'gblPrefsLoadedFlg = False
    IsLoaded = False
    prefsDynamicSizingFlg = False
    lastFormHeight = 0

    On Error GoTo 0
    Exit Sub

Form_Initialize_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Initialize of Form widgetPrefs"
    
    End Sub




'---------------------------------------------------------------------------------------
' Procedure : Form_Load     WidgetPrefs
' Author    : beededea
' Date      : 25/04/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()
    Dim prefsFormHeight As Long: prefsFormHeight = 0
    Dim prefsFormMonitorID As Long: prefsFormMonitorID = 0

    On Error GoTo Form_Load_Error
    
    'Me.Show
        
    prefsStartupFlg = True ' this is used to prevent some control initialisations from running code at startup
    prefsDynamicSizingFlg = False
    'gblPrefsLoadedFlg = True ' this is a variable tested by an added form property to indicate whether the form is loaded or not
    IsLoaded = True
    gblWindowLevelWasChanged = False
        
    gblPrefsCurrentWidth = cPrefsFormWidth
    gblPrefsCurrentHeight = cPrefsFormHeight
    
    prefsFormHeight = gblPrefsCurrentHeight
    
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
    
    Me.mnuAbout.Caption = "About Steampunk Clock Calendar Cairo " & gblCodingEnvironment & " widget"
        
    ' subclass controls that need additional functionality that VB6 does not provide (scrollwheel/balloon tooltips)
    Call subClassControls
    
    ' note the monitor primary at the preferences form_load and store as oldClockFormMonitorPrimary
    prefsMonitorStruct = formScreenProperties(widgetPrefs, prefsFormMonitorID)
    oldPrefsFormMonitorPrimary = prefsMonitorStruct.IsPrimary ' -1 true
      
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
    
    ' set the height of the whole form not higher than the screen size, cause a form_resize event
    If gblDpiAwareness = "1" Then
        gblPrefsFormResizedInCode = True
        If prefsFormHeight < physicalScreenHeightTwips Then
            widgetPrefs.Height = prefsFormHeight ' 16450
        Else
            widgetPrefs.Height = physicalScreenHeightTwips - 1000
        End If
    End If
    
    ' position the prefs on the current monitor
    Call positionPrefsMonitor
    
    ' set the Z order of the prefs form
    Call setPrefsFormZordering
    
    ' make the busy sand timer invisible
    Call hideBusyTimer
    
    ' start the timer that records the prefs position every 10 seconds
    tmrWritePosition.Enabled = True
    
    ' start the timer that detects a MOVE event on the preferences form
    tmrPrefsScreenResolution.Enabled = True
    
    ' end the startup by un-setting the start flag
    prefsStartupFlg = False
    
    btnSave.Enabled = False

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form widgetPrefs"

End Sub


Private Sub hideBusyTimer()

    fClock.clockForm.Widgets("busy1").Widget.Alpha = 0
    fClock.clockForm.Widgets("busy2").Widget.Alpha = 0
    fClock.clockForm.Widgets("busy3").Widget.Alpha = 0
    fClock.clockForm.Widgets("busy4").Widget.Alpha = 0
    fClock.clockForm.Widgets("busy5").Widget.Alpha = 0
    fClock.clockForm.Widgets("busy6").Widget.Alpha = 0
    
    fClock.clockForm.Widgets("busy1").Widget.Refresh
    fClock.clockForm.Widgets("busy2").Widget.Refresh
    fClock.clockForm.Widgets("busy3").Widget.Refresh
    fClock.clockForm.Widgets("busy4").Widget.Refresh
    fClock.clockForm.Widgets("busy5").Widget.Refresh
    fClock.clockForm.Widgets("busy6").Widget.Refresh
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : subClassControls
' Author    : beededea
' Date      : 16/07/2024
' Purpose   : sub classing code to mostly intercept messages to the comboboxes to provide missing balloon tooltips functionality
'---------------------------------------------------------------------------------------
'
Private Sub subClassControls()
    
   On Error GoTo subClassControls_Error

    If InIDE Then
        MsgBox "NOTE: Running in IDE so Sub classing is disabled" & vbCrLf & "Mousewheel will not scroll icon maps and balloon tooltips will not display on comboboxes" & vbCrLf & vbCrLf & _
            "In addition, the display screen will not show messages as it currently crashes when run within the IDE."
    Else
        ' sub classing code to intercept messages to the form itself in order to capture WM_EXITSIZEMOVE messages
        
        Call SubclassForm(widgetPrefs.hWnd, ObjPtr(widgetPrefs))
        
        'now the comboboxes in order to capture the mouseOver and display the balloon tooltips
        
        Call SubclassComboBox(cmbMultiMonitorResize.hWnd, ObjPtr(cmbMultiMonitorResize))
        Call SubclassComboBox(cmbScrollWheelDirection.hWnd, ObjPtr(cmbScrollWheelDirection))
        Call SubclassComboBox(cmbWindowLevel.hWnd, ObjPtr(cmbWindowLevel))
        Call SubclassComboBox(cmbHidingTime.hWnd, ObjPtr(cmbHidingTime))
        
        Call SubclassComboBox(cmbWidgetLandscape.hWnd, ObjPtr(cmbWidgetLandscape))
        Call SubclassComboBox(cmbWidgetPortrait.hWnd, ObjPtr(cmbWidgetPortrait))
        Call SubclassComboBox(cmbWidgetPosition.hWnd, ObjPtr(cmbWidgetPosition))
        Call SubclassComboBox(cmbAspectHidden.hWnd, ObjPtr(cmbAspectHidden))
        Call SubclassComboBox(cmbDebug.hWnd, ObjPtr(cmbDebug))
        
        Call SubclassComboBox(cmbAlarm1Day.hWnd, ObjPtr(cmbAlarm1Day))
        Call SubclassComboBox(cmbAlarm2Day.hWnd, ObjPtr(cmbAlarm2Day))
        Call SubclassComboBox(cmbAlarm3Day.hWnd, ObjPtr(cmbAlarm3Day))
        Call SubclassComboBox(cmbAlarm4Day.hWnd, ObjPtr(cmbAlarm4Day))
        Call SubclassComboBox(cmbAlarm5Day.hWnd, ObjPtr(cmbAlarm5Day))
        
        Call SubclassComboBox(cmbAlarm1Month.hWnd, ObjPtr(cmbAlarm1Month))
        Call SubclassComboBox(cmbAlarm2Month.hWnd, ObjPtr(cmbAlarm2Month))
        Call SubclassComboBox(cmbAlarm3Month.hWnd, ObjPtr(cmbAlarm3Month))
        Call SubclassComboBox(cmbAlarm4Month.hWnd, ObjPtr(cmbAlarm4Month))
        Call SubclassComboBox(cmbAlarm5Month.hWnd, ObjPtr(cmbAlarm5Month))
        
        Call SubclassComboBox(cmbAlarm1Year.hWnd, ObjPtr(cmbAlarm1Year))
        Call SubclassComboBox(cmbAlarm2Year.hWnd, ObjPtr(cmbAlarm2Year))
        Call SubclassComboBox(cmbAlarm3Year.hWnd, ObjPtr(cmbAlarm3Year))
        Call SubclassComboBox(cmbAlarm4Year.hWnd, ObjPtr(cmbAlarm4Year))
        Call SubclassComboBox(cmbAlarm5Year.hWnd, ObjPtr(cmbAlarm5Year))
        
        Call SubclassComboBox(cmbAlarm1Hours.hWnd, ObjPtr(cmbAlarm1Hours))
        Call SubclassComboBox(cmbAlarm2Hours.hWnd, ObjPtr(cmbAlarm2Hours))
        Call SubclassComboBox(cmbAlarm3Hours.hWnd, ObjPtr(cmbAlarm3Hours))
        Call SubclassComboBox(cmbAlarm4Hours.hWnd, ObjPtr(cmbAlarm4Hours))
        Call SubclassComboBox(cmbAlarm5Hours.hWnd, ObjPtr(cmbAlarm5Hours))
        
        Call SubclassComboBox(cmbAlarm1Minutes.hWnd, ObjPtr(cmbAlarm1Minutes))
        Call SubclassComboBox(cmbAlarm2Minutes.hWnd, ObjPtr(cmbAlarm2Minutes))
        Call SubclassComboBox(cmbAlarm3Minutes.hWnd, ObjPtr(cmbAlarm3Minutes))
        Call SubclassComboBox(cmbAlarm4Minutes.hWnd, ObjPtr(cmbAlarm4Minutes))
        Call SubclassComboBox(cmbAlarm5Minutes.hWnd, ObjPtr(cmbAlarm5Minutes))
        
    End If

    On Error GoTo 0
    Exit Sub

subClassControls_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure subClassControls of Form rDIconConfigForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : MouseMoveOnComboText
' Author    : beededea
' Date      : 16/07/2024
' Purpose   : Note: Each control must also be added to the subClassControls routine
'---------------------------------------------------------------------------------------
'
Public Sub MouseMoveOnComboText(sComboName As String)
    Dim sTitle As String
    Dim sText As String

    On Error GoTo MouseMoveOnComboText_Error
    
    Select Case sComboName
        Case "cmbMultiMonitorResize"
            sTitle = "Help on the Drop Down Icon Filter"
            sText = gCmbMultiMonitorResizeBalloonTooltip
            If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip cmbMultiMonitorResize.hWnd, sText, , sTitle, , , , True
        Case "cmbScrollWheelDirection"
            sTitle = "Help on the Scroll Wheel Direction"
            sText = gCmbScrollWheelDirectionBalloonTooltip
            If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip cmbScrollWheelDirection.hWnd, sText, , sTitle, , , , True
        Case "cmbWindowLevel"
            sTitle = "Help on the Window Level"
            sText = gCmbWindowLevelBalloonTooltip
            If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip cmbWindowLevel.hWnd, sText, , sTitle, , , , True
        Case "cmbHidingTime"
            sTitle = "Help on the Hiding Time"
            sText = gCmbHidingTimeBalloonTooltip
            If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip cmbHidingTime.hWnd, sText, , sTitle, , , , True
            
        Case "cmbAspectHidden"
            sTitle = "Help on Hiding in Landscape/Portrait Mode"
            sText = gCmbAspectHiddenBalloonTooltip
            If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip cmbAspectHidden.hWnd, sText, , sTitle, , , , True
        Case "cmbWidgetPosition"
            sTitle = "Help on Widget Position in Landscape/Portrait Modes"
            sText = gCmbWidgetPositionBalloonTooltip
            If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip cmbWidgetPosition.hWnd, sText, , sTitle, , , , True
        Case "cmbWidgetLandscape"
            sTitle = "Help on Widget Locking in Landscape Mode"
            sText = gCmbWidgetLandscapeBalloonTooltip
            If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip cmbWidgetLandscape.hWnd, sText, , sTitle, , , , True
        Case "cmbWidgetPortrait"
            sTitle = "Help on Widget Locking in Portrait Mode"
            sText = gCmbWidgetPortraitBalloonTooltip
            If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip cmbWidgetPortrait.hWnd, sText, , sTitle, , , , True
        Case "cmbDebug"
            sTitle = "Help on Debug Mode"
            sText = gCmbDebugBalloonTooltip
            If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip cmbDebug.hWnd, sText, , sTitle, , , , True
        
        Case "cmbAlarm1Day"
            sTitle = "Help on Alarm Number One Day"
            sText = gCmbAlarmDayBalloonTooltip
            If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip cmbAlarm1Day.hWnd, sText, , sTitle, , , , True
        Case "cmbAlarm2Day"
            sTitle = "Help on Alarm Number Two Day"
            sText = gCmbAlarmDayBalloonTooltip
            If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip cmbAlarm2Day.hWnd, sText, , sTitle, , , , True
        Case "cmbAlarm3Day"
            sTitle = "Help on Alarm Number Three Day"
            sText = gCmbAlarmDayBalloonTooltip
            If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip cmbAlarm3Day.hWnd, sText, , sTitle, , , , True
        Case "cmbAlarm4Day"
            sTitle = "Help on Alarm Number Four Day"
            sText = gCmbAlarmDayBalloonTooltip
            If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip cmbAlarm4Day.hWnd, sText, , sTitle, , , , True
        Case "cmbAlarm5Day"
            sTitle = "Help on Alarm Number Five Day"
            sText = gCmbAlarmDayBalloonTooltip
            If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip cmbAlarm5Day.hWnd, sText, , sTitle, , , , True
        
        Case "cmbAlarm1Month"
            sTitle = "Help on Alarm Number One Month"
            sText = gCmbAlarmMonthBalloonTooltip
            If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip cmbAlarm1Month.hWnd, sText, , sTitle, , , , True
        Case "cmbAlarm2Month"
            sTitle = "Help on Alarm Number Two Month"
            sText = gCmbAlarmMonthBalloonTooltip
            If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip cmbAlarm2Month.hWnd, sText, , sTitle, , , , True
        Case "cmbAlarm3Month"
            sTitle = "Help on Alarm Number Three Month"
            sText = gCmbAlarmMonthBalloonTooltip
            If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip cmbAlarm3Month.hWnd, sText, , sTitle, , , , True
        Case "cmbAlarm4Month"
            sTitle = "Help on Alarm Number Four Month"
            sText = gCmbAlarmMonthBalloonTooltip
            If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip cmbAlarm4Month.hWnd, sText, , sTitle, , , , True
        Case "cmbAlarm5Month"
            sTitle = "Help on Alarm Number Five Month"
            sText = gCmbAlarmMonthBalloonTooltip
            If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip cmbAlarm5Month.hWnd, sText, , sTitle, , , , True
        
        Case "cmbAlarm1Year"
            sTitle = "Help on Alarm Number One Year"
            sText = gCmbAlarmYearBalloonTooltip
            If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip cmbAlarm1Year.hWnd, sText, , sTitle, , , , True
        Case "cmbAlarm2Year"
            sTitle = "Help on Alarm Number Two Year"
            sText = gCmbAlarmYearBalloonTooltip
            If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip cmbAlarm2Year.hWnd, sText, , sTitle, , , , True
        Case "cmbAlarm3Year"
            sTitle = "Help on Alarm Number Three Year"
            sText = gCmbAlarmYearBalloonTooltip
            If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip cmbAlarm3Year.hWnd, sText, , sTitle, , , , True
        Case "cmbAlarm4Year"
            sTitle = "Help on Alarm Number Four Year"
            sText = gCmbAlarmYearBalloonTooltip
            If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip cmbAlarm4Year.hWnd, sText, , sTitle, , , , True
        Case "cmbAlarm5Year"
            sTitle = "Help on Alarm Number Five Year"
            sText = gCmbAlarmYearBalloonTooltip
            If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip cmbAlarm5Year.hWnd, sText, , sTitle, , , , True
        
        Case "cmbAlarm1Hours"
            sTitle = "Help on Alarm Number One Hour"
            sText = gCmbAlarmHoursBalloonTooltip
            If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip cmbAlarm1Hours.hWnd, sText, , sTitle, , , , True
        Case "cmbAlarm2Hours"
            sTitle = "Help on Alarm Number Two Hour"
            sText = gCmbAlarmHoursBalloonTooltip
            If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip cmbAlarm2Hours.hWnd, sText, , sTitle, , , , True
        Case "cmbAlarm3Hours"
            sTitle = "Help on Alarm Number Three Hour"
            sText = gCmbAlarmHoursBalloonTooltip
            If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip cmbAlarm3Hours.hWnd, sText, , sTitle, , , , True
        Case "cmbAlarm4Hours"
            sTitle = "Help on Alarm Number Four Hour"
            sText = gCmbAlarmHoursBalloonTooltip
            If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip cmbAlarm4Hours.hWnd, sText, , sTitle, , , , True
        Case "cmbAlarm5Hours"
            sTitle = "Help on Alarm Number Five Hour"
            sText = gCmbAlarmHoursBalloonTooltip
            If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip cmbAlarm5Hours.hWnd, sText, , sTitle, , , , True
        
        Case "cmbAlarm1Minutes"
            sTitle = "Help on Alarm Number One Minute"
            sText = gCmbAlarmMinutesBalloonTooltip
            If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip cmbAlarm1Minutes.hWnd, sText, , sTitle, , , , True
        Case "cmbAlarm2Minutes"
            sTitle = "Help on Alarm Number Two Minute"
            sText = gCmbAlarmMinutesBalloonTooltip
            If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip cmbAlarm2Minutes.hWnd, sText, , sTitle, , , , True
        Case "cmbAlarm3Minutes"
            sTitle = "Help on Alarm Number Three Minute"
            sText = gCmbAlarmMinutesBalloonTooltip
            If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip cmbAlarm3Minutes.hWnd, sText, , sTitle, , , , True
        Case "cmbAlarm4Minutes"
            sTitle = "Help on Alarm Number Four Minute"
            sText = gCmbAlarmMinutesBalloonTooltip
            If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip cmbAlarm4Minutes.hWnd, sText, , sTitle, , , , True
        Case "cmbAlarm5Minutes"
            sTitle = "Help on Alarm Number Five Minute"
            sText = gCmbAlarmMinutesBalloonTooltip
            If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip cmbAlarm5Minutes.hWnd, sText, , sTitle, , , , True
    End Select
    
   On Error GoTo 0
   Exit Sub

MouseMoveOnComboText_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MouseMoveOnComboText of Form rDIconConfigForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Form_Moved
' Author    : beededea
' Date      : 16/07/2024
' Purpose   : Non VB6-standard event caught by subclassing and intercepting the WM_EXITSIZEMOVE (WM_MOVED) event
'---------------------------------------------------------------------------------------
'
Public Sub Form_Moved(sForm As String)
    Dim sTitle As String
    Dim sText As String

    On Error GoTo Form_Moved_Error
        
    'passing a form name as it allows us to potentially subclass another form's movement
    Select Case sForm
        Case "widgetPrefs"
            ' call a resize of all controls only when the form resize (by dragging) has completed (mouseUP)
            If gPrefsFormResizedByDrag = True Then
                Call PrefsForm_resize
                gPrefsFormResizedByDrag = False
                
            End If
            
            ' call the procedure to resize the form automatically if it now resides on a different sized monitor
            Call positionPrefsByMonitorSize
           
        Case Else
    End Select
    
   On Error GoTo 0
   Exit Sub

Form_Moved_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Moved of Form rDIconConfigForm"
End Sub



' ---------------------------------------------------------------------------------------
' Procedure : positionPrefsMonitor
' Author    : beededea
' Date      : 06/10/2023
' Purpose   : position the prefs on the curent monitor
'---------------------------------------------------------------------------------------
'
Public Sub positionPrefsMonitor()

    Dim formLeftTwips As Long: formLeftTwips = 0
    Dim formTopTwips As Long: formTopTwips = 0
    'Dim monitorCount As Long: monitorCount = 0
    
    On Error GoTo positionPrefsMonitor_Error
    
    If gblDpiAwareness = "1" Then
        formLeftTwips = Val(gblPrefsHighDpiXPosTwips)
        formTopTwips = Val(gblPrefsHighDpiYPosTwips)
    Else
        formLeftTwips = Val(gblPrefsLowDpiXPosTwips)
        formTopTwips = Val(gblPrefsLowDpiYPosTwips)
    End If
    
    If formLeftTwips = 0 Then
        If ((fClock.clockForm.Left + fClock.clockForm.Width) * screenTwipsPerPixelX) + 200 + widgetPrefs.Width > physicalScreenWidthTwips Then
            widgetPrefs.Left = (fClock.clockForm.Left * screenTwipsPerPixelX) - (widgetPrefs.Width + 200)
        End If
    End If

    ' if a current location not stored then position to the middle of the screen
    
    If formLeftTwips <> 0 Then
        widgetPrefs.Left = formLeftTwips
    Else
        widgetPrefs.Left = physicalScreenWidthTwips / 2 - widgetPrefs.Width / 2
    End If
    
    If formTopTwips <> 0 Then
        widgetPrefs.Top = formTopTwips
    Else
        widgetPrefs.Top = Screen.Height / 2 - widgetPrefs.Height / 2
    End If
    
    'monitorCount = fGetMonitorCount
    If monitorCount > 1 Then Call SetFormOnMonitor(Me.hWnd, formLeftTwips / fTwipsPerPixelX, formTopTwips / fTwipsPerPixelY)
    
    ' calculate the on-screen widget position
    If Me.Left < 0 Then
        widgetPrefs.Left = 10
    End If
    If Me.Top < 0 Then
        widgetPrefs.Top = 0
    End If
    If Me.Left > virtualScreenWidthTwips - 2500 Then
        widgetPrefs.Left = virtualScreenWidthTwips - 2500
    End If
    If Me.Top > virtualScreenHeightTwips - 2500 Then
        widgetPrefs.Top = virtualScreenHeightTwips - 2500
    End If
    
    
    ' if just one monitor or the global switch is off then exit
    If monitorCount > 1 And LTrim$(gblMultiMonitorResize) = "2" Then

        If prefsMonitorStruct.IsPrimary = True Then
            gblPrefsFormResizedInCode = True
            gblPrefsPrimaryHeightTwips = fGetINISetting("Software\SteampunkClockCalendar", "prefsPrimaryHeightTwips", gblSettingsFile)
            If Val(gblPrefsPrimaryHeightTwips) <= 0 Then
                widgetPrefs.Height = gblPrefsCurrentHeight
            Else
                widgetPrefs.Height = Val(gblPrefsPrimaryHeightTwips)
            End If
        Else
            gblPrefsSecondaryHeightTwips = fGetINISetting("Software\SteampunkClockCalendar", "prefsSecondaryHeightTwips", gblSettingsFile)
            gblPrefsFormResizedInCode = True
            If Val(gblPrefsSecondaryHeightTwips) <= 0 Then
                widgetPrefs.Height = gblPrefsCurrentHeight
            Else
                widgetPrefs.Height = Val(gblPrefsSecondaryHeightTwips)
            End If
        End If
    End If

    fClock.RotateBusyTimer = True
    
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

    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip btnResetMessages.hWnd, "The various pop-up messages that this program generates can be manually hidden. This button restores them to their original visible state.", _
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
' Procedure : optEnablePrefsTooltips_Click
' Author    : beededea
' Date      : 07/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub optEnablePrefsTooltips_Click()

   On Error GoTo optEnablePrefsTooltips_Click_Error

    btnSave.Enabled = True ' enable the save button
    
    If prefsStartupFlg = False Then
        If optEnablePrefsTooltips.Value = True Then
            gblEnablePrefsTooltips = "True"
        Else
            gblEnablePrefsTooltips = "False"
        End If

    End If
        
    sPutINISetting "Software\SteampunkClockCalendar", "enablePrefsTooltips", gblEnablePrefsTooltips, gblSettingsFile
    
    ' set the tooltips on the prefs screen
    Call setPrefsTooltips

   On Error GoTo 0
   Exit Sub

optEnablePrefsTooltips_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optEnablePrefsTooltips_Click of Form widgetPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : optEnableTooltips_MouseMove
' Author    : beededea
' Date      : 05/10/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub optEnableTooltips_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip optEnableTooltips.hWnd, "This setting enables/disables the RichClient-enhanced tooltips for elements within the Steampunk GUI. These tooltips allow for a larger font size on high DPI screens.", _
                  TTIconInfo, "Help on Enabling/Disabling RC tooltips on the GUI", , , , True
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
    
       ' If resizeProportionTimed > 0 Then widgetPrefs.Height = resizeProportionTimed

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
' Procedure : optEnableBalloonTooltips_Click
' Author    : beededea
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub optEnableBalloonTooltips_Click()
   On Error GoTo optEnableBalloonTooltips_Click_Error

    btnSave.Enabled = True ' enable the save button
    If optEnableBalloonTooltips.Value = True Then
        gblEnableBalloonTooltips = "True" 'true
        gblEnableTooltips = "False"
    Else
        gblEnableBalloonTooltips = "False"
        gblEnableTooltips = "True"
    End If
        
    sPutINISetting "Software\SteampunkClockCalendar", "enableBalloonTooltips", gblEnableBalloonTooltips, gblSettingsFile

   On Error GoTo 0
   Exit Sub

optEnableBalloonTooltips_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optEnableBalloonTooltips_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : optEnableTooltips_Click
' Author    : beededea
' Date      : 19/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub optEnableTooltips_Click()
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString
    On Error GoTo optEnableTooltips_Click_Error

    btnSave.Enabled = True ' enable the save button
    
    If prefsStartupFlg = False Then
        If optEnableTooltips.Value = True Then
            gblEnableTooltips = "True" 'true
            gblEnableBalloonTooltips = "False"
        Else
            gblEnableTooltips = "False"
            gblEnableBalloonTooltips = "True"
        End If
        
        sPutINISetting "Software\SteampunkClockCalendar", "enableTooltips", gblEnableTooltips, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "enableBalloonTooltips", gblEnableBalloonTooltips, gblSettingsFile

        answer = vbYes
        answerMsg = "You must soft reload this widget, in order to change the tooltip setting, do you want me to reload this widget? I can do it now for you."
        answer = msgBoxA(answerMsg, vbYesNo, "Request to Enable Tooltips", True, "optEnableTooltipsClick")
        If answer = vbNo Then
            Exit Sub
        Else
            Call reloadWidget
        End If
    End If


   On Error GoTo 0
   Exit Sub

optEnableTooltips_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optEnableTooltips_Click of Form widgetPrefs"

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
' Procedure : IsLoaded
' Author    : beededea
' Date      : 16/12/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get IsLoaded() As Boolean
 
   On Error GoTo IsLoaded_Error

    IsLoaded = mIsLoaded

   On Error GoTo 0
   Exit Property

IsLoaded_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsLoaded of Form widgetPrefs"
 
End Property

'---------------------------------------------------------------------------------------
' Procedure : IsLoaded
' Author    : beededea
' Date      : 16/12/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let IsLoaded(ByVal newValue As Boolean)
 
   On Error GoTo IsLoaded_Error

   If mIsLoaded <> newValue Then mIsLoaded = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

IsLoaded_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsLoaded of Form widgetPrefs"
 
End Property


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

    'If gblPrefsLoadedFlg = True Then
    If IsLoaded = True Then
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

    fClock.RotateBusyTimer = True
    
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
    widgetPrefs.Width = rightHandAlignment + leftHandGutterWidth + 75 ' (not quite sure why we need the 75 twips padding) ' this triggers a resize '
    
    ' 9053 start
    
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
    
    fClock.RotateBusyTimer = True

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
    
    Call adjustPrefsControls(True)

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
            Call ShellExecute(Me.hWnd, "Open", App.path & "\help\Help.chm", vbNullString, App.path, 1)
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
    gblEnableTooltips = CStr(optEnableTooltips.Value)
    gblEnableBalloonTooltips = CStr(optEnableBalloonTooltips.Value)
    
    gblEnablePrefsTooltips = CStr(optEnablePrefsTooltips.Value)
    gblEnablePrefsBalloonTooltips = CStr(optEnablePrefsBalloonTooltips.Value)
    
    gblShowTaskbar = CStr(chkShowTaskbar.Value)
    gblShowHelp = CStr(chkShowHelp.Value)
    
    gblTogglePendulum = CStr(chkTogglePendulum.Value)
    gbl24HourClockMode = CStr(chk24HourClockMode.Value)
    
    gblDpiAwareness = CStr(chkDpiAwareness.Value)
    gblGaugeSize = CStr(sliGaugeSize.Value)
    gblScrollWheelDirection = CStr(cmbScrollWheelDirection.ListIndex)
    
    ' general
    gblWidgetFunctions = CStr(chkWidgetFunctions.Value)
    gblStartup = CStr(chkGenStartup.Value)
    
    ' Validate all the alarm variables
    alarmTest = validateAlarmVars
    If alarmTest = False Then Exit Sub ' END the save if the alarms are malformed
    
    gblAlarm1 = gblAlarm1Date & " " & gblAlarm1Time
    gblAlarm2 = gblAlarm2Date & " " & gblAlarm2Time
    gblAlarm3 = gblAlarm3Date & " " & gblAlarm3Time
    gblAlarm4 = gblAlarm4Date & " " & gblAlarm4Time
    gblAlarm5 = gblAlarm5Date & " " & gblAlarm5Time
    
    ' sounds
    gblEnableSounds = CStr(chkEnableSounds.Value)
    gblEnableTicks = CStr(chkEnableTicks.Value)
    gblEnableChimes = CStr(chkEnableChimes.Value)
    gblVolumeBoost = CStr(chkVolumeBoost.Value)
    
    'development
    gblDebug = CStr(cmbDebug.ListIndex)
    gblDblClickCommand = txtDblClickCommand.Text
    gblOpenFile = txtOpenFile.Text
    gblDefaultEditor = txtDefaultEditor.Text
    
    ' position
    gblAspectHidden = CStr(cmbAspectHidden.ListIndex)
    gblWidgetPosition = CStr(cmbWidgetPosition.ListIndex)
    gblWidgetLandscape = CStr(cmbWidgetLandscape.ListIndex)
    gblWidgetPortrait = CStr(cmbWidgetPortrait.ListIndex)
    gblLandscapeFormHoffset = txtLandscapeHoffset.Text
    gblLandscapeFormVoffset = txtLandscapeVoffset.Text
    gblPortraitHoffset = txtPortraitHoffset.Text
    gblPortraitYoffset = txtPortraitYoffset.Text
    
'    gblvLocationPercPrefValue
'    gblhLocationPercPrefValue

    ' fonts
    gblPrefsFont = txtPrefsFont.Text
    gblClockFont = gblPrefsFont
    
    gblDisplayScreenFont = txtDisplayScreenFont.Text
    gblDisplayScreenFontSize = txtDisplayScreenFontSize.Text
    
'    gblDisplayScreenFontSize
'    gblDisplayScreenFontItalics
'    gblDisplayScreenFontColour

    ' the sizing is not saved here again as it saved during the setting phase.
    
'    If gblDpiAwareness = "1" Then
'        gblPrefsFontSizeHighDPI = txtPrefsFontSize.Text
'    Else
'        gblPrefsFontSizeLowDPI = txtPrefsFontSize.Text
'    End If
    'gblPrefsFontItalics = txtFontSize.Text

    ' Windows
    gblWindowLevel = CStr(cmbWindowLevel.ListIndex)
    gblPreventDragging = CStr(chkPreventDragging.Value)
    gblOpacity = CStr(sliOpacity.Value)
    gblWidgetHidden = CStr(chkWidgetHidden.Value)
    gblHidingTime = CStr(cmbHidingTime.ListIndex)
    gblIgnoreMouse = CStr(chkIgnoreMouse.Value)
    
    gblMultiMonitorResize = CStr(cmbMultiMonitorResize.ListIndex)
     
            
    
    'development
    gblDebug = CStr(cmbDebug.ListIndex)
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
        sPutINISetting "Software\SteampunkClockCalendar", "enablePrefsBalloonTooltips", gblEnablePrefsBalloonTooltips, gblSettingsFile

        sPutINISetting "Software\SteampunkClockCalendar", "showTaskbar", gblShowTaskbar, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "showHelp", gblShowHelp, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "togglePendulum", gblTogglePendulum, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "24HourClockMode", gbl24HourClockMode, gblSettingsFile
        
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
        
        sPutINISetting "Software\SteampunkClockCalendar", "displayScreenFont", gblDisplayScreenFont, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "displayScreenFontSize", gblDisplayScreenFontSize, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "displayScreenFontItalics", gblDisplayScreenFontItalics, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "displayScreenFontColour", gblDisplayScreenFontColour, gblSettingsFile

        'save the values from the Windows Config Items
        sPutINISetting "Software\SteampunkClockCalendar", "windowLevel", gblWindowLevel, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "preventDragging", gblPreventDragging, gblSettingsFile
        
        sPutINISetting "Software\SteampunkClockCalendar", "opacity", gblOpacity, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "widgetHidden", gblWidgetHidden, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "hidingTime", gblHidingTime, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "ignoreMouse", gblIgnoreMouse, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "multiMonitorResize", gblMultiMonitorResize, gblSettingsFile
        
        
        sPutINISetting "Software\SteampunkClockCalendar", "startup", gblStartup, gblSettingsFile

        sPutINISetting "Software\SteampunkClockCalendar", "enableSounds", gblEnableSounds, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "enableTicks", gblEnableTicks, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "enableChimes", gblEnableChimes, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "volumeBoost", gblVolumeBoost, gblSettingsFile
        
        sPutINISetting "Software\SteampunkClockCalendar", "lastSelectedTab", gblLastSelectedTab, gblSettingsFile
        
        sPutINISetting "Software\SteampunkClockCalendar", "debug", gblDebug, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "dblClickCommand", gblDblClickCommand, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "openFile", gblOpenFile, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "defaultEditor", gblDefaultEditor, gblSettingsFile
        
        sPutINISetting "Software\SteampunkClockCalendar", "clockHighDpiXPos", gblClockHighDpiXPos, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "clockHighDpiYPos", gblClockHighDpiYPos, gblSettingsFile
        
        sPutINISetting "Software\SteampunkClockCalendar", "clockLowDpiXPos", gblClockLowDpiXPos, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "clockLowDpiYPos", gblClockLowDpiYPos, gblSettingsFile
        
        'sPutINISetting "Software\SteampunkClockCalendar", "setToggleEnabled", gblSetToggleEnabled, gblSettingsFile
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
        'alarmDateStatus = fVerifyAlarmDate(txtAlarm1Date.Text)
        alarmTimeStatus = fVerifyAlarmTime(txtAlarm1Date.Text, txtAlarm1Time.Text)
        If alarmTimeStatus = False Then
            btnSave.Enabled = False
            
            answerMsg = "Alarm no.1 is invalid, saving FAILED. Please correct and re-save."
            answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Message", True, "validateAlarmVars1")
            
            validateAlarmVars = False
            Exit Function
        End If
    End If
    gblAlarm1Date = txtAlarm1Date.Text
    gblAlarm1Time = txtAlarm1Time.Text
    gblAlarm1 = gblAlarm1Date + " " & gblAlarm1Time
            
    If txtAlarm2Date.Text <> "Alarm not yet set" Then
        ' alarmDateStatusfVerifyAlarmDate(txtAlarm2Date.Text)
        alarmTimeStatus = fVerifyAlarmTime(txtAlarm2Date.Text, txtAlarm2Time.Text)
        If alarmTimeStatus = False Then
            btnSave.Enabled = False
            
            answerMsg = "Alarm no.2 is invalid, saving FAILED. Please correct and re-save."
            answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Message", True, "validateAlarmVars2")
            
            validateAlarmVars = False
            Exit Function
        End If
    End If
    gblAlarm2Date = txtAlarm2Date.Text
    gblAlarm2Time = txtAlarm2Time.Text
    gblAlarm2 = gblAlarm2Date + " " & gblAlarm2Time
    
    If txtAlarm3Date.Text <> "Alarm not yet set" Then
        ' alarmDateStatusfVerifyAlarmDate(txtAlarm3Date.Text)
        alarmTimeStatus = fVerifyAlarmTime(txtAlarm3Date.Text, txtAlarm3Time.Text)
        If alarmTimeStatus = False Then
            btnSave.Enabled = False
            
            answerMsg = "Alarm no.3 is invalid, saving FAILED. Please correct and re-save."
            answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Message", True, "validateAlarmVars3")
            
            validateAlarmVars = False
            Exit Function
        End If
    End If
    gblAlarm3Date = txtAlarm3Date.Text
    gblAlarm3Time = txtAlarm3Time.Text
    gblAlarm3 = gblAlarm3Date + " " & gblAlarm3Time
    
    If txtAlarm4Date.Text <> "Alarm not yet set" Then
        ' alarmDateStatusfVerifyAlarmDate(txtAlarm4Date.Text)
        alarmTimeStatus = fVerifyAlarmTime(txtAlarm4Date.Text, txtAlarm4Time.Text)
        If alarmTimeStatus = False Then
            btnSave.Enabled = False
            
            answerMsg = "Alarm no.4 is invalid, saving FAILED. Please correct and re-save."
            answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Message", True, "validateAlarmVars4")
            
            validateAlarmVars = False
            Exit Function
        End If
    End If
    gblAlarm4Date = txtAlarm4Date.Text
    gblAlarm4Time = txtAlarm4Time.Text
    gblAlarm4 = gblAlarm4Date + " " & gblAlarm4Time
            
    If txtAlarm5Date.Text <> "Alarm not yet set" Then
        ' alarmDateStatusfVerifyAlarmDate(txtAlarm5Date.Text)
        alarmTimeStatus = fVerifyAlarmTime(txtAlarm5Date.Text, txtAlarm5Time.Text)
        If alarmTimeStatus = False Then
            btnSave.Enabled = False
            
            answerMsg = "Alarm no.5 is invalid, saving FAILED. Please correct and re-save."
            answer = msgBoxA(answerMsg, vbExclamation + vbOK, "Alarm Message", True, "validateAlarmVars5")
            
            validateAlarmVars = False
            Exit Function
        End If
    End If
    gblAlarm5Date = txtAlarm5Date.Text
    gblAlarm5Time = txtAlarm5Time.Text
    gblAlarm5 = gblAlarm5Date + " " & gblAlarm5Time

   On Error GoTo 0
   Exit Function

validateAlarmVars_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure validateAlarmVars of Form widgetPrefs"
    
End Function



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
    
    ' changes the displayed font to an adjusted base font size after a resize
    Call PrefsForm_resize

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
' Procedure : btnDisplayScreenFont_Click
' Author    : beededea
' Date      : 02/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnDisplayScreenFont_Click()

    Dim fntFont As String: fntFont = vbNullString
    Dim fntSize As Integer: fntSize = 0
    Dim fntWeight As Integer: fntWeight = 0
    Dim fntStyle As Boolean: fntStyle = False
    Dim fntColour As Long: fntColour = 0
    Dim fntItalics As Boolean: fntItalics = False
    Dim fntUnderline As Boolean: fntUnderline = False
    Dim fntFontResult As Boolean: fntFontResult = False
    
    On Error GoTo btnDisplayScreenFont_Click_Error

    btnSave.Enabled = True ' enable the save button
    
    ' set the preliminary vars to feed and populate the changefont routine
    fntFont = gblDisplayScreenFont
    
    fntSize = Val(gblDisplayScreenFontSize)
    If fntSize = 0 Then fntSize = 8
    fntItalics = CBool(gblDisplayScreenFontItalics)
    fntColour = CLng(gblDisplayScreenFontColour)
    
    displayFontSelector fntFont, fntSize, fntWeight, fntStyle, fntColour, fntItalics, fntUnderline, fntFontResult
    If fntFontResult = False Then Exit Sub
            
    gblDisplayScreenFont = CStr(fntFont)
    gblDisplayScreenFontSize = CStr(fntSize)
    gblDisplayScreenFontItalics = CStr(fntItalics)
    gblDisplayScreenFontColour = CStr(fntColour)
    
    If gblFClockAvailable = True Then
        fClock.clockForm.Widgets("lblTerminalText").Widget.FontSize = gblDisplayScreenFontSize
        fClock.clockForm.Widgets("lblTerminalText").Widget.FontName = gblDisplayScreenFont
    End If

    If fFExists(gblSettingsFile) Then ' does the tool's own settings.ini exist?
        sPutINISetting "Software\SteampunkClockCalendar", "displayScreenFont", gblDisplayScreenFont, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "displayScreenFontSize", gblDisplayScreenFontSize, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "displayScreenFontItalics", gblDisplayScreenFontItalics, gblSettingsFile
        sPutINISetting "Software\SteampunkClockCalendar", "displayScreenFontColour", gblDisplayScreenFontColour, gblSettingsFile
    End If
    
    If fntFont = vbNullString Then fntFont = "courier new"
    txtDisplayScreenFont.Text = fntFont
    txtDisplayScreenFont.Font.Name = fntFont
    txtDisplayScreenFont.Font.Italic = fntItalics
    txtDisplayScreenFont.ForeColor = fntColour
    txtDisplayScreenFontSize.Text = fntSize

   On Error GoTo 0
   Exit Sub

btnDisplayScreenFont_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnDisplayScreenFont_Click of Form widgetPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : adjustPrefsControls
' Author    : beededea
' Date      : 12/05/2020
' Purpose   : adjust the controls so their startup position matches the last write of the config file
'---------------------------------------------------------------------------------------
'
Private Sub adjustPrefsControls(Optional ByVal restart As Boolean)
    
    Dim fntWeight As Integer: fntWeight = 0
    Dim fntStyle As Boolean: fntStyle = False
    Dim sliGaugeSizeOldValue As Long: sliGaugeSizeOldValue = 0
    Dim prefsFormMonitorID As Long: prefsFormMonitorID = 0
    
    On Error GoTo adjustPrefsControls_Error
    
    ' note the monitor ID at PrefsForm form_load and store as the prefsFormMonitorID
    'prefsMonitorStruct = formScreenProperties(widgetPrefs, prefsFormMonitorID)
            
    ' general tab
    chkWidgetFunctions.Value = Val(gblWidgetFunctions)
    chkGenStartup.Value = Val(gblStartup)
        
    txtAlarm1Date.Text = gblAlarm1Date
    txtAlarm2Date.Text = gblAlarm2Date
    txtAlarm3Date.Text = gblAlarm3Date
    txtAlarm4Date.Text = gblAlarm4Date
    txtAlarm5Date.Text = gblAlarm5Date
        
    txtAlarm1Time.Text = gblAlarm1Time
    txtAlarm2Time.Text = gblAlarm2Time
    txtAlarm3Time.Text = gblAlarm3Time
    txtAlarm4Time.Text = gblAlarm4Time
    txtAlarm5Time.Text = gblAlarm5Time
    
    ' fill the comboboxes for the alarms
        
    If gblAlarm1Date <> "Alarm not yet set" Then
        Call setAlarmDayValues(cmbAlarm1Day, gblAlarm1Date)
        Call setAlarmMonthValues(cmbAlarm1Month, gblAlarm1Date)
        Call setAlarmYearValues(cmbAlarm1Year, gblAlarm1Date)
        Call setAlarmHourValues(cmbAlarm1Hours, gblAlarm1Time)
        Call setAlarmMinuteValues(cmbAlarm1Minutes, gblAlarm1Time)
   End If
    If gblAlarm2Date <> "Alarm not yet set" Then
        Call setAlarmDayValues(cmbAlarm2Day, gblAlarm2Date)
        Call setAlarmMonthValues(cmbAlarm2Month, gblAlarm2Date)
        Call setAlarmYearValues(cmbAlarm2Year, gblAlarm2Date)
        Call setAlarmHourValues(cmbAlarm2Hours, gblAlarm2Time)
        Call setAlarmMinuteValues(cmbAlarm2Minutes, gblAlarm2Time)
    End If
    If gblAlarm3Date <> "Alarm not yet set" Then
        Call setAlarmDayValues(cmbAlarm3Day, gblAlarm3Date)
        Call setAlarmMonthValues(cmbAlarm3Month, gblAlarm3Date)
        Call setAlarmYearValues(cmbAlarm3Year, gblAlarm3Date)
        Call setAlarmHourValues(cmbAlarm3Hours, gblAlarm3Time)
        Call setAlarmMinuteValues(cmbAlarm3Minutes, gblAlarm3Time)
    End If
    If gblAlarm4Date <> "Alarm not yet set" Then
        Call setAlarmDayValues(cmbAlarm4Day, gblAlarm4Date)
        Call setAlarmMonthValues(cmbAlarm4Month, gblAlarm4Date)
        Call setAlarmYearValues(cmbAlarm4Year, gblAlarm4Date)
        Call setAlarmHourValues(cmbAlarm4Hours, gblAlarm4Time)
        Call setAlarmMinuteValues(cmbAlarm4Minutes, gblAlarm4Time)
    End If
    If gblAlarm5Date <> "Alarm not yet set" Then
        Call setAlarmDayValues(cmbAlarm5Day, gblAlarm5Date)
        Call setAlarmMonthValues(cmbAlarm5Month, gblAlarm5Date)
        Call setAlarmYearValues(cmbAlarm5Year, gblAlarm5Date)
        Call setAlarmHourValues(cmbAlarm5Hours, gblAlarm5Time)
        Call setAlarmMinuteValues(cmbAlarm5Minutes, gblAlarm5Time)
    End If
    
    ' configuration tab
   
    fClock.RotateBusyTimer = True
    
    ' check whether the size has been previously altered via ctrl+mousewheel on the widget
    sliGaugeSizeOldValue = sliGaugeSize.Value
    sliGaugeSize.Value = Val(gblGaugeSize)
    If sliGaugeSize.Value <> sliGaugeSizeOldValue Then
        btnSave.Visible = True
    End If
        
    cmbScrollWheelDirection.ListIndex = Val(gblScrollWheelDirection)
    
    optEnableTooltips.Value = CBool(gblEnableTooltips)
    optEnableBalloonTooltips.Value = CBool(gblEnableBalloonTooltips)
    
    optEnablePrefsTooltips.Value = CBool(gblEnablePrefsTooltips)
    optEnablePrefsBalloonTooltips.Value = CBool(gblEnablePrefsBalloonTooltips)
    
    chkShowTaskbar.Value = Val(gblShowTaskbar)
    chkShowHelp.Value = Val(gblShowHelp)
    chkTogglePendulum.Value = Val(gblTogglePendulum)
    chk24HourClockMode.Value = Val(gbl24HourClockMode)
    
    chkDpiAwareness.Value = Val(gblDpiAwareness)
    
    
    ' sounds tab
    fClock.RotateBusyTimer = True

    chkEnableSounds.Value = Val(gblEnableSounds)
    chkEnableTicks.Value = Val(gblEnableTicks)
    chkEnableChimes.Value = Val(gblEnableChimes)
    chkVolumeBoost.Value = Val(gblVolumeBoost)
    
    
    ' development
    fClock.RotateBusyTimer = True
    
    cmbDebug.ListIndex = Val(gblDebug)
    txtDblClickCommand.Text = gblDblClickCommand
    txtOpenFile.Text = gblOpenFile
    txtDefaultEditor.Text = gblDefaultEditor
    lblGitHub.Caption = "You can find the code for the Steampunk Clock Calendar on github, visit by double-clicking this link https://github.com/yereverluvinunclebert/ Steampunk-Clock-Calendar"
     
     
     If Not restart = True Then
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
        
        txtDisplayScreenFontSize.Text = gblDisplayScreenFontSize
    
        txtDisplayScreenFont.Font.Name = gblDisplayScreenFont
        'txtDisplayScreenFont.Font.Size = Val(gblDisplayScreenFont)
    End If
    
    ' position tab
    fClock.RotateBusyTimer = True
    
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
    fClock.RotateBusyTimer = True
    
    cmbWidgetLandscape.ListIndex = Val(gblWidgetLandscape)
    cmbWidgetPortrait.ListIndex = Val(gblWidgetPortrait)
    txtLandscapeHoffset.Text = gblLandscapeFormHoffset
    txtLandscapeVoffset.Text = gblLandscapeFormVoffset
    txtPortraitHoffset.Text = gblPortraitHoffset
    txtPortraitYoffset.Text = gblPortraitYoffset

    ' Windows tab
    fClock.RotateBusyTimer = True
    
    cmbWindowLevel.ListIndex = Val(gblWindowLevel)
    chkIgnoreMouse.Value = Val(gblIgnoreMouse)
    chkPreventDragging.Value = Val(gblPreventDragging)
    sliOpacity.Value = Val(gblOpacity)
    chkWidgetHidden.Value = Val(gblWidgetHidden)
    cmbHidingTime.ListIndex = Val(gblHidingTime)
    cmbMultiMonitorResize.ListIndex = Val(gblMultiMonitorResize)
    
    If monitorCount > 1 Then
        cmbMultiMonitorResize.Visible = True
        lblWindowLevel(10).Visible = True
        lblWindowLevel(11).Visible = True
    Else
        cmbMultiMonitorResize.Visible = False
        lblWindowLevel(10).Visible = False
        lblWindowLevel(11).Visible = False
    End If
    
    fClock.RotateBusyTimer = True
    
   On Error GoTo 0
   Exit Sub

adjustPrefsControls_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure adjustPrefsControls of Form widgetPrefs on line " & Erl

End Sub

'---------------------------------------------------------------------------------------
' Procedure : setAlarmHourValues
' Author    : beededea
' Date      : 01/11/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setAlarmHourValues(ByRef thisComboBox As ComboBox, ByVal thisDateString As String)
    Dim useloop As Integer: useloop = 0
    Dim thisHour As String: thisHour = vbNullString
    
    On Error GoTo setAlarmHourValues_Error

    thisHour = Left$(thisDateString, 2)

    'Iterate through items.
    For useloop = 0 To thisComboBox.ListCount - 1
        'Compare value.
        If thisComboBox.List(useloop) = thisHour Then
            'Select it and leave loop.
            thisComboBox.ListIndex = useloop
            Exit For
        End If
    Next useloop

   On Error GoTo 0
   Exit Sub

setAlarmHourValues_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setAlarmHourValues of Form widgetPrefs"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : setAlarmMinuteValues
' Author    : beededea
' Date      : 01/11/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setAlarmMinuteValues(ByRef thisComboBox As ComboBox, ByVal thisDateString As String)
    Dim useloop As Integer: useloop = 0
    Dim thisMinutes As String: thisMinutes = vbNullString
    
    On Error GoTo setAlarmMinuteValues_Error

    thisMinutes = Right$(thisDateString, 2)

    'Iterate through items.
    For useloop = 0 To thisComboBox.ListCount - 1
        'Compare value.
        If thisComboBox.List(useloop) = thisMinutes Then
            'Select it and leave loop.
            thisComboBox.ListIndex = useloop
            Exit For
        End If
    Next useloop

   On Error GoTo 0
   Exit Sub

setAlarmMinuteValues_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setAlarmMinuteValues of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : setAlarmYearValues
' Author    : beededea
' Date      : 01/11/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setAlarmYearValues(ByRef thisComboBox As ComboBox, ByVal thisDateString As String)
    Dim useloop As Integer: useloop = 0
    Dim thisYear As String: thisYear = vbNullString
    
    On Error GoTo setAlarmYearValues_Error

    thisYear = Right$(thisDateString, 4)

    'Iterate through items.
    For useloop = 0 To thisComboBox.ListCount - 1
        'Compare value.
        If thisComboBox.List(useloop) = thisYear Then
            'Select it and leave loop.
            thisComboBox.ListIndex = useloop
            Exit For
        End If
    Next useloop

   On Error GoTo 0
   Exit Sub

setAlarmYearValues_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setAlarmYearValues of Form widgetPrefs"

End Sub
    
'---------------------------------------------------------------------------------------
' Procedure : setAlarmDayValues
' Author    : beededea
' Date      : 01/11/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setAlarmDayValues(ByRef thisComboBox As ComboBox, ByVal thisDateString As String)
    Dim useloop As Integer: useloop = 0
    Dim thisDay As String: thisDay = vbNullString
    
   On Error GoTo setAlarmDayValues_Error

    thisDay = Left$(thisDateString, 2)

    'Iterate through items.
    For useloop = 0 To thisComboBox.ListCount - 1
        'Compare value.
        If thisComboBox.List(useloop) = thisDay Then
            'Select it and leave loop.
            thisComboBox.ListIndex = useloop
            Exit For
        End If
    Next useloop

   On Error GoTo 0
   Exit Sub

setAlarmDayValues_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setAlarmDayValues of Form widgetPrefs"

End Sub




    
'---------------------------------------------------------------------------------------
' Procedure : setAlarmMonthValues
' Author    : beededea
' Date      : 01/11/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setAlarmMonthValues(ByRef thisComboBox As ComboBox, ByVal thisDateString As String)
    Dim useloop As Integer: useloop = 0
    Dim monthNo As String: monthNo = vbNullString
    Dim newMonth As String: newMonth = vbNullString
    Dim thisMonth As String: thisMonth = vbNullString
    
    On Error GoTo setAlarmMonthValues_Error
    
    monthNo = CStr(DatePart("m", thisDateString)) ' extract the month number integer from the date string
    newMonth = MonthName(monthNo)  ' convert to full month name
    thisMonth = Left$(newMonth, 3) ' extract the first three characters

    ' iterate through combobox list
    For useloop = 0 To thisComboBox.ListCount - 1
        ' compare values
        If thisComboBox.List(useloop) = thisMonth Then
            ' select the value in to the combobox and leave.
            thisComboBox.ListIndex = useloop
            Exit For
        End If
    Next useloop

   On Error GoTo 0
   Exit Sub

setAlarmMonthValues_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setAlarmMonthValues of Form widgetPrefs"

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

    Dim useloop As Integer: useloop = 0
    Dim minString As String: minString = vbNullString
'    Dim hrsString As String: hrsString = vbNullString
'    Dim dayString As String: dayString = vbNullString
    
    On Error GoTo populatePrefsComboBoxes_Error
    
    With fClock.clockForm.Widgets("busy1").Widget
        .Alpha = 1
        .Refresh
    End With
    
    cmbScrollWheelDirection.AddItem "up", 0
    cmbScrollWheelDirection.ItemData(0) = 0
    cmbScrollWheelDirection.AddItem "down", 1
    cmbScrollWheelDirection.ItemData(1) = 1
    
    fClock.RotateBusyTimer = True
    
    cmbAspectHidden.AddItem "none", 0
    cmbAspectHidden.ItemData(0) = 0
    cmbAspectHidden.AddItem "portrait", 1
    cmbAspectHidden.ItemData(1) = 1
    cmbAspectHidden.AddItem "landscape", 2
    cmbAspectHidden.ItemData(2) = 2

    fClock.RotateBusyTimer = True
    
    cmbWidgetPosition.AddItem "disabled", 0
    cmbWidgetPosition.ItemData(0) = 0
    cmbWidgetPosition.AddItem "enabled", 1
    cmbWidgetPosition.ItemData(1) = 1
    
    fClock.RotateBusyTimer = True
    
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
    
    fClock.RotateBusyTimer = True
    
    ' populate comboboxes in the windows tab
    cmbWindowLevel.AddItem "Keep on top of other windows", 0
    cmbWindowLevel.ItemData(0) = 0
    cmbWindowLevel.AddItem "Normal", 0
    cmbWindowLevel.ItemData(1) = 1
    cmbWindowLevel.AddItem "Keep below all other windows", 0
    cmbWindowLevel.ItemData(2) = 2

    fClock.RotateBusyTimer = True
    
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
    
    fClock.RotateBusyTimer = True
    
    ' populate the multi monitor combobox
    cmbMultiMonitorResize.AddItem "Disabled", 0
    cmbMultiMonitorResize.ItemData(0) = 0
    cmbMultiMonitorResize.AddItem "Automatic Resizing Enabled", 1
    cmbMultiMonitorResize.ItemData(1) = 1
    cmbMultiMonitorResize.AddItem "Manual Sizing Stored Per Monitor", 2
    cmbMultiMonitorResize.ItemData(2) = 2

    fClock.RotateBusyTimer = True

    
    ' add the day options to the day combobox
    Call fillComboAlarmDay(cmbAlarm1Day)
    Call fillComboAlarmDay(cmbAlarm2Day)
    Call fillComboAlarmDay(cmbAlarm3Day)
    Call fillComboAlarmDay(cmbAlarm4Day)
    Call fillComboAlarmDay(cmbAlarm5Day)
    
    ' add the month options to the month combobox
    Call fillComboAlarmMonth(cmbAlarm1Month)
    Call fillComboAlarmMonth(cmbAlarm2Month)
    Call fillComboAlarmMonth(cmbAlarm3Month)
    Call fillComboAlarmMonth(cmbAlarm4Month)
    Call fillComboAlarmMonth(cmbAlarm5Month)
        
    ' add the next hundred year options to the year combobox
    Call fillComboAlarmYear(cmbAlarm1Year)
    Call fillComboAlarmYear(cmbAlarm2Year)
    Call fillComboAlarmYear(cmbAlarm3Year)
    Call fillComboAlarmYear(cmbAlarm4Year)
    Call fillComboAlarmYear(cmbAlarm5Year)
    
    ' add the hour options to the hour combobox
    Call fillComboAlarmHour(cmbAlarm1Hours)
    Call fillComboAlarmHour(cmbAlarm2Hours)
    Call fillComboAlarmHour(cmbAlarm3Hours)
    Call fillComboAlarmHour(cmbAlarm4Hours)
    Call fillComboAlarmHour(cmbAlarm5Hours)
    
    ' add the minute options to the minute combobox
    Call fillComboAlarmMinute(cmbAlarm1Minutes)
    Call fillComboAlarmMinute(cmbAlarm2Minutes)
    Call fillComboAlarmMinute(cmbAlarm3Minutes)
    Call fillComboAlarmMinute(cmbAlarm4Minutes)
    Call fillComboAlarmMinute(cmbAlarm5Minutes)
    
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
' Procedure : fillComboAlarmMinute
' Author    : beededea
' Date      : 31/10/2024
' Purpose   : add the minute options to the minute combobox
'---------------------------------------------------------------------------------------
'
Private Sub fillComboAlarmMinute(ByRef thisComboBox As ComboBox)
    Dim useloop As Integer: useloop = 0
    Dim minString As String: minString = vbNullString
    
    On Error GoTo fillComboAlarmMinute_Error
    
    thisComboBox.AddItem "", 0
    thisComboBox.ItemData(0) = 0
    
    For useloop = 0 To 59
        minString = CStr(useloop)
        If useloop <= 9 Then
            minString = "0" & CStr(useloop)
        End If
        thisComboBox.AddItem minString, useloop + 1
        thisComboBox.ItemData(useloop + 1) = useloop
        If useloop Mod (5) = 0 Then fClock.RotateBusyTimer = True
    Next useloop
   
    fClock.RotateBusyTimer = True
       
   On Error GoTo 0
   Exit Sub

fillComboAlarmMinute_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fillComboAlarmMinute of Form widgetPrefs"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : fillComboAlarmHour
' Author    : beededea
' Date      : 31/10/2024
' Purpose   : add the hour options to the hour combobox
'---------------------------------------------------------------------------------------
'
Private Sub fillComboAlarmHour(ByRef thisComboBox As ComboBox)
    Dim useloop As Integer: useloop = 0
    Dim hrsString As String: hrsString = vbNullString
    
    On Error GoTo fillComboAlarmHour_Error
    
    thisComboBox.AddItem "", 0
    thisComboBox.ItemData(0) = 0
        
    For useloop = 0 To 23
        hrsString = CStr(useloop)
        If useloop <= 9 Then
            hrsString = "0" & CStr(useloop)
        End If
        thisComboBox.AddItem hrsString, useloop + 1
        thisComboBox.ItemData(useloop + 1) = useloop
    Next useloop

    fClock.RotateBusyTimer = True
   
   On Error GoTo 0
   Exit Sub

fillComboAlarmHour_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fillComboAlarmHour of Form widgetPrefs"

End Sub
   

    
'---------------------------------------------------------------------------------------
' Procedure : fillComboAlarmYear
' Author    : beededea
' Date      : 31/10/2024
' Purpose   :  add the next hundred year options to the year combobox
'---------------------------------------------------------------------------------------
'
Private Sub fillComboAlarmYear(ByRef thisComboBox As ComboBox)
    Dim useloop As Integer: useloop = 0

    On Error GoTo fillComboAlarmYear_Error

    thisComboBox.AddItem "", 0
    thisComboBox.ItemData(0) = 0
        
    For useloop = 1 To 100
        thisComboBox.AddItem CStr(Year(Now) + useloop - 1), useloop
        thisComboBox.ItemData(useloop) = Year(Now) + useloop - 1
        If useloop Mod (10) = 0 Then fClock.RotateBusyTimer = True
    Next useloop

    fClock.RotateBusyTimer = True
       
  On Error GoTo 0
   Exit Sub

fillComboAlarmYear_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fillComboAlarmYear of Form widgetPrefs"
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : fillComboAlarmDay
' Author    : beededea
' Date      : 30/10/2024
' Purpose   : add the day options to the day combobox
'---------------------------------------------------------------------------------------
'
Private Sub fillComboAlarmDay(ByRef thisComboBox As ComboBox)
    Dim useloop As Integer: useloop = 0
    Dim dayString As String: dayString = vbNullString
    
    On Error GoTo fillComboAlarmDay_Error
    
    thisComboBox.AddItem "", 0
    thisComboBox.ItemData(0) = 0

    For useloop = 1 To 31
        dayString = CStr(useloop)
        If useloop <= 9 Then
            dayString = "0" & CStr(useloop)
        End If
        thisComboBox.AddItem dayString, useloop
        thisComboBox.ItemData(useloop) = useloop
        
        If useloop Mod (5) = 0 Then fClock.RotateBusyTimer = True
    Next useloop
   
    fClock.RotateBusyTimer = True
       
   On Error GoTo 0
   Exit Sub

fillComboAlarmDay_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fillComboAlarmDay of Form widgetPrefs"
End Sub

    
'---------------------------------------------------------------------------------------
' Procedure : fillComboAlarmMonth
' Author    : beededea
' Date      : 31/10/2024
' Purpose   : add the month options to the month combobox
'---------------------------------------------------------------------------------------
'
Private Sub fillComboAlarmMonth(ByRef thisComboBox As ComboBox)
    Dim useloop As Integer: useloop = 0

    On Error GoTo fillComboAlarmMonth_Error

    thisComboBox.AddItem "", 0
    thisComboBox.ItemData(0) = 0
    
    For useloop = 1 To 12
        thisComboBox.AddItem Left$(MonthName(useloop), 3), useloop
        thisComboBox.ItemData(useloop) = useloop
    Next useloop

    fClock.RotateBusyTimer = True
   
   On Error GoTo 0
   Exit Sub

fillComboAlarmMonth_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fillComboAlarmMonth of Form widgetPrefs"
    
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
' Purpose   : Called at every twip of resising, goodness knows what interval, we barely use this, instead we subclass and look for WM_EXITSIZEMOVE
'---------------------------------------------------------------------------------------
'
Private Sub Form_Resize()
    
    gPrefsFormResizedByDrag = True
    
    ' do not call the resizing function when the form is resized by dragging the border
    ' only call this if the resize is done in code
    If gblPrefsFormResizedInCode = True Then Call PrefsForm_resize
            
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
' Procedure : PrefsForm_resize
' Author    : beededea
' Date      : 10/10/2024
' Purpose   : routine to resize the prefs called by subclassing om a mouseUp form event during a bottom-right corner-drag or in code as required.
'---------------------------------------------------------------------------------------
'
Public Sub PrefsForm_resize()

    Dim ratio As Double: ratio = 0
    Dim currentFontSize As Long: currentFontSize = 0
    
    On Error GoTo PrefsForm_resize_Error

    ' When minimised and a resize is called then simply exit.
    If Me.WindowState = vbMinimized Then Exit Sub
    
    btnSave.Enabled = True ' enable the save button
    
    ' move the drag corner label along with the form's bottom right corner
    lblDragCorner.Move Me.ScaleLeft + Me.ScaleWidth - (lblDragCorner.Width + 40), _
               Me.ScaleTop + Me.ScaleHeight - (lblDragCorner.Height + 40)

    ' constrain the height/width ratio
    ratio = cPrefsFormHeight / cPrefsFormWidth
    
    If prefsDynamicSizingFlg = True Then
    
        widgetPrefs.Width = widgetPrefs.Height / ratio ' maintain the aspect ratio, note: this change calls this routine again...
        
        If gblDpiAwareness = "1" Then
            currentFontSize = gblPrefsFontSizeHighDPI
        Else
            currentFontSize = gblPrefsFontSizeLowDPI
        End If
        Call resizeControls(Me, prefsControlPositions(), gblPrefsCurrentWidth, gblPrefsCurrentHeight, currentFontSize)
        Call tweakPrefsControlPositions(Me, gblPrefsCurrentWidth, gblPrefsCurrentHeight)
        
        Call loadHigherResPrefsImages

    Else
        If Me.WindowState = 0 Then ' normal
            If widgetPrefs.Width > 9090 Then widgetPrefs.Width = 9090
            If widgetPrefs.Width < 9085 Then widgetPrefs.Width = 9090
            If lastFormHeight <> 0 Then
               gblPrefsFormResizedInCode = True
               widgetPrefs.Height = lastFormHeight
            End If
        End If
    End If
    
    gblPrefsFormResizedInCode = False
    
    'lblSize.Caption = "topIconWidth = " & topIconWidth & " imgGeneral width = " & imgGeneral.Width

   On Error GoTo 0
   Exit Sub

PrefsForm_resize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PrefsForm_resize of Form widgetPrefs"

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
    
    btnHelp.Top = fraGeneral.Top + fraGeneral.Height + (100 * y_scale)
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

    'gblPrefsLoadedFlg = False
    
    IsLoaded = False
    
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
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip fraAbout.hWnd, "The About tab tells you all about this program and its creation using " & gblCodingEnvironment & ".", _
                  TTIconInfo, "Help on the About Tab", , , , True
End Sub
Private Sub fraConfigInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraConfigInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip fraConfigInner.hWnd, "The configuration panel is the location for optional configuration items. These items change how the widget operates, configure them to suit your needs and your mode of operation.", _
                  TTIconInfo, "Help on Configuration", , , , True
End Sub
Private Sub fraConfig_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraConfig_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip fraConfig.hWnd, "The configuration panel is the location for important configuration items. These items change how the widget operates, configure them to suit your needs and your mode of operation.", _
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
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip fraDevelopment.hWnd, "This tab contains elements that will assist in debugging and developing this program further. ", _
                  TTIconInfo, "Help on the Development Tab", , , , True
End Sub
Private Sub fraDevelopmentInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraDevelopmentInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip fraDevelopmentInner.hWnd, "This tab contains elements that will assist in debugging and developing this program further. ", _
                  TTIconInfo, "Help on the Development Tab", , , , True

End Sub
Private Sub fraFonts_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip fraFonts.hWnd, "This tab allows you to set a specific font for the preferences only as there are no textual elements in the main program. We suggest Centurion Light SF at 8pt, which you will find bundled in the gbl program folder. Choose a small 8pt font for each.", _
                  TTIconInfo, "Help on Setting the Fonts", , , , True

End Sub
Private Sub fraFontsInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraFontsInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip fraFontsInner.hWnd, "This tab allows you to set a specific font for the preferences only as there are no textual elements in the main program. We suggest Centurion Light SF at 8pt, which you will find bundled in the gbl program folder. Choose a small 8pt font for each.", _
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
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip fraGeneral.hWnd, "The General Panel contains the most important user-configurable items required for the program to operate correctly.", _
                  TTIconInfo, "Help on Essential Configuration", , , , True
End Sub

Private Sub fraGeneralInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraGeneralInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip fraGeneralInner.hWnd, "The General Panel contains the most important user-configurable items required for the program to operate correctly.", _
                  TTIconInfo, "Help on Essential Configuration", , , , True
End Sub

Private Sub fraPosition_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip fraPosition.hWnd, "This tab allows you to determine the X and Y positioning of your widget in landscape and portrait screen modes. Best left well alone unless you use Windows on a tablet.", _
                  TTIconInfo, "Help on Tablet Positioning", , , , True
End Sub
Private Sub fraPositionInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraPositionInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip fraPositionInner.hWnd, "This tab allows you to determine the X and Y positioning of your widget in landscape and portrait screen modes. Best left well alone unless you use Windows on a tablet.", _
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
 If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip fraSounds.hWnd, "The sound panel allows you to configure the sounds that occur within gbl. Some of the animations have associated sounds, you can control these here..", _
                  TTIconInfo, "Help on Configuring Sounds", , , , True
End Sub
Private Sub fraSoundsInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraSoundsInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip fraSoundsInner.hWnd, "The sound panel allows you to configure the sounds that occur within gbl. Some of the animations have associated sounds, you can control these here..", _
                  TTIconInfo, "Help on Configuring Sounds", , , , True
End Sub

Private Sub fraWindow_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraWindow_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip fraWindow.hWnd, "The Opacity and Window Level of the program are rather strange characteristics to change in a Windows program, however this widget is a copy of a Yahoo Widget of the same name. All widgets have similar window tab options including the capability to change the opacity and window level. Whether these options are useful to you or anyone is a moot point but as this tool aims to replicate the YWE version functionality it has been reproduced here. It is here as more of an experiment as to how to implement a feature, one carried over from the Yahoo Widget (javascript) version of this program.", _
                  TTIconInfo, "Help on YWE Quirk Mode Options", , , , True
End Sub
Private Sub fraWindowInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraWindowInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip fraWindowInner.hWnd, "The Opacity and Window Level of the program are rather strange characteristics to change in a Windows program, however this widget is a copy of a Yahoo Widget of the same name. All widgets have similar window tab options including the capability to change the opacity and window level. Whether these options are useful to you or anyone is a moot point but as this tool aims to replicate the YWE version functionality it has been reproduced here. It is here as more of an experiment as to how to implement a feature, one carried over from the Yahoo Widget (javascript) version of this program.", _
                  TTIconInfo, "Help on YWE Quirk Mode Options", , , , True
End Sub

   Private Sub imgGeneral_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
        #If TWINBASIC Then
            Exit Sub
        #Else
            imgGeneral.Visible = False
            imgGeneralClicked.Visible = True
        #End If
    End Sub
    Private Sub imgAbout_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
        #If TWINBASIC Then
            Exit Sub
        #Else
            imgAbout.Visible = False
            imgAboutClicked.Visible = True
        #End If
    End Sub
    Private Sub imgDevelopment_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
        #If TWINBASIC Then
            Exit Sub
        #Else
            imgDevelopment.Visible = False
            imgDevelopmentClicked.Visible = True
        #End If
    End Sub
    Private Sub imgFonts_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
        #If TWINBASIC Then
            Exit Sub
        #Else
            imgFonts.Visible = False
            imgFontsClicked.Visible = True
        #End If
    End Sub
    Private Sub imgConfig_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
        #If TWINBASIC Then
            Exit Sub
        #Else
            imgConfig.Visible = False
            imgConfigClicked.Visible = True
        #End If
    End Sub
    Private Sub imgPosition_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

        #If TWINBASIC Then
            Exit Sub
        #Else
            imgPosition.Visible = False
            imgPositionClicked.Visible = True
        #End If
    End Sub
    Private Sub imgSounds_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
        #If TWINBASIC Then
            Exit Sub
        #Else
            imgSounds.Visible = False
            imgSoundsClicked.Visible = True
        #End If
    End Sub
    Private Sub imgWindow_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
        #If TWINBASIC Then
            Exit Sub
        #Else
            imgWindow.Visible = False
            imgWindowClicked.Visible = True
        #End If
    End Sub

Private Sub fraGeneralButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip fraGeneralButton.hWnd, "Clicking on the General icon reveals the General Tab where the essential items can be configured, alarms, startup &c.", _
                  TTIconInfo, "Help on the General Tab Icon", , , , True
End Sub

Private Sub fraConfigButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip fraConfigButton.hWnd, "Clicking on the Config icon reveals the Configuration Tab where the optional items can be configured, DPI, tooltips &c.", _
                  TTIconInfo, "Help on the Configuration Tab Icon", , , , True
End Sub

Private Sub fraFontsButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip fraFontsButton.hWnd, "Clicking on the Fonts icon reveals the Fonts Tab where the font related items can be configured, size, type, popups &c.", _
                  TTIconInfo, "Help on the Font Tab Icon", , , , True
End Sub
Private Sub fraSoundsButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip fraSoundsButton.hWnd, "Clicking on the Sounds icon reveals the Sounds Tab where sound related items can be configured, volume, type &c.", _
                  TTIconInfo, "Help on the Sounds Tab Icon", , , , True
End Sub
Private Sub fraPositionButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip fraPositionButton.hWnd, "Clicking on the Position icon reveals the Position Tab where items related to Positioning can be configured, aspect ratios, landscape, &c.", _
                  TTIconInfo, "Help on the Position Tab Icon", , , , True
End Sub
Private Sub fraDevelopmentButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip fraDevelopmentButton.hWnd, "Clicking on the Development icon reveals the Development Tab where items relating to Development can be configured, debug, VBP location, &c.", _
                  TTIconInfo, "Help on the Development Tab Icon", , , , True
End Sub
Private Sub fraWindowButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip fraWindowButton.hWnd, "Clicking on the Window icon reveals the Window Tab where items relating to window sizing and layering can be configured &c.", _
                  TTIconInfo, "Help on the Window Tab Icon", , , , True
End Sub
Private Sub fraAboutButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip fraAboutButton.hWnd, "Clicking on the About icon reveals the About Tab where information about this desktop widget can be revealed.", _
                  TTIconInfo, "Help on the About Tab Icon", , , , True
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
    
    If gblWidgetFunctions = "0" Or gblIgnoreMouse = "1" Then Exit Sub

    answer = vbYes
    answerMsg = "This option opens a browser window and take you straight to Github. Proceed?"
    answer = msgBoxA(answerMsg, vbExclamation + vbYesNo, "Proceed to Github? ", True, "lblGitHubDblClick")
    If answer = vbYes Then
        Call ShellExecute(Me.hWnd, "Open", "https://github.com/yereverluvinunclebert/Steampunk-Clock-Calendar-" & gblCodingEnvironment, vbNullString, App.path, 1)
    End If

   On Error GoTo 0
   Exit Sub

lblGitHub_dblClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblGitHub_dblClick of Form widgetPrefs"
End Sub

Private Sub lblGitHub_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblGitHub.ForeColor = &H8000000D
End Sub



Private Sub sliGaugeSize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip sliGaugeSize.hWnd, "Adjust to a percentage of the original size. Any adjustment in size made here takes place instantly (you can also use Ctrl+Mousewheel when hovering over the clock itself).", _
                  TTIconInfo, "Help on the Size Slider", , , , True
End Sub

Private Sub sliOpacity_Change()
    btnSave.Enabled = True ' enable the save button
End Sub
Private Sub sliOpacity_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip sliOpacity.hWnd, "Sliding this causes the program's opacity to change from solidly opaque to fully transparent or some way in-between. Seemingly, a strange option for a windows program, a useful left-over from the Yahoo Widgets days that offered this additional option. Replicated here as a homage to the old widget platform.", _
                  TTIconInfo, "Help on the Opacity Slider", , , , True

End Sub

'---------------------------------------------------------------------------------------
' Procedure : tmrPrefsScreenResolution_Timer
' Author    : beededea
' Date      : 05/05/2023
' Purpose   : for handling rotation of the screen in tablet mode or a resolution change
'             possibly due to an old game in full screen mode.
        ' when the timer frequency is reduced caused some weird async. effects, 500 ms seems fine, disable the timer first
'---------------------------------------------------------------------------------------
'
Private Sub tmrPrefsScreenResolution_Timer()

'    Static oldWidgetPrefsLeft As Long
'    Static oldWidgetPrefsTop As Long
'    Static beenMovingFlg As Boolean
'
'    Static oldPrefsFormMonitorID As Long
''    Static oldPrefsFormMonitorPrimary As Long
'    Static oldPrefsMonitorStructWidthTwips As Long
'    Static oldPrefsMonitorStructHeightTwips As Long
'    Static oldPrefsClockLeftPixels As Long
'
'    Dim prefsFormMonitorID As Long: prefsFormMonitorID = 0
'    'Dim prefsFormMonitorPrimary As Long: prefsFormMonitorPrimary = 0
'    Dim monitorStructWidthTwips As Long: monitorStructWidthTwips = 0
'    Dim monitorStructHeightTwips As Long: monitorStructHeightTwips = 0
'    Dim resizeProportion As Double: resizeProportion = 0
'
'    Dim answer As VbMsgBoxResult: answer = vbNo
'    Dim answerMsg As String: answerMsg = vbNullString
'
'    On Error GoTo tmrPrefsScreenResolution_Timer_Error
'
'    ' calls a routine that tests for a change in the monitor upon which the form sits, if so, resizes
'    If widgetPrefs.IsVisible = False Then Exit Sub
'
'    ' prefs hasn't moved at all
'    If widgetPrefs.Left = oldWidgetPrefsLeft Then Exit Sub  ' this can only work if the reposition is being performed by the timer
'    ' we are also hopefully calling this routine on a mouseUP event after a form move, where the above line will not apply.
'
'    ' if just one monitor or the global switch is off then exit
'    If monitorCount > 1 And (LTrim$(gblMultiMonitorResize) = "1" Or LTrim$(gblMultiMonitorResize) = "2") Then
'
'        ' turn off the timer that saves the prefs height and position
'        tmrPrefsMonitorSaveHeight.Enabled = False
'        tmrWritePosition.Enabled = False
'        tmrPrefsScreenResolution.Enabled = False ' turn off this very timer here
'
'        ' populate the OLD vars if empty, to allow valid comparison next run
'        If oldWidgetPrefsLeft <= 0 Then oldWidgetPrefsLeft = widgetPrefs.Left
'        If oldWidgetPrefsTop <= 0 Then oldWidgetPrefsTop = widgetPrefs.Top
'
'        ' test whether the form has been moved (VB6 has no FORM_MOVING nor FORM_MOVED EVENTS)
'        If widgetPrefs.Left <> oldWidgetPrefsLeft Or widgetPrefs.Top <> oldWidgetPrefsTop Then
'
'               ' note the monitor ID at PrefsForm form_load and store as the prefsFormMonitorID
'                prefsMonitorStruct = formScreenProperties(widgetPrefs, prefsFormMonitorID)
'
'                'prefsFormMonitorPrimary = prefsMonitorStruct.IsPrimary ' -1 true
'
'                ' sample the physical monitor resolution
'                monitorStructWidthTwips = prefsMonitorStruct.Width
'                monitorStructHeightTwips = prefsMonitorStruct.Height
'
'                'if the old monitor ID has not been stored already (form load) then do so now
'                If oldPrefsFormMonitorID = 0 Then oldPrefsFormMonitorID = prefsFormMonitorID
'
'                ' same with other 'old' vars
'                If oldPrefsMonitorStructWidthTwips = 0 Then oldPrefsMonitorStructWidthTwips = monitorStructWidthTwips
'                If oldPrefsMonitorStructHeightTwips = 0 Then oldPrefsMonitorStructHeightTwips = monitorStructHeightTwips
'                If oldPrefsClockLeftPixels = 0 Then oldPrefsClockLeftPixels = widgetPrefs.Left
'
'                ' if the monitor ID has changed
'                If oldPrefsFormMonitorID <> prefsFormMonitorID Then
'                'If oldPrefsFormMonitorPrimary <> prefsFormMonitorPrimary Then
'
''                    screenWrite ("Prefs Stored monitor primary status = " & CBool(oldPrefsFormMonitorPrimary))
''                    screenWrite ("Prefs Current monitor primary status = " & CBool(prefsFormMonitorPrimary))
'
'                    If LTrim$(gblMultiMonitorResize) = "1" Then
'                        'if the resolution is different then calculate new size proportion
'                        If monitorStructWidthTwips <> oldPrefsMonitorStructWidthTwips Or monitorStructHeightTwips <> oldPrefsMonitorStructHeightTwips Then
'                            'now calculate the size of the widget according to the screen HeightTwips.
'                            resizeProportion = prefsMonitorStruct.Height / oldPrefsMonitorStructHeightTwips
'                            newPrefsHeight = widgetPrefs.Height * resizeProportion
'                            widgetPrefs.Height = newPrefsHeight
'                        End If
'                    ElseIf LTrim$(gblMultiMonitorResize) = "2" Then
'                        ' set the size according to saved values
'                        If prefsMonitorStruct.IsPrimary = True Then
'                            widgetPrefs.Height = Val(gblPrefsPrimaryHeightTwips)
'                        Else
'                            'gblPrefsSecondaryHeightTwips = "15000"
'                            widgetPrefs.Height = Val(gblPrefsSecondaryHeightTwips)
'                        End If
'                    End If
'
'                End If
'
'                ' set the current values as 'old' for comparison on next run
'                'oldPrefsFormMonitorPrimary = prefsFormMonitorPrimary
'                oldPrefsFormMonitorID = prefsFormMonitorID
'                oldPrefsMonitorStructWidthTwips = monitorStructWidthTwips
'                oldPrefsMonitorStructHeightTwips = monitorStructHeightTwips
'                oldPrefsClockLeftPixels = widgetPrefs.Left
'            End If
'
'    End If
'
'    oldWidgetPrefsLeft = widgetPrefs.Left
'    oldWidgetPrefsTop = widgetPrefs.Top
'
'    tmrPrefsScreenResolution.Enabled = True
'    tmrPrefsMonitorSaveHeight.Enabled = True
'    tmrWritePosition.Enabled = True
    
    On Error GoTo 0
    Exit Sub

tmrPrefsScreenResolution_Timer_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrPrefsScreenResolution_Timer of Form frmTimer"
            Resume Next
          End If
    End With
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


Private Sub imgAbout_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picButtonMouseUpEvent("about", imgAbout, imgAboutClicked, fraAbout, fraAboutButton)
End Sub

Private Sub imgDevelopment_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picButtonMouseUpEvent("development", imgDevelopment, imgDevelopmentClicked, fraDevelopment, fraDevelopmentButton)
End Sub


Private Sub imgFonts_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picButtonMouseUpEvent("fonts", imgFonts, imgFontsClicked, fraFonts, fraFontsButton)
End Sub


Private Sub imgConfig_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picButtonMouseUpEvent("config", imgConfig, imgConfigClicked, fraConfig, fraConfigButton) ' was imgConfigMouseUpEvent
End Sub




Private Sub imgPosition_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picButtonMouseUpEvent("position", imgPosition, imgPositionClicked, fraPosition, fraPositionButton)
End Sub


Private Sub imgSounds_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Call imgSoundsMouseUpEvent
    Call picButtonMouseUpEvent("sounds", imgSounds, imgSoundsClicked, fraSounds, fraSoundsButton)
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
'    'overlayWidget.overlaySkewDeg = 'sliWidgetSkew.Value
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
        gblOpacity = CStr(sliOpacity.Value)
    
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

Private Sub txtAlarm1Time_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip txtAlarm1Time.hWnd, "Enter a valid Time here by typing a future time in 24hr military format, eg: 23:45. When you have entered both a date in the previous field and a time here, then click the > key to validate.", _
                  TTIconInfo, "Help on Alarm Time Number One", , , , True
End Sub


Private Sub txtAlarm2Time_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip txtAlarm2Time.hWnd, "Enter a valid Time here by typing a future time in 24hr military format, eg: 23:45. When you have entered both a date in the previous field and a time here, then click the > key to validate.", _
                  TTIconInfo, "Help on Alarm Time Number Two", , , , True
End Sub

Private Sub txtAlarm3Time_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip txtAlarm3Time.hWnd, "Enter a valid Time here by typing a future time in 24hr military format, eg: 23:45. When you have entered both a date in the previous field and a time here, then click the > key to validate.", _
                  TTIconInfo, "Help on Alarm Time Number Three", , , , True
End Sub

Private Sub txtAlarm4Time_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip txtAlarm4Time.hWnd, "Enter a valid Time here by typing a future time in 24hr military format, eg: 23:45. When you have entered both a date in the previous field and a time here, then click the > key to validate.", _
                  TTIconInfo, "Help on Alarm Time Number Four", , , , True
End Sub

Private Sub txtAlarm5Time_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip txtAlarm5Time.hWnd, "Enter a valid Time here by typing a future time in 24hr military format, eg: 23:45. When you have entered both a date in the previous field and a time here, then click the > key to validate.", _
                  TTIconInfo, "Help on Alarm Time Number Five", , , , True
End Sub


Private Sub txtAlarm1Date_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip txtAlarm1Date.hWnd, "Enter a valid date here by typing a future date in format dd/mm/yyyy. When you have entered both a date here and a time in the adjacent field, then click the > key to validate.", _
                  TTIconInfo, "Help on Alarm Date Number One", , , , True
End Sub


Private Sub txtAlarm2Date_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip txtAlarm2Date.hWnd, "Enter a valid date here by typing a future date in format dd/mm/yyyy. When you have entered both a date here and a time in the adjacent field, then click the > key to validate.", _
                  TTIconInfo, "Help on Alarm Date Number Two", , , , True
End Sub

Private Sub txtAlarm3Date_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip txtAlarm3Date.hWnd, "Enter a valid date here by typing a future date in format dd/mm/yyyy. When you have entered both a date here and a time in the adjacent field, then click the > key to validate.", _
                  TTIconInfo, "Help on Alarm Date Number Three", , , , True
End Sub

Private Sub txtAlarm4Date_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip txtAlarm4Date.hWnd, "Enter a valid date here by typing a future date in format dd/mm/yyyy. When you have entered both a date here and a time in the adjacent field, then click the > key to validate.", _
                  TTIconInfo, "Help on Alarm Date Number Four", , , , True
End Sub

Private Sub txtAlarm5Date_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip txtAlarm5Date.hWnd, "Enter a valid date here by typing a future date in format dd/mm/yyyy. When you have entered both a date here and a time in the adjacent field, then click the > key to validate.", _
                  TTIconInfo, "Help on Alarm Date Number Five", , , , True
End Sub
Private Sub txtAlarm1Time_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtAlarm2Date_Change()
    btnSave.Enabled = True ' enable the save button
End Sub



Private Sub txtAlarm2Time_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtAlarm3Date_Change()
    btnSave.Enabled = True ' enable the save button
End Sub



Private Sub txtAlarm3Time_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtAlarm4Date_Change()
    btnSave.Enabled = True ' enable the save button
End Sub





Private Sub txtAlarm4Time_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtAlarm5Date_Change()
    btnSave.Enabled = True ' enable the save button

End Sub



Private Sub txtAlarm5Time_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtDblClickCommand_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtDblClickCommand_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip txtDblClickCommand.hWnd, "Field to hold the any double click command that you have assigned to this widget. For example: taskmgr or %systemroot%\syswow64\ncpa.cpl", _
                  TTIconInfo, "Help on the Double Click Command", , , , True
End Sub

Private Sub txtDefaultEditor_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtDefaultEditor_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip txtDefaultEditor.hWnd, "Field to hold the path to a Visual Basic Project (VBP) file you would like to execute on a right click menu, edit option, if you select the adjacent button a file explorer will appear allowing you to select the VBP file, this field is automatically filled with the chosen file.", _
                  TTIconInfo, "Help on the Default Editor Field", , , , True
End Sub

Private Sub txtDisplayScreenFont_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip txtDisplayScreenFont.hWnd, "This is a read-only text box. It displays the current font - as set when you click the font selector button. This field is in operation for informational purposes only. When resizing the main clock (CTRL+ mouse scroll wheel) the font size will change in relation to clock size. The base font determines the initial size, the resulting resized font will dynamically change. My preferred font for the display screen is Courier New at 6pt size.", _
                  TTIconInfo, "Help on the Display Screen Font", , , , True
End Sub



Private Sub txtDisplayScreenFontSize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip txtDisplayScreenFontSize.hWnd, "This is a read-only text box. It displays the current base font size as set when dynamic form resizing is enabled. The adjacent text box will display the automatically resized font currently in operation, for informational purposes only.", _
                  TTIconInfo, "Help on the Base Font Size for Display Screen", , , , True
End Sub

Private Sub txtLandscapeHoffset_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtLandscapeHoffset_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip txtLandscapeHoffset.hWnd, "Field to hold the horizontal offset for the widget position in landscape mode. When you lock the widget using the lock button above, this field is automatically filled.", _
                  TTIconInfo, "Help on the Landscape X Horizontal Field", , , , True
End Sub

Private Sub txtLandscapeVoffset_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtLandscapeVoffset_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip txtLandscapeVoffset.hWnd, "Field to hold the vertical offset for the widget position in landscape mode. When you lock the widget using the lock button above, this field is automatically filled.", _
                  TTIconInfo, "Help on the Landscape Y Vertical Field", , , , True
End Sub

Private Sub txtOpenFile_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtOpenFile_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip txtOpenFile.hWnd, "Field to hold the path to a file you would like to execute on a shift+DBlClick, if you select the adjacent button a file explorer will appear allowing you to select any file, this field is automatically filled with the chosen file.", _
                  TTIconInfo, "Help on the shift+DBlClick Field", , , , True
End Sub

Private Sub txtPortraitHoffset_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtPortraitHoffset_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip txtPortraitHoffset.hWnd, "Field to hold the horizontal offset for the widget position in Portrait mode. When you lock the widget using the lock button above, this field is automatically filled.", _
                  TTIconInfo, "Help on the Portrait X Horizontal Field", , , , True
End Sub

Private Sub txtPortraitYoffset_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtPortraitYoffset_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip txtPortraitYoffset.hWnd, "Field to hold the vertical offset for the widget position in Portrait mode. When you lock the widget using the lock button above, this field is automatically filled.", _
                  TTIconInfo, "Help on the Portrait Y Vertical Field", , , , True
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

    If optEnablePrefsBalloonTooltips.Value = True Then
        ' module level balloon tooltip variables for subclassed comboBoxes ONLY.
        gCmbMultiMonitorResizeBalloonTooltip = "This option will only appear on multi-monitor systems. This dropdown has three choices that affect the automatic sizing of both the main clock and the preference utility. " & vbCrLf & vbCrLf & _
            "For monitors of different sizes, this allows you to resize the widget to suit the monitor it is currently sitting on. The automatic option resizes according to the relative proportions of the two screens.  " & vbCrLf & vbCrLf & _
            "The manual option resizes according to sizes that you set manually. Just resize the clock on the monitor of your choice and the program will store it. This option only works for no more than TWO monitors."
   
        gCmbScrollWheelDirectionBalloonTooltip = "This option will allow you to change the direction of the mouse scroll wheel when resizing the clock gauge. IF you want to resize the clock on your desktop, hold the CTRL key along with moving the scroll wheel UP/DOWN. Some prefer scrolling UP rather than DOWN. You configure that here."
        gCmbWindowLevelBalloonTooltip = "You can determine the window level here. You can keep it above all other windows or you can set it to bottom to keep the widget below all other windows."
        gCmbHidingTimeBalloonTooltip = "The hiding time that you can set here determines how long the widget will disappear when you click the menu option to hide the widget."
        
        gCmbWidgetLandscapeBalloonTooltip = "The widget can be locked into landscape mode, it ensures that the widget always appears where you want it to. Using the fields below, you can assign a default x/y position for Landscape mode. "
        gCmbWidgetPortraitBalloonTooltip = "The widget can be locked into portrait mode, it ensures that the widget always appears where you want it to. Using the fields below, you can assign a default x/y position for portrait mode. "
        gCmbWidgetPositionBalloonTooltip = "Tablets only. The widget can be positioned proportionally when switching between portrait/landscape. If you want to enable this, disable the options below."
        gCmbAspectHiddenBalloonTooltip = "Here you can choose whether the widget is hidden by default in either landscape or portrait mode or not at all. This allows you to have certain widgets that do not obscure the screen in one mode or another. If you accidentally set it so you can't find it on screen then change the setting here to none."
        
        gCmbDebugBalloonTooltip = "Here you can set debug mode. This will enable the editor field and allow you to assign a VBP file for the " & gblCodingEnvironment & " IDE editor"
        
        gCmbAlarmDayBalloonTooltip = "Enter a valid day of the month here. When you have entered both a date here and a time in the adjacent field, then click the > key to validate."
        gCmbAlarmMonthBalloonTooltip = "Enter a valid month here. When you have entered both a date here and a time in the adjacent field, then click the > key to validate."
        gCmbAlarmYearBalloonTooltip = "Enter a valid year here. When you have entered both a valid year here and a time in the adjacent fields, then click the > key to validate."
        gCmbAlarmHoursBalloonTooltip = "Enter a valid hour here by typing a future time in 24hr military format, eg: 23:45. When you have entered both a date in the previous fields and a time here in these two fields, then click the > key to validate."
        gCmbAlarmMinutesBalloonTooltip = "Enter valid minutes here by typing a future time in 24hr military format, eg: 23:45. When you have entered both a date in the previous fields and a time here, then click the > key to validate."
    Else
        ' module level balloon tooltip variables for subclassed comboBoxes ONLY.
        
        gCmbMultiMonitorResizeBalloonTooltip = vbNullString
        gCmbScrollWheelDirectionBalloonTooltip = vbNullString
        gCmbWindowLevelBalloonTooltip = vbNullString
        gCmbHidingTimeBalloonTooltip = vbNullString
        
        gCmbWidgetLandscapeBalloonTooltip = vbNullString
        gCmbWidgetPortraitBalloonTooltip = vbNullString
        gCmbWidgetPositionBalloonTooltip = vbNullString
        gCmbAspectHiddenBalloonTooltip = vbNullString
        gCmbDebugBalloonTooltip = vbNullString
        gCmbAlarmDayBalloonTooltip = vbNullString
        gCmbAlarmMonthBalloonTooltip = vbNullString
        gCmbAlarmYearBalloonTooltip = vbNullString
        
        gCmbAlarmHoursBalloonTooltip = vbNullString
        gCmbAlarmMinutesBalloonTooltip = vbNullString
        
        ' for some reason, the balloon tooltip on the checkbox used to dismiss the balloon tooltips does not disappear, this forces it go away.
        CreateToolTip optEnablePrefsBalloonTooltips.hWnd, "", _
                  TTIconInfo, "Help", , , , True
        
    End If
     
     If optEnablePrefsTooltips.Value = True Then
       
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
        lblPosition(6).ToolTipText = "Tablets only. Don't fiddle with this unless you really know what you are doing. Here you can choose whether this the widget widget is hidden by default in either landscape or portrait mode or not at all. This option allows you to have certain widgets that do not obscure the screen in either landscape or portrait. If you accidentally set it so you can't find your widget on screen then change the setting here to NONE."
        chkGenStartup.ToolTipText = "Check this box to enable the automatic start of the program when Windows is started."
        chkTogglePendulum.ToolTipText = "Decide whether the pendulum swings or not. The math behind the graphics of a swinging pendulum requires processor power. Older single-core machines may experience a lot of cpu usage."
        chk24HourClockMode.ToolTipText = "Decide whether the digital clock display should use military/twenty-four hour time or the 12-hour clock."
        chkWidgetFunctions.ToolTipText = "When checked this box enables the spinning earth functionality. Any adjustment takes place instantly. "

        btnNow(1).ToolTipText = "Fill all alarm fields in number 1 alarm with the date and time now."
        btnNow(2).ToolTipText = "Fill all alarm fields in number 2 alarm with the date and time now."
        btnNow(3).ToolTipText = "Fill all alarm fields in number 3 alarm with the date and time now."
        btnNow(4).ToolTipText = "Fill all alarm fields in number 4 alarm with the date and time now."
        btnNow(5).ToolTipText = "Fill all alarm fields in number 5 alarm with the date and time now."
        
        btnclearAlarm(1).ToolTipText = "Clear all alarm fields in number 1 alarm."
        btnclearAlarm(2).ToolTipText = "Clear all alarm fields in number 2 alarm."
        btnclearAlarm(3).ToolTipText = "Clear all alarm fields in number 3 alarm."
        btnclearAlarm(4).ToolTipText = "Clear all alarm fields in number 4 alarm."
        btnclearAlarm(5).ToolTipText = "Clear all alarm fields in number 5 alarm."
        
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
        
        chkVolumeBoost.ToolTipText = "Sets the volume of the various sound elements, you can boost from quiet to loud."
        btnDefaultEditor.ToolTipText = "Click to select the .vbp file to edit the program - You need to have access to the source!"
        txtDblClickCommand.ToolTipText = "Enter a Windows command for the gauge to operate when double-clicked."
        btnOpenFile.ToolTipText = "Click to select a particular file for the gauge to run or open when double-clicked."
        txtOpenFile.ToolTipText = "Enter a particular file for the gauge to run or open when double-clicked."
        cmbDebug.ToolTipText = "Choose to set debug mode."
        
        txtPrefsFontSize.ToolTipText = "Disabled for manual input. Choose a font size using the font selector to be used within this preferences window only"
        btnPrefsFont.ToolTipText = "The Font Selector."
        txtPrefsFont.ToolTipText = "Disabled for manual input. Choose a font via the font selector to be used only for this preferences window"
        txtPrefsFontSize.ToolTipText = "Disabled for manual input. Choose a font size via the font selector that fits the text boxes"
        
         
        txtDisplayScreenFontSize.ToolTipText = "Disabled for manual input. Choose a font size using the font selector to be used within the clock display screen only"
        btnDisplayScreenFont.ToolTipText = "The Font Selector."
        txtDisplayScreenFont.ToolTipText = "Disabled for manual input. Choose a font size using the font selector to be used within the clock display screen only"
        
        cmbWindowLevel.ToolTipText = "You can determine the window position here. Set to bottom to keep the widget below other windows."
        cmbHidingTime.ToolTipText = "The hiding time that you can set here determines how long the widget will disappear when you click the menu option to hide the widget."
        cmbMultiMonitorResize.ToolTipText = "When you have a multi-monitor set-up, the widget can auto-resize on a smaller secondary monitor. Here you determine the proportion of the resize."
        
        chkEnableResizing.ToolTipText = "Provides an alternative method of supporting high DPI screens."
        chkPreventDragging.ToolTipText = "Checking this box turns off the ability to drag the program with the mouse. The locking in position effect takes place instantly."
        chkIgnoreMouse.ToolTipText = "Checking this box causes the program to ignore all mouse events."
        sliOpacity.ToolTipText = "Set the transparency of the program. Any change in opacity takes place instantly."
        cmbScrollWheelDirection.ToolTipText = "To change the direction of the mouse scroll wheel when resizing the clock gauge."
        optEnableBalloonTooltips.ToolTipText = "Check the box to enable larger balloon tooltips for all controls on the main program"
        chkShowTaskbar.ToolTipText = "Check the box to show the widget in the taskbar"
        chkShowHelp.ToolTipText = "Check the box to show the help page on startup"
        optEnableTooltips.ToolTipText = "Check the box to enable tooltips for all controls on the main program"
        sliGaugeSize.ToolTipText = "Adjust to a percentage of the original size. Any adjustment in size takes place instantly (you can also use Ctrl+Mousewheel hovering over the clock itself)."
        btnFacebook.ToolTipText = "This will link you to the our Steampunk/Dieselpunk program users Group."
        imgAbout.ToolTipText = "Opens the About tab"
        btnAboutDebugInfo.ToolTipText = "This gives access to the debugging tool"
        btnDonate.ToolTipText = "Buy me a Kofi! This button opens a browser window and connects to Kofi donation page"
        btnUpdate.ToolTipText = "Here you can visit the update location where you can download new versions of the programs."
        
        btnGithubHome.ToolTipText = "Here you can visit the widget's home page on github, when you click the button it will open a browser window and take you to the github home page."

        txtPrefsFontCurrentSize.ToolTipText = "Disabled for manual input. Shows the current font size when form resizing is enabled."
        'lblCurrentFontsTab.ToolTipText = "Disabled for manual input. Shows the current font size when form resizing is enabled."
        
        chkDpiAwareness.ToolTipText = " Check the box to make the program DPI aware. RESTART required."
        optEnablePrefsTooltips.ToolTipText = "Check the box to enable tooltips for all controls in the preferences utility"
        optEnablePrefsBalloonTooltips.ToolTipText = "This setting enables/disables the enhanced balloon tooltips for elements within this Preference Utility. These tooltips are multi-line and in general more attractive, note that their font size will match the Windows system font size."

        btnResetMessages.ToolTipText = "This button restores the pop-up messages to their original visible state."
        
        cmbAlarm1Day.ToolTipText = "Enter a valid day of the month here. When you have entered both a date here and a time in the adjacent field, then click the > key to validate."
        cmbAlarm2Day.ToolTipText = "Enter a valid day of the month here. When you have entered both a date here and a time in the adjacent field, then click the > key to validate."
        cmbAlarm3Day.ToolTipText = "Enter a valid day of the month here. When you have entered both a date here and a time in the adjacent field, then click the > key to validate."
        cmbAlarm4Day.ToolTipText = "Enter a valid day of the month here. When you have entered both a date here and a time in the adjacent field, then click the > key to validate."
        cmbAlarm5Day.ToolTipText = "Enter a valid day of the month here. When you have entered both a date here and a time in the adjacent field, then click the > key to validate."
        
        cmbAlarm1Month.ToolTipText = "Enter a valid month here. When you have entered both a date here and a time in the adjacent field, then click the > key to validate."
        cmbAlarm2Month.ToolTipText = "Enter a valid month here. When you have entered both a date here and a time in the adjacent field, then click the > key to validate."
        cmbAlarm3Month.ToolTipText = "Enter a valid month here. When you have entered both a date here and a time in the adjacent field, then click the > key to validate."
        cmbAlarm4Month.ToolTipText = "Enter a valid month here. When you have entered both a date here and a time in the adjacent field, then click the > key to validate."
        cmbAlarm5Month.ToolTipText = "Enter a valid month here. When you have entered both a date here and a time in the adjacent field, then click the > key to validate."
        
        cmbAlarm1Year.ToolTipText = "Enter a valid year here. When you have entered both a valid year here and a time in the adjacent fields, then click the > key to validate."
        cmbAlarm2Year.ToolTipText = "Enter a valid year here. When you have entered both a valid year here and a time in the adjacent fields, then click the > key to validate."
        cmbAlarm3Year.ToolTipText = "Enter a valid year here. When you have entered both a valid year here and a time in the adjacent fields, then click the > key to validate."
        cmbAlarm4Year.ToolTipText = "Enter a valid year here. When you have entered both a valid year here and a time in the adjacent fields, then click the > key to validate."
        cmbAlarm5Year.ToolTipText = "Enter a valid year here. When you have entered both a valid year here and a time in the adjacent fields, then click the > key to validate."
        
        cmbAlarm1Hours.ToolTipText = "Enter a valid hour here by typing a future time in 24hr military format, eg: 23:45. When you have entered both a date in the previous fields and a time here in these two fields, then click the > key to validate."
        cmbAlarm2Hours.ToolTipText = "Enter a valid hour here by typing a future time in 24hr military format, eg: 23:45. When you have entered both a date in the previous fields and a time here in these two fields, then click the > key to validate."
        cmbAlarm3Hours.ToolTipText = "Enter a valid hour here by typing a future time in 24hr military format, eg: 23:45. When you have entered both a date in the previous fields and a time here in these two fields, then click the > key to validate."
        cmbAlarm4Hours.ToolTipText = "Enter a valid hour here by typing a future time in 24hr military format, eg: 23:45. When you have entered both a date in the previous fields and a time here in these two fields, then click the > key to validate."
        cmbAlarm5Hours.ToolTipText = "Enter a valid hour here by typing a future time in 24hr military format, eg: 23:45. When you have entered both a date in the previous fields and a time here in these two fields, then click the > key to validate."
        
        cmbAlarm1Minutes.ToolTipText = "Enter valid minutes here by typing a future time in 24hr military format, eg: 23:45. When you have entered both a date in the previous fields and a time here, then click the > key to validate."
        cmbAlarm2Minutes.ToolTipText = "Enter valid minutes here by typing a future time in 24hr military format, eg: 23:45. When you have entered both a date in the previous fields and a time here, then click the > key to validate."
        cmbAlarm3Minutes.ToolTipText = "Enter valid minutes here by typing a future time in 24hr military format, eg: 23:45. When you have entered both a date in the previous fields and a time here, then click the > key to validate."
        cmbAlarm4Minutes.ToolTipText = "Enter valid minutes here by typing a future time in 24hr military format, eg: 23:45. When you have entered both a date in the previous fields and a time here, then click the > key to validate."
        cmbAlarm5Minutes.ToolTipText = "Enter valid minutes here by typing a future time in 24hr military format, eg: 23:45. When you have entered both a date in the previous fields and a time here, then click the > key to validate."
        
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
        chkTogglePendulum.ToolTipText = vbNullString
        chk24HourClockMode.ToolTipText = vbNullString
        chkWidgetFunctions.ToolTipText = vbNullString
        

        btnNow(1).ToolTipText = vbNullString
        btnNow(2).ToolTipText = vbNullString
        btnNow(3).ToolTipText = vbNullString
        btnNow(4).ToolTipText = vbNullString
        btnNow(5).ToolTipText = vbNullString
        
        btnclearAlarm(1).ToolTipText = vbNullString
        btnclearAlarm(2).ToolTipText = vbNullString
        btnclearAlarm(3).ToolTipText = vbNullString
        btnclearAlarm(4).ToolTipText = vbNullString
        btnclearAlarm(5).ToolTipText = vbNullString
        
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
        chkVolumeBoost.ToolTipText = vbNullString
        
        btnDefaultEditor.ToolTipText = vbNullString
        txtDblClickCommand.ToolTipText = vbNullString
        btnOpenFile.ToolTipText = vbNullString
        txtOpenFile.ToolTipText = vbNullString
        cmbDebug.ToolTipText = vbNullString
        txtPrefsFontSize.ToolTipText = vbNullString
        btnPrefsFont.ToolTipText = vbNullString
        txtPrefsFont.ToolTipText = vbNullString
        txtPrefsFontCurrentSize.ToolTipText = vbNullString
        
        
        txtDisplayScreenFontSize.ToolTipText = vbNullString
        btnDisplayScreenFont.ToolTipText = vbNullString
        txtDisplayScreenFont.ToolTipText = vbNullString
        
        cmbWindowLevel.ToolTipText = vbNullString
        cmbHidingTime.ToolTipText = vbNullString
        cmbMultiMonitorResize.ToolTipText = vbNullString
        
        chkEnableResizing.ToolTipText = vbNullString
        chkPreventDragging.ToolTipText = vbNullString
        chkIgnoreMouse.ToolTipText = vbNullString
        sliOpacity.ToolTipText = vbNullString
        cmbScrollWheelDirection.ToolTipText = vbNullString
        optEnableBalloonTooltips.ToolTipText = vbNullString
        chkShowTaskbar.ToolTipText = vbNullString
        chkShowHelp.ToolTipText = vbNullString
        optEnableTooltips.ToolTipText = vbNullString
        sliGaugeSize.ToolTipText = vbNullString
        btnFacebook.ToolTipText = vbNullString
        imgAbout.ToolTipText = vbNullString
        btnAboutDebugInfo.ToolTipText = vbNullString
        btnDonate.ToolTipText = vbNullString
        btnUpdate.ToolTipText = vbNullString
        btnGithubHome.ToolTipText = vbNullString
        
        chkDpiAwareness.ToolTipText = vbNullString
        optEnablePrefsTooltips.ToolTipText = vbNullString
        optEnablePrefsBalloonTooltips.ToolTipText = vbNullString
        btnResetMessages.ToolTipText = vbNullString
    
        cmbAlarm1Day.ToolTipText = vbNullString
        cmbAlarm2Day.ToolTipText = vbNullString
        cmbAlarm3Day.ToolTipText = vbNullString
        cmbAlarm4Day.ToolTipText = vbNullString
        cmbAlarm5Day.ToolTipText = vbNullString
        
        cmbAlarm1Month.ToolTipText = vbNullString
        cmbAlarm2Month.ToolTipText = vbNullString
        cmbAlarm3Month.ToolTipText = vbNullString
        cmbAlarm4Month.ToolTipText = vbNullString
        cmbAlarm5Month.ToolTipText = vbNullString
        
        cmbAlarm1Year.ToolTipText = vbNullString
        cmbAlarm2Year.ToolTipText = vbNullString
        cmbAlarm3Year.ToolTipText = vbNullString
        cmbAlarm4Year.ToolTipText = vbNullString
        cmbAlarm5Year.ToolTipText = vbNullString
        
        cmbAlarm1Hours.ToolTipText = vbNullString
        cmbAlarm2Hours.ToolTipText = vbNullString
        cmbAlarm3Hours.ToolTipText = vbNullString
        cmbAlarm4Hours.ToolTipText = vbNullString
        cmbAlarm5Hours.ToolTipText = vbNullString
        
        cmbAlarm1Minutes.ToolTipText = vbNullString
        cmbAlarm2Minutes.ToolTipText = vbNullString
        cmbAlarm3Minutes.ToolTipText = vbNullString
        cmbAlarm4Minutes.ToolTipText = vbNullString
        cmbAlarm5Minutes.ToolTipText = vbNullString
        
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

'    lblFontsTab(0).Caption = "When resizing the form (drag bottom right) the font size will in relation to form height. The base font determines the initial size, the resulting resized font will dynamically change." & vbCrLf & vbCrLf & _
'        "" & vbCrLf & vbCrLf & _
'        "My preferred font for this utility is Centurion Light SF at 8pt size."

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
    
    lblAbout(1).Caption = "(32bit WoW64 using " & gblCodingEnvironment & ")"
    
    Call LoadFileToTB(txtAboutText, App.path & "\resources\txt\about.txt", False)
    
    fClock.RotateBusyTimer = True
    
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
    Dim BorderWidth As Long: BorderWidth = 0
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
    y_scale = Me.ScaleHeight / gblPrefsCurrentHeight
    
    If gblDpiAwareness = "1" Then
        btnHelp.Top = fraGeneral.Top + fraGeneral.Height + (100 * y_scale)
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
    
    BorderWidth = (widgetPrefs.Width - Me.ScaleWidth) / 2
    captionHeight = widgetPrefs.Height - Me.ScaleHeight - BorderWidth
        
    ' under windows 10+ the internal window calcs are all wrong due to the bigger title bars
    If prefsDynamicSizingFlg = False Then
        padding = 200 ' add normal padding below the help button to position the bottom of the form

        lastFormHeight = btnHelp.Top + btnHelp.Height + captionHeight + BorderWidth + padding
        gblPrefsFormResizedInCode = True
        widgetPrefs.Height = lastFormHeight
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
    
    If widgetPrefs.IsVisible = False Then Exit Sub

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




'---------------------------------------------------------------------------------------
' Procedure : mnuClosePreferences_Click
' Author    : beededea
' Date      : 06/09/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuClosePreferences_Click()
   On Error GoTo mnuClosePreferences_Click_Error

    Call btnClose_Click

   On Error GoTo 0
   Exit Sub

mnuClosePreferences_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuClosePreferences_Click of Form widgetPrefs"
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
    
    fClock.RotateBusyTimer = True

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
    
    'the widgetPrefs.mnuAuto.Caption = "Auto Theme Selection Cannot be Enabled"

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
    
    fClock.RotateBusyTimer = True
    
   On Error GoTo 0
   Exit Sub

loadHigherResPrefsImages_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadHigherResPrefsImages of Form widgetPrefs"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : tmrWritePosition_Timer
' Author    : beededea
' Date      : 27/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub tmrWritePosition_Timer()
    ' save the current X and y position of this form to allow repositioning when restarting
    On Error GoTo tmrWritePosition_Timer_Error
   
    If widgetPrefs.IsVisible = True Then Call writePrefsPosition

   On Error GoTo 0
   Exit Sub

tmrWritePosition_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrWritePosition_Timer of Form widgetPrefs"

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
        'lblCurrentFontsTab.Visible = True
        'Call writePrefsPosition
        chkEnableResizing.Caption = "Disable Corner Resizing"
    Else
        prefsDynamicSizingFlg = False
        txtPrefsFontCurrentSize.Visible = False
        'lblCurrentFontsTab.Visible = False
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
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip chkEnableResizing.hWnd, "This allows you to resize the whole prefs window by dragging the bottom right corner of the window. It provides an alternative method of supporting high DPI screens.", _
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
            Call SaveSizes(widgetPrefs, prefsControlPositions(), gblPrefsCurrentWidth, gblPrefsCurrentHeight)
        'End If
    Else
        fraGeneral.Height = 7737
        fraConfig.Height = 7134
        fraSounds.Height = 3376
        fraPosition.Height = 7544
        fraFonts.Height = 4533
        
        ' the lowest window controls are not displayed on a single monitor
        If monitorCount > 1 Then
            fraWindow.Height = 8138
            fraWindowInner.Height = 7500
        Else
            fraWindow.Height = 6586
            fraWindowInner.Height = 6000
        End If

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
    
    #If TWINBASIC Then
    
        Set imgGeneral.Picture = LoadPicture(App.path & "\Resources\images\general-icon-dark-1010.jpg")
        Set imgConfig.Picture = LoadPicture(App.path & "\Resources\images\config-icon-dark-1010.jpg")
        Set imgFonts.Picture = LoadPicture(App.path & "\Resources\images\font-icon-dark-1010.jpg")
        Set imgSounds.Picture = LoadPicture(App.path & "\Resources\images\sounds-icon-dark-1010.jpg")
        Set imgPosition.Picture = LoadPicture(App.path & "\Resources\images\position-icon-dark-1010.jpg")
        Set imgDevelopment.Picture = LoadPicture(App.path & "\Resources\images\development-icon-dark-1010.jpg")
        Set imgWindow.Picture = LoadPicture(App.path & "\Resources\images\windows-icon-dark-1010.jpg")
        Set imgAbout.Picture = LoadPicture(App.path & "\Resources\images\about-icon-dark-1010.jpg")
    '
        Set imgGeneralClicked.Picture = LoadPicture(App.path & "\Resources\images\general-icon-dark-600-clicked.jpg")
        Set imgConfigClicked.Picture = LoadPicture(App.path & "\Resources\images\config-icon-dark-600-clicked.jpg")
        Set imgFontsClicked.Picture = LoadPicture(App.path & "\Resources\images\font-icon-dark-600-clicked.jpg")
        Set imgSoundsClicked.Picture = LoadPicture(App.path & "\Resources\images\sounds-icon-dark-600-clicked.jpg")
        Set imgPositionClicked.Picture = LoadPicture(App.path & "\Resources\images\position-icon-dark-600-clicked.jpg")
        Set imgDevelopmentClicked.Picture = LoadPicture(App.path & "\Resources\images\development-icon-dark-600-clicked.jpg")
        Set imgWindowClicked.Picture = LoadPicture(App.path & "\Resources\images\windows-icon-dark-600-clicked.jpg")
        Set imgAboutClicked.Picture = LoadPicture(App.path & "\Resources\images\about-icon-dark-600-clicked.jpg")
    
    #Else
        
        Set imgGeneral.Picture = Cairo.ImageList("general-icon-dark").Picture
        fClock.RotateBusyTimer = True
    
        Set imgConfig.Picture = Cairo.ImageList("config-icon-dark").Picture
        fClock.RotateBusyTimer = True
        
        Set imgConfig.Picture = LoadPicture(App.path & "\Resources\images\config-icon-dark-1010.jpg")
        fClock.RotateBusyTimer = True
        
        Set imgFonts.Picture = Cairo.ImageList("font-icon-dark").Picture
        fClock.RotateBusyTimer = True
        
        Set imgSounds.Picture = Cairo.ImageList("sounds-icon-dark").Picture
        fClock.RotateBusyTimer = True
        
        Set imgPosition.Picture = Cairo.ImageList("position-icon-dark").Picture
        fClock.RotateBusyTimer = True
        
        Set imgDevelopment.Picture = Cairo.ImageList("development-icon-dark").Picture
        fClock.RotateBusyTimer = True
        
        Set imgWindow.Picture = Cairo.ImageList("windows-icon-dark").Picture
        fClock.RotateBusyTimer = True
        
        Set imgAbout.Picture = Cairo.ImageList("about-icon-dark").Picture
        fClock.RotateBusyTimer = True
        
    '
        Set imgGeneralClicked.Picture = Cairo.ImageList("general-icon-dark-clicked").Picture
        fClock.RotateBusyTimer = True
        
        Set imgConfigClicked.Picture = Cairo.ImageList("config-icon-dark-clicked").Picture
        fClock.RotateBusyTimer = True
        
        Set imgFontsClicked.Picture = Cairo.ImageList("font-icon-dark-clicked").Picture
        fClock.RotateBusyTimer = True
        
        Set imgSoundsClicked.Picture = Cairo.ImageList("sounds-icon-dark-clicked").Picture
        fClock.RotateBusyTimer = True
        
        Set imgPositionClicked.Picture = Cairo.ImageList("position-icon-dark-clicked").Picture
        fClock.RotateBusyTimer = True
        
        Set imgDevelopmentClicked.Picture = Cairo.ImageList("development-icon-dark-clicked").Picture
        fClock.RotateBusyTimer = True
        
        Set imgWindowClicked.Picture = Cairo.ImageList("windows-icon-dark-clicked").Picture
        fClock.RotateBusyTimer = True
        
        Set imgAboutClicked.Picture = Cairo.ImageList("about-icon-dark-clicked").Picture
        fClock.RotateBusyTimer = True
        
    
    #End If

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
    
    #If TWINBASIC Then
    
        Set imgGeneral.Picture = LoadPicture(App.path & "\Resources\images\general-icon-light-1010.jpg")
        Set imgConfig.Picture = LoadPicture(App.path & "\Resources\images\config-icon-light-1010.jpg")
        Set imgFonts.Picture = LoadPicture(App.path & "\Resources\images\font-icon-light-1010.jpg")
        Set imgSounds.Picture = LoadPicture(App.path & "\Resources\images\sounds-icon-light-1010.jpg")
        Set imgPosition.Picture = LoadPicture(App.path & "\Resources\images\position-icon-light-1010.jpg")
        Set imgDevelopment.Picture = LoadPicture(App.path & "\Resources\images\development-icon-light-1010.jpg")
        Set imgWindow.Picture = LoadPicture(App.path & "\Resources\images\windows-icon-light-1010.jpg")
        Set imgAbout.Picture = LoadPicture(App.path & "\Resources\images\about-icon-light-1010.jpg")
    '
        Set imgGeneralClicked.Picture = LoadPicture(App.path & "\Resources\images\general-icon-light-600-clicked.jpg")
        Set imgConfigClicked.Picture = LoadPicture(App.path & "\Resources\images\config-icon-light-600-clicked.jpg")
        Set imgFontsClicked.Picture = LoadPicture(App.path & "\Resources\images\font-icon-light-600-clicked.jpg")
        Set imgSoundsClicked.Picture = LoadPicture(App.path & "\Resources\images\sounds-icon-light-600-clicked.jpg")
        Set imgPositionClicked.Picture = LoadPicture(App.path & "\Resources\images\position-icon-light-600-clicked.jpg")
        Set imgDevelopmentClicked.Picture = LoadPicture(App.path & "\Resources\images\development-icon-light-600-clicked.jpg")
        Set imgWindowClicked.Picture = LoadPicture(App.path & "\Resources\images\windows-icon-light-600-clicked.jpg")
        Set imgAboutClicked.Picture = LoadPicture(App.path & "\Resources\images\about-icon-light-600-clicked.jpg")
    
    #Else
        
        Set imgGeneral.Picture = Cairo.ImageList("general-icon-light").Picture
        Set imgConfig.Picture = Cairo.ImageList("config-icon-light").Picture
        
        Set imgConfig.Picture = LoadPicture(App.path & "\Resources\images\config-icon-light-1010.jpg")
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
    
    #End If
        
   On Error GoTo 0
   Exit Sub

setPrefsIconImagesLight_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setPrefsIconImagesLight of Form widgetPrefs"

End Sub



Private Sub txtPrefsFont_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip txtPrefsFont.hWnd, "This is a read-only text box. It displays the current font as set when you click the font selector button. This is in operation for informational purposes only. When resizing the form (drag bottom right) the font size will change in relation to form height. The base font determines the initial size, the resulting resized font will dynamically change.  My preferred font for this utility is Centurion Light SF at 8pt size.", _
                  TTIconInfo, "Help on the Currently Selected Font", , , , True
End Sub

Private Sub txtPrefsFontCurrentSize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip txtPrefsFontCurrentSize.hWnd, "This is a read-only text box. It displays the current font size as set when dynamic form resizing is enabled. Drag the right hand corner of the window downward and the form will auto-resize. This text box will display the resized font currently in operation for informational purposes only.", _
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
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, 0
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


'---------------------------------------------------------------------------------------
' Procedure : tmrPrefsMonitorSaveHeight_Timer
' Author    : beededea
' Date      : 26/08/2024
' Purpose   : save the current height of this form to allow resizing when restarting or placing on another monitor
'---------------------------------------------------------------------------------------
'
Private Sub tmrPrefsMonitorSaveHeight_Timer()

    'Dim prefsFormMonitorID As Long: prefsFormMonitorID = 0
    
    On Error GoTo tmrPrefsMonitorSaveHeight_Timer_Error
    
    If widgetPrefs.IsVisible = False Then Exit Sub

    If LTrim$(gblMultiMonitorResize) <> "2" Then Exit Sub
 
    If prefsMonitorStruct.IsPrimary = True Then
        gblPrefsPrimaryHeightTwips = Trim$(CStr(widgetPrefs.Height))
        sPutINISetting "Software\SteampunkClockCalendar", "prefsPrimaryHeightTwips", gblPrefsPrimaryHeightTwips, gblSettingsFile
    Else
        gblPrefsSecondaryHeightTwips = Trim$(CStr(widgetPrefs.Height))
        sPutINISetting "Software\SteampunkClockCalendar", "prefsSecondaryHeightTwips", gblPrefsSecondaryHeightTwips, gblSettingsFile
    End If

   On Error GoTo 0
   Exit Sub

tmrPrefsMonitorSaveHeight_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrPrefsMonitorSaveHeight_Timer of Form widgetPrefs"

End Sub



Private Sub txtPrefsFontSize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gblEnablePrefsBalloonTooltips = "True" Then CreateToolTip txtPrefsFontSize.hWnd, "This is a read-only text box. It displays the current base font size as set when dynamic form resizing is enabled. The adjacent text box will display the automatically resized font currently in operation, for informational purposes only.", _
                  TTIconInfo, "Help on the Base Font Size", , , , True
End Sub

' Procedure : btnNow_Click
' Author    : beededea
' Date      : 26/11/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnNow_Click(Index As Integer)
    Dim dateAndTime As String: dateAndTime = vbNullString
    Dim alarmDateString As String: alarmDateString = vbNullString
    Dim alarmTimeString As String: alarmTimeString = vbNullString
    
    On Error GoTo btnNow_Click_Error

    dateAndTime = Now()
    
    alarmDateString = Left$(dateAndTime, 10)
    alarmTimeString = Mid$(dateAndTime, 12, 5)
    
    ' fill the comboboxes for the current time
    Select Case Index
        Case 1
            Call setAlarmDayValues(cmbAlarm1Day, alarmDateString)
            Call setAlarmMonthValues(cmbAlarm1Month, alarmDateString)
            Call setAlarmYearValues(cmbAlarm1Year, alarmDateString)
            Call setAlarmHourValues(cmbAlarm1Hours, alarmTimeString)
            Call setAlarmMinuteValues(cmbAlarm1Minutes, alarmTimeString)
            txtAlarm1Date.Text = "Alarm not yet set"
            txtAlarm1Time.Text = ""
        Case 2
            Call setAlarmDayValues(cmbAlarm2Day, alarmDateString)
            Call setAlarmMonthValues(cmbAlarm2Month, alarmDateString)
            Call setAlarmYearValues(cmbAlarm2Year, alarmDateString)
            Call setAlarmHourValues(cmbAlarm2Hours, alarmTimeString)
            Call setAlarmMinuteValues(cmbAlarm2Minutes, alarmTimeString)
            txtAlarm2Date.Text = "Alarm not yet set"
            txtAlarm2Time.Text = ""
        Case 3
            Call setAlarmDayValues(cmbAlarm3Day, alarmDateString)
            Call setAlarmMonthValues(cmbAlarm3Month, alarmDateString)
            Call setAlarmYearValues(cmbAlarm3Year, alarmDateString)
            Call setAlarmHourValues(cmbAlarm3Hours, alarmTimeString)
            Call setAlarmMinuteValues(cmbAlarm3Minutes, alarmTimeString)
            txtAlarm3Date.Text = "Alarm not yet set"
            txtAlarm1Time.Text = ""
        Case 4
            Call setAlarmDayValues(cmbAlarm4Day, alarmDateString)
            Call setAlarmMonthValues(cmbAlarm4Month, alarmDateString)
            Call setAlarmYearValues(cmbAlarm4Year, alarmDateString)
            Call setAlarmHourValues(cmbAlarm4Hours, alarmTimeString)
            Call setAlarmMinuteValues(cmbAlarm4Minutes, alarmTimeString)
            txtAlarm4Date.Text = "Alarm not yet set"
            txtAlarm4Time.Text = ""
        Case 5
            Call setAlarmDayValues(cmbAlarm5Day, alarmDateString)
            Call setAlarmMonthValues(cmbAlarm5Month, alarmDateString)
            Call setAlarmYearValues(cmbAlarm5Year, alarmDateString)
            Call setAlarmHourValues(cmbAlarm5Hours, alarmTimeString)
            Call setAlarmMinuteValues(cmbAlarm5Minutes, alarmTimeString)
            txtAlarm5Date.Text = "Alarm not yet set"
            txtAlarm5Time.Text = ""
    
    End Select
    
   On Error GoTo 0
   Exit Sub

btnNow_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnNow_Click of Form widgetPrefs"
    
End Sub



'---------------------------------------------------------------------------------------
' Procedure : btnclearAlarm_Click
' Author    : beededea
' Date      : 26/11/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnclearAlarm_Click(Index As Integer)
    On Error GoTo btnclearAlarm_Click_Error

    Select Case Index
        Case 1
            cmbAlarm1Day.ListIndex = 0
            cmbAlarm1Month.ListIndex = 0
            cmbAlarm1Year.ListIndex = 0
            cmbAlarm1Hours.ListIndex = 0
            cmbAlarm1Minutes.ListIndex = 0
            txtAlarm1Date.Text = "Alarm not yet set"
            txtAlarm1Time.Text = ""
        Case 2
            cmbAlarm2Day.ListIndex = 0
            cmbAlarm2Month.ListIndex = 0
            cmbAlarm2Year.ListIndex = 0
            cmbAlarm2Hours.ListIndex = 0
            cmbAlarm2Minutes.ListIndex = 0
            txtAlarm2Date.Text = "Alarm not yet set"
            txtAlarm2Time.Text = ""
        Case 3
            cmbAlarm3Day.ListIndex = 0
            cmbAlarm3Month.ListIndex = 0
            cmbAlarm3Year.ListIndex = 0
            cmbAlarm3Hours.ListIndex = 0
            cmbAlarm3Minutes.ListIndex = 0
            txtAlarm3Date.Text = "Alarm not yet set"
            txtAlarm3Time.Text = ""
            
        Case 4
            cmbAlarm4Day.ListIndex = 0
            cmbAlarm4Month.ListIndex = 0
            cmbAlarm4Year.ListIndex = 0
            cmbAlarm4Hours.ListIndex = 0
            cmbAlarm4Minutes.ListIndex = 0
            txtAlarm4Date.Text = "Alarm not yet set"
            txtAlarm4Time.Text = ""
            
        Case 5
            cmbAlarm5Day.ListIndex = 0
            cmbAlarm5Month.ListIndex = 0
            cmbAlarm5Year.ListIndex = 0
            cmbAlarm5Hours.ListIndex = 0
            cmbAlarm5Minutes.ListIndex = 0
            txtAlarm5Date.Text = "Alarm not yet set"
            txtAlarm5Time.Text = ""
            
        
    End Select
        

   On Error GoTo 0
   Exit Sub

btnclearAlarm_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnclearAlarm_Click of Form widgetPrefs"
End Sub


'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'  --- All folded content will be temporary put under this lines ---
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'CODEFOLD STORAGE:
'CODEFOLD STORAGE END:
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'--- If you're Subclassing: Move the CODEFOLD STORAGE up as needed ---
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\




