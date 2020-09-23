VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Universe File Editor"
   ClientHeight    =   6960
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab sstUniverse 
      Height          =   6855
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   13
      Tab             =   1
      TabsPerRow      =   13
      TabHeight       =   520
      TabCaption(0)   =   "Player"
      TabPicture(0)   =   "frmMain.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Objects"
      TabPicture(1)   =   "frmMain.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "scrObjects"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtObjectNum"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmbObjectExists"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Frame2"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Frame3"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Frame4"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmdObjectCopy"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cmdObjectDelete"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Races"
      TabPicture(2)   =   "frmMain.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Armour"
      TabPicture(3)   =   "frmMain.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Cannons"
      TabPicture(4)   =   "frmMain.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label24"
      Tab(4).Control(1)=   "Label25"
      Tab(4).Control(2)=   "Label26"
      Tab(4).Control(3)=   "Label27"
      Tab(4).Control(4)=   "Label28"
      Tab(4).Control(5)=   "Label29"
      Tab(4).Control(6)=   "Label30"
      Tab(4).Control(7)=   "Label31"
      Tab(4).Control(8)=   "Label32"
      Tab(4).Control(9)=   "Label33"
      Tab(4).Control(10)=   "Label34"
      Tab(4).Control(11)=   "Label41"
      Tab(4).Control(12)=   "Label105"
      Tab(4).Control(13)=   "txtCannonSpeed"
      Tab(4).Control(14)=   "txtCannonRadiation"
      Tab(4).Control(15)=   "txtCannonConcussive"
      Tab(4).Control(16)=   "txtCannonInstantaneous"
      Tab(4).Control(17)=   "txtCannonConsumption"
      Tab(4).Control(18)=   "txtCannonName"
      Tab(4).Control(19)=   "txtCannonNum"
      Tab(4).Control(20)=   "scrCannons"
      Tab(4).Control(21)=   "txtCannonSound"
      Tab(4).Control(22)=   "txtCannonFireRate"
      Tab(4).Control(23)=   "txtCannonMaxEnergy"
      Tab(4).Control(24)=   "txtCannonDuration"
      Tab(4).Control(25)=   "txtCannonSprite"
      Tab(4).Control(26)=   "txtCannonType"
      Tab(4).ControlCount=   27
      TabCaption(5)   =   "Engines"
      TabPicture(5)   =   "frmMain.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label35"
      Tab(5).Control(1)=   "Label36"
      Tab(5).Control(2)=   "Label37"
      Tab(5).Control(3)=   "Label38"
      Tab(5).Control(4)=   "Label39"
      Tab(5).Control(5)=   "Label40"
      Tab(5).Control(6)=   "scrEngines"
      Tab(5).Control(7)=   "txtEngineNum"
      Tab(5).Control(8)=   "txtEngineName"
      Tab(5).Control(9)=   "txtEngineConsumption"
      Tab(5).Control(10)=   "txtEngineThrust"
      Tab(5).Control(11)=   "txtEngineMaxEnergy"
      Tab(5).Control(12)=   "txtEngineSound"
      Tab(5).ControlCount=   13
      TabCaption(6)   =   "Generators"
      TabPicture(6)   =   "frmMain.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      TabCaption(7)   =   "Hulls"
      TabPicture(7)   =   "frmMain.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Label69"
      Tab(7).Control(1)=   "Label70"
      Tab(7).Control(2)=   "Label71"
      Tab(7).Control(3)=   "Label72"
      Tab(7).Control(4)=   "Label73"
      Tab(7).Control(5)=   "Label74"
      Tab(7).Control(6)=   "Label75"
      Tab(7).Control(7)=   "Label76"
      Tab(7).Control(8)=   "Label77"
      Tab(7).Control(9)=   "Label78"
      Tab(7).Control(10)=   "Label79"
      Tab(7).Control(11)=   "Label80"
      Tab(7).Control(12)=   "label231"
      Tab(7).Control(13)=   "Label82"
      Tab(7).Control(14)=   "Label83"
      Tab(7).Control(15)=   "Label84"
      Tab(7).Control(16)=   "Label85"
      Tab(7).Control(17)=   "Label86"
      Tab(7).Control(18)=   "Label87"
      Tab(7).Control(19)=   "Label88"
      Tab(7).Control(20)=   "Label89"
      Tab(7).Control(21)=   "Label90"
      Tab(7).Control(22)=   "Label81"
      Tab(7).Control(23)=   "Label91"
      Tab(7).Control(24)=   "Label92"
      Tab(7).Control(25)=   "Label93"
      Tab(7).Control(26)=   "Label95"
      Tab(7).Control(27)=   "txtHullSpriteFrameAmt"
      Tab(7).Control(28)=   "txtHullSpriteHeight"
      Tab(7).Control(29)=   "txtHullSpriteWidth"
      Tab(7).Control(30)=   "txtHullSprite"
      Tab(7).Control(31)=   "txtHullSpriteAnimAmt"
      Tab(7).Control(32)=   "txtHullMaxMines"
      Tab(7).Control(33)=   "txtHullMaxMissile"
      Tab(7).Control(34)=   "txtHullMaxCrew"
      Tab(7).Control(35)=   "txtHullMaxCargo"
      Tab(7).Control(36)=   "txtHullMass"
      Tab(7).Control(37)=   "txtHullName"
      Tab(7).Control(38)=   "txtHullNum"
      Tab(7).Control(39)=   "scrHulls"
      Tab(7).Control(40)=   "txtHullArmour"
      Tab(7).Control(41)=   "txtHullRotationRate"
      Tab(7).Control(42)=   "txtHullMaxSpeed"
      Tab(7).Control(43)=   "txtHullMaxFuel"
      Tab(7).Control(44)=   "txtHullCannon"
      Tab(7).Control(45)=   "txtHullEngine"
      Tab(7).Control(46)=   "txtHullGenerator"
      Tab(7).Control(47)=   "txtHullShield"
      Tab(7).Control(48)=   "txtHullMissile"
      Tab(7).Control(49)=   "txtHullLaser"
      Tab(7).Control(50)=   "cmbHullARCD"
      Tab(7).Control(51)=   "cmbHullFTLD"
      Tab(7).Control(52)=   "cmbHullMines"
      Tab(7).Control(53)=   "cmbHullJammer"
      Tab(7).Control(54)=   "txtHullSpriteAnimRate"
      Tab(7).ControlCount=   55
      TabCaption(8)   =   "Jammer"
      TabPicture(8)   =   "frmMain.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).ControlCount=   0
      TabCaption(9)   =   "Lasers"
      TabPicture(9)   =   "frmMain.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Label42"
      Tab(9).Control(1)=   "Label43"
      Tab(9).Control(2)=   "Label44"
      Tab(9).Control(3)=   "Label46"
      Tab(9).Control(4)=   "Label47"
      Tab(9).Control(5)=   "Label48"
      Tab(9).Control(6)=   "Label49"
      Tab(9).Control(7)=   "Label50"
      Tab(9).Control(8)=   "Label51"
      Tab(9).Control(9)=   "Label52"
      Tab(9).Control(10)=   "txtLaserMaxEnergy"
      Tab(9).Control(11)=   "txtLaserColour"
      Tab(9).Control(12)=   "txtLaserSound"
      Tab(9).Control(13)=   "scrLasers"
      Tab(9).Control(14)=   "txtLaserNum"
      Tab(9).Control(15)=   "txtLaserName"
      Tab(9).Control(16)=   "txtLaserConsumption"
      Tab(9).Control(17)=   "txtLaserFireConsumption"
      Tab(9).Control(18)=   "txtLaserConcussive"
      Tab(9).Control(19)=   "txtLaserRadiation"
      Tab(9).Control(20)=   "txtLaserRange"
      Tab(9).ControlCount=   21
      TabCaption(10)  =   "Missiles"
      TabPicture(10)  =   "frmMain.frx":0118
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Label45"
      Tab(10).Control(1)=   "Label53"
      Tab(10).Control(2)=   "Label54"
      Tab(10).Control(3)=   "Label55"
      Tab(10).Control(4)=   "Label56"
      Tab(10).Control(5)=   "Label57"
      Tab(10).Control(6)=   "Label58"
      Tab(10).Control(7)=   "Label59"
      Tab(10).Control(8)=   "Label60"
      Tab(10).Control(9)=   "Label61"
      Tab(10).Control(10)=   "Label62"
      Tab(10).Control(11)=   "Label63"
      Tab(10).Control(12)=   "Label64"
      Tab(10).Control(13)=   "Label65"
      Tab(10).Control(14)=   "Label66"
      Tab(10).Control(15)=   "Label67"
      Tab(10).Control(16)=   "Label68"
      Tab(10).Control(17)=   "Label94"
      Tab(10).Control(18)=   "txtMissileSound"
      Tab(10).Control(19)=   "txtMissileThrust"
      Tab(10).Control(20)=   "txtMissileDuration"
      Tab(10).Control(21)=   "txtMissileTargetBias"
      Tab(10).Control(22)=   "txtMissileSeekDist"
      Tab(10).Control(23)=   "scrMissiles"
      Tab(10).Control(24)=   "txtMissileNum"
      Tab(10).Control(25)=   "txtMissileName"
      Tab(10).Control(26)=   "txtMissileConcussive"
      Tab(10).Control(27)=   "txtMissileRadiation"
      Tab(10).Control(28)=   "txtMissileRotationRate"
      Tab(10).Control(29)=   "txtMissileFireRate"
      Tab(10).Control(30)=   "txtMissileMaxSpeed"
      Tab(10).Control(31)=   "txtMissileSpriteAnimAmt"
      Tab(10).Control(32)=   "txtMissileSprite"
      Tab(10).Control(33)=   "txtMissileSpriteWidth"
      Tab(10).Control(34)=   "txtMissileSpriteHeight"
      Tab(10).Control(35)=   "txtMissileSpriteFrameAmt"
      Tab(10).Control(36)=   "txtMissileSpriteAnimRate"
      Tab(10).ControlCount=   37
      TabCaption(11)  =   "Scanners"
      TabPicture(11)  =   "frmMain.frx":0134
      Tab(11).ControlEnabled=   0   'False
      Tab(11).ControlCount=   0
      TabCaption(12)  =   "Shields"
      TabPicture(12)  =   "frmMain.frx":0150
      Tab(12).ControlEnabled=   0   'False
      Tab(12).ControlCount=   0
      Begin VB.CommandButton cmdObjectDelete 
         Caption         =   "Delete"
         Height          =   315
         Left            =   7800
         TabIndex        =   224
         Top             =   750
         Width           =   1215
      End
      Begin VB.CommandButton cmdObjectCopy 
         Caption         =   "&Copy"
         Height          =   315
         Left            =   9090
         TabIndex        =   223
         Top             =   750
         Width           =   1215
      End
      Begin VB.TextBox txtCannonType 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   221
         Top             =   4800
         Width           =   1755
      End
      Begin VB.Frame Frame4 
         Caption         =   "Systems"
         Height          =   3345
         Left            =   7080
         TabIndex        =   202
         Top             =   1110
         Width           =   3255
         Begin VB.ComboBox cmbObjectScanner 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   218
            Top             =   2550
            Width           =   1755
         End
         Begin VB.ComboBox cmbObjectShield 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   217
            Top             =   2880
            Width           =   1755
         End
         Begin VB.ComboBox cmbObjectArmour 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   209
            Top             =   210
            Width           =   1755
         End
         Begin VB.ComboBox cmbObjectCannon 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   208
            Top             =   570
            Width           =   1755
         End
         Begin VB.ComboBox cmbObjectEngine 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   207
            Top             =   900
            Width           =   1755
         End
         Begin VB.ComboBox cmbObjectGenerator 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   206
            Top             =   1230
            Width           =   1755
         End
         Begin VB.ComboBox cmbObjectHull 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   205
            Top             =   1560
            Width           =   1755
         End
         Begin VB.ComboBox cmbObjectLaser 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   204
            Top             =   1890
            Width           =   1755
         End
         Begin VB.ComboBox cmbObjectMissile 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   203
            Top             =   2220
            Width           =   1755
         End
         Begin VB.Label Label104 
            Caption         =   "Scanner:"
            Height          =   195
            Left            =   150
            TabIndex        =   220
            Top             =   2640
            Width           =   1095
         End
         Begin VB.Label Label103 
            Caption         =   "Sheild:"
            Height          =   195
            Left            =   150
            TabIndex        =   219
            Top             =   2970
            Width           =   1095
         End
         Begin VB.Label Label102 
            Caption         =   "Armour:"
            Height          =   195
            Left            =   150
            TabIndex        =   216
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label Label101 
            Caption         =   "Cannon:"
            Height          =   195
            Left            =   150
            TabIndex        =   215
            Top             =   660
            Width           =   1095
         End
         Begin VB.Label Label100 
            Caption         =   "Engine:"
            Height          =   195
            Left            =   150
            TabIndex        =   214
            Top             =   990
            Width           =   1095
         End
         Begin VB.Label Label99 
            Caption         =   "Generator:"
            Height          =   195
            Left            =   150
            TabIndex        =   213
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label98 
            Caption         =   "Hull:"
            Height          =   195
            Left            =   150
            TabIndex        =   212
            Top             =   1650
            Width           =   1095
         End
         Begin VB.Label Label97 
            Caption         =   "Laser:"
            Height          =   195
            Left            =   150
            TabIndex        =   211
            Top             =   1980
            Width           =   1095
         End
         Begin VB.Label Label96 
            Caption         =   "Missile:"
            Height          =   195
            Left            =   150
            TabIndex        =   210
            Top             =   2310
            Width           =   1095
         End
      End
      Begin VB.TextBox txtHullSpriteAnimRate 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -70560
         TabIndex        =   200
         Top             =   2820
         Width           =   1755
      End
      Begin VB.TextBox txtMissileSpriteAnimRate 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -70560
         TabIndex        =   198
         Top             =   2820
         Width           =   1755
      End
      Begin VB.ComboBox cmbHullJammer 
         Height          =   315
         Left            =   -70530
         Style           =   2  'Dropdown List
         TabIndex        =   193
         Top             =   3480
         Width           =   1755
      End
      Begin VB.ComboBox cmbHullMines 
         Height          =   315
         Left            =   -70530
         Style           =   2  'Dropdown List
         TabIndex        =   192
         Top             =   3840
         Width           =   1755
      End
      Begin VB.ComboBox cmbHullFTLD 
         Height          =   315
         Left            =   -70530
         Style           =   2  'Dropdown List
         TabIndex        =   191
         Top             =   4170
         Width           =   1755
      End
      Begin VB.ComboBox cmbHullARCD 
         Height          =   315
         Left            =   -70530
         Style           =   2  'Dropdown List
         TabIndex        =   190
         Top             =   4500
         Width           =   1755
      End
      Begin VB.TextBox txtHullLaser 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   184
         Top             =   6120
         Width           =   1755
      End
      Begin VB.TextBox txtHullMissile 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   183
         Top             =   4800
         Width           =   1755
      End
      Begin VB.TextBox txtHullShield 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   182
         Top             =   5130
         Width           =   1755
      End
      Begin VB.TextBox txtHullGenerator 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   181
         Top             =   5460
         Width           =   1755
      End
      Begin VB.TextBox txtHullEngine 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   180
         Top             =   5790
         Width           =   1755
      End
      Begin VB.TextBox txtHullCannon 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   162
         Top             =   4470
         Width           =   1755
      End
      Begin VB.TextBox txtHullMaxFuel 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   161
         Top             =   3150
         Width           =   1755
      End
      Begin VB.TextBox txtHullMaxSpeed 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   160
         Top             =   3480
         Width           =   1755
      End
      Begin VB.TextBox txtHullRotationRate 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   159
         Top             =   3810
         Width           =   1755
      End
      Begin VB.TextBox txtHullArmour 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   158
         Top             =   4140
         Width           =   1755
      End
      Begin VB.HScrollBar scrHulls 
         Height          =   225
         Left            =   -74880
         Max             =   0
         TabIndex        =   157
         Top             =   420
         Width           =   10695
      End
      Begin VB.TextBox txtHullNum 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   156
         Text            =   "0"
         Top             =   750
         Width           =   1755
      End
      Begin VB.TextBox txtHullName 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   155
         Top             =   1170
         Width           =   1755
      End
      Begin VB.TextBox txtHullMass 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   154
         Top             =   1500
         Width           =   1755
      End
      Begin VB.TextBox txtHullMaxCargo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   153
         Top             =   1830
         Width           =   1755
      End
      Begin VB.TextBox txtHullMaxCrew 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   152
         Top             =   2160
         Width           =   1755
      End
      Begin VB.TextBox txtHullMaxMissile 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   151
         Top             =   2490
         Width           =   1755
      End
      Begin VB.TextBox txtHullMaxMines 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   150
         Top             =   2820
         Width           =   1755
      End
      Begin VB.TextBox txtHullSpriteAnimAmt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -70560
         TabIndex        =   149
         Top             =   2490
         Width           =   1755
      End
      Begin VB.TextBox txtHullSprite 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -70560
         TabIndex        =   148
         Top             =   1170
         Width           =   1755
      End
      Begin VB.TextBox txtHullSpriteWidth 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -70560
         TabIndex        =   147
         Top             =   1500
         Width           =   1755
      End
      Begin VB.TextBox txtHullSpriteHeight 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -70560
         TabIndex        =   146
         Top             =   1830
         Width           =   1755
      End
      Begin VB.TextBox txtHullSpriteFrameAmt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -70560
         TabIndex        =   145
         Top             =   2160
         Width           =   1755
      End
      Begin VB.TextBox txtMissileSpriteFrameAmt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -70560
         TabIndex        =   139
         Top             =   2160
         Width           =   1755
      End
      Begin VB.TextBox txtMissileSpriteHeight 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -70560
         TabIndex        =   138
         Top             =   1830
         Width           =   1755
      End
      Begin VB.TextBox txtMissileSpriteWidth 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -70560
         TabIndex        =   137
         Top             =   1500
         Width           =   1755
      End
      Begin VB.TextBox txtMissileSprite 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -70560
         TabIndex        =   136
         Top             =   1170
         Width           =   1755
      End
      Begin VB.TextBox txtMissileSpriteAnimAmt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -70560
         TabIndex        =   135
         Top             =   2490
         Width           =   1755
      End
      Begin VB.TextBox txtMissileMaxSpeed 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   122
         Top             =   2820
         Width           =   1755
      End
      Begin VB.TextBox txtMissileFireRate 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   121
         Top             =   2490
         Width           =   1755
      End
      Begin VB.TextBox txtMissileRotationRate 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   120
         Top             =   2160
         Width           =   1755
      End
      Begin VB.TextBox txtMissileRadiation 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   119
         Top             =   1830
         Width           =   1755
      End
      Begin VB.TextBox txtMissileConcussive 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   118
         Top             =   1500
         Width           =   1755
      End
      Begin VB.TextBox txtMissileName 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   117
         Top             =   1170
         Width           =   1755
      End
      Begin VB.TextBox txtMissileNum 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   116
         Text            =   "0"
         Top             =   750
         Width           =   1755
      End
      Begin VB.HScrollBar scrMissiles 
         Height          =   225
         Left            =   -74880
         Max             =   0
         TabIndex        =   115
         Top             =   420
         Width           =   10695
      End
      Begin VB.TextBox txtMissileSeekDist 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   114
         Top             =   4140
         Width           =   1755
      End
      Begin VB.TextBox txtMissileTargetBias 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   113
         Top             =   3810
         Width           =   1755
      End
      Begin VB.TextBox txtMissileDuration 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   112
         Top             =   3480
         Width           =   1755
      End
      Begin VB.TextBox txtMissileThrust 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   111
         Top             =   3150
         Width           =   1755
      End
      Begin VB.TextBox txtMissileSound 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   110
         Top             =   4470
         Width           =   1755
      End
      Begin VB.TextBox txtLaserRange 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   99
         Top             =   2820
         Width           =   1755
      End
      Begin VB.TextBox txtLaserRadiation 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   98
         Top             =   2490
         Width           =   1755
      End
      Begin VB.TextBox txtLaserConcussive 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   97
         Top             =   2160
         Width           =   1755
      End
      Begin VB.TextBox txtLaserFireConsumption 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   96
         Top             =   1830
         Width           =   1755
      End
      Begin VB.TextBox txtLaserConsumption 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   95
         Top             =   1500
         Width           =   1755
      End
      Begin VB.TextBox txtLaserName 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   94
         Top             =   1170
         Width           =   1755
      End
      Begin VB.TextBox txtLaserNum 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   93
         Text            =   "0"
         Top             =   750
         Width           =   1755
      End
      Begin VB.HScrollBar scrLasers 
         Height          =   225
         Left            =   -74880
         Max             =   0
         TabIndex        =   92
         Top             =   420
         Width           =   10695
      End
      Begin VB.TextBox txtLaserSound 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   91
         Top             =   3810
         Width           =   1755
      End
      Begin VB.TextBox txtLaserColour 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   90
         Top             =   3480
         Width           =   1755
      End
      Begin VB.TextBox txtLaserMaxEnergy 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   89
         Top             =   3150
         Width           =   1755
      End
      Begin VB.TextBox txtCannonSprite 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   87
         Top             =   4470
         Width           =   1755
      End
      Begin VB.TextBox txtEngineSound 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   80
         Top             =   2490
         Width           =   1755
      End
      Begin VB.TextBox txtEngineMaxEnergy 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   79
         Top             =   2160
         Width           =   1755
      End
      Begin VB.TextBox txtEngineThrust 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   78
         Top             =   1830
         Width           =   1755
      End
      Begin VB.TextBox txtEngineConsumption 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   77
         Top             =   1500
         Width           =   1755
      End
      Begin VB.TextBox txtEngineName 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   76
         Top             =   1170
         Width           =   1755
      End
      Begin VB.TextBox txtEngineNum 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   75
         Text            =   "0"
         Top             =   750
         Width           =   1755
      End
      Begin VB.HScrollBar scrEngines 
         Height          =   225
         Left            =   -74880
         Max             =   0
         TabIndex        =   74
         Top             =   420
         Width           =   10695
      End
      Begin VB.TextBox txtCannonDuration 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   69
         Top             =   3150
         Width           =   1755
      End
      Begin VB.TextBox txtCannonMaxEnergy 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   68
         Top             =   3480
         Width           =   1755
      End
      Begin VB.TextBox txtCannonFireRate 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   67
         Top             =   3810
         Width           =   1755
      End
      Begin VB.TextBox txtCannonSound 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   66
         Top             =   4140
         Width           =   1755
      End
      Begin VB.HScrollBar scrCannons 
         Height          =   225
         Left            =   -74880
         Max             =   0
         TabIndex        =   64
         Top             =   420
         Width           =   10695
      End
      Begin VB.TextBox txtCannonNum 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   63
         Text            =   "0"
         Top             =   750
         Width           =   1755
      End
      Begin VB.TextBox txtCannonName 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   56
         Top             =   1170
         Width           =   1755
      End
      Begin VB.TextBox txtCannonConsumption 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   55
         Top             =   1500
         Width           =   1755
      End
      Begin VB.TextBox txtCannonInstantaneous 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   54
         Top             =   1830
         Width           =   1755
      End
      Begin VB.TextBox txtCannonConcussive 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   53
         Top             =   2160
         Width           =   1755
      End
      Begin VB.TextBox txtCannonRadiation 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   52
         Top             =   2490
         Width           =   1755
      End
      Begin VB.TextBox txtCannonSpeed 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73650
         TabIndex        =   51
         Top             =   2820
         Width           =   1755
      End
      Begin VB.Frame Frame3 
         Caption         =   "AI"
         Height          =   2685
         Left            =   3630
         TabIndex        =   36
         Top             =   1110
         Width           =   3255
         Begin VB.TextBox txtObjectLengthTargetLock 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1350
            TabIndex        =   49
            Top             =   2220
            Width           =   1755
         End
         Begin VB.ComboBox cmbObjectAction 
            Height          =   315
            Left            =   1350
            TabIndex        =   48
            Top             =   240
            Width           =   1755
         End
         Begin VB.TextBox txtObjectTargetBias 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1350
            TabIndex        =   41
            Top             =   570
            Width           =   1755
         End
         Begin VB.TextBox txtObjectSeekDist 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1350
            TabIndex        =   40
            Top             =   900
            Width           =   1755
         End
         Begin VB.TextBox txtObjectMinDist 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1350
            TabIndex        =   39
            Top             =   1230
            Width           =   1755
         End
         Begin VB.TextBox txtObjectCannonDist 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1350
            TabIndex        =   38
            Top             =   1560
            Width           =   1755
         End
         Begin VB.TextBox txtObjectAimTolerance 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1350
            TabIndex        =   37
            Top             =   1890
            Width           =   1755
         End
         Begin VB.Label Label23 
            Caption         =   "Lock Duration:"
            Height          =   195
            Left            =   180
            TabIndex        =   50
            Top             =   2280
            Width           =   1095
         End
         Begin VB.Label Label22 
            Caption         =   "Action:"
            Height          =   195
            Left            =   180
            TabIndex        =   47
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label Label21 
            Caption         =   "Target Bias:"
            Height          =   195
            Left            =   180
            TabIndex        =   46
            Top             =   630
            Width           =   1095
         End
         Begin VB.Label Label20 
            Caption         =   "Seek Dist:"
            Height          =   195
            Left            =   180
            TabIndex        =   45
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label19 
            Caption         =   "Min Dist:"
            Height          =   195
            Left            =   180
            TabIndex        =   44
            Top             =   1290
            Width           =   1095
         End
         Begin VB.Label Label18 
            Caption         =   "Cannon Dist:"
            Height          =   195
            Left            =   180
            TabIndex        =   43
            Top             =   1620
            Width           =   1095
         End
         Begin VB.Label Label17 
            Caption         =   "Aim Tolerance:"
            Height          =   195
            Left            =   180
            TabIndex        =   42
            Top             =   1950
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Physics"
         Height          =   2355
         Left            =   180
         TabIndex        =   23
         Top             =   4290
         Width           =   3255
         Begin VB.TextBox txtObjectMass 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1350
            TabIndex        =   29
            Top             =   1890
            Width           =   1755
         End
         Begin VB.TextBox txtObjectSpeed 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1350
            TabIndex        =   28
            Top             =   1560
            Width           =   1755
         End
         Begin VB.TextBox txtObjectHeading 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1350
            TabIndex        =   27
            Top             =   1230
            Width           =   1755
         End
         Begin VB.TextBox txtObjectFacing 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1350
            TabIndex        =   26
            Top             =   900
            Width           =   1755
         End
         Begin VB.TextBox txtObjectYCoord 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1350
            TabIndex        =   25
            Top             =   570
            Width           =   1755
         End
         Begin VB.TextBox txtObjectXCoord 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1350
            TabIndex        =   24
            Top             =   240
            Width           =   1755
         End
         Begin VB.Label Label16 
            Caption         =   "Mass:"
            Height          =   195
            Left            =   180
            TabIndex        =   35
            Top             =   1950
            Width           =   1095
         End
         Begin VB.Label Label15 
            Caption         =   "Speed:"
            Height          =   195
            Left            =   180
            TabIndex        =   34
            Top             =   1620
            Width           =   1095
         End
         Begin VB.Label Label14 
            Caption         =   "Heading:"
            Height          =   195
            Left            =   180
            TabIndex        =   33
            Top             =   1290
            Width           =   1095
         End
         Begin VB.Label Label13 
            Caption         =   "Facing:"
            Height          =   195
            Left            =   180
            TabIndex        =   32
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label12 
            Caption         =   "Y Coord:"
            Height          =   195
            Left            =   180
            TabIndex        =   31
            Top             =   630
            Width           =   1095
         End
         Begin VB.Label Label11 
            Caption         =   "X Coord:"
            Height          =   195
            Left            =   180
            TabIndex        =   30
            Top             =   300
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Info"
         Height          =   3075
         Left            =   180
         TabIndex        =   6
         Top             =   1110
         Width           =   3255
         Begin VB.ComboBox cmbObjectStarbase 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   2610
            Width           =   1755
         End
         Begin VB.ComboBox cmbObjectStar 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   2280
            Width           =   1755
         End
         Begin VB.ComboBox cmbObjectPlanet 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   1950
            Width           =   1755
         End
         Begin VB.ComboBox cmbObjectFighter 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1620
            Width           =   1755
         End
         Begin VB.ComboBox cmbObjectCarrier 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1290
            Width           =   1755
         End
         Begin VB.ComboBox cmbObjectCanMove 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   960
            Width           =   1755
         End
         Begin VB.ComboBox cmbObjectRace 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   600
            Width           =   1755
         End
         Begin VB.TextBox txtObjectName 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1350
            TabIndex        =   7
            Top             =   270
            Width           =   1755
         End
         Begin VB.Label Label10 
            Caption         =   "Starbase:"
            Height          =   195
            Left            =   180
            TabIndex        =   22
            Top             =   2700
            Width           =   1095
         End
         Begin VB.Label Label9 
            Caption         =   "Star:"
            Height          =   195
            Left            =   180
            TabIndex        =   21
            Top             =   2370
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "Planet:"
            Height          =   195
            Left            =   180
            TabIndex        =   20
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "Fighter:"
            Height          =   195
            Left            =   180
            TabIndex        =   19
            Top             =   1710
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Carrier:"
            Height          =   195
            Left            =   180
            TabIndex        =   18
            Top             =   1380
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "Can Move:"
            Height          =   195
            Left            =   180
            TabIndex        =   17
            Top             =   1050
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Race:"
            Height          =   195
            Left            =   180
            TabIndex        =   16
            Top             =   690
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Name:"
            Height          =   195
            Left            =   180
            TabIndex        =   15
            Top             =   330
            Width           =   1095
         End
      End
      Begin VB.ComboBox cmbObjectExists 
         Height          =   315
         Left            =   4740
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   720
         Width           =   1485
      End
      Begin VB.TextBox txtObjectNum 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1530
         TabIndex        =   3
         Text            =   "0"
         Top             =   750
         Width           =   1755
      End
      Begin VB.HScrollBar scrObjects 
         Height          =   225
         Left            =   120
         Max             =   0
         TabIndex        =   1
         Top             =   420
         Width           =   10695
      End
      Begin VB.Label Label105 
         Caption         =   "Cannon Type:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   222
         Top             =   4860
         Width           =   1095
      End
      Begin VB.Label Label95 
         Caption         =   "Anim Rate:"
         Height          =   195
         Left            =   -71730
         TabIndex        =   201
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label94 
         Caption         =   "Anim Rate:"
         Height          =   195
         Left            =   -71730
         TabIndex        =   199
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label93 
         Caption         =   "Comm Jammer:"
         Height          =   195
         Left            =   -71700
         TabIndex        =   197
         Top             =   3570
         Width           =   1095
      End
      Begin VB.Label Label92 
         Caption         =   "Mines:"
         Height          =   195
         Left            =   -71700
         TabIndex        =   196
         Top             =   3930
         Width           =   1095
      End
      Begin VB.Label Label91 
         Caption         =   "FTLD:"
         Height          =   195
         Left            =   -71700
         TabIndex        =   195
         Top             =   4260
         Width           =   1095
      End
      Begin VB.Label Label81 
         Caption         =   "ARCD:"
         Height          =   195
         Left            =   -71700
         TabIndex        =   194
         Top             =   4590
         Width           =   1095
      End
      Begin VB.Label Label90 
         Caption         =   "Lasers:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   189
         Top             =   6180
         Width           =   1095
      End
      Begin VB.Label Label89 
         Caption         =   "Missiles:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   188
         Top             =   4860
         Width           =   1095
      End
      Begin VB.Label Label88 
         Caption         =   "Sheilds:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   187
         Top             =   5190
         Width           =   1095
      End
      Begin VB.Label Label87 
         Caption         =   "Generators:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   186
         Top             =   5520
         Width           =   1095
      End
      Begin VB.Label Label86 
         Caption         =   "Engines:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   185
         Top             =   5850
         Width           =   1095
      End
      Begin VB.Label Label85 
         Caption         =   "Cannons:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   179
         Top             =   4530
         Width           =   1095
      End
      Begin VB.Label Label84 
         Caption         =   "Max Fuel:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   178
         Top             =   3210
         Width           =   1095
      End
      Begin VB.Label Label83 
         Caption         =   "Max Speed:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   177
         Top             =   3540
         Width           =   1095
      End
      Begin VB.Label Label82 
         Caption         =   "Rotation Rate:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   176
         Top             =   3870
         Width           =   1095
      End
      Begin VB.Label label231 
         Caption         =   "Armour:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   175
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Label80 
         Caption         =   "Hull Num:"
         Height          =   255
         Left            =   -74820
         TabIndex        =   174
         Top             =   780
         Width           =   1005
      End
      Begin VB.Label Label79 
         Caption         =   "Name:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   173
         Top             =   1230
         Width           =   1095
      End
      Begin VB.Label Label78 
         Caption         =   "Mass:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   172
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label77 
         Caption         =   "Max Cargo:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   171
         Top             =   1890
         Width           =   1095
      End
      Begin VB.Label Label76 
         Caption         =   "Max Crew:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   170
         Top             =   2220
         Width           =   1095
      End
      Begin VB.Label Label75 
         Caption         =   "Max Missile:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   169
         Top             =   2550
         Width           =   1095
      End
      Begin VB.Label Label74 
         Caption         =   "Max Mines:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   168
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label73 
         Caption         =   "Anim Amount:"
         Height          =   195
         Left            =   -71730
         TabIndex        =   167
         Top             =   2550
         Width           =   1095
      End
      Begin VB.Label Label72 
         Caption         =   "Sprite Name:"
         Height          =   195
         Left            =   -71730
         TabIndex        =   166
         Top             =   1230
         Width           =   1095
      End
      Begin VB.Label Label71 
         Caption         =   "Sprite Width:"
         Height          =   195
         Left            =   -71730
         TabIndex        =   165
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label70 
         Caption         =   "Sprite Height:"
         Height          =   195
         Left            =   -71730
         TabIndex        =   164
         Top             =   1890
         Width           =   1095
      End
      Begin VB.Label Label69 
         Caption         =   "Frame Amount:"
         Height          =   195
         Left            =   -71730
         TabIndex        =   163
         Top             =   2220
         Width           =   1095
      End
      Begin VB.Label Label68 
         Caption         =   "Frame Amount:"
         Height          =   195
         Left            =   -71730
         TabIndex        =   144
         Top             =   2220
         Width           =   1095
      End
      Begin VB.Label Label67 
         Caption         =   "Sprite Height:"
         Height          =   195
         Left            =   -71730
         TabIndex        =   143
         Top             =   1890
         Width           =   1095
      End
      Begin VB.Label Label66 
         Caption         =   "Sprite Width:"
         Height          =   195
         Left            =   -71730
         TabIndex        =   142
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label65 
         Caption         =   "Sprite Name:"
         Height          =   195
         Left            =   -71730
         TabIndex        =   141
         Top             =   1230
         Width           =   1095
      End
      Begin VB.Label Label64 
         Caption         =   "Anim Amount:"
         Height          =   195
         Left            =   -71730
         TabIndex        =   140
         Top             =   2550
         Width           =   1095
      End
      Begin VB.Label Label63 
         Caption         =   "Max Speed:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   134
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label62 
         Caption         =   "Fire Rate:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   133
         Top             =   2550
         Width           =   1095
      End
      Begin VB.Label Label61 
         Caption         =   "Rotation Rate:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   132
         Top             =   2220
         Width           =   1095
      End
      Begin VB.Label Label60 
         Caption         =   "Radiation:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   131
         Top             =   1890
         Width           =   1095
      End
      Begin VB.Label Label59 
         Caption         =   "Concussive:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   130
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label58 
         Caption         =   "Name:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   129
         Top             =   1230
         Width           =   1095
      End
      Begin VB.Label Label57 
         Caption         =   "Missile Num:"
         Height          =   255
         Left            =   -74820
         TabIndex        =   128
         Top             =   780
         Width           =   1005
      End
      Begin VB.Label Label56 
         Caption         =   "Seek Dist:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   127
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Label55 
         Caption         =   "Target Bias:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   126
         Top             =   3870
         Width           =   1095
      End
      Begin VB.Label Label54 
         Caption         =   "Duration:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   125
         Top             =   3540
         Width           =   1095
      End
      Begin VB.Label Label53 
         Caption         =   "Thrust:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   124
         Top             =   3210
         Width           =   1095
      End
      Begin VB.Label Label45 
         Caption         =   "Sound File:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   123
         Top             =   4530
         Width           =   1095
      End
      Begin VB.Label Label52 
         Caption         =   "Range:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   109
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label51 
         Caption         =   "Radiation:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   108
         Top             =   2550
         Width           =   1095
      End
      Begin VB.Label Label50 
         Caption         =   "Concussive:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   107
         Top             =   2220
         Width           =   1095
      End
      Begin VB.Label Label49 
         Caption         =   "F Consumption:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   106
         Top             =   1890
         Width           =   1095
      End
      Begin VB.Label Label48 
         Caption         =   "Consumption:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   105
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label47 
         Caption         =   "Name:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   104
         Top             =   1230
         Width           =   1095
      End
      Begin VB.Label Label46 
         Caption         =   "Laser Num:"
         Height          =   255
         Left            =   -74820
         TabIndex        =   103
         Top             =   780
         Width           =   1005
      End
      Begin VB.Label Label44 
         Caption         =   "Sound File:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   102
         Top             =   3870
         Width           =   1095
      End
      Begin VB.Label Label43 
         Caption         =   "Colour:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   101
         Top             =   3540
         Width           =   1095
      End
      Begin VB.Label Label42 
         Caption         =   "Max Energy:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   100
         Top             =   3210
         Width           =   1095
      End
      Begin VB.Label Label41 
         Caption         =   "Sprite Name:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   88
         Top             =   4530
         Width           =   1095
      End
      Begin VB.Label Label40 
         Caption         =   "Sound File:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   86
         Top             =   2550
         Width           =   1095
      End
      Begin VB.Label Label39 
         Caption         =   "Max Energy:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   85
         Top             =   2220
         Width           =   1095
      End
      Begin VB.Label Label38 
         Caption         =   "Thrust:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   84
         Top             =   1890
         Width           =   1095
      End
      Begin VB.Label Label37 
         Caption         =   "Consumption:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   83
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label36 
         Caption         =   "Name:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   82
         Top             =   1230
         Width           =   1095
      End
      Begin VB.Label Label35 
         Caption         =   "Engine Num:"
         Height          =   255
         Left            =   -74820
         TabIndex        =   81
         Top             =   780
         Width           =   1005
      End
      Begin VB.Label Label34 
         Caption         =   "Duration:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   73
         Top             =   3210
         Width           =   1095
      End
      Begin VB.Label Label33 
         Caption         =   "Max Energy:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   72
         Top             =   3540
         Width           =   1095
      End
      Begin VB.Label Label32 
         Caption         =   "Fire Rate:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   71
         Top             =   3870
         Width           =   1095
      End
      Begin VB.Label Label31 
         Caption         =   "Sound File:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   70
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Label30 
         Caption         =   "Cannon Num:"
         Height          =   255
         Left            =   -74820
         TabIndex        =   65
         Top             =   780
         Width           =   1005
      End
      Begin VB.Label Label29 
         Caption         =   "Name:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   62
         Top             =   1230
         Width           =   1095
      End
      Begin VB.Label Label28 
         Caption         =   "Consumption:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   61
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label27 
         Caption         =   "Instantaneous:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   60
         Top             =   1890
         Width           =   1095
      End
      Begin VB.Label Label26 
         Caption         =   "Concussive:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   59
         Top             =   2220
         Width           =   1095
      End
      Begin VB.Label Label25 
         Caption         =   "Radiation:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   58
         Top             =   2550
         Width           =   1095
      End
      Begin VB.Label Label24 
         Caption         =   "Speed:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   57
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Object Exists:"
         Height          =   225
         Left            =   3660
         TabIndex        =   4
         Top             =   780
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Object Num:"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   780
         Width           =   945
      End
   End
   Begin MSComDlg.CommonDialog cdlUniverse 
      Left            =   8610
      Top             =   7140
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'TODO:
' - Need to add some kind of integrity check, to make sure people aren't
'   tampering with the universe files (sum of all the bits?)

Option Explicit

Private Sub cmbHullARCD_Click()

    'Update hull ARCD
    gudtHull(scrHulls.Value).blnARCD = GetBoolCombo(cmbHullARCD)

End Sub

Private Sub cmbHullFTLD_Click()

    'Update hull FTLD
    gudtHull(scrHulls.Value).blnFTLD = GetBoolCombo(cmbHullFTLD)

End Sub

Private Sub cmbHullJammer_Click()

    'Update hull jammer
    gudtHull(scrHulls.Value).blnCommJammer = GetBoolCombo(cmbHullJammer)

End Sub

Private Sub cmbHullMines_Click()

    'Update hull mines
    gudtHull(scrHulls.Value).blnMines = GetBoolCombo(cmbHullMines)

End Sub

Private Sub cmbObjectAction_Click()

    'The object's action value has changed
    gudtObject(scrObjects.Value).udtAI.bytAction = cmbObjectAction.ListIndex

End Sub

Private Sub cmbObjectArmour_Click()

    'Update armour system
    gudtObject(scrObjects.Value).udtSystems.bytArmour = cmbObjectArmour.ListIndex

End Sub

Private Sub cmbObjectCanMove_Click()

    'The object's Can Move value has changed
    gudtObject(scrObjects.Value).udtInfo.blnCanMove = GetBoolCombo(cmbObjectCanMove)

End Sub

Private Sub cmbObjectCannon_Click()

    'Update cannon system
    gudtObject(scrObjects.Value).udtSystems.bytCannon = cmbObjectCannon.ListIndex

End Sub

Private Sub cmbObjectCarrier_Click()

    'The object's carrier value has changed
    gudtObject(scrObjects.Value).udtInfo.blnCarrier = GetBoolCombo(cmbObjectCarrier)

End Sub

Private Sub cmbObjectEngine_Click()

    'Update engine system
    gudtObject(scrObjects.Value).udtSystems.bytEngine = cmbObjectEngine.ListIndex

End Sub

Private Sub cmbObjectExists_Click()

    'The object's existence has been modified
    gudtObject(scrObjects.Value).blnExists = GetBoolCombo(cmbObjectExists)
    
End Sub

Private Sub cmbObjectFighter_Click()

    'The object's fighter value has changed
    gudtObject(scrObjects.Value).udtInfo.blnFighter = GetBoolCombo(cmbObjectFighter)

End Sub

Private Sub cmbObjectGenerator_Click()

    'Update generator system
    gudtObject(scrObjects.Value).udtSystems.bytGenerator = cmbObjectGenerator.ListIndex

End Sub

Private Sub cmbObjectHull_Click()

    'Update hull system
    gudtObject(scrObjects.Value).udtSystems.bytHull = cmbObjectHull.ListIndex

End Sub

Private Sub cmbObjectLaser_Click()

    'Update laser system
    gudtObject(scrObjects.Value).udtSystems.bytLaser = cmbObjectLaser.ListIndex

End Sub

Private Sub cmbObjectMissile_Click()

    'Update missile system
    gudtObject(scrObjects.Value).udtSystems.bytMissile = cmbObjectMissile.ListIndex

End Sub

Private Sub cmbObjectPlanet_Click()

    'The object's planet value has changed
    gudtObject(scrObjects.Value).udtInfo.blnPlanet = GetBoolCombo(cmbObjectPlanet)

End Sub

Private Sub cmbObjectRace_Click()

    'The object's race has changed
    gudtObject(scrObjects.Value).udtInfo.bytRace = cmbObjectRace.ListIndex

End Sub

Private Sub cmbObjectScanner_Click()

    'Update scanner system
    gudtObject(scrObjects.Value).udtSystems.bytScanner = cmbObjectScanner.ListIndex

End Sub

Private Sub cmbObjectShield_Click()

    'Update shield system
    gudtObject(scrObjects.Value).udtSystems.bytShield = cmbObjectShield.ListIndex

End Sub

Private Sub cmbObjectStar_Click()

    'The object's star value has changed
    gudtObject(scrObjects.Value).udtInfo.blnStar = GetBoolCombo(cmbObjectStar)

End Sub

Private Sub cmbObjectStarbase_Click()

    'The object's starbase value has changed
    gudtObject(scrObjects.Value).udtInfo.blnStarBase = GetBoolCombo(cmbObjectStarbase)

End Sub

Private Sub cmdObjectCopy_Click()

    'Make a copy of the current object
    ReDim Preserve gudtObject(UBound(gudtObject) + 1)
    CopyObject scrObjects.Value, UBound(gudtObject)
    UpdateDisplay

End Sub

Private Sub cmdObjectDelete_Click()

Dim i As Long

    'Copy all of the objects above
    For i = scrObjects.Value To UBound(gudtObject) - 1
        CopyObject i + 1, i
    Next i
    
    'Remove the last
    ReDim Preserve gudtObject(UBound(gudtObject) - 1)
    
    'Refresh
    UpdateDisplay

End Sub

Private Sub CopyObject(lngObject As Long, lngTargetObject As Long)

Dim i As Long

    'Copy the first object over top of the second one
    With gudtObject(lngTargetObject)
        .blnExists = gudtObject(lngObject).blnExists
        .udtAI.bytAction = gudtObject(lngObject).udtAI.bytAction
        .udtAI.lngLengthTargetLock = gudtObject(lngObject).udtAI.lngLengthTargetLock
        .udtAI.sngAimTolerance = gudtObject(lngObject).udtAI.sngAimTolerance
        .udtAI.sngCannonDist = gudtObject(lngObject).udtAI.sngCannonDist
        .udtAI.sngMinDist = gudtObject(lngObject).udtAI.sngMinDist
        .udtAI.sngSeekDist = gudtObject(lngObject).udtAI.sngSeekDist
        .udtAI.sngTargetBias = gudtObject(lngObject).udtAI.sngTargetBias
        .udtAI.sngTolerance = gudtObject(lngObject).udtAI.sngTolerance
        .udtCarrier.intFighters = gudtObject(lngObject).udtCarrier.intFighters
        .udtCarrier.intMaxFighters = gudtObject(lngObject).udtCarrier.intMaxFighters
        .udtFighter.lngFighterOwner = gudtObject(lngObject).udtFighter.lngFighterOwner
        .udtInfo.blnCanMove = gudtObject(lngObject).udtInfo.blnCanMove
        .udtInfo.blnCarrier = gudtObject(lngObject).udtInfo.blnCarrier
        .udtInfo.blnFighter = gudtObject(lngObject).udtInfo.blnFighter
        .udtInfo.blnPlanet = gudtObject(lngObject).udtInfo.blnPlanet
        .udtInfo.blnStar = gudtObject(lngObject).udtInfo.blnStar
        .udtInfo.blnStarBase = gudtObject(lngObject).udtInfo.blnStarBase
        .udtInfo.bytRace = gudtObject(lngObject).udtInfo.bytRace
        .udtInfo.strName = gudtObject(lngObject).udtInfo.strName
        .udtPhysics.dblX = gudtObject(lngObject).udtPhysics.dblX
        .udtPhysics.dblY = gudtObject(lngObject).udtPhysics.dblY
        .udtPhysics.lngMass = gudtObject(lngObject).udtPhysics.lngMass
        .udtPhysics.sngFacing = gudtObject(lngObject).udtPhysics.sngFacing
        .udtPhysics.sngHeading = gudtObject(lngObject).udtPhysics.sngHeading
        .udtPhysics.sngSpeed = gudtObject(lngObject).udtPhysics.sngSpeed
        .udtSprite.bytAnimAmt = gudtObject(lngObject).udtSprite.bytAnimAmt
        .udtSprite.bytAnimNum = gudtObject(lngObject).udtSprite.bytAnimNum
        .udtSprite.bytFrameAmt = gudtObject(lngObject).udtSprite.bytFrameAmt
        .udtSprite.bytFrameNum = gudtObject(lngObject).udtSprite.bytFrameNum
        .udtSprite.intHeight = gudtObject(lngObject).udtSprite.intHeight
        .udtSprite.intWidth = gudtObject(lngObject).udtSprite.intWidth
        .udtSprite.lngAnimRate = gudtObject(lngObject).udtSprite.lngAnimRate
        .udtSprite.strResName = gudtObject(lngObject).udtSprite.strResName
        .udtStarbase.lngARCDPrice = gudtObject(lngObject).udtStarbase.lngARCDPrice
        .udtStarbase.lngCommJammerPrice = gudtObject(lngObject).udtStarbase.lngCommJammerPrice
        .udtStarbase.lngMinesPrice = gudtObject(lngObject).udtStarbase.lngMinesPrice
        .udtStarbase.lngFTLDPrice = gudtObject(lngObject).udtStarbase.lngFTLDPrice
        For i = 0 To COMMODITY_NUM
            .udtStarbase.lngCommodityPrice(i) = .udtStarbase.lngARCDPrice = gudtObject(lngObject).udtStarbase.lngCommodityPrice(i)
        Next i
        For i = 0 To LASER_NUM
            .udtStarbase.lngLaserPrice(i) = .udtStarbase.lngARCDPrice = gudtObject(lngObject).udtStarbase.lngLaserPrice(i)
        Next i
        For i = 0 To CANNON_NUM
            .udtStarbase.lngCannonPrice(i) = .udtStarbase.lngARCDPrice = gudtObject(lngObject).udtStarbase.lngCannonPrice(i)
        Next i
        For i = 0 To MISSILE_NUM
            .udtStarbase.lngMissilePrice(i) = .udtStarbase.lngARCDPrice = gudtObject(lngObject).udtStarbase.lngMissilePrice(i)
        Next i
        For i = 0 To HULL_NUM
            .udtStarbase.lngHullPrice(i) = .udtStarbase.lngARCDPrice = gudtObject(lngObject).udtStarbase.lngHullPrice(i)
        Next i
        For i = 0 To SHIELD_NUM
            .udtStarbase.lngShieldPrice(i) = .udtStarbase.lngARCDPrice = gudtObject(lngObject).udtStarbase.lngShieldPrice(i)
        Next i
        For i = 0 To GENERATOR_NUM
            .udtStarbase.lngGeneratorPrice(i) = .udtStarbase.lngARCDPrice = gudtObject(lngObject).udtStarbase.lngGeneratorPrice(i)
        Next i
        For i = 0 To ENGINE_NUM
            .udtStarbase.lngEnginePrice(i) = .udtStarbase.lngARCDPrice = gudtObject(lngObject).udtStarbase.lngEnginePrice(i)
        Next i
        For i = 0 To ARMOUR_NUM
            .udtStarbase.lngArmourPrice(i) = .udtStarbase.lngARCDPrice = gudtObject(lngObject).udtStarbase.lngArmourPrice(i)
        Next i
        .udtSystems.blnARCD = gudtObject(lngObject).udtSystems.blnARCD
        .udtSystems.blnFTLD = gudtObject(lngObject).udtSystems.blnFTLD
        .udtSystems.blnJammer = gudtObject(lngObject).udtSystems.blnJammer
        .udtSystems.bytArmour = gudtObject(lngObject).udtSystems.bytArmour
        .udtSystems.bytCannon = gudtObject(lngObject).udtSystems.bytCannon
        .udtSystems.bytEngine = gudtObject(lngObject).udtSystems.bytEngine
        .udtSystems.bytGenerator = gudtObject(lngObject).udtSystems.bytGenerator
        .udtSystems.bytHull = gudtObject(lngObject).udtSystems.bytHull
        .udtSystems.bytLaser = gudtObject(lngObject).udtSystems.bytLaser
        .udtSystems.bytMissile = gudtObject(lngObject).udtSystems.bytMissile
        .udtSystems.bytScanner = gudtObject(lngObject).udtSystems.bytScanner
        .udtSystems.bytShield = gudtObject(lngObject).udtSystems.bytShield
        .udtSystems.intMineNum = gudtObject(lngObject).udtSystems.intMineNum
        .udtSystems.intMissileNum = gudtObject(lngObject).udtSystems.intMissileNum
        .udtSystems.lngArmour = gudtObject(lngObject).udtSystems.lngArmour
        .udtSystems.lngCrew = gudtObject(lngObject).udtSystems.lngCrew
        .udtSystems.sngEnergy = gudtObject(lngObject).udtSystems.sngEnergy
        .udtSystems.sngEngineEnergy = gudtObject(lngObject).udtSystems.sngEngineEnergy
        .udtSystems.sngFuel = gudtObject(lngObject).udtSystems.sngFuel
        .udtSystems.sngGeneratorEnergy = gudtObject(lngObject).udtSystems.sngGeneratorEnergy
        .udtSystems.sngRotationRate = gudtObject(lngObject).udtSystems.sngRotationRate
        .udtSystems.sngShieldEnergy = gudtObject(lngObject).udtSystems.sngShieldEnergy
        .udtSystems.sngWeaponEnergy = gudtObject(lngObject).udtSystems.sngWeaponEnergy
    End With

End Sub

Private Sub mnuFileOpen_Click()

    'Open a universe file
    cdlUniverse.InitDir = App.Path
    cdlUniverse.FileName = ""
    cdlUniverse.flags = HideReadOnly Or FileMustExist
    cdlUniverse.Filter = "Universe Files (*.uni)|*.uni"
    On Error Resume Next
    cdlUniverse.ShowOpen
    'If the user canceled out of the dialog, exit sub
    If Err.Number = cdlCancel Or cdlUniverse.FileName = "" Then
        Exit Sub
    End If
    On Error GoTo 0
    
    'Initialize the display
    InitializeDisplay
    
    'Extract the data
    gstrUniverse = cdlUniverse.FileName
    LoadUniverse False
    
    'Update the display
    UpdateDisplay

End Sub

Private Sub InitializeDisplay()

    'Initialize the various displays
    InitializeObjectsDisplay
    InitializeHullsDisplay
    
End Sub

Private Sub InitializeHullsDisplay()

    'Populate all of the known-quantity fields
    InitBoolCombo cmbHullJammer
    InitBoolCombo cmbHullMines
    InitBoolCombo cmbHullFTLD
    InitBoolCombo cmbHullARCD
    
End Sub

Private Sub InitializeObjectsDisplay()

    'Populate all of the known-quantity fields
    InitBoolCombo cmbObjectExists
    InitBoolCombo cmbObjectCanMove
    InitBoolCombo cmbObjectCarrier
    InitBoolCombo cmbObjectFighter
    InitBoolCombo cmbObjectPlanet
    InitBoolCombo cmbObjectStar
    InitBoolCombo cmbObjectStarbase
    InitializeObjectAICombo
    
    'Populate the system combos
    Dim i As Long
    'Armour
    cmbObjectArmour.Clear
    For i = 0 To ARMOUR_NUM
        cmbObjectArmour.AddItem gudtArmour(i).strName, i
    Next i
    'Cannon
    cmbObjectCannon.Clear
    For i = 0 To CANNON_NUM
        cmbObjectCannon.AddItem gudtCannon(i).strName, i
    Next i
    'Engine
    cmbObjectEngine.Clear
    For i = 0 To ENGINE_NUM
        cmbObjectEngine.AddItem gudtEngine(i).strName, i
    Next i
    'Generator
    cmbObjectGenerator.Clear
    For i = 0 To GENERATOR_NUM
        cmbObjectGenerator.AddItem gudtGenerator(i).strName, i
    Next i
    'Hull
    cmbObjectHull.Clear
    For i = 0 To HULL_NUM
        cmbObjectHull.AddItem gudtHull(i).strName, i
    Next i
    'Laser
    cmbObjectLaser.Clear
    For i = 0 To LASER_NUM
        cmbObjectLaser.AddItem gudtLaser(i).strName, i
    Next i
    'Missile
    cmbObjectMissile.Clear
    For i = 0 To MISSILE_NUM
        cmbObjectMissile.AddItem gudtMissile(i).strName, i
    Next i
    'Scanner
    cmbObjectScanner.Clear
    For i = 0 To SCANNER_NUM
        cmbObjectScanner.AddItem gudtScanner(i).strName, i
    Next i
    'Shield
    cmbObjectShield.Clear
    For i = 0 To SHIELD_NUM
        cmbObjectShield.AddItem gudtShield(i).strName, i
    Next i

End Sub

Private Sub InitializeObjectAICombo()

    'Add all of the known AI states to the list
    cmbObjectAction.Clear
    cmbObjectAction.AddItem "0 - None", 0
    cmbObjectAction.AddItem "1 - Flee", 1
    cmbObjectAction.AddItem "2 - Attack", 2
    cmbObjectAction.AddItem "3 - Patrol", 3
    cmbObjectAction.AddItem "4 - Trade", 4
    cmbObjectAction.AddItem "5 - Seek", 5
    cmbObjectAction.AddItem "6 - All Stop", 6
    cmbObjectAction.AddItem "7 - Autopilot", 7

End Sub

Private Sub UpdateDisplay()

    'Update the scroll bars
    scrObjects.Max = UBound(gudtObject)
    scrObjects.Value = 0
    scrCannons.Max = CANNON_NUM
    scrCannons.Value = 0
    scrEngines.Max = ENGINE_NUM
    scrEngines.Value = 0
    scrLasers.Max = LASER_NUM
    scrLasers.Value = 0
    scrMissiles.Max = MISSILE_NUM
    scrMissiles.Value = 0
    scrHulls.Value = 0
    scrHulls.Max = HULL_NUM

    'Update the various displays
    UpdateObjectsDisplay
    UpdateCannonsDisplay
    UpdateEnginesDisplay
    UpdateLasersDisplay
    UpdateMissilesDisplay
    UpdateHullsDisplay

End Sub

Private Sub mnuFileSave_Click()

    'Save the data
    SaveUniverse

End Sub

Private Sub mnuFileSaveAs_Click()

    'Save a universe file
    cdlUniverse.InitDir = App.Path
    cdlUniverse.FileName = ""
    cdlUniverse.flags = HideReadOnly Or FileMustExist
    cdlUniverse.Filter = "Universe Files (*.uni)|*.uni"
    On Error Resume Next
    cdlUniverse.ShowSave
    'If the user canceled out of the dialog, exit sub
    If Err.Number = cdlCancel Or cdlUniverse.FileName = "" Then
        Exit Sub
    End If
    On Error GoTo 0
    
    'Save the data
    gstrUniverse = cdlUniverse.FileName
    SaveUniverse

End Sub

Private Sub scrCannons_Change()

    'User has scrolled to another cannon, update the cannons display
    UpdateCannonsDisplay

End Sub

Private Sub scrEngines_Change()

    'User has scrolled to another engine, update the engines display
    UpdateEnginesDisplay

End Sub

Private Sub scrHulls_Change()

    'User has scrolled to another hull
    UpdateHullsDisplay

End Sub

Private Sub scrLasers_Change()

    'User has scrolled to another laser, update the lasers display
    UpdateLasersDisplay

End Sub

Private Sub scrMissiles_Change()

    'User has scrolled to another missile
    UpdateMissilesDisplay

End Sub

Private Sub scrObjects_Change()

    'User has scrolled to another object, update the objects display
    UpdateObjectsDisplay

End Sub

Private Sub txtCannonConcussive_Change()

    'The concussive damage value has changed
    On Local Error Resume Next
    gudtCannon(scrCannons.Value).sngConcussiveDamage = txtCannonConcussive.Text

End Sub

Private Sub txtCannonConsumption_Change()

    'The consumption value has changed
    On Local Error Resume Next
    gudtCannon(scrCannons.Value).sngConsumption = txtCannonConsumption.Text

End Sub

Private Sub txtCannonDuration_Change()

    'The bullet duration value has changed
    On Local Error Resume Next
    gudtCannon(scrCannons.Value).lngDuration = txtCannonDuration.Text

End Sub

Private Sub txtCannonFireRate_Change()

    'The cannon fire rate value has changed
    On Local Error Resume Next
    gudtCannon(scrCannons.Value).lngFireRate = txtCannonFireRate.Text

End Sub

Private Sub txtCannonInstantaneous_Change()

    'The instantaneous consumption value has changed
    On Local Error Resume Next
    gudtCannon(scrCannons.Value).sngInstantaneousConsumption = txtCannonInstantaneous.Text

End Sub

Private Sub txtCannonMaxEnergy_Change()

    'The max energy value has changed
    On Local Error Resume Next
    gudtCannon(scrCannons.Value).lngMaxEnergy = txtCannonMaxEnergy.Text

End Sub

Private Sub txtCannonName_Change()

    'The object name has changed
    On Local Error Resume Next
    gudtCannon(scrCannons.Value).strName = txtCannonName.Text

End Sub

Private Sub txtCannonNum_Change()

    'User has changed the cannon number, update scroll bar
    ItemNumChange txtCannonNum, scrCannons

End Sub

Private Sub txtCannonRadiation_Change()

    'The radiation damage value has changed
    On Local Error Resume Next
    gudtCannon(scrCannons.Value).sngRadiationDamage = txtCannonRadiation.Text

End Sub

Private Sub txtCannonSound_Change()

    'The cannon sound file value has changed
    On Local Error Resume Next
    gstrCannonSound(scrCannons.Value) = txtCannonSound.Text

End Sub

Private Sub txtCannonSpeed_Change()

    'The bullet speed value has changed
    On Local Error Resume Next
    gudtCannon(scrCannons.Value).sngSpeed = txtCannonSpeed.Text

End Sub

Private Sub txtCannonSprite_Change()

    'The cannon sprite value has changed
    On Local Error Resume Next
    gstrCannonSprite(scrCannons.Value) = txtCannonSprite.Text

End Sub

Private Sub txtCannonType_Change()

    'The cannon type value has changed
    On Local Error Resume Next
    gudtCannon(scrCannons.Value).lngCannonType = txtCannonType.Text

End Sub

Private Sub txtEngineConsumption_Change()

    'The engine consumption value has changed
    On Local Error Resume Next
    gudtEngine(scrEngines.Value).sngConsumption = txtEngineConsumption.Text

End Sub

Private Sub txtEngineMaxEnergy_Change()

    'The engine max energy value has changed
    On Local Error Resume Next
    gudtEngine(scrEngines.Value).lngMaxEnergy = txtEngineMaxEnergy.Text

End Sub

Private Sub txtEngineName_Change()

    'The engine name value has changed
    On Local Error Resume Next
    gudtEngine(scrEngines.Value).strName = txtEngineName.Text

End Sub

Private Sub txtEngineNum_Change()

    'User has changed the engine number, update scroll bar
    ItemNumChange txtEngineNum, scrEngines

End Sub

Private Sub ItemNumChange(txtNum As TextBox, scrNum As HScrollBar)

    'User has changed the item number, update scroll bar
    If IsNumeric(txtNum.Text) Then
        Dim lngTemp As Long
        lngTemp = CLng(txtNum.Text)
        If (lngTemp >= 0) And (lngTemp <= scrNum.Max) Then scrNum.Value = lngTemp
    End If

End Sub

Private Sub txtEngineSound_Change()

    'The engine sound file value has changed
    On Local Error Resume Next
    gstrEngineSound(scrEngines.Value) = txtEngineSound.Text

End Sub

Private Sub txtEngineThrust_Change()

    'The engine thrust value has changed
    On Local Error Resume Next
    gudtEngine(scrEngines.Value).sngThrust = txtEngineThrust.Text

End Sub

Private Sub txtHullArmour_Change()

    'Update hull armour
    On Local Error Resume Next
    gudtHull(scrHulls.Value).lngArmour = txtHullArmour.Text

End Sub

Private Sub txtHullCannon_Change()

    'Update hull cannons
    On Local Error Resume Next
    gudtHull(scrHulls.Value).lngCannon = txtHullCannon.Text

End Sub

Private Sub txtHullEngine_Change()

    'Update hull engines
    On Local Error Resume Next
    gudtHull(scrHulls.Value).lngEngine = txtHullEngine.Text

End Sub

Private Sub txtHullGenerator_Change()

    'Update hull generators
    On Local Error Resume Next
    gudtHull(scrHulls.Value).lngGenerator = txtHullGenerator.Text

End Sub

Private Sub txtHullLaser_Change()

    'Update hull lasers
    On Local Error Resume Next
    gudtHull(scrHulls.Value).lngLaser = txtHullLaser.Text

End Sub

Private Sub txtHullMass_Change()

    'Update hull mass
    On Local Error Resume Next
    gudtHull(scrHulls.Value).lngMass = txtHullMass.Text

End Sub

Private Sub txtHullMaxCargo_Change()

    'Update hull max cargo
    On Local Error Resume Next
    gudtHull(scrHulls.Value).lngMaxCargo = txtHullMaxCargo.Text

End Sub

Private Sub txtHullMaxCrew_Change()

    'Update hull max crew
    On Local Error Resume Next
    gudtHull(scrHulls.Value).lngMaxCrew = txtHullMaxCrew.Text

End Sub

Private Sub txtHullMaxFuel_Change()

    'Update hull max fuel
    On Local Error Resume Next
    gudtHull(scrHulls.Value).lngMaxFuel = txtHullMaxFuel.Text

End Sub

Private Sub txtHullMaxMines_Change()

    'Update hull max mines
    On Local Error Resume Next
    gudtHull(scrHulls.Value).lngMaxMines = txtHullMaxMines.Text

End Sub

Private Sub txtHullMaxMissile_Change()

    'Update hull max missile
    On Local Error Resume Next
    gudtHull(scrHulls.Value).lngMaxMissile = txtHullMaxMissile.Text

End Sub

Private Sub txtHullMaxSpeed_Change()

    'Update hull max speed
    On Local Error Resume Next
    gudtHull(scrHulls.Value).sngMaxSpeed = txtHullMaxSpeed.Text

End Sub

Private Sub txtHullMissile_Change()

    'Update hull missiles
    On Local Error Resume Next
    gudtHull(scrHulls.Value).lngMissile = txtHullMissile.Text

End Sub

Private Sub txtHullName_Change()

    'Update hull name
    On Local Error Resume Next
    gudtHull(scrHulls.Value).strName = txtHullName.Text

End Sub

Private Sub txtHullNum_Change()

    'User has changed the hull number, update scroll bar
    ItemNumChange txtHullNum, scrHulls

End Sub

Private Sub txtHullRotationRate_Change()

    'Update hull rotation rate
    On Local Error Resume Next
    gudtHull(scrHulls.Value).sngRotationRate = txtHullRotationRate.Text

End Sub

Private Sub txtHullShield_Change()

    'Update hull shields
    On Local Error Resume Next
    gudtHull(scrHulls.Value).lngShield = txtHullShield.Text

End Sub

Private Sub txtHullSprite_Change()

    'Update hull sprite
    On Local Error Resume Next
    gudtHull(scrHulls.Value).udtSprite.strResName = txtHullSprite.Text

End Sub

Private Sub txtHullSpriteAnimAmt_Change()

    'Update hull sprite anim amount
    On Local Error Resume Next
    gudtHull(scrHulls.Value).udtSprite.bytAnimAmt = txtHullSpriteAnimAmt.Text

End Sub

Private Sub txtHullSpriteAnimRate_Change()

    'Update hull sprite anim rate
    On Local Error Resume Next
    gudtHull(scrHulls.Value).udtSprite.lngAnimRate = txtHullSpriteAnimRate.Text

End Sub

Private Sub txtHullSpriteFrameAmt_Change()

    'Update hull sprite frame amount
    On Local Error Resume Next
    gudtHull(scrHulls.Value).udtSprite.bytFrameAmt = txtHullSpriteFrameAmt.Text

End Sub

Private Sub txtHullSpriteHeight_Change()

    'Update hull sprite height
    On Local Error Resume Next
    gudtHull(scrHulls.Value).udtSprite.intHeight = txtHullSpriteHeight.Text

End Sub

Private Sub txtHullSpriteWidth_Change()

    'Update hull sprite width
    On Local Error Resume Next
    gudtHull(scrHulls.Value).udtSprite.intWidth = txtHullSpriteWidth.Text

End Sub

Private Sub txtLaserColour_Change()

    'Set colour
    On Local Error Resume Next
    gudtLaser(scrLasers.Value).lngColour = txtLaserColour.Text

End Sub

Private Sub txtLaserConcussive_Change()

    'Set concussive damage
    On Local Error Resume Next
    gudtLaser(scrLasers.Value).sngConcussiveDamage = txtLaserConcussive.Text

End Sub

Private Sub txtLaserConsumption_Change()

    'Set consumption
    On Local Error Resume Next
    gudtLaser(scrLasers.Value).sngConsumption = txtLaserConsumption.Text

End Sub

Private Sub txtLaserFireConsumption_Change()

    'Set firing consumption
    On Local Error Resume Next
    gudtLaser(scrLasers.Value).sngFireConsumption = txtLaserFireConsumption.Text

End Sub

Private Sub txtLaserMaxEnergy_Change()

    'Set max energy
    On Local Error Resume Next
    gudtLaser(scrLasers.Value).lngMaxEnergy = txtLaserMaxEnergy.Text

End Sub

Private Sub txtLaserName_Change()

    'Set name
    On Local Error Resume Next
    gudtLaser(scrLasers.Value).strName = txtLaserName.Text

End Sub

Private Sub txtLaserNum_Change()

    'User has changed the laser number, update scroll bar
    ItemNumChange txtLaserNum, scrLasers

End Sub

Private Sub txtLaserRadiation_Change()

    'Set radiation damage
    On Local Error Resume Next
    gudtLaser(scrLasers.Value).sngRadiationDamage = txtLaserRadiation.Text

End Sub

Private Sub txtLaserRange_Change()

    'Set range
    On Local Error Resume Next
    gudtLaser(scrLasers.Value).lngRange = txtLaserRange.Text

End Sub

Private Sub txtLaserSound_Change()

    'Set sound file
    On Local Error Resume Next
    gstrLaserSound(scrLasers.Value) = txtLaserSound.Text

End Sub

Private Sub txtMissileConcussive_Change()

    'Update missile concussive damage
    On Local Error Resume Next
    gudtMissile(scrMissiles.Value).sngConcussiveDamage = txtMissileConcussive.Text

End Sub

Private Sub txtMissileDuration_Change()

    'Update missile duration
    On Local Error Resume Next
    gudtMissile(scrMissiles.Value).lngDuration = txtMissileDuration.Text

End Sub

Private Sub txtMissileFireRate_Change()

    'Update missile fire rate
    On Local Error Resume Next
    gudtMissile(scrMissiles.Value).lngFireRate = txtMissileFireRate.Text

End Sub

Private Sub txtMissileMaxSpeed_Change()

    'Update missile max speed
    On Local Error Resume Next
    gudtMissile(scrMissiles.Value).sngMaxSpeed = txtMissileMaxSpeed.Text

End Sub

Private Sub txtMissileName_Change()

    'Update missile name
    On Local Error Resume Next
    gudtMissile(scrMissiles.Value).strName = txtMissileName.Text

End Sub

Private Sub txtMissileNum_Change()

    'User has changed the missile number, update scroll bar
    ItemNumChange txtMissileNum, scrMissiles

End Sub

Private Sub txtMissileRadiation_Change()

    'Update missile radiation damage
    On Local Error Resume Next
    gudtMissile(scrMissiles.Value).sngRadiationDamage = txtMissileRadiation.Text

End Sub

Private Sub txtMissileRotationRate_Change()

    'Update missile rotation rate
    On Local Error Resume Next
    gudtMissile(scrMissiles.Value).sngRotationRate = txtMissileRotationRate.Text

End Sub

Private Sub txtMissileSeekDist_Change()

    'Update missile seek dist
    On Local Error Resume Next
    gudtMissile(scrMissiles.Value).sngSeekDist = txtMissileSeekDist.Text

End Sub

Private Sub txtMissileSound_Change()

    'Update missile sound file
    On Local Error Resume Next
    gstrMissileSound(scrMissiles.Value) = txtMissileSound.Text

End Sub

Private Sub txtMissileSprite_Change()

    'Update missile sprite name
    On Local Error Resume Next
    gudtMissile(scrMissiles.Value).udtSprite.strResName = txtMissileSprite.Text

End Sub

Private Sub txtMissileSpriteAnimAmt_Change()

    'Update missile animation amount
    On Local Error Resume Next
    gudtMissile(scrMissiles.Value).udtSprite.bytAnimAmt = txtMissileSpriteAnimAmt.Text

End Sub

Private Sub txtMissileSpriteAnimRate_Change()

    'Update missile anim rate
    On Local Error Resume Next
    gudtMissile(scrMissiles.Value).udtSprite.lngAnimRate = txtMissileSpriteAnimRate.Text

End Sub

Private Sub txtMissileSpriteFrameAmt_Change()

    'Update missile frame amount
    On Local Error Resume Next
    gudtMissile(scrMissiles.Value).udtSprite.bytFrameAmt = txtMissileSpriteFrameAmt.Text

End Sub

Private Sub txtMissileSpriteHeight_Change()

    'Update missile sprite height
    On Local Error Resume Next
    gudtMissile(scrMissiles.Value).udtSprite.intHeight = txtMissileSpriteHeight.Text

End Sub

Private Sub txtMissileSpriteWidth_Change()

    'Update missile sprite width
    On Local Error Resume Next
    gudtMissile(scrMissiles.Value).udtSprite.intWidth = txtMissileSpriteWidth.Text

End Sub

Private Sub txtMissileTargetBias_Change()

    'Update missile target bias
    On Local Error Resume Next
    gudtMissile(scrMissiles.Value).sngTargetBias = txtMissileTargetBias.Text

End Sub

Private Sub txtMissileThrust_Change()

    'Update missile thrust
    On Local Error Resume Next
    gudtMissile(scrMissiles.Value).sngThrust = txtMissileThrust.Text

End Sub

Private Sub txtObjectAimTolerance_Change()

    'The object's aim tolerance has changed
    On Local Error Resume Next
    gudtObject(scrObjects.Value).udtAI.sngAimTolerance = txtObjectAimTolerance.Text

End Sub

Private Sub txtObjectCannonDist_Change()

    'The object's cannon distance has changed
    On Local Error Resume Next
    gudtObject(scrObjects.Value).udtAI.sngCannonDist = txtObjectCannonDist.Text

End Sub

Private Sub txtObjectFacing_Change()

    'The object facing has changed
    On Local Error Resume Next
    gudtObject(scrObjects.Value).udtPhysics.sngFacing = txtObjectFacing.Text

End Sub

Private Sub txtObjectHeading_Change()

    'The object heading has changed
    On Local Error Resume Next
    gudtObject(scrObjects.Value).udtPhysics.sngHeading = txtObjectHeading.Text
    
End Sub

Private Sub txtObjectLengthTargetLock_Change()

    'The object's lock duration has changed
    On Local Error Resume Next
    gudtObject(scrObjects.Value).udtAI.lngLengthTargetLock = txtObjectLengthTargetLock.Text

End Sub

Private Sub txtObjectMass_Change()

    'The object mass has changed
    On Local Error Resume Next
    gudtObject(scrObjects.Value).udtPhysics.lngMass = txtObjectMass.Text
    
End Sub

Private Sub txtObjectMinDist_Change()

    'The object's min distance has changed
    On Local Error Resume Next
    gudtObject(scrObjects.Value).udtAI.sngMinDist = txtObjectMinDist.Text
    
End Sub

Private Sub txtObjectName_Change()

    'The object name has changed
    On Local Error Resume Next
    gudtObject(scrObjects.Value).udtInfo.strName = txtObjectName.Text

End Sub

Private Sub txtObjectNum_Change()

    'User has changed the object number, update scroll bar
    ItemNumChange txtObjectNum, scrObjects

End Sub

Private Sub UpdateHullsDisplay()

    'Set the hull num
    txtHullNum.Text = scrHulls.Value
    
    'Set the values
    txtHullName.Text = gudtHull(scrHulls.Value).strName
    txtHullMass.Text = gudtHull(scrHulls.Value).lngMass
    txtHullMaxCargo.Text = gudtHull(scrHulls.Value).lngMaxCargo
    txtHullMaxCrew.Text = gudtHull(scrHulls.Value).lngMaxCrew
    txtHullMaxMissile.Text = gudtHull(scrHulls.Value).lngMaxMissile
    txtHullMaxMines.Text = gudtHull(scrHulls.Value).lngMaxMines
    txtHullMaxFuel.Text = gudtHull(scrHulls.Value).lngMaxFuel
    txtHullMaxSpeed.Text = gudtHull(scrHulls.Value).sngMaxSpeed
    txtHullRotationRate.Text = gudtHull(scrHulls.Value).sngRotationRate
    txtHullArmour.Text = gudtHull(scrHulls.Value).lngArmour
    txtHullCannon.Text = gudtHull(scrHulls.Value).lngCannon
    txtHullMissile.Text = gudtHull(scrHulls.Value).lngMissile
    txtHullShield.Text = gudtHull(scrHulls.Value).lngShield
    txtHullGenerator.Text = gudtHull(scrHulls.Value).lngGenerator
    txtHullEngine.Text = gudtHull(scrHulls.Value).lngEngine
    txtHullLaser.Text = gudtHull(scrHulls.Value).lngLaser
    cmbHullJammer.ListIndex = SetBoolCombo(gudtHull(scrHulls.Value).blnCommJammer)
    cmbHullMines.ListIndex = SetBoolCombo(gudtHull(scrHulls.Value).blnMines)
    cmbHullFTLD.ListIndex = SetBoolCombo(gudtHull(scrHulls.Value).blnFTLD)
    cmbHullARCD.ListIndex = SetBoolCombo(gudtHull(scrHulls.Value).blnARCD)
    txtHullSprite.Text = gudtHull(scrHulls.Value).udtSprite.strResName
    txtHullSpriteWidth.Text = gudtHull(scrHulls.Value).udtSprite.intWidth
    txtHullSpriteHeight.Text = gudtHull(scrHulls.Value).udtSprite.intHeight
    txtHullSpriteFrameAmt.Text = gudtHull(scrHulls.Value).udtSprite.bytFrameAmt
    txtHullSpriteAnimAmt.Text = gudtHull(scrHulls.Value).udtSprite.bytAnimAmt
    txtHullSpriteAnimRate.Text = gudtHull(scrHulls.Value).udtSprite.lngAnimRate

End Sub

Private Sub UpdateMissilesDisplay()

    'Set the missile num
    txtMissileNum.Text = scrMissiles.Value
    
    'Set the values
    txtMissileName.Text = gudtMissile(scrMissiles.Value).strName
    txtMissileConcussive.Text = gudtMissile(scrMissiles.Value).sngConcussiveDamage
    txtMissileRadiation.Text = gudtMissile(scrMissiles.Value).sngRadiationDamage
    txtMissileRotationRate.Text = gudtMissile(scrMissiles.Value).sngRotationRate
    txtMissileFireRate.Text = gudtMissile(scrMissiles.Value).lngFireRate
    txtMissileMaxSpeed.Text = gudtMissile(scrMissiles.Value).sngMaxSpeed
    txtMissileThrust.Text = gudtMissile(scrMissiles.Value).sngThrust
    txtMissileDuration.Text = gudtMissile(scrMissiles.Value).lngDuration
    txtMissileTargetBias.Text = gudtMissile(scrMissiles.Value).sngTargetBias
    txtMissileSeekDist.Text = gudtMissile(scrMissiles.Value).sngSeekDist
    txtMissileSound.Text = gstrMissileSound(scrMissiles.Value)
    txtMissileSprite.Text = gudtMissile(scrMissiles.Value).udtSprite.strResName
    txtMissileSpriteWidth.Text = gudtMissile(scrMissiles.Value).udtSprite.intWidth
    txtMissileSpriteHeight.Text = gudtMissile(scrMissiles.Value).udtSprite.intHeight
    txtMissileSpriteFrameAmt.Text = gudtMissile(scrMissiles.Value).udtSprite.bytFrameAmt
    txtMissileSpriteAnimAmt.Text = gudtMissile(scrMissiles.Value).udtSprite.bytAnimAmt
    txtMissileSpriteAnimRate.Text = gudtMissile(scrMissiles.Value).udtSprite.lngAnimRate

End Sub

Private Sub UpdateLasersDisplay()

    'Set the laser num
    txtLaserNum.Text = scrLasers.Value
    
    'Set the name
    txtLaserName.Text = gudtLaser(scrLasers.Value).strName
    
    'Set the consumption
    txtLaserConsumption.Text = gudtLaser(scrLasers.Value).sngConsumption
    
    'Set the firing consumption
    txtLaserFireConsumption.Text = gudtLaser(scrLasers.Value).sngFireConsumption
    
    'Set the concussive damage
    txtLaserConcussive.Text = gudtLaser(scrLasers.Value).sngConcussiveDamage
    
    'Set the radiation damage
    txtLaserRadiation.Text = gudtLaser(scrLasers.Value).sngRadiationDamage
    
    'Set the range
    txtLaserRange.Text = gudtLaser(scrLasers.Value).lngRange
    
    'Set the max energy
    txtLaserMaxEnergy.Text = gudtLaser(scrLasers.Value).lngMaxEnergy
    
    'Set the colour
    txtLaserColour.Text = gudtLaser(scrLasers.Value).lngColour
    
    'Set the sound file
    txtLaserSound.Text = gstrLaserSound(scrLasers.Value)

End Sub

Private Sub UpdateEnginesDisplay()

    'Set the engine num
    txtEngineNum.Text = scrEngines.Value
    
    'Set the name
    txtEngineName.Text = gudtEngine(scrEngines.Value).strName
    
    'Set the consumption
    txtEngineConsumption.Text = gudtEngine(scrEngines.Value).sngConsumption
    
    'Set the thrust
    txtEngineThrust.Text = gudtEngine(scrEngines.Value).sngThrust
    
    'Set the max energy
    txtEngineMaxEnergy.Text = gudtEngine(scrEngines.Value).lngMaxEnergy
    
    'Set the sound file
    txtEngineSound.Text = gstrEngineSound(scrEngines.Value)

End Sub

Private Sub UpdateCannonsDisplay()

    'Set the cannon num
    txtCannonNum.Text = scrCannons.Value

    'Set the name
    txtCannonName.Text = gudtCannon(scrCannons.Value).strName
    
    'Set the consumption
    txtCannonConsumption.Text = gudtCannon(scrCannons.Value).sngConsumption
    
    'Set the instantaneous consumption
    txtCannonInstantaneous.Text = gudtCannon(scrCannons.Value).sngInstantaneousConsumption
    
    'Set the concussive damage
    txtCannonConcussive.Text = gudtCannon(scrCannons.Value).sngConcussiveDamage
    
    'Set the radiation damage
    txtCannonRadiation.Text = gudtCannon(scrCannons.Value).sngRadiationDamage
    
    'Set the bullet speed
    txtCannonSpeed.Text = gudtCannon(scrCannons.Value).sngSpeed
    
    'Set the bullet duration
    txtCannonDuration.Text = gudtCannon(scrCannons.Value).lngDuration
    
    'Set the max energy
    txtCannonMaxEnergy.Text = gudtCannon(scrCannons.Value).lngMaxEnergy
    
    'Set the fire rate
    txtCannonFireRate.Text = gudtCannon(scrCannons.Value).lngFireRate
    
    'Set the sound file
    txtCannonSound.Text = gstrCannonSound(scrCannons.Value)
    
    'Set the sprite file
    txtCannonSprite.Text = gstrCannonSprite(scrCannons.Value)
    
    'Set the cannon type
    txtCannonType.Text = gudtCannon(scrCannons.Value).lngCannonType

End Sub

Private Sub UpdateObjectsDisplay()

    'Set the object num
    txtObjectNum.Text = scrObjects.Value

    'Exists?
    cmbObjectExists.ListIndex = SetBoolCombo(gudtObject(scrObjects.Value).blnExists)
    
    'Update the sub displays
    UpdateObjectInfoDisplay
    UpdateObjectPhysicsDisplay
    UpdateObjectAIDisplay
    UpdateObjectSystemsDisplay

End Sub

Private Sub UpdateObjectSystemsDisplay()

    'Set the values
    cmbObjectArmour.ListIndex = gudtObject(scrObjects.Value).udtSystems.bytArmour
    cmbObjectCannon.ListIndex = gudtObject(scrObjects.Value).udtSystems.bytCannon
    cmbObjectEngine.ListIndex = gudtObject(scrObjects.Value).udtSystems.bytEngine
    cmbObjectGenerator.ListIndex = gudtObject(scrObjects.Value).udtSystems.bytGenerator
    cmbObjectHull.ListIndex = gudtObject(scrObjects.Value).udtSystems.bytHull
    cmbObjectLaser.ListIndex = gudtObject(scrObjects.Value).udtSystems.bytLaser
    cmbObjectMissile.ListIndex = gudtObject(scrObjects.Value).udtSystems.bytMissile
    cmbObjectScanner.ListIndex = gudtObject(scrObjects.Value).udtSystems.bytScanner
    cmbObjectShield.ListIndex = gudtObject(scrObjects.Value).udtSystems.bytShield

End Sub

Private Sub UpdateObjectAIDisplay()

    'Set the action
    cmbObjectAction.ListIndex = gudtObject(scrObjects.Value).udtAI.bytAction
    
    'Set the target bias
    txtObjectTargetBias.Text = gudtObject(scrObjects.Value).udtAI.sngTargetBias
    
    'Set the seek distance
    txtObjectSeekDist.Text = gudtObject(scrObjects.Value).udtAI.sngSeekDist
    
    'Set the min distnace
    txtObjectMinDist.Text = gudtObject(scrObjects.Value).udtAI.sngMinDist
    
    'Set the cannon distance
    txtObjectCannonDist.Text = gudtObject(scrObjects.Value).udtAI.sngCannonDist
    
    'Set the aim tolerance
    txtObjectAimTolerance.Text = gudtObject(scrObjects.Value).udtAI.sngAimTolerance
    
    'Set the lock duration
    txtObjectLengthTargetLock.Text = gudtObject(scrObjects.Value).udtAI.lngLengthTargetLock

End Sub

Private Sub UpdateObjectPhysicsDisplay()

    'Set the X coord
    txtObjectXCoord.Text = gudtObject(scrObjects.Value).udtPhysics.dblX
    
    'Set the Y coord
    txtObjectYCoord.Text = gudtObject(scrObjects.Value).udtPhysics.dblY
    
    'Set the facing
    txtObjectFacing.Text = gudtObject(scrObjects.Value).udtPhysics.sngFacing
    
    'Set the heading
    txtObjectHeading.Text = gudtObject(scrObjects.Value).udtPhysics.sngHeading
    
    'Set the speed
    txtObjectSpeed.Text = gudtObject(scrObjects.Value).udtPhysics.sngSpeed
    
    'Set the mass
    txtObjectMass.Text = gudtObject(scrObjects.Value).udtPhysics.lngMass

End Sub

Private Sub UpdateObjectInfoDisplay()

Dim i As Long

    'Set the name
    txtObjectName.Text = gudtObject(scrObjects.Value).udtInfo.strName

    'Set the race list
    cmbObjectRace.Clear
    For i = 0 To UBound(gudtRace)
        cmbObjectRace.AddItem gudtRace(i).strName
    Next i

    'Set the race
    cmbObjectRace.ListIndex = gudtObject(scrObjects.Value).udtInfo.bytRace
    
    'Set the Can Move value
    cmbObjectCanMove.ListIndex = SetBoolCombo(gudtObject(scrObjects.Value).udtInfo.blnCanMove)
    
    'Set the carrier value
    cmbObjectCarrier.ListIndex = SetBoolCombo(gudtObject(scrObjects.Value).udtInfo.blnCarrier)
    
    'Set the fighter value
    cmbObjectFighter.ListIndex = SetBoolCombo(gudtObject(scrObjects.Value).udtInfo.blnFighter)
    
    'Set the planet value
    cmbObjectPlanet.ListIndex = SetBoolCombo(gudtObject(scrObjects.Value).udtInfo.blnPlanet)
    
    'Set the star value
    cmbObjectStar.ListIndex = SetBoolCombo(gudtObject(scrObjects.Value).udtInfo.blnStar)
    
    'Set the starbase value
    cmbObjectStarbase.ListIndex = SetBoolCombo(gudtObject(scrObjects.Value).udtInfo.blnStarBase)

End Sub

Private Sub txtObjectSeekDist_Change()

    'The object's seek distance has changed
    On Local Error Resume Next
    gudtObject(scrObjects.Value).udtAI.sngSeekDist = txtObjectSeekDist.Text

End Sub

Private Sub txtObjectSpeed_Change()

    'The object speed has changed
    On Local Error Resume Next
    gudtObject(scrObjects.Value).udtPhysics.sngSpeed = txtObjectSpeed.Text
    
End Sub

Private Sub txtObjectTargetBias_Change()

    'The object's target bias has been changed
    On Local Error Resume Next
    gudtObject(scrObjects.Value).udtAI.sngTargetBias = txtObjectTargetBias.Text

End Sub

Private Sub txtObjectXCoord_Change()

    'The object X coord has changed
    On Local Error Resume Next
    gudtObject(scrObjects.Value).udtPhysics.dblX = txtObjectXCoord.Text

End Sub

Private Sub txtObjectYCoord_Change()

    'The object Y coord has changed
    On Local Error Resume Next
    gudtObject(scrObjects.Value).udtPhysics.dblY = txtObjectYCoord.Text

End Sub
